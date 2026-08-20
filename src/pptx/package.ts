import { readFile } from "node:fs/promises";
import JSZip from "jszip";
import { attr, parseXml, type XmlElement } from "./xml.js";

/**
 * `.pptx`（OOXML パッケージ）の読み取り。
 *
 * PPTX は **信頼できない入力** として扱う。
 *  - ZIP のパストラバーサル（'..' / 絶対パス）を拒否する
 *  - エントリ数・展開後サイズに上限を設ける（ZIP 爆弾対策）
 *  - マクロ (vbaProject.bin) は実行しないどころか読み込まず、非対応として報告する
 *  - 外部リレーション (TargetMode="External") は fetch せず、記録だけする
 */

export const MAX_ZIP_ENTRIES = 5000;
export const MAX_PART_BYTES = 24 * 1024 * 1024;
export const MAX_TOTAL_BYTES = 200 * 1024 * 1024;

export class PptxReadError extends Error {
  constructor(message: string) {
    super(message);
    this.name = "PptxReadError";
  }
}

export interface PptxRelationship {
  id: string;
  type: string;
  /** パッケージ内の正規化されたパス（外部リンクの場合は undefined） */
  target?: string;
  /** TargetMode="External" のときの生の値。fetch はしない。 */
  externalTarget?: string;
}

/** 'ppt/slides/slide1.xml' + '../media/img.png' → 'ppt/media/img.png' */
export function resolvePartPath(basePart: string, target: string): string {
  if (target.startsWith("/")) return normalizeZipPath(target.slice(1));
  const baseDir = basePart.includes("/") ? basePart.slice(0, basePart.lastIndexOf("/")) : "";
  const segments = `${baseDir}/${target}`.split("/");
  const out: string[] = [];
  for (const segment of segments) {
    if (segment === "" || segment === ".") continue;
    if (segment === "..") {
      if (out.length === 0) {
        throw new PptxReadError(`パッケージ外を指すパスです: ${target}`);
      }
      out.pop();
      continue;
    }
    out.push(segment);
  }
  return out.join("/");
}

export function normalizeZipPath(name: string): string {
  const normalized = name.replace(/\\/g, "/");
  if (normalized.startsWith("/") || /^[a-zA-Z]:/.test(normalized)) {
    throw new PptxReadError(`絶対パスのエントリは受け付けません: ${name}`);
  }
  if (normalized.split("/").includes("..")) {
    throw new PptxReadError(`'..' を含むエントリは受け付けません: ${name}`);
  }
  if (normalized.includes("\0")) {
    throw new PptxReadError(`不正なエントリ名です: ${name}`);
  }
  return normalized;
}

export class PptxPackage {
  private totalRead = 0;
  private readonly xmlCache = new Map<string, XmlElement>();

  private constructor(
    private readonly zip: JSZip,
    readonly entries: string[],
  ) {}

  static async open(input: Buffer | string): Promise<PptxPackage> {
    const buffer = typeof input === "string" ? await readFile(input) : input;
    if (buffer.length < 4 || buffer.subarray(0, 2).toString("latin1") !== "PK") {
      throw new PptxReadError("PPTX (ZIP) として読み込めません: シグネチャが不正です");
    }

    let zip: JSZip;
    try {
      zip = await JSZip.loadAsync(buffer);
    } catch (error) {
      throw new PptxReadError(
        `PPTX を展開できませんでした: ${error instanceof Error ? error.message : String(error)}`,
      );
    }

    const entries: string[] = [];
    for (const [name, entry] of Object.entries(zip.files)) {
      if (entry.dir) continue;
      entries.push(normalizeZipPath(name));
    }
    if (entries.length === 0) {
      throw new PptxReadError("PPTX の中身が空です");
    }
    if (entries.length > MAX_ZIP_ENTRIES) {
      throw new PptxReadError(`エントリ数が多すぎます (${entries.length} > ${MAX_ZIP_ENTRIES})`);
    }
    if (!entries.includes("[Content_Types].xml")) {
      throw new PptxReadError("[Content_Types].xml がありません（PPTX ではない可能性があります）");
    }
    if (!entries.some((e) => e.startsWith("ppt/slides/slide"))) {
      throw new PptxReadError("スライドが 1 枚も含まれていません");
    }

    return new PptxPackage(zip, entries.sort());
  }

  has(part: string): boolean {
    return this.entries.includes(part);
  }

  list(prefix: string): string[] {
    return this.entries.filter((e) => e.startsWith(prefix));
  }

  /** スライドを番号順に返す */
  slideParts(): string[] {
    return this.list("ppt/slides/slide")
      .filter((f) => /^ppt\/slides\/slide\d+\.xml$/.test(f))
      .sort((a, b) => slideNumberOf(a) - slideNumberOf(b));
  }

  private track(bytes: number, part: string): void {
    if (bytes > MAX_PART_BYTES) {
      throw new PptxReadError(`パートが大きすぎます: ${part} (${bytes} bytes)`);
    }
    this.totalRead += bytes;
    if (this.totalRead > MAX_TOTAL_BYTES) {
      throw new PptxReadError("展開後の合計サイズが上限を超えました（ZIP 爆弾の可能性）");
    }
  }

  async text(part: string): Promise<string> {
    const file = this.zip.file(part);
    if (!file) throw new PptxReadError(`パートが見つかりません: ${part}`);
    const content = await file.async("string");
    this.track(Buffer.byteLength(content), part);
    return content;
  }

  async binary(part: string): Promise<Buffer> {
    const file = this.zip.file(part);
    if (!file) throw new PptxReadError(`パートが見つかりません: ${part}`);
    const content = await file.async("nodebuffer");
    this.track(content.length, part);
    return content;
  }

  async xml(part: string): Promise<XmlElement> {
    const cached = this.xmlCache.get(part);
    if (cached) return cached;
    const parsed = parseXml(await this.text(part));
    this.xmlCache.set(part, parsed);
    return parsed;
  }

  /** パートの .rels を読む（存在しなければ空） */
  async relationships(part: string): Promise<Map<string, PptxRelationship>> {
    const dir = part.includes("/") ? part.slice(0, part.lastIndexOf("/")) : "";
    const file = part.slice(part.lastIndexOf("/") + 1);
    const relPart = `${dir}/_rels/${file}.rels`;
    const out = new Map<string, PptxRelationship>();
    if (!this.has(relPart)) return out;

    const root = await this.xml(relPart);
    for (const rel of root.children) {
      const id = attr(rel, "Id");
      const type = attr(rel, "Type");
      const target = attr(rel, "Target");
      if (!id || !type || !target) continue;
      if (attr(rel, "TargetMode") === "External") {
        out.set(id, { id, type, externalTarget: target });
        continue;
      }
      out.set(id, { id, type, target: resolvePartPath(part, target) });
    }
    return out;
  }

  /** マクロなど「読み込まないが存在は報告する」要素 */
  detectPackageLevelUnsupported(): string[] {
    const found: string[] = [];
    if (this.entries.some((e) => e.endsWith("vbaProject.bin"))) {
      found.push("VBA マクロ (vbaProject.bin) — 取り込まず、実行もしません");
    }
    if (this.entries.some((e) => e.startsWith("ppt/embeddings/"))) {
      found.push("埋め込みオブジェクト (ppt/embeddings) — 非対応");
    }
    if (this.entries.some((e) => e.startsWith("ppt/media/") && /\.(mp4|wmv|avi|mp3|wav)$/i.test(e))) {
      found.push("動画・音声メディア — 非対応");
    }
    return found;
  }
}

export function slideNumberOf(part: string): number {
  return Number(/slide(\d+)\.xml$/.exec(part)?.[1] ?? 0);
}

export const RELATIONSHIP_TYPES = {
  slide: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide",
  slideLayout: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout",
  slideMaster: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster",
  theme: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme",
  image: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
  notesSlide: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide",
} as const;
