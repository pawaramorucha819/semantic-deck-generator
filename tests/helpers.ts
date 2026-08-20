import { readFile } from "node:fs/promises";
import { fileURLToPath } from "node:url";
import { dirname, join } from "node:path";
import JSZip from "jszip";
import { parsePresentation } from "../src/domain/schema.js";
import type {
  ContentItem,
  PresentationSpec,
  SemanticContent,
  SlideSpec,
} from "../src/domain/types.js";

const here = dirname(fileURLToPath(import.meta.url));
export const projectRoot = join(here, "..");
export const examplePath = join(projectRoot, "examples", "inventory-system.json");

export async function loadExample(): Promise<PresentationSpec> {
  return parsePresentation(JSON.parse(await readFile(examplePath, "utf8")));
}

export async function loadExampleJson(): Promise<Record<string, unknown>> {
  return JSON.parse(await readFile(examplePath, "utf8")) as Record<string, unknown>;
}

/** テスト用の最小 SlideSpec */
export function slide(partial: Partial<SlideSpec> = {}): SlideSpec {
  return {
    id: "slide-test",
    keyMessage: "テスト用のキーメッセージ",
    slideRole: "analysis",
    logicalRelation: "claim_evidence",
    evidence: [],
    ...partial,
  };
}

export function textEvidence(id: string, label: string, statement = `${label}に関する事実`) {
  return {
    id,
    type: "text" as const,
    statement,
    data: { label },
  };
}

export function metricEvidence(id: string, label: string, value: number, unit = "%") {
  return {
    id,
    type: "metric" as const,
    statement: `${label}は${value}${unit}である`,
    data: { label, value, unit },
  };
}

/** content (kind=sequence) を組み立てる */
export function sequenceContent(
  ...steps: Array<string | ContentItem>
): SemanticContent {
  return {
    kind: "sequence",
    steps: steps.map((s) => (typeof s === "string" ? { label: s } : s)),
  };
}

/** content (kind=points) を組み立てる */
export function pointsContent(...items: Array<string | ContentItem>): SemanticContent {
  return {
    kind: "points",
    items: items.map((s) => (typeof s === "string" ? { label: s } : s)),
  };
}

export interface PptxContents {
  zip: JSZip;
  slideXmls: string[];
  presentationXml: string;
  /** ディレクトリエントリを除いた実ファイルのみ */
  files: string[];
  /** Slide Master（Template 検査用） */
  slideMasterXmls: string[];
  /** Slide Layout（Template の master 定義はここへ出る） */
  slideLayoutXmls: string[];
  /** slideN.xml → 参照している slideLayout の番号 */
  slideLayoutRefs: number[];
}

export async function readPptx(buffer: Buffer): Promise<PptxContents> {
  const zip = await JSZip.loadAsync(buffer);
  const files = Object.entries(zip.files)
    .filter(([, entry]) => !entry.dir)
    .map(([name]) => name)
    .sort();
  const slideNames = files
    .filter((f) => /^ppt\/slides\/slide\d+\.xml$/.test(f))
    .sort((a, b) => slideIndex(a) - slideIndex(b));
  const slideXmls = await Promise.all(
    slideNames.map((name) => zip.file(name)!.async("string")),
  );
  const presentationXml = await zip.file("ppt/presentation.xml")!.async("string");

  const masterNames = files.filter((f) => /^ppt\/slideMasters\/slideMaster\d+\.xml$/.test(f));
  const slideMasterXmls = await Promise.all(
    masterNames.map((name) => zip.file(name)!.async("string")),
  );
  const layoutNames = files
    .filter((f) => /^ppt\/slideLayouts\/slideLayout\d+\.xml$/.test(f))
    .sort((a, b) => layoutIndex(a) - layoutIndex(b));
  const slideLayoutXmls = await Promise.all(
    layoutNames.map((name) => zip.file(name)!.async("string")),
  );

  const slideLayoutRefs: number[] = [];
  for (const name of slideNames) {
    const relName = name.replace(/slides\/(slide\d+)\.xml$/, "slides/_rels/$1.xml.rels");
    const rel = await zip.file(relName)?.async("string");
    slideLayoutRefs.push(Number(/slideLayout(\d+)\.xml/.exec(rel ?? "")?.[1] ?? 0));
  }

  return {
    zip,
    slideXmls,
    presentationXml,
    files,
    slideMasterXmls,
    slideLayoutXmls,
    slideLayoutRefs,
  };
}

function layoutIndex(name: string): number {
  return Number(/slideLayout(\d+)\.xml$/.exec(name)?.[1] ?? 0);
}

/** 1x1 の有効な PNG（テンプレートアセットのテスト用） */
export const TINY_PNG_BASE64 =
  "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==";

export function tinyPngBytes(): Buffer {
  return Buffer.from(TINY_PNG_BASE64, "base64");
}

function slideIndex(name: string): number {
  return Number(/slide(\d+)\.xml$/.exec(name)?.[1] ?? 0);
}

/** スライド XML から <a:t> のテキストをすべて抜き出す */
export function extractTexts(slideXml: string): string[] {
  return [...slideXml.matchAll(/<a:t>([\s\S]*?)<\/a:t>/g)].map((m) =>
    (m[1] ?? "")
      .replace(/&amp;/g, "&")
      .replace(/&lt;/g, "<")
      .replace(/&gt;/g, ">")
      .replace(/&quot;/g, '"')
      .replace(/&apos;/g, "'"),
  );
}

export function countMatches(xml: string, pattern: RegExp): number {
  return [...xml.matchAll(pattern)].length;
}
