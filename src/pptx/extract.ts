import type { Rect } from "../render/geometry.js";
import {
  emuToInch,
  emuToPoint,
  type ExtractedDeck,
  type ExtractedElement,
  type ExtractedFill,
  type ExtractedRun,
  type ExtractedSlide,
  type ExtractedText,
  type ExtractedTheme,
} from "./model.js";
import { PptxPackage, RELATIONSHIP_TYPES, slideNumberOf, type PptxRelationship } from "./package.js";
import { attr, findChild, findChildren, findDescendants, findPath, numAttr, textOf, type XmlElement } from "./xml.js";

/**
 * PPTX → 中立なスライド表現。
 *
 * **対応する範囲を意図的に絞っている**（V2 スコープ）:
 *   text / 図形 (prstGeom) / 直線・コネクタ / 画像 / 単純なグループ
 *
 * SmartArt・グラフ・表・OLE・3D・アニメーション・特殊効果は解析せず、
 * `unsupported` に必ず記録する。黙って落とすことはしない。
 */

const GRAPHIC_DATA_URIS: Record<string, string> = {
  "http://schemas.openxmlformats.org/drawingml/2006/table": "表 (table)",
  "http://schemas.openxmlformats.org/drawingml/2006/chart": "グラフ (chart)",
  "http://schemas.openxmlformats.org/drawingml/2006/diagram": "SmartArt (diagram)",
  "http://schemas.openxmlformats.org/presentationml/2006/ole": "OLE 埋め込みオブジェクト",
};

/** slideMaster の clrMap を通した「参照名 → clrScheme のキー」 */
const DEFAULT_COLOR_MAP: Record<string, string> = {
  bg1: "lt1",
  tx1: "dk1",
  bg2: "lt2",
  tx2: "dk2",
  accent1: "accent1",
  accent2: "accent2",
  accent3: "accent3",
  accent4: "accent4",
  accent5: "accent5",
  accent6: "accent6",
  hlink: "hlink",
  folHlink: "folHlink",
};

export interface ExtractContext {
  theme?: ExtractedTheme;
  colorMap?: Record<string, string>;
  relationships?: Map<string, PptxRelationship>;
}

function transformOf(spPr: XmlElement | undefined): { box: Rect; rotation?: number } | undefined {
  const xfrm = spPr ? findChild(spPr, "xfrm") : undefined;
  if (!xfrm) return undefined;
  const off = findChild(xfrm, "off");
  const ext = findChild(xfrm, "ext");
  const x = numAttr(off, "x");
  const y = numAttr(off, "y");
  const cx = numAttr(ext, "cx");
  const cy = numAttr(ext, "cy");
  if (x === undefined || y === undefined || cx === undefined || cy === undefined) return undefined;
  const rot = numAttr(xfrm, "rot");
  return {
    box: { x: emuToInch(x), y: emuToInch(y), w: emuToInch(cx), h: emuToInch(cy) },
    ...(rot !== undefined ? { rotation: rot / 60000 } : {}),
  };
}

function resolveColorElement(
  parent: XmlElement | undefined,
  ctx: ExtractContext,
): string | undefined {
  if (!parent) return undefined;
  const srgb = findChild(parent, "srgbClr");
  if (srgb) return attr(srgb, "val")?.toUpperCase();
  const scheme = findChild(parent, "schemeClr");
  if (scheme) {
    const ref = attr(scheme, "val");
    if (!ref) return undefined;
    const map = ctx.colorMap ?? DEFAULT_COLOR_MAP;
    const key = map[ref] ?? ref;
    return ctx.theme?.colors[key];
  }
  const sys = findChild(parent, "sysClr");
  if (sys) return attr(sys, "lastClr")?.toUpperCase();
  return undefined;
}

function fillOf(spPr: XmlElement | undefined, ctx: ExtractContext): ExtractedFill | undefined {
  if (!spPr) return undefined;
  if (findChild(spPr, "noFill")) return { type: "none" };
  const solid = findChild(spPr, "solidFill");
  if (solid) {
    const color = resolveColorElement(solid, ctx);
    return color ? { type: "solid", color } : { type: "unsupported", detail: "解決できない solidFill" };
  }
  if (findChild(spPr, "gradFill")) return { type: "unsupported", detail: "グラデーション塗り" };
  if (findChild(spPr, "blipFill")) return { type: "unsupported", detail: "画像塗り" };
  if (findChild(spPr, "pattFill")) return { type: "unsupported", detail: "パターン塗り" };
  return undefined;
}

function lineOf(
  spPr: XmlElement | undefined,
  ctx: ExtractContext,
): { color?: string; widthPt?: number } | undefined {
  const ln = spPr ? findChild(spPr, "ln") : undefined;
  if (!ln) return undefined;
  if (findChild(ln, "noFill")) return undefined;
  const color = resolveColorElement(findChild(ln, "solidFill"), ctx);
  const width = numAttr(ln, "w");
  if (!color && width === undefined) return undefined;
  return {
    ...(color ? { color } : {}),
    ...(width !== undefined ? { widthPt: emuToPoint(width) } : {}),
  };
}

function textOfBody(txBody: XmlElement | undefined, ctx: ExtractContext): ExtractedText | undefined {
  if (!txBody) return undefined;
  const paragraphs: string[] = [];
  const runs: ExtractedRun[] = [];
  let align: ExtractedText["align"];

  for (const p of findChildren(txBody, "p")) {
    const pPr = findChild(p, "pPr");
    const algn = attr(pPr, "algn");
    if (!align && algn) {
      align = algn === "ctr" ? "center" : algn === "r" ? "right" : "left";
    }

    let paragraph = "";
    for (const child of p.children) {
      const local = child.name.includes(":") ? child.name.slice(child.name.indexOf(":") + 1) : child.name;
      if (local !== "r" && local !== "fld" && local !== "br") continue;
      if (local === "br") {
        paragraph += "\n";
        continue;
      }
      const t = findChild(child, "t");
      const value = t ? textOf(t) : "";
      if (value.length === 0) continue;
      const rPr = findChild(child, "rPr");
      const size = numAttr(rPr, "sz");
      const latin = findChild(rPr ?? child, "latin");
      runs.push({
        text: value,
        ...(size !== undefined ? { fontSize: size / 100 } : {}),
        ...(attr(rPr, "b") === "1" ? { bold: true } : {}),
        ...(attr(rPr, "i") === "1" ? { italic: true } : {}),
        ...(resolveColorElement(findChild(rPr ?? child, "solidFill"), ctx)
          ? { color: resolveColorElement(findChild(rPr ?? child, "solidFill"), ctx)! }
          : {}),
        ...(attr(latin, "typeface") ? { fontFace: attr(latin, "typeface")! } : {}),
      });
      paragraph += value;
    }
    if (paragraph.length > 0) paragraphs.push(paragraph);
  }

  if (paragraphs.length === 0) return undefined;
  const bodyPr = findChild(txBody, "bodyPr");
  const anchorRaw = attr(bodyPr, "anchor");
  const sizes = runs.map((r) => r.fontSize).filter((s): s is number => s !== undefined);

  return {
    paragraphs,
    runs,
    plain: paragraphs.join("\n"),
    ...(align ? { align } : {}),
    ...(anchorRaw
      ? { anchor: anchorRaw === "ctr" ? "middle" : anchorRaw === "b" ? "bottom" : "top" }
      : {}),
    ...(sizes.length > 0 ? { maxFontSize: Math.max(...sizes) } : {}),
  };
}

interface GroupTransform {
  scaleX: number;
  scaleY: number;
  offsetX: number;
  offsetY: number;
}

const IDENTITY: GroupTransform = { scaleX: 1, scaleY: 1, offsetX: 0, offsetY: 0 };

function applyTransform(box: Rect, t: GroupTransform): Rect {
  return {
    x: t.offsetX + box.x * t.scaleX,
    y: t.offsetY + box.y * t.scaleY,
    w: box.w * t.scaleX,
    h: box.h * t.scaleY,
  };
}

/** グループの子座標系 → 親座標系の変換を求める */
function groupTransform(grpSpPr: XmlElement | undefined, parent: GroupTransform): GroupTransform {
  const xfrm = grpSpPr ? findChild(grpSpPr, "xfrm") : undefined;
  if (!xfrm) return parent;
  const off = findChild(xfrm, "off");
  const ext = findChild(xfrm, "ext");
  const chOff = findChild(xfrm, "chOff");
  const chExt = findChild(xfrm, "chExt");
  const x = numAttr(off, "x") ?? 0;
  const y = numAttr(off, "y") ?? 0;
  const cx = numAttr(ext, "cx") ?? 0;
  const cy = numAttr(ext, "cy") ?? 0;
  const chX = numAttr(chOff, "x") ?? 0;
  const chY = numAttr(chOff, "y") ?? 0;
  const chCx = numAttr(chExt, "cx") ?? cx;
  const chCy = numAttr(chExt, "cy") ?? cy;

  const scaleX = chCx > 0 ? cx / chCx : 1;
  const scaleY = chCy > 0 ? cy / chCy : 1;
  return {
    scaleX: parent.scaleX * scaleX,
    scaleY: parent.scaleY * scaleY,
    offsetX: parent.offsetX + (emuToInch(x) - emuToInch(chX) * scaleX) * parent.scaleX,
    offsetY: parent.offsetY + (emuToInch(y) - emuToInch(chY) * scaleY) * parent.scaleY,
  };
}

interface ShapeTreeResult {
  elements: ExtractedElement[];
  unsupported: string[];
  warnings: string[];
}

function localName(name: string): string {
  const colon = name.indexOf(":");
  return colon >= 0 ? name.slice(colon + 1) : name;
}

function walkShapeTree(
  tree: XmlElement,
  ctx: ExtractContext,
  transform: GroupTransform,
  result: ShapeTreeResult,
  depth = 0,
): void {
  if (depth > 8) {
    result.warnings.push("グループのネストが深すぎるため、一部を展開していません");
    return;
  }

  for (const node of tree.children) {
    const kind = localName(node.name);

    if (kind === "sp") {
      const spPr = findChild(node, "spPr");
      const nv = findChild(node, "nvSpPr");
      const cNvPr = nv ? findChild(nv, "cNvPr") : undefined;
      const ph = nv ? findPath(nv, "nvPr", "ph") : undefined;
      const t = transformOf(spPr);
      const txBody = findChild(node, "txBody");
      const text = textOfBody(txBody, ctx);

      if (!t) {
        // 継承（レイアウト側の位置）に依存するプレースホルダは座標を持たない
        if (ph) {
          result.warnings.push(
            `プレースホルダ '${attr(cNvPr, "name") ?? attr(ph, "type") ?? "?"}' はレイアウト継承の位置のため取り込めません`,
          );
        }
        continue;
      }

      if (findChild(spPr!, "custGeom")) {
        result.unsupported.push("カスタム図形 (custGeom) — 矩形として近似します");
      }
      if (findChild(spPr!, "effectLst")?.children.length) {
        result.warnings.push("図形の効果 (影・光彩など) は取り込まれません");
      }
      if (findChild(spPr!, "scene3d") || findChild(spPr!, "sp3d")) {
        result.unsupported.push("3D 効果 — 非対応");
      }

      const prst = attr(findPath(spPr!, "prstGeom"), "prst");
      const fill = fillOf(spPr, ctx);
      const line = lineOf(spPr, ctx);
      if (fill?.type === "unsupported") {
        result.unsupported.push(`${fill.detail ?? "非対応の塗り"} — 単色として扱えません`);
      }

      result.elements.push({
        id: attr(cNvPr, "id") ?? `sp-${result.elements.length + 1}`,
        ...(attr(cNvPr, "name") ? { name: attr(cNvPr, "name")! } : {}),
        kind: ph ? "placeholder" : text ? "text" : "shape",
        box: applyTransform(t.box, transform),
        ...(t.rotation !== undefined ? { rotation: t.rotation } : {}),
        ...(prst ? { geometry: prst } : {}),
        ...(text ? { text } : {}),
        ...(fill ? { fill } : {}),
        ...(line ? { line } : {}),
        ...(ph
          ? {
              placeholder: {
                ...(attr(ph, "type") ? { type: attr(ph, "type")! } : {}),
                ...(attr(ph, "idx") ? { index: attr(ph, "idx")! } : {}),
              },
            }
          : {}),
        ...(transform !== IDENTITY ? { fromGroup: true } : {}),
      });
      continue;
    }

    if (kind === "pic") {
      const spPr = findChild(node, "spPr");
      const t = transformOf(spPr);
      if (!t) continue;
      const blip = findPath(node, "blipFill", "blip");
      const embed = attr(blip, "embed");
      const link = attr(blip, "link");
      const rel = embed ? ctx.relationships?.get(embed) : undefined;
      if (link) {
        result.unsupported.push("外部リンク画像 — 取得しません（セキュリティ上 fetch しない）");
      }
      const nvPr = findPath(node, "nvPicPr", "cNvPr");
      result.elements.push({
        id: attr(nvPr, "id") ?? `pic-${result.elements.length + 1}`,
        ...(attr(nvPr, "name") ? { name: attr(nvPr, "name")! } : {}),
        kind: "image",
        box: applyTransform(t.box, transform),
        image: {
          relationshipId: embed ?? link ?? "",
          ...(rel?.target ? { part: rel.target } : {}),
          ...(rel?.target ? { fileName: rel.target.slice(rel.target.lastIndexOf("/") + 1) } : {}),
          ...(rel?.externalTarget ? { externalTarget: rel.externalTarget } : {}),
        },
        ...(transform !== IDENTITY ? { fromGroup: true } : {}),
      });
      continue;
    }

    if (kind === "cxnSp") {
      const spPr = findChild(node, "spPr");
      const t = transformOf(spPr);
      if (!t) continue;
      const line = lineOf(spPr, ctx);
      result.elements.push({
        id: attr(findPath(node, "nvCxnSpPr", "cNvPr"), "id") ?? `cxn-${result.elements.length + 1}`,
        kind: "line",
        box: applyTransform(t.box, transform),
        ...(line ? { line } : {}),
        ...(transform !== IDENTITY ? { fromGroup: true } : {}),
      });
      continue;
    }

    if (kind === "grpSp") {
      const childTransform = groupTransform(findChild(node, "grpSpPr"), transform);
      walkShapeTree(node, ctx, childTransform, result, depth + 1);
      continue;
    }

    if (kind === "graphicFrame") {
      const uri = attr(findPath(node, "graphic", "graphicData"), "uri");
      const label = (uri && GRAPHIC_DATA_URIS[uri]) ?? `graphicFrame (${uri ?? "不明"})`;
      result.unsupported.push(`${label} — V2 では取り込みません`);
      continue;
    }

    if (kind === "AlternateContent") {
      result.unsupported.push("AlternateContent（新旧混在コンテンツ） — 非対応");
      continue;
    }
  }
}

/** スライド XML から要素を取り出す */
export function extractShapeTree(slideXml: XmlElement, ctx: ExtractContext): ShapeTreeResult {
  const tree = findPath(slideXml, "cSld", "spTree");
  const result: ShapeTreeResult = { elements: [], unsupported: [], warnings: [] };
  if (!tree) {
    result.warnings.push("spTree が見つかりません");
    return result;
  }
  walkShapeTree(tree, ctx, IDENTITY, result);

  for (const animTag of ["anim", "animEffect", "animMotion", "animRot", "animScale"]) {
    if (findDescendants(slideXml, animTag).length > 0) {
      result.unsupported.push("アニメーション — 取り込みません");
      break;
    }
  }
  if (findDescendants(slideXml, "transition").length > 0) {
    result.unsupported.push("画面切り替え効果 — 取り込みません");
  }

  return result;
}

/** theme1.xml からカラースキームとフォントを取り出す */
export function extractTheme(themeXml: XmlElement): ExtractedTheme {
  const colors: Record<string, string> = {};
  const scheme = findPath(themeXml, "themeElements", "clrScheme");
  for (const child of scheme?.children ?? []) {
    const key = localName(child.name);
    const value =
      attr(findChild(child, "srgbClr"), "val")?.toUpperCase() ??
      attr(findChild(child, "sysClr"), "lastClr")?.toUpperCase();
    if (value) colors[key] = value;
  }

  const fontScheme = findPath(themeXml, "themeElements", "fontScheme");
  const major = findChild(fontScheme ?? themeXml, "majorFont");
  const minor = findChild(fontScheme ?? themeXml, "minorFont");

  return {
    colors,
    fonts: {
      ...(attr(findChild(major ?? themeXml, "latin"), "typeface")
        ? { major: attr(findChild(major!, "latin"), "typeface")! }
        : {}),
      ...(attr(findChild(minor ?? themeXml, "latin"), "typeface")
        ? { minor: attr(findChild(minor!, "latin"), "typeface")! }
        : {}),
      ...(attr(findChild(major ?? themeXml, "ea"), "typeface")
        ? { majorEastAsian: attr(findChild(major!, "ea"), "typeface")! }
        : {}),
      ...(attr(findChild(minor ?? themeXml, "ea"), "typeface")
        ? { minorEastAsian: attr(findChild(minor!, "ea"), "typeface")! }
        : {}),
    },
  };
}

/** slideMaster の clrMap（bg1 → lt1 など） */
export function extractColorMap(masterXml: XmlElement): Record<string, string> {
  const clrMap = findChild(masterXml, "clrMap");
  if (!clrMap) return { ...DEFAULT_COLOR_MAP };
  const map: Record<string, string> = { ...DEFAULT_COLOR_MAP };
  for (const [key, value] of Object.entries(clrMap.attrs)) map[key] = value;
  return map;
}

export interface ReadDeckOptions {
  /** 読み込むスライド番号（1 始まり）。省略時は全スライド。 */
  slideNumbers?: number[];
  /** ノートも読むか（既定 true） */
  includeNotes?: boolean;
}

/** PPTX 全体を中立表現へ読み出す */
export async function readDeck(
  pkg: PptxPackage,
  options: ReadDeckOptions = {},
): Promise<ExtractedDeck> {
  const unsupported = pkg.detectPackageLevelUnsupported();
  const warnings: string[] = [];

  const presentation = await pkg.xml("ppt/presentation.xml");
  const sldSz = findChild(presentation, "sldSz");
  const slideSize = {
    widthIn: emuToInch(numAttr(sldSz, "cx") ?? 9144000),
    heightIn: emuToInch(numAttr(sldSz, "cy") ?? 5143500),
  };

  // マスター → テーマ / カラーマップ
  let theme: ExtractedTheme | undefined;
  let colorMap: Record<string, string> | undefined;
  const masterPart = pkg.list("ppt/slideMasters/").find((p) => /slideMaster\d+\.xml$/.test(p));
  if (masterPart) {
    const masterXml = await pkg.xml(masterPart);
    colorMap = extractColorMap(masterXml);
    const masterRels = await pkg.relationships(masterPart);
    const themeRel = [...masterRels.values()].find((r) => r.type === RELATIONSHIP_TYPES.theme);
    if (themeRel?.target && pkg.has(themeRel.target)) {
      theme = extractTheme(await pkg.xml(themeRel.target));
    }
  }
  if (!theme) {
    const themePart = pkg.list("ppt/theme/").find((p) => /theme\d+\.xml$/.test(p));
    if (themePart) theme = extractTheme(await pkg.xml(themePart));
  }
  if (!theme) warnings.push("テーマ (theme1.xml) を読み取れませんでした");

  const metadata: ExtractedDeck["metadata"] = {};
  if (pkg.has("docProps/core.xml")) {
    const core = await pkg.xml("docProps/core.xml");
    const pick = (name: string): string | undefined => {
      const el = findChild(core, name);
      const value = el ? textOf(el).trim() : "";
      return value.length > 0 ? value : undefined;
    };
    Object.assign(metadata, {
      ...(pick("title") ? { title: pick("title") } : {}),
      ...(pick("creator") ? { creator: pick("creator") } : {}),
      ...(pick("created") ? { created: pick("created") } : {}),
      ...(pick("modified") ? { modified: pick("modified") } : {}),
    });
  }
  if (pkg.has("docProps/app.xml")) {
    const app = await pkg.xml("docProps/app.xml");
    const company = findChild(app, "Company");
    const value = company ? textOf(company).trim() : "";
    if (value) metadata.company = value;
  }

  const wanted = options.slideNumbers ? new Set(options.slideNumbers) : undefined;
  const slides: ExtractedSlide[] = [];

  for (const part of pkg.slideParts()) {
    const slideNumber = slideNumberOf(part);
    if (wanted && !wanted.has(slideNumber)) continue;

    const relationships = await pkg.relationships(part);
    const slideXml = await pkg.xml(part);
    const ctx: ExtractContext = {
      ...(theme ? { theme } : {}),
      ...(colorMap ? { colorMap } : {}),
      relationships,
    };
    const extracted = extractShapeTree(slideXml, ctx);

    let notes: string | undefined;
    if (options.includeNotes !== false) {
      const notesRel = [...relationships.values()].find(
        (r) => r.type === RELATIONSHIP_TYPES.notesSlide,
      );
      if (notesRel?.target && pkg.has(notesRel.target)) {
        const notesXml = await pkg.xml(notesRel.target);
        const text = findDescendants(notesXml, "t")
          .map((t) => textOf(t))
          .join("\n")
          .trim();
        if (text.length > 0) notes = text;
      }
    }

    const titleElement = extracted.elements.find(
      (e) => e.placeholder?.type === "title" || e.placeholder?.type === "ctrTitle",
    );

    slides.push({
      slideNumber,
      part,
      elements: extracted.elements,
      unsupported: [...new Set(extracted.unsupported)],
      warnings: [...new Set(extracted.warnings)],
      ...(notes ? { notes } : {}),
      ...(titleElement?.text?.plain ? { title: titleElement.text.plain } : {}),
    });
  }

  return {
    slideSize,
    ...(theme ? { theme } : {}),
    slides,
    unsupported: [...new Set(unsupported)],
    warnings: [...new Set(warnings)],
    metadata,
  };
}
