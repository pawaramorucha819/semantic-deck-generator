import { mkdir, writeFile } from "node:fs/promises";
import { join } from "node:path";
import { PptxPackage, RELATIONSHIP_TYPES } from "../../pptx/package.js";
import {
  extractColorMap,
  extractShapeTree,
  extractTheme,
  type ExtractContext,
} from "../../pptx/extract.js";
import {
  emuToInch,
  hasRenderableExtent,
  isLineElement,
  type ExtractedElement,
  type ExtractedTheme,
} from "../../pptx/model.js";
import { attr, findChild, findPath, numAttr } from "../../pptx/xml.js";
import { nearestColorToken } from "../../render/tokens.js";
import { defaultTheme, mix, resolveTheme, tint, type PresentationTheme } from "../../theme/theme.js";
import { parseTemplate } from "../schema.js";
import {
  DEFAULT_SLIDE_HEIGHT_IN,
  DEFAULT_SLIDE_WIDTH_IN,
  TEMPLATE_SCHEMA_VERSION,
  type PresentationTemplate,
  type SlideMasterDefinition,
  type TemplateElement,
} from "../template.js";

/**
 * 既存 `.pptx` → Semantic Deck Generator の Template。
 *
 * **完全互換の Import は目指さない**（要件どおり）。
 * 取り込むのは「企業の紙面を再現するのに実際に効く部分」だけで、
 * 取り込めなかったものは必ず `unsupportedFeatures` / `warnings` に残す。
 *
 * 取り込む: スライドサイズ / テーマ配色 / フォント / マスター背景 /
 *           マスターとレイアウトの固定図形・固定テキスト・画像 /
 *           ページ番号プレースホルダの位置 / 本文プレースホルダから求めた余白
 * 取り込まない: SmartArt / グラフ / 表 / OLE / アニメーション / 3D / 特殊効果 /
 *               グラデーション・画像塗り / プレースホルダの継承レイアウト詳細
 */

export const TEMPLATE_IMPORT_REPORT_VERSION = "1.0";

export interface TemplateImportReport {
  schemaVersion: string;
  source: string;
  templateId: string;
  slideSize: { widthIn: number; heightIn: number };
  extracted: {
    themeColors: number;
    masters: number;
    elements: number;
    assets: number;
  };
  /** 取り込めなかった PowerPoint 機能 */
  unsupportedFeatures: string[];
  warnings: string[];
}

export interface ImportedAsset {
  id: string;
  /** テンプレートディレクトリからの相対パス */
  path: string;
  bytes: Buffer;
}

export interface ImportedTemplate {
  template: PresentationTemplate;
  assets: ImportedAsset[];
  report: TemplateImportReport;
}

export interface ImportTemplateOptions {
  /** 生成する Template の id（英小文字・数字・- _） */
  id: string;
  name?: string;
  organization?: string;
  department?: string;
  confidentialityLabel?: string;
  /** 取り込み元の表示名（既定: 入力パス） */
  sourceLabel?: string;
}

function themeFromPptx(
  extracted: ExtractedTheme | undefined,
  id: string,
  warnings: string[],
): PresentationTheme {
  if (!extracted) {
    warnings.push("テーマを読み取れなかったため既定テーマを使用します");
    return { ...defaultTheme, id, name: `${id} (default fallback)` };
  }

  const colors = extracted.colors;
  const background = colors.lt1 ?? colors.bg1 ?? defaultTheme.colors.background;
  const text = colors.dk1 ?? colors.tx1 ?? defaultTheme.colors.text;
  const primary = colors.accent1 ?? defaultTheme.colors.primary;
  const secondary = colors.accent2 ?? mix(primary, background, 0.3);
  const accent = colors.accent3 ?? colors.accent6 ?? defaultTheme.colors.accent;
  const mutedText = mix(text, background, 0.45);

  const heading =
    extracted.fonts.majorEastAsian ?? extracted.fonts.major ?? defaultTheme.fonts.heading;
  const body = extracted.fonts.minorEastAsian ?? extracted.fonts.minor ?? defaultTheme.fonts.body;

  warnings.push(
    "フォントサイズ・余白の体系は PPTX から一意に決まらないため既定値を使用します（template.json で調整してください）",
  );

  return {
    id,
    name: `${id} (imported)`,
    fonts: {
      heading,
      body,
      numeric: extracted.fonts.major ?? defaultTheme.fonts.numeric!,
    },
    fontSizes: { ...defaultTheme.fontSizes },
    spacing: { ...defaultTheme.spacing },
    colors: {
      background,
      text,
      mutedText,
      primary,
      secondary,
      accent,
      surface: tint(primary, 0.92),
      border: tint(mutedText, 0.7),
      positive: colors.accent6 ?? defaultTheme.colors.positive!,
      negative: colors.accent5 ?? defaultTheme.colors.negative!,
      onPrimary: background,
    },
  };
}

interface ConversionResult {
  elements: TemplateElement[];
  slideNumber?: SlideMasterDefinition["slideNumber"];
  contentMargins?: SlideMasterDefinition["contentMargins"];
  assets: ImportedAsset[];
  warnings: string[];
  unsupported: string[];
}

/** 取り込んだ要素 → Template 要素（色は必ず Theme トークンへ寄せる） */
function convertElements(
  elements: ExtractedElement[],
  theme: PresentationTheme,
  slideSize: { widthIn: number; heightIn: number },
  assetPrefix: string,
  imageParts: Map<string, Buffer>,
): ConversionResult {
  const resolved = resolveTheme(theme);
  const out: ConversionResult = { elements: [], assets: [], warnings: [], unsupported: [] };
  let assetIndex = 0;

  for (const element of elements) {
    // 線は w か h が 0 になりうるので、通常図形と同じ条件で弾かない
    if (!hasRenderableExtent(element)) continue;
    if (
      element.box.x < -1 ||
      element.box.y < -1 ||
      element.box.x > slideSize.widthIn + 1 ||
      element.box.y > slideSize.heightIn + 1
    ) {
      out.warnings.push(`スライド外の要素を除外しました: ${element.name ?? element.id}`);
      continue;
    }

    if (element.kind === "placeholder") {
      const type = element.placeholder?.type;
      if (type === "sldNum") {
        out.slideNumber = {
          box: element.box,
          style: { role: "caption", font: "numeric", color: "mutedText", align: "right", valign: "middle" },
        };
        continue;
      }
      if (type === "dt") {
        out.warnings.push("日付プレースホルダは決定論性のため取り込みません");
        continue;
      }
      if (type === "ftr") {
        out.elements.push({
          kind: "text",
          box: element.box,
          text: "{{confidentiality}}",
          style: { role: "caption", color: "mutedText", valign: "middle" },
        });
        continue;
      }
      if (type === "body" || type === "obj" || type === undefined) {
        // 本文プレースホルダの矩形が「Pattern が描いてよい領域」になる
        out.contentMargins = {
          top: round(element.box.y),
          left: round(element.box.x),
          right: round(Math.max(0, slideSize.widthIn - element.box.x - element.box.w)),
          bottom: round(Math.max(0, slideSize.heightIn - element.box.y - element.box.h)),
        };
        continue;
      }
      if (type === "title" || type === "ctrTitle" || type === "subTitle") {
        out.warnings.push(
          `タイトル系プレースホルダ (${type}) は Pattern が描くため Master へは取り込みません`,
        );
        continue;
      }
      out.warnings.push(`未対応のプレースホルダ (${type}) を無視しました`);
      continue;
    }

    if (element.kind === "image") {
      const part = element.image?.part;
      if (element.image?.externalTarget) {
        out.unsupported.push("外部リンク画像 — 取得しません");
        continue;
      }
      const bytes = part ? imageParts.get(part) : undefined;
      if (!bytes || !part) {
        out.warnings.push(`画像を解決できませんでした: ${element.name ?? element.id}`);
        continue;
      }
      const extension = part.slice(part.lastIndexOf("."));
      if (![".png", ".jpg", ".jpeg", ".gif"].includes(extension.toLowerCase())) {
        out.unsupported.push(`対応していない画像形式です (${extension})`);
        continue;
      }
      assetIndex += 1;
      const id = `${assetPrefix}image-${assetIndex}`;
      const path = `assets/${id}${extension.toLowerCase()}`;
      out.assets.push({ id, path, bytes });
      out.elements.push({ kind: "image", assetId: id, box: element.box, sizing: "contain" });
      continue;
    }

    // コネクタ (p:cxnSp) と線プリセットの図形 (p:sp + prstGeom line) の両方を線として扱う。
    // 線は塗りを持たないので、塗りベースの分岐より前に処理する必要がある。
    if (isLineElement(element)) {
      const color = element.line?.color
        ? nearestColorToken(resolved, element.line.color)
        : undefined;
      out.elements.push({
        kind: "shape",
        shape: "line",
        box: element.box,
        style: {
          line: {
            color: color?.token ?? "border",
            ...(element.line?.widthPt !== undefined ? { width: round(element.line.widthPt) } : {}),
          },
        },
      });
      if (color && !color.exact) {
        out.warnings.push(
          `線の色 #${element.line!.color} を Theme トークン '${color.token}' に寄せました`,
        );
      }
      continue;
    }

    if (element.kind === "text" && element.text) {
      const color = element.text.runs.find((r) => r.color)?.color;
      const token = color ? nearestColorToken(resolved, color) : undefined;
      if (token && !token.exact) {
        out.warnings.push(`文字色 #${color} を Theme トークン '${token.token}' に寄せました`);
      }
      out.elements.push({
        kind: "text",
        box: element.box,
        text: element.text.plain.replace(/\n/g, " ").slice(0, 400),
        style: {
          role: "caption",
          ...(element.text.maxFontSize ? { fontSize: round(element.text.maxFontSize) } : {}),
          color: token?.token ?? "mutedText",
          ...(element.text.align ? { align: element.text.align } : {}),
          ...(element.text.anchor ? { valign: element.text.anchor } : {}),
          ...(element.text.runs.some((r) => r.bold) ? { bold: true } : {}),
        },
      });
      continue;
    }

    // 図形（背景の帯など）
    const fill = element.fill;
    if (!fill || fill.type === "none") continue;
    if (fill.type === "unsupported") {
      out.unsupported.push(`${fill.detail ?? "非対応の塗り"} — 図形を取り込めません`);
      continue;
    }
    const token = nearestColorToken(resolved, fill.color!);
    if (!token.exact) {
      out.warnings.push(`図形の色 #${fill.color} を Theme トークン '${token.token}' に寄せました`);
    }
    if (element.geometry && element.geometry !== "rect" && element.geometry !== "line") {
      out.unsupported.push(
        `図形 '${element.geometry}' は Slide Master では矩形として扱われます`,
      );
    }
    out.elements.push({
      kind: "shape",
      shape: element.geometry === "line" ? "line" : "rect",
      box: element.box,
      style: { fill: token.token },
    });
  }

  return out;
}

const round = (n: number): number => Number(n.toFixed(3));

/** 既存 PPTX から Template を取り込む */
export async function importTemplateFromPptx(
  input: Buffer | string,
  options: ImportTemplateOptions,
): Promise<ImportedTemplate> {
  const pkg = await PptxPackage.open(input);
  const sourceLabel =
    options.sourceLabel ?? (typeof input === "string" ? input : "(buffer)");

  const unsupported = [...pkg.detectPackageLevelUnsupported()];
  const warnings: string[] = [];

  const presentation = await pkg.xml("ppt/presentation.xml");
  const sldSz = findChild(presentation, "sldSz");
  const slideSize = {
    widthIn: round(emuToInch(numAttr(sldSz, "cx") ?? DEFAULT_SLIDE_WIDTH_IN * 914400)),
    heightIn: round(emuToInch(numAttr(sldSz, "cy") ?? DEFAULT_SLIDE_HEIGHT_IN * 914400)),
  };

  const masterPart = pkg.list("ppt/slideMasters/").find((p) => /slideMaster\d+\.xml$/.test(p));
  if (!masterPart) {
    throw new Error("slideMaster が見つかりません（Template として取り込めません）");
  }
  const masterXml = await pkg.xml(masterPart);
  const colorMap = extractColorMap(masterXml);
  const masterRels = await pkg.relationships(masterPart);

  let extractedTheme: ExtractedTheme | undefined;
  const themeRel = [...masterRels.values()].find((r) => r.type === RELATIONSHIP_TYPES.theme);
  if (themeRel?.target && pkg.has(themeRel.target)) {
    extractedTheme = extractTheme(await pkg.xml(themeRel.target));
  }
  const theme = themeFromPptx(extractedTheme, options.id, warnings);

  /** そのパートが参照する画像を実バイト列で集める */
  const imagesOf = async (part: string): Promise<Map<string, Buffer>> => {
    const rels = await pkg.relationships(part);
    const map = new Map<string, Buffer>();
    for (const rel of rels.values()) {
      if (rel.type !== RELATIONSHIP_TYPES.image || !rel.target) continue;
      if (!pkg.has(rel.target)) continue;
      map.set(rel.target, await pkg.binary(rel.target));
    }
    return map;
  };

  const convertPart = async (
    part: string,
    xml: Awaited<ReturnType<typeof pkg.xml>>,
    prefix: string,
  ): Promise<ConversionResult> => {
    const rels = await pkg.relationships(part);
    const ctx: ExtractContext = {
      ...(extractedTheme ? { theme: extractedTheme } : {}),
      colorMap,
      relationships: rels,
    };
    const extracted = extractShapeTree(xml, ctx);
    unsupported.push(...extracted.unsupported);
    warnings.push(...extracted.warnings);
    return convertElements(extracted.elements, theme, slideSize, prefix, await imagesOf(part));
  };

  const masterConversion = await convertPart(masterPart, masterXml, "master-");
  unsupported.push(...masterConversion.unsupported);
  warnings.push(...masterConversion.warnings);

  const masterBackground = resolveBackgroundToken(masterXml, theme, extractedTheme, colorMap);

  const assets: ImportedAsset[] = [...masterConversion.assets];
  const masters: SlideMasterDefinition[] = [];

  // どのレイアウトが「本文の紙面」かは、実際にスライドが最も多く参照している
  // レイアウトで決める（PowerPoint 製・ツール製のどちらでも同じ判断ができる）。
  const layoutParts = pkg
    .list("ppt/slideLayouts/")
    .filter((p) => /slideLayout\d+\.xml$/.test(p))
    .sort();

  const layoutUsage = new Map<string, number>();
  const layoutOfSlide = new Map<number, string>();
  for (const slidePart of pkg.slideParts()) {
    const rels = await pkg.relationships(slidePart);
    const layoutRel = [...rels.values()].find((r) => r.type === RELATIONSHIP_TYPES.slideLayout);
    if (!layoutRel?.target) continue;
    layoutUsage.set(layoutRel.target, (layoutUsage.get(layoutRel.target) ?? 0) + 1);
    layoutOfSlide.set(Number(/slide(\d+)\.xml$/.exec(slidePart)?.[1] ?? 0), layoutRel.target);
  }

  const layoutByRole = new Map<"title" | "section", string>();
  for (const part of layoutParts) {
    const type = attr(await pkg.xml(part), "type");
    if (type === "title" && !layoutByRole.has("title")) layoutByRole.set("title", part);
    if (type === "secHead" && !layoutByRole.has("section")) layoutByRole.set("section", part);
  }

  const mostUsedLayout = [...layoutUsage.entries()].sort(
    (a, b) => b[1] - a[1] || a[0].localeCompare(b[0]),
  )[0]?.[0];
  const contentLayoutPart =
    layoutParts.find((p) => p !== layoutByRole.get("title") && p === mostUsedLayout) ??
    mostUsedLayout ??
    layoutParts[0];

  // タイトル用レイアウトが type で分からない場合は「1 枚目が使っていて
  // 本文レイアウトとは違うもの」をタイトル扱いにする
  if (!layoutByRole.has("title")) {
    const firstSlideLayout = layoutOfSlide.get(1);
    if (firstSlideLayout && firstSlideLayout !== contentLayoutPart) {
      layoutByRole.set("title", firstSlideLayout);
    }
  }

  let contentConversion = masterConversion;
  if (contentLayoutPart) {
    const layoutConversion = await convertPart(
      contentLayoutPart,
      await pkg.xml(contentLayoutPart),
      "content-",
    );
    unsupported.push(...layoutConversion.unsupported);
    warnings.push(...layoutConversion.warnings);
    assets.push(...layoutConversion.assets);
    contentConversion = {
      // マスターの装飾 + 本文レイアウトの装飾（PowerPoint と同じ重なり順）
      elements: [...masterConversion.elements, ...layoutConversion.elements],
      assets: [],
      warnings: [],
      unsupported: [],
      ...(layoutConversion.slideNumber ?? masterConversion.slideNumber
        ? { slideNumber: layoutConversion.slideNumber ?? masterConversion.slideNumber }
        : {}),
      ...(layoutConversion.contentMargins ?? masterConversion.contentMargins
        ? { contentMargins: layoutConversion.contentMargins ?? masterConversion.contentMargins }
        : {}),
    };
  }

  const contentMaster: SlideMasterDefinition = {
    id: `${options.id}-content`,
    role: "content",
    name: "Content",
    background: {
      color: contentLayoutPart
        ? resolveBackgroundToken(await pkg.xml(contentLayoutPart), theme, extractedTheme, colorMap)
        : masterBackground,
    },
    contentMargins: contentConversion.contentMargins ?? defaultMargins(slideSize),
    elements: contentConversion.elements,
    ...(contentConversion.slideNumber ? { slideNumber: contentConversion.slideNumber } : {}),
  };
  masters.push(contentMaster);

  for (const [role, layoutPart] of layoutByRole) {
    if (layoutPart === contentLayoutPart) continue;
    const layoutXml = await pkg.xml(layoutPart);
    const conversion = await convertPart(layoutPart, layoutXml, `${role}-`);
    unsupported.push(...conversion.unsupported);
    warnings.push(...conversion.warnings);
    assets.push(...conversion.assets);
    masters.push({
      id: `${options.id}-${role}`,
      role,
      name: role === "title" ? "Title" : "Section",
      background: {
        color: resolveBackgroundToken(layoutXml, theme, extractedTheme, colorMap),
      },
      contentMargins: conversion.contentMargins ?? titleMargins(slideSize),
      elements: [...masterConversion.elements, ...conversion.elements],
    });
  }

  const draft = {
    schemaVersion: TEMPLATE_SCHEMA_VERSION,
    id: options.id,
    version: "1.0.0",
    name: options.name ?? `${options.id} (imported)`,
    theme,
    ...(slideSize.widthIn !== DEFAULT_SLIDE_WIDTH_IN || slideSize.heightIn !== DEFAULT_SLIDE_HEIGHT_IN
      ? { slideSize }
      : {}),
    masters,
    defaultMasterId: contentMaster.id,
    ...(masters.some((m) => m.role === "title") ? { titleMasterId: `${options.id}-title` } : {}),
    ...(masters.some((m) => m.role === "section")
      ? { sectionMasterId: `${options.id}-section` }
      : {}),
    ...(assets.length > 0
      ? { assets: assets.map((a) => ({ id: a.id, path: a.path, kind: "image" as const })) }
      : {}),
    metadata: {
      ...(options.organization ? { organization: options.organization } : {}),
      ...(options.department ? { department: options.department } : {}),
      ...(options.confidentialityLabel
        ? { confidentialityLabel: options.confidentialityLabel }
        : {}),
      description: `${sourceLabel} から取り込んだ Template`,
    },
    import: {
      importedFrom: sourceLabel,
      ...(unsupported.length > 0 ? { unsupportedFeatures: [...new Set(unsupported)] } : {}),
      ...(warnings.length > 0 ? { warnings: [...new Set(warnings)] } : {}),
    },
  };

  // 生成物も必ず自分のスキーマを通す（取り込み結果が壊れていたらここで落ちる）
  const template = parseTemplate(draft);

  return {
    template,
    assets,
    report: {
      schemaVersion: TEMPLATE_IMPORT_REPORT_VERSION,
      source: sourceLabel,
      templateId: options.id,
      slideSize,
      extracted: {
        themeColors: Object.keys(extractedTheme?.colors ?? {}).length,
        masters: masters.length,
        elements: masters.reduce((sum, m) => sum + (m.elements?.length ?? 0), 0),
        assets: assets.length,
      },
      unsupportedFeatures: [...new Set(unsupported)],
      warnings: [...new Set(warnings)],
    },
  };
}

function defaultMargins(slideSize: { widthIn: number; heightIn: number }) {
  const horizontal = round(slideSize.widthIn * 0.06);
  const vertical = round(slideSize.heightIn * 0.075);
  return { top: vertical, right: horizontal, bottom: vertical, left: horizontal };
}

function titleMargins(slideSize: { widthIn: number; heightIn: number }) {
  const horizontal = round(slideSize.widthIn * 0.09);
  return {
    top: round(slideSize.heightIn * 0.25),
    right: horizontal,
    bottom: round(slideSize.heightIn * 0.18),
    left: horizontal,
  };
}

function resolveBackgroundToken(
  xml: { children: unknown } & Parameters<typeof findPath>[0],
  theme: PresentationTheme,
  extractedTheme: ExtractedTheme | undefined,
  colorMap: Record<string, string>,
): "background" | "primary" | "surface" | "text" | "secondary" | "accent" {
  const bg = findPath(xml, "cSld", "bg");
  if (!bg) return "background";
  const solid = findPath(bg, "bgPr", "solidFill");
  const srgb = solid ? findChild(solid, "srgbClr") : undefined;
  const scheme = solid ? findChild(solid, "schemeClr") : undefined;
  const hex =
    attr(srgb, "val") ??
    (scheme ? extractedTheme?.colors[colorMap[attr(scheme, "val") ?? ""] ?? ""] : undefined);
  if (!hex) return "background";
  const nearest = nearestColorToken(resolveTheme(theme), hex, [
    "background",
    "surface",
    "primary",
    "secondary",
    "accent",
    "text",
  ]);
  return nearest.token as "background" | "primary" | "surface" | "text" | "secondary" | "accent";
}

/** 取り込んだ Template を `templates/<id>/` へ書き出す */
export async function writeTemplatePackage(
  rootDir: string,
  imported: ImportedTemplate,
): Promise<{ dir: string; templatePath: string; reportPath: string; assetPaths: string[] }> {
  const dir = join(rootDir, imported.template.id);
  await mkdir(dir, { recursive: true });

  const templatePath = join(dir, "template.json");
  await writeFile(templatePath, `${JSON.stringify(imported.template, null, 2)}\n`, "utf8");

  const reportPath = join(dir, "import-report.json");
  await writeFile(reportPath, `${JSON.stringify(imported.report, null, 2)}\n`, "utf8");

  const assetPaths: string[] = [];
  if (imported.assets.length > 0) {
    await mkdir(join(dir, "assets"), { recursive: true });
    for (const asset of imported.assets) {
      const path = join(dir, asset.path);
      await writeFile(path, asset.bytes);
      assetPaths.push(path);
    }
  }

  return { dir, templatePath, reportPath, assetPaths };
}
