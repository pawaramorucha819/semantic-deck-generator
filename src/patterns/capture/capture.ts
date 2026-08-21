import { mkdir, writeFile } from "node:fs/promises";
import { join } from "node:path";
import { readDeck } from "../../pptx/extract.js";
import { PptxPackage, RELATIONSHIP_TYPES } from "../../pptx/package.js";
import type { ExtractedElement, ExtractedSlide } from "../../pptx/model.js";
import { defaultTheme, type PresentationTheme } from "../../theme/theme.js";
import { parseDeclarativePattern, type DeclarativePatternDocument } from "../declarative/schema.js";
import { buildCapturedPattern } from "./build.js";
import {
  heuristicSlotInferencer,
  SlotInferenceError,
  validateInference,
  type SlotInferenceResult,
  type SlotInferencer,
} from "./infer.js";
import {
  assignmentsOf,
  describeAssignment,
  interpretFromSlotInference,
  parseCaptureInterpretation,
  type CaptureInterpretation,
} from "./interpretation.js";

/**
 * Pattern Capture — 人が PowerPoint で作った良いスライドを Pattern 化する。
 *
 *   PPTX → 要素抽出 → 解釈（Interpretation） → コンパイル
 *        → 下書き（Declarative Pattern JSON） → **人間レビュー** → 登録
 *
 * 解釈の出どころは差し替えられる。
 *   heuristic … 既定。幾何とフォントサイズだけを見る決定論的推定
 *   human     … --interpretation で JSON を渡す（Phase 2 で追加）
 *   ai        … 将来。同じ CaptureInterpretation スキーマを返させる（未実装）
 *
 * 「AI が勝手に Pattern を増やす」仕組みにはしない。
 * 出力される pattern.json は必ず `metadata.status = "draft"` であり、
 * 人がレビューして approved に変えるまで PatternRegistry は受け付けない。
 */

export const PATTERN_CAPTURE_REPORT_VERSION = "1.0";

/** 人間（や将来の AI）が解釈を書くための、要素 1 つぶんの目録 */
export interface CapturedElementSummary {
  index: number;
  /** 元 PPTX の描画順（Z-order）。index と同じ並び。 */
  sourceOrder: number;
  id: string;
  name?: string;
  kind: string;
  geometry?: string;
  box: { x: number; y: number; w: number; h: number };
  text?: string;
  fontSize?: number;
  fill?: string;
  line?: string;
  /** この要素に割り当てられた役割（未割り当てなら省略） */
  role?: string;
}

export interface PatternCaptureReport {
  schemaVersion: string;
  patternId: string;
  source: { file: string; slideNumber: number };
  elements: {
    total: number;
    text: number;
    shape: number;
    line: number;
    image: number;
    placeholder: number;
  };
  /** 抽出結果そのもの（解釈を書くための材料） */
  inventory: CapturedElementSummary[];
  /** 実際に採用された解釈 */
  interpretation: {
    source: string;
    inferencer?: string;
    headline?: string;
    cells: number;
    statics: number;
    /** セルごとの役割（どの要素をどう解釈したか） */
    cellRoles: Array<Record<string, number | number[] | undefined>>;
    notes: string[];
  };
  inference: {
    inferencer: string;
    headline?: string;
    cells: number;
    statics: number;
    notes: string[];
  };
  unsupportedFeatures: string[];
  warnings: string[];
  /** 人間レビューで必ず確認してほしい項目 */
  reviewChecklist: string[];
}

export interface CapturedPattern {
  document: DeclarativePatternDocument;
  report: PatternCaptureReport;
  assets: Array<{ id: string; path: string; bytes: Buffer }>;
}

/** CLI などから Slot 推定を上書きするための指定 */
export interface CaptureOverrides {
  headlineIndex?: number;
  sublineIndex?: number;
  eyebrowIndex?: number;
  sourceIndex?: number;
  /** 固定デザインとして扱う要素インデックス（推定に追加する） */
  staticIndexes?: number[];
  /** 反復セルを明示する */
  cells?: Array<{ backgroundIndex?: number; textIndexes: number[] }>;
}

export interface CapturePatternOptions {
  id: string;
  slideNumber: number;
  description?: string;
  organization?: string;
  author?: string;
  scope?: "company" | "team" | "project" | "user";
  /** 色・サイズを寄せる先の Theme（既定: ビルトイン default） */
  theme?: PresentationTheme;
  /** Slot 推定器（AI 実装へ差し替え可能。既定は決定論的ヒューリスティクス） */
  inferencer?: SlotInferencer;
  overrides?: CaptureOverrides;
  /**
   * 人間（将来は AI）が書いた解釈。検証前の値でよく、内部で必ず
   * `parseCaptureInterpretation()` を通す（untrusted input として扱う）。
   * 指定するとヒューリスティック推定は使われない。
   */
  interpretation?: unknown;
  /** 取り込み元の表示名 */
  sourceLabel?: string;
  /** 記録する取り込み日時（決定論性のため呼び出し側が渡す） */
  importedAt?: string;
}

function buildInventory(
  elements: readonly ExtractedElement[],
  interpretation: CaptureInterpretation,
): CapturedElementSummary[] {
  const assignments = assignmentsOf(interpretation);
  return elements.map((element, index) => {
    const assignment = assignments.get(index);
    const text = element.text?.plain.replace(/\n/g, " ").trim();
    return {
      index,
      sourceOrder: index,
      id: element.id,
      ...(element.name ? { name: element.name } : {}),
      kind: element.kind,
      ...(element.geometry ? { geometry: element.geometry } : {}),
      box: {
        x: Number(element.box.x.toFixed(3)),
        y: Number(element.box.y.toFixed(3)),
        w: Number(element.box.w.toFixed(3)),
        h: Number(element.box.h.toFixed(3)),
      },
      ...(text ? { text: text.slice(0, 120) } : {}),
      ...(element.text?.maxFontSize ? { fontSize: element.text.maxFontSize } : {}),
      ...(element.fill?.color ? { fill: element.fill.color } : {}),
      ...(element.line?.color ? { line: element.line.color } : {}),
      ...(assignment ? { role: describeAssignment(assignment) } : {}),
    };
  });
}

function cellRolesOf(interpretation: CaptureInterpretation): Array<
  Record<string, number | number[] | undefined>
> {
  return interpretation.cells.map((cell) => ({
    background: cell.backgroundIndex,
    index: cell.content.indexIndex,
    label: cell.content.labelIndex,
    metric: cell.content.metricIndex,
    descriptions: cell.content.descriptionIndexes,
    accents: cell.decoration?.accentIndexes,
    dividers: cell.decoration?.dividerIndexes,
  }));
}

function countKinds(elements: ExtractedElement[]): PatternCaptureReport["elements"] {
  return {
    total: elements.length,
    text: elements.filter((e) => e.kind === "text").length,
    shape: elements.filter((e) => e.kind === "shape").length,
    line: elements.filter((e) => e.kind === "line").length,
    image: elements.filter((e) => e.kind === "image").length,
    placeholder: elements.filter((e) => e.kind === "placeholder").length,
  };
}

/**
 * 推定結果に CLI 等からの明示指定を重ねる。
 *
 * 明示指定が常に優先で、同じ要素が別の役割にも推定されていた場合はそちらを外す。
 * 役割を外された要素は「固定デザイン」へ回す（元スライドの見た目を保つため）。
 */
function applyOverrides(
  inference: SlotInferenceResult,
  overrides: CaptureOverrides | undefined,
  elementCount: number,
): SlotInferenceResult {
  if (!overrides) return inference;

  const roles = ["headlineIndex", "sublineIndex", "eyebrowIndex", "sourceIndex"] as const;
  type RoleKey = (typeof roles)[number];

  const assignment = new Map<RoleKey, number | undefined>(
    roles.map((role) => [role, inference[role]]),
  );
  const displaced: number[] = [];
  const notes = [...inference.notes];
  const explicit: number[] = [];

  for (const role of roles) {
    const value = overrides[role];
    if (value === undefined) continue;
    if (value >= elementCount) {
      throw new SlotInferenceError(
        `${role}: 存在しない要素インデックスです (${value} >= ${elementCount})`,
      );
    }
    explicit.push(value);
    // 同じ要素が別の役割に割り当たっていたら外す
    for (const other of roles) {
      if (other !== role && assignment.get(other) === value) {
        assignment.set(other, undefined);
        notes.push(`要素 ${value} は ${role} として指定されたため ${other} から外しました`);
      }
    }
    const previous = assignment.get(role);
    if (previous !== undefined && previous !== value) displaced.push(previous);
    assignment.set(role, value);
  }

  let cells = overrides.cells ?? inference.cells;
  if (overrides.cells) {
    for (const cell of overrides.cells) {
      explicit.push(...cell.textIndexes);
      if (cell.backgroundIndex !== undefined) explicit.push(cell.backgroundIndex);
    }
  }
  if (explicit.length > 0) {
    notes.push(`CLI 指定により ${[...new Set(explicit)].sort((a, b) => a - b).join(", ")} 番の役割を上書きしました`);
  }

  const roleIndexes = new Set<number>(
    roles.map((role) => assignment.get(role)).filter((i): i is number => i !== undefined),
  );

  // 役割として明示指定された要素はセルから外す
  cells = cells
    .map((cell) => ({
      ...(cell.backgroundIndex !== undefined && !roleIndexes.has(cell.backgroundIndex)
        ? { backgroundIndex: cell.backgroundIndex }
        : {}),
      textIndexes: cell.textIndexes.filter((i) => !roleIndexes.has(i)),
    }))
    .filter((cell) => cell.textIndexes.length > 0);

  const cellIndexes = new Set<number>(
    cells.flatMap((c) => [
      ...(c.backgroundIndex !== undefined ? [c.backgroundIndex] : []),
      ...c.textIndexes,
    ]),
  );
  const taken = new Set<number>([...roleIndexes, ...cellIndexes]);

  const staticIndexes = [
    ...new Set([...(overrides.staticIndexes ?? []), ...inference.staticIndexes, ...displaced]),
  ]
    .filter((i) => !taken.has(i) && i < elementCount)
    .sort((a, b) => a - b);

  return {
    ...(assignment.get("headlineIndex") !== undefined
      ? { headlineIndex: assignment.get("headlineIndex")! }
      : {}),
    ...(assignment.get("sublineIndex") !== undefined
      ? { sublineIndex: assignment.get("sublineIndex")! }
      : {}),
    ...(assignment.get("eyebrowIndex") !== undefined
      ? { eyebrowIndex: assignment.get("eyebrowIndex")! }
      : {}),
    ...(assignment.get("sourceIndex") !== undefined
      ? { sourceIndex: assignment.get("sourceIndex")! }
      : {}),
    cells,
    staticIndexes,
    notes,
  };
}

/** PPTX の 1 スライドから Pattern 下書きを作る */
export async function capturePatternFromPptx(
  input: Buffer | string,
  options: CapturePatternOptions,
): Promise<CapturedPattern> {
  const pkg = await PptxPackage.open(input);
  const deck = await readDeck(pkg, { slideNumbers: [options.slideNumber], includeNotes: false });
  const slide: ExtractedSlide | undefined = deck.slides[0];
  if (!slide) {
    throw new Error(
      `スライド ${options.slideNumber} が見つかりません（このファイルのスライド数を確認してください）`,
    );
  }

  const sourceLabel =
    options.sourceLabel ?? (typeof input === "string" ? input : "(buffer)");

  // --- 解釈（Interpretation）を決める -------------------------------------
  // 人間の解釈が渡されたらそれを使い、無ければヒューリスティック推定を解釈へ変換する。
  // どちらの経路でも Compiler が受け取るのは同じ CaptureInterpretation。
  let inference: SlotInferenceResult | undefined;
  let interpretation: CaptureInterpretation;

  if (options.interpretation !== undefined) {
    if (options.overrides) {
      throw new SlotInferenceError(
        "interpretation と overrides は同時に指定できません（overrides はヒューリスティック推定への上書きです）",
      );
    }
    interpretation = parseCaptureInterpretation(options.interpretation, slide.elements.length);
    if (interpretation.interpretationSource !== "human") {
      throw new SlotInferenceError(
        `外部から渡せる interpretationSource は 'human' だけです (指定=${interpretation.interpretationSource})`,
      );
    }
  } else {
    const inferencer = options.inferencer ?? heuristicSlotInferencer;
    const rawInference = await inferencer.infer({
      slideNumber: slide.slideNumber,
      slideSize: deck.slideSize,
      elements: slide.elements,
    });
    const validated = validateInference(rawInference, slide.elements.length);
    inference = applyOverrides(validated, options.overrides, slide.elements.length);
    // 上書き後も整合していることを再検証する
    validateInference(inference, slide.elements.length);
    // ヒューリスティクスの結果も同じスキーマ検証を通す（経路を揃える）
    interpretation = parseCaptureInterpretation(
      interpretFromSlotInference(slide.elements, inference),
      slide.elements.length,
    );
  }

  // 画像アセットを取り出す
  const rels = await pkg.relationships(slide.part);
  const assets: CapturedPattern["assets"] = [];
  const imageAssetIds = new Map<number, string>();
  const warnings = [...slide.warnings];

  for (const [index, element] of slide.elements.entries()) {
    if (element.kind !== "image") continue;
    if (element.image?.externalTarget) {
      warnings.push("外部リンク画像は取り込みません（外部 URL を fetch しないため）");
      continue;
    }
    const target =
      element.image?.part ??
      (element.image?.relationshipId
        ? rels.get(element.image.relationshipId)?.target
        : undefined);
    if (!target || !pkg.has(target)) {
      warnings.push(`画像を解決できませんでした: ${element.name ?? element.id}`);
      continue;
    }
    const extension = target.slice(target.lastIndexOf(".")).toLowerCase();
    if (![".png", ".jpg", ".jpeg", ".gif"].includes(extension)) {
      warnings.push(`対応していない画像形式のため除外しました (${extension})`);
      continue;
    }
    const id = `image-${assets.length + 1}`;
    assets.push({ id, path: `assets/${id}${extension}`, bytes: await pkg.binary(target) });
    imageAssetIds.set(index, id);
  }

  const built = buildCapturedPattern(slide.elements, interpretation, {
    id: options.id,
    ...(options.description ? { description: options.description } : {}),
    ...(options.scope ? { scope: options.scope as never } : {}),
    ...(options.organization ? { organization: options.organization } : {}),
    ...(options.author ? { author: options.author } : {}),
    theme: options.theme ?? defaultTheme,
    slideSize: deck.slideSize,
    source: {
      file: sourceLabel,
      slideNumber: options.slideNumber,
      ...(options.importedAt ? { importedAt: options.importedAt } : {}),
    },
    unsupportedFeatures: [...deck.unsupported, ...slide.unsupported],
    warnings,
    assets: assets.map((a) => ({ id: a.id, path: a.path })),
    imageAssetIds,
  });

  // 生成物も必ず自分のスキーマを通す（壊れた下書きを書き出さない）
  const document = parseDeclarativePattern(built.document);

  const headlineText =
    interpretation.headlineIndex !== undefined
      ? slide.elements[interpretation.headlineIndex]?.text?.plain
      : undefined;
  const inferencerId =
    options.interpretation !== undefined
      ? undefined
      : (options.inferencer ?? heuristicSlotInferencer).id;

  const report: PatternCaptureReport = {
    schemaVersion: PATTERN_CAPTURE_REPORT_VERSION,
    patternId: options.id,
    source: { file: sourceLabel, slideNumber: options.slideNumber },
    elements: countKinds(slide.elements),
    inventory: buildInventory(slide.elements, interpretation),
    interpretation: {
      source: interpretation.interpretationSource,
      ...(inferencerId ? { inferencer: inferencerId } : {}),
      ...(headlineText ? { headline: headlineText } : {}),
      cells: interpretation.cells.length,
      statics: interpretation.staticIndexes.length,
      cellRoles: cellRolesOf(interpretation),
      notes: interpretation.notes ?? [],
    },
    // 後方互換: 既存の読み手のために従来の inference セクションも残す
    inference: {
      inferencer: inferencerId ?? "human",
      ...(headlineText ? { headline: headlineText } : {}),
      cells: interpretation.cells.length,
      statics: interpretation.staticIndexes.length,
      notes: interpretation.notes ?? [],
    },
    unsupportedFeatures: [...new Set([...deck.unsupported, ...slide.unsupported])],
    warnings: [...new Set(built.warnings)],
    reviewChecklist: [
      "inventory を見て、各要素の解釈（label / metric / description / accent / divider）が正しいか確認する",
      "違っていれば interpretation JSON を書き、--interpretation で指定し直す",
      "compatibility.logicalRelations / slideRoles を実際の用途に合わせて絞る",
      "match.baseScore と rules を調整する（既定は 0.5 の控えめな値）",
      "slot の source（keyMessage / supportingMessage / content / evidence）が意味と合っているか確認する",
      "固定テキスト（literal）として残った文字列が本当に毎回同じか確認する",
      "レビューが済んだら metadata.status を 'approved' に変更する",
    ],
  };

  return { document, report, assets };
}

/** 下書きを `patterns/<id>/` へ書き出す */
export async function writePatternPackage(
  rootDir: string,
  captured: { document: DeclarativePatternDocument; report?: PatternCaptureReport; assets?: CapturedPattern["assets"] },
): Promise<{ dir: string; patternPath: string; reportPath?: string; assetPaths: string[] }> {
  const dir = join(rootDir, captured.document.id);
  await mkdir(dir, { recursive: true });

  const patternPath = join(dir, "pattern.json");
  await writeFile(patternPath, `${JSON.stringify(captured.document, null, 2)}\n`, "utf8");

  let reportPath: string | undefined;
  if (captured.report) {
    reportPath = join(dir, "capture-report.json");
    await writeFile(reportPath, `${JSON.stringify(captured.report, null, 2)}\n`, "utf8");
  }

  const assetPaths: string[] = [];
  if (captured.assets && captured.assets.length > 0) {
    await mkdir(join(dir, "assets"), { recursive: true });
    for (const asset of captured.assets) {
      const path = join(dir, asset.path);
      await writeFile(path, asset.bytes);
      assetPaths.push(path);
    }
  }

  return { dir, patternPath, ...(reportPath ? { reportPath } : {}), assetPaths };
}

export { RELATIONSHIP_TYPES };
