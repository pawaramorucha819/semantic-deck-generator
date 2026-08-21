import { z } from "zod";
import type { ExtractedElement } from "../../pptx/model.js";
import type { SlotInferenceResult } from "./infer.js";

/**
 * Capture Interpretation — 「抽出した要素が何を意味するか」の中間表現。
 *
 * Pattern Capture は 3 段に分かれる。
 *
 *   1. Extraction     … PowerPoint から機械的に取れる事実
 *                       （element index / text / geometry / font / color / Z-order）
 *   2. Interpretation … その要素が何なのかの解釈  ← このファイル
 *   3. Compilation    … 解釈を Declarative Pattern へ変換（build.ts）
 *
 * 解釈の出どころは差し替えられる。
 *
 *   heuristic（既定・AI 不要） ─┐
 *   human（--interpretation）  ─┼→ CaptureInterpretation → Pattern Compiler
 *   将来の AI                  ─┘
 *
 * **Compiler は解釈をしない**。ここに書かれた役割をそのまま構造へ写すだけ。
 * だからこのスキーマが、将来 AI に返させる Structured Output の契約になる。
 */

export const CAPTURE_INTERPRETATION_SCHEMA_VERSION = "1.0";

/**
 * Content Role — 「何を伝える要素か」。
 * 再利用時に別の内容へ差し替わる部分。
 */
export const CONTENT_ROLES = ["index", "label", "metric", "description"] as const;
export type ContentRole = (typeof CONTENT_ROLES)[number];

/**
 * Visual Structural Role — 「どう見せるかの枠組み」。
 * 内容ではなくカードの構造そのもの。
 */
export const STRUCTURAL_ROLES = ["background", "accent", "divider"] as const;
export type StructuralRole = (typeof STRUCTURAL_ROLES)[number];

/**
 * 解釈の出どころ。
 * 将来 "ai" / "human-reviewed-ai" を足せるようにしてあるが、今は未実装。
 */
export const INTERPRETATION_SOURCES = ["heuristic", "human"] as const;
export type InterpretationSource = (typeof INTERPRETATION_SOURCES)[number];

/** 反復セル（カード 1 枚）の解釈 */
export interface InterpretedCell {
  /** Structural: カードの背景図形 */
  backgroundIndex?: number;
  /** Content: 差し替わる中身 */
  content: {
    /** 連番ラベル（01 / 02 …）。値は再利用時に item の順番から作る。 */
    indexIndex?: number;
    /** 見出し語。セルに必ず 1 つ必要。 */
    labelIndex: number;
    /** 数値・強調値 */
    metricIndex?: number;
    /** 説明文。複数指定できるが、実際に描画されるのは先頭のみ（下記参照）。 */
    descriptionIndexes?: number[];
  };
  /** Structural: 内容を持たない装飾 */
  decoration?: {
    accentIndexes?: number[];
    dividerIndexes?: number[];
  };
}

export interface CaptureInterpretation {
  schemaVersion: string;
  interpretationSource: InterpretationSource;

  /** スライド全体の役割（セルの外側） */
  headlineIndex?: number;
  sublineIndex?: number;
  eyebrowIndex?: number;
  sourceTextIndex?: number;

  cells: InterpretedCell[];
  /** 固定デザインとして残す要素（ロゴ・固定ラベルなど） */
  staticIndexes: number[];
  notes?: string[];
}

const MAX_CELLS = 12;
const MAX_DESCRIPTIONS = 4;
const MAX_DECORATIONS = 8;

const elementIndex = z.number().int().min(0).max(9999);

export const InterpretedCellSchema = z
  .object({
    backgroundIndex: elementIndex.optional(),
    content: z
      .object({
        indexIndex: elementIndex.optional(),
        labelIndex: elementIndex,
        metricIndex: elementIndex.optional(),
        descriptionIndexes: z.array(elementIndex).max(MAX_DESCRIPTIONS).optional(),
      })
      .strict(),
    decoration: z
      .object({
        accentIndexes: z.array(elementIndex).max(MAX_DECORATIONS).optional(),
        dividerIndexes: z.array(elementIndex).max(MAX_DECORATIONS).optional(),
      })
      .strict()
      .optional(),
  })
  .strict();

export const CaptureInterpretationSchema = z
  .object({
    schemaVersion: z.string().trim().min(1).max(20),
    interpretationSource: z.enum(INTERPRETATION_SOURCES),
    headlineIndex: elementIndex.optional(),
    sublineIndex: elementIndex.optional(),
    eyebrowIndex: elementIndex.optional(),
    sourceTextIndex: elementIndex.optional(),
    cells: z.array(InterpretedCellSchema).max(MAX_CELLS),
    staticIndexes: z.array(elementIndex).max(400),
    notes: z.array(z.string().max(400)).max(50).optional(),
  })
  .strict();

export class CaptureInterpretationError extends Error {
  constructor(
    summary: string,
    readonly issues: string[] = [],
  ) {
    super(issues.length > 0 ? `${summary}\n${issues.map((i) => `  - ${i}`).join("\n")}` : summary);
    this.name = "CaptureInterpretationError";
  }
}

/** どの要素がどの役割に割り当てられたかの逆引き（重複検出とレポート用） */
export type RoleAssignment =
  | { scope: "slide"; role: "headline" | "subline" | "eyebrow" | "sourceText" }
  | { scope: "cell"; cell: number; kind: "content"; role: ContentRole }
  | { scope: "cell"; cell: number; kind: "structural"; role: StructuralRole }
  | { scope: "static" };

export function describeAssignment(assignment: RoleAssignment): string {
  if (assignment.scope === "slide") return `slide.${assignment.role}`;
  if (assignment.scope === "static") return "static";
  return `cells[${assignment.cell}].${assignment.kind}.${assignment.role}`;
}

/**
 * 解釈を検証する。
 *
 * 将来 AI が返す JSON もここを必ず通るので、
 * 「Pattern Compiler へ壊れた解釈が到達しない」境界はこの関数が担う。
 *
 * @param elementCount 抽出された要素数（インデックスの上限チェックに使う）
 */
export function parseCaptureInterpretation(
  input: unknown,
  elementCount: number,
): CaptureInterpretation {
  const parsed = CaptureInterpretationSchema.safeParse(input);
  if (!parsed.success) {
    throw new CaptureInterpretationError(
      `Capture Interpretation の形式が不正です (${parsed.error.issues.length} 件)`,
      parsed.error.issues.map((i) => `${i.path.join(".") || "(root)"}: ${i.message}`),
    );
  }
  const interpretation = parsed.data as CaptureInterpretation;

  if (interpretation.schemaVersion !== CAPTURE_INTERPRETATION_SCHEMA_VERSION) {
    throw new CaptureInterpretationError(
      `未対応の Capture Interpretation schemaVersion です (指定=${interpretation.schemaVersion}, 対応=${CAPTURE_INTERPRETATION_SCHEMA_VERSION})`,
    );
  }

  const issues: string[] = [];
  const assigned = new Map<number, RoleAssignment>();

  const claim = (index: number | undefined, assignment: RoleAssignment): void => {
    if (index === undefined) return;
    if (index >= elementCount) {
      issues.push(
        `${describeAssignment(assignment)}: 存在しない要素インデックスです (${index} >= ${elementCount})`,
      );
      return;
    }
    const existing = assigned.get(index);
    if (existing) {
      issues.push(
        `要素 ${index} が複数の役割へ割り当てられています (${describeAssignment(existing)} と ${describeAssignment(assignment)})`,
      );
      return;
    }
    assigned.set(index, assignment);
  };

  claim(interpretation.headlineIndex, { scope: "slide", role: "headline" });
  claim(interpretation.sublineIndex, { scope: "slide", role: "subline" });
  claim(interpretation.eyebrowIndex, { scope: "slide", role: "eyebrow" });
  claim(interpretation.sourceTextIndex, { scope: "slide", role: "sourceText" });

  interpretation.cells.forEach((cell, i) => {
    claim(cell.backgroundIndex, { scope: "cell", cell: i, kind: "structural", role: "background" });
    claim(cell.content.indexIndex, { scope: "cell", cell: i, kind: "content", role: "index" });
    claim(cell.content.labelIndex, { scope: "cell", cell: i, kind: "content", role: "label" });
    claim(cell.content.metricIndex, { scope: "cell", cell: i, kind: "content", role: "metric" });
    for (const index of cell.content.descriptionIndexes ?? []) {
      claim(index, { scope: "cell", cell: i, kind: "content", role: "description" });
    }
    for (const index of cell.decoration?.accentIndexes ?? []) {
      claim(index, { scope: "cell", cell: i, kind: "structural", role: "accent" });
    }
    for (const index of cell.decoration?.dividerIndexes ?? []) {
      claim(index, { scope: "cell", cell: i, kind: "structural", role: "divider" });
    }
  });

  for (const index of interpretation.staticIndexes) {
    claim(index, { scope: "static" });
  }

  // セルは同じ構造でなければ「反復」として扱えない
  if (interpretation.cells.length === 1) {
    issues.push("反復セルは 2 件以上必要です（1 件だけなら staticIndexes へ入れてください）");
  }
  const shapes = new Set(
    interpretation.cells.map((cell) =>
      [
        cell.backgroundIndex !== undefined ? "bg" : "",
        cell.content.indexIndex !== undefined ? "index" : "",
        cell.content.metricIndex !== undefined ? "metric" : "",
        (cell.content.descriptionIndexes ?? []).length > 0 ? "desc" : "",
      ].join("|"),
    ),
  );
  if (shapes.size > 1) {
    issues.push(
      "セルごとに役割の構成が異なります（反復として扱えるよう、全セルで同じ役割を指定してください）",
    );
  }

  if (issues.length > 0) {
    throw new CaptureInterpretationError(
      `Capture Interpretation の内容が不正です (${issues.length} 件)`,
      issues,
    );
  }

  return interpretation;
}

/** 要素インデックス → 割り当てられた役割（Capture Report 用） */
export function assignmentsOf(
  interpretation: CaptureInterpretation,
): Map<number, RoleAssignment> {
  const assigned = new Map<number, RoleAssignment>();
  const put = (index: number | undefined, assignment: RoleAssignment): void => {
    if (index !== undefined) assigned.set(index, assignment);
  };

  put(interpretation.headlineIndex, { scope: "slide", role: "headline" });
  put(interpretation.sublineIndex, { scope: "slide", role: "subline" });
  put(interpretation.eyebrowIndex, { scope: "slide", role: "eyebrow" });
  put(interpretation.sourceTextIndex, { scope: "slide", role: "sourceText" });
  interpretation.cells.forEach((cell, i) => {
    put(cell.backgroundIndex, { scope: "cell", cell: i, kind: "structural", role: "background" });
    put(cell.content.indexIndex, { scope: "cell", cell: i, kind: "content", role: "index" });
    put(cell.content.labelIndex, { scope: "cell", cell: i, kind: "content", role: "label" });
    put(cell.content.metricIndex, { scope: "cell", cell: i, kind: "content", role: "metric" });
    for (const index of cell.content.descriptionIndexes ?? []) {
      put(index, { scope: "cell", cell: i, kind: "content", role: "description" });
    }
    for (const index of cell.decoration?.accentIndexes ?? []) {
      put(index, { scope: "cell", cell: i, kind: "structural", role: "accent" });
    }
    for (const index of cell.decoration?.dividerIndexes ?? []) {
      put(index, { scope: "cell", cell: i, kind: "structural", role: "divider" });
    }
  });
  for (const index of interpretation.staticIndexes) put(index, { scope: "static" });
  return assigned;
}

// ---------------------------------------------------------------------------
// Heuristic → Interpretation（AI 無しの既定経路）
// ---------------------------------------------------------------------------

const NUMERIC_PATTERN = /^[\s0-9.,+\-−%¥$€£×xX億万千人年月日時間件社]+$/;

/** 数値・指標らしい文字列か（ヒューリスティクスの判断であり、Compiler は使わない） */
export function looksNumeric(text: string): boolean {
  return /\d/.test(text) && NUMERIC_PATTERN.test(text.trim());
}

/**
 * 既存のヒューリスティック推定結果 → Interpretation。
 *
 * ここが「テキストの並びから label / metric / description を決める」唯一の場所。
 * Phase 2 では推定の賢さは変えず、判断の置き場所だけを Compiler から移した。
 * 「01 は連番だ」「この線は区切り線だ」といった判別はまだ行わない（Phase 3 の課題）。
 * 現時点でそれらを正しく解釈したい場合は、人間が interpretation JSON を書く。
 */
export function interpretFromSlotInference(
  elements: readonly ExtractedElement[],
  inference: SlotInferenceResult,
): CaptureInterpretation {
  const notes = [...inference.notes];
  const usableCells = inference.cells.filter((cell) => cell.textIndexes.length > 0);
  const cells: InterpretedCell[] = [];

  for (const cell of usableCells) {
    const texts = cell.textIndexes.filter((index) => elements[index]?.text?.plain);
    if (texts.length === 0) continue;

    let labelIndex: number | undefined;
    let metricIndex: number | undefined;
    const descriptionIndexes: number[] = [];

    for (const index of texts) {
      const plain = elements[index]?.text?.plain ?? "";
      if (labelIndex === undefined) {
        labelIndex = index;
        continue;
      }
      if (metricIndex === undefined && looksNumeric(plain)) {
        metricIndex = index;
        continue;
      }
      descriptionIndexes.push(index);
    }

    if (labelIndex === undefined) continue;
    cells.push({
      ...(cell.backgroundIndex !== undefined ? { backgroundIndex: cell.backgroundIndex } : {}),
      content: {
        labelIndex,
        ...(metricIndex !== undefined ? { metricIndex } : {}),
        ...(descriptionIndexes.length > 0 ? { descriptionIndexes } : {}),
      },
    });
  }

  // セルごとに役割の構成がずれる場合は反復として扱えないので、解釈を空にする
  const shapes = new Set(
    cells.map((cell) =>
      [
        cell.backgroundIndex !== undefined ? "bg" : "",
        cell.content.indexIndex !== undefined ? "index" : "",
        cell.content.metricIndex !== undefined ? "metric" : "",
        (cell.content.descriptionIndexes ?? []).length > 0 ? "desc" : "",
      ].join("|"),
    ),
  );
  const usable = cells.length >= 2 && shapes.size === 1;
  if (cells.length >= 2 && !usable) {
    notes.push(
      "セルごとに推定された役割の構成が揃わなかったため、反復セルとして扱いませんでした",
    );
  }

  const claimed = new Set<number>();
  const add = (index: number | undefined): void => {
    if (index !== undefined) claimed.add(index);
  };
  add(inference.headlineIndex);
  add(inference.sublineIndex);
  add(inference.eyebrowIndex);
  add(inference.sourceIndex);
  if (usable) {
    for (const cell of cells) {
      add(cell.backgroundIndex);
      add(cell.content.indexIndex);
      add(cell.content.labelIndex);
      add(cell.content.metricIndex);
      for (const index of cell.content.descriptionIndexes ?? []) add(index);
    }
  }

  // ヒューリスティクスは accent / divider を判別しない（Phase 3 の課題）。
  // 判別できなかった要素は「固定デザイン」として元の見た目のまま残す。
  const staticIndexes = [
    ...new Set([
      ...inference.staticIndexes,
      ...(usable ? [] : cells.flatMap((cell) => cellElementIndexes(cell))),
    ]),
  ]
    .filter((index) => !claimed.has(index) && index < elements.length)
    .sort((a, b) => a - b);

  return {
    schemaVersion: CAPTURE_INTERPRETATION_SCHEMA_VERSION,
    interpretationSource: "heuristic",
    ...(inference.headlineIndex !== undefined ? { headlineIndex: inference.headlineIndex } : {}),
    ...(inference.sublineIndex !== undefined ? { sublineIndex: inference.sublineIndex } : {}),
    ...(inference.eyebrowIndex !== undefined ? { eyebrowIndex: inference.eyebrowIndex } : {}),
    ...(inference.sourceIndex !== undefined ? { sourceTextIndex: inference.sourceIndex } : {}),
    cells: usable ? cells : [],
    staticIndexes,
    notes,
  };
}

function cellElementIndexes(cell: InterpretedCell): number[] {
  return [
    ...(cell.backgroundIndex !== undefined ? [cell.backgroundIndex] : []),
    ...(cell.content.indexIndex !== undefined ? [cell.content.indexIndex] : []),
    cell.content.labelIndex,
    ...(cell.content.metricIndex !== undefined ? [cell.content.metricIndex] : []),
    ...(cell.content.descriptionIndexes ?? []),
  ];
}

