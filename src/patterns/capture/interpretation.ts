import { z } from "zod";
import type { ExtractedElement } from "../../pptx/model.js";
import type { Rect } from "../../render/geometry.js";
import {
  canRenderAsAccent,
  canRenderAsBackground,
  canRenderAsDivider,
  canRenderAsText,
  describeElement,
  isTemplatePlaceholder,
} from "./element-roles.js";
import { inferRepeatedCellRoles, type HeuristicCell } from "./heuristic-roles.js";
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
 *
 * 契約を守るため、解釈は 2 段階で検証される。
 *
 *   1. `parseCaptureInterpretation()`                   … 形式・範囲・重複・セル構造
 *   2. `validateCaptureInterpretationAgainstElements()` … 役割 ↔ 要素種別の整合
 *
 * どの経路（human / heuristic / 将来の AI）も、両方をまとめて通す
 * `parseAndValidateCaptureInterpretation()` を唯一の入口として使うこと。
 */

/** 現行の Capture Interpretation スキーマ。divider は単数 (`dividerIndex`)。 */
export const CAPTURE_INTERPRETATION_SCHEMA_VERSION = "1.1";

/**
 * 読み込みだけ受け付ける旧バージョン。
 * 内部モデルへ移行してから使い、移行した痕跡は `notes` に残す。
 */
export const MIGRATABLE_CAPTURE_INTERPRETATION_VERSIONS = ["1.0"] as const;

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
    /** 説明文。複数指定できるが、実際に描画されるのは先頭のみ。 */
    descriptionIndexes?: number[];
  };
  /** Structural: 内容を持たない装飾 */
  decoration?: {
    accentIndexes?: number[];
    /**
     * 区切り線は 1 セルに 1 本。
     * Compiler が 1 本しか描けないため、スキーマ側も単数にしてある
     * （表現できない指定を受け取って黙って捨てないため）。
     */
    dividerIndex?: number;
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
const MAX_NOTES = 50;

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
        dividerIndex: elementIndex.optional(),
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
    notes: z.array(z.string().max(400)).max(MAX_NOTES).optional(),
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

// ---------------------------------------------------------------------------
// セルの構造シグネチャ — 「何を同じ Visual Structure とみなすか」
// ---------------------------------------------------------------------------

/**
 * Compiler は 1 セル目を反復テンプレートの基準にする。
 * ここが揃っていないと「解釈では divider を持たないセルにも divider が描かれる」
 * ような *意味の書き換え* が起きるため、Content Role だけでなく
 * Structural Role（background / accent / divider）も必ず比較する。
 */
export interface CellRoleSignature {
  background: boolean;
  index: boolean;
  /** label は必須なので常に true。比較対象を明示するために持つ。 */
  label: boolean;
  metric: boolean;
  description: boolean;
  /** accent は個数まで一致していないと同じ見た目の反復にならない */
  accentCount: number;
  divider: boolean;
}

export function cellRoleSignature(cell: InterpretedCell): CellRoleSignature {
  return {
    background: cell.backgroundIndex !== undefined,
    index: cell.content.indexIndex !== undefined,
    label: true,
    metric: cell.content.metricIndex !== undefined,
    description: (cell.content.descriptionIndexes ?? []).length > 0,
    accentCount: (cell.decoration?.accentIndexes ?? []).length,
    divider: cell.decoration?.dividerIndex !== undefined,
  };
}

/** 比較用の決定論的なキー */
export function cellRoleSignatureKey(cell: InterpretedCell): string {
  const signature = cellRoleSignature(cell);
  return (Object.keys(signature) as Array<keyof CellRoleSignature>)
    .map((key) => `${key}=${String(signature[key])}`)
    .join(",");
}

/** 差分だけを人が読める形で返す（"divider: expected=true, actual=false"） */
export function describeSignatureDiff(
  expected: CellRoleSignature,
  actual: CellRoleSignature,
): string[] {
  return (Object.keys(expected) as Array<keyof CellRoleSignature>)
    .filter((key) => expected[key] !== actual[key])
    .map((key) => `${key}: expected=${String(expected[key])}, actual=${String(actual[key])}`);
}

// ---------------------------------------------------------------------------
// バージョン移行（1.0 → 1.1）
// ---------------------------------------------------------------------------

const isPlainObject = (value: unknown): value is Record<string, unknown> =>
  typeof value === "object" && value !== null && !Array.isArray(value);

/**
 * 旧形式の `dividerIndexes: number[]` を現行の `dividerIndex?: number` へ正規化する。
 *
 * 空配列は「divider 無し」、1 件はその 1 本。
 * **2 件以上は黙って先頭を採らない**。Compiler が 1 本しか描けない以上、
 * どれを使うかは人（や AI）に決め直してもらう方が安全なため、明示エラーにする。
 */
function migrateToCurrentVersion(input: unknown): { value: unknown; notes: string[] } {
  if (!isPlainObject(input)) return { value: input, notes: [] };
  if (input.schemaVersion !== "1.0") return { value: input, notes: [] };

  const issues: string[] = [];
  const cells = Array.isArray(input.cells)
    ? input.cells.map((cell, i) => migrateCell(cell, i, issues))
    : input.cells;

  if (issues.length > 0) {
    throw new CaptureInterpretationError(
      `Capture Interpretation 1.0 を ${CAPTURE_INTERPRETATION_SCHEMA_VERSION} へ移行できません (${issues.length} 件)`,
      issues,
    );
  }

  return {
    value: { ...input, schemaVersion: CAPTURE_INTERPRETATION_SCHEMA_VERSION, cells },
    notes: [
      `Capture Interpretation 1.0 から ${CAPTURE_INTERPRETATION_SCHEMA_VERSION} へ移行しました (dividerIndexes → dividerIndex)`,
    ],
  };
}

function migrateCell(cell: unknown, cellIndex: number, issues: string[]): unknown {
  if (!isPlainObject(cell)) return cell;
  const decoration = cell.decoration;
  if (!isPlainObject(decoration) || !("dividerIndexes" in decoration)) return cell;

  const { dividerIndexes, ...rest } = decoration;
  const path = `cells[${cellIndex}].decoration.dividerIndexes`;

  if (!Array.isArray(dividerIndexes)) {
    issues.push(`${path}: 1.0 では配列である必要があります`);
    return cell;
  }
  if (dividerIndexes.length > 1) {
    issues.push(
      `${path}: Capture Interpretation ${CAPTURE_INTERPRETATION_SCHEMA_VERSION} では divider は単数です。` +
        ` [${dividerIndexes.join(", ")}] のうち 1 本を選び、decoration.dividerIndex に指定してください`,
    );
    return cell;
  }

  return {
    ...cell,
    decoration:
      dividerIndexes.length === 1 ? { ...rest, dividerIndex: dividerIndexes[0] } : { ...rest },
  };
}

// ---------------------------------------------------------------------------
// 検証（1）形式・範囲・重複・セル構造
// ---------------------------------------------------------------------------

/**
 * 解釈の形式を検証する。
 *
 * 要素の *種別* との整合は見ない（要素配列が必要なため）。
 * Compiler へ渡す前は必ず `parseAndValidateCaptureInterpretation()` を使うこと。
 *
 * @param elementCount 抽出された要素数（インデックスの上限チェックに使う）
 */
export function parseCaptureInterpretation(
  input: unknown,
  elementCount: number,
): CaptureInterpretation {
  const migrated = migrateToCurrentVersion(input);
  const parsed = CaptureInterpretationSchema.safeParse(migrated.value);
  if (!parsed.success) {
    throw new CaptureInterpretationError(
      `Capture Interpretation の形式が不正です (${parsed.error.issues.length} 件)`,
      parsed.error.issues.map((i) => `${i.path.join(".") || "(root)"}: ${i.message}`),
    );
  }
  const interpretation = parsed.data as CaptureInterpretation;

  if (interpretation.schemaVersion !== CAPTURE_INTERPRETATION_SCHEMA_VERSION) {
    throw new CaptureInterpretationError(
      `未対応の Capture Interpretation schemaVersion です (指定=${interpretation.schemaVersion}, 対応=${CAPTURE_INTERPRETATION_SCHEMA_VERSION}, 自動移行=${MIGRATABLE_CAPTURE_INTERPRETATION_VERSIONS.join(" / ")})`,
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
    claim(cell.decoration?.dividerIndex, {
      scope: "cell",
      cell: i,
      kind: "structural",
      role: "divider",
    });
  });

  for (const index of interpretation.staticIndexes) {
    claim(index, { scope: "static" });
  }

  // セルは同じ構造でなければ「反復」として扱えない。
  // Content Role だけでなく Structural Role（background / accent / divider）も比較する。
  if (interpretation.cells.length === 1) {
    issues.push("反復セルは 2 件以上必要です（1 件だけなら staticIndexes へ入れてください）");
  }
  const first = interpretation.cells[0];
  if (first) {
    const expected = cellRoleSignature(first);
    interpretation.cells.forEach((cell, i) => {
      if (i === 0) return;
      const diff = describeSignatureDiff(expected, cellRoleSignature(cell));
      if (diff.length > 0) {
        issues.push(
          `cells[${i}] の Visual Role 構成が cells[0] と一致しません: ${diff.join(" / ")}`,
        );
      }
    });
  }

  if (issues.length > 0) {
    throw new CaptureInterpretationError(
      `Capture Interpretation の内容が不正です (${issues.length} 件)`,
      issues,
    );
  }

  if (migrated.notes.length > 0) {
    interpretation.notes = [...(interpretation.notes ?? []), ...migrated.notes].slice(0, MAX_NOTES);
  }

  return interpretation;
}

// ---------------------------------------------------------------------------
// 検証（2）役割 ↔ 要素種別の整合
// ---------------------------------------------------------------------------

/** 役割ごとの「描画できる要素」の契約。判定は element-roles.ts と共有する。 */
const ROLE_REQUIREMENTS = {
  text: { can: canRenderAsText, need: "テキストを持つ要素が必要です" },
  background: { can: canRenderAsBackground, need: "塗り (solid fill) を持つ図形が必要です" },
  accent: { can: canRenderAsAccent, need: "塗りを持つ図形または線が必要です" },
  divider: { can: canRenderAsDivider, need: "線 (line) 要素が必要です" },
} as const satisfies Record<string, { can: (element: ExtractedElement) => boolean; need: string }>;

type RoleRequirement = keyof typeof ROLE_REQUIREMENTS;

/**
 * 「その要素をその役割として描けるか」を検証する。
 *
 * 見るのは客観的な描画可能性だけで、意味の妥当性は判定しない。
 * 例えば "30%" を label と解釈するのは（不自然でも）text である以上通す。
 * それを判断するのは human / heuristic / 将来の AI の責務。
 *
 * staticIndexes は「意味の役割へ分類しなかったが元デザインとして残す要素」なので
 * text / image / shape / line を広く許可する（描画できないものは Compiler が警告する）。
 */
export function validateCaptureInterpretationAgainstElements(
  interpretation: CaptureInterpretation,
  elements: readonly ExtractedElement[],
): CaptureInterpretation {
  const issues: string[] = [];

  const requireRole = (
    index: number | undefined,
    path: string,
    role: RoleRequirement,
    roleLabel: string,
  ): void => {
    if (index === undefined) return;
    const element = elements[index];
    if (!element) {
      issues.push(`${path}: 存在しない要素インデックスです (${index} >= ${elements.length})`);
      return;
    }
    const requirement = ROLE_REQUIREMENTS[role];
    if (!requirement.can(element)) {
      issues.push(
        `${path}: 要素 ${index} (${describeElement(element)}) は ${roleLabel} として利用できません (${requirement.need})`,
      );
    }
  };

  requireRole(interpretation.headlineIndex, "headlineIndex", "text", "headline");
  requireRole(interpretation.sublineIndex, "sublineIndex", "text", "subline");
  requireRole(interpretation.eyebrowIndex, "eyebrowIndex", "text", "eyebrow");
  requireRole(interpretation.sourceTextIndex, "sourceTextIndex", "text", "sourceText");

  interpretation.cells.forEach((cell, i) => {
    const at = `cells[${i}]`;
    requireRole(cell.backgroundIndex, `${at}.backgroundIndex`, "background", "background");
    requireRole(cell.content.indexIndex, `${at}.content.indexIndex`, "text", "index");
    requireRole(cell.content.labelIndex, `${at}.content.labelIndex`, "text", "label");
    requireRole(cell.content.metricIndex, `${at}.content.metricIndex`, "text", "metric");
    (cell.content.descriptionIndexes ?? []).forEach((index, j) => {
      requireRole(index, `${at}.content.descriptionIndexes[${j}]`, "text", "description");
    });
    (cell.decoration?.accentIndexes ?? []).forEach((index, j) => {
      requireRole(index, `${at}.decoration.accentIndexes[${j}]`, "accent", "accent");
    });
    requireRole(
      cell.decoration?.dividerIndex,
      `${at}.decoration.dividerIndex`,
      "divider",
      "divider",
    );
  });

  if (issues.length > 0) {
    throw new CaptureInterpretationError(
      `Capture Interpretation の役割が要素の種類と一致しません (${issues.length} 件)`,
      issues,
    );
  }

  return interpretation;
}

/**
 * **Compiler へ渡してよい解釈を作る唯一の入口。**
 *
 * human / heuristic / 将来の AI のどの経路でも必ずここを通す。
 * 形式検証と要素整合検証の両方を通過した解釈だけが Compiler へ到達する。
 */
export function parseAndValidateCaptureInterpretation(
  input: unknown,
  elements: readonly ExtractedElement[],
): CaptureInterpretation {
  return validateCaptureInterpretationAgainstElements(
    parseCaptureInterpretation(input, elements.length),
    elements,
  );
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
    put(cell.decoration?.dividerIndex, {
      scope: "cell",
      cell: i,
      kind: "structural",
      role: "divider",
    });
  });
  for (const index of interpretation.staticIndexes) put(index, { scope: "static" });
  return assigned;
}


// ---------------------------------------------------------------------------
// Heuristic → Interpretation（AI 無しの既定経路）
// ---------------------------------------------------------------------------

/**
 * 既存のヒューリスティック推定結果 → Interpretation。
 *
 * ここが「抽出結果から役割を決める」唯一の場所で、**Compiler は関与しない**。
 * Phase 3 で、反復カードを横断した強い規則（連番・数値・フォント階層・
 * 縁の帯・繰り返される線）だけを `heuristic-roles.ts` で判定するようにした。
 *
 * 方針は Precision 優先。根拠が弱ければ役割を付けず、
 * 要素を「固定デザイン」として元の見た目のまま残し、理由を notes に書く。
 * 一般的なスライドの意味理解はここでは扱わない（Phase 4 の AI 解釈の領域）。
 */
export function interpretFromSlotInference(
  elements: readonly ExtractedElement[],
  inference: SlotInferenceResult,
): CaptureInterpretation {
  const notes = [...inference.notes];
  /** 役割として使えなかったため固定デザインへ回す要素 */
  const demoted: number[] = [];

  /** その要素をその役割として描けるときだけ採用する */
  const usableAs = (
    index: number | undefined,
    role: RoleRequirement,
    label: string,
  ): number | undefined => {
    if (index === undefined) return undefined;
    const element = elements[index];
    if (element && ROLE_REQUIREMENTS[role].can(element)) return index;
    if (element) {
      demoted.push(index);
      notes.push(
        `要素 ${index} (${describeElement(element)}) は ${label} として描画できないため、固定デザインとして残しました`,
      );
    }
    return undefined;
  };

  const headlineIndex = usableAs(inference.headlineIndex, "text", "headline");
  const sublineIndex = usableAs(inference.sublineIndex, "text", "subline");
  const eyebrowIndex = usableAs(inference.eyebrowIndex, "text", "eyebrow");
  const sourceTextIndex = usableAs(inference.sourceIndex, "text", "sourceText");

  const claimedBySlide = new Set(
    [headlineIndex, sublineIndex, eyebrowIndex, sourceTextIndex].filter(
      (index): index is number => index !== undefined,
    ),
  );

  // --- 反復セルの下ごしらえ（領域と、役割になりうるテキスト） -----------------
  const heuristicCells: HeuristicCell[] = [];
  for (const cell of inference.cells) {
    const textIndexes = cell.textIndexes.filter((index) => {
      const element = elements[index];
      return element !== undefined && canRenderAsText(element) && !claimedBySlide.has(index);
    });
    if (textIndexes.length === 0) continue;
    const backgroundIndex = usableAs(cell.backgroundIndex, "background", "background");
    const box =
      backgroundIndex !== undefined
        ? elements[backgroundIndex]!.box
        : boundingBoxOf(textIndexes.map((index) => elements[index]!.box));
    heuristicCells.push({
      box,
      ...(backgroundIndex !== undefined ? { backgroundIndex } : {}),
      textIndexes,
    });
  }

  const claimedByCells = new Set<number>();
  for (const cell of heuristicCells) {
    if (cell.backgroundIndex !== undefined) claimedByCells.add(cell.backgroundIndex);
    for (const index of cell.textIndexes) claimedByCells.add(index);
  }

  // 装飾（accent / divider）の候補は、まだどの役割にもなっていない要素だけ
  const structuralCandidates = elements
    .map((_, index) => index)
    .filter(
      (index) =>
        !claimedBySlide.has(index) &&
        !claimedByCells.has(index) &&
        !isTemplatePlaceholder(elements[index]!),
    );

  const roles = inferRepeatedCellRoles(elements, heuristicCells, structuralCandidates);
  notes.push(...roles.notes);

  const inferredContent = roles.content;
  const cells: InterpretedCell[] = inferredContent
    ? heuristicCells.map((cell, c) => {
        const content = inferredContent[c]!;
        const structural = roles.structural[c]!;
        const decoration = {
          ...(structural.accentIndexes.length > 0
            ? { accentIndexes: structural.accentIndexes }
            : {}),
          ...(structural.dividerIndex !== undefined
            ? { dividerIndex: structural.dividerIndex }
            : {}),
        };
        return {
          ...(cell.backgroundIndex !== undefined ? { backgroundIndex: cell.backgroundIndex } : {}),
          content: {
            ...(content.indexIndex !== undefined ? { indexIndex: content.indexIndex } : {}),
            labelIndex: content.labelIndex,
            ...(content.metricIndex !== undefined ? { metricIndex: content.metricIndex } : {}),
            ...(content.descriptionIndexes.length > 0
              ? { descriptionIndexes: content.descriptionIndexes }
              : {}),
          },
          ...(Object.keys(decoration).length > 0 ? { decoration } : {}),
        };
      })
    : [];

  // セルごとに役割の構成がずれる場合は反復として扱えないので、解釈を空にする
  const signatures = new Set(cells.map((cell) => cellRoleSignatureKey(cell)));
  const usable = cells.length >= 2 && signatures.size === 1;
  if (cells.length >= 2 && !usable) {
    notes.push(
      "セルごとに推定された役割の構成が揃わなかったため、反復セルとして扱いませんでした",
    );
  }

  // 役割として採用した要素は固定デザインへ残さない（二重割り当てになるため）
  const claimed = new Set<number>(claimedBySlide);
  if (usable) {
    for (const cell of cells) {
      for (const index of interpretedCellIndexes(cell)) claimed.add(index);
    }
  }

  const staticIndexes = [
    ...new Set([...inference.staticIndexes, ...demoted, ...(usable ? [] : [...claimedByCells])]),
  ]
    .filter(
      (index) =>
        index < elements.length && !claimed.has(index) && !isTemplatePlaceholder(elements[index]!),
    )
    .sort((a, b) => a - b);

  const dropped = elements
    .map((_, index) => index)
    .filter((index) => isTemplatePlaceholder(elements[index]!) && !claimed.has(index));
  if (dropped.length > 0) {
    notes.push(
      `ページ番号・フッターのプレースホルダ（要素 ${dropped.join(", ")}）は Template の責務なので Pattern へ取り込みませんでした`,
    );
  }

  return {
    schemaVersion: CAPTURE_INTERPRETATION_SCHEMA_VERSION,
    interpretationSource: "heuristic",
    ...(headlineIndex !== undefined ? { headlineIndex } : {}),
    ...(sublineIndex !== undefined ? { sublineIndex } : {}),
    ...(eyebrowIndex !== undefined ? { eyebrowIndex } : {}),
    ...(sourceTextIndex !== undefined ? { sourceTextIndex } : {}),
    cells: usable ? cells : [],
    staticIndexes,
    notes: notes.slice(0, MAX_NOTES),
  };
}

/** そのセルが占有している要素（Content + Structural のすべて） */
function interpretedCellIndexes(cell: InterpretedCell): number[] {
  return [
    ...(cell.backgroundIndex !== undefined ? [cell.backgroundIndex] : []),
    ...(cell.content.indexIndex !== undefined ? [cell.content.indexIndex] : []),
    cell.content.labelIndex,
    ...(cell.content.metricIndex !== undefined ? [cell.content.metricIndex] : []),
    ...(cell.content.descriptionIndexes ?? []),
    ...(cell.decoration?.accentIndexes ?? []),
    ...(cell.decoration?.dividerIndex !== undefined ? [cell.decoration.dividerIndex] : []),
  ];
}

function boundingBoxOf(boxes: readonly Rect[]): Rect {
  const x = Math.min(...boxes.map((b) => b.x));
  const y = Math.min(...boxes.map((b) => b.y));
  const right = Math.max(...boxes.map((b) => b.x + b.w));
  const bottom = Math.max(...boxes.map((b) => b.y + b.h));
  return { x, y, w: right - x, h: bottom - y };
}
