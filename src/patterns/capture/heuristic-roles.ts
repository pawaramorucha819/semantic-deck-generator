import type { ExtractedElement } from "../../pptx/model.js";
import type { Rect } from "../../render/geometry.js";
import { canRenderAsAccent, canRenderAsDivider } from "./element-roles.js";

/**
 * Phase 3: 反復カードの「明白な役割」だけを決定論的に推定する。
 *
 * ここは Interpretation 層の一部であり、**Compiler は一切これを知らない**。
 * 推定できたものは `CaptureInterpretation` の役割になり、
 * 推定できなければ役割を付けずに固定デザインとして残す（abstain）。
 *
 * 設計方針は Recall ではなく **Precision**。
 *
 *  - 単独の要素だけを見て意味を決めない（"01" 1 つで index と決めない）
 *  - **反復セルを横断した一致**（相対位置・フォント・種別・書式）を根拠にする
 *  - 根拠が弱ければ推定せず、理由を notes に残す
 *
 * 「人が見ればまず迷わない」強い反復規則だけを対象にしている。
 * 一般的なスライドの意味理解はここでは扱わない（Phase 4 の AI 解釈の領域）。
 */

/**
 * 判断に使う閾値。散らかると意味が追えなくなるのでここへ集約する。
 * 外部設定にはしない（決定論性と再現性を優先する）。
 */
export const HEURISTIC_THRESHOLDS = {
  /**
   * セル内の相対位置のずれ許容量（セルの幅・高さに対する比率）。
   * PowerPoint で手作業に並べたカードは数 pt ずれるので、その程度は許す。
   */
  relativePositionTolerance: 0.06,
  /**
   * 同じ font tier とみなす相対差。17pt と 17.5pt のような揺れを吸収する。
   * これを超えたら「別の階層のテキスト」として扱う。
   */
  fontSizeTolerance: 0.12,
  /**
   * label が description より「明確に大きい」と言える最小比率。
   * これ未満は視覚的な階層として弱く、人でも迷うので abstain する
   * （「上にあるから label」という 1 条件だけで決めないため）。
   */
  labelToDescriptionFontRatio: 1.2,
  /** index / metric として扱うテキストの最大文字数。文章を弾く。 */
  maxNumericTextLength: 12,
  /** metric の数値以外（単位・記号）の最大文字数。"¥100万" までは許す。 */
  maxMetricUnitLength: 3,
  /** index として認める最大桁数。2024 のような西暦を構造的に弾く。 */
  maxOrdinalDigits: 3,
  /**
   * divider の「軸に沿っている度合い」。短辺が長辺のこの比率以下なら
   * 水平／垂直とみなす。斜めのコネクタや矢印はここで落ちる。
   */
  dividerMaxAspectRatio: 0.1,
  /** divider の相対長さがセル間でどれだけ違ってよいか。 */
  dividerLengthTolerance: 0.15,
  /** accent の「細さ」。セルの対応する辺に対してこの比率以下なら帯とみなす。 */
  accentMaxThicknessRatio: 0.12,
  /**
   * accent の「長さ」。帯はセルの辺をおおよそ覆う。
   * アイコン・バッジのような小さな装飾を弾くための下限。
   */
  accentMinLengthRatio: 0.5,
  /** accent がセルの縁に接していると言える距離（セルの辺に対する比率）。 */
  accentEdgeTolerance: 0.06,
} as const;

// ---------------------------------------------------------------------------
// テキストの形（内容ではなく書式だけを見る）
// ---------------------------------------------------------------------------

/**
 * 「連番として使える裸の数字」か。
 *
 * 数字だけ・符号無し・小数点無し・単位無しで、桁数が上限以下のものだけ。
 * `2024` のような西暦は桁数で、`30%` `2.4x` `-25%` は文字種で落ちる。
 */
export function ordinalToken(raw: string): { value: number; width: number } | undefined {
  const text = raw.trim();
  if (!new RegExp(`^[0-9]{1,${HEURISTIC_THRESHOLDS.maxOrdinalDigits}}$`).test(text)) {
    return undefined;
  }
  return { value: Number(text), width: text.length };
}

/**
 * 「数値が主役の短いテキスト」か（metric 候補）。
 *
 * 数字を含み、短く、数字以外（単位・記号）がごく僅かなもの。
 * "30%" "−25%" "2.4×" "12件" "¥100万" は通り、
 * "前年比30%改善した" のような文章は単位部分が長すぎて落ちる。
 */
export function looksLikeMetric(raw: string): boolean {
  const text = raw.trim();
  if (text.length === 0 || text.length > HEURISTIC_THRESHOLDS.maxNumericTextLength) return false;
  if (!/[0-9]/.test(text)) return false;
  const unit = text.replace(/[0-9\s.,]/g, "").replace(/^[+\-−]/, "");
  return unit.length <= HEURISTIC_THRESHOLDS.maxMetricUnitLength;
}

// ---------------------------------------------------------------------------
// セルを横断した比較（相対幾何・フォント）
// ---------------------------------------------------------------------------

/** セル矩形に対する相対座標・相対サイズ（0..1） */
export function relativeBox(box: Rect, cell: Rect): Rect {
  const w = cell.w > 0 ? cell.w : 1;
  const h = cell.h > 0 ? cell.h : 1;
  return { x: (box.x - cell.x) / w, y: (box.y - cell.y) / h, w: box.w / w, h: box.h / h };
}

/** 相対位置が同じとみなせるか（サイズは見ない） */
export function sameRelativePosition(
  a: Rect,
  b: Rect,
  tolerance = HEURISTIC_THRESHOLDS.relativePositionTolerance,
): boolean {
  return Math.abs(a.x - b.x) <= tolerance && Math.abs(a.y - b.y) <= tolerance;
}

/** 相対位置とサイズの両方が同じとみなせるか */
export function sameRelativeBox(
  a: Rect,
  b: Rect,
  tolerance = HEURISTIC_THRESHOLDS.relativePositionTolerance,
): boolean {
  return (
    sameRelativePosition(a, b, tolerance) &&
    Math.abs(a.w - b.w) <= tolerance &&
    Math.abs(a.h - b.h) <= tolerance
  );
}

const fontSizeOf = (element: ExtractedElement | undefined): number =>
  element?.text?.maxFontSize ?? 12;

/** 同じフォント階層とみなせるか */
export function sameFontTier(
  a: ExtractedElement | undefined,
  b: ExtractedElement | undefined,
  tolerance = HEURISTIC_THRESHOLDS.fontSizeTolerance,
): boolean {
  const [x, y] = [fontSizeOf(a), fontSizeOf(b)];
  const max = Math.max(x, y);
  return max <= 0 || Math.abs(x - y) / max <= tolerance;
}

const centerOf = (box: Rect): { x: number; y: number } => ({
  x: box.x + box.w / 2,
  y: box.y + box.h / 2,
});

const containsCenter = (cell: Rect, box: Rect): boolean => {
  const point = centerOf(box);
  return (
    point.x >= cell.x - 0.02 &&
    point.x <= cell.x + cell.w + 0.02 &&
    point.y >= cell.y - 0.02 &&
    point.y <= cell.y + cell.h + 0.02
  );
};

// ---------------------------------------------------------------------------
// 入出力
// ---------------------------------------------------------------------------

/** 推定対象の反復セル 1 つ */
export interface HeuristicCell {
  /** セルの領域（背景図形があればその矩形、無ければテキストの外接矩形） */
  box: Rect;
  backgroundIndex?: number;
  /** セル内のテキスト要素（上から順） */
  textIndexes: number[];
}

export interface CellContentRoles {
  indexIndex?: number;
  labelIndex: number;
  metricIndex?: number;
  descriptionIndexes: number[];
}

export interface CellStructuralRoles {
  accentIndexes: number[];
  dividerIndex?: number;
}

export interface HeuristicRoleResult {
  /** 全セル分の Content Role。推定できなければ undefined（abstain） */
  content?: CellContentRoles[];
  /** 全セル分の Structural Role（推定できなければ空） */
  structural: CellStructuralRoles[];
  notes: string[];
}

/** 反復セルを横断して役割を推定する。曖昧なら推定しない。 */
export function inferRepeatedCellRoles(
  elements: readonly ExtractedElement[],
  cells: readonly HeuristicCell[],
  structuralCandidates: readonly number[],
): HeuristicRoleResult {
  const notes: string[] = [];
  const empty = (): CellStructuralRoles[] => cells.map(() => ({ accentIndexes: [] }));

  if (cells.length < 2) return { structural: empty(), notes };

  const rows = textRowsOf(elements, cells, notes);
  if (!rows) return { structural: empty(), notes };

  const content = assignContentRoles(elements, cells, rows, notes);
  if (!content) return { structural: empty(), notes };

  const structural = assignStructuralRoles(elements, cells, structuralCandidates, notes);
  return { content, structural, notes };
}

// ---------------------------------------------------------------------------
// テキストを「行」（セル横断の対応関係）へまとめる
// ---------------------------------------------------------------------------

/**
 * 反復セルの同じ相対位置にあるテキストを 1 行として束ねる。
 *
 * `rows[k][c]` = セル c の k 行目の要素インデックス。
 * テキスト数が揃わない・相対位置が揃わない場合は反復として扱えないので undefined。
 */
function textRowsOf(
  elements: readonly ExtractedElement[],
  cells: readonly HeuristicCell[],
  notes: string[],
): number[][] | undefined {
  const counts = new Set(cells.map((cell) => cell.textIndexes.length));
  if (counts.size > 1) {
    notes.push(
      "セルごとのテキスト数が揃わないため、反復セルの役割推定を中止しました（要素は固定デザインとして残します）",
    );
    return undefined;
  }
  const rowCount = cells[0]!.textIndexes.length;
  if (rowCount === 0) return undefined;

  const rows: number[][] = [];
  for (let k = 0; k < rowCount; k += 1) {
    const row = cells.map((cell) => cell.textIndexes[k]!);
    const first = relativeBox(elements[row[0]!]!.box, cells[0]!.box);
    for (let c = 1; c < row.length; c += 1) {
      const other = relativeBox(elements[row[c]!]!.box, cells[c]!.box);
      if (!sameRelativePosition(first, other)) {
        notes.push(
          `セル間でテキストの相対位置が揃わないため、反復セルの役割推定を中止しました（${k + 1} 行目）`,
        );
        return undefined;
      }
      if (!sameFontTier(elements[row[0]!], elements[row[c]!])) {
        notes.push(
          `セル間でテキストのフォントサイズが揃わないため、反復セルの役割推定を中止しました（${k + 1} 行目）`,
        );
        return undefined;
      }
    }
    rows.push(row);
  }
  return rows;
}

const textOf = (elements: readonly ExtractedElement[], index: number): string =>
  elements[index]?.text?.plain.replace(/\s+/g, " ").trim() ?? "";

const rowFontSize = (elements: readonly ExtractedElement[], row: number[]): number =>
  fontSizeOf(elements[row[0]!]);

const rowTop = (elements: readonly ExtractedElement[], row: number[], cells: readonly HeuristicCell[]): number =>
  relativeBox(elements[row[0]!]!.box, cells[0]!.box).y;

// ---------------------------------------------------------------------------
// Content Role: index → metric → label / description の順に決める
// ---------------------------------------------------------------------------

function assignContentRoles(
  elements: readonly ExtractedElement[],
  cells: readonly HeuristicCell[],
  rows: number[][],
  notes: string[],
): CellContentRoles[] | undefined {
  const remaining = [...rows];

  // --- index: 連番であることを「系列」として確認する -------------------------
  const indexRow = detectIndexRow(elements, remaining, notes);
  if (indexRow) remaining.splice(remaining.indexOf(indexRow), 1);

  // --- metric: index を除いた残りから探す ------------------------------------
  const metricRow = detectMetricRow(elements, remaining, notes);
  if (metricRow) remaining.splice(remaining.indexOf(metricRow), 1);

  // --- label / description ---------------------------------------------------
  if (remaining.length === 0) {
    notes.push(
      "label にできるテキストが残らなかったため、反復セルの役割推定を中止しました",
    );
    return undefined;
  }

  const labelRow = remaining[0]!;
  const descriptionRows = remaining.slice(1);

  for (const row of descriptionRows) {
    const labelSize = rowFontSize(elements, labelRow);
    const descriptionSize = rowFontSize(elements, row);
    const ratio = descriptionSize > 0 ? labelSize / descriptionSize : 0;
    if (ratio < HEURISTIC_THRESHOLDS.labelToDescriptionFontRatio) {
      notes.push(
        `label と description のフォント階層が不明確なため（${labelSize}pt と ${descriptionSize}pt）、` +
          "反復セルの役割推定を中止しました",
      );
      return undefined;
    }
    if (rowTop(elements, labelRow, cells) >= rowTop(elements, row, cells)) {
      notes.push(
        "label 候補が description 候補より上に配置されていないため、反復セルの役割推定を中止しました",
      );
      return undefined;
    }
  }

  if (indexRow) {
    notes.push(
      `${cells.map((_, c) => textOf(elements, indexRow[c]!)).join(" / ")} が同じ相対位置に並ぶ連番だったため index と推定しました`,
    );
  }
  if (metricRow) {
    notes.push(
      `${cells.map((_, c) => textOf(elements, metricRow[c]!)).join(" / ")} が同じ相対位置の数値テキストだったため metric と推定しました`,
    );
  }
  notes.push(
    descriptionRows.length > 0
      ? `フォント階層（${rowFontSize(elements, labelRow)}pt > ${rowFontSize(elements, descriptionRows[0]!)}pt）と配置順から label / description を区別しました`
      : "残ったテキストは 1 行だったため label と推定しました",
  );

  return cells.map((_, c) => ({
    ...(indexRow ? { indexIndex: indexRow[c]! } : {}),
    labelIndex: labelRow[c]!,
    ...(metricRow ? { metricIndex: metricRow[c]! } : {}),
    descriptionIndexes: descriptionRows.map((row) => row[c]!),
  }));
}

/**
 * 連番の行を探す。
 *
 * 「数字だけのテキスト」というだけでは採らない。セルの順に並べた値が
 * 1, 2, 3 … になっていること、桁の書式（ゼロ埋めの幅）が揃っていることまで要求する。
 * これで 10% / 20% / 30%、2024 / 2025 / 2026、1 / 3 / 5、01 / 02 / 04 は落ちる。
 */
function detectIndexRow(
  elements: readonly ExtractedElement[],
  rows: readonly number[][],
  notes: string[],
): number[] | undefined {
  const matches: number[][] = [];

  for (const row of rows) {
    const tokens = row.map((index) => ordinalToken(textOf(elements, index)));
    if (tokens.some((token) => token === undefined)) continue;

    const widths = new Set(tokens.map((token) => token!.width));
    if (widths.size > 1) {
      notes.push(
        `連番候補（${row.map((i) => textOf(elements, i)).join(" / ")}）は桁の書式が揃わないため index と推定しませんでした`,
      );
      continue;
    }
    if (tokens.some((token, i) => token!.value !== i + 1)) {
      notes.push(
        `連番候補（${row.map((i) => textOf(elements, i)).join(" / ")}）は 1 から始まる連番ではないため index と推定しませんでした`,
      );
      continue;
    }
    matches.push(row);
  }

  if (matches.length === 0) return undefined;
  if (matches.length > 1) {
    notes.push("連番に見える行が複数あったため index と推定しませんでした");
    return undefined;
  }
  return matches[0]!;
}

/**
 * metric の行を探す。
 *
 * 候補が複数あるときは「明確に一番大きいフォント」の行だけを採る。
 * 決め手が無ければ metric を推定しない（label / description 側へ回す）。
 */
function detectMetricRow(
  elements: readonly ExtractedElement[],
  rows: readonly number[][],
  notes: string[],
): number[] | undefined {
  const matches = rows.filter((row) =>
    row.every((index) => looksLikeMetric(textOf(elements, index))),
  );
  if (matches.length === 0) return undefined;
  if (matches.length === 1) return matches[0]!;

  const sorted = [...matches].sort(
    (a, b) => rowFontSize(elements, b) - rowFontSize(elements, a),
  );
  if (sameFontTier(elements[sorted[0]![0]!], elements[sorted[1]![0]!])) {
    notes.push(
      "数値テキストの行が複数あり、どれを metric とすべきか決め手が無いため metric と推定しませんでした",
    );
    return undefined;
  }
  return sorted[0]!;
}

// ---------------------------------------------------------------------------
// Structural Role: divider → accent
// ---------------------------------------------------------------------------

function assignStructuralRoles(
  elements: readonly ExtractedElement[],
  cells: readonly HeuristicCell[],
  candidates: readonly number[],
  notes: string[],
): CellStructuralRoles[] {
  const dividers = detectDividers(elements, cells, candidates, notes);
  const claimed = new Set(dividers ?? []);
  const accents = detectAccents(
    elements,
    cells,
    candidates.filter((index) => !claimed.has(index)),
    notes,
  );

  return cells.map((_, c) => ({
    accentIndexes: accents ? accents[c]! : [],
    ...(dividers ? { dividerIndex: dividers[c]! } : {}),
  }));
}

type Orientation = "horizontal" | "vertical";

/** 軸に沿った線か（斜めのコネクタ・矢印を弾く） */
function orientationOf(box: Rect): Orientation | undefined {
  const aspect = HEURISTIC_THRESHOLDS.dividerMaxAspectRatio;
  if (box.w > 0 && box.h <= box.w * aspect) return "horizontal";
  if (box.h > 0 && box.w <= box.h * aspect) return "vertical";
  return undefined;
}

/**
 * 区切り線を探す。
 *
 * 「セル内に線がある」だけでは採らない。全セルに 1 本ずつあり、
 * 向き・相対位置・相対長さ・線の色が揃っていることまで要求する。
 */
function detectDividers(
  elements: readonly ExtractedElement[],
  cells: readonly HeuristicCell[],
  candidates: readonly number[],
  notes: string[],
): number[] | undefined {
  const perCell = cells.map((cell) =>
    candidates.filter((index) => {
      const element = elements[index];
      if (!element || !canRenderAsDivider(element)) return false;
      if (!containsCenter(cell.box, element.box)) return false;
      return orientationOf(element.box) !== undefined;
    }),
  );

  if (perCell.every((found) => found.length === 0)) return undefined;
  if (perCell.some((found) => found.length !== 1)) {
    notes.push(
      "セルごとに区切り線の本数が揃わないため divider と推定しませんでした（線は固定デザインとして残します）",
    );
    return undefined;
  }

  const chosen = perCell.map((found) => found[0]!);
  const first = elements[chosen[0]!]!;
  const firstRelative = relativeBox(first.box, cells[0]!.box);
  const firstOrientation = orientationOf(first.box)!;

  for (let c = 1; c < chosen.length; c += 1) {
    const element = elements[chosen[c]!]!;
    const relative = relativeBox(element.box, cells[c]!.box);
    if (orientationOf(element.box) !== firstOrientation) {
      notes.push("セル間で区切り線の向きが揃わないため divider と推定しませんでした");
      return undefined;
    }
    if (!sameRelativePosition(firstRelative, relative)) {
      notes.push("セル間で区切り線の相対位置が揃わないため divider と推定しませんでした");
      return undefined;
    }
    const length = firstOrientation === "horizontal" ? "w" : "h";
    if (
      Math.abs(firstRelative[length] - relative[length]) >
      HEURISTIC_THRESHOLDS.dividerLengthTolerance
    ) {
      notes.push("セル間で区切り線の長さが揃わないため divider と推定しませんでした");
      return undefined;
    }
    if (first.line?.color !== element.line?.color) {
      notes.push("セル間で区切り線の色が揃わないため divider と推定しませんでした");
      return undefined;
    }
  }

  notes.push(
    `各セルの同じ位置にある${firstOrientation === "horizontal" ? "水平" : "垂直"}線を divider と推定しました`,
  );
  return chosen;
}

/** セルの縁に接した細長い帯か（アイコン・バッジ・大きな色面を弾く） */
function isEdgeBand(box: Rect, cell: Rect): boolean {
  const { accentMaxThicknessRatio, accentMinLengthRatio, accentEdgeTolerance } =
    HEURISTIC_THRESHOLDS;
  const relative = relativeBox(box, cell);

  const vertical =
    relative.w <= accentMaxThicknessRatio &&
    relative.h >= accentMinLengthRatio &&
    (Math.abs(relative.x) <= accentEdgeTolerance ||
      Math.abs(relative.x + relative.w - 1) <= accentEdgeTolerance);

  const horizontal =
    relative.h <= accentMaxThicknessRatio &&
    relative.w >= accentMinLengthRatio &&
    (Math.abs(relative.y) <= accentEdgeTolerance ||
      Math.abs(relative.y + relative.h - 1) <= accentEdgeTolerance);

  return vertical || horizontal;
}

const styleKeyOf = (element: ExtractedElement): string =>
  `${element.fill?.type ?? "-"}:${element.fill?.color ?? "-"}:${element.line?.color ?? "-"}`;

/**
 * アクセント（縁の細い帯）を探す。
 *
 * 「色が付いた図形」というだけでは採らない。セルの縁に接した細長い帯で、
 * 全セルに同数あり、相対位置・相対サイズ・塗りが揃っていることまで要求する。
 */
function detectAccents(
  elements: readonly ExtractedElement[],
  cells: readonly HeuristicCell[],
  candidates: readonly number[],
  notes: string[],
): number[][] | undefined {
  const perCell = cells.map((cell) =>
    candidates
      .filter((index) => {
        const element = elements[index];
        if (!element || !canRenderAsAccent(element)) return false;
        if (!containsCenter(cell.box, element.box)) return false;
        return isEdgeBand(element.box, cell.box);
      })
      .sort((a, b) => {
        const [x, y] = [elements[a]!.box, elements[b]!.box];
        return x.y - y.y || x.x - y.x;
      }),
  );

  if (perCell.every((found) => found.length === 0)) return undefined;

  const counts = new Set(perCell.map((found) => found.length));
  if (counts.size > 1 || perCell[0]!.length === 0) {
    notes.push(
      "セルごとにアクセントの数が揃わないため accent と推定しませんでした（図形は固定デザインとして残します）",
    );
    return undefined;
  }

  for (let k = 0; k < perCell[0]!.length; k += 1) {
    const first = elements[perCell[0]![k]!]!;
    const firstRelative = relativeBox(first.box, cells[0]!.box);
    for (let c = 1; c < perCell.length; c += 1) {
      const element = elements[perCell[c]![k]!]!;
      if (!sameRelativeBox(firstRelative, relativeBox(element.box, cells[c]!.box))) {
        notes.push("セル間でアクセントの相対位置・大きさが揃わないため accent と推定しませんでした");
        return undefined;
      }
      if (styleKeyOf(first) !== styleKeyOf(element)) {
        notes.push("セル間でアクセントの塗り・線が揃わないため accent と推定しませんでした");
        return undefined;
      }
    }
  }

  notes.push(
    `各セルの縁にある細い装飾${perCell[0]!.length > 1 ? `（${perCell[0]!.length} 個）` : ""}を accent と推定しました`,
  );
  return perCell;
}
