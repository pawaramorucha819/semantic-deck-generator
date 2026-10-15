import { contentItemEntries } from "../domain/content.js";
import { parsePresentation } from "../domain/schema.js";
import type { PresentationSpec, SemanticContent } from "../domain/types.js";
import type { VisualPlan } from "../plan/visual-plan.js";

export type RelationshipFamily = "enumeration" | "radial" | "flow" | "background" | "convergence" | "cycle";
export interface RelationshipEdge { from: string; to: string; relation: "next" | "returns_to" | "distributes_to" | "contributes_to" | "contains" }
const family: Record<SemanticContent["kind"], RelationshipFamily> = {
  points: "enumeration", sequence: "flow", distribution: "radial",
  context: "background", convergence: "convergence", cycle: "cycle",
};

/** Describe input relations, without guessing edges from wording or position. */
export function relationshipGraph(content: SemanticContent) {
  const entries = contentItemEntries(content);
  const nodes = entries.map(({ item, path }) => ({ ...item, id: item.id ?? path.join("."), inputPath: path.join(".") }));
  const edges: RelationshipEdge[] = [];
  switch (content.kind) {
    case "points": break;
    case "sequence":
    case "cycle":
      content.steps.slice(1).forEach((_, i) => edges.push({ from: `steps.${i}`, to: `steps.${i + 1}`, relation: "next" }));
      if (content.kind === "cycle") edges.push({ from: `steps.${content.steps.length - 1}`, to: "steps.0", relation: "returns_to" });
      break;
    case "distribution":
      content.items.forEach((_, i) => edges.push({ from: "center", to: `items.${i}`, relation: "distributes_to" })); break;
    case "convergence":
      content.inputs.forEach((_, i) => edges.push({ from: `inputs.${i}`, to: "outcome", relation: "contributes_to" })); break;
    case "context":
      content.items.forEach((_, i) => edges.push({ from: "environment", to: `items.${i}`, relation: "contains" })); break;
  }
  const ids = new Map(nodes.map(n => [n.inputPath, n.id]));
  return { family: family[content.kind], nodes, edges: edges.map(e => ({ ...e, from: ids.get(e.from)!, to: ids.get(e.to)! })) };
}

/** Membership tests apply to siblings, not their center/base/outcome. */
function siblings(content?: SemanticContent) {
  if (!content) return [];
  switch (content.kind) {
    case "points": case "distribution": case "context": return content.items;
    case "sequence": case "cycle": return content.steps;
    case "convergence": return content.inputs;
  }
}
function membershipReview(slide: PresentationSpec["slides"][number]) {
  const universe = slide.decomposition?.universe;
  const items = siblings(slide.content);
  if (!universe || !items.length || items.some(item => item.members === undefined)) {
    return { status: "not-verifiable" as const, reason: "有限の対象集合と全ての兄弟要素への所属が揃っていません" };
  }
  const owners = new Map<string, string[]>();
  items.forEach((item, i) => item.members!.forEach(member => owners.set(member, [...(owners.get(member) ?? []), item.id ?? item.label ?? String(i)])));
  const overlap = [...owners.entries()].filter(([, nodes]) => nodes.length > 1).map(([member, nodes]) => ({ member, nodes }));
  const missing = universe.filter(member => !owners.has(member));
  const outsideScope = [...owners.keys()].filter(member => !universe.includes(member));
  return { status: overlap.length || missing.length || outsideScope.length ? "review-required" as const : "declared-sets-pass" as const, overlap, missing, outsideScope,
    limitation: "宣言された集合の分割のみ検査。集合自体の正しさや分解軸の妥当性は未確認です" };
}

/** Mechanical review is evidence for a human review, never a semantic MECE certificate. */
export function buildStructureReview(spec: PresentationSpec, plan?: VisualPlan) {
  const validated = parsePresentation(spec);
  return {
    schemaVersion: "1.0" as const,
    title: validated.title,
    assessment: "requires-human-review" as const,
    limitation: "ラベル重複・参照・方向は機械検査。意味上の重複や範囲の網羅性は、この報告だけでは判定できません。",
    slides: validated.slides.map(slide => {
      const graph = slide.content ? relationshipGraph(slide.content) : undefined;
      const labels = graph?.nodes.map(n => n.label.normalize("NFKC").trim().toLocaleLowerCase("en")) ?? [];
      const duplicateLabels = [...new Set(labels.filter((label, i) => labels.indexOf(label) !== i))];
      const chosen = plan?.slides.find(s => s.slideId === slide.id);
      const questions = [
        ...(slide.decomposition ? [] : ["分解の対象範囲と基準が未記載です"]),
        ...(slide.decomposition?.unresolvedQuestions ?? []),
        "同じ分解基準で、複数の要素に同じ対象が含まれていないか確認してください",
        "対象範囲に対して、要素が抜けていないか確認してください",
        "矢印や包含が、原文・根拠の関係を正しく表しているか確認してください",
      ];
      return {
        slideId: slide.id,
        input: { keyMessage: slide.keyMessage, logicalRelation: slide.logicalRelation, content: slide.content, decomposition: slide.decomposition },
        ...(graph ? { graph } : {}),
        mechanicalChecks: { schemaAndReferences: "passed", duplicateLabels, status: duplicateLabels.length ? "review-required" : "passed" },
        membershipChecks: membershipReview(slide),
        semanticMECE: "unverified" as const,
        reviewQuestions: questions,
        ...(chosen ? { selection: { patternId: chosen.patternId, reason: chosen.reason, fallbackUsed: chosen.fallbackUsed } } : {}),
        warnings: chosen?.fallbackUsed ? ["専用の関係図が選択されていません。関係を省略していないか、描画結果を確認してください"] : [],
      };
    }),
  };
}
