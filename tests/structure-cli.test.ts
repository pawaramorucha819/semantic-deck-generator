import { mkdtemp, readFile, writeFile, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { afterEach, describe, expect, it, vi } from "vitest";
import { runStructure } from "../src/cli/commands/structure.js";
import { BodyDecompositionAdvisor } from "../src/ai/body-decomposition-advisor.js";
import { nullAIProvider } from "../src/ai/provider.js";
import { replayBodyDecomposition } from "../src/api.js";
import {
  sourceFingerprint, resolveExactQuote, buildDecompositionReview,
} from "../src/structure/index.js";

const raw = {
  schemaVersion: "1.0" as const, id: "workflow", title: "確認工程",
  body: "収集して検証し、結果を共有する。",
};
function candidate() {
  const nodes = [{ id: "collect", label: "収集" }, { id: "check", label: "検証" }, { id: "share", label: "結果を共有" }];
  return {
    schemaVersion: "1.0", id: "workflow-candidate", status: "draft",
    sourceHash: sourceFingerprint(raw), author: { kind: "fixture", id: "offline-test" },
    slide: {
      id: "workflow", keyMessage: "収集・検証の後に結果を共有する", slideRole: "recommendation",
      logicalRelation: "process", content: { kind: "sequence", steps: nodes },
      decomposition: { scope: "確認業務", criterion: "実施順序", rationale: "原文の順番に沿う", unresolvedQuestions: ["担当者は未確認"] },
      evidence: [],
    },
    nodeSources: nodes.map(node => ({ nodeId: node.id, spans: [resolveExactQuote(raw, node.label)], origin: "source", rationale: "原文の工程" })),
    messageSources: [resolveExactQuote(raw, raw.body)],
    relation: { rationale: "原文の実施順序を保持", spans: [resolveExactQuote(raw, raw.body)] },
    excludedSpans: [],
  };
}

const temporary: string[] = [];
afterEach(async () => {
  vi.restoreAllMocks();
  await Promise.all(temporary.splice(0).map(path => rm(path, { recursive: true, force: true })));
});
async function workspace() {
  const dir = await mkdtemp(join(tmpdir(), "sdg-structure-test-")); temporary.push(dir); return dir;
}

describe("offline structure workflow", () => {
  it("freezes plaintext without normalizing CRLF or emoji and never invents a draft", async () => {
    const dir = await workspace(); const source = join(dir, "body.txt"); const out = join(dir, "request.json");
    await writeFile(source, "第一行\r\n🙂第二行\n");
    await runStructure(["prepare", source, "--title", "原文", "--out", out]);
    const frozen = JSON.parse(await readFile(join(dir, "request.raw.json"), "utf8"));
    expect(frozen.body).toBe("第一行\r\n🙂第二行\n");
    const request = JSON.parse(await readFile(out, "utf8"));
    expect(request.input.sourceHash).toBe(sourceFingerprint(frozen));
    expect(request).not.toHaveProperty("draft");
    expect(request.instructions).toContain("untrusted source material");
  });

  it("replays, reviews and compiles with explicit fixture identity and no external calls", async () => {
    const network = vi.spyOn(globalThis, "fetch").mockRejectedValue(new Error("network forbidden"));
    const dir = await workspace(); const source = join(dir, "raw.json"); const draftPath = join(dir, "candidate.json");
    const replayed = join(dir, "replayed.json"); const packetPath = join(dir, "packet.json");
    await writeFile(source, JSON.stringify(raw)); await writeFile(draftPath, JSON.stringify(candidate()));
    await runStructure(["replay", source, "--draft", draftPath, "--out", replayed]);
    await runStructure(["review", source, "--draft", replayed, "--out", packetPath]);
    const template = JSON.parse(await readFile(join(dir, "packet.review-template.json"), "utf8"));
    expect(template.decision).toBe("request-changes"); expect(template.actor.id).toBe("");
    await expect(runStructure(["compile", source, "--draft", replayed, "--review", join(dir, "packet.review-template.json"), "--out", join(dir, "blocked.json")])).rejects.toThrow();
    const packet = buildDecompositionReview(raw, candidate());
    const approval = {
      schemaVersion: "1.0", sourceHash: packet.sourceHash, draftHash: packet.draftHash,
      decision: "approve", actor: { kind: "fixture", id: "test-reviewer" }, rationale: "Synthetic regression fixture, not user approval",
      dispositions: packet.warnings.map(w => ({ id: w.id, decision: "accepted-limitation", reason: "Fixture retains this uncertainty without certification" })),
    };
    const reviewPath = join(dir, "review.json"); await writeFile(reviewPath, JSON.stringify(approval));
    const presentation = join(dir, "presentation.json");
    await runStructure(["compile", source, "--draft", replayed, "--review", reviewPath, "--out", presentation]);
    const spec = JSON.parse(await readFile(presentation, "utf8"));
    expect(spec.slides[0].content.steps.map((n: { id: string }) => n.id)).toEqual(["collect", "check", "share"]);
    expect(spec).not.toHaveProperty("approval");
    expect(await readFile(join(dir, "presentation.lineage.json"), "utf8")).toContain('"semanticMECE": "unverified"');
    expect(network).not.toHaveBeenCalled();
  });

  it("protects operator edits instead of overwriting output", async () => {
    const dir = await workspace(); const source = join(dir, "source.json"); const out = join(dir, "request.json");
    await writeFile(source, JSON.stringify(raw)); await writeFile(out, "operator edit");
    await expect(runStructure(["prepare", source, "--out", out])).rejects.toThrow("Refusing to overwrite");
    expect(await readFile(out, "utf8")).toBe("operator edit");
  });

  it("rejects numeric evidence before compilation writes any output, without stripping it", async () => {
    const dir = await workspace(); const source = join(dir, "raw.json"); const draftPath = join(dir, "draft.json");
    const out = join(dir, "presentation.json");
    const numeric = { ...candidate(), slide: { ...candidate().slide, evidence: [{ id: "metric", type: "metric", statement: "処理時間は5分", data: { value: 5, unit: "分" } }] } };
    await writeFile(source, JSON.stringify(raw)); await writeFile(draftPath, JSON.stringify(numeric));
    await expect(runStructure(["compile", source, "--draft", draftPath, "--review", join(dir, "not-read.json"), "--out", out])).rejects.toThrow(/evidence/);
    await expect(readFile(out, "utf8")).rejects.toMatchObject({ code: "ENOENT" });
    expect(JSON.parse(await readFile(draftPath, "utf8")).slide.evidence).toHaveLength(1);
  });

  it("rechecks provider output and rejects a provider-supplied approval field", async () => {
    const unsafe = { ...candidate(), approval: { decision: "approve" } };
    await expect(replayBodyDecomposition(raw, unsafe)).rejects.toThrow();
    const advisor = new BodyDecompositionAdvisor({ id: "unsafe-test", available: true, generateStructured: async () => unsafe as never });
    await expect(advisor.advise(raw)).rejects.toThrow();
    await expect(new BodyDecompositionAdvisor(nullAIProvider).advise(raw)).rejects.toThrow("explicitly supplied provider");
  });
});
