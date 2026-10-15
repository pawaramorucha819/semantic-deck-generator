import { readFile, writeFile, mkdir } from "node:fs/promises";
import { basename, dirname, extname } from "node:path";
import { prepareBodyDecompositionRequest, replayBodyDecomposition } from "../../api.js";
import { createSafeDeckId } from "../../knowledge/deck-id.js";
import {
  parseRawBody, validateDecompositionDraft, buildDecompositionReview, compileDecomposition,
} from "../../structure/index.js";
import { boolFlag, parseArgs, requireString, stringFlag, CliArgumentError } from "../args.js";

export const STRUCTURE_USAGE = `
semantic-deck structure <prepare|import|replay|review|compile> <raw.json> [options]

  prepare <raw.json|body.txt> --out <request.json> [--title <title>]
    Freeze source and export authoring instructions. Does not infer a decomposition.
  import <raw.json> --draft <candidate.json> --out <draft.json>
    Validate an assistant/human-authored candidate without API calls.
  replay <raw.json> --draft <candidate.json> --out <draft.json>
    Replay a saved response through the provider validation boundary, offline.
  review <raw.json> --draft <draft.json> --out <review-packet.json>
    Write source/node/coverage comparison and a request-changes review template.
  compile <raw.json> --draft <draft.json> --review <review.json> --out <presentation.json>
    Require a separate explicit, revision-bound operator review; write lineage too.

Raw JSON contains schemaVersion, id, title, body and optional source metadata.
prepare also writes <request>.raw.json; use that frozen raw file in later commands.
No command calls an external AI service. Review actors are self-declared, not authenticated.
Output files are not overwritten with different content. Choose a new path for revisions.
Semantic MECE remains unverified even after approval. See docs/body-decomposition.md.
`.trim();

async function readJson(path: string): Promise<unknown> {
  return JSON.parse(await readFile(path, "utf8"));
}

function siblingPath(path: string, suffix: string) {
  return `${path.slice(0, path.length - extname(path).length)}.${suffix}.json`;
}

/** Idempotent replay is safe; differing files may contain operator edits. */
async function writeJson(path: string, value: unknown) {
  const body = `${JSON.stringify(value, null, 2)}\n`;
  await mkdir(dirname(path), { recursive: true });
  try {
    await writeFile(path, body, { flag: "wx" });
  } catch (error) {
    if ((error as NodeJS.ErrnoException).code !== "EEXIST") throw error;
    if (await readFile(path, "utf8") !== body) {
      throw new CliArgumentError(`Refusing to overwrite changed output: ${path}. Choose a new path.`);
    }
  }
  console.log(`Written: ${path}`);
}

export async function runStructure(argv: readonly string[]): Promise<number> {
  const args = parseArgs(argv, {
    boolean: ["help"], string: ["out", "draft", "review", "title"],
  });
  if (boolFlag(args, "help") || args.positional.length === 0) {
    console.log(STRUCTURE_USAGE);
    return 0;
  }
  const [command, sourcePath, ...extra] = args.positional;
  if (!sourcePath || extra.length || !["prepare", "import", "replay", "review", "compile"].includes(command!)) {
    throw new CliArgumentError(STRUCTURE_USAGE);
  }
  const out = requireString(args, "out");
  const raw = command === "prepare" && extname(sourcePath).toLowerCase() !== ".json"
    ? parseRawBody({
      schemaVersion: "1.0",
      id: createSafeDeckId(basename(sourcePath, extname(sourcePath))),
      title: stringFlag(args, "title") ?? basename(sourcePath, extname(sourcePath)),
      body: await readFile(sourcePath, "utf8"),
    })
    : parseRawBody(await readJson(sourcePath));
  if (command === "prepare") {
    await writeJson(out, prepareBodyDecompositionRequest(raw));
    await writeJson(siblingPath(out, "raw"), raw);
    console.log("Authoring is pending. No decomposition or approval was inferred.");
    return 0;
  }
  const inputDraft = await readJson(requireString(args, "draft"));
  const draft = command === "replay"
    ? await replayBodyDecomposition(raw, inputDraft)
    : validateDecompositionDraft(raw, inputDraft);
  if (command === "import" || command === "replay") {
    await writeJson(out, draft);
    console.log("Validated draft only; explicit review is still required.");
  } else if (command === "review") {
    const packet = buildDecompositionReview(raw, draft);
    await writeJson(out, packet);
    await writeJson(siblingPath(out, "review-template"), {
      schemaVersion: "1.0", sourceHash: packet.sourceHash, draftHash: packet.draftHash,
      decision: "request-changes",
      actor: { kind: "human", id: "" }, rationale: "",
      dispositions: packet.warnings.map(warning => ({
        id: warning.id, decision: "accepted-limitation", reason: "",
      })),
    });
    console.log("Template is deliberately unapproved/incomplete. Record the actual reviewer and reasons.");
  } else {
    const result = compileDecomposition(raw, draft, await readJson(requireString(args, "review")));
    await writeJson(out, result.presentation);
    await writeJson(siblingPath(out, "lineage"), result.lineage);
    console.log("Compiled reviewed structure; semantic MECE and source entailment are not certified.");
  }
  return 0;
}
