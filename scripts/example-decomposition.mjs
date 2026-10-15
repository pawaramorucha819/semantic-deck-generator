import { spawnSync } from "node:child_process";
import { fileURLToPath } from "node:url";
import { dirname, resolve } from "node:path";

const root = resolve(dirname(fileURLToPath(import.meta.url)), "..");
const base = "examples/decomposition";
const raw = `${base}/nec-raw.json`;
const draft = `${base}/nec-candidate.json`;
const generated = `${base}/generated`;
// The recorded assistant review is an input fixture, never generated or approved here.
const commands = [
  ["structure", "prepare", raw, "--out", `${generated}/sample-request.json`],
  ["structure", "replay", raw, "--draft", draft, "--out", `${generated}/sample-draft.json`],
  ["structure", "review", raw, "--draft", `${generated}/sample-draft.json`, "--out", `${generated}/sample-review-packet.json`],
  ["structure", "compile", raw, "--draft", `${generated}/sample-draft.json`, "--review", `${base}/nec-assistant-review.json`, "--out", `${generated}/presentation.json`],
  ["generate", `${generated}/presentation.json`, "--no-title-slide", "--ai-mode", "off", "--report", "--strict", "--theme", "portable", "--out-dir", generated],
];
for (const args of commands) {
  const result = spawnSync(process.execPath, ["--import", "tsx", "src/cli/main.ts", ...args], { cwd: root, stdio: "inherit" });
  if (result.error || result.status !== 0) {
    console.error(result.error?.message ?? `Example stopped at ${args[0]} ${args[1]}; review inputs and errors before continuing.`);
    process.exit(result.status ?? 1);
  }
}
