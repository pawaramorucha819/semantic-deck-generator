import { mkdir, readFile, writeFile } from "node:fs/promises";
import { dirname } from "node:path";
import { buildArgumentReviewPacket } from "../../review/argument.js";
import { parseArgs, requireString, boolFlag, CliArgumentError } from "../args.js";
export const ARGUMENT_USAGE = "semantic-deck argument-review <presentation.json> --out <packet.json>\nWrites an unapproved, revision-bound review packet. No semantic approval is inferred.";
export async function runArgumentReview(argv: readonly string[]): Promise<number> {
  const args = parseArgs(argv, { boolean: ["help"], string: ["out"] });
  if (boolFlag(args, "help")) { console.log(ARGUMENT_USAGE); return 0; }
  if (args.positional.length !== 1) throw new CliArgumentError(ARGUMENT_USAGE);
  const packet = buildArgumentReviewPacket(JSON.parse(await readFile(args.positional[0]!, "utf8")));
  if (!packet.slides.length) throw new CliArgumentError("No explicit argument in input");
  const out = requireString(args, "out"), body = `${JSON.stringify(packet,null,2)}\n`;
  await mkdir(dirname(out), { recursive: true });
  try { await writeFile(out, body, { flag: "wx" }); }
  catch (error) {
    if ((error as NodeJS.ErrnoException).code !== "EEXIST" || await readFile(out,"utf8") !== body) throw error;
  }
  console.log(`Written unapproved argument review packet: ${out}`);
  return 0;
}
