import { spawnSync } from "node:child_process";
import { existsSync, mkdtempSync, mkdirSync, rmSync } from "node:fs";
import { basename, extname, join, resolve } from "node:path";
import { tmpdir } from "node:os";
import { pathToFileURL } from "node:url";

// Optional local review tool. Generation itself does not require LibreOffice/Poppler.
const input = process.argv[2];
if (!input || !existsSync(input) || extname(input).toLowerCase() !== ".pptx") {
  console.error("Usage: npm run preview -- <existing.pptx> [output-directory]");
  process.exit(1);
}
const out = resolve(process.argv[3] ?? ".preview");
mkdirSync(out, { recursive: true });
const profile = mkdtempSync(join(tmpdir(), "sdg-preview-"));
const env = { ...process.env, HOME: profile, XDG_CACHE_HOME: join(profile, "cache") };
function run(command, args) {
  const result = spawnSync(command, args, { encoding: "utf8", env, timeout: 120000 });
  if (result.error || result.status !== 0) {
    throw new Error(`${command} failed: ${result.error?.message ?? result.stderr}. Install LibreOffice and Poppler to create previews.`);
  }
  if (result.stdout) console.log(result.stdout.trim());
}
try {
  run("soffice", [`-env:UserInstallation=${pathToFileURL(join(profile, "office")).href}`, "--headless", "--convert-to", "pdf", "--outdir", out, resolve(input)]);
  const stem = basename(input, extname(input));
  const pdf = join(out, `${stem}.pdf`);
  if (!existsSync(pdf)) throw new Error("LibreOffice returned without producing a PDF.");
  run("pdftoppm", ["-scale-to", "1600", "-png", pdf, join(out, `${stem}-slide`)]);
  console.log(`Preview: ${pdf}\nInspect every ${stem}-slide-*.png. A clean render report is not visual approval.`);
} catch (error) {
  console.error(error instanceof Error ? error.message : error);
  process.exitCode = 1;
} finally {
  rmSync(profile, { recursive: true, force: true });
}
