#!/usr/bin/env node
import { runGenerate } from "./commands/generate.js";
import { reportError } from "./main.js";

/**
 * 従来の入口（`npm run generate -- presentation.json`）。
 *
 * V2 のサブコマンド CLI（src/cli/main.ts）と同じ実装を呼ぶだけの薄いラッパで、
 * V1 からの利用方法をそのまま維持するために残している。
 */
runGenerate(process.argv.slice(2))
  .then((code) => {
    process.exitCode = code;
  })
  .catch((error: unknown) => {
    reportError(error);
    process.exitCode = 1;
  });
