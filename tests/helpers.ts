import { readFile } from "node:fs/promises";
import { fileURLToPath } from "node:url";
import { dirname, join } from "node:path";
import JSZip from "jszip";
import { parsePresentation } from "../src/domain/schema.js";
import type { PresentationSpec, SlideSpec } from "../src/domain/types.js";

const here = dirname(fileURLToPath(import.meta.url));
export const projectRoot = join(here, "..");
export const examplePath = join(projectRoot, "examples", "inventory-system.json");

export async function loadExample(): Promise<PresentationSpec> {
  return parsePresentation(JSON.parse(await readFile(examplePath, "utf8")));
}

export async function loadExampleJson(): Promise<Record<string, unknown>> {
  return JSON.parse(await readFile(examplePath, "utf8")) as Record<string, unknown>;
}

/** テスト用の最小 SlideSpec */
export function slide(partial: Partial<SlideSpec> = {}): SlideSpec {
  return {
    id: "slide-test",
    keyMessage: "テスト用のキーメッセージ",
    slideRole: "analysis",
    logicalRelation: "claim_evidence",
    evidence: [],
    ...partial,
  };
}

export function textEvidence(id: string, label: string, statement = `${label}に関する事実`) {
  return {
    id,
    type: "text" as const,
    statement,
    data: { label },
  };
}

export function metricEvidence(id: string, label: string, value: number, unit = "%") {
  return {
    id,
    type: "metric" as const,
    statement: `${label}は${value}${unit}である`,
    data: { label, value, unit },
  };
}

export interface PptxContents {
  zip: JSZip;
  slideXmls: string[];
  presentationXml: string;
  /** ディレクトリエントリを除いた実ファイルのみ */
  files: string[];
}

export async function readPptx(buffer: Buffer): Promise<PptxContents> {
  const zip = await JSZip.loadAsync(buffer);
  const files = Object.entries(zip.files)
    .filter(([, entry]) => !entry.dir)
    .map(([name]) => name)
    .sort();
  const slideNames = files
    .filter((f) => /^ppt\/slides\/slide\d+\.xml$/.test(f))
    .sort((a, b) => slideIndex(a) - slideIndex(b));
  const slideXmls = await Promise.all(
    slideNames.map((name) => zip.file(name)!.async("string")),
  );
  const presentationXml = await zip.file("ppt/presentation.xml")!.async("string");
  return { zip, slideXmls, presentationXml, files };
}

function slideIndex(name: string): number {
  return Number(/slide(\d+)\.xml$/.exec(name)?.[1] ?? 0);
}

/** スライド XML から <a:t> のテキストをすべて抜き出す */
export function extractTexts(slideXml: string): string[] {
  return [...slideXml.matchAll(/<a:t>([\s\S]*?)<\/a:t>/g)].map((m) =>
    (m[1] ?? "")
      .replace(/&amp;/g, "&")
      .replace(/&lt;/g, "<")
      .replace(/&gt;/g, ">")
      .replace(/&quot;/g, '"')
      .replace(/&apos;/g, "'"),
  );
}

export function countMatches(xml: string, pattern: RegExp): number {
  return [...xml.matchAll(pattern)].length;
}
