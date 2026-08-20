import { readFile, readdir, stat } from "node:fs/promises";
import { extname, isAbsolute, join, relative, resolve, sep } from "node:path";
import type { AssetResolver, ResolvedAsset } from "../render/context.js";
import { parseTemplate } from "./schema.js";
import type { LoadedTemplateAsset, PresentationTemplate } from "./template.js";

/**
 * Template パッケージ（ディレクトリ）のロード。
 *
 *   templates/<id>/
 *     template.json   … Template 定義（必須）
 *     theme.json      … theme を別ファイルに分ける場合（任意）
 *     assets/logo.png … ロゴ等（任意）
 *
 * セキュリティ上の前提: `.pptx` から取り込んだものを含め、外部 Template は
 * 「信頼できない入力」として扱う。ここではコードを一切実行せず、
 * アセットのパスがテンプレートディレクトリの外を指していないことを検証する。
 */

/** 1 アセットあたりの上限（DoS・巨大ファイル混入対策） */
export const MAX_ASSET_BYTES = 4 * 1024 * 1024;

const MIME_BY_EXTENSION: Record<string, string> = {
  ".png": "image/png",
  ".jpg": "image/jpeg",
  ".jpeg": "image/jpeg",
  ".gif": "image/gif",
  ".svg": "image/svg+xml",
};

export interface LoadedTemplate {
  template: PresentationTemplate;
  assets: LoadedTemplateAsset[];
  /** 読み込み元ディレクトリ（builtin の場合は undefined） */
  dir?: string;
}

export class TemplateLoadError extends Error {
  constructor(message: string) {
    super(message);
    this.name = "TemplateLoadError";
  }
}

/** dir 配下に収まっている絶対パスだけを返す（パストラバーサル対策） */
export function resolveInsideDirectory(dir: string, relativePath: string): string {
  if (isAbsolute(relativePath)) {
    throw new TemplateLoadError(`絶対パスのアセットは参照できません: ${relativePath}`);
  }
  const base = resolve(dir);
  const target = resolve(base, relativePath);
  const rel = relative(base, target);
  if (rel.startsWith("..") || isAbsolute(rel) || rel.split(sep).includes("..")) {
    throw new TemplateLoadError(
      `テンプレートディレクトリの外を参照しています: ${relativePath}`,
    );
  }
  return target;
}

export function toDataUri(bytes: Buffer, fileName: string): string {
  const mime = MIME_BY_EXTENSION[extname(fileName).toLowerCase()];
  if (!mime) {
    throw new TemplateLoadError(
      `対応していない画像形式です: ${fileName} (対応: ${Object.keys(MIME_BY_EXTENSION).join(", ")})`,
    );
  }
  return `${mime};base64,${bytes.toString("base64")}`;
}

async function readJson(path: string): Promise<unknown> {
  const raw = await readFile(path, "utf8");
  try {
    return JSON.parse(raw);
  } catch (error) {
    throw new TemplateLoadError(
      `JSON として読み込めませんでした: ${path}\n${error instanceof Error ? error.message : String(error)}`,
    );
  }
}

/** 1 つの Template ディレクトリを読み込む */
export async function loadTemplateFromDirectory(dir: string): Promise<LoadedTemplate> {
  const absoluteDir = resolve(dir);
  const templatePath = join(absoluteDir, "template.json");

  let raw: unknown;
  try {
    raw = await readJson(templatePath);
  } catch (error) {
    if ((error as NodeJS.ErrnoException).code === "ENOENT") {
      throw new TemplateLoadError(`template.json が見つかりません: ${templatePath}`);
    }
    throw error;
  }

  // theme を別ファイルに切り出している場合はここで合成する
  if (raw && typeof raw === "object" && !("theme" in (raw as Record<string, unknown>))) {
    const themePath = join(absoluteDir, "theme.json");
    try {
      const theme = await readJson(themePath);
      (raw as Record<string, unknown>).theme = theme;
    } catch (error) {
      if ((error as NodeJS.ErrnoException).code !== "ENOENT") throw error;
    }
  }

  const template = parseTemplate(raw);
  const assets: LoadedTemplateAsset[] = [];

  for (const asset of template.assets ?? []) {
    const path = resolveInsideDirectory(absoluteDir, asset.path);
    const info = await stat(path).catch(() => undefined);
    if (!info?.isFile()) {
      throw new TemplateLoadError(
        `Template '${template.id}' のアセットが見つかりません: ${asset.path}`,
      );
    }
    if (info.size > MAX_ASSET_BYTES) {
      throw new TemplateLoadError(
        `アセットが大きすぎます (${asset.path}: ${info.size} bytes > ${MAX_ASSET_BYTES} bytes)`,
      );
    }
    const bytes = await readFile(path);
    assets.push({ ...asset, data: toDataUri(bytes, asset.path) });
  }

  return { template, assets, dir: absoluteDir };
}

/** ルートディレクトリ配下の Template ディレクトリをすべて読み込む */
export async function loadTemplatesFromRoot(rootDir: string): Promise<LoadedTemplate[]> {
  const absoluteRoot = resolve(rootDir);
  const info = await stat(absoluteRoot).catch(() => undefined);
  if (!info?.isDirectory()) {
    throw new TemplateLoadError(`Template ディレクトリが見つかりません: ${absoluteRoot}`);
  }

  // ルート自体が 1 つの Template パッケージであるケースも許容する
  const selfTemplate = await stat(join(absoluteRoot, "template.json")).catch(() => undefined);
  if (selfTemplate?.isFile()) {
    return [await loadTemplateFromDirectory(absoluteRoot)];
  }

  const entries = await readdir(absoluteRoot, { withFileTypes: true });
  const loaded: LoadedTemplate[] = [];
  for (const entry of entries.filter((e) => e.isDirectory()).sort((a, b) => a.name.localeCompare(b.name))) {
    const dir = join(absoluteRoot, entry.name);
    const has = await stat(join(dir, "template.json")).catch(() => undefined);
    if (has?.isFile()) loaded.push(await loadTemplateFromDirectory(dir));
  }
  return loaded;
}

/** assetId → data URI の解決器を作る */
export function createAssetResolver(assets: readonly LoadedTemplateAsset[]): AssetResolver {
  const byId = new Map<string, ResolvedAsset>(
    assets.map((a) => [a.id, { id: a.id, data: a.data, fileName: a.path }]),
  );
  return {
    has: (id) => byId.has(id),
    resolve(id) {
      const asset = byId.get(id);
      if (!asset) {
        throw new Error(
          `未登録の画像アセットです: ${id} (登録済み: ${[...byId.keys()].join(", ") || "なし"})`,
        );
      }
      return asset;
    },
  };
}
