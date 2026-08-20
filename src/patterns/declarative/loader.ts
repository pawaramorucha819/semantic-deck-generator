import { readFile, readdir, stat } from "node:fs/promises";
import { join, resolve } from "node:path";
import {
  MAX_ASSET_BYTES,
  resolveInsideDirectory,
  TemplateLoadError,
  toDataUri,
} from "../../templates/loader.js";
import type { LoadedTemplateAsset } from "../../templates/template.js";
import {
  compileDeclarativePattern,
  patternAssetId,
  type CompiledDeclarativePattern,
} from "./compile.js";
import type { DeclarativePatternDocument, PatternScope } from "./schema.js";

/**
 * Pattern パッケージ（ディレクトリ）のロード。
 *
 *   patterns/company/three-column/
 *     pattern.json   … Declarative Pattern 定義（必須）
 *     preview.png    … 人が中身を確認するためのプレビュー（任意）
 *     assets/…       … Pattern が使う画像（任意）
 *
 * 外部 Pattern も「信頼できない入力」として扱う。コードは実行せず、
 * アセットのパスがパッケージ外を指していないことを検証する。
 */

export class PatternPackageLoadError extends Error {
  constructor(message: string) {
    super(message);
    this.name = "PatternPackageLoadError";
  }
}

export interface LoadedPatternPackage {
  pattern: CompiledDeclarativePattern;
  document: DeclarativePatternDocument;
  dir: string;
  /** ID は `pattern:<patternId>:<assetId>` へ名前空間化されている */
  assets: LoadedTemplateAsset[];
  /** プレビュー画像の絶対パス（存在する場合） */
  previewImagePath?: string;
  scope: PatternScope;
}

export interface LoadPatternPackageOptions {
  /** draft（人間レビュー前）も読み込むか。既定 false。 */
  allowDraft?: boolean;
  /** metadata.scope が無い場合に使う scope。既定 'company'。 */
  defaultScope?: PatternScope;
}

export async function loadPatternPackage(
  dir: string,
  options: LoadPatternPackageOptions = {},
): Promise<LoadedPatternPackage> {
  const absoluteDir = resolve(dir);
  const patternPath = join(absoluteDir, "pattern.json");

  let raw: unknown;
  try {
    raw = JSON.parse(await readFile(patternPath, "utf8"));
  } catch (error) {
    if ((error as NodeJS.ErrnoException).code === "ENOENT") {
      throw new PatternPackageLoadError(`pattern.json が見つかりません: ${patternPath}`);
    }
    throw new PatternPackageLoadError(
      `pattern.json を読み込めませんでした: ${patternPath}\n${error instanceof Error ? error.message : String(error)}`,
    );
  }

  const pattern = compileDeclarativePattern(raw, {
    ...(options.allowDraft !== undefined ? { allowDraft: options.allowDraft } : {}),
  });
  const doc = pattern.document;

  const assets: LoadedTemplateAsset[] = [];
  for (const asset of doc.assets ?? []) {
    const path = resolveInsideDirectory(absoluteDir, asset.path);
    const info = await stat(path).catch(() => undefined);
    if (!info?.isFile()) {
      throw new PatternPackageLoadError(
        `Pattern '${doc.id}' のアセットが見つかりません: ${asset.path}`,
      );
    }
    if (info.size > MAX_ASSET_BYTES) {
      throw new PatternPackageLoadError(
        `アセットが大きすぎます (${asset.path}: ${info.size} bytes > ${MAX_ASSET_BYTES} bytes)`,
      );
    }
    assets.push({
      id: patternAssetId(doc.id, asset.id),
      path: asset.path,
      kind: "image",
      data: toDataUri(await readFile(path), asset.path),
    });
  }

  let previewImagePath: string | undefined;
  if (doc.metadata?.previewImage) {
    const candidate = resolveInsideDirectory(absoluteDir, doc.metadata.previewImage);
    const info = await stat(candidate).catch(() => undefined);
    if (!info?.isFile()) {
      throw new PatternPackageLoadError(
        `Pattern '${doc.id}' のプレビュー画像が見つかりません: ${doc.metadata.previewImage}`,
      );
    }
    previewImagePath = candidate;
  }

  return {
    pattern,
    document: doc,
    dir: absoluteDir,
    assets,
    ...(previewImagePath ? { previewImagePath } : {}),
    scope: doc.metadata?.scope ?? options.defaultScope ?? "company",
  };
}

/** ルートディレクトリ配下の Pattern パッケージをすべて読み込む */
export async function loadPatternPackages(
  rootDir: string,
  options: LoadPatternPackageOptions = {},
): Promise<LoadedPatternPackage[]> {
  const absoluteRoot = resolve(rootDir);
  const info = await stat(absoluteRoot).catch(() => undefined);
  if (!info?.isDirectory()) {
    throw new PatternPackageLoadError(`Pattern ディレクトリが見つかりません: ${absoluteRoot}`);
  }

  const self = await stat(join(absoluteRoot, "pattern.json")).catch(() => undefined);
  if (self?.isFile()) return [await loadPatternPackage(absoluteRoot, options)];

  const entries = await readdir(absoluteRoot, { withFileTypes: true });
  const loaded: LoadedPatternPackage[] = [];
  for (const entry of entries
    .filter((e) => e.isDirectory())
    .sort((a, b) => a.name.localeCompare(b.name))) {
    const dir = join(absoluteRoot, entry.name);
    const has = await stat(join(dir, "pattern.json")).catch(() => undefined);
    if (has?.isFile()) loaded.push(await loadPatternPackage(dir, options));
  }
  return loaded;
}

export { TemplateLoadError as AssetPathError };
