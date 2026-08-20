import { readFile } from "node:fs/promises";
import { dirname, isAbsolute, resolve } from "node:path";
import { createAnthropicProvider } from "../ai/anthropic.js";
import { PatternAdvisor } from "../ai/pattern-advisor.js";
import {
  ScriptedAIProvider,
  type AIProvider,
  type ScriptedResponse,
} from "../ai/provider.js";
import {
  DeterministicRetriever,
  type KnowledgeRetriever,
} from "../knowledge/retrieval/retrieval.js";
import { JsonKnowledgeStore } from "../knowledge/store/json-store.js";
import {
  CompositeKnowledgeStore,
  type DeckKnowledgeStore,
} from "../knowledge/store/store.js";
import type { PatternScope } from "../patterns/declarative/schema.js";
import { PatternRegistry, builtinPatterns } from "../patterns/registry.js";
import type { GenerateOptions } from "../pipeline.js";
import type { SelectionMode } from "../plan/decision-trace.js";
import type { SelectionPolicyOptions } from "../selector/selection-policy.js";
import { TemplateRegistry } from "../templates/registry.js";
import {
  GenerationConfigError,
  parseGenerationConfig,
  type GenerationConfig,
} from "./schema.js";

/**
 * Generation Config → 実行コンテキスト。
 *
 * 「設定を読む」と「実際に何を組み立てるか」をここに閉じ込め、
 * CLI もライブラリ利用者も同じ経路を通るようにする。
 */

export interface GenerationContext {
  config: GenerationConfig;
  /** config ファイルのあるディレクトリ（相対パスの基準） */
  baseDir: string;
  templateRegistry: TemplateRegistry;
  templateId?: string;
  patternRegistry: PatternRegistry;
  knowledgeStore?: DeckKnowledgeStore;
  /** 設定順に並んだ個別ストア（recordKnowledge の追記先を選ぶのに使う） */
  knowledgeStores: JsonKnowledgeStore[];
  retriever?: KnowledgeRetriever;
  aiProvider?: AIProvider;
  advisor?: PatternAdvisor;
  policy?: SelectionPolicyOptions;
  /** generate() へそのまま渡せるオプション */
  generateOptions: GenerateOptions;
  outputDir?: string;
  writeRenderReport: boolean;
  writeAdvisorReport: boolean;
  /** ロード中に起きた注意事項（CLI が表示する） */
  notes: string[];
}

export interface LoadGenerationConfigResult {
  config: GenerationConfig;
  baseDir: string;
  path: string;
}

export async function loadGenerationConfig(path: string): Promise<LoadGenerationConfigResult> {
  const absolute = resolve(path);
  let raw: unknown;
  try {
    raw = JSON.parse(await readFile(absolute, "utf8"));
  } catch (error) {
    if ((error as NodeJS.ErrnoException).code === "ENOENT") {
      throw new GenerationConfigError(`generation.config.json が見つかりません: ${absolute}`);
    }
    throw new GenerationConfigError(
      `generation.config.json を読み込めませんでした: ${absolute}\n${
        error instanceof Error ? error.message : String(error)
      }`,
    );
  }
  return { config: parseGenerationConfig(raw), baseDir: dirname(absolute), path: absolute };
}

function resolvePath(baseDir: string, path: string): string {
  return isAbsolute(path) ? path : resolve(baseDir, path);
}

export interface BuildContextOptions {
  /** CLI からの上書き */
  templateId?: string;
  themeId?: string;
  aiMode?: SelectionMode | "off";
  outputDir?: string;
  templateDirs?: string[];
  patternDirs?: string[];
  knowledgeDirs?: string[];
  writeRenderReport?: boolean;
  writeAdvisorReport?: boolean;
  includeTitleSlide?: boolean;
  includeCandidates?: boolean;
  sectionSlideIds?: string[];
}

async function createProvider(
  config: GenerationConfig,
  baseDir: string,
): Promise<AIProvider | undefined> {
  const provider = config.ai?.provider;
  if (!provider || provider.type === "none") return undefined;

  if (provider.type === "anthropic") {
    const apiKey = process.env[provider.apiKeyEnv ?? "ANTHROPIC_API_KEY"];
    return createAnthropicProvider({
      ...(apiKey ? { apiKey } : {}),
      ...(provider.model ? { model: provider.model } : {}),
      ...(provider.effort ? { effort: provider.effort } : {}),
    });
  }

  // scripted: オフラインでも AI 経路を通せるようにする（デモ・CI 用）
  const path = resolvePath(baseDir, provider.responses);
  let responses: Record<string, ScriptedResponse>;
  try {
    responses = JSON.parse(await readFile(path, "utf8")) as Record<string, ScriptedResponse>;
  } catch (error) {
    throw new GenerationConfigError(
      `scripted provider の応答ファイルを読み込めませんでした: ${path}\n${
        error instanceof Error ? error.message : String(error)
      }`,
    );
  }
  return new ScriptedAIProvider({ id: `scripted:${provider.responses}`, responses });
}

/** Generation Config から Registry / Store / Provider / Policy を組み立てる */
export async function buildGenerationContext(
  loaded: LoadGenerationConfigResult,
  overrides: BuildContextOptions = {},
): Promise<GenerationContext> {
  const { config, baseDir } = loaded;
  const notes: string[] = [];

  // --- Template ------------------------------------------------------------
  const templateRegistry = new TemplateRegistry();
  const templateDirs = [
    ...(config.template?.dirs ?? []).map((dir) => resolvePath(baseDir, dir)),
    ...(overrides.templateDirs ?? []).map((dir) => resolve(dir)),
  ];
  for (const dir of templateDirs) {
    const loadedTemplates = await templateRegistry.registerDirectory(dir);
    notes.push(
      `Template 読込: ${dir} → ${loadedTemplates.map((t) => t.template.id).join(", ") || "(なし)"}`,
    );
  }
  const templateId = overrides.templateId ?? config.template?.id;
  if (templateId && !templateRegistry.has(templateId)) {
    throw new GenerationConfigError(
      `template.id '${templateId}' が見つかりません (登録済み: ${templateRegistry.ids().join(", ")})`,
    );
  }

  // --- Pattern -------------------------------------------------------------
  const includeBuiltin = config.patterns?.builtin ?? true;
  const patternRegistry = new PatternRegistry(includeBuiltin ? builtinPatterns : []);
  if (!includeBuiltin) {
    notes.push("ビルトイン Pattern は無効化されています（patterns.builtin=false）");
  }
  const patternSources: Array<{ dir: string; scope?: PatternScope }> = [
    ...(config.patterns?.sources ?? []).map((source) => ({
      dir: resolvePath(baseDir, source.dir),
      ...(source.scope ? { scope: source.scope } : {}),
    })),
    ...(overrides.patternDirs ?? []).map((dir) => ({ dir: resolve(dir) })),
  ];
  for (const source of patternSources) {
    const packages = await patternRegistry.registerDirectory(source.dir, {
      ...(config.patterns?.allowDraft !== undefined
        ? { allowDraft: config.patterns.allowDraft }
        : {}),
      ...(source.scope ? { defaultScope: source.scope } : {}),
    });
    notes.push(
      `Pattern 読込: ${source.dir} → ${packages.map((p) => p.document.id).join(", ") || "(なし)"}`,
    );
  }

  // --- Knowledge -----------------------------------------------------------
  const knowledgeDirs = [
    ...(config.knowledge?.stores ?? []).map((store) => resolvePath(baseDir, store.dir)),
    ...(overrides.knowledgeDirs ?? []).map((dir) => resolve(dir)),
  ];
  const stores = knowledgeDirs.map((dir) => new JsonKnowledgeStore(dir));
  const knowledgeStore =
    stores.length === 0
      ? undefined
      : stores.length === 1
        ? stores[0]!
        : new CompositeKnowledgeStore(stores);
  const retriever = knowledgeStore
    ? new DeterministicRetriever(() => knowledgeStore.all())
    : undefined;

  // --- AI ------------------------------------------------------------------
  const configuredMode = (overrides.aiMode ?? config.ai?.mode ?? "off") as SelectionMode | "off";
  const usePatternAdvisor = config.ai?.patternSelection ?? true;
  const aiProvider = await createProvider(config, baseDir);

  if ((configuredMode === "assist" || configuredMode === "auto") && usePatternAdvisor) {
    if (!aiProvider) {
      throw new GenerationConfigError(
        `--ai-mode ${configuredMode} には AI Provider の設定が必要です（generation.config.json の ai.provider を設定してください）`,
      );
    }
    if (!aiProvider.available) {
      throw new GenerationConfigError(
        `AI Provider '${aiProvider.id}' が利用できません（API キーが設定されていない可能性があります）。--ai-mode off で AI 無しの生成ができます`,
      );
    }
  }
  const advisor = aiProvider && usePatternAdvisor ? new PatternAdvisor(aiProvider) : undefined;

  // --- Selection Policy ----------------------------------------------------
  const overridesMap = config.selection?.overrides;
  const needsPolicy =
    configuredMode !== "off" || !!overridesMap || !!knowledgeStore;
  const policyMode: SelectionMode = configuredMode === "off" ? "manual" : configuredMode;

  const policy: SelectionPolicyOptions | undefined = needsPolicy
    ? {
        mode: policyMode,
        ...(advisor && configuredMode !== "off" ? { advisor } : {}),
        ...(retriever ? { retriever } : {}),
        ...(overridesMap ? { overrides: overridesMap } : {}),
        ...(config.knowledge?.limit !== undefined ? { knowledgeLimit: config.knowledge.limit } : {}),
        ...(config.knowledge?.context ? { knowledgeContext: config.knowledge.context } : {}),
      }
    : undefined;

  if (configuredMode === "off" && needsPolicy) {
    notes.push(
      "AI は使用しません（Knowledge の参照情報と人間の指定のみを Decision Trace に記録します）",
    );
  }

  // --- generate() 用オプション ---------------------------------------------
  const sectionSlideIds = overrides.sectionSlideIds ?? config.selection?.sectionSlides;
  const themeId = overrides.themeId ?? config.theme?.id;

  const generateOptions: GenerateOptions = {
    registry: patternRegistry,
    templateRegistry,
    ...(templateId ? { templateId } : {}),
    ...(themeId ? { themeId } : {}),
    ...(sectionSlideIds ? { sectionSlideIds } : {}),
    ...(policy ? { policy } : {}),
    ...(knowledgeStore ? { knowledgeStoreId: knowledgeStore.id } : {}),
    ...(config.output?.titleSlide !== undefined
      ? { includeTitleSlide: config.output.titleSlide }
      : {}),
    ...(overrides.includeTitleSlide !== undefined
      ? { includeTitleSlide: overrides.includeTitleSlide }
      : {}),
    ...(config.output?.slideNumbers !== undefined
      ? { showSlideNumbers: config.output.slideNumbers }
      : {}),
    ...(config.output?.candidates !== undefined
      ? { includeCandidates: config.output.candidates }
      : {}),
    ...(overrides.includeCandidates !== undefined
      ? { includeCandidates: overrides.includeCandidates }
      : {}),
    ...(config.selection?.minimumScore !== undefined
      ? { minimumScore: config.selection.minimumScore }
      : {}),
  };

  return {
    config,
    baseDir,
    templateRegistry,
    ...(templateId ? { templateId } : {}),
    patternRegistry,
    ...(knowledgeStore ? { knowledgeStore } : {}),
    knowledgeStores: stores,
    ...(retriever ? { retriever } : {}),
    ...(aiProvider ? { aiProvider } : {}),
    ...(advisor ? { advisor } : {}),
    ...(policy ? { policy } : {}),
    generateOptions,
    ...(overrides.outputDir ?? config.output?.dir
      ? { outputDir: resolvePath(baseDir, overrides.outputDir ?? config.output!.dir!) }
      : {}),
    writeRenderReport: overrides.writeRenderReport ?? config.output?.renderReport ?? false,
    writeAdvisorReport: overrides.writeAdvisorReport ?? config.output?.advisorReport ?? false,
    notes,
  };
}
