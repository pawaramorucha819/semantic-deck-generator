import type { AssetResolver } from "../render/context.js";
import { corporateTemplate } from "./builtin/corporate.js";
import { defaultTemplate } from "./builtin/default.js";
import { createAssetResolver, loadTemplatesFromRoot, type LoadedTemplate } from "./loader.js";
import type { LoadedTemplateAsset, PresentationTemplate } from "./template.js";

/**
 * Template レジストリ。
 *
 * PatternRegistry と同じ考え方で、ビルトインと企業提供（外部ディレクトリ）の
 * Template を同じ入口から引けるようにする。
 * 同じ id が登録された場合は「後から登録した外部 Template が勝つ」＝
 * 企業がビルトインを上書きできる、という運用を明示的に許可する。
 */

export type TemplateSource = "builtin" | "external";

export interface TemplateRegistryEntry {
  template: PresentationTemplate;
  source: TemplateSource;
  assets: LoadedTemplateAsset[];
  /** 外部 Template の読み込み元 */
  dir?: string;
}

export const builtinTemplates: PresentationTemplate[] = [defaultTemplate, corporateTemplate];

export class TemplateRegistry {
  private readonly byId = new Map<string, TemplateRegistryEntry>();

  constructor(templates: PresentationTemplate[] = builtinTemplates) {
    for (const template of templates) {
      this.registerBuiltin(template);
    }
  }

  registerBuiltin(template: PresentationTemplate): this {
    return this.register({ template, source: "builtin", assets: [] });
  }

  register(entry: TemplateRegistryEntry): this {
    this.byId.set(entry.template.id, entry);
    return this;
  }

  registerLoaded(loaded: LoadedTemplate): this {
    return this.register({
      template: loaded.template,
      source: "external",
      assets: loaded.assets,
      ...(loaded.dir ? { dir: loaded.dir } : {}),
    });
  }

  /** ディレクトリ配下の Template をすべて登録する */
  async registerDirectory(rootDir: string): Promise<LoadedTemplate[]> {
    const loaded = await loadTemplatesFromRoot(rootDir);
    for (const item of loaded) this.registerLoaded(item);
    return loaded;
  }

  has(id: string): boolean {
    return this.byId.has(id);
  }

  get(id: string): TemplateRegistryEntry {
    const entry = this.byId.get(id);
    if (!entry) {
      throw new Error(
        `未登録の templateId です: ${id} (登録済み: ${this.ids().join(", ") || "なし"})`,
      );
    }
    return entry;
  }

  template(id: string): PresentationTemplate {
    return this.get(id).template;
  }

  ids(): string[] {
    return [...this.byId.keys()];
  }

  all(): TemplateRegistryEntry[] {
    return [...this.byId.values()];
  }

  /** 一覧表示用の要約 */
  list(): Array<{
    id: string;
    name: string;
    version: string;
    source: TemplateSource;
    organization?: string;
    masters: number;
    dir?: string;
  }> {
    return this.all().map((entry) => ({
      id: entry.template.id,
      name: entry.template.name,
      version: entry.template.version,
      source: entry.source,
      ...(entry.template.metadata?.organization
        ? { organization: entry.template.metadata.organization }
        : {}),
      masters: entry.template.masters.length,
      ...(entry.dir ? { dir: entry.dir } : {}),
    }));
  }

  assetResolverFor(id: string): AssetResolver {
    return createAssetResolver(this.get(id).assets);
  }
}

export const defaultTemplateRegistry = new TemplateRegistry();
