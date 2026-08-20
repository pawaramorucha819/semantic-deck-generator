import {
  extractJsonObject,
  validateAIResponse,
  AIUnavailableError,
  type AIProvider,
  type AIRequest,
} from "./provider.js";

/**
 * Anthropic Claude 用の Adapter。
 *
 * 公式 SDK (`@anthropic-ai/sdk`) を **動的 import** で読み込む。
 *  - コア機能（生成パイプライン）は AI を使わないので、SDK を必須依存にしない
 *  - 未インストール / API キー未設定なら、その場で明確なエラーを出す
 *    （「AI を使ったふりをして黙って fallback する」ことはしない）
 *
 * 他ベンダーへ移るときは、このファイルと同じ形の Adapter を足すだけでよい。
 * Domain / Pattern / Renderer からこのファイルを import してはいけない。
 */

/** SDK のうち、この Adapter が使う部分だけを構造的に表した型 */
interface MessagesClient {
  messages: {
    create(params: Record<string, unknown>): Promise<{
      content: Array<{ type: string; text?: string }>;
      stop_reason?: string | null;
    }>;
  };
}

export interface AnthropicProviderOptions {
  /** 省略時は環境変数 ANTHROPIC_API_KEY */
  apiKey?: string;
  /** 既定 'claude-opus-5' */
  model?: string;
  /** 推論の深さ。小さな構造化タスクなので既定は 'medium'。 */
  effort?: "low" | "medium" | "high" | "xhigh" | "max";
  maxOutputTokens?: number;
  /** テスト用に SDK クライアントを差し替える口 */
  client?: MessagesClient;
}

export const DEFAULT_ANTHROPIC_MODEL = "claude-opus-5";

export class AnthropicAIProvider implements AIProvider {
  readonly id: string;
  private client: MessagesClient | undefined;

  constructor(private readonly options: AnthropicProviderOptions = {}) {
    this.id = `anthropic:${options.model ?? DEFAULT_ANTHROPIC_MODEL}`;
    this.client = options.client;
  }

  get available(): boolean {
    return !!this.client || !!(this.options.apiKey ?? process.env.ANTHROPIC_API_KEY);
  }

  private async ensureClient(): Promise<MessagesClient> {
    if (this.client) return this.client;

    const apiKey = this.options.apiKey ?? process.env.ANTHROPIC_API_KEY;
    if (!apiKey) {
      throw new AIUnavailableError(
        "ANTHROPIC_API_KEY が設定されていません（--ai-mode off で AI を使わずに生成できます）",
      );
    }

    let module: { default: new (options: { apiKey: string }) => MessagesClient };
    try {
      module = (await import("@anthropic-ai/sdk")) as unknown as typeof module;
    } catch (error) {
      throw new AIUnavailableError(
        `@anthropic-ai/sdk が見つかりません。AI 機能を使う場合は 'npm install @anthropic-ai/sdk' を実行してください\n${
          error instanceof Error ? error.message : String(error)
        }`,
      );
    }

    this.client = new module.default({ apiKey });
    return this.client;
  }

  async generateStructured<T>(request: AIRequest<T>): Promise<T> {
    const client = await this.ensureClient();

    const params: Record<string, unknown> = {
      model: this.options.model ?? DEFAULT_ANTHROPIC_MODEL,
      max_tokens: request.maxOutputTokens ?? this.options.maxOutputTokens ?? 4096,
      system: `${request.instructions}\n\n必ず JSON オブジェクトだけを出力してください。説明文やコードフェンスは不要です。`,
      messages: [{ role: "user", content: JSON.stringify(request.input, null, 2) }],
      output_config: {
        effort: this.options.effort ?? "medium",
        ...(request.jsonSchema
          ? { format: { type: "json_schema", schema: request.jsonSchema } }
          : {}),
      },
    };

    const response = await client.messages.create(params);
    const text = response.content
      .filter((block) => block.type === "text")
      .map((block) => block.text ?? "")
      .join("");

    if (text.trim().length === 0) {
      throw new AIUnavailableError(
        `Anthropic API がテキストを返しませんでした (stop_reason=${response.stop_reason ?? "unknown"})`,
      );
    }

    // Provider が何を返しても、採用前に必ずスキーマ検証を通す
    return validateAIResponse(request, extractJsonObject(text));
  }
}

export function createAnthropicProvider(options: AnthropicProviderOptions = {}): AIProvider {
  return new AnthropicAIProvider(options);
}
