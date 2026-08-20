import type { z } from "zod";

/**
 * AI Provider — LLM への唯一の入口。
 *
 * 設計の原則:
 *  1. **ベンダー SDK を Domain / Pattern / Renderer へ持ち込まない**。
 *     具体 Provider は Adapter として差し替える。
 *  2. **出力は必ず構造化 + 検証**。自由テキストをそのまま信用しない。
 *  3. **AI は optional**。Provider が無くても生成パイプラインは完全に動く。
 */

export interface AIRequest<T> {
  /** 用途識別子（'pattern-selection' など）。ログ・テストの分岐に使う。 */
  purpose: string;
  /** モデルへの指示（システムプロンプト相当） */
  instructions: string;
  /** 構造化された入力。JSON 化して渡される。 */
  input: unknown;
  /** 出力の検証スキーマ。これを通らない応答は採用されない。 */
  schema: z.ZodType<T>;
  /** 制約付きデコードに対応した Provider 向けの JSON Schema（任意） */
  jsonSchema?: Record<string, unknown>;
  maxOutputTokens?: number;
}

export interface AIProvider {
  readonly id: string;
  /** false の場合、この Provider は呼び出せない（API キー未設定など） */
  readonly available: boolean;
  generateStructured<T>(request: AIRequest<T>): Promise<T>;
}

export class AIUnavailableError extends Error {
  constructor(message: string) {
    super(message);
    this.name = "AIUnavailableError";
  }
}

export class AIResponseValidationError extends Error {
  constructor(
    message: string,
    readonly raw?: unknown,
  ) {
    super(message);
    this.name = "AIResponseValidationError";
  }
}

/** AI を使わないときの既定 Provider。呼ばれたら必ず明示的に失敗する。 */
export const nullAIProvider: AIProvider = {
  id: "none",
  available: false,
  generateStructured(): never {
    throw new AIUnavailableError(
      "AI Provider が設定されていません（--ai-mode off で実行するか、provider を設定してください）",
    );
  },
};

/** 応答（未検証）を schema へ通す共通処理 */
export function validateAIResponse<T>(request: AIRequest<T>, raw: unknown): T {
  const result = request.schema.safeParse(raw);
  if (!result.success) {
    throw new AIResponseValidationError(
      `AI 応答が期待するスキーマに適合しません (${request.purpose}):\n${result.error.issues
        .map((i) => `  - ${i.path.join(".") || "(root)"}: ${i.message}`)
        .join("\n")}`,
      raw,
    );
  }
  return result.data;
}

/** JSON テキストから最初の JSON オブジェクトを取り出す（前後の説明文を許容する） */
export function extractJsonObject(text: string): unknown {
  const trimmed = text.trim();
  const fenced = /```(?:json)?\s*([\s\S]*?)```/.exec(trimmed);
  const body = fenced?.[1]?.trim() ?? trimmed;
  const start = body.indexOf("{");
  const end = body.lastIndexOf("}");
  if (start === -1 || end === -1 || end <= start) {
    throw new AIResponseValidationError("AI 応答から JSON を取り出せませんでした", text);
  }
  try {
    return JSON.parse(body.slice(start, end + 1));
  } catch (error) {
    throw new AIResponseValidationError(
      `AI 応答の JSON を解析できませんでした: ${error instanceof Error ? error.message : String(error)}`,
      text,
    );
  }
}

export type ScriptedResponder = (input: unknown, request: AIRequest<unknown>) => unknown;

/** 応答は「値そのもの」か「入力から値を作る関数」 */
export type ScriptedResponse =
  | ScriptedResponder
  | Record<string, unknown>
  | unknown[]
  | string
  | number
  | boolean
  | null;

export interface ScriptedAIProviderOptions {
  id?: string;
  /** purpose → 応答（値そのもの、または入力から作る関数） */
  responses: Record<string, ScriptedResponse>;
  /** 未定義の purpose を呼ばれたときに例外にするか（既定 true） */
  strict?: boolean;
}

/**
 * 決定論的なテスト・デモ用 Provider。
 *
 * 外部 API を一切呼ばないので、CI でも API キー無しの環境でも
 * 「AI を使ったときの挙動」を再現できる。
 * 出力は本物の Provider とまったく同じ検証を通る。
 */
export class ScriptedAIProvider implements AIProvider {
  readonly id: string;
  readonly available = true;
  /** 呼び出し履歴（テストで「AI に何を渡したか」を検証するため） */
  readonly calls: Array<{ purpose: string; input: unknown }> = [];

  constructor(private readonly options: ScriptedAIProviderOptions) {
    this.id = options.id ?? "scripted";
  }

  // eslint-disable-next-line @typescript-eslint/require-await
  async generateStructured<T>(request: AIRequest<T>): Promise<T> {
    this.calls.push({ purpose: request.purpose, input: request.input });
    const responder = this.options.responses[request.purpose];
    if (responder === undefined) {
      if (this.options.strict === false) {
        throw new AIUnavailableError(`purpose '${request.purpose}' の応答が定義されていません`);
      }
      throw new AIUnavailableError(
        `ScriptedAIProvider に purpose '${request.purpose}' の応答が定義されていません`,
      );
    }
    const raw =
      typeof responder === "function"
        ? (responder as ScriptedResponder)(request.input, request as AIRequest<unknown>)
        : responder;
    // スクリプト応答でも本番と同じ検証を通す（テストが検証を素通りしないように）
    return validateAIResponse(request, raw);
  }
}
