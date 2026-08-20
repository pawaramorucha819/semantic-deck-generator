/**
 * 最小限の引数パーサ。
 * サブコマンドごとに「受け付けるフラグ」を宣言し、未知のフラグは必ず拒否する
 * （typo を黙って無視しない）。
 */

export interface FlagSpec {
  /** 値を取らないフラグ */
  boolean?: readonly string[];
  /** 値を取るフラグ */
  string?: readonly string[];
  /** 複数回指定できるフラグ */
  repeatable?: readonly string[];
}

export interface ParsedArgs {
  positional: string[];
  flags: Record<string, string | boolean | string[] | undefined>;
}

export class CliArgumentError extends Error {
  constructor(message: string) {
    super(message);
    this.name = "CliArgumentError";
  }
}

export function parseArgs(argv: readonly string[], spec: FlagSpec): ParsedArgs {
  const booleans = new Set(spec.boolean ?? []);
  const strings = new Set(spec.string ?? []);
  const repeatables = new Set(spec.repeatable ?? []);
  const positional: string[] = [];
  const flags: Record<string, string | boolean | string[] | undefined> = {};

  for (let i = 0; i < argv.length; i += 1) {
    const token = argv[i]!;
    if (!token.startsWith("--")) {
      if (token.startsWith("-") && token !== "-") {
        if (token === "-h") {
          flags.help = true;
          continue;
        }
        throw new CliArgumentError(`未知のオプションです: ${token}`);
      }
      positional.push(token);
      continue;
    }

    const [rawName, inlineValue] = splitFlag(token.slice(2));
    const name = rawName;

    if (booleans.has(name)) {
      flags[name] = inlineValue === undefined ? true : inlineValue !== "false";
      continue;
    }
    if (strings.has(name) || repeatables.has(name)) {
      const value = inlineValue ?? argv[++i];
      if (value === undefined) throw new CliArgumentError(`--${name} には値が必要です`);
      if (repeatables.has(name)) {
        const current = (flags[name] as string[] | undefined) ?? [];
        current.push(value);
        flags[name] = current;
      } else {
        flags[name] = value;
      }
      continue;
    }
    throw new CliArgumentError(`未知のオプションです: --${name}`);
  }

  return { positional, flags };
}

function splitFlag(token: string): [string, string | undefined] {
  const eq = token.indexOf("=");
  return eq === -1 ? [token, undefined] : [token.slice(0, eq), token.slice(eq + 1)];
}

export function stringFlag(args: ParsedArgs, name: string): string | undefined {
  const value = args.flags[name];
  return typeof value === "string" ? value : undefined;
}

export function boolFlag(args: ParsedArgs, name: string): boolean {
  return args.flags[name] === true;
}

export function listFlag(args: ParsedArgs, name: string): string[] {
  const value = args.flags[name];
  return Array.isArray(value) ? value : [];
}

export function requireString(args: ParsedArgs, name: string): string {
  const value = stringFlag(args, name);
  if (!value) throw new CliArgumentError(`--${name} は必須です`);
  return value;
}

export function numberFlag(args: ParsedArgs, name: string): number | undefined {
  const value = stringFlag(args, name);
  if (value === undefined) return undefined;
  const parsed = Number(value);
  if (!Number.isFinite(parsed)) {
    throw new CliArgumentError(`--${name} には数値を指定してください: ${value}`);
  }
  return parsed;
}
