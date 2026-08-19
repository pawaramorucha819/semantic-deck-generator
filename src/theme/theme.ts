/**
 * Theme — 「見た目」だけを担当する層。
 *
 * Pattern は *構造* を、Theme は *見た目* を担当します。
 * Pattern 実装の中に色・フォント名・サイズをハードコードしないでください。
 */

export interface PresentationTheme {
  id: string;
  name: string;

  fonts: {
    heading: string;
    body: string;
    /** 数値・単位に使う書体（省略時 heading にフォールバック） */
    numeric?: string;
  };

  /** すべて pt */
  fontSizes: {
    keyMessage: number;
    heading: number;
    body: number;
    caption: number;
    /** BigNumber 等の巨大数値用（省略時 keyMessage * 2） */
    display?: number;
  };

  /** すべて inch（PptxGenJS の座標系と同じ単位） */
  spacing: {
    xs: number;
    sm: number;
    md: number;
    lg: number;
    xl: number;
  };

  colors: {
    background: string;
    text: string;
    mutedText: string;
    primary: string;
    secondary: string;
    accent: string;
    /** カード等の面色。省略時 primary の淡色から導出。 */
    surface?: string;
    /** 罫線色。省略時 mutedText の淡色から導出。 */
    border?: string;
    /** 良い変化。省略時 secondary。 */
    positive?: string;
    /** 悪い変化。省略時 accent。 */
    negative?: string;
    /** 濃色面の上に載せる文字色。省略時 background。 */
    onPrimary?: string;
  };

  /** スライド外周の余白（inch） */
  margins?: {
    top: number;
    right: number;
    bottom: number;
    left: number;
  };

  shape?: {
    /** roundRect の角丸（0.0 - 1.0） */
    cornerRadius: number;
    /** 罫線の太さ（pt） */
    borderWidth: number;
  };
}

/** 省略可能項目を埋めた、Renderer / Pattern が実際に使う Theme */
export interface ResolvedTheme extends PresentationTheme {
  fonts: Required<PresentationTheme["fonts"]>;
  fontSizes: Required<PresentationTheme["fontSizes"]>;
  colors: Required<PresentationTheme["colors"]>;
  margins: NonNullable<PresentationTheme["margins"]>;
  shape: NonNullable<PresentationTheme["shape"]>;
}

/** '#RRGGBB' / 'RRGGBB' を pptxgenjs が要求する 'RRGGBB' に正規化する */
export function normalizeHex(hex: string): string {
  const raw = hex.trim().replace(/^#/, "").toUpperCase();
  if (/^[0-9A-F]{3}$/.test(raw)) {
    return raw
      .split("")
      .map((c) => c + c)
      .join("");
  }
  if (!/^[0-9A-F]{6}$/.test(raw)) {
    throw new Error(`不正なカラーコードです: ${hex}`);
  }
  return raw;
}

/** 2 色を線形補間する。amount = 0 で from、1 で to。 */
export function mix(from: string, to: string, amount: number): string {
  const a = normalizeHex(from);
  const b = normalizeHex(to);
  const t = Math.min(1, Math.max(0, amount));
  const channel = (i: number) => {
    const start = parseInt(a.slice(i * 2, i * 2 + 2), 16);
    const end = parseInt(b.slice(i * 2, i * 2 + 2), 16);
    return Math.round(start + (end - start) * t)
      .toString(16)
      .padStart(2, "0")
      .toUpperCase();
  };
  return `${channel(0)}${channel(1)}${channel(2)}`;
}

/** hex を白（amount=1）方向へ混色する。0 = そのまま、1 = 白。 */
export function tint(hex: string, amount: number): string {
  return mix(hex, "FFFFFF", amount);
}

/** hex を黒（amount=1）方向へ混色する。 */
export function shade(hex: string, amount: number): string {
  return mix(hex, "000000", amount);
}

export function resolveTheme(theme: PresentationTheme): ResolvedTheme {
  const colors = theme.colors;
  return {
    ...theme,
    fonts: {
      ...theme.fonts,
      numeric: theme.fonts.numeric ?? theme.fonts.heading,
    },
    fontSizes: {
      ...theme.fontSizes,
      display: theme.fontSizes.display ?? Math.round(theme.fontSizes.keyMessage * 2),
    },
    colors: {
      ...colors,
      surface: colors.surface ?? tint(colors.primary, 0.92),
      border: colors.border ?? tint(colors.mutedText, 0.7),
      positive: colors.positive ?? colors.secondary,
      negative: colors.negative ?? colors.accent,
      onPrimary: colors.onPrimary ?? colors.background,
    },
    margins: theme.margins ?? { top: 0.42, right: 0.6, bottom: 0.4, left: 0.6 },
    shape: theme.shape ?? { cornerRadius: 0.06, borderWidth: 1 },
  };
}

/** 標準テーマ（コンサル資料寄りの落ち着いた配色） */
export const defaultTheme: PresentationTheme = {
  id: "default",
  name: "Default (Corporate Blue)",
  fonts: {
    heading: "Yu Gothic",
    body: "Yu Gothic",
    numeric: "Segoe UI",
  },
  fontSizes: {
    keyMessage: 22,
    heading: 15,
    body: 13,
    caption: 10,
    display: 48,
  },
  spacing: {
    xs: 0.06,
    sm: 0.12,
    md: 0.22,
    lg: 0.36,
    xl: 0.56,
  },
  colors: {
    background: "FFFFFF",
    text: "1A2230",
    mutedText: "6B7686",
    primary: "0B4F8A",
    secondary: "2E8FD0",
    accent: "E2703A",
    surface: "F2F6FA",
    border: "D5DCE4",
    positive: "1F8A70",
    negative: "C0503C",
    onPrimary: "FFFFFF",
  },
};

/** 差し替え例（テーマ交換が Pattern を触らずに成立することの実証） */
export const midnightTheme: PresentationTheme = {
  id: "midnight",
  name: "Midnight",
  fonts: {
    heading: "Meiryo",
    body: "Meiryo",
    numeric: "Segoe UI",
  },
  fontSizes: {
    keyMessage: 22,
    heading: 15,
    body: 13,
    caption: 10,
    display: 46,
  },
  spacing: { xs: 0.06, sm: 0.12, md: 0.22, lg: 0.36, xl: 0.56 },
  colors: {
    background: "10151F",
    text: "F2F5F9",
    mutedText: "94A2B5",
    primary: "5AA9F0",
    secondary: "3E7FBF",
    accent: "F0A05A",
    surface: "1B2433",
    border: "2C3849",
    positive: "4FC3A1",
    negative: "F07A6A",
    onPrimary: "10151F",
  },
};

export const builtinThemes: Record<string, PresentationTheme> = {
  [defaultTheme.id]: defaultTheme,
  [midnightTheme.id]: midnightTheme,
};

export function getTheme(id: string | undefined): PresentationTheme {
  if (!id) return defaultTheme;
  const theme = builtinThemes[id];
  if (!theme) {
    throw new Error(
      `未知のテーマです: ${id} (利用可能: ${Object.keys(builtinThemes).join(", ")})`,
    );
  }
  return theme;
}
