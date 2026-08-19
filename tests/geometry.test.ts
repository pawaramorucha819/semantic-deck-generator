import { describe, expect, it } from "vitest";
import {
  columns,
  columnsWeighted,
  contains,
  estimateLineCount,
  estimateTextHeight,
  inset,
  intersects,
  LINE_PITCH_FACTOR,
  measureEm,
  splitBottom,
  splitTop,
} from "../src/render/geometry.js";
import { mix, normalizeHex, resolveTheme, tint } from "../src/theme/theme.js";
import { defaultTheme } from "../src/theme/theme.js";

describe("geometry", () => {
  it("columns は gap を含めて等分する", () => {
    const cols = columns({ x: 0, y: 0, w: 10, h: 2 }, 3, 1);
    expect(cols).toHaveLength(3);
    expect(cols[0]!.w).toBeCloseTo(8 / 3);
    expect(cols[2]!.x + cols[2]!.w).toBeCloseTo(10);
  });

  it("columnsWeighted は重み比で分割する", () => {
    const cols = columnsWeighted({ x: 0, y: 0, w: 11, h: 2 }, [1, 1], 1);
    expect(cols[0]!.w).toBeCloseTo(5);
    expect(cols[1]!.x).toBeCloseTo(6);
  });

  it("splitTop / splitBottom は領域を分割する", () => {
    const [top, rest] = splitTop({ x: 0, y: 0, w: 4, h: 4 }, 1, 0.5);
    expect(top.h).toBe(1);
    expect(rest.y).toBe(1.5);
    expect(rest.h).toBe(2.5);

    const [bottomRect, remaining] = splitBottom({ x: 0, y: 0, w: 4, h: 4 }, 1, 0.5);
    expect(bottomRect.y).toBe(3);
    expect(remaining.h).toBe(2.5);
  });

  it("inset は四辺を縮める", () => {
    expect(inset({ x: 1, y: 1, w: 4, h: 4 }, 0.5)).toEqual({ x: 1.5, y: 1.5, w: 3, h: 3 });
  });

  it("intersects / contains が判定できる", () => {
    expect(intersects({ x: 0, y: 0, w: 2, h: 2 }, { x: 1, y: 1, w: 2, h: 2 })).toBe(true);
    expect(intersects({ x: 0, y: 0, w: 2, h: 2 }, { x: 3, y: 0, w: 2, h: 2 })).toBe(false);
    expect(contains({ x: 0, y: 0, w: 5, h: 5 }, { x: 1, y: 1, w: 2, h: 2 })).toBe(true);
    expect(contains({ x: 0, y: 0, w: 5, h: 5 }, { x: 4, y: 4, w: 2, h: 2 })).toBe(false);
  });

  it("全角と半角で文字幅を区別する", () => {
    expect(measureEm("あいう")).toBeCloseTo(3);
    expect(measureEm("abc")).toBeCloseTo(1.56);
    expect(measureEm("あa")).toBeGreaterThan(measureEm("aa"));
  });

  it("行数とテキスト高さを概算できる", () => {
    // widthIn=2in / 12pt → 名目 12em/行。安全係数 1.08 で実効約 11.1em/行。
    const opts = { widthIn: 2, fontSize: 12 };
    expect(estimateLineCount("あ".repeat(11), opts)).toBe(1);
    expect(estimateLineCount("あ".repeat(12), opts)).toBe(2); // 12*1.08 = 12.96em > 12em
    expect(estimateLineCount("あ".repeat(23), opts)).toBe(3); // 名目では 2 行だが安全係数で 3 行
    expect(estimateLineCount("a\nb", opts)).toBe(2);
    expect(estimateTextHeight("あ".repeat(11), opts)).toBeCloseTo(
      (12 * 1.32 * LINE_PITCH_FACTOR) / 72,
      3,
    );
  });
});

describe("theme", () => {
  it("hex を正規化する", () => {
    expect(normalizeHex("#0b4f8a")).toBe("0B4F8A");
    expect(normalizeHex("abc")).toBe("AABBCC");
    expect(() => normalizeHex("xyz123")).toThrow();
  });

  it("mix / tint は 2 色を補間する", () => {
    expect(mix("000000", "FFFFFF", 0.5)).toBe("808080");
    expect(tint("000000", 1)).toBe("FFFFFF");
    expect(tint("0B4F8A", 0)).toBe("0B4F8A");
  });

  it("resolveTheme が省略項目を補完する", () => {
    const resolved = resolveTheme({
      ...defaultTheme,
      fonts: { heading: "Arial", body: "Arial" },
      colors: {
        background: "FFFFFF",
        text: "000000",
        mutedText: "888888",
        primary: "0B4F8A",
        secondary: "2E8FD0",
        accent: "E2703A",
      },
      fontSizes: { keyMessage: 22, heading: 15, body: 12, caption: 9 },
    });

    expect(resolved.fonts.numeric).toBe("Arial");
    expect(resolved.fontSizes.display).toBe(44);
    expect(resolved.colors.surface).toBeDefined();
    expect(resolved.colors.positive).toBe("2E8FD0");
    expect(resolved.margins.left).toBeGreaterThan(0);
    expect(resolved.shape.cornerRadius).toBeGreaterThan(0);
  });
});
