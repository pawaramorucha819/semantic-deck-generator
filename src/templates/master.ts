import type PptxGenJS from "pptxgenjs";
import type { AssetResolver } from "../render/context.js";
import { defaultFontRoleFor, resolveColor, resolveFontFace } from "../render/tokens.js";
import type { ResolvedTheme } from "../theme/theme.js";
import {
  expandTemplateText,
  type PresentationTemplate,
  type SlideMasterDefinition,
  type TemplateTextStyle,
  type TemplateTextVariables,
} from "./template.js";

/**
 * Template の Master 定義 → PptxGenJS の Slide Master。
 *
 * ここが「企業の共通装飾」の唯一の描画点。Pattern はここを一切知らないし、
 * 逆にここは Pattern の中身を知らない。
 *
 * PowerPoint の Slide Master として書き出すため、ロゴ・フッター・機密表示は
 * スライド本体ではなくマスター側に入る（= PowerPoint 上で一括編集できる）。
 */

export interface DefineMastersOptions {
  pptx: PptxGenJS;
  template: PresentationTemplate;
  theme: ResolvedTheme;
  assets: AssetResolver;
  variables: TemplateTextVariables;
  /** false ならページ番号フィールドを Master へ入れない（CLI の --no-slide-numbers 相当） */
  includeSlideNumbers?: boolean;
}

/** master.id → PptxGenJS の master 名 */
export type MasterNameMap = ReadonlyMap<string, string>;

export function masterNameOf(template: PresentationTemplate, master: SlideMasterDefinition): string {
  return `${template.id}__${master.id}`;
}

function textOptions(
  theme: ResolvedTheme,
  box: { x: number; y: number; w: number; h: number },
  style: TemplateTextStyle | undefined,
): PptxGenJS.TextPropsOptions {
  const role = style?.role ?? "caption";
  return {
    x: box.x,
    y: box.y,
    w: box.w,
    h: box.h,
    fontFace: resolveFontFace(theme, style?.font ?? defaultFontRoleFor(role)),
    fontSize: style?.fontSize ?? theme.fontSizes[role],
    color: resolveColor(theme, style?.color, "mutedText"),
    bold: style?.bold ?? false,
    italic: style?.italic ?? false,
    align: style?.align ?? "left",
    valign: style?.valign ?? "middle",
    margin: 0,
    isTextBox: true,
    ...(style?.lineHeight ? { lineSpacingMultiple: Number(style.lineHeight.toFixed(2)) } : {}),
    ...(style?.charSpacing !== undefined ? { charSpacing: style.charSpacing } : {}),
  };
}

/**
 * Template の全 Master を PptxGenJS へ登録する。
 * 戻り値は `master.id → PptxGenJS master 名` の対応表。
 */
export function defineTemplateMasters(options: DefineMastersOptions): MasterNameMap {
  const { pptx, template, theme, assets, variables } = options;
  const names = new Map<string, string>();

  for (const master of template.masters) {
    const objects: NonNullable<PptxGenJS.SlideMasterProps["objects"]> = [];

    for (const element of master.elements ?? []) {
      if (element.kind === "text") {
        const text = expandTemplateText(element.text, variables);
        // 変数が空（例: organization 未設定）なら要素ごと省く。
        // 空のテキストボックスを残すより、無いことが分かる方がよい。
        if (text.trim().length === 0) continue;
        objects.push({ text: { text, options: textOptions(theme, element.box, element.style) } });
        continue;
      }

      if (element.kind === "shape") {
        const shapeProps: PptxGenJS.ShapeProps = {
          x: element.box.x,
          y: element.box.y,
          w: element.box.w,
          h: element.box.h,
        };
        if (element.style?.fill) {
          shapeProps.fill = {
            color: resolveColor(theme, element.style.fill),
            ...(element.style.transparency !== undefined
              ? { transparency: element.style.transparency }
              : {}),
          };
        }
        if (element.style?.line) {
          shapeProps.line = {
            color: resolveColor(theme, element.style.line.color),
            width: element.style.line.width ?? theme.shape.borderWidth,
            dashType: element.style.line.dashType ?? "solid",
          };
        }
        objects.push(element.shape === "line" ? { line: shapeProps } : { rect: shapeProps });
        continue;
      }

      // image: 事前に登録済みのアセットしか参照できない（外部 URL は不可）
      const asset = assets.resolve(element.assetId);
      objects.push({
        image: {
          data: asset.data,
          x: element.box.x,
          y: element.box.y,
          w: element.box.w,
          h: element.box.h,
          ...(element.transparency !== undefined ? { transparency: element.transparency } : {}),
          ...(element.sizing && element.sizing !== "stretch"
            ? { sizing: { type: element.sizing, w: element.box.w, h: element.box.h } }
            : {}),
        },
      });
    }

    const name = masterNameOf(template, master);
    const props: PptxGenJS.SlideMasterProps = {
      title: name,
      background: { color: resolveColor(theme, master.background?.color ?? "background") },
      objects,
    };

    if (master.slideNumber && (options.includeSlideNumbers ?? true)) {
      const style = master.slideNumber.style;
      const role = style?.role ?? "caption";
      props.slideNumber = {
        x: master.slideNumber.box.x,
        y: master.slideNumber.box.y,
        w: master.slideNumber.box.w,
        h: master.slideNumber.box.h,
        fontFace: resolveFontFace(theme, style?.font ?? "numeric"),
        fontSize: style?.fontSize ?? theme.fontSizes[role],
        color: resolveColor(theme, style?.color, "mutedText"),
        align: style?.align ?? "right",
        valign: style?.valign ?? "middle",
        margin: 0,
      };
    }

    pptx.defineSlideMaster(props);
    names.set(master.id, name);
  }

  return names;
}
