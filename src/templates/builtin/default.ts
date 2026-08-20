import { defaultTheme } from "../../theme/theme.js";
import { TEMPLATE_SCHEMA_VERSION, type PresentationTemplate } from "../template.js";

/**
 * ビルトイン Template 'default'。
 *
 * V1 の紙面（余白・ページ番号のみ）を Template として表現したもの。
 * Template を明示しない生成は今まで通り Template 非使用の経路を通るため、
 * これは「Template の最小例」であり、既定値ではない。
 */
export const defaultTemplate: PresentationTemplate = {
  schemaVersion: TEMPLATE_SCHEMA_VERSION,
  id: "default",
  version: "1.0.0",
  name: "Default (V1 相当の余白とページ番号のみ)",
  theme: defaultTheme,
  masters: [
    {
      id: "default-title",
      role: "title",
      name: "Title",
      background: { color: "background" },
      elements: [{ kind: "shape", shape: "rect", box: { x: 0, y: 0, w: 0.34, h: 5.625 }, style: { fill: "primary" } }],
      contentMargins: { top: 0.42, right: 0.6, bottom: 0.4, left: 0.94 },
    },
    {
      id: "default-section",
      role: "section",
      name: "Section",
      background: { color: "surface" },
      elements: [{ kind: "shape", shape: "rect", box: { x: 0, y: 2.68, w: 1.4, h: 0.06 }, style: { fill: "primary" } }],
      contentMargins: { top: 0.42, right: 0.6, bottom: 0.4, left: 0.6 },
    },
    {
      id: "default-content",
      role: "content",
      name: "Content",
      background: { color: "background" },
      contentMargins: { top: 0.42, right: 0.6, bottom: 0.4, left: 0.6 },
      slideNumber: {
        box: { x: 8.2, y: 5.23, w: 1.2, h: 0.26 },
        style: { role: "caption", font: "numeric", color: "mutedText", align: "right", valign: "middle" },
      },
    },
  ],
  defaultMasterId: "default-content",
  titleMasterId: "default-title",
  sectionMasterId: "default-section",
  metadata: {
    description: "V1 と同じ余白・ページ番号だけを持つ最小の Template。",
  },
};
