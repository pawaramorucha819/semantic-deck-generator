import type { PresentationTheme } from "../../theme/theme.js";
import { TEMPLATE_SCHEMA_VERSION, type PresentationTemplate } from "../template.js";

/** corporate Template が持つ Theme（企業ブランドカラー相当） */
export const corporateTheme: PresentationTheme = {
  id: "corporate",
  name: "Corporate (Deep Navy)",
  fonts: {
    heading: "Yu Gothic",
    body: "Yu Gothic",
    numeric: "Segoe UI",
  },
  fontSizes: {
    keyMessage: 21,
    heading: 14,
    body: 12,
    caption: 9,
    display: 44,
  },
  spacing: { xs: 0.06, sm: 0.12, md: 0.2, lg: 0.32, xl: 0.5 },
  colors: {
    background: "FFFFFF",
    text: "16202E",
    mutedText: "63707F",
    primary: "10365F",
    secondary: "2C7BB6",
    accent: "C8622F",
    surface: "F3F6F9",
    border: "D2DAE3",
    positive: "1B7F62",
    negative: "B5462F",
    onPrimary: "FFFFFF",
  },
  margins: { top: 0.62, right: 0.6, bottom: 0.62, left: 0.6 },
  shape: { cornerRadius: 0.04, borderWidth: 1 },
};

/**
 * ビルトイン Template 'corporate'。
 *
 * 「企業テンプレートとは何か」を最小構成で示すサンプル。
 * ヘッダー帯・フッター罫線・組織名・機密表示・ページ番号を Master 側で定義しており、
 * Pattern 側には会社固有の要素が一切入らないことを実演する。
 */
export const corporateTemplate: PresentationTemplate = {
  schemaVersion: TEMPLATE_SCHEMA_VERSION,
  id: "corporate",
  version: "1.0.0",
  name: "Corporate Standard",
  theme: corporateTheme,
  masters: [
    {
      id: "corporate-title",
      role: "title",
      name: "Title",
      background: { color: "primary" },
      contentMargins: { top: 1.4, right: 0.9, bottom: 1.0, left: 0.9 },
      elements: [
        { kind: "shape", shape: "rect", box: { x: 0.9, y: 4.55, w: 1.6, h: 0.05 }, style: { fill: "accent" } },
        {
          kind: "text",
          box: { x: 0.9, y: 4.75, w: 8.2, h: 0.28 },
          text: "{{organization}}",
          style: { role: "caption", color: "onPrimary", charSpacing: 1, valign: "middle" },
        },
      ],
    },
    {
      id: "corporate-section",
      role: "section",
      name: "Section",
      background: { color: "surface" },
      contentMargins: { top: 2.0, right: 0.9, bottom: 1.2, left: 0.9 },
      elements: [
        { kind: "shape", shape: "rect", box: { x: 0, y: 0, w: 10, h: 0.09 }, style: { fill: "primary" } },
        { kind: "shape", shape: "rect", box: { x: 0.9, y: 1.6, w: 1.2, h: 0.05 }, style: { fill: "accent" } },
      ],
    },
    {
      id: "corporate-content",
      role: "content",
      name: "Content",
      background: { color: "background" },
      contentMargins: { top: 0.62, right: 0.6, bottom: 0.66, left: 0.6 },
      elements: [
        { kind: "shape", shape: "rect", box: { x: 0, y: 0, w: 10, h: 0.09 }, style: { fill: "primary" } },
        {
          kind: "text",
          box: { x: 5.4, y: 0.16, w: 4.0, h: 0.26 },
          text: "{{organization}}",
          style: { role: "caption", color: "mutedText", align: "right", valign: "middle" },
        },
        { kind: "shape", shape: "rect", box: { x: 0.6, y: 5.14, w: 8.8, h: 0.012 }, style: { fill: "border" } },
        {
          kind: "text",
          box: { x: 0.6, y: 5.2, w: 5.4, h: 0.26 },
          text: "{{confidentiality}}",
          style: { role: "caption", color: "mutedText", valign: "middle" },
        },
      ],
      slideNumber: {
        box: { x: 8.4, y: 5.2, w: 1.0, h: 0.26 },
        style: { role: "caption", font: "numeric", color: "mutedText", align: "right", valign: "middle" },
      },
    },
  ],
  defaultMasterId: "corporate-content",
  titleMasterId: "corporate-title",
  sectionMasterId: "corporate-section",
  metadata: {
    organization: "Example Corporation",
    description: "ヘッダー帯・フッター・機密表示・ページ番号を Master 側に持つ企業標準テンプレート。",
    confidentialityLabel: "社外秘 / Confidential",
  },
};
