/**
 * 最小限の XML パーサ。
 *
 * なぜ自前なのか:
 *  - 必要なのは OOXML の「要素・属性・テキスト」だけで、汎用 XML の全機能は要らない
 *  - `.pptx` は信頼できない入力なので、DOCTYPE / ENTITY を **構造的に拒否** したい
 *    （外部実体参照 = XXE の経路を最初から作らない）
 *  - PPTX の解析のためだけに大きな依存を増やしたくない
 *
 * 対応するもの: 要素 / 属性 / テキスト / 自己終了タグ / コメント / CDATA / 基本実体参照
 * 対応しないもの: DTD / 外部実体 / 名前空間解決（接頭辞は文字列として扱う）
 */

export interface XmlElement {
  /** 接頭辞込みのタグ名（例: 'p:sp'） */
  name: string;
  attrs: Record<string, string>;
  children: XmlElement[];
  /** 直下のテキストノードを連結したもの */
  text: string;
}

export interface ParseXmlOptions {
  /** 入力サイズの上限（文字数） */
  maxLength?: number;
  /** ネストの上限 */
  maxDepth?: number;
  /** 要素数の上限 */
  maxElements?: number;
}

export const DEFAULT_MAX_XML_LENGTH = 16 * 1024 * 1024;
export const DEFAULT_MAX_XML_DEPTH = 256;
export const DEFAULT_MAX_XML_ELEMENTS = 200_000;

export class XmlParseError extends Error {
  constructor(message: string) {
    super(message);
    this.name = "XmlParseError";
  }
}

const NAMED_ENTITIES: Record<string, string> = {
  amp: "&",
  lt: "<",
  gt: ">",
  quot: '"',
  apos: "'",
};

export function decodeXmlText(raw: string): string {
  return raw.replace(/&(#x?[0-9a-fA-F]+|[a-zA-Z]+);/g, (match, body: string) => {
    if (body.startsWith("#x") || body.startsWith("#X")) {
      const code = Number.parseInt(body.slice(2), 16);
      return Number.isFinite(code) ? String.fromCodePoint(code) : match;
    }
    if (body.startsWith("#")) {
      const code = Number.parseInt(body.slice(1), 10);
      return Number.isFinite(code) ? String.fromCodePoint(code) : match;
    }
    return NAMED_ENTITIES[body] ?? match;
  });
}

function parseAttributes(source: string): Record<string, string> {
  const attrs: Record<string, string> = {};
  const pattern = /([^\s=/>]+)\s*=\s*("([^"]*)"|'([^']*)')/g;
  let match: RegExpExecArray | null;
  while ((match = pattern.exec(source)) !== null) {
    const key = match[1]!;
    // プロトタイプ汚染を避けるため、危険なキーは捨てる
    if (key === "__proto__" || key === "constructor" || key === "prototype") continue;
    attrs[key] = decodeXmlText(match[3] ?? match[4] ?? "");
  }
  return attrs;
}

/** XML 文字列を要素ツリーへ変換する。ルート要素が無ければ例外。 */
export function parseXml(source: string, options: ParseXmlOptions = {}): XmlElement {
  const maxLength = options.maxLength ?? DEFAULT_MAX_XML_LENGTH;
  const maxDepth = options.maxDepth ?? DEFAULT_MAX_XML_DEPTH;
  const maxElements = options.maxElements ?? DEFAULT_MAX_XML_ELEMENTS;

  if (source.length > maxLength) {
    throw new XmlParseError(`XML が大きすぎます (${source.length} > ${maxLength})`);
  }
  if (/<!DOCTYPE/i.test(source) || /<!ENTITY/i.test(source)) {
    throw new XmlParseError("DOCTYPE / ENTITY を含む XML は受け付けません（外部実体参照の防止）");
  }

  const root: XmlElement = { name: "#document", attrs: {}, children: [], text: "" };
  const stack: XmlElement[] = [root];
  let elements = 0;
  let i = 0;

  const current = (): XmlElement => stack[stack.length - 1]!;

  while (i < source.length) {
    const lt = source.indexOf("<", i);
    if (lt === -1) {
      current().text += decodeXmlText(source.slice(i));
      break;
    }
    if (lt > i) current().text += decodeXmlText(source.slice(i, lt));

    if (source.startsWith("<!--", lt)) {
      const end = source.indexOf("-->", lt + 4);
      i = end === -1 ? source.length : end + 3;
      continue;
    }
    if (source.startsWith("<![CDATA[", lt)) {
      const end = source.indexOf("]]>", lt + 9);
      const body = source.slice(lt + 9, end === -1 ? source.length : end);
      current().text += body;
      i = end === -1 ? source.length : end + 3;
      continue;
    }
    if (source.startsWith("<?", lt)) {
      const end = source.indexOf("?>", lt + 2);
      i = end === -1 ? source.length : end + 2;
      continue;
    }

    const gt = source.indexOf(">", lt);
    if (gt === -1) throw new XmlParseError("閉じられていないタグがあります");
    const raw = source.slice(lt + 1, gt).trim();

    if (raw.startsWith("/")) {
      const name = raw.slice(1).trim();
      if (stack.length <= 1) {
        throw new XmlParseError(`対応する開始タグがありません: </${name}>`);
      }
      const open = stack.pop()!;
      if (open.name !== name) {
        throw new XmlParseError(`タグの対応が取れていません: <${open.name}> ... </${name}>`);
      }
      i = gt + 1;
      continue;
    }

    const selfClosing = raw.endsWith("/");
    const body = selfClosing ? raw.slice(0, -1) : raw;
    const spaceAt = body.search(/[\s/]/);
    const name = spaceAt === -1 ? body : body.slice(0, spaceAt);
    const attrs = spaceAt === -1 ? {} : parseAttributes(body.slice(spaceAt));

    elements += 1;
    if (elements > maxElements) {
      throw new XmlParseError(`XML の要素数が多すぎます (> ${maxElements})`);
    }

    const element: XmlElement = { name, attrs, children: [], text: "" };
    current().children.push(element);
    if (!selfClosing) {
      stack.push(element);
      if (stack.length > maxDepth) {
        throw new XmlParseError(`XML のネストが深すぎます (> ${maxDepth})`);
      }
    }
    i = gt + 1;
  }

  if (stack.length !== 1) {
    throw new XmlParseError(`閉じられていない要素があります: <${current().name}>`);
  }
  const documentRoot = root.children[0];
  if (!documentRoot) throw new XmlParseError("ルート要素がありません");
  return documentRoot;
}

/** 'a:off' のような完全名、または 'off' のようなローカル名でマッチする */
function matches(element: XmlElement, name: string): boolean {
  if (element.name === name) return true;
  if (name.includes(":")) return false;
  const colon = element.name.indexOf(":");
  return colon >= 0 && element.name.slice(colon + 1) === name;
}

export function findChild(element: XmlElement, name: string): XmlElement | undefined {
  return element.children.find((child) => matches(child, name));
}

export function findChildren(element: XmlElement, name: string): XmlElement[] {
  return element.children.filter((child) => matches(child, name));
}

/** パス（例: 'spPr', 'xfrm', 'off'）をたどる */
export function findPath(element: XmlElement, ...names: string[]): XmlElement | undefined {
  let current: XmlElement | undefined = element;
  for (const name of names) {
    if (!current) return undefined;
    current = findChild(current, name);
  }
  return current;
}

/** 子孫を深さ優先で全探索する */
export function findDescendants(element: XmlElement, name: string): XmlElement[] {
  const out: XmlElement[] = [];
  const stack: XmlElement[] = [...element.children];
  while (stack.length > 0) {
    const node = stack.pop()!;
    if (matches(node, name)) out.push(node);
    for (let i = node.children.length - 1; i >= 0; i -= 1) stack.push(node.children[i]!);
  }
  return out;
}

export function hasDescendant(element: XmlElement, name: string): boolean {
  return findDescendants(element, name).length > 0;
}

export function attr(element: XmlElement | undefined, name: string): string | undefined {
  if (!element) return undefined;
  if (element.attrs[name] !== undefined) return element.attrs[name];
  // 接頭辞付き属性（r:embed など）をローカル名でも引けるようにする
  for (const [key, value] of Object.entries(element.attrs)) {
    const colon = key.indexOf(":");
    if (colon >= 0 && key.slice(colon + 1) === name) return value;
  }
  return undefined;
}

export function numAttr(element: XmlElement | undefined, name: string): number | undefined {
  const raw = attr(element, name);
  if (raw === undefined) return undefined;
  const value = Number(raw);
  return Number.isFinite(value) ? value : undefined;
}

/** 要素以下のテキストをすべて連結する */
export function textOf(element: XmlElement): string {
  let out = element.text;
  for (const child of element.children) out += textOf(child);
  return out;
}
