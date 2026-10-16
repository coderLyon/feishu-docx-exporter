import { sanitizeXmlText } from '../shared/util.js';

/**
 * 飞书富文本（attributed text）解析。
 *
 * 数据形如：
 *   {
 *     apool: { nextNum, numToAttrib: { "0": ["bold","true"], "1": ["author",".."] } },
 *     initialAttributedTexts: {
 *       attribs: { "0": "*0*1+8*0+3i" },   // 属性区间
 *       text:    { "0": "..." }            // 对应文本
 *     }
 *   }
 *
 * attribs 语法：
 *   *N        引用 numToAttrib 中的第 N 个属性
 *   |N        行/段标记（不参与文本切分，忽略）
 *   +LEN      接下来 LEN 个字符（LEN 为 36 进制，可多字符，如 +3i = 126）
 */

const BASE36 = /[0-9a-zA-Z]/;

/** 把属性区间字符串拆成 op 列表 */
export function tokenizeAttribs(source) {
  const ops = [];
  let i = 0;

  while (i < source.length) {
    const attrs = [];
    let line = null;

    while (source[i] === '*') {
      let j = i + 1;
      while (j < source.length && BASE36.test(source[j])) j += 1;
      if (j === i + 1) break;
      attrs.push(source.slice(i + 1, j));
      i = j;
    }

    if (source[i] === '|') {
      let j = i + 1;
      while (j < source.length && BASE36.test(source[j])) j += 1;
      line = source.slice(i + 1, j);
      i = j;
    }

    if (source[i] === '+') {
      let j = i + 1;
      while (j < source.length && BASE36.test(source[j])) j += 1;
      const raw = source.slice(i + 1, j);
      const len = Number.parseInt(raw, 36);
      if (Number.isFinite(len) && len >= 0) {
        ops.push({ attrs, line, len });
      }
      i = j;
      continue;
    }

    i += 1;
  }

  return ops;
}

/** 解析单个属性值：true/"..."/JSON */
function parseAttrValue(raw) {
  if (raw === 'true') return true;
  if (raw === 'false') return false;
  if (typeof raw === 'string' && raw.startsWith('{')) {
    try {
      return JSON.parse(raw);
    } catch {
      return raw;
    }
  }
  return raw;
}

const ATTR_ALIASES = {
  bold: 'bold',
  italic: 'italic',
  strikethrough: 'strike',
  strike: 'strike',
  underline: 'underline',
  inlinecode: 'inlineCode',
  code: 'inlineCode',
  link: 'link',
  url: 'link',
  formula: 'formula',
  equation: 'formula',
  mention: 'mention',
  'text-color': 'color',
  fontcolor: 'color',
  color: 'color',
  'background-color': 'bgColor',
  backcolor: 'bgColor',
  highlight: 'bgColor',
  // 飞书的「高亮底色」属性名就是这个（值为 rgba(255,246,122,0.8) 这类），
  // 漏掉它会导致 2.2 鉴权方式说明等黄色底色在 Word 里丢失
  texthighlightbackground: 'bgColor',
  'text-background': 'bgColor',
  'background-color-highlight': 'bgColor',
};

/**
 * 解析富文本块，返回 run 列表：
 *   [{ text, bold, italic, strike, underline, inlineCode, link, color, bgColor, formula, mention }]
 */
export function parseAttributedText(textBlock) {
  const pool = textBlock?.apool?.numToAttrib ?? {};
  const texts = textBlock?.initialAttributedTexts?.text;
  const attribs = textBlock?.initialAttributedTexts?.attribs;

  if (!texts || !attribs) return [];

  const runs = [];
  for (const segmentKey of Object.keys(texts)) {
    const segment = texts[segmentKey] ?? '';
    const pattern = attribs[segmentKey];
    if (!pattern) {
      if (segment) runs.push({ text: sanitizeXmlText(segment) });
      continue;
    }

    let cursor = 0;
    for (const op of tokenizeAttribs(pattern)) {
      if (op.len <= 0) continue;
      const slice = segment.slice(cursor, cursor + op.len);
      cursor += op.len;
      if (!slice) continue;

      const run = { text: sanitizeXmlText(slice) };
      for (const attrId of op.attrs) {
        const entry = pool[attrId];
        if (!entry) continue;
        const [name, rawValue] = entry;
        const key = ATTR_ALIASES[String(name).toLowerCase()];
        if (!key) continue;
        const value = parseAttrValue(rawValue);
        if (key === 'color' || key === 'bgColor') {
          run[key] = typeof value === 'string' ? value : value?.color ?? value;
        } else {
          run[key] = value;
        }
      }
      runs.push(run);
    }

    // 属性区间比文本短时，剩余部分按无样式输出
    if (cursor < segment.length) {
      runs.push({ text: segment.slice(cursor) });
    }
  }

  return runs;
}

/** 取块的纯文本（用于占位、文件名等） */
export function plainText(textBlock, max = 120) {
  const runs = parseAttributedText(textBlock);
  const text = runs.map((run) => run.text).join('').trim();
  return text.length > max ? `${text.slice(0, max)}…` : text;
}
