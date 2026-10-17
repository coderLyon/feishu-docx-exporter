/**
 * 段落 / 字符样式体系。
 *
 * 设计原则（r005）：**每类文本一种段落样式，段落上不写任何与之冲突的直接格式**。
 * 这样用户在 Word 里改一次「Fde 正文」样式，全文正文就跟着变，
 * 而不是被成百上千个 run 级 size/font 覆盖住。
 *
 * 行距统一单倍（SINGLE_LINE = line 240 / auto），只在这里定义一次。
 * 颜色：标题与序号纯黑 000000；蓝色 3370FF 只出现在三处白名单——
 *   ① Hyperlink 样式 ② @提及 ③ 引用块左边框（在线就是品牌蓝）。
 */
import { BorderStyle, ShadingType } from 'docx';
import { FONT_SCHEMES, SINGLE_LINE, TYPO } from '../../shared/constants.js';

/** 样式 id 常量：渲染层只引用这里，不写字面量 */
export const STYLE = {
  body: 'FdeBody',
  // 列表段落沿用 Word 内建的 ListParagraph id：
  // docx 对 bullet/numbering 段落会自动写 pStyle=ListParagraph，
  // 自定义 id 会导致一个 pPr 里出现两个 pStyle（Word 会判定文档损坏并修复），
  // 因此改为覆盖内建样式本身，列表就只有一种样式
  list: 'ListParagraph',
  // 引用直接挂到 Word 内建的「引用」样式（id=Quote），用户在样式面板里改它即可
  quote: 'Quote',
  code: 'FdeCode',
  caption: 'FdeCaption',
  // 代码块标题左对齐（飞书如此），图片题注才居中，两者不能共用一个样式
  codeCaption: 'FdeCodeCaption',
  tableText: 'FdeTableText',
  title: 'FdeTitle',
  inlineCode: 'FdeInlineCode',
  heading: (level) => `FdeHeading${Math.min(9, Math.max(1, level))}`,
};

export const INK = '000000';
export const BODY_INK = '1F2329';
export const BRAND_BLUE = '3370FF';
const MUTED = '646A73';
const CODE_FILL = 'F5F6F7';

const MONO = { ascii: 'Consolas', eastAsia: 'Consolas', hAnsi: 'Consolas', cs: 'Consolas' };

function runFont(font) {
  return { ascii: font.ascii, eastAsia: font.eastAsia, hAnsi: font.ascii, cs: font.ascii };
}

const spacing = (before = 0, after = 0) => ({ ...SINGLE_LINE, before, after });

/**
 * 生成 Document 的 styles 配置。
 * 中文必须显式指定 eastAsia，否则 Word 会回退成宋体。
 */
export function buildStyles(options) {
  const font = FONT_SCHEMES[options.fontScheme] ?? FONT_SCHEMES.yahei;
  const bodyFont = runFont(font);

  /** 标题样式：name 用 "heading N"，docx 的目录生成靠 name 匹配层级 */
  const headingStyles = [1, 2, 3, 4, 5, 6, 7, 8, 9].map((level) => {
    const spec = TYPO.heading[level];
    return {
      id: STYLE.heading(level),
      name: `heading ${level}`,
      run: { font: bodyFont, size: spec.size, bold: true, color: INK },
      paragraph: {
        outlineLevel: level - 1,
        spacing: spacing(spec.before, spec.after),
      },
    };
  });

  return {
    default: {
      document: {
        run: { font: bodyFont, size: TYPO.body.size, color: BODY_INK },
        paragraph: { spacing: spacing(TYPO.body.before, TYPO.body.after) },
      },
      hyperlink: {
        run: { color: BRAND_BLUE, underline: { type: 'single' } },
      },
    },

    paragraphStyles: [
      {
        id: STYLE.title,
        name: 'Fde 文档标题',
        run: { font: bodyFont, size: 45, bold: true, color: INK },
        paragraph: { spacing: spacing(0, 240) },
      },
      {
        id: STYLE.body,
        name: 'Fde 正文',
        run: { font: bodyFont, size: TYPO.body.size, color: BODY_INK },
        paragraph: { spacing: spacing(TYPO.body.before, TYPO.body.after) },
      },
      {
        // bullet / numbering 段落由 docx 自动应用，覆盖它以统一行距与字号
        id: STYLE.list,
        name: 'Fde 列表',
        run: { font: bodyFont, size: TYPO.body.size, color: BODY_INK },
        paragraph: { spacing: spacing(0, 0) },
      },
      {
        id: STYLE.tableText,
        name: 'Fde 表格文字',
        run: { font: bodyFont, size: TYPO.table.size, color: BODY_INK },
        paragraph: { spacing: spacing(0, 0) },
      },
      {
        id: STYLE.quote,
        name: '引用',
        basedOn: STYLE.body,
        run: { font: bodyFont, size: TYPO.body.size, color: BODY_INK },
        paragraph: {
          // 引用块的竖线改由外层单格表格的左边框承载（见 renderQuoteContainer）：
          // 加在样式上会和列表编号的缩进打架，出现两条不对齐的竖线。
          // 这里只保留段落间距，缩进交给单元格内边距（在线为 12px）
          spacing: spacing(60, 60),
        },
      },
      {
        id: STYLE.code,
        name: 'Fde 代码',
        // eastAsia 也要写：代码块里常有中文注释，不写会回退宋体
        run: { font: MONO, size: 20, color: BODY_INK },
        paragraph: {
          spacing: spacing(0, 0),
          indent: { left: 120, right: 120 },
          shading: { type: ShadingType.CLEAR, fill: CODE_FILL },
        },
      },
      {
        id: STYLE.codeCaption,
        name: 'Fde 代码标题',
        run: { font: bodyFont, size: 18, color: MUTED, italics: true },
        paragraph: { spacing: spacing(40, 60), alignment: 'left' },
      },
      {
        id: STYLE.caption,
        name: 'Fde 题注',
        run: { font: bodyFont, size: 18, color: MUTED, italics: true },
        paragraph: { spacing: spacing(40, 120), alignment: 'center' },
      },
      ...headingStyles,
    ],

    characterStyles: [
      {
        id: STYLE.inlineCode,
        name: 'Fde 行内代码',
        run: {
          font: MONO,
          color: BODY_INK,
          shading: { type: ShadingType.CLEAR, fill: 'F2F3F5' },
        },
      },
    ],
  };
}
