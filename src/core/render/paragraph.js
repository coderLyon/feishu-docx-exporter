/** 段落类块渲染：正文、标题（含自动编号）、列表、待办、引用、高亮块、代码块、分割线 */
import { feishuOnly, t } from '../../shared/messages.js';

/** 占位文字颜色（与既有的橙色降级提示区分：飞书专有内容用中性灰） */
const MUTED_COLOR = '646A73';
import {
  AlignmentType,
  BorderStyle,
  Paragraph,
  ShadingType,
  Table,
  TableCell,
  TableRow,
  TextRun,
  WidthType,
} from 'docx';
import { parseAttributedText } from '../text.js';
import { runsToDocxChildren } from '../docx/runs.js';
import { QUOTE_INK, STYLE } from '../docx/styles.js';
import { bulletReference, listIndent, listLevel, orderedNumbering } from '../docx/numbering.js';
import { orderedSequence } from './list-instance.js';
import { highlightLine, isDarkBackground, safeBackground } from '../docx/highlight.js';
import { toHexColor } from '../../shared/util.js';

const ALIGN = {
  center: AlignmentType.CENTER,
  right: AlignmentType.RIGHT,
  justify: AlignmentType.JUSTIFIED,
  left: AlignmentType.LEFT,
};

/** 文本里已经自带编号前缀（如 "1.1 编写说明"）时不再重复添加 */
const NUMERIC_PREFIX = /^\s*\d+(?:[.\-、]\s*\d+)*[.\-、]?\s+/;

function alignOf(data) {
  return ALIGN[data?.align] ?? undefined;
}

/**
 * 按上下文挑段落样式：引用 > 表格 > 正文。
 * 字号 / 字体 / 行距 / 颜色全部由样式决定，段落上不再写直接格式。
 */
function styleIdFor(ctx) {
  if (ctx.quoteDepth > 0) return STYLE.quote;
  // 标注块内的正文单独成类，便于用户只调标注块字号/行距
  if (ctx.calloutDepth > 0) return STYLE.callout;
  if (ctx.tableDepth > 0) return STYLE.tableText;
  return STYLE.body;
}

// 竖线颜色/粗细取自飞书官方导出（#BBBFC4 / sz=18）；正文为官方同款灰 #646A73（QUOTE_INK）。
// 若要与竖线**完全相同**，把这里改成 QUOTE_INK 即可。
const QUOTE_BAR = 'BBBFC4';
const QUOTE_BAR_SIZE = 18;

/** 正文 / 标题 */
export function renderTextLike(node, ctx, depth = 0) {
  const runs = parseAttributedText(node.data?.text);

  if (node.type.startsWith('heading')) {
    const level = Number(node.type.slice(7)) || 1;
    const text = runs.map((run) => run.text).join('');
    // Word 样式生成章节编号，手写前缀可能跨多个不同格式的文字片段。
    const isAttachment = /^\s*(附件|附录)/.test(text);
    let skip = isAttachment ? 0 : (text.match(NUMERIC_PREFIX)?.[0].length ?? 0);
    const headingRuns = runs.map((run) => {
      const sliced = run.text.slice(skip);
      skip -= run.text.length - sliced.length;
      return { ...run, text: sliced };
    }).filter((run) => run.text);

    return [
      new Paragraph({
        children: runsToDocxChildren(headingRuns),
        style: STYLE.heading(level),
        alignment: alignOf(node.data),
        numbering: isAttachment ? false : undefined,
      }),
      ...ctx.renderChildren(node.children, depth),
    ];
  }

  const children = runsToDocxChildren(runs);
  return [
    new Paragraph({
      children: children.length ? children : [new TextRun({ text: '' })],
      style: styleIdFor(ctx),
      alignment: alignOf(node.data),
      indent: depth > 0 ? { left: listIndent(depth - 1, compactList(ctx)).left } : undefined,
    }),
    ...ctx.renderChildren(node.children, depth),
  ];
}

function compactList(ctx) {
  return ctx.quoteDepth > 0 || ctx.tableDepth > 0 || ctx.calloutDepth > 0;
}

export function renderBullet(node, ctx, level = 0) {
  const runs = parseAttributedText(node.data?.text);
  return [
    new Paragraph({
      children: runsToDocxChildren(runs),
      numbering: { reference: bulletReference(compactList(ctx)), level: listLevel(level), custom: true },
      // 单元格里的列表要跟随表格字号，否则同一格内 10.5pt 与 12pt 混排
      style: ctx.quoteDepth > 0 ? STYLE.quote : ctx.tableDepth > 0 ? STYLE.tableText : STYLE.listBullet,
      alignment: alignOf(node.data),
    }),
    ...ctx.renderChildren(node.children, level + 1),
  ];
}

/** 有序列表使用可在 Word 中继续编辑的原生编号。 */
export function renderOrdered(node, ctx, level = 0, list) {
  const runs = parseAttributedText(node.data?.text);
  const start = list?.start ?? orderedSequence(node) ?? 1;
  return [
    new Paragraph({
      children: runsToDocxChildren(runs),
      numbering: {
        reference: orderedNumbering(start, compactList(ctx)),
        instance: list?.instance ?? ++ctx.listInstance,
        level: listLevel(level),
        custom: true,
      },
      style: ctx.quoteDepth > 0 ? STYLE.quote : ctx.tableDepth > 0 ? STYLE.tableText : STYLE.listOrdered,
      alignment: alignOf(node.data),
    }),
    ...ctx.renderChildren(node.children, level + 1),
  ];
}

/** 待办 */
export function renderTodo(node, ctx, level = 0) {
  const runs = parseAttributedText(node.data?.text);
  const done = node.data?.done === true;
  const children = runsToDocxChildren(runs);
  if (done) children.forEach((run) => { if (run instanceof TextRun) run.strike = true; });
  return [
    new Paragraph({
      children: [new TextRun({ text: done ? '☑ ' : '☐ ', bold: true }), ...children],
      style: ctx.quoteDepth > 0 ? STYLE.quote : ctx.tableDepth > 0 ? STYLE.tableText : STYLE.todo,
      indent: listIndent(level, compactList(ctx)),
    }),
    ...ctx.renderChildren(node.children, level + 1),
  ];
}

/**
 * 引用容器。
 *
 * 上一版在这里「重建段落」：`new Paragraph({ children: element.options?.children })`
 * —— docx@9 的 Paragraph 实例没有 options 属性，取到恒为 undefined，
 * 于是引用块里所有文字都被丢掉（5.1 末尾那 3 条 bullet 全空）。
 *
 * 正确做法：只推进 quoteDepth，让子块自己渲染，缩进与蓝色左边框由 FdeQuote 样式承载。
 */
/**
 * 引用块：渲染成**单格表格**，竖线用表格的左边框。
 *
 * 为什么不把边框加在段落样式上：引用里常有列表，列表段落带 numPr，
 * 而**编号自带的缩进会覆盖样式的缩进**，导致各段左边框落在不同的 x 上——
 * 看起来就是「两条不对齐的竖线」（用户实测反馈）。表格边框位置固定，一次解决。
 *
 * 规格取自在线文档实测（tools/probe-quote.mjs）：
 *   竖线 2px = 1.5pt、颜色 #BBBFC4（灰）、容器无底色、
 *   内容左缩进 12px（=180 twips）、文字仍为正常深色 #1F2329。
 */
export function renderQuoteContainer(node, ctx, level = 0) {
  ctx.quoteDepth += 1;
  let inner;
  try {
    inner = ctx.renderChildren(node.children, level);
  } finally {
    ctx.quoteDepth -= 1;
  }

  return [
    new Table({
      // 不左移：表格左边框（即引用竖线）必须**与版心正文左边界对齐**（在线文档就是这样）。
      // 引用正文的缩进由单元格左内边距提供，视觉上竖线贴着正文列、引用文字内缩约 10pt。
      // （注：飞书官方导出的做法相反——竖线甩到版心左侧 7pt，正文贴版心；用户明确要求按在线来）
      width: { size: 100, type: WidthType.PERCENTAGE },
      borders: {
        top: { style: BorderStyle.NONE },
        bottom: { style: BorderStyle.NONE },
        right: { style: BorderStyle.NONE },
        left: { style: BorderStyle.SINGLE, size: QUOTE_BAR_SIZE, color: QUOTE_BAR },
        insideHorizontal: { style: BorderStyle.NONE },
        insideVertical: { style: BorderStyle.NONE },
      },
      rows: [
        new TableRow({
          children: [
            new TableCell({
              // 左内边距 180 twips（9pt）≈ 在线的引用文字缩进 12px；
              // 其余按官方：上 60 / 下 30 / 右 120
              margins: { top: 60, bottom: 30, left: 180, right: 120 },
              children: inner.length
                ? inner
                : [new Paragraph({ children: [], style: STYLE.body })],
            }),
          ],
        }),
      ],
    }),
  ];
}

/** 高亮块（callout）：用单单元格表格模拟带底色的容器 */
export function renderCallout(node, ctx, level = 0) {
  // 作者可能选了深色背景；照搬会让正文变成「深底深字」，打印一片黑 —— 统一提亮
  const fill = safeBackground(toHexColor(node.data?.background_color) ?? 'F5F6F7');
  // 标注块内正文用 FdeCallout 样式（单独一类，便于统一调整）
  ctx.calloutDepth = (ctx.calloutDepth ?? 0) + 1;
  const inner = ctx.renderChildren(node.children, level);
  ctx.calloutDepth -= 1;
  const emoji = node.data?.emoji_value ? String(node.data.emoji_value) : '';

  return [
    new Table({
      width: { size: 100, type: WidthType.PERCENTAGE },
      borders: {
        top: { style: BorderStyle.SINGLE, size: 4, color: 'E1E4E8' },
        bottom: { style: BorderStyle.SINGLE, size: 4, color: 'E1E4E8' },
        left: { style: BorderStyle.SINGLE, size: 4, color: 'E1E4E8' },
        right: { style: BorderStyle.SINGLE, size: 4, color: 'E1E4E8' },
        insideHorizontal: { style: BorderStyle.NONE },
        insideVertical: { style: BorderStyle.NONE },
      },
      rows: [
        new TableRow({
          children: [
            new TableCell({
              shading: { type: ShadingType.CLEAR, fill },
              margins: { top: 140, bottom: 140, left: 200, right: 200 },
              children: [
                ...(emoji
                  ? [new Paragraph({ children: [new TextRun({ text: emoji })], style: styleIdFor(ctx) })]
                  : []),
                ...inner,
              ],
            }),
          ],
        }),
      ],
    }),
    new Paragraph({ text: '', style: STYLE.body }),
  ];
}

/** 代码块：等宽字体 + 灰底，保留换行；caption 放在代码块上方（与飞书一致） */
export function renderCode(node, ctx) {
  const runs = parseAttributedText(node.data?.text);
  const text = runs.map((run) => run.text).join('');
  const captionRuns = parseAttributedText(node.data?.caption?.text);
  const caption = captionRuns.map((run) => run.text).join('').trim();

  const lines = text.replace(/\n$/, '').split('\n');
  const language = node.data?.language ?? '';
  const background = 'F5F6F7';

  // 逐行着色：飞书块数据里没有 token 颜色，按语言做规则着色（字符串/注释/关键字/数字），
  // 目标是分段可辨，不追求还原在线主题
  const paragraphs = lines.map((line) => {
    const runs = highlightLine(line, language, background).map(
      (seg) =>
        new TextRun({
          text: seg.text,
          // 深色底必须配浅色字，否则打印出来是一团黑
          color: isDarkBackground(background) || seg.color !== '24292E' ? seg.color : undefined,
        }),
    );
    return new Paragraph({ children: runs.length ? runs : [new TextRun({ text: '' })], style: STYLE.code });
  });

  const result = [
    new Table({
      width: { size: 100, type: WidthType.PERCENTAGE },
      borders: {
        top: { style: BorderStyle.SINGLE, size: 4, color: 'E1E4E8' },
        bottom: { style: BorderStyle.SINGLE, size: 4, color: 'E1E4E8' },
        left: { style: BorderStyle.SINGLE, size: 4, color: 'E1E4E8' },
        right: { style: BorderStyle.SINGLE, size: 4, color: 'E1E4E8' },
        insideHorizontal: { style: BorderStyle.NONE },
        insideVertical: { style: BorderStyle.NONE },
      },
      rows: [
        new TableRow({
          // 代码块不允许跨页断开：否则下一页开头会孤零零剩下几行花括号
          cantSplit: true,
          children: [
            new TableCell({
              shading: { type: ShadingType.CLEAR, fill: 'F5F6F7' },
              margins: { top: 100, bottom: 100, left: 160, right: 160 },
              children: paragraphs.length ? paragraphs : [new Paragraph({ text: '', style: STYLE.code })],
            }),
          ],
        }),
      ],
    }),
  ];

  if (caption) {
    // 飞书把代码块标题显示在代码框上方；上一版 push 到下方，看起来像「被挪到末尾」。
    // 代码标题用左对齐的 FdeCodeCaption（图片题注才居中），否则「PLAN_WO_JDZJ 原建表语句」会居中
    result.unshift(
      new Paragraph({
        children: [new TextRun({ text: caption })],
        style: STYLE.codeCaption,
      }),
    );
  }

  // 相邻代码块之间要有间距：3.6 是两个 caption 为空的独立 code 块，
  // 直接相邻会渲染成两个无缝相接的框，看着像同一个框里的两段
  result.push(new Paragraph({ text: '', style: STYLE.body, spacing: { after: 120 } }));
  return result;
}

/** 分割线 */
export function renderDivider() {
  return [
    new Paragraph({
      text: '',
      style: STYLE.body,
      border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: 'D0D3D6', space: 1 } },
    }),
  ];
}

/** 未知/不支持块的兜底占位 */
/**
 * 飞书专有内容的占位（Word 无对应形式，不做强行兼容）。
 * 与 renderPlaceholder 的区别：文案说清「这是什么」，而不是「未支持的块 xxx」。
 */
export function renderFeishuOnly(node, ctx) {
  const label = feishuOnly(node.type, ctx.lang);
  ctx.report.degraded.push(label);
  return [
    new Paragraph({
      children: [new TextRun({ text: `[${label}]`, italics: true, color: MUTED_COLOR })],
      style: styleIdFor(ctx),
    }),
  ];
}

/** 未知/不支持块的兜底占位 */
export function renderPlaceholder(node, ctx, message) {
  const label = message ?? t('unsupported', ctx.lang)(node.type);
  ctx.report.degraded.push(label);
  return [
    new Paragraph({
      children: [new TextRun({ text: `[${label}]`, italics: true, color: 'FF8800' })],
      style: styleIdFor(ctx),
    }),
  ];
}
