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
import { STYLE, BRAND_BLUE } from '../docx/styles.js';
import { BULLET_REFERENCE } from '../docx/numbering.js';
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

function indentFor(level) {
  return { left: 360 * level + 420, hanging: 360 };
}

/**
 * 按上下文挑段落样式：引用 > 表格 > 正文。
 * 字号 / 字体 / 行距 / 颜色全部由样式决定，段落上不再写直接格式。
 */
function styleIdFor(ctx) {
  if (ctx.quoteDepth > 0) return STYLE.quote;
  if (ctx.tableDepth > 0) return STYLE.tableText;
  return STYLE.body;
}

/**
 * 计算标题自动编号。
 *
 * 判据来自飞书块数据本身：**data.seq === 'auto' 的标题才参与自动编号**。
 * 正文章节的 7 个 heading1 都带 seq='auto'；「附件A/B/C」没有 seq 字段，
 * 因此既不补号也不占用计数器（这正是上一版用「附件」正则猜测会误判的原因）。
 *
 * 有 seq 但文本里已经写了编号（heading2/3 的编号打在文本里）时跳过，防止重复。
 */
function headingPrefix(node, ctx, level, text) {
  if (node.data?.seq !== 'auto') return '';

  ctx.headingCounters[level - 1] = (ctx.headingCounters[level - 1] ?? 0) + 1;
  for (let i = level; i < ctx.headingCounters.length; i += 1) ctx.headingCounters[i] = 0;

  if (NUMERIC_PREFIX.test(text)) return '';
  return `${ctx.headingCounters.slice(0, level).join('.')}. `;
}

/** 引用块左边框：飞书在线是品牌蓝，属于 3370FF 保留白名单 */
const QUOTE_BORDER = {
  left: { style: BorderStyle.SINGLE, size: 18, color: BRAND_BLUE, space: 12 },
};

/** 正文 / 标题 */
export function renderTextLike(node, ctx) {
  const runs = parseAttributedText(node.data?.text);
  const children = runsToDocxChildren(runs);

  if (node.type.startsWith('heading')) {
    const level = Math.min(9, Math.max(1, Number(node.type.slice(7)) || 1));
    const text = runs.map((run) => run.text).join('');
    const prefix = headingPrefix(node, ctx, level, text);
    // 编号 run 不写 size / color / font：继承 Heading 样式 => 三号黑色加粗，
    // 与标题文字自动同字号同色（r005 需求 2 + 需求 3）
    const headingChildren = prefix ? [new TextRun({ text: prefix, bold: true }), ...children] : children;

    return [
      new Paragraph({
        children: headingChildren,
        style: STYLE.heading(level),
        alignment: alignOf(node.data),
      }),
      ...ctx.renderChildren(node.children, 0),
    ];
  }

  return [
    new Paragraph({
      children: children.length ? children : [new TextRun({ text: '' })],
      style: styleIdFor(ctx),
      alignment: alignOf(node.data),
    }),
    ...ctx.renderChildren(node.children, 0),
  ];
}

/**
 * 无序列表。
 * 用自建的 fde-bullet 编号定义 + custom:true：这样段落上还能指定 FdeQuote /
 * FdeTableText 等样式，docx 不会额外插入一个 ListParagraph 的 pStyle
 * （一个 pPr 里出现两个 pStyle 会被 Word 判为文档损坏）。
 */
export function renderBullet(node, ctx, level = 0) {
  const runs = parseAttributedText(node.data?.text);
  return [
    new Paragraph({
      children: runsToDocxChildren(runs),
      numbering: { reference: BULLET_REFERENCE, level, custom: true },
      style: styleIdFor(ctx),
      alignment: alignOf(node.data),
    }),
    ...ctx.renderChildren(node.children, level + 1),
  ];
}

/** 有序列表（Word 原生编号，按列表段自动重新从 1 计数） */
export function renderOrdered(node, ctx, level = 0) {
  const runs = parseAttributedText(node.data?.text);
  return [
    new Paragraph({
      children: runsToDocxChildren(runs),
      numbering: { reference: `fde-ordered-${ctx.ordered.current}`, level },
      style: styleIdFor(ctx),
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
      style: styleIdFor(ctx),
      indent: indentFor(level),
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
export function renderQuoteContainer(node, ctx, level = 0) {
  ctx.quoteDepth += 1;
  try {
    return ctx.renderChildren(node.children, level);
  } finally {
    ctx.quoteDepth -= 1;
  }
}

/** 高亮块（callout）：用单单元格表格模拟带底色的容器 */
export function renderCallout(node, ctx, level = 0) {
  const fill = toHexColor(node.data?.background_color) ?? 'F5F6F7';
  const inner = ctx.renderChildren(node.children, level);
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

  const paragraphs = lines.map(
    (line) =>
      new Paragraph({
        children: [new TextRun({ text: line })],
        style: STYLE.code,
      }),
  );

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

export { alignOf, indentFor, QUOTE_BORDER };
