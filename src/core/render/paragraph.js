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
import { BULLET_REFERENCE } from '../docx/numbering.js';
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

function indentFor(level) {
  return { left: 360 * level + 420, hanging: 360 };
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

// 竖线颜色/粗细取自飞书官方导出（#BBBFC4 / sz=18）；正文为官方同款灰 #646A73（QUOTE_INK）。
// 若要与竖线**完全相同**，把这里改成 QUOTE_INK 即可。
const QUOTE_BAR = 'BBBFC4';
const QUOTE_BAR_SIZE = 18;

/** 正文 / 标题 */
export function renderTextLike(node, ctx) {
  const runs = parseAttributedText(node.data?.text);
  const children = runsToDocxChildren(runs);

  if (node.type.startsWith('heading')) {
    const level = Math.min(9, Math.max(1, Number(node.type.slice(7)) || 1));
    const text = runs.map((run) => run.text).join('');
    // 计数器照旧维护：别处（如报告摘要）可能读它
    headingPrefix(node, ctx, level, text);

    // 编号交给 Word 的 numbering：插入/删除标题后编号自动重排。
    //
    // 关键点：飞书里大量标题的编号是**作者手打在文本里**的（"5. 子表更新"、"6.2 状态流转"），
    // 只给 seq==='auto' 挂编号会让大部分标题没有编号。所以这里统一处理：
    //   1) 附件/附录类标题保留原样、不编号
    //   2) 其余标题剥掉手打的数字前缀，改由 Word 编号
    // 这样插入/删除标题时编号会自动重排 —— 手打编号做不到这一点。
    const isAttachment = /^\s*(附件|附录)/.test(text);
    let headingChildren = children;
    if (!isAttachment && NUMERIC_PREFIX.test(text)) {
      const stripped = runs.map((run, i) =>
        i === 0 ? { ...run, text: String(run.text ?? '').replace(NUMERIC_PREFIX, '') } : run,
      );
      headingChildren = runsToDocxChildren(stripped);
    }

    return [
      new Paragraph({
        children: headingChildren,
        style: STYLE.heading(level),
        alignment: alignOf(node.data),
        // 编号由标题样式自带（见 styles.js），附件标题在打包后取消编号
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
/**
 * 引用内列表的缩进覆盖。
 *
 * 编号定义里的 level0 是 left=720 / hanging=360（符号落在 18pt 处），
 * 这在正文里没问题，但引用里正文从头排版（符号应与正文同 x 起），
 * 否则「适用场景」在 x=7.2、下面两个 bullet 却在 x=25.2 —— 看起来就是没对齐。
 * 在线与飞书官方导出都是「符号与引用正文同 x，文字悬出」。
 * pPr 里的 w:ind 会覆盖编号级缩进，所以这里直接给段落写缩进即可。
 */
function quoteListIndent(ctx, level) {
  if (!ctx.quoteDepth) return undefined;
  const base = 360 + level * 420; // 0 级符号贴齐正文，逐级 21pt
  return { left: base, hanging: base };
}

export function renderBullet(node, ctx, level = 0) {
  const runs = parseAttributedText(node.data?.text);
  return [
    new Paragraph({
      children: runsToDocxChildren(runs),
      // 树深度从 1 开始，编号定义是 0 基 —— 不减 1 的话一级列表会套上二级样式（○、缩进 720）
      numbering: { reference: BULLET_REFERENCE, level: Math.max(0, level - 1), custom: true },
      indent: quoteListIndent(ctx, level),
      // 单元格里的列表要跟随表格字号，否则同一格内 10.5pt 与 12pt 混排
      style: ctx.quoteDepth > 0 ? STYLE.quote : ctx.tableDepth > 0 ? STYLE.tableText : STYLE.listBullet,
      alignment: alignOf(node.data),
    }),
    ...ctx.renderChildren(node.children, level + 1),
  ];
}

/** 有序列表（Word 原生编号，按列表段自动重新从 1 计数） */
export function renderOrdered(node, ctx, level = 0, instance = ctx.ordered.instance) {
  const runs = parseAttributedText(node.data?.text);
  return [
    new Paragraph({
      children: runsToDocxChildren(runs),
      numbering: { reference: `fde-ordered-${instance}`, level },
      indent: quoteListIndent(ctx, level),
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

export { alignOf, indentFor, QUOTE_BAR };
