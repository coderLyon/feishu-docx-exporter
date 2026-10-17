/** 原生表格渲染：合并单元格、列宽、单元格内嵌块 */
import {
  BorderStyle,
  Paragraph,
  ShadingType,
  Table,
  TableCell,
  TableLayoutType,
  TableRow,
  TextRun,
  WidthType,
} from 'docx';
import { PX_TO_TWIP } from '../../shared/constants.js';
import { STYLE } from '../docx/styles.js';

// 黑色 + size 4（0.5pt，Word 默认粗细）：在线是浅灰，但用户要求 Word 里用默认黑框线
// 框线用黑色 0.5pt（用户明确要求；曾一度跟随官方导出改成浅灰 DEE0E3，被指出"框线丢了"，已改回）。
// 表头底纹沿用官方导出的 EFF0F1。
const TABLE_BORDER_COLOR = '000000';
const TABLE_HEADER_FILL = 'EFF0F1';
const BORDER = { style: BorderStyle.SINGLE, size: 4, color: TABLE_BORDER_COLOR };

/**
 * 按页面可用宽度缩放列宽。
 *
 * 不能简单等比缩放：等比会把「字段编码」这类窄列压到连 VARCHAR(40) 都放不下，
 * 英文单词被硬切断行。这里迭代缩放，先压宽列；一旦某列降到最小可用宽度
 * （MIN，约 1.27cm，能容纳 8 个半角字符）就锁定不再压。
 */
/**
 * 计算表格各列宽度（twips）。
 *
 * 老做法是「按原比例缩放 + 统一 MIN=720」，结果是把「日期」「编码」这类
 * 不可断词的长 token 压到折行（2026/09/21 被拆成两行），看起来很难受。
 * 现在改为**内容感知**：
 *   1. 先给每列一个下限 = 该列最长不可断 token 的宽度（不低于 720）
 *   2. 剩下的空间按原始比例分给各列（宽列多吃）
 *   3. 若下限之和已超出版心，等比压缩（此时只能接受折行）
 */
/** 估算一段文本的显示宽度（px）：中文 1em，ASCII 约 0.52em */
function textWidthPx(text, fontSizePt = 10.5) {
  const em = fontSizePt * (96 / 72);
  let px = 0;
  for (const ch of String(text ?? '')) {
    px += /[\u4e00-\u9fa5\u3000-\u303f\uff00-\uffef]/.test(ch) ? em : em * 0.52;
  }
  return px;
}

/**
 * 计算表格列宽（twips），目标：**尽量不折行**。
 *
 * 两档宽度：
 *   tokenW[i]   —— 该列最长「不可断 token」的宽度（物理下限，再窄就断词）
 *   contentW[i] —— 该列最长单元格整行放下的宽度（理想值，放得下就一行）
 *
 * 分配策略：
 *   1. 所有 contentW 之和 ≤ 版心 → 直接按 contentW 分配（**零折行**），余量按内容比例摊
 *   2. 装不下 → 先给每列 tokenW 保底，再把剩余空间按「缺口比例」补给最需要的列
 *      （这样短列不浪费、长列也能多得，整体折行最少）
 */
export function computeColumnWidths(columnTexts, availableTwips, widthsPx = [], fontSizePt = 10.5) {
  const MIN = 600; // 约 30pt，再窄也放不下内容
  const PAD = 12; // 左右单元格内边距（px）
  const cols = columnTexts.map((texts) => {
    let tokenMax = 0;
    let contentMax = 0;
    for (const raw of texts) {
      const text = String(raw ?? '');
      if (!text) continue;
      const runs = text.match(/[A-Za-z0-9_\-/.%:]+/g) ?? [];
      const longestToken = runs.reduce((m, r) => Math.max(m, r.length), 0);
      tokenMax = Math.max(tokenMax, textWidthPx('x'.repeat(longestToken), fontSizePt));
      contentMax = Math.max(contentMax, textWidthPx(text, fontSizePt));
    }
    return {
      token: Math.max(MIN, Math.round((tokenMax + PAD) * PX_TO_TWIP)),
      content: Math.max(MIN, Math.round((contentMax + PAD) * PX_TO_TWIP)),
    };
  });

  // 基准用**飞书原始列宽**（作者已按内容调过），只在必要时做两件事：
  //   1. 把「装不下最长不可断 token」的列补到下限（避免 2026/09/21 被拆行）
  //   2. 若补齐后超出，从最富余的列（width - token 最大）按比例借
  // 这样既尊重原版面，又消除断词；纯内容驱动的版本实测会让页数变多。
  const hasSrc = Array.isArray(widthsPx) && widthsPx.length === cols.length;
  const src = hasSrc
    ? widthsPx.map((w) => Math.max(1, Math.round(w * PX_TO_TWIP)))
    : cols.map((c) => c.content); // 没有原始列宽时按内容等比
  const srcSum = src.reduce((a, b) => a + b, 0) || 1;
  const width = cols.map((c, i) => {
    const scaled = Math.floor((availableTwips * (src[i] ?? 1)) / srcSum);
    // 内容极短时不必撑太宽：不超过该列内容的整行宽度
    return Math.max(c.token, Math.min(scaled, Math.max(c.token, c.content)));
  });

  let sum = width.reduce((a, b) => a + b, 0);
  if (sum > availableTwips) {
    const over = sum - availableTwips;
    const slack = width.reduce((a, v, i) => a + Math.max(0, v - cols[i].token), 0) || 1;
    sum = 0;
    return width.map((v, i) => {
      const room = Math.max(0, v - cols[i].token);
      return v - Math.floor((over * room) / slack);
    });
  }

  // 还有余量：给「内容放不下」的列加宽，减少折行（按缺口比例分）
  const left = availableTwips - sum;
  const gaps = cols.map((c, i) => Math.max(0, c.content - width[i]));
  const gapSum = gaps.reduce((a, b) => a + b, 0);
  if (gapSum <= 0) return width;
  return width.map((w, i) => w + Math.floor((left * gaps[i]) / gapSum));
}

/** 兼容旧调用：按 px 宽度列表 + 每列内容下限计算 */
function fitWidths(widthsPx, availableTwips, columnTexts) {
  return computeColumnWidths(columnTexts, availableTwips, widthsPx);
}


/**
 * @param {Object} node table 块
 * @param {Object} ctx 渲染上下文
 */
export function renderTable(node, ctx) {
  const data = node.data ?? {};
  const columnsId = data.columns_id ?? [];
  const rowsId = data.rows_id ?? [];
  const cellSet = data.cell_set ?? {};
  const columnSet = data.column_set ?? {};

  if (!rowsId.length || !columnsId.length) return [];

  const widthsPx = columnsId.map((id) => columnSet[id]?.column_width ?? 120);

  // 逐列统计单元格文本，算出「不让长 token 折断」所需的列宽下限
  const columnTexts = columnsId.map((cid) => {
    const texts = [];
    for (const rid of rowsId) {
      const info = cellSet[`${rid}${cid}`];
      const block = info?.block_id ? ctx.byId.get(info.block_id) : null;
      if (!block) continue;
      const plain = block.data?.text
        ? Object.values(block.data.text.initialAttributedTexts?.text ?? {}).join('')
        : '';
      if (plain) texts.push(plain);
      for (const child of block.children ?? []) {
        const cb = ctx.byId.get(child);
        const childText = cb?.data?.text
          ? Object.values(cb.data.text.initialAttributedTexts?.text ?? {}).join('')
          : '';
        if (childText) texts.push(childText);
      }
    }
    return texts;
  });
  const widths = computeColumnWidths(columnTexts, ctx.pageContentWidth, widthsPx);

  // 表头判据来自飞书数据本身：data.header_row === true 表示 rows_id[0] 是表头。
  // false 的表首行是「请求参数（HEAD）」这类跨列合并的分节行，按表头渲染会出错。
  // 字段缺失时保守兜底：首行无列向合并、且每个单元格都有内容才当表头。
  const firstRowCells = columnsId.map((cid) => cellSet[`${rowsId[0]}${cid}`]);
  const firstRowNoMerge = firstRowCells.every(
    (info) => Math.max(1, info?.merge_info?.col_span ?? 1) === 1,
  );
  const firstRowFilled = firstRowCells.every((info) => {
    const b = info?.block_id ? ctx.byId.get(info.block_id) : null;
    return Boolean(b && (b.data?.text || b.children?.length));
  });
  const hasHeader =
    data.header_row === true ||
    (data.header_row === undefined && firstRowNoMerge && firstRowFilled);

  const covered = new Set();
  const rows = rowsId.map((rowId, rowIndex) => {
    const cells = [];

    for (let colIndex = 0; colIndex < columnsId.length; colIndex += 1) {
      const key = `${rowIndex}:${colIndex}`;
      if (covered.has(key)) continue;

      const info = cellSet[`${rowId}${columnsId[colIndex]}`];
      const rowSpan = Math.max(1, info?.merge_info?.row_span ?? 1);
      const colSpan = Math.max(1, info?.merge_info?.col_span ?? 1);

      for (let i = 0; i < rowSpan; i += 1) {
        for (let j = 0; j < colSpan; j += 1) {
          if (i || j) covered.add(`${rowIndex + i}:${colIndex + j}`);
        }
      }

      let children;
      const cellBlock = info?.block_id ? ctx.byId.get(info.block_id) : null;
      // 进入单元格：切换为表格专用字号/行距，并记录该列可用宽度。
      // 单元格里的图片必须按列宽缩放，否则会按版心宽度撑破表格（实测某文档图示列被拉成竖条）
      ctx.tableDepth += 1;
      const prevCellWidth = ctx.cellWidthPx;
      ctx.cellWidthPx = Math.round((widths[colIndex] ?? ctx.pageContentWidth) / 15);
      try {
        if (cellBlock && cellBlock.children?.length) {
          children = ctx.renderChildren(cellBlock.children, 0);
        } else {
          children = [new Paragraph({ children: [new TextRun({ text: '' })], style: STYLE.tableText })];
        }
      } finally {
        ctx.tableDepth -= 1;
        ctx.cellWidthPx = prevCellWidth;
      }
      if (!children.length) children.push(new Paragraph({ text: '', style: STYLE.tableText }));

      // 飞书把首行渲染成表头：浅灰底 + 跨页重复。
      // 接口类文档的表格首行几乎都是表头，块数据里没有表头标记，只能按此约定还原
      const isHeaderRow = hasHeader && rowIndex === 0;
      cells.push(
        new TableCell({
          children,
          ...(rowSpan > 1 ? { rowSpan } : {}),
          ...(colSpan > 1 ? { columnSpan: colSpan } : {}),
          ...(isHeaderRow
            ? { shading: { type: ShadingType.CLEAR, fill: TABLE_HEADER_FILL } }
            : {}),
          width: { size: widths[colIndex] ?? 1200, type: WidthType.DXA },
          margins: { top: 60, bottom: 60, left: 80, right: 80 },
        }),
      );
    }

    if (!cells.length) cells.push(new Paragraph({ text: '', style: STYLE.tableText }));
    return new TableRow({
      // 一行不允许被劈成两半跨页（否则单元格文字会在页边界处截断）
      cantSplit: true,
      children: cells,
      ...(hasHeader && rowIndex === 0 ? { tableHeader: true } : {}),
    });
  });

  return [
    new Table({
      rows,
      columnWidths: widths,
      // FIXED 布局：columnWidths 的 DXA 值才会被严格采用。
      // AUTOFIT 会让 Word 按内容重新分配列宽，在线的列宽比例就丢了
      layout: TableLayoutType.FIXED,
      width: { size: 100, type: WidthType.PERCENTAGE },
      borders: {
        top: BORDER,
        bottom: BORDER,
        left: BORDER,
        right: BORDER,
        insideHorizontal: BORDER,
        insideVertical: BORDER,
      },
    }),
    new Paragraph({ text: '', style: STYLE.body }),
  ];
}
