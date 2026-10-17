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
function fitWidths(widthsPx, availableTwips, contentMinTwips = []) {
  const MIN = 720;
  const base = widthsPx.map((w) => Math.max(1, Math.round(w * PX_TO_TWIP)));
  const mins = base.map((_, i) => Math.max(MIN, Math.round(contentMinTwips[i] ?? MIN)));
  const minSum = mins.reduce((a, b) => a + b, 0);

  if (minSum >= availableTwips) {
    const k = availableTwips / minSum;
    return mins.map((w) => Math.max(240, Math.floor(w * k)));
  }

  const rest = availableTwips - minSum;
  const baseSum = base.reduce((a, b) => a + b, 0) || 1;
  return mins.map((m, i) => m + Math.floor((rest * base[i]) / baseSum));
}

/**
 * 估算一列需要的下限宽度：取该列最长「不可断 token」的宽度。
 * 中文可任意断行，只按 2 个字算；ASCII/数字串（日期、编码、URL）整段不可断。
 */
function contentMinFor(cellTexts, fontSizePt = 10.5) {
  const pxPerAscii = fontSizePt * 0.52 * (96 / 72);
  const pxPerCjk = fontSizePt * (96 / 72);
  const PAD = 12; // 左右单元格内边距
  let maxPx = 0;
  for (const text of cellTexts) {
    const raw = String(text ?? '');
    // 拆出 ASCII 连续串（含数字、下划线、斜杠、冒号等），中文单独计 2 字宽
    const asciiRuns = raw.match(/[A-Za-z0-9_\-/.%:]+/g) ?? [];
    const longest = asciiRuns.reduce((m, r) => Math.max(m, r.length), 0);
    const width = longest * pxPerAscii + 2 * pxPerCjk + PAD;
    if (width > maxPx) maxPx = width;
  }
  return Math.round(maxPx * PX_TO_TWIP);
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
  const contentMins = columnTexts.map((texts) => contentMinFor(texts));
  const widths = fitWidths(widthsPx, ctx.pageContentWidth, contentMins);

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
