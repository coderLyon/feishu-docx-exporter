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
  const PAD = 22; // 左右单元格内边距 + 边框占位（px）；实测 12 偏小会让 token 仍被拆
  const cols = columnTexts.map((texts) => {
    let tokenMax = 0;
    let contentMax = 0;
    for (const raw of texts) {
      const text = String(raw ?? '');
      if (!text) continue;
      const runs = text.match(/[A-Za-z0-9_\-/.%:()\[\]]+/g) ?? [];
      const longestToken = runs.reduce((m, r) => Math.max(m, r.length), 0);
      tokenMax = Math.max(tokenMax, textWidthPx('x'.repeat(longestToken), fontSizePt));
      contentMax = Math.max(contentMax, textWidthPx(text, fontSizePt));
    }
    // token 下限要封顶：像 PLAN_WO_JDZJ.DEVICE_PLAN_ID 这种 27 字符的标识符，
    // 如果硬要整行放会吃掉版心 1/3，把说明列挤到每行两三个字。
    // 飞书原稿里这类长串本来也是折行的，所以上限取「16 字符」与「版心 22%」的较小值。
    const capPx = Math.min(16 * fontSizePt * 0.52 * (96 / 72), availableTwips * 0.22 / PX_TO_TWIP);
    return {
      token: Math.max(MIN, Math.round((Math.min(tokenMax, capPx) + PAD) * PX_TO_TWIP)),
      content: Math.max(MIN, Math.round((contentMax + PAD) * PX_TO_TWIP)),
    };
  });

  // 排版策略（三轮实测后的结论）：
  //   1. 交给 Word AutoFit 会把握不住的列压成竖排（「字段名称」四个字一字一行）
  //   2. 纯手算又容易把说明列挤死或让表格变窄
  // 所以：固定布局 + 精确下限。每条列宽都要满足三件事：
  //   a) ≥ 表头/短内容能横排的最小宽度（避免一字一行）
  //   b) ≥ 关键 token 宽度（日期、编码不折断；超长标识符允许折行，故封顶）
  //   c) 合计正好等于版心宽度（否则表格变窄、右侧留白）
  const hasSrc = Array.isArray(widthsPx) && widthsPx.length === cols.length;
  const src = hasSrc
    ? widthsPx.map((w) => Math.max(1, Math.round(w * PX_TO_TWIP)))
    : cols.map((c) => c.content || 1);
  const srcSum = src.reduce((a, b) => a + b, 0) || 1;

  // 每列的下限：minChars（表头 3~4 字能横排）与 token 宽度的较大者
  const MIN_CHARS = 4;
  const floorPx = MIN_CHARS * fontSizePt * (96 / 72) + PAD;
  const floor = cols.map((c) => Math.max(Math.round(floorPx * PX_TO_TWIP), Math.min(c.token, Math.round(availableTwips * 0.24))));

  // 起点：原比例，但不低于下限
  const width = cols.map((c, i) => Math.max(floor[i], Math.round((availableTwips * src[i]) / srcSum)));

  let sum = width.reduce((a, b) => a + b, 0);

  // 超宽：从「余量最多」的列按比例削（不削到下限以下）
  if (sum > availableTwips) {
    const over = sum - availableTwips;
    const slack = width.map((w, i) => Math.max(0, w - floor[i]));
    const slackSum = slack.reduce((a, b) => a + b, 0) || 1;
    for (let i = 0; i < width.length; i += 1) {
      width[i] -= Math.floor((over * slack[i]) / slackSum);
    }
    sum = width.reduce((a, b) => a + b, 0);
  }

  // 安全阀：下限之和本身就超版心时（如 11 列每列都要 4 字宽），等比压缩；
  // 绝不能让某列被减成负数 —— 实测出现过 -344，Word 直接渲染失败
  let sum2 = width.reduce((a, b) => a + b, 0);
  if (sum2 > availableTwips) {
    const k = availableTwips / sum2;
    for (let i = 0; i < width.length; i += 1) width[i] = Math.max(120, Math.floor(width[i] * k));
  }

  // 归一化：差额补给「内容缺口最大」的列，保证正好铺满版心
  let sum3 = width.reduce((a, b) => a + b, 0);
  if (sum3 !== availableTwips) {
    const diff = availableTwips - sum3;
    let best = 0;
    let bestGap = -Infinity;
    for (let i = 0; i < width.length; i += 1) {
      const gap = cols[i].content - width[i];
      if (gap > bestGap) {
        bestGap = gap;
        best = i;
      }
    }
    width[best] = Math.max(120, width[best] + diff);
  }

  return width.map((w) => Math.max(120, Math.round(w)));
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
  // 注意：表格单元格块**不在 blockSequence 里**（只在 blockMap 中），
  // 所以不能只用 ctx.byId —— 之前正是这里取不到文本，导致列宽估算退化成「按原比例」，
  // 长 token（devicePlanId / VARCHAR(64)）照样被拆行。
  const getData = (id) => {
    if (!id) return null;
    const node = ctx.byId.get(id);
    if (node?.data) return node.data;
    const raw = ctx.blockMap?.[id];
    return raw?.data ?? raw ?? null;
  };
  const textOf = (data) =>
    data?.text ? Object.values(data.text.initialAttributedTexts?.text ?? {}).join('') : '';

  const columnTexts = columnsId.map((cid) => {
    const texts = [];
    for (const rid of rowsId) {
      const info = cellSet[`${rid}${cid}`];
      const cellData = getData(info?.block_id);
      if (!cellData) continue;
      const plain = textOf(cellData);
      if (plain) texts.push(plain);
      for (const child of cellData.children ?? []) {
        const childText = textOf(getData(child));
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
      // 固定布局：columnWidths 才会被严格采用。
      // （试过 AUTOFIT：Word 会把「字段名称」「必输」这类短列表头压成竖排，
      //   实测反而更糟；AutoFit 只适合作为参考，最终仍要给出精确列宽）
      columnWidths: widths,
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
