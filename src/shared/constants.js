/** 飞书块类型与样式常量 */

/** client_vars 的字符串块类型 */
export const BLOCK_TYPE = {
  PAGE: 'page',
  TEXT: 'text',
  HEADING: 'heading',
  BULLET: 'bullet',
  ORDERED: 'ordered',
  CODE: 'code',
  QUOTE: 'quote',
  QUOTE_CONTAINER: 'quote_container',
  TODO: 'todo',
  CALLOUT: 'callout',
  DIVIDER: 'divider',
  IMAGE: 'image',
  TABLE: 'table',
  TABLE_CELL: 'table_cell',
  SHEET: 'sheet',
  BITABLE: 'bitable',
  FILE: 'file',
  GRID: 'grid',
  GRID_COLUMN: 'grid_column',
  VIEW: 'view',
  WHITEBOARD: 'whiteboard',
  DIAGRAM: 'diagram',
  MINDNOTE: 'mindnote',
  IFRAME: 'iframe',
  ISV: 'isv',
  CHAT_CARD: 'chat_card',
  OKR: 'okr',
  TASK: 'task',
  BOARD: 'board',
};

/** 标题层级：heading1..heading9 */
export const HEADING_LEVEL = {
  heading1: 'Heading1',
  heading2: 'Heading2',
  heading3: 'Heading3',
  heading4: 'Heading4',
  heading5: 'Heading5',
  heading6: 'Heading6',
  heading7: 'Heading7',
  heading8: 'Heading7',
  heading9: 'Heading7',
};

/**
 * 排版基准：字号来自在线文档实测（见 samples/style-probe.json），
 * 标题字号按 Word 中文字号口径对齐：一号标题=三号(16pt)、二号=四号(14pt)、三号=小四(12pt)。
 *
 * ⚠️ 行距：全部改单倍（SINGLE_LINE = { line: 240, lineRule: auto }）。
 *   上一轮用「绝对行高 + atLeast」（正文 432 twips）是为了压住微软雅黑 1.714em 的
 *   自然行高，但 atLeast 会让 Word 忽略样式里的行距设置、也不跟随字号变化。
 *   本轮按用户要求回到单倍，行距只由段落样式统一承载（styles.js），
 *   段落自身不再写任何 spacing.line，保证「改样式即全文生效」。
 *
 * 只保留 size / before / after；before、after 单位为 twips（px × 15）。
 */
export const TYPO = {
  body: { size: 24, before: 0, after: 0 },
  table: { size: 21, before: 0, after: 0 },
  heading: {
    1: { size: 32, before: 390, after: 150 },
    2: { size: 28, before: 330, after: 120 },
    3: { size: 24, before: 270, after: 100 },
    4: { size: 24, before: 240, after: 90 },
    5: { size: 23, before: 210, after: 80 },
    6: { size: 22, before: 180, after: 70 },
    7: { size: 22, before: 160, after: 60 },
    8: { size: 22, before: 140, after: 60 },
    9: { size: 22, before: 140, after: 60 },
  },
};

/**
 * 单倍行距。Word 的 auto 倍数以 240 为 1.0 倍（12pt = 240 twips）基准。
 * 所有段落样式统一引用它，段落级不再另行指定行距。
 */
export const SINGLE_LINE = { line: 240, lineRule: 'auto' };

/** 字体方案 */
export const FONT_SCHEMES = {
  yahei: { ascii: 'Microsoft YaHei', eastAsia: '微软雅黑', label: '微软雅黑' },
  songti: { ascii: 'SimSun', eastAsia: '宋体', label: '宋体' },
  fangsong: { ascii: 'FangSong', eastAsia: '仿宋', label: '仿宋' },
  kaiti: { ascii: 'KaiTi', eastAsia: '楷体', label: '楷体' },
};

/** 页面尺寸（twips） */
export const PAGE_SIZES = {
  A4: { width: 11906, height: 16838, label: 'A4' },
  A3: { width: 16838, height: 23811, label: 'A3' },
  letter: { width: 12240, height: 15840, label: 'Letter' },
};

export const PAGE_MARGINS = {
  // 默认左右 900（0.625"）：飞书正文列约 824px，A4 版心越宽，表格列宽/换行越接近在线。
  // 1800 的旧默认会把版心压到 553px，导致 VARCHAR(40) 这类字段名被硬切断行
  normal: { top: 1440, bottom: 1440, left: 900, right: 900, label: '标准' },
  narrow: { top: 720, bottom: 720, left: 720, right: 720, label: '窄' },
  wide: { top: 1800, bottom: 1800, left: 1800, right: 1800, label: '宽' },
};

/** 默认导出选项 */
export const DEFAULT_OPTIONS = {
  embedImages: true,
  removeWatermark: true,
  includeToc: false,
  // 内嵌电子表格：默认导出为图片（版式忠实、零风险）；
  // 选 'table' 则导出为可编辑的 Word 原生表格（需要先用 extract-sheets 取到单元格数据）
  embeddedSheet: 'image',
  pageSize: 'A4',
  pageMargin: 'normal',
  fontScheme: 'yahei',
};

/** px -> twips（96 DPI 下 1px = 15 twips） */
export const PX_TO_TWIP = 15;

/** px -> EMU（1px = 9525 EMU） */
export const PX_TO_EMU = 9525;

/** 图片并发下载数 */
export const IMAGE_CONCURRENCY = 5;
