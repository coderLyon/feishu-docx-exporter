/**
 * 由网格模型生成最小可用的 .xlsx（SpreadsheetML），不依赖任何 xlsx 库。
 *
 * 目标不是做完整 Excel 实现，而是让「Word 里内嵌的 Excel 对象」双击打开后
 * 能看到与在线一致的表格：底色/字色/字号/加粗/对齐/合并/列宽/边框 + 能还原的数字格式。
 *
 * 包结构（6 个部件，用 inlineStr 省掉 sharedStrings，颜色全用 rgb 省掉 theme）：
 *   [Content_Types].xml
 *   _rels/.rels
 *   xl/workbook.xml
 *   xl/_rels/workbook.xml.rels
 *   xl/styles.xml
 *   xl/worksheets/sheet1.xml
 */
import JSZip from 'jszip';
import { colLetter } from '../sheet/grid.js';

const NS_MAIN = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const NS_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const NS_PKG_REL = 'http://schemas.openxmlformats.org/package/2006/relationships';
const XML_HEAD = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';

/** XML 转义 + 去掉非法控制字符（正文里曾出现过 0x08，会让 Word 拒绝打开） */
export function esc(text) {
  return String(text ?? '')
    // eslint-disable-next-line no-control-regex
    .replace(/[\u0000-\u0008\u000b\u000c\u000e-\u001f]/g, '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

/**
 * Excel 日期序列值：1899-12-30 起算。
 * 必须两端都按「本地墙钟时间」换算 —— 一边 Date.UTC、一边 getTime() 会差一个时区，
 * 实测北京时区下 2026/09/21 会算成 46285.667（少 1 天 8 小时）。
 */
function toSerial(date) {
  const epoch = Date.UTC(1899, 11, 30);
  const wall = Date.UTC(
    date.getFullYear(), date.getMonth(), date.getDate(),
    date.getHours(), date.getMinutes(), date.getSeconds(),
  );
  return (wall - epoch) / 86400000;
}

function decimalsOf(numText) {
  const m = String(numText).match(/\.(\d+)/);
  return m ? m[1].length : 0;
}

/**
 * 单元格取值 + 数字格式。
 *
 * 判定顺序：data-formatter → 文本形态 → 保守不转数字。
 * 说明：剪贴板只给显示文本，所以「货币/百分比/日期」是靠格式名 + 文本反推的，
 * 会丢隐藏精度（如 12.35% 实际可能是 0.123456），这是剪贴板路线的固有限制。
 */
export function inferCell(cell) {
  const raw = String(cell.text ?? '');
  const fmt = String(cell.formatter ?? '');
  const trimmed = raw.trim();

  if (!trimmed) return { kind: 'blank' };

  // 1) 百分比
  if (/percent/i.test(fmt) || /^-?[\d,]+(\.\d+)?\s*%$/.test(trimmed)) {
    const num = Number(trimmed.replace(/[%\s,]/g, ''));
    if (Number.isFinite(num)) {
      const d = decimalsOf(trimmed.replace(/[^\d.]/g, ''));
      return { kind: 'number', value: num / 100, numFmt: `0.${'0'.repeat(Math.max(d, 2))}%` };
    }
  }

  // 2) 货币（含负数的会计括号写法）
  const currencyMatch = trimmed.match(/^([(（])?\s*([¥$€£])?\s*(-?[\d,]+(?:\.\d+)?)\s*([)）])?$/);
  const isCurrency = /currenc/i.test(fmt) || (currencyMatch && currencyMatch[2]);
  if (isCurrency && currencyMatch) {
    const symbol = currencyMatch[2] || '¥';
    const neg = Boolean(currencyMatch[1]) || String(currencyMatch[3]).startsWith('-');
    const num = Number(String(currencyMatch[3]).replace(/[,\s-]/g, ''));
    if (Number.isFinite(num)) {
      const d = Math.max(decimalsOf(currencyMatch[3]), 2);
      return {
        kind: 'number',
        value: neg ? -num : num,
        numFmt: `${symbol}#,##0.${'0'.repeat(d)};[Red](${symbol}#,##0.${'0'.repeat(d)})`,
      };
    }
  }

  // 3) 日期 / 日期时间
  const dateMatch = trimmed.match(
    /^(\d{4})[-/年](\d{1,2})[-/月](\d{1,2})日?(?:[ T](\d{1,2}):(\d{2})(?::(\d{2}))?)?$/,
  );
  if ((/date|time/i.test(fmt) && dateMatch) || (!fmt && dateMatch)) {
    const [, y, mo, d, hh, mm, ss] = dateMatch;
    const date = new Date(Number(y), Number(mo) - 1, Number(d), Number(hh ?? 0), Number(mm ?? 0), Number(ss ?? 0));
    const serial = toSerial(date);
    const hasTime = hh != null;
    return {
      kind: 'number',
      value: serial,
      numFmt: hasTime ? 'yyyy/m/d h:mm' : 'yyyy/m/d',
    };
  }

  // 4) 普通数字：保守判定，宁可当文本
  const plain = /^-?(?:0|[1-9]\d{0,14})(?:\.\d{1,10})?$/.test(trimmed);
  const keepAsText =
    /^0\d/.test(trimmed) || // 前导零：00123
    trimmed.replace(/\D/g, '').length > 15 || // 长编号
    /^\d{8}$/.test(trimmed) || // 20260921 这种日期形态
    /%|¥|\$|€|£|,/.test(trimmed);
  if (plain && !keepAsText) {
    return { kind: 'number', value: Number(trimmed), numFmt: null };
  }

  return { kind: 'text', value: raw };
}

/** 列宽：像素 → Excel 字符宽（96DPI、Calibri 11 基准：px = 7 * width + 5） */
function widthFromPx(px) {
  const w = (Number(px) - 5) / 7;
  return Math.min(80, Math.max(3, Number.isFinite(w) ? w : 8.43));
}

/**
 * @param {ReturnType<import('../sheet/grid.js').parseSheetHtml>} grid
 * @param {{sheetName?: string, fontName?: string, freezeHeader?: boolean}} [opts]
 * @returns {Promise<Uint8Array>}
 */
export async function buildXlsx(grid, opts = {}) {
  const { sheetName = 'Sheet1', fontName = '微软雅黑', freezeHeader = true } = opts;

  // ---- 样式池 ----
  const fonts = [{ key: 'base', xml: `<font><sz val="10"/><color rgb="FF112233"/><name val="${esc(fontName)}"/><charset val="134"/><family val="2"/></font>` }];
  const fills = [
    { key: 'none', xml: '<fill><patternFill patternType="none"/></fill>' },
    { key: 'gray125', xml: '<fill><patternFill patternType="gray125"/></fill>' },
  ];
  const borders = [{ key: 'none', xml: '<border><left/><right/><top/><bottom/><diagonal/></border>' }];
  const xfs = [{ key: 'base', xml: '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>' }];
  const numFmts = [];
  const indexOf = (pool, key, make) => {
    const hit = pool.findIndex((p) => p.key === key);
    if (hit >= 0) return hit;
    pool.push({ key, xml: make() });
    return pool.length - 1;
  };

  const fontId = (cell) =>
    indexOf(fonts, `f|${cell.bold}|${cell.color}|${cell.sizePt}`, () => {
      const parts = [];
      if (cell.bold) parts.push('<b/>');
      parts.push(`<sz val="${cell.sizePt || 10}"/>`);
      parts.push(`<color rgb="FF${cell.color || '112233'}"/>`);
      parts.push(`<name val="${esc(fontName)}"/><charset val="134"/><family val="2"/>`);
      return `<font>${parts.join('')}</font>`;
    });

  const fillId = (cell) => {
    if (!cell.bg) return 0;
    return indexOf(fills, `b|${cell.bg}`, () =>
      `<fill><patternFill patternType="solid"><fgColor rgb="FF${cell.bg}"/><bgColor indexed="64"/></patternFill></fill>`);
  };

  const borderId = (cell, color) => {
    if (!grid.hasBorder && !cell.borderColor) return 0;
    const c = color || 'DEE0E3';
    return indexOf(borders, `d|${c}`, () =>
      `<border><left style="thin"><color rgb="FF${c}"/></left><right style="thin"><color rgb="FF${c}"/></right>` +
      `<top style="thin"><color rgb="FF${c}"/></top><bottom style="thin"><color rgb="FF${c}"/></bottom><diagonal/></border>`);
  };

  const numFmtId = (code) => {
    if (!code) return 0;
    // 164 起是自定义格式区的惯例起点
    const hit = numFmts.findIndex((n) => n.code === code);
    if (hit >= 0) return 164 + hit;
    numFmts.push({ code });
    return 164 + numFmts.length - 1;
  };

  const xfId = (cell, nf) => {
    const f = fontId(cell);
    const fl = fillId(cell);
    const bd = borderId(cell, cell.borderColor);
    const n = numFmtId(nf);
    const align = cell.align || 'left';
    const key = `x|${f}|${fl}|${bd}|${n}|${align}`;
    return indexOf(xfs, key, () =>
      `<xf numFmtId="${n}" fontId="${f}" fillId="${fl}" borderId="${bd}" xfId="0"` +
      ` applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1" applyNumberFormat="${n ? 1 : 0}">` +
      `<alignment horizontal="${align}" vertical="center" wrapText="1"/></xf>`);
  };

  // ---- 行/单元格 ----
  const byRow = new Map();
  for (const cell of grid.cells) {
    if (!byRow.has(cell.r)) byRow.set(cell.r, []);
    byRow.get(cell.r).push(cell);
  }

  const rowXml = [];
  for (let r = 0; r < grid.rowCount; r += 1) {
    const list = (byRow.get(r) ?? []).sort((a, b) => a.c - b.c);
    if (!list.length) {
      rowXml.push(`<row r="${r + 1}"/>`);
      continue;
    }
    const cellsXml = list.map((cell) => {
      const addr = `${colLetter(cell.c)}${r + 1}`;
      const inferred = inferCell(cell);
      const s = xfId(cell, inferred.numFmt);
      if (inferred.kind === 'blank') return `<c r="${addr}" s="${s}"/>`;
      if (inferred.kind === 'number') return `<c r="${addr}" s="${s}"><v>${inferred.value}</v></c>`;
      return `<c r="${addr}" s="${s}" t="inlineStr"><is><t xml:space="preserve">${esc(inferred.value)}</t></is></c>`;
    });
    rowXml.push(`<row r="${r + 1}">${cellsXml.join('')}</row>`);
  }

  const colsXml = (grid.colsPx.length ? grid.colsPx : Array.from({ length: grid.colCount }, () => 80))
    .slice(0, grid.colCount)
    .map((px, i) => `<col min="${i + 1}" max="${i + 1}" width="${widthFromPx(px).toFixed(2)}" customWidth="1"/>`)
    .join('');

  const mergesXml = grid.merges.length
    ? `<mergeCells count="${grid.merges.length}">${grid.merges
        .map((m) => `<mergeCell ref="${colLetter(m.c0)}${m.r0 + 1}:${colLetter(m.c1)}${m.r1 + 1}"/>`)
        .join('')}</mergeCells>`
    : '';

  const paneXml = freezeHeader
    ? '<pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/>' +
      '<selection pane="bottomLeft" activeCell="A2" sqref="A2"/>'
    : '';

  const sheetXml =
    `${XML_HEAD}\n<worksheet xmlns="${NS_MAIN}" xmlns:r="${NS_REL}">` +
    `<dimension ref="A1:${colLetter(grid.colCount - 1)}${grid.rowCount}"/>` +
    `<sheetViews><sheetView showGridLines="0" tabSelected="1" workbookViewId="0">${paneXml}</sheetView></sheetViews>` +
    '<sheetFormatPr defaultRowHeight="16"/>' +
    (colsXml ? `<cols>${colsXml}</cols>` : '') +
    `<sheetData>${rowXml.join('')}</sheetData>` +
    mergesXml +
    '<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>' +
    '</worksheet>';

  const stylesXml =
    `${XML_HEAD}\n<styleSheet xmlns="${NS_MAIN}">` +
    (numFmts.length
      ? `<numFmts count="${numFmts.length}">${numFmts
          .map((n, i) => `<numFmt numFmtId="${164 + i}" formatCode="${esc(n.code)}"/>`)
          .join('')}</numFmts>`
      : '') +
    `<fonts count="${fonts.length}">${fonts.map((f) => f.xml).join('')}</fonts>` +
    `<fills count="${fills.length}">${fills.map((f) => f.xml).join('')}</fills>` +
    `<borders count="${borders.length}">${borders.map((b) => b.xml).join('')}</borders>` +
    '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>' +
    `<cellXfs count="${xfs.length}">${xfs.map((x) => x.xml).join('')}</cellXfs>` +
    '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>' +
    '<dxfs count="0"/><tableStyles count="0" defaultTableStyle="TableStyleMedium2" defaultPivotStyle="PivotStyleLight16"/>' +
    '</styleSheet>';

  const safeName = esc(String(sheetName).replace(/[[\]:*?/\\]/g, '').slice(0, 31) || 'Sheet1');

  const contentTypes =
    `${XML_HEAD}\n<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
    '<Default Extension="xml" ContentType="application/xml"/>' +
    '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>' +
    '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>' +
    '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>' +
    '</Types>';

  const rootRels =
    `${XML_HEAD}\n<Relationships xmlns="${NS_PKG_REL}">` +
    `<Relationship Id="rId1" Type="${NS_REL}/officeDocument" Target="xl/workbook.xml"/>` +
    '</Relationships>';

  const workbook =
    `${XML_HEAD}\n<workbook xmlns="${NS_MAIN}" xmlns:r="${NS_REL}">` +
    '<bookViews><workbookView/></bookViews>' +
    `<sheets><sheet name="${safeName}" sheetId="1" r:id="rId1"/></sheets>` +
    '</workbook>';

  const workbookRels =
    `${XML_HEAD}\n<Relationships xmlns="${NS_PKG_REL}">` +
    `<Relationship Id="rId1" Type="${NS_REL}/worksheet" Target="worksheets/sheet1.xml"/>` +
    `<Relationship Id="rId2" Type="${NS_REL}/styles" Target="styles.xml"/>` +
    '</Relationships>';

  const zip = new JSZip();
  zip.file('[Content_Types].xml', contentTypes);
  zip.file('_rels/.rels', rootRels);
  zip.file('xl/workbook.xml', workbook);
  zip.file('xl/_rels/workbook.xml.rels', workbookRels);
  zip.file('xl/styles.xml', stylesXml);
  zip.file('xl/worksheets/sheet1.xml', sheetXml);

  const buf = await zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' });
  return buf;
}
