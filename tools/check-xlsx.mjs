/**
 * 离线校验 xlsx 生成器：造一个「刁钻」表格，导出 xlsx，打印部件与关键单元格。
 *
 *   node tools/check-xlsx.mjs [输出路径]
 *
 * 覆盖：底色/字色/加粗/合并/列宽/边框 + 货币/百分比/日期/前导零/长编号/普通数字
 */
import fs from 'node:fs';
import path from 'node:path';
import JSZip from 'jszip';
import { buildXlsx } from '../src/core/ole/xlsx.js';

const out = process.argv[2] ?? 'samples/out/check.xlsx';

const cell = (r, c, text, extra = {}) => ({
  r,
  c,
  rowSpan: 1,
  colSpan: 1,
  text,
  bg: null,
  color: null,
  sizePt: 10,
  bold: false,
  align: 'left',
  borderColor: 'DEE0E3',
  formatter: '',
  type: '',
  ...extra,
});

const grid = {
  rowCount: 6,
  colCount: 5,
  colsPx: [57, 260, 238, 106, 177],
  hasBorder: true,
  covered: new Set(['1,0', '2,0']),
  merges: [{ r0: 0, c0: 0, r1: 3, c1: 0 }],
  cells: [
    cell(0, 0, '模块', { rowSpan: 4, bg: 'D8EAD3', bold: true, align: 'center' }),
    cell(0, 1, '金额', { bg: 'D8EAD3', bold: true, align: 'center', formatter: 'General' }),
    cell(0, 2, '占比', { bg: 'D8EAD3', bold: true, align: 'center' }),
    cell(0, 3, '日期', { bg: 'D8EAD3', bold: true, align: 'center' }),
    cell(0, 4, '编号', { bg: 'D8EAD3', bold: true, align: 'center' }),

    cell(1, 1, '¥1,234.50', { formatter: 'Currency' }),
    cell(1, 2, '12.35%', { formatter: 'Percent' }),
    cell(1, 3, '2026/09/21', { formatter: 'DateTime' }),
    cell(1, 4, '00123'),

    cell(2, 1, '(¥99.00)', { formatter: 'Currency', color: 'C02F1D' }),
    cell(2, 2, '0.50%', { formatter: 'Percent' }),
    cell(2, 3, '2026/09/21 13:45:30', { formatter: 'DateTime' }),
    cell(2, 4, '13800138000'),

    cell(3, 1, '1234.5'),
    cell(3, 2, '75%', {}),
    cell(3, 3, '2026年9月21日', {}),
    cell(3, 4, '20260921'),

    cell(4, 0, '合计', { bg: 'E1EAFF', bold: true }),
    cell(4, 1, '¥1,333.50', { formatter: 'Currency', bg: 'E1EAFF' }),
    cell(4, 2, '12.85%', { formatter: 'Percent', bg: 'E1EAFF' }),
    cell(4, 3, '', { bg: 'E1EAFF' }),
    cell(4, 4, '', { bg: 'E1EAFF' }),

    cell(5, 0, '备注', {}),
    cell(5, 1, '这里是较长的说明文本，用来检查 wrapText 与列宽', { colSpan: 4 }),
  ],
};

const bytes = await buildXlsx(grid, { sheetName: 'GO44OU', fontName: '微软雅黑' });
fs.mkdirSync(path.dirname(out), { recursive: true });
fs.writeFileSync(out, Buffer.from(bytes));
console.log(`已生成 ${out}（${bytes.length} 字节，PK 头=${Buffer.from(bytes.slice(0, 2)).toString('hex')}）`);

const zip = await JSZip.loadAsync(bytes);
console.log('部件:', Object.keys(zip.files).join(' '));

const styles = await zip.file('xl/styles.xml').async('string');
const numFmts = [...styles.matchAll(/<numFmt numFmtId="(\d+)" formatCode="([^"]*)"/g)].map((m) => `${m[1]}=${m[2]}`);
const xfCount = (styles.match(/<xf /g) || []).length;
console.log('numFmts:', numFmts.join(' | ') || '(无)');
console.log('fonts/fills/borders/xfs:', (styles.match(/<font>/g) || []).length,
  (styles.match(/<fill>/g) || []).length, (styles.match(/<border>/g) || []).length, xfCount);

const sheet = await zip.file('xl/worksheets/sheet1.xml').async('string');
console.log('dimension:', sheet.match(/<dimension ref="([^"]+)"/)?.[1]);
console.log('mergeCells:', sheet.match(/<mergeCells[^>]*>(.*?)<\/mergeCells>/s)?.[1]?.trim() ?? '(无)');
const cells = [...sheet.matchAll(/<c r="([A-Z]+\d+)" s="(\d+)"[^>]*?(?:\/>|>(?:<v>([^<]*)<\/v>|<is><t[^>]*>([^<]*)<\/t><\/is>)?<\/c>)/g)]
  .slice(0, 14)
  .map((m) => `${m[1]}[s${m[2]}]=${m[3] ?? m[4] ?? ''}`);
console.log('前 14 格:', cells.join('  '));
