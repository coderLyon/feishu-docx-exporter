/**
 * 剪贴板 HTML → 规整网格模型（Word 原生表格与 xlsx 共用）。
 *
 * 飞书复制出来的 HTML 形如：
 *   <table><colgroup><col width="57">…</colgroup>
 *     <tr height="27"><td rowspan="4" style="border:0.5pt solid rgb(222,224,227);
 *       font-size:10pt; background-color:rgb(216,234,211); color:rgb(17,34,51)">模块</td>…
 *
 * 关键点：rowspan/colspan 会让每行 td 数不一致，必须用占位矩阵归一化，
 * 否则列会错位（实测 11 列被挤成 8 列）。
 */

/** rgb(1,2,3) / #abc / #aabbcc → 'RRGGBB'，取不到返回 null */
export function cssColorToHex(value) {
  if (!value) return null;
  const v = String(value).trim().toLowerCase();
  const rgb = v.match(/rgba?\(\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)/);
  if (rgb) {
    return [rgb[1], rgb[2], rgb[3]]
      .map((n) => Math.min(255, Number(n)).toString(16).padStart(2, '0'))
      .join('')
      .toUpperCase();
  }
  const hex = v.match(/^#([0-9a-f]{3}|[0-9a-f]{6})$/);
  if (hex) {
    const h = hex[1];
    return (h.length === 3 ? h.split('').map((c) => c + c).join('') : h).toUpperCase();
  }
  return null;
}

/** 'a:1; b:2' → { a:'1', b:'2' }（比逐条正则更稳，不会把 border-color 当成 color） */
export function parseStyleDecl(style) {
  const out = {};
  for (const part of String(style ?? '').split(';')) {
    const i = part.indexOf(':');
    if (i < 0) continue;
    const key = part.slice(0, i).trim().toLowerCase();
    if (key) out[key] = part.slice(i + 1).trim();
  }
  return out;
}

/**
 * @param {string} html 剪贴板 text/html
 * @param {{maxCols?: number}} [opts]
 * @returns {null | {
 *   rowCount:number, colCount:number, colsPx:number[],
 *   cells:Array<{r:number,c:number,rowSpan:number,colSpan:number,text:string,
 *     bg:string|null,color:string|null,sizePt:number,bold:boolean,
 *     align:'left'|'center'|'right',borderColor:string|null,formatter:string}>,
 *   merges:Array<{r0:number,c0:number,r1:number,c1:number}>,
 *   covered:Set<string>, hasBorder:boolean
 * }}
 */
export function parseSheetHtml(html, { maxCols = 100 } = {}) {
  if (!html || typeof DOMParser === 'undefined') return null;
  const doc = new DOMParser().parseFromString(html, 'text/html');
  const tableEl = doc.querySelector('table');
  if (!tableEl) return null;

  const trEls = [...tableEl.querySelectorAll('tr')];
  if (!trEls.length) return null;

  const colEls = [...tableEl.querySelectorAll('colgroup col')];
  const colCount = Math.max(
    colEls.length,
    ...trEls.map((tr) =>
      [...tr.children].reduce((sum, td) => sum + (Number(td.getAttribute('colspan')) || 1), 0),
    ),
  );
  if (!colCount || colCount > maxCols) return null;

  const colsPx = colEls.map((c) => Number(c.getAttribute('width')) || 0).filter((w) => w > 0);

  const rows = trEls.length;
  const occupied = Array.from({ length: rows }, () => new Array(colCount).fill(false));
  const owner = Array.from({ length: rows }, () => new Array(colCount).fill(null));
  const cells = [];
  const covered = new Set();
  let hasBorder = false;

  trEls.forEach((tr, ri) => {
    let ci = 0;
    for (const td of [...tr.children]) {
      if (td.tagName !== 'TD' && td.tagName !== 'TH') continue;
      while (ci < colCount && occupied[ri][ci]) ci += 1;
      if (ci >= colCount) break;

      const colSpan = Math.min(Number(td.getAttribute('colspan')) || 1, colCount - ci);
      const rowSpan = Math.min(Number(td.getAttribute('rowspan')) || 1, rows - ri);
      const decl = parseStyleDecl(td.getAttribute('style'));

      const bg = cssColorToHex(decl['background-color'] ?? decl.background ?? '');
      const color = cssColorToHex(decl.color ?? '');
      const sizePt = Number(String(decl['font-size'] ?? '').match(/([\d.]+)pt/i)?.[1]) || 10;
      const bold = /^(bold|[6-9]00)$/i.test(String(decl['font-weight'] ?? ''));
      const alignRaw = String(decl['text-align'] ?? '').toLowerCase();
      const align = alignRaw === 'center' ? 'center' : alignRaw === 'right' ? 'right' : 'left';

      const borderMatch = String(decl.border ?? '').match(/solid\s+(rgb\([^)]+\)|#[0-9a-f]{3,6})/i);
      const borderColor = cssColorToHex(borderMatch?.[1] ?? '');
      if (decl.border) hasBorder = true;

      const text = (td.textContent ?? '').replace(/\u200b/g, '').replace(/\u00a0/g, ' ').trim();
      const formatter = td.getAttribute('data-formatter') ?? '';
      const type = td.getAttribute('data-type') ?? '';

      const start = { r: ri, c: ci, rowSpan, colSpan, text, bg, color, sizePt, bold, align, borderColor, formatter, type };
      cells.push(start);

      for (let r = ri; r < ri + rowSpan; r += 1) {
        for (let c = ci; c < ci + colSpan; c += 1) {
          if (!occupied[r]) continue;
          occupied[r][c] = true;
          if (!owner[r][c]) owner[r][c] = start;
          if (r !== ri || c !== ci) covered.add(`${r},${c}`);
        }
      }
      ci += colSpan;
    }
  });

  const merges = cells
    .filter((c) => c.rowSpan > 1 || c.colSpan > 1)
    .map((c) => ({ r0: c.r, c0: c.c, r1: c.r + c.rowSpan - 1, c1: c.c + c.colSpan - 1 }));

  return { rowCount: rows, colCount, colsPx, cells, merges, covered, hasBorder };
}

/** 列字母：0 → A，26 → AA */
export function colLetter(index) {
  let n = index + 1;
  let s = '';
  while (n > 0) {
    const rem = (n - 1) % 26;
    s = String.fromCharCode(65 + rem) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}
