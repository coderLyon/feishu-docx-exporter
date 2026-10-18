/** 图片 / 附件 / 内嵌电子表格 / 画板 渲染 */
import {
  AlignmentType,
  BorderStyle,
  ExternalHyperlink,
  ImageRun,
  Paragraph,
  ShadingType,
  Table,
  TableCell,
  TableRow,
  TableLayoutType,
  TextRun,
  VerticalAlign,
  WidthType,
} from 'docx';
import { PX_TO_EMU, PX_TO_TWIP } from '../../shared/constants.js';
import { SHEET_MARKER } from '../ole/inject.js';
import { t } from '../../shared/messages.js';
import { sanitizeXmlText } from '../../shared/util.js';
import { STYLE } from '../docx/styles.js';

/** 按页面宽度等比缩放 */
function fitImage(width, height, maxWidthPx) {
  if (!width || !height) return { width: maxWidthPx, height: Math.round(maxWidthPx * 0.6) };
  const scale = Math.min(1, maxWidthPx / width);
  return { width: Math.round(width * scale), height: Math.round(height * scale) };
}

function imageParagraph(resource, ctx, caption, { allowLandscape = false } = {}) {
  // 只有「宽表格截图」才配横页；普通插图即便很宽也不该把整篇文档切成横的
  // （实测：无条件判定会把 845 块文档切成 11 节、5 个横向页）
  const contentPx = (ctx.pageContentWidth ?? 10106) / 15;
  const isWide = allowLandscape && (resource?.width ?? 0) > contentPx * 1.15;
  const markWide = (el) => {
    if (isWide && el) el.__fdeLandscape = true;
    return el;
  };
  if (!resource?.data) return [];

  // 在表格单元格里按列宽缩放（再留一点单元格内边距），否则会撑破表格
  // 宽表进横页后要按**横向版心**缩放，否则只占纸面一半（实测横页里表格只有 70% 宽）
  const contentWidthPx = isWide && ctx.landscapeContentWidth
    ? Math.round(ctx.landscapeContentWidth / 15)
    : ctx.contentWidthPx;
  const available = ctx.cellWidthPx
    ? Math.max(60, Math.min(contentWidthPx, ctx.cellWidthPx - 12))
    : contentWidthPx;
  const { width, height } = fitImage(resource.width, resource.height, available);
  const run = new ImageRun({
    type: resource.mimeType?.includes('jpeg') ? 'jpg' : 'png',
    data: resource.data,
    transformation: { width, height },
  });

  // 飞书图片块默认居中
  const result = [markWide(new Paragraph({ children: [run], alignment: AlignmentType.CENTER }))];
  if (caption) {
    result.push(
      new Paragraph({
        // 题注不写 size/color：由 FdeCaption 样式承载（灰、小、斜体、居中）。
        // alignment 必须用 AlignmentType 字符串枚举，写数字 1 会产出非法的
        // <w:jc w:val="1"/>，Word 打开时直接判定文档损坏
        children: [new TextRun({ text: caption })],
        style: STYLE.caption,
        alignment: AlignmentType.CENTER,
      }),
    );
  }
  return result;
}

/** 图片块 */
export function renderImage(node, ctx) {
  const captionRuns = node.data?.image?.caption?.text
    ? node.data.image.caption.text
    : null;
  const caption = captionRuns
    ? (captionRuns.initialAttributedTexts?.text?.['0'] ?? '').trim()
    : '';

  const resource = ctx.takeResource('image');
  if (!resource) {
    const token = node.data?.image?.token ?? '';
    ctx.report.degraded.push(t('imageMissing', ctx.lang)(token));
    return imageParagraph(null, ctx);
  }
  return imageParagraph(resource, ctx, caption || undefined);
}

/** 附件块：以链接占位 */
export function renderFile(node, ctx) {
  const file = node.data?.file ?? {};
  const name = sanitizeXmlText(file.name ?? '附件');
  const sizeKb = file.size ? ` (${Math.round(file.size / 1024)} KB)` : '';
  const url = file.token ? `https://${ctx.host}/space/api/box/stream/download/all/${file.token}` : '';

  // 取到附件字节 → 输出占位段，打包后由 ole/inject.js 换成真正的 OLE 文件对象。
  // 双击即可用关联程序打开，而不是跳转浏览器链接。
  const attachment = ctx.attachments?.get?.(file.token);
  if (ctx.options?.embedAttachments && attachment?.bytes?.length) {
    const index = (ctx.oleSeq = (ctx.oleSeq ?? 0) + 1);
    ctx.oleObjects = ctx.oleObjects ?? [];
    ctx.oleObjects.push({ index, name, bytes: attachment.bytes });
    return [new Paragraph({ children: [new TextRun({ text: `\u0001FDE_OLE_${index}\u0001` })], style: STYLE.body })];
  }

  ctx.report.degraded.push(t('fileAsLink', ctx.lang)(name));

  return [
    new Paragraph({
      children: [
        new TextRun({ text: t('fileLabel', ctx.lang)() }),
        url
          ? new ExternalHyperlink({
              link: url,
              children: [new TextRun({ text: `${name}${sizeKb}`, style: 'Hyperlink' })],
            })
          : new TextRun({ text: `${name}${sizeKb}`, bold: true }),
      ],
      style: STYLE.body,
    }),
  ];
}

/** 内嵌电子表格：整表截图嵌入，彻底避免官方导出「显示不全」 */
/** rgb(216, 234, 211) / #d8ead3 → 'D8EAD3' */
function cssColorToHex(value) {
  if (!value) return null;
  const rgb = value.match(/rgba?\(\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)/i);
  if (rgb) {
    const [r, g, b] = [Number(rgb[1]), Number(rgb[2]), Number(rgb[3])];
    return [r, g, b].map((n) => n.toString(16).padStart(2, '0')).join('').toUpperCase();
  }
  const hex = value.trim().replace(/^#/, '');
  return /^[0-9a-f]{6}$/i.test(hex) ? hex.toUpperCase() : null;
}

/**
 * 用飞书「复制」得到的 **HTML flavor** 还原表格完整样式。
 *
 * 飞书给出的单元格形如：
 *   <td rowspan="4" style="border: 0.5pt solid rgb(222,224,227); font-size: 10pt;
 *        background-color: rgb(216,234,211); color: rgb(17,34,51)">模块</td>
 *
 * 因此底色、字色、字号、加粗、对齐、行列合并、列宽都能 1:1 搬过来，
 * 而不是只搬文字（那正是「样式相去甚远」的原因）。
 */
/**
 * 把文本按「emoji/符号」与「普通文字」切开分别建 run。
 * 原因：✅❌🔜 这类字符不在微软雅黑里，Word 回退后字形会错（实测变成 ☑/⊠/↔），
 * 指定 Segoe UI Emoji 才能正确显示。
 */
const EMOJI_RE = /[\u2190-\u21FF\u2600-\u27BF\u2B00-\u2BFF\uFE0F\u{1F000}-\u{1FAFF}]/u;
function splitEmojiRuns(text, cell) {
  const size = Math.round(cell.fontSizePt * 2);
  const base = { bold: cell.bold, size, ...(cell.color ? { color: cell.color } : {}) };
  if (!EMOJI_RE.test(text)) return [new TextRun({ text, ...base })];

  const runs = [];
  let buf = '';
  for (const ch of text) {
    if (EMOJI_RE.test(ch)) {
      if (buf) {
        runs.push(new TextRun({ text: buf, ...base }));
        buf = '';
      }
      runs.push(new TextRun({ text: ch, ...base, font: 'Segoe UI Emoji' }));
    } else {
      buf += ch;
    }
  }
  if (buf) runs.push(new TextRun({ text: buf, ...base }));
  return runs.length ? runs : [new TextRun({ text, ...base })];
}

function buildSheetFromHtml(html, ctx) {
  const parsed = new DOMParser().parseFromString(html, 'text/html');
  const tableEl = parsed.querySelector('table');
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
  if (!colCount || colCount > 40) return null;

  // 列数多（≥7）的表格在竖版里每列只有 50~60px，必然折行 —— 放进横向节。
  // 横向页版心 = 长边 - 左右边距，比竖版宽 40% 左右。
  const wide = colCount >= 7;
  const availableTw = wide
    ? (ctx.landscapeContentWidth ?? ctx.pageContentWidth ?? 10106)
    : (ctx.pageContentWidth ?? 10106);
  let widthsPx = colEls.map((c) => Number(c.getAttribute('width')) || 0).filter((w) => w > 0);
  if (widthsPx.length !== colCount) {
    widthsPx = Array.from({ length: colCount }, () => Math.round(availableTw / colCount / PX_TO_TWIP));
  }

  const styleOf = (td) => {
    const style = td.getAttribute('style') ?? '';
    const pick = (re) => (style.match(re)?.[1] ?? '').trim();
    const bg = cssColorToHex(pick(/background-color:\s*([^;]+)/i));
    const color = cssColorToHex(pick(/[^-]color:\s*([^;]+)/i));
    const fontSizePt = Number(pick(/font-size:\s*([\d.]+)pt/i)) || 10;
    const bold = /font-weight:\s*(bold|[6-9]00)/i.test(style);
    const alignRaw = pick(/text-align:\s*([a-z]+)/i);
    return { bg, color, fontSizePt, bold, alignRaw };
  };

  // HTML 表格 → 规整网格：用占位矩阵处理 rowspan/colspan，
  // 否则合并单元格会让每行 td 数量不一致，Word 里列就错位（实测 11 列变 8 列）。
  const occupied = Array.from({ length: trEls.length }, () => new Array(colCount).fill(false));
  const grid = Array.from({ length: trEls.length }, () => new Array(colCount).fill(null));
  // owner[ri][ci] = 覆盖该位置的「起始单元格」；用于区分「被上方 rowspan 覆盖」与「本来就是空」
  const owner = Array.from({ length: trEls.length }, () => new Array(colCount).fill(null));

  trEls.forEach((tr, ri) => {
    let ci = 0;
    for (const td of [...tr.children]) {
      if (td.tagName !== 'TD' && td.tagName !== 'TH') continue;
      while (ci < colCount && occupied[ri][ci]) ci += 1;
      if (ci >= colCount) break;
      const colspan = Math.min(Number(td.getAttribute('colspan')) || 1, colCount - ci);
      const rowspan = Math.min(Number(td.getAttribute('rowspan')) || 1, trEls.length - ri);
      const st = styleOf(td);
      const text = (td.textContent ?? '').replace(/\u200b/g, '').trim();
      grid[ri][ci] = { ...st, text, colspan, rowspan };
      for (let r = ri; r < ri + rowspan; r += 1) {
        for (let c = ci; c < ci + colspan; c += 1) {
          if (!occupied[r]) continue;
          occupied[r][c] = true;
          if (!owner[r][c]) owner[r][c] = grid[ri][ci];
        }
      }
      ci += colspan;
    }
  });

  const rows = [];
  for (let ri = 0; ri < grid.length; ri += 1) {
    const cells = [];
    for (let ci = 0; ci < colCount; ci += 1) {
      const cell = grid[ri][ci];
      // 被上方 rowspan 覆盖：**不再生成单元格**（docx 的 rowSpan 会自动补 vMerge continue，
      // 这里再补一个就会多出一列 —— 实测 11 列变 12 列）
      if (!cell && owner[ri][ci]) continue;
      if (!cell) {
        cells.push(new TableCell({ children: [new Paragraph({ style: STYLE.sheetCell, children: [] })] }));
        continue;
      }
      const align =
        cell.alignRaw === 'center' ? AlignmentType.CENTER : cell.alignRaw === 'right' ? AlignmentType.RIGHT : undefined;
      cells.push(
        new TableCell({
          ...(cell.colspan > 1 ? { columnSpan: cell.colspan } : {}),
          ...(cell.rowspan > 1 ? { rowSpan: cell.rowspan } : {}),
          verticalAlign: VerticalAlign.CENTER,
          margins: { top: 30, bottom: 30, left: 60, right: 60 },
          ...(cell.bg ? { shading: { type: ShadingType.CLEAR, fill: cell.bg } } : {}),
          children: [
            new Paragraph({
              alignment: align,
              style: STYLE.sheetCell,
              children: splitEmojiRuns(cell.text, cell),
            }),
          ],
        }),
      );
    }
    if (cells.length) rows.push(new TableRow({ cantSplit: true, children: cells }));
  }
  if (!rows.length) return null;

  // 列宽：**直接用 colgroup 的比例**（这正是飞书自己的列宽，等于「样式一致」），
  // 只做等比映射 + 极小值保护。不要套普通表格的「四字下限」——
  // 11 列时下限之和会超过版心，把窄列压到 8px（实测过）。
  const srcSum = widthsPx.reduce((a, b) => a + b, 0) || 1;
  const colWidths = widthsPx.map((w) => Math.round((w / srcSum) * availableTw));

  // 最小宽度：至少放得下 5 个字（表头「Windows」「模块」不被拆成两行）。
  // 差额从最宽的列按比例扣，保持整体比例和 = 版心。
  // 下限要按列数自适应：11 列时「5 个字」的下限之和会超过版心，把所有列拉平。
  // 取 min(5 字宽, 均分宽 × 0.85)
  const fiveChar = Math.round(5 * 10 * (96 / 72) * PX_TO_TWIP);
  const MIN_W = Math.min(fiveChar, Math.floor((availableTw / colCount) * 0.85));
  for (let i = 0; i < colWidths.length; i += 1) {
    if (colWidths[i] < MIN_W) {
      const need = MIN_W - colWidths[i];
      let pool = 0;
      for (let j = 0; j < colWidths.length; j += 1) pool += Math.max(0, colWidths[j] - MIN_W);
      if (pool > 0) {
        for (let j = 0; j < colWidths.length; j += 1) {
          const room = Math.max(0, colWidths[j] - MIN_W);
          colWidths[j] -= Math.floor((need * room) / pool);
        }
      }
      colWidths[i] = MIN_W;
    }
  }

  const border = { style: BorderStyle.SINGLE, size: 4, color: 'DEE0E3' };
  const table = new Table({
    rows,
    columnWidths: colWidths,
    layout: TableLayoutType.FIXED,
    width: { size: 100, type: WidthType.PERCENTAGE },
    borders: {
      top: border,
      bottom: border,
      left: border,
      right: border,
      insideHorizontal: border,
      insideVertical: border,
    },
  });
  if (wide) {
    // 交给 build.js：把这个表格单独放进横向节
    table.__fdeLandscape = true;
  }
  return [table, new Paragraph({ children: [], style: STYLE.body })];
}

/**
 * 把 TSV 单元格数据建成 Word 原生表格。
 * 返回值：段落/表格数组；数据不合法时返回 null（调用方退回截图）。
 */
function buildSheetTable(tsv, ctx) {
  const rows = String(tsv)
    .replace(/\r/g, '')
    .split('\n')
    .filter((line) => line.length)
    .map((line) => line.split('\t'));
  if (!rows.length) return null;

  const colCount = Math.max(...rows.map((r) => r.length));
  if (colCount < 1 || colCount > 40 || rows.length > 400) return null;

  // 列宽：与普通表格同一套「尽量不折行」算法；连不可断 token 都放不下就回退截图
  const columnTexts = Array.from({ length: colCount }, (_, ci) => rows.map((r) => (r[ci] ?? '').trim()));
  const available = ctx.pageContentWidth ?? 10106;
  const widths = null; // 不再手算列宽，交给 Word AutoFit
  const sumToken = 0;
  if (sumToken > available * 1.02) {
    ctx.report.degraded.push(t('sheetTooWide', ctx.lang)(colCount));
    return null;
  }

  const border = { style: BorderStyle.SINGLE, size: 4, color: '000000' };
  const table = new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    borders: { top: border, bottom: border, left: border, right: border, insideHorizontal: border, insideVertical: border },
    // 交给 Word 自动调整列宽（同上，手算不如 AutoFit）
    rows: rows.map(
      (cells, rowIndex) =>
        new TableRow({
          cantSplit: true,
          ...(rowIndex === 0 ? { tableHeader: true } : {}),
          children: (() => {
            // 首行只有一个非空单元格 → 视为跨列标题（飞书表格常见），合并成一格
            const filled = cells.filter((c2) => (c2 ?? '').trim()).length;
            if (rowIndex === 0 && colCount > 1 && filled === 1) {
              const title = (cells.find((c2) => (c2 ?? '').trim()) ?? '').trim();
              return [
                new TableCell({
                  columnSpan: colCount,
                  margins: { top: 40, bottom: 40, left: 80, right: 80 },
                  shading: { type: ShadingType.CLEAR, fill: 'EFF0F1' },
                  children: [
                    new Paragraph({
                      children: [new TextRun({ text: title, bold: true })],
                      style: STYLE.tableText,
                    }),
                  ],
                }),
              ];
            }
            return Array.from({ length: colCount }, (_, ci) => {
              const text = (cells[ci] ?? '').trim();
              return new TableCell({
              margins: { top: 40, bottom: 40, left: 80, right: 80 },
              ...(rowIndex === 0 ? { shading: { type: ShadingType.CLEAR, fill: 'EFF0F1' } } : {}),
                children: [
                  new Paragraph({
                    children: [new TextRun({ text })],
                    style: STYLE.tableText,
                  }),
                ],
              });
            });
          })(),
        }),
    ),
  });
  return [table, new Paragraph({ children: [], style: STYLE.body })];
}

export function renderSheet(node, ctx) {
  const token = node.data?.token ?? '';
  const resource = ctx.takeResource('sheet');

  // 用户可在面板选择内嵌表格的导出方式：
  //   'image'（默认）—— 截图，版式最忠实
  //   'table'        —— Word 原生表格，可编辑可搜索（需要先提取到单元格数据）
  //   'excel'       —— Word 里内嵌真 Excel 对象（预览图为整表截图，双击进 Excel）
  if (ctx.options?.embeddedSheet === 'excel') {
    const html = typeof resource?.table === 'object' ? resource.table.html : null;
    if (html) {
      // 渲染是同步的、xlsx 生成是异步的 → 这里只登记任务，由 export.js 打包前统一生成
      const index = (ctx.xlsxSeq = (ctx.xlsxSeq ?? 0) + 1);
      ctx.xlsxObjects = ctx.xlsxObjects ?? [];
      const contentPt = (ctx.pageContentWidth ?? 10106) / 20;
      const wide = (ctx.landscapeContentWidth ?? 0) / 20;
      const pxW = resource?.width || 800;
      const pxH = resource?.height || 400;
      // 宽表按横向版心算，避免在横向页里只占一半
      const maxW = pxW / pxH > 1.6 && wide ? wide : contentPt;
      let widthPt = Math.min(maxW, contentPt * 1.6, wide || contentPt * 1.6);
      let heightPt = (widthPt * pxH) / pxW;
      if (heightPt > 620) {
        widthPt *= 620 / heightPt;
        heightPt = 620;
      }
      ctx.xlsxObjects.push({
        index,
        html,
        name: token,
        preview: resource?.data ?? null,
        widthPt,
        heightPt,
        wide,
      });
      const placeholder = new Paragraph({
        children: [new TextRun({ text: `${SHEET_MARKER}${index}\u0001` })],
        alignment: AlignmentType.CENTER,
        style: STYLE.body,
      });
      // 预览图比竖版版心宽 → 这一节转成横向页，否则会溢出到纸外
      if (widthPt > contentPt + 1) placeholder.__fdeLandscape = true;
      return [placeholder];
    }
    if (!resource?.data) ctx.report.degraded.push(t('sheetNoCellData', ctx.lang)());
  }

  if (ctx.options?.embeddedSheet === 'table') {
    // 优先用带样式的 HTML（底色/字色/字号/合并/列宽），退回纯文本 TSV
    const html = typeof resource?.table === 'object' ? resource.table.html : null;
    if (html) {
      const built = buildSheetFromHtml(html, ctx);
      if (built) return built;
    }
    if (resource?.table) {
      const tsv = typeof resource.table === 'object' ? resource.table.tsv : resource.table;
      const built = buildSheetTable(tsv, ctx);
      if (built) return built;
      // 数据拿到了但版心放不下/格式异常 → 继续走图片，下面会补一条降级说明
    } else if (resource?.data) {
      ctx.report.degraded.push(t('sheetNoCellData', ctx.lang)());
    }
  }

  if (resource?.data) {
    return imageParagraph(resource, ctx, null, { allowLandscape: true });
  }

  ctx.report.degraded.push(t('sheetAsLink', ctx.lang)());
  return [
    new Paragraph({
      children: [
        new TextRun({ text: t('sheetLabel', ctx.lang)(), bold: true }),
        token
          ? new ExternalHyperlink({
              link: `https://${ctx.host}/docx/${ctx.objToken ?? ''}?sheet=${token}`,
              children: [new TextRun({ text: token, style: 'Hyperlink' })],
            })
          : new TextRun({ text: t('noToken', ctx.lang)() }),
      ],
      style: STYLE.body,
    }),
  ];
}

/** 画板 / 流程图 / 思维导图 */
export function renderWhiteboard(node, ctx) {
  const resource = ctx.takeResource('whiteboard');
  if (resource?.data) return imageParagraph(resource, ctx);
  ctx.report.degraded.push(t('whiteboardMissing', ctx.lang)());
  return [
    new Paragraph({
      children: [new TextRun({ text: t('whiteboardFallback', ctx.lang)(), italics: true, color: 'FF8800' })],
      style: STYLE.body,
    }),
  ];
}

/** 多维表格：以链接占位 */
export function renderBitable(node, ctx) {
  // 有预取数据就渲染成 Word 原生表格（表头灰底、跨页重复）；没有则退回链接卡片
  const data = ctx.bitables?.get?.(node.data?.token ?? '');
  if (data?.headers?.length) {
    const cols = data.headers.length;
    const wide = cols >= 7;
    const availableTw = wide
      ? (ctx.landscapeContentWidth ?? ctx.pageContentWidth ?? 10106)
      : (ctx.pageContentWidth ?? 10106);
    const cell = (text, header) =>
      new TableCell({
        children: [
          new Paragraph({
            style: STYLE.sheetCell,
            children: [new TextRun({ text: String(text ?? ''), bold: header })],
          }),
        ],
        shading: header ? { fill: 'F5F6F7' } : undefined,
        verticalAlign: VerticalAlign.CENTER,
      });
    const table = new Table({
      rows: [
        new TableRow({
          tableHeader: true,
          children: data.headers.map((h) => cell(h, true)),
        }),
        ...data.rows.map(
          (row) =>
            new TableRow({
              children: row.map((v) => cell(v, false)),
            }),
        ),
      ],
      width: { size: availableTw, type: WidthType.DXA },
      columnWidths: Array.from({ length: cols }, () => Math.floor(availableTw / cols)),
      layout: TableLayoutType.FIXED,
      borders: {
        top: { style: BorderStyle.SINGLE, size: 4, color: 'DEE0E3' },
        bottom: { style: BorderStyle.SINGLE, size: 4, color: 'DEE0E3' },
        left: { style: BorderStyle.SINGLE, size: 4, color: 'DEE0E3' },
        right: { style: BorderStyle.SINGLE, size: 4, color: 'DEE0E3' },
        insideHorizontal: { style: BorderStyle.SINGLE, size: 4, color: 'DEE0E3' },
        insideVertical: { style: BorderStyle.SINGLE, size: 4, color: 'DEE0E3' },
      },
    });
    if (wide) table.__fdeLandscape = true;
    return [table];
  }

  // bitable / base_refer 的 token 形如 <baseToken>_<tableId>，
  // 第一个下划线之前是 base token；有 view_id 时一并带上，点开能落到相同视图
  const raw = node.data?.token ?? '';
  const token = raw.split('_')[0];
  const tableId = raw.includes('_') ? raw.split('_').slice(1).join('_') : '';
  const viewId = node.data?.view_id ?? '';
  const query = [tableId ? `table=${tableId}` : '', viewId ? `view=${viewId}` : '']
    .filter(Boolean)
    .join('&');
  const url = token ? `https://${ctx.host}/base/${token}${query ? `?${query}` : ''}` : '';

  ctx.report.degraded.push(t('bitableAsLink', ctx.lang)());
  return [
    new Paragraph({
      children: [
        new TextRun({ text: t('bitableLabel', ctx.lang)(), bold: true }),
        token
          ? new ExternalHyperlink({
              link: url,
              children: [new TextRun({ text: token, style: 'Hyperlink' })],
            })
          : new TextRun({ text: t('noToken', ctx.lang)() }),
      ],
      style: STYLE.body,
    }),
  ];
}

export { PX_TO_EMU };
