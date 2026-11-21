/** 导出编排：解析 -> 取数 -> 渲染 -> 打包 */
import { fetchBitableClientVars, loadDocument } from './api.js';
import { t } from '../shared/messages.js';
import { buildBlockTree } from './tree.js';
import { createRenderContext, renderNode } from './render/index.js';
import { buildDocx } from './docx/build.js';
import { injectOleObjects, injectXlsxObjects } from './ole/inject.js';
import { buildXlsx } from './ole/xlsx.js';
import { parseSheetHtml } from './sheet/grid.js';
import { FONT_SCHEMES, PAGE_MARGINS, PAGE_SIZES, TYPO } from '../shared/constants.js';
import { runPool } from '../shared/util.js';

const MAX_ATTACHMENT_BYTES = 200 * 1024 * 1024;

/**
 * @param {Object} params
 * @param {string} params.url 当前文档地址
 * @param {Object} params.options 导出选项
 * @param {Object} [params.resources] 页面采集到的资源 { images, sheets, whiteboards }
 * @param {Function} [params.onProgress]
 * @returns {Promise<{blob: Blob, title: string, report: Object}>}
 */
/**
 * 抽样文档正文（前若干字符）用于语言检测。
 * 只读块数据里的纯文本，不触发渲染。
 */
function sampleDocText(doc, limit = 3000) {
  const parts = [];
  let len = 0;
  for (const block of Object.values(doc?.blockMap ?? {})) {
    const text = block?.data?.text?.initialAttributedTexts?.text?.['0'];
    if (typeof text !== 'string' || !text) continue;
    parts.push(text);
    len += text.length;
    if (len >= limit) break;
  }
  return parts.join('\n').slice(0, limit);
}


function resourceBlockIds(root, byId) {
  const expected = { image: [], sheet: [], whiteboard: [] };
  const seen = new Set();
  function walk(node) {
    if (!node || seen.has(node.id)) return;
    seen.add(node.id);
    const kind = ['diagram', 'mindnote'].includes(node.type) ? 'whiteboard' : node.type;
    if (Object.hasOwn(expected, kind)) expected[kind].push(node.id);

    if (node.type === 'table') {
      for (const row of node.data.rows_id ?? []) {
        for (const column of node.data.columns_id ?? []) {
          walk(byId.get(node.data.cell_set?.[row + column]?.block_id));
        }
      }
    } else {
      node.children.forEach(walk);
    }
  }
  walk(root);
  return expected;
}

async function downloadAttachments(byId, attachments, report) {
  const files = new Map();
  for (const node of byId.values()) {
    const file = node.data?.file;
    if (node.type === 'file' && file?.token && (file.size ?? 0) <= MAX_ATTACHMENT_BYTES) {
      files.set(file.token, file);
    }
  }

  await runPool([...files.values()], 3, async (file) => {
    try {
      const response = await fetch(`/space/api/box/stream/download/all/${file.token}`, { credentials: 'include' });
      if (!response.ok) {
        report.degraded.push(`· 附件 ${file.name} 下载返回 ${response.status}`);
        return;
      }
      const blob = await response.blob();
      if (blob.size && blob.size <= MAX_ATTACHMENT_BYTES) {
        attachments.set(file.token, { bytes: new Uint8Array(await blob.arrayBuffer()) });
      } else {
        report.degraded.push(`· 附件 ${file.name} 超过 ${MAX_ATTACHMENT_BYTES / 1048576}MB 或为空，已改为链接`);
      }
    } catch (error) {
      report.degraded.push(`· 附件 ${file.name} 取字节失败：${error?.message ?? error}`);
    }
  });
}

export async function exportToDocx({ url, options, resources, collect, onProgress }) {
  onProgress?.({ stage: 'resolve', percent: 5 });

  const { node, doc } = await loadDocument({ url, onProgress });
  const { root, byId } = buildBlockTree(doc);
  const expected = resourceBlockIds(root, byId);
  const report = { degraded: [], errors: [] };
  const attachments = new Map();
  const [collected] = await Promise.all([
    collect ? collect(expected) : resources,
    downloadAttachments(byId, attachments, report),
  ]);
  resources = collected;
  onProgress?.({ stage: 'render', percent: 45, detail: node.title });
  if (resources?.liveSheetError) report.degraded.push(`· 表格读取失败：${resources.liveSheetError}`);

  const size = PAGE_SIZES[options.pageSize] ?? PAGE_SIZES.A4;
  const margin = PAGE_MARGINS[options.pageMargin] ?? PAGE_MARGINS.normal;
  const pageContentWidth = size.width - margin.left - margin.right;
  // 横向页的版心宽度：长边（A4 的 height）减去左右边距。
  // 宽表会进横向节，列宽要按这个宽度算，否则横向页里表格只占一半。
  const landscapeContentWidth = size.height - margin.left - margin.right;
  const font = FONT_SCHEMES[options.fontScheme] ?? FONT_SCHEMES.yahei;

  // 多维表格数据（在渲染前取好：渲染是同步的，接口是异步的）
  const bitables = new Map();

  const ctx = createRenderContext({
    byId,
    // 表格单元格块不在 blockSequence 里（只存于 blockMap），列宽估算需要它兜底
    blockMap: doc.blockMap ?? {},
    attachments,
    bitables,
    options,
    resources,
    host: new URL(url ?? location.href).hostname,
    objToken: node.objToken,
    report,
    pageContentWidth,
    landscapeContentWidth,
    bodySize: TYPO.body.size,
    bodyFont: { ascii: font.ascii, eastAsia: font.eastAsia, hAnsi: font.ascii },
    // 抽样正文用于判断文档语言（决定占位/降级文案用中文还是英文）
    textSample: sampleDocText(doc),
  });

  // 多维表格：预取视图数据，失败就保持原来的链接卡片
  if (options.embedBitable !== false) {
    const nodes = [...byId.values()].filter((n) => n.type === 'bitable' && n.data?.token).slice(0, 5);
    for (const n of nodes) {
      try {
        const data = await fetchBitableClientVars({ token: n.data.token, viewId: n.data.view_id ?? '' });
        if (data) bitables.set(n.data.token, data);
      } catch (error) {
        report.degraded.push(`· 多维表格读取失败：${error?.message ?? error}`);
      }
    }
  }

  const children = root ? renderNode(root, ctx, 0) : [];

  const got = Object.fromEntries(Object.keys(expected).map((kind) => [kind,
    resources?.byBlock?.[kind]?.size ?? resources?.[`${kind}s`]?.length ?? 0,
  ]));
  const missing = Object.keys(expected).filter((kind) => expected[kind].length > got[kind]);
  if (missing.length) {
    const label = t('resourceHint', ctx.lang)(missing.map((k) => `${k} ${got[k]}/${expected[k].length}`).join('、'));
    report.degraded.push(label);
  }

  onProgress?.({ stage: 'pack', percent: 80, detail: '正在打包 docx' });
  const blob = await buildDocx({
    title: node.title,
    children,
    options,
  });

  let finalBlob = blob;

  // 内嵌表格 → Excel 对象：渲染期只登记任务，这里才异步生成 xlsx
  if (ctx.xlsxObjects?.length) {
    onProgress?.({ stage: 'pack', percent: 90, detail: `正在生成 Excel 对象（${ctx.xlsxObjects.length} 个）` });
    const sheets = [];
    for (const task of ctx.xlsxObjects) {
      try {
        const grid = parseSheetHtml(task.html, { maxRows: 500, maxCols: 100 });
        if (!grid) {
          report.degraded.push('· 内嵌表格单元格数据解析失败，已保留截图');
          continue;
        }
        const bytes = await buildXlsx(grid, {
          sheetName: task.name || 'Sheet1',
          fontName: font.eastAsia || '微软雅黑',
          freezeHeader: false,
        });
        sheets.push({ index: task.index, bytes, name: task.name,
          widthPt: task.widthPt, heightPt: task.heightPt });
      } catch (err) {
        report.degraded.push(`· Excel 对象生成失败：${err?.message ?? err}`);
      }
    }
    if (sheets.length) {
      try {
        const embedded = await injectXlsxObjects(finalBlob, sheets);
        finalBlob = embedded.blob;
        report.excelObjects = embedded.count;
        if (embedded.count < sheets.length) report.degraded.push('· 部分 Excel 对象未能嵌入，已保留截图');
      } catch (err) {
        report.degraded.push(`· Excel 对象注入失败：${err?.message ?? err}`);
      }
    }
  }

  if (ctx.oleObjects?.length) {
    onProgress?.({ stage: 'pack', percent: 95, detail: `正在嵌入附件（${ctx.oleObjects.length} 个）` });
    try {
      finalBlob = await injectOleObjects(finalBlob, ctx.oleObjects);
    } catch (err) {
      report.degraded.push(`· 附件嵌入失败：${err?.message ?? err}`);
    }
  }

  onProgress?.({ stage: 'done', percent: 100 });
  return {
    blob: finalBlob,
    title: node.title || '未命名文档',
    report,
    stats: {
      blocks: doc.blockSequence?.length ?? 0,
      images: got.image,
      sheets: got.sheet,
      whiteboards: got.whiteboard,
    },
    liveSheets: resources?.liveSheets ?? null,
  };
}

/** 生成安全文件名 */
export function buildFilename(title) {
  const clean = (title || '未命名文档')
    .replace(/[\\/:*?"<>|]/g, '_')
    .replace(/\s+/g, ' ')
    .trim()
    .slice(0, 80);
  return `${clean || '未命名文档'}.docx`;
}
