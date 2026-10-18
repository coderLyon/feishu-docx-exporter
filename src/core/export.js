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

export async function exportToDocx({ url, options, resources, onProgress }) {
  onProgress?.({ stage: 'resolve', percent: 5 });

  const { node, doc } = await loadDocument({ url, onProgress });
  onProgress?.({ stage: 'render', percent: 45, detail: node.title });

  const { root, byId } = buildBlockTree(doc);
  const report = { degraded: [], errors: [] };

  const size = PAGE_SIZES[options.pageSize] ?? PAGE_SIZES.A4;
  const margin = PAGE_MARGINS[options.pageMargin] ?? PAGE_MARGINS.normal;
  const pageContentWidth = size.width - margin.left - margin.right;
  // 横向页的版心宽度：长边（A4 的 height）减去左右边距。
  // 宽表会进横向节，列宽要按这个宽度算，否则横向页里表格只占一半。
  const landscapeContentWidth = size.height - margin.left - margin.right;
  const font = FONT_SCHEMES[options.fontScheme] ?? FONT_SCHEMES.yahei;

  // 附件字节容器（稍后填充；Map 引用传给 ctx，填充即可见）
  const attachments = new Map();
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
    host: location.hostname,
    objToken: node.objToken,
    report,
    pageContentWidth,
    landscapeContentWidth,
    bodySize: TYPO.body.size,
    bodyFont: { ascii: font.ascii, eastAsia: font.eastAsia, hAnsi: font.ascii },
    // 抽样正文用于判断文档语言（决定占位/降级文案用中文还是英文）
    textSample: sampleDocText(doc),
  });

  // 附件字节：用于把附件做成 OLE 文件对象（双击直接打开）。
  // 上限 200MB：Ole10Native 的 dataSize 是 u32（<4GB 都合法），真正的瓶颈是内存，
  // 200MB 在 8GB 机器上仍安全（峰值约 2 倍）。超过则回退成链接。
  if (options.embedAttachments) {
    const fileNodes = [...byId.values()].filter((n) => n.type === 'file' && n.data?.file?.token);
    const MAX = 200 * 1024 * 1024;
    let done = 0;
    for (const node of fileNodes) {
      const f = node.data.file;
      if ((f.size ?? 0) > MAX) continue;
      try {
        const res = await fetch(`/space/api/box/stream/download/all/${f.token}`, { credentials: 'include' });
        if (!res.ok) report.degraded.push(`· 附件 ${f.name} 下载返回 ${res.status}`);
        if (res.ok) {
          const buf = await res.arrayBuffer();
          if (buf.byteLength && buf.byteLength <= MAX) attachments.set(f.token, { bytes: new Uint8Array(buf) });
        else report.degraded.push(`· 附件 ${f.name} 字节为空或过大（${buf.byteLength}）`);
        }
      } catch (err) {
        report.degraded.push(`· 附件 ${f.name} 取字节失败：${err?.message ?? err}`);
      }
      done += 1;
      onProgress?.({ stage: 'capture', percent: 60 + Math.round((done / fileNodes.length) * 15), detail: `正在获取附件 ${done}/${fileNodes.length}` });
    }
  }

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

  // 资源完整性自检：拿文档里实际有多少图片/内嵌表格/画板块，与采集到的数量比较。
  // 采集偶尔会漏（懒渲染未完成就翻过去了），用户只看到「降级为链接」却不知道该怎么办，
  // 这里直接给出可执行的建议。
  const countBlocks = (type) =>
    Object.values(doc.blockMap ?? {}).filter((b) => b?.data?.type === type).length;
  const expect = {
    image: countBlocks('image'),
    sheet: countBlocks('sheet'),
    whiteboard: countBlocks('whiteboard'),
  };
  const got = {
    image: resources?.images?.length ?? 0,
    sheet: resources?.sheets?.length ?? 0,
    whiteboard: resources?.whiteboards?.length ?? 0,
  };
  const missing = Object.keys(expect).filter((k) => expect[k] > got[k]);
  if (missing.length) {
    const label = t('resourceHint', ctx.lang)(missing.map((k) => `${k} ${got[k]}/${expect[k]}`).join('、'));
    report.degraded.push(label);
  }

  onProgress?.({ stage: 'pack', percent: 80, detail: '正在打包 docx' });
  const blob = await buildDocx({
    title: node.title,
    children,
    options,
    orderedInstances: ctx.ordered.instance + 1,
  });

  let finalBlob = blob;

  // 内嵌表格 → Excel 对象：渲染期只登记任务，这里才异步生成 xlsx
  if (ctx.xlsxObjects?.length) {
    onProgress?.({ stage: 'pack', percent: 90, detail: `正在生成 Excel 对象（${ctx.xlsxObjects.length} 个）` });
    const sheets = [];
    for (const task of ctx.xlsxObjects) {
      const grid = parseSheetHtml(task.html, { maxCols: 100 });
      if (!grid) {
        report.degraded.push('· 内嵌表格单元格数据解析失败，已回退为截图');
        continue;
      }
      try {
        const bytes = await buildXlsx(grid, {
          sheetName: task.name || 'Sheet1',
          fontName: font.eastAsia || '微软雅黑',
        });
        sheets.push({ index: task.index, bytes, name: task.name, preview: task.preview,
          widthPt: task.widthPt, heightPt: task.heightPt });
      } catch (err) {
        report.degraded.push(`· Excel 对象生成失败：${err?.message ?? err}`);
      }
    }
    if (sheets.length) {
      try {
        finalBlob = await injectXlsxObjects(finalBlob, sheets);
        report.excelObjects = sheets.length;
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
      images: resources?.images?.length ?? 0,
      sheets: resources?.sheets?.length ?? 0,
      whiteboards: resources?.whiteboards?.length ?? 0,
    },
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
