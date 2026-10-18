/** 导出编排：解析 -> 取数 -> 渲染 -> 打包 */
import { loadDocument } from './api.js';
import { t } from '../shared/messages.js';
import { buildBlockTree } from './tree.js';
import { createRenderContext, renderNode } from './render/index.js';
import { buildDocx } from './docx/build.js';
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

  const ctx = createRenderContext({
    byId,
    // 表格单元格块不在 blockSequence 里（只存于 blockMap），列宽估算需要它兜底
    blockMap: doc.blockMap ?? {},
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

  onProgress?.({ stage: 'done', percent: 100 });
  return {
    blob,
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
