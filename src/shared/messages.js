/**
 * 导出产物中的文案（占位 / 降级提示 / 链接标签）。
 *
 * 为什么单独抽出来：这些字符串会**写进导出的 Word 正文**，
 * 如果写死中文，导出英文文档时会冒出中文占位。
 * 这里按文档语言自动选择（中文 / 英文），面板 UI 仍为中文。
 */

export const MSG = {
  zh: {
    unsupported: (type) => `未支持的内容块：${type}`,
    blockError: (type) => `${type} 渲染异常`,
    imageMissing: (token) => `图片未能从页面获取${token ? `（token ${token}）` : ''}`,
    fileAsLink: (name) => `附件「${name}」以链接形式保留`,
    sheetAsLink: () => '内嵌电子表格未能截图，已降级为链接',
    whiteboardMissing: () => '画板内容未能截图，已降级为占位',
    whiteboardFallback: () => '[画板内容未能导出]',
    bitableAsLink: () => '多维表格以链接形式保留',
    sheetLabel: () => '📊 内嵌电子表格：',
    bitableLabel: () => '🔗 多维表格：',
    fileLabel: () => '📎 ',
    noToken: () => '（缺少 token）',
  },
  en: {
    unsupported: (type) => `Unsupported block: ${type}`,
    blockError: (type) => `${type} render failed`,
    imageMissing: (token) => `Image not captured${token ? ` (token ${token})` : ''}`,
    fileAsLink: (name) => `Attachment "${name}" kept as a link`,
    sheetAsLink: () => 'Embedded spreadsheet not captured, kept as a link',
    whiteboardMissing: () => 'Whiteboard not captured, kept as a placeholder',
    whiteboardFallback: () => '[Whiteboard content not exported]',
    bitableAsLink: () => 'Bitable kept as a link',
    sheetLabel: () => '📊 Embedded spreadsheet: ',
    bitableLabel: () => '🔗 Bitable: ',
    fileLabel: () => '📎 ',
    noToken: () => '(missing token)',
  },
};

/**
 * 按文档正文判断语言：中文字符占比超过 5% 视为中文文档。
 * @param {string} sample 文档正文抽样（取前若干字符即可）
 */
export function detectLang(sample = '') {
  const text = String(sample).slice(0, 4000);
  if (!text) return 'zh';
  const cjk = (text.match(/[㐀-䶿一-鿿豈-﫿]/g) ?? []).length;
  return cjk / text.length > 0.05 ? 'zh' : 'en';
}

/** 取文案：`t('fileAsLink', lang)('手册.pdf')` */
export function t(key, lang = 'zh') {
  const dict = MSG[lang] ?? MSG.zh;
  return dict[key] ?? MSG.zh[key] ?? (() => key);
}

/**
 * 飞书专有块的占位文案（Word 无对应形式，按用户约定不强求兼容，
 * 但要说清「这是什么、去哪看」，而不是笼统的「未支持的内容块」）。
 */
const FEISHU_ONLY = {
  zh: {
    chart_refer_host_perm: () => '图表（飞书图表，需在原文档中查看）',
    task_list: () => '任务列表（飞书任务，Word 无对应形式）',
    synced_source: () => '同步块（源内容为空）',
  },
  en: {
    chart_refer_host_perm: () => 'Chart (Feishu chart, view it in the original document)',
    task_list: () => 'Task list (Feishu Tasks, no Word equivalent)',
    synced_source: () => 'Synced block (source is empty)',
  },
};

export function feishuOnly(type, lang = 'zh') {
  const dict = FEISHU_ONLY[lang] ?? FEISHU_ONLY.zh;
  const fn = dict[type];
  if (fn) return fn();
  return t('unsupported', lang)(type);
}
