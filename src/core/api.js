/**
 * 飞书文档数据接口层。
 *
 * 运行在浏览器上下文（扩展 content script，或离线工具注入到页面后的脚本），
 * 复用页面自身的登录态与 CSRF token，只发只读 GET/查询类请求，不修改任何文档数据。
 *
 * 已验证的链路：
 *   1. /space/api/wiki/v2/tree/get_node/  wiki_token -> obj_token + obj_type
 *   2. /space/api/docx/pages/client_vars  按 cursor 分页拉取全部块（block_map + block_sequence）
 *   3. /space/api/v3/sheet/*              内嵌电子表格的全量数据
 */

const BASE_HEADERS = {
  Accept: 'application/json, text/plain, */*',
  'x-lgw-app-id': '1161',
  'x-lsc-bizid': '2',
  'x-lsc-terminal': 'web',
};

/** 接口异常，带 url / code / status 便于上层归类降级 */
export class FeishuApiError extends Error {
  constructor(message, meta = {}) {
    super(message);
    this.name = 'FeishuApiError';
    Object.assign(this, meta);
  }
}

export function readCookie(name) {
  const match = document.cookie.match(new RegExp(`(?:^|;\\s*)${name}=([^;]*)`));
  return match ? decodeURIComponent(match[1]) : '';
}

export function csrfToken() {
  return readCookie('_csrf_token') || readCookie('swp_csrf_token') || '';
}

export function requestHeaders(extra = {}) {
  const headers = { ...BASE_HEADERS, ...extra };
  const token = csrfToken();
  if (token) {
    headers['x-csrftoken'] = token;
    headers['X-Csrf-Token'] = token;
  }
  return headers;
}

/** 统一 JSON 请求：带 cookie、注入 csrf、归一化错误 */
export async function requestJson(url, { method = 'GET', body, headers, signal } = {}) {
  const init = { method, credentials: 'include', headers: requestHeaders(headers), signal };
  if (body !== undefined) {
    init.body = typeof body === 'string' ? body : JSON.stringify(body);
    init.headers['Content-Type'] = 'application/json';
  }

  let res;
  try {
    res = await fetch(url, init);
  } catch (error) {
    throw new FeishuApiError(`网络请求失败: ${error.message}`, { url });
  }
  if (!res.ok) {
    throw new FeishuApiError(`HTTP ${res.status}`, { url, status: res.status });
  }

  const text = await res.text();
  let json;
  try {
    json = JSON.parse(text);
  } catch {
    throw new FeishuApiError('响应不是合法 JSON', { url, payload: text.slice(0, 200) });
  }

  const code = json.code ?? 0;
  if (code !== 0 && code !== 200) {
    throw new FeishuApiError(`接口返回 code=${code} msg=${json.msg ?? ''}`, {
      url,
      code,
      payload: json,
    });
  }
  return json;
}

/**
 * 解析当前页面 URL，识别文档类型与 token。
 * 支持：
 *   /wiki/<wikiToken>
 *   /docx/<objToken>
 *   /docs/<objToken>
 *   /sheet/<sheetToken>
 */
export function parseDocUrl(input = location.href) {
  const url = new URL(input);
  const host = url.hostname;
  const parts = url.pathname.split('/').filter(Boolean);

  if (parts[0] === 'wiki' && parts[1]) {
    return { kind: 'wiki', host, wikiToken: parts[1], containerId: parts[1] };
  }
  if (parts[0] === 'docx' && parts[1]) {
    return { kind: 'docx', host, objToken: parts[1] };
  }
  if (parts[0] === 'docs' && parts[1]) {
    return { kind: 'docx', host, objToken: parts[1] };
  }
  if (parts[0] === 'sheet' && parts[1]) {
    return { kind: 'sheet', host, sheetToken: parts[1] };
  }
  throw new FeishuApiError(`无法识别的文档链接: ${url.pathname}`, { url: input });
}

/**
 * wiki_token -> 文档节点信息（obj_token / obj_type / 标题 / space_id）
 */
export async function resolveWikiNode(wikiToken, spaceId = '', signal) {
  const qs = new URLSearchParams({
    wiki_token: wikiToken,
    expand_shortcut: 'true',
    with_deleted: 'true',
  });
  if (spaceId) qs.set('space_id', spaceId);

  const json = await requestJson(`/space/api/wiki/v2/tree/get_node/?${qs}`, { signal });
  const data = json?.data ?? {};
  const node = data.node ?? data;

  if (!node?.obj_token) {
    throw new FeishuApiError('解析 wiki 节点失败：缺少 obj_token', { url: `get_node/${wikiToken}` });
  }
  return {
    objToken: node.obj_token,
    objType: node.obj_type,
    title: node.title ?? '',
    spaceId: node.space_id ?? spaceId ?? '',
    wikiToken,
  };
}

/**
 * 分页拉取文档全部块数据。
 * 每页返回 { block_map, block_sequence, cursor, has_more }。
 */
export async function fetchDocxClientVars(
  { objToken, spaceId = '', containerId = '', limit = 500, signal, onPage } = {},
) {
  const pages = [];
  let cursor = '';

  for (let guard = 0; guard < 500; guard += 1) {
    const qs = new URLSearchParams({ id: objToken, mode: '7', limit: String(limit) });
    if (cursor) qs.set('cursor', cursor);
    if (spaceId) qs.set('wiki_space_id', spaceId);
    if (containerId) {
      qs.set('container_type', 'wiki2.0');
      qs.set('container_id', containerId);
    }

    const json = await requestJson(`/space/api/docx/pages/client_vars?${qs}`, { signal });
    const data = json?.data ?? {};
    pages.push(data);
    onPage?.(data, pages.length);

    if (!data.has_more || !data.cursor) break;
    if (data.cursor === cursor) break; // 防御：游标不再推进
    cursor = data.cursor;
  }

  return pages;
}

/** 把分页结果合并为单一文档数据 */
export function mergeClientVars(pages) {
  const blockMap = {};
  const metaMap = {};
  const sequence = [];
  const seen = new Set();

  for (const page of pages) {
    Object.assign(blockMap, page.block_map ?? {});
    Object.assign(metaMap, page.meta_map ?? {});
    for (const id of page.block_sequence ?? []) {
      if (!seen.has(id)) {
        seen.add(id);
        sequence.push(id);
      }
    }
  }

  return {
    documentId: pages[0]?.id ?? '',
    blockMap,
    metaMap,
    blockSequence: sequence,
    structureVersion: pages[0]?.structure_version ?? 0,
  };
}

/** 内嵌电子表格：拉取全量单元格数据（解决官方导出「显示不全」的关键） */
export async function fetchSheetClientVars({ sheetToken, signal, body = {} } = {}) {
  const json = await requestJson('/space/api/v3/sheet/client_vars', {
    method: 'POST',
    body: { token: sheetToken, ...body },
    signal,
  });
  return json?.data ?? {};
}

/**
 * 按页面地址自动解析并拉取完整文档数据。
 * 返回 { node, pages, doc }，其中 doc 为合并后的 { blockMap, blockSequence, metaMap }。
 */
/**
 * 从块数据里取文档大标题：根 `page` 块（没有则用 blockSequence 首块）的文本。
 * 飞书文档模型里，标题不是独立块，而是挂在根块 text 上。
 */
function extractPageTitle(doc) {
  const map = doc?.blockMap ?? {};
  const seq = doc?.blockSequence ?? [];
  const candidates = [
    map[doc?.documentId],
    ...seq.slice(0, 3).map((id) => map[id]),
  ].filter(Boolean);
  const root = candidates.find((b) => b?.data?.type === 'page') ?? candidates[0];
  const texts = root?.data?.text?.initialAttributedTexts?.text ?? {};
  return Object.keys(texts)
    .sort((a, b) => Number(a) - Number(b))
    .map((k) => texts[k] ?? '')
    .join('')
    .trim();
}

export async function loadDocument({ url = location.href, signal, onProgress } = {}) {
  const target = parseDocUrl(url);
  onProgress?.({ stage: 'resolve', detail: target.kind });


  let node;
  if (target.kind === 'wiki') {
    node = await resolveWikiNode(target.wikiToken, '', signal);
  } else {
    node = { objToken: target.objToken, objType: 22, title: '', spaceId: '', wikiToken: '' };
  }

  onProgress?.({ stage: 'fetch', detail: node.title || node.objToken });

  const pages = await fetchDocxClientVars({
    objToken: node.objToken,
    spaceId: node.spaceId,
    containerId: node.wikiToken,
    signal,
    onPage: (page, index) => onProgress?.({ stage: 'fetch', page: index, blocks: Object.keys(page.block_map ?? {}).length }),
  });

  const doc = mergeClientVars(pages);

  // 标题兜底：/docx/<token> 链接拿不到 node.title（get_node 只服务 wiki），
  // 但文档大标题就在根 `page` 块的文本里。不兜底会退化成「未命名文档」，
  // 而且会被当成大标题写进产物（实测模板文档全部中招）。
  if (!node.title) {
    node.title = extractPageTitle(doc);
  }

  return { node, pages, doc };
}

/** gzip + base64 → 解析后的 JSON（多维表格的 base/table 就是这么打包的） */
async function gunzipBase64ToJson(b64) {
  if (!b64) return null;
  const bin = atob(b64);
  const bytes = new Uint8Array(bin.length);
  for (let i = 0; i < bin.length; i += 1) bytes[i] = bin.charCodeAt(i);
  if (typeof DecompressionStream === 'undefined') return null;
  const ds = new DecompressionStream('gzip');
  const stream = new Blob([bytes]).stream().pipeThrough(ds);
  const text = await new Response(stream).text();
  return JSON.parse(text);
}

/**
 * 多维表格（bitable）数据。
 *
 * 接口与字段来自真实抓包（samples/bodies/0050_*.json）：
 *   GET /space/api/v1/bitable/<baseToken>_<tableId>/clientvars?tableID=...&recordLimit=200...
 *   data.base / data.table 是 base64(gzip(JSON))：
 *     viewMap[view].property.fields → 列顺序
 *     fieldMap[fid]                 → 字段名/类型
 *     groupList[].recordIDList      → 行顺序
 *     recordMap[rid][fid].value     → 单元格值
 */
export async function fetchBitableClientVars({ token, viewId = '', recordLimit = 200, signal } = {}) {
  if (!token) return null;
  const [baseToken, tableId] = String(token).split('_');
  if (!baseToken || !tableId) return null;

  const qs = new URLSearchParams({
    tableID: tableId,
    viewID: viewId ?? '',
    recordLimit: String(recordLimit),
    needBase: 'true',
    openType: '0',
    viewLazyLoad: 'true',
    ondemandVer: '2',
    ondemandLimit: String(recordLimit),
    noMissCS: 'true',
    optimizationFlag: '1',
    removeFmlExtra: 'true',
  });
  // 路径里只放 baseToken，tableId 走 query（写成 <base>_<table> 会返回 code=3 Not Found，实测）
  const json = await requestJson(`/space/api/v1/bitable/${baseToken}/clientvars?${qs}`, { signal });
  const data = json?.data ?? {};

  const [base, table] = await Promise.all([
    gunzipBase64ToJson(data.base),
    gunzipBase64ToJson(data.table),
  ]);
  return normalizeBitable(base, table);
}

/** 把 base/table 两份 JSON 合并成「列 + 行 + 单元格文本」的简单结构 */
function normalizeBitable(base, table) {
  if (!base && !table) return null;
  const fieldMap = { ...(base?.fieldMap ?? table?.fieldMap ?? {}) };
  const recordMap = { ...(base?.recordMap ?? table?.recordMap ?? {}) };

  // 列顺序：当前视图的字段列表 → 否则用 fieldMap 的键顺序
  const viewMap = table?.viewMap ?? base?.viewMap ?? {};
  const viewId = table?.currentViewId ?? base?.currentViewId ?? Object.keys(viewMap)[0];
  const viewFields = viewMap?.[viewId]?.property?.fields;
  let fids = Array.isArray(viewFields) && viewFields.length
    ? viewFields.map((f) => (typeof f === 'string' ? f : f?.id)).filter(Boolean)
    : Object.keys(fieldMap);

  // 行顺序
  let rids = [];
  const groups = table?.groupList ?? base?.groupList;
  if (Array.isArray(groups) && groups.length) {
    for (const g of groups) rids.push(...(g?.recordIDList ?? []));
  }
  if (!rids.length) rids = Object.keys(recordMap);

  // 单元格取值：value 可能是数组/对象/字符串，统一转成显示文本
  const toText = (v) => {
    if (v == null) return '';
    if (typeof v === 'string') return v;
    if (Array.isArray(v)) return v.map(toText).filter(Boolean).join(', ');
    if (typeof v === 'object') {
      if (typeof v.text === 'string') return v.text;
      if (typeof v.value === 'string') return v.value;
      if (typeof v.name === 'string') return v.name;
      if (Array.isArray(v.value)) return toText(v.value);
      return '';
    }
    return String(v);
  };

  const headers = fids.map((fid) => String(fieldMap[fid]?.name ?? fid));
  const rows = rids.map((rid) => fids.map((fid) => toText(recordMap[rid]?.[fid]?.value)));

  if (!headers.length) return null;
  return { headers, rows, fieldTypes: fids.map((fid) => fieldMap[fid]?.type ?? null) };
}
