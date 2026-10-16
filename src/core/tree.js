/**
 * 把 client_vars 的扁平 block_map 还原成块树。
 *
 * client_vars 块形如：
 *   { id, version, data: { type, parent_id, children: [...], ...类型专属字段 } }
 */

/**
 * @param {{blockMap: Object<string, any>, documentId: string, blockSequence?: string[]}} doc
 * @returns {{root: Object, byId: Map<string, Object>}}
 */
export function buildBlockTree(doc) {
  const { blockMap, documentId, blockSequence } = doc;
  const byId = new Map();

  for (const id of blockSequence ?? Object.keys(blockMap)) {
    const block = blockMap[id];
    if (!block) continue;
    const data = block.data ?? block;
    byId.set(id, {
      id,
      type: String(data.type ?? block.block_type ?? 'unknown'),
      data,
      children: [],
    });
  }

  // 兜底：sequence 里没有的块也纳入（防止分页遗漏）
  for (const [id, block] of Object.entries(blockMap)) {
    if (!byId.has(id)) {
      const data = block.data ?? block;
      byId.set(id, {
        id,
        type: String(data.type ?? block.block_type ?? 'unknown'),
        data,
        children: [],
      });
    }
  }

  // 显式 children 顺序优先；只有父块未显式声明 children 时才按 parent_id 归位，
  // 否则会把同一子块挂两次，导致导出内容重复。
  for (const node of byId.values()) {
    const order = node.data?.children;
    if (Array.isArray(order)) {
      node.children = order.map((id) => byId.get(id)).filter(Boolean);
      continue;
    }
    const parentId = node.data?.parent_id;
    const parent = parentId ? byId.get(parentId) : null;
    if (parent && !Array.isArray(parent.data?.children)) {
      parent.children.push(node);
    }
  }

  // 根节点：优先用 documentId 对应的 page 块
  let root = byId.get(documentId);
  if (!root || root.type !== 'page') {
    root = [...byId.values()].find((n) => n.type === 'page') ?? null;
  }
  if (root && root.data?.children?.length) {
    root.children = root.data.children.map((id) => byId.get(id)).filter(Boolean);
  }

  return { root, byId };
}

/** 判断块是否有内容（用于跳过空段落） */
export function hasTextContent(node) {
  const runs = node.data?.text?.initialAttributedTexts?.text;
  if (!runs) return false;
  return Object.values(runs).some((value) => typeof value === 'string' && value.trim().length > 0);
}
