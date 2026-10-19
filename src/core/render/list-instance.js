/**
 * 有序列表「实例」判定：决定某个列表项是开启新列表，还是延续上一个。
 *
 * 判据来自飞书数据本身：
 *   seq === '1'    → 新列表的第一项（作者在此重新开始编号）
 *   seq === 'auto' → 续编号，即使中间隔着补充说明段落（3.4 的情况）
 *   上一个兄弟不是列表项（标题/表格/代码/引用/图片/其他列表） → 也开启新列表
 *
 * 单独抽成纯函数，便于单测覆盖这两类历史缺陷：
 *   只按「上一项是否列表项」判断 → 两个独立列表被正文隔开时会粘成一个（7.3 的缺陷）；
 *   只按 seq 判断 → 列表紧跟标题且首项为 auto 时会漏掉重开。
 */

/** @param {{data?: {seq?: unknown}}} node @param {boolean} prevWasOrdered */
export function startsNewOrderedInstance(node, prevWasOrdered) {
  const raw = node?.data?.seq;
  const hasExplicitSeq = raw != null && raw !== 'auto' && String(raw).trim() !== '';
  const explicitStart = hasExplicitSeq && Number(raw) === 1;
  return explicitStart || !prevWasOrdered;
}
