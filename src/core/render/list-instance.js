export function orderedSequence(node) {
  const value = node.data?.seq;
  if (!/^\d+$/.test(String(value ?? '').trim())) return null;
  const number = Number(value);
  return number > 0 && number <= 2147483647 ? number : null;
}

/** 同一组兄弟项中，显式编号偏离自然顺序时开启新编号实例。 */
export function startsNewOrderedInstance(node, previous) {
  const sequence = orderedSequence(node);
  return previous == null || (sequence !== null && sequence !== previous + 1);
}
