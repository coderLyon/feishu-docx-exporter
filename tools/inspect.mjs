/**
 * 结构分析工具：打印任意 JSON 文件的骨架（类型 / 数组长度 / 键值样本），用于快速摸清接口响应。
 *
 * 用法：node tools/inspect.mjs <file.json> [maxDepth] [grepKey]
 */
import fs from 'node:fs';
import path from 'node:path';

const [, , fileArg, depthArg, grepArg] = process.argv;
const maxDepth = Number(depthArg ?? 3);
const grep = grepArg ? new RegExp(grepArg, 'i') : null;
const file = path.resolve(process.cwd(), fileArg);

let MAX_LINES = 400;
let lines = 0;

function skeleton(value, depth, key, out) {
  if (lines >= MAX_LINES) return;
  const pad = '  '.repeat(depth);
  if (grep && !grep.test(key)) {
    if (typeof value !== 'object' || value === null) return;
    const hasDeep = JSON.stringify(value).includes(key);
    if (!hasDeep) return;
  }
  if (value === null || value === undefined) {
    out.push(`${pad}${key}: ${value === null ? 'null' : 'undefined'}`);
    lines += 1;
    return;
  }
  if (Array.isArray(value)) {
    out.push(`${pad}${key}: Array(${value.length})`);
    lines += 1;
    if (depth < maxDepth) {
      value.slice(0, 2).forEach((item, i) => skeleton(item, depth + 1, `[${i}]`, out));
    }
    return;
  }
  if (typeof value === 'object') {
    out.push(`${pad}${key}: {${Object.keys(value).length}}`);
    lines += 1;
    if (depth < maxDepth) {
      for (const [k, v] of Object.entries(value)) skeleton(v, depth + 1, k, out);
    }
    return;
  }
  const text = typeof value === 'string' ? JSON.stringify(value.length > 70 ? `${value.slice(0, 70)}…` : value) : String(value);
  out.push(`${pad}${key}: ${typeof value} = ${text}`);
  lines += 1;
}

const raw = fs.readFileSync(file, 'utf8');
let data;
try {
  data = JSON.parse(raw);
} catch {
  // 部分接口返回 JSON Lines 或带前缀
  const start = raw.search(/[[{]/);
  data = JSON.parse(raw.slice(start));
}

const out = [];
skeleton(data, 0, 'root', out);
console.log(out.join('\n'));
if (lines >= MAX_LINES) console.log(`\n…(输出已达 ${MAX_LINES} 行上限，调小 maxDepth 或用 grepKey 过滤)`);
