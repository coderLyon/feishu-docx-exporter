/**
 * Protobuf wire 格式解析/转储工具：用于摸清飞书内嵌电子表格的二进制块结构。
 * 无需 .proto 文件，按 wire type 递归解析并猜测嵌套结构。
 *
 * 用法：
 *   node tools/proto-dump.mjs <file.json> [maxDepth] [maxLines]
 *   node tools/proto-dump.mjs <file.json> --strings       只列出所有可打印字符串及其字段路径
 */
import fs from 'node:fs';
import path from 'node:path';
import zlib from 'node:zlib';

const [, , fileArg, arg2, arg3] = process.argv;
const STRINGS_ONLY = arg2 === '--strings';
const MAX_DEPTH = STRINGS_ONLY ? 10 : Number(arg2 ?? 4);
const MAX_LINES = Number(arg3 ?? 200);
const file = path.resolve(process.cwd(), fileArg);

let lines = 0;

function readVarint(buf, pos) {
  let result = 0;
  let shift = 0;
  for (let i = 0; i < 10; i += 1) {
    const byte = buf[pos + i];
    if (byte === undefined) throw new Error('varint 越界');
    result += (byte & 0x7f) * 2 ** shift;
    if ((byte & 0x80) === 0) return [result, pos + i + 1];
    shift += 7;
  }
  throw new Error('varint 过长');
}

const PRINTABLE = /^[^\u0000-\u0008\u000b\u000c\u000e-\u001f]+$/u;

function isText(slice) {
  if (slice.length === 0) return false;
  const text = slice.toString('utf8');
  if (!PRINTABLE.test(text)) return false;
  // 避免把二进制字节误判为中文
  return !/�/.test(text);
}

function parse(buf, start, end, depth, out, path) {
  let pos = start;
  while (pos < end) {
    try {
      const [key, afterKey] = readVarint(buf, pos);
      pos = afterKey;
      const field = key >>> 3;
      const wire = key & 7;

      if (wire === 0) {
        const [value, next] = readVarint(buf, pos);
        pos = next;
        out.push({ kind: 'varint', field, value, path });
      } else if (wire === 2) {
        const [len, afterLen] = readVarint(buf, pos);
        pos = afterLen;
        const slice = buf.subarray(pos, pos + len);
        pos += len;
        const text = isText(slice) ? slice.toString('utf8') : null;
        out.push({ kind: 'bytes', field, len, text, path });
        if (text === null && depth < MAX_DEPTH) {
          // 即使嵌套解析中途失败，也保留已解析出的部分结果
          const children = [];
          try {
            parse(slice, 0, slice.length, depth + 1, children, `${path}/${field}`);
          } catch {
            /* 不是合法嵌套消息，按原始字节处理 */
          }
          if (children.length) {
            out.push({ kind: 'message', field, children, path: `${path}/${field}` });
          }
        }
      } else if (wire === 5) {
        out.push({ kind: 'fixed32', field, path });
        pos += 4;
      } else if (wire === 1) {
        out.push({ kind: 'fixed64', field, path });
        pos += 8;
      } else {
        out.push({ kind: `wire${wire}`, field, path });
        break;
      }
    } catch (error) {
      out.push({ kind: 'error', message: error.message, offset: pos, path });
      break;
    }
  }
  return out;
}

function print(nodes, indent = 0) {
  for (const node of nodes) {
    if (lines >= MAX_LINES) return;
    lines += 1;
    if (node.kind === 'bytes') {
      console.log(`${'  '.repeat(indent)}f${node.field} bytes(${node.len})${node.text ? ` = ${JSON.stringify(node.text.length > 60 ? `${node.text.slice(0, 60)}…` : node.text)}` : ''}`);
    } else if (node.kind === 'varint') {
      console.log(`${'  '.repeat(indent)}f${node.field} varint = ${node.value}`);
    } else if (node.kind === 'message') {
      console.log(`${'  '.repeat(indent)}f${node.field} {`);
      print(node.children, indent + 1);
      console.log(`${'  '.repeat(indent)}}`);
    } else {
      console.log(`${'  '.repeat(indent)}f${node.field} ${node.kind}`);
    }
  }
}

const json = JSON.parse(fs.readFileSync(file, 'utf8'));

/** 收集所有 gzip+base64 的块 */
function collectBlocks(node, out = {}, keyPath = '') {
  if (typeof node === 'string' && node.length > 80 && /^H4sIA/.test(node)) {
    out[keyPath] = node;
    return out;
  }
  if (Array.isArray(node)) {
    node.forEach((item, i) => collectBlocks(item, out, `${keyPath}[${i}]`));
    return out;
  }
  if (node && typeof node === 'object') {
    for (const [key, value] of Object.entries(node)) {
      collectBlocks(value, out, keyPath ? `${keyPath}.${key}` : key);
    }
  }
  return out;
}

const blocks = collectBlocks(json);

for (const [key, encoded] of Object.entries(blocks)) {
  const buf = zlib.gunzipSync(Buffer.from(encoded, 'base64'));
  console.log(`\n########## ${key}  (${buf.length} bytes) ##########`);

  if (STRINGS_ONLY) {
    const nodes = [];
    parse(buf, 0, buf.length, 0, nodes, '');
    const strings = nodes.filter((n) => n.kind === 'bytes' && n.text);
    console.log(`共 ${strings.length} 个可打印字符串（字段路径统计如下）`);
    const byPath = new Map();
    for (const s of strings) byPath.set(s.path, (byPath.get(s.path) ?? 0) + 1);
    for (const [p, count] of byPath) console.log(`  ${p} x${count}`);
    console.log('\n前 60 个字符串:');
    strings.slice(0, 60).forEach((s, i) => console.log(`  [${i}] f${s.field}@${s.path} = ${JSON.stringify(s.text)}`));
  } else {
    const nodes = [];
    try {
      parse(buf, 0, buf.length, 0, nodes, '');
      print(nodes);
    } catch (error) {
      console.log(`解析失败: ${error.message}`);
    }
  }
}
