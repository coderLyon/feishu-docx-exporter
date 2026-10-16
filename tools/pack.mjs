/**
 * 把 dist/ 打包为可分发的 zip（manifest.json 位于压缩包根目录）。
 *
 *   node tools/pack.mjs
 *
 * 覆盖前会把旧包按时间戳备份到 output/zip-backup/。
 */
import fs from 'node:fs';
import fsp from 'node:fs/promises';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import JSZip from 'jszip';

const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const DIST = path.join(ROOT, 'dist');
const ZIP_NAME = '飞书文档导出Word扩展.zip';
const TARGET = path.join(ROOT, ZIP_NAME);
const BACKUP_DIR = path.join(ROOT, 'output', 'zip-backup');

function stamp() {
  const d = new Date();
  const p = (n) => String(n).padStart(2, '0');
  return `${d.getFullYear()}${p(d.getMonth() + 1)}${p(d.getDate())}-${p(d.getHours())}${p(d.getMinutes())}${p(d.getSeconds())}`;
}

/** 递归收集目录下所有文件，返回相对路径 */
function walk(dir, base = dir) {
  const out = [];
  for (const entry of fs.readdirSync(dir, { withFileTypes: true })) {
    const full = path.join(dir, entry.name);
    if (entry.isDirectory()) out.push(...walk(full, base));
    else out.push(path.relative(base, full).split(path.sep).join('/'));
  }
  return out;
}

async function main() {
  if (!fs.existsSync(path.join(DIST, 'manifest.json'))) {
    throw new Error('dist/manifest.json 不存在，请先执行 npm run build');
  }

  const zip = new JSZip();
  let total = 0;
  for (const rel of walk(DIST)) {
    const buf = await fsp.readFile(path.join(DIST, rel));
    zip.file(rel, buf);
    total += buf.length;
    console.log(`  + ${rel}  (${(buf.length / 1024).toFixed(1)} KB)`);
  }

  const buffer = await zip.generateAsync({
    type: 'nodebuffer',
    compression: 'DEFLATE',
    compressionOptions: { level: 9 },
  });

  if (fs.existsSync(TARGET)) {
    await fsp.mkdir(BACKUP_DIR, { recursive: true });
    const backup = path.join(BACKUP_DIR, `飞书文档导出Word扩展-${stamp()}.zip`);
    await fsp.copyFile(TARGET, backup);
    console.log(`旧包已备份 -> ${path.relative(ROOT, backup)}`);
  }

  await fsp.writeFile(TARGET, buffer);
  console.log(`\n打包完成 -> ${ZIP_NAME}`);
  console.log(`  文件数 ${walk(DIST).length} · 原始 ${(total / 1024).toFixed(1)} KB · 压缩后 ${(buffer.length / 1024).toFixed(1)} KB`);
}

main().catch((error) => {
  console.error('打包失败:', error.message);
  process.exit(1);
});
