/**
 * 构建脚本：esbuild 打包 MV3 扩展产物到 dist/。
 *
 *   node build.mjs          一次性构建
 *   node build.mjs --watch  监听模式
 */
import { build, context } from 'esbuild';
import fs from 'node:fs/promises';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const ROOT = path.dirname(fileURLToPath(import.meta.url));
const DIST = path.join(ROOT, 'dist');
const watch = process.argv.includes('--watch');

const common = {
  bundle: true,
  platform: 'browser',
  format: 'iife',
  target: ['chrome110'],
  loader: { '.css': 'text' },
  logLevel: 'info',
  sourcemap: false,
  minify: false,
  legalComments: 'none',
  define: { 'process.env.NODE_ENV': '"production"' },
};

const entryPoints = {
  'content/index': path.join(ROOT, 'src/content/index.js'),
  'background/service-worker': path.join(ROOT, 'src/background/service-worker.js'),
};

async function copyStatic() {
  await fs.mkdir(DIST, { recursive: true });
  await fs.copyFile(path.join(ROOT, 'src/manifest.json'), path.join(DIST, 'manifest.json'));
  // 图标等静态资源
  const iconsSrc = path.join(ROOT, 'src/icons');
  if (await fs.stat(iconsSrc).then(() => true).catch(() => false)) {
    await fs.cp(iconsSrc, path.join(DIST, 'icons'), { recursive: true });
  }
}

if (watch) {
  const ctx = await context({ ...common, entryPoints, outdir: DIST });
  await ctx.watch();
  await copyStatic();
  console.log('监听中… 修改 src/ 下文件将自动重建到 dist/');
} else {
  await build({ ...common, entryPoints, outdir: DIST });
  await copyStatic();
  console.log(`构建完成 -> ${DIST}`);
}
