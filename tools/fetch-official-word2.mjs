/**
 * 用真实鼠标事件走「••• → 下载为 → Word」，并把下载目录重定向到 samples/official。
 *
 *   node tools/fetch-official-word2.mjs
 */
import fs from 'node:fs/promises';
import path from 'node:path';
import { Cdp, PORT, sleep } from './cdp.mjs';

const OUT = path.resolve(import.meta.dirname, '..', 'samples', 'official');
const cdp = await Cdp.connect(PORT);
const list = await cdp.targets();
const page = list.find((t) => t.type === 'page' && /feishu\.cn/.test(t.url));
if (!page) throw new Error('找不到飞书文档标签页');
const { sessionId } = await cdp.send('Target.attachToTarget', { targetId: page.id, flatten: true }, null);
cdp.sessionId = sessionId;
await cdp.send('Runtime.enable');

await fs.mkdir(OUT, { recursive: true });
// 注意：Browser.* 是浏览器级命令，不能带 sessionId
await cdp.send('Browser.setDownloadBehavior', { behavior: 'allow', downloadPath: OUT, eventsEnabled: true }, null)
  .then(() => console.log('下载目录已重定向:', OUT))
  .catch((e) => console.log('Browser.setDownloadBehavior 失败:', e.message));

const clickAt = async (x, y) => {
  await cdp.send('Input.dispatchMouseEvent', { type: 'mouseMoved', x, y, button: 'none' });
  await sleep(120);
  await cdp.send('Input.dispatchMouseEvent', { type: 'mousePressed', x, y, button: 'left', clickCount: 1 });
  await sleep(60);
  await cdp.send('Input.dispatchMouseEvent', { type: 'mouseReleased', x, y, button: 'left', clickCount: 1 });
};

const centerOf = (code) => cdp.eval(`(() => {
  const el = ${code};
  if (!el) return null;
  const r = el.getBoundingClientRect();
  if (!r.width) return null;
  return { x: Math.round(r.x + r.width / 2), y: Math.round(r.y + r.height / 2), text: (el.textContent||'').trim().slice(0,10) };
})()`);

const before = new Set(await fs.readdir(OUT).catch(() => []));

// 1) 打开 ••• 菜单
let pos = await centerOf(`document.querySelector('.suite-more-menu')`);
console.log('••• 位置:', JSON.stringify(pos));
if (!pos) throw new Error('未找到 .suite-more-menu');
await clickAt(pos.x, pos.y);
await sleep(1500);

// 2) 悬停「下载为」（子菜单需要 hover 展开）
pos = await centerOf(`[...document.querySelectorAll('li,div,span')].find(e => (e.textContent||'').trim() === '下载为')`);
console.log('「下载为」位置:', JSON.stringify(pos));
if (pos) {
  await cdp.send('Input.dispatchMouseEvent', { type: 'mouseMoved', x: pos.x, y: pos.y });
  await sleep(1600);
}

// 3) 点 Word（子菜单项）
pos = await centerOf(`[...document.querySelectorAll('li,div,span')].find(e => (e.textContent||'').trim() === 'Word' && e.getBoundingClientRect().width > 20)`);
console.log('「Word」位置:', JSON.stringify(pos));
if (!pos) {
  // 打印子菜单里可见的短文本，便于排错
  const txts = await cdp.eval(`[...document.querySelectorAll('li,div,span')]
    .filter(e => e.children.length === 0 && (e.textContent||'').trim().length < 10)
    .map(e => (e.textContent||'').trim()).filter(Boolean).slice(0, 30)`);
  console.log('可见短文本:', JSON.stringify(txts));
} else {
  await clickAt(pos.x, pos.y);
}

// 3.5) 处理「导出 Word 设置」对话框：选「仅正文」并点「导出」
await sleep(1500);
const onlyBody = await centerOf(`[...document.querySelectorAll('label,span,div')].find(e => (e.textContent||'').trim() === '仅正文' && e.getBoundingClientRect().width > 10)`);
if (onlyBody) {
  await clickAt(onlyBody.x, onlyBody.y);
  console.log('已选「仅正文」:', JSON.stringify(onlyBody));
  await sleep(400);
}
const exportBtn = await centerOf(`[...document.querySelectorAll('button')].find(e => (e.textContent||'').trim() === '导出' && e.getBoundingClientRect().width > 20)`);
console.log('「导出」按钮:', JSON.stringify(exportBtn));
if (exportBtn) await clickAt(exportBtn.x, exportBtn.y);

// 4) 等文件落盘
let got = null;
for (let i = 0; i < 60; i += 1) {
  await sleep(1000);
  const now = await fs.readdir(OUT).catch(() => []);
  const fresh = now.filter((f) => !before.has(f) && !f.endsWith('.crdownload'));
  if (fresh.length) {
    const f = fresh[0];
    if ((await fs.stat(path.join(OUT, f))).size > 2000) { got = f; break; }
  }
}
console.log(got ? `✓ 下载完成: ${got}` : '✗ 未检测到下载');
cdp.close();
