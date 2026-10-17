/**
 * 触发飞书官方「下载为 Word」，并把产物落到 samples/official。
 * 兼容两种起始状态：① 设置对话框已打开 ② 什么都没有。
 *
 *   node tools/fetch-official-word3.mjs
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
await cdp.send('Page.enable');

await fs.mkdir(OUT, { recursive: true });
await cdp
  .send('Browser.setDownloadBehavior', { behavior: 'allow', downloadPath: OUT, eventsEnabled: true }, null)
  .catch(() => {});
console.log('下载目录:', OUT);

const shot = async (tag) => {
  const s = await cdp.send('Page.captureScreenshot', { format: 'png' });
  await fs.writeFile(path.join(OUT, `..`, 'out', `step-${tag}.png`), Buffer.from(s.data, 'base64'));
};

const clickAt = async (x, y) => {
  await cdp.send('Input.dispatchMouseEvent', { type: 'mouseMoved', x, y, button: 'none' });
  await sleep(120);
  await cdp.send('Input.dispatchMouseEvent', { type: 'mousePressed', x, y, button: 'left', clickCount: 1 });
  await sleep(60);
  await cdp.send('Input.dispatchMouseEvent', { type: 'mouseReleased', x, y, button: 'left', clickCount: 1 });
};

const centerOf = (expr) =>
  cdp.eval(`(() => {
    const el = ${expr};
    if (!el) return null;
    const r = el.getBoundingClientRect();
    if (!r.width || !r.height) return null;
    return { x: Math.round(r.x + r.width / 2), y: Math.round(r.y + r.height / 2), text: (el.textContent || '').trim().slice(0, 12) };
  })()`);

const hasDialog = () =>
  cdp.eval(`(() => {
    const btns = [...document.querySelectorAll('button')].filter((b) => (b.textContent || '').trim() === '导出');
    return btns.length > 0 && btns.some((b) => b.getBoundingClientRect().width > 20);
  })()`);

const before = new Set(await fs.readdir(OUT).catch(() => []));

// ① 已有「导出 Word 设置」对话框 → 直接确认
let isDialog = await hasDialog();
console.log('设置对话框已存在:', isDialog);

if (!isDialog) {
  // ② 走完整流程：••• → 下载为 → Word
  const more = await centerOf(`document.querySelector('.suite-more-menu')`);
  console.log('••• 位置:', JSON.stringify(more));
  if (!more) throw new Error('未找到 ••• 菜单');
  await clickAt(more.x, more.y);
  await sleep(1500);

  const dl = await centerOf(`[...document.querySelectorAll('li,div,span')].find(e => (e.textContent||'').trim() === '下载为')`);
  console.log('「下载为」:', JSON.stringify(dl));
  if (!dl) { await shot('no-download-item'); throw new Error('未找到「下载为」菜单项'); }
  await cdp.send('Input.dispatchMouseEvent', { type: 'mouseMoved', x: dl.x, y: dl.y });
  await sleep(1600);

  const word = await centerOf(`[...document.querySelectorAll('li,div,span')].find(e => (e.textContent||'').trim() === 'Word' && e.getBoundingClientRect().width > 20)`);
  console.log('「Word」:', JSON.stringify(word));
  if (!word) { await shot('no-word-item'); throw new Error('未找到 Word 子项'); }
  await clickAt(word.x, word.y);
  await sleep(1800);
  isDialog = await hasDialog();
  console.log('设置对话框出现:', isDialog);
  if (!isDialog) { await shot('no-dialog'); throw new Error('未出现导出设置对话框'); }
}

// ③ 选「仅正文」→ 点「导出」
const body = await centerOf(`[...document.querySelectorAll('label,span,div')].find(e => (e.textContent||'').trim() === '仅正文' && e.getBoundingClientRect().width > 10)`);
if (body) { await clickAt(body.x, body.y); console.log('已选「仅正文」'); await sleep(400); }

const btn = await centerOf(`[...document.querySelectorAll('button')].find(e => (e.textContent||'').trim() === '导出' && e.getBoundingClientRect().width > 20)`);
console.log('「导出」按钮:', JSON.stringify(btn));
if (!btn) { await shot('no-export-btn'); throw new Error('未找到「导出」按钮'); }
await clickAt(btn.x, btn.y);

// ④ 等文件
let got = null;
for (let i = 0; i < 90; i += 1) {
  await sleep(1000);
  const now = await fs.readdir(OUT).catch(() => []);
  const fresh = now.filter((f) => !before.has(f) && !f.endsWith('.crdownload'));
  if (fresh.length) {
    const f = fresh[0];
    if ((await fs.stat(path.join(OUT, f))).size > 2000) { got = f; break; }
  }
}
console.log(got ? `✓ 官方 Word 已下载: ${got}` : '✗ 未检测到下载文件');
cdp.close();
