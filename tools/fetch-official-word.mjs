/**
 * 自动化飞书「下载为 Word」：点右上角 ••• → 下载为 → Word，
 * 并用 CDP 把下载重定向到 samples/official。
 *
 *   node tools/fetch-official-word.mjs
 *
 * 前置：调试 Chrome 已打开一篇**已登录**的飞书文档（/wiki/ 或 /docx/）。
 */
import fs from 'node:fs/promises';
import path from 'node:path';
import { Cdp, PORT, sleep } from './cdp.mjs';

const OUT = path.resolve(import.meta.dirname, '..', 'samples', 'official');

const cdp = await Cdp.connect(PORT);
const list = await cdp.targets();
const page = list.find((t) => t.type === 'page' && /feishu\.cn|feishu\.cn\/|larkoffice/.test(t.url));
if (!page) throw new Error('没有找到飞书文档标签页');
console.log('页面:', page.url);
const { sessionId } = await cdp.send('Target.attachToTarget', { targetId: page.id, flatten: true }, null);
cdp.sessionId = sessionId;
await cdp.send('Runtime.enable');
await cdp.send('Page.enable');

await fs.mkdir(OUT, { recursive: true });
// 下载目录重定向（事件级别，页面感知不到）
await cdp.send('Browser.setDownloadBehavior', {
  behavior: 'allow',
  downloadPath: OUT,
  eventsEnabled: true,
}).catch(async () => {
  await cdp.send('Page.setDownloadBehavior', { behavior: 'allow', downloadPath: OUT });
});
console.log('下载目录:', OUT);

const before = new Set(await fs.readdir(OUT).catch(() => []));

// 1) 点「更多」按钮：优先 aria-label，其次按位置（右上角、图标按钮）
const clickedMore = await cdp.eval(`(() => {
  const cands = [...document.querySelectorAll('button, [role="button"], span, div')].filter((el) => {
    const r = el.getBoundingClientRect();
    if (r.width < 12 || r.height < 12 || r.width > 60) return false;
    if (r.top > 110 || r.left < window.innerWidth - 420) return false;
    const label = (el.getAttribute('aria-label') || '') + (el.getAttribute('title') || '') + (el.getAttribute('data-testid') || '');
    return /更多|more|ellipsis|operat/i.test(label) || el.className.toString().includes('more');
  });
  if (!cands.length) return { ok: false, tried: 0 };
  const el = cands[0];
  el.click();
  return { ok: true, cls: String(el.className).slice(0, 40), label: el.getAttribute('aria-label') || '' };
})()`);
console.log('点击「更多」:', JSON.stringify(clickedMore));
await sleep(1200);

// 2) 找菜单里的「下载为」，悬停展开子菜单
const hoverDownload = await cdp.eval(`(() => {
  const nodes = [...document.querySelectorAll('div, li, span')];
  const hit = nodes.find((el) => (el.textContent || '').trim() === '下载为');
  if (!hit) return { ok: false, sample: nodes.filter(e => (e.textContent||'').trim().length < 8 && e.children.length === 0).slice(0, 12).map(e => (e.textContent||'').trim()) };
  const target = hit.closest('[role="menuitem"], li, div');
  target.dispatchEvent(new MouseEvent('mouseenter', { bubbles: true }));
  target.dispatchEvent(new MouseEvent('mouseover', { bubbles: true }));
  return { ok: true, tag: target.tagName, cls: String(target.className).slice(0, 40) };
})()`);
console.log('悬停「下载为」:', JSON.stringify(hoverDownload).slice(0, 400));
await sleep(1500);

// 3) 点 Word
const clickWord = await cdp.eval(`(() => {
  const nodes = [...document.querySelectorAll('div, li, span')];
  const hit = nodes.find((el) => (el.textContent || '').trim() === 'Word');
  if (!hit) return { ok: false, texts: nodes.filter(e => (e.textContent||'').trim().length < 6 && e.children.length === 0).map(e => (e.textContent||'').trim()).filter(Boolean).slice(0, 20) };
  hit.click();
  return { ok: true, tag: hit.tagName };
})()`);
console.log('点击 Word:', JSON.stringify(clickWord).slice(0, 400));

// 4) 等下载出现新文件
let downloaded = null;
for (let i = 0; i < 40; i += 1) {
  await sleep(1000);
  const now = await fs.readdir(OUT);
  const fresh = now.filter((f) => !before.has(f) && !f.endsWith('.crdownload'));
  if (fresh.length) {
    const f = fresh[0];
    const st = await fs.stat(path.join(OUT, f));
    if (st.size > 1000) { downloaded = f; break; }
  }
}
if (downloaded) {
  console.log('✓ 已下载:', downloaded, Math.round((await fs.stat(path.join(OUT, downloaded))).size / 1024), 'KB');
} else {
  console.log('✗ 未检测到下载文件，目录内容:', (await fs.readdir(OUT)).join(', ') || '(空)');
}
cdp.close();
