/**
 * 读取内嵌表格复制后的 **HTML flavor**（含背景色/边框/字体等样式）。
 *
 *   node tools/probe-sheet-html.mjs
 *
 * 剪贴板通常同时有 text/plain（TSV）与 text/html（带样式表格）。
 * 若能拿到 HTML，就能还原飞书表格的配色与边框，而不是只搬文字。
 */
import fs from 'node:fs/promises';
import { Cdp, PORT, sleep } from './cdp.mjs';

const cdp = await Cdp.connect(PORT);
const list = await cdp.targets();
const page = list.find((t) => t.type === 'page');
const { sessionId } = await cdp.send('Target.attachToTarget', { targetId: page.id, flatten: true }, null);
cdp.sessionId = sessionId;
await cdp.send('Runtime.enable');
await cdp.send('Page.enable');
await cdp.send('Browser.grantPermissions', {
  origin: new URL(page.url).origin,
  permissions: ['clipboardReadWrite', 'clipboardSanitizedWrite'],
}, null).catch(() => {});

// 点进表格 → 全选 → 复制（用 CDP trust 输入）
const box = await cdp.eval(`(() => {
  const b = document.querySelector('[data-block-type="sheet"], .docx-sheet-block');
  if (!b) return null;
  b.scrollIntoView({ block: 'center' });
  const r = b.getBoundingClientRect();
  return { x: Math.round(r.x + r.width / 2), y: Math.round(r.y + Math.min(120, r.height / 2)) };
})()`);
if (!box) { console.log('未找到表格'); process.exit(1); }
await sleep(800);

await cdp.send('Input.dispatchMouseEvent', { type: 'mousePressed', x: box.x, y: box.y, button: 'left', clickCount: 1 });
await sleep(60);
await cdp.send('Input.dispatchMouseEvent', { type: 'mouseReleased', x: box.x, y: box.y, button: 'left', clickCount: 1 });
await sleep(1400);

const key = async (code, keyCode, mods = 0) => {
  await cdp.send('Input.dispatchKeyEvent', { type: 'keyDown', windowsVirtualKeyCode: keyCode, nativeVirtualKeyCode: keyCode, code, key: code, modifiers: mods });
  await sleep(60);
  await cdp.send('Input.dispatchKeyEvent', { type: 'keyUp', windowsVirtualKeyCode: keyCode, nativeVirtualKeyCode: keyCode, code, key: code, modifiers: mods });
};
await key('KeyA', 65, 2);
await sleep(700);
await key('KeyC', 67, 2);
await sleep(1200);

const result = await cdp.eval(`(async () => {
  try {
    const items = await navigator.clipboard.read();
    const out = [];
    for (const it of items) {
      for (const type of it.types) {
        let val = '';
        if (type === 'text/plain') val = await (await it.getType(type)).text();
        else if (type === 'text/html') val = await (await it.getType(type)).text();
        else val = '<' + type + ' ' + (await it.getType(type)).size + ' bytes>';
        out.push({ type, len: val.length, head: val.slice(0, 700) });
      }
    }
    return out;
  } catch (e) {
    return [{ err: String(e) }];
  }
})()`);

console.log(JSON.stringify(result, null, 1).slice(0, 3000));
// 把 HTML 存下来供解析
const html = result.find((r) => r.type === 'text/html')?.head ?? '';
if (html) {
  const htmlItem = result.find((r) => r.type === 'text/html');
  const full = await cdp.eval(`(async () => {
    const items = await navigator.clipboard.read();
    for (const it of items) if (it.types.includes('text/html')) return await (await it.getType('text/html')).text();
    return '';
  })()`);
  await fs.writeFile('samples/out/sheet-clip.html', full, 'utf8');
  console.log('\nHTML 片段已存 samples/out/sheet-clip.html');
}
cdp.close();
