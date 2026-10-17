/**
 * 尝试不碰系统剪贴板，直接拦截飞书 sheet 的 copy 事件拿到 TSV。
 *
 *   node tools/probe-sheet-copyevent.mjs
 *
 * 原理：点进表格后派发一个合成的 copy 事件（带空 DataTransfer），
 * 若飞书的复制处理器监听在 document 上并填充 clipboardData，
 * 就能在事件回调里直接读到 TSV，无需系统剪贴板与额外权限。
 */
import { Cdp, PORT, sleep } from './cdp.mjs';

const cdp = await Cdp.connect(PORT);
const list = await cdp.targets();
const page = list.find((t) => t.type === 'page');
const { sessionId } = await cdp.send('Target.attachToTarget', { targetId: page.id, flatten: true }, null);
cdp.sessionId = sessionId;
await cdp.send('Runtime.enable');
await cdp.send('Page.enable');

const rect = await cdp.eval(`(() => {
  const b = document.querySelector('[data-block-type="sheet"]');
  const r = b.getBoundingClientRect();
  return { x: Math.round(r.x + r.width / 2), y: Math.round(r.y + 200) };
})()`);

// 先把焦点给表格
await cdp.send('Input.dispatchMouseEvent', { type: 'mousePressed', x: rect.x, y: rect.y, button: 'left', clickCount: 1 });
await sleep(60);
await cdp.send('Input.dispatchMouseEvent', { type: 'mouseReleased', x: rect.x, y: rect.y, button: 'left', clickCount: 1 });
await sleep(1200);

// 装一个监听器，记录所有 copy 事件带来的数据
const result = await cdp.eval(`(async () => {
  window.__fdeCopied = [];
  const onCopy = (e) => {
    try {
      const t = e.clipboardData ? e.clipboardData.getData('text/plain') : '';
      window.__fdeCopied.push({ len: (t || '').length, head: (t || '').slice(0, 120) });
    } catch (err) {
      window.__fdeCopied.push({ err: String(err) });
    }
  };
  document.addEventListener('copy', onCopy, true);

  // 派发合成 copy 事件（带 DataTransfer）
  const dt = new DataTransfer();
  const ev = new ClipboardEvent('copy', { clipboardData: dt, bubbles: true, cancelable: true });
  document.dispatchEvent(ev);
  await new Promise(r => setTimeout(r, 600));

  // 也试一下在活动元素上派发
  const active = document.activeElement;
  if (active && active !== document.body) {
    const dt2 = new DataTransfer();
    active.dispatchEvent(new ClipboardEvent('copy', { clipboardData: dt2, bubbles: true, cancelable: true }));
    await new Promise(r => setTimeout(r, 400));
  }

  document.removeEventListener('copy', onCopy, true);
  return {
    events: window.__fdeCopied,
    dtData: (() => { try { return dt.getData('text/plain').slice(0, 120); } catch { return '(读取失败)'; } })(),
    activeTag: active ? active.tagName + '.' + String(active.className).slice(0, 30) : null,
  };
})()`);

console.log(JSON.stringify(result, null, 1).slice(0, 1600));
cdp.close();
