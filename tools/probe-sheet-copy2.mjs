/**
 * 测试「页面内合成按键 + execCommand('copy') + 读剪贴板」是否可行
 * （这是扩展里能落地的路径：不依赖 CDP 的 trusted 输入）。
 *
 *   node tools/probe-sheet-copy2.mjs
 */
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

// 只在页内操作：合成点击 → 合成 Ctrl+A → execCommand('copy') → 读剪贴板
const result = await cdp.eval(`(async () => {
  const log = [];
  const b = document.querySelector('[data-block-type="sheet"]');
  if (!b) return { err: '无表格块' };
  const r = b.getBoundingClientRect();
  const cx = r.x + r.width / 2;
  const cy = r.y + 200;

  // 1) 合成点击（在页内派发，非 CDP trusted 输入）
  const target = document.elementFromPoint(cx, cy);
  log.push('命中元素: ' + (target ? target.tagName + '.' + String(target.className).slice(0, 28) : '无'));
  for (const type of ['mousedown', 'mouseup', 'click']) {
    target.dispatchEvent(new MouseEvent(type, { bubbles: true, cancelable: true, clientX: cx, clientY: cy, button: 0 }));
  }
  await new Promise(r2 => setTimeout(r2, 900));
  log.push('activeElement: ' + (document.activeElement?.tagName || '?'));

  // 2) 合成 Ctrl+A
  const el = document.activeElement || document.body;
  for (const type of ['keydown', 'keyup']) {
    el.dispatchEvent(new KeyboardEvent(type, { key: 'a', code: 'KeyA', keyCode: 65, ctrlKey: true, bubbles: true, cancelable: true }));
  }
  await new Promise(r2 => setTimeout(r2, 500));

  // 3) execCommand('copy')
  let ok = false;
  try { ok = document.execCommand('copy'); } catch (e) { log.push('execCommand 异常: ' + e.message); }
  log.push('execCommand(copy) = ' + ok);
  await new Promise(r2 => setTimeout(r2, 800));

  // 4) 读剪贴板
  let text = '';
  try { text = await navigator.clipboard.readText(); } catch (e) { log.push('读剪贴板失败: ' + e.message); }
  const lines = text.split(/\\r?\\n/);
  return { log, len: text.length, lineCount: lines.length, head: lines.slice(0, 3).map(l => l.slice(0, 100)) };
})()`);

console.log(JSON.stringify(result, null, 1).slice(0, 1800));
cdp.close();
