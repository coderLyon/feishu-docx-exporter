/**
 * 尝试用「全选 + 复制」把内嵌表格的数据取出来（剪贴板里是 TSV）。
 *
 *   node tools/probe-sheet-clipboard.mjs
 */
import { Cdp, PORT, sleep } from './cdp.mjs';

const cdp = await Cdp.connect(PORT);
const list = await cdp.targets();
const page = list.find((t) => t.type === 'page');
const { sessionId } = await cdp.send('Target.attachToTarget', { targetId: page.id, flatten: true }, null);
cdp.sessionId = sessionId;
await cdp.send('Runtime.enable');
await cdp.send('Page.enable');

// 授权剪贴板读写（浏览器级命令，不能带 sessionId）
await cdp.send('Browser.grantPermissions', {
  origin: new URL(page.url).origin,
  permissions: ['clipboardReadWrite', 'clipboardSanitizedWrite'],
}, null).then(() => console.log('剪贴板权限已授予')).catch((e) => console.log('授权失败:', e.message));

// 1) 找到表格块并点击中心，让它获得焦点
const rect = await cdp.eval(`(() => {
  const b = document.querySelector('[data-block-type="sheet"]');
  if (!b) return null;
  const r = b.getBoundingClientRect();
  return { x: Math.round(r.x + r.width / 2), y: Math.round(r.y + Math.min(r.height / 2, 250)), w: Math.round(r.width), h: Math.round(r.height) };
})()`);
console.log('表格区域:', JSON.stringify(rect));
if (!rect) process.exit(1);

const click = async (x, y) => {
  await cdp.send('Input.dispatchMouseEvent', { type: 'mouseMoved', x, y, button: 'none' });
  await cdp.send('Input.dispatchMouseEvent', { type: 'mousePressed', x, y, button: 'left', clickCount: 1 });
  await sleep(60);
  await cdp.send('Input.dispatchMouseEvent', { type: 'mouseReleased', x, y, button: 'left', clickCount: 1 });
};
await click(rect.x, rect.y);
await sleep(1500);

const key = async (code, keyCode, modifiers = 0) => {
  await cdp.send('Input.dispatchKeyEvent', { type: 'keyDown', windowsVirtualKeyCode: keyCode, nativeVirtualKeyCode: keyCode, code, key: code, modifiers });
  await sleep(50);
  await cdp.send('Input.dispatchKeyEvent', { type: 'keyUp', windowsVirtualKeyCode: keyCode, nativeVirtualKeyCode: keyCode, code, key: code, modifiers });
};

// 2) Ctrl+A 全选，Ctrl+C 复制
await key('KeyA', 65, 2);
await sleep(600);
await key('KeyC', 67, 2);
await sleep(900);

// 3) 读剪贴板
const text = await cdp.eval(`(async () => {
  try { return await navigator.clipboard.readText(); } catch (e) { return 'ERR:' + e.message; }
})()`);
if (typeof text === 'string' && text.startsWith('ERR:')) {
  console.log('读剪贴板失败:', text.slice(0, 120));
  process.exit(1);
}
const lines = String(text).split(/\r?\n/);
console.log(`剪贴板长度 ${text.length}，${lines.length} 行`);
console.log('前 6 行:');
for (const line of lines.slice(0, 6)) {
  console.log(`  ${line.split('\t').length} 列 | ${line.slice(0, 120)}`);
}
cdp.close();
