/**
 * 通过「点进表格 → 全选 → 复制 → 读剪贴板」拿到内嵌电子表格的真实单元格数据，
 * 写入页面 localStorage，供内容脚本在采集阶段取用（同源共享）。
 *
 *   node tools/extract-sheets.mjs [--keep-clipboard]
 *
 * 为什么走剪贴板：飞书表格是 canvas 渲染 + 稀疏 protobuf 快照，
 * 直接在二进制里还原网格风险极高（错位比截图更糟）；而「复制」是飞书自己
 * 实现的序列化，剪贴板里就是权威的 TSV（含空单元格，与网格一一对应）。
 *
 * 局限：依赖 CDP 的 trusted 按键（扩展内无法发送），所以目前只用于自动化导出链路。
 */
import { Cdp, PORT, sleep } from './cdp.mjs';

const cdp = await Cdp.connect(PORT);
const list = await cdp.targets();
const page = list.find((t) => t.type === 'page' && /feishu/.test(t.url));
if (!page) throw new Error('找不到飞书文档页');
const { sessionId } = await cdp.send('Target.attachToTarget', { targetId: page.id, flatten: true }, null);
cdp.sessionId = sessionId;
await cdp.send('Runtime.enable');
await cdp.send('Page.enable');
await cdp.send('Browser.grantPermissions', {
  origin: new URL(page.url).origin,
  permissions: ['clipboardReadWrite', 'clipboardSanitizedWrite'],
}, null).catch(() => {});

const key = async (code, keyCode, mods = 0) => {
  await cdp.send('Input.dispatchKeyEvent', { type: 'keyDown', windowsVirtualKeyCode: keyCode, nativeVirtualKeyCode: keyCode, code, key: code, modifiers: mods });
  await sleep(60);
  await cdp.send('Input.dispatchKeyEvent', { type: 'keyUp', windowsVirtualKeyCode: keyCode, nativeVirtualKeyCode: keyCode, code, key: code, modifiers: mods });
};
const clickAt = async (x, y) => {
  await cdp.send('Input.dispatchMouseEvent', { type: 'mouseMoved', x, y, button: 'none' });
  await sleep(80);
  await cdp.send('Input.dispatchMouseEvent', { type: 'mousePressed', x, y, button: 'left', clickCount: 1 });
  await sleep(60);
  await cdp.send('Input.dispatchMouseEvent', { type: 'mouseReleased', x, y, button: 'left', clickCount: 1 });
};

// 飞书虚拟化：先慢速滚一遍让所有块挂载，否则深层的表格根本不在 DOM 里
console.log('预热滚动，等待表格挂载…');
await cdp.eval(`(async () => {
  const s = document.querySelector('.bear-web-x-container') || document.scrollingElement;
  const step = Math.max(300, Math.floor(s.clientHeight * 0.6));
  for (let y = 0; y <= s.scrollHeight; y += step) {
    s.scrollTop = y;
    await new Promise((r) => setTimeout(r, 420));
  }
  s.scrollTop = 0;
  await new Promise((r) => setTimeout(r, 600));
})()`).catch(() => {});
await sleep(1500);

// 虚拟化：表格块只有滚到时才挂载。边滚边处理，处理完按 data-block-id 存数据。
const SEL = '[data-block-type="sheet"], .docx-sheet-block';
const done = new Set();
let processed = 0;

async function handleVisibleSheets() {
  const list = await cdp.eval(`(() => {
    return [...document.querySelectorAll(${JSON.stringify(SEL)})].map((el) => {
      const r = el.getBoundingClientRect();
      return {
        id: el.getAttribute('data-block-id') || '',
        x: Math.round(r.x + r.width / 2),
        y: Math.round(Math.max(80, Math.min(r.y + 120, window.innerHeight - 80))),
        w: Math.round(r.width), h: Math.round(r.height),
        // 表格常常比视口还高，不能要求「完整可见」；只要有足够区域可点就处理
        inView: r.top < window.innerHeight - 120 && r.bottom > 80,
      };
    }).filter((s2) => s2.id && s2.inView);
  })()`);
  for (const box of list) {
    if (done.has(box.id)) continue;
    done.add(box.id);
    console.log(`  块 ${box.id}: ${box.w}×${box.h} @(${box.x},${box.y})`);
    await clickAt(box.x, box.y);
    await sleep(1400);
    await key('KeyA', 65, 2);
    await sleep(650);
    await key('KeyC', 67, 2);
    await sleep(1100);
    const tsv = await cdp.eval(`(async () => { try { return await navigator.clipboard.readText(); } catch { return ''; } })()`);
    const lines = String(tsv).replace(/\r/g, '').split('\n').filter((l) => l.length);
    if (!lines.length) { console.log('    ✗ 剪贴板为空'); continue; }
    console.log(`    ✓ ${lines.length} 行 × ${lines[0].split('\t').length} 列`);
    await cdp.eval(`(() => { localStorage.setItem('fde-sheet-' + ${JSON.stringify(box.id)}, ${JSON.stringify(tsv)}); return true; })()`);
    processed += 1;
  }
}

const geo = await cdp.eval(`(() => {
  const s = document.querySelector('.bear-web-x-container') || document.scrollingElement;
  return { sh: s.scrollHeight, ch: s.clientHeight };
})()`);
const step = Math.max(280, Math.floor(geo.ch * 0.55));
for (let y = 0; y <= geo.sh; y += step) {
  await cdp.eval(`(() => { const s = document.querySelector('.bear-web-x-container') || document.scrollingElement; s.scrollTop = ${y}; return s.scrollTop; })()`);
  await sleep(430);
  await handleVisibleSheets();
}
// 到底后再补一轮（有些块在最后才挂载）
await handleVisibleSheets();

// 退出表格编辑态：否则后续导出的页面负载会大幅上升（实测会让采集超时）
await key('Escape', 27);
await sleep(300);
await cdp.eval(`(() => {
  const s = document.querySelector('.bear-web-x-container') || document.scrollingElement;
  s.scrollTop = 0;
  const el = document.querySelector('[data-block-type="text"]');
  if (el) el.dispatchEvent(new MouseEvent('mousedown', { bubbles: true }));
  return true;
})()`).catch(() => {});
await sleep(800);

console.log(`完成：${processed} 个表格的数据已写入 localStorage（键名 fde-sheet-<blockId>）`);
cdp.close();
