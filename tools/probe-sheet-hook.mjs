/**
 * 探针：验证「扩展内（无 CDP trusted 按键）能否拿到内嵌表格的 HTML」。
 *
 *   node tools/probe-sheet-hook.mjs
 *
 * 做法（全部在页面主世界执行）：
 *   1. hook DataTransfer.setData / navigator.clipboard.write / execCommand('copy')
 *   2. 对内嵌表 canvas 派发合成鼠标序列 → 合成 Ctrl+A / Ctrl+C
 *   3. 看三个出口哪个命中、拿到什么（长度 + 前 80 字）
 *
 * 结论会打印成 PASS / FAIL + 各出口命中情况。
 */
import { Cdp, PORT } from './cdp.mjs';

const cdp = await Cdp.connect(PORT);
const targets = await cdp.targets();
const page = targets.find((t) => t.type === 'page' && /feishu/.test(t.url));
if (!page) throw new Error('没有飞书标签页');
const { sessionId } = await cdp.send('Target.attachToTarget', { targetId: page.id, flatten: true }, null);
cdp.sessionId = sessionId;
await cdp.send('Runtime.enable');

const result = await cdp.eval(`(async () => {
  const seen = [];
  const push = (kind, mime, data) => {
    if (typeof data === 'string' && data.length > 20) seen.push({ kind, mime, len: data.length, head: data.slice(0, 80) });
  };

  // ---- 1) 装 hook ----
  if (!window.__fdeHookInstalled) {
    window.__fdeHookInstalled = true;
    try {
      const origSetData = DataTransfer.prototype.setData;
      DataTransfer.prototype.setData = function (type, data) {
        try { window.__fdeSeen = window.__fdeSeen || []; window.__fdeSeen.push({ kind: 'setData', mime: type, data }); } catch (e) {}
        return origSetData.apply(this, arguments);
      };
    } catch (e) {}
    try {
      const origWrite = navigator.clipboard.write.bind(navigator.clipboard);
      navigator.clipboard.write = async function (items) {
        try {
          for (const item of items) {
            for (const t of item.types) {
              if (!/html|text/.test(t)) continue;
              const blob = await item.getType(t);
              const text = await blob.text();
              window.__fdeSeen = window.__fdeSeen || [];
              window.__fdeSeen.push({ kind: 'clipboard.write', mime: t, data: text });
            }
          }
        } catch (e) {}
        return origWrite(items);
      };
    } catch (e) {}
    try {
      const origExec = document.execCommand.bind(document);
      document.execCommand = function (cmd) {
        if (String(cmd).toLowerCase() === 'copy') {
          try {
            const sel = window.getSelection();
            if (sel && sel.rangeCount) {
              const div = document.createElement('div');
              div.appendChild(sel.getRangeAt(0).cloneContents());
              if (/<table/i.test(div.innerHTML)) {
                window.__fdeSeen = window.__fdeSeen || [];
                window.__fdeSeen.push({ kind: 'execCommand', mime: 'text/html', data: div.innerHTML });
              }
            }
          } catch (e) {}
        }
        return origExec(cmd);
      };
    } catch (e) {}
  }

  // ---- 2) 找到内嵌表 ----
  const blocks = [...document.querySelectorAll('.docx-sheet-block, [data-block-type="sheet"]')];
  const target = blocks.find((b) => b.getBoundingClientRect().height > 60) || blocks[0];
  if (!target) return { error: '页面上找不到内嵌表块（block 数=' + blocks.length + '）' };
  target.scrollIntoView({ block: 'center' });
  await new Promise((r) => setTimeout(r, 600));
  const canvas = target.querySelector('canvas');
  const box = (canvas || target).getBoundingClientRect();
  const x = box.left + Math.min(box.width / 2, 200);
  const y = box.top + Math.min(box.height / 2, 120);

  // ---- 3) 合成鼠标 + 键盘 ----
  const fire = (el, type, extra = {}) =>
    el.dispatchEvent(new MouseEvent(type, { bubbles: true, cancelable: true, view: window, clientX: x, clientY: y, button: 0, ...extra }));
  const el = canvas || target;
  ['mouseover', 'mousemove', 'mousedown', 'mouseup', 'click'].forEach((t) => fire(el, t));
  await new Promise((r) => setTimeout(r, 500));

  const key = (init) => {
    const ev = new KeyboardEvent('keydown', { bubbles: true, cancelable: true, ...init });
    (document.activeElement || document.body).dispatchEvent(ev);
    (document.activeElement || document.body).dispatchEvent(new KeyboardEvent('keyup', { bubbles: true, cancelable: true, ...init }));
  };
  key({ key: 'a', code: 'KeyA', keyCode: 65, which: 65, ctrlKey: true, metaKey: true });
  await new Promise((r) => setTimeout(r, 300));
  key({ key: 'c', code: 'KeyC', keyCode: 67, which: 67, ctrlKey: true, metaKey: true });

  // 等数据
  for (let i = 0; i < 20; i += 1) {
    if ((window.__fdeSeen || []).length) break;
    await new Promise((r) => setTimeout(r, 150));
  }
  // 退出编辑态
  key({ key: 'Escape', code: 'Escape', keyCode: 27, which: 27 });
  await new Promise((r) => setTimeout(r, 200));

  const got = window.__fdeSeen || [];
  const html = got.find((g) => /html/i.test(g.mime));
  const text = got.find((g) => /plain/i.test(g.mime));
  return {
    blocks: blocks.length,
    hits: got.map((g) => ({ kind: g.kind, mime: g.mime, len: g.data.length })),
    hasHtml: Boolean(html && /<table/i.test(html.data)),
    hasTsv: Boolean(text),
    htmlHead: html ? html.data.slice(0, 100) : '',
    tsvHead: text ? text.data.slice(0, 100) : '',
  };
})()`);

console.log(JSON.stringify(result, null, 1));
cdp.close();
