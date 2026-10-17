/**
 * 隔离实验：不调用扩展导出，只按「预热滚动」的节奏把文档滚一遍，
 * 观察页面是否自己变卡（用于判断卡死是飞书渲染的问题还是采集代码的问题）。
 *
 *   node tools/probe-scroll-cost.mjs
 */
import { Cdp, PORT, sleep } from './cdp.mjs';

const cdp = await Cdp.connect(PORT);
const list = await cdp.targets();
const page = list.find((t) => /\/wiki\/|\/docx\//.test(t.url));
const { sessionId } = await cdp.send('Target.attachToTarget', { targetId: page.id, flatten: true }, null);
cdp.sessionId = sessionId;
await cdp.send('Runtime.enable');

const meta = await cdp.eval(`(() => {
  const s = document.querySelector('.bear-web-x-container') || document.scrollingElement;
  return { scrollHeight: s.scrollHeight, clientHeight: s.clientHeight, blocks: document.querySelectorAll('[data-block-id]').length };
})()`);
console.log('初始:', JSON.stringify(meta));

const step = Math.max(240, Math.floor(meta.clientHeight * 0.5));
const total = meta.scrollHeight;
console.log(`半步长 ${step}px，预计 ${Math.ceil(total / step)} 步`);

for (let y = 0, i = 1; y <= total; y += step, i += 1) {
  const t0 = Date.now();
  try {
    const r = await cdp.eval(`(() => {
      const s = document.querySelector('.bear-web-x-container') || document.scrollingElement;
      s.scrollTop = ${y};
      return { top: s.scrollTop, sh: s.scrollHeight, blocks: document.querySelectorAll('[data-block-id]').length,
               heap: performance.memory ? Math.round(performance.memory.usedJSHeapSize / 1048576) : -1 };
    })()`);
    const rtt = Date.now() - t0;
    if (i % 5 === 0 || rtt > 3000) {
      console.log(`步 ${i}: y=${r.top}/${r.sh} blocks=${r.blocks} heap=${r.heap}MB 往返=${rtt}ms`);
    }
  } catch (e) {
    console.log(`步 ${i} 失败: ${e.message}`);
    break;
  }
  await sleep(450);
}
console.log('滚动完成（未触发导出）');
cdp.close();
