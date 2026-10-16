/**
 * 采集在线文档基线：定位正文列 → 自上而下滚动 → 分段截图落盘。
 *
 *   node tools/shot-online.mjs [outDir] [--max 40] [--warm]
 *
 * 设计要点：
 *   1. 飞书用内部容器滚动（document 本身不滚），所以 clip 直接用视口坐标即可
 *   2. 懒加载必须「自上而下」走一遍触发，禁止来回跳（虚拟滚动会回收节点）
 *   3. 每张图记录 scrollTop，配合 meta.json 还原文档纵向位置；不做像素拼接
 *      （845 块文档拼成单张长图会到 3 万像素高，8GB 内存下必崩）
 *   4. 截图直接落盘，不在内存里驻留
 */
import fs from 'node:fs/promises';
import path from 'node:path';
import { Cdp, PORT, sleep, pngWidth, pngHeight } from './cdp.mjs';

const argv = process.argv.slice(2);
const ROOT = path.resolve(import.meta.dirname, '..');
// 不写死任何文档：优先用命令行给的 URL（或路径片段），否则自动挑第一个飞书文档页
const DOC_URL_KEY = argv.filter((a) => !a.startsWith('--'))[1] ?? '';

const pos = argv.find((a) => !a.startsWith('--'));
const OUT_DIR = path.resolve(ROOT, pos ?? 'samples/runs/base/online');
const MAX_SHOTS = Number(argv.find((a) => a.startsWith('--max='))?.slice(6) ?? 60);
const WARMUP = argv.includes('--warm');

/** 冻结 UI：隐藏悬浮工具栏/评论/滚动条，固定 hover 态 */
const FREEZE_CSS = `
  .bear-web-x-container .toolbar, .docx-toolbar, .comment-anchor,
  .bear-web-x-container [class*="suspend"], [class*="float-toolbar"] { display:none !important; }
  ::-webkit-scrollbar { display:none !important; }
  * { caret-color: transparent !important; }
`;

const probeLayout = () => {
  // 只统计「宽块」：正文列里的段落/表格都是整列宽（≥300px），
  // 否则中位数会被列表项、行内代码等窄块带偏（实测会把列宽算成 140px）
  const blocks = [...document.querySelectorAll('[data-block-id]')].filter((el) => {
    const r = el.getBoundingClientRect();
    return r.width >= 300 && r.height > 0;
  });
  const rects = blocks.map((el) => el.getBoundingClientRect());
  // 取所有宽块的并集边界：表格通常比正文列更宽，用中位数会把表格裁掉一半
  const minLeft = Math.min(...rects.map((r) => Math.round(r.left)));
  const maxRight = Math.max(...rects.map((r) => Math.round(r.right)));
  const med = (arr) => arr[Math.floor(arr.length / 2)] ?? 0;

  let scroller = null;
  let best = 0;
  document.querySelectorAll('*').forEach((el) => {
    const d = el.scrollHeight - el.clientHeight;
    if (d > 300 && d > best) {
      best = d;
      scroller = el;
    }
  });
  const sr = scroller?.getBoundingClientRect();
  return {
    blockCount: blocks.length,
    columnLeft: minLeft,
    columnWidth: maxRight - minLeft,
    scrollHeight: scroller?.scrollHeight ?? 0,
    clientHeight: scroller?.clientHeight ?? window.innerHeight,
    scrollerTop: sr ? Math.round(sr.top) : 0,
    scrollerHeight: sr ? Math.round(sr.height) : window.innerHeight,
    viewportHeight: window.innerHeight,
    viewportWidth: window.innerWidth,
  };
};

const scrollTo = (y) => {
  let scroller = null;
  let best = 0;
  document.querySelectorAll('*').forEach((el) => {
    const d = el.scrollHeight - el.clientHeight;
    if (d > 300 && d > best) {
      best = d;
      scroller = el;
    }
  });
  if (!scroller) return 0;
  scroller.scrollTop = y;
  return scroller.scrollTop;
};

const readCurrentScroll = () => {
  let scroller = null;
  let best = 0;
  document.querySelectorAll('*').forEach((el) => {
    const d = el.scrollHeight - el.clientHeight;
    if (d > 300 && d > best) {
      best = d;
      scroller = el;
    }
  });
  return scroller?.scrollTop ?? 0;
};

/** 采集当前视口内可见的标题锚点（虚拟渲染下只能逐屏累积） */
const collectAnchors = () => {
  const getScroll = () => {
    let s = null;
    let b = 0;
    document.querySelectorAll('*').forEach((el) => {
      const d = el.scrollHeight - el.clientHeight;
      if (d > 300 && d > b) {
        b = d;
        s = el;
      }
    });
    return s?.scrollTop ?? 0;
  };
  const base = getScroll();
  return [...document.querySelectorAll('.heading-content')]
    .map((el) => ({
      text: (el.textContent ?? '').trim().slice(0, 60),
      y: Math.round(el.getBoundingClientRect().top + base),
    }))
    .filter((a) => a.text.length > 0);
};

async function main() {
  await fs.mkdir(OUT_DIR, { recursive: true });
  const cdp = await Cdp.connect(PORT);
  try {
    if (DOC_URL_KEY) {
      await cdp.attachToPage(DOC_URL_KEY);
    } else {
      // 自动挑第一个飞书文档页（/wiki/ 或 /docx/），保证换文档也能用
      const list = await cdp.targets();
      const hit = list.find((t) => /\/(wiki|docx|docs)\//.test(t.url));
      if (!hit) throw new Error('没有打开的飞书文档页，请在参数里给出文档地址');
      const { sessionId } = await cdp.send('Target.attachToTarget', { targetId: hit.id, flatten: true }, null);
      cdp.sessionId = sessionId;
      cdp.url = hit.url;
      await cdp.send('Page.enable');
      await cdp.send('Runtime.enable');
    }
    const css = FREEZE_CSS;
    await cdp.eval((injectedCss) => {
      const s = document.createElement('style');
      s.id = 'fde-freeze';
      s.textContent = injectedCss;
      document.head.appendChild(s);
    }, [css]);

    const layout = await cdp.eval(probeLayout);
    console.log('版面探测:', JSON.stringify(layout));

    const step = Math.floor(layout.clientHeight * 0.9);
    const total = Math.min(MAX_SHOTS, Math.ceil(layout.scrollHeight / step) || 1);
    console.log(`共 ${total} 段，步长 ${step}px`);

    // 懒加载预热：自上而下走一遍，不截图
    if (WARMUP) {
      for (let i = 0; i < total; i += 1) {
        await cdp.eval(scrollTo, [i * step]);
        await sleep(320);
      }
      process.stdout.write('\n预热完成\n');
      await cdp.eval(scrollTo, [0]);
      await sleep(1500);
    }

    const shots = [];
    const anchorMap = new Map();
    for (let i = 0; i < total; i += 1) {
      const want = i * step;
      const actual = await cdp.eval(scrollTo, [want]);
      await sleep(900);
      const scrollTop = await cdp.eval(readCurrentScroll);
      for (const a of await cdp.eval(collectAnchors)) {
        anchorMap.set(`${a.y}:${a.text}`, a);
      }

      const clip = {
        x: Math.max(0, layout.columnLeft - 8),
        y: Math.max(0, layout.scrollerTop),
        width: Math.min(layout.columnWidth + 16, layout.viewportWidth - Math.max(0, layout.columnLeft - 8)),
        height: Math.min(layout.viewportHeight - Math.max(0, layout.scrollerTop), layout.scrollerHeight),
      };
      const shot = await cdp.screenshot({ clip });
      const file = `${String(i).padStart(3, '0')}.png`;
      await fs.writeFile(path.join(OUT_DIR, file), shot.buffer);
      shots.push({
        file,
        want,
        scrollTop,
        clip,
        pxWidth: pngWidth(shot.buffer),
        pxHeight: pngHeight(shot.buffer),
        scale: shot.scale,
      });
      process.stdout.write(`\r  已采 ${i + 1}/${total}  scale=${shot.scale.toFixed(2)}  `);
      shot.buffer = null;
    }
    console.log('');

    const anchors = [...anchorMap.values()].sort((a, b) => a.y - b.y);
    const meta = {
      capturedAt: new Date().toISOString(),
      url: cdp.url,
      layout,
      step,
      scale: shots[0]?.scale ?? 1,
      effectiveDpi: Math.round(96 * (shots[0]?.scale ?? 1)),
      shots,
      anchors,
    };
    await fs.writeFile(path.join(OUT_DIR, 'meta.json'), JSON.stringify(meta, null, 2));
    console.log(`\n完成 → ${path.relative(ROOT, OUT_DIR)}  (${shots.length} 张, ${meta.effectiveDpi} DPI, 锚点 ${anchors.length} 个)`);
  } finally {
    cdp.close();
  }
}

main().catch((error) => {
  console.error('截图失败:', error.message);
  process.exit(1);
});
