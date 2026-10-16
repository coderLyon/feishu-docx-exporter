/**
 * 探测飞书「画板下载为图片」接口（r005 批 0）。
 *
 *   node tools/probe-wb.mjs [--click]
 *
 * 与上一版的区别：不靠类名猜画板位置，而是**先用 client_vars 拿到画板块的
 * block_id / token，再到 DOM 里按 [data-block-id] 精确定位**，滚过去之后再
 * hover 出画板工具条，点「下载为图片」，用 Network 域抓请求。
 *
 * 判定：
 *   - 响应 Content-Type 为 image/*  → 二进制直返，capture 可直接 fetch
 *   - 响应是 JSON { ticket, url }   → 需要二次换票
 *   - 什么都没抓到                  → 走 capture.js 的「只合成一层 canvas」降级路径
 *
 * 探测失败属于预期内，不影响其余 7 项需求。
 */
import fs from 'node:fs';
import path from 'node:path';
import { Cdp, PORT, sleep } from './cdp.mjs';

const ROOT = path.resolve(import.meta.dirname, '..');
const OUT = path.join(ROOT, 'samples', 'wb-probe.json');
const DOC_KEY = process.env.DOC_KEY ?? '';
const DO_CLICK = process.argv.includes('--click');

/** 页面内：分页拉全量块，返回画板/流程图/思维导图块的 id 与 token */
const FETCH_BOARDS = `(async () => {
  const wiki = '${DOC_KEY}';
  const g = await fetch('/space/api/wiki/v2/tree/get_node/?wiki_token=' + wiki, { credentials: 'include' }).then(r => r.json());
  const objToken = (g?.data?.node ?? g?.data)?.obj_token;
  if (!objToken) return { err: 'no obj_token' };
  const map = {};
  let cursor = '';
  for (let i = 0; i < 30; i++) {
    const qs = new URLSearchParams({ id: objToken, mode: '7', limit: '500', container_type: 'wiki2.0', container_id: wiki });
    if (cursor) qs.set('cursor', cursor);
    const j = await fetch('/space/api/docx/pages/client_vars?' + qs, { credentials: 'include' }).then(r => r.json());
    Object.assign(map, j?.data?.block_map ?? {});
    if (!j?.data?.has_more || !j?.data?.cursor) break;
    cursor = j.data.cursor;
  }
  const boards = [];
  for (const [id, b] of Object.entries(map)) {
    const t = b?.data?.type ?? '';
    if (!/whiteboard|diagram|mindnote|board/i.test(t)) continue;
    boards.push({ id, type: t, token: b?.data?.token ?? '', w: b?.data?.width ?? 0, h: b?.data?.height ?? 0 });
  }
  return { objToken, total: Object.keys(map).length, boards };
})()`;

/** 飞书文档真正的滚动容器不是 window，而是 .bear-web-x-container 这类 div */
const FIND_SCROLLER = `(() => {
  let best = null;
  document.querySelectorAll('*').forEach((el) => {
    const delta = el.scrollHeight - el.clientHeight;
    if (delta > 400 && (!best || delta > best.delta)) best = { delta, el };
  });
  (best?.el ?? document.scrollingElement).setAttribute('data-fde-probe-scroller', '1');
  return String(best?.el?.className ?? '').slice(0, 60);
})()`;

const HOVER_AT = (x, y) => `(() => {
  const el = document.elementFromPoint(${x}, ${y});
  if (!el) return 'no-element';
  for (const type of ['mouseover', 'mouseenter', 'mousemove']) {
    el.dispatchEvent(new MouseEvent(type, { bubbles: true, clientX: ${x}, clientY: ${y}, view: window }));
  }
  return String(el.className).slice(0, 80);
})()`;

/** 页面内：列出所有与「下载 / 导出 / 图片 / 更多」有关的小元素 */
const FIND_MENU = `(() => {
  const hit = [];
  for (const el of document.querySelectorAll('*')) {
    if (el.children.length > 2) continue;
    const t = (el.textContent || '').trim();
    if (!t || t.length > 24) continue;
    if (!/下载|导出|图片|另存|复制/.test(t)) continue;
    const r = el.getBoundingClientRect();
    if (r.width < 8 || r.height < 8) continue;
    hit.push({ tag: el.tagName, cls: String(el.className).slice(0, 90), text: t, x: Math.round(r.x + r.width / 2), y: Math.round(r.y + r.height / 2) });
  }
  return hit.slice(0, 40);
})()`;

const CLICK_AT = (x, y) => `(() => {
  const el = document.elementFromPoint(${x}, ${y});
  if (!el) return 'no-element';
  el.dispatchEvent(new MouseEvent('mousedown', { bubbles: true }));
  el.dispatchEvent(new MouseEvent('mouseup', { bubbles: true }));
  el.click();
  return (el.textContent || '').trim().slice(0, 30) || el.tagName;
})()`;

async function main() {
  const cdp = await Cdp.connect(PORT);
  const byId = new Map();

  cdp.listeners.push((msg) => {
    if (msg.method === 'Network.requestWillBeSent') {
      const r = msg.params.request;
      byId.set(msg.params.requestId, {
        url: r.url, method: r.method, type: msg.params.type,
        postData: r.postData?.slice(0, 600) ?? '',
      });
    }
    if (msg.method === 'Network.responseReceived') {
      const rec = byId.get(msg.params.requestId);
      if (rec) {
        rec.status = msg.params.response.status;
        rec.mime = msg.params.response.mimeType;
      }
    }
  });

  try {
    await cdp.attachToPage(DOC_KEY);
    await cdp.send('Network.enable', { maxPostDataSize: 8192 });
    console.log(`已连接: ${cdp.url}`);

    const meta = await cdp.eval(FETCH_BOARDS);
    if (meta.err) throw new Error(meta.err);
    console.log(`块总数 ${meta.total}，画板块 ${meta.boards.length} 个`);
    meta.boards.forEach((b, i) => console.log(`  #${i} ${b.type} ${b.id} token=${b.token || '-'} ${b.w}x${b.h}`));

    // 飞书虚拟化渲染：块只有滚到视口附近才进入 DOM，必须一遍滚一遍找。
    // 实测只设容器的 scrollTop 不够，同时 window.scrollTo 并给足 1.3s 才会触发懒渲染。
    const located = new Map();
    await cdp.eval(FIND_SCROLLER);
    const total = await cdp.eval(`document.querySelector('[data-fde-probe-scroller="1"]').scrollHeight`);
    for (let y = 0; y < Number(total) + 20000; y += 600) {
      await cdp.eval(`(() => {
        const el = document.querySelector('[data-fde-probe-scroller="1"]');
        el.scrollTop = ${y}; window.scrollTo(0, ${y});
        return document.querySelector('[data-fde-probe-scroller="1"]').scrollHeight;
      })()`);
      await sleep(900);
      const found = await cdp.eval(
        `(() => [${meta.boards.map((b) => `"${b.id}"`).join(',')}].map((id) => {
          const el = document.querySelector('[data-block-id="' + id + '"]');
          if (!el) return null;
          const r = el.getBoundingClientRect();
          return { id, w: Math.round(r.width), h: Math.round(r.height), x: Math.round(r.x + r.width / 2), y: Math.round(r.y + r.height / 2), cls: String(el.className).slice(0, 70) };
        }).filter(Boolean))()`,
      );
      for (const f of found ?? []) {
        if (!located.has(f.id)) {
          located.set(f.id, f);
          console.log(`  命中 ${f.id} @${f.x},${f.y} ${f.w}x${f.h} ${f.cls}`);
        }
      }
      if (located.size === meta.boards.length) break;
      if (y > Number(total)) {
        const grown = await cdp.eval(`document.querySelector('[data-fde-probe-scroller="1"]').scrollHeight`);
        if (grown <= Number(total)) break;
      }
    }
    if (!located.size) console.log('  未能在 DOM 中定位任何画板块（虚拟滚动未渲染到）');

    let click = null;
    const before = byId.size;
    if (DO_CLICK && located.length) {
      const target = [...located.values()][0];
      await cdp.eval(HOVER_AT(target.x, target.y));
      await sleep(1200);
      const menu = await cdp.eval(FIND_MENU);
      console.log(`候选菜单 ${menu.length} 项`);
      menu.forEach((m) => console.log(`   "${m.text}" @${m.x},${m.y} ${m.cls}`));

      const item = menu.find((m) => /下载为图片|下载图片|导出为图片|导出图片/.test(m.text))
        ?? menu.find((m) => /下载|导出/.test(m.text));
      if (item) {
        const hit = await cdp.eval(CLICK_AT(item.x, item.y));
        console.log(`已点「${item.text}」-> ${hit}`);
        await sleep(4000);
        click = { item, hit };
      } else {
        console.log('未找到「下载为图片」菜单项');
      }
    }

    const all = [...byId.values()].filter(Boolean);
    const interesting = all.filter(
      (r) => /board|whiteboard|export|download|image|drive|ticket|attachment/i.test(r.url)
        && !/\.(js|css|woff2?|ttf|svg|png)(\?|$)/i.test(r.url),
    );

    const payload = {
      doneAt: new Date().toISOString(),
      objToken: meta.objToken,
      boards: [...located.values()],
      click,
      newRequests: all.length - before,
      interesting: interesting.map((r) => ({
        method: r.method, url: r.url.slice(0, 300), type: r.type, status: r.status, mime: r.mime, postData: r.postData,
      })),
      totalRequests: all.length,
    };
    fs.writeFileSync(OUT, JSON.stringify(payload, null, 2));
    console.log(`\n落盘 -> ${path.relative(ROOT, OUT)}  总请求 ${all.length}，命中 ${interesting.length}`);
    interesting.slice(0, 30).forEach((r) => console.log(`   ${r.method} ${r.status ?? '-'} ${r.mime ?? '-'} ${r.url.slice(0, 150)}`));
  } finally {
    cdp.close();
  }
}

main().catch((e) => {
  console.error('探测失败:', e.message);
  process.exit(1);
});
