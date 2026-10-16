/** 快速探查当前飞书页面状态（是否已加载、content script 是否注入） */
import { Cdp, PORT } from './cdp.mjs';

const KEY = process.argv[2] ?? '';

const cdp = await Cdp.connect(PORT);
try {
  await cdp.attachToPage(KEY);
  const info = await cdp.eval(() => {
    let scroller = null;
    let best = 0;
    document.querySelectorAll('*').forEach((el) => {
      const d = el.scrollHeight - el.clientHeight;
      if (d > 300 && d > best) {
        best = d;
        scroller = el;
      }
    });
    return {
      url: location.href.slice(0, 80),
      title: document.title,
      blocks: document.querySelectorAll('[data-block-id]').length,
      headings: document.querySelectorAll('.heading-content').length,
      ready: document.documentElement.getAttribute('data-fde-ready'),
      scrollHeight: scroller?.scrollHeight ?? 0,
      clientHeight: scroller?.clientHeight ?? 0,
      firstText: (document.body.innerText || '').slice(0, 150).replace(/\n/g, ' | '),
    };
  });
  console.log(JSON.stringify(info, null, 2));
} finally {
  cdp.close();
}
