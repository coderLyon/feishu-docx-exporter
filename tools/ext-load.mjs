/**
 * 用 CDP 的 Extensions.loadUnpacked 装载 dist/。
 *
 * Chrome 137+ 禁用了 --load-extension 命令行参数，官方替代方案是
 * CDP 的 Extensions.loadUnpacked（需配合 --enable-unsafe-extension-debugging）。
 */
import path from 'node:path';
import { ROOT, PORT } from './browser.mjs';

const DIST = path.join(ROOT, 'dist');

class Cdp {
  constructor(ws) {
    this.ws = ws;
    this.nextId = 0;
    this.pending = new Map();
    ws.addEventListener('message', (event) => {
      const message = JSON.parse(event.data);
      const slot = this.pending.get(message.id);
      if (!slot) return;
      this.pending.delete(message.id);
      if (message.error) slot.reject(new Error(`${message.error.message}`));
      else slot.resolve(message.result);
    });
  }

  static async connect() {
    const version = await (await fetch(`http://127.0.0.1:${PORT}/json/version`)).json();
    const ws = new WebSocket(version.webSocketDebuggerUrl);
    await new Promise((resolve, reject) => {
      ws.addEventListener('open', resolve, { once: true });
      ws.addEventListener('error', reject, { once: true });
    });
    return new Cdp(ws);
  }

  send(method, params = {}) {
    const id = (this.nextId += 1);
    return new Promise((resolve, reject) => {
      this.pending.set(id, { resolve, reject });
      this.ws.send(JSON.stringify({ id, method, params }));
      setTimeout(() => {
        if (this.pending.delete(id)) reject(new Error(`${method} 超时`));
      }, 15000);
    });
  }

  close() {
    this.ws.close();
  }
}

/** 装载（或重新装载）扩展，返回扩展 id */
export async function loadExtension() {
  const cdp = await Cdp.connect();
  try {
    const result = await cdp.send('Extensions.loadUnpacked', { path: DIST });
    console.log(`扩展已装载: ${result.id}`);
    return result.id;
  } finally {
    cdp.close();
  }
}

/** 列出已装载的扩展，便于确认状态 */
export async function listExtensions() {
  const cdp = await Cdp.connect();
  try {
    const result = await cdp.send('Extensions.getForBrowser');
    return result ?? {};
  } finally {
    cdp.close();
  }
}

import { pathToFileURL } from 'node:url';

if (import.meta.url === pathToFileURL(process.argv[1] ?? '').href) {
  loadExtension().catch((error) => {
    console.error(`装载失败: ${String(error.message ?? error)}`);
    process.exitCode = 1;
  });
}
