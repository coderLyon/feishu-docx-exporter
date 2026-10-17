/**
 * 极简 CDP 客户端（不依赖 Playwright）。
 *
 * 之所以不用 Playwright：本机只有 8GB 内存，connectOverCDP 需要拉起完整的
 * browser/context 抽象，内存与启动开销都远大于直连 WebSocket。而我们要做的
 * 只是「eval 一段 JS」和「截一张图」，原始 CDP 足够。
 */
import os from 'node:os';
import path from 'node:path';

export const PORT = Number(process.env.FDE_DEBUG_PORT ?? 9222);

const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

export class Cdp {
  constructor(ws, port) {
    this.ws = ws;
    this.port = port;
    this.id = 0;
    this.pending = new Map();
    this.sessionId = null;
    this.listeners = [];
  }

  static async connect(port = PORT, timeout = 15000) {
    const res = await fetch(`http://127.0.0.1:${port}/json/version`, { cache: 'no-store' });
    if (!res.ok) throw new Error(`调试端口 ${port} 返回 HTTP ${res.status}`);
    const version = await res.json();
    const ws = new WebSocket(version.webSocketDebuggerUrl);
    const cdp = new Cdp(ws, port);
    await new Promise((resolve, reject) => {
      const timer = setTimeout(() => reject(new Error('WebSocket 连接超时')), timeout);
      ws.onopen = () => {
        clearTimeout(timer);
        resolve();
      };
      ws.onerror = (e) => {
        clearTimeout(timer);
        reject(new Error(`WebSocket 失败: ${e?.message ?? e}`));
      };
    });
    ws.onmessage = (event) => cdp.#onMessage(event.data);
    return cdp;
  }

  #onMessage(raw) {
    let msg;
    try {
      msg = JSON.parse(raw);
    } catch {
      return;
    }
    if (msg.id && this.pending.has(msg.id)) {
      const { resolve, reject } = this.pending.get(msg.id);
      this.pending.delete(msg.id);
      if (msg.error) reject(new Error(`${msg.error.message} (code ${msg.error.code})`));
      else resolve(msg.result);
      return;
    }
    for (const fn of this.listeners) fn(msg);
  }

  send(method, params = {}, sessionId = this.sessionId) {
    const id = (this.id += 1);
    const payload = { id, method, params };
    if (sessionId) payload.sessionId = sessionId;
    return new Promise((resolve, reject) => {
      this.pending.set(id, { resolve, reject });
      this.ws.send(JSON.stringify(payload));
      setTimeout(() => {
        if (this.pending.has(id)) {
          this.pending.delete(id);
          reject(new Error(`CDP ${method} 超时`));
        }
      }, 180000); // 大文档采集期间主线程可能被长时间占用，60s 容易误判超时
    });
  }

  /** 列出所有 page 类型的 target */
  async targets() {
    const res = await fetch(`http://127.0.0.1:${this.port}/json/list`, { cache: 'no-store' });
    const list = await res.json();
    return list.filter((t) => t.type === 'page');
  }

  /** 按 URL 片段找到目标页并 attach，返回 sessionId */
  async attachToPage(urlFragment) {
    const list = await this.targets();
    const target = list.find((t) => (urlFragment ? t.url.includes(urlFragment) : true));
    if (!target) throw new Error(`未找到匹配「${urlFragment}」的页面，当前 ${list.length} 个标签页`);
    const { sessionId } = await this.send('Target.attachToTarget', {
      targetId: target.id,
      flatten: true,
    });
    this.sessionId = sessionId;
    this.targetId = target.id;
    this.url = target.url;
    await this.send('Page.enable');
    await this.send('Runtime.enable');
    return sessionId;
  }

  /**
   * 在页面主世界执行 JS。
   * @param {string|Function} fn 函数（会被转成 IIFE 并内联参数）或完整表达式字符串
   * @param {Array} args 传给函数的参数（会被 JSON 序列化后内联）
   */
  async eval(fn, args = []) {
    // 注意：字符串按「完整表达式」原样执行，绝不能再包一层——
    // 否则 "() => ..." 会被包成「返回函数的函数」，返回值永远是 undefined
    const expression =
      typeof fn === 'function' ? `(${fn.toString()})(...${JSON.stringify(args)})` : fn;
    const result = await this.send('Runtime.evaluate', {
      expression,
      awaitPromise: true,
      returnByValue: true,
    });
    if (result.exceptionDetails) {
      const ex = result.exceptionDetails.exception?.description ?? result.exceptionDetails.text;
      throw new Error(`页面内执行失败: ${ex}`);
    }
    return result.result?.value;
  }

  /** 截当前视口（可选裁剪），返回 { buffer, width, height, scale } */
  async screenshot({ clip, format = 'png' } = {}) {
    const params = { format };
    if (clip) params.clip = { ...clip, scale: clip.scale ?? 1 };
    const result = await this.send('Page.captureScreenshot', params);
    const buffer = Buffer.from(result.data, 'base64');
    // 用图片实际尺寸反推真实 dpr，避免依赖系统显示缩放
    const width = pngWidth(buffer);
    const scale = clip ? width / clip.width : 1;
    return { buffer, width, height: pngHeight(buffer), scale };
  }

  close() {
    try {
      this.ws.close();
    } catch {
      /* 忽略 */
    }
  }
}

/** 读 PNG 头拿宽高（无需图像库） */
function pngSize(buffer) {
  if (buffer.readUInt32BE(0) !== 0x89504e47) return { width: 0, height: 0 };
  return { width: buffer.readUInt32BE(16), height: buffer.readUInt32BE(20) };
}
export const pngWidth = (b) => pngSize(b).width;
export const pngHeight = (b) => pngSize(b).height;

/**
 * Chrome 硬性限制：默认 profile 不允许开远程调试
 * （实测报错 "DevTools remote debugging requires a non-default data directory"）。
 * 因此必须用独立目录；需要登录态时从默认 profile 定向复制 Cookie/LocalStorage。
 */
export const DEFAULT_PROFILE = path.join(
  os.homedir(),
  'AppData',
  'Local',
  'Google',
  'Chrome',
  'User Data',
);

export const FDE_PROFILE = path.join(
  os.homedir(),
  'CodeBuddy',
  '20261014140311',
  '.chrome-user',
);

export const CHROME_CANDIDATES = [
  'C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe',
  'C:\\Program Files (x86)\\Google\\Chrome\\Application\\chrome.exe',
];

export { sleep };
