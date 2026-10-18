/**
 * 内嵌表格「实时提取」：让后台用 chrome.debugger 发 trusted 输入，
 * 点进表格 → Ctrl+A / Ctrl+C → 读剪贴板 → 返回 HTML/TSV。
 *
 * 这是扩展内拿到单元格数据的唯一可行方式。实测（tools/probe-sheet-hook.mjs）：
 * 合成的 MouseEvent / PointerEvent / 合成 copy 事件全部零命中 ——
 * 飞书的 canvas 表格只认真实输入。
 *
 * 代价：调用期间浏览器顶部会出现「正在调试此浏览器」提示条。
 * 失败一律回退，不阻塞导出。
 */

/** 后台是否已获得 debugger 权限 */
export async function hasDebuggerPermission() {
  try {
    const res = await chrome.runtime.sendMessage({ type: 'has-debugger' });
    return Boolean(res?.ok);
  } catch {
    return false;
  }
}

/** 申请 debugger 权限（必须在用户手势里调用，例如点「开始导出」） */
export async function requestDebuggerPermission() {
  try {
    const res = await chrome.runtime.sendMessage({ type: 'request-debugger' });
    return Boolean(res?.ok);
  } catch {
    return false;
  }
}

export async function ensureDebuggerPermission() {
  try {
    if (await hasDebuggerPermission()) return true;
    const res = await chrome.runtime.sendMessage({ type: 'request-debugger' });
    return Boolean(res?.ok);
  } catch {
    return false;
  }
}

/**
 * 依次提取多张内嵌表的单元格数据。
 *
 * @param {string[]} blockIds 内嵌表块的 data-block-id（虚拟化下必须按 id 重查元素）
 * @param {{onProgress?: (done:number,total:number)=>void, timeoutMs?:number}} [opts]
 * @returns {Promise<Array<{html:string|null, tsv:string|null, error:string|null}>>}
 */
export async function extractSheetsLive(blockIds, { onProgress, timeoutMs = 12000 } = {}) {
  const out = [];
  for (let i = 0; i < blockIds.length; i += 1) {
    onProgress?.(i, blockIds.length);
    const id = blockIds[i];
    let block = document.querySelector(`[data-block-id="${id}"]`);
    if (!block) {
      out.push({ html: null, tsv: null, error: 'block 未挂载' });
      continue;
    }
    // 滚到视口中央：debugger 的点击坐标是视口坐标
    try {
      block.scrollIntoView({ block: 'center' });
    } catch {
      /* 忽略 */
    }
    await new Promise((r) => setTimeout(r, 450));
    block = document.querySelector(`[data-block-id="${id}"]`) ?? block;

    const canvas = block.querySelector('canvas') ?? block;
    const box = canvas.getBoundingClientRect();
    const rect = {
      x: box.left + Math.min(box.width / 2, 150),
      y: box.top + Math.min(box.height / 2, 70),
    };

    try {
      const res = await Promise.race([
        chrome.runtime.sendMessage({ type: 'extract-sheet', rect }),
        new Promise((r) => setTimeout(() => r({ ok: false, error: 'timeout' }), timeoutMs)),
      ]);
      if (!res?.ok) {
        out.push({ html: null, tsv: null, error: res?.error ?? '后台未完成按键' });
        continue;
      }
      // 后台已发完 trusted 的 Ctrl+A / Ctrl+C，这里读剪贴板（扩展有 clipboardRead 权限）
      const clip = await readClipboard();
      out.push({
        html: clip.html,
        tsv: clip.text,
        error: clip.html || clip.text ? null : clip.error ?? '剪贴板为空',
      });
    } catch (error) {
      out.push({ html: null, tsv: null, error: String(error?.message ?? error) });
    }
  }
  onProgress?.(blockIds.length, blockIds.length);
  return out;
}


/** 读剪贴板：优先 HTML（带样式），退回纯文本 TSV */
async function readClipboard() {
  try {
    const items = await navigator.clipboard.read();
    let html = null;
    let text = null;
    for (const item of items) {
      for (const type of item.types) {
        if (/html/i.test(type)) html = await (await item.getType(type)).text();
        else if (/plain/i.test(type)) text = await (await item.getType(type)).text();
      }
    }
    return { html, text };
  } catch (error) {
    try {
      const text = await navigator.clipboard.readText();
      return { html: null, text, error: text ? null : String(error?.message ?? error) };
    } catch (inner) {
      return { html: null, text: null, error: String(inner?.message ?? inner) };
    }
  }
}
