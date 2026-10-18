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
export async function extractSheetsLive(blockIds, { onProgress, timeoutMs = 12000, scrollTo } = {}) {
  const out = [];
  for (let i = 0; i < blockIds.length; i += 1) {
    onProgress?.(i, blockIds.length);
    const item = blockIds[i];
    // 先滚到该表的纵向位置：虚拟化渲染下滚出视口的块会被卸载。
    // 不能用块 id 定位 —— id 是虚拟列表动态分配的，两次渲染可能不同（实测 6 → 562）
    if (typeof item === 'object' && item?.y != null) scrollTo?.(item.y);
    await new Promise((r) => setTimeout(r, 700));

    // 取当前已挂载、且最接近目标位置的那个表格块
    const mounted = [...document.querySelectorAll('.docx-sheet-block')];
    const wantTop = 120;
    let block = null;
    let best = Infinity;
    for (const el of mounted) {
      const d = Math.abs(el.getBoundingClientRect().top - wantTop);
      if (d < best) {
        best = d;
        block = el;
      }
    }
    if (!block) {
      out.push({ html: null, tsv: null, error: '表格块未挂载（文档可能还没渲染到该位置）' });
      continue;
    }
    // 滚到视口中央：debugger 的点击坐标是视口坐标
    try {
      block.scrollIntoView({ block: 'center' });
    } catch {
      /* 忽略 */
    }
    await new Promise((r) => setTimeout(r, 450));

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
        const err = res?.error ?? '后台未完成按键';
        out.push({ html: null, tsv: null, error: err });
        // 附加失败是环境问题，后面的表也一样会失败 —— 直接停止，别让用户干等
        if (/附加|attach|DevTools/i.test(err)) {
          for (let j = i + 1; j < blockIds.length; j += 1) {
            out.push({ html: null, tsv: null, error: err });
          }
          break;
        }
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
