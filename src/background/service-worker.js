/**
 * 后台服务：负责下载落盘与跨域二进制抓取（content script 受 CORS 限制）。
 *
 * 注意：MV3 service worker 里没有 URL.createObjectURL，
 * 因此这里只接收 content script 已创建好的 blob: URL。
 */

/** tabId -> port，用于把工具栏图标点击转发给对应页面的 content script */
const ports = new Map();

/** 短期提示：图标角标 */
function flashBadge(tabId, text = '!') {
  try {
    chrome.action.setBadgeBackgroundColor({ color: '#FF8800', tabId });
    chrome.action.setBadgeText({ text, tabId });
    setTimeout(() => chrome.action.setBadgeText({ text: '', tabId }), 4000);
  } catch {
    /* 忽略 */
  }
}

chrome.runtime.onConnect.addListener((port) => {
  if (port.name !== 'fde-panel') return;
  const tabId = port.sender?.tab?.id;
  if (tabId == null) return;
  ports.set(tabId, port);
  port.onDisconnect.addListener(() => ports.delete(tabId));
});

/** 点击扩展图标 -> 显示/隐藏导出面板 */
chrome.action.onClicked.addListener((tab) => {
  const port = ports.get(tab.id);
  if (!port) {
    // 页面还没注入脚本（通常是安装后未刷新页面）
    flashBadge(tab.id);
    return;
  }
  port.postMessage({ type: 'toggle-panel' });
});

/** 等到下载真正完成后取回绝对路径（conflictAction 可能改写文件名） */
function waitForDownloadPath(downloadId, timeout = 10000) {
  return new Promise((resolve) => {
    const started = Date.now();
    const tick = () => {
      chrome.downloads.search({ id: downloadId }, (items) => {
        const item = items?.[0];
        if (item?.filename && item.state === 'complete') {
          resolve(item.filename);
          return;
        }
        if (Date.now() - started > timeout) {
          resolve(item?.filename ?? '');
          return;
        }
        setTimeout(tick, 250);
      });
    };
    tick();
  });
}

const sleepMs = (ms) => new Promise((r) => setTimeout(r, ms));

/**
 * 用 chrome.debugger 发 trusted 输入，把内嵌表格的单元格数据读出来。
 *
 * 为什么必须 trusted：飞书表格是 canvas + 自绘选区，合成事件一律不触发
 * （三种组合都实测过，见 tools/probe-sheet-hook.mjs）。
 */
async function extractSheetWithDebugger(tabId, rect) {
  const cmd = (method, params) => chrome.debugger.sendCommand({ tabId }, method, params);
  // attach 在「标签页已被 DevTools/CDP 接管」时会卡住不返回，必须加超时快速失败
  await Promise.race([
    chrome.debugger.attach({ tabId }, '1.3'),
    new Promise((_, reject) =>
      setTimeout(() => reject(new Error('无法附加调试器：该标签页可能开着 DevTools，或已有其他调试连接')), 4000),
    ),
  ]);
  try {
    // 1) 双击进入单元格编辑态。
    //    单击只会选中块（出现蓝色边框），不会进入编辑，后续 Ctrl+A 就选不到单元格。
    await cmd('Input.dispatchMouseEvent', { type: 'mouseMoved', x: Math.round(rect.x), y: Math.round(rect.y) });
    for (const clickCount of [1, 2]) {
      await cmd('Input.dispatchMouseEvent', {
        type: 'mousePressed', x: Math.round(rect.x), y: Math.round(rect.y), button: 'left', clickCount,
      });
      await cmd('Input.dispatchMouseEvent', {
        type: 'mouseReleased', x: Math.round(rect.x), y: Math.round(rect.y), button: 'left', clickCount,
      });
      await sleepMs(60);
    }
    await sleepMs(500);

    // 2) Ctrl+A / Ctrl+C
    for (const k of [
      { key: 'a', code: 'KeyA', vk: 65 },
      { key: 'c', code: 'KeyC', vk: 67 },
    ]) {
      // 必须用 rawKeyDown：'keyDown' 不带 text 时 Chrome 不会把它当快捷键派发给页面
      await cmd('Input.dispatchKeyEvent', {
        type: 'rawKeyDown', modifiers: 2, key: k.key, code: k.code,
        windowsVirtualKeyCode: k.vk, nativeVirtualKeyCode: k.vk,
      });
      await cmd('Input.dispatchKeyEvent', {
        type: 'keyUp', modifiers: 2, key: k.key, code: k.code,
        windowsVirtualKeyCode: k.vk, nativeVirtualKeyCode: k.vk,
      });
      await sleepMs(320);
    }
    await sleepMs(500);

    // 剪贴板由内容脚本读（扩展有 clipboardRead 权限）；
    // 这里只负责把「真实按键」发出去 —— 这是飞书唯一认的输入方式
    return { ok: true };
  } finally {
    // 4) 退出编辑态：不退出的话页面负载上升，后续采集容易超时
    try {
      await cmd('Input.dispatchKeyEvent', { type: 'rawKeyDown', key: 'Escape', code: 'Escape', windowsVirtualKeyCode: 27 });
      await cmd('Input.dispatchKeyEvent', { type: 'keyUp', key: 'Escape', code: 'Escape', windowsVirtualKeyCode: 27 });
    } catch {
      /* 忽略 */
    }
    try {
      await chrome.debugger.detach({ tabId });
    } catch {
      /* 忽略 */
    }
  }
}

chrome.runtime.onMessage.addListener((message, sender, sendResponse) => {
  // 健康检查：内容脚本用它确认后台真的活着（后台挂了时 sendMessage 会静默返回 undefined）
  if (message?.type === 'ping') {
    sendResponse({ ok: true, at: Date.now() });
    return undefined;
  }

  if (message?.type === 'save-docx') {
    chrome.downloads.download(
      {
        url: message.url,
        filename: message.filename,
        saveAs: false,
        conflictAction: 'uniquify',
      },
      (downloadId) => {
        if (chrome.runtime.lastError) {
          sendResponse({ ok: false, error: chrome.runtime.lastError.message });
          return;
        }
        waitForDownloadPath(downloadId).then((path) =>
          sendResponse({ ok: true, downloadId, path }),
        );
      },
    );
    return true;
  }

  if (message?.type === 'capture-region') {
    // WebGL 画布无法通过 toDataURL 取出内容（默认不保留绘制缓冲），
    // 这里截整屏，由 content script 按元素位置裁剪。
    chrome.tabs.captureVisibleTab(null, { format: 'png' }, (dataUrl) => {
      if (chrome.runtime.lastError || !dataUrl) {
        sendResponse({
          ok: false,
          error: chrome.runtime.lastError?.message ?? '截图失败',
        });
      } else {
        sendResponse({ ok: true, dataUrl });
      }
    });
    return true;
  }

  if (message?.type === 'not-doc-page') {
    if (sender.tab?.id != null) flashBadge(sender.tab.id);
    return undefined;
  }

  if (message?.type === 'fetch-binary') {
    (async () => {
      try {
        const response = await fetch(message.url, {
          credentials: 'include',
          headers: message.headers ?? {},
        });
        if (!response.ok) {
          sendResponse({ ok: false, error: `HTTP ${response.status}` });
          return;
        }
        const buffer = await response.arrayBuffer();
        sendResponse({
          ok: true,
          buffer,
          mimeType: response.headers.get('content-type') ?? '',
        });
      } catch (error) {
        sendResponse({ ok: false, error: String(error.message ?? error) });
      }
    })();
    return true;
  }

  if (message?.type === 'has-debugger') {
    chrome.permissions.contains({ permissions: ['debugger'] }).then((ok) => sendResponse({ ok }));
    return true;
  }

  if (message?.type === 'request-debugger') {
    chrome.permissions
      .request({ permissions: ['debugger'] })
      .then((ok) => sendResponse({ ok }))
      .catch((error) => sendResponse({ ok: false, error: String(error?.message ?? error) }));
    return true;
  }

  if (message?.type === 'extract-sheet') {
    const tabId = sender.tab?.id;
    if (tabId == null) {
      sendResponse({ ok: false, error: '没有 tabId' });
      return undefined;
    }
    extractSheetWithDebugger(tabId, message.rect ?? { x: 0, y: 0 })
      .then((data) => sendResponse({ ok: Boolean(data?.ok) }))
      .catch((error) => sendResponse({ ok: false, error: String(error?.message ?? error) }));
    return true;
  }

  return undefined;
});
