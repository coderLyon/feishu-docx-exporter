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

chrome.runtime.onMessage.addListener((message, sender, sendResponse) => {
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

  return undefined;
});
