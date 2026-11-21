/** 扩展入口：识别文档页、连接后台、按需弹出导出面板并编排导出 */
import { parseDocUrl } from '../core/api.js';
import { collectResources } from './capture.js';
import { buildFilename, exportToDocx } from '../core/export.js';
import { DOCX_MIME } from '../core/ole/inject.js';
import { ExportPanel } from './ui.js';

const MOUNT_ID = 'fde-export-mount';

let panel = null;
let port = null;

function isDocPage() {
  try {
    parseDocUrl(location.href);
    return true;
  } catch {
    return false;
  }
}

function ensurePanel() {
  if (panel) return panel;
  if (document.getElementById(MOUNT_ID)) return panel;
  const mount = document.createElement('div');
  mount.id = MOUNT_ID;
  document.documentElement.appendChild(mount);

  panel = new ExportPanel(mount, {
    onExport: (options) => runExport(panel, options),
    onCancel: () => {},
  });
  return panel;
}

/** 供工具栏图标点击调用：显示/隐藏面板 */
function togglePanel() {
  if (!isDocPage()) {
    try {
      port?.postMessage({ type: 'not-doc-page' });
    } catch {
      /* 忽略 */
    }
    return;
  }
  const target = ensurePanel();
  if (!target) return;
  target.setVisible(!target.visible);
}

async function runExport(currentPanel, options) {
  const url = location.href;
  currentPanel.setState({ mode: 'busy', visible: true, progress: { percent: 2, stage: 'resolve' } });

  try {
    const exported = await exportToDocx({
      url,
      options,
      collect: (expected) => collectResources({
        expected,
        options,
        onProgress: (info) => {
          if (info.stage === 'scroll') {
            const ratio = Math.max(0, Math.min(1, info.loaded / info.total));
            currentPanel.setProgress({ stage: 'scroll', percent: 8 + Math.round(ratio * 30) });
          } else {
            currentPanel.setProgress(info);
          }
        },
      }),
      onProgress: (info) => {
        if (info?.stage) currentPanel.setProgress(info);
      },
    });
    if (new URL(location.href).pathname !== new URL(url).pathname) {
      throw new Error('导出期间切换了文档，请在目标页面重新导出');
    }

    currentPanel.setProgress({ stage: 'pack', percent: 96 });

    const filename = buildFilename(exported.title);
    // 防御：任何后处理（OLE/zip 重打包）都可能把 MIME 换成 application/zip，
    // 而 Chrome 会据此改扩展名 —— 这里统一归一到 docx
    const docxBlob = new Blob([exported.blob], { type: DOCX_MIME });
    const objectUrl = URL.createObjectURL(docxBlob);

    // 优先走 chrome.downloads（可拿到绝对路径）；失败则在页面里直接触发下载
    let saved = null;
    try {
      saved = await chrome.runtime.sendMessage({
        type: 'save-docx',
        url: objectUrl,
        filename,
      });
    } catch {
      saved = null;
    }

    if (!saved?.ok) {
      const anchor = document.createElement('a');
      anchor.href = objectUrl;
      anchor.download = filename;
      document.body.appendChild(anchor);
      anchor.click();
      anchor.remove();
    }

    setTimeout(() => URL.revokeObjectURL(objectUrl), 60000);

    currentPanel.setState({
      mode: 'result',
      visible: true,
      result: {
        filename,
        path: saved?.path ?? '',
        size: exported.blob.size,
        report: exported.report,
        stats: exported.stats,
      },
    });

    // 记录上一次导出结果：排查问题时可据此读取降级原因（自动化验证也依赖它）
    try {
      localStorage.setItem(
        'fde-last-run',
        JSON.stringify({
          ok: true,
          at: Date.now(),
          filename,
          size: exported.blob.size,
          stats: exported.stats,
          degraded: exported.report?.degraded ?? [],
          liveSheets: exported.liveSheets,
        }),
      );
    } catch {
      /* 配额不足时忽略 */
    }
  } catch (error) {
    console.error('[fde] 导出失败', error);
    try {
      localStorage.setItem('fde-last-run', JSON.stringify({ ok: false, at: Date.now(), error: String(error.message ?? error) }));
    } catch {
      /* 忽略 */
    }
    currentPanel.setState({ mode: 'error', visible: true, error: String(error.message ?? error) });
  }
}

/** 与后台建立长连接，接收工具栏图标点击 */
function connectPort() {
  try {
    port = chrome.runtime.connect({ name: 'fde-panel' });
    port.onMessage.addListener((message) => {
      if (message?.type === 'toggle-panel') togglePanel();
    });
    port.onDisconnect.addListener(() => {
      port = null;
      // 后台重启后重连
      setTimeout(connectPort, 1000);
    });
  } catch {
    /* 忽略 */
  }
}

connectPort();

// DOM 是各世界共享的，打个标记便于外部确认 content script 是否已注入
document.documentElement.setAttribute('data-fde-ready', '1');

// 飞书是 SPA，路由切换后清掉旧面板，下次点击图标会重新挂载
let lastUrl = location.href;
setInterval(() => {
  if (location.href === lastUrl) return;
  lastUrl = location.href;
  if (!isDocPage() && panel) {
    panel.destroy();
    document.getElementById(MOUNT_ID)?.remove();
    panel = null;
  }
}, 1500);

// 供自动化测试调用
window.__fdeToggle = togglePanel;

/**
 * content script 运行在隔离世界，主世界的 window 变量互不可见，
 * 但 DOM 事件是跨世界共享的，因此用自定义事件作为触发入口。
 */
document.addEventListener('fde-toggle', () => togglePanel());
