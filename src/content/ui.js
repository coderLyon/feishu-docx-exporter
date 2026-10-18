/** 导出面板：点击扩展图标后弹出（渲染在 Shadow DOM 内，与宿主页面隔离） */
import styles from './ui.css';
import { DEFAULT_OPTIONS, FONT_SCHEMES, PAGE_MARGINS, PAGE_SIZES } from '../shared/constants.js';

const STAGES = [
  { id: 'scroll', label: '渲染页面资源' },
  { id: 'resolve', label: '解析文档链接' },
  { id: 'fetch', label: '拉取块数据' },
  { id: 'render', label: '渲染排版' },
  { id: 'pack', label: '打包 docx' },
];

export class ExportPanel {
  /**
   * @param {HTMLElement} mount 挂载点
   * @param {{onExport: Function, onCancel: Function}} handlers
   */
  constructor(mount, { onExport, onCancel }) {
    this.shadow = mount.attachShadow({ mode: 'open' });
    this.handlers = { onExport, onCancel };
    this.state = {
      visible: false,
      mode: 'panel', // panel | busy | result | error
      options: { ...DEFAULT_OPTIONS, ...this.loadPersisted() },
      progress: { percent: 0, stage: '' },
      result: null,
      error: '',
    };
    this.render();
  }

  loadPersisted() {
    try {
      return JSON.parse(localStorage.getItem('fde-options') ?? '{}');
    } catch {
      return {};
    }
  }

  persist() {
    try {
      localStorage.setItem('fde-options', JSON.stringify(this.state.options));
    } catch {
      /* 忽略 */
    }
  }

  get visible() {
    return this.state.visible;
  }

  setVisible(visible) {
    // 关闭后再打开时，若上次停在终态（导出完成 / 失败），回到选项菜单，
    // 否则用户想「再导一次」只能刷新页面
    if (visible && (this.state.mode === 'result' || this.state.mode === 'error')) {
      this.setState({ visible, mode: 'panel', result: null, progress: null });
      return;
    }
    this.setState({ visible });
  }

  setState(patch) {
    this.state = { ...this.state, ...patch };
    this.render();
  }

  setProgress(progress) {
    // 进度条只进不退：采集会跑两遍（第二遍补漏），若两遍都映射 20→38%
    // 就会出现「跳回去」的观感，用户以为卡住了重新开始
    const prev = this.state.progress?.percent ?? 0;
    const next = { ...progress };
    if (typeof next.percent === 'number' && next.percent < prev && prev < 100) {
      next.percent = prev;
    }
    this.state.progress = { ...this.state.progress, ...next };
    if (this.state.mode === 'busy') this.render();
  }

  destroy() {
    this.shadow.innerHTML = '';
  }

  render() {
    const { visible, mode } = this.state;
    let html = '';
    if (visible) {
      if (mode === 'panel') html = this.panelHtml();
      else if (mode === 'busy') html = this.busyHtml();
      else if (mode === 'result') html = this.resultHtml();
      else html = this.errorHtml();
    }
    this.shadow.innerHTML = `<style>${styles}</style><div class="fde-root">${html}</div>`;
    this.bind();
  }

  panelHtml() {
    const o = this.state.options;
    const pageSizeOptions = Object.entries(PAGE_SIZES)
      .map(([key, value]) => `<option value="${key}" ${o.pageSize === key ? 'selected' : ''}>${value.label}</option>`)
      .join('');
    const marginOptions = Object.entries(PAGE_MARGINS)
      .map(([key, value]) => `<option value="${key}" ${o.pageMargin === key ? 'selected' : ''}>${value.label}</option>`)
      .join('');
    const fontOptions = Object.entries(FONT_SCHEMES)
      .map(([key, value]) => `<option value="${key}" ${o.fontScheme === key ? 'selected' : ''}>${value.label}</option>`)
      .join('');

    return `<div class="fde-panel">
      <div class="fde-panel-head">
        <span class="fde-panel-title">导出为 Word</span>
        <button class="fde-close" data-act="close">✕</button>
      </div>
      <div class="fde-body">
        <div class="fde-group">
          <label class="fde-label">图片</label>
          <div class="fde-row">
            <div class="fde-seg">
              <button data-act="opt" data-key="embedImages" data-value="1" class="${o.embedImages ? 'fde-on' : ''}">内嵌</button>
              <button data-act="opt" data-key="embedImages" data-value="0" class="${!o.embedImages ? 'fde-on' : ''}">外链</button>
            </div>
          </div>
        </div>
        <div class="fde-switch">
          <span>去除水印</span>
          <button class="fde-toggle ${o.removeWatermark ? 'fde-on' : ''}" data-act="toggle" data-key="removeWatermark"></button>
        </div>
        <div class="fde-switch">
          <span>包含目录</span>
          <button class="fde-toggle ${o.includeToc ? 'fde-on' : ''}" data-act="toggle" data-key="includeToc"></button>
        </div>
        <div class="fde-group">
          <label class="fde-label">纸张</label>
          <select class="fde-select" data-act="select" data-key="pageSize">${pageSizeOptions}</select>
        </div>
        <div class="fde-group">
          <label class="fde-label">页边距</label>
          <select class="fde-select" data-act="select" data-key="pageMargin">${marginOptions}</select>
        </div>
        <div class="fde-group">
          <label class="fde-label">字体</label>
          <select class="fde-select" data-act="select" data-key="fontScheme">${fontOptions}</select>
        </div>
        <div class="fde-switch">
          <span>附件嵌入为文件对象</span>
          <button class="fde-toggle ${o.embedAttachments ? 'fde-on' : ''}" data-act="toggle" data-key="embedAttachments"></button>
        </div>
        <div class="fde-group">
          <label class="fde-label">内嵌表格</label>
          <select class="fde-select" data-act="select" data-key="embeddedSheet">
            <option value="excel" ${o.embeddedSheet === 'excel' ? 'selected' : ''}>导出为 Excel 对象（双击可编辑）</option>
            <option value="image" ${o.embeddedSheet === 'image' ? 'selected' : ''}>导出为图片</option>
            <option value="table" ${o.embeddedSheet === 'table' ? 'selected' : ''}>导出为可编辑表格</option>
          </select>
        </div>
        <p class="fde-hint">再次点击工具栏图标可收起面板；仅发送只读请求，不会修改文档。</p>
      </div>
      <div class="fde-foot">
        <button class="fde-btn fde-btn-ghost" data-act="close">取消</button>
        <button class="fde-btn fde-btn-primary" data-act="export">开始导出</button>
      </div>
    </div>`;
  }

  busyHtml() {
    const { stage } = this.state.progress;
    const percent = Math.max(0, Math.min(100, Math.round(Number(this.state.progress.percent) || 0)));
    const currentIndex = STAGES.findIndex((s) => s.id === stage);
    return `<div class="fde-panel">
      <div class="fde-panel-head">
        <span class="fde-panel-title">正在导出</span>
      </div>
      <div class="fde-body">
        <div class="fde-stages">
          ${STAGES.map((item, index) => {
            const cls =
              index < currentIndex ? 'fde-done' : index === currentIndex ? 'fde-active' : '';
            return `<div class="fde-stage ${cls}"><span class="fde-dot">${
              index < currentIndex ? '✓' : ''
            }</span>${item.label}</div>`;
          }).join('')}
        </div>
        <div class="fde-bar"><i style="width:${Math.max(4, percent)}%"></i></div>
        <p class="fde-hint" style="margin-top:8px">${percent}% · 请勿关闭页面</p>
      </div>
    </div>`;
  }

  resultHtml() {
    const { result } = this.state;
    const size = result?.size ? ` · ${Math.round(result.size / 1024)} KB` : '';
    const degraded = result?.report?.degraded ?? [];
    const errors = result?.report?.errors ?? [];
    const stats = result?.stats;
    const statText = stats
      ? `共 ${stats.blocks} 个块 · 图片 ${stats.images} · 表格截图 ${stats.sheets} · 画板 ${stats.whiteboards}`
      : '';
    const path = result?.path ?? '';

    return `<div class="fde-panel">
      <div class="fde-panel-head">
        <span class="fde-panel-title">导出完成</span>
        <button class="fde-close" data-act="close">✕</button>
      </div>
      <div class="fde-body">
        <div class="fde-result-ok">
          <div class="fde-check">✓</div>
          <div><b>${result?.filename ?? ''}</b><small>${size ? size.replace(' · ', '') : ''}</small></div>
        </div>
        ${
          path
            ? `<div class="fde-path">
                 <span class="fde-path-text">${path}</span>
                 <button class="fde-copy" data-act="copy-path" data-path="${path}">复制</button>
               </div>`
            : `<p class="fde-hint" style="margin:0 0 8px">已保存到浏览器默认下载目录</p>`
        }
        ${statText ? `<p class="fde-hint" style="margin:0 0 8px">${statText}</p>` : ''}
        ${
          degraded.length
            ? `<ul class="fde-degraded">${degraded
                .slice(0, 30)
                .map((item) => `<li>· ${item}</li>`)
                .join('')}</ul>`
            : ''
        }
        ${
          errors.length
            ? `<ul class="fde-degraded" style="background:rgba(245,74,69,.08)">${errors
                .slice(0, 20)
                .map((item) => `<li>· ${item}</li>`)
                .join('')}</ul>`
            : ''
        }
      </div>
      <div class="fde-foot">
        <button class="fde-btn fde-btn-ghost" data-act="again">重新导出</button>
        <button class="fde-btn fde-btn-primary" data-act="close">完成</button>
      </div>
    </div>`;
  }

  errorHtml() {
    return `<div class="fde-panel">
      <div class="fde-panel-head">
        <span class="fde-panel-title">导出失败</span>
        <button class="fde-close" data-act="close">✕</button>
      </div>
      <div class="fde-body">
        <div class="fde-error">${this.state.error}</div>
        <p class="fde-hint">请确认当前页面是飞书文档且仍有访问权限，然后重试。</p>
      </div>
      <div class="fde-foot">
        <button class="fde-btn fde-btn-primary" data-act="again">重试</button>
      </div>
    </div>`;
  }

  bind() {
    const root = this.shadow.querySelector('.fde-root');
    if (!root) return;

    root.addEventListener('click', (event) => {
      const target = event.target.closest('[data-act]');
      if (!target) return;
      const action = target.dataset.act;
      const { key, value } = target.dataset;

      if (action === 'close') {
        this.setVisible(false);
        return;
      }
      if (action === 'export') {
        this.persist();
        this.handlers.onExport?.(this.state.options);
        return;
      }
      if (action === 'again') {
        this.setState({ mode: 'panel' });
        return;
      }
      if (action === 'cancel') {
        this.handlers.onCancel?.();
        this.setVisible(false);
        return;
      }
      if (action === 'copy-path') {
        const path = target.dataset.path ?? '';
        const original = target.textContent;
        navigator.clipboard
          ?.writeText(path)
          .then(() => {
            target.textContent = '已复制';
            setTimeout(() => {
              target.textContent = original;
            }, 1500);
          })
          .catch(() => {});
        return;
      }
      if (action === 'toggle') {
        const options = { ...this.state.options, [key]: !this.state.options[key] };
        this.setState({ options });
        return;
      }
      if (action === 'opt') {
        const options = { ...this.state.options, [key]: value === '1' };
        this.setState({ options });
        return;
      }
      if (action === 'select') {
        target.addEventListener('change', () => {
          const options = { ...this.state.options, [key]: target.value };
          this.setState({ options });
        });
      }
    });
  }
}
