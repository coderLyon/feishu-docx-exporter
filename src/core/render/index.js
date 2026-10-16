/** 块类型分发器 + 渲染上下文 */
import { detectLang } from '../../shared/messages.js';
import { t } from '../../shared/messages.js';
import {
  renderBullet,
  renderCallout,
  renderCode,
  renderDivider,
  renderOrdered,
  renderPlaceholder,
  renderQuoteContainer,
  renderTextLike,
  renderTodo,
} from './paragraph.js';
import { renderTable } from './table.js';
import {
  renderBitable,
  renderFile,
  renderImage,
  renderSheet,
  renderWhiteboard,
} from './media.js';

/**
 * 创建渲染上下文。
 * ctx 持有资源队列（图片 / 表格截图 / 画板截图），按文档出现顺序逐个消费。
 */
export function createRenderContext({
  byId,
  options,
  resources,
  host,
  objToken,
  report,
  pageContentWidth,
  bodySize,
  bodyFont,
  textSample = '',
}) {
  const ctx = {
    byId,
    options,
    resources: resources ?? {},
    host,
    objToken,
    report,
    pageContentWidth,
    contentWidthPx: Math.round(pageContentWidth / 15),
    bodySize: bodySize ?? 22,
    bodyFont,
    // 文档语言：决定占位/降级文案用中文还是英文（按正文中文占比自动判断）
    lang: detectLang(textSample),
    ordered: { instance: 0, current: 0 },
    counters: { image: 0, sheet: 0, whiteboard: 0 },
    // 标题自动编号计数器（1..9 级）
    headingCounters: [0, 0, 0, 0, 0, 0, 0, 0, 0],
    // >0 表示正在渲染表格单元格，用于切换表格专用段落样式
    tableDepth: 0,
    // >0 表示正在渲染引用容器，子段落改用 FdeQuote 样式
    quoteDepth: 0,

    /** 按顺序取下一个页面资源（collectResources 返回的是复数键名） */
    takeResource(kind) {
      const list = this.resources[`${kind}s`] ?? [];
      const index = this.counters[kind] ?? 0;
      this.counters[kind] = index + 1;
      return list[index] ?? null;
    },

    /** 渲染一组子块；有序列表在此处维护编号实例 */
    renderChildren(nodes, level = 0) {
      const out = [];
      let prevWasOrdered = false;
      for (const node of nodes ?? []) {
        if (!node) continue;
        if (node.type === 'ordered') {
          if (!prevWasOrdered) this.ordered.instance += 1;
          this.ordered.current = this.ordered.instance;
          out.push(...renderOrdered(node, this, level));
          prevWasOrdered = true;
        } else {
          prevWasOrdered = false;
          out.push(...renderNode(node, this, level));
        }
      }
      return out;
    },
  };

  return ctx;
}

/** 单块渲染分发 */
export function renderNode(node, ctx, level = 0) {
  try {
    switch (node.type) {
      case 'page':
        return ctx.renderChildren(node.children, level);
      case 'text':
        return renderTextLike(node, ctx);
      case 'heading1':
      case 'heading2':
      case 'heading3':
      case 'heading4':
      case 'heading5':
      case 'heading6':
      case 'heading7':
      case 'heading8':
      case 'heading9':
        return renderTextLike(node, ctx);
      case 'bullet':
        return renderBullet(node, ctx, level);
      case 'ordered':
        return renderOrdered(node, ctx, level);
      case 'todo':
        return renderTodo(node, ctx, level);
      case 'code':
        return renderCode(node, ctx);
      case 'quote':
      case 'quote_container':
        return renderQuoteContainer(node, ctx, level);
      case 'callout':
        return renderCallout(node, ctx, level);
      case 'divider':
        return renderDivider();
      case 'table':
        return renderTable(node, ctx);
      case 'table_cell':
        return ctx.renderChildren(node.children, level);
      case 'image':
        return renderImage(node, ctx);
      case 'file':
        return renderFile(node, ctx);
      case 'sheet':
        return renderSheet(node, ctx);
      case 'whiteboard':
      case 'diagram':
      case 'mindnote':
        return renderWhiteboard(node, ctx);
      case 'bitable':
        return renderBitable(node, ctx);
      case 'grid':
      case 'grid_column':
      case 'view':
        return ctx.renderChildren(node.children, level);
      default:
        return renderPlaceholder(node, ctx);
    }
  } catch (error) {
    ctx.report.errors.push(`${node.type} 渲染失败: ${error.message}`);
    return renderPlaceholder(node, ctx, t('blockError', ctx.lang)(node.type));
  }
}
