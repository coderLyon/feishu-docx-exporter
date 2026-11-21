/** 块类型分发器 + 渲染上下文 */
import { orderedSequence, startsNewOrderedInstance } from './list-instance.js';
import { detectLang, t } from '../../shared/messages.js';
import {
  renderBullet,
  renderCallout,
  renderCode,
  renderDivider,
  renderOrdered,
  renderFeishuOnly,
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
  landscapeContentWidth,
  attachments,
  bitables,
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
    landscapeContentWidth,
    attachments,
    bitables,
    contentWidthPx: Math.round(pageContentWidth / 15),
    bodySize: bodySize ?? 22,
    bodyFont,
    // 文档语言：决定占位/降级文案用中文还是英文（按正文中文占比自动判断）
    lang: detectLang(textSample),
    listInstance: 0,
    counters: { image: 0, sheet: 0, whiteboard: 0 },
    // >0 表示正在渲染表格单元格，用于切换表格专用段落样式
    tableDepth: 0,
    // >0 表示正在渲染引用容器，子段落改用 FdeQuote 样式
    quoteDepth: 0,

    /** 浏览器采集按块匹配，离线工具也可传入已有的资源队列。 */
    takeResource(kind, blockId) {
      const matched = this.resources.byBlock?.[kind];
      if (matched) return matched.get(blockId) ?? null;
      const list = this.resources[`${kind}s`] ?? [];
      const index = this.counters[kind] ?? 0;
      this.counters[kind] = index + 1;
      return list[index] ?? null;
    },

    /** 每组兄弟块维护自己的编号。子列表和其他单元格各自计数。 */
    renderChildren(nodes, level = 0) {
      const out = [];
      let list = null;
      for (const node of nodes ?? []) {
        if (!node) continue;
        if (node.type === 'ordered') {
          if (startsNewOrderedInstance(node, list?.last)) {
            const start = orderedSequence(node) ?? 1;
            list = { start, last: start, instance: ++this.listInstance };
          } else {
            list.last += 1;
          }
          out.push(...renderNode(node, this, level, list));
        } else {
          // 补充说明文字可位于相邻编号项之间。
          if (node.type !== 'text') list = null;
          out.push(...renderNode(node, this, level));
        }
      }
      return out;
    },
  };

  return ctx;
}

/** 单块渲染分发 */
export function renderNode(node, ctx, level = 0, list) {
  try {
    switch (node.type) {
      case 'page':
        return ctx.renderChildren(node.children, level);
      case 'text':
      case 'heading1':
      case 'heading2':
      case 'heading3':
      case 'heading4':
      case 'heading5':
      case 'heading6':
      case 'heading7':
      case 'heading8':
      case 'heading9':
        return renderTextLike(node, ctx, level);
      case 'bullet':
        return renderBullet(node, ctx, level);
      case 'ordered':
        return renderOrdered(node, ctx, level, list);
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
      case 'base_refer':
        // base_refer 是「引用某个多维表格视图」，与 bitable 同源，复用同一渲染（带 table/view 参数）
        return renderBitable(node, ctx);
      case 'synced_source':
        // 同步块：源块的内容就在它的 children 里，直接渲染真实内容
        return node.children?.length
          ? ctx.renderChildren(node.children, level)
          : renderFeishuOnly(node, ctx);
      case 'chart_refer_host_perm':
      case 'task_list':
        // 飞书专有内容（图表引用需宿主权限 / 飞书任务），Word 无对应形式 → 说明性占位
        return renderFeishuOnly(node, ctx);
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
