/** 块类型分发器 + 渲染上下文 */
import { startsNewOrderedInstance } from './list-instance.js';
import { detectLang } from '../../shared/messages.js';
import { t } from '../../shared/messages.js';
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
    ordered: { instance: 0, current: 0, byLevel: [] },
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

    /**
     * 渲染一组子块；有序列表在此处维护编号实例。
     *
     * 实例号必须**按层级**记录：嵌套子列表会递归进 renderChildren，
     * 若共用同一个全局 current，子列表的实例号会顶掉父级的，
     * 父级后续项就被并进子列表实例、从 1 重新计数
     *（实测：1、2、a、然后 1、2、3 —— 编号倒退）。
     */
    renderChildren(nodes, level = 0) {
      const out = [];
      let prevWasOrdered = false;
      for (const node of nodes ?? []) {
        if (!node) continue;
        if (node.type === 'ordered') {
          if (startsNewOrderedInstance(node, prevWasOrdered)) {
            this.ordered.instance += 1;
            this.ordered.byLevel[level] = this.ordered.instance;
          }
          const instance = this.ordered.byLevel[level] ?? this.ordered.instance;
          out.push(...renderOrdered(node, this, level, instance));
          prevWasOrdered = true;
        } else {
          // 普通段落不打断有序列表的连续编号：飞书里列表项后的补充说明就是一个 text 块，
          // 若在此处当作「列表结束」，后面的项会重开实例、编号从 1 重新开始
          //（线上表现：1、2、3、然后又是 1、2、3…）。标题/表格/代码/引用/图片等仍会打断。
          if (node.type !== 'text') prevWasOrdered = false;
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
