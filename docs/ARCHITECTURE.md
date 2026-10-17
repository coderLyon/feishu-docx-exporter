# 架构说明

改代码前请先读本文的「硬约束」一节——里面每一条都是踩过的坑。

## 一、数据流

```
文档 URL
  │  parseDocUrl：/wiki/<token> 或 /docx/<token>
  │  wiki 走 GET /space/api/wiki/v2/tree/get_node/ → obj_token + obj_type
  ▼
块数据分页拉取
  GET /space/api/docx/pages/client_vars?id=<objToken>&mode=7&limit=500&cursor=<cursor>
  │  mergeClientVars：把各页 block_map 合并
  ▼
buildBlockTree：扁平 block_map → 树（优先按根块 children，缺失时按 parent_id 兜底）
  ▼
renderNode：按块类型分发到 paragraph / table / media 渲染器
  ▼
docx 元素数组（Paragraph / Table / ImageRun）
  │  buildDocx：套用 styles 段落样式 + numbering 编号配置，new Document(...)
  ▼
Packer.toBlob() → content script 建 blob URL → service-worker chrome.downloads 落盘
  │  （download 回调只给 id，conflictAction 可能改名 → 轮询 chrome.downloads.search 到 complete 才拿得到绝对路径）
  ▼
面板显示结果路径
```

**资源（图片 / 内嵌表格 canvas / 画板 canvas）走另一条路**：导出前由 `content/capture.js` 滚动两遍页面完成采集，按文档出现顺序入队；渲染时由 `ctx.takeResource(kind)` 按下标消费。

## 二、模块地图

| 文件 | 职责 |
| --- | --- |
| `content/index.js` | 入口：识别文档页、连后台 port、编排导出、挂载面板（隔离世界，靠 `data-fde-ready` 属性与 DOM 事件与外界通信） |
| `content/ui.js` + `ui.css` | Shadow DOM 面板：选项 / 进度 / 结果，选项存 `localStorage['fde-options']` |
| `content/capture.js` | 两遍扫描 + 按 `docY` 去重排序，采集图片、内嵌表格 canvas、画板 canvas |
| `core/api.js` | wiki 解析、client_vars 分页、统一请求（CSRF 双 header、错误归一化） |
| `core/tree.js` | block_map → 块树 |
| `core/text.js` | 富文本 attribs 解析（`*N` 属性引用、`+LEN` 36 进制长度、行内代码 `|N`） |
| `core/export.js` | 导出编排：选项 → 页面尺寸/版心宽度/字体 → 调用渲染 → 打包 |
| `core/render/index.js` | 渲染分发 + 渲染上下文（含样式栈 `quoteDepth` / `tableDepth`） |
| `core/render/paragraph.js` | 段落、标题、列表、待办、引用、callout、代码块 |
| `core/render/table.js` | 原生表格：合并单元格、列宽、表头判定 |
| `core/render/media.js` | 图片 / 附件 / 内嵌表格 / 画板 / 多维表格 |
| `core/docx/styles.js` | **样式体系**（Fde* 段落样式 + 行内代码字符样式） |
| `core/docx/numbering.js` | 有序列表编号定义 |
| `core/docx/runs.js` | 行内 run 生成（粗斜体/颜色/底色/链接/行内代码） |
| `core/docx/build.js` | 组装 Document、目录、打包 |
| `shared/constants.js` | 块类型、TYPO 排版基准、纸张/边距/字体、样式 id 常量 |

## 三、关键设计决策

### 1. 数据优先，而非截图
页面 DOM 只用于采集「拿不到数据」的内容（图片、canvas 类）。因此产物无水印、表格原生可编辑、标题是 Word 真标题（可被目录识别）。

### 2. 样式体系：段落级样式承载一切
**每类文本只有一种段落样式，段落与 run 上不写与之冲突的直接格式**（详见 [STYLE-SYSTEM.md](STYLE-SYSTEM.md)）。这样做的好处是用户在 Word 里改一次「标题 1」样式，全文同步生效——已用深度验证证实（改「标题 1」字号 16→26pt，全文 H1 同步变化）。

### 3. 行距用绝对行高还是倍数？
历史教训：微软雅黑自然行高 ≈ **1.714em**（实测表格 10.5pt 单倍行距渲染出 18pt），因此 `auto` 倍数会乘在这个自然行高上，`1.625 倍`实际渲染成 2.79em。
现按用户要求统一为**单倍行距**：`{ line: 240, lineRule: LineRuleType.AUTO }`，只写在样式里，段落不带。

### 4. 标题编号来自数据，不靠猜
飞书用 `data.seq === 'auto'` 标记参与自动编号的标题；「附件/附录」类标题**没有** seq 字段，因此既不补号也不占计数器。早期版本用「附件」正则猜测会误判，已改为读数据。

### 5. 表格表头来自 `data.header_row`
- `true` → 首行渲染为表头（`tableHeader` + 浅灰底 + 跨页重复）
- `false` → **不加**（这些表的首行是「请求参数（HEAD）」这类跨列合并的分节行）
- 字段缺失 → 保守兜底：首行无列向合并且每格都有内容才当表头

### 6. 资源按下标顺序消费
`ctx.takeResource('image' / 'sheet' / 'whiteboard')` 依赖「采集顺序 == 文档出现顺序」。任何改变采集顺序的改动都必须同步验证配图是否错位。

## 四、硬约束（改代码前必读）

1. **必须边滚动边采集**：飞书虚拟化渲染，块移出视口即销毁；懒渲染有竞态 → 两遍扫描 + 按 `docY` 去重排序；抓取过程会滚动文档，**抓完必须恢复 scrollTop**，否则同块重复采集。
1b. **扫描前必须先做慢速预热滚动**：内嵌电子表格、画板是重型懒挂载块，扫描节奏太快时它们还没渲染就翻过去了。实测某文档（370 块、11 张内嵌表格）预热前 **11 张全部漏采**，加 `warmUpScroll`（半步长 / 450ms，且循环里重读 `scrollHeight`——它会随内容挂载增长）后 **11/11 全部采到、降级项清零**。
2. **画板采集只能从 canvas 反查容器**：`canvas.closest('[class*="whiteboard"]')`。曾尝试「按类名找容器」的定点采集，命中的是 `whiteboard-protal-container` 这个 0 高占位容器（永远没有 canvas），导致画板全部丢失。**画板是重型懒挂载，容器先存在、canvas 后绘制**。
3. **画板两层等大 canvas，只合成 content 层**：合成两层会文字重影；`auxiliary` 层全空白，用「墨量探测」选层，空白时加长重试（1200/2000ms）。
4. **画板拼接不能用滚动容器切块**：容器 scrollHeight 略大于画布时会切成 2 块，同一幅图画两次错位叠加（6.2 状态流转图重影的根因）→ 白板走 `singleTile` 单块拼接；内嵌电子表格仍需滚动拼接。
5. **飞书表格 `cell_set` 只有 merge_info + block_id**，没有单元格底色/对齐字段；表头是表级布尔 `header_row`。
6. **`mergeClientVars` 用 `Object.assign` 合并**，后页覆盖前页；若某页返回被裁剪过的根块，顶层顺序会乱（已加固但仍需留意）。
7. **service worker 无 `URL.createObjectURL`** → blob URL 由 content script 创建后传后台。
8. **保存路径要轮询**：download 回调只给 id 且 `conflictAction` 会改名。
9. **清单坑**：空 `default_popup`、声明 `default_locale` 却缺 `_locales/` 都会加载失败。
10. **隔离世界**：content script 的 `window` 与主世界不可见，DOM 与 DOM 事件才是通道（自动化用 `window.__fdeToggle` 或 `fde-toggle` 事件）。
11. **离屏画布必须 `willReadFrequently`**；不要读页面画布像素。
12. **块渲染异常一律捕获降级**进 `report.degraded / errors`，绝不中断整体导出。

13. **文档标题不是独立块**：飞书把标题挂在根 `page` 块的 text 上。`/docx/<token>` 链接取不到 `node.title`（`get_node` 只服务 wiki），必须用 `extractPageTitle()` 从根块兜底，否则标题会退化成「未命名文档」。
14. **飞书图片的 src 多为 `blob:` URL**：content script 的 `fetch` 受页面 CSP 限制会失败 → 采集必须**优先走 canvas**（读已解码的 `<img>`），跨域直链才退 fetch。
15. **单元格内的图片按列宽缩放**（`ctx.cellWidthPx`），不能用版心宽度，否则撑破表格。

16. **引用块是单格表格**（与飞书官方导出一致）：竖线用表格左边框（位置固定），
    官方数值 = 竖线 `BBBFC4`/sz18、正文 `646A73`、内边距 60/120/30/120、段前后 120、行距 288 auto。不要把边框加在 `Quote` 样式上——
    引用内的列表段落带 `numPr`，编号缩进会覆盖样式缩进，导致边框错位成「两条竖线」（已踩过）。
17. **采集必须有三道资源闸**：瓦片数 ≤9、输出画布 ≤8M 像素、采集总预算 4 分钟。
    缺任何一道，大文档（800+ 块）会因巨型画布/超大 base64 把渲染进程拖死（已踩过）。

18. **页面必须可见，否则采不到东西**：Chrome 对被遮挡/后台窗口会节流渲染，
    `document.visibilityState` 变成 `hidden` 时飞书虚拟化只渲染个位数的块
    （实测 6/90），采集结果全空。`serve-chrome` 已加
    `--disable-backgrounding-occluded-windows / --disable-renderer-backgrounding /
    --disable-background-timer-throttling`，采集前还会 `window.focus()`。
    **「首屏偶发漏采」的真凶就是它**，不是时序。
19. **排版三件套**：标题 `keepNext`（不与正文分离）、表格行与代码块 `cantSplit`（不跨页劈开）、
    表头 `tableHeader`（跨页重复）。缺 keepNext 会出现「标题孤悬页底」，长文档里最伤连贯性。

## 五、尚未打通的两条路（附证据）

### 内嵌电子表格 → 原生表格
- 接口：`POST /space/api/v3/sheet/client_vars`，**必须带 `memberId`（传 0 即可）**，否则返回 400 `Parameter Error`
- 请求体：
  ```json
  { "memberId": 0, "schemaVersion": 9, "openType": 1,
    "token": "<sheetToken>", "sheetRange": { "sheetId": "<sheetId>" }, "clientVersion": "v0.0.1" }
  ```
- `data.snapshot.gzipBlockMeta`（gzip + base64，JSON）给出 sheetId → blockId 与行列范围；`data.snapshot.blocks[blockId]`（gzip + base64）是单元格内容
- **阻塞点**：解出来的块是**私有二进制格式**（非标准 protobuf，递归解析不可靠），无法保证「空单元格不占位」时的行列对齐。错位的表格比截图更糟，故维持截图方案
- 证据：`samples/embed-probe.json`、`samples/sheet-live.json`

### html 附件 → 真内嵌
- 取址：`POST /space/api/box/file/cdn_url/`，body `[{ "file_token": "<token>", "policy": "near" }]` → 返回 `lf*-drive.feishucdn.com` 直链（带 cipher/secret/nonce）
- **阻塞点**：docx@9 无 OLE 封装 API；Word 双击激活内嵌对象需要 OLE 复合文档（Compound File）格式，裸字节内嵌大概率只能显示图标且打不开
- 现状：以「📎 文件名 + 大小 + 超链接」的文件卡片呈现
