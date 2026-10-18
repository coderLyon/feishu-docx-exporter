# 工具脚本清单

> 约定：所有脚本都**不写死测试文档**。需要文档地址的一律从命令行传入，或从当前打开的页面自动识别。
> 若必须保留默认值，也只是便捷 fallback。

## 一、构建与打包

| 命令 | 说明 |
| --- | --- |
| `npm run build` | esbuild 打包到 `dist/`（`node build.mjs`） |
| `npm run watch` | 监听模式 |
| `npm run pack` | 打成 `飞书文档导出Word扩展.zip`，旧包自动备份到 `output/zip-backup/` |

## 二、日常迭代（核心链路）

| 脚本 | 用法 | 说明 |
| --- | --- | --- |
> **防泄漏机制（v0.2.2）**：`serve-chrome` 带空闲自毁（默认 10 分钟无 CDP 使用即自动关 Chrome 退出），
> 并在收到 SIGINT/SIGTERM/exit 时用 `taskkill /T` 杀掉整棵 Chrome 进程树；
> 任何 CDP 使用方（导出/下载/探测）都会自动刷新心跳。即使忘了停后台任务也不会长期占内存。

| `tools/open-doc.mjs` | `node tools/open-doc.mjs <URL>` | 把已打开的调试 Chrome 导航到指定文档（复用同一标签页；**不要关闭唯一标签页**，否则窗口会一起关掉） |
| `npm run chrome:stop` | 一键停止：先停 serve-chrome 常驻进程（否则它会重启 Chrome），再杀 Chrome 进程树，最后清心跳文件 |
| `tools/serve-chrome.mjs --idle=10` | `node tools/serve-chrome.mjs [文档URL]` | **后台常驻启动 Chrome**（远程调试端口 9222，profile `.chrome-user`）。必须用后台任务持有，否则 spawn 的进程会被回收 |
| `tools/export-run.mjs` | `node tools/export-run.mjs r021 [--reload]` | 真实加载扩展跑一次导出，产出 `samples/runs/<id>/source.docx` |
| `tools/check-styles.mjs` | `node tools/check-styles.mjs <docx> [--json f] [--expect tables=24]` | 样式/结构断言，退出码 0 通过 / 1 失败 / 2 用法错误 |
| `tools/render-offline.mjs` | `node tools/render-offline.mjs [out.docx] [--sample=...] [--host=...]` | L0 离线渲染（不启 Chrome） |
| `tools/docx2pdf.ps1` | `powershell -File tools\docx2pdf.ps1 -Src <docx> -Dst <pdf>` | Word COM 转 PDF |
| `tools/pdf2png.py` | `python tools/pdf2png.py <pdf> <outdir> --dpi 144 [--pages 1-25]` | PDF 逐页转 PNG（逐页释放，内存友好） |
| `tools/pdf-probe.py` | `python tools/pdf-probe.py <pdf> [在线锚点meta.json]` | PDF 级实测断言 + 产出 `word/anchors.json` |
| `tools/make-manifest.mjs` | `node tools/make-manifest.mjs <runId> [--strict]` | 按锚点配对「在线截图 / Word 页」生成 `pair/` |
| `tools/shot-online.mjs` | `node tools/shot-online.mjs <outdir> [--warm] [文档URL路径片段]` | 沿正文列分段截在线基线；`--warm` 只预热滚动不截图（用于触发懒渲染） |
| `tools/verify-style-editable.ps1` | `powershell -File tools\verify-style-editable.ps1 -Src <docx> -PdfOut <pdf>` | 深度验证：改样式后全文是否同步 |
| `tools/run-iter.mjs` | `node tools/run-iter.mjs r021 --from=pdf` | 串跑 build→export→pdf→png→manifest，支持 `--from=` 断点 |
| `tools/check-docx.mjs` | `node tools/check-docx.mjs <docx> [--extract dir]` | 基础统计（段落/表格/图片/合并/链接） |
| `tools/check-layout.mjs` | `node tools/check-layout.mjs <docx>` | 版面几何：图片/表格是否溢出版心（版心从 sectPr 实读） |

## 三、Chrome / CDP 基础设施

| 脚本 | 说明 |
| --- | --- |
| `tools/cdp.mjs` | 极简 CDP 客户端（`connect / targets / attachToPage / send / eval`） |
| `tools/browser.mjs` | 共享常量（ROOT、默认文档 URL 等） |
| `tools/launch-chrome.mjs` | 单独启动一个带调试端口的 Chrome |
| `tools/ext-load.mjs` | 通过 CDP `Extensions.loadUnpacked` 装载扩展（Chrome 137+ 已禁用 `--load-extension`） |


## 四、多文档体检

| 脚本 | 用法 | 说明 |
| --- | --- | --- |
| `tools/multi-doc-audit.mjs` | `node tools/multi-doc-audit.mjs [--outdir=samples/cov] <url1> [url2] ...` | 逐篇抓块数据 → 统计块类型覆盖率与未支持类型。支持清单直接从渲染分发器的 switch 提取，避免与代码漂移 |
| `tools/coverage.mjs` | `node tools/coverage.mjs [doc.json]` | 就本地已有的一份块数据做同样的覆盖率体检（不用开浏览器） |

> 换文档测试是发现「特化代码」的唯一可靠手段——本轮一次就暴露了 4 类未支持块、
> 1 个采集缺陷（重型块漏采）、3 处断言特化。


## 五、与官方产物对比

| 脚本 | 用法 | 说明 |
| --- | --- | --- |
| `tools/fetch-official-word3.mjs` | `node tools/fetch-official-word3.mjs` | 自动走飞书 UI「••• → 下载为 → Word → 导出设置 → 导出」，把官方 docx 落到 `samples/official/`。前置：调试 Chrome 已登录并打开目标文档 |
| `tools/extract-sheets.mjs` | `node tools/extract-sheets.mjs` | 提取内嵌表格单元格数据（点进表格→Ctrl+A/C→读剪贴板→存 localStorage），供「导出为可编辑表格」使用 |
| `tools/audit-align.mjs` | `node tools/audit-align.mjs <pdf> [--census]` | 左对齐体检：量各元素左边界与版心的差值；`--census` 做全文分布统计 |
| `tools/probe-quote-align.mjs` | `node tools/probe-quote-align.mjs` | 量在线引用块的竖线/正文/符号坐标，确定对齐基准 |
| `tools/compare-official.mjs` | `node tools/compare-official.mjs <官方.docx> <我们的.docx>` | 逐项对比页面设置、样式、引用、表格、图片尺寸。注意 `<w:tbl>` 计数含嵌套表，会重复计入 |
| `tools/probe-quote.mjs` | `node tools/probe-quote.mjs` | 读取在线引用块的计算样式（竖线颜色/宽度/内边距），用于 1:1 复刻 |
| `tools/probe-scroll-cost.mjs` | `node tools/probe-scroll-cost.mjs` | 隔离实验：只滚动不采集，判断卡顿来自飞书渲染还是我们的采集代码 |

## 六、诊断脚本（按需使用，不参与正式验证）

`probe.mjs`、`probe-page.mjs`、`probe-sheet.mjs`、`probe-wb.mjs`、`probe-font.mjs`、`probe-style.mjs`、
`probe-embed.mjs`、`inspect.mjs`、`shot-state.mjs`（截当前视口）、`diag-resources.mjs`、`analyze-blocks.mjs`、`proto-dump.mjs`。

## 七、探测与研究

| 脚本 | 用途 |
| --- | --- |
| `tools/probe-embed.mjs` | 探测内嵌表格取数接口 + 附件直链接口（`--sheet=token --file=token`） |
| `tools/probe-wb.mjs` | 探测画板「下载为图片」接口（点击菜单 + 抓包） |
| `tools/probe-sheet.mjs` / `decode-sheet.mjs` | 电子表格取数与快照解码 |
| `tools/probe.mjs` / `probe-page.mjs` / `probe-font.mjs` / `probe-style.mjs` | 页面结构、字体、样式的实测取样 |
| `tools/analyze-blocks.mjs` / `inspect.mjs` / `proto-dump.mjs` | 块数据与人信息分析 |
| `tools/diag-resources.mjs` | 资源采集诊断 |
| `tools/e2e-export.mjs` / `e2e-real.mjs` | 早期端到端试验（注入式 / 真实加载） |
| `tools/fetch-doc.mjs` / `screenshot-doc.mjs` | 抓样本数据与整页截图 |

## 八、重要约定

1. **导出前绝不能 `Page.reload`**：刷新会清掉页面懒渲染状态，导致画板/内嵌表格/图片全部采不到（r016 因此全灭）。面板若停在上一次的「导出完成」结果态，再发两次 `fde-toggle` 事件即可回到菜单视图。
2. **装载扩展的顺序**：必须先 `Extensions.loadUnpacked`，再 `Page.reload`，content script 才会注入（顺序颠倒永不注入）。
3. **`cdp.eval` 的参数语义**：传函数是「函数 + 内联参数」，传字符串是「完整表达式原样执行」。传 `() => ...` 字符串会被当成返回函数的函数，结果恒为 undefined。
4. **Chrome 关不掉**：后台任务会重启 Chrome，且 Chrome 后台应用会保活进程 → 先停后台任务，再循环 `taskkill /F /IM chrome.exe` 2-3 轮。
5. **内存**：8GB 机器上 Chrome(~2GB)/Word COM(~500MB)/pymupdf(<300MB 逐页) **必须串行**；导出完先关 Chrome 再转 PDF。
