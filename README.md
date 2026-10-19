# 飞书文档导出 Word

将飞书（Lark）在线文档导出为高保真 `.docx` 的 Chrome 扩展（MV3）。导出的文档无水印、内容可编辑，标题、正文、表格、引用等均映射到统一的 Word 样式，后续修改样式即可全局生效。

本扩展采用**数据优先**的实现方式：不截取页面 DOM，而是读取飞书文档的结构化块数据并重建为 Word 原生元素；仅图片、内嵌电子表格（canvas 渲染）与画板通过页面采集。因此导出结果默认为无水印，表格可编辑，且全过程只发送只读请求。

---

## 安装

1. 在 [Releases](../../releases) 页面下载最新版本的 `feishu-docx-exporter-vX.Y.Z.zip`；
2. 解压到任意固定目录（安装后请勿删除该目录）；
3. 在 Chrome 或 Edge 地址栏打开 `chrome://extensions`（Edge 为 `edge://extensions`），启用右上角的「开发者模式」；
4. 点击「加载已解压的扩展程序」，选择解压后的目录（包含 `manifest.json` 的那一层）；
5. 打开任意飞书文档（`/wiki/*`、`/docx/*`），右下角出现悬浮按钮即表示安装成功。

安装后，扩展管理页会显示「读取和更改你在 feishu.cn 上的数据」以及「调试此浏览器」权限说明，用途见「权限与隐私」一节。

**升级**：下载新版本 → 覆盖解压至同一目录 → 在 `chrome://extensions` 点击该扩展的「重新加载」。

**卸载**：在扩展管理页点击「移除」，并删除上述目录。

## 使用

1. 打开要导出的飞书文档，点击右下角悬浮按钮唤出面板；
2. 按需选择纸张、页边距、字体与内嵌表格方式，点击「开始导出」；
3. 导出过程中扩展会自动滚动全文以加载懒加载内容，此期间请勿切换标签页或手动滚动；
4. 完成后文件保存至浏览器下载目录，面板会显示页数与降级项。

导出过程只发送只读请求，不会修改原文档。

### 导出选项

| 选项 | 说明 | 默认值 |
| --- | --- | --- |
| 纸张 | A4 / Letter | A4 |
| 页边距 | 标准（左右 900 twips）/ 窄 / 宽 | 标准 |
| 字体 | 微软雅黑 / 宋体 | 微软雅黑 |
| 包含目录 | 插入 Word 目录域，打开后按 F9 更新 | 关闭 |
| 附件嵌入 | 开启：附件以 OLE 对象嵌入，双击打开原文件；关闭：输出超链接 | 开启 |
| 内嵌表格 | 图片（默认）/ 可编辑表格 / Excel 对象 | 图片 |
| 宽表横向页 | 宽高比大于 2.2 的表格单独使用横向页 | 关闭 |

三种内嵌表格方式的差异：

- **图片**（默认）：版式最贴近原文，不依赖浏览器调试接口，也不会受「禁止复制」限制；
- **可编辑表格**：生成 Word 原生表格，可检索、可调整字号，列数较多时自动回退为图片；
- **Excel 对象**：文档中显示整表预览，双击进入 Excel 编辑，保留底色、字号、合并与数字格式。

选项保存在浏览器 `localStorage`（键 `fde-options`），下次导出自动沿用。

## 输出效果

| 内容 | 处理方式 |
| --- | --- |
| 标题 | 使用 Word 内置「标题 1–9」样式，编号由 Word 自动维护，插入或删除标题后自动重排；「附件 / 附录」类标题不参与编号 |
| 正文 | 统一段落样式，保留粗体、斜体、删除线、行内代码、高亮与超链接 |
| 表格 | Word 原生表格，首行作为表头并跨页重复 |
| 引用 | 使用 Word 内置引用样式，保留左侧竖线与内部列表 |
| 代码块 | 等宽字体与浅灰底色，按 `language` 做语法着色，标题显示在代码块上方 |
| 附件 | 居中卡片，含类型徽标、文件名、大小；以 OLE 对象嵌入，双击打开原文件 |
| 图片 / 画板 | 按版心宽度等比缩放并居中 |
| 内嵌电子表格 | Excel 嵌入对象（默认）/ 图片 / Word 表格，按选项决定 |
| 多维表格 | Word 原生表格，表头灰底、跨页重复 |

所有带背景色的块在导出时会校验背景亮度，确保正文与背景的对比度适合打印。

## 文档索引

| 文档 | 内容 |
| --- | --- |
| [docs/ARCHITECTURE.md](docs/ARCHITECTURE.md) | 数据流、模块划分与关键设计决策（修改代码前建议阅读） |
| [docs/STYLE-SYSTEM.md](docs/STYLE-SYSTEM.md) | Word 样式体系、字号与行距映射、自定义样式的修改方式 |
| [docs/VERIFICATION.md](docs/VERIFICATION.md) | 验证体系、断言清单与验收结果 |
| [docs/TOOLS.md](docs/TOOLS.md) | `tools/` 目录下脚本的用途与用法 |
| [docs/TROUBLESHOOTING.md](docs/TROUBLESHOOTING.md) | 常见问题与排查方法 |
| [docs/COMPARISON.md](docs/COMPARISON.md) | 与飞书官方「下载为 Word」的对比 |
| [docs/CHANGELOG.md](docs/CHANGELOG.md) | 版本更新日志 |
| [docs/THIRD-PARTY.md](docs/THIRD-PARTY.md) | 第三方依赖与许可声明 |
| [samples/README.md](samples/README.md) | 样本数据说明 |

## 开发

```bash
npm install
npm run build     # 使用 esbuild 打包到 dist/
npm run watch     # 监听文件变更并增量构建
npm run pack      # 打包为 zip（旧包自动备份至 output/zip-backup/）
npm test          # 离线渲染并执行结构断言
```

开发时加载 `dist/` 目录。修改代码后需在 `chrome://extensions` 点击「重新加载」，并刷新文档页，否则页面中运行的仍是旧版内容脚本。

### 常用命令

```bash
node tools/serve-chrome.mjs <文档URL>    # 启动常驻 Chrome 实例（供自动化使用）
node tools/export-run.mjs <runId> <URL>  # 加载扩展并执行一次真实导出
node tools/check-styles.mjs <docx>       # 样式与结构断言
node tools/docx2pdf.ps1 -Src <docx> -Dst <pdf>   # 通过 Word COM 转为 PDF
```

### 环境要求

- Node.js 22 或以上（构建）
- Chrome 137 或以上（该版本起禁用 `--load-extension`，自动化装载需使用 CDP `Extensions.loadUnpacked`）
- 视觉验证需要本机安装 Word（用于转 PDF）以及 Python `pymupdf`

## 权限与隐私

本项目为非官方第三方工具，与飞书官方无任何关联，不使用开放平台 API Key。

扩展申请的权限及用途：

| 权限 | 用途 |
| --- | --- |
| `host_permissions`（`*.feishu.cn`、`*.larkoffice.com`） | 读取当前用户正在浏览的文档数据 |
| `downloads` | 保存导出的 docx 文件 |
| `clipboardRead` | 读取内嵌表格的单元格内容 |
| `debugger` | 向内嵌表格发送真实鼠标与键盘事件，以触发其复制能力（表格由 canvas 渲染，无其他数据出口） |

数据处理方式：

- 只发送只读请求，不修改任何文档内容；
- 全部处理在本地浏览器内完成，不上传文档内容至任何第三方服务器；
- 不采集使用统计，选项仅保存在本机 `localStorage`；
- 文件通过浏览器原生下载保存，路径由用户选择。

## 许可

代码以 MIT 许可发布，详见 [LICENSE](LICENSE)。第三方依赖与许可见 [docs/THIRD-PARTY.md](docs/THIRD-PARTY.md)。

## 已知限制

1. 选择「Excel 对象」时，单元格内容依赖浏览器调试接口读取，导出时该标签页不可同时开启开发者工具；文档设置为「禁止复制」时无法读取，将回退为图片。默认的「图片」方式不受此限制。
2. 代码块着色依据语言规则生成，与在线文档的配色主题不完全一致。
3. 多维表格中的人员、附件等字段以文本形式呈现。
4. 画板内容通过 canvas 采集，需页面完成懒渲染；建议人工浏览一遍后再导出。
