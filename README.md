# 飞书文档导出 Word（Chrome MV3 扩展）

把飞书（Lark）在线文档导出为高保真 `.docx`：**无水印、原生可编辑**，标题/正文/表格/引用各有统一 Word 样式，导出后改一次样式即全文生效。

核心策略是**数据优先**：不截页面 DOM，而是读飞书内部接口的结构化块数据，重建为 Word 原生元素；只有图片、内嵌电子表格（canvas）与画板才回到 DOM 采集。因此默认无水印、原生表格可编辑、全程只发只读请求。

---

## 安装（普通用户）

1. 打开本仓库的 [Releases](../../releases) 页，下载最新版 `飞书文档导出Word扩展.zip`；
2. 解压到任意**固定目录**（例如 `D://Tools//feishu-word`，装好后不要删）；
3. Chrome / Edge 地址栏输入 `chrome://extensions`（Edge 为 `edge://extensions`），打开右上角「开发者模式」；
4. 点「加载已解压的扩展程序」，选择**解压出来的目录**（里面有 `manifest.json` 那一层）；
5. 打开任意飞书文档（`/wiki/*`、`/docx/*`），右下角出现蓝色悬浮按钮即安装成功。

> 升级：下载新版 zip → 覆盖解压到同一目录 → 回 `chrome://extensions` 点该扩展的「重新加载」。
> 卸载：在扩展页点「移除」，再删掉那个目录即可。

## 使用

1. 打开要导出的飞书文档，点右下角悬浮按钮唤出面板；
2. 按需选择纸张 / 页边距 / 字体 / 内嵌表格方式，点「开始导出」；
3. 导出过程会自动滚动全文采集懒加载内容（**请勿中途切换标签页或手动滚动**）；
4. 完成后文件落到浏览器下载目录，面板会显示页数与降级项。

> 导出只发送**只读**请求，不会修改原文档。

## 开发

```bash
npm install
npm run build     # esbuild 打包到 dist/
npm run watch     # 监听改动
npm run pack      # 打成 飞书文档导出Word扩展.zip（旧包自动备份到 output/zip-backup/）
npm test          # 离线渲染 + 16 条结构断言
```

开发时加载 `dist/` 目录；改完代码需要在 `chrome://extensions` 点「重新加载」，
**并刷新文档页**，否则页面里跑的还是旧内容脚本。

### 导出选项

| 选项 | 说明 | 默认值 |
| --- | --- | --- |
| 纸张 | A4 / Letter | A4 |
| 页边距 | 标准（左右 900 twips，版心约 674px）/ 窄 / 宽 | 标准 |
| 字体 | 雅黑 / 宋体 | 雅黑 |
| 包含目录 | 插入 Word 目录域，打开后 Ctrl+A、F9 更新页码 | 关闭 |
| 内嵌表格 | 「导出为图片」版式最忠实；「导出为可编辑表格」生成 Word 原生表格（列数过多时自动回退图片） | 图片 |

选项保存在浏览器 `localStorage`（键 `fde-options`），下次自动沿用。

---

## 文档索引

| 文档 | 内容 |
| --- | --- |
| **[docs/ARCHITECTURE.md](docs/ARCHITECTURE.md)** | 数据流、模块地图、关键设计决策与坑（**改代码前必读**） |
| **[docs/STYLE-SYSTEM.md](docs/STYLE-SYSTEM.md)** | Word 样式体系（Fde* 样式表）、字号/行距映射、如何在 Word 里改样式 |
| **[docs/VERIFICATION.md](docs/VERIFICATION.md)** | 三层验证体系、断言清单、命令与当前验收结果 |
| **[docs/TOOLS.md](docs/TOOLS.md)** | tools/ 下脚本的分类与用法（构建/导出/渲染/验证/探测） |
| **[docs/TROUBLESHOOTING.md](docs/TROUBLESHOOTING.md)** | 排障手册：Chrome 注入、画板采集、内存、面板卡死、编码乱码 |
| **[docs/CHANGELOG.md](docs/CHANGELOG.md)** | r001 → r020 迭代记录 |
| [docs/THIRD-PARTY.md](docs/THIRD-PARTY.md) | 第三方依赖与许可声明 |
| **[docs/COMPARISON.md](docs/COMPARISON.md)** | 与飞书官方「下载为 Word」的逐项对比（6 篇文档实测） |
| [samples/README.md](samples/README.md) | 样本数据与各轮产物说明 |

## 许可与隐私

- 代码许可：**MIT**（`LICENSE`），第三方依赖见 [docs/THIRD-PARTY.md](docs/THIRD-PARTY.md)
- 本项目是**非官方**第三方工具，与飞书官方无关联，不使用开放平台 API Key

隐私说明（扩展实际做了什么）：

- 只申请 `https://*.feishu.cn/*` 与 `https://*.larkoffice.com/*` 的主机权限
- **只发只读请求**：读取当前登录用户自己在看的文档块数据，不修改任何文档内容
- 所有处理在**本地浏览器内**完成，不向任何第三方服务器发送文档内容
- 不采集、不上传任何使用统计；选项仅存于本机 `localStorage`
- 文档落盘走浏览器原生下载，路径由你自己选择

> `samples/` 下的抓包样本含登录态请求，**已列入 `.gitignore`，请勿提交**。

---

## 目录结构

```
src/
  manifest.json                 扩展清单（MV3）
  background/service-worker.js  下载落盘、视口截图兜底、跨域抓取
  content/
    index.js                    入口：识别文档页、连后台、编排导出、挂载面板
    ui.js + ui.css              Shadow DOM 面板（选项/进度/结果）
    capture.js                  两遍扫描采集图片、内嵌表格 canvas、画板 canvas
  core/
    api.js                      wiki 解析 + client_vars 分页 + 统一请求/CSRF
    tree.js                     block_map → 块树
    text.js                     富文本 attribs 解析（*N 属性引用、+LEN 36 进制长度）
    export.js                   导出编排（选项 → 页面尺寸/版心 → 渲染 → 打包）
    render/                     index 分发 / paragraph 段落标题列表 / table 原生表格 / media 图片附件表格画板
    docx/                       build 组装 / styles 样式体系 / numbering 编号 / runs 行内样式
  shared/constants.js           块类型、TYPO 排版基准、纸张/边距/字体方案、样式 id 常量
  shared/util.js                颜色/尺寸换算等通用工具
tools/                          构建、导出、渲染、验证、探测脚本（见 docs/TOOLS.md）
samples/                        样本数据与各轮迭代产物（见 samples/README.md）
output/                         交付物：最新 docx + zip 备份
docs/                           上述文档
```

---

## 当前保真度（r020，845 块 / 25 页）

- **标题**：H1 三号(16pt) / H2 四号(14pt) / H3 小四(12pt)，黑色；编号按飞书 `data.seq === 'auto'` 判定，「附件/附录」类不编号
- **正文**：12pt 雅黑、单倍行距；粗体/斜体/删除线/行内代码/高亮底色/超链接均保留
- **表格**：原生可编辑，黑色 0.5pt 框线；首行按飞书 `header_row` 渲染为表头（灰底 + 跨页重复）
- **引用**：挂到 Word 内建「引用」样式，黑色左竖线，内部列表完整
- **代码块**：等宽字体 + 灰底，**标题（caption）在代码块上方**
- **图片 / 画板 / 内嵌电子表格**：按页宽等比缩放居中

---

## 已知限制

1. **内嵌电子表格仍为截图**（Word 内不可编辑）：单元格数据接口已探通（`client_vars` 必须带 `memberId`），但快照内是私有二进制格式，无法可靠还原行列（详见 [docs/ARCHITECTURE.md](docs/ARCHITECTURE.md)）
2. **附件（如 html）以链接形式保留**：docx 库无 OLE 封装 API，真内嵌需自实现 OLE Compound File
3. **画板走 canvas 采集**：依赖页面已完成懒渲染，真人浏览文档后再导出最稳；自动化场景可能降级为占位
4. 多维表格（bitable）以链接形式保留

---

## 常用命令

```bash
npm run build                      # 构建
npm run pack                       # 打包 zip（自动备份旧包）

# 自动化迭代（需要 Chrome 已在跑，见 docs/TOOLS.md）
node tools/serve-chrome.mjs        # 后台常驻启动 Chrome（必须，否则进程会被回收）
node tools/export-run.mjs r021     # 真实加载扩展跑一次导出
node tools/check-styles.mjs <docx> # 样式/结构断言（14 项）
node tools/docx2pdf.ps1 -Src <docx> -Dst <pdf>   # Word COM 转 PDF
```

## 环境要求

- Node 22+（构建）
- Chrome 137+（已禁用 `--load-extension`，自动化装载需用 CDP `Extensions.loadUnpacked`）
- 仅做视觉验证时需要：本机 Word（COM 转 PDF）+ Python pymupdf（装在隔离 venv）

> 仓库尚未纳入版本控制，建议先 `git init`（`.gitignore` 已备好）。
