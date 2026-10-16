# 排障手册

## 一、扩展装载与 content script 注入

| 现象 | 原因与处理 |
| --- | --- |
| `chrome://extensions` 加载 `dist/` 报「清单文件无效」 | 检查三处：① `default_popup` 不能是空串（不需要 popup 就删掉该字段）；② 声明了 `default_locale` 就必须有 `_locales/`；③ JSON 不能有注释 |
| 自动化时提示「content script 未注入」 | 顺序必须是**先 `Extensions.loadUnpacked`、再 `Page.reload`**。颠倒顺序永远注入不上 |
| 注入后仍报「未找到开始导出按钮」 | 面板停在上一次的「导出完成」结果态 → 再发两次 `fde-toggle` 事件回到菜单视图。**不要用 reload 重置**（会清掉懒渲染内容） |
| Chrome 137+ 无法用 `--load-extension` | 改用 CDP `Extensions.loadUnpacked`（启动参数需 `--enable-unsafe-extension-debugging`） |
| Chrome 默认 profile 拒绝远程调试 | 必须指定独立 `--user-data-dir`（本项目用 `.chrome-user`） |

## 二、采集相关

| 现象 | 原因与处理 |
| --- | --- |
| 画板全部丢失（0 张） | 不要用「按类名找容器」的定点采集——会命中 `whiteboard-protal-container` 这个 **0 高占位容器**（永远没有 canvas）。正确做法：扫描 canvas，用 `canvas.closest('[class*="whiteboard"]')` 反查真实容器 |
| 画板有重影 | 两个独立成因，都要处理：① 合成了两层等大 canvas（用墨量探测只选有内容的那一层）；② 拼接时按滚动容器切块，容器比画布略大就切成 2 块，同一幅图画两次错位叠加（白板走 `singleTile` 单块拼接） |
| 画板/图片采集为空白 | 懒渲染未完成。内容层空白时逐次加长等待重试（1200 / 2000ms）；仍空白则回退视口截图 |
| 图片与文档位置对不上 | 资源按下标顺序消费，采集顺序必须等于文档出现顺序。检查两遍扫描的 `docY` 去重与排序，以及抓完是否恢复了 `scrollTop` |
| 内嵌电子表格只有一部分 | 表格内容确实超出可见画布，需滚动拼接（与白板的单块拼接不同，不要误用 `singleTile`） |

## 三、渲染 / 样式

| 现象 | 原因与处理 |
| --- | --- |
| 标题编号与标题文字字号不一致 | 标题段落上不应有任何 run 级 `size`/`font`。检查是否漏加；编号 run 也应继承样式，不写死字号 |
| 在 Word 里改样式全文没反应 | run 级直接格式覆盖了样式 → 跑 `node tools/check-styles.mjs <docx>` 看 B6/B7 |
| 正文行距偏大 | 雅黑自然行高 ≈ 1.714em，`auto` 倍数会乘在它上面。统一用单倍 `{line:240, lineRule:'auto'}`，且只写在样式里 |
| 引用块内容为空（只剩竖线） | docx@9 的 `Paragraph` **没有 `options` 属性**，`element.options?.children` 恒为 undefined → 重建段落会变空。正确做法是用 `quoteDepth` 标志让子渲染器直接产出带样式的段落，不要重建 |
| 附件/附录被误加编号 | 必须用飞书 `data.seq === 'auto'` 判定，不要用「附件」文本正则猜测 |
| 「附件C 建表语句」标题居中 | 代码标题与图片题注不能共用样式：图片题注 `FdeCaption`（居中），代码标题 `FdeCodeCaption`（左对齐） |
| 在线黄色底色丢失 | 飞书属性名是 `textHighlightBackground`（值 `rgba(255,246,122,0.8)`），不是 `background-color`/`highlight`，需在 `ATTR_ALIASES` 里映射 |
| 表格首行被误当成表头 | 按 `data.header_row` 判断：`false` 的表首行是跨列合并的分节行，不能当表头；字段缺失时用「首行无列向合并 + 每格非空」兜底 |
| 相邻代码块粘在一起 | 两个 caption 为空的相邻 code 块会渲染成无缝相接的表 → 代码块后补间距段落 |

## 四、Word / PDF / 编码

| 现象 | 原因与处理 |
| --- | --- |
| **Word 说「无法打开/文件损坏」，一点都打不开** | 正文里夹带 XML 非法控制字符（0x00-0x08、0x0B、0x0C、0x0E-0x1F、U+FFFE/FFFF）。实测某模板正文含一个 `0x08` 退格符，原样写进 `<w:t>` 就废了。修复：`sanitizeXmlText()`（`src/shared/util.js`）在文本抽取咽喉处净化（`text.js` 的富文本解析 + `runs.js` 兜底 + 附件名）。**守卫**：`check-styles.mjs` 的 `F1` 会逐部件扫描 XML 字符合法性 |
| Word 打开提示「文档需要修复」 | 常见三种：① 一个 `pPr` 里出现两个 `pStyle`（如自定义样式 id 又叠加了 ListParagraph）；② 引用了 styles.xml 中不存在的样式 id；③ 非法枚举值（如 `w:jc w:val="1"`，应为字符串 `"center"`）。`check-styles.mjs` 的 B3/B8 专门查这些 |
| PowerShell 脚本里的中文变乱码 | `.ps1` 无 UTF-8 BOM 时 PowerShell 5.1 会按本地编码解析 → 脚本内避免使用中文字面量，改用 Word 常量（如 `Styles.Item(-2)` 表示「标题 1」） |
| Word COM 调用报「无法转换参数」 | COM 参数**不要加 `[ref]`**；`ExportAsFixedFormat` 的 Range 必须是 `0`（`wdExportAllDocument`），传 1 会报「值不在预期范围」 |
| `ExportAsFixedFormat` 失败 | 改用 `$doc.SaveAs2($path, 17)`（17 = wdFormatPDF），本项目 `docx2pdf.ps1` 已用此法 |
| 拿不到落盘路径 | download 回调只给 id，且 `conflictAction` 可能改名 → 轮询 `chrome.downloads.search` 直到 `complete` |

## 五、环境与资源

| 现象 | 原因与处理 |
| --- | --- |
| Chrome 关不掉 | ① 后台常驻任务（serve-chrome）会重启 Chrome → 先停任务；② Chrome 后台应用保活 → 循环 `taskkill /F /IM chrome.exe` 2-3 轮 |
| 磁盘被 Chrome profile 占满 | `.chrome-user/OptGuideOnDeviceModel` 是 Chrome 端侧 AI 模型缓存，可达数 GB，删掉不影响使用；该目录已加入 `.gitignore` |
| 内存不足 / CDP 超时 | 8GB 机器上 Chrome、Word COM、pymupdf 必须串行；导出完先关 Chrome 再转 PDF；渲染 PNG 时逐页释放 |
| `samples/` 体积膨胀 | 每轮迭代的全页 PNG 约 12MB，旧轮次只保留 `source.docx` + `export.json` 即可（`.gitignore` 已忽略 `samples/runs/*/word/`） |
