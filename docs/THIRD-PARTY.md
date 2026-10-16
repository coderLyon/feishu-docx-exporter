# 第三方依赖与许可

本项目自身采用 **MIT**（见仓库根目录 `LICENSE`）。

## 运行时依赖

| 依赖 | 版本 | 许可 | 用途 |
| --- | --- | --- | --- |
| [docx](https://github.com/dolanmiu/docx) | ^9.11.0 | MIT | 生成 OOXML（.docx）文档 |
| [jszip](https://github.com/Stuk/jszip) | ^3.10.1 | MIT（另有 GPLv3 双许可可选） | 打包 zip（用于 npm run pack） |

## 开发依赖

| 依赖 | 版本 | 许可 | 用途 |
| --- | --- | --- | --- |
| [esbuild](https://github.com/evanw/esbuild) | ^0.25.10 | MIT | 打包 content script 与 service worker |
| [playwright](https://github.com/microsoft/playwright) | ^1.56.0 | Apache-2.0 | 早期研究阶段用于驱动浏览器（非运行时依赖） |

## 验证环节的外部工具（不随代码分发）

- **Microsoft Word**（COM 自动化）：仅在本机做视觉验证时用于 docx → PDF
- **PyMuPDF**（AGPL-3.0，可选商用授权）：仅用于把 PDF 渲染为 PNG 与反查字号/颜色，**不链接进产物、不随扩展分发**

## 与飞书（Lark）的关系

本项目是**非官方**的第三方工具，与飞书官方无任何关联，未使用飞书开放平台的官方 API Key，
仅复用已在浏览器中登录的会话读取自己的文档数据（只读）。请遵守飞书的服务条款，
仅导出你有权访问与使用的文档。
