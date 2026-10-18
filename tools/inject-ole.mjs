/**
 * 把 OLE 嵌入对象注入 docx（验证 OLE 结构是否被 Word 接受）。
 *
 *   node tools/inject-ole.mjs <源docx> <输出docx> <要嵌入的文件> [显示名]
 *
 * 做的事：
 *   1. 加入 word/embeddings/oleObject1.bin（由模板复刻）
 *   2. 加入图标 word/media/oleicon.emf（沿用模板里的 Packager 图标）
 *   3. 补 rels：oleObject + image
 *   4. 补 [Content_Types].xml 的 bin / emf 声明
 *   5. 把正文里的占位字样 [[OLE]] 替换成 <w:object> 结构
 */
import fs from 'node:fs';
import path from 'node:path';
import JSZip from 'jszip';
import { buildOleFromTemplate } from './build-ole.mjs';

const [srcDocx, outDocx, attachFile, displayName] = process.argv.slice(2);
if (!srcDocx || !outDocx || !attachFile) {
  console.error('用法: node tools/inject-ole.mjs <源docx> <输出docx> <附件> [显示名]');
  process.exit(2);
}

const TEMPLATE_DOCX = 'C:\\Users\\Admin\\Downloads\\IIoT试验任务接口开发文档V2.0 副本 (28).docx';
const TEMPLATE_BIN = 'samples/out/ole-template.bin';

const zip = await JSZip.loadAsync(fs.readFileSync(srcDocx));
const tzip = await JSZip.loadAsync(fs.readFileSync(TEMPLATE_DOCX));

// 1) OLE 二进制
const fileBytes = fs.readFileSync(attachFile);
const name = displayName || path.basename(attachFile);
const { bin } = buildOleFromTemplate(TEMPLATE_BIN, fileBytes, name);
zip.file('word/embeddings/oleObject1.bin', bin);

// 2) 图标（直接用模板里的 Packager 图标 EMF）
const emf = await tzip.file('word/media/image6.emf')?.async('nodebuffer');
if (emf) zip.file('word/media/oleicon.emf', emf);

// 3) rels
let rels = await zip.file('word/_rels/document.xml.rels').async('string');
if (!rels.includes('oleObject1.bin')) {
  rels = rels.replace('</Relationships>',
    '<Relationship Id="rIdOle1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/oleObject" Target="embeddings/oleObject1.bin"/>' +
    (emf ? '<Relationship Id="rIdOleImg" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/oleicon.emf"/>' : '') +
    '</Relationships>');
  zip.file('word/_rels/document.xml.rels', rels);
}

// 4) content types
let ct = await zip.file('[Content_Types].xml').async('string');
if (!ct.includes('Extension="bin"')) {
  ct = ct.replace('</Types>', '<Default Extension="bin" ContentType="application/vnd.openxmlformats-officedocument.oleObject"/></Types>');
}
if (emf && !ct.includes('Extension="emf"')) {
  ct = ct.replace('</Types>', '<Default Extension="emf" ContentType="image/x-emf"/></Types>');
}
zip.file('[Content_Types].xml', ct);

// 5) 正文：把占位段替换成 OLE 对象
let doc = await zip.file('word/document.xml').async('string');
// 直接照抄模板里的 <w:object> 结构（含完整 shapetype/formulas/stroke），
// 只替换两个 rId —— 自己简化的 VML Word 不渲染（实测图标消失）。
const tplDoc = await tzip.file('word/document.xml').async('string');
const objStart = tplDoc.indexOf('<w:object');
const objEnd = tplDoc.indexOf('</w:object>', objStart) + '</w:object>'.length;
if (objStart < 0) throw new Error('模板里找不到 w:object');
const tplObject = tplDoc
  .slice(objStart, objEnd)
  .replace(/r:id="rId12"/g, 'r:id="rIdOleImg"')
  .replace(/r:id="rId13"/g, 'r:id="rIdOle1"')
  .replace(/id="_x0000_i1025"/g, 'id="oleShape1"')
  .replace(/ShapeID="_x0000_i1025"/g, 'ShapeID="oleShape1"');
const shapeXml = `<w:r>${tplObject}</w:r>`;

const idx = doc.indexOf('[[OLE]]');
if (idx < 0) {
  console.log('⚠ 正文里没有 [[OLE]] 占位，跳过替换');
} else {
  // 把占位所在的整段 run 段替换掉
  const pStart = doc.lastIndexOf('<w:p ', idx) >= 0 ? doc.lastIndexOf('<w:p ', idx) : doc.lastIndexOf('<w:p>', idx);
  const pEnd = doc.indexOf('</w:p>', idx) + 6;
  const para = doc.slice(pStart, pEnd);
  const pPrMatch = para.match(/<w:pPr>[\s\S]*?<\/w:pPr>/);
  const newPara = `<w:p>${pPrMatch ? pPrMatch[0] : ''}${shapeXml}</w:p>`;
  doc = doc.slice(0, pStart) + newPara + doc.slice(pEnd);
  zip.file('word/document.xml', doc);
  console.log('✓ 已替换占位段为 OLE 对象');
}

const outBuf = await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' });
fs.writeFileSync(outDocx, outBuf);
console.log(`已生成 ${outDocx}（${(outBuf.length / 1024).toFixed(0)} KB，嵌入 ${name} ${(fileBytes.length / 1024).toFixed(1)} KB）`);
