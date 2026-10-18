/**
 * docx 后处理：把 OLE 附件注入已打包好的 docx。
 *
 * docx 库不支持 OLE，所以在 Packer 产出之后再改 zip：
 *   1. word/embeddings/oleObjectN.bin
 *   2. 图标 word/media/fdeole*.png
 *   3. word/_rels/document.xml.rels 增加两条关系
 *   4. [Content_Types].xml 声明 bin / png
 *   5. 把正文里的占位文本替换成 <w:object> 结构
 */
import JSZip from 'jszip';
import { buildOleBin } from './build.js';
import { makeFileCardPng, makeIconPng } from './icon.js';

/** 占位标记：渲染阶段写这段文字，后处理时整段替换成 OLE 对象 */
export const OLE_MARKER = '\u0001FDE_OLE_';

/** 内嵌表格（Excel 对象）的占位标记 */
export const SHEET_MARKER = '\u0001FDE_XLSX_';

/** docx 的 MIME，任何后处理都必须保持它，否则下载后缀会变成 .zip */
export const DOCX_MIME =
  'application/vnd.openxmlformats-officedocument.wordprocessingml.document';

function oleObjectXml(shapeId, oleRelId, imgRelId, widthPt = 48, heightPt = 48) {
  return `<w:r><w:object w:dxaOrig="${Math.round(widthPt * 20)}" w:dyaOrig="${Math.round(heightPt * 20)}">
<v:shapetype id="_x0000_t75" coordsize="21600,21600" o:spt="75" o:preferrelative="t" path="m@4@5l@4@11@9@11@9@5xe" filled="f" stroked="f">
<v:stroke joinstyle="miter"/><v:formulas><v:f eqn="if lineDrawn pixelLineWidth 0"/><v:f eqn="sum @0 1 0"/><v:f eqn="sum 0 0 @1"/><v:f eqn="prod @2 1 2"/><v:f eqn="prod @3 21600 pixelWidth"/><v:f eqn="prod @3 21600 pixelHeight"/><v:f eqn="sum @0 0 1"/><v:f eqn="prod @6 1 2"/><v:f eqn="prod @7 21600 pixelWidth"/><v:f eqn="sum @8 21600 0"/><v:f eqn="prod @7 21600 pixelHeight"/><v:f eqn="sum @10 21600 0"/></v:formulas>
<v:path o:extrusionok="f" gradientshapeok="t" o:connecttype="rect"/><o:lock v:ext="edit" aspectratio="t"/>
</v:shapetype>
<v:shape id="${shapeId}" type="#_x0000_t75" style="width:${widthPt.toFixed(1)}pt;height:${heightPt.toFixed(1)}pt" o:ole="">
<v:imagedata r:id="${imgRelId}" o:title=""/></v:shape>
<o:OLEObject Type="Embed" ProgID="Package" ShapeID="${shapeId}" DrawAspect="Icon" ObjectID="_fde${shapeId}" r:id="${oleRelId}"/>
</w:object></w:r>`;
}

/**
 * @param {Blob} blob Packer 产出的 docx
 * @param {Array<{index:number, name:string, bytes:Uint8Array}>} attachments
 * @returns {Promise<Blob>}
 */
export async function injectOleObjects(blob, attachments) {
  if (!attachments.length) return blob;

  const zip = await JSZip.loadAsync(await blob.arrayBuffer());
  let doc = await zip.file('word/document.xml').async('string');
  let rels = await zip.file('word/_rels/document.xml.rels').async('string');
  let ct = await zip.file('[Content_Types].xml').async('string');

  const relAdds = [];
  const ctAdds = [];

  for (const att of attachments) {
    const n = att.index;
    const oleBytes = buildOleBin(att.bytes, att.name);
    // 大附件用 STORE：本身多是压缩包/二进制，再 DEFLATE 既慢又几乎没收益
    zip.file(`word/embeddings/oleObject${n}.bin`, oleBytes, {
      compression: oleBytes.length > 4 * 1024 * 1024 ? 'STORE' : 'DEFLATE',
    });
    // 卡片式预览：长方形 + 类型徽标 + 文件名 + 大小（与飞书在线一致）
    let card = null;
    try {
      card = await makeFileCardPng(att.name, att.size ?? 0);
    } catch {
      // 退化：只画类型图标
      try {
        const icon = await makeIconPng(att.name);
        card = { bytes: icon, widthPt: 48, heightPt: 48 };
      } catch {
        card = null;
      }
    }
    if (card?.bytes?.length > 1) zip.file(`word/media/fdeole${n}.png`, card.bytes);

    const oleRel = `rIdFdeOle${n}`;
    const imgRel = `rIdFdeOleImg${n}`;
    relAdds.push(
      `<Relationship Id="${oleRel}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/oleObject" Target="embeddings/oleObject${n}.bin"/>`,
      `<Relationship Id="${imgRel}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/fdeole${n}.png"/>`,
    );

    // 替换占位段
    const marker = `${OLE_MARKER}${n}\u0001`;
    const at = doc.indexOf(marker);
    if (at < 0) continue;
    const pStart = Math.max(doc.lastIndexOf('<w:p ', at), doc.lastIndexOf('<w:p>', at));
    const pEnd = doc.indexOf('</w:p>', at) + '</w:p>'.length;
    const para = doc.slice(pStart, pEnd);
    let pPr = para.match(/<w:pPr>[\s\S]*?<\/w:pPr>/)?.[0] ?? '';
    // 卡片居中：与飞书在线的附件块一致（原 pPr 里的对齐一律覆盖）
    pPr = pPr.replace(/<w:jc[^/]*\/>/g, '');
    const centeredPr = pPr
      ? pPr.replace('</w:pPr>', '<w:jc w:val="center"/></w:pPr>')
      : '<w:pPr><w:jc w:val="center"/></w:pPr>';
    const shapeId = `fdeOleShape${n}`;
    const cardW = card?.widthPt ?? 48;
    const cardH = card?.heightPt ?? 48;
    doc =
      doc.slice(0, pStart) +
      `<w:p>${centeredPr}${oleObjectXml(shapeId, oleRel, imgRel, cardW, cardH)}</w:p>` +
      doc.slice(pEnd);
  }

  if (relAdds.length) {
    rels = rels.replace('</Relationships>', `${relAdds.join('')}</Relationships>`);
    zip.file('word/_rels/document.xml.rels', rels);
  }

  if (!ct.includes('Extension="bin"')) {
    ctAdds.push('<Default Extension="bin" ContentType="application/vnd.openxmlformats-officedocument.oleObject"/>');
  }
  if (!ct.includes('Extension="png"')) {
    ctAdds.push('<Default Extension="png" ContentType="image/png"/>');
  }
  if (ctAdds.length) {
    ct = ct.replace('</Types>', `${ctAdds.join('')}</Types>`);
    zip.file('[Content_Types].xml', ct);
  }

  zip.file('word/document.xml', doc);
  // ⚠ 必须显式指定 docx MIME：JSZip 默认是 application/zip，
  // 而 Chrome 下载会按 Blob 的 MIME 覆盖扩展名 —— 含附件的文档会落盘成 .zip（实测）
  return zip.generateAsync({
    type: 'blob',
    mimeType: DOCX_MIME,
    compression: 'DEFLATE',
  });
}

/** 内嵌 Excel 对象的 <w:object>（结构与飞书官方导出一致：package 关系 + Excel.Sheet.12 + 预览图） */
function xlsxObjectXml({ shapeId, oleRelId, imgRelId, widthPt, heightPt, title }) {
  const dxa = Math.round(widthPt * 20);
  const dya = Math.round(heightPt * 20);
  return `<w:r><w:object w:dxaOrig="${dxa}" w:dyaOrig="${dya}">
<v:shapetype id="_x0000_t75" coordsize="21600,21600" o:spt="75" o:preferrelative="t" path="m@4@5l@4@11@9@11@9@5xe" filled="f" stroked="f">
<v:stroke joinstyle="miter"/><v:formulas><v:f eqn="if lineDrawn pixelLineWidth 0"/><v:f eqn="sum @0 1 0"/><v:f eqn="sum 0 0 @1"/><v:f eqn="prod @2 1 2"/><v:f eqn="prod @3 21600 pixelWidth"/><v:f eqn="prod @3 21600 pixelHeight"/><v:f eqn="sum @0 0 1"/><v:f eqn="prod @6 1 2"/><v:f eqn="prod @7 21600 pixelWidth"/><v:f eqn="sum @8 21600 0"/><v:f eqn="prod @7 21600 pixelHeight"/><v:f eqn="sum @10 21600 0"/></v:formulas>
<v:path o:extrusionok="f" gradientshapeok="t" o:connecttype="rect"/><o:lock v:ext="edit" aspectratio="t"/>
</v:shapetype>
<v:shape id="${shapeId}" type="#_x0000_t75" style="width:${widthPt.toFixed(1)}pt;height:${heightPt.toFixed(1)}pt" o:ole="">
<v:imagedata r:id="${imgRelId}" o:title="${title}"/></v:shape>
<o:OLEObject Type="Embed" ProgID="Excel.Sheet.12" ShapeID="${shapeId}" DrawAspect="Icon" ObjectID="_${shapeId}" r:id="${oleRelId}"/>
</w:object></w:r>`;
}

/**
 * 把内嵌表格以「Excel 对象」注入 docx。
 *
 * @param {Blob} blob Packer 产出
 * @param {Array<{index:number, bytes:Uint8Array, name:string,
 *   preview?:Uint8Array, widthPt:number, heightPt:number}>} sheets
 */
export async function injectXlsxObjects(blob, sheets) {
  if (!sheets.length) return blob;

  const zip = await JSZip.loadAsync(await blob.arrayBuffer());
  let doc = await zip.file('word/document.xml').async('string');
  let rels = await zip.file('word/_rels/document.xml.rels').async('string');
  let ct = await zip.file('[Content_Types].xml').async('string');

  const relAdds = [];
  for (const sh of sheets) {
    const n = sh.index;
    const fileName = `Microsoft_Excel_Worksheet${n}.xlsx`;
    zip.file(`word/embeddings/${fileName}`, sh.bytes);
    // 已压缩过的 xlsx 不再二次压缩
    zip.file(`word/media/fdexlsx${n}.png`, sh.preview ?? new Uint8Array([0]), { compression: 'STORE' });

    const oleRel = `rIdFdeXlsx${n}`;
    const imgRel = `rIdFdeXlsxImg${n}`;
    relAdds.push(
      `<Relationship Id="${oleRel}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/package" Target="embeddings/${fileName}"/>`,
      `<Relationship Id="${imgRel}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/fdexlsx${n}.png"/>`,
    );

    const marker = `${SHEET_MARKER}${n}\u0001`;
    const at = doc.indexOf(marker);
    if (at < 0) continue;
    const pStart = Math.max(doc.lastIndexOf('<w:p ', at), doc.lastIndexOf('<w:p>', at));
    const pEnd = doc.indexOf('</w:p>', at) + '</w:p>'.length;
    const para = doc.slice(pStart, pEnd);
    const pPr = para.match(/<w:pPr>[\s\S]*?<\/w:pPr>/);
    const shapeId = `fdeXlsxShape${n}`;
    doc =
      doc.slice(0, pStart) +
      `<w:p>${pPr ? pPr[0] : ''}${xlsxObjectXml({
        shapeId,
        oleRelId: oleRel,
        imgRelId: imgRel,
        widthPt: sh.widthPt,
        heightPt: sh.heightPt,
        title: sh.name ?? '',
      })}</w:p>` +
      doc.slice(pEnd);
  }

  if (relAdds.length) {
    rels = rels.replace('</Relationships>', `${relAdds.join('')}</Relationships>`);
    zip.file('word/_rels/document.xml.rels', rels);
  }
  if (!ct.includes('Extension="xlsx"')) {
    ct = ct.replace(
      '</Types>',
      '<Default Extension="xlsx" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"/></Types>',
    );
    zip.file('[Content_Types].xml', ct);
  }
  zip.file('word/document.xml', doc);

  return zip.generateAsync({ type: 'blob', mimeType: DOCX_MIME, compression: 'DEFLATE' });
}
