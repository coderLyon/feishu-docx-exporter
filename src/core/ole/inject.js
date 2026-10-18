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
import { makeIconPng } from './icon.js';

/** 占位标记：渲染阶段写这段文字，后处理时整段替换成 OLE 对象 */
export const OLE_MARKER = '\u0001FDE_OLE_';

function oleObjectXml(shapeId, oleRelId, imgRelId, sizePt = 48) {
  return `<w:r><w:object w:dxaOrig="${sizePt * 20}" w:dyaOrig="${sizePt * 20}">
<v:shapetype id="_x0000_t75" coordsize="21600,21600" o:spt="75" o:preferrelative="t" path="m@4@5l@4@11@9@11@9@5xe" filled="f" stroked="f">
<v:stroke joinstyle="miter"/><v:formulas><v:f eqn="if lineDrawn pixelLineWidth 0"/><v:f eqn="sum @0 1 0"/><v:f eqn="sum 0 0 @1"/><v:f eqn="prod @2 1 2"/><v:f eqn="prod @3 21600 pixelWidth"/><v:f eqn="prod @3 21600 pixelHeight"/><v:f eqn="sum @0 0 1"/><v:f eqn="prod @6 1 2"/><v:f eqn="prod @7 21600 pixelWidth"/><v:f eqn="sum @8 21600 0"/><v:f eqn="prod @7 21600 pixelHeight"/><v:f eqn="sum @10 21600 0"/></v:formulas>
<v:path o:extrusionok="f" gradientshapeok="t" o:connecttype="rect"/><o:lock v:ext="edit" aspectratio="t"/>
</v:shapetype>
<v:shape id="${shapeId}" type="#_x0000_t75" style="width:${sizePt}pt;height:${sizePt}pt" o:ole="">
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
    zip.file(`word/embeddings/oleObject${n}.bin`, buildOleBin(att.bytes, att.name));
    // 按扩展名现画图标（无法生成时退回内置底板）
    let icon = null;
    try {
      icon = await makeIconPng(att.name);
    } catch {
      icon = new Uint8Array([0]);
    }
    if (icon.length > 1) zip.file(`word/media/fdeole${n}.png`, icon);

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
    const pPr = para.match(/<w:pPr>[\s\S]*?<\/w:pPr>/);
    const shapeId = `fdeOleShape${n}`;
    doc =
      doc.slice(0, pStart) +
      `<w:p>${pPr ? pPr[0] : ''}${oleObjectXml(shapeId, oleRel, imgRel)}</w:p>` +
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
  return zip.generateAsync({ type: 'blob', compression: 'DEFLATE' });
}
