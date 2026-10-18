/**
 * 由「Word 生成的 OLE 模板」复刻出可嵌入任意文件的 oleObject.bin。
 *
 *   node tools/build-ole.mjs <模板bin> <要嵌入的文件> <输出bin> [显示名]
 *
 * Ole10Native 布局（从真实模板逆向确认）：
 *   u32 totalSize | u16 flags(2) | cstr label | u16 0 | u16 3 |
 *   u32 tempPathLen | bytes tempPath | u32 dataSize | bytes data
 */
import fs from 'node:fs';
import CFB from 'cfb';

export function buildOleFromTemplate(templatePath, fileBytes, fileName) {
  const cfb = CFB.read(fs.readFileSync(templatePath), { type: 'buffer' });
  const pick = (name) => {
    const e = cfb.FileIndex.find((f) => f.name === name);
    if (!e) throw new Error(`模板缺少流 ${JSON.stringify(name)}`);
    return Buffer.from(e.content);
  };
  const compObj = pick('\u0001CompObj');
  const objInfo = pick('\u0003ObjInfo');
  const tpl = pick('\u0001Ole10Native');

  // 解析模板里的 label / tempPath（沿用它们的结构，只换文件名与数据）
  let i = 4 + 2;
  const readStr = () => {
    const s = i;
    while (i < tpl.length && tpl[i] !== 0) i += 1;
    const v = tpl.toString('latin1', s, i);
    i += 1;
    return v;
  };
  readStr(); // 原 label（短名），丢弃
  readStr(); // 原 fileName（Word 缓存里的完整路径），丢弃
  const zero = tpl.readUInt16LE(i);
  i += 2;
  const unk = tpl.readUInt16LE(i);
  i += 2;
  const tempLen = tpl.readUInt32LE(i);
  i += 4;
  const tempPath = tpl.subarray(i, i + tempLen);
  i += tempLen;
  const tplDataSize = tpl.readUInt32LE(i);

  // 文件名转 ANSI（非 ASCII 用 ? 代替，和 Windows Packager 的行为一致）
  const ansi = (s) => Buffer.from([...s].map((ch) => (ch.charCodeAt(0) < 128 ? ch : '?')).join(''), 'latin1');

  const label = Buffer.concat([ansi(fileName), Buffer.from([0])]);
  // 第二个字符串是完整路径（Packager 用它决定解压目录），给个中性路径
  const fullPath = Buffer.concat([ansi('C:\\Temp\\' + fileName), Buffer.from([0])]);

  // 模板的 tempPath 末段是**原文件名**，必须先切掉再拼我们的名字，
  // 否则会变成 "...????.htmlreadme.txt" —— 路径非法，双击激活直接失败（实测）
  const tplTempStr = tempPath.toString('latin1').replace(/\u0000+$/, '');
  const cut = tplTempStr.lastIndexOf('\\');
  const tempStr = (cut >= 0 ? tplTempStr.slice(0, cut + 1) : 'C:////Temp////') + fileName;
  const temp = Buffer.concat([ansi(tempStr), Buffer.from([0])]);

  const dataSize = fileBytes.length;
  const head = Buffer.alloc(4 + 2);
  head.writeUInt16LE(2, 4);
  const mid1 = Buffer.alloc(2);
  mid1.writeUInt16LE(zero, 0);
  const mid2 = Buffer.alloc(2);
  mid2.writeUInt16LE(unk, 0);
  const tempLenBuf = Buffer.alloc(4);
  tempLenBuf.writeUInt32LE(temp.length, 0);
  const sizeBuf = Buffer.alloc(4);
  sizeBuf.writeUInt32LE(dataSize, 0);

  // 数据之后还有一段：u32 长度 + UTF-16LE 的临时路径（Packager 用它决定解压到哪里）。
  // 少了这段，Word 认得对象（ClassType=Package）但双击激活会失败 —— 实测对照得出。
  const tempUtf16 = Buffer.from(temp.toString('latin1').replace(/\0$/, ''), 'utf16le');
  const utf16LenBuf = Buffer.alloc(4);
  utf16LenBuf.writeUInt32LE(tempUtf16.length, 0);

  const body = Buffer.concat([
    head.subarray(4),
    label,
    fullPath,
    mid1,
    mid2,
    tempLenBuf,
    temp,
    sizeBuf,
    fileBytes,
    utf16LenBuf,
    tempUtf16,
  ]);
  const out = Buffer.concat([Buffer.alloc(4), body]);
  out.writeUInt32LE(body.length, 0);

  const streams = [
    { name: '\u0001CompObj', data: compObj },
    { name: '\u0003ObjInfo', data: objInfo },
    { name: '\u0001Ole10Native', data: out },
  ];
  const newCfb = CFB.utils.cfb_new();
  // 根存储的 CLSID 必须是 Package 类标识，否则 Word 认得对象但**激活会失败**
  // （模板里是 0c000300-0000-0000-c000-000000000046）
  newCfb.FileIndex[0].clsid = '0c00030000000000c000000000000046';
  for (const s of streams) CFB.utils.cfb_add(newCfb, s.name, s.data);
  // cfb 库默认会塞一个 \x01Sh33tJ5 标记流和几个空条目，Word 不需要，清掉
  try {
    CFB.utils.cfb_del(newCfb, '\u0001Sh33tJ5');
  } catch {
    /* 没有就忽略 */
  }
  return {
    bin: Buffer.from(CFB.write(newCfb, { type: 'buffer' })),
    meta: { tplDataSize, dataSize, labelLen: label.length, tempLen: temp.length },
  };
}

// CLI
if (process.argv[1]?.endsWith('build-ole.mjs')) {
  const [tpl, file, outBin, display] = process.argv.slice(2);
  const bytes = fs.readFileSync(file);
  const name = display || file.split(/[\\/]/).pop();
  const { bin, meta } = buildOleFromTemplate(tpl, bytes, name);
  fs.writeFileSync(outBin, bin);
  console.log('已生成', outBin, bin.length, '字节', JSON.stringify(meta));
}
