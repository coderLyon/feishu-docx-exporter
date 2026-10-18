/**
 * 在浏览器里生成 OLE 复合文档（Ole10Native），把附件真正嵌入 docx。
 *
 * 结构来自真实 Word 文档逆向（见 docs/ARCHITECTURE.md 5.2）：
 *   CFB(root CLSID = Package) { \x01CompObj, \x03ObjInfo, \x01Ole10Native }
 *
 * 依赖 cfb（纯 JS，可打包进扩展）。
 */
import { Buffer } from 'buffer';
import CFB from 'cfb';

// cfb 内部依赖 Node 的 Buffer，浏览器里没有 —— 在模块加载时补上。
// （实测缺失时报 "Buffer is not defined"，整个附件嵌入失败）
if (typeof globalThis.Buffer === 'undefined') {
  globalThis.Buffer = Buffer;
}

/** 模板里的类信息与对象标志，直接内联（76B / 6B） */
const COMP_OBJ = Uint8Array.from(
  atob('AQD+/wMKAAD/////DAADAAAAAADAAAAAAAAARgwAAABPTEUgUGFja2FnZQAAAAAACAAAAFBhY2thZ2UA9DmycQAAAAAAAAAAAAAAAA=='),
  (c) => c.charCodeAt(0),
);
const OBJ_INFO = Uint8Array.from(atob('QAADAAEA'), (c) => c.charCodeAt(0));
const PACKAGE_CLSID = '0c00030000000000c000000000000046';

/** 非 ASCII 字符按 Windows Packager 的行为替换成 ?（保持 ANSI 语义） */
function ansi(str) {
  const bytes = [];
  for (const ch of str) {
    const code = ch.codePointAt(0);
    bytes.push(code < 128 ? code : 0x3f);
  }
  bytes.push(0);
  return new Uint8Array(bytes);
}

function utf16(str) {
  const out = new Uint8Array(str.length * 2);
  for (let i = 0; i < str.length; i += 1) {
    const code = str.charCodeAt(i);
    out[i * 2] = code & 0xff;
    out[i * 2 + 1] = (code >> 8) & 0xff;
  }
  return out;
}

function u32(value) {
  const b = new Uint8Array(4);
  new DataView(b.buffer).setUint32(0, value, true);
  return b;
}

function u16(value) {
  const b = new Uint8Array(2);
  new DataView(b.buffer).setUint16(0, value, true);
  return b;
}

function concat(chunks) {
  const total = chunks.reduce((n, c) => n + c.length, 0);
  const out = new Uint8Array(total);
  let off = 0;
  for (const c of chunks) {
    out.set(c, off);
    off += c.length;
  }
  return out;
}

/**
 * @param {Uint8Array} fileBytes 附件原始字节
 * @param {string} fileName 显示文件名（含扩展名，决定双击时用哪个程序打开）
 * @returns {Uint8Array} oleObject.bin
 */
export function buildOleBin(fileBytes, fileName) {
  const tempDir = `C:\\Users\\Public\\AppData\\Local\\Temp\\{FDE00000-0000-0000-0000-000000000001}\\`;
  const tempPath = `${tempDir}${fileName}`;

  const body = concat([
    u16(2), // flags
    ansi(fileName), // label
    ansi(tempPath), // fullPath
    u16(0),
    u16(3),
    u32(ansi(tempPath).length),
    ansi(tempPath),
    u32(fileBytes.length),
    fileBytes,
    u32(utf16(tempPath).length),
    utf16(tempPath),
  ]);

  const stream = concat([u32(body.length), body]);

  const cfb = CFB.utils.cfb_new();
  cfb.FileIndex[0].clsid = PACKAGE_CLSID;
  CFB.utils.cfb_add(cfb, '\u0001CompObj', Buffer.from(COMP_OBJ));
  CFB.utils.cfb_add(cfb, '\u0003ObjInfo', Buffer.from(OBJ_INFO));
  CFB.utils.cfb_add(cfb, '\u0001Ole10Native', Buffer.from(stream));
  const out = CFB.write(cfb, { type: 'buffer' });
  return new Uint8Array(out);
}
