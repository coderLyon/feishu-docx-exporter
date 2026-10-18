/**
 * 代码块轻量着色 + 背景对比度保护。
 *
 * 为什么需要「轻量」：飞书的块数据里只有纯文本 + language 字段，**没有**逐 token 的颜色信息
 * （颜色只在 canvas 渲染层）。所以这里按语言做规则着色，目标是「一眼能分清字符串/注释/关键字」，
 * 而不是像素级还原在线主题。
 *
 * 对比度保护（重要）：如果背景是深色、文字还是深色，打印出来就是一片黑。
 * 这里按背景亮度自动切换深/浅两套配色，保证任何背景都能看清。
 */

/** 浅色底用的配色（GitHub Light 风格，克制、不刺眼） */
const LIGHT = {
  plain: '24292E',
  comment: '6A737D',
  string: '032F62',
  number: '005CC5',
  keyword: 'D73A49',
  key: '22863A',
};

/** 深色底用的配色 */
const DARK = {
  plain: 'D4D4D4',
  comment: '6A9955',
  string: 'CE9178',
  number: 'B5CEA8',
  keyword: '569CD6',
  key: '9CDCFE',
};

/** #RRGGBB 的相对亮度（0=黑，1=白） */
export function luminance(hex) {
  const h = String(hex ?? '').replace('#', '');
  if (h.length !== 6) return 1;
  const rgb = [0, 2, 4].map((i) => parseInt(h.slice(i, i + 2), 16) / 255);
  const lin = rgb.map((c) => (c <= 0.03928 ? c / 12.92 : ((c + 0.055) / 1.055) ** 2.4));
  return 0.2126 * lin[0] + 0.7152 * lin[1] + 0.0722 * lin[2];
}

/** 背景是否属于「深色」，深色底必须换浅色字，否则打印不可读 */
export function isDarkBackground(hex) {
  return luminance(hex) < 0.5;
}

const KEYWORDS = {
  json: new Set(['true', 'false', 'null']),
  sql: new Set(
    ('select from where insert into update delete join left right inner outer on as and or not null is in ' +
      'group by order having limit offset create table alter drop index primary key foreign references ' +
      'varchar char int bigint decimal date datetime timestamp boolean default constraint unique values set')
      .split(' '),
  ),
  js: new Set(
    ('const let var function return if else for while do switch case break continue new class extends super ' +
      'this try catch finally throw async await import export from default typeof instanceof delete void yield')
      .split(' '),
  ),
  python: new Set(
    ('def class return if elif else for while break continue import from as try except finally raise with ' +
      'lambda yield global nonlocal pass assert del in is not and or None True False self')
      .split(' '),
  ),
  java: new Set(
    ('public private protected class interface extends implements static final void int long double float ' +
      'boolean char String if else for while return new try catch finally throw import package')
      .split(' '),
  ),
  shell: new Set(['if', 'then', 'else', 'fi', 'for', 'do', 'done', 'while', 'case', 'esac', 'export', 'echo', 'cd']),
};

/** 语言名 → 关键字表 key */
function langFamily(language) {
  const l = String(language ?? '').toLowerCase();
  if (l.includes('json')) return 'json';
  if (l.includes('sql')) return 'sql';
  if (l.includes('python')) return 'python';
  if (l.includes('java') && !l.includes('script')) return 'java';
  if (l.includes('shell') || l.includes('bash') || l.includes('sh')) return 'shell';
  if (l.includes('javascript') || l.includes('typescript') || l === 'js' || l === 'ts') return 'js';
  return null;
}

// 注释 | 字符串 | 数字 | 标识符
const TOKEN_RX =
  /(\/\/[^\n]*|#[^\n]*|--[^\n]*)|("(?:[^"\\]|\\.)*"|'(?:[^'\\]|\\.)*'|`(?:[^`\\]|\\.)*`)|(\b\d+(?:\.\d+)?\b)|([A-Za-z_$][\w$]*)/g;

/**
 * 把一行代码切成带颜色的片段。
 *
 * @param {string} line
 * @param {string} language 飞书的 language 字段（如 "JSON"、"SQL"）
 * @param {string} background 代码块底色（#RRGGBB），用于决定深/浅配色
 * @returns {Array<{text: string, color: string}>}
 */
export function highlightLine(line, language, background = 'F5F6F7') {
  const palette = isDarkBackground(background) ? DARK : LIGHT;
  const family = langFamily(language);
  const keywords = family ? KEYWORDS[family] : null;
  const text = String(line ?? '');
  if (!text) return [{ text: '', color: palette.plain }];

  const out = [];
  let last = 0;
  TOKEN_RX.lastIndex = 0;
  let m;
  while ((m = TOKEN_RX.exec(text))) {
    if (m.index > last) out.push({ text: text.slice(last, m.index), color: palette.plain });
    const [whole, comment, str, num, ident] = m;
    let color = palette.plain;
    if (comment) color = palette.comment;
    else if (str) {
      // JSON 里 "key": 的键单独着色，读起来更像结构化数据
      const rest = text.slice(m.index + whole.length);
      color = family === 'json' && /^\s*:/.test(rest) ? palette.key : palette.string;
    } else if (num) color = palette.number;
    else if (ident && keywords?.has(ident.toLowerCase())) color = palette.keyword;
    out.push({ text: whole, color });
    last = m.index + whole.length;
  }
  if (last < text.length) out.push({ text: text.slice(last), color: palette.plain });
  return out;
}


/**
 * 把背景色「提亮」到可打印的浅色，保留色相。
 *
 * 场景：callout 的背景色由作者选择，可能是深蓝/深红；如果照搬，Word 里就是深底 +
 * 深色正文，打印出来一片黑。这里按亮度向白色混合，直到足够浅为止。
 *
 * @param {string} hex 原始背景（#RRGGBB 或 RRGGBB）
 * @returns {string} 安全的浅色（RRGGBB，无 #）
 */
export function safeBackground(hex) {
  const h = String(hex ?? '').replace('#', '');
  if (!/^[0-9a-fA-F]{6}$/.test(h)) return 'F5F6F7';
  const rgb = [0, 2, 4].map((i) => parseInt(h.slice(i, i + 2), 16));
  let cur = rgb;
  // 每次向白色靠 20%，最多 6 次；亮度达标即停
  for (let i = 0; i < 6; i += 1) {
    const hexNow = cur.map((c) => c.toString(16).padStart(2, '0')).join('');
    if (!isDarkBackground(hexNow)) return hexNow.toUpperCase();
    cur = cur.map((c) => Math.round(c + (255 - c) * 0.35));
  }
  return cur.map((c) => c.toString(16).padStart(2, '0')).join('').toUpperCase();
}
