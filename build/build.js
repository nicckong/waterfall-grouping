/* Builds waterfall_grouping.html: src/page.html + src/core.js + the ExcelJS library, compressed into
   checksummed parts. Run with: npm run build */
const fs = require('fs'), path = require('path'), zlib = require('zlib'), vm = require('vm');
const ROOT = path.join(__dirname, '..');
const OUT = path.join(ROOT, 'waterfall_grouping.html');
const read = (...p) => fs.readFileSync(path.join(ROOT, ...p), 'utf8');

let lib = read('node_modules', 'exceljs', 'dist', 'exceljs.min.js').replace(/\/\/# sourceMappingURL=.*$/m, '');
// ExcelJS strips backslashes from number formats when reading, which can change what a format means
// (0.0\% shows a literal %, 0.0% multiplies by 100). Keep them as written.
const STRIP = String.raw`formatCode:e.attributes.formatCode.replace(/[\\](.)/g,"$1")`;
if (!lib.includes(STRIP)) throw new Error('ExcelJS number-format patch target not found: check the ExcelJS version');
lib = lib.replace(STRIP, () => 'formatCode:e.attributes.formatCode');
const core = read('src', 'core.js');
const licence = read('node_modules', 'exceljs', 'LICENSE').trim();

// Same CRC32 as the page uses (over the base64 text with whitespace removed).
function crc32(text) {
  let c, crc = 0xffffffff;
  for (let i = 0; i < text.length; i++) {
    c = (crc ^ text.charCodeAt(i)) & 0xff;
    for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
    crc = (crc >>> 8) ^ c;
  }
  return ((crc ^ 0xffffffff) >>> 0).toString(16).padStart(8, '0');
}

// Compressed and split into parts, each with its own checksum, so the file is small, survives copying,
// and a damaged copy can say which part broke.
const b64 = zlib.gzipSync(Buffer.from(lib, 'utf8'), { level: 9 }).toString('base64');
const PART = 90000, LINE = 100, n = Math.ceil(b64.length / PART);
const blocks = [`<!-- ExcelJS licence:\n${licence.replace(/--/g, '- -')}\n-->`];
for (let i = 0; i < n; i++) {
  const chunk = b64.slice(i * PART, (i + 1) * PART);
  const lines = chunk.match(new RegExp(`.{1,${LINE}}`, 'g')).join('\n');
  blocks.push(`<!-- Excel reader, part ${i + 1} of ${n} -->\n<script type="text/x-wg-part" data-part="${i + 1}" data-of="${n}" data-crc="${crc32(chunk)}">\n${lines}\n</script>`);
}

let html = read('src', 'page.html').replace('/*CORE*/', () => core).replace('/*EXCELJS*/', () => blocks.join('\n'))
  .replace('/*PARTS*/0', () => String(n));
html = html.replace('/*SIZE*/0', () => '/*SIZE*/0000000');
html = html.replace('/*SIZE*/0000000', () => '/*SIZE*/' + String(html.length).padStart(7, ' ')); // same length, so the number is exact

// Checks: every script parses, the parts unpack to exactly the library, and the file is plain ASCII.
for (const m of html.matchAll(/<script>([\s\S]*?)<\/script>/g)) new vm.Script(m[1]);
const back = [...html.matchAll(/data-part="(\d+)" data-of="\d+" data-crc="([0-9a-f]{8})">([\s\S]*?)<\/script>/g)].map((m) => {
  const t = m[3].replace(/\s+/g, ''); if (crc32(t) !== m[2]) throw new Error('crc mismatch in part ' + m[1]); return t;
}).join('');
if (zlib.gunzipSync(Buffer.from(back, 'base64')).toString('utf8') !== lib) throw new Error('round trip failed');
if (/[^\x09\x0a\x0d\x20-\x7e]/.test(html)) throw new Error('non-ASCII character in output');

fs.writeFileSync(OUT, html);
const lens = html.split('\n').map((l) => l.length);
console.log(`built ${path.relative(ROOT, OUT)}: ${(html.length / 1024).toFixed(0)} KB, ${n} library parts, longest line ${Math.max(...lens)} chars`);
