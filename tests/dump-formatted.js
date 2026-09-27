const fs = require('fs'), path = require('path');
const { DOMParser, XMLSerializer } = require('@xmldom/xmldom');
global.DOMParser = DOMParser; global.XMLSerializer = XMLSerializer;
global.JSZip = require(path.resolve(__dirname, '../js/jszip.min.js'));
const Core = require(path.resolve(__dirname, '../js/docx-core.js'));

function kid(el, n) { if (!el) return null; const c = el.childNodes; for (let i = 0; i < c.length; i++) if (c[i].nodeType === 1 && c[i].localName === n) return c[i]; return null; }
function A(el, n) { return el ? el.getAttribute(n) : null; }

(async () => {
  const file = process.argv[2];
  const b = fs.readFileSync(file);
  const res = await Core.applyFormat(b.buffer.slice(b.byteOffset, b.byteOffset + b.byteLength),
    Core.BUILTIN_PRESETS['report-cn'].spec, {});
  console.log('=== 处理报告 ===');
  console.log(JSON.stringify(res.report, null, 2));

  const zip = await global.JSZip.loadAsync(res.data);
  const doc = new DOMParser().parseFromString(await zip.file('word/document.xml').async('string'), 'text/xml');
  const ps = doc.getElementsByTagName('w:p');
  const stat = { 1: 0, 2: 0, 3: 0, body: 0, blank: 0 };
  const list = [];
  for (let i = 0; i < ps.length; i++) {
    const p = ps[i]; let txt = ''; const ts = p.getElementsByTagName('w:t');
    for (let j = 0; j < ts.length; j++) txt += ts[j].textContent;
    txt = txt.trim();
    const pPr = kid(p, 'pPr'); const sp = kid(pPr, 'spacing'); const jc = kid(pPr, 'jc'); const ind = kid(pPr, 'ind');
    const sid = A(kid(pPr, 'pStyle'), 'w:val'); const ol = A(kid(pPr, 'outlineLvl'), 'w:val');
    const r = p.getElementsByTagName('w:r')[0]; const rPr = r ? kid(r, 'rPr') : null; const rf = kid(rPr, 'rFonts');
    const sz = A(kid(rPr, 'sz'), 'w:val');
    const lvl = ol == null ? null : (+ol + 1);
    if (!txt) stat.blank++; else if (lvl) stat[lvl] = (stat[lvl] || 0) + 1; else stat.body++;
    list.push({ i, txt, sid, lvl, jc: A(jc, 'w:val'), font: A(rf, 'w:eastAsia'), sz, line: A(sp, 'w:line'), lineRule: A(sp, 'w:lineRule'), ind: ind ? (A(ind, 'w:firstLineChars') || A(ind, 'w:firstLine')) : null });
  }
  console.log('\n=== 层级统计 ===', JSON.stringify(stat));
  console.log('\n=== 所有非空段落 ===');
  list.filter(x => x.txt).forEach(x => {
    console.log('[' + x.i + '] L' + (x.lvl || '正文') + ' st=' + (x.sid || '-') + ' jc=' + (x.jc || '-') +
      ' ea=' + (x.font || '-') + ' sz=' + (x.sz || '-') + ' line=' + (x.line || '-') + '/' + (x.lineRule || '-') +
      ' ind=' + (x.ind || '-') + ' | ' + x.txt.slice(0, 40));
  });

  // 导出，便于在 Word 里直接查看
  const outPath = path.resolve(__dirname, '..', 'samples', 'out', 'verify-user-formatted.docx');
  fs.mkdirSync(path.dirname(outPath), { recursive: true });
  fs.writeFileSync(outPath, Buffer.from(res.data));
  console.log('\n产物已写入: ' + outPath);
})().catch(e => { console.error(e); process.exit(1); });
