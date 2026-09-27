/**
 * 用「课程设计报告规范」预设处理一个真实文档。
 * 用法：node tests/format-one.js <输入.docx> <输出.docx> [--no-chapter]
 */
const fs = require('fs'), path = require('path');
const { DOMParser, XMLSerializer } = require('@xmldom/xmldom');
global.DOMParser = DOMParser; global.XMLSerializer = XMLSerializer;
global.JSZip = require(path.resolve(__dirname, '../js/jszip.min.js'));
const Core = require(path.resolve(__dirname, '../js/docx-core.js'));

function kid(el, n) { if (!el) return null; const c = el.childNodes; for (let i = 0; i < c.length; i++) if (c[i].nodeType === 1 && c[i].localName === n) return c[i]; return null; }
function A(el, n) { return el ? el.getAttribute(n) : null; }

(async () => {
  const input = process.argv[2];
  const output = process.argv[3];
  if (!input || !output) { console.error('用法: node tests/format-one.js <输入.docx> <输出.docx> [--no-chapter]'); process.exit(2); }
  const noChapter = process.argv.indexOf('--no-chapter') >= 0;

  const b = fs.readFileSync(input);
  const opts = noChapter
    ? { chapter: { pageBreakBefore: false, blankLinesBefore: 0, blankLinesAfter: 0 } }
    : {};
  const res = await Core.applyFormat(b.buffer.slice(b.byteOffset, b.byteOffset + b.byteLength),
    Core.BUILTIN_PRESETS['report-cn'].spec, opts);

  const r = res.report;
  console.log('=== 处理报告 ===');
  console.log('段落 ' + r.paragraphs + ' | 一级 ' + r.headings[1] + ' 二级 ' + r.headings[2] +
    ' 三级 ' + r.headings[3] + ' 题注 ' + r.caption + ' 正文 ' + r.body);
  console.log('清掉水平线 ' + r.removedRules + ' | 章分页 ' + r.chapterBreaks + ' | 章前后空行 ' + r.chapterBlanks);
  if (r.warnings.length) console.log('警告: ' + r.warnings.join('; '));

  const zip = await global.JSZip.loadAsync(res.data);
  const doc = new DOMParser().parseFromString(await zip.file('word/document.xml').async('string'), 'text/xml');
  const ps = doc.getElementsByTagName('w:p');
  console.log('\n=== 章标题前后结构（前3个）===');
  let shown = 0;
  for (let i = 0; i < ps.length && shown < 3; i++) {
    const p = ps[i]; let txt = ''; const ts = p.getElementsByTagName('w:t');
    for (let j = 0; j < ts.length; j++) txt += ts[j].textContent;
    if (!/^第\d+章/.test(txt.trim())) continue;
    shown++;
    [i - 1, i, i + 1].forEach(function (k) {
      const q = ps[k]; if (!q) return;
      let s = ''; const t2 = q.getElementsByTagName('w:t');
      for (let j = 0; j < t2.length; j++) s += t2[j].textContent;
      const pPr = kid(q, 'pPr'); const r0 = q.getElementsByTagName('w:r')[0];
      const rf = r0 ? kid(kid(r0, 'rPr'), 'rFonts') : null;
      const sz = r0 ? A(kid(kid(r0, 'rPr'), 'sz'), 'w:val') : null;
      console.log('  [' + k + '] ' + (k === i ? '章标题' : (s.trim() ? '正文' : '空行')) +
        ' pStyle=' + (A(kid(pPr, 'pStyle'), 'w:val') || '-') +
        ' pageBreak=' + (kid(pPr, 'pageBreakBefore') ? 'Y' : '-') +
        ' ea=' + (A(rf, 'w:eastAsia') || '-') + ' sz=' + (sz || '-') +
        ' | ' + s.trim().slice(0, 26));
    });
  }

  fs.writeFileSync(output, Buffer.from(res.data));
  console.log('\n产物已写入: ' + output);
})().catch(e => { console.error(e); process.exit(1); });
