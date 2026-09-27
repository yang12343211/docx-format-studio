const path = require('path');
const fs = require('fs');
const { DOMParser, XMLSerializer } = require('@xmldom/xmldom');
global.DOMParser = DOMParser;
global.XMLSerializer = XMLSerializer;
global.JSZip = require(path.resolve(__dirname, '../js/jszip.min.js'));
const DocxCore = require(path.resolve(__dirname, '../js/docx-core.js'));

const ROOT = path.resolve(__dirname, '..');
const toBuf = p => { const b = fs.readFileSync(p); return b.buffer.slice(b.byteOffset, b.byteOffset + b.byteLength); };
const norm = s => s.replace(/(<\/?)\w+:/g, '$1');

let pass = 0, fail = 0;
function t(name, cond, extra) {
  if (cond) { pass++; console.log('  \u2713 ' + name); }
  else { fail++; console.log('  \u2717 ' + name + (extra ? '  -> ' + extra : '')); }
}

(async () => {
  const preset = DocxCore.BUILTIN_PRESETS['report-cn'];
  console.log('=== 预设：' + preset.label + ' ===');
  console.log(JSON.stringify(DocxCore.describeSpec(preset.spec), null, 2).slice(0, 1600));

  const res = await DocxCore.applyFormat(toBuf(path.join(ROOT, 'samples/target.docx')), preset.spec, {});
  const zip = await global.JSZip.loadAsync(res.data);
  const xml = norm(await zip.file('word/document.xml').async('string'));

  console.log('\n=== XML 校验 ===');
  // 页边距 twips: 上3.5cm=1985 右2.0cm=1134 下3.0cm=1701 左2.5cm=1418
  t('页边距 上1985/右1134/下1701/左1418',
    /w:top="1985"/.test(xml) && /w:right="1134"/.test(xml) && /w:bottom="1701"/.test(xml) && /w:left="1418"/.test(xml),
    (xml.match(/<w:pgMar[^>]*>/) || [''])[0]);
  t('页眉1418 / 页脚1134', /w:header="1418"/.test(xml) && /w:footer="1134"/.test(xml));

  const ps = [];
  const re = /<p>[\s\S]*?<\/p>/g; let m;
  while ((m = re.exec(xml))) ps.push(m[0]);

  const h1 = ps.find(p => /w:val="Heading1"/.test(p));
  const h2 = ps.find(p => /w:val="Heading2"/.test(p));
  const h3 = ps.find(p => /w:val="Heading3"/.test(p));
  const body = ps.find(p => /w:val="Normal"/.test(p) && /firstLineChars="200"/.test(p));

  t('一级标题 黑体 小二(36)', !!h1 && /eastAsia="黑体"/.test(h1) && /<sz w:val="36"\/>/.test(h1), h1 && h1.slice(0, 260));
  t('一级标题 居中', !!h1 && /<jc w:val="center"\/>/.test(h1), h1 && h1.slice(0, 200));
  t('一级标题 单倍行距(240)+段前/段后0', !!h1 && /w:line="240" w:lineRule="auto"/.test(h1) && !/w:(before|after)="[1-9]/.test(h1), h1 && h1.slice(0, 200));

  t('二级标题 黑体 小三(30)', !!h2 && /eastAsia="黑体"/.test(h2) && /<sz w:val="30"\/>/.test(h2), h2 && h2.slice(0, 260));
  t('二级标题 段前240(12磅)/段后160(8磅)/行距240', !!h2 && /w:before="240"/.test(h2) && /w:after="160"/.test(h2) && /w:line="240" w:lineRule="auto"/.test(h2), h2 && h2.slice(0, 260));

  t('正文 宋体 小四(24)', !!body && /eastAsia="宋体"/.test(body) && /<sz w:val="24"\/>/.test(body));
  t('正文 行距固定值20磅(400 exact)', !!body && /w:line="400" w:lineRule="exact"/.test(body), body && body.slice(0, 300));
  t('正文 首行缩进2字符', !!body && /w:firstLineChars="200"/.test(body) && /w:firstLine="480"/.test(body));
  t('正文 两端对齐', /<jc w:val="both"\/>/.test(body));

  console.log('\n---------------------------------------');
  console.log('  通过 ' + pass + ' 项，失败 ' + fail + ' 项');
  console.log('---------------------------------------');
  process.exit(fail ? 1 : 0);
})().catch(e => { console.error(e); process.exit(1); });
