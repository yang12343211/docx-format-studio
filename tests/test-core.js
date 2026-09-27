/* 核心引擎 Node 端测试 —— 验证「模板规范提取」与「格式套用」是否正确 */
const path = require('path');
const fs = require('fs');

const { DOMParser, XMLSerializer } = require('@xmldom/xmldom');
global.DOMParser = DOMParser;
global.XMLSerializer = XMLSerializer;
global.JSZip = require(path.resolve(__dirname, '../js/jszip.min.js'));

const DocxCore = require(path.resolve(__dirname, '../js/docx-core.js'));

const SAMPLES = path.resolve(__dirname, '../samples');
const OUT = path.resolve(__dirname, '../samples/out');
if (!fs.existsSync(OUT)) fs.mkdirSync(OUT, { recursive: true });

let pass = 0, fail = 0;
function check(name, cond, extra) {
  if (cond) { pass++; console.log('  \u2713 ' + name); }
  else { fail++; console.log('  \u2717 ' + name + (extra ? '  -> ' + extra : '')); }
}

function toBuf(file) {
  const b = fs.readFileSync(file);
  return b.buffer.slice(b.byteOffset, b.byteOffset + b.byteLength);
}

(async function main() {
  console.log('\n=== 1. 分析模板 template.docx ===');
  const spec = await DocxCore.analyzeTemplate(toBuf(path.join(SAMPLES, 'template.docx')),
    { fileName: 'template.docx' });

  console.log('  页面:', JSON.stringify(spec.page), JSON.stringify(spec.margin));
  console.log('  识别到的样式:', spec.foundNames.join(' | ') || '(无)');

  check('页边距上 = 2.54cm(1440)', spec.margin && spec.margin.top === 1440, JSON.stringify(spec.margin));
  check('页边距左 = 3.17cm(1800)', spec.margin && spec.margin.left === 1800);
  check('纸张 A4', spec.page && spec.page.w === 11906 && spec.page.h === 16838);
  check('识别出标题 1', !!spec.headings[1], 'headings=' + Object.keys(spec.headings));
  check('标题1 字体=黑体', spec.headings[1] && spec.headings[1].rPr.font.eastAsia === '黑体',
    spec.headings[1] && JSON.stringify(spec.headings[1].rPr.font));
  check('标题1 字号=32(三号)', spec.headings[1] && spec.headings[1].rPr.size === 32);
  check('标题1 居中', spec.headings[1] && spec.headings[1].pPr.align === 'center');
  check('标题1 继承到 1.5 倍行距', spec.headings[1] && spec.headings[1].pPr.spacing &&
    spec.headings[1].pPr.spacing.line === 360, JSON.stringify(spec.headings[1] && spec.headings[1].pPr.spacing));
  check('识别出标题 2（中文名 "标题 2"）', !!(spec.headings[2] && spec.headings[2].rPr.size === 28),
    JSON.stringify(spec.headings[2] && spec.headings[2].rPr));
  check('识别出标题 3（styleId 为 "3"）', !!(spec.headings[3] && spec.headings[3].rPr.size === 24),
    JSON.stringify(spec.headings[3] && spec.headings[3].rPr));
  check('标题3 加粗', spec.headings[3] && spec.headings[3].rPr.bold === true);
  check('正文 字体=宋体', spec.normal && spec.normal.rPr.font.eastAsia === '宋体');
  check('正文 字号=24(小四)', spec.normal && spec.normal.rPr.size === 24);
  check('正文 首行缩进 2 字符', spec.normal && spec.normal.pPr.ind && spec.normal.pPr.ind.firstLineChars === 200);
  check('识别出题注', !!spec.caption && spec.caption.rPr.size === 18);

  console.log('\n=== 2. 套用到 target.docx ===');
  const res = await DocxCore.applyFormat(toBuf(path.join(SAMPLES, 'target.docx')), spec, {
    applyPage: true, applyHeading: true, applyBody: true, applyCaption: true,
    skipTableIndent: true, heuristic: true, clearManual: true
  });

  const r = res.report;
  console.log('  报告:', JSON.stringify({
    paragraphs: r.paragraphs, headings: r.headings, body: r.body,
    caption: r.caption, tableParagraphs: r.tableParagraphs,
    heuristicHits: r.heuristicHits, skippedFields: r.skippedFields,
    pageApplied: r.pageApplied
  }));
  if (r.warnings.length) console.log('  warnings:', r.warnings);

  check('总段落数 = 14（8 正文段 + 表格 6 单元格段）', r.paragraphs === 14, 'got ' + r.paragraphs);
  check('识别出 1 个一级标题', r.headings[1] === 1, 'got ' + r.headings[1]);
  check('识别出 2 个二级标题（含启发式）', r.headings[2] === 2, 'got ' + r.headings[2]);
  check('启发式命中 1 处（"2.1 系统需求分析"）', r.heuristicHits === 1, 'got ' + r.heuristicHits);
  check('题注 1 处', r.caption === 1, 'got ' + r.caption);
  check('表格内段落被统计', r.tableParagraphs === 6, 'got ' + r.tableParagraphs);
  check('目录段落被跳过', r.skippedFields === 1, 'got ' + r.skippedFields);
  check('页面设置已应用', r.pageApplied === true);
  check('无警告', r.warnings.length === 0, r.warnings.join('; '));
  check('输出为有效 zip 大小 > 2KB', res.data.byteLength > 2048, res.data.byteLength);

  /* 把结果写盘，再用 JSZip 打开检查 XML */
  const outPath = path.join(OUT, 'target-formatted.docx');
  fs.writeFileSync(outPath, Buffer.from(res.data));

  const zip = await global.JSZip.loadAsync(res.data);
  const xml = await zip.file('word/document.xml').async('string');
  fs.writeFileSync(path.join(OUT, 'document.xml'), xml);

  console.log('\n=== 3. 检查输出 XML ===');
  check('XML 声明存在', /^<\?xml[^>]*\?>/.test(xml.trim()));
  check('页边距写入 1440/1800', /w:top="1440"/.test(xml) && /w:left="1800"/.test(xml), '');
  check('docGrid 写入', /w:docGrid/.test(xml));
  check('标题1 黑体已写入', /w:eastAsia="黑体"/.test(xml));
  check('标题1 三号(32) 已写入', /<w:sz w:val="32"\/>/.test(xml) || /w:sz w:val="32"/.test(xml));
  check('正文 小四(24) 已写入', /w:sz w:val="24"/.test(xml));
  check('首行缩进 200 已写入', /w:firstLineChars="200"/.test(xml));
  check('标题1 居中已写入', /<w:jc w:val="center"\/>/.test(xml) || /w:jc w:val="center"/.test(xml));
  check('旧的手动红色 FF0000 已被清除', !/FF0000/.test(xml));
  check('旧的微软雅黑已被清除', !/微软雅黑/.test(xml));
  check('旧纸张以外的旧页边距 720 已清除', !/w:top="720"/.test(xml));
  check('outlineLvl 已写入（导航可识别）', /w:outlineLvl/.test(xml));
  check('文本内容完整保留', /第一章/.test(xml) && /示波器/.test(xml) && /系统需求分析/.test(xml));
  check('表格结构保留', /<w:tbl>/.test(xml));

  /* ---- 更精细的结构检查 ---- */
  console.log('\n=== 3b. 逐段结构检查 ===');
  const WNS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
  const d = new DOMParser().parseFromString(xml, 'text/xml');
  const ps = d.getElementsByTagNameNS(WNS, 'p');
  function attrOf(el, name) {
    if (!el) return null;
    return el.getAttributeNS(WNS, name) || el.getAttribute('w:' + name);
  }
  function find(el, name) { return el ? el.getElementsByTagNameNS(WNS, name)[0] : null; }

  const p0Run = find(ps[0], 'r');
  const p0Rpr = find(p0Run, 'rPr');
  check('标题1 的 run 实际使用黑体',
    attrOf(find(p0Rpr, 'rFonts'), 'eastAsia') === '黑体',
    attrOf(find(p0Rpr, 'rFonts'), 'eastAsia'));
  check('标题1 的 run 实际字号 32',
    attrOf(find(p0Rpr, 'sz'), 'val') === '32', attrOf(find(p0Rpr, 'sz'), 'val'));
  check('标题1 的 run 已去掉加粗（模板要求不加粗）',
    !find(p0Rpr, 'b') || attrOf(find(p0Rpr, 'b'), 'val') === '0',
    find(p0Rpr, 'b') ? new XMLSerializer().serializeToString(find(p0Rpr, 'b')) : 'none');

  const p1RunRpr = find(find(ps[1], 'r'), 'rPr');
  check('正文的 run 实际使用宋体', attrOf(find(p1RunRpr, 'rFonts'), 'eastAsia') === '宋体');
  check('正文的 run 实际字号 24', attrOf(find(p1RunRpr, 'sz'), 'val') === '24');

  const tblXml = (xml.match(/<w:tbl>[\s\S]*?<\/w:tbl>/) || [''])[0];
  check('表格内段落不加首行缩进', tblXml.length > 0 && !/w:firstLineChars/.test(tblXml));
  check('表格内段落仍套用宋体小四', /w:eastAsia="宋体"/.test(tblXml) && /w:sz w:val="24"/.test(tblXml));

  /* 文档顺序：0-3 正文/标题，4-9 表格单元格，10 题注，11 启发式标题，12 正文，13 目录 */
  const captionP = ps[10];
  check('题注段落居中且不缩进',
    /<w:jc w:val="center"\/>/.test(new XMLSerializer().serializeToString(captionP)) &&
    !/w:firstLineChars="200"/.test(new XMLSerializer().serializeToString(captionP)),
    new XMLSerializer().serializeToString(captionP).slice(0, 260));

  const heuristicP = ps[11];
  const hXml = new XMLSerializer().serializeToString(heuristicP);
  check('启发式识别出的标题带 outlineLvl=1', /<w:outlineLvl w:val="1"\/>/.test(hXml), hXml.slice(0, 260));

  /* ---- 内置预设也能工作 ---- */
  console.log('\n=== 3c. 内置预设 ===');
  const preset = DocxCore.BUILTIN_PRESETS['thesis-cn'].spec;
  const res2 = await DocxCore.applyFormat(toBuf(path.join(SAMPLES, 'target.docx')), preset, {});
  check('内置预设可套用且无警告', res2.report.warnings.length === 0,
    res2.report.warnings.join('; '));
  const zip2 = await global.JSZip.loadAsync(res2.data);
  const xml2 = await zip2.file('word/document.xml').async('string');
  check('内置预设写入页边距 1440/1800', /w:left="1800"/.test(xml2));

  /* 二次校验：把结果文档当作"目标文档"再解析一次，确认 Word 结构未破坏 */
  console.log('\n=== 4. 二次解析（结构完整性） ===');
  const spec2 = await DocxCore.analyzeTemplate(res.data, { fileName: 'roundtrip' });
  check('结果文档可被重新解析', !!spec2);
  check('二次解析读到的页边距一致', spec2.margin && spec2.margin.top === 1440);

  /* 段落实样输出，人工核对 */
  console.log('\n=== 5. 抽样输出（前 2 个段落） ===');
  for (let i = 0; i < Math.min(2, ps.length); i++) {
    const x = new XMLSerializer().serializeToString(ps[i]);
    console.log('  [' + i + '] ' + x.slice(0, 460).replace(/\r?\n/g, '') + (x.length > 460 ? ' …' : ''));
  }

  console.log('\n---------------------------------------');
  console.log('  通过 ' + pass + ' 项，失败 ' + fail + ' 项');
  console.log('  产物: ' + outPath);
  console.log('---------------------------------------\n');
  process.exit(fail ? 1 : 0);
})().catch(function (e) {
  console.error('测试异常:', e && e.stack || e);
  process.exit(2);
});
