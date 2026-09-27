/**
 * 标题层级识别 / 题注识别 / 空标题样式段落 的回归测试
 *
 * 用例背景：
 *   1) 有文档把「带标题样式的空段落」当空行用，照着样式套会把空行变成小二黑体居中。
 *   2) 旧规则按顺序遍历正则，浅规则 ^数字+[.、] 会先把 "1.1.1" 抢走 → 三级标题被套成二级。
 *   3) 样式里没写「题注」时，「表4-1 / 图2-1」这类题注会被套成正文格式。
 *
 * 运行：node tests/test-headings.js
 */
const path = require('path');
const fs = require('fs');
const { DOMParser, XMLSerializer } = require('@xmldom/xmldom');

global.DOMParser = DOMParser;
global.XMLSerializer = XMLSerializer;
global.JSZip = require(path.resolve(__dirname, '../js/jszip.min.js'));

const DocxCore = require(path.resolve(__dirname, '../js/docx-core.js'));

let pass = 0, fail = 0;
function check(name, cond, extra) {
  if (cond) { pass++; console.log('  \u2713 ' + name); }
  else { fail++; console.log('  \u2717 ' + name + (extra !== undefined ? '  -> ' + String(extra).slice(0, 300) : '')); }
}

/* ============================ 1. 纯函数：层级判定 ============================ */

function testHeuristic() {
  console.log('=== 1. 标题层级判定 ===');
  const H = DocxCore._internal.heuristicLevel;
  const cases = [
    /* [文本, 期望层级] —— 0 = 不是标题 */
    ['第1章 绪论', 1],
    ['第1章  绪   论', 1],
    ['第一章 绪论', 1],
    ['第12章 系统设计', 1],
    ['第1节 概述', 2],

    ['1.1 课题背景与意义', 2],
    ['1.1.1 系统功能模块', 3],          /* ← 旧版会误判成 2 */
    ['1.1.1.1 更深一层', 3],            /* 四级收敛到三级 */
    ['3.2.1 数据库表设计', 3],
    ['10.1.2 硬件连接', 3],
    ['5.3.7 程序段7—位置感应信号生成', 3],
    ['2.3 有关公式、查表和插图', 2],

    ['一、项目背景', 2],
    ['（一）总体方案', 3],
    ['（1）硬件设计', 3],
    ['1. 工艺对象分析', 3],
    ['1、概述说明', 3],

    ['参考文献', 1],
    ['摘要', 1],
    ['目录', 1],
    ['致谢', 1],
    ['附录A', 1],

    /* 不能误判的 */
    ['2023.5 版本说明', 0],             /* 四位数年份不是章节号 */
    ['图3 所示为主电路', 0],            /* 含句末说明，弱规则要拦掉 */
    ['本课题以1000MW超超临界机组电厂化学水处理系统为被控对象，基于实验室', 0],
    ['普通的一段正文，讲了很多东西。', 0]
  ];
  cases.forEach(function (c) {
    const got = H(c[0]);
    check('「' + c[0].slice(0, 26) + '」→ ' + c[1] + ' 级', got === c[1], '实际 ' + got);
  });
}

/* ============================ 2. 端到端：构造文档 ============================ */

const { buildDocx, DOC_ITEMS, PLAIN_ITEMS, RULE_ITEMS, CHAPTER_ITEMS, CHAPTER_FIRST_ITEMS, PRE_BLANKED_ITEMS, GRID_ITEMS, GRID_SECTPR, NOGRID_SECTPR, BARE_STYLES_XML } = require('./lib/build-docx.js');

/* 「课设模板规范」预设现在自带「章版式」（另起一页 + 前后各空一行）。
   2~7 节只关心标题识别、格式写入、缩进继承、水平线清理这些事，
   章版式会插分页和空行、把段落下标顶歪，所以这些用例统一把它关掉；
   章版式本身由第 8 节单独覆盖。 */
const NO_CHAPTER = { chapter: { pageBreakBefore: false, blankLinesBefore: 0, blankLinesAfter: 0 } };

function getParas(xml) {
  const doc = new DOMParser().parseFromString(xml, 'text/xml');
  const out = [];
  const ps = doc.getElementsByTagName('w:p');
  for (let i = 0; i < ps.length; i++) {
    const p = ps[i];
    let txt = '';
    const ts = p.getElementsByTagName('w:t');
    for (let j = 0; j < ts.length; j++) txt += (ts[j].textContent || '');
    out.push({ el: p, xml: new XMLSerializer().serializeToString(p), text: txt });
  }
  return out;
}

async function testEndToEnd() {
  console.log('\n=== 2. 端到端：空标题样式段 / 手打编号标题 / 无样式题注 ===');
  const buf = await buildDocx(DOC_ITEMS);
  const spec = DocxCore.BUILTIN_PRESETS['report-cn'].spec;
  const res = await DocxCore.applyFormat(buf.buffer.slice(buf.byteOffset, buf.byteOffset + buf.byteLength),
    spec, NO_CHAPTER);
  const r = res.report;

  console.log('  报告: ' + JSON.stringify(r));

  /* ---- 统计 ---- */
  check('一级标题 = 2（第1章、第2章；空段不算）', r.headings[1] === 2, r.headings[1]);
  check('二级标题 = 2（1.1、2.1）', r.headings[2] === 2, r.headings[2]);
  check('三级标题 = 2（1.1.1、1.1.2）', r.headings[3] === 2, r.headings[3]);
  check('题注 = 2（表2-1、图2-1）', r.caption === 2, r.caption);
  check('识别到 2 个空段落', r.blankParagraphs === 2, r.blankParagraphs);
  check('其中 2 处空段的标题样式被摘掉', r.demotedHeadings === 2, r.demotedHeadings);
  check('无警告', r.warnings.length === 0, r.warnings.join('; '));

  const zip = await global.JSZip.loadAsync(res.data);
  const ps = getParas(await zip.file('word/document.xml').async('string'));

  /* ---- 空段：不再是标题 ---- */
  /* 空段原来挂着标题样式，现在必须换成正文样式（styleId "1" = Normal）。
     只要不指向标题样式即可——否则 Word 导航窗格 / 更新目录里会冒出空条目。 */
  check('空段 0 不再是标题样式', !/w:pStyle w:val="[2345]"\/>/.test(ps[0].xml), ps[0].xml.slice(0, 200));
  check('空段 0 套的是正文小四(24)', /<w:sz w:val="24"\/>/.test(ps[0].xml), ps[0].xml.slice(0, 240));
  check('空段 0 两端对齐且不居中', /<w:jc w:val="both"\/>/.test(ps[0].xml) && !/<w:jc w:val="center"\/>/.test(ps[0].xml));
  check('空段 0 没有首行缩进',
    !/<w:ind/.test(ps[0].xml) || /<w:ind\b[^>]*w:firstLineChars="0"/.test(ps[0].xml), ps[0].xml.slice(0, 240));
  check('空段 0 不再带大纲级别', !/<w:outlineLvl/.test(ps[0].xml));
  check('空段 2 同样不再是标题样式', !/w:pStyle w:val="[2345]"\/>/.test(ps[2].xml));

  /* ---- 一级标题：小二黑体居中 ---- */
  check('第1章 → 小二(36) 黑体居中',
    /<w:sz w:val="36"\/>/.test(ps[1].xml) && /w:eastAsia="黑体"/.test(ps[1].xml) && /<w:jc w:val="center"\/>/.test(ps[1].xml),
    ps[1].xml.slice(0, 260));
  check('第1章 → outlineLvl 0',
    /<w:outlineLvl w:val="0"\/>/.test(ps[1].xml), ps[1].xml.slice(0, 260));

  /* ---- 二级标题：小三黑体 ---- */
  check('1.1 → 小三(30) 黑体，段前12磅/段后8磅',
    /<w:sz w:val="30"\/>/.test(ps[3].xml) && /w:before="240"/.test(ps[3].xml) && /w:after="160"/.test(ps[3].xml),
    ps[3].xml.slice(0, 260));

  /* ---- 三级标题：四号黑体 ---- */
  check('1.1.1 → 四号(28) 黑体，段前8磅/段后6磅',
    /<w:sz w:val="28"\/>/.test(ps[4].xml) && /w:before="160"/.test(ps[4].xml) && /w:after="120"/.test(ps[4].xml),
    ps[4].xml.slice(0, 260));

  /* ---- 手打编号的三级标题也必须落在三级，而不是二级 ---- */
  check('「1.1.2 手打的三级标题」→ 四号(28) 而不是小三(30)',
    /<w:sz w:val="28"\/>/.test(ps[5].xml) && !/<w:sz w:val="30"\/>/.test(ps[5].xml),
    ps[5].xml.slice(0, 260));
  check('「1.1.2 …」→ outlineLvl 2',
    /<w:outlineLvl w:val="2"\/>/.test(ps[5].xml), ps[5].xml.slice(0, 260));

  /* ---- 手打的章标题 ---- */
  check('「第2章 系统设计」→ 小二(36) 居中',
    /<w:sz w:val="36"\/>/.test(ps[6].xml) && /<w:jc w:val="center"\/>/.test(ps[6].xml));

  /* ---- 无样式题注 ---- */
  check('「表2-1 …」→ 题注 五号(21) 居中',
    /<w:sz w:val="21"\/>/.test(ps[8].xml) && /<w:jc w:val="center"\/>/.test(ps[8].xml),
    ps[8].xml.slice(0, 260));
  check('「图2-1 …」→ 题注 五号(21) 居中',
    /<w:sz w:val="21"\/>/.test(ps[9].xml) && /<w:jc w:val="center"\/>/.test(ps[9].xml),
    ps[9].xml.slice(0, 260));

  /* ---- 正文 ---- */
  check('正文 → 小四(24) + 首行缩进 2 字符',
    /<w:sz w:val="24"\/>/.test(ps[10].xml) && /firstLineChars="200"/.test(ps[10].xml),
    ps[10].xml.slice(0, 260));

  /* ---- 结果仍可被解析 ---- */
  const again = await DocxCore.analyzeTemplate(res.data, { fileName: 'roundtrip' });
  check('产物可被重新解析', !!again);
}

/* ============================ 3. 样式整体错位的文档 ============================ */

/**
 * 真实案例：一份用 AI 生成的中期报告，作者把「第1章」套成 heading 2、
 * 「1.1」套成 heading 3、「1.1.1」套成 heading 4 —— 样式整体错位一级。
 * 引擎若纯按样式给级别，输出的标题就会整体小一号（小二变小三、小三变四号）。
 */
const SHIFTED_ITEMS = [
  { text: '第1章 绪论', style: '3' },        /* 0  heading 2 样式，但文字是章标题 */
  { text: '1.1 课题背景与意义', style: '4' }, /* 1  heading 3 样式 */
  { text: '1.1.1 课题背景', style: '5' },     /* 2  heading 4 样式 */
  { text: '一、项目背景', style: '2' },       /* 3  heading 1 样式 + 弱编号，应保持一级 */
  { text: '正文段落，这里是普通的说明文字，讲清楚了系统要做什么。' }  /* 4 */
];

async function testShiftedStyles() {
  console.log('\n=== 3. 标题样式整体错位一级的文档（编号规律应纠正它） ===');
  const buf = await buildDocx(SHIFTED_ITEMS);
  const res = await DocxCore.applyFormat(buf.buffer.slice(buf.byteOffset, buf.byteOffset + buf.byteLength),
    DocxCore.BUILTIN_PRESETS['report-cn'].spec, NO_CHAPTER);
  const r = res.report;
  console.log('  报告: ' + JSON.stringify(r));

  check('「第1章」被纠正为一级（而不是 heading 2 的二级）', r.headings[1] === 2, JSON.stringify(r.headings));
  check('「1.1」被纠正为二级', r.headings[2] === 1, JSON.stringify(r.headings));
  check('「1.1.1」被纠正为三级', r.headings[3] === 1, JSON.stringify(r.headings));
  check('共纠正 3 处层级（一、那条不该被改）', r.levelFixes === 3, r.levelFixes);

  const zip = await global.JSZip.loadAsync(res.data);
  const ps = getParas(await zip.file('word/document.xml').async('string'));

  check('第1章 → 小二(36) 居中 + outlineLvl 0',
    /<w:sz w:val="36"\/>/.test(ps[0].xml) && /<w:jc w:val="center"\/>/.test(ps[0].xml) &&
    /<w:outlineLvl w:val="0"\/>/.test(ps[0].xml), ps[0].xml.slice(0, 260));
  check('1.1 → 小三(30) + outlineLvl 1',
    /<w:sz w:val="30"\/>/.test(ps[1].xml) && /<w:outlineLvl w:val="1"\/>/.test(ps[1].xml),
    ps[1].xml.slice(0, 260));
  check('1.1.1 → 四号(28) + outlineLvl 2',
    /<w:sz w:val="28"\/>/.test(ps[2].xml) && /<w:outlineLvl w:val="2"\/>/.test(ps[2].xml),
    ps[2].xml.slice(0, 260));
  check('「一、项目背景」（弱编号）不被压低，仍是一级 小二(36) 居中',
    /<w:sz w:val="36"\/>/.test(ps[3].xml) && /<w:jc w:val="center"\/>/.test(ps[3].xml),
    ps[3].xml.slice(0, 260));

  /* ---- pStyle 必须跟着层级走（否则 Word 样式库 / 样式检查器显示错的名字） ---- */
  check('第1章 的 pStyle 指向「标题 1」（styleId 2）', /w:pStyle w:val="2"\/>/.test(ps[0].xml), ps[0].xml.slice(0, 120));
  check('1.1 的 pStyle 指向「标题 2」（styleId 3）', /w:pStyle w:val="3"\/>/.test(ps[1].xml), ps[1].xml.slice(0, 120));
  check('1.1.1 的 pStyle 指向「标题 3」（styleId 4）', /w:pStyle w:val="4"\/>/.test(ps[2].xml), ps[2].xml.slice(0, 120));
  check('「一、项目背景」的 pStyle 也回到「标题 1」', /w:pStyle w:val="2"\/>/.test(ps[3].xml), ps[3].xml.slice(0, 120));

  /* ---- styles.xml 里的样式定义也要被重写成 spec ---- */
  const sty = await zip.file('word/styles.xml').async('string');
  const h1 = (sty.match(/<w:style [^>]*w:styleId="2"[\s\S]*?<\/w:style>/) || [''])[0];
  const h2 = (sty.match(/<w:style [^>]*w:styleId="3"[\s\S]*?<\/w:style>/) || [''])[0];
  check('样式「标题 1」定义已重写为 小二(36) + outlineLvl 0',
    /<w:sz w:val="36"\/>/.test(h1) && /<w:outlineLvl w:val="0"\/>/.test(h1) && /eastAsia="黑体"/.test(h1),
    h1.slice(0, 300));
  check('样式「标题 2」定义已重写为 小三(30) + outlineLvl 1',
    /<w:sz w:val="30"\/>/.test(h2) && /<w:outlineLvl w:val="1"\/>/.test(h2),
    h2.slice(0, 300));
  check('样式名没有被改掉（仍是 heading 1 / heading 2）',
    /w:name w:val="heading 1"\/>/.test(h1) && /w:name w:val="heading 2"\/>/.test(h2));
}

/* ============================ 4. 强弱编号规则 ============================ */

function testRuleStrength() {
  console.log('\n=== 4. 编号规则的可信度分级 ===');
  const I = DocxCore._internal.heuristicInfo;
  const strong = ['第1章 绪论', '1.1 概述', '1.1.1 概述', '1.1.1.1 概述', '第2节 概述', '参考文献'];
  const weak = ['一、项目背景', '（一）总体方案', '（1）硬件设计', '1. 工艺对象分析', '1、概述说明'];
  strong.forEach(function (s) {
    var r = I(s);
    check('「' + s + '」= 强规则（可覆盖样式层级）', !!r && r.strong === true, JSON.stringify(r));
  });
  weak.forEach(function (s) {
    var r = I(s);
    check('「' + s + '」= 弱规则（不覆盖样式层级）', !!r && r.strong === false, JSON.stringify(r));
  });
}

/* ============================ 5. 样式同步 ============================ */

async function testStyleSync() {
  console.log('\n=== 5. 目标文档完全没有标题样式时，应新建并挂上 ===');
  const buf = await buildDocx(PLAIN_ITEMS, { styles: BARE_STYLES_XML });
  const res = await DocxCore.applyFormat(buf.buffer.slice(buf.byteOffset, buf.byteOffset + buf.byteLength),
    DocxCore.BUILTIN_PRESETS['report-cn'].spec, NO_CHAPTER);
  const r = res.report;
  console.log('  新建样式: ' + JSON.stringify(r.stylesCreated));

  check('三级标题都被认出来了（1/1/1）',
    r.headings[1] === 1 && r.headings[2] === 1 && r.headings[3] === 1, JSON.stringify(r.headings));
  check('提示里列出了新建的样式（标题 1/2/3 + 题注）',
    r.stylesCreated.length >= 3 && r.stylesCreated.some(function (s) { return /标题 1/.test(s); }),
    JSON.stringify(r.stylesCreated));

  const zip = await global.JSZip.loadAsync(res.data);
  const ps = getParas(await zip.file('word/document.xml').async('string'));
  const sty = await zip.file('word/styles.xml').async('string');

  /* 段落挂的 pStyle 必须真的在 styles.xml 里存在 */
  const used = (ps.map(function (p) { return (p.xml.match(/w:pStyle w:val="([^"]+)"/) || [])[1]; }))
    .filter(Boolean);
  const missing = used.filter(function (id) { return sty.indexOf('w:styleId="' + id + '"') < 0; });
  check('所有 pStyle 引用都能在 styles.xml 里找到', missing.length === 0, JSON.stringify(missing));
  check('第1章 / 1.1 / 1.1.1 分别挂了 3 个不同的标题样式',
    used.length >= 3 && new Set(used.slice(0, 3)).size === 3, JSON.stringify(used));

  const styleBlocks = sty.split('<w:style ').slice(1).map(function (s) { return '<w:style ' + s; });
  const h1 = styleBlocks.filter(function (s) { return /<w:name w:val="标题 1"\/>/.test(s); })[0] || '';
  check('新建的「标题 1」定义 = 黑体 小二(36) 居中 outlineLvl 0',
    /<w:sz w:val="36"\/>/.test(h1) && /<w:jc w:val="center"\/>/.test(h1) && /<w:outlineLvl w:val="0"\/>/.test(h1),
    h1.slice(0, 300));
}

/* ============================ 6. 标题不该继承正文的首行缩进 ============================ */

async function testHeadingIndent() {
  console.log('\n=== 6. 标题不得继承正文样式的首行缩进 ===');
  /* 文档里 标题N 样式都 basedOn 正文（Normal）。模板的正文带「首行缩进 2 字符」，
     所以只把标题段落的 w:ind 删掉是不够的 —— Word 会顺着样式链继承出两格缩进。 */
  const buf = await buildDocx(DOC_ITEMS);
  const res = await DocxCore.applyFormat(buf.buffer.slice(buf.byteOffset, buf.byteOffset + buf.byteLength),
    DocxCore.BUILTIN_PRESETS['report-cn'].spec, NO_CHAPTER);
  const zip = await global.JSZip.loadAsync(res.data);
  const ps = getParas(await zip.file('word/document.xml').async('string'));
  const sty = await zip.file('word/styles.xml').async('string');

  [['一级', 1], ['二级', 3], ['三级', 4]].forEach(function (pair) {
    const name = pair[0], idx = pair[1];
    const ind = (ps[idx].xml.match(/<w:ind\b[^>]*\/>/) || [''])[0];
    check('「' + name + '标题」显式写了 0 缩进（不给样式链留继承口子）',
      /w:firstLineChars="0"/.test(ind) && /w:firstLine="0"/.test(ind), ps[idx].xml.slice(0, 200));
  });

  check('正文段仍带首行缩进 2 字符',
    /w:firstLineChars="200"/.test(ps[10].xml), ps[10].xml.slice(0, 200));

  const blocks = sty.split('<w:style ').slice(1).map(function (s) { return '<w:style ' + s; });
  const pick = function (n) { return blocks.filter(function (s) { return new RegExp('<w:name w:val="' + n + '"').test(s); })[0] || ''; };
  const h2 = pick('heading 2');
  const h3 = pick('heading 3');
  const normal = pick('Normal');
  check('样式「标题 2」定义里也显式写了 0 缩进',
    /<w:ind\b[^>]*w:firstLineChars="0"/.test(h2), h2.slice(0, 320));
  check('样式「标题 3」定义里也显式写了 0 缩进',
    /<w:ind\b[^>]*w:firstLineChars="0"/.test(h3), h3.slice(0, 320));
  check('样式「正文」定义里保留首行缩进 2 字符',
    /<w:ind\b[^>]*w:firstLineChars="200"/.test(normal), normal.slice(0, 320));
}

/* ============================ 7. 清掉遗留的水平线 ============================ */

async function testRemoveRules() {
  console.log('\n=== 7. 清掉网页/AI 粘过来的水平线（<hr>） ===');
  const buf = await buildDocx(RULE_ITEMS);
  const ab = buf.buffer.slice(buf.byteOffset, buf.byteOffset + buf.byteLength);

  const res = await DocxCore.applyFormat(ab, DocxCore.BUILTIN_PRESETS['report-cn'].spec, NO_CHAPTER);
  check('报告统计到 2 条水平线', res.report.removedRules === 2, res.report.removedRules);

  const zip = await global.JSZip.loadAsync(res.data);
  const xml = await zip.file('word/document.xml').async('string');
  check('产物里已无 o:hr 水平线图形', !/o:hr\s*=/.test(xml));
  check('产物里已无 w:pict', !/<w:pict/.test(xml));

  const ps = getParas(xml);
  check('只装水平线的那一段被整段删除（6 段 → 5 段）', ps.length === 5, ps.length);
  check('「文字 + 水平线」的段落保留了文字',
    ps.some(function (p) { return /带线的正文段落/.test(p.xml); }));
  check('第1章 / 第2章 两章标题仍完好',
    ps.filter(function (p) { return /第[12]章/.test(p.xml) && /w:outlineLvl/.test(p.xml); }).length === 2);

  /* 关掉这个开关时应原样保留 */
  const keep = await DocxCore.applyFormat(ab, DocxCore.BUILTIN_PRESETS['report-cn'].spec,
    { removeRules: false, chapter: NO_CHAPTER.chapter });
  const zip2 = await global.JSZip.loadAsync(keep.data);
  const xml2 = await zip2.file('word/document.xml').async('string');
  check('取消勾选后水平线原样保留', (xml2.match(/o:hr="t"/g) || []).length === 2);
  check('取消勾选后 removedRules 为 0', keep.report.removedRules === 0);
}

/* ============================ 8. 一级标题（章）版式 ============================ */

async function testChapterLayout() {
  console.log('\n=== 8. 一级标题：另起一页 + 上下各空一行（空行同用一级标题格式） ===');

  const buf = await buildDocx(CHAPTER_ITEMS);
  const ab = buf.buffer.slice(buf.byteOffset, buf.byteOffset + buf.byteLength);
  const res = await DocxCore.applyFormat(ab, DocxCore.BUILTIN_PRESETS['report-cn'].spec, {});
  const r = res.report;
  console.log('  报告: ' + JSON.stringify({ breaks: r.chapterBreaks, blanks: r.chapterBlanks, h: r.headings }));

  check('三章都识别为一级标题', r.headings[1] === 3, JSON.stringify(r.headings));
  check('三章都加了分页（第1章前面本来就有内容）', r.chapterBreaks === 3, r.chapterBreaks);
  check('补了 6 个空行（3 章 × 前后各 1）', r.chapterBlanks === 6, r.chapterBlanks);

  const zip = await global.JSZip.loadAsync(res.data);
  const ps = getParas(await zip.file('word/document.xml').async('string'));
  const seq = ps.map(function (p) { return p.text.trim() ? p.text.trim().slice(0, 14) : '(空行)'; });
  console.log('  段落序列: ' + seq.join(' | '));

  /* 第1章 前一个空行带分页；每个章标题前后各有一个空行 */
  const idx = {};
  ps.forEach(function (p, i) { if (/^第[123]章/.test(p.text.trim())) idx[p.text.trim().slice(0, 2)] = i; });
  [['第1章', '第1'], ['第2章', '第2'], ['第3章', '第3']].forEach(function (pair) {
    const i = idx[pair[1]];
    check(pair[0] + ' 前面是一个空行', i > 0 && !ps[i - 1].text.trim(), ps[i - 1] && ps[i - 1].xml.slice(0, 120));
    check(pair[0] + ' 后面是一个空行', i < ps.length - 1 && !ps[i + 1].text.trim());
    check(pair[0] + ' 前面那个空行带分页符',
      i > 0 && /<w:pageBreakBefore\/>/.test(ps[i - 1].xml), ps[i - 1] && ps[i - 1].xml.slice(0, 200));
    check(pair[0] + ' 前面的空行套的是一级标题格式（黑体 小二 36）',
      i > 0 && /w:eastAsia="黑体"/.test(ps[i - 1].xml) && /<w:sz w:val="36"\/>/.test(ps[i - 1].xml));
    check(pair[0] + ' 后面的空行同样是一级标题格式',
      i < ps.length - 1 && /w:eastAsia="黑体"/.test(ps[i + 1].xml) && /<w:sz w:val="36"\/>/.test(ps[i + 1].xml));
    check(pair[0] + ' 章标题本身仍居中',
      i !== undefined && /<w:jc w:val="center"\/>/.test(ps[i].xml));
  });

  /* 空行不能带首行缩进 */
  const blanks = ps.filter(function (p) { return !p.text.trim() && /w:eastAsia="黑体"/.test(p.xml); });
  check('所有标题格式空行都没有首行缩进',
    blanks.length === 6 && blanks.every(function (p) { return /w:firstLineChars="0"/.test(p.xml); }),
    blanks.length);

  /* 二级标题本身不该被波及（它后面紧跟章前空行是正常的，所以只查它自己那一行） */
  const h2i = ps.findIndex(function (p) { return /^1\.1 /.test(p.text.trim()); });
  check('二级标题自己没被加分页符',
    h2i > 0 && !/<w:pageBreakBefore\/>/.test(ps[h2i].xml), ps[h2i] && ps[h2i].xml.slice(0, 200));
  check('二级标题前一行不是「标题格式的空行」',
    h2i > 0 && !/w:eastAsia="黑体"/.test(ps[h2i - 1].xml), ps[h2i - 1] && ps[h2i - 1].xml.slice(0, 200));
  check('原有正文段落一个都没少',
    ps.filter(function (p) { return /正文段落/.test(p.text); }).length === 2);

  /* 原文标题上下本来就留了空行时：就地改造，不再叠出第二行 */
  const buf3 = await buildDocx(PRE_BLANKED_ITEMS);
  const res3 = await DocxCore.applyFormat(
    buf3.buffer.slice(buf3.byteOffset, buf3.byteOffset + buf3.byteLength),
    DocxCore.BUILTIN_PRESETS['report-cn'].spec, {});
  check('已有的空行被复用而不是叠加（2 章 → 4 个空行，不新增段落）',
    res3.report.chapterBlanks === 4 && res3.report.chapterBreaks === 2,
    JSON.stringify({ blanks: res3.report.chapterBlanks, breaks: res3.report.chapterBreaks }));
  const ps3 = getParas(await (await global.JSZip.loadAsync(res3.data)).file('word/document.xml').async('string'));
  check('复用后的段落数 = 原文 6 + 新增 2', ps3.length === 8, ps3.length);
  check('被复用的那个空行拿到了分页符',
    /<w:pageBreakBefore\/>/.test(ps3[1].xml) && !ps3[1].text.trim(), ps3[1].xml.slice(0, 160));
  check('第2章前那一行是新建的标题格式空行',
    /<w:pageBreakBefore\/>/.test(ps3[5].xml) && /w:eastAsia="黑体"/.test(ps3[5].xml), ps3[5].xml.slice(0, 200));

  /* 第1章就是文档开头时：不再分页，但空行照加 */
  const buf2 = await buildDocx(CHAPTER_FIRST_ITEMS);
  const res2 = await DocxCore.applyFormat(
    buf2.buffer.slice(buf2.byteOffset, buf2.byteOffset + buf2.byteLength),
    DocxCore.BUILTIN_PRESETS['report-cn'].spec, {});
  check('第1章在文档开头时少一次分页（2 章 → 1 次分页）', res2.report.chapterBreaks === 1, res2.report.chapterBreaks);
  check('第1章在文档开头时空行照加（2 章 × 2 = 4）', res2.report.chapterBlanks === 4, res2.report.chapterBlanks);
  const ps2 = getParas(await (await global.JSZip.loadAsync(res2.data)).file('word/document.xml').async('string'));
  const firstH1 = ps2.findIndex(function (p) { return !p.text.trim() && /w:eastAsia="黑体"/.test(p.xml); });
  check('文档开头那个空行不带分页符', firstH1 >= 0 && !/<w:pageBreakBefore\/>/.test(ps2[firstH1].xml));

  /* 关掉这项版式应完全不生效 */
  const off = await DocxCore.applyFormat(ab, DocxCore.BUILTIN_PRESETS['report-cn'].spec, {
    chapter: { pageBreakBefore: false, blankLinesBefore: 0, blankLinesAfter: 0 }
  });
  check('关闭后既不分页也不插空行',
    off.report.chapterBreaks === 0 && off.report.chapterBlanks === 0,
    JSON.stringify({ b: off.report.chapterBreaks, n: off.report.chapterBlanks }));
  const offPs = getParas(await (await global.JSZip.loadAsync(off.data)).file('word/document.xml').async('string'));
  check('关闭后段落数等于原文段落数', offPs.length === CHAPTER_ITEMS.length, offPs.length);

  /* 其它预设不受影响 */
  const ver = DocxCore.BUILTIN_PRESETS['thesis-cn'];
  check('「中文学术论文通用规范」预设不带这项版式', !ver.spec.chapter);
}

/* ============================ 9. 关闭文档网格 ============================ */

async function testDisableDocGrid() {
  console.log('\n=== 9. 自动关闭文档网格（标题被撑高的元凶） ===');

  /* 开着「指定行网格」的文档 → 应当被关掉 */
  const buf = await buildDocx(GRID_ITEMS, { sectPr: GRID_SECTPR });
  const ab = buf.buffer.slice(buf.byteOffset, buf.byteOffset + buf.byteLength);
  const res = await DocxCore.applyFormat(ab, DocxCore.BUILTIN_PRESETS['report-cn'].spec, NO_CHAPTER);

  check('报告统计到 1 处网格被关闭', res.report.docGridOff === 1, res.report.docGridOff);
  check('报告记录了原网格类型', res.report.docGridBefore === 'lines', res.report.docGridBefore);

  const zip = await global.JSZip.loadAsync(res.data);
  const xml = await zip.file('word/document.xml').async('string');
  check('产物里 docGrid 已无 w:type 属性（= 无网格）',
    /<w:docGrid(?![^>]*w:type)[^>]*>/.test(xml), (xml.match(/<w:docGrid[^>]*>/) || [])[0]);
  check('产物里不含 w:type="lines"', !/w:type="lines"/.test(xml));
  check('linePitch 等无关属性未被顺手删掉', /w:linePitch="312"/.test(xml));
  check('正文与标题段落仍然完好', getParas(xml).length === GRID_ITEMS.length);

  /* 关掉这个开关时应原样保留网格 */
  const keep = await DocxCore.applyFormat(ab, DocxCore.BUILTIN_PRESETS['report-cn'].spec,
    { disableGrid: false, chapter: NO_CHAPTER.chapter });
  const xml2 = await (await global.JSZip.loadAsync(keep.data)).file('word/document.xml').async('string');
  check('取消勾选后网格原样保留', /w:type="lines"/.test(xml2));
  check('取消勾选后 docGridOff 为 0', keep.report.docGridOff === 0, keep.report.docGridOff);

  /* 本来就无网格的文档：不该被写进任何改动 */
  const noGrid = await buildDocx(GRID_ITEMS, { sectPr: NOGRID_SECTPR });
  const nm = noGrid.buffer.slice(noGrid.byteOffset, noGrid.byteOffset + noGrid.byteLength);
  const res3 = await DocxCore.applyFormat(nm, DocxCore.BUILTIN_PRESETS['report-cn'].spec, NO_CHAPTER);
  check('本来无网格的文档统计为 0（不虚报）', res3.report.docGridOff === 0, res3.report.docGridOff);
  const xml3 = await (await global.JSZip.loadAsync(res3.data)).file('word/document.xml').async('string');
  check('本来无网格的文档产物仍无 w:type', !/[^>]*w:type=/.test((xml3.match(/<w:docGrid[^>]*>/) || [''])[0]));

  /* 纯函数单测：多种 type 取值 */
  const { DOMParser: DP } = require('@xmldom/xmldom');
  ['lines', 'linesAndChars', 'snapToChars'].forEach(function (t) {
    const dom = new DP().parseFromString(
      '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>'
      + '<w:sectPr><w:docGrid w:type="' + t + '" w:linePitch="312"/></w:sectPr></w:body></w:document>', 'text/xml');
    const r = DocxCore.disableDocGrid(dom);
    check('纯函数：type="' + t + '" 被关闭', r.changed === 1 && r.modeBefore === t, JSON.stringify(r));
  });

  const domDefault = new DP().parseFromString(
    '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>'
    + '<w:sectPr><w:docGrid w:type="default"/></w:sectPr></w:body></w:document>', 'text/xml');
  check('纯函数：type="default" 视为无网格、跳过',
    DocxCore.disableDocGrid(domDefault).changed === 0);

  const domEmpty = new DP().parseFromString(
    '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>'
    + '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/></w:sectPr></w:body></w:document>', 'text/xml');
  check('纯函数：没有 docGrid 元素也不报错',
    DocxCore.disableDocGrid(domEmpty).changed === 0);
}

/* ============================ 运行 ============================ */

(async function () {
  testHeuristic();
  await testEndToEnd();
  await testShiftedStyles();
  testRuleStrength();
  await testStyleSync();
  await testHeadingIndent();
  await testRemoveRules();
  await testChapterLayout();
  await testDisableDocGrid();
  console.log('\n---------------------------------------');
  console.log('  通过 ' + pass + ' 项，失败 ' + fail + ' 项');
  console.log('---------------------------------------\n');
  process.exit(fail ? 1 : 0);
})().catch(function (e) {
  console.error('测试异常:', (e && e.stack) || e);
  process.exit(1);
});
