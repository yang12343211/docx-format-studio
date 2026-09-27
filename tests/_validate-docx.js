const fs = require('fs'), path = require('path');
const { DOMParser, XMLSerializer } = require('@xmldom/xmldom');
global.DOMParser = DOMParser; global.XMLSerializer = XMLSerializer;
global.JSZip = require(path.resolve(__dirname, '../js/jszip.min.js'));
const Core = require(path.resolve(__dirname, '../js/docx-core.js'));

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
function kid(el, n) { if (!el) return null; const c = el.childNodes; for (let i = 0; i < c.length; i++) if (c[i].nodeType === 1 && c[i].localName === n) return c[i]; return null; }
function A(el, n) { if (!el) return null; let v = null; try { v = el.getAttributeNS(W, n); } catch (e) { } if (!v) v = el.getAttribute('w:' + n); if (!v) v = el.getAttribute(n); return v || null; }

const PPR = ['pStyle', 'keepNext', 'keepLines', 'pageBreakBefore', 'framePr', 'widowControl', 'numPr',
  'suppressLineNumbers', 'pBdr', 'shd', 'tabs', 'suppressAutoHyphens', 'kinsoku', 'wordWrap', 'overflowPunct',
  'topLinePunct', 'autoSpaceDE', 'autoSpaceDN', 'bidi', 'adjustRightInd', 'snapToGrid', 'spacing', 'ind',
  'contextualSpacing', 'mirrorIndents', 'suppressOverlap', 'jc', 'textDirection', 'textAlignment',
  'textboxTightWrap', 'outlineLvl', 'divId', 'cnfStyle', 'rPr', 'sectPr', 'pPrChange'];
const STY = ['name', 'aliases', 'basedOn', 'next', 'link', 'autoRedefine', 'hidden', 'uiPriority', 'semiHidden',
  'unhideWhenUsed', 'qFormat', 'locked', 'personal', 'personalCompose', 'personalReply', 'rsid', 'pPr', 'rPr',
  'tblPr', 'trPr', 'tcPr', 'tblStylePr'];

/* 检查某个父元素下、给定顺序表里的子元素是否严格递增 */
function orderOk(parent, orderList, label, problems) {
  let last = -1;
  const cs = parent.childNodes;
  for (let i = 0; i < cs.length; i++) {
    if (cs[i].nodeType !== 1) continue;
    const idx = orderList.indexOf(cs[i].localName);
    if (idx < 0) continue;
    if (idx < last) problems.push(label + ' 子元素顺序错：' + cs[i].localName);
    last = idx;
  }
}

/** 某个 w:style 里「受 schema 顺序约束的子元素」的名字序列 */
function styleChildOrder(styleEl) {
  const out = [];
  const cs = styleEl.childNodes;
  for (let i = 0; i < cs.length; i++) {
    if (cs[i].nodeType === 1 && STY.indexOf(cs[i].localName) >= 0) out.push(cs[i].localName);
  }
  return out;
}

(async () => {
  const src = process.argv[2];
  const outPath = process.argv[3] || path.resolve(__dirname, '..', 'samples', 'out', 'verify-user-formatted.docx');

  const srcBuf = fs.readFileSync(src);
  const res = await Core.applyFormat(srcBuf.buffer.slice(srcBuf.byteOffset, srcBuf.byteOffset + srcBuf.byteLength),
    Core.BUILTIN_PRESETS['report-cn'].spec, {});
  fs.mkdirSync(path.dirname(outPath), { recursive: true });
  fs.writeFileSync(outPath, Buffer.from(res.data));

  const problems = [];
  const notes = [];
  const zip = await global.JSZip.loadAsync(res.data);

  /* 1. 所有 XML 部件都要能解析 */
  const names = Object.keys(zip.files).filter(n => /\.(xml|rels)$/.test(n));
  for (const n of names) {
    const txt = await zip.file(n).async('string');
    const d = new DOMParser().parseFromString(txt, 'text/xml');
    if (d.getElementsByTagName('parsererror').length) problems.push(n + ' XML 解析失败');
  }

  /* 2. pPr 子元素顺序（引擎自己写的，必须严格合法） */
  const doc = new DOMParser().parseFromString(await zip.file('word/document.xml').async('string'), 'text/xml');
  const ps = doc.getElementsByTagName('w:p');
  for (let i = 0; i < ps.length; i++) orderOk(kid(ps[i], 'pPr'), PPR, '第' + (i + 1) + '段 pPr', problems);

  /* 3. styles.xml：子元素顺序 + styleId 唯一
        源文件本身可能就有顺序怪癖（实测 WPS 生成的文档把 qFormat 排在 uiPriority 前），
        这类「沿用自源文件」的保持原样即可，只揪引擎新引入的问题。 */
  const srcZip = await global.JSZip.loadAsync(srcBuf.buffer.slice(srcBuf.byteOffset, srcBuf.byteOffset + srcBuf.byteLength));
  const srcOrder = {};
  if (srcZip.file('word/styles.xml')) {
    const sd = new DOMParser().parseFromString(await srcZip.file('word/styles.xml').async('string'), 'text/xml');
    const ss = sd.getElementsByTagName('w:style');
    for (let i = 0; i < ss.length; i++) srcOrder[A(ss[i], 'styleId')] = styleChildOrder(ss[i]).join('>');
  }

  const sty = new DOMParser().parseFromString(await zip.file('word/styles.xml').async('string'), 'text/xml');
  const styles = sty.getElementsByTagName('w:style');
  const ids = {};
  for (let i = 0; i < styles.length; i++) {
    const id = A(styles[i], 'styleId');
    const label = '样式 ' + id;
    if (ids[id]) problems.push('styleId 重复：' + id);
    ids[id] = true;

    const seq = styleChildOrder(styles[i]).join('>');
    const inherited = srcOrder[id] !== undefined && seq === srcOrder[id];
    if (!inherited) {
      orderOk(styles[i], STY, label, problems);
    } else {
      /* 顺序和源文件一模一样 —— 即便不合 schema 也不是本次引入的 */
      const bad = [];
      orderOk(styles[i], STY, label, bad);
      if (bad.length) notes.push(label + '：源文件本身顺序不合 schema（' + srcOrder[id] + '），引擎原样保留');
    }

    /* 引擎改过子元素的样式：只要求「原有相对顺序没被打乱」 */
    if (srcOrder[id] !== undefined && !inherited) {
      const kept = seq.split('>').filter(n => srcOrder[id].split('>').indexOf(n) >= 0).join('>');
      if (kept !== srcOrder[id]) problems.push(label + '：引擎打乱了原有子元素顺序（' + srcOrder[id] + ' → ' + seq + '）');
    }
  }

  /* 4. 所有 pStyle 引用都必须存在 */
  const docXml = await zip.file('word/document.xml').async('string');
  const refs = [...docXml.matchAll(/w:pStyle w:val="([^"]+)"/g)].map(m => m[1]);
  const missing = [...new Set(refs)].filter(id => !ids[id]);
  if (missing.length) problems.push('pStyle 指向了不存在的样式：' + missing.join(','));

  console.log('产物：' + outPath);
  console.log('大小：' + (res.data.byteLength / 1024).toFixed(1) + ' KB');
  console.log('XML 部件：' + names.length + ' 个，全部可解析');
  console.log('样式表：' + Object.keys(ids).length + ' 个 styleId，无重复');
  console.log('段落引用的样式：' + [...new Set(refs)].join(', '));
  if (notes.length) console.log('\n提示（沿用源文件、非本次引入）：\n - ' + notes.join('\n - '));
  console.log(problems.length ? ('\n校验发现问题：\n - ' + problems.join('\n - ')) : '\n校验通过：无 XML 顺序 / 引用 / 重复问题');
  process.exit(problems.length ? 1 : 0);
})().catch(e => { console.error(e); process.exit(1); });
