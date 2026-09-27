/**
 * 测试用最小 docx 构造器
 * 中文版 Word 的常见形态：styleId 是数字，"标题 N" 只出现在 w:name 里。
 */
const path = require('path');

function getJSZip() {
  if (global.JSZip) return global.JSZip;
  return require(path.resolve(__dirname, '../../js/jszip.min.js'));
}

const CT_XML = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
  + '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
  + '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
  + '<Default Extension="xml" ContentType="application/xml"/>'
  + '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
  + '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>'
  + '</Types>';

const RELS_XML = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
  + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
  + '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>'
  + '</Relationships>';

const DOC_RELS_XML = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
  + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
  + '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
  + '</Relationships>';

/* styleId 2/3/4 = heading 1/2/3（对应大纲级别 0/1/2），16 = toc 1 */
const STYLES_XML = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
  + '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
  + '<w:docDefaults><w:rPrDefault><w:rPr>'
  + '<w:rFonts w:ascii="Times New Roman" w:eastAsia="宋体" w:hAnsi="Times New Roman"/>'
  + '<w:sz w:val="24"/><w:szCs w:val="24"/>'
  + '</w:rPr></w:rPrDefault></w:docDefaults>'
  + '<w:style w:type="paragraph" w:default="1" w:styleId="1"><w:name w:val="Normal"/></w:style>'
  + '<w:style w:type="paragraph" w:styleId="2"><w:name w:val="heading 1"/><w:basedOn w:val="1"/>'
  + '<w:pPr><w:outlineLvl w:val="0"/></w:pPr><w:rPr><w:sz w:val="32"/></w:rPr></w:style>'
  + '<w:style w:type="paragraph" w:styleId="3"><w:name w:val="heading 2"/><w:basedOn w:val="1"/>'
  + '<w:pPr><w:outlineLvl w:val="1"/></w:pPr><w:rPr><w:sz w:val="28"/></w:rPr></w:style>'
  + '<w:style w:type="paragraph" w:styleId="4"><w:name w:val="heading 3"/><w:basedOn w:val="1"/>'
  + '<w:pPr><w:outlineLvl w:val="2"/></w:pPr><w:rPr><w:sz w:val="24"/></w:rPr></w:style>'
  + '<w:style w:type="paragraph" w:styleId="5"><w:name w:val="heading 4"/><w:basedOn w:val="1"/>'
  + '<w:pPr><w:outlineLvl w:val="3"/></w:pPr><w:rPr><w:sz w:val="24"/></w:rPr></w:style>'
  + '<w:style w:type="paragraph" w:styleId="16"><w:name w:val="toc 1"/><w:basedOn w:val="1"/></w:style>'
  + '</w:styles>';

/**
 * 覆盖场景：
 *  - 空段套着「标题 1」当空行用（作者常见做法）
 *  - 有样式的一/二/三级标题
 *  - 手打编号、完全没有样式的标题
 *  - 无样式的「表X-Y / 图X-Y」题注
 */
const DOC_ITEMS = [
  { name: '空段落(标题1样式)', text: '', style: '2' },
  { name: '第1章 绪论', text: '第1章 绪论', style: '2' },
  { name: '空段落(标题1样式)', text: '', style: '2' },
  { name: '1.1 课题背景', text: '1.1 课题背景', style: '3' },
  { name: '1.1.1 系统功能', text: '1.1.1 系统功能', style: '4' },
  { name: '1.1.2 手打的三级标题', text: '1.1.2 手打的三级标题' },
  { name: '第2章 系统设计', text: '第2章 系统设计' },
  { name: '2.1 需求分析', text: '2.1 需求分析' },
  { name: '表2-1 题注', text: '表2-1 线性五杆结构各自由度随机反应数值特征' },
  { name: '图2-1 题注', text: '图2-1 试样的膨胀变化曲线' },
  { name: '正文段落', text: '正文段落，这里是普通的说明文字，讲清楚了系统要做什么。' }
];

function documentXml(items, opts) {
  opts = opts || {};
  const body = items.map(function (it) {
    const pPr = it.style ? ('<w:pPr><w:pStyle w:val="' + it.style + '"/></w:pPr>') : '';
    const runs = it.text
      ? '<w:r><w:rPr><w:rFonts w:hint="eastAsia"/></w:rPr><w:t xml:space="preserve">' + it.text + '</w:t></w:r>'
      : '';
    return '<w:p>' + pPr + runs + (it.raw || '') + '</w:p>';
  }).join('');

  return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    + '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
    + ' xmlns:v="urn:schemas-microsoft-com:vml"'
    + ' xmlns:o="urn:schemas-microsoft-com:office:office"'
    + ' xmlns:w10="urn:schemas-microsoft-com:office:word">'
    + '<w:body>' + body
    + (opts.sectPr || ('<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>'
      + '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="851" w:footer="992" w:gutter="0"/>'
      + '</w:sectPr>'))
    + '</w:body></w:document>';
}

/** Word 把 HTML / markdown 的 <hr> 存成这个形状（o:hr = horizontal rule） */
const HR_PICT = '<w:r><w:pict><v:rect id="_x0000_i1025" o:spt="1"'
  + ' style="height:1.5pt;width:432pt;" fillcolor="#A0A0A0" filled="t" stroked="f"'
  + ' coordsize="21600,21600" o:hr="t" o:hrstd="t" o:hralign="center">'
  + '<v:path/><v:fill on="t" focussize="0,0"/><v:stroke on="f"/><v:imagedata o:title=""/>'
  + '<o:lock v:ext="edit"/><w10:wrap type="none"/><w10:anchorlock/>'
  + '</v:rect></w:pict></w:r>';

/** 夹着水平线的文档：第2段只装一条线，第5段是「有文字 + 一条线」 */
const RULE_ITEMS = [
  { text: '第1章 绪论' },
  { raw: HR_PICT },
  { text: '正文段落，这里是普通的说明文字。' },
  { text: '1.1 课题背景' },
  { text: '带线的正文段落，文字和水平线在同一个段落里。', raw: HR_PICT },
  { text: '第2章 系统设计' }
];

/** 开着「指定行网格」的文档：这条 sectPr 会让标题这类倍数行距的段落被撑高 */
const GRID_SECTPR = '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>'
  + '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="851" w:footer="992" w:gutter="0"/>'
  + '<w:docGrid w:type="lines" w:linePitch="312" w:charSpace="0"/>'
  + '</w:sectPr>';

/** 本来就是「无网格」的文档（w:type 缺省）—— 引擎不该去动它 */
const NOGRID_SECTPR = '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>'
  + '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="851" w:footer="992" w:gutter="0"/>'
  + '<w:docGrid w:linePitch="312" w:charSpace="0"/>'
  + '</w:sectPr>';

const GRID_ITEMS = [
  { text: '第1章 绪论' },
  { text: '正文段落，这里是普通的说明文字，讲清楚了系统要做什么。' },
  { text: '1.1 课题背景' },
  { text: '第2章 系统设计' }
];

/** 返回 Uint8Array */
async function buildDocx(items, opts) {
  items = items || DOC_ITEMS;
  opts = opts || {};
  const JSZip = getJSZip();
  const zip = new JSZip();
  zip.file('[Content_Types].xml', CT_XML);
  zip.file('_rels/.rels', RELS_XML);
  zip.file('word/_rels/document.xml.rels', DOC_RELS_XML);
  zip.file('word/document.xml', documentXml(items, opts));
  zip.file('word/styles.xml', opts.styles || STYLES_XML);
  return await zip.generateAsync({ type: 'uint8array' });
}

/** 只有 Normal 的最小样式表：用于验证「目标文档没有标题样式」时能新建出来 */
const BARE_STYLES_XML = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
  + '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
  + '<w:docDefaults><w:rPrDefault><w:rPr>'
  + '<w:rFonts w:ascii="Times New Roman" w:eastAsia="宋体" w:hAnsi="Times New Roman"/>'
  + '<w:sz w:val="21"/><w:szCs w:val="21"/>'
  + '</w:rPr></w:rPrDefault></w:docDefaults>'
  + '<w:style w:type="paragraph" w:default="1" w:styleId="1"><w:name w:val="Normal"/></w:style>'
  + '</w:styles>';

/** 完全没有样式、全靠手打编号的标题（最考验新建样式的能力） */
const PLAIN_ITEMS = [
  { text: '第1章 绪论' },
  { text: '1.1 课题背景' },
  { text: '1.1.1 系统功能' },
  { text: '正文段落，这里是普通的说明文字，讲清楚了系统要做什么。' }
];

/** 一级标题（章）版式：第1章前面还有内容，所以三章都该另起一页 */
const CHAPTER_ITEMS = [
  { text: '本报告依据课程设计任务书编写。' },
  { text: '第1章 绪论' },
  { text: '正文段落，这里是普通的说明文字，讲清楚了系统要做什么。' },
  { text: '1.1 课题背景' },
  { text: '第2章 系统设计' },
  { text: '正文段落，讲一下系统整体怎么设计。' },
  { text: '第3章 结论' }
];

/** 第1章本身就是文档开头：前面没内容，不该再分页（否则空出一整页） */
const CHAPTER_FIRST_ITEMS = [
  { text: '第1章 绪论' },
  { text: '正文段落，这里是普通的说明文字。' },
  { text: '第2章 系统设计' }
];

/** 标题上下原本就留了空行：应当就地改造，不该再叠出第二行 */
const PRE_BLANKED_ITEMS = [
  { text: '前置说明。' },
  { text: '' },
  { text: '第1章 绪论' },
  { text: '' },
  { text: '正文段落，这里是普通的说明文字。' },
  { text: '第2章 系统设计' }
];

/** 样式整体错位一级的用例：第1章套 heading 2、1.1 套 heading 3、1.1.1 套 heading 4 */
const SHIFTED_ITEMS = [
  { text: '第1章 绪论', style: '3' },
  { text: '1.1 课题背景与意义', style: '4' },
  { text: '1.1.1 课题背景', style: '5' },
  { text: '一、项目背景', style: '2' },
  { text: '正文段落，这里是普通的说明文字，讲清楚了系统要做什么。' }
];

module.exports = { buildDocx, documentXml, DOC_ITEMS, SHIFTED_ITEMS, PLAIN_ITEMS, RULE_ITEMS,
  CHAPTER_ITEMS, CHAPTER_FIRST_ITEMS, PRE_BLANKED_ITEMS, GRID_ITEMS,
  GRID_SECTPR, NOGRID_SECTPR,
  HR_PICT, STYLES_XML, BARE_STYLES_XML };
