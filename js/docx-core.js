/*!
 * docx-core.js — Word / WPS 文档格式规范「提取 → 套用」核心引擎
 *
 * 设计原则：
 *   1. 无第三方依赖（JSZip 由调用方以全局变量形式提供）
 *   2. 浏览器 / Node 通用（UMD）
 *   3. 只做「格式」不做「内容」：段落、run 的文本一个字都不动
 *
 * 对外 API：
 *   DocxCore.analyzeTemplate(arrayBuffer)            -> Promise<Spec>
 *   DocxCore.applyFormat(arrayBuffer, spec, options) -> Promise<{blob|buffer, report}>
 *   DocxCore.describeSpec(spec)                      -> 可读的规范描述对象
 *   DocxCore.BUILTIN_PRESETS                         -> 内置规范预设
 */
(function (root, factory) {
  if (typeof module === 'object' && typeof module.exports === 'object') {
    module.exports = factory();
  } else {
    root.DocxCore = factory();
  }
})(typeof globalThis !== 'undefined' ? globalThis : this, function () {
  'use strict';

  /* ================================================================== *
   * 一、常量
   * ================================================================== */

  var W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
  var DOC_PATH = 'word/document.xml';
  var STY_PATH = 'word/styles.xml';
  var DOCX_MIME = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document';

  /* pPr 子元素合法顺序（OOXML schema 规定，顺序错误会让 Word 报错） */
  var PPR_ORDER = ['pStyle', 'keepNext', 'keepLines', 'pageBreakBefore', 'framePr', 'widowControl',
    'numPr', 'suppressLineNumbers', 'pBdr', 'shd', 'tabs', 'suppressAutoHyphens', 'kinsoku',
    'wordWrap', 'overflowPunct', 'topLinePunct', 'autoSpaceDE', 'autoSpaceDN', 'bidi',
    'adjustRightInd', 'snapToGrid', 'spacing', 'ind', 'contextualSpacing', 'mirrorIndents',
    'suppressOverlap', 'jc', 'textDirection', 'textAlignment', 'textboxTightWrap', 'outlineLvl',
    'divId', 'cnfStyle', 'rPr', 'sectPr', 'pPrChange'];

  /* rPr 子元素合法顺序 */
  var RPR_ORDER = ['rStyle', 'rFonts', 'b', 'bCs', 'i', 'iCs', 'caps', 'smallCaps', 'strike',
    'dstrike', 'outline', 'shadow', 'emboss', 'imprint', 'noProof', 'snapToGrid', 'vanish',
    'webHidden', 'color', 'spacing', 'w', 'kern', 'position', 'sz', 'szCs', 'highlight', 'u',
    'effect', 'bdr', 'shd', 'fitText', 'vertAlign', 'rtl', 'cs', 'em', 'lang', 'eastAsianLayout',
    'specVanish', 'oMath'];

  /* w:style 子元素合法顺序（CT_Style，顺序错误 Word 会报错） */
  var STYLE_ORDER = ['name', 'aliases', 'basedOn', 'next', 'link', 'autoRedefine', 'hidden',
    'uiPriority', 'semiHidden', 'unhideWhenUsed', 'qFormat', 'locked', 'personal',
    'personalCompose', 'personalReply', 'rsid', 'pPr', 'rPr', 'tblPr', 'trPr', 'tcPr', 'tblStylePr'];

  /* 显式"零缩进"。
     为什么不能只把 w:ind 删掉了事：标题样式通常 basedOn 正文，而模板的正文带
     「首行缩进 2 字符」。删掉段落的 w:ind 之后，Word 会顺着样式链把正文的缩进
     继承下来 —— 于是"没设首行缩进"的二级/三级标题反而缩进了两格。
     必须显式写 0 才能截断继承链（Word 自己在「段落」对话框里把缩进清零时，
     写的就是这组属性）。 */
  var ZERO_IND = { firstLine: 0, firstLineChars: 0, left: 0, leftChars: 0 };

  /* 中文字号 <-> 半磅 */
  var CN_SIZES = [
    [84, '初号'], [72, '小初'], [52, '一号'], [48, '小一'], [44, '二号'], [36, '小二'],
    [32, '三号'], [30, '小三'], [28, '四号'], [24, '小四'], [21, '五号'], [18, '小五'],
    [15, '六号'], [13, '小六'], [11, '七号'], [10, '八号']
  ];

  /* ================================================================== *
   * 二、DOM / XML 基础工具
   * ================================================================== */

  function isEl(n) { return n && n.nodeType === 1; }

  function parseXml(text) {
    if (typeof DOMParser === 'undefined') throw new Error('当前环境缺少 DOMParser');
    var doc = new DOMParser().parseFromString(text, 'text/xml');
    if (!doc || !doc.documentElement) throw new Error('XML 解析失败');
    /* xmldom / 浏览器对解析错误的处理不同，这里做双保险 */
    if (doc.documentElement.nodeName === 'parsererror' ||
        (doc.getElementsByTagName && doc.getElementsByTagName('parsererror').length)) {
      throw new Error('XML 解析失败：文档结构异常');
    }
    return doc;
  }

  function serializeXml(node) {
    var s = new XMLSerializer().serializeToString(node);
    if (!/^\s*<\?xml/.test(s)) {
      s = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n' + s;
    }
    return s;
  }

  /** 取直接子元素 */
  function kid(el, localName) {
    if (!el) return null;
    var cs = el.childNodes;
    for (var i = 0; i < cs.length; i++) {
      var c = cs[i];
      if (c.nodeType === 1 && c.localName === localName) return c;
    }
    return null;
  }

  /** 取直接子元素列表 */
  function kids(el, localName) {
    var out = [];
    if (!el) return out;
    var cs = el.childNodes;
    for (var i = 0; i < cs.length; i++) {
      var c = cs[i];
      if (c.nodeType === 1 && (!localName || c.localName === localName)) out.push(c);
    }
    return out;
  }

  /** 读取 w:xxx 属性值 */
  function attr(el, name) {
    if (!el || !el.getAttribute) return null;
    var v = null;
    try { v = el.getAttributeNS ? el.getAttributeNS(W_NS, name) : null; } catch (e) { v = null; }
    if (v === null || v === undefined || v === '') v = el.getAttribute('w:' + name);
    if (v === null || v === undefined || v === '') v = el.getAttribute(name);
    return (v === null || v === undefined || v === '') ? null : v;
  }

  function setAttr(el, name, value) {
    el.setAttributeNS(W_NS, 'w:' + name, String(value));
    return el;
  }

  function num(v) {
    if (v === null || v === undefined || v === '') return null;
    var n = parseInt(v, 10);
    return isNaN(n) ? null : n;
  }

  /** w:b / w:i 这类开关型元素：存在即 true，除非 val 显式为假 */
  function onOff(el) {
    if (!el) return undefined;
    var v = attr(el, 'val');
    if (v === null) return true;
    v = String(v).toLowerCase();
    if (v === '0' || v === 'false' || v === 'off' || v === 'no') return false;
    return true;
  }

  var NSMAP = {
    w: 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
    r: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    wp: 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
    a: 'http://schemas.openxmlformats.org/drawingml/2006/main',
    pic: 'http://schemas.openxmlformats.org/drawingml/2006/picture',
    m: 'http://schemas.openxmlformats.org/officeDocument/2006/math',
    mc: 'http://schemas.openxmlformats.org/markup-compatibility/2006',
    v: 'urn:schemas-microsoft-com:vml',
    o: 'urn:schemas-microsoft-com:office:office',
    w10: 'urn:schemas-microsoft-com:office:word'
  };

  /** 创建带命名空间的元素，created 元素通过 setAttributeNS 自动带上前缀 */
  function createEl(doc, localName, ns) {
    return doc.createElementNS(ns || W_NS, 'w:' + localName);
  }

  /* ================================================================== *
   * 三、格式属性解析（把 XML 片段翻译成普通对象）
   * ================================================================== */

  function parseRPr(rPr) {
    var o = {};
    if (!rPr) return o;

    var rf = kid(rPr, 'rFonts');
    if (rf) {
      var f = {};
      ['ascii', 'eastAsia', 'hAnsi', 'cs', 'hint'].forEach(function (k) {
        var v = attr(rf, k);
        if (v) f[k] = v;
      });
      if (Object.keys(f).length) o.font = f;
    }

    var sz = num(attr(kid(rPr, 'sz'), 'val'));
    if (sz !== null) o.size = sz;
    var szCs = num(attr(kid(rPr, 'szCs'), 'val'));
    if (szCs !== null) o.sizeCs = szCs;

    var b = onOff(kid(rPr, 'b'));
    if (b !== undefined) o.bold = b;
    var bCs = onOff(kid(rPr, 'bCs'));
    if (bCs !== undefined) o.boldCs = bCs;
    var i = onOff(kid(rPr, 'i'));
    if (i !== undefined) o.italic = i;
    var iCs = onOff(kid(rPr, 'iCs'));
    if (iCs !== undefined) o.italicCs = iCs;

    var color = attr(kid(rPr, 'color'), 'val');
    if (color && color.toLowerCase() !== 'auto') o.color = color;

    var u = kid(rPr, 'u');
    if (u) {
      var uv = attr(u, 'val');
      o.underline = (uv && uv !== 'none') ? uv : null;
    }

    var spacing = num(attr(kid(rPr, 'spacing'), 'val'));
    if (spacing !== null) o.charSpacing = spacing;
    var kern = num(attr(kid(rPr, 'kern'), 'val'));
    if (kern !== null) o.kern = kern;
    var position = num(attr(kid(rPr, 'position'), 'val'));
    if (position !== null) o.position = position;

    var lang = kid(rPr, 'lang');
    if (lang) {
      var lg = {};
      ['val', 'eastAsia'].forEach(function (k) { var v = attr(lang, k); if (v) lg[k] = v; });
      if (Object.keys(lg).length) o.lang = lg;
    }

    var vert = attr(kid(rPr, 'vertAlign'), 'val');
    if (vert) o.vertAlign = vert;

    return o;
  }

  function parsePPr(pPr) {
    var o = {};
    if (!pPr) return o;

    var jc = attr(kid(pPr, 'jc'), 'val');
    if (jc) o.align = jc;

    var sp = kid(pPr, 'spacing');
    if (sp) {
      var s = {};
      [['before', 'before'], ['after', 'after'], ['line', 'line'], ['lineRule', 'lineRule'],
       ['beforeLines', 'beforeLines'], ['afterLines', 'afterLines'],
       ['beforeAutospacing', 'beforeAutospacing'], ['afterAutospacing', 'afterAutospacing']
      ].forEach(function (pair) {
        var v = attr(sp, pair[1]);
        if (v === null) return;
        s[pair[0]] = (pair[1] === 'lineRule') ? v : num(v);
      });
      if (Object.keys(s).length) o.spacing = s;
    }

    var ind = kid(pPr, 'ind');
    if (ind) {
      var m = {};
      ['firstLine', 'firstLineChars', 'hanging', 'hangingChars', 'left', 'leftChars',
       'right', 'rightChars', 'start', 'startChars', 'end', 'endChars'
      ].forEach(function (k) {
        var v = num(attr(ind, k));
        if (v !== null) m[k] = v;
      });
      if (Object.keys(m).length) o.ind = m;
    }

    var ol = num(attr(kid(pPr, 'outlineLvl'), 'val'));
    if (ol !== null) o.outlineLvl = ol;

    var kn = onOff(kid(pPr, 'keepNext'));
    if (kn !== undefined) o.keepNext = kn;
    var kl = onOff(kid(pPr, 'keepLines'));
    if (kl !== undefined) o.keepLines = kl;
    var pbb = onOff(kid(pPr, 'pageBreakBefore'));
    if (pbb !== undefined) o.pageBreakBefore = pbb;
    var ctx = onOff(kid(pPr, 'contextualSpacing'));
    if (ctx !== undefined) o.contextualSpacing = ctx;

    var sg = onOff(kid(pPr, 'snapToGrid'));
    if (sg !== undefined) o.snapToGrid = sg;

    return o;
  }

  function parseSectPr(sectPr) {
    var o = {};
    if (!sectPr) return o;

    var pgSz = kid(sectPr, 'pgSz');
    if (pgSz) {
      o.page = {};
      var w = num(attr(pgSz, 'w')), h = num(attr(pgSz, 'h'));
      var orient = attr(pgSz, 'orient');
      if (w !== null) o.page.w = w;
      if (h !== null) o.page.h = h;
      if (orient) o.page.orient = orient;
    }

    var pgMar = kid(sectPr, 'pgMar');
    if (pgMar) {
      o.margin = {};
      ['top', 'right', 'bottom', 'left', 'header', 'footer', 'gutter'].forEach(function (k) {
        var v = num(attr(pgMar, k));
        if (v !== null) o.margin[k] = v;
      });
    }

    var dg = kid(sectPr, 'docGrid');
    if (dg) {
      o.docGrid = {};
      var t = attr(dg, 'type'), lp = num(attr(dg, 'linePitch')), cs = num(attr(dg, 'charSpace'));
      if (t) o.docGrid.type = t;
      if (lp !== null) o.docGrid.linePitch = lp;
      if (cs !== null) o.docGrid.charSpace = cs;
    }

    var tp = onOff(kid(sectPr, 'titlePg'));
    if (tp !== undefined) o.titlePg = tp;

    return o;
  }

  /* ================================================================== *
   * 四、样式表解析
   * ================================================================== */

  function classifyStyle(name, styleId) {
    var n = (name || '').trim();
    var id = (styleId || '').trim();
    var low = n.toLowerCase().replace(/\s+/g, '');

    /* 标题 N —— 兼容 "标题 1" / "标题1" / "Heading 1" / "heading1" */
    var m = low.match(/^heading(\d)$/);
    if (m) return { type: 'heading', level: +m[1] };
    m = n.replace(/\s+/g, '').match(/^标题(\d)$/);
    if (m) return { type: 'heading', level: +m[1] };
    m = id.toLowerCase().match(/^heading(\d)$/);
    if (m) return { type: 'heading', level: +m[1] };

    /* WPS 中文版 styleId 常常就是 "1".."9" 且样式名为 "标题 N" */
    if (/^[1-9]$/.test(id) && /标题/.test(n)) return { type: 'heading', level: +id };

    if (low === 'normal' || n === '正文' || n === '普通' || low === 'body') return { type: 'normal' };
    if (n === '题注' || low === 'caption') return { type: 'caption' };
    if (n === '页眉' || low === 'header') return { type: 'header' };
    if (n === '页脚' || low === 'footer') return { type: 'footer' };
    if (/^(toc|目录)\s*\d?$/.test(low) || /^目录\d$/.test(n.replace(/\s+/g, ''))) return { type: 'toc' };
    if (low === 'footnote text' || n === '脚注文本') return { type: 'footnote' };
    if (low === 'strong' || n === '强调' && false) return { type: 'character' };
    if (/表格|图表/.test(n) && /文字|内容/.test(n)) return { type: 'tableText' };

    return null;
  }

  /**
   * 解析 styles.xml，返回：
   * {
   *   docDefaults: { rPr, pPr },
   *   byId: { styleId: { id, name, type, basedOn, rPr, pPr, resolvedRPr, resolvedPPr, info } },
   *   order: [styleId...]
   * }
   */
  function parseStylesXml(xmlText) {
    var doc = parseXml(xmlText);
    var rootEl = doc.documentElement;

    var result = {
      docDefaults: { rPr: {}, pPr: {} },
      byId: {},
      order: []
    };

    var dd = kid(rootEl, 'docDefaults');
    if (dd) {
      var rprd = kid(dd, 'rPrDefault');
      if (rprd) result.docDefaults.rPr = parseRPr(kid(rprd, 'rPr'));
      var pprd = kid(dd, 'pPrDefault');
      if (pprd) result.docDefaults.pPr = parsePPr(kid(pprd, 'pPr'));
    }

    var styleEls = rootEl.getElementsByTagNameNS ?
      rootEl.getElementsByTagNameNS(W_NS, 'style') : [];
    if (!styleEls.length) styleEls = rootEl.getElementsByTagName('w:style');

    for (var i = 0; i < styleEls.length; i++) {
      var st = styleEls[i];
      var id = attr(st, 'styleId');
      if (!id) continue;
      var nameEl = kid(st, 'name');
      var rec = {
        id: id,
        name: nameEl ? attr(nameEl, 'val') : id,
        type: attr(st, 'type') || 'paragraph',
        basedOn: attr(kid(st, 'basedOn'), 'val'),
        link: attr(kid(st, 'link'), 'val'),
        next: attr(kid(st, 'next'), 'val'),
        rPr: parseRPr(kid(st, 'rPr')),
        pPr: parsePPr(kid(st, 'pPr')),
        custom: onOff(kid(st, 'customStyle')) === true
      };
      rec.info = classifyStyle(rec.name, rec.id);
      result.byId[id] = rec;
      result.order.push(id);
    }

    /* 解析继承链（basedOn），带环检测 */
    function resolve(id, rPrFmt, pPrFmt, seen) {
      var rec = result.byId[id];
      if (!rec) return;
      if (seen.indexOf(id) >= 0) return;
      seen.push(id);
      if (rec.basedOn && result.byId[rec.basedOn]) resolve(rec.basedOn, rPrFmt, pPrFmt, seen);
      Object.keys(rec.rPr).forEach(function (k) { rPrFmt[k] = rec.rPr[k]; });
      Object.keys(rec.pPr).forEach(function (k) { pPrFmt[k] = rec.pPr[k]; });
    }

    result.order.forEach(function (id) {
      var rPrFmt = {};
      var pPrFmt = {};
      ['rPr'].forEach(function (k) {});
      /* 底层：docDefaults */
      Object.keys(result.docDefaults.rPr).forEach(function (k) { rPrFmt[k] = result.docDefaults.rPr[k]; });
      Object.keys(result.docDefaults.pPr).forEach(function (k) { pPrFmt[k] = result.docDefaults.pPr[k]; });
      resolve(id, rPrFmt, pPrFmt, []);
      var rec = result.byId[id];
      /* 字体等对象型属性做浅合并，避免被子级不完整覆盖 */
      rec.resolvedRPr = rPrFmt;
      rec.resolvedPPr = pPrFmt;
    });

    return result;
  }

  /* ================================================================== *
   * 五、目标文档段落解析
   * ================================================================== */

  function isInsideTable(p) {
    var n = p.parentNode;
    while (n && n.nodeType === 1) {
      if (n.localName === 'tbl') return true;
      if (n.localName === 'body') return false;
      n = n.parentNode;
    }
    return false;
  }

  function paragraphText(p) {
    var t = p.getElementsByTagNameNS ?
      p.getElementsByTagNameNS(W_NS, 't') : p.getElementsByTagName('w:t');
    var s = '';
    for (var i = 0; i < t.length; i++) s += t[i].textContent || '';
    return s;
  }

  /** 段落里是否有域代码（TOC / PAGE 等），这类段落不套正文格式 */
  function paragraphHasField(p, codes) {
    var its = p.getElementsByTagNameNS ?
      p.getElementsByTagNameNS(W_NS, 'instrText') : p.getElementsByTagName('w:instrText');
    if (!its.length) return false;
    var buf = '';
    for (var i = 0; i < its.length; i++) buf += (its[i].textContent || '') + ' ';
    buf = buf.toUpperCase();
    if (!codes) return true;
    for (var j = 0; j < codes.length; j++) if (buf.indexOf(codes[j]) >= 0) return true;
    return false;
  }

  /** 段落里所有 run 是否都只含域字符 / 书签（无真实文本） */
  function runHasRealText(r) {
    var ts = r.getElementsByTagNameNS ?
      r.getElementsByTagNameNS(W_NS, 't') : r.getElementsByTagName('w:t');
    for (var i = 0; i < ts.length; i++) if ((ts[i].textContent || '').length) return true;
    return false;
  }

  /* ================================================================== *
   * 六、标题层级启发式识别（针对没有使用样式的文档）
   * ================================================================== */

  /* 层级约定（与《课程设计报告模板》一致）：
   *   第X章 / 第X篇 / 第X部   → 一级标题
   *   X.Y                     → 二级标题
   *   X.Y.Z（更深如 X.Y.Z.W） → 三级标题
   *   第X节                   → 二级标题
   *   一、二、三、             → 二级标题
   *   （一）/（1）             → 三级标题
   *   1. / 1、                 → 三级标题（序号型列表项）
   *   摘要 / 目录 / 参考文献 … → 一级标题
   *
   * ⚠️ 多级数字编号必须「先深后浅」匹配：早先按顺序遍历正则时，
   *    ^数字+[.、] 这条浅规则会先把 "1.1.1" 抢走，导致三级标题被判成二级。
   *    现在改成先解析出整个编号串、再按点分组数定级。
   */

  var ZH_NUM = '[一二三四五六七八九十百千零〇]';
  var AR_NUM = '[0-9０-９]';
  var NUM_ANY = '(?:' + ZH_NUM + '|' + AR_NUM + ')+';
  var DOT_CHARS = '.．';                 /* 半角点 + 全角句点 */
  var DOT = '[' + DOT_CHARS + ']';
  var RE_MULTI_NUM = new RegExp('^(' + AR_NUM + '{1,2}(?:\\s*' + DOT + '\\s*' + AR_NUM + '{1,2})+)');
  var RE_MORE_NUM = new RegExp('^\\s*' + DOT + '\\s*' + AR_NUM);

  /** 解析多级阿拉伯编号（1.1 / 1.1.1 …），返回层级；不是多级编号则返回 0 */
  function numberedLevel(t) {
    var m = t.match(RE_MULTI_NUM);
    if (!m) return 0;
    /* 后面若还接着「.数字」，说明整个编号更长，交给下一轮匹配 */
    if (RE_MORE_NUM.test(t.slice(m[0].length))) return 0;
    var groups = m[1].split(new RegExp(DOT)).length;
    return groups >= 3 ? 3 : 2;   /* 三组及以上统一收敛到三级 */
  }

  /** 返回 { level, strong }；strong = 模式足够可靠，无需再做「像不像标题」的校验 */
  function heuristicInfo(t) {
    /* 第X章 / 第X节 */
    if (new RegExp('^第\\s*' + NUM_ANY + '\\s*[章篇部]').test(t)) return { level: 1, strong: true };
    if (new RegExp('^第\\s*' + NUM_ANY + '\\s*节').test(t)) return { level: 2, strong: true };

    /* X.Y → 二级；X.Y.Z → 三级 */
    var nv = numberedLevel(t);
    if (nv) return { level: nv, strong: true };

    /* 一、二、三、 */
    if (new RegExp('^' + ZH_NUM + '+\\s*[、' + DOT_CHARS + ']\\s*\\S').test(t)) return { level: 2, strong: false };

    /* （一）/（1） */
    if (new RegExp('^[（(]\\s*' + NUM_ANY + '\\s*[)）]').test(t)) return { level: 3, strong: false };

    /* 1. / 1、 —— 序号型列表项（限定 1~2 位数字，避免把「2023.5 版本」当标题） */
    if (new RegExp('^' + AR_NUM + '{1,2}\\s*[、' + DOT_CHARS + ']\\s*\\S').test(t)) return { level: 3, strong: false };

    /* 无编号的常见章节名 */
    if (/^(摘\s*要|abstract|引\s*言|绪\s*论|前\s*言|结\s*论|结束语|参考文献|致\s*谢|目\s*录|课程设计(概述|任务|要求))\s*$/i.test(t)) {
      return { level: 1, strong: true };
    }
    if (/^附\s*[录錄]\s*[A-Za-z0-9一二三四五六七八九十]?\s*$/.test(t)) return { level: 1, strong: true };

    return null;
  }

  function heuristicLevel(text) {
    if (!text) return 0;
    var t = text.trim();
    if (!t || t.length > 60) return 0;
    /* 含句末标点的基本是正文 */
    if (/[。；！？]$/.test(t) && t.length > 18) return 0;
    var r = heuristicInfo(t);
    if (!r) return 0;
    /* 弱规则要求这段文字"看起来像标题"：较短、不含中文句号 */
    if (!r.strong) {
      if (t.length > 30) return 0;
      if (/[。；]/.test(t)) return 0;
    }
    return r.level;
  }

  /* 图表题注兜底：形如「表4-1 …」「图 2-1 …」「图A-1 …」「表1 …」
   * 样式里没写「题注」时靠这个认，否则表题会被套成正文格式 */
  var RE_CAP_NUM = AR_NUM + '{1,2}(?:\\s*[-–—−]\\s*' + AR_NUM + '+)?';
  var RE_CAPTION = new RegExp('^[图表]\\s*[A-Za-z]?\\s*' + RE_CAP_NUM +
    '(?!\\s*' + DOT + '\\s*' + AR_NUM + ')');
  /* 编号后紧跟这些字的多半是正文引用（「图3 所示…」），不是题注。
     注意别把「系统」「中间继电器」这类正常图名误伤——这里只收真正表示"话没说完"的词头 */
  var RE_CAPTION_TAIL = /^(?:如|见|所|示|说|给|列|即|分别是|表示|说明|给出|列出)/;

  function isCaptionText(text) {
    if (!text) return false;
    var t = text.trim();
    if (!t || t.length > 60) return false;
    if (/[。；]$/.test(t)) return false;
    var m = t.match(RE_CAPTION);
    if (!m) return false;
    var rest = t.slice(m[0].length).trim();
    if (!rest) return false;                 /* 光一个「图1」不算题注 */
    if (RE_CAPTION_TAIL.test(rest)) return false;
    return true;
  }

  /* ================================================================== *
   * 七、规范对象构建
   * ================================================================== */

  function emptySpec() {
    return {
      source: { file: '', generatedAt: '' },
      page: null,
      margin: null,
      docGrid: null,
      normal: null,
      headings: {},       /* level -> style spec */
      caption: null,
      tableText: null,
      toc: null,
      foundNames: [],
      missing: []
    };
  }

  function styleSpecOf(rec) {
    if (!rec) return null;
    return {
      styleId: rec.id,
      name: rec.name,
      rPr: rec.resolvedRPr || rec.rPr || {},
      pPr: rec.resolvedPPr || rec.pPr || {}
    };
  }

  /* ================================================================== *
   * 八、核心 API — 分析模板
   * ================================================================== */

  function getZip() {
    var Z = (typeof JSZip !== 'undefined') ? JSZip : (typeof globalThis !== 'undefined' ? globalThis.JSZip : null);
    if (!Z) throw new Error('缺少 JSZip 依赖');
    return Z;
  }

  async function readZipEntry(zip, path) {
    var f = zip.file(path);
    if (!f) return null;
    return await f.async('string');
  }

  async function resolvePath(zip, path) {
    if (zip.file(path)) return path;
    return null;
  }

  async function analyzeTemplate(buffer, options) {
    options = options || {};
    var JSZipRef = getZip();
    var zip;
    try {
      zip = await JSZipRef.loadAsync(buffer);
    } catch (e) {
      throw new Error('文件无法作为 DOCX 解析，请确认是 .docx 格式（不是 .doc / .wps）');
    }

    var docPath = await resolvePath(zip, DOC_PATH);
    if (!docPath) throw new Error('DOCX 内部缺少 word/document.xml，文件可能已损坏');

    var docXml = await readZipEntry(zip, docPath);
    var styXml = await readZipEntry(zip, STY_PATH);

    var spec = emptySpec();
    spec.source.file = options.fileName || '';

    /* --- 页面设置：取 body 末尾的 sectPr --- */
    var docDom = parseXml(docXml);
    var body = docDom.documentElement.getElementsByTagNameNS ?
      docDom.documentElement.getElementsByTagNameNS(W_NS, 'body')[0] : null;
    if (!body) body = docDom.getElementsByTagName('w:body')[0];

    var sect = null;
    if (body) {
      var direct = kids(body, 'sectPr');
      if (direct.length) sect = direct[direct.length - 1];
    }
    if (!sect) {
      var all = docDom.getElementsByTagNameNS ?
        docDom.getElementsByTagNameNS(W_NS, 'sectPr') : docDom.getElementsByTagName('w:sectPr');
      if (all && all.length) sect = all[all.length - 1];
    }
    var sectSpec = parseSectPr(sect);
    spec.page = sectSpec.page || null;
    spec.margin = sectSpec.margin || null;
    spec.docGrid = sectSpec.docGrid || null;

    /* --- 样式 --- */
    if (styXml) {
      var styles = parseStylesXml(styXml);

      var headings = {};
      var normal = null, caption = null, toc = null, tableText = null;

      styles.order.forEach(function (id) {
        var rec = styles.byId[id];
        if (!rec || !rec.info) return;
        if (rec.type === 'character') return;
        if (rec.info.type === 'heading') {
          if (!headings[rec.info.level]) headings[rec.info.level] = styleSpecOf(rec);
        } else if (rec.info.type === 'normal' && !normal) {
          normal = styleSpecOf(rec);
        } else if (rec.info.type === 'caption' && !caption) {
          caption = styleSpecOf(rec);
        } else if (rec.info.type === 'toc' && !toc) {
          toc = styleSpecOf(rec);
        } else if (rec.info.type === 'tableText' && !tableText) {
          tableText = styleSpecOf(rec);
        }
      });

      spec.headings = headings;
      spec.normal = normal;
      spec.caption = caption;
      spec.toc = toc;
      spec.tableText = tableText;
      spec.foundNames = Object.keys(headings).sort().map(function (lv) {
        return '标题 ' + lv + '（' + headings[lv].name + '）';
      });
      if (normal) spec.foundNames.unshift('正文（' + normal.name + '）');
      if (caption) spec.foundNames.push('题注（' + caption.name + '）');

      for (var lv = 1; lv <= 3; lv++) if (!headings[lv]) spec.missing.push('标题 ' + lv);
      if (!normal) spec.missing.push('正文');
    }

    /* --- 兜底：模板里没定义正文样式时，用 docDefaults --- */
    if (!spec.normal) {
      spec.normal = { styleId: null, name: '（文档默认）', rPr: {}, pPr: {} };
      spec.autoNormal = true;
    }

    spec.source.generatedAt = new Date().toISOString();
    return spec;
  }

  /* ================================================================== *
   * 九、XML 写入工具（保持 OOXML 元素顺序）
   * ================================================================== */

  function insertOrdered(parent, el, orderList) {
    var idx = orderList.indexOf(el.localName);
    if (idx < 0) { parent.appendChild(el); return el; }
    var cs = parent.childNodes;
    for (var i = 0; i < cs.length; i++) {
      var c = cs[i];
      if (c.nodeType !== 1) continue;
      var ci = orderList.indexOf(c.localName);
      if (ci >= 0 && ci > idx) { parent.insertBefore(el, c); return el; }
    }
    parent.appendChild(el);
    return el;
  }

  function removeChild(parent, localName) {
    if (!parent) return;
    var cs = parent.childNodes;
    for (var i = cs.length - 1; i >= 0; i--) {
      if (cs[i].nodeType === 1 && cs[i].localName === localName) parent.removeChild(cs[i]);
    }
  }

  /** 删除属性（兼容 xmldom 不支持 removeAttributeNS 的情况） */
  function dropAttr(el, name) {
    if (!el || !el.removeAttribute) return;
    try { el.removeAttributeNS(W_NS, name); } catch (e) { /* ignore */ }
    el.removeAttribute('w:' + name);
  }

  /* ================================================================== *
   * 十、核心 API — 套用格式
   * ================================================================== */

  /**
   * options:
   *   applyPage        是否套用页面设置（默认 true）
   *   applyHeading     是否套用标题格式（默认 true）
   *   applyBody        是否套用正文格式（默认 true）
   *   applyCaption     是否套用题注格式（默认 true）
   *   skipTableIndent  表格内段落不加首行缩进（默认 true）
   *   heuristic        智能识别无样式标题（默认 true）
   *   clearManual      覆盖手工格式（默认 true）
   *   syncStyles       同步 styles.xml 并纠正 pStyle 指向（默认 true）
   */

  /* ================================================================== *
   * 六·五、目标文档样式同步
   *
   *   只写「直接格式」的坏处：段落明明被排成一级标题（小二黑体居中、
   *   outlineLvl=0），pStyle 却还指着源文档的「标题 2」——Word 的样式库、
   *   样式检查器都会显示错的名字，用户打开一看就以为没排对。
   *   所以这里把 spec 一并写进 styles.xml，并把每个段落的 pStyle
   *   改挂到与判定结果一致的样式上。
   * ================================================================== */

  function normStyleName(s) {
    return String(s || '').trim().toLowerCase().replace(/\s+/g, ' ');
  }

  /** 各级标题在 Word / WPS 里的样式名候选（中英双语） */
  function headingNameCandidates(level) {
    return ['标题 ' + level, '标题' + level, 'heading ' + level, 'heading' + level];
  }
  var NORMAL_NAME_CANDIDATES = ['正文', 'Normal', 'Body Text', 'body text', 'Normal (Web)'];
  var CAPTION_NAME_CANDIDATES = ['题注', 'Caption', 'caption'];

  function paraStylesOf(stylesRoot) {
    var all = stylesRoot.getElementsByTagNameNS ?
      stylesRoot.getElementsByTagNameNS(W_NS, 'style') : stylesRoot.getElementsByTagName('w:style');
    var out = [];
    for (var i = 0; i < all.length; i++) {
      var t = attr(all[i], 'type');
      if (!t || t === 'paragraph') out.push(all[i]);
    }
    return out;
  }

  /** 按名称候选顺序找段落样式（先按第一个候选全表扫，再找第二个…） */
  function findParaStyleByName(stylesRoot, candidates) {
    var styles = paraStylesOf(stylesRoot);
    for (var ci = 0; ci < candidates.length; ci++) {
      var want = normStyleName(candidates[ci]);
      for (var i = 0; i < styles.length; i++) {
        if (normStyleName(attr(kid(styles[i], 'name'), 'val')) === want) return styles[i];
      }
    }
    return null;
  }

  function makeParaStyleEl(doc, id, name, basedOnId) {
    var st = createEl(doc, 'style');
    setAttr(st, 'type', 'paragraph');
    setAttr(st, 'styleId', id);
    var nm = createEl(doc, 'name'); setAttr(nm, 'val', name);
    insertOrdered(st, nm, STYLE_ORDER);
    if (basedOnId) {
      var bo = createEl(doc, 'basedOn'); setAttr(bo, 'val', basedOnId);
      insertOrdered(st, bo, STYLE_ORDER);
    }
    var ui = createEl(doc, 'uiPriority'); setAttr(ui, 'val', '99');
    insertOrdered(st, ui, STYLE_ORDER);
    insertOrdered(st, createEl(doc, 'qFormat'), STYLE_ORDER);
    return st;
  }

  /** 把 spec 的段落属性写进「样式定义」里的 pPr */
  function writeStylePPrInto(doc, pPr, targetPPr, outlineLvl, zeroIndent) {
    targetPPr = targetPPr || {};

    if (targetPPr.align) {
      var jc = createEl(doc, 'jc'); setAttr(jc, 'val', targetPPr.align);
      insertOrdered(pPr, jc, PPR_ORDER);
    }
    if (targetPPr.spacing) {
      var sp = createEl(doc, 'spacing');
      var ts = targetPPr.spacing;
      ['before', 'after', 'line', 'lineRule', 'beforeLines', 'afterLines'].forEach(function (k) {
        if (ts[k] !== undefined && ts[k] !== null) setAttr(sp, k, ts[k]);
      });
      insertOrdered(pPr, sp, PPR_ORDER);
    }
    /* 标题 / 题注样式通常 basedOn 正文样式，而模板的正文带「首行缩进 2 字符」。
       这里必须写死 0 —— 否则"没设首行缩进"的标题会顺着样式链继承出两格缩进。 */
    var indSrc = zeroIndent ? ZERO_IND : (targetPPr.ind || {});
    if (Object.keys(indSrc).length) {
      var ind = createEl(doc, 'ind');
      Object.keys(indSrc).forEach(function (k) { setAttr(ind, k, indSrc[k]); });
      insertOrdered(pPr, ind, PPR_ORDER);
    }
    if (outlineLvl !== undefined && outlineLvl !== null) {
      var ol = createEl(doc, 'outlineLvl'); setAttr(ol, 'val', outlineLvl);
      insertOrdered(pPr, ol, PPR_ORDER);
    }
  }

  /** 重写某个样式的格式定义（保留 styleId / 名称 / basedOn） */
  function rewriteStyleDef(doc, styleEl, specStyle, outlineLvl, zeroIndent) {
    removeChild(styleEl, 'pPr');
    removeChild(styleEl, 'rPr');

    var pPr = createEl(doc, 'pPr');
    writeStylePPrInto(doc, pPr, specStyle.pPr || {}, outlineLvl, zeroIndent);
    insertOrdered(styleEl, pPr, STYLE_ORDER);

    var rPr = createEl(doc, 'rPr');
    writeRPrInto(doc, rPr, specStyle.rPr || {}, { clearManual: true });
    insertOrdered(styleEl, rPr, STYLE_ORDER);
  }

  /**
   * 让 styles.xml 与 spec 对齐，返回每类段落该挂的 styleId。
   * @returns {{normal, headings:{}, caption, created:[], reused:[], missing:[]}}
   */
  function syncSpecStyles(stylesDom, spec, reports) {
    var doc = stylesDom;
    var root = doc.documentElement;
    var out = { normal: null, headings: {}, caption: null, created: [], reused: [] };

    var used = {};
    var existing = paraStylesOf(root);
    for (var i = 0; i < existing.length; i++) {
      var id = attr(existing[i], 'styleId');
      if (id) used[id] = true;
    }
    function freshId(base) {
      var n = 1, cand = base;
      while (used[cand]) { n++; cand = base + n; }
      used[cand] = true;
      return cand;
    }
    function take(candidates, baseId, preferredName, specStyle, outlineLvl, label, zeroIndent) {
      var el = findParaStyleByName(root, candidates);
      if (el) {
        out.reused.push(label + ' → ' + attr(el, 'styleId') + '(' + attr(kid(el, 'name'), 'val') + ')');
      } else {
        var nid = freshId(baseId);
        var name = specStyle && specStyle.name ? specStyle.name : preferredName;
        el = makeParaStyleEl(doc, nid, name, out.normal);
        root.appendChild(el);
        out.created.push(label + ' → ' + nid + '(' + name + ')');
      }
      if (specStyle) rewriteStyleDef(doc, el, specStyle, outlineLvl, zeroIndent);
      return attr(el, 'styleId');
    }

    /* 先正文：标题样式要 basedOn 它 */
    out.normal = take(NORMAL_NAME_CANDIDATES, 'WBBody', '正文', spec.normal, null, '正文', false);

    if (spec.headings) {
      Object.keys(spec.headings).forEach(function (k) {
        var lv = parseInt(k, 10);
        if (!lv) return;
        out.headings[lv] = take(headingNameCandidates(lv), 'WBHeading' + lv,
          '标题 ' + lv, spec.headings[k], lv - 1, '标题 ' + lv, true);
      });
    }
    if (spec.caption) {
      out.caption = take(CAPTION_NAME_CANDIDATES, 'WBCaption', '题注', spec.caption, null, '题注', true);
    }

    if (reports) {
      reports.stylesCreated = out.created;
      reports.stylesReused = out.reused;
    }
    return out;
  }

  /* ================================================================== *
   * 六·六、清掉 HTML / markdown 遗留的水平线
   *
   *   从网页、AI 对话里粘进 Word 的 markdown「---」分隔线，会被 Word 存成
   *     <w:pict><v:rect o:hr="t" o:hrstd="t" w10:wrap type="none"/></w:pict>
   *   `o:hr` 是 Word 专门给「水平线」打的标记（hr = horizontal rule），
   *   不会出现在正常的图形 / 图片上。这类线不属于任何模板规范，跟着走就会
   *   在章与章之间留下一道灰色横杠——用户很容易把它当成"页眉下面那条线"。
   * ================================================================== */

  function outerXml(el) {
    try {
      return new XMLSerializer().serializeToString(el);
    } catch (e) {
      return '';
    }
  }

  /** 找出所有「水平线」pict 元素 */
  function collectHorizontalRules(docDom) {
    var out = [];
    var pics = docDom.getElementsByTagNameNS ?
      docDom.getElementsByTagNameNS(W_NS, 'pict') : docDom.getElementsByTagName('w:pict');
    for (var i = 0; i < pics.length; i++) {
      var s = outerXml(pics[i]);
      /* 新版 Word / WPS：<v:rect o:hr="t">；老版：<v:hr> */
      if (/\bo:hr\s*=\s*"(1|t|true|on)"/i.test(s) || /<v:hr[\s/>]/i.test(s)) out.push(pics[i]);
    }
    return out;
  }

  /**
   * 删除水平线。如果某个段落只装了这一条线，整段一起删掉
   * （否则会空出一整行，看起来像多了个空行）。
   */
  function removeHorizontalRules(docDom) {
    var pics = collectHorizontalRules(docDom);
    var removed = 0, droppedParas = 0;

    for (var i = 0; i < pics.length; i++) {
      var pict = pics[i];
      var run = pict.parentNode;
      var para = run;
      while (para && para.localName !== 'p') para = para.parentNode;
      if (!para) continue;

      /* 段落里除了这一条线，还有没有别的东西 */
      var others = 0;
      var cs = para.childNodes;
      for (var j = 0; j < cs.length; j++) {
        var c = cs[j];
        if (c.nodeType !== 1 || c.localName === 'pPr') continue;
        if (c !== run) others++;
      }

      if (others === 0 && run && run.localName === 'r') {
        para.parentNode.removeChild(para);
        droppedParas++;
      } else if (run && run.localName === 'r') {
        run.parentNode.removeChild(run);
      } else {
        pict.parentNode.removeChild(pict);
      }
      removed++;
    }
    return { removed: removed, droppedParas: droppedParas };
  }

  /* ================================================================== *
   * 五·五、关闭文档网格
   *
   *   为什么值得单独做一件事：
   *     sectPr 里的 <w:docGrid w:type="lines"> 会开启「指定行网格」。开了之后，
   *     凡是「倍数行距」的段落（标题基本都是单倍行距），实际行高会被向上吸附到
   *     网格间距（w:linePitch）的整数倍。中文文档常见 linePitch=312 twips=15.6 磅，
   *     于是 18 磅的一级标题会被从约 23 磅撑到约 31 磅 —— 标题越发虚、页面越发松。
   *
   *   为什么「正文看着没事、只有标题遭殃」：
   *     正文多用「固定值行距」（w:lineRule="exact"），固定行距不参与网格吸附。
   *     所以同一份文档里，正文行高正常、标题却被撑开，用户会描述成
   *     「明明设置一样，为什么它更紧凑」。
   *
   *   实现要点：
   *     ① 把 w:type 属性整个删掉，而不是写成 type="default" —— 「无网格」正是
   *        属性的缺省状态，删掉才和天然无网格的文档字节一致。
   *     ② 只处理确实开了网格的（有 w:type 且不是 default），本来就无网格的不动，
   *        免得往用户文档里写无意义的改动。
   *     ③ 节标题（header/footer）里的 sectPr 一起处理，否则页眉页脚内部仍按网格排。
   * ================================================================== */

  function disableDocGrid(docDom) {
    var list = docDom.getElementsByTagNameNS ?
      docDom.getElementsByTagNameNS(W_NS, 'sectPr') : docDom.getElementsByTagName('w:sectPr');

    var changed = 0, sections = 0, modeBefore = '';

    for (var i = 0; i < list.length; i++) {
      var sectPr = list[i];
      sections++;
      var dg = kid(sectPr, 'docGrid');
      if (!dg) continue;

      var t = attr(dg, 'type');
      if (t === null || t === '' || t === 'default') continue;   /* 本来就是无网格 */

      if (modeBefore) modeBefore += ' / ';
      modeBefore += t;

      dropAttr(dg, 'type');
      changed++;
    }

    return { changed: changed, sections: sections, modeBefore: modeBefore };
  }

  /* ================================================================== *
   * 六·七、一级标题（章）的附加版式
   *
   *   课设这类模板常额外要求：
   *     ① 每章另起一页（一级标题不允许出现在页面下半部分）
   *     ② 一级标题上下各空一行，而且这两个空行也要套「一级标题」格式
   *
   *   为什么空行要套标题格式：空行的高度由「段落标记」的字符格式决定。如果
   *   按正文（小四、固定值 20 磅）算，空行高度和小二号标题对不上，空出来的
   *   那行明显偏小；按一级标题格式算，空行的行高才和标题一致。
   *
   *   分页符挂在「最上面那个空行」上：这样空行和标题一定一起落在新页顶部，
   *   不会被甩到上一页的页尾去（挂在标题上则空行可能留在上一页，时隐时现）。
   * ================================================================== */

  function paragraphsOf(doc) {
    var list = doc.getElementsByTagNameNS ?
      doc.getElementsByTagNameNS(W_NS, 'p') : doc.getElementsByTagName('w:p');
    var out = [];
    for (var i = 0; i < list.length; i++) out.push(list[i]);
    return out;
  }

  /** 取上一个元素兄弟 */
  function prevElement(node) {
    var prev = node && node.previousSibling;
    while (prev && prev.nodeType !== 1) prev = prev.previousSibling;
    return prev || null;
  }

  /** 段落里是否还塞着图形 / 图片 / 表格（这种「看着空」的段落不能当空行复用） */
  function hasBlockContent(el) {
    var cs = el.childNodes;
    for (var i = 0; i < cs.length; i++) {
      var c = cs[i];
      if (c.nodeType !== 1) continue;
      var nm = c.localName;
      if (nm === 'drawing' || nm === 'pict' || nm === 'object' || nm === 'tbl' ||
        nm === 'txbxContent' || nm === 'pict' || nm === 'AlternateContent') return true;
      if (hasBlockContent(c)) return true;
    }
    return false;
  }

  /** 是不是一个真正意义上的「空段落」（没有任何可见内容） */
  function isBlankParagraphNode(p) {
    if (!p || p.nodeType !== 1) return null;
    if (p.localName !== 'p') return null;
    if (hasBlockContent(p)) return null;
    if (paragraphText(p).replace(/[\s\u00a0\u3000]/g, '').length) return null;
    return p;
  }

  /** 把一段段落属性写成「一级标题格式的空行」（含分页标记、零缩进、段落标记字符格式） */
  function fillChapterBlankPPr(doc, pPr, headingSpec, styleId, pageBreak) {
    if (styleId) {
      var ps = createEl(doc, 'pStyle'); setAttr(ps, 'val', styleId);
      insertOrdered(pPr, ps, PPR_ORDER);
    }
    if (pageBreak) insertOrdered(pPr, createEl(doc, 'pageBreakBefore'), PPR_ORDER);

    var targetPPr = headingSpec.pPr || {};
    if (targetPPr.spacing) {
      var sp = createEl(doc, 'spacing');
      ['before', 'after', 'line', 'lineRule'].forEach(function (k) {
        if (targetPPr.spacing[k] !== undefined && targetPPr.spacing[k] !== null) {
          setAttr(sp, k, targetPPr.spacing[k]);
        }
      });
      insertOrdered(pPr, sp, PPR_ORDER);
    }
    if (targetPPr.align) {
      var jc = createEl(doc, 'jc'); setAttr(jc, 'val', targetPPr.align);
      insertOrdered(pPr, jc, PPR_ORDER);
    }
    /* 空行同样不能有首行缩进 */
    var ind = createEl(doc, 'ind');
    Object.keys(ZERO_IND).forEach(function (k) { setAttr(ind, k, ZERO_IND[k]); });
    insertOrdered(pPr, ind, PPR_ORDER);

    /* 段落标记的字符格式决定这一空行的高度 —— 必须和一级标题一致 */
    var markRPr = createEl(doc, 'rPr');
    writeRPrInto(doc, markRPr, headingSpec.rPr || {}, { clearManual: true });
    insertOrdered(pPr, markRPr, PPR_ORDER);
  }

  /** 造一个「一级标题格式的空段落」 */
  function makeChapterBlank(doc, headingSpec, styleId, pageBreak) {
    var p = createEl(doc, 'p');
    var pPr = createEl(doc, 'pPr');
    p.appendChild(pPr);
    fillChapterBlankPPr(doc, pPr, headingSpec, styleId, pageBreak);

    var r = createEl(doc, 'r');
    var rPr = createEl(doc, 'rPr');
    writeRPrInto(doc, rPr, headingSpec.rPr || {}, { clearManual: true });
    r.appendChild(rPr);
    p.appendChild(r);
    return p;
  }

  /**
   * 把「已经存在的空段落」改造成一级标题格式的空行（复用，不新增段落）。
   * 场景：原文档在标题上下本来就留了空行，若再插一行就会变成两行。
   */
  function dressChapterBlank(doc, p, headingSpec, styleId, pageBreak) {
    var pPr = kid(p, 'pPr');
    if (!pPr) { pPr = createEl(doc, 'pPr'); p.insertBefore(pPr, p.firstChild); }
    /* 清掉段落内容（空行不该留 run） */
    var ch = p.firstChild;
    while (ch) {
      var nx = ch.nextSibling;
      if (ch !== pPr) p.removeChild(ch);
      ch = nx;
    }
    /* 段落属性整段重写，避免残留上一轮的间距 / 缩进 */
    var pc = pPr.firstChild;
    while (pc) {
      var nx2 = pc.nextSibling;
      pPr.removeChild(pc);
      pc = nx2;
    }
    fillChapterBlankPPr(doc, pPr, headingSpec, styleId, pageBreak);

    var r = createEl(doc, 'r');
    var rPr = createEl(doc, 'rPr');
    writeRPrInto(doc, rPr, headingSpec.rPr || {}, { clearManual: true });
    r.appendChild(rPr);
    p.appendChild(r);
    return p;
  }

  function setPageBreakBefore(doc, p) {
    var pPr = kid(p, 'pPr');
    if (!pPr) { pPr = createEl(doc, 'pPr'); p.insertBefore(pPr, p.firstChild); }
    if (kid(pPr, 'pageBreakBefore')) return;
    insertOrdered(pPr, createEl(doc, 'pageBreakBefore'), PPR_ORDER);
  }

  /**
   * 给每个一级标题前后补空行、并在章首加分页。
   * 靠「段落里已写入的 outlineLvl=0 + pStyle 指向标题 1 样式」来认定一级标题，
   * 这样不会把正文里残留的 outlineLvl 误当章标题。
   *
   * 标题上下本来就留了空行的话，就地把它改造成标题格式的空行，不再叠加第二行。
   */
  function applyChapterLayout(doc, spec, styleIds, cfg) {
    var before = cfg.blankLinesBefore === undefined ? 0 : cfg.blankLinesBefore;
    var after = cfg.blankLinesAfter === undefined ? 0 : cfg.blankLinesAfter;
    var wantBreak = !!cfg.pageBreakBefore;
    var res = { breaks: 0, blanks: 0 };
    if (!before && !after && !wantBreak) return res;

    var headingSpec = (spec.headings && spec.headings[1]) || spec.normal;
    var styleId = (styleIds && styleIds.headings) ? styleIds.headings[1] : null;
    var touched = [];   /* 本轮创建或改造过的空行 —— 下一章可直接复用，不叠行 */
    var seenContent = false;

    paragraphsOf(doc).forEach(function (p) {
      var pPr = kid(p, 'pPr');
      var ol = num(attr(kid(pPr, 'outlineLvl'), 'val'));
      var blank = !paragraphText(p).replace(/[\s\u00a0\u3000]/g, '').length;

      var psId = attr(kid(pPr, 'pStyle'), 'val');
      var looksLikeChapter = (ol === 0) && !isInsideTable(p) && !blank &&
        (!styleIds ? true : (psId === styleId));

      if (!looksLikeChapter) {
        if (!blank) seenContent = true;
        return;
      }

      /* 章首才分页：文档最开头的那一章前面没内容，再分页就会空出一整页 */
      var needBreak = wantBreak && seenContent;

      /* ① 前空行 */
      if (before > 0) {
        var prev = prevElement(p);
        var reused = 0;
        if (prev && (touched.indexOf(prev) >= 0 || isBlankParagraphNode(prev))) {
          /* 上一章的后空行 / 原文自带的空行 —— 直接复用 */
          dressChapterBlank(doc, prev, headingSpec, styleId, needBreak);
          if (touched.indexOf(prev) < 0) touched.push(prev);
          reused = 1;
          res.blanks++;
        }
        for (var bi = reused; bi < before; bi++) {
          var nb = makeChapterBlank(doc, headingSpec, styleId, needBreak && bi === 0);
          p.parentNode.insertBefore(nb, p);
          touched.push(nb);
          res.blanks++;
        }
        if (needBreak) res.breaks++;
      } else if (needBreak) {
        setPageBreakBefore(doc, p);
        res.breaks++;
      }

      /* ② 后空行 */
      if (after > 0) {
        var next = p.nextSibling;
        while (next && next.nodeType !== 1) next = next.nextSibling;
        if (next && isBlankParagraphNode(next) && touched.indexOf(next) < 0) {
          dressChapterBlank(doc, next, headingSpec, styleId, false);
          touched.push(next);
          res.blanks++;
        } else {
          for (var ai = 0; ai < after; ai++) {
            var na = makeChapterBlank(doc, headingSpec, styleId, false);
            if (p.nextSibling) p.parentNode.insertBefore(na, p.nextSibling);
            else p.parentNode.appendChild(na);
            touched.push(na);
            res.blanks++;
          }
        }
      }
      seenContent = true;
    });

    return res;
  }

  async function applyFormat(buffer, spec, options) {
    options = Object.assign({
      applyPage: true,
      applyHeading: true,
      applyBody: true,
      applyCaption: true,
      skipTableIndent: true,
      heuristic: true,
      clearManual: true,
      syncStyles: true,
      removeRules: true,
      disableGrid: true
    }, options || {});

    var JSZipRef = getZip();
    var zip;
    try {
      zip = await JSZipRef.loadAsync(buffer);
    } catch (e) {
      throw new Error('文件无法作为 DOCX 解析，请确认是 .docx 格式（不是 .doc / .wps）');
    }

    if (!zip.file(DOC_PATH)) throw new Error('DOCX 内部缺少 word/document.xml，文件可能已损坏');

    var docXml = await readZipEntry(zip, DOC_PATH);
    var styXml = await readZipEntry(zip, STY_PATH);

    /* 目标文档的样式索引：styleId -> 分类信息 */
    var targetStyles = { byId: {} };
    if (styXml) {
      var ts = parseStylesXml(styXml);
      ts.order.forEach(function (id) {
        var rec = ts.byId[id];
        targetStyles.byId[id] = {
          id: id,
          name: rec.name,
          type: rec.type,
          info: rec.info,
          outlineLvl: (rec.resolvedPPr && rec.resolvedPPr.outlineLvl !== undefined) ?
            rec.resolvedPPr.outlineLvl : (rec.pPr ? rec.pPr.outlineLvl : undefined)
        };
      });
    }

    var docDom = parseXml(docXml);
    var reports = {
      paragraphs: 0,
      headings: { 1: 0, 2: 0, 3: 0, 4: 0, 5: 0, 6: 0, 7: 0, 8: 0, 9: 0 },
      body: 0,
      caption: 0,
      tableParagraphs: 0,
      skippedFields: 0,
      heuristicHits: 0,
      blankParagraphs: 0,
      demotedHeadings: 0,
      levelFixes: 0,
      stylesCreated: [],
      stylesReused: [],
      removedRules: 0,
      docGridOff: 0,
      docGridBefore: '',
      chapterBreaks: 0,
      chapterBlanks: 0,
      pageApplied: false,
      warnings: []
    };

    /* ---------- 0. 样式定义同步（决定每个段落该挂哪个 pStyle） ---------- */
    var styleIds = null;
    if (options.syncStyles) {
      if (styXml) {
        var styDom = parseXml(styXml);
        styleIds = syncSpecStyles(styDom, spec, reports);
      } else {
        reports.warnings.push('目标文档没有 word/styles.xml，已跳过样式同步（仅写入直接格式）');
      }
    }

    /* ---------- 1.5 清掉 HTML / markdown 遗留的水平线 ----------
       （必须在统计段落之前做，删掉的段落不应计入 headings / body） */
    if (options.removeRules) {
      var ruleRes = removeHorizontalRules(docDom);
      reports.removedRules = ruleRes.removed;
    }

    /* ---------- 1.4 关闭文档网格 ----------
       必须放在套用页面设置之前：套页面时可能按规范写回 docGrid，
       先关掉才不会被重新打开。 */
    if (options.disableGrid) {
      var gridRes = disableDocGrid(docDom);
      reports.docGridOff = gridRes.changed;
      reports.docGridBefore = gridRes.modeBefore;
    }

    /* ---------- 1. 页面设置 ---------- */
    if (options.applyPage && (spec.page || spec.margin || spec.docGrid)) {
      var allSectPr = docDom.getElementsByTagNameNS ?
        docDom.getElementsByTagNameNS(W_NS, 'sectPr') : docDom.getElementsByTagName('w:sectPr');
      if (!allSectPr.length) {
        reports.warnings.push('目标文档未找到节属性（sectPr），已跳过页面设置');
      } else {
        for (var si = 0; si < allSectPr.length; si++) {
          applySectionProperties(docDom, allSectPr[si], spec);
        }
        reports.pageApplied = true;
        reports.sections = allSectPr.length;
      }
    }

    /* ---------- 2. 段落格式 ----------
       paragraphsOf() 会拷贝成数组：浏览器返回的 NodeList 是「活的」，
       后面往文档里插空段落时，活列表会变长，遍历就会失控。 */
    var allP = paragraphsOf(docDom);

    var bodyLevels = {};
    for (var lv = 1; lv <= 9; lv++) {
      if (spec.headings && spec.headings[lv]) bodyLevels[lv] = true;
    }

    for (var pi = 0; pi < allP.length; pi++) {
      var p = allP[pi];
      try {
        processParagraph(docDom, p, spec, targetStyles, options, reports, styleIds);
      } catch (err) {
        reports.warnings.push('第 ' + (pi + 1) + ' 个段落处理失败：' + (err && err.message));
        if (reports.warnings.length > 20) break;
      }
    }

    /* ---------- 2.5 一级标题（章）附加版式：另起一页 + 前后各空一行 ----------
       必须放在段落处理之后 —— 靠段落里已写好的 outlineLvl / pStyle 来认定章标题。
       新插入的空段落不再参与段落处理（它们的格式在 makeChapterBlank 里直接写好）。 */
    var chapterCfg = options.chapter !== undefined ? options.chapter : spec.chapter;
    if (chapterCfg) {
      var chRes = applyChapterLayout(docDom, spec, styleIds, chapterCfg);
      reports.chapterBreaks = chRes.breaks;
      reports.chapterBlanks = chRes.blanks;
    }

    /* ---------- 3. 回写 ---------- */
    zip.file(DOC_PATH, serializeXml(docDom));
    if (styleIds && styDom) zip.file(STY_PATH, serializeXml(styDom));

    var out = await zip.generateAsync({
      type: 'arraybuffer',
      compression: 'DEFLATE',
      compressionOptions: { level: 6 },
      mimeType: DOCX_MIME
    });

    return {
      data: out,
      mime: DOCX_MIME,
      report: reports
    };
  }

  /* ---------- 页面设置 ---------- */
  function applySectionProperties(doc, sectPr, spec) {
    if (spec.page) {
      var pgSz = kid(sectPr, 'pgSz');
      if (!pgSz) { pgSz = createEl(doc, 'pgSz'); insertOrdered(sectPr, pgSz, SECTPR_ORDER); }
      if (spec.page.w) setAttr(pgSz, 'w', spec.page.w);
      if (spec.page.h) setAttr(pgSz, 'h', spec.page.h);
      if (spec.page.orient) setAttr(pgSz, 'orient', spec.page.orient);
      else dropAttr(pgSz, 'orient');
    }
    if (spec.margin) {
      var pgMar = kid(sectPr, 'pgMar');
      if (!pgMar) { pgMar = createEl(doc, 'pgMar'); insertOrdered(sectPr, pgMar, SECTPR_ORDER); }
      ['top', 'right', 'bottom', 'left', 'header', 'footer', 'gutter'].forEach(function (k) {
        if (spec.margin[k] !== undefined && spec.margin[k] !== null) setAttr(pgMar, k, spec.margin[k]);
      });
    }
    if (spec.docGrid) {
      var dg = kid(sectPr, 'docGrid');
      if (!dg) { dg = createEl(doc, 'docGrid'); sectPr.appendChild(dg); }
      if (spec.docGrid.type) setAttr(dg, 'type', spec.docGrid.type);
      if (spec.docGrid.linePitch !== undefined) setAttr(dg, 'linePitch', spec.docGrid.linePitch);
      if (spec.docGrid.charSpace !== undefined) setAttr(dg, 'charSpace', spec.docGrid.charSpace);
    }
  }

  /* ---------- 单个段落 ---------- */
  function processParagraph(doc, p, spec, targetStyles, options, reports, styleIds) {
    reports.paragraphs++;

    var inTable = isInsideTable(p);
    if (inTable) reports.tableParagraphs++;

    var pPrOld = kid(p, 'pPr');
    var styleId = pPrOld ? attr(kid(pPrOld, 'pStyle'), 'val') : null;
    var styleRec = styleId && targetStyles.byId[styleId] ? targetStyles.byId[styleId] : null;
    var info = styleRec ? styleRec.info : null;

    /* 目录段落与目录域：整体跳过，保持原样 */
    if (info && info.type === 'toc') { reports.skippedFields++; return; }
    if (paragraphHasField(p, ['TOC', 'PAGEREF'])) { reports.skippedFields++; return; }
    /* 页眉页脚 / 脚注类样式不处理 */
    if (info && (info.type === 'header' || info.type === 'footer' || info.type === 'footnote')) {
      reports.skippedFields++;
      return;
    }

    var txt = paragraphText(p);
    /* 空段落（或只含空白）绝不能算标题：很多文档拿「带标题样式的空段」当空行用，
       照着样式套会把这些空行变成小二黑体居中，整篇版式就散了 */
    var blank = !txt.replace(/[\s\u00a0\u3000]/g, '').length;
    if (blank) reports.blankParagraphs++;

    var level = 0;         /* 0 = 正文 */
    var kind = 'body';     /* body | heading | caption */

    /* 先看文字里的编号规律。很多文档的标题样式与实际层级对不上
       （实测：「第1章」套 heading 2、「1.1」套 heading 3），
       照着样式套就会整体错位一级——所以编号规律优先于样式。 */
    var numInfo = (!blank && !inTable) ? heuristicInfo(txt) : null;

    if (info && !blank) {
      if (info.type === 'heading') { level = info.level; kind = 'heading'; }
      else if (info.type === 'caption') { kind = 'caption'; }
      else { kind = 'body'; }
    } else if (!blank && styleRec && styleRec.outlineLvl !== undefined && styleRec.outlineLvl !== null) {
      var ol = styleRec.outlineLvl;
      if (ol >= 0 && ol <= 8) { level = ol + 1; kind = 'heading'; }
    }

    /* 段落级 outlineLvl 兜底 */
    if (!blank && kind === 'body' && pPrOld) {
      var olv = num(attr(kid(pPrOld, 'outlineLvl'), 'val'));
      if (olv !== null && olv >= 0 && olv <= 8) { level = olv + 1; kind = 'heading'; }
    }

    /* 可靠编号（第X章 / X.Y / X.Y.Z …）一锤定音：
       - 段落已是标题 → 用它校正层级（样式错位也能救回来）
       - 段落本是正文 → 提升为标题
       弱编号（一、/（一）/1.）只用于「没样式可依据」的探测，不覆盖样式，
       否则会把用「一、二、三」当章标题的文档整体压低一级。 */
    if (numInfo && numInfo.strong) {
      if (kind === 'heading') {
        if (level !== numInfo.level) reports.levelFixes++;
      } else {
        reports.heuristicHits++;
      }
      level = numInfo.level;
      kind = 'heading';
    } else if (!blank && kind === 'body' && options.heuristic && !inTable) {
      /* 无样式可依据时的启发式识别 */
      var hl = heuristicLevel(txt);
      if (hl > 0) {
        level = hl; kind = 'heading';
        reports.heuristicHits++;
      }
    }

    /* 题注兜底：样式没标出「题注」时，按「表4-1 …」「图2-1 …」的文字规律认 */
    if (!blank && kind === 'body' && options.heuristic && !inTable && isCaptionText(txt)) {
      kind = 'caption';
    }

    /* 原本是标题样式、却被判定不该当标题的段落（主要是空段），
       下面会把它的 pStyle 摘掉——否则 Word 导航窗格和「更新目录」里
       仍然会冒出这些空条目 */
    var demoteHeading = !!(info && info.type === 'heading' && kind !== 'heading');

    /* 选择目标格式 */
    var specStyle = null;
    var forceIndent = false;
    var styleLevel = 0;     /* 实际用上的标题级别（模板缺级时会回退） */

    if (kind === 'heading') {
      if (!options.applyHeading) return;
      specStyle = spec.headings && spec.headings[level];
      styleLevel = level;
      if (!specStyle) {
        /* 模板没定义该级标题：退回最接近的低级标题，pStyle 也跟着回退 */
        for (var k = level - 1; k >= 1; k--) {
          if (spec.headings && spec.headings[k]) { specStyle = spec.headings[k]; styleLevel = k; break; }
        }
      }
      if (!specStyle) { specStyle = spec.normal; styleLevel = 0; }
      forceIndent = true;   /* 标题不加首行缩进 */
    } else if (kind === 'caption') {
      if (!options.applyCaption) return;
      specStyle = spec.caption || spec.normal;
      forceIndent = true;
    } else {
      if (!options.applyBody) return;
      specStyle = spec.normal;
      forceIndent = blank;   /* 空段落不留首行缩进 */
    }

    if (!specStyle) return;

    /* ---- 写入 pPr ---- */
    var pPr = pPrOld;
    if (!pPr) {
      pPr = createEl(doc, 'pPr');
      p.insertBefore(pPr, p.firstChild);
    } else if (p.firstChild !== pPr) {
      p.insertBefore(pPr, p.firstChild);
    }

    /* ---- 把 pStyle 改挂到与判定结果一致的样式 ----
       源文档常见「第1章套 heading 2」这类整体错位；只写直接格式的话，
       Word 样式库 / 样式检查器里仍显示错的名字，用户打开会以为没排对。
       同理，原本挂着标题样式的空段落要摘掉，免得导航窗格和「更新目录」
       里冒出一堆空条目。 */
    if (styleIds) {
      var wantedId = null;
      if (demoteHeading) { reports.demotedHeadings++; wantedId = styleIds.normal; }
      else if (kind === 'heading') wantedId = styleIds.headings && styleIds.headings[styleLevel];
      else if (kind === 'caption') wantedId = styleIds.caption;
      else wantedId = styleIds.normal;

      if (wantedId) {
        if (attr(kid(pPr, 'pStyle'), 'val') !== wantedId) {
          removeChild(pPr, 'pStyle');
          var ps = createEl(doc, 'pStyle'); setAttr(ps, 'val', wantedId);
          insertOrdered(pPr, ps, PPR_ORDER);
        }
      } else {
        removeChild(pPr, 'pStyle');
      }
    } else if (demoteHeading) {
      removeChild(pPr, 'pStyle');
      reports.demotedHeadings++;
    }

    var targetPPr = specStyle.pPr || {};

    /* 对齐 */
    if (options.clearManual) {
      removeChild(pPr, 'jc');
      if (targetPPr.align) {
        var jc = createEl(doc, 'jc'); setAttr(jc, 'val', targetPPr.align);
        insertOrdered(pPr, jc, PPR_ORDER);
      }
      /* 间距 */
      removeChild(pPr, 'spacing');
      if (targetPPr.spacing) {
        var sp = createEl(doc, 'spacing');
        var ts = targetPPr.spacing;
        ['before', 'after', 'line', 'lineRule', 'beforeLines', 'afterLines'].forEach(function (key) {
          if (ts[key] !== undefined && ts[key] !== null) setAttr(sp, key, ts[key]);
        });
        insertOrdered(pPr, sp, PPR_ORDER);
      }
      /* 缩进 */
      removeChild(pPr, 'ind');
      removeChild(pPr, 'hanging');          /* w:ind 上的挂起缩进属性，与 firstLine 互斥 */
      var indSrc = targetPPr.ind || {};
      var indObj = {};
      if (forceIndent) {
        /* 标题 / 题注：模板没给缩进 → 显式写 0，截断样式链上的首行缩进继承 */
        indObj = Object.assign({}, ZERO_IND);
      } else if (inTable && options.skipTableIndent) {
        if (indSrc.leftChars !== undefined) indObj.leftChars = indSrc.leftChars;
        if (indSrc.left !== undefined) indObj.left = indSrc.left;
      } else {
        indObj = Object.assign({}, indSrc);
      }
      if (Object.keys(indObj).length) {
        var ind = createEl(doc, 'ind');
        Object.keys(indObj).forEach(function (key) { setAttr(ind, key, indObj[key]); });
        insertOrdered(pPr, ind, PPR_ORDER);
      }
      /* 其他段落属性 */
      applyOnOff(pPr, doc, 'keepNext', targetPPr.keepNext, PPR_ORDER);
      applyOnOff(pPr, doc, 'keepLines', targetPPr.keepLines, PPR_ORDER);
      applyOnOff(pPr, doc, 'contextualSpacing', targetPPr.contextualSpacing, PPR_ORDER);
    }

    /* 大纲级别：让导航窗格能识别（仅 heading） */
    if (kind === 'heading') {
      removeChild(pPr, 'outlineLvl');
      var olEl = createEl(doc, 'outlineLvl');
      setAttr(olEl, 'val', level - 1);
      insertOrdered(pPr, olEl, PPR_ORDER);
    }

    /* ---- 写入 run 的 rPr ---- */
    var targetRPr = specStyle.rPr || {};
    var runs = p.getElementsByTagNameNS ?
      p.getElementsByTagNameNS(W_NS, 'r') : p.getElementsByTagName('w:r');

    for (var ri = 0; ri < runs.length; ri++) {
      var r = runs[ri];
      if (r.parentNode && r.parentNode.localName === 'del') continue;   /* 修订删除内容 */
      applyRunFormat(doc, r, targetRPr, options);
    }

    /* ---- 段落标记本身的格式（决定空行高度） ---- */
    applyParagraphMarkFormat(doc, pPr, targetRPr, options);

    /* ---- 统计 ---- */
    if (kind === 'heading') reports.headings[level] = (reports.headings[level] || 0) + 1;
    else if (kind === 'caption') reports.caption++;
    else reports.body++;
  }

  function applyOnOff(parent, doc, name, value, order) {
    removeChild(parent, name);
    if (value === undefined || value === null) return;
    var el = createEl(doc, name);
    if (value === false) setAttr(el, 'val', '0');
    insertOrdered(parent, el, order);
  }

  /* clearManual 开启时，这些字符级属性一律以模板为准，不再保留目标文档的手工设置。
     注意 vertAlign（上下标）、lang、noProof 等语义属性不在此列，绝不能删。 */
  var FORCE_CLEAR_RPR = ['b', 'bCs', 'i', 'iCs', 'color', 'u', 'strike', 'dstrike',
    'highlight', 'shd', 'caps', 'smallCaps', 'outline', 'shadow', 'emboss', 'imprint', 'w'];

  function applyRunFormat(doc, r, targetRPr, options) {
    if (!runHasRealText(r)) return;   /* 域字符、书签等 run 不动 */

    var rPr = kid(r, 'rPr');
    if (!rPr) {
      rPr = createEl(doc, 'rPr');
      r.insertBefore(rPr, r.firstChild);
    } else if (r.firstChild !== rPr) {
      r.insertBefore(rPr, r.firstChild);
    }

    writeRPrInto(doc, rPr, targetRPr, options);
  }

  /**
   * 把一组字符格式写进任意 rPr 元素。
   * 既用于 run 的直接格式，也用于 styles.xml 里的样式定义（保证两处规则一致）。
   */
  function writeRPrInto(doc, rPr, targetRPr, options) {
    targetRPr = targetRPr || {};

    /* 先清掉与模板冲突的手工格式 */
    if (options.clearManual) {
      FORCE_CLEAR_RPR.forEach(function (n) { removeChild(rPr, n); });
    }

    /* 字体 */
    if (targetRPr.font) {
      var rf = kid(rPr, 'rFonts');
      if (!rf) { rf = createEl(doc, 'rFonts'); insertOrdered(rPr, rf, RPR_ORDER); }
      if (targetRPr.font.ascii) setAttr(rf, 'ascii', targetRPr.font.ascii);
      if (targetRPr.font.hAnsi) setAttr(rf, 'hAnsi', targetRPr.font.hAnsi);
      if (targetRPr.font.eastAsia) setAttr(rf, 'eastAsia', targetRPr.font.eastAsia);
      if (targetRPr.font.cs) setAttr(rf, 'cs', targetRPr.font.cs);
      if (targetRPr.font.hint) setAttr(rf, 'hint', targetRPr.font.hint);
    }

    /* 字号 */
    if (targetRPr.size) {
      var sz = kid(rPr, 'sz');
      if (!sz) { sz = createEl(doc, 'sz'); insertOrdered(rPr, sz, RPR_ORDER); }
      setAttr(sz, 'val', targetRPr.size);
      var szCs = kid(rPr, 'szCs');
      if (!szCs) { szCs = createEl(doc, 'szCs'); insertOrdered(rPr, szCs, RPR_ORDER); }
      setAttr(szCs, 'val', targetRPr.sizeCs || targetRPr.size);
    }

    /* 粗体 / 斜体 / 颜色：模板定义了什么就写什么，没定义就是"清除" */
    if (options.clearManual) {
      writeToggle(rPr, doc, 'b', targetRPr.bold);
      writeToggle(rPr, doc, 'bCs', targetRPr.bold);
      writeToggle(rPr, doc, 'i', targetRPr.italic);
      writeToggle(rPr, doc, 'iCs', targetRPr.italic);
      if (targetRPr.color) {
        var color = createEl(doc, 'color'); setAttr(color, 'val', targetRPr.color);
        insertOrdered(rPr, color, RPR_ORDER);
      }
      if (targetRPr.underline) {
        var u = createEl(doc, 'u'); setAttr(u, 'val', targetRPr.underline);
        insertOrdered(rPr, u, RPR_ORDER);
      }
      if (targetRPr.charSpacing !== undefined && targetRPr.charSpacing !== null) {
        var csp = createEl(doc, 'spacing'); setAttr(csp, 'val', targetRPr.charSpacing);
        insertOrdered(rPr, csp, RPR_ORDER);
      }
    }
  }

  function writeToggle(rPr, doc, name, value) {
    removeChild(rPr, name);
    if (value === undefined || value === null) return;
    if (value === true) {
      insertOrdered(rPr, createEl(doc, name), RPR_ORDER);
    } else {
      var el = createEl(doc, name); setAttr(el, 'val', '0');
      insertOrdered(rPr, el, RPR_ORDER);
    }
  }

  function applyParagraphMarkFormat(doc, pPr, targetRPr, options) {
    if (!options.clearManual) return;
    var rPr = kid(pPr, 'rPr');
    if (!rPr) {
      rPr = createEl(doc, 'rPr');
      insertOrdered(pPr, rPr, PPR_ORDER);
    }
    FORCE_CLEAR_RPR.forEach(function (n) { removeChild(rPr, n); });
    if (targetRPr.font) {
      var rf = kid(rPr, 'rFonts');
      if (!rf) { rf = createEl(doc, 'rFonts'); insertOrdered(rPr, rf, RPR_ORDER); }
      if (targetRPr.font.eastAsia) setAttr(rf, 'eastAsia', targetRPr.font.eastAsia);
      if (targetRPr.font.ascii) setAttr(rf, 'ascii', targetRPr.font.ascii);
      if (targetRPr.font.hAnsi) setAttr(rf, 'hAnsi', targetRPr.font.hAnsi);
    }
    if (targetRPr.size) {
      var sz = kid(rPr, 'sz');
      if (!sz) { sz = createEl(doc, 'sz'); insertOrdered(rPr, sz, RPR_ORDER); }
      setAttr(sz, 'val', targetRPr.size);
    }
  }

  /* ================================================================== *
   * 十一、规范描述（给 UI 展示用）
   * ================================================================== */

  function halfPointToPt(hp) { return hp / 2; }
  function twipToPt(t) { return t / 20; }

  function cnSizeName(halfPt) {
    for (var i = 0; i < CN_SIZES.length; i++) if (CN_SIZES[i][0] === halfPt) return CN_SIZES[i][1];
    return null;
  }

  function fontLabel(font) {
    if (!font) return '（跟随样式）';
    return [font.eastAsia || font.ascii || font.hAnsi, font.ascii && font.eastAsia ? '(' + font.ascii + ')' : '']
      .filter(Boolean).join(' ');
  }

  function sizeLabel(halfPt) {
    if (!halfPt) return '—';
    var cn = cnSizeName(halfPt);
    var pt = halfPointToPt(halfPt);
    return pt + ' 磅' + (cn ? '（' + cn + '）' : '');
  }

  function lineSpacingLabel(spacing) {
    if (!spacing || spacing.line === undefined || spacing.line === null) return '单倍（默认）';
    var rule = spacing.lineRule || 'auto';
    if (rule === 'auto') {
      var m = Math.round(spacing.line / 240 * 100) / 100;
      return (m === 1 ? '单倍' : m + ' 倍行距');
    }
    if (rule === 'exact') return '固定值 ' + twipToPt(spacing.line) + ' 磅';
    return '最小值 ' + twipToPt(spacing.line) + ' 磅';
  }

  function spacingLabel(v) {
    if (!v && v !== 0) return '0 磅';
    return twipToPt(v) + ' 磅';
  }

  function indentLabel(ind, baseSize) {
    if (!ind) return '无';
    if (ind.firstLineChars !== undefined) return (ind.firstLineChars / 100) + ' 字符（首行缩进）';
    if (ind.firstLine !== undefined && ind.firstLine > 0) {
      var pt = twipToPt(ind.firstLine);
      if (baseSize) return (pt / halfPointToPt(baseSize)).toFixed(1) + ' 字符（首行缩进）';
      return pt + ' 磅（首行缩进）';
    }
    if (ind.hanging !== undefined && ind.hanging > 0) return twipToPt(ind.hanging) + ' 磅（悬挂缩进）';
    if (ind.leftChars !== undefined) return (ind.leftChars / 100) + ' 字符（左缩进）';
    if (ind.left !== undefined) return twipToPt(ind.left) + ' 磅（左缩进）';
    return '无';
  }

  function alignLabel(a) {
    return { left: '左对齐', center: '居中', right: '右对齐', both: '两端对齐', distribute: '分散对齐' }[a] || '默认';
  }

  function describeSpec(spec) {
    var out = { page: [], styles: [], source: spec.source, missing: spec.missing || [], foundNames: spec.foundNames || [] };

    if (spec.page) {
      var w = spec.page.w ? (twipToPt(spec.page.w)).toFixed(1) : '?';
      var h = spec.page.h ? (twipToPt(spec.page.h)).toFixed(1) : '?';
      var name = '自定义';
      if (spec.page.w === 11906 && spec.page.h === 16838) name = 'A4';
      else if (spec.page.w === 12240 && spec.page.h === 15840) name = 'Letter';
      out.page.push({ k: '纸张', v: name + '（' + w + ' × ' + h + ' 磅）' });
    }
    if (spec.margin) {
      var m = spec.margin;
      var fmt = function (v) { return v === undefined || v === null ? '—' : (twipToPt(v)).toFixed(2).replace(/\.00$/, '') + ' 厘米'; };
      out.page.push({
        k: '页边距',
        v: '上 ' + cm(m.top) + ' / 下 ' + cm(m.bottom) + ' / 左 ' + cm(m.left) + ' / 右 ' + cm(m.right)
      });
      if (m.header !== undefined || m.footer !== undefined) {
        out.page.push({ k: '页眉 / 页脚', v: '距边界 ' + cm(m.header) + ' / ' + cm(m.footer) });
      }
    }
    if (spec.docGrid && spec.docGrid.linePitch) {
      out.page.push({ k: '文档网格', v: '每页行数由网格控制（行间距 ' + twipToPt(spec.docGrid.linePitch).toFixed(1) + ' 磅）' });
    }

    function pushStyle(label, s) {
      if (!s) return;
      var r = s.rPr || {}, p = s.pPr || {};
      out.styles.push({
        label: label,
        name: s.name,
        font: fontLabel(r.font),
        size: sizeLabel(r.size),
        bold: r.bold === true ? '加粗' : (r.bold === false ? '常规' : '—'),
        align: alignLabel(p.align),
        spacing: lineSpacingLabel(p.spacing),
        before: p.spacing && p.spacing.before !== undefined ? spacingLabel(p.spacing.before) : '0 磅',
        after: p.spacing && p.spacing.after !== undefined ? spacingLabel(p.spacing.after) : '0 磅',
        indent: indentLabel(p.ind, r.size)
      });
    }

    pushStyle('正文', spec.normal);
    for (var lv = 1; lv <= 9; lv++) if (spec.headings && spec.headings[lv]) pushStyle('标题 ' + lv, spec.headings[lv]);
    if (spec.caption) pushStyle('题注', spec.caption);
    if (spec.tableText) pushStyle('表格文字', spec.tableText);

    if (spec.chapter) {
      var ch = spec.chapter;
      var bits = [];
      if (ch.pageBreakBefore) bits.push('每章另起一页');
      var nb = ch.blankLinesBefore || 0, na = ch.blankLinesAfter || 0;
      if (nb || na) {
        bits.push('标题' + (nb ? '上' : '') + (na ? '下' : '') + '各空 ' + Math.max(nb, na) + ' 行（空行同样是一级标题格式）');
      }
      if (bits.length) out.page.push({ k: '一级标题版式', v: bits.join('，') });
    }

    return out;
  }

  function cm(twips) {
    if (twips === undefined || twips === null) return '—';
    return (twips / 567).toFixed(2).replace(/\.00$/, '') + ' 厘米';
  }

  /* ================================================================== *
   * 十二、内置预设（模板缺失时的兜底）
   * ================================================================== */

  function makeRPr(font, sizeHalfPt, bold) {
    var o = { font: { eastAsia: font, ascii: 'Times New Roman', hAnsi: 'Times New Roman' }, size: sizeHalfPt, sizeCs: sizeHalfPt };
    if (bold) o.bold = true;
    return o;
  }

  var BUILTIN_PRESETS = {
    'thesis-cn': {
      label: '中文学术论文通用规范',
      spec: {
        source: { file: '内置预设：中文学术论文通用规范', generatedAt: '' },
        page: { w: 11906, h: 16838 },
        margin: { top: 1440, right: 1440, bottom: 1440, left: 1800, header: 851, footer: 992 },
        docGrid: null,
        normal: { styleId: null, name: '正文', rPr: makeRPr('宋体', 24, false), pPr: { align: 'both', spacing: { line: 360, lineRule: 'auto' }, ind: { firstLineChars: 200, firstLine: 480 } } },
        headings: {
          1: { styleId: null, name: '标题 1', rPr: makeRPr('黑体', 30, false), pPr: { align: 'left', spacing: { before: 240, after: 180, line: 360, lineRule: 'auto' } } },
          2: { styleId: null, name: '标题 2', rPr: makeRPr('黑体', 28, false), pPr: { align: 'left', spacing: { before: 180, after: 120, line: 360, lineRule: 'auto' } } },
          3: { styleId: null, name: '标题 3', rPr: makeRPr('黑体', 24, false), pPr: { align: 'left', spacing: { before: 120, after: 60, line: 360, lineRule: 'auto' } } }
        },
        caption: { styleId: null, name: '题注', rPr: makeRPr('宋体', 21, false), pPr: { align: 'center', spacing: { before: 60, after: 60, line: 300, lineRule: 'auto' } } },
        tableText: null,
        foundNames: [], missing: []
      }
    },
    /* 课程设计报告规范 —— 依据用户提供的《课程设计报告模板.doc》逐项提取
     *   页面：A4，上 3.5cm / 下 3.0cm / 左 2.5cm / 右 2.0cm，页眉 2.5cm / 页脚 2.0cm
     *   正文：小四号宋体（西文 Times New Roman），两端对齐，行距固定值 20 磅，首行缩进 2 字符
     *   一级标题（章）：小二号黑体居中，段前 0 / 段后 0，单倍行距
     *   二级标题（节）：小三号黑体，段前 12 磅 / 段后 8 磅，单倍行距
     *   三级标题（条）：四号黑体，段前 8 磅 / 段后 6 磅，单倍行距
     *   图题表题：五号宋体居中
     *   目录：各章题序及标题 小四号黑体，其余 小四号宋体
     *   页码：页脚阿拉伯数字（-1-、-2-） */
    'report-cn': {
      label: '课程设计报告规范',
      spec: {
        source: { file: '内置预设：课程设计报告规范（依据《课程设计报告模板》）', generatedAt: '' },
        page: { w: 11906, h: 16838 },
        margin: { top: 1985, right: 1134, bottom: 1701, left: 1418, header: 1418, footer: 1134 },
        docGrid: null,
        normal: { styleId: null, name: '正文', rPr: makeRPr('宋体', 24, false), pPr: { align: 'both', spacing: { line: 400, lineRule: 'exact' }, ind: { firstLineChars: 200, firstLine: 480 } } },
        headings: {
          1: { styleId: null, name: '标题 1', rPr: makeRPr('黑体', 36, false), pPr: { align: 'center', spacing: { before: 0, after: 0, line: 240, lineRule: 'auto' } } },
          2: { styleId: null, name: '标题 2', rPr: makeRPr('黑体', 30, false), pPr: { align: 'left', spacing: { before: 240, after: 160, line: 240, lineRule: 'auto' } } },
          3: { styleId: null, name: '标题 3', rPr: makeRPr('黑体', 28, false), pPr: { align: 'left', spacing: { before: 160, after: 120, line: 240, lineRule: 'auto' } } }
        },
        caption: { styleId: null, name: '题注', rPr: makeRPr('宋体', 21, false), pPr: { align: 'center', spacing: { before: 120, after: 120, line: 300, lineRule: 'auto' } } },
        /* 一级标题（章）的附加版式：另起一页，上下各空一行，
           两个空行同样套一级标题格式（小二黑体居中，行高才和标题一致） */
        chapter: { pageBreakBefore: true, blankLinesBefore: 1, blankLinesAfter: 1 },
        tableText: null,
        foundNames: [], missing: []
      }
    }
  };

  /* ================================================================== *
   * 导出
   * ================================================================== */

  return {
    analyzeTemplate: analyzeTemplate,
    applyFormat: applyFormat,
    disableDocGrid: disableDocGrid,
    describeSpec: describeSpec,
    parseStylesXml: parseStylesXml,
    BUILTIN_PRESETS: BUILTIN_PRESETS,
    DOCX_MIME: DOCX_MIME,
    _internal: {
      parseRPr: parseRPr,
      parsePPr: parsePPr,
      parseSectPr: parseSectPr,
      heuristicLevel: heuristicLevel,
      heuristicInfo: heuristicInfo,
      classifyStyle: classifyStyle,
      syncSpecStyles: syncSpecStyles,
      writeRPrInto: writeRPrInto,
      writeStylePPrInto: writeStylePPrInto,
      collectHorizontalRules: collectHorizontalRules,
      applyChapterLayout: applyChapterLayout,
      paragraphsOf: paragraphsOf
    }
  };
});
