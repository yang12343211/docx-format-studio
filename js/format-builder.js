/*!
 * format-builder.js — 手动设置文档格式规范
 *
 * 对外 API（挂在 window.FormatBuilder）：
 *   init(options)      初始化（options.onChange 在规范变化时回调，参数是 spec）
 *   getSpec()          取出当前设置对应的 spec（与 docx-core 的 spec 结构一致）
 *   applySpec(spec)    用一份 spec 回填表单（用于套用预设）
 *   setActive(key)     切换当前编辑的样式（normal / h1 / h2 / h3 / caption）
 */
(function (root) {
  'use strict';

  /* ============================ 选项数据 ============================ */

  var FONTS_CJK = ['宋体', '黑体', '楷体', '仿宋', '微软雅黑', '等线', '华文中宋', '华文楷体', '幼圆'];

  var FONTS_LATIN = ['Times New Roman', 'Arial', 'Calibri', 'Cambria', 'Georgia', 'Verdana', 'Courier New', 'Tahoma'];

  /* 中文字号 -> 半磅 */
  var SIZES = [
    [84, '初号　42 磅'], [72, '小初　36 磅'], [52, '一号　26 磅'], [48, '小一　24 磅'],
    [44, '二号　22 磅'], [36, '小二　18 磅'], [32, '三号　16 磅'], [30, '小三　15 磅'],
    [28, '四号　14 磅'], [24, '小四　12 磅'], [21, '五号　10.5 磅'], [18, '小五　9 磅'],
    [15, '六号　7.5 磅'], [13, '小六　6.5 磅'], [11, '七号　5.5 磅'], [10, '八号　5 磅']
  ];

  /* 行距：数值即 w:line（twips），240 = 单倍 */
  var LINE_MODES = [
    [240, '单倍行距'], [276, '1.15 倍'], [300, '1.25 倍'], [360, '1.5 倍'],
    [420, '1.75 倍'], [480, '2 倍'], [576, '2.4 倍'],
    ['exact', '固定值'], ['atLeast', '最小值']
  ];

  var ALIGNS = [['left', '左对齐'], ['center', '居中'], ['right', '右对齐'], ['both', '两端对齐']];

  /* 纸张尺寸（twips） */
  var PAPERS = {
    A4: [11906, 16838],
    A3: [16838, 23811],
    B5: [9979, 14170],
    Letter: [12240, 15840]
  };

  /* ============================ 默认规范 ============================ */

  function mkStyle(o) {
    return Object.assign({
      cjk: '宋体', latin: 'Times New Roman', size: 24,
      bold: false, italic: false, align: 'both',
      lineMode: '360', lineValue: 12,
      before: 0, after: 0, indent: 0
    }, o);
  }

  function defaultState() {
    return {
      active: 'normal',
      styles: {
        /* 正文默认按中文学术论文的习惯：首行缩进 2 字符 */
        normal: mkStyle({ indent: 2 }),
        h1: mkStyle({
          cjk: '黑体', size: 32, align: 'center', before: 12, after: 12
        }),
        h2: mkStyle({
          cjk: '黑体', size: 28, align: 'left', before: 6, after: 6
        }),
        h3: mkStyle({
          cjk: '黑体', size: 24, bold: true, align: 'left', before: 6, after: 6
        }),
        caption: mkStyle({
          size: 21, align: 'center', lineMode: '300', before: 3, after: 3
        })
      },
      page: {
        paper: 'A4', w: 11906, h: 16838,
        top: 2.54, bottom: 2.54, left: 3.17, right: 2.54,
        header: 1.5, footer: 1.75
      },
      /* 一级标题（章）的附加版式，默认关闭；课设模板规范预设会打开它 */
      chapter: { pageBreakBefore: false, blanks: false }
    };
  }

  var STYLE_META = [
    { key: 'normal', label: '正文', tag: 'p' },
    { key: 'h1', label: '一级标题', tag: 'h1', level: 1 },
    { key: 'h2', label: '二级标题', tag: 'h2', level: 2 },
    { key: 'h3', label: '三级标题', tag: 'h3', level: 3 },
    { key: 'caption', label: '图题表题', tag: 'span' }
  ];

  var state = defaultState();
  var onChange = null;

  /* ============================ 工具 ============================ */

  function $(id) { return document.getElementById(id); }

  function esc(s) {
    return String(s === undefined || s === null ? '' : s)
      .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;').replace(/'/g, '&#39;');
  }

  function cmToTwips(cm) { return Math.round((parseFloat(cm) || 0) * 567); }

  /* ============================ spec 生成 ============================ */

  function styleToSpec(s, name, isHeading, level) {
    var sizePt = s.size / 2;
    var pPr = {
      align: s.align,
      spacing: {}
    };

    if (s.lineMode === 'exact' || s.lineMode === 'atLeast') {
      pPr.spacing.line = Math.round((parseFloat(s.lineValue) || 12) * 20);
      pPr.spacing.lineRule = s.lineMode;
    } else {
      pPr.spacing.line = parseInt(s.lineMode, 10);
      pPr.spacing.lineRule = 'auto';
    }
    if (s.before) pPr.spacing.before = Math.round(s.before * 20);
    if (s.after) pPr.spacing.after = Math.round(s.after * 20);

    /* 标题与题注不缩进；正文按设置的字符数缩进 */
    if (!isHeading && s.indent > 0) {
      pPr.ind = {
        firstLineChars: Math.round(s.indent * 100),
        firstLine: Math.round(s.indent * sizePt * 20)
      };
    }

    var spec = {
      styleId: null,
      name: name,
      rPr: {
        font: { eastAsia: s.cjk, ascii: s.latin, hAnsi: s.latin },
        size: s.size,
        sizeCs: s.size,
        bold: !!s.bold
      },
      pPr: pPr
    };
    if (s.italic) spec.rPr.italic = true;
    return spec;
  }

  function getSpec() {
    var st = state.styles;
    return {
      source: { file: '手动设置', generatedAt: new Date().toISOString() },
      page: { w: state.page.w, h: state.page.h },
      margin: {
        top: cmToTwips(state.page.top),
        bottom: cmToTwips(state.page.bottom),
        left: cmToTwips(state.page.left),
        right: cmToTwips(state.page.right),
        header: cmToTwips(state.page.header),
        footer: cmToTwips(state.page.footer)
      },
      docGrid: null,
      normal: styleToSpec(st.normal, '正文', false),
      headings: {
        1: styleToSpec(st.h1, '标题 1', true, 1),
        2: styleToSpec(st.h2, '标题 2', true, 2),
        3: styleToSpec(st.h3, '标题 3', true, 3)
      },
      caption: styleToSpec(st.caption, '题注', true),
      /* 一级标题（章）附加版式：每章另起一页 + 上下各空一行（空行同用一级标题格式） */
      chapter: (state.chapter.pageBreakBefore || state.chapter.blanks) ? {
        pageBreakBefore: !!state.chapter.pageBreakBefore,
        blankLinesBefore: state.chapter.blanks ? 1 : 0,
        blankLinesAfter: state.chapter.blanks ? 1 : 0
      } : null,
      foundNames: [],
      missing: []
    };
  }

  /* ============================ 表单渲染 ============================ */

  function optList(list, cur, cast) {
    return list.map(function (item) {
      var v, label;
      if (Array.isArray(item)) { v = item[0]; label = item[1]; }
      else { v = item; label = item; }
      var vv = cast ? cast(v) : v;
      var sel = String(vv) === String(cur) ? ' selected' : '';
      return '<option value="' + esc(v) + '"' + sel + '>' + esc(label) + '</option>';
    }).join('');
  }

  function fieldsHtml(s) {
    var isBody = state.active === 'normal';
    var fixedLine = (s.lineMode === 'exact' || s.lineMode === 'atLeast');

    return '' +
      '<div class="fb-grid">' +

      '<label class="fb-field"><span>中文字体</span>' +
      '<select data-f="cjk">' + optList(FONTS_CJK, s.cjk) + '</select></label>' +

      '<label class="fb-field"><span>西文字体</span>' +
      '<select data-f="latin">' + optList(FONTS_LATIN, s.latin) + '</select></label>' +

      '<label class="fb-field"><span>字号</span>' +
      '<select data-f="size">' + optList(SIZES, s.size, Number) + '</select></label>' +

      '<label class="fb-field"><span>对齐方式</span>' +
      '<select data-f="align">' + optList(ALIGNS, s.align) + '</select></label>' +

      '<label class="fb-field"><span>行距</span>' +
      '<select data-f="lineMode">' + optList(
        LINE_MODES.map(function (m) { return [m[0], m[1]]; }), s.lineMode,
        function (v) { return /^\d+$/.test(v) ? Number(v) : v; }) + '</select></label>' +

      '<label class="fb-field' + (fixedLine ? '' : ' fb-off') + '" data-when="fixedLine">' +
      '<span>行距数值</span>' +
      '<span class="fb-num"><input type="number" data-f="lineValue" min="1" max="200" step="0.5" value="' + esc(s.lineValue) + '"><i>磅</i></span>' +
      '</label>' +

      '<label class="fb-field"><span>段前间距</span>' +
      '<span class="fb-num"><input type="number" data-f="before" min="0" max="100" step="1" value="' + esc(s.before) + '"><i>磅</i></span></label>' +

      '<label class="fb-field"><span>段后间距</span>' +
      '<span class="fb-num"><input type="number" data-f="after" min="0" max="100" step="1" value="' + esc(s.after) + '"><i>磅</i></span></label>' +

      '<label class="fb-field' + (isBody ? '' : ' fb-off') + '" data-when="indent">' +
      '<span>首行缩进</span>' +
      '<span class="fb-num"><input type="number" data-f="indent" min="0" max="8" step="0.5" value="' + esc(s.indent) + '"><i>字符</i></span>' +
      '</label>' +

      '<div class="fb-field fb-field-shape"><span>字形</span>' +
      '<span class="fb-checks">' +
      '<label class="fb-check"><input type="checkbox" data-f="bold"' + (s.bold ? ' checked' : '') + '><i>加粗</i></label>' +
      '<label class="fb-check"><input type="checkbox" data-f="italic"' + (s.italic ? ' checked' : '') + '><i>倾斜</i></label>' +
      '</span></div>' +

      (state.active === 'h1' ? chapterFieldsHtml() : '') +

      '</div>';
  }

  /* 一级标题（章）的附加版式：只在一级标题这一页出现 */
  function chapterFieldsHtml() {
    var c = state.chapter;
    return '' +
      '<div class="fb-field fb-wide"><span>一级标题（章）版式</span>' +
      '<span class="fb-checks">' +
      '<label class="fb-check"><input type="checkbox" data-f="chapterPageBreak"' + (c.pageBreakBefore ? ' checked' : '') + '><i>每章另起一页</i></label>' +
      '<label class="fb-check"><input type="checkbox" data-f="chapterBlanks"' + (c.blanks ? ' checked' : '') + '><i>标题上下各空一行</i></label>' +
      '</span>' +
      '<span class="fb-hint">空行同样套上面这一级标题的字体字号，所以空出来的行高和标题一致；' +
      '文档最开头那一章前面没有内容，不会再加分页。</span></div>';
  }

  function renderFields() {
    var s = state.styles[state.active];
    var box = $('fb-fields');
    if (!box) return;
    box.innerHTML = fieldsHtml(s);
  }

  /* ============================ 预览 ============================ */

  var PREVIEW_MAP = {
    normal: 'pv-body',
    h1: 'pv-h1',
    h2: 'pv-h2',
    h3: 'pv-h3',
    caption: 'pv-caption'
  };

  function applyToEl(el, s) {
    if (!el) return;
    el.style.fontFamily = '"' + s.cjk + '", "' + s.latin + '", serif';
    el.style.fontSize = (s.size / 2) + 'pt';
    el.style.fontWeight = s.bold ? '700' : '400';
    el.style.fontStyle = s.italic ? 'italic' : 'normal';
    el.style.textAlign = s.align === 'both' ? 'justify' : s.align;
    el.style.marginTop = s.before + 'pt';
    el.style.marginBottom = s.after + 'pt';

    if (s.lineMode === 'exact' || s.lineMode === 'atLeast') {
      el.style.lineHeight = (parseFloat(s.lineValue) || 12) + 'pt';
    } else {
      el.style.lineHeight = (parseInt(s.lineMode, 10) / 240).toFixed(2);
    }

    var isBody = el.id === 'pv-body';
    el.style.textIndent = (isBody && s.indent > 0) ? (s.indent + 'em') : '0';
  }

  function renderPreview() {
    Object.keys(PREVIEW_MAP).forEach(function (key) {
      applyToEl($(PREVIEW_MAP[key]), state.styles[key]);
    });
    /* 高亮当前编辑的样式 */
    Object.keys(PREVIEW_MAP).forEach(function (key) {
      var el = $(PREVIEW_MAP[key]);
      if (el) el.classList.toggle('pv-active', key === state.active);
    });
    var pv = $('fb-preview');
    if (pv) pv.style.padding = '20px 22px';
  }

  /* ============================ 事件 ============================ */

  var emitTimer = null;

  function emit() {
    renderPreview();
    if (!onChange) return;
    clearTimeout(emitTimer);
    emitTimer = setTimeout(function () { onChange(getSpec()); }, 160);
  }

  function bindFields() {
    var box = $('fb-fields');
    if (!box) return;

    box.addEventListener('input', handleFieldEvent);
    box.addEventListener('change', handleFieldEvent);
  }

  function handleFieldEvent(e) {
    var el = e.target;
    var f = el && el.dataset ? el.dataset.f : null;
    if (!f) return;

    /* 一级标题（章）版式：存在 state.chapter 上，不属于任何单个样式。
       注意 input 的 data-f 是 chapterPageBreak / chapterBlanks，
       而 state.chapter 的字段名是 pageBreakBefore / blanks —— 必须显式映射。 */
    if (f === 'chapterPageBreak') { state.chapter.pageBreakBefore = el.checked; emit(); return; }
    if (f === 'chapterBlanks') { state.chapter.blanks = el.checked; emit(); return; }

    var s = state.styles[state.active];

    if (el.type === 'checkbox') {
      s[f] = el.checked;
    } else if (f === 'size') {
      s.size = Number(el.value);
    } else if (f === 'lineMode') {
      s.lineMode = el.value;
      /* 切到固定值/最小值时重绘，露出数值输入框 */
      renderFields();
      var fresh = $('fb-fields');
      if (fresh && (s.lineMode === 'exact' || s.lineMode === 'atLeast')) {
        var focusEl = fresh.querySelector('[data-f="lineValue"]');
        if (focusEl) focusEl.focus();
      }
      emit();
      return;
    } else if (f === 'before' || f === 'after' || f === 'indent' || f === 'lineValue') {
      s[f] = el.value === '' ? 0 : Number(el.value);
    } else {
      s[f] = el.value;
    }

    emit();
  }

  function bindTabs() {
    var bar = $('fb-style-tabs');
    if (!bar) return;
    bar.addEventListener('click', function (e) {
      var btn = e.target.closest ? e.target.closest('.fb-tab') : null;
      if (!btn) return;
      setActive(btn.dataset.style);
    });
  }

  function setActive(key) {
    if (!state.styles[key]) return;
    state.active = key;
    Array.prototype.forEach.call(document.querySelectorAll('#fb-style-tabs .fb-tab'), function (b) {
      b.classList.toggle('active', b.dataset.style === key);
    });
    renderFields();
    renderPreview();
  }

  function bindPage() {
    var box = $('fb-page');
    if (!box) return;

    box.addEventListener('input', handlePageEvent);
    box.addEventListener('change', handlePageEvent);
  }

  function handlePageEvent(e) {
    var el = e.target;
    var f = el && el.dataset ? el.dataset.p : null;
    if (!f) return;

    if (f === 'paper') {
      state.page.paper = el.value;
      var dim = PAPERS[el.value];
      if (dim) { state.page.w = dim[0]; state.page.h = dim[1]; }
      renderPage();
      emit();
      return;
    }
    state.page[f] = el.value === '' ? 0 : Number(el.value);
    emit();
  }

  function renderPage() {
    var box = $('fb-page');
    if (!box) return;
    var p = state.page;
    var core = '' +
      '<label class="fb-field"><span>纸张大小</span><select data-p="paper">' +
      optList(Object.keys(PAPERS), p.paper) + '</select></label>' +
      '<label class="fb-field"><span>上边距</span><span class="fb-num"><input type="number" data-p="top" min="0" max="10" step="0.1" value="' + esc(p.top) + '"><i>厘米</i></span></label>' +
      '<label class="fb-field"><span>下边距</span><span class="fb-num"><input type="number" data-p="bottom" min="0" max="10" step="0.1" value="' + esc(p.bottom) + '"><i>厘米</i></span></label>' +
      '<label class="fb-field"><span>左边距</span><span class="fb-num"><input type="number" data-p="left" min="0" max="10" step="0.1" value="' + esc(p.left) + '"><i>厘米</i></span></label>' +
      '<label class="fb-field"><span>右边距</span><span class="fb-num"><input type="number" data-p="right" min="0" max="10" step="0.1" value="' + esc(p.right) + '"><i>厘米</i></span></label>' +
      '<label class="fb-field"><span>页眉距边界</span><span class="fb-num"><input type="number" data-p="header" min="0" max="10" step="0.1" value="' + esc(p.header) + '"><i>厘米</i></span></label>' +
      '<label class="fb-field"><span>页脚距边界</span><span class="fb-num"><input type="number" data-p="footer" min="0" max="10" step="0.1" value="' + esc(p.footer) + '"><i>厘米</i></span></label>' +
      '<div class="fb-field fb-field-shape"><span>纸张尺寸</span>' +
      '<span class="fb-hint" id="fb-paper-size"></span></div>';
    box.innerHTML = core;

    var hint = $('fb-paper-size');
    if (hint) {
      hint.textContent = (p.w / 567).toFixed(1) + ' × ' + (p.h / 567).toFixed(1) + ' 厘米';
    }
  }

  function bindPresets() {
    var bar = $('fb-presets');
    if (!bar) return;
    bar.addEventListener('click', function (e) {
      var btn = e.target.closest ? e.target.closest('[data-preset]') : null;
      if (!btn) return;
      var key = btn.dataset.preset;
      var preset = root.DocxCore && root.DocxCore.BUILTIN_PRESETS[key];
      if (!preset) return;
      Array.prototype.forEach.call(bar.querySelectorAll('[data-preset]'), function (b) {
        b.classList.toggle('active', b === btn);
      });
      applySpec(preset.spec);
    });
  }

  /* ============================ 用 spec 回填 ============================ */

  function specToStyle(sp, target) {
    var r = sp.rPr || {}, p = sp.pPr || {};
    if (r.font) {
      if (r.font.eastAsia) target.cjk = r.font.eastAsia;
      if (r.font.ascii) target.latin = r.font.ascii;
    }
    if (r.size) target.size = r.size;
    target.bold = r.bold === true;
    target.italic = r.italic === true;
    if (p.align) target.align = p.align;

    var spc = p.spacing || {};
    if (spc.line !== undefined && spc.line !== null) {
      if (spc.lineRule === 'exact' || spc.lineRule === 'atLeast') {
        target.lineMode = spc.lineRule;
        target.lineValue = spc.line / 20;
      } else {
        target.lineMode = String(spc.line);
      }
    }
    target.before = spc.before ? spc.before / 20 : 0;
    target.after = spc.after ? spc.after / 20 : 0;

    var ind = p.ind || {};
    if (ind.firstLineChars) target.indent = ind.firstLineChars / 100;
    else if (ind.firstLine) target.indent = +(ind.firstLine / 20 / (target.size / 2)).toFixed(1);
    else target.indent = 0;
  }

  function applySpec(spec) {
    if (!spec) return;
    var s = state.styles;

    if (spec.normal) specToStyle(spec.normal, s.normal);
    if (spec.headings) {
      if (spec.headings[1]) specToStyle(spec.headings[1], s.h1);
      if (spec.headings[2]) specToStyle(spec.headings[2], s.h2);
      if (spec.headings[3]) specToStyle(spec.headings[3], s.h3);
    }
    if (spec.caption) specToStyle(spec.caption, s.caption);

    /* 一级标题版式跟着预设走：课设模板规范会打开，其它预设则关掉 */
    if (spec.chapter) {
      state.chapter.pageBreakBefore = !!spec.chapter.pageBreakBefore;
      state.chapter.blanks = !!(spec.chapter.blankLinesBefore || spec.chapter.blankLinesAfter);
    } else {
      state.chapter.pageBreakBefore = false;
      state.chapter.blanks = false;
    }

    if (spec.page && spec.page.w && spec.page.h) {
      state.page.w = spec.page.w;
      state.page.h = spec.page.h;
      var found = 'A4';
      Object.keys(PAPERS).forEach(function (k) {
        if (PAPERS[k][0] === spec.page.w && PAPERS[k][1] === spec.page.h) found = k;
      });
      state.page.paper = found;
    }
    if (spec.margin) {
      var m = spec.margin;
      var toCm = function (t) { return t === undefined || t === null ? 0 : +(t / 567).toFixed(2); };
      state.page.top = toCm(m.top);
      state.page.bottom = toCm(m.bottom);
      state.page.left = toCm(m.left);
      state.page.right = toCm(m.right);
      state.page.header = toCm(m.header);
      state.page.footer = toCm(m.footer);
    }

    renderFields();
    renderPage();
    emit();
  }

  /* ============================ 初始化 ============================ */

  function init(options) {
    options = options || {};
    onChange = options.onChange || null;
    bindTabs();
    bindFields();
    bindPage();
    bindPresets();
    renderFields();
    renderPage();
    renderPreview();
    if (options.spec) applySpec(options.spec);
  }

  root.FormatBuilder = {
    init: init,
    getSpec: getSpec,
    applySpec: applySpec,
    setActive: setActive,
    STYLE_META: STYLE_META
  };
})(window);
