/* ==========================================================================
   文档格式自动排版 — 页面交互
   ========================================================================== */
(function () {
  'use strict';

  var $ = function (id) { return document.getElementById(id); };

  var state = {
    mode: 'upload',     /* upload = 用模板文档解析；manual = 手动设置 */
    tplFile: null,
    tplBuffer: null,
    tplSpec: null,
    docFile: null,
    docBuffer: null,
    result: null,
    resultName: ''
  };

  /** 当前生效的格式规范 */
  function currentSpec() {
    if (state.mode === 'manual') {
      return (window.FormatBuilder && window.FormatBuilder.getSpec) ? FormatBuilder.getSpec() : null;
    }
    return state.tplSpec;
  }

  var MAX_SIZE = 30 * 1024 * 1024;

  /* ---------------- 工具 ---------------- */

  function esc(s) {
    return String(s === undefined || s === null ? '' : s)
      .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;').replace(/'/g, '&#39;');
  }

  function fmtSize(bytes) {
    if (bytes < 1024) return bytes + ' B';
    if (bytes < 1024 * 1024) return (bytes / 1024).toFixed(1) + ' KB';
    return (bytes / 1024 / 1024).toFixed(2) + ' MB';
  }

  function readBuffer(file) {
    if (file.arrayBuffer) return file.arrayBuffer();
    return new Promise(function (resolve, reject) {
      var fr = new FileReader();
      fr.onload = function () { resolve(fr.result); };
      fr.onerror = function () { reject(fr.error || new Error('读取文件失败')); };
      fr.readAsArrayBuffer(file);
    });
  }

  function checkFile(file) {
    if (!file) return '没有选择文件';
    var name = (file.name || '').toLowerCase();
    if (!/\.docx$/.test(name)) {
      if (/\.(doc|wps|wpt)$/.test(name)) {
        return '这是老格式（.doc / .wps），浏览器无法直接读取。请在 WPS 或 Word 里「另存为」成 .docx 再试。';
      }
      if (/\.(dotx|docm)$/.test(name)) {
        return '请使用普通的 .docx 文件（.dotx 模板 / .docm 宏文档请先另存为 .docx）。';
      }
      return '只支持 .docx 格式的文件，你选的是「' + file.name + '」。';
    }
    if (file.size > MAX_SIZE) {
      return '文件太大了（' + fmtSize(file.size) + '），请先精简文档内容再试。';
    }
    if (file.size === 0) return '这是一个空文件。';
    return null;
  }

  function showError(msg) {
    var box = $('err-box');
    if (!msg) { box.hidden = true; box.innerHTML = ''; return; }
    box.hidden = false;
    box.innerHTML = '<b>处理不了这个文件</b>' + esc(msg);
    box.scrollIntoView({ behavior: 'smooth', block: 'center' });
  }

  function baseName(name) {
    return String(name || 'document').replace(/\.docx$/i, '');
  }

  /* ---------------- 状态同步 ---------------- */

  function refreshRunBtn() {
    var ready = !!(currentSpec() && state.docBuffer);
    var btn = $('run');
    btn.disabled = !ready || btn.dataset.busy === '1';
    $('run-hint').textContent = ready
      ? '一切就绪，点击开始'
      : (!currentSpec()
        ? (state.mode === 'manual' ? '请上传要处理的文档' : '请先上传模板，或切到「手动设置格式」')
        : '请上传要处理的文档');
  }

  function setChip(kind, file) {
    var chip = $('chip-' + kind);
    if (!file) { chip.hidden = true; return; }
    $('name-' + kind).textContent = file.name;
    $('meta-' + kind).textContent = fmtSize(file.size);
    chip.hidden = false;
  }

  /* ---------------- 规范渲染 ---------------- */

  function labelHtml(label) {
    var m = String(label).match(/^标题 (\d)$/);
    if (m) return '<span class="lv-badge">H' + m[1] + '</span>标题 ' + m[1];
    return esc(label);
  }

  function renderSpec(spec) {
    var d = DocxCore.describeSpec(spec);

    var pagesHtml = d.page.map(function (b) {
      return '<div class="page-box"><div class="k">' + esc(b.k) + '</div><div class="v">' + esc(b.v) + '</div></div>';
    }).join('');
    if (!pagesHtml) {
      pagesHtml = '<div class="page-box"><div class="k">页面设置</div><div class="v">模板中未读到，将保持不变</div></div>';
    }
    $('spec-pages').innerHTML = pagesHtml;

    var rows = d.styles.map(function (s) {
      return '<tr>' +
        '<td class="row-label">' + labelHtml(s.label) + '</td>' +
        '<td>' + esc(s.font) + '</td>' +
        '<td>' + esc(s.size) + '</td>' +
        '<td>' + esc(s.bold) + '</td>' +
        '<td>' + esc(s.align) + '</td>' +
        '<td>' + esc(s.spacing) + '</td>' +
        '<td>' + esc(s.before) + ' / ' + esc(s.after) + '</td>' +
        '<td>' + esc(s.indent) + '</td>' +
        '</tr>';
    }).join('');

    if (!rows) {
      rows = '<tr><td colspan="8" style="text-align:center;color:#8d9bb0;padding:22px">' +
        '模板中没有读到可用的样式信息，将只能套用页面设置。</td></tr>';
    }
    $('spec-tbody').innerHTML = rows;

    /* 提示 */
    var warns = [];
    if (spec.autoNormal) {
      warns.push('模板里没有定义「正文」样式，正文格式将沿用文档自身的默认设置。');
    }
    var hasHeading = false;
    for (var i = 1; i <= 9; i++) if (spec.headings && spec.headings[i]) hasHeading = true;
    if (!hasHeading) {
      warns.push('模板里没有识别到任何标题样式。如果你的模板是靠手工加粗、调大字号做的标题，程序读不到格式；建议先在模板里用「标题 1 / 标题 2」样式刷一遍。');
    }

    var warnEl = $('spec-warn');
    if (warns.length) {
      warnEl.hidden = false;
      warnEl.innerHTML = warns.map(function (w) { return '• ' + esc(w); }).join('<br>');
    } else {
      warnEl.hidden = true;
    }

    $('card-spec').hidden = false;
  }

  /* ---------------- 上传处理 ---------------- */

  function bindDropzone(dz, input, onFile) {
    dz.addEventListener('click', function () { input.click(); });

    input.addEventListener('change', function () {
      if (input.files && input.files[0]) onFile(input.files[0]);
      input.value = '';
    });

    ['dragenter', 'dragover'].forEach(function (ev) {
      dz.addEventListener(ev, function (e) {
        e.preventDefault(); e.stopPropagation();
        dz.classList.add('over');
      });
    });
    ['dragleave', 'drop'].forEach(function (ev) {
      dz.addEventListener(ev, function (e) {
        e.preventDefault(); e.stopPropagation();
        if (ev === 'dragleave' && dz.contains(e.relatedTarget)) return;
        dz.classList.remove('over');
      });
    });
    dz.addEventListener('drop', function (e) {
      var dt = e.dataTransfer;
      if (dt && dt.files && dt.files[0]) onFile(dt.files[0]);
    });
  }

  async function handleTemplate(file) {
    var err = checkFile(file);
    if (err) { showError(err); return; }

    showError(null);
    state.tplFile = file;
    setChip('template', file);

    var dz = $('dz-template');
    dz.classList.add('loading');
    $('meta-template').textContent = '正在读取格式…';

    try {
      state.tplBuffer = await readBuffer(file);
      var spec = await DocxCore.analyzeTemplate(state.tplBuffer, { fileName: file.name });
      state.tplSpec = spec;
      $('meta-template').textContent = fmtSize(file.size) + ' · 已读取格式规范';
      renderSpec(spec);
      refreshRunBtn();
      $('card-spec').scrollIntoView({ behavior: 'smooth', block: 'start' });
    } catch (e) {
      console.error(e);
      state.tplSpec = null;
      state.tplBuffer = null;
      setChip('template', null);
      showError((e && e.message) || '读取模板失败');
      refreshRunBtn();
    } finally {
      dz.classList.remove('loading');
    }
  }

  async function handleDoc(file) {
    var err = checkFile(file);
    if (err) { showError(err); return; }

    showError(null);
    state.docFile = file;
    setChip('doc', file);

    var dz = $('dz-doc');
    dz.classList.add('loading');
    $('meta-doc').textContent = '正在读取…';

    try {
      state.docBuffer = await readBuffer(file);
      $('meta-doc').textContent = fmtSize(file.size) + ' · 已就绪';
      $('card-result').hidden = true;
      refreshRunBtn();
    } catch (e) {
      console.error(e);
      state.docBuffer = null;
      setChip('doc', null);
      showError('读取文档失败：' + ((e && e.message) || '未知错误'));
      refreshRunBtn();
    } finally {
      dz.classList.remove('loading');
    }
  }

  /* ---------------- 执行 ---------------- */

  async function run() {
    var spec = currentSpec();
    if (!spec || !state.docBuffer) return;

    var btn = $('run');
    btn.dataset.busy = '1';
    btn.disabled = true;
    btn.querySelector('.run-label').textContent = '正在处理…';
    btn.querySelector('.run-spin').hidden = false;
    $('run-hint').textContent = '正在逐段套用格式，请稍候';
    showError(null);

    var options = {
      applyPage: $('opt-page').checked,
      applyHeading: $('opt-heading').checked,
      applyBody: $('opt-body').checked,
      applyCaption: $('opt-caption').checked,
      skipTableIndent: $('opt-tableindent').checked,
      heuristic: $('opt-heuristic').checked,
      clearManual: $('opt-clear').checked,
      removeRules: $('opt-rules') ? $('opt-rules').checked : true
    };

    /* 让浏览器有机会把 loading 画面画出来 */
    await new Promise(function (r) { setTimeout(r, 30); });

    try {
      var res = await DocxCore.applyFormat(state.docBuffer, spec, options);
      state.result = res.data;
      state.resultName = baseName(state.docFile.name) + '-已排版.docx';
      renderResult(res.report);
    } catch (e) {
      console.error(e);
      showError((e && e.message) || '套用格式时出错');
    } finally {
      btn.dataset.busy = '0';
      btn.querySelector('.run-label').textContent = '开始套用格式';
      btn.querySelector('.run-spin').hidden = true;
      refreshRunBtn();
    }
  }

  function renderResult(report) {
    var headingTotal = 0, headingDetail = [];
    for (var lv = 1; lv <= 9; lv++) {
      var n = report.headings[lv] || 0;
      if (n) { headingTotal += n; headingDetail.push('标题 ' + lv + ' × ' + n); }
    }

    var stats = [
      { n: report.paragraphs, l: '处理段落总数' },
      { n: headingTotal, l: '识别为标题' },
      { n: report.body, l: '正文段落' },
      { n: report.caption, l: '题注' }
    ];
    $('stats').innerHTML = stats.map(function (s) {
      return '<div class="stat"><div class="n">' + s.n + '</div><div class="l">' + esc(s.l) + '</div></div>';
    }).join('');

    var sub = [];
    if (report.pageApplied) sub.push('页面设置已套用');
    if (headingDetail.length) sub.push(headingDetail.join('、'));
    $('result-sub').textContent = sub.join(' · ') || '处理完成';

    var notes = [];
    if (report.levelFixes > 0) {
      notes.push('原文档里有 <b>' + report.levelFixes + '</b> 处<b>标题样式和实际层级对不上</b>（比如把「第1章」套成了「标题 2」、把「1.1」套成了「标题 3」）。这种情况如果照样式套，整篇标题会集体错位一级，所以已按文字编号（第X章 → 一级、X.Y → 二级、X.Y.Z → 三级）纠正过来，段落的样式名也一并改挂到正确的「标题 1 / 2 / 3」上——打开 Word 的样式库、导航窗格看到的都是对的。');
    }
    if (report.stylesReused && report.stylesReused.length) {
      notes.push('文档原有的样式已按规范重写定义：' + report.stylesReused.map(function (s) { return esc(s); }).join('、') + '。');
    }
    if (report.stylesCreated && report.stylesCreated.length) {
      notes.push('文档里原本缺这些样式，已新建并套用：' + report.stylesCreated.map(function (s) { return esc(s); }).join('、') + '。');
    }
    if (report.removedRules > 0) {
      notes.push('清掉了 <b>' + report.removedRules + '</b> 条<b>遗留的水平线</b>。这类灰横杠是网页或 AI 对话里的「---」分隔线粘进 Word 后留下的（在 XML 里是带 <code>o:hr</code> 标记的图形），不属于文档内容；如果你确实需要它们，把下面的「清掉遗留的水平线」取消勾选再处理一次。');
    }
    if (report.chapterBlanks > 0 || report.chapterBreaks > 0) {
      var cp = [];
      if (report.chapterBreaks > 0) cp.push('给 <b>' + report.chapterBreaks + '</b> 个章标题加了<b>分页</b>（每章另起一页）');
      if (report.chapterBlanks > 0) cp.push('为 <b>' + report.chapterBlanks + '</b> 处章标题备好了<b>一级标题格式的空行</b>（标题上下各一行）');
      notes.push(cp.join('、') + '。这些空行的字体字号跟一级标题一致（小二黑体），所以空出来的行高和标题对得上；原文档在标题上下本来就留了空行的，会就地改成这个格式，不会再叠出第二行。文档最开头的那一章前面没有内容，不会再加分页，免得空出一整页。');
    }
    if (report.demotedHeadings > 0) {
      notes.push('有 <b>' + report.demotedHeadings + '</b> 处<b>空段落原本套着标题样式</b>（常见于拿空行垫版面的文档）。这类空行如果照着样式套，会变成「小二黑体居中」把版式撑乱，所以已按正文处理并摘掉了标题样式，Word 导航窗格和自动目录里也不会再冒出空条目。');
    }
    if (report.heuristicHits > 0) {
      notes.push('其中 <b>' + report.heuristicHits + '</b> 处标题是按编号规律自动认出来的（「第X章」→ 一级、「1.1」→ 二级、「1.1.1」→ 三级）。这部分属于推断，建议打开文档核对一下层级对不对。');
    }
    if (headingTotal === 0) {
      notes.push('这次<b>没有识别到任何标题</b>。如果文档里确实有标题，说明它们既没用 Word 标题样式，编号也不符合常见规律——可以手动把标题刷成「标题 1 / 标题 2」再处理一次。');
    }
    if (report.warnings && report.warnings.length) {
      notes.push('处理中有 ' + report.warnings.length + ' 处小问题：' + esc(report.warnings.slice(0, 3).join('；')));
    }

    var noteEl = $('result-note');
    if (notes.length) {
      noteEl.hidden = false;
      noteEl.innerHTML = notes.map(function (t) { return '• ' + t; }).join('<br>');
    } else {
      noteEl.hidden = true;
    }

    $('card-result').hidden = false;
    $('card-result').scrollIntoView({ behavior: 'smooth', block: 'center' });
  }

  function download() {
    if (!state.result) return;
    var blob = new Blob([state.result], { type: DocxCore.DOCX_MIME });
    var url = URL.createObjectURL(blob);
    var a = document.createElement('a');
    a.href = url;
    a.download = state.resultName;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    setTimeout(function () { URL.revokeObjectURL(url); }, 2000);
  }

  /* ---------------- 初始化 ---------------- */

  document.addEventListener('DOMContentLoaded', function () {
    bindDropzone($('dz-template'), $('file-template'), handleTemplate);
    bindDropzone($('dz-doc'), $('file-doc'), handleDoc);

    Array.prototype.forEach.call(document.querySelectorAll('.chip-x'), function (btn) {
      btn.addEventListener('click', function () {
        var kind = btn.dataset.clear;
        if (kind === 'template') {
          state.tplFile = null; state.tplBuffer = null; state.tplSpec = null;
          setChip('template', null);
          $('card-spec').hidden = true;
        } else {
          state.docFile = null; state.docBuffer = null;
          setChip('doc', null);
          $('card-result').hidden = true;
        }
        showError(null);
        refreshRunBtn();
      });
    });

    $('run').addEventListener('click', run);
    $('download').addEventListener('click', download);

    $('use-preset').addEventListener('click', function () {
      var preset = DocxCore.BUILTIN_PRESETS['thesis-cn'];
      state.tplSpec = JSON.parse(JSON.stringify(preset.spec));
      state.tplSpec.source = { file: preset.label, generatedAt: new Date().toISOString() };
      state.tplFile = null; state.tplBuffer = null;

      $('name-template').textContent = '内置规范 · ' + preset.label;
      $('meta-template').textContent = '已载入，无需上传模板';
      $('chip-template').hidden = false;

      showError(null);
      renderSpec(state.tplSpec);
      refreshRunBtn();
      $('card-spec').scrollIntoView({ behavior: 'smooth', block: 'start' });
    });

    /* 全局拖放兜底：防止拖到页面空白处时浏览器直接打开文件 */
    ['dragover', 'drop'].forEach(function (ev) {
      window.addEventListener(ev, function (e) {
        if (!e.target.closest || !e.target.closest('.dropzone')) e.preventDefault();
      });
    });

    /* ---------- 手动设置格式模块 ---------- */
    if (window.FormatBuilder) {
      FormatBuilder.init({
        onChange: function () {
          if (state.mode === 'manual') refreshRunBtn();
        }
      });
    }

    /* ---------- 两种模式切换 ---------- */
    var modeTabs = document.querySelectorAll('.mode-tab');
    Array.prototype.forEach.call(modeTabs, function (tab) {
      tab.addEventListener('click', function () {
        var mode = tab.dataset.mode;
        if (mode === state.mode) return;
        state.mode = mode;
        Array.prototype.forEach.call(modeTabs, function (t) {
          t.classList.toggle('active', t === tab);
        });
        $('pane-upload').hidden = (mode !== 'upload');
        $('pane-manual').hidden = (mode !== 'manual');
        /* 手动模式自带实时预览，就不再重复展示解析出来的规范表 */
        $('card-spec').hidden = (mode !== 'upload') || !state.tplSpec;
        showError(null);
        refreshRunBtn();
      });
    });

    refreshRunBtn();
  });
})();
