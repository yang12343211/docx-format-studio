/**
 * 打包出可以直接发给别人的两种形态（产物写到 dist/）：
 *
 *   dist/DOCX排版工具-单文件版.html   —— CSS/JS 全部内联，双击浏览器就能用
 *   dist/DOCX排版工具-离线版.zip      —— 解压到任意位置，双击里面的 index.html
 *
 * 用法：node tools/build-dist.js      （或 npm run dist）
 */
const fs = require('fs');
const path = require('path');

const ROOT = path.resolve(__dirname, '..');
const DIST = path.join(ROOT, 'dist');
const JSZip = require(path.join(ROOT, 'js', 'jszip.min.js'));

/* 顺序不能变：jszip 先、app 最后 */
const JS_ORDER = ['js/jszip.min.js', 'js/docx-core.js', 'js/format-builder.js', 'js/app.js'];
const ZIP_FILES = ['index.html', 'css/style.css'].concat(JS_ORDER);
const TOP_DIR = 'DOCX排版工具';

function read(rel) {
  return fs.readFileSync(path.join(ROOT, rel.split('/').join(path.sep)), 'utf8');
}

function reEscape(s) {
  return s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

/** 把 index.html 里的 css/js 外链换成内联，产出一个自包含的 html 字符串 */
function buildSingle() {
  let html = read('index.html');
  const replacements = [];

  const css = read('css/style.css');
  html = html.replace(/<link[^>]*href="css\/style\.css[^"]*"[^>]*>/,
    function () { return '<style>\n' + css + '\n</style>'; });
  replacements.push('css/style.css');

  JS_ORDER.forEach(function (rel) {
    const js = read(rel);
    /* </script 会提前终止脚本块 —— 内联前必须拦下来，不能静默产出坏文件 */
    if (/<\/script/i.test(js)) throw new Error(rel + ' 里出现了 </script，无法安全内联');
    if (/<!--/.test(js)) throw new Error(rel + ' 里出现了 <!--，会触发 HTML 的转义脚本状态');
    const pat = new RegExp('<script src="' + reEscape(rel) + '(\\?[^"]*)?"\\s*></script>');
    if (!pat.test(html)) throw new Error('index.html 里找不到对 ' + rel + ' 的引用');
    html = html.replace(pat, function () {
      return '<script>\n/* ' + rel + ' */\n' + js + '\n</script>';
    });
    replacements.push(rel);
  });

  /* 自检：内联完不该再有任何外部 js/css 引用 */
  const left = html.match(/(?:src|href)="(?!data:|#)[^"]*\.(?:js|css)[^"]*"/g);
  if (left) throw new Error('单文件版仍有外链：' + left.join(', '));

  html = html.replace('<!DOCTYPE html>',
    '<!DOCTYPE html>\n<!-- DOCX 排版工作台 · 单文件版（CSS/JS 已全部内联，双击即可使用，无需解压或联网） -->');

  return { html: html, parts: replacements };
}

async function buildZip() {
  const zip = new JSZip();
  ZIP_FILES.forEach(function (rel) {
    zip.file(TOP_DIR + '/' + rel, fs.readFileSync(path.join(ROOT, rel.split('/').join(path.sep))));
  });
  zip.file(TOP_DIR + '/使用说明.txt', fs.readFileSync(path.join(__dirname, 'dist-readme.txt'), 'utf8'));
  return await zip.generateAsync({
    type: 'nodebuffer',
    compression: 'DEFLATE',
    compressionOptions: { level: 9 }
  });
}

function kb(n) { return (n / 1024).toFixed(1) + ' KB'; }

(async function () {
  fs.mkdirSync(DIST, { recursive: true });

  const single = buildSingle();
  const singlePath = path.join(DIST, 'DOCX排版工具-单文件版.html');
  fs.writeFileSync(singlePath, single.html, 'utf8');
  console.log('✓ ' + path.relative(ROOT, singlePath) + '   ' + kb(Buffer.byteLength(single.html)) +
    '  (内联 ' + single.parts.join(' / ') + ')');

  const zipBuf = await buildZip();
  const zipPath = path.join(DIST, 'DOCX排版工具-离线版.zip');
  fs.writeFileSync(zipPath, zipBuf);
  console.log('✓ ' + path.relative(ROOT, zipPath) + '   ' + kb(zipBuf.length) +
    '  (' + (ZIP_FILES.length + 1) + ' 个文件)');
})().catch(function (e) {
  console.error('打包失败：' + ((e && e.message) || e));
  process.exit(1);
});
