/**
 * 用真实文档验证「关闭文档网格」：
 * 处理用户的《中期(1)-已排版 (1)-已排版.docx》——它原来开着 w:type="lines"。
 *
 * 运行： node tests/grid-real.js <输入.docx> [输出.docx]
 */
const path = require('path');
const fs = require('fs');
const { DOMParser, XMLSerializer } = require('@xmldom/xmldom');

global.DOMParser = DOMParser;
global.XMLSerializer = XMLSerializer;
global.JSZip = require(path.resolve(__dirname, '../js/jszip.min.js'));

const DocxCore = require(path.resolve(__dirname, '../js/docx-core.js'));

const SRC = process.argv[2];
const OUT = process.argv[3] || path.join(require('os').tmpdir(), 'grid-off-out.docx');

(async function () {
  const buf = fs.readFileSync(SRC);
  const ab = buf.buffer.slice(buf.byteOffset, buf.byteOffset + buf.byteLength);

  const zip0 = await global.JSZip.loadAsync(ab);
  const before = await zip0.file('word/document.xml').async('string');
  const bGrid = before.match(/<w:docGrid[^>]*>/g) || [];
  console.log('输入:', path.basename(SRC));
  console.log('  处理前 docGrid :', bGrid.join(' ; ') || '(无)');

  const res = await DocxCore.applyFormat(ab, DocxCore.BUILTIN_PRESETS['report-cn'].spec, {});
  console.log('  报告 docGridOff :', res.report.docGridOff);
  console.log('  报告 原网格类型 :', res.report.docGridBefore || '(无)');

  const zip1 = await global.JSZip.loadAsync(res.data);
  const after = await zip1.file('word/document.xml').async('string');
  const aGrid = after.match(/<w:docGrid[^>]*>/g) || [];
  console.log('  处理后 docGrid :', aGrid.join(' ; ') || '(无)');

  const ok1 = aGrid.length > 0 && aGrid.every(function (g) { return !/w:type=/.test(g); });
  const ok2 = /w:linePitch="312"/.test(after);
  console.log('\n  ' + (ok1 ? 'OK ' : 'FAIL') + ' 所有 docGrid 的 w:type 已移除（= 无网格）');
  console.log('  ' + (ok2 ? 'OK ' : 'FAIL') + ' linePitch 等无关属性保留');

  fs.writeFileSync(OUT, Buffer.from(new Uint8Array(res.data)));
  console.log('\n产物:', OUT, '(' + fs.statSync(OUT).size + ' bytes)');

  const keys = Object.keys(res.report).filter(function (k) {
    const v = res.report[k];
    return typeof v === 'number' && v > 0;
  });
  console.log('报告非零项:', keys.map(function (k) { return k + '=' + res.report[k]; }).join(', '));

  process.exit(ok1 && ok2 ? 0 : 1);
})().catch(function (e) {
  console.error('异常:', (e && e.stack) || e);
  process.exit(1);
});
