/**
 * 一次跑完所有 Node 端测试：node tests/run-all.js
 * （浏览器端自测见 tests/heading-selftest.html 与 tests/preset-selftest.html）
 */
const path = require('path');
const { spawnSync } = require('child_process');

const SUITES = [
  ['test-core.js', '核心引擎：页面设置 / 段落格式 / 往返解析'],
  ['check-report-preset.js', '内置预设「课程设计报告规范」逐项核对'],
  ['test-headings.js', '标题层级识别 / 题注 / 空标题段 / 样式同步']
];

/* @xmldom/xmldom 装在受管 Node 工作区里，测试通过 NODE_PATH 找它 */
const managed = path.resolve(process.env.USERPROFILE || process.env.HOME || '', '.workbuddy/binaries/node/workspace/node_modules');
const env = Object.assign({}, process.env);
if (!env.NODE_PATH) env.NODE_PATH = managed;

let failed = 0;
for (const [file, label] of SUITES) {
  const r = spawnSync(process.execPath, [path.resolve(__dirname, file)], {
    encoding: 'utf8', env, cwd: path.resolve(__dirname, '..')
  });
  const out = (r.stdout || '') + (r.stderr || '');
  const summary = (out.match(/^\s*通过 \d+ 项，失败 \d+ 项\s*$/m) || ['(未产出结果)'])[0].trim();
  const bad = /\u2717/.test(out) || r.status !== 0;
  console.log((bad ? '\u2717 ' : '\u2713 ') + file.padEnd(24) + summary + '   ' + label);
  if (bad) {
    failed++;
    console.log(out.split('\n').filter(l => /\u2717|Error|错误/.test(l)).slice(0, 12).map(l => '    ' + l).join('\n'));
  }
}
console.log('\n' + (failed ? (failed + ' 个套件有失败项') : '全部套件通过'));
process.exit(failed ? 1 : 0);
