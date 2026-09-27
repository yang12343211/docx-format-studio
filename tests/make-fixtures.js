/** 生成浏览器自测用的静态夹具：node tests/make-fixtures.js */
const fs = require('fs');
const path = require('path');
const { buildDocx, DOC_ITEMS, SHIFTED_ITEMS, RULE_ITEMS, CHAPTER_ITEMS, CHAPTER_FIRST_ITEMS, PRE_BLANKED_ITEMS, PLAIN_ITEMS } = require('./lib/build-docx.js');

(async function () {
  const dir = path.join(__dirname, 'fixtures');
  fs.mkdirSync(dir, { recursive: true });
  const jobs = [
    ['headings.docx', DOC_ITEMS],
    ['shifted-styles.docx', SHIFTED_ITEMS],
    ['rules.docx', RULE_ITEMS],
    ['chapters.docx', CHAPTER_ITEMS],
    ['chapters-first.docx', CHAPTER_FIRST_ITEMS],
    ['pre-blanked.docx', PRE_BLANKED_ITEMS],
    ['plain.docx', PLAIN_ITEMS]
  ];
  for (const [name, items] of jobs) {
    const bytes = await buildDocx(items);
    const out = path.join(dir, name);
    fs.writeFileSync(out, Buffer.from(bytes));
    console.log('已生成 ' + out + ' (' + bytes.length + ' 字节)');
  }
})().catch(e => { console.error(e); process.exit(1); });
