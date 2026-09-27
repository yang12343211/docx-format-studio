/**
 * 检视一份 .docx 的段落结构（排查"标题识别不对"时先跑这个）
 *
 *   node tests/inspect-docx.js "C:/path/to/报告.docx"
 *
 * 输出每段的：序号 | pStyle(styleId) | outlineLvl | 是否带编号 | 文字前 44 字
 * pStyle 为 "-" 表示该段没有显式段落样式（标题只能靠 outlineLvl / 编号规律认）。
 * 想知道 styleId 对应哪个样式名，再看 word/styles.xml 里的 w:name。
 */
const fs=require('fs'), path=require('path');
const {DOMParser}=require('@xmldom/xmldom');
global.DOMParser=DOMParser;
global.JSZip=require(path.resolve(__dirname,'../js/jszip.min.js'));
const file=process.argv[2];
(async()=>{
  const b=fs.readFileSync(file);
  const zip=await global.JSZip.loadAsync(new Uint8Array(b.buffer.slice(b.byteOffset,b.byteOffset+b.byteLength)));
  const xml=await zip.file('word/document.xml').async('string');
  const doc=new DOMParser().parseFromString(xml,'text/xml');
  const ps=doc.getElementsByTagName('w:p');
  for(let i=0;i<ps.length;i++){
    const p=ps[i];
    let txt='';
    const ts=p.getElementsByTagName('w:t');
    for(let j=0;j<ts.length;j++) txt+=(ts[j].textContent||'');
    const st=p.getElementsByTagName('w:pStyle')[0];
    const ol=p.getElementsByTagName('w:outlineLvl')[0];
    const npr=p.getElementsByTagName('w:numPr')[0];
    const sid=st?st.getAttribute('w:val'):'-';
    const lvl=ol?ol.getAttribute('w:val'):'-';
    const num=npr?'num':'';
    const t=txt.trim();
    if(!t) continue;
    console.log([i,sid,'outline='+lvl,num,t.slice(0,44)].join(' | '));
  }
})().catch(e=>{console.error(e);process.exit(1)});
