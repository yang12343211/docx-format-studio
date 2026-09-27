# docx-format-studio · DOCX 排版工作台

一个**纯前端**的 Word 文档批量排版工具。把排版规范（字体、字号、行距、缩进、页边距、标题层级）一次性套用到已有的 `.docx` 上，
自动纠正"标题样式和实际层级对不上"、"标题继承了正文首行缩进"、"网页粘过来的水平线"这类常见毛病。

所有处理都在**浏览器本地**完成（JSZip 直接读写 OOXML），文档不上传任何服务器。

## 功能

**排版规范（Spec）来源**
- 内置预设：`中文学术论文通用规范`、`课程设计报告规范`（后者额外带"一级标题章版式"）
- 手动设置：逐级调字体/字号/字形/对齐/行距/段前段后/首行缩进
- 上传模板文档，从模板里**反解**出规范（读页边距、各级标题样式、题注样式）
- 上传目标文档，从目标文档里**推断**出它自己原本的规范

**识别与套用**
- 标题层级识别：优先看 `outlineLvl` + `w:pStyle`，无样式时按编号规律兜底
  （`第X章` → 一级、`X.Y` → 二级、`X.Y.Z` → 三级；`一、`/`（一）`/`1.` 这类弱编号不覆盖已知样式层级）
- **样式整体错位一级**的文档（如「第1章」被套成了"标题 2"）会被编号规律纠正回来
- 空段落套着标题样式（作者拿它当空行用）→ 摘掉标题样式按正文处理，避免撑乱版式、避免 Word 导航窗格里冒出空条目
- 题注识别：`表X-Y` / `图X-Y`，样式表里没有"题注"样式时也能认
- 同步纠正 `w:pStyle` 与 `styles.xml` 里的样式定义，样式表缺失的样式会**新建**并挂上
- 「**一级标题（章）版式**」：每章另起一页、标题上下各空一行，且两个空行套一级标题的字体字号
  （空行高度由段落标记的字符格式决定，按正文算会比标题矮一截）
- 清理从网页/AI 对话粘进来的 markdown `---`（在 Word 里是带 `o:hr` 标记的灰横杠），可开关

**输出**：处理完给出逐项结果的说明（识别到多少级标题、新建/改写了哪些样式、纠正了几处、清掉几条水平线等），并下载排版后的 `.docx`。

## 在线使用

**<https://yang12343211.github.io/docx-format-studio/>** —— 点开就能用，手机也能开，不需要安装任何东西。

备用地址：<https://docx-format-studio.app.workbuddy.host/>

### 部署到 GitHub Pages

本仓库已经开启了 Pages（`main` 分支 / 根目录），每次推送到 `main` 会自动重新构建。
如果要在别处重新部署：这是个纯静态站，仓库根目录就是站点根目录（`index.html` 在最外层），**不需要任何构建步骤**——
仓库 **Settings → Pages** → Source 选 `Deploy from a branch` → Branch 选 `main` + `/ (root)` → 保存，
等一分多钟访问 `https://<用户名>.github.io/<仓库名>/`。

仓库里带了 `.nojekyll`，这样 GitHub Pages 不会用 Jekyll 处理、下划线开头的文件（如 `tests/_validate-docx.js`）也不会被跳过。

线上还能直接跑浏览器自检页（它们用的都是相对路径，放在子路径下也能用）：

```
https://yang12343211.github.io/docx-format-studio/tests/heading-selftest.html
https://yang12343211.github.io/docx-format-studio/tests/preset-selftest.html
```

## 分享给别人

三种方式，按方便程度排：

| 方式 | 别人需要做什么 |
| --- | --- |
| **发在线链接** | 点开就用，手机也能开。不需要装任何东西 |
| **发单文件 HTML** | 收到一个 `.html` 文件，双击用浏览器打开即可（CSS/JS 全内联，不依赖其它文件、不联网） |
| **发离线压缩包** | 解压到任意位置，双击里面的 `index.html` |

后两种都是纯前端、断网可用，文档不上传服务器。自己打包：

```bash
npm run dist     # 产物写到 dist/
#   dist/DOCX排版工具-单文件版.html   把 index.html + css + js 内联成一个 .html
#   dist/DOCX排版工具-离线版.zip      解压后双击里面的 index.html
```

打包脚本会做自检：内联后若还存在任何外部 `js`/`css` 引用、或源码里出现会破坏内联的
`</script`、`<!--` 序列，就直接报错退出，不会静默产出坏文件。

## 本地运行

纯静态站点，任意静态服务器即可：

```bash
python -m http.server 8899 --bind 127.0.0.1   # 或 npm run serve
# 浏览器打开 http://127.0.0.1:8899/index.html
```

## 测试

```bash
npm install     # 只装一个测试依赖 @xmldom/xmldom
npm test        # 跑全部套件
```

三个 Node 套件（直接对引擎断言，不经过界面）：

| 套件 | 覆盖 |
| --- | --- |
| `tests/test-core.js` | 页面设置 / 段落格式 / 往返解析 |
| `tests/check-report-preset.js` | 内置预设「课程设计报告规范」逐项核对 |
| `tests/test-headings.js` | 标题层级识别 / 题注 / 空标题段 / 样式同步 / 缩进继承 / 水平线清理 / 章版式 |

两个**浏览器自检页**（在真实浏览器里点真实按钮、断言真实产物，覆盖界面到引擎的接线）：

- `tests/heading-selftest.html`
- `tests/preset-selftest.html`

用 headless Chrome/Edge 跑：

```bash
npm run serve
# 另开一个终端
msedge --headless=new --dump-dom http://127.0.0.1:8899/tests/heading-selftest.html | grep DONE
```

测试夹具（`tests/fixtures/*.docx`）由 JSZip 现场构造，用 `npm run fixtures` 重新生成。

处理一个真实文档并打印结构报告：

```bash
node tests/format-one.js 输入.docx 输出.docx              # 含章版式
node tests/format-one.js 输入.docx 输出.docx --no-chapter # 关掉章版式
```

## 目录结构

```
index.html              界面
css/style.css           样式
js/docx-core.js         排版引擎（OOXML 读写、识别、套用、章版式）
js/format-builder.js    规范编辑器（预设/手动/模板反解）与预览
js/app.js               上传、执行、结果说明
js/jszip.min.js         vendored 第三方库
tests/                  Node 测试套件 + 浏览器自检页 + 夹具生成
tools/build-dist.js     打包脚本（单文件版 / 离线压缩包）
samples/                输入样例（模板/目标文档）
```

## 实现要点

- **两遍架构**：先写段落直接格式并同步 `styles.xml`、重映射 `pStyle`；再把章版式（分页 + 空行）作为独立一遍跑，
  因为章标题是靠"段落里已经写好的 `outlineLvl` + `pStyle`"认出来的，不能只看文字像不像章标题。
- **删掉 `w:ind` ≠ 没有缩进**。中文 Word 里标题样式通常 `basedOn` 正文，正文带「首行缩进 2 字符」时，
  把标题段落自己的 `w:ind` 删掉，Word 会顺着样式链把缩进继承回来。必须**显式写 `firstLineChars="0"`**，
  段落和样式定义两处都要写。
- **分页符挂在"标题上面那个空行"上**，而不是标题自己身上——挂在标题上的话前置空行会留在上一页页尾、标题单独跑到新页顶。
- **元素顺序**：`w:pPr` / `w:rPr` / `w:style` 的子元素都必须按 OOXML schema 顺序插入，不能 `appendChild` 乱拼。
  测试里有一道产物校验器专门查这个，并且和源文件对比，只报"引擎新引入"的顺序问题。
- 处理段落前必须把 `getElementsByTagName` 的结果**拷成数组**：浏览器返回的是活 `NodeList`，
  往文档里插空段落会让它变长、遍历失控。

## 已知限制

- 只支持 `.docx`（OOXML）。老式 `.doc` 是 OLE 复合文档，不是 zip，前端读不了。
- 标题层级识别是启发式的，编号规律不常规的文档建议处理完核对一遍层级。
- 不做目录域（TOC）更新——那是 Word 打开时自己刷新的。

## License

MIT
