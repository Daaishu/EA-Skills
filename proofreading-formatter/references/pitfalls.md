# Pitfalls · 已知问题与规避方法

构建 .docx 文件时踩过的坑，按问题表象分类。修改 `build_docx.js` 之前务必读完本文。

## 生成的 Word 文件打不开

### 表现

用户在 MS Word 中双击 .docx 文件，提示"Word 在试图打开文件时遇到错误""文件已损坏"等。但 LibreOffice 可以正常打开。

### 排查步骤

1. **先用 validate.py 校验**：`python /mnt/skills/public/docx/scripts/office/validate.py <file>.docx`。如果报"新出现的 schema 错误"，按报错信息定位。
2. **用 LibreOffice 转 PDF 试开**：`python /mnt/skills/public/docx/scripts/office/soffice.py --headless --convert-to pdf <file>.docx`。LibreOffice 容错更高；如果它也报错，问题肯定在 XML 层面。
3. **二分法定位**：把文档的 body blocks 数组对半切，分别构建两份文档，看哪一半有问题。继续二分直到找到问题 block。

### 已知原因及规避方法

#### 坑 1：docx-js 的 Table 在 MS Word 中边框冲突

**症状**：用 Table 做"照片 + 文字"并排布局，LibreOffice 能开，MS Word 打不开。

**根因**：docx-js 生成的表格 XML 中 `<w:tblBorders>` 默认是可见边框（single, size 4），而单元格级别的 `<w:tcBorders>` 设置为 none。MS Word 对这种冲突更严格。此外 docx-js 的 `BorderStyle.NONE` 生成的 XML 不带 `space` 属性，进一步触发 MS Word 的校验失败。

**规避**：不要用 Table 做任何布局。所有"并排"或"缩进块"都用独立段落+居中对齐+合适的 spacing 实现。本技能里的照片-姓名-职务组合，用三个独立的居中段落实现，不用 Table。

#### 坑 2：图片段落加 left/right 边框

**症状**：给含 ImageRun 的 Paragraph 设置了四面边框（上下左右都有），validate.py 报错 "Element '{...}w:left': This element is not expected"。

**根因**：Paragraph 的 pBdr 合法子元素只有 top/bottom/between/bar，没有 left/right。给段落加四面边框在 docx schema 里不合法。

**规避**：段落边框只用 top 和 bottom。需要四面框的场景应该用 Table（如果非要用的话），但本技能不用 Table（见坑 1）。

#### 坑 3：footer 字段写成 footer: 1134 导致 MS Word 拒绝打开

**症状**：页面边距设置时给 `margin` 对象加了 `footer: 1134`，docx 能生成，validate.py 能过，LibreOffice 能开，但 MS Word 打不开。

**根因**：某些 docx-js 版本生成的 `<w:pgMar>` 属性顺序不符合 OOXML schema，`footer` 放在某些位置时 MS Word 严格校验会拒绝。

**规避**：目前格式 A 确实需要 footer 距页底距离，但要写成独立的 section property 而不是塞进 margin 对象。或者直接省略 footer 距离，用默认值。

## 中文字符在 JavaScript 字符串中引发语法错误

### 表现

写脚本时在 JS 字符串里直接拼中文，运行时报 `SyntaxError: Invalid or unexpected token` 或 `Unexpected string`。

### 根因

JavaScript 字符串字面量的界定符是 `"` `'` 和反引号。中文"智能引号"（U+201C、U+201D）在视觉上和半角引号很像但是不同字符，编辑器有时会自动替换。破折号"——"（U+2014）在紧贴英文符号时也会引发解析歧义。

### 规避

**所有文本内容都通过 JSON 文件传入 `build_docx.js`**，不要在脚本代码里嵌入中文字符串。JSON 解析器天然兼容所有 Unicode 字符，包括智能引号、破折号、省略号。

本技能的 `build_docx.js` 只做渲染，不包含任何中文字符串常量（除了字体名）。

## JSON 文件里中文引号导致 JSON 解析失败

### 表现

`Error: failed to parse JSON: Expected ',' or '}' after property value` 类错误。

### 根因

JSON 字符串用 `"` 界定。如果文本内需要出现引号，只能是转义的 `\"` 或者使用**中文全角引号** `"` `"`（U+201C、U+201D）。直接嵌半角引号会导致 JSON 结构破坏。

### 规避

在 `content.json` 里写中文内容时，所有引号都用全角 `"` `"` 而非半角 `"`，这样既符合中文排版习惯又不会破坏 JSON。例如：

```json
"text": "在教育部的指导下，开展\u201c因材施教\u201d项目"
```

或直接写 Unicode 字符：

```json
"text": "在教育部的指导下，开展"因材施教"项目"
```

（注意外层仍是半角 `"`）。

## 字体不正确或不显示

### 表现

生成的文档在 Word 中打开，发现字体变成了宋体或系统默认字体。

### 根因

1. 系统未安装指定字体（例如 Linux 环境没有"方正小标宋简体"）。
2. `TextRun` 的 `font` 属性只设了 `name` 没设 `eastAsia`，中文字符会回退到默认。

### 规避

1. **字体名必须同时设置 `name` 和 `eastAsia`**：
   ```javascript
   font: { name: "仿宋", eastAsia: "仿宋" }
   ```
2. 接受字体降级：方正小标宋简体在非 Windows 办公环境会自动降级到华文中宋或宋体，这是预期行为。打印排版的视觉差异可以接受。
3. 不要假设运行环境装了政府公文字体。脚本里可以同时指定多个字体作为降级链。

## 标题一行放不下，Word 自动折行很难看

### 表现

长标题如"关于申请 2026 年度城市绿化专项资金的请示"，Word 自动在某个词之间断行，断句位置语义上不合理。

### 规避

在 content.json 里用 `titleLines` 数组显式指定每行文本：

```json
"titleLines": [
  "关于申请 2026 年度城市绿化",
  "专项资金的请示"
]
```

格式 A 和 B 都支持。断句原则见 `format-specs.md` 格式 A "标题断行规则"一节。

## 页眉 Logo 图片定位异常

### 表现

早期尝试给议程格式 C 在页眉加科大讯飞 Logo，发现生成的 docx 在某些环境里 Logo 位置漂移、边框线错位、或页眉整体空白。

### 规避

**本技能刻意不嵌入任何 Logo**。原因有三：
1. 不同机构 Logo 尺寸、装饰线样式、位置要求都不一样，做通用模板不现实
2. 固定 Logo 在 docx 页眉中的绝对定位有多个坑（inline vs anchor、wp:wrap 属性、与边框线的相对位置）
3. 使用者后期自己在 Word 里插入 Logo 更灵活可控

如果确实要嵌入 Logo，建议的做法是让使用者在生成的 docx 里自行添加到页眉，而不是写死在构建脚本里。

## Linux 环境缺少中文字体导致 LibreOffice 渲染异常

### 表现

在 Linux 容器中用 LibreOffice 把 docx 转 PDF 时，中文变成乱码或全是方块。

### 规避

这是 LibreOffice 字体配置问题，和 docx 文件本身无关。在受控环境（如 Claude 的 computer use 环境）通常不是问题；在自己部署时需要安装 `fonts-noto-cjk` 或类似包。

文档本身生成没有问题——用 MS Word 或 WPS 打开是正常的。

## 小结：修改 build_docx.js 时的 checklist

改完脚本后跑一遍这几件事：

1. 三个 example 都重新构建一遍，检查退出码
2. 对每个 output.docx 跑 `validate.py`
3. 用 LibreOffice 转一份 PDF，肉眼看字体、断行、间距有没有明显问题
4. 如果有条件，用 MS Word 或 WPS 双击打开试试
5. 改了字体/字号/间距，把 format-specs.md 对应章节也更新了

本技能走过的所有 v1 → v7 迭代，绝大多数问题都是在第 3、4 步暴露的。validate.py 通过不代表 MS Word 能打开。
