# Content JSON Schema

所有三种格式共用一个内容 JSON 结构。`build_docx.js` 读取此 JSON 并按指定格式渲染。

## 顶层结构

```json
{
  "title": "文档主标题",
  "titleLines": ["可选", "如需手动控制标题断行，填此字段"],
  "author": "可选，作者署名（格式 B 常用）",
  "date": "可选，日期字符串",
  "meta": {
    "time": "可选，用于格式 C 的时间",
    "location": "可选，用于格式 C 的地点",
    "agendaItem": "可选，用于格式 C 的议程条目"
  },
  "body": [
    { "type": "...", ... }
  ],
  "signature": {
    "org": "可选，落款单位",
    "date": "可选，落款日期"
  },
  "nameList": {
    "groups": [
      { "title": "单位/组别名称", "members": [{ "name": "...", "title": "..." }] }
    ]
  }
}
```

说明：

- `title` 必填；`titleLines` 优先级高于 `title`，用于公文标题主动断行（见 format-specs.md 格式 A 的标题断行规则）
- `body` 是正文段落数组，每项是一个 block 对象
- `signature` 仅格式 A 使用
- `nameList` 仅格式 C 使用，渲染在名单页
- `meta` 仅格式 C 使用，渲染在议程封面页

## Body Block 类型

### `h1` - 一级标题

```json
{ "type": "h1", "text": "一、标题内容" }
```

- 格式 A：黑体三号，首行缩进 2 字符
- 格式 B：黑体加粗三号，左对齐不缩进
- 格式 C：一般不用

### `h2` - 二级标题

```json
{ "type": "h2", "text": "（一）标题内容" }
```

- 格式 A：楷体_GB2312 三号，首行缩进 2 字符
- 格式 B：黑体加粗三号，左对齐不缩进
- 格式 C：一般不用

### `h3` - 三级标题

```json
{ "type": "h3", "text": "1．标题内容。" }
```

- 格式 A：仿宋_GB2312 三号，首行缩进 2 字符
- 格式 B：黑体加粗三号
- 注意格式 A 要求三级标题末尾加句号"。"

### `p` - 正文段落

```json
{
  "type": "p",
  "text": "正文内容，可嵌入 {b:加粗文本} 和 {u:下划线文本} 标记",
  "align": "justified"
}
```

- 所有正文段落默认首行缩进两字符（640 DXA），不可覆盖——这是格式 B 的设计要求
- 可选字段 `align`：`"justified"`（默认）/ `"left"` / `"center"` / `"right"`
- 文本内支持两种行内标记：
  - `{b:文本}` 渲染为**加粗**
  - `{u:文本}` 渲染为下划线
- 不支持颜色标记。格式 B 设计要求纯黑配色，重点用加粗/下划线突出
- `{` 和 `}` 如需字面量，用 `\u007B` 和 `\u007D` 的 Unicode 转义（实际场景几乎不需要）

### `emp` - 空行

```json
{ "type": "emp", "height": 200 }
```

- `height` 为行距 DXA 单位，默认 200，可选
- 用于段落间隔

### `photo` - 居中嵌入图片（仅格式 B）

```json
{ "type": "photo", "path": "examples/format-b-project-brief/sample-portrait.png", "width": 200, "height": 260 }
```

- `path`：PNG 图片相对路径或绝对路径
- `width` / `height`：可选，单位为像素，默认 200×260（接近证件照宽高比）
- 图片以独立段落居中显示
- 通常配合 `name` 和 `role` 块使用，构成"照片 + 姓名 + 职务 + 简历"的人物介绍版式

### `name` - 居中姓名（仅格式 B）

```json
{ "type": "name", "text": "Dr. Michael Weber" }
```

- 黑体三号加粗居中，通常出现在 `photo` 块下方
- 用于人物照片的姓名标注

### `role` - 居中职务（仅格式 B）

```json
{ "type": "role", "text": "慕尼黑工业大学副校长（国际合作）" }
```

- 仿宋三号居中，通常出现在 `name` 块下方
- 用于人物姓名下的职务说明

### `list` - 名单条目（仅格式 C）

由 `nameList.groups[].members` 提供数据，不在 `body` 中使用。渲染时：
- 两字姓名中间补全角空格使其与三字姓名对齐
- 姓名与职务之间用两个全角空格分隔

## 完整示例

见 `examples/` 目录下三个范例的 `content.json`。

## 为什么用 JSON

1. 避免 JavaScript 字符串中中文引号、破折号等字符引发的语法错误
2. 内容与渲染逻辑解耦，同一份内容可渲染为不同格式
3. 便于程序化生成或其他工具链接入
