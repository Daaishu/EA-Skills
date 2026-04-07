#!/usr/bin/env node
/**
 * build_docx.js — Unified Chinese document builder for three formats.
 *
 * Usage:
 *   node build_docx.js <content.json> <output.docx> <format>
 *
 * Arguments:
 *   content.json  Input content file (see references/content-schema.md)
 *   output.docx   Output Word file path
 *   format        One of: a (policy report), b (internal brief), c (agenda)
 *
 * Requirements:
 *   Node.js 18+, npm package `docx` (>= 9.x)
 *
 * Design notes:
 *   - All text content lives in JSON; no literals in code paths.
 *   - Three formats share paragraph/run helpers but have format-specific layouts.
 *   - No colors in Format B — emphasis is {b:bold} and {u:underline} only.
 *   - No table-based layouts (past docx-js table-border bugs break MS Word).
 *   - No embedded logos (generic template; users add their own post-build).
 */

const fs = require("fs");
const path = require("path");
const {
  Document, Packer, Paragraph, TextRun, AlignmentType, PageNumber,
  Header, Footer, BorderStyle, LineRuleType, UnderlineType, ImageRun
} = require("docx");

// ─── Constants ──────────────────────────────────────────────────────────

const FONTS = {
  // Format A (government document) — uses GB2312 variants
  fangsongGB: { name: "仿宋_GB2312", eastAsia: "仿宋_GB2312" },
  kaitiGB: { name: "楷体_GB2312", eastAsia: "楷体_GB2312" },
  xiaoBiaoSong: { name: "方正小标宋简体", eastAsia: "方正小标宋简体" },
  // General-purpose Chinese fonts
  fangsong: { name: "仿宋", eastAsia: "仿宋" },
  heiti: { name: "黑体", eastAsia: "黑体" },
  songti: { name: "宋体", eastAsia: "宋体" },
  kaiti: { name: "楷体", eastAsia: "楷体" },
};

const PAGE_A4 = { width: 11906, height: 16838 };

// ─── Inline markup parser ───────────────────────────────────────────────
// Supports {b:text} for bold and {u:text} for underline.

function parseInlineMarkup(text, baseFont, baseSize, kaitiFont = null) {
  const runs = [];
  // {b:...} bold, {u:...} underline, {k:...} emphasis-via-楷体 (Format A only)
  const regex = /\{(b|u|k):([^}]+)\}/g;
  let lastIndex = 0;
  let match;
  while ((match = regex.exec(text)) !== null) {
    if (match.index > lastIndex) {
      runs.push(new TextRun({
        text: text.slice(lastIndex, match.index),
        font: baseFont,
        size: baseSize,
      }));
    }
    const props = { text: match[2], font: baseFont, size: baseSize };
    if (match[1] === "b") props.bold = true;
    if (match[1] === "u") props.underline = { type: UnderlineType.SINGLE };
    if (match[1] === "k") {
      // Government-document emphasis: switch to 楷体, no bold, no color.
      // Falls back to baseFont if no kaitiFont was provided (i.e. format B passes none).
      if (kaitiFont) props.font = kaitiFont;
    }
    runs.push(new TextRun(props));
    lastIndex = regex.lastIndex;
  }
  if (lastIndex < text.length) {
    runs.push(new TextRun({
      text: text.slice(lastIndex),
      font: baseFont,
      size: baseSize,
    }));
  }
  return runs;
}

// ─── Shared helpers ─────────────────────────────────────────────────────

function emptyPara(height = 200) {
  return new Paragraph({
    spacing: { line: height, lineRule: LineRuleType.EXACT },
    children: [],
  });
}

function centerPara(text, opts) {
  return new Paragraph({
    spacing: {
      before: opts.before || 0,
      after: opts.after || 0,
      line: opts.line || 560,
      lineRule: LineRuleType.EXACT,
    },
    alignment: AlignmentType.CENTER,
    children: [new TextRun({
      text,
      font: opts.font,
      size: opts.size,
      bold: opts.bold || false,
    })],
  });
}

function alignmentFromString(s) {
  switch (s) {
    case "left": return AlignmentType.LEFT;
    case "right": return AlignmentType.RIGHT;
    case "center": return AlignmentType.CENTER;
    case "justified":
    default: return AlignmentType.JUSTIFIED;
  }
}

function footerWithPageNumber() {
  return new Footer({
    children: [new Paragraph({
      alignment: AlignmentType.CENTER,
      children: [
        new TextRun({ text: "— ", font: FONTS.songti, size: 28 }),
        new TextRun({ children: [PageNumber.CURRENT], font: FONTS.songti, size: 28 }),
        new TextRun({ text: " —", font: FONTS.songti, size: 28 }),
      ],
    })],
  });
}

// ═══════════════════════════════════════════════════════════════════════
// FORMAT A — Policy Report
// ═══════════════════════════════════════════════════════════════════════

function buildFormatA(content) {
  const LINE_GW = 560; // 28pt exact line spacing
  const INDENT_2CH = 640;

  const gwTitle = (text, opts = {}) => new Paragraph({
    spacing: {
      before: opts.before !== undefined ? opts.before : 312,
      after: opts.after !== undefined ? opts.after : 312,
      line: LINE_GW, lineRule: LineRuleType.EXACT,
    },
    alignment: AlignmentType.CENTER,
    children: [new TextRun({
      text, font: FONTS.xiaoBiaoSong, size: 44, color: "000000",
    })],
  });

  const h1 = (text) => new Paragraph({
    spacing: { line: LINE_GW, lineRule: LineRuleType.EXACT },
    indent: { firstLine: INDENT_2CH },
    children: [new TextRun({ text, font: FONTS.heiti, size: 32, color: "000000" })],
  });

  const h2 = (text) => new Paragraph({
    spacing: { line: LINE_GW, lineRule: LineRuleType.EXACT },
    indent: { firstLine: INDENT_2CH },
    children: [new TextRun({ text, font: FONTS.kaitiGB, size: 32, color: "000000" })],
  });

  const h3 = (text) => new Paragraph({
    spacing: { line: LINE_GW, lineRule: LineRuleType.EXACT },
    indent: { firstLine: INDENT_2CH },
    children: [new TextRun({ text, font: FONTS.fangsongGB, size: 32, color: "000000" })],
  });

  const body = (text, opts = {}) => new Paragraph({
    spacing: { line: LINE_GW, lineRule: LineRuleType.EXACT },
    indent: { firstLine: opts.noIndent ? 0 : INDENT_2CH },
    alignment: alignmentFromString(opts.align),
    children: parseInlineMarkup(text, FONTS.fangsongGB, 32, FONTS.kaitiGB),
  });

  const rightPara = (text) => new Paragraph({
    spacing: { line: LINE_GW, lineRule: LineRuleType.EXACT },
    alignment: AlignmentType.RIGHT,
    children: [new TextRun({ text, font: FONTS.fangsongGB, size: 32, color: "000000" })],
  });

  const children = [];

  // Title — supports multi-line via titleLines, falls back to single line from title
  const titleLines = content.titleLines && content.titleLines.length > 0
    ? content.titleLines
    : [content.title];
  titleLines.forEach((line, idx) => {
    const isFirst = idx === 0;
    const isLast = idx === titleLines.length - 1;
    children.push(gwTitle(line, {
      before: isFirst ? 312 : 0,
      after: isLast ? 312 : 0,
    }));
  });

  // Body blocks
  for (const block of content.body || []) {
    switch (block.type) {
      case "h1": children.push(h1(block.text)); break;
      case "h2": children.push(h2(block.text)); break;
      case "h3": children.push(h3(block.text)); break;
      case "p": children.push(body(block.text, block)); break;
      case "emp": children.push(emptyPara(block.height)); break;
    }
  }

  // Signature block
  if (content.signature) {
    children.push(emptyPara());
    if (content.signature.org) children.push(rightPara(content.signature.org));
    if (content.signature.date) children.push(rightPara(content.signature.date));
  }

  return new Document({
    styles: { default: { document: { run: { font: "仿宋_GB2312", size: 32 } } } },
    sections: [{
      properties: {
        page: {
          size: PAGE_A4,
          margin: { top: 2098, bottom: 1985, left: 1588, right: 1474, footer: 1134 },
        },
      },
      footers: { default: footerWithPageNumber() },
      children,
    }],
  });
}

// ═══════════════════════════════════════════════════════════════════════
// FORMAT B — Internal Brief
// ═══════════════════════════════════════════════════════════════════════

function buildFormatB(content) {
  const LINE = 560;        // 28pt fixed line spacing — used by body AND headings AND title
  const ROW = 560;         // "1 行" of paragraph spacing == one line height
  const INDENT_2CH = 640;

  // Title — 黑体 二号
  // Single-line:           before=1行, after=1行
  // Multi-line, first line: before=1行, after=0
  // Multi-line, middle:    before=0, after=0
  // Multi-line, last line: before=0, after=1行
  // Line spacing: 28pt fixed (same as body) so the title block visually integrates
  const docTitle = (text, opts = {}) => new Paragraph({
    spacing: {
      before: opts.before ?? ROW,
      after: opts.after ?? ROW,
      line: LINE,
      lineRule: LineRuleType.EXACT,
    },
    alignment: AlignmentType.CENTER,
    children: [new TextRun({ text, font: FONTS.heiti, size: 44, bold: true })],
  });

  // Section headings — 黑体 三号 (h1/h2) or 小三 (h3)
  // before=0, after=0, line=28pt fixed — same line spacing as body, no extra space
  const h1 = (text) => new Paragraph({
    spacing: { before: 0, after: 0, line: LINE, lineRule: LineRuleType.EXACT },
    children: [new TextRun({ text, font: FONTS.heiti, size: 32, bold: true })],
  });

  const h2 = (text) => new Paragraph({
    spacing: { before: 0, after: 0, line: LINE, lineRule: LineRuleType.EXACT },
    children: [new TextRun({ text, font: FONTS.heiti, size: 32, bold: true })],
  });

  const h3 = (text) => new Paragraph({
    spacing: { before: 0, after: 0, line: LINE, lineRule: LineRuleType.EXACT },
    children: [new TextRun({ text, font: FONTS.heiti, size: 28, bold: true })],
  });

  // Body — 仿宋 三号, first-line indent two characters, fixed 28pt line spacing
  const body = (text, opts = {}) => new Paragraph({
    spacing: { before: 0, after: 0, line: LINE, lineRule: LineRuleType.EXACT },
    indent: { firstLine: INDENT_2CH },
    alignment: alignmentFromString(opts.align),
    children: parseInlineMarkup(text, FONTS.fangsong, 32),
  });

  // Photo block — centered image
  // Used for portraits in leader-bio sections.
  // The path can be absolute or relative to the cwd of the build script invocation.
  // IMPORTANT: do NOT use fixed line spacing here — the paragraph must auto-size
  // to the image height, otherwise docx-js will squash the image to one line of text.
  const photoBlock = (path, opts = {}) => {
    const data = fs.readFileSync(path);
    const width = opts.width ?? 150;   // ~4cm at 96dpi (one-inch ID-photo width)
    const height = opts.height ?? 200; // 4cm × 5.3cm — standard 一寸 ID-photo aspect
    return new Paragraph({
      spacing: { before: 240, after: 120 },  // small breathing room around the photo
      alignment: AlignmentType.CENTER,
      children: [new ImageRun({
        data,
        transformation: { width, height },
        type: "png",
      })],
    });
  };

  // Centered name (under a photo) — 黑体 三号 加粗
  const photoName = (text) => new Paragraph({
    spacing: { before: 0, after: 0, line: LINE, lineRule: LineRuleType.EXACT },
    alignment: AlignmentType.CENTER,
    children: [new TextRun({ text, font: FONTS.heiti, size: 32, bold: true })],
  });

  // Centered title/role line (under the name) — 仿宋 三号
  const photoRole = (text) => new Paragraph({
    spacing: { before: 0, after: 0, line: LINE, lineRule: LineRuleType.EXACT },
    alignment: AlignmentType.CENTER,
    children: [new TextRun({ text, font: FONTS.fangsong, size: 32 })],
  });

  const children = [];

  // Title block — multi-line aware spacing per the梯形/菱形 rule
  const titleLines = content.titleLines && content.titleLines.length > 0
    ? content.titleLines
    : [content.title];

  if (titleLines.length === 1) {
    // Single-line: standard before=1行, after=1行
    children.push(docTitle(titleLines[0]));
  } else {
    // Multi-line: first line before=1行, after=0; middle 0/0; last line 0/after=1行
    titleLines.forEach((line, idx) => {
      const isFirst = idx === 0;
      const isLast = idx === titleLines.length - 1;
      children.push(docTitle(line, {
        before: isFirst ? ROW : 0,
        after: isLast ? ROW : 0,
      }));
    });
  }

  // Optional author line (centered)
  if (content.author) {
    children.push(new Paragraph({
      spacing: { before: 0, after: 0, line: LINE, lineRule: LineRuleType.EXACT },
      alignment: AlignmentType.CENTER,
      children: [new TextRun({ text: content.author, font: FONTS.fangsong, size: 32 })],
    }));
  }

  // Body blocks
  for (const block of content.body || []) {
    switch (block.type) {
      case "h1": children.push(h1(block.text)); break;
      case "h2": children.push(h2(block.text)); break;
      case "h3": children.push(h3(block.text)); break;
      case "p":  children.push(body(block.text, block)); break;
      case "emp": children.push(emptyPara(block.height)); break;
      case "photo": children.push(photoBlock(block.path, block)); break;
      case "name":  children.push(photoName(block.text)); break;
      case "role":  children.push(photoRole(block.text)); break;
    }
  }

  return new Document({
    styles: { default: { document: { run: { font: "仿宋", size: 32 } } } },
    sections: [{
      properties: {
        page: {
          size: PAGE_A4,
          margin: { top: 1418, bottom: 1418, left: 1596, right: 1482 },
        },
      },
      headers: {
        default: new Header({
          children: [new Paragraph({
            alignment: AlignmentType.CENTER,
            border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: "CCCCCC", space: 1 } },
            children: [new TextRun({
              text: "内部资料  注意保密",
              font: FONTS.fangsong, size: 18, color: "999999",
            })],
          })],
        }),
      },
      footers: {
        default: new Footer({
          children: [new Paragraph({
            alignment: AlignmentType.CENTER,
            children: [
              new TextRun({ text: "— ", font: FONTS.songti, size: 20 }),
              new TextRun({ children: [PageNumber.CURRENT], font: FONTS.songti, size: 20 }),
              new TextRun({ text: " —", font: FONTS.songti, size: 20 }),
            ],
          })],
        }),
      },
      children,
    }],
  });
}

// ═══════════════════════════════════════════════════════════════════════
// FORMAT C — Agenda
// ═══════════════════════════════════════════════════════════════════════

function buildFormatC(content) {
  const LINE_AGENDA = 800;  // 40pt — generous for cover page items
  const LINE_LIST = 560;    // 28pt — standard body line spacing for name list

  const agendaTitle = (text) => new Paragraph({
    spacing: { line: LINE_AGENDA, lineRule: LineRuleType.EXACT },
    alignment: AlignmentType.CENTER,
    children: [new TextRun({ text, font: FONTS.songti, size: 44, bold: true })],
  });

  const agendaInfo = (text) => new Paragraph({
    spacing: { line: LINE_AGENDA, lineRule: LineRuleType.EXACT },
    indent: {
      // Hanging indent so a long location wraps under the colon position.
      // "时　　间：" / "地　　点：" = 5 full-width chars at 16pt ≈ 1600 DXA.
      left: 1600,
      hanging: 1600,
    },
    children: [new TextRun({ text, font: FONTS.fangsong, size: 32, bold: true })],
  });

  // Unified agenda item — three-size font (sz 32 = 三号), consistent spacing,
  // no emp paragraphs between items. Visual separation comes from the line height.
  // Hanging indent so a long activity name wraps under the time slot.
  const agendaItem = (text) => new Paragraph({
    spacing: { line: LINE_AGENDA, lineRule: LineRuleType.EXACT },
    indent: {
      // "HH:MM-HH:MM" + 2 full-width spaces = 11 half-width + 2 full-width
      // ≈ 11 × 160 + 2 × 320 = 2400 DXA at 16pt sz 32.
      left: 2400,
      hanging: 2400,
    },
    children: [new TextRun({ text, font: FONTS.fangsong, size: 32, bold: true })],
  });

  const listPageTitle = (text) => new Paragraph({
    spacing: { before: 0, after: 240, line: LINE_LIST, lineRule: LineRuleType.EXACT },
    alignment: AlignmentType.CENTER,
    children: [new TextRun({ text, font: FONTS.fangsong, size: 44, bold: true })],
  });

  const listGroupTitle = (text) => new Paragraph({
    spacing: { before: 240, after: 120, line: LINE_LIST, lineRule: LineRuleType.EXACT },
    children: [new TextRun({
      text, font: FONTS.fangsong, size: 32, bold: true, color: "000000",
    })],
  });

  // listMember accepts a `compact` flag — when set, uses tighter line spacing
  // so a long list can still fit on one page if the user prefers.
  // Hanging indent: when a long title wraps, the wrapped lines align under the
  // title column. Name (3 chars wide after padding) + 2 full-width spaces
  // = 5 full-width chars × 320 DXA = 1600 DXA at 16pt sz 32.
  const listMember = (name, title, compact = false) => {
    const paddedName = name.length === 2 ? name[0] + "\u3000" + name[1] : name;
    const text = title ? paddedName + "\u3000\u3000" + title : paddedName;
    return new Paragraph({
      spacing: {
        line: compact ? 480 : LINE_LIST,
        lineRule: LineRuleType.EXACT,
      },
      indent: {
        left: 1600,
        hanging: 1600,
      },
      children: [new TextRun({ text, font: FONTS.fangsong, size: 32, color: "000000" })],
    });
  };

  // ── Cover page ─────────────────────────────────────────────
  const coverChildren = [];
  const titleLines = content.titleLines && content.titleLines.length > 0
    ? content.titleLines
    : [content.title];
  titleLines.forEach(line => coverChildren.push(agendaTitle(line)));
  coverChildren.push(emptyPara(800));

  if (content.meta) {
    if (content.meta.time) {
      coverChildren.push(agendaInfo(`时\u3000\u3000间：${content.meta.time}`));
    }
    if (content.meta.location) {
      coverChildren.push(agendaInfo(`地\u3000\u3000点：${content.meta.location}`));
    }
    coverChildren.push(emptyPara(800));
    if (content.meta.agendaItem) {
      coverChildren.push(agendaItem(content.meta.agendaItem));
    }
  }

  // Body blocks render as agenda items with the SAME size and spacing as meta.agendaItem.
  // We deliberately ignore "emp" blocks here — separation should come from line height,
  // not extra empty paragraphs (which create the visual unevenness reported in v1).
  for (const block of content.body || []) {
    if (block.type === "p") {
      coverChildren.push(agendaItem(block.text));
    }
    // emp blocks intentionally skipped
  }

  // ── Name list pages ─────────────────────────────────────────
  // Auto-paginate: if the total member count exceeds the single-page threshold,
  // put each group on its own section (i.e. its own page).
  const sectionProps = {
    page: {
      size: PAGE_A4,
      margin: { top: 1800, bottom: 1418, left: 1596, right: 1482 },
    },
  };

  const sections = [{
    properties: sectionProps,
    footers: { default: footerWithPageNumber() },
    children: coverChildren,
  }];

  if (content.nameList && content.nameList.groups && content.nameList.groups.length > 0) {
    const groups = content.nameList.groups;
    const totalMembers = groups.reduce((sum, g) => sum + (g.members || []).length, 0);

    // Heuristic: with 28pt line + heading per group + 人员名单 title,
    // a single page comfortably holds about 18-20 entries.
    // If we're over that AND there are multiple groups, split each group onto its own page.
    const SINGLE_PAGE_THRESHOLD = 18;
    const splitByGroup = totalMembers > SINGLE_PAGE_THRESHOLD && groups.length >= 2;

    if (splitByGroup) {
      // One page per group. Each page has its own "人员名单" title.
      groups.forEach(group => {
        const pageChildren = [];
        pageChildren.push(listPageTitle("人员名单"));
        pageChildren.push(listGroupTitle(group.title));
        for (const member of group.members || []) {
          pageChildren.push(listMember(member.name, member.title));
        }
        sections.push({
          properties: sectionProps,
          footers: { default: footerWithPageNumber() },
          children: pageChildren,
        });
      });
    } else {
      // All groups on one page. If we're slightly over the threshold use compact mode.
      const compact = totalMembers > 14;
      const listChildren = [];
      listChildren.push(listPageTitle("人员名单"));

      groups.forEach((group) => {
        listChildren.push(listGroupTitle(group.title));
        for (const member of group.members || []) {
          listChildren.push(listMember(member.name, member.title, compact));
        }
      });

      sections.push({
        properties: sectionProps,
        footers: { default: footerWithPageNumber() },
        children: listChildren,
      });
    }
  }

  return new Document({
    styles: { default: { document: { run: { font: "仿宋", size: 32 } } } },
    sections,
  });
}

// ═══════════════════════════════════════════════════════════════════════
// Main entry
// ═══════════════════════════════════════════════════════════════════════

function main() {
  const args = process.argv.slice(2);
  if (args.length < 3) {
    console.error("Usage: node build_docx.js <content.json> <output.docx> <format>");
    console.error("  format: 'a' (policy report) | 'b' (internal brief) | 'c' (agenda)");
    process.exit(1);
  }
  const [inputPath, outputPath, formatArg] = args;
  const format = formatArg.toLowerCase();
  if (!["a", "b", "c"].includes(format)) {
    console.error(`Error: format must be 'a', 'b', or 'c', got: ${formatArg}`);
    process.exit(1);
  }
  if (!fs.existsSync(inputPath)) {
    console.error(`Error: input file not found: ${inputPath}`);
    process.exit(1);
  }

  let content;
  try {
    content = JSON.parse(fs.readFileSync(inputPath, "utf8"));
  } catch (e) {
    console.error(`Error: failed to parse JSON: ${e.message}`);
    process.exit(1);
  }

  if (!content.title && !content.titleLines) {
    console.error("Error: content must have 'title' or 'titleLines'");
    process.exit(1);
  }

  let doc;
  if (format === "a") doc = buildFormatA(content);
  else if (format === "b") doc = buildFormatB(content);
  else doc = buildFormatC(content);

  Packer.toBuffer(doc).then(buffer => {
    // Ensure output directory exists
    const outDir = path.dirname(outputPath);
    if (outDir && !fs.existsSync(outDir)) {
      fs.mkdirSync(outDir, { recursive: true });
    }
    fs.writeFileSync(outputPath, buffer);
    console.log(`✓ Generated ${outputPath} (format ${format.toUpperCase()}, ${buffer.length} bytes)`);
  }).catch(err => {
    console.error(`Error generating document: ${err.message}`);
    process.exit(1);
  });
}

main();
