/**
 * jrtt-skill · Word文章生成器
 * 用法：node gen_docx.js <json输入文件> [输出路径]
 *
 * JSON输入格式：
 * {
 *   "title": "文章标题",
 *   "sections": [
 *     { "type": "body",    "text": "正文段落" },
 *     { "type": "bold",    "text": "加粗段落" },
 *     { "type": "quote",   "text": "引用框内容" },
 *     { "type": "heading", "emoji": "📊", "text": "章节标题" },
 *     { "type": "divider" },
 *     { "type": "spacer",  "pt": 10 },
 *     { "type": "slogan",  "text": "居中金句" },
 *     { "type": "debate",  "text": "争议互动框内容" },
 *     { "type": "image",   "text": "中文关键词 → https://unsplash.com/..." }
 *   ]
 * }
 */

const {
  Document, Packer, Paragraph, TextRun, AlignmentType,
  BorderStyle, ShadingType, WidthType,
  Table, TableRow, TableCell, Footer, PageNumber,
  ExternalHyperlink,
} = require('docx');
const fs = require('fs');

// ── 颜色 ─────────────────────────────────────────────
const C = {
  red:       "C0392B",
  darkGray:  "2C3E50",
  lightGray: "7F8C8D",
  blue:      "2980B9",
  bgBlue:    "E8F4FD",
  bgWarm:    "FDF6F0",
  bgImage:   "F0F7F0",
  border:    "E0E0E0",
  white:     "FFFFFF",
};

// ── 辅助函数 ──────────────────────────────────────────
const spacer = (pt = 6) => new Paragraph({
  children: [],
  spacing: { before: 0, after: pt * 20 },
});

const divider = () => new Paragraph({
  children: [new TextRun("")],
  border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: C.border, space: 1 } },
  spacing: { before: 100, after: 100 },
});

const bodyPara = (text, bold = false) => new Paragraph({
  children: [new TextRun({ text, font: "Arial", size: 24, color: C.darkGray, bold })],
  spacing: { before: 0, after: 200, line: 420 },
});

const sectionTitle = (emoji, title) => new Paragraph({
  children: [new TextRun({ text: `${emoji}  ${title}`, font: "Arial", size: 26, bold: true, color: C.red })],
  spacing: { before: 400, after: 160 },
  border: { bottom: { style: BorderStyle.SINGLE, size: 2, color: C.red, space: 4 } },
});

const sloganPara = (text) => new Paragraph({
  children: [new TextRun({ text, font: "Arial", size: 26, bold: true, color: C.red })],
  alignment: AlignmentType.CENTER,
  spacing: { before: 0, after: 200 },
});

// ── 引用框（蓝底，用Table实现左粗边） ─────────────────
const quoteBox = (text) => new Table({
  width: { size: 9026, type: WidthType.DXA },
  columnWidths: [9026],
  rows: [new TableRow({
    children: [new TableCell({
      shading: { fill: C.bgBlue, type: ShadingType.CLEAR },
      borders: {
        top:    { style: BorderStyle.SINGLE, size: 1, color: C.border },
        bottom: { style: BorderStyle.SINGLE, size: 1, color: C.border },
        left:   { style: BorderStyle.THICK,  size: 8, color: C.red },
        right:  { style: BorderStyle.SINGLE, size: 1, color: C.border },
      },
      margins: { top: 160, bottom: 160, left: 220, right: 200 },
      width: { size: 9026, type: WidthType.DXA },
      children: [new Paragraph({
        children: [new TextRun({ text, font: "Arial", size: 22, color: C.darkGray, italics: true })],
        spacing: { line: 360 },
      })],
    })],
  })],
});

// ── 争议框：改用普通段落，避免表格在某些编辑器里变形 ──
const debateBox = (text) => new Paragraph({
  children: [
    new TextRun({ text: "💬 互动话题    ", font: "Arial", size: 22, bold: true, color: C.red }),
    new TextRun({ text, font: "Arial", size: 22, color: C.darkGray }),
  ],
  border: {
    top:    { style: BorderStyle.SINGLE, size: 1, color: C.border, space: 4 },
    bottom: { style: BorderStyle.SINGLE, size: 1, color: C.border, space: 4 },
    left:   { style: BorderStyle.THICK,  size: 8, color: C.red,    space: 4 },
    right:  { style: BorderStyle.SINGLE, size: 1, color: C.border, space: 4 },
  },
  shading: { fill: C.bgWarm, type: ShadingType.CLEAR },
  spacing: { before: 200, after: 200, line: 420 },
  indent: { left: 160, right: 160 },
});

// ── 配图链接行（绿底，可点击超链接） ──────────────────
const imageLink = (text) => {
  // 格式：「中文关键词 → https://unsplash.com/...」
  const arrowIdx = text.indexOf('→');
  if (arrowIdx === -1) return bodyPara(text);
  const label = text.slice(0, arrowIdx).trim();
  const url   = text.slice(arrowIdx + 1).trim();
  return new Paragraph({
    children: [
      new TextRun({ text: `📷  ${label}    `, font: "Arial", size: 22, bold: true, color: C.darkGray }),
      new ExternalHyperlink({
        link: url,
        children: [new TextRun({
          text: url,
          font: "Arial", size: 22, color: C.blue,
          underline: { type: "single", color: C.blue },
          style: "Hyperlink",
        })],
      }),
    ],
    spacing: { before: 100, after: 100, line: 360 },
    shading: { fill: C.bgImage, type: ShadingType.CLEAR },
    indent: { left: 160, right: 160 },
  });
};

// ── 渲染分发 ──────────────────────────────────────────
function renderSection(s) {
  switch (s.type) {
    case "body":    return bodyPara(s.text, false);
    case "bold":    return bodyPara(s.text, true);
    case "quote":   return quoteBox(s.text);
    case "heading": return sectionTitle(s.emoji || "▶", s.text);
    case "divider": return divider();
    case "spacer":  return spacer(s.pt || 8);
    case "slogan":  return sloganPara(s.text);
    case "debate":  return debateBox(s.text);
    case "image":   return imageLink(s.text);
    default:        return bodyPara(s.text || "");
  }
}

// ── 配图区块（标题 + 4条链接） ────────────────────────
function buildImageBlock(images) {
  if (!images || images.length === 0) return [];
  return [
    spacer(16),
    divider(),
    spacer(8),
    sectionTitle("📸", "相关配图搜索（Unsplash 免费图库）"),
    spacer(4),
    ...images.map(item => imageLink(item)),
    spacer(8),
    new Paragraph({
      children: [new TextRun({
        text: "💡 点击链接在 Unsplash 搜索免费高清配图，下载后直接上传今日头条",
        font: "Arial", size: 20, color: C.lightGray, italics: true,
      })],
      spacing: { before: 0, after: 0 },
    }),
  ];
}

// ── 主函数 ────────────────────────────────────────────
function buildDoc(inputPath, outputPath) {
  const data = JSON.parse(fs.readFileSync(inputPath, "utf8"));

  // 标题只保留文字，不加栏目标签和元信息行
  const headerBlock = [
    ...data.title.split(/[？?，,。！!]/u)
      .filter(t => t.trim())
      .map((line, i) => new Paragraph({
        children: [new TextRun({ text: line.trim(), font: "Arial", size: 40, bold: true, color: C.red })],
        spacing: { before: i === 0 ? 0 : 0, after: 120 },
      })),
    // 红色分隔线
    new Paragraph({
      children: [new TextRun("")],
      border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: C.red, space: 1 } },
      spacing: { before: 120, after: 240 },
    }),
  ];

  // 正文sections（过滤掉image类型，单独处理）
  const bodySections = data.sections.filter(s => s.type !== "image");
  const imageSections = data.sections.filter(s => s.type === "image");

  // 如果JSON里没有image条目，从images字段读取
  const images = imageSections.map(s => s.text).concat(data.images || []);

  const bodyBlocks  = bodySections.map(renderSection);
  const imageBlock  = buildImageBlock(images);

  const doc = new Document({
    styles: {
      default: { document: { run: { font: "Arial", size: 24, color: C.darkGray } } },
    },
    sections: [{
      properties: {
        page: {
          size: { width: 11906, height: 16838 },
          margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 },
        },
      },
      footers: {
        default: new Footer({
          children: [new Paragraph({
            children: [
              new TextRun({ text: "今日头条 · 时事评论    ", font: "Arial", size: 18, color: C.lightGray }),
              new TextRun({ children: [PageNumber.CURRENT], font: "Arial", size: 18, color: C.lightGray }),
            ],
            alignment: AlignmentType.CENTER,
          })],
        }),
      },
      children: [...headerBlock, ...bodyBlocks, ...imageBlock],
    }],
  });

  Packer.toBuffer(doc).then(buf => {
    fs.writeFileSync(outputPath, buf);
    console.log(`✅ Word文档已生成：${outputPath}`);
  });
}

// ── 入口 ──────────────────────────────────────────────
const DEFAULT_OUTPUT_DIR = "E:/VSCODE/文章";

const [,, inputPath, outputArg] = process.argv;
if (!inputPath) {
  console.error("用法：node gen_docx.js <input.json> [output.docx]");
  process.exit(1);
}

let resolvedOutput = outputArg;
if (!resolvedOutput) {
  const data = JSON.parse(fs.readFileSync(inputPath, "utf8"));
  const safeName = (data.title || "头条文章")
    .replace(/[\\/:*?"<>|]/g, "")
    .slice(0, 50);
  resolvedOutput = `${DEFAULT_OUTPUT_DIR}/${safeName}.docx`;
}

buildDoc(inputPath, resolvedOutput);
