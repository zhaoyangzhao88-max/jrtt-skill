/**
 * jrtt-skill · Word文章生成器 v4
 * 全面排版优化版
 */

const {
  Document, Packer, Paragraph, TextRun, AlignmentType,
  BorderStyle, ShadingType, WidthType,
  Table, TableRow, TableCell, Footer, PageNumber,
  ExternalHyperlink, UnderlineType,
} = require('docx');
const fs   = require('fs');
const path = require('path');
const os   = require('os');

// ── 颜色 ─────────────────────────────────────────────
const C = {
  red:       "C0392B",
  darkGray:  "2C3E50",
  midGray:   "566573",
  lightGray: "AAB7B8",
  blue:      "1A5276",
  bgBlue:    "EBF5FB",
  bgWarm:    "FEF9E7",
  bgImage:   "EAFAF1",
  border:    "CCD1D1",
  white:     "FFFFFF",
};

// ── 尺寸常量 ─────────────────────────────────────────
// 字号（half-points，24=12pt）
const SZ_TITLE  = 36;   // 标题 18pt（缩小，避免太大）
const SZ_H1     = 26;   // 章节标题 13pt
const SZ_BODY   = 23;   // 正文 11.5pt（略小更舒适）
const SZ_SMALL  = 20;   // 辅助信息 10pt
const SZ_META   = 18;   // 页脚 9pt

// 字体
const FONT    = "微软雅黑";
const FONT_EN = "Arial";

// 行距（240=单倍，360=1.5倍，480=2倍）
const LINE_BODY  = 400;  // 正文 约1.67倍
const LINE_BOX   = 400;  // 框内 约1.67倍

// 段落间距（twips，20twips=1pt）
const PARA_AFTER  = 200;  // 段后 10pt
const PARA_BEFORE = 0;

// 页边距（DXA，1440=1英寸）
const MARGIN = { top: 1800, right: 1800, bottom: 1800, left: 1800 };

// ── 辅助：空行 ────────────────────────────────────────
const spacer = (pt = 8) => new Paragraph({
  children: [new TextRun("")],
  spacing: { before: 0, after: pt * 20 },
});

// ── 辅助：分隔线 ──────────────────────────────────────
const divider = (color = C.border, size = 4) => new Paragraph({
  children: [new TextRun("")],
  border: { bottom: { style: BorderStyle.SINGLE, size, color, space: 2 } },
  spacing: { before: 200, after: 200 },
});

// ── 正文段落 ──────────────────────────────────────────
const bodyPara = (text, bold = false) => new Paragraph({
  children: [new TextRun({
    text, bold,
    size: SZ_BODY,
    color: C.darkGray,
    font: { name: FONT },
  })],
  spacing: { before: PARA_BEFORE, after: PARA_AFTER, line: LINE_BODY, lineRule: "auto" },
});

// ── 章节标题（左侧红色色块+文字，比border更醒目） ─────
const sectionTitle = (emoji, title) => {
  // 用Table实现左侧色块效果，比border更稳定
  return new Table({
    width: { size: 8506, type: WidthType.DXA },
    columnWidths: [120, 8386],
    rows: [new TableRow({
      children: [
        // 左侧红色色块
        new TableCell({
          shading: { fill: C.red, type: ShadingType.CLEAR },
          borders: {
            top:    { style: BorderStyle.NONE },
            bottom: { style: BorderStyle.NONE },
            left:   { style: BorderStyle.NONE },
            right:  { style: BorderStyle.NONE },
          },
          width: { size: 120, type: WidthType.DXA },
          children: [new Paragraph({ children: [] })],
        }),
        // 右侧标题文字
        new TableCell({
          shading: { fill: "F8F9FA", type: ShadingType.CLEAR },
          borders: {
            top:    { style: BorderStyle.NONE },
            bottom: { style: BorderStyle.SINGLE, size: 1, color: C.border },
            left:   { style: BorderStyle.NONE },
            right:  { style: BorderStyle.NONE },
          },
          margins: { top: 80, bottom: 80, left: 200, right: 120 },
          width: { size: 8386, type: WidthType.DXA },
          children: [new Paragraph({
            children: [
              new TextRun({ text: `${emoji}  `, size: SZ_H1, font: { name: FONT_EN } }),
              new TextRun({ text: title, bold: true, size: SZ_H1, color: C.red, font: { name: FONT } }),
            ],
          })],
        }),
      ],
    })],
    margins: { top: 480, bottom: 160 },
  });
};

// ── 居中金句 ──────────────────────────────────────────
const sloganPara = (text) => new Paragraph({
  children: [new TextRun({
    text, bold: true,
    size: SZ_H1,
    color: C.red,
    font: { name: FONT },
  })],
  alignment: AlignmentType.CENTER,
  spacing: { before: 320, after: 320 },
  border: {
    top:    { style: BorderStyle.SINGLE, size: 2, color: C.border, space: 8 },
    bottom: { style: BorderStyle.SINGLE, size: 2, color: C.border, space: 8 },
  },
});

// ── 引用框（用Table，彻底解决变形问题） ───────────────
const quoteBox = (text) => new Table({
  width: { size: 8506, type: WidthType.DXA },
  columnWidths: [8506],
  rows: [new TableRow({
    children: [new TableCell({
      shading: { fill: C.bgBlue, type: ShadingType.CLEAR },
      borders: {
        top:    { style: BorderStyle.SINGLE, size: 1, color: C.border },
        bottom: { style: BorderStyle.SINGLE, size: 1, color: C.border },
        left:   { style: BorderStyle.THICK,  size: 10, color: C.red },
        right:  { style: BorderStyle.SINGLE, size: 1, color: C.border },
      },
      margins: { top: 160, bottom: 160, left: 240, right: 240 },
      width: { size: 8506, type: WidthType.DXA },
      children: [new Paragraph({
        children: [new TextRun({
          text,
          italics: true,
          size: SZ_BODY,
          color: C.midGray,
          font: { name: FONT },
        })],
        spacing: { line: LINE_BOX },
      })],
    })],
  })],
});

// ── 争议框（用Table，彻底解决变形问题） ───────────────
const debateBox = (text) => new Table({
  width: { size: 8506, type: WidthType.DXA },
  columnWidths: [8506],
  rows: [new TableRow({
    children: [new TableCell({
      shading: { fill: C.bgWarm, type: ShadingType.CLEAR },
      borders: {
        top:    { style: BorderStyle.SINGLE, size: 1, color: C.border },
        bottom: { style: BorderStyle.SINGLE, size: 1, color: C.border },
        left:   { style: BorderStyle.THICK,  size: 10, color: C.red },
        right:  { style: BorderStyle.SINGLE, size: 1, color: C.border },
      },
      margins: { top: 200, bottom: 200, left: 240, right: 240 },
      width: { size: 8506, type: WidthType.DXA },
      children: [
        new Paragraph({
          children: [new TextRun({
            text: "💬  互动话题",
            bold: true,
            size: SZ_BODY,
            color: C.red,
            font: { name: FONT },
          })],
          spacing: { after: 120 },
        }),
        new Paragraph({
          children: [new TextRun({
            text,
            size: SZ_BODY,
            color: C.darkGray,
            font: { name: FONT },
          })],
          spacing: { line: LINE_BOX },
        }),
      ],
    })],
  })],
});

// ── 配图链接行 ────────────────────────────────────────
const imageLink = (text) => {
  const arrowIdx = text.indexOf('→');
  if (arrowIdx === -1) return bodyPara(text);
  const label = text.slice(0, arrowIdx).trim();
  const url   = text.slice(arrowIdx + 1).trim();
  return new Paragraph({
    children: [
      new TextRun({ text: `📷  ${label}    `, bold: true, size: SZ_SMALL, color: C.darkGray, font: { name: FONT } }),
      new ExternalHyperlink({
        link: url,
        children: [new TextRun({
          text: url,
          size: SZ_SMALL,
          color: C.blue,
          underline: { type: UnderlineType.SINGLE, color: C.blue },
          font: { name: FONT_EN },
          style: "Hyperlink",
        })],
      }),
    ],
    shading: { fill: C.bgImage, type: ShadingType.CLEAR },
    spacing: { before: 80, after: 80 },
    indent: { left: 200, right: 200 },
  });
};

// ── 配图区块 ──────────────────────────────────────────
function buildImageBlock(images) {
  if (!images || images.length === 0) return [];
  return [
    spacer(16),
    divider(C.border, 4),
    spacer(8),
    sectionTitle("📸", "相关配图搜索（Unsplash 免费图库）"),
    spacer(4),
    ...images.map(imageLink),
    spacer(8),
    new Paragraph({
      children: [new TextRun({
        text: "💡 点击链接在 Unsplash 搜索免费高清配图，下载后直接上传今日头条",
        size: SZ_META,
        color: C.lightGray,
        italics: true,
        font: { name: FONT },
      })],
    }),
  ];
}

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

// ── 主函数 ────────────────────────────────────────────
function buildDoc(inputPath, outputPath) {
  const data = JSON.parse(fs.readFileSync(inputPath, "utf8"));

  // ── 标题（按标点拆行，18pt，不过大） ─────────────────
  const titleLines = data.title
    .split(/(?<=[？?！!，,。])/u)
    .map(s => s.trim())
    .filter(Boolean);

  const headerBlock = [
    spacer(8),
    ...titleLines.map((line, i) => new Paragraph({
      children: [new TextRun({
        text: line,
        bold: true,
        size: SZ_TITLE,
        color: C.red,
        font: { name: FONT },
      })],
      spacing: { before: 0, after: i < titleLines.length - 1 ? 60 : 200 },
    })),
    // 红色粗分隔线
    new Paragraph({
      children: [new TextRun("")],
      border: { bottom: { style: BorderStyle.SINGLE, size: 10, color: C.red, space: 2 } },
      spacing: { before: 0, after: 400 },
    }),
  ];

  const bodySections  = data.sections.filter(s => s.type !== "image");
  const imageSections = data.sections.filter(s => s.type === "image");
  const images        = imageSections.map(s => s.text).concat(data.images || []);

  const bodyBlocks = bodySections.map(renderSection);
  const imageBlock = buildImageBlock(images);

  const doc = new Document({
    styles: {
      default: {
        document: {
          run: { font: { name: FONT }, size: SZ_BODY, color: C.darkGray },
          paragraph: { spacing: { line: LINE_BODY, lineRule: "auto", after: PARA_AFTER } },
        },
      },
    },
    sections: [{
      properties: {
        page: {
          size:   { width: 11906, height: 16838 },
          margin: MARGIN,
        },
      },
      footers: {
        default: new Footer({
          children: [new Paragraph({
            children: [
              new TextRun({ text: "今日头条 · 时事评论    ", size: SZ_META, color: C.lightGray, font: { name: FONT } }),
              new TextRun({ children: [PageNumber.CURRENT], size: SZ_META, color: C.lightGray, font: { name: FONT_EN } }),
            ],
            alignment: AlignmentType.CENTER,
          })],
        }),
      },
      children: [...headerBlock, ...bodyBlocks, ...imageBlock],
    }],
  });

  // ── 确保输出目录存在 ──────────────────────────────
  const outDir = path.dirname(outputPath);
  try {
    if (!fs.existsSync(outDir)) {
      fs.mkdirSync(outDir, { recursive: true });
      console.log(`📁 已创建目录：${outDir}`);
    }
  } catch (e) {
    console.warn(`⚠️  无法创建目录 ${outDir}，将保存到当前目录`);
    outputPath = path.basename(outputPath);
  }

  Packer.toBuffer(doc).then(buf => {
    fs.writeFileSync(outputPath, buf);
    console.log(`✅ Word文档已生成：${outputPath}`);
  }).catch(err => {
    console.error("❌ 生成失败：", err.message);
    process.exit(1);
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
  const raw  = JSON.parse(fs.readFileSync(inputPath, "utf8"));
  const safe = (raw.title || "头条文章")
    .replace(/[\\/:*?"<>|]/g, "").trim().slice(0, 50);
  resolvedOutput = path.join(DEFAULT_OUTPUT_DIR, `${safe}.docx`).replace(/\\/g, '/');
}

buildDoc(inputPath, resolvedOutput);
