const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
        Header, Footer, AlignmentType, LevelFormat,
        HeadingLevel, BorderStyle, WidthType, ShadingType,
        PageNumber, PageBreak, TabStopType, TabStopPosition } = require('docx');
const fs = require('fs');

// Color palette (PPT final version)
const C = {
  gov: "2563EB", mgmt: "059669", ent: "D97706", pub: "7C3AED",
  acc: "F59E0B", text: "1E293B", gray: "64748B", line: "E2E8F0",
};

const border = { style: BorderStyle.SINGLE, size: 1, color: C.line };
const borders = { top: border, bottom: border, left: border, right: border };

// Helper: section header with color
function sectionHeader(text, color) {
  return new Paragraph({
    spacing: { before: 360, after: 200 },
    border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: color, space: 4 } },
    children: [new TextRun({ text, bold: true, size: 32, font: "Microsoft YaHei", color })],
  });
}

// Helper: sub header
function subHeader(text, color) {
  return new Paragraph({
    spacing: { before: 240, after: 120 },
    children: [new TextRun({ text, bold: true, size: 26, font: "Microsoft YaHei", color })],
  });
}

// Helper: feature item with bullet
function featItem(title, desc, color) {
  return new Paragraph({
    spacing: { before: 80, after: 40 },
    indent: { left: 360 },
    children: [
      new TextRun({ text: "\u25b8 ", bold: true, size: 22, font: "Microsoft YaHei", color }),
      new TextRun({ text: title, bold: true, size: 22, font: "Microsoft YaHei", color }),
    ],
  });
}
function featDesc(desc) {
  return new Paragraph({
    spacing: { before: 0, after: 80 },
    indent: { left: 720 },
    children: [new TextRun({ text: desc, size: 20, font: "Microsoft YaHei", color: C.text })],
  });
}

// Helper: body text
function bodyText(text, opts) {
  return new Paragraph({
    spacing: { before: 60, after: 60 },
    children: [new TextRun({ text, size: 20, font: "Microsoft YaHei", color: C.text, ...(opts||{}) })],
  });
}

const doc = new Document({
  styles: {
    default: { document: { run: { font: "Microsoft YaHei", size: 20 } } },
    paragraphStyles: [
      { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 36, bold: true, font: "Microsoft YaHei", color: C.text },
        paragraph: { spacing: { before: 360, after: 200 }, outlineLevel: 0 } },
      { id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 28, bold: true, font: "Microsoft YaHei" },
        paragraph: { spacing: { before: 240, after: 160 }, outlineLevel: 1 } },
      { id: "Heading3", name: "Heading 3", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 24, bold: true, font: "Microsoft YaHei" },
        paragraph: { spacing: { before: 200, after: 120 }, outlineLevel: 2 } },
    ]
  },
  numbering: {
    config: [
      { reference: "bullets",
        levels: [{ level: 0, format: LevelFormat.BULLET, text: "\u2022", alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] },
    ]
  },
  sections: [
    // ===== COVER =====
    {
      properties: {
        page: {
          size: { width: 11906, height: 16838 },
          margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 },
        },
      },
      children: [
        new Paragraph({ spacing: { before: 3000 }, children: [] }),
        new Paragraph({
          alignment: AlignmentType.CENTER,
          children: [new TextRun({ text: "\u9752\u5ddd\u53bf\u5e84\u5b50\u4e0a\u5de5\u4e1a\u56ed\u533a", size: 28, font: "Microsoft YaHei", color: C.gray })],
        }),
        new Paragraph({
          alignment: AlignmentType.CENTER,
          spacing: { before: 120, after: 120 },
          children: [new TextRun({ text: "\u667a\u6167\u5e73\u53f0\u4ea7\u54c1\u624b\u518c", size: 48, bold: true, font: "Microsoft YaHei", color: C.text })],
        }),
        new Paragraph({
          alignment: AlignmentType.CENTER,
          border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: C.acc, space: 8 } },
          spacing: { after: 400 },
          children: [new TextRun({ text: "\u56db\u7c7b\u89d2\u8272 \u00b7 \u5168\u6a21\u5757\u667a\u6167\u5316 \u00b7 \u6570\u636e\u9a71\u52a8\u51b3\u7b56", size: 22, font: "Microsoft YaHei", color: C.gray })],
        }),
        new Paragraph({ spacing: { before: 600 }, children: [] }),
        // Role tags
        ...[
          [C.gov, "\u653f\u5e9c\u9886\u5bfc\u5c42"],
          [C.mgmt, "\u56ed\u533a\u7ba1\u7406\u5c42"],
          [C.ent, "\u5165\u9a7b\u4f01\u4e1a"],
          [C.pub, "\u516c\u4f17\u7528\u6237"],
        ].map(r => new Paragraph({
          alignment: AlignmentType.CENTER,
          spacing: { before: 80 },
          children: [
            new TextRun({ text: "\u25a0 ", color: r[0], bold: true, size: 24 }),
            new TextRun({ text: r[1], size: 22, font: "Microsoft YaHei", bold: true, color: r[0] }),
          ],
        })),
        new Paragraph({ spacing: { before: 600 }, children: [] }),
        new Paragraph({
          alignment: AlignmentType.CENTER,
          children: [new TextRun({ text: "\u56db\u5ddd\u7701\u5e7f\u5143\u5e02\u9752\u5ddd\u53bf \u00b7 \u7af9\u56ed\u9547\u5e84\u5b50\u4e0a\u5de5\u4e1a\u56ed\u533a", size: 20, font: "Microsoft YaHei", color: C.gray })],
        }),
        new Paragraph({
          alignment: AlignmentType.CENTER,
          children: [new TextRun({ text: "\u6280\u672f\u652f\u6301\uff1a\u9752\u8fc5\u79d1\u6280", size: 20, font: "Microsoft YaHei", color: C.gray })],
        }),
        new Paragraph({
          alignment: AlignmentType.CENTER,
          children: [new TextRun({ text: "2026\u5e745\u6708", size: 20, font: "Microsoft YaHei", color: C.gray })],
        }),
      ],
    },
    // ===== TOC =====
    {
      properties: {
        page: { size: { width: 11906, height: 16838 }, margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 } },
      },
      headers: { default: new Header({ children: [new Paragraph({
        alignment: AlignmentType.RIGHT, border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: C.gov, space: 4 } },
        children: [new TextRun({ text: "\u9752\u5ddd\u53bf\u5e84\u5b50\u4e0a\u667a\u6167\u5e73\u53f0\u4ea7\u54c1\u624b\u518c", size: 16, font: "Microsoft YaHei", color: C.gray })]
      })] }) },
      footers: { default: new Footer({ children: [new Paragraph({
        alignment: AlignmentType.CENTER,
        children: [new TextRun({ text: "\u7b2c ", size: 16, font: "Microsoft YaHei", color: C.gray }), new TextRun({ children: [PageNumber.CURRENT], size: 16, font: "Microsoft YaHei", color: C.gray }), new TextRun({ text: " \u9875", size: 16, font: "Microsoft YaHei", color: C.gray })]
      })] }) },
      children: [
        new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun({ text: "\u76ee\u5f55", bold: true })] }),
        ...[
          ["\u4e00\u3001\u4ea7\u54c1\u6982\u8ff0", ""],
          ["\u4e8c\u3001\u653f\u5e9c\u9886\u5bfc\u5c42\u89c6\u89d2", "\u6570\u636e\u9a71\u52a8\u51b3\u7b56"],
          ["\u4e09\u3001\u56ed\u533a\u7ba1\u7406\u5c42\u89c6\u89d2", "\u667a\u6167\u8fd0\u8425\u7ba1\u63a7"],
          ["\u56db\u3001\u5165\u9a7b\u4f01\u4e1a\u89c6\u89d2", "\u4e00\u7ad9\u5f0f\u4f01\u4e1a\u670d\u52a1"],
          ["\u4e94\u3001\u516c\u4f17\u7528\u6237\u89c6\u89d2", "\u4fbf\u6c11\u4fe1\u606f\u670d\u52a1"],
          ["\u516d\u3001\u6848\u4f8b\u4e0e\u5b9e\u8df5", ""],
        ].map((t, i) => new Paragraph({
          spacing: { before: 120 },
          tabStops: [{ type: TabStopType.RIGHT, position: TabStopPosition.MAX }],
          children: [
            new TextRun({ text: t[0], bold: true, size: 22, font: "Microsoft YaHei", color: C.text }),
            t[1] ? new TextRun({ text: "\t" + t[1], size: 18, font: "Microsoft YaHei", color: C.gray }) : new TextRun({ text: "", size: 18 }),
          ],
        })),
        new Paragraph({ children: [new PageBreak()] }),

        // ===== SECTION 1: OVERVIEW =====
        new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun({ text: "\u4e00\u3001\u4ea7\u54c1\u6982\u8ff0", bold: true })] }),
        bodyText("\u9752\u5ddd\u53bf\u5e84\u5b50\u4e0a\u5de5\u4e1a\u56ed\u533a\u667a\u6167\u5e73\u53f0\u662f\u4e00\u4e2a\u96c6AI\u62db\u5546\u3001\u5168\u57df\u5b89\u5168\u76d1\u63a7\u3001\u4f01\u4e1a\u670d\u52a1\u3001\u8fd0\u8425\u7ba1\u7406\u3001\u80fd\u6548\u76d1\u6d4b\u7b49\u591a\u6a21\u5757\u4e8e\u4e00\u4f53\u7684\u7efc\u5408\u6027\u667a\u6167\u56ed\u533a\u89e3\u51b3\u65b9\u6848\u3002\u5e73\u53f0\u4e3a\u56ed\u533a\u653f\u5e9c\u3001\u7ba1\u7406\u65b9\u3001\u5165\u9a7b\u4f01\u4e1a\u548c\u516c\u4f17\u7528\u6237\u56db\u7c7b\u89d2\u8272\u63d0\u4f9b\u5b9a\u5236\u5316\u7684\u529f\u80fd\u4e0e\u754c\u9762\u3002"),
        bodyText("\u5e73\u53f0\u91c7\u7528\u4e94\u5c42\u67b6\u6784\uff1aC\u7aef\u5c55\u793a\u5c42\u3001\u5e94\u7528\u670d\u52a1\u5c42\u3001\u6570\u636e\u670d\u52a1\u5c42\u3001\u8bbe\u65bd\u5c42\u548c\u6570\u636e\u5bf9\u63a5\u5c42\uff0c\u5bf9\u63a512\u5927\u4e09\u65b9\u6570\u636e\u6e90\u3002\u56ed\u533a\u89c4\u5212\u9762\u79ef2.1\u33a1\uff0c\u5df2\u5165\u9a7b47\u5bb6\u4f01\u4e1a\uff0c\u4ece\u4e1a\u4eba\u54583,000+\u4eba\u3002"),

        // ====== \u67b6\u6784\u8868 ======
        sectionHeader("\u5e73\u53f0\u67b6\u6784", C.gov),
        bodyText("\u5e73\u53f0\u5206\u4e3a\u4e94\u5c42\u67b6\u6784\uff0c\u4ece\u5e95\u5c42\u6570\u636e\u5bf9\u63a5\u5230\u9876\u5c42\u5c55\u793a\uff1a"),
        ...[
          ["C\u7aef\u5c55\u793a\u5c42", "\u5927\u5c4f\u6570\u636e\u9a7e\u9a76\u8231 | \u56ed\u533a\u5b98\u7f51 | \u5411\u5bfc\u9875 | H5\u79fb\u52a8\u7aef | \u5fae\u4fe1\u5c0f\u7a0b\u5e8f", C.gov],
          ["\u5e94\u7528\u670d\u52a1\u5c42", "AI\u62db\u5546 | \u5b89\u5168\u76d1\u63a7 | \u4f01\u4e1a\u670d\u52a1 | \u8fd0\u8425\u7ba1\u7406 | \u80fd\u6548\u76d1\u6d4b | \u6570\u5b57\u5316\u529e\u516c", C.mgmt],
          ["\u6570\u636e\u670d\u52a1\u5c42", "\u6570\u636e\u6e56 | \u6570\u636e\u5206\u6790 | \u62a5\u8868\u5f15\u64ce | \u5b9e\u65f6\u8ba1\u7b97 | AI\u7b97\u6cd5 | GIS\u5e73\u53f0", C.ent],
          ["\u8bbe\u65bd\u5c42", "\u89c6\u9891\u76d1\u63a7 | \u4f20\u611f\u5668 | \u667a\u80fd\u8868\u8ba1 | GPS\u5b9a\u4f4d | \u73af\u5883\u76d1\u6d4b", C.pub],
          ["\u6570\u636e\u5bf9\u63a5\u5c42", "\u5929\u7136\u6c14 | \u7535\u529b | \u81ea\u6765\u6c34 | \u7efc\u5408\u6267\u6cd5 | \u73af\u4fdd | \u5e94\u6025 | \u4ea4\u901a | \u516c\u5b89 | \u81ea\u7136\u8d44\u6e90 | \u6c14\u8c61 | \u4eba\u793e | \u8fd0\u8425\u5546", "475569"],
        ].map((a, i) => {
          const table = new Table({
            width: { size: 9026, type: WidthType.DXA },
            columnWidths: [1800, 7226],
            rows: [
              new TableRow({
                children: [
                  new TableCell({
                    borders, width: { size: 1800, type: WidthType.DXA },
                    shading: { fill: a[2], type: ShadingType.CLEAR },
                    margins: { top: 40, bottom: 40, left: 80, right: 80 },
                    verticalAlign: "center",
                    children: [new Paragraph({
                      children: [new TextRun({ text: a[0], bold: true, size: 18, font: "Microsoft YaHei", color: "FFFFFF" })],
                    })],
                  }),
                  new TableCell({
                    borders, width: { size: 7226, type: WidthType.DXA },
                    margins: { top: 40, bottom: 40, left: 80, right: 80 },
                    verticalAlign: "center",
                    children: [new Paragraph({
                      children: [new TextRun({ text: a[1], size: 18, font: "Microsoft YaHei", color: C.text })],
                    })],
                  }),
                ],
              }),
            ],
          });
          return i === 0 ? [new Paragraph({ spacing: { before: 120 } }), table] : table;
        }).flat(),

        new Paragraph({ children: [new PageBreak()] }),

        // ===== SECTION 2: GOVERNMENT =====
        new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun({ text: "\u4e8c\u3001\u653f\u5e9c\u9886\u5bfc\u5c42\u89c6\u89d2", color: C.gov, bold: true })] }),
        bodyText("\u4e3a\u56ed\u533a\u653f\u5e9c\u7ba1\u7406\u90e8\u95e8\u63d0\u4f9b\u6570\u636e\u9a71\u52a8\u7684\u51b3\u7b56\u652f\u6301\uff0c\u6db5\u76d6\u62db\u5546\u5f15\u8d44\u3001\u4ea7\u4e1a\u89c4\u5212\u3001\u5b89\u5168\u76d1\u7ba1\u3001\u73af\u4fdd\u76d1\u63a7\u7b49\u6838\u5fc3\u80fd\u529b\u3002"),

        sectionHeader("AI\u62db\u5546\u670d\u52a1", C.gov),
        featItem("AI\u667a\u80fd\u62db\u5546", "\u5927\u6570\u636e+AI\u7b97\u6cd5\u7cbe\u51c6\u5339\u914d\u76ee\u6807\u4f01\u4e1a\uff0c\u667a\u80fd\u63a8\u8350\u4e0e\u8ddf\u8e2a", C.gov),
        featDesc("\u591a\u7ef4\u6570\u636e\u6e90+\u667a\u80fd\u7b97\u6cd5\u81ea\u52a8\u63a8\u8350\u4e0e\u56ed\u533a\u4ea7\u4e1a\u9ad8\u5ea6\u5339\u914d\u7684\u76ee\u6807\u4f01\u4e1a\uff0c\u5b9e\u73b0\u7cbe\u51c6\u62db\u5546\u5f15\u8d44\u3002"),
        featItem("\u6570\u5b57\u6c99\u76d8", "3D\u53ef\u89c6\u5316\u5c55\u793a\u56ed\u533a\u89c4\u5212\u4e0e\u571f\u5730\u8d44\u6e90", C.gov),
        featDesc("\u6c89\u6d78\u5f0f\u62db\u5546\u5c55\u793a\uff0c\u5e2e\u52a9\u6f5c\u5728\u6295\u8d44\u8005\u76f4\u89c2\u4e86\u89e3\u56ed\u533a\u89c4\u5212\u3002"),
        featItem("\u4ea7\u4e1a\u94fe\u56fe\u8c31", "\u53ef\u89c6\u5316\u5c55\u793a\u4ea7\u4e1a\u94fe\u4e0a\u4e0b\u6e38\u5173\u7cfb", C.gov),
        featDesc("\u7cbe\u51c6\u5f3a\u94fe\u8865\u94fe\uff0c\u4e3a\u4ea7\u4e1a\u89c4\u5212\u548c\u62db\u5546\u63d0\u4f9b\u51b3\u7b56\u4f9d\u636e\u3002"),
        featItem("\u76ee\u6807\u4f01\u4e1a\u6316\u6398", "\u591a\u7ef4\u5ea6\u7b97\u6cd5\u8bc4\u5206\u4e0e\u667a\u80fd\u63a8\u8350", C.gov),
        featItem("\u6295\u8d44\u9879\u76ee\u7ba1\u7406", "\u62db\u5546\u9879\u76ee\u5168\u751f\u547d\u5468\u671f\u7ba1\u7406", C.gov),
        featItem("\u62db\u5546\u5165\u53e3/\u4e91\u5c55\u9986", "\u4f01\u4e1a\u5f62\u8c61\u7ebf\u4e0a\u5c55\u793a", C.gov),

        sectionHeader("\u5168\u57df\u5b89\u5168\u76d1\u63a7\u4f53\u7cfb", C.gov),
        featItem("\u5929\u773c - \u9053\u8def\u4ea4\u901a\u5b89\u5168\u76d1\u63a7", "\u667a\u80fd\u4ea4\u901a\u76d1\u63a7\u3001\u8f66\u8f86\u8bc6\u522b\u3001\u8fdd\u7ae0\u6355\u6349", C.gov),
        featItem("\u54e8\u5175 - \u5730\u707e\u5b89\u5168\u76d1\u6d4b", "\u5730\u8d28\u707e\u5bb3\u5b9e\u65f6\u76d1\u6d4b\u3001\u591a\u4f20\u611f\u5668\u8054\u52a8\u9884\u8b66", C.gov),
        featItem("\u6167\u773c - \u73af\u4fdd\u76d1\u63a7\u68c0\u6d4b", "\u73af\u5883\u6307\u6807\u5b9e\u65f6\u76d1\u6d4b\u3001\u6392\u653e\u6570\u636e\u8ffd\u8e2a", C.gov),
        featItem("\u89c6\u9891\u76d1\u63a7\u603b\u89c8", "\u7edf\u4e00\u89c6\u9891\u5e73\u53f0\uff0c16\u5bab\u683c\u667a\u80fd\u6392\u5217", C.gov),
        featItem("\u5e94\u6025\u6307\u6325\u8054\u52a8", "\u4e00\u952e\u8c03\u5ea6\u3001\u9884\u6848\u7ba1\u7406\u3001\u5b9e\u65f6\u901a\u8baf", C.gov),

        sectionHeader("\u6295\u8d44\u6982\u7b97\u4e0e\u521b\u5efa\u8def\u5f84", C.gov),
        bodyText("\u5e73\u53f0\u652f\u6301\u56ed\u533a\u6cbf\u7740\u201c\u7701\u7ea7\u7ecf\u6d4e\u5f00\u53d1\u533a\u2192\u7eff\u8272\u5de5\u4e1a\u56ed\u533a\u2192\u56fd\u5bb6\u7ea7\u7ecf\u6d4e\u5f00\u53d1\u533a\u57f9\u80b2\u5bf9\u8c61\u201d\u7684\u521b\u5efa\u8def\u5f84\u524d\u8fdb\uff0c\u63d0\u4f9b\u6570\u636e\u652f\u6491\u4e0e\u7ba1\u7406\u5de5\u5177\u3002"),

        new Paragraph({ children: [new PageBreak()] }),

        // ===== SECTION 3: MANAGEMENT =====
        new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun({ text: "\u4e09\u3001\u56ed\u533a\u7ba1\u7406\u5c42\u89c6\u89d2", color: C.mgmt, bold: true })] }),
        bodyText("\u4e3a\u56ed\u533a\u65e5\u5e38\u8fd0\u8425\u7ba1\u7406\u4eba\u5458\u63d0\u4f9b\u5168\u9762\u7684\u667a\u6167\u8fd0\u8425\u5de5\u5177\uff0c\u6db5\u76d6\u5b89\u5168\u76d1\u63a7\u3001\u80fd\u6548\u7ba1\u63a7\u3001\u8bbe\u5907\u7ba1\u7406\u3001\u6570\u636e\u5206\u6790\u7b49\u65e5\u5e38\u8fd0\u7ef4\u80fd\u529b\u3002"),

        sectionHeader("\u667a\u6167\u8fd0\u8425\u4e0e\u76d1\u63a7", C.mgmt),
        featItem("\u9053\u8def\u4ea4\u901a\u76d1\u63a7", "\u8f66\u8f86\u8bc6\u522b\u3001\u8fdd\u7ae0\u6355\u6349\u3001\u4ea4\u901a\u72b6\u6001\u53ef\u89c6\u5316", C.mgmt),
        featItem("\u5730\u707e\u5b89\u5168\u76d1\u6d4b", "\u5730\u8d28\u707e\u5bb3\u5b9e\u65f6\u76d1\u6d4b\u3001\u591a\u4f20\u611f\u5668\u8054\u52a8\u9884\u8b66", C.mgmt),
        featItem("\u73af\u4fdd\u76d1\u63a7\u68c0\u6d4b", "\u73af\u5883\u6307\u6807\u5b9e\u65f6\u76d1\u6d4b\u3001\u6392\u653e\u6570\u636e\u8ffd\u8e2a", C.mgmt),
        featItem("\u89c6\u9891\u76d1\u63a7\u603b\u89c8", "\u7edf\u4e00\u89c6\u9891\u5e73\u53f0\u3001\u667a\u80fd\u6392\u5217\u3001\u5168\u56ed\u533a\u8986\u76d6", C.mgmt),
        featItem("\u5e94\u6025\u6307\u6325\u8054\u52a8", "\u4e00\u952e\u8c03\u5ea6\u3001\u9884\u6848\u7ba1\u7406\u3001\u5b9e\u65f6\u901a\u8baf", C.mgmt),

        sectionHeader("\u80fd\u6548\u76d1\u6d4b\u4e0e\u4ea9\u5747\u8bba\u82f1\u96c4", C.mgmt),
        featItem("\u4f01\u4e1a\u7528\u80fd\u76d1\u6d4b", "\u667a\u80fd\u8ba1\u91cf\u8868\u5b9e\u65f6\u91c7\u96c6\u80fd\u8017\u6570\u636e", C.mgmt),
        featDesc("\u652f\u6301\u6c34\u3001\u7535\u3001\u6c14\u591a\u7ef4\u5ea6\u76d1\u63a7\uff0c\u53ef\u89c6\u5316\u5c55\u793a\u80fd\u8017\u8d8b\u52bf\u4e0e\u5f02\u5e38\u544a\u8b66\u3002"),
        featItem("\u4ea9\u5747\u8bba\u82f1\u96c4", "\u5927\u6570\u636e\u8bc4\u4ef7\u4f53\u7cfb\uff0c\u8d44\u6e90\u8981\u7d20\u5dee\u5f02\u5316\u914d\u7f6e", C.mgmt),
        featItem("\u80fd\u6548\u5206\u6790", "\u4f01\u4e1a\u80fd\u8017\u6548\u7387\u591a\u7ef4\u5ea6\u5206\u6790\uff0c\u540c\u6bd4\u73af\u6bd4\u8d8b\u52bf\u53ef\u89c6\u5316", C.mgmt),
        featItem("\u5f02\u5e38\u9884\u8b66", "\u7528\u80fd\u5f02\u5e38\u81ea\u52a8\u544a\u8b66\uff0c\u78b3\u6392\u7ba1\u7406\u8ffd\u8e2a", C.mgmt),

        sectionHeader("\u8fd0\u8425\u7ba1\u7406\u4e0e\u6570\u5b57\u5316\u529e\u516c", C.mgmt),
        featItem("\u8fd0\u8425\u670d\u52a1\u4e2d\u5fc3", "\u5927\u5c4f\u6307\u6325+\u56ed\u5c0f\u4e8c\u56e2\u961f+\u957f\u6548\u670d\u52a1", C.mgmt),
        featItem("\u65e5\u5e38\u8fd0\u8425\u7ba1\u7406", "\u6570\u5b57\u5316\u5ba1\u6279\u3001\u65e5\u5e38\u5de1\u68c0\u3001\u8d44\u4ea7\u7ba1\u7406\u3001\u6863\u6848\u7ba1\u7406", C.mgmt),
        featItem("\u8bbe\u5907\u4e0e\u7269\u8054\u7f51", "\u6444\u50cf\u5934\u3001\u4f20\u611f\u5668\u3001\u667a\u80fd\u8bbe\u5907\u7edf\u4e00\u7ba1\u63a7", C.mgmt),
        featItem("\u6570\u636e\u4e2d\u5fc3", "\u6c47\u805a\u56ed\u533a\u5168\u91cf\u6570\u636e\uff0c\u6784\u5efa\u7edf\u4e00\u6570\u636e\u6e56", C.mgmt),
        featItem("\u6570\u636e\u8d28\u91cf\u4e0e\u62a5\u8868", "\u6570\u636e\u8d28\u91cf\u76d1\u63a7\u3001\u81ea\u5b9a\u4e49\u62a5\u8868\u3001\u591a\u7ef4\u5ea6\u5206\u6790", C.mgmt),

        new Paragraph({ children: [new PageBreak()] }),

        // ===== SECTION 4: ENTERPRISE =====
        new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun({ text: "\u56db\u3001\u5165\u9a7b\u4f01\u4e1a\u89c6\u89d2", color: C.ent, bold: true })] }),
        bodyText("\u4e3a\u56ed\u533a\u5165\u9a7b\u4f01\u4e1a\u63d0\u4f9b\u5168\u6d41\u7a0b\u6570\u5b57\u5316\u670d\u52a1\uff0c\u6db5\u76d6\u8bc9\u6c42\u5904\u7406\u3001\u7f34\u8d39\u670d\u52a1\u3001\u4f9b\u9700\u5bf9\u63a5\u3001\u573a\u9986\u9910\u996e\u3001\u7528\u5de5\u4fe1\u606f\u7b49\u591a\u4e2a\u7ef4\u5ea6\u3002"),

        sectionHeader("\u4f01\u4e1a\u670d\u52a1\u4e2d\u5fc3", C.ent),
        featItem("\u8bc9\u6c42\u76f4\u8fbe\uff08\u968f\u624b\u62cd\uff09", "\u6587\u5b57+\u56fe\u7247+\u5b9a\u4f4d\u5feb\u901f\u63d0\u4ea4\u8bc9\u6c42\uff0c2\u5c0f\u65f6\u54cd\u5e94", C.ent),
        featItem("\u670d\u52a1\u76f4\u8fbe", "\u5929\u7136\u6c14/\u6c34/\u7535\u8d39\u5728\u7ebf\u7f34\u8d39\uff0c\u4e00\u7ad9\u5f0f\u670d\u52a1", C.ent),
        featItem("\u56ed\u4f01\u4ea4\u4e92", "\u5728\u7ebf\u6c9f\u901a\u3001\u901a\u77e5\u516c\u544a\u3001\u653f\u7b56\u63a8\u9001\u3001\u53cc\u5411\u4e92\u52a8", C.ent),
        featItem("\u4f01\u4e1a\u5165\u9a7b", "\u7ebf\u4e0a\u7533\u8bf7\u3001\u8bc1\u7167\u7ba1\u7406\u3001\u5165\u9a7b\u5168\u6d41\u7a0b\u7ebf\u4e0a\u5316", C.ent),
        featItem("\u4fe1\u7528\u8bc4\u4ef7", "\u4f01\u4e1a\u4fe1\u7528\u591a\u7ef4\u5ea6\u8bc4\u4ef7\u4f53\u7cfb\uff0c\u52a9\u529b\u878d\u8d44\u5408\u4f5c", C.ent),

        sectionHeader("\u4f9b\u9700\u5bf9\u63a5\u4e0e\u7269\u6d41\u670d\u52a1", C.ent),
        featItem("\u7269\u6d41\u4f9b\u9700", "\u56ed\u533a\u8d27\u6e90+\u8fd4\u7a0b\u7a7a\u8f66+\u8f66\u8f86\u8d44\u6e90\u667a\u80fd\u5339\u914d", C.ent),
        featItem("\u7279\u8272\u4f9b\u9700", "\u9752\u5ddd\u5c71\u73cd\u7b49\u7279\u8272\u4ea7\u54c1\u5c55\u793a\u4e0e\u4f9b\u9700\u5bf9\u63a5", C.ent),
        featItem("\u653f\u7b56\u63a8\u9001", "\u60e0\u4f01\u653f\u7b56\u7cbe\u51c6\u5339\u914d\u4e0e\u4e3b\u52a8\u63a8\u9001", C.ent),
        featItem("\u505c\u8f66\u670d\u52a1", "\u9ad8\u5fb7\u5730\u56fe\u5bfc\u822a+\u5b9e\u65f6\u4f59\u4f4d+\u5728\u7ebf\u7f34\u8d39", C.ent),

        sectionHeader("\u573a\u9986\u3001\u6d3b\u52a8\u4e0e\u9910\u996e", C.ent),
        featItem("\u573a\u9986\u670d\u52a1", "8\u5927\u573a\u9986\uff08\u4f1a\u8bae\u5385/\u5c55\u5385/\u4f53\u80b2\u9986/\u57f9\u8bad\u6559\u5ba4\u7b49\uff09\u5728\u7ebf\u9884\u7ea6", C.ent),
        featItem("\u56ed\u533a\u6d3b\u52a8", "\u6587\u4f53/\u57f9\u8bad/\u516c\u76ca/\u4ea4\u6d41\u7b49\u591a\u7c7b\u6d3b\u52a8\u53c2\u4e0e", C.ent),
        featItem("\u56ed\u533a\u9910\u996e", "6\u5bb6\u9910\u5385\u300120+\u9053\u83dc\u54c1\u4e00\u89c8\uff0c\u652f\u6301\u5206\u7c7b\u7b5b\u9009", C.ent),
        featItem("\u7528\u5de5\u4fe1\u606f", "\u62db\u8058+\u5c97\u4f4d\u5339\u914d\uff0c\u5e2e\u52a9\u4f01\u4e1a\u89e3\u51b3\u7528\u5de5\u95ee\u9898", C.ent),

        sectionHeader("\u7f34\u8d39\u670d\u52a1\u4e0e\u5ba1\u6279\u6d41\u7a0b", C.ent),
        featItem("\u5929\u7136\u6c14\u7f34\u8d39", "\u5728\u7ebf\u67e5\u770b\u6c14\u8d39\u8d26\u5355\u3001\u652f\u4ed8\u5386\u53f2\u8bb0\u5f55", C.ent),
        featItem("\u81ea\u6765\u6c34\u7f34\u8d39", "\u5b9e\u65f6\u67e5\u770b\u7528\u6c34\u91cf\u3001\u6c34\u8d39\u8d26\u5355", C.ent),
        featItem("\u7535\u8d39\u7f34\u8d39", "\u667a\u80fd\u8868\u8ba1\u5b9e\u65f6\u76d1\u63a7\u7528\u7535\u91cf\uff0c\u5728\u7ebf\u7f34\u8d39", C.ent),
        featItem("\u5ba1\u6279\u6d41\u7a0b", "\u56ed\u533a\u5ba1\u6279\u4e2d\u5fc3\u3001\u9879\u76ee\u70b9\u5ba1\u6279\u3001\u65e5\u62a5\u7ba1\u7406\u7b49", C.ent),

        new Paragraph({ children: [new PageBreak()] }),

        // ===== SECTION 5: PUBLIC =====
        new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun({ text: "\u4e94\u3001\u516c\u4f17\u7528\u6237\u89c6\u89d2", color: C.pub, bold: true })] }),
        bodyText("\u4e3a\u56ed\u533a\u516c\u4f17\u63d0\u4f9b\u4fbf\u6c11\u4fe1\u606f\u670d\u52a1\uff0c\u5305\u62ec\u56ed\u533a\u516c\u544a\u3001\u6d3b\u52a8\u4fe1\u606f\u3001\u5546\u4e1a\u670d\u52a1\u3001\u7528\u5de5\u4fe1\u606f\u7b49\u3002"),

        sectionHeader("\u4fe1\u606f\u670d\u52a1\u4e0e\u56ed\u533a\u53c2\u4e0e", C.pub),
        featItem("\u56ed\u533a\u516c\u544a", "\u5b9e\u65f6\u67e5\u770b\u56ed\u533a\u901a\u77e5\uff08\u65bd\u5de5\u3001\u68c0\u67e5\u3001\u7ef4\u62a4\u7b49\uff09", C.pub),
        featItem("\u56ed\u533a\u6d3b\u52a8", "\u67e5\u770b\u5e76\u53c2\u4e0e\u6587\u4f53/\u516c\u76ca\u7b49\u6d3b\u52a8\uff0c\u8ddf\u8e2a\u6d3b\u52a8\u72b6\u6001", C.pub),
        featItem("\u56ed\u533a\u653f\u7b56", "\u67e5\u770b\u516c\u5f00\u653f\u7b56\u4fe1\u606f\uff0c\u4e86\u89e3\u56ed\u533a\u53d1\u5c55\u52a8\u6001", C.pub),
        featItem("\u4fe1\u606f\u53d1\u5e03", "\u56ed\u533a\u53d1\u5c55\u6210\u5c31\u3001\u4f18\u52bf\u4ea7\u4e1a\u591a\u7ef4\u5ea6\u5c55\u793a", C.pub),

        sectionHeader("\u751f\u6d3b\u670d\u52a1\u4e0e\u4e92\u52a8", C.pub),
        featItem("\u56ed\u533a\u9910\u996e", "6\u5bb6\u9910\u5385\u4fe1\u606f+\u4eca\u65e5\u83dc\u5355+\u4ef7\u683c+\u5206\u7c7b\u7b5b\u9009", C.pub),
        featItem("\u505c\u8f66\u670d\u52a1", "\u5730\u56fe\u67e5\u770b\u505c\u8f66\u4f4d\u5b9e\u65f6\u4f59\u4f4d\uff0c\u5728\u7ebf\u7f34\u8d39", C.pub),
        featItem("\u573a\u9986\u670d\u52a1", "\u6d4f\u89c8\u56ed\u533a\u573a\u9986\u4fe1\u606f\uff0c\u4e86\u89e3\u5f00\u653e\u65f6\u95f4\u4e0e\u9884\u7ea6", C.pub),
        featItem("\u7279\u8272\u4f9b\u9700", "\u6d4f\u89c8\u9752\u5ddd\u5c71\u73cd\u7b49\u7279\u8272\u4ea7\u54c1\uff0c\u76f4\u63a5\u8054\u7cfb\u4f9b\u5e94\u5546", C.pub),
        featItem("\u7528\u5de5\u4fe1\u606f", "\u67e5\u770b\u56ed\u533a\u4f01\u4e1a\u62db\u8058\u4fe1\u606f\uff0c\u5c97\u4f4d\u5339\u914d", C.pub),
        featItem("\u5fae\u4fe1\u5c0f\u7a0b\u5e8f", "\u5feb\u901f\u63a5\u5165\u667a\u6167\u5e73\u53f0\uff0c\u4fbf\u6377\u670d\u52a1\u5165\u53e3", C.pub),

        new Paragraph({ children: [new PageBreak()] }),

        // ===== SECTION 6: CASES =====
        new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun({ text: "\u516d\u3001\u6848\u4f8b\u4e0e\u5b9e\u8df5", bold: true })] }),
        bodyText("\u5e73\u53f0\u7d2f\u8ba1\u670d\u52a118+\u4e2a\u786e\u8ba4\u6848\u4f8b\uff0c\u8986\u76d6\u7802\u77f3\u884c\u4e1a\u667a\u80fd\u5316\u7ba1\u63a7\u591a\u4e2a\u573a\u666f\u3002"),
        ...[
          ["\u7ef5\u7af9\u5efa\u6295\u5efa\u6750", "\u957f\u6c5f\u6cb3\u9053\u7802\u77f3\u7535\u5b50\u91c7\u8fd0\u7ba1\u7406\u5355\u670d\u52a1", C.mgmt],
          ["\u7ef5\u7af9\u5efa\u6295\u5efa\u6750", "\u5806\u6599\u573a\u96f7\u8fbe\u6d4b\u91cf\u9879\u76ee", C.mgmt],
          ["\u4e2d\u56fd\u79fb\u52a8\u5fb7\u9633\u5206\u516c\u53f8", "\u7ef5\u7af9\u5e02\u91c7\u7802\u89c6\u9891\u76d1\u63a7\u548c\u7535\u5b50\u56f4\u680f\u9879\u76ee", C.gov],
          ["\u6210\u90fd\u7a7a\u6e2f\u5174\u57ce\u5efa\u6750", "\u7802\u77f3\u5806\u573a\u667a\u6167\u5316\u7ba1\u7406\u5e73\u53f0\u5efa\u8bbe", C.ent],
          ["\u7ef5\u7af9\u5efa\u6295\u5efa\u6750", "\u7802\u77f3\u8d44\u6e90\u7ecf\u8425\u7ba1\u7406\u5927\u6570\u636e\u5e73\u53f0\u8fd0\u7ef4", C.mgmt],
          ["\u7ef5\u7af9\u5efa\u6295\u5efa\u6750", "\u96f6\u78b3\u8fd0\u8f93\u8f66\u8f86\u5b9a\u4f4d\u4e0e\u8f68\u8ff9\u5f00\u53d1", C.mgmt],
        ].map(c => new Paragraph({
          numbering: { reference: "bullets", level: 0 },
          spacing: { before: 60 },
          children: [
            new TextRun({ text: c[0], bold: true, size: 20, font: "Microsoft YaHei", color: c[2] }),
            new TextRun({ text: " - " + c[1], size: 20, font: "Microsoft YaHei", color: C.text }),
          ],
        })),

        new Paragraph({ spacing: { before: 400 } }),
        bodyText("\u5e73\u53f0\u6301\u7eed\u8fdb\u5316\uff0c\u672a\u6765\u5c06\u6df1\u5316AI\u591a\u573a\u666f\u5e94\u7528\u3001\u6570\u5b57\u5b6a\u751f\u5168\u8986\u76d6\u3001\u8de8\u56ed\u533a\u534f\u540c\u7b49\u65b9\u5411\uff0c\u9a71\u52a8\u56ed\u533a\u6cbf\u7740\u6570\u5b57\u5316\u3001\u667a\u80fd\u5316\u65b9\u5411\u5168\u901f\u524d\u8fdb\u3002"),
      ],
    },
  ],
});

Packer.toBuffer(doc).then(buffer => {
  const out = "/Users/wangcaihua/WorkBuddy/20260427142154/slides/output/\xe9\x9d\x92\xe5\xb7\x9d\xe5\x8e\xbf\xe5\xba\x84\xe5\xad\x90\xe4\xb8\x8a\xe5\xb7\xa5\xe4\xb8\x9a\xe5\x9b\xad\xe5\x8c\xba\xe6\x99\xba\xe6\x85\xa7\xe5\xb9\xb3\xe5\x8f\xb0\xe4\xba\xa7\xe5\x93\x81\xe6\x89\x8b\xe5\x86\x8c.docx";
  fs.writeFileSync(out, buffer);
  console.log("DOCX OK: " + fs.statSync(out).size + " bytes at " + out);
});
