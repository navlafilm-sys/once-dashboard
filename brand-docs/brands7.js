const fs = require('fs');
const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
        Header, Footer, AlignmentType, LevelFormat,
        HeadingLevel, BorderStyle, WidthType, ShadingType,
        PageNumber, PageBreak } = require('docx');

const border = { style: BorderStyle.SINGLE, size: 1, color: "DDDDDD" };
const borders = { top: border, bottom: border, left: border, right: border };
const cm = { top: 80, bottom: 80, left: 120, right: 120 };

function hCell(text, width) {
  return new TableCell({
    borders, width: { size: width, type: WidthType.DXA },
    shading: { fill: "1B1B1B", type: ShadingType.CLEAR },
    margins: cm, verticalAlign: "center",
    children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text, bold: true, color: "FFFFFF", font: "Arial", size: 19 })] })]
  });
}
function bCell(text, width, fill) {
  return new TableCell({
    borders, width: { size: width, type: WidthType.DXA },
    shading: fill ? { fill, type: ShadingType.CLEAR } : undefined,
    margins: cm,
    children: [new Paragraph({ spacing: { after: 40 }, children: [new TextRun({ text, font: "Arial", size: 19 })] })]
  });
}
function bCellBold(text, width, fill) {
  return new TableCell({
    borders, width: { size: width, type: WidthType.DXA },
    shading: fill ? { fill, type: ShadingType.CLEAR } : undefined,
    margins: cm,
    children: [new Paragraph({ spacing: { after: 40 }, children: [new TextRun({ text, font: "Arial", size: 19, bold: true })] })]
  });
}

function h1(t) { return new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun(t)] }); }
function h2(t) { return new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun(t)] }); }
function p(t, opts={}) { return new Paragraph({ spacing: { after: opts.after||120 }, children: [new TextRun({ text: t, size: 22, font: "Arial", ...opts })] }); }
function pb(label, text) {
  return new Paragraph({ spacing: { after: 100 }, children: [
    new TextRun({ text: label, size: 22, font: "Arial", bold: true }),
    new TextRun({ text, size: 22, font: "Arial" }),
  ]});
}
function bullet(label, text, ref) {
  return new Paragraph({ numbering: { reference: ref||"bullets", level: 0 }, spacing: { after: 80 }, children: [
    new TextRun({ text: label, size: 22, font: "Arial", bold: true }),
    new TextRun({ text, size: 22, font: "Arial" }),
  ]});
}

function brandBlock(num, name, subtitle, accent, rows) {
  return [
    h2(`${num}. ${name}`),
    new Paragraph({ spacing: { after: 80 }, children: [new TextRun({ text: subtitle, size: 20, color: "888888", italics: true })] }),
    new Table({
      width: { size: 9840, type: WidthType.DXA },
      columnWidths: [2460, 7380],
      rows: rows.map((r, i) => new TableRow({ children: [
        bCellBold(r[0], 2460, i%2===0 ? accent : undefined),
        bCell(r[1], 7380, i%2===0 ? accent : undefined),
      ]}))
    }),
  ];
}

const doc = new Document({
  styles: {
    default: { document: { run: { font: "Arial", size: 22 } } },
    paragraphStyles: [
      { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 34, bold: true, font: "Arial", color: "1B1B1B" },
        paragraph: { spacing: { before: 360, after: 200 }, outlineLevel: 0 } },
      { id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 26, bold: true, font: "Arial", color: "333333" },
        paragraph: { spacing: { before: 280, after: 140 }, outlineLevel: 1 } },
    ]
  },
  numbering: {
    config: [
      { reference: "bullets", levels: [{ level: 0, format: LevelFormat.BULLET, text: "\u2022", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] },
      { reference: "nums", levels: [{ level: 0, format: LevelFormat.DECIMAL, text: "%1.", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] },
    ]
  },
  sections: [{
    properties: {
      page: { size: { width: 12240, height: 15840 }, margin: { top: 1440, right: 1200, bottom: 1440, left: 1200 } }
    },
    headers: { default: new Header({ children: [new Paragraph({ alignment: AlignmentType.RIGHT, children: [new TextRun({ text: "\uC628\uC138\uC2A4\uD29C\uB514\uC624 \uC2E0\uADDC 7\uAC1C \uBE0C\uB79C\uB4DC \uAE30\uD68D\uC11C", italics: true, color: "999999", size: 17 })] })] }) },
    footers: { default: new Footer({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "Page ", size: 17, color: "999999" }), new TextRun({ children: [PageNumber.CURRENT], size: 17, color: "999999" })] })] }) },
    children: [

      // ===== COVER =====
      new Paragraph({ spacing: { before: 2400 }, children: [] }),
      new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 80 }, children: [new TextRun({ text: "\uC628\uC138\uC2A4\uD29C\uB514\uC624", size: 48, bold: true, color: "1B1B1B" })] }),
      new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 100 }, children: [new TextRun({ text: "\uC2E0\uADDC 7\uAC1C \uBE0C\uB79C\uB4DC \uAE30\uD68D\uC11C", size: 40, bold: true, color: "E04040" })] }),
      new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 300 }, border: { top: { style: BorderStyle.SINGLE, size: 4, color: "E04040", space: 1 } }, children: [] }),
      new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 80 }, children: [new TextRun({ text: "\uAE30\uC874 3\uAC1C \uBE0C\uB79C\uB4DC + \uC2E0\uADDC 7\uAC1C = \uCD1D 10\uAC1C \uBE0C\uB79C\uB4DC \uD3EC\uD2B8\uD3F4\uB9AC\uC624", size: 22, color: "888888" })] }),
      new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "\uAC10\uC131 \uC2A4\uBAB0\uBE0C\uB79C\uB4DC 10\uAC1C \uB808\uD37C\uB7F0\uC2A4 \uAE30\uBC18 | 2026\uB144 4\uC6D4", size: 20, color: "AAAAAA" })] }),

      new Paragraph({ children: [new PageBreak()] }),

      // ===== 개요 =====
      h1("\uAC1C\uC694: 7\uAC1C \uBE0C\uB79C\uB4DC \uC804\uCCB4 \uC124\uACC4"),
      p("\uAC10\uC131 \uC2A4\uBAB0\uBE0C\uB79C\uB4DC 10\uAC1C \uB808\uD37C\uB7F0\uC2A4\uB97C \uAE30\uBC18\uC73C\uB85C, \uC628\uC138\uC2A4\uD29C\uB514\uC624 \uBA54\uD0C0\uAD11\uACE0\uC5D0 \uC801\uC6A9\uD560 7\uAC1C \uC2E0\uADDC \uBE0C\uB79C\uB4DC/\uB77C\uC778\uC744 \uC124\uACC4\uD588\uC2B5\uB2C8\uB2E4. \uAC01 \uBE0C\uB79C\uB4DC\uB294 \uB3C5\uB9BD\uC801\uC778 \uB514\uC790\uC778 \uC815\uCCB4\uC131\uACFC \uD0C0\uAC9F\uC744 \uAC00\uC9C0\uBA70, \uBA54\uD0C0\uAD11\uACE0\uC5D0\uC11C \uBC14\uB85C \uC9D1\uD589 \uAC00\uB2A5\uD55C \uC218\uC900\uC73C\uB85C \uAD6C\uCCB4\uD654\uD588\uC2B5\uB2C8\uB2E4."),

      new Table({
        width: { size: 9840, type: WidthType.DXA },
        columnWidths: [600, 2000, 2400, 2440, 2400],
        rows: [
          new TableRow({ children: [ hCell("#", 600), hCell("\uBE0C\uB79C\uB4DC\uBA85", 2000), hCell("\uCEE8\uC149", 2400), hCell("\uB808\uD37C\uB7F0\uC2A4", 2440), hCell("\uD0C0\uAC9F", 2400) ] }),
          new TableRow({ children: [
            bCell("1", 600, "FFF5F5"), bCell("MUSE ONCE", 2000, "FFF5F5"), bCell("\uBA85\uD654 \uAC10\uC131 \uD504\uB9AC\uBBF8\uC5C4", 2400, "FFF5F5"), bCell("\uC544\uC6B0\uB810 + \uC18C\uC720\uB9C8\uC2E4", 2440, "FFF5F5"), bCell("20~30\uB300 \uC5EC\uC131, \uC608\uC220 \uAC10\uC131", 2400, "FFF5F5"),
          ] }),
          new TableRow({ children: [
            bCell("2", 600), bCell("LACE MOOD", 2000), bCell("\uBC1C\uB808\uCF54\uC5B4/\uB9AC\uBCF8 \uD328\uC158", 2400), bCell("\uC138\uCEE8\uB4DC\uC720\uB2C8\uD06C\uB124\uC784", 2440), bCell("10~20\uB300 \uC5EC\uC131, K-\uD31D", 2400),
          ] }),
          new TableRow({ children: [
            bCell("3", 600, "FFF5F5"), bCell("TONE DAILY", 2000, "FFF5F5"), bCell("\uBBF8\uB2C8\uBA40 \uBB34\uB4DC \uCEEC\uB7EC", 2400, "FFF5F5"), bCell("\uD558\uC6B0\uC704 + \uD558\uC774\uC6B0", 2440, "FFF5F5"), bCell("20~30\uB300 \uC131\uBCC4\uBB34\uAD00", 2400, "FFF5F5"),
          ] }),
          new TableRow({ children: [
            bCell("4", 600), bCell("PETIT ONCE", 2000), bCell("\uD0A4\uCE58 \uCE90\uB9AD\uD130 \uC800\uAC00 \uC9C4\uC785", 2400), bCell("\uAC00\uB974\uC1A1\uD2F0\uBBF8\uB4DC + \uC5B4\uD504\uC5B4\uD504", 2440), bCell("10~20\uB300, \uCCAB\uAD6C\uB9E4 \uC720\uC785", 2400),
          ] }),
          new TableRow({ children: [
            bCell("5", 600, "FFF5F5"), bCell("BLOOM ONCE", 2000, "FFF5F5"), bCell("\uBE48\uD2F0\uC9C0 \uD50C\uB85C\uB7F4 \uC2DC\uC98C", 2400, "FFF5F5"), bCell("\uC544\uC6B0\uB810 + \uBABD\uBABD\uB4DC", 2440, "FFF5F5"), bCell("20\uB300 \uC5EC\uC131, \uBD04/\uAC00\uC744 \uC2DC\uC98C", 2400, "FFF5F5"),
          ] }),
          new TableRow({ children: [
            bCell("6", 600), bCell("SILVER EDIT", 2000), bCell("\uBA54\uD0C8\uB9AD/\uBBF8\uB7EC \uD2B8\uB80C\uB4DC", 2400), bCell("\uB514\uC790\uC778\uC2A4\uD0A8 + \uC18C\uC720\uB9C8\uC2E4", 2440), bCell("\uC804 \uC5F0\uB839, \uC778\uC2A4\uD0C0 \uC18C\uC7AC\uC6A9", 2400),
          ] }),
          new TableRow({ children: [
            bCell("7", 600, "FFF5F5"), bCell("ONCE SET", 2000, "FFF5F5"), bCell("\uC138\uD2B8 \uAD6C\uC131 \uC804\uBB38", 2400, "FFF5F5"), bCell("\uB514\uC790\uC778\uC2A4\uD0A8 + \uB358\uD0C0\uC6B4", 2440, "FFF5F5"), bCell("\uC120\uBB3C/\uC138\uD2B8 \uAD6C\uB9E4\uCE35", 2400, "FFF5F5"),
          ] }),
        ]
      }),

      new Paragraph({ children: [new PageBreak()] }),

      // ===== 브랜드 1: MUSE ONCE =====
      ...brandBlock("Brand 1", "MUSE ONCE \u2014 \uBA85\uD654 \uAC10\uC131 \uD504\uB9AC\uBBF8\uC5C4",
        "\uB808\uD37C\uB7F0\uC2A4: \uC544\uC6B0\uB810(\uC2DC\uC801 \uC81C\uD488\uBA85) + \uC18C\uC720\uB9C8\uC2E4(\uCEEC\uB7EC\uBA85 \uBE0C\uB79C\uB529) + \uC628\uC138\uC2A4\uD29C\uB514\uC624 \uBA85\uD654 \uC2DC\uB9AC\uC988 \uAC15\uD654", "FFF8F5", [
        ["\uCEE8\uC149", "\uBBF8\uC220\uAD00\uC5D0\uC11C \uAC78\uC5B4\uB098\uC628 \uCF00\uC774\uC2A4. \uBA85\uD654 \uBAA8\uD2F0\uD504\uB97C \uC5D0\uD3ED\uC2DC/\uC824\uD558\uB4DC\uB85C \uD504\uB9AC\uBBF8\uC5C4\uD654. \uAC01 \uC81C\uD488\uC5D0 \uC2DC\uC801 \uC81C\uD488\uBA85\uC744 \uBD99\uC5EC \uAD11\uACE0 \uCE74\uD53C\uB85C \uBC14\uB85C \uD65C\uC6A9"],
        ["\uB514\uC790\uC778 \uBC29\uD5A5", "\uBAA8\uB124 \uC218\uB828/\uD638\uC548\uBBF8\uB85C \uCD94\uC0C1\uD654/\uD074\uB9BC\uD2B8 \uAE08\uBC15 \uB4F1 \uBA85\uD654 \uBAA8\uD2F0\uD504\uB97C \uC5D0\uD3ED\uC2DC \uBC94\uD37C(\uC2E4\uBC84/\uBE14\uB799)\uC640 \uC824\uD558\uB4DC\uB85C \uC81C\uC791. \uC720\uD654 \uD130\uCE58\uAC10\uC744 \uC0B4\uB824\uC11C \uC608\uC220 \uAC10\uC131 \uADF9\uB300\uD654"],
        ["\uCEEC\uB7EC \uD314\uB808\uD2B8", "Monet Blue(\uC218\uB828 \uD30C\uB791), Mir\u00F3 Pink(\uD638\uC548\uBBF8\uB85C \uD551\uD06C), Klimt Gold(\uAE08\uBC15 \uD1A4), Starry Navy(\uBCC4\uC774 \uBE5B\uB098\uB294 \uBC24 \uB124\uC774\uBE44). \uAC01 \uCEEC\uB7EC\uC5D0 \uBA85\uD654 \uC774\uB984\uC744 \uBD99\uC5EC \uBE0C\uB79C\uB529"],
        ["\uC81C\uD488 \uB77C\uC778\uC5C5", "\uC5D0\uD3ED\uC2DC\uBC94\uD37C(\uC2E4\uBC84) 25,000\uC6D0 / \uC5D0\uD3ED\uC2DC\uBC94\uD37C(\uBE14\uB799) 25,000\uC6D0 / \uC824\uD558\uB4DC 20,000\uC6D0 / \uCE74\uB4DC\uD615 28,000\uC6D0 / \uB9E5\uC138\uC774\uD504 \uC824\uD558\uB4DC 25,000\uC6D0"],
        ["\uAD11\uACE0 \uCE74\uD53C", "\"\uBAA8\uB124\uC758 \uD55C\uB09F\uC744 \uC190\uC5D0 \uB2F4\uB2E4\" / \"\uBBF8\uC220\uAD00\uC5D0\uC11C \uAC78\uC5B4\uB098\uC628 \uCF00\uC774\uC2A4\" / \"MUSE ONCE - \uC608\uC220\uC744 \uC77C\uC0C1\uC5D0\" / \"Monet Blue \uC2E0\uC0C1 \uCD9C\uC2DC\""],
        ["\uBA54\uD0C0\uAD11\uACE0 \uC804\uB7B5", "\uB9B4\uC2A4 15\uCD08: \uBA85\uD654 \uC6D0\uC791 \u2192 \uCF00\uC774\uC2A4 \uD074\uB85C\uC988\uC5C5 \uC804\uD658. \uCE90\uB7EC\uC140 3\uC7A5: \uCEEC\uB7EC\uBCC4 \uBC14\uB9AC\uC5D0\uC774\uC158. \uD0C0\uAC9F: 20~30\uB300 \uC5EC\uC131, \uC608\uC220/\uBBF8\uC220\uAD00 \uAD00\uC2EC\uC0AC"],
        ["\uD575\uC2EC \uCC28\uBCC4\uD654", "\uCEEC\uB7EC\uBA85 \uC790\uCCB4\uAC00 \uBA85\uD654 \uC774\uB984 + \uC81C\uD488\uBA85\uC774 \uC2DC\uC801 \uCE90\uCE58\uD504\uB808\uC774\uC988 = \uAD11\uACE0 \uCE74\uD53C\uAC00 \uC790\uB3D9\uC73C\uB85C \uC644\uC131\uB418\uB294 \uAD6C\uC870"],
      ]),

      new Paragraph({ children: [new PageBreak()] }),

      // ===== 브랜드 2: LACE MOOD =====
      ...brandBlock("Brand 2", "LACE MOOD \u2014 \uBC1C\uB808\uCF54\uC5B4 / \uB9AC\uBCF8 \uD328\uC158",
        "\uB808\uD37C\uB7F0\uC2A4: \uC138\uCEE8\uB4DC\uC720\uB2C8\uD06C\uB124\uC784(\uD328\uC158 \uC545\uC138\uC11C\uB9AC \uD3EC\uC9C0\uC154\uB2DD) + \uC628\uC138\uC2A4\uD29C\uB514\uC624 \uB9AC\uBCF8 \uD328\uD134 \uAC15\uD654", "F5F0FF", [
        ["\uCEE8\uC149", "\uD3F0\uCF00\uC774\uC2A4\uB97C \uD328\uC158 \uC545\uC138\uC11C\uB9AC\uB85C. \uBC1C\uB808\uCF54\uC5B4/\uB9AC\uBCF8/\uB808\uC774\uC2A4 \uD2B8\uB80C\uB4DC\uB97C \uCF00\uC774\uC2A4\uC5D0 \uC811\uBAA9. wonyoungism, \uCF54\uC5B4 \uAC10\uC131\uC744 \uD0C0\uAC9F\uD55C \uD398\uBBF8\uB2CC \uB77C\uC778"],
        ["\uB514\uC790\uC778 \uBC29\uD5A5", "\uC587\uC740 \uB9AC\uBCF8 \uD328\uD134, \uB808\uC774\uC2A4 \uD14D\uC2A4\uCC98, \uBBFC\uD2B8/\uD654\uC774\uD2B8/\uBCA0\uC774\uBE44\uD551\uD06C \uD30C\uC2A4\uD154 \uBC30\uACBD. \uD22C\uBA85+\uD551\uD06C\uD504\uB808\uC784, \uC5C7\uC740 \uB9AC\uBCF8 \uD328\uD134 \uD504\uB9B0\uD2B8. Z\uD50C\uB9BD\uC6A9 \uB9AC\uBCF8 \uCF00\uC774\uC2A4\uB3C4 \uAD6C\uC131"],
        ["\uCEEC\uB7EC \uD314\uB808\uD2B8", "Ballet Pink, Lace White, Mint Ribbon, Cream Blush. \uD30C\uC2A4\uD154 \uD1A4 \uC911\uC2EC\uC73C\uB85C \uBD80\uB4DC\uB7EC\uC6B4 \uC5EC\uC131\uC2A4\uB7EC\uC6C0. \uC0AC\uC9C4 \uCC0D\uC5B4\uB3C4 \uC608\uC05C \uCEEC\uB7EC"],
        ["\uC81C\uD488 \uB77C\uC778\uC5C5", "\uC824\uD558\uB4DC(\uB9AC\uBCF8\uD328\uD134) 18,000\uC6D0 / \uD22C\uBA85\uD504\uB808\uC784(\uBBFC\uD2B8/\uD551\uD06C) 20,000\uC6D0 / \uC5D0\uD3ED\uC2DC\uBC94\uD37C(\uB808\uC774\uC2A4) 25,000\uC6D0 / Z\uD50C\uB9BD \uB9AC\uBCF8 22,000\uC6D0"],
        ["\uAD11\uACE0 \uCE74\uD53C", "\"\uC587\uC740 \uB9AC\uBCF8 \uD328\uD134 \uBC1C\uB808\uCF54\uC5B4 \uCF00\uC774\uC2A4\" / \"\uC624\uB298\uC758 OOTD\uC5D0 \uC5B4\uC6B8\uB9AC\uB294 \uCF00\uC774\uC2A4\" / \"Ballet Pink \uC2E0\uC0C1\" / \"\uC6D0\uC601\uC774\uC998 \uD551\uD06C \uB9AC\uBCF8\""],
        ["\uBA54\uD0C0\uAD11\uACE0 \uC804\uB7B5", "\uB9B4\uC2A4: OOTD \uC2A4\uD0C0\uC77C\uB9C1\uC5D0 \uCF00\uC774\uC2A4 \uB9E4\uCE6D. \uCE90\uB7EC\uC140: \uCEEC\uB7EC\uBCC4 \uCF54\uB514 \uB9E4\uCE6D. \uD0C0\uAC9F: 10~20\uB300 \uC5EC\uC131, K-\uD31D/\uBC1C\uB808\uCF54\uC5B4 \uAD00\uC2EC\uC0AC"],
        ["\uD575\uC2EC \uCC28\uBCC4\uD654", "\uBC1C\uB808\uCF54\uC5B4/\uB9AC\uBCF8\uC774\uB77C\uB294 \uBA85\uD655\uD55C \uD0C0\uAC9F \uD0A4\uC6CC\uB4DC + OOTD \uCF54\uB514 \uB9E4\uCE6D \uAD11\uACE0\uB85C \uD328\uC158 \uC720\uC800 \uACF5\uB7B5"],
      ]),

      new Paragraph({ children: [new PageBreak()] }),

      // ===== 브랜드 3: TONE DAILY =====
      ...brandBlock("Brand 3", "TONE DAILY \u2014 \uBBF8\uB2C8\uBA40 \uBB34\uB4DC \uCEEC\uB7EC",
        "\uB808\uD37C\uB7F0\uC2A4: \uD558\uC6B0\uC704(\uBBF8\uB2C8\uBA40 \uBB34\uB4DC) + \uD558\uC774\uC6B0(\uD0C0\uC784\uB808\uC2A4 \uC808\uC81C\uBBF8) + \uB274\uD2B8\uB7F4 \uD1A4 \uD2B8\uB80C\uB4DC", "F5FAFF", [
        ["\uCEE8\uC149", "\uACFC\uD558\uC9C0 \uC54A\uC740 \uCEEC\uB7EC \uD558\uB098\uB85C \uC77C\uC0C1\uC744 \uC644\uC131\uD558\uB294 \uBB34\uB4DC \uCF00\uC774\uC2A4. \uBB34\uC9C0 \uCEEC\uB7EC + \uBBF8\uC138\uD55C \uD14D\uC2A4\uCC98 \uCC28\uC774\uB85C \uAC10\uC131 \uC804\uB2EC. \uC131\uBCC4/\uB098\uC774 \uBB34\uAD00\uD55C \uBCF4\uD3B8\uC801 \uB514\uC790\uC778"],
        ["\uB514\uC790\uC778 \uBC29\uD5A5", "\uBB34\uC9C0 \uB2E8\uC0C9 + \uBBF8\uC138\uD55C \uD14D\uC2A4\uCC98(\uB9E4\uD2B8/\uBC18\uAD11/\uC0E4\uBE0C). \uB808\uB354 \uD14D\uC2A4\uCC98, \uD328\uBE0C\uB9AD \uD130\uCE58, \uC2A4\uD1A4 \uC9C8\uAC10 \uB4F1 \uC18C\uC7AC\uAC10\uC73C\uB85C \uCC28\uBCC4\uD654. \uBB38\uC790/\uD328\uD134 \uC5C6\uC774 \uCEEC\uB7EC\uB9CC\uC73C\uB85C \uC2B9\uBD80"],
        ["\uCEEC\uB7EC \uD314\uB808\uD2B8", "Sand, Clay, Sage, Stone Grey, Cream Beige, Deep Forest, Dusty Rose. \uAC01 \uCEEC\uB7EC\uC5D0 \uAC10\uC131\uBA85: \"Morning Sand\", \"Calm Sage\", \"Deep Stone\" \uB4F1"],
        ["\uC81C\uD488 \uB77C\uC778\uC5C5", "\uB9E4\uD2B8 \uD558\uB4DC 16,000\uC6D0 / \uBC18\uAD11 \uC824\uD558\uB4DC 18,000\uC6D0 / \uB808\uB354\uD14D\uC2A4\uCC98 \uBC94\uD37C 22,000\uC6D0 / \uB9E5\uC138\uC774\uD504 \uBC94\uD37C 24,000\uC6D0"],
        ["\uAD11\uACE0 \uCE74\uD53C", "\"\uC624\uB298\uC758 \uD1A4, Morning Sand\" / \"\uC5B4\uB5A4 \uC21C\uAC04\uC5D0\uB3C4 \uC790\uC5F0\uC2A4\uB7EC\uC6B4\" / \"\uC9C8\uB9AC\uC9C0 \uC54A\uB294 \uCEEC\uB7EC\" / \"TONE DAILY - \uB0A0\uB9C8\uB2E4 \uB4E4\uC5B4\uB3C4 \uC88B\uC740\""],
        ["\uBA54\uD0C0\uAD11\uACE0 \uC804\uB7B5", "\uB9B4\uC2A4: \uCEEC\uB7EC \uC2A4\uC640\uCE58 + \uC77C\uC0C1 \uC2A4\uD0C0\uC77C\uB9C1 \uBBF8\uB2C8\uBA40 \uBB34\uB4DC. \uCE90\uB7EC\uC140: 7\uAC00\uC9C0 \uCEEC\uB7EC \uBC14\uB9AC\uC5D0\uC774\uC158. \uD0C0\uAC9F: 20~30\uB300 \uC131\uBCC4\uBB34\uAD00, \uBBF8\uB2C8\uBA40 \uC120\uD638"],
        ["\uD575\uC2EC \uCC28\uBCC4\uD654", "\uBB34\uC9C0 \uCEEC\uB7EC\uB9CC\uC73C\uB85C \uC2B9\uBD80 + \uD14D\uC2A4\uCC98 \uCC28\uC774\uB85C \uAC00\uACA9\uB300 \uAD6C\uBD84 + \uC131\uBCC4\uBB34\uAD00 \uBCF4\uD3B8\uC131\uC73C\uB85C \uD0C0\uAC9F \uD655\uC7A5"],
      ]),

      new Paragraph({ children: [new PageBreak()] }),

      // ===== 브랜드 4: PETIT ONCE =====
      ...brandBlock("Brand 4", "PETIT ONCE \u2014 \uD0A4\uCE58 \uCE90\uB9AD\uD130 \uC800\uAC00 \uC9C4\uC785",
        "\uB808\uD37C\uB7F0\uC2A4: \uAC00\uB974\uC1A1\uD2F0\uBBF8\uB4DC(\uB79C\uB364\uBC15\uC2A4 \uBBF8\uB07C\uC0C1\uD488 + 2\uB2E8\uACC4 \uAC00\uACA9) + \uC5B4\uD504\uC5B4\uD504(\uCE90\uB9AD\uD130 \uC2DC\uB9AC\uC988)", "FFFBF0", [
        ["\uCEE8\uC149", "\uBA54\uD0C0\uAD11\uACE0 \uC2E0\uADDC\uACE0\uAC1D \uC720\uC785 \uC804\uC6A9 \uB77C\uC778. \uC790\uCCB4 \uCE90\uB9AD\uD130/\uD0A4\uCE58 \uB514\uC790\uC778\uC744 \uC800\uAC00\uC5D0 \uC81C\uACF5\uD574\uC11C \uCCAB \uAD6C\uB9E4\uB97C \uC720\uB3C4\uD558\uACE0, \uC5C5\uC140\uB85C \uC5F0\uACB0"],
        ["\uB514\uC790\uC778 \uBC29\uD5A5", "\uAC04\uB2E8\uD55C \uB450\uB4E4/\uC77C\uB7EC\uC2A4\uD2B8 \uCE90\uB9AD\uD130. \uBE44\uBE44\uB4DC \uCEEC\uB7EC + \uD22C\uBA85 \uBC30\uACBD. \uD3FC\uC774 \uACFC\uD558\uC9C0 \uC54A\uC9C0\uB9CC \uADC0\uC5EC\uC6B4 \uAC10\uC131. \uB79C\uB364\uBC15\uC2A4 \uAD6C\uC131\uC73C\uB85C \"\uC5B4\uB5A4 \uB514\uC790\uC778\uC774 \uC62C\uC9C0 \uBAB0\uB77C\uC694\" \uCEE8\uC149"],
        ["\uCEEC\uB7EC \uD314\uB808\uD2B8", "Bubblegum Pink, Lime, Sky Blue, Peach, Lavender. \uBE44\uBE44\uB4DC\uD558\uC9C0\uB9CC \uAE68\uB057\uD55C \uD30C\uC2A4\uD154 \uD1A4"],
        ["\uC81C\uD488 \uB77C\uC778\uC5C5", "\uD22C\uBA85\uD558\uB4DC 12,000\uC6D0 / \uD22C\uBA85\uC824\uD558\uB4DC 14,000\uC6D0 / \uB79C\uB364\uBC15\uC2A4 9,999\uC6D0 / \uC5C5\uC140\uC6A9 \uC5D0\uD3ED\uC2DC 22,000\uC6D0"],
        ["\uAD11\uACE0 \uCE74\uD53C", "\"\uB79C\uB364\uBC15\uC2A4 9,999\uC6D0 - \uC5B4\uB5A4 \uB514\uC790\uC778\uC774 \uC62C\uC9C0 \uBAB0\uB77C\uC694\" / \"\uCCAB \uAD6C\uB9E4 \uD2B9\uAC00\" / \"PETIT ONCE \uD22C\uBA85\uCF00\uC774\uC2A4 12,000\uC6D0\uBD80\uD130\" / \"\uADC0\uC5EC\uC6B4 \uAC83\uC774 \uC138\uC0C1\uC744 \uAD6C\uD55C\uB2E4\""],
        ["\uBA54\uD0C0\uAD11\uACE0 \uC804\uB7B5", "\uC804\uD658 \uCE90\uD398\uC778 \uC804\uC6A9. \uB79C\uB364\uBC15\uC2A4 9,999\uC6D0\uC73C\uB85C CPA \uCD5C\uC801\uD654. \uAD6C\uB9E4 \uD6C4 \uB9AC\uD0C0\uAC9F\uC73C\uB85C MUSE ONCE / LACE MOOD \uC5C5\uC140 \uC720\uB3C4. \uD0C0\uAC9F: 10~20\uB300, \uCCAB\uAD6C\uB9E4 \uC720\uC785"],
        ["\uD575\uC2EC \uCC28\uBCC4\uD654", "\uC800\uAC00 \uBBF8\uB07C\uC0C1\uD488\uC73C\uB85C \uC2E0\uADDC \uACE0\uAC1D \uD655\uBCF4 \u2192 \uB2E4\uB978 \uBE0C\uB79C\uB4DC \uB77C\uC778\uC73C\uB85C \uC5C5\uC140. \uD37C\uB110 \uAD6C\uC870\uC758 \uC9C4\uC785\uC810"],
      ]),

      // ===== 브랜드 5: BLOOM ONCE =====
      ...brandBlock("Brand 5", "BLOOM ONCE \u2014 \uBE48\uD2F0\uC9C0 \uD50C\uB85C\uB7F4 \uC2DC\uC98C",
        "\uB808\uD37C\uB7F0\uC2A4: \uC544\uC6B0\uB810(\uC77C\uB7EC\uC2A4\uD2B8 \uC544\uD2B8\uC6CC\uD06C) + \uBABD\uBABD\uB4DC(\uBBF4\uB4DC \uAC10\uC131 \uBC00\uB3C4) + \uC2DC\uC98C\uAC10 \uCEEC\uB809\uC158", "F0FFF5", [
        ["\uCEE8\uC149", "\uBD04/\uAC00\uC744 \uC2DC\uC98C\uB9C8\uB2E4 \uC2E0\uADDC \uD50C\uB85C\uB7F4 \uCEEC\uB809\uC158 \uCD9C\uC2DC. \uC218\uCC44\uD654 \uD130\uCE58 + \uBE48\uD2F0\uC9C0 \uAF43\uBB34\uB2EC\uB85C \uC190\uADF8\uB9BC \uAC10\uC131. AI \uD53C\uB85C\uAC10\uC5D0 \uC9C0\uCE5C \uC18C\uBE44\uC790\uB4E4\uC5D0\uAC8C \uC5B4\uD544"],
        ["\uB514\uC790\uC778 \uBC29\uD5A5", "\uC218\uCC44\uD654 \uD130\uCE58\uC758 \uC740\uC740\uD55C \uAF43\uBB34\uB2EC. \uBD04: \uBCC7\uAF43/\uD280\uB9BD/\uB370\uC774\uC9C0, \uAC00\uC744: \uCF54\uC2A4\uBAA8\uC2A4/\uB2E8\uD48D\uB098\uBB34/\uB77C\uBCA4\uB354. \uAE08\uC120 \uC544\uC6C3\uB77C\uC778\uC73C\uB85C \uACE0\uAE09\uAC10 \uCD94\uAC00. \uC824\uD558\uB4DC + \uC5D0\uD3ED\uC2DC \uBC94\uD37C"],
        ["\uCEEC\uB7EC \uD314\uB808\uD2B8", "\uBD04: Cherry Blossom, Tulip Cream, Daisy Mint / \uAC00\uC744: Autumn Rose, Cosmos Purple, Maple Gold. \uC2DC\uC98C\uBCC4 \uCEEC\uB7EC \uD314\uB808\uD2B8 \uAD50\uCCB4"],
        ["\uC81C\uD488 \uB77C\uC778\uC5C5", "\uC824\uD558\uB4DC 18,000\uC6D0 / \uC5D0\uD3ED\uC2DC\uBC94\uD37C 24,000\uC6D0 / \uD50C\uB85C\uB7F4 3\uC885 \uC138\uD2B8 48,000\uC6D0 / \uD55C\uC815\uD310 \uC2DC\uC98C \uCEEC\uB809\uC158 \uBC15\uC2A4 25,000\uC6D0"],
        ["\uAD11\uACE0 \uCE74\uD53C", "\"\uBD04\uC5D0 \uC5B4\uC6B8\uB9AC\uB294 Cherry Blossom \uCEEC\uB809\uC158\" / \"\uC218\uCC44\uD654\uCC98\uB7FC \uD53C\uC5B4\uB098\uB294 \uAF43\uBB34\uB2EC \uCF00\uC774\uC2A4\" / \"\uC2DC\uC98C \uD55C\uC815 \uD50C\uB85C\uB7F4 \uBC15\uC2A4 25,000\uC6D0\""],
        ["\uBA54\uD0C0\uAD11\uACE0 \uC804\uB7B5", "\uC2DC\uC98C \uC804\uD658\uAE30(3\uC6D4/9\uC6D4)\uC5D0 \uC9D1\uC911 \uC9D1\uD589. \uC2DC\uC98C \uD55C\uC815 \uAE34\uBC15\uAC10\uC73C\uB85C \uC804\uD658 \uADF9\uB300\uD654. \uD0C0\uAC9F: 20\uB300 \uC5EC\uC131, \uAF43/\uC790\uC5F0 \uAC10\uC131 \uC120\uD638"],
        ["\uD575\uC2EC \uCC28\uBCC4\uD654", "\uC2DC\uC98C\uBCC4 \uCEEC\uB809\uC158 \uAD50\uCCB4 \u2192 \uC7AC\uAD6C\uB9E4 \uC720\uB3C4 + \uC2DC\uC98C \uD55C\uC815 \uAE34\uBC15\uAC10 + \uC190\uADF8\uB9BC \uAC10\uC131\uC73C\uB85C AI\uD53BC\uB85C\uAC10 \uB300\uC751"],
      ]),

      new Paragraph({ children: [new PageBreak()] }),

      // ===== 브랜드 6: SILVER EDIT =====
      ...brandBlock("Brand 6", "SILVER EDIT \u2014 \uBA54\uD0C8\uB9AD / \uBBF8\uB7EC \uD2B8\uB80C\uB4DC",
        "\uB808\uD37C\uB7F0\uC2A4: \uB514\uC790\uC778\uC2A4\uD0A8(WHAT IS YOUR COLOR?) + \uC18C\uC720\uB9C8\uC2E4(\uC5D0\uD3ED\uC2DC \uD14D\uC2A4\uCC98)", "F8F8FF", [
        ["\uCEE8\uC149", "\uC140\uD53C \uBBF8\uB7EC \uD6A8\uACFC + \uBA54\uD0C8\uB9AD \uAC10\uC131. \uC778\uC2A4\uD0C0 \uC0AC\uC9C4 \uCC0D\uC744 \uB54C \uBE5B \uBC18\uC0AC \uD6A8\uACFC\uB85C \uC8FC\uBAA9\uB3C4 \uADF9\uB300\uD654. \uD06C\uB86C/\uC2E4\uBC84/\uD640\uB85C\uADF8\uB7A8 \uC18C\uC7AC\uB85C \uD504\uB9AC\uBBF8\uC5C4"],
        ["\uB514\uC790\uC778 \uBC29\uD5A5", "\uC2E4\uBC84 \uBBF8\uB7EC, \uD06C\uB86C \uBC18\uC0AC, \uD640\uB85C\uADF8\uB7A8 \uC18C\uC7AC. \uBA85\uD654 \uBAA8\uD2F0\uD504 + \uC2E4\uBC84 \uBC30\uACBD\uC73C\uB85C MUSE ONCE\uC640 \uD06C\uB85C\uC2A4. \uC5D0\uD3ED\uC2DC \uC7A5\uC2DD \uCD94\uAC00\uB85C \uCC28\uBCC4\uD654"],
        ["\uCEEC\uB7EC \uD314\uB808\uD2B8", "Mirror Silver, Chrome Gold, Hologram, Ice Blue Mirror. \uBA54\uD0C8\uB9AD \uC18C\uC7AC \uC790\uCCB4\uAC00 \uCEEC\uB7EC"],
        ["\uC81C\uD488 \uB77C\uC778\uC5C5", "\uBBF8\uB7EC \uD558\uB4DC 20,000\uC6D0 / \uBBF8\uB7EC \uC5D0\uD3ED\uC2DC 26,000\uC6D0 / \uD640\uB85C\uADF8\uB7A8 \uC5D0\uD3ED\uC2DC 28,000\uC6D0 / \uD06C\uB86C \uBC94\uD37C 24,000\uC6D0"],
        ["\uAD11\uACE0 \uCE74\uD53C", "\"\uBBF8\uB7EC\uCF00\uC774\uC2A4 \uC2E0\uC0C1 - \uBE5B\uC774 \uB2FF\uB294 \uC21C\uAC04\" / \"SILVER EDIT - \uC624\uB298\uC758 \uBE5B\" / \"\uC140\uD53C \uCC0D\uC744 \uB54C \uAC00\uC7A5 \uC608\uC05C \uCF00\uC774\uC2A4\" / \"Chrome Gold \uD55C\uC815\""],
        ["\uBA54\uD0C0\uAD11\uACE0 \uC804\uB7B5", "\uB9B4\uC2A4: \uBE5B \uBC18\uC0AC \uD6A8\uACFC \uBCF4\uC5EC\uC8FC\uB294 15\uCD08 \uC601\uC0C1(\uD074\uB9AD\uC728 \uADF9\uB300\uD654). \uC2A4\uD1A0\uB9AC: \uC140\uD53C \uBBF8\uB7EC \uD6A8\uACFC \uAC15\uC870. \uD0C0\uAC9F: \uC804 \uC5F0\uB839, \uC778\uC2A4\uD0C0 \uC18C\uC7AC\uC6A9"],
        ["\uD575\uC2EC \uCC28\uBCC4\uD654", "\uBE5B \uBC18\uC0AC/\uBBF8\uB7EC \uD6A8\uACFC\uB294 \uB9B4\uC2A4 \uAD11\uACE0\uC5D0\uC11C \uC2DC\uC120 \uC7A1\uAE30 \uCD5C\uACE0 + \uC778\uC2A4\uD0C0 \uC0AC\uC9C4\uC6A9 \uC218\uC694 \uADF9\uAC15"],
      ]),

      // ===== 브랜드 7: ONCE SET =====
      ...brandBlock("Brand 7", "ONCE SET \u2014 \uC138\uD2B8 \uAD6C\uC131 \uC804\uBB38",
        "\uB808\uD37C\uB7F0\uC2A4: \uB514\uC790\uC778\uC2A4\uD0A8(\uC138\uD2B8 \uC804\uB7B5) + \uB358\uD0C0\uC6B4(\uD06C\uB85C\uC2A4\uC140\uB9C1) + \uC804 \uBE0C\uB79C\uB4DC \uD1B5\uD569 \uC138\uD2B8", "FFF5F5", [
        ["\uCEE8\uC149", "\uC628\uC138\uC2A4\uD29C\uB514\uC624 \uC804 \uBE0C\uB79C\uB4DC\uC758 \uBCA0\uC2A4\uD2B8 \uC81C\uD488\uC744 \uC138\uD2B8\uB85C \uBB36\uC5B4 \uAC1D\uB2E8\uAC00 \uC0C1\uC2B9. \uCF00\uC774\uC2A4+\uCE74\uB4DC\uD640\uB354+\uC2A4\uB9C8\uD2B8\uD1A1+\uD0A4\uB9C1 \uC870\uD569. \uC120\uBB3C\uC6A9 \uC218\uC694 \uACF5\uB7B5"],
        ["\uB514\uC790\uC778 \uBC29\uD5A5", "\uAC01 \uBE0C\uB79C\uB4DC \uBCA0\uC2A4\uD2B8\uC140\uB7EC\uC758 \uC138\uD2B8 \uAD6C\uC131. MUSE ONCE \uBA85\uD654\uCF00\uC774\uC2A4+\uCE74\uB4DC\uD640\uB354, BLOOM ONCE \uD50C\uB85C\uB7F4 3\uC885 SET, TONE DAILY \uCEEC\uB7EC 3\uC885 SET \uB4F1. \uD328\uD0A4\uC9C0\uB3C4 \uAC10\uC131\uC801\uC73C\uB85C \uAD6C\uC131"],
        ["\uCEEC\uB7EC \uD314\uB808\uD2B8", "\uAC01 \uBE0C\uB79C\uB4DC\uC758 \uB300\uD45C \uCEEC\uB7EC\uB97C \uC138\uD2B8\uB85C \uBB36\uC74C. \uD1B5\uC77C\uAC10 \uC788\uB294 \uC138\uD2B8 \uD328\uD0A4\uC9C0 \uCEEC\uB7EC"],
        ["\uC81C\uD488 \uB77C\uC778\uC5C5", "\uCF00\uC774\uC2A4+\uCE74\uB4DC\uD640\uB354 35,000\uC6D0 / \uCF00\uC774\uC2A4+\uC2A4\uB9C8\uD2B8\uD1A1+\uD0A4\uB9C1 38,000\uC6D0 / \uD50C\uB85C\uB7F4 3\uC885 SET 48,000\uC6D0 / \uCF5C\uB809\uC158 5\uC885 SET 78,000\uC6D0 / \uC120\uBB3C\uC6A9 \uD328\uD0A4\uC9C0 +3,000\uC6D0"],
        ["\uAD11\uACE0 \uCE74\uD53C", "\"\uC27C\uC544\uC9C0\uB294 \uC608\uC220 \uC138\uD2B8\" / \"\uC120\uBB3C\uD558\uAE30 \uC88B\uC740 \uAC10\uC131 \uC138\uD2B8\" / \"ONCE SET - \uCF00\uC774\uC2A4+\uCE74\uB4DC\uD640\uB354 35,000\uC6D0\" / \"\uBCA0\uC2A4\uD2B8 3\uC885 \uBB36\uC74C \uD2B9\uAC00\""],
        ["\uBA54\uD0C0\uAD11\uACE0 \uC804\uB7B5", "\uCE90\uB7EC\uC140 \uAD11\uACE0: \uC138\uD2B8 \uAD6C\uC131 \uBCF4\uC5EC\uC8FC\uAE30. \uB9B4\uC2A4: \uD328\uD0A4\uC9C0 \uAC1C\uBD09 \uAC10\uC131. \uD0C0\uAC9F: \uC120\uBB3C\uC6A9 \uAD6C\uB9E4\uCE35, \uC138\uD2B8 \uAD6C\uB9E4\uB85C \uAC1D\uB2E8\uAC00 \uC0C1\uC2B9"],
        ["\uD575\uC2EC \uCC28\uBCC4\uD654", "\uC804 \uBE0C\uB79C\uB4DC \uD1B5\uD569 \uC138\uD2B8\uB85C \uAC1D\uB2E8\uAC00 \uC0C1\uC2B9 + \uC120\uBB3C\uC6A9 \uD328\uD0A4\uC9C0\uB85C \uC0C8\uB85C\uC6B4 \uC218\uC694 \uCC3D\uCD9C + \uD06C\uB85C\uC2A4\uC140\uB9C1 \uAD6C\uC870"],
      ]),

      new Paragraph({ children: [new PageBreak()] }),

      // ===== 전체 포트폴리오 요약 =====
      h1("\uC804\uCCB4 \uD3EC\uD2B8\uD3F4\uB9AC\uC624 \uC694\uC57D: 10\uAC1C \uBE0C\uB79C\uB4DC"),
      p("\uAE30\uC874 3\uAC1C + \uC2E0\uADDC 7\uAC1C = \uCD1D 10\uAC1C \uBE0C\uB79C\uB4DC \uD3EC\uD2B8\uD3F4\uB9AC\uC624 \uC644\uC131:"),

      new Table({
        width: { size: 9840, type: WidthType.DXA },
        columnWidths: [600, 1800, 2200, 1700, 1840, 1700],
        rows: [
          new TableRow({ children: [ hCell("#", 600), hCell("\uBE0C\uB79C\uB4DC", 1800), hCell("\uCEE8\uC149", 2200), hCell("\uAC00\uACA9\uB300", 1700), hCell("\uC5ED\uD560", 1840), hCell("\uC0C1\uD0DC", 1700) ] }),
          new TableRow({ children: [ bCell("1", 600, "E8F5E9"), bCell("\uAE30\uC874 \uBE0C\uB79C\uB4DC A", 1800, "E8F5E9"), bCell("(\uAE30\uC874 \uC6B4\uC601 \uC911)", 2200, "E8F5E9"), bCell("-", 1700, "E8F5E9"), bCell("\uBA54\uC778", 1840, "E8F5E9"), bCell("\uC644\uB8CC", 1700, "E8F5E9") ] }),
          new TableRow({ children: [ bCell("2", 600, "E8F5E9"), bCell("\uAE30\uC874 \uBE0C\uB79C\uB4DC B", 1800, "E8F5E9"), bCell("(\uAE30\uC874 \uC6B4\uC601 \uC911)", 2200, "E8F5E9"), bCell("-", 1700, "E8F5E9"), bCell("\uBA54\uC778", 1840, "E8F5E9"), bCell("\uC644\uB8CC", 1700, "E8F5E9") ] }),
          new TableRow({ children: [ bCell("3", 600, "E8F5E9"), bCell("\uAE30\uC874 \uBE0C\uB79C\uB4DC C", 1800, "E8F5E9"), bCell("(\uAE30\uC874 \uC6B4\uC601 \uC911)", 2200, "E8F5E9"), bCell("-", 1700, "E8F5E9"), bCell("\uBA54\uC778", 1840, "E8F5E9"), bCell("\uC644\uB8CC", 1700, "E8F5E9") ] }),
          new TableRow({ children: [ bCell("4", 600, "FFF8F0"), bCell("MUSE ONCE", 1800, "FFF8F0"), bCell("\uBA85\uD654 \uAC10\uC131 \uD504\uB9AC\uBBF8\uC5C4", 2200, "FFF8F0"), bCell("20,000~28,000", 1700, "FFF8F0"), bCell("\uD504\uB9AC\uBBF8\uC5C4", 1840, "FFF8F0"), bCell("NEW", 1700, "FFF8F0") ] }),
          new TableRow({ children: [ bCell("5", 600, "FFF8F0"), bCell("LACE MOOD", 1800, "FFF8F0"), bCell("\uBC1C\uB808\uCF54\uC5B4/\uB9AC\uBCF8", 2200, "FFF8F0"), bCell("18,000~25,000", 1700, "FFF8F0"), bCell("K-\uD31D \uD0C0\uAC9F", 1840, "FFF8F0"), bCell("NEW", 1700, "FFF8F0") ] }),
          new TableRow({ children: [ bCell("6", 600, "FFF8F0"), bCell("TONE DAILY", 1800, "FFF8F0"), bCell("\uBBF8\uB2C8\uBA40 \uBB34\uB4DC\uCEEC\uB7EC", 2200, "FFF8F0"), bCell("16,000~24,000", 1700, "FFF8F0"), bCell("\uBCF4\uD3B8\uC131 \uD655\uC7A5", 1840, "FFF8F0"), bCell("NEW", 1700, "FFF8F0") ] }),
          new TableRow({ children: [ bCell("7", 600, "FFF8F0"), bCell("PETIT ONCE", 1800, "FFF8F0"), bCell("\uD0A4\uCE58/\uC800\uAC00 \uC9C4\uC785", 2200, "FFF8F0"), bCell("9,999~22,000", 1700, "FFF8F0"), bCell("CPA \uCD5C\uC801\uD654", 1840, "FFF8F0"), bCell("NEW", 1700, "FFF8F0") ] }),
          new TableRow({ children: [ bCell("8", 600, "FFF8F0"), bCell("BLOOM ONCE", 1800, "FFF8F0"), bCell("\uBE48\uD2F0\uC9C0 \uD50C\uB85C\uB7F4", 2200, "FFF8F0"), bCell("18,000~48,000", 1700, "FFF8F0"), bCell("\uC2DC\uC98C \uC7AC\uAD6C\uB9E4", 1840, "FFF8F0"), bCell("NEW", 1700, "FFF8F0") ] }),
          new TableRow({ children: [ bCell("9", 600, "FFF8F0"), bCell("SILVER EDIT", 1800, "FFF8F0"), bCell("\uBA54\uD0C8\uB9AD/\uBBF8\uB7EC", 2200, "FFF8F0"), bCell("20,000~28,000", 1700, "FFF8F0"), bCell("\uC778\uC2A4\uD0C0 \uC18C\uC7AC", 1840, "FFF8F0"), bCell("NEW", 1700, "FFF8F0") ] }),
          new TableRow({ children: [ bCell("10", 600, "FFF8F0"), bCell("ONCE SET", 1800, "FFF8F0"), bCell("\uC138\uD2B8 \uAD6C\uC131 \uC804\uBB38", 2200, "FFF8F0"), bCell("35,000~78,000", 1700, "FFF8F0"), bCell("\uAC1D\uB2E8\uAC00 \uC0C1\uC2B9", 1840, "FFF8F0"), bCell("NEW", 1700, "FFF8F0") ] }),
        ]
      }),

      new Paragraph({ spacing: { before: 300 }, children: [] }),
      h2("\uD37C\uB110 \uAD6C\uC870"),
      p("PETIT ONCE(9,999\uC6D0 \uB79C\uB364\uBC15\uC2A4) \u2192 \uCCAB\uAD6C\uB9E4 \u2192 \uB9AC\uD0C0\uAC9F\uC73C\uB85C MUSE ONCE / LACE MOOD / BLOOM ONCE \uC5C5\uC140 \u2192 ONCE SET\uB85C \uAC1D\uB2E8\uAC00 \uC0C1\uC2B9"),
      p("TONE DAILY\uB294 \uC131\uBCC4\uBB34\uAD00 \uBCF4\uD3B8\uC131\uC73C\uB85C \uD0C0\uAC9F \uD655\uC7A5, SILVER EDIT\uB294 \uC778\uC2A4\uD0C0 \uC18C\uC7AC\uC6A9\uC73C\uB85C SNS \uD655\uC0B0 \uC5ED\uD560"),

      new Paragraph({ spacing: { before: 400 }, border: { top: { style: BorderStyle.SINGLE, size: 3, color: "CCCCCC", space: 1 } }, children: [] }),
      new Paragraph({ spacing: { before: 100 }, children: [new TextRun({ text: "\uAC10\uC131 \uC2A4\uBAB0\uBE0C\uB79C\uB4DC 10\uAC1C \uB808\uD37C\uB7F0\uC2A4 \uAC00\uC774\uB4DC \uAE30\uBC18 \uC124\uACC4 | \uC628\uC138\uC2A4\uD29C\uB514\uC624", size: 18, color: "999999", italics: true })] }),
      new Paragraph({ children: [new TextRun({ text: "\uC791\uC131\uC77C: 2026\uB144 4\uC6D4 14\uC77C", size: 18, color: "999999", italics: true })] }),
    ]
  }]
});

Packer.toBuffer(doc).then(buffer => {
  fs.writeFileSync("/sessions/zealous-confident-pasteur/mnt/outputs/온세스튜디오_신규7개브랜드_기획서.docx", buffer);
  console.log("Brand plan created successfully!");
});
