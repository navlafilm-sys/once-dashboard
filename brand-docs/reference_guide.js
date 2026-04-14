const fs = require('fs');
const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
        Header, Footer, AlignmentType, LevelFormat, ExternalHyperlink,
        HeadingLevel, BorderStyle, WidthType, ShadingType,
        PageNumber, PageBreak } = require('docx');

const border = { style: BorderStyle.SINGLE, size: 1, color: "DDDDDD" };
const borders = { top: border, bottom: border, left: border, right: border };
const cm = { top: 80, bottom: 80, left: 120, right: 120 };
const noBorderBottom = { top: border, bottom: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" }, left: border, right: border };
const noBorderTop = { top: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" }, bottom: border, left: border, right: border };

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
function bCellMulti(runs, width, fill) {
  return new TableCell({
    borders, width: { size: width, type: WidthType.DXA },
    shading: fill ? { fill, type: ShadingType.CLEAR } : undefined,
    margins: cm,
    children: [new Paragraph({ spacing: { after: 40 }, children: runs })]
  });
}

function heading1(text) {
  return new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun(text)] });
}
function heading2(text) {
  return new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun(text)] });
}
function body(text, opts = {}) {
  return new Paragraph({ spacing: { after: opts.after || 120 }, children: [new TextRun({ text, size: 22, font: "Arial", ...opts })] });
}
function boldBody(label, text) {
  return new Paragraph({ spacing: { after: 100 }, children: [
    new TextRun({ text: label, size: 22, font: "Arial", bold: true }),
    new TextRun({ text, size: 22, font: "Arial" }),
  ]});
}
function bullet(text, ref) {
  return new Paragraph({ numbering: { reference: ref || "bullets", level: 0 }, spacing: { after: 80 }, children: [new TextRun({ text, size: 22, font: "Arial" })] });
}
function bulletBold(label, text, ref) {
  return new Paragraph({ numbering: { reference: ref || "bullets", level: 0 }, spacing: { after: 80 }, children: [
    new TextRun({ text: label, size: 22, font: "Arial", bold: true }),
    new TextRun({ text, size: 22, font: "Arial" }),
  ]});
}

const doc = new Document({
  styles: {
    default: { document: { run: { font: "Arial", size: 22 } } },
    paragraphStyles: [
      { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 34, bold: true, font: "Arial", color: "1B1B1B" },
        paragraph: { spacing: { before: 360, after: 200 }, outlineLevel: 0 } },
      { id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 26, bold: true, font: "Arial", color: "444444" },
        paragraph: { spacing: { before: 240, after: 140 }, outlineLevel: 1 } },
    ]
  },
  numbering: {
    config: [
      { reference: "bullets", levels: [{ level: 0, format: LevelFormat.BULLET, text: "\u2022", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] },
      { reference: "nums", levels: [{ level: 0, format: LevelFormat.DECIMAL, text: "%1.", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] },
      { reference: "nums2", levels: [{ level: 0, format: LevelFormat.DECIMAL, text: "%1.", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] },
    ]
  },
  sections: [{
    properties: {
      page: {
        size: { width: 12240, height: 15840 },
        margin: { top: 1440, right: 1200, bottom: 1440, left: 1200 }
      }
    },
    headers: { default: new Header({ children: [new Paragraph({ alignment: AlignmentType.RIGHT, children: [new TextRun({ text: "\uD3F0\uCF00\uC774\uC2A4 \uAC10\uC131 \uC2A4\uBAB0\uBE0C\uB79C\uB4DC \uB808\uD37C\uB7F0\uC2A4 \uAC00\uC774\uB4DC", italics: true, color: "999999", size: 17 })] })] }) },
    footers: { default: new Footer({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "Page ", size: 17, color: "999999" }), new TextRun({ children: [PageNumber.CURRENT], size: 17, color: "999999" })] })] }) },
    children: [

      // ===== COVER =====
      new Paragraph({ spacing: { before: 2400 }, children: [] }),
      new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 80 }, children: [new TextRun({ text: "\uD3F0\uCF00\uC774\uC2A4 \uAC10\uC131 \uC2A4\uBAB0\uBE0C\uB79C\uB4DC", size: 48, bold: true, color: "1B1B1B" })] }),
      new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 100 }, children: [new TextRun({ text: "\uC2E4\uC804 \uB808\uD37C\uB7F0\uC2A4 \uAC00\uC774\uB4DC", size: 40, bold: true, color: "E04040" })] }),
      new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 300 }, border: { top: { style: BorderStyle.SINGLE, size: 4, color: "E04040", space: 1 } }, children: [] }),
      new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 80 }, children: [new TextRun({ text: "\uB514\uC790\uC778 \uC694\uC18C \u00B7 \uCEEC\uB7EC\uCF54\uB4DC \u00B7 \uAD11\uACE0 \uCE74\uD53C \u00B7 \uAC00\uACA9 \uC804\uB7B5 \u00B7 \uBA54\uD0C0\uAD11\uACE0 \uC801\uC6A9\uBC95", size: 22, color: "888888" })] }),
      new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "2026\uB144 4\uC6D4 | \uD55C\uAD6D \uC18C\uB9E4 \uC2DC\uC7A5 \uAE30\uC900", size: 20, color: "AAAAAA" })] }),

      new Paragraph({ children: [new PageBreak()] }),

      // ===== 1. 이 가이드 사용법 =====
      heading1("1. \uC774 \uAC00\uC774\uB4DC \uC0AC\uC6A9\uBC95"),
      body("\uC774 \uBB38\uC11C\uB294 \uD55C\uAD6D \uBA54\uD0C0\uAD11\uACE0\uC5D0\uC11C \uAC10\uC131\uC73C\uB85C \uC131\uACF5\uD55C \uC2A4\uBAB0\uBE0C\uB79C\uB4DC 8\uACF3\uC758 \uC2E4\uC81C \uC81C\uD488\u00B7\uAD11\uACE0\u00B7\uBE0C\uB79C\uB529\uC744 \uBD84\uC11D\uD574\uC11C, \uBC14\uB85C \uB808\uD37C\uB7F0\uC2A4\uB85C \uC801\uC6A9\uD560 \uC218 \uC788\uB3C4\uB85D \uC815\uB9AC\uD588\uC2B5\uB2C8\uB2E4."),
      body("\uAC01 \uBE0C\uB79C\uB4DC\uB9C8\uB2E4 \uB2E4\uC74C\uC744 \uB2F4\uC558\uC2B5\uB2C8\uB2E4:"),
      bulletBold("\uB514\uC790\uC778 DNA: ", "\uBE0C\uB79C\uB4DC \uC815\uCCB4\uC131\uC744 \uB9CC\uB4DC\uB294 \uD575\uC2EC \uB514\uC790\uC778 \uC694\uC18C"),
      bulletBold("\uCEEC\uB7EC \uD314\uB808\uD2B8: ", "\uC2E4\uC81C \uC0AC\uC6A9\uD558\uB294 \uCEEC\uB7EC\uC640 \uC870\uD569"),
      bulletBold("\uBCA0\uC2A4\uD2B8\uC140\uB7EC \uAD6C\uC870: ", "\uAC00\uC7A5 \uC798 \uD314\uB9AC\uB294 \uC81C\uD488\uC758 \uAD6C\uC131 \uC694\uC18C"),
      bulletBold("\uAC00\uACA9 \uC804\uB7B5: ", "\uAC00\uACA9\uB300\uC640 \uD560\uC778/\uC138\uD2B8 \uAD6C\uC131"),
      bulletBold("\uAD11\uACE0 \uCE74\uD53C \uB808\uD37C\uB7F0\uC2A4: ", "\uBA54\uD0C0\uAD11\uACE0\uC5D0 \uBC14\uB85C \uC4F8 \uC218 \uC788\uB294 \uBB38\uAD6C/\uCEE8\uC149"),
      bulletBold("\uC0AC\uC774\uD2B8 \uD1A4\uC564\uB9E4\uB108: ", "\uBE0C\uB79C\uB4DC \uBD84\uC704\uAE30\uB97C \uB9CC\uB4DC\uB294 \uBE44\uC8FC\uC5BC \uC694\uC18C"),

      new Paragraph({ children: [new PageBreak()] }),

      // ===== 2. 브랜드별 레퍼런스 시트 =====
      heading1("2. \uBE0C\uB79C\uB4DC\uBCC4 \uB808\uD37C\uB7F0\uC2A4 \uC2DC\uD2B8"),

      // --- 2-1. 어프어프 ---
      heading2("2-1. \uC5B4\uD504\uC5B4\uD504 (EARPEARP)"),
      new Paragraph({ spacing: { after: 60 }, children: [new TextRun({ text: "earpearp.com | @earp_earp (59K)", size: 20, color: "888888", italics: true })] }),

      new Table({
        width: { size: 9840, type: WidthType.DXA },
        columnWidths: [2460, 7380],
        rows: [
          new TableRow({ children: [
            bCellBold("\uBE0C\uB79C\uB4DC DNA", 2460, "FFF5F5"),
            bCell("\uD0A4\uCE58\uD55C \uCEEC\uB7EC\uAC10 + \uC790\uCCB4 \uCE90\uB9AD\uD130 '\uCF54\uBE44' \uC911\uC2EC. \uADC0\uC5EC\uC6C0\uACFC \uD3FD \uAC10\uC131\uC744 \uB3D9\uC2DC\uC5D0 \uC7A1\uC740 \uBE44\uC8FC\uC5BC. \uCE90\uB9AD\uD130\uBCC4 \uC2DC\uB9AC\uC988(8\uC885: \uCF54\uBE44, \uCE58\uCE58, \uD3EC\uD3EC, \uD30C\uCF54 \uB4F1)\uB85C \uCDE8\uD5A5 \uC138\uBD84\uD654", 7380, "FFF5F5"),
          ] }),
          new TableRow({ children: [
            bCellBold("\uCEEC\uB7EC \uD314\uB808\uD2B8", 2460),
            bCell("\uD06C\uB9BC/\uD654\uC774\uD2B8 \uBC30\uACBD + \uD551\uD06C\u00B7\uD37C\uD50C\u00B7\uBBFC\uD2B8 \uD3EC\uC778\uD2B8 \uCEEC\uB7EC. \uC2E4\uBC84(SILVER) \uBBF8\uB7EC \uB9C8\uAC10\uC774 \uD504\uB9AC\uBBF8\uC5C4 \uB77C\uC778. \uD22C\uBA85(CLEAR) + \uBE44\uBE44\uB4DC \uCEEC\uB7EC \uD504\uB808\uC784\uC774 \uAE30\uBCF8 \uAD6C\uC131", 7380),
          ] }),
          new TableRow({ children: [
            bCellBold("\uBCA0\uC2A4\uD2B8\uC140\uB7EC \uAD6C\uC870", 2460, "FFF5F5"),
            bCell("\uC6E8\uC774\uBE0C\uB77C\uBCA8 \uCF00\uC774\uC2A4(29,000\uC6D0) > \uC5D0\uD3ED\uC2DC(28,000\uC6D0) > \uD558\uB4DC(19,000\uC6D0). \uCF5C\uB77C\uBCF4 \uD55C\uC815\uD310(\uC5B4\uD504\uC5B4\uD504X\uD770\uB514 \uBE0C\uB7EC\uD50C\uB77C\uC774 39,000\u219229,800\uC6D0)\uC774 \uD654\uC81C\uC131 \uC8FC\uB3C4. \uCF00\uC774\uC2A4+\uC5D0\uC5B4\uD31F+\uD30C\uC6B0\uCE58\uB85C \uB77C\uC774\uD504\uC2A4\uD0C0\uC77C \uD655\uC7A5", 7380, "FFF5F5"),
          ] }),
          new TableRow({ children: [
            bCellBold("\uAC00\uACA9 \uC804\uB7B5", 2460),
            bCell("\uD558\uB4DC 19,000 / \uC5D0\uD3ED\uC2DC 28,000 / \uC6E8\uC774\uBE0C\uB77C\uBCA8 29,000 / \uCF5C\uB77C\uBCF4 39,000\uC6D0. \uAE30\uAC04\uD55C\uC815 \uC138\uC77C(20~40% \uD560\uC778)\uB85C \uAE34\uBC15\uAC10 \uC870\uC131. \uD0A4\uCEA1\uD0A4\uB9C1 7,900\uC6D0\uC73C\uB85C \uC800\uAC00 \uC9C4\uC785\uC0C1\uD488 \uC6B4\uC601", 7380),
          ] }),
          new TableRow({ children: [
            bCellBold("\uAD11\uACE0 \uCE74\uD53C \uB808\uD37C\uB7F0\uC2A4", 2460, "FFF5F5"),
            bCellMulti([
              new TextRun({ text: "\"\uCF54\uBE44\uB791 \uAC19\uC774 \uBD04 \uB9DE\uC774\uD574\uC694\" ", font: "Arial", size: 19, bold: true }),
              new TextRun({ text: "/ \"\uAE30\uAC04\uD55C\uC815 \uC138\uC77C\" / \"\uC5B4\uD504\uC5B4\uD504X\uD770\uB514 \uCF5C\uB77C\uBCF4 \uD55C\uC815\uD310\" / \"\uBBF8\uB7EC\uCF00\uC774\uC2A4 \uC2E0\uC0C1 \uCD9C\uC2DC\". \uCE90\uB9AD\uD130 \uAC10\uC131 + \uD55C\uC815/\uCF5C\uB77C\uBCF4\uB85C \uD76C\uC18C\uC131 \uAC15\uC870\uD558\uB294 \uD328\uD134", font: "Arial", size: 19 }),
            ], 7380, "FFF5F5"),
          ] }),
          new TableRow({ children: [
            bCellBold("\uC0AC\uC774\uD2B8 \uD1A4\uC564\uB9E4\uB108", 2460),
            bCell("\uBC1D\uC740 \uD06C\uB9BC\uC0C9 \uBC30\uACBD + Pretendard \uD3F0\uD2B8. \uD551\uD06C/\uD654\uC774\uD2B8 \uD1A4. \uCE90\uB9AD\uD130\uBCC4 \uCE74\uD14C\uACE0\uB9AC \uBD84\uB958\uB85C \uCDE8\uD5A5 \uD0D0\uC0C9 \uC720\uB3C4. \uBB34\uC2E0\uC0AC\u00B7\uC9C0\uADF8\uC7AC\uADF8 \uC785\uC810", 7380),
          ] }),
        ]
      }),
      body("\uB808\uD37C\uB7F0\uC2A4 \uD3EC\uC778\uD2B8: \uCE90\uB9AD\uD130 \uAE30\uBC18 \uC2DC\uB9AC\uC988 \uC804\uB7B5 + \uCF5C\uB77C\uBCF4/\uD55C\uC815\uD310 \uD654\uC81C\uC131 \uB9C8\uCF00\uD305 + \uC800\uAC00 \uC9C4\uC785\uC0C1\uD488(\uD0A4\uB9C1)\uC73C\uB85C \uC2E0\uADDC\uACE0\uAC1D \uC720\uC785", { bold: true, color: "E04040" }),

      new Paragraph({ children: [new PageBreak()] }),

      // --- 2-2. 세컨드유니크네임 ---
      heading2("2-2. \uC138\uCEE8\uB4DC\uC720\uB2C8\uD06C\uB124\uC784 (SECOND UNIQUE NAME)"),
      new Paragraph({ spacing: { after: 60 }, children: [new TextRun({ text: "youngboyz.co.kr | @youngboyz_sun (18K)", size: 20, color: "888888", italics: true })] }),

      new Table({
        width: { size: 9840, type: WidthType.DXA },
        columnWidths: [2460, 7380],
        rows: [
          new TableRow({ children: [
            bCellBold("\uBE0C\uB79C\uB4DC DNA", 2460, "F5F0FF"),
            bCell("\uD3F0\uCF00\uC774\uC2A4 = \uD328\uC158 \uC545\uC138\uC11C\uB9AC. \uCEEC\uB7EC\uBE14\uB85D + \uC2A4\uD2B8\uB7A9/\uD328\uCE58/\uB9AC\uBCF8\uC73C\uB85C \uC774\uBBF8\uC9C0 \uBCC0\uD615. \uC704\uD2B8 \uC788\uB294 \uCEE8\uC149\uC73C\uB85C \uD328\uC158 \uAC10\uC131 \uC804\uB2EC", 7380, "F5F0FF"),
          ] }),
          new TableRow({ children: [
            bCellBold("\uCEEC\uB7EC \uD314\uB808\uD2B8", 2460),
            bCell("\uBE44\uBE44\uB4DC \uB2E8\uC0C9 \uCEEC\uB7EC\uBE14\uB85D(Yellow, Red, Pink, Sky, Purple, Green). \uC544\uC774\uBCF4\uB9AC + \uADF8\uB808\uC774 \uB274\uD2B8\uB7F4 \uBCA0\uC774\uC2A4. \uB370\uB2D8/\uB2C8\uD2B8/\uCCB4\uD06C \uD14D\uC2A4\uCC98 \uD65C\uC6A9", 7380),
          ] }),
          new TableRow({ children: [
            bCellBold("\uBCA0\uC2A4\uD2B8\uC140\uB7EC \uAD6C\uC870", 2460, "F5F0FF"),
            bCell("SUN CASE \uB77C\uC778: BELT(\uC2A4\uD2B8\uB7A9), PATCH(\uD328\uCE58\uBD99\uC774\uAE30), STRING(\uC2A4\uD2B8\uB9C1), CLEAR(\uD22C\uBA85), GRAPHIC(\uADF8\uB798\uD53D). \uAC01 \uB77C\uC778 24,000~33,000\uC6D0. \uD3E8\uCF00\uC774\uC2A4+\uD30C\uC6B0\uCE58+\uD5E4\uB4DC\uC6E8\uC5B4 \uD06C\uB85C\uC2A4\uC140\uB9C1", 7380, "F5F0FF"),
          ] }),
          new TableRow({ children: [
            bCellBold("\uAC00\uACA9 \uC804\uB7B5", 2460),
            bCell("GRAPHIC 24,000 / CLEAR PATCH 30,000 / COZY BEAR 30,000~33,000\uC6D0. STAR POUCH 16,000\uC6D0\uC73C\uB85C \uD3EC\uC778\uD2B8 \uC0C1\uD488 \uC6B4\uC601. \uBB34\uC2E0\uC0AC\u00B729CM\u00B7W\uCEE8\uC149 \uC785\uC810\uC73C\uB85C \uCC44\uB110 \uB2E4\uBCC0\uD654", 7380),
          ] }),
          new TableRow({ children: [
            bCellBold("\uAD11\uACE0 \uCE74\uD53C \uB808\uD37C\uB7F0\uC2A4", 2460, "F5F0FF"),
            bCellMulti([
              new TextRun({ text: "\"\uB098\uB9CC\uC758 SUN CASE \uC870\uD569 \uB9CC\uB4E4\uAE30\" ", font: "Arial", size: 19, bold: true }),
              new TextRun({ text: "/ \"\uC2A4\uD2B8\uB7A9 \uBC14\uAFB8\uACE0 \uD328\uCE58 \uBD99\uC774\uACE0 \uB0A0\uB9C8\uB2E4 \uB2E4\uB978 \uCF00\uC774\uC2A4\" / \"\uCF54\uC9C0\uBCA0\uC5B4 \uC2E0\uC0C1 \uCD9C\uC2DC\". \uCEE4\uC2A4\uD130\uB9C8\uC774\uC9D5 \uCEE8\uC149 + \uC2DC\uC98C\uAC10 \uC2E0\uC0C1\uC774 \uD575\uC2EC", font: "Arial", size: 19 }),
            ], 7380, "F5F0FF"),
          ] }),
          new TableRow({ children: [
            bCellBold("\uC0AC\uC774\uD2B8 \uD1A4\uC564\uB9E4\uB108", 2460),
            bCell("\uD654\uC774\uD2B8 \uBC30\uACBD + \uBBF8\uB2C8\uBA40 \uADF8\uB9AC\uB4DC. \uC601\uBB38 \uB300\uBB38\uC790 \uC0B0\uC138\uB9AC\uD504. \uC81C\uD488 \uC911\uC2EC \uB808\uC774\uC544\uC6C3. \uBAA8\uB358\uD558\uACE0 \uCEA0\uC8FC\uC5BC\uD55C \uC80A\uC740 \uAC10\uC131", 7380),
          ] }),
        ]
      }),
      body("\uB808\uD37C\uB7F0\uC2A4 \uD3EC\uC778\uD2B8: \uD328\uC158 \uC545\uC138\uC11C\uB9AC \uD3EC\uC9C0\uC154\uB2DD + \uCEE4\uC2A4\uD130\uB9C8\uC774\uC9D5 \uCEE8\uC149 + \uD06C\uB85C\uC2A4\uC140\uB9C1 \uC804\uB7B5", { bold: true, color: "E04040" }),

      new Paragraph({ children: [new PageBreak()] }),

      // --- 2-3. 소유마실 ---
      heading2("2-3. \uC18C\uC720\uB9C8\uC2E4 (SOYOUMASIL)"),
      new Paragraph({ spacing: { after: 60 }, children: [new TextRun({ text: "soyoumasil.com | \uBB34\uC2E0\uC0AC\u00B7W\uCEE8\uC149 \uC785\uC810", size: 20, color: "888888", italics: true })] }),

      new Table({
        width: { size: 9840, type: WidthType.DXA },
        columnWidths: [2460, 7380],
        rows: [
          new TableRow({ children: [
            bCellBold("\uBE0C\uB79C\uB4DC DNA", 2460, "F0FFF5"),
            bCell("\uC790\uC5F0\uC5D0\uC11C \uC601\uAC10\uBC1B\uC740 \uC5D0\uCF54 \uAC10\uC131. \uBD80\uB4DC\uB7EC\uC6B4 \uD1A4\uACFC \uD14D\uC2A4\uCC98, \uD328\uBE0C\uB9AD \uAC10\uAC01\uC744 \uB514\uC9C0\uD138 \uC561\uC138\uC11C\uB9AC\uC5D0 \uC811\uBAA9. \uC8FC\uBB38\uC81C\uC791 4\uC77C \uC774\uB0B4 \uCD9C\uACE0\uB85C \uD504\uB9AC\uBBF8\uC5C4 \uAC10\uC131 \uC720\uC9C0", 7380, "F0FFF5"),
          ] }),
          new TableRow({ children: [
            bCellBold("\uCEEC\uB7EC \uD314\uB808\uD2B8", 2460),
            bCell("Haze Blue, Peach Cream, Sahara Beige, Soft Leopard Stone. \uD30C\uC2A4\uD154+\uB274\uD2B8\uB7F4 \uD1A4 \uC911\uC2EC. \uBE14\uB8E8/\uD551\uD06C/\uBCA0\uC774\uC9C0 \uC870\uD569\uC774 \uC8FC\uB825", 7380),
          ] }),
          new TableRow({ children: [
            bCellBold("\uBCA0\uC2A4\uD2B8\uC140\uB7EC \uAD6C\uC870", 2460, "F0FFF5"),
            bCell("\uC5D0\uD3ED\uC2DC\uBC94\uD37C\uCF00\uC774\uC2A4(\uBE14\uB799/\uC2E4\uBC84) > \uD22C\uBA85\uC824\uD558\uB4DC > \uD130\uD504\uBC94\uD37C+\uB9E5\uC138\uC774\uD504 > \uC5D0\uD3ED\uC2DC\uCE74\uB4DC\uCF00\uC774\uC2A4(\uCE74\uB4DC2\uC7A5). Full of Love, Wave, Gradient \uB4F1 \uC2DC\uC98C \uCEEC\uB809\uC158 \uC6B4\uC601", 7380, "F0FFF5"),
          ] }),
          new TableRow({ children: [
            bCellBold("\uAC00\uACA9 \uC804\uB7B5", 2460),
            bCell("\uD22C\uBA85\uC824\uD558\uB4DC 17,850~21,000 / \uC5D0\uD3ED\uC2DC\uBC94\uD37C 21,000~25,500 / \uC5D0\uD3ED\uC2DC\uCE74\uB4DC 25,500~30,000\uC6D0. 15% \uC0C1\uC2DC\uD560\uC778\uC73C\uB85C \uBCA0\uC774\uC2A4 \uAC00\uACA9 \uC815\uCC45. \uC5D0\uCF54\uBC31\uACFC \uC2A4\uB9C8\uD2B8\uD1A1\uC73C\uB85C \uBD80\uAC00 \uC0C1\uD488 \uAD6C\uC131", 7380),
          ] }),
          new TableRow({ children: [
            bCellBold("\uAD11\uACE0 \uCE74\uD53C \uB808\uD37C\uB7F0\uC2A4", 2460, "F0FFF5"),
            bCellMulti([
              new TextRun({ text: "\"NEW IN! Peach Cream\" ", font: "Arial", size: 19, bold: true }),
              new TextRun({ text: "/ \"BEST! \uC8FC\uBB38\uC774 \uB9CE\uC544\uC694\" / \"Blue haze plaid \uC2E0\uC0C1\". \uC2E0\uC0C1 \uC54C\uB9BC + \uC0AC\uD68C\uC801 \uC99D\uAC70(\uC8FC\uBB38\uB9CE\uC544\uC694) + \uCEEC\uB7EC\uBA85 \uC790\uCCB4\uAC00 \uCE90\uCE58\uD504\uB808\uC774\uC988\uC778 \uD328\uD134", font: "Arial", size: 19 }),
            ], 7380, "F0FFF5"),
          ] }),
          new TableRow({ children: [
            bCellBold("\uC0AC\uC774\uD2B8 \uD1A4\uC564\uB9E4\uB108", 2460),
            bCell("\uAC10\uC131\uC801\uC774\uACE0 \uCE5C\uADFC\uD55C \uD1A4. \uC774\uBAA8\uC9C0 \uD65C\uC6A9\uD55C \uCE90\uC8FC\uC5BC\uD568. \uC5D0\uCF54/\uD328\uBE0C\uB9AD \uD0A4\uC6CC\uB4DC\uB85C \uBE0C\uB79C\uB4DC \uBD84\uC704\uAE30 \uC870\uC131. \uBB34\uC2E0\uC0AC\u00B7W\uCEE8\uC149 \uC785\uC810\uC73C\uB85C \uC2E0\uB8B0\uAC10 \uD655\uBCF4", 7380),
          ] }),
        ]
      }),
      body("\uB808\uD37C\uB7F0\uC2A4 \uD3EC\uC778\uD2B8: \uCEEC\uB7EC\uBA85 \uC790\uCCB4\uAC00 \uBE0C\uB79C\uB529 + \uC790\uC5F0 \uAC10\uC131 \uD14D\uC2A4\uCC98 + \uC2DC\uC98C \uCEEC\uB809\uC158 \uC804\uB7B5", { bold: true, color: "E04040" }),

      // --- 2-4. 하우위 ---
      heading2("2-4. \uD558\uC6B0\uC704 (howie)"),
      new Paragraph({ spacing: { after: 60 }, children: [new TextRun({ text: "howie.co.kr", size: 20, color: "888888", italics: true })] }),

      new Table({
        width: { size: 9840, type: WidthType.DXA },
        columnWidths: [2460, 7380],
        rows: [
          new TableRow({ children: [
            bCellBold("\uBE0C\uB79C\uB4DC DNA", 2460, "F5FAFF"),
            bCell("\uBBF8\uB2C8\uBA40 + \uBB34\uB4DC \uCEEC\uB7EC. \uBCF5\uC7A1\uD558\uC9C0 \uC54A\uC740 \uC790\uC5F0\uC2A4\uB7EC\uC6C0. \uD3EC\uC778\uD2B8\uAC00 \uB418\uB294 \uCEEC\uB7EC\uC640 \uC9C1\uAD00\uC801\uC778 \uB514\uC790\uC778\uC73C\uB85C \uC77C\uC0C1 \uC18D \uC870\uD654\uB97C \uCD94\uAD6C", 7380, "F5FAFF"),
          ] }),
          new TableRow({ children: [
            bCellBold("\uCEEC\uB7EC \uD314\uB808\uD2B8", 2460),
            bCell("\uB274\uD2B8\uB7F4 \uD1A4 \uC911\uC2EC: Sand, Clay, Cream, Stone, Sage. \uD3EC\uC778\uD2B8 \uCEEC\uB7EC\uB85C \uBB34\uB4DC\uAC10 \uC804\uB2EC. \uACFC\uD558\uC9C0 \uC54A\uC740 \uCEEC\uB7EC \uC870\uD569", 7380),
          ] }),
          new TableRow({ children: [
            bCellBold("\uBCA0\uC2A4\uD2B8\uC140\uB7EC \uAD6C\uC870", 2460, "F5FAFF"),
            bCell("DIGITAL(\uD3F0\uCF00\uC774\uC2A4/\uADF8\uB9BD\uD1A1) + LIVING + ACC + BAG \uCE74\uD14C\uACE0\uB9AC. HOWIE SELECT(\uD050\uB808\uC774\uC158 \uC544\uC774\uD15C) \uCE74\uD14C\uACE0\uB9AC\uB85C \uBE0C\uB79C\uB4DC \uCDE8\uD5A5 \uC804\uB2EC. \uB77C\uC774\uD504\uC2A4\uD0C0\uC77C \uD655\uC7A5 \uC804\uB7B5", 7380, "F5FAFF"),
          ] }),
          new TableRow({ children: [
            bCellBold("\uAD11\uACE0 \uCE74\uD53C \uB808\uD37C\uB7F0\uC2A4", 2460),
            bCellMulti([
              new TextRun({ text: "\"\uD3B8\uC548\uD558\uACE0 \uC870\uD654\uB85C\uC6B4 \uC77C\uC0C1\" ", font: "Arial", size: 19, bold: true }),
              new TextRun({ text: "/ \"\uC5B4\uB5A4 \uC21C\uAC04\uC5D0\uB3C4 \uC790\uC5F0\uC2A4\uB7EC\uC6B4\". \uBBF8\uB2C8\uBA40 \uBB34\uB4DC \uCEEC\uB7EC + \uC77C\uC0C1 \uC2A4\uD0C0\uC77C\uB9C1 \uCEE8\uC149. \uC81C\uD488 \uC790\uCCB4\uBCF4\uB2E4 '\uBD84\uC704\uAE30'\uB97C \uD30C\uB294 \uD1A4", font: "Arial", size: 19 }),
            ], 7380),
          ] }),
          new TableRow({ children: [
            bCellBold("\uC0AC\uC774\uD2B8 \uD1A4\uC564\uB9E4\uB108", 2460, "F5FAFF"),
            bCell("\uD654\uC774\uD2B8 \uBC30\uACBD + \uC601\uBB38 \uB300\uBB38\uC790 \uC0B0\uC138\uB9AC\uD504. \uD55C/\uC601 \uC774\uC911\uC5B8\uC5B4. \uAE68\uB057\uD558\uACE0 \uC138\uB828\uB41C \uB77C\uC774\uD504\uC2A4\uD0C0\uC77C \uD50C\uB7AB\uD3FC \uBD84\uC704\uAE30", 7380, "F5FAFF"),
          ] }),
        ]
      }),
      body("\uB808\uD37C\uB7F0\uC2A4 \uD3EC\uC778\uD2B8: \uBBF8\uB2C8\uBA40 \uD1A4 \uBE0C\uB79C\uB529 + '\uBD84\uC704\uAE30' \uD30C\uB294 \uAD11\uACE0 + \uB77C\uC774\uD504\uC2A4\uD0C0\uC77C \uD655\uC7A5 \uC804\uB7B5", { bold: true, color: "E04040" }),

      new Paragraph({ children: [new PageBreak()] }),

      // --- 2-5. 가르송티미드 ---
      heading2("2-5. \uAC00\uB974\uC1A1\uD2F0\uBBF8\uB4DC (GARCONTIMIDE)"),
      new Paragraph({ spacing: { after: 60 }, children: [new TextRun({ text: "garcontimide.com | \uBB34\uC2E0\uC0AC\u00B729CM \uC785\uC810", size: 20, color: "888888", italics: true })] }),

      new Table({
        width: { size: 9840, type: WidthType.DXA },
        columnWidths: [2460, 7380],
        rows: [
          new TableRow({ children: [
            bCellBold("\uBE0C\uB79C\uB4DC DNA", 2460, "FFFBF0"),
            bCell("\"\uADC0\uC5EC\uC6B4 \uAC83\uC774 \uC138\uC0C1\uC744 \uAD6C\uD55C\uB2E4\". \uC11C\uD234\uD55C \uB450\uB4E4 \uC544\uD2B8\uB97C \uC791\uD488\uC73C\uB85C. \uC790\uCCB4 \uCE90\uB9AD\uD130 '\uAF2C\uC21C'(\uAC15\uC544\uC9C0) \uC911\uC2EC\uC758 \uB530\uB73B\uD55C \uC720\uBA38", 7380, "FFFBF0"),
          ] }),
          new TableRow({ children: [
            bCellBold("\uCEEC\uB7EC \uD314\uB808\uD2B8", 2460),
            bCell("\uBCA0\uC774\uCEE4\uB9AC \uBE0C\uB77C\uC6B4, \uC2A4\uCE74\uC774\uBE14\uB8E8, \uBC84\uAC74\uB514, \uBE14\uB799, \uD654\uC774\uD2B8, \uD551\uD06C. \uC2E4\uBC84 \uC5D0\uD3ED\uC2DC\uAC00 \uD504\uB9AC\uBBF8\uC5C4 \uB77C\uC778. \uBB34\uC9C0 \uCEEC\uB7EC + \uCE90\uB9AD\uD130 \uC870\uD569", 7380),
          ] }),
          new TableRow({ children: [
            bCellBold("\uBCA0\uC2A4\uD2B8\uC140\uB7EC \uAD6C\uC870", 2460, "FFFBF0"),
            bCell("\uC2E4\uBC84\uC5D0\uD3ED\uC2DC(25,000\uC6D0) > \uD22C\uBA85\uCF00\uC774\uC2A4(20,000\uC6D0). \uAF2C\uC21C \uC5C9\uB369\uC774 \uC2DC\uB9AC\uC988, \uC0D0\uB728 \uB3C4\uD2B8, \uB7EC\uBE0C\uBBF8\uD37C\uC2A4\uD2B8, \uC11C\uD551\uAF2C\uC21C \uB4F1 \uD14C\uB9C8\uBCC4 \uC2DC\uB9AC\uC988. \uB79C\uB364 \uB514\uC790\uC778 9,999\uC6D0\uC73C\uB85C \uC800\uAC00 \uC9C4\uC785", 7380, "FFFBF0"),
          ] }),
          new TableRow({ children: [
            bCellBold("\uAC00\uACA9 \uC804\uB7B5", 2460),
            bCell("\uD22C\uBA85 20,000 / \uC2E4\uBC84\uC5D0\uD3ED\uC2DC 25,000\uC6D0. 2\uB2E8\uACC4 \uAC00\uACA9\uC81C. \uB79C\uB364\uB514\uC790\uC778 9,999\uC6D0\uC740 \uC8FC\uBB38 \uC720\uC785\uC6A9 \uBBF8\uB07C\uC0C1\uD488. \uBB34\uC2E0\uC0AC\u00B729CM \uC785\uC810\uC73C\uB85C \uCC44\uB110 \uC2E0\uB8B0\uAC10 \uD655\uBCF4", 7380),
          ] }),
          new TableRow({ children: [
            bCellBold("\uAD11\uACE0 \uCE74\uD53C \uB808\uD37C\uB7F0\uC2A4", 2460, "FFFBF0"),
            bCellMulti([
              new TextRun({ text: "\"\uADC0\uC5EC\uC6B4 \uAC83\uC774 \uC138\uC0C1\uC744 \uAD6C\uD55C\uB2E4\" ", font: "Arial", size: 19, bold: true }),
              new TextRun({ text: "/ \"\uAF2C\uC21C \uC5C9\uB369\uC774 \uC2E4\uBC84 \uC5D0\uD3ED\uC2DC \uC2E0\uC0C1\" / \"\uB79C\uB364 \uB514\uC790\uC778 9,999\uC6D0\". \uCE90\uB9AD\uD130 \uAC10\uC131 + \uC800\uAC00 \uBBF8\uB07C \uC0C1\uD488\uC73C\uB85C \uC2E0\uADDC \uC720\uC785 \uC720\uB3C4", font: "Arial", size: 19 }),
            ], 7380, "FFFBF0"),
          ] }),
          new TableRow({ children: [
            bCellBold("\uC0AC\uC774\uD2B8 \uD1A4\uC564\uB9E4\uB108", 2460),
            bCell("\uB530\uB73B\uD558\uACE0 \uADC0\uC5EC\uC6B4 \uD1A4. \uB450\uB4E4 \uC77C\uB7EC\uC2A4\uD2B8 \uC911\uC2EC. \uCE90\uB9AD\uD130 \uC2DC\uB9AC\uC988\uBCC4 \uCE74\uD14C\uACE0\uB9AC \uAD6C\uC131. \uBB34\uC2E0\uC0AC\u00B729CM \uC785\uC810\uC73C\uB85C 2030 \uC5EC\uC131 \uD0C0\uAC9F", 7380),
          ] }),
        ]
      }),
      body("\uB808\uD37C\uB7F0\uC2A4 \uD3EC\uC778\uD2B8: \uCE90\uB9AD\uD130 \uC2DC\uB9AC\uC988 \uC804\uB7B5 + \uB79C\uB364\uBC15\uC2A4 \uBBF8\uB07C\uC0C1\uD488 + 2\uB2E8\uACC4 \uAC00\uACA9\uC81C(\uD22C\uBA85/\uC5D0\uD3ED\uC2DC)", { bold: true, color: "E04040" }),

      // --- 2-6. 아우렐 ---
      heading2("2-6. \uC544\uC6B0\uB810 (Aurel)"),
      new Paragraph({ spacing: { after: 60 }, children: [new TextRun({ text: "aurel.kr | \uC9C0\uADF8\uC7AC\uADF8 \uC785\uC810", size: 20, color: "888888", italics: true })] }),

      new Table({
        width: { size: 9840, type: WidthType.DXA },
        columnWidths: [2460, 7380],
        rows: [
          new TableRow({ children: [
            bCellBold("\uBE0C\uB79C\uB4DC DNA", 2460, "FFF8F5"),
            bCell("\uD504\uB9AC\uBBF8\uC5C4 \uC77C\uB7EC\uC2A4\uD2B8 \uAC10\uC131. \uC81C\uD488\uBA85\uC774 \uACE7 \uC2DC\uC801 \uCE90\uCE58\uD504\uB808\uC774\uC988(\"\uC783\uD600\uC9C4 \uAFC8\uC758 \uD754\uC801\", \"\uB2EC\uCF64\uD55C \uC0C9\uC758 \uAFC8\", \"\uC2EC\uC5F0\uC758 \uC0AC\uB9C9\"). \uBABD\uD658\uC801 \uBB34\uB4DC + \uACE0\uC591\uC774 \uCE90\uB9AD\uD130", 7380, "FFF8F5"),
          ] }),
          new TableRow({ children: [
            bCellBold("\uCEEC\uB7EC \uD314\uB808\uD2B8", 2460),
            bCell("\uD30C\uC2A4\uD154 \uD1A4: \uB77C\uBCA4\uB354, \uBBF8\uC2A4\uD2F0 \uBE14\uB8E8, \uB85C\uC988 \uD551\uD06C, \uB180 \uC624\uB80C\uC9C0. \uC790\uC5F0 \uBAA8\uD2F0\uD504: \uBC14\uB2E4 \uBB3C\uACB0, \uC5BC\uC74C, \uB178\uC744, \uBCC4\uBE5B. \uCD94\uC0C1\uD654/\uC720\uD654 \uD130\uCE58\uC758 \uC544\uD2B8\uC6CC\uD06C \uC2A4\uD0C0\uC77C", 7380),
          ] }),
          new TableRow({ children: [
            bCellBold("\uBCA0\uC2A4\uD2B8\uC140\uB7EC \uAD6C\uC870", 2460, "FFF8F5"),
            bCell("\uC824\uD558\uB4DC \uCF00\uC774\uC2A4 \uC911\uC2EC. \uC815\uAC00 27,900 \u2192 \uD310\uB9E4\uAC00 19,500\uC6D0(30% \uD560\uC778). \uAC10\uC131 \uC77C\uB7EC\uC2A4\uD2B8 6\uC885, \uBC14\uB2E4 \uBB3C\uACB0 6\uC885, \uC560\uB2C8\uBA54\uC774\uC158 6\uC885 \uB4F1 \uD14C\uB9C8\uBCC4 \uBB36\uC74C \uD310\uB9E4. \uBB34\uB8CC\uBC30\uC1A1", 7380, "FFF8F5"),
          ] }),
          new TableRow({ children: [
            bCellBold("\uAD11\uACE0 \uCE74\uD53C \uB808\uD37C\uB7F0\uC2A4", 2460),
            bCellMulti([
              new TextRun({ text: "\"\uC783\uD600\uC9C4 \uAFC8\uC758 \uD754\uC801\uC744 \uC190\uC5D0 \uB2F4\uB2E4\" ", font: "Arial", size: 19, bold: true }),
              new TextRun({ text: "/ \"\uD55C \uC0C1\uD488\uB9CC \uAD6C\uB9E4\uD574\uB3C4 \uBB34\uB8CC\uBC30\uC1A1\" / \"\uC120\uBB3C\uD558\uAE30 \uC88B\uC740 \uAC10\uC131 \uCF00\uC774\uC2A4\". \uC2DC\uC801 \uC81C\uD488\uBA85 + \uBB34\uB8CC\uBC30\uC1A1 + \uC120\uBB3C\uC6A9 \uCEE8\uC149", font: "Arial", size: 19 }),
            ], 7380),
          ] }),
          new TableRow({ children: [
            bCellBold("\uC0AC\uC774\uD2B8 \uD1A4\uC564\uB9E4\uB108", 2460, "FFF8F5"),
            bCell("\uAC10\uC815 \uD45C\uD604\uD615 \uC81C\uD488\uBA85. \uD504\uB9AC\uBBF8\uC5C4 \uAC10\uC131 \uC720\uC9C0. \uACE0\uC591\uC774 \uCE90\uB9AD\uD130\uB85C \uCE90\uC8FC\uC5BC\uD568 \uD130\uCE58. \uC544\uD2B8\uC6CC\uD06C \uAC10\uAC01\uC758 \uC81C\uD488 \uC0AC\uC9C4\uC774 \uD575\uC2EC", 7380, "FFF8F5"),
          ] }),
        ]
      }),
      body("\uB808\uD37C\uB7F0\uC2A4 \uD3EC\uC778\uD2B8: \uC2DC\uC801 \uC81C\uD488\uBA85 \uBE0C\uB79C\uB529 + \uD14C\uB9C8\uBCC4 \uBB36\uC74C \uD310\uB9E4 + \uC120\uBB3C\uC6A9 \uCEE8\uC149 \uAD11\uACE0", { bold: true, color: "E04040" }),

      new Paragraph({ children: [new PageBreak()] }),

      // --- 2-7. 하이우 ---
      heading2("2-7. \uD558\uC774\uC6B0 (hioo)"),
      new Paragraph({ spacing: { after: 60 }, children: [new TextRun({ text: "hioo.kr", size: 20, color: "888888", italics: true })] }),

      new Table({
        width: { size: 9840, type: WidthType.DXA },
        columnWidths: [2460, 7380],
        rows: [
          new TableRow({ children: [
            bCellBold("\uBE0C\uB79C\uB4DC DNA", 2460, "F8F5FF"),
            bCell("\"\uC624\uB798 \uACE1\uC5D0 \uB450\uACE0 \uC2F6\uC740 \uB514\uC790\uC778\". \uC808\uC81C\uB41C \uAC10\uC131, \uACFC\uD558\uC9C0 \uC54A\uC740 \uD3FC\uACFC \uCEEC\uB7EC. \uD2B8\uB80C\uB4DC\uBCF4\uB2E4 \uD0C0\uC784\uB808\uC2A4\uD568\uC744 \uCD94\uAD6C\uD558\uB294 \uBE0C\uB79C\uB4DC \uCCA0\uD559", 7380, "F8F5FF"),
          ] }),
          new TableRow({ children: [
            bCellBold("\uCEEC\uB7EC \uD314\uB808\uD2B8", 2460),
            bCell("\uBBF8\uB2C8\uBA40 \uD1A4: \uD654\uC774\uD2B8, \uBE14\uB799, \uB124\uC774\uBE44, \uBCA0\uC774\uC9C0. \uACFC\uD558\uC9C0 \uC54A\uC740 \uCEEC\uB7EC\uB85C \uC2DC\uC990\uC744 \uD0C0\uC9C0 \uC54A\uB294 \uD0C0\uC784\uB808\uC2A4 \uB514\uC790\uC778", 7380),
          ] }),
          new TableRow({ children: [
            bCellBold("\uBCA0\uC2A4\uD2B8\uC140\uB7EC \uAD6C\uC870", 2460, "F8F5FF"),
            bCell("2025 \uD3F0\uCF00\uC774\uC2A4 \uB9DB\uC9D1 TOP5 \uC120\uC815(@ahyunfrom). \uBBF8\uB2C8\uBA40 \uB77C\uC778\uC5C5 \uC911\uC2EC. \uC790\uC0AC\uBAB0 \uC9C1\uC811 \uD310\uB9E4\uB85C \uBE0C\uB79C\uB4DC \uD1B5\uC81C\uB825 \uC720\uC9C0. \uC7AC\uAD6C\uB9E4\uC728 \uB192\uC740 \uCDA9\uC131 \uACE0\uAC1D\uCE35", 7380, "F8F5FF"),
          ] }),
          new TableRow({ children: [
            bCellBold("\uAD11\uACE0 \uCE74\uD53C \uB808\uD37C\uB7F0\uC2A4", 2460),
            bCellMulti([
              new TextRun({ text: "\"\uC624\uB798 \uACE4\uC5D0 \uB450\uACE0 \uC2F6\uC740\" ", font: "Arial", size: 19, bold: true }),
              new TextRun({ text: "/ \"\uC9C8\uB9AC\uC9C0 \uC54A\uB294 \uB514\uC790\uC778\" / \"\uB9E4\uC77C \uB4E4\uC5B4\uB3C4 \uC88B\uC740\". \uD0C0\uC784\uB808\uC2A4 \uCEE8\uC149 + \uC7AC\uAD6C\uB9E4 \uC720\uB3C4 \uBA54\uC2DC\uC9C0", font: "Arial", size: 19 }),
            ], 7380),
          ] }),
          new TableRow({ children: [
            bCellBold("\uC0AC\uC774\uD2B8 \uD1A4\uC564\uB9E4\uB108", 2460, "F8F5FF"),
            bCell("\uD578\uD551\uD06C(#ea3394) + \uBE14\uB799 \uD3EC\uC778\uD2B8 \uCEEC\uB7EC. \uBBF8\uB2C8\uBA40\uD558\uC9C0\uB9CC \uAC1C\uC131 \uC788\uB294 \uBE0C\uB79C\uB4DC \uCEEC\uB7EC \uC2DC\uC2A4\uD15C", 7380, "F8F5FF"),
          ] }),
        ]
      }),
      body("\uB808\uD37C\uB7F0\uC2A4 \uD3EC\uC778\uD2B8: \uD0C0\uC784\uB808\uC2A4 \uBE0C\uB79C\uB529 + \uC790\uC0AC\uBAB0 \uC9C1\uD310 \uC804\uB7B5 + \uC7AC\uAD6C\uB9E4 \uC720\uB3C4 \uBA54\uC2DC\uC9C0", { bold: true, color: "E04040" }),

      // --- 2-8. 디자인스킨 ---
      heading2("2-8. \uB514\uC790\uC778\uC2A4\uD0A8 (DESIGNSKIN)"),
      new Paragraph({ spacing: { after: 60 }, children: [new TextRun({ text: "designskin.com", size: 20, color: "888888", italics: true })] }),

      new Table({
        width: { size: 9840, type: WidthType.DXA },
        columnWidths: [2460, 7380],
        rows: [
          new TableRow({ children: [
            bCellBold("\uBE0C\uB79C\uB4DC DNA", 2460, "F5F5F0"),
            bCell("2011\uB144 \uCC3D\uB9BD. \"\uC6B0\uC544\uD558\uACE0 \uC544\uB984\uB2E4\uC6B4 \uCF00\uC774\uC2A4 \uB514\uC790\uC778\uC774 \uC5C6\uC744\uAE4C?\" \uC5D0\uC11C \uCD9C\uBC1C. \uCF00\uC774\uC2A4 = \uD328\uC158\uC744 \uC815\uC758\uD55C \uD504\uB9AC\uBBF8\uC5C4 \uBE0C\uB79C\uB4DC. WHAT IS YOUR COLOR? \uCEE0\uD398\uC778", 7380, "F5F5F0"),
          ] }),
          new TableRow({ children: [
            bCellBold("\uCEEC\uB7EC \uD314\uB808\uD2B8", 2460),
            bCell("\uD53C\uB2C9\uC2A4 \uB9E4\uD2B8 \uCEEC\uB7EC \uC2DC\uB9AC\uC988(10+\uC0C9\uC0C1). \uD074\uB798\uC2DD \uC790\uC218 \uD328\uD134. \uC5D0\uD3ED\uC2DC/\uAE00\uB77C\uC2A4 \uADF8\uB798\uD53D \uC7A5\uC2DD. \uB514\uC988\uB2C8/\uC6F9\uD230 \uCF5C\uB77C\uBCF4 \uB77C\uC778", 7380),
          ] }),
          new TableRow({ children: [
            bCellBold("\uBCA0\uC2A4\uD2B8\uC140\uB7EC \uAD6C\uC870", 2460, "F5F5F0"),
            bCell("\uB9E5\uC138\uC774\uD504 \uC561\uC138\uC11C\uB9AC \uC138\uD2B8(\uB4C0\uC5BC\uC2A4\uB9C8\uD2B8\uB9C1 34,800 + \uCE74\uB4DC\uD3EC\uCF13 34,800). \uD480\uCEE4\uBC84\uCF00\uC774\uC2A4+\uCE74\uB4DC\uD3EC\uCF13 \uC138\uD2B8 74,600\u219252,200(\uD68C\uC6D0\uAC00). \uD074\uB798\uC2DD \uC790\uC218\uD3EC\uCF13 39,800\uC6D0. \uD504\uB9AC\uBBF8\uC5C4 \uAC00\uACA9\uB300", 7380, "F5F5F0"),
          ] }),
          new TableRow({ children: [
            bCellBold("\uAD11\uACE0 \uCE74\uD53C \uB808\uD37C\uB7F0\uC2A4", 2460),
            bCellMulti([
              new TextRun({ text: "\"WHAT IS YOUR COLOR?\" ", font: "Arial", size: 19, bold: true }),
              new TextRun({ text: "/ \"\uB9E5\uC138\uC774\uD504 \uD480\uCEE4\uBC84+\uCE74\uB4DC\uD3EC\uCF13 \uC138\uD2B8\" / \"\uD68C\uC6D0\uAC00 30% \uD560\uC778\". \uCEEC\uB7EC \uCEE8\uC149 + \uC138\uD2B8 \uAD6C\uC131 + \uD68C\uC6D0\uD600\uD0DD \uACB0\uD569 \uD328\uD134", font: "Arial", size: 19 }),
            ], 7380),
          ] }),
          new TableRow({ children: [
            bCellBold("\uC0AC\uC774\uD2B8 \uD1A4\uC564\uB9E4\uB108", 2460, "F5F5F0"),
            bCell("\uD504\uB9AC\uBBF8\uC5C4/\uB7ED\uC154\uB9AC \uD3EC\uC9C0\uC154\uB2DD. \uAE54\uB054\uD55C \uADF8\uB9AC\uB4DC. \uBA85\uD655\uD55C \uAC00\uACA9\uD45C\uC2DC + \uD68C\uC6D0\uAC00 \uD560\uC778. COLLABO \uCE74\uD14C\uACE0\uB9AC(\uB514\uC988\uB2C8/\uC6F9\uD230)\uB85C \uD654\uC81C\uC131", 7380, "F5F5F0"),
          ] }),
        ]
      }),
      body("\uB808\uD37C\uB7F0\uC2A4 \uD3EC\uC778\uD2B8: \uD504\uB9AC\uBBF8\uC5C4 \uD3EC\uC9C0\uC154\uB2DD + \uB9E5\uC138\uC774\uD504 \uC138\uD2B8 \uC804\uB7B5 + \uD68C\uC6D0\uAC00 \uD560\uC778 \uAD6C\uC870", { bold: true, color: "E04040" }),

      // --- 2-9. 던타운 ---
      heading2("2-9. \uB358\uD0C0\uC6B4 (Dawntown)"),
      new Paragraph({ spacing: { after: 60 }, children: [new TextRun({ text: "dawntown.co.kr | @dawntown.kr | \uBB34\uC2E0\uC0AC \uC785\uC810", size: 20, color: "888888", italics: true })] }),

      new Table({
        width: { size: 9840, type: WidthType.DXA },
        columnWidths: [2460, 7380],
        rows: [
          new TableRow({ children: [
            bCellBold("\uBE0C\uB79C\uB4DC DNA", 2460, "F0F8FF"),
            bCell("\"\uC190\uB05D\uC5D0\uC11C \uB290\uAEF4\uC9C0\uB294 \uC791\uACE0 \uC18C\uC18C\uD55C \uD589\uBCF5\". \uD0A4\uCE58\uD55C \uB3D9\uBB3C \uCE90\uB9AD\uD130(\uBC84\uAC70\uBA4D, \uD37C\uADF8\uB0AB\uB514\uC2A4\uD130\uBE0C) \uC911\uC2EC. \uC8FC\uBB38\uC81C\uC791 1:1 \uC624\uB354\uBA54\uC774\uB4DC\uB85C \uD504\uB9AC\uBBF8\uC5C4 \uAC10\uC131 \uC720\uC9C0", 7380, "F0F8FF"),
          ] }),
          new TableRow({ children: [
            bCellBold("\uCEEC\uB7EC \uD314\uB808\uD2B8", 2460),
            bCell("\uD30C\uC2A4\uD154 \uD1A4 + \uCE90\uB9AD\uD130 \uD3EC\uC778\uD2B8 \uCEEC\uB7EC. \uD074\uB9AC\uC5B4/\uD22C\uBA85 \uBCA0\uC774\uC2A4\uC5D0 \uC790\uCCB4 \uCE90\uB9AD\uD130 \uBC30\uCE58. \uB525\uD37C\uD50C, \uD654\uC774\uD2B8, \uC18C\uD504\uD2B8 \uD551\uD06C \uC8FC\uB825", 7380),
          ] }),
          new TableRow({ children: [
            bCellBold("\uBCA0\uC2A4\uD2B8\uC140\uB7EC \uAD6C\uC870", 2460, "F0F8FF"),
            bCell("\uD480\uCEE4\uBC84 \uD074\uB9AC\uC5B4\uCF00\uC774\uC2A4(\uC624\uB514\uB108\uB9AC\uB77C\uC774\uD504, \uAC8C\uC774\uBC0D\uB9C8\uC6B0\uC2A4 \uB4F1 \uD14C\uB9C8). \uCEE4\uC2A4\uD140\uCF00\uC774\uC2A4(1:1 \uC81C\uC791). Z\uD50C\uB9BD/\uC5D0\uC5B4\uD31F \uCF00\uC774\uC2A4\uB85C \uD655\uC7A5. SHUSH X \uB358\uD0C0\uC6B4 \uCF5C\uB77C\uBCF4 \uD504\uB85C\uC81D\uD2B8", 7380, "F0F8FF"),
          ] }),
          new TableRow({ children: [
            bCellBold("\uAD11\uACE0 \uCE74\uD53C \uB808\uD37C\uB7F0\uC2A4", 2460),
            bCellMulti([
              new TextRun({ text: "\"\uC190\uB05D\uC5D0\uC11C \uB290\uAEF4\uC9C0\uB294 \uC18C\uC18C\uD55C \uD589\uBCF5\" ", font: "Arial", size: 19, bold: true }),
              new TextRun({ text: "/ \"\uBC84\uAC70\uBA4D \uD480\uCEE4\uBC84 \uCF00\uC774\uC2A4 \uC2E0\uC0C1\" / \"\uC778\uC2A4\uD0C0 \uD6C4 \uC8FC\uBB38\uD3ED\uC8FC \uC778\uAE30 \uC0C1\uD488\". \uCE90\uB9AD\uD130 \uAC10\uC131 + \uC0AC\uD68C\uC801 \uC99D\uAC70 \uACB0\uD569 \uD328\uD134", font: "Arial", size: 19 }),
            ], 7380),
          ] }),
          new TableRow({ children: [
            bCellBold("\uC0AC\uC774\uD2B8 \uD1A4\uC564\uB9E4\uB108", 2460, "F0F8FF"),
            bCell("\uADC0\uC5EC\uC6B4 \uD1A4 + \uCE90\uB9AD\uD130 \uC911\uC2EC. \uC8FC\uBB38\uC81C\uC791 1~2\uC77C \uCD9C\uACE0. \uBB34\uC2E0\uC0AC \uC785\uC810\uC73C\uB85C \uC2E0\uB8B0\uAC10 \uD655\uBCF4. \uCF5C\uB77C\uBCF4 \uD504\uB85C\uC81D\uD2B8\uB85C \uD654\uC81C\uC131", 7380, "F0F8FF"),
          ] }),
        ]
      }),
      body("\uB808\uD37C\uB7F0\uC2A4 \uD3EC\uC778\uD2B8: \uC790\uCCB4 \uCE90\uB9AD\uD130 \uAE30\uBC18 \uD480\uCEE4\uBC84 + \uCF5C\uB77C\uBCF4 \uD654\uC81C\uC131 + \uC8FC\uBB38\uC81C\uC791 \uD504\uB9AC\uBBF8\uC5C4", { bold: true, color: "E04040" }),

      // --- 2-10. 몽몽드 ---
      heading2("2-10. \uBABD\uBABD\uB4DC (monmonde)"),
      new Paragraph({ spacing: { after: 60 }, children: [new TextRun({ text: "@monmonde.official | 2025 \uD3F0\uCF00\uC774\uC2A4 \uB9DB\uC9D1 TOP5 \uC120\uC815", size: 20, color: "888888", italics: true })] }),

      new Table({
        width: { size: 9840, type: WidthType.DXA },
        columnWidths: [2460, 7380],
        rows: [
          new TableRow({ children: [
            bCellBold("\uBE0C\uB79C\uB4DC DNA", 2460, "FFF5F0"),
            bCell("\"All about your moments\". \uC21C\uAC04\uC758 \uAC10\uC131\uC744 \uB2F4\uB294 \uBE0C\uB79C\uB4DC. \uC2E0\uC0DD \uBE0C\uB79C\uB4DC\uC784\uC5D0\uB3C4 2025 \uD3F0\uCF00\uC774\uC2A4 \uB9DB\uC9D1 \uCD94\uCC9C 5\uC120\uC5D0 \uC120\uC815\uB420 \uB9CC\uD07C \uC8FC\uBAA9\uBC1B\uB294 \uC911. \uC18C\uADDC\uBAA8\uC774\uC9C0\uB9CC \uAC10\uC131\uC758 \uBC00\uB3C4\uAC00 \uB192\uC740 \uBE0C\uB79C\uB4DC", 7380, "FFF5F0"),
          ] }),
          new TableRow({ children: [
            bCellBold("\uCEEC\uB7EC \uD314\uB808\uD2B8", 2460),
            bCell("\uBBF4\uB4DC \uCEEC\uB7EC \uC911\uC2EC. \uBD80\uB4DC\uB7EC\uC6B4 \uD30C\uC2A4\uD154 \uD1A4 + \uC218\uCC44\uD654 \uD130\uCE58. \uC790\uC5F0 \uBAA8\uD2F0\uD504\uC640 \uCD94\uC0C1\uC801 \uD328\uD134\uC758 \uACB0\uD569", 7380),
          ] }),
          new TableRow({ children: [
            bCellBold("\uBCA0\uC2A4\uD2B8\uC140\uB7EC \uAD6C\uC870", 2460, "FFF5F0"),
            bCell("\uC18C\uADDC\uBAA8 \uC790\uC0AC\uBAB0 \uC911\uC2EC. \uC2E0\uC0DD \uBE0C\uB79C\uB4DC\uC774\uC9C0\uB9CC \uC778\uD50C\uB8E8\uC5B8\uC11C \uCD94\uCC9C\uC744 \uD1B5\uD574 \uBE60\uB974\uAC8C \uC778\uC9C0\uB3C4 \uC0C1\uC2B9 \uC911. \uC2E0\uC0DD\uBE0C\uB79C\uB4DC\uAC00 \uD3F0\uCF00\uC774\uC2A4 \uB9DB\uC9D1\uC5D0 \uC120\uC815\uB418\uB294 \uACFC\uC815 \uC790\uCCB4\uAC00 \uB808\uD37C\uB7F0\uC2A4", 7380, "FFF5F0"),
          ] }),
          new TableRow({ children: [
            bCellBold("\uAD11\uACE0 \uCE74\uD53C \uB808\uD37C\uB7F0\uC2A4", 2460),
            bCellMulti([
              new TextRun({ text: "\"All about your moments\" ", font: "Arial", size: 19, bold: true }),
              new TextRun({ text: "/ \"\uC21C\uAC04\uC744 \uB2F4\uB294 \uCF00\uC774\uC2A4\" / \"\uD3F0\uCF00\uC774\uC2A4 \uB9DB\uC9D1 TOP5 \uC120\uC815\". \uC2E0\uC0DD \uBE0C\uB79C\uB4DC\uC758 \uBE60\uB978 \uC131\uC7A5 \uC2A4\uD1A0\uB9AC + \uBBF4\uB4DC \uAC10\uC131 \uCE74\uD53C", font: "Arial", size: 19 }),
            ], 7380),
          ] }),
          new TableRow({ children: [
            bCellBold("\uC0AC\uC774\uD2B8 \uD1A4\uC564\uB9E4\uB108", 2460, "FFF5F0"),
            bCell("\uBBF4\uB4DC \uAC10\uC131\uC758 \uBBF8\uB2C8\uBA40 \uC0AC\uC774\uD2B8. \uC18C\uADDC\uBAA8\uC774\uC9C0\uB9CC \uBE0C\uB79C\uB4DC \uC815\uCCB4\uC131\uC774 \uAC15\uD568. \uC2E0\uC0DD \uBE0C\uB79C\uB4DC\uAC00 \uC5B4\uB5BB\uAC8C \uAC10\uC131\uC73C\uB85C \uC8FC\uBAA9\uBC1B\uB294\uC9C0\uC758 \uC88B\uC740 \uC0AC\uB840", 7380, "FFF5F0"),
          ] }),
        ]
      }),
      body("\uB808\uD37C\uB7F0\uC2A4 \uD3EC\uC778\uD2B8: \uC2E0\uC0DD \uBE0C\uB79C\uB4DC \uC131\uC7A5 \uC804\uB7B5 + \uBBF4\uB4DC \uAC10\uC131 \uBC00\uB3C4 + \uC778\uD50C\uB8E8\uC5B8\uC11C \uCD94\uCC9C \uD65C\uC6A9\uBC95", { bold: true, color: "E04040" }),

      new Paragraph({ children: [new PageBreak()] }),

      // ===== 온세스튜디오 적용 섹션 =====
      heading1("3. \uC628\uC138\uC2A4\uD29C\uB514\uC624 \uBA54\uD0C0\uAD11\uACE0 \uC801\uC6A9 \uAC00\uC774\uB4DC"),
      body("\uC628\uC138\uC2A4\uD29C\uB514\uC624\uC758 \uD604\uC7AC \uAC15\uC810(\uBA85\uD654 \uC2DC\uB9AC\uC988, \uB9AC\uBCF8 \uD328\uD134, \uBC1C\uB808\uCF54\uC5B4 \uB4F1 \uC544\uD2B8 \uAC10\uC131)\uC5D0 \uC704 10\uAC1C \uBE0C\uB79C\uB4DC\uC758 \uC804\uB7B5\uC744 \uC811\uBAA9\uD558\uB294 \uBC29\uBC95\uC785\uB2C8\uB2E4:"),

      heading2("3-1. \uD604\uC7AC \uC628\uC138\uC2A4\uD29C\uB514\uC624 \uAC15\uC810 \uBD84\uC11D"),
      bulletBold("\uBA85\uD654 \uC2DC\uB9AC\uC988: ", "\uBAA8\uB124/\uD638\uC548\uBBF8\uB85C \uB4F1 \uBA85\uD654 \uBAA8\uD2F0\uD504 \u2192 \uC544\uC6B0\uB810\uCC98\uB7FC \uC2DC\uC801 \uC81C\uD488\uBA85\uC744 \uBD99\uC774\uBA74 \uAD11\uACE0 \uCE74\uD53C\uB85C \uBC14\uB85C \uC4F8 \uC218 \uC788\uC74C"),
      bulletBold("\uB9AC\uBCF8 \uD328\uD134 / \uBC1C\uB808\uCF54\uC5B4: ", "\uC138\uCEE8\uB4DC\uC720\uB2C8\uD06C\uB124\uC784\uCC98\uB7FC \uD328\uC158 \uC545\uC138\uC11C\uB9AC\uB85C \uD3EC\uC9C0\uC154\uB2DD \uAC00\uB2A5. \"\uBC1C\uB808\uCF54\uC5B4 \uCF00\uC774\uC2A4\" \uC790\uCCB4\uAC00 \uD0C0\uAC9F \uD0A4\uC6CC\uB4DC"),
      bulletBold("LOVE CONQUERS ALL \uC2AC\uB85C\uAC74: ", "\uD558\uC774\uC6B0\uCC98\uB7FC \uBE0C\uB79C\uB4DC \uCCA0\uD559\uC774 \uBA85\uD655. \uC774\uAC78 \uAD11\uACE0 \uCE74\uD53C \uC804\uBA74\uC5D0 \uD65C\uC6A9\uD558\uBA74 \uC815\uCCB4\uC131 \uAC15\uD654"),

      heading2("3-2. 10\uAC1C \uBE0C\uB79C\uB4DC\uC5D0\uC11C \uBC14\uB85C \uAC00\uC838\uC62C \uAC83"),
      new Table({
        width: { size: 9840, type: WidthType.DXA },
        columnWidths: [2460, 3690, 3690],
        rows: [
          new TableRow({ children: [ hCell("\uCC38\uACE0 \uBE0C\uB79C\uB4DC", 2460), hCell("\uAC00\uC838\uC62C \uC804\uB7B5", 3690), hCell("\uC628\uC138\uC2A4\uD29C\uB514\uC624 \uC801\uC6A9\uC548", 3690) ] }),
          new TableRow({ children: [
            bCellBold("\uC544\uC6B0\uB810", 2460, "FFF8F0"),
            bCell("\uC2DC\uC801 \uC81C\uD488\uBA85 + \uD14C\uB9C8\uBCC4 \uBB36\uC74C \uD310\uB9E4", 3690, "FFF8F0"),
            bCell("\uBA85\uD654 \uC2DC\uB9AC\uC988\uC5D0 \uAC10\uC131 \uC81C\uD488\uBA85 \uBD99\uC774\uAE30: \"\uBAA8\uB124\uC758 \uD55C\uB09F\", \"\uBBF8\uB85C\uC758 \uC0C1\uC0C1\" \uB4F1", 3690, "FFF8F0"),
          ] }),
          new TableRow({ children: [
            bCellBold("\uC18C\uC720\uB9C8\uC2E4", 2460),
            bCell("\uCEEC\uB7EC\uBA85 \uBE0C\uB79C\uB529 + \uC5D0\uD3ED\uC2DC \uBC94\uD37C", 3690),
            bCell("\uCEEC\uB7EC\uBCC4 \uAC10\uC131\uBA85 \uB9CC\uB4E4\uAE30: \"Monet Blue\", \"Mir\u00F3 Pink\" \uB4F1 \uBA85\uD654+\uCEEC\uB7EC \uACB0\uD569", 3690),
          ] }),
          new TableRow({ children: [
            bCellBold("\uAC00\uB974\uC1A1\uD2F0\uBBF8\uB4DC", 2460, "FFF8F0"),
            bCell("\uB79C\uB364\uBC15\uC2A4 \uBBF8\uB07C\uC0C1\uD488 + 2\uB2E8\uACC4 \uAC00\uACA9\uC81C", 3690, "FFF8F0"),
            bCell("\uBA85\uD654 \uB79C\uB364\uBC15\uC2A4 9,999\uC6D0 \u2192 \uC5D0\uD3ED\uC2DC/\uCE74\uB4DC\uD615 \uC5C5\uC140 \uC720\uB3C4", 3690, "FFF8F0"),
          ] }),
          new TableRow({ children: [
            bCellBold("\uC5B4\uD504\uC5B4\uD504", 2460),
            bCell("\uCF5C\uB77C\uBCF4/\uD55C\uC815\uD310 \uAE34\uBC15 \uB9C8\uCF00\uD305", 3690),
            bCell("\uC791\uAC00 \uCF5C\uB77C\uBCF4 \uD55C\uC815\uD310 \uAE30\uD68D: \"\uC628\uC138 X [OO] \uC791\uAC00 \uCF5C\uB77C\uBCF4\"", 3690),
          ] }),
          new TableRow({ children: [
            bCellBold("\uC138\uCEE8\uB4DC\uC720\uB2C8\uD06C\uB124\uC784", 2460, "FFF8F0"),
            bCell("\uD328\uC158 \uC545\uC138\uC11C\uB9AC \uD3EC\uC9C0\uC154\uB2DD", 3690, "FFF8F0"),
            bCell("\uBC1C\uB808\uCF54\uC5B4/\uB9AC\uBCF8 \uB77C\uC778\uC744 \"\uD328\uC158 \uCF00\uC774\uC2A4\"\uB85C \uD3EC\uC9C0\uC154\uB2DD", 3690, "FFF8F0"),
          ] }),
          new TableRow({ children: [
            bCellBold("\uB514\uC790\uC778\uC2A4\uD0A8", 2460),
            bCell("\uC138\uD2B8 \uAD6C\uC131 + \uD68C\uC6D0\uAC00 \uD560\uC778", 3690),
            bCell("\uBA85\uD654\uCF00\uC774\uC2A4+\uCE74\uB4DC\uD640\uB354+\uC2A4\uB9C8\uD2B8\uD1A1 \uC138\uD2B8 \uAD6C\uC131, \uD68C\uC6D0\uAC00 \uD560\uC778 \uC801\uC6A9", 3690),
          ] }),
        ]
      }),

      heading2("3-3. \uC628\uC138\uC2A4\uD29C\uB514\uC624 \uBA54\uD0C0\uAD11\uACE0 \uCE74\uD53C \uC608\uC2DC"),
      new Paragraph({ numbering: { reference: "nums", level: 0 }, spacing: { after: 100 }, children: [
        new TextRun({ text: "\uBA85\uD654 \uAC10\uC131 \uCEE8\uC149: ", bold: true, size: 22 }), new TextRun({ text: "\"\uBAA8\uB124\uC758 \uD55C\uB09F\uC744 \uC190\uC5D0 \uB2F4\uB2E4\" / \"\uBBF8\uC220\uAD00\uC5D0\uC11C \uAC78\uC5B4\uB098\uC628 \uCF00\uC774\uC2A4\" / \"LOVE CONQUERS ALL - \uBA85\uD654 \uC2DC\uB9AC\uC988 \uC2E0\uC0C1\"", size: 22 })
      ] }),
      new Paragraph({ numbering: { reference: "nums", level: 0 }, spacing: { after: 100 }, children: [
        new TextRun({ text: "\uBC1C\uB808\uCF54\uC5B4 \uCEE8\uC149: ", bold: true, size: 22 }), new TextRun({ text: "\"\uC5C7\uC740 \uB9AC\uBCF8 \uD328\uD134 \uBC1C\uB808\uCF54\uC5B4 \uCF00\uC774\uC2A4\" / \"\uC6D0\uC601\uC774\uC998 \uD551\uD06C \uB9AC\uBCF8\" / \"\uC624\uB298\uC758 OOTD\uC5D0 \uC5B4\uC6B8\uB9AC\uB294 \uCF00\uC774\uC2A4\"", size: 22 })
      ] }),
      new Paragraph({ numbering: { reference: "nums", level: 0 }, spacing: { after: 100 }, children: [
        new TextRun({ text: "\uC800\uAC00 \uC9C4\uC785 \uCEE8\uC149: ", bold: true, size: 22 }), new TextRun({ text: "\"\uBA85\uD654 \uB79C\uB364\uBC15\uC2A4 9,999\uC6D0\" / \"\uC5B4\uB5A4 \uBA85\uD654\uAC00 \uC62C\uC9C0 \uBAB0\uB77C\uC694\" / \"\uCCAB \uAD6C\uB9E4 \uD2B9\uAC00\"", size: 22 })
      ] }),
      new Paragraph({ numbering: { reference: "nums", level: 0 }, spacing: { after: 100 }, children: [
        new TextRun({ text: "\uC138\uD2B8 \uAD6C\uC131 \uCEE8\uC149: ", bold: true, size: 22 }), new TextRun({ text: "\"\uBA85\uD654\uCF00\uC774\uC2A4 + \uCE74\uB4DC\uD640\uB354 \uC138\uD2B8\" / \"\uBAA8\uB124 \uCF5C\uB809\uC158 3\uC885 SET\" / \"\uC27C\uC544\uC9C0\uB294 \uC608\uC220 \uC138\uD2B8\"", size: 22 })
      ] }),
      new Paragraph({ numbering: { reference: "nums", level: 0 }, spacing: { after: 100 }, children: [
        new TextRun({ text: "\uC2DC\uC98C\uAC10 \uCEE8\uC149: ", bold: true, size: 22 }), new TextRun({ text: "\"\uBD04\uC5D0 \uC5B4\uC6B8\uB9AC\uB294 \uBA85\uD654 \uCEEC\uB809\uC158\" / \"\uC5EC\uB984 \uD55C\uC815 \uD074\uB9AC\uC5B4 \uC2DC\uB9AC\uC988\" / \"\uAC00\uC744 \uB4E4\uD310 \uCF5C\uB809\uC158\"", size: 22 })
      ] }),

      new Paragraph({ children: [new PageBreak()] }),

      // ===== 4. 바로 적용할 수 있는 액션 플랜 =====
      heading1("4. \uBC14\uB85C \uC801\uC6A9\uD560 \uC218 \uC788\uB294 \uC561\uC158 \uD50C\uB79C"),

      heading2("4-1. \uB514\uC790\uC778 \uB808\uD37C\uB7F0\uC2A4 \uC801\uC6A9\uBC95"),
      new Paragraph({ numbering: { reference: "nums", level: 0 }, spacing: { after: 100 }, children: [
        new TextRun({ text: "\uCEEC\uB7EC\uBA85\uC744 \uBE0C\uB79C\uB529\uD558\uB77C: ", bold: true, size: 22 }), new TextRun({ text: "\uC18C\uC720\uB9C8\uC2E4\uCC98\uB7FC Haze Blue, Peach Cream \uAC19\uC740 \uAC10\uC131\uC801 \uCEEC\uB7EC\uBA85\uC744 \uB9CC\uB4E4\uBA74 \uCEEC\uB7EC\uBA85 \uC790\uCCB4\uAC00 \uAD11\uACE0 \uCE74\uD53C\uAC00 \uB429\uB2C8\uB2E4", size: 22 })
      ] }),
      new Paragraph({ numbering: { reference: "nums", level: 0 }, spacing: { after: 100 }, children: [
        new TextRun({ text: "\uC2DC\uB9AC\uC988\uB85C \uBB36\uC5B4\uB77C: ", bold: true, size: 22 }), new TextRun({ text: "\uC5B4\uD504\uC5B4\uD504\uC758 \uCE90\uB9AD\uD130 \uC2DC\uB9AC\uC988, \uAC00\uB974\uC1A1\uD2F0\uBBF8\uB4DC\uC758 \uAF2C\uC21C \uC2DC\uB9AC\uC988\uCC98\uB7FC \uB514\uC790\uC778\uC744 \uC2DC\uB9AC\uC988/\uCEEC\uB809\uC158\uC73C\uB85C \uBB36\uC73C\uBA74 \uCE90\uB7EC\uC140 \uAD11\uACE0 \uC18C\uC7AC\uAC00 \uB429\uB2C8\uB2E4", size: 22 })
      ] }),
      new Paragraph({ numbering: { reference: "nums", level: 0 }, spacing: { after: 100 }, children: [
        new TextRun({ text: "\uC81C\uD488\uBA85\uC5D0 \uAC10\uC131\uC744 \uB2F4\uC544\uB77C: ", bold: true, size: 22 }), new TextRun({ text: "\uC544\uC6B0\uB810\uC758 \"\uC783\uD600\uC9C4 \uAFC8\uC758 \uD754\uC801\" \"\uB2EC\uCF64\uD55C \uC0C9\uC758 \uAFC8\" \uAC19\uC740 \uC2DC\uC801 \uC81C\uD488\uBA85\uC740 \uAD11\uACE0 \uCE74\uD53C\uB85C \uADF8\uB300\uB85C \uC4F8 \uC218 \uC788\uC2B5\uB2C8\uB2E4", size: 22 })
      ] }),
      new Paragraph({ numbering: { reference: "nums", level: 0 }, spacing: { after: 100 }, children: [
        new TextRun({ text: "2\uB2E8\uACC4 \uAC00\uACA9\uC81C\uB97C \uD65C\uC6A9\uD558\uB77C: ", bold: true, size: 22 }), new TextRun({ text: "\uAC00\uB974\uC1A1\uD2F0\uBBF8\uB4DC\uC758 \uD22C\uBA85(20,000)/\uC5D0\uD3ED\uC2DC(25,000) \uAD6C\uC870\uCC98\uB7FC, \uAC19\uC740 \uB514\uC790\uC778\uC744 \uC18C\uC7AC\uBCC4\uB85C \uAC00\uACA9 \uCC28\uB4F1\uD654\uD558\uBA74 \uC5C5\uC140 \uC720\uB3C4\uAC00 \uC790\uC5F0\uC2A4\uB7FD\uC2B5\uB2C8\uB2E4", size: 22 })
      ] }),
      new Paragraph({ numbering: { reference: "nums", level: 0 }, spacing: { after: 100 }, children: [
        new TextRun({ text: "\uC800\uAC00 \uC9C4\uC785\uC0C1\uD488\uC744 \uB9CC\uB4E4\uC5B4\uB77C: ", bold: true, size: 22 }), new TextRun({ text: "\uC5B4\uD504\uC5B4\uD504 \uD0A4\uCEA1\uD0A4\uB9C1 7,900\uC6D0, \uAC00\uB974\uC1A1\uD2F0\uBBF8\uB4DC \uB79C\uB364\uBC15\uC2A4 9,999\uC6D0\uCC98\uB7FC \uC800\uAC00 \uBBF8\uB07C \uC0C1\uD488\uC744 \uBA54\uD0C0\uAD11\uACE0 \uC9C4\uC785\uC6A9\uC73C\uB85C \uC6B4\uC601\uD558\uBA74 \uC2E0\uADDC\uACE0\uAC1D \uD655\uBCF4\uC5D0 \uD6A8\uACFC\uC801", size: 22 })
      ] }),

      heading2("4-2. \uBA54\uD0C0\uAD11\uACE0 \uCE74\uD53C \uD15C\uD50C\uB9BF"),
      body("\uC704 \uBE0C\uB79C\uB4DC\uB4E4\uC758 \uAD11\uACE0 \uCE74\uD53C \uD328\uD134\uC744 \uC885\uD569\uD574\uC11C, \uBC14\uB85C \uC801\uC6A9\uD560 \uC218 \uC788\uB294 \uD15C\uD50C\uB9BF\uC785\uB2C8\uB2E4:"),

      new Table({
        width: { size: 9840, type: WidthType.DXA },
        columnWidths: [2460, 3690, 3690],
        rows: [
          new TableRow({ children: [ hCell("\uCE74\uD53C \uC720\uD615", 2460), hCell("\uD15C\uD50C\uB9BF \uC608\uC2DC", 3690), hCell("\uCC38\uACE0 \uBE0C\uB79C\uB4DC", 3690) ] }),
          new TableRow({ children: [
            bCellBold("\uC2E0\uC0C1 \uCD9C\uC2DC", 2460, "FFF8F0"),
            bCell("\"[\uCEEC\uB7EC\uBA85] \uC2E0\uC0C1 \uCD9C\uC2DC\"\n\"\uBD04\uC5D0 \uC5B4\uC6B8\uB9AC\uB294 [OO] \uCEEC\uB809\uC158\"", 3690, "FFF8F0"),
            bCell("\uC18C\uC720\uB9C8\uC2E4, \uC544\uC6B0\uB810", 3690, "FFF8F0"),
          ] }),
          new TableRow({ children: [
            bCellBold("\uD55C\uC815/\uCF5C\uB77C\uBCF4", 2460),
            bCell("\"[OO]X[OO] \uCF5C\uB77C\uBCF4 \uD55C\uC815\uD310\"\n\"\uAE30\uAC04\uD55C\uC815 \uC138\uC77C [N]% OFF\"", 3690),
            bCell("\uC5B4\uD504\uC5B4\uD504", 3690),
          ] }),
          new TableRow({ children: [
            bCellBold("\uCEE4\uC2A4\uD130\uB9C8\uC774\uC9D5", 2460, "FFF8F0"),
            bCell("\"\uB098\uB9CC\uC758 [OO] \uC870\uD569 \uB9CC\uB4E4\uAE30\"\n\"\uC2A4\uD2B8\uB7A9 \uBC14\uAFB8\uACE0 \uD328\uCE58 \uBD99\uC774\uACE0\"", 3690, "FFF8F0"),
            bCell("\uC138\uCEE8\uB4DC\uC720\uB2C8\uD06C\uB124\uC784", 3690, "FFF8F0"),
          ] }),
          new TableRow({ children: [
            bCellBold("\uAC10\uC131 \uBB38\uAD6C", 2460),
            bCell("\"\uC783\uD600\uC9C4 \uAFC8\uC758 \uD754\uC801\uC744 \uC190\uC5D0 \uB2F4\uB2E4\"\n\"\uC624\uB798 \uACE4\uC5D0 \uB450\uACE0 \uC2F6\uC740 \uB514\uC790\uC778\"", 3690),
            bCell("\uC544\uC6B0\uB810, \uD558\uC774\uC6B0", 3690),
          ] }),
          new TableRow({ children: [
            bCellBold("\uC0AC\uD68C\uC801 \uC99D\uAC70", 2460, "FFF8F0"),
            bCell("\"BEST! \uC8FC\uBB38\uC774 \uB9CE\uC544\uC694\"\n\"[N]K \uD314\uB85C\uC6CC\uAC00 \uC120\uD0DD\uD55C\"", 3690, "FFF8F0"),
            bCell("\uC18C\uC720\uB9C8\uC2E4", 3690, "FFF8F0"),
          ] }),
          new TableRow({ children: [
            bCellBold("\uC138\uD2B8 \uAD6C\uC131", 2460),
            bCell("\"\uD480\uCEE4\uBC84+\uCE74\uB4DC\uD3EC\uCF13 \uC138\uD2B8\"\n\"\uCF00\uC774\uC2A4+\uC2A4\uB9C8\uD2B8\uD1A1+\uD0A4\uB9C1 \uC138\uD2B8\"", 3690),
            bCell("\uB514\uC790\uC778\uC2A4\uD0A8", 3690),
          ] }),
          new TableRow({ children: [
            bCellBold("\uC800\uAC00 \uC9C4\uC785", 2460, "FFF8F0"),
            bCell("\"\uB79C\uB364 \uB514\uC790\uC778 9,999\uC6D0\"\n\"\uD0A4\uCEA1\uD0A4\uB9C1 7,900\uC6D0\uBD80\uD130\"", 3690, "FFF8F0"),
            bCell("\uAC00\uB974\uC1A1\uD2F0\uBBF8\uB4DC, \uC5B4\uD504\uC5B4\uD504", 3690, "FFF8F0"),
          ] }),
        ]
      }),

      heading2("4-3. \uC0C8 \uBE0C\uB79C\uB4DC/\uB77C\uC778 \uB9CC\uB4E4 \uB54C \uCCB4\uD06C\uB9AC\uC2A4\uD2B8"),
      new Paragraph({ numbering: { reference: "nums2", level: 0 }, spacing: { after: 80 }, children: [
        new TextRun({ text: "\uBE0C\uB79C\uB4DC \uCEEC\uB7EC \uC2DC\uC2A4\uD15C \uC815\uD588\uB294\uAC00? ", bold: true, size: 22 }), new TextRun({ text: "(\uD558\uC774\uC6B0: \uD578\uD551\uD06C+\uBE14\uB799 / \uC5B4\uD504\uC5B4\uD504: \uD551\uD06C+\uD654\uC774\uD2B8)", size: 22 })
      ] }),
      new Paragraph({ numbering: { reference: "nums2", level: 0 }, spacing: { after: 80 }, children: [
        new TextRun({ text: "\uCEEC\uB7EC\uBA85\uC5D0 \uAC10\uC131\uC774 \uB2F4\uACA8 \uC788\uB294\uAC00? ", bold: true, size: 22 }), new TextRun({ text: "(\uC18C\uC720\uB9C8\uC2E4: Haze Blue / \uC544\uC6B0\uB810: \uC2EC\uC5F0\uC758 \uC0AC\uB9C9)", size: 22 })
      ] }),
      new Paragraph({ numbering: { reference: "nums2", level: 0 }, spacing: { after: 80 }, children: [
        new TextRun({ text: "\uC2DC\uB9AC\uC988/\uCEEC\uB809\uC158 \uAD6C\uC131\uC774 \uC788\uB294\uAC00? ", bold: true, size: 22 }), new TextRun({ text: "(\uC5B4\uD504\uC5B4\uD504: 8\uC885 \uCE90\uB9AD\uD130 / \uAC00\uB974\uC1A1\uD2F0\uBBF8\uB4DC: \uD14C\uB9C8\uBCC4 \uC2DC\uB9AC\uC988)", size: 22 })
      ] }),
      new Paragraph({ numbering: { reference: "nums2", level: 0 }, spacing: { after: 80 }, children: [
        new TextRun({ text: "\uC800\uAC00 \uC9C4\uC785\uC0C1\uD488\uC774 \uC788\uB294\uAC00? ", bold: true, size: 22 }), new TextRun({ text: "(1\uB9CC\uC6D0 \uC774\uD558 \uBBF8\uB07C\uC0C1\uD488\uC73C\uB85C \uC2E0\uADDC\uACE0\uAC1D \uC720\uC785)", size: 22 })
      ] }),
      new Paragraph({ numbering: { reference: "nums2", level: 0 }, spacing: { after: 80 }, children: [
        new TextRun({ text: "\uC5C5\uC140 \uAD6C\uC870\uAC00 \uC788\uB294\uAC00? ", bold: true, size: 22 }), new TextRun({ text: "(\uAC00\uB974\uC1A1\uD2F0\uBBF8\uB4DC: \uD22C\uBA85\u2192\uC5D0\uD3ED\uC2DC / \uB514\uC790\uC778\uC2A4\uD0A8: \uB2E8\uD488\u2192\uC138\uD2B8)", size: 22 })
      ] }),
      new Paragraph({ numbering: { reference: "nums2", level: 0 }, spacing: { after: 80 }, children: [
        new TextRun({ text: "\uD310\uB9E4\uCC44\uB110 \uC804\uB7B5\uC774 \uC788\uB294\uAC00? ", bold: true, size: 22 }), new TextRun({ text: "(\uC790\uC0AC\uBAB0+\uBB34\uC2E0\uC0AC/29CM/\uC9C0\uADF8\uC7AC\uADF8 \uC785\uC810 \uBCD1\uD589)", size: 22 })
      ] }),
      new Paragraph({ numbering: { reference: "nums2", level: 0 }, spacing: { after: 80 }, children: [
        new TextRun({ text: "\uC0AC\uC774\uD2B8 \uD1A4\uC564\uB9E4\uB108\uAC00 \uC77C\uAD00\uC801\uC778\uAC00? ", bold: true, size: 22 }), new TextRun({ text: "(\uBC30\uACBD\uC0C9+\uD3F0\uD2B8+\uCEEC\uB7EC \uC2DC\uC2A4\uD15C = \uBE0C\uB79C\uB4DC \uC815\uCCB4\uC131)", size: 22 })
      ] }),

      new Paragraph({ spacing: { before: 400 }, border: { top: { style: BorderStyle.SINGLE, size: 3, color: "CCCCCC", space: 1 } }, children: [] }),
      new Paragraph({ spacing: { before: 100 }, children: [new TextRun({ text: "\uCC38\uACE0: \uAC01 \uBE0C\uB79C\uB4DC \uACF5\uC2DD \uC0AC\uC774\uD2B8 + \uBB34\uC2E0\uC0AC/29CM/W\uCEE8\uC149 \uC785\uC810 \uD398\uC774\uC9C0 + \uC778\uC2A4\uD0C0\uADF8\uB7A8 \uACF5\uC2DD \uACC4\uC815 \uAE30\uBC18 \uBD84\uC11D", size: 18, color: "999999", italics: true })] }),
      new Paragraph({ children: [new TextRun({ text: "\uC791\uC131\uC77C: 2026\uB144 4\uC6D4 14\uC77C", size: 18, color: "999999", italics: true })] }),
    ]
  }]
});

Packer.toBuffer(doc).then(buffer => {
  fs.writeFileSync("/sessions/zealous-confident-pasteur/mnt/outputs/폰케이스_감성브랜드_레퍼런스가이드.docx", buffer);
  console.log("Reference guide created successfully!");
});
