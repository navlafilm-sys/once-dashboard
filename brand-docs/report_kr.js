const fs = require('fs');
const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
        Header, Footer, AlignmentType, LevelFormat,
        HeadingLevel, BorderStyle, WidthType, ShadingType,
        PageNumber, PageBreak } = require('docx');

const border = { style: BorderStyle.SINGLE, size: 1, color: "DDDDDD" };
const borders = { top: border, bottom: border, left: border, right: border };
const cellMargins = { top: 80, bottom: 80, left: 120, right: 120 };

function hCell(text, width) {
  return new TableCell({
    borders, width: { size: width, type: WidthType.DXA },
    shading: { fill: "2D2D2D", type: ShadingType.CLEAR },
    margins: cellMargins, verticalAlign: "center",
    children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text, bold: true, color: "FFFFFF", font: "Arial", size: 20 })] })]
  });
}
function bCell(text, width, fill) {
  return new TableCell({
    borders, width: { size: width, type: WidthType.DXA },
    shading: fill ? { fill, type: ShadingType.CLEAR } : undefined,
    margins: cellMargins,
    children: [new Paragraph({ children: [new TextRun({ text, font: "Arial", size: 20 })] })]
  });
}

const doc = new Document({
  styles: {
    default: { document: { run: { font: "Arial", size: 22 } } },
    paragraphStyles: [
      { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 34, bold: true, font: "Arial", color: "1B1B1B" },
        paragraph: { spacing: { before: 360, after: 200 }, outlineLevel: 0 } },
      { id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 28, bold: true, font: "Arial", color: "333333" },
        paragraph: { spacing: { before: 240, after: 160 }, outlineLevel: 1 } },
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
      page: {
        size: { width: 12240, height: 15840 },
        margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 }
      }
    },
    headers: { default: new Header({ children: [new Paragraph({ alignment: AlignmentType.RIGHT, children: [new TextRun({ text: "\uD55C\uAD6D \uC18C\uB9E4 \uD3F0\uCF00\uC774\uC2A4 \uBA54\uD0C0\uAD11\uACE0 \uB514\uC790\uC778 \uD2B8\uB80C\uB4DC", italics: true, color: "999999", size: 18 })] })] }) },
    footers: { default: new Footer({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "Page ", size: 18, color: "999999" }), new TextRun({ children: [PageNumber.CURRENT], size: 18, color: "999999" })] })] }) },
    children: [
      // ===== COVER =====
      new Paragraph({ spacing: { before: 2000 }, children: [] }),
      new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 100 }, children: [new TextRun({ text: "\uD55C\uAD6D \uC18C\uB9E4 \uD3F0\uCF00\uC774\uC2A4", size: 48, bold: true, color: "1B1B1B" })] }),
      new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 100 }, children: [new TextRun({ text: "\uBA54\uD0C0\uAD11\uACE0 \uC798 \uD314\uB9AC\uB294 \uB514\uC790\uC778 \uD2B8\uB80C\uB4DC", size: 40, bold: true, color: "E04040" })] }),
      new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 300 }, border: { top: { style: BorderStyle.SINGLE, size: 4, color: "E04040", space: 1 } }, children: [] }),
      new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 80 }, children: [new TextRun({ text: "\uB77C\uC774\uC120\uC2A4 \uCE90\uB9AD\uD130 / \uCEE4\uC2A4\uD140 \uC81C\uC791 / \uC911\uAD6D \uC0AC\uC785 \uC81C\uC678", size: 22, color: "888888" })] }),
      new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 80 }, children: [new TextRun({ text: "\uC790\uCCB4 \uB514\uC790\uC778 \uC18C\uB9E4 \uBE0C\uB79C\uB4DC \uAE30\uC900 | 2026\uB144 4\uC6D4", size: 20, color: "AAAAAA" })] }),
      new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "Meta Ad Library + \uD55C\uAD6D \uC18C\uB9E4 \uC1FC\uD551\uBAB0 + \uC5C5\uACC4 \uB9AC\uC11C\uCE58 \uC885\uD569", size: 20, color: "AAAAAA" })] }),

      new Paragraph({ children: [new PageBreak()] }),

      // ===== 1. \uD55C\uAD6D \uBA54\uD0C0\uAD11\uACE0 \uD3F0\uCF00\uC774\uC2A4 \uC2DC\uC7A5 \uD604\uD669 =====
      new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun("1. \uD55C\uAD6D \uBA54\uD0C0\uAD11\uACE0 \uD3F0\uCF00\uC774\uC2A4 \uC2DC\uC7A5 \uD604\uD669")] }),
      new Paragraph({ spacing: { after: 120 }, children: [
        new TextRun({ text: "Meta Ad Library \uAE30\uC900 \uD55C\uAD6D \uD0C0\uAC9F \"\uD3F0\uCF00\uC774\uC2A4\" \uD65C\uC131 \uAD11\uACE0\uB294 \uC57D 350\uAC1C\uB85C, \uD3F0\uCF00\uC774\uC2A4\uB294 10\uB300~30\uB300 \uC5EC\uC131\uC774 \uC8FC \uD0C0\uAC9F\uC778 \uAC10\uC131 \uC18C\uBE44 \uC81C\uD488\uC785\uB2C8\uB2E4. \uACE0\uAC1D\uC740 \uAE30\uB2A5\uC774\uB098 \uAC00\uACA9\uBCF4\uB2E4 '\uB0B4\uAC00 \uC88B\uC544\uD558\uB294 \uB514\uC790\uC778\uC778\uAC00?'\uB85C \uAD6C\uB9E4\uB97C \uACB0\uC815\uD569\uB2C8\uB2E4.", size: 22 })
      ] }),
      new Paragraph({ spacing: { after: 120 }, children: [
        new TextRun({ text: "\uC131\uACF5 \uC0AC\uB840: MBTI \uD3F0\uCF00\uC774\uC2A4 + \uAC10\uC131 \uD0A4\uB9C1\uC744 \uC2DC\uC98C \uAC10\uC131 \uC18C\uC7AC\uB85C \uAE30\uD68D\uD558\uC5EC \uC778\uC2A4\uD0C0 DA \uC6B4\uC601 \u2192 ROAS 620% \uB2EC\uC131. \uC2A4\uD0C0\uC77C\uBCC4 \uAE30\uD68D\uC804 \uD615\uD0DC\uB85C \uAD6C\uC131\uD55C \uC140\uB7EC\uB294 \uD68C\uC218\uC728 470% \uC774\uC0C1, \uBC29\uBB38\uC790 2.3\uBC30 \uC99D\uAC00 \uAE30\uB85D.", size: 22 })
      ] }),
      new Paragraph({ spacing: { after: 200 }, children: [
        new TextRun({ text: "\uD575\uC2EC: \uD3F0\uCF00\uC774\uC2A4\uB294 '\uB514\uC790\uC778\uC5D0 \uB9E4\uB825\uC744 \uB290\uAEF4 \uAD6C\uB9E4\uD558\uB294' \uAC10\uC131 \uC18C\uBE44 \uC81C\uD488\uC774\uBBC0\uB85C, \uAD11\uACE0 \uC18C\uC7AC \uC790\uCCB4\uAC00 \uACE7 \uC804\uD658\uC728\uC744 \uACB0\uC815\uD569\uB2C8\uB2E4.", size: 22, bold: true })
      ] }),

      // ===== 2. \uBA54\uD0C0\uAD11\uACE0\uC5D0\uC11C \uC798 \uD314\uB9AC\uB294 \uB514\uC790\uC778 \uC720\uD615 TOP 8 =====
      new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun("2. \uBA54\uD0C0\uAD11\uACE0\uC5D0\uC11C \uC798 \uD314\uB9AC\uB294 \uB514\uC790\uC778 \uC720\uD615 TOP 8")] }),
      new Paragraph({ spacing: { after: 160 }, children: [
        new TextRun({ text: "Meta Ad Library \uD55C\uAD6D \uD3F0\uCF00\uC774\uC2A4 \uAD11\uACE0 350\uAC1C + \uD55C\uAD6D \uC18C\uB9E4 \uC1FC\uD551\uBAB0 \uBCA0\uC2A4\uD2B8\uC140\uB7EC + \uC5C5\uACC4 \uB9AC\uC11C\uCE58\uB97C \uC885\uD569 \uBD84\uC11D\uD55C \uACB0\uACFC\uC785\uB2C8\uB2E4. \uB77C\uC774\uC120\uC2A4 \uCE90\uB9AD\uD130, \uCEE4\uC2A4\uD140 \uC81C\uC791, \uC911\uAD6D \uC0AC\uC785 \uC81C\uD488\uC740 \uC81C\uC678\uD588\uC2B5\uB2C8\uB2E4.", size: 22 })
      ] }),

      new Table({
        width: { size: 9360, type: WidthType.DXA },
        columnWidths: [500, 1800, 3530, 3530],
        rows: [
          new TableRow({ children: [ hCell("#", 500), hCell("\uB514\uC790\uC778 \uC720\uD615", 1800), hCell("\uC124\uBA85 & \uC608\uC2DC", 3530), hCell("\uAD11\uACE0 \uD3EC\uC778\uD2B8", 3530) ] }),
          new TableRow({ children: [
            bCell("1", 500, "FFF8F0"), bCell("\uCEEC\uB7EC \uD22C\uBA85 \uCF00\uC774\uC2A4", 1800, "FFF8F0"),
            bCell("변색 없는 투명 케이스 + 비비드/파스텔 컬러 프레임. 소유마실·하우위 등 감성 브랜드가 자체 컬러로 차별화. 투명 케이스 검색량 1분기 상승, 유튜브 리뷰 평균 12만뷰", 3530, "FFF8F0"),
            bCell("\"\uBCC0\uC0C9 \uC5C6\uB294 \uD22C\uBA85\uCF00\uC774\uC2A4\" \uAC15\uC870 + \uD3F4\uB9AC\uCE74\uBCF4\uB124\uC774\uD2B8 & \uD569\uC131 \uC18C\uC7AC \uCC28\uBCC4\uD654. \uC2E4\uC81C \uC190\uC5D0 \uB4E4\uACE0 \uC788\uB294 UGC \uC2A4\uD0C0\uC77C \uC601\uC0C1\uC774 \uC804\uD658\uC5D0 \uD6A8\uACFC\uC801", 3530, "FFF8F0"),
          ] }),
          new TableRow({ children: [
            bCell("2", 500), bCell("\uCCB4\uD06C & \uACA9\uC790 \uD328\uD134", 1800),
            bCell("\uD074\uB798\uC2DD \uD751\uBC31 \uCCB4\uD06C\uBD80\uD130 \uBB3C\uACB0 \uBCC0\uD615, \uD558\uD2B8 \uCCB4\uD06C, \uCEEC\uB7EC \uACA9\uC790\uAE4C\uC9C0. \uB354\uB098\uC778\uBAB0\uC758 \uCEEC\uB7EC \uD3EC\uC778\uD2B8 \uCCB4\uD06C \uC2E0\uC0C1\uC774 \uBA54\uD0C0\uAD11\uACE0\uC5D0\uC11C \uC9D1\uD589\uB428. 2025\uB144\uBD80\uD130 \uAFB8\uC900\uD55C \uC778\uAE30", 3530),
            bCell("\uCCB4\uD06C \uD328\uD134\uC740 \uC5B4\uB5A4 \uC637\uACFC\uB3C4 \uC798 \uC5B4\uC6B8\uB9AC\uB294 \uC810 \uAC15\uC870. \uCEEC\uB7EC \uBC30\uB9AC\uC5D0\uC774\uC158\uC73C\uB85C \uCE90\uB7EC\uC140 \uAD11\uACE0 \uC81C\uC791\uD558\uBA74 \uB2E4\uC591\uD55C \uCDE8\uD5A5 \uD0C0\uAC9F \uAC00\uB2A5", 3530),
          ] }),
          new TableRow({ children: [
            bCell("3", 500, "FFF8F0"), bCell("\uBE48\uD2F0\uC9C0 \uD50C\uB85C\uB7F4", 1800, "FFF8F0"),
            bCell("은은한 톤의 빈티지 꽃무늬 디자인. 아우렐의 감성 일러스트 플로럴, 소유마실의 자연 모티프 등. 수채화 터치 + 금선 아웃라인이 고급스러움 극대화", 3530, "FFF8F0"),
            bCell("\uBD04/\uC5EC\uB984/\uAC00\uC744 \uC2DC\uC98C \uAC10\uC131\uC73C\uB85C \uAD11\uACE0 \uC18C\uC7AC \uAD50\uCCB4. \"\uC0C8 \uACC4\uC808 \uC0C8 \uCF00\uC774\uC2A4\" \uD6C5\uC73C\uB85C \uC2DC\uC98C\uB9C8\uB2E4 \uC7AC\uAD6C\uB9E4 \uC720\uB3C4 \uAC00\uB2A5", 3530, "FFF8F0"),
          ] }),
          new TableRow({ children: [
            bCell("4", 500), bCell("\uBBF8\uB7EC / \uC2E4\uBC84 \uCF00\uC774\uC2A4", 1800),
            bCell("\uAC70\uC6B8\uCC98\uB7FC \uBC18\uC0AC\uB418\uB294 \uBBF8\uB7EC \uCF00\uC774\uC2A4, \uC2E4\uBC84/\uD06C\uB86C \uBA54\uD0C8\uB9AD \uCF00\uC774\uC2A4. \uB208\uCF54\u2661 \uBE0C\uB79C\uB4DC\uAC00 \uC2E4\uBC84 \uC5D0\uD3ED\uC2DC \uB514\uC790\uC778\uC73C\uB85C \uBA54\uD0C0\uAD11\uACE0 \uC9D1\uD589 \uC911. \uC140\uD53C \uBBF8\uB7EC \uD6A8\uACFC\uB85C MZ\uC138\uB300 \uC778\uAE30", 3530),
            bCell("\uC778\uC2A4\uD0C0 \uB9B4\uC2A4\uC5D0\uC11C \uC140\uD53C \uC7AC\uC2E0\uAC19\uC740 \uBBF8\uB7EC \uD6A8\uACFC \uBCF4\uC5EC\uC8FC\uB294 \uC21F\uD3FC \uAD11\uACE0\uAC00 \uD074\uB9AD\uC728 \uB192\uC74C. \uD2B9\uD788 \uC5EC\uC131 \uD0C0\uAC9F\uC5D0 \uAC15\uB825\uD55C \uC5B4\uD544", 3530),
          ] }),
          new TableRow({ children: [
            bCell("5", 500, "FFF8F0"), bCell("\uBBF4\uB4DC / \uAC10\uC131 \uBB38\uAD6C", 1800, "FFF8F0"),
            bCell("\"I'm Fine\", MBTI \uD3F0\uCF00\uC774\uC2A4, \uC9E7\uC740 \uC601\uBB38 \uBB38\uAD6C, \uD55C\uAE00 \uAC10\uC131 \uBB38\uAD6C \uB4F1. \uACF5\uAC10\uD615 \uBB38\uAD6C\uAC00 SNS \uACF5\uC720\uB97C \uC720\uBC1C\uD558\uBA70, \uD3F0\uCF00\uC774\uC2A4 \uC5C5\uC885\uC5D0\uC11C ROAS 620%\uB97C \uB2EC\uC131\uD55C \uC0AC\uB840\uC758 \uD575\uC2EC \uD0A4\uC6CC\uB4DC\uAC00 'MBTI \uD3F0\uCF00\uC774\uC2A4'\uC600\uC74C", 3530, "FFF8F0"),
            bCell("\"\uB098\uB97C \uD45C\uD604\uD558\uB294 \uCF00\uC774\uC2A4\" \uCEE8\uC149 \uAD11\uACE0. \uD0C0\uAC9F\uC758 \uAD00\uC2EC\uC0AC/\uC131\uACA9\uACFC \uC5F0\uACB0\uB41C \uBB38\uAD6C\uAC00 \uC804\uD658\uC744 \uB192\uC784. \uCE90\uB7EC\uC140 \uD615\uD0DC\uB85C MBTI\uBCC4 \uCD94\uCC9C", 3530, "FFF8F0"),
          ] }),
          new TableRow({ children: [
            bCell("6", 500), bCell("\uBBF4\uB4DC \uCEEC\uB7EC \uB2E8\uC0C9", 1800),
            bCell("\uB9E4\uD2B8 \uB2E8\uC0C9 \uCF00\uC774\uC2A4(Sand, Clay, Sage Green, Lavender, Dusty Rose). \uB2E8\uC0C9 30% \uC774\uC0C1 \uC18C\uBE44\uC790 \uC120\uD638\uB3C4 \uAE30\uB85D. \uAD11\uD0DD\uBCF4\uB2E4 \uBB34\uAD11 \uD14D\uC2A4\uCC98\uAC00 \uC555\uB3C4\uC801 \uC120\uD638. \uBE0C\uB9AC\uC988\uD53C\uAC00 \uCEEC\uB7EC \uBC14\uB9AC\uC5D0\uC774\uC158\uC73C\uB85C \uBA54\uD0C0\uAD11\uACE0 \uC9D1\uD589", 3530),
            bCell("\"\uB098\uB9CC\uC758 \uD3EC\uC778\uD2B8 \uCEEC\uB7EC\" \uCEE8\uC149. \uCEEC\uB7EC\uBCC4\uB85C \uBD84\uB9AC\uD55C \uCE90\uB7EC\uC140 \uAD11\uACE0\uAC00 \uD6A8\uACFC\uC801. \uACC4\uC808\uAC10\uACFC \uC5F0\uACC4\uD558\uBA74 \uB354 \uC88B\uC74C(\uBD04=\uD30C\uC2A4\uD154, \uAC00\uC744=\uB525\uD1A4)", 3530),
          ] }),
          new TableRow({ children: [
            bCell("7", 500, "FFF8F0"), bCell("\uC5D0\uD3ED\uC2DC / 3D \uC7A5\uC2DD", 1800, "FFF8F0"),
            bCell("\uC5D0\uD3ED\uC2DC \uC2A4\uD2F0\uCEE4, 3D \uC7A5\uC2DD, \uBE44\uC988 \uC7A5\uC2DD \uB4F1 \uC785\uCCB4\uC801 \uC7A5\uC2DD \uC694\uC18C. \uB208\uCF54\u2661 \uC2E4\uBC84 \uC5D0\uD3ED\uC2DC, \uD3F0\uBF40 \uC2A4\uB9C8\uD2B8\uD1A1 ZIP \uB4F1\uC774 \uBA54\uD0C0\uAD11\uACE0\uC5D0\uC11C \uD65C\uBC1C\uD558\uAC8C \uC9D1\uD589. \uD3F0\uAFB8 \uBB38\uD654\uC640 \uC5F0\uACB0", 3530, "FFF8F0"),
            bCell("\"\uD3F0\uAFB8\" \uD0A4\uC6CC\uB4DC\uB85C \uCF58\uD150\uCE20 \uC81C\uC791. \uCF00\uC774\uC2A4+\uC2A4\uB9C8\uD2B8\uD1A1+\uD0A4\uB9C1 \uC138\uD2B8 \uAD6C\uC131\uC73C\uB85C \uAC1D\uB2E8\uAC00 \uC0C1\uC2B9. \uB9B4\uC2A4 \uC601\uC0C1\uC73C\uB85C \uAFB8\uBBF8\uB294 \uACFC\uC815 \uBCF4\uC5EC\uC8FC\uAE30", 3530, "FFF8F0"),
          ] }),
          new TableRow({ children: [
            bCell("8", 500), bCell("\uB808\uD2B8\uB85C / Y2K", 1800),
            bCell("\uD22C\uBA85 \uD2F4\uD2F0\uB4DC \uCF00\uC774\uC2A4(\uBC84\uBE14\uAC80 \uD551\uD06C, \uB77C\uC784 \uADF8\uB9B0), \uB808\uD2B8\uB85C \uD3F0\uD2B8, \uAE43\uD568 \uD328\uD134 \uBCC0\uD615. 2000\uB144\uB300 \uB178\uC2A4\uD0E4\uC9C0\uC544 \uD65C\uC6A9. \uD3F0 \uB0B4\uBD80\uAC00 \uBE44\uCE58\uB294 \uD22C\uBA85 \uCEEC\uB7EC \uCF00\uC774\uC2A4\uAC00 MZ\uC138\uB300\uC5D0\uAC8C \uC778\uAE30", 3530),
            bCell("\"\uC637\uACFC \uB9E4\uCE6D\uD558\uB294 \uD3F0\uCF00\uC774\uC2A4\" \uCEE8\uC149. \uD328\uC158 \uCF54\uB514 + \uCF00\uC774\uC2A4 \uB9E4\uCE6D \uC774\uBBF8\uC9C0\uAC00 \uC778\uC2A4\uD0C0\uC5D0\uC11C \uC800\uC7A5/\uACF5\uC720 \uC720\uB3C4\uC5D0 \uD6A8\uACFC\uC801", 3530),
          ] }),
        ]
      }),

      new Paragraph({ children: [new PageBreak()] }),

      // ===== 3. 감성 스몰브랜드 =====
      new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun("3. 주목할 감성 스몰브랜드 (자체 디자인 기반)")] }),
      new Paragraph({ spacing: { after: 160 }, children: [
        new TextRun({ text: "인스타그램/메타 광고에서 자체 디자인으로 브랜딩에 성공한 한국 감성 스몰브랜드입니다. 대형 쇼핑몰·사입·라이선스 캐릭터 브랜드를 제외하고, 독자적 디자인 아이덴티티를 가진 브랜드만 선별했습니다.", size: 22 })
      ] }),

      new Table({
        width: { size: 9360, type: WidthType.DXA },
        columnWidths: [1600, 2200, 2780, 2780],
        rows: [
          new TableRow({ children: [ hCell("브랜드", 1600), hCell("디자인 정체성", 2200), hCell("대표 스타일 & 제품", 2780), hCell("참고 포인트", 2780) ] }),
          new TableRow({ children: [
            bCell("어프어프\n(EARPEARP)", 1600, "F5F9FF"),
            bCell("키치 + 유니크 컬러감. 자체 캐릭터 '코비' 중심의 팝한 감성", 2200, "F5F9FF"),
            bCell("웨이브 라벨 케이스(29,000원), 에폭시(28,000원), 하드(19,000원). 에어팟·아이패드 파우치까지 라이프스타일 확장", 2780, "F5F9FF"),
            bCell("인스타 @earp_earp 59K 팔로워. 무신사·지그재그 입점. 키치한 컬러 조합이 10~20대에 강한 소구력", 2780, "F5F9FF"),
          ] }),
          new TableRow({ children: [
            bCell("세컨드유니크네임\n(SUN)", 1600),
            bCell("컬러블록 + 스트랩/패치. 위트 있는 패션 악세서리 감성", 2200),
            bCell("SUN CASE 라인 - 비비드 컬러 + 스트랩·패치로 이미지 변형. 가격대 26,000~30,000원", 2780),
            bCell("무신사·29CM·W컨셉 입점. @youngboyz_sun 18K. 폰케이스를 '패션 아이템'으로 포지셔닝한 대표 사례", 2780),
          ] }),
          new TableRow({ children: [
            bCell("소유마실\n(SOYOUMASIL)", 1600, "F5F9FF"),
            bCell("자연에서 영감받은 에코 감성. 부드러운 톤과 텍스처", 2200, "F5F9FF"),
            bCell("에폭시범퍼 케이스(Blue haze plaid 등), 투명젤하드(Heart bean). 맥세이프 터프범퍼. 18,000~27,000원", 2780, "F5F9FF"),
            bCell("무신사·W컨셉 입점. 에폭시 소재 + 자연 컬러 조합이 20~30대 여성에 인기. 시즌 감성 컬렉션 운영", 2780, "F5F9FF"),
          ] }),
          new TableRow({ children: [
            bCell("하우위\n(howie)", 1600),
            bCell("미니멀 + 무드 컬러. 복잡하지 않은 자연스러움 추구", 2200),
            bCell("포인트 컬러 케이스, 아이패드 파우치, 그립톡. 직관적 디자인으로 일상 조화 강조", 2780),
            bCell("howie.co.kr 자사몰 운영. '편안하고 조화로운 분위기'를 브랜드 철학으로. 미니멀 무드 좋아하는 2030 타겟", 2780),
          ] }),
          new TableRow({ children: [
            bCell("가르송티미드\n(GARCONTIMIDE)", 1600, "F5F9FF"),
            bCell("감성 두들 아트. 심플한 일러스트에 따뜻한 유머", 2200, "F5F9FF"),
            bCell("왕눈이 캐릭터, 무지 컬러(베이커리 브라운, 스카이블루 등), 아티스트 블랙 투명 케이스. 20,000~22,000원", 2780, "F5F9FF"),
            bCell("무신사·29CM 입점. \"귀여운 것이 세상을 구한다\" 슬로건. 자체 캐릭터로 브랜딩 성공 사례", 2780, "F5F9FF"),
          ] }),
          new TableRow({ children: [
            bCell("아우렐\n(Aurel)", 1600),
            bCell("프리미엄 일러스트 감성. 우아함 + 몽환적 무드", 2200),
            bCell("감성 일러스트 젤하드(잊혀진 꿈의 흔적 등), 바다 물결 시리즈, 데칼코마니 블럭. 자체제작 디자인", 2780),
            bCell("aurel.kr 자사몰 + 지그재그 입점. 일러스트레이터 협업 느낌의 아트워크가 차별점. 선물용 수요 높음", 2780),
          ] }),
          new TableRow({ children: [
            bCell("하이우\n(hioo)", 1600, "F5F9FF"),
            bCell("오래 곁에 두고 싶은 디자인. 절제된 감성", 2200, "F5F9FF"),
            bCell("미니멀 라인업 중심. 과하지 않은 컬러와 형태. hioo.kr에서 직접 판매", 2780, "F5F9FF"),
            bCell("2025 폰케이스 맛집 추천 5선 선정(@ahyunfrom). 절제미가 브랜드 정체성. 재구매율 높은 충성 고객층", 2780, "F5F9FF"),
          ] }),
          new TableRow({ children: [
            bCell("디자인스킨\n(DESIGNSKIN)", 1600),
            bCell("케이스 = 패션. 2011년부터 프리미엄 디자인 추구", 2200),
            bCell("우아한 텍스처 + 슬림 디자인. 아이폰·갤럭시 프리미엄 악세서리 라인. 글로벌 수출까지 확장", 2780),
            bCell("designskin.com 운영. \"왜 우아한 케이스가 없을까?\"에서 출발. 케이스를 패션으로 재정의한 브랜드", 2780),
          ] }),
        ]
      }),

      new Paragraph({ children: [new PageBreak()] }),

      // ===== 4. \uCEEC\uB7EC \uD2B8\uB80C\uB4DC =====
      new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun("4. 2026\uB144 \uD55C\uAD6D \uC18C\uBE44\uC790 \uCEEC\uB7EC \uC120\uD638\uB3C4")] }),
      new Table({
        width: { size: 9360, type: WidthType.DXA },
        columnWidths: [2340, 2340, 2340, 2340],
        rows: [
          new TableRow({ children: [ hCell("\uCEEC\uB7EC \uADF8\uB8F9", 2340), hCell("\uB300\uD45C \uC0C9\uC0C1", 2340), hCell("\uC801\uD569\uD55C \uB514\uC790\uC778", 2340), hCell("\uBE44\uACE0", 2340) ] }),
          new TableRow({ children: [
            bCell("\uD30C\uC2A4\uD154 \uBB34\uB4DC", 2340, "FFF0F5"), bCell("Dusty Rose, Lavender, Mint, Haze Blue", 2340, "FFF0F5"), bCell("\uD50C\uB85C\uB7F4, \uCCB4\uD06C, \uD558\uD2B8, \uBBF4\uB4DC \uBB38\uAD6C", 2340, "FFF0F5"), bCell("10~20\uB300 \uC5EC\uC131 \uC8FC \uD0C0\uAC9F", 2340, "FFF0F5"),
          ] }),
          new TableRow({ children: [
            bCell("\uB274\uD2B8\uB7F4 \uD1A4", 2340), bCell("Sand, Clay, Cream Beige, Stone", 2340), bCell("\uBBF8\uB2C8\uBA40, \uBB34\uC9C0, \uB808\uB354 \uD14D\uC2A4\uCC98", 2340), bCell("20~30\uB300 \uC5EC\uC131, \uC131\uBCC4 \uBB34\uAD00", 2340),
          ] }),
          new TableRow({ children: [
            bCell("\uBE44\uBE44\uB4DC \uD3F8\uCE58", 2340, "F0FFF0"), bCell("Bubblegum Pink, Lime Green, Electric Blue", 2340, "F0FFF0"), bCell("Y2K, \uD22C\uBA85 \uCEEC\uB7EC, \uCCB4\uD06C", 2340, "F0FFF0"), bCell("10~20\uB300, SNS \uC0AC\uC9C4\uBC1C \uC88B\uC74C", 2340, "F0FFF0"),
          ] }),
          new TableRow({ children: [
            bCell("\uB525 \uD1A4", 2340), bCell("Slate Grey, Charcoal, Deep Green", 2340), bCell("\uBBF8\uB2C8\uBA40, \uB808\uB354, \uBA54\uD0C8\uB9AD", 2340), bCell("20~30\uB300 \uB0A8\uB140 \uD3EC\uD568", 2340),
          ] }),
          new TableRow({ children: [
            bCell("\uBA54\uD0C8\uB9AD / \uD06C\uB86C", 2340, "F8F8FF"), bCell("Silver, Chrome, Hologram", 2340, "F8F8FF"), bCell("\uBBF8\uB7EC, \uC5D0\uD3ED\uC2DC, \uD2B9\uBCC4\uD310", 2340, "F8F8FF"), bCell("\uC804 \uC5F0\uB839, \uC778\uC2A4\uD0C0 \uC18C\uC7AC\uC6A9 \uADF9\uAC15", 2340, "F8F8FF"),
          ] }),
        ]
      }),

      new Paragraph({ children: [new PageBreak()] }),

      // ===== 5. \uBA54\uD0C0\uAD11\uACE0 \uC18C\uC7AC \uC81C\uC791 \uD301 =====
      new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun("5. \uBA54\uD0C0\uAD11\uACE0 \uC18C\uC7AC \uC81C\uC791 \uD301 (\uD55C\uAD6D \uC18C\uB9E4 \uAE30\uC900)")] }),

      new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun("\uAD11\uACE0 \uD3EC\uB9F7")] }),
      new Paragraph({ numbering: { reference: "nums", level: 0 }, spacing: { after: 80 }, children: [
        new TextRun({ text: "\uB9B4\uC2A4 15\uCD08 \uC774\uD558 \uC21F\uD3FC \uBE44\uB514\uC624: ", bold: true, size: 22 }), new TextRun({ text: "2026\uB144 \uAC00\uC7A5 \uB192\uC740 \uC804\uD658\uC728. \uC18C\uB9AC \uC5C6\uC774\uB3C4 \uC774\uD574 \uAC00\uB2A5\uD55C \uBE44\uC8FC\uC5BC \uC911\uC2EC. \uCF00\uC774\uC2A4\uB97C \uC190\uC5D0 \uB4E4\uACE0 \uBCF4\uC5EC\uC8FC\uB294 UGC \uC2A4\uD0C0\uC77C\uC774 \uD6A8\uACFC\uC801", size: 22 })
      ] }),
      new Paragraph({ numbering: { reference: "nums", level: 0 }, spacing: { after: 80 }, children: [
        new TextRun({ text: "\uCE90\uB7EC\uC140 \uAD11\uACE0 (3~5\uC7A5): ", bold: true, size: 22 }), new TextRun({ text: "\uCEEC\uB7EC \uBC14\uB9AC\uC5D0\uC774\uC158 \uBCF4\uC5EC\uC8FC\uAE30, \uC138\uD2B8 \uAD6C\uC131 \uBCF4\uC5EC\uC8FC\uAE30\uC5D0 \uCD5C\uC801. \"\uCDE8\uD5A5\uBCC4 \uACE8\uB77C\uBCF4\uC138\uC694\" \uCEE8\uC149\uC73C\uB85C \uD074\uB9AD \uC720\uB3C4", size: 22 })
      ] }),
      new Paragraph({ numbering: { reference: "nums", level: 0 }, spacing: { after: 80 }, children: [
        new TextRun({ text: "\uC2A4\uD1A0\uB9AC \uAD11\uACE0: ", bold: true, size: 22 }), new TextRun({ text: "\uD480\uC2A4\uD06C\uB9B0 \uC81C\uD488 \uC774\uBBF8\uC9C0 + \uAC04\uACB0\uD55C CTA(\"\uC9C0\uAE08 \uAD6C\uB9E4\uD558\uAE30\"). \uD2B9\uAC00/\uC2DC\uC98C \uD504\uB85C\uBAA8\uC158\uC5D0 \uD6A8\uACFC\uC801", size: 22 })
      ] }),

      new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun("\uD55C\uAD6D \uD0C0\uAC9F \uD575\uC2EC \uC804\uB7B5")] }),
      new Paragraph({ numbering: { reference: "bullets", level: 0 }, spacing: { after: 80 }, children: [
        new TextRun({ text: "\uBE0C\uB79C\uB4DC \uC815\uCCB4\uC131\uC774 \uACE7 \uACBD\uC7C1\uB825: ", bold: true, size: 22 }), new TextRun({ text: "\uC218\uBC31 \uBA85\uC758 \uC140\uB7EC\uAC00 \uBE44\uC2B7\uD55C \uC81C\uD488\uC744 \uD314\uAE30 \uB54C\uBB38\uC5D0, \uACE0\uAC1D\uC740 \uC81C\uD488\uBCF4\uB2E4 '\uC77C\uAD00\uB41C \uD1A4, \uD328\uD0A4\uC9C0, \uAC10\uC131'\uC5D0 \uB044\uB9BC", size: 22 })
      ] }),
      new Paragraph({ numbering: { reference: "bullets", level: 0 }, spacing: { after: 80 }, children: [
        new TextRun({ text: "\uC138\uD2B8 \uAD6C\uC131\uC73C\uB85C \uAC1D\uB2E8\uAC00 \uC0C1\uC2B9: ", bold: true, size: 22 }), new TextRun({ text: "\uCF00\uC774\uC2A4+\uC2A4\uB9C8\uD2B8\uD1A1+\uD0A4\uB9C1 \uC138\uD2B8, \uCF00\uC774\uC2A4+\uC561\uC138\uC11C\uB9AC \uC138\uD2B8 \uB4F1\uC73C\uB85C \uAC1D\uB2E8\uAC00\uB97C \uB192\uC774\uB294 \uC804\uB7B5\uC774 \uBA54\uD0C0\uAD11\uACE0\uC5D0\uC11C \uD65C\uBC1C", size: 22 })
      ] }),
      new Paragraph({ numbering: { reference: "bullets", level: 0 }, spacing: { after: 80 }, children: [
        new TextRun({ text: "\uC2DC\uC98C \uAC10\uC131 \uC18C\uC7AC \uAD50\uCCB4: ", bold: true, size: 22 }), new TextRun({ text: "\uBD04=\uD50C\uB85C\uB7F4/\uD30C\uC2A4\uD154, \uC5EC\uB984=\uBE44\uBE44\uB4DC/\uD22C\uBA85, \uAC00\uC744=\uB525\uD1A4/\uBE48\uD2F0\uC9C0, \uACA8\uC6B8=\uBA54\uD0C8\uB9AD/\uD06C\uB86C \uC73C\uB85C \uACC4\uC808\uB9C8\uB2E4 \uAD11\uACE0 \uC18C\uC7AC\uB97C \uBC14\uAFB8\uBA74 \uC7AC\uAD6C\uB9E4 \uC720\uB3C4 \uAC00\uB2A5", size: 22 })
      ] }),
      new Paragraph({ numbering: { reference: "bullets", level: 0 }, spacing: { after: 80 }, children: [
        new TextRun({ text: "60\uC77C \uC774\uC0C1 \uC9D1\uD589\uB41C \uACBD\uC7C1\uC0AC \uAD11\uACE0 = \uC218\uC775\uC131 \uAC80\uC99D\uB428: ", bold: true, size: 22 }), new TextRun({ text: "Meta Ad Library\uC5D0\uC11C \uC624\uB798 \uB3CC\uC544\uAC00\uB294 \uACBD\uC7C1\uC0AC \uAD11\uACE0\uB97C \uCC38\uACE0\uD558\uC138\uC694", size: 22 })
      ] }),

      new Paragraph({ children: [new PageBreak()] }),

      // ===== 6. \uC2E0\uADDC \uB514\uC790\uC778 \uBC29\uD5A5 \uC81C\uC548 =====
      new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun("6. \uC2E0\uADDC \uB514\uC790\uC778 \uBC29\uD5A5 \uC81C\uC548")] }),
      new Paragraph({ spacing: { after: 160 }, children: [
        new TextRun({ text: "\uD604\uC7AC \uD55C\uAD6D \uBA54\uD0C0\uAD11\uACE0 \uC2DC\uC7A5\uC5D0\uC11C \uC2E4\uC81C\uB85C \uC798 \uD314\uB9AC\uACE0 \uC788\uB294 \uD328\uD134\uACFC, \uC544\uC9C1 \uACBD\uC7C1\uC774 \uB35C\uD55C \uD2C8\uC0C8 \uB514\uC790\uC778\uC744 \uC885\uD569\uD574\uC11C \uC81C\uC548\uD569\uB2C8\uB2E4:", size: 22 })
      ] }),
      new Paragraph({ numbering: { reference: "nums", level: 0 }, spacing: { after: 120 }, children: [
        new TextRun({ text: "\uBBF4\uB4DC \uCEEC\uB7EC \uD22C\uBA85 \uCF00\uC774\uC2A4 \uB77C\uC778\uC5C5: ", bold: true, size: 22 }), new TextRun({ text: "\uD22C\uBA85\uCF00\uC774\uC2A4\uB294 \uD55C\uAD6D \uBA54\uD0C0\uAD11\uACE0\uC5D0\uC11C \uAC00\uC7A5 \uB9CE\uC774 \uC9D1\uD589\uB418\uB294 \uCE74\uD14C\uACE0\uB9AC. \uD30C\uC2A4\uD154/\uBBF4\uB4DC \uCEEC\uB7EC \uD22C\uBA85\uCF00\uC774\uC2A4\uB97C 5~8\uAC1C \uCEEC\uB7EC\uB85C \uAD6C\uC131\uD574\uC11C \uCE90\uB7EC\uC140 \uAD11\uACE0\uB85C \uD14C\uC2A4\uD2B8", size: 22 })
      ] }),
      new Paragraph({ numbering: { reference: "nums", level: 0 }, spacing: { after: 120 }, children: [
        new TextRun({ text: "빈티지 플로럴 컬렉션: ", bold: true, size: 22 }), new TextRun({ text: "수채화 터치 + 은은한 톤의 꽃무늬. 아우렐처럼 일러스트 아트워크 기반 시즌 컬렉션으로 구성하면 효과적. AI 피로감으로 인해 '손그림 느낌' 디자인 수요가 증가 중", size: 22 })
      ] }),
      new Paragraph({ numbering: { reference: "nums", level: 0 }, spacing: { after: 120 }, children: [
        new TextRun({ text: "\uCCB4\uD06C \uD328\uD134 \uBCC0\uD615: ", bold: true, size: 22 }), new TextRun({ text: "\uD558\uD2B8 \uCCB4\uD06C, \uCEEC\uB7EC \uACA9\uC790, \uBB3C\uACB0 \uCCB4\uD06C \uB4F1 \uBCC0\uD615\uC744 3~5\uAC1C \uB77C\uC778\uC5C5\uC73C\uB85C. \uCF54\uB514 \uB9E4\uCE6D \uCEE8\uC149\uC73C\uB85C \uC778\uC2A4\uD0C0 \uAD11\uACE0 \uC81C\uC791\uD558\uBA74 \uC800\uC7A5/\uACF5\uC720 \uB192\uC74C", size: 22 })
      ] }),
      new Paragraph({ numbering: { reference: "nums", level: 0 }, spacing: { after: 120 }, children: [
        new TextRun({ text: "\uBBF8\uB7EC / \uC2E4\uBC84 \uCF00\uC774\uC2A4: ", bold: true, size: 22 }), new TextRun({ text: "\uC140\uD53C \uBBF8\uB7EC \uD6A8\uACFC + \uBA54\uD0C8\uB9AD \uAC10\uC131. \uB9B4\uC2A4 \uAD11\uACE0\uC5D0\uC11C \uBE5B \uBC18\uC0AC \uD6A8\uACFC\uB97C \uBCF4\uC5EC\uC8FC\uBA74 \uD074\uB9AD\uC728 \uADF9\uB300\uD654. \uB208\uCF54\u2661\uCC98\uB7FC \uC5D0\uD3ED\uC2DC \uC7A5\uC2DD \uCD94\uAC00\uD558\uBA74 \uCC28\uBCC4\uD654", size: 22 })
      ] }),
      new Paragraph({ numbering: { reference: "nums", level: 0 }, spacing: { after: 120 }, children: [
        new TextRun({ text: "\uBB38\uAD6C / \uAC10\uC131 \uB77C\uC778: ", bold: true, size: 22 }), new TextRun({ text: "MBTI, \uC9E7\uC740 \uC601\uBB38/\uD55C\uAE00 \uBB38\uAD6C, \uACF5\uAC10\uD615 \uC720\uBA38 \uBB38\uAD6C. ROAS 620% \uC131\uACF5\uC0AC\uB840\uCC98\uB7FC '\uB098\uB97C \uD45C\uD604\uD558\uB294 \uCF00\uC774\uC2A4' \uCEE8\uC149\uC774 \uD55C\uAD6D\uC5D0\uC11C \uAC15\uB825\uD568", size: 22 })
      ] }),
      new Paragraph({ numbering: { reference: "nums", level: 0 }, spacing: { after: 120 }, children: [
        new TextRun({ text: "\uCF00\uC774\uC2A4+\uC561\uC138\uC11C\uB9AC \uC138\uD2B8: ", bold: true, size: 22 }), new TextRun({ text: "\uCF00\uC774\uC2A4+\uC2A4\uB9C8\uD2B8\uD1A1+\uD0A4\uB9C1, \uCF00\uC774\uC2A4+\uCE74\uB4DC\uD640\uB354 \uB4F1 \uC138\uD2B8 \uAD6C\uC131. \uAC1D\uB2E8\uAC00 \uC0C1\uC2B9\uC2DC\uD0A4\uBA74\uC11C \"\uC4F0\uC544\uC9C0\uB294 \uC0AC\uC740\uD488\" \uCEE8\uC149\uC73C\uB85C \uC804\uD658 \uC720\uB3C4", size: 22 })
      ] }),

      new Paragraph({ spacing: { before: 400 }, border: { top: { style: BorderStyle.SINGLE, size: 3, color: "CCCCCC", space: 1 } }, children: [] }),
      new Paragraph({ spacing: { before: 100 }, children: [new TextRun({ text: "\uCC38\uACE0: Meta Ad Library \uD55C\uAD6D \uD0C0\uAC9F \uD3F0\uCF00\uC774\uC2A4 \uAD11\uACE0 \uBD84\uC11D, Accio \uD55C\uAD6D \uD3F0\uCF00\uC774\uC2A4 \uD2B8\uB80C\uB4DC \uB9AC\uD3EC\uD2B8, AMPM\uAE00\uB85C\uBC8C \uD3F0\uCF00\uC774\uC2A4 \uAD11\uACE0 \uC131\uACF5\uC0AC\uB840, \uD2B8\uB80C\uB4DC\uBAA8\uB2C8\uD130 \uC2A4\uB9C8\uD2B8\uD3F0 \uCF00\uC774\uC2A4 \uC870\uC0AC, Brunch \uD575\uB4DC\uD3F0\uCF00\uC774\uC2A4 \uB514\uC790\uC778 \uD2B8\uB80C\uB4DC, Alibaba \uD55C\uAD6D \uD3F0\uCF00\uC774\uC2A4 \uBD84\uC11D", size: 18, color: "999999", italics: true })] }),
      new Paragraph({ children: [new TextRun({ text: "\uC791\uC131\uC77C: 2026\uB144 4\uC6D4 14\uC77C", size: 18, color: "999999", italics: true })] }),
    ]
  }]
});

Packer.toBuffer(doc).then(buffer => {
  fs.writeFileSync("/sessions/zealous-confident-pasteur/mnt/outputs/한국_소매_폰케이스_메타광고_디자인트렌드.docx", buffer);
  console.log("Report created successfully!");
});
