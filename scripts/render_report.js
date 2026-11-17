// analysis.json -> 美化 docx 报告（fakao-answer-review 配套）
// 用法: node render_report.js <analysis.json> <output.docx>
// 依赖: npm i -g docx；运行时设置 NODE_PATH=$(npm root -g)
const fs = require("fs");
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  AlignmentType, HeadingLevel, BorderStyle, WidthType, ShadingType,
  Header, Footer, PageNumber, PageBreak
} = require("docx");

if (process.argv.length < 4) {
  console.log("usage: node render_report.js <analysis.json> <output.docx>");
  process.exit(1);
}
const data = JSON.parse(fs.readFileSync(process.argv[2], "utf8"));
const OUT = process.argv[3];

const FONT = { ascii: "Times New Roman", eastAsia: "宋体", hAnsi: "Times New Roman" };
const HFONT = { ascii: "Arial", eastAsia: "黑体", hAnsi: "Arial" };
const BLUE = "1F4E79", GRAY = "595959", RED = "C00000", GREEN = "2E7D32", ORANGE = "BF6900";
const LINE = { style: BorderStyle.SINGLE, size: 2, color: "BFBFBF" };
const BORDERS = { top: LINE, bottom: LINE, left: LINE, right: LINE };
const PAGE_W = 9026; // A4 可用宽度 (DXA)

function tr(text, o = {}) {
  return new TextRun({ text, font: o.h ? HFONT : FONT, size: o.size || 21, bold: !!o.bold, color: o.color || "000000" });
}
function P(runs, o = {}) {
  return new Paragraph({
    spacing: { before: o.before || 0, after: o.after === undefined ? 120 : o.after, line: o.line || 360 },
    indent: o.indent ? { left: o.indent } : undefined,
    children: runs,
  });
}
function H1(text) {
  return new Paragraph({ heading: HeadingLevel.HEADING_1, spacing: { before: 360, after: 200 },
    border: { bottom: { style: BorderStyle.SINGLE, size: 12, color: BLUE, space: 4 } },
    children: [new TextRun({ text, font: HFONT, size: 30, bold: true, color: BLUE })] });
}
function H2(text) {
  return new Paragraph({ heading: HeadingLevel.HEADING_2, spacing: { before: 280, after: 160 },
    children: [new TextRun({ text, font: HFONT, size: 25, bold: true, color: BLUE })] });
}
function H3(text) {
  return new Paragraph({ heading: HeadingLevel.HEADING_3, spacing: { before: 200, after: 120 },
    children: [new TextRun({ text, font: HFONT, size: 22, bold: true, color: "333333" })] });
}

const DIAG = {
  ok:    { label: "【答对之处】", color: GREEN },
  error: { label: "【定性错误】", color: RED },
  miss:  { label: "【层次缺失】", color: ORANGE },
  warn:  { label: "【边界提醒】", color: ORANGE },
};
function diag(d) {
  const m = DIAG[d.type] || DIAG.miss;
  return new Paragraph({ spacing: { after: 100, line: 360 }, indent: { left: 240 },
    children: [tr(d.label || m.label, { bold: true, color: m.color }), tr(d.text)] });
}

function cellPara(text, o = {}) {
  return new Paragraph({ alignment: o.align || AlignmentType.LEFT, spacing: { line: 320, after: 40 },
    children: [tr(text, { size: 20, bold: o.bold, color: o.color || (o.header ? "FFFFFF" : "000000"), h: o.header })] });
}
function cell(paras, w, o = {}) {
  return new TableCell({ borders: BORDERS, width: { size: w, type: WidthType.DXA },
    shading: o.fill ? { fill: o.fill, type: ShadingType.CLEAR } : undefined,
    margins: { top: 80, bottom: 80, left: 120, right: 120 },
    children: Array.isArray(paras) ? paras : [paras] });
}
function markCell(mark) {
  const map = { hit: ["✔", GREEN], mid: ["部分", ORANGE], miss: ["✘", RED] };
  const [t, c] = map[mark] || [mark, "000000"];
  return cellPara(t, { bold: true, color: c, align: AlignmentType.CENTER });
}

const children = [];

// ---- 标题区 ----
children.push(new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 240, after: 80 },
  children: [new TextRun({ text: data.title, font: HFONT, size: 40, bold: true, color: BLUE })] }));
if (data.subtitle) children.push(new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 60 },
  children: [new TextRun({ text: data.subtitle, font: FONT, size: 22, color: GRAY })] }));
if (data.meta) children.push(new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 120 },
  children: [new TextRun({ text: data.meta, font: FONT, size: 20, color: GRAY })] }));
children.push(new Paragraph({ spacing: { after: 200 }, border: { bottom: { style: BorderStyle.SINGLE, size: 12, color: BLUE, space: 1 } }, children: [] }));

// ---- 一、总体诊断 ----
children.push(H1("一、总体诊断"));
children.push(P([tr("总体印象：", { bold: true }), tr(data.overall.impression)]));
(data.overall.gaps || []).forEach(g => {
  children.push(P([tr(g.label, { bold: true, color: g.sev === "error" ? RED : ORANGE }), tr(g.text)], { indent: 240 }));
});
if (data.overall.scoreNote) children.push(P([tr(data.overall.scoreNote, { bold: true })], { before: 120 }));

// 估分表
const sw = [1300, 900, 1500, PAGE_W - 1300 - 900 - 1500];
const sRows = [new TableRow({ tableHeader: true, children: ["设问", "分值", "预估得分", "核心失分点"].map((h, i) =>
  cell(cellPara(h, { header: true, bold: true, align: AlignmentType.CENTER }), sw[i], { fill: BLUE })) })];
data.scoreTable.rows.forEach(r => sRows.push(new TableRow({ children: r.map((c, i) =>
  cell(cellPara(c, { align: i === 0 ? AlignmentType.LEFT : (i === 3 ? AlignmentType.LEFT : AlignmentType.CENTER) }), sw[i], {})) })));
if (data.scoreTable.total) {
  const t = data.scoreTable.total;
  sRows.push(new TableRow({ children: [
    cell(cellPara(t[0], { bold: true, align: AlignmentType.CENTER }), sw[0], {}),
    cell(cellPara(t[1], { bold: true, align: AlignmentType.CENTER }), sw[1], {}),
    cell(cellPara(t[2], { bold: true, color: RED, align: AlignmentType.CENTER }), sw[2], {}),
    cell(cellPara(t[3], {}), sw[3], {}),
  ] }));
}
children.push(new Table({ width: { size: PAGE_W, type: WidthType.DXA }, columnWidths: sw, rows: sRows }));

// ---- 二、逐问精细对比 ----
children.push(H1("二、逐问精细对比"));
const qw = [4400, 1100, PAGE_W - 4400 - 1100];
data.questions.forEach(q => {
  children.push(H2(q.heading));
  children.push(P([tr("考点定位：", { bold: true }), tr(q.points)]));
  children.push(P([tr("你的答案要点：", { bold: true, color: BLUE }), tr(q.mine)]));
  children.push(P([tr("参考答案要点：", { bold: true, color: BLUE }), tr(q.reference)]));
  children.push(P([tr("差异诊断：", { bold: true })]));
  q.diagnosis.forEach(d => children.push(diag(d)));
  const rows = [new TableRow({ tableHeader: true, children: ["采分点", "命中", "说明"].map((h, i) =>
    cell(cellPara(h, { header: true, bold: true, align: AlignmentType.CENTER }), qw[i], { fill: BLUE })) })];
  q.rubric.forEach(r => rows.push(new TableRow({ children: [
    cell(cellPara(r[0], {}), qw[0], {}),
    cell(markCell(r[1]), qw[1], {}),
    cell(cellPara(r[2], {}), qw[2], {}),
  ] })));
  children.push(new Table({ width: { size: PAGE_W, type: WidthType.DXA }, columnWidths: qw, rows }));
});

// ---- 三、案件涉及法条全文（法条定位练习） ----
if (Array.isArray(data.statutes) && data.statutes.length) {
  children.push(new Paragraph({ children: [new PageBreak()] }));
  children.push(H1("三、案件涉及法条全文（法条定位练习）"));
  data.statutes.forEach(s => {
    children.push(H3(s.title));
    if (s.source) children.push(P([tr(s.source, { size: 18, color: GRAY })], { after: 60 }));
    if (s.text) children.push(P([tr(s.text)], { indent: 240, after: 160 }));
  });
}

// ---- 四、知识漏洞清单 ----
children.push(new Paragraph({ children: [new PageBreak()] }));
children.push(H1("四、知识漏洞清单（按优先级排序，须逐项补齐）"));
data.holes.forEach(h => {
  children.push(H3(h.title));
  children.push(P([tr(h.yours, { color: RED })], { indent: 240, after: 60 }));
  children.push(P([tr(h.rule)], { indent: 240, after: 60 }));
  children.push(P([tr(h.anchor, { bold: true, color: BLUE })], { indent: 240 }));
});

// ---- 五、答题技巧 ----
children.push(H1("五、答题技巧：如何向参考答案学习"));
data.tips.forEach(t => children.push(P([tr(t.title + "。", { bold: true, color: BLUE }), tr(t.body)], { after: 140 })));

// ---- 六、复盘计划 ----
children.push(H1("六、针对性复盘计划"));
const pw = [1100, 2900, 1300, PAGE_W - 1100 - 2900 - 1300];
function starColor(s) { return s.startsWith("★★★") ? RED : s.startsWith("★★☆") ? ORANGE : GREEN; }
const planRows = [new TableRow({ tableHeader: true, children: data.plan.headers.map((h, i) =>
  cell(cellPara(h, { header: true, bold: true, align: AlignmentType.CENTER }), pw[i], { fill: BLUE })) })];
data.plan.rows.forEach(r => planRows.push(new TableRow({ children: r.map((c, i) =>
  cell(i === 0 ? cellPara(c, { bold: true, color: starColor(c), align: AlignmentType.CENTER }) : cellPara(c, {}), pw[i], {})) })));
children.push(new Table({ width: { size: PAGE_W, type: WidthType.DXA }, columnWidths: pw, rows: planRows }));

// ---- 结语 ----
if (data.conclusion) children.push(P([tr("结语：", { bold: true }), tr(data.conclusion)], { before: 160 }));

const doc = new Document({
  styles: { default: { document: { run: { font: FONT, size: 21 } } } },
  sections: [{
    properties: { page: { margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 } } },
    headers: { default: new Header({ children: [new Paragraph({ alignment: AlignmentType.RIGHT,
      children: [new TextRun({ text: data.headerText || "法考主观题 · 答案对比分析", font: FONT, size: 18, color: GRAY })] })] }) },
    footers: { default: new Footer({ children: [new Paragraph({ alignment: AlignmentType.CENTER,
      children: [new TextRun({ text: "第 ", font: FONT, size: 18, color: GRAY }),
        new TextRun({ children: [PageNumber.CURRENT], font: FONT, size: 18, color: GRAY }),
        new TextRun({ text: " 页", font: FONT, size: 18, color: GRAY })] })] }) },
    children,
  }],
});

Packer.toBuffer(doc).then(buf => { fs.writeFileSync(OUT, buf); console.log("OK bytes=" + buf.length); });
