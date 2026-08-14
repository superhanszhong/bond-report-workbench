import {
  AlignmentType, BorderStyle, Document, Packer, Paragraph, Table, TableCell, TableLayoutType,
  TableRow, TextRun, VerticalAlign, WidthType,
} from "docx";
import { fridayOf } from "./workbench.ts";
import type { ParsedBondRecord } from "./workbench.ts";

type ReportInput = {
  weekStart: string;
  summary: string;
  localRecords: ParsedBondRecord[];
  spreadRecords: ParsedBondRecord[];
  previousSpreadRecords?: ParsedBondRecord[];
  ytdLocalRecords?: ParsedBondRecord[];
};

const BLACK = "000000";
const WHITE = "FFFFFF";
const ORANGE = "F4B083";
const LIGHT_BORDER = "BEBEBE";
const FONT = { ascii: "SimSun", hAnsi: "SimSun", eastAsia: "SimSun", cs: "SimSun", hint: "eastAsia" } as const;
const noneBorder = { style: BorderStyle.NONE, size: 0, color: WHITE };
const lightBorder = { style: BorderStyle.SINGLE, size: 4, color: LIGHT_BORDER };
const darkBorder = { style: BorderStyle.SINGLE, size: 8, color: BLACK };

function mmdd(date: string) { return date.slice(5).replace("-", ""); }
function yyyymmdd(date: string) { return date.replaceAll("-", ""); }
function formatMd(value: string) {
  const date = new Date(`${value}T12:00:00`);
  return `${date.getMonth() + 1}/${date.getDate()}`;
}
function weekday(value: string) {
  return ["周日", "周一", "周二", "周三", "周四", "周五", "周六"][new Date(`${value}T12:00:00`).getDay()];
}
function text(value: unknown, fallback = "-") {
  if (value === null || value === undefined || value === "") return fallback;
  if (typeof value === "number") return String(Math.round(value * 10000) / 10000);
  return String(value);
}
function amount(rows: ParsedBondRecord[]) {
  return rows.reduce((sum, row) => sum + (row.amount || 0), 0);
}
function display(row: ParsedBondRecord, key: string) {
  const map = row.raw?.__display;
  return map && typeof map === "object" ? text((map as Record<string, unknown>)[key]) : "-";
}
function displayRate(row: ParsedBondRecord, key: string) {
  return display(row, key).replace(/%$/, "");
}
function lines(value: unknown, options: { bold?: boolean; color?: string; size?: number } = {}) {
  return text(value).split("\n").map((part, index) => new TextRun({
    text: part, break: index ? 1 : undefined, bold: options.bold, color: options.color || BLACK,
    size: options.size || 18, font: FONT,
  }));
}
function reportCell(value: unknown, options: {
  bold?: boolean; fill?: string; color?: string; size?: number; width?: number;
  align?: (typeof AlignmentType)[keyof typeof AlignmentType]; borders?: typeof lightBorder;
  columnSpan?: number;
} = {}) {
  const border = options.borders || lightBorder;
  return new TableCell({
    columnSpan: options.columnSpan,
    width: options.width ? { size: options.width, type: WidthType.DXA } : undefined,
    verticalAlign: VerticalAlign.CENTER,
    shading: options.fill ? { fill: options.fill } : undefined,
    borders: { top: border, bottom: border, left: border, right: border },
    margins: { top: 55, bottom: 55, left: 55, right: 55 },
    children: [new Paragraph({
      alignment: options.align || AlignmentType.CENTER,
      spacing: { before: 0, after: 0, line: 240 },
      children: lines(value, { bold: options.bold, color: options.color, size: options.size }),
    })],
  });
}
function titleParagraph(value: string) {
  return new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { after: 80 },
    children: [new TextRun({ text: value, bold: true, size: 24, color: BLACK, font: FONT })],
  });
}
function sectionHeading(value: string) {
  return new Paragraph({
    keepNext: true,
    spacing: { before: 120, after: 0, line: 360 },
    children: [new TextRun({ text: value, bold: true, size: 21, color: BLACK, font: FONT })],
  });
}
function bodyParagraph(value: string, keepNext = false) {
  return new Paragraph({
    keepNext,
    spacing: { before: 0, after: 0, line: 315 },
    children: [new TextRun({ text: value, size: 21, color: BLACK, font: FONT })],
  });
}
function localNature(row: ParsedBondRecord) {
  const nature = text(row.raw?.["性质"], "");
  const kind = text(row.raw?.["类型"], "");
  if (/置换/.test(nature)) return "置换债";
  if (/再融资/.test(nature)) return /一般/.test(kind) ? "再融资一般债" : "再融资专项债";
  return /一般/.test(kind) ? "新增一般债" : "新增专项债";
}
function localSummaryTable(current: ParsedBondRecord[], ytd: ParsedBondRecord[]) {
  const groups = ["新增一般债", "再融资一般债", "新增专项债", "再融资专项债", "置换债"];
  const quota: Record<string, number> = { 新增一般债: 8000, 新增专项债: 44000, 置换债: 20000 };
  const rows = groups.map((label) => {
    const weekly = amount(current.filter((row) => localNature(row) === label));
    const annual = amount(ytd.filter((row) => localNature(row) === label));
    const limit = quota[label];
    return [label, text(weekly), text(annual), limit ? `${(annual / limit * 100).toFixed(2)}%` : "/", "/", limit ? text(limit - annual) : "/"];
  });
  const weeklyTotal = amount(current);
  const annualTotal = amount(ytd);
  rows.push(["合计(不含普通再融资债)", text(weeklyTotal), text(annualTotal), "/", "/", "/"]);
  const headers = [" ", "本周（亿）", "今年以来（亿）", "今年发行进度", "去年发行进度", "剩余（亿）"];
  return new Table({
    width: { size: 100, type: WidthType.PERCENTAGE }, layout: TableLayoutType.FIXED,
    rows: [
      new TableRow({ tableHeader: true, cantSplit: true, children: headers.map((value) => reportCell(value, { bold: true, fill: ORANGE, size: 18, borders: darkBorder })) }),
      ...rows.map((values, index) => new TableRow({ cantSplit: true, children: values.map((value, column) => reportCell(value, { bold: index === rows.length - 1 || column === 0, size: 18, align: column === 0 ? AlignmentType.LEFT : AlignmentType.CENTER, borders: darkBorder })) })),
    ],
  });
}
function tenorLabel(row: ParsedBondRecord) {
  const raw = text(row.tenor, "");
  const unit = /[DMY]$/i.test(raw) ? raw.toUpperCase() : `${raw}Y`;
  return /浮息/.test(row.summaryMeta?.rateType || row.remark || "") ? `${unit}浮息` : unit;
}
function planLabel(row: ParsedBondRecord) {
  if (row.bondType === "国债") {
    if (/^269/.test(row.bondCode || "")) return "贴现国债";
    if ((Number.parseFloat(row.tenor || "") || 0) >= 20) return "超长特国";
    return "附息国债";
  }
  const short = row.bondType === "国开债" ? "国开" : row.bondType === "口行债" ? "进出" : row.bondType === "农发债" ? "农发" : row.bondType || "利率债";
  return (row.summaryMeta?.route === "报价发行" || row.issuanceRoute === "报价发行") ? `${short}清发（报价）` : short;
}
function dailyPlan(rows: ParsedBondRecord[]) {
  if (!rows.length) return "-";
  const groups = new Map<string, string[]>();
  rows.forEach((row) => {
    const label = planLabel(row);
    const item = `${tenorLabel(row)}:${text(row.amount)}亿`;
    groups.set(label, [...(groups.get(label) || []), item]);
  });
  return [...groups].map(([label, items]) => `${label}\n${items.join("\n")}`).join("\n");
}
function varietyTotals(rows: ParsedBondRecord[]) {
  const order = ["贴现国债", "农发", "国开", "进出", "超长特国", "附息国债"];
  const totals = new Map<string, number>();
  rows.forEach((row) => {
    const label = planLabel(row).replace("清发（报价）", "");
    totals.set(label, (totals.get(label) || 0) + (row.amount || 0));
  });
  return order.filter((label) => totals.has(label)).map((label) => `${label}:${text(totals.get(label))}亿`).join("　");
}
function weeklyPlanTable(weekStart: string, spreads: ParsedBondRecord[], locals: ParsedBondRecord[]) {
  const dates = Array.from({ length: 5 }, (_, index) => {
    const date = new Date(`${weekStart}T12:00:00`); date.setDate(date.getDate() + index);
    return date.toISOString().slice(0, 10);
  });
  const totalRate = amount(spreads);
  const totalLocal = amount(locals);
  const headers = [" ", "周一", "周二", "周三", "周四", "周五"];
  const dailyRate = dates.map((date) => spreads.filter((row) => row.tradeDate === date));
  const dailyLocal = dates.map((date) => locals.filter((row) => row.tradeDate === date));
  return new Table({
    width: { size: 100, type: WidthType.PERCENTAGE }, layout: TableLayoutType.FIXED,
    rows: [
      new TableRow({ tableHeader: true, cantSplit: true, children: headers.map((value) => reportCell(value, { bold: true, fill: ORANGE, size: 18 })) }),
      new TableRow({ cantSplit: true, children: [reportCell("发行安排", { bold: true, size: 17 }), ...dailyRate.map((rows) => reportCell(dailyPlan(rows), { size: 16 }))] }),
      new TableRow({ cantSplit: true, children: [reportCell(`国债政金债合计\n${text(totalRate)}亿`, { bold: true, size: 17 }), ...dailyRate.map((rows) => reportCell(text(amount(rows)), { size: 17 }))] }),
      new TableRow({ cantSplit: true, children: [reportCell("本周合计", { bold: true, size: 17 }), reportCell(varietyTotals(spreads) || "-", { size: 16, columnSpan: 5 })] }),
      new TableRow({ cantSplit: true, children: [reportCell(`地方债\n${text(totalLocal)}亿`, { bold: true, size: 17 }), ...dailyLocal.map((rows) => reportCell(rows.length ? text(amount(rows)) : "-", { size: 17 }))] }),
    ],
  });
}
function routeLabel(row: ParsedBondRecord) {
  const route = row.summaryMeta?.route || row.issuanceRoute || "中债招标";
  if (route === "报价发行") return row.bondType === "农发债" ? "直投报价发行" : "报价发行";
  return row.bondType === "国债" ? "发行" : "招标发行";
}
function reviewBondLabel(row: ParsedBondRecord) {
  if (/^09/.test(row.bondCode || "") && row.bondType === "农发债") return "农发清发债";
  return row.bondType || "利率债";
}
function dailyLead(rows: ParsedBondRecord[]) {
  const counts = new Map<string, number>();
  rows.forEach((row) => {
    const key = `${routeLabel(row)}|${reviewBondLabel(row)}`;
    counts.set(key, (counts.get(key) || 0) + 1);
  });
  const pieces = [...counts.entries()].map(([key, count]) => {
    const [route, type] = key.split("|");
    return `${route}${count}只${type}`;
  });
  return `今日${pieces.join("，")}，招标结果如下：`;
}
function reviewTable(rows: ParsedBondRecord[]) {
  const headers = ["代码", "期限", "发行量(亿元)", "中标利率", "前一日估值", "二级成交", "全场倍数", "边际倍数", "边际投标量", "边际中标量"];
  const widths = [1600, 550, 950, 850, 850, 900, 700, 700, 950, 950];
  return new Table({
    alignment: AlignmentType.CENTER, width: { size: 100, type: WidthType.PERCENTAGE }, layout: TableLayoutType.FIXED,
    rows: [
      new TableRow({ tableHeader: true, cantSplit: true, children: headers.map((value, index) => reportCell(value, { bold: true, fill: ORANGE, color: WHITE, size: 15, width: widths[index] })) }),
      ...rows.map((row) => new TableRow({ cantSplit: true, children: [
        `${row.bondCode || ""}${row.remark ? `（${row.remark}）` : ""}`,
        row.tenor, row.amount, displayRate(row, "中标利率"), displayRate(row, "前一日估值"),
        displayRate(row, "二级"), display(row, "全场倍数"), display(row, "边际倍数"),
        display(row, "边际投标量"), display(row, "边际中标量"),
      ].map((value, index) => reportCell(value, { size: 15, width: widths[index] })) })),
    ],
  });
}
function dailySection(date: string, rows: ParsedBondRecord[]) {
  const outer = new TableCell({
    borders: { top: noneBorder, bottom: noneBorder, left: noneBorder, right: noneBorder },
    margins: { top: 0, bottom: 0, left: 0, right: 0 },
    children: [
      new Paragraph({ keepNext: true, spacing: { before: 140, after: 0, line: 300 }, children: [new TextRun({ text: `${formatMd(date)} 回顾（${weekday(date)}）`, bold: true, size: 21, color: BLACK, font: FONT })] }),
      new Paragraph({ keepNext: true, spacing: { before: 0, after: 60, line: 300 }, children: [new TextRun({ text: dailyLead(rows), size: 20, color: BLACK, font: FONT })] }),
      reviewTable(rows),
    ],
  });
  return new Table({ width: { size: 100, type: WidthType.PERCENTAGE }, rows: [new TableRow({ cantSplit: true, children: [outer] })] });
}

export async function buildWeeklyReportBlob({
  weekStart, summary, localRecords, spreadRecords, previousSpreadRecords = [], ytdLocalRecords = localRecords,
}: ReportInput) {
  const weekEnd = fridayOf(weekStart);
  const currentAmount = amount(spreadRecords);
  const previousAmount = amount(previousSpreadRecords);
  const change = previousAmount ? (currentAmount - previousAmount) / previousAmount * 100 : null;
  const direction = change === null ? "" : change >= 0 ? `，较上周增加${Math.abs(change).toFixed(2)}%` : `，较上周减少${Math.abs(change).toFixed(2)}%`;
  const ytdReplacement = amount(ytdLocalRecords.filter((row) => localNature(row) === "置换债"));
  const summaryHeadings = /^(?:国债|政金债|中债招标|上清所|报价发行|DR浮息债|主题债(?:含绿色债)?)$/;
  const dailyDates = Array.from(new Set(spreadRecords.map((row) => row.tradeDate))).sort();
  const children: (Paragraph | Table)[] = [
    titleParagraph(`利率债发行周报${yyyymmdd(weekStart)}-${mmdd(weekEnd)}`),
    sectionHeading("本周利率债发行情况总结"),
    bodyParagraph(`本周共发${spreadRecords.length}期利率债，国债政金债发行总额达${text(currentAmount)}亿（不包含储蓄国债）${direction}。`),
    bodyParagraph(`本周地方债发行${text(amount(localRecords))}亿${ytdReplacement ? `；今年以来置换债已发行${text(ytdReplacement)}亿` : ""}。`),
    sectionHeading("地方债发行统计"),
    localSummaryTable(localRecords, ytdLocalRecords),
    new Paragraph({ alignment: AlignmentType.CENTER, keepNext: true, spacing: { before: 140, after: 80 }, children: [new TextRun({ text: `${yyyymmdd(weekStart)}-${yyyymmdd(weekEnd)}本周利率债发行情况表`, size: 21, color: BLACK, font: FONT })] }),
    weeklyPlanTable(weekStart, spreadRecords, localRecords),
    sectionHeading("周报发行小结"),
    ...summary.split(/\n+/).filter(Boolean).map((line) => {
      const isHeading = summaryHeadings.test(line.trim());
      return new Paragraph({ keepNext: isHeading, spacing: { before: isHeading ? 70 : 0, after: 0, line: 315 }, children: [new TextRun({ text: line, size: isHeading ? 20 : 19, bold: isHeading, color: BLACK, font: FONT })] });
    }),
    sectionHeading("本周发行回顾"),
    ...dailyDates.map((date) => dailySection(date, spreadRecords.filter((row) => row.tradeDate === date))),
  ];
  const document = new Document({
    styles: { default: { document: { run: { font: FONT, size: 21, color: BLACK }, paragraph: { spacing: { line: 315, before: 0, after: 0 } } } } },
    sections: [{
      properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 1440, right: 1800, bottom: 1440, left: 1800, header: 850, footer: 992 } } },
      children,
    }],
  });
  return Packer.toBlob(document);
}
