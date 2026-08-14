import {
  AlignmentType, Document, HeadingLevel, Packer, Paragraph, Table, TableCell, TableRow,
  TextRun, VerticalAlign, WidthType,
} from "docx";
import { fridayOf } from "./workbench.ts";
import type { ParsedBondRecord } from "./workbench.ts";

type ReportInput = {
  weekStart: string;
  summary: string;
  localRecords: ParsedBondRecord[];
  spreadRecords: ParsedBondRecord[];
};

const BLACK = "000000";
const HEADER_FILL = "F9E4D4";
const FONT = { ascii: "宋体", hAnsi: "宋体", eastAsia: "宋体", cs: "宋体", hint: "eastAsia" } as const;

function mmdd(date: string) { return date.slice(5).replace("-", ""); }
function formatMd(value: string) {
  const date = new Date(`${value}T12:00:00`);
  return `${date.getMonth() + 1}/${date.getDate()}`;
}
function weekday(value: string) {
  return ["周日", "周一", "周二", "周三", "周四", "周五", "周六"][new Date(`${value}T12:00:00`).getDay()];
}
function text(value: unknown) {
  if (value === null || value === undefined || value === "") return "-";
  if (typeof value === "number") return String(Math.round(value * 10000) / 10000);
  return String(value);
}
function display(row: ParsedBondRecord, key: string) {
  const map = row.raw?.__display;
  return map && typeof map === "object" ? text((map as Record<string, unknown>)[key]) : "-";
}
function cell(value: unknown, bold = false, fill?: string) {
  return new TableCell({
    verticalAlign: VerticalAlign.CENTER,
    shading: fill ? { fill } : undefined,
    children: [new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 40, after: 40, line: 240 },
      children: [new TextRun({ text: text(value), bold, color: BLACK, size: 17, font: FONT })],
    })],
  });
}
function heading(textValue: string, size = 28) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_1,
    keepNext: true,
    spacing: { before: 180, after: 100 },
    children: [new TextRun({ text: textValue, bold: true, color: BLACK, size, font: FONT })],
  });
}
function routeLabel(row: ParsedBondRecord) {
  const route = row.summaryMeta?.route || row.issuanceRoute || "中债招标";
  return route === "报价发行" ? "报价发行" : route === "上清所" ? "上清所发行" : "招标发行";
}
function dailyLead(rows: ParsedBondRecord[]) {
  const counts = new Map<string, number>();
  rows.forEach((row) => {
    const key = `${routeLabel(row)}|${row.bondType || "利率债"}`;
    counts.set(key, (counts.get(key) || 0) + 1);
  });
  const pieces = [...counts.entries()].map(([key, count]) => {
    const [route, type] = key.split("|");
    return `${route}${count}只${type}`;
  });
  return `今日${pieces.join("，")}，发行结果如下：`;
}

export async function buildWeeklyReportBlob({ weekStart, summary, localRecords, spreadRecords }: ReportInput) {
  const weekEnd = fridayOf(weekStart);
  const summaryHeadings = /^(?:国债|政金债|中债招标|上清所|报价发行|DR浮息债|主题债(?:含绿色债)?)$/;
  const localHeaders = ["日期", "代码", "简称", "期限", "发行量（亿元）", "下限", "场所"];
  const reviewHeaders = ["代码", "期限", "发行量\n(亿元)", "中标利率", "前一日估值", "二级成交", "综收比二级\n(BP)", "全场倍数", "边际倍数", "边际投标量", "边际中标量"];
  const dailyDates = Array.from(new Set(spreadRecords.map(row => row.tradeDate))).sort();
  const children: (Paragraph | Table)[] = [
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { after: 180 },
      children: [new TextRun({ text: `利率债发行周报${mmdd(weekStart)}-${mmdd(weekEnd)}`, bold: true, size: 36, color: BLACK, font: FONT })],
    }),
    heading("周报发行小结"),
    ...summary.split(/\n+/).filter(Boolean).map((line) => {
      const isHeading = summaryHeadings.test(line.trim());
      return new Paragraph({
        keepNext: isHeading,
        spacing: { before: isHeading ? 100 : 0, after: isHeading ? 60 : 100, line: 300 },
        children: [new TextRun({ text: line, size: isHeading ? 23 : 21, bold: isHeading, color: BLACK, font: FONT })],
      });
    }),
    heading("本周地方债发行明细"),
    new Table({
      width: { size: 100, type: WidthType.PERCENTAGE },
      rows: [
        new TableRow({ tableHeader: true, cantSplit: true, children: localHeaders.map(value => cell(value, true, HEADER_FILL)) }),
        ...localRecords.map(row => new TableRow({ cantSplit: true, children: [row.tradeDate, row.bondCode, row.shortName, row.tenor, row.amount, row.floorRate, row.venue].map(value => cell(value)) })),
      ],
    }),
    heading("本周发行回顾"),
  ];
  dailyDates.forEach((date) => {
    const rows = spreadRecords.filter(row => row.tradeDate === date);
    children.push(new Paragraph({
      pageBreakBefore: true,
      keepNext: true,
      spacing: { before: 160, after: 70 },
      children: [new TextRun({ text: `${formatMd(date)} 回顾（${weekday(date)}）`, bold: true, size: 24, color: BLACK, font: FONT })],
    }));
    children.push(new Paragraph({
      keepNext: true,
      spacing: { after: 80, line: 280 },
      children: [new TextRun({ text: dailyLead(rows), size: 20, color: BLACK, font: FONT })],
    }));
    children.push(new Table({
      width: { size: 100, type: WidthType.PERCENTAGE },
      rows: [
        new TableRow({ tableHeader: true, cantSplit: true, children: reviewHeaders.map(value => cell(value, true, HEADER_FILL)) }),
        ...rows.map(row => new TableRow({ cantSplit: true, children: [
          `${row.bondCode || ""}${row.remark ? `（${row.remark}）` : ""}`,
          row.tenor, row.amount, display(row, "中标利率"), display(row, "前一日估值"),
          display(row, "二级"), row.summaryMeta?.displaySpreadText || row.spread,
          display(row, "全场倍数"), display(row, "边际倍数"),
          display(row, "边际投标量"), display(row, "边际中标量"),
        ].map(value => cell(value)) })),
      ],
    }));
  });
  const document = new Document({ sections: [{ properties: {}, children }] });
  return Packer.toBlob(document);
}
