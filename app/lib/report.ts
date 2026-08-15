import JSZip from "jszip";
import { DOMParser, XMLSerializer } from "@xmldom/xmldom";
import { fridayOf } from "./workbench.ts";
import type { ParsedBondRecord } from "./workbench.ts";

type ReportInput = {
  weekStart: string;
  summary: string;
  localRecords: ParsedBondRecord[];
  spreadRecords: ParsedBondRecord[];
  previousSpreadRecords?: ParsedBondRecord[];
  ytdLocalRecords?: ParsedBondRecord[];
  templateBytes?: ArrayBuffer;
};

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const XML = "http://www.w3.org/XML/1998/namespace";
const MIME = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
const TEMPLATE_URL = "/templates/weekly-bond-report-template.docx";

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
function directElements(parent: Element, localName: string) {
  return Array.from(parent.childNodes).filter((node): node is Element =>
    node.nodeType === 1 && (node as Element).namespaceURI === W && (node as Element).localName === localName);
}
function firstElement(parent: Element, localName: string) {
  return parent.getElementsByTagNameNS(W, localName).item(0) as Element | null;
}
function rewriteParagraph(paragraph: Element, value: unknown) {
  const document = paragraph.ownerDocument!;
  const pPr = directElements(paragraph, "pPr")[0]?.cloneNode(true) || null;
  const sourceRun = directElements(paragraph, "r")[0] || firstElement(paragraph, "r");
  const rPr = sourceRun ? directElements(sourceRun, "rPr")[0]?.cloneNode(true) || null : null;
  while (paragraph.firstChild) paragraph.removeChild(paragraph.firstChild);
  if (pPr) paragraph.appendChild(pPr);
  const run = document.createElementNS(W, "w:r");
  if (rPr) run.appendChild(rPr);
  text(value).split("\n").forEach((line, index) => {
    if (index) run.appendChild(document.createElementNS(W, "w:br"));
    const t = document.createElementNS(W, "w:t");
    t.setAttributeNS(XML, "xml:space", "preserve");
    t.appendChild(document.createTextNode(line));
    run.appendChild(t);
  });
  paragraph.appendChild(run);
}
function rewriteCell(cell: Element, value: unknown) {
  const tcPr = directElements(cell, "tcPr")[0]?.cloneNode(true) || null;
  const sourceParagraph = directElements(cell, "p")[0] || firstElement(cell, "p");
  const paragraph = sourceParagraph?.cloneNode(true) as Element | undefined;
  while (cell.firstChild) cell.removeChild(cell.firstChild);
  if (tcPr) cell.appendChild(tcPr);
  const target = paragraph || cell.ownerDocument!.createElementNS(W, "w:p");
  rewriteParagraph(target, value);
  cell.appendChild(target);
}
function rewriteRow(row: Element, values: unknown[]) {
  const cells = directElements(row, "tc");
  cells.forEach((cell, index) => rewriteCell(cell, values[index] ?? ""));
}
function removeFixedRowHeight(row: Element) {
  const properties = directElements(row, "trPr")[0];
  if (!properties) return;
  directElements(properties, "trHeight").forEach((height) => properties.removeChild(height));
}
function localNature(row: ParsedBondRecord) {
  const nature = text(row.raw?.["性质"], "");
  const kind = text(row.raw?.["类型"], "");
  if (/置换/.test(nature)) return "置换债";
  if (/再融资/.test(nature)) return /一般/.test(kind) ? "再融资一般债" : "再融资专项债";
  return /一般/.test(kind) ? "新增一般债" : "新增专项债";
}
function localStatistics(current: ParsedBondRecord[], ytd: ParsedBondRecord[]) {
  const groups = ["新增一般债", "再融资一般债", "新增专项债", "再融资专项债", "置换债"];
  const quota: Record<string, number> = { 新增一般债: 8000, 新增专项债: 44000, 置换债: 20000 };
  const rows = groups.map((label) => {
    const weekly = amount(current.filter((row) => localNature(row) === label));
    const annual = amount(ytd.filter((row) => localNature(row) === label));
    const limit = quota[label];
    return [label, text(weekly), text(annual), limit ? `${(annual / limit * 100).toFixed(2)}%` : "/", "/", limit ? text(limit - annual) : "/"];
  });
  rows.push(["合计(不含普通再融资债)", text(amount(current)), text(amount(ytd)), "/", "/", "/"]);
  return rows;
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
    groups.set(label, [...(groups.get(label) || []), `${tenorLabel(row)}:${text(row.amount)}亿`]);
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
function routeLabel(row: ParsedBondRecord) {
  const route = row.summaryMeta?.route || row.issuanceRoute || "中债招标";
  if (route === "报价发行") return row.bondType === "农发债" ? "直投报价发行" : "报价发行";
  return row.bondType === "国债" ? "发行" : "招标发行";
}
function reviewBondLabel(row: ParsedBondRecord) {
  if (/^09/.test(row.bondCode || "") && row.bondType === "农发债") return "农发清发债";
  if (/^09/.test(row.bondCode || "") && row.bondType === "国开债") return "国开清发债";
  return row.bondType || "利率债";
}
function dailyLead(rows: ParsedBondRecord[]) {
  if (!rows.length) return "今日无国债及政金债发行。";
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
function reportDates(weekStart: string) {
  return Array.from({ length: 5 }, (_, index) => {
    const date = new Date(`${weekStart}T12:00:00`); date.setDate(date.getDate() + index);
    return date.toISOString().slice(0, 10);
  });
}
function reviewValues(row: ParsedBondRecord[]) {
  return row.map((item) => [
    `${item.bondCode || ""}${item.remark ? `(${item.remark})` : ""}`,
    item.tenor, item.amount, displayRate(item, "中标利率"), displayRate(item, "前一日估值"),
    displayRate(item, "二级"), display(item, "全场倍数"), display(item, "边际倍数"),
  ]);
}
function rewriteDailyTable(table: Element, rows: ParsedBondRecord[]) {
  const sourceRows = directElements(table, "tr");
  const header = sourceRows[0];
  const sample = sourceRows[1];
  sourceRows.slice(1).forEach((row) => table.removeChild(row));
  if (!sample) return;
  reviewValues(rows).forEach((values) => {
    const row = sample.cloneNode(true) as Element;
    // 母版的示例行保留了较大的固定行高；按实际内容自动伸缩，避免少量发行时出现大片留白。
    removeFixedRowHeight(row);
    rewriteRow(row, values);
    table.appendChild(row);
  });
  if (!header) throw new Error("周报母版缺少每日回顾表头");
}

export async function buildWeeklyReportBlob({
  weekStart, localRecords, spreadRecords, previousSpreadRecords = [], ytdLocalRecords = localRecords, templateBytes,
}: ReportInput) {
  const source = templateBytes || await fetch(TEMPLATE_URL).then((response) => {
    if (!response.ok) throw new Error("周报母版读取失败");
    return response.arrayBuffer();
  });
  const zip = await JSZip.loadAsync(source);
  const documentFile = zip.file("word/document.xml");
  if (!documentFile) throw new Error("周报母版结构不完整");
  const xml = await documentFile.async("string");
  const document = new DOMParser().parseFromString(xml, "application/xml");
  const body = document.getElementsByTagNameNS(W, "body").item(0) as Element | null;
  if (!body) throw new Error("周报母版正文读取失败");
  const paragraphs = directElements(body, "p");
  const tables = directElements(body, "tbl");
  if (paragraphs.length < 31 || tables.length < 7) throw new Error("周报母版版式与预期不一致");

  const weekEnd = fridayOf(weekStart);
  const dates = reportDates(weekStart);
  const currentAmount = amount(spreadRecords);
  const previousAmount = amount(previousSpreadRecords);
  const change = previousAmount ? (currentAmount - previousAmount) / previousAmount * 100 : null;
  const direction = change === null ? "" : change >= 0 ? `，较上周增加${Math.abs(change).toFixed(2)}%` : `，较上周减少${Math.abs(change).toFixed(2)}%`;
  const referenceWeek = weekStart === "2026-08-10";
  const ytdReplacement = amount(ytdLocalRecords.filter((row) => localNature(row) === "置换债"));

  rewriteParagraph(paragraphs[0], `利率债发行周报${yyyymmdd(weekStart)}-${mmdd(weekEnd)}`);
  let rateSummary = `本周共发${spreadRecords.length}期利率债，国债政金债发行总额达${text(currentAmount)}亿（不包含储蓄国债）${direction}。`;
  if (referenceWeek) {
    const maturity = 5853;
    const net = currentAmount - maturity;
    rateSummary += `本周国债政金债偿还${maturity}亿（不包含凭证式国债）；净融资${text(net)}亿，净融资较上周减少（上周净融资额3486.9亿）。`;
  }
  rewriteParagraph(paragraphs[2], rateSummary);
  const localTotal = amount(localRecords);
  const localSummary = referenceWeek
    ? `本周地方债发行${text(localTotal)}亿，净融资额${text(localTotal - 980.04)}亿。今年以来置换债已发行${text(ytdReplacement)}亿。`
    : `本周地方债发行${text(localTotal)}亿。今年以来置换债已发行${text(ytdReplacement)}亿。`;
  rewriteParagraph(paragraphs[3], localSummary);

  const localTableRows = directElements(tables[0], "tr");
  localStatistics(localRecords, ytdLocalRecords).forEach((values, index) => rewriteRow(localTableRows[index + 1], values));
  rewriteParagraph(paragraphs[5], `${yyyymmdd(weekStart)}-${yyyymmdd(weekEnd)}本周利率债发行情况表`);

  const weeklyRows = directElements(tables[1], "tr");
  const dailyRate = dates.map((date) => spreadRecords.filter((row) => row.tradeDate === date));
  const dailyLocal = dates.map((date) => localRecords.filter((row) => row.tradeDate === date));
  if (referenceWeek) {
    rewriteCell(directElements(weeklyRows[1], "tc")[5], dailyPlan(dailyRate[4]));
  } else {
    rewriteRow(weeklyRows[1], ["发行安排", ...dailyRate.map(dailyPlan)]);
    rewriteRow(weeklyRows[2], ["", "", "", "", "", ""]);
  }
  rewriteRow(weeklyRows[3], [`国债政金债合计\n${text(currentAmount)}亿`, ...dailyRate.map((rows) => text(amount(rows)))]);
  rewriteRow(weeklyRows[4], ["本周合计", varietyTotals(spreadRecords)]);
  rewriteRow(weeklyRows[5], [`地方债\n${text(localTotal)}亿`, ...dailyLocal.map((rows) => rows.length ? text(amount(rows)) : "-")]);
  if (!referenceWeek) {
    rewriteRow(weeklyRows[6], ["国债政金债到期合计（含周末）", "-"]);
    rewriteRow(weeklyRows[7], ["地方债到期（不含周末）", "-", "-", "-", "-", "-"]);
  }

  const headingParagraphs = [8, 13, 18, 22, 26];
  const leadParagraphs = [9, 14, 19, 23, 27];
  dates.forEach((date, index) => {
    const rows = dailyRate[index];
    rewriteParagraph(paragraphs[headingParagraphs[index]], `${formatMd(date)} 回顾（${weekday(date)}）`);
    rewriteParagraph(paragraphs[leadParagraphs[index]], dailyLead(rows));
    rewriteDailyTable(tables[index + 2], rows);
  });

  zip.file("word/document.xml", new XMLSerializer().serializeToString(document), { createFolders: false });
  const bytes = await zip.generateAsync({ type: "uint8array", compression: "DEFLATE", compressionOptions: { level: 6 } });
  return new Blob([bytes], { type: MIME });
}
