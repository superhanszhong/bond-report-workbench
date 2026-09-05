import JSZip from "jszip";
import { DOMParser, XMLSerializer } from "@xmldom/xmldom";
import { fridayOf, usesDrPrice, spreadValueForMetric, basisPointValue } from "./workbench.ts";
import type { ParsedBondRecord } from "./workbench.ts";
import { consolidateBondRecords, recordKey } from "./record-merge.ts";
import { baseBondCode, bondIssueKey } from "./bond-code.ts";

type ReportInput = {
  weekStart: string;
  summary: string;
  localRecords: ParsedBondRecord[];
  spreadRecords: ParsedBondRecord[];
  scheduleRecords?: ParsedBondRecord[];
  previousSpreadRecords?: ParsedBondRecord[];
  ytdLocalRecords?: ParsedBondRecord[];
  templateBytes?: ArrayBuffer;
  maturity?: {
    rateTotal: number;
    rateBreakdown: string;
    localDaily: Record<string, number>;
    localTotal?: number;
    previousRateNet?: number;
  };
};

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const XML = "http://www.w3.org/XML/1998/namespace";
const MIME = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
const TEMPLATE_URL = `${import.meta.env?.BASE_URL || "/"}templates/weekly-bond-report-template.docx`;

// 已经人工核定的历史口径，用于连续周报在上周明细不完整时仍能准确衔接。
// 新一周仍优先根据当周一二级发行量和到期 Excel 计算当周净融资。
const VERIFIED_RATE_FINANCING: Record<string, {
  maturity: number;
  net: number;
  previousNet: number;
}> = {
  "2026-08-10": { maturity: 5853, net: -1643, previousNet: 3486.9 },
};

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
function ensureChild(parent: Element, localName: string, first = false) {
  const existing = directElements(parent, localName)[0];
  if (existing) return existing;
  const child = parent.ownerDocument!.createElementNS(W, `w:${localName}`);
  if (first && parent.firstChild) parent.insertBefore(child, parent.firstChild);
  else parent.appendChild(child);
  return child;
}
function setParagraphFlag(paragraph: Element, localName: string) {
  const properties = ensureChild(paragraph, "pPr", true);
  ensureChild(properties, localName);
}
function removeParagraphFlag(paragraph: Element, localName: string) {
  const properties = directElements(paragraph, "pPr")[0];
  if (!properties) return;
  directElements(properties, localName).forEach((node) => properties.removeChild(node));
}
function resetParagraphIndent(paragraph: Element) {
  const properties = ensureChild(paragraph, "pPr", true);
  const indent = ensureChild(properties, "ind");
  indent.setAttributeNS(W, "w:left", "0");
  indent.setAttributeNS(W, "w:right", "0");
  indent.setAttributeNS(W, "w:firstLine", "0");
  indent.removeAttributeNS(W, "hanging");
}
function setRowCantSplit(row: Element) {
  const properties = ensureChild(row, "trPr", true);
  ensureChild(properties, "cantSplit");
}
function keepTableTogether(table: Element) {
  const rows = directElements(table, "tr");
  rows.forEach((row, rowIndex) => {
    setRowCantSplit(row);
    const properties = ensureChild(row, "trPr", true);
    if (rowIndex === 0) ensureChild(properties, "tblHeader");
    if (rowIndex === rows.length - 1) return;
    directElements(row, "tc").forEach((cell) => {
      directElements(cell, "p").forEach((paragraph) => setParagraphFlag(paragraph, "keepNext"));
    });
  });
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
  rows.push(["合计", text(amount(current)), text(amount(ytd)), "/", "/", "/"]);
  return rows;
}
function tenorLabel(row: ParsedBondRecord) {
  const raw = text(row.tenor, "");
  const unit = /[DMY]$/i.test(raw) ? raw.toUpperCase() : `${raw}Y`;
  return /浮息/.test(row.summaryMeta?.rateType || row.remark || "") ? `${unit}浮息` : unit;
}
function tenorYears(value = "") {
  const match = value.trim().match(/^(\d+(?:\.\d+)?)\s*(D|M|Y|天|月|年)?(?:\s*浮息(?:债)?)?$/i);
  if (!match) return Number.POSITIVE_INFINITY;
  const unit = (match[2] || "Y").toUpperCase();
  return Number(match[1]) / (unit === "D" || unit === "天" ? 365 : unit === "M" || unit === "月" ? 12 : 1);
}
export function sortReportBonds(rows: ParsedBondRecord[], groupBy = (row: ParsedBondRecord) => row.bondType || "利率债") {
  // Preserve the first-seen variety order and original order of equal tenors; never reorder stored records.
  const groups = new Map<string, ParsedBondRecord[]>();
  rows.forEach(row => {
    const key = groupBy(row);
    if (!groups.has(key)) groups.set(key, []);
    groups.get(key)!.push(row);
  });
  return [...groups.values()].flatMap(group => group.sort((a, b) => {
    const left = tenorYears(a.tenor), right = tenorYears(b.tenor);
    return left === right ? 0 : left < right ? -1 : 1;
  }));
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
export function planSession(row: ParsedBondRecord): "上午" | "下午" {
  if (row.summaryMeta?.route === "报价发行" || row.issuanceRoute === "报价发行" || /清发|报价/.test(row.remark || "")) return "上午";
  const hour = Number.parseInt((row.bidTime || "").match(/\d{1,2}/)?.[0] || "", 10);
  if (Number.isFinite(hour)) return hour >= 12 ? "下午" : "上午";
  const day = new Date(`${row.tradeDate}T12:00:00`).getDay();
  if (row.bondType === "农发债") return "下午";
  if (row.bondType === "国开债" && day === 4) return "下午";
  return "上午";
}
function dailyPlan(rows: ParsedBondRecord[], session: "上午" | "下午") {
  rows = sortReportBonds(rows.filter((row) => planSession(row) === session), planLabel);
  if (!rows.length) return "-";
  const groups = new Map<string, string[]>();
  rows.forEach((row) => {
    const label = planLabel(row);
    groups.set(label, [...(groups.get(label) || []), `${tenorLabel(row)}:${text(row.amount)}亿`]);
  });
  return [...groups].map(([label, items]) => `${label}\n${items.join("\n")}`).join("\n");
}
function basePlanCode(code = "") { return baseBondCode(code); }
export function mergeIssuanceSessions(rows: ParsedBondRecord[], schedules: ParsedBondRecord[]) {
  if (!schedules.length) return rows;
  const scheduleByKey = new Map(schedules.map((row) => [recordKey(row), row]));
  return rows.map((row) => {
    const schedule = scheduleByKey.get(recordKey(row));
    return schedule?.bidTime ? { ...row, bidTime: schedule.bidTime } : row;
  });
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
  if (route === "报价发行" || /报价发行/.test(row.remark || "")) return row.bondType === "农发债" ? "直投报价发行" : "报价发行";
  if (/直投招标/.test(row.remark || "")) return "直投招标发行";
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
function previousWeekStart(weekStart: string) {
  const date = new Date(`${weekStart}T12:00:00`);
  date.setDate(date.getDate() - 7);
  return date.toISOString().slice(0, 10);
}
function reviewValues(row: ParsedBondRecord[]) {
  return row.map((item) => [
    `${item.bondCode || ""}${item.remark ? `(${item.remark})` : ""}`,
    item.tenor, item.amount, reportWinningResult(item), reportBenchmark(item, "前一日估值"),
    reportBenchmark(item, "二级"), display(item, "全场倍数"), display(item, "边际倍数"),
  ]);
}

function usesDrNetPrice(row: ParsedBondRecord) {
  return usesDrPrice(row);
}

function numericQuote(value: string) {
  if (!/^\s*[-+]?\d+(?:\.\d+)?\s*[%元]?\s*$/.test(value)) return null;
  return Number.parseFloat(value);
}

function netPriceQuote(value: string) {
  // Older DR history stores the payment price in the all-in column. A yield must not be labelled as yuan.
  const numeric = numericQuote(value);
  return numeric !== null && (/元/.test(value) || numeric >= 50) ? value.replace(/元$/, "") + "元" : null;
}

function reportWinningResult(row: ParsedBondRecord) {
  if (!usesDrNetPrice(row)) return displayRate(row, "中标利率");
  const explicit = display(row, "中标净价");
  const rawPrice = row.raw?.["中标净价"];
  const price = numericQuote(explicit) !== null ? explicit.replace(/元$/, "") + "元"
    : typeof rawPrice === "number" ? `${rawPrice}元`
      : netPriceQuote(display(row, "综收")) || netPriceQuote(row.summaryMeta?.allInText || "");
  return price ? `净价 ${price}` : "净价未提供";
}

function reportBenchmark(row: ParsedBondRecord, field: string) {
  const value = display(row, field);
  if (!usesDrNetPrice(row) || numericQuote(value) === null) return displayRate(row, field);
  const price = netPriceQuote(value);
  return price ? `净价 ${price}` : `${value.replace(/%$/, "")}%`;
}

export function reportDataWarnings(rows: ParsedBondRecord[], schedules: ParsedBondRecord[]) {
  rows = consolidateBondRecords(rows);
  schedules = consolidateBondRecords(schedules);
  const warnings: string[] = [];
  // X/Z are interchangeable for the same reopening; a bare suffix is reopening 1.
  const comparisonCode = bondIssueKey;
  for (const row of rows) {
    if (usesDrNetPrice(row) && reportWinningResult(row).includes("净价未提供")) {
      warnings.push(`${row.bondCode} 缺少中标净价，请在一二级表补充“中标净价”列后上传；周报将标注未提供，不以收益率代替。`);
    }
    const sameIssue = schedules.filter(plan => plan.tradeDate === row.tradeDate && basePlanCode(plan.bondCode) === basePlanCode(row.bondCode));
    if (sameIssue.length && !sameIssue.some(plan => comparisonCode(plan.bondCode) === comparisonCode(row.bondCode))) {
      warnings.push(`${row.tradeDate} 增发期次待核对：一二级表为 ${row.bondCode}，发行计划为 ${sameIssue.map(plan => plan.bondCode).join(" / ")}。保留一二级表原值，不自动改写增发期数。`);
    }
    const exact = schedules.find(plan => recordKey(plan) === recordKey(row));
    if (!exact && !sameIssue.length && schedules.length) warnings.push(`${row.tradeDate} ${row.bondCode}：一二级已收录，上传的发行计划未列示；仍采用一二级 ${text(row.amount)} 亿，由你核定。`);
    if (exact && row.amount != null && exact.amount != null && Math.abs(row.amount - exact.amount) > 0.00001) {
      warnings.push(`${row.tradeDate} ${row.bondCode}：一二级发行量 ${text(row.amount)} 亿，发行计划 ${text(exact.amount)} 亿（可能含上弹或追加等口径差异）；周报采用一二级 ${text(row.amount)} 亿，由你核定。`);
    }
    if (exact && !usesDrPrice(row)) {
      const actual = reportYield(row.raw?.中标利率 ?? row.summaryMeta?.winningRateText);
      const planned = reportYield(exact.raw?.加权利率 ?? exact.raw?.发行利率);
      if (actual !== null && planned !== null && Math.abs(actual - planned) > 0.00005) warnings.push(`${row.tradeDate} ${row.bondCode}：一二级中标利率 ${text(actual)}%，发行计划 ${text(planned)}%；采用一二级，由你核定。`);
    }
    const raw = row.raw || {};
    if (usesDrPrice(row)) {
      const winningPrice = numericQuote(reportWinningResult(row).replace(/^净价\s*/, ""));
      const secondaryPrice = numericQuote(display(row, "二级"));
      const description = String(raw.招标利差 || row.summaryMeta?.auctionSpreadText || "");
      const stated = description.match(/^\s*([+-]?\d+(?:\.\d+)?)\s*元/) || description.match(/\)\s*([+-]?\d+(?:\.\d+)?)\s*元/);
      if (winningPrice !== null && secondaryPrice !== null && stated && Math.abs(Math.abs(winningPrice-secondaryPrice)-Math.abs(Number(stated[1]))) > 0.00005) warnings.push(`${row.tradeDate} ${row.bondCode}：净价差复算为 ${text(winningPrice-secondaryPrice)} 元，与一二级文字描述不一致；保留原始净价，由你核定。`);
    } else {
      const winning = reportYield(raw.中标利率 ?? row.summaryMeta?.winningRateText);
      const secondary = reportYield(raw.二级 ?? row.summaryMeta?.secondaryText);
      const supplied = spreadValueForMetric(row, "winning");
      if (winning !== null && secondary !== null && supplied !== null && Math.abs((winning-secondary)*100-supplied)>0.05) warnings.push(`${row.tradeDate} ${row.bondCode}：一二级中标利差 ${text(supplied)} BP，与中标率减二级复算 ${text((winning-secondary)*100)} BP 不一致，由你核定。`);
      const allIn = reportYield(raw.综收);
      const suppliedAllIn = basisPointValue(raw.利差);
      if (allIn !== null && secondary !== null && suppliedAllIn !== null && Math.abs((allIn-secondary)*100-suppliedAllIn)>0.05) warnings.push(`${row.tradeDate} ${row.bondCode}：一二级综收利差文字与收益率复算不一致，图表使用同表收益率复算值 ${text((allIn-secondary)*100)} BP，由你核定。`);
    }
  }
  const actualKeys = new Set(rows.map(recordKey));
  for (const plan of schedules) {
    if (!actualKeys.has(recordKey(plan)) && !rows.some(row => row.tradeDate === plan.tradeDate && basePlanCode(row.bondCode) === basePlanCode(plan.bondCode))) warnings.push(`${plan.tradeDate} ${plan.bondCode} ${plan.shortName || ""}：发行计划有 ${text(plan.amount)} 亿，一二级未收录（可能为特殊小团等口径）；本周不纳入，由你决定是否补入一二级表。`);
  }
  return warnings;
}

function reportYield(value: unknown) {
  const s = String(value ?? "").trim();
  if (!s || /元|净价|--/.test(s)) return null;
  const n = Number(s.replace(/%$/, ""));
  return Number.isFinite(n) && Math.abs(n) < 50 ? (s.endsWith("%") ? n : Math.abs(n) < 0.2 ? n*100 : n) : null;
}
function rewriteDailyTable(table: Element, rows: ParsedBondRecord[]) {
  const sourceRows = directElements(table, "tr");
  const header = sourceRows[0];
  const sample = sourceRows[1];
  sourceRows.slice(1).forEach((row) => table.removeChild(row));
  if (!sample) return;
  if (header && rows.some(usesDrNetPrice)) rewriteRow(header, ["代码", "期限", "发行量(亿元)", "中标结果", "前一日估值", "二级成交", "全场倍数", "边际倍数"]);
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
  weekStart, localRecords, spreadRecords, scheduleRecords = [], previousSpreadRecords = [], ytdLocalRecords = localRecords, templateBytes, maturity,
}: ReportInput) {
  spreadRecords = consolidateBondRecords(spreadRecords);
  previousSpreadRecords = consolidateBondRecords(previousSpreadRecords);
  localRecords = consolidateBondRecords(localRecords);
  ytdLocalRecords = consolidateBondRecords(ytdLocalRecords);
  scheduleRecords = consolidateBondRecords(scheduleRecords);
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
  for (const row of [...spreadRecords, ...localRecords, ...ytdLocalRecords]) {
    if (row.amount === null || row.amount === undefined || !Number.isFinite(row.amount) || row.amount < 0) throw new Error(`${row.tradeDate} ${row.bondCode} 发行量缺失或无效，请更新对应底稿`);
  }
  const ytdByKey = new Map(ytdLocalRecords.map(row => [recordKey(row), row]));
  for (const row of localRecords) {
    const annual = ytdByKey.get(recordKey(row));
    if (!annual || Math.abs((annual.amount || 0) - (row.amount || 0)) > 0.00001 || localNature(annual) !== localNature(row)) throw new Error(`地方债 ${row.bondCode} 本周与年度明细不一致，请刷新或重新上传完整底稿`);
  }
  const previousAmount = amount(previousSpreadRecords);
  const change = previousAmount ? (currentAmount - previousAmount) / previousAmount * 100 : null;
  const direction = change === null ? "" : change >= 0 ? `，较上周增加${Math.abs(change).toFixed(2)}%` : `，较上周减少${Math.abs(change).toFixed(2)}%`;
  const verifiedCurrent = VERIFIED_RATE_FINANCING[weekStart];
  const verifiedPrevious = VERIFIED_RATE_FINANCING[previousWeekStart(weekStart)];
  const referenceWeek = Boolean(verifiedCurrent);
  const ytdReplacement = amount(ytdLocalRecords.filter((row) => localNature(row) === "置换债"));

  rewriteParagraph(paragraphs[0], `利率债发行周报${mmdd(weekStart)}-${mmdd(weekEnd)}`);
  let rateSummary = `本周共发${spreadRecords.length}期利率债，国债政金债发行总额达${text(currentAmount)}亿（不包含储蓄国债）${direction}。`;
  if (maturity) {
    const rateMaturity = verifiedCurrent?.maturity ?? maturity.rateTotal;
    const net = verifiedCurrent?.net ?? currentAmount - rateMaturity;
    const priorNet = verifiedCurrent?.previousNet ?? verifiedPrevious?.net ?? maturity.previousRateNet;
    const netDirection = priorNet === undefined ? "" : net >= priorNet ? "增加" : "减少";
    const comparison = priorNet === undefined ? "" : !verifiedCurrent && verifiedPrevious
      ? `，上周净融资额${text(priorNet)}亿为已核定总偿还口径，不作直接比较`
      : `，净融资较上周${netDirection}（上周${verifiedCurrent ? "净融资额" : "到期口径测算净融资额"}${text(priorNet)}亿）`;
    rateSummary += verifiedCurrent
      ? `本周国债政金债偿还${text(rateMaturity)}亿（不包含凭证式国债）；净融资${text(net)}亿${comparison}。`
      : `本周国债政金债到期${text(rateMaturity)}亿（不包含凭证式国债）；按到期明细测算，净融资${text(net)}亿${comparison}。`;
  } else if (verifiedCurrent) {
    const netDirection = verifiedCurrent.net >= verifiedCurrent.previousNet ? "增加" : "减少";
    rateSummary += `本周国债政金债偿还${text(verifiedCurrent.maturity)}亿（不包含凭证式国债）；净融资${text(verifiedCurrent.net)}亿，净融资较上周${netDirection}（上周净融资额${text(verifiedCurrent.previousNet)}亿）。`;
  }
  rewriteParagraph(paragraphs[2], rateSummary);
  const localTotal = amount(localRecords);
  const localMaturityTotal = maturity
    ? maturity.localTotal ?? Object.values(maturity.localDaily).reduce((sum, value) => sum + value, 0)
    : referenceWeek ? 980.04 : null;
  const localSummary = localMaturityTotal === null
    ? `本周地方债发行${text(localTotal)}亿。今年以来置换债已发行${text(ytdReplacement)}亿。`
    : `本周地方债发行${text(localTotal)}亿，按到期明细测算，净融资额${text(localTotal - localMaturityTotal)}亿。今年以来置换债已发行${text(ytdReplacement)}亿。`;
  rewriteParagraph(paragraphs[3], localSummary);

  const localTableRows = directElements(tables[0], "tr");
  localStatistics(localRecords, ytdLocalRecords).forEach((values, index) => rewriteRow(localTableRows[index + 1], values));
  rewriteParagraph(paragraphs[5], `${yyyymmdd(weekStart)}-${yyyymmdd(weekEnd)}本周利率债发行情况表`);

  const weeklyRows = directElements(tables[1], "tr");
  const dailyRate = dates.map((date) => sortReportBonds(spreadRecords.filter((row) => row.tradeDate === date)));
  // 一二级表是发行量与债券明细的唯一口径；发行计划表只补充上午/下午时段。
  const plannedRate = mergeIssuanceSessions(spreadRecords, scheduleRecords);
  const dailyPlannedRate = dates.map((date) => plannedRate.filter((row) => row.tradeDate === date));
  const dailyLocal = dates.map((date) => localRecords.filter((row) => row.tradeDate === date));
  rewriteRow(weeklyRows[1], ["上午", ...dailyPlannedRate.map((rows) => dailyPlan(rows, "上午"))]);
  rewriteRow(weeklyRows[2], ["下午", ...dailyPlannedRate.map((rows) => dailyPlan(rows, "下午"))]);
  rewriteRow(weeklyRows[3], [`国债政金债合计\n${text(amount(plannedRate))}亿`, ...dailyPlannedRate.map((rows) => text(amount(rows)))]);
  rewriteRow(weeklyRows[4], ["本周合计", varietyTotals(plannedRate)]);
  rewriteRow(weeklyRows[5], [`地方债\n${text(localTotal)}亿`, ...dailyLocal.map((rows) => rows.length ? text(amount(rows)) : "-")]);
  if (maturity) {
    rewriteRow(weeklyRows[6], [`国债政金债到期合计（含周末）\n${text(maturity.rateTotal)}亿`, maturity.rateBreakdown]);
    rewriteRow(weeklyRows[7], [`地方债到期（不含周末）\n${text(localMaturityTotal)}亿`, ...dates.map((date) => maturity.localDaily[date] ? text(maturity.localDaily[date]) : "-")]);
  } else if (!referenceWeek) {
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
    // 每日回顾连续排版；仅在剩余空间不足时让标题、导语和整张表自然移至下一页。
    resetParagraphIndent(paragraphs[headingParagraphs[index]]);
    resetParagraphIndent(paragraphs[leadParagraphs[index]]);
    removeParagraphFlag(paragraphs[headingParagraphs[index]], "pageBreakBefore");
    setParagraphFlag(paragraphs[headingParagraphs[index]], "keepNext");
    setParagraphFlag(paragraphs[leadParagraphs[index]], "keepNext");
    keepTableTogether(tables[index + 2]);
  });

  zip.file("word/document.xml", new XMLSerializer().serializeToString(document), { createFolders: false });
  const bytes = await zip.generateAsync({ type: "uint8array", compression: "DEFLATE", compressionOptions: { level: 6 } });
  return new Blob([bytes], { type: MIME });
}
