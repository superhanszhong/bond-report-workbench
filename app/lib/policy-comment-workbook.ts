import XLSX from "xlsx-js-style";
import type { ParsedBondRecord } from "./workbench";
import { policyDraftSpreadRecords, type PolicyCommentDraft } from "./policy-comment";
import { recordKey } from "./record-merge";
import { normalizeBondCode } from "./bond-code";

const SPREAD_HEADERS = [
  "发行日期", "代码", "期限", "发行量", "中标利率", "综收", "中标比二级(bp)", "全场倍数", "边际倍数",
  "首场边际投标量（亿）", "首场边际中标量（亿）", "追加场倍数", "备注", "修正久期", "每100亏几毛", "发行人",
  "代码1", "估值获取日期", "前一日估值", "前一日估值1", "截标前二级价格", "修正久期1", "综收-二级", "中标-二级",
  "募集用途", "备注", "综收比二级(bp)", "", "", "", "债券代码", "久期",
] as const;

function numeric(value = "") {
  const clean = value.trim();
  if (!clean) return "";
  const parsed = Number(clean);
  return Number.isFinite(parsed) ? parsed : "";
}

function dateText(value: unknown) {
  if (value instanceof Date && !Number.isNaN(value.valueOf())) return value.toISOString().slice(0, 10);
  const source = String(value || "").trim();
  const match = source.match(/^(\d{4})[-年/](\d{1,2})[-月/](\d{1,2})/);
  return match ? `${match[1]}-${match[2].padStart(2, "0")}-${match[3].padStart(2, "0")}` : source;
}

function maturityYears(value = "") {
  const source = value.trim().toUpperCase();
  const number = Number.parseFloat(source);
  if (!Number.isFinite(number)) return Number.POSITIVE_INFINITY;
  if (/D$/.test(source)) return number / 365;
  if (/M$/.test(source)) return number / 12;
  return number;
}

function excelDate(value: string) {
  const [year, month, day] = value.split("-").map(Number);
  return new Date(year, month - 1, day, 12);
}

function rawNumber(raw: Record<string, unknown>, ...keys: string[]) {
  for (const key of keys) {
    const source = raw[key];
    if (source === null || source === undefined || source === "") continue;
    const parsed = Number(String(source).replace(/[^\d.-]/g, ""));
    if (Number.isFinite(parsed)) return parsed;
  }
  return "";
}

function plansForDate(plans: ParsedBondRecord[], tradeDate: string) {
  const sessionOrder = new Map<string, number>();
  const selected = plans.filter((plan) => plan.tradeDate === tradeDate && ["国债", "国开债", "口行债", "农发债"].includes(plan.bondType || ""));
  selected.forEach((plan) => {
    const session = `${plan.issuer}|${plan.issuanceRoute}|${plan.bidTime}`;
    if (!sessionOrder.has(session)) sessionOrder.set(session, sessionOrder.size);
  });
  return selected.sort((left, right) => {
    const leftSession = `${left.issuer}|${left.issuanceRoute}|${left.bidTime}`;
    const rightSession = `${right.issuer}|${right.issuanceRoute}|${right.bidTime}`;
    return sessionOrder.get(leftSession)! - sessionOrder.get(rightSession)!
      || maturityYears(left.tenor || "") - maturityYears(right.tenor || "")
      || String(left.bondCode).localeCompare(String(right.bondCode));
  });
}

export function buildDailySpreadInputWorkbook(drafts: PolicyCommentDraft[], plans: ParsedBondRecord[], tradeDate: string) {
  const draftByKey = new Map(drafts.map((draft) => [recordKey(draft), draft]));
  const completedByKey = new Map(policyDraftSpreadRecords(drafts, plans).map((record) => [recordKey(record), record]));
  const rows = plansForDate(plans, tradeDate).map((plan) => {
    const key = recordKey(plan);
    const draft = draftByKey.get(key);
    const completed = completedByKey.get(key);
    const dr = Boolean(draft && /[XZ]\d*$/i.test(draft.bondCode) && /^DR(?:001|007)?浮息债$/i.test(draft.rateType));
    const ownSecondary = Boolean(draft && !draft.referenceBond?.trim() && draft.benchmarkType === "二级");
    const ownValuation = Boolean(draft && !draft.referenceBond?.trim() && /估值|估价/.test(draft.benchmarkType));
    const finalValue = draft ? numeric(draft.finalValue) : "";
    const benchmarkValue = draft ? numeric(draft.benchmarkValue) : "";
    const numericSpread = ownSecondary && !dr && finalValue !== "" && benchmarkValue !== ""
      ? Number(((finalValue - benchmarkValue) * 100).toFixed(2)) : "";
    const routeNote = plan.issuanceRoute === "报价发行" ? "前台报价发行" : plan.issuanceRoute === "上清所" ? "上清所" : "";
    const raw = plan.raw || {};
    const row = Array.from({ length: SPREAD_HEADERS.length }, () => "" as string | number | Date);
    row[0] = excelDate(tradeDate);
    row[1] = normalizeBondCode(plan.bondCode);
    row[2] = Number.isFinite(Number(plan.tenor)) ? Number(plan.tenor) : plan.tenor || "";
    row[3] = plan.amount ?? "";
    if (finalValue !== "") row[dr ? 5 : 4] = dr ? `${finalValue}元` : finalValue / 100;
    if (completed?.summaryMeta?.auctionSpreadText) row[6] = numericSpread === "" ? completed.summaryMeta.auctionSpreadText : numericSpread;
    row[7] = rawNumber(raw, "全场倍数");
    row[8] = rawNumber(raw, "边际倍数");
    row[9] = rawNumber(raw, "首场边际投标量（亿）", "边际投标量（亿）");
    row[10] = rawNumber(raw, "首场边际中标量（亿）", "边际中标量（亿）");
    row[11] = rawNumber(raw, "追加场倍数");
    row[12] = [draft?.rateType, routeNote].filter(Boolean).join("，");
    row[15] = plan.issuer || "";
    if (ownValuation && benchmarkValue !== "") {
      row[18] = benchmarkValue;
      row[19] = benchmarkValue / 100;
    }
    if (ownSecondary && benchmarkValue !== "") {
      row[20] = dr ? `${benchmarkValue}元` : benchmarkValue / 100;
      if (!dr && finalValue !== "") row[23] = Number(((finalValue - benchmarkValue) * 100).toFixed(2));
    }
    row[30] = `${normalizeBondCode(plan.bondCode)}.IB`;
    return row;
  });
  const sheet = XLSX.utils.aoa_to_sheet([SPREAD_HEADERS, ...rows]);
  sheet["!autofilter"] = { ref: `A1:AF${Math.max(rows.length + 1, 1)}` };
  sheet["!freeze"] = { xSplit: 2, ySplit: 1, topLeftCell: "C2", activePane: "bottomRight", state: "frozen" };
  sheet["!cols"] = [12, 17, 10, 11, 13, 13, 20, 11, 11, 18, 18, 13, 22, 12, 14, 22, 16, 14, 14, 14, 18, 14, 14, 14, 16, 14, 18, 5, 5, 5, 18, 12]
    .map((width) => ({ wch: width }));
  for (let column = 0; column < SPREAD_HEADERS.length; column += 1) {
    const cell = sheet[XLSX.utils.encode_cell({ r: 0, c: column })];
    if (cell) cell.s = {
      font: { name: "微软雅黑", sz: 10, bold: true, color: { rgb: "FFFFFF" } },
      fill: { fgColor: { rgb: "8F5148" } },
      alignment: { horizontal: "center", vertical: "center", wrapText: true },
    };
  }
  rows.forEach((_, index) => {
    const row = index + 1;
    const dateCell = sheet[XLSX.utils.encode_cell({ r: row, c: 0 })];
    if (dateCell) dateCell.z = "yyyy-mm-dd";
    [4, 20].forEach((column) => {
      const cell = sheet[XLSX.utils.encode_cell({ r: row, c: column })];
      if (cell && typeof cell.v === "number") cell.z = "0.0000%";
    });
  });
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, sheet, "今日待粘贴");
  return XLSX.write(workbook, { type: "array", bookType: "xlsx", cellStyles: true }) as ArrayBuffer;
}

function cloneStyle(value: unknown) {
  return value && typeof value === "object" ? JSON.parse(JSON.stringify(value)) : value;
}

export function buildUpdatedSpreadWorkbook(templateBytes: ArrayBuffer, drafts: PolicyCommentDraft[], plans: ParsedBondRecord[], tradeDate: string) {
  const workbook = XLSX.read(templateBytes, { type: "array", cellStyles: true, cellFormula: true, cellDates: true });
  const sheet = workbook.Sheets.Sheet1 || workbook.Sheets[workbook.SheetNames[0]];
  if (!sheet) throw new Error("一二级底稿缺少数据工作表");
  const requiredHeaders = ["发行日期", "代码", "期限", "发行量", "中标利率", "综收", "中标比二级(bp)", "发行人", "截标前二级价格"];
  const header = new Map<string, number>();
  for (let column = 0; column < 40; column += 1) {
    const value = String(sheet[XLSX.utils.encode_cell({ r: 0, c: column })]?.v || "").replace(/\s+/g, "");
    if (value && !header.has(value)) header.set(value, column);
  }
  if (requiredHeaders.some((name) => !header.has(name.replace(/\s+/g, "")))) throw new Error("最新一二级文件与当前模板列不一致，请上传与0923底稿同结构的文件");
  const range = XLSX.utils.decode_range(sheet["!ref"] || "A1:AF1");
  // Some source workbooks carry formatting residue all the way to XFA. Keeping
  // that artificial range makes XLSX serialization block the browser for tens
  // of seconds. Retain every cell with an actual value/formula, but drop blank
  // far-column residue from the written range.
  let meaningfulEndColumn = 31;
  let meaningfulEndRow = 0;
  Object.entries(sheet).forEach(([address, cell]) => {
    if (address.startsWith("!") || !cell || typeof cell !== "object") return;
    const value = (cell as XLSX.CellObject).v;
    const formulaValue = (cell as XLSX.CellObject).f;
    if ((value === "" || value === null || value === undefined) && !formulaValue) return;
    const decoded = XLSX.utils.decode_cell(address);
    meaningfulEndColumn = Math.max(meaningfulEndColumn, decoded.c);
    meaningfulEndRow = Math.max(meaningfulEndRow, decoded.r);
  });
  let lastDataRow = 0;
  const existingRows = new Map<string, number>();
  for (let row = 1; row <= range.e.r; row += 1) {
    const code = sheet[XLSX.utils.encode_cell({ r: row, c: 1 })]?.v;
    if (!code) continue;
    lastDataRow = row;
    const date = dateText(sheet[XLSX.utils.encode_cell({ r: row, c: 0 })]?.v);
    existingRows.set(recordKey({ tradeDate: date, bondCode: normalizeBondCode(code) }), row);
  }
  if (!lastDataRow) throw new Error("一二级底稿没有可复制的数据行");
  const draftByKey = new Map(drafts.map((draft) => [recordKey(draft), draft]));
  const completedByKey = new Map(policyDraftSpreadRecords(drafts, plans).map((record) => [recordKey(record), record]));
  const selected = plansForDate(plans, tradeDate);
  const assign = (row: number, column: number, value: unknown, format?: string) => {
    const address = XLSX.utils.encode_cell({ r: row, c: column });
    const prior = sheet[address];
    const cell: XLSX.CellObject = prior ? { ...prior, s: cloneStyle(prior.s) } : { t: "z", v: "" };
    delete cell.f;
    if (value === "" || value === null || value === undefined) { cell.t = "z"; cell.v = ""; }
    else if (typeof value === "number") { cell.t = "n"; cell.v = value; }
    else { cell.t = "s"; cell.v = String(value); }
    if (format) cell.z = format;
    sheet[address] = cell;
  };
  const formula = (row: number, column: number, expression: string) => {
    const address = XLSX.utils.encode_cell({ r: row, c: column });
    const prior = sheet[address];
    sheet[address] = { ...(prior || {}), t: "n", f: expression, v: 0, s: cloneStyle(prior?.s) } as XLSX.CellObject;
  };
  let appendRow = lastDataRow + 1;
  for (const plan of selected) {
    const key = recordKey(plan);
    const existing = existingRows.get(key);
    const row = existing ?? appendRow++;
    if (existing === undefined) {
      for (let column = 0; column <= 31; column += 1) {
        const source = sheet[XLSX.utils.encode_cell({ r: lastDataRow, c: column })];
        if (source) sheet[XLSX.utils.encode_cell({ r: row, c: column })] = { t: "z", v: "", s: cloneStyle(source.s), z: source.z } as XLSX.CellObject;
      }
    }
    const draft = draftByKey.get(key);
    const completed = completedByKey.get(key);
    const dr = Boolean(draft && /[XZ]\d*$/i.test(draft.bondCode) && /^DR(?:001|007)?浮息债$/i.test(draft.rateType));
    const ownSecondary = Boolean(draft && !draft.referenceBond?.trim() && draft.benchmarkType === "二级");
    const ownValuation = Boolean(draft && !draft.referenceBond?.trim() && /估值|估价/.test(draft.benchmarkType));
    const finalValue = draft ? numeric(draft.finalValue) : "";
    const benchmarkValue = draft ? numeric(draft.benchmarkValue) : "";
    assign(row, 0, tradeDate, "m/d/yy");
    assign(row, 1, normalizeBondCode(plan.bondCode));
    assign(row, 2, Number.isFinite(Number(plan.tenor)) ? Number(plan.tenor) : plan.tenor || "");
    assign(row, 3, plan.amount ?? "");
    if (finalValue !== "") assign(row, 4, dr ? "" : finalValue / 100, "0.0000%");
    if (dr && finalValue !== "") assign(row, 5, `${finalValue}元`);
    if (completed?.summaryMeta?.auctionSpreadText) {
      const numericSpread = ownSecondary && !dr && finalValue !== "" && benchmarkValue !== "" ? Number(((finalValue - benchmarkValue) * 100).toFixed(2)) : null;
      assign(row, 6, numericSpread ?? completed.summaryMeta.auctionSpreadText);
    }
    const routeNote = plan.issuanceRoute === "报价发行" ? "前台报价发行" : plan.issuanceRoute === "上清所" ? "上清所" : "";
    const note = [draft?.rateType, routeNote].filter(Boolean).join("，");
    if (note) assign(row, 12, note);
    assign(row, 15, plan.issuer || "");
    if (ownValuation && benchmarkValue !== "") assign(row, 18, benchmarkValue);
    if (ownSecondary && benchmarkValue !== "") assign(row, 20, dr ? `${benchmarkValue}元` : benchmarkValue / 100, dr ? "General" : "0.0000%");
    assign(row, 30, `${normalizeBondCode(plan.bondCode)}.IB`);
    const excelRow = row + 1;
    if (!dr) {
      formula(row, 14, `IF(OR(N${excelRow}="",G${excelRow}=""),"",N${excelRow}*IFERROR(LEFT(G${excelRow},FIND("(",G${excelRow})-1),G${excelRow})/10)`);
      formula(row, 19, `IF(S${excelRow}="","",S${excelRow}/100)`);
      formula(row, 22, `IF(OR(F${excelRow}="",U${excelRow}=""),"",(F${excelRow}-U${excelRow})*10000)`);
      formula(row, 23, `IF(OR(E${excelRow}="",U${excelRow}=""),"",(E${excelRow}-U${excelRow})*10000)`);
      formula(row, 26, `IF(W${excelRow}="","",W${excelRow})`);
    }
  }
  range.e.r = Math.max(meaningfulEndRow, appendRow - 1);
  range.e.c = meaningfulEndColumn;
  sheet["!ref"] = XLSX.utils.encode_range(range);
  if (sheet["!autofilter"]) sheet["!autofilter"] = { ref: `A1:AF${Math.max(lastDataRow + 1, appendRow)}` };
  return XLSX.write(workbook, { type: "array", bookType: "xlsx", cellStyles: true, cellFormula: true }) as ArrayBuffer;
}
