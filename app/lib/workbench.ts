import XLSX from "xlsx-js-style";

export type ParsedBondRecord = {
  tradeDate: string;
  bondCode?: string;
  shortName?: string;
  fullName?: string;
  issuer?: string;
  region?: string;
  bondType?: string;
  issuanceRoute?: string;
  venue?: string;
  bidTime?: string;
  tenor?: string;
  amount?: number | null;
  spread?: number | null;
  floorRate?: number | null;
  fee?: number | null;
  distributionDate?: string;
  remark?: string;
  summaryMeta?: SpreadSummaryMeta;
  raw?: Record<string, unknown>;
};

export type SpreadSummaryMeta = {
  baseCode: string;
  route: string;
  rateType: string;
  displaySpreadText: string;
  auctionSpreadText: string;
  allInText: string;
  secondaryText: string;
  winningRateText: string;
  note: string;
  proceeds: string;
  previous?: {
    date: string;
    code: string;
    comparisonType: "same_bond" | "discount_comparator";
    displaySpreadText: string;
    auctionSpreadText: string;
    allInText: string;
    secondaryText: string;
    note: string;
    spread: number | null;
  };
};

export type SpreadAudit = {
  status: "provided" | "recalculated" | "derived_treasury" | "missing" | "excluded";
  source: string;
  provided: number | null;
  derived: number | null;
  reason: string;
};

const LOCAL_ALIASES: Record<string, string[]> = {
  招标日: ["招标日", "发行日", "发行日期"],
  招标时间: ["招标时间"],
  招标场所: ["招标场所", "招标系统"],
  债券代码: ["债券代码", "代码"],
  债券简称: ["债券简称", "简称"],
  类型: ["类型"],
  性质: ["性质"],
  期限: ["期限"],
  发行量: ["发行量(亿元)", "发行量(亿)", "发行量"],
  手续费: ["手续费(%)", "手续费"],
  财政部基准: ["财政部5日均值(%)", "财政部同期国债(5日)", "财政部基准"],
  加点: ["加点(bp)", "投标加点(bp)", "加点"],
  投标下限: ["投标下限", "投标下限(%)", "下限"],
  缴款日: ["缴款日", "分销日"],
  债券全称: ["债券全称", "全称"],
  区域名称: ["区域名称", "地区"],
};

const ISSUER_MAP: Record<string, string> = {
  中华人民共和国财政部: "国债",
  国家开发银行: "国开债",
  中国进出口银行: "口行债",
  中国农业发展银行: "农发债",
};

const MATURITY_ALIASES: Record<string, string[]> = {
  发行状态: ["发行状态", "状态"],
  债券简称: ["债券简称", "简称"],
  债券代码: ["债券代码", "代码"],
  发行规模: ["发行规模", "发行规模(亿元)", "发行量", "发行量(亿元)"],
  期限: ["期限"],
  实际到期日: ["实际到期日", "到期日", "兑付日"],
  发行人: ["发行人"],
  所属区域: ["所属区域", "区域名称", "地区"],
};

const ISSUANCE_PLAN_ALIASES: Record<string, string[]> = {
  债券简称: ["债券简称", "简称"],
  债券代码: ["债券代码", "代码"],
  发行期限: ["发行期限", "期限"],
  计划发行量: ["计划发行量(亿)", "计划发行量（亿）", "计划发行量"],
  实际发行量: ["实际发行量(亿)", "实际发行量（亿）", "实际发行量"],
  招标时间: ["招标时间", "发行时间"],
  招标标的: ["招标标的"],
  招标方式: ["招标方式"],
  发行利率: ["发行利率"],
  加权利率: ["加权利率(%)", "加权利率（%）", "加权利率"],
  发行起始日: ["发行起始日", "招标日", "发行日期"],
  托管机构: ["托管机构", "招标场所"],
};

const SPECIAL_SPREAD_PATTERN = /绿债|绿色|主题债|浮息债/;

export function isSpecialSpreadBond(row: Pick<ParsedBondRecord, "remark" | "summaryMeta">) {
  return SPECIAL_SPREAD_PATTERN.test(`${row.remark || ""}${row.summaryMeta?.rateType || ""}${row.summaryMeta?.note || ""}`);
}

function clean(value: unknown) {
  if (value === null || value === undefined || value === "" || value === "--") return "";
  return String(value).trim();
}

function numberValue(value: unknown): number | null {
  const normalized = clean(value).replace(/,/g, "").replace(/%/g, "");
  if (!normalized) return null;
  const n = Number(normalized);
  return Number.isFinite(n) ? n : null;
}

function basisPointValue(value: unknown): number | null {
  const direct = numberValue(value);
  if (direct !== null) return direct;
  const text = clean(value);
  if (!text || /净价/.test(text)) return null;
  if (/平/.test(text) && !/^-?\d/.test(text)) return 0;
  const match = text.match(/^(-?\d+(?:\.\d+)?)/);
  return match ? Number(match[1]) : null;
}

function percentYieldValue(value: unknown): number | null {
  const text = clean(value);
  if (!text || /净价|价格|元/.test(text)) return null;
  const parsed = numberValue(value);
  if (parsed === null) return null;
  return Math.abs(parsed) < 0.2 ? parsed * 100 : parsed;
}

export function resolveSpreadBp(input: {
  provided: unknown;
  allIn: unknown;
  secondary: unknown;
  winningRate: unknown;
  issuer: string;
  remark: string;
}): { spread: number | null; audit: SpreadAudit } {
  const provided = basisPointValue(input.provided);
  const allIn = percentYieldValue(input.allIn);
  const secondary = percentYieldValue(input.secondary);
  const winningRate = percentYieldValue(input.winningRate);
  const derivedFromAllIn = allIn !== null && secondary !== null
    ? Number(((allIn - secondary) * 100).toFixed(6)) : null;
  const derivedTreasury = input.issuer === "中华人民共和国财政部" && winningRate !== null && secondary !== null
    ? Number(((winningRate - secondary) * 100).toFixed(6)) : null;
  const derived = derivedFromAllIn ?? derivedTreasury;

  if (SPECIAL_SPREAD_PATTERN.test(input.remark)) {
    return { spread: provided, audit: { status: "excluded", source: "备注筛选", provided, derived, reason: "绿债、主题债或浮息债不纳入普通利差图" } };
  }
  if (provided !== null && derivedFromAllIn !== null && Math.abs(provided - derivedFromAllIn) > 0.05) {
    return { spread: derivedFromAllIn, audit: { status: "recalculated", source: "综收与二级复算", provided, derived: derivedFromAllIn, reason: "表内利差与收益率复算差异超过0.05BP，采用复算值" } };
  }
  if (provided !== null) {
    return { spread: provided, audit: { status: "provided", source: "综收-二级列", provided, derived, reason: "使用表内利差并完成收益率交叉核对" } };
  }
  if (derivedTreasury !== null) {
    return { spread: derivedTreasury, audit: { status: "derived_treasury", source: "国债中标利率与二级复算", provided, derived: derivedTreasury, reason: "国债综收利差为空，按无发行手续费口径使用中标利率减二级" } };
  }
  return { spread: null, audit: { status: "missing", source: "无可用利差", provided, derived, reason: "缺少可比收益率，未纳入散点图" } };
}

function excelDate(value: unknown): Date | null {
  if (value instanceof Date && !Number.isNaN(value.valueOf())) return value;
  if (typeof value === "number") {
    const compact = String(Math.trunc(value));
    if (/^(?:19|20)\d{6}$/.test(compact)) {
      return new Date(Number(compact.slice(0, 4)), Number(compact.slice(4, 6)) - 1, Number(compact.slice(6, 8)));
    }
    const parsed = XLSX.SSF.parse_date_code(value);
    if (parsed) return new Date(parsed.y, parsed.m - 1, parsed.d);
  }
  const source = clean(value);
  if (/^(?:19|20)\d{6}$/.test(source)) {
    return new Date(Number(source.slice(0, 4)), Number(source.slice(4, 6)) - 1, Number(source.slice(6, 8)));
  }
  const text = source.slice(0, 10).replace(/[/.]/g, "-");
  const parts = text.split("-").map(Number);
  if (parts.length === 3 && parts.every(Number.isFinite)) return new Date(parts[0], parts[1] - 1, parts[2]);
  return null;
}

export function isoDate(value: Date | null) {
  if (!value) return "";
  const y = value.getFullYear();
  const m = String(value.getMonth() + 1).padStart(2, "0");
  const d = String(value.getDate()).padStart(2, "0");
  return `${y}-${m}-${d}`;
}

export function mondayOf(value: string | Date) {
  const date = typeof value === "string" ? new Date(`${value}T12:00:00`) : new Date(value);
  const day = date.getDay() || 7;
  date.setDate(date.getDate() - day + 1);
  return isoDate(date);
}

export function fridayOf(weekStart: string) {
  const date = new Date(`${weekStart}T12:00:00`);
  date.setDate(date.getDate() + 4);
  return isoDate(date);
}

export function maturityWeekStart(value: string | Date) {
  const date = typeof value === "string" ? new Date(`${value}T12:00:00`) : new Date(value);
  const day = date.getDay();
  if (day === 6) date.setDate(date.getDate() - 5);
  if (day === 0) date.setDate(date.getDate() - 6);
  return mondayOf(date);
}

function actualSheetRange(sheet: XLSX.WorkSheet) {
  let minRow = Number.POSITIVE_INFINITY;
  let minCol = Number.POSITIVE_INFINITY;
  let maxRow = -1;
  let maxCol = -1;

  Object.keys(sheet).forEach((address) => {
    if (address.startsWith("!")) return;
    const cell = sheet[address] as XLSX.CellObject | undefined;
    if (!cell || ((cell.v === undefined || cell.v === null || cell.v === "") && !cell.f)) return;
    const { r, c } = XLSX.utils.decode_cell(address);
    minRow = Math.min(minRow, r);
    minCol = Math.min(minCol, c);
    maxRow = Math.max(maxRow, r);
    maxCol = Math.max(maxCol, c);
  });

  if (maxRow < 0 || maxCol < 0) return undefined;
  return { s: { r: minRow, c: minCol }, e: { r: maxRow, c: maxCol } };
}

function boundedSheetRows(sheet: XLSX.WorkSheet) {
  const range = actualSheetRange(sheet);
  if (!range) return [];
  return XLSX.utils.sheet_to_json<unknown[]>(sheet, { header: 1, raw: true, defval: null, range });
}

function sheetRows(workbook: XLSX.WorkBook) {
  return boundedSheetRows(workbook.Sheets[workbook.SheetNames[0]]);
}

function headerIndex(rows: unknown[][], aliases: Record<string, string[]>) {
  let bestRow = -1;
  let bestIndex: Record<string, number> = {};
  let bestScore = -1;
  rows.slice(0, 16).forEach((row, rowIndex) => {
    const cells = new Map<string, number>();
    (row || []).forEach((value, col) => {
      const label = clean(value).replace(/\s+/g, "");
      if (label && !cells.has(label)) cells.set(label, col);
    });
    const index: Record<string, number> = {};
    Object.entries(aliases).forEach(([key, options]) => {
      const match = options.find((option) => cells.has(option.replace(/\s+/g, "")));
      if (match) index[key] = cells.get(match.replace(/\s+/g, ""))!;
    });
    const score = Object.keys(index).length;
    if (score > bestScore) {
      bestScore = score;
      bestRow = rowIndex;
      bestIndex = index;
    }
  });
  return { row: bestRow, index: bestIndex };
}

function rowObject(row: unknown[], index: Record<string, number>) {
  const output: Record<string, unknown> = {};
  Object.entries(index).forEach(([name, col]) => { output[name] = row[col]; });
  return output;
}

function shortVenue(value: unknown) {
  const text = clean(value);
  if (text.includes("深圳")) return "深交所";
  if (text.includes("上海证券")) return "上交所";
  if (text.includes("中央国债")) return "中债登";
  return text;
}

function regionFromName(fullName: string) {
  const match = fullName.match(/^\d{4}年([\u4e00-\u9fa5]{2,12}?)(?:壮族自治区|回族自治区|维吾尔自治区|自治区|省|市)/);
  return match?.[1] || "";
}

function cleanBondType(fullName: string, fallback: string, region: string) {
  let text = fullName.replace(/^\d{4}年/, "");
  const area = region || regionFromName(fullName);
  if (area) text = text.replace(new RegExp(`^${area}(?:壮族自治区|回族自治区|维吾尔自治区|自治区|省|市)?`), "");
  text = text.replace(/壮族自治区|回族自治区|维吾尔自治区|自治区/g, "")
    .replace(/地方政府|政府/g, "").replace(/债券/g, "债").trim();
  return text || fallback || "地方债";
}

function routeFromRemark(remark: string, bondCode = "") {
  if (/报价发行|前台报价/.test(remark)) return "报价发行";
  if (/清发|上清所/.test(remark)) return "上清所";
  if (/^09/.test(bondCode)) return "上清所";
  return "中债招标";
}

function baseBondCode(code: string) {
  return code.replace(/[XZ]\d*$/i, "");
}

function isDiscountBond(record: Pick<ParsedBondRecord, "tenor">) {
  return /D$/i.test(record.tenor || "");
}

function isReopenedBond(record: Pick<ParsedBondRecord, "bondCode">) {
  return /[XZ]\d*$/i.test(record.bondCode || "");
}

function discountComparatorKey(record: Pick<ParsedBondRecord, "bondType" | "tenor" | "summaryMeta" | "remark">) {
  return [record.bondType || "", (record.tenor || "").toUpperCase(), record.summaryMeta?.rateType || rateTypeFromRemark(record.remark || "")].join("|");
}

function rateTypeFromRemark(remark: string) {
  if (/DR0{1,2}1/i.test(remark)) return "DR001浮息债";
  if (/DR007/i.test(remark)) return "DR007浮息债";
  if (/DR/i.test(remark)) return "DR007浮息债";
  if (/LPR/i.test(remark)) return "LPR浮息债";
  return "固息或贴现";
}

export async function parseLocalBondFile(file: File) {
  const workbook = XLSX.read(await file.arrayBuffer(), { type: "array", cellDates: true });
  const rows = sheetRows(workbook);
  const { row: headerRow, index } = headerIndex(rows, LOCAL_ALIASES);
  const required = ["招标日", "债券代码", "债券简称", "期限", "发行量"];
  const missing = required.filter((name) => index[name] === undefined);
  if (headerRow < 0 || missing.length) throw new Error(`地方债表缺少：${missing.join("、")}`);

  const records: ParsedBondRecord[] = [];
  for (let r = headerRow + 1; r < rows.length; r += 1) {
    const raw = rowObject(rows[r], index);
    if (!clean(raw.债券代码)) continue;
    const date = excelDate(raw.招标日);
    if (!date) continue;
    const spread = numberValue(raw.加点) || 0;
    const benchmark = numberValue(raw.财政部基准);
    const explicitFloor = numberValue(raw.投标下限);
    const floorRate = benchmark === null ? explicitFloor : benchmark + spread / 100;
    const fullName = clean(raw.债券全称);
    const region = clean(raw.区域名称) || regionFromName(fullName);
    records.push({
      tradeDate: isoDate(date),
      bondCode: clean(raw.债券代码).replace(/\.IB$/i, ""),
      shortName: clean(raw.债券简称),
      fullName,
      region,
      bondType: cleanBondType(fullName, `${clean(raw.性质)}${clean(raw.类型)}`, region),
      venue: shortVenue(raw.招标场所),
      bidTime: clean(raw.招标时间),
      tenor: clean(raw.期限).replace(/Y$/i, ""),
      amount: numberValue(raw.发行量),
      floorRate,
      fee: numberValue(raw.手续费),
      distributionDate: isoDate(excelDate(raw.缴款日)),
      raw,
    });
  }
  if (!records.length) throw new Error("未读取到地方债发行明细");
  records.sort((a, b) => `${a.tradeDate}${a.bidTime}${a.venue}`.localeCompare(`${b.tradeDate}${b.bidTime}${b.venue}`));
  return records;
}

export async function parseMaturityFile(file: File) {
  const workbook = XLSX.read(await file.arrayBuffer(), { type: "array", cellDates: true });
  const records: ParsedBondRecord[] = [];
  workbook.SheetNames.forEach((sheetName) => {
    const rows = boundedSheetRows(workbook.Sheets[sheetName]);
    const { row: headerRow, index } = headerIndex(rows, MATURITY_ALIASES);
    const required = ["债券简称", "债券代码", "发行规模", "实际到期日", "发行人"];
    if (headerRow < 0 || required.some((name) => index[name] === undefined)) return;
    for (let r = headerRow + 1; r < rows.length; r += 1) {
      const raw = rowObject(rows[r], index);
      const date = excelDate(raw.实际到期日);
      const bondCode = clean(raw.债券代码).replace(/\.(?:IB|SH|SZ)$/i, "");
      const amount = numberValue(raw.发行规模);
      const issuer = clean(raw.发行人);
      if (!date || !bondCode || amount === null) continue;
      const rateType = ISSUER_MAP[issuer];
      records.push({
        tradeDate: isoDate(date), bondCode, shortName: clean(raw.债券简称), issuer,
        region: clean(raw.所属区域), bondType: rateType || "地方债",
        tenor: clean(raw.期限), amount, remark: clean(raw.发行状态),
        raw: { ...raw, __maturityKind: rateType ? "rate" : "local" },
      });
    }
  });
  if (!records.length) throw new Error("未识别到到期明细，请使用包含债券简称、债券代码、发行规模、实际到期日和发行人的 Excel");
  records.sort((a, b) => `${a.tradeDate}|${a.bondCode}`.localeCompare(`${b.tradeDate}|${b.bondCode}`));
  return records;
}

export function maturityKind(row: ParsedBondRecord) {
  const kind = row.raw?.__maturityKind;
  if (kind === "rate" || kind === "local") return kind;
  return ["国债", "国开债", "口行债", "农发债"].includes(row.bondType || "") ? "rate" : "local";
}

export function maturityDailyTotals(rows: ParsedBondRecord[], weekStart: string) {
  return Object.fromEntries(Array.from({ length: 5 }, (_, index) => {
    const date = new Date(`${weekStart}T12:00:00`);
    date.setDate(date.getDate() + index);
    const iso = isoDate(date);
    return [iso, rows.filter((row) => row.tradeDate === iso).reduce((sum, row) => sum + (row.amount || 0), 0)];
  }));
}

function rateIssuerFromShortName(shortName: string) {
  if (/国债/.test(shortName)) return { issuer: "中华人民共和国财政部", bondType: "国债" };
  if (/国开/.test(shortName)) return { issuer: "国家开发银行", bondType: "国开债" };
  if (/进出/.test(shortName)) return { issuer: "中国进出口银行", bondType: "口行债" };
  if (/农发/.test(shortName)) return { issuer: "中国农业发展银行", bondType: "农发债" };
  return null;
}

export async function parseIssuancePlanFile(file: File) {
  const workbook = XLSX.read(await file.arrayBuffer(), { type: "array", cellDates: true });
  const records: ParsedBondRecord[] = [];
  workbook.SheetNames.forEach((sheetName) => {
    const rows = boundedSheetRows(workbook.Sheets[sheetName]);
    const { row: headerRow, index } = headerIndex(rows, ISSUANCE_PLAN_ALIASES);
    const required = ["债券简称", "债券代码", "发行期限", "招标时间", "发行起始日"];
    if (headerRow < 0 || required.some((name) => index[name] === undefined)) return;
    for (let r = headerRow + 1; r < rows.length; r += 1) {
      const raw = rowObject(rows[r], index);
      const shortName = clean(raw.债券简称);
      const issuer = rateIssuerFromShortName(shortName);
      const date = excelDate(raw.发行起始日);
      const bondCode = clean(raw.债券代码).replace(/\.IB$/i, "");
      if (!issuer || !date || !bondCode) continue;
      const auctionTarget = clean(raw.招标标的);
      const custody = clean(raw.托管机构);
      const routeSource = `${shortName}${auctionTarget}${custody}`;
      const route = /报价发行|前台报价/.test(auctionTarget)
        ? "报价发行"
        : /^09/.test(bondCode) || /清发|上清所/.test(routeSource)
          ? "上清所"
          : "中债招标";
      records.push({
        tradeDate: isoDate(date), bondCode, shortName, issuer: issuer.issuer, bondType: issuer.bondType,
        issuanceRoute: route, venue: shortVenue(raw.托管机构), bidTime: clean(raw.招标时间),
        tenor: clean(raw.发行期限).replace(/Y$/i, ""),
        amount: numberValue(raw.实际发行量) ?? numberValue(raw.计划发行量),
        remark: route === "报价发行" ? "前台报价发行" : route === "上清所" ? "上清所" : "", raw,
      });
    }
  });
  if (!records.length) throw new Error("未识别到国债及政金债发行计划，请确认文件包含发行起始日、招标时间和债券代码");
  records.sort((a, b) => `${a.tradeDate}|${a.bidTime}|${a.bondCode}`.localeCompare(`${b.tradeDate}|${b.bidTime}|${b.bondCode}`));
  return records;
}

export function maturityCategory(row: ParsedBondRecord) {
  const name = `${row.shortName || ""}${row.fullName || ""}`;
  if (row.bondType === "国债") {
    if (/贴现/.test(name) || /^269/.test(row.bondCode || "")) return "贴现国债";
    if (/超长|特别国债/.test(name)) return "超长特国";
    return "附息国债";
  }
  if (row.bondType === "国开债") return /贴现|清发/.test(name) ? "贴现国开" : "国开";
  if (row.bondType === "口行债") return /贴现|清发/.test(name) ? "进出清发" : "进出";
  if (row.bondType === "农发债") return /贴现|清发/.test(name) ? "农发清发" : "农发";
  return row.bondType || "利率债";
}

export function rateMaturityBreakdown(rows: ParsedBondRecord[]) {
  const order = ["贴现国债", "附息国债", "超长特国", "贴现国开", "国开", "进出清发", "进出", "农发清发", "农发"];
  const totals = new Map<string, number>();
  rows.forEach((row) => {
    const label = maturityCategory(row);
    totals.set(label, (totals.get(label) || 0) + (row.amount || 0));
  });
  return order.filter((label) => totals.has(label))
    .map((label) => `${label}:${Number((totals.get(label) || 0).toFixed(4))}亿`).join("　") || "-";
}

export async function parseSpreadFile(file: File) {
  const workbook = XLSX.read(await file.arrayBuffer(), { type: "array", cellDates: true });
  const aliases = {
    发行日期: ["发行日期", "招标日"], 代码: ["代码", "债券代码"], 简称: ["简称", "债券简称"],
    期限: ["期限"], 发行人: ["发行人"], 备注: ["备注"],
    利差: ["综收比二级(bp)", "综收比二级（bp）", "综收-二级", "综收－二级"],
    发行量: ["发行量", "发行量(亿元)"], 中标利率: ["中标利率"], 综收: ["综收"],
    招标利差: ["中标比二级(bp)", "中标比二级（bp）"], 二级: ["截标前二级价格"],
    前一日估值: ["前一日估值"],
    募集用途: ["募集用途"], 全场倍数: ["全场倍数"], 边际倍数: ["边际倍数"],
    边际投标量: ["首场边际投标量（亿）", "边际投标量（亿）"],
    边际中标量: ["首场边际中标量（亿）", "边际中标量（亿）"],
  };
  const candidates: Array<ParsedBondRecord & { order: number }> = [];
  let order = 0;
  workbook.SheetNames.forEach((sheetName) => {
    const sheet = workbook.Sheets[sheetName];
    const range = actualSheetRange(sheet);
    if (!range) return;
    const rows = XLSX.utils.sheet_to_json<unknown[]>(sheet, { header: 1, raw: true, defval: null, range });
    const displayRows = XLSX.utils.sheet_to_json<unknown[]>(sheet, { header: 1, raw: false, defval: "", range });
    const { row: headerRow, index } = headerIndex(rows, aliases);
    if (headerRow < 0 || index.发行日期 === undefined || index.代码 === undefined || index.利差 === undefined) return;
    for (let r = headerRow + 1; r < rows.length; r += 1) {
      const raw = rowObject(rows[r], index);
      const display = rowObject(displayRows[r] || [], index);
      const date = excelDate(raw.发行日期);
      const bondCode = clean(raw.代码).replace(/\.IB$/i, "");
      if (!date || !bondCode) continue;
      const issuer = clean(raw.发行人);
      const remark = clean(raw.备注);
      const route = routeFromRemark(remark, bondCode);
      const resolvedSpread = resolveSpreadBp({
        provided: raw.利差, allIn: raw.综收, secondary: raw.二级,
        winningRate: raw.中标利率, issuer, remark,
      });
      const meta: SpreadSummaryMeta = {
        baseCode: baseBondCode(bondCode), route, rateType: rateTypeFromRemark(remark),
        displaySpreadText: clean(display.利差), auctionSpreadText: clean(display.招标利差),
        allInText: clean(display.综收), secondaryText: clean(display.二级),
        winningRateText: clean(display.中标利率), note: remark, proceeds: clean(display.募集用途),
      };
      if (!meta.displaySpreadText && resolvedSpread.spread !== null) {
        meta.displaySpreadText = `${resolvedSpread.spread > 0 ? "+" : ""}${resolvedSpread.spread.toFixed(2)}（复算）`;
      }
      const fullRaw = { ...raw, __display: display, __summaryMeta: meta, __spreadAudit: resolvedSpread.audit };
      candidates.push({
        tradeDate: isoDate(date), bondCode, shortName: clean(raw.简称), issuer,
        bondType: ISSUER_MAP[issuer] || issuer, issuanceRoute: route,
        tenor: clean(raw.期限), amount: numberValue(raw.发行量), spread: resolvedSpread.spread,
        remark, summaryMeta: meta, raw: fullRaw, order: order++,
      });
    }
  });
  candidates.sort((a, b) => `${a.tradeDate}|${String(a.order).padStart(8, "0")}`.localeCompare(`${b.tradeDate}|${String(b.order).padStart(8, "0")}`));
  const sameBondHistory = new Map<string, ParsedBondRecord>();
  const discountHistory = new Map<string, ParsedBondRecord>();
  const records: ParsedBondRecord[] = candidates.map(({ order: _order, ...record }) => {
    void _order;
    const meta = record.summaryMeta!;
    const comparisonType = isDiscountBond(record) ? "discount_comparator" : isReopenedBond(record) ? "same_bond" : null;
    const previous = comparisonType === "discount_comparator"
      ? discountHistory.get(discountComparatorKey(record))
      : comparisonType === "same_bond" ? sameBondHistory.get(meta.baseCode) : undefined;
    if (previous?.summaryMeta) {
      meta.previous = {
        date: previous.tradeDate, code: previous.bondCode || "", comparisonType, displaySpreadText: previous.summaryMeta.displaySpreadText,
        auctionSpreadText: previous.summaryMeta.auctionSpreadText, allInText: previous.summaryMeta.allInText,
        secondaryText: previous.summaryMeta.secondaryText, note: previous.remark || "", spread: previous.spread ?? null,
      };
    }
    record.raw = { ...(record.raw || {}), __summaryMeta: meta };
    sameBondHistory.set(meta.baseCode, record);
    if (isDiscountBond(record)) discountHistory.set(discountComparatorKey(record), record);
    return record;
  });
  if (!records.length) throw new Error("未识别到一二级利差明细，请检查表头");
  return records;
}

function cnDate(value: string) {
  const date = new Date(`${value}T12:00:00`);
  return `${date.getMonth() + 1}月${date.getDate()}日`;
}

function compact(value: number | null | undefined, decimals = 4) {
  if (value === null || value === undefined || !Number.isFinite(value)) return "";
  return String(Math.round(value * 10 ** decimals) / 10 ** decimals);
}

export function localPlanText(records: ParsedBondRecord[], date: string) {
  const daily = records.filter((row) => row.tradeDate === date);
  const lines = [`【${cnDate(date)}地方债发行计划】`];
  let group = "";
  daily.forEach((row) => {
    const next = `【${[row.venue, row.bidTime, row.region ? `${row.region}债` : "地方债"].filter(Boolean).join(" ")}】`;
    if (next !== group) { lines.push(next); group = next; }
    lines.push(`【${row.bondType || "地方债"}】${row.shortName || row.bondCode}，${row.tenor || ""}${/Y|D|M/i.test(row.tenor || "") ? "" : "Y"}，${compact(row.amount)}亿，下限${compact(row.floorRate, 2)}，${row.distributionDate ? `${cnDate(row.distributionDate)}分销` : ""}，手续费${compact(row.fee)}`);
  });
  return lines.join("\n");
}

export type RollingSpreadAnalysis = {
  text: string;
  benchmarkStart: string;
  benchmarkEnd: string;
  comparableGroups: number;
  normalGroups: number;
  notableGroups: number;
  specialBonds: number;
};

function rollingTenor(value = "") {
  const text = value.trim().toUpperCase();
  const numeric = Number.parseFloat(text);
  if (!Number.isFinite(numeric)) return text || "未知期限";
  if (text.includes("D")) {
    if (numeric <= 45) return "1M";
    if (numeric <= 120) return "3M";
    if (numeric <= 240) return "6M";
    if (numeric <= 300) return "9M";
    const years = numeric / 365;
    const rounded = Math.round(years);
    return Math.abs(years - rounded) <= 0.08 ? `${rounded}Y` : `${Number(years.toFixed(2))}Y`;
  }
  if (text.includes("M")) return `${Number(numeric.toFixed(2))}M`;
  return `${Number(numeric.toFixed(2))}Y`;
}

function shiftIsoDate(value: string, days: number) {
  const date = new Date(`${value}T12:00:00`);
  date.setDate(date.getDate() + days);
  return isoDate(date);
}

function signedBp(value: number) {
  return `${value > 0 ? "+" : ""}${value.toFixed(2)}BP`;
}

function specialBondResult(row: ParsedBondRecord) {
  const identity = `${row.bondType || "利率债"}${rollingTenor(row.tenor || "")}${row.shortName || row.bondCode || ""}`;
  const source = row.summaryMeta?.auctionSpreadText || row.summaryMeta?.displaySpreadText || "";
  const type = `${row.remark || ""}${row.summaryMeta?.rateType || ""}${row.summaryMeta?.note || ""}`;
  if (/DR/i.test(type)) {
    return source ? `${identity}净价结果为${source}` : `${identity}已发行，表内暂无可比净价结果`;
  }
  if (row.spread !== null && row.spread !== undefined && Number.isFinite(row.spread)) {
    return `${identity}一二级利差为${signedBp(row.spread)}`;
  }
  return source ? `${identity}发行结果为${source}` : `${identity}已发行，表内暂无可比一二级利差`;
}

export function rollingSpreadAnalysis(
  currentRecords: ParsedBondRecord[],
  windowRecords: ParsedBondRecord[],
  startDate: string,
  endDate: string,
  threshold = 1.5,
): RollingSpreadAnalysis {
  const supported = (row: ParsedBondRecord) => ["国债", "国开债", "口行债", "农发债"].includes(row.bondType || "");
  const weekday = (row: ParsedBondRecord) => {
    const day = new Date(`${row.tradeDate}T12:00:00`).getDay();
    return day >= 1 && day <= 5;
  };
  const benchmarkStart = shiftIsoDate(startDate, -28);
  const benchmarkEnd = shiftIsoDate(startDate, -1);
  const inCurrent = currentRecords.filter((row) => row.tradeDate >= startDate && row.tradeDate <= endDate && supported(row) && weekday(row));
  const ordinaryCurrent = inCurrent.filter((row) => !isSpecialSpreadBond(row) && row.spread !== null && row.spread !== undefined && Number.isFinite(row.spread));
  const ordinaryBenchmark = windowRecords.filter((row) => row.tradeDate >= benchmarkStart && row.tradeDate <= benchmarkEnd && supported(row) && weekday(row) && !isSpecialSpreadBond(row) && row.spread !== null && row.spread !== undefined && Number.isFinite(row.spread));
  const special = inCurrent.filter(isSpecialSpreadBond);
  const keyOf = (row: ParsedBondRecord) => `${row.bondType || ""}|${rollingTenor(row.tenor || "")}`;
  const averageByKey = (rows: ParsedBondRecord[]) => {
    const groups = new Map<string, { label: string; values: number[] }>();
    rows.forEach((row) => {
      const key = keyOf(row);
      const entry = groups.get(key) || { label: `${row.bondType}${rollingTenor(row.tenor || "")}`, values: [] };
      entry.values.push(Number(row.spread));
      groups.set(key, entry);
    });
    return new Map([...groups.entries()].map(([key, entry]) => [key, {
      label: entry.label,
      count: entry.values.length,
      average: entry.values.reduce((sum, value) => sum + value, 0) / entry.values.length,
    }]));
  };
  const current = averageByKey(ordinaryCurrent);
  const benchmark = averageByKey(ordinaryBenchmark);
  const comparable = [...current.entries()].map(([key, value]) => {
    const prior = benchmark.get(key);
    return prior ? { ...value, benchmarkAverage: prior.average, benchmarkCount: prior.count, change: value.average - prior.average } : null;
  }).filter((item): item is NonNullable<typeof item> => Boolean(item));
  const notable = comparable.filter((item) => Math.abs(item.change) > threshold).sort((a, b) => Math.abs(b.change) - Math.abs(a.change));
  const normal = comparable.filter((item) => Math.abs(item.change) <= threshold);
  const currentRanking = [...current.values()].sort((a, b) => b.average - a.average);
  const paragraphs: string[] = [];
  if (!ordinaryCurrent.length) {
    paragraphs.push("本期暂无纳入散点图口径的普通国债及政金债发行记录，无法形成四周滚动比较。");
  } else if (!comparable.length) {
    paragraphs.push(`本期共有${current.size}个品种期限组合，但前四周窗口（${benchmarkStart}至${benchmarkEnd}）暂无同品种同期限可比样本。`);
    if (currentRanking.length > 1) {
      const highest = currentRanking[0];
      const lowest = currentRanking.at(-1)!;
      paragraphs.push(`横向看，本期平均一二级利差最高为${highest.label}${signedBp(highest.average)}，最低为${lowest.label}${signedBp(lowest.average)}。`);
    }
  } else {
    paragraphs.push(`本期共有${comparable.length}个同品种同期限组合具备前四周可比样本，其中${normal.length}个处于±${threshold.toFixed(2)}BP正常波动区间，${notable.length}个变动较为突出。`);
    if (currentRanking.length > 1) {
      const highest = currentRanking[0];
      const lowest = currentRanking.at(-1)!;
      paragraphs.push(`横向看，本期平均一二级利差最高为${highest.label}${signedBp(highest.average)}，最低为${lowest.label}${signedBp(lowest.average)}。`);
    }
    if (notable.length) {
      paragraphs.push(`重点变化方面，${notable.map((item) => `${item.label}平均利差由前四周${signedBp(item.benchmarkAverage)}变为本期${signedBp(item.average)}，${item.change > 0 ? "上行" : "下行"}${Math.abs(item.change).toFixed(2)}BP`).join("；")}。`);
    } else {
      paragraphs.push(`各可比品种期限组合相对前四周均值的变化均未超过±${threshold.toFixed(2)}BP，整体属于正常波动。`);
    }
  }
  if (special.length) paragraphs.push(`特殊债券发行汇总：${special.map(specialBondResult).join("；")}。`);
  else paragraphs.push("本期无被散点图口径排除的浮息、绿色或主题债发行记录。");
  return {
    text: paragraphs.join("\n\n"), benchmarkStart, benchmarkEnd,
    comparableGroups: comparable.length, normalGroups: normal.length,
    notableGroups: notable.length, specialBonds: special.length,
  };
}

export function spreadSummary(records: ParsedBondRecord[], history: ParsedBondRecord[] = []) {
  const supported = records.filter((row) => ["国债", "国开债", "口行债", "农发债"].includes(row.bondType || ""));
  if (!supported.length) return "本周暂无符合口径的国债及政金债发行记录。";

  const tenor = (value = "") => /[DMY]$/i.test(value) ? value.toUpperCase() : `${value}Y`;
  const benchmark = (text: string) => text.includes("远期") ? "远期" : text.includes("估值") ? "估值" : "二级";
  const reference = (text: string) => {
    const match = text.match(/(?:较|平)([^()]*?)(估值|二级|远期)/);
    return match?.[1]?.trim() || "";
  };
  const resultPhrase = (row: ParsedBondRecord, text?: string, value?: number | null) => {
    const source = text || row.summaryMeta?.displaySpreadText || "";
    const numeric = value ?? row.spread;
    const target = `${reference(source)}${benchmark(source)}`;
    if (/持平|平/.test(source) && (numeric === null || numeric === undefined || numeric === 0)) return `与${target}持平`;
    if (numeric === null || numeric === undefined || !Number.isFinite(numeric)) return source || "暂无可比利差";
    if (numeric > 0) return `高${target}${numeric.toFixed(2)}BP`;
    if (numeric < 0) return `低${target}${Math.abs(numeric).toFixed(2)}BP`;
    return `与${target}持平`;
  };
  const direction = (previous: number, current: number) => {
    if (previous < 0 && current < 0) return current < previous ? "负利差走扩" : "负利差收窄";
    if (previous > 0 && current > 0) return current > previous ? "正利差走扩" : "正利差收窄";
    if (previous >= 0 && current < 0) return "由正转负";
    if (previous <= 0 && current > 0) return "由负转正";
    return current === 0 ? "收窄至持平" : "由持平转为利差";
  };
  const allCandidates = [...history, ...supported];
  const comparablePrevious = (row: ParsedBondRecord) => {
    if (!isDiscountBond(row) && !isReopenedBond(row)) return undefined;
    const embedded = row.summaryMeta?.previous;
    if (embedded && embedded.date < row.tradeDate) return embedded;
    const candidates = allCandidates.filter((candidate) => {
      if (candidate === row || candidate.tradeDate >= row.tradeDate || candidate.spread === null || candidate.spread === undefined) return false;
      if (isDiscountBond(row)) return isDiscountBond(candidate) && discountComparatorKey(candidate) === discountComparatorKey(row);
      return baseBondCode(candidate.bondCode || "") === baseBondCode(row.bondCode || "");
    }).sort((a, b) => b.tradeDate.localeCompare(a.tradeDate));
    const previous = candidates[0];
    if (!previous) return undefined;
    return {
      date: previous.tradeDate,
      code: previous.bondCode || "",
      comparisonType: isDiscountBond(row) ? "discount_comparator" as const : "same_bond" as const,
      displaySpreadText: previous.summaryMeta?.displaySpreadText || "",
      auctionSpreadText: previous.summaryMeta?.auctionSpreadText || "",
      allInText: previous.summaryMeta?.allInText || "",
      secondaryText: previous.summaryMeta?.secondaryText || "",
      note: previous.remark || "",
      spread: previous.spread ?? null,
    };
  };
  const comparableRows = (rows: ParsedBondRecord[]) => rows.filter((row) =>
    (isDiscountBond(row) || isReopenedBond(row)) && row.spread !== null && row.spread !== undefined);
  const weekStart = mondayOf([...supported].sort((a, b) => a.tradeDate.localeCompare(b.tradeDate))[0].tradeDate);
  const previousWeekDate = new Date(`${weekStart}T12:00:00`);
  previousWeekDate.setDate(previousWeekDate.getDate() - 7);
  const previousWeekStart = isoDate(previousWeekDate);
  previousWeekDate.setDate(previousWeekDate.getDate() + 4);
  const previousWeekEnd = isoDate(previousWeekDate);
  const previousWeek = history.filter((row) => row.tradeDate >= previousWeekStart && row.tradeDate <= previousWeekEnd);
  const average = (rows: ParsedBondRecord[]) => rows.reduce((sum, row) => sum + (row.spread || 0), 0) / rows.length;
  const weeklyLead = (currentRows: ParsedBondRecord[], priorRows: ParsedBondRecord[], label: string) => {
    const current = comparableRows(currentRows);
    const prior = comparableRows(priorRows);
    if (!current.length) return "";
    const currentAverage = average(current);
    if (!prior.length) return `按续发及贴现可比券口径，本周${label}平均一二级利差为${currentAverage >= 0 ? "+" : ""}${currentAverage.toFixed(2)}BP，上周暂无可比发行记录。`;
    const priorAverage = average(prior);
    const change = currentAverage - priorAverage;
    return `按续发及贴现可比券口径，本周${label}平均一二级利差为${currentAverage >= 0 ? "+" : ""}${currentAverage.toFixed(2)}BP，上周为${priorAverage >= 0 ? "+" : ""}${priorAverage.toFixed(2)}BP，整体${direction(priorAverage, currentAverage)}${Math.abs(change).toFixed(2)}BP。`;
  };
  const sections: string[] = [];
  const maturityOrder = (value = "") => {
    const numeric = Number.parseFloat(value);
    if (!Number.isFinite(numeric)) return 999;
    if (/D$/i.test(value)) return numeric / 365;
    if (/M$/i.test(value)) return numeric / 12;
    return numeric;
  };
  const treasury = supported.filter((row) => row.bondType === "国债").sort((a, b) =>
    a.tradeDate.localeCompare(b.tradeDate) || maturityOrder(a.tenor) - maturityOrder(b.tenor));
  const treasuryChanges = treasury.map((row) => {
    const previous = comparablePrevious(row);
    const previousSpread = previous?.spread;
    const currentSpread = row.spread;
    if (previousSpread === null || previousSpread === undefined || currentSpread === null || currentSpread === undefined) return null;
    return { row, previous, previousSpread, currentSpread, change: currentSpread - previousSpread, movement: direction(previousSpread, currentSpread) };
  }).filter((item): item is NonNullable<typeof item> => Boolean(item));
  if (treasuryChanges.length) {
    const weekOverview = weeklyLead(treasury, previousWeek.filter((row) => row.bondType === "国债"), "国债");
    const opening = `分期限看，${treasuryChanges.map(({ row, movement }) => `${tenor(row.tenor)}${isDiscountBond(row) ? "贴现国债" : "国债"}${movement}`).join("，")}。`;
    const details = treasuryChanges.map(({ row, previous, previousSpread, currentSpread, change, movement }) => {
      const comparator = previous.comparisonType === "discount_comparator"
        ? `对比同期限贴现国债${previous.code}（${previous.date}）`
        : `上次发行${previous.code}（${previous.date}）`;
      return `${tenor(row.tenor)}${isDiscountBond(row) ? "贴现国债" : "国债"}本次${resultPhrase(row, row.summaryMeta?.displaySpreadText, currentSpread)}，${comparator}${resultPhrase(row, previous.displaySpreadText, previousSpread)}，${movement}${Math.abs(change).toFixed(2)}BP。`;
    }).join("\n");
    sections.push(`国债\n${[weekOverview, opening, details].filter(Boolean).join("\n")}`);
  }

  const policy = supported.filter((row) => row.bondType !== "国债" && !/DR浮息债/i.test(row.summaryMeta?.rateType || row.remark || ""));
  const routeSections = ["中债招标", "上清所", "报价发行"].map((route) => {
    const material = policy.filter((row) => (row.summaryMeta?.route || row.issuanceRoute || "中债招标") === route).map((row) => {
      const previous = comparablePrevious(row);
      const previousSpread = previous?.spread;
      const currentSpread = row.spread;
      if (previousSpread === null || previousSpread === undefined || currentSpread === null || currentSpread === undefined) return null;
      const change = currentSpread - previousSpread;
      if (Math.abs(change) <= 1) return null;
      return { row, previous, previousSpread, currentSpread, change, movement: direction(previousSpread, currentSpread) };
    }).filter((item): item is NonNullable<typeof item> => Boolean(item));
    const routeRows = policy.filter((row) => (row.summaryMeta?.route || row.issuanceRoute || "中债招标") === route);
    const priorRouteRows = previousWeek.filter((row) => row.bondType !== "国债" && (row.summaryMeta?.route || row.issuanceRoute || "中债招标") === route);
    const weekOverview = weeklyLead(routeRows, priorRouteRows, `${route}政金债`);
    if (!material.length && !weekOverview) return "";
    const lead = material.length ? material.map(({ row, movement }) => `${row.bondType}${tenor(row.tenor)}品种${movement}`).join("，") + "。" : "";
    const details = material.map(({ row, previous, previousSpread, currentSpread, change, movement }) => {
      const comparator = previous.comparisonType === "discount_comparator" ? `对比同期限${row.bondType}${previous.code}（${previous.date}）` : "上次发行";
      return `${row.bondType}${tenor(row.tenor)}品种本次${resultPhrase(row, row.summaryMeta?.displaySpreadText, currentSpread)}，${comparator}${resultPhrase(row, previous.displaySpreadText, previousSpread)}，${movement}${Math.abs(change).toFixed(2)}BP。`;
    }).join("\n");
    return `${route}\n${[weekOverview, lead, details].filter(Boolean).join("\n")}`;
  }).filter(Boolean);
  if (routeSections.length) sections.push(`政金债\n${routeSections.join("\n\n")}`);

  const dr = supported.filter((row) => /DR浮息债/i.test(row.summaryMeta?.rateType || row.remark || ""));
  if (dr.length) {
    const details = dr.map((row) => {
      const meta = row.summaryMeta;
      const label = `${row.bondType}${tenor(row.tenor)}DR浮息债`;
      const text = meta?.displaySpreadText || "";
      const route = meta?.route || row.issuanceRoute || "中债招标";
      const routePhrase = route === "报价发行" ? "报价方式" : route;
      const net = text.match(/(?:低|高)(?:估值|估价)净价\([^)]*\)([\d.]+)元/);
      if (net) return `${label}通过${routePhrase}发行，缴款净价${text.startsWith("低") ? "低于" : "高于"}估值净价${Number(net[1]).toFixed(4)}元。`;
      if (/平(?:估值|估价)净价/.test(text)) return `${label}通过${routePhrase}发行，缴款净价与估值净价持平。`;
      return `${label}通过${routePhrase}发行；表内未提供可比缴款净价，本次不作净价差比较。`;
    });
    sections.push(`DR浮息债\n${details.join("\n")}`);
  }
  return sections.join("\n\n");
}
