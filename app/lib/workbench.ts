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
  raw?: Record<string, unknown>;
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

function excelDate(value: unknown): Date | null {
  if (value instanceof Date && !Number.isNaN(value.valueOf())) return value;
  if (typeof value === "number") {
    const parsed = XLSX.SSF.parse_date_code(value);
    if (parsed) return new Date(parsed.y, parsed.m - 1, parsed.d);
  }
  const text = clean(value).slice(0, 10).replace(/[/.]/g, "-");
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

function sheetRows(workbook: XLSX.WorkBook) {
  const sheet = workbook.Sheets[workbook.SheetNames[0]];
  return XLSX.utils.sheet_to_json<unknown[]>(sheet, { header: 1, raw: true, defval: null });
}

function headerIndex(rows: unknown[][], aliases: Record<string, string[]>) {
  let bestRow = -1;
  let bestIndex: Record<string, number> = {};
  let bestScore = -1;
  rows.slice(0, 16).forEach((row, rowIndex) => {
    const cells = new Map<string, number>();
    (row || []).forEach((value, col) => {
      const label = clean(value).replace(/\s+/g, "");
      if (label) cells.set(label, col);
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

function routeFromRemark(remark: string) {
  if (/报价发行|前台报价/.test(remark)) return "报价发行";
  if (/清发|上清所/.test(remark)) return "上清所";
  return "中债招标";
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

export async function parseSpreadFile(file: File) {
  const workbook = XLSX.read(await file.arrayBuffer(), { type: "array", cellDates: true });
  const aliases = {
    发行日期: ["发行日期", "招标日"], 代码: ["代码", "债券代码"], 简称: ["简称", "债券简称"],
    期限: ["期限"], 发行人: ["发行人"], 备注: ["备注"], 利差: ["综收-二级", "综收－二级"],
    发行量: ["发行量", "发行量(亿元)"],
  };
  const records: ParsedBondRecord[] = [];
  workbook.SheetNames.forEach((sheetName) => {
    const rows = XLSX.utils.sheet_to_json<unknown[]>(workbook.Sheets[sheetName], { header: 1, raw: true, defval: null });
    const { row: headerRow, index } = headerIndex(rows, aliases);
    if (headerRow < 0 || index.发行日期 === undefined || index.代码 === undefined || index.利差 === undefined) return;
    for (let r = headerRow + 1; r < rows.length; r += 1) {
      const raw = rowObject(rows[r], index);
      const date = excelDate(raw.发行日期);
      const bondCode = clean(raw.代码).replace(/\.IB$/i, "");
      if (!date || !bondCode) continue;
      const issuer = clean(raw.发行人);
      const remark = clean(raw.备注);
      records.push({
        tradeDate: isoDate(date), bondCode, shortName: clean(raw.简称), issuer,
        bondType: ISSUER_MAP[issuer] || issuer, issuanceRoute: routeFromRemark(remark),
        tenor: clean(raw.期限), amount: numberValue(raw.发行量), spread: numberValue(raw.利差), remark, raw,
      });
    }
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

export function spreadSummary(records: ParsedBondRecord[]) {
  const ordinary = records.filter((row) => row.spread !== null && row.spread !== undefined &&
    ["国债", "国开债", "口行债", "农发债"].includes(row.bondType || "") &&
    !/绿债|绿色|主题债|浮息债/.test(row.remark || ""));
  if (!ordinary.length) return "本周暂无符合口径的国债及政金债发行利差记录。";
  const describe = (rows: ParsedBondRecord[]) => {
    const positive = rows.filter((row) => Number(row.spread) > 0).length;
    const negative = rows.filter((row) => Number(row.spread) < 0).length;
    const avg = rows.reduce((sum, row) => sum + Number(row.spread || 0), 0) / rows.length;
    const direction = positive && negative ? "正负利差并存" : positive ? "整体高于二级" : negative ? "整体低于二级" : "整体与二级基本持平";
    const details = rows.map((row) => {
      const value = Number(row.spread || 0);
      const comparison = value > 0 ? `高二级${value.toFixed(2)}BP` : value < 0 ? `低二级${Math.abs(value).toFixed(2)}BP` : "与二级持平";
      return `${row.bondType || ""}${row.tenor || ""}${row.shortName ? `（${row.shortName}）` : ""}${comparison}`;
    }).join("，");
    return `${direction}，平均利差${avg.toFixed(2)}BP。${details}。`;
  };
  const paragraphs: string[] = [];
  const treasury = ordinary.filter((row) => row.bondType === "国债");
  if (treasury.length) paragraphs.push(`国债：${describe(treasury)}`);
  const policy = ordinary.filter((row) => row.bondType !== "国债");
  if (policy.length) {
    const routeParagraphs = ["中债招标", "上清所", "报价发行"].map((route) => {
      const routeRows = policy.filter((row) => (row.issuanceRoute || "中债招标") === route);
      return routeRows.length ? `${route}：${describe(routeRows)}` : "";
    }).filter(Boolean);
    paragraphs.push(`政金债：\n${routeParagraphs.join("\n")}`);
  }
  return paragraphs.join("\n\n");
}
