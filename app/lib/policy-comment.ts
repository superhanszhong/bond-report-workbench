import type { ParsedBondRecord, SpreadSummaryMeta } from "./workbench";

export type PolicyShortComment = {
  key: string;
  tradeDate: string;
  bondCode: string;
  displayCode: string;
  tenor: string;
  issuer: string;
  currentSpread: number | null;
  previousSpread: number | null;
  previousCode: string;
  movement: string;
  firstLine: string;
  secondLine: string;
  text: string;
  warning?: string;
};

export type PolicyCommentDraft = {
  id: string;
  tradeDate: string;
  bondCode: string;
  shortName: string;
  issuer: string;
  bondType: string;
  tenor: string;
  rateType: string;
  route: string;
  benchmarkType: string;
  benchmarkValue: string;
  finalValue: string;
};

export const POLICY_FLOAT_RATE_OPTIONS = ["", "LPR浮息债", "DR007浮息债", "DR001浮息债"] as const;

export type PolicyDraftResult = {
  draft: PolicyCommentDraft;
  comment: PolicyShortComment | null;
  missing: string[];
};

type PreviousSnapshot = NonNullable<SpreadSummaryMeta["previous"]>;

function normalized(text = "") {
  return text.replaceAll("（", "(").replaceAll("）", ")").replaceAll("％", "%").trim();
}

function compactNumber(value: number, decimals = 4) {
  return String(Number(value.toFixed(decimals)));
}

function firstNumber(text = "") {
  const match = normalized(text).match(/-?\d+(?:\.\d+)?/);
  return match ? Number(match[0]) : null;
}

function percentNumber(text = "") {
  const value = firstNumber(text);
  if (value === null) return null;
  return Math.abs(value) < 0.2 ? value * 100 : value;
}

function baseBondCode(code = "") {
  return code.replace(/\.(?:IB|SH|SZ)$/i, "").replace(/[XZ]\d*$/i, "");
}

function isReopenedBondCode(code = "") {
  return /[XZ]\d*$/i.test(code.replace(/\.(?:IB|SH|SZ)$/i, ""));
}

function displayBondCode(code = "") {
  const clean = code.replace(/\.(?:IB|SH|SZ)$/i, "");
  return /^09/.test(clean) ? clean : clean.replace(/Z(\d*)$/i, "X$1");
}

function issuerLabel(row: ParsedBondRecord) {
  const source = `${row.bondType || ""}${row.issuer || ""}`;
  if (/农发|农业发展/.test(source)) return "农发";
  if (/口行|进出口/.test(source)) return "口行";
  if (/国开|国家开发/.test(source)) return "国开";
  return row.bondType?.replace(/债$/, "") || "政金";
}

function tenorLabel(value = "") {
  const clean = value.trim().toUpperCase();
  return /[DMY]$/.test(clean) ? clean : `${clean}Y`;
}

function isPolicy(row: ParsedBondRecord) {
  return ["国开债", "口行债", "农发债"].includes(row.bondType || "") || /国家开发|进出口|农业发展/.test(row.issuer || "");
}

function isDr(row: ParsedBondRecord) {
  return isReopenedBondCode(row.bondCode || "") && /DR(?:001|007)?浮息债/i.test(`${row.summaryMeta?.rateType || ""}${row.remark || ""}`);
}

function selectedRateType(row: ParsedBondRecord) {
  const source = `${row.summaryMeta?.rateType || ""}${row.remark || ""}`;
  if (/DR0{1,2}1/i.test(source)) return "DR001浮息债";
  if (/DR007/i.test(source)) return "DR007浮息债";
  if (/DR浮息债/i.test(source)) return "DR007浮息债";
  if (/LPR/i.test(source)) return "LPR浮息债";
  return "";
}

function bpValue(text = "") {
  const source = normalized(text);
  if (!source) return null;
  const direct = source.match(/^\s*(-?\d+(?:\.\d+)?)/);
  if (direct) return Number(direct[1]);
  const magnitude = source.match(/(\d+(?:\.\d+)?)\s*BP/i);
  if (/平/.test(source) && !magnitude) return 0;
  if (!magnitude) return null;
  return /低|负/.test(source) ? -Number(magnitude[1]) : Number(magnitude[1]);
}

function comparison(text = "", fallbackQuote = "") {
  const source = normalized(text);
  const compared = source.match(/较([^()]+?)(?:\(([^()]*)\))?\)?(?:\s|$)/);
  const direct = source.match(/(?:高|低|平|与)([^()]*?(?:二级|估值|中间价|价格|曲线))(?:\(([^()]*)\))?/);
  let label = (compared?.[1] || direct?.[1] || "二级").trim();
  label = label.replace(/^较/, "").replace(/净价$/, "").trim();
  const embeddedQuote = compared?.[2] || direct?.[2] || "";
  const quoteValue = percentNumber(embeddedQuote) ?? percentNumber(fallbackQuote);
  return { label, quote: quoteValue === null ? "" : compactNumber(quoteValue, 4) };
}

function netValue(text = "") {
  const source = normalized(text).replace(/^\(+|\)+$/g, "");
  if (!source || !/净价|元/.test(source)) return null;
  const side = source.includes("低") ? -1 : source.includes("高") ? 1 : 0;
  const label = source.match(/[高低平]([^()]*?净价)/)?.[1]?.trim() || "估价净价";
  const benchmark = source.match(/\((\d+(?:\.\d+)?)\)/)?.[1] || "";
  const afterBenchmark = source.match(/\)(\d+(?:\.\d+)?)\s*元/)?.[1];
  const allNumbers = [...source.matchAll(/\d+(?:\.\d+)?/g)].map((match) => Number(match[0]));
  const magnitude = afterBenchmark ? Number(afterBenchmark) : allNumbers.at(-1) ?? null;
  if (magnitude === null) return null;
  return { value: side * magnitude, magnitude, side, label, benchmark };
}

function direction(previous: number, current: number) {
  const tolerance = 0.0000001;
  if (Math.abs(previous - current) <= tolerance) return "持平";
  if (Math.abs(previous) <= tolerance) return Math.abs(current) <= tolerance ? "持平" : "反转";
  if (Math.abs(current) <= tolerance) return "收窄";
  if (Math.sign(previous) !== Math.sign(current)) return "反转";
  return Math.abs(current) > Math.abs(previous) ? "走阔" : "收窄";
}

function previousFor(row: ParsedBondRecord, records: ParsedBondRecord[]): PreviousSnapshot | undefined {
  if (row.summaryMeta?.previous && row.summaryMeta.previous.date < row.tradeDate) return row.summaryMeta.previous;
  const base = baseBondCode(row.bondCode || "");
  const previous = records.filter((candidate) => candidate !== row && candidate.tradeDate < row.tradeDate && baseBondCode(candidate.bondCode || "") === base)
    .sort((a, b) => b.tradeDate.localeCompare(a.tradeDate))[0];
  if (!previous) return undefined;
  return {
    date: previous.tradeDate,
    code: previous.bondCode || "",
    comparisonType: "same_bond",
    displaySpreadText: previous.summaryMeta?.displaySpreadText || "",
    auctionSpreadText: previous.summaryMeta?.auctionSpreadText || "",
    allInText: previous.summaryMeta?.allInText || "",
    secondaryText: previous.summaryMeta?.secondaryText || "",
    note: previous.remark || "",
    spread: previous.spread ?? null,
  };
}

function issueDescription(row: ParsedBondRecord) {
  const code = row.bondCode || "";
  const reopened = isReopenedBondCode(code);
  const action = reopened ? "增发" : "新发";
  const issuer = issuerLabel(row);
  const maturity = tenorLabel(row.tenor || "");
  const rateType = selectedRateType(row);
  const route = row.summaryMeta?.route || row.issuanceRoute || "";
  const clearing = /^DR/.test(rateType) || /^09/.test(code) || /清发|上清所/.test(`${row.shortName || ""}${route}${row.remark || ""}`);
  return `今日${action}${issuer}${clearing ? "清发" : ""}${maturity}${rateType ? ` ${rateType}` : ""}`;
}

function ordinaryComment(row: ParsedBondRecord, records: ParsedBondRecord[]): PolicyShortComment {
  const meta = row.summaryMeta;
  const currentText = meta?.auctionSpreadText || "";
  const current = bpValue(currentText) ?? (() => {
    const winning = percentNumber(meta?.winningRateText || "");
    const secondary = percentNumber(meta?.secondaryText || "");
    return winning === null || secondary === null ? null : Number(((winning - secondary) * 100).toFixed(2));
  })();
  const reopened = isReopenedBondCode(row.bondCode || "");
  const previous = reopened ? previousFor(row, records) : undefined;
  const previousText = normalized(previous?.auctionSpreadText || "");
  const previousValue = bpValue(previousText);
  const compare = comparison(currentText, meta?.secondaryText || "");
  const side = current === null ? "" : current > 0 ? "高" : current < 0 ? "低" : "平";
  const resultLabel = (meta?.route || row.issuanceRoute) === "报价发行" ? "报价利率" : "中标利率";
  const winning = percentNumber(meta?.winningRateText || "");
  const resultPrefix = `${resultLabel}${winning === null ? "" : `${compactNumber(winning, 4)}%`}`;
  const firstLine = current === null
    ? "缺少中标利差，无法生成"
    : current === 0
      ? `${resultPrefix} 平${compare.label}${compare.quote ? `(${compare.quote})` : ""}`
      : `${resultPrefix} ${side}${compare.label}${compare.quote ? `(${compare.quote})` : ""}${Math.abs(current).toFixed(2)}BP`;
  const movement = !reopened ? "不判断" : current !== null && previousValue !== null ? direction(previousValue, current) : "待确认";
  const secondLine = reopened
    ? `${issueDescription(row)},利差${movement}(${previousText ? `上次${previousText}` : "暂无上次同券发行记录"})`
    : issueDescription(row);
  return {
    key: `${row.tradeDate}|${row.bondCode || ""}`,
    tradeDate: row.tradeDate,
    bondCode: row.bondCode || "",
    displayCode: displayBondCode(row.bondCode || ""),
    tenor: tenorLabel(row.tenor || ""),
    issuer: issuerLabel(row),
    currentSpread: current,
    previousSpread: previousValue,
    previousCode: displayBondCode(previous?.code || ""),
    movement,
    firstLine,
    secondLine,
    text: `${firstLine}\n${secondLine}`,
    warning: current === null ? "未识别到中标比二级利差" : reopened && previousValue === null ? "未找到可解析的上次同券利差" : undefined,
  };
}

function drComment(row: ParsedBondRecord, records: ParsedBondRecord[]): PolicyShortComment {
  const meta = row.summaryMeta;
  const currentNet = netValue(meta?.auctionSpreadText || meta?.displaySpreadText || "");
  const previous = previousFor(row, records);
  const previousText = normalized(previous?.auctionSpreadText || previous?.displaySpreadText || "");
  const previousNet = netValue(previousText);
  const payment = firstNumber(meta?.allInText || "");
  const benchmark = currentNet?.benchmark || (() => {
    const value = firstNumber(meta?.secondaryText || "");
    return value === null ? "" : compactNumber(value, 4);
  })();
  const side = !currentNet ? "" : currentNet.side > 0 ? "高" : currentNet.side < 0 ? "低" : "平";
  const firstLine = !currentNet || payment === null
    ? "缺少中标净价或估价净价差，无法生成"
    : `中标净价${compactNumber(payment, 4)}元 ${side}${currentNet.label}${benchmark ? `(${benchmark})` : ""}${compactNumber(currentNet.magnitude, 4)}元`;
  const movement = currentNet && previousNet ? direction(previousNet.value, currentNet.value) : "不判断";
  const previousCompact = previousNet
    ? `${previousNet.side > 0 ? "高" : previousNet.side < 0 ? "低" : "平"}${previousNet.label}${compactNumber(previousNet.magnitude, 4)}元`
    : "";
  const secondLine = previousCompact
    ? `${issueDescription(row)},利差${movement}(上次${previousCompact})`
    : `${issueDescription(row)},本次不作利差变化判断(上次发行无可比净价差)`;
  return {
    key: `${row.tradeDate}|${row.bondCode || ""}`,
    tradeDate: row.tradeDate,
    bondCode: row.bondCode || "",
    displayCode: displayBondCode(row.bondCode || ""),
    tenor: tenorLabel(row.tenor || ""),
    issuer: issuerLabel(row),
    currentSpread: currentNet?.value ?? null,
    previousSpread: previousNet?.value ?? null,
    previousCode: displayBondCode(previous?.code || ""),
    movement,
    firstLine,
    secondLine,
    text: `${firstLine}\n${secondLine}`,
    warning: !currentNet || payment === null ? "未识别到中标净价或估价净价差" : !previousNet ? "未找到可解析的上次同券净价差" : undefined,
  };
}

export function policyCommentDates(records: ParsedBondRecord[]) {
  return [...new Set(records.filter(isPolicy).map((row) => row.tradeDate))].sort().reverse();
}

export function policyComments(records: ParsedBondRecord[], date: string) {
  return records.filter((row) => row.tradeDate === date && isPolicy(row))
    .map((row) => isDr(row) ? drComment(row, records) : ordinaryComment(row, records));
}

function benchmarkKind(text = "", dr = false) {
  const source = normalized(text);
  if (/二级/.test(source)) return dr ? "二级" : "二级";
  if (/中间价/.test(source)) return "中间价";
  if (/价格/.test(source) && !/估价/.test(source)) return "价格";
  if (/估值曲线/.test(source)) return "估值曲线";
  if (/估值|估价/.test(source)) return dr ? "估价" : "估值";
  return dr ? "估价" : "二级";
}

function latestSameBond(row: ParsedBondRecord, history: ParsedBondRecord[]) {
  const base = baseBondCode(row.bondCode || "");
  return history.filter((candidate) => candidate.tradeDate <= row.tradeDate && baseBondCode(candidate.bondCode || "") === base)
    .sort((a, b) => b.tradeDate.localeCompare(a.tradeDate))[0];
}

function historicalRateType(row: ParsedBondRecord, history: ParsedBondRecord[]) {
  const latest = latestSameBond(row, history);
  const source = `${latest?.remark || ""}${latest?.summaryMeta?.note || ""}${latest?.summaryMeta?.previous?.note || ""}`;
  if (/DR0{1,2}1/i.test(source)) return "DR001浮息债";
  if (/DR007/i.test(source)) return "DR007浮息债";
  if (/DR/i.test(source)) return "DR007浮息债";
  if (/LPR/i.test(source)) return "LPR浮息债";
  return "";
}

function inputRate(value: unknown) {
  if (value === null || value === undefined || value === "" || value === "--") return "";
  const parsed = percentNumber(String(value));
  return parsed === null ? "" : compactNumber(parsed, 4);
}

export function createPolicyCommentDrafts(planRecords: ParsedBondRecord[], history: ParsedBondRecord[], existing: PolicyCommentDraft[] = []) {
  const existingById = new Map(existing.map((draft) => [draft.id, draft]));
  return planRecords.filter(isPolicy).map((row) => {
    const id = `${row.tradeDate}|${row.bondCode || row.shortName || ""}`;
    const previous = latestSameBond(row, history);
    const previousDraft = existingById.get(id);
    const rateType = previousDraft?.rateType ?? historicalRateType(row, history);
    const drPricing = isReopenedBondCode(row.bondCode || "") && /^DR(?:001|007)?浮息债$/i.test(rateType);
    const previousText = previous?.summaryMeta?.auctionSpreadText || previous?.summaryMeta?.displaySpreadText || "";
    const raw = row.raw || {};
    return {
      id,
      tradeDate: row.tradeDate,
      bondCode: row.bondCode || "",
      shortName: row.shortName || "",
      issuer: row.issuer || "",
      bondType: row.bondType || "",
      tenor: tenorLabel(row.tenor || ""),
      rateType,
      route: row.issuanceRoute || "中债招标",
      benchmarkType: previousDraft?.benchmarkType || benchmarkKind(previousText, drPricing),
      benchmarkValue: previousDraft?.benchmarkValue || "",
      finalValue: previousDraft?.finalValue || (drPricing ? "" : inputRate(raw.加权利率 ?? raw.发行利率)),
    } satisfies PolicyCommentDraft;
  });
}

function draftRecord(draft: PolicyCommentDraft): ParsedBondRecord | null {
  if (!draft.finalValue.trim() || !draft.benchmarkValue.trim()) return null;
  const finalValue = Number(draft.finalValue);
  const benchmarkValue = Number(draft.benchmarkValue);
  if (!Number.isFinite(finalValue) || !Number.isFinite(benchmarkValue)) return null;
  const dr = isReopenedBondCode(draft.bondCode) && /^DR(?:001|007)?浮息债$/i.test(draft.rateType);
  let auctionSpreadText = "";
  let allInText = "";
  if (dr) {
    const difference = Number((finalValue - benchmarkValue).toFixed(4));
    const side = difference > 0 ? "高" : difference < 0 ? "低" : "平";
    const label = `${draft.benchmarkType.replace(/净价$/, "")}净价`;
    auctionSpreadText = `${side}${label}(${compactNumber(benchmarkValue, 4)})${compactNumber(Math.abs(difference), 4)}元`;
    allInText = `${compactNumber(finalValue, 4)}元`;
  } else {
    const difference = Number(((finalValue - benchmarkValue) * 100).toFixed(2));
    auctionSpreadText = `${difference > 0 ? "+" : ""}${difference.toFixed(2)}(较${draft.benchmarkType}(${compactNumber(benchmarkValue, 4)}))`;
  }
  return {
    tradeDate: draft.tradeDate,
    bondCode: draft.bondCode,
    shortName: draft.shortName,
    issuer: draft.issuer,
    bondType: draft.bondType,
    issuanceRoute: draft.route,
    tenor: draft.tenor,
    remark: draft.rateType,
    summaryMeta: {
      baseCode: baseBondCode(draft.bondCode),
      route: draft.route,
      rateType: draft.rateType,
      displaySpreadText: auctionSpreadText,
      auctionSpreadText,
      allInText,
      secondaryText: dr ? `${compactNumber(benchmarkValue, 4)}元` : `${compactNumber(benchmarkValue, 4)}%`,
      winningRateText: dr ? "" : `${compactNumber(finalValue, 4)}%`,
      note: draft.rateType,
      proceeds: "",
    },
  };
}

export function policyDraftResults(drafts: PolicyCommentDraft[], history: ParsedBondRecord[]): PolicyDraftResult[] {
  const completeRecords = drafts.map(draftRecord).filter((row): row is ParsedBondRecord => Boolean(row));
  const combined = [...history, ...completeRecords];
  const resultMap = new Map<string, PolicyShortComment>();
  completeRecords.forEach((record) => {
    const result = policyComments(combined, record.tradeDate).find((item) => item.bondCode === record.bondCode);
    if (result) resultMap.set(`${record.tradeDate}|${record.bondCode}`, result);
  });
  return drafts.map((draft) => {
    const missing: string[] = [];
    const drPricing = isReopenedBondCode(draft.bondCode) && /^DR(?:001|007)?浮息债$/i.test(draft.rateType);
    if (!draft.benchmarkValue.trim()) missing.push(drPricing ? "估价净价" : "二级/估值");
    if (!draft.finalValue.trim()) missing.push(drPricing ? "中标净价" : "最终中标率");
    return { draft, comment: resultMap.get(`${draft.tradeDate}|${draft.bondCode}`) || null, missing };
  });
}
