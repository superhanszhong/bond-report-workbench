"use client";

import { useEffect, useMemo, useRef, useState } from "react";
import {
  BarChart3, CalendarDays, Check, ChevronLeft, ChevronRight, CircleAlert,
  Copy, Download, FileSpreadsheet, FileText, House, LoaderCircle, RefreshCw, Upload, X,
} from "lucide-react";
import {
  fridayOf, maturityDailyTotals, maturityKind, maturityWeekStart, mondayOf, parseIssuancePlanFile, parseLocalBondFile, parseMaturityFile,
  inferredRateType, isSpecialSpreadBond, parseSpreadFile, ParsedBondRecord, rateMaturityBreakdown, recordsForSpreadMetric, resolveSpreadBp,
  rollingSpreadAnalysis, SpreadMetric,
} from "./lib/workbench";
import { buildWeeklyReportBlob, reportDataWarnings } from "./lib/report";
import {
  createPolicyCommentDrafts, policyDraftResults, PolicyCommentDraft, POLICY_FLOAT_RATE_OPTIONS,
} from "./lib/policy-comment";
import { LOCAL_STORAGE_MODE, workbenchRequest } from "./lib/workbench-request";
import { recordKey } from "./lib/record-merge";

type StoredRecord = ParsedBondRecord & {
  id: number;
  dataset_type: string;
  trade_date: string;
  week_start: string;
  bond_code?: string;
  short_name?: string;
  full_name?: string;
  bond_type?: string;
  issuance_route?: string;
  bid_time?: string;
  floor_rate?: number | null;
  distribution_date?: string;
  raw_json?: string;
};

type ImportRow = {
  id: string;
  dataset_type: string;
  trade_date: string;
  file_name: string;
  record_count: number;
  created_at: string;
};

type WeekData = {
  imports: ImportRow[];
  records: StoredRecord[];
  draft?: { summary_text?: string; review_text?: string; updated_at?: string } | null;
};

type LatestDates = { local_bond?: string; spread?: string; maturity?: string; issuance_plan?: string };
type DatasetType = "local_bond" | "spread" | "maturity" | "issuance_plan";
const ASSET_BASE = import.meta.env?.BASE_URL || "/";

const tabGroups = [
  { label: "工作台", items: [["overview", "首页"]] },
  { label: "地方债", items: [["local", "日表转换器"]] },
  { label: "利率债", items: [["spread", "一二级利差"], ["comment", "发行小结生成"], ["report", "周报生成"]] },
] as const;

type TabKey = (typeof tabGroups)[number]["items"][number][0];

const viewCopy: Record<TabKey, { eyebrow: string; title: string; description: string }> = {
  overview: { eyebrow: "ISSUANCE DESK", title: "利率债一级工作台", description: "本周数据总览与文件入口" },
  local: { eyebrow: "LOCAL BOND TOOL", title: "地方债日表转换器", description: "在浏览器本地转换 DM 地方债文件" },
  spread: { eyebrow: "PRIMARY / SECONDARY", title: "一二级利差分析", description: "按自选日期区间生成散点图" },
  comment: { eyebrow: "ISSUANCE SUMMARY", title: "政金债发行小结生成", description: "从一二级利差历史表自动匹配上次同券发行" },
  report: { eyebrow: "FINAL OUTPUT", title: "客户版周报生成", description: "使用发行、到期及一二级数据生成 Word" },
};

function formatMd(value: string) {
  const date = new Date(`${value}T12:00:00`);
  return `${date.getMonth() + 1}/${date.getDate()}`;
}

function displayWeek(start: string) {
  return `${formatMd(start)}—${formatMd(fridayOf(start))}`;
}

function shiftWeek(start: string, offset: number) {
  const date = new Date(`${start}T12:00:00`);
  date.setDate(date.getDate() + offset * 7);
  return mondayOf(date);
}

function shiftDate(value: string, offset: number) {
  const date = new Date(`${value}T12:00:00`);
  date.setDate(date.getDate() + offset);
  return isoDateValue(date);
}

function isoDateValue(date: Date) {
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, "0");
  const day = String(date.getDate()).padStart(2, "0");
  return `${year}-${month}-${day}`;
}

function normalize(row: StoredRecord): ParsedBondRecord {
  let raw: Record<string, unknown> = row.raw || {};
  if (row.raw_json) {
    try { raw = JSON.parse(row.raw_json) as Record<string, unknown>; } catch { raw = {}; }
  }
  const storedMeta = raw.__summaryMeta && typeof raw.__summaryMeta === "object"
    ? raw.__summaryMeta as ParsedBondRecord["summaryMeta"] : undefined;
  const rateType = inferredRateType({ bondCode: row.bond_code || row.bondCode, remark: row.remark, summaryMeta: storedMeta });
  const fallbackRoute = row.remark && /报价发行|前台报价/.test(row.remark)
    ? "报价发行" : /^09/.test(row.bond_code || row.bondCode || "") ? "上清所" : (row.issuance_route || row.issuanceRoute);
  const spreadResolution = row.dataset_type === "spread" ? resolveSpreadBp({
    provided: raw["利差"] ?? row.spread,
    allIn: raw["综收"], secondary: raw["二级"], winningRate: raw["中标利率"],
    issuer: row.issuer || "", remark: `${row.remark || ""} ${rateType}`,
  }) : null;
  if (spreadResolution && !raw.__spreadAudit) raw = { ...raw, __spreadAudit: spreadResolution.audit };
  return {
    tradeDate: row.trade_date || row.tradeDate,
    bondCode: row.bond_code || row.bondCode,
    shortName: row.short_name || row.shortName,
    fullName: row.full_name || row.fullName,
    issuer: row.issuer,
    region: row.region,
    bondType: row.bond_type || row.bondType,
    issuanceRoute: storedMeta?.route || fallbackRoute,
    venue: row.venue,
    bidTime: row.bid_time || row.bidTime,
    tenor: row.tenor,
    amount: row.amount,
    spread: spreadResolution ? spreadResolution.spread : row.spread,
    floorRate: row.floor_rate ?? row.floorRate,
    fee: row.fee,
    distributionDate: row.distribution_date || row.distributionDate,
    remark: row.remark,
    summaryMeta: storedMeta ? { ...storedMeta, rateType } : undefined,
    raw,
  };
}

function tone(type: string) {
  return type === "国债" ? "treasury" : type === "国开债" ? "cdb" : type === "口行债" ? "exim" : "adbc";
}

function spreadAuditStatus(row: ParsedBondRecord) {
  const audit = row.raw?.__spreadAudit;
  if (!audit || typeof audit !== "object") return "provided";
  return String((audit as Record<string, unknown>).status || "provided");
}

function maturityInfo(tenor: string) {
  const text = String(tenor || "").trim().toUpperCase();
  const value = Number.parseFloat(text);
  if (!Number.isFinite(value)) return null;
  if (text.includes("D")) {
    if (value <= 45) return { key: "1M", label: "1M", order: 1 / 12 };
    if (value <= 120) return { key: "3M", label: "3M", order: 0.25 };
    if (value <= 240) return { key: "6M", label: "6M", order: 0.5 };
    if (value <= 300) return { key: "9M", label: "9M", order: 0.75 };
    const years = value / 365;
    const rounded = Math.round(years);
    if (Math.abs(years - rounded) <= 0.08) return { key: `${rounded}Y`, label: `${rounded}Y`, order: rounded };
    const label = `${Number(years.toFixed(2))}Y`;
    return { key: label, label, order: years };
  }
  const rounded = Math.round(value);
  const years = Math.abs(value - rounded) <= 0.08 ? rounded : Number(value.toFixed(2));
  const label = `${years}Y`;
  return { key: label, label, order: years };
}

function niceStep(raw: number) {
  const safe = Math.max(raw, 0.0001);
  const exponent = Math.floor(Math.log10(safe));
  const fraction = safe / 10 ** exponent;
  const nice = fraction <= 1 ? 1 : fraction <= 2 ? 2 : fraction <= 5 ? 5 : 10;
  return nice * 10 ** exponent;
}

function SpreadChart({ records, svgRef, startDate, endDate, metricLabel }: { records: ParsedBondRecord[]; svgRef: React.RefObject<SVGSVGElement | null>; startDate: string; endDate: string; metricLabel: string }) {
  const issuerTypes = ["国债", "国开债", "口行债", "农发债"];
  const eligible = records.filter(row => issuerTypes.includes(row.bondType || "") && (() => {
    const day = new Date(`${row.tradeDate}T12:00:00`).getDay();
    return day >= 1 && day <= 5;
  })());
  const excluded = eligible.filter(isSpecialSpreadBond);
  const normalized = eligible.filter(row => row.spread !== null && row.spread !== undefined && !isSpecialSpreadBond(row)).map((row, sourceIndex) => {
    const maturity = maturityInfo(row.tenor || "");
    return maturity ? { ...row, ...maturity, sourceIndex } : null;
  }).filter((row): row is NonNullable<typeof row> => Boolean(row));
  const maturities = [...new Map(normalized.map(row => [row.key, { key: row.key, label: row.label, order: row.order }])).values()].sort((a, b) => a.order - b.order);
  const maturityIndex = new Map(maturities.map((item, index) => [item.key, index]));
  const dates = Array.from(new Set(normalized.map(row => row.tradeDate))).sort();
  const values = normalized.map(row => Number(row.spread));
  const avg = normalized.length ? values.reduce((a, b) => a + b, 0) / normalized.length : 0;
  const minSpread = Math.min(...values, 0);
  const maxSpread = Math.max(...values, 0);
  const tickStep = niceStep((maxSpread - minSpread) / 7);
  let yMin = Math.floor((minSpread - tickStep * 0.5) / tickStep) * tickStep;
  let yMax = Math.ceil((maxSpread + tickStep * 0.5) / tickStep) * tickStep;
  if (yMin === yMax) { yMin -= tickStep; yMax += tickStep; }
  const ticks: number[] = [];
  for (let tick = yMin; tick <= yMax + tickStep / 10; tick += tickStep) ticks.push(Math.abs(tick) < tickStep / 100 ? 0 : Number(tick.toFixed(6)));
  const chartLeft = 105;
  const chartRight = 1615;
  const chartTop = 195;
  const chartBottom = 870;
  const categoryWidth = (chartRight - chartLeft) / Math.max(1, maturities.length);
  const xCenter = (key: string) => chartLeft + categoryWidth * ((maturityIndex.get(key) || 0) + 0.5);
  const dateOffset = (date: string) => dates.length <= 1 ? 0 : (((dates.indexOf(date) / (dates.length - 1)) - 0.5) * categoryWidth * 0.52);
  const varietyOffset: Record<string, number> = { 国债: -13, 国开债: -4, 口行债: 4, 农发债: 13 };
  const y = (value: number) => chartTop + ((yMax - value) / (yMax - yMin)) * (chartBottom - chartTop);
  const colors: Record<string, string> = { treasury: "#F4CDC3", cdb: "#FAE1CC", exim: "#E5E5E3", adbc: "#D6EFF5" };
  const darkColors: Record<string, string> = { treasury: "#C18476", cdb: "#C99A76", exim: "#A6A6A2", adbc: "#8CBCC7" };
  const points = normalized.map(row => ({ ...row, cx: xCenter(row.key) + dateOffset(row.tradeDate) + (varietyOffset[row.bondType || ""] || 0), cy: y(Number(row.spread)) }));
  type Box = { x: number; y: number; w: number; h: number };
  const placed: Box[] = [];
  const overlaps = (a: Box, b: Box) => a.x < b.x + b.w + 6 && a.x + a.w + 6 > b.x && a.y < b.y + b.h + 6 && a.y + a.h + 6 > b.y;
  const callouts = [...points.filter(row => Number(row.spread) > 0).sort((a, b) => Number(b.spread) - Number(a.spread)), ...points.filter(row => Number(row.spread) <= -1).sort((a, b) => Number(a.spread) - Number(b.spread))].map((row, labelIndex) => {
    const label = `${row.shortName || row.bondCode || ""}  ${Number(row.spread) > 0 ? "+" : ""}${Number(row.spread).toFixed(2)}bp`;
    const w = Math.max(132, label.length * 9.7 + 34);
    const h = 36;
    const candidates: { x: number; y: number }[] = [];
    const sideFirst = (row.sourceIndex + labelIndex) % 2 === 0;
    for (let lane = 0; lane < 8; lane += 1) {
      const vertical = 18 + lane * 40;
      const xs = sideFirst ? [row.cx + 16, row.cx - w - 16] : [row.cx - w - 16, row.cx + 16];
      xs.forEach(x => candidates.push({ x, y: row.cy - h - vertical }));
      [...xs].reverse().forEach(x => candidates.push({ x, y: row.cy + vertical }));
    }
    let best: Box | undefined;
    let bestScore = Number.POSITIVE_INFINITY;
    candidates.forEach(candidate => {
      const trial = { x: Math.min(Math.max(candidate.x, chartLeft + 8), chartRight - w - 8), y: Math.min(Math.max(candidate.y, chartTop + 8), chartBottom - h - 8), w, h };
      const score = placed.filter(box => overlaps(trial, box)).length * 10000 + Math.abs(trial.x + w / 2 - row.cx) + Math.abs(trial.y + h / 2 - row.cy);
      if (score < bestScore) { best = trial; bestScore = score; }
    });
    const box = best || { x: row.cx + 16, y: row.cy - h - 14, w, h };
    placed.push(box);
    return { row, label, box };
  });
  return (
    <div className="chart-wrap">
      <svg ref={svgRef} className="spread-chart" viewBox="0 0 1680 1050" role="img" aria-label="本周国债、政金债发行一二级利差散点图">
        <defs>
          <filter id="bubble-shadow" x="-30%" y="-40%" width="160%" height="190%">
            <feDropShadow dx="0" dy="4" stdDeviation="4" floodColor="#5C514B" floodOpacity=".16" />
          </filter>
        </defs>
        <rect width="1680" height="1050" fill="#fff" />
        <rect width="1680" height="96" fill="#FAE7DA" />
        <text x="70" y="43" fontSize="29" fontWeight="700" fill="#9A5748">本周国债、政金债发行一二级利差散点图（{metricLabel === "综收－二级" ? "综收口径" : "中标口径"}）</text>
        <text x="70" y="76" fontSize="19" fill="#6B6662">交易日：{startDate} 至 {endDate}｜利差口径：{metricLabel}（bp）｜散点口径：单券</text>
        <svg x="1238" y="14" width="370" height="70" viewBox="970 350 1950 370" preserveAspectRatio="xMidYMid meet" aria-label="东方证券 FICC">
          <image data-chart-logo="true" href={`${ASSET_BASE}orient-ficc-logo.png`} width="3878" height="1071" />
        </svg>
        <text x="70" y="140" fontSize="22" fill="#68717b">普通债券样本 <tspan fontSize="26" fontWeight="700" fill="#D28A4D">{normalized.length}只</tspan> | 正利差债券 <tspan fontSize="26" fontWeight="700" fill="#D28A4D">{normalized.filter(r => Number(r.spread) > 0).length}只</tspan> | 平均利差 <tspan fontSize="26" fontWeight="700" fill="#D28A4D">{avg.toFixed(2)}bp</tspan> | 已排除特殊债券 <tspan fontSize="26" fontWeight="700" fill="#D28A4D">{excluded.length}只</tspan></text>
        <rect x={chartLeft} y={chartTop} width={chartRight-chartLeft} height={Math.max(0, y(0)-chartTop)} fill="#FFFCFA" />
        {ticks.map((tick) => (
          <g key={tick}>
            <line x1={chartLeft} y1={y(tick)} x2={chartRight} y2={y(tick)} stroke={tick === 0 ? "#EAB6A8" : "#E7E6E6"} strokeWidth={tick === 0 ? 2.3 : 1} strokeDasharray={tick === 0 ? "9 6" : undefined} />
            <text x={chartLeft-15} y={y(tick)+7} textAnchor="end" fontSize="20" fontWeight="600" fill="#7F7F7F">{tick > 0 ? "+" : ""}{Number(tick.toFixed(2))}</text>
          </g>
        ))}
        <line x1={chartLeft} y1={y(avg)} x2={chartRight} y2={y(avg)} stroke="#7F7F7F" strokeWidth="1.5" strokeDasharray="5 5" />
        <text x={chartRight-8} y={y(avg)-8} textAnchor="end" fontSize="17" fill="#7F7F7F">区间均值 {avg.toFixed(2)}bp</text>
        {maturities.map((item, index) => <g key={item.key}><line x1={chartLeft+index*categoryWidth} y1={chartTop} x2={chartLeft+index*categoryWidth} y2={chartBottom} stroke="#F2F2F2"/><text x={xCenter(item.key)} y="906" textAnchor="middle" fontSize="20" fontWeight="600" fill="#7F7F7F">{item.label}</text></g>)}
        <line x1={chartRight} y1={chartTop} x2={chartRight} y2={chartBottom} stroke="#F2F2F2"/>
        {points.map((row, index) => {
          const notable = Number(row.spread) > 0 || Number(row.spread) <= -1;
          return <g key={`${row.bondCode}-${index}`}>
            <circle cx={row.cx} cy={row.cy} r={notable ? 9 : 6.5} fill={colors[tone(row.bondType || "")]} fillOpacity={notable ? 1 : .76} stroke={notable ? darkColors[tone(row.bondType || "")] : "#FFFFFF"} strokeWidth={notable ? 2.8 : 1.2} />
          </g>;
        })}
        {callouts.map(({ row, label, box }, index) => <g key={`callout-${row.bondCode}-${index}`}>
          <path d={`M ${row.cx} ${row.cy} Q ${(row.cx + box.x + box.w / 2) / 2} ${row.cy - 10} ${box.x + box.w / 2} ${box.y + box.h / 2}`} fill="none" stroke={darkColors[tone(row.bondType || "")]} strokeWidth="1.5" opacity=".7"/>
          <rect x={box.x} y={box.y} width={box.w} height={box.h} rx="16" fill="#FFFFFF" stroke={darkColors[tone(row.bondType || "")]} strokeWidth="1.2" filter="url(#bubble-shadow)"/>
          <circle cx={box.x+16} cy={box.y+box.h/2} r="5" fill={colors[tone(row.bondType || "")]} stroke={darkColors[tone(row.bondType || "")]} strokeWidth="1.3"/>
          <text x={box.x+29} y={box.y+24} fontSize="17" fontWeight="700" fill="#3F3F3F">{label}</text>
        </g>)}
        <text x="860" y="950" textAnchor="middle" fontSize="21" fontWeight="600" fill="#626a73">发行期限</text>
        <text x="25" y="535" transform="rotate(-90 25 535)" textAnchor="middle" fontSize="21" fontWeight="600" fill="#626a73">{metricLabel}（bp）</text>
        {[["国债","#F4CDC3"],["国开债","#FAE1CC"],["口行债","#E5E5E3"],["农发债","#D6EFF5"]].map(([label,color],i) => <g key={label} transform={`translate(${485+i*190},1005)`}><circle r="11" fill={color}/><text x="22" y="7" fontSize="24" fontWeight="700" fill="#555b63">{label}</text></g>)}
      </svg>
      {!normalized.length && <div className="chart-empty">上传一二级表后，这里按所选区间生成利差图</div>}
    </div>
  );
}

function BrandLogo() {
  return <div className="brand-logo">
    <svg className="brand-logo-full" viewBox="970 350 1950 370" role="img" aria-label="东方证券 FICC" preserveAspectRatio="xMidYMid meet">
      <image href={`${ASSET_BASE}orient-ficc-logo.png`} width="3878" height="1071" />
    </svg>
  </div>;
}

export default function Workbench() {
  const [weekStart, setWeekStart] = useState(() => mondayOf(new Date()));
  const [active, setActive] = useState<TabKey>("overview");
  const [data, setData] = useState<WeekData>({ imports: [], records: [] });
  const [loading, setLoading] = useState(true);
  const [message, setMessage] = useState("");
  const [dragging, setDragging] = useState<DatasetType | null>(null);
  const [homeDragActive, setHomeDragActive] = useState(false);
  const [spreadMetric, setSpreadMetric] = useState<SpreadMetric>("all_in");
  const [analysisStart, setAnalysisStart] = useState(() => mondayOf(new Date()));
  const [analysisEnd, setAnalysisEnd] = useState(() => fridayOf(mondayOf(new Date())));
  const [chartRange, setChartRange] = useState(() => ({ start: mondayOf(new Date()), end: fridayOf(mondayOf(new Date())) }));
  const [chartRecords, setChartRecords] = useState<ParsedBondRecord[]>([]);
  const [chartAnalysisRecords, setChartAnalysisRecords] = useState<ParsedBondRecord[]>([]);
  const [spreadSourceRecords, setSpreadSourceRecords] = useState<ParsedBondRecord[]>([]);
  const [historicalSpreadRecords, setHistoricalSpreadRecords] = useState<ParsedBondRecord[]>([]);
  const [commentHistoryRecords, setCommentHistoryRecords] = useState<ParsedBondRecord[]>([]);
  const [commentPlanRecords, setCommentPlanRecords] = useState<ParsedBondRecord[]>([]);
  const [commentDrafts, setCommentDrafts] = useState<PolicyCommentDraft[]>([]);
  const [commentPlanFileName, setCommentPlanFileName] = useState("");
  const [commentSelectedDate, setCommentSelectedDate] = useState("");
  const [selectedCommentIds, setSelectedCommentIds] = useState<string[]>([]);
  const [commentDragTarget, setCommentDragTarget] = useState<"plan" | null>(null);
  const [commentLoading, setCommentLoading] = useState(false);
  const [chartLoading, setChartLoading] = useState(false);
  const [reportLoading, setReportLoading] = useState(false);
  const [reportStatus, setReportStatus] = useState("");
  const [confirmedNoLocalWeek, setConfirmedNoLocalWeek] = useState("");
  const [reportDownload, setReportDownload] = useState<{ url: string; name: string } | null>(null);
  const [latestDates, setLatestDates] = useState<LatestDates>({});
  const spreadInput = useRef<HTMLInputElement>(null);
  const localInput = useRef<HTMLInputElement>(null);
  const smartInput = useRef<HTMLInputElement>(null);
  const maturityInput = useRef<HTMLInputElement>(null);
  const issuancePlanInput = useRef<HTMLInputElement>(null);
  const commentPlanInput = useRef<HTMLInputElement>(null);
  const svgRef = useRef<SVGSVGElement>(null);
  const uploadInProgress = useRef(false);

  useEffect(() => () => {
    if (reportDownload) URL.revokeObjectURL(reportDownload.url);
  }, [reportDownload]);

  const records = useMemo(() => data.records.map(normalize), [data.records]);
  const localRecords = records.filter((_, i) => data.records[i]?.dataset_type === "local_bond");
  const spreadRecords = records.filter((_, i) => data.records[i]?.dataset_type === "spread");
  const maturityRecords = records.filter((_, i) => data.records[i]?.dataset_type === "maturity");
  const issuancePlanRecords = records.filter((_, i) => data.records[i]?.dataset_type === "issuance_plan");
  const reportWarnings = reportDataWarnings(spreadRecords, issuancePlanRecords);
  const weekEnd = fridayOf(weekStart);
  const spreadAmount = spreadRecords.reduce((sum, row) => sum + (row.amount || 0), 0);
  const latestSpreadDate = latestDates.spread || "暂无数据";
  const latestLocalDate = latestDates.local_bond || "暂无数据";
  const latestMaturityDate = latestDates.maturity || "暂无数据";
  const latestIssuancePlanDate = latestDates.issuance_plan || "暂无数据";
  const rateMaturityRecords = maturityRecords.filter(row => maturityKind(row) === "rate");
  const localMaturityRecords = maturityRecords.filter(row => maturityKind(row) === "local");
  const chartMetricRecords = useMemo(() => recordsForSpreadMetric(chartRecords, spreadMetric), [chartRecords, spreadMetric]);
  const chartMetricAnalysisRecords = useMemo(() => recordsForSpreadMetric(chartAnalysisRecords, spreadMetric), [chartAnalysisRecords, spreadMetric]);
  const spreadMetricLabel = spreadMetric === "all_in" ? "综收－二级" : "中标－二级";
  const chartAudit = useMemo(() => {
    const supported = chartMetricRecords.filter(row => ["国债", "国开债", "口行债", "农发债"].includes(row.bondType || ""));
    return {
      plotted: supported.filter(row => row.spread !== null && row.spread !== undefined && !isSpecialSpreadBond(row)).length,
      excluded: supported.filter(isSpecialSpreadBond).length,
      derived: supported.filter(row => ["derived_treasury", "recalculated"].includes(spreadAuditStatus(row))).length,
      missing: supported.filter(row => (row.spread === null || row.spread === undefined) && !isSpecialSpreadBond(row)).length,
    };
  }, [chartMetricRecords]);
  const rollingAnalysis = useMemo(
    () => rollingSpreadAnalysis(chartMetricRecords, chartMetricAnalysisRecords, chartRange.start, chartRange.end),
    [chartMetricRecords, chartMetricAnalysisRecords, chartRange],
  );
  const legacySpreadData = spreadRecords.length > 0 && spreadRecords.some((row) => !row.summaryMeta);
  const commentHistorySource = useMemo(() => {
    const merged = new Map<string, ParsedBondRecord>();
    [...historicalSpreadRecords, ...commentHistoryRecords].forEach((row) => merged.set(recordKey(row), row));
    return [...merged.values()].sort((a, b) => `${a.tradeDate}|${a.bondCode || ""}`.localeCompare(`${b.tradeDate}|${b.bondCode || ""}`));
  }, [historicalSpreadRecords, commentHistoryRecords]);
  const commentHistoryLatestDate = commentHistorySource.map((row) => row.tradeDate).sort().at(-1) || latestDates.spread || "暂无数据";
  const commentResults = useMemo(() => policyDraftResults(commentDrafts, commentHistorySource), [commentDrafts, commentHistorySource]);
  const commentDates = useMemo(() => [...new Set(commentDrafts.map((draft) => draft.tradeDate))].sort(), [commentDrafts]);
  const effectiveCommentDate = commentDates.includes(commentSelectedDate) ? commentSelectedDate : commentDates.at(-1) || "";
  const visibleCommentResults = commentResults.filter((item) => item.draft.tradeDate === effectiveCommentDate);
  const completedComments = visibleCommentResults.filter((item) => item.comment);
  const selectedCompletedComments = completedComments.filter((item) => selectedCommentIds.includes(item.draft.id));
  const allCompletedSelected = completedComments.length > 0 && selectedCompletedComments.length === completedComments.length;

  async function loadWeek() {
    setLoading(true);
    try {
      const historyStart = shiftDate(weekStart, -370);
      const historyEnd = shiftDate(weekStart, -1);
      const [response, latestResponse, historyResponse] = await Promise.all([
        workbenchRequest(`/api/workbench?weekStart=${weekStart}`),
        workbenchRequest("/api/workbench?meta=latest"),
        workbenchRequest(`/api/workbench?startDate=${historyStart}&endDate=${historyEnd}`),
      ]);
      const payload = await response.json() as WeekData & { error?: string };
      const latestPayload = await latestResponse.json() as { latestDates?: LatestDates; error?: string };
      const historyPayload = await historyResponse.json() as { records?: StoredRecord[]; error?: string };
      if (!response.ok) throw new Error(payload.error || "读取失败");
      if (!latestResponse.ok) throw new Error(latestPayload.error || "读取最新日期失败");
      if (!historyResponse.ok) throw new Error(historyPayload.error || "读取历史一二级数据失败");
      setData(payload);
      setLatestDates(latestPayload.latestDates || {});
      const loadedSpread = payload.records.filter(r => r.dataset_type === "spread").map(normalize);
      const historicalSpread = (historyPayload.records || []).filter((row) => row.dataset_type === "spread").map(normalize);
      setHistoricalSpreadRecords(historicalSpread);
      setAnalysisStart(weekStart);
      setAnalysisEnd(fridayOf(weekStart));
      setChartRange({ start: weekStart, end: fridayOf(weekStart) });
      setChartRecords(loadedSpread);
      setChartAnalysisRecords([...historicalSpread, ...loadedSpread]);
      setMessage("");
    } catch (error) {
      setMessage(error instanceof Error ? error.message : "读取失败");
    } finally { setLoading(false); }
  }

  useEffect(() => {
    const frame = window.requestAnimationFrame(() => { void loadWeek(); });
    return () => window.cancelAnimationFrame(frame);
    // loadWeek intentionally follows the selected week only.
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [weekStart]);

  async function upload(file: File, type: DatasetType) {
    if (uploadInProgress.current) {
      setMessage("上一份文件仍在保存，请完成后再上传下一份。");
      return;
    }
    uploadInProgress.current = true;
    setMessage(`正在解析 ${file.name}…`);
    try {
      const parsed = type === "local_bond" ? await parseLocalBondFile(file)
        : type === "spread" ? await parseSpreadFile(file)
          : type === "maturity" ? await parseMaturityFile(file) : await parseIssuancePlanFile(file);
      if (type === "spread") {
        setSpreadSourceRecords(parsed);
        setCommentHistoryRecords(parsed);
        if (commentPlanRecords.length) setCommentDrafts(createPolicyCommentDrafts(commentPlanRecords, [...historicalSpreadRecords, ...parsed], commentDrafts));
      }
      const recordWeekOf = (row: ParsedBondRecord) => type === "maturity" ? maturityWeekStart(row.tradeDate) : mondayOf(row.tradeDate);
      const detectedWeeks = new Set(parsed.map(recordWeekOf));
      const isHistoricalBase = detectedWeeks.size > 1;
      const groups = new Map<string, ParsedBondRecord[]>();
      if (isHistoricalBase || type === "local_bond") {
        parsed.forEach((row) => {
          const recordWeek = recordWeekOf(row);
          groups.set(recordWeek, [...(groups.get(recordWeek) || []), row]);
        });
      } else {
        const inWeek = parsed.filter((row) => row.tradeDate >= weekStart && row.tradeDate <= weekEnd);
        if (!inWeek.length) throw new Error(`文件中没有 ${displayWeek(weekStart)} 的记录`);
        groups.set(weekStart, inWeek);
      }
      if (isHistoricalBase) setMessage(`已识别 ${groups.size} 个交易周，正在保存到${LOCAL_STORAGE_MODE ? "当前浏览器" : "共享数据库"}…`);
      let added = 0;
      let updated = 0;
      let unchanged = 0;
      const entries = [...groups.entries()];
      let cursor = 0;
      async function saveNext() {
        while (cursor < entries.length) {
          const [recordWeek, group] = entries[cursor++];
          const tradeDate = group.map((r) => r.tradeDate).sort().at(-1)!;
          const response = await workbenchRequest("/api/workbench", {
            method: "POST", headers: { "Content-Type": "application/json" },
            body: JSON.stringify({ datasetType: type, tradeDate, weekStart: recordWeek, fileName: file.name, records: group }),
          });
          const payload = await response.json() as { error?: string; inserted?: number; updated?: number; unchanged?: number };
          if (!response.ok) throw new Error(payload.error || "保存失败");
          added += payload.inserted || 0;
          updated += payload.updated || 0;
          unchanged += payload.unchanged || 0;
        }
      }
      const saved = await Promise.allSettled(Array.from({ length: Math.min(6, entries.length) }, () => saveNext()));
      const failures = saved.filter((result): result is PromiseRejectedResult => result.status === "rejected");
      if (failures.length) {
        await loadWeek();
        throw new Error(`部分记录未能保存（已新增${added}条、更新${updated}条）。${failures.map(result => String(result.reason instanceof Error ? result.reason.message : result.reason)).join("；")}。请重新上传，已成功记录不会重复计入。`);
      }
      const result = `新增${added}条，更新${updated}条，保留${unchanged}条未变化记录`;
      await loadWeek();
      const sourceDates = parsed.map(row => row.tradeDate).sort();
      const sourceTotal = parsed.reduce((sum, row) => sum + (row.amount || 0), 0);
      setMessage(`已识别 ${parsed.length} 条，合计 ${Number(sourceTotal.toFixed(4))} 亿元，覆盖 ${sourceDates[0]} 至 ${sourceDates.at(-1)}。已保存 ${groups.size} 个交易周：${result}${type === "local_bond" ? "。地方债明细已接入周报，续发券一并计入" : ""}`);
    } catch (error) { setMessage(error instanceof Error ? error.message : "上传失败"); }
    finally { uploadInProgress.current = false; }
  }

  function dropFile(event: React.DragEvent<HTMLButtonElement>, type: DatasetType) {
    event.preventDefault();
    setDragging(null);
    const file = event.dataTransfer.files?.[0];
    if (file) void upload(file, type);
  }

  function uploadDetectedFile(file: File) {
    if (!/\.(?:xlsx|xlsm)$/i.test(file.name)) {
      setMessage("请上传 .xlsx 或 .xlsm 文件");
      return;
    }
    const name = file.name;
    const detected: DatasetType | null = /到期|偿还/.test(name) ? "maturity"
      : /地方政府债|地方债/.test(name) ? "local_bond"
      : /新债发行|发行计划|招标计划/.test(name) ? "issuance_plan"
        : /一二级|利差|短评/.test(name) ? "spread" : null;
    if (!detected) {
      setMessage("未能从文件名识别用途，请拖到对应的上传卡片");
      return;
    }
    void upload(file, detected);
  }

  function dropHomepageFile(event: React.DragEvent<HTMLElement>) {
    event.preventDefault();
    setHomeDragActive(false);
    const file = event.dataTransfer.files?.[0];
    if (file) uploadDetectedFile(file);
  }

  async function generateSpreadRange() {
    if (!analysisStart || !analysisEnd || analysisStart > analysisEnd) {
      setMessage("请选择有效的利差分析日期区间");
      return;
    }
    setChartLoading(true);
    try {
      const benchmarkStart = shiftDate(analysisStart, -28);
      const response = await workbenchRequest(`/api/workbench?startDate=${benchmarkStart}&endDate=${analysisEnd}`);
      const payload = await response.json() as { records?: StoredRecord[]; error?: string };
      if (!response.ok) throw new Error(payload.error || "读取区间数据失败");
      const stored = (payload.records || []).map(normalize);
      const fromCurrentFile = spreadSourceRecords.filter(row => row.tradeDate >= benchmarkStart && row.tradeDate <= analysisEnd);
      const merged = new Map<string, ParsedBondRecord>();
      [...stored, ...fromCurrentFile].forEach(row => merged.set(`${row.tradeDate}|${row.bondCode || row.shortName || ""}`, row));
      const analysisRows = [...merged.values()].sort((a, b) => `${a.tradeDate}${a.bondCode}`.localeCompare(`${b.tradeDate}${b.bondCode}`));
      const selected = analysisRows.filter(row => row.tradeDate >= analysisStart && row.tradeDate <= analysisEnd);
      setChartRecords(selected);
      setChartAnalysisRecords(analysisRows);
      setChartRange({ start: analysisStart, end: analysisEnd });
      setMessage(selected.length ? `已生成 ${analysisStart} 至 ${analysisEnd} 的利差图，共 ${selected.length} 条记录` : "所选区间暂无已入库的一二级数据");
    } catch (error) {
      setMessage(error instanceof Error ? error.message : "生成区间利差图失败");
    } finally { setChartLoading(false); }
  }

  async function loadCommentPlanFile(file: File) {
    setCommentLoading(true);
    setMessage(`正在转换 ${file.name} 为今日待填写区…`);
    try {
      const parsed = await parseIssuancePlanFile(file);
      const drafts = createPolicyCommentDrafts(parsed, commentHistorySource, commentDrafts);
      if (!drafts.length) throw new Error("文件中没有识别到国开、口行或农发债");
      setCommentPlanRecords(parsed);
      setCommentDrafts(drafts);
      setSelectedCommentIds([]);
      setCommentSelectedDate([...new Set(drafts.map((draft) => draft.tradeDate))].sort().at(-1) || "");
      setCommentPlanFileName(file.name);
      setMessage(`今日待填写区已生成：${drafts.length} 只政金债，可填写二级价格和最终中标率`);
    } catch (error) {
      setMessage(error instanceof Error ? error.message : "新债发行文件转换失败");
    } finally {
      setCommentLoading(false);
    }
  }

  function dropCommentFile(event: React.DragEvent<HTMLElement>) {
    event.preventDefault();
    event.stopPropagation();
    setCommentDragTarget(null);
    const file = event.dataTransfer.files?.[0];
    if (!file) return;
    if (!/\.(?:xlsx|xlsm)$/i.test(file.name)) {
      setMessage("请拖入 .xlsx 或 .xlsm 文件");
      return;
    }
    void loadCommentPlanFile(file);
  }

  function toggleCommentSelection(id: string) {
    setSelectedCommentIds((current) => current.includes(id) ? current.filter((item) => item !== id) : [...current, id]);
  }

  function toggleAllCompletedComments() {
    const completedIds = completedComments.map((item) => item.draft.id);
    setSelectedCommentIds((current) => allCompletedSelected
      ? current.filter((id) => !completedIds.includes(id))
      : [...new Set([...current, ...completedIds])]);
  }

  function updateCommentDraft(id: string, field: "rateType" | "benchmarkType" | "referenceBond" | "benchmarkValue" | "finalValue", value: string) {
    setCommentDrafts((current) => current.map((draft) => {
      if (draft.id !== id) return draft;
      if (field !== "rateType") return { ...draft, [field]: value };
      const reopened = /[XZ]\d*$/i.test(draft.bondCode);
      const wasDr = reopened && /^DR(?:001|007)?浮息债$/i.test(draft.rateType);
      const nextDr = reopened && /^DR(?:001|007)?浮息债$/i.test(value);
      return {
        ...draft,
        rateType: value,
        ...(wasDr !== nextDr ? { benchmarkType: nextDr ? "估价" : "二级", benchmarkValue: "", finalValue: "" } : {}),
      };
    }));
  }

  async function copyComment(text: string, label: string) {
    try {
      await navigator.clipboard.writeText(text);
      setMessage(`已复制 ${label} 的发行小结`);
    } catch {
      setMessage("复制失败，请手动选中文字复制");
    }
  }

  async function copySelectedComments() {
    if (!selectedCompletedComments.length) return;
    await copyComment(selectedCompletedComments.map((row) => row.comment!.text).join("\n\n"), `${effectiveCommentDate}已选${selectedCompletedComments.length}只政金债`);
  }

  async function copyRollingAnalysis() {
    try {
      await navigator.clipboard.writeText(rollingAnalysis.text);
      setMessage("已复制四周滚动分析文字");
    } catch {
      setMessage("复制失败，请手动选中文字复制");
    }
  }

  async function deleteImport(importId: string) {
    const response = await workbenchRequest(`/api/workbench?importId=${encodeURIComponent(importId)}`, { method: "DELETE" });
    if (response.ok) await loadWeek();
  }

  async function downloadChart() {
    if (!svgRef.current) return;
    const chart = svgRef.current.cloneNode(true) as SVGSVGElement;
    const chartLogo = chart.querySelector("image[data-chart-logo]");
    if (chartLogo) {
      try {
        const logoBlob = await fetch(`${ASSET_BASE}orient-ficc-logo.png`).then((response) => response.blob());
        const logoDataUrl = await new Promise<string>((resolve, reject) => {
          const reader = new FileReader();
          reader.onload = () => resolve(String(reader.result || ""));
          reader.onerror = () => reject(reader.error);
          reader.readAsDataURL(logoBlob);
        });
        chartLogo.setAttribute("href", logoDataUrl);
      } catch {
        // The chart remains downloadable even if the logo asset cannot be embedded.
      }
    }
    const xml = new XMLSerializer().serializeToString(chart);
    const blob = new Blob([xml], { type: "image/svg+xml;charset=utf-8" });
    const url = URL.createObjectURL(blob);
    const image = new Image();
    image.onload = () => {
      const canvas = document.createElement("canvas"); canvas.width = 1680; canvas.height = 1050;
      canvas.getContext("2d")?.drawImage(image, 0, 0, 1680, 1050);
      canvas.toBlob((png) => {
        if (!png) return;
        const a = document.createElement("a"); a.href = URL.createObjectURL(png);
        a.download = `国债政金债${spreadMetricLabel}_${chartRange.start}_${chartRange.end}.png`; a.click();
        URL.revokeObjectURL(a.href); URL.revokeObjectURL(url);
      }, "image/png");
    };
    image.src = url;
  }

  async function exportDocx() {
    if (!localRecords.length && confirmedNoLocalWeek !== weekStart) {
      setReportStatus("请先上传本周地方债发行明细；若确无发行，请勾选确认后再生成，避免将缺失数据计为零。");
      return;
    }
    setReportLoading(true);
    setReportStatus("正在按今日 Word 母版填充数据…");
    try {
      const mmdd = (d: string) => d.slice(5).replace("-", "");
      const previousStart = shiftWeek(weekStart, -1);
      const yearStart = `${weekStart.slice(0, 4)}-01-01`;
      const [previousResponse, ytdLocalResponse] = await Promise.all([
        workbenchRequest(`/api/workbench?weekStart=${previousStart}`),
        workbenchRequest(`/api/workbench?startDate=${yearStart}&endDate=${weekEnd}&datasetType=local_bond`),
      ]);
      const previousPayload = await previousResponse.json() as WeekData & { error?: string };
      const ytdLocalPayload = await ytdLocalResponse.json() as { records?: StoredRecord[]; error?: string };
      if (!previousResponse.ok) throw new Error(previousPayload.error || "读取上周数据失败");
      if (!ytdLocalResponse.ok) throw new Error(ytdLocalPayload.error || "读取地方债年度数据失败");
      const previousSpreadRecords = previousPayload.records.filter((row) => row.dataset_type === "spread").map(normalize);
      const previousMaturityRecords = previousPayload.records.filter((row) => row.dataset_type === "maturity").map(normalize);
      const ytdLocalRecords = (ytdLocalPayload.records || []).map(normalize);
      const rateTotal = rateMaturityRecords.reduce((sum, row) => sum + (row.amount || 0), 0);
      const localDaily = maturityDailyTotals(localMaturityRecords, weekStart);
      const previousRateMaturity = previousMaturityRecords.filter(row => maturityKind(row) === "rate").reduce((sum, row) => sum + (row.amount || 0), 0);
      const previousRateIssuance = previousSpreadRecords.reduce((sum, row) => sum + (row.amount || 0), 0);
      const maturity = maturityRecords.length ? {
        rateTotal,
        rateBreakdown: rateMaturityBreakdown(rateMaturityRecords),
        localDaily,
        localTotal: localMaturityRecords.reduce((sum, row) => sum + (row.amount || 0), 0),
        previousRateNet: previousMaturityRecords.length ? previousRateIssuance - previousRateMaturity : undefined,
      } : undefined;
      const blob = await buildWeeklyReportBlob({ weekStart, summary: rollingAnalysis.text, localRecords, spreadRecords, scheduleRecords: issuancePlanRecords, previousSpreadRecords, ytdLocalRecords, maturity });
      const url = URL.createObjectURL(blob);
      const fileName = `利率债发行周报${weekStart.replaceAll("-", "")}-${mmdd(weekEnd)}.docx`;
      setReportDownload({ url, name: fileName });
      const a = document.createElement("a"); a.href = url;
      a.download = fileName;
      document.body.appendChild(a);
      a.click();
      a.remove();
      setReportStatus(legacySpreadData ? "已生成；当前为旧版入库数据，空缺字段以“-”显示，重新上传一二级表可补全。" : maturityRecords.length ? "周报已按今日母版生成，并已写入本周到期数据。若未自动下载，请使用下方下载链接。" : "周报已生成，但本周尚未上传到期数据；到期及净融资项目未填充。" );
    } catch (error) {
      const reason = error instanceof Error ? error.message : "生成失败";
      setReportStatus(`生成失败：${reason}`);
    } finally {
      setReportLoading(false);
    }
  }

  return (
    <main className="app-shell">
      <aside className="sidebar">
        <div className="brand"><BrandLogo/><div className="brand-copy"><strong>利率债发行工作台</strong><small>Issuance Desk</small></div></div>
        <nav>{tabGroups.map(group => <div className="nav-group" key={group.label}><span className="nav-group-label">{group.label}</span>{group.items.map(([key, label]) => <button key={key} className={active === key ? "nav-active" : ""} onClick={() => setActive(key)}>{key === "overview" ? <House/> : key === "spread" ? <BarChart3/> : <FileText/>}<span>{label}</span></button>)}</div>)}</nav>
        <div className="side-note"><Check size={16}/><span>{LOCAL_STORAGE_MODE ? <>数据保存在当前浏览器<br/>换设备需重新上传</> : <>公开共享数据<br/>周度记录集中留存</>}</span></div>
      </aside>

      <section className="workspace">
        <header className="topbar">
          <div><p className="eyebrow">{viewCopy[active].eyebrow}</p><h1>{viewCopy[active].title}</h1><p>{viewCopy[active].description}{active === "local" || active === "comment" ? "" : ` · ${displayWeek(weekStart)}`}</p></div>
          {active !== "local" && active !== "comment" && <div className="week-picker"><button aria-label="上一周" onClick={() => setWeekStart(shiftWeek(weekStart,-1))}><ChevronLeft/></button><input type="date" value={weekStart} onChange={e => setWeekStart(mondayOf(e.target.value))}/><button aria-label="下一周" onClick={() => setWeekStart(shiftWeek(weekStart,1))}><ChevronRight/></button></div>}
        </header>

        {message && <div className={`notice ${/失败|缺少|没有/.test(message) ? "notice-error" : ""}`}><CircleAlert size={17}/>{message}<button onClick={() => setMessage("")}><X size={15}/></button></div>}

        {active === "overview" && <button className={`home-smart-drop ${homeDragActive ? "drag-active" : ""}`}
          onClick={() => smartInput.current?.click()}
          onDragEnter={event => { event.preventDefault(); setHomeDragActive(true); }}
          onDragOver={event => { event.preventDefault(); event.dataTransfer.dropEffect = "copy"; setHomeDragActive(true); }}
          onDragLeave={() => setHomeDragActive(false)}
          onDrop={dropHomepageFile}>
          <Upload/><span><strong>拖入 Excel，自动识别文件用途</strong><small>支持地方债发行明细、一二级利差、新债发行计划和到期明细</small></span><em>智能上传</em>
          <input ref={smartInput} hidden type="file" accept=".xlsx,.xlsm" onChange={event => { const file = event.target.files?.[0]; if (file) uploadDetectedFile(file); event.currentTarget.value = ""; }}/>
        </button>}

        {active === "overview" && <section className="hero-grid">
          <div className="upload-card">
            <div className="card-head"><div><span className="section-label">功能入口</span><h2>选择工具或直接上传文件</h2></div><Upload/></div>
            <div className="upload-actions">
              <div className="upload-lane">
                <button
                  className={`drop-zone local-drop ${dragging === "local_bond" ? "drag-active" : ""}`}
                  onClick={() => localInput.current?.click()}
                  onDragEnter={event => { event.preventDefault(); setDragging("local_bond"); }}
                  onDragOver={event => { event.preventDefault(); event.dataTransfer.dropEffect = "copy"; }}
                  onDragLeave={() => setDragging(null)}
                  onDrop={event => dropFile(event, "local_bond")}
                ><FileSpreadsheet/><span><em>周报取数 · 增量更新</em><strong>地方债发行明细</strong><small>拖入或点击上传 Excel<br/>自动去重，保留历史并更新修订值</small></span></button>
                <div className="latest-date"><CalendarDays/><span>地方债明细最新日期</span><strong>{latestLocalDate}</strong></div>
              </div>
              <div className="upload-lane">
                <button
                  className={`drop-zone rate-drop ${dragging === "spread" ? "drag-active" : ""}`}
                  onClick={() => spreadInput.current?.click()}
                  onDragEnter={event => { event.preventDefault(); setDragging("spread"); }}
                  onDragOver={event => { event.preventDefault(); event.dataTransfer.dropEffect = "copy"; }}
                  onDragLeave={() => setDragging(null)}
                  onDrop={event => dropFile(event, "spread")}
                ><BarChart3/><span><em>利率债板块</em><strong>国债及政金债一二级表</strong><small>拖入或点击选择 Excel<br/>生成利差图与四周滚动分析</small></span></button>
                <div className="latest-date spread-date"><CalendarDays/><span>一二级利差最新日期</span><strong>{latestSpreadDate}</strong></div>
              </div>
            </div>
            <input ref={spreadInput} hidden type="file" accept=".xlsx,.xlsm" onChange={e => { const file=e.target.files?.[0]; if(file) void upload(file,"spread"); e.currentTarget.value=""; }}/>
            <input ref={localInput} hidden type="file" accept=".xlsx,.xlsm" onChange={e => { const file=e.target.files?.[0]; if(file) void upload(file,"local_bond"); e.currentTarget.value=""; }}/>
            <p>地方债明细用于周报发行统计；按发行日期和债券代码识别增量，重复上传不重复计量。一二级表提供国债政金债数据。</p>
            <div className="converter-shortcut"><button className="secondary" onClick={() => setActive("local")}><FileSpreadsheet/>地方债日表转换</button><span>独立转换工具 · 不自动计入周报</span></div>
          </div>
          <div className="report-source-card">
            <div className="card-head"><div><span className="section-label">周报生成数据</span><h2>发行时段与到期明细</h2><p>发行Excel只补充上午/下午；到期Excel提供全部到期量、到期结构和净融资口径。</p></div><FileText/></div>
            <div className="upload-actions report-upload-actions">
              <div className="upload-lane">
                <button
                  className={`drop-zone plan-drop ${dragging === "issuance_plan" ? "drag-active" : ""}`}
                  onClick={() => issuancePlanInput.current?.click()}
                  onDragEnter={event => { event.preventDefault(); setDragging("issuance_plan"); }}
                  onDragOver={event => { event.preventDefault(); event.dataTransfer.dropEffect = "copy"; }}
                  onDragLeave={() => setDragging(null)}
                  onDrop={event => dropFile(event, "issuance_plan")}
                ><CalendarDays/><span><em>发行时段补充</em><strong>新债发行计划</strong><small>只读取发行日期、债券代码和招标时间<br/>不采用文件中的发行量</small></span></button>
                <div className="latest-date plan-date"><CalendarDays/><span>发行时段文件最新日期</span><strong>{latestIssuancePlanDate}</strong></div>
              </div>
              <div className="upload-lane">
                <button
                  className={`drop-zone maturity-drop ${dragging === "maturity" ? "drag-active" : ""}`}
                  onClick={() => maturityInput.current?.click()}
                  onDragEnter={event => { event.preventDefault(); setDragging("maturity"); }}
                  onDragOver={event => { event.preventDefault(); event.dataTransfer.dropEffect = "copy"; }}
                  onDragLeave={() => setDragging(null)}
                  onDrop={event => dropFile(event, "maturity")}
                ><CalendarDays/><span><em>到期明细</em><strong>本周到期数据</strong><small>自动区分地方债与国债政金债<br/>细分贴现、附息和清发品种</small></span></button>
                <div className="latest-date maturity-date"><CalendarDays/><span>到期数据最新日期</span><strong>{latestMaturityDate}</strong></div>
              </div>
            </div>
            <input ref={maturityInput} hidden type="file" accept=".xlsx,.xlsm" onChange={e => { const file=e.target.files?.[0]; if(file) void upload(file,"maturity"); e.currentTarget.value=""; }}/>
            <input ref={issuancePlanInput} hidden type="file" accept=".xlsx,.xlsm" onChange={e => { const file=e.target.files?.[0]; if(file) void upload(file,"issuance_plan"); e.currentTarget.value=""; }}/>
          </div>
        </section>}

        {active === "overview" && <section className="overview-section">
          <div className="overview-heading"><span className="section-label">WEEKLY OVERVIEW</span><h2>本周发行与到期总览</h2></div>
          <div className="metrics">
            <article><span>地方债发行量</span><strong>{localRecords.reduce((sum,row)=>sum+(row.amount||0),0).toFixed(2)}</strong><small>亿元 · {localRecords.length}条明细</small></article>
            <article><span>国债政金债发行量</span><strong>{spreadAmount.toFixed(2)}</strong><small>亿元 · 一二级表口径</small></article>
            <article><span>国债及政金债到期</span><strong>{rateMaturityRecords.reduce((sum,row)=>sum+(row.amount||0),0).toFixed(2)}</strong><small>亿元 · {rateMaturityRecords.length}条明细</small></article>
            <article><span>地方债到期</span><strong>{localMaturityRecords.reduce((sum,row)=>sum+(row.amount||0),0).toFixed(2)}</strong><small>亿元 · {localMaturityRecords.length}条明细</small></article>
          </div>
        </section>}

        {active === "local" && <section className="converter-panel">
          <div className="converter-intro">
            <div><span className="section-label">LOCAL CONVERTER</span><h2>地方债日表转换器</h2><p>拖入 DM 下载的“地方政府债+日期.xlsx”，自动下载每日发行计划并生成可复制文本。文件仅在浏览器本地处理；周报取数请另从首页上传地方债发行明细。</p></div>
          </div>
          <iframe className="converter-frame" title="地方债日表转换器" src={`${ASSET_BASE}local-bond-daily-converter.html`} />
        </section>}

        {active === "spread" && <section className="panel focus-panel">
          <div className="panel-head spread-panel-head"><div><span className="section-label">PRIMARY / SECONDARY</span><h2>一二级利差可视化</h2></div><button className="secondary" onClick={downloadChart} disabled={!chartRecords.length}><Download/>下载 PNG</button></div>
          <div className="range-toolbar">
            <label><span>开始日期</span><input type="date" value={analysisStart} onChange={event => setAnalysisStart(event.target.value)}/></label>
            <span className="range-divider">至</span>
            <label><span>结束日期</span><input type="date" value={analysisEnd} onChange={event => setAnalysisEnd(event.target.value)}/></label>
            <div className="metric-toggle" aria-label="利差口径">
              <span>利差口径</span>
              <div><button className={spreadMetric === "all_in" ? "metric-active" : ""} onClick={() => setSpreadMetric("all_in")}>综收－二级</button><button className={spreadMetric === "winning" ? "metric-active" : ""} onClick={() => setSpreadMetric("winning")}>中标－二级</button></div>
            </div>
            <button className="range-generate" onClick={generateSpreadRange} disabled={chartLoading}>{chartLoading ? <LoaderCircle className="spin"/> : <BarChart3/>}生成图表</button>
            <p>当前区间：{chartRange.start} 至 {chartRange.end} · {chartRecords.length} 条记录</p>
          </div>
          <div className="chart-audit" aria-label="利差数据校验结果">
            <span>有效样本 <strong>{chartAudit.plotted}</strong></span>
            <span>按口径排除 <strong>{chartAudit.excluded}</strong></span>
            <span>自动复算 <strong>{chartAudit.derived}</strong></span>
            <span className={chartAudit.missing ? "audit-warning" : ""}>缺少可比值 <strong>{chartAudit.missing}</strong></span>
          </div>
          <SpreadChart records={chartMetricRecords} svgRef={svgRef} startDate={chartRange.start} endDate={chartRange.end} metricLabel={spreadMetricLabel}/>
          <section className="rolling-analysis" aria-label="四周滚动利差分析">
            <div className="rolling-analysis-head">
              <div><span className="section-label">FOUR-WEEK ROLLING</span><h3>同品种同期限四周滚动分析</h3><p>口径：{spreadMetricLabel} · 基准窗口：{rollingAnalysis.benchmarkStart} 至 {rollingAnalysis.benchmarkEnd} · 偏离阈值 ±0.60BP</p></div>
              <button className="secondary" onClick={() => void copyRollingAnalysis()} disabled={!rollingAnalysis.text}><Copy/>复制分析</button>
            </div>
            <div className="rolling-analysis-metrics">
              <span>可比组合 <strong>{rollingAnalysis.comparableGroups}</strong></span>
              <span>正常波动 <strong>{rollingAnalysis.normalGroups}</strong></span>
              <span className={rollingAnalysis.notableGroups ? "rolling-notable" : ""}>突出变化 <strong>{rollingAnalysis.notableGroups}</strong></span>
              <span>特殊债券 <strong>{rollingAnalysis.specialBonds}</strong></span>
            </div>
            <pre>{rollingAnalysis.text}</pre>
          </section>
        </section>}

        {active === "comment" && <section className="panel focus-panel comment-panel"
          onDragOver={event => { event.preventDefault(); event.dataTransfer.dropEffect = "copy"; }}
          onDrop={dropCommentFile}>
          <div className="panel-head comment-panel-head">
            <div><span className="section-label">DAILY INPUT</span><h2>今日政金债填写区</h2><p>先上传“新债发行”Excel，系统只提取国开、口行和农发债。你填写比较基准与最终中标率后，发行小结会即时生成。</p></div>
            <button className="secondary" onClick={() => commentPlanInput.current?.click()} disabled={commentLoading}>{commentLoading ? <LoaderCircle className="spin"/> : <Upload/>}{commentPlanFileName ? "更换发行文件" : "上传新债发行"}</button>
          </div>
          <input ref={commentPlanInput} hidden type="file" accept=".xlsx,.xlsm" onChange={event => { const file = event.target.files?.[0]; if (file) void loadCommentPlanFile(file); event.currentTarget.value = ""; }}/>
          <div className="comment-source-grid">
            <button className={`comment-source-card comment-plan-source ${commentDragTarget === "plan" ? "drag-active" : ""}`} onClick={() => commentPlanInput.current?.click()} disabled={commentLoading}
              onDragEnter={event => { event.preventDefault(); setCommentDragTarget("plan"); }}
              onDragOver={event => { event.preventDefault(); event.dataTransfer.dropEffect = "copy"; }}
              onDragLeave={() => setCommentDragTarget(null)}
              onDrop={dropCommentFile}>
              <FileSpreadsheet/><span><small>第一步 · 今日发行</small><strong>{commentPlanFileName || "上传新债发行 Excel"}</strong><em>{commentDrafts.length ? `已转换 ${commentDrafts.length} 只政金债` : "自动过滤地方债与信用债"}</em></span>
            </button>
            <div className="comment-source-card comment-history-note">
              <RefreshCw/><span><small>历史来源 · 首页统一管理</small><strong>一二级利差历史库</strong><em>最新日期 {commentHistoryLatestDate} · 当前可用 {commentHistorySource.length} 条记录</em></span>
            </div>
          </div>
          {!commentDrafts.length ? <button className={`comment-drop ${commentDragTarget === "plan" ? "drag-active" : ""}`} onClick={() => commentPlanInput.current?.click()} disabled={commentLoading}
            onDragEnter={event => { event.preventDefault(); setCommentDragTarget("plan"); }}
            onDragOver={event => { event.preventDefault(); event.dataTransfer.dropEffect = "copy"; }}
            onDragLeave={() => setCommentDragTarget(null)}
            onDrop={dropCommentFile}>
            {commentLoading ? <LoaderCircle className="spin"/> : <Upload/>}
            <span><strong>拖入或点击上传“新债发行-日期.xlsx”</strong><small>转换后将在下方显示债券代码、期限、比较基准、二级价格、最终中标率和自动发行小结。</small></span>
          </button> : <>
            <div className="comment-toolbar comment-entry-toolbar">
              <div><Check/><span><small>当日填写区</small><strong>{commentPlanFileName} · 全部 {commentDrafts.length} 只政金债</strong></span></div>
              <label><span>发行日期</span><select value={effectiveCommentDate} onChange={event => setCommentSelectedDate(event.target.value)}>{commentDates.map(date => <option key={date} value={date}>{date}</option>)}</select></label>
              <button className="range-generate" onClick={copySelectedComments} disabled={!selectedCompletedComments.length}><Copy/>复制已选{selectedCompletedComments.length ? `（${selectedCompletedComments.length}）` : ""}</button>
            </div>
            <div className="comment-audit"><Check/><span>代码末尾无 X/Z 的新券统一按收益率差计算，只写“今日新发”，不判断利差变化；增发 DR007/DR001 浮息债按净价差处理。</span></div>
            <div className="comment-entry-table-wrap"><table className="comment-entry-table">
              <thead><tr><th className="comment-select-column"><input type="checkbox" aria-label="全选已完成的小结" checked={allCompletedSelected} disabled={!completedComments.length} onChange={toggleAllCompletedComments}/></th><th>债券</th><th>发行类型</th><th>浮息类型</th><th>比较基准</th><th>参考券（选填）</th><th>二级收益率 / 净价</th><th>最终中标率 / 中标净价</th><th>自动文字详情</th></tr></thead>
              <tbody>{visibleCommentResults.map(({ draft, comment, missing }) => {
                const reopened = /[XZ]\d*$/i.test(draft.bondCode);
                const drPricing = reopened && /^DR(?:001|007)?浮息债$/i.test(draft.rateType);
                const benchmarkOptions = drPricing ? ["估价", "二级", "中间价"] : ["二级", "估值", "估值曲线", "中间价", "价格"];
                const selected = selectedCommentIds.includes(draft.id) && Boolean(comment);
                return <tr key={draft.id} className={selected ? "comment-row-selected" : ""}>
                  <td className="comment-select-column"><input type="checkbox" aria-label={`选择 ${draft.bondCode} 小结`} checked={selected} disabled={!comment} onChange={() => toggleCommentSelection(draft.id)}/></td>
                  <td><strong>{draft.tenor} {draft.bondCode}</strong><small>{draft.shortName}<br/>{draft.tradeDate}</small>{draft.sequenceCheck && <span className={`sequence-check sequence-check-${draft.sequenceCheck.status}`}>{draft.sequenceCheck.message}</span>}</td>
                  <td><span className={`comment-kind ${reopened ? "comment-kind-reopened" : "comment-kind-new"}`}>{reopened ? "增发券" : "新券"}</span></td>
                  <td><select aria-label={`${draft.bondCode} 浮息类型`} value={draft.rateType} onChange={event => updateCommentDraft(draft.id,"rateType",event.target.value)}>{POLICY_FLOAT_RATE_OPTIONS.map(option => <option key={option || "empty"} value={option}>{option || "空值"}</option>)}</select></td>
                  <td><select aria-label={`${draft.bondCode} 比较基准`} value={draft.benchmarkType} onChange={event => updateCommentDraft(draft.id,"benchmarkType",event.target.value)}>{benchmarkOptions.map(option => <option key={option} value={option}>{option}</option>)}</select></td>
                  <td><input aria-label={`${draft.bondCode} 参考券代码`} value={draft.referenceBond || ""} onChange={event => updateCommentDraft(draft.id,"referenceBond",event.target.value)} placeholder="留空为本券，如 260214"/></td>
                  <td><input aria-label={`${draft.bondCode} 二级收益率或净价`} inputMode="decimal" value={draft.benchmarkValue} onChange={event => updateCommentDraft(draft.id,"benchmarkValue",event.target.value)} placeholder={drPricing ? "如 99.9500" : "如 1.4730"}/></td>
                  <td><input aria-label={`${draft.bondCode} 最终中标率或中标净价`} inputMode="decimal" value={draft.finalValue} onChange={event => updateCommentDraft(draft.id,"finalValue",event.target.value)} placeholder={drPricing ? "中标净价" : "如 1.4381"}/></td>
                  <td className="comment-output-cell">{comment ? <><pre>{comment.text}</pre><div><span>{reopened ? `上次同券：${comment.previousCode || "未找到"}` : "新券：不做历史利差判断"}</span><button onClick={() => void copyComment(comment.text, comment.displayCode)}><Copy/>复制</button></div></> : <span className="comment-missing">待填写：{missing.join("、")}</span>}</td>
                </tr>;
              })}</tbody>
            </table></div>
          </>}
        </section>}

        {active === "report" && <section className="report-panel focus-panel">
          <div className="local-report-source"><p>国债政金债发行量、利差及结果以一二级表为准。发行计划用于核对和补充时段；上弹、追加或特殊小团等差异由你核定，修改一二级表后重新上传即可。净融资按发行量减到期明细金额测算。</p></div>
          {reportWarnings.length > 0 && <details className="local-report-source report-reconciliation" open><summary>一二级与发行计划核对 · {reportWarnings.length} 项</summary><p>以下差异由你判断，生成周报继续采用一二级数据。</p><ul>{reportWarnings.map(note => <li key={note}>{note}</li>)}</ul></details>}
          <div className="local-report-source">
            <p>地方债发行取数：本周 {localRecords.length} 条明细，合计 {localRecords.reduce((sum,row)=>sum+(row.amount||0),0).toFixed(2)} 亿元。最新明细日期：{latestLocalDate}。</p>
            {!localRecords.length && <><p className="report-status report-status-error">本周尚无地方债明细，请从首页上传。数据缺失不等于零发行。</p><label><input type="checkbox" checked={confirmedNoLocalWeek === weekStart} onChange={event => setConfirmedNoLocalWeek(event.target.checked ? weekStart : "")}/> 已核实本周确无地方债发行</label></>}
          </div>
          <div><span className="section-label">FINAL OUTPUT</span><h2>生成本周客户版周报</h2><p>地方债发行量及分类统计取自首页上传的地方债发行明细；国债政金债发行量及招标结果取自一二级表；四周滚动分析可在“一二级利差”页面查看和复制；新债发行Excel用于时段补充和差异核对，不覆盖一二级数据。</p>{!issuancePlanRecords.length && <p className="report-status report-status-error">请先在“周报生成数据”模块上传新债发行计划，系统才能准确区分上午和下午。</p>}{!maturityRecords.length && <p className="report-status report-status-error">请先上传本周到期明细，否则到期及净融资项目无法完整填充。</p>}{!spreadRecords.length && <p className="report-status report-status-error">请先上传一二级表，用于发行量、滚动分析和每日招标结果。</p>}{reportStatus && <p className={reportStatus.startsWith("生成失败") ? "report-status report-status-error" : "report-status"}>{reportStatus}</p>}{reportDownload && <a className="report-download" href={reportDownload.url} download={reportDownload.name}><Download/>下载已生成周报</a>}</div>
          <button onClick={exportDocx} disabled={reportLoading || !issuancePlanRecords.length || !maturityRecords.length || !spreadRecords.length || (!localRecords.length && confirmedNoLocalWeek !== weekStart)}>{reportLoading ? <LoaderCircle className="spin"/> : <FileText/>}{reportLoading ? "生成中" : "生成 Word"}</button>
        </section>}

        {active === "overview" && <section className="panel compact-panel">
          <div className="panel-head"><div><span className="section-label">SOURCE LOG</span><h2>本周文件</h2></div><button className="icon-button" onClick={loadWeek} aria-label="刷新"><RefreshCw/></button></div>
          {loading ? <div className="loading"><LoaderCircle className="spin"/>读取中</div> : data.imports.length ? <div className="file-list">{data.imports.map(item => <div key={item.id}><FileSpreadsheet/><span><strong>{item.file_name}</strong><small>{item.dataset_type === "local_bond" ? "地方债发行" : item.dataset_type === "maturity" ? "到期明细" : item.dataset_type === "issuance_plan" ? "新债发行计划" : "一二级"} · {item.record_count}条 · {item.trade_date}</small></span><button aria-label="删除批次" onClick={()=>deleteImport(item.id)}><X/></button></div>)}</div> : <Empty text="本周暂无入库文件" />}
        </section>}
      </section>
    </main>
  );
}

function Empty({text}:{text:string}) { return <div className="empty"><FileSpreadsheet/><span>{text}</span></div>; }
