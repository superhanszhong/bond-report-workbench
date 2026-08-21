"use client";

import { useEffect, useMemo, useRef, useState } from "react";
import {
  BarChart3, CalendarDays, Check, ChevronLeft, ChevronRight, CircleAlert,
  Download, FileSpreadsheet, FileText, LoaderCircle, RefreshCw, Save, Upload, X,
} from "lucide-react";
import {
  fridayOf, localPlanText, maturityWeekStart, mondayOf, parseLocalBondFile, parseMaturityFile,
  parseSpreadFile, ParsedBondRecord, resolveSpreadBp, spreadSummary,
} from "./lib/workbench";
import { buildWeeklyReportBlob } from "./lib/report";

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

type LatestDates = { local_bond?: string; spread?: string; maturity?: string };
type DatasetType = "local_bond" | "spread" | "maturity";

const SUMMARY_DRAFT_VERSION = "weekly-bond-summary-v2";

function encodeSummaryDraft(text: string) {
  return `${SUMMARY_DRAFT_VERSION}\n${text}`;
}

function decodeSummaryDraft(text?: string) {
  if (!text?.startsWith(`${SUMMARY_DRAFT_VERSION}\n`)) return null;
  return text.slice(SUMMARY_DRAFT_VERSION.length + 1);
}

const tabGroups = [
  { label: "工作台", items: [["overview", "本周总览"], ["maturity", "到期数据"]] },
  { label: "地方债", items: [["local", "发行明细"]] },
  { label: "利率债", items: [["spread", "一二级利差"], ["summary", "发行小结"], ["report", "周报生成"]] },
] as const;

type TabKey = (typeof tabGroups)[number]["items"][number][0];

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

function normalize(row: StoredRecord): ParsedBondRecord {
  let raw: Record<string, unknown> = row.raw || {};
  if (row.raw_json) {
    try { raw = JSON.parse(row.raw_json) as Record<string, unknown>; } catch { raw = {}; }
  }
  const storedMeta = raw.__summaryMeta && typeof raw.__summaryMeta === "object"
    ? raw.__summaryMeta as ParsedBondRecord["summaryMeta"] : undefined;
  const fallbackRoute = row.remark && /报价发行|前台报价/.test(row.remark)
    ? "报价发行" : /^09/.test(row.bond_code || row.bondCode || "") ? "上清所" : (row.issuance_route || row.issuanceRoute);
  const spreadResolution = row.dataset_type === "spread" ? resolveSpreadBp({
    provided: raw["利差"] ?? row.spread,
    allIn: raw["综收"], secondary: raw["二级"], winningRate: raw["中标利率"],
    issuer: row.issuer || "", remark: row.remark || "",
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
    spread: spreadResolution?.spread ?? row.spread,
    floorRate: row.floor_rate ?? row.floorRate,
    fee: row.fee,
    distributionDate: row.distribution_date || row.distributionDate,
    remark: row.remark,
    summaryMeta: storedMeta,
    raw,
  };
}

function tone(type: string) {
  return type === "国债" ? "treasury" : type === "国开债" ? "cdb" : type === "口行债" ? "exim" : "adbc";
}

function maturityKind(row: ParsedBondRecord) {
  const kind = row.raw?.__maturityKind;
  if (kind === "rate" || kind === "local") return kind;
  return ["国债", "国开债", "口行债", "农发债"].includes(row.bondType || "") ? "rate" : "local";
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

function SpreadChart({ records, svgRef, startDate, endDate }: { records: ParsedBondRecord[]; svgRef: React.RefObject<SVGSVGElement | null>; startDate: string; endDate: string }) {
  const issuerTypes = ["国债", "国开债", "口行债", "农发债"];
  const specialPattern = /绿债|绿色|主题债|浮息债/;
  const eligible = records.filter(row => issuerTypes.includes(row.bondType || "") && (() => {
    const day = new Date(`${row.tradeDate}T12:00:00`).getDay();
    return day >= 1 && day <= 5;
  })());
  const excluded = eligible.filter(row => specialPattern.test(row.remark || ""));
  const normalized = eligible.filter(row => row.spread !== null && row.spread !== undefined && !specialPattern.test(row.remark || "")).map((row, sourceIndex) => {
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
    const w = Math.max(112, label.length * 9.5 + 20);
    const h = 30;
    const candidates: { x: number; y: number }[] = [];
    const sideFirst = (row.sourceIndex + labelIndex) % 2 === 0;
    for (let lane = 0; lane < 8; lane += 1) {
      const vertical = 14 + lane * 34;
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
        <rect width="1680" height="1050" fill="#fff" />
        <rect width="1680" height="96" fill="#FAE7DA" />
        <text x="70" y="43" fontSize="34" fontWeight="700" fill="#9A5748">本周国债、政金债发行一二级利差散点图</text>
        <text x="70" y="76" fontSize="20" fill="#6B6662">交易日：{startDate} 至 {endDate}｜利差口径：综收－二级（bp）｜散点口径：单券</text>
        <text x="70" y="140" fontSize="22" fill="#68717b">普通债券样本 <tspan fontSize="26" fontWeight="700" fill="#D28A4D">{normalized.length}只</tspan>　｜　正利差债券 <tspan fontSize="26" fontWeight="700" fill="#D28A4D">{normalized.filter(r => Number(r.spread) > 0).length}只</tspan>　｜　平均利差 <tspan fontSize="26" fontWeight="700" fill="#D28A4D">{avg.toFixed(2)}bp</tspan>　｜　已排除{excluded.some(row => /绿债|绿色/.test(row.remark || "")) ? "绿债、" : ""}浮息债和主题债 <tspan fontSize="26" fontWeight="700" fill="#D28A4D">{excluded.length}只</tspan></text>
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
          <line x1={row.cx} y1={row.cy} x2={box.x+box.w/2} y2={box.y+box.h/2} stroke={darkColors[tone(row.bondType || "")]} strokeWidth="1.2" opacity=".75"/>
          <rect x={box.x} y={box.y} width={box.w} height={box.h} rx="7" fill={colors[tone(row.bondType || "")]} fillOpacity=".92" stroke={darkColors[tone(row.bondType || "")]} strokeWidth="1"/>
          <text x={box.x+box.w/2} y={box.y+20} textAnchor="middle" fontSize="18" fontWeight="700" fill="#3F3F3F">{label}</text>
        </g>)}
        <text x="860" y="950" textAnchor="middle" fontSize="21" fontWeight="600" fill="#626a73">发行期限</text>
        <text x="25" y="535" transform="rotate(-90 25 535)" textAnchor="middle" fontSize="21" fontWeight="600" fill="#626a73">综收－二级（bp）</text>
        {[["国债","#F4CDC3"],["国开债","#FAE1CC"],["口行债","#E5E5E3"],["农发债","#D6EFF5"]].map(([label,color],i) => <g key={label} transform={`translate(${485+i*190},1005)`}><circle r="11" fill={color}/><text x="22" y="7" fontSize="24" fontWeight="700" fill="#555b63">{label}</text></g>)}
      </svg>
      {!normalized.length && <div className="chart-empty">上传一二级表后，这里按所选区间生成利差图</div>}
    </div>
  );
}

export default function Workbench() {
  const [weekStart, setWeekStart] = useState(() => mondayOf(new Date()));
  const [active, setActive] = useState<TabKey>("overview");
  const [data, setData] = useState<WeekData>({ imports: [], records: [] });
  const [loading, setLoading] = useState(true);
  const [message, setMessage] = useState("");
  const [summary, setSummary] = useState("");
  const [saving, setSaving] = useState(false);
  const [dragging, setDragging] = useState<DatasetType | null>(null);
  const [analysisStart, setAnalysisStart] = useState(() => mondayOf(new Date()));
  const [analysisEnd, setAnalysisEnd] = useState(() => fridayOf(mondayOf(new Date())));
  const [chartRange, setChartRange] = useState(() => ({ start: mondayOf(new Date()), end: fridayOf(mondayOf(new Date())) }));
  const [chartRecords, setChartRecords] = useState<ParsedBondRecord[]>([]);
  const [spreadSourceRecords, setSpreadSourceRecords] = useState<ParsedBondRecord[]>([]);
  const [chartLoading, setChartLoading] = useState(false);
  const [reportLoading, setReportLoading] = useState(false);
  const [reportStatus, setReportStatus] = useState("");
  const [reportDownload, setReportDownload] = useState<{ url: string; name: string } | null>(null);
  const [latestDates, setLatestDates] = useState<LatestDates>({});
  const localInput = useRef<HTMLInputElement>(null);
  const spreadInput = useRef<HTMLInputElement>(null);
  const maturityInput = useRef<HTMLInputElement>(null);
  const svgRef = useRef<SVGSVGElement>(null);

  useEffect(() => () => {
    if (reportDownload) URL.revokeObjectURL(reportDownload.url);
  }, [reportDownload]);

  const records = useMemo(() => data.records.map(normalize), [data.records]);
  const localRecords = records.filter((_, i) => data.records[i]?.dataset_type === "local_bond");
  const spreadRecords = records.filter((_, i) => data.records[i]?.dataset_type === "spread");
  const maturityRecords = records.filter((_, i) => data.records[i]?.dataset_type === "maturity");
  const weekEnd = fridayOf(weekStart);
  const dailyDates = Array.from(new Set(localRecords.map((row) => row.tradeDate))).sort();
  const localAmount = localRecords.reduce((sum, row) => sum + (row.amount || 0), 0);
  const ordinary = spreadRecords.filter((row) => row.spread !== null && row.spread !== undefined && !/绿债|绿色|主题债|浮息债/.test(row.remark || ""));
  const latestLocalDate = latestDates.local_bond || "暂无数据";
  const latestSpreadDate = latestDates.spread || "暂无数据";
  const latestMaturityDate = latestDates.maturity || "暂无数据";
  const rateMaturityRecords = maturityRecords.filter(row => maturityKind(row) === "rate");
  const localMaturityRecords = maturityRecords.filter(row => maturityKind(row) === "local");
  const maturityTotal = maturityRecords.reduce((sum, row) => sum + (row.amount || 0), 0);
  const chartAudit = useMemo(() => {
    const supported = chartRecords.filter(row => ["国债", "国开债", "口行债", "农发债"].includes(row.bondType || ""));
    return {
      plotted: supported.filter(row => row.spread !== null && row.spread !== undefined && !/绿债|绿色|主题债|浮息债/.test(row.remark || "")).length,
      excluded: supported.filter(row => /绿债|绿色|主题债|浮息债/.test(row.remark || "")).length,
      derived: supported.filter(row => ["derived_treasury", "recalculated"].includes(spreadAuditStatus(row))).length,
      missing: supported.filter(row => (row.spread === null || row.spread === undefined) && !/绿债|绿色|主题债|浮息债/.test(row.remark || "")).length,
    };
  }, [chartRecords]);
  const legacySpreadData = spreadRecords.length > 0 && spreadRecords.some((row) => !row.summaryMeta);

  async function loadWeek() {
    setLoading(true);
    try {
      const [response, latestResponse] = await Promise.all([
        fetch(`/api/workbench?weekStart=${weekStart}`),
        fetch("/api/workbench?meta=latest"),
      ]);
      const payload = await response.json() as WeekData & { error?: string };
      const latestPayload = await latestResponse.json() as { latestDates?: LatestDates; error?: string };
      if (!response.ok) throw new Error(payload.error || "读取失败");
      if (!latestResponse.ok) throw new Error(latestPayload.error || "读取最新日期失败");
      setData(payload);
      setLatestDates(latestPayload.latestDates || {});
      const loadedSpread = payload.records.filter(r => r.dataset_type === "spread").map(normalize);
      const hasLegacySpread = loadedSpread.length > 0 && loadedSpread.some((row) => !row.summaryMeta);
      const generated = hasLegacySpread
        ? "当前一二级数据由旧版解析器入库，缺少发行渠道、备注、前次结果及边际投/中标量等字段。请重新上传原始一二级表，系统将按最新规则生成完整发行小结。"
        : spreadSummary(loadedSpread);
      const latestSpreadImport = payload.imports.filter(item => item.dataset_type === "spread")
        .map(item => item.created_at).sort().at(-1) || "";
      const currentDraft = decodeSummaryDraft(payload.draft?.summary_text);
      const draftIsCurrent = Boolean(currentDraft && payload.draft?.updated_at && payload.draft.updated_at >= latestSpreadImport && !hasLegacySpread);
      setSummary(draftIsCurrent ? currentDraft || generated : generated);
      setAnalysisStart(weekStart);
      setAnalysisEnd(fridayOf(weekStart));
      setChartRange({ start: weekStart, end: fridayOf(weekStart) });
      setChartRecords(loadedSpread);
      setMessage("");
    } catch (error) {
      setMessage(error instanceof Error ? error.message : "读取失败");
    } finally { setLoading(false); }
  }

  useEffect(() => { void loadWeek(); }, [weekStart]);

  async function upload(file: File, type: DatasetType) {
    setMessage(`正在解析 ${file.name}…`);
    try {
      const parsed = type === "local_bond" ? await parseLocalBondFile(file)
        : type === "spread" ? await parseSpreadFile(file) : await parseMaturityFile(file);
      if (type === "spread") setSpreadSourceRecords(parsed);
      const recordWeekOf = (row: ParsedBondRecord) => type === "maturity" ? maturityWeekStart(row.tradeDate) : mondayOf(row.tradeDate);
      const detectedWeeks = new Set(parsed.map(recordWeekOf));
      const isHistoricalBase = detectedWeeks.size > 1;
      const groups = new Map<string, ParsedBondRecord[]>();
      if (isHistoricalBase) {
        parsed.forEach((row) => {
          const recordWeek = recordWeekOf(row);
          groups.set(recordWeek, [...(groups.get(recordWeek) || []), row]);
        });
      } else {
        const inWeek = parsed.filter((row) => row.tradeDate >= weekStart && row.tradeDate <= weekEnd);
        if (!inWeek.length) throw new Error(`文件中没有 ${displayWeek(weekStart)} 的记录`);
        groups.set(weekStart, inWeek);
      }
      if (isHistoricalBase) setMessage(`已识别 ${groups.size} 个交易周，正在同步到共享数据库…`);
      let added = 0;
      let updated = 0;
      let unchanged = 0;
      const entries = [...groups.entries()];
      let cursor = 0;
      async function saveNext() {
        while (cursor < entries.length) {
          const [recordWeek, group] = entries[cursor++];
          const tradeDate = group.map((r) => r.tradeDate).sort().at(-1)!;
          const response = await fetch("/api/workbench", {
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
      await Promise.all(Array.from({ length: Math.min(6, entries.length) }, () => saveNext()));
      const result = `新增${added}条，更新${updated}条，保留${unchanged}条未变化记录`;
      setMessage(isHistoricalBase ? `历史数据已同步至共享数据库：${groups.size} 个交易周，${result}` : `已入库：${result}`);
      await loadWeek();
    } catch (error) { setMessage(error instanceof Error ? error.message : "上传失败"); }
  }

  function dropFile(event: React.DragEvent<HTMLButtonElement>, type: DatasetType) {
    event.preventDefault();
    setDragging(null);
    const file = event.dataTransfer.files?.[0];
    if (file) void upload(file, type);
  }

  async function generateSpreadRange() {
    if (!analysisStart || !analysisEnd || analysisStart > analysisEnd) {
      setMessage("请选择有效的利差分析日期区间");
      return;
    }
    setChartLoading(true);
    try {
      const response = await fetch(`/api/workbench?startDate=${analysisStart}&endDate=${analysisEnd}`);
      const payload = await response.json() as { records?: StoredRecord[]; error?: string };
      if (!response.ok) throw new Error(payload.error || "读取区间数据失败");
      const stored = (payload.records || []).map(normalize);
      const fromCurrentFile = spreadSourceRecords.filter(row => row.tradeDate >= analysisStart && row.tradeDate <= analysisEnd);
      const merged = new Map<string, ParsedBondRecord>();
      [...stored, ...fromCurrentFile].forEach(row => merged.set(`${row.tradeDate}|${row.bondCode || row.shortName || ""}`, row));
      const selected = [...merged.values()].sort((a, b) => `${a.tradeDate}${a.bondCode}`.localeCompare(`${b.tradeDate}${b.bondCode}`));
      setChartRecords(selected);
      setChartRange({ start: analysisStart, end: analysisEnd });
      setMessage(selected.length ? `已生成 ${analysisStart} 至 ${analysisEnd} 的利差图，共 ${selected.length} 条记录` : "所选区间暂无已入库的一二级数据");
    } catch (error) {
      setMessage(error instanceof Error ? error.message : "生成区间利差图失败");
    } finally { setChartLoading(false); }
  }

  async function deleteImport(importId: string) {
    const response = await fetch(`/api/workbench?importId=${encodeURIComponent(importId)}`, { method: "DELETE" });
    if (response.ok) await loadWeek();
  }

  async function saveDraft() {
    setSaving(true);
    try {
      const response = await fetch("/api/workbench", {
        method: "POST", headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ action: "saveDraft", weekStart, summaryText: encodeSummaryDraft(summary), reviewText: "" }),
      });
      if (!response.ok) throw new Error("保存失败");
      setMessage("发行小结草稿已保存");
    } finally { setSaving(false); }
  }

  function downloadChart() {
    if (!svgRef.current) return;
    const xml = new XMLSerializer().serializeToString(svgRef.current);
    const blob = new Blob([xml], { type: "image/svg+xml;charset=utf-8" });
    const url = URL.createObjectURL(blob);
    const image = new Image();
    image.onload = () => {
      const canvas = document.createElement("canvas"); canvas.width = 1680; canvas.height = 1050;
      canvas.getContext("2d")?.drawImage(image, 0, 0, 1680, 1050);
      canvas.toBlob((png) => {
        if (!png) return;
        const a = document.createElement("a"); a.href = URL.createObjectURL(png);
        a.download = `国债政金债一二级利差_${chartRange.start}_${chartRange.end}.png`; a.click();
        URL.revokeObjectURL(a.href); URL.revokeObjectURL(url);
      }, "image/png");
    };
    image.src = url;
  }

  async function exportDocx() {
    setReportLoading(true);
    setReportStatus("正在按今日 Word 母版填充数据…");
    try {
      const mmdd = (d: string) => d.slice(5).replace("-", "");
      const previousStart = shiftWeek(weekStart, -1);
      const yearStart = `${weekStart.slice(0, 4)}-01-01`;
      const [previousResponse, ytdLocalResponse] = await Promise.all([
        fetch(`/api/workbench?weekStart=${previousStart}`),
        fetch(`/api/workbench?startDate=${yearStart}&endDate=${weekEnd}&datasetType=local_bond`),
      ]);
      const previousPayload = await previousResponse.json() as WeekData & { error?: string };
      const ytdLocalPayload = await ytdLocalResponse.json() as { records?: StoredRecord[]; error?: string };
      if (!previousResponse.ok) throw new Error(previousPayload.error || "读取上周数据失败");
      if (!ytdLocalResponse.ok) throw new Error(ytdLocalPayload.error || "读取地方债年度数据失败");
      const previousSpreadRecords = previousPayload.records.filter((row) => row.dataset_type === "spread").map(normalize);
      const previousMaturityRecords = previousPayload.records.filter((row) => row.dataset_type === "maturity").map(normalize);
      const ytdLocalRecords = (ytdLocalPayload.records || []).map(normalize);
      const rateTotal = rateMaturityRecords.reduce((sum, row) => sum + (row.amount || 0), 0);
      const localDaily = Object.fromEntries(Array.from({ length: 5 }, (_, index) => {
        const date = new Date(`${weekStart}T12:00:00`); date.setDate(date.getDate() + index);
        const iso = date.toISOString().slice(0, 10);
        return [iso, localMaturityRecords.filter(row => row.tradeDate === iso).reduce((sum, row) => sum + (row.amount || 0), 0)];
      }));
      const rateLabels = new Map<string, number>();
      rateMaturityRecords.forEach(row => rateLabels.set(row.bondType || "利率债", (rateLabels.get(row.bondType || "利率债") || 0) + (row.amount || 0)));
      const previousRateMaturity = previousMaturityRecords.filter(row => maturityKind(row) === "rate").reduce((sum, row) => sum + (row.amount || 0), 0);
      const previousRateIssuance = previousSpreadRecords.reduce((sum, row) => sum + (row.amount || 0), 0);
      const maturity = maturityRecords.length ? {
        rateTotal,
        rateBreakdown: [...rateLabels].map(([label, value]) => `${label}:${Number(value.toFixed(4))}亿`).join("　") || "-",
        localDaily,
        localTotal: localMaturityRecords.reduce((sum, row) => sum + (row.amount || 0), 0),
        previousRateNet: previousMaturityRecords.length ? previousRateIssuance - previousRateMaturity : undefined,
      } : undefined;
      const blob = await buildWeeklyReportBlob({ weekStart, summary, localRecords, spreadRecords, previousSpreadRecords, ytdLocalRecords, maturity });
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
        <div className="brand"><span className="brand-mark">债</span><div><strong>利率债发行工作台</strong><small>Issuance Desk</small></div></div>
        <nav>{tabGroups.map(group => <div className="nav-group" key={group.label}><span className="nav-group-label">{group.label}</span>{group.items.map(([key, label]) => <button key={key} className={active === key ? "nav-active" : ""} onClick={() => setActive(key)}>{key === "overview" ? <CalendarDays/> : key === "local" ? <FileSpreadsheet/> : key === "spread" ? <BarChart3/> : <FileText/>}<span>{label}</span></button>)}</div>)}</nav>
        <a className="legacy-link" href="https://superhanszhong.github.io/local-bond-daily-converter/" target="_blank" rel="noreferrer">原日表转换器 ↗</a>
        <div className="side-note"><Check size={16}/><span>公开共享数据<br/>周度记录集中留存</span></div>
      </aside>

      <section className="workspace">
        <header className="topbar">
          <div><p className="eyebrow">WEEKLY ISSUANCE</p><h1>{displayWeek(weekStart)} 利率债发行</h1><p>{weekStart} 至 {weekEnd} · 仅统计交易日</p></div>
          <div className="week-picker"><button aria-label="上一周" onClick={() => setWeekStart(shiftWeek(weekStart,-1))}><ChevronLeft/></button><input type="date" value={weekStart} onChange={e => setWeekStart(mondayOf(e.target.value))}/><button aria-label="下一周" onClick={() => setWeekStart(shiftWeek(weekStart,1))}><ChevronRight/></button></div>
        </header>

        {message && <div className={`notice ${/失败|缺少|没有/.test(message) ? "notice-error" : ""}`}><CircleAlert size={17}/>{message}<button onClick={() => setMessage("")}><X size={15}/></button></div>}

        <section className="hero-grid">
          <div className="upload-card">
            <div className="card-head"><div><span className="section-label">双板块入口</span><h2>拖入文件，自动归入对应板块</h2></div><Upload/></div>
            <div className="upload-actions">
              <div className="upload-lane">
                <button
                  className={`drop-zone local-drop ${dragging === "local_bond" ? "drag-active" : ""}`}
                  onClick={() => localInput.current?.click()}
                  onDragEnter={event => { event.preventDefault(); setDragging("local_bond"); }}
                  onDragOver={event => { event.preventDefault(); event.dataTransfer.dropEffect = "copy"; }}
                  onDragLeave={() => setDragging(null)}
                  onDrop={event => dropFile(event, "local_bond")}
                ><FileSpreadsheet/><span><em>地方债板块</em><strong>地方债发行明细</strong><small>拖入或点击选择 Excel<br/>生成日表并纳入周汇总</small></span></button>
                <div className="latest-date"><CalendarDays/><span>地方债最新发行日</span><strong>{latestLocalDate}</strong></div>
              </div>
              <div className="upload-lane">
                <button
                  className={`drop-zone rate-drop ${dragging === "spread" ? "drag-active" : ""}`}
                  onClick={() => spreadInput.current?.click()}
                  onDragEnter={event => { event.preventDefault(); setDragging("spread"); }}
                  onDragOver={event => { event.preventDefault(); event.dataTransfer.dropEffect = "copy"; }}
                  onDragLeave={() => setDragging(null)}
                  onDrop={event => dropFile(event, "spread")}
                ><BarChart3/><span><em>利率债板块</em><strong>国债及政金债一二级表</strong><small>拖入或点击选择 Excel<br/>生成利差图与发行小结</small></span></button>
                <div className="latest-date spread-date"><CalendarDays/><span>一二级利差最新日期</span><strong>{latestSpreadDate}</strong></div>
              </div>
              <div className="upload-lane">
                <button
                  className={`drop-zone maturity-drop ${dragging === "maturity" ? "drag-active" : ""}`}
                  onClick={() => maturityInput.current?.click()}
                  onDragEnter={event => { event.preventDefault(); setDragging("maturity"); }}
                  onDragOver={event => { event.preventDefault(); event.dataTransfer.dropEffect = "copy"; }}
                  onDragLeave={() => setDragging(null)}
                  onDrop={event => dropFile(event, "maturity")}
                ><CalendarDays/><span><em>周报数据</em><strong>本周到期明细</strong><small>拖入或点击选择 Excel<br/>自动区分地方债与国债政金债</small></span></button>
                <div className="latest-date maturity-date"><CalendarDays/><span>到期数据最新日期</span><strong>{latestMaturityDate}</strong></div>
              </div>
            </div>
            <input ref={localInput} hidden type="file" accept=".xlsx,.xlsm" onChange={e => { const file=e.target.files?.[0]; if(file) void upload(file,"local_bond"); e.currentTarget.value=""; }}/>
            <input ref={spreadInput} hidden type="file" accept=".xlsx,.xlsm" onChange={e => { const file=e.target.files?.[0]; if(file) void upload(file,"spread"); e.currentTarget.value=""; }}/>
            <input ref={maturityInput} hidden type="file" accept=".xlsx,.xlsm" onChange={e => { const file=e.target.files?.[0]; if(file) void upload(file,"maturity"); e.currentTarget.value=""; }}/>
            <p>将文件拖入对应色块即可上传；重复上传同一债券时，仅更新发生变化的字段。</p>
          </div>
        </section>

        <section className="overview-section">
          <div className="overview-heading"><span className="section-label">WEEKLY OVERVIEW</span><h2>本周发行概览</h2></div>
          <div className="metrics">
            <article><span>地方债发行</span><strong>{localAmount.toFixed(2)}</strong><small>亿元 · {localRecords.length}只</small></article>
            <article><span>普通利率债样本</span><strong>{ordinary.length}</strong><small>只 · 已按周报口径筛选</small></article>
            <article><span>正利差债券</span><strong>{ordinary.filter(r => (r.spread || 0)>0).length}</strong><small>只 · 综收高于二级</small></article>
            <article><span>本周到期</span><strong>{maturityTotal.toFixed(2)}</strong><small>亿元 · 利率债{rateMaturityRecords.length}只 / 地方债{localMaturityRecords.length}只</small></article>
          </div>
        </section>

        {(active === "overview" || active === "local") && <section className="panel">
          <div className="panel-head"><div><span className="section-label">DAILY RECORDS</span><h2>本周地方债发行明细</h2></div><span className="pill">{localRecords.length} 条</span></div>
          {dailyDates.length ? dailyDates.map(date => <div className="day-block" key={date}><div className="day-title"><strong>{formatMd(date)}</strong><span>{localRecords.filter(r=>r.tradeDate===date).length}只 · {localRecords.filter(r=>r.tradeDate===date).reduce((s,r)=>s+(r.amount||0),0).toFixed(2)}亿元</span></div><pre>{localPlanText(localRecords,date)}</pre></div>) : <Empty text="尚未上传本周地方债发行表" />}
        </section>}

        {(active === "overview" || active === "maturity") && <section className="panel compact-panel">
          <div className="panel-head"><div><span className="section-label">MATURITY DATA</span><h2>本周到期数据</h2></div><span className="pill">{maturityRecords.length} 条 · {maturityTotal.toFixed(4)}亿元</span></div>
          {maturityRecords.length ? <div className="maturity-grid">
            <article><span>国债及政金债</span><strong>{rateMaturityRecords.reduce((sum,row)=>sum+(row.amount||0),0).toFixed(4)}亿元</strong><small>{rateMaturityRecords.map(row=>`${row.shortName || row.bondCode} ${row.amount}亿`).join(" · ") || "暂无"}</small></article>
            <article><span>地方政府债</span><strong>{localMaturityRecords.reduce((sum,row)=>sum+(row.amount||0),0).toFixed(4)}亿元</strong><small>{Array.from(new Set(localMaturityRecords.map(row=>row.tradeDate))).sort().map(date=>`${formatMd(date)} ${localMaturityRecords.filter(row=>row.tradeDate===date).reduce((sum,row)=>sum+(row.amount||0),0).toFixed(4)}亿`).join(" · ") || "暂无"}</small></article>
          </div> : <Empty text="尚未上传本周到期明细" />}
        </section>}

        {(active === "overview" || active === "spread") && <section className="panel">
          <div className="panel-head spread-panel-head"><div><span className="section-label">PRIMARY / SECONDARY</span><h2>一二级利差可视化</h2></div><button className="secondary" onClick={downloadChart} disabled={!chartRecords.length}><Download/>下载 PNG</button></div>
          <div className="range-toolbar">
            <label><span>开始日期</span><input type="date" value={analysisStart} onChange={event => setAnalysisStart(event.target.value)}/></label>
            <span className="range-divider">至</span>
            <label><span>结束日期</span><input type="date" value={analysisEnd} onChange={event => setAnalysisEnd(event.target.value)}/></label>
            <button className="range-generate" onClick={generateSpreadRange} disabled={chartLoading}>{chartLoading ? <LoaderCircle className="spin"/> : <BarChart3/>}生成图表</button>
            <p>当前区间：{chartRange.start} 至 {chartRange.end} · {chartRecords.length} 条记录</p>
          </div>
          <div className="chart-audit" aria-label="利差数据校验结果">
            <span>有效样本 <strong>{chartAudit.plotted}</strong></span>
            <span>按口径排除 <strong>{chartAudit.excluded}</strong></span>
            <span>自动复算 <strong>{chartAudit.derived}</strong></span>
            <span className={chartAudit.missing ? "audit-warning" : ""}>缺少可比值 <strong>{chartAudit.missing}</strong></span>
          </div>
          <SpreadChart records={chartRecords} svgRef={svgRef} startDate={chartRange.start} endDate={chartRange.end}/>
        </section>}

        {(active === "overview" || active === "summary") && <section className="panel">
          <div className="panel-head"><div><span className="section-label">CLIENT COPY</span><h2>周报发行小结</h2></div><button className="secondary" onClick={saveDraft} disabled={saving}>{saving?<LoaderCircle className="spin"/>:<Save/>}保存草稿</button></div>
          <textarea className="summary-editor" value={summary} onChange={e=>setSummary(e.target.value)} placeholder="上传一二级表后自动生成，可在此复核和修改。"/>
          <div className="audit-row">{legacySpreadData ? <CircleAlert/> : <Check/>}<span>{legacySpreadData ? "旧版存量数据缺少新版小结字段，请重新上传原始一二级表后生成。" : "国债逐只保留；政金债按发行渠道分层；绿债、主题债和浮息债从普通散点图中剔除。"}</span></div>
        </section>}

        {(active === "overview" || active === "report") && <section className="report-panel">
          <div><span className="section-label">FINAL OUTPUT</span><h2>生成本周客户版周报</h2><p>直接使用今日 Word 原文件作为母版，自动填入发行、到期、净融资、每日回顾和表格数据。</p>{!maturityRecords.length && <p className="report-status report-status-error">生成前请先上传本周到期明细，否则到期及净融资项目无法完整填充。</p>}{reportStatus && <p className={reportStatus.startsWith("生成失败") ? "report-status report-status-error" : "report-status"}>{reportStatus}</p>}{reportDownload && <a className="report-download" href={reportDownload.url} download={reportDownload.name}><Download/>下载已生成周报</a>}</div>
          <button onClick={exportDocx} disabled={reportLoading || (!localRecords.length && !spreadRecords.length)}>{reportLoading ? <LoaderCircle className="spin"/> : <FileText/>}{reportLoading ? "生成中" : "生成 Word"}</button>
        </section>}

        <section className="panel compact-panel">
          <div className="panel-head"><div><span className="section-label">SOURCE LOG</span><h2>本周文件</h2></div><button className="icon-button" onClick={loadWeek} aria-label="刷新"><RefreshCw/></button></div>
          {loading ? <div className="loading"><LoaderCircle className="spin"/>读取中</div> : data.imports.length ? <div className="file-list">{data.imports.map(item => <div key={item.id}><FileSpreadsheet/><span><strong>{item.file_name}</strong><small>{item.dataset_type === "local_bond" ? "地方债发行" : item.dataset_type === "maturity" ? "到期明细" : "一二级"} · {item.record_count}条 · {item.trade_date}</small></span><button aria-label="删除批次" onClick={()=>deleteImport(item.id)}><X/></button></div>)}</div> : <Empty text="本周暂无入库文件" />}
        </section>
      </section>
    </main>
  );
}

function Empty({text}:{text:string}) { return <div className="empty"><FileSpreadsheet/><span>{text}</span></div>; }
