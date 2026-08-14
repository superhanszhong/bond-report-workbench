"use client";

import { useEffect, useMemo, useRef, useState } from "react";
import {
  BarChart3, CalendarDays, Check, ChevronLeft, ChevronRight, CircleAlert,
  Download, FileSpreadsheet, FileText, LoaderCircle, RefreshCw, Save, Upload, X,
} from "lucide-react";
import {
  fridayOf, localPlanText, mondayOf, parseLocalBondFile, parseSpreadFile,
  ParsedBondRecord, spreadSummary,
} from "./lib/workbench";

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
  draft?: { summary_text?: string; review_text?: string } | null;
};

const tabGroups = [
  { label: "工作台", items: [["overview", "本周总览"]] },
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
  return {
    tradeDate: row.trade_date || row.tradeDate,
    bondCode: row.bond_code || row.bondCode,
    shortName: row.short_name || row.shortName,
    fullName: row.full_name || row.fullName,
    issuer: row.issuer,
    region: row.region,
    bondType: row.bond_type || row.bondType,
    issuanceRoute: row.issuance_route || row.issuanceRoute,
    venue: row.venue,
    bidTime: row.bid_time || row.bidTime,
    tenor: row.tenor,
    amount: row.amount,
    spread: row.spread,
    floorRate: row.floor_rate ?? row.floorRate,
    fee: row.fee,
    distributionDate: row.distribution_date || row.distributionDate,
    remark: row.remark,
  };
}

function tone(type: string) {
  return type === "国债" ? "treasury" : type === "国开债" ? "cdb" : type === "口行债" ? "exim" : "adbc";
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
  const eligible = records.filter(row => row.spread !== null && row.spread !== undefined && issuerTypes.includes(row.bondType || "") && (() => {
    const day = new Date(`${row.tradeDate}T12:00:00`).getDay();
    return day >= 1 && day <= 5;
  })());
  const excluded = eligible.filter(row => specialPattern.test(row.remark || ""));
  const normalized = eligible.filter(row => !specialPattern.test(row.remark || "")).map((row, sourceIndex) => {
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
  const colors: Record<string, string> = { treasury: "#F4D8D3", cdb: "#F7E3CF", exim: "#E6E7E5", adbc: "#D8EEF5" };
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
        <rect width="1680" height="96" fill="#F7E3CF" />
        <text x="70" y="43" fontSize="34" fontWeight="700" fill="#9B642F">国债、政金债发行一二级利差散点图</text>
        <text x="70" y="76" fontSize="20" fill="#6B6662">交易日：{startDate} 至 {endDate}｜利差口径：综收－二级（bp）｜散点口径：单券</text>
        <text x="70" y="140" fontSize="22" fill="#68717b">普通债券样本 <tspan fontSize="26" fontWeight="700" fill="#D28A4D">{normalized.length}只</tspan>　｜　正利差债券 <tspan fontSize="26" fontWeight="700" fill="#D28A4D">{normalized.filter(r => Number(r.spread) > 0).length}只</tspan>　｜　平均利差 <tspan fontSize="26" fontWeight="700" fill="#D28A4D">{avg.toFixed(2)}bp</tspan>　｜　已排除{excluded.some(row => /绿债|绿色/.test(row.remark || "")) ? "绿债、" : ""}浮息债和主题债 <tspan fontSize="26" fontWeight="700" fill="#D28A4D">{excluded.length}只</tspan></text>
        <rect x={chartLeft} y={chartTop} width={chartRight-chartLeft} height={Math.max(0, y(0)-chartTop)} fill="#FFFCFA" />
        {ticks.map((tick) => (
          <g key={tick}>
            <line x1={chartLeft} y1={y(tick)} x2={chartRight} y2={y(tick)} stroke={tick === 0 ? "#D8B864" : "#E7E6E6"} strokeWidth={tick === 0 ? 2.3 : 1} strokeDasharray={tick === 0 ? "9 6" : undefined} />
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
        {[["国债","#F4D8D3"],["国开债","#F7E3CF"],["口行债","#E6E7E5"],["农发债","#D8EEF5"]].map(([label,color],i) => <g key={label} transform={`translate(${485+i*190},1005)`}><circle r="11" fill={color}/><text x="22" y="7" fontSize="24" fontWeight="700" fill="#555b63">{label}</text></g>)}
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
  const [dragging, setDragging] = useState<"local_bond" | "spread" | null>(null);
  const [analysisStart, setAnalysisStart] = useState(() => mondayOf(new Date()));
  const [analysisEnd, setAnalysisEnd] = useState(() => fridayOf(mondayOf(new Date())));
  const [chartRange, setChartRange] = useState(() => ({ start: mondayOf(new Date()), end: fridayOf(mondayOf(new Date())) }));
  const [chartRecords, setChartRecords] = useState<ParsedBondRecord[]>([]);
  const [spreadSourceRecords, setSpreadSourceRecords] = useState<ParsedBondRecord[]>([]);
  const [chartLoading, setChartLoading] = useState(false);
  const localInput = useRef<HTMLInputElement>(null);
  const spreadInput = useRef<HTMLInputElement>(null);
  const svgRef = useRef<SVGSVGElement>(null);

  const records = useMemo(() => data.records.map(normalize), [data.records]);
  const localRecords = records.filter((_, i) => data.records[i]?.dataset_type === "local_bond");
  const spreadRecords = records.filter((_, i) => data.records[i]?.dataset_type === "spread");
  const weekEnd = fridayOf(weekStart);
  const dailyDates = Array.from(new Set(localRecords.map((row) => row.tradeDate))).sort();
  const localAmount = localRecords.reduce((sum, row) => sum + (row.amount || 0), 0);
  const ordinary = spreadRecords.filter((row) => row.spread !== null && row.spread !== undefined && !/绿债|绿色|主题债|浮息债/.test(row.remark || ""));

  async function loadWeek() {
    setLoading(true);
    try {
      const response = await fetch(`/api/workbench?weekStart=${weekStart}`);
      const payload = await response.json() as WeekData & { error?: string };
      if (!response.ok) throw new Error(payload.error || "读取失败");
      setData(payload);
      const loadedSpread = payload.records.filter(r => r.dataset_type === "spread").map(normalize);
      const generated = spreadSummary(loadedSpread);
      setSummary(payload.draft?.summary_text || generated);
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

  async function upload(file: File, type: "local_bond" | "spread") {
    setMessage(`正在解析 ${file.name}…`);
    try {
      const parsed = type === "local_bond" ? await parseLocalBondFile(file) : await parseSpreadFile(file);
      if (type === "spread") setSpreadSourceRecords(parsed);
      const detectedWeeks = new Set(parsed.map((row) => mondayOf(row.tradeDate)));
      const isHistoricalLocalBase = type === "local_bond" && detectedWeeks.size > 1;
      const groups = new Map<string, ParsedBondRecord[]>();
      if (isHistoricalLocalBase) {
        parsed.forEach((row) => {
          const recordWeek = mondayOf(row.tradeDate);
          groups.set(recordWeek, [...(groups.get(recordWeek) || []), row]);
        });
      } else {
        const inWeek = parsed.filter((row) => row.tradeDate >= weekStart && row.tradeDate <= weekEnd);
        if (!inWeek.length) throw new Error(`文件中没有 ${displayWeek(weekStart)} 的记录`);
        groups.set(weekStart, inWeek);
      }
      let saved = 0;
      for (const [recordWeek, group] of groups) {
        const tradeDate = group.map((r) => r.tradeDate).sort().at(-1)!;
        const response = await fetch("/api/workbench", {
          method: "POST", headers: { "Content-Type": "application/json" },
          body: JSON.stringify({ datasetType: type, tradeDate, weekStart: recordWeek, fileName: file.name, records: group }),
        });
        const payload = await response.json() as { error?: string; count?: number };
        if (!response.ok) throw new Error(payload.error || "保存失败");
        saved += payload.count || 0;
      }
      setMessage(isHistoricalLocalBase ? `全年底库已拆分为 ${groups.size} 个交易周，共入库 ${saved} 条记录` : `已入库 ${saved} 条记录`);
      await loadWeek();
    } catch (error) { setMessage(error instanceof Error ? error.message : "上传失败"); }
  }

  function dropFile(event: React.DragEvent<HTMLButtonElement>, type: "local_bond" | "spread") {
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
        body: JSON.stringify({ action: "saveDraft", weekStart, summaryText: summary, reviewText: "" }),
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
    const { Document, Packer, Paragraph, Table, TableCell, TableRow, TextRun, WidthType, AlignmentType, HeadingLevel } = await import("docx");
    const mmdd = (d: string) => d.slice(5).replace("-", "");
    const children: (InstanceType<typeof Paragraph> | InstanceType<typeof Table>)[] = [
      new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: `利率债发行周报${mmdd(weekStart)}-${mmdd(weekEnd)}`, bold: true, size: 36, color: "000000" })] }),
      new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun({ text: "周报发行小结", bold: true, color: "000000" })] }),
      ...summary.split(/\n+/).filter(Boolean).map((text) => new Paragraph({ children: [new TextRun({ text, size: 21, color: "000000" })], spacing: { after: 120, line: 300 } })),
      new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun({ text: "本周地方债发行明细", bold: true, color: "000000" })] }),
      new Table({ width: { size: 100, type: WidthType.PERCENTAGE }, rows: [
        new TableRow({ tableHeader: true, children: ["日期","代码","简称","期限","发行量（亿元）","下限","场所"].map(t => new TableCell({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children:[new TextRun({text:t,bold:true,color:"000000"})] })] })) }),
        ...localRecords.map((row) => new TableRow({ cantSplit: true, children: [row.tradeDate,row.bondCode||"",row.shortName||"",row.tenor||"",String(row.amount??""),String(row.floorRate??""),row.venue||""].map(t => new TableCell({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children:[new TextRun({text:t,color:"000000"})] })] })) })),
      ] }),
      new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun({ text: "本周发行回顾", bold: true, color: "000000" })] }),
      ...dailyDates.flatMap((date) => [
        new Paragraph({ keepNext: true, children: [new TextRun({ text: `${formatMd(date)} 发行回顾`, bold: true, color: "000000" })] }),
        ...localPlanText(localRecords, date).split("\n").map(text => new Paragraph({ keepNext: true, children:[new TextRun({text,color:"000000"})] })),
      ]),
    ];
    const doc = new Document({ sections: [{ properties: {}, children }] });
    const blob = await Packer.toBlob(doc);
    const a = document.createElement("a"); a.href = URL.createObjectURL(blob);
    a.download = `利率债发行周报${mmdd(weekStart)}-${mmdd(weekEnd)}.docx`; a.click(); URL.revokeObjectURL(a.href);
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
          <div className="status-card">
            <div className="section-label">本周完成度</div>
            <div className="progress-ring" style={{"--progress": `${Math.min(100,(data.imports.length>0?45:0)+(spreadRecords.length?35:0)+(summary?20:0))}%`} as React.CSSProperties}><span>{Math.min(100,(data.imports.length>0?45:0)+(spreadRecords.length?35:0)+(summary?20:0))}%</span></div>
            <div className="checklist">
              <div className={localRecords.length ? "done" : ""}><span>{localRecords.length ? <Check/> : "1"}</span>地方债日表入库</div>
              <div className={spreadRecords.length ? "done" : ""}><span>{spreadRecords.length ? <Check/> : "2"}</span>一二级表入库</div>
              <div className={summary ? "done" : ""}><span>{summary ? <Check/> : "3"}</span>小结草稿生成</div>
            </div>
          </div>
          <div className="upload-card">
            <div className="card-head"><div><span className="section-label">双板块入口</span><h2>拖入文件，自动归入对应板块</h2></div><Upload/></div>
            <div className="upload-actions">
              <button
                className={`drop-zone local-drop ${dragging === "local_bond" ? "drag-active" : ""}`}
                onClick={() => localInput.current?.click()}
                onDragEnter={event => { event.preventDefault(); setDragging("local_bond"); }}
                onDragOver={event => { event.preventDefault(); event.dataTransfer.dropEffect = "copy"; }}
                onDragLeave={() => setDragging(null)}
                onDrop={event => dropFile(event, "local_bond")}
              ><FileSpreadsheet/><span><em>地方债板块</em><strong>地方债发行明细</strong><small>拖入或点击选择 Excel<br/>生成日表并纳入周汇总</small></span></button>
              <button
                className={`drop-zone rate-drop ${dragging === "spread" ? "drag-active" : ""}`}
                onClick={() => spreadInput.current?.click()}
                onDragEnter={event => { event.preventDefault(); setDragging("spread"); }}
                onDragOver={event => { event.preventDefault(); event.dataTransfer.dropEffect = "copy"; }}
                onDragLeave={() => setDragging(null)}
                onDrop={event => dropFile(event, "spread")}
              ><BarChart3/><span><em>利率债板块</em><strong>国债及政金债一二级表</strong><small>拖入或点击选择 Excel<br/>生成利差图与发行小结</small></span></button>
            </div>
            <input ref={localInput} hidden type="file" accept=".xlsx,.xlsm" onChange={e => { const file=e.target.files?.[0]; if(file) void upload(file,"local_bond"); e.currentTarget.value=""; }}/>
            <input ref={spreadInput} hidden type="file" accept=".xlsx,.xlsm" onChange={e => { const file=e.target.files?.[0]; if(file) void upload(file,"spread"); e.currentTarget.value=""; }}/>
            <p>将文件拖入对应色块即可上传；重复上传同一债券时更新已有记录。</p>
          </div>
        </section>

        <section className="metrics">
          <article><span>地方债发行</span><strong>{localAmount.toFixed(2)}</strong><small>亿元 · {localRecords.length}只</small></article>
          <article><span>普通利率债样本</span><strong>{ordinary.length}</strong><small>只 · 已按周报口径筛选</small></article>
          <article><span>正利差债券</span><strong>{ordinary.filter(r => (r.spread || 0)>0).length}</strong><small>只 · 综收高于二级</small></article>
          <article><span>已入库批次</span><strong>{data.imports.length}</strong><small>个 · {dailyDates.length}个发行日</small></article>
        </section>

        {(active === "overview" || active === "local") && <section className="panel">
          <div className="panel-head"><div><span className="section-label">DAILY RECORDS</span><h2>本周地方债发行明细</h2></div><span className="pill">{localRecords.length} 条</span></div>
          {dailyDates.length ? dailyDates.map(date => <div className="day-block" key={date}><div className="day-title"><strong>{formatMd(date)}</strong><span>{localRecords.filter(r=>r.tradeDate===date).length}只 · {localRecords.filter(r=>r.tradeDate===date).reduce((s,r)=>s+(r.amount||0),0).toFixed(2)}亿元</span></div><pre>{localPlanText(localRecords,date)}</pre></div>) : <Empty text="尚未上传本周地方债发行表" />}
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
          <SpreadChart records={chartRecords} svgRef={svgRef} startDate={chartRange.start} endDate={chartRange.end}/>
        </section>}

        {(active === "overview" || active === "summary") && <section className="panel">
          <div className="panel-head"><div><span className="section-label">CLIENT COPY</span><h2>周报发行小结</h2></div><button className="secondary" onClick={saveDraft} disabled={saving}>{saving?<LoaderCircle className="spin"/>:<Save/>}保存草稿</button></div>
          <textarea className="summary-editor" value={summary} onChange={e=>setSummary(e.target.value)} placeholder="上传一二级表后自动生成，可在此复核和修改。"/>
          <div className="audit-row"><Check/><span>国债逐只保留；政金债按发行渠道分层；绿债、主题债和浮息债从普通散点图中剔除。</span></div>
        </section>}

        {(active === "overview" || active === "report") && <section className="report-panel">
          <div><span className="section-label">FINAL OUTPUT</span><h2>生成本周客户版周报</h2><p>标题自动使用“利率债发行周报MMDD-MMDD”，正文黑色极简，地方债明细与发行回顾来自已入库数据。</p></div>
          <button onClick={exportDocx} disabled={!localRecords.length && !spreadRecords.length}><FileText/>生成 Word</button>
        </section>}

        <section className="panel compact-panel">
          <div className="panel-head"><div><span className="section-label">SOURCE LOG</span><h2>本周文件</h2></div><button className="icon-button" onClick={loadWeek} aria-label="刷新"><RefreshCw/></button></div>
          {loading ? <div className="loading"><LoaderCircle className="spin"/>读取中</div> : data.imports.length ? <div className="file-list">{data.imports.map(item => <div key={item.id}><FileSpreadsheet/><span><strong>{item.file_name}</strong><small>{item.dataset_type === "local_bond" ? "地方债" : "一二级"} · {item.record_count}条 · {item.trade_date}</small></span><button aria-label="删除批次" onClick={()=>deleteImport(item.id)}><X/></button></div>)}</div> : <Empty text="本周暂无入库文件" />}
        </section>
      </section>
    </main>
  );
}

function Empty({text}:{text:string}) { return <div className="empty"><FileSpreadsheet/><span>{text}</span></div>; }
