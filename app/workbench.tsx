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

const tabs = [
  ["overview", "本周总览"], ["local", "地方债日表"], ["spread", "一二级利差"],
  ["summary", "发行小结"], ["report", "周报生成"],
] as const;

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

function SpreadChart({ records, svgRef }: { records: ParsedBondRecord[]; svgRef: React.RefObject<SVGSVGElement | null> }) {
  const filtered = records.filter((row) => row.spread !== null && row.spread !== undefined &&
    ["国债", "国开债", "口行债", "农发债"].includes(row.bondType || "") &&
    !/绿债|绿色|主题债|浮息债/.test(row.remark || ""));
  const tenors = Array.from(new Set(filtered.map((row) => row.tenor || "—")));
  const values = filtered.map((row) => row.spread || 0);
  const min = Math.min(-1, ...values) - 0.8;
  const max = Math.max(1, ...values) + 0.8;
  const x = (tenor: string) => 110 + Math.max(0, tenors.indexOf(tenor)) * (1320 / Math.max(1, tenors.length - 1));
  const y = (value: number) => 780 - ((value - min) / (max - min)) * 490;
  const avg = filtered.length ? values.reduce((a, b) => a + b, 0) / filtered.length : 0;
  return (
    <div className="chart-wrap">
      <svg ref={svgRef} className="spread-chart" viewBox="0 0 1680 1050" role="img" aria-label="本周国债、政金债发行一二级利差散点图">
        <rect width="1680" height="1050" fill="#fff" />
        <rect width="1680" height="126" fill="#FAE7DA" />
        <text x="78" y="72" fontSize="38" fontWeight="700" fill="#9A5748">本周国债、政金债发行一二级利差散点图</text>
        <text x="80" y="166" fontSize="21" fill="#6b7280">利差口径：综收－二级（bp）｜散点口径：单券</text>
        <text x="80" y="220" fontSize="23" fill="#68717b">普通债券样本 <tspan fontWeight="700" fill="#9A5748">{filtered.length}只</tspan> ｜ 正利差债券 <tspan fontWeight="700" fill="#9A5748">{filtered.filter(r => (r.spread || 0) > 0).length}只</tspan> ｜ 平均利差 <tspan fontWeight="700" fill="#9A5748">{avg.toFixed(2)}bp</tspan></text>
        <rect x="80" y={y(max)} width="1500" height={y(0) - y(max)} fill="#FFFCFA" />
        {[min, min + (max-min)/4, min + (max-min)/2, min + 3*(max-min)/4, max].map((tick) => (
          <g key={tick}>
            <line x1="80" y1={y(tick)} x2="1580" y2={y(tick)} stroke="#e5e7eb" />
            <text x="62" y={y(tick)+7} textAnchor="end" fontSize="20" fill="#747b84">{tick.toFixed(1)}</text>
          </g>
        ))}
        <line x1="80" y1={y(0)} x2="1580" y2={y(0)} stroke="#EAB6A8" strokeWidth="3" strokeDasharray="10 8" />
        <line x1="80" y1={y(avg)} x2="1580" y2={y(avg)} stroke="#8b9299" strokeWidth="2" strokeDasharray="8 8" />
        {tenors.map((tenor) => <text key={tenor} x={x(tenor)} y="848" textAnchor="middle" fontSize="22" fontWeight="600" fill="#626a73">{tenor}</text>)}
        {filtered.map((row, index) => {
          const jitter = ((index % 5) - 2) * 13;
          const cy = y(row.spread || 0);
          const cx = x(row.tenor || "—") + jitter;
          const colors: Record<string, string> = { treasury: "#F4CDC3", cdb: "#FAE1CC", exim: "#E5E5E3", adbc: "#D6EFF5" };
          const notable = (row.spread || 0) > 0 || (row.spread || 0) <= -1;
          return <g key={`${row.bondCode}-${index}`}>
            <circle cx={cx} cy={cy} r={notable ? 13 : 9} fill={colors[tone(row.bondType || "")]} stroke="#8e776f" strokeWidth={notable ? 3 : 1} />
            {notable && <text x={cx + 17} y={cy - 12} fontSize="17" fontWeight="600" fill="#42464d">{row.shortName || row.bondCode} {(row.spread || 0).toFixed(2)}</text>}
          </g>;
        })}
        <text x="830" y="920" textAnchor="middle" fontSize="23" fontWeight="600" fill="#626a73">发行期限</text>
        <text x="25" y="540" transform="rotate(-90 25 540)" textAnchor="middle" fontSize="22" fontWeight="600" fill="#626a73">综收－二级（bp）</text>
        {[["国债","#F4CDC3"],["国开债","#FAE1CC"],["口行债","#E5E5E3"],["农发债","#D6EFF5"]].map(([label,color],i) => <g key={label} transform={`translate(${485+i*190},995)`}><circle r="11" fill={color}/><text x="22" y="7" fontSize="22" fontWeight="600" fill="#555b63">{label}</text></g>)}
      </svg>
      {!filtered.length && <div className="chart-empty">上传一二级表后，这里自动生成本周利差图</div>}
    </div>
  );
}

export default function Workbench() {
  const [weekStart, setWeekStart] = useState(() => mondayOf(new Date()));
  const [active, setActive] = useState<(typeof tabs)[number][0]>("overview");
  const [data, setData] = useState<WeekData>({ imports: [], records: [] });
  const [loading, setLoading] = useState(true);
  const [message, setMessage] = useState("");
  const [summary, setSummary] = useState("");
  const [saving, setSaving] = useState(false);
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
      const generated = spreadSummary(payload.records.filter(r => r.dataset_type === "spread").map(normalize));
      setSummary(payload.draft?.summary_text || generated);
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
        a.download = `本周国债政金债一二级利差_${weekStart}.png`; a.click();
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
        <nav>{tabs.map(([key, label]) => <button key={key} className={active === key ? "nav-active" : ""} onClick={() => setActive(key)}>{key === "overview" ? <CalendarDays/> : key === "local" ? <FileSpreadsheet/> : key === "spread" ? <BarChart3/> : <FileText/>}<span>{label}</span></button>)}</nav>
        <a className="legacy-link" href="https://superhanszhong.github.io/local-bond-daily-converter/" target="_blank" rel="noreferrer">原日表转换器 ↗</a>
        <div className="side-note"><Check size={16}/><span>数据按账号隔离<br/>周度草稿自动保存</span></div>
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
            <div className="card-head"><div><span className="section-label">每日入口</span><h2>上传并记录发行数据</h2></div><Upload/></div>
            <div className="upload-actions">
              <button onClick={() => localInput.current?.click()}><FileSpreadsheet/><span><strong>地方债发行表</strong><small>生成日表并纳入周汇总</small></span></button>
              <button onClick={() => spreadInput.current?.click()}><BarChart3/><span><strong>一二级分析表</strong><small>生成利差图与发行小结</small></span></button>
            </div>
            <input ref={localInput} hidden type="file" accept=".xlsx,.xlsm" onChange={e => e.target.files?.[0] && upload(e.target.files[0],"local_bond")}/>
            <input ref={spreadInput} hidden type="file" accept=".xlsx,.xlsm" onChange={e => e.target.files?.[0] && upload(e.target.files[0],"spread")}/>
            <p>同一天可重复上传；需要替换时先在“本周文件”中删除旧批次。</p>
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
          <div className="panel-head"><div><span className="section-label">PRIMARY / SECONDARY</span><h2>一二级利差可视化</h2></div><button className="secondary" onClick={downloadChart} disabled={!ordinary.length}><Download/>下载 PNG</button></div>
          <SpreadChart records={spreadRecords} svgRef={svgRef}/>
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
