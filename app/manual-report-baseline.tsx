"use client";

import { useState } from "react";
import { manualReportTotals, previousReportWeek, reportBaseline, type ReportSnapshot, type ReportTotals } from "./lib/report-baseline";
import { saveReportSnapshot } from "./lib/workbench-request";

export default function ManualReportBaseline({ weekStart, baseline, disabled, onSaved }: {
  weekStart: string; baseline: ReportTotals | null; disabled: boolean; onSaved: (value: ReportSnapshot) => void;
}) {
  const [open, setOpen] = useState(false);
  const [issuance, setIssuance] = useState("");
  const [net, setNet] = useState("");
  const [saving, setSaving] = useState(false);
  const [notice, setNotice] = useState("");
  return <div className="manual-baseline">
    <button type="button" className="secondary" disabled={disabled || saving} aria-expanded={open} onClick={() => {
      if (!open) { setIssuance(baseline ? String(baseline.rateIssuance) : ""); setNet(baseline?.rateNet == null ? "" : String(baseline.rateNet)); setNotice(""); }
      setOpen(!open);
    }}>{open ? "收起手动填写" : "手动填写／修改上周基准"}</button>
    {open && <form onSubmit={async event => {
      event.preventDefault(); setSaving(true); setNotice("");
      try {
        const totals = manualReportTotals(weekStart, issuance, net);
        const saved = await saveReportSnapshot(totals, null);
        const actual = reportBaseline(saved, weekStart);
        if (actual?.savedAt !== totals.savedAt) throw new Error("已有更新一期的周报，请切换至最近一期后填写上周基准");
        onSaved(saved); setNotice("上周基准已保存，生成本周周报时会自动使用。");
      } catch (error) { setNotice(error instanceof Error ? error.message : "保存失败，请重试"); }
      finally { setSaving(false); }
    }}>
      <p>对应上周：{previousReportWeek(weekStart)} 当周 · 单位：亿元。仅用于环比，不计入本周发行量或地方债年度累计。</p>
      <div className="manual-baseline-fields">
        <label>上周国债政金债发行量（必填）<input inputMode="decimal" required value={issuance} disabled={disabled || saving} onChange={event => setIssuance(event.target.value)} placeholder="按上周周报填写"/></label>
        <label>上周国债政金债净融资（选填）<input inputMode="decimal" value={net} disabled={disabled || saving} onChange={event => setNet(event.target.value)} placeholder="可为负数；未知请留空"/></label>
      </div>
      {baseline && <p>保存将更新已显示的上周比较基准。净融资请使用与本周一致的到期口径。</p>}
      <button type="submit" disabled={disabled || saving}>{saving ? "保存中…" : "保存上周基准"}</button>
    </form>}
    {notice && <p role="status">{notice}</p>}
  </div>;
}
