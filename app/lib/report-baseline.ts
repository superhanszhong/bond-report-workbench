export type ReportTotals = {
  weekStart: string;
  savedAt: string;
  rateIssuance: number;
  rateNet: number | null;
  localIssuance: number | null;
  localNet: number | null;
  source?: "manual";
};

// One rolling record, with only the immediately preceding comparison retained
// so regenerating the current week's report never compares it against itself.
export type ReportSnapshot = { current: ReportTotals; comparison: ReportTotals | null };

export function previousReportWeek(weekStart: string) {
  const date = new Date(`${weekStart}T00:00:00Z`);
  date.setUTCDate(date.getUTCDate() - 7);
  return date.toISOString().slice(0, 10);
}

export function manualReportTotals(weekStart: string, issuance: string, net: string): ReportTotals {
  const parse = (value: string) => {
    const text = value.trim().replace(/[,，]/g, "");
    if (!text) return null;
    if (!/^[+-]?(?:\d+(?:\.\d*)?|\.\d+)$/.test(text) || !Number.isFinite(Number(text))) throw new Error("请填写有效数字，单位为亿元");
    return Number(text);
  };
  const rateIssuance = parse(issuance);
  const rateNet = parse(net);
  if (rateIssuance === null || rateIssuance < 0) throw new Error("请填写上周发行量，且不能为负数");
  return { weekStart: previousReportWeek(weekStart), savedAt: new Date().toISOString(), rateIssuance, rateNet,
    localIssuance: null, localNet: null, source: "manual" };
}

export function reportBaseline(snapshot: ReportSnapshot | null, weekStart: string): ReportTotals | null {
  const wanted = previousReportWeek(weekStart);
  return [snapshot?.current, snapshot?.comparison].find(row => row?.weekStart === wanted) || null;
}

export function advanceReportSnapshot(saved: ReportSnapshot | null, current: ReportTotals, comparison: ReportTotals | null): ReportSnapshot {
  if (![current.rateIssuance, current.localIssuance, current.rateNet, current.localNet].every(value => value === null || Number.isFinite(value))) throw new Error("周报比较数据无效，未保存");
  if (saved && current.weekStart < saved.current.weekStart) {
    // Revising last week updates its baseline without replacing a newer report.
    return current.weekStart === previousReportWeek(saved.current.weekStart)
      ? { ...saved, comparison: current } : saved;
  }
  const prior = reportBaseline(saved, current.weekStart);
  return { current, comparison: prior || (comparison?.weekStart === previousReportWeek(current.weekStart) ? comparison : null) };
}
