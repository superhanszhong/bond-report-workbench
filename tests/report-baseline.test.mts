import assert from "node:assert/strict";
import test from "node:test";
import { readFile } from "node:fs/promises";
import JSZip from "jszip";
import { advanceReportSnapshot, manualReportTotals, reportBaseline, type ReportTotals } from "../app/lib/report-baseline.ts";
import { buildWeeklyReportBlob, reportTotals } from "../app/lib/report.ts";

const totals = (weekStart: string, rateIssuance = 100): ReportTotals => ({ weekStart, savedAt: "2026-09-11T00:00:00Z", rateIssuance, rateNet: rateIssuance - 10, localIssuance: 20, localNet: 5 });

test("manual first-week baseline permits negative net and unknown net without inventing local totals", () => {
  const manual = manualReportTotals("2026-09-07", "4,260", "-180.5");
  assert.equal(manual.weekStart, "2026-08-31");
  assert.equal(manual.rateIssuance, 4260);
  assert.equal(manual.rateNet, -180.5);
  assert.equal(manual.localIssuance, null);
  assert.equal(manualReportTotals("2026-09-07", "0", "").rateNet, null);
  assert.throws(() => manualReportTotals("2026-09-07", "", "0"), /发行量/);
  assert.throws(() => manualReportTotals("2026-09-07", "-1", "0"), /发行量/);
  assert.throws(() => manualReportTotals("2026-09-07", "100", "NaN"), /有效数字/);
  const first = advanceReportSnapshot(null, manual, null);
  const generated = advanceReportSnapshot(first, totals("2026-09-07", 5000), manual);
  assert.equal(reportBaseline(generated, "2026-09-07")?.rateIssuance, 4260);
  const revised = advanceReportSnapshot(generated, manualReportTotals("2026-09-07", "4300", "-140.5"), null);
  assert.equal(revised.current.rateIssuance, 5000);
  assert.equal(reportBaseline(revised, "2026-09-07")?.rateIssuance, 4300);
});

test("next week uses saved totals, and same-week regeneration retains the previous-week baseline", () => {
  const first = advanceReportSnapshot(null, totals("2026-08-31"), null);
  assert.equal(reportBaseline(first, "2026-09-07")?.rateIssuance, 100);
  const next = advanceReportSnapshot(first, totals("2026-09-07", 200), reportBaseline(first, "2026-09-07"));
  const reload = JSON.parse(JSON.stringify(next));
  const revised = advanceReportSnapshot(reload, totals("2026-09-07", 250), reportBaseline(reload, "2026-09-07"));
  assert.equal(revised.current.rateIssuance, 250);
  assert.equal(reportBaseline(revised, "2026-09-07")?.rateIssuance, 100);
  assert.equal(reportBaseline(revised, "2026-09-14")?.rateIssuance, 250);
  const third = advanceReportSnapshot(revised, totals("2026-09-14"), null);
  assert.equal(third.comparison?.weekStart, "2026-09-07");
  assert.doesNotMatch(JSON.stringify(third), /2026-08-31/);
});

test("missing or skipped weeks do not become zero; old report generation does not replace the latest", () => {
  assert.equal(reportBaseline(null, "2026-09-07"), null);
  const saved = advanceReportSnapshot(null, totals("2026-08-31", 0), null);
  assert.equal(reportBaseline(saved, "2026-09-07")?.rateIssuance, 0);
  assert.equal(reportBaseline(saved, "2026-09-14"), null);
  const skipped = advanceReportSnapshot(saved, totals("2026-09-14"), null);
  assert.equal(skipped.comparison, null);
  assert.deepEqual(advanceReportSnapshot(skipped, totals("2026-08-31"), null), skipped);
  const correctedPrior = advanceReportSnapshot(skipped, totals("2026-09-07", 300), null);
  assert.equal(correctedPrior.current.weekStart, "2026-09-14");
  assert.equal(reportBaseline(correctedPrior, "2026-09-14")?.rateIssuance, 300);
});

test("snapshot totals follow issuance authority, deduplicate aliases and distinguish missing maturity", () => {
  const row = { tradeDate: "2026-09-01", bondCode: "260308X2", amount: 120 };
  const input = { weekStart: "2026-08-31", summary: "", spreadRecords: [row, { ...row, bondCode: "260308Z2" }], localRecords: [] };
  assert.equal(reportTotals(input).rateIssuance, 120);
  assert.equal(reportTotals(input).rateNet, null);
  assert.equal(reportTotals({ ...input, maturity: { rateTotal: 0, rateBreakdown: "", localDaily: {}, localTotal: 0 } }).rateNet, 120);
});

test("Word uses the saved previous issuance and net even without previous detail rows", async () => {
  const template = await readFile("public/templates/weekly-bond-report-template.docx");
  const blob = await buildWeeklyReportBlob({ weekStart: "2026-09-07", summary: "", localRecords: [],
    spreadRecords: [{ tradeDate: "2026-09-08", bondCode: "260308X4", amount: 100, tenor: "1", bondType: "口行债" }],
    previousRateIssuance: 200, previousSpreadRecords: [],
    maturity: { rateTotal: 10, rateBreakdown: "进出:10亿", localDaily: {}, localTotal: 0, previousRateNet: 50 },
    templateBytes: template.buffer.slice(template.byteOffset, template.byteOffset + template.byteLength),
  });
  const xml = await (await JSZip.loadAsync(await blob.arrayBuffer())).file("word/document.xml")!.async("string");
  assert.match(xml, /较上周减少50.00%/);
  assert.match(xml, /净融资90亿/);
  assert.match(xml, /净融资较上周增加（上周到期口径测算净融资额50亿）/);
});
