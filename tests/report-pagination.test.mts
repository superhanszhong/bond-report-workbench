import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import test from "node:test";
import JSZip from "jszip";
import { DOMParser } from "@xmldom/xmldom";
import { buildWeeklyReportBlob, mergeIssuanceSessions, reportDataWarnings, sortReportBonds } from "../app/lib/report.ts";
import type { ParsedBondRecord } from "../app/lib/workbench.ts";

async function reportXml(spreadRecords: ParsedBondRecord[], localRecords: ParsedBondRecord[] = [], previousSpreadRecords: ParsedBondRecord[] = [], ytdLocalRecords = localRecords) {
  const template = await readFile("public/templates/weekly-bond-report-template.docx");
  const blob = await buildWeeklyReportBlob({ weekStart: "2026-08-31", summary: "", spreadRecords, localRecords, ytdLocalRecords, previousSpreadRecords,
    templateBytes: template.buffer.slice(template.byteOffset, template.byteOffset + template.byteLength) as ArrayBuffer,
    maturity: { rateTotal: 4080, rateBreakdown: "农发:4080亿", localDaily: {}, localTotal: 0 } });
  return (await JSZip.loadAsync(await blob.arrayBuffer())).file("word/document.xml")!.async("string");
}

test("report deduplication recomputes issue counts, daily/weekly totals, net financing and prior-week change", async () => {
  const bond = { tradeDate: "2026-09-01", bondType: "农发债", bondCode: "09260411", tenor: "1.112", amount: 60 };
  const xml = await reportXml([bond, { ...bond, bondCode: "9260411" }, { ...bond, bondCode: "260308Z2", bondType: "口行债", amount: 120 }], [],
    [{ ...bond, tradeDate: "2026-08-25", amount: 200 }, { ...bond, tradeDate: "2026-08-25", bondCode: "9260411", amount: 200 }]);
  assert.match(xml, /本周共发2期利率债/);
  assert.match(xml, /发行总额达180亿/);
  assert.match(xml, /较上周减少10\.00%/);
  assert.match(xml, /净融资-3900亿/);
  assert.equal((xml.match(/>1\.112Y:60亿</g) || []).length, 1);
  assert.doesNotMatch(xml, />9260411</);
});

test("0831-0904 supplied report regression removes exactly the duplicate 60-亿元 entry", async () => {
  // Amounts are in 亿元, transcribed from the supplied report; no market values are changed.
  const days: [string, string, number, string?][][] = [
    [["250409Z35", "3", 20], ["260403Z30", "3", 70], ["260405Z14", "5", 50], ["260410Z14", "10", 90]],
    [["09260402Z11", "2", 30], ["09260407Z07", "7", 20], ["09260409Z21", "2", 60], ["09260411", "1.112", 60], ["260308Z2", "1", 120, "口行债"], ["9260411", "1.112", 60]],
    [["092603001Z04", "1", 20, "口行债"], ["092603002", "2", 50, "口行债"], ["2600006X", "30", 730, "国债"], ["260403Z31", "3", 70], ["260405Z15", "5", 50], ["260410Z15", "10", 90], ["269955", "182D", 600, "国债"], ["269956", "91D", 350, "国债"]],
    [["09260403Z07", "3", 10], ["09260405Z06", "5", 10], ["09260410Z05", "10", 10], ["09260411Z01", "1.112", 70]],
    [["2600005X2", "20", 240, "国债"], ["260014X2", "1", 1300, "国债"], ["260308Z3", "1", 140, "口行债"]],
  ];
  const dates = ["2026-08-31", "2026-09-01", "2026-09-02", "2026-09-03", "2026-09-04"];
  const records = days.flatMap((rows, day) => rows.map(([bondCode, tenor, amount, bondType]) => ({ tradeDate: dates[day], bondCode, tenor, amount, bondType: bondType || "农发债" })));
  const xml = await reportXml(records);
  assert.match(xml, /本周共发24期利率债/);
  assert.match(xml, /发行总额达4260亿/);
  assert.match(xml, /净融资180亿/);
  assert.match(xml, /按到期明细测算，净融资180亿/);
  assert.doesNotMatch(xml, /正式净融资需以总偿还口径核对/);
  assert.match(xml, /农发:710亿/);
  assert.match(xml, />290</);
  assert.equal((xml.match(/>1\.112Y:60亿</g) || []).length, 1);
});

test("local totals include all five categories for both weekly and year-to-date columns", async () => {
  const rows = [
    ["新增", "一般", 334.1997], ["再融资", "一般", 318.3583], ["新增", "专项", 619.5839], ["再融资", "专项", 194.5952], ["置换", "专项", 93.5186],
  ].map(([nature, kind, amount], index) => ({ tradeDate: "2026-08-31", bondCode: `267100${index}`, amount: Number(amount), raw: { 性质: nature, 类型: kind } }));
  const ytd = [...rows, { ...rows[1], tradeDate: "2026-08-24", amount: 200 }];
  const xml = await reportXml([], rows, [], ytd);
  assert.match(xml, /本周地方债发行1560\.2557亿/);
  assert.doesNotMatch(xml, /不含普通再融资债/);
  const document = new DOMParser().parseFromString(xml, "application/xml");
  const table = document.getElementsByTagName("w:tbl").item(0)!;
  const values = Array.from(table.getElementsByTagName("w:tr")).map(row =>
    Array.from(row.getElementsByTagName("w:tc")).map(cell => cell.textContent));
  assert.deepEqual(values[6], ["合计", "1560.2557", "1760.2557", "/", "/", "/"]);
  for (const column of [1, 2]) {
    const sum = values.slice(1, 6).reduce((total, row) => total + Number(row[column]), 0);
    assert.ok(Math.abs(sum - Number(values[6][column])) < 0.00001);
  }
});

test("DR review follows the edited reference: net-price result only and percentage valuation", async () => {
  const bond: ParsedBondRecord = { tradeDate: "2026-09-01", bondCode: "09260409Z21", tenor: "2", bondType: "农发债", amount: 60, remark: "DR浮息，直投招标",
    raw: { __display: { 中标利率: "1.4930", 中标净价: "99.9767", 前一日估值: "1.4987", 二级: "99.9775元", 全场倍数: "4.82", 边际倍数: "2.75" } } };
  const xml = await reportXml([bond]);
  assert.match(xml, /中标结果/);
  assert.doesNotMatch(xml, /1\.4930/);
  assert.match(xml, /净价 99\.9767元/);
  assert.match(xml, />1\.4987%<\/w:t>/);
  assert.match(xml, /净价 99\.9775元/);
  assert.match(xml, /今日直投招标发行1只农发清发债/);
  assert.equal(bond.raw?.__display && (bond.raw.__display as Record<string, string>).中标利率, "1.4930");
  const missing = { ...bond, raw: { __display: { 中标利率: "1.4930", 综收: "1.49" } } };
  assert.match(await reportXml([missing]), /净价未提供/);
  const legacy = { ...bond, raw: { __display: { 中标利率: "1.4930", 综收: "99.9767" } } };
  assert.match(await reportXml([legacy]), /净价 99\.9767元/);
  assert.doesNotMatch(await reportXml([{ ...bond, bondCode: "09260409" }]), /净价 99\.9767元/);
  const quote = { ...bond, remark: "DR浮息，直投招标，前台报价发行" };
  assert.match(await reportXml([quote]), /今日直投报价发行1只农发清发债/);
});

test("report sorting groups varieties and compares D/M/Y numerically without mutating inputs", () => {
  const rows = [
    ["A", "农发债", "10Y"], ["B", "口行债", "2"], ["C", "农发债", "1.112"],
    ["D", "国债", "182D"], ["E", "农发债", "3年"], ["F", "国债", "91D"],
    ["G", "口行债", "12M"], ["H", "国债", "30Y"], ["I", "农发债", "3Y浮息"],
    ["J", "农发债", "待定"], ["K", "国债", "1Y"],
  ].map(([bondCode, bondType, tenor]) => ({ tradeDate: "2026-09-02", bondCode, bondType, tenor }));
  const before = structuredClone(rows);
  assert.deepEqual(sortReportBonds(rows).map(row => row.bondCode), ["C", "E", "I", "A", "J", "G", "B", "F", "D", "K", "H"]);
  assert.deepEqual(rows, before);
});

test("weekly schedule cells and daily review rows both sort same-variety tenors short to long", async () => {
  const rows = [
    ["09260410Z05", "农发债", "10Y", 10], ["09260403Z07", "农发债", "3Y", 10],
    ["09260405Z06", "农发债", "5Y", 10], ["09260411Z01", "农发债", "1.112Y", 70],
    ["269955", "国债", "182D", 600], ["269956", "国债", "91D", 350],
  ].map(([bondCode, bondType, tenor, amount]) => ({ tradeDate: "2026-09-03", bondCode: String(bondCode), bondType: String(bondType), tenor: String(tenor), amount: Number(amount), bidTime: "10:00" }));
  const xml = await reportXml(rows);
  const document = new DOMParser().parseFromString(xml, "application/xml");
  const tables = document.getElementsByTagName("w:tbl");
  const scheduleCell = tables.item(1)!.getElementsByTagName("w:tr").item(1)!.getElementsByTagName("w:tc").item(4)!;
  const scheduleText = Array.from(scheduleCell.getElementsByTagName("w:t")).map(node => node.textContent).join("\n");
  assert.match(scheduleText, /农发\n1\.112Y:70亿\n3Y:10亿\n5Y:10亿\n10Y:10亿/);
  assert.match(scheduleText, /贴现国债\n91D:350亿\n182D:600亿/);
  const review = Array.from(tables.item(5)!.getElementsByTagName("w:tr")).slice(1).map(row =>
    Array.from(row.getElementsByTagName("w:tc")).map(cell => cell.textContent));
  assert.deepEqual(review.map(row => row[0]), ["09260411Z01", "09260403Z07", "09260405Z06", "09260410Z05", "269956", "269955"]);
  assert.deepEqual(review.map(row => row[2]), ["70", "10", "10", "10", "350", "600"]);
  assert.match(xml, /发行总额达1050亿/);
});

test("report warnings accept X/Z aliases of the same reopening but flag different periods", () => {
  const row = { tradeDate: "2026-09-01", bondCode: "260308Z2" };
  assert.deepEqual(reportDataWarnings([row], [{ ...row, bondCode: "260308X2" }]), []);
  assert.deepEqual(reportDataWarnings([{ ...row, bondCode: "260308X" }], [{ ...row, bondCode: "260308X1" }]), []);
  assert.deepEqual(reportDataWarnings([{ ...row, bondCode: "260308X" }], [{ ...row, bondCode: "260308Z01" }]), []);
  assert.deepEqual(reportDataWarnings([{ ...row, bondCode: "260308Z" }], [{ ...row, bondCode: "260308X1" }]), []);
  assert.match(reportDataWarnings([row], [{ ...row, bondCode: "260308X3" }]).join(""), /增发期次待核对/);
  assert.match(reportDataWarnings([{ ...row, bondCode: "260308" }], [{ ...row, bondCode: "260308X" }]).join(""), /增发期次待核对/);
  assert.equal(row.bondCode, "260308Z2");
  assert.match(reportDataWarnings([{ ...row, bondCode: "09260409Z21", remark: "DR浮息" }], []).join(""), /缺少中标净价/);
});

test("plans flag upsize and special syndicates but never change authoritative amounts or add issuances", async () => {
  const actual = { tradeDate: "2026-09-04", bondCode: "260308Z3", amount: 140, tenor: "1", bondType: "口行债", raw: { 中标利率: 1.2 } };
  const plans = [
    { ...actual, bondCode: "260308X3", amount: 120, bidTime: "10:00", raw: { 加权利率: 1.19 } },
    { tradeDate: "2026-09-03", bondCode: "092602002", amount: 50, tenor: "3", bondType: "国开债", shortName: "绿色小团债" },
  ];
  const merged = mergeIssuanceSessions([actual], plans);
  assert.equal(merged.length, 1);
  assert.equal(merged[0].amount, 140);
  assert.equal(merged[0].raw?.中标利率, 1.2);
  const messages = reportDataWarnings([actual], plans).join("\n");
  assert.match(messages, /上弹/);
  assert.match(messages, /发行计划有 50 亿/);
  assert.match(messages, /本周不纳入/);
  assert.match(messages, /中标利率 1.2%/);
  const xml = await reportXml(merged);
  assert.match(xml, /本周共发1期利率债/);
  assert.match(xml, /发行总额达140亿/);
  assert.doesNotMatch(xml, /092602002|由你核定|待核对/);
  assert.equal(actual.amount, 140);
});

test("report rejects a current local bond missing or inconsistent in the annual detail", async () => {
  const current = { tradeDate: "2026-08-31", bondCode: "2671001X", amount: 20, raw: { 性质: "新增", 类型: "一般" } };
  await assert.rejects(() => reportXml([], [current], [], []), /本周与年度明细不一致/);
  await assert.rejects(() => reportXml([], [current], [], [{ ...current, amount: 21 }]), /本周与年度明细不一致/);
  await assert.rejects(() => reportXml([], [current], [], [{ ...current, raw: { 性质: "再融资", 类型: "一般" } }]), /本周与年度明细不一致/);
  assert.match(await reportXml([], [current], [], [{ ...current, bondCode: "2671001Z01" }]), /本周地方债发行20亿/);
});

test("weekly report includes local issuance totals and net financing from uploaded records", async () => {
  const template = await readFile("public/templates/weekly-bond-report-template.docx");
  const localRecords: ParsedBondRecord[] = [
    { tradeDate: "2026-08-31", bondCode: "2671001", shortName: "26北京47", bondType: "一般债", tenor: "7", amount: 7.3 },
    { tradeDate: "2026-09-04", bondCode: "2671002", shortName: "26北京48", bondType: "专项债", tenor: "10", amount: 10 },
  ];
  const blob = await buildWeeklyReportBlob({
    weekStart: "2026-08-31", summary: "", localRecords, spreadRecords: [],
    templateBytes: template.buffer.slice(template.byteOffset, template.byteOffset + template.byteLength) as ArrayBuffer,
    maturity: { rateTotal: 0, rateBreakdown: "", localDaily: { "2026-08-31": 2 }, localTotal: 2 },
  });
  const zip = await JSZip.loadAsync(await blob.arrayBuffer());
  const xml = await zip.file("word/document.xml")!.async("string");
  assert.match(xml, /本周地方债发行17\.3亿，按到期明细测算，净融资额15\.3亿/);
  assert.match(xml, /地方债/);
  assert.match(xml, />7\.3<\/w:t>/);
});

test("daily reviews flow continuously without forced page breaks", async () => {
  const template = await readFile("public/templates/weekly-bond-report-template.docx");
  const templateBytes = template.buffer.slice(template.byteOffset, template.byteOffset + template.byteLength) as ArrayBuffer;
  const dates = ["2026-08-17", "2026-08-18", "2026-08-19", "2026-08-20", "2026-08-21"];
  const spreadRecords: ParsedBondRecord[] = dates.flatMap((tradeDate, dayIndex) =>
    Array.from({ length: 4 }, (_, rowIndex) => ({
      tradeDate,
      bondCode: `26${dayIndex}${rowIndex}00X1`,
      bondType: "国债",
      tenor: "10",
      amount: 100,
      spread: -0.5,
      raw: { __display: { 中标利率: "1.50", 前一日估值: "1.51", 二级: "1.51", 全场倍数: "3.20", 边际倍数: "1.10" } },
    })),
  );
  const blob = await buildWeeklyReportBlob({
    weekStart: dates[0], summary: "", localRecords: [], spreadRecords, templateBytes,
  });
  const zip = await JSZip.loadAsync(await blob.arrayBuffer());
  const xml = await zip.file("word/document.xml")!.async("string");

  assert.doesNotMatch(xml, /<w:br\b[^>]*w:type="page"/);
  assert.doesNotMatch(xml, /<w:pageBreakBefore\b/);
  assert.match(xml, /<w:cantSplit\b/);
  assert.match(xml, /<w:tblHeader\b/);
  assert.match(xml, /<w:keepNext\b/);
});

test("rate financing paragraph continues from the verified prior-week baseline", async () => {
  const template = await readFile("public/templates/weekly-bond-report-template.docx");
  const templateBytes = template.buffer.slice(template.byteOffset, template.byteOffset + template.byteLength) as ArrayBuffer;
  const spreadRecords: ParsedBondRecord[] = [{
    tradeDate: "2026-08-17", bondCode: "260017", bondType: "国债", tenor: "10", amount: 100,
  }];
  const blob = await buildWeeklyReportBlob({
    weekStart: "2026-08-17", summary: "", localRecords: [], spreadRecords, templateBytes,
    maturity: {
      rateTotal: 200,
      rateBreakdown: "国债:200亿",
      localDaily: {},
      // 即使旧入库明细计算出了不一致值，也优先使用已核定的 8/10–8/14 口径。
      previousRateNet: -1638,
    },
  });
  const zip = await JSZip.loadAsync(await blob.arrayBuffer());
  const xml = await zip.file("word/document.xml")!.async("string");

  assert.match(xml, /本周国债政金债到期200亿（不包含凭证式国债）；按到期明细测算，净融资-100亿/);
  assert.match(xml, /上周净融资额-1643亿为已核定总偿还口径，不作直接比较/);
  assert.doesNotMatch(xml, /净融资较上周增加/);
});
