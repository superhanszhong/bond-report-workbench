import assert from "node:assert/strict";
import test from "node:test";
import XLSX from "xlsx-js-style";
import { basisPointValue, inferredRateType, isSpecialSpreadBond, maturityDailyTotals, parseLocalBondFile, parseIssuancePlanFile, parseMaturityFile, parseSpreadFile, rateMaturityBreakdown, resolveSpreadBp, spreadValueForMetric } from "../app/lib/workbench.ts";
import { mergeIssuanceSessions, planSession } from "../app/lib/report.ts";
import { mergeRecord } from "../app/lib/record-merge.ts";

test("spread import coalesces numeric and textual SHCH codes before building history and keeps explicit net price", async () => {
  const book = XLSX.utils.book_new();
  const header = ["发行日期", "代码", "期限", "发行人", "发行量", "综收比二级(bp)", "中标净价"];
  XLSX.utils.book_append_sheet(book, XLSX.utils.aoa_to_sheet([
    header, ["2026/09/01", 9260411, "1.112", "中国农业发展银行", 60, 0.16, null],
    ["2026/09/01", "09260411", "1.112", "中国农业发展银行", 60, 0.16, null],
    ["2026/09/01", "09260409Z21", "2", "中国农业发展银行", 60, null, 99.9767],
  ]), "历史");
  const rows = await parseSpreadFile(new File([XLSX.write(book, { type: "array", bookType: "xlsx" })], "一二级.xlsx"));
  assert.equal(rows.length, 2);
  assert.equal(rows[0].bondCode, "09260411");
  assert.equal(rows[0].summaryMeta?.route, "上清所");
  assert.equal(rows[0].summaryMeta?.previous, undefined);
  assert.equal(rows[1].raw?.中标净价, 99.9767);
});

test("local issuance import reads multiple detail sheets and removes duplicate date/code rows", async () => {
  const book = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(book, XLSX.utils.aoa_to_sheet([["地方债说明"]]), "说明");
  const header = ["发行起始日", "债券代码", "债券简称", "发行期限", "发行规模(亿元)", "债券全称", "所属区域"];
  const first = [new Date(2026, 7, 31, 12), "2671001.IB", "26北京47", "7Y", 7.3, "2026年北京市政府一般债券", "北京"];
  const second = ["2026/09/04", "2671002.SH", "26北京48", "10Y", 10, "2026年北京市政府专项债券", "北京"];
  XLSX.utils.book_append_sheet(book, XLSX.utils.aoa_to_sheet([header, first]), "周一");
  XLSX.utils.book_append_sheet(book, XLSX.utils.aoa_to_sheet([header, first, second]), "累计");
  const rows = await parseLocalBondFile(new File([XLSX.write(book, {type: "array", bookType: "xlsx"})], "地方债发行明细.xlsx"));
  assert.equal(rows.length, 2);
  assert.equal(rows[0].bondCode, "2671001");
  assert.equal(rows[1].bondCode, "2671002");
  assert.equal(rows[0].tenor, "7");
  assert.equal(rows[0].raw?.招标日, "2026-08-31");
  const persisted = { ...rows[0], raw: undefined, raw_json: JSON.stringify(rows[0].raw) };
  assert.equal(mergeRecord(rows[0], persisted).changed, false);
  assert.equal(rows.reduce((sum, row) => sum + row.amount!, 0), 17.3);
});

test("local issuance import rejects missing amounts instead of treating them as zero", async () => {
  const book = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(book, XLSX.utils.aoa_to_sheet([
    ["招标日", "债券代码", "债券简称", "期限", "发行量"],
    ["2026/09/04", "2671001", "26北京47", "7", null],
  ]), "地方债");
  await assert.rejects(() => parseLocalBondFile(new File([XLSX.write(book, {type: "array", bookType: "xlsx"})], "地方债.xlsx")), /发行量.*无效/);
});

test("local reopenings are retained while equivalent X/Z spellings are deduplicated", async () => {
  const makeFile = (amount = 30) => {
    const book = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(book, XLSX.utils.aoa_to_sheet([
      ["招标日", "债券代码", "债券简称", "期限", "发行量"],
      ["2026/09/01", "2671001", "26北京47", "7", 20],
      ["2026/09/04", "2671001X", "26北京47续发", "7", 30],
      ["2026/09/04", "2671001Z01.IB", "26北京47续发", "7", amount],
      ["2026/09/04", "2671001X2", "26北京47续发2", "7", 40],
    ]), "地方债");
    return new File([XLSX.write(book, { type: "array", bookType: "xlsx" })], "地方债.xlsx");
  };
  const rows = await parseLocalBondFile(makeFile());
  assert.equal(rows.length, 3);
  assert.equal(rows.reduce((sum, row) => sum + row.amount!, 0), 90);
  await assert.rejects(() => parseLocalBondFile(makeFile(31)), /发行量冲突/);
});

test("spread signs follow high/low descriptions and never interpret net prices as BP", () => {
  assert.equal(basisPointValue("0.07BP(较低二级(1.465))"), -0.07);
  assert.equal(basisPointValue("低2603001估值0.35BP"), -0.35);
  assert.equal(basisPointValue("0.16(较高估值)"), 0.16);
  assert.equal(basisPointValue("平二级(1.42)"), 0);
  assert.equal(basisPointValue("低二级净价(99.9775)0.0008元"), null);
  const dr = { tradeDate: "2026-09-01", bondCode: "09260409Z21", spread: 0.0008, raw: { 中标利率: 1.49, 二级: 99.9775, 招标利差: 0.0008 } };
  assert.equal(spreadValueForMetric(dr, "winning"), null);
  assert.equal(spreadValueForMetric(dr, "all_in"), null);
  assert.equal(spreadValueForMetric({ ...dr, bondCode: "09260409", raw: { 中标利率: "0.1%", 二级: "0.11%" } }, "winning"), -1);
  assert.equal(inferredRateType({ bondCode: "092603001Z04", remark: "DR浮息" }), "DR001浮息债");
});

test("one historical special note does not contaminate later ordinary reopenings", async () => {
  const book = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(book, XLSX.utils.aoa_to_sheet([
    ["发行日期", "代码", "期限", "发行人", "发行量", "综收比二级(bp)", "备注"],
    ["2026/07/01", "260403Z17", "3", "中国农业发展银行", 70, -1, "LPR浮息债"],
    ["2026/08/17", "260403Z26", "3", "中国农业发展银行", 70, -1, "京津冀协同发展主题债"],
    ["2026/08/31", "260403Z30", "3", "中国农业发展银行", 70, -1, ""],
    ["2026/08/31", "250409Z35", "3", "中国农业发展银行", 20, -1, ""],
  ]), "历史");
  const rows = await parseSpreadFile(new File([XLSX.write(book, { type: "array", bookType: "xlsx" })], "一二级.xlsx"));
  assert.equal(isSpecialSpreadBond(rows[1]), true);
  assert.equal(isSpecialSpreadBond(rows[2]), false);
  assert.equal(rows[2].summaryMeta?.rateType, "固息或贴现");
  assert.equal(rows[3].summaryMeta?.rateType, "LPR浮息债");
});

test("uses the supplied spread when it reconciles with all-in and secondary yields", () => {
  const result = resolveSpreadBp({
    provided: -0.18, allIn: 0.015532, secondary: 0.01555,
    winningRate: 0.015356, issuer: "中国农业发展银行", remark: "",
  });
  assert.equal(result.spread, -0.18);
  assert.equal(result.audit.status, "provided");
});

test("recalculates a mismatched spread from all-in and secondary yields", () => {
  const result = resolveSpreadBp({
    provided: 18, allIn: 0.015532, secondary: 0.01555,
    winningRate: 0.015356, issuer: "中国农业发展银行", remark: "",
  });
  assert.equal(result.spread, -0.18);
  assert.equal(result.audit.status, "recalculated");
});

test("fills a missing Treasury spread from winning and secondary yields", () => {
  const result = resolveSpreadBp({
    provided: "", allIn: "", secondary: 0.01388,
    winningRate: 0.013776, issuer: "中华人民共和国财政部", remark: "追加10.1亿",
  });
  assert.equal(result.spread, -1.04);
  assert.equal(result.audit.status, "derived_treasury");
});

test("parses the supplied maturity workbook layout and classifies rate versus local bonds", async () => {
  const sheet = XLSX.utils.aoa_to_sheet([
    ["发行状态", "债券简称", "债券代码", "发行规模", "期限", "实际到期日", "发行人", "所属区域"],
    ["已到期", "26贴现国债32", "269932.IB", 300, "91D", "20260820", "中华人民共和国财政部", "北京"],
    ["已到期", "24北京53", "809173.IB", 2, "2Y", "20260821", "北京市人民政府", "北京"],
  ]);
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, sheet, "按债券");
  const bytes = XLSX.write(workbook, { type: "array", bookType: "xlsx" });
  const records = await parseMaturityFile(new File([bytes], "到期明细.xlsx"));
  assert.equal(records.length, 2);
  assert.equal(records[0].bondType, "国债");
  assert.equal(records[1].bondType, "地方债");
  assert.equal(records[1].amount, 2);
});

test("parses the new-bond workbook schedule and keeps only Treasury and policy-bank issues", async () => {
  const sheet = XLSX.utils.aoa_to_sheet([
    ["债券简称", "债券代码", "发行期限", "计划发行量(亿)", "实际发行量(亿)", "招标时间", "招标标的", "发行起始日", "托管机构"],
    ["26农发03(增发26)", "260403X26.IB", "3Y", 80, 80, "14:00-15:00", "价格招标", "2026/08/17", "中债登"],
    ["26进出清发001(增发2)", "092603001Z02.IB", "1Y", 20, 20, "10:00-16:30", "报价发行", "2026/08/19", "上清所"],
    ["26国开清发02(增发13)", "09260202Z13.IB", "2Y", 30, 30, "10:00-11:00", "利率招标", "2026/08/20", "上清所"],
    ["26北京债47", "2671001.IB", "7Y", 7.3, 7.3, "09:30-10:10", "利率招标", "2026/08/17", "中债登"],
  ]);
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, sheet, "利率债");
  const bytes = XLSX.write(workbook, { type: "array", bookType: "xlsx" });
  const records = await parseIssuancePlanFile(new File([bytes], "新债发行.xlsx"));
  assert.equal(records.length, 3);
  assert.equal(records[0].bidTime, "14:00-15:00");
  assert.equal(planSession(records[0]), "下午");
  assert.equal(planSession(records[1]), "上午");
  assert.equal(records[1].issuanceRoute, "报价发行");
  assert.equal(records[2].issuanceRoute, "上清所");
  assert.equal(records[2].remark, "上清所");
});

test("uses the primary-secondary table for issuance amounts and the schedule only for session", () => {
  const spread = [{ tradeDate: "2026-08-17", bondCode: "260403X26", bondType: "农发债", tenor: "3", amount: 80 }];
  const schedules = [
    { tradeDate: "2026-08-17", bondCode: "260403X26", bondType: "农发债", tenor: "3", amount: 999, bidTime: "14:00-15:00" },
    { tradeDate: "2026-08-17", bondCode: "260999X1", bondType: "农发债", tenor: "5", amount: 500, bidTime: "10:00-11:00" },
  ];
  const merged = mergeIssuanceSessions(spread, schedules);
  assert.equal(merged.length, 1);
  assert.equal(merged[0].amount, 80);
  assert.equal(merged[0].bidTime, "14:00-15:00");
  assert.equal(planSession(merged[0]), "下午");
});

test("aggregates local-bond maturities by the actual maturity date from the workbook", () => {
  const totals = maturityDailyTotals([
    { tradeDate: "2026-08-17", bondType: "地方债", amount: 10.25 },
    { tradeDate: "2026-08-17", bondType: "地方债", amount: 2.75 },
    { tradeDate: "2026-08-19", bondType: "地方债", amount: 6 },
  ], "2026-08-17");
  assert.equal(totals["2026-08-17"], 13);
  assert.equal(totals["2026-08-18"], 0);
  assert.equal(totals["2026-08-19"], 6);
});

test("refines Treasury and policy-bank maturity categories", () => {
  const text = rateMaturityBreakdown([
    { tradeDate: "2026-08-20", bondType: "国债", shortName: "26贴现国债32", bondCode: "269932", amount: 300 },
    { tradeDate: "2026-08-20", bondType: "国债", shortName: "23附息国债17", bondCode: "230017", amount: 3055.1 },
    { tradeDate: "2026-08-21", bondType: "口行债", shortName: "26进出清发贴现02", bondCode: "092603002", amount: 80 },
  ]);
  assert.equal(text, "贴现国债:300亿　附息国债:3055.1亿　进出清发:80亿");
});
