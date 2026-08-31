import assert from "node:assert/strict";
import test from "node:test";
import { rollingSpreadAnalysis, spreadSummary, type ParsedBondRecord } from "../app/lib/workbench.ts";

const meta = (previous?: ParsedBondRecord["summaryMeta"]["previous"]) => ({
  baseCode: "2600006", route: "中债招标", rateType: "固息或贴现",
  displaySpreadText: "较二级0.50BP", auctionSpreadText: "", allInText: "", secondaryText: "", winningRateText: "", note: "", proceeds: "", previous,
});

test("does not write a spread-change conclusion for a new fixed-rate bond", () => {
  const text = spreadSummary([{ tradeDate: "2026-08-14", bondCode: "260016", bondType: "国债", tenor: "10", spread: 0.5, summaryMeta: meta() }]);
  assert.equal(text, "");
});

test("compares an old bond with the prior same-bond issuance", () => {
  const previous = { date: "2026-07-14", code: "2600006", comparisonType: "same_bond" as const, displaySpreadText: "较二级0.20BP", auctionSpreadText: "", allInText: "", secondaryText: "", note: "", spread: 0.2 };
  const text = spreadSummary([{ tradeDate: "2026-08-14", bondCode: "2600006X1", bondType: "国债", tenor: "30", spread: -0.6, summaryMeta: meta(previous) }]);
  assert.match(text, /上次发行2600006/);
  assert.match(text, /由正转负/);
});

test("compares a discount bond with the prior compatible discount bond", () => {
  const previous = { date: "2026-08-05", code: "269940", comparisonType: "discount_comparator" as const, displaySpreadText: "较二级-0.20BP", auctionSpreadText: "", allInText: "", secondaryText: "", note: "", spread: -0.2 };
  const text = spreadSummary([{ tradeDate: "2026-08-12", bondCode: "269949", bondType: "国债", tenor: "63D", spread: -0.8, summaryMeta: meta(previous) }]);
  assert.match(text, /对比同期限贴现国债269940/);
  assert.match(text, /负利差走扩/);
});

test("finds the prior same-bond issuance from separately loaded history", () => {
  const current = { tradeDate: "2026-08-14", bondCode: "2600006X2", bondType: "国债", tenor: "30", spread: -0.6, summaryMeta: meta() };
  const history = [{ tradeDate: "2026-07-14", bondCode: "2600006X1", bondType: "国债", tenor: "30", spread: 0.2, summaryMeta: meta() }];
  const text = spreadSummary([current], history);
  assert.match(text, /上次发行2600006X1/);
  assert.match(text, /由正转负0\.80BP/);
});

test("summarizes this week's comparable spread against last week's", () => {
  const previous = { date: "2026-07-14", code: "2600006X1", comparisonType: "same_bond" as const, displaySpreadText: "较二级0.20BP", auctionSpreadText: "", allInText: "", secondaryText: "", note: "", spread: 0.2 };
  const current = { tradeDate: "2026-08-14", bondCode: "2600006X2", bondType: "国债", tenor: "30", spread: -0.6, summaryMeta: meta(previous) };
  const lastWeek = { tradeDate: "2026-08-07", bondCode: "2600005X1", bondType: "国债", tenor: "10", spread: 0.4, summaryMeta: { ...meta(), baseCode: "2600005" } };
  const text = spreadSummary([current], [lastWeek]);
  assert.match(text, /本周国债平均一二级利差为-0\.60BP，上周为\+0\.40BP/);
  assert.match(text, /整体由正转负1\.00BP/);
});

test("compares same issuer and tenor against a rolling four-week average", () => {
  const history: ParsedBondRecord[] = ["2026-08-03", "2026-08-10", "2026-08-17", "2026-08-24"].flatMap((tradeDate, index) => [
    { tradeDate, bondCode: `260405X${index + 1}`, shortName: "26农发05", bondType: "农发债", tenor: "3", spread: -2 },
    { tradeDate, bondCode: `260210X${index + 1}`, shortName: "26国开10", bondType: "国开债", tenor: "5", spread: -1 },
  ]);
  const current: ParsedBondRecord[] = [
    { tradeDate: "2026-08-31", bondCode: "260405X5", shortName: "26农发05", bondType: "农发债", tenor: "3", spread: -4 },
    { tradeDate: "2026-08-31", bondCode: "260210X5", shortName: "26国开10", bondType: "国开债", tenor: "5", spread: -0.2 },
    { tradeDate: "2026-08-31", bondCode: "250409X35", shortName: "25农发09", bondType: "农发债", tenor: "3", spread: -3.8, remark: "LPR浮息债" },
  ];
  const result = rollingSpreadAnalysis(current, [...history, ...current], "2026-08-31", "2026-09-04");
  assert.equal(result.comparableGroups, 2);
  assert.equal(result.normalGroups, 1);
  assert.equal(result.notableGroups, 1);
  assert.equal(result.specialBonds, 1);
  assert.match(result.text, /农发债3Y平均利差由前四周-2\.00BP变为本期-4\.00BP，下行2\.00BP/);
  assert.match(result.text, /国开债5Y\+?-0\.20BP|国开债5Y-0\.20BP/);
  assert.match(result.text, /特殊债券发行汇总.*25农发09/);
});
