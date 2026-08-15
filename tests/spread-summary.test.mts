import assert from "node:assert/strict";
import test from "node:test";
import { spreadSummary, type ParsedBondRecord } from "../app/lib/workbench.ts";

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
