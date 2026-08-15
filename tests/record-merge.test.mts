import assert from "node:assert/strict";
import test from "node:test";
import { mergeRecord } from "../app/lib/record-merge.ts";

test("keeps stored fields when a later upload leaves them blank", () => {
  const { record, changed } = mergeRecord({
    tradeDate: "2026-08-14", bondCode: "260016", amount: 900,
    raw: { __display: { 中标利率: "" } },
  }, {
    trade_date: "2026-08-14", bond_code: "260016", amount: 900,
    raw_json: JSON.stringify({ __display: { 中标利率: "1.6800", 二级: "1.6800", 全场倍数: "6.12" }, __summaryMeta: { route: "中债招标" } }),
  });
  assert.equal(changed, false);
  assert.deepEqual(record.raw?.__display, { 中标利率: "1.6800", 二级: "1.6800", 全场倍数: "6.12" });
});

test("updates a stored field when the later upload supplies a revised value", () => {
  const { record, changed } = mergeRecord({
    tradeDate: "2026-08-14", bondCode: "260016", amount: 900,
    raw: { __display: { 中标利率: "1.6810", 二级: "1.6800" } },
  }, {
    trade_date: "2026-08-14", bond_code: "260016", amount: 900,
    raw_json: JSON.stringify({ __display: { 中标利率: "1.6800", 二级: "1.6800" } }),
  });
  assert.equal(changed, true);
  assert.equal((record.raw?.__display as Record<string, unknown>).中标利率, "1.6810");
});
