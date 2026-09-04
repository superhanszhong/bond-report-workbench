import assert from "node:assert/strict";
import test from "node:test";
import { mergeRecord, recordKey, type RecordPayload } from "../app/lib/record-merge.ts";

test("local-bond incremental uploads preserve history, skip repeats and update revised amounts", () => {
  const stored = new Map<string, RecordPayload>();
  function upload(rows: RecordPayload[]) {
    const counts = { added: 0, updated: 0, unchanged: 0 };
    for (const row of rows) {
      const key = recordKey(row);
      const previous = stored.get(key);
      const merged = mergeRecord(row, previous ? { ...previous } : undefined);
      if (!merged.changed) counts.unchanged++;
      else if (previous) counts.updated++;
      else counts.added++;
      stored.set(key, merged.record);
    }
    return counts;
  }
  const first = { tradeDate: "2026-08-31", bondCode: "2671001", amount: 7.3, tenor: "7", region: "北京" };
  assert.deepEqual(upload([first]), { added: 1, updated: 0, unchanged: 0 });
  assert.deepEqual(upload([first]), { added: 0, updated: 0, unchanged: 1 });
  assert.deepEqual(upload([{ ...first, amount: 8 }, { ...first, tradeDate: "2026-09-04", bondCode: "2671002", amount: 10 }]), { added: 1, updated: 1, unchanged: 0 });
  assert.equal(stored.size, 2);
  assert.equal([...stored.values()].reduce((sum, row) => sum + row.amount!, 0), 18);
});

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
