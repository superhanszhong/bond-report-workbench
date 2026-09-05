import assert from "node:assert/strict";
import test from "node:test";
import { consolidateBondRecords, mergeRecord, prepareRecordUpdate, recordKey, withImportHistory, type RecordPayload } from "../app/lib/record-merge.ts";
import { normalizeBondCode } from "../app/lib/bond-code.ts";

test("normalizes only missing SHCH policy-bank leading zeros, preserving reopenings", () => {
  assert.equal(normalizeBondCode(9260411), "09260411");
  assert.equal(normalizeBondCode("92603001z04.IB"), "092603001Z04");
  assert.equal(normalizeBondCode("2600006X"), "2600006X");
  assert.equal(normalizeBondCode("2671001"), "2671001");
  assert.equal(recordKey({ bondCode: "260308X2" }), recordKey({ bondCode: "260308Z2" }));
  assert.equal(recordKey({ bondCode: "260308X" }), recordKey({ bondCode: "260308Z01" }));
  assert.notEqual(recordKey({ bondCode: "260308X2" }), recordKey({ bondCode: "260308Z3" }));
  assert.notEqual(recordKey({ bondCode: "09260411" }), recordKey({ bondCode: "09260411Z01" }));
});

test("reconciles persisted leading-zero duplicates without mutating historical records", () => {
  const common = { trade_date: "2026-09-01", dataset_type: "spread", amount: 60, tenor: "1.112", raw_json: JSON.stringify({ __display: { 中标利率: "1.4200", 二级: "1.4184" }, __summaryMeta: { baseCode: "9260411", route: "中债招标" } }) };
  const original = [{ ...common, id: "a", bond_code: "09260411" }, { ...common, id: "b", bond_code: "9260411" }];
  const before = JSON.stringify(original);
  const rows = consolidateBondRecords(original);
  assert.equal(rows.length, 1);
  assert.equal(rows[0].bond_code, "09260411");
  assert.equal(JSON.parse(rows[0].raw_json).__summaryMeta.route, "上清所");
  assert.equal(JSON.stringify(original), before);
  assert.equal(consolidateBondRecords([...original, { ...common, bond_code: "09260411Z01" }, { ...common, trade_date: "2026-09-03", bond_code: "09260411" }]).length, 3);
});

test("duplicate economic conflicts block reporting rather than choosing a value", () => {
  const first = { tradeDate: "2026-09-01", bondCode: "09260411", amount: 60, raw: { __display: { 中标利率: "1.4200" } } };
  assert.throws(() => consolidateBondRecords([first, { ...first, bondCode: "9260411", amount: 70 }]), /发行量冲突/);
  assert.throws(() => consolidateBondRecords([first, { ...first, bondCode: "9260411", raw: { __display: { 中标利率: "1.4300" } } }]), /中标利率冲突/);
  assert.equal(consolidateBondRecords([first, { ...first, bondCode: "9260411", raw: { __display: { 中标利率: "1.42" } } }]).length, 1);
});

test("reupload corrects aliases while retaining complementary fields and not guessing unresolved conflicts", () => {
  const first = { trade_date: "2026-09-01", bond_code: "09260411", amount: 60, raw_json: JSON.stringify({ __display: { 中标利率: "1.42", 二级: "1.4184" } }) };
  const alias = { ...first, bond_code: "9260411", raw_json: JSON.stringify({ __display: { 中标利率: "1.43", 全场倍数: "3.13" } }) };
  assert.throws(() => prepareRecordUpdate({ tradeDate: "2026-09-01", bondCode: "09260411", amount: 60 }, [first, alias]), /中标利率冲突/);
  const update = prepareRecordUpdate({ tradeDate: "2026-09-01", bondCode: "09260411", raw: { __display: { 中标利率: "1.4200" } } }, [first, alias]);
  assert.deepEqual(update.record.raw?.__display, { 中标利率: "1.4200", 二级: "1.4184", 全场倍数: "3.13" });
  assert.equal(update.changed, true);
  const stored = { ...update.record };
  assert.equal(prepareRecordUpdate(update.record, [stored, stored]).changed, false);
});

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

test("revised imports retain the previous value and source without recording unchanged reuploads", () => {
  const first = { tradeDate: "2026-09-04", bondCode: "260308X3", amount: 120 };
  const stored = withImportHistory(first, undefined, "旧一二级.xlsx", "2026-09-04T08:00:00Z");
  const revised = prepareRecordUpdate({ ...first, bondCode: "260308Z03", amount: 140 }, [stored]);
  assert.equal(revised.changed, true);
  const traced = withImportHistory(revised.record, stored, "新一二级.xlsx", "2026-09-05T08:00:00Z");
  const history = traced.raw.__history;
  assert.equal(history.length, 1);
  assert.equal(history[0].previous.amount, 120);
  assert.equal(history[0].previous.raw.__source.fileName, "旧一二级.xlsx");
  assert.equal(traced.raw.__source.fileName, "新一二级.xlsx");
  assert.equal(traced.amount, 140);
  assert.equal(stored.amount, 120);
  assert.equal(prepareRecordUpdate({ ...first, bondCode: "260308Z03", amount: 140 }, [traced]).changed, false);
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
