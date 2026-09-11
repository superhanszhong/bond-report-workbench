import assert from "node:assert/strict";
import test from "node:test";
import { loadSpreadHistory } from "../app/lib/workbench-request.ts";
import { createPolicyCommentDrafts, policyDraftResults } from "../app/lib/policy-comment.ts";
import type { ParsedBondRecord } from "../app/lib/workbench.ts";

function bond(date: string, code: string, spread: string): ParsedBondRecord {
  return { tradeDate: date, bondCode: code, bondType: "农发债", tenor: "3", summaryMeta: {
    baseCode: "260403", rateType: "固息或贴现", route: "中债招标", displaySpreadText: spread,
    auctionSpreadText: spread, allInText: "", secondaryText: "1.5", winningRateText: "1.48", note: "", proceeds: "",
  } };
}

test("comment history reload includes saved current-week data and is independent of the selected report week", async () => {
  const rows = [bond("2026-09-04", "260403X31", "-1"), bond("2026-09-10", "260403X32", "-2")];
  const records = await loadSpreadHistory<ParsedBondRecord>(async input => {
    const url = new URL(input, "http://localhost");
    assert.equal(url.searchParams.get("datasetType"), "spread");
    assert.equal(url.searchParams.has("weekStart"), false);
    return Response.json({ records: rows.filter(row => row.tradeDate >= url.searchParams.get("startDate")! && row.tradeDate <= url.searchParams.get("endDate")!) });
  });
  assert.deepEqual(records, rows);
  const plan = bond("2026-09-11", "260403X33", "");
  const draft = { ...createPolicyCommentDrafts([plan], records)[0], finalValue: "1.49", benchmarkValue: "1.5" };
  const result = policyDraftResults([draft], records)[0].comment!;
  assert.equal(result.previousCode, "260403X32");
  assert.equal(result.previousSpread, -2);
});

test("saved corrections replace prior history without erasing manually filled quotes", async () => {
  const plan = bond("2026-09-11", "260403X33", "");
  const original = [bond("2026-09-10", "260403X32", "-2")];
  const entered = { ...createPolicyCommentDrafts([plan], original)[0], finalValue: "1.49", benchmarkValue: "1.5", referenceBond: "260405" };
  const refreshed = await loadSpreadHistory<ParsedBondRecord>(async () => Response.json({ records: [bond("2026-09-10", "260403Z32", "-0.5")] }));
  const drafts = createPolicyCommentDrafts([plan], refreshed, [entered]);
  assert.equal(drafts[0].finalValue, "1.49");
  assert.equal(drafts[0].benchmarkValue, "1.5");
  assert.equal(drafts[0].referenceBond, "260405");
  assert.equal(policyDraftResults(drafts, refreshed)[0].comment!.previousSpread, -0.5);
});

test("history load failures are surfaced instead of claiming the latest date was synchronized", async () => {
  await assert.rejects(() => loadSpreadHistory(async () => Response.json({ error: "读取失败" }, { status: 500 })), /读取失败/);
  await assert.rejects(() => loadSpreadHistory(async () => Response.json({ latestDates: { spread: "2026-09-11" } })), /读取首页一二级历史库失败/);
  assert.deepEqual(await loadSpreadHistory(async () => Response.json({ records: [] })), []);
});
