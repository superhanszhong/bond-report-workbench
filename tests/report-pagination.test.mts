import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import test from "node:test";
import JSZip from "jszip";
import { buildWeeklyReportBlob } from "../app/lib/report.ts";
import type { ParsedBondRecord } from "../app/lib/workbench.ts";

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

  assert.match(xml, /本周国债政金债偿还200亿（不包含凭证式国债）；净融资-100亿，净融资较上周增加（上周净融资额-1643亿）。/);
});
