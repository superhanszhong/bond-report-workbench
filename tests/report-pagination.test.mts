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
  assert.equal((xml.match(/<w:pageBreakBefore\b/g) || []).length, 1);
  assert.match(xml, /<w:cantSplit\b/);
  assert.match(xml, /<w:tblHeader\b/);
  assert.match(xml, /<w:keepNext\b/);
});
