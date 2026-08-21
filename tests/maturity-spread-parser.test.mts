import assert from "node:assert/strict";
import test from "node:test";
import XLSX from "xlsx-js-style";
import { parseMaturityFile, resolveSpreadBp } from "../app/lib/workbench.ts";

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
