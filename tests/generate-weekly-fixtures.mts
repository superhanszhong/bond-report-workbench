import { mkdir, readFile, writeFile } from "node:fs/promises";
import { basename, resolve } from "node:path";
import { buildWeeklyReportBlob } from "../app/lib/report.ts";
import { parseLocalBondFile, parseSpreadFile, spreadSummary } from "../app/lib/workbench.ts";

const spreadPath = "/Users/zhonghanji/Library/Containers/com.tencent.xinWeChat/Data/Documents/xwechat_files/wxid_o1ibaqan2uon21_d019/temp/drag/利率债一二级分析0814(5).xlsx";
const localPath = "/Users/zhonghanji/Downloads/地方政府债+2026-08-15.xlsx";
const templatePath = "/Users/zhonghanji/Desktop/利率债发行周报20260814(8).docx";
const outputDir = resolve("../outputs/site-regression");

const asFile = async (path: string) => new File([await readFile(path)], basename(path));
const [allSpread, allLocal] = await Promise.all([
  parseSpreadFile(await asFile(spreadPath)),
  parseLocalBondFile(await asFile(localPath)),
]);
const templateBuffer = await readFile(templatePath);
const templateBytes = templateBuffer.buffer.slice(templateBuffer.byteOffset, templateBuffer.byteOffset + templateBuffer.byteLength) as ArrayBuffer;

const fixtures = [
  { start: "2026-08-03", end: "2026-08-07", expectedSpread: 29, expectedLocal: 31, expectedLocalAmount: 1818.4003 },
  { start: "2026-08-10", end: "2026-08-14", expectedSpread: 31, expectedLocal: 39, expectedLocalAmount: 2828.2280 },
];

await mkdir(outputDir, { recursive: true });
for (const fixture of fixtures) {
  const spreadRecords = allSpread.filter(row => row.tradeDate >= fixture.start && row.tradeDate <= fixture.end);
  const localRecords = allLocal.filter(row => row.tradeDate >= fixture.start && row.tradeDate <= fixture.end);
  const localAmount = localRecords.reduce((sum, row) => sum + (row.amount || 0), 0);
  if (spreadRecords.length !== fixture.expectedSpread) throw new Error(`${fixture.start} 利率债数量不符`);
  if (localRecords.length !== fixture.expectedLocal) throw new Error(`${fixture.start} 地方债数量不符`);
  if (Math.abs(localAmount - fixture.expectedLocalAmount) > 0.0001) throw new Error(`${fixture.start} 地方债规模不符`);
  const summary = spreadSummary(spreadRecords);
  if (!summary.includes("国债") || !summary.includes("DR浮息债")) throw new Error(`${fixture.start} 小结结构不完整`);
  const previousStartDate = new Date(`${fixture.start}T12:00:00`);
  previousStartDate.setDate(previousStartDate.getDate() - 7);
  const previousEndDate = new Date(`${fixture.end}T12:00:00`);
  previousEndDate.setDate(previousEndDate.getDate() - 7);
  const previousStart = previousStartDate.toISOString().slice(0, 10);
  const previousEnd = previousEndDate.toISOString().slice(0, 10);
  const previousSpreadRecords = allSpread.filter(row => row.tradeDate >= previousStart && row.tradeDate <= previousEnd);
  const ytdLocalRecords = allLocal.filter(row => row.tradeDate >= `${fixture.start.slice(0, 4)}-01-01` && row.tradeDate <= fixture.end);
  const blob = await buildWeeklyReportBlob({ weekStart: fixture.start, summary, localRecords, spreadRecords, previousSpreadRecords, ytdLocalRecords, templateBytes });
  const name = `利率债发行周报${fixture.start.slice(5).replace("-", "")}-${fixture.end.slice(5).replace("-", "")}_网站回归测试.docx`;
  await writeFile(resolve(outputDir, name), Buffer.from(await blob.arrayBuffer()));
  console.log(JSON.stringify({ name, spreadCount: spreadRecords.length, localCount: localRecords.length, localAmount, summary }));
}
