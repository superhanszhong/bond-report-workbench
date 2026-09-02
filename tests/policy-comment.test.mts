import assert from "node:assert/strict";
import test from "node:test";
import type { ParsedBondRecord, SpreadSummaryMeta } from "../app/lib/workbench";
import { createPolicyCommentDrafts, policyComments, policyDraftResults, type PolicyCommentDraft } from "../app/lib/policy-comment";

function meta(overrides: Partial<SpreadSummaryMeta>): SpreadSummaryMeta {
  return {
    baseCode: "", route: "中债招标", rateType: "固息或贴现", displaySpreadText: "",
    auctionSpreadText: "", allInText: "", secondaryText: "", winningRateText: "", note: "", proceeds: "",
    ...overrides,
  };
}

function history(code: string, date: string, auctionSpreadText: string, rateType = "固息或贴现"): ParsedBondRecord {
  return {
    tradeDate: date, bondCode: code, issuer: "中国农业发展银行", bondType: "农发债", tenor: code.startsWith("09260409") ? "2" : "3",
    remark: rateType, summaryMeta: meta({ baseCode: code.replace(/[XZ]\d*$/i, ""), rateType, auctionSpreadText }),
  };
}

function current(code: string, tenor: string, date: string, auctionSpreadText: string, winningRateText: string, secondaryText: string, rateType = "固息或贴现", allInText = ""): ParsedBondRecord {
  return {
    tradeDate: date, bondCode: code, issuer: "中国农业发展银行", bondType: "农发债", tenor,
    remark: rateType, summaryMeta: meta({ baseCode: code.replace(/[XZ]\d*$/i, ""), rateType, auctionSpreadText, winningRateText, secondaryText, allInText }),
  };
}

test("政金债短评复现工作样例", () => {
  const rows: ParsedBondRecord[] = [
    history("250409Z30", "2026-07-27", "-4.77(较估值(1.46))", "LPR浮息债"),
    history("260403Z22", "2026-07-29", "-0.94(较二级(1.495))"),
    history("260405Z5", "2026-07-29", "-2.23(较二级(1.5875))"),
    history("260410Z5", "2026-07-29", "-1.43(较二级(1.808))"),
    history("09260409Z15", "2026-07-21", "高估价净价(99.9528)0.0009元", "DR浮息债"),
    current("250409Z31", "3", "2026-08-03", "-3.49(较二级(1.473))", "1.4381%", "1.473%", "LPR浮息债"),
    current("260403Z23", "3", "2026-08-03", "-1.91(较二级(1.5025))", "1.4834%", "1.5025%"),
    current("260405Z6", "5", "2026-08-03", "-1.53(较估值(1.575))", "1.5597%", "1.575%"),
    current("260410Z6", "10", "2026-08-03", "-1.85(较二级(1.805))", "1.7865%", "1.805%"),
    current("09260409Z16", "2", "2026-07-28", "高估价净价(99.95)0.0139元", "1.4791%", "99.95元", "DR浮息债", "99.9639元"),
  ];
  const results = [...policyComments(rows, "2026-08-03"), ...policyComments(rows, "2026-07-28")];
  const byCode = new Map(results.map((row) => [row.displayCode, row.text]));
  assert.equal(byCode.get("250409X31"), "中标利率1.4381% 低二级(1.473)3.49BP\n今日增发农发3Y LPR浮息债,利差收窄(上次-4.77(较估值(1.46)))");
  assert.equal(byCode.get("260403X23"), "中标利率1.4834% 低二级(1.5025)1.91BP\n今日增发农发3Y,利差走阔(上次-0.94(较二级(1.495)))");
  assert.equal(byCode.get("260405X6"), "中标利率1.5597% 低估值(1.575)1.53BP\n今日增发农发5Y,利差收窄(上次-2.23(较二级(1.5875)))");
  assert.equal(byCode.get("260410X6"), "中标利率1.7865% 低二级(1.805)1.85BP\n今日增发农发10Y,利差走阔(上次-1.43(较二级(1.808)))");
  assert.equal(byCode.get("09260409Z16"), "中标净价99.9639元 高估价净价(99.95)0.0139元\n今日增发农发清发2Y DR007浮息债,利差走阔(上次高估价净价0.0009元)");
});

test("新券只计算一二级收益率差，不判断历史利差", () => {
  const rows: ParsedBondRecord[] = [
    history("269999Z1", "2026-08-20", "-3.00(较二级(1.5))", "DR007浮息债"),
    current("269999", "3", "2026-08-31", "-1.50(较二级(1.5))", "1.485%", "1.5%", "DR007浮息债"),
  ];
  const result = policyComments(rows, "2026-08-31")[0];
  assert.equal(result.text, "中标利率1.485% 低二级(1.5)1.50BP\n今日新发农发清发3Y DR007浮息债");
  assert.equal(result.movement, "不判断");
  assert.equal(result.previousCode, "");
  assert.equal(result.warning, undefined);
});

test("DR001 增发券按净价差生成文字", () => {
  const prior = history("09260101Z1", "2026-08-24", "高估价净价(99.9528)0.0009元", "DR001浮息债");
  const draft: PolicyCommentDraft = {
    id: "2026-08-31|09260101Z2", tradeDate: "2026-08-31", bondCode: "09260101Z2", shortName: "26农发清发01",
    issuer: "中国农业发展银行", bondType: "农发债", tenor: "2Y", rateType: "DR001浮息债", route: "上清所",
    benchmarkType: "估价", benchmarkValue: "99.95", finalValue: "99.9639",
  };
  const result = policyDraftResults([draft], [prior])[0].comment!;
  assert.equal(result.text, "中标净价99.9639元 高估价净价(99.95)0.0139元\n今日增发农发清发2Y DR001浮息债,利差走阔(上次高估价净价0.0009元)");
});

test("DR 净价债也可填写参考券", () => {
  const draft: PolicyCommentDraft = {
    id: "2026-09-01|09260409Z21", tradeDate: "2026-09-01", bondCode: "09260409Z21", shortName: "26农发清发09(增发21)",
    issuer: "中国农业发展银行", bondType: "农发债", tenor: "2Y", rateType: "DR007浮息债", route: "上清所",
    benchmarkType: "二级", referenceBond: "260214", benchmarkValue: "99.9775", finalValue: "99.9767",
  };
  const result = policyDraftResults([draft], [])[0].comment!;
  assert.match(result.firstLine, /低较260214二级净价\(99\.9775\)0\.0008元/);
});

test("DR 增发券的上次记录若只有收益率差，不误判为净价差", () => {
  const prior = history("260217", "2026-08-11", "-1.36(较260214估值(1.5036))", "DR浮息债");
  const draft: PolicyCommentDraft = {
    id: "2026-08-25|260217Z1", tradeDate: "2026-08-25", bondCode: "260217Z1", shortName: "26国开清发17(增发1)",
    issuer: "国家开发银行", bondType: "国开债", tenor: "3Y", rateType: "DR007浮息债", route: "上清所",
    benchmarkType: "估价", benchmarkValue: "100.0052", finalValue: "100.0033",
  };
  const result = policyDraftResults([draft], [prior])[0].comment!;
  assert.equal(result.movement, "不判断");
  assert.match(result.text, /今日增发国开清发3Y DR007浮息债,本次不作利差变化判断/);
  assert.doesNotMatch(result.text, /1\.5036元|利差走[阔窄]/);
});

test("报价发行优先使用报价利率，普通清发债文案保留清发", () => {
  const quote = current("09260401Z2", "1", "2026-08-31", "-1.00(较二级(1.5))", "1.49%", "1.5%");
  quote.shortName = "26农发清发01(增发2)";
  quote.issuanceRoute = "报价发行";
  quote.summaryMeta = meta({ ...quote.summaryMeta, route: "报价发行" });
  const result = policyComments([quote], "2026-08-31")[0];
  assert.equal(result.firstLine, "报价利率1.49% 低二级(1.5)1.00BP");
  assert.match(result.secondLine, /^今日增发农发清发1Y/);
});

test("历史备注将 DR01 归一为 DR001，将普通 DR 浮息债归一为 DR007", () => {
  const plans: ParsedBondRecord[] = [
    { tradeDate: "2026-08-31", bondCode: "092603001Z2", shortName: "26进出清发001(增发2)", issuer: "中国进出口银行", bondType: "口行债", tenor: "1Y" },
    { tradeDate: "2026-08-31", bondCode: "260217Z2", shortName: "26国开清发17(增发2)", issuer: "国家开发银行", bondType: "国开债", tenor: "3Y" },
  ];
  const histories = [
    history("092603001Z1", "2026-08-24", "高估价净价(99.99)0.001元", "DR01浮息债 前台报价发行"),
    { ...history("260217Z1", "2026-08-25", "高估价净价(100.0052)0.0019元", "DR浮息债"), issuer: "国家开发银行", bondType: "国开债" },
  ];
  const drafts = createPolicyCommentDrafts(plans, histories);
  assert.deepEqual(drafts.map((draft) => draft.rateType), ["DR001浮息债", "DR007浮息债"]);
});

test("参考券默认留空，仅在手动填写后按参考收益率生成小结", () => {
  const plan: ParsedBondRecord = {
    tradeDate: "2026-08-25", bondCode: "260217Z2", shortName: "26国开清发17(增发2)",
    issuer: "国家开发银行", bondType: "国开债", tenor: "3Y",
  };
  const prior = {
    ...history("260217Z1", "2026-08-18", "-1.36(较260214估值(1.5036))"),
    issuer: "国家开发银行", bondType: "国开债",
  };
  const [draft] = createPolicyCommentDrafts([plan], [prior]);
  assert.equal(draft.benchmarkType, "估值");
  assert.equal(draft.referenceBond, "");
  const [result] = policyDraftResults([{ ...draft, referenceBond: "260214", benchmarkValue: "1.5036", finalValue: "1.49" }], [prior]);
  assert.match(result.comment!.firstLine, /低较260214估值\(1\.5036\)1\.36BP/);
});

test("参考券留空时默认按本券基准生成，不增加额外处理", () => {
  const draft: PolicyCommentDraft = {
    id: "2026-09-01|260308X2", tradeDate: "2026-09-01", bondCode: "260308X2", shortName: "26进出08(增发2)",
    issuer: "中国进出口银行", bondType: "口行债", tenor: "1Y", rateType: "", route: "中债招标",
    benchmarkType: "二级", referenceBond: "", benchmarkValue: "1.42", finalValue: "1.3801",
  };
  const result = policyDraftResults([draft], [])[0];
  assert.deepEqual(result.missing, []);
  assert.match(result.comment!.firstLine, /低二级\(1\.42\)3\.99BP/);
});

test("短评核对增发期数并提示跳号", () => {
  const plans: ParsedBondRecord[] = [
    { tradeDate: "2026-09-01", bondCode: "260308X2", issuer: "中国进出口银行", bondType: "口行债", tenor: "1Y" },
    { tradeDate: "2026-09-01", bondCode: "260405X6", issuer: "中国农业发展银行", bondType: "农发债", tenor: "5Y" },
  ];
  const histories = [
    { ...history("260308Z1", "2026-08-25", "-1.00(较二级(1.4))"), issuer: "中国进出口银行", bondType: "口行债" },
    history("260405Z3", "2026-08-25", "-1.00(较二级(1.5))"),
  ];
  const drafts = createPolicyCommentDrafts(plans, histories);
  assert.equal(drafts[0].sequenceCheck?.status, "ok");
  assert.match(drafts[0].sequenceCheck?.message || "", /260308X(?:\D|$)/);
  assert.equal(drafts[1].sequenceCheck?.status, "warning");
  assert.match(drafts[1].sequenceCheck?.message || "", /跳号/);
});

test("首期增发代码 X1 按市场习惯显示为 X，X 仍识别为第一期", () => {
  const first = current("260308X1", "1", "2026-08-25", "-1.00(较二级(1.4))", "1.39%", "1.4%");
  const comments = policyComments([first], "2026-08-25");
  assert.equal(comments[0].displayCode, "260308X");
  const plan = { ...first, tradeDate: "2026-09-01", bondCode: "260309X2" };
  const prior = { ...first, bondCode: "260309X" };
  const [draft] = createPolicyCommentDrafts([plan], [prior]);
  assert.equal(draft.sequenceCheck?.status, "ok");
  assert.match(draft.sequenceCheck?.message || "", /260309X(?:\D|$)/);
});

test("同一基础代码一周多次发行时按日期和完整代码保留各自行情填写", () => {
  const plans: ParsedBondRecord[] = [
    { tradeDate: "2026-08-24", bondCode: "260403Z28", shortName: "26农发03(增发28)", issuer: "中国农业发展银行", bondType: "农发债", tenor: "3Y" },
    { tradeDate: "2026-08-27", bondCode: "260403Z29", shortName: "26农发03(增发29)", issuer: "中国农业发展银行", bondType: "农发债", tenor: "3Y" },
  ];
  const existing: PolicyCommentDraft[] = plans.map((plan, index) => ({
    id: `${plan.tradeDate}|${plan.bondCode}`, tradeDate: plan.tradeDate, bondCode: plan.bondCode!, shortName: plan.shortName!,
    issuer: plan.issuer!, bondType: plan.bondType!, tenor: "3Y", rateType: "", route: "中债招标", benchmarkType: "二级",
    benchmarkValue: index ? "1.52" : "1.50", finalValue: index ? "1.51" : "1.49",
  }));
  const rebuilt = createPolicyCommentDrafts(plans, [], existing);
  assert.deepEqual(rebuilt.map((draft) => [draft.benchmarkValue, draft.finalValue]), [["1.50", "1.49"], ["1.52", "1.51"]]);
});

test("二级价格为空时不把空字符串当成 0 提前生成文字", () => {
  const draft: PolicyCommentDraft = {
    id: "2026-08-31|260308Z1", tradeDate: "2026-08-31", bondCode: "260308Z1", shortName: "26进出08(增发1)",
    issuer: "中国进出口银行", bondType: "口行债", tenor: "1Y", rateType: "", route: "中债招标",
    benchmarkType: "二级", benchmarkValue: "", finalValue: "1.3909",
  };
  const result = policyDraftResults([draft], [] as ParsedBondRecord[])[0];
  assert.equal(result.comment, null);
  assert.deepEqual(result.missing, ["二级/估值"]);
});

test("2026-09-01 五只政金债发行小结回归", () => {
  const histories: ParsedBondRecord[] = [
    { ...history("260308Z1", "2026-08-28", "-3.91(较二级(1.425))"), issuer: "中国进出口银行", bondType: "口行债", tenor: "1" },
    { ...history("09260402Z10", "2026-08-25", "-0.25(较估值)"), tenor: "2" },
    { ...history("09260407Z06", "2026-08-25", "0(平二级)"), tenor: "7" },
    history("09260409Z20", "2026-08-27", "高估价净价(99.95)0.0158元", "DR浮息债"),
  ];
  const drafts: PolicyCommentDraft[] = [
    {
      id: "2026-09-01|260308X2", tradeDate: "2026-09-01", bondCode: "260308X2", shortName: "26进出08(增发2)",
      issuer: "中国进出口银行", bondType: "口行债", tenor: "1Y", rateType: "", route: "中债招标",
      benchmarkType: "二级", benchmarkValue: "1.42", finalValue: "1.3801",
    },
    {
      id: "2026-09-01|09260411", tradeDate: "2026-09-01", bondCode: "09260411", shortName: "26农发清发11",
      issuer: "中国农业发展银行", bondType: "农发债", tenor: "1.1123Y", rateType: "", route: "上清所",
      benchmarkType: "估值曲线", benchmarkValue: "1.4184", finalValue: "1.42",
    },
    {
      id: "2026-09-01|09260409Z21", tradeDate: "2026-09-01", bondCode: "09260409Z21", shortName: "26农发清发09(增发21)",
      issuer: "中国农业发展银行", bondType: "农发债", tenor: "2Y", rateType: "DR007浮息债", route: "上清所",
      benchmarkType: "二级", benchmarkValue: "99.9775", finalValue: "99.9767",
    },
    {
      id: "2026-09-01|09260402Z11", tradeDate: "2026-09-01", bondCode: "09260402Z11", shortName: "26农发清发02(增发11)",
      issuer: "中国农业发展银行", bondType: "农发债", tenor: "2Y", rateType: "", route: "上清所",
      benchmarkType: "二级", benchmarkValue: "1.465", finalValue: "1.4643",
    },
    {
      id: "2026-09-01|09260407Z07", tradeDate: "2026-09-01", bondCode: "09260407Z07", shortName: "26农发清发07(增发7)",
      issuer: "中国农业发展银行", bondType: "农发债", tenor: "7Y", rateType: "", route: "上清所",
      benchmarkType: "二级", benchmarkValue: "1.6625", finalValue: "1.6635",
    },
  ];
  const byCode = new Map(policyDraftResults(drafts, histories).map(({ comment }) => [comment?.bondCode, comment?.text]));
  assert.equal(byCode.get("260308X2"), "中标利率1.3801% 低二级(1.42)3.99BP\n今日增发口行1Y,利差走阔(上次-3.91(较二级(1.425)))");
  assert.equal(byCode.get("09260411"), "中标利率1.42% 高估值曲线(1.4184)0.16BP\n今日新发农发清发1.1123Y");
  assert.equal(byCode.get("09260409Z21"), "中标净价99.9767元 低二级净价(99.9775)0.0008元\n今日增发农发清发2Y DR007浮息债,利差反转(上次高估价净价0.0158元)");
  assert.equal(byCode.get("09260402Z11"), "中标利率1.4643% 低二级(1.465)0.07BP\n今日增发农发清发2Y,利差收窄(上次-0.25(较估值))");
  assert.equal(byCode.get("09260407Z07"), "中标利率1.6635% 高二级(1.6625)0.10BP\n今日增发农发清发7Y,利差反转(上次0(平二级))");
});
