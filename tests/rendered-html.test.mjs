import assert from "node:assert/strict";
import test from "node:test";

async function render() {
  const workerUrl = new URL("../dist/server/index.js", import.meta.url);
  workerUrl.searchParams.set("test", `${process.pid}-${Date.now()}`);
  const { default: worker } = await import(workerUrl.href);
  return worker.fetch(
    new Request("http://localhost/", { headers: { accept: "text/html" } }),
    { ASSETS: { fetch: async () => new Response("Not found", { status: 404 }) } },
    { waitUntil() {}, passThroughOnException() {} },
  );
}

test("server-renders the bond issuance workbench", async () => {
  const response = await render();
  assert.equal(response.status, 200);
  assert.match(response.headers.get("content-type") ?? "", /^text\/html\b/i);
  const html = await response.text();
  assert.match(html, /<title>利率债发行工作台<\/title>/i);
  assert.match(html, /利率债发行工作台/);
  assert.match(html, /利率债一级工作台/);
  assert.match(html, /地方债发行明细/);
  assert.match(html, /地方债日表转换/);
  assert.match(html, /日表转换器/);
  assert.match(html, /独立转换工具 · 不自动计入周报/);
  assert.match(html, /一二级利差/);
  assert.match(html, /增量更新/);
  assert.match(html, /地方债明细最新日期/);
  assert.match(html, /一二级利差最新日期/);
  assert.match(html, /新债发行计划/);
  assert.match(html, /不采用文件中的发行量/);
  assert.match(html, /一二级表口径/);
  assert.match(html, /周报生成数据/);
  assert.match(html, /发行时段与到期明细/);
  assert.match(html, /本周发行与到期总览/);
  assert.doesNotMatch(html, /本周完成度/);
  assert.doesNotMatch(html, /progress-ring/);
  assert.doesNotMatch(html, /当前区间/);
  assert.doesNotMatch(html, /周报发行小结/);
  assert.match(html, /周报生成/);
  assert.doesNotMatch(html, /codex-preview|Your site is taking shape|react-loading-skeleton/i);
});
