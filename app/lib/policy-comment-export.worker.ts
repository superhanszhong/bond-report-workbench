/// <reference lib="webworker" />

import { buildUpdatedSpreadWorkbook } from "./policy-comment-workbook";
import type { PolicyCommentDraft } from "./policy-comment";
import type { ParsedBondRecord } from "./workbench";

type ExportRequest = {
  templateBytes: ArrayBuffer;
  drafts: PolicyCommentDraft[];
  plans: ParsedBondRecord[];
  tradeDate: string;
};

self.addEventListener("message", (event: MessageEvent<ExportRequest>) => {
  try {
    const { templateBytes, drafts, plans, tradeDate } = event.data;
    const bytes = buildUpdatedSpreadWorkbook(templateBytes, drafts, plans, tradeDate);
    self.postMessage({ ok: true, bytes }, { transfer: [bytes] });
  } catch (error) {
    self.postMessage({ ok: false, error: error instanceof Error ? error.message : "生成完整一二级表失败" });
  }
});

