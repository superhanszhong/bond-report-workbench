import { bondIssueKey, normalizeBondCode } from "./bond-code.ts";

export type RecordPayload = {
  tradeDate?: string;
  bondCode?: string;
  shortName?: string;
  fullName?: string;
  issuer?: string;
  region?: string;
  bondType?: string;
  issuanceRoute?: string;
  venue?: string;
  bidTime?: string;
  tenor?: string;
  amount?: number | null;
  spread?: number | null;
  floorRate?: number | null;
  fee?: number | null;
  distributionDate?: string;
  remark?: string;
  raw?: Record<string, unknown>;
};

export type StoredRecord = Record<string, unknown> & { raw_json?: string | null };

const FIELD_NAMES: Array<[keyof RecordPayload, string]> = [
  ["tradeDate", "trade_date"], ["bondCode", "bond_code"], ["shortName", "short_name"],
  ["fullName", "full_name"], ["issuer", "issuer"], ["region", "region"], ["bondType", "bond_type"],
  ["issuanceRoute", "issuance_route"], ["venue", "venue"], ["bidTime", "bid_time"], ["tenor", "tenor"],
  ["amount", "amount"], ["spread", "spread"], ["floorRate", "floor_rate"], ["fee", "fee"],
  ["distributionDate", "distribution_date"], ["remark", "remark"],
];

function isPlainObject(value: unknown): value is Record<string, unknown> {
  return Boolean(value) && typeof value === "object" && !Array.isArray(value);
}

function hasInformation(value: unknown) {
  return value !== null && value !== undefined
    && (typeof value !== "string" || !["", "-", "—"].includes(value.trim()));
}

function parseRaw(value: unknown) {
  if (!value) return {};
  if (isPlainObject(value)) return value;
  if (typeof value !== "string") return {};
  try {
    const parsed = JSON.parse(value) as unknown;
    return isPlainObject(parsed) ? parsed : {};
  } catch { return {}; }
}

function mergeRaw(existing: Record<string, unknown>, incoming: Record<string, unknown>): Record<string, unknown> {
  const merged: Record<string, unknown> = { ...existing };
  Object.entries(incoming).forEach(([key, value]) => {
    const previous = merged[key];
    if (isPlainObject(previous) && isPlainObject(value)) {
      merged[key] = mergeRaw(previous, value);
    } else if (hasInformation(value)) {
      merged[key] = value;
    }
  });
  return merged;
}

function stable(value: unknown): string {
  if (Array.isArray(value)) return `[${value.map(stable).join(",")}]`;
  if (isPlainObject(value)) return `{${Object.keys(value).filter(key => value[key] !== undefined).sort().map(key => `${JSON.stringify(key)}:${stable(value[key])}`).join(",")}}`;
  return JSON.stringify(value);
}

function previousValue(record: StoredRecord, camel: string, snake: string) {
  return record[camel] ?? record[snake];
}

export function recordKey(record: Pick<RecordPayload, "tradeDate" | "bondCode">) {
  return `${record.tradeDate || ""}|${bondIssueKey(record.bondCode)}`;
}

export function mergeRecord(incoming: RecordPayload, existing?: StoredRecord) {
  const output: RecordPayload = {};
  FIELD_NAMES.forEach(([camel, snake]) => {
    const next = incoming[camel];
    const previous = existing ? previousValue(existing, camel, snake) : undefined;
    output[camel] = (hasInformation(next) ? next : previous) as never;
  });
  output.raw = mergeRaw(existing ? parseRaw(existing.raw_json ?? existing.raw) : {}, incoming.raw || {});
  if (output.bondCode) output.bondCode = normalizeBondCode(output.bondCode);
  const changed = !existing || FIELD_NAMES.some(([camel, snake]) => stable(output[camel]) !== stable(previousValue(existing, camel, snake)))
    || stable(output.raw) !== stable(parseRaw(existing.raw_json ?? existing.raw));
  return { record: output, changed };
}

function comparable(value: unknown, tenor = false) {
  if (!hasInformation(value)) return null;
  const string = String(value).trim().replace(/[,，]/g, "");
  const numeric = (tenor ? string.replace(/Y$/i, "") : string.replace(/[%元]$/, "")).trim();
  return /^[-+]?\d+(?:\.\d+)?$/.test(numeric) ? Number(numeric) : string;
}

function payloadOf(row: StoredRecord): RecordPayload {
  const payload: RecordPayload = {};
  FIELD_NAMES.forEach(([camel, snake]) => { payload[camel] = previousValue(row, camel, snake) as never; });
  payload.raw = parseRaw(row.raw_json ?? row.raw);
  payload.bondCode = normalizeBondCode(payload.bondCode);
  return payload;
}

// Non-destructive read reconciliation: raw stored records remain recoverable.
// Never merge different dates or reopenings, or silently choose conflicting results.
export function consolidateBondRecords<T extends object>(rows: T[]): T[] {
  const unique = new Map<string, T>();
  rows.forEach((item, index) => {
    const source = item as StoredRecord;
    const incoming = payloadOf(source);
    const key = incoming.bondCode ? `${source.dataset_type || ""}|${recordKey(incoming)}` : `missing:${index}`;
    const previous = unique.get(key) as StoredRecord | undefined;
    if (previous) {
      const prior = payloadOf(previous);
      const checks: [string, unknown, unknown, boolean?][] = [
        ["发行量", prior.amount, incoming.amount], ["期限", prior.tenor, incoming.tenor, true],
      ];
      const oldDisplay = parseRaw(prior.raw?.__display);
      const newDisplay = parseRaw(incoming.raw?.__display);
      for (const field of ["中标利率", "综收", "中标净价", "前一日估值", "二级", "全场倍数", "边际倍数"]) {
        checks.push([field, oldDisplay[field] ?? prior.raw?.[field], newDisplay[field] ?? incoming.raw?.[field]]);
      }
      const conflict = checks.find(([, a, b, tenor]) => comparable(a, tenor) !== null && comparable(b, tenor) !== null && comparable(a, tenor) !== comparable(b, tenor));
      if (conflict) throw new Error(`${incoming.tradeDate} ${incoming.bondCode} 存在重复记录且${conflict[0]}冲突，请核对后重新上传该券完整记录`);
    }
    const merged = mergeRecord(incoming, previous).record;
    const result: StoredRecord = { ...previous, ...source, ...merged };
    const meta = parseRaw(source.summaryMeta ?? merged.raw?.__summaryMeta);
    if (Object.keys(meta).length) {
      const normalizedMeta = { ...meta, baseCode: merged.bondCode?.replace(/[XZ]\d*$/, ""),
        route: meta.route === "报价发行" ? "报价发行" : /^09/.test(merged.bondCode || "") ? "上清所" : meta.route };
      result.summaryMeta = normalizedMeta;
      merged.raw = { ...merged.raw, __summaryMeta: normalizedMeta };
      result.raw = merged.raw;
    }
    if ("bond_code" in source) result.bond_code = merged.bondCode;
    if ("raw_json" in source) result.raw_json = JSON.stringify(merged.raw || {});
    unique.set(key, result as T);
  });
  return [...unique.values()];
}

export function prepareRecordUpdate(incoming: RecordPayload, aliases: StoredRecord[]) {
  if (!aliases.length) return mergeRecord(incoming);
  // Apply the explicit new values to every alias before checking conflicts.
  // Missing incoming values do not erase complementary historical fields or resolve conflicts by accident.
  const revised = aliases.map(alias => ({ ...alias, ...mergeRecord(incoming, alias).record, raw_json: undefined }));
  const unified = consolidateBondRecords(revised)[0];
  const record = mergeRecord(incoming, unified).record;
  return { record, changed: aliases.some(alias => mergeRecord(record, alias).changed) };
}

export function withImportHistory(record: RecordPayload, existing: StoredRecord | undefined, fileName: string, importedAt: string) {
  const oldRaw = existing ? parseRaw(existing.raw_json ?? existing.raw) : {};
  const { __history: history, ...previousRaw } = oldRaw;
  const previous = existing ? { ...payloadOf(existing), raw: previousRaw } : null;
  return { ...record, raw: { ...record.raw,
    __source: { fileName, importedAt },
    __history: [...(Array.isArray(history) ? history : []), ...(previous ? [{ replacedAt: importedAt, sourceImportId: existing?.import_id, previous }] : [])],
  } };
}
