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
  if (isPlainObject(value)) return `{${Object.keys(value).sort().map(key => `${JSON.stringify(key)}:${stable(value[key])}`).join(",")}}`;
  return JSON.stringify(value);
}

function previousValue(record: StoredRecord, camel: string, snake: string) {
  return record[camel] ?? record[snake];
}

export function recordKey(record: Pick<RecordPayload, "tradeDate" | "bondCode">) {
  return `${record.tradeDate || ""}|${record.bondCode || ""}`;
}

export function mergeRecord(incoming: RecordPayload, existing?: StoredRecord) {
  const output: RecordPayload = {};
  FIELD_NAMES.forEach(([camel, snake]) => {
    const next = incoming[camel];
    const previous = existing ? previousValue(existing, camel, snake) : undefined;
    output[camel] = (hasInformation(next) ? next : previous) as never;
  });
  output.raw = mergeRaw(existing ? parseRaw(existing.raw_json ?? existing.raw) : {}, incoming.raw || {});
  const changed = !existing || FIELD_NAMES.some(([camel, snake]) => stable(output[camel]) !== stable(previousValue(existing, camel, snake)))
    || stable(output.raw) !== stable(parseRaw(existing.raw_json ?? existing.raw));
  return { record: output, changed };
}
