import { mergeRecord, recordKey, type RecordPayload, type StoredRecord } from "./record-merge";

export const LOCAL_STORAGE_MODE = import.meta.env?.VITE_STORAGE_MODE === "local";

type LocalImport = {
  id: string;
  dataset_type: string;
  trade_date: string;
  week_start: string;
  file_name: string;
  record_count: number;
  created_at: string;
};

type LocalRecord = StoredRecord & {
  id: string;
  import_id: string;
  dataset_type: string;
  trade_date: string;
  week_start: string;
  bond_code: string;
};

type LocalDraft = {
  week_start: string;
  summary_text: string;
  review_text: string;
  updated_at: string;
};

type StoreName = "imports" | "records" | "drafts";
const DATABASE_NAME = "bond-report-workbench";
const DATABASE_VERSION = 1;

function requestResult<T>(request: IDBRequest<T>) {
  return new Promise<T>((resolve, reject) => {
    request.onsuccess = () => resolve(request.result);
    request.onerror = () => reject(request.error || new Error("浏览器数据读取失败"));
  });
}

function transactionDone(transaction: IDBTransaction) {
  return new Promise<void>((resolve, reject) => {
    transaction.oncomplete = () => resolve();
    transaction.onerror = () => reject(transaction.error || new Error("浏览器数据保存失败"));
    transaction.onabort = () => reject(transaction.error || new Error("浏览器数据保存已取消"));
  });
}

function openDatabase() {
  return new Promise<IDBDatabase>((resolve, reject) => {
    const request = indexedDB.open(DATABASE_NAME, DATABASE_VERSION);
    request.onupgradeneeded = () => {
      const database = request.result;
      if (!database.objectStoreNames.contains("imports")) database.createObjectStore("imports", { keyPath: "id" });
      if (!database.objectStoreNames.contains("records")) database.createObjectStore("records", { keyPath: "id" });
      if (!database.objectStoreNames.contains("drafts")) database.createObjectStore("drafts", { keyPath: "week_start" });
    };
    request.onsuccess = () => resolve(request.result);
    request.onerror = () => reject(request.error || new Error("无法打开浏览器历史库"));
  });
}

async function getAll<T>(database: IDBDatabase, storeName: StoreName) {
  const transaction = database.transaction(storeName, "readonly");
  return requestResult(transaction.objectStore(storeName).getAll()) as Promise<T[]>;
}

function response(payload: unknown, status = 200) {
  return new Response(JSON.stringify(payload), {
    status,
    headers: { "Content-Type": "application/json; charset=utf-8" },
  });
}

function storedRecord(payload: RecordPayload, datasetType: string, weekStart: string, importId: string): LocalRecord {
  const tradeDate = payload.tradeDate || "";
  const bondCode = payload.bondCode || "";
  return {
    id: `${datasetType}|${tradeDate}|${bondCode}`,
    import_id: importId,
    dataset_type: datasetType,
    trade_date: tradeDate,
    week_start: weekStart,
    bond_code: bondCode,
    short_name: payload.shortName || null,
    full_name: payload.fullName || null,
    issuer: payload.issuer || null,
    region: payload.region || null,
    bond_type: payload.bondType || null,
    issuance_route: payload.issuanceRoute || null,
    venue: payload.venue || null,
    bid_time: payload.bidTime || null,
    tenor: payload.tenor || null,
    amount: payload.amount ?? null,
    spread: payload.spread ?? null,
    floor_rate: payload.floorRate ?? null,
    fee: payload.fee ?? null,
    distribution_date: payload.distributionDate || null,
    remark: payload.remark || null,
    raw_json: JSON.stringify(payload.raw || {}),
  };
}

async function localGet(url: URL) {
  const database = await openDatabase();
  try {
    const records = await getAll<LocalRecord>(database, "records");
    if (url.searchParams.get("meta") === "latest") {
      const latestDates: Record<string, string> = {};
      records.forEach((row) => {
        if (!latestDates[row.dataset_type] || row.trade_date > latestDates[row.dataset_type]) latestDates[row.dataset_type] = row.trade_date;
      });
      return response({ latestDates });
    }
    const startDate = url.searchParams.get("startDate");
    const endDate = url.searchParams.get("endDate");
    if (startDate || endDate) {
      if (!startDate || !endDate || startDate > endDate) return response({ error: "分析日期区间无效" }, 400);
      const datasetType = url.searchParams.get("datasetType") === "local_bond" ? "local_bond" : "spread";
      return response({ records: records.filter((row) => row.dataset_type === datasetType && row.trade_date >= startDate && row.trade_date <= endDate)
        .sort((a, b) => `${a.trade_date}|${a.id}`.localeCompare(`${b.trade_date}|${b.id}`)) });
    }
    const weekStart = url.searchParams.get("weekStart");
    if (!weekStart) return response({ error: "weekStart 必须为 YYYY-MM-DD" }, 400);
    const [imports, drafts] = await Promise.all([
      getAll<LocalImport>(database, "imports"),
      getAll<LocalDraft>(database, "drafts"),
    ]);
    return response({
      imports: imports.filter((row) => row.week_start === weekStart).sort((a, b) => `${a.trade_date}|${a.created_at}`.localeCompare(`${b.trade_date}|${b.created_at}`)),
      records: records.filter((row) => row.week_start === weekStart).sort((a, b) => `${a.trade_date}|${a.id}`.localeCompare(`${b.trade_date}|${b.id}`)),
      draft: drafts.find((row) => row.week_start === weekStart) || null,
    });
  } finally {
    database.close();
  }
}

async function localPost(init?: RequestInit) {
  const payload = JSON.parse(String(init?.body || "{}")) as {
    action?: string;
    datasetType?: string;
    tradeDate?: string;
    weekStart?: string;
    fileName?: string;
    records?: RecordPayload[];
    summaryText?: string;
    reviewText?: string;
  };
  const database = await openDatabase();
  try {
    if (payload.action === "saveDraft") {
      if (!payload.weekStart) return response({ error: "周起始日无效" }, 400);
      const transaction = database.transaction("drafts", "readwrite");
      transaction.objectStore("drafts").put({
        week_start: payload.weekStart,
        summary_text: payload.summaryText || "",
        review_text: payload.reviewText || "",
        updated_at: new Date().toISOString(),
      } satisfies LocalDraft);
      await transactionDone(transaction);
      return response({ ok: true });
    }
    if (!payload.datasetType || !payload.tradeDate || !payload.weekStart || !payload.fileName || !payload.records?.length) {
      return response({ error: "上传数据不完整" }, 400);
    }
    const allRecords = await getAll<LocalRecord>(database, "records");
    const existingRows = allRecords.filter((row) => row.dataset_type === payload.datasetType && row.week_start === payload.weekStart);
    const existingByKey = new Map(existingRows.map((row) => [recordKey({ tradeDate: row.trade_date, bondCode: row.bond_code }), row]));
    const mergedRows = payload.records.map((row) => {
      const normalized = { ...row, tradeDate: row.tradeDate || payload.tradeDate };
      const existing = existingByKey.get(recordKey(normalized));
      return { existing, ...mergeRecord(normalized, existing) };
    });
    const changes = mergedRows.filter((row) => row.changed);
    const inserted = changes.filter((row) => !row.existing).length;
    const updated = changes.length - inserted;
    if (!changes.length) return response({ ok: true, count: 0, inserted: 0, updated: 0, unchanged: payload.records.length });
    const importId = crypto.randomUUID();
    const transaction = database.transaction(["imports", "records"], "readwrite");
    transaction.objectStore("imports").put({
      id: importId,
      dataset_type: payload.datasetType,
      trade_date: payload.tradeDate,
      week_start: payload.weekStart,
      file_name: payload.fileName,
      record_count: changes.length,
      created_at: new Date().toISOString(),
    } satisfies LocalImport);
    changes.forEach(({ record }) => transaction.objectStore("records").put(storedRecord(record, payload.datasetType!, payload.weekStart!, importId)));
    await transactionDone(transaction);
    return response({ ok: true, importId, count: changes.length, inserted, updated, unchanged: payload.records.length - changes.length }, 201);
  } finally {
    database.close();
  }
}

async function localDelete(url: URL) {
  const importId = url.searchParams.get("importId");
  if (!importId) return response({ error: "缺少 importId" }, 400);
  const database = await openDatabase();
  try {
    const records = await getAll<LocalRecord>(database, "records");
    const transaction = database.transaction(["imports", "records"], "readwrite");
    transaction.objectStore("imports").delete(importId);
    records.filter((row) => row.import_id === importId).forEach((row) => transaction.objectStore("records").delete(row.id));
    await transactionDone(transaction);
    return response({ ok: true });
  } finally {
    database.close();
  }
}

async function localRequest(input: string, init?: RequestInit) {
  const url = new URL(input, window.location.origin);
  const method = (init?.method || "GET").toUpperCase();
  if (method === "GET") return localGet(url);
  if (method === "POST") return localPost(init);
  if (method === "DELETE") return localDelete(url);
  return response({ error: "不支持的请求" }, 405);
}

export function workbenchRequest(input: string, init?: RequestInit) {
  return LOCAL_STORAGE_MODE ? localRequest(input, init) : fetch(input, init);
}
