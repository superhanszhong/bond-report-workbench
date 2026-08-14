import { env } from "cloudflare:workers";

type RecordPayload = {
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

interface WorkbenchEnv {
  DB: D1Database;
}

function ownerId(request: Request) {
  return request.headers.get("oai-authenticated-user-id") || "local-preview";
}

function isDate(value: unknown): value is string {
  return typeof value === "string" && /^\d{4}-\d{2}-\d{2}$/.test(value);
}

async function ensureSchema(db: D1Database) {
  await db.batch([
    db.prepare(`CREATE TABLE IF NOT EXISTS imports (
      id TEXT PRIMARY KEY,
      owner_id TEXT NOT NULL,
      dataset_type TEXT NOT NULL,
      trade_date TEXT NOT NULL,
      week_start TEXT NOT NULL,
      file_name TEXT NOT NULL,
      record_count INTEGER NOT NULL,
      created_at TEXT NOT NULL
    )`),
    db.prepare(`CREATE TABLE IF NOT EXISTS bond_records (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      import_id TEXT NOT NULL,
      owner_id TEXT NOT NULL,
      dataset_type TEXT NOT NULL,
      trade_date TEXT NOT NULL,
      week_start TEXT NOT NULL,
      bond_code TEXT,
      short_name TEXT,
      full_name TEXT,
      issuer TEXT,
      region TEXT,
      bond_type TEXT,
      issuance_route TEXT,
      venue TEXT,
      bid_time TEXT,
      tenor TEXT,
      amount REAL,
      spread REAL,
      floor_rate REAL,
      fee REAL,
      distribution_date TEXT,
      remark TEXT,
      raw_json TEXT NOT NULL
    )`),
    db.prepare(`CREATE TABLE IF NOT EXISTS weekly_drafts (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      owner_id TEXT NOT NULL,
      week_start TEXT NOT NULL,
      summary_text TEXT NOT NULL DEFAULT '',
      review_text TEXT NOT NULL DEFAULT '',
      updated_at TEXT NOT NULL
    )`),
    db.prepare("CREATE INDEX IF NOT EXISTS idx_imports_owner_week ON imports(owner_id, week_start)"),
    db.prepare("CREATE INDEX IF NOT EXISTS idx_records_owner_week_type ON bond_records(owner_id, week_start, dataset_type)"),
    db.prepare("CREATE INDEX IF NOT EXISTS idx_records_owner_code_date ON bond_records(owner_id, bond_code, trade_date)"),
    db.prepare("CREATE UNIQUE INDEX IF NOT EXISTS idx_records_unique_bond_day ON bond_records(owner_id, dataset_type, trade_date, bond_code)"),
    db.prepare("CREATE UNIQUE INDEX IF NOT EXISTS idx_drafts_owner_week ON weekly_drafts(owner_id, week_start)"),
  ]);
}

export async function GET(request: Request) {
  try {
    const db = (env as unknown as WorkbenchEnv).DB;
    await ensureSchema(db);
    const url = new URL(request.url);
    const weekStart = url.searchParams.get("weekStart");
    if (!isDate(weekStart)) {
      return Response.json({ error: "weekStart 必须为 YYYY-MM-DD" }, { status: 400 });
    }
    const owner = ownerId(request);
    const [imports, records, draft] = await Promise.all([
      db.prepare("SELECT * FROM imports WHERE owner_id = ? AND week_start = ? ORDER BY trade_date, created_at")
        .bind(owner, weekStart).all(),
      db.prepare("SELECT * FROM bond_records WHERE owner_id = ? AND week_start = ? ORDER BY trade_date, id")
        .bind(owner, weekStart).all(),
      db.prepare("SELECT * FROM weekly_drafts WHERE owner_id = ? AND week_start = ? LIMIT 1")
        .bind(owner, weekStart).first(),
    ]);
    return Response.json({ imports: imports.results, records: records.results, draft });
  } catch (error) {
    const message = error instanceof Error ? error.message : "读取周度数据失败";
    return Response.json({ error: message }, { status: 500 });
  }
}

export async function POST(request: Request) {
  try {
    const db = (env as unknown as WorkbenchEnv).DB;
    await ensureSchema(db);
    const payload = await request.json() as {
      action?: string;
      datasetType?: string;
      tradeDate?: string;
      weekStart?: string;
      fileName?: string;
      records?: RecordPayload[];
      summaryText?: string;
      reviewText?: string;
    };
    const owner = ownerId(request);

    if (payload.action === "saveDraft") {
      if (!isDate(payload.weekStart)) {
        return Response.json({ error: "周起始日无效" }, { status: 400 });
      }
      const now = new Date().toISOString();
      await db.prepare(`INSERT INTO weekly_drafts (owner_id, week_start, summary_text, review_text, updated_at)
        VALUES (?, ?, ?, ?, ?)
        ON CONFLICT(owner_id, week_start) DO UPDATE SET
          summary_text = excluded.summary_text,
          review_text = excluded.review_text,
          updated_at = excluded.updated_at`)
        .bind(owner, payload.weekStart, payload.summaryText || "", payload.reviewText || "", now).run();
      return Response.json({ ok: true });
    }

    if (!isDate(payload.tradeDate) || !isDate(payload.weekStart)) {
      return Response.json({ error: "交易日或周起始日无效" }, { status: 400 });
    }
    if (!payload.datasetType || !payload.fileName || !Array.isArray(payload.records) || payload.records.length === 0) {
      return Response.json({ error: "上传数据不完整" }, { status: 400 });
    }
    const importId = crypto.randomUUID();
    const now = new Date().toISOString();
    const statements = [
      db.prepare(`INSERT INTO imports
        (id, owner_id, dataset_type, trade_date, week_start, file_name, record_count, created_at)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?)`)
        .bind(importId, owner, payload.datasetType, payload.tradeDate, payload.weekStart, payload.fileName, payload.records.length, now),
      ...payload.records.map((row) => db.prepare(`INSERT INTO bond_records
        (import_id, owner_id, dataset_type, trade_date, week_start, bond_code, short_name, full_name,
         issuer, region, bond_type, issuance_route, venue, bid_time, tenor, amount, spread, floor_rate,
         fee, distribution_date, remark, raw_json)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        ON CONFLICT(owner_id, dataset_type, trade_date, bond_code) DO UPDATE SET
          import_id = excluded.import_id,
          week_start = excluded.week_start,
          short_name = excluded.short_name,
          full_name = excluded.full_name,
          issuer = excluded.issuer,
          region = excluded.region,
          bond_type = excluded.bond_type,
          issuance_route = excluded.issuance_route,
          venue = excluded.venue,
          bid_time = excluded.bid_time,
          tenor = excluded.tenor,
          amount = excluded.amount,
          spread = excluded.spread,
          floor_rate = excluded.floor_rate,
          fee = excluded.fee,
          distribution_date = excluded.distribution_date,
          remark = excluded.remark,
          raw_json = excluded.raw_json`)
        .bind(
          importId, owner, payload.datasetType, row.tradeDate || payload.tradeDate, payload.weekStart,
          row.bondCode || null, row.shortName || null, row.fullName || null, row.issuer || null,
          row.region || null, row.bondType || null, row.issuanceRoute || null, row.venue || null,
          row.bidTime || null, row.tenor || null, row.amount ?? null, row.spread ?? null,
          row.floorRate ?? null, row.fee ?? null, row.distributionDate || null, row.remark || null,
          JSON.stringify(row.raw || {})
        )),
    ];
    await db.batch(statements);
    return Response.json({ ok: true, importId, count: payload.records.length }, { status: 201 });
  } catch (error) {
    const message = error instanceof Error ? error.message : "保存上传数据失败";
    return Response.json({ error: message }, { status: 500 });
  }
}

export async function DELETE(request: Request) {
  try {
    const db = (env as unknown as WorkbenchEnv).DB;
    await ensureSchema(db);
    const url = new URL(request.url);
    const importId = url.searchParams.get("importId");
    if (!importId) return Response.json({ error: "缺少 importId" }, { status: 400 });
    const owner = ownerId(request);
    await db.batch([
      db.prepare("DELETE FROM bond_records WHERE owner_id = ? AND import_id = ?").bind(owner, importId),
      db.prepare("DELETE FROM imports WHERE owner_id = ? AND id = ?").bind(owner, importId),
    ]);
    return Response.json({ ok: true });
  } catch (error) {
    const message = error instanceof Error ? error.message : "删除失败";
    return Response.json({ error: message }, { status: 500 });
  }
}
