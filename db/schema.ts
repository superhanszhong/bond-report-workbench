import { index, integer, real, sqliteTable, text, uniqueIndex } from "drizzle-orm/sqlite-core";

export const imports = sqliteTable("imports", {
  id: text("id").primaryKey(),
  ownerId: text("owner_id").notNull(),
  datasetType: text("dataset_type").notNull(),
  tradeDate: text("trade_date").notNull(),
  weekStart: text("week_start").notNull(),
  fileName: text("file_name").notNull(),
  recordCount: integer("record_count").notNull(),
  createdAt: text("created_at").notNull(),
}, (table) => [index("idx_imports_owner_week").on(table.ownerId, table.weekStart)]);

export const bondRecords = sqliteTable("bond_records", {
  id: integer("id").primaryKey({ autoIncrement: true }),
  importId: text("import_id").notNull(),
  ownerId: text("owner_id").notNull(),
  datasetType: text("dataset_type").notNull(),
  tradeDate: text("trade_date").notNull(),
  weekStart: text("week_start").notNull(),
  bondCode: text("bond_code"),
  shortName: text("short_name"),
  fullName: text("full_name"),
  issuer: text("issuer"),
  region: text("region"),
  bondType: text("bond_type"),
  issuanceRoute: text("issuance_route"),
  venue: text("venue"),
  bidTime: text("bid_time"),
  tenor: text("tenor"),
  amount: real("amount"),
  spread: real("spread"),
  floorRate: real("floor_rate"),
  fee: real("fee"),
  distributionDate: text("distribution_date"),
  remark: text("remark"),
  rawJson: text("raw_json").notNull(),
}, (table) => [
  index("idx_records_owner_week_type").on(table.ownerId, table.weekStart, table.datasetType),
  index("idx_records_owner_code_date").on(table.ownerId, table.bondCode, table.tradeDate),
  uniqueIndex("idx_records_unique_bond_day").on(table.ownerId, table.datasetType, table.tradeDate, table.bondCode),
]);

export const weeklyDrafts = sqliteTable("weekly_drafts", {
  id: integer("id").primaryKey({ autoIncrement: true }),
  ownerId: text("owner_id").notNull(),
  weekStart: text("week_start").notNull(),
  summaryText: text("summary_text").notNull().default(""),
  reviewText: text("review_text").notNull().default(""),
  updatedAt: text("updated_at").notNull(),
}, (table) => [uniqueIndex("idx_drafts_owner_week").on(table.ownerId, table.weekStart)]);
