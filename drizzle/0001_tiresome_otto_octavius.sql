CREATE INDEX `idx_records_owner_week_type` ON `bond_records` (`owner_id`,`week_start`,`dataset_type`);--> statement-breakpoint
CREATE INDEX `idx_records_owner_code_date` ON `bond_records` (`owner_id`,`bond_code`,`trade_date`);--> statement-breakpoint
CREATE INDEX `idx_imports_owner_week` ON `imports` (`owner_id`,`week_start`);--> statement-breakpoint
CREATE UNIQUE INDEX `idx_drafts_owner_week` ON `weekly_drafts` (`owner_id`,`week_start`);