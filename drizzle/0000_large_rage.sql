CREATE TABLE `bond_records` (
	`id` integer PRIMARY KEY AUTOINCREMENT NOT NULL,
	`import_id` text NOT NULL,
	`owner_id` text NOT NULL,
	`dataset_type` text NOT NULL,
	`trade_date` text NOT NULL,
	`week_start` text NOT NULL,
	`bond_code` text,
	`short_name` text,
	`full_name` text,
	`issuer` text,
	`region` text,
	`bond_type` text,
	`issuance_route` text,
	`venue` text,
	`bid_time` text,
	`tenor` text,
	`amount` real,
	`spread` real,
	`floor_rate` real,
	`fee` real,
	`distribution_date` text,
	`remark` text,
	`raw_json` text NOT NULL
);
--> statement-breakpoint
CREATE TABLE `imports` (
	`id` text PRIMARY KEY NOT NULL,
	`owner_id` text NOT NULL,
	`dataset_type` text NOT NULL,
	`trade_date` text NOT NULL,
	`week_start` text NOT NULL,
	`file_name` text NOT NULL,
	`record_count` integer NOT NULL,
	`created_at` text NOT NULL
);
--> statement-breakpoint
CREATE TABLE `weekly_drafts` (
	`id` integer PRIMARY KEY AUTOINCREMENT NOT NULL,
	`owner_id` text NOT NULL,
	`week_start` text NOT NULL,
	`summary_text` text DEFAULT '' NOT NULL,
	`review_text` text DEFAULT '' NOT NULL,
	`updated_at` text NOT NULL
);
