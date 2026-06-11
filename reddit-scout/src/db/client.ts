import Database from "better-sqlite3";
import path from "node:path";
import fs from "node:fs";

const dbPath = process.env.DB_PATH ?? path.join(process.cwd(), "data", "scout.db");

// Ensure the directory exists (important on first Railway deploy)
fs.mkdirSync(path.dirname(dbPath), { recursive: true });

export const db = new Database(dbPath);

// WAL mode: allows reads while a write is happening — safer for cron + streams
db.pragma("journal_mode = WAL");
db.pragma("foreign_keys = ON");

// Auto-migrate on startup — no separate migrate step needed
db.exec(`
  CREATE TABLE IF NOT EXISTS opportunities (
    id            TEXT PRIMARY KEY,
    post_id       TEXT NOT NULL UNIQUE,
    subreddit     TEXT NOT NULL,
    app_id        TEXT NOT NULL,
    title         TEXT NOT NULL,
    body          TEXT NOT NULL,
    author        TEXT NOT NULL,
    url           TEXT NOT NULL,
    status        TEXT NOT NULL DEFAULT 'ingested',
    llm_score     REAL,
    llm_reasoning TEXT,
    created_at    TEXT NOT NULL DEFAULT (strftime('%Y-%m-%dT%H:%M:%SZ', 'now'))
  );

  CREATE INDEX IF NOT EXISTS opp_status_idx ON opportunities (status);
  CREATE INDEX IF NOT EXISTS opp_app_idx    ON opportunities (app_id);

  CREATE TABLE IF NOT EXISTS drafts (
    id               TEXT PRIMARY KEY,
    opportunity_id   TEXT NOT NULL REFERENCES opportunities(id) ON DELETE CASCADE,
    version          INTEGER NOT NULL DEFAULT 1,
    content          TEXT NOT NULL,
    edited_by_human  INTEGER NOT NULL DEFAULT 0,
    created_at       TEXT NOT NULL DEFAULT (strftime('%Y-%m-%dT%H:%M:%SZ', 'now'))
  );

  CREATE INDEX IF NOT EXISTS draft_opp_idx ON drafts (opportunity_id);

  CREATE TABLE IF NOT EXISTS posts (
    id                 TEXT PRIMARY KEY,
    opportunity_id     TEXT NOT NULL REFERENCES opportunities(id) ON DELETE CASCADE,
    permalink          TEXT NOT NULL,
    posted_at          TEXT NOT NULL DEFAULT (strftime('%Y-%m-%dT%H:%M:%SZ', 'now')),
    comment_score_24h  INTEGER,
    comment_score_7d   INTEGER,
    removed            INTEGER NOT NULL DEFAULT 0,
    clicks             INTEGER NOT NULL DEFAULT 0
  );

  CREATE INDEX IF NOT EXISTS post_opp_idx ON posts (opportunity_id);

  CREATE TABLE IF NOT EXISTS subreddit_state (
    subreddit      TEXT PRIMARY KEY,
    cooldown_until TEXT,
    removals_30d   INTEGER NOT NULL DEFAULT 0,
    notes          TEXT
  );
`);

console.log(`[db] SQLite ready at ${dbPath}`);
