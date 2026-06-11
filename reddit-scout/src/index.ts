import "node:process";
import cron from "node-cron";
import { getAllSubreddits } from "./apps.config.js";
import { startStreams } from "./ingest/reddit.js";
import { handlePost, flushDigest } from "./pipeline.js";
import { runCheckIns } from "./monitor/circuit-breaker.js";
import { startBot } from "./approval/telegram.js";

const required = [
  "REDDIT_CLIENT_ID",
  "REDDIT_CLIENT_SECRET",
  "REDDIT_USERNAME",
  "REDDIT_PASSWORD",
  "REDDIT_USER_AGENT",
  "SUPABASE_URL",
  "SUPABASE_SERVICE_ROLE_KEY",
  "OPENROUTER_API_KEY",
  "TELEGRAM_BOT_TOKEN",
  "TELEGRAM_CHAT_ID",
];

for (const key of required) {
  if (!process.env[key]) {
    console.error(`Missing required env var: ${key}`);
    process.exit(1);
  }
}

console.log("[scout] Starting Reddit Scout…");

// ── Telegram approval bot (long-polling) ─────────────────────────────────────
startBot();

// ── Snoostorm: one SubmissionStream per subreddit ─────────────────────────────
// Replaces the manual 5-min cron + getNew() loop. Snoostorm handles polling
// and deduplication within a process run; Supabase handles cross-restart dedup.
const subreddits = getAllSubreddits();
const stopStreams = startStreams(subreddits, handlePost);

process.on("SIGTERM", () => {
  console.log("[scout] SIGTERM received — stopping streams");
  stopStreams();
  process.exit(0);
});

// ── Digest flush: every 2 hours ───────────────────────────────────────────────
// Medium-priority leads (score 7-8) accumulate in memory and are sent as a
// batch to avoid notification spam.
cron.schedule("0 */2 * * *", async () => {
  try {
    await flushDigest();
  } catch (err) {
    console.error("[cron:digest] error:", err);
  }
});

// ── Comment check-ins: daily at 09:00 UTC ────────────────────────────────────
// Updates 24h/7d scores and removal status; triggers circuit breaker if needed.
cron.schedule("0 9 * * *", async () => {
  console.log("[cron] Running check-ins…");
  try {
    await runCheckIns();
  } catch (err) {
    console.error("[cron:checkins] error:", err);
  }
});

console.log(`[scout] Streaming ${subreddits.length} subreddit(s). Waiting for posts…`);
