import "node:process";
import cron from "node-cron";
import { runPipeline } from "./pipeline.js";
import { runCheckIns } from "./monitor/circuit-breaker.js";
import { startBot } from "./approval/telegram.js";

// Validate required env vars on startup
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

// Start the Telegram approval bot (long-polling)
startBot();

// Ingest + filter + score + draft: every 5 minutes
cron.schedule("*/5 * * * *", async () => {
  console.log("[cron] Running pipeline…");
  try {
    await runPipeline();
  } catch (err) {
    console.error("[cron] Pipeline error:", err);
  }
});

// Check-in on posted comments: every day at 09:00 UTC
cron.schedule("0 9 * * *", async () => {
  console.log("[cron] Running check-ins…");
  try {
    await runCheckIns();
  } catch (err) {
    console.error("[cron] Check-in error:", err);
  }
});

// Run pipeline immediately on startup (helps during development)
runPipeline().catch((err) => console.error("[startup] Pipeline error:", err));

console.log("[scout] Crons scheduled. Waiting for events…");
