import { APP_PROFILES, getProfile } from "./apps.config.js";
import { applyDeterministicFilter } from "./filters/deterministic.js";
import { scoreOpportunity } from "./llm/score.js";
import { draftReply } from "./llm/draft.js";
import {
  upsertOpportunity,
  setOpportunityStatus,
  insertDraft,
  isSubredditOnCooldown,
  promoCountToday,
  promoCountThisWeek,
} from "./db/queries.js";
import { sendForApproval, sendDigest } from "./approval/telegram.js";
import type { Opportunity, Draft, RedditPost } from "./types.js";

const SCORE_THRESHOLD = Number(process.env.SCORE_THRESHOLD ?? 7);
const MAX_PROMO_PER_DAY_PER_SUB = 1;
const MAX_PROMO_PER_WEEK_PER_APP = 5;

// Score ≥ 9 → sent immediately; 7-8 → batched into digest every 2h
const digestQueue: { opportunity: Opportunity; draft: Draft }[] = [];

/**
 * Entry point called by each Snoostorm "item" event.
 * Fully async — errors are caught internally so the stream stays alive.
 */
export async function handlePost(post: RedditPost, subreddit: string): Promise<void> {
  const onCooldown = await isSubredditOnCooldown(subreddit);
  if (onCooldown) return;

  // ── 1. Deterministic filter (zero LLM cost) ───────────────────────────────
  const filter = applyDeterministicFilter(post, APP_PROFILES);
  if (!filter.passed || !filter.appId) return;

  // ── 2. DB upsert — deduplication via unique post_id constraint ────────────
  const row = await upsertOpportunity({
    post_id: post.id,
    subreddit,
    app_id: filter.appId,
    title: post.title,
    body: post.selftext,
    author: post.author,
    url: post.url,
    status: "ingested",
    llm_score: null,
    llm_reasoning: null,
  });
  if (!row) return; // already processed

  const profile = getProfile(filter.appId)!;

  // ── 3. Score ───────────────────────────────────────────────────────────────
  let scoreResult;
  try {
    scoreResult = await scoreOpportunity(post, profile);
  } catch (err) {
    console.error(`[pipeline] Score failed for ${post.id}:`, err);
    await setOpportunityStatus(row.id, "filtered_out");
    return;
  }

  await setOpportunityStatus(row.id, "scored", {
    llm_score: scoreResult.score,
    llm_reasoning: scoreResult.reasoning,
  });

  if (scoreResult.score < SCORE_THRESHOLD) {
    await setOpportunityStatus(row.id, "filtered_out");
    console.log(`[pipeline] ${post.id} scored ${scoreResult.score} — below threshold`);
    return;
  }

  // ── 4. Rate limits ────────────────────────────────────────────────────────
  const todayCount = await promoCountToday(subreddit, filter.appId);
  if (todayCount >= MAX_PROMO_PER_DAY_PER_SUB) {
    await setOpportunityStatus(row.id, "filtered_out");
    console.log(`[pipeline] r/${subreddit} hit daily promo limit`);
    return;
  }

  const weekCount = await promoCountThisWeek(filter.appId);
  if (weekCount >= MAX_PROMO_PER_WEEK_PER_APP) {
    await setOpportunityStatus(row.id, "filtered_out");
    console.log(`[pipeline] ${filter.appId} hit weekly promo limit`);
    return;
  }

  // ── 5. Draft ───────────────────────────────────────────────────────────────
  let draftContent;
  try {
    draftContent = await draftReply(post, profile, subreddit);
  } catch (err) {
    console.error(`[pipeline] Draft failed for ${post.id}:`, err);
    return;
  }

  const draft = await insertDraft({
    opportunity_id: row.id,
    version: 1,
    content: draftContent,
    edited_by_human: false,
  });

  await setOpportunityStatus(row.id, "pending_approval");

  const opportunity: Opportunity = {
    ...row,
    status: "pending_approval",
    llm_score: scoreResult.score,
    llm_reasoning: scoreResult.reasoning,
  };

  // ── 6. Telegram routing ───────────────────────────────────────────────────
  if (scoreResult.score >= 9) {
    await sendForApproval(opportunity, draft);
    console.log(`[pipeline] ${post.id} → immediate approval (score ${scoreResult.score})`);
  } else {
    digestQueue.push({ opportunity, draft });
    console.log(`[pipeline] ${post.id} → digest queue (score ${scoreResult.score})`);
  }
}

/**
 * Flush the digest queue to Telegram. Called by the 2-hour cron in index.ts.
 * No-ops if there's nothing queued.
 */
export async function flushDigest(): Promise<void> {
  if (digestQueue.length === 0) return;
  const batch = digestQueue.splice(0);
  console.log(`[pipeline] Flushing digest — ${batch.length} item(s)`);
  await sendDigest(batch);
}
