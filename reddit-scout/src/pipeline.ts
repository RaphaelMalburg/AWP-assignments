import { APP_PROFILES, getAllSubreddits, getProfile } from "./apps.config.js";
import { fetchNewPosts } from "./ingest/reddit.js";
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
import type { Opportunity, Draft } from "./types.js";

const SCORE_THRESHOLD = Number(process.env.SCORE_THRESHOLD ?? 7);
const MAX_PROMO_PER_DAY_PER_SUB = 1;
const MAX_PROMO_PER_WEEK_PER_APP = 5;

// Separate digest queues: score 9-10 → immediate, score 7-8 → batch
const immediateQueue: { opportunity: Opportunity; draft: Draft }[] = [];
const digestQueue: { opportunity: Opportunity; draft: Draft }[] = [];

/**
 * One full ingest-filter-score-draft cycle.
 * Called by the 5-minute cron.
 */
export async function runPipeline(): Promise<void> {
  const subreddits = getAllSubreddits();
  console.log(`[pipeline] Checking ${subreddits.length} subreddits…`);

  for (const subreddit of subreddits) {
    const onCooldown = await isSubredditOnCooldown(subreddit);
    if (onCooldown) {
      console.log(`[pipeline] r/${subreddit} is on cooldown — skipping`);
      continue;
    }

    let posts;
    try {
      posts = await fetchNewPosts(subreddit);
    } catch (err) {
      console.error(`[pipeline] Failed to fetch r/${subreddit}:`, err);
      continue;
    }

    console.log(`[pipeline] r/${subreddit}: ${posts.length} posts fetched`);

    for (const post of posts) {
      await processPost(post, subreddit);
    }
  }

  // Send digest for medium-priority leads accumulated this cycle
  if (digestQueue.length > 0) {
    await sendDigest(digestQueue.splice(0));
  }

  // Immediate sends happen inline in processPost — drain any that queued
  for (const item of immediateQueue.splice(0)) {
    await sendForApproval(item.opportunity, item.draft);
  }
}

async function processPost(
  post: Parameters<typeof applyDeterministicFilter>[0],
  subreddit: string
): Promise<void> {
  // ── 1. Deterministic filter ───────────────────────────────────────────────
  const filter = applyDeterministicFilter(post, APP_PROFILES);
  if (!filter.passed || !filter.appId) return;

  // ── 2. DB upsert (deduplication) ──────────────────────────────────────────
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
  if (!row) return; // duplicate

  const profile = getProfile(filter.appId)!;

  // ── 3. Score ──────────────────────────────────────────────────────────────
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
    console.log(`[pipeline] Post ${post.id} scored ${scoreResult.score} — below threshold`);
    await setOpportunityStatus(row.id, "filtered_out");
    return;
  }

  // ── 4. Rate limits check ──────────────────────────────────────────────────
  const todayCount = await promoCountToday(subreddit, filter.appId);
  if (todayCount >= MAX_PROMO_PER_DAY_PER_SUB) {
    console.log(`[pipeline] r/${subreddit} hit daily promo limit`);
    await setOpportunityStatus(row.id, "filtered_out");
    return;
  }

  const weekCount = await promoCountThisWeek(filter.appId);
  if (weekCount >= MAX_PROMO_PER_WEEK_PER_APP) {
    console.log(`[pipeline] ${filter.appId} hit weekly promo limit`);
    await setOpportunityStatus(row.id, "filtered_out");
    return;
  }

  // ── 5. Draft ──────────────────────────────────────────────────────────────
  let draftContent;
  try {
    draftContent = await draftReply(post, profile, subreddit);
  } catch (err) {
    console.error(`[pipeline] Draft failed for ${post.id}:`, err);
    await setOpportunityStatus(row.id, "scored"); // leave in scored state
    return;
  }

  const draft = await insertDraft({
    opportunity_id: row.id,
    version: 1,
    content: draftContent,
    edited_by_human: false,
  });

  await setOpportunityStatus(row.id, "pending_approval");

  // ── 6. Queue for Telegram ─────────────────────────────────────────────────
  const opportunity: Opportunity = {
    ...row,
    status: "pending_approval",
    llm_score: scoreResult.score,
    llm_reasoning: scoreResult.reasoning,
  };

  if (scoreResult.score >= 9) {
    // High priority: send immediately
    await sendForApproval(opportunity, draft);
  } else {
    // Medium priority: batch into next digest
    digestQueue.push({ opportunity, draft });
  }

  console.log(`[pipeline] Post ${post.id} queued for approval (score ${scoreResult.score})`);
}
