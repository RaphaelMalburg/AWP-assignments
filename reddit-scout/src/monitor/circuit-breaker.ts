import {
  getPostsDueForCheckIn,
  updatePostScores,
  incrementSubredditRemovals,
  setSubredditCooldown,
  getSubredditState,
  setOpportunityStatus,
} from "../db/queries.js";
import { getCommentScore, isCommentRemoved } from "../ingest/reddit.js";
import { sendAlert } from "../approval/telegram.js";

const REMOVAL_THRESHOLD = 2;       // removals before cooldown triggers
const COOLDOWN_DAYS = 60;
const WINDOW_DAYS = 30;

/**
 * Check-in on posted comments for score and removal status.
 * Called by a daily cron. Triggers circuit breaker when needed.
 */
export async function runCheckIns(): Promise<void> {
  const posts = await getPostsDueForCheckIn();
  if (posts.length === 0) return;

  console.log(`[circuit-breaker] Checking ${posts.length} posts…`);

  for (const post of posts) {
    // Derive comment id from permalink: last segment is comment id
    const segments = post.permalink.split("/").filter(Boolean);
    const commentId = `t1_${segments[segments.length - 1]}`;

    const removed = await isCommentRemoved(commentId);
    const score = await getCommentScore(commentId);

    const now = new Date();
    const postedAt = new Date(post.posted_at);
    const ageHours = (now.getTime() - postedAt.getTime()) / 3600000;

    const update: { comment_score_24h?: number; comment_score_7d?: number; removed?: boolean } = {};
    if (ageHours >= 25 && post.comment_score_24h === null) {
      update.comment_score_24h = score ?? 0;
    }
    if (ageHours >= 8 * 24 && post.comment_score_7d === null) {
      update.comment_score_7d = score ?? 0;
    }
    if (removed) {
      update.removed = true;
    }

    if (Object.keys(update).length > 0) {
      await updatePostScores(post.opportunity_id, update);
    }

    if (removed) {
      await setOpportunityStatus(post.opportunity_id, "removed");
      await incrementSubredditRemovals(post.subreddit);
      await checkCircuitBreaker(post.subreddit);
    }
  }
}

async function checkCircuitBreaker(subreddit: string): Promise<void> {
  const state = await getSubredditState(subreddit);
  const removals = state?.removals_30d ?? 0;

  if (removals >= REMOVAL_THRESHOLD) {
    const cooldownUntil = new Date(Date.now() + COOLDOWN_DAYS * 24 * 3600 * 1000);
    const notes = `Circuit breaker triggered: ${removals} removals in ${WINDOW_DAYS} days. Cooldown until ${cooldownUntil.toISOString().split("T")[0]}.`;
    await setSubredditCooldown(subreddit, cooldownUntil, notes);
    await sendAlert(
      `🚨 Circuit breaker triggered for r/${subreddit}!\n${removals} comments removed in 30 days.\nCooldown until ${cooldownUntil.toDateString()}.`
    );
    console.warn(`[circuit-breaker] Cooldown set for r/${subreddit}`);
  }
}
