import { db } from "./client.js";
import type { Opportunity, OpportunityStatus, Draft, Post, SubredditState } from "../types.js";

// ── Opportunities ──────────────────────────────────────────────────────────

export async function upsertOpportunity(
  data: Omit<Opportunity, "id" | "created_at">
): Promise<Opportunity | null> {
  const { data: row, error } = await db
    .from("opportunities")
    .upsert(data, { onConflict: "post_id", ignoreDuplicates: true })
    .select()
    .single();

  if (error?.code === "23505") return null; // duplicate — already tracked
  if (error) throw new Error(`upsertOpportunity: ${error.message}`);
  return row as Opportunity;
}

export async function setOpportunityStatus(
  id: string,
  status: OpportunityStatus,
  extras?: Partial<Pick<Opportunity, "llm_score" | "llm_reasoning">>
): Promise<void> {
  const { error } = await db
    .from("opportunities")
    .update({ status, ...extras })
    .eq("id", id);

  if (error) throw new Error(`setOpportunityStatus: ${error.message}`);
}

export async function getOpportunity(id: string): Promise<Opportunity | null> {
  const { data, error } = await db
    .from("opportunities")
    .select("*")
    .eq("id", id)
    .single();

  if (error) throw new Error(`getOpportunity: ${error.message}`);
  return data as Opportunity;
}

// ── Drafts ─────────────────────────────────────────────────────────────────

export async function insertDraft(
  data: Omit<Draft, "id" | "created_at">
): Promise<Draft> {
  const { data: row, error } = await db
    .from("drafts")
    .insert(data)
    .select()
    .single();

  if (error) throw new Error(`insertDraft: ${error.message}`);
  return row as Draft;
}

export async function getLatestDraft(opportunityId: string): Promise<Draft | null> {
  const { data, error } = await db
    .from("drafts")
    .select("*")
    .eq("opportunity_id", opportunityId)
    .order("version", { ascending: false })
    .limit(1)
    .single();

  if (error && error.code !== "PGRST116") throw new Error(`getLatestDraft: ${error.message}`);
  return (data as Draft) ?? null;
}

// ── Posts ──────────────────────────────────────────────────────────────────

export async function insertPost(data: Omit<Post, "id">): Promise<Post> {
  const { data: row, error } = await db
    .from("posts")
    .insert(data)
    .select()
    .single();

  if (error) throw new Error(`insertPost: ${error.message}`);
  return row as Post;
}

export async function updatePostScores(
  opportunityId: string,
  scores: { comment_score_24h?: number; comment_score_7d?: number; removed?: boolean }
): Promise<void> {
  const { error } = await db
    .from("posts")
    .update(scores)
    .eq("opportunity_id", opportunityId);

  if (error) throw new Error(`updatePostScores: ${error.message}`);
}

/** Returns posts that need a 24h or 7d score check-in */
export async function getPostsDueForCheckIn(): Promise<(Post & { opportunity_id: string; subreddit: string })[]> {
  const now = new Date();
  const twentyFiveHoursAgo = new Date(now.getTime() - 25 * 60 * 60 * 1000).toISOString();
  const eightDaysAgo = new Date(now.getTime() - 8 * 24 * 60 * 60 * 1000).toISOString();

  const { data, error } = await db
    .from("posts")
    .select("*, opportunities!inner(subreddit)")
    .or(`and(comment_score_24h.is.null,posted_at.lte.${twentyFiveHoursAgo}),and(comment_score_7d.is.null,posted_at.lte.${eightDaysAgo})`)
    .eq("removed", false);

  if (error) throw new Error(`getPostsDueForCheckIn: ${error.message}`);
  return (data ?? []) as (Post & { opportunity_id: string; subreddit: string })[];
}

// ── Subreddit state ────────────────────────────────────────────────────────

export async function getSubredditState(subreddit: string): Promise<SubredditState | null> {
  const { data, error } = await db
    .from("subreddit_state")
    .select("*")
    .eq("subreddit", subreddit)
    .single();

  if (error && error.code !== "PGRST116") throw new Error(`getSubredditState: ${error.message}`);
  return (data as SubredditState) ?? null;
}

export async function incrementSubredditRemovals(subreddit: string): Promise<void> {
  const state = await getSubredditState(subreddit);
  const current = state?.removals_30d ?? 0;

  const { error } = await db
    .from("subreddit_state")
    .upsert({ subreddit, removals_30d: current + 1 }, { onConflict: "subreddit" });

  if (error) throw new Error(`incrementSubredditRemovals: ${error.message}`);
}

export async function setSubredditCooldown(
  subreddit: string,
  cooldownUntil: Date,
  notes: string
): Promise<void> {
  const { error } = await db
    .from("subreddit_state")
    .upsert({ subreddit, cooldown_until: cooldownUntil.toISOString(), notes }, { onConflict: "subreddit" });

  if (error) throw new Error(`setSubredditCooldown: ${error.message}`);
}

/** Returns true if the subreddit is under cooldown */
export async function isSubredditOnCooldown(subreddit: string): Promise<boolean> {
  const state = await getSubredditState(subreddit);
  if (!state?.cooldown_until) return false;
  return new Date(state.cooldown_until) > new Date();
}

/** How many promotional comments were posted today for this app in this subreddit */
export async function promoCountToday(subreddit: string, appId: string): Promise<number> {
  const startOfDay = new Date();
  startOfDay.setUTCHours(0, 0, 0, 0);

  const { count, error } = await db
    .from("posts")
    .select("id", { count: "exact" })
    .eq("opportunities.subreddit", subreddit)
    .eq("opportunities.app_id", appId)
    .gte("posted_at", startOfDay.toISOString());

  if (error) throw new Error(`promoCountToday: ${error.message}`);
  return count ?? 0;
}

/** How many promotional comments were posted this week for this app total */
export async function promoCountThisWeek(appId: string): Promise<number> {
  const weekAgo = new Date(Date.now() - 7 * 24 * 60 * 60 * 1000).toISOString();

  const { count, error } = await db
    .from("posts")
    .select("id", { count: "exact" })
    .gte("posted_at", weekAgo);

  if (error) throw new Error(`promoCountThisWeek: ${error.message}`);
  return count ?? 0;
}
