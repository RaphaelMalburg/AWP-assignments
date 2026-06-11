import { randomUUID } from "node:crypto";
import { db } from "./client.js";
import type { Opportunity, OpportunityStatus, Draft, Post, SubredditState } from "../types.js";

// ── Opportunities ──────────────────────────────────────────────────────────

export function upsertOpportunity(
  data: Omit<Opportunity, "id" | "created_at">
): Opportunity | null {
  const existing = db
    .prepare("SELECT * FROM opportunities WHERE post_id = ?")
    .get(data.post_id) as Opportunity | undefined;

  if (existing) return null; // already tracked

  const id = randomUUID();
  db.prepare(`
    INSERT INTO opportunities
      (id, post_id, subreddit, app_id, title, body, author, url, status, llm_score, llm_reasoning)
    VALUES
      (@id, @post_id, @subreddit, @app_id, @title, @body, @author, @url, @status, @llm_score, @llm_reasoning)
  `).run({ id, ...data });

  return db
    .prepare("SELECT * FROM opportunities WHERE id = ?")
    .get(id) as Opportunity;
}

export function setOpportunityStatus(
  id: string,
  status: OpportunityStatus,
  extras?: Partial<Pick<Opportunity, "llm_score" | "llm_reasoning">>
): void {
  db.prepare(`
    UPDATE opportunities
    SET status = @status,
        llm_score     = COALESCE(@llm_score, llm_score),
        llm_reasoning = COALESCE(@llm_reasoning, llm_reasoning)
    WHERE id = @id
  `).run({ id, status, llm_score: extras?.llm_score ?? null, llm_reasoning: extras?.llm_reasoning ?? null });
}

export function getOpportunity(id: string): Opportunity | null {
  return (db.prepare("SELECT * FROM opportunities WHERE id = ?").get(id) as Opportunity) ?? null;
}

// ── Drafts ─────────────────────────────────────────────────────────────────

export function insertDraft(data: Omit<Draft, "id" | "created_at">): Draft {
  const id = randomUUID();
  db.prepare(`
    INSERT INTO drafts (id, opportunity_id, version, content, edited_by_human)
    VALUES (@id, @opportunity_id, @version, @content, @edited_by_human)
  `).run({ id, ...data, edited_by_human: data.edited_by_human ? 1 : 0 });

  return db.prepare("SELECT * FROM drafts WHERE id = ?").get(id) as Draft;
}

export function getLatestDraft(opportunityId: string): Draft | null {
  const row = db
    .prepare("SELECT * FROM drafts WHERE opportunity_id = ? ORDER BY version DESC LIMIT 1")
    .get(opportunityId) as Record<string, unknown> | undefined;

  if (!row) return null;
  return { ...(row as unknown as Draft), edited_by_human: Boolean(row.edited_by_human) };
}

// ── Posts ──────────────────────────────────────────────────────────────────

export function insertPost(data: Omit<Post, "id">): Post {
  const id = randomUUID();
  db.prepare(`
    INSERT INTO posts (id, opportunity_id, permalink, posted_at, comment_score_24h, comment_score_7d, removed, clicks)
    VALUES (@id, @opportunity_id, @permalink, @posted_at, @comment_score_24h, @comment_score_7d, @removed, @clicks)
  `).run({ id, ...data, removed: data.removed ? 1 : 0 });

  return db.prepare("SELECT * FROM posts WHERE id = ?").get(id) as Post;
}

export function updatePostScores(
  opportunityId: string,
  scores: { comment_score_24h?: number; comment_score_7d?: number; removed?: boolean }
): void {
  const sets: string[] = [];
  const params: Record<string, unknown> = { opportunity_id: opportunityId };

  if (scores.comment_score_24h !== undefined) {
    sets.push("comment_score_24h = @comment_score_24h");
    params.comment_score_24h = scores.comment_score_24h;
  }
  if (scores.comment_score_7d !== undefined) {
    sets.push("comment_score_7d = @comment_score_7d");
    params.comment_score_7d = scores.comment_score_7d;
  }
  if (scores.removed !== undefined) {
    sets.push("removed = @removed");
    params.removed = scores.removed ? 1 : 0;
  }

  if (sets.length === 0) return;
  db.prepare(`UPDATE posts SET ${sets.join(", ")} WHERE opportunity_id = @opportunity_id`).run(params);
}

/** Posts needing a 24h or 7d check-in */
export function getPostsDueForCheckIn(): (Post & { subreddit: string })[] {
  const now = Date.now();
  const h25 = new Date(now - 25 * 3600 * 1000).toISOString();
  const d8  = new Date(now - 8 * 24 * 3600 * 1000).toISOString();

  return db.prepare(`
    SELECT p.*, o.subreddit
    FROM posts p
    JOIN opportunities o ON o.id = p.opportunity_id
    WHERE p.removed = 0
      AND (
        (p.comment_score_24h IS NULL AND p.posted_at <= @h25)
        OR
        (p.comment_score_7d  IS NULL AND p.posted_at <= @d8)
      )
  `).all({ h25, d8 }) as (Post & { subreddit: string })[];
}

// ── Subreddit state ────────────────────────────────────────────────────────

export function getSubredditState(subreddit: string): SubredditState | null {
  return (
    (db.prepare("SELECT * FROM subreddit_state WHERE subreddit = ?").get(subreddit) as SubredditState) ?? null
  );
}

export function incrementSubredditRemovals(subreddit: string): void {
  db.prepare(`
    INSERT INTO subreddit_state (subreddit, removals_30d) VALUES (@subreddit, 1)
    ON CONFLICT(subreddit) DO UPDATE SET removals_30d = removals_30d + 1
  `).run({ subreddit });
}

export function setSubredditCooldown(subreddit: string, cooldownUntil: Date, notes: string): void {
  db.prepare(`
    INSERT INTO subreddit_state (subreddit, cooldown_until, notes) VALUES (@subreddit, @cooldown_until, @notes)
    ON CONFLICT(subreddit) DO UPDATE SET cooldown_until = @cooldown_until, notes = @notes
  `).run({ subreddit, cooldown_until: cooldownUntil.toISOString(), notes });
}

export function isSubredditOnCooldown(subreddit: string): boolean {
  const state = getSubredditState(subreddit);
  if (!state?.cooldown_until) return false;
  return new Date(state.cooldown_until) > new Date();
}

export function promoCountToday(subreddit: string, appId: string): number {
  const startOfDay = new Date();
  startOfDay.setUTCHours(0, 0, 0, 0);

  const row = db.prepare(`
    SELECT COUNT(*) as cnt
    FROM posts p
    JOIN opportunities o ON o.id = p.opportunity_id
    WHERE o.subreddit = @subreddit
      AND o.app_id = @app_id
      AND p.posted_at >= @start
  `).get({ subreddit, app_id: appId, start: startOfDay.toISOString() }) as { cnt: number };

  return row.cnt;
}

export function promoCountThisWeek(appId: string): number {
  const weekAgo = new Date(Date.now() - 7 * 24 * 3600 * 1000).toISOString();

  const row = db.prepare(`
    SELECT COUNT(*) as cnt
    FROM posts p
    JOIN opportunities o ON o.id = p.opportunity_id
    WHERE o.app_id = @app_id
      AND p.posted_at >= @week_ago
  `).get({ app_id: appId, week_ago: weekAgo }) as { cnt: number };

  return row.cnt;
}
