import Snoowrap from "snoowrap";
import { SubmissionStream } from "snoostorm";
import type { RedditPost } from "../types.js";

let client: Snoowrap | null = null;

function getClient(): Snoowrap {
  if (client) return client;

  const required = [
    "REDDIT_CLIENT_ID",
    "REDDIT_CLIENT_SECRET",
    "REDDIT_USERNAME",
    "REDDIT_PASSWORD",
    "REDDIT_USER_AGENT",
  ];
  for (const key of required) {
    if (!process.env[key]) throw new Error(`${key} env var is required`);
  }

  client = new Snoowrap({
    userAgent: process.env.REDDIT_USER_AGENT!,
    clientId: process.env.REDDIT_CLIENT_ID!,
    clientSecret: process.env.REDDIT_CLIENT_SECRET!,
    username: process.env.REDDIT_USERNAME!,
    password: process.env.REDDIT_PASSWORD!,
  });

  client.config({ requestDelay: 1100, continueAfterRatelimitError: true });
  return client;
}

/**
 * Start one SubmissionStream per subreddit (Snoostorm handles polling).
 * Calls `onPost` for each new submission that passes the age gate.
 * Returns a cleanup function that removes all listeners.
 */
export function startStreams(
  subreddits: string[],
  onPost: (post: RedditPost, subreddit: string) => Promise<void>
): () => void {
  const r = getClient();
  const maxAgeHours = Number(process.env.MAX_POST_AGE_HOURS ?? 24);
  // Poll each subreddit every 60 s — stays well under Reddit's 60 req/min limit
  const pollTime = Number(process.env.REDDIT_POLL_MS ?? 60_000);

  const streams: SubmissionStream[] = [];

  for (const subreddit of subreddits) {
    const stream = new SubmissionStream(r, { subreddit, limit: 25, pollTime });

    stream.on("item", async (submission: Snoowrap.Submission) => {
      const cutoff = Date.now() / 1000 - maxAgeHours * 3600;
      if (submission.created_utc < cutoff) return;
      if (submission.author.name === process.env.REDDIT_USERNAME) return;

      const post: RedditPost = {
        id: submission.id,
        title: submission.title,
        selftext: submission.selftext ?? "",
        author: submission.author.name,
        subreddit_name_prefixed: submission.subreddit_name_prefixed,
        url: `https://reddit.com${submission.permalink}`,
        created_utc: submission.created_utc,
        score: submission.score,
        permalink: submission.permalink,
      };

      try {
        await onPost(post, subreddit);
      } catch (err) {
        console.error(`[stream:${subreddit}] handler error:`, err);
      }
    });

    streams.push(stream);
    console.log(`[ingest] Streaming r/${subreddit} (poll every ${pollTime / 1000}s)`);
  }

  return () => {
    for (const s of streams) s.removeAllListeners();
  };
}

/** Post a comment on a submission. Returns the permalink of the new comment. */
export async function postComment(
  submissionId: string,
  text: string
): Promise<string> {
  const r = getClient();
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const submission = r.getSubmission(submissionId) as any;
  const comment = await (submission.reply(text) as Promise<unknown>);
  return (comment as { permalink: string }).permalink;
}

/** Fetch the current karma score of a comment. */
export async function getCommentScore(commentId: string): Promise<number | null> {
  const r = getClient();
  try {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const comment = await ((r.getComment(commentId) as any).fetch() as Promise<unknown>);
    return (comment as { score: number }).score;
  } catch {
    return null;
  }
}

/** Returns true if the comment has been removed or deleted. */
export async function isCommentRemoved(commentId: string): Promise<boolean> {
  const r = getClient();
  try {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const comment = await ((r.getComment(commentId) as any).fetch() as Promise<unknown>);
    return (comment as { body: string }).body === "[removed]";
  } catch {
    return true;
  }
}
