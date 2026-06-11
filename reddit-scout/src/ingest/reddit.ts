import Snoowrap from "snoowrap";
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

  // Be conservative with rate limiting
  client.config({ requestDelay: 1100, continueAfterRatelimitError: true });
  return client;
}

/** Fetch up to `limit` new posts from a subreddit. */
export async function fetchNewPosts(
  subreddit: string,
  limit = 50
): Promise<RedditPost[]> {
  const r = getClient();
  const sub = r.getSubreddit(subreddit);
  const listing = await sub.getNew({ limit });

  const maxAgeHours = Number(process.env.MAX_POST_AGE_HOURS ?? 24);
  const cutoffEpoch = (Date.now() / 1000) - maxAgeHours * 3600;

  const posts: RedditPost[] = [];
  for (const post of listing) {
    if (post.created_utc < cutoffEpoch) continue;
    if (post.author.name === process.env.REDDIT_USERNAME) continue;

    posts.push({
      id: post.id,
      title: post.title,
      selftext: post.selftext ?? "",
      author: post.author.name,
      subreddit_name_prefixed: post.subreddit_name_prefixed,
      url: `https://reddit.com${post.permalink}`,
      created_utc: post.created_utc,
      score: post.score,
      permalink: post.permalink,
    });
  }

  return posts;
}

/** Post a comment on a submission by its full-name ID (e.g. "t3_abc123"). */
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

/** Fetch the current score of a comment by its full-name ID (e.g. "t1_abc123"). */
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

/**
 * Check whether a comment still exists (not deleted/removed).
 * Returns false if the API returns a 404 or the comment body is "[removed]".
 */
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
