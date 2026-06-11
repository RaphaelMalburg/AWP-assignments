import type { AppProfile, RedditPost } from "../types.js";

interface FilterResult {
  passed: boolean;
  appId: string | null;
  reason: string;
}

/**
 * Zero-cost deterministic gate. Returns the first matching app profile or
 * rejects the post. The LLM is never called for posts that fail here.
 */
export function applyDeterministicFilter(
  post: RedditPost,
  profiles: AppProfile[],
  blocklist: Set<string> = new Set()
): FilterResult {
  if (blocklist.has(post.author.toLowerCase())) {
    return { passed: false, appId: null, reason: "author_blocklisted" };
  }

  const haystack = `${post.title} ${post.selftext}`.toLowerCase();

  for (const profile of profiles) {
    if (profile.subreddits.length === 0 || profile.keywords.length === 0) continue;

    const subredditName = post.subreddit_name_prefixed.replace(/^r\//, "");
    if (!profile.subreddits.map((s) => s.toLowerCase()).includes(subredditName.toLowerCase())) {
      continue;
    }

    const hasNegative = profile.negativeKeywords.some((kw) =>
      haystack.includes(kw.toLowerCase())
    );
    if (hasNegative) {
      return { passed: false, appId: profile.id, reason: "negative_keyword" };
    }

    const hasPositive = profile.keywords.some((kw) =>
      haystack.includes(kw.toLowerCase())
    );
    if (hasPositive) {
      return { passed: true, appId: profile.id, reason: "keyword_match" };
    }
  }

  return { passed: false, appId: null, reason: "no_keyword_match" };
}
