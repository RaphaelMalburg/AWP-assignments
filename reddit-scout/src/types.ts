export interface AppProfile {
  id: string;
  name: string;
  oneLiner: string;
  problemsSolved: string[];
  subreddits: string[];
  keywords: string[];
  negativeKeywords: string[];
  landingUrl: string;
  /** Tone guidance: what voice to use, what never to say */
  voiceNotes: string;
  /** Known self-promo rules per subreddit, e.g. { "r/smallbusiness": "links allowed in comments" } */
  selfPromoRules: Record<string, string>;
}

export type OpportunityStatus =
  | "ingested"
  | "filtered_out"
  | "scored"
  | "drafted"
  | "pending_approval"
  | "approved"
  | "posted"
  | "discarded"
  | "removed";

export interface Opportunity {
  id: string;
  post_id: string;
  subreddit: string;
  app_id: string;
  title: string;
  body: string;
  author: string;
  url: string;
  status: OpportunityStatus;
  llm_score: number | null;
  llm_reasoning: string | null;
  created_at: string;
}

export interface Draft {
  id: string;
  opportunity_id: string;
  version: number;
  content: string;
  edited_by_human: boolean;
  created_at: string;
}

export interface Post {
  id: string;
  opportunity_id: string;
  permalink: string;
  posted_at: string;
  comment_score_24h: number | null;
  comment_score_7d: number | null;
  removed: boolean;
  clicks: number;
}

export interface SubredditState {
  subreddit: string;
  cooldown_until: string | null;
  removals_30d: number;
  notes: string | null;
}

export interface ScoreResult {
  score: number;
  appId: string;
  reasoning: string;
}

export interface RedditPost {
  id: string;
  title: string;
  selftext: string;
  author: string;
  subreddit_name_prefixed: string;
  url: string;
  created_utc: number;
  score: number;
  permalink: string;
}
