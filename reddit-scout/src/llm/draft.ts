import { chatCompletion } from "./client.js";
import type { AppProfile, RedditPost } from "../types.js";

function buildSystemPrompt(profile: AppProfile, subreddit: string): string {
  const promoRule = profile.selfPromoRules[subreddit] ?? "Links allowed in comments if genuinely helpful";

  return `You are writing a Reddit comment on behalf of a founder who built ${profile.name}.

Rules (non-negotiable):
1. Start by genuinely answering or helping with the question — useful even if they never use the app.
2. Only in the second or third paragraph, mention ${profile.name} as one option. Use "I built ${profile.name}" or "I'm the founder of ${profile.name}".
3. Hard limit: 120 words total. Count them.
4. Never use: "streamline", "leverage", "empower", "game-changer", "revolutionize", "seamlessly", "robust".
5. No bullet lists unless the post is asking for a comparison of tools.
6. Self-promo rule for r/${subreddit}: ${promoRule}
7. Sound like a person, not a press release.

Voice notes for this app:
${profile.voiceNotes.trim()}

Respond with ONLY the comment text — no preamble, no quotation marks, no code fences.`;
}

export async function draftReply(
  post: RedditPost,
  profile: AppProfile,
  subreddit: string
): Promise<string> {
  const model = process.env.OPENROUTER_DRAFT_MODEL ?? "google/gemma-3-27b-it";

  const userMessage = `
Reddit post from r/${subreddit}:
Title: ${post.title}
Body: ${post.selftext || "(no body — title only)"}

App: ${profile.name}
What it does: ${profile.oneLiner}
Problems it solves: ${profile.problemsSolved.join("; ")}
Landing URL: ${profile.landingUrl}${subreddit}

Write the reply comment now.
`.trim();

  return chatCompletion(
    model,
    [
      { role: "system", content: buildSystemPrompt(profile, subreddit) },
      { role: "user", content: userMessage },
    ],
    0.7
  );
}

/** Regenerate a draft given human feedback on what to change. */
export async function redraftWithFeedback(
  originalDraft: string,
  humanFeedback: string,
  post: RedditPost,
  profile: AppProfile,
  subreddit: string
): Promise<string> {
  const model = process.env.OPENROUTER_DRAFT_MODEL ?? "google/gemma-3-27b-it";

  const userMessage = `
Original Reddit post from r/${subreddit}:
Title: ${post.title}
Body: ${post.selftext || "(no body)"}

Previous draft:
${originalDraft}

Human feedback (what to change):
${humanFeedback}

Please rewrite the comment applying the feedback. Keep all original rules (120 words max, disclose "I built", genuine helpfulness first).
`.trim();

  return chatCompletion(
    model,
    [
      { role: "system", content: buildSystemPrompt(profile, subreddit) },
      { role: "user", content: userMessage },
    ],
    0.7
  );
}
