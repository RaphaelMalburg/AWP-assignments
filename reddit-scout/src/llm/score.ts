import { chatCompletion } from "./client.js";
import type { AppProfile, RedditPost, ScoreResult } from "../types.js";

// Scoring weights mirror the Liftof/reddit-scout rubric
const SYSTEM_PROMPT = `You are a lead-qualification specialist for a SaaS product.
Read a Reddit post and score it as a lead opportunity.

Respond ONLY with a single valid JSON object — no markdown, no explanation, no code fences.

JSON schema:
{
  "score": <integer 0-10>,
  "appId": "<the app id from context>",
  "reasoning": "<one sentence>"
}

Scoring rubric (weights add to 100%):
- Relevance to the product (30%): Does the post topic directly match what the product does?
- User intent (25%): Is the person actively looking for a solution right now, or just venting?
- Recency signal (20%): Is this a fresh, unanswered question or an old thread with accepted answers?
- Engagement level (15%): Some comments already? That means the thread has visibility.
- Natural fit for a value-first reply (10%): Can we answer helpfully without it feeling like a pitch?

Score mapping:
10   = Perfect: actively asking for exactly this solution, no accepted answer yet
8-9  = Strong: clear pain point, open to recommendations, good thread visibility
6-7  = Moderate: tangentially relevant; could work with the right angle
3-5  = Weak: loosely related keyword match, not a real buying signal
0-2  = Not a lead: coincidental match, wrong context, or would look like spam`;

export async function scoreOpportunity(
  post: RedditPost,
  profile: AppProfile
): Promise<ScoreResult> {
  const model = process.env.OPENROUTER_SCORE_MODEL ?? "google/gemma-3-27b-it";

  const userMessage = `
App context:
- App ID: ${profile.id}
- App name: ${profile.name}
- What it does: ${profile.oneLiner}
- Problems it solves: ${profile.problemsSolved.join(" | ")}

Reddit post (${post.subreddit_name_prefixed}):
Title: ${post.title}
Body: ${post.selftext || "(no body — title only)"}
`.trim();

  const raw = await chatCompletion(model, [
    { role: "system", content: SYSTEM_PROMPT },
    { role: "user", content: userMessage },
  ]);

  try {
    const parsed = JSON.parse(raw) as ScoreResult;
    if (
      typeof parsed.score !== "number" ||
      typeof parsed.appId !== "string" ||
      typeof parsed.reasoning !== "string"
    ) {
      throw new Error("Missing fields");
    }
    parsed.score = Math.max(0, Math.min(10, Math.round(parsed.score)));
    return parsed;
  } catch {
    throw new Error(`scoreOpportunity: could not parse LLM response: ${raw}`);
  }
}
