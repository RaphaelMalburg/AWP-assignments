import { chatCompletion } from "./client.js";
import type { AppProfile, RedditPost, ScoreResult } from "../types.js";

const SYSTEM_PROMPT = `You are a lead-qualification specialist for a SaaS app.
Your job is to read a Reddit post and determine how strong a lead it is.

Respond ONLY with a single valid JSON object — no markdown, no explanation, no code fences.

JSON schema:
{
  "score": <integer 0-10>,
  "appId": "<string — the app id provided in the context>",
  "reasoning": "<one sentence explaining the score>"
}

Scoring guide:
10 = Person is explicitly searching for exactly this type of solution right now
8-9 = Clear pain point that the app solves; they're open to recommendations
6-7 = Tangentially relevant; might be interested if approached right
3-5 = Loosely related topic but not a real lead
0-2 = Not a lead; keyword match was coincidental`;

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

Reddit post from r/${post.subreddit_name_prefixed}:
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
