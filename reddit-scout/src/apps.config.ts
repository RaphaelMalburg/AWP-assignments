import type { AppProfile } from "./types.js";

/**
 * Keywords are split into categories so the deterministic filter can be tuned
 * per signal strength and the scoring prompt gets richer context.
 * All categories are optional — just use `keywords` for a flat list if preferred.
 */
export interface KeywordCategories {
  /** "inventory app", "food cost" — direct product queries */
  direct?: string[];
  /** "running out of stock", "waste is killing me" — pain language */
  problem?: string[];
  /** "MarketMan alternative", "vs BlueCart" — comparison/switch intent */
  competitor?: string[];
}

export const APP_PROFILES: AppProfile[] = [
  {
    id: "kitchen",
    name: "KitchenOS",
    oneLiner: "Inventory and food-cost management for restaurant kitchens",
    problemsSolved: [
      "Can't track ingredient usage and keep running out of stock",
      "Food cost is too high and I don't know where the waste is",
      "No easy way to calculate recipe costs or update prices",
      "Manually counting inventory every week takes hours",
      "POS doesn't talk to my supplier invoices",
    ],
    subreddits: [
      "KitchenConfidential",
      "restaurantowners",
      "Chefit",
      "smallbusiness",
      "mealprep",
      "finedining",
      "foodtrucks",
    ],
    // Flat keyword list used by the deterministic filter
    keywords: [
      // direct
      "inventory app",
      "food cost",
      "track ingredients",
      "kitchen management",
      "waste tracking",
      "recipe costing",
      "food inventory",
      "inventory management",
      "menu costing",
      "ingredient tracking",
      "cost of goods",
      "stock management",
      // problem
      "running out of stock",
      "food waste",
      "how do you track",
      "how do you guys manage",
      "is there an app",
      "what app do you use",
      // comparison / switch intent
      "MarketMan alternative",
      "BlueCart alternative",
      "instead of MarketMan",
    ],
    negativeKeywords: [
      "hiring",
      "recipe request",
      "restaurant depot is great",
      "BlueCart is great",
      "MarketMan is great",
      "looking for a job",
      "quit my job",
    ],
    landingUrl: "https://kitchenos.app?utm_source=reddit&utm_medium=organic&utm_campaign=",
    voiceNotes: `
      - Speak as a fellow restaurant person, not a SaaS salesperson.
      - Lead with the actual answer to their question — be useful even if they never use the app.
      - Mention the app in the second or third paragraph, never the first sentence.
      - Always say "I built" or "I'm the founder of" — never pretend to be a random user.
      - Never use buzzwords like "streamline", "leverage", "empower", "game-changer".
      - Keep it under 120 words. No bullet lists unless the post is asking for comparisons.
      - If the subreddit bans self-promotion, mention the app name only — no link.
    `,
    selfPromoRules: {
      KitchenConfidential: "No self-promotion links; mention app name only without URL",
      restaurantowners: "Links allowed in comments if genuinely helpful",
      Chefit: "Light self-promo OK if you answer the question first",
      smallbusiness: "Links allowed; rule 4 requires disclosure",
      mealprep: "No promo posts; comment links generally tolerated if helpful",
    },
  },

  // ─────────────────────────────────────────────
  // App B — FILL IN BEFORE PRODUCTION
  // ─────────────────────────────────────────────
  {
    id: "app2",
    name: "App B",
    oneLiner: "TODO — fill in the one-liner for App B",
    problemsSolved: [
      "TODO — list the pains App B solves in the user's own words",
    ],
    subreddits: [],
    keywords: [],
    negativeKeywords: [],
    landingUrl: "https://app-b.example.com?utm_source=reddit&utm_medium=organic&utm_campaign=",
    voiceNotes: "TODO — describe the voice/tone and what never to say",
    selfPromoRules: {},
  },
];

export function getProfile(appId: string): AppProfile | undefined {
  return APP_PROFILES.find((p) => p.id === appId);
}

/** All subreddits across all apps, deduplicated */
export function getAllSubreddits(): string[] {
  const seen = new Set<string>();
  for (const profile of APP_PROFILES) {
    for (const sub of profile.subreddits) {
      seen.add(sub);
    }
  }
  return [...seen];
}
