import { Bot, InlineKeyboard } from "grammy";
import type { Opportunity, Draft } from "../types.js";
import {
  getOpportunity,
  getLatestDraft,
  setOpportunityStatus,
  insertDraft,
} from "../db/queries.js";
import { redraftWithFeedback } from "../llm/draft.js";
import { getProfile } from "../apps.config.js";
import { postComment } from "../ingest/reddit.js";
import { insertPost } from "../db/queries.js";

const CHAT_ID = process.env.TELEGRAM_CHAT_ID;
if (!CHAT_ID) throw new Error("TELEGRAM_CHAT_ID is not set");

let bot: Bot | null = null;

/** Lazy-init the bot instance (so import doesn't fail if token is missing at module load) */
function getBot(): Bot {
  if (bot) return bot;
  const token = process.env.TELEGRAM_BOT_TOKEN;
  if (!token) throw new Error("TELEGRAM_BOT_TOKEN is not set");
  bot = new Bot(token);
  registerHandlers(bot);
  return bot;
}

/**
 * Send a draft to the human approval queue in Telegram.
 * Messages contain the post link, score, reasoning, and the draft,
 * plus Approve / Edit / Discard inline buttons.
 */
export async function sendForApproval(
  opportunity: Opportunity,
  draft: Draft
): Promise<void> {
  const b = getBot();

  const keyboard = new InlineKeyboard()
    .text("✅ Approve", `approve:${opportunity.id}`)
    .text("✏️ Edit", `edit:${opportunity.id}`)
    .text("🗑 Discard", `discard:${opportunity.id}`);

  const score = opportunity.llm_score ?? "?";
  const scoreBar = buildScoreBar(Number(score));

  const text =
    `*[${opportunity.app_id.toUpperCase()}]* New lead — score ${score}/10 ${scoreBar}\n` +
    `📌 *${escMd(opportunity.title)}*\n` +
    `🔗 [View post](${opportunity.url})\n` +
    `_Reasoning: ${escMd(opportunity.llm_reasoning ?? "")}_ \n\n` +
    `*Draft reply:*\n${escMd(draft.content)}`;

  await b.api.sendMessage(CHAT_ID!, text, {
    parse_mode: "MarkdownV2",
    reply_markup: keyboard,
  });
}

/** Batch digest for lower-priority leads (score 7-8). Groups up to 5 at once. */
export async function sendDigest(items: { opportunity: Opportunity; draft: Draft }[]): Promise<void> {
  const b = getBot();
  const chunks = chunk(items, 5);

  for (const batch of chunks) {
    let text = `*Reddit Scout digest — ${batch.length} lead${batch.length > 1 ? "s" : ""}*\n\n`;
    for (const { opportunity, draft } of batch) {
      text +=
        `▸ *[${opportunity.app_id}]* score ${opportunity.llm_score}/10\n` +
        `  ${escMd(opportunity.title)}\n` +
        `  [View post](${opportunity.url})\n` +
        `  _Draft:_ ${escMd(draft.content.slice(0, 80))}…\n\n`;
    }

    await b.api.sendMessage(CHAT_ID!, text, { parse_mode: "MarkdownV2" });
  }
}

/** Send a plain alert (circuit breaker, posting success, etc.) */
export async function sendAlert(message: string): Promise<void> {
  const b = getBot();
  await b.api.sendMessage(CHAT_ID!, escMd(message), { parse_mode: "MarkdownV2" });
}

/** Register callback query handlers for the approval buttons. */
function registerHandlers(b: Bot): void {
  // ── Approve ──────────────────────────────────────────────────────────────
  b.callbackQuery(/^approve:(.+)$/, async (ctx) => {
    const opportunityId = ctx.match[1];
    await ctx.answerCallbackQuery("Posting…");

    try {
      const opp = await getOpportunity(opportunityId);
      if (!opp) return ctx.editMessageText("⚠️ Opportunity not found.");

      const draft = await getLatestDraft(opportunityId);
      if (!draft) return ctx.editMessageText("⚠️ No draft found.");

      const permalink = await postComment(`t3_${opp.post_id}`, draft.content);

      await setOpportunityStatus(opportunityId, "posted");
      await insertPost({
        opportunity_id: opportunityId,
        permalink,
        posted_at: new Date().toISOString(),
        comment_score_24h: null,
        comment_score_7d: null,
        removed: false,
        clicks: 0,
      });

      await ctx.editMessageText(
        `✅ Posted!\n🔗 https://reddit.com${permalink}\n\n${draft.content}`,
        { reply_markup: undefined }
      );
    } catch (err) {
      await ctx.editMessageText(`❌ Error posting: ${String(err)}`);
    }
  });

  // ── Discard ───────────────────────────────────────────────────────────────
  b.callbackQuery(/^discard:(.+)$/, async (ctx) => {
    const opportunityId = ctx.match[1];
    await setOpportunityStatus(opportunityId, "discarded");
    await ctx.answerCallbackQuery("Discarded");
    await ctx.editMessageText("🗑 Discarded.", { reply_markup: undefined });
  });

  // ── Edit — start conversation ─────────────────────────────────────────────
  const pendingEdits = new Map<string, string>(); // chatId → opportunityId

  b.callbackQuery(/^edit:(.+)$/, async (ctx) => {
    const opportunityId = ctx.match[1];
    pendingEdits.set(String(ctx.chat?.id), opportunityId);
    await ctx.answerCallbackQuery();
    await ctx.reply(
      "What should I change in the draft? (describe the edit, e.g. \"make the tone more casual\" or \"remove the link\")"
    );
  });

  b.on("message:text", async (ctx) => {
    const chatId = String(ctx.chat.id);
    const opportunityId = pendingEdits.get(chatId);
    if (!opportunityId) return;

    pendingEdits.delete(chatId);
    await ctx.reply("Regenerating…");

    try {
      const opp = await getOpportunity(opportunityId);
      if (!opp) return ctx.reply("⚠️ Opportunity not found.");

      const profile = getProfile(opp.app_id);
      if (!profile) return ctx.reply("⚠️ App profile not found.");

      const prevDraft = await getLatestDraft(opportunityId);
      if (!prevDraft) return ctx.reply("⚠️ No previous draft.");

      const fakePost = {
        id: opp.post_id,
        title: opp.title,
        selftext: opp.body,
        author: opp.author,
        subreddit_name_prefixed: `r/${opp.subreddit}`,
        url: opp.url,
        created_utc: 0,
        score: 0,
        permalink: "",
      };

      const newContent = await redraftWithFeedback(
        prevDraft.content,
        ctx.message.text,
        fakePost,
        profile,
        opp.subreddit
      );

      const newDraft = await insertDraft({
        opportunity_id: opportunityId,
        version: prevDraft.version + 1,
        content: newContent,
        edited_by_human: true,
      });

      const keyboard = new InlineKeyboard()
        .text("✅ Approve", `approve:${opportunityId}`)
        .text("✏️ Edit again", `edit:${opportunityId}`)
        .text("🗑 Discard", `discard:${opportunityId}`);

      await ctx.reply(`New draft (v${newDraft.version}):\n\n${newContent}`, {
        reply_markup: keyboard,
      });
    } catch (err) {
      await ctx.reply(`❌ Error regenerating: ${String(err)}`);
    }
  });
}

/** Start the bot (long-polling) — call this from your process entrypoint. */
export function startBot(): void {
  const b = getBot();
  b.start();
  console.log("[telegram] Bot started (long-polling)");
}

// ── Helpers ────────────────────────────────────────────────────────────────

function escMd(text: string): string {
  // Escape special MarkdownV2 characters
  return text.replace(/[_*[\]()~`>#+\-=|{}.!\\]/g, "\\$&");
}

function buildScoreBar(score: number): string {
  const filled = Math.round(score);
  return "█".repeat(filled) + "░".repeat(10 - filled);
}

function chunk<T>(arr: T[], size: number): T[][] {
  const out: T[][] = [];
  for (let i = 0; i < arr.length; i += size) {
    out.push(arr.slice(i, i + size));
  }
  return out;
}
