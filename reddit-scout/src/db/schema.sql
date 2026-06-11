-- Reddit Scout — Supabase schema
-- Run once against your project via the Supabase SQL editor or psql.

create table if not exists opportunities (
  id              uuid primary key default gen_random_uuid(),
  post_id         text not null unique,
  subreddit       text not null,
  app_id          text not null,
  title           text not null,
  body            text not null,
  author          text not null,
  url             text not null,
  status          text not null default 'ingested',
  llm_score       numeric(4,1),
  llm_reasoning   text,
  created_at      timestamptz not null default now()
);

create index if not exists opp_status_idx on opportunities (status);
create index if not exists opp_app_idx    on opportunities (app_id);

-- ─────────────────────────────────────────────────────────
create table if not exists drafts (
  id               uuid primary key default gen_random_uuid(),
  opportunity_id   uuid not null references opportunities(id) on delete cascade,
  version          int not null default 1,
  content          text not null,
  edited_by_human  boolean not null default false,
  created_at       timestamptz not null default now()
);

create index if not exists draft_opp_idx on drafts (opportunity_id);

-- ─────────────────────────────────────────────────────────
create table if not exists posts (
  id                   uuid primary key default gen_random_uuid(),
  opportunity_id       uuid not null references opportunities(id) on delete cascade,
  permalink            text not null,
  posted_at            timestamptz not null default now(),
  comment_score_24h    int,
  comment_score_7d     int,
  removed              boolean not null default false,
  clicks               int not null default 0
);

create index if not exists post_opp_idx on posts (opportunity_id);

-- ─────────────────────────────────────────────────────────
create table if not exists subreddit_state (
  subreddit       text primary key,
  cooldown_until  timestamptz,
  removals_30d    int not null default 0,
  notes           text
);

-- ─────────────────────────────────────────────────────────
-- Convenience view: pending approval queue
create or replace view pending_approvals as
  select
    o.id            as opportunity_id,
    o.post_id,
    o.subreddit,
    o.app_id,
    o.title,
    o.url,
    o.llm_score,
    o.llm_reasoning,
    d.id            as draft_id,
    d.content       as draft_content,
    d.version       as draft_version
  from opportunities o
  join drafts d on d.opportunity_id = o.id
  where o.status = 'pending_approval'
  order by o.llm_score desc, o.created_at asc;
