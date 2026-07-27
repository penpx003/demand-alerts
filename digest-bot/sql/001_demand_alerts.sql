-- Migration 001: weekly Demand Planning alert snapshots + generated digests.
--
-- The `Demand Alerts` Office Script rebuilds four alert tables each week in the
-- DemandAlertsScripts workbook. A Power Automate flow POSTs them to
-- /api/demand-digest, which stores one snapshot row per alert row and keeps the
-- narrative it generated.
--
-- Snapshots exist so the digest can say "third week running" rather than
-- describing every week as if it were new.

create table if not exists demand_alert_snapshots (
  id          bigserial   primary key,
  week_of     date        not null,          -- Monday of the week the digest covers
  alert       text        not null,          -- alert1 | alert2 | alert3 | alert4
  agg_level   text        not null,          -- 'Product-Customer' | 'Product'
  market      text        not null default '',
  product     text        not null default '',
  customer    text        not null default '',
  bucket      text        not null default '',  -- future week label (Alert 1), else ''
  metric      double precision,              -- headline metric (variation %, accuracy diff, ...)
  volume      double precision,              -- headline volume
  alert_type  text,
  details     jsonb       not null default '{}',
  created_at  timestamptz not null default now()
);

alter table demand_alert_snapshots enable row level security;

-- Re-POSTing the same week must not duplicate rows (flows get retried).
create unique index if not exists demand_alert_snapshots_key_idx
  on demand_alert_snapshots (week_of, alert, agg_level, market, product, customer, bucket);

create index if not exists demand_alert_snapshots_week_idx
  on demand_alert_snapshots (week_of);
create index if not exists demand_alert_snapshots_entity_idx
  on demand_alert_snapshots (alert, agg_level, market, product, customer);

create table if not exists demand_digests (
  id          bigserial   primary key,
  week_of     date        not null unique,
  narrative   text        not null,
  stats       jsonb       not null default '{}',
  model       text,
  created_at  timestamptz not null default now()
);

alter table demand_digests enable row level security;

-- How many of the recent weeks each combination has been alerting.
-- PostgREST cannot GROUP BY, so this is exposed as an RPC.
create or replace function demand_alert_recurrence(
  p_alert text,
  p_level text,
  p_since date
)
returns table (
  market      text,
  product     text,
  customer    text,
  weeks_seen  bigint,
  first_seen  date,
  last_seen   date
)
language sql stable
as $$
  select market,
         product,
         customer,
         count(distinct week_of) as weeks_seen,
         min(week_of)            as first_seen,
         max(week_of)            as last_seen
  from demand_alert_snapshots
  where alert = p_alert
    and agg_level = p_level
    and week_of >= p_since
  group by market, product, customer
  having count(distinct week_of) > 1
  order by count(distinct week_of) desc, product
  limit 500;
$$;
