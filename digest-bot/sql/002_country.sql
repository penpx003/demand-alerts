-- Migration 002: one digest per country per week.
--
-- The first design assumed a single Demand Planning scope: `demand_digests`
-- had a UNIQUE constraint on week_of alone. With one flow per country all
-- writing to the same service, the last country to run each week would
-- OVERWRITE every other country's digest, and recurrence lookups would mix
-- countries together as if they were one book of business.
--
-- Country is stored as a plain label supplied by the flow (e.g. 'ES', 'PT').
-- Existing rows get '' — a single-country installation keeps working unchanged.

-- ---------------------------------------------------------------- snapshots
alter table demand_alert_snapshots
  add column if not exists country text not null default '';

drop index if exists demand_alert_snapshots_key_idx;
create unique index if not exists demand_alert_snapshots_key_idx
  on demand_alert_snapshots (
    country, week_of, alert, agg_level, market, product, customer, bucket
  );

create index if not exists demand_alert_snapshots_country_week_idx
  on demand_alert_snapshots (country, week_of);

-- ------------------------------------------------------------------ digests
alter table demand_digests
  add column if not exists country text not null default '';

-- week_of was UNIQUE on its own; that is exactly the overwrite bug.
alter table demand_digests drop constraint if exists demand_digests_week_of_key;

create unique index if not exists demand_digests_country_week_idx
  on demand_digests (country, week_of);

-- ---------------------------------------------------------------------- RPC
-- Recurrence must be scoped to one country, otherwise "alerted in 3 of the last
-- 8 weeks" silently counts other countries' weeks.
drop function if exists demand_alert_recurrence(text, text, date);

create or replace function demand_alert_recurrence(
  p_alert text,
  p_level text,
  p_since date,
  p_country text default ''
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
    and country = p_country
  group by market, product, customer
  having count(distinct week_of) > 1
  order by count(distinct week_of) desc, product
  limit 500;
$$;
