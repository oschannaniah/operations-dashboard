-- Capacity forecasting v2 — replaces the crude "3+ items due soon" count with a real weighted
-- load score per person: urgency (days to due) x cost bracket x project-type weight, divided by
-- how many people are actually on the project (shared load counts less per person). All the
-- multipliers are Central-configurable rather than hardcoded, since the right thresholds are an
-- operational judgment call that will need tuning against real data, not something to bake into
-- code and require a rebuild to adjust.

alter table projects add column project_type text not null default 'General';

-- One settings row per org (not per campus — the weighting model is an org-wide policy).
-- JSONB for the bracket lists keeps this to one row instead of several small tables; brackets
-- are evaluated ascending by their "max" key, first match wins, and a null max means "and up."
create table capacity_weight_settings (
  organization_id uuid primary key references organizations(id) on delete cascade,
  type_weights jsonb not null default '{"Event":1,"Facilities":1,"Administrative":1,"Strategic Initiative":1.3,"General":1}'::jsonb,
  cost_brackets jsonb not null default '[{"maxCost":1000,"weight":1},{"maxCost":5000,"weight":1.5},{"maxCost":null,"weight":2}]'::jsonb,
  urgency_brackets jsonb not null default '[{"maxDays":0,"weight":3},{"maxDays":7,"weight":2},{"maxDays":30,"weight":1.5},{"maxDays":null,"weight":1}]'::jsonb,
  heavy_load_threshold numeric not null default 3,
  over_capacity_threshold numeric not null default 6,
  updated_at timestamptz not null default now(),
  updated_by text
);

alter table capacity_weight_settings enable row level security;

-- Same shape as playbook_templates: Central curates the weighting policy, od/staff need read
-- access since the forecast itself is computed client-side wherever a roster is shown.
create policy "central manages weight settings" on capacity_weight_settings for all
  using (organization_id = my_org() and my_tier() = 'central');
create policy "campus reads weight settings" on capacity_weight_settings for select
  using (organization_id = my_org() and my_tier() in ('od', 'staff'));

grant select, insert, update, delete on capacity_weight_settings to authenticated;
grant all on capacity_weight_settings to service_role;

insert into capacity_weight_settings (organization_id) values ('00000000-0000-0000-0000-000000000001')
  on conflict do nothing;
