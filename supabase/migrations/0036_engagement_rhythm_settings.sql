-- Engagement rhythm rubric — how consistently a leader (an OD, or Central for their own team)
-- is actually checking in / pulsing their people, scored against a cadence Central sets. Same
-- "one row per org" shape as capacity_weight_settings (0026): Central curates the target, every
-- tier needs read access since the score is computed client-side wherever a roster is shown,
-- same posture as capacity forecasting.
--
-- Two numbers, not brackets — this rubric is simpler than Capacity Weights by design: "days
-- since the last check-in or pulse" compared against a target and a grace period is the whole
-- model. target_days is the cadence Central wants (e.g. 30 = monthly); grace_days is how much
-- slack before "on track" becomes "overdue" rather than flipping the instant the target passes.

create table engagement_rhythm_settings (
  organization_id uuid primary key references organizations(id) on delete cascade,
  target_days numeric not null default 30,
  grace_days numeric not null default 14,
  updated_at timestamptz not null default now(),
  updated_by text
);

alter table engagement_rhythm_settings enable row level security;

create policy "central manages engagement rhythm settings" on engagement_rhythm_settings for all
  using (organization_id = my_org() and my_tier() = 'central');
create policy "campus reads engagement rhythm settings" on engagement_rhythm_settings for select
  using (organization_id = my_org() and my_tier() in ('od', 'staff'));

grant select, insert, update, delete on engagement_rhythm_settings to authenticated;
grant all on engagement_rhythm_settings to service_role;

insert into engagement_rhythm_settings (organization_id) values ('00000000-0000-0000-0000-000000000001')
  on conflict do nothing;
