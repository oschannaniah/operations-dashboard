-- Capacity-load snapshots — capacityForecast() in the frontend is computed live from whatever
-- is currently assigned; it has no memory of what someone's load looked like last month. A
-- trend line needs actual history, which only exists from the day this starts recording
-- forward — there's no way to retroactively reconstruct a past load score without knowing what
-- was assigned to who at each past date, which was never captured as events.
--
-- One row per staff member per snapshot run (a scheduled job, not per-request). Visibility
-- mirrors margin_scores exactly (OD of that campus + Central only, never the staff member
-- themselves) since capacityForecast()'s live badge in StaffPanel is already gated the same
-- way — this just persists what that badge already shows, it doesn't widen who sees it.

create table capacity_load_snapshots (
  id uuid primary key default gen_random_uuid(),
  organization_id uuid not null references organizations(id) on delete cascade,
  campus_id text not null references campuses(id),
  staff_id bigint not null references staff(id) on delete cascade,
  load_score numeric not null,
  status text, -- 'heavy_load' | 'over_capacity' | null, mirrors capacityForecast()'s two tiers
  snapshotted_at timestamptz not null default now()
);
create index capacity_load_snapshots_staff_idx on capacity_load_snapshots (staff_id, snapshotted_at desc);

alter table capacity_load_snapshots enable row level security;

create policy "central od read capacity snapshots" on capacity_load_snapshots for select
  using (organization_id = my_org() and (my_tier() = 'central' or (my_tier() = 'od' and campus_id = my_campus())));

-- Written only by the scheduled snapshot job via service_role — no client insert path.
grant select on capacity_load_snapshots to authenticated;
grant all on capacity_load_snapshots to service_role;
