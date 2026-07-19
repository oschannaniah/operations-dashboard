-- Health Trends v1 — staff.flag and staff.last_contact are overwritten in place on every
-- change, so there was no way to see drift over time (a flag set once looks the same as a flag
-- set every week). These two append-only logs give Team Health a real history to read instead
-- of just a snapshot. Insert-only by design: authenticated gets select+insert, no update/delete,
-- so the trend can't be edited after the fact.

create table staff_flag_history (
  id uuid primary key default gen_random_uuid(),
  organization_id uuid not null references organizations(id) on delete cascade,
  campus_id text not null references campuses(id),
  staff_id bigint not null references staff(id) on delete cascade,
  flag text, -- null = cleared
  set_by text,
  set_at timestamptz not null default now()
);
create index staff_flag_history_staff_idx on staff_flag_history (staff_id, set_at desc);

create table staff_checkin_log (
  id uuid primary key default gen_random_uuid(),
  organization_id uuid not null references organizations(id) on delete cascade,
  campus_id text not null references campuses(id),
  staff_id bigint not null references staff(id) on delete cascade,
  logged_by text,
  logged_at timestamptz not null default now()
);
create index staff_checkin_log_staff_idx on staff_checkin_log (staff_id, logged_at desc);

alter table staff_flag_history enable row level security;
alter table staff_checkin_log enable row level security;

-- Same shape as staff's own policies (0002_rls.sql).
create policy "central manages all flag history" on staff_flag_history for all
  using (organization_id = my_org() and my_tier() = 'central');
create policy "campus manages own flag history" on staff_flag_history for all
  using (organization_id = my_org() and my_tier() in ('od', 'staff') and campus_id = my_campus());

create policy "central manages all checkin log" on staff_checkin_log for all
  using (organization_id = my_org() and my_tier() = 'central');
create policy "campus manages own checkin log" on staff_checkin_log for all
  using (organization_id = my_org() and my_tier() in ('od', 'staff') and campus_id = my_campus());

grant select, insert on staff_flag_history, staff_checkin_log to authenticated;
grant all on staff_flag_history, staff_checkin_log to service_role;
