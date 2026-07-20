-- Staff profile expansion: communication preferences, self-reported working-style results, a
-- role/campus assignment history, and a note field on check-ins.
--
-- staff_style_profile is deliberately self-write-only (via the same staff.user_id link pattern
-- as mobile_checkins) — Working Genius, Enneagram, and DISC results are self-reported by the
-- person taking OpsCore's own short in-house surveys (original content, not a reproduction of
-- any commercial/copyrighted assessment), so nobody else — not even their own OD — should be
-- able to overwrite what someone reported about themselves. OD/Central get read access, same
-- shape as everything else on a profile.
--
-- comm_preferences lives directly on staff (like flag, lastContact, etc. already do) and
-- inherits that table's existing, deliberately broad "campus writes own staff" policy
-- (0002_rls.sql) — any od/staff-tier person on the campus can technically write it, same as
-- every other staff column; the UI is what actually restricts editing to OD/Central. Not a new
-- gap, just consistent with how this table already works.
--
-- staff_assignment_history is written automatically whenever roles or campus change (from
-- existing OD/Central actions), not filled in by hand — same shape as staff_flag_history.
-- Like every other history table in this schema, it only starts accumulating from today
-- forward; there's no way to backfill assignment history that was never recorded as events.

alter table staff add column comm_preferences jsonb; -- ordered array, e.g. ["text","email"]

create table staff_style_profile (
  staff_id bigint primary key references staff(id) on delete cascade,
  organization_id uuid not null references organizations(id) on delete cascade,
  campus_id text not null references campuses(id),
  working_genius_geniuses jsonb, -- 2 of: Wonder, Invention, Discernment, Galvanizing, Enablement, Tenacity
  working_genius_competencies jsonb, -- 2 more, same list
  working_genius_frustrations jsonb, -- remaining 2, same list
  working_genius_completed_at timestamptz,
  enneagram_type smallint check (enneagram_type between 1 and 9),
  enneagram_wing smallint check (enneagram_wing between 1 and 9),
  enneagram_completed_at timestamptz,
  disc_primary text, -- 'D' | 'I' | 'S' | 'C'
  disc_secondary text,
  disc_completed_at timestamptz,
  updated_at timestamptz not null default now()
);

alter table staff_style_profile enable row level security;

create policy "self writes own style profile" on staff_style_profile for all
  using (exists (select 1 from staff s where s.id = staff_style_profile.staff_id and s.user_id = auth.uid()))
  with check (exists (select 1 from staff s where s.id = staff_style_profile.staff_id and s.user_id = auth.uid()));
create policy "od reads campus style profiles" on staff_style_profile for select
  using (my_tier() = 'od' and campus_id = my_campus());
create policy "central reads all style profiles" on staff_style_profile for select
  using (organization_id = my_org() and my_tier() = 'central');

grant select, insert, update, delete on staff_style_profile to authenticated;
grant all on staff_style_profile to service_role;

create table staff_assignment_history (
  id uuid primary key default gen_random_uuid(),
  organization_id uuid not null references organizations(id) on delete cascade,
  campus_id text not null references campuses(id),
  staff_id bigint not null references staff(id) on delete cascade,
  roles jsonb not null default '[]'::jsonb,
  changed_by text,
  changed_at timestamptz not null default now()
);
create index staff_assignment_history_staff_idx on staff_assignment_history (staff_id, changed_at desc);

alter table staff_assignment_history enable row level security;

create policy "central manages all assignment history" on staff_assignment_history for all
  using (organization_id = my_org() and my_tier() = 'central');
create policy "campus manages own assignment history" on staff_assignment_history for all
  using (organization_id = my_org() and my_tier() in ('od', 'staff') and campus_id = my_campus());

grant select, insert on staff_assignment_history to authenticated;
grant all on staff_assignment_history to service_role;

alter table staff_checkin_log add column note text;
