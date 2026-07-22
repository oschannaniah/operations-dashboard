-- Task Groups — a Central-curated list of named sub-groupings for organizing the tasks inside
-- ONE project (e.g. "Safety", "Admin", "Facilities"). This is distinct from Ministry Area, which
-- categorizes whole projects across the org — Task Groups cluster tasks *within* a single
-- project's list, the same role Asana's per-project "Sections" play. Central curates which
-- names exist (same governance pattern as Ministry Area/Custom Taxonomy); any tier creating or
-- editing a task picks from that list.

create table task_group_options (
  id uuid primary key default gen_random_uuid(),
  organization_id uuid not null references organizations(id) on delete cascade,
  name text not null,
  sort_order numeric not null default 0,
  created_at timestamptz not null default now(),
  created_by text
);

alter table task_group_options enable row level security;

create policy "central manages task group options" on task_group_options for all
  using (organization_id = my_org() and my_tier() = 'central');
create policy "campus reads task group options" on task_group_options for select
  using (organization_id = my_org() and my_tier() in ('od', 'staff'));

grant select, insert, update, delete on task_group_options to authenticated;
grant all on task_group_options to service_role;

-- Plain nullable text, not an FK — same lightweight pattern as projects.section against
-- MINISTRY_AREA_OPTIONS: validated client-side against task_group_options, not enforced in the
-- schema. No RLS change needed on subtasks itself; this just adds a column to an existing,
-- already-governed table.
alter table subtasks add column task_group text;
