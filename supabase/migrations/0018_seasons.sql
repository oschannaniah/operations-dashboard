-- Seasonal / Initiative model — a named, time-bound window (Easter, Christmas, VBS, ...) that
-- projects across every campus can be tagged into, so Central can see one season's full
-- cross-campus picture in a single view instead of stitching together eight campuses' project
-- boards by hand. Deliberately v1-scoped: create a season, tag projects to it, see it rolled
-- up. "Auto-suggest next year's checklist from last year's season" is a real feature but a
-- separate one, left for once there's enough real season history to template from.

create table seasons (
  id uuid primary key default gen_random_uuid(),
  organization_id uuid not null references organizations(id) on delete cascade,
  name text not null,
  starts_on date,
  ends_on date,
  created_by uuid references profiles(id),
  created_at timestamptz not null default now()
);

alter table projects add column season_id uuid references seasons(id);

alter table seasons enable row level security;

-- Central creates/manages seasons; everyone in the org can see what seasons exist (so a
-- campus OD tagging a project into "Easter 2027" has something to pick from) without being
-- able to create or edit seasons themselves — mirrors how campuses/campus_config already work.
create policy "central manages seasons" on seasons for all
  using (organization_id = my_org() and my_tier() = 'central');
create policy "org reads seasons" on seasons for select
  using (organization_id = my_org() and my_tier() in ('od', 'staff'));

grant select, insert, update, delete on seasons to authenticated;
grant all on seasons to service_role;
