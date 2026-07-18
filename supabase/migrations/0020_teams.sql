-- Named Teams within Staff & Team — distinct from Team Lanes (the 4 fixed ministry
-- categories a role belongs to) and from Ministry Area (what a project belongs to). A Team
-- here is an ad-hoc, campus-created group (Worship Team, Guest Services...) that existing
-- staff members can be added to, each with their own role within that specific team — someone
-- can be "Team Lead" on one and just a member on another, and belong to more than one team at
-- once. This is the "Teams grouping" feature that was deliberately deferred earlier in the
-- project and is now being built for real.

create table teams (
  id uuid primary key default gen_random_uuid(),
  organization_id uuid not null references organizations(id) on delete cascade,
  campus_id text not null references campuses(id),
  name text not null,
  created_at timestamptz not null default now()
);

create table team_members (
  id uuid primary key default gen_random_uuid(),
  team_id uuid not null references teams(id) on delete cascade,
  staff_id bigint not null references staff(id) on delete cascade,
  role_in_team text,
  added_at timestamptz not null default now(),
  unique (team_id, staff_id)
);

alter table teams enable row level security;
alter table team_members enable row level security;

-- Same shape as staff's own policies (0002_rls.sql) — central manages every campus's teams,
-- od/staff manage only their own campus's.
create policy "central manages all teams" on teams for all
  using (organization_id = my_org() and my_tier() = 'central');
create policy "campus manages own teams" on teams for all
  using (organization_id = my_org() and my_tier() in ('od', 'staff') and campus_id = my_campus());

-- team_members has no campus_id of its own — scope is inherited by joining back to the parent
-- team, same pattern project_id-scoped tables would use if this app had any (it doesn't yet;
-- subtasks scope through projects the same conceptual way).
create policy "central manages all team_members" on team_members for all
  using (exists (select 1 from teams t where t.id = team_members.team_id and t.organization_id = my_org() and my_tier() = 'central'));
create policy "campus manages own team_members" on team_members for all
  using (exists (select 1 from teams t where t.id = team_members.team_id and t.organization_id = my_org() and my_tier() in ('od', 'staff') and t.campus_id = my_campus()));

grant select, insert, update, delete on teams, team_members to authenticated;
grant all on teams, team_members to service_role;
