-- OpsCore — Row Level Security policies.
--
-- Every policy checks organization_id first (tenant isolation — the most safety-critical
-- boundary, a bug here means one church seeing another's data), then tier/campus logic that
-- mirrors filterRowsForAuth_/canWriteRow_ from the old Code.gs. 'unassigned' tier gets no
-- policies at all (default-deny), matching `if (auth.tier === "unassigned") return [];`.
--
-- Design note on profiles: role/campus/tier changes (bootstrap, Team Accounts assigning
-- someone's campus+role, admin account creation) all go through Edge Functions using the
-- service role, NOT direct client writes — Postgres RLS is row-scoped, not column-scoped, so
-- a policy permissive enough to let Central update someone else's campus/role would need
-- column-level grants to stop a regular user self-elevating their own tier, and column grants
-- apply per-database-role rather than per-policy — which would end up blocking Central too.
-- Routing every sensitive profile write through a service-role Edge Function sidesteps that
-- entirely. Only self-service, non-sensitive fields (phone, name, calendar prefs) are
-- updatable directly by the row's own owner.

create or replace function my_org() returns uuid as $$
  select organization_id from profiles where id = auth.uid()
$$ language sql stable security definer;

create or replace function my_tier() returns text as $$
  select tier from profiles where id = auth.uid()
$$ language sql stable security definer;

create or replace function my_campus() returns text as $$
  select campus_id from profiles where id = auth.uid()
$$ language sql stable security definer;

create or replace function my_name() returns text as $$
  select trim(coalesce(first_name,'') || ' ' || coalesce(last_name,'')) from profiles where id = auth.uid()
$$ language sql stable security definer;

alter table organizations enable row level security;
alter table campuses enable row level security;
alter table profiles enable row level security;
alter table google_connections enable row level security;
alter table campus_config enable row level security;
alter table projects enable row level security;
alter table subtasks enable row level security;
alter table staff enable row level security;
alter table notifications enable row level security;
alter table central_threads enable row level security;

-- ---------- organizations ----------
-- Everyone can read their own org's row (name/slug display) — never anyone else's.
create policy "read own org" on organizations for select
  using (id = my_org());

-- ---------- campuses ----------
create policy "read campuses in own org" on campuses for select
  using (organization_id = my_org());

create policy "central manages campuses" on campuses for insert with check (organization_id = my_org() and my_tier() = 'central');
create policy "central updates campuses" on campuses for update using (organization_id = my_org() and my_tier() = 'central');
create policy "central deletes campuses" on campuses for delete using (organization_id = my_org() and my_tier() = 'central');

-- ---------- profiles ----------
create policy "central sees org profiles" on profiles for select
  using (organization_id = my_org() and my_tier() = 'central');
create policy "od sees own campus profiles" on profiles for select
  using (organization_id = my_org() and my_tier() = 'od' and campus_id = my_campus());
create policy "everyone sees own profile" on profiles for select
  using (id = auth.uid());

-- Self-service update, restricted to non-sensitive columns via column grants below.
create policy "update own profile" on profiles for update
  using (id = auth.uid()) with check (id = auth.uid());

revoke update on profiles from authenticated;
grant update (first_name, last_name, phone, google_calendar_ids, google_calendar_names) on profiles to authenticated;

-- ---------- google_connections ----------
-- Never selectable by a regular client at all — only Edge Functions (service role, which
-- bypasses RLS entirely) ever read or write this table.
-- (No policies created = no access for the "authenticated"/"anon" roles by default.)

-- ---------- campus_config ----------
create policy "read campus_config in own org" on campus_config for select
  using (exists (select 1 from campuses c where c.id = campus_config.campus_id and c.organization_id = my_org()));
create policy "od/central write own campus_config" on campus_config for all
  using (exists (select 1 from campuses c where c.id = campus_config.campus_id and c.organization_id = my_org()
                 and (my_tier() = 'central' or c.id = my_campus())));

-- ---------- projects ----------
create policy "central sees all org projects" on projects for select
  using (organization_id = my_org() and my_tier() = 'central');
create policy "campus sees own/shared/central projects" on projects for select
  using (organization_id = my_org() and my_tier() in ('od','staff')
         and (campus = my_campus() or campus = 'central'
              or (shared and (shared_with ? my_campus() or shared_with ? 'all'))));
create policy "staff further limited to assigned projects" on projects for select
  using (organization_id = my_org() and my_tier() = 'staff'
         and (owner = my_name() or created_by = my_name() or team ? my_name()));

create policy "central writes all org projects" on projects for all
  using (organization_id = my_org() and my_tier() = 'central');
create policy "od writes own campus projects" on projects for all
  using (organization_id = my_org() and my_tier() = 'od' and campus = my_campus());
create policy "staff writes own assigned projects" on projects for all
  using (organization_id = my_org() and my_tier() = 'staff' and campus = my_campus()
         and (owner = my_name() or created_by = my_name() or team ? my_name()));

-- ---------- subtasks ----------
-- Visibility/write access is entirely inherited from the parent project, exactly like
-- filterRowsForAuth_/canWriteRow_ did (subtasks carry no campus of their own).
create policy "subtasks follow parent project visibility" on subtasks for select
  using (exists (
    select 1 from projects p where p.id = subtasks.project_id
    and p.organization_id = my_org()
    and (my_tier() = 'central'
         or (my_tier() in ('od','staff') and (p.campus = my_campus() or p.campus = 'central'
             or (p.shared and (p.shared_with ? my_campus() or p.shared_with ? 'all'))))
    )
    and (my_tier() != 'staff' or created_by = my_name() or p.owner = my_name() or p.created_by = my_name() or p.team ? my_name())
  ));
create policy "subtasks follow parent project write access" on subtasks for all
  using (exists (
    select 1 from projects p where p.id = subtasks.project_id
    and p.organization_id = my_org()
    and (my_tier() = 'central' or (my_tier() in ('od','staff') and p.campus = my_campus()))
    and (my_tier() != 'staff' or created_by = my_name() or p.owner = my_name() or p.created_by = my_name() or p.team ? my_name())
  ));

-- ---------- staff ----------
create policy "central sees all org staff" on staff for select
  using (organization_id = my_org() and my_tier() = 'central');
create policy "campus sees own staff" on staff for select
  using (organization_id = my_org() and my_tier() in ('od','staff') and campus_id = my_campus());
create policy "central writes all org staff" on staff for all
  using (organization_id = my_org() and my_tier() = 'central');
create policy "campus writes own staff" on staff for all
  using (organization_id = my_org() and my_tier() in ('od','staff') and campus_id = my_campus());

-- ---------- notifications ----------
-- Always personal, regardless of tier — matches `forUser === auth.name` in the old code.
create policy "read own notifications" on notifications for select
  using (organization_id = my_org() and for_user_name = my_name());
create policy "insert notifications in own org" on notifications for insert
  with check (organization_id = my_org());
create policy "mark own notifications read" on notifications for update
  using (organization_id = my_org() and for_user_name = my_name());

-- ---------- central_threads ----------
-- Central-only, matches the old frontend's structural invisibility to campus-scoped views.
create policy "central reads/writes threads" on central_threads for all
  using (organization_id = my_org() and my_tier() = 'central');
