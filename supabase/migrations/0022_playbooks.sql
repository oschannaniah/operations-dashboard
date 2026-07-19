-- Playbooks — reusable checklists (onboarding, opening procedures, season-launch) that get
-- "run" against a target and tracked to completion. Deliberately relational rather than jsonb
-- for the checklist items: runs get checked off incrementally, sometimes by different people,
-- and a per-row RLS-scoped item avoids the read-modify-write race a single jsonb array would
-- have if two people toggled different items around the same time.
--
-- A template is org-wide and Central-curated (od/staff can read and apply, not edit). A run
-- snapshots the template's items at apply-time into its own rows — editing or deleting the
-- template later never changes an in-flight run, same principle as a project's subtasks not
-- retroactively changing when some other project's template changes.

create table playbook_templates (
  id uuid primary key default gen_random_uuid(),
  organization_id uuid not null references organizations(id) on delete cascade,
  name text not null,
  type text not null, -- 'onboarding' (targets a staff member) | 'project' (targets a project) | 'standing' (no target, e.g. opening procedures)
  created_by text,
  created_at timestamptz not null default now()
);

create table playbook_template_items (
  id uuid primary key default gen_random_uuid(),
  template_id uuid not null references playbook_templates(id) on delete cascade,
  position int not null default 0,
  text text not null
);
create index playbook_template_items_template_idx on playbook_template_items (template_id, position);

create table playbook_runs (
  id uuid primary key default gen_random_uuid(),
  organization_id uuid not null references organizations(id) on delete cascade,
  campus_id text not null references campuses(id),
  template_id uuid references playbook_templates(id) on delete set null,
  template_name text not null, -- snapshot — survives the template being renamed or deleted
  type text not null,
  target_staff_id bigint references staff(id) on delete cascade,
  target_project_id bigint references projects(id) on delete cascade,
  started_by text,
  started_at timestamptz not null default now()
);
create index playbook_runs_campus_idx on playbook_runs (campus_id, started_at desc);

create table playbook_run_items (
  id uuid primary key default gen_random_uuid(),
  run_id uuid not null references playbook_runs(id) on delete cascade,
  position int not null default 0,
  text text not null,
  done boolean not null default false,
  done_by text,
  done_at timestamptz
);
create index playbook_run_items_run_idx on playbook_run_items (run_id, position);

alter table playbook_templates enable row level security;
alter table playbook_template_items enable row level security;
alter table playbook_runs enable row level security;
alter table playbook_run_items enable row level security;

-- Templates: Central curates the library; od/staff can only read it (to apply, not edit).
create policy "central manages playbook templates" on playbook_templates for all
  using (organization_id = my_org() and my_tier() = 'central');
create policy "campus reads playbook templates" on playbook_templates for select
  using (organization_id = my_org() and my_tier() in ('od', 'staff'));

create policy "central manages template items" on playbook_template_items for all
  using (exists (select 1 from playbook_templates t where t.id = playbook_template_items.template_id and t.organization_id = my_org() and my_tier() = 'central'));
create policy "campus reads template items" on playbook_template_items for select
  using (exists (select 1 from playbook_templates t where t.id = playbook_template_items.template_id and t.organization_id = my_org() and my_tier() in ('od', 'staff')));

-- Runs: same shape as staff/teams — central sees and manages every campus's runs, od/staff
-- manage runs and check off items only on their own campus.
create policy "central manages all playbook runs" on playbook_runs for all
  using (organization_id = my_org() and my_tier() = 'central');
create policy "campus manages own playbook runs" on playbook_runs for all
  using (organization_id = my_org() and my_tier() in ('od', 'staff') and campus_id = my_campus());

create policy "central manages all run items" on playbook_run_items for all
  using (exists (select 1 from playbook_runs r where r.id = playbook_run_items.run_id and r.organization_id = my_org() and my_tier() = 'central'));
create policy "campus manages own run items" on playbook_run_items for all
  using (exists (select 1 from playbook_runs r where r.id = playbook_run_items.run_id and r.organization_id = my_org() and my_tier() in ('od', 'staff') and r.campus_id = my_campus()));

grant select, insert, update, delete on playbook_templates, playbook_template_items, playbook_runs, playbook_run_items to authenticated;
grant all on playbook_templates, playbook_template_items, playbook_runs, playbook_run_items to service_role;
