-- Custom taxonomy — lets a business rename the two fixed vocabularies baked into this app
-- (the "Campus" noun, and the 5 Ministry Area category labels) from settings instead of a code
-- change. Deliberately a *display* relabeling only: ministry_area_labels maps the existing
-- stored values (Operations/Ministry/Next Gen/Programming/Central — unchanged in the database)
-- to whatever text Central wants shown instead, so nothing about how projects are actually
-- tagged or filtered ever has to migrate. Same "one row per org" shape as
-- capacity_weight_settings (0026).
--
-- Scope note: this covers the primary nav/page chrome and the Ministry Area field itself, not
-- every prose mention of "campus" scattered through descriptive copy elsewhere in the app, and
-- not role titles (Campus Operations Director etc.) or Playbooks' own separate category list —
-- those are out of scope for this pass.

create table taxonomy_settings (
  organization_id uuid primary key references organizations(id) on delete cascade,
  location_singular text not null default 'Campus',
  location_plural text not null default 'Campuses',
  ministry_area_field_label text not null default 'Ministry Area',
  ministry_area_labels jsonb not null default '{"Operations":"Operations","Ministry":"Ministry","Next Gen":"Next Gen","Programming":"Programming","Central":"Central"}'::jsonb,
  updated_at timestamptz not null default now(),
  updated_by text
);

alter table taxonomy_settings enable row level security;

create policy "central manages taxonomy settings" on taxonomy_settings for all
  using (organization_id = my_org() and my_tier() = 'central');
create policy "campus reads taxonomy settings" on taxonomy_settings for select
  using (organization_id = my_org() and my_tier() in ('od', 'staff'));

grant select, insert, update, delete on taxonomy_settings to authenticated;
grant all on taxonomy_settings to service_role;

insert into taxonomy_settings (organization_id) values ('00000000-0000-0000-0000-000000000001')
  on conflict do nothing;
