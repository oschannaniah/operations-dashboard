-- OpsCore — initial schema migration (Sheets/Apps Script -> Supabase)
--
-- Design notes:
-- * Campus ids stay as short TEXT slugs ("abv", "central", ...) matching the existing
--   frontend convention (CAMPUSES const, the "central" sentinel checked all over the app) —
--   forcing UUID campus ids would mean rewriting every campus-id comparison in the frontend,
--   not just the data-access layer. A real "central" row is seeded below so the sentinel value
--   satisfies the foreign key instead of being a magic unchecked string.
-- * Projects/Subtasks/Staff ids stay as client-supplied bigint (Date.now()-based) — the
--   frontend renders optimistically using an id it generates itself before the write
--   round-trips, and rewriting that pattern is out of scope for this migration.
-- * Notifications/central_threads ids are DB-generated uuids instead — they're fire-and-forget
--   inserts with no optimistic-id-tracking need (see notes on notifyInvolved in the frontend).

create extension if not exists pgcrypto;

-- ---------- organizations & campuses ----------

create table organizations (
  id uuid primary key default gen_random_uuid(),
  name text not null,
  slug text unique not null,
  created_at timestamptz not null default now()
);

create table campuses (
  id text primary key,
  organization_id uuid not null references organizations(id) on delete cascade,
  name text not null,
  abbr text not null,
  lead_name text, -- display name, same "identity by string" convention as owner/createdBy
                   -- elsewhere — most campus ODs don't have a real login yet (only Han does
                   -- as of this migration), so this isn't forced into a hard FK.
  lead_profile_id uuid, -- optional: filled in once that OD actually registers/links an account
  phase int not null default 1,
  color text not null default '#2B4C7E',
  created_at timestamptz not null default now()
);

-- ---------- profiles (replaces Users) ----------

create table profiles (
  id uuid primary key references auth.users(id) on delete cascade,
  organization_id uuid not null references organizations(id) on delete cascade,
  first_name text,
  last_name text,
  email text not null,
  phone text,
  campus_id text references campuses(id),
  role text,
  tier text not null default 'unassigned',
  google_calendar_ids jsonb not null default '[]'::jsonb,
  google_calendar_names jsonb not null default '[]'::jsonb,
  created_at timestamptz not null default now()
);

alter table campuses
  add constraint campuses_lead_profile_id_fkey foreign key (lead_profile_id) references profiles(id);

-- Every new Supabase Auth user gets a profile row automatically, landing as tier
-- 'unassigned' (mirrors handleRegister_ in the old Code.gs) — organization_id and
-- everything else is filled in by whatever invite/signup flow the frontend uses.
create or replace function handle_new_auth_user() returns trigger as $$
begin
  insert into public.profiles (id, organization_id, email, tier)
  values (new.id, (new.raw_user_meta_data->>'organization_id')::uuid, new.email, 'unassigned');
  return new;
end;
$$ language plpgsql security definer;

create trigger on_auth_user_created
  after insert on auth.users
  for each row execute function handle_new_auth_user();

-- ---------- google_connections (split out from profiles for stricter RLS) ----------

create table google_connections (
  profile_id uuid primary key references profiles(id) on delete cascade,
  refresh_token text not null,
  granted_scopes text,
  updated_at timestamptz not null default now()
);

-- ---------- campus_config ----------

create table campus_config (
  campus_id text primary key references campuses(id) on delete cascade,
  slides_link text
);

-- ---------- projects / subtasks / staff ----------

create table projects (
  id bigint primary key,
  organization_id uuid not null references organizations(id) on delete cascade,
  campus text not null,
  title text not null,
  stage text not null default 'Pending',
  owner text,
  created_by text,
  created_at_str text, -- display-format date string the frontend already uses (TODAY_STR shape)
  completed_on text,
  team jsonb not null default '[]'::jsonb,
  collaborators jsonb not null default '[]'::jsonb,
  due text,
  due_time text,
  cost numeric not null default 0,
  spent numeric not null default 0,
  shared boolean not null default false,
  shared_with jsonb not null default '[]'::jsonb,
  section text not null default 'General',
  recurrence jsonb,
  photos jsonb not null default '[]'::jsonb,
  notes jsonb not null default '[]'::jsonb
);

create table subtasks (
  id bigint primary key,
  project_id bigint not null references projects(id) on delete cascade,
  t text not null,
  done boolean not null default false,
  due text,
  due_time text,
  created_by text,
  cost numeric not null default 0,
  spent numeric not null default 0,
  recurrence jsonb,
  photos jsonb not null default '[]'::jsonb,
  notes jsonb not null default '[]'::jsonb,
  assignees jsonb not null default '[]'::jsonb
);

create table staff (
  id bigint primary key,
  organization_id uuid not null references organizations(id) on delete cascade,
  campus_id text not null references campuses(id),
  name text not null,
  roles jsonb not null default '[]'::jsonb,
  reports_to text,
  next_meeting text,
  last_contact text,
  flag text,
  calendar_synced boolean not null default false,
  calendars jsonb not null default '[]'::jsonb,
  email text,
  phone text,
  lane text,
  user_id uuid references profiles(id)
);

-- ---------- new persistence: notifications & central_threads ----------

create table notifications (
  id uuid primary key default gen_random_uuid(),
  organization_id uuid not null references organizations(id) on delete cascade,
  for_user_name text not null,
  actor text not null,
  summary text not null,
  project_id bigint,
  ts text not null,
  read boolean not null default false
);

create table central_threads (
  id uuid primary key default gen_random_uuid(),
  organization_id uuid not null references organizations(id) on delete cascade,
  thread_key text not null,
  tags jsonb not null default '[]'::jsonb,
  messages jsonb not null default '[]'::jsonb,
  unique (organization_id, thread_key)
);
