-- Scaffolding for per-user Asana and Slack connections, mirroring google_connections exactly
-- (0001_init.sql) — never selectable by a regular client, only the asana-connect/slack-connect
-- Edge Functions (service role) ever touch these. Genuinely inert until ASANA_CLIENT_ID/
-- ASANA_CLIENT_SECRET and SLACK_CLIENT_ID/SLACK_CLIENT_SECRET exist as real Supabase secrets —
-- every Edge Function call fails cleanly with a clear "not configured" message until then.

create table asana_connections (
  profile_id uuid primary key references profiles(id) on delete cascade,
  access_token text not null,
  refresh_token text,
  workspace_gid text,
  updated_at timestamptz not null default now()
);

create table slack_connections (
  profile_id uuid primary key references profiles(id) on delete cascade,
  access_token text not null,
  team_id text,
  team_name text,
  updated_at timestamptz not null default now()
);

alter table asana_connections enable row level security;
alter table slack_connections enable row level security;
-- No policies created = no access for authenticated/anon by default, same as google_connections.

grant all on asana_connections, slack_connections to service_role;
