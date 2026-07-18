-- OpsCore — base table-level privileges.
--
-- Postgres checks these BEFORE RLS is ever evaluated — a role with no GRANT at all on a table
-- gets "permission denied" regardless of how permissive its RLS policies are. Tables created
-- via a CLI-pushed migration (running as the postgres superuser) don't automatically pick up
-- the default grants Supabase's dashboard/Studio table editor would normally apply.
--
-- `anon` gets nothing — every real flow in this app (even self-registration) goes through
-- Supabase Auth endpoints or a service-role Edge Function, never a direct anon table read/write.

grant usage on schema public to anon, authenticated, service_role;

grant select, insert, update, delete on
  organizations, campuses, profiles, campus_config, projects, subtasks, staff, notifications, central_threads
  to authenticated;

grant all on
  organizations, campuses, profiles, google_connections, campus_config, projects, subtasks, staff, notifications, central_threads
  to service_role;
