-- Deleting a staff row today is a real DELETE, which cascades (on delete cascade) through every
-- history table tied to them — flag history, check-in log, assignment history, Margin scores/
-- surveys/pulses, mobile check-ins, capacity snapshots, working-style profile — permanently
-- destroying it. "Removing" someone from the roster should never do that: it becomes an
-- archive (a flag + timestamp), not a delete, so every table above keeps its foreign key intact
-- and the person's history survives for as long as the org needs to look back at it.
--
-- No RLS change needed — archiving is just an UPDATE on staff, already covered by the existing
-- "campus writes own staff" / "central writes all org staff" policies (0002_rls.sql).

alter table staff add column archived boolean not null default false;
alter table staff add column archived_at timestamptz;
alter table staff add column archived_by text;
create index staff_archived_idx on staff (campus_id, archived);
