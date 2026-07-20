-- Reconsiders 0032's "service_role only" write posture: recomputing capacityForecast()
-- server-side would mean porting its weighted formula into a second implementation (SQL),
-- which drifts from the JS version the moment either one changes — a real risk, not a
-- hypothetical one, in a formula that's already been tuned once this session. capacityForecast
-- is already documented (0026) as "computed client-side wherever a roster is shown" and trusted
-- there; letting od/central write their own campus's snapshot directly extends that same trust
-- level rather than introducing a new one. It does NOT change who can ever READ a snapshot —
-- 0032's select policy (OD own campus + Central) is untouched.
--
-- snapshot_date + a unique constraint turns repeat writes from the same day (e.g. two different
-- ODs opening the roster) into one upsert instead of duplicate rows, so a trend chart doesn't
-- have to de-dupe client-side.

alter table capacity_load_snapshots add column snapshot_date date not null default current_date;
alter table capacity_load_snapshots add constraint capacity_load_snapshots_staff_date_key unique (staff_id, snapshot_date);

create policy "od central write own scope capacity snapshots" on capacity_load_snapshots for insert
  with check (organization_id = my_org() and (my_tier() = 'central' or (my_tier() = 'od' and campus_id = my_campus())));

create policy "od central update own scope capacity snapshots" on capacity_load_snapshots for update
  using (organization_id = my_org() and (my_tier() = 'central' or (my_tier() = 'od' and campus_id = my_campus())))
  with check (organization_id = my_org() and (my_tier() = 'central' or (my_tier() = 'od' and campus_id = my_campus())));

grant insert, update on capacity_load_snapshots to authenticated;
