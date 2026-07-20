-- Mobile check-ins — a self-initiated, few-second pulse a frontline team member can fire off
-- from their phone between shifts, distinct from both Margin's OD-initiated pulse (not
-- anonymous, tied to a person, but pushed BY someone else) and Org Pulse (anonymous, quarterly,
-- Central-initiated). This is the missing self-serve leg: any time, by the person themselves,
-- visible to their own OD and Central so a "rough" answer can actually get a follow-up — same
-- visibility shape as Margin, not Org Pulse's anonymity.
--
-- staff_id/campus_id are never client-supplied — both are resolved server-side from the
-- submitter's own staff record (profiles.id = auth.uid() joined through staff.user_id), same
-- discipline as campus_id being trusted from my_campus() elsewhere in this schema. This stops
-- someone from ever submitting a check-in as a different person or a different campus.

create table mobile_checkins (
  id uuid primary key default gen_random_uuid(),
  organization_id uuid not null references organizations(id) on delete cascade,
  campus_id text not null references campuses(id),
  staff_id bigint not null references staff(id) on delete cascade,
  profile_id uuid not null references profiles(id) on delete cascade,
  mood smallint not null check (mood between 1 and 5),
  energy smallint not null check (energy between 1 and 5),
  pace smallint not null check (pace between 1 and 5),
  support smallint not null check (support between 1 and 5),
  note text,
  created_at timestamptz not null default now()
);
create index mobile_checkins_staff_idx on mobile_checkins (staff_id, created_at desc);
create index mobile_checkins_campus_idx on mobile_checkins (campus_id, created_at desc);

alter table mobile_checkins enable row level security;

create policy "self reads own checkins" on mobile_checkins for select
  using (profile_id = auth.uid());
create policy "od reads campus checkins" on mobile_checkins for select
  using (my_tier() = 'od' and campus_id = my_campus());
create policy "central reads all checkins" on mobile_checkins for select
  using (organization_id = my_org() and my_tier() = 'central');

create policy "self submits own checkin" on mobile_checkins for insert
  with check (
    organization_id = my_org()
    and profile_id = auth.uid()
    and exists (
      select 1 from staff s
      where s.id = staff_id
        and s.user_id = auth.uid()
        and s.campus_id = mobile_checkins.campus_id
        and s.organization_id = mobile_checkins.organization_id
    )
  );

-- No update/delete policy for anyone but service_role — a submitted check-in is immutable,
-- same posture as pulse_responses.
grant select, insert on mobile_checkins to authenticated;
grant all on mobile_checkins to service_role;

-- Push subscriptions — one row per browser/device a person has opted into notifications on.
-- Strictly self-owned: nobody, including Central, ever reads another person's raw subscription
-- (endpoint/keys are bearer credentials for sending them a push). Only service_role, which
-- sends the actual notifications, needs broad access.
create table push_subscriptions (
  id uuid primary key default gen_random_uuid(),
  profile_id uuid not null references profiles(id) on delete cascade,
  endpoint text not null unique,
  p256dh text not null,
  auth_key text not null,
  created_at timestamptz not null default now()
);
create index push_subscriptions_profile_idx on push_subscriptions (profile_id);

alter table push_subscriptions enable row level security;

create policy "own push subscription row" on push_subscriptions for all
  using (profile_id = auth.uid())
  with check (profile_id = auth.uid());

grant select, insert, delete on push_subscriptions to authenticated;
grant all on push_subscriptions to service_role;
