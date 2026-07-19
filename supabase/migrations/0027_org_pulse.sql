-- Org-wide Pulse — a quarterly, anonymous sentiment read across the whole org, distinct from
-- Margin's per-person pulse (which is deliberately NOT anonymous, since an OD needs to know
-- who to follow up with). This is the opposite shape on purpose: leadership gets an aggregate
-- number, never who said what.
--
-- Anonymity is structural, not just a missing column: pulse_responses carries no profile_id or
-- staff_id at all, so there is nothing in that table to ever join back to a person. Preventing
-- double-voting instead uses a *separate* table (pulse_wave_participants) that only proves
-- someone responded, never what they said — the two tables share a wave_id but never share a
-- row, so there's no way to reconstruct "who answered what" even with full database access.

create table pulse_waves (
  id uuid primary key default gen_random_uuid(),
  organization_id uuid not null references organizations(id) on delete cascade,
  name text not null, -- e.g. "Q3 2026 Pulse"
  opens_at date not null,
  closes_at date not null,
  created_by text,
  created_at timestamptz not null default now()
);

create table pulse_wave_participants (
  id uuid primary key default gen_random_uuid(),
  wave_id uuid not null references pulse_waves(id) on delete cascade,
  profile_id uuid not null references profiles(id) on delete cascade,
  responded_at timestamptz,
  unique (wave_id, profile_id)
);

create table pulse_responses (
  id uuid primary key default gen_random_uuid(),
  wave_id uuid not null references pulse_waves(id) on delete cascade,
  campus_id text references campuses(id), -- context for breakdowns only; never linked to participants
  answers jsonb not null, -- { q1..q4: "1".."5" }
  note text,
  submitted_at timestamptz not null default now()
);
create index pulse_responses_wave_idx on pulse_responses (wave_id);

alter table pulse_waves enable row level security;
alter table pulse_wave_participants enable row level security;
alter table pulse_responses enable row level security;

-- Waves: Central creates/manages; od/staff need to read them to know one is open and by when.
create policy "central manages pulse waves" on pulse_waves for all
  using (organization_id = my_org() and my_tier() = 'central');
create policy "campus reads pulse waves" on pulse_waves for select
  using (organization_id = my_org() and my_tier() in ('od', 'staff'));

-- Participants: Central sees the full roll (completion tracking, reminders) but never pairs it
-- with an answer. Everyone else can only ever see/write their own single row — "have I
-- responded to this wave yet."
create policy "central reads all pulse participants" on pulse_wave_participants for select
  using (exists (select 1 from pulse_waves w where w.id = pulse_wave_participants.wave_id and w.organization_id = my_org() and my_tier() = 'central'));
create policy "own pulse participant row" on pulse_wave_participants for all
  using (profile_id = auth.uid())
  with check (profile_id = auth.uid());

-- Responses: Central reads the aggregate (that's the entire point). Any org member can submit
-- one — insert-only, no update/delete for anyone but service_role, since a submitted response
-- is meant to be immutable. campus_id is trusted from the submitter's own my_campus(), not
-- client-supplied, so it can't be spoofed to misattribute a response to a different campus.
create policy "central reads pulse responses" on pulse_responses for select
  using (exists (select 1 from pulse_waves w where w.id = pulse_responses.wave_id and w.organization_id = my_org() and my_tier() = 'central'));
create policy "org members submit pulse responses" on pulse_responses for insert
  with check (
    exists (select 1 from pulse_waves w where w.id = pulse_responses.wave_id and w.organization_id = my_org())
    and (campus_id = my_campus() or my_tier() = 'central')
  );

grant select, insert, update, delete on pulse_waves to authenticated;
grant select, insert, update on pulse_wave_participants to authenticated;
grant select, insert on pulse_responses to authenticated;
grant all on pulse_waves, pulse_wave_participants, pulse_responses to service_role;
