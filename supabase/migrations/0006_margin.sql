-- Margin v1 — per-person workload signal, deliberately separate from Budget, visible only to
-- Central and each team member's own Campus Operations Director, never the team member
-- themselves. Two instruments feed it: a 10-question survey the OD completes about each team
-- member (sets the benchmark), and a 4-5 question pulse the OD can push to a team member any
-- time, answered by that person, which blends into the score. Landing on "over_capacity" fires
-- a notification to that person's OD carrying a check-in script.
--
-- Why three tables instead of columns on staff: the existing "campus writes own staff" policy
-- (0002_rls.sql) grants BOTH 'od' and 'staff' tier read/write on a campus's staff rows, so
-- margin data can't live there without leaking scores to the team member and their peers. It
-- needs its own tables with RLS scoped to 'od'/'central' only — no policy at all for 'staff'
-- tier on margin_scores/margin_surveys, same default-deny shape 'unassigned' gets elsewhere.
--
-- The one place a 'staff'-tier write is genuinely needed is answering their own pending pulse.
-- That reuses the exact column-vs-row pattern already solved for profiles: revoke the general
-- update grant, grant back only the specific columns they're allowed to touch, paired with a
-- row policy that only matches their own row while it's still pending.

create table margin_scores (
  staff_id bigint primary key references staff(id) on delete cascade,
  organization_id uuid not null references organizations(id) on delete cascade,
  campus_id text not null references campuses(id),
  score numeric not null,
  status text not null, -- 'comfortable' | 'full' | 'stretched' | 'over_capacity'
  calibration_gap boolean not null default false, -- Q9/Q10 disagreed sharply with the computed score
  last_survey_at timestamptz,
  last_pulse_at timestamptz,
  updated_at timestamptz not null default now()
);

create table margin_surveys (
  id uuid primary key default gen_random_uuid(),
  organization_id uuid not null references organizations(id) on delete cascade,
  campus_id text not null references campuses(id),
  staff_id bigint not null references staff(id) on delete cascade,
  od_profile_id uuid not null references profiles(id),
  answers jsonb not null, -- { q1..q10: <answer key> }
  score numeric, -- base score (Q1-8 average) at time of submission, kept for history
  created_at timestamptz not null default now()
);
create index margin_surveys_staff_idx on margin_surveys (staff_id, created_at desc);

create table margin_pulses (
  id uuid primary key default gen_random_uuid(),
  organization_id uuid not null references organizations(id) on delete cascade,
  campus_id text not null references campuses(id),
  staff_id bigint not null references staff(id) on delete cascade,
  sent_by uuid not null references profiles(id),
  sent_at timestamptz not null default now(),
  status text not null default 'pending', -- 'pending' | 'answered'
  answers jsonb, -- { p1..p3: <answer key>, p4: bool, p5: free text } — null until answered
  score numeric, -- pulse-only sub-score at time of response
  responded_at timestamptz
);
create index margin_pulses_staff_idx on margin_pulses (staff_id, status);

-- ---------- scoring ----------

-- One lookup table for every fixed-choice answer this feature uses, survey and pulse alike —
-- keeps the mapping in one place instead of duplicated across CASE expressions.
create or replace function margin_answer_points(q text, a text) returns numeric as $$
  select case q
    when 'q1' then case a when 'fewer' then 100 when 'typical' then 75 when 'more' then 40 when 'significantly_more' then 15 else null end
    when 'q2' then case a when 'down' then 100 when 'steady' then 70 when 'up' then 35 else null end
    when 'q3' then case a when 'well_ahead' then 100 when 'on_time' then 75 when 'occasionally_late' then 45 when 'often_late' then 15 else null end
    when 'q4' then case a when 'mostly_routine' then 100 when 'mixed' then 60 when 'mostly_high_stakes' then 25 else null end
    when 'q5' then case a when 'easily' then 100 when 'with_rampup' then 60 when 'no_one_else' then 20 else null end
    when 'q6' then case a when 'over_1yr' then 100 when '3_12mo' then 65 when 'under_3mo' then 35 else null end
    when 'q7' then case a when 'not_that_ive_seen' then 100 when 'a_little' then 55 when 'clearly_more_than_once' then 15 else null end
    when 'q8' then case a when 'no' then 100 when 'some' then 60 when 'significant' then 25 else null end
    when 'q9' then case a when 'a_lot_more' then 100 when 'a_little_more' then 70 when 'nothing_more' then 40 when 'less_not_more' then 10 else null end
    when 'q10' then case a when 'plenty_of_room' then 100 when 'comfortably_full' then 70 when 'stretched' then 35 when 'over_capacity' then 5 else null end
    when 'p1' then case a when 'light' then 100 when 'comfortably_full' then 70 when 'stretched' then 35 when 'overwhelmed' then 5 else null end
    when 'p2' then case a when 'yes_easily' then 100 when 'yes_but_something_slips' then 55 when 'no_not_right_now' then 10 else null end
    when 'p3' then case a when 'no' then 100 when 'yes' then 50 else null end
    else null
  end
$$ language sql immutable set search_path = public;

-- Recomputes and upserts one staff member's margin_scores row from their latest survey
-- (required) blended with their latest answered pulse from the last 14 days (if any), then
-- fires an OD notification when the result lands on over_capacity. Runs as security definer so
-- it can write margin_scores/notifications regardless of which caller (OD, or the team member
-- answering a pulse) triggered the recompute — those two callers have very different RLS
-- privileges, and this function is the one place that's allowed to bridge them.
create or replace function recompute_margin_score(p_staff_id bigint) returns void as $$
declare
  v_staff record;
  v_survey record;
  v_pulse record;
  v_base numeric;
  v_pulse_score numeric;
  v_final numeric;
  v_status text;
  v_calibration_gap boolean;
  v_od record;
begin
  select * into v_staff from staff where id = p_staff_id;
  if not found then return; end if;

  select * into v_survey from margin_surveys where staff_id = p_staff_id order by created_at desc limit 1;
  if v_survey is null then return; end if; -- nothing to score until the first survey lands

  v_base := (
    margin_answer_points('q1', v_survey.answers->>'q1') +
    margin_answer_points('q2', v_survey.answers->>'q2') +
    margin_answer_points('q3', v_survey.answers->>'q3') +
    margin_answer_points('q4', v_survey.answers->>'q4') +
    margin_answer_points('q5', v_survey.answers->>'q5') +
    margin_answer_points('q6', v_survey.answers->>'q6') +
    margin_answer_points('q7', v_survey.answers->>'q7') +
    margin_answer_points('q8', v_survey.answers->>'q8')
  ) / 8.0;

  -- Q9/Q10 are the OD's own calibration read — compared against the computed base, never
  -- blended into it, so a sharp disagreement is a signal in its own right rather than
  -- something silently averaged away.
  v_calibration_gap := abs(v_base - margin_answer_points('q10', v_survey.answers->>'q10')) > 30;

  select * into v_pulse from margin_pulses
    where staff_id = p_staff_id and status = 'answered' and responded_at > now() - interval '14 days'
    order by responded_at desc limit 1;

  if v_pulse is not null then
    v_pulse_score := (
      margin_answer_points('p1', v_pulse.answers->>'p1') +
      margin_answer_points('p2', v_pulse.answers->>'p2') +
      coalesce(margin_answer_points('p3', v_pulse.answers->>'p3'), 100)
    ) / 3.0;
    v_final := (v_base * 0.6) + (v_pulse_score * 0.4);
  else
    v_final := v_base;
  end if;

  v_status := case
    when v_final >= 70 then 'comfortable'
    when v_final >= 45 then 'full'
    when v_final >= 25 then 'stretched'
    else 'over_capacity'
  end;

  insert into margin_scores (staff_id, organization_id, campus_id, score, status, calibration_gap, last_survey_at, last_pulse_at, updated_at)
  values (p_staff_id, v_staff.organization_id, v_staff.campus_id, v_final, v_status, v_calibration_gap, v_survey.created_at, v_pulse.responded_at, now())
  on conflict (staff_id) do update set
    score = excluded.score,
    status = excluded.status,
    calibration_gap = excluded.calibration_gap,
    last_survey_at = excluded.last_survey_at,
    last_pulse_at = coalesce(excluded.last_pulse_at, margin_scores.last_pulse_at),
    updated_at = now();

  if v_status = 'over_capacity' then
    for v_od in select * from profiles where tier = 'od' and campus_id = v_staff.campus_id and organization_id = v_staff.organization_id loop
      insert into notifications (organization_id, for_user_name, actor, summary, project_id, ts, read)
      values (
        v_staff.organization_id,
        trim(coalesce(v_od.first_name, '') || ' ' || coalesce(v_od.last_name, '')),
        'OpsCore',
        v_staff.name || ' is showing as over capacity. Check in: "How are you doing with everything on your plate right now — really?" · "Is there anything on your list we could delay, hand off, or take off entirely?" · "What would actually make this week feel more manageable?"',
        null,
        to_char(now(), 'YYYY-MM-DD'),
        false
      );
    end loop;
  end if;
end;
$$ language plpgsql security definer set search_path = public;

create or replace function trg_margin_survey_recompute() returns trigger as $$
begin
  perform recompute_margin_score(new.staff_id);
  return new;
end;
$$ language plpgsql security definer set search_path = public;

create trigger margin_survey_after_insert
  after insert on margin_surveys
  for each row execute function trg_margin_survey_recompute();

create or replace function trg_margin_pulse_recompute() returns trigger as $$
begin
  if new.status = 'answered' and (old.status is distinct from new.status) then
    perform recompute_margin_score(new.staff_id);
  end if;
  return new;
end;
$$ language plpgsql security definer set search_path = public;

create trigger margin_pulse_after_update
  after update on margin_pulses
  for each row execute function trg_margin_pulse_recompute();

-- ---------- RLS ----------

alter table margin_scores enable row level security;
alter table margin_surveys enable row level security;
alter table margin_pulses enable row level security;

-- margin_scores is only ever written by recompute_margin_score() above, which runs as security
-- definer and so bypasses RLS/grants for its own writes — the only client-facing policy needed
-- here is read access for the people allowed to see the number at all.
create policy "central od read margin_scores" on margin_scores for select
  using (organization_id = my_org() and (my_tier() = 'central' or (my_tier() = 'od' and campus_id = my_campus())));

create policy "central od insert margin_surveys" on margin_surveys for insert
  with check (organization_id = my_org() and (my_tier() = 'central' or (my_tier() = 'od' and campus_id = my_campus())));
create policy "central od read margin_surveys" on margin_surveys for select
  using (organization_id = my_org() and (my_tier() = 'central' or (my_tier() = 'od' and campus_id = my_campus())));

create policy "central od insert margin_pulses" on margin_pulses for insert
  with check (organization_id = my_org() and (my_tier() = 'central' or (my_tier() = 'od' and campus_id = my_campus())));
create policy "central od read margin_pulses" on margin_pulses for select
  using (organization_id = my_org() and (my_tier() = 'central' or (my_tier() = 'od' and campus_id = my_campus())));

-- The narrow staff-tier exception: a team member can see and answer only their own pulse, and
-- only while it's still pending — the moment status flips to 'answered' this policy stops
-- matching, so they lose visibility to the row (and the score column they were never granted
-- update rights to in the first place).
create policy "staff sees own pending pulse" on margin_pulses for select
  using (status = 'pending' and staff_id in (select id from staff where user_id = auth.uid()));
create policy "staff responds to own pending pulse" on margin_pulses for update
  using (status = 'pending' and staff_id in (select id from staff where user_id = auth.uid()));

-- ---------- grants ----------
-- (CLI-pushed tables don't get Supabase's default dashboard grants — base GRANT is checked
-- before RLS ever runs, the same gotcha hit and fixed in 0004_grants.sql.)

grant select on margin_scores to authenticated;
grant select, insert on margin_surveys to authenticated;
grant select, insert on margin_pulses to authenticated;
revoke update on margin_pulses from authenticated;
grant update (answers, status, responded_at) on margin_pulses to authenticated;

grant all on margin_scores, margin_surveys, margin_pulses to service_role;
