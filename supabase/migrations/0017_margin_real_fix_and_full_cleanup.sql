-- The actual fix. Root cause isolated by bisecting working vs. broken variants: populating a
-- SECOND `record`-typed variable via `select * into` (the pulse lookup), after a first
-- `record`-typed variable (the survey lookup) had several of its jsonb fields extracted and
-- passed through margin_answer_points() calls, silently produced a record whose jsonb column
-- read back as NULL — despite `record is not null` still (correctly) reporting the row as
-- found. Every attempt to work around this by restructuring the *arithmetic* (0006, 0010)
-- failed for the same underlying reason. The actual fix is avoiding `record` variables for
-- these two lookups entirely: scalar-subquery assignment into plain `jsonb` variables instead
-- of `select * into a_record`. Verified via an isolated repro (migrations 0011-0016, all
-- dropped below) — this exact pattern, and only this pattern, produces the correct blended
-- score.

create or replace function recompute_margin_score(p_staff_id bigint) returns void as $$
declare
  v_staff record;
  v_survey_answers jsonb;
  v_survey_created_at timestamptz;
  v_pulse_answers jsonb;
  v_pulse_responded_at timestamptz;
  v_q1 numeric; v_q2 numeric; v_q3 numeric; v_q4 numeric;
  v_q5 numeric; v_q6 numeric; v_q7 numeric; v_q8 numeric;
  v_base numeric;
  v_p1 numeric; v_p2 numeric; v_p3 numeric;
  v_pulse_score numeric;
  v_final numeric;
  v_status text;
  v_calibration_gap boolean;
  v_od record;
begin
  select * into v_staff from staff where id = p_staff_id;
  if not found then return; end if;

  select answers, created_at into v_survey_answers, v_survey_created_at
    from margin_surveys where staff_id = p_staff_id order by created_at desc limit 1;
  if v_survey_answers is null then return; end if; -- nothing to score until the first survey lands

  v_q1 := margin_answer_points('q1', v_survey_answers->>'q1');
  v_q2 := margin_answer_points('q2', v_survey_answers->>'q2');
  v_q3 := margin_answer_points('q3', v_survey_answers->>'q3');
  v_q4 := margin_answer_points('q4', v_survey_answers->>'q4');
  v_q5 := margin_answer_points('q5', v_survey_answers->>'q5');
  v_q6 := margin_answer_points('q6', v_survey_answers->>'q6');
  v_q7 := margin_answer_points('q7', v_survey_answers->>'q7');
  v_q8 := margin_answer_points('q8', v_survey_answers->>'q8');
  v_base := (v_q1 + v_q2 + v_q3 + v_q4 + v_q5 + v_q6 + v_q7 + v_q8) / 8.0;

  -- Q9/Q10 are the OD's own calibration read — compared against the computed base, never
  -- blended into it, so a sharp disagreement is a signal in its own right rather than
  -- something silently averaged away.
  v_calibration_gap := abs(v_base - margin_answer_points('q10', v_survey_answers->>'q10')) > 30;

  select answers, responded_at into v_pulse_answers, v_pulse_responded_at
    from margin_pulses
    where staff_id = p_staff_id and status = 'answered' and responded_at > now() - interval '14 days'
    order by responded_at desc limit 1;

  if v_pulse_answers is not null then
    v_p1 := margin_answer_points('p1', v_pulse_answers->>'p1');
    v_p2 := margin_answer_points('p2', v_pulse_answers->>'p2');
    v_p3 := coalesce(margin_answer_points('p3', v_pulse_answers->>'p3'), 100);
    v_pulse_score := (v_p1 + v_p2 + v_p3) / 3.0;
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
  values (p_staff_id, v_staff.organization_id, v_staff.campus_id, v_final, v_status, v_calibration_gap, v_survey_created_at, v_pulse_responded_at, now())
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

drop function if exists debug_margin_recompute(bigint);
drop function if exists debug_margin_old_style(bigint);
drop function if exists debug_margin_new_style(bigint);
drop function if exists debug_margin_v3(bigint);
drop function if exists debug_margin_v4(bigint);
drop function if exists debug_margin_v5(bigint);
drop function if exists debug_margin_v6(bigint);
