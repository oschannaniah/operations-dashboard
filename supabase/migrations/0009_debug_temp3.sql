create or replace function debug_margin_recompute(p_staff_id bigint) returns jsonb as $$
declare
  v_survey record;
  v_pulse record;
  v_base numeric;
  v_p1 numeric;
  v_p2 numeric;
  v_p3 numeric;
  v_pulse_score numeric;
  v_final numeric;
begin
  select * into v_survey from margin_surveys where staff_id = p_staff_id order by created_at desc limit 1;
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

  select * into v_pulse from margin_pulses
    where staff_id = p_staff_id and status = 'answered' and responded_at > now() - interval '14 days'
    order by responded_at desc limit 1;

  v_p1 := margin_answer_points('p1', v_pulse.answers->>'p1');
  v_p2 := margin_answer_points('p2', v_pulse.answers->>'p2');
  v_p3 := coalesce(margin_answer_points('p3', v_pulse.answers->>'p3'), 100);
  v_pulse_score := (v_p1 + v_p2 + v_p3) / 3.0;
  v_final := (v_base * 0.6) + (v_pulse_score * 0.4);

  return jsonb_build_object(
    'v_base', v_base, 'v_p1', v_p1, 'v_p2', v_p2, 'v_p3', v_p3,
    'v_pulse_score', v_pulse_score, 'v_final', v_final,
    'v_survey_answers', v_survey.answers, 'v_pulse_answers', v_pulse.answers
  );
end;
$$ language plpgsql security definer set search_path = public;

grant execute on function debug_margin_recompute(bigint) to service_role;
