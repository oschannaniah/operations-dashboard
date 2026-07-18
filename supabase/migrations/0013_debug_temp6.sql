create or replace function debug_margin_v3(p_staff_id bigint) returns jsonb as $$
declare
  v_pulse record;
  v_answers jsonb;
  v_p1_text text;
  v_p1 numeric;
begin
  select * into v_pulse from margin_pulses
    where staff_id = p_staff_id and status = 'answered' and responded_at > now() - interval '14 days'
    order by responded_at desc limit 1;

  v_answers := v_pulse.answers;
  v_p1_text := v_answers->>'p1';
  v_p1 := margin_answer_points('p1', v_p1_text);

  return jsonb_build_object(
    'v_answers', v_answers,
    'v_answers_is_null', v_answers is null,
    'v_p1_text', v_p1_text,
    'v_p1_text_is_null', v_p1_text is null,
    'v_p1', v_p1,
    'direct_v_pulse_answers', v_pulse.answers,
    'direct_extraction', v_pulse.answers->>'p1',
    'direct_lookup', margin_answer_points('p1', v_pulse.answers->>'p1'),
    'pg_typeof_answers', pg_typeof(v_pulse.answers)::text
  );
end;
$$ language plpgsql security definer set search_path = public;

grant execute on function debug_margin_v3(bigint) to service_role;
