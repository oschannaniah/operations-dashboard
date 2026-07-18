create or replace function debug_margin_v4(p_staff_id bigint) returns jsonb as $$
declare
  v_pulse record;
  v_p1 numeric;
  v_p2 numeric;
  v_p3 numeric;
  v_sum numeric;
begin
  select * into v_pulse from margin_pulses
    where staff_id = p_staff_id and status = 'answered' and responded_at > now() - interval '14 days'
    order by responded_at desc limit 1;

  v_p1 := margin_answer_points('p1', v_pulse.answers->>'p1');
  v_p2 := margin_answer_points('p2', v_pulse.answers->>'p2');
  v_p3 := margin_answer_points('p3', v_pulse.answers->>'p3');
  v_sum := v_p1 + v_p2 + v_p3;

  return jsonb_build_object(
    'v_p1', v_p1, 'v_p2', v_p2, 'v_p3', v_p3, 'v_sum', v_sum,
    'p2_text', v_pulse.answers->>'p2', 'p3_text', v_pulse.answers->>'p3'
  );
end;
$$ language plpgsql security definer set search_path = public;

grant execute on function debug_margin_v4(bigint) to service_role;
