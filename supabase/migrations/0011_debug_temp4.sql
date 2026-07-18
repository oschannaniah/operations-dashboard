create or replace function debug_margin_recompute(p_staff_id bigint) returns jsonb as $$
declare
  v_pulse record;
  v_pulse_count int;
begin
  select count(*) into v_pulse_count from margin_pulses where staff_id = p_staff_id;
  select * into v_pulse from margin_pulses
    where staff_id = p_staff_id and status = 'answered' and responded_at > now() - interval '14 days'
    order by responded_at desc limit 1;
  return jsonb_build_object(
    'total_pulse_rows_for_staff', v_pulse_count,
    'v_pulse_is_null', v_pulse is null,
    'v_pulse_id', v_pulse.id,
    'v_pulse_status', v_pulse.status,
    'v_pulse_responded_at', v_pulse.responded_at,
    'now_minus_14d', now() - interval '14 days',
    'all_rows', (select jsonb_agg(row_to_json(t)) from (select id, status, responded_at, answers from margin_pulses where staff_id = p_staff_id) t)
  );
end;
$$ language plpgsql security definer set search_path = public;

grant execute on function debug_margin_recompute(bigint) to service_role;
