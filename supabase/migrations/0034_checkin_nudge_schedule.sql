-- Schedules the checkin-nudge Edge Function (supabase/functions/checkin-nudge) to run daily.
-- pg_cron can't call an Edge Function directly, so this uses pg_net (also Supabase-managed) to
-- fire an HTTP POST at it — the pattern Supabase documents for this exact case.
--
-- The service role key the job needs for its Authorization header is pulled from Supabase
-- Vault (encrypted at rest, readable only via direct SQL access to this project) rather than
-- ever appearing as a literal value in a migration file that gets committed to git. The secret
-- itself (name: checkin_nudge_service_key) was created out-of-band with vault.create_secret()
-- before this migration ran — this file only ever references it by name.

create extension if not exists pg_cron;
create extension if not exists pg_net;

select cron.schedule(
  'checkin-nudge-daily',
  '0 15 * * *', -- 15:00 UTC daily — mid-morning US time, well clear of any campus's overnight
  $$
  select net.http_post(
    url := 'https://pcuadpgamkoaytksbkcl.supabase.co/functions/v1/checkin-nudge',
    headers := jsonb_build_object(
      'Authorization', 'Bearer ' || (select decrypted_secret from vault.decrypted_secrets where name = 'checkin_nudge_service_key'),
      'Content-Type', 'application/json'
    ),
    body := '{}'::jsonb
  );
  $$
);
