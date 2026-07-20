// OpsCore — scheduled nudge for Mobile Check-ins. Finds anyone with a linked login who hasn't
// submitted a self-initiated check-in (mobile_checkins, 0031) within the trailing window and
// sends them a Web Push notification if they've opted in.
//
// Not user-invoked — this runs on a schedule (pg_cron -> net.http_post, see the migration this
// ships alongside) carrying a dedicated shared secret (CHECKIN_NUDGE_SECRET) as its bearer
// token, not the project's service role key. Supabase currently issues two different formats
// of that key (legacy JWT vs. newer sb_secret_...) and which one an Edge Function sees
// auto-injected doesn't necessarily match what's stored anywhere else — a purpose-built secret
// sidesteps that ambiguity instead of trying to keep two representations of "the same" key in
// sync. auth.getUser() isn't used here either, since it only resolves real user sessions.
//
// Email is deliberately NOT wired in yet — the only email capability this project has today is
// Supabase's fixed invite-template sender (see admin-accounts/index.ts), not a general
// transactional sender. sendEmailNudge() below is a real, callable stub: the moment a
// RESEND_API_KEY (or similar) secret exists, fill in the fetch call and remove the early
// return — everything else (who's eligible, how often) is already correct and doesn't change.

import { createClient } from "jsr:@supabase/supabase-js@2";
import webpush from "npm:web-push@3";

const supabaseAdmin = createClient(
  Deno.env.get("SUPABASE_URL")!,
  Deno.env.get("SUPABASE_SERVICE_ROLE_KEY")!,
);

webpush.setVapidDetails(
  Deno.env.get("VAPID_SUBJECT")!,
  Deno.env.get("VAPID_PUBLIC_KEY")!,
  Deno.env.get("VAPID_PRIVATE_KEY")!,
);

// How long someone can go without a self-initiated check-in before this nudges them. A plain
// constant, not a Central-configurable setting (unlike Capacity Weights) — this is a first cut
// and easy to tune here directly once there's real data on what cadence actually lands well.
const NUDGE_WINDOW_DAYS = 3;

async function sendEmailNudge(_email: string, _name: string): Promise<void> {
  // No transactional email provider configured yet — see the file header. This is a no-op
  // until a real provider key exists as a secret.
  return;
}

Deno.serve(async (req) => {
  const authHeader = req.headers.get("Authorization") ?? "";
  const bearer = authHeader.replace("Bearer ", "");
  if (bearer !== Deno.env.get("CHECKIN_NUDGE_SECRET")) {
    return jsonOut({ error: "Unauthorized" }, 401);
  }

  try {
    const cutoff = new Date(Date.now() - NUDGE_WINDOW_DAYS * 24 * 60 * 60 * 1000).toISOString();

    const [{ data: staffRows, error: staffErr }, { data: recentCheckins, error: checkinErr }] = await Promise.all([
      supabaseAdmin.from("staff").select("id, name, email, user_id").not("user_id", "is", null),
      supabaseAdmin.from("mobile_checkins").select("staff_id, created_at").gte("created_at", cutoff),
    ]);
    if (staffErr) return jsonOut({ error: staffErr.message }, 500);
    if (checkinErr) return jsonOut({ error: checkinErr.message }, 500);

    const recentlyCheckedIn = new Set((recentCheckins ?? []).map((c) => c.staff_id));
    const eligible = (staffRows ?? []).filter((s) => !recentlyCheckedIn.has(s.id));

    let notified = 0, pruned = 0, emailed = 0;
    for (const person of eligible) {
      const { data: subs } = await supabaseAdmin.from("push_subscriptions").select("*").eq("profile_id", person.user_id);
      for (const sub of subs ?? []) {
        try {
          await webpush.sendNotification(
            { endpoint: sub.endpoint, keys: { p256dh: sub.p256dh, auth: sub.auth_key } },
            JSON.stringify({ title: "Quick check-in?", body: "It's been a few days — takes about 15 seconds.", url: "/?checkin=1" }),
          );
          notified++;
        } catch (err) {
          // 404/410 means the browser has invalidated this subscription (uninstalled,
          // permissions revoked, etc.) — dead weight, not a retry candidate.
          const status = (err as { statusCode?: number })?.statusCode;
          if (status === 404 || status === 410) {
            await supabaseAdmin.from("push_subscriptions").delete().eq("id", sub.id);
            pruned++;
          }
        }
      }
      if (person.email) {
        await sendEmailNudge(person.email, person.name);
        emailed++;
      }
    }

    return jsonOut({ ok: true, eligible: eligible.length, notified, pruned, emailed });
  } catch (err) {
    return jsonOut({ error: err instanceof Error ? err.message : String(err) }, 500);
  }
});

function jsonOut(body: unknown, status = 200) {
  return new Response(JSON.stringify(body), { status, headers: { "Content-Type": "application/json" } });
}
