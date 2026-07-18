// OpsCore — permanent per-user Slack connection. Same shape and same scaffolding-only status
// as asana-connect/index.ts — SLACK_CLIENT_ID/SLACK_CLIENT_SECRET/SLACK_REDIRECT_URI aren't
// set as Supabase secrets yet, so every action returns a clean "not configured" error until
// they are. Nothing here posts messages yet either — that's the fast-follow once a live
// connection exists.

import { createClient } from "jsr:@supabase/supabase-js@2";

const SLACK_CLIENT_ID = Deno.env.get("SLACK_CLIENT_ID");
const SLACK_CLIENT_SECRET = Deno.env.get("SLACK_CLIENT_SECRET");
const REDIRECT_URI = Deno.env.get("SLACK_REDIRECT_URI") ?? "";
// Minimal scope for now: identify the user and allow posting on their behalf later. Widen
// deliberately, not by default, once there's a real feature that needs more.
const SLACK_SCOPES = "chat:write,users:read";

const supabaseAdmin = createClient(
  Deno.env.get("SUPABASE_URL")!,
  Deno.env.get("SUPABASE_SERVICE_ROLE_KEY")!,
);

Deno.serve(async (req) => {
  try {
    const authHeader = req.headers.get("Authorization") ?? "";
    const jwt = authHeader.replace("Bearer ", "");
    const { data: userData, error: userErr } = await supabaseAdmin.auth.getUser(jwt);
    if (userErr || !userData.user) return jsonOut({ error: "Unauthorized" }, 401);
    const userId = userData.user.id;

    const { action, data } = await req.json();

    if (action === "status") {
      const { data: conn } = await supabaseAdmin.from("slack_connections").select("profile_id, team_name").eq("profile_id", userId).maybeSingle();
      return jsonOut({ connected: !!conn, teamName: conn?.team_name ?? null, configured: !!(SLACK_CLIENT_ID && SLACK_CLIENT_SECRET) });
    }

    if (!SLACK_CLIENT_ID || !SLACK_CLIENT_SECRET) {
      return jsonOut({ error: "Slack isn't set up yet — SLACK_CLIENT_ID and SLACK_CLIENT_SECRET need to be added as Supabase secrets first (from a real Slack app)." });
    }

    if (action === "authorizeUrl") {
      const url = new URL("https://slack.com/oauth/v2/authorize");
      url.searchParams.set("client_id", SLACK_CLIENT_ID);
      url.searchParams.set("scope", SLACK_SCOPES);
      url.searchParams.set("redirect_uri", REDIRECT_URI);
      url.searchParams.set("state", userId);
      return jsonOut({ url: url.toString() });
    }

    if (action === "connect") {
      const res = await fetch("https://slack.com/api/oauth.v2.access", {
        method: "POST",
        headers: { "Content-Type": "application/x-www-form-urlencoded" },
        body: new URLSearchParams({
          client_id: SLACK_CLIENT_ID,
          client_secret: SLACK_CLIENT_SECRET,
          redirect_uri: REDIRECT_URI,
          code: data.code,
        }),
      });
      const result = await res.json();
      if (!result.ok) return jsonOut({ error: result.error || "Slack authorization failed" });
      await supabaseAdmin.from("slack_connections").upsert({
        profile_id: userId,
        access_token: result.authed_user?.access_token || result.access_token,
        team_id: result.team?.id ?? null,
        team_name: result.team?.name ?? null,
        updated_at: new Date().toISOString(),
      });
      return jsonOut({ ok: true });
    }

    if (action === "disconnect") {
      await supabaseAdmin.from("slack_connections").delete().eq("profile_id", userId);
      return jsonOut({ ok: true });
    }

    return jsonOut({ error: `Unknown action: ${action}` });
  } catch (err) {
    return jsonOut({ error: err instanceof Error ? err.message : String(err) });
  }
});

function jsonOut(body: unknown, status = 200) {
  return new Response(JSON.stringify(body), { status, headers: { "Content-Type": "application/json" } });
}
