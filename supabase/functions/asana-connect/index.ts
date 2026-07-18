// OpsCore — permanent per-user Asana connection. Mirrors google-calendar/index.ts's shape
// exactly (same OAuth-code-exchange pattern, same service-role trust boundary), but the whole
// point of this file existing yet is scaffolding: ASANA_CLIENT_ID/ASANA_CLIENT_SECRET/
// ASANA_REDIRECT_URI are not set as Supabase secrets yet, so every action below returns a
// clean "not configured" error until they are. Nothing here reads Asana task data yet either —
// that's the fast-follow once a live connection actually exists to read from.

import { createClient } from "jsr:@supabase/supabase-js@2";

const ASANA_CLIENT_ID = Deno.env.get("ASANA_CLIENT_ID");
const ASANA_CLIENT_SECRET = Deno.env.get("ASANA_CLIENT_SECRET");
const REDIRECT_URI = Deno.env.get("ASANA_REDIRECT_URI") ?? "";

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
      const { data: conn } = await supabaseAdmin.from("asana_connections").select("profile_id").eq("profile_id", userId).maybeSingle();
      return jsonOut({ connected: !!conn, configured: !!(ASANA_CLIENT_ID && ASANA_CLIENT_SECRET) });
    }

    if (!ASANA_CLIENT_ID || !ASANA_CLIENT_SECRET) {
      return jsonOut({ error: "Asana isn't set up yet — ASANA_CLIENT_ID and ASANA_CLIENT_SECRET need to be added as Supabase secrets first (from a real Asana developer app)." });
    }

    if (action === "authorizeUrl") {
      const url = new URL("https://app.asana.com/-/oauth_authorize");
      url.searchParams.set("client_id", ASANA_CLIENT_ID);
      url.searchParams.set("redirect_uri", REDIRECT_URI);
      url.searchParams.set("response_type", "code");
      url.searchParams.set("state", userId);
      return jsonOut({ url: url.toString() });
    }

    if (action === "connect") {
      const res = await fetch("https://app.asana.com/-/oauth_token", {
        method: "POST",
        headers: { "Content-Type": "application/x-www-form-urlencoded" },
        body: new URLSearchParams({
          grant_type: "authorization_code",
          client_id: ASANA_CLIENT_ID,
          client_secret: ASANA_CLIENT_SECRET,
          redirect_uri: REDIRECT_URI,
          code: data.code,
        }),
      });
      const tokens = await res.json();
      if (tokens.error) return jsonOut({ error: tokens.error_description || tokens.error });
      await supabaseAdmin.from("asana_connections").upsert({
        profile_id: userId,
        access_token: tokens.access_token,
        refresh_token: tokens.refresh_token ?? null,
        workspace_gid: tokens.data?.gid ?? null,
        updated_at: new Date().toISOString(),
      });
      return jsonOut({ ok: true });
    }

    if (action === "disconnect") {
      await supabaseAdmin.from("asana_connections").delete().eq("profile_id", userId);
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
