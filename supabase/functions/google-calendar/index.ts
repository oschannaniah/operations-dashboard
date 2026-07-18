// OpsCore — permanent per-user Google Calendar connection.
// Ports handleGoogleCalendarAction_ from the old Code.gs almost line for line: Apps Script's
// UrlFetchApp was never doing anything Deno's built-in fetch() can't do directly. Runs with
// the service role (same trust level Apps Script had) — the browser never sees a refresh
// token or the Client Secret, only ever calls this function with its own user JWT.

import { createClient } from "jsr:@supabase/supabase-js@2";

const GOOGLE_CLIENT_ID = Deno.env.get("GOOGLE_CLIENT_ID")!;
const GOOGLE_CLIENT_SECRET = Deno.env.get("GOOGLE_CLIENT_SECRET")!;

const supabaseAdmin = createClient(
  Deno.env.get("SUPABASE_URL")!,
  Deno.env.get("SUPABASE_SERVICE_ROLE_KEY")!,
);

async function exchangeAuthCode(code: string) {
  const res = await fetch("https://oauth2.googleapis.com/token", {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body: new URLSearchParams({
      code,
      client_id: GOOGLE_CLIENT_ID,
      client_secret: GOOGLE_CLIENT_SECRET,
      redirect_uri: "postmessage", // matches the popup code-client flow the frontend uses
      grant_type: "authorization_code",
    }),
  });
  const data = await res.json();
  if (data.error) throw new Error(data.error_description || data.error);
  return data as { access_token: string; refresh_token?: string; scope?: string };
}

async function refreshAccessToken(refreshToken: string) {
  const res = await fetch("https://oauth2.googleapis.com/token", {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body: new URLSearchParams({
      refresh_token: refreshToken,
      client_id: GOOGLE_CLIENT_ID,
      client_secret: GOOGLE_CLIENT_SECRET,
      grant_type: "refresh_token",
    }),
  });
  const data = await res.json();
  if (data.error) throw new Error(data.error_description || data.error);
  return data.access_token as string;
}

Deno.serve(async (req) => {
  try {
    const authHeader = req.headers.get("Authorization") ?? "";
    const jwt = authHeader.replace("Bearer ", "");
    const { data: userData, error: userErr } = await supabaseAdmin.auth.getUser(jwt);
    if (userErr || !userData.user) return jsonOut({ error: "Unauthorized" }, 401);
    const userId = userData.user.id;

    const { action, data } = await req.json();

    if (action === "connect") {
      const tokens = await exchangeAuthCode(data.code);
      if (!tokens.refresh_token) {
        return jsonOut({
          error: "Google didn't return a lasting connection — this can happen on a second attempt. Revoke access at myaccount.google.com/permissions for this app, then try connecting again.",
        });
      }
      await supabaseAdmin.from("google_connections").upsert({
        profile_id: userId, refresh_token: tokens.refresh_token, granted_scopes: tokens.scope ?? null, updated_at: new Date().toISOString(),
      });
      await supabaseAdmin.from("profiles").update({}).eq("id", userId); // no-op touch; connected flag lives in google_connections existence
      return jsonOut({ ok: true });
    }

    const { data: conn } = await supabaseAdmin.from("google_connections").select("refresh_token").eq("profile_id", userId).maybeSingle();

    if (action === "listCalendars") {
      if (!conn) return jsonOut({ error: "Not connected." });
      const accessToken = await refreshAccessToken(conn.refresh_token);
      const res = await fetch("https://www.googleapis.com/calendar/v3/users/me/calendarList?maxResults=250", {
        headers: { Authorization: `Bearer ${accessToken}` },
      });
      const result = await res.json();
      if (result.error) return jsonOut({ error: result.error.message || "Google Calendar API error" });
      return jsonOut({ calendars: (result.items || []).map((c: any) => ({ id: c.id, name: c.summaryOverride || c.summary, primary: !!c.primary })) });
    }

    if (action === "listEvents") {
      if (!conn) return jsonOut({ error: "Not connected." });
      const accessToken = await refreshAccessToken(conn.refresh_token);
      const timeMin = new Date(); timeMin.setDate(timeMin.getDate() - 7);
      const timeMax = new Date(); timeMax.setDate(timeMax.getDate() + 60);
      const allEvents: any[] = [];
      for (const id of (data.calendarIds || [])) {
        const params = new URLSearchParams({ timeMin: timeMin.toISOString(), timeMax: timeMax.toISOString(), singleEvents: "true", orderBy: "startTime", maxResults: "250" });
        const res = await fetch(`https://www.googleapis.com/calendar/v3/calendars/${encodeURIComponent(id)}/events?${params}`, {
          headers: { Authorization: `Bearer ${accessToken}` },
        });
        const result = await res.json();
        if (!result.error) allEvents.push(...(result.items || []));
      }
      return jsonOut({ events: allEvents });
    }

    if (action === "disconnect") {
      if (conn) {
        await fetch(`https://oauth2.googleapis.com/revoke?token=${encodeURIComponent(conn.refresh_token)}`, { method: "POST" });
      }
      await supabaseAdmin.from("google_connections").delete().eq("profile_id", userId);
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
