// OpsCore — reads a campus's org chart straight out of its Google Slides deck (the slide
// matching that campus's current phase), following the deck's actual connector lines to
// resolve reporting structure. Ports importOrgChartFromSlides_ from the old Code.gs — the one
// real change is that Apps Script's Slides.Presentations.get() advanced service becomes a
// plain authenticated REST call, since Deno doesn't have that abstraction, just fetch().
// Uses the SAME stored Google connection as the Calendar integration (whoever connects Google
// Calendar also grants Slides read access in the same consent) — no second OAuth flow needed.

import { createClient } from "jsr:@supabase/supabase-js@2";

const GOOGLE_CLIENT_ID = Deno.env.get("GOOGLE_CLIENT_ID")!;
const GOOGLE_CLIENT_SECRET = Deno.env.get("GOOGLE_CLIENT_SECRET")!;

const supabaseAdmin = createClient(
  Deno.env.get("SUPABASE_URL")!,
  Deno.env.get("SUPABASE_SERVICE_ROLE_KEY")!,
);

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

function extractPresentationId(url: string): string {
  const m = String(url).match(/\/presentation\/d\/([a-zA-Z0-9-_]+)/);
  if (!m) throw new Error("That doesn't look like a Google Slides link.");
  return m[1];
}

function shapeText(shape: any): string {
  if (!shape?.text?.textElements) return "";
  return shape.text.textElements.map((t: any) => t.textRun?.content || "").join("");
}

Deno.serve(async (req) => {
  try {
    const authHeader = req.headers.get("Authorization") ?? "";
    const jwt = authHeader.replace("Bearer ", "");
    const { data: userData, error: userErr } = await supabaseAdmin.auth.getUser(jwt);
    if (userErr || !userData.user) return jsonOut({ error: "Unauthorized" }, 401);
    const userId = userData.user.id;

    const { data: profile } = await supabaseAdmin.from("profiles").select("tier, campus_id").eq("id", userId).single();
    if (!profile || (profile.tier !== "central" && profile.tier !== "od")) return jsonOut({ error: "Forbidden." }, 403);

    const { data } = await req.json();
    const { url, phase: rawPhase, campusId } = data || {};
    if (profile.tier === "od" && campusId !== profile.campus_id) return jsonOut({ error: "Forbidden: outside your campus." }, 403);
    if (!url) return jsonOut({ error: "No Slides link saved for this campus yet." });

    const { data: conn } = await supabaseAdmin.from("google_connections").select("refresh_token").eq("profile_id", userId).maybeSingle();
    if (!conn) return jsonOut({ error: "Connect your Google Calendar first (Calendar tab) — the Slides import reuses that same connection." });

    let presentationId: string;
    try {
      presentationId = extractPresentationId(url);
    } catch (e) {
      return jsonOut({ error: e instanceof Error ? e.message : String(e) });
    }

    const accessToken = await refreshAccessToken(conn.refresh_token);
    const presRes = await fetch(`https://slides.googleapis.com/v1/presentations/${presentationId}`, {
      headers: { Authorization: `Bearer ${accessToken}` },
    });
    const presentation = await presRes.json();
    if (presentation.error) {
      return jsonOut({ error: `Couldn't open that Slides deck: ${presentation.error.message}. Check the link is correct and shared with the Google account you connected, and that it was granted Slides access.` });
    }

    const phase = Number(rawPhase) || 1;
    const slides = presentation.slides || [];
    const slide = slides[phase - 1];
    if (!slide) {
      return jsonOut({ error: `This deck only has ${slides.length} slide(s) — expected at least ${phase} (one per phase, phase ${phase} is slide #${phase}).` });
    }

    // First pass: every text box becomes a person — first line of text is their name, second
    // line (if any) is their role. Vertical position is kept only to resolve reporting
    // direction in the second pass, never shown or stored.
    const people: Record<string, { name: string; role: string; y: number; reportsTo: string | null }> = {};
    for (const el of slide.pageElements || []) {
      if (!el.shape) continue;
      const lines = shapeText(el.shape).split("\n").map((s: string) => s.trim()).filter(Boolean);
      if (lines.length === 0) continue;
      const y = el.transform?.translateY || 0;
      people[el.objectId] = { name: lines[0], role: lines[1] || "", y, reportsTo: null };
    }

    // Second pass: every connector line links two boxes — whichever box sits higher on the
    // slide (smaller Y) is who the other one reports to.
    for (const el of slide.pageElements || []) {
      const conn2 = el.line?.lineProperties;
      const startId = conn2?.startConnection?.connectedObjectId;
      const endId = conn2?.endConnection?.connectedObjectId;
      if (!startId || !endId || !people[startId] || !people[endId]) continue;
      const [higherId, lowerId] = people[startId].y <= people[endId].y ? [startId, endId] : [endId, startId];
      people[lowerId].reportsTo = people[higherId].name;
    }

    const result = Object.values(people)
      .filter((p) => p.name)
      .map((p) => ({ name: p.name, role: p.role, reportsTo: p.reportsTo }));

    if (result.length === 0) return jsonOut({ error: "Didn't find any named boxes on that slide." });
    return jsonOut({ people: result });
  } catch (err) {
    return jsonOut({ error: err instanceof Error ? err.message : String(err) });
  }
});

function jsonOut(body: unknown, status = 200) {
  return new Response(JSON.stringify(body), { status, headers: { "Content-Type": "application/json" } });
}
