// OpsCore — admin-level account management (Team Accounts panel + Staff & Team "Link Login").
//
// Why this exists (a gap found while implementing, not in the original plan): creating a login
// FOR someone else with a chosen password requires Supabase's admin API
// (auth.admin.createUser), which needs the service role key — never safe to expose to the
// browser. Likewise, assigning someone's campus/role (which recomputes their tier) has to
// bypass the column-grant restriction on profiles that stops a regular user self-elevating
// their own tier — see the comment in migrations/0002_rls.sql. Both operations are Central-only
// (mirrors the old Code.gs: "Only Central can create user accounts").

import { createClient } from "jsr:@supabase/supabase-js@2";

const supabaseAdmin = createClient(
  Deno.env.get("SUPABASE_URL")!,
  Deno.env.get("SUPABASE_SERVICE_ROLE_KEY")!,
);

// Mirrors tierForRole_ in the old Code.gs.
function tierForRole(campusId: string | null, role: string | null): string {
  if (campusId === "central") return "central";
  if (!campusId) return "unassigned";
  if (role === "Campus Operations Director") return "od";
  return "staff";
}

Deno.serve(async (req) => {
  try {
    const authHeader = req.headers.get("Authorization") ?? "";
    const jwt = authHeader.replace("Bearer ", "");
    const { data: userData, error: userErr } = await supabaseAdmin.auth.getUser(jwt);
    if (userErr || !userData.user) return jsonOut({ error: "Unauthorized" }, 401);
    const callerId = userData.user.id;

    const { data: callerProfile } = await supabaseAdmin.from("profiles").select("tier, organization_id").eq("id", callerId).single();
    if (!callerProfile || callerProfile.tier !== "central") return jsonOut({ error: "Only Central can manage accounts." }, 403);

    const { action, data } = await req.json();

    if (action === "create") {
      const tier = tierForRole(data.campusId, data.role);
      const { data: created, error: createErr } = await supabaseAdmin.auth.admin.createUser({
        email: data.email,
        password: data.password,
        email_confirm: true,
        user_metadata: { organization_id: callerProfile.organization_id },
      });
      if (createErr) return jsonOut({ error: createErr.message });

      // The auth.users insert trigger already created a bare profiles row (tier
      // 'unassigned') — fill in the real details now.
      const { error: updateErr } = await supabaseAdmin.from("profiles").update({
        first_name: data.firstName, last_name: data.lastName, phone: data.phone || "",
        campus_id: data.campusId, role: data.role, tier,
      }).eq("id", created.user.id);
      if (updateErr) return jsonOut({ error: updateErr.message });

      return jsonOut({ ok: true, id: created.user.id });
    }

    if (action === "updateAccess") {
      const tier = tierForRole(data.campusId, data.role);
      const { error } = await supabaseAdmin.from("profiles")
        .update({ campus_id: data.campusId, role: data.role, tier })
        .eq("id", data.userId).eq("organization_id", callerProfile.organization_id);
      if (error) return jsonOut({ error: error.message });
      return jsonOut({ ok: true });
    }

    if (action === "delete") {
      const { error } = await supabaseAdmin.auth.admin.deleteUser(data.userId);
      if (error) return jsonOut({ error: error.message });
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
