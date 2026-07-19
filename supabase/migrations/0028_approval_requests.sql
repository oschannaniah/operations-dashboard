-- Approval workflows v1 — Budget, Purchase, and PTO requests: submit, route to an approver,
-- approve/deny, done. Deliberately single-step (no multi-level chains) for v1.
--
-- Approver resolution happens client-side at submit time (who gets notified), but WHO IS
-- ALLOWED TO DECIDE is enforced here, not just in the UI: an od-tier approver can only decide
-- requests on their own campus, Central can decide anything org-wide, and — critically — nobody
-- can decide their own request, regardless of tier. That last rule is what forces a solo
-- campus's OD's own request to escalate to Central: there's no peer admin on that campus who
-- could approve it, and the requester themselves is explicitly excluded.

create table approval_requests (
  id uuid primary key default gen_random_uuid(),
  organization_id uuid not null references organizations(id) on delete cascade,
  campus_id text not null references campuses(id),
  type text not null, -- 'budget' | 'purchase' | 'pto'
  amount numeric, -- budget/purchase only
  starts_on date, -- pto only
  ends_on date, -- pto only
  reason text not null,
  requested_by text not null, -- display name, same identity-by-string convention as elsewhere
  requested_by_profile_id uuid not null references profiles(id) on delete cascade,
  status text not null default 'pending', -- 'pending' | 'approved' | 'denied'
  decided_by text,
  decided_at timestamptz,
  decision_note text,
  created_at timestamptz not null default now()
);
create index approval_requests_campus_idx on approval_requests (campus_id, created_at desc);

alter table approval_requests enable row level security;

create policy "central manages all requests" on approval_requests for all
  using (organization_id = my_org() and my_tier() = 'central');

-- od/staff share full read/insert visibility on their own campus's requests — same shape as
-- staff/teams elsewhere in this schema (a campus is a shared visibility boundary, not a
-- per-person silo).
create policy "campus reads own requests" on approval_requests for select
  using (organization_id = my_org() and my_tier() in ('od', 'staff') and campus_id = my_campus());
create policy "campus creates own requests" on approval_requests for insert
  with check (organization_id = my_org() and my_tier() in ('od', 'staff') and campus_id = my_campus() and requested_by_profile_id = auth.uid());

-- Withdrawing is just deleting your own still-pending request — no self-approval risk in a
-- delete, so this doesn't need the column-grant treatment below.
create policy "requester withdraws own pending request" on approval_requests for delete
  using (my_tier() in ('od', 'staff') and campus_id = my_campus() and requested_by_profile_id = auth.uid() and status = 'pending');

-- Deciding (approve/deny) is its own policy, deliberately excluding the requester from ever
-- matching it themselves.
create policy "od central decide requests" on approval_requests for update
  using (
    organization_id = my_org()
    and requested_by_profile_id != auth.uid()
    and (my_tier() = 'central' or (my_tier() = 'od' and campus_id = my_campus()))
  );

grant select, insert, delete on approval_requests to authenticated;
revoke update on approval_requests from authenticated;
grant update (status, decided_by, decided_at, decision_note) on approval_requests to authenticated;
grant all on approval_requests to service_role;
