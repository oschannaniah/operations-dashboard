-- Fixes a real bug found during verification: the original "central manages all requests" FOR
-- ALL policy let a Central user self-approve their own request. Postgres OR's every permissive
-- policy that matches a command together — that broad policy had no self-exclusion, so it
-- silently overrode "od central decide requests"' correct restriction whenever the decider was
-- Central, since either policy passing is enough to allow the UPDATE.
--
-- Splits Central's access into one policy per command instead. The decide policy (unchanged)
-- is now the *only* UPDATE policy on this table, so its self-exclusion can never be bypassed by
-- a broader policy elsewhere.

drop policy "central manages all requests" on approval_requests;

create policy "central reads all requests" on approval_requests for select
  using (organization_id = my_org() and my_tier() = 'central');

create policy "central creates own requests" on approval_requests for insert
  with check (organization_id = my_org() and my_tier() = 'central' and requested_by_profile_id = auth.uid());

create policy "central deletes any request" on approval_requests for delete
  using (organization_id = my_org() and my_tier() = 'central');
