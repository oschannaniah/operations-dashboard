-- Splits per-step ownership on a playbook run into two roles. assigned_to (0023) already
-- covers who's scheduled to complete a step; this adds who's managing/overseeing it — often a
-- different person. On an onboarding run, for example, the OD might manage "set up email"
-- (make sure IT follows through) while the new hire's mentor is the one actually scheduled to
-- do the walkthrough steps. done_by (who actually checked it off) stays separate from both —
-- that's history, not an assignment.
alter table playbook_run_items add column managed_by text;
