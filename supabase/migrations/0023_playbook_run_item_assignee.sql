-- Per-step ownership on a playbook run — separate from the run's own target (e.g. an
-- onboarding run's target is the new hire, but "set up their email" and "give the building
-- tour" are usually owned by different people). Free-text, same "identity by string"
-- convention as projects.owner / staff.reports_to elsewhere — not every assignee has a login.
alter table playbook_run_items add column assigned_to text;
