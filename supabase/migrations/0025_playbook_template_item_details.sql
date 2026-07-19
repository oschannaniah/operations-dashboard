-- Extends checklist items with three more template-level fields, each copied into a run item
-- as a starting default when a run is applied (same snapshot principle as text/position):
--
-- category   — freeform label Central defines per item (e.g. "HR", "IT", "Week 1") so a long
--              checklist can be grouped instead of read as one flat list.
-- managed_by — a role/title suggestion (e.g. "Campus Operations Director"), since a template is
--              org-wide and can't name a specific person the way a run (tied to one real
--              campus) can. Copied as the run item's starting managed_by text, still fully
--              editable there against that campus's actual roster.
-- due_offset_days — "due within N days of the run starting," not an absolute date, because a
--              template gets reused indefinitely — a fixed calendar date would go stale after
--              its first use. Converted into a real due_date on the run item at apply-time.
alter table playbook_template_items add column category text;
alter table playbook_template_items add column managed_by text;
alter table playbook_template_items add column due_offset_days integer;

alter table playbook_run_items add column category text;
alter table playbook_run_items add column due_date date;
