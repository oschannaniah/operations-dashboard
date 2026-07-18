-- OpsCore — seed data: OSC's organization + its 8 campuses + the "central" pseudo-campus.
-- Values match the current CAMPUSES const in campus-ops-dashboard.jsx exactly.

insert into organizations (id, name, slug)
values ('00000000-0000-0000-0000-000000000001', 'OSC Family Church', 'osc')
on conflict (id) do nothing;

insert into campuses (id, organization_id, name, abbr, lead_name, phase, color) values
  ('abv', '00000000-0000-0000-0000-000000000001', 'Abbeville',    'AvC', 'Angel Lormand',      1, '#E91E8C'),
  ('opl', '00000000-0000-0000-0000-000000000001', 'Opelousas',    'OpC', 'Shaina Broussard',   3, '#8FA05C'),
  ('laf', '00000000-0000-0000-0000-000000000001', 'Lafayette',    'LC',  'Katie Wilbanks',     3, '#D6472A'),
  ('mid', '00000000-0000-0000-0000-000000000001', 'Midtown',      'MtC', 'SuAn McClure',       2, '#9B87C4'),
  ('brx', '00000000-0000-0000-0000-000000000001', 'Broussard',    'BrC', 'Lauren Smith',       2, '#3E7CC2'),
  ('nib', '00000000-0000-0000-0000-000000000001', 'New Iberia',   'NIC', 'Jared Robicheaux',   3, '#E8A868'),
  ('vpl', '00000000-0000-0000-0000-000000000001', 'Ville Platte', 'VPC', 'William Reiszner',   1, '#F2B705'),
  ('ynv', '00000000-0000-0000-0000-000000000001', 'Youngsville',  'YvC', 'Pastor Josh Mesa',   2, '#17A2A0'),
  -- Synthetic pseudo-campus for org-wide/Central-scoped data — satisfies the FK on projects/
  -- staff/profiles for the "central" sentinel that's checked all over the existing frontend
  -- (campus === "central"), instead of that being an unvalidated magic string.
  ('central', '00000000-0000-0000-0000-000000000001', 'Central', 'Central', null, 0, '#2B4C7E')
on conflict (id) do nothing;
