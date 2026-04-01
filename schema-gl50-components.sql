-- ════════════════════════════════════════════════════════════════════════════
-- GL50 COMPONENTS TABLE
-- Replaces the hardcoded GL50_COMPONENTS array in gl50.html.
-- Add/remove cards by inserting rows or toggling `active`.
-- Reorder cards by updating `sort_order`.
-- ════════════════════════════════════════════════════════════════════════════

create table if not exists gl50_components (
  id             uuid        default gen_random_uuid() primary key,
  set_id         uuid        references sets(id) on delete cascade,
  set_name       text        not null,            -- denormalized for display
  player         text        not null,
  variant        text        not null,
  active         boolean     not null default true, -- toggle without deleting
  included_since date        not null default current_date,
  sort_order     int,                              -- controls table display order
  created_at     timestamptz not null default now(),

  unique (set_id, player, variant)                -- prevent duplicate components
);

-- Public read access (same as every other table in this project)
alter table gl50_components enable row level security;

drop policy if exists "public read gl50_components" on gl50_components;
create policy "public read gl50_components"
  on gl50_components for select using (true);

-- ── Seed: 77 components sourced from CL50 Components CSV (GL50 = 1) ──────────
-- included_since = 2025-11-01 (GL50 index inception date)

insert into gl50_components (set_id, set_name, player, variant, included_since, sort_order) values

  -- MLB 2025 Game Face
  ('0401c373-55a1-43fd-84d0-944bb2421527', 'MLB 2025 Game Face',    'Aaron Judge',          'Base',         '2025-11-01',  1),
  ('0401c373-55a1-43fd-84d0-944bb2421527', 'MLB 2025 Game Face',    'Aaron Judge',          'Victory',      '2025-11-01',  2),
  ('0401c373-55a1-43fd-84d0-944bb2421527', 'MLB 2025 Game Face',    'Aaron Judge',          'Emerald',      '2025-11-01',  3),
  ('0401c373-55a1-43fd-84d0-944bb2421527', 'MLB 2025 Game Face',    'Aaron Judge',          'Gold',         '2025-11-01',  4),
  ('0401c373-55a1-43fd-84d0-944bb2421527', 'MLB 2025 Game Face',    'Aaron Judge',          'Chrome',       '2025-11-01',  5),
  ('0401c373-55a1-43fd-84d0-944bb2421527', 'MLB 2025 Game Face',    'Mike Trout',           'Victory',      '2025-11-01',  6),
  ('0401c373-55a1-43fd-84d0-944bb2421527', 'MLB 2025 Game Face',    'Mike Trout',           'Emerald',      '2025-11-01',  7),
  ('0401c373-55a1-43fd-84d0-944bb2421527', 'MLB 2025 Game Face',    'Mike Trout',           'Gold',         '2025-11-01',  8),
  ('0401c373-55a1-43fd-84d0-944bb2421527', 'MLB 2025 Game Face',    'Mike Trout',           'Chrome',       '2025-11-01',  9),
  ('0401c373-55a1-43fd-84d0-944bb2421527', 'MLB 2025 Game Face',    'Paul Skenes',          'Base',         '2025-11-01', 10),
  ('0401c373-55a1-43fd-84d0-944bb2421527', 'MLB 2025 Game Face',    'Paul Skenes',          'Victory',      '2025-11-01', 11),
  ('0401c373-55a1-43fd-84d0-944bb2421527', 'MLB 2025 Game Face',    'Paul Skenes',          'Emerald',      '2025-11-01', 12),
  ('0401c373-55a1-43fd-84d0-944bb2421527', 'MLB 2025 Game Face',    'Paul Skenes',          'Gold',         '2025-11-01', 13),
  ('0401c373-55a1-43fd-84d0-944bb2421527', 'MLB 2025 Game Face',    'Paul Skenes',          'Chrome',       '2025-11-01', 14),
  ('0401c373-55a1-43fd-84d0-944bb2421527', 'MLB 2025 Game Face',    'Shohei Ohtani',        'Base',         '2025-11-01', 15),
  ('0401c373-55a1-43fd-84d0-944bb2421527', 'MLB 2025 Game Face',    'Shohei Ohtani',        'Victory',      '2025-11-01', 16),
  ('0401c373-55a1-43fd-84d0-944bb2421527', 'MLB 2025 Game Face',    'Shohei Ohtani',        'Emerald',      '2025-11-01', 17),
  ('0401c373-55a1-43fd-84d0-944bb2421527', 'MLB 2025 Game Face',    'Shohei Ohtani',        'Gold',         '2025-11-01', 18),
  ('0401c373-55a1-43fd-84d0-944bb2421527', 'MLB 2025 Game Face',    'Shohei Ohtani',        'Chrome',       '2025-11-01', 19),
  ('0401c373-55a1-43fd-84d0-944bb2421527', 'MLB 2025 Game Face',    'Trophy',               'Emerald',      '2025-11-01', 20),
  ('0401c373-55a1-43fd-84d0-944bb2421527', 'MLB 2025 Game Face',    'Trophy',               'Gold',         '2025-11-01', 21),
  ('0401c373-55a1-43fd-84d0-944bb2421527', 'MLB 2025 Game Face',    'Trophy',               'Chrome',       '2025-11-01', 22),

  -- WNBA 2024 Game Face
  ('723c8d7d-e1be-42a3-a704-0547878664c6', 'WNBA 2024 Game Face',  'A''Ja Wilson',         'Base',         '2025-11-01', 23),
  ('723c8d7d-e1be-42a3-a704-0547878664c6', 'WNBA 2024 Game Face',  'A''Ja Wilson',         'Victory',      '2025-11-01', 24),
  ('723c8d7d-e1be-42a3-a704-0547878664c6', 'WNBA 2024 Game Face',  'A''Ja Wilson',         'Bet On Women', '2025-11-01', 25),
  ('723c8d7d-e1be-42a3-a704-0547878664c6', 'WNBA 2024 Game Face',  'A''Ja Wilson',         'Gold',         '2025-11-01', 26),
  ('723c8d7d-e1be-42a3-a704-0547878664c6', 'WNBA 2024 Game Face',  'A''Ja Wilson',         'Chrome',       '2025-11-01', 27),
  ('723c8d7d-e1be-42a3-a704-0547878664c6', 'WNBA 2024 Game Face',  'Angel Reese',          'Base',         '2025-11-01', 28),
  ('723c8d7d-e1be-42a3-a704-0547878664c6', 'WNBA 2024 Game Face',  'Angel Reese',          'Victory',      '2025-11-01', 29),
  ('723c8d7d-e1be-42a3-a704-0547878664c6', 'WNBA 2024 Game Face',  'Angel Reese',          'Bet On Women', '2025-11-01', 30),
  ('723c8d7d-e1be-42a3-a704-0547878664c6', 'WNBA 2024 Game Face',  'Angel Reese',          'Gold',         '2025-11-01', 31),
  ('723c8d7d-e1be-42a3-a704-0547878664c6', 'WNBA 2024 Game Face',  'Angel Reese',          'Chrome',       '2025-11-01', 32),
  ('723c8d7d-e1be-42a3-a704-0547878664c6', 'WNBA 2024 Game Face',  'Aubrey Plaza',         'Base',         '2025-11-01', 33),
  ('723c8d7d-e1be-42a3-a704-0547878664c6', 'WNBA 2024 Game Face',  'Aubrey Plaza',         'Victory',      '2025-11-01', 34),
  ('723c8d7d-e1be-42a3-a704-0547878664c6', 'WNBA 2024 Game Face',  'Aubrey Plaza',         'Gold',         '2025-11-01', 35),
  ('723c8d7d-e1be-42a3-a704-0547878664c6', 'WNBA 2024 Game Face',  'Aubrey Plaza',         'Chrome',       '2025-11-01', 36),
  ('723c8d7d-e1be-42a3-a704-0547878664c6', 'WNBA 2024 Game Face',  'Caitlin Clark',        'Base',         '2025-11-01', 37),
  ('723c8d7d-e1be-42a3-a704-0547878664c6', 'WNBA 2024 Game Face',  'Caitlin Clark',        'Victory',      '2025-11-01', 38),
  ('723c8d7d-e1be-42a3-a704-0547878664c6', 'WNBA 2024 Game Face',  'Caitlin Clark',        'Bet On Women', '2025-11-01', 39),
  ('723c8d7d-e1be-42a3-a704-0547878664c6', 'WNBA 2024 Game Face',  'Caitlin Clark',        'Gold',         '2025-11-01', 40),
  ('723c8d7d-e1be-42a3-a704-0547878664c6', 'WNBA 2024 Game Face',  'Caitlin Clark',        'Chrome',       '2025-11-01', 41),
  ('723c8d7d-e1be-42a3-a704-0547878664c6', 'WNBA 2024 Game Face',  'Sabrina Ionescu',      'Base',         '2025-11-01', 42),
  ('723c8d7d-e1be-42a3-a704-0547878664c6', 'WNBA 2024 Game Face',  'Sabrina Ionescu',      'Victory',      '2025-11-01', 43),
  ('723c8d7d-e1be-42a3-a704-0547878664c6', 'WNBA 2024 Game Face',  'Sabrina Ionescu',      'Bet On Women', '2025-11-01', 44),
  ('723c8d7d-e1be-42a3-a704-0547878664c6', 'WNBA 2024 Game Face',  'Sabrina Ionescu',      'Gold',         '2025-11-01', 45),
  ('723c8d7d-e1be-42a3-a704-0547878664c6', 'WNBA 2024 Game Face',  'Sabrina Ionescu',      'Chrome',       '2025-11-01', 46),
  ('723c8d7d-e1be-42a3-a704-0547878664c6', 'WNBA 2024 Game Face',  'Trophy',               'Gold',         '2025-11-01', 47),
  ('723c8d7d-e1be-42a3-a704-0547878664c6', 'WNBA 2024 Game Face',  'Trophy',               'Chrome',       '2025-11-01', 48),

  -- NBA 2024-25 Game Face
  ('af4115ee-e310-4538-af0a-0386d1cb189d', 'NBA 2024-25 Game Face', 'Cooper Flagg',         'SP',           '2025-11-01', 49),
  ('af4115ee-e310-4538-af0a-0386d1cb189d', 'NBA 2024-25 Game Face', 'Eminem',               'Base',         '2025-11-01', 50),
  ('af4115ee-e310-4538-af0a-0386d1cb189d', 'NBA 2024-25 Game Face', 'Eminem',               'Emerald',      '2025-11-01', 51),
  ('af4115ee-e310-4538-af0a-0386d1cb189d', 'NBA 2024-25 Game Face', 'Eminem',               'Gold',         '2025-11-01', 52),
  ('af4115ee-e310-4538-af0a-0386d1cb189d', 'NBA 2024-25 Game Face', 'Eminem',               'Chrome',       '2025-11-01', 53),
  ('af4115ee-e310-4538-af0a-0386d1cb189d', 'NBA 2024-25 Game Face', 'Lebron James',         'SP',           '2025-11-01', 54),
  ('af4115ee-e310-4538-af0a-0386d1cb189d', 'NBA 2024-25 Game Face', 'Luka Doncic',          'Base',         '2025-11-01', 55),
  ('af4115ee-e310-4538-af0a-0386d1cb189d', 'NBA 2024-25 Game Face', 'Luka Doncic',          'Victory',      '2025-11-01', 56),
  ('af4115ee-e310-4538-af0a-0386d1cb189d', 'NBA 2024-25 Game Face', 'Luka Doncic',          'Emerald',      '2025-11-01', 57),
  ('af4115ee-e310-4538-af0a-0386d1cb189d', 'NBA 2024-25 Game Face', 'Luka Doncic',          'Gold',         '2025-11-01', 58),
  ('af4115ee-e310-4538-af0a-0386d1cb189d', 'NBA 2024-25 Game Face', 'Luka Doncic',          'Chrome',       '2025-11-01', 59),
  ('af4115ee-e310-4538-af0a-0386d1cb189d', 'NBA 2024-25 Game Face', 'Stephen Curry',        'Base',         '2025-11-01', 60),
  ('af4115ee-e310-4538-af0a-0386d1cb189d', 'NBA 2024-25 Game Face', 'Stephen Curry',        'Victory',      '2025-11-01', 61),
  ('af4115ee-e310-4538-af0a-0386d1cb189d', 'NBA 2024-25 Game Face', 'Stephen Curry',        'Emerald',      '2025-11-01', 62),
  ('af4115ee-e310-4538-af0a-0386d1cb189d', 'NBA 2024-25 Game Face', 'Stephen Curry',        'Gold',         '2025-11-01', 63),
  ('af4115ee-e310-4538-af0a-0386d1cb189d', 'NBA 2024-25 Game Face', 'Stephen Curry',        'Chrome',       '2025-11-01', 64),
  ('af4115ee-e310-4538-af0a-0386d1cb189d', 'NBA 2024-25 Game Face', 'Trophy',               'Emerald',      '2025-11-01', 65),
  ('af4115ee-e310-4538-af0a-0386d1cb189d', 'NBA 2024-25 Game Face', 'Trophy',               'Gold',         '2025-11-01', 66),
  ('af4115ee-e310-4538-af0a-0386d1cb189d', 'NBA 2024-25 Game Face', 'Trophy',               'Chrome',       '2025-11-01', 67),
  ('af4115ee-e310-4538-af0a-0386d1cb189d', 'NBA 2024-25 Game Face', 'Victor Wembanyama',    'Base',         '2025-11-01', 68),
  ('af4115ee-e310-4538-af0a-0386d1cb189d', 'NBA 2024-25 Game Face', 'Victor Wembanyama',    'Victory',      '2025-11-01', 69),
  ('af4115ee-e310-4538-af0a-0386d1cb189d', 'NBA 2024-25 Game Face', 'Victor Wembanyama',    'Emerald',      '2025-11-01', 70),
  ('af4115ee-e310-4538-af0a-0386d1cb189d', 'NBA 2024-25 Game Face', 'Victor Wembanyama',    'Gold',         '2025-11-01', 71),
  ('af4115ee-e310-4538-af0a-0386d1cb189d', 'NBA 2024-25 Game Face', 'Victor Wembanyama',    'Chrome',       '2025-11-01', 72),

  -- Eminem Home
  ('d9d613cd-efef-48e0-9fcd-9eb3faceed12', 'Eminem Home',           'Eminem',               'Base',         '2025-11-01', 73),
  ('d9d613cd-efef-48e0-9fcd-9eb3faceed12', 'Eminem Home',           'Eminem',               'Saphire',      '2025-11-01', 74),
  ('d9d613cd-efef-48e0-9fcd-9eb3faceed12', 'Eminem Home',           'Eminem',               'Emerald',      '2025-11-01', 75),
  ('d9d613cd-efef-48e0-9fcd-9eb3faceed12', 'Eminem Home',           'Eminem',               'Gold',         '2025-11-01', 76),
  ('d9d613cd-efef-48e0-9fcd-9eb3faceed12', 'Eminem Home',           'Eminem',               'Chrome',       '2025-11-01', 77)

on conflict (set_id, player, variant) do nothing; -- safe to re-run
