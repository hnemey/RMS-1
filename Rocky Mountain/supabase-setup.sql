-- ============================================================
--  RMS Platform — Supabase Database Setup
--  Run this entire file in the Supabase SQL Editor
--  (Dashboard → SQL Editor → New query → paste → Run)
-- ============================================================

-- ── 1. EMPLOYEES ─────────────────────────────────────────────
CREATE TABLE IF NOT EXISTS employees (
  id            UUID PRIMARY KEY DEFAULT gen_random_uuid(),
  created_at    TIMESTAMPTZ DEFAULT NOW(),
  name          TEXT NOT NULL UNIQUE,
  role          TEXT DEFAULT 'Field Staff',
  pin           TEXT DEFAULT '1234',
  commission_rate DECIMAL(5,4) DEFAULT 0.05,
  ytd_goal      DECIMAL(10,2) DEFAULT 20000,
  active        BOOLEAN DEFAULT TRUE
);

-- ── 2. LEADS ─────────────────────────────────────────────────
CREATE TABLE IF NOT EXISTS leads (
  id               BIGSERIAL PRIMARY KEY,
  created_at       TIMESTAMPTZ DEFAULT NOW(),
  submitter_name   TEXT NOT NULL,
  opportunity_type TEXT NOT NULL,
  description      TEXT NOT NULL,
  urgency          TEXT NOT NULL DEFAULT 'Medium',
  contact_name     TEXT,
  phone            TEXT,
  address          TEXT,
  lat              DECIMAL(9,6),
  lng              DECIMAL(9,6),
  photo_urls       TEXT[],
  status           TEXT DEFAULT 'New',
  assigned_to      TEXT,
  notes            TEXT,
  job_id           BIGINT
);

-- ── 3. JOBS ──────────────────────────────────────────────────
CREATE TABLE IF NOT EXISTS jobs (
  id             BIGSERIAL PRIMARY KEY,
  created_at     TIMESTAMPTZ DEFAULT NOW(),
  job_name       TEXT NOT NULL,
  client_name    TEXT,
  job_type       TEXT,
  location       TEXT,
  bid_number     TEXT,
  job_date       DATE,
  status         TEXT DEFAULT 'Active',
  quoted_price   DECIMAL(10,2),
  direct_cost    DECIMAL(10,2),
  margin_pct     DECIMAL(5,2),
  actual_cost    DECIMAL(10,2),
  actual_revenue DECIMAL(10,2),
  salesperson    TEXT,
  foreman        TEXT,
  costing_data   JSONB,
  lead_id        BIGINT REFERENCES leads(id),
  completed_date DATE
);

-- Link leads → jobs
ALTER TABLE leads ADD CONSTRAINT leads_job_id_fk
  FOREIGN KEY (job_id) REFERENCES jobs(id);

-- ── 4. JOB NOTES ─────────────────────────────────────────────
CREATE TABLE IF NOT EXISTS job_notes (
  id          BIGSERIAL PRIMARY KEY,
  created_at  TIMESTAMPTZ DEFAULT NOW(),
  job_id      BIGINT NOT NULL REFERENCES jobs(id) ON DELETE CASCADE,
  author      TEXT NOT NULL,
  note_text   TEXT,
  photo_urls  TEXT[]
);

-- ── 5. COMMISSIONS ───────────────────────────────────────────
CREATE TABLE IF NOT EXISTS commissions (
  id             BIGSERIAL PRIMARY KEY,
  created_at     TIMESTAMPTZ DEFAULT NOW(),
  job_id         BIGINT REFERENCES jobs(id),
  job_name       TEXT,
  employee_name  TEXT NOT NULL,
  service_type   TEXT,
  period         TEXT NOT NULL,
  revenue        DECIMAL(10,2),
  rate           DECIMAL(5,4),
  amount         DECIMAL(10,2)
);

-- ── 6. ROW LEVEL SECURITY (open for internal tools) ──────────
ALTER TABLE employees   ENABLE ROW LEVEL SECURITY;
ALTER TABLE leads        ENABLE ROW LEVEL SECURITY;
ALTER TABLE jobs         ENABLE ROW LEVEL SECURITY;
ALTER TABLE job_notes    ENABLE ROW LEVEL SECURITY;
ALTER TABLE commissions  ENABLE ROW LEVEL SECURITY;

-- Allow anon key full access (internal tool — no public sign-up)
CREATE POLICY "anon_all" ON employees   FOR ALL TO anon USING (true) WITH CHECK (true);
CREATE POLICY "anon_all" ON leads        FOR ALL TO anon USING (true) WITH CHECK (true);
CREATE POLICY "anon_all" ON jobs         FOR ALL TO anon USING (true) WITH CHECK (true);
CREATE POLICY "anon_all" ON job_notes    FOR ALL TO anon USING (true) WITH CHECK (true);
CREATE POLICY "anon_all" ON commissions  FOR ALL TO anon USING (true) WITH CHECK (true);

-- ── 7. STORAGE BUCKET ────────────────────────────────────────
-- Run this too (or create the bucket via Dashboard → Storage → New Bucket)
INSERT INTO storage.buckets (id, name, public)
VALUES ('rms-photos', 'rms-photos', true)
ON CONFLICT DO NOTHING;

CREATE POLICY "anon_upload" ON storage.objects
  FOR INSERT TO anon
  WITH CHECK (bucket_id = 'rms-photos');

CREATE POLICY "anon_read" ON storage.objects
  FOR SELECT TO anon
  USING (bucket_id = 'rms-photos');

-- ── 8. SEED DATA — Add your real employees here ───────────────
-- Replace with actual names, roles, PINs, and commission rates
INSERT INTO employees (name, role, pin, commission_rate, ytd_goal) VALUES
  ('Harry Nemelka',    'Sales / Management',  '1111', 0.06, 30000),
  ('Crew Lead 1',      'Crew Lead',           '2222', 0.04, 20000),
  ('Crew Lead 2',      'Crew Lead',           '3333', 0.04, 20000),
  ('Sales Rep 1',      'Sales Representative','4444', 0.05, 25000)
ON CONFLICT (name) DO NOTHING;

-- ============================================================
--  DONE! After running this:
--  1. Copy your Project URL and anon key from:
--     Dashboard → Settings → API
--  2. Paste them into the SUPABASE_URL and SUPABASE_ANON_KEY
--     constants at the top of:
--       - lead.html
--       - portal.html
--       - estimator.html
--       - (RMS-Commission-Card) index.html
-- ============================================================
