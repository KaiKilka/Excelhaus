-- Halloween Trip 2025 – Supabase Setup
-- Dieses SQL im Supabase SQL-Editor ausführen

CREATE TABLE IF NOT EXISTS votes (
  session_id TEXT PRIMARY KEY,
  gender TEXT,
  ratings JSONB DEFAULT '{}',
  updated_at TIMESTAMPTZ DEFAULT NOW()
);

-- Realtime aktivieren
ALTER PUBLICATION supabase_realtime ADD TABLE votes;

-- Row Level Security aktivieren (öffentlicher Lese- und Schreibzugriff für die Abstimmung)
ALTER TABLE votes ENABLE ROW LEVEL SECURITY;

CREATE POLICY "Public read" ON votes FOR SELECT USING (true);
CREATE POLICY "Public upsert" ON votes FOR INSERT WITH CHECK (true);
CREATE POLICY "Public update" ON votes FOR UPDATE USING (true) WITH CHECK (true);
