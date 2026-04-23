-- ════════════════════════════════════════════
-- Supabase Database Setup
-- LINE Construction Report Bot
-- ════════════════════════════════════════════
-- วิธีใช้: ไปที่ Supabase Dashboard → SQL Editor
--          วางโค้ดนี้แล้วกด "Run"
-- ════════════════════════════════════════════

-- 1. สร้างตาราง line_reports
CREATE TABLE IF NOT EXISTS line_reports (
    id              BIGINT GENERATED ALWAYS AS IDENTITY PRIMARY KEY,
    timestamp       TIMESTAMPTZ     NOT NULL DEFAULT now(),
    user_id         TEXT,
    message_type    TEXT            CHECK (message_type IN ('text', 'image')),
    raw_text        TEXT,
    work_date       DATE,
    activities      JSONB           DEFAULT '[]'::jsonb,
    quantities      JSONB           DEFAULT '[]'::jsonb,
    workers         INTEGER,
    weather         TEXT,
    image_url       TEXT,
    image_filename  TEXT,
    created_at      TIMESTAMPTZ     NOT NULL DEFAULT now()
);

-- 2. Index สำหรับค้นหาตามวันที่ (ทำให้ query เร็วขึ้น)
CREATE INDEX IF NOT EXISTS idx_line_reports_work_date  ON line_reports (work_date);
CREATE INDEX IF NOT EXISTS idx_line_reports_timestamp  ON line_reports (timestamp DESC);
CREATE INDEX IF NOT EXISTS idx_line_reports_user_id    ON line_reports (user_id);
CREATE INDEX IF NOT EXISTS idx_line_reports_msg_type   ON line_reports (message_type);

-- 3. ปิด Row Level Security (สำหรับ server-side ใช้ service_role key)
--    ถ้าต้องการความปลอดภัยเพิ่มเติมสามารถเปิดได้ภายหลัง
ALTER TABLE line_reports DISABLE ROW LEVEL SECURITY;

-- 4. Grant สิทธิ์ให้ anon role (ใช้กับ SUPABASE_KEY)
GRANT ALL ON line_reports TO anon;
GRANT ALL ON line_reports TO authenticated;
GRANT USAGE, SELECT ON ALL SEQUENCES IN SCHEMA public TO anon;
GRANT USAGE, SELECT ON ALL SEQUENCES IN SCHEMA public TO authenticated;

-- ════════════════════════════════════════════
-- สร้าง Storage Bucket สำหรับรูปภาพ
-- (ทำใน Supabase Dashboard → Storage → New bucket)
-- ชื่อ bucket: construction-images
-- Public bucket: ✅ เปิด (เพื่อให้ดูรูปได้)
-- ════════════════════════════════════════════

-- ตัวอย่าง Query ที่ใช้บ่อย:

-- ดูรายงานวันนี้
-- SELECT * FROM line_reports WHERE work_date = CURRENT_DATE ORDER BY timestamp;

-- ดูรายงานสัปดาห์นี้
-- SELECT * FROM line_reports
--   WHERE work_date >= DATE_TRUNC('week', CURRENT_DATE)
--   ORDER BY work_date, timestamp;

-- สรุปงานตามวัน
-- SELECT work_date,
--        COUNT(*) FILTER (WHERE message_type = 'text')  AS text_count,
--        COUNT(*) FILTER (WHERE message_type = 'image') AS image_count,
--        STRING_AGG(DISTINCT weather, ', ')              AS weather
-- FROM line_reports
-- GROUP BY work_date
-- ORDER BY work_date DESC;
