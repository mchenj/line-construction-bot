-- ════════════════════════════════════════════════════════════
-- Supabase Security: Enable Row-Level Security (RLS)
-- ════════════════════════════════════════════════════════════
-- รันใน Supabase SQL Editor ครั้งเดียว
-- หลังรัน: bot ใช้ service_role key → ทำงานได้เต็ม
--          คนนอกมี anon key → เข้าตารางไม่ได้
-- ════════════════════════════════════════════════════════════

-- 1) เปิด RLS ทุกตาราง
ALTER TABLE daily_reports      ENABLE ROW LEVEL SECURITY;
ALTER TABLE report_activities  ENABLE ROW LEVEL SECURITY;
ALTER TABLE report_images      ENABLE ROW LEVEL SECURITY;
ALTER TABLE line_reports       ENABLE ROW LEVEL SECURITY;
ALTER TABLE IF EXISTS project_issues ENABLE ROW LEVEL SECURITY;
ALTER TABLE IF EXISTS project_issue_comments ENABLE ROW LEVEL SECURITY;

-- 2) ลบ policy เก่าก่อน (กันรันซ้ำแล้ว error)
DROP POLICY IF EXISTS "service_role_all" ON daily_reports;
DROP POLICY IF EXISTS "service_role_all" ON report_activities;
DROP POLICY IF EXISTS "service_role_all" ON report_images;
DROP POLICY IF EXISTS "service_role_all" ON line_reports;

-- 3) สร้าง policy: service_role เข้าได้เต็ม (CRUD)
CREATE POLICY "service_role_all" ON daily_reports
    FOR ALL TO service_role USING (true) WITH CHECK (true);

CREATE POLICY "service_role_all" ON report_activities
    FOR ALL TO service_role USING (true) WITH CHECK (true);

CREATE POLICY "service_role_all" ON report_images
    FOR ALL TO service_role USING (true) WITH CHECK (true);

CREATE POLICY "service_role_all" ON line_reports
    FOR ALL TO service_role USING (true) WITH CHECK (true);

DO $$
BEGIN
    IF to_regclass('public.project_issues') IS NOT NULL THEN
        DROP POLICY IF EXISTS "service_role_all" ON project_issues;
        CREATE POLICY "service_role_all" ON project_issues
            FOR ALL TO service_role USING (true) WITH CHECK (true);
    END IF;

    IF to_regclass('public.project_issue_comments') IS NOT NULL THEN
        DROP POLICY IF EXISTS "service_role_all" ON project_issue_comments;
        CREATE POLICY "service_role_all" ON project_issue_comments
            FOR ALL TO service_role USING (true) WITH CHECK (true);
    END IF;
END $$;

-- ════════════════════════════════════════════════════════════
-- ตรวจผลหลังรัน (optional — รันแยกเพื่อ verify)
-- ════════════════════════════════════════════════════════════
-- SELECT tablename, rowsecurity
-- FROM pg_tables
-- WHERE schemaname = 'public'
--   AND tablename IN ('daily_reports','report_activities','report_images','line_reports');
--
-- ผลที่ควรได้: rowsecurity = true ทุกแถว ✓
