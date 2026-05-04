-- Problem & Decision Board tables for Hua Ro Project Control Platform
-- Run once in Supabase SQL Editor.

CREATE TABLE IF NOT EXISTS project_issues (
    id              BIGINT GENERATED ALWAYS AS IDENTITY PRIMARY KEY,
    created_at      TIMESTAMPTZ NOT NULL DEFAULT now(),
    updated_at      TIMESTAMPTZ NOT NULL DEFAULT now(),
    work_date       DATE,
    source_id       BIGINT REFERENCES line_reports(id) ON DELETE SET NULL,
    title           TEXT NOT NULL,
    description     TEXT,
    area            TEXT,
    owner           TEXT,
    due_date        DATE,
    status          TEXT NOT NULL DEFAULT 'open'
                    CHECK (status IN ('open', 'in_progress', 'waiting', 'resolved', 'closed')),
    impact          TEXT NOT NULL DEFAULT 'medium'
                    CHECK (impact IN ('low', 'medium', 'high', 'critical')),
    next_action     TEXT,
    reported_by     TEXT,
    source_channel  TEXT NOT NULL DEFAULT 'line'
                    CHECK (source_channel IN ('line', 'admin', 'system')),
    CONSTRAINT project_issues_title_not_blank CHECK (length(trim(title)) > 0)
);

CREATE TABLE IF NOT EXISTS project_issue_comments (
    id          BIGINT GENERATED ALWAYS AS IDENTITY PRIMARY KEY,
    issue_id    BIGINT NOT NULL REFERENCES project_issues(id) ON DELETE CASCADE,
    created_at  TIMESTAMPTZ NOT NULL DEFAULT now(),
    author      TEXT,
    comment     TEXT NOT NULL,
    CONSTRAINT project_issue_comments_not_blank CHECK (length(trim(comment)) > 0)
);

CREATE INDEX IF NOT EXISTS idx_project_issues_status
    ON project_issues (status);
CREATE INDEX IF NOT EXISTS idx_project_issues_due_date
    ON project_issues (due_date);
CREATE INDEX IF NOT EXISTS idx_project_issues_work_date
    ON project_issues (work_date);
CREATE INDEX IF NOT EXISTS idx_project_issue_comments_issue_id
    ON project_issue_comments (issue_id, created_at DESC);

ALTER TABLE project_issues ENABLE ROW LEVEL SECURITY;
ALTER TABLE project_issue_comments ENABLE ROW LEVEL SECURITY;

DROP POLICY IF EXISTS "service_role_all" ON project_issues;
DROP POLICY IF EXISTS "service_role_all" ON project_issue_comments;

CREATE POLICY "service_role_all" ON project_issues
    FOR ALL TO service_role USING (true) WITH CHECK (true);

CREATE POLICY "service_role_all" ON project_issue_comments
    FOR ALL TO service_role USING (true) WITH CHECK (true);

-- Optional verification after running:
-- SELECT tablename, rowsecurity
-- FROM pg_tables
-- WHERE schemaname = 'public'
--   AND tablename IN ('project_issues', 'project_issue_comments');
