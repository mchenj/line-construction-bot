"""
Admin control center for LINE Construction Bot.

Routes:
- GET  /admin
- GET  /admin/api/overview
- GET  /admin/download/{kind}
- POST /admin/upload/{kind}
- POST /admin/trigger/weekly
- GET  /admin/check_fonts
"""

from __future__ import annotations

import os
import subprocess
from collections import Counter
from datetime import date, datetime, timedelta, timezone
from html import escape
from pathlib import Path
from urllib.parse import quote

from fastapi import APIRouter, File, HTTPException, UploadFile
from fastapi.responses import FileResponse, HTMLResponse, JSONResponse, RedirectResponse

try:
    from zoneinfo import ZoneInfo
except Exception:  # pragma: no cover - Python always has zoneinfo on Railway's current stack.
    ZoneInfo = None

try:
    from supabase import create_client
except Exception:  # pragma: no cover - lets the admin page render even if deps are missing locally.
    create_client = None


_THIS_DIR = Path(__file__).parent
_DATA_DIR = _THIS_DIR / "data"
_DATA_DIR.mkdir(exist_ok=True)

DATA_FILES = {
    "plan": (_DATA_DIR / "construction_plan.xlsx", "Construction Plan"),
    "cm": (_DATA_DIR / "cm_personnel.xlsx", "CM Personnel"),
}

PROJECT_DISPLAY_NAME = "โครงการพัฒนาพื้นที่ชุมชนหัวรอและพื้นที่ต่อเนื่อง ตำบลหัวรอ อำเภอเมืองหัวรอฯ"

router = APIRouter(prefix="/admin", tags=["admin"])


def _now_bkk() -> datetime:
    if ZoneInfo:
        try:
            return datetime.now(ZoneInfo("Asia/Bangkok"))
        except Exception:
            pass
    return datetime.now(timezone(timedelta(hours=7)))


def _check_token(token: str):
    expected = os.getenv("ADMIN_TOKEN", "")
    if not expected:
        raise HTTPException(403, "ADMIN_TOKEN is not configured.")
    if token != expected:
        raise HTTPException(401, "Invalid admin token.")


def _file_info(path: Path) -> dict:
    if not path.exists():
        return {"exists": False, "size": 0, "modified": None, "age_hours": None}
    st = path.stat()
    modified = datetime.fromtimestamp(st.st_mtime, tz=_now_bkk().tzinfo)
    age_hours = max(0, (_now_bkk() - modified).total_seconds() / 3600)
    return {
        "exists": True,
        "size": st.st_size,
        "size_label": _format_bytes(st.st_size),
        "modified": modified.strftime("%Y-%m-%d %H:%M"),
        "age_hours": age_hours,
    }


def _format_bytes(size: int) -> str:
    value = float(size)
    for unit in ("B", "KB", "MB", "GB"):
        if value < 1024 or unit == "GB":
            return f"{value:.1f} {unit}" if unit != "B" else f"{int(value)} B"
        value /= 1024
    return f"{value:.1f} GB"


def _parse_iso_dt(value: str | None) -> datetime | None:
    if not value:
        return None
    try:
        text = str(value).replace("Z", "+00:00")
        parsed = datetime.fromisoformat(text)
        if parsed.tzinfo is None:
            parsed = parsed.replace(tzinfo=timezone.utc)
        return parsed.astimezone(_now_bkk().tzinfo)
    except Exception:
        return None


def _parse_date(value: str | None) -> date | None:
    if not value:
        return None
    try:
        return date.fromisoformat(str(value)[:10])
    except Exception:
        return None


def _short(text: object, length: int = 92) -> str:
    clean = " ".join(str(text or "").split())
    if len(clean) <= length:
        return clean
    return clean[: max(0, length - 1)].rstrip() + "..."


def _get_supabase_client():
    if create_client is None:
        return None, "The supabase package is not available."
    url = os.getenv("SUPABASE_URL", "").strip()
    key = os.getenv("SUPABASE_KEY", "").strip()
    if not url or not key:
        return None, "SUPABASE_URL or SUPABASE_KEY is missing."
    try:
        return create_client(url, key), None
    except Exception as exc:
        return None, str(exc)


def _safe_query(label: str, fn):
    try:
        result = fn()
        return list(result.data or []), None
    except Exception as exc:
        return [], f"{label}: {exc}"


def _env_status() -> dict:
    weekly_enabled = os.getenv("WEEKLY_CRON_ENABLED", "false").lower() in ("true", "1", "yes")
    railway_name = os.getenv("RAILWAY_SERVICE_NAME") or os.getenv("RAILWAY_PROJECT_NAME") or "Local runtime"
    railway_domain = os.getenv("RAILWAY_PUBLIC_DOMAIN") or os.getenv("RAILWAY_STATIC_URL") or ""
    weekly_hour = _safe_int(os.getenv("WEEKLY_CRON_HOUR", "17"), 17)
    weekly_minute = _safe_int(os.getenv("WEEKLY_CRON_MINUTE", "0"), 0)
    return {
        "line_secret": bool(os.getenv("LINE_CHANNEL_SECRET", "").strip()),
        "line_token": bool(os.getenv("LINE_CHANNEL_ACCESS_TOKEN", "").strip()),
        "supabase_url": bool(os.getenv("SUPABASE_URL", "").strip()),
        "supabase_key": bool(os.getenv("SUPABASE_KEY", "").strip()),
        "admin_token": bool(os.getenv("ADMIN_TOKEN", "").strip()),
        "weekly_enabled": weekly_enabled,
        "weekly_schedule": (
            f"{os.getenv('WEEKLY_CRON_DAY', 'fri').upper()} "
            f"{weekly_hour:02d}:"
            f"{weekly_minute:02d}"
        ),
        "weekly_format": os.getenv("WEEKLY_CRON_FORMAT", "zip").upper(),
        "weekly_targets": len([x for x in os.getenv("WEEKLY_CRON_USER_IDS", "").split(",") if x.strip()]),
        "railway_name": railway_name,
        "railway_domain": railway_domain,
        "port": os.getenv("PORT", "local"),
        "project_name": os.getenv("PROJECT_NAME", PROJECT_DISPLAY_NAME),
    }


def _safe_int(value: object, fallback: int) -> int:
    try:
        return int(str(value).strip().split()[0])
    except Exception:
        return fallback


def _collect_dashboard() -> dict:
    now = _now_bkk()
    today = now.date()
    start_14 = today - timedelta(days=13)
    start_30 = today - timedelta(days=29)
    env = _env_status()
    client, client_error = _get_supabase_client()

    data = {
        "generated_at": now.strftime("%Y-%m-%d %H:%M:%S"),
        "today": today.isoformat(),
        "env": env,
        "files": {kind: {"label": label, **_file_info(path)} for kind, (path, label) in DATA_FILES.items()},
        "db_ok": False,
        "db_error": client_error,
        "query_errors": [],
        "daily": [],
        "line_reports": [],
        "activities": [],
        "images": [],
    }

    if client is not None:
        daily, err = _safe_query(
            "daily_reports",
            lambda: client.table("daily_reports")
            .select(
                "work_date,total_workers,engineers,foremen,skilled_workers,"
                "laborers,water_level,weather_morning,report_status,updated_at"
            )
            .gte("work_date", start_30.isoformat())
            .order("work_date")
            .execute(),
        )
        data["daily"] = daily
        if err:
            data["query_errors"].append(err)

        line_reports, err = _safe_query(
            "line_reports",
            lambda: client.table("line_reports")
            .select("id,timestamp,user_id,message_type,work_date,raw_text,image_url")
            .gte("timestamp", (now - timedelta(days=30)).astimezone(timezone.utc).isoformat())
            .order("timestamp", desc=True)
            .limit(300)
            .execute(),
        )
        data["line_reports"] = line_reports
        if err:
            data["query_errors"].append(err)

        activities, err = _safe_query(
            "report_activities",
            lambda: client.table("report_activities")
            .select("work_date,activity_type,description,seq_no")
            .gte("work_date", start_30.isoformat())
            .order("work_date", desc=True)
            .limit(500)
            .execute(),
        )
        data["activities"] = activities
        if err:
            data["query_errors"].append(err)

        images, err = _safe_query(
            "report_images",
            lambda: client.table("report_images")
            .select("work_date,image_url,caption,taken_at")
            .gte("work_date", start_30.isoformat())
            .order("taken_at", desc=True)
            .limit(24)
            .execute(),
        )
        data["images"] = images
        if err:
            data["query_errors"].append(err)

        data["db_ok"] = any([daily, line_reports, activities, images]) or not data["query_errors"]
        if data["db_ok"]:
            data["db_error"] = None

    data.update(_summarize_dashboard(data, start_14, start_30))
    return data


def _summarize_dashboard(data: dict, start_14: date, start_30: date) -> dict:
    today = _parse_date(data["today"]) or _now_bkk().date()
    line_reports = data["line_reports"]
    daily = data["daily"]
    activities = data["activities"]
    images = data["images"]

    by_day = {start_14 + timedelta(days=i): {"reports": 0, "workers": 0, "images": 0} for i in range(14)}
    recent_by_type = Counter()
    unique_users = set()
    today_messages = 0
    week_messages = 0

    for row in line_reports:
        work_date = _parse_date(row.get("work_date")) or _parse_iso_dt(row.get("timestamp")) and _parse_iso_dt(row.get("timestamp")).date()
        msg_type = row.get("message_type") or "unknown"
        recent_by_type[msg_type] += 1
        if row.get("user_id"):
            unique_users.add(row.get("user_id"))
        if work_date == today:
            today_messages += 1
        if work_date and work_date >= today - timedelta(days=6):
            week_messages += 1
        if work_date in by_day:
            by_day[work_date]["reports"] += 1

    latest_daily = None
    worker_total_7d = 0
    water_points = []
    for row in daily:
        work_date = _parse_date(row.get("work_date"))
        if work_date in by_day:
            by_day[work_date]["workers"] = int(row.get("total_workers") or 0)
        if work_date and work_date >= today - timedelta(days=6):
            worker_total_7d += int(row.get("total_workers") or 0)
        if row.get("water_level") is not None and work_date:
            try:
                water_points.append({"date": work_date.isoformat(), "value": float(row.get("water_level"))})
            except Exception:
                pass
        if work_date and (latest_daily is None or work_date > _parse_date(latest_daily.get("work_date"))):
            latest_daily = row

    for img in images:
        work_date = _parse_date(img.get("work_date"))
        if work_date in by_day:
            by_day[work_date]["images"] += 1

    activity_counts = Counter((row.get("activity_type") or "general") for row in activities)
    activity_top = activity_counts.most_common(8)

    plan_ok = data["files"]["plan"]["exists"]
    cm_ok = data["files"]["cm"]["exists"]
    health_items = [
        data["env"]["line_secret"],
        data["env"]["line_token"],
        data["env"]["supabase_url"],
        data["env"]["supabase_key"],
        data["env"]["admin_token"],
        data["db_ok"],
        plan_ok,
        cm_ok,
    ]
    health_score = round((sum(1 for item in health_items if item) / len(health_items)) * 100)

    return {
        "health_score": health_score,
        "metrics": {
            "today_messages": today_messages,
            "week_messages": week_messages,
            "active_users": len(unique_users),
            "photos_30d": len(images),
            "worker_total_7d": worker_total_7d,
            "activity_types": len(activity_counts),
        },
        "latest_daily": latest_daily,
        "daily_series": [
            {
                "date": day.isoformat(),
                "label": day.strftime("%d %b"),
                "reports": values["reports"],
                "workers": values["workers"],
                "images": values["images"],
            }
            for day, values in by_day.items()
        ],
        "message_mix": dict(recent_by_type),
        "activity_top": activity_top,
        "water_points": water_points[-14:],
        "recent_events": _recent_events(line_reports[:12]),
    }


def _recent_events(rows: list[dict]) -> list[dict]:
    events = []
    for row in rows:
        dt = _parse_iso_dt(row.get("timestamp"))
        label = dt.strftime("%d %b %H:%M") if dt else str(row.get("timestamp") or "")
        message_type = str(row.get("message_type") or "message")
        events.append(
            {
                "time": label,
                "type": message_type.title(),
                "date": str(row.get("work_date") or ""),
                "text": _short(row.get("raw_text") or row.get("image_url") or "No text", 110),
            }
        )
    return events


def _status_word(ok: bool, paused: bool = False) -> str:
    if paused:
        return "Paused"
    return "Online" if ok else "Needs setup"


def _status_class(ok: bool, paused: bool = False) -> str:
    if paused:
        return "warn"
    return "ok" if ok else "bad"


def _render_metric(title: str, value: object, note: str, accent: str) -> str:
    return f"""
    <article class="metric" style="--accent:{escape(accent)}">
      <span class="metric-key">{escape(title)}</span>
      <strong>{escape(str(value))}</strong>
      <small>{escape(note)}</small>
    </article>
    """


def _render_report_chart(series: list[dict]) -> str:
    width, height = 720, 236
    left, top, bottom = 44, 22, 42
    plot_h = height - top - bottom
    plot_w = width - left - 20
    max_reports = max([int(item["reports"]) for item in series] + [1])
    max_workers = max([int(item["workers"]) for item in series] + [1])
    slot = plot_w / max(1, len(series))
    bar_w = max(14, min(28, slot * 0.45))
    parts = [
        f'<svg viewBox="0 0 {width} {height}" role="img" aria-label="Fourteen day report trend">'
        f'<rect width="{width}" height="{height}" fill="#fffaf0"/>'
    ]
    for i in range(5):
        y = top + plot_h * i / 4
        parts.append(f'<line x1="{left}" y1="{y:.1f}" x2="{width-12}" y2="{y:.1f}" stroke="#f1d7ad" stroke-width="1"/>')
    for idx, item in enumerate(series):
        x = left + idx * slot + (slot - bar_w) / 2
        report_h = (int(item["reports"]) / max_reports) * (plot_h - 8)
        worker_h = (int(item["workers"]) / max_workers) * (plot_h - 8)
        y1 = top + plot_h - report_h
        y2 = top + plot_h - worker_h
        parts.append(f'<rect x="{x:.1f}" y="{y1:.1f}" width="{bar_w:.1f}" height="{report_h:.1f}" rx="4" fill="#b91c1c"/>')
        parts.append(f'<rect x="{x + bar_w + 3:.1f}" y="{y2:.1f}" width="{max(5, bar_w * 0.32):.1f}" height="{worker_h:.1f}" rx="3" fill="#f59e0b"/>')
        label = escape(str(item["label"]))
        parts.append(
            f'<text x="{x + bar_w / 2:.1f}" y="{height-18}" text-anchor="middle" '
            f'font-size="11" fill="#7f1d1d">{label}</text>'
        )
    parts.append(f'<text x="{left}" y="15" font-size="12" fill="#7f1d1d">Reports and workers</text>')
    parts.append(f'<circle cx="{width-180}" cy="12" r="5" fill="#b91c1c"/><text x="{width-170}" y="16" font-size="11" fill="#7f1d1d">Messages</text>')
    parts.append(f'<circle cx="{width-95}" cy="12" r="5" fill="#f59e0b"/><text x="{width-85}" y="16" font-size="11" fill="#7f1d1d">Workers</text>')
    parts.append("</svg>")
    return "".join(parts)


def _render_activity_chart(activity_top: list[tuple[str, int]]) -> str:
    if not activity_top:
        activity_top = [("No activity yet", 1)]
    max_value = max(count for _, count in activity_top) or 1
    rows = []
    for name, count in activity_top:
        pct = max(5, round((count / max_value) * 100))
        rows.append(
            f"""
            <div class="bar-row">
              <span>{escape(str(name).replace("_", " ").title())}</span>
              <div class="bar-track"><i style="width:{pct}%"></i></div>
              <b>{count}</b>
            </div>
            """
        )
    return "".join(rows)


def _render_waterline(points: list[dict]) -> str:
    width, height = 560, 150
    if len(points) < 2:
        return (
            f'<svg viewBox="0 0 {width} {height}" role="img" aria-label="Water level chart">'
            f'<rect width="{width}" height="{height}" fill="#fffaf0"/>'
            f'<text x="24" y="78" font-size="13" fill="#7f1d1d">Waiting for water level data</text>'
            "</svg>"
        )
    values = [float(item["value"]) for item in points]
    min_v, max_v = min(values), max(values)
    span = max(max_v - min_v, 1)
    coords = []
    for idx, item in enumerate(points):
        x = 22 + idx * ((width - 44) / max(1, len(points) - 1))
        y = 20 + (max_v - float(item["value"])) / span * (height - 52)
        coords.append((x, y))
    poly = " ".join(f"{x:.1f},{y:.1f}" for x, y in coords)
    area = f"22,{height-24} {poly} {width-22},{height-24}"
    dots = "".join(f'<circle cx="{x:.1f}" cy="{y:.1f}" r="3.5" fill="#f59e0b"/>' for x, y in coords)
    return (
        f'<svg viewBox="0 0 {width} {height}" role="img" aria-label="Water level chart">'
        f'<rect width="{width}" height="{height}" fill="#fffaf0"/>'
        f'<line x1="22" y1="{height-24}" x2="{width-22}" y2="{height-24}" stroke="#f1d7ad"/>'
        f'<polygon points="{area}" fill="#fef3c7"/>'
        f'<polyline points="{poly}" fill="none" stroke="#b91c1c" stroke-width="3" stroke-linecap="round" stroke-linejoin="round"/>'
        f"{dots}"
        f'<text x="22" y="18" font-size="12" fill="#7f1d1d">Water level: {min_v:.2f} to {max_v:.2f} m</text>'
        "</svg>"
    )


def _render_files(files: dict, token_q: str) -> str:
    rows = []
    for kind, info in files.items():
        exists = bool(info["exists"])
        status = "Ready" if exists else "Missing"
        status_class = "ok" if exists else "bad"
        modified = info.get("modified") or "No file"
        size = info.get("size_label") or "0 B"
        rows.append(
            f"""
            <tr>
              <td>
                <strong>{escape(info["label"])}</strong>
                <span>{escape(DATA_FILES[kind][0].name)}</span>
              </td>
              <td><mark class="{status_class}">{status}</mark></td>
              <td>{escape(size)}</td>
              <td>{escape(modified)}</td>
              <td class="actions">
                <a class="icon-link" href="/admin/download/{escape(kind)}?token={token_q}" aria-label="Download {escape(info["label"])}">Download</a>
                <form action="/admin/upload/{escape(kind)}?token={token_q}" method="post" enctype="multipart/form-data">
                  <input type="file" name="file" accept=".xlsx" required>
                  <button type="submit">Upload</button>
                </form>
              </td>
            </tr>
            """
        )
    return "".join(rows)


def _render_recent(events: list[dict]) -> str:
    if not events:
        return '<p class="empty">No recent LINE messages found.</p>'
    items = []
    for event in events:
        items.append(
            f"""
            <li>
              <time>{escape(event["time"])}</time>
              <strong>{escape(event["type"])}</strong>
              <span>{escape(event["date"])}</span>
              <p>{escape(event["text"])}</p>
            </li>
            """
        )
    return "<ol class=\"timeline\">" + "".join(items) + "</ol>"


def _render_flow(data: dict) -> str:
    env = data["env"]
    line_ok = env["line_secret"] and env["line_token"]
    db_ok = data["db_ok"]
    admin_ok = env["admin_token"]
    cron_paused = not env["weekly_enabled"]
    nodes = [
        ("LINE", "Messaging API", line_ok, False),
        ("FastAPI", "Webhook + admin", admin_ok, False),
        ("Supabase", "Reports + images", db_ok, False),
        ("Railway", env["railway_name"], True, False),
        ("Weekly", f'{env["weekly_schedule"]} {env["weekly_format"]}', env["weekly_enabled"], cron_paused),
    ]
    html_nodes = []
    for title, note, ok, paused in nodes:
        cls = _status_class(ok, paused)
        html_nodes.append(
            f"""
            <div class="flow-node {cls}">
              <i></i>
              <strong>{escape(title)}</strong>
              <span>{escape(note)}</span>
              <b>{escape(_status_word(ok, paused))}</b>
            </div>
            """
        )
    return "".join(html_nodes)


def _render_latest_daily(row: dict | None) -> str:
    if not row:
        return '<p class="empty">No daily report summary found.</p>'
    fields = [
        ("Date", row.get("work_date")),
        ("Status", row.get("report_status") or "draft"),
        ("Workers", row.get("total_workers") or 0),
        ("Engineers", row.get("engineers") or 0),
        ("Foremen", row.get("foremen") or 0),
        ("Skilled", row.get("skilled_workers") or 0),
        ("Laborers", row.get("laborers") or 0),
        ("Water", "" if row.get("water_level") is None else f'{float(row.get("water_level")):.2f} m'),
        ("Weather", row.get("weather_morning") or "-"),
    ]
    cells = "".join(
        f"<div><span>{escape(label)}</span><strong>{escape(str(value))}</strong></div>" for label, value in fields
    )
    return f'<div class="fact-grid">{cells}</div>'


def _page_css() -> str:
    return """
    :root {
      --ink: #18110f;
      --muted: #75635a;
      --line: #ead8bd;
      --page: #fff7ed;
      --panel: #ffffff;
      --armor: #7f1d1d;
      --red: #b91c1c;
      --hot: #dc2626;
      --gold: #f59e0b;
      --arc: #facc15;
      --steel: #24201e;
      --ok: #15803d;
      font-family: Inter, ui-sans-serif, system-ui, -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif;
      color: var(--ink);
      background: var(--page);
    }
    * { box-sizing: border-box; }
    body {
      margin: 0;
      background:
        linear-gradient(180deg, #2b0d0d 0, #7f1d1d 132px, #fff7ed 133px, #fff7ed 100%);
    }
    a { color: inherit; }
    .shell { min-height: 100vh; }
    .topbar {
      position: sticky;
      top: 0;
      z-index: 5;
      display: flex;
      justify-content: space-between;
      gap: 18px;
      align-items: center;
      padding: 16px clamp(18px, 3vw, 42px);
      border-bottom: 1px solid rgba(250, 204, 21, .18);
      background: rgba(43, 13, 13, .94);
      backdrop-filter: blur(12px);
    }
    .brand { display: flex; gap: 12px; align-items: center; min-width: 0; color: #fff7ed; }
    .brand-mark {
      width: 42px;
      height: 42px;
      border-radius: 8px;
      background:
        linear-gradient(135deg, #facc15 0 42%, transparent 43%),
        linear-gradient(135deg, #dc2626 0 63%, #18110f 64% 100%);
      box-shadow: inset 0 0 0 1px rgba(250, 204, 21, .45), 0 10px 28px rgba(0, 0, 0, .22);
      flex: 0 0 auto;
    }
    .brand h1 { font-size: 18px; line-height: 1.15; margin: 0; letter-spacing: 0; }
    .brand span { display: block; color: #f8d594; font-size: 12px; margin-top: 3px; max-width: 720px; overflow-wrap: anywhere; }
    .top-actions { display: flex; gap: 10px; align-items: center; flex-wrap: wrap; justify-content: flex-end; }
    .pill {
      display: inline-flex;
      gap: 8px;
      align-items: center;
      border: 1px solid rgba(250, 204, 21, .25);
      border-radius: 999px;
      padding: 7px 11px;
      background: rgba(255, 247, 237, .09);
      color: #fff7ed;
      font-size: 12px;
      white-space: nowrap;
    }
    .dot { width: 8px; height: 8px; border-radius: 50%; background: var(--ok); display: inline-block; }
    .dot.warn { background: var(--gold); }
    .dot.bad { background: var(--hot); }
    main { padding: 24px clamp(18px, 3vw, 42px) 48px; }
    .project-band {
      display: grid;
      grid-template-columns: minmax(0, 1fr) auto;
      gap: 18px;
      align-items: end;
      margin-bottom: 18px;
      padding: 18px 20px;
      border: 1px solid rgba(250, 204, 21, .28);
      border-left: 5px solid var(--gold);
      border-radius: 8px;
      background:
        linear-gradient(135deg, rgba(127, 29, 29, .96), rgba(24, 17, 15, .96)),
        #7f1d1d;
      color: #fff7ed;
      box-shadow: 0 18px 42px rgba(63, 12, 12, .22);
    }
    .project-band span {
      display: block;
      color: #facc15;
      font-size: 12px;
      font-weight: 800;
      margin-bottom: 8px;
      text-transform: uppercase;
      letter-spacing: 0;
    }
    .project-band h2 {
      margin: 0;
      font-size: clamp(18px, 2.5vw, 30px);
      line-height: 1.35;
      letter-spacing: 0;
      max-width: 980px;
    }
    .project-band p {
      margin: 0;
      color: #f8d594;
      font-size: 13px;
      text-align: right;
      white-space: nowrap;
    }
    .summary {
      display: grid;
      grid-template-columns: minmax(260px, 1.2fr) minmax(280px, 2fr);
      gap: 18px;
      align-items: stretch;
    }
    .hero-panel, .chart-panel, .section {
      background: var(--panel);
      border: 1px solid var(--line);
      border-radius: 8px;
      box-shadow: 0 12px 30px rgba(127, 29, 29, .08);
    }
    .hero-panel {
      padding: 22px;
      display: grid;
      grid-template-rows: auto 1fr auto;
      gap: 20px;
      overflow: hidden;
      min-height: 300px;
    }
    .health-ring {
      width: min(220px, 52vw);
      aspect-ratio: 1;
      border-radius: 50%;
      margin: 4px auto;
      display: grid;
      place-items: center;
      background:
        radial-gradient(circle at center, #fff 0 57%, transparent 58%),
        conic-gradient(var(--gold) calc(var(--score) * 1%), #f1e2c7 0);
      border: 1px solid #f1d7ad;
    }
    .health-ring strong { display: block; font-size: 46px; line-height: 1; }
    .health-ring span { color: var(--muted); font-size: 12px; display: block; text-align: center; margin-top: 6px; }
    .hero-panel h2, .chart-panel h2, .section h2 {
      margin: 0;
      font-size: 14px;
      letter-spacing: 0;
      text-transform: uppercase;
      color: var(--armor);
    }
    .runtime { display: grid; grid-template-columns: 1fr 1fr; gap: 10px; }
    .runtime div { border-top: 1px solid var(--line); padding-top: 10px; min-width: 0; }
    .runtime span, .fact-grid span { display: block; color: var(--muted); font-size: 12px; }
    .runtime strong, .fact-grid strong { display: block; margin-top: 4px; overflow-wrap: anywhere; }
    .chart-panel { padding: 18px; overflow: hidden; }
    .chart-panel svg { width: 100%; height: auto; display: block; margin-top: 10px; }
    .metrics {
      display: grid;
      grid-template-columns: repeat(6, minmax(140px, 1fr));
      gap: 12px;
      margin: 18px 0;
    }
    .metric {
      background: var(--panel);
      border: 1px solid var(--line);
      border-top: 4px solid var(--accent);
      border-radius: 8px;
      padding: 14px;
      min-height: 112px;
      box-shadow: 0 8px 22px rgba(127, 29, 29, .06);
    }
    .metric-key { color: var(--muted); font-size: 12px; display: block; }
    .metric strong { display: block; font-size: clamp(24px, 3vw, 34px); margin: 8px 0 5px; letter-spacing: 0; }
    .metric small { color: var(--muted); line-height: 1.35; display: block; }
    .grid-2 { display: grid; grid-template-columns: 1fr 1fr; gap: 18px; align-items: start; }
    .section { padding: 18px; min-width: 0; }
    .flow {
      display: grid;
      grid-template-columns: repeat(5, minmax(140px, 1fr));
      gap: 12px;
      margin-top: 14px;
    }
    .flow-node {
      border: 1px solid var(--line);
      border-radius: 8px;
      padding: 14px;
      min-height: 136px;
      position: relative;
      background: linear-gradient(180deg, #fff, #fffaf0);
    }
    .flow-node i {
      width: 34px;
      height: 34px;
      display: block;
      border-radius: 8px;
      background: #fee2e2;
      margin-bottom: 12px;
      position: relative;
    }
    .flow-node i:after {
      content: "";
      position: absolute;
      inset: 9px;
      border: 2px solid var(--red);
      border-radius: 3px;
    }
    .flow-node.ok i { background: #fef3c7; }
    .flow-node.ok i:after { border-color: var(--gold); }
    .flow-node.warn i { background: #fef3c7; }
    .flow-node.warn i:after { border-color: var(--gold); }
    .flow-node.bad i { background: #fee2e2; }
    .flow-node.bad i:after { border-color: var(--red); }
    .flow-node strong { display: block; font-size: 15px; }
    .flow-node span { display: block; color: var(--muted); font-size: 12px; margin: 5px 0 12px; min-height: 32px; overflow-wrap: anywhere; }
    .flow-node b { font-size: 12px; color: var(--armor); }
    .bar-row {
      display: grid;
      grid-template-columns: minmax(96px, 150px) 1fr 38px;
      gap: 10px;
      align-items: center;
      margin: 12px 0;
      font-size: 13px;
    }
    .bar-row span { color: var(--steel); overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
    .bar-track { height: 11px; background: #f5e8d0; border-radius: 999px; overflow: hidden; }
    .bar-track i { display: block; height: 100%; border-radius: 999px; background: linear-gradient(90deg, #7f1d1d, #dc2626, #f59e0b); }
    .bar-row b { text-align: right; }
    table { width: 100%; border-collapse: collapse; margin-top: 14px; font-size: 13px; }
    th { text-align: left; color: var(--armor); font-weight: 700; border-bottom: 1px solid var(--line); padding: 10px 8px; }
    td { border-bottom: 1px solid #f3e3c8; padding: 11px 8px; vertical-align: middle; }
    td strong { display: block; }
    td span { color: var(--muted); font-size: 12px; display: block; margin-top: 2px; }
    mark {
      border-radius: 999px;
      padding: 5px 9px;
      font-size: 12px;
      color: var(--steel);
      background: #f3e8d1;
    }
    mark.ok { background: #fef3c7; color: #7c2d12; }
    mark.bad { background: #fee2e2; color: #991b1b; }
    .actions { display: flex; gap: 8px; align-items: center; flex-wrap: wrap; }
    .actions form { display: flex; gap: 8px; align-items: center; flex-wrap: wrap; }
    input[type="file"] {
      max-width: 210px;
      font-size: 12px;
      color: var(--armor);
    }
    button, .icon-link {
      appearance: none;
      border: 1px solid #e4c993;
      background: #fff;
      color: var(--armor);
      min-height: 34px;
      padding: 7px 12px;
      border-radius: 8px;
      font-weight: 650;
      font-size: 13px;
      text-decoration: none;
      cursor: pointer;
      white-space: nowrap;
    }
    button.primary {
      border-color: var(--armor);
      background: linear-gradient(135deg, var(--armor), var(--red));
      color: #fff;
    }
    button:hover, .icon-link:hover { border-color: var(--gold); box-shadow: 0 0 0 3px rgba(245, 158, 11, .16); }
    .timeline {
      list-style: none;
      padding: 0;
      margin: 14px 0 0;
      display: grid;
      gap: 10px;
    }
    .timeline li {
      border-left: 3px solid var(--gold);
      padding: 2px 0 4px 12px;
    }
    .timeline time { color: var(--muted); font-size: 12px; display: block; }
    .timeline strong { display: inline-block; margin-top: 2px; }
    .timeline span { color: var(--muted); font-size: 12px; margin-left: 8px; }
    .timeline p { margin: 5px 0 0; color: var(--steel); line-height: 1.35; }
    .fact-grid {
      display: grid;
      grid-template-columns: repeat(3, minmax(0, 1fr));
      gap: 12px;
      margin-top: 14px;
    }
    .fact-grid div {
      border: 1px solid var(--line);
      border-radius: 8px;
      padding: 12px;
      background: #fffaf0;
      min-width: 0;
    }
    .split-head {
      display: flex;
      align-items: center;
      justify-content: space-between;
      gap: 12px;
      margin-bottom: 2px;
    }
    .error {
      margin-top: 12px;
      border: 1px solid #fecaca;
      background: #fff7f7;
      color: #991b1b;
      border-radius: 8px;
      padding: 10px 12px;
      font-size: 13px;
      overflow-wrap: anywhere;
    }
    .empty { color: var(--muted); margin: 14px 0 0; }
    .footnote { color: var(--muted); font-size: 12px; margin-top: 14px; }
    @media (max-width: 1120px) {
      .metrics { grid-template-columns: repeat(3, 1fr); }
      .summary, .grid-2 { grid-template-columns: 1fr; }
      .flow { grid-template-columns: repeat(2, minmax(150px, 1fr)); }
    }
    @media (max-width: 720px) {
      .topbar { align-items: flex-start; flex-direction: column; }
      .project-band { grid-template-columns: 1fr; align-items: start; }
      .project-band p { text-align: left; white-space: normal; }
      .metrics { grid-template-columns: 1fr 1fr; }
      .runtime, .fact-grid { grid-template-columns: 1fr; }
      .flow { grid-template-columns: 1fr; }
      .bar-row { grid-template-columns: 1fr 54px; }
      .bar-row .bar-track { grid-column: 1 / -1; order: 3; }
      table, thead, tbody, tr, th, td { display: block; }
      thead { display: none; }
      tr { border: 1px solid var(--line); border-radius: 8px; margin: 12px 0; background: #fff; }
      td { border-bottom: 0; }
      .actions { align-items: stretch; }
    }
    @media (max-width: 480px) {
      .metrics { grid-template-columns: 1fr; }
      main { padding-left: 14px; padding-right: 14px; }
      .brand h1 { font-size: 16px; }
    }
    """


def _render_admin_page(data: dict, token: str) -> str:
    token_q = quote(token, safe="")
    env = data["env"]
    metrics = data["metrics"]
    db_dot = "ok" if data["db_ok"] else "bad"
    weekly_dot = "ok" if env["weekly_enabled"] else "warn"
    query_errors = data.get("query_errors") or []
    error_html = ""
    if data.get("db_error"):
        error_html += f'<div class="error">{escape(data["db_error"])}</div>'
    if query_errors:
        joined = " | ".join(_short(err, 180) for err in query_errors[:3])
        error_html += f'<div class="error">{escape(joined)}</div>'

    html = f"""<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>Hua Ro Project Command Dashboard</title>
  <style>{_page_css()}</style>
</head>
<body>
<div class="shell">
  <header class="topbar">
    <div class="brand">
      <div class="brand-mark" aria-hidden="true"></div>
      <div>
        <h1>Project Command Dashboard</h1>
        <span>LINE bot, reports, Supabase and Railway operations</span>
      </div>
    </div>
    <div class="top-actions">
      <span class="pill"><i class="dot {db_dot}"></i> Supabase</span>
      <span class="pill"><i class="dot {weekly_dot}"></i> Weekly {escape(env["weekly_format"])}</span>
      <span class="pill">Updated {escape(data["generated_at"])}</span>
    </div>
  </header>

  <main>
    <section class="project-band">
      <div>
        <span>Active Construction Project</span>
        <h2>{escape(PROJECT_DISPLAY_NAME)}</h2>
      </div>
      <p>Minimal control room · Bangkok time</p>
    </section>

    <section class="summary">
      <div class="hero-panel">
        <h2>System Readiness</h2>
        <div class="health-ring" style="--score:{int(data["health_score"])}">
          <div><strong>{int(data["health_score"])}</strong><span>health score</span></div>
        </div>
        <div class="runtime">
          <div><span>Runtime</span><strong>{escape(env["railway_name"])}</strong></div>
          <div><span>Port</span><strong>{escape(str(env["port"]))}</strong></div>
          <div><span>Weekly schedule</span><strong>{escape(env["weekly_schedule"])}</strong></div>
          <div><span>Recipients</span><strong>{escape(str(env["weekly_targets"]))}</strong></div>
        </div>
      </div>
      <div class="chart-panel">
        <div class="split-head">
          <h2>Fourteen Day Field Signal</h2>
          <span class="pill">Messages + workforce</span>
        </div>
        {_render_report_chart(data["daily_series"])}
      </div>
    </section>

    <section class="metrics" aria-label="Key metrics">
      {_render_metric("Today Messages", metrics["today_messages"], "LINE entries dated today", "#b91c1c")}
      {_render_metric("7 Day Messages", metrics["week_messages"], "Recent report traffic", "#f59e0b")}
      {_render_metric("Active Users", metrics["active_users"], "Unique LINE user ids", "#7f1d1d")}
      {_render_metric("Photos", metrics["photos_30d"], "Stored progress images", "#dc2626")}
      {_render_metric("Workers 7D", metrics["worker_total_7d"], "Sum from daily reports", "#facc15")}
      {_render_metric("Activity Types", metrics["activity_types"], "Detected work categories", "#18110f")}
    </section>

    <section class="section">
      <div class="split-head">
        <h2>Operational Topology</h2>
        <form action="/admin/trigger/weekly?token={token_q}" method="post">
          <button class="primary" type="submit">Run Weekly Report</button>
        </form>
      </div>
      <div class="flow">{_render_flow(data)}</div>
      {error_html}
    </section>

    <section class="grid-2" style="margin-top:18px">
      <div class="section">
        <h2>Activity Mix</h2>
        {_render_activity_chart(data["activity_top"])}
      </div>
      <div class="section">
        <h2>Water Level Trend</h2>
        {_render_waterline(data["water_points"])}
      </div>
    </section>

    <section class="grid-2" style="margin-top:18px">
      <div class="section">
        <h2>Latest Daily Snapshot</h2>
        {_render_latest_daily(data["latest_daily"])}
      </div>
      <div class="section">
        <h2>Recent LINE Activity</h2>
        {_render_recent(data["recent_events"])}
      </div>
    </section>

    <section class="section" style="margin-top:18px">
      <h2>Data Workbooks</h2>
      <table>
        <thead>
          <tr><th>File</th><th>Status</th><th>Size</th><th>Modified</th><th>Controls</th></tr>
        </thead>
        <tbody>{_render_files(data["files"], token_q)}</tbody>
      </table>
      <p class="footnote">The admin token stays server side. Supabase and LINE credentials are never printed on this page.</p>
    </section>
  </main>
</div>
</body>
</html>"""
    return html


@router.get("", response_class=HTMLResponse)
async def admin_home(token: str = ""):
    _check_token(token)
    return HTMLResponse(_render_admin_page(_collect_dashboard(), token))


@router.get("/api/overview")
async def admin_api_overview(token: str = ""):
    _check_token(token)
    data = _collect_dashboard()
    public = {
        "generated_at": data["generated_at"],
        "db_ok": data["db_ok"],
        "query_errors": data["query_errors"],
        "health_score": data["health_score"],
        "metrics": data["metrics"],
        "daily_series": data["daily_series"],
        "activity_top": data["activity_top"],
        "water_points": data["water_points"],
        "files": data["files"],
    }
    return JSONResponse(public)


@router.get("/download/{kind}")
async def admin_download(kind: str, token: str = ""):
    _check_token(token)
    if kind not in DATA_FILES:
        raise HTTPException(404, f"Unknown data file kind: {kind}")
    path, _ = DATA_FILES[kind]
    if not path.exists():
        raise HTTPException(404, f"Missing file: {path.name}")
    return FileResponse(
        path,
        filename=path.name,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


@router.post("/upload/{kind}")
async def admin_upload(kind: str, token: str = "", file: UploadFile = File(...)):
    _check_token(token)
    if kind not in DATA_FILES:
        raise HTTPException(404, f"Unknown data file kind: {kind}")
    if not file.filename or not file.filename.lower().endswith(".xlsx"):
        raise HTTPException(400, "Only .xlsx uploads are supported.")

    path, _ = DATA_FILES[kind]
    content = await file.read()
    if path.exists():
        backup = path.with_suffix(f".{datetime.now().strftime('%Y%m%d_%H%M%S')}.bak.xlsx")
        path.rename(backup)
    path.write_bytes(content)
    return RedirectResponse(f"/admin?token={quote(token, safe='')}", status_code=303)


@router.post("/trigger/weekly")
async def admin_trigger_weekly(token: str = ""):
    _check_token(token)
    try:
        from scheduler import run_weekly_for_current_week

        await run_weekly_for_current_week()
        return JSONResponse({"status": "ok", "message": "Weekly report triggered. Check LINE."})
    except Exception as exc:
        return JSONResponse({"status": "error", "error": str(exc)}, status_code=500)


@router.get("/check_fonts")
async def admin_check_fonts(token: str = ""):
    _check_token(token)
    result = {"thai_fonts": [], "soffice": "not_found", "errors": []}
    try:
        probe = subprocess.run(["fc-list", ":lang=th"], capture_output=True, timeout=10)
        if probe.returncode == 0:
            lines = probe.stdout.decode("utf-8", errors="ignore").splitlines()
            result["thai_fonts"] = sorted(set(line.split(":")[1].strip() for line in lines if ":" in line))
        else:
            result["errors"].append(f"fc-list failed: {probe.stderr.decode(errors='ignore')[:200]}")
    except Exception as exc:
        result["errors"].append(f"fc-list error: {exc}")

    for command in ("soffice", "libreoffice", "/usr/bin/soffice"):
        try:
            probe = subprocess.run([command, "--version"], capture_output=True, timeout=10)
            if probe.returncode == 0:
                result["soffice"] = probe.stdout.decode("utf-8", errors="ignore").strip()
                break
        except Exception:
            continue
    return JSONResponse(result)
