"""
scheduler.py
Weekly cron job — auto-generate และส่งรายงานสัปดาห์ทุกศุกร์เย็น

Config (env vars):
- WEEKLY_CRON_ENABLED=true       เปิด/ปิด cron (default: false)
- WEEKLY_CRON_DAY=fri            วันที่จะรัน (mon/tue/wed/thu/fri/sat/sun)
- WEEKLY_CRON_HOUR=17            ชั่วโมง (0-23, default 17)
- WEEKLY_CRON_MINUTE=0           นาที (0-59, default 0)
- WEEKLY_CRON_USER_IDS=Uxxx,Uyyy LINE user IDs ที่จะส่งให้ (comma-separated)
- WEEKLY_CRON_FORMAT=zip         "zip" หรือ "pdf" (default zip)
"""

import os, asyncio
from datetime import date, timedelta, datetime
from typing import Optional

try:
    from apscheduler.schedulers.asyncio import AsyncIOScheduler
    from apscheduler.triggers.cron import CronTrigger
    SCHEDULER_AVAILABLE = True
except ImportError:
    SCHEDULER_AVAILABLE = False
    AsyncIOScheduler = None
    CronTrigger = None


_scheduler: Optional["AsyncIOScheduler"] = None


def _parse_user_ids() -> list:
    raw = os.getenv("WEEKLY_CRON_USER_IDS", "").strip()
    return [uid.strip() for uid in raw.split(",") if uid.strip()]


async def run_weekly_for_current_week():
    """สร้าง weekly report ของสัปดาห์ปัจจุบัน + ส่งไปยัง user IDs ที่กำหนด"""
    user_ids = _parse_user_ids()
    if not user_ids:
        print("⚠️ scheduler: ไม่มี WEEKLY_CRON_USER_IDS — ข้าม")
        return

    # หาสัปดาห์ปัจจุบัน (อิงตาม get_current_week_no logic จาก main.py)
    today = date.today()
    if today.day <= 7:    wk = 1; ws_day, we_day = 1, 7
    elif today.day <= 15: wk = 2; ws_day, we_day = 8, 15
    elif today.day <= 23: wk = 3; ws_day, we_day = 16, 23
    else:
        wk = 4
        import calendar
        ws_day = 24
        we_day = calendar.monthrange(today.year, today.month)[1]

    ws = date(today.year, today.month, ws_day)
    we = date(today.year, today.month, we_day)

    # delayed import เพื่อหลีกเลี่ยง circular import
    from main import (get_week_daily_list, push_file_to_line, PROJECT_NAME,
                      reply_to_line)
    from weekly_phase1 import generate_weekly_phase1

    daily_list = get_week_daily_list(str(ws), str(we))
    if not daily_list:
        print(f"⚠️ scheduler: ไม่มีข้อมูลสัปดาห์ {ws}-{we}")
        for uid in user_ids:
            try:
                # ส่ง push message แบบ raw แทน (ไม่มี reply_token)
                await _push_text_to_user(uid, f"⚠️ Cron weekly: ไม่พบข้อมูลในช่วง {ws} ถึง {we}")
            except Exception as e:
                print(f"❌ push warn: {e}")
        return

    fmt = os.getenv("WEEKLY_CRON_FORMAT", "zip").lower()
    try:
        if fmt == "pdf":
            from pdf_merger import generate_weekly_phase1_pdf
            fb = await generate_weekly_phase1_pdf(week_no=wk, week_start=str(ws),
                                                  daily_list=daily_list, project_name=PROJECT_NAME)
            ext = "pdf"
        else:
            fb = await generate_weekly_phase1(week_no=wk, week_start=str(ws),
                                              daily_list=daily_list, project_name=PROJECT_NAME)
            ext = "zip"
        fn = f"Weekly_Auto_{ws.strftime('%Y%m%d')}-{we.strftime('%Y%m%d')}.{ext}"
        for uid in user_ids:
            try:
                await push_file_to_line(uid, fn, fb)
                print(f"✅ scheduler: ส่งให้ {uid[:8]}... สำเร็จ")
            except Exception as e:
                print(f"❌ scheduler push to {uid[:8]}: {e}")
    except Exception as e:
        print(f"❌ scheduler generate failed: {e}")
        for uid in user_ids:
            try:
                await _push_text_to_user(uid, f"❌ Cron weekly error: {e}")
            except Exception:
                pass


async def _push_text_to_user(user_id: str, text: str):
    """ส่ง text message ไป user (push, ไม่ใช่ reply)"""
    import httpx, os
    token = os.getenv("LINE_CHANNEL_ACCESS_TOKEN", "")
    if not token:
        return
    async with httpx.AsyncClient(timeout=10) as c:
        await c.post("https://api.line.me/v2/bot/message/push",
            headers={"Content-Type": "application/json",
                     "Authorization": f"Bearer {token}"},
            json={"to": user_id, "messages": [{"type": "text", "text": text}]})


def start_scheduler():
    """เรียกตอน FastAPI startup → สร้าง scheduler + register weekly job"""
    global _scheduler
    if not SCHEDULER_AVAILABLE:
        print("⚠️ APScheduler ไม่ได้ติดตั้ง — cron จะไม่ทำงาน")
        return None
    if os.getenv("WEEKLY_CRON_ENABLED", "false").lower() not in ("true", "1", "yes"):
        print("ℹ️ scheduler: WEEKLY_CRON_ENABLED=false — ข้าม")
        return None

    _scheduler = AsyncIOScheduler(timezone="Asia/Bangkok")

    day = os.getenv("WEEKLY_CRON_DAY", "fri").lower()
    hour = int(os.getenv("WEEKLY_CRON_HOUR", "17"))
    minute = int(os.getenv("WEEKLY_CRON_MINUTE", "0"))

    _scheduler.add_job(
        run_weekly_for_current_week,
        CronTrigger(day_of_week=day, hour=hour, minute=minute),
        id="weekly_auto",
        replace_existing=True,
    )
    _scheduler.start()
    print(f"✅ scheduler: weekly cron registered — {day} {hour:02d}:{minute:02d} (Asia/Bangkok)")
    return _scheduler


def stop_scheduler():
    global _scheduler
    if _scheduler:
        _scheduler.shutdown(wait=False)
        _scheduler = None
        print("✅ scheduler: stopped")
