"""
LINE Construction Report Bot - Webhook Server v3
=================================================
ครบวงจร: รับรายงานจาก LINE → เก็บ Supabase → สร้าง Word → ส่งกลับ LINE

คำสั่งใน LINE:
  /daily          → Daily Report วันนี้
  /daily 23/04    → Daily Report วันที่ระบุ
  /weekly         → Weekly Report สัปดาห์นี้
  /monthly        → Monthly Report เดือนนี้
  /help           → แสดงคำสั่งทั้งหมด

Deploy บน Railway.app
"""

import os
import json
import re
import hmac
import hashlib
import base64
import httpx
from datetime import datetime, date, timedelta, timezone
from fastapi import FastAPI, Request, HTTPException
from fastapi.responses import JSONResponse
from supabase import create_client, Client
from report_generator import generate_daily, generate_weekly, generate_monthly

# ─────────────────────────────────────────
# Environment Variables
# ─────────────────────────────────────────
LINE_CHANNEL_SECRET       = os.environ.get("LINE_CHANNEL_SECRET", "")
LINE_CHANNEL_ACCESS_TOKEN = os.environ.get("LINE_CHANNEL_ACCESS_TOKEN", "")
SUPABASE_URL              = os.environ.get("SUPABASE_URL", "")
SUPABASE_KEY              = os.environ.get("SUPABASE_KEY", "")
PROJECT_NAME              = os.environ.get("PROJECT_NAME", "โครงการก่อสร้าง")

supabase: Client | None = None
if SUPABASE_URL and SUPABASE_KEY:
    supabase = create_client(SUPABASE_URL, SUPABASE_KEY)

LINE_REPLY_URL   = "https://api.line.me/v2/bot/message/reply"
LINE_PUSH_URL    = "https://api.line.me/v2/bot/message/push"
LINE_CONTENT_URL = "https://api-data.line.me/v2/bot/message/{msg_id}/content"

app = FastAPI(title="LINE Construction Report Bot", version="3.0.0")

# in-memory: จำข้อความล่าสุดต่อ user (จับคู่ข้อความ ↔ รูปภาพ)
last_text_by_user: dict[str, dict] = {}

# ─────────────────────────────────────────
# ชื่อเดือนภาษาไทย
# ─────────────────────────────────────────
THAI_MONTHS = {
    "มกราคม": "01",    "ม.ค.": "01",
    "กุมภาพันธ์": "02","ก.พ.": "02",
    "มีนาคม": "03",    "มี.ค.": "03",
    "เมษายน": "04",    "เม.ย.": "04",
    "พฤษภาคม": "05",   "พ.ค.": "05",
    "มิถุนายน": "06",  "มิ.ย.": "06",
    "กรกฎาคม": "07",   "ก.ค.": "07",
    "สิงหาคม": "08",   "ส.ค.": "08",
    "กันยายน": "09",   "ก.ย.": "09",
    "ตุลาคม": "10",    "ต.ค.": "10",
    "พฤศจิกายน": "11", "พ.ย.": "11",
    "ธันวาคม": "12",   "ธ.ค.": "12",
}

ACTIVITY_KEYWORDS = {
    "ตอกเสาเข็ม": "foundation",  "เสาเข็ม": "foundation",
    "เทคอนกรีต": "concrete",     "คอนกรีต": "concrete",
    "ผูกเหล็ก": "rebar",          "โครงเหล็ก": "rebar",
    "งานแบบหล่อ": "formwork",    "ไม้แบบ": "formwork",
    "ขุดดิน": "earthwork",        "ถมดิน": "earthwork",
    "บดอัด": "earthwork",         "งานดิน": "earthwork",
    "งานก่ออิฐ": "masonry",      "งานฉาบปูน": "plaster",
    "งานทาสี": "paint",           "งานโครงสร้าง": "structure",
    "งานฐานราก": "foundation",   "งานเสา": "structure",
    "งานคาน": "structure",        "งานพื้น": "structure",
    "งานหลังคา": "roofing",       "งานผนัง": "wall",
    "งานกำแพง": "wall",            "งานราง": "drainage",
    "งานท่อ": "plumbing",         "งานไฟฟ้า": "electrical",
    "งานประปา": "plumbing",       "งานระบบ": "mep",
    "ติดตั้ง": "installation",    "ขนส่ง": "logistics",
    "รื้อถอน": "demolition",      "สำรวจ": "survey",
    "วัดระดับ": "survey",          "ตรวจสอบ": "inspection",
    "ทดสอบ": "testing",            "งานถนน": "road",
    "Shop Drawing": "shop_drawing","shop drawing": "shop_drawing",
    "Shopdrawing": "shop_drawing", "shopdrawing": "shop_drawing",
    "แบบก่อสร้าง": "shop_drawing","แบบขยาย": "shop_drawing",
    "ประชุม": "meeting",           "รายงาน": "report",
}


# ════════════════════════════════════════
# Helper Functions
# ════════════════════════════════════════

def verify_line_signature(body: bytes, signature: str) -> bool:
    if not LINE_CHANNEL_SECRET:
        return True
    hash_val = hmac.new(
        LINE_CHANNEL_SECRET.encode("utf-8"), body, hashlib.sha256
    ).digest()
    expected = base64.b64encode(hash_val).decode("utf-8")
    return hmac.compare_digest(expected, signature)


def parse_thai_date(text: str) -> str | None:
    for month_th, month_num in THAI_MONTHS.items():
        pattern = rf'(\d{{1,2}})\s*{re.escape(month_th)}\s*(\d{{2,4}})'
        m = re.search(pattern, text)
        if m:
            day  = m.group(1).zfill(2)
            year = int(m.group(2))
            if year < 100:
                year += 2500
            if year > 2400:
                year -= 543
            return f"{year}-{month_num}-{day}"
    return None


def parse_date_arg(arg: str) -> str | None:
    """แปลง '23/04' หรือ '2026-04-23' เป็น YYYY-MM-DD"""
    arg = arg.strip()
    m = re.match(r'^(\d{1,2})/(\d{1,2})$', arg)
    if m:
        today = date.today()
        return f"{today.year}-{m.group(2).zfill(2)}-{m.group(1).zfill(2)}"
    if re.match(r'^\d{4}-\d{2}-\d{2}$', arg):
        return arg
    return None


def parse_construction_report(text: str) -> dict:
    result = {
        "work_date":  parse_thai_date(text),
        "activities": [],
        "quantities": [],
        "workers":    None,
        "weather":    None,
        "raw_text":   text,
    }
    seen = set()
    for kw, act_type in ACTIVITY_KEYWORDS.items():
        if kw in text and kw not in seen:
            seen.add(kw)
            result["activities"].append({
                "keyword": kw, "type": act_type, "description": kw,
            })
    if not result["activities"]:
        clean = re.sub(
            r'\d{1,2}\s*(?:' + '|'.join(re.escape(k) for k in THAI_MONTHS.keys()) + r')\s*\d{2,4}',
            '', text
        ).strip()
        if clean:
            result["activities"].append({
                "keyword": "งานทั่วไป", "type": "general",
                "description": clean[:200],
            })
    qty_pattern = (
        r'(\d+(?:,\d{3})*(?:\.\d+)?)\s*'
        r'(ต้น|เมตร|ม\.|ตัน|ลบ\.ม\.|ตร\.ม\.|'
        r'แผ่น|ชุด|หลัง|จุด|คัน|ชิ้น|ราย|แห่ง|กม\.|ลิตร|ถุง|กระสอบ|ชั้น|เสา)'
    )
    for m in re.finditer(qty_pattern, text):
        ctx_start = max(0, m.start() - 35)
        ctx_end   = min(len(text), m.end() + 15)
        result["quantities"].append({
            "amount": float(m.group(1).replace(",", "")),
            "unit":   m.group(2),
            "context": text[ctx_start:ctx_end].strip(),
        })
    worker_m = re.search(
        r'(?:คนงาน|แรงงาน|ช่าง|วิศวกร|โฟร์แมน)\s*(\d+)\s*(?:คน|นาย)', text
    )
    if worker_m:
        result["workers"] = int(worker_m.group(1))
    for kw, label in {
        "ฝนตก": "ฝนตก", "ฝน": "มีฝน", "แดดจ้า": "แดดจ้า",
        "แดด": "แดด", "เมฆมาก": "เมฆมาก", "เมฆ": "มีเมฆ",
        "แจ่มใส": "แจ่มใส", "ร้อน": "อากาศร้อน", "หมอก": "มีหมอก",
    }.items():
        if kw in text:
            result["weather"] = label
            break
    return result


async def reply_to_line(reply_token: str, text: str):
    headers = {
        "Content-Type":  "application/json",
        "Authorization": f"Bearer {LINE_CHANNEL_ACCESS_TOKEN}",
    }
    async with httpx.AsyncClient(timeout=10) as client:
        await client.post(LINE_REPLY_URL, headers=headers, json={
            "replyToken": reply_token,
            "messages":   [{"type": "text", "text": text}],
        })


async def push_file_to_line(user_id: str, filename: str, file_bytes: bytes):
    """อัพโหลดไฟล์ Word ขึ้น Supabase Storage แล้วส่ง download link กลับ LINE"""
    headers = {
        "Content-Type":  "application/json",
        "Authorization": f"Bearer {LINE_CHANNEL_ACCESS_TOKEN}",
    }
    try:
        timestamp    = datetime.now().strftime("%Y%m%d_%H%M%S")
        storage_path = f"reports/{timestamp}_{filename}"
        if supabase:
            supabase.storage.from_("construction-images").upload(
                storage_path, file_bytes,
                file_options={
                    "content-type": "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                }
            )
            download_url = supabase.storage.from_("construction-images").get_public_url(storage_path)
            async with httpx.AsyncClient(timeout=10) as client:
                await client.post(LINE_PUSH_URL, headers=headers, json={
                    "to": user_id,
                    "messages": [{
                        "type": "text",
                        "text": (
                            f"✅ ไฟล์รายงานพร้อมแล้ว!\n"
                            f"📄 {filename}\n\n"
                            f"🔗 กดลิงก์เพื่อดาวน์โหลด:\n{download_url}"
                        )
                    }]
                })
    except Exception as e:
        print(f"❌ push_file_to_line: {e}")


async def fetch_line_image(message_id: str) -> bytes:
    url = LINE_CONTENT_URL.format(msg_id=message_id)
    headers = {"Authorization": f"Bearer {LINE_CHANNEL_ACCESS_TOKEN}"}
    async with httpx.AsyncClient(timeout=30) as client:
        resp = await client.get(url, headers=headers)
        resp.raise_for_status()
        return resp.content


# ════════════════════════════════════════
# Database Writers
# ════════════════════════════════════════

def save_raw_report(record: dict) -> int | None:
    if not supabase:
        return None
    try:
        result = supabase.table("line_reports").insert(record).execute()
        return result.data[0].get("id") if result.data else None
    except Exception as e:
        print(f"❌ save_raw_report: {e}")
        return None


def upsert_daily_report(work_date: str, parsed: dict) -> bool:
    if not supabase or not work_date:
        return False
    try:
        existing = (
            supabase.table("daily_reports")
            .select("id, total_workers")
            .eq("work_date", work_date)
            .execute()
        )
        if existing.data:
            upd: dict = {"updated_at": datetime.now(timezone.utc).isoformat()}
            if parsed.get("weather"):
                upd["weather_morning"] = parsed["weather"]
            if parsed.get("workers") and not existing.data[0].get("total_workers"):
                upd["total_workers"] = parsed["workers"]
            supabase.table("daily_reports").update(upd).eq("work_date", work_date).execute()
        else:
            supabase.table("daily_reports").insert({
                "work_date":       work_date,
                "weather_morning": parsed.get("weather"),
                "total_workers":   parsed.get("workers") or 0,
                "report_status":   "draft",
            }).execute()
        return True
    except Exception as e:
        print(f"❌ upsert_daily_report: {e}")
        return False


def save_activities(work_date: str, source_id: int | None, activities: list, raw_text: str) -> bool:
    if not supabase or not work_date or not activities:
        return False
    try:
        supabase.table("report_activities").insert([{
            "work_date":     work_date,
            "source_id":     source_id,
            "activity_type": a.get("type", "general"),
            "description":   a.get("description", raw_text[:200]),
            "seq_no":        i + 1,
        } for i, a in enumerate(activities)]).execute()
        return True
    except Exception as e:
        print(f"❌ save_activities: {e}")
        return False


def save_image_record(work_date: str, source_id: int | None, image_url: str, caption: str = "") -> bool:
    if not supabase or not work_date:
        return False
    try:
        supabase.table("report_images").insert({
            "work_date": work_date,
            "source_id": source_id,
            "image_url": image_url,
            "caption":   caption,
            "category":  "progress",
            "taken_at":  datetime.now(timezone.utc).isoformat(),
        }).execute()
        return True
    except Exception as e:
        print(f"❌ save_image_record: {e}")
        return False


def upload_image(image_bytes: bytes, filename: str) -> str | None:
    if not supabase:
        return None
    try:
        supabase.storage.from_("construction-images").upload(
            filename, image_bytes,
            file_options={"content-type": "image/jpeg"},
        )
        return supabase.storage.from_("construction-images").get_public_url(filename)
    except Exception as e:
        print(f"❌ upload_image: {e}")
        return None


# ════════════════════════════════════════
# Report Query Helpers
# ════════════════════════════════════════

def get_daily_data(work_date: str) -> dict:
    if not supabase:
        return {"work_date": work_date, "activities": [], "images": []}
    try:
        result = (
            supabase.table("v_daily_report_full")
            .select("*")
            .eq("work_date", work_date)
            .execute()
        )
        return result.data[0] if result.data else {"work_date": work_date, "activities": [], "images": []}
    except Exception as e:
        print(f"❌ get_daily_data: {e}")
        return {"work_date": work_date, "activities": [], "images": []}


def get_week_daily_list(week_start: str) -> list[dict]:
    if not supabase:
        return []
    try:
        ws = date.fromisoformat(week_start)
        we = ws + timedelta(days=6)
        result = (
            supabase.table("v_daily_report_full")
            .select("*")
            .gte("work_date", str(ws))
            .lte("work_date", str(we))
            .order("work_date")
            .execute()
        )
        return result.data or []
    except Exception as e:
        print(f"❌ get_week_daily_list: {e}")
        return []


def get_month_daily_list(month_str: str) -> list[dict]:
    if not supabase:
        return []
    try:
        year, month = int(month_str[:4]), int(month_str[5:7])
        month_start = f"{year}-{month:02d}-01"
        month_end   = f"{year+1}-01-01" if month == 12 else f"{year}-{month+1:02d}-01"
        result = (
            supabase.table("v_daily_report_full")
            .select("*")
            .gte("work_date", month_start)
            .lt("work_date", month_end)
            .order("work_date")
            .execute()
        )
        return result.data or []
    except Exception as e:
        print(f"❌ get_month_daily_list: {e}")
        return []


# ════════════════════════════════════════
# Command Handler
# ════════════════════════════════════════

def _help_text() -> str:
    return (
        "📋 คำสั่งสร้างรายงาน:\n"
        "━━━━━━━━━━━━━━━\n"
        "/daily          → รายงานประจำวัน (วันนี้)\n"
        "/daily 23/04    → รายงานวันที่ 23 เม.ย.\n"
        "/weekly         → รายงานประจำสัปดาห์\n"
        "/monthly        → รายงานประจำเดือน\n"
        "━━━━━━━━━━━━━━━\n"
        "📸 วิธีบันทึกรายงาน:\n"
        "1. ส่งข้อความก่อน เช่น\n"
        "   'วันที่ 23 เมษายน 2568 ผู้รับจ้าง\n"
        "    จัดทำแบบขยาย Shop Drawing'\n"
        "2. ส่งรูปตามทันที (ภายใน 10 นาที)\n"
        "   → รูปจะผูกกับข้อความอัตโนมัติ"
    )


async def handle_command(cmd: str, reply_token: str, user_id: str):
    parts   = cmd.strip().split()
    command = parts[0].lower()
    arg     = parts[1] if len(parts) > 1 else None

    await reply_to_line(reply_token, "⏳ กำลังสร้างรายงาน กรุณารอสักครู่...")

    try:
        if command == "/daily":
            target_date = parse_date_arg(arg) if arg else str(date.today())
            data        = get_daily_data(target_date)
            file_bytes  = await generate_daily(target_date, data, PROJECT_NAME)
            filename    = f"Daily_Report_{target_date}.docx"

        elif command == "/weekly":
            if arg:
                week_start = parse_date_arg(arg) or str(date.today())
            else:
                today      = date.today()
                week_start = str(today - timedelta(days=today.weekday()))
            daily_list = get_week_daily_list(week_start)
            if not daily_list:
                await reply_to_line(reply_token, "❌ ไม่พบข้อมูลในสัปดาห์นี้ครับ\nลองส่งรายงานในไลน์ก่อน แล้วค่อยขอรายงานอีกครั้ง")
                return
            file_bytes = await generate_weekly(week_start, daily_list, PROJECT_NAME)
            filename   = f"Weekly_Report_{week_start}.docx"

        elif command == "/monthly":
            month_str  = arg or date.today().strftime("%Y-%m")
            daily_list = get_month_daily_list(month_str)
            if not daily_list:
                await reply_to_line(reply_token, "❌ ไม่พบข้อมูลในเดือนนี้ครับ")
                return
            file_bytes = await generate_monthly(month_str, daily_list, PROJECT_NAME)
            filename   = f"Monthly_Report_{month_str}.docx"

        else:
            await reply_to_line(reply_token, _help_text())
            return

        await push_file_to_line(user_id, filename, file_bytes)

    except Exception as e:
        print(f"❌ handle_command: {e}")
        await reply_to_line(reply_token, f"❌ เกิดข้อผิดพลาด กรุณาลองใหม่\n({str(e)[:80]})")


# ════════════════════════════════════════
# API Routes
# ════════════════════════════════════════

@app.get("/")
def root():
    return {
        "status":   "✅ LINE Construction Report Bot v3 กำลังทำงาน",
        "version":  "3.0.0",
        "commands": ["/daily", "/weekly", "/monthly", "/help"],
    }


@app.post("/webhook")
async def webhook(request: Request):
    signature = request.headers.get("X-Line-Signature", "")
    body      = await request.body()
    if not verify_line_signature(body, signature):
        raise HTTPException(status_code=400, detail="Invalid LINE signature")

    events = json.loads(body).get("events", [])

    for event in events:
        if event.get("type") != "message":
            continue

        msg          = event.get("message", {})
        msg_type     = msg.get("type")
        reply_token  = event.get("replyToken", "")
        user_id      = event.get("source", {}).get("userId", "unknown")
        timestamp_tz = datetime.now(timezone.utc).isoformat()
        today_str    = datetime.now().strftime("%Y-%m-%d")

        # ── TEXT ──────────────────────────────────
        if msg_type == "text":
            text = msg.get("text", "").strip()

            # คำสั่ง
            if text.startswith("/"):
                if text.lower() in ("/help", "/start"):
                    await reply_to_line(reply_token, _help_text())
                else:
                    await handle_command(text, reply_token, user_id)
                continue

            # รายงานปกติ
            parsed    = parse_construction_report(text)
            work_date = parsed["work_date"] or today_str

            source_id = save_raw_report({
                "timestamp":      timestamp_tz,
                "user_id":        user_id,
                "message_type":   "text",
                "raw_text":       text,
                "work_date":      work_date,
                "activities":     json.dumps(
                    [a["keyword"] for a in parsed["activities"]], ensure_ascii=False
                ),
                "quantities":     json.dumps(parsed["quantities"], ensure_ascii=False),
                "workers":        parsed["workers"],
                "weather":        parsed["weather"],
                "image_url":      None,
                "image_filename": None,
            })
            upsert_daily_report(work_date, parsed)
            save_activities(work_date, source_id, parsed["activities"], text)

            # จำไว้จับคู่รูป
            last_text_by_user[user_id] = {
                "text":      text,
                "work_date": work_date,
                "timestamp": datetime.now(timezone.utc),
                "source_id": source_id,
            }

            acts_str = ", ".join(a["keyword"] for a in parsed["activities"]) or "ไม่ระบุ"
            qtys_str = "\n".join(
                f"  • {q['amount']:g} {q['unit']}" for q in parsed["quantities"]
            ) or "  ไม่ระบุ"
            await reply_to_line(reply_token, (
                f"✅ บันทึกรายงานเรียบร้อย\n"
                f"━━━━━━━━━━━━━━━\n"
                f"📅 วันที่: {work_date}\n"
                f"🔨 งาน: {acts_str}\n"
                f"📊 ปริมาณ:\n{qtys_str}\n"
                f"👷 คนงาน: {parsed['workers'] or 'ไม่ระบุ'} คน\n"
                f"☁️ อากาศ: {parsed['weather'] or 'ไม่ระบุ'}\n"
                f"━━━━━━━━━━━━━━━\n"
                f"📸 ส่งรูปต่อได้เลย จะผูกกับรายงานนี้อัตโนมัติ"
            ))

        # ── IMAGE ─────────────────────────────────
        elif msg_type == "image":
            message_id = msg.get("id", "")
            try:
                img_bytes   = await fetch_line_image(message_id)
                date_prefix = datetime.now().strftime("%Y%m%d_%H%M%S")
                filename    = f"{date_prefix}_{message_id}.jpg"
                image_url   = upload_image(img_bytes, filename)

                # จับคู่กับข้อความล่าสุด (ภายใน 10 นาที)
                last      = last_text_by_user.get(user_id)
                caption   = ""
                work_date = today_str
                if last:
                    diff = (datetime.now(timezone.utc) - last["timestamp"]).total_seconds()
                    if diff < 600:
                        caption   = last["text"]
                        work_date = last["work_date"]

                save_raw_report({
                    "timestamp":      timestamp_tz,
                    "user_id":        user_id,
                    "message_type":   "image",
                    "raw_text":       f"[รูปภาพ: {filename}]",
                    "work_date":      work_date,
                    "activities":     "[]",
                    "quantities":     "[]",
                    "workers":        None,
                    "weather":        None,
                    "image_url":      image_url,
                    "image_filename": filename,
                })
                upsert_daily_report(work_date, {})
                if image_url:
                    save_image_record(work_date, None, image_url, caption=caption)

                cap_preview = f"\n📝 {caption[:50]}..." if caption else ""
                await reply_to_line(reply_token, (
                    f"✅ บันทึกรูปภาพเรียบร้อย 📸{cap_preview}\n"
                    f"📅 วันที่: {work_date}"
                ))

            except Exception as e:
                print(f"❌ Image error: {e}")
                await reply_to_line(reply_token, "❌ บันทึกรูปไม่ได้ กรุณาส่งใหม่")

    return {"status": "ok"}


@app.get("/reports")
def get_reports(date: str = None, limit: int = 100, offset: int = 0):
    if not supabase:
        return JSONResponse(status_code=503, content={"error": "Supabase not configured"})
    try:
        q = (
            supabase.table("line_reports")
            .select("*")
            .order("timestamp", desc=True)
            .limit(limit)
            .offset(offset)
        )
        if date:
            q = q.eq("work_date", date)
        result = q.execute()
        return {"count": len(result.data), "offset": offset, "reports": result.data}
    except Exception as e:
        return JSONResponse(status_code=500, content={"error": str(e)})
