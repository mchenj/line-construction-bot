"""
LINE Construction Report Bot - Webhook Server
=============================================
รับข้อมูลจาก LINE (ข้อความ + รูปภาพ) และบันทึกลง Supabase
Deploy บน Railway.app (ฟรี)

โดย: ระบบรายงานก่อสร้างอัตโนมัติ
"""

import os
import json
import re
import hmac
import hashlib
import base64
import httpx
from datetime import datetime, timezone
from fastapi import FastAPI, Request, HTTPException
from fastapi.responses import JSONResponse
from supabase import create_client, Client

# ─────────────────────────────────────────
# ตั้งค่า Environment Variables
# (ใส่ค่าใน Railway Dashboard → Variables)
# ─────────────────────────────────────────
LINE_CHANNEL_SECRET       = os.environ.get("LINE_CHANNEL_SECRET", "")
LINE_CHANNEL_ACCESS_TOKEN = os.environ.get("LINE_CHANNEL_ACCESS_TOKEN", "")
SUPABASE_URL              = os.environ.get("SUPABASE_URL", "")
SUPABASE_KEY              = os.environ.get("SUPABASE_KEY", "")

# Supabase client
supabase: Client | None = None
if SUPABASE_URL and SUPABASE_KEY:
    supabase = create_client(SUPABASE_URL, SUPABASE_KEY)

# LINE API endpoints
LINE_REPLY_URL   = "https://api.line.me/v2/bot/message/reply"
LINE_CONTENT_URL = "https://api-data.line.me/v2/bot/message/{msg_id}/content"

app = FastAPI(title="LINE Construction Report Bot", version="1.0.0")

# ─────────────────────────────────────────
# ชื่อเดือนภาษาไทย → เลขเดือน
# ─────────────────────────────────────────
THAI_MONTHS = {
    "มกราคม": "01",  "ม.ค.": "01",
    "กุมภาพันธ์": "02", "ก.พ.": "02",
    "มีนาคม": "03",  "มี.ค.": "03",
    "เมษายน": "04",  "เม.ย.": "04",
    "พฤษภาคม": "05", "พ.ค.": "05",
    "มิถุนายน": "06", "มิ.ย.": "06",
    "กรกฎาคม": "07", "ก.ค.": "07",
    "สิงหาคม": "08", "ส.ค.": "08",
    "กันยายน": "09", "ก.ย.": "09",
    "ตุลาคม": "10",  "ต.ค.": "10",
    "พฤศจิกายน": "11","พ.ย.": "11",
    "ธันวาคม": "12", "ธ.ค.": "12",
}

# คำหลักกิจกรรมงานก่อสร้าง
ACTIVITY_KEYWORDS = [
    "ตอกเสาเข็ม", "เทคอนกรีต", "ผูกเหล็ก", "งานแบบหล่อ", "ไม้แบบ",
    "ขุดดิน", "ถมดิน", "บดอัด", "งานก่ออิฐ", "งานฉาบปูน", "งานทาสี",
    "งานโครงสร้าง", "งานฐานราก", "งานเสา", "งานคาน", "งานพื้น",
    "งานหลังคา", "งานผนัง", "งานกำแพง", "งานราง", "งานท่อ",
    "งานไฟฟ้า", "งานประปา", "งานระบบ", "ติดตั้ง", "ขนส่ง",
    "รื้อถอน", "สำรวจ", "วัดระดับ", "ตรวจสอบ", "ทดสอบ",
    "เสาเข็ม", "โครงเหล็ก", "งานดิน", "งานถนน",
]


# ════════════════════════════════════════
# ฟังก์ชันช่วย (Helper Functions)
# ════════════════════════════════════════

def verify_line_signature(body: bytes, signature: str) -> bool:
    """ตรวจสอบ LINE Webhook signature"""
    if not LINE_CHANNEL_SECRET:
        return True  # Skip ถ้ายังไม่ตั้งค่า (สำหรับ dev)
    hash_val = hmac.new(
        LINE_CHANNEL_SECRET.encode("utf-8"),
        body,
        hashlib.sha256
    ).digest()
    expected = base64.b64encode(hash_val).decode("utf-8")
    return hmac.compare_digest(expected, signature)


def parse_thai_date(text: str) -> str | None:
    """แปลงวันที่ภาษาไทยเป็น YYYY-MM-DD

    รองรับ:
    - "1 เมษายน 2569"  → "2026-04-01"
    - "1 เมษายน 69"    → "2026-04-01"  (เพิ่ม 2500-543)
    - "1 เม.ย. 2569"   → "2026-04-01"
    """
    for month_th, month_num in THAI_MONTHS.items():
        pattern = rf'(\d{{1,2}})\s*{re.escape(month_th)}\s*(\d{{2,4}})'
        m = re.search(pattern, text)
        if m:
            day  = m.group(1).zfill(2)
            year = int(m.group(2))
            # ปีสองหลัก เช่น 69 → พ.ศ. 2569
            if year < 100:
                year += 2500
            # พ.ศ. → ค.ศ.
            if year > 2400:
                year -= 543
            return f"{year}-{month_num}-{day}"
    return None


def parse_construction_report(text: str) -> dict:
    """แยกข้อมูลรายงานก่อสร้างจากข้อความภาษาไทย"""
    result = {
        "work_date":  parse_thai_date(text),
        "activities": [],
        "quantities": [],
        "workers":    None,
        "weather":    None,
        "raw_text":   text,
    }

    # งานที่ดำเนินการ
    for kw in ACTIVITY_KEYWORDS:
        if kw in text and kw not in result["activities"]:
            result["activities"].append(kw)

    # ปริมาณงาน: ตัวเลข + หน่วย
    qty_pattern = (
        r'(\d+(?:,\d{3})*(?:\.\d+)?)\s*'
        r'(ต้น|เมตร|ม\.|ตัน|ลบ\.ม\.|ตร\.ม\.|'
        r'แผ่น|ชุด|หลัง|จุด|คัน|ชิ้น|ราย|แห่ง|'
        r'กม\.|ลิตร|ถุง|กระสอบ|ชั้น|เสา)'
    )
    for m in re.finditer(qty_pattern, text):
        ctx_start = max(0, m.start() - 35)
        ctx_end   = min(len(text), m.end() + 15)
        result["quantities"].append({
            "amount":  float(m.group(1).replace(",", "")),
            "unit":    m.group(2),
            "context": text[ctx_start:ctx_end].strip(),
        })

    # จำนวนคนงาน
    worker_m = re.search(r'(?:คนงาน|แรงงาน|ช่าง|วิศวกร|โฟร์แมน)\s*(\d+)\s*(?:คน|นาย)', text)
    if worker_m:
        result["workers"] = int(worker_m.group(1))

    # สภาพอากาศ
    weather_map = {
        "ฝนตก": "ฝนตก", "ฝน": "มีฝน", "แดดจ้า": "แดดจ้า",
        "แดด": "แดด", "เมฆมาก": "เมฆมาก", "เมฆ": "มีเมฆ",
        "แจ่มใส": "แจ่มใส", "ร้อน": "อากาศร้อน", "หมอก": "มีหมอก",
    }
    for kw, label in weather_map.items():
        if kw in text:
            result["weather"] = label
            break

    return result


async def reply_to_line(reply_token: str, text: str):
    """ส่งข้อความตอบกลับไป LINE"""
    headers = {
        "Content-Type":  "application/json",
        "Authorization": f"Bearer {LINE_CHANNEL_ACCESS_TOKEN}",
    }
    payload = {
        "replyToken": reply_token,
        "messages":   [{"type": "text", "text": text}],
    }
    async with httpx.AsyncClient(timeout=10) as client:
        await client.post(LINE_REPLY_URL, headers=headers, json=payload)


async def fetch_line_image(message_id: str) -> bytes:
    """ดาวน์โหลดรูปภาพจาก LINE"""
    url = LINE_CONTENT_URL.format(msg_id=message_id)
    headers = {"Authorization": f"Bearer {LINE_CHANNEL_ACCESS_TOKEN}"}
    async with httpx.AsyncClient(timeout=30) as client:
        resp = await client.get(url, headers=headers)
        resp.raise_for_status()
        return resp.content


def save_report(record: dict) -> bool:
    """บันทึกข้อมูลลง Supabase"""
    if not supabase:
        print("⚠️  Supabase ยังไม่ได้ตั้งค่า — ข้อมูลไม่ถูกบันทึก")
        return False
    try:
        supabase.table("line_reports").insert(record).execute()
        return True
    except Exception as e:
        print(f"❌ Supabase error: {e}")
        return False


def upload_image(image_bytes: bytes, filename: str) -> str | None:
    """อัพโหลดรูปภาพขึ้น Supabase Storage"""
    if not supabase:
        return None
    try:
        supabase.storage.from_("construction-images").upload(
            filename, image_bytes,
            file_options={"content-type": "image/jpeg"},
        )
        url = supabase.storage.from_("construction-images").get_public_url(filename)
        return url
    except Exception as e:
        print(f"❌ Image upload error: {e}")
        return None


# ════════════════════════════════════════
# API Routes
# ════════════════════════════════════════

@app.get("/")
def root():
    """Health check"""
    return {
        "status":  "✅ LINE Construction Report Bot กำลังทำงาน",
        "version": "1.0.0",
        "routes": {
            "POST /webhook":         "LINE webhook endpoint",
            "GET  /reports":         "ดึงรายงานทั้งหมด",
            "GET  /reports?date=…":  "ดึงรายงานตามวันที่ (YYYY-MM-DD)",
            "GET  /reports?limit=…": "จำกัดจำนวนรายการ (default=100)",
        },
    }


@app.post("/webhook")
async def webhook(request: Request):
    """LINE Webhook — รับข้อความและรูปภาพ"""
    # ตรวจสอบ signature
    signature = request.headers.get("X-Line-Signature", "")
    body      = await request.body()
    if not verify_line_signature(body, signature):
        raise HTTPException(status_code=400, detail="Invalid LINE signature")

    data   = json.loads(body)
    events = data.get("events", [])

    for event in events:
        if event.get("type") != "message":
            continue

        msg          = event.get("message", {})
        msg_type     = msg.get("type")
        reply_token  = event.get("replyToken", "")
        user_id      = event.get("source", {}).get("userId", "unknown")
        timestamp_tz = datetime.now(timezone.utc).isoformat()

        # ── ข้อความ TEXT ──────────────────────────
        if msg_type == "text":
            text   = msg.get("text", "").strip()
            parsed = parse_construction_report(text)

            record = {
                "timestamp":    timestamp_tz,
                "user_id":      user_id,
                "message_type": "text",
                "raw_text":     text,
                "work_date":    parsed["work_date"],
                "activities":   json.dumps(parsed["activities"],  ensure_ascii=False),
                "quantities":   json.dumps(parsed["quantities"],  ensure_ascii=False),
                "workers":      parsed["workers"],
                "weather":      parsed["weather"],
                "image_url":    None,
                "image_filename": None,
            }
            save_report(record)

            # สร้างข้อความตอบกลับ
            date_str  = parsed["work_date"] or "ไม่ระบุ"
            acts_str  = ", ".join(parsed["activities"]) if parsed["activities"] else "ไม่ระบุ"
            qtys_str  = "\n".join(
                f"  • {q['amount']:g} {q['unit']}"
                for q in parsed["quantities"]
            ) or "  ไม่ระบุ"
            reply_txt = (
                f"✅ บันทึกรายงานเรียบร้อย\n"
                f"━━━━━━━━━━━━━━━\n"
                f"📅 วันที่: {date_str}\n"
                f"🔨 งาน: {acts_str}\n"
                f"📊 ปริมาณ:\n{qtys_str}\n"
                f"👷 คนงาน: {parsed['workers'] or 'ไม่ระบุ'} คน\n"
                f"☁️ อากาศ: {parsed['weather'] or 'ไม่ระบุ'}"
            )
            await reply_to_line(reply_token, reply_txt)

        # ── รูปภาพ IMAGE ──────────────────────────
        elif msg_type == "image":
            message_id = msg.get("id", "")
            try:
                img_bytes = await fetch_line_image(message_id)
                date_prefix = datetime.now().strftime("%Y%m%d_%H%M%S")
                filename    = f"{date_prefix}_{message_id}.jpg"
                image_url   = upload_image(img_bytes, filename)

                record = {
                    "timestamp":      timestamp_tz,
                    "user_id":        user_id,
                    "message_type":   "image",
                    "raw_text":       f"[รูปภาพ: {filename}]",
                    "work_date":      datetime.now().strftime("%Y-%m-%d"),
                    "activities":     "[]",
                    "quantities":     "[]",
                    "workers":        None,
                    "weather":        None,
                    "image_url":      image_url,
                    "image_filename": filename,
                }
                save_report(record)
                await reply_to_line(reply_token, f"✅ บันทึกรูปภาพเรียบร้อย 📸\n📁 ไฟล์: {filename}")

            except Exception as e:
                print(f"❌ Image handling error: {e}")
                await reply_to_line(reply_token, "❌ ไม่สามารถบันทึกรูปได้ กรุณาส่งใหม่อีกครั้ง")

    return {"status": "ok"}


@app.get("/reports")
def get_reports(date: str = None, limit: int = 100, offset: int = 0):
    """ดึงข้อมูลรายงาน (สำหรับ download script)"""
    if not supabase:
        return JSONResponse(status_code=503, content={"error": "Supabase ยังไม่ได้ตั้งค่า"})

    try:
        q = (
            supabase.table("line_reports")
            .select("*")
            .order("timestamp", desc=True)
            .limit(limit)
            .offset(offset)
        )
        if date:
            # กรองตามวันที่ทำงาน
            q = q.eq("work_date", date)

        result = q.execute()
        return {
            "count":   len(result.data),
            "offset":  offset,
            "reports": result.data,
        }
    except Exception as e:
        return JSONResponse(status_code=500, content={"error": str(e)})
