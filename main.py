"""
LINE Construction Report Bot - v4
เพิ่ม: parser กำลังพล (วิศวกร/หัวหน้า/ช่าง/กรรมกร) และเครื่องจักร
"""

import os, json, re, hmac, hashlib, base64, httpx, calendar
from datetime import datetime, date, timedelta, timezone
from fastapi import FastAPI, Request, HTTPException
from fastapi.responses import JSONResponse
from supabase import create_client, Client
from report_generator import generate_daily, generate_weekly, generate_monthly
from weekly_phase1 import generate_weekly_phase1
try:
    from pdf_merger import generate_weekly_phase1_pdf
    _PDF_AVAILABLE = True
except ImportError:
    _PDF_AVAILABLE = False
    generate_weekly_phase1_pdf = None

LINE_CHANNEL_SECRET       = os.environ.get("LINE_CHANNEL_SECRET", "")
LINE_CHANNEL_ACCESS_TOKEN = os.environ.get("LINE_CHANNEL_ACCESS_TOKEN", "")
SUPABASE_URL              = os.environ.get("SUPABASE_URL", "")
SUPABASE_KEY              = os.environ.get("SUPABASE_KEY", "")
PROJECT_NAME              = os.environ.get("PROJECT_NAME", "โครงการพัฒนาพื้นที่ชุมชนหัวรอและพื้นที่ต่อเนื่อง ตำบลหัวรอ อำเภอเมืองพิษณุโลก จังหวัดพิษณุโลก")

supabase: Client | None = None
if SUPABASE_URL and SUPABASE_KEY:
    supabase = create_client(SUPABASE_URL, SUPABASE_KEY)

LINE_REPLY_URL   = "https://api.line.me/v2/bot/message/reply"
LINE_PUSH_URL    = "https://api.line.me/v2/bot/message/push"
LINE_CONTENT_URL = "https://api-data.line.me/v2/bot/message/{msg_id}/content"

app = FastAPI(title="LINE Construction Report Bot", version="4.0.0")
last_text_by_user: dict[str, dict] = {}

THAI_MONTHS = {
    "มกราคม":"01","ม.ค.":"01","กุมภาพันธ์":"02","ก.พ.":"02",
    "มีนาคม":"03","มี.ค.":"03","เมษายน":"04","เม.ย.":"04",
    "พฤษภาคม":"05","พ.ค.":"05","มิถุนายน":"06","มิ.ย.":"06",
    "กรกฎาคม":"07","ก.ค.":"07","สิงหาคม":"08","ส.ค.":"08",
    "กันยายน":"09","ก.ย.":"09","ตุลาคม":"10","ต.ค.":"10",
    "พฤศจิกายน":"11","พ.ย.":"11","ธันวาคม":"12","ธ.ค.":"12",
}

ACTIVITY_KEYWORDS = {
    "ตอกเสาเข็ม":"foundation","เสาเข็ม":"foundation",
    "เทคอนกรีต":"concrete","คอนกรีต":"concrete",
    "ผูกเหล็ก":"rebar","โครงเหล็ก":"rebar",
    "งานแบบหล่อ":"formwork","ไม้แบบ":"formwork",
    "ขุดดิน":"earthwork","ถมดิน":"earthwork","บดอัด":"earthwork",
    "งานก่ออิฐ":"masonry","งานฉาบปูน":"plaster","งานทาสี":"paint",
    "งานโครงสร้าง":"structure","งานฐานราก":"foundation",
    "งานเสา":"structure","งานคาน":"structure","งานพื้น":"structure",
    "งานหลังคา":"roofing","งานผนัง":"wall","งานกำแพง":"wall",
    "งานราง":"drainage","งานท่อ":"plumbing","งานไฟฟ้า":"electrical",
    "งานประปา":"plumbing","งานระบบ":"mep",
    "ติดตั้ง":"installation","ขนส่ง":"logistics",
    "รื้อถอน":"demolition","สำรวจ":"survey","วัดระดับ":"survey",
    "ตรวจสอบ":"inspection","ทดสอบ":"testing","งานถนน":"road",
    "Shop Drawing":"shop_drawing","shop drawing":"shop_drawing",
    "Shopdrawing":"shop_drawing","shopdrawing":"shop_drawing",
    "แบบก่อสร้าง":"shop_drawing","แบบขยาย":"shop_drawing",
    "ประชุม":"meeting","รายงาน":"report",
}

# keyword → field ใน daily_reports
LABOR_FIELDS = {
    "วิศวกร":"engineers","engineer":"engineers",
    "โฟร์แมน":"foremen","หัวหน้าคนงาน":"foremen","หัวหน้า":"foremen","foreman":"foremen",
    "ช่างฝีมือ":"skilled_workers","ช่าง":"skilled_workers",
    "แรงงาน":"laborers","กรรมกร":"laborers","คนงาน":"laborers",
}

EQUIPMENT_KEYWORDS = [
    "รถแบ็คโฮ","แบ็คโฮ","รถขุด",
    "รถบรรทุก 10 ล้อ","รถบรรทุก ๑๐ ล้อ",
    "รถบรรทุก 6 ล้อ","รถบรรทุก ๖ ล้อ",
    "รถบรรทุก","รถดั้ม","รถเทรเลอร์",
    "รถเกรด","รถบด","รถบดล้อยาง","รถสั่นสะเทือน",
    "รถเครน","เครน","ปั่นจั่น",
    "รถนํ้า","รถน้ำ","รถสูบน้ำ",    # รองรับทั้ง นํ้า (nikhahit) และ น้ำ (sara am)
    "รถสกัดคอนกรีตเสาเข็ม","รถสกัดคอนกรีต",
    "รถแทร็กเตอร์","แทร็กเตอร์",
    "รถ PRIME COAT","รถ PAVE",
    "กล้องสำรวจแนว","กล้องระดับ","กล้องสำรวจ",
    "เครื่องจี้คอนกรีต","เครื่องเชื่อม",
    "เครื่องตบดิน","เครื่องสูบน้ำ",
]


def parse_thai_date(text):
    for m_th, m_num in THAI_MONTHS.items():
        pat = rf'(\d{{1,2}})\s*{re.escape(m_th)}\s*(\d{{2,4}})'
        m = re.search(pat, text)
        if m:
            day = m.group(1).zfill(2)
            yr  = int(m.group(2))
            if yr < 100:  yr += 2500
            if yr > 2400: yr -= 543
            return f"{yr}-{m_num}-{day}"
    return None


def parse_water_level(text: str):
    """แปลงค่าระดับน้ำ เช่น +97.50, -2.30, ระดับน้ำ +97.50"""
    m = re.search(r'(?:ระดับน้ำ\s*)?([+-]\d+(?:\.\d+)?)\s*(?:ม\.|เมตร|m)?', text)
    if m:
        return float(m.group(1))
    return None


def parse_labor(text: str) -> dict:
    result = {"engineers":0,"foremen":0,"skilled_workers":0,"laborers":0,"total_workers":0}
    for kw, field in LABOR_FIELDS.items():
        pat = rf'{re.escape(kw)}\s*(?:[/ๆ]\s*\w+\s*)?(\d+)\s*(?:คน|นาย)?'
        m = re.search(pat, text, re.IGNORECASE)
        if m and result[field] == 0:
            result[field] = int(m.group(1))
    total_m = re.search(r'รวม\s*(?:ทั้งหมด)?\s*(\d+)\s*คน', text)
    if total_m:
        result["total_workers"] = int(total_m.group(1))
    else:
        s = sum(result[k] for k in ["engineers","foremen","skilled_workers","laborers"])
        if s > 0: result["total_workers"] = s
    return result


def parse_equipment(text: str) -> list:
    equip, seen = [], set()
    # รองรับ: คัน / ค้น / ค่น (typo), เครื่อง, ตัว, แห่ง, ชุด
    unit_pat = r'(คัน|ค้น|ค่น|เครื่อง|ตัว|แห่ง|ชุด)'
    # normalize ชื่อเครื่องจักรในข้อความ: นํ้า → น้ำ (Unicode variant)
    text_norm = text.replace('นํ้า', 'น้ำ')
    for kw in EQUIPMENT_KEYWORDS:
        kw_norm = kw.replace('นํ้า', 'น้ำ')  # normalize keyword ด้วย
        if kw_norm in text_norm and kw_norm not in seen:
            if any(kw_norm in s for s in seen):
                seen.add(kw_norm)
                continue
            seen.add(kw_norm)
            m = re.search(rf'{re.escape(kw_norm)}\s*(\d+)\s*{unit_pat}', text_norm)
            # normalize ชื่อที่เก็บ: ใช้ น้ำ เสมอ
            name_stored = kw_norm
            qty = int(m.group(1)) if m else 1
            unit = m.group(2) if m else "คัน"
            if unit in ("ค้น", "ค่น"):  # normalize unit typo
                unit = "คัน"
            equip.append({"name": name_stored, "qty": qty, "unit": unit})
    return equip


def parse_construction_report(text: str) -> dict:
    result = {
        "work_date": parse_thai_date(text),
        "activities":[], "quantities":[], "workers":None,
        "weather":None, "raw_text":text,
        "labor": parse_labor(text),
        "equipment": parse_equipment(text),
        "water_level": parse_water_level(text),
    }
    # ลอง parse รายการเลขกำกับ (1. ... 2. ... 3. ...) ก่อน
    numbered_items = []
    for line in text.split('\n'):
        stripped = line.strip()
        m2 = re.match(r'^(\d+)[.)]\s*(.+)', stripped)
        if m2:
            numbered_items.append(m2.group(2).strip())

    if numbered_items:
        for item in numbered_items:
            act_type = "general"
            for kw, t in ACTIVITY_KEYWORDS.items():
                if kw.lower() in item.lower():
                    act_type = t
                    break
            result["activities"].append({"keyword": item, "type": act_type, "description": item})
    else:
        seen_kw, seen_desc = set(), set()
        for kw, act_type in ACTIVITY_KEYWORDS.items():
            if kw in text and kw not in seen_kw:
                seen_kw.add(kw)
                desc = kw
                for line in text.split('\n'):
                    if kw in line:
                        desc = line.strip()
                        break
                if desc not in seen_desc:
                    seen_desc.add(desc)
                    result["activities"].append({"keyword": kw, "type": act_type, "description": desc})
        if not result["activities"]:
            clean = re.sub(r'\d{1,2}\s*(?:'+
                '|'.join(re.escape(k) for k in THAI_MONTHS)+r')\s*\d{2,4}','',text).strip()
            if clean:
                result["activities"].append({"keyword":"งานทั่วไป","type":"general","description":clean[:200]})
    for kw, label in {
            "ฝนตกหนัก":    "ฝนตกหนัก",
            "ฝนตกเล็กน้อย": "ฝนตกเล็กน้อย",
            "ฝนตก":        "ฝนตกเล็กน้อย",
            "ฝน":          "ฝนตกเล็กน้อย",
            "เมฆมาก":      "เมฆมาก",
            "มืดครึ้ม":    "เมฆมาก",
            "แดดจ้า":      "แจ่มใส",
            "แดด":         "แจ่มใส",
            "แจ่มใส":      "แจ่มใส",
            "ร้อน":        "แจ่มใส",
            "หมอก":        "แจ่มใส",
    }.items():
        if kw in text: result["weather"] = label; break
    labor = result["labor"]
    result["workers"] = labor["total_workers"] if labor["total_workers"] > 0 else None
    return result


def build_image_caption(text: str) -> str:
    """ตัดบรรทัดกำลังพล เครื่องจักร และระดับน้ำออก เหลือแค่วันที่และรายการงาน"""
    lines = []
    for line in text.split('\n'):
        stripped = line.strip()
        if not stripped:
            continue
        if any(kw in stripped for kw in LABOR_FIELDS):
            continue
        # normalize นํ้า → น้ำ ก่อนเช็ค equipment keyword
        stripped_norm = stripped.replace('นํ้า', 'น้ำ')
        if any(kw.replace('นํ้า', 'น้ำ') in stripped_norm for kw in EQUIPMENT_KEYWORDS):
            continue
        # ตัดบรรทัดระดับน้ำ เช่น "+92.50", "ระดับน้ำ +92.60", "ระดับนํ้า +92.60"
        if re.match(r'^[+-]\d+(?:\.\d+)?\s*(?:ม\.|เมตร|m)?$', stripped):
            continue
        if re.search(r'ระดับน[^\s]*า', stripped) or re.search(r'ระดับน้ำ', stripped):
            continue
        lines.append(stripped)
    return '\n'.join(lines)


THAI_MONTHS_FULL = ["","มกราคม","กุมภาพันธ์","มีนาคม","เมษายน","พฤษภาคม","มิถุนายน",
                    "กรกฎาคม","สิงหาคม","กันยายน","ตุลาคม","พฤศจิกายน","ธันวาคม"]

def thai_date_str(work_date: str) -> str:
    try:
        d = date.fromisoformat(work_date)
        return f"{d.day} {THAI_MONTHS_FULL[d.month]} {d.year+543}"
    except:
        return work_date


def parse_date_arg(arg):
    arg = arg.strip()
    # รูปแบบ DD/MM เช่น 25/04
    m = re.match(r'^(\d{1,2})/(\d{1,2})$', arg)
    if m:
        t = date.today()
        return f"{t.year}-{m.group(2).zfill(2)}-{m.group(1).zfill(2)}"
    # รูปแบบวันเดียว เช่น "25" → วันที่ 25 เดือนปัจจุบัน
    m2 = re.match(r'^(\d{1,2})$', arg)
    if m2:
        t = date.today()
        return f"{t.year}-{t.month:02d}-{int(m2.group(1)):02d}"
    # รูปแบบ ISO YYYY-MM-DD
    return arg if re.match(r'^\d{4}-\d{2}-\d{2}$', arg) else None


def verify_line_signature(body, sig):
    if not LINE_CHANNEL_SECRET: return True
    h = hmac.new(LINE_CHANNEL_SECRET.encode(), body, hashlib.sha256).digest()
    return hmac.compare_digest(base64.b64encode(h).decode(), sig)


async def reply_to_line(reply_token, text):
    async with httpx.AsyncClient(timeout=10) as c:
        await c.post(LINE_REPLY_URL,
            headers={"Content-Type":"application/json","Authorization":f"Bearer {LINE_CHANNEL_ACCESS_TOKEN}"},
            json={"replyToken":reply_token,"messages":[{"type":"text","text":text}]})


_CONTENT_TYPE_MAP = {
    ".docx": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    ".pptx": "application/vnd.openxmlformats-officedocument.presentationml.presentation",
    ".xlsx": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    ".pdf":  "application/pdf",
    ".zip":  "application/zip",
}

async def push_file_to_line(user_id, filename, file_bytes):
    try:
        ts   = datetime.now().strftime("%Y%m%d_%H%M%S")
        path = f"reports/{ts}_{filename}"
        # ตรวจ content-type จากนามสกุลไฟล์ (default = docx)
        ext = "." + filename.rsplit(".", 1)[-1].lower() if "." in filename else ""
        content_type = _CONTENT_TYPE_MAP.get(ext, _CONTENT_TYPE_MAP[".docx"])
        if supabase:
            supabase.storage.from_("construction-images").upload(path, file_bytes,
                file_options={"content-type": content_type})
            url = supabase.storage.from_("construction-images").get_public_url(path)
            async with httpx.AsyncClient(timeout=10) as c:
                await c.post(LINE_PUSH_URL,
                    headers={"Content-Type":"application/json","Authorization":f"Bearer {LINE_CHANNEL_ACCESS_TOKEN}"},
                    json={"to":user_id,"messages":[{"type":"text",
                        "text":f"✅ ไฟล์รายงานพร้อมแล้ว!\n📄 {filename}\n\n🔗 กดลิงก์ดาวน์โหลด:\n{url}"}]})
    except Exception as e:
        print(f"❌ push_file: {e}")


async def fetch_line_image(message_id):
    async with httpx.AsyncClient(timeout=30) as c:
        r = await c.get(LINE_CONTENT_URL.format(msg_id=message_id),
            headers={"Authorization":f"Bearer {LINE_CHANNEL_ACCESS_TOKEN}"})
        r.raise_for_status(); return r.content


def save_raw_report(record):
    if not supabase: return None
    try:
        r = supabase.table("line_reports").insert(record).execute()
        return r.data[0].get("id") if r.data else None
    except Exception as e:
        print(f"❌ save_raw: {e}"); return None


def delete_daily_data(work_date: str) -> dict:
    """ลบข้อมูลทั้งหมดของวันที่กำหนดออกจาก 3 ตาราง"""
    if not supabase:
        return {"ok": False, "msg": "ไม่ได้เชื่อมต่อ Supabase"}
    results = {}
    # ลบ child tables ก่อน (FK constraint) แล้วค่อยลบ daily_reports
    try:
        r = supabase.table("report_activities").delete().eq("work_date", work_date).execute()
        results["report_activities"] = len(r.data or [])
    except Exception as e:
        print(f"❌ del report_activities: {e}"); results["report_activities"] = -1
    try:
        r = supabase.table("report_images").delete().eq("work_date", work_date).execute()
        results["report_images"] = len(r.data or [])
    except Exception as e:
        print(f"❌ del report_images: {e}"); results["report_images"] = -1
    try:
        r = supabase.table("line_reports").delete().eq("work_date", work_date).execute()
        results["line_reports"] = len(r.data or [])
    except Exception as e:
        print(f"❌ del line_reports: {e}"); results["line_reports"] = -1
    try:
        r = supabase.table("daily_reports").delete().eq("work_date", work_date).execute()
        results["daily_reports"] = len(r.data or [])
    except Exception as e:
        print(f"❌ del daily_reports: {e}"); results["daily_reports"] = -1
    return {"ok": True, "results": results}


def upsert_daily_report(work_date, parsed):
    if not supabase or not work_date: return False
    try:
        labor = parsed.get("labor", {})
        equip = parsed.get("equipment", [])
        ex = supabase.table("daily_reports").select("id,total_workers").eq("work_date",work_date).execute()
        water_level = parsed.get("water_level")
        if ex.data:
            upd = {"updated_at": datetime.now(timezone.utc).isoformat()}
            if parsed.get("weather"):           upd["weather_morning"]  = parsed["weather"]
            if labor.get("total_workers",0)>0:
                upd.update({"total_workers":labor["total_workers"],
                    "engineers":labor.get("engineers",0),"foremen":labor.get("foremen",0),
                    "skilled_workers":labor.get("skilled_workers",0),"laborers":labor.get("laborers",0)})
            if equip: upd["equipment"] = json.dumps(equip, ensure_ascii=False)
            if water_level is not None: upd["water_level"] = water_level
            supabase.table("daily_reports").update(upd).eq("work_date",work_date).execute()
        else:
            supabase.table("daily_reports").insert({
                "work_date":work_date,"weather_morning":parsed.get("weather"),
                "total_workers":labor.get("total_workers") or parsed.get("workers") or 0,
                "engineers":labor.get("engineers",0),"foremen":labor.get("foremen",0),
                "skilled_workers":labor.get("skilled_workers",0),"laborers":labor.get("laborers",0),
                "equipment":json.dumps(equip,ensure_ascii=False),"report_status":"draft",
                "water_level":water_level,
            }).execute()
        return True
    except Exception as e:
        print(f"❌ upsert_daily: {e}"); return False


def save_activities(work_date, source_id, activities, raw_text):
    if not supabase or not work_date or not activities: return False
    try:
        supabase.table("report_activities").insert([{
            "work_date":work_date,"source_id":source_id,
            "activity_type":a.get("type","general"),
            "description":a.get("description",raw_text[:200]),"seq_no":i+1,
        } for i,a in enumerate(activities)]).execute()
        return True
    except Exception as e:
        print(f"❌ save_act: {e}"); return False


def save_image_record(work_date, source_id, image_url, caption=""):
    if not supabase or not work_date: return False
    try:
        supabase.table("report_images").insert({
            "work_date":work_date,"source_id":source_id,"image_url":image_url,
            "caption":caption,"category":"progress","taken_at":datetime.now(timezone.utc).isoformat(),
        }).execute(); return True
    except Exception as e:
        print(f"❌ save_img: {e}"); return False


def upload_image(img_bytes, filename):
    if not supabase: return None
    try:
        supabase.storage.from_("construction-images").upload(filename, img_bytes,
            file_options={"content-type":"image/jpeg"})
        return supabase.storage.from_("construction-images").get_public_url(filename)
    except Exception as e:
        print(f"❌ upload_img: {e}"); return None


def get_daily_data(work_date):
    if not supabase: return {"work_date":work_date,"activities":[],"images":[]}
    try:
        # ดึงข้อมูลหลักจาก daily_reports โดยตรง (ไม่พึ่ง view)
        eq = supabase.table("daily_reports").select(
            "engineers,foremen,skilled_workers,laborers,equipment,water_level,total_workers,weather_morning"
        ).eq("work_date", work_date).execute()
        data = {"work_date": work_date, "activities": [], "images": []}
        if eq.data:
            data.update(eq.data[0])

        # ดึง activities จาก report_activities (deduplicate ด้วย description)
        acts = supabase.table("report_activities").select(
            "description,seq_no,activity_type"
        ).eq("work_date", work_date).order("seq_no").execute()
        if acts.data:
            seen_desc, unique_acts = set(), []
            for i, a in enumerate(acts.data):
                desc = (a.get("description") or "").strip()
                if desc and desc not in seen_desc:
                    seen_desc.add(desc)
                    unique_acts.append({"desc": desc, "seq": a.get("seq_no", i+1)})
            data["activities"] = unique_acts

        # ดึง images จาก report_images
        imgs = supabase.table("report_images").select(
            "image_url,caption"
        ).eq("work_date", work_date).execute()
        if imgs.data:
            data["images"] = [
                {"url": img.get("image_url"), "caption": img.get("caption", "")}
                for img in imgs.data
            ]

        return data
    except Exception as e:
        print(f"❌ get_daily: {e}"); return {"work_date":work_date,"activities":[],"images":[]}


def get_last_text_context(user_id: str) -> dict | None:
    """ดึง context ข้อความล่าสุดของ user จาก Supabase (fallback เมื่อ server restart)"""
    if not supabase:
        return None
    try:
        cutoff = (datetime.now(timezone.utc) - timedelta(minutes=10)).isoformat()
        r = (supabase.table("line_reports")
             .select("work_date,raw_text")
             .eq("user_id", user_id)
             .eq("message_type", "text")
             .gte("timestamp", cutoff)
             .order("timestamp", desc=True)
             .limit(1)
             .execute())
        if r.data:
            return {"work_date": r.data[0]["work_date"], "text": r.data[0]["raw_text"]}
    except Exception as e:
        print(f"❌ get_last_context: {e}")
    return None


def _enrich_days(days):
    """เติม labor+equipment แต่ละวัน"""
    if not supabase: return days
    for d in days:
        try:
            eq = supabase.table("daily_reports").select(
                "engineers,foremen,skilled_workers,laborers,equipment,water_level").eq("work_date",d["work_date"]).execute()
            if eq.data: d.update(eq.data[0])
        except: pass
    return days


def get_week_range(week_no: int, year: int, month: int):
    """คืน (start_date, end_date) ของสัปดาห์ที่ week_no (1-4) ในเดือนนั้น
    สัปดาห์ที่ 1: 1-7  |  2: 8-15  |  3: 16-23  |  4: 24-สิ้นเดือน
    """
    last_day = calendar.monthrange(year, month)[1]
    ranges = {1: (1, 7), 2: (8, 15), 3: (16, 23), 4: (24, last_day)}
    s, e = ranges.get(week_no, (1, 7))
    return date(year, month, s), date(year, month, e)


def get_current_week_no(d: date) -> int:
    if d.day <= 7:  return 1
    if d.day <= 15: return 2
    if d.day <= 23: return 3
    return 4


def parse_weekly_arg(arg):
    """แปล arg ของ /weekly → (week_no, year, month)
    รูปแบบ: (ว่าง) | "2" | "2/04" | "2/04/2569"
    """
    today = date.today()
    if not arg:
        return get_current_week_no(today), today.year, today.month
    parts = arg.split('/')
    try:
        wk = int(parts[0])
        if not 1 <= wk <= 4:
            return None
        mo = int(parts[1]) if len(parts) > 1 else today.month
        yr_raw = int(parts[2]) if len(parts) > 2 else today.year
        yr = yr_raw - 543 if yr_raw > 2400 else yr_raw
        return wk, yr, mo
    except:
        return None


def get_week_daily_list(start_date: str, end_date: str):
    if not supabase: return []
    try:
        r = supabase.table("v_daily_report_full").select("*").gte("work_date", start_date).lte("work_date", end_date).order("work_date").execute()
        return _enrich_days(r.data or [])
    except Exception as e:
        print(f"❌ get_week: {e}"); return []


def get_month_daily_list(month_str):
    if not supabase: return []
    try:
        yr,mo = int(month_str[:4]),int(month_str[5:7])
        ms = f"{yr}-{mo:02d}-01"
        me = f"{yr+1}-01-01" if mo==12 else f"{yr}-{mo+1:02d}-01"
        r = supabase.table("v_daily_report_full").select("*").gte("work_date",ms).lt("work_date",me).order("work_date").execute()
        return _enrich_days(r.data or [])
    except Exception as e:
        print(f"❌ get_month: {e}"); return []


def _help_text():
    return (
        "📖 วิธีส่งรายงานประจำวัน\n"
        "━━━━━━━━━━━━━━━\n"
        "① พิมพ์ข้อความรายงาน แล้วกดส่ง\n"
        "② ส่งรูปภาพหน้างานตามได้เลย\n"
        "   (รูปจะผูกกับรายงานอัตโนมัติ ภายใน 10 นาที)\n"
        "③ พิมพ์ /daily แล้วกดส่ง\n"
        "   → บอทจะส่งลิงก์ดาวน์โหลดไฟล์ .docx\n"
        "━━━━━━━━━━━━━━━\n"
        "📝 ตัวอย่างข้อความรายงาน:\n"
        "วันที่ 23 เมษายน 2569 อากาศแจ่มใส\n"
        "1. ผู้รับจ้างจัดทำ Shop Drawing\n"
        "2. ดำเนินการตัดหัวเสาเข็ม\n"
        "วิศวกร 2 คน หัวหน้าคนงาน 3 คน กรรมกร 2 คน รวม 7 คน\n"
        "รถแบ็คโฮ 1 คัน รถบรรทุก 2 คัน\n"
        "+92.70\n"
        "━━━━━━━━━━━━━━━\n"
        "📋 คำสั่งสร้างรายงาน:\n"
        "/daily          → รายงานวันนี้\n"
        "/daily 23/04    → รายงานวันที่ 23 เม.ย.\n"
        "/weekly         → รายงานสัปดาห์ปัจจุบัน (เก่า)\n"
        "/weekly 2       → รายงานสัปดาห์ที่ 2 เดือนนี้\n"
        "/weekly2        → รายงานสัปดาห์ฉบับเต็ม (ZIP) ⭐\n"
        "/weekly2 85     → สัปดาห์ที่ 85 (ฉบับเต็ม)\n"
        "/weekly2pdf 85  → ฉบับเต็มรวม PDF เดียว ⭐\n"
        "/monthly        → รายงานเดือนนี้\n"
        "━━━━━━━━━━━━━━━\n"
        "🌊 บันทึกระดับน้ำ (พิมพ์ตัวเลขอย่างเดียว):\n"
        "+92.70          → ระดับน้ำวันนี้ +92.70 ม.\n"
        "-2.30           → ระดับน้ำติดลบ\n"
        "━━━━━━━━━━━━━━━\n"
        "🗑️ คำสั่งลบข้อมูล:\n"
        "/delete         → ลบข้อมูลวันนี้\n"
        "/delete 22/04   → ลบข้อมูลวันที่ 22 เม.ย.\n"
        "━━━━━━━━━━━━━━━\n"
        "ℹ️ หมายเหตุ:\n"
        "• รายการงานใช้ 1. 2. 3. นำหน้าเพื่อให้บอทแยกถูก\n"
        "• ส่งรูปได้หลายรูปต่อเนื่องกัน\n"
        "• พิมพ์ /help เพื่อดูคำแนะนำนี้อีกครั้ง"
    )


async def handle_command(cmd, reply_token, user_id):
    parts = cmd.strip().split()
    command, arg = parts[0].lower(), (parts[1] if len(parts)>1 else None)

    if command == "/delete":
        td = parse_date_arg(arg) if arg else str(date.today())
        if not td:
            await reply_to_line(reply_token, "❌ รูปแบบวันที่ไม่ถูกต้อง\nตัวอย่าง: /delete 22/04"); return
        d_thai = thai_date_str(td)
        await reply_to_line(reply_token, f"⏳ กำลังลบข้อมูลวันที่ {d_thai}...")
        res = delete_daily_data(td)
        if res["ok"]:
            r = res["results"]
            await reply_to_line(reply_token, (
                f"🗑️ ลบข้อมูลวันที่ {d_thai} เรียบร้อย\n"
                f"━━━━━━━━━━━━━━━\n"
                f"• รายงานประจำวัน: {r.get('daily_reports',0)} รายการ\n"
                f"• รายการงาน: {r.get('report_activities',0)} รายการ\n"
                f"• ข้อความดิบ: {r.get('line_reports',0)} รายการ"
            ))
            # Auto-regenerate weekly report สำหรับสัปดาห์ที่มีวันที่ถูกลบ
            try:
                d_obj = date.fromisoformat(td)
                wk = get_current_week_no(d_obj)
                ws, we = get_week_range(wk, d_obj.year, d_obj.month)
                dl = get_week_daily_list(str(ws), str(we))
                if dl:
                    await reply_to_line(reply_token, f"⏳ กำลังสร้างรายงานสัปดาห์ที่ {wk} ใหม่...")
                    fb = await generate_weekly(str(ws), dl, PROJECT_NAME, week_no=wk, week_end=str(we))
                    fn = f"Weekly_Report_{d_obj.year}-{d_obj.month:02d}_W{wk}.docx"
                    await push_file_to_line(user_id, fn, fb)
                else:
                    await reply_to_line(reply_token, f"ℹ️ ไม่มีข้อมูลเหลือในสัปดาห์ที่ {wk} แล้ว")
            except Exception as e:
                print(f"❌ auto weekly: {e}")
        else:
            await reply_to_line(reply_token, f"❌ ลบไม่สำเร็จ: {res['msg']}")
        return

    await reply_to_line(reply_token, "⏳ กำลังสร้างรายงาน กรุณารอสักครู่...")
    try:
        if command == "/daily":
            td = parse_date_arg(arg) if arg else str(date.today())
            if not td:
                await reply_to_line(reply_token, "❌ รูปแบบวันที่ไม่ถูกต้อง\nตัวอย่าง: /daily 23/04")
                return
            daily_data = get_daily_data(td)
            has_data = (daily_data.get("activities") or daily_data.get("total_workers")
                        or daily_data.get("water_level") or daily_data.get("images"))
            if not has_data:
                await reply_to_line(reply_token,
                    f"⚠️ ไม่พบข้อมูลรายงานวันที่ {thai_date_str(td)}\n"
                    f"━━━━━━━━━━━━━━━\n"
                    f"กรุณาตรวจสอบ:\n"
                    f"• ส่งข้อความรายงานก่อนใช้ /daily แล้วหรือยัง?\n"
                    f"• วันที่ในข้อความถูกต้องหรือไม่?\n"
                    f"• ลองพิมพ์ /daily (ไม่ใส่วันที่) สำหรับวันนี้")
                return
            fb = await generate_daily(td, daily_data, PROJECT_NAME)
            fn = f"Daily_Report_{td}.docx"
        elif command == "/weekly":
            parsed_w = parse_weekly_arg(arg)
            if not parsed_w:
                await reply_to_line(reply_token,
                    "❌ รูปแบบไม่ถูกต้อง\nตัวอย่าง:\n/weekly        → สัปดาห์ปัจจุบัน\n/weekly 2      → สัปดาห์ที่ 2 เดือนนี้\n/weekly 2/04   → สัปดาห์ที่ 2 เดือนเม.ย.")
                return
            wk, yr, mo = parsed_w
            ws, we = get_week_range(wk, yr, mo)
            dl = get_week_daily_list(str(ws), str(we))
            if not dl: await reply_to_line(reply_token,"❌ ไม่พบข้อมูลในสัปดาห์นี้"); return
            fb = await generate_weekly(str(ws), dl, PROJECT_NAME, week_no=wk, week_end=str(we))
            fn = f"Weekly_Report_{yr}-{mo:02d}_W{wk}.docx"
        elif command == "/weekly2":
            # ใหม่: ใช้ template บริษัทจริง → ZIP รวม cover/TOC/รายละเอียด/ภาพถ่าย/รายงานประจำวัน 8 ใบ
            parsed_w = parse_weekly_arg(arg)
            if not parsed_w:
                await reply_to_line(reply_token,
                    "❌ รูปแบบไม่ถูกต้อง\nตัวอย่าง:\n/weekly2       → สัปดาห์ปัจจุบัน\n/weekly2 85    → สัปดาห์ที่ 85 เดือนนี้\n/weekly2 85/03 → สัปดาห์ที่ 85 เดือน มี.ค.")
                return
            wk, yr, mo = parsed_w
            ws, we = get_week_range(wk, yr, mo)
            dl = get_week_daily_list(str(ws), str(we))
            if not dl: await reply_to_line(reply_token,"❌ ไม่พบข้อมูลในสัปดาห์นี้"); return
            fb = await generate_weekly_phase1(week_no=wk, week_start=str(ws),
                                              daily_list=dl, project_name=PROJECT_NAME)
            fn = f"Weekly_Report_W{wk}_{yr}-{mo:02d}.zip"
        elif command == "/weekly2pdf":
            # ใหม่: generate weekly + รวมเป็น PDF เดียว (ต้องมี LibreOffice)
            if not _PDF_AVAILABLE:
                await reply_to_line(reply_token, "❌ PDF merger ไม่พร้อมใช้งาน (ต้องติดตั้ง LibreOffice + pypdf)"); return
            parsed_w = parse_weekly_arg(arg)
            if not parsed_w:
                await reply_to_line(reply_token,
                    "❌ รูปแบบไม่ถูกต้อง\nตัวอย่าง:\n/weekly2pdf       → สัปดาห์ปัจจุบัน (PDF)\n/weekly2pdf 85    → สัปดาห์ที่ 85 (PDF)")
                return
            wk, yr, mo = parsed_w
            ws, we = get_week_range(wk, yr, mo)
            dl = get_week_daily_list(str(ws), str(we))
            if not dl: await reply_to_line(reply_token,"❌ ไม่พบข้อมูลในสัปดาห์นี้"); return
            try:
                fb = await generate_weekly_phase1_pdf(week_no=wk, week_start=str(ws),
                                                      daily_list=dl, project_name=PROJECT_NAME)
                fn = f"Weekly_Report_W{wk}_{yr}-{mo:02d}.pdf"
            except Exception as e:
                await reply_to_line(reply_token, f"❌ สร้าง PDF ไม่สำเร็จ: {e}\n(ลองใช้ /weekly2 แทน)"); return
        elif command == "/monthly":
            ms = arg or date.today().strftime("%Y-%m")
            dl = get_month_daily_list(ms)
            if not dl: await reply_to_line(reply_token,"❌ ไม่พบข้อมูลในเดือนนี้"); return
            fb = await generate_monthly(ms, dl, PROJECT_NAME)
            fn = f"Monthly_Report_{ms}.docx"
        else:
            await reply_to_line(reply_token, _help_text()); return
        await push_file_to_line(user_id, fn, fb)
    except Exception as e:
        print(f"❌ handle_cmd: {e}")
        await reply_to_line(reply_token, f"❌ เกิดข้อผิดพลาด\n({str(e)[:80]})")


@app.get("/")
def root():
    return {"status":"✅ LINE Construction Report Bot v4","version":"4.0.0"}


@app.post("/webhook")
async def webhook(request: Request):
    sig  = request.headers.get("X-Line-Signature","")
    body = await request.body()
    if not verify_line_signature(body, sig):
        raise HTTPException(status_code=400, detail="Invalid signature")

    for event in json.loads(body).get("events",[]):
        if event.get("type") != "message": continue
        msg         = event.get("message",{})
        msg_type    = msg.get("type")
        reply_token = event.get("replyToken","")
        user_id     = event.get("source",{}).get("userId","unknown")
        ts_tz       = datetime.now(timezone.utc).isoformat()
        today_str   = datetime.now().strftime("%Y-%m-%d")

        if msg_type == "text":
            text = msg.get("text","").strip()
            if text.startswith("/"):
                await (reply_to_line(reply_token,_help_text()) if text.lower() in ("/help","/start")
                       else handle_command(text,reply_token,user_id))
                continue

            # ตรวจสอบว่าเป็นข้อความระดับน้ำแบบ standalone เช่น "+97.50" หรือ "-2.30"
            if re.match(r'^[+-]\d+(?:\.\d+)?$', text):
                wl = float(text)
                last = last_text_by_user.get(user_id)
                wl_date = (last["work_date"] if last and
                           (datetime.now(timezone.utc)-last["timestamp"]).total_seconds() < 600
                           else today_str)
                if supabase:
                    try:
                        supabase.table("daily_reports").update(
                            {"water_level": wl, "updated_at": datetime.now(timezone.utc).isoformat()}
                        ).eq("work_date", wl_date).execute()
                    except Exception as e:
                        print(f"❌ water_level: {e}")
                sign = "+" if wl >= 0 else ""
                await reply_to_line(reply_token,
                    f"✅ บันทึกระดับน้ำเรียบร้อย\n📅 วันที่: {wl_date}\n🌊 ระดับน้ำ: {sign}{wl:.2f} ม.")
                continue

            parsed    = parse_construction_report(text)
            work_date = parsed["work_date"] or today_str
            labor     = parsed.get("labor",{})
            equipment = parsed.get("equipment",[])

            source_id = save_raw_report({
                "timestamp":ts_tz,"user_id":user_id,"message_type":"text",
                "raw_text":text,"work_date":work_date,
                "activities":json.dumps([a["keyword"] for a in parsed["activities"]],ensure_ascii=False),
                "quantities":json.dumps(parsed["quantities"],ensure_ascii=False),
                "workers":parsed["workers"],"weather":parsed["weather"],
                "image_url":None,"image_filename":None,
            })
            upsert_daily_report(work_date, parsed)
            save_activities(work_date, source_id, parsed["activities"], text)
            last_text_by_user[user_id] = {
                "text":text,"work_date":work_date,
                "timestamp":datetime.now(timezone.utc),"source_id":source_id,
            }

            acts_str = ", ".join(a["keyword"] for a in parsed["activities"]) or "ไม่ระบุ"
            labor_lines = []
            if labor.get("engineers"):       labor_lines.append(f"วิศวกร {labor['engineers']} คน")
            if labor.get("foremen"):         labor_lines.append(f"หัวหน้า {labor['foremen']} คน")
            if labor.get("skilled_workers"): labor_lines.append(f"ช่าง {labor['skilled_workers']} คน")
            if labor.get("laborers"):        labor_lines.append(f"กรรมกร {labor['laborers']} คน")
            labor_str = f"\n👷 กำลังพล: {', '.join(labor_lines)} (รวม {labor['total_workers']} คน)" if labor_lines else ""
            equip_str = ""
            if equipment:
                eq_str = ", ".join(f"{e['name']} {e['qty']} {e['unit']}" for e in equipment[:4])
                equip_str = f"\n🚜 เครื่องจักร: {eq_str}"

            await reply_to_line(reply_token,(
                f"✅ บันทึกรายงานเรียบร้อย\n━━━━━━━━━━━━━━━\n"
                f"📅 วันที่: {work_date}\n☁️ อากาศ: {parsed['weather'] or 'ไม่ระบุ'}\n"
                f"🔨 งาน: {acts_str}{labor_str}{equip_str}\n"
                f"━━━━━━━━━━━━━━━\n📸 ส่งรูปต่อได้เลย"
            ))

        elif msg_type == "image":
            message_id = msg.get("id","")
            try:
                img_bytes   = await fetch_line_image(message_id)
                fn          = f"{datetime.now().strftime('%Y%m%d_%H%M%S')}_{message_id}.jpg"
                image_url   = upload_image(img_bytes, fn)
                caption, work_date = "", today_str

                # ลอง in-memory ก่อน (เร็วที่สุด)
                last = last_text_by_user.get(user_id)
                if last and (datetime.now(timezone.utc)-last["timestamp"]).total_seconds() < 600:
                    caption    = build_image_caption(last["text"])
                    work_date  = last["work_date"]
                else:
                    # fallback: ดึงจาก Supabase (รับมือ server restart / redeploy)
                    db_last = get_last_text_context(user_id)
                    if db_last:
                        caption   = build_image_caption(db_last["text"])
                        work_date = db_last["work_date"]
                save_raw_report({"timestamp":ts_tz,"user_id":user_id,"message_type":"image",
                    "raw_text":f"[รูปภาพ: {fn}]","work_date":work_date,
                    "activities":"[]","quantities":"[]","workers":None,"weather":None,
                    "image_url":image_url,"image_filename":fn})
                upsert_daily_report(work_date, {})
                if image_url: save_image_record(work_date, None, image_url, caption=caption)
                cap_p = f"\n📝 {caption[:50]}..." if caption else ""
                await reply_to_line(reply_token, f"✅ บันทึกรูปภาพเรียบร้อย 📸{cap_p}\n📅 {work_date}")
            except Exception as e:
                print(f"❌ img: {e}"); await reply_to_line(reply_token,"❌ บันทึกรูปไม่ได้ กรุณาส่งใหม่")

    return {"status":"ok"}


@app.get("/reports")
def get_reports(date: str = None, limit: int = 100, offset: int = 0):
    if not supabase: return JSONResponse(status_code=503,content={"error":"Supabase not configured"})
    try:
        q = supabase.table("line_reports").select("*").order("timestamp",desc=True).limit(limit).offset(offset)
        if date: q = q.eq("work_date",date)
        r = q.execute()
        return {"count":len(r.data),"offset":offset,"reports":r.data}
    except Exception as e:
        return JSONResponse(status_code=500,content={"error":str(e)})
