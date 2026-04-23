"""
LINE Construction Report Bot - v4
เพิ่ม: parser กำลังพล (วิศวกร/หัวหน้า/ช่าง/กรรมกร) และเครื่องจักร
"""

import os, json, re, hmac, hashlib, base64, httpx
from datetime import datetime, date, timedelta, timezone
from fastapi import FastAPI, Request, HTTPException
from fastapi.responses import JSONResponse
from supabase import create_client, Client
from report_generator import generate_daily, generate_weekly, generate_monthly

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
    "รถน้ำ","รถสูบน้ำ",
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
    unit_pat = r'(คัน|เครื่อง|ตัว|แห่ง|ชุด)'
    for kw in EQUIPMENT_KEYWORDS:
        if kw in text and kw not in seen:
            seen.add(kw)
            m = re.search(rf'{re.escape(kw)}\s*(\d+)\s*{unit_pat}', text)
            equip.append({"name":kw, "qty": int(m.group(1)) if m else 1, "unit": m.group(2) if m else "คัน"})
    return equip


def parse_construction_report(text: str) -> dict:
    result = {
        "work_date": parse_thai_date(text),
        "activities":[], "quantities":[], "workers":None,
        "weather":None, "raw_text":text,
        "labor": parse_labor(text),
        "equipment": parse_equipment(text),
    }
    seen = set()
    for kw, act_type in ACTIVITY_KEYWORDS.items():
        if kw in text and kw not in seen:
            seen.add(kw)
            result["activities"].append({"keyword":kw,"type":act_type,"description":kw})
    if not result["activities"]:
        clean = re.sub(r'\d{1,2}\s*(?:'+
            '|'.join(re.escape(k) for k in THAI_MONTHS)+r')\s*\d{2,4}','',text).strip()
        if clean:
            result["activities"].append({"keyword":"งานทั่วไป","type":"general","description":clean[:200]})
    for kw, label in {"ฝนตก":"ฝนตก","ฝน":"มีฝน","แดดจ้า":"แดดจ้า","แดด":"แดด",
                      "เมฆมาก":"เมฆมาก","เมฆ":"มีเมฆ","แจ่มใส":"แจ่มใส",
                      "ร้อน":"อากาศร้อน","หมอก":"มีหมอก"}.items():
        if kw in text: result["weather"] = label; break
    labor = result["labor"]
    result["workers"] = labor["total_workers"] if labor["total_workers"] > 0 else None
    return result


def parse_date_arg(arg):
    arg = arg.strip()
    m = re.match(r'^(\d{1,2})/(\d{1,2})$', arg)
    if m:
        t = date.today()
        return f"{t.year}-{m.group(2).zfill(2)}-{m.group(1).zfill(2)}"
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


async def push_file_to_line(user_id, filename, file_bytes):
    try:
        ts   = datetime.now().strftime("%Y%m%d_%H%M%S")
        path = f"reports/{ts}_{filename}"
        if supabase:
            supabase.storage.from_("construction-images").upload(path, file_bytes,
                file_options={"content-type":"application/vnd.openxmlformats-officedocument.wordprocessingml.document"})
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


def upsert_daily_report(work_date, parsed):
    if not supabase or not work_date: return False
    try:
        labor = parsed.get("labor", {})
        equip = parsed.get("equipment", [])
        ex = supabase.table("daily_reports").select("id,total_workers").eq("work_date",work_date).execute()
        if ex.data:
            upd = {"updated_at": datetime.now(timezone.utc).isoformat()}
            if parsed.get("weather"):           upd["weather_morning"]  = parsed["weather"]
            if labor.get("total_workers",0)>0:
                upd.update({"total_workers":labor["total_workers"],
                    "engineers":labor.get("engineers",0),"foremen":labor.get("foremen",0),
                    "skilled_workers":labor.get("skilled_workers",0),"laborers":labor.get("laborers",0)})
            if equip: upd["equipment"] = json.dumps(equip, ensure_ascii=False)
            supabase.table("daily_reports").update(upd).eq("work_date",work_date).execute()
        else:
            supabase.table("daily_reports").insert({
                "work_date":work_date,"weather_morning":parsed.get("weather"),
                "total_workers":labor.get("total_workers") or parsed.get("workers") or 0,
                "engineers":labor.get("engineers",0),"foremen":labor.get("foremen",0),
                "skilled_workers":labor.get("skilled_workers",0),"laborers":labor.get("laborers",0),
                "equipment":json.dumps(equip,ensure_ascii=False),"report_status":"draft",
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
        r = supabase.table("v_daily_report_full").select("*").eq("work_date",work_date).execute()
        data = r.data[0] if r.data else {"work_date":work_date,"activities":[],"images":[]}
        eq = supabase.table("daily_reports").select(
            "engineers,foremen,skilled_workers,laborers,equipment").eq("work_date",work_date).execute()
        if eq.data: data.update(eq.data[0])
        return data
    except Exception as e:
        print(f"❌ get_daily: {e}"); return {"work_date":work_date,"activities":[],"images":[]}


def _enrich_days(days):
    """เติม labor+equipment แต่ละวัน"""
    if not supabase: return days
    for d in days:
        try:
            eq = supabase.table("daily_reports").select(
                "engineers,foremen,skilled_workers,laborers,equipment").eq("work_date",d["work_date"]).execute()
            if eq.data: d.update(eq.data[0])
        except: pass
    return days


def get_week_daily_list(week_start):
    if not supabase: return []
    try:
        ws,we = date.fromisoformat(week_start), date.fromisoformat(week_start)+timedelta(days=6)
        r = supabase.table("v_daily_report_full").select("*").gte("work_date",str(ws)).lte("work_date",str(we)).order("work_date").execute()
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
        "📋 คำสั่งสร้างรายงาน:\n"
        "━━━━━━━━━━━━━━━\n"
        "/daily          → รายงานวันนี้\n"
        "/daily 23/04    → รายงานวันที่ 23 เม.ย.\n"
        "/weekly         → รายงานสัปดาห์นี้\n"
        "/monthly        → รายงานเดือนนี้\n"
        "━━━━━━━━━━━━━━━\n"
        "📝 ตัวอย่างการบันทึก:\n"
        "วันที่ 23 เมษายน 2568 อากาศแดด\n"
        "ผู้รับจ้างจัดทำ Shop Drawing\n"
        "วิศวกร 2 คน ช่าง 5 คน กรรมกร 10 คน รวม 17 คน\n"
        "รถแบ็คโฮ 2 คัน รถบรรทุก 3 คัน\n"
        "แล้วส่งรูปตาม (ภายใน 10 นาที)"
    )


async def handle_command(cmd, reply_token, user_id):
    parts = cmd.strip().split()
    command, arg = parts[0].lower(), (parts[1] if len(parts)>1 else None)
    await reply_to_line(reply_token, "⏳ กำลังสร้างรายงาน กรุณารอสักครู่...")
    try:
        if command == "/daily":
            td = parse_date_arg(arg) if arg else str(date.today())
            fb = await generate_daily(td, get_daily_data(td), PROJECT_NAME)
            fn = f"Daily_Report_{td}.docx"
        elif command == "/weekly":
            ws = parse_date_arg(arg) if arg else str(date.today()-timedelta(days=date.today().weekday()))
            dl = get_week_daily_list(ws)
            if not dl: await reply_to_line(reply_token,"❌ ไม่พบข้อมูลในสัปดาห์นี้"); return
            fb = await generate_weekly(ws, dl, PROJECT_NAME)
            fn = f"Weekly_Report_{ws}.docx"
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
                last        = last_text_by_user.get(user_id)
                caption, work_date = "", today_str
                if last and (datetime.now(timezone.utc)-last["timestamp"]).total_seconds() < 600:
                    caption, work_date = last["text"], last["work_date"]
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
