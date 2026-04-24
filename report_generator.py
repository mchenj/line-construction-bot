"""
report_generator.py v2
เพิ่ม: ตารางกำลังพล (วิศวกร/หัวหน้า/ช่าง/กรรมกร) และตารางเครื่องจักร
"""

import io, json, re, httpx
from datetime import datetime, date, timedelta
from docx import Document
from docx.shared import Inches, Pt, RGBColor, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

THAI_MONTHS_FULL = ["","มกราคม","กุมภาพันธ์","มีนาคม","เมษายน","พฤษภาคม","มิถุนายน",
                    "กรกฎาคม","สิงหาคม","กันยายน","ตุลาคม","พฤศจิกายน","ธันวาคม"]

_MONTHS_PAT = "|".join(THAI_MONTHS_FULL[1:])

def clean_caption(text: str) -> str:
    """ตัดบรรทัดวันที่ กำลังพล และเครื่องจักรออกจาก caption ใต้ภาพ"""
    if not text:
        return ""
    lines = []
    for line in text.split('\n'):
        s = line.strip()
        if not s:
            continue
        # ตัดบรรทัดวันที่ เช่น "วันที่ 22 เมษายน 2569 อากาศแจ่มใส"
        if re.search(rf'(?:วันที่).*(?:{_MONTHS_PAT}).*\d{{4}}', s):
            continue
        # ตัดบรรทัดกำลังพล เช่น "วิศวกร 2 คน หัวหน้าคนงาน 3 คน รวม 7 คน"
        if re.search(r'(?:วิศวกร|หัวหน้าคนงาน|หัวหน้า|กรรมกร|ช่างฝีมือ|ช่าง|คนงาน)\s*\d+\s*คน', s):
            continue
        # ตัดบรรทัดเครื่องจักร เช่น "รถแบ็คโฮ 1 คัน"
        if re.search(r'(?:รถแบ็คโฮ|แบ็คโฮ|รถขุด|รถบรรทุก|รถเครน|รถบด|รถน้ำ|รถเกรด|รถสูบน้ำ|รถแทร็กเตอร์)\s*\d+\s*คัน', s):
            continue
        lines.append(s)
    result = "\n".join(lines)
    # แก้ "วันที่ วันที่" ซ้ำ กรณีที่หลุดผ่านมา
    result = re.sub(r'วันที่\s+วันที่', 'วันที่', result)
    return result

def thai_date(d):
    if isinstance(d, str): d = date.fromisoformat(d)
    return f"{d.day} {THAI_MONTHS_FULL[d.month]} {d.year+543}"

def thai_date_short(d):
    if isinstance(d, str): d = date.fromisoformat(d)
    return f"{d.day:02d}/{d.month:02d}/{str(d.year+543)[2:]}"

def set_cell_bg(cell, hex_color):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:fill"), hex_color)
    shd.set(qn("w:val"), "clear")
    tcPr.append(shd)

def add_header_row(table, headers, bg="1F4E79"):
    row = table.rows[0]
    for i, h in enumerate(headers):
        cell = row.cells[i]
        cell.text = ""
        set_cell_bg(cell, bg)
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p.add_run(h)
        run.bold = True
        run.font.color.rgb = RGBColor(0xFF,0xFF,0xFF)
        run.font.size = Pt(11)

def style_doc(doc):
    s = doc.styles["Normal"]
    s.font.name = "TH Sarabun New"
    s.font.size = Pt(14)
    for sec in doc.sections:
        sec.page_height = Cm(29.7); sec.page_width = Cm(21.0)
        sec.left_margin = Cm(2.5);  sec.right_margin = Cm(2.0)
        sec.top_margin  = Cm(2.0);  sec.bottom_margin = Cm(2.0)

def add_title_block(doc, title, subtitle, project_name):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run(title); r.bold = True; r.font.size = Pt(18)
    r.font.color.rgb = RGBColor(0x1F,0x4E,0x79)
    p2 = doc.add_paragraph()
    p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p2.add_run(subtitle).font.size = Pt(13)
    if project_name:
        p3 = doc.add_paragraph()
        p3.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r3 = p3.add_run(f"โครงการ: {project_name}")
        r3.bold = True; r3.font.size = Pt(13)
    p4 = doc.add_paragraph()
    p4.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r4 = p4.add_run("─"*60); r4.font.size = Pt(10)
    r4.font.color.rgb = RGBColor(0x1F,0x4E,0x79)
    doc.add_paragraph()

async def download_image_bytes(url):
    try:
        async with httpx.AsyncClient(timeout=15) as c:
            r = await c.get(url); r.raise_for_status(); return r.content
    except Exception as e:
        print(f"⚠️ img download: {e}"); return None


def add_labor_table(doc, data: dict):
    """ตารางกำลังพล: วิศวกร / หัวหน้า / ช่าง / กรรมกร / รวม"""
    engineers       = data.get("engineers") or 0
    foremen         = data.get("foremen") or 0
    skilled_workers = data.get("skilled_workers") or 0
    laborers        = data.get("laborers") or 0
    total           = data.get("total_workers") or (engineers+foremen+skilled_workers+laborers)

    if total == 0 and not any([engineers, foremen, skilled_workers, laborers]):
        p = doc.add_paragraph("— ไม่มีข้อมูลกำลังพล —")
        p.runs[0].font.size = Pt(12)
        return

    tbl = doc.add_table(rows=2, cols=5)
    tbl.style = "Table Grid"
    add_header_row(tbl, ["วิศวกร/ช่าง (คน)", "หัวหน้าคนงาน (คน)", "ช่างฝีมือ (คน)", "กรรมกร (คน)", "รวม (คน)"])
    row = tbl.rows[1]
    for i, v in enumerate([engineers, foremen, skilled_workers, laborers, total]):
        row.cells[i].text = str(v) if v else "—"
        row.cells[i].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        if v and i == 4:  # รวม — สีเน้น
            set_cell_bg(row.cells[i], "D6E4F0")
            if row.cells[i].paragraphs[0].runs:
                row.cells[i].paragraphs[0].runs[0].bold = True


def add_equipment_table(doc, data: dict):
    """ตารางเครื่องจักร"""
    equip_raw = data.get("equipment")
    if isinstance(equip_raw, str):
        try: equipment = json.loads(equip_raw)
        except: equipment = []
    else:
        equipment = equip_raw or []

    if not equipment:
        doc.add_paragraph("— ไม่มีข้อมูลเครื่องจักร —").runs[0].font.size = Pt(12)
        return

    tbl = doc.add_table(rows=len(equipment)+1, cols=3)
    tbl.style = "Table Grid"
    add_header_row(tbl, ["ลำดับ", "ประเภทเครื่องจักร/ยานพาหนะ", "จำนวน"])
    for i, eq in enumerate(equipment):
        row = tbl.rows[i+1]
        row.cells[0].text = str(i+1)
        row.cells[1].text = eq.get("name", "—")
        row.cells[2].text = f"{eq.get('qty',1)} {eq.get('unit','คัน')}"
        row.cells[0].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        row.cells[2].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        if i % 2 == 0:
            for cell in row.cells: set_cell_bg(cell, "EBF3FB")


def add_signature_table(doc, left="ผู้รายงาน (Reported by)", right="ผู้ตรวจสอบ (Checked by)"):
    tbl = doc.add_table(rows=3, cols=2)
    tbl.style = "Table Grid"
    tbl.rows[0].cells[0].text = left
    tbl.rows[0].cells[1].text = right
    tbl.rows[1].cells[0].text = "\n\n"
    tbl.rows[1].cells[1].text = "\n\n"
    tbl.rows[2].cells[0].text = "วันที่: ____________________"
    tbl.rows[2].cells[1].text = "วันที่: ____________________"
    for cell in tbl.rows[0].cells:
        set_cell_bg(cell, "D6E4F0")
        if cell.paragraphs[0].runs:
            cell.paragraphs[0].runs[0].bold = True


# ════════════════════════════════════════
# DAILY REPORT
# ════════════════════════════════════════

async def generate_daily(work_date: str, daily_data: dict, project_name: str = "โครงการก่อสร้าง") -> bytes:
    doc = Document()
    style_doc(doc)
    d = date.fromisoformat(work_date)
    add_title_block(doc, "รายงานประจำวัน (DAILY REPORT)",
        f"วันที่ {thai_date(d)}  |  ฉบับที่ DR-{d.strftime('%Y%m%d')}", project_name)

    # 1. ข้อมูลทั่วไป
    doc.add_heading("1. ข้อมูลทั่วไป", level=2)
    info = doc.add_table(rows=2, cols=4)
    info.style = "Table Grid"
    add_header_row(info, ["วันที่ทำงาน","สภาพอากาศ","จำนวนคนงานรวม","ผู้ควบคุมงาน"])
    for i,v in enumerate([thai_date(d),
                           daily_data.get("weather_morning") or "—",
                           str(daily_data.get("total_workers") or "—")+" คน",
                           daily_data.get("supervisor") or "—"]):
        info.rows[1].cells[i].text = v
        info.rows[1].cells[i].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
    doc.add_paragraph()

    # 2. กำลังพล
    doc.add_heading("2. กำลังพล", level=2)
    add_labor_table(doc, daily_data)
    doc.add_paragraph()

    # 3. เครื่องจักร
    doc.add_heading("3. เครื่องจักรและยานพาหนะ", level=2)
    add_equipment_table(doc, daily_data)
    doc.add_paragraph()

    # 4. งานที่ดำเนินการ
    doc.add_heading("4. งานที่ดำเนินการ", level=2)
    activities = daily_data.get("activities") or []
    if activities:
        act_tbl = doc.add_table(rows=len(activities)+1, cols=4)
        act_tbl.style = "Table Grid"
        add_header_row(act_tbl, ["ลำดับ","รายการงาน","สถานที่","หมายเหตุ"])
        for i, act in enumerate(activities):
            row = act_tbl.rows[i+1]
            row.cells[0].text = str(act.get("seq") or i+1)
            row.cells[1].text = act.get("desc") or act.get("description") or "—"
            row.cells[2].text = act.get("location") or "—"
            row.cells[3].text = f"{act['qty']:g} {act['unit']}" if act.get("qty") and act.get("unit") else "—"
            row.cells[0].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
            row.cells[3].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
            if i%2==0:
                for cell in row.cells: set_cell_bg(cell,"EBF3FB")
    else:
        doc.add_paragraph("— ไม่มีข้อมูลกิจกรรม —")
    doc.add_paragraph()

    # 5. รูปภาพ
    images = daily_data.get("images") or []
    if images:
        doc.add_heading("5. รูปภาพประกอบ", level=2)
        for img_info in images:
            url     = img_info.get("url") or img_info.get("image_url")
            acts_text = clean_caption(img_info.get("caption") or "")
            caption = f"วันที่ {thai_date(d)}"
            if acts_text:
                caption += f"\n\n{acts_text}"
            if not url: continue
            img_bytes = await download_image_bytes(url)
            if img_bytes:
                p = doc.add_paragraph()
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                p.add_run().add_picture(io.BytesIO(img_bytes), width=Inches(5.5))
            cap_p = doc.add_paragraph(caption)
            cap_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            if cap_p.runs:
                cap_p.runs[0].italic = True; cap_p.runs[0].font.size = Pt(12)
            doc.add_paragraph()

    # ลายเซ็น
    doc.add_heading("ลงชื่อ / Signature", level=2)
    add_signature_table(doc)

    buf = io.BytesIO(); doc.save(buf); return buf.getvalue()


# ════════════════════════════════════════
# WEEKLY REPORT
# ════════════════════════════════════════

async def generate_weekly(week_start: str, daily_list: list, project_name: str = "โครงการก่อสร้าง",
                          week_no: int = None, week_end: str = None) -> bytes:
    doc = Document()
    style_doc(doc)
    ws = date.fromisoformat(week_start)
    we = date.fromisoformat(week_end) if week_end else ws + timedelta(days=6)
    if week_no:
        subtitle = (f"สัปดาห์ที่ {week_no}  |  {thai_date(ws)} — {thai_date(we)}"
                    f"  |  ฉบับที่ WR-{ws.strftime('%Y%m')}-W{week_no}")
    else:
        subtitle = f"{thai_date(ws)} — {thai_date(we)}  |  ฉบับที่ WR-{ws.strftime('%Y%m%d')}"
    add_title_block(doc, "รายงานความก้าวหน้าประจำสัปดาห์ (WEEKLY PROGRESS REPORT)",
        subtitle, project_name)

    # สรุปสัปดาห์
    doc.add_heading("1. สรุปภาพรวมสัปดาห์", level=2)
    total_w = sum(d.get("total_workers") or 0 for d in daily_list)
    total_e = sum(d.get("engineers") or 0 for d in daily_list)
    total_f = sum(d.get("foremen") or 0 for d in daily_list)
    total_s = sum(d.get("skilled_workers") or 0 for d in daily_list)
    total_l = sum(d.get("laborers") or 0 for d in daily_list)

    sum_tbl = doc.add_table(rows=2, cols=5)
    sum_tbl.style = "Table Grid"
    add_header_row(sum_tbl, ["วันทำงาน","วิศวกร (คน-วัน)","หัวหน้า (คน-วัน)","ช่าง (คน-วัน)","กรรมกร (คน-วัน)"])
    r = sum_tbl.rows[1]
    for i,v in enumerate([f"{len(daily_list)} วัน", str(total_e), str(total_f), str(total_s), str(total_l)]):
        r.cells[i].text = v
        r.cells[i].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
    doc.add_paragraph()

    # ตารางสรุปรายวัน
    doc.add_heading("2. ตารางสรุปงานรายวัน", level=2)
    day_tbl = doc.add_table(rows=len(daily_list)+1, cols=5)
    day_tbl.style = "Table Grid"
    add_header_row(day_tbl, ["วันที่","อากาศ","คนงานรวม (คน)","กิจกรรมหลัก","เครื่องจักร"])
    for i, d in enumerate(daily_list):
        acts = d.get("activities") or []
        act_str = ", ".join(a.get("desc") or a.get("description","") for a in acts[:2])
        if len(acts)>2: act_str += f" (+{len(acts)-2})"

        equip_raw = d.get("equipment")
        if isinstance(equip_raw, str):
            try: equip = json.loads(equip_raw)
            except: equip = []
        else: equip = equip_raw or []
        eq_str = ", ".join(f"{e['name']} {e['qty']}{e['unit']}" for e in equip[:2]) or "—"

        row = day_tbl.rows[i+1]
        row.cells[0].text = thai_date_short(d.get("work_date",""))
        row.cells[1].text = d.get("weather_morning") or "—"
        row.cells[2].text = str(d.get("total_workers") or "—")
        row.cells[3].text = act_str or "—"
        row.cells[4].text = eq_str
        row.cells[0].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        row.cells[2].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        if i%2==0:
            for cell in row.cells: set_cell_bg(cell,"EBF3FB")
    doc.add_paragraph()

    # รูปภาพ + caption รายวัน
    doc.add_heading("3. ภาพความก้าวหน้าประจำสัปดาห์", level=2)
    for day_data in daily_list:
        images = day_data.get("images") or []
        acts   = day_data.get("activities") or []
        if not images and not acts: continue

        wp = doc.add_paragraph()
        wr = wp.add_run(f"▶  {thai_date(day_data.get('work_date',''))}")
        wr.bold = True; wr.font.size = Pt(14)
        wr.font.color.rgb = RGBColor(0x1F,0x4E,0x79)

        # กำลังพลและเครื่องจักรของวัน
        labor_parts = []
        if day_data.get("engineers"):       labor_parts.append(f"วิศวกร {day_data['engineers']} คน")
        if day_data.get("foremen"):         labor_parts.append(f"หัวหน้า {day_data['foremen']} คน")
        if day_data.get("skilled_workers"): labor_parts.append(f"ช่าง {day_data['skilled_workers']} คน")
        if day_data.get("laborers"):        labor_parts.append(f"กรรมกร {day_data['laborers']} คน")
        if labor_parts:
            lp = doc.add_paragraph(f"👷 {', '.join(labor_parts)}")
            lp.runs[0].font.size = Pt(12)

        if images:
            for img_info in images:
                url     = img_info.get("url") or img_info.get("image_url")
                acts_text = clean_caption(img_info.get("caption") or "")
                _wd = day_data.get("work_date", "")
                caption = f"วันที่ {thai_date(_wd)}" if _wd else ""
                if acts_text:
                    caption += f"\n\n{acts_text}" if caption else acts_text
                if not url: continue
                img_bytes = await download_image_bytes(url)
                if img_bytes:
                    p = doc.add_paragraph()
                    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    p.add_run().add_picture(io.BytesIO(img_bytes), width=Inches(5.0))
                cap_p = doc.add_paragraph(caption)
                cap_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                if cap_p.runs:
                    cap_p.runs[0].italic = True; cap_p.runs[0].font.size = Pt(12)
                else:
                    r2 = cap_p.add_run(caption); r2.italic = True; r2.font.size = Pt(12)
        elif acts:
            for act in acts:
                bp = doc.add_paragraph(act.get("desc") or act.get("description",""), style="List Bullet")
        doc.add_paragraph()

    doc.add_heading("ลงชื่อ / Signature", level=2)
    add_signature_table(doc, "ผู้รับจ้าง (Contractor)", "ผู้ควบคุมงาน (Inspector)")

    buf = io.BytesIO(); doc.save(buf); return buf.getvalue()


# ════════════════════════════════════════
# MONTHLY REPORT
# ════════════════════════════════════════

async def generate_monthly(month_str: str, daily_list: list, project_name: str = "โครงการก่อสร้าง") -> bytes:
    doc = Document()
    style_doc(doc)
    yr, mo = int(month_str[:4]), int(month_str[5:7])
    month_label = f"{THAI_MONTHS_FULL[mo]} {yr+543}"
    add_title_block(doc, "รายงานความก้าวหน้าประจำเดือน (MONTHLY PROGRESS REPORT)",
        f"{month_label}  |  ฉบับที่ MPR-{month_str.replace('-','')}", project_name)

    # สรุปเดือน
    doc.add_heading("1. สรุปภาพรวมประจำเดือน", level=2)
    total_w = sum(d.get("total_workers") or 0 for d in daily_list)
    sum_tbl = doc.add_table(rows=2, cols=4)
    sum_tbl.style = "Table Grid"
    add_header_row(sum_tbl,["วันทำงาน","คนงานรวม (คน-วัน)","กิจกรรมรวม","รูปภาพรวม"])
    r = sum_tbl.rows[1]
    for i,v in enumerate([f"{len(daily_list)} วัน", str(total_w),
                           str(sum(len(d.get("activities") or []) for d in daily_list)),
                           str(sum(len(d.get("images") or []) for d in daily_list))]):
        r.cells[i].text = v
        r.cells[i].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
    doc.add_paragraph()

    # ตารางสรุปรายวัน
    doc.add_heading("2. ตารางสรุปงานประจำเดือน", level=2)
    day_tbl = doc.add_table(rows=len(daily_list)+1, cols=5)
    day_tbl.style = "Table Grid"
    add_header_row(day_tbl,["วันที่","อากาศ","คนงาน (คน)","กิจกรรมหลัก","เครื่องจักร"])
    for i,d in enumerate(daily_list):
        acts = d.get("activities") or []
        act_str = ", ".join(a.get("desc") or a.get("description","") for a in acts[:2])
        if len(acts)>2: act_str += f"(+{len(acts)-2})"

        equip_raw = d.get("equipment")
        if isinstance(equip_raw, str):
            try: equip = json.loads(equip_raw)
            except: equip = []
        else: equip = equip_raw or []
        eq_str = ", ".join(f"{e['name']} {e['qty']}{e['unit']}" for e in equip[:2]) or "—"

        row = day_tbl.rows[i+1]
        row.cells[0].text = thai_date_short(d.get("work_date",""))
        row.cells[1].text = d.get("weather_morning") or "—"
        row.cells[2].text = str(d.get("total_workers") or "—")
        row.cells[3].text = act_str or "—"
        row.cells[4].text = eq_str
        row.cells[0].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        row.cells[2].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        if i%2==0:
            for cell in row.cells: set_cell_bg(cell,"EBF3FB")
    doc.add_paragraph()

    # รูปภาพ
    doc.add_heading("3. ภาพความก้าวหน้าประจำเดือน", level=2)
    for day_data in daily_list:
        images = day_data.get("images") or []
        if not images: continue
        dp = doc.add_paragraph()
        dr = dp.add_run(f"▶  {thai_date(day_data.get('work_date',''))}")
        dr.bold = True; dr.font.size = Pt(13)
        dr.font.color.rgb = RGBColor(0x1F,0x4E,0x79)
        for img_info in images:
            url     = img_info.get("url") or img_info.get("image_url")
            acts_text = clean_caption(img_info.get("caption") or "")
            _wd = day_data.get("work_date", "")
            caption = f"วันที่ {thai_date(_wd)}" if _wd else ""
            if acts_text:
                caption += f"\n{acts_text}" if caption else acts_text
            if not url: continue
            img_bytes = await download_image_bytes(url)
            if img_bytes:
                p = doc.add_paragraph()
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                p.add_run().add_picture(io.BytesIO(img_bytes), width=Inches(4.5))
            cap_p = doc.add_paragraph(caption)
            cap_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            if cap_p.runs:
                cap_p.runs[0].italic = True; cap_p.runs[0].font.size = Pt(11)
            else:
                r2 = cap_p.add_run(caption); r2.italic = True; r2.font.size = Pt(11)
        doc.add_paragraph()

    doc.add_heading("ลงชื่อ / Signature", level=2)
    add_signature_table(doc, "ผู้รับจ้าง (Contractor)", "ผู้ควบคุมงาน (Inspector)")

    buf = io.BytesIO(); doc.save(buf); return buf.getvalue()
