"""
report_generator.py
===================
สร้างไฟล์ Word รายงานก่อสร้างจากข้อมูลใน Supabase
รองรับ: Daily / Weekly / Monthly Report

ใช้งาน:
  from report_generator import generate_daily, generate_weekly, generate_monthly
"""

import io
import httpx
from datetime import datetime, date, timedelta
from docx import Document
from docx.shared import Inches, Pt, RGBColor, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

# ─────────────────────────────────────────
# ชื่อเดือนภาษาไทย
# ─────────────────────────────────────────
THAI_MONTHS_FULL = [
    "", "มกราคม", "กุมภาพันธ์", "มีนาคม", "เมษายน",
    "พฤษภาคม", "มิถุนายน", "กรกฎาคม", "สิงหาคม",
    "กันยายน", "ตุลาคม", "พฤศจิกายน", "ธันวาคม",
]

def thai_date(d: date | str) -> str:
    """แปลง date เป็น '23 เมษายน 2568'"""
    if isinstance(d, str):
        d = date.fromisoformat(d)
    return f"{d.day} {THAI_MONTHS_FULL[d.month]} {d.year + 543}"

def thai_date_short(d: date | str) -> str:
    """แปลง date เป็น '23/04/68'"""
    if isinstance(d, str):
        d = date.fromisoformat(d)
    return f"{d.day:02d}/{d.month:02d}/{str(d.year + 543)[2:]}"


# ─────────────────────────────────────────
# Styling Helpers
# ─────────────────────────────────────────

def set_cell_bg(cell, hex_color: str):
    """ตั้งสีพื้นหลัง cell"""
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:fill"), hex_color)
    shd.set(qn("w:val"), "clear")
    tcPr.append(shd)

def set_col_width(table, col_idx: int, width_cm: float):
    """กำหนดความกว้าง column"""
    for row in table.rows:
        row.cells[col_idx].width = Cm(width_cm)

def add_header_row(table, headers: list, bg_color="1F4E79"):
    """เพิ่ม header row สีน้ำเงินเข้ม"""
    row = table.rows[0]
    for i, h in enumerate(headers):
        cell = row.cells[i]
        cell.text = h
        set_cell_bg(cell, bg_color)
        para = cell.paragraphs[0]
        para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = para.runs[0] if para.runs else para.add_run(h)
        run.text = h
        run.bold = True
        run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
        run.font.size = Pt(10)

def style_doc(doc: Document, project_name: str = ""):
    """ตั้งค่า font และ margin เริ่มต้น"""
    style = doc.styles["Normal"]
    style.font.name = "TH Sarabun New"
    style.font.size = Pt(14)
    # margin A4
    for section in doc.sections:
        section.page_height = Cm(29.7)
        section.page_width  = Cm(21.0)
        section.left_margin   = Cm(2.5)
        section.right_margin  = Cm(2.0)
        section.top_margin    = Cm(2.0)
        section.bottom_margin = Cm(2.0)

def add_title_block(doc: Document, title: str, subtitle: str, project_name: str):
    """เพิ่มหัวรายงาน"""
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run(title)
    run.bold = True
    run.font.size = Pt(18)
    run.font.color.rgb = RGBColor(0x1F, 0x4E, 0x79)

    p2 = doc.add_paragraph()
    p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run2 = p2.add_run(subtitle)
    run2.font.size = Pt(13)

    if project_name:
        p3 = doc.add_paragraph()
        p3.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run3 = p3.add_run(f"โครงการ: {project_name}")
        run3.font.size = Pt(13)
        run3.bold = True

    # เส้นคั่น
    p4 = doc.add_paragraph()
    p4.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run4 = p4.add_run("─" * 60)
    run4.font.size = Pt(10)
    run4.font.color.rgb = RGBColor(0x1F, 0x4E, 0x79)
    doc.add_paragraph()


async def download_image_bytes(url: str) -> bytes | None:
    """ดาวน์โหลดรูปภาพจาก URL"""
    try:
        async with httpx.AsyncClient(timeout=15) as client:
            resp = await client.get(url)
            resp.raise_for_status()
            return resp.content
    except Exception as e:
        print(f"⚠️ ดาวน์โหลดรูปไม่ได้: {url} — {e}")
        return None


# ════════════════════════════════════════
# DAILY REPORT
# ════════════════════════════════════════

async def generate_daily(
    work_date: str,
    daily_data: dict,
    project_name: str = "โครงการก่อสร้าง",
) -> bytes:
    """
    สร้าง Daily Report Word
    daily_data มาจาก v_daily_report_full:
      work_date, weather_morning, total_workers, supervisor,
      activities: [{seq, type, desc, qty, unit, location}],
      images:     [{url, caption, category}]
    """
    doc = Document()
    style_doc(doc, project_name)

    d = date.fromisoformat(work_date)
    doc_no = f"DR-{d.strftime('%Y%m%d')}"

    add_title_block(
        doc,
        "รายงานประจำวัน (DAILY REPORT)",
        f"วันที่ {thai_date(d)}  |  ฉบับที่ {doc_no}",
        project_name,
    )

    # ── ข้อมูลโครงการ ──────────────────────────
    doc.add_heading("1. ข้อมูลทั่วไป", level=2)
    info_table = doc.add_table(rows=2, cols=4)
    info_table.style = "Table Grid"
    labels = ["วันที่ทำงาน", "สภาพอากาศ", "จำนวนคนงาน", "ผู้ควบคุมงาน"]
    values = [
        thai_date(d),
        daily_data.get("weather_morning") or "—",
        str(daily_data.get("total_workers") or "—") + " คน",
        daily_data.get("supervisor") or "—",
    ]
    add_header_row(info_table, labels)
    for i, v in enumerate(values):
        cell = info_table.rows[1].cells[i]
        cell.text = v
        cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
    doc.add_paragraph()

    # ── กิจกรรมงาน ─────────────────────────────
    doc.add_heading("2. งานที่ดำเนินการ", level=2)
    activities = daily_data.get("activities") or []
    if activities:
        act_table = doc.add_table(rows=len(activities) + 1, cols=4)
        act_table.style = "Table Grid"
        add_header_row(act_table, ["ลำดับ", "รายการงาน", "สถานที่", "หมายเหตุ"])
        for i, act in enumerate(activities):
            row = act_table.rows[i + 1]
            row.cells[0].text = str(act.get("seq") or i + 1)
            row.cells[1].text = act.get("desc") or act.get("description") or "—"
            row.cells[2].text = act.get("location") or "—"
            row.cells[3].text = (
                f"{act['qty']:g} {act['unit']}"
                if act.get("qty") and act.get("unit")
                else "—"
            )
            row.cells[0].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
            row.cells[3].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
            if i % 2 == 0:
                for cell in row.cells:
                    set_cell_bg(cell, "EBF3FB")
    else:
        doc.add_paragraph("— ไม่มีข้อมูลกิจกรรม —")
    doc.add_paragraph()

    # ── หมายเหตุ ──────────────────────────────
    remarks = daily_data.get("remarks")
    if remarks:
        doc.add_heading("3. หมายเหตุ / ปัญหาที่พบ", level=2)
        doc.add_paragraph(remarks)
        doc.add_paragraph()

    # ── รูปภาพ ─────────────────────────────────
    images = daily_data.get("images") or []
    if images:
        doc.add_heading(f"{'4' if remarks else '3'}. รูปภาพประกอบ", level=2)
        for img_info in images:
            url     = img_info.get("url") or img_info.get("image_url")
            caption = img_info.get("caption") or ""
            if not url:
                continue
            img_bytes = await download_image_bytes(url)
            if img_bytes:
                img_stream = io.BytesIO(img_bytes)
                p = doc.add_paragraph()
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                run = p.add_run()
                run.add_picture(img_stream, width=Inches(5.5))
            # caption ใต้รูป
            cap_p = doc.add_paragraph(caption)
            cap_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            cap_run = cap_p.runs[0] if cap_p.runs else cap_p.add_run(caption)
            cap_run.italic = True
            cap_run.font.size = Pt(12)
            doc.add_paragraph()

    # ── ลายเซ็น ────────────────────────────────
    doc.add_heading("ลงชื่อ / Signature", level=2)
    sig_table = doc.add_table(rows=3, cols=2)
    sig_table.style = "Table Grid"
    sig_table.rows[0].cells[0].text = "ผู้รายงาน (Reported by)"
    sig_table.rows[0].cells[1].text = "ผู้ตรวจสอบ (Checked by)"
    sig_table.rows[1].cells[0].text = "\n\n"
    sig_table.rows[1].cells[1].text = "\n\n"
    sig_table.rows[2].cells[0].text = "วันที่: ____________________"
    sig_table.rows[2].cells[1].text = "วันที่: ____________________"
    for cell in sig_table.rows[0].cells:
        set_cell_bg(cell, "D6E4F0")
        cell.paragraphs[0].runs[0].bold = True if cell.paragraphs[0].runs else None

    buf = io.BytesIO()
    doc.save(buf)
    return buf.getvalue()


# ════════════════════════════════════════
# WEEKLY REPORT
# ════════════════════════════════════════

async def generate_weekly(
    week_start: str,
    daily_list: list[dict],
    project_name: str = "โครงการก่อสร้าง",
) -> bytes:
    """
    สร้าง Weekly Report Word
    daily_list: list ของ daily_data แต่ละวัน (จาก v_daily_report_full)
    """
    doc = Document()
    style_doc(doc, project_name)

    ws = date.fromisoformat(week_start)
    we = ws + timedelta(days=6)
    doc_no = f"WR-{ws.strftime('%Y%m%d')}"

    add_title_block(
        doc,
        "รายงานความก้าวหน้าประจำสัปดาห์ (WEEKLY PROGRESS REPORT)",
        f"{thai_date(ws)} — {thai_date(we)}  |  ฉบับที่ {doc_no}",
        project_name,
    )

    # ── สรุปสัปดาห์ ───────────────────────────
    doc.add_heading("1. สรุปภาพรวมสัปดาห์", level=2)
    total_workers = sum(d.get("total_workers") or 0 for d in daily_list)
    total_acts    = sum(len(d.get("activities") or []) for d in daily_list)
    total_imgs    = sum(len(d.get("images") or []) for d in daily_list)
    working_days  = len([d for d in daily_list if d])

    sum_table = doc.add_table(rows=2, cols=4)
    sum_table.style = "Table Grid"
    add_header_row(sum_table, ["วันทำงาน", "คนงานรวม (คน-วัน)", "กิจกรรมรวม", "รูปภาพรวม"])
    r = sum_table.rows[1]
    for i, v in enumerate([
        f"{working_days} วัน",
        str(total_workers),
        str(total_acts),
        str(total_imgs),
    ]):
        r.cells[i].text = v
        r.cells[i].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
    doc.add_paragraph()

    # ── รูปภาพ + คำบรรยาย รายวัน ──────────────
    doc.add_heading("2. ภาพความก้าวหน้าและกิจกรรมประจำสัปดาห์", level=2)

    for day_data in daily_list:
        if not day_data:
            continue
        work_date = day_data.get("work_date", "")
        images    = day_data.get("images") or []
        activities= day_data.get("activities") or []

        if not images and not activities:
            continue

        # หัวข้อวัน
        day_heading = doc.add_paragraph()
        day_run = day_heading.add_run(f"▶  {thai_date(work_date)}")
        day_run.bold = True
        day_run.font.size = Pt(14)
        day_run.font.color.rgb = RGBColor(0x1F, 0x4E, 0x79)

        # รูปภาพ + caption
        if images:
            for img_info in images:
                url     = img_info.get("url") or img_info.get("image_url")
                caption = img_info.get("caption") or ""
                if not url:
                    continue
                img_bytes = await download_image_bytes(url)
                if img_bytes:
                    img_stream = io.BytesIO(img_bytes)
                    p = doc.add_paragraph()
                    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    run = p.add_run()
                    run.add_picture(img_stream, width=Inches(5.0))
                # caption ใต้รูป
                cap_p = doc.add_paragraph(caption)
                cap_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                if cap_p.runs:
                    cap_p.runs[0].italic = True
                    cap_p.runs[0].font.size = Pt(12)
                else:
                    r = cap_p.add_run(caption)
                    r.italic = True
                    r.font.size = Pt(12)
        elif activities:
            # ถ้าไม่มีรูป แสดง text กิจกรรมแทน
            for act in activities:
                bp = doc.add_paragraph(
                    act.get("desc") or act.get("description") or "",
                    style="List Bullet"
                )
        doc.add_paragraph()

    # ── ลายเซ็น ────────────────────────────────
    doc.add_heading("ลงชื่อ / Signature", level=2)
    sig_table = doc.add_table(rows=3, cols=2)
    sig_table.style = "Table Grid"
    sig_table.rows[0].cells[0].text = "ผู้รับจ้าง (Contractor)"
    sig_table.rows[0].cells[1].text = "ผู้ควบคุมงาน (Inspector)"
    sig_table.rows[1].cells[0].text = "\n\n"
    sig_table.rows[1].cells[1].text = "\n\n"
    sig_table.rows[2].cells[0].text = "วันที่: ____________________"
    sig_table.rows[2].cells[1].text = "วันที่: ____________________"
    for cell in sig_table.rows[0].cells:
        set_cell_bg(cell, "D6E4F0")

    buf = io.BytesIO()
    doc.save(buf)
    return buf.getvalue()


# ════════════════════════════════════════
# MONTHLY REPORT
# ════════════════════════════════════════

async def generate_monthly(
    month_str: str,
    daily_list: list[dict],
    project_name: str = "โครงการก่อสร้าง",
) -> bytes:
    """
    สร้าง Monthly Report Word
    month_str: "2026-04"
    daily_list: list ของ daily_data ทั้งเดือน
    """
    doc = Document()
    style_doc(doc, project_name)

    year, month = int(month_str[:4]), int(month_str[5:7])
    month_label = f"{THAI_MONTHS_FULL[month]} {year + 543}"
    doc_no = f"MPR-{month_str.replace('-', '')}"

    add_title_block(
        doc,
        "รายงานความก้าวหน้าประจำเดือน (MONTHLY PROGRESS REPORT)",
        f"{month_label}  |  ฉบับที่ {doc_no}",
        project_name,
    )

    # ── สรุปเดือน ──────────────────────────────
    doc.add_heading("1. สรุปภาพรวมประจำเดือน", level=2)
    working_days  = len([d for d in daily_list if d])
    total_workers = sum(d.get("total_workers") or 0 for d in daily_list)
    total_acts    = sum(len(d.get("activities") or []) for d in daily_list)
    total_imgs    = sum(len(d.get("images") or []) for d in daily_list)

    sum_table = doc.add_table(rows=2, cols=4)
    sum_table.style = "Table Grid"
    add_header_row(sum_table, ["วันทำงาน", "คนงานรวม (คน-วัน)", "กิจกรรมรวม", "รูปภาพรวม"])
    r = sum_table.rows[1]
    for i, v in enumerate([
        f"{working_days} วัน",
        str(total_workers),
        str(total_acts),
        str(total_imgs),
    ]):
        r.cells[i].text = v
        r.cells[i].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
    doc.add_paragraph()

    # ── ตารางสรุปรายวัน ────────────────────────
    doc.add_heading("2. ตารางสรุปงานประจำเดือน", level=2)
    sum2 = doc.add_table(rows=len(daily_list) + 1, cols=4)
    sum2.style = "Table Grid"
    add_header_row(sum2, ["วันที่", "สภาพอากาศ", "คนงาน (คน)", "กิจกรรมหลัก"])
    for i, d in enumerate(daily_list):
        row = sum2.rows[i + 1]
        acts = d.get("activities") or []
        act_str = ", ".join(
            a.get("desc") or a.get("description") or ""
            for a in acts[:2]
        )
        if len(acts) > 2:
            act_str += f" (+{len(acts)-2})"
        row.cells[0].text = thai_date_short(d.get("work_date", ""))
        row.cells[1].text = d.get("weather_morning") or "—"
        row.cells[2].text = str(d.get("total_workers") or "—")
        row.cells[3].text = act_str or "—"
        row.cells[0].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        row.cells[2].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        if i % 2 == 0:
            for cell in row.cells:
                set_cell_bg(cell, "EBF3FB")
    doc.add_paragraph()

    # ── รูปภาพ + caption รายวัน ─────────────────
    doc.add_heading("3. ภาพความก้าวหน้าประจำเดือน", level=2)

    for day_data in daily_list:
        if not day_data:
            continue
        images = day_data.get("images") or []
        if not images:
            continue
        work_date = day_data.get("work_date", "")

        day_p = doc.add_paragraph()
        dr = day_p.add_run(f"▶  {thai_date(work_date)}")
        dr.bold = True
        dr.font.size = Pt(13)
        dr.font.color.rgb = RGBColor(0x1F, 0x4E, 0x79)

        for img_info in images:
            url     = img_info.get("url") or img_info.get("image_url")
            caption = img_info.get("caption") or ""
            if not url:
                continue
            img_bytes = await download_image_bytes(url)
            if img_bytes:
                img_stream = io.BytesIO(img_bytes)
                p = doc.add_paragraph()
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                run = p.add_run()
                run.add_picture(img_stream, width=Inches(4.5))
            cap_p = doc.add_paragraph(caption)
            cap_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            if cap_p.runs:
                cap_p.runs[0].italic = True
                cap_p.runs[0].font.size = Pt(11)
            else:
                r2 = cap_p.add_run(caption)
                r2.italic = True
                r2.font.size = Pt(11)
        doc.add_paragraph()

    # ── ลายเซ็น ────────────────────────────────
    doc.add_heading("ลงชื่อ / Signature", level=2)
    sig_table = doc.add_table(rows=3, cols=2)
    sig_table.style = "Table Grid"
    sig_table.rows[0].cells[0].text = "ผู้รับจ้าง (Contractor)"
    sig_table.rows[0].cells[1].text = "ผู้ควบคุมงาน (Inspector)"
    sig_table.rows[1].cells[0].text = "\n\n"
    sig_table.rows[1].cells[1].text = "\n\n"
    sig_table.rows[2].cells[0].text = "วันที่: ____________________"
    sig_table.rows[2].cells[1].text = "วันที่: ____________________"
    for cell in sig_table.rows[0].cells:
        set_cell_bg(cell, "D6E4F0")

    buf = io.BytesIO()
    doc.save(buf)
    return buf.getvalue()
