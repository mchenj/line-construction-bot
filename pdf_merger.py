"""
pdf_merger.py
รวม DOCX + PPTX → PDF เดียว ใช้ LibreOffice headless

Requirements:
- LibreOffice ต้องติดตั้งบน Railway (ใช้ nixpacks หรือ apt-get install libreoffice)
- pypdf สำหรับ merge PDF
"""

import os, io, tempfile, subprocess, zipfile
from typing import List, Tuple


def _find_soffice() -> str:
    """หา path ของ LibreOffice/soffice executable"""
    candidates = [
        "soffice",
        "libreoffice",
        "/usr/bin/soffice",
        "/usr/bin/libreoffice",
        "/Applications/LibreOffice.app/Contents/MacOS/soffice",
        r"C:\Program Files\LibreOffice\program\soffice.exe",
        r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
    ]
    for c in candidates:
        try:
            r = subprocess.run([c, "--version"], capture_output=True, timeout=10)
            if r.returncode == 0:
                return c
        except (FileNotFoundError, subprocess.TimeoutExpired, OSError):
            continue
    raise RuntimeError("ไม่พบ LibreOffice (soffice). ต้องติดตั้งก่อน")


def docx_pptx_to_pdf(input_path: str, output_dir: str) -> str:
    """แปลง DOCX/PPTX → PDF ผ่าน LibreOffice headless (ใช้ font ตามที่กำหนดใน docx)
    คืนค่า path ของ PDF ที่ได้
    """
    soffice = _find_soffice()
    cmd = [soffice, "--headless", "--convert-to", "pdf",
           "--outdir", output_dir, input_path]
    r = subprocess.run(cmd, capture_output=True, timeout=180)
    if r.returncode != 0:
        raise RuntimeError(f"LibreOffice convert failed: {r.stderr.decode('utf-8', errors='ignore')}")

    # PDF ออกมาในชื่อเดียวกับ input แต่นามสกุล .pdf
    base = os.path.splitext(os.path.basename(input_path))[0]
    pdf_path = os.path.join(output_dir, base + ".pdf")
    if not os.path.exists(pdf_path):
        raise RuntimeError(f"PDF output not found: {pdf_path}")
    return pdf_path


def merge_pdfs(pdf_paths: List[str], output_path: str):
    """รวม PDF หลายไฟล์เป็นไฟล์เดียว ใช้ pypdf"""
    try:
        from pypdf import PdfWriter
    except ImportError:
        from PyPDF2 import PdfWriter
    writer = PdfWriter()
    for p in pdf_paths:
        if os.path.exists(p):
            writer.append(p)
    with open(output_path, "wb") as f:
        writer.write(f)


def zip_to_pdf(zip_bytes: bytes, file_order: List[str] = None) -> bytes:
    """รับ ZIP bytes (จาก generate_weekly_phase1) → คืน PDF เดียว
    file_order: list ของชื่อไฟล์ตามลำดับที่ต้องการเรียง
                ถ้าไม่ระบุ จะเรียงตาม alphabetical (00_, 01_, 02_, ...)
    """
    with tempfile.TemporaryDirectory() as tmpdir:
        extracted = []
        # extract zip
        with zipfile.ZipFile(io.BytesIO(zip_bytes)) as zf:
            names = zf.namelist()
            if file_order:
                # filter ตามลำดับที่ระบุ
                names = [n for n in file_order if n in names]
            else:
                # default: เรียงชื่อไฟล์ (รับประกัน 00_, 01_, ... มาก่อน)
                names = sorted([n for n in names if not n.startswith("ERROR_")])

            for name in names:
                # ตัด path-unsafe chars และเก็บ original ext
                safe_name = "".join(c if c.isalnum() or c in "._-" else "_" for c in name)
                fp = os.path.join(tmpdir, safe_name)
                with open(fp, "wb") as f:
                    f.write(zf.read(name))
                extracted.append(fp)

        # convert each → PDF
        pdf_dir = os.path.join(tmpdir, "pdfs")
        os.makedirs(pdf_dir, exist_ok=True)
        pdf_paths = []
        for src in extracted:
            ext = os.path.splitext(src)[1].lower()
            if ext in (".docx", ".pptx", ".xlsx"):
                try:
                    pdf_paths.append(docx_pptx_to_pdf(src, pdf_dir))
                except Exception as e:
                    print(f"⚠️ skip {src}: {e}")

        # merge
        out_path = os.path.join(tmpdir, "merged.pdf")
        merge_pdfs(pdf_paths, out_path)
        with open(out_path, "rb") as f:
            return f.read()


async def generate_weekly_phase1_pdf(week_no: int, week_start: str,
                                     daily_list: list,
                                     project_name: str = "โครงการก่อสร้าง",
                                     week_end: str = None) -> bytes:
    """ครบทุกอย่าง: generate weekly + รวมเป็น PDF เดียว"""
    from weekly_phase1 import generate_weekly_phase1
    zip_bytes = await generate_weekly_phase1(week_no, week_start, daily_list,
                                             project_name, week_end=week_end)
    return zip_to_pdf(zip_bytes)
