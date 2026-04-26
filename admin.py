"""
admin.py
Web admin panel — จัดการ data files (construction_plan.xlsx, cm_personnel.xlsx)
และทดสอบ trigger weekly report manually

Auth: ส่ง ?token=xxx ใน URL (env: ADMIN_TOKEN)

Endpoints:
- GET  /admin                    หน้า HTML แสดงสถานะ + form upload
- GET  /admin/download/{kind}    ดาวน์โหลดไฟล์ปัจจุบัน (kind=plan|cm)
- POST /admin/upload/{kind}      อัปโหลดไฟล์ใหม่
- POST /admin/trigger/weekly     สั่ง generate weekly + push ทันที
"""

import os
from pathlib import Path
from datetime import datetime
from fastapi import APIRouter, UploadFile, File, HTTPException, Request, Form
from fastapi.responses import FileResponse, HTMLResponse, JSONResponse


_THIS_DIR = Path(__file__).parent
_DATA_DIR = _THIS_DIR / "data"
_DATA_DIR.mkdir(exist_ok=True)

DATA_FILES = {
    "plan": (_DATA_DIR / "construction_plan.xlsx", "แผนงานก่อสร้าง"),
    "cm":   (_DATA_DIR / "cm_personnel.xlsx",      "บุคลากร CM"),
}

router = APIRouter(prefix="/admin", tags=["admin"])


def _check_token(token: str):
    expected = os.getenv("ADMIN_TOKEN", "")
    if not expected:
        raise HTTPException(403, "ADMIN_TOKEN ไม่ได้ตั้งค่าใน env")
    if token != expected:
        raise HTTPException(401, "Token ไม่ถูกต้อง")


def _file_info(path: Path) -> dict:
    if not path.exists():
        return {"exists": False, "size": 0, "modified": None}
    st = path.stat()
    return {
        "exists": True,
        "size": st.st_size,
        "modified": datetime.fromtimestamp(st.st_mtime).strftime("%Y-%m-%d %H:%M:%S"),
    }


@router.get("", response_class=HTMLResponse)
async def admin_home(token: str = ""):
    _check_token(token)
    rows = []
    for kind, (path, label) in DATA_FILES.items():
        info = _file_info(path)
        if info["exists"]:
            size_kb = info["size"] / 1024
            status = f"✅ {size_kb:.1f} KB · แก้ไขล่าสุด {info['modified']}"
        else:
            status = "❌ ยังไม่มีไฟล์"
        rows.append(f"""
        <tr>
          <td><b>{label}</b><br><code>{path.name}</code></td>
          <td>{status}</td>
          <td>
            <a href="/admin/download/{kind}?token={token}" class="btn">ดาวน์โหลด</a>
          </td>
          <td>
            <form action="/admin/upload/{kind}?token={token}" method="post" enctype="multipart/form-data" style="display:inline">
              <input type="file" name="file" accept=".xlsx" required style="font-size:13px">
              <button type="submit" class="btn primary">อัปโหลด</button>
            </form>
          </td>
        </tr>
        """)

    html = f"""<!DOCTYPE html>
<html lang="th"><head>
<meta charset="UTF-8"><title>Admin · LINE Construction Bot</title>
<style>
  body {{ font-family: 'Segoe UI', Tahoma, sans-serif; max-width: 900px; margin: 30px auto; padding: 0 20px; color: #222; }}
  h1 {{ color: #1F4E79; }}
  table {{ width: 100%; border-collapse: collapse; margin: 20px 0; background: #fff; box-shadow: 0 1px 3px rgba(0,0,0,.08); }}
  th, td {{ padding: 12px; border-bottom: 1px solid #eee; text-align: left; vertical-align: middle; }}
  th {{ background: #1F4E79; color: white; font-size: 14px; }}
  td code {{ background: #f5f5f5; padding: 2px 6px; border-radius: 3px; font-size: 12px; }}
  .btn {{ display: inline-block; padding: 6px 14px; border: 1px solid #1F4E79; color: #1F4E79;
         text-decoration: none; border-radius: 4px; font-size: 13px; cursor: pointer; background: white; }}
  .btn.primary {{ background: #1F4E79; color: white; }}
  .btn:hover {{ opacity: 0.85; }}
  .section {{ background: #f9f9f9; padding: 20px; border-radius: 6px; margin-top: 30px; }}
  .section h2 {{ margin-top: 0; color: #1F4E79; }}
  small {{ color: #666; }}
</style></head><body>

<h1>📊 Admin · LINE Construction Bot</h1>
<p>จัดการข้อมูลแผนงาน + บุคลากร CM ที่ใช้สร้างรายงานประจำสัปดาห์</p>

<table>
  <thead><tr><th>ไฟล์</th><th>สถานะ</th><th>ดาวน์โหลด</th><th>อัปโหลดใหม่</th></tr></thead>
  <tbody>{"".join(rows)}</tbody>
</table>

<div class="section">
  <h2>🔄 ทดสอบ Cron Weekly</h2>
  <p>ส่งรายงานสัปดาห์ปัจจุบันให้ <code>WEEKLY_CRON_USER_IDS</code> ทันที (ไม่รอเวลา cron)</p>
  <form action="/admin/trigger/weekly?token={token}" method="post">
    <button type="submit" class="btn primary">🚀 Trigger Weekly Report ตอนนี้</button>
  </form>
</div>

<div class="section">
  <h2>📋 วิธีใช้</h2>
  <ol>
    <li><b>ดาวน์โหลด</b> ไฟล์ปัจจุบันมาแก้ไข</li>
    <li>แก้ไขใน Excel — อัปเดต % ผลงาน หรือ attendance ในคอลัมน์ของวันที่ใหม่</li>
    <li><b>อัปโหลด</b> กลับมา (ทับไฟล์เดิม)</li>
    <li>สั่ง <code>/weekly2</code> ใน LINE เพื่อสร้างรายงานด้วยข้อมูลใหม่</li>
  </ol>
  <small>Tip: ใช้คำสั่ง <code>/upload_plan</code> หรือ <code>/upload_cm</code> ใน LINE แล้วส่งไฟล์ Excel ตามมา ก็อัปโหลดได้เหมือนกัน</small>
</div>

</body></html>"""
    return HTMLResponse(html)


@router.get("/download/{kind}")
async def admin_download(kind: str, token: str = ""):
    _check_token(token)
    if kind not in DATA_FILES:
        raise HTTPException(404, f"ไม่รู้จัก kind '{kind}'")
    path, _ = DATA_FILES[kind]
    if not path.exists():
        raise HTTPException(404, f"ไม่มีไฟล์ {path.name}")
    return FileResponse(path, filename=path.name,
                        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


@router.post("/upload/{kind}", response_class=HTMLResponse)
async def admin_upload(kind: str, token: str = "", file: UploadFile = File(...)):
    _check_token(token)
    if kind not in DATA_FILES:
        raise HTTPException(404, f"ไม่รู้จัก kind '{kind}'")
    if not file.filename.lower().endswith(".xlsx"):
        raise HTTPException(400, "รองรับเฉพาะไฟล์ .xlsx")
    path, label = DATA_FILES[kind]
    content = await file.read()
    # backup ไฟล์เก่า
    if path.exists():
        backup = path.with_suffix(f".{datetime.now().strftime('%Y%m%d_%H%M%S')}.bak.xlsx")
        path.rename(backup)
    path.write_bytes(content)
    return HTMLResponse(f"""<!DOCTYPE html>
<html><head><meta charset="UTF-8"><meta http-equiv="refresh" content="2;url=/admin?token={token}"></head>
<body style="font-family:sans-serif; padding:40px; text-align:center">
<h2>✅ อัปโหลดสำเร็จ</h2>
<p>{label}: <code>{file.filename}</code> ({len(content)/1024:.1f} KB)</p>
<p>กำลังกลับสู่หน้าหลัก...</p>
</body></html>""")


@router.post("/trigger/weekly")
async def admin_trigger_weekly(token: str = ""):
    _check_token(token)
    try:
        from scheduler import run_weekly_for_current_week
        await run_weekly_for_current_week()
        return JSONResponse({"status": "ok", "message": "Weekly report triggered — check LINE"})
    except Exception as e:
        return JSONResponse({"status": "error", "error": str(e)}, status_code=500)


@router.get("/check_fonts")
async def admin_check_fonts(token: str = ""):
    """ตรวจดูว่า Thai fonts ติดตั้งครบไหมบน server (สำหรับ debug PDF rendering)"""
    _check_token(token)
    import subprocess
    result = {"thai_fonts": [], "soffice": "not_found", "errors": []}
    # 1) check fonts via fc-list
    try:
        r = subprocess.run(["fc-list", ":lang=th"], capture_output=True, timeout=10)
        if r.returncode == 0:
            lines = r.stdout.decode("utf-8", errors="ignore").splitlines()
            result["thai_fonts"] = sorted(set(line.split(":")[1].strip() for line in lines if ":" in line))
        else:
            result["errors"].append(f"fc-list failed: {r.stderr.decode()[:200]}")
    except Exception as e:
        result["errors"].append(f"fc-list error: {e}")
    # 2) check soffice
    for c in ("soffice", "libreoffice", "/usr/bin/soffice"):
        try:
            r = subprocess.run([c, "--version"], capture_output=True, timeout=10)
            if r.returncode == 0:
                result["soffice"] = r.stdout.decode("utf-8", errors="ignore").strip()
                break
        except Exception:
            continue
    return JSONResponse(result)
