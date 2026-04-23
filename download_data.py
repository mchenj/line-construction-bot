"""
ดาวน์โหลดข้อมูลรายงานจาก Server มาเก็บที่เครื่อง
================================================
วิธีใช้:
  python download_data.py              → ดาวน์โหลดทั้งหมด
  python download_data.py 2026-04-23   → ดาวน์โหลดเฉพาะวันที่ระบุ
  python download_data.py week         → ดาวน์โหลด 7 วันล่าสุด
  python download_data.py month        → ดาวน์โหลด 30 วันล่าสุด
"""

import sys
import json
import os
import httpx
from datetime import datetime, date, timedelta
from pathlib import Path

# ════════════════════════════════════════
# ⚙️  ตั้งค่าก่อนใช้งาน
# ════════════════════════════════════════
SERVER_URL  = "https://YOUR-APP-NAME.railway.app"   # ← ใส่ URL จาก Railway
OUTPUT_DIR  = Path("construction_reports")           # โฟลเดอร์เก็บข้อมูล
BATCH_SIZE  = 100                                    # ดึงครั้งละกี่รายการ
# ════════════════════════════════════════


def ensure_dirs():
    """สร้างโฟลเดอร์เก็บข้อมูล"""
    (OUTPUT_DIR / "json").mkdir(parents=True, exist_ok=True)
    (OUTPUT_DIR / "images").mkdir(parents=True, exist_ok=True)
    (OUTPUT_DIR / "summary").mkdir(parents=True, exist_ok=True)


def fetch_reports(target_date: str = None, limit: int = BATCH_SIZE, offset: int = 0) -> list:
    """ดึงข้อมูลจาก server"""
    params = {"limit": limit, "offset": offset}
    if target_date:
        params["date"] = target_date

    try:
        resp = httpx.get(f"{SERVER_URL}/reports", params=params, timeout=30)
        resp.raise_for_status()
        return resp.json().get("reports", [])
    except Exception as e:
        print(f"❌ เชื่อมต่อ server ไม่ได้: {e}")
        print(f"   ตรวจสอบ SERVER_URL ในไฟล์ download_data.py")
        return []


def fetch_all_reports(target_date: str = None) -> list:
    """ดึงข้อมูลทั้งหมด (รองรับข้อมูลจำนวนมาก)"""
    all_reports = []
    offset = 0
    while True:
        batch = fetch_reports(target_date, BATCH_SIZE, offset)
        if not batch:
            break
        all_reports.extend(batch)
        if len(batch) < BATCH_SIZE:
            break
        offset += BATCH_SIZE
        print(f"  ดึงข้อมูลแล้ว {len(all_reports)} รายการ...")
    return all_reports


def download_image(url: str, filename: str) -> bool:
    """ดาวน์โหลดรูปภาพ"""
    save_path = OUTPUT_DIR / "images" / filename
    if save_path.exists():
        return True  # มีแล้ว ข้าม
    try:
        resp = httpx.get(url, timeout=30, follow_redirects=True)
        resp.raise_for_status()
        save_path.write_bytes(resp.content)
        return True
    except Exception as e:
        print(f"  ⚠️  ดาวน์โหลดรูป {filename} ไม่ได้: {e}")
        return False


def save_json_by_date(reports: list):
    """บันทึก JSON แยกตามวันที่"""
    by_date: dict[str, list] = {}
    for r in reports:
        d = r.get("work_date") or r.get("timestamp", "")[:10] or "unknown"
        by_date.setdefault(d, []).append(r)

    for d, day_reports in by_date.items():
        # แยก text และ image
        texts  = [r for r in day_reports if r.get("message_type") == "text"]
        images = [r for r in day_reports if r.get("message_type") == "image"]

        output = {
            "date":         d,
            "total_text":   len(texts),
            "total_images": len(images),
            "text_reports": texts,
            "image_reports": images,
        }
        filepath = OUTPUT_DIR / "json" / f"report_{d}.json"
        filepath.write_text(json.dumps(output, ensure_ascii=False, indent=2), encoding="utf-8")
        print(f"  💾 {filepath.name}  ({len(texts)} ข้อความ, {len(images)} รูป)")
    return by_date


def create_summary_csv(reports: list):
    """สร้างไฟล์สรุป CSV สำหรับนำเข้า Excel"""
    import csv
    rows = []
    for r in reports:
        if r.get("message_type") != "text":
            continue
        try:
            acts = json.loads(r.get("activities", "[]"))
            qtys = json.loads(r.get("quantities", "[]"))
        except Exception:
            acts, qtys = [], []

        qty_str = "; ".join(f"{q['amount']:g} {q['unit']}" for q in qtys)
        rows.append({
            "วันที่ทำงาน":  r.get("work_date", ""),
            "เวลาบันทึก":   r.get("timestamp", "")[:19].replace("T", " "),
            "งานที่ทำ":     ", ".join(acts),
            "ปริมาณงาน":   qty_str,
            "จำนวนคนงาน":  r.get("workers", ""),
            "สภาพอากาศ":   r.get("weather", ""),
            "ข้อความต้นฉบับ": r.get("raw_text", ""),
        })

    if not rows:
        return

    today_str = date.today().strftime("%Y%m%d")
    csv_path  = OUTPUT_DIR / "summary" / f"summary_{today_str}.csv"
    with open(csv_path, "w", newline="", encoding="utf-8-sig") as f:
        writer = csv.DictWriter(f, fieldnames=rows[0].keys())
        writer.writeheader()
        writer.writerows(rows)
    print(f"  📊 สรุป CSV: {csv_path}")


def main():
    ensure_dirs()

    # ─── กำหนดช่วงวันที่ ───────────────────────
    arg = sys.argv[1] if len(sys.argv) > 1 else None

    if arg == "week":
        target_date = None
        since_date  = date.today() - timedelta(days=7)
        print(f"⏳ ดาวน์โหลดรายงาน 7 วันล่าสุด (ตั้งแต่ {since_date})...")
    elif arg == "month":
        target_date = None
        since_date  = date.today() - timedelta(days=30)
        print(f"⏳ ดาวน์โหลดรายงาน 30 วันล่าสุด (ตั้งแต่ {since_date})...")
    elif arg and arg != "all":
        target_date = arg          # เช่น "2026-04-23"
        since_date  = None
        print(f"⏳ ดาวน์โหลดรายงานวัน {target_date}...")
    else:
        target_date = None
        since_date  = None
        print("⏳ ดาวน์โหลดรายงานทั้งหมด...")

    # ─── ดึงข้อมูล ─────────────────────────────
    reports = fetch_all_reports(target_date)

    # กรองตามช่วงวัน (ถ้ามี)
    if since_date:
        reports = [
            r for r in reports
            if (r.get("work_date") or "") >= str(since_date)
        ]

    if not reports:
        print("📭 ไม่พบข้อมูล")
        return

    print(f"\n✅ พบ {len(reports)} รายการ\n")

    # ─── บันทึก JSON ───────────────────────────
    print("📁 บันทึกไฟล์ JSON:")
    by_date = save_json_by_date(reports)

    # ─── สร้าง CSV ─────────────────────────────
    print("\n📊 สร้างไฟล์สรุป CSV:")
    create_summary_csv(reports)

    # ─── ดาวน์โหลดรูปภาพ ───────────────────────
    image_reports = [r for r in reports if r.get("image_url")]
    if image_reports:
        print(f"\n📸 ดาวน์โหลดรูปภาพ ({len(image_reports)} ไฟล์):")
        ok_count = 0
        for r in image_reports:
            fname = r.get("image_filename", f"{r['id']}.jpg")
            if download_image(r["image_url"], fname):
                ok_count += 1
                print(f"  ✅ {fname}")
        print(f"  ดาวน์โหลดสำเร็จ {ok_count}/{len(image_reports)} ไฟล์")

    # ─── สรุปผล ────────────────────────────────
    abs_path = OUTPUT_DIR.resolve()
    print(f"\n{'═'*50}")
    print(f"✅ เสร็จสิ้น! ข้อมูลถูกบันทึกที่:")
    print(f"   📁 {abs_path}")
    print(f"   ├── json/      ← รายงาน JSON แยกตามวัน")
    print(f"   ├── images/    ← รูปภาพจากไซต์งาน")
    print(f"   └── summary/   ← ไฟล์สรุป CSV (เปิดด้วย Excel)")
    print(f"{'═'*50}")


if __name__ == "__main__":
    main()
