# Custom Fonts for PDF Rendering

วางไฟล์ฟอนต์ `.ttf` ที่ใช้ใน template ที่นี่ — Nixpacks จะ copy ไปติดตั้งบน Railux/Linux ตอน build

## ฟอนต์ที่ต้องการ

| ฟอนต์ | ใช้ที่ | ที่ดาวน์โหลด |
|---|---|---|
| **TH SarabunIT๙** (`THSarabunIT9.ttf` + Bold/Italic) | template ทั้งหมด | https://www.f0nt.com/release/th-sarabun-it๙/ |
| **TH SarabunPSK** (สำรอง) | บางตำแหน่ง | https://www.f0nt.com/release/th-sarabun-psk/ |

## วิธีติดตั้ง

### บนเครื่อง Local (Windows)
1. ดาวน์โหลด `.ttf` จากลิงก์ด้านบน  
2. คลิกขวา → Install for all users (ติดตั้งใน Windows)  
3. ก็อปไฟล์ `.ttf` ทั้ง 4 ไฟล์ (Regular, Bold, Italic, BoldItalic) มาวางใน `fonts/` นี้

### บน Railway
1. Push folder `fonts/` ที่มีไฟล์ `.ttf` ขึ้น git  
2. `nixpacks.toml` จะ auto-copy ไป `/usr/share/fonts/truetype/custom/`  
3. รัน `fc-cache -f -v` ตอน build  
4. LibreOffice จะใช้ฟอนต์ตามที่ docx กำหนดได้ถูกต้อง  

## ตรวจสอบ

หลัง deploy เปิด `/admin/check_fonts?token=xxx` ควรเห็น "TH SarabunIT๙" ในรายการ
