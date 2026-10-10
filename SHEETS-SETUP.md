# กระดานคะแนนกลาง Lean Quest บน Google Sheets (ตั้งค่าครั้งเดียว ~3 นาที)

1. สร้าง Google Sheets ใหม่ (ชีตเปล่า) ตั้งชื่อเช่น "Lean Quest Scores"
2. เมนู **Extensions → Apps Script** ลบโค้ดเดิมทั้งหมด แล้ววางเนื้อหาจากไฟล์ `Code.gs` → กด 💾 Save
3. กด **Deploy → New deployment** → ไอคอน ⚙️ เลือก **Web app**
   - Description: Lean Quest
   - Execute as: **Me**
   - Who has access: **Anyone**
   - กด Deploy → อนุญาตสิทธิ์ (Authorize access → เลือกบัญชี → Advanced → Go to … (unsafe) → Allow)
4. คัดลอก **Web app URL** (ลงท้ายด้วย `/exec`)
5. เปิดเกม → แท็บ **ฉัน** → ⚙️ กระดานคะแนนกลาง → วาง URL → **บันทึกและทดสอบ** ต้องขึ้น ✅
6. ส่ง URL ให้ Claude ฝังลงในตัวเกม เพื่อให้ผู้เล่นทุกคนส่งคะแนนอัตโนมัติโดยไม่ต้องตั้งค่า

**หลังบ้านวิทยากร:** ในเกม แท็บ ฉัน → 🔐 หลังบ้านวิทยากร → ใส่ PIN (ค่าเริ่มต้น `2468` เปลี่ยนได้ที่บรรทัด ADMIN_PIN ในสคริปต์ แล้ว Deploy เวอร์ชันใหม่) ดูอันดับทุกซีซัน ประวัติการเล่นทุกเกม และดาวน์โหลด Excel

ชีตที่ระบบสร้างให้เอง: `scores` (ทุกผลการเล่น), `season` (อันดับซีซันปัจจุบัน พร้อมคะแนนแยกส่วน), `players`
ซีซันละ 14 วัน เริ่ม 12 ต.ค. 2026 (แก้ได้ที่บรรทัด SEASON_EPOCH ในสคริปต์)
ถ้าแก้สคริปต์ภายหลัง ต้อง Deploy → Manage deployments → ✏️ → Version: New version → Deploy (URL เดิม)
