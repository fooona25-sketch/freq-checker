# คู่มือติดตั้งแอปตรวจสอบความถี่ดาวเทียม (เฟส 1)

> เฟส 1 = เครื่องมือเทียบความถี่ล้วน **ฟรี ไม่ต้องใช้ Claude API key**
> คุณไม่ต้องเขียนโค้ดเลย ทำตามทีละขั้น

ภาพรวม 3 ส่วน:
- **โค้ดแอป** → เก็บบน GitHub (อัปครั้งเดียว ไม่ต้องแตะอีก)
- **ฐานความถี่ไทย** → Google Sheet (คุณแก้เองได้ตลอด เหมือน Excel)
- **เว็บแอป** → โฮสต์ฟรีบน Streamlit Community Cloud (ได้ลิงก์เปิดจากบ้าน)

---

## ขั้นที่ 1 — สมัคร GitHub และอัปไฟล์โค้ด
1. ไปที่ https://github.com แล้วสมัครบัญชีฟรี
2. กดปุ่ม **New** เพื่อสร้าง repository ใหม่ ตั้งชื่อเช่น `thaisat-freq-checker` เลือก **Private** แล้วกด Create
3. ในหน้า repo กด **Add file > Upload files** แล้วลาก **ทุกไฟล์ในโฟลเดอร์นี้** เข้าไป
   (`app.py`, `engine.py`, `requirements.txt`, `packages.txt`, โฟลเดอร์ `.streamlit`)
4. กด **Commit changes**

## ขั้นที่ 2 — สร้างฐานความถี่ไทยใน Google Sheet
1. เปิด https://sheets.google.com สร้างชีตใหม่
2. **File > Import > Upload** เลือกไฟล์ `thai_reference_template.csv` (เลือก "Replace current sheet")
   - คอลัมน์: `network_name | type | direction | freq_start_mhz | freq_end_mhz | note`
   - ช่อง `type` (GSO/NGSO) และ `direction` (Tx/Rx) เว้นว่างได้ ใส่ทีหลังเพื่อให้กรองคมขึ้น
3. **File > Share > Publish to web** → เลือกทั้งชีต → รูปแบบ **Comma-separated values (.csv)** → กด Publish
4. คัดลอกลิงก์ที่ได้ (ลงท้ายด้วย `output=csv`) เก็บไว้ใช้ในขั้นที่ 3

> เวลามีข่ายใหม่/ย่านใหม่/ต้องลบ — แค่แก้ในชีตนี้ แอปจะใช้ของล่าสุดทันที ไม่ต้องทำอะไรเพิ่ม

## ขั้นที่ 3 — Deploy บน Streamlit (ฟรี)
1. ไปที่ https://share.streamlit.io กด **Sign in with GitHub**
2. กด **Create app > Deploy a public app from GitHub**
3. เลือก repo `thaisat-freq-checker`, branch `main`, main file `app.py`
4. กด **Advanced settings > Secrets** แล้ววางข้อความนี้ (แก้ค่าให้เป็นของคุณ):
   ```toml
   password = "ตั้งรหัสผ่านของคุณ"
   thai_sheet_csv_url = "วางลิงก์ CSV จากขั้นที่ 2 ที่นี่"
   ```
5. กด **Deploy** รอสักครู่ จะได้ลิงก์เว็บ เช่น `https://xxxx.streamlit.app`

## ขั้นที่ 4 — ใช้งานและแชร์
- เปิดลิงก์ → ใส่รหัสผ่าน → อัปโหลด `ificXXXX.mdb` ของรอบนั้น → กด "เริ่มเปรียบเทียบ" → ดาวน์โหลด Excel
- แชร์ลิงก์ + รหัสผ่าน ให้เพื่อนร่วมกลุ่ม ดบ. ใช้จากบ้านได้เลย

---
### หมายเหตุ
- ถ้าไฟล์ .mdb ใหญ่จนเครื่องฟรีไม่ไหว แจ้งได้ จะย้ายไป Hugging Face Spaces ที่แรงกว่า
- เฟส 2 (ปุ่มร่างบันทึก/หนังสือด้วย Claude API) ค่อยเพิ่มภายหลังโดยใส่ API key ใน Secrets
