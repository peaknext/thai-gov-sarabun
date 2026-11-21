# สเปก JSON สำหรับ `scripts/build_docx.py`

```bash
python3 scripts/build_docx.py spec.json output.docx [--pdf] [--font "TH Sarabun New"] [--arabic-digits]
```

- ตัวเลขอารบิกในทุกฟิลด์จะถูกแปลงเป็นเลขไทยอัตโนมัติ (ยกเว้นอีเมล URL ชื่อไฟล์ PDF) — ใช้ `--arabic-digits` หรือ `"digits": "arabic"` ถ้าหน่วยงานใช้เลขอารบิก
- `--pdf` แปลงเป็น PDF ด้วย LibreOffice เพื่อดูตัวอย่าง (ถ้าเครื่องไม่มีฟอนต์ TH SarabunPSK ภาพตัวอย่างจะใช้ฟอนต์อื่นแทน — ไฟล์ docx ยังถูกต้อง)
- ตัวอย่างครบทุกชนิดอยู่ใน `assets/examples/*.json`
- ครุฑใช้ภาพ `assets/garuda.png` (ตัดจากแบบที่ ๒๖ ในเอกสารอ้างอิง สูง ๓ ซม. / ย่อเป็น ๑.๕ ซม. สำหรับบันทึกข้อความ) — หน่วยงานควรแทนด้วยไฟล์ครุฑความละเอียดสูงของตนเองเมื่อพิมพ์ใช้จริง

## ฟิลด์ร่วม

| ฟิลด์ | ความหมาย |
|---|---|
| `type` | `external` · `internal` (= `memo`) · `stamped` · `order` · `regulation` · `bylaw` · `announcement` · `statement` · `news` · `certificate` · `minutes` (รับชื่อภาษาไทยได้ เช่น `"หนังสือภายนอก"`) |
| `secrecy` | ชั้นความลับ: `"ลับ"` `"ลับมาก"` `"ลับที่สุด"` — พิมพ์สีแดงกึ่งกลางบน/ล่างทุกหน้า |
| `urgency` | ชั้นความเร็ว: `"ด่วน"` `"ด่วนมาก"` `"ด่วนที่สุด"` (สีแดง ≥ ๓๒ pt) · `urgency_deadline`: ข้อความต่อท้าย เช่น `"ภายใน ๒๕ พฤศจิกายน ๒๕๖๙ เวลา ๑๒.๐๐ น."` (ใช้กับ urgency `"ด่วน"`) |
| `body` | รายการย่อหน้า (string) — ย่อหน้าแรกบรรทัด ๒.๕ ซม. อัตโนมัติ; ย่อหน้าที่ขึ้นต้น `๑.๑` / `(๑)` ย่อลึกขึ้นอัตโนมัติ · ใช้ object `{"text": "...", "indent": 3.5, "align": "center", "bold": true}` เพื่อกำหนดเอง |
| `signer` | `{"name": "นายสมชาย ใจดี", "position": ["ตำแหน่ง", "ปฏิบัติราชการแทน..."], "rank": "พันตำรวจเอก"}` — ชื่อจะถูกใส่วงเล็บให้อัตโนมัติ |
| `page_numbers` | `true` (ค่าเริ่มต้น) — เลขหน้า "- ๒ -" ตั้งแต่หน้า ๒ |
| `font`, `font_size` | ค่าเริ่มต้น `TH SarabunPSK`, `16` |

## ฟิลด์เฉพาะชนิด

| type | ฟิลด์ |
|---|---|
| `external` | `number` (ไม่ต้องใส่ "ที่"), `agency` (list บรรทัด: ชื่อส่วนราชการ + ที่ตั้ง), `date`, `subject`, `salutation` (รวมคำขึ้นต้น+ผู้รับ เช่น `"เรียน อธิบดีกรม..."`), `reference` (string หรือ list), `enclosures` (list), `body`, `closing`, `signer`, `contact` (list บรรทัด: ส่วนราชการเจ้าของเรื่อง, โทร., โทรสาร, อีเมล), `cc` (สำเนาส่ง) |
| `internal` | `department` (บรรทัดส่วนราชการ รวมเบอร์โทร), `number`, `date`, `subject`, `salutation`, `body`, `signer` (ไม่มี closing) |
| `stamped` | `number`, `to`, `body`, `issuer` (ชื่อส่วนราชการที่ส่งหนังสือออก), `date`, `contact` |
| `order` | `issuer` (→ "คำสั่ง{issuer}") หรือ `title` (ทั้งบรรทัด), `number` ("๑๒๕/๒๕๖๙"), `subject`, `body`, `effective` ("ทั้งนี้ ตั้งแต่..."), `date` (ใส่ "พ.ศ." ด้วย), `signer` |
| `regulation` / `bylaw` | `issuer` หรือ `title`, `topic` (ว่าด้วย…), `edition` ("ฉบับที่ ๒", ถ้ามี), `year`, `preamble` (list), `articles` (list ของ string "ข้อ ๑ …" หรือ `{"chapter": "หมวด ๑", "chapter_title": "บททั่วไป"}`), `date`, `signer` |
| `announcement` | `issuer` หรือ `title`, `subject`, `body`, `date`, `signer`, `word` (`"แจ้งความ"` ถ้ากฎหมายกำหนด) |
| `statement` / `news` | `issuer` หรือ `title`, `subject`, `edition` (ถ้ามี), `body`, `date` — ไม่มีลายมือชื่อ |
| `certificate` | `number` (→ "เลขที่ …"), `agency`, `body` (ขึ้นต้น "หนังสือฉบับนี้ให้ไว้เพื่อรับรองว่า…"), `date`, `signer`, `photo` (`true` หรือ `{"name": "ชื่อผู้ได้รับการรับรอง"}` → กล่องรูป ๔×๖ ซม.) |
| `minutes` | `meeting` หรือ `title`, `session`, `when`, `where`, `present` / `absent` / `attendees` (list ของ `{"name", "role", "note"}`), `start_time`, `body` (ย่อหน้าเปิด), `agenda` (list ของ `{"title", "body", "resolution"}`), `end_time`, `recorder` (`{"name", "position"}`), `checker` (ถ้ามี) |

## ข้อจำกัดที่ควรรู้

- สคริปต์จัดรูปแบบตามแบบท้ายระเบียบ แต่ไม่ได้ตรวจเนื้อหา — ตรวจด้วย `scripts/check_draft.py spec.json` ก่อน
- ภาพตัวอย่าง PDF จาก LibreOffice อาจตัดคำไทยต่างจาก Microsoft Word เล็กน้อย; ไฟล์ docx ตั้งค่าภาษาไทยไว้ให้ Word ตัดคำถูกต้อง
- ตราชื่อส่วนราชการ (หนังสือประทับตรา) และลายมือชื่อ เว้นที่ว่างไว้ให้ประทับ/ลงนามจริง
