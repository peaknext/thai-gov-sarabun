# กฎการตั้งค่า .docx ภาษาไทยให้ Microsoft Word ตัดคำและตรวจตัวสะกดถูกต้อง

**ใช้ทุกครั้งที่สร้างไฟล์ Word (.docx) ที่มีข้อความภาษาไทย** — ไม่ว่าจะสร้างด้วย `scripts/build_docx.py`, python-docx เขียนเอง,
docx-js, หรือแก้ไฟล์ที่ผู้ใช้ส่งมา วิธีนี้ผู้ใช้ทดสอบกับ Microsoft Word แล้วว่าใช้ได้ (2026-11-19)

ถ้าตั้งค่าผิด Word จะ **ตัดบรรทัดกลางคำภาษาไทย** และขีดเส้นแดงตรวจตัวสะกดผิด ๆ ทั้งที่ไฟล์ดูปกติใน LibreOffice
(ภาพจาก LibreOffice จึงพิสูจน์ไม่ได้ว่า Word จะตัดคำถูก)

## กฎ ๔ ข้อ

1. **ภาษา** — ใน `docDefaults` และ **ทุก style** ใน `styles.xml`:
   ```xml
   <w:lang w:val="en-US" w:eastAsia="en-US" w:bidi="th-TH"/>
   ```
   ภาษาไทยอยู่ใน `bidi` (complex script) **เท่านั้น** · **ห้าม** ใส่ `th-TH` ใน `eastAsia` เด็ดขาด — Word จะใช้กฎตัดบรรทัดแบบเอเชียตะวันออกและตัดคำไทยผิดที่
   (docx-js: `language: { value: "en-US", eastAsia: "en-US", bidirectional: "th-TH" }`)

2. **ทุก run ที่มีอักษรไทย (U+0E00–U+0E7F)** ต้องมีใน `rPr`:
   ```xml
   <w:rFonts w:hint="cs" .../>  ...  <w:cs/>
   ```
   `<w:cs/>` ต้องอยู่ **ก่อน** `<w:em>`/`<w:lang>` ตามลำดับ schema (ลำดับลูกของ rPr: rStyle, rFonts, b, bCs, i, iCs, … color, … sz, szCs, … u, … rtl, **cs**, em, lang …)
   docx-js และ python-docx ไม่ใส่ให้ → ต้อง post-process `document.xml` (และ header/footer) ด้วย lxml

3. **settings.xml** — เพิ่มต่อจาก `</w:compat>` ทันที:
   ```xml
   <w:themeFontLang w:val="en-US" w:eastAsia="en-US" w:bidi="th-TH"/>
   ```
   (ถ้าไฟล์มี `w:docVars`/`w:rsids`/`m:mathPr` ซึ่ง schema กำหนดให้อยู่ระหว่าง compat กับ themeFontLang ให้วางต่อจากตัวสุดท้ายของกลุ่มนั้น —
   แม่แบบของ python-docx มี themeFontLang ที่ eastAsia เป็นภาษาจีนอยู่แล้ว ต้อง **แทนที่** ไม่ใช่ข้าม)

4. **ฟอนต์และขนาด** — TH SarabunPSK (หรือ TH Sarabun New) ที่ `ascii`/`hAnsi`/`cs`/`eastAsia` ครบทั้ง ๔ ช่อง ·
   ตั้งทั้ง `w:sz` **และ** `w:szCs` (๑๖ pt = `32`) · ตัวหนาใส่ `w:b` **และ** `w:bCs` (ตัวเอียง `w:i` + `w:iCs`) —
   เพราะ run ที่มี `<w:cs/>` Word จะใช้ค่าฝั่ง complex script (szCs, bCs, ฟอนต์ cs)

## วิธีทำ

- `scripts/build_docx.py` เรียก `scripts/thai_docx_fix.py` ให้อัตโนมัติหลังบันทึกไฟล์ และพิมพ์ปัญหาที่ยังเหลือ (ถ้ามี)
- ไฟล์ที่สร้างด้วยวิธีอื่น / ไฟล์ของผู้ใช้:
  ```bash
  python3 scripts/thai_docx_fix.py input.docx output.docx --font "TH SarabunPSK"   # แก้
  python3 scripts/thai_docx_fix.py --check output.docx                              # ตรวจอย่างเดียว
  ```
  สคริปต์จัดลำดับลูกของ rPr/pPr/tblPr ตาม schema, เติม szCs/bCs/iCs ที่ขาด, แก้ eastAsia="th-…" ที่ผิด และใส่ lang ให้ทุก style
- ส่งไฟล์ให้ผู้ใช้แล้ว แนะนำให้เปิดใน Microsoft Word เพื่อยืนยันการตัดคำ (LibreOffice ใช้ตรวจหน้าตาได้ แต่ไม่ใช่ตัวพิสูจน์การตัดคำของ Word)

## ข้อสังเกตเรื่องตัวตรวจ schema

ตัวตรวจ XSD ของ ISO/IEC 29500 (เช่น `validate.py` ของสกิล docx) จะแจ้งว่า `w:hint="cs"` ไม่อยู่ใน enumeration (`default`, `eastAsia`)
ค่า `cs` เป็นค่าที่ Word เขียนและอ่านเองตามปกติ และผู้ใช้ยืนยันแล้วว่าทำให้ตัดคำถูก — **ให้คงไว้** และถือข้อความเตือนนี้เป็น false positive
ข้อผิดพลาดอื่นจากตัวตรวจควรแก้ให้หมด
