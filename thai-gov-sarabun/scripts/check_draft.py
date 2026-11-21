#!/usr/bin/env python3
"""ตรวจร่างหนังสือราชการ (spec JSON ของ build_docx.py) ตามระเบียบงานสารบรรณ

Usage:
    python3 check_draft.py spec.json [more.json ...]

พิมพ์ผลเป็น ERROR (ผิดระเบียบ ต้องแก้) / WARN (ควรตรวจ) / INFO (ข้อสังเกต)
exit code 1 ถ้ามี ERROR
"""
from __future__ import annotations

import json
import re
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
from build_docx import TYPE_ALIASES  # noqa: E402
from thai_format import MONTHS, arabic_digits  # noqa: E402

# ภาคผนวก ๒ ข้อ ๓.๑ (แก้ไขตามหนังสือ นร ๐๑๐๖/ว ๙๑๓ ลงวันที่ ๘ พ.ค. ๒๕๕๕) — ใช้ "กราบเรียน"
KRAB_RIAN = [
    "ประธานองคมนตรี", "นายกรัฐมนตรี", "ประธานรัฐสภา", "ประธานสภาผู้แทนราษฎร", "ประธานวุฒิสภา",
    "ประธานศาลฎีกา", "ประธานศาลรัฐธรรมนูญ", "ประธานศาลปกครองสูงสุด", "ประธานกรรมการการเลือกตั้ง",
    "ประธานกรรมการสิทธิมนุษยชนแห่งชาติ", "ประธานกรรมการป้องกันและปราบปรามการทุจริตแห่งชาติ",
    "ประธานกรรมการตรวจเงินแผ่นดิน", "ประธานผู้ตรวจการแผ่นดิน", "ผู้ตรวจการแผ่นดิน", "อัยการสูงสุด", "รัฐบุรุษ",
]
PAIRS = {  # คำขึ้นต้น -> คำลงท้ายที่เข้าคู่ (ภาคผนวก ๒)
    "กราบเรียน": ["ขอแสดงความนับถืออย่างยิ่ง"],
    "เรียน": ["ขอแสดงความนับถือ"],
    "นมัสการ": ["ขอนมัสการด้วยความเคารพอย่างยิ่ง", "ขอนมัสการด้วยความเคารพอย่างสูง", "ขอนมัสการด้วยความเคารพ"],
    "กราบทูล": ["ควรมิควรแล้วแต่จะโปรด"],
    "ทูล": ["ควรมิควรแล้วแต่จะโปรด", "แล้วแต่จะโปรด"],
    "ขอประทานกราบทูล": ["ควรมิควรแล้วแต่จะโปรดเกล้าโปรดกระหม่อม"],
    "ขอพระราชทานกราบทูล": ["ควรมิควรแล้วแต่จะโปรดเกล้าโปรดกระหม่อม"],
    "ขอพระราชทานกราบบังคมทูล": ["ควรมิควรแล้วแต่จะทรงพระกรุณาโปรดเกล้าโปรดกระหม่อม"],
    "ขอเดชะ": ["ควรมิควรแล้วแต่จะทรงพระกรุณาโปรดเกล้าโปรดกระหม่อม ขอเดชะ"],
}
URGENCY = {"ด่วน", "ด่วนมาก", "ด่วนที่สุด"}
SECRECY = {"ลับ", "ลับมาก", "ลับที่สุด"}
COLLOQUIAL = ["ครับ", "ค่ะ", "คะ ", "นะคะ", "นะครับ", "เยอะ", "ไหม", "หรือเปล่า", "เช็ค", "อัพเดท", "อัปเดต",
              "โอเค", "ประมาณว่า", "ก็เลย", "แล้วก็", "!", "?", "เดี๋ยว", "ตอนนี้"]
NAME_ABBR = ["ดร.", "นพ.", "พญ.", "ผศ.", "รศ.", "ศ.", "ม.ร.ว.", "ม.ล.", "พ.ต.อ.", "พล.ต.อ."]
SUBST_WORDS = ["ปฏิบัติหน้าที่แทน", "รักษาราชการแทน", "รักษาการแทน", "ปฏิบัติราชการแทน", "รักษาการในตำแหน่ง",
               "ทำการแทน", "แทน"]

REQUIRED = {
    "external": ["number", "agency", "date", "subject", "salutation", "body", "closing", "signer", "contact"],
    "internal": ["department", "number", "date", "subject", "salutation", "body", "signer"],
    "stamped": ["number", "to", "body", "issuer", "date", "contact"],
    "order": ["number", "subject", "body", "date", "signer"],
    "regulation": ["topic", "year", "articles", "date", "signer"],
    "bylaw": ["topic", "year", "preamble", "articles", "date", "signer"],
    "announcement": ["subject", "body", "date", "signer"],
    "statement": ["subject", "body", "date"],
    "news": ["subject", "body", "date"],
    "certificate": ["number", "agency", "body", "date", "signer"],
    "minutes": ["session", "when", "where", "present", "body", "recorder"],
}
NEEDS_ISSUER = {"order", "regulation", "bylaw", "announcement", "statement", "news"}
DATE_FULL = re.compile(r"^\d{1,2} (" + "|".join(MONTHS) + r") \d{4}$")
DATE_BE = re.compile(r"^\d{1,2} (" + "|".join(MONTHS) + r") พ\.ศ\. \d{4}$")
NUM_LETTER = re.compile(r"^[ก-ฮ]{2}(\s*\([ก-ฮ]{1,4}\))?\s*\d{4,5}(\.\d{1,3})?/(ว\s*)?\d*$")
NUM_ORDER = re.compile(r"^\d+/\d{4}$")


class Report:
    def __init__(self, name):
        self.name = name
        self.items = []

    def add(self, level, msg):
        self.items.append((level, msg))

    def error(self, m):
        self.add("ERROR", m)

    def warn(self, m):
        self.add("WARN", m)

    def info(self, m):
        self.add("INFO", m)


def text_of(v) -> str:
    if v is None:
        return ""
    if isinstance(v, str):
        return v
    if isinstance(v, list):
        return "\n".join(text_of(x) for x in v)
    if isinstance(v, dict):
        return "\n".join(text_of(x) for x in v.values())
    return str(v)


def opener_is(sal: str, word: str) -> bool:
    return sal.startswith(word)


def num_raw(spec) -> str:
    return re.sub(r"\s+", "", arabic_digits(spec.get("number") or ""))


def check(spec: dict, rep: Report):
    kind = TYPE_ALIASES.get(spec.get("type", ""))
    if not kind:
        rep.error(f"ไม่รู้จักชนิดหนังสือ type={spec.get('type')!r}")
        return
    rep.info(f"ชนิด: {kind}")
    for f in REQUIRED[kind]:
        if not spec.get(f):
            rep.error(f"ขาดส่วน '{f}' ซึ่งต้องมีในแบบ {kind}")
    if kind in NEEDS_ISSUER and not (spec.get("issuer") or spec.get("title")):
        rep.error("ขาดชื่อส่วนราชการ/ตำแหน่งผู้ออก (issuer หรือ title)")
    if kind == "minutes" and not (spec.get("meeting") or spec.get("title")):
        rep.error("ขาดชื่อคณะที่ประชุม (meeting)")

    # ---- ชั้นความเร็ว / ความลับ
    u = spec.get("urgency")
    if u and u not in URGENCY:
        rep.error(f"ชั้นความเร็ว '{u}' ไม่ตรงระเบียบข้อ ๒๘ (ด่วนที่สุด / ด่วนมาก / ด่วน)")
    if u and kind not in {"external", "internal", "stamped"}:
        rep.warn("ระเบียบกำหนดชั้นความเร็วไว้ในแบบที่ ๑, ๒, ๓ (และซอง) เท่านั้น")
    s = spec.get("secrecy")
    if s and s not in SECRECY:
        rep.error(f"ชั้นความลับ '{s}' ไม่ถูกต้อง (ลับ / ลับมาก / ลับที่สุด)")
    if s == "ลับที่สุด":
        rep.warn("ชั้นลับที่สุด ห้ามรับส่งด้วยระบบสารบรรณอิเล็กทรอนิกส์/อีเมล (ข้อ ๒๙)")

    # ---- คำขึ้นต้น / คำลงท้าย
    sal = (spec.get("salutation") or "").strip()
    clo = (spec.get("closing") or "").strip()
    if sal:
        opener = next((k for k in sorted(PAIRS, key=len, reverse=True) if sal.startswith(k)), None)
        if not opener:
            rep.warn(f"คำขึ้นต้น '{sal[:20]}…' ไม่อยู่ในภาคผนวก ๒")
        else:
            recipient = sal[len(opener):].strip()
            if not recipient:
                rep.error("หลังคำขึ้นต้นต้องลงตำแหน่งหรือชื่อผู้รับ (ข้อ ๑๑.๕)")
            in_31 = any(t in recipient for t in KRAB_RIAN)
            if opener == "เรียน" and in_31:
                rep.error(f"ผู้รับ '{recipient}' อยู่ในภาคผนวก ๒ ข้อ ๓.๑ ต้องใช้ 'กราบเรียน' และ 'ขอแสดงความนับถืออย่างยิ่ง'")
            if opener == "กราบเรียน" and not in_31:
                rep.warn(f"'{recipient}' ไม่อยู่ในรายการ ๓.๑ ตามระเบียบใช้ 'เรียน' (ใช้ 'กราบเรียน' ได้เฉพาะถ้าเป็นธรรมเนียมของหน่วยงาน)")
            if kind == "external":
                if not clo:
                    rep.error("หนังสือภายนอกต้องมีคำลงท้าย")
                elif clo not in PAIRS[opener]:
                    rep.error(f"คำลงท้าย '{clo}' ไม่เข้าคู่กับคำขึ้นต้น '{opener}' (ควรเป็น: {' / '.join(PAIRS[opener])})")
    if opener_is(sal, "นมัสการ") and clo:
        r = sal[len("นมัสการ"):].strip()
        want = None
        if re.match(r"^(สมเด็จพระ|พระพรหม)", r):
            want = "ขอนมัสการด้วยความเคารพอย่างยิ่ง"   # สมเด็จพระราชาคณะ / รองสมเด็จ
        elif re.match(r"^(พระธรรม|พระเทพ|พระราช)", r) and not r.startswith("พระราชวัง"):
            want = "ขอนมัสการด้วยความเคารพอย่างสูง"   # พระราชาคณะ
        elif re.match(r"^(พระครู|พระมหา|พระใบฎีกา|พระสมุห์|พระปลัด|พระอธิการ|เจ้าอาวาส)", r):
            want = "ขอนมัสการด้วยความเคารพ"            # พระภิกษุสงฆ์ทั่วไป
        if want and clo != want:
            rep.warn(f"ผู้รับดูเหมือนเป็นพระชั้น '{r.split()[0]}' — คำลงท้ายตามภาคผนวก ๒ ข้อ ๒ ควรเป็น '{want}' (ตรวจสมณศักดิ์อีกครั้ง)")
        for bad, good in (("เชิญ", "อาราธนา/นิมนต์"), ("เงิน", "ปัจจัย"), ("รับประทาน", "ฉัน")):
            if bad in text_of(spec.get("body")) + (spec.get("subject") or ""):
                rep.warn(f"หนังสือถึงพระสงฆ์: พบคำ '{bad}' — ควรใช้ '{good}' (to-monks.md)")
    circular = bool(spec.get("circular") or spec.get("distribution"))
    if circular and kind in {"external", "internal", "stamped"} and "/ว" not in num_raw(spec):
        rep.error("หนังสือเวียนต้องเติม 'ว' หน้าเลขทะเบียนหนังสือส่ง เช่น 'สธ ๐๔๐๑/ว ๑๒๓' (ข้อ ๓๒)")
    if kind == "internal" and clo:
        rep.warn("หนังสือภายใน/บันทึกข้อความ (แบบที่ ๒) ไม่มีคำลงท้าย")
    if kind in {"stamped", "order", "regulation", "bylaw", "announcement", "statement", "news", "certificate"} and sal:
        rep.warn(f"แบบ {kind} ไม่มีคำขึ้นต้น")

    # ---- วันที่
    date = arabic_digits((spec.get("date") or "").strip())
    if date:
        if kind in {"order", "regulation", "bylaw", "announcement", "certificate"}:
            if not DATE_BE.match(date):
                rep.warn(f"วันที่ '{spec['date']}' ควรอยู่ในรูป 'สั่ง/ประกาศ/ให้ไว้ ณ วันที่ ๒๑ พฤศจิกายน พ.ศ. ๒๕๖๙' (มี พ.ศ.)")
        elif kind != "minutes" and not DATE_FULL.match(date):
            rep.error(f"วันที่ '{spec['date']}' ต้องเป็น ตัวเลขวันที่ + ชื่อเต็มของเดือน + ปี พ.ศ. เช่น '๒๑ พฤศจิกายน ๒๕๖๙'")
        y = re.search(r"(\d{4})$", date)
        if y and int(y.group(1)) < 2400:
            rep.error("ปีต้องเป็นพุทธศักราช ไม่ใช่คริสต์ศักราช")

    # ---- เลขที่
    num = arabic_digits((spec.get("number") or "").strip())
    if num.startswith("ที่ ") or num.startswith("เลขที่"):
        rep.warn("ไม่ต้องใส่คำว่า 'ที่'/'เลขที่' ในฟิลด์ number (สคริปต์ใส่ให้)")
        num = re.sub(r"^(เลขที่|ที่)\s*", "", num)
    if num:
        if kind in {"external", "internal", "stamped"} and not NUM_LETTER.match(num):
            rep.warn(f"เลขที่ '{spec['number']}' ไม่ตรงรูปแบบภาคผนวก ๑: รหัสพยัญชนะ ๒ ตัว + เลขประจำ ๔ หลัก / เลขทะเบียนส่ง เช่น 'ศธ ๐๒๐๑/๑๒๓'")
        if kind in {"external", "internal", "stamped"} and num.endswith("/"):
            rep.info("ยังไม่มีเลขทะเบียนหนังสือส่ง — หน่วยสารบรรณกลางจะออกเลขให้เมื่อส่ง")
        if kind == "order" and not NUM_ORDER.match(num):
            rep.warn(f"เลขที่คำสั่ง '{spec['number']}' ควรเป็น เลขลำดับ/ปี พ.ศ. เช่น '๑๒๕/๒๕๖๙' (ข้อ ๑๖.๒)")

    # ---- เนื้อความ
    body = text_of(spec.get("body")) + "\n" + text_of(spec.get("preamble")) + "\n" + text_of(spec.get("articles")) \
        + "\n" + text_of(spec.get("agenda"))
    for w in COLLOQUIAL:
        if w in body:
            rep.warn(f"พบคำ/เครื่องหมายแบบภาษาพูด '{w.strip()}' ในเนื้อความ")
    for para in (spec.get("body") or []):
        t = text_of(para)
        if t.startswith("ตามที่") and "นั้น" not in t:
            rep.warn("ย่อหน้าที่ขึ้นต้นด้วย 'ตามที่' ควรปิดอนุประโยคด้วย '… นั้น'")
    if spec.get("reference") and "อ้างถึง" not in body:
        rep.warn("มี 'อ้างถึง' แต่เนื้อความไม่ได้กล่าวถึงหนังสือที่อ้างถึง")
    if spec.get("enclosures") and "สิ่งที่ส่งมาด้วย" not in body:
        rep.warn("มี 'สิ่งที่ส่งมาด้วย' แต่เนื้อความไม่ได้กล่าวถึง")
    if kind in {"external", "internal"}:
        last = text_of((spec.get("body") or [""])[-1])
        if not re.search(r"^จึง|มาเพื่อ|จะขอบคุณ|ขอขอบคุณ", last):
            rep.warn("ย่อหน้าสุดท้ายควรเป็นส่วนสรุปความประสงค์ เช่น 'จึงเรียนมาเพื่อโปรดพิจารณา'")
    if re.search(r"\d", body) and re.search(r"[๐-๙]", body) and not spec.get("digits"):
        rep.info("เนื้อความมีทั้งเลขอารบิกและเลขไทย — build_docx จะแปลงเป็นเลขไทยทั้งหมดอัตโนมัติ")
    leftovers = sorted(set(re.findall(r"\[[^\]]{1,40}\]", text_of(spec))))
    if leftovers:
        rep.info("ยังมีช่องที่ต้องเติม: " + ", ".join(leftovers))

    # ---- ระเบียบ / ข้อบังคับ
    if kind in {"regulation", "bylaw"}:
        arts = [text_of(a) for a in spec.get("articles", []) if not (isinstance(a, dict) and a.get("chapter"))]
        if arts:
            if "เรียกว่า" not in arts[0]:
                rep.error("ข้อ ๑ ต้องเป็นชื่อระเบียบ/ข้อบังคับ ('…นี้เรียกว่า “…”') (ข้อ ๑๗.๖/๑๘.๖)")
            if len(arts) > 1 and "ใช้บังคับ" not in arts[1]:
                rep.error("ข้อ ๒ ต้องเป็นวันใช้บังคับ")
            if "รักษาการ" not in arts[-1] and not any(isinstance(a, dict) and a.get("chapter") for a in spec["articles"]):
                rep.error("ข้อสุดท้ายต้องเป็นข้อผู้รักษาการ")
        if kind == "bylaw" and "อาศัยอำนาจ" not in text_of(spec.get("preamble")):
            rep.error("ข้อบังคับต้องอ้างกฎหมายที่ให้อำนาจออก (ข้อ ๑๘.๕)")
    if kind == "order":
        if not spec.get("effective") and "ตั้งแต่" not in body:
            rep.warn("คำสั่งควรระบุวันใช้บังคับ ('ทั้งนี้ ตั้งแต่…') (ข้อ ๑๖.๔)")
    if kind == "certificate" and not text_of(spec.get("body")).startswith("หนังสือฉบับนี้ให้ไว้เพื่อรับรองว่า"):
        rep.error("หนังสือรับรองต้องขึ้นต้นข้อความว่า 'หนังสือฉบับนี้ให้ไว้เพื่อรับรองว่า' (ข้อ ๒๔.๓)")

    # ---- ผู้ลงชื่อ
    signer = spec.get("signer") or {}
    name = signer.get("name", "")
    for ab in NAME_ABBR:
        if ab in name:
            rep.warn(f"ชื่อใต้ลายมือชื่อมีคำย่อ '{ab}' — ภาคผนวก ๓ ให้พิมพ์คำนำนาม/ยศ/บรรดาศักดิ์เป็นคำเต็ม")
    if name and not name.startswith("[") and not re.match(r"^\(?(นาย|นาง|นางสาว|ท่านผู้หญิง|คุณหญิง|หม่อม|ศาสตราจารย์|รองศาสตราจารย์|ผู้ช่วยศาสตราจารย์|"
                             r"นายแพทย์|แพทย์หญิง|ทันตแพทย์|พระ|พล|พัน|ร้อย|จ่า|สิบ|พลตำรวจ|ว่าที่)", name) \
            and not signer.get("rank"):
        rep.warn("ชื่อเต็มใต้ลายมือชื่อควรมีคำนำนาม นาย/นาง/นางสาว (ภาคผนวก ๓ ข้อ ๓)")
    pos = signer.get("position") or []
    pos = [pos] if isinstance(pos, str) else pos
    for line in pos[1:]:
        if line.endswith("แทน") or ("แทน" in line and not any(line.startswith(w) for w in SUBST_WORDS)):
            rep.warn(f"บรรทัดตำแหน่ง '{line}' — คำลงชื่อแทนต้องตามด้วยตำแหน่งเจ้าของหนังสือ เช่น 'ปฏิบัติราชการแทนอธิบดี…'")


def main(argv):
    if not argv:
        print(__doc__)
        return 2
    bad = False
    for path in argv:
        spec = json.loads(Path(path).read_text(encoding="utf-8"))
        rep = Report(path)
        check(spec, rep)
        print(f"== {path}")
        for level, msg in rep.items:
            print(f"  {level:5} {msg}")
        if any(l == "ERROR" for l, _ in rep.items):
            bad = True
        else:
            print("  OK    ไม่พบข้อผิดพลาดตามระเบียบ (ยังควรอ่านทานเนื้อหาเอง)")
    return 1 if bad else 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
