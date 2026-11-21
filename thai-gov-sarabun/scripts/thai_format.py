#!/usr/bin/env python3
"""Thai official-document formatting helpers (เลขไทย, วันที่แบบหนังสือราชการ).

ใช้เป็นโมดูล:
    from thai_format import thai_digits, thai_date
    thai_digits("ที่ ศธ 0201/123")      -> "ที่ ศธ ๐๒๐๑/๑๒๓"
    thai_date("2026-11-21")             -> "๒๑ พฤศจิกายน ๒๕๖๙"
    thai_date("2026-11-21", short=True) -> "๒๑ พ.ย. ๖๙"

ใช้จากบรรทัดคำสั่ง:
    python3 thai_format.py digits "ข้อความ 123"
    python3 thai_format.py date 2026-11-21 [--arabic] [--short]
    python3 thai_format.py today
    python3 thai_format.py weekday 2569-12-05     -> วันเสาร์ — วันหยุด: เสาร์–อาทิตย์, วันพ่อแห่งชาติ
"""
from __future__ import annotations

import datetime as _dt
import re
import sys

ARABIC = "0123456789"
THAI = "๐๑๒๓๔๕๖๗๘๙"
_TO_THAI = str.maketrans(ARABIC, THAI)
_TO_ARABIC = str.maketrans(THAI, ARABIC)

MONTHS = [
    "มกราคม", "กุมภาพันธ์", "มีนาคม", "เมษายน", "พฤษภาคม", "มิถุนายน",
    "กรกฎาคม", "สิงหาคม", "กันยายน", "ตุลาคม", "พฤศจิกายน", "ธันวาคม",
]
MONTHS_SHORT = [
    "ม.ค.", "ก.พ.", "มี.ค.", "เม.ย.", "พ.ค.", "มิ.ย.",
    "ก.ค.", "ส.ค.", "ก.ย.", "ต.ค.", "พ.ย.", "ธ.ค.",
]

WEEKDAYS = ["วันจันทร์", "วันอังคาร", "วันพุธ", "วันพฤหัสบดี", "วันศุกร์", "วันเสาร์", "วันอาทิตย์"]
# วันหยุดราชการที่วันที่ตายตัว (ไม่รวมวันหยุดทางพุทธศาสนาที่เลื่อนตามจันทรคติ และวันหยุดพิเศษตามมติ ครม.)
FIXED_HOLIDAYS = {
    (1, 1): "วันขึ้นปีใหม่", (4, 6): "วันจักรี", (4, 13): "วันสงกรานต์", (4, 14): "วันสงกรานต์",
    (4, 15): "วันสงกรานต์", (5, 1): "วันแรงงานแห่งชาติ", (5, 4): "วันฉัตรมงคล",
    (6, 3): "วันเฉลิมพระชนมพรรษาสมเด็จพระราชินี", (7, 28): "วันเฉลิมพระชนมพรรษาพระบาทสมเด็จพระเจ้าอยู่หัว",
    (8, 12): "วันแม่แห่งชาติ", (10, 13): "วันนวมินทรมหาราช", (10, 23): "วันปิยมหาราช",
    (12, 5): "วันพ่อแห่งชาติ", (12, 10): "วันรัฐธรรมนูญ", (12, 31): "วันสิ้นปี",
}

# Text that must keep Arabic digits even inside a Thai document
# (URLs, e-mail addresses, file names such as 2564_OPM0913_56.pdf).
_KEEP = re.compile(r"(https?://\S+|[\w.+-]+@[\w.-]+|\b\d{4}_[A-Z]+\d*_\d+(?:_\d+)?\.pdf\b)")


def thai_digits(text: str) -> str:
    """Convert Arabic digits to Thai digits, leaving URLs/e-mails/file names alone."""
    parts = _KEEP.split(text)
    return "".join(p if i % 2 else p.translate(_TO_THAI) for i, p in enumerate(parts))


def arabic_digits(text: str) -> str:
    return text.translate(_TO_ARABIC)


def be_year(year: int) -> int:
    """ค.ศ. -> พ.ศ. (ถ้าให้ปีที่เป็น พ.ศ. อยู่แล้ว (>2400) คืนค่าเดิม)."""
    return year if year > 2400 else year + 543


def _parse(d) -> _dt.date:
    if isinstance(d, _dt.datetime):
        return d.date()
    if isinstance(d, _dt.date):
        return d
    s = arabic_digits(str(d)).strip()
    y, m, dd = (int(x) for x in re.split(r"[-/.]", s)[:3])
    if y > 2400:  # given in พ.ศ.
        y -= 543
    return _dt.date(y, m, dd)


def thai_date(d, *, arabic: bool = False, short: bool = False, with_be_prefix: bool = False) -> str:
    """Format a date the way หนังสือราชการ does: '๒๑ พฤศจิกายน ๒๕๖๙'.

    - ระเบียบข้อ ๑๑.๓: "ลงตัวเลขของวันที่ ชื่อเต็มของเดือน และตัวเลขของปีพุทธศักราช"
    - with_be_prefix=True gives '๒๑ พฤศจิกายน พ.ศ. ๒๕๖๙' (ใช้ใน "สั่ง ณ วันที่ ..." / "ประกาศ ณ วันที่ ...")
    - short=True gives '๒๑ พ.ย. ๖๙' (ใช้ในบันทึกต่อเนื่อง/การลงวันที่กำกับลายมือชื่อเท่านั้น)
    """
    dt = _parse(d)
    y = be_year(dt.year)
    if short:
        s = f"{dt.day} {MONTHS_SHORT[dt.month - 1]} {str(y)[-2:]}"
    elif with_be_prefix:
        s = f"{dt.day} {MONTHS[dt.month - 1]} พ.ศ. {y}"
    else:
        s = f"{dt.day} {MONTHS[dt.month - 1]} {y}"
    return s if arabic else thai_digits(s)


def weekday_info(d) -> str:
    """'วันเสาร์ — วันหยุด: วันพ่อแห่งชาติ' — ใช้ตรวจว่าวันกำหนด (กำหนดส่ง วันประชุม วันงาน) ตรงวันหยุดหรือไม่."""
    dt = _parse(d)
    info = WEEKDAYS[dt.weekday()]
    notes = []
    if dt.weekday() >= 5:
        notes.append("เสาร์–อาทิตย์")
    if (dt.month, dt.day) in FIXED_HOLIDAYS:
        notes.append(FIXED_HOLIDAYS[(dt.month, dt.day)])
    return info + (" — วันหยุด: " + ", ".join(notes) if notes else "")


def find_dates(text: str):
    """หาวันที่รูปแบบ '๕ ธันวาคม ๒๕๖๙' ในข้อความ คืน list ของ (ข้อความ, date)."""
    t = arabic_digits(text)
    out = []
    for m in re.finditer(r"(\d{1,2})\s*(" + "|".join(MONTHS) + r")\s*(?:พ\.ศ\.\s*)?(\d{4})", t):
        try:
            out.append((m.group(0), _parse(f"{m.group(3)}-{MONTHS.index(m.group(2)) + 1}-{m.group(1)}"), m.start()))
        except ValueError:
            pass
    return out


def main(argv: list[str]) -> int:
    if not argv or argv[0] in {"-h", "--help"}:
        print(__doc__)
        return 0
    cmd, rest = argv[0], argv[1:]
    flags = {a for a in rest if a.startswith("--")}
    args = [a for a in rest if not a.startswith("--")]
    if cmd == "digits":
        print(thai_digits(" ".join(args)))
    elif cmd == "date":
        print(thai_date(args[0], arabic="--arabic" in flags, short="--short" in flags,
                        with_be_prefix="--be" in flags))
    elif cmd == "weekday":
        print(weekday_info(args[0]))
    elif cmd == "today":
        print(thai_date(_dt.date.today(), arabic="--arabic" in flags, short="--short" in flags,
                        with_be_prefix="--be" in flags))
    else:
        print(f"unknown command: {cmd}", file=sys.stderr)
        return 2
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
