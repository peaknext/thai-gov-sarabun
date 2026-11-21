#!/usr/bin/env python3
"""ตั้งค่าไฟล์ .docx ที่มีข้อความภาษาไทยให้ Microsoft Word ตัดคำและตรวจตัวสะกดภาษาไทยถูกต้อง

Usage:
    python3 thai_docx_fix.py input.docx [output.docx] [--font "TH SarabunPSK"]

(ถ้าไม่ระบุ output จะเขียนทับไฟล์เดิม) ใช้ได้กับ .docx ที่สร้างจากเครื่องมือใดก็ได้ (python-docx, docx-js, ฯลฯ)
build_docx.py เรียกสคริปต์นี้ให้อัตโนมัติทุกครั้ง

กฎที่บังคับใช้ (วิธีที่ผู้ใช้ยืนยันแล้วว่าใช้ได้กับ Word, 2026-11-19):
 1. ภาษาใน docDefaults และ **ทุก style**: <w:lang w:val="en-US" w:eastAsia="en-US" w:bidi="th-TH"/>
    — ภาษาไทยอยู่ใน bidi (complex script) เท่านั้น ห้ามใส่ th-TH ใน eastAsia (Word จะตัดบรรทัดภาษาไทยผิด)
 2. ทุก run ที่มีอักษรไทย (U+0E00–U+0E7F) ต้องมี <w:rFonts w:hint="cs"/> และ <w:cs/> ใน rPr
    (<w:cs/> อยู่ก่อน <w:em>/<w:lang> ตามลำดับของ schema) — สคริปต์นี้จัดลำดับลูกของ rPr/pPr/tblPr ตาม schema
    หมายเหตุ: ตัวตรวจ XSD ของ ISO 29500 จะแจ้งว่า w:hint="cs" ไม่อยู่ใน enumeration — เป็นค่าที่ Word เองเขียนและอ่านได้
    (ผู้ใช้ยืนยันแล้ว) ให้ถือว่าเป็น false positive
 3. settings.xml: <w:themeFontLang w:val="en-US" w:eastAsia="en-US" w:bidi="th-TH"/> ต่อท้าย </w:compat> ทันที
    (ถ้าไฟล์มี docVars/rsids/mathPr ซึ่ง schema กำหนดให้อยู่ระหว่าง compat กับ themeFontLang จะวางต่อจากตัวสุดท้ายของกลุ่มนั้น)
 4. ฟอนต์ TH SarabunPSK (ascii/hAnsi/cs/eastAsia) ตั้งทั้ง sz และ szCs และใส่ bCs คู่กับตัวหนา (iCs คู่กับตัวเอียง)
"""
from __future__ import annotations

import argparse
import re
import sys
import zipfile

from lxml import etree

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NS = {"w": W}


def q(tag: str) -> str:
    return f"{{{W}}}{tag}"


THAI = re.compile(r"[฀-๿]")
LANG = {"val": "en-US", "eastAsia": "en-US", "bidi": "th-TH"}

# CT_RPr child order (ECMA-376 Part 1, EG_RPrBase) — paragraph-mark change-tracking elements first
RPR_ORDER = [
    "ins", "del", "moveFrom", "moveTo",
    "rStyle", "rFonts", "b", "bCs", "i", "iCs", "caps", "smallCaps", "strike", "dstrike", "outline",
    "shadow", "emboss", "imprint", "noProof", "snapToGrid", "vanish", "webHidden", "color", "spacing",
    "w", "kern", "position", "sz", "szCs", "highlight", "u", "effect", "bdr", "shd", "fitText",
    "vertAlign", "rtl", "cs", "em", "lang", "eastAsianLayout", "specVanish", "oMath", "rPrChange",
]
RPR_RANK = {t: i for i, t in enumerate(RPR_ORDER)}
# CT_Style child order
STYLE_ORDER = [
    "name", "aliases", "basedOn", "next", "link", "autoRedefine", "hidden", "uiPriority", "semiHidden",
    "unhideWhenUsed", "qFormat", "locked", "personal", "personalCompose", "personalReply", "rsid",
    "pPr", "rPr", "tblPr", "trPr", "tcPr", "tblStylePr",
]
STYLE_RANK = {t: i for i, t in enumerate(STYLE_ORDER)}
# CT_PPr / CT_TblPr child order — generators often append these out of order (Word tolerates it, validators don't)
PPR_ORDER = [
    "pStyle", "keepNext", "keepLines", "pageBreakBefore", "framePr", "widowControl", "numPr",
    "suppressLineNumbers", "pBdr", "shd", "tabs", "suppressAutoHyphens", "kinsoku", "wordWrap",
    "overflowPunct", "topLinePunct", "autoSpaceDE", "autoSpaceDN", "bidi", "adjustRightInd", "snapToGrid",
    "spacing", "ind", "contextualSpacing", "mirrorIndents", "suppressOverlap", "jc", "textDirection",
    "textAlignment", "textboxTightWrap", "outlineLvl", "divId", "cnfStyle", "rPr", "sectPr", "pPrChange",
]
PPR_RANK = {t: i for i, t in enumerate(PPR_ORDER)}
TBLPR_ORDER = [
    "tblStyle", "tblpPr", "tblOverlap", "bidiVisual", "tblStyleRowBandSize", "tblStyleColBandSize", "tblW",
    "jc", "tblCellSpacing", "tblInd", "tblBorders", "shd", "tblLayout", "tblCellMar", "tblLook",
    "tblCaption", "tblDescription", "tblPrChange",
]
TBLPR_RANK = {t: i for i, t in enumerate(TBLPR_ORDER)}
# settings.xml elements that the schema puts between w:compat and w:themeFontLang
AFTER_COMPAT = ("compat", "docVars", "rsids", "mathPr", "attachedSchema")


def local(el) -> str:
    return etree.QName(el).localname


def sort_children(parent, rank):
    kids = list(parent)
    kids.sort(key=lambda e: rank.get(local(e), len(rank)) if isinstance(e.tag, str) else len(rank))
    for k in kids:
        parent.remove(k)
    for k in kids:
        parent.append(k)


def child(parent, tag, create=True):
    el = parent.find(q(tag))
    if el is None and create:
        el = etree.SubElement(parent, q(tag))
    return el


def set_lang(rpr):
    lang = child(rpr, "lang")
    for k, v in LANG.items():
        lang.set(q(k), v)


def set_fonts(rpr, font):
    if not font:
        return
    rf = child(rpr, "rFonts")
    for k in ("ascii", "hAnsi", "cs", "eastAsia"):
        rf.set(q(k), font)
    for k in ("asciiTheme", "hAnsiTheme", "cstheme", "eastAsiaTheme"):
        rf.attrib.pop(q(k), None)


def on(el) -> bool:
    """True if a toggle property like <w:b/> is switched on."""
    if el is None:
        return False
    v = el.get(q("val"))
    return v is None or v.lower() not in ("0", "false", "off")


def complete_pairs(rpr):
    """sz -> szCs, b -> bCs, i -> iCs (complex-script twins)."""
    sz = rpr.find(q("sz"))
    if sz is not None and rpr.find(q("szCs")) is None:
        etree.SubElement(rpr, q("szCs")).set(q("val"), sz.get(q("val")))
    for a, b in (("b", "bCs"), ("i", "iCs")):
        if on(rpr.find(q(a))) and rpr.find(q(b)) is None:
            etree.SubElement(rpr, q(b))


def fix_run_part(xml: bytes, font: str | None) -> bytes:
    root = etree.fromstring(xml)
    for r in root.iter(q("r")):
        text = "".join(t.text or "" for t in r.iter(q("t")))
        rpr = r.find(q("rPr"))
        if THAI.search(text):
            if rpr is None:
                rpr = etree.Element(q("rPr"))
                r.insert(0, rpr)
            rf = child(rpr, "rFonts")
            rf.set(q("hint"), "cs")
            if font and not rf.get(q("cs")):
                set_fonts(rpr, font)
            child(rpr, "cs")
        if rpr is not None:
            complete_pairs(rpr)
            for lang in rpr.findall(q("lang")):
                if (lang.get(q("eastAsia")) or "").lower().startswith("th"):
                    lang.set(q("eastAsia"), "en-US")
                    lang.set(q("bidi"), "th-TH")
            sort_children(rpr, RPR_RANK)
    # paragraph-mark run properties
    for rpr in root.iter(q("rPr")):
        if rpr.getparent() is not None and local(rpr.getparent()) == "pPr":
            complete_pairs(rpr)
            sort_children(rpr, RPR_RANK)
    for ppr in root.iter(q("pPr")):
        sort_children(ppr, PPR_RANK)
    for tblpr in root.iter(q("tblPr")):
        sort_children(tblpr, TBLPR_RANK)
    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)


def fix_styles(xml: bytes, font: str | None) -> bytes:
    root = etree.fromstring(xml)
    dd = root.find(q("docDefaults"))
    if dd is None:
        dd = etree.Element(q("docDefaults"))
        root.insert(0, dd)
    rprd = child(dd, "rPrDefault")
    rpr = child(rprd, "rPr")
    set_lang(rpr)
    set_fonts(rpr, font)
    if font and rpr.find(q("sz")) is None:
        etree.SubElement(rpr, q("sz")).set(q("val"), "32")
    complete_pairs(rpr)
    sort_children(rpr, RPR_RANK)
    for st in root.findall(q("style")):
        if st.get(q("type")) == "numbering":  # numbering styles cannot carry rPr
            continue
        srpr = child(st, "rPr")
        set_lang(srpr)
        complete_pairs(srpr)
        sort_children(srpr, RPR_RANK)
        sort_children(st, STYLE_RANK)
    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)


def fix_settings(xml: bytes) -> bytes:
    root = etree.fromstring(xml)
    for old in root.findall(q("themeFontLang")):
        root.remove(old)
    el = etree.Element(q("themeFontLang"))
    for k, v in LANG.items():
        el.set(q(k), v)
    # Right after </w:compat>; if the file also has docVars/rsids/mathPr/attachedSchema (which the schema
    # orders between compat and themeFontLang) go after the last of those so the file stays schema-valid.
    anchors = [e for e in root if isinstance(e.tag, str) and local(e) in AFTER_COMPAT]
    if anchors:
        anchors[-1].addnext(el)
    else:
        root.append(el)
    zoom = root.find(q("zoom"))
    if zoom is not None and zoom.get(q("percent")) is None:
        zoom.set(q("percent"), "100")
    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)


RUN_PARTS = re.compile(r"^word/(document|header\d*|footer\d*|footnotes|endnotes|comments)\.xml$")


def fix_docx(src: str, dst: str | None = None, font: str | None = "TH SarabunPSK") -> str:
    dst = dst or src
    with zipfile.ZipFile(src) as zin:
        items = [(i, zin.read(i.filename)) for i in zin.infolist()]
    out = []
    for info, data in items:
        name = info.filename
        if RUN_PARTS.match(name):
            data = fix_run_part(data, font)
        elif name == "word/styles.xml":
            data = fix_styles(data, font)
        elif name == "word/settings.xml":
            data = fix_settings(data)
        out.append((info, data))
    with zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as zout:
        for info, data in out:
            if not info.filename.endswith("/"):
                zout.writestr(info, data)
    return dst


def verify(path: str) -> list[str]:
    """Return a list of rule violations (empty list = all four rules satisfied)."""
    problems = []
    with zipfile.ZipFile(path) as z:
        names = z.namelist()
        styles = etree.fromstring(z.read("word/styles.xml"))
        settings = etree.fromstring(z.read("word/settings.xml"))
        parts = {n: etree.fromstring(z.read(n)) for n in names if RUN_PARTS.match(n)}
    for lang in styles.iter(q("lang")):
        if lang.get(q("bidi")) != "th-TH" or (lang.get(q("eastAsia")) or "").startswith("th"):
            problems.append("styles.xml: w:lang without bidi=th-TH or with th in eastAsia")
            break
    dd = styles.find("w:docDefaults/w:rPrDefault/w:rPr/w:lang", NS)
    if dd is None:
        problems.append("docDefaults has no w:lang")
    for st in styles.findall(q("style")):
        if st.get(q("type")) != "numbering" and st.find("w:rPr/w:lang", NS) is None:
            problems.append(f"style {st.get(q('styleId'))} has no w:lang")
            break
    tfl = settings.find(q("themeFontLang"))
    if tfl is None or tfl.get(q("bidi")) != "th-TH":
        problems.append("settings.xml: themeFontLang missing or bidi != th-TH")
    else:
        prev = tfl.getprevious()
        if settings.find(q("compat")) is not None and (prev is None or local(prev) not in AFTER_COMPAT):
            problems.append("settings.xml: themeFontLang is not right after w:compat (or rsids/mathPr)")
    for n, root in parts.items():
        for r in root.iter(q("r")):
            text = "".join(t.text or "" for t in r.iter(q("t")))
            rpr = r.find(q("rPr"))
            if THAI.search(text):
                if rpr is None or rpr.find(q("cs")) is None or (rpr.find(q("rFonts")) is None) \
                        or rpr.find(q("rFonts")).get(q("hint")) != "cs":
                    problems.append(f"{n}: Thai run without rFonts hint=cs + w:cs: {text[:20]!r}")
                    break
            if rpr is not None:
                order = [RPR_RANK.get(local(e), 999) for e in rpr if isinstance(e.tag, str)]
                if order != sorted(order):
                    problems.append(f"{n}: rPr children out of schema order: {[local(e) for e in rpr]}")
                    break
                if on(rpr.find(q("b"))) and rpr.find(q("bCs")) is None:
                    problems.append(f"{n}: bold run without bCs")
                    break
                if rpr.find(q("sz")) is not None and rpr.find(q("szCs")) is None:
                    problems.append(f"{n}: sz without szCs")
                    break
    return problems


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("input")
    ap.add_argument("output", nargs="?")
    ap.add_argument("--font", default="TH SarabunPSK",
                    help='font for docDefaults and Thai runs without a cs font (use "" to leave fonts alone)')
    ap.add_argument("--check", action="store_true", help="only report rule violations, do not modify")
    a = ap.parse_args()
    if a.check:
        probs = verify(a.input)
        print("\n".join(probs) if probs else "OK: all Thai Word rules satisfied")
        sys.exit(1 if probs else 0)
    out = fix_docx(a.input, a.output, a.font or None)
    probs = verify(out)
    print(f"wrote {out}" + ("" if not probs else "\nremaining issues:\n" + "\n".join(probs)))


if __name__ == "__main__":
    main()
