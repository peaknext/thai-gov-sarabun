#!/usr/bin/env python3
"""สร้างไฟล์ Word (.docx) ของหนังสือราชการไทยตามแบบท้ายระเบียบสำนักนายกรัฐมนตรีว่าด้วยงานสารบรรณ

Usage:
    python3 build_docx.py spec.json output.docx [--pdf] [--font "TH Sarabun New"] [--arabic-digits]

spec.json อธิบายเนื้อหาหนังสือ (ดูรายละเอียดฟิลด์ใน references/docx-spec.md) ตัวอย่างย่อ:
    {
      "type": "external",
      "number": "ศธ ๐๒๐๑/๑๒๓",
      "agency": ["กระทรวงศึกษาธิการ", "ถนนราชดำเนินนอก กทม. ๑๐๓๐๐"],
      "date": "๒๑ พฤศจิกายน ๒๕๖๙",
      "subject": "ขอเชิญเป็นวิทยากร",
      "salutation": "เรียน อธิบดีกรมควบคุมโรค",
      "body": ["ด้วย...", "จึงเรียนมาเพื่อโปรดพิจารณา"],
      "closing": "ขอแสดงความนับถือ",
      "signer": {"name": "นายสมชาย ใจดี", "position": ["ปลัดกระทรวงศึกษาธิการ"]},
      "contact": ["สำนักงานปลัดกระทรวง", "โทร. ๐ ๒๖๒๘ ๕๕๕๕"]
    }

type ที่รองรับ: external (หนังสือภายนอก แบบที่ ๑), internal / memo (หนังสือภายใน, บันทึก — แบบที่ ๒/๒๙),
stamped (หนังสือประทับตรา แบบที่ ๓), order (คำสั่ง แบบที่ ๔), regulation (ระเบียบ แบบที่ ๕),
bylaw (ข้อบังคับ แบบที่ ๖), announcement (ประกาศ แบบที่ ๗), statement (แถลงการณ์ แบบที่ ๘),
news (ข่าว แบบที่ ๙), certificate (หนังสือรับรอง แบบที่ ๑๐), minutes (รายงานการประชุม แบบที่ ๑๑)

Layout values that the regulation itself fixes (Garuda 3 cm / 1.5 cm, A4, red urgency text >= 32 pt)
are followed exactly. Margins, font size and indents follow common government practice
(TH SarabunPSK 16 pt, left 3 cm, right 2 cm, paragraph indent 2.5 cm) — see references/layout-spec.md.
After saving, scripts/thai_docx_fix.py is applied so Word breaks and spell-checks Thai correctly
(lang bidi=th-TH in docDefaults and every style, w:cs + hint on Thai runs, themeFontLang, sz/szCs, bCs).
"""
from __future__ import annotations

import argparse
import json
import os
import re
import subprocess
import sys
from pathlib import Path

from docx import Document
from docx.enum.section import WD_SECTION
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_TAB_ALIGNMENT, WD_BREAK
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor

sys.path.insert(0, str(Path(__file__).resolve().parent))
from thai_format import thai_digits  # noqa: E402
from thai_docx_fix import fix_docx, verify  # noqa: E402

SKILL_DIR = Path(__file__).resolve().parent.parent
GARUDA = SKILL_DIR / "assets" / "garuda.png"

THAI_RE = re.compile(r"[฀-๿]")
RED = RGBColor(0xC0, 0x00, 0x00)

# ---- page geometry (cm) ----
PAGE_W, PAGE_H = 21.0, 29.7
LEFT, RIGHT = 3.0, 2.0
TEXT_W = PAGE_W - LEFT - RIGHT          # 16 cm
PARA_INDENT = 2.5                        # ย่อหน้า
SIGN_INDENT = 2.0                        # centred signature block: axis = (indent + TEXT_W)/2 ≈ 9 cm from left margin
DATE_INDENT = 7.5                        # วัน เดือน ปี starts just left of page centre (under the Garuda)

TYPE_ALIASES = {
    "external": "external", "หนังสือภายนอก": "external",
    "internal": "internal", "หนังสือภายใน": "internal", "memo": "internal", "บันทึก": "internal",
    "บันทึกข้อความ": "internal",
    "stamped": "stamped", "หนังสือประทับตรา": "stamped",
    "order": "order", "คำสั่ง": "order",
    "regulation": "regulation", "ระเบียบ": "regulation",
    "bylaw": "bylaw", "ข้อบังคับ": "bylaw",
    "announcement": "announcement", "ประกาศ": "announcement", "แจ้งความ": "announcement",
    "statement": "statement", "แถลงการณ์": "statement",
    "news": "news", "ข่าว": "news",
    "certificate": "certificate", "หนังสือรับรอง": "certificate",
    "minutes": "minutes", "รายงานการประชุม": "minutes",
}


class Builder:
    def __init__(self, spec: dict, font: str, arabic: bool):
        self.spec = spec
        self.font = spec.get("font", font)
        self.size = float(spec.get("font_size", 16))
        self.arabic = arabic or spec.get("digits") == "arabic"
        self.doc = Document()
        self._setup_document()

    # ------------------------------------------------------------------ helpers
    def t(self, s) -> str:
        if s is None:
            return ""
        s = str(s)
        return s if self.arabic else thai_digits(s)

    def _setup_document(self):
        doc = self.doc
        sec = doc.sections[0]
        sec.page_width, sec.page_height = Cm(PAGE_W), Cm(PAGE_H)
        sec.left_margin, sec.right_margin = Cm(LEFT), Cm(RIGHT)
        sec.bottom_margin = Cm(2.0)
        sec.header_distance, sec.footer_distance = Cm(1.0), Cm(1.0)
        # Normal style: Thai font everywhere + Thai only in bidi language slot
        st = doc.styles["Normal"]
        st.font.name = self.font
        st.font.size = Pt(self.size)
        rpr = st.element.get_or_add_rPr()
        self._set_fonts(rpr)
        self._set_size_cs(rpr, self.size)
        lang = rpr.find(qn("w:lang"))
        if lang is None:
            lang = OxmlElement("w:lang")
            rpr.append(lang)
        lang.set(qn("w:val"), "en-US")
        lang.set(qn("w:eastAsia"), "en-US")
        lang.set(qn("w:bidi"), "th-TH")
        pf = st.paragraph_format
        pf.space_before = Pt(0)
        pf.space_after = Pt(0)
        pf.line_spacing = 1.0
        # docDefaults language too
        styles_el = doc.styles.element
        rpr_default = styles_el.find(qn("w:docDefaults"))
        if rpr_default is not None:
            for l in rpr_default.iter(qn("w:lang")):
                l.set(qn("w:val"), "en-US")
                l.set(qn("w:eastAsia"), "en-US")
                l.set(qn("w:bidi"), "th-TH")

    def _set_fonts(self, rpr, hint_cs=False):
        rf = rpr.find(qn("w:rFonts"))
        if rf is None:
            rf = OxmlElement("w:rFonts")
            rpr.insert(0, rf)
        for k in ("w:ascii", "w:hAnsi", "w:cs", "w:eastAsia"):
            rf.set(qn(k), self.font)
        if hint_cs:
            rf.set(qn("w:hint"), "cs")

    @staticmethod
    def _set_size_cs(rpr, size):
        for tag in ("w:sz", "w:szCs"):
            el = rpr.find(qn(tag))
            if el is None:
                el = OxmlElement(tag)
                rpr.append(el)
            el.set(qn("w:val"), str(int(round(size * 2))))

    def run(self, p, text, bold=False, size=None, color=None, underline=False, italic=False):
        r = p.add_run(self.t(text))
        size = size or self.size
        r.font.size = Pt(size)
        r.font.name = self.font
        r.bold = bold or None
        r.italic = italic or None
        r.underline = underline or None
        if color is not None:
            r.font.color.rgb = color
        rpr = r._r.get_or_add_rPr()
        self._set_fonts(rpr, hint_cs=bool(THAI_RE.search(text or "")))
        self._set_size_cs(rpr, size)
        if bold:
            rpr.append(OxmlElement("w:bCs"))
        if italic:
            rpr.append(OxmlElement("w:iCs"))
        if THAI_RE.search(text or ""):
            rpr.append(OxmlElement("w:cs"))
        return r

    def para(self, text="", *, align=None, first=None, left=None, right=None, before=0, after=0,
             bold=False, size=None, color=None, keep_next=False, underline=False, italic=False):
        p = self.doc.add_paragraph()
        pf = p.paragraph_format
        if align == "center":
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        elif align == "right":
            p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        elif align == "justify":
            p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY  # Word uses Thai distributed justify for th
        if first is not None:
            pf.first_line_indent = Cm(first)
        if left is not None:
            pf.left_indent = Cm(left)
        if right is not None:
            pf.right_indent = Cm(right)
        pf.space_before = Pt(before)
        pf.space_after = Pt(after)
        if keep_next:
            pf.keep_with_next = True
        if text:
            self.run(p, text, bold=bold, size=size, color=color, underline=underline, italic=italic)
        return p

    def blank(self, n=1):
        for _ in range(n):
            self.para("")

    def labelled(self, label, value, *, before=6, tab=None, bold_label=False, label_size=None):
        """'เรื่อง  xxx' style line; continuation lines hang under the value."""
        p = self.para(before=before)
        if tab:
            pf = p.paragraph_format
            pf.left_indent = Cm(tab)
            pf.first_line_indent = Cm(-tab)
            pf.tab_stops.add_tab_stop(Cm(tab))
            self.run(p, label, bold=bold_label, size=label_size)
            self.run(p, "\t" + (value or ""))
        else:
            self.run(p, label, bold=bold_label, size=label_size)
            self.run(p, ("  " if value else "") + (value or ""))
        return p

    def short_rule(self, before=0, after=6):
        p = self.para(before=before, after=after, left=TEXT_W / 2 - 2.25, right=TEXT_W / 2 - 2.25)
        pPr = p._p.get_or_add_pPr()
        bdr = OxmlElement("w:pBdr")
        b = OxmlElement("w:bottom")
        b.set(qn("w:val"), "single")
        b.set(qn("w:sz"), "6")
        b.set(qn("w:space"), "1")
        b.set(qn("w:color"), "000000")
        bdr.append(b)
        pPr.append(bdr)
        return p

    def garuda(self, height_cm=3.0, align="center", before=0):
        p = self.para(align=align, before=before)
        if GARUDA.exists():
            p.add_run().add_picture(str(GARUDA), height=Cm(height_cm))
        else:
            self.run(p, f"[ตราครุฑ {height_cm} ซม.]")
        return p

    @staticmethod
    def _no_borders(table):
        tbl = table._tbl
        tblPr = tbl.tblPr
        borders = OxmlElement("w:tblBorders")
        for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
            el = OxmlElement(f"w:{edge}")
            el.set(qn("w:val"), "nil")
            borders.append(el)
        tblPr.append(borders)

    def cell_para(self, cell, text="", align=None, bold=False, size=None, color=None, first=True):
        p = cell.paragraphs[0] if first else cell.add_paragraph()
        if align == "center":
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        elif align == "right":
            p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        if text:
            self.run(p, text, bold=bold, size=size, color=color)
        return p

    # ---------------------------------------------------- header/footer marks
    def secrecy_and_pages(self):
        """ชั้นความลับ top & bottom centre of every page (red); page number '- ๒ -' from page 2."""
        sec = self.doc.sections[0]
        secrecy = self.spec.get("secrecy")
        sec.different_first_page_header_footer = True
        headers = [(sec.first_page_header, False), (sec.header, True)]
        for hdr, with_page in headers:
            p = hdr.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            if with_page and self.spec.get("page_numbers", True):
                self.run(p, "- ")
                self._page_field(p)
                self.run(p, " -")
            if secrecy:
                if with_page and self.spec.get("page_numbers", True):
                    p = hdr.add_paragraph()
                    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                self.run(p, secrecy, bold=True, size=self.spec.get("secrecy_size", 20), color=RED)
        if secrecy:
            for ftr in (sec.first_page_footer, sec.footer):
                p = ftr.paragraphs[0]
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                self.run(p, secrecy, bold=True, size=self.spec.get("secrecy_size", 20), color=RED)

    def _page_field(self, p):
        # Word shows Thai digits with the \* ThaiArabic switch
        fmt = "" if self.arabic else r" \* ThaiArabic"
        for kind, txt in (("begin", None), ("instr", f" PAGE{fmt} "), ("separate", None), ("text", "2"),
                          ("end", None)):
            r = p.add_run()
            r.font.size = Pt(self.size)
            if kind in ("begin", "separate", "end"):
                fc = OxmlElement("w:fldChar")
                fc.set(qn("w:fldCharType"), kind)
                r._r.append(fc)
            elif kind == "instr":
                it = OxmlElement("w:instrText")
                it.set(qn("xml:space"), "preserve")
                it.text = txt
                r._r.append(it)
            else:
                r.text = self.t(txt)

    # --------------------------------------------------------------- shared blocks
    def urgency_text(self):
        u = self.spec.get("urgency")
        if u and self.spec.get("urgency_deadline"):
            u = f"{u} {self.spec['urgency_deadline']}"
        return u

    def top_garuda_row(self):
        """Centred 3 cm Garuda with ชั้นความเร็ว at the left (as in แบบที่ ๑ and ๓)."""
        urg = self.urgency_text()
        if not urg:
            return self.garuda(3.0)
        table = self.doc.add_table(rows=1, cols=3)
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        table.autofit = False
        self._no_borders(table)
        widths = [Cm(6.5), Cm(3.0), Cm(6.5)]
        for i, w in enumerate(widths):
            table.columns[i].width = w
            table.rows[0].cells[i].width = w
        c0, c1, _ = table.rows[0].cells
        c0.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.BOTTOM
        self.cell_para(c0, urg, bold=True, size=max(32, self.spec.get("urgency_size", 32)), color=RED)
        p = self.cell_para(c1, align="center")
        if GARUDA.exists():
            p.add_run().add_picture(str(GARUDA), height=Cm(3.0))
        self._zero_cell_margins(table)
        return table

    @staticmethod
    def _zero_cell_margins(table):
        tblPr = table._tbl.tblPr
        mar = OxmlElement("w:tblCellMar")
        for edge in ("left", "right"):
            el = OxmlElement(f"w:{edge}")
            el.set(qn("w:w"), "0")
            el.set(qn("w:type"), "dxa")
            mar.append(el)
        tblPr.append(mar)

    def number_agency(self, label="ที่"):
        """'ที่ xx ๐๐๐๐/๑๒๓' left, ส่วนราชการเจ้าของหนังสือ flush right (multi-line allowed)."""
        agency = self.spec.get("agency") or []
        if isinstance(agency, str):
            agency = [agency]
        number = self.spec.get("number", "")
        p = self.para(before=0)
        p.paragraph_format.tab_stops.add_tab_stop(Cm(TEXT_W), WD_TAB_ALIGNMENT.RIGHT)
        self.run(p, f"{label} {number}".rstrip() if number else f"{label} ")
        if agency:
            self.run(p, "\t" + agency[0])
        for line in agency[1:]:
            self.para(line, align="right")

    def body(self, paragraphs, before=6):
        """Body paragraphs. Strings get a 2.5 cm first-line indent; sub-items such as
        '๑.๑ ...' or '(๑) ...' get deeper indents. A dict may override:
        {"text": "...", "indent": 3.5, "align": "center", "bold": true}."""
        for i, item in enumerate(paragraphs or []):
            if isinstance(item, dict):
                txt = item.get("text", "")
                first = item.get("indent")
                align = item.get("align", "justify")
                bold = item.get("bold", False)
                # a continuation line (explicit indent, e.g. a member's position) sits tight under its item
                b = item.get("before", 0 if "indent" in item else before)
            else:
                txt, first, align, bold, b = item, None, "justify", False, before
                if "\t" in txt:
                    b = 3
            if first is None:
                first = PARA_INDENT + self._level(txt) * 0.75
            if "\t" in txt and align == "justify":
                align = None  # justified tab lines look broken
            p = self.para(txt, first=first, before=b, align=align, bold=bold)
            if "\t" in txt:
                tab = item.get("tab") if isinstance(item, dict) else None
                p.paragraph_format.tab_stops.add_tab_stop(Cm(tab or self.spec.get("tab_stop", 10.0)))

    @staticmethod
    def _level(txt: str) -> int:
        s = thai_digits(txt.strip())
        if re.match(r"^[๐-๙]+\.[๐-๙]+\.[๐-๙]+", s):
            return 2
        if re.match(r"^[๐-๙]+\.[๐-๙]+", s) or re.match(r"^\([๐-๙]+\)", s):
            return 1
        return 0

    def keep_last_with_next(self, n=1):
        """Stop the signature block from being stranded alone on a new page."""
        paras = self.doc.paragraphs
        for p in paras[-n:]:
            p.paragraph_format.keep_with_next = True

    def signature(self, closing=None, before=12, gap_lines=3):
        signer = self.spec.get("signer") or {}
        self.keep_last_with_next(1)
        if closing:
            self.para(closing, align="center", left=SIGN_INDENT, before=before, keep_next=True)
        else:
            self.blank(1)
        for _ in range(gap_lines):
            self.para("", keep_next=True)
        name = signer.get("name", "")
        if name and not name.startswith("("):
            name = f"({name})"
        rank = signer.get("rank")  # ยศ — ภาคผนวก ๓ ข้อ ๓.๓: พิมพ์คำเต็มของยศไว้หน้าลายมือชื่อ
        if rank:
            self.para(rank, left=SIGN_INDENT + 1.0, keep_next=True)
        self.para(name or "(.....................................)", align="center", left=SIGN_INDENT,
                  keep_next=True)
        pos = signer.get("position") or []
        if isinstance(pos, str):
            pos = [pos]
        for line in pos:
            self.para(line, align="center", left=SIGN_INDENT, keep_next=True)

    def contact_block(self):
        contact = self.spec.get("contact") or []
        if isinstance(contact, str):
            contact = [contact]
        first = True
        for line in contact:
            self.para(line, before=24 if first else 0)
            first = False
        cc = self.spec.get("cc")
        if cc:
            if isinstance(cc, str):
                cc = [cc]
            self.para("สำเนาส่ง  " + cc[0], before=6)
            for line in cc[1:]:
                self.para(line, left=1.75)

    # ---------------------------------------------------------------- types
    def distribution(self):
        """บัญชีรายชื่อผู้รับแนบท้าย (หนังสือเวียน / สำเนาส่งตามรายชื่อที่แนบ)."""
        names = self.spec.get("distribution")
        if not names:
            return
        p = self.para()
        p.add_run().add_break(WD_BREAK.PAGE)
        title = self.spec.get("distribution_title", "บัญชีรายชื่อผู้รับ")
        self.para(title, align="center", bold=True)
        num = self.spec.get("number")
        if num:
            self.para(f"แนบท้ายหนังสือ ที่ {num}" + (f" ลงวันที่ {self.spec['date']}" if self.spec.get("date") else ""),
                      align="center")
        self.short_rule()
        for i, n in enumerate(names, 1):
            self.para(f"{i}. {n}", left=1.0, before=3)

    def build_external(self):
        s = self.spec
        self.top_garuda_row()
        self.number_agency("ที่")
        self.para(s.get("date", ""), left=DATE_INDENT, before=6)
        self.labelled("เรื่อง", s.get("subject"), before=6, tab=1.25)
        self.para(s.get("salutation", ""), before=6)
        ref = s.get("reference")
        if ref:
            ref = ref if isinstance(ref, list) else [ref]
            self.labelled("อ้างถึง", ref[0], before=6, tab=1.75)
            for line in ref[1:]:
                self.para(line, left=1.75)
        encl = s.get("enclosures")
        if encl:
            encl = encl if isinstance(encl, list) else [encl]
            if len(encl) == 1:
                self.labelled("สิ่งที่ส่งมาด้วย", encl[0], before=6, tab=2.75)
            else:
                for i, e in enumerate(encl):
                    lab = "สิ่งที่ส่งมาด้วย" if i == 0 else ""
                    self.labelled(lab, f"{i + 1}. {e}", before=6 if i == 0 else 0, tab=2.75)
        self.body(s.get("body"))
        self.signature(s.get("closing", "ขอแสดงความนับถือ"))
        self.contact_block()
        self.distribution()

    def build_internal(self):
        s = self.spec
        table = self.doc.add_table(rows=1, cols=3)
        table.autofit = False
        self._no_borders(table)
        self._zero_cell_margins(table)
        widths = [Cm(3.5), Cm(9.0), Cm(3.5)]
        for i, w in enumerate(widths):
            table.columns[i].width = w
            table.rows[0].cells[i].width = w
        c0, c1, _ = table.rows[0].cells
        c0.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.BOTTOM
        c1.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.BOTTOM
        p = self.cell_para(c0)
        if GARUDA.exists():
            p.add_run().add_picture(str(GARUDA), height=Cm(1.5))
        urg = self.urgency_text()
        if urg:
            self.run(p, " " + urg, bold=True, size=max(32, s.get("urgency_size", 32)), color=RED)
        self.cell_para(c1, "บันทึกข้อความ", align="center", bold=True, size=29)
        # header lines: labels 20 pt bold (common practice)
        p = self.para(before=6)
        self.run(p, "ส่วนราชการ", bold=True, size=20)
        self.run(p, "  " + s.get("department", s.get("agency", "") if isinstance(s.get("agency"), str) else ""))
        p = self.para()
        p.paragraph_format.tab_stops.add_tab_stop(Cm(8.5))
        self.run(p, "ที่", bold=True, size=20)
        self.run(p, "  " + s.get("number", "") + "\t")
        self.run(p, "วันที่", bold=True, size=20)
        self.run(p, "  " + s.get("date", ""))
        p = self.para()
        p.paragraph_format.left_indent = Cm(1.25)
        p.paragraph_format.first_line_indent = Cm(-1.25)
        p.paragraph_format.tab_stops.add_tab_stop(Cm(1.25))
        self.run(p, "เรื่อง", bold=True, size=20)
        self.run(p, "\t" + s.get("subject", ""))
        self.para(s.get("salutation", ""), before=12)
        self.body(s.get("body"))
        self.signature(s.get("closing"), gap_lines=3)
        self.distribution()

    def build_stamped(self):
        s = self.spec
        self.top_garuda_row()
        self.para(f"ที่ {s.get('number', '')}")
        self.para(f"ถึง {s.get('to', '')}")
        self.body(s.get("body"), before=12)
        issuer = s.get("issuer", "")
        self.para(issuer, align="center", left=SIGN_INDENT + 2, before=24, keep_next=True)
        self.blank(3)  # space for ตราชื่อส่วนราชการ (หมึกแดง) + ลายมือชื่อย่อกำกับตรา
        self.para(s.get("date", ""), align="center", left=SIGN_INDENT + 2)
        contact = s.get("contact") or []
        if isinstance(contact, str):
            contact = [contact]
        for i, line in enumerate(contact):
            self.para(line, before=24 if i == 0 else 0)
        self.distribution()

    def _title_block(self, lines):
        for i, (txt, bold) in enumerate(lines):
            if txt:
                self.para(txt, align="center", bold=bold, before=6 if i == 0 else 0)
        self.short_rule()

    def _dated_signature(self, label, date):
        self.keep_last_with_next(2)
        self.para(f"{label} {date}".strip(), left=4.0, before=18, keep_next=True)
        self.signature(None, before=0, gap_lines=2)

    def build_order(self):
        s = self.spec
        self.garuda(3.0)
        self._title_block([
            (s.get("title") or f"คำสั่ง{s.get('issuer', '')}", True),
            (f"ที่ {s.get('number', '')}", False),
            (f"เรื่อง {s.get('subject', '')}", False),
        ])
        self.body(s.get("body"))
        if s.get("effective"):
            self.para(s["effective"], first=PARA_INDENT, before=6, keep_next=True)
        self._dated_signature("สั่ง ณ วันที่", s.get("date", ""))

    def _articles(self):
        for a in self.spec.get("articles", []):
            if isinstance(a, dict) and a.get("chapter"):
                self.para(a["chapter"], align="center", bold=True, before=12)
                if a.get("chapter_title"):
                    self.para(a["chapter_title"], align="center", bold=True)
                continue
            self.body([a])

    def build_regulation(self, word="ระเบียบ"):
        s = self.spec
        self.garuda(3.0)
        lines = [
            (s.get("title") or f"{word}{s.get('issuer', '')}", True),
            (f"ว่าด้วย{s.get('topic', '')}", True),
        ]
        if s.get("edition"):
            lines.append((f"({s['edition']})" if not s["edition"].startswith("(") else s["edition"], True))
        lines.append((f"พ.ศ. {s.get('year', '')}", True))
        self._title_block(lines)
        self.body(s.get("preamble") or s.get("body"))
        self._articles()
        self._dated_signature("ประกาศ ณ วันที่", s.get("date", ""))

    def build_announcement(self):
        s = self.spec
        self.garuda(3.0)
        word = s.get("word", "ประกาศ")  # หรือ "แจ้งความ" (ข้อ ๒๐ วรรคท้าย)
        self._title_block([
            (s.get("title") or f"{word}{s.get('issuer', '')}", True),
            (f"เรื่อง {s.get('subject', '')}", True),
        ])
        self.body(s.get("body"))
        self._dated_signature(f"{word} ณ วันที่", s.get("date", ""))

    def _issuer_date_close(self):
        s = self.spec
        self.para(s.get("issuer", ""), align="center", left=SIGN_INDENT + 2, before=24)
        self.para(s.get("date", ""), align="center", left=SIGN_INDENT + 2)

    def build_statement(self, word="แถลงการณ์", garuda=True):
        s = self.spec
        if garuda:
            self.garuda(3.0)
        else:
            self.doc.sections[0].top_margin = Cm(2.5)
        lines = [
            (s.get("title") or f"{word}{s.get('issuer', '')}", True),
            (f"เรื่อง {s.get('subject', '')}", True),
        ]
        if s.get("edition"):
            lines.append((s["edition"], False))
        self._title_block(lines)
        self.body(s.get("body"))
        self._issuer_date_close()

    def build_certificate(self):
        s = self.spec
        self.garuda(3.0)
        self.number_agency("เลขที่")
        self.body(s.get("body"), before=18)
        self.para(f"ให้ไว้ ณ วันที่ {s.get('date', '')}", left=4.0, before=18, keep_next=True)
        photo = s.get("photo")
        if not photo:
            self.signature(None, before=0, gap_lines=2)
            return
        # photo box 4 x 6 cm on the left, signature block on the right (ข้อ ๒๔.๗)
        table = self.doc.add_table(rows=1, cols=2)
        table.autofit = False
        self._no_borders(table)
        table.columns[0].width = table.rows[0].cells[0].width = Cm(6.0)
        table.columns[1].width = table.rows[0].cells[1].width = Cm(10.0)
        left, right = table.rows[0].cells
        self.cell_para(left, "(ส่วนนี้ใช้สำหรับเรื่องสำคัญ)", size=14)
        box = left.add_table(rows=1, cols=1)
        box.style = self.doc.styles["Table Grid"]
        box.rows[0].height = Cm(6.0)
        box.columns[0].width = box.rows[0].cells[0].width = Cm(4.0)
        bc = box.rows[0].cells[0]
        bc.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
        self.cell_para(bc, "รูปถ่าย", align="center", size=14)
        holder = photo if isinstance(photo, dict) else {}
        self.cell_para(left, "", first=False)
        self.cell_para(left, "(ลงชื่อผู้ได้รับการรับรอง)", size=14, first=False)
        if holder.get("name"):
            self.cell_para(left, f"({holder['name']})", size=14, first=False)
        signer = s.get("signer") or {}
        for _ in range(3):
            self.cell_para(right, "", first=False)
        name = signer.get("name", "")
        if name and not name.startswith("("):
            name = f"({name})"
        self.cell_para(right, name, align="center", first=False)
        pos = signer.get("position") or []
        for line in ([pos] if isinstance(pos, str) else pos):
            self.cell_para(right, line, align="center", first=False)

    def build_minutes(self):
        s = self.spec
        self.doc.sections[0].top_margin = Cm(2.5)
        self._title_block([
            (s.get("title") or f"รายงานการประชุม{s.get('meeting', '')}", True),
            (f"ครั้งที่ {s.get('session', '')}", True),
            (f"เมื่อ{s.get('when', '')}", True),
            (f"ณ {s.get('where', '')}", True),
        ])

        def people(label, items):
            if items is None:
                return
            self.para(label, bold=True, underline=True, before=12)
            for i, it in enumerate(items, 1):
                if isinstance(it, dict):
                    nm = it.get("name", "")
                    role = it.get("role", "")
                    note = it.get("note", "")
                else:
                    nm, role, note = str(it), "", ""
                p = self.para(left=0.75)
                p.paragraph_format.tab_stops.add_tab_stop(Cm(8.0))
                self.run(p, f"{i}. {nm}" + (f"\t{role}" if role else ""))
                if note:
                    self.para(note, left=1.25)

        people("ผู้มาประชุม", s.get("present"))
        people("ผู้ไม่มาประชุม", s.get("absent"))
        people("ผู้เข้าร่วมประชุม", s.get("attendees"))
        if s.get("start_time"):
            p = self.para(before=12)
            self.run(p, "เริ่มประชุมเวลา", bold=True, underline=True)
            self.run(p, "  " + s["start_time"])
        self.body(s.get("body"))
        for ag in s.get("agenda", []):
            self.para(ag.get("title", ""), bold=True, before=12)
            self.body(ag.get("body"))
            if ag.get("resolution"):
                p = self.para(first=PARA_INDENT, before=6, align="justify")
                self.run(p, "มติที่ประชุม", bold=True)
                self.run(p, "  " + ag["resolution"])
        if s.get("end_time"):
            p = self.para(before=12)
            self.run(p, "เลิกประชุมเวลา", bold=True, underline=True)
            self.run(p, "  " + s["end_time"])
        rec = s.get("recorder") or {}
        self.blank(2)
        name = rec.get("name", "")
        if name and not name.startswith("("):
            name = f"({name})"
        self.para(name or ".........................................", align="center", left=7.0)
        if rec.get("position"):
            self.para(rec["position"], align="center", left=7.0)
        self.para("ผู้จดรายงานการประชุม", align="center", left=7.0)
        chk = s.get("checker")
        if chk:
            self.blank(2)
            n2 = chk.get("name", "")
            self.para(f"({n2})" if n2 and not n2.startswith("(") else n2, align="center", left=7.0)
            if chk.get("position"):
                self.para(chk["position"], align="center", left=7.0)
            self.para(chk.get("label", "ผู้ตรวจรายงานการประชุม"), align="center", left=7.0)

    # ---------------------------------------------------------------- driver
    def build(self):
        kind = TYPE_ALIASES.get(self.spec.get("type", "external"))
        if not kind:
            raise SystemExit(f"unknown type: {self.spec.get('type')}")
        sec = self.doc.sections[0]
        # Garuda pages: Garuda sits 1.5 cm from the top edge; others start at 2.5 cm
        sec.top_margin = Cm(1.5)
        self.secrecy_and_pages()
        {
            "external": self.build_external,
            "internal": self.build_internal,
            "stamped": self.build_stamped,
            "order": self.build_order,
            "regulation": lambda: self.build_regulation("ระเบียบ"),
            "bylaw": lambda: self.build_regulation("ข้อบังคับ"),
            "announcement": self.build_announcement,
            "statement": lambda: self.build_statement("แถลงการณ์", True),
            "news": lambda: self.build_statement("ข่าว", False),
            "certificate": self.build_certificate,
            "minutes": self.build_minutes,
        }[kind]()
        return self.doc


def to_pdf(docx_path: Path) -> Path | None:
    try:
        subprocess.run(["soffice", "--headless", "--convert-to", "pdf", "--outdir", str(docx_path.parent),
                        str(docx_path)], check=True, capture_output=True, timeout=180)
        return docx_path.with_suffix(".pdf")
    except Exception as e:  # noqa: BLE001
        print(f"PDF conversion skipped: {e}", file=sys.stderr)
        return None


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("spec")
    ap.add_argument("output")
    ap.add_argument("--pdf", action="store_true", help="also convert to PDF with LibreOffice (preview)")
    ap.add_argument("--font", default="TH SarabunPSK")
    ap.add_argument("--arabic-digits", action="store_true", help="keep Arabic digits instead of Thai digits")
    a = ap.parse_args()
    spec = json.loads(Path(a.spec).read_text(encoding="utf-8"))
    out = Path(a.output)
    out.parent.mkdir(parents=True, exist_ok=True)
    builder = Builder(spec, a.font, a.arabic_digits)
    builder.build().save(out)
    # Thai Word rules: bidi=th-TH language everywhere, w:cs + hint on Thai runs, themeFontLang, sz/szCs, bCs
    fix_docx(str(out), str(out), builder.font)
    problems = verify(str(out))
    print(f"wrote {out}" + ("" if not problems else "\nThai-Word check problems:\n  " + "\n  ".join(problems)))
    if a.pdf:
        pdf = to_pdf(out)
        if pdf:
            print(f"wrote {pdf}")


if __name__ == "__main__":
    main()
