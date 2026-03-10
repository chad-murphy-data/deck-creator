# -*- coding: utf-8 -*-
"""
Slick Minimal template v2 - with Fidelity Slab / Fidelity Sans fonts,
increased text sizes, adaptive layouts, bold lead-ins, rounded cards.
"""

import os
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.chart import XL_CHART_TYPE
from pptx.oxml.ns import qn
from lxml import etree

W = Inches(10)
H = Inches(5.625)

GREEN = RGBColor(0x36, 0x87, 0x27)
GREEN_LIGHT = RGBColor(0xF7, 0xF4, 0xE4)
GREEN_MID = RGBColor(0x1D, 0xE4, 0xCA)
BLUE = RGBColor(0x38, 0x80, 0xF3)
PURPLE = RGBColor(0x5B, 0x2C, 0x8F)
COBALT = RGBColor(0x04, 0x54, 0x7C)
GOLD = RGBColor(0xD4, 0xA8, 0x43)
DARK = RGBColor(0x40, 0x3F, 0x3E)
MID = RGBColor(0x55, 0x55, 0x55)
LIGHT = RGBColor(0xE5, 0xE5, 0xE5)
OFF_WHITE = RGBColor(0xF9, 0xF7, 0xF5)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
COLORS = [GREEN, BLUE, PURPLE, COBALT, GOLD]

TITLE_FONT = "Fidelity Slab"
TITLE_FONT_FALLBACK = "Georgia"
BODY_FONT = "Fidelity Sans"
BODY_FONT_FALLBACK = "Calibri"

LOGO_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "assets", "logo.png")

LM = Inches(0.9)
ACC_W = Inches(0.25)
CW = Inches(8.6)


# -- helpers --

def _add_rect(slide, x, y, w, h, fill_color):
    shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, x, y, w, h)
    shape.fill.solid(); shape.fill.fore_color.rgb = fill_color
    shape.line.fill.background(); return shape

def _add_rounded_rect(slide, x, y, w, h, fill_color):
    shape = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, y, w, h)
    shape.fill.solid(); shape.fill.fore_color.rgb = fill_color
    shape.line.fill.background()
    pg = shape._element.find('.//' + qn('a:prstGeom'))
    if pg is not None:
        av = pg.find(qn('a:avLst'))
        if av is None: av = etree.SubElement(pg, qn('a:avLst'))
        for g in av.findall(qn('a:gd')): av.remove(g)
        g = etree.SubElement(av, qn('a:gd')); g.set('name','adj'); g.set('fmla','val 5000')
    return shape

def _add_oval(slide, x, y, w, h, fill_color):
    shape = slide.shapes.add_shape(MSO_SHAPE.OVAL, x, y, w, h)
    shape.fill.solid(); shape.fill.fore_color.rgb = fill_color
    shape.line.fill.background(); return shape

def _font_fallback(font_name):
    """Return (primary, fallback) for a font name."""
    if font_name == TITLE_FONT:
        return TITLE_FONT, TITLE_FONT_FALLBACK
    if font_name == BODY_FONT:
        return BODY_FONT, BODY_FONT_FALLBACK
    return font_name, None

def _set_font_with_fallback(font_obj, font_name):
    """Set font name and add XML-level fallback via cs element."""
    primary, fallback = _font_fallback(font_name)
    font_obj.name = primary
    if fallback:
        el = font_obj._element  # rPr or defRPr element
        latin = el.find(qn('a:latin'))
        if latin is None:
            latin = etree.SubElement(el, qn('a:latin'))
            latin.set('typeface', primary)
        cs = el.find(qn('a:cs'))
        if cs is None:
            cs = etree.SubElement(el, qn('a:cs'))
        cs.set('typeface', fallback)

def _add_text_box(slide, x, y, w, h, text, font_name=None, font_size=12,
                  color=None, bold=False, italic=False, align=PP_ALIGN.LEFT,
                  valign=MSO_ANCHOR.TOP, line_spacing=None):
    font_name = font_name or BODY_FONT; color = color or DARK
    txBox = slide.shapes.add_textbox(x, y, w, h)
    tf = txBox.text_frame; tf.word_wrap = True
    p = tf.paragraphs[0]; p.text = text
    _set_font_with_fallback(p.font, font_name)
    p.font.size = Pt(font_size); p.font.color.rgb = color
    p.font.bold = bold; p.font.italic = italic; p.alignment = align
    if line_spacing: p.line_spacing = Pt(line_spacing)
    bp = tf._txBody.find(qn('a:bodyPr'))
    if bp is not None:
        bp.set('anchor', {MSO_ANCHOR.TOP:'t', MSO_ANCHOR.MIDDLE:'ctr', MSO_ANCHOR.BOTTOM:'b'}.get(valign,'t'))
    return txBox

def _accent(slide):
    _add_rect(slide, Inches(0), Inches(0), ACC_W, H, GREEN)

def _slide_title(slide, title, y=None, size=28):
    """Title at top with green rule. Moved UP for more content space."""
    _add_text_box(slide, LM, Inches(0.15), CW, Inches(0.65), title,
                  TITLE_FONT, size, DARK, bold=True, valign=MSO_ANCHOR.BOTTOM)
    _add_rect(slide, LM, Inches(0.85), Inches(2.5), Inches(0.04), GREEN)

# Content starts at 1.05" (below title + rule + breathing room)
CONTENT_TOP = Inches(1.05)

def _find_split(text):
    """Find where to split text into bold lead-in + regular continuation."""
    for delim in [' — ', ' – ']:
        idx = text.find(delim)
        if idx > 0: return text[:idx + len(delim)], text[idx + len(delim):]
    idx = text.find(': ')
    if idx > 0: return text[:idx + 1], text[idx + 2:]
    words = text.split(); char_count = 0
    for i, w in enumerate(words):
        char_count += len(w) + 1
        if i >= 3:
            ci = text.find(',', char_count - 1)
            if ci > 0 and ci < len(text) * 0.55: return text[:ci + 1], text[ci + 2:]
    if len(words) > 10:
        n = min(8, len(words) // 2)
        pos = sum(len(words[i]) + 1 for i in range(n))
        return text[:pos].rstrip(), text[pos:]
    return text, ""

def _add_split_text(slide, x, y, w, h, text, font_name, font_size, color, line_spacing=None):
    """Text with bold lead-in clause."""
    tb = slide.shapes.add_textbox(x, y, w, h); tf = tb.text_frame; tf.word_wrap = True
    bold_part, rest_part = _find_split(text)
    p = tf.paragraphs[0]
    r1 = p.add_run(); r1.text = bold_part + (" " if rest_part else "")
    _set_font_with_fallback(r1.font, font_name)
    r1.font.size = Pt(font_size)
    r1.font.color.rgb = color; r1.font.bold = True
    if rest_part:
        r2 = p.add_run(); r2.text = rest_part
        _set_font_with_fallback(r2.font, font_name)
        r2.font.size = Pt(font_size)
        r2.font.color.rgb = color; r2.font.bold = False
    p.alignment = PP_ALIGN.LEFT
    if line_spacing: p.line_spacing = Pt(line_spacing)
    bp = tf._txBody.find(qn('a:bodyPr'))
    if bp is not None: bp.set('anchor', 'ctr')
    return tb

def _adaptive_layout(n, content_top, slide_h, bottom_margin=Inches(0.35)):
    """Calculate card positions - fills available space, biased slightly up."""
    avail = slide_h - content_top - bottom_margin
    gap = Inches(0.12)
    rowH = int((avail - (n - 1) * gap) / n)
    max_row = Inches(1.15)
    if rowH > max_row: rowH = max_row
    total = n * rowH + (n - 1) * gap
    remaining = avail - total
    startY = Emu(content_top + int(remaining * 0.3))
    fs = 16 if n <= 3 else 15
    ls = 20 if n <= 3 else 19
    return startY, rowH, gap, fs, ls


# -- slide builders --

def build_title(prs, c):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    _add_rect(slide, Inches(0), Inches(0), ACC_W, H, GREEN)
    _add_text_box(slide, LM, Inches(1.0), Inches(6.5), Inches(1.5), c.get("title", "Title"),
                  TITLE_FONT, 40, DARK, bold=True, valign=MSO_ANCHOR.BOTTOM)
    _add_rect(slide, LM, Inches(2.65), Inches(2.5), Inches(0.04), GREEN)
    if c.get("subtitle"):
        _add_text_box(slide, LM, Inches(2.85), Inches(6.5), Inches(0.6), c["subtitle"],
                      BODY_FONT, 16, MID)
    if c.get("author"):
        _add_text_box(slide, LM, Inches(4.1), CW, Inches(0.35), c["author"],
                      BODY_FONT, 12, MID, bold=True)
    if c.get("date"):
        _add_text_box(slide, LM, Inches(4.45), CW, Inches(0.3), c["date"],
                      BODY_FONT, 11, MID)
    # Logo in upper-right
    logo = c.get("logo_path", LOGO_PATH)
    if logo and os.path.isfile(logo):
        slide.shapes.add_picture(logo, Inches(7.6), Inches(0.3), Inches(2.0))


def build_in_brief(prs, c):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    _accent(slide)
    _slide_title(slide, c.get("title", "In Brief"))
    bullets = c.get("bullets", [])
    n = len(bullets)
    startY, rowH, gap, fs, ls = _adaptive_layout(n, CONTENT_TOP, H)

    for i, b in enumerate(bullets):
        y = Emu(startY + i * (rowH + gap))
        col = COLORS[i % len(COLORS)]
        _add_rounded_rect(slide, LM, y, CW, rowH, OFF_WHITE)
        _add_rect(slide, LM, y, Inches(0.10), rowH, col)
        cs = Inches(0.40); cx = Emu(LM + Inches(0.22)); cy = Emu(y + (rowH - cs) // 2)
        _add_oval(slide, cx, cy, cs, cs, col)
        _add_text_box(slide, cx, cy, cs, cs, str(i+1), TITLE_FONT, 14, WHITE,
                      bold=True, align=PP_ALIGN.CENTER, valign=MSO_ANCHOR.MIDDLE)
        _add_split_text(slide, Emu(LM + Inches(0.8)), y, Emu(CW - Inches(0.95)), rowH,
                        b, BODY_FONT, fs, DARK, line_spacing=ls)


def build_section_divider(prs, c):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    bg = slide.background; fill = bg.fill; fill.solid(); fill.fore_color.rgb = GREEN
    if c.get("sectionNumber"):
        _add_text_box(slide, Inches(0.8), Inches(1.2), Inches(2), Inches(0.8),
                      f"0{c['sectionNumber']}", TITLE_FONT, 48, GREEN_MID, bold=True,
                      valign=MSO_ANCHOR.BOTTOM)
    _add_text_box(slide, Inches(0.8), Inches(2.1), Inches(8), Inches(1.0),
                  c.get("title", "Section"), TITLE_FONT, 36, WHITE, bold=True,
                  valign=MSO_ANCHOR.MIDDLE)
    _add_rect(slide, Inches(0.8), Inches(3.2), Inches(2.0), Inches(0.04), GREEN_MID)
    if c.get("subtitle"):
        _add_text_box(slide, Inches(0.8), Inches(3.4), Inches(8), Inches(0.5),
                      c["subtitle"], BODY_FONT, 15, GREEN_LIGHT)


def build_stat_callout(prs, c):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    _accent(slide); _slide_title(slide, c.get("title", "Key Metric"), size=22)
    _add_text_box(slide, LM, Inches(1.3), CW, Inches(1.8), c.get("stat", "—"),
                  TITLE_FONT, 80, GREEN, bold=True, align=PP_ALIGN.CENTER,
                  valign=MSO_ANCHOR.MIDDLE)
    if c.get("headline"):
        _add_text_box(slide, Inches(1.5), Inches(3.1), Inches(7), Inches(0.6),
                      c["headline"], BODY_FONT, 17, DARK, bold=True, align=PP_ALIGN.CENTER)
    if c.get("detail"):
        _add_text_box(slide, Inches(1.5), Inches(3.75), Inches(7), Inches(0.8),
                      c["detail"], BODY_FONT, 13, MID, align=PP_ALIGN.CENTER, line_spacing=17)
    if c.get("source"):
        _add_text_box(slide, Inches(0.5), Inches(4.85), Inches(9), Inches(0.3),
                      c["source"], BODY_FONT, 9, MID, italic=True, align=PP_ALIGN.CENTER)


def build_quote(prs, c):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    _accent(slide); _slide_title(slide, c.get("title", "In Their Words"), size=22)
    _add_text_box(slide, Emu(LM - Inches(0.1)), Inches(1.2), Inches(0.8), Inches(0.8),
                  "\u201C", "Georgia", 64, GREEN_LIGHT, bold=True)
    _add_text_box(slide, Emu(LM + Inches(0.5)), Inches(1.5), Inches(7.5), Inches(2.0),
                  c.get("quote", ""), BODY_FONT, 18, DARK, italic=True,
                  valign=MSO_ANCHOR.MIDDLE, line_spacing=25)
    _add_rect(slide, Emu(LM + Inches(0.5)), Inches(3.7), Inches(1.5), Inches(0.04), GREEN)
    if c.get("attribution"):
        _add_text_box(slide, Emu(LM + Inches(0.5)), Inches(3.85), Inches(7), Inches(0.35),
                      c["attribution"], BODY_FONT, 13, MID)
    if c.get("context"):
        _add_text_box(slide, Emu(LM + Inches(0.5)), Inches(4.2), Inches(7), Inches(0.35),
                      c["context"], BODY_FONT, 11, MID, italic=True)


def build_comparison(prs, c):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    _accent(slide); _slide_title(slide, c.get("title", "Comparison"), size=24)
    cardY = CONTENT_TOP; cardH = Inches(3.6)
    _add_rounded_rect(slide, LM, cardY, Inches(4.0), cardH, OFF_WHITE)
    _add_rect(slide, LM, cardY, Inches(0.08), cardH, COBALT)
    _add_text_box(slide, Emu(LM + Inches(0.2)), Emu(cardY + Inches(0.15)),
                  Inches(3.6), Inches(0.4), c.get("leftLabel", "Before"),
                  TITLE_FONT, 15, COBALT, bold=True)
    for i, item in enumerate(c.get("leftItems", [])):
        _add_text_box(slide, Emu(LM + Inches(0.2)), Emu(cardY + Inches(0.65) + i * Inches(0.6)),
                      Inches(3.6), Inches(0.55), item, BODY_FONT, 13, DARK,
                      valign=MSO_ANCHOR.MIDDLE)
    _add_rect(slide, Inches(5.05), Emu(cardY + Inches(0.2)), Inches(0.03), Emu(cardH - Inches(0.4)), LIGHT)
    _add_rounded_rect(slide, Inches(5.25), cardY, Inches(4.25), cardH, OFF_WHITE)
    _add_rect(slide, Inches(5.25), cardY, Inches(0.08), cardH, GREEN)
    _add_text_box(slide, Inches(5.5), Emu(cardY + Inches(0.15)),
                  Inches(3.8), Inches(0.4), c.get("rightLabel", "After"),
                  TITLE_FONT, 15, GREEN, bold=True)
    for i, item in enumerate(c.get("rightItems", [])):
        _add_text_box(slide, Inches(5.5), Emu(cardY + Inches(0.65) + i * Inches(0.6)),
                      Inches(3.85), Inches(0.55), item, BODY_FONT, 13, DARK,
                      valign=MSO_ANCHOR.MIDDLE)


def build_text_graph(prs, c):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    _accent(slide); _slide_title(slide, c.get("title", "Title"), size=24)
    texts = c.get("text", []); texts = [texts] if not isinstance(texts, list) else texts
    for i, t in enumerate(texts):
        _add_text_box(slide, LM, Emu(CONTENT_TOP + i * Inches(1.2)),
                      Inches(4.0), Inches(1.1), t, BODY_FONT, 13, DARK, line_spacing=18)
    _add_rect(slide, Inches(5.1), CONTENT_TOP, Inches(0.03), Inches(3.5), LIGHT)
    chart_data_raw = c.get("chartData", [{"name": "S1", "labels": ["A","B","C"], "values": [25,45,30]}])
    from pptx.chart.data import CategoryChartData
    chart_data = CategoryChartData()
    if chart_data_raw:
        cd = chart_data_raw[0]
        chart_data.categories = cd.get("labels", ["A","B","C"])
        chart_data.add_series(cd.get("name","Series 1"), cd.get("values",[25,45,30]))
    ct = {"line": XL_CHART_TYPE.LINE, "pie": XL_CHART_TYPE.PIE}.get(c.get("chartType","bar"), XL_CHART_TYPE.COLUMN_CLUSTERED)
    slide.shapes.add_chart(ct, Inches(5.3), CONTENT_TOP, Inches(4.2), Inches(3.8), chart_data)
    if c.get("note"):
        _add_text_box(slide, Inches(5.3), Inches(5.0), Inches(4.2), Inches(0.3),
                      c["note"], BODY_FONT, 8, MID, italic=True)


def build_process_flow(prs, c):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    _accent(slide); _slide_title(slide, c.get("title", "Process"), size=24)
    steps = c.get("steps", []); count = min(len(steps), 5)
    if count == 0: return
    total_w = 8.6; arrow_w = 0.3
    step_w = (total_w - (count - 1) * arrow_w) / count
    startX = 0.9; stepY = 1.05; stepH = 3.9
    for i, step in enumerate(steps[:count]):
        x = Inches(startX + i * (step_w + arrow_w)); y = Inches(stepY)
        sw = Inches(step_w); sh = Inches(stepH)
        _add_rounded_rect(slide, x, y, sw, sh, OFF_WHITE)
        _add_rect(slide, x, y, Inches(0.06), sh, COLORS[i % len(COLORS)])
        _add_text_box(slide, Emu(x + Inches(0.15)), Emu(y + Inches(0.15)),
                      Inches(0.35), Inches(0.35), str(i + 1),
                      TITLE_FONT, 16, COLORS[i % len(COLORS)], bold=True)
        _add_text_box(slide, Emu(x + Inches(0.15)), Emu(y + Inches(0.55)),
                      Emu(sw - Inches(0.3)), Inches(0.65),
                      step.get("title", ""), BODY_FONT, 13, DARK, bold=True)
        if step.get("detail"):
            _add_text_box(slide, Emu(x + Inches(0.15)), Emu(y + Inches(1.25)),
                          Emu(sw - Inches(0.3)), Inches(2.3),
                          step["detail"], BODY_FONT, 11, MID, line_spacing=15)
        if i < count - 1:
            _add_text_box(slide, Emu(x + sw + Inches(0.02)), Emu(y + Inches(1.5)),
                          Emu(Inches(arrow_w) - Inches(0.04)), Inches(0.4),
                          "\u2192", BODY_FONT, 16, LIGHT, align=PP_ALIGN.CENTER,
                          valign=MSO_ANCHOR.MIDDLE)


def build_matrix(prs, c):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    _accent(slide); _slide_title(slide, c.get("title", "Framework"), size=24)
    quads = c.get("quadrants", [{},{},{},{}])
    accents = [GREEN, BLUE, PURPLE, COBALT]
    qW = Inches(4.05); qH = Inches(1.85); gap = Inches(0.2)
    sX = LM; sY = CONTENT_TOP
    for i, q in enumerate(quads[:4]):
        col = i % 2; row = i // 2
        x = Emu(sX + col * (qW + gap)); y = Emu(sY + row * (qH + gap))
        _add_rounded_rect(slide, x, y, qW, qH, OFF_WHITE)
        _add_rect(slide, x, y, Inches(0.06), qH, accents[i])
        _add_text_box(slide, Emu(x + Inches(0.2)), Emu(y + Inches(0.12)),
                      Emu(qW - Inches(0.4)), Inches(0.35),
                      q.get("label", ""), TITLE_FONT, 12, accents[i], bold=True)
        _add_text_box(slide, Emu(x + Inches(0.2)), Emu(y + Inches(0.5)),
                      Emu(qW - Inches(0.4)), Emu(qH - Inches(0.65)),
                      q.get("detail", ""), BODY_FONT, 11, DARK, line_spacing=15)


def build_methods(prs, c):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    _accent(slide); _slide_title(slide, c.get("title", "Approach"), size=24)
    fields = c.get("fields", []); n = len(fields)
    avail = H - CONTENT_TOP - Inches(0.3)
    rowH = min(int(avail / n), Inches(0.82))
    gap = Inches(0.06) if n > 4 else Inches(0.1)
    total = n * rowH + (n - 1) * gap
    startY = Emu(CONTENT_TOP + int((avail - total) * 0.3))
    for i, f in enumerate(fields):
        y = Emu(startY + i * (rowH + gap))
        col = COLORS[i % len(COLORS)]
        _add_rounded_rect(slide, LM, y, CW, rowH, OFF_WHITE)
        _add_rect(slide, LM, y, Inches(0.10), rowH, col)
        _add_text_box(slide, Emu(LM + Inches(0.22)), y, Inches(1.8), rowH,
                      f.get("label", ""), BODY_FONT, 13, col, bold=True,
                      valign=MSO_ANCHOR.MIDDLE)
        _add_text_box(slide, Emu(LM + Inches(2.1)), y, Inches(6.3), rowH,
                      f.get("value", ""), BODY_FONT, 13, DARK,
                      valign=MSO_ANCHOR.MIDDLE)


def build_hypotheses(prs, c):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    _accent(slide); _slide_title(slide, c.get("title", "Hypotheses"), size=24)
    hyps = c.get("hypotheses", []); n = len(hyps)
    avail = H - CONTENT_TOP - Inches(0.3)
    rowH = min(int(avail / n), Inches(0.72))
    gap = Inches(0.08)
    total = n * rowH + (n - 1) * gap
    startY = Emu(CONTENT_TOP + int((avail - total) * 0.3))
    for i, h in enumerate(hyps):
        y = Emu(startY + i * (rowH + gap))
        col = COLORS[i % len(COLORS)]
        _add_rounded_rect(slide, LM, y, CW, rowH, OFF_WHITE)
        _add_rect(slide, LM, y, Inches(0.10), rowH, col)
        cs = Inches(0.36); cx = Emu(LM + Inches(0.18)); cy = Emu(y + (rowH - cs) // 2)
        _add_oval(slide, cx, cy, cs, cs, col)
        _add_text_box(slide, cx, cy, cs, cs, f"H{i+1}", TITLE_FONT, 11, WHITE,
                      bold=True, align=PP_ALIGN.CENTER, valign=MSO_ANCHOR.MIDDLE)
        _add_text_box(slide, Emu(LM + Inches(0.72)), y, Inches(5.95), rowH,
                      h.get("text", ""), BODY_FONT, 12, DARK,
                      valign=MSO_ANCHOR.MIDDLE)
        if h.get("status"):
            sc = GREEN if h["status"] == "Confirmed" else (COBALT if h["status"] == "Rejected" else MID)
            _add_text_box(slide, Inches(7.8), y, Inches(1.2), rowH,
                          h["status"], BODY_FONT, 10, sc, bold=True,
                          align=PP_ALIGN.CENTER, valign=MSO_ANCHOR.MIDDLE)


def build_wsn_dense(prs, c):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    _accent(slide); _slide_title(slide, c.get("title", "Key Finding"), size=24)
    cols = [("What", GREEN, c.get("what", {})),
            ("So What", BLUE, c.get("soWhat", {})),
            ("Now What", PURPLE, c.get("nowWhat", {}))]
    colW = Inches(2.75); gap = Inches(0.2); cardH = Inches(3.8)
    startX = LM; startY = CONTENT_TOP
    for i, (label, color, data) in enumerate(cols):
        x = Emu(startX + i * (colW + gap))
        _add_rounded_rect(slide, x, startY, colW, cardH, OFF_WHITE)
        _add_rect(slide, x, startY, Inches(0.06), cardH, color)
        _add_text_box(slide, Emu(x + Inches(0.2)), Emu(startY + Inches(0.15)),
                      Emu(colW - Inches(0.4)), Inches(0.35),
                      label, TITLE_FONT, 14, color, bold=True)
        _add_text_box(slide, Emu(x + Inches(0.2)), Emu(startY + Inches(0.55)),
                      Emu(colW - Inches(0.4)), Inches(1.1),
                      data.get("headline", ""), BODY_FONT, 12, DARK, bold=True,
                      line_spacing=16)
        if data.get("detail"):
            _add_text_box(slide, Emu(x + Inches(0.2)), Emu(startY + Inches(1.7)),
                          Emu(colW - Inches(0.4)), Inches(1.9),
                          data["detail"], BODY_FONT, 11, MID, line_spacing=14)


def build_wsn_reveal(prs, c):
    """Builds 3 progressive slides: What -> So What -> Now What."""
    WSN_COLORS = [GREEN, BLUE, PURPLE]
    WSN_LABELS = ["What We Found", "So What", "Now What"]

    def _draw_step_indicator(slide, x, y, num, label, color, active=True):
        cs = Inches(0.42)
        fill = color if active else LIGHT
        _add_oval(slide, x, y, cs, cs, fill)
        tc = WHITE if active else MID
        _add_text_box(slide, x, y, cs, cs, str(num), TITLE_FONT, 15, tc,
                      bold=True, align=PP_ALIGN.CENTER, valign=MSO_ANCHOR.MIDDLE)
        lw = Inches(1.2)
        _add_text_box(slide, Emu(x - (lw - cs) // 2), Emu(y + cs + Inches(0.04)),
                      lw, Inches(0.25), label, BODY_FONT, 9, color if active else MID,
                      bold=active, align=PP_ALIGN.CENTER)

    def _draw_connector(slide, x1, x2, y):
        _add_rect(slide, Emu(x1), y, Emu(x2 - x1), Inches(0.025), LIGHT)

    def _draw_step_bar(slide, active_count):
        bar_y = Emu(CONTENT_TOP + Inches(0.05))
        spacing = Inches(2.5)
        start_x = Emu(LM + Inches(1.5))
        cs = Inches(0.42)
        for i in range(3):
            sx = Emu(start_x + i * spacing)
            active = i < active_count
            _draw_step_indicator(slide, sx, bar_y, i + 1, WSN_LABELS[i],
                                 WSN_COLORS[i], active)
            if i < 2:
                conn_x1 = Emu(sx + cs + Inches(0.08))
                conn_x2 = Emu(start_x + (i + 1) * spacing - Inches(0.08))
                _draw_connector(slide, conn_x1, conn_x2, Emu(bar_y + cs // 2))

    def _draw_zone(slide, x, w, color, data, emphasis=False):
        card_y = Emu(CONTENT_TOP + Inches(0.95))
        card_h = Inches(3.7)
        _add_rounded_rect(slide, x, card_y, w, card_h, OFF_WHITE)
        _add_rect(slide, x, card_y, Inches(0.10), card_h, color)
        fs_head = 14 if emphasis else 13
        fs_detail = 12 if emphasis else 11
        _add_text_box(slide, Emu(x + Inches(0.25)), Emu(card_y + Inches(0.2)),
                      Emu(w - Inches(0.5)), Inches(0.85),
                      data.get("headline", ""), BODY_FONT, fs_head, DARK, bold=True,
                      line_spacing=18 if emphasis else 17)
        if data.get("detail"):
            _add_text_box(slide, Emu(x + Inches(0.25)), Emu(card_y + Inches(1.15)),
                          Emu(w - Inches(0.5)), Inches(2.2),
                          data["detail"], BODY_FONT, fs_detail, MID, line_spacing=16)

    s1 = prs.slides.add_slide(prs.slide_layouts[6])
    _accent(s1); _slide_title(s1, c.get("title", "Key Finding"), size=24)
    _draw_step_bar(s1, 1)
    _draw_zone(s1, LM, Inches(5.5), GREEN, c.get("what", {}), emphasis=True)

    s2 = prs.slides.add_slide(prs.slide_layouts[6])
    _accent(s2); _slide_title(s2, c.get("title", "Key Finding"), size=24)
    _draw_step_bar(s2, 2)
    _draw_zone(s2, LM, Inches(4.05), GREEN, c.get("what", {}))
    _draw_zone(s2, Inches(5.2), Inches(4.3), BLUE, c.get("soWhat", {}), emphasis=True)

    s3 = prs.slides.add_slide(prs.slide_layouts[6])
    _accent(s3); _slide_title(s3, c.get("title", "Key Finding"), size=24)
    _draw_step_bar(s3, 3)
    cond_y = Emu(CONTENT_TOP + Inches(0.95)); cond_h = Inches(1.75)
    _add_rounded_rect(s3, LM, cond_y, Inches(4.05), cond_h, OFF_WHITE)
    _add_rect(s3, LM, cond_y, Inches(0.08), cond_h, GREEN)
    _add_text_box(s3, Emu(LM + Inches(0.2)), Emu(cond_y + Inches(0.12)),
                  Inches(3.6), Inches(0.5),
                  c.get("what", {}).get("headline", ""), BODY_FONT, 11, DARK, bold=True)
    _add_text_box(s3, Emu(LM + Inches(0.2)), Emu(cond_y + Inches(0.65)),
                  Inches(3.6), Inches(0.9),
                  c.get("what", {}).get("detail", ""), BODY_FONT, 10, MID)

    _add_rounded_rect(s3, Inches(5.2), cond_y, Inches(4.3), cond_h, OFF_WHITE)
    _add_rect(s3, Inches(5.2), cond_y, Inches(0.08), cond_h, BLUE)
    _add_text_box(s3, Inches(5.45), Emu(cond_y + Inches(0.12)),
                  Inches(3.8), Inches(0.5),
                  c.get("soWhat", {}).get("headline", ""), BODY_FONT, 11, DARK, bold=True)
    _add_text_box(s3, Inches(5.45), Emu(cond_y + Inches(0.65)),
                  Inches(3.8), Inches(0.9),
                  c.get("soWhat", {}).get("detail", ""), BODY_FONT, 10, MID)

    nwY = Emu(cond_y + cond_h + Inches(0.15)); nwH = Inches(1.65)
    _add_rounded_rect(s3, LM, nwY, CW, nwH, OFF_WHITE)
    _add_rect(s3, LM, nwY, Inches(0.10), nwH, PURPLE)
    nw = c.get("nowWhat", {})
    _add_text_box(s3, Emu(LM + Inches(0.25)), Emu(nwY + Inches(0.15)),
                  Emu(CW - Inches(0.5)), Inches(0.6),
                  nw.get("headline", ""), BODY_FONT, 16, DARK, bold=True)
    if nw.get("detail"):
        _add_text_box(s3, Emu(LM + Inches(0.25)), Emu(nwY + Inches(0.8)),
                      Emu(CW - Inches(0.5)), Inches(0.65),
                      nw["detail"], BODY_FONT, 12, MID)

def build_findings_recs(prs, c):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    _accent(slide); _slide_title(slide, c.get("title", "Findings & Recommendations"), size=24)
    items = c.get("items", []); n = min(len(items), 5)
    avail = H - CONTENT_TOP - Inches(0.3)
    rowH = min(int((avail - (n - 1) * Inches(0.12)) / n), Inches(0.85))
    gap = Inches(0.12)
    total = n * rowH + (n - 1) * gap
    startY = Emu(CONTENT_TOP + int((avail - total) * 0.3))
    for i, item in enumerate(items[:5]):
        y = Emu(startY + i * (rowH + gap))
        col = COLORS[i % len(COLORS)]
        _add_rounded_rect(slide, LM, y, Inches(3.9), rowH, OFF_WHITE)
        _add_rect(slide, LM, y, Inches(0.10), rowH, col)
        _add_text_box(slide, Emu(LM + Inches(0.2)), y, Inches(3.5), rowH,
                      item.get("finding", ""), BODY_FONT, 12, DARK,
                      valign=MSO_ANCHOR.MIDDLE)
        acs = Inches(0.30); acx = Emu(Inches(4.97)); acy = Emu(y + (rowH - acs) // 2)
        _add_oval(slide, acx, acy, acs, acs, col)
        _add_text_box(slide, acx, acy, acs, acs,
                      "\u2192", BODY_FONT, 13, WHITE, align=PP_ALIGN.CENTER,
                      valign=MSO_ANCHOR.MIDDLE)
        _add_rounded_rect(slide, Inches(5.4), y, Inches(4.1), rowH, OFF_WHITE)
        _add_rect(slide, Inches(5.4), y, Inches(0.10), rowH, col)
        _add_text_box(slide, Inches(5.6), y, Inches(3.75), rowH,
                      item.get("recommendation", ""), BODY_FONT, 12, DARK,
                      valign=MSO_ANCHOR.MIDDLE)


def build_findings_recs_dense(prs, c):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    _accent(slide); _slide_title(slide, c.get("title", "Complete Findings"), size=22)
    items = c.get("items", []); n = min(len(items), 8)
    avail = H - CONTENT_TOP - Inches(0.2)
    rowH = min(int((avail - (n - 1) * Inches(0.06)) / n), Inches(0.48))
    gap = Inches(0.06)
    total = n * rowH + (n - 1) * gap
    startY = Emu(CONTENT_TOP + int((avail - total) * 0.3))
    for i, item in enumerate(items[:8]):
        y = Emu(startY + i * (rowH + gap))
        bg = OFF_WHITE if i % 2 == 0 else WHITE
        _add_rect(slide, LM, y, Inches(4.1), rowH, bg)
        _add_rect(slide, LM, y, Inches(0.04), rowH, GREEN)
        _add_text_box(slide, Emu(LM + Inches(0.15)), y, Inches(3.85), rowH,
                      item.get("finding", ""), BODY_FONT, 10, DARK,
                      valign=MSO_ANCHOR.MIDDLE)
        _add_rect(slide, Inches(5.15), y, Inches(4.35), rowH, bg)
        _add_rect(slide, Inches(5.15), y, Inches(0.04), rowH, BLUE)
        _add_text_box(slide, Inches(5.3), y, Inches(4.1), rowH,
                      item.get("recommendation", ""), BODY_FONT, 10, DARK,
                      valign=MSO_ANCHOR.MIDDLE)


def build_open_questions(prs, c):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    _accent(slide); _slide_title(slide, c.get("title", "Open Questions"), size=26)
    questions = c.get("questions", [])
    cardW = Inches(4.1); cardH = Inches(1.7); gX = Inches(0.4); gY = Inches(0.2)
    gridY = CONTENT_TOP
    for i, question in enumerate(questions[:4]):
        col = i % 2; row = i // 2
        x = Emu(LM + col * (cardW + gX)); y = Emu(gridY + row * (cardH + gY))
        _add_rounded_rect(slide, x, y, cardW, cardH, OFF_WHITE)
        _add_rect(slide, x, y, Inches(0.06), cardH, COLORS[i % len(COLORS)])
        _add_oval(slide, Emu(x + Inches(0.15)), Emu(y + Inches(0.12)),
                  Inches(0.35), Inches(0.35), COLORS[i % len(COLORS)])
        _add_text_box(slide, Emu(x + Inches(0.15)), Emu(y + Inches(0.12)),
                      Inches(0.35), Inches(0.35), str(i + 1),
                      TITLE_FONT, 14, WHITE, bold=True, align=PP_ALIGN.CENTER,
                      valign=MSO_ANCHOR.MIDDLE)
        _add_text_box(slide, Emu(x + Inches(0.2)), Emu(y + Inches(0.55)),
                      Emu(cardW - Inches(0.4)), Inches(1.0),
                      question, BODY_FONT, 13, DARK, line_spacing=17)


def build_agenda(prs, c):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    _accent(slide); _slide_title(slide, c.get("title", "Agenda"), size=26)
    items = c.get("items", []); n = len(items)
    avail = H - CONTENT_TOP - Inches(0.3)
    rowH = min(int(avail / n), Inches(0.72))
    gap = Inches(0.08)
    total = n * rowH + (n - 1) * gap
    startY = Emu(CONTENT_TOP + int((avail - total) * 0.3))
    for i, item in enumerate(items):
        y = Emu(startY + i * (rowH + gap))
        col = COLORS[i % len(COLORS)]
        _add_rounded_rect(slide, LM, y, CW, rowH, OFF_WHITE)
        _add_rect(slide, LM, y, Inches(0.10), rowH, col)
        cs = Inches(0.40); cx = Emu(LM + Inches(0.22)); cy = Emu(y + (rowH - cs) // 2)
        _add_oval(slide, cx, cy, cs, cs, col)
        _add_text_box(slide, cx, cy, cs, cs, str(i + 1), TITLE_FONT, 14, WHITE,
                      bold=True, align=PP_ALIGN.CENTER, valign=MSO_ANCHOR.MIDDLE)
        title_text = item if isinstance(item, str) else item.get("title", "")
        _add_text_box(slide, Emu(LM + Inches(0.8)), y, Inches(5.5), rowH,
                      title_text, BODY_FONT, 15, DARK, bold=True,
                      valign=MSO_ANCHOR.MIDDLE)
        if isinstance(item, dict) and item.get("detail"):
            _add_text_box(slide, Inches(7.5), y, Inches(1.8), rowH,
                          item["detail"], BODY_FONT, 12, MID,
                          align=PP_ALIGN.RIGHT, valign=MSO_ANCHOR.MIDDLE)


def build_progressive_reveal(prs, c):
    takeaways = c.get("takeaways", [])
    for n in range(min(len(takeaways), 5)):
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        _accent(slide); _slide_title(slide, c.get("title", "Building the Picture"))
        cur = takeaways[n]
        _add_rounded_rect(slide, LM, CONTENT_TOP, CW, Inches(2.4), OFF_WHITE)
        _add_rect(slide, LM, CONTENT_TOP, Inches(0.06), Inches(2.4), GREEN)
        _add_text_box(slide, Emu(LM + Inches(0.2)), Emu(CONTENT_TOP + Inches(0.1)),
                      Emu(CW - Inches(0.4)), Inches(0.6),
                      cur.get("headline", ""), BODY_FONT, 16, DARK, bold=True)
        if cur.get("detail"):
            _add_text_box(slide, Emu(LM + Inches(0.2)), Emu(CONTENT_TOP + Inches(0.75)),
                          Emu(CW - Inches(0.4)), Inches(1.4),
                          cur["detail"], BODY_FONT, 12, MID, line_spacing=16)
        _add_rect(slide, Inches(0), Inches(3.65), W, Inches(0.04), GREEN)
        _add_text_box(slide, LM, Inches(3.75), Inches(3), Inches(0.25),
                      "Running Takeaways", TITLE_FONT, 10, GREEN, bold=True)
        for j in range(n + 1):
            ty = Emu(Inches(4.05) + j * Inches(0.35))
            active = j == n
            _add_rect(slide, LM, Emu(ty + Inches(0.05)), Inches(0.12), Inches(0.12), GREEN)
            _add_text_box(slide, Emu(LM + Inches(0.25)), ty,
                          Emu(CW - Inches(0.25)), Inches(0.32),
                          takeaways[j].get("summary", takeaways[j].get("headline", "")),
                          BODY_FONT, 10, DARK if active else MID, bold=active,
                          valign=MSO_ANCHOR.MIDDLE)


def build_closer(prs, c):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    bg = slide.background; fill = bg.fill; fill.solid(); fill.fore_color.rgb = GREEN
    _add_text_box(slide, Inches(0.5), Inches(1.4), Inches(9), Inches(1.2),
                  c.get("title", "Thank You"), TITLE_FONT, 44, WHITE, bold=True,
                  align=PP_ALIGN.CENTER, valign=MSO_ANCHOR.BOTTOM)
    _add_rect(slide, Inches(3.75), Inches(2.75), Inches(2.5), Inches(0.04), GREEN_MID)
    if c.get("subtitle"):
        _add_text_box(slide, Inches(0.5), Inches(2.95), Inches(9), Inches(0.5),
                      c["subtitle"], BODY_FONT, 16, WHITE, align=PP_ALIGN.CENTER)
    if c.get("contact"):
        _add_text_box(slide, Inches(0.5), Inches(3.8), Inches(9), Inches(0.4),
                      c["contact"], BODY_FONT, 12, GREEN_LIGHT, align=PP_ALIGN.CENTER)


def build_timeline(prs, c):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    _accent(slide); _slide_title(slide, c.get("title", "Timeline"), size=24)
    milestones = c.get("milestones", [])
    n = len(milestones)
    if n == 0:
        return
    STATUS_CLR = {"complete": GREEN, "current": GOLD, "upcoming": MID}
    line_y = Emu(CONTENT_TOP + Inches(1.2))
    usable_w = CW - Inches(0.4)
    step = int(usable_w / max(n - 1, 1)) if n > 1 else 0
    start_x = Emu(LM + Inches(0.2))
    # horizontal line
    _add_rect(slide, start_x, Emu(line_y + Inches(0.12)), Emu(usable_w), Inches(0.04), LIGHT)
    for i, m in enumerate(milestones):
        status = m.get("status", "upcoming")
        clr = STATUS_CLR.get(status, MID)
        cx = Emu(start_x + i * step) if n > 1 else Emu(start_x + int(usable_w / 2))
        dot_sz = Inches(0.30)
        _add_oval(slide, Emu(cx - dot_sz // 2), Emu(line_y - dot_sz // 2 + Inches(0.14)),
                  dot_sz, dot_sz, clr)
        col_w = Inches(1.5) if n <= 5 else Inches(1.2)
        tx = Emu(cx - int(col_w) // 2)
        # date above
        _add_text_box(slide, tx, Emu(line_y - Inches(0.55)), col_w, Inches(0.35),
                      m.get("date", ""), BODY_FONT, 9, clr, bold=True, align=PP_ALIGN.CENTER)
        # title below
        _add_text_box(slide, tx, Emu(line_y + Inches(0.55)), col_w, Inches(0.4),
                      m.get("title", ""), BODY_FONT, 11, DARK, bold=True, align=PP_ALIGN.CENTER)
        # detail below title
        if m.get("detail"):
            _add_text_box(slide, tx, Emu(line_y + Inches(0.95)), col_w, Inches(1.0),
                          m["detail"], BODY_FONT, 9, MID, align=PP_ALIGN.CENTER, line_spacing=12)


def build_data_table(prs, c):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    _accent(slide); _slide_title(slide, c.get("title", "Data"), size=24)
    headers = c.get("headers", [])
    rows = c.get("rows", [])
    highlightCol = c.get("highlightCol", None)
    n_cols = len(headers)
    n_rows = len(rows)
    if n_cols == 0:
        return
    col_w = int(CW / n_cols)
    row_h = Inches(0.42)
    header_h = Inches(0.45)
    start_y = CONTENT_TOP
    # header row
    for j, hdr in enumerate(headers):
        x = Emu(LM + j * col_w)
        _add_rect(slide, x, start_y, Emu(col_w), header_h, GREEN)
        _add_text_box(slide, x, start_y, Emu(col_w), header_h,
                      hdr, BODY_FONT, 11, WHITE, bold=True,
                      align=PP_ALIGN.CENTER, valign=MSO_ANCHOR.MIDDLE)
    # data rows
    for i, row in enumerate(rows):
        y = Emu(start_y + header_h + i * row_h)
        bg = OFF_WHITE if i % 2 == 0 else WHITE
        for j, cell in enumerate(row):
            x = Emu(LM + j * col_w)
            cell_bg = bg
            cell_clr = DARK
            cell_bold = False
            if highlightCol is not None and j == highlightCol:
                cell_bg = RGBColor(0xE8, 0xF5, 0xE9)
                cell_clr = GREEN
                cell_bold = True
            _add_rect(slide, x, y, Emu(col_w), row_h, cell_bg)
            _add_text_box(slide, x, y, Emu(col_w), row_h,
                          str(cell), BODY_FONT, 10, cell_clr, bold=cell_bold,
                          align=PP_ALIGN.CENTER, valign=MSO_ANCHOR.MIDDLE)
    if c.get("note"):
        note_y = Emu(start_y + header_h + n_rows * row_h + Inches(0.15))
        _add_text_box(slide, LM, note_y, CW, Inches(0.3),
                      c["note"], BODY_FONT, 9, MID, italic=True)


def build_multi_stat(prs, c):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    _accent(slide); _slide_title(slide, c.get("title", "Key Metrics"), size=24)
    stats = c.get("stats", [])
    n = len(stats)
    if n == 0:
        return
    gap = Inches(0.2)
    card_w = int((CW - (n - 1) * gap) / n)
    card_h = Inches(3.2)
    start_y = Emu(CONTENT_TOP + Inches(0.15))
    for i, s in enumerate(stats):
        x = Emu(LM + i * (card_w + gap))
        col = COLORS[i % len(COLORS)]
        _add_rounded_rect(slide, x, start_y, Emu(card_w), card_h, OFF_WHITE)
        _add_rect(slide, x, start_y, Inches(0.06), card_h, col)
        # large value
        _add_text_box(slide, Emu(x + Inches(0.2)), Emu(start_y + Inches(0.3)),
                      Emu(card_w - Inches(0.4)), Inches(1.2),
                      s.get("value", ""), TITLE_FONT, 40, col, bold=True,
                      align=PP_ALIGN.CENTER, valign=MSO_ANCHOR.MIDDLE)
        # label
        _add_text_box(slide, Emu(x + Inches(0.2)), Emu(start_y + Inches(1.5)),
                      Emu(card_w - Inches(0.4)), Inches(0.5),
                      s.get("label", ""), BODY_FONT, 13, DARK, bold=True,
                      align=PP_ALIGN.CENTER, valign=MSO_ANCHOR.MIDDLE)
        # detail
        if s.get("detail"):
            _add_text_box(slide, Emu(x + Inches(0.2)), Emu(start_y + Inches(2.1)),
                          Emu(card_w - Inches(0.4)), Inches(0.9),
                          s["detail"], BODY_FONT, 10, MID,
                          align=PP_ALIGN.CENTER, line_spacing=14)
    if c.get("source"):
        _add_text_box(slide, LM, Inches(5.0), CW, Inches(0.3),
                      c["source"], BODY_FONT, 9, MID, italic=True, align=PP_ALIGN.CENTER)


def build_persona(prs, c):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    _accent(slide); _slide_title(slide, c.get("title", "Persona"), size=24)
    left_w = Inches(3.8); right_w = Inches(4.5); gap = Inches(0.3)
    card_h = Inches(3.8)
    card_y = CONTENT_TOP
    # Left panel - name, archetype, traits
    _add_rounded_rect(slide, LM, card_y, left_w, card_h, OFF_WHITE)
    _add_rect(slide, LM, card_y, Inches(0.08), card_h, GREEN)
    _add_text_box(slide, Emu(LM + Inches(0.25)), Emu(card_y + Inches(0.2)),
                  Emu(left_w - Inches(0.5)), Inches(0.5),
                  c.get("name", ""), TITLE_FONT, 18, DARK, bold=True)
    if c.get("archetype"):
        _add_text_box(slide, Emu(LM + Inches(0.25)), Emu(card_y + Inches(0.7)),
                      Emu(left_w - Inches(0.5)), Inches(0.35),
                      c["archetype"], BODY_FONT, 12, GREEN, bold=True)
    traits = c.get("traits", [])
    for i, trait in enumerate(traits):
        ty = Emu(card_y + Inches(1.2) + i * Inches(0.38))
        col = COLORS[i % len(COLORS)]
        _add_rect(slide, Emu(LM + Inches(0.25)), Emu(ty + Inches(0.06)),
                  Inches(0.10), Inches(0.10), col)
        _add_text_box(slide, Emu(LM + Inches(0.5)), ty,
                      Emu(left_w - Inches(0.7)), Inches(0.35),
                      trait, BODY_FONT, 11, DARK)
    # Right panel - strategy + detail
    rx = Emu(LM + left_w + gap)
    _add_rounded_rect(slide, rx, card_y, right_w, card_h, OFF_WHITE)
    _add_rect(slide, rx, card_y, Inches(0.08), card_h, BLUE)
    _add_text_box(slide, Emu(rx + Inches(0.25)), Emu(card_y + Inches(0.2)),
                  Emu(right_w - Inches(0.5)), Inches(0.35),
                  "Strategy", TITLE_FONT, 13, BLUE, bold=True)
    _add_text_box(slide, Emu(rx + Inches(0.25)), Emu(card_y + Inches(0.6)),
                  Emu(right_w - Inches(0.5)), Inches(1.2),
                  c.get("strategy", ""), BODY_FONT, 13, DARK, bold=True, line_spacing=18)
    if c.get("detail"):
        _add_text_box(slide, Emu(rx + Inches(0.25)), Emu(card_y + Inches(1.9)),
                      Emu(right_w - Inches(0.5)), Inches(1.6),
                      c["detail"], BODY_FONT, 11, MID, line_spacing=15)


def build_risk_tradeoff(prs, c):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    _accent(slide); _slide_title(slide, c.get("title", "Risks & Rewards"), size=24)
    SEV = {
        "high": RGBColor(0xC2, 0x3B, 0x22),
        "medium": GOLD,
        "low": GREEN,
    }
    risks = c.get("risks", [])
    rewards = c.get("rewards", [])
    col_w = Inches(4.1)
    gap_x = Inches(0.4)
    start_y = CONTENT_TOP
    # Left side - Risks header
    _add_text_box(slide, LM, start_y, col_w, Inches(0.35),
                  "Risks", TITLE_FONT, 14, RGBColor(0xC2, 0x3B, 0x22), bold=True)
    row_h = Inches(0.65)
    row_gap = Inches(0.08)
    for i, r in enumerate(risks):
        y = Emu(start_y + Inches(0.45) + i * (row_h + row_gap))
        sev = r.get("severity", "medium")
        sev_clr = SEV.get(sev, MID)
        _add_rounded_rect(slide, LM, y, col_w, row_h, OFF_WHITE)
        _add_rect(slide, LM, y, Inches(0.06), row_h, sev_clr)
        _add_text_box(slide, Emu(LM + Inches(0.2)), y, Emu(col_w - Inches(0.4)), Emu(row_h // 2),
                      r.get("label", ""), BODY_FONT, 11, DARK, bold=True,
                      valign=MSO_ANCHOR.BOTTOM)
        if r.get("detail"):
            _add_text_box(slide, Emu(LM + Inches(0.2)), Emu(y + row_h // 2),
                          Emu(col_w - Inches(0.4)), Emu(row_h // 2),
                          r["detail"], BODY_FONT, 9, MID, valign=MSO_ANCHOR.TOP)
    # Right side - Rewards header
    rx = Emu(LM + col_w + gap_x)
    _add_text_box(slide, rx, start_y, col_w, Inches(0.35),
                  "Rewards", TITLE_FONT, 14, GREEN, bold=True)
    for i, rw in enumerate(rewards):
        y = Emu(start_y + Inches(0.45) + i * (row_h + row_gap))
        _add_rounded_rect(slide, rx, y, col_w, row_h, OFF_WHITE)
        _add_rect(slide, rx, y, Inches(0.06), row_h, GREEN)
        _add_text_box(slide, Emu(rx + Inches(0.2)), y, Emu(col_w - Inches(0.4)), Emu(row_h // 2),
                      rw.get("label", ""), BODY_FONT, 11, DARK, bold=True,
                      valign=MSO_ANCHOR.BOTTOM)
        if rw.get("detail"):
            _add_text_box(slide, Emu(rx + Inches(0.2)), Emu(y + row_h // 2),
                          Emu(col_w - Inches(0.4)), Emu(row_h // 2),
                          rw["detail"], BODY_FONT, 9, MID, valign=MSO_ANCHOR.TOP)


def build_appendix(prs, c):
    """Compact appendix slide - no accent bar, muted styling."""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    # No _accent(slide) - appendix has no green bar
    _add_text_box(slide, LM, Inches(0.2), CW, Inches(0.55),
                  c.get("title", "Appendix"), TITLE_FONT, 22, MID, bold=True,
                  valign=MSO_ANCHOR.BOTTOM)
    _add_rect(slide, LM, Inches(0.8), Inches(1.5), Inches(0.03), LIGHT)
    sections = c.get("sections", [])
    row_h = Inches(0.55)
    gap = Inches(0.06)
    start_y = Emu(Inches(0.95))
    for i, sec in enumerate(sections):
        y = Emu(start_y + i * (row_h + gap))
        _add_text_box(slide, LM, y, Inches(2.0), row_h,
                      sec.get("label", ""), BODY_FONT, 10, MID, bold=True,
                      valign=MSO_ANCHOR.TOP)
        _add_text_box(slide, Emu(LM + Inches(2.2)), y, Inches(6.2), row_h,
                      sec.get("content", ""), BODY_FONT, 10, DARK,
                      valign=MSO_ANCHOR.TOP, line_spacing=14)


def build_before_after(prs, c):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    _accent(slide); _slide_title(slide, c.get("title", "Before & After"), size=24)
    card_h = Inches(3.5)
    card_y = Emu(CONTENT_TOP + Inches(0.1))
    before = c.get("before", {})
    after = c.get("after", {})
    intervention = c.get("intervention", "")
    bw = Inches(3.5)
    aw = Inches(3.5)
    mid_w = Inches(1.2)
    # Before card
    _add_rounded_rect(slide, LM, card_y, bw, card_h, OFF_WHITE)
    _add_rect(slide, LM, card_y, Inches(0.08), card_h, COBALT)
    _add_text_box(slide, Emu(LM + Inches(0.25)), Emu(card_y + Inches(0.2)),
                  Emu(bw - Inches(0.5)), Inches(0.45),
                  before.get("label", "Before"), TITLE_FONT, 14, COBALT, bold=True)
    _add_text_box(slide, Emu(LM + Inches(0.25)), Emu(card_y + Inches(0.75)),
                  Emu(bw - Inches(0.5)), Inches(2.4),
                  before.get("detail", ""), BODY_FONT, 12, DARK, line_spacing=17)
    # Middle arrow + intervention
    mid_x = Emu(LM + bw + Inches(0.1))
    arrow_y = Emu(card_y + Inches(1.2))
    _add_text_box(slide, mid_x, arrow_y, mid_w, Inches(0.5),
                  "\u2192", TITLE_FONT, 28, GREEN, bold=True,
                  align=PP_ALIGN.CENTER, valign=MSO_ANCHOR.MIDDLE)
    if intervention:
        _add_text_box(slide, mid_x, Emu(arrow_y + Inches(0.55)), mid_w, Inches(0.8),
                      intervention, BODY_FONT, 9, GREEN, bold=True,
                      align=PP_ALIGN.CENTER, line_spacing=12)
    # After card
    ax = Emu(LM + bw + mid_w + Inches(0.2))
    _add_rounded_rect(slide, ax, card_y, aw, card_h, OFF_WHITE)
    _add_rect(slide, ax, card_y, Inches(0.08), card_h, GREEN)
    _add_text_box(slide, Emu(ax + Inches(0.25)), Emu(card_y + Inches(0.2)),
                  Emu(aw - Inches(0.5)), Inches(0.45),
                  after.get("label", "After"), TITLE_FONT, 14, GREEN, bold=True)
    _add_text_box(slide, Emu(ax + Inches(0.25)), Emu(card_y + Inches(0.75)),
                  Emu(aw - Inches(0.5)), Inches(2.4),
                  after.get("detail", ""), BODY_FONT, 12, DARK, line_spacing=17)


def build_summary(prs, c):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    _accent(slide); _slide_title(slide, c.get("title", "Summary"), size=24)
    sections = c.get("sections", [])
    n = len(sections)
    if n == 0:
        return
    gap = Inches(0.2)
    col_w = int((CW - (n - 1) * gap) / n)
    card_h = Inches(3.8)
    card_y = CONTENT_TOP
    for i, sec in enumerate(sections):
        x = Emu(LM + i * (col_w + gap))
        col = COLORS[i % len(COLORS)]
        _add_rounded_rect(slide, x, card_y, Emu(col_w), card_h, OFF_WHITE)
        _add_rect(slide, x, card_y, Inches(0.06), card_h, col)
        _add_text_box(slide, Emu(x + Inches(0.2)), Emu(card_y + Inches(0.15)),
                      Emu(col_w - Inches(0.4)), Inches(0.4),
                      sec.get("heading", ""), TITLE_FONT, 13, col, bold=True)
        points = sec.get("points", [])
        for j, pt in enumerate(points):
            py = Emu(card_y + Inches(0.65) + j * Inches(0.5))
            _add_rect(slide, Emu(x + Inches(0.2)), Emu(py + Inches(0.07)),
                      Inches(0.08), Inches(0.08), col)
            _add_text_box(slide, Emu(x + Inches(0.4)), py,
                          Emu(col_w - Inches(0.6)), Inches(0.45),
                          pt, BODY_FONT, 10, DARK, line_spacing=13)


def build_quote_full(prs, c):
    """Full-bleed dark background quote. No title bar."""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    bg = slide.background; fill = bg.fill; fill.solid(); fill.fore_color.rgb = DARK
    # Big opening quotation mark
    _add_text_box(slide, Inches(1.0), Inches(0.8), Inches(1.0), Inches(1.0),
                  "\u201C", "Georgia", 72, GREEN, bold=True)
    # Quote text
    _add_text_box(slide, Inches(1.3), Inches(1.5), Inches(7.4), Inches(2.2),
                  c.get("quote", ""), BODY_FONT, 22, WHITE, italic=True,
                  align=PP_ALIGN.CENTER, valign=MSO_ANCHOR.MIDDLE, line_spacing=30)
    # Rule
    _add_rect(slide, Inches(4.0), Inches(3.9), Inches(2.0), Inches(0.04), GREEN)
    # Attribution
    if c.get("attribution"):
        _add_text_box(slide, Inches(1.0), Inches(4.1), Inches(8.0), Inches(0.4),
                      c["attribution"], BODY_FONT, 14, WHITE, bold=True,
                      align=PP_ALIGN.CENTER)
    if c.get("context"):
        _add_text_box(slide, Inches(1.0), Inches(4.5), Inches(8.0), Inches(0.35),
                      c["context"], BODY_FONT, 11, MID, italic=True,
                      align=PP_ALIGN.CENTER)


def build_stat_hero(prs, c):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    _accent(slide); _slide_title(slide, c.get("title", "Key Metric"), size=22)
    hero = c.get("hero", {})
    # Giant stat
    _add_text_box(slide, LM, Emu(CONTENT_TOP + Inches(0.1)), CW, Inches(1.6),
                  hero.get("value", ""), TITLE_FONT, 72, GREEN, bold=True,
                  align=PP_ALIGN.CENTER, valign=MSO_ANCHOR.MIDDLE)
    _add_text_box(slide, LM, Emu(CONTENT_TOP + Inches(1.7)), CW, Inches(0.45),
                  hero.get("label", ""), BODY_FONT, 16, DARK, bold=True,
                  align=PP_ALIGN.CENTER)
    # Supporting stats
    supporting = c.get("supporting", [])
    if supporting:
        n = len(supporting)
        gap = Inches(0.2)
        sw = int((CW - (n - 1) * gap) / n)
        sy = Emu(CONTENT_TOP + Inches(2.5))
        sh = Inches(1.3)
        for i, s in enumerate(supporting):
            x = Emu(LM + i * (sw + gap))
            col = COLORS[(i + 1) % len(COLORS)]
            _add_rounded_rect(slide, x, sy, Emu(sw), sh, OFF_WHITE)
            _add_rect(slide, x, sy, Inches(0.06), sh, col)
            _add_text_box(slide, Emu(x + Inches(0.15)), Emu(sy + Inches(0.1)),
                          Emu(sw - Inches(0.3)), Inches(0.7),
                          s.get("value", ""), TITLE_FONT, 24, col, bold=True,
                          align=PP_ALIGN.CENTER, valign=MSO_ANCHOR.MIDDLE)
            _add_text_box(slide, Emu(x + Inches(0.15)), Emu(sy + Inches(0.8)),
                          Emu(sw - Inches(0.3)), Inches(0.35),
                          s.get("label", ""), BODY_FONT, 10, DARK,
                          align=PP_ALIGN.CENTER, valign=MSO_ANCHOR.MIDDLE)
    if c.get("source"):
        _add_text_box(slide, LM, Inches(5.0), CW, Inches(0.3),
                      c["source"], BODY_FONT, 9, MID, italic=True, align=PP_ALIGN.CENTER)


def build_in_brief_featured(prs, c):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    _accent(slide); _slide_title(slide, c.get("title", "In Brief"), size=24)
    featured = c.get("featured", "")
    supporting = c.get("supporting", [])
    # Featured insight - large card
    feat_h = Inches(1.4)
    _add_rounded_rect(slide, LM, CONTENT_TOP, CW, feat_h, OFF_WHITE)
    _add_rect(slide, LM, CONTENT_TOP, Inches(0.10), feat_h, GREEN)
    _add_split_text(slide, Emu(LM + Inches(0.25)), CONTENT_TOP,
                    Emu(CW - Inches(0.5)), feat_h,
                    featured, BODY_FONT, 15, DARK, line_spacing=21)
    # Supporting bullets below
    bullet_y = Emu(CONTENT_TOP + feat_h + Inches(0.2))
    row_h = Inches(0.52)
    gap = Inches(0.08)
    for i, pt in enumerate(supporting):
        y = Emu(bullet_y + i * (row_h + gap))
        col = COLORS[(i + 1) % len(COLORS)]
        cs = Inches(0.28); cx = Emu(LM + Inches(0.1)); cy = Emu(y + (row_h - cs) // 2)
        _add_oval(slide, cx, cy, cs, cs, col)
        _add_text_box(slide, cx, cy, cs, cs, str(i + 1), TITLE_FONT, 10, WHITE,
                      bold=True, align=PP_ALIGN.CENTER, valign=MSO_ANCHOR.MIDDLE)
        _add_text_box(slide, Emu(LM + Inches(0.55)), y, Emu(CW - Inches(0.6)), row_h,
                      pt, BODY_FONT, 12, DARK, valign=MSO_ANCHOR.MIDDLE)


def build_persona_duo(prs, c):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    _accent(slide); _slide_title(slide, c.get("title", "Personas"), size=24)
    personas = c.get("personas", [])[:2]
    col_w = Inches(4.1); gap = Inches(0.4); card_h = Inches(3.8)
    card_y = CONTENT_TOP
    for i, p in enumerate(personas):
        x = Emu(LM + i * (col_w + gap))
        col = COLORS[i % len(COLORS)]
        _add_rounded_rect(slide, x, card_y, col_w, card_h, OFF_WHITE)
        _add_rect(slide, x, card_y, Inches(0.08), card_h, col)
        # Name
        _add_text_box(slide, Emu(x + Inches(0.25)), Emu(card_y + Inches(0.15)),
                      Emu(col_w - Inches(0.5)), Inches(0.45),
                      p.get("name", ""), TITLE_FONT, 16, DARK, bold=True)
        # Archetype
        if p.get("archetype"):
            _add_text_box(slide, Emu(x + Inches(0.25)), Emu(card_y + Inches(0.6)),
                          Emu(col_w - Inches(0.5)), Inches(0.3),
                          p["archetype"], BODY_FONT, 11, col, bold=True)
        # Traits
        traits = p.get("traits", [])
        for j, trait in enumerate(traits):
            ty = Emu(card_y + Inches(1.05) + j * Inches(0.32))
            _add_rect(slide, Emu(x + Inches(0.25)), Emu(ty + Inches(0.06)),
                      Inches(0.08), Inches(0.08), col)
            _add_text_box(slide, Emu(x + Inches(0.45)), ty,
                          Emu(col_w - Inches(0.7)), Inches(0.3),
                          trait, BODY_FONT, 10, DARK)
        # Strategy
        strat_y = Emu(card_y + Inches(1.05) + len(traits) * Inches(0.32) + Inches(0.15))
        _add_rect(slide, Emu(x + Inches(0.25)), strat_y,
                  Emu(col_w - Inches(0.5)), Inches(0.03), LIGHT)
        if p.get("strategy"):
            _add_text_box(slide, Emu(x + Inches(0.25)), Emu(strat_y + Inches(0.1)),
                          Emu(col_w - Inches(0.5)), Inches(1.0),
                          p["strategy"], BODY_FONT, 10, MID, line_spacing=14)


def build_process_flow_vertical(prs, c):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    _accent(slide); _slide_title(slide, c.get("title", "Process"), size=24)
    steps = c.get("steps", [])
    n = len(steps)
    if n == 0:
        return
    avail = H - CONTENT_TOP - Inches(0.3)
    arrow_h = Inches(0.25)
    total_arrows = (n - 1) * arrow_h
    row_h = int((avail - total_arrows) / n)
    max_row = Inches(0.9)
    if row_h > max_row:
        row_h = max_row
    total = n * row_h + (n - 1) * arrow_h
    start_y = Emu(CONTENT_TOP + int((avail - total) * 0.2))
    for i, step in enumerate(steps):
        y = Emu(start_y + i * (row_h + arrow_h))
        col = COLORS[i % len(COLORS)]
        _add_rounded_rect(slide, LM, y, CW, row_h, OFF_WHITE)
        _add_rect(slide, LM, y, Inches(0.10), row_h, col)
        # Number circle
        cs = Inches(0.36); cx = Emu(LM + Inches(0.18)); cy = Emu(y + (row_h - cs) // 2)
        _add_oval(slide, cx, cy, cs, cs, col)
        _add_text_box(slide, cx, cy, cs, cs, str(i + 1), TITLE_FONT, 12, WHITE,
                      bold=True, align=PP_ALIGN.CENTER, valign=MSO_ANCHOR.MIDDLE)
        # Title
        _add_text_box(slide, Emu(LM + Inches(0.7)), y, Inches(2.5), row_h,
                      step.get("title", ""), BODY_FONT, 13, DARK, bold=True,
                      valign=MSO_ANCHOR.MIDDLE)
        # Detail
        if step.get("detail"):
            _add_text_box(slide, Emu(LM + Inches(3.3)), y, Inches(5.1), row_h,
                          step["detail"], BODY_FONT, 11, MID,
                          valign=MSO_ANCHOR.MIDDLE)
        # Down arrow between steps
        if i < n - 1:
            arrow_y = Emu(y + row_h)
            _add_text_box(slide, Emu(LM + Inches(0.2)), arrow_y, Inches(0.5), arrow_h,
                          "\u2193", BODY_FONT, 14, LIGHT, align=PP_ALIGN.CENTER,
                          valign=MSO_ANCHOR.MIDDLE)


def build_text_cards(prs, c):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    _accent(slide); _slide_title(slide, c.get("title", "Key Points"), size=24)
    items = c.get("items", [])
    n = len(items)
    if n == 0:
        return
    # Determine grid: 2x2 for <=4, 2x3 for 5-6
    cols = 2
    rows = (n + cols - 1) // cols
    rows = min(rows, 3)
    gap_x = Inches(0.25); gap_y = Inches(0.2)
    card_w = int((CW - (cols - 1) * gap_x) / cols)
    avail_h = H - CONTENT_TOP - Inches(0.3)
    card_h = int((avail_h - (rows - 1) * gap_y) / rows)
    max_h = Inches(1.6)
    if card_h > max_h:
        card_h = max_h
    for i, item in enumerate(items[:cols * rows]):
        col_idx = i % cols
        row_idx = i // cols
        x = Emu(LM + col_idx * (card_w + gap_x))
        y = Emu(CONTENT_TOP + row_idx * (card_h + gap_y))
        color = COLORS[i % len(COLORS)]
        _add_rounded_rect(slide, x, y, Emu(card_w), Emu(card_h), OFF_WHITE)
        _add_rect(slide, x, y, Inches(0.06), Emu(card_h), color)
        if isinstance(item, dict):
            _add_text_box(slide, Emu(x + Inches(0.2)), Emu(y + Inches(0.12)),
                          Emu(card_w - Inches(0.4)), Inches(0.4),
                          item.get("title", ""), BODY_FONT, 13, DARK, bold=True)
            if item.get("detail"):
                _add_text_box(slide, Emu(x + Inches(0.2)), Emu(y + Inches(0.55)),
                              Emu(card_w - Inches(0.4)), Emu(card_h - Inches(0.7)),
                              item["detail"], BODY_FONT, 11, MID, line_spacing=15)
        else:
            _add_split_text(slide, Emu(x + Inches(0.2)), y,
                            Emu(card_w - Inches(0.4)), Emu(card_h),
                            str(item), BODY_FONT, 12, DARK, line_spacing=16)


def build_text_columns(prs, c):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    _accent(slide); _slide_title(slide, c.get("title", "Details"), size=24)
    columns = c.get("columns", [])
    n = min(len(columns), 3)
    if n == 0:
        return
    gap = Inches(0.25)
    col_w = int((CW - (n - 1) * gap) / n)
    card_h = Inches(3.8)
    card_y = CONTENT_TOP
    for i, col_data in enumerate(columns[:n]):
        x = Emu(LM + i * (col_w + gap))
        color = COLORS[i % len(COLORS)]
        _add_rect(slide, x, card_y, Emu(col_w), Inches(0.04), color)
        if isinstance(col_data, dict):
            heading = col_data.get("heading", "")
            body = col_data.get("body", "")
        else:
            heading = ""
            body = str(col_data)
        if heading:
            _add_text_box(slide, x, Emu(card_y + Inches(0.15)),
                          Emu(col_w), Inches(0.35),
                          heading, TITLE_FONT, 13, color, bold=True)
            _add_text_box(slide, x, Emu(card_y + Inches(0.55)),
                          Emu(col_w), Emu(card_h - Inches(0.6)),
                          body, BODY_FONT, 11, DARK, line_spacing=16)
        else:
            _add_text_box(slide, x, Emu(card_y + Inches(0.15)),
                          Emu(col_w), card_h,
                          body, BODY_FONT, 11, DARK, line_spacing=16)


def build_text_narrative(prs, c):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    _accent(slide); _slide_title(slide, c.get("title", "Narrative"), size=24)
    lede = c.get("lede", "")
    body = c.get("body", "")
    # Big lede text
    _add_text_box(slide, LM, CONTENT_TOP, CW, Inches(1.2),
                  lede, BODY_FONT, 18, DARK, bold=True, line_spacing=25)
    _add_rect(slide, LM, Emu(CONTENT_TOP + Inches(1.35)), Inches(2.0), Inches(0.03), GREEN)
    # Body text
    body_y = Emu(CONTENT_TOP + Inches(1.55))
    if isinstance(body, list):
        body_text = "\n\n".join(body)
    else:
        body_text = body
    _add_text_box(slide, LM, body_y, CW, Inches(2.8),
                  body_text, BODY_FONT, 12, DARK, line_spacing=17)


def build_text_nested(prs, c):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    _accent(slide); _slide_title(slide, c.get("title", "Overview"), size=24)
    items = c.get("items", [])
    n = len(items)
    if n == 0:
        return
    avail = H - CONTENT_TOP - Inches(0.3)
    gap = Inches(0.1)
    row_h = int((avail - (n - 1) * gap) / n)
    max_h = Inches(1.2)
    if row_h > max_h:
        row_h = max_h
    label_w = Inches(2.0)
    child_w = Emu(CW - label_w - Inches(0.2))
    for i, item in enumerate(items):
        y = Emu(CONTENT_TOP + i * (row_h + gap))
        col = COLORS[i % len(COLORS)]
        # Label block on the left
        _add_rounded_rect(slide, LM, y, label_w, Emu(row_h), col)
        _add_text_box(slide, Emu(LM + Inches(0.15)), y, Emu(label_w - Inches(0.3)), Emu(row_h),
                      item.get("text", ""), BODY_FONT, 12, WHITE, bold=True,
                      valign=MSO_ANCHOR.MIDDLE)
        # Children text on the right
        children = item.get("children", [])
        if children:
            child_x = Emu(LM + label_w + Inches(0.2))
            child_h = int(row_h / max(len(children), 1))
            for j, ch in enumerate(children):
                cy = Emu(y + j * child_h)
                ch_text = ch.get("text", ch) if isinstance(ch, dict) else str(ch)
                _add_rect(slide, child_x, Emu(cy + Inches(0.06)),
                          Inches(0.06), Inches(0.06), col)
                _add_text_box(slide, Emu(child_x + Inches(0.15)), cy,
                              child_w, Emu(child_h),
                              ch_text, BODY_FONT, 10, DARK, valign=MSO_ANCHOR.MIDDLE)
                # Nested grandchildren
                nested = ch.get("children", []) if isinstance(ch, dict) else []
                # Show nested as inline suffix if present
                if nested:
                    nested_texts = [gc.get("text", gc) if isinstance(gc, dict) else str(gc) for gc in nested]
                    nested_str = " | ".join(nested_texts)
                    _add_text_box(slide, Emu(child_x + Inches(0.15)), Emu(cy + Emu(child_h) - Inches(0.2)),
                                  child_w, Inches(0.2),
                                  nested_str, BODY_FONT, 8, MID, italic=True)


def build_text_split(prs, c):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    _accent(slide); _slide_title(slide, c.get("title", "Key Message"), size=24)
    left_w = Inches(4.0); right_w = Inches(4.3); gap = Inches(0.3)
    card_h = Inches(3.8)
    card_y = CONTENT_TOP
    # Left: big headline + detail
    _add_rounded_rect(slide, LM, card_y, left_w, card_h, OFF_WHITE)
    _add_rect(slide, LM, card_y, Inches(0.08), card_h, GREEN)
    _add_text_box(slide, Emu(LM + Inches(0.25)), Emu(card_y + Inches(0.3)),
                  Emu(left_w - Inches(0.5)), Inches(1.5),
                  c.get("headline", ""), TITLE_FONT, 20, DARK, bold=True,
                  line_spacing=27, valign=MSO_ANCHOR.TOP)
    if c.get("detail"):
        _add_text_box(slide, Emu(LM + Inches(0.25)), Emu(card_y + Inches(1.9)),
                      Emu(left_w - Inches(0.5)), Inches(1.6),
                      c["detail"], BODY_FONT, 11, MID, line_spacing=16)
    # Right: evidence points
    rx = Emu(LM + left_w + gap)
    points = c.get("points", [])
    row_h = Inches(0.65)
    row_gap = Inches(0.08)
    for i, pt in enumerate(points):
        y = Emu(card_y + Inches(0.1) + i * (row_h + row_gap))
        col = COLORS[i % len(COLORS)]
        _add_rounded_rect(slide, rx, y, right_w, row_h, OFF_WHITE)
        _add_rect(slide, rx, y, Inches(0.06), row_h, col)
        cs = Inches(0.28); cx = Emu(rx + Inches(0.15)); cy = Emu(y + (row_h - cs) // 2)
        _add_oval(slide, cx, cy, cs, cs, col)
        _add_text_box(slide, cx, cy, cs, cs, str(i + 1), TITLE_FONT, 10, WHITE,
                      bold=True, align=PP_ALIGN.CENTER, valign=MSO_ANCHOR.MIDDLE)
        _add_text_box(slide, Emu(rx + Inches(0.55)), y, Emu(right_w - Inches(0.7)), row_h,
                      pt, BODY_FONT, 11, DARK, valign=MSO_ANCHOR.MIDDLE)


def build_text_annotated(prs, c):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    _accent(slide); _slide_title(slide, c.get("title", "Annotations"), size=24)
    items = c.get("items", [])
    n = len(items)
    if n == 0:
        return
    avail = H - CONTENT_TOP - Inches(0.3)
    gap = Inches(0.1)
    row_h = int((avail - (n - 1) * gap) / n)
    max_h = Inches(0.85)
    if row_h > max_h:
        row_h = max_h
    label_w = Inches(1.8)
    text_w = Emu(CW - label_w - Inches(0.25))
    for i, item in enumerate(items):
        y = Emu(CONTENT_TOP + i * (row_h + gap))
        col = COLORS[i % len(COLORS)]
        # Colored label rectangle
        _add_rounded_rect(slide, LM, y, label_w, Emu(row_h), col)
        _add_text_box(slide, Emu(LM + Inches(0.1)), y, Emu(label_w - Inches(0.2)), Emu(row_h),
                      item.get("label", ""), BODY_FONT, 11, WHITE, bold=True,
                      valign=MSO_ANCHOR.MIDDLE, align=PP_ALIGN.CENTER)
        # Text on the right
        tx = Emu(LM + label_w + Inches(0.25))
        _add_text_box(slide, tx, y, text_w, Emu(row_h),
                      item.get("text", ""), BODY_FONT, 12, DARK,
                      valign=MSO_ANCHOR.MIDDLE, line_spacing=16)


BUILDERS = {
    "title": build_title, "in_brief": build_in_brief,
    "section_divider": build_section_divider, "stat_callout": build_stat_callout,
    "quote": build_quote, "comparison": build_comparison,
    "text_graph": build_text_graph, "process_flow": build_process_flow,
    "matrix": build_matrix, "methods": build_methods,
    "hypotheses": build_hypotheses, "wsn_dense": build_wsn_dense,
    "wsn_reveal": build_wsn_reveal, "findings_recs": build_findings_recs,
    "findings_recs_dense": build_findings_recs_dense,
    "open_questions": build_open_questions, "agenda": build_agenda,
    "progressive_reveal": build_progressive_reveal, "closer": build_closer,
    "timeline": build_timeline, "data_table": build_data_table,
    "multi_stat": build_multi_stat, "persona": build_persona,
    "risk_tradeoff": build_risk_tradeoff, "appendix": build_appendix,
    "before_after": build_before_after, "summary": build_summary,
    "quote_full": build_quote_full, "stat_hero": build_stat_hero,
    "in_brief_featured": build_in_brief_featured, "persona_duo": build_persona_duo,
    "process_flow_vertical": build_process_flow_vertical, "text_cards": build_text_cards,
    "text_columns": build_text_columns, "text_narrative": build_text_narrative,
    "text_nested": build_text_nested, "text_split": build_text_split,
    "text_annotated": build_text_annotated,
}

def build_deck(slide_configs, output_path):
    prs = Presentation(); prs.slide_width = W; prs.slide_height = H
    for slide_type, data in slide_configs:
        if slide_type == "skip":
            continue
        builder = BUILDERS.get(slide_type)
        if builder: builder(prs, data)
    prs.save(output_path); return output_path
