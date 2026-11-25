"""
Builds presentation/Profit-Portfolio-Analyzer.pptx

Interactive PowerPoint deck explaining the Profit Portfolio Analyzer:
what it does, how every file works, how it is deployed on AWS Amplify,
and what makes it different.

Usage:
    pip install python-pptx
    python presentation/build_deck.py

Edit LIVE_URL below once you have your Amplify URL, then re-run.
"""
import os
import re

from lxml import etree
from pptx import Presentation
from pptx.chart.data import BubbleChartData, CategoryChartData
from pptx.dml.color import RGBColor
from pptx.enum.chart import (XL_CHART_TYPE, XL_LABEL_POSITION, XL_LEGEND_POSITION,
                            XL_TICK_LABEL_POSITION, XL_TICK_MARK)
from pptx.enum.shapes import MSO_CONNECTOR, MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, MSO_AUTO_SIZE, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Emu, Inches, Pt

# ----------------------------------------------------------------- settings
LIVE_URL = "https://main.YOUR-APP-ID.amplifyapp.com"   # <- replace with your Amplify URL
LIVE_LABEL = "main.YOUR-APP-ID.amplifyapp.com"
GITHUB_URL = "https://github.com/Umbrella912/profit-portfolio-analyzer"
OUT = os.path.join(os.path.dirname(os.path.abspath(__file__)), "Profit-Portfolio-Analyzer.pptx")

# ------------------------------------------------------------------- theme
def rgb(h):
    return RGBColor.from_string(h.lstrip("#").upper())

BG = rgb("0C0E14")
PANEL = rgb("14171F")
PANEL2 = rgb("1A1D28")
CODEBG = rgb("0F1118")
BORDER = rgb("262A36")
GRID = rgb("1F2330")
MUTED = rgb("8A8FA3")
SOFT = rgb("B9BDCC")
TEXT = rgb("ECEDF5")
PURPLE = rgb("AC99FF")
MINT = rgb("54D6B3")
AMBER = rgb("EDBD6C")
ROSE = rgb("EF8395")
SKY = rgb("82B9F4")
CYAN = rgb("68D3E5")
LIME = rgb("C9DB7B")
PINK = rgb("EEA5DC")
DARKPURPLE = rgb("2A2347")

HEAD = "Segoe UI Semibold"
BODY = "Segoe UI"
MONO = "Consolas"

SW, SH = 13.333, 7.5

# ----------------------------------------------------------- slide registry
SECTIONS = [
    ("project", "Project"),
    ("features", "Features"),
    ("arch", "How it works"),
    ("files", "Files"),
    ("deploy", "Deployment"),
    ("diff", "Why different"),
]
SECTION_NAME = dict(SECTIONS)

ORDER = [
    ("title", None),
    ("agenda", None),
    ("problem", "project"),
    ("solution", "project"),
    ("hub", "features"),
    ("f1", "features"), ("f2", "features"), ("f3", "features"), ("f4", "features"),
    ("f5", "features"), ("f6", "features"), ("f7", "features"), ("f8", "features"),
    ("demo", "features"),
    ("architecture", "arch"),
    ("pipeline", "arch"),
    ("modules", "arch"),
    ("tree", "files"),
    ("htmlcss", "files"),
    ("scriptjs", "files"),
    ("amplifycsv", "files"),
    ("deployflow", "deploy"),
    ("deploysteps", "deploy"),
    ("compare", "diff"),
    ("different", "diff"),
    ("skills", "diff"),
    ("thanks", "diff"),
]
SECTION_FIRST = {}
for k, s in ORDER:
    if s and s not in SECTION_FIRST:
        SECTION_FIRST[s] = k

prs = Presentation()
prs.slide_width = Inches(SW)
prs.slide_height = Inches(SH)
BLANK = prs.slide_layouts[6]
S = {}
for key, _ in ORDER:
    S[key] = prs.slides.add_slide(BLANK)
KEYS = [k for k, _ in ORDER]


# ------------------------------------------------------------------ helpers
def I(v):
    return Inches(v)


def no_shadow(shape):
    try:
        shape.shadow.inherit = False
    except Exception:
        pass


def set_bg(slide):
    slide.background.fill.solid()
    slide.background.fill.fore_color.rgb = BG


def link(shapes, target):
    """Make shapes clickable: target is a slide key or an external URL."""
    if not isinstance(shapes, (list, tuple)):
        shapes = [shapes]
    for sh in shapes:
        if isinstance(target, str) and target.startswith("http"):
            sh.click_action.hyperlink.address = target
        else:
            sh.click_action.target_slide = S[target]


def _apply_run(run, size, color, bold, font, italic=False):
    f = run.font
    f.size = Pt(size)
    f.color.rgb = color
    f.bold = bold
    f.italic = italic
    f.name = font


def set_bullet(p, char="•", color=PURPLE, indent=0.22):
    pPr = p._p.get_or_add_pPr()
    pPr.set("marL", str(int(Inches(indent))))
    pPr.set("indent", str(-int(Inches(indent))))
    for tag in ("a:buClr", "a:buFont", "a:buChar", "a:buNone"):
        el = pPr.find(qn(tag))
        if el is not None:
            pPr.remove(el)
    buClr = etree.SubElement(pPr, qn("a:buClr"))
    srgb = etree.SubElement(buClr, qn("a:srgbClr"))
    srgb.set("val", str(color))
    buFont = etree.SubElement(pPr, qn("a:buFont"))
    buFont.set("typeface", "Arial")
    buChar = etree.SubElement(pPr, qn("a:buChar"))
    buChar.set("char", char)


def write(tf, paras, size=14, color=TEXT, bold=False, font=BODY, align=PP_ALIGN.LEFT,
          spacing=None, space_after=0, bullet=None, bullet_color=PURPLE):
    """paras: str | list of (str | list of (text, opts))."""
    if isinstance(paras, str):
        paras = [paras]
    tf.word_wrap = True
    first = True
    for para in paras:
        p = tf.paragraphs[0] if first else tf.add_paragraph()
        first = False
        p.alignment = align
        if spacing:
            p.line_spacing = spacing
        if space_after:
            p.space_after = Pt(space_after)
        runs = [(para, {})] if isinstance(para, str) else para
        popts = {}
        for text, o in runs:
            r = p.add_run()
            r.text = text
            _apply_run(r, o.get("size", size), o.get("color", color), o.get("bold", bold),
                       o.get("font", font), o.get("italic", False))
            if o.get("url"):
                r.hyperlink.address = o["url"]
            popts.update({k: v for k, v in o.items() if k in ("bullet", "align", "after")})
        if popts.get("align") is not None:
            p.alignment = popts["align"]
        if popts.get("after") is not None:
            p.space_after = Pt(popts["after"])
        b = popts.get("bullet", bullet)
        if b:
            set_bullet(p, b, bullet_color)
    return tf


def text(slide, x, y, w, h, paras, size=14, color=TEXT, bold=False, font=BODY,
         align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP, spacing=None, space_after=0,
         bullet=None, bullet_color=PURPLE, margin=0.0):
    tb = slide.shapes.add_textbox(I(x), I(y), I(w), I(h))
    tf = tb.text_frame
    tf.auto_size = MSO_AUTO_SIZE.NONE
    tf.vertical_anchor = anchor
    for side in ("left", "right", "top", "bottom"):
        setattr(tf, "margin_" + side, I(margin))
    write(tf, paras, size, color, bold, font, align, spacing, space_after, bullet, bullet_color)
    return tb


def box(slide, x, y, w, h, fill=PANEL, line=BORDER, radius=0.06, shape=MSO_SHAPE.ROUNDED_RECTANGLE,
        line_w=0.75):
    sh = slide.shapes.add_shape(shape, I(x), I(y), I(w), I(h))
    no_shadow(sh)
    if fill is None:
        sh.fill.background()
    else:
        sh.fill.solid()
        sh.fill.fore_color.rgb = fill
    if line is None:
        sh.line.fill.background()
    else:
        sh.line.color.rgb = line
        sh.line.width = Pt(line_w)
    if shape == MSO_SHAPE.ROUNDED_RECTANGLE:
        sh.adjustments[0] = min(0.5, radius / min(w, h))
    sh.text_frame.text = ""
    return sh


def box_text(sh, paras, size=13, color=TEXT, bold=False, font=BODY, align=PP_ALIGN.LEFT,
             anchor=MSO_ANCHOR.TOP, margin=(0.18, 0.14, 0.18, 0.14), spacing=None,
             space_after=0, bullet=None, bullet_color=PURPLE):
    tf = sh.text_frame
    tf.auto_size = MSO_AUTO_SIZE.NONE
    tf.vertical_anchor = anchor
    l, t, r, b = margin
    tf.margin_left, tf.margin_top, tf.margin_right, tf.margin_bottom = I(l), I(t), I(r), I(b)
    write(tf, paras, size, color, bold, font, align, spacing, space_after, bullet, bullet_color)
    return sh


def button(slide, x, y, w, h, label, target, primary=False, size=12, color=None):
    fill = PURPLE if primary else PANEL2
    line = PURPLE if primary else rgb("3A3550")
    fg = rgb("14111F") if primary else (color or TEXT)
    sh = box(slide, x, y, w, h, fill=fill, line=line, radius=h / 2)
    box_text(sh, label, size=size, color=fg, bold=True, font=HEAD, align=PP_ALIGN.CENTER,
             anchor=MSO_ANCHOR.MIDDLE, margin=(0.08, 0.02, 0.08, 0.02))
    link(sh, target)
    return sh


def line(slide, x1, y1, x2, y2, color=BORDER, width=1.25, arrow=False, dash=False):
    c = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, I(x1), I(y1), I(x2), I(y2))
    c.line.color.rgb = color
    c.line.width = Pt(width)
    ln = c.line._get_or_add_ln()
    if dash:
        pd = etree.SubElement(ln, qn("a:prstDash"))
        pd.set("val", "dash")
    if arrow:
        te = etree.SubElement(ln, qn("a:tailEnd"))
        te.set("type", "triangle")
        te.set("w", "med")
        te.set("len", "med")
    return c


def notes(slide, t):
    slide.notes_slide.notes_text_frame.text = t.strip()


def fade(slide):
    sld = slide._element
    for old in sld.findall(qn("p:transition")):
        sld.remove(old)
    tr = etree.SubElement(sld, qn("p:transition"))
    tr.set("spd", "med")
    etree.SubElement(tr, qn("p:fade"))
    # p:transition must come right after p:clrMapOvr (before p:timing / p:extLst)
    clr = sld.find(qn("p:clrMapOvr"))
    if clr is not None:
        clr.addnext(tr)


# -------------------------------------------------------------- chrome
def nav(slide, key, section):
    # top bar
    bar = box(slide, 0, 0, SW, 0.58, fill=rgb("10121A"), line=None, shape=MSO_SHAPE.RECTANGLE)
    line(slide, 0, 0.58, SW, 0.58, color=BORDER, width=0.75)
    brand = text(slide, 0.4, 0.1, 1.9, 0.4,
                 [[("↗ ", {"color": PURPLE}), ("profit", {}), (".", {"color": PURPLE})]],
                 size=16, bold=True, font=HEAD, anchor=MSO_ANCHOR.MIDDLE)
    link(brand, "title")
    menu = box(slide, 2.25, 0.14, 1.05, 0.32, fill=PANEL2, line=rgb("3A3550"), radius=0.16)
    box_text(menu, "☰  Menu", size=10, color=SOFT, bold=True, font=HEAD, align=PP_ALIGN.CENTER,
             anchor=MSO_ANCHOR.MIDDLE, margin=(0.04, 0, 0.04, 0))
    link(menu, "agenda")
    tw = 1.48
    x0 = SW - 0.35 - tw * len(SECTIONS)
    for i, (sid, name) in enumerate(SECTIONS):
        active = sid == section
        t = text(slide, x0 + i * tw, 0.08, tw, 0.42, name, size=11.5,
                 color=TEXT if active else MUTED, bold=active, font=HEAD if active else BODY,
                 align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        link(t, SECTION_FIRST[sid])
        if active:
            u = box(slide, x0 + i * tw + 0.3, 0.52, tw - 0.6, 0.06, fill=PURPLE, line=None,
                    shape=MSO_SHAPE.RECTANGLE)
            link(u, SECTION_FIRST[sid])
    # bottom bar
    idx = KEYS.index(key)
    text(slide, 0.4, 7.02, 6, 0.3,
         [[("Profit Portfolio Analyzer", {"color": SOFT, "bold": True}),
           ("   ·   " + SECTION_NAME.get(section, "Menu"), {})]],
         size=10, color=MUTED, anchor=MSO_ANCHOR.MIDDLE)
    text(slide, SW - 2.55, 7.02, 1.0, 0.3, f"{idx + 1:02d} / {len(KEYS):02d}", size=10,
         color=MUTED, align=PP_ALIGN.RIGHT, anchor=MSO_ANCHOR.MIDDLE)
    if idx > 0:
        button(slide, SW - 1.45, 6.98, 0.48, 0.36, "‹", KEYS[idx - 1], size=14)
    if idx < len(KEYS) - 1:
        button(slide, SW - 0.9, 6.98, 0.48, 0.36, "›", KEYS[idx + 1], primary=True, size=14)


def heading(slide, eyebrow, title, sub=None, y=0.85, color=PURPLE):
    text(slide, 0.6, y, 9, 0.3, eyebrow, size=11, color=color, bold=True, font=HEAD)
    text(slide, 0.6, y + 0.28, 12.2, 0.7,
         [[(title, {}), (".", {"color": color})]], size=30, bold=True, font=HEAD)
    if sub:
        text(slide, 0.6, y + 0.98, 12.2, 0.4, sub, size=14, color=MUTED)


# --------------------------------------------------------------- charts
def chart_base(chart, legend=False, size=10):
    chart.has_title = False
    chart.font.size = Pt(size)
    chart.font.color.rgb = SOFT
    chart.font.name = BODY
    cs = chart._chartSpace
    for old in cs.findall(qn("c:spPr")):
        cs.remove(old)
    sp = etree.Element(qn("c:spPr"))
    etree.SubElement(sp, qn("a:noFill"))
    ln = etree.SubElement(sp, qn("a:ln"))
    etree.SubElement(ln, qn("a:noFill"))
    cs.find(qn("c:chart")).addnext(sp)
    pa = cs.find(qn("c:chart")).find(qn("c:plotArea"))
    if pa.find(qn("c:spPr")) is None:
        sp2 = etree.Element(qn("c:spPr"))
        etree.SubElement(sp2, qn("a:noFill"))
        ext = pa.find(qn("c:extLst"))
        if ext is not None:
            ext.addprevious(sp2)
        else:
            pa.append(sp2)
    chart.has_legend = legend
    if legend:
        chart.legend.position = XL_LEGEND_POSITION.BOTTOM
        chart.legend.include_in_layout = False
        chart.legend.font.size = Pt(size)
        chart.legend.font.color.rgb = SOFT


def axes(chart, value_fmt=None, low_labels=False, gridlines=True, cat_fmt=None):
    ca, va = chart.category_axis, chart.value_axis
    for ax in (ca, va):
        ax.tick_labels.font.size = Pt(10)
        ax.tick_labels.font.color.rgb = MUTED
        ax.major_tick_mark = XL_TICK_MARK.NONE
        ax.minor_tick_mark = XL_TICK_MARK.NONE
    ca.format.line.color.rgb = BORDER
    va.format.line.fill.background()
    va.has_major_gridlines = gridlines
    if gridlines:
        va.major_gridlines.format.line.color.rgb = GRID
        va.major_gridlines.format.line.width = Pt(0.75)
    if value_fmt:
        va.tick_labels.number_format = value_fmt
        va.tick_labels.number_format_is_linked = False
    if cat_fmt:
        ca.tick_labels.number_format = cat_fmt
        ca.tick_labels.number_format_is_linked = False
    if low_labels:
        ca.tick_label_position = XL_TICK_LABEL_POSITION.LOW


def color_series(series, color):
    series.format.fill.solid()
    series.format.fill.fore_color.rgb = color
    series.format.line.fill.background()


def add_chart(slide, kind, x, y, w, h, data):
    return slide.shapes.add_chart(kind, I(x), I(y), I(w), I(h), data).chart


# ------------------------------------------------------------ code blocks
JS_KW = r"\b(function|const|let|return|if|else|new|for|of|map|forEach|true|false)\b"


def code_block(slide, x, y, w, h, src, lang="js", size=11, title=None):
    sh = box(slide, x, y, w, h, fill=CODEBG, line=BORDER, radius=0.08)
    tf = sh.text_frame
    tf.auto_size = MSO_AUTO_SIZE.NONE
    tf.vertical_anchor = MSO_ANCHOR.TOP
    tf.word_wrap = True
    tf.margin_left, tf.margin_top, tf.margin_right, tf.margin_bottom = I(0.22), I(0.16), I(0.15), I(0.12)
    lines_ = src.strip("\n").split("\n")
    if title:
        lines_ = ["\x00" + title] + lines_
    first = True
    for ln_ in lines_:
        p = tf.paragraphs[0] if first else tf.add_paragraph()
        first = False
        if ln_.startswith("\x00"):
            r = p.add_run()
            r.text = ln_[1:]
            _apply_run(r, size - 1, MUTED, True, HEAD)
            p.space_after = Pt(4)
            continue
        for tok, col in highlight(ln_, lang):
            r = p.add_run()
            r.text = tok
            _apply_run(r, size, col, False, MONO)
        if not p.runs:
            r = p.add_run()
            r.text = " "
            _apply_run(r, size, TEXT, False, MONO)
    return sh


def highlight(line_, lang):
    if lang == "yaml":
        m = re.match(r"^(\s*)(#.*)$", line_)
        if m:
            return [(m.group(1), TEXT), (m.group(2), MUTED)]
        m = re.match(r"^(\s*)(- )?([A-Za-z_]+)(:)(.*)$", line_)
        if m:
            out = [(m.group(1), TEXT)]
            if m.group(2):
                out.append((m.group(2), MUTED))
            out += [(m.group(3), PURPLE), (m.group(4), MUTED), (m.group(5), MINT)]
            return out
        m = re.match(r"^(\s*)(- )(.*)$", line_)
        if m:
            return [(m.group(1), TEXT), (m.group(2), MUTED), (m.group(3), TEXT)]
        return [(line_, TEXT)]
    # js
    out = []
    pos = 0
    pattern = re.compile(r"(//.*$)|('[^']*')|" + JS_KW + r"|(\b\d+\b)")
    for m in pattern.finditer(line_):
        if m.start() > pos:
            out.append((line_[pos:m.start()], TEXT))
        if m.group(1):
            out.append((m.group(1), MUTED))
        elif m.group(2):
            out.append((m.group(2), MINT))
        elif m.group(3):
            out.append((m.group(3), PURPLE))
        else:
            out.append((m.group(0), AMBER))
        pos = m.end()
    if pos < len(line_):
        out.append((line_[pos:], TEXT))
    return out


# ====================================================================== DATA
# Real numbers computed from sample-data/superstore.csv (9,994 order lines)
SUBCATS = {"Bookcases": (114880, -3473), "Chairs": (328449, 26590), "Labels": (12486, 5546),
           "Tables": (206966, -17725), "Storage": (223844, 21279), "Furnishings": (91705, 13059),
           "Art": (27119, 6528), "Phones": (330007, 44516), "Binders": (203413, 30222),
           "Appliances": (107532, 18138), "Paper": (78479, 34054), "Accessories": (167380, 41937),
           "Envelopes": (16476, 6964), "Fasteners": (3024, 950), "Supplies": (46674, -1189),
           "Machines": (189239, 3385), "Copiers": (149528, 55618)}
MEAN_SUBCAT_SALES = 135129
QUARTERS = [("2014 Q1", 74448, 3811), ("2014 Q2", 86539, 11204), ("2014 Q3", 143633, 12805),
            ("2014 Q4", 179628, 21724), ("2015 Q1", 68852, 9265), ("2015 Q2", 89124, 12191),
            ("2015 Q3", 130260, 16854), ("2015 Q4", 182297, 23309), ("2016 Q1", 93237, 11441),
            ("2016 Q2", 136082, 16390), ("2016 Q3", 143787, 15824), ("2016 Q4", 236099, 38140),
            ("2017 Q1", 123145, 23506), ("2017 Q2", 133764, 15499), ("2017 Q3", 196252, 26985),
            ("2017 Q4", 280054, 27449)]
CAT_REGION = {("Furniture", "Central"): -2871, ("Furniture", "East"): 3046, ("Furniture", "South"): 6771,
              ("Furniture", "West"): 11505, ("Office Supplies", "Central"): 8880,
              ("Office Supplies", "East"): 41015, ("Office Supplies", "South"): 19986,
              ("Office Supplies", "West"): 52610, ("Technology", "Central"): 33697,
              ("Technology", "East"): 47462, ("Technology", "South"): 19992, ("Technology", "West"): 44304}
CAT_YEAR = {("Furniture", "2016"): 6960, ("Furniture", "2017"): 3018, ("Office Supplies", "2016"): 35061,
            ("Office Supplies", "2017"): 39737, ("Technology", "2016"): 39774, ("Technology", "2017"): 50684}
STATES = [("California", 76381), ("New York", 74039), ("Washington", 33403),
          ("Pennsylvania", -15560), ("Ohio", -16971), ("Texas", -25729)]
DISCOUNTS = [("0%", 66.9), ("1–20%", 26.5), ("21–40%", -77.9), ("41–60%", -134.6), (">60%", -98.3)]


def quadrant(sales, profit):
    if profit >= 0:
        return "star" if sales >= MEAN_SUBCAT_SALES else "cow"
    return "question" if sales < MEAN_SUBCAT_SALES else "loss"


# ===================================================================== SLIDES
for key, section in ORDER:
    set_bg(S[key])
    fade(S[key])
    if section:
        nav(S[key], key, section)

# ---------------------------------------------------------------- 1. title
s = S["title"]
glow = box(s, 8.2, -2.2, 7.5, 7.5, fill=rgb("151129"), line=None, shape=MSO_SHAPE.OVAL)
text(s, 0.75, 0.55, 4, 0.4, [[("↗ ", {"color": PURPLE}), ("profit", {}), (".", {"color": PURPLE})]],
     size=20, bold=True, font=HEAD)
text(s, 0.75, 1.75, 6.4, 0.3, "PORTFOLIO PROJECT  ·  WEB APP  ·  HOSTED ON AWS", size=11,
     color=PURPLE, bold=True, font=HEAD)
text(s, 0.75, 2.1, 6.6, 1.6, [[("Profit Portfolio", {})], [("Analyzer", {}), (".", {"color": PURPLE})]],
     size=44, bold=True, font=HEAD, spacing=0.95)
text(s, 0.75, 3.85, 6.0, 1.0,
     "A privacy-first analytics app that turns any sales CSV into 60+ interactive charts, "
     "side-by-side comparisons and PDF reports. Everything runs in the browser.",
     size=15, color=SOFT, spacing=1.1)
button(s, 0.75, 5.15, 2.25, 0.5, "▶  Start the tour", "agenda", primary=True, size=13)
button(s, 3.15, 5.15, 1.85, 0.5, "Live app  ↗", LIVE_URL, size=12)
button(s, 5.15, 5.15, 1.75, 0.5, "GitHub  ↗", GITHUB_URL, size=12)
text(s, 0.75, 6.55, 6.5, 0.3, "HTML  ·  CSS  ·  JavaScript  ·  Chart.js  ·  Papa Parse  ·  jsPDF  ·  AWS Amplify",
     size=10.5, color=MUTED)

# mock dashboard
p = box(s, 7.55, 1.25, 5.2, 5.25, fill=PANEL, line=BORDER, radius=0.14)
text(s, 7.85, 1.45, 3.4, 0.3, [[("▤  ", {"color": PURPLE}), ("Sample - Superstore.csv", {"bold": True})]],
     size=11.5)
pill = box(s, 11.35, 1.47, 1.1, 0.27, fill=DARKPURPLE, line=None, radius=0.13)
box_text(pill, "YOUR DATA", size=8.5, color=PURPLE, bold=True, font=HEAD, align=PP_ALIGN.CENTER,
         anchor=MSO_ANCHOR.MIDDLE, margin=(0.02, 0, 0.02, 0))
text(s, 7.85, 1.78, 4.5, 0.25, "9,994 transactions  ·  2014–2017", size=9.5, color=MUTED)
for i, (lbl, val, col) in enumerate([("Total revenue", "$2.30M", PURPLE), ("Total profit", "$286.4K", MINT),
                                       ("Profit margin", "12.5%", SKY)]):
    k = box(s, 7.85 + i * 1.6, 2.2, 1.48, 1.05, fill=PANEL2, line=BORDER, radius=0.08)
    box_text(k, [[(lbl, {"size": 9, "color": MUTED})], [(val, {"size": 19, "bold": True, "font": HEAD, "color": TEXT})]],
             margin=(0.12, 0.1, 0.08, 0.06))
    box(s, 7.97 + i * 1.6, 3.1, 0.5, 0.05, fill=col, line=None, shape=MSO_SHAPE.RECTANGLE)
text(s, 7.85, 3.45, 4.6, 0.25, "Profit by quarter", size=10.5, bold=True, font=HEAD)
cd = CategoryChartData()
cd.categories = [q[0].replace(" ", "\n") if q[0].endswith("Q1") else "" for q in QUARTERS]
cd.add_series("Profit", [q[2] for q in QUARTERS])
ch = add_chart(s, XL_CHART_TYPE.COLUMN_CLUSTERED, 7.7, 3.7, 4.9, 2.7, cd)
chart_base(ch)
axes(ch, value_fmt='$#,##0,"K"')
ch.plots[0].gap_width = 45
color_series(ch.plots[0].series[0], PURPLE)
ch.category_axis.tick_labels.font.size = Pt(8)
notes(s, """
Hi, I'm presenting the Profit Portfolio Analyzer, a web app I built and deployed on AWS.
In one sentence: you drop in a sales CSV, and within seconds you get a full profitability analysis with more than 60 interactive charts, side-by-side comparisons and a PDF report. Nothing is uploaded to a server; it all runs in your browser.
The dashboard on the right uses real numbers from the Superstore demo dataset the app ships with.
This deck is interactive: click Start the tour, use the tabs at the top to jump to any section, or click the cards to drill down.
""")

# ---------------------------------------------------------------- 2. agenda
s = S["agenda"]
nav(s, "agenda", None)
heading(s, "INTERACTIVE MENU", "Where would you like to go", "Click any card to jump straight to that part of the story.")
cards = [
    ("01", "The project", "The problem I wanted to solve, and the idea behind the app.", "problem", PURPLE),
    ("02", "What the site can do", "Eight feature areas, each with a live example from real data.", "hub", MINT),
    ("03", "How it works", "Client-side architecture, the data pipeline and the modular code design.", "architecture", SKY),
    ("04", "How each file works", "index.html, style.css, script.js, amplify.yml and the sample data.", "tree", AMBER),
    ("05", "How I deployed it", "GitHub to AWS Amplify, automatic CI/CD, and debugging the first build.", "deployflow", ROSE),
    ("06", "Why it's different", "Compared with spreadsheets and BI tools, plus the skills I used.", "compare", CYAN),
]
cw, chh = 3.9, 1.84
for i, (num, title, desc, tgt, col) in enumerate(cards):
    cx = 0.6 + (i % 3) * (cw + 0.22)
    cy = 2.55 + (i // 3) * (chh + 0.22)
    c = box(s, cx, cy, cw, chh, fill=PANEL, line=BORDER, radius=0.12)
    box_text(c, [[(num, {"size": 12, "color": col, "bold": True, "font": HEAD, "after": 6})],
                 [(title, {"size": 18, "bold": True, "font": HEAD, "after": 4})],
                 [(desc, {"size": 12, "color": MUTED, "after": 8})],
                 [("Open  →", {"size": 11, "color": col, "bold": True, "font": HEAD})]],
             margin=(0.28, 0.24, 0.25, 0.15))
    stripe = box(s, cx + 0.28, cy, 0.9, 0.06, fill=col, line=None, shape=MSO_SHAPE.RECTANGLE)
    link([c, stripe], tgt)
text(s, 0.6, 6.55, 12, 0.3,
     [[("Tip: ", {"bold": True, "color": SOFT}),
       ("this works in Slide Show mode (F5). The tabs at the top and the ‹ › buttons at the bottom are clickable on every slide.", {})]],
     size=10.5, color=MUTED)
notes(s, """
This is the menu for the presentation. I can go through it start to finish, or jump straight to whatever you're most interested in: the features, the code, the AWS deployment, or how it compares to other tools.
""")

# ---------------------------------------------------------------- 3. problem
s = S["problem"]
heading(s, "01  ·  THE PROJECT", "Sales data is everywhere. Insight isn't",
        "Most small businesses and analysts already have the data. Turning it into decisions is the hard part.")
pains = [
    ("Spreadsheets are slow", "Pivot tables, manual charts and formulas that break as soon as the columns change. "
                              "Every new question means rebuilding the analysis.", AMBER, "01"),
    ("BI tools are heavy", "Installs, licences, accounts, data models and a learning curve, all before you see "
                           "your first chart. Too much for a quick question.", SKY, "02"),
    ("Privacy is a real risk", "Many online dashboard tools require you to upload confidential sales and customer "
                               "data to a third-party server.", ROSE, "03"),
]
for i, (t, d, col, n) in enumerate(pains):
    cx = 0.6 + i * 4.12
    c = box(s, cx, 2.55, 3.9, 2.5, fill=PANEL, line=BORDER, radius=0.12)
    box_text(c, [[(n, {"size": 26, "bold": True, "font": HEAD, "color": col, "after": 4})],
                 [(t, {"size": 17, "bold": True, "font": HEAD, "after": 6})],
                 [(d, {"size": 12.5, "color": MUTED})]], margin=(0.28, 0.22, 0.25, 0.15), spacing=1.05)
q = box(s, 0.6, 5.35, 12.13, 1.25, fill=rgb("161227"), line=rgb("3A3060"), radius=0.12)
box_text(q, [[("The question I asked:  ", {"color": PURPLE, "bold": True, "font": HEAD}),
              ("What if you could drop in a CSV and get a full profitability analysis in seconds, "
               "without your data ever leaving your computer?", {"color": TEXT})]],
         size=16, anchor=MSO_ANCHOR.MIDDLE, margin=(0.35, 0.1, 0.35, 0.1))
notes(s, """
I started from a simple observation. Every business exports sales data as CSV files, but getting answers out of them is painful.
Spreadsheets need manual work for every question. BI tools like Power BI or Tableau are powerful but heavy for a quick analysis. And many online tools make you upload sensitive data.
So my goal was: instant, private, profit-focused analysis with zero setup.
""")

# ---------------------------------------------------------------- 4. solution
s = S["solution"]
heading(s, "01  ·  THE PROJECT", "One page. Any CSV. Instant answers",
        "Profit Portfolio Analyzer is a single-page web app. The whole workflow takes five steps.")
steps = [("Upload", "Drop one or more CSV files"), ("Map", "Columns are auto-detected"),
         ("Explore", "KPIs, quadrants, 60+ charts"), ("Compare", "Files side by side"),
         ("Export", "One-click PDF report")]
cols = [PURPLE, MINT, SKY, AMBER, ROSE]
for i, (t, d) in enumerate(steps):
    cx = 0.6 + i * 2.48
    c = box(s, cx, 2.55, 2.12, 1.45, fill=PANEL, line=BORDER, radius=0.1)
    box_text(c, [[(f"STEP {i + 1}", {"size": 10, "bold": True, "color": cols[i], "font": HEAD, "after": 3})],
                 [(t, {"size": 18, "bold": True, "font": HEAD, "after": 2})],
                 [(d, {"size": 11.5, "color": MUTED})]], margin=(0.2, 0.16, 0.12, 0.1))
    if i < 4:
        text(s, cx + 2.12, 2.95, 0.36, 0.5, "→", size=20, color=MUTED, align=PP_ALIGN.CENTER,
             anchor=MSO_ANCHOR.MIDDLE)
stats = [("60+", "interactive charts and views", PURPLE), ("0", "bytes of your data sent to a server", MINT),
         ("3", "core files, with no build step and no backend", SKY), ("9,994", "order lines in the bundled demo dataset", AMBER)]
for i, (big, small, col) in enumerate(stats):
    cx = 0.6 + i * 3.08
    c = box(s, cx, 4.35, 2.9, 2.2, fill=PANEL, line=BORDER, radius=0.12)
    box_text(c, [[(big, {"size": 40, "bold": True, "font": HEAD, "color": col})],
                 [(small, {"size": 13, "color": SOFT})]], margin=(0.28, 0.26, 0.2, 0.15))
notes(s, """
The solution is a single web page. You upload a CSV; the app recognises the columns automatically; you explore a full dashboard; you can compare several files; and you export a PDF.
Some numbers: more than 60 chart types, zero bytes uploaded, and the whole site is just three files with no backend at all, which makes it fast, private and almost free to host.
""")

# ---------------------------------------------------------------- 5. hub
s = S["hub"]
heading(s, "02  ·  WHAT THE SITE CAN DO", "Eight things the site can do",
        "Click a tile to open that feature. Every detail slide has a button back to this hub.")
feats = [
    ("01", "Smart CSV import", "Auto-detects columns and cleans messy numbers and dates", "f1", PURPLE),
    ("02", "Portfolio overview", "KPIs, quadrant bubble chart, top and bottom 10", "f2", MINT),
    ("03", "Slice & filter", "Group by any level, plus 5 filters that apply everywhere", "f3", SKY),
    ("04", "Compare portfolios", "Multi-file comparison and A/B side-by-side", "f4", AMBER),
    ("05", "Geography", "World map, US states and city bubbles", "f5", ROSE),
    ("06", "Time intelligence", "Trends, growth, seasonality and rankings", "f6", CYAN),
    ("07", "Customer intelligence", "RFM segments, cohorts and Pareto", "f7", LIME),
    ("08", "Diagnostics & PDF", "Discounts, flows, outliers and reports", "f8", PINK),
]
tw_, th_ = 2.88, 1.85
for i, (n, t, d, tgt, col) in enumerate(feats):
    cx = 0.6 + (i % 4) * (tw_ + 0.2)
    cy = 2.5 + (i // 4) * (th_ + 0.2)
    c = box(s, cx, cy, tw_, th_, fill=PANEL, line=BORDER, radius=0.12)
    box_text(c, [[(n, {"size": 11, "color": col, "bold": True, "font": HEAD, "after": 4})],
                 [(t, {"size": 16, "bold": True, "font": HEAD, "after": 3})],
                 [(d, {"size": 11.5, "color": MUTED})]], margin=(0.22, 0.2, 0.18, 0.1))
    stripe = box(s, cx + 0.22, cy, 0.7, 0.06, fill=col, line=None, shape=MSO_SHAPE.RECTANGLE)
    link([c, stripe], tgt)
b = button(s, 9.55, 6.43, 3.18, 0.42, "See it on real data  →", "demo", primary=True, size=12)
text(s, 0.6, 6.45, 8.5, 0.38, "Every number in these slides comes from the Superstore dataset bundled with the app.",
     size=11, color=MUTED, anchor=MSO_ANCHOR.MIDDLE)
notes(s, """
Here are the eight feature areas. I'll click into whichever is most relevant. Each detail slide shows a real chart made from the Superstore dataset, the same data you get when you click Load Superstore data on the live site.
""")


# ------------------------------------------------------ feature detail slides
def feature_slide(key, idx, eyebrow_color, title, sub, bullets, under_hood, visual_title, visual_caption):
    s = S[key]
    heading(s, f"FEATURE {idx:02d} / 08", title, sub, color=eyebrow_color)
    text(s, 0.6, 2.45, 5.6, 0.3, "WHAT IT DOES", size=10.5, color=MUTED, bold=True, font=HEAD)
    text(s, 0.6, 2.78, 5.6, 2.6, bullets, size=12.5, color=TEXT, bullet="•", bullet_color=eyebrow_color,
         space_after=5, spacing=1.0)
    uh = box(s, 0.6, 5.45, 5.6, 1.15, fill=CODEBG, line=BORDER, radius=0.08)
    box_text(uh, [[("UNDER THE HOOD", {"size": 9.5, "bold": True, "color": eyebrow_color, "font": HEAD, "after": 3})],
                  [(under_hood, {"size": 10.5, "color": SOFT})]], margin=(0.2, 0.12, 0.18, 0.08))
    panel = box(s, 6.55, 2.45, 6.18, 4.15, fill=PANEL, line=BORDER, radius=0.12)
    text(s, 6.8, 2.6, 5.7, 0.3, visual_title, size=13, bold=True, font=HEAD)
    text(s, 6.8, 2.9, 5.7, 0.28, visual_caption, size=10, color=MUTED)
    # nav buttons
    back = button(s, 3.3, 6.98, 1.75, 0.36, "←  All features", "hub", size=10.5)
    nxt_key = f"f{idx + 1}" if idx < 8 else "demo"
    nxt_label = "Next feature  →" if idx < 8 else "Demo results  →"
    button(s, 5.15, 6.98, 1.85, 0.36, nxt_label, nxt_key, size=10.5, color=eyebrow_color)
    return s


# F1 ------------------------------------------------------------------------
s = feature_slide("f1", 1, PURPLE, "Smart CSV import & column mapping",
                  "Bring any sales export. The app works out which column is which.",
                  ["Upload one or many CSVs at once. Files are queued and imported one by one.",
                   "Auto-matches headers like “Product Name”, “Revenue”, “Net Profit” or “Order Date”.",
                   "Only Product, Sales and Profit are required. Region, Date, Category, Segment and Country unlock more charts.",
                   "Cleans messy numbers: $ £ € ¥ ₹, commas, % and accounting negatives like (1,200).",
                   "Validates dates in several formats and skips bad values instead of guessing."],
                  "Papa Parse reads the file inside the browser. openMapping() pre-fills dropdowns with regex and priority "
                  "matching, then numeric() and dateKey() normalise every value.",
                  "Column mapping, automatically", "Superstore headers (left) matched to app fields (right)")
left = ["Product Name", "Sales", "Profit", "Order Date", "Region", "Category", "Sub-Category", "Segment"]
right = [("Product / Item *", True), ("Sales / Revenue *", True), ("Profit *", True), ("Order date", False),
         ("Region", False), ("Category", False), ("Sub-category", False), ("Customer segment", False)]
for i, h in enumerate(left):
    yy = 3.35 + i * 0.4
    c = box(s, 6.85, yy, 2.0, 0.32, fill=PANEL2, line=BORDER, radius=0.08)
    box_text(c, h, size=10.5, color=SOFT, font=MONO, anchor=MSO_ANCHOR.MIDDLE, margin=(0.12, 0, 0.05, 0))
    lbl, req = right[i]
    c2 = box(s, 10.35, yy, 2.15, 0.32, fill=DARKPURPLE if req else PANEL2,
             line=PURPLE if req else BORDER, radius=0.08)
    box_text(c2, lbl, size=10.5, color=PURPLE if req else TEXT, bold=req, anchor=MSO_ANCHOR.MIDDLE,
             margin=(0.12, 0, 0.05, 0))
    line(s, 8.88, yy + 0.16, 10.32, yy + 0.16, color=PURPLE if req else rgb("3A3F52"), width=1.5 if req else 1,
         arrow=True)
text(s, 6.85, 6.6 - 0.35, 5.7, 0.3, [[("* required   ", {"color": PURPLE, "bold": True}),
                                       ("Every other field is optional and unlocks extra charts.", {})]],
     size=10, color=MUTED)
notes(s, """
Feature one is the import. Real-world CSVs are messy, so the app auto-detects which column is product, sales, profit, date and so on, using pattern matching with priorities.
Only three fields are required. The optional ones unlock more charts: a date column unlocks all the time charts, a customer ID unlocks RFM and cohorts.
It also cleans currency symbols, thousands separators and accounting-style negatives, and it never guesses invalid dates.
""")

# F2 ------------------------------------------------------------------------
s = feature_slide("f2", 2, MINT, "Portfolio overview",
                  "The big picture on one screen: what makes money and what loses it.",
                  ["KPI cards for total revenue, total profit and profit margin.",
                   "A quadrant bubble chart sorts every product into Stars, Cash Cows, Question Marks or Loss-Makers.",
                   "Top 10 “profit leaders” and bottom 10 “room to improve”.",
                   "Profit by region and by month.",
                   "Auto-written insights: your best performer, your profit leaks and your next opportunity."],
                  "Quadrants split at average revenue and zero profit:  profit ≥ 0 → Star (above avg) or Cash Cow;  "
                  "profit < 0 → Loss-Maker (above avg) or Question Mark.",
                  "Superstore portfolio quadrants", "17 sub-categories  ·  bubble size = revenue  ·  split at $135K average revenue")
groups = {"star": ("Stars", PURPLE), "cow": ("Cash Cows", MINT), "question": ("Question Marks", AMBER),
          "loss": ("Loss-Makers", ROSE)}
bd = BubbleChartData()
series_points = {}
for g, (name, col) in groups.items():
    ser = bd.add_series(name)
    series_points[g] = []
    for sc, (sa, pr) in SUBCATS.items():
        if quadrant(sa, pr) == g:
            ser.add_data_point(sa / 1000, pr / 1000, sa)
            series_points[g].append(sc)
ch = add_chart(s, XL_CHART_TYPE.BUBBLE, 6.7, 3.2, 5.9, 3.35, bd)
chart_base(ch, legend=True, size=9.5)
axes(ch, value_fmt='$#,##0"K"', cat_fmt='$#,##0"K"')
ch.category_axis.has_major_gridlines = False
ch.plots[0].bubble_scale = 35
label_these = {"Copiers", "Phones", "Tables", "Bookcases", "Supplies", "Chairs", "Paper"}
for i, (g, (name, col)) in enumerate(groups.items()):
    ser = ch.plots[0].series[i]
    ser.format.fill.solid()
    ser.format.fill.fore_color.rgb = col
    ser.format.line.color.rgb = col
    for j, sc in enumerate(series_points[g]):
        if sc in label_these:
            dl = ser.points[j].data_label
            dl.has_text_frame = True
            dl.text_frame.text = sc
            dl.position = XL_LABEL_POSITION.RIGHT if sc not in ("Phones", "Chairs") else XL_LABEL_POSITION.ABOVE
            dl.font.size = Pt(9)
            dl.font.color.rgb = SOFT
notes(s, """
The overview is the heart of the app. The quadrant chart works like a BCG matrix for profitability: x is revenue, y is profit, and the split is at average revenue and zero profit.
On Superstore you can immediately see Tables as a Loss-Maker: $207K in sales but $17.7K in losses. Copiers and Phones are the Stars.
The app also writes three insights automatically: back your best performer, stop the profit leaks, and find your next opportunity.
""")

# F3 ------------------------------------------------------------------------
s = feature_slide("f3", 3, SKY, "Slice & filter",
                  "Ask follow-up questions with a click, without rebuilding anything.",
                  ["Analyze by Product, Category or Sub-category.",
                   "Filter by Category, Sub-category, Customer segment, Region and Period.",
                   "Every chart, insight and the PDF report respects the active filters.",
                   "One-click reset back to the full portfolio.",
                   "Dropdowns fill themselves from the values in your data."],
                  "A single filteredRows() function is the source of truth. Every view (core and add-on modules) "
                  "reads from it, so all charts stay consistent.",
                  "Profit by region, split by category", "Superstore  ·  Furniture in Central is the only combination losing money (−$2.9K)")
cd = CategoryChartData()
regions = ["Central", "East", "South", "West"]
cd.categories = regions
for cat in ["Furniture", "Office Supplies", "Technology"]:
    cd.add_series(cat, [CAT_REGION[(cat, r)] / 1000 for r in regions])
ch = add_chart(s, XL_CHART_TYPE.COLUMN_CLUSTERED, 6.7, 3.2, 5.9, 3.35, cd)
chart_base(ch, legend=True)
axes(ch, value_fmt='$#,##0"K"', low_labels=True)
ch.plots[0].gap_width = 70
ch.plots[0].overlap = -10
for i, col in enumerate([AMBER, PURPLE, MINT]):
    ser = ch.plots[0].series[i]
    color_series(ser, col)
    ser.invert_if_negative = False
pt = ch.plots[0].series[0].points[0]
pt.format.fill.solid()
pt.format.fill.fore_color.rgb = ROSE
notes(s, """
Filters make the dashboard conversational. Pick a category, a segment, a region or a year, and every chart, insight and even the PDF export updates together.
For example, splitting by region and category shows Furniture in the Central region is the only combination that actually loses money.
Technically there's one filtering function that every module reads from, so the views never disagree.
""")

# F4 ------------------------------------------------------------------------
s = feature_slide("f4", 4, AMBER, "Compare portfolios",
                  "Put files side by side. They're never merged, so comparisons stay fair.",
                  ["A data library in the sidebar: open, check and compare any number of files.",
                   "The comparison view shows totals, top groups, regions and months, with one colour per file.",
                   "Side-by-side A/B mode uses the same scale on both sides and calculates the deltas (B − A).",
                   "Swap sides, or switch files, straight from the dropdowns.",
                   "Missing groups are shown as gaps, never filled in."],
                  "renderComparison() and renderPair() analyse each dataset separately with the same filters. "
                  "Shared axis bounds keep both sides visually comparable.",
                  "Example: 2016 export vs 2017 export", "Profit by category  ·  automatic deltas, as shown in the app")
cd = CategoryChartData()
cats = ["Furniture", "Office Supplies", "Technology"]
cd.categories = cats
cd.add_series("2016 (A)", [CAT_YEAR[(c, "2016")] / 1000 for c in cats])
cd.add_series("2017 (B)", [CAT_YEAR[(c, "2017")] / 1000 for c in cats])
ch = add_chart(s, XL_CHART_TYPE.COLUMN_CLUSTERED, 6.7, 3.2, 3.9, 3.35, cd)
chart_base(ch, legend=True)
axes(ch, value_fmt='$#,##0"K"')
ch.plots[0].gap_width = 60
color_series(ch.plots[0].series[0], SKY)
color_series(ch.plots[0].series[1], AMBER)
for i, (lbl, val, col) in enumerate([("Revenue", "+$124,009", MINT), ("Profit", "+$11,644", MINT),
                                       ("Margin", "−0.7 pp", ROSE)]):
    c = box(s, 10.75, 3.3 + i * 1.08, 1.78, 0.95, fill=PANEL2, line=BORDER, radius=0.08)
    box_text(c, [[(lbl + " (B − A)", {"size": 9, "color": MUTED})],
                 [(val, {"size": 17, "bold": True, "font": HEAD, "color": col})]],
             margin=(0.14, 0.1, 0.08, 0.05))
notes(s, """
Comparison is something spreadsheets make painful. You can load several files, say one per year or one per store, and compare them.
Files are never merged: each one is analysed separately with the same filters and the same scales, and the app calculates the deltas.
Here, 2017 versus 2016 shows revenue up $124K and profit up $11.6K, but margin down 0.7 points. That's the kind of nuance it surfaces automatically.
""")

# F5 ------------------------------------------------------------------------
s = feature_slide("f5", 5, ROSE, "Geography",
                  "See where the business happens, from continents down to cities.",
                  ["World bubble map: marker size shows revenue. Scroll to zoom and drag to pan.",
                   "A US state choropleth coloured by profit.",
                   "City sales bubbles using 580 built-in US city coordinates.",
                   "A Region → State → City sunburst.",
                   "Unrecognised regions are excluded and listed, never guessed."],
                  "World and US TopoJSON atlases are embedded in script.js (~245 KB), so maps render offline via "
                  "chartjs-chart-geo with custom zoom and pan controls.",
                  "Most and least profitable states", "Superstore  ·  top 3 vs bottom 3 of 49 states")
cd = CategoryChartData()
st = list(reversed(STATES))
cd.categories = [x[0] for x in st]
cd.add_series("Profit", [x[1] / 1000 for x in st])
ch = add_chart(s, XL_CHART_TYPE.BAR_CLUSTERED, 6.7, 3.2, 5.9, 3.35, cd)
chart_base(ch)
axes(ch, value_fmt='$#,##0"K"', low_labels=True)
ch.plots[0].gap_width = 55
ser = ch.plots[0].series[0]
color_series(ser, MINT)
ser.invert_if_negative = False
for i, (_, v) in enumerate(st):
    pt = ser.points[i]
    pt.format.fill.solid()
    pt.format.fill.fore_color.rgb = ROSE if v < 0 else MINT
ch.plots[0].has_data_labels = True
dls = ch.plots[0].data_labels
dls.number_format = '$#,##0.0"K"'
dls.number_format_is_linked = False
dls.font.size = Pt(9)
dls.font.color.rgb = SOFT
dls.position = XL_LABEL_POSITION.OUTSIDE_END
notes(s, """
Geography works at three levels: a world map for regions, a US state choropleth, and city bubbles.
The map data is embedded inside the JavaScript, so maps work without calling any map API.
On Superstore, California and New York lead with over $74K profit each, while Texas loses $25.7K, which is an immediate question for the business.
""")

# F6 ------------------------------------------------------------------------
s = feature_slide("f6", 6, CYAN, "Time intelligence",
                  "Trends, seasonality and growth, calculated automatically.",
                  ["Monthly sales and profit with a moving average.",
                   "Month-over-month and year-over-year growth.",
                   "A Year × Month heatmap, a daily calendar and weekday performance.",
                   "Cumulative sales and profit, plus category sales over time.",
                   "A “ranking over time” bump chart for top products."],
                  "Part 5 module: builds a continuous month timeline (gaps stay empty), calculates the growth rates, "
                  "and draws heatmaps with custom Chart.js plugins.",
                  "Profit by quarter, 2014–2017", "Superstore  ·  Q4 is the most profitable quarter in every year")
cd = CategoryChartData()
cd.categories = [q[0][-2:] + ("\n" + q[0][:4] if q[0].endswith("Q1") else "") for q in QUARTERS]
cd.add_series("Profit", [q[2] / 1000 for q in QUARTERS])
ch = add_chart(s, XL_CHART_TYPE.LINE_MARKERS, 6.7, 3.2, 5.9, 3.35, cd)
chart_base(ch)
axes(ch, value_fmt='$#,##0"K"')
ser = ch.plots[0].series[0]
ser.format.line.color.rgb = CYAN
ser.format.line.width = Pt(2.5)
ser.smooth = False
ser.marker.size = 7
ser.marker.format.fill.solid()
ser.marker.format.fill.fore_color.rgb = CYAN
ser.marker.format.line.color.rgb = CYAN
for i, q in enumerate(QUARTERS):
    if q[0].endswith("Q4"):
        pt = ser.points[i]
        pt.marker.size = 10
        pt.marker.format.fill.solid()
        pt.marker.format.fill.fore_color.rgb = AMBER
        pt.marker.format.line.color.rgb = AMBER
ch.category_axis.tick_labels.font.size = Pt(8.5)
notes(s, """
With a date column, the app unlocks a whole time-intelligence module: trends with moving averages, MoM and YoY growth, heatmaps, calendars and rank changes.
On Superstore, the quarterly view shows a clear seasonal pattern: Q4 is the most profitable quarter every single year, highlighted here in amber. A business would plan stock and marketing around that.
""")

# F7 ------------------------------------------------------------------------
s = feature_slide("f7", 7, LIME, "Customer intelligence",
                  "Know which customers drive the business, and who's slipping away.",
                  ["RFM scoring (Recency, Frequency, Monetary on a 1–5 scale).",
                   "Segments: Champions, Loyal, Recent, Developing, At risk and Dormant.",
                   "An interactive 3D RFM view and a Recency × Frequency heatmap.",
                   "A cohort retention heatmap and curves, plus new vs repeat customers.",
                   "Pareto (80/20) charts for products and customers."],
                  "Part 2 enriches each row with order and customer metrics, using percentile scores for R, F and M. "
                  "Part 6 adds cohorts and a hand-built 3D projection on a canvas.",
                  "Revenue concentration (Pareto)", "Superstore  ·  793 customers  ·  5,009 orders  ·  1,850 products")
for i, (share, label, col) in enumerate([(48.1, "of sales come from the\ntop 20% of customers", LIME),
                                          (76.8, "of sales come from the\ntop 20% of products", PURPLE)]):
    cx = 6.75 + i * 2.95
    cd = CategoryChartData()
    cd.categories = ["Top 20%", "Rest"]
    cd.add_series("Share", [share, 100 - share])
    ch = add_chart(s, XL_CHART_TYPE.DOUGHNUT, cx, 3.25, 2.7, 2.45, cd)
    chart_base(ch)
    dn = ch._chartSpace.find(".//" + qn("c:doughnutChart"))
    hs = dn.find(qn("c:holeSize"))
    if hs is None:
        hs = etree.SubElement(dn, qn("c:holeSize"))
    hs.set("val", "68")
    ser = ch.plots[0].series[0]
    for j, c_ in enumerate([col, rgb("2A2E3B")]):
        pt = ser.points[j]
        pt.format.fill.solid()
        pt.format.fill.fore_color.rgb = c_
        pt.format.line.color.rgb = PANEL
    text(s, cx, 4.2, 2.7, 0.55, f"{share:.0f}%", size=24, bold=True, font=HEAD, align=PP_ALIGN.CENTER,
         anchor=MSO_ANCHOR.MIDDLE)
    text(s, cx, 5.72, 2.7, 0.6, label.split("\n"), size=11, color=SOFT, align=PP_ALIGN.CENTER)
notes(s, """
If the file has customer and order IDs, the app turns into a customer analytics tool: RFM scoring, segments like Champions or At risk, cohort retention and Pareto analysis.
On Superstore, the top 20% of customers generate 48% of sales, and the top 20% of products generate 77%, which is very close to the classic 80/20 rule.
""")

# F8 ------------------------------------------------------------------------
s = feature_slide("f8", 8, PINK, "Diagnostics & PDF reports",
                  "Find out why profit leaks, then share the answer.",
                  ["Discount analysis: average profit by discount bucket, and loss vs non-loss orders.",
                   "A correlation matrix, pair plots and a sales × profit density view.",
                   "A profit waterfall, a Sankey flow (Region → Category → Segment), radar and parallel coordinates.",
                   "Z-score outlier detection, plus shipping days by mode and region.",
                   "A one-click PDF report with KPIs, findings and every chart, built in the browser."],
                  "jsPDF builds an A4 report from live canvases, including the active filters in the header. "
                  "The blob URL is created during the click so browsers allow the download.",
                  "Average profit per order line, by discount", "Superstore  ·  every discount above 20% loses money on average")
cd = CategoryChartData()
cd.categories = [d[0] for d in DISCOUNTS]
cd.add_series("Avg profit", [d[1] for d in DISCOUNTS])
ch = add_chart(s, XL_CHART_TYPE.COLUMN_CLUSTERED, 6.7, 3.2, 5.9, 3.35, cd)
chart_base(ch)
axes(ch, value_fmt='$#,##0', low_labels=True)
ch.plots[0].gap_width = 55
ser = ch.plots[0].series[0]
color_series(ser, MINT)
ser.invert_if_negative = False
for i, (_, v) in enumerate(DISCOUNTS):
    pt = ser.points[i]
    pt.format.fill.solid()
    pt.format.fill.fore_color.rgb = ROSE if v < 0 else MINT
ch.plots[0].has_data_labels = True
dls = ch.plots[0].data_labels
dls.number_format = '$#,##0'
dls.number_format_is_linked = False
dls.font.size = Pt(9.5)
dls.font.color.rgb = SOFT
dls.position = XL_LABEL_POSITION.OUTSIDE_END
notes(s, """
The diagnostics answer why. The strongest finding on Superstore is about discounts: with no discount, the average order line earns $67, but every discount bucket above 20% loses money, down to minus $135.
Then you can export everything, including KPIs, findings and charts with the current filters, as a PDF generated entirely in the browser.
""")

# ---------------------------------------------------------------- demo
s = S["demo"]
heading(s, "02  ·  LIVE DEMO RESULTS", "What it found in 9,994 Superstore orders",
        "Click “Load Superstore data” on the live site to reproduce every number on this slide.")
kpis = [("Revenue", "$2,297,201", PURPLE), ("Profit", "$286,397", MINT), ("Margin", "12.5%", SKY),
        ("Loss-making lines", "18.7%", ROSE)]
for i, (l_, v, col) in enumerate(kpis):
    cx = 0.6 + (i % 2) * 2.62
    cy = 2.5 + (i // 2) * 1.2
    c = box(s, cx, cy, 2.48, 1.06, fill=PANEL, line=BORDER, radius=0.1)
    box_text(c, [[(l_, {"size": 10, "color": MUTED})], [(v, {"size": 21, "bold": True, "font": HEAD})]],
             margin=(0.18, 0.12, 0.1, 0.06))
    box(s, cx + 0.18, cy + 0.92, 0.55, 0.05, fill=col, line=None, shape=MSO_SHAPE.RECTANGLE)
finds = [("Tables lose $17.7K", "despite $207K in sales. That's the biggest profit leak.", ROSE),
         ("Copiers earn the most", "with $55.6K profit from only $150K in sales.", MINT),
         ("Discounts above 20% backfire", "and average order lines turn negative.", AMBER)]
for i, (t, d, col) in enumerate(finds):
    c = box(s, 0.6, 4.95 + i * 0.57, 5.1, 0.5, fill=PANEL, line=BORDER, radius=0.08)
    box_text(c, [[("● ", {"color": col}), (t + "  ", {"bold": True, "font": HEAD}), (d, {"color": MUTED})]],
             size=11, anchor=MSO_ANCHOR.MIDDLE, margin=(0.15, 0, 0.1, 0))
panel = box(s, 6.0, 2.5, 6.73, 4.15, fill=PANEL, line=BORDER, radius=0.12)
text(s, 6.25, 2.62, 6, 0.3, "Profit by sub-category", size=13, bold=True, font=HEAD)
text(s, 6.25, 2.92, 6, 0.28, "All 17 sub-categories, sorted  ·  red = loss-making", size=10, color=MUTED)
items = sorted(SUBCATS.items(), key=lambda kv: kv[1][1])
cd = CategoryChartData()
cd.categories = [k for k, _ in items]
cd.add_series("Profit", [v[1] / 1000 for _, v in items])
ch = add_chart(s, XL_CHART_TYPE.BAR_CLUSTERED, 6.1, 3.15, 6.5, 3.45, cd)
chart_base(ch, size=9)
axes(ch, value_fmt='$#,##0"K"', low_labels=True)
ch.category_axis.tick_labels.font.size = Pt(8.5)
ch.plots[0].gap_width = 35
ser = ch.plots[0].series[0]
color_series(ser, PURPLE)
ser.invert_if_negative = False
for i, (_, v) in enumerate(items):
    pt = ser.points[i]
    pt.format.fill.solid()
    pt.format.fill.fore_color.rgb = ROSE if v[1] < 0 else (MINT if v[1] > 30000 else PURPLE)
button(s, 3.3, 6.98, 1.75, 0.36, "←  All features", "hub", size=10.5)
notes(s, """
Here's a summary of what the app finds on the bundled dataset, in seconds and with no manual work: $2.3M revenue, $286K profit, a 12.5% margin, and almost one in five order lines losing money.
The three headline findings: Tables are the biggest leak, Copiers are the most profitable, and heavy discounting destroys profit.
That's exactly the kind of decision-ready output the app was designed for.
""")

# ---------------------------------------------------------------- architecture
s = S["architecture"]
heading(s, "03  ·  HOW IT WORKS", "100% client-side architecture",
        "There's no backend, no database and no API keys. AWS serves static files, and the browser does all the work.")
# browser container
bc = box(s, 0.6, 2.5, 7.55, 4.1, fill=rgb("10121A"), line=rgb("3A3550"), radius=0.14, line_w=1.25)
text(s, 0.85, 2.6, 7, 0.32, [[("YOUR BROWSER  ", {"color": PURPLE, "bold": True, "font": HEAD}),
                               ("·  everything below runs on the user's device", {"color": MUTED})]], size=10.5)
files3 = [("index.html", "Structure: sidebar, panels and <canvas> elements", PURPLE),
          ("style.css", "Look: dark theme, grid layout, responsive design", MINT),
          ("script.js", "Logic: parsing, analytics, 60+ charts and the PDF report", SKY)]
for i, (f, d, col) in enumerate(files3):
    c = box(s, 0.85, 3.05 + i * 0.78, 3.45, 0.7, fill=PANEL, line=BORDER, radius=0.08)
    box_text(c, [[(f, {"font": MONO, "bold": True, "color": col, "size": 12})], [(d, {"size": 9.5, "color": MUTED})]],
             margin=(0.15, 0.06, 0.08, 0.04))
csv = box(s, 0.85, 5.48, 3.45, 0.92, fill=rgb("13201C"), line=rgb("2C5A4C"), radius=0.08)
box_text(csv, [[("your-sales.csv", {"font": MONO, "bold": True, "color": MINT, "size": 12})],
               [("Read locally with the File API. Never uploaded.", {"size": 9.5, "color": SOFT})]],
         margin=(0.15, 0.1, 0.08, 0.05))
eng = box(s, 4.7, 3.05, 3.2, 3.35, fill=PANEL, line=BORDER, radius=0.1)
box_text(eng, [[("IN-MEMORY ENGINE", {"size": 9.5, "bold": True, "color": SKY, "font": HEAD, "after": 6})],
               [("Parse → map → clean", {"size": 11.5, "after": 3})],
               [("Filter → aggregate", {"size": 11.5, "after": 3})],
               [("Enrich: orders, customers, RFM, growth", {"size": 11.5, "after": 3})],
               [("Render with Chart.js", {"size": 11.5, "after": 3})],
               [("Export with jsPDF", {"size": 11.5, "after": 10})],
               [("Data lives only in this tab. Close it and it's gone.", {"size": 10, "color": MUTED})]],
         margin=(0.2, 0.16, 0.15, 0.1))
line(s, 4.32, 3.4, 4.68, 3.4, color=SKY, arrow=True)
line(s, 4.32, 5.9, 4.68, 5.9, color=MINT, arrow=True)
# right side services
aws = box(s, 8.75, 2.5, 3.98, 1.85, fill=rgb("1E1A10"), line=rgb("6B5524"), radius=0.12)
box_text(aws, [[("AWS AMPLIFY HOSTING", {"size": 10, "bold": True, "color": AMBER, "font": HEAD, "after": 4})],
               [("Serves index.html, style.css, script.js and the sample CSV over HTTPS from a global CDN.",
                 {"size": 11, "color": SOFT, "after": 4})],
               [("Static files only, with no server code.", {"size": 10, "color": MUTED})]],
         margin=(0.2, 0.15, 0.15, 0.1))
cdn = box(s, 8.75, 4.55, 3.98, 2.05, fill=PANEL, line=BORDER, radius=0.12)
box_text(cdn, [[("OPEN-SOURCE LIBRARIES (jsDelivr CDN)", {"size": 10, "bold": True, "color": PURPLE, "font": HEAD, "after": 5})],
               [("Papa Parse 5.4.1", {"bold": True}), ("  CSV parsing", {"color": MUTED})],
               [("Chart.js 4.4.8", {"bold": True}), ("  charts", {"color": MUTED})],
               [("chartjs-chart-geo 4.3.6", {"bold": True}), ("  maps", {"color": MUTED})],
               [("jsPDF 2.5.2", {"bold": True}), ("  PDF reports", {"color": MUTED})],
               [("Google Fonts", {"bold": True}), ("  DM Sans & Manrope", {"color": MUTED})]],
         size=11, margin=(0.2, 0.15, 0.15, 0.08), space_after=2)
line(s, 8.72, 3.4, 8.18, 3.4, color=AMBER, arrow=True)
line(s, 8.72, 5.55, 8.18, 5.55, color=PURPLE, arrow=True)
notes(s, """
Architecturally, this is a pure client-side app. AWS Amplify just serves static files. Libraries like Chart.js and Papa Parse load from the jsDelivr CDN.
When a user selects a CSV, it's read locally through the browser's File API, then parsed, analysed and charted in memory. No request ever carries the user's data.
Benefits: privacy by design, instant performance, nothing to scale, and hosting costs close to zero.
""")

# ---------------------------------------------------------------- pipeline
s = S["pipeline"]
heading(s, "03  ·  HOW IT WORKS", "The data pipeline, step by step",
        "From a raw CSV to charts in six functions.")
pipe = [("1  Read", "Papa.parse(file)", "Turns CSV text into row objects", PURPLE),
        ("2  Map", "openMapping()", "Matches headers to fields", MINT),
        ("3  Clean", "numeric() · dateKey()", "Numbers, negatives, months", SKY),
        ("4  Filter", "filteredRows()", "Category, segment, region, period", AMBER),
        ("5  Aggregate", "aggregate(rows, key)", "Sales, profit, count by group", ROSE),
        ("6  Render", "buildChart() · jsPDF", "Charts, insights, PDF", CYAN)]
for i, (t, fn, d, col) in enumerate(pipe):
    cx = 0.6 + i * 2.05
    c = box(s, cx, 2.5, 1.88, 1.55, fill=PANEL, line=BORDER, radius=0.1)
    box_text(c, [[(t, {"size": 14, "bold": True, "font": HEAD, "color": col, "after": 4})],
                 [(fn, {"size": 9.5, "font": MONO, "color": TEXT, "after": 4})],
                 [(d, {"size": 10, "color": MUTED})]], margin=(0.15, 0.14, 0.1, 0.08))
    if i < 5:
        line(s, cx + 1.89, 3.27, cx + 2.04, 3.27, color=MUTED, arrow=True)
code_block(s, 0.6, 4.3, 6.55, 2.3, """
function aggregate(rows, key) {
  const map = new Map();
  rows.forEach(r => {
    const k = r[key]; if (!k) return;
    if (!map.has(k)) map.set(k, {name:k, sales:0, profit:0, count:0});
    const a = map.get(k);
    a.sales += r.sales; a.profit += r.profit; a.count++;
  });
  return [...map.values()];
}""", size=9.5, title="script.js  ·  grouping engine (formatted)")
code_block(s, 7.35, 4.3, 5.38, 2.3, """
// Quadrant rule used by the bubble chart
const mean = sales / products.length;
p.group = p.profit >= 0
  ? (p.sales >= mean ? 'star' : 'cow')
  : (p.sales <  mean ? 'question' : 'loss');

// Accounting negatives: '(1,200)' -> -1200
numeric('(1,200)')  // -1200""", size=9.5, title="script.js  ·  business rules (formatted)")
notes(s, """
Here's the pipeline. Papa Parse reads the file, the mapping step matches columns, numeric and dateKey clean the values, filteredRows applies the active filters, aggregate groups the data, and Chart.js renders it.
The aggregate function is simple on purpose: a single pass with a Map, so it's linear time. That's why 10,000 rows feel instant.
On the right is the actual quadrant rule and an example of number cleaning.
""")

# ---------------------------------------------------------------- modules
s = S["modules"]
heading(s, "03  ·  HOW IT WORKS", "Built in layers: a core plus six add-on modules",
        "Each module is a self-contained IIFE that adds charts without modifying the code before it.")
mods = [("Core", "Import, mapping, KPIs, quadrants, filters, comparison, world map, PDF", PURPLE, 1.0),
        ("Part 2", "Data enrichment: orders, customers, RFM, growth · distributions & counts", MINT, 0.92),
        ("Part 3", "Breakdowns, box plots, donuts, treemap, sunburst", SKY, 0.84),
        ("Part 4", "Correlation, heatmaps, discount analysis, density", AMBER, 0.84),
        ("Part 5", "Time intelligence: trends, growth, calendars, rankings", CYAN, 0.84),
        ("Part 6", "Customer intelligence: RFM 3D, cohorts, Pareto", LIME, 0.84),
        ("Part 7", "Shipping, waterfall, Sankey, outliers, state & city maps", ROSE, 0.84)]
for i, (n, d, col, _) in enumerate(mods):
    yy = 2.5 + i * 0.59
    c = box(s, 0.6, yy, 7.1, 0.5, fill=PANEL, line=BORDER, radius=0.08)
    box_text(c, [[(n.ljust(8), {"bold": True, "font": HEAD, "color": col}), ("  " + d, {"color": SOFT})]],
             size=11.5, anchor=MSO_ANCHOR.MIDDLE, margin=(0.18, 0, 0.1, 0))
    box(s, 0.6, yy + 0.08, 0.06, 0.34, fill=col, line=None, shape=MSO_SHAPE.RECTANGLE)
principles = [("Append-only", "New parts never replace earlier functions, so features can't break each other.", PURPLE),
              ("Fault isolated", "Each part renders inside try/catch. One failing chart shows a message, not a blank app.", MINT),
              ("Efficient", "Re-renders are batched with requestAnimationFrame, and old charts are destroyed to free memory.", SKY),
              ("Shared API", "window.PPAAdditions and PPAPart3 to PPAPart7 expose helpers to later modules.", AMBER)]
for i, (t, d, col) in enumerate(principles):
    cx = 8.0 + (i % 2) * 2.4
    cy = 2.5 + (i // 2) * 2.06
    c = box(s, cx, cy, 2.28, 1.94, fill=PANEL, line=BORDER, radius=0.1)
    box_text(c, [[(t, {"size": 13.5, "bold": True, "font": HEAD, "color": col, "after": 5})],
                 [(d, {"size": 10.5, "color": MUTED})]], margin=(0.17, 0.15, 0.12, 0.08))
notes(s, """
As the app grew, I structured script.js as a core plus six add-on modules. Each one is an immediately-invoked function, so its variables are private.
Four principles: modules only append, so earlier features can't break; each module is wrapped in try/catch, so one failure doesn't take down the page; renders are batched per animation frame; and modules talk through a small shared API on window.
That made it safe to keep adding features incrementally.
""")

# ---------------------------------------------------------------- tree
s = S["tree"]
heading(s, "04  ·  HOW EACH FILE WORKS", "The repository: six files, no build step",
        "Click a file to see how it works.")
tree = [("profit-portfolio-analyzer/", "", None, MUTED, ""),
        ("├─ index.html", "14 KB", "htmlcss", PURPLE, "Page structure"),
        ("├─ style.css", "31 KB", "htmlcss", MINT, "Visual design"),
        ("├─ script.js", "435 KB", "scriptjs", SKY, "All app logic and map data"),
        ("├─ amplify.yml", "0.5 KB", "amplifycsv", AMBER, "AWS build instructions"),
        ("├─ sample-data/superstore.csv", "2.3 MB", "amplifycsv", ROSE, "Demo dataset"),
        ("└─ README.md", "<1 KB", None, MUTED, "Project title")]
tp = box(s, 0.6, 2.45, 6.4, 4.15, fill=CODEBG, line=BORDER, radius=0.12)
for i, (name, size_, tgt, col, desc) in enumerate(tree):
    yy = 2.65 + i * 0.55
    if tgt:
        row = box(s, 0.8, yy, 6.0, 0.47, fill=PANEL, line=BORDER, radius=0.08)
    else:
        row = box(s, 0.8, yy, 6.0, 0.47, fill=None, line=None, radius=0.08)
    box_text(row, [[(name, {"font": MONO, "color": col if tgt else MUTED, "bold": bool(tgt)}),
                    ("   " + desc, {"color": MUTED, "size": 10})]], size=12, anchor=MSO_ANCHOR.MIDDLE,
             margin=(0.15, 0, 0.1, 0))
    if size_:
        t_ = text(s, 5.55, yy, 1.1, 0.47, size_ + ("  →" if tgt else ""), size=10, color=col if tgt else MUTED,
                  align=PP_ALIGN.RIGHT, anchor=MSO_ANCHOR.MIDDLE)
        if tgt:
            link([row, t_], tgt)
# how they connect
hc = box(s, 7.3, 2.45, 5.43, 4.15, fill=PANEL, line=BORDER, radius=0.12)
text(s, 7.55, 2.6, 5, 0.3, "How the files connect", size=13, bold=True, font=HEAD)
flows = [("index.html", "loads", "style.css + script.js", PURPLE),
         ("index.html", "loads", "4 libraries from the CDN", PURPLE),
         ("script.js", "fills", "the <canvas> elements in index.html", SKY),
         ("script.js", "fetches", "sample-data/superstore.csv", SKY),
         ("amplify.yml", "copies", "everything into site-output/", AMBER),
         ("AWS Amplify", "publishes", "site-output/ to a public URL", AMBER)]
for i, (a, verb, b_, col) in enumerate(flows):
    yy = 3.05 + i * 0.57
    c = box(s, 7.55, yy, 1.55, 0.42, fill=PANEL2, line=BORDER, radius=0.08)
    box_text(c, a, size=10, font=MONO, color=col, bold=True, anchor=MSO_ANCHOR.MIDDLE, align=PP_ALIGN.CENTER,
             margin=(0.03, 0, 0.03, 0))
    text(s, 9.15, yy, 0.85, 0.42, verb + " →", size=9.5, color=MUTED, anchor=MSO_ANCHOR.MIDDLE,
         align=PP_ALIGN.CENTER)
    text(s, 10.05, yy, 2.6, 0.42, b_, size=10.5, color=SOFT, anchor=MSO_ANCHOR.MIDDLE)
notes(s, """
The whole project is six files. There's no package.json, no bundler and no framework. You could open index.html straight from disk.
index.html is the skeleton, style.css the look, script.js the brain, amplify.yml tells AWS how to publish, and the CSV is the demo data.
I'll click into each one.
""")

# ---------------------------------------------------------------- html + css
s = S["htmlcss"]
heading(s, "04  ·  HOW EACH FILE WORKS", "index.html & style.css",
        "The skeleton and the skin. Together they're about 45 KB.")
for i, (fname, tag, col, pts, snippet) in enumerate([
    ("index.html", "THE SKELETON  ·  14 KB", PURPLE,
     ["Single page: a sidebar (data library, compare buttons, PDF export) and a main area.",
      "Header actions: Load Superstore, Load sample, Upload CSVs (hidden file input).",
      "Column-mapping panel, analysis controls and filters.",
      "Single-file view and comparison view, toggled with the hidden attribute.",
      "Every chart is a <canvas>. Libraries load with defer, so the page paints first."],
     '<canvas id="bubbleChart"></canvas>\n<script defer src="script.js"></script>'),
    ("style.css", "THE SKIN  ·  31 KB", MINT,
     ["A dark theme built on CSS variables: --bg, --panel, --purple, --mint, --rose.",
      "Reusable “panel” cards, KPI tiles, pills and buttons.",
      "CSS Grid layouts with 9 responsive breakpoints (580 px to 1500 px).",
      "Respects prefers-reduced-motion for accessibility.",
      "Typography: DM Sans for text and Manrope for headings."],
     ":root{--bg:#0c0e14; --panel:#14171f;\n  --purple:#ac99ff; --mint:#54d6b3}")]):
    cx = 0.6 + i * 6.18
    c = box(s, cx, 2.45, 5.95, 4.15, fill=PANEL, line=BORDER, radius=0.12)
    text(s, cx + 0.25, 2.6, 5.4, 0.3, tag, size=10, color=col, bold=True, font=HEAD)
    text(s, cx + 0.25, 2.88, 5.4, 0.45, fname, size=20, bold=True, font=MONO, color=TEXT)
    text(s, cx + 0.25, 3.42, 5.5, 2.2, pts, size=11.5, color=SOFT, bullet="•", bullet_color=col, space_after=4)
    cb = box(s, cx + 0.25, 5.72, 5.45, 0.72, fill=CODEBG, line=BORDER, radius=0.06)
    box_text(cb, snippet.split("\n"), size=10, font=MONO, color=col, anchor=MSO_ANCHOR.MIDDLE,
             margin=(0.15, 0.04, 0.1, 0.04))
button(s, 3.3, 6.98, 1.75, 0.36, "←  File tree", "tree", size=10.5)
notes(s, """
index.html defines every section of the interface: the sidebar data library, the header actions, the mapping panel, the filters, and both the single-file and comparison views. Charts are empty canvas elements that JavaScript fills.
style.css gives it the dark, modern look using CSS variables for the colour palette, CSS Grid for layout, nine responsive breakpoints, and reduced-motion support for accessibility.
""")

# ---------------------------------------------------------------- script.js
s = S["scriptjs"]
heading(s, "04  ·  HOW EACH FILE WORKS", "script.js: the brain of the app",
        "435 KB in total: about 190 KB of hand-written logic (187 functions) and 245 KB of embedded map data.", color=SKY)
fns = [("numeric(value)", "Cleans currency, commas and (negatives)"),
       ("dateKey(value)", "Normalises dates to YYYY-MM"),
       ("openMapping()", "Auto-matches CSV headers to fields"),
       ("aggregate(rows, key)", "Groups sales, profit and count"),
       ("render()", "KPIs, quadrants, bars and insights"),
       ("renderComparison()", "Multi-file comparison view"),
       ("renderPair()", "A/B side-by-side with deltas"),
       ("renderWorldMap()", "Geo bubble map with zoom and pan"),
       ("downloadPdf.onclick", "Builds the A4 PDF report with jsPDF")]
tb_ = s.shapes.add_table(len(fns) + 1, 2, I(0.6), I(2.45), I(7.0), I(4.1)).table
tb_.columns[0].width = I(2.65)
tb_.columns[1].width = I(4.35)
tb_.first_row = True
tb_.horz_banding = False
for r in range(len(fns) + 1):
    tb_.rows[r].height = I(0.41)
    for cidx in range(2):
        cell = tb_.cell(r, cidx)
        cell.fill.solid()
        cell.fill.fore_color.rgb = PANEL2 if r == 0 else (PANEL if r % 2 else rgb("171A23"))
        cell.margin_left = I(0.15)
        cell.margin_right = I(0.08)
        cell.margin_top = I(0.03)
        cell.margin_bottom = I(0.03)
        cell.vertical_anchor = MSO_ANCHOR.MIDDLE
        tf = cell.text_frame
        tf.word_wrap = True
        if r == 0:
            write(tf, ["KEY FUNCTION", "WHAT IT DOES"][cidx], size=10, color=MUTED, bold=True, font=HEAD)
        else:
            fn, d = fns[r - 1]
            if cidx == 0:
                write(tf, fn, size=10.5, color=SKY, font=MONO, bold=True)
            else:
                write(tf, d, size=11, color=SOFT)
# composition chart
panel = box(s, 7.9, 2.45, 4.83, 4.15, fill=PANEL, line=BORDER, radius=0.12)
text(s, 8.12, 2.58, 4.4, 0.3, "What's inside the 435 KB", size=13, bold=True, font=HEAD)
cd = CategoryChartData()
cd.categories = ["App logic", "World atlas", "US states atlas + 580 city centres"]
cd.add_series("KB", [190, 108, 137])
ch = add_chart(s, XL_CHART_TYPE.DOUGHNUT, 8.0, 2.9, 2.35, 2.35, cd)
chart_base(ch)
dn = ch._chartSpace.find(".//" + qn("c:doughnutChart"))
hs = dn.find(qn("c:holeSize"))
if hs is None:
    hs = etree.SubElement(dn, qn("c:holeSize"))
hs.set("val", "62")
for j, c_ in enumerate([SKY, PURPLE, MINT]):
    pt = ch.plots[0].series[0].points[j]
    pt.format.fill.solid()
    pt.format.fill.fore_color.rgb = c_
    pt.format.line.color.rgb = PANEL
text(s, 10.45, 3.15, 2.2, 1.9,
     [[("● ", {"color": SKY}), ("App logic  ", {"bold": True}), ("~190 KB", {"color": MUTED})],
      [("● ", {"color": PURPLE}), ("World atlas  ", {"bold": True}), ("~108 KB", {"color": MUTED})],
      [("● ", {"color": MINT}), ("US atlas + cities  ", {"bold": True}), ("~137 KB", {"color": MUTED})]],
     size=10.5, space_after=8)
text(s, 8.12, 5.35, 4.45, 1.2,
     ["Layout: core (lines 1–133), then Parts 2 to 7 as isolated modules.",
      "Map data is embedded, so maps work with no extra network requests."],
     size=10.5, color=SOFT, bullet="•", bullet_color=SKY, space_after=4)
button(s, 3.3, 6.98, 1.75, 0.36, "←  File tree", "tree", size=10.5)
notes(s, """
script.js is where everything happens. More than half its size is embedded map data, the world and US atlases plus 580 US city coordinates, so maps never depend on an external map service.
The logic itself is around 190 KB across 187 functions. The table shows the key ones: cleaning, mapping, aggregating, rendering the single and comparison views, the map, and the PDF export.
""")

# ---------------------------------------------------------------- amplify + csv
s = S["amplifycsv"]
heading(s, "04  ·  HOW EACH FILE WORKS", "amplify.yml & the sample dataset",
        "The build recipe AWS follows, and the demo data behind “Load Superstore data”.", color=AMBER)
code_block(s, 0.6, 2.45, 6.3, 3.0, """
version: 1
frontend:
  phases:
    build:
      commands:
        - mkdir -p site-output/sample-data
        - cp index.html style.css script.js site-output/
        - if ls sample-data/*.csv ...; then cp ...; fi
  artifacts:
    baseDirectory: site-output
    files:
      - '**/*'""", lang="yaml", size=11, title="amplify.yml")
ann = [("1", "Create an output folder for the published site.", AMBER),
       ("2", "Copy the three site files into it.", AMBER),
       ("3", "Copy the sample CSVs, only if they exist (my fix).", MINT),
       ("4", "Publish everything in site-output/ to the CDN.", AMBER)]
for i, (n, d, col) in enumerate(ann):
    yy = 5.6 + (i // 2) * 0.5
    cx = 0.6 + (i % 2) * 3.18
    c = box(s, cx, yy, 3.08, 0.43, fill=PANEL, line=BORDER, radius=0.08)
    box_text(c, [[(n + "  ", {"bold": True, "color": col, "font": HEAD}), (d, {"color": SOFT})]], size=9.5,
             anchor=MSO_ANCHOR.MIDDLE, margin=(0.12, 0, 0.06, 0))
p2 = box(s, 7.2, 2.45, 5.53, 4.15, fill=PANEL, line=BORDER, radius=0.12)
text(s, 7.45, 2.6, 5, 0.3, "DEMO DATA  ·  2.3 MB", size=10, color=ROSE, bold=True, font=HEAD)
text(s, 7.45, 2.88, 5, 0.45, "superstore.csv", size=20, bold=True, font=MONO)
facts = [("9,994", "order lines"), ("21", "columns"), ("2014–17", "four years"), ("793", "customers"),
         ("5,009", "orders"), ("49", "US states")]
for i, (big, small) in enumerate(facts):
    cx = 7.45 + (i % 3) * 1.7
    cy = 3.45 + (i // 3) * 0.95
    c = box(s, cx, cy, 1.58, 0.82, fill=PANEL2, line=BORDER, radius=0.08)
    box_text(c, [[(big, {"size": 16, "bold": True, "font": HEAD})], [(small, {"size": 9.5, "color": MUTED})]],
             margin=(0.12, 0.08, 0.05, 0.04))
code_block(s, 7.45, 5.4, 5.05, 1.0, """
// "Load Superstore data" button
fetch('sample-data/superstore.csv')""", size=10.5)
button(s, 3.3, 6.98, 1.75, 0.36, "←  File tree", "tree", size=10.5)
notes(s, """
amplify.yml is the build specification AWS Amplify runs on every deploy. It creates an output folder, copies the site files, conditionally copies the sample CSVs, and publishes the folder.
Step 3 is my fix. Originally the copy was unconditional, and the build failed because the folder didn't exist yet. More on that in the deployment section.
The Superstore dataset is a well-known retail dataset: about 10,000 order lines over four years, fetched when you click Load Superstore data.
""")

# ---------------------------------------------------------------- deploy flow
s = S["deployflow"]
heading(s, "05  ·  HOW I DEPLOYED IT", "From git push to a live website",
        "Hosted on AWS Amplify with continuous deployment from GitHub.", color=ROSE)
flow = [("GitHub", "Source code lives in the main branch of the GitHub repository", PURPLE, "1"),
        ("Webhook", "Amplify detects every push to main automatically", SKY, "2"),
        ("Build", "Amplify runs amplify.yml and assembles site-output/", AMBER, "3"),
        ("Deploy", "Files are published to a global CDN with HTTPS", MINT, "4"),
        ("Live", "Anyone can open the app at the public URL", ROSE, "5")]
for i, (t, d, col, n) in enumerate(flow):
    cx = 0.6 + i * 2.48
    circ = box(s, cx + 0.73, 2.5, 0.66, 0.66, fill=PANEL2, line=col, shape=MSO_SHAPE.OVAL, line_w=1.75)
    box_text(circ, n, size=16, bold=True, font=HEAD, color=col, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE,
             margin=(0, 0, 0, 0))
    if i < 4:
        line(s, cx + 1.45, 2.83, cx + 3.15, 2.83, color=rgb("3A3F52"), width=1.5, arrow=True)
    c = box(s, cx, 3.35, 2.12, 1.75, fill=PANEL, line=BORDER, radius=0.1)
    box_text(c, [[(t, {"size": 16, "bold": True, "font": HEAD, "color": col, "after": 4})],
                 [(d, {"size": 11, "color": MUTED})]], margin=(0.18, 0.15, 0.12, 0.1), align=PP_ALIGN.LEFT)
ci = box(s, 0.6, 5.35, 7.6, 1.25, fill=rgb("161227"), line=rgb("3A3060"), radius=0.12)
box_text(ci, [[("CI/CD:  ", {"bold": True, "color": PURPLE, "font": HEAD}),
               ("every merge to main redeploys the site automatically in about a minute. "
                "No servers to patch, nothing to scale, and it fits in the AWS free tier.", {"color": SOFT})]],
         size=13, anchor=MSO_ANCHOR.MIDDLE, margin=(0.3, 0.1, 0.3, 0.1))
lv = box(s, 8.45, 5.35, 4.28, 1.25, fill=PANEL, line=BORDER, radius=0.12)
box_text(lv, [[("LIVE URL", {"size": 9.5, "bold": True, "color": ROSE, "font": HEAD, "after": 3})],
              [(LIVE_LABEL, {"size": 12, "font": MONO, "color": TEXT, "url": LIVE_URL, "after": 3})],
              [("Click to open the live app  ↗", {"size": 10, "color": MUTED})]],
         margin=(0.22, 0.14, 0.15, 0.1))
link(lv, LIVE_URL)
notes(s, """
Here's how it's deployed. The code is on GitHub. I connected the repository to AWS Amplify Hosting, which installs a webhook. Every push to main triggers a build that runs amplify.yml, and the result is published to a global CDN with HTTPS.
So deployment is fully automated: merge a pull request and the live site updates within about a minute. Because it's static, there are no servers to manage and the cost is essentially zero.
""")

# ---------------------------------------------------------------- deploy steps
s = S["deploysteps"]
heading(s, "05  ·  HOW I DEPLOYED IT", "Step by step, including a real bug I fixed",
        "Setting up AWS Amplify took six steps, and the first build would have failed.", color=ROSE)
steps6 = ["Push the project to GitHub (main branch).",
          "AWS Console → AWS Amplify → Create new app.",
          "Choose GitHub, authorise, then pick the repo and the main branch.",
          "Amplify picks up amplify.yml automatically, so the build settings stay as they are.",
          "Save and deploy. About a minute later the app is live on amplifyapp.com.",
          "Optional: add a custom domain under Hosting → Custom domains."]
for i, st_ in enumerate(steps6):
    yy = 2.45 + i * 0.7
    circ = box(s, 0.6, yy + 0.06, 0.46, 0.46, fill=PANEL2, line=ROSE, shape=MSO_SHAPE.OVAL, line_w=1.25)
    box_text(circ, str(i + 1), size=12, bold=True, font=HEAD, color=ROSE, align=PP_ALIGN.CENTER,
             anchor=MSO_ANCHOR.MIDDLE, margin=(0, 0, 0, 0))
    text(s, 1.22, yy, 4.9, 0.6, st_, size=12.5, color=SOFT, anchor=MSO_ANCHOR.MIDDLE)
bug = box(s, 6.45, 2.45, 6.28, 4.15, fill=PANEL, line=BORDER, radius=0.12)
text(s, 6.7, 2.6, 5.8, 0.3, "DEBUGGING STORY", size=10, color=AMBER, bold=True, font=HEAD)
text(s, 6.7, 2.88, 5.8, 0.4, "The build that would have failed", size=16, bold=True, font=HEAD)
text(s, 6.7, 3.32, 5.8, 0.55, [[("Problem: ", {"bold": True, "color": ROSE}),
                                ("amplify.yml copied sample-data/*.csv, but that folder wasn't in the repo yet.", {})]],
     size=11, color=SOFT)
code_block(s, 6.7, 3.92, 5.8, 0.48, "cp: cannot stat 'sample-data/*.csv': No such file", size=10)
text(s, 6.7, 4.5, 5.8, 0.3, [[("Fix: ", {"bold": True, "color": MINT}),
                              ("copy the CSVs only if they exist.", {})]], size=11, color=SOFT)
code_block(s, 6.7, 4.85, 5.8, 0.48, "if ls sample-data/*.csv; then cp sample-data/*.csv ...; fi", size=10)
text(s, 6.7, 5.45, 5.8, 1.1,
     [[("Result: ", {"bold": True, "color": PURPLE}),
       ("I tested the build locally both ways (with and without the folder). Then I merged PR #1 (the fix) and "
        "PR #2 (added superstore.csv) to main, and Amplify redeploys automatically.", {})]],
     size=11, color=SOFT)
notes(s, """
The setup itself is six steps in the AWS console.
The interesting part is the debugging. The original build spec copied the sample CSVs unconditionally, but that folder wasn't in the repository, so the cp command would have failed and stopped the build.
I reproduced it locally, made the copy conditional, tested both cases, and merged it through a pull request. Then I added the dataset in a second PR. It's a small example of how I approach problems: reproduce, fix, verify, then ship.
""")

# ---------------------------------------------------------------- compare table
s = S["compare"]
heading(s, "06  ·  WHAT MAKES IT DIFFERENT", "How it compares", "An honest comparison for the job it was built for: quick, private profitability analysis.",
        color=CYAN)
hdr = ["", "Profit Portfolio Analyzer", "Spreadsheets\n(Excel / Sheets)", "BI tools\n(Power BI / Tableau)", "Online dashboard\ntools"]
rows_ = [("Setup", "Open a URL", "Installed app or account", "Install, licence, account", "Sign-up"),
         ("Where data goes", "Stays in your browser", "Local file or cloud", "Often a cloud workspace", "Uploaded to the vendor"),
         ("Time to first insight", "Seconds (auto-mapping)", "Manual pivots & charts", "Build a data model first", "Templates, then setup"),
         ("Quadrants, RFM, cohorts", "Built in", "Build it yourself", "Build it yourself", "Varies"),
         ("Compare files side by side", "Built in, never merged", "Manual", "Possible with modelling", "Varies"),
         ("Cost", "Free", "Licence or free tier", "Free tiers, paid for sharing", "Usually a subscription"),
         ("Learning curve", "None", "Medium", "High", "Low to medium")]
gf = s.shapes.add_table(len(rows_) + 1, 5, I(0.6), I(2.45), I(12.13), I(3.6))
tb_ = gf.table
widths = [2.35, 2.6, 2.39, 2.39, 2.4]
for i, w_ in enumerate(widths):
    tb_.columns[i].width = I(w_)
tb_.first_row = True
tb_.horz_banding = False
for r in range(len(rows_) + 1):
    tb_.rows[r].height = I(0.6 if r == 0 else 0.42)
    for cidx in range(5):
        cell = tb_.cell(r, cidx)
        cell.fill.solid()
        hl = cidx == 1
        if r == 0:
            cell.fill.fore_color.rgb = DARKPURPLE if hl else PANEL2
        else:
            cell.fill.fore_color.rgb = rgb("1A1630") if hl else (PANEL if r % 2 else rgb("171A23"))
        cell.margin_left = I(0.12)
        cell.margin_right = I(0.06)
        cell.margin_top = I(0.02)
        cell.margin_bottom = I(0.02)
        cell.vertical_anchor = MSO_ANCHOR.MIDDLE
        tf = cell.text_frame
        tf.word_wrap = True
        if r == 0:
            write(tf, hdr[cidx].split("\n"), size=10.5, color=PURPLE if hl else SOFT, bold=True, font=HEAD)
        else:
            v = rows_[r - 1][cidx]
            if cidx == 0:
                write(tf, v, size=10.5, color=TEXT, bold=True, font=HEAD)
            elif hl:
                write(tf, [[("✓ ", {"color": MINT}), (v, {})]], size=10.5, color=TEXT, bold=True)
            else:
                write(tf, v, size=10.5, color=MUTED)
text(s, 0.6, 6.25, 12.1, 0.4,
     [[("Fair note: ", {"bold": True, "color": SOFT}),
       ("BI tools are far more powerful for enterprise data modelling and live databases. This app is not a replacement for them. "
        "It's the fastest, most private way to answer “what makes us money?” from a CSV.", {})]],
     size=10.5, color=MUTED)
notes(s, """
Here's an honest comparison. Spreadsheets and BI tools are excellent, but for this specific job (I have a CSV, I want to know what's profitable, and I don't want to upload it) this app wins on setup, privacy, time to insight and cost.
I'm careful to say it isn't a Power BI replacement. It's a focused tool that does one thing very well.
""")

# ---------------------------------------------------------------- differentiators
s = S["different"]
heading(s, "06  ·  WHAT MAKES IT DIFFERENT", "Four things that set it apart", color=CYAN)
diffs = [("01", "Privacy by design", "Files are read with the browser's File API and analysed in memory. There's no upload, no database and "
                                      "no tracking of your data. Close the tab and it's gone.", MINT),
         ("02", "Zero setup, near-zero cost", "No install, no account, no backend. It's a static site on AWS Amplify, "
                                               "so hosting costs almost nothing and nothing needs scaling.", PURPLE),
         ("03", "Honest analytics", "Files are never silently merged. Missing months show as gaps, unrecognised regions are listed "
                                     "instead of guessed, and every chart says what it measures.", AMBER),
         ("04", "BI-level depth in 3 files", "60+ views including RFM, cohorts, Sankey, waterfall, correlation and maps, "
                                              "all in vanilla JavaScript with no framework.", SKY)]
for i, (n, t, d, col) in enumerate(diffs):
    cx = 0.6 + (i % 2) * 6.18
    cy = 2.0 + (i // 2) * 2.38
    c = box(s, cx, cy, 5.95, 2.2, fill=PANEL, line=BORDER, radius=0.12)
    box_text(c, [[(n, {"size": 30, "bold": True, "font": HEAD, "color": col})],
                 [(t, {"size": 18, "bold": True, "font": HEAD, "after": 5})],
                 [(d, {"size": 12, "color": MUTED})]], margin=(0.3, 0.18, 0.3, 0.12), spacing=1.05)
notes(s, """
To sum up the differentiators: privacy by design, because data never leaves the browser. Zero setup and near-zero cost, because it's a static site. Honest analytics, because the app never invents data. And depth: over 60 analytical views in three plain files.
""")

# ---------------------------------------------------------------- skills
s = S["skills"]
heading(s, "06  ·  WHAT MAKES IT DIFFERENT", "Skills this project demonstrates",
        "What I built, and the technologies behind it.", color=CYAN)
skills = [("Front-end engineering", ["HTML5 semantic layout", "CSS3: variables, Grid, responsive design", "Vanilla JavaScript (ES2020+)", "Accessibility (aria, reduced motion)"], PURPLE),
          ("Data visualisation", ["Chart.js 4 and chartjs-chart-geo", "Custom canvas plugins: quadrants, heatmaps", "Treemap layout and a 3D RFM projection", "Interactive zoom and pan maps"], MINT),
          ("Data analysis", ["Parsing and cleaning messy CSVs", "RFM, cohorts, Pareto, z-score outliers", "Pearson correlation, growth (MoM, YoY)", "Profitability quadrant modelling"], AMBER),
          ("Cloud & DevOps", ["AWS Amplify Hosting with a global CDN", "CI/CD from GitHub on every push", "YAML build specs and debugging", "Git pull-request workflow"], ROSE)]
for i, (t, pts, col) in enumerate(skills):
    cx = 0.6 + i * 3.08
    c = box(s, cx, 2.45, 2.9, 4.15, fill=PANEL, line=BORDER, radius=0.12)
    box(s, cx + 0.25, 2.45, 0.8, 0.06, fill=col, line=None, shape=MSO_SHAPE.RECTANGLE)
    text(s, cx + 0.25, 2.7, 2.5, 0.75, t, size=16, bold=True, font=HEAD, color=TEXT)
    text(s, cx + 0.25, 3.5, 2.5, 3.0, pts, size=11.5, color=SOFT, bullet="✓", bullet_color=col, space_after=9)
notes(s, """
For a recruiter, here's what the project shows: front-end engineering without relying on a framework, serious data visualisation including custom canvas plugins, real analytical methods like RFM and cohort analysis, and cloud deployment with CI/CD on AWS.
""")

# ---------------------------------------------------------------- thanks
s = S["thanks"]
glow = box(s, 8.6, 1.0, 5.5, 5.5, fill=rgb("141027"), line=None, shape=MSO_SHAPE.OVAL)
heading(s, "WHAT'S NEXT", "Roadmap & thank you", color=PURPLE)
road = [("Custom domain and usage analytics", "Privacy-friendly page analytics on AWS"),
        ("Multi-currency support", "Today every value is formatted in USD"),
        ("Save sessions locally", "IndexedDB, so files survive a page refresh"),
        ("Excel (.xlsx) import", "Alongside CSV"),
        ("AI-written narrative summaries", "Generated on the device, to keep the privacy promise")]
for i, (t, d) in enumerate(road):
    yy = 2.15 + i * 0.68
    c = box(s, 0.6, yy, 6.6, 0.58, fill=PANEL, line=BORDER, radius=0.08)
    box_text(c, [[("→  ", {"color": PURPLE, "bold": True}), (t, {"bold": True, "font": HEAD}), ("   " + d, {"color": MUTED, "size": 10.5})]],
             size=12.5, anchor=MSO_ANCHOR.MIDDLE, margin=(0.2, 0, 0.1, 0))
text(s, 7.8, 2.2, 5, 0.9, [[("Thank you", {}), (".", {"color": PURPLE})]], size=40, bold=True, font=HEAD)
text(s, 7.8, 3.1, 4.9, 0.7, "Try it yourself: load the Superstore data and explore.", size=14, color=SOFT)
button(s, 7.8, 4.0, 2.4, 0.5, "Open live app  ↗", LIVE_URL, primary=True, size=12.5)
button(s, 10.35, 4.0, 2.3, 0.5, "View on GitHub  ↗", GITHUB_URL, size=12.5)
button(s, 7.8, 4.7, 2.4, 0.45, "↺  Back to menu", "agenda", size=11.5)
button(s, 10.35, 4.7, 2.3, 0.45, "⌂  Title slide", "title", size=11.5)
text(s, 7.8, 5.45, 4.9, 0.8, [[("Questions?  ", {"bold": True, "color": TEXT}),
                              ("Use the menu to jump back to any section.", {})]], size=12, color=MUTED)
notes(s, """
To close, here's where I'd take it next: multi-currency, saving sessions locally, Excel import, and on-device AI summaries that keep the privacy promise.
Thank you. I'm happy to open the live app and demo it, or jump back to any section you'd like to discuss.
""")

prs.save(OUT)
print("Saved", OUT, "with", len(prs.slides), "slides")
