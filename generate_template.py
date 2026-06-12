#!/usr/bin/env python3
"""
Futsal National Team — Session Plan Template Generator
======================================================
Generates `futsal_session_template.pptx`: a fully editable, premium,
one-page session-plan template system for national-team futsal staff.

Slides:
  1. Master template      (keep clean, duplicate weekly)
  2. Working copy         (identical, ready to fill)
  3. Annotated guide      (same layout + deletable "HOW TO" callouts)
  4. Component library    (reusable futsal tactical symbols + courts)

Everything is native PowerPoint vector shapes / text boxes / tables.
No images, no locked elements, no hidden masking shapes.

Usage:  python3 generate_template.py
"""

import copy

from pptx import Presentation
from pptx.util import Emu, Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE, MSO_CONNECTOR
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.lang import MSO_LANGUAGE_ID
from pptx.enum.dml import MSO_LINE_DASH_STYLE
from pptx.oxml import parse_xml
from pptx.oxml.ns import qn

# ---------------------------------------------------------------- constants

EMU_IN = 914400


def IN(v):
    """Inches -> Emu."""
    return Emu(int(round(v * EMU_IN)))


SLIDE_W = 13.333
SLIDE_H = 7.5

C_DARK = RGBColor(0x1A, 0x1A, 0x1A)    # headings, primary text
C_MID = RGBColor(0x6B, 0x6B, 0x6B)     # secondary text
C_LITE = RGBColor(0xF2, 0xF2, 0xF2)    # focus bar / panel fills
C_FAINT = RGBColor(0xFA, 0xFA, 0xFA)   # subtle row tint
C_LINE = RGBColor(0xD9, 0xD9, 0xD9)    # hairline dividers
C_ACCENT = RGBColor(0x00, 0x57, 0xB8)  # single accent (hierarchy only)
C_COURT = RGBColor(0xA9, 0xB1, 0xBB)   # court line grey
C_WHITE = RGBColor(0xFF, 0xFF, 0xFF)

FONT = "Calibri"

# Court-scale marker size.
# On a 3.26" wide court (=40m), 1m = 0.0815".
# 0.14" ≈ 1.7 m — standard for elite coaching diagrams.
MARKER_D = 0.14

# page geometry (all inches; arithmetic validated to sum exactly)
MARGIN = 0.15
HEADER_H = 0.55
RULE_H = 0.02
FOCUS_Y = HEADER_H + RULE_H            # 0.57
FOCUS_H = 0.65
BODY_Y = 1.28
BODY_H = 6.07                          # bottom edge 7.35
BLOCK_W = 5.265
BLOCK_H = 3.00
BLOCK_GAP = 0.07
SQUAD_X = 10.83
SQUAD_W = 2.35

_shape_id = [9000]                     # id counter for hand-built group XML


def _next_id():
    _shape_id[0] += 1
    return _shape_id[0]


# ---------------------------------------------------------------- low-level helpers

def _no_shadow(shp):
    shp.shadow.inherit = False
    return shp


def add_box(shapes, x, y, w, h, fill=None, line=None, line_w=0.75,
            shape=MSO_SHAPE.RECTANGLE, dash=None):
    """Add a rectangle/autoshape with explicit fill + outline (None = off)."""
    shp = shapes.add_shape(shape, IN(x), IN(y), IN(w), IN(h))
    if fill is None:
        shp.fill.background()
    else:
        shp.fill.solid()
        shp.fill.fore_color.rgb = fill
    if line is None:
        shp.line.fill.background()
    else:
        shp.line.color.rgb = line
        shp.line.width = Pt(line_w)
        if dash is not None:
            shp.line.dash_style = dash
    return _no_shadow(shp)


def add_line(shapes, x1, y1, x2, y2, color=C_LINE, w=0.75, dash=None):
    ln = shapes.add_connector(MSO_CONNECTOR.STRAIGHT,
                              IN(x1), IN(y1), IN(x2), IN(y2))
    ln.line.color.rgb = color
    ln.line.width = Pt(w)
    if dash is not None:
        ln.line.dash_style = dash
    return _no_shadow(ln)


def add_arrowhead(line_shape):
    """python-pptx has no arrowhead API; append <a:tailEnd> to the line XML."""
    ln = line_shape.line._get_or_add_ln()
    tail = ln.makeelement(qn("a:tailEnd"),
                          {"type": "triangle", "w": "med", "len": "med"})
    ln.append(tail)
    return line_shape


def set_fill_alpha(shp, opacity_pct):
    """Set transparency on a solid fill via <a:alpha> (no python-pptx API)."""
    solid = shp.fill._xPr.find(qn("a:solidFill"))
    clr = solid.find(qn("a:srgbClr"))
    clr.append(clr.makeelement(qn("a:alpha"),
                               {"val": str(int(opacity_pct * 1000))}))


def add_text(shapes, x, y, w, h, paras, align=PP_ALIGN.LEFT,
             anchor=MSO_ANCHOR.TOP, wrap=True):
    """
    Text box. `paras` = list of paragraphs; each paragraph = list of run
    tuples (text, size_pt, bold, color). One-click editable, no autofit.
    """
    tb = shapes.add_textbox(IN(x), IN(y), IN(w), IN(h))
    tf = tb.text_frame
    tf.word_wrap = wrap
    tf.vertical_anchor = anchor
    tf.margin_left = IN(0.03)
    tf.margin_right = IN(0.03)
    tf.margin_top = IN(0.015)
    tf.margin_bottom = IN(0.015)
    for i, runs in enumerate(paras):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        for text, size, bold, color in runs:
            r = p.add_run()
            r.text = text
            f = r.font
            f.name = FONT
            f.size = Pt(size)
            f.bold = bold
            f.color.rgb = color
            f.language_id = MSO_LANGUAGE_ID.ENGLISH_UK
    return _no_shadow(tb)


def group_elements(slide, shps, x, y, w, h, name):
    """Wrap already-placed shapes into a <p:grpSp> with an identity
    child-coordinate transform, so the group moves/copies as one unit."""
    ex, ey, ew, eh = (IN(x), IN(y), IN(w), IN(h))
    grp = parse_xml(
        '<p:grpSp xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" '
        'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
        f'<p:nvGrpSpPr><p:cNvPr id="{_next_id()}" name="{name}"/>'
        '<p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>'
        '<p:grpSpPr><a:xfrm>'
        f'<a:off x="{ex}" y="{ey}"/><a:ext cx="{ew}" cy="{eh}"/>'
        f'<a:chOff x="{ex}" y="{ey}"/><a:chExt cx="{ew}" cy="{eh}"/>'
        '</a:xfrm></p:grpSpPr></p:grpSp>'
    )
    slide.shapes._spTree.append(grp)
    for s in shps:
        grp.append(s._element)
    return grp


def set_adj_angles(shp, start_deg, end_deg):
    """Set adj1/adj2 angle adjustments (clockwise from 3 o'clock) via avLst
    XML — python-pptx exposes no adjustment spec for arc/pie/chord presets."""
    prst = shp._element.spPr.find(qn("a:prstGeom"))
    av = prst.find(qn("a:avLst"))
    if av is None:
        av = prst.makeelement(qn("a:avLst"), {})
        prst.append(av)
    else:
        for gd in list(av):
            av.remove(gd)
    for nm, deg in (("adj1", start_deg), ("adj2", end_deg)):
        gd = av.makeelement(qn("a:gd"), {"name": nm,
                                         "fmla": f"val {int(deg * 60000)}"})
        av.append(gd)
    return shp


def add_arc(shapes, cx, cy, r, start_deg, end_deg, color=C_COURT, w=0.75):
    """MSO ARC autoshape with explicit start/end angles."""
    shp = shapes.add_shape(MSO_SHAPE.ARC,
                           IN(cx - r), IN(cy - r), IN(2 * r), IN(2 * r))
    set_adj_angles(shp, start_deg, end_deg)
    shp.fill.background()
    shp.line.color.rgb = color
    shp.line.width = Pt(w)
    return _no_shadow(shp)


# ---------------------------------------------------------------- futsal court

def draw_futsal_court(slide, x, y, w, group_name="Futsal court"):
    """
    Futsal-correct court at true 40m x 20m proportions (h = w/2).
    Penalty areas: two 6m quarter-arcs centred on each goalpost joined by a
    straight 3m segment (the real futsal D — not a football box, not a
    plain semicircle). Returns the grouped court.
    """
    h = w / 2.0
    s = w / 40.0                       # inches per metre
    cy = y + h / 2.0
    cx = x + w / 2.0
    goal_d = 0.07                      # goal stub depth behind goal line
    dot = 0.035                        # mark diameter
    parts = []

    # touchlines / goal lines
    parts.append(add_box(slide.shapes, x, y, w, h, fill=None,
                         line=C_COURT, line_w=1.0))
    # halfway line + centre circle + centre mark
    parts.append(add_line(slide.shapes, cx, y, cx, y + h, C_COURT, 0.75))
    parts.append(add_box(slide.shapes, cx - 3 * s, cy - 3 * s, 6 * s, 6 * s,
                         fill=None, line=C_COURT, line_w=0.75,
                         shape=MSO_SHAPE.OVAL))
    parts.append(add_box(slide.shapes, cx - dot / 2, cy - dot / 2, dot, dot,
                         fill=C_COURT, line=None, shape=MSO_SHAPE.OVAL))

    for end in ("L", "R"):
        gl = x if end == "L" else x + w          # goal line x
        sgn = 1 if end == "L" else -1            # into-court direction
        post_top = y + 8.5 * s                   # goal is 3m, centred
        post_bot = y + 11.5 * s
        # penalty area: quarter arc off each post + straight 6m segment
        if end == "L":
            parts.append(add_arc(slide.shapes, gl, post_top, 6 * s, 270, 360))
            parts.append(add_arc(slide.shapes, gl, post_bot, 6 * s, 0, 90))
        else:
            parts.append(add_arc(slide.shapes, gl, post_top, 6 * s, 180, 270))
            parts.append(add_arc(slide.shapes, gl, post_bot, 6 * s, 90, 180))
        seg_x = gl + sgn * 6 * s
        parts.append(add_line(slide.shapes, seg_x, post_top, seg_x, post_bot,
                              C_COURT, 0.75))
        # 6m penalty mark + 10m second penalty mark
        for dist in (6, 10):
            mx = gl + sgn * dist * s
            parts.append(add_box(slide.shapes, mx - dot / 2, cy - dot / 2,
                                 dot, dot, fill=C_COURT, line=None,
                                 shape=MSO_SHAPE.OVAL))
        # goal (3m wide stub behind the goal line)
        gx = gl - goal_d if end == "L" else gl
        parts.append(add_box(slide.shapes, gx, post_top, goal_d, 3 * s,
                             fill=None, line=C_COURT, line_w=1.0))

    return group_elements(slide, parts, x - goal_d, y,
                          w + 2 * goal_d, h, group_name)


def draw_futsal_half_court(slide, x, y, w, group_name="Futsal half court"):
    """Half court (20m x 20m square): goal end left, halfway line right."""
    s = w / 20.0
    h = w
    cy = y + h / 2.0
    dot = 0.035
    goal_d = 0.07
    parts = []

    parts.append(add_box(slide.shapes, x, y, w, h, fill=None,
                         line=C_COURT, line_w=1.0))
    # half centre circle on the halfway (right) edge
    parts.append(add_arc(slide.shapes, x + w, cy, 3 * s, 90, 270))
    parts.append(add_box(slide.shapes, x + w - dot / 2, cy - dot / 2,
                         dot, dot, fill=C_COURT, line=None,
                         shape=MSO_SHAPE.OVAL))
    post_top = y + h / 2 - 1.5 * s
    post_bot = y + h / 2 + 1.5 * s
    parts.append(add_arc(slide.shapes, x, post_top, 6 * s, 270, 360))
    parts.append(add_arc(slide.shapes, x, post_bot, 6 * s, 0, 90))
    parts.append(add_line(slide.shapes, x + 6 * s, post_top,
                          x + 6 * s, post_bot, C_COURT, 0.75))
    for dist in (6, 10):
        parts.append(add_box(slide.shapes, x + dist * s - dot / 2,
                             cy - dot / 2, dot, dot,
                             fill=C_COURT, line=None, shape=MSO_SHAPE.OVAL))
    parts.append(add_box(slide.shapes, x - goal_d, post_top, goal_d, 3 * s,
                         fill=None, line=C_COURT, line_w=1.0))

    return group_elements(slide, parts, x - goal_d, y, w + goal_d, h,
                          group_name)


# ---------------------------------------------------------------- page zones

def draw_header(slide):
    """White masthead: team + session title left, logistics fields right."""
    sh = slide.shapes
    add_text(sh, MARGIN, 0.02, 4.30, 0.51, [
        [("[INSERT TEAM]  ·  FUTSAL SESSION PLAN", 8, True, C_MID)],
        [("[Insert session title]", 20, True, C_DARK)],
    ])
    fields = [
        ("DATE", "[Insert date]", 4.62, 1.45),
        ("VENUE", "[Insert venue]", 6.17, 1.95),
        ("DURATION", "[Insert duration]", 8.22, 1.45),
        ("STAFF", "[Insert staff]", 9.77, 3.41),
    ]
    for label, value, fx, fw in fields:
        add_text(sh, fx, 0.07, fw, 0.44, [
            [(label, 7, True, C_MID)],
            [(value, 9.5, label == "DURATION", C_DARK)],
        ])
    # accent rule separates masthead from the working document
    add_box(sh, 0, HEADER_H, SLIDE_W, RULE_H, fill=C_ACCENT, line=None)


def draw_focus_bar(slide):
    """Tier-1 focus strip: theme, outcomes, constraints, reminders."""
    sh = slide.shapes
    add_box(sh, 0, FOCUS_Y, SLIDE_W, FOCUS_H, fill=C_LITE, line=None)
    cols = [
        ("SESSION THEME / OBJECTIVE", "[Insert focus]", MARGIN, 3.30, True),
        ("KEY COACHING OUTCOMES", "[Insert key outcomes]", 3.60, 3.45, False),
        ("KEY CONSTRAINTS", "[Insert constraints]", 7.20, 3.00, False),
        ("KEY REMINDERS", "[Insert reminders]", 10.35, 2.83, False),
    ]
    for label, value, fx, fw, is_theme in cols:
        add_text(sh, fx, FOCUS_Y + 0.05, fw, FOCUS_H - 0.10, [
            [(label, 7, True, C_ACCENT if is_theme else C_MID)],
            [(value, 11 if is_theme else 8.5, is_theme,
              C_DARK if is_theme else C_DARK)],
        ])
        if fx > MARGIN:
            add_line(sh, fx - 0.08, FOCUS_Y + 0.10,
                     fx - 0.08, FOCUS_Y + FOCUS_H - 0.10, C_LINE, 0.75)


def draw_activity_block(slide, x, y, number, title):
    """One session block: title bar / objective / court + coaching points /
    notes. Identical structure for all four blocks."""
    sh = slide.shapes
    w, h = BLOCK_W, BLOCK_H
    # container + title bar
    add_box(sh, x, y, w, h, fill=C_WHITE, line=C_LINE, line_w=0.75)
    bar_h = 0.28
    add_box(sh, x, y, w, bar_h, fill=C_LITE, line=None)
    add_box(sh, x, y, 0.05, bar_h, fill=C_ACCENT, line=None)
    add_text(sh, x + 0.07, y + 0.015, 0.36, 0.25,
             [[(number, 11, True, C_ACCENT)]])
    add_text(sh, x + 0.44, y + 0.03, 1.82, 0.23,
             [[(title, 9, True, C_DARK)]], wrap=False)
    add_text(sh, x + 2.30, y + 0.035, w - 1.40 - 2.30, 0.22, [[
        ("LEAD  ", 7, True, C_MID),
        ("[Insert responsibility]", 8, False, C_DARK),
    ]], align=PP_ALIGN.RIGHT, wrap=False)
    chip = add_box(sh, x + w - 1.30, y + 0.035, 1.22, 0.21,
                   fill=C_WHITE, line=C_ACCENT, line_w=1.0,
                   shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    tf = chip.text_frame
    tf.word_wrap = False
    tf.margin_left = IN(0.02)
    tf.margin_right = IN(0.02)
    tf.margin_top = 0
    tf.margin_bottom = 0
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    r = p.add_run()
    r.text = "[Insert timing]"
    r.font.name = FONT
    r.font.size = Pt(8)
    r.font.bold = True
    r.font.color.rgb = C_ACCENT

    # objective row
    obj_y = y + bar_h + 0.02
    add_text(sh, x + 0.07, obj_y, w - 0.14, 0.20, [[
        ("OBJECTIVE  ", 7, True, C_MID),
        ("[Insert objective]", 8.5, False, C_DARK),
    ]])

    # main area: court left / coaching points right
    main_y = obj_y + 0.21
    main_h = 1.92
    court_w = 3.26
    court_h = court_w / 2
    draw_futsal_court(slide, x + 0.12, main_y + (main_h - court_h) / 2,
                      court_w, f"Court {number}")
    div_x = x + 0.12 + court_w + 0.12
    add_line(sh, div_x, main_y + 0.06, div_x, main_y + main_h - 0.06,
             C_LINE, 0.75)
    cp_x = div_x + 0.06
    cp_w = x + w - 0.07 - cp_x
    add_text(sh, cp_x, main_y + 0.03, cp_w, main_h - 0.06, [
        [("COACHING POINTS", 7, True, C_MID)],
        [("[Insert coaching points]", 8, False, C_DARK)],
    ])

    # notes strip
    notes_y = main_y + main_h + 0.02
    notes_h = y + h - notes_y - 0.04
    add_line(sh, x + 0.07, notes_y, x + w - 0.07, notes_y, C_LINE, 0.75)
    add_text(sh, x + 0.07, notes_y + 0.02, w - 0.14, notes_h, [[
        ("NOTES / CONSTRAINTS / PROGRESSIONS   ", 7, True, C_MID),
        ("[Insert notes]", 8, False, C_DARK),
    ]])


def _style_cell(cell, text, size, bold, color, fill, align=PP_ALIGN.LEFT):
    cell.fill.solid()
    cell.fill.fore_color.rgb = fill
    cell.margin_left = IN(0.04)
    cell.margin_right = IN(0.02)
    cell.margin_top = IN(0.005)
    cell.margin_bottom = IN(0.005)
    cell.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = cell.text_frame.paragraphs[0]
    p.alignment = align
    r = p.add_run()
    r.text = text
    f = r.font
    f.name = FONT
    f.size = Pt(size)
    f.bold = bold
    f.color.rgb = color


def draw_squad_panel(slide):
    """Operational availability panel: counts, position-seeded player table
    (tab-through editing), squad notes."""
    sh = slide.shapes
    x, y, w, h = SQUAD_X, BODY_Y, SQUAD_W, BODY_H
    add_box(sh, x, y, w, h, fill=C_WHITE, line=C_LINE, line_w=0.75)
    add_box(sh, x, y, w, 0.30, fill=C_DARK, line=None)
    add_text(sh, x + 0.05, y + 0.045, w - 0.10, 0.24,
             [[("AVAILABLE PLAYERS / POSITIONS", 8, True, C_WHITE)]],
             wrap=False)
    # counts row — squad balance at a glance
    add_text(sh, x + 0.05, y + 0.33, w - 0.10, 0.24, [[
        ("TOTAL ", 7, True, C_MID), ("[#]   ", 9, True, C_DARK),
        ("GK ", 7, True, C_MID), ("[#]   ", 9, True, C_DARK),
        ("OUTFIELD ", 7, True, C_MID), ("[#]", 9, True, C_DARK),
    ]])

    # player table: header + 14 rows seeded with futsal positions
    positions = (["GK"] * 2 + ["FIXO"] * 3 + ["ALA"] * 6 + ["PIVOT"] * 3)
    rows = 1 + len(positions)
    tbl_y = y + 0.60
    tbl_h = 4.05                       # 15 rows x 0.27"; slack below is
    #                                    deliberate room for late additions
    gframe = sh.add_table(rows, 3, IN(x + 0.05), IN(tbl_y),
                          IN(w - 0.10), IN(tbl_h))
    tbl = gframe.table
    tbl.first_row = False
    tbl.horz_banding = False
    # "No Style, No Grid" — fills only, no theme-blue styling
    tblPr = tbl._tbl.tblPr
    for el in tblPr.findall(qn("a:tableStyleId")):
        tblPr.remove(el)
    tsid = tblPr.makeelement(qn("a:tableStyleId"), {})
    tsid.text = "{2D5ABB26-0587-4C30-8999-92F81FD0307C}"
    tblPr.append(tsid)

    tbl.columns[0].width = IN(0.52)
    tbl.columns[1].width = IN(1.08)
    tbl.columns[2].width = IN(0.65)
    row_h = IN(tbl_h / rows)
    for r in tbl.rows:
        r.height = row_h

    for c, head in enumerate(("POS", "PLAYER", "STATUS")):
        _style_cell(tbl.cell(0, c), head, 7, True, C_DARK, C_LITE)
    for i, pos in enumerate(positions, start=1):
        tint = C_FAINT if pos in ("FIXO", "PIVOT") else C_WHITE
        _style_cell(tbl.cell(i, 0), pos, 7.5, True, C_DARK, tint)
        _style_cell(tbl.cell(i, 1), "[Player]", 7.5, False, C_MID, tint)
        _style_cell(tbl.cell(i, 2), "", 7.5, False, C_MID, tint)

    # squad notes (anchored to panel bottom, clear of any table reflow)
    notes_y = y + 5.25
    add_line(sh, x + 0.05, notes_y, x + w - 0.05, notes_y, C_LINE, 0.75)
    add_text(sh, x + 0.05, notes_y + 0.03, w - 0.10,
             y + h - notes_y - 0.08, [
                 [("SQUAD NOTES", 7, True, C_MID)],
                 [("[Insert status notes]", 8, False, C_DARK)],
             ])


def add_footer(slide, text):
    add_text(slide.shapes, MARGIN, 7.355, SLIDE_W - 2 * MARGIN, 0.135,
             [[(text, 6.5, False, C_MID)]])


# ---------------------------------------------------------------- slide builders

BLOCKS = [
    ("01", "WARM UP"),
    ("02", "TECHNICAL / POSSESSION"),
    ("03", "TACTICAL ACTIVITY"),
    ("04", "GAME"),
]


def build_session_slide(prs, footer):
    slide = prs.slides.add_slide(prs.slide_layouts[6])   # blank layout
    draw_header(slide)
    draw_focus_bar(slide)
    for i, (num, title) in enumerate(BLOCKS):
        bx = MARGIN + (i % 2) * (BLOCK_W + BLOCK_GAP)
        by = BODY_Y + (i // 2) * (BLOCK_H + BLOCK_GAP)
        draw_activity_block(slide, bx, by, num, title)
    draw_squad_panel(slide)
    add_footer(slide, footer)
    return slide


def clone_slide(prs, source, footer):
    """Exact structural copy: deep-copy every shape element into a fresh
    blank slide. Safe here because all content is native shapes (no rels)."""
    new = prs.slides.add_slide(prs.slide_layouts[6])
    for shp in list(source.shapes):
        new.shapes._spTree.append(copy.deepcopy(shp._element))
    # swap the footer (last shape added by build order is the footer textbox)
    for shp in new.shapes:
        if shp.has_text_frame and shp.text_frame.text.startswith("MASTER"):
            shp.text_frame.paragraphs[0].runs[0].text = footer
    return new


def add_annotations(slide):
    """Deletable dashed accent 'HOW TO' callouts placed in spacious zones."""
    notes = [
        # (x, y, w, h, text) — sited inside diagram/notes zones
        (0.55, 2.10, 2.55, 0.62,
         "HOW TO — Set timing + lead coach in each block header. "
         "Keep the objective to one line."),
        (5.90, 2.10, 2.55, 0.62,
         "HOW TO — Copy players, arrows and zones from Slide 4 "
         "(Component Library) onto the court."),
        (0.55, 5.20, 2.55, 0.62,
         "HOW TO — Use the notes strip for constraints, progressions "
         "and setup / equipment reminders."),
        (5.90, 5.20, 2.55, 0.62,
         "HOW TO — Duplicate Slide 2 each week. Delete these blue "
         "notes — they are single text boxes."),
        (10.93, 3.30, 2.15, 0.80,
         "HOW TO — Update counts, names and status here for fast late "
         "changes. Tab moves between table cells."),
    ]
    for x, y, w, h, text in notes:
        box = add_box(slide.shapes, x, y, w, h, fill=C_WHITE,
                      line=C_ACCENT, line_w=1.0, dash=MSO_LINE_DASH_STYLE.DASH)
        tf = box.text_frame
        tf.word_wrap = True
        tf.margin_left = IN(0.05)
        tf.margin_right = IN(0.05)
        tf.margin_top = IN(0.03)
        tf.margin_bottom = IN(0.03)
        p = tf.paragraphs[0]
        p.alignment = PP_ALIGN.LEFT
        r = p.add_run()
        r.text = text
        r.font.name = FONT
        r.font.size = Pt(7.5)
        r.font.bold = False
        r.font.color.rgb = C_ACCENT


# ---------------------------------------------------------------- component library

def _marker(sh, x, y, fill, line, text, text_color, d=MARKER_D, dash=None):
    """Plain (non-directional) player circle. d defaults to court-scale size."""
    m = add_box(sh, x, y, d, d, fill=fill, line=line, line_w=1.0,
                shape=MSO_SHAPE.OVAL, dash=dash)
    tf = m.text_frame
    tf.word_wrap = False
    tf.margin_left = 0
    tf.margin_right = 0
    tf.margin_top = 0
    tf.margin_bottom = 0
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    r = p.add_run()
    r.text = text
    r.font.name = FONT
    r.font.size = Pt(7)
    r.font.bold = True
    r.font.color.rgb = text_color
    return m


def _marker_text(shp, label, text_color, size=7, lift=0.0):
    """lift > 0 nudges the centred text upward (extra bottom margin) —
    used by the chord puck whose visual centre sits above its bbox centre."""
    tf = shp.text_frame
    tf.word_wrap = False
    tf.margin_left = tf.margin_right = tf.margin_top = 0
    tf.margin_bottom = IN(lift)
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    r = p.add_run()
    r.text = str(label)
    f = r.font
    f.name = FONT
    f.size = Pt(size)
    f.bold = True
    f.color.rgb = text_color
    return shp


def draw_d_puck(sh, x, y, fill, line_color, label, text_color, dash=None):
    """
    Body-shape marker: a circle with a flattened back — the flat edge IS the
    shoulder line, the curved edge is the chest. Default orientation faces UP.
    A single CHORD autoshape: no grouping, rotates cleanly in one click, and
    reads as open/closed body shape instantly at court scale.
    """
    d = MARKER_D
    shp = sh.add_shape(MSO_SHAPE.CHORD, IN(x), IN(y), IN(d), IN(d))
    # chord cut across the lower edge: keep the sweep from 150° -> 30°
    # (through north), leaving a flat shoulder line at the back/south
    set_adj_angles(shp, 150, 30)
    if fill is None:
        shp.fill.background()
    else:
        shp.fill.solid()
        shp.fill.fore_color.rgb = fill
    if line_color is None:
        shp.line.fill.background()
    else:
        shp.line.color.rgb = line_color
        shp.line.width = Pt(1.0)
        if dash is not None:
            shp.line.dash_style = dash
    _no_shadow(shp)
    return _marker_text(shp, label, text_color, lift=0.03)


def draw_vision_marker(slide, x, y, fill, line_color, label, text_color,
                       dash=None):
    """
    Player circle + translucent vision cone (70° wedge, default facing UP),
    grouped as one object. Rotate the group to set body orientation — the
    cone shows facing AND field of view. Placed on a defender it doubles as
    a cover-shadow tool (the passing lanes the body blocks).
    x, y = top-left of the player circle. Cone radius ~2.5x marker.
    """
    d = MARKER_D
    cone_r = d * 1.75                  # wedge reach ≈ 4 m at court scale
    ccx = x + d / 2.0                  # player centre
    ccy = y + d / 2.0

    pie = slide.shapes.add_shape(MSO_SHAPE.PIE,
                                 IN(ccx - cone_r), IN(ccy - cone_r),
                                 IN(2 * cone_r), IN(2 * cone_r))
    set_adj_angles(pie, 235, 305)      # 70° wedge centred on north (270°)
    pie.fill.solid()
    cone_color = fill if fill not in (C_WHITE, None) else (line_color or C_DARK)
    pie.fill.fore_color.rgb = cone_color
    set_fill_alpha(pie, 20)
    pie.line.fill.background()
    _no_shadow(pie)

    circ = add_box(slide.shapes, x, y, d, d, fill=fill, line=line_color,
                   line_w=1.0, shape=MSO_SHAPE.OVAL, dash=dash)
    _marker_text(circ, label, text_color)

    return group_elements(slide, [pie, circ],
                          ccx - cone_r, ccy - cone_r,
                          2 * cone_r, 2 * cone_r, f"vm_{label}")


def _lib_label(sh, x, y, w, text):
    add_text(sh, x, y, w, 0.18, [[(text, 6.5, False, C_MID)]],
             align=PP_ALIGN.CENTER, wrap=False)


def _section_head(sh, x, y, text):
    add_box(sh, x, y + 0.03, 0.04, 0.16, fill=C_ACCENT, line=None)
    add_text(sh, x + 0.08, y, 2.6, 0.22, [[(text, 9, True, C_DARK)]],
             wrap=False)


def build_component_library(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    sh = slide.shapes

    add_text(sh, MARGIN, 0.05, 10.0, 0.50, [
        [("COMPONENT LIBRARY", 16, True, C_DARK)],
        [("All markers are at true court scale (≈1.7 m on a 40 m court). "
          "Copy and paste directly onto any diagram — they will be "
          "proportionally correct.", 8.5, False, C_MID)],
    ])
    add_box(sh, 0, HEADER_H, SLIDE_W, RULE_H, fill=C_ACCENT, line=None)

    # ── Row 1: plain markers | equipment | lines | zones ────────────────────
    r1y = 0.72

    # PLAIN PLAYER MARKERS (no direction indicator — for quick annotation)
    px, py = 0.40, r1y
    _section_head(sh, px, py, "PLAYER MARKERS  (plain)")
    plain = [
        (C_ACCENT, None,  "A", C_WHITE, None,                   "Attacker"),
        (C_WHITE,  C_MID, "D", C_MID,  None,                   "Defender"),
        (C_DARK,   None,  "G", C_WHITE, None,                   "GK"),
        (C_WHITE,  C_MID, "N", C_MID,  MSO_LINE_DASH_STYLE.DASH, "Neutral"),
    ]
    for i, (fill, line, t, tc, dash, lbl) in enumerate(plain):
        mx = px + 0.06 + i * 0.50
        _marker(sh, mx, py + 0.26, fill, line, t, tc, dash=dash)
        _lib_label(sh, mx - 0.04, py + 0.44, 0.44, lbl)

    # EQUIPMENT (court-scale)
    ex, ey = 2.55, r1y
    _section_head(sh, ex, ey, "EQUIPMENT")
    # Cone ≈1.1 m tall
    add_box(sh, ex + 0.08, ey + 0.28, 0.09, 0.08, fill=C_MID, line=None,
            shape=MSO_SHAPE.ISOSCELES_TRIANGLE)
    _lib_label(sh, ex + 0.01, ey + 0.44, 0.40, "Cone")
    # Pole ≈0.5 m wide × 2.5 m tall
    add_box(sh, ex + 0.62, ey + 0.24, 0.04, 0.20, fill=C_MID, line=None)
    _lib_label(sh, ex + 0.47, ey + 0.44, 0.40, "Pole")
    # Ball ≈0.85 m
    add_box(sh, ex + 1.10, ey + 0.30, 0.07, 0.07, fill=C_DARK, line=None,
            shape=MSO_SHAPE.OVAL)
    _lib_label(sh, ex + 0.95, ey + 0.44, 0.40, "Ball")

    # LINES
    lx, ly = 3.75, r1y
    _section_head(sh, lx, ly, "LINES")
    add_arrowhead(add_line(sh, lx + 0.05, ly + 0.38,
                           lx + 0.80, ly + 0.38, C_DARK, 1.5))
    _lib_label(sh, lx, ly + 0.44, 0.90, "Run")
    add_arrowhead(add_line(sh, lx + 1.00, ly + 0.38,
                           lx + 1.75, ly + 0.38,
                           C_DARK, 1.5, dash=MSO_LINE_DASH_STYLE.DASH))
    _lib_label(sh, lx + 0.95, ly + 0.44, 0.90, "Pass")
    add_arrowhead(add_line(sh, lx + 1.95, ly + 0.38,
                           lx + 2.70, ly + 0.38,
                           C_DARK, 1.5, dash=MSO_LINE_DASH_STYLE.ROUND_DOT))
    _lib_label(sh, lx + 1.90, ly + 0.44, 0.90, "Dribble")

    # ZONES & SCREENS
    zx, zy = 7.05, r1y
    _section_head(sh, zx, zy, "ZONES / SCREENS")
    zone = add_box(sh, zx + 0.05, zy + 0.26, 0.70, 0.24, fill=C_ACCENT, line=None)
    set_fill_alpha(zone, 18)
    _lib_label(sh, zx + 0.02, zy + 0.44, 0.76, "Zone")
    add_box(sh, zx + 1.00, zy + 0.34, 0.36, 0.06, fill=C_DARK, line=None,
            shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    _lib_label(sh, zx + 0.88, zy + 0.44, 0.62, "Screen")

    add_line(sh, MARGIN, 1.65, SLIDE_W - MARGIN, 1.65, C_LINE, 0.5)

    # ── Row 2: body shape & orientation ─────────────────────────────────────
    r2y = 1.72
    _section_head(sh, 0.40, r2y, "BODY SHAPE  —  shoulder-line pucks")

    puck_items = [
        (C_ACCENT, None,  "A", C_WHITE, None,                   "Attacker"),
        (C_WHITE,  C_MID, "D", C_MID,  None,                   "Defender"),
        (C_DARK,   None,  "G", C_WHITE, None,                   "GK"),
        (C_WHITE,  C_MID, "N", C_MID,  MSO_LINE_DASH_STYLE.DASH, "Neutral"),
    ]
    for i, (fill, line, t, tc, dash, lbl) in enumerate(puck_items):
        mx = 0.47 + i * 0.50
        draw_d_puck(sh, mx, r2y + 0.28, fill, line, t, tc, dash=dash)
        _lib_label(sh, mx - 0.06, r2y + 0.46, 0.44, lbl)

    add_text(sh, 0.40, r2y + 0.66, 2.55, 0.36, [[
        ("Flat edge = shoulders, curve = chest. One shape — rotate it to "
         "show open / closed body shape.", 7.5, False, C_MID),
    ]])

    _section_head(sh, 3.30, r2y, "VISION CONE  —  facing + field of view")
    cone_items = [
        (C_ACCENT, None,  "A", C_WHITE, "Attacker"),
        (C_WHITE,  C_MID, "D", C_MID,  "Defender"),
        (C_DARK,   None,  "G", C_WHITE, "GK"),
    ]
    for i, (fill, line, t, tc, lbl) in enumerate(cone_items):
        mx = 3.55 + i * 0.62
        draw_vision_marker(slide, mx, r2y + 0.42, fill, line, t, tc)
        _lib_label(sh, mx - 0.15, r2y + 0.60, 0.44, lbl)

    add_text(sh, 3.30, r2y + 0.80, 2.85, 0.30, [[
        ("Rotate the group: cone shows where the player faces and what "
         "they see. On a defender it doubles as a cover shadow.",
         7.5, False, C_MID),
    ]])

    # rotation instruction callout
    instr = add_box(sh, 6.55, r2y + 0.24, 3.30, 0.62,
                    fill=C_FAINT, line=C_LINE, line_w=0.75)
    tf = instr.text_frame
    tf.word_wrap = True
    tf.margin_left = IN(0.06); tf.margin_right = IN(0.06)
    tf.margin_top = IN(0.04); tf.margin_bottom = IN(0.04)
    p = tf.paragraphs[0]
    r = p.add_run()
    r.text = ("ROTATION  ·  All orientation markers face UP by default. "
              "Select → drag the rotation handle, or Format Shape → "
              "Rotation and type the angle (0 = up, 90 = right, "
              "180 = down, 270 = left).")
    r.font.name = FONT; r.font.size = Pt(7.5); r.font.color.rgb = C_MID

    add_line(sh, MARGIN, 2.92, SLIDE_W - MARGIN, 2.92, C_LINE, 0.5)

    # ── Row 3: Team A + Team B numbered sets (body-shape pucks) ─────────────
    r3y = 2.98

    # Team A (blue — your team)
    add_text(sh, 0.40, r3y, 2.4, 0.20, [[
        ("TEAM A  ", 8, True, C_ACCENT),
        ("your team", 7.5, False, C_MID),
    ]])
    for n in range(1, 6):
        draw_d_puck(sh, 0.45 + (n - 1) * 0.38, r3y + 0.24,
                    C_ACCENT, None, str(n), C_WHITE)

    # Team B (dark — opposition)
    add_text(sh, 2.75, r3y, 3.0, 0.20, [[
        ("TEAM B  ", 8, True, C_DARK),
        ("opposition — change fill colour to match", 7.5, False, C_MID),
    ]])
    for n in range(1, 6):
        draw_d_puck(sh, 2.80 + (n - 1) * 0.38, r3y + 0.24,
                    C_DARK, None, str(n), C_WHITE)

    # Scale note
    add_text(sh, 5.50, r3y, 7.50, 0.55, [
        [("SCALE + COLOUR", 7, True, C_MID)],
        [("Markers are at true court scale — copy straight onto a diagram. "
          "Team pucks carry body shape too: rotate any puck to angle the "
          "shoulders. To recolour Team B: select all 5 → Format Shape → "
          "Fill.", 8, False, C_DARK)],
    ])

    add_line(sh, MARGIN, 3.58, SLIDE_W - MARGIN, 3.58, C_LINE, 0.5)

    # ── Row 4: courts ───────────────────────────────────────────────────────
    cy0 = 3.65
    _section_head(sh, 0.40, cy0, "FULL COURT  (40m × 20m)")
    draw_futsal_court(slide, 0.55, cy0 + 0.32, 4.40, "Library full court")
    _section_head(sh, 6.20, cy0, "HALF COURT  (20m × 20m)")
    draw_futsal_half_court(slide, 6.35, cy0 + 0.32, 2.20,
                           "Library half court")

    add_text(sh, 9.65, cy0 + 0.32, 3.30, 3.20, [
        [("USAGE", 7, True, C_MID)],
        [("Courts are grouped — one click selects the whole court. "
          "Copy and paste onto any diagram, then resize to fit. "
          "Ungroup (Ctrl+Shift+G) to edit individual markings.", 8, False, C_DARK)],
        [("", 4, False, C_DARK)],
        [("To change the accent colour on a slide: select all blue shapes "
          "(click one, then Edit → Select All Same Fill Colour) and choose "
          "Format Shape → Fill.", 8, False, C_DARK)],
        [("", 4, False, C_DARK)],
        [("All shapes are unlocked. Numbers inside markers are "
          "editable text — double-click to change.", 8, False, C_DARK)],
    ])

    add_footer(slide, "COMPONENT LIBRARY  —  court-scale tactical symbols  ·  "
                      "copy into any diagram zone  ·  pucks + vision cones: "
                      "rotate to show body shape and field of view")
    return slide


# ---------------------------------------------------------------- main

def main():
    prs = Presentation()
    prs.slide_width = IN(SLIDE_W)
    prs.slide_height = IN(SLIDE_H)

    master = build_session_slide(
        prs, "MASTER TEMPLATE  —  keep clean  ·  duplicate Slide 2 for "
             "each session")
    clone_slide(prs, master,
                "WORKING COPY  —  ready to edit  ·  duplicate weekly")
    guide = clone_slide(prs, master,
                        "GUIDE  —  annotated example  ·  delete the blue "
                        "notes after reading")
    add_annotations(guide)
    build_component_library(prs)

    out = "futsal_session_template.pptx"
    prs.save(out)
    print(f"Saved {out}: {len(prs.slides._sldIdLst)} slides")


if __name__ == "__main__":
    main()
