#!/usr/bin/env python3
"""Reusable python-pptx helpers for minimalist pure-white academic seminar slides."""
from PIL import Image
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn
from pptx.oxml import parse_xml

# Restrained white-canvas academic palette
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
INK = RGBColor(0x0F, 0x17, 0x2A)
BODY = RGBColor(0x33, 0x41, 0x55)
MUTED = RGBColor(0x64, 0x74, 0x8B)
LIGHT_GRAY = RGBColor(0x94, 0xA3, 0xB8)
BORDER = RGBColor(0xE2, 0xE8, 0xF0)
SURFACE = RGBColor(0xF8, 0xFA, 0xFC)
BLUE = RGBColor(0x1D, 0x4E, 0xD8)
BLUE_DARK = RGBColor(0x1E, 0x3A, 0x8A)
BLUE_TINT = RGBColor(0xEF, 0xF6, 0xFF)
TEAL = RGBColor(0x0F, 0x76, 0x6E)
ORANGE = RGBColor(0xD9, 0x48, 0x0F)
ORANGE_DARK = RGBColor(0x9A, 0x34, 0x12)
ORANGE_TINT = RGBColor(0xFF, 0xF7, 0xED)
RED = RGBColor(0xDC, 0x26, 0x26)
RED_TINT = RGBColor(0xFE, 0xF2, 0xF2)
AMBER = RGBColor(0xD9, 0x77, 0x06)

DEFAULT_FONT = "Microsoft YaHei"


def new_deck():
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    return prs, prs.slide_layouts[6]


def set_slide_bg(slide, color=WHITE):
    fill = slide.background.fill
    fill.solid()
    fill.fore_color.rgb = color


def add_rect(slide, x, y, w, h, *, fill=None, line=None, line_pt=0.75, rounded=False, dashed=False):
    shape_type = MSO_SHAPE.ROUNDED_RECTANGLE if rounded else MSO_SHAPE.RECTANGLE
    sh = slide.shapes.add_shape(shape_type, Inches(x), Inches(y), Inches(w), Inches(h))
    if rounded:
        sh.adjustments[0] = 0.06
    if fill is not None:
        sh.fill.solid()
        sh.fill.fore_color.rgb = fill
    else:
        sh.fill.background()
    if line is not None:
        sh.line.color.rgb = line
        sh.line.width = Pt(line_pt)
        if dashed:
            ln = sh._element.spPr.find(qn("a:ln"))
            if ln is not None:
                ln.append(parse_xml('<a:prstDash xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" val="dash"/>'))
    else:
        sh.line.fill.background()
    sh.shadow.inherit = False
    return sh


def add_line(slide, x1, y1, x2, y2, color=BORDER, width_pt=0.75):
    sh = slide.shapes.add_connector(1, Inches(x1), Inches(y1), Inches(x2), Inches(y2))
    sh.line.color.rgb = color
    sh.line.width = Pt(width_pt)
    sh.shadow.inherit = False
    return sh


def add_textbox(slide, x, y, w, h, lines, *, align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP,
                margin=(0.03, 0.02, 0.03, 0.02), font_name=DEFAULT_FONT):
    tb = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = anchor
    tf.margin_left = Inches(margin[0])
    tf.margin_top = Inches(margin[1])
    tf.margin_right = Inches(margin[2])
    tf.margin_bottom = Inches(margin[3])

    for i, item in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        if isinstance(item, dict):
            p.alignment = item.get("align", align)
            if "space_after" in item:
                p.space_after = Pt(item["space_after"])
            if "space_before" in item:
                p.space_before = Pt(item["space_before"])
            if "line_spacing" in item:
                p.line_spacing = item["line_spacing"]
            for rtxt, rsz, rbold, rcol in item["runs"]:
                r = p.add_run()
                r.text = rtxt
                r.font.name = font_name
                r.font.size = Pt(rsz)
                r.font.bold = rbold
                r.font.color.rgb = rcol
        else:
            txt, sz, bold, col = item
            p.alignment = align
            p.line_spacing = 1.25
            r = p.add_run()
            r.text = txt
            r.font.name = font_name
            r.font.size = Pt(sz)
            r.font.bold = bold
            r.font.color.rgb = col
    return tb


def add_picture_fit(slide, img_path, x, y, max_w, max_h, *, center_x=True, center_y=True, border_color=None):
    with Image.open(img_path) as im:
        iw, ih = im.size
    aspect = iw / ih
    if max_w / max_h > aspect:
        h = max_h
        w = h * aspect
    else:
        w = max_w
        h = w / aspect
    px = x + (max_w - w) / 2 if center_x else x
    py = y + (max_h - h) / 2 if center_y else y
    if border_color is not None:
        add_rect(slide, px - 0.04, py - 0.04, w + 0.08, h + 0.08, fill=WHITE, line=border_color, line_pt=0.75)
    pic = slide.shapes.add_picture(img_path, Inches(px), Inches(py), width=Inches(w), height=Inches(h))
    return pic, (px, py, w, h)


def add_text_section(slide, x, y, w, h, lines, *, top_rule=None, top_rule_pt=1.0, rule_width=None):
    """Preferred pure-white borderless content block: typography + whitespace + optional thin top rule."""
    top_pad = 0.0
    if top_rule is not None:
        rw = rule_width if rule_width is not None else w
        add_line(slide, x, y, x + rw, y, color=top_rule, width_pt=top_rule_pt)
        top_pad = 0.12
    return add_textbox(slide, x, y + top_pad, w, max(0.2, h - top_pad), lines, margin=(0, 0, 0, 0))


def add_callout(slide, x, y, w, h, lines, *, fill=SURFACE, border=BORDER, accent=None, bar_side=None):
    """Sparingly-used quiet tinted box (max 1 per slide, <=3 per deck; no thick accent bars by default)."""
    add_rect(slide, x, y, w, h, fill=fill, line=border, line_pt=0.75)
    if bar_side == "left" and accent is not None:
        add_rect(slide, x, y, 0.05, h, fill=accent)
        add_textbox(slide, x + 0.20, y + 0.10, w - 0.32, h - 0.16, lines, margin=(0, 0, 0, 0))
    elif bar_side == "top" and accent is not None:
        add_rect(slide, x, y, w, 0.04, fill=accent)
        add_textbox(slide, x + 0.18, y + 0.14, w - 0.36, h - 0.22, lines, margin=(0, 0, 0, 0))
    else:
        add_textbox(slide, x + 0.18, y + 0.11, w - 0.36, h - 0.20, lines, margin=(0, 0, 0, 0))


def style_cell(cell, text, *, size=10.5, bold=False, color=BODY, fill=None,
               align=PP_ALIGN.LEFT, border_bottom="E2E8F0", border_top=None,
               pad=(0.06, 0.05, 0.06, 0.05), font_name=DEFAULT_FONT):
    cell.vertical_anchor = MSO_ANCHOR.MIDDLE
    cell.margin_left = Inches(pad[0])
    cell.margin_top = Inches(pad[1])
    cell.margin_right = Inches(pad[2])
    cell.margin_bottom = Inches(pad[3])
    if fill is not None:
        cell.fill.solid()
        cell.fill.fore_color.rgb = fill
    else:
        cell.fill.background()

    tf = cell.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.alignment = align
    p.line_spacing = 1.15
    if isinstance(text, list):
        for rtxt, rbold, rcol in text:
            r = p.add_run()
            r.text = rtxt
            r.font.name = font_name
            r.font.size = Pt(size)
            r.font.bold = rbold
            r.font.color.rgb = rcol
    else:
        r = p.add_run()
        r.text = text
        r.font.name = font_name
        r.font.size = Pt(size)
        r.font.bold = bold
        r.font.color.rgb = color

    tcPr = cell._tc.get_or_add_tcPr()
    for tag in ("a:lnL", "a:lnR", "a:lnT", "a:lnB"):
        old = tcPr.find(qn(tag))
        if old is not None:
            tcPr.remove(old)
    for tag in ("a:lnL", "a:lnR"):
        tcPr.append(parse_xml(f'<{tag} xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:noFill/></{tag}>'))
    if border_top:
        tcPr.append(parse_xml(
            f'<a:lnT xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" w="12700">'
            f'<a:solidFill><a:srgbClr val="{border_top}"/></a:solidFill></a:lnT>'
        ))
    else:
        tcPr.append(parse_xml('<a:lnT xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:noFill/></a:lnT>'))
    if border_bottom:
        tcPr.append(parse_xml(
            f'<a:lnB xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" w="9525">'
            f'<a:solidFill><a:srgbClr val="{border_bottom}"/></a:solidFill></a:lnB>'
        ))
    else:
        tcPr.append(parse_xml('<a:lnB xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:noFill/></a:lnB>'))


def add_header_footer(slide, page_num, total_pages, kicker, title, *, footer_left="", accent=BLUE):
    set_slide_bg(slide, WHITE)
    add_textbox(slide, 0.60, 0.26, 6.5, 0.26, [(kicker, 10.0, True, accent)], margin=(0, 0, 0, 0))
    add_textbox(slide, 0.60, 0.52, 12.13, 0.48, [(title, 19.5, True, INK)], margin=(0, 0, 0, 0))
    add_line(slide, 0.60, 1.04, 12.73, 1.04, color=BORDER, width_pt=0.8)

    add_line(slide, 0.60, 7.08, 12.73, 7.08, color=BORDER, width_pt=0.5)
    if footer_left:
        add_textbox(slide, 0.60, 7.13, 8.5, 0.26, [(footer_left, 8.5, False, LIGHT_GRAY)], margin=(0, 0, 0, 0))
    add_textbox(slide, 11.4, 7.13, 1.33, 0.26,
                [(f"{page_num} / {total_pages}", 9.0, True, MUTED)], align=PP_ALIGN.RIGHT, margin=(0, 0, 0, 0))


def add_notes(slide, text):
    slide.notes_slide.notes_text_frame.text = text.strip()
