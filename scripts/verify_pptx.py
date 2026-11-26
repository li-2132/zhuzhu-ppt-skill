#!/usr/bin/env python3
"""Render every slide of a .pptx to PNG via Playwright Chromium and verify:
  1. OOXML ZIP integrity (0 duplicate entries, 0 broken slide rels)
  2. Zero vertical text-box overflows (>8px)
  3. Zero accidental AI-cliché / trailing-ellipsis text flags
"""
import base64
import html
import os
import posixpath
import re
import sys
import zipfile
import xml.etree.ElementTree as ET
from collections import Counter
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.enum.text import PP_ALIGN
from playwright.sync_api import sync_playwright

EMU_PER_PX = 914400 / 96.0
AI_CLICHE_PATTERNS = [
    r"！", r"杀手", r"死穴", r"悬崖", r"暗藏", r"怎么打", r"定锚",
    r"抢占先机", r"赋能", r"抓手", r"可复用积木", r"精准卡位", r"敬请批评指正",
]


def rgb_hex(color_obj):
    try:
        if color_obj and color_obj.type == 1 and color_obj.rgb:
            return f"#{color_obj.rgb}"
    except Exception:
        pass
    return None


def fill_hex(fill):
    try:
        if fill and fill.type == 1:
            return rgb_hex(fill.fore_color)
    except Exception:
        pass
    return None


def align_css(al):
    if al == PP_ALIGN.CENTER:
        return "center"
    if al == PP_ALIGN.RIGHT:
        return "right"
    return "left"


def check_ooxml_integrity(pptx_path: str) -> list[str]:
    errs = []
    with zipfile.ZipFile(pptx_path) as z:
        names = z.namelist()
        dups = {k: v for k, v in Counter(names).items() if v > 1}
        if dups:
            errs.append(f"Duplicate ZIP entries: {dups}")
        name_set = set(names)
        ct = ET.fromstring(z.read("[Content_Types].xml"))
        for el in ct:
            if el.tag.endswith("Override"):
                pn = el.attrib["PartName"].lstrip("/")
                if pn not in name_set:
                    errs.append(f"Missing Content_Types part: {pn}")
        for s in sorted(n for n in name_set if n.startswith("ppt/slides/slide") and n.endswith(".xml")):
            rf = s.replace("ppt/slides/", "ppt/slides/_rels/") + ".rels"
            if rf in name_set:
                sr = ET.fromstring(z.read(rf))
                for el in sr:
                    if el.attrib.get("TargetMode") == "External":
                        continue
                    resolved = posixpath.normpath(posixpath.join("ppt/slides", el.attrib["Target"]))
                    if resolved not in name_set:
                        errs.append(f"Broken rel in {rf}: {resolved}")
    return errs


def render_text_frame(tf):
    paras = []
    for p in tf.paragraphs:
        al = align_css(p.alignment)
        lh = p.line_spacing if isinstance(p.line_spacing, (int, float)) and p.line_spacing <= 3.0 else 1.22
        mb = (p.space_after.pt * 96 / 72) if p.space_after else 0
        runs = []
        for r in p.runs:
            if not r.text:
                continue
            sz = (r.font.size.pt * 96 / 72) if r.font.size else 14
            col = rgb_hex(r.font.color) or "#0F172A"
            wt = "700" if r.font.bold else "400"
            runs.append(f'<span style="font-size:{sz:.1f}px;color:{col};font-weight:{wt};">'
                        f'{html.escape(r.text).replace(chr(10), "<br>")}</span>')
        paras.append(f'<div style="text-align:{al};line-height:{lh:.2f};margin-bottom:{mb:.1f}px;word-break:break-word;">'
                     + "".join(runs) + "</div>")
    return "".join(paras)


def verify_and_render(pptx_path: str, out_dir: str):
    ooxml_errs = check_ooxml_integrity(pptx_path)
    prs = Presentation(pptx_path)
    w_px = int(round(prs.slide_width / EMU_PER_PX))
    h_px = int(round(prs.slide_height / EMU_PER_PX))
    os.makedirs(out_dir, exist_ok=True)

    style_notes = []
    pages_html = []

    for s_idx, slide in enumerate(prs.slides, 1):
        elems = []
        for sh in slide.shapes:
            x = sh.left / EMU_PER_PX
            y = sh.top / EMU_PER_PX
            w = sh.width / EMU_PER_PX
            h = sh.height / EMU_PER_PX

            if sh.shape_type == MSO_SHAPE_TYPE.PICTURE:
                blob = sh.image.blob
                ext = sh.image.ext or "png"
                mime = "image/jpeg" if ext.lower() in ("jpg", "jpeg") else f"image/{ext}"
                b64 = base64.b64encode(blob).decode("ascii")
                elems.append(
                    f'<img src="data:{mime};base64,{b64}" style="position:absolute;'
                    f'left:{x:.1f}px;top:{y:.1f}px;width:{w:.1f}px;height:{h:.1f}px;object-fit:fill;"/>'
                )
            elif sh.shape_type == MSO_SHAPE_TYPE.TABLE:
                tbl = sh.table
                col_ws = [c.width / EMU_PER_PX for c in tbl.columns]
                rows_html = []
                for r in tbl.rows:
                    rh = r.height / EMU_PER_PX
                    tds = []
                    for ci, cell in enumerate(r.cells):
                        txt = cell.text_frame.text.strip()
                        for pat in AI_CLICHE_PATTERNS:
                            if re.search(pat, txt):
                                style_notes.append(f"Slide {s_idx} table: matches /{pat}/ -> {txt[:30]!r}")
                        bg = fill_hex(cell.fill) or "transparent"
                        ml = cell.margin_left / EMU_PER_PX
                        mt = cell.margin_top / EMU_PER_PX
                        mr = cell.margin_right / EMU_PER_PX
                        mb = cell.margin_bottom / EMU_PER_PX
                        inner = render_text_frame(cell.text_frame)
                        tds.append(
                            f'<td style="width:{col_ws[ci]:.1f}px;background:{bg};'
                            f'padding:{mt:.1f}px {mr:.1f}px {mb:.1f}px {ml:.1f}px;'
                            f'border-bottom:1px solid #E2E8F0;vertical-align:middle;">{inner}</td>'
                        )
                    rows_html.append(f'<tr style="height:{rh:.1f}px;">{"".join(tds)}</tr>')
                elems.append(
                    f'<table style="position:absolute;left:{x:.1f}px;top:{y:.1f}px;'
                    f'width:{w:.1f}px;border-collapse:collapse;table-layout:fixed;">'
                    + "".join(rows_html) + "</table>"
                )
            elif sh.shape_type == MSO_SHAPE_TYPE.LINE:
                lc = rgb_hex(sh.line.color) or "#E2E8F0"
                lw = max(1.0, (sh.line.width.pt * 96 / 72) if sh.line.width else 1.0)
                if w >= h:
                    elems.append(f'<div style="position:absolute;left:{x:.1f}px;top:{y:.1f}px;width:{w:.1f}px;height:{lw:.1f}px;background:{lc};"></div>')
                else:
                    elems.append(f'<div style="position:absolute;left:{x:.1f}px;top:{y:.1f}px;width:{lw:.1f}px;height:{h:.1f}px;background:{lc};"></div>')
            else:
                bg = fill_hex(sh.fill) if hasattr(sh, "fill") else None
                lc, lw, dashed = None, 1.0, "solid"
                try:
                    if sh.line.fill.type == 1:
                        lc = rgb_hex(sh.line.color)
                        lw = max(1.0, (sh.line.width.pt * 96 / 72) if sh.line.width else 1.0)
                        if 'val="dash"' in sh._element.xml:
                            dashed = "dashed"
                except Exception:
                    pass
                border_css = f"{lw:.1f}px {dashed} {lc}" if lc else "none"
                rad = "6px" if "roundRect" in sh._element.xml else "0"
                inner, cls, pad_css = "", "shape", "0"
                if sh.has_text_frame and sh.text_frame.text.strip():
                    cls = "tbox"
                    tf = sh.text_frame
                    txt = tf.text.strip()
                    for pat in AI_CLICHE_PATTERNS:
                        if re.search(pat, txt):
                            style_notes.append(f"Slide {s_idx} sid={sh.shape_id}: matches /{pat}/ -> {txt[:30]!r}")
                    if txt.endswith(("...", "…", "等等")):
                        style_notes.append(f"Slide {s_idx} sid={sh.shape_id}: trailing ellipsis -> {txt[-20:]!r}")
                    pad_css = (f"{tf.margin_top/EMU_PER_PX:.1f}px {tf.margin_right/EMU_PER_PX:.1f}px "
                               f"{tf.margin_bottom/EMU_PER_PX:.1f}px {tf.margin_left/EMU_PER_PX:.1f}px")
                    inner = render_text_frame(tf)
                elems.append(
                    f'<div class="{cls}" data-sid="{sh.shape_id}" style="position:absolute;'
                    f'left:{x:.1f}px;top:{y:.1f}px;width:{w:.1f}px;height:{h:.1f}px;'
                    f'background:{bg or "transparent"};border:{border_css};border-radius:{rad};'
                    f'padding:{pad_css};overflow:visible;">{inner}</div>'
                )

        phtml = f"""<!DOCTYPE html><html><head><meta charset="utf-8"><style>
* {{ box-sizing:border-box; margin:0; padding:0; }}
body {{
  width:{w_px}px; height:{h_px}px; background:#FFFFFF; position:relative; overflow:hidden;
  font-family:"WenQuanYi Zen Hei","Microsoft YaHei","DejaVu Sans",sans-serif;
  -webkit-font-smoothing:antialiased;
}}
</style></head><body>{"".join(elems)}</body></html>"""
        pages_html.append((s_idx, phtml))

    overflows = []
    with sync_playwright() as p:
        browser = p.chromium.launch()
        page = browser.new_page(viewport={"width": w_px, "height": h_px}, device_scale_factor=1.5)
        for s_idx, phtml in pages_html:
            page.set_content(phtml)
            checks = page.evaluate("""() => Array.from(document.querySelectorAll('.tbox')).map(el => ({
                sid: el.dataset.sid, ch: el.clientHeight, sh: el.scrollHeight, txt: el.innerText.slice(0, 35)
            }))""")
            for c in checks:
                if c["sh"] > c["ch"] + 8:
                    overflows.append(f"Slide {s_idx} sid={c['sid']}: scrollH {c['sh']} > boxH {c['ch']} (+{c['sh']-c['ch']}px) | {c['txt']!r}")
            page.screenshot(path=os.path.join(out_dir, f"slide_{s_idx:02d}.png"))
        browser.close()

    print(f"Rendered {len(pages_html)} slides to {out_dir}")
    print(f"OOXML integrity issues: {len(ooxml_errs)}")
    for e in ooxml_errs:
        print("  OOXML:", e)
    print(f"Text box overflow alerts (>8px): {len(overflows)}")
    for o in overflows:
        print("  OVERFLOW:", o)
    print(f"De-AI / ellipsis lint notes: {len(style_notes)}")
    for n in style_notes:
        print("  STYLE:", n)


if __name__ == "__main__":
    if len(sys.argv) != 3:
        print("Usage: python3 verify_pptx.py <input.pptx> <out_preview_dir>", file=sys.stderr)
        sys.exit(2)
    verify_and_render(sys.argv[1], sys.argv[2])
