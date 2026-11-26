#!/usr/bin/env python3
"""Render PDF pages for visual inspection and crop crisp 300-DPI paper figures.

Usage:
  # 1. List figure captions in a PDF
  python3 crop_pdf_figure.py index paper.pdf

  # 2. Render candidate pages to locate figure bounds
  python3 crop_pdf_figure.py preview paper.pdf --pages 1,3,5 --out-dir ./page_renders

  # 3. Crop a specific region using normalized fractions (x0,y0,x1,y1 in [0,1])
  python3 crop_pdf_figure.py crop paper.pdf --page 3 --rect 0.51,0.08,0.87,0.255 --out ./fig.png
"""
import argparse
import os
import re
import fitz
from PIL import Image, ImageChops


def trim_white(img: Image.Image, pad: int = 16) -> Image.Image:
    bg = Image.new(img.mode, img.size, (255, 255, 255))
    diff = ImageChops.difference(img, bg)
    bbox = diff.getbbox()
    if not bbox:
        return img
    x0 = max(0, bbox[0] - pad)
    y0 = max(0, bbox[1] - pad)
    x1 = min(img.width, bbox[2] + pad)
    y1 = min(img.height, bbox[3] + pad)
    return img.crop((x0, y0, x1, y1))


def cmd_index(args):
    doc = fitz.open(args.pdf)
    print(f"{os.path.basename(args.pdf)} ({len(doc)} pages)")
    for pi in range(len(doc)):
        txt = doc[pi].get_text("text")
        for m in re.finditer(r"(Fig(?:ure|\.)\s*\d+[\.:][^\n]+(?:\n[^\n]+)?)", txt):
            cap = " ".join(m.group(1).split())[:160]
            print(f"  [p.{pi+1}] {cap}")


def cmd_preview(args):
    os.makedirs(args.out_dir, exist_ok=True)
    doc = fitz.open(args.pdf)
    stem = os.path.splitext(os.path.basename(args.pdf))[0]
    pages = [int(p.strip()) for p in args.pages.split(",") if p.strip()]
    for p in pages:
        if 1 <= p <= len(doc):
            out_path = os.path.join(args.out_dir, f"{stem}_p{p}.png")
            doc[p - 1].get_pixmap(dpi=args.dpi).save(out_path)
            print(f"Saved {out_path}")


def cmd_crop(args):
    os.makedirs(os.path.dirname(os.path.abspath(args.out)), exist_ok=True)
    fx0, fy0, fx1, fy1 = [float(v.strip()) for v in args.rect.split(",")]
    doc = fitz.open(args.pdf)
    page = doc[args.page - 1]
    r = page.rect
    clip = fitz.Rect(r.width * fx0, r.height * fy0, r.width * fx1, r.height * fy1)
    pix = page.get_pixmap(dpi=args.dpi, clip=clip)
    img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
    if not args.no_trim:
        img = trim_white(img, pad=args.pad)
    img.save(args.out)
    print(f"Saved {args.out}: {img.size[0]}x{img.size[1]} (aspect={img.size[0]/img.size[1]:.2f})")


def main():
    ap = argparse.ArgumentParser(description="Inspect and crop 300-DPI figures from academic PDFs.")
    sub = ap.add_subparsers(dest="cmd", required=True)

    p_idx = sub.add_parser("index", help="List Figure captions and page numbers")
    p_idx.add_argument("pdf")
    p_idx.set_defaults(func=cmd_index)

    p_prev = sub.add_parser("preview", help="Render selected pages as PNGs")
    p_prev.add_argument("pdf")
    p_prev.add_argument("--pages", required=True, help="Comma-separated 1-based page numbers")
    p_prev.add_argument("--out-dir", required=True)
    p_prev.add_argument("--dpi", type=int, default=200)
    p_prev.set_defaults(func=cmd_preview)

    p_crop = sub.add_parser("crop", help="Crop a figure region at high DPI")
    p_crop.add_argument("pdf")
    p_crop.add_argument("--page", type=int, required=True)
    p_crop.add_argument("--rect", required=True, help="Normalized x0,y0,x1,y1 in [0,1]")
    p_crop.add_argument("--out", required=True)
    p_crop.add_argument("--dpi", type=int, default=300)
    p_crop.add_argument("--pad", type=int, default=16)
    p_crop.add_argument("--no-trim", action="store_true")
    p_crop.set_defaults(func=cmd_crop)

    args = ap.parse_args()
    args.func(args)


if __name__ == "__main__":
    main()
