#!/usr/bin/env python3
"""Find a word or phrase in PDF, Office, HTML and image files: in diagrams and images, body text and hyperlinks.

Selectable text inside vector diagrams is read from the PDF text layer. Words that
exist only as pixels or outlines (pictures, screenshots, logos, outlined vector art)
are found with OCR (RapidOCR). Body text and hyperlink targets are reported separately.
Office files are converted to PDF first so page and slide numbers match the rendered
document; HTML is scanned image by image, then its text and links.

Exit status: 0 = found, 1 = not found, 2 = error.
"""
from __future__ import annotations

import argparse
import base64
import gzip
import hashlib
import io
import json
import os
import posixpath
import re
import shutil
import subprocess
import sys
import tempfile
import time
import traceback
import unicodedata
import urllib.parse
import urllib.request
import zipfile
from dataclasses import asdict, dataclass, field
from html.parser import HTMLParser
from xml.etree import ElementTree as ET


def die(message):
    print(f"error: {message}", file=sys.stderr)
    sys.exit(2)


try:
    import pymupdf
except ImportError:
    try:
        import fitz as pymupdf
    except ImportError:
        die("PyMuPDF is required: pip install pymupdf")

WORD_EXTS = {".doc", ".docx", ".docm", ".dot", ".dotx", ".rtf", ".odt"}
SLIDE_EXTS = {".ppt", ".pptx", ".pptm", ".pps", ".ppsx", ".pot", ".potx", ".odp"}
HTML_EXTS = {".html", ".htm", ".xhtml"}
SVG_EXTS = {".svg", ".svgz"}
IMAGE_EXTS = {".png", ".apng", ".jpg", ".jpeg", ".jpe", ".jfif", ".gif", ".bmp", ".dib", ".tif", ".tiff", ".webp",
              ".avif", ".jp2", ".j2k", ".jpx", ".ico", ".cur", ".pcx", ".ppm", ".pgm", ".pbm", ".pnm", ".tga",
              ".psd", ".qoi", ".dds", ".emf", ".wmf"}
TERM_COLORS = ["#e11d48", "#2563eb", "#16a34a", "#d97706", "#7c3aed", "#0891b2", "#db2777", "#4d7c0f"]
DISPLAY_SIDE = 1800   # px; longest side of screenshots in the HTML report and --crops
THUMB_WIDTH = 520     # px; page thumbnails in the HTML report
LINK_ZOOM = 1.6       # render scale of the page strip shown for a hyperlink or repeated text
TEXT_ZOOM = 1.4       # render scale of the whole-page screenshot shown for body-text matches

OCR_CONFUSABLES = {"1": "l", "|": "l", "!": "l", "i": "l", "0": "o"}
FIGURE_CAPTION = re.compile(r"^\s*(?:figure|fig\.?)\s*\d", re.I)
TABLE_CAPTION = re.compile(r"^\s*table\s*\d", re.I)
MAX_OCR_SIDE = 4000   # px; RapidOCR shrinks larger images, which loses small labels
LABEL_MARGIN = 6      # pt around a graphic whose text still counts as its labels
MAX_PATHS = 5000      # beyond this, a page is treated as one big drawing
LABELS = {
    ("text", "vector"): "diagram text",
    ("ocr", "vector"): "vector art (OCR)",
    ("text", "image"): "text on image",
    ("ocr", "image"): "image (OCR)",
    ("link", "link"): "hyperlink",
    ("body", "text"): "body text",
}
NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS_W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


# ---------------------------------------------------------------- matching

def fold(text, case_sensitive=False, ocr=False):
    """Normalise text for matching; index[k] is the position in text of folded[k]."""
    chars, index = [], []
    for i, ch in enumerate(text):
        for c in unicodedata.normalize("NFKD", ch):
            for cc in c if case_sensitive else c.lower():
                if unicodedata.combining(cc):
                    continue
                if ocr and not case_sensitive:
                    cc = OCR_CONFUSABLES.get(cc, cc)
                chars.append(cc)
                index.append(i)
    return "".join(chars), index


def _boundary(left, right):
    if not (left.isalnum() and right.isalnum()):
        return True
    # OCR and diagram labels often drop spaces: "ClientIntelFPGA", "Intel64"
    return (left.islower() and right.isupper()) or left.isdigit() != right.isdigit()


class Matcher:
    def __init__(self, terms, case_sensitive=False, substring=False):
        self.terms = terms
        self.case_sensitive = case_sensitive
        self.substring = substring
        self.multiword = any(len(t.split()) > 1 for t in terms)
        self._rx = {}
        for term in terms:
            for ocr in (False, True):
                words = list(map(re.escape, fold(term, case_sensitive, ocr)[0].split()))
                self._rx[term, ocr, False] = re.compile(r"\s+".join(words))
                self._rx[term, ocr, True] = re.compile(r"\s*".join(words))

    def find(self, text, ocr=False, joined=None):
        """Yield (term, start, end) spans of text that match a search term.

        joined lets the words of a phrase run together ("IntelFPGA"); it is on by default for OCR text.
        """
        folded, index = fold(text, self.case_sensitive, ocr)
        joined = ocr if joined is None else joined
        for term in self.terms:
            rx, pos = self._rx[term, ocr, joined], 0
            while m := rx.search(folded, pos):
                a, b = m.start(), m.end()
                start, end = index[a], index[b - 1] + 1
                pos = a + 1
                if self.substring or self._whole_word(text, folded, index, a, b, start, end):
                    yield term, start, end
                    pos = b

    @staticmethod
    def _whole_word(text, folded, index, a, b, start, end):
        if (a > 0 and index[a - 1] == start) or (b < len(folded) and index[b] == end - 1):
            return False  # match starts or ends inside an expanded character such as a ligature
        return (start == 0 or _boundary(text[start - 1], text[start])) and \
               (end == len(text) or _boundary(text[end - 1], text[end]))


@dataclass
class Line:
    text: str
    bbox: tuple
    score: float | None = None
    chars: list | None = None  # per-character boxes aligned with text (None entries for spaces)


@dataclass
class Found:
    term: str
    text: str
    bbox: tuple       # the matched characters
    line_bbox: tuple  # the whole line(s) containing them
    score: float | None = None


@dataclass
class Hit:
    term: str
    page: int | None        # 1-based page or slide; None when the format has no pages
    text: str               # the whole label, line or paragraph that contains the term, or the link target
    source: str             # "text" = selectable text in a graphic, "ocr" = read from pixels or outlines,
                            # "link" = link target, "body" = ordinary document text
    region: str             # "vector" (drawing), "image" (raster picture), "link" or "text"
    where: str = ""         # figure caption, header/footer, image name, link text or text location
    bbox: list | None = None
    score: float | None = None
    page_approx: bool = False
    crop: str | None = None
    repeated: bool = False


def _min_score(*scores):
    known = [s for s in scores if s is not None]
    return min(known) if known else None


def _below(a, lines):
    """The line directly under a, i.e. where a wrapped label continues."""
    h, best, best_gap = a.bbox[3] - a.bbox[1], None, None
    for b in lines:
        gap = b.bbox[1] - a.bbox[3]
        if b is not a and -0.3 * h <= gap <= 0.8 * h and min(a.bbox[2], b.bbox[2]) > max(a.bbox[0], b.bbox[0]) \
                and (best is None or gap < best_gap):
            best, best_gap = b, gap
    return best


def _span_bbox(ln, s, e):
    boxes = [b for b in (ln.chars or [])[s:e] if b]
    if boxes:
        bbox = boxes[0]
        for b in boxes[1:]:
            bbox = union(bbox, b)
        return bbox
    n, (x0, y0, x1, y1) = max(len(ln.text), 1), ln.bbox
    return (x0 + (x1 - x0) * s / n, y0, x0 + (x1 - x0) * e / n, y1)


def search_lines(matcher, lines, ocr=False, max_wrap=4):
    found = []
    for ln in lines:
        for term, s, e in matcher.find(ln.text, ocr):
            found.append(Found(term, ln.text, _span_bbox(ln, s, e), ln.bbox, ln.score))
    if matcher.multiword:  # phrases wrapped over stacked lines, e.g. "Avalon / Memory-Mapped / Interface"
        below = {id(ln): _below(ln, lines) for ln in lines}
        for first in lines:
            chain, joined, nxt = [first], first.text.strip(), below[id(first)]
            while nxt is not None and len(chain) < max_wrap and all(nxt is not c for c in chain):
                last_start = len(joined) + 1
                chain.append(nxt)
                joined = f"{joined} {nxt.text.strip()}"
                for term, s, e in matcher.find(joined, ocr):
                    if s < last_start < e:  # crosses into the newest line; shorter chains found the rest
                        bbox = chain[0].bbox
                        for c in chain[1:]:
                            bbox = union(bbox, c.bbox)
                        found.append(Found(term, joined, bbox, bbox, _min_score(*(c.score for c in chain))))
                nxt = below[id(nxt)]
    return found


# ---------------------------------------------------------------- geometry

def union(a, b):
    return (min(a[0], b[0]), min(a[1], b[1]), max(a[2], b[2]), max(a[3], b[3]))


def overlaps(a, b, pad=2.0):
    return a[0] - pad < b[2] and b[0] - pad < a[2] and a[1] - pad < b[3] and b[1] - pad < a[3]


def iou(a, b):
    w, h = min(a[2], b[2]) - max(a[0], b[0]), min(a[3], b[3]) - max(a[1], b[1])
    if w <= 0 or h <= 0:
        return 0.0
    area = lambda r: (r[2] - r[0]) * (r[3] - r[1])
    return w * h / (area(a) + area(b) - w * h)


def center_in(bbox, rect):
    x, y = (bbox[0] + bbox[2]) / 2, (bbox[1] + bbox[3]) / 2
    return rect.x0 <= x <= rect.x1 and rect.y0 <= y <= rect.y1


def _mid(a, b):
    return pymupdf.Point((a.x + b.x) / 2, (a.y + b.y) / 2)


def _item_rect(item):
    if item[0] == "re":
        return item[1]
    if item[0] == "qu" and item[1].is_rectangular:
        return item[1].rect
    return None


def _item_points(item):
    if item[0] == "l":
        return [item[1], item[2], _mid(item[1], item[2])]
    if item[0] == "c":
        return list(item[1:5])
    r = _item_rect(item) or item[1].rect
    return [r.tl, r.tr, r.bl, r.br, _mid(r.tl, r.tr), _mid(r.bl, r.br), _mid(r.tl, r.bl), _mid(r.tr, r.br)]


def _segments(item):
    if item[0] == "l":
        return [(item[1], item[2])]
    r = _item_rect(item)
    if r is None:
        return None  # curve or rotated quad
    return [(r.tl, r.tr), (r.bl, r.br), (r.tl, r.bl), (r.tr, r.br)]


def _covered(intervals):
    total, end = 0.0, None
    for a, b in sorted(intervals):
        if end is None or a > end:
            total, end = total + b - a, b
        elif b > end:
            total, end = total + b - end, b
    return total


# ---------------------------------------------------------------- PDF pages

@dataclass
class Region:
    kind: str                 # "vector" or "image"
    rect: object              # pymupdf.Rect
    zoom: float
    report_text: bool = True  # False for full-page backgrounds/scans: their live text is body text
    where: str = ""
    ocr: bool = True          # False for vector diagrams whose words are all selectable text


def page_lines(page):
    """Visible text-layer lines; invisible OCR layers of scanned pages are dropped."""
    hidden = []
    try:
        hidden = [pymupdf.Rect(s["bbox"]) for s in page.get_texttrace()
                  if s.get("type") == 3 or s.get("opacity", 1) == 0]
    except Exception:
        pass
    lines = []
    for block in page.get_text("rawdict")["blocks"]:
        for line in block.get("lines", ()):
            chars = [ch for span in line["spans"] for ch in span["chars"]]
            text = "".join(ch["c"] for ch in chars)
            if not text.strip():
                continue
            r = pymupdf.Rect(line["bbox"])
            if not any(_mid(r.tl, r.br) in h for h in hidden):
                lines.append(Line(text, tuple(r), chars=[tuple(ch["bbox"]) for ch in chars]))
    return lines


def _captions(lines):
    """[(caption line, "figure"|"table")] joined with the title text on the same baseline, plus member indexes."""
    captions, members = [], set()
    for ln in lines:
        if not (FIGURE_CAPTION.match(ln.text) or TABLE_CAPTION.match(ln.text)):
            continue
        same = sorted((j for j, other in enumerate(lines)
                       if abs(other.bbox[1] - ln.bbox[1]) < 2 and other.bbox[0] >= ln.bbox[0]),
                      key=lambda j: lines[j].bbox[0])
        bbox = ln.bbox
        for j in same:
            bbox = union(bbox, lines[j].bbox)
        text = " ".join(lines[j].text.strip() for j in same)
        captions.append((Line(text, bbox), "table" if TABLE_CAPTION.match(ln.text) else "figure"))
        members.update(same)
    return captions, members


def _big_enough(r):
    return r.width >= 8 and r.height >= 6 and r.width * r.height >= 120


def _visible_path(path, page_area):
    if abs(pymupdf.Rect(path["rect"])) >= 0.85 * page_area:
        return False  # page background or full-page border
    kind = path.get("type") or ""
    fill = path.get("fill") if "f" in kind and (path.get("fill_opacity") or 1) > 0 else None
    stroke = path.get("color") if "s" in kind and (path.get("stroke_opacity") or 1) > 0 else None
    if fill is not None and stroke is None and min(fill) > 0.97:
        return False  # white mask
    return fill is not None or stroke is not None


def _within(rect, r, tol=1.0):
    rect = pymupdf.Rect(rect)
    return rect.x0 >= r.x0 - tol and rect.y0 >= r.y0 - tol and rect.x1 <= r.x1 + tol and rect.y1 <= r.y1 + tol


def _is_frame(paths, r, tol=2.0):
    """Every point lies on the outline: a border, shaded code block or text-box background."""
    def on_outline(pt):
        return min(abs(pt.x - r.x0), abs(pt.x - r.x1)) <= tol or min(abs(pt.y - r.y0), abs(pt.y - r.y1)) <= tol
    return all(on_outline(pt) for p in paths for item in p["items"] for pt in _item_points(item))


def _looks_like_table(paths, r):
    """A ruled grid: only horizontal/vertical strokes and several rules spanning the whole width."""
    rows, cols = {}, {}
    for p in paths:
        for item in p["items"]:
            segments = _segments(item)
            if segments is None:
                return False
            for a, b in segments:
                if abs(a.y - b.y) <= 0.5:
                    rows.setdefault(round(a.y), []).append((min(a.x, b.x), max(a.x, b.x)))
                elif abs(a.x - b.x) <= 0.5:
                    cols.setdefault(round(a.x), []).append((min(a.y, b.y), max(a.y, b.y)))
                else:
                    return False  # diagonal: arrow head, diamond, connector
    full_rows = sum(_covered(v) >= 0.9 * r.width for v in rows.values())
    full_cols = sum(_covered(v) >= 0.9 * r.height for v in cols.values())
    return full_rows >= 4 or (full_rows >= 3 and full_cols >= 2)


def _needs_ocr(paths):
    """Vector art whose words may not be selectable: outlined glyphs, logos, stroked CAD text."""
    curves = small = 0
    for p in paths:
        if len(p["items"]) >= 10:
            return True
        curves += sum(1 for item in p["items"] if item[0] == "c")
        r = pymupdf.Rect(p["rect"])
        small += max(r.width, r.height) < 10
    return curves >= 20 or small >= 30


def _caption_for(r, captions, max_gap=50):
    """Nearest figure caption above/below, or table caption above, the region."""
    best, kind, best_gap = "", "", max_gap
    for ln, k in captions:
        x0, y0, x1, y1 = ln.bbox
        if x1 < r.x0 - 30 or x0 > r.x1 + 30:
            continue
        if y1 <= r.y0 + 2:
            gap = r.y0 - y1
        elif y0 >= r.y1 - 2 and k == "figure":
            gap = y0 - r.y1
        else:
            continue
        if gap <= best_gap:
            best, kind, best_gap = ln.text.strip(), k, gap
    return best, kind


def _attach_labels(r, labels, max_gap=20, max_len=40):
    """Grow r over short text labels beside it (bus names left of a brace) or touching its top/bottom."""
    r = pymupdf.Rect(r)
    pending = [ln for ln in labels if len(ln.text.strip()) <= max_len]
    while True:
        for ln in pending:
            x0, y0, x1, y1 = ln.bbox
            beside = r.y0 <= (y0 + y1) / 2 <= r.y1 and r.x0 - x1 <= max_gap and x0 - r.x1 <= max_gap
            touching = r.x0 <= x0 and x1 <= r.x1 and (r.y0 - 4 <= y1 <= r.y0 + 2 or r.y1 - 2 <= y0 <= r.y1 + 4)
            if beside or touching:
                r |= ln.bbox
                pending.remove(ln)
                break
        else:
            return r


def _band(page, r):
    h = page.rect.height
    if r.y1 <= page.rect.y0 + 0.12 * h:
        return "page header"
    if r.y0 >= page.rect.y1 - 0.12 * h:
        return "page footer"
    return ""


def find_regions(page, captions, labels, paths, images, zoom, ocr_all=False):
    """Raster images plus clusters of vector drawing that look like diagrams (not tables or frames)."""
    area = abs(page.rect)
    regions = []
    for info in images:
        r = pymupdf.Rect(info["bbox"]) & page.rect
        if r.is_empty or not _big_enough(r):
            continue
        native = max(info["width"] / r.width, info["height"] / r.height)
        foreground = abs(r) < 0.6 * area
        where = _band(page, r) or _caption_for(r, captions)[0] or ("" if foreground else "full-page image (scan or background)")
        regions.append(Region("image", _attach_labels(r, labels) if foreground else r,
                              zoom=min(max(native, 2.0), 4.0), report_text=foreground, where=where))
    if len(paths) > MAX_PATHS:
        r = pymupdf.Rect()
        for p in paths:
            r |= p["rect"]
        r &= page.rect
        regions.append(Region("vector", r, zoom=zoom, where=_band(page, r) or _caption_for(r, captions)[0]))
    elif paths:
        for r in page.cluster_drawings(drawings=paths):
            if not _big_enough(r):
                continue
            members = [p for p in paths if _within(p["rect"], r)]
            band = _band(page, r)
            caption, kind = ("", "") if band else _caption_for(r, captions)
            if not members or _is_frame(members, r) or kind == "table" \
                    or (kind != "figure" and _looks_like_table(members, r)):
                continue
            regions.append(Region("vector", _attach_labels(r, labels), zoom=zoom, where=band or caption,
                                  ocr=ocr_all or _needs_ocr(members)))
    return regions


def _signature(clip, zoom, paths, lines, images):
    """Content key for a clip, stable across pages where the same art sits at a slightly different spot."""
    ox, oy, box = clip.x0, clip.y0, tuple(clip)
    rel = lambda x, y: (round(x - ox, 1), round(y - oy, 1))
    parts = [round(clip.width, 1), round(clip.height, 1), zoom]
    for p in paths:
        if overlaps(tuple(p["rect"]), box, 0):
            parts.append((p.get("type"), p.get("fill"), p.get("color"),
                          [rel(pt.x, pt.y) for item in p["items"] for pt in _item_points(item)]))
    for info in images:
        if overlaps(tuple(info["bbox"]), box, 0):
            parts.append((info.get("digest"), rel(info["bbox"][0], info["bbox"][1])))
    for ln in lines:
        if overlaps(ln.bbox, box, 0):
            parts.append((ln.text, rel(ln.bbox[0], ln.bbox[1])))
    return hashlib.sha1(repr(parts).encode()).hexdigest()


def _pix_array(pix):
    import numpy as np
    return np.frombuffer(pix.samples, dtype=np.uint8).reshape(pix.height, pix.width, pix.n)


def _dedupe(hits):
    kept = []
    for h in sorted(hits, key=lambda h: h.source != "text"):
        if not any(k.term == h.term and iou(k.bbox, h.bbox) > 0.3 for k in kept):
            kept.append(h)
    return kept


def scan_page(page, page_no, ctx, whole_page=None, approx=False):
    """Search one PDF/SVG page; whole_page treats the entire page as a single graphic."""
    if page.rotation:
        page.remove_rotation()
    lines = page_lines(page)
    area = abs(page.rect)
    paths = [p for p in page.get_drawings() if _visible_path(p, area)]
    images = page.get_image_info(hashes=True)
    captions, caption_lines = _captions(lines)
    labels = [ln for i, ln in enumerate(lines) if i not in caption_lines]
    regions = [whole_page] if whole_page else find_regions(page, captions, labels, paths, images, ctx.zoom, ctx.ocr_all)
    body = ctx.text and whole_page is None
    if not regions and not body:
        return []
    all_text = search_lines(ctx.matcher, lines)
    label_text = search_lines(ctx.matcher, labels)
    hits, in_graphics = [], []
    ctx.checked += len(regions)
    for reg in regions:
        clip = (reg.rect + (-LABEL_MARGIN, -LABEL_MARGIN, LABEL_MARGIN, LABEL_MARGIN)) & page.rect
        found = [(f, "text") for f in label_text if center_in(f.bbox, clip)] if reg.report_text else []
        in_graphics += [f for f, _ in found]
        zoom = min(reg.zoom, MAX_OCR_SIDE / max(clip.width, clip.height))
        render = lambda: page.get_pixmap(clip=clip, matrix=pymupdf.Matrix(zoom, zoom),
                                         colorspace=pymupdf.csRGB, alpha=False)
        key = None
        if ctx.ocr and reg.ocr:
            key = _signature(clip, zoom, paths, lines, images)
            to_page = lambda b: (clip.x0 + b[0] / zoom, clip.y0 + b[1] / zoom, clip.x0 + b[2] / zoom, clip.y0 + b[3] / zoom)
            ocr_lines = [Line(ln.text, to_page(ln.bbox), ln.score, ln.chars and [c and to_page(c) for c in ln.chars])
                         for ln in ctx.ocr.read(key, lambda: _pix_array(render()))]
            for f in search_lines(ctx.matcher, ocr_lines, ocr=True):
                if not any(t.term == f.term and overlaps(t.line_bbox, f.bbox) for t in all_text):
                    found.append((f, "ocr"))
        region_hits = [Hit(f.term, page_no, f.text.strip(), source, reg.kind, reg.where,
                           [round(v, 1) for v in f.bbox], None if f.score is None else round(f.score, 2), approx)
                       for f, source in found]
        if region_hits and (ctx.crops or ctx.report):
            def image():
                from PIL import Image
                pix = render()
                return Image.frombytes("RGB", (pix.width, pix.height), pix.samples)
            to_px = lambda b: ((b[0] - clip.x0) * zoom, (b[1] - clip.y0) * zoom,
                               (b[2] - clip.x0) * zoom, (b[3] - clip.y0) * zoom)
            thumb = None
            if whole_page is None and abs(clip) < 0.6 * area:
                thumb = lambda: _page_thumb(page, clip, region_hits, ctx)
            record_graphic(ctx, key or _signature(clip, zoom, paths, lines, images), image, region_hits, to_px,
                           f"page{page_no:03d}" if page_no else "graphic", thumb)
        hits += region_hits
    hits = _dedupe(hits)
    if body:
        ctx.text_checked += len(lines)
        for f in all_text:
            if any(g.term == f.term and iou(g.bbox, f.bbox) > 0.5 for g in in_graphics):
                continue
            hits.append(Hit(f.term, page_no, f.text.strip(), "body", "text", _text_where(page, f),
                            [round(v, 1) for v in f.bbox], page_approx=approx))
    return hits


def _text_where(page, f):
    if FIGURE_CAPTION.match(f.text):
        return "figure caption"
    if TABLE_CAPTION.match(f.text):
        return "table caption"
    # tighter than _band: body text of a Word page starts 1 inch (9%) from the top
    top, h = page.rect.y0, page.rect.height
    if f.line_bbox[3] <= top + 0.095 * h:
        return "page header"
    if f.line_bbox[1] >= top + 0.915 * h:
        return "page footer"
    return ""


def parse_pages(spec, count):
    if not spec:
        return list(range(count))
    pages = set()
    for part in spec.split(","):
        part = part.strip()
        if not part:
            continue
        lo, dash, hi = part.partition("-")
        lo = int(lo or 1)
        hi = int(hi) if hi else (count if dash else lo)
        pages.update(range(max(lo, 1) - 1, min(hi, count)))
    return sorted(pages)


def scan_pdf(path, ctx, pages_spec, info):
    with pymupdf.open(path) as doc:
        if doc.needs_pass:
            die("the PDF is password protected")
        info["pages"] = doc.page_count
        selected = parse_pages(pages_spec, doc.page_count)
        info["scanned_pages"] = len(selected)
        hits = []
        for i, pno in enumerate(selected, 1):
            page = doc[pno]
            hits += scan_page(page, pno + 1, ctx)
            if ctx.links:
                hits += scan_page_links(page, pno + 1, ctx)
            if len(selected) > 20 and i % 20 == 0:
                print(f"  scanned {i}/{len(selected)} pages", file=sys.stderr, flush=True)
        record_text_shots(doc, [h for h in hits if h.source == "body"], ctx)
    return hits


def _render(page, clip, zoom):
    from PIL import Image
    pix = page.get_pixmap(clip=clip, matrix=pymupdf.Matrix(zoom, zoom), colorspace=pymupdf.csRGB, alpha=False)
    return Image.frombytes("RGB", (pix.width, pix.height), pix.samples)


def record_text_shots(doc, hits, ctx):
    """Screenshots for body-text matches: each page with its matches boxed, except text repeated on many pages
    (running headers and footers), which gets one strip of its first occurrence."""
    if not (ctx.crops or ctx.report):
        return
    mark_repeated(hits)
    by_page, repeated = {}, {}
    for h in hits:
        if h.repeated:
            repeated.setdefault((h.text.lower(), h.where), []).append(h)
        else:
            by_page.setdefault(h.page, []).append(h)
    for page_no, hs in by_page.items():
        page = doc[page_no - 1]
        to_px = lambda b, o=page.rect: tuple((v - p) * TEXT_ZOOM for v, p in zip(b, (o.x0, o.y0, o.x0, o.y0)))
        record_graphic(ctx, ("text page", page_no), lambda page=page: _render(page, page.rect, TEXT_ZOOM), hs,
                       to_px, f"page{page_no:03d}_text", photo=True)
    for hs in repeated.values():
        first = min(h.page for h in hs)
        page = doc[first - 1]
        marks = [(h.term, h.bbox) for h in hs if h.page == first]
        area = pymupdf.Rect(marks[0][1])
        for _, b in marks[1:]:
            area |= b
        clip = (pymupdf.Rect(page.rect.x0 + 36, area.y0 - 14, page.rect.x1 - 36, area.y1 + 14) | area) & page.rect
        to_px = lambda b, c=clip: tuple((v - o) * LINK_ZOOM for v, o in zip(b, (c.x0, c.y0, c.x0, c.y0)))
        record_graphic(ctx, ("text strip", hs[0].text.lower(), tuple(round(v) for v in area)),
                       lambda page=page, clip=clip: _render(page, clip, LINK_ZOOM), hs, to_px,
                       f"page{first:03d}_text", marks=marks)


# ---------------------------------------------------------------- hyperlinks

def link_spans(matcher, target):
    """(term, start, end) in a link target; URL punctuation such as / . - _ = + separates words, and the words
    of a phrase may run together as in URL slugs ("intel-fpga", "IntelFPGA")."""
    return list(matcher.find(re.sub(r"[\W_]", " ", target), joined=True))


def link_hits(ctx, target, page, where, bbox=None, approx=False):
    """One hit per term found in a hyperlink target (URL, e-mail address or file path)."""
    ctx.links_checked += 1
    target = urllib.parse.unquote(target.strip())
    terms = dict.fromkeys(t for t, _, _ in link_spans(ctx.matcher, target))
    return [Hit(t, page, target, "link", "link", where, bbox, page_approx=approx) for t in terms]


def link_where(text, place=""):
    text = " ".join(text.split())
    label = f'link text "{text if len(text) <= 70 else text[:67] + "..."}"' if text else "linked picture or shape"
    return f"{place}, {label}" if place else label


def _link_pieces(rects, gap=4):
    """Group the rectangles of one link target into links: a link that wraps has one rectangle per line."""
    groups = []
    for r in sorted(rects, key=lambda r: (r.y0, r.x0)):
        if groups and r.y0 - groups[-1][-1].y1 <= gap:
            groups[-1].append(r)
        else:
            groups.append([r])
    return groups


def scan_page_links(page, page_no, ctx):
    """Hyperlinks on a PDF page whose target contains a term; each is boxed on a strip of the page."""
    targets = {}
    for link in page.get_links():
        if target := link.get("uri") or link.get("file"):  # other kinds jump within the document
            targets.setdefault(target, []).append(pymupdf.Rect(link["from"]))
    hits, words = [], None
    for target, rects in targets.items():
        for pieces in _link_pieces(rects):
            area = pymupdf.Rect(pieces[0])
            for r in pieces[1:]:
                area |= r
            found = link_hits(ctx, target, page_no, "", [round(v, 1) for v in area])
            if not found:
                continue
            if words is None:
                words = page.get_text("words")
            # PowerPoint also writes an invisible copy of the text about 1 pt high at the link's corner
            text = " ".join(w[4] for r in pieces for w in words if center_in(w[:4], r) and w[3] - w[1] >= 2)
            for h in found:
                h.where = link_where(text)
            clip = (pymupdf.Rect(page.rect.x0 + 36, area.y0 - 14, page.rect.x1 - 36, area.y1 + 14) | area) & page.rect
            if (ctx.crops or ctx.report) and not clip.is_empty:
                def image():
                    from PIL import Image
                    pix = page.get_pixmap(clip=clip, matrix=pymupdf.Matrix(LINK_ZOOM, LINK_ZOOM),
                                          colorspace=pymupdf.csRGB, alpha=False)
                    return Image.frombytes("RGB", (pix.width, pix.height), pix.samples)
                to_px = lambda b: tuple((v - o) * LINK_ZOOM for v, o in zip(b, (clip.x0, clip.y0, clip.x0, clip.y0)))
                record_graphic(ctx, ("link", found[0].text, tuple(round(v) for v in area)), image, found, to_px,
                               f"page{page_no:03d}_link", marks=[(h.term, r) for h in found for r in pieces])
            hits += found
    return hits


# ---------------------------------------------------------------- OCR and standalone graphics

class Ocr:
    def __init__(self, min_score):
        import importlib.util
        if importlib.util.find_spec("rapidocr") is None:
            die("OCR needs RapidOCR: pip install rapidocr onnxruntime   (or pass --no-ocr)")
        self.min_score = min_score
        self.engine = None
        self.cache = {}

    def read(self, key, get_rgb):
        """OCR the HxWx3 RGB array from get_rgb(); returns Lines in pixel coordinates."""
        if key not in self.cache:
            import numpy as np
            if self.engine is None:
                from rapidocr import RapidOCR
                self.engine = RapidOCR(params={"Global.log_level": "critical", "Global.text_score": self.min_score,
                                               "Global.max_side_len": MAX_OCR_SIDE})
            result = self.engine(np.ascontiguousarray(get_rgb()[:, :, ::-1]),
                                 return_word_box=True, return_single_char_box=True)
            box_of = lambda pts: (float(np.min(pts[:, 0])), float(np.min(pts[:, 1])),
                                  float(np.max(pts[:, 0])), float(np.max(pts[:, 1])))
            out = []
            if result is not None and result.boxes is not None and result.txts is not None:
                words = getattr(result, "word_results", None) or [None] * len(result.txts)
                for box, text, score, glyphs in zip(result.boxes, result.txts, result.scores, words):
                    out.append(Line(text, box_of(np.asarray(box)), float(score), self._align(text, glyphs, box_of)))
            self.cache[key] = out
        return self.cache[key]

    @staticmethod
    def _align(text, glyphs, box_of):
        """Per-character boxes for text, or None when RapidOCR's glyph list does not line up."""
        import numpy as np
        slots = [i for i, ch in enumerate(text) if not ch.isspace()]
        if not glyphs or len(glyphs) != len(slots) or any(len(g[0]) != 1 for g in glyphs):
            return None
        chars = [None] * len(text)
        for i, glyph in zip(slots, glyphs):
            chars[i] = box_of(np.asarray(glyph[2], dtype=float))
        return chars


@dataclass
class Context:
    matcher: Matcher
    ocr: Ocr | None
    zoom: float = 3.0
    crops: str | None = None
    ocr_all: bool = False
    report: bool = False
    links: bool = True                                # also search hyperlink targets
    text: bool = True                                 # also search body text
    warnings: list = field(default_factory=list)
    skipped_images: int = 0
    checked: int = 0                                  # graphics examined
    links_checked: int = 0
    text_checked: int = 0                             # lines or paragraphs of body text
    graphics: dict = field(default_factory=dict)      # annotated screenshots by content key
    hit_graphic: dict = field(default_factory=dict)   # id(hit) -> key into graphics

    def color(self, term):
        return TERM_COLORS[self.matcher.terms.index(term) % len(TERM_COLORS)]


@dataclass
class Graphic:
    """Annotated screenshot of one graphic; the same graphic on several pages shares one."""
    image: bytes | None = None  # PNG or JPEG embedded in the HTML report
    mime: str = "image/png"
    size: tuple = (0, 0)
    thumb: bytes | None = None  # JPEG of the whole page with the graphic outlined
    path: str | None = None     # file written for --crops


def _rgb(color):
    return tuple(int(color[i:i + 2], 16) for i in (1, 3, 5))


def _encode(img, photo=False):
    """(bytes, mime): PNG keeps diagram text crisp; JPEG for thumbnails and very large screenshots."""
    buf = io.BytesIO()
    if photo or img.width * img.height > 1_500_000:
        img.convert("RGB").save(buf, "JPEG", quality=72 if photo else 85)
        return buf.getvalue(), "image/jpeg"
    img.save(buf, "PNG", optimize=True)
    return buf.getvalue(), "image/png"


def _draw_boxes(img, boxes, scale=1.0, width=3):
    from PIL import ImageDraw
    draw = ImageDraw.Draw(img)
    pad = width + 1
    for color, (x0, y0, x1, y1) in boxes:
        box = (max(x0 * scale - pad, 1), max(y0 * scale - pad, 1),
               min(x1 * scale + pad, img.width - 2), min(y1 * scale + pad, img.height - 2))
        if box[2] > box[0] and box[3] > box[1]:
            draw.rectangle(box, outline=_rgb(color), width=width)
    return img


def record_graphic(ctx, key, get_image, hits, to_px, stem, get_thumb=None, marks=None, photo=False):
    """Screenshot get_image() with each hit boxed in its term colour, for --crops and --html.

    marks, if given, are the (term, bbox) boxes to draw instead of the hits' own boxes; photo embeds a JPEG.
    """
    if not (ctx.crops or ctx.report):
        return
    from PIL import Image
    if marks is None:
        marks = [(h.term, h.bbox) for h in hits if h.bbox]
    boxes = tuple(sorted((ctx.color(term), tuple(round(v) for v in to_px(b))) for term, b in marks))
    key = (key, boxes)
    g = ctx.graphics.get(key)
    if g is None:
        g = ctx.graphics[key] = Graphic()
        img = get_image()
        scale = min(1.0, DISPLAY_SIDE / max(img.size))
        img = img.resize((max(1, round(img.width * scale)), max(1, round(img.height * scale))), Image.LANCZOS) \
            if scale < 1 else img.copy()
        _draw_boxes(img, boxes, scale)
        if ctx.crops:
            os.makedirs(ctx.crops, exist_ok=True)
            g.path = os.path.abspath(os.path.join(ctx.crops, f"{stem}_{len(ctx.graphics):03d}.png"))
            img.save(g.path)
        if ctx.report:
            g.image, g.mime = _encode(img, photo)
            g.size = img.size
            g.thumb = get_thumb() if get_thumb else None
    for h in hits:
        h.crop = g.path
        ctx.hit_graphic[id(h)] = key


def _page_thumb(page, clip, hits, ctx):
    """JPEG of the whole page: the graphic outlined in amber, matches in their term colours."""
    from PIL import Image, ImageDraw
    zoom = THUMB_WIDTH / page.rect.width
    pix = page.get_pixmap(matrix=pymupdf.Matrix(zoom, zoom), colorspace=pymupdf.csRGB, alpha=False)
    img = Image.frombytes("RGB", (pix.width, pix.height), pix.samples)
    ox, oy = page.rect.x0, page.rect.y0
    px = lambda b: ((b[0] - ox) * zoom, (b[1] - oy) * zoom, (b[2] - ox) * zoom, (b[3] - oy) * zoom)
    ImageDraw.Draw(img).rectangle(px(tuple(clip)), outline=(245, 158, 11), width=3)
    _draw_boxes(img, [(ctx.color(h.term), px(h.bbox)) for h in hits if h.bbox], width=2)
    return _encode(img, photo=True)[0]


def _is_svg(data, name=""):
    if name.lower().split("?")[0].endswith((".svg", ".svgz")):
        return True
    if data[:2] == b"\x1f\x8b":
        try:
            data = gzip.GzipFile(fileobj=io.BytesIO(data)).read(2048)
        except (OSError, EOFError):
            return False
    head = data[:2048].lstrip().lower()
    return head.startswith((b"<svg", b"<?xml", b"<!doctype svg")) and b"<svg" in head


def scan_svg(data, ctx, where, page=None, approx=False):
    if data[:2] == b"\x1f\x8b":
        try:
            data = gzip.decompress(data)
        except (OSError, EOFError) as e:
            ctx.warnings.append(f"{where}: compressed SVG could not be unpacked ({e}); not checked")
            return []
    if b"xmlns=" not in data[:4096]:
        data = data.replace(b"<svg", b'<svg xmlns="http://www.w3.org/2000/svg"', 1)
    try:
        doc = pymupdf.open(stream=data, filetype="svg")
    except Exception:
        text = re.sub(r"<[^>]+>", " ", data.decode("utf-8", "replace"))
        lines = [Line(t, (0, 0, 0, 0)) for t in text.splitlines() if t.strip()]
        ctx.warnings.append(f"{where}: SVG could not be rendered; only its text was searched")
        return [Hit(f.term, page, f.text.strip(), "text", "vector", where, page_approx=approx)
                for f in search_lines(ctx.matcher, lines)]
    pg = doc[0]
    zoom = min(max(2000 / max(pg.rect.width, pg.rect.height, 1), 1.0), 6.0)
    return scan_page(pg, page, ctx, whole_page=Region("vector", pg.rect, zoom=zoom, where=where), approx=approx)


def _flatten(img):
    from PIL import Image
    if img.mode in ("RGBA", "LA", "P", "PA") or "transparency" in img.info:
        img = img.convert("RGBA")
        bg = Image.new("RGB", img.size, (255, 255, 255))
        bg.paste(img, mask=img.getchannel("A"))
        return bg
    return img.convert("RGB")


def _open_image(data):
    from PIL import Image
    img = Image.open(io.BytesIO(data))
    if img.format in ("WMF", "EMF"):
        img.load(dpi=200)  # metafiles are vector; the default 72 dpi is too coarse to read
    img.load()
    return img


def scan_raster(data, ctx, where, page=None, approx=False, img=None):
    """OCR a picture given as encoded bytes, or as an already decoded PIL image (one frame of a file)."""
    if ctx.ocr is None:
        ctx.skipped_images += 1
        return []
    import numpy as np
    from PIL import Image, ImageOps
    if img is None:
        try:
            img = _open_image(data)
        except Exception as e:
            ctx.warnings.append(f"{where}: image could not be decoded ({e.__class__.__name__}); not checked")
            return []
    ctx.checked += 1
    try:
        img = ImageOps.exif_transpose(img)
    except Exception:
        pass
    img = _flatten(img)
    key = hashlib.sha1(data if data is not None else img.tobytes()).hexdigest()
    scale = 1.0
    if max(img.size) < 800:
        scale = 2.0
    elif max(img.size) > MAX_OCR_SIDE:
        scale = MAX_OCR_SIDE / max(img.size)
    if scale != 1.0:
        img = img.resize((max(1, round(img.width * scale)), max(1, round(img.height * scale))), Image.LANCZOS)
    found = search_lines(ctx.matcher, ctx.ocr.read(key, lambda: np.asarray(img)), ocr=True)
    hits = [Hit(f.term, page, f.text.strip(), "ocr", "image", where, [round(v / scale, 1) for v in f.bbox],
                round(f.score, 2), approx) for f in found]
    if hits:
        record_graphic(ctx, key, lambda: img, hits, lambda b: tuple(v * scale for v in b),
                       f"frame{page:03d}" if page else "image")
    return hits


def scan_blob(data, name, ctx, where, page=None, approx=False):
    if _is_svg(data, name):
        return scan_svg(data, ctx, where, page, approx)
    return scan_raster(data, ctx, where, page, approx)


def _svg_links(data):
    """(href, link text) of the <a> elements in SVG markup."""
    try:
        root = ET.fromstring(gzip.decompress(data) if data[:2] == b"\x1f\x8b" else data)
    except (ET.ParseError, OSError, EOFError):
        return []
    links = []
    for el in root.iter():
        href = el.get("href") or el.get("{http://www.w3.org/1999/xlink}href")
        if _local(el.tag) == "a" and href:
            links.append((href, "".join(el.itertext())))
    return links


def looks_like_image(path):
    """Content sniffing for image files with an unfamiliar extension."""
    with open(path, "rb") as f:
        head = f.read(4096)
    if _is_svg(head):
        return True
    try:
        from PIL import Image
        with Image.open(path):
            return True
    except Exception:
        return False


def scan_image_file(path, ctx, pages_spec, info):
    """A standalone SVG or picture; each frame of a multi-page TIFF or animated GIF counts as a page."""
    with open(path, "rb") as f:
        data = f.read()
    name = os.path.basename(path)
    if _is_svg(data, path):
        info["type"] = "svgz" if data[:2] == b"\x1f\x8b" else "svg"
        hits = scan_svg(data, ctx, name)
        if ctx.links:
            hits += [h for n, (href, text) in enumerate(_svg_links(data), 1)
                     for h in link_hits(ctx, href, None, link_where(text, f"<a> #{n}"))]
        return hits
    try:
        img = _open_image(data)
    except Exception as e:
        die(f"{name} could not be read as an image ({e})")
    info["type"] = (img.format or info["type"]).lower()
    turned = img.getexif().get(0x0112) in (5, 6, 7, 8)  # EXIF orientation: displayed rotated by 90 degrees
    info["size"] = f"{img.height} x {img.width} px" if turned else f"{img.width} x {img.height} px"
    frames = getattr(img, "n_frames", 1)
    if frames == 1:
        return scan_raster(data, ctx, name, img=img)
    info["pages"] = frames
    selected = parse_pages(pages_spec, frames)
    info["scanned_pages"] = len(selected)
    hits = []
    for i in selected:
        img.seek(i)
        hits += scan_raster(None, ctx, name, page=i + 1, img=img.copy())
    return hits


# ---------------------------------------------------------------- HTML

HTML_BLOCKS = {"p", "div", "li", "ul", "ol", "dl", "dt", "dd", "h1", "h2", "h3", "h4", "h5", "h6", "table", "tr",
               "td", "th", "caption", "thead", "tbody", "tfoot", "section", "article", "header", "footer", "nav",
               "aside", "main", "blockquote", "pre", "figure", "figcaption", "form", "fieldset", "legend", "label",
               "button", "option", "select", "textarea", "title", "body", "hr", "address", "details", "summary"}
HTML_HIDDEN = {"script", "style", "noscript", "template"}


class _HtmlAssets(HTMLParser):
    def __init__(self, raw):
        super().__init__(convert_charrefs=True)
        self.raw = raw
        self._line_starts = [0] + [m.end() for m in re.finditer("\n", raw)]
        self.assets = []  # ("src", url) or ("svg", markup), in document order
        self.links = []   # [href, link text], in document order
        self.blocks = []  # (visible text, where), in document order
        self._svg_depth, self._svg_start = 0, 0
        self._link = None
        self._hidden, self._buf, self._block, self._heading = 0, [], None, ""

    def _offset(self):
        line, col = self.getpos()
        return self._line_starts[line - 1] + col

    def _flush(self):
        text = " ".join("".join(self._buf).split())
        self._buf = []
        if not text:
            return
        if self._block == "title":
            where = "page title"
        elif self._block in ("h1", "h2", "h3", "h4", "h5", "h6"):
            where, self._heading = "heading", text
        else:
            where = f'section "{_clip(self._heading, 60)}"' if self._heading else ""
        self.blocks.append((text, where))

    def handle_data(self, data):
        if self._link is not None:
            self._link[1] += data
        if not (self._hidden or self._svg_depth):
            self._buf.append(data)

    def handle_starttag(self, tag, attrs):
        if tag in HTML_HIDDEN:
            self._hidden += 1
        elif tag in HTML_BLOCKS:
            self._flush()
            self._block = tag
        elif tag in ("br", "img", "wbr"):
            self._buf.append(" ")
        self._tag(tag, attrs)

    def _tag(self, tag, attrs):
        a = dict(attrs)
        href = (a.get("href") or a.get("xlink:href") or "").strip()
        if tag in ("a", "area") and href:
            text = (a.get("alt") or "") if tag == "area" else ""
            self.links.append([href, text])
            self._link = self.links[-1] if tag == "a" else None
        elif tag == "img" and self._link is not None:
            self._link[1] += " " + (a.get("alt") or "")
        if tag == "svg":
            if self._svg_depth == 0:
                self._svg_start = self._offset()
            self._svg_depth += 1
        elif self._svg_depth == 0:
            if tag in ("img", "embed"):
                src = a.get("src") or a.get("data-src")
            else:
                src = a.get("data") if tag == "object" else None
            if src and src.strip():
                self.assets.append(("src", src.strip()))

    def handle_endtag(self, tag):
        if tag in HTML_HIDDEN:
            self._hidden = max(0, self._hidden - 1)
        elif tag in HTML_BLOCKS:
            self._flush()
            self._block = None
        if tag == "a":
            self._link = None
        if tag == "svg" and self._svg_depth:
            self._svg_depth -= 1
            if self._svg_depth == 0:
                self.assets.append(("svg", self.raw[self._svg_start:self.raw.find(">", self._offset()) + 1]))


def _clip(text, n):
    return text if len(text) <= n else text[:n - 3] + "..."


def _read_html(path):
    data = open(path, "rb").read()
    m = re.search(rb"""<meta[^>]+charset=["']?([\w-]+)""", data[:4096], re.I)
    try:
        return data.decode(m.group(1).decode() if m else "utf-8", "replace")
    except LookupError:
        return data.decode("utf-8", "replace")


def _load_src(src, base_dir):
    if src.startswith("data:"):
        header, _, payload = src.partition(",")
        return base64.b64decode(payload) if ";base64" in header else urllib.parse.unquote_to_bytes(payload)
    parsed = urllib.parse.urlparse(src)
    if parsed.scheme in ("http", "https"):
        req = urllib.request.Request(src, headers={"User-Agent": "Mozilla/5.0"})
        with urllib.request.urlopen(req, timeout=20) as resp:
            return resp.read()
    if parsed.scheme == "file":
        path = urllib.request.url2pathname(parsed.path)
    elif len(parsed.scheme) == 1:  # Windows drive letter, e.g. C:\img.png
        path = src
    else:
        path = os.path.join(base_dir, urllib.parse.unquote(parsed.path))
    with open(path, "rb") as f:
        return f.read()


def _short(src, n=60):
    if src.startswith("data:"):
        return src.split(",", 1)[0] + ",..."
    return src if len(src) <= n else "..." + src[-(n - 3):]


def scan_html(path, ctx, info):
    raw = _read_html(path)
    assets = _HtmlAssets(raw)
    assets.feed(raw)
    assets.close()
    assets._flush()
    counts = {"src": 0, "svg": 0}
    base = os.path.dirname(path)
    hits = []
    for kind, value in assets.assets:
        counts[kind] += 1
        if kind == "svg":
            hits += scan_svg(value.encode("utf-8"), ctx, f"inline <svg> #{counts[kind]}")
            continue
        where = f"<img> #{counts[kind]} {_short(value)}"
        try:
            data = _load_src(value, base)
        except Exception as e:
            ctx.warnings.append(f"{where}: could not be loaded ({e}); not checked")
            continue
        hits += scan_blob(data, value, ctx, where)
    info.update(images=counts["src"], inline_svgs=counts["svg"])
    if ctx.links:
        info["links"] = len(assets.links)
        for n, (href, text) in enumerate(assets.links, 1):
            hits += link_hits(ctx, href, None, link_where(text, f"<a> #{n}"))
    if ctx.text:
        hits += body_hits(ctx, assets.blocks, None)
    return hits


def body_hits(ctx, blocks, page, approx=False):
    """Hits in (text, where) blocks of body text that have no page picture (HTML, DOCX/PPTX read directly)."""
    ctx.text_checked += len(blocks)
    return [Hit(f.term, page, f.text.strip(), "body", "text", where, page_approx=approx)
            for text, where in blocks for f in search_lines(ctx.matcher, [Line(text, (0, 0, 0, 0))])]


# ---------------------------------------------------------------- Office

_PS_WORD = r"""
$ErrorActionPreference = 'Stop'
$app = New-Object -ComObject Word.Application
try {
    $app.Visible = $false
    $app.DisplayAlerts = 0
    $doc = $app.Documents.Open($env:DOCOCR_IN, $false, $true, $false)
    $doc.ExportAsFixedFormat($env:DOCOCR_OUT, 17)
    $doc.Close(0)
} finally {
    if ($app.Documents.Count -eq 0) { $app.Quit() }
}
"""

_PS_POWERPOINT = r"""
$ErrorActionPreference = 'Stop'
$app = New-Object -ComObject PowerPoint.Application
try {
    $pres = $app.Presentations.Open($env:DOCOCR_IN, -1, 0, 0)
    foreach ($slide in $pres.Slides) { $slide.SlideShowTransition.Hidden = 0 }
    $pres.SaveAs($env:DOCOCR_OUT, 32)
    $pres.Close()
} finally {
    if ($app.Presentations.Count -eq 0) { $app.Quit() }
}
"""


def _find_soffice():
    for name in ("soffice", "libreoffice"):
        if found := shutil.which(name):
            return found
    for candidate in (r"C:\Program Files\LibreOffice\program\soffice.exe",
                      r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
                      "/Applications/LibreOffice.app/Contents/MacOS/soffice"):
        if os.path.exists(candidate):
            return candidate
    return None


def convert_to_pdf(path, outdir, slides):
    """Return (pdf_path, tool). Microsoft Office is tried first on Windows, then LibreOffice."""
    out = os.path.join(outdir, os.path.splitext(os.path.basename(path))[0] + ".pdf")
    errors = []
    if os.name == "nt":
        script = os.path.join(outdir, "convert.ps1")
        with open(script, "w", encoding="utf-8") as f:
            f.write(_PS_POWERPOINT if slides else _PS_WORD)
        try:
            proc = subprocess.run(["powershell", "-NoProfile", "-NonInteractive", "-ExecutionPolicy", "Bypass",
                                   "-File", script], env=dict(os.environ, DOCOCR_IN=path, DOCOCR_OUT=out),
                                  capture_output=True, text=True, timeout=900)
            if os.path.exists(out):
                return out, "Microsoft PowerPoint" if slides else "Microsoft Word"
            errors.append("Microsoft Office: " + ((proc.stderr or proc.stdout).strip().splitlines() or ["no PDF produced"])[0])
        except (OSError, subprocess.TimeoutExpired) as e:
            errors.append(f"Microsoft Office: {e}")
    soffice = _find_soffice()
    if soffice:
        targets = ['pdf:impress_pdf_Export:{"ExportHiddenSlides":{"type":"boolean","value":"true"}}', "pdf"] if slides else ["pdf"]
        for target in targets:
            subprocess.run([soffice, "--headless", "--convert-to", target, "--outdir", outdir, path],
                           capture_output=True, timeout=900)
            if os.path.exists(out):
                return out, "LibreOffice"
        errors.append("LibreOffice: no PDF produced")
    else:
        errors.append("LibreOffice not found")
    raise RuntimeError("; ".join(errors))


def _local(tag):
    return tag.rsplit("}", 1)[-1]


def _rels(z, part):
    folder, name = posixpath.split(part)
    try:
        root = ET.fromstring(z.read(posixpath.join(folder, "_rels", name + ".rels")))
    except KeyError:
        return {}
    out = {}
    for rel in root:
        if rel.get("TargetMode") == "External":
            continue
        target = rel.get("Target", "")
        target = target.lstrip("/") if target.startswith("/") else posixpath.normpath(posixpath.join(folder, target))
        out[rel.get("Id")] = (rel.get("Type", "").rsplit("/", 1)[-1], target)
    return out


def _hyperlink_rels(z, part):
    """rId -> target of the part's external hyperlinks."""
    folder, name = posixpath.split(part)
    try:
        root = ET.fromstring(z.read(posixpath.join(folder, "_rels", name + ".rels")))
    except KeyError:
        return {}
    return {rel.get("Id"): rel.get("Target", "") for rel in root
            if rel.get("TargetMode") == "External" and rel.get("Type", "").endswith("/hyperlink")}


def _text(el):
    return "".join(t.text or "" for t in el.iter() if _local(t.tag) == "t")


def _field_links(instr):
    """Targets of HYPERLINK field instructions; \\l switches jump to a bookmark in the same document."""
    return [m.group(2) for m in re.finditer(r'HYPERLINK\s+(\\l\s+)?"?([^"\s]+)', instr or "") if not m.group(1)]


SHAPES = {"sp", "pic", "graphicFrame", "cxnSp", "inline", "anchor", "wsp"}


def _xml_links(root, hlinks):
    """(target, where) for w:hyperlink, a:hlinkClick and HYPERLINK fields under root, in document order."""
    found = []  # [target, link text, label]; label, when set, describes a link that has no text

    def walk(el, ancestors):
        name = _local(el.tag)
        if name == "Fallback":  # mc:Fallback repeats the mc:Choice content
            return
        target = hlinks.get(el.get(f"{{{NS_R}}}id"))
        parent = ancestors[-1] if ancestors else None
        if target and name == "hyperlink":
            found.append([target, _text(el), None])
        elif target and name == "hlinkClick" and parent is not None and _local(parent.tag) == "rPr":
            text = _text(ancestors[-2])  # a:r/a:rPr/a:hlinkClick: a link on one text run
            if found and found[-1][0] == target and found[-1][2] is None:
                found[-1][1] += text  # one link split over several runs
            else:
                found.append([target, text, None])
        elif target and name == "hlinkClick":  # a link on a whole shape or picture
            shape = next((a for a in reversed(ancestors) if _local(a.tag) in SHAPES), None)
            text = _text(shape) if shape is not None else ""
            alt = (parent.get("descr") or parent.get("name")) if parent is not None else ""
            found.append([target, text, f'linked picture or shape "{alt}"' if alt and not text else None])
        elif name == "instrText":
            found.extend([t, "", "HYPERLINK field"] for t in _field_links(el.text))
        elif name == "fldSimple":
            found.extend([t, _text(el), None] for t in _field_links(el.get(f"{{{NS_W}}}instr")))
        for child in el:
            walk(child, ancestors + [el])

    walk(root, [])
    return [(target, label or link_where(text)) for target, text, label in found]


def _paragraphs(root):
    out = []
    for el in root.iter():
        name = _local(el.tag)
        if name == "p":
            text = "".join(t.text or "" for t in el.iter() if _local(t.tag) == "t")
            if text.strip():
                out.append(text)
        elif name == "v" and el.text and el.text.strip():  # chart category/series labels
            out.append(el.text)
    return out


def _para_texts(root):
    """Text of each w:p / a:p paragraph under root; a paragraph inside a text box is a paragraph of its own."""
    out = []

    def walk(el, buf):
        name = _local(el.tag)
        if name == "Fallback":
            return
        if name == "p":
            mine = []
            for child in el:
                walk(child, mine)
            if "".join(mine).strip():
                out.append("".join(mine))
            return
        if buf is not None and name == "t":
            buf.append(el.text or "")
        elif buf is not None and name in ("tab", "br", "cr"):
            buf.append(" ")
        for child in el:
            walk(child, buf)

    walk(root, None)
    return out


def _pptx_notes(z):
    """(slide number, speaker-notes paragraphs) for each slide that has notes."""
    pres = "ppt/presentation.xml"
    prels = _rels(z, pres)
    slides = [prels[el.get(f"{{{NS_R}}}id")][1] for el in ET.fromstring(z.read(pres)).iter() if _local(el.tag) == "sldId"]
    out = []
    for n, slide in enumerate(slides, 1):
        for kind, target in _rels(z, slide).values():
            if kind == "notesSlide":
                root = ET.fromstring(z.read(target))
                # skip the slide-image and slide-number placeholders
                texts = [t for sp in root.iter() if _local(sp.tag) == "sp"
                         and not any(_local(e.tag) == "ph" and e.get("type") in ("sldNum", "sldImg", "hdr", "ftr", "dt")
                                     for e in sp.iter())
                         for t in _para_texts(sp)]
                if texts:
                    out.append((n, texts))
    return out


def _linked_parts(elements, rels, page):
    items, seen = [], set()
    for el in elements:
        if _local(el.tag) == "svgBlip":  # Office always keeps a PNG fallback of SVG pictures
            continue
        for key, value in el.attrib.items():
            if key.startswith(f"{{{NS_R}}}") and value in rels and value not in seen:
                seen.add(value)
                kind, target = rels[value]
                if kind == "image":
                    items.append((page, "picture", "image", target))
                elif kind in ("chart", "diagramData"):
                    items.append((page, "chart" if kind == "chart" else "SmartArt", "xml", target))
    return items


def _part_links(z, part, place):
    """(where, target) of the hyperlinks in one XML part, where prefixed with place."""
    return [(f"{place}, {where}" if place else where, target)
            for target, where in _xml_links(ET.fromstring(z.read(part)), _hyperlink_rels(z, part))]


def _pptx_items(z):
    """(slide, where, kind, payload) for pictures, charts, SmartArt, grouped shapes and hyperlinks."""
    pres = "ppt/presentation.xml"
    prels = _rels(z, pres)
    slides = [prels[el.get(f"{{{NS_R}}}id")][1] for el in ET.fromstring(z.read(pres)).iter() if _local(el.tag) == "sldId"]
    items, shared, shared_links = [], {}, {}
    for n, slide in enumerate(slides, 1):
        srels = _rels(z, slide)
        root = ET.fromstring(z.read(slide))
        items += _linked_parts(root.iter(), srels, n)
        for tree in (el for el in root.iter() if _local(el.tag) == "spTree"):
            for shape in tree:
                if _local(shape.tag) == "grpSp":
                    if lines := _paragraphs(shape):
                        items.append((n, "grouped shapes", "lines", lines))
                elif texts := _para_texts(shape):
                    kinds = {_local(e.tag): e for e in shape.iter() if _local(e.tag) in ("ph", "tbl")}
                    ph = kinds.get("ph")
                    where = "slide title" if ph is not None and ph.get("type") in ("title", "ctrTitle") \
                        else "table" if "tbl" in kinds else ""
                    items.append((n, where, "body", texts))
        items += [(n, where, "link", target) for where, target in _part_links(z, slide, "")]
        for kind, layout in srels.values():
            if kind != "slideLayout":
                continue
            for link in _part_links(z, layout, "slide layout"):
                shared_links.setdefault(link, None)
            for kind2, target2 in _rels(z, layout).values():
                if kind2 == "image":
                    shared.setdefault(target2, "slide layout")
                elif kind2 == "slideMaster":
                    for link in _part_links(z, target2, "slide master"):
                        shared_links.setdefault(link, None)
                    for kind3, target3 in _rels(z, target2).values():
                        if kind3 == "image":
                            shared.setdefault(target3, "slide master")
    items += [(None, label, "image", target) for target, label in shared.items()]
    items += [(None, where, "link", target) for where, target in shared_links]
    items += [(n, "speaker notes", "body", texts) for n, texts in _pptx_notes(z)]
    return items


def _docx_items(z):
    """(estimated page, where, kind, payload); pages come from Word's last-rendered page breaks."""
    doc = "word/document.xml"
    rels = _rels(z, doc)
    hlinks = _hyperlink_rels(z, doc)
    data = z.read(doc)
    rendered = b"lastRenderedPageBreak" in data
    items, page, depth, fallback, in_link = [], 1, 0, 0, 0
    para, para_page, tables = [], 1, 0
    links = lambda el: [(page, where, "link", target) for target, where in _xml_links(el, hlinks)]
    for event, el in ET.iterparse(io.BytesIO(data), events=("start", "end")):
        name = _local(el.tag)
        if name == "Fallback":  # mc:Fallback repeats the mc:Choice drawing as VML
            fallback += 1 if event == "start" else -1
            continue
        if fallback:
            continue
        if depth == 0:  # body text outside drawings and text boxes
            if event == "start" and name == "p":
                para, para_page = [], page
            elif event == "start" and name == "tbl":
                tables += 1
            elif event == "end" and name == "tbl":
                tables -= 1
            elif event == "end" and name == "t":
                para.append(el.text or "")
            elif event == "end" and name in ("tab", "br", "cr"):
                para.append(" ")
            elif event == "end" and name == "p" and "".join(para).strip():
                items.append((para_page, "table" if tables else "", "body", ["".join(para)]))
        if event == "start":
            if name in ("drawing", "pict", "object"):
                depth += 1
            elif depth == 0 and name in ("hyperlink", "fldSimple"):
                in_link += 1
            elif depth == 0 and ((rendered and name == "lastRenderedPageBreak")
                                 or (not rendered and name == "br" and el.get(f"{{{NS_W}}}type") == "page")):
                page += 1
        elif depth == 0 and name in ("hyperlink", "fldSimple"):
            in_link -= 1
            if in_link == 0:
                items += links(el)
        elif depth == 0 and name == "instrText" and in_link == 0:
            items += links(el)
        elif name in ("drawing", "pict", "object"):
            depth -= 1
            if depth == 0:
                items += _linked_parts(el.iter(), rels, page)
                items += links(el)  # links on pictures and inside text boxes
                if {_local(e.tag) for e in el.iter()} & {"wgp", "wpc", "grpSp", "group"}:
                    if lines := _paragraphs(el):
                        items.append((page, "grouped shapes", "lines", lines))
                elif texts := _para_texts(el):
                    items.append((page, "text box", "body", texts))
    seen = set()
    for kind, part in rels.values():
        if kind in ("header", "footer"):
            items += [(None, f"page {kind}", "image", target)
                      for kind2, target in _rels(z, part).values() if kind2 == "image"]
        if kind in ("header", "footer", "footnotes", "endnotes"):
            place = f"page {kind}" if kind in ("header", "footer") else kind
            for where, target in _part_links(z, part, place):
                if (where, target) not in seen:  # first-page, even and default headers often repeat a link
                    seen.add((where, target))
                    items.append((None, where, "link", target))
            texts = [t for t in _para_texts(ET.fromstring(z.read(part))) if (place, t) not in seen]
            seen.update((place, t) for t in texts)
            if texts:
                items.append((None, place, "body", texts))
    return items


def pptx_notes_hits(path, ctx, pages_spec, count):
    """Body-text hits in speaker notes, which a PDF export of the slides leaves out."""
    try:
        with zipfile.ZipFile(path) as z:
            notes = _pptx_notes(z)
    except (zipfile.BadZipFile, KeyError, ET.ParseError) as e:
        ctx.warnings.append(f"speaker notes could not be read ({e}); not checked")
        return []
    keep = {p + 1 for p in parse_pages(pages_spec, count)}
    return [h for n, texts in notes if n in keep
            for h in body_hits(ctx, [(t, "speaker notes") for t in texts], n)]


def scan_office_native(path, ctx, info):
    ext = os.path.splitext(path)[1].lower()
    if ext not in {".docx", ".docm", ".dotx", ".pptx", ".pptm", ".ppsx", ".potx"}:
        die(f"{ext} files must be converted to PDF; install Microsoft Office or LibreOffice")
    info["method"] = "read directly (no PDF conversion)"
    slides = ext in SLIDE_EXTS
    hits = []
    with zipfile.ZipFile(path) as z:
        for page, where, kind, payload in (_pptx_items(z) if slides else _docx_items(z)):
            approx = not slides and page is not None
            if kind == "link":
                if ctx.links:
                    hits += link_hits(ctx, payload, page, where, approx=approx)
                continue
            if kind == "body":
                if ctx.text:
                    hits += body_hits(ctx, [(t, where) for t in payload], page, approx)
                continue
            if kind == "image":
                hits += scan_blob(z.read(payload), payload, ctx, f"{where} {posixpath.basename(payload)}", page, approx)
                continue
            lines = payload if kind == "lines" else _paragraphs(ET.fromstring(z.read(payload)))
            hits += [Hit(f.term, page, f.text.strip(), "text", "vector", where, page_approx=approx)
                     for f in search_lines(ctx.matcher, [Line(t, (0, 0, 0, 0)) for t in lines])]
    return hits


# ---------------------------------------------------------------- reporting

def mark_repeated(hits, min_pages=3):
    """Flag hits that recur on many pages: graphics at the same spot (header/footer logos), or the same link."""
    groups = {}
    for h in hits:
        if h.page is not None and h.source == "link":
            groups.setdefault((h.term, h.source, h.text), []).append(h)
        elif h.page is not None and h.bbox:
            x0, y0, x1, y1 = h.bbox
            key = (h.term, h.source, h.region, h.text.lower(), round(y0 / 10), round((x1 - x0) / 5), round((y1 - y0) / 5))
            groups.setdefault(key, []).append(h)
    for group in groups.values():
        if len({h.page for h in group}) >= min_pages:
            for h in group:
                h.repeated = True


def page_ranges(pages):
    out, pages = [], sorted(set(pages))
    i = 0
    while i < len(pages):
        j = i
        while j + 1 < len(pages) and pages[j + 1] == pages[j] + 1:
            j += 1
        out.append(str(pages[i]) if i == j else f"{pages[i]}-{pages[j]}")
        i = j + 1
    return ", ".join(out)


def _sort_key(h):
    if h.page is None or not h.bbox:  # keep discovery order (document order for HTML), after positioned hits
        return (h.page or 0, float("inf"), 0)
    return (h.page, h.bbox[1], h.bbox[0])


def coverage_notes(ctx):
    notes = list(ctx.warnings)
    if ctx.ocr is None:
        notes.append("OCR disabled: words inside pictures and outlined vector art were not checked")
    if ctx.skipped_images:
        notes.append(f"{ctx.skipped_images} raster image(s) skipped because OCR is disabled")
    return notes


def _excerpt(text, start, n=140):
    """text cut to about n characters, keeping position start in view."""
    if len(text) <= n:
        return text
    a = max(0, min(start - 30, len(text) - n))
    return ("..." if a else "") + text[a:a + n] + ("..." if a + n < len(text) else "")


def _print_hits(hits, unit, matcher):
    def shown(h):
        if h.source == "link":
            spans = link_spans(matcher, h.text)
        elif h.source == "body":
            spans = matcher.find(h.text)
        else:
            return h.text
        return _excerpt(h.text, min((s for t, s, _ in spans if t == h.term), default=0))

    rows = {}
    for h in sorted((h for h in hits if not h.repeated), key=_sort_key):
        rows.setdefault((h.page, h.source, h.region, h.text, h.where), []).append(h)
    last_crop = None
    for group in rows.values():
        h = group[0]
        where = f"{unit} {'~' if h.page_approx else ''}{h.page}" if h.page is not None else "-"
        extra = f"  [{h.where}]" if h.where else ""
        times = f" x{len(group)}" if len(group) > 1 else ""
        scores = [g.score for g in group if g.score is not None]
        conf = f"  (OCR confidence {min(scores):.2f})" if h.source == "ocr" and scores else ""
        print(f'  {where:<10} {LABELS[h.source, h.region]:<17} "{shown(h)}"{times}{extra}{conf}')
        if h.crop and h.crop != last_crop:
            print(f"  {'':<10} crop: {h.crop}")
            last_crop = h.crop
    groups = {}
    for h in (h for h in hits if h.repeated):
        groups.setdefault((h.source, h.region, h.text.lower()), []).append(h)
    for (source, region, _), group in groups.items():
        group.sort(key=_sort_key)
        gpages = sorted({h.page for h in group})
        where = group[0].where or "same position on each page"
        print(f'  repeated on {len(gpages)} {unit}s ({page_ranges(gpages)}): {LABELS[source, region]} '
              f'"{shown(group[0])}"  [{where}]')
        if group[0].crop:
            print(f"  {'':<10} crop: {group[0].crop}")


def search_scope(text, links):
    kinds = ["diagram", "image"] + (["text"] if text else []) + (["hyperlink"] if links else [])
    return ", ".join(kinds[:-1]) + " or " + kinds[-1]


def print_report(info, terms, hits, notes, unit, matcher, links=True, text=True):
    details = [info["type"].upper()]
    if "size" in info:
        details.append(info["size"])
    if "pages" in info:
        details.append(f"{info['pages']} {unit}s")
    if info.get("converted_with"):
        details.append(f"converted to PDF with {info['converted_with']}")
    if info.get("method"):
        details.append(info["method"])
    if "images" in info:
        details.append(f"{info['images']} images, {info['inline_svgs']} inline SVGs")
    if "links" in info:
        details.append(f"{info['links']} links")
    print(f"File: {info['file']} ({', '.join(details)})")
    scope = search_scope(text, links)
    for term in terms:
        th = [h for h in hits if h.term == term]
        if not th:
            print(f'\n"{term}": NOT FOUND in any {scope}')
            continue
        pages = sorted({h.page for h in th if h.page is not None})
        in_text = [h for h in th if h.source == "body"]
        in_links = [h for h in th if h.source == "link"]
        split = [f"{len(th) - len(in_text) - len(in_links)} in graphics"] if in_text or in_links else []
        split += [f"{len(in_text)} in text"] if in_text else []
        split += [f"{len(in_links)} in hyperlinks"] if in_links else []
        print(f'\n"{term}": FOUND - {len(th)} match(es)' + (f" ({', '.join(split)})" if split else "")
              + (f" on {len(pages)} {unit}(s): {page_ranges(pages)}" if pages else ""))
        _print_hits([h for h in th if h.source not in ("link", "body")], unit, matcher)
        if in_text:
            print("  in body text:")
            _print_hits(in_text, unit, matcher)
        if in_links:
            print("  in hyperlink targets:")
            _print_hits(in_links, unit, matcher)
    for note in notes:
        print(f"note: {note}")
    if info.get("report"):
        print(f"\nHTML report: {info['report']}")


def main(argv=None):
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(encoding="utf-8", errors="replace")
        except AttributeError:
            pass
    ap = argparse.ArgumentParser(description="Report pages where diagrams, images, body text or hyperlink targets "
                                             "contain a word or phrase.")
    ap.add_argument("file", help="PDF, DOCX/DOC, PPTX/PPT, HTML, SVG or image file (PNG, JPG, TIFF, ...)")
    ap.add_argument("terms", nargs="+", metavar="TERM", help='word or phrase, e.g. "Intel" (several allowed)')
    ap.add_argument("--json", action="store_true", help="print machine-readable JSON")
    ap.add_argument("--case-sensitive", action="store_true")
    ap.add_argument("--substring", action="store_true", help="also match inside longer words (Intel in Intelligent)")
    ap.add_argument("--pages", help="pages/slides to scan, e.g. 1-5,9")
    ap.add_argument("--no-ocr", action="store_true", help="text layer only: fast, but misses words in pictures")
    ap.add_argument("--ocr-all", action="store_true",
                    help="also OCR vector diagrams that look fully selectable (slower, catches odd outlined text)")
    ap.add_argument("--min-score", type=float, default=0.5, help="minimum OCR confidence (default 0.5)")
    ap.add_argument("--zoom", type=float, default=3.0, help="render scale for vector graphics before OCR (default 3 = 216 dpi)")
    ap.add_argument("--native", action="store_true", help="read DOCX/PPTX directly instead of converting to PDF")
    ap.add_argument("--no-links", action="store_true", help="do not search hyperlink targets (URLs, e-mail, files)")
    ap.add_argument("--no-text", action="store_true", help="do not search body text (paragraphs, tables, headers)")
    ap.add_argument("--crops", metavar="DIR", help="save an annotated PNG of every graphic with a match")
    ap.add_argument("--html", nargs="?", const="", metavar="FILE",
                    help="write a self-contained HTML report with marked screenshots "
                         "(default FILE: <input name>_report.html next to the input)")
    ap.add_argument("-o", "--output", metavar="FILE", help="write the report to FILE (UTF-8) instead of stdout")
    args = ap.parse_args(argv)
    if args.output:
        sys.stdout = open(args.output, "w", encoding="utf-8")

    path = os.path.abspath(args.file)
    if not os.path.isfile(path):
        die(f"file not found: {path}")
    terms = list(dict.fromkeys(t.strip() for t in args.terms if t.strip()))
    if not terms:
        die("the search term is empty")
    ext = os.path.splitext(path)[1].lower()
    report = None
    if args.html is not None:
        report = os.path.abspath(args.html or os.path.splitext(path)[0] + "_report.html")
    ctx = Context(Matcher(terms, args.case_sensitive, args.substring),
                  None if args.no_ocr else Ocr(args.min_score), args.zoom, args.crops, args.ocr_all, report is not None,
                  links=not args.no_links, text=not args.no_text)
    info = {"file": path, "type": ext.lstrip(".")}
    unit = "slide" if ext in SLIDE_EXTS else "page"
    can_link = can_text = True  # image files have no body text; raster images cannot carry hyperlinks
    started = time.time()
    try:
        if ext == ".pdf":
            hits = scan_pdf(path, ctx, args.pages, info)
        elif ext in WORD_EXTS | SLIDE_EXTS:
            if args.native:
                hits = scan_office_native(path, ctx, info)
            else:
                with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as tmp:
                    try:
                        pdf, tool = convert_to_pdf(path, tmp, ext in SLIDE_EXTS)
                    except RuntimeError as e:
                        ctx.warnings.append(f"PDF conversion failed ({e}); file was read directly"
                                            + ("" if unit == "slide" else ", so page numbers are estimates"))
                        hits = scan_office_native(path, ctx, info)
                    else:
                        info["converted_with"] = tool
                        hits = scan_pdf(pdf, ctx, args.pages, info)
                        if ctx.text and ext in {".pptx", ".pptm", ".ppsx", ".potx"}:
                            hits += pptx_notes_hits(path, ctx, args.pages, info["pages"])
        elif ext in HTML_EXTS:
            hits = scan_html(path, ctx, info)
        elif ext in SVG_EXTS or ext in IMAGE_EXTS or (ext not in WORD_EXTS | SLIDE_EXTS and looks_like_image(path)):
            hits = scan_image_file(path, ctx, args.pages, info)
            can_link, can_text = info["type"] in ("svg", "svgz"), False
        else:
            die(f"unsupported file type: {ext or '(none)'}")
    except SystemExit:
        raise
    except Exception:
        traceback.print_exc()
        sys.exit(2)

    mark_repeated(hits)
    hits.sort(key=lambda h: (terms.index(h.term),) + _sort_key(h))
    notes = coverage_notes(ctx)
    links_checked = ctx.links_checked if ctx.links and can_link else None
    text_checked = ctx.text_checked if ctx.text and can_text else None
    if report:
        from html_report import write_html_report
        options = {"OCR": "off" if ctx.ocr is None else f"on (min confidence {args.min_score:g})",
                   "Matching": ("case-sensitive" if args.case_sensitive else "case-insensitive")
                               + (", substrings" if args.substring else ", whole words")}
        if args.pages:
            options[f"{unit.capitalize()}s"] = args.pages
        if not ctx.text and can_text:
            options["Body text"] = "not checked"
        if not ctx.links and can_link:
            options["Hyperlinks"] = "not checked"
        stats = {"elapsed": time.time() - started, "checked": ctx.checked,
                 "ocr_runs": len(ctx.ocr.cache) if ctx.ocr else 0, "links": links_checked, "text": text_checked}

        def highlight(text, term, source):
            spans = link_spans(ctx.matcher, text) if source == "link" else ctx.matcher.find(text, source == "ocr")
            return [(s, e) for t, s, e in spans if t == term]

        write_html_report(report, info=info, unit=unit, terms=terms, hits=hits, notes=notes, options=options,
                          stats=stats, colors={t: ctx.color(t) for t in terms}, graphics=ctx.graphics,
                          hit_graphic=ctx.hit_graphic, labels=LABELS, page_ranges=page_ranges, highlight=highlight)
        info["report"] = report
    info["links_checked"] = links_checked
    info["text_checked"] = text_checked
    if args.json:
        summary = {}
        for term in terms:
            th = [h for h in hits if h.term == term]
            summary[term] = {"found": bool(th), "matches": len(th),
                             "in_text": sum(h.source == "body" for h in th),
                             "in_links": sum(h.source == "link" for h in th),
                             "pages": sorted({h.page for h in th if h.page is not None})}
        print(json.dumps({**info, "unit": unit, "terms": terms, "found": bool(hits), "summary": summary,
                          "matches": [asdict(h) for h in hits], "warnings": notes},
                         indent=2, ensure_ascii=False))
    else:
        print_report(info, terms, hits, notes, unit, ctx.matcher, links_checked is not None, text_checked is not None)
    sys.exit(0 if hits else 1)


if __name__ == "__main__":
    main()
