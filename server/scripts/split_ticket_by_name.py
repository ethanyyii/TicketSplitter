#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Scoot itinerary splitter (dynamic: direct + connecting)

Per passenger:
  - Interleaved: outbound itinerary (through FLIGHT 2) + first half of add-ons,
    then return itinerary + second half of add-ons (when FLIGHT 2 exists and ≥2 add-on blocks).
  - Otherwise: full header clips + all add-on row clips

Add-ons region ends before "Fare Terms and Conditions" or (if absent) before the
last "Click here to Purchase more add-ons" — no legal/footer pages are emitted.

Output pages: width = source doc page width (w0), height = A4 height; clips pasted
at 1:1 (no scaling). Natural page breaks when content exceeds printable height.

Filename: <ORIGINAL_STEM>_<PASSENGER_NAME>.pdf
"""

import re
import argparse
from pathlib import Path
import math
from collections import defaultdict

import fitz  # PyMuPDF

TITLE_RE = re.compile(r"^(MR|MS|MRS|MISS|MSTR)\b", re.I)
CLICK_NEEDLE = "Click here to Purchase more add-ons"
FARE_TERMS_NEEDLE = "Fare Terms and Conditions"
PASSENGER_ADDONS_NEEDLES = (
    "Passenger & Add-ons on this flight",
    "Passenger & Add-ons on this",
)

A4_HEIGHT = float(fitz.paper_size("A4")[1])
MARGIN_TOP = 20.0
MARGIN_BOTTOM = 30.0
BLOCK_GAP = 10.0


def normalize_spaces(s: str) -> str:
    return re.sub(r"\s+", " ", (s or "")).strip()


def passenger_name_only(full: str) -> str:
    s = normalize_spaces(full)
    s = re.sub(r"^(MR|MS|MRS|MISS|MSTR)\s+", "", s, flags=re.I).strip()
    s = re.sub(r"\b(Check[- ]?in)\b", "", s, flags=re.I).strip()
    return normalize_spaces(s)


def safe_filename_component(s: str) -> str:
    s = normalize_spaces(s).upper()
    s = re.sub(r"[^A-Z0-9]+", "_", s).strip("_")
    return s[:120] if len(s) > 120 else s


def _finite_rect(r: fitz.Rect) -> bool:
    return all(math.isfinite(v) for v in (r.x0, r.y0, r.x1, r.y1))


def _ensure_valid_clip(r: fitz.Rect, page_rect: fitz.Rect, min_h: float = 5.0) -> fitz.Rect:
    rr = fitz.Rect(
        max(page_rect.x0, r.x0),
        max(page_rect.y0, r.y0),
        min(page_rect.x1, r.x1),
        min(page_rect.y1, r.y1),
    )
    if (not _finite_rect(rr)) or rr.is_empty or rr.width < 5 or rr.height < min_h:
        raise ValueError(f"Invalid clip rect: {rr}")
    area = rr.width * rr.height
    if area <= 0:
        raise ValueError(f"Invalid clip rect(area<=0): {rr}")
    return rr


def trim_clip_bottom(page: fitz.Page, rect: fitz.Rect, padding: float = 30.0) -> fitz.Rect:
    """
    Scan text blocks inside rect and shrink rect.y1 to the lowest content bottom + padding,
    removing excess whitespace below the itinerary text.
    """
    pr = page.rect
    max_y1 = float(rect.y0)
    has_content = False
    for b in page.get_text("blocks") or []:
        if len(b) < 5:
            continue
        block_type = int(b[6]) if len(b) > 6 else 0
        if block_type != 0:
            continue
        b_y0, b_y1 = float(b[1]), float(b[3])
        if b_y0 >= rect.y0 - 5 and b_y0 < rect.y1:
            if b_y1 > max_y1:
                max_y1 = b_y1
            has_content = True

    if has_content:
        new_y1 = min(float(rect.y1), float(pr.y1), max_y1 + padding)
        if new_y1 <= rect.y0 + 1e-3:
            return rect
        return fitz.Rect(rect.x0, rect.y0, rect.x1, new_y1)
    return rect


def page_lines(page: fitz.Page):
    d = page.get_text("dict")
    lines = []
    for b in d.get("blocks", []):
        if b.get("type") != 0:
            continue
        for l in b.get("lines", []):
            txt = "".join(span.get("text", "") for span in l.get("spans", [])).strip()
            if txt:
                lines.append((fitz.Rect(l["bbox"]), txt))
    lines.sort(key=lambda x: (x[0].y0, x[0].x0))
    return lines


def find_line_bbox(page: fitz.Page, contains: str, y_gt=None, y_lt=None):
    for bbox, txt in page_lines(page):
        if y_gt is not None and bbox.y0 <= y_gt:
            continue
        if y_lt is not None and bbox.y0 >= y_lt:
            continue
        if contains in txt:
            return bbox, txt
    return None


def find_first_line_in_doc(doc: fitz.Document, needle: str):
    for pno in range(len(doc)):
        hit = find_line_bbox(doc[pno], needle, y_gt=None)
        if hit:
            return pno, hit[0], hit[1]
    return None


def find_passenger_addons_heading(doc: fitz.Document):
    for needle in PASSENGER_ADDONS_NEEDLES:
        r = find_first_line_in_doc(doc, needle)
        if r:
            return r[0], r[1], needle
    return None


def find_first_fare_terms(doc: fitz.Document):
    return find_first_line_in_doc(doc, FARE_TERMS_NEEDLE)


def find_last_click_in_doc(doc: fitz.Document):
    best = None
    best_tup = None
    for pno in range(len(doc)):
        for bbox, txt in page_lines(doc[pno]):
            if CLICK_NEEDLE in txt:
                t = (pno, bbox.y0)
                if best_tup is None or t > best_tup:
                    best_tup = t
                    best = (pno, bbox)
    return best


def resolve_addons_region_end(doc: fitz.Document):
    """
    Exclusive end of add-ons vertical range: first Fare Terms y0, or else last Click y0.
    If neither exists, end of document.
    """
    ft = find_first_fare_terms(doc)
    if ft:
        p_terms, bbox_terms, _txt = ft
        return p_terms, bbox_terms.y0, "fare_terms"
    lc = find_last_click_in_doc(doc)
    if lc:
        p_click, bbox_click = lc
        return p_click, bbox_click.y0, "click"
    last_pno = len(doc) - 1
    return last_pno, doc[last_pno].rect.height, "eof"


def extract_name_occurrences_page(page: fitz.Page):
    lines = page_lines(page)
    occ = []
    i = 0
    while i < len(lines):
        bbox, txt = lines[i]
        t = txt.strip()
        if TITLE_RE.match(t):
            parts = [t]
            bbox_comb = fitz.Rect(bbox)
            j = i + 1
            while j < len(lines):
                bb2, nxt = lines[j]
                n = nxt.strip()
                if TITLE_RE.match(n):
                    break
                up = n.upper()
                if re.match(r"^[A-Z][A-Z \-']*$", up) and len(up) <= 40:
                    parts += n.split()
                    bbox_comb = bbox_comb | bb2
                    j += 1
                    continue
                break
            full = " ".join(parts).strip()
            if len(full.split()) >= 2:
                occ.append((full, bbox_comb))
            i = j
        else:
            i += 1
    occ.sort(key=lambda x: x[1].y0)
    return occ


def text_spans(page: fitz.Page):
    d = page.get_text("dict")
    spans = []
    for b in d.get("blocks", []):
        if b.get("type") != 0:
            continue
        for l in b.get("lines", []):
            for s in l.get("spans", []):
                t = (s.get("text", "") or "").strip()
                if t:
                    spans.append((fitz.Rect(s["bbox"]), t))
    spans.sort(key=lambda x: (x[0].y0, x[0].x0))
    return spans


def find_baggage_bottom(page: fitz.Page, y0: float, y1: float):
    bottoms = []
    for r, t in text_spans(page):
        if r.y0 >= y0 - 2 and r.y1 <= y1 + 2 and re.fullmatch(r"\d+KG", t.strip().upper()):
            bottoms.append(r.y1)
    return max(bottoms) if bottoms else None


def row_clip_single_passenger(
    page: fitz.Page,
    name_bb: fitz.Rect,
    region_candidates,
    region_top_y0: float,
    region_bottom_y1: float,
):
    page_rect = page.rect
    prev_y1s = [bb.y1 for _nm, bb in region_candidates if bb.y0 < name_bb.y0]
    prev_safe_top = max(prev_y1s) + 2 if prev_y1s else region_top_y0
    y0 = max(region_top_y0, prev_safe_top, name_bb.y0 - 4)
    next_y0s = [bb.y0 for _nm, bb in region_candidates if bb.y0 > name_bb.y0]
    hard_cap = region_bottom_y1
    if next_y0s:
        hard_cap = min(hard_cap, min(next_y0s) - 2)
    search_bottom = min(hard_cap, y0 + 900)
    b_bot = find_baggage_bottom(page, y0 + 10, search_bottom)
    if b_bot is not None:
        y1 = min(hard_cap, b_bot + 18)
    else:
        y1 = min(hard_cap, name_bb.y1 + 220)
    if y1 <= y0 + 20:
        y1 = min(hard_cap, y0 + 240)
    clip = fitz.Rect(0, y0, page_rect.width, y1)
    return _ensure_valid_clip(clip, page_rect, min_h=20.0)


def split_global_itinerary(doc: fitz.Document, header_end_pno: int, header_end_y0: float):
    """
    Split itinerary header at the first "FLIGHT 2" line (outbound vs return).
    part2 is empty if not found (e.g. one-way).
    """
    f2_pno = -1
    f2_y0 = -1.0
    for p in range(header_end_pno + 1):
        y_lt = header_end_y0 if p == header_end_pno else None
        found = find_line_bbox(doc[p], "FLIGHT 2", y_lt=y_lt)
        if found:
            f2_pno = p
            prp = doc[p].rect
            f2_y0 = max(float(prp.y0), float(found[0].y0) - 10.0)
            break

    part1 = []
    part2 = []

    if f2_pno < 0:
        for p in range(header_end_pno):
            pr = doc[p].rect
            tc = trim_clip_bottom(doc[p], pr)
            part1.append((p, _ensure_valid_clip(tc, pr, min_h=1.0)))
        prh = doc[header_end_pno].rect
        head_clip = fitz.Rect(0, 0, prh.width, header_end_y0)
        tc_h = trim_clip_bottom(doc[header_end_pno], head_clip)
        part1.append((header_end_pno, _ensure_valid_clip(tc_h, prh, min_h=1.0)))
        return part1, part2

    for p in range(f2_pno):
        pr = doc[p].rect
        tc = trim_clip_bottom(doc[p], pr)
        part1.append((p, _ensure_valid_clip(tc, pr, min_h=1.0)))
    pr_f2 = doc[f2_pno].rect
    clip1 = fitz.Rect(0, 0, pr_f2.width, f2_y0)
    tc1 = trim_clip_bottom(doc[f2_pno], clip1)
    part1.append((f2_pno, _ensure_valid_clip(tc1, pr_f2, min_h=1.0)))

    if f2_pno == header_end_pno:
        clip2 = fitz.Rect(0, f2_y0, pr_f2.width, header_end_y0)
        tc2 = trim_clip_bottom(doc[f2_pno], clip2)
        part2.append((f2_pno, _ensure_valid_clip(tc2, pr_f2, min_h=1.0)))
    else:
        clip2 = fitz.Rect(0, f2_y0, pr_f2.width, pr_f2.height)
        tc2b = trim_clip_bottom(doc[f2_pno], clip2)
        part2.append((f2_pno, _ensure_valid_clip(tc2b, pr_f2, min_h=1.0)))
        for p in range(f2_pno + 1, header_end_pno):
            pr = doc[p].rect
            tc = trim_clip_bottom(doc[p], pr)
            part2.append((p, _ensure_valid_clip(tc, pr, min_h=1.0)))
        pr_end = doc[header_end_pno].rect
        clip_end = fitz.Rect(0, 0, pr_end.width, header_end_y0)
        tc_e = trim_clip_bottom(doc[header_end_pno], clip_end)
        part2.append((header_end_pno, _ensure_valid_clip(tc_e, pr_end, min_h=1.0)))

    return part1, part2


def addon_vertical_slices(doc: fitz.Document, p0: int, y0_top: float, p1: int, y1_end: float):
    if p0 > p1 or (p0 == p1 and y1_end <= y0_top):
        return []
    out = []
    if p0 == p1:
        out.append((p0, y0_top, y1_end))
        return out
    h0 = doc[p0].rect.height
    out.append((p0, y0_top, h0))
    for p in range(p0 + 1, p1):
        hp = doc[p].rect.height
        out.append((p, 0.0, hp))
    out.append((p1, 0.0, y1_end))
    return out


def names_in_vertical_band(page: fitz.Page, y_top: float, y_bottom: float, tol: float = 2.0):
    occ = extract_name_occurrences_page(page)
    band = []
    for nm, bb in occ:
        if bb.y0 >= y_top - tol and bb.y0 < y_bottom - tol:
            band.append((nm, bb))
    band.sort(key=lambda x: x[1].y0)
    return band


def collect_addon_occurrences_by_key(doc: fitz.Document, slices):
    by_key = defaultdict(list)
    for pno, yt, yb in slices:
        for nm, bb in names_in_vertical_band(doc[pno], yt, yb):
            k = passenger_name_only(nm).upper()
            if k:
                by_key[k].append((pno, bb))
    for k in by_key:
        by_key[k].sort(key=lambda x: (x[0], x[1].y0))
    return by_key


def all_passenger_keys(doc: fitz.Document):
    keys = []
    seen = set()
    for pno in range(len(doc)):
        for nm, _bb in extract_name_occurrences_page(doc[pno]):
            k = passenger_name_only(nm).upper()
            if k and k not in seen:
                seen.add(k)
                keys.append(k)
    return keys


def build_addon_clips_for_passenger(doc: fitz.Document, key: str, occ_by_key: dict, slices):
    occs = occ_by_key.get(key, [])
    clips = []
    for pno, bb in occs:
        yt_yb = None
        for p, yt, yb in slices:
            if p == pno:
                yt_yb = (yt, yb)
                break
        if yt_yb is None:
            continue
        yt, yb = yt_yb
        page = doc[pno]
        cands = names_in_vertical_band(page, yt, yb)
        clip = row_clip_single_passenger(page, bb, cands, yt, yb)
        clips.append((pno, clip))
    return clips


def paginate_clips_one_to_one(
    out_doc: fitz.Document,
    src_doc: fitz.Document,
    parts: list,
    w0: float,
):
    """
    Paste each clip at 1:1: target height = clip height, target width = w0.
    New pages use (w0, A4_HEIGHT). Natural breaks when content exceeds bottom margin.
    """
    y_limit = A4_HEIGHT - MARGIN_BOTTOM
    out_page = None
    current_y = MARGIN_TOP

    def new_sheet():
        nonlocal out_page, current_y
        out_page = out_doc.new_page(width=w0, height=A4_HEIGHT)
        current_y = MARGIN_TOP

    for pno, clip in parts:
        if clip.width <= 0 or clip.height <= 0:
            continue

        h = clip.height
        if out_page is None:
            new_sheet()

        if current_y > MARGIN_TOP and current_y + h > y_limit:
            new_sheet()

        target = fitz.Rect(0, current_y, w0, current_y + h)
        out_page.show_pdf_page(target, src_doc, pno, clip=clip)
        current_y += h + BLOCK_GAP

    return len(out_doc)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("pdf", type=Path, help="Input PDF")
    ap.add_argument("-o", "--out", type=Path, default=None,
                    help="Output directory (default: <stem>_split)")
    args = ap.parse_args()

    pdf_path: Path = args.pdf
    out_dir: Path = args.out if args.out else pdf_path.with_suffix("").with_name(pdf_path.stem + "_split")
    out_dir.mkdir(parents=True, exist_ok=True)

    doc = fitz.open(pdf_path)
    try:
        head = find_passenger_addons_heading(doc)
        if not head:
            raise RuntimeError("Cannot find Passenger & Add-ons heading in PDF.")
        header_end_pno, header_end_bbox, head_needle = head
        header_end_y0 = header_end_bbox.y0

        print(f"Found header ending at Page {header_end_pno + 1} (0-based pno={header_end_pno}), "
              f"y0={header_end_y0:.2f} (matched: {head_needle!r})")

        end_pno, end_y, end_kind = resolve_addons_region_end(doc)
        print(f"Add-ons region ends before: {end_kind!r} at page {end_pno + 1}, y0={end_y:.2f} (exclusive).")

        slices = addon_vertical_slices(doc, header_end_pno, header_end_y0, end_pno, end_y)
        if not slices:
            print("Warning: empty add-ons vertical range.")

        header_part1, header_part2 = split_global_itinerary(doc, header_end_pno, header_end_y0)
        print(f"Global itinerary: part1 {len(header_part1)} clip(s), part2 {len(header_part2)} clip(s); "
              f"no footer/terms in output.")

        occ_by_key = collect_addon_occurrences_by_key(doc, slices)
        all_keys = all_passenger_keys(doc)
        for k in all_keys:
            if k not in occ_by_key:
                occ_by_key[k] = []

        w0 = float(doc[0].rect.width)

        stem = pdf_path.stem
        exported = 0

        for key in sorted(set(all_keys) | set(occ_by_key.keys())):
            addon_clips = build_addon_clips_for_passenger(doc, key, occ_by_key, slices)
            mid = len(addon_clips) // 2
            if header_part2 and len(addon_clips) >= 2:
                parts = list(header_part1) + list(addon_clips[:mid]) + list(header_part2) + list(addon_clips[mid:])
            else:
                parts = list(header_part1) + list(header_part2) + list(addon_clips)

            print(f"Passenger {key!r}: {len(addon_clips)} add-on block(s).")

            out = fitz.open()
            n_pages = paginate_clips_one_to_one(out, doc, parts, w0)
            print(f"  -> Output: {n_pages} page(s), width={w0:.2f} pt, height={A4_HEIGHT:.2f} (A4), "
                  f"1:1 paste, margin_top={MARGIN_TOP}, margin_bottom={MARGIN_BOTTOM}, gap={BLOCK_GAP}.")

            passenger_component = safe_filename_component(key)
            out_name = f"{stem}_{passenger_component}.pdf"
            out.save(out_dir / out_name)
            out.close()
            exported += 1

        print(f"Done. Exported {exported} file(s) to: {out_dir}")
    finally:
        doc.close()


if __name__ == "__main__":
    main()
