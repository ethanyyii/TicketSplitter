#!/usr/bin/env python3
"""
泰國獅子航空（Thai Lion Air）多人機票 PDF 分割工具。

需要安裝：pip install pymupdf

用法：
  python split_thailionair.py INPUT.pdf
  python split_thailionair.py INPUT.pdf -o /path/to/output

若未指定 -o/--out，輸出目錄預設為 <原檔名>_split/

多航段：PREPARED FOR 頁 + 其後第一個旅客明細頁為一 Segment；同一旅客以正規化姓名合併。
各航段依「出發日期」排序（早的在上），並繪製於同一張 A4 單頁（垂直堆疊，必要時縮放）。
"""

from __future__ import annotations

import argparse
import logging
import os
import re
import sys
from collections import defaultdict
from dataclasses import dataclass
from datetime import date
from typing import DefaultDict, List, Optional, Sequence, Tuple

import fitz

# ── 版面常數 ─────────────────────────────────────────────────────────────────

HEADER_CLIP_Y1 = 90.0  # 標頭裁切高度（與原邏輯一致）
A4_HEIGHT = 841.89
A4_WIDTH = 595.276
SEGMENT_GAP = 24.0  # 航段之間間距（像素）
PAGE_TOP_MARGIN = 8.0
PAGE_BOTTOM_MARGIN = 10.0

# ── 日誌 ──────────────────────────────────────────────────────────────────────

LOG = logging.getLogger("thailionair.split")


def _setup_logging() -> None:
    logging.basicConfig(
        level=logging.INFO,
        format="%(message)s",
        stream=sys.stdout,
    )


# ── 資料結構 ──────────────────────────────────────────────────────────────────

PassengerRow = Tuple[str, str]  # (display_name, e_ticket_number)


@dataclass(frozen=True)
class Segment:
    """單一航段：標頭頁 + 明細頁，以及用於排序的出發日期。"""

    header_page: int
    detail_page: int
    passengers: Tuple[PassengerRow, ...]
    sort_date: date


@dataclass
class PaxSegmentRecord:
    segment_index: int
    display_name: str
    ticket: str
    detail_page: int


@dataclass
class PreparedSegment:
    """已套用遮罩／紅塊／單一旅客列的暫存文件（僅供 show_pdf_page 使用）。"""

    comp: fitz.Document
    header_page: int
    detail_page: int
    pw: float
    header_h: float
    dep_y0: float
    content_bottom: float

    @property
    def detail_h(self) -> float:
        return max(0.0, self.content_bottom - self.dep_y0)


# ── 日期解析（航段排序）──────────────────────────────────────────────────────

_MONTH_MAP = {
    "JAN": 1,
    "FEB": 2,
    "MAR": 3,
    "APR": 4,
    "MAY": 5,
    "JUN": 6,
    "JUL": 7,
    "AUG": 8,
    "SEP": 9,
    "OCT": 10,
    "NOV": 11,
    "DEC": 12,
}

# 10 APR 2026、10 Apr. 2026、01APR2026 等
_RE_DATE_DMY = re.compile(
    r"\b(\d{1,2})\s*[-/]?\s*(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*\.?\s*[-/]?\s*(\d{2,4})\b",
    re.IGNORECASE,
)
_RE_DATE_COMPACT = re.compile(
    r"\b(\d{1,2})(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)(\d{4})\b",
    re.IGNORECASE,
)


def _month_num(m: str) -> int:
    key = m.upper()[:3]
    if key not in _MONTH_MAP:
        raise ValueError(m)
    return _MONTH_MAP[key]


def _normalize_year(y: int) -> int:
    if y < 100:
        return 2000 + y if y < 50 else 1900 + y
    return y


def _parse_first_date_in_text(text: str) -> Optional[date]:
    for rx in (_RE_DATE_DMY, _RE_DATE_COMPACT):
        m = rx.search(text)
        if not m:
            continue
        try:
            d = int(m.group(1))
            mon = _month_num(m.group(2))
            y = _normalize_year(int(m.group(3)))
            return date(y, mon, d)
        except (ValueError, OverflowError):
            continue
    return None


def parse_segment_sort_date(doc: fitz.Document, header_page: int, detail_page: int) -> date:
    """
    優先從明細頁「DEPARTURE」附近的文字取第一個日期，其次整頁明細、再標頭頁。
    若皆無則回傳極晚日期，排序時落在後面。
    """
    fallback = date(9999, 12, 31)
    dp = doc[detail_page]
    full = dp.get_text()
    up = full.upper()
    dep_i = up.find("DEPARTURE")
    if dep_i >= 0:
        window = full[dep_i : dep_i + 800]
        d = _parse_first_date_in_text(window)
        if d:
            return d
    d = _parse_first_date_in_text(full)
    if d:
        return d
    hp = doc[header_page]
    d = _parse_first_date_in_text(hp.get_text())
    if d:
        return d
    return fallback


# ── 文字／版面工具 ────────────────────────────────────────────────────────────


def find_line_bbox(page: fitz.Page, keyword: str):
    for b in page.get_text("dict")["blocks"]:
        for line in b.get("lines", []):
            txt = "".join(s["text"] for s in line["spans"])
            if keyword in txt:
                x0 = min(s["bbox"][0] for s in line["spans"])
                y0 = min(s["bbox"][1] for s in line["spans"])
                x1 = max(s["bbox"][2] for s in line["spans"])
                y1 = max(s["bbox"][3] for s in line["spans"])
                size = line["spans"][0]["size"]
                return x0, y0, x1, y1, size
    return None, None, None, None, 8


def get_names_x1(page: fitz.Page, y_top: float, y_bottom: float) -> Optional[float]:
    max_x1 = 0.0
    for b in page.get_text("dict")["blocks"]:
        for line in b.get("lines", []):
            for span in line["spans"]:
                sx0, sy0, sx1, sy1 = span["bbox"]
                if sy0 >= y_top and sy1 <= y_bottom:
                    if sx1 > max_x1:
                        max_x1 = sx1
    return max_x1 if max_x1 > 0 else None


def extract_passengers(page: fitz.Page) -> List[PassengerRow]:
    all_spans = []
    for b in page.get_text("dict")["blocks"]:
        for line in b.get("lines", []):
            for s in line["spans"]:
                all_spans.append(
                    {
                        "text": s["text"].strip(),
                        "y0": s["bbox"][1],
                        "y1": s["bbox"][3],
                        "x0": s["bbox"][0],
                    }
                )

    ticket_spans = [s for s in all_spans if s["text"].isdigit() and len(s["text"]) >= 10]
    passengers: List[PassengerRow] = []
    skip_words = {
        "Check-In",
        "Required",
        "Check-In Required",
        "Passenger",
        "Name:",
        "Seats:",
        "eTicket",
        "Receipt(s):",
    }

    for ts in ticket_spans:
        row_texts = []
        for s in all_spans:
            if s["text"] in skip_words or not s["text"]:
                continue
            if s["text"].isdigit():
                continue
            if s["y0"] < ts["y1"] and s["y1"] > ts["y0"]:
                row_texts.append((s["x0"], s["text"]))
        row_texts.sort()
        name = " ".join(t for _, t in row_texts)
        name = name.replace("»", "").replace("\u00bb", "").strip()
        if name:
            passengers.append((name, ts["text"]))

    return passengers


def reservation_code_from_header(page: fitz.Page) -> str:
    res_code = "UNKNOWN"
    for b in page.get_text("dict")["blocks"]:
        for line in b.get("lines", []):
            txt = "".join(s["text"] for s in line["spans"])
            if "RESERVATION CODE" in txt:
                tokens = txt.strip().split()
                if tokens:
                    res_code = tokens[-1]
                return res_code
    return res_code


# ── 姓名正規化 ────────────────────────────────────────────────────────────────

_TITLE_PREFIXES = frozenset({"MR", "MS", "MRS", "MISS"})


def normalize_passenger_stem(name: str) -> str:
    parts = name.split()
    if parts:
        p0 = parts[0].upper().rstrip(".")
        if p0 in _TITLE_PREFIXES:
            parts = parts[1:]
    raw = " ".join(parts).upper()
    raw = raw.replace("»", "").replace("\u00bb", "")
    stem = re.sub(r"[^A-Z0-9]+", "_", raw)
    stem = re.sub(r"_+", "_", stem).strip("_")
    return stem or "UNKNOWN"


def normalize_name_key(name: str) -> str:
    s = normalize_passenger_stem(name)
    return s.replace("_", " ") if s != "UNKNOWN" else ""


def passenger_index_for_name_key(
    passengers: Sequence[PassengerRow], name_key: str
) -> Optional[int]:
    nk = name_key.upper().replace("_", " ")
    for i, (display_name, _) in enumerate(passengers):
        if normalize_name_key(display_name).upper() == nk:
            return i
    return None


# ── 航段探索 ──────────────────────────────────────────────────────────────────


def discover_segments(doc: fitz.Document) -> List[Segment]:
    n = doc.page_count
    header_indices = sorted(
        i for i in range(n) if doc[i].search_for("PREPARED FOR")
    )
    segments: List[Segment] = []
    consumed_detail: set[int] = set()

    for h in header_indices:
        d: Optional[int] = None
        for p in range(h + 1, n):
            if p in consumed_detail:
                continue
            if doc[p].search_for("Passenger Name:"):
                d = p
                break
        if d is None:
            LOG.warning(
                "第 %d 頁有 PREPARED FOR，但之後找不到可用的旅客明細頁（Passenger Name:），略過。",
                h + 1,
            )
            continue

        detail_page = doc[d]
        if not detail_page.search_for("DEPARTURE"):
            LOG.warning(
                "第 %d 頁（明細）未偵測到 DEPARTURE，仍嘗試處理。",
                d + 1,
            )

        pax = extract_passengers(detail_page)
        if not pax:
            LOG.warning("第 %d 頁（明細）無法解析旅客清單，略過此航段。", d + 1)
            continue

        sort_d = parse_segment_sort_date(doc, h, d)
        segments.append(
            Segment(
                header_page=h,
                detail_page=d,
                passengers=tuple(pax),
                sort_date=sort_d,
            )
        )
        consumed_detail.add(d)

    return segments


def build_passenger_segment_map(
    segments: Sequence[Segment],
) -> Tuple[DefaultDict[str, List[PaxSegmentRecord]], List[str]]:
    by_name: DefaultDict[str, List[PaxSegmentRecord]] = defaultdict(list)
    first_seen_order: List[str] = []

    for si, seg in enumerate(segments):
        for display_name, ticket in seg.passengers:
            nk = normalize_name_key(display_name)
            if not nk:
                LOG.warning(
                    "航段 %d（明細第 %d 頁）有無法正規化姓名的列，略過該列。",
                    si + 1,
                    seg.detail_page + 1,
                )
                continue
            if nk not in first_seen_order:
                first_seen_order.append(nk)
            by_name[nk].append(
                PaxSegmentRecord(
                    segment_index=si,
                    display_name=display_name,
                    ticket=ticket,
                    detail_page=seg.detail_page,
                )
            )

    for nk in by_name:
        by_name[nk].sort(
            key=lambda r: (
                segments[r.segment_index].sort_date,
                segments[r.segment_index].detail_page,
                r.segment_index,
            )
        )

    return by_name, first_seen_order


# ── 準備單一航段（遮罩後的暫存 PDF）──────────────────────────────────────────


def prepare_passenger_segment(
    input_pdf: str,
    segment: Segment,
    name_key: str,
    passengers: Sequence[PassengerRow],
) -> Optional[PreparedSegment]:
    pax_idx = passenger_index_for_name_key(passengers, name_key)
    if pax_idx is None:
        return None

    _, target_ticket = passengers[pax_idx]

    doc = fitz.open(input_pdf)
    src_doc = fitz.open(input_pdf)

    ph = doc[segment.header_page]
    pt = doc[segment.detail_page]
    src_pt = src_doc[segment.detail_page]

    prep_x0, prep_y0, _, _, prep_size = find_line_bbox(ph, "PREPARED FOR")
    _, res_y0, _, res_y1, res_size = find_line_bbox(ph, "RESERVATION CODE")

    if prep_y0 is None:
        prep_y0, prep_x0, prep_size = 60.0, 30.0, 8.0
    if res_y0 is None:
        res_y0, res_x1, res_y1, res_size = 500.0, 200.0, 515.0, 8.0

    res_code = reservation_code_from_header(ph)
    names_x1 = get_names_x1(ph, prep_y0, res_y0) or (ph.rect.width * 0.5)

    ph.draw_rect(
        fitz.Rect(0, prep_y0, names_x1, res_y0),
        color=(1, 1, 1),
        fill=(1, 1, 1),
        width=0,
    )
    ph.draw_rect(
        fitz.Rect(0, res_y0 - 1, ph.rect.width, res_y1 + 1),
        color=(1, 1, 1),
        fill=(1, 1, 1),
        width=0,
    )
    ph.insert_text(
        (prep_x0, prep_y0 + prep_size),
        f"RESERVATION CODE   {res_code}",
        fontsize=res_size,
        color=(0, 0, 0),
    )

    pw = pt.rect.width

    pax_hdr_y1 = None
    for b in src_pt.get_text("dict")["blocks"]:
        for line in b.get("lines", []):
            txt = "".join(s["text"] for s in line["spans"])
            if "Passenger Name" in txt:
                pax_hdr_y1 = max(s["bbox"][3] for s in line["spans"])
                break
        if pax_hdr_y1:
            break

    first_row_name = passengers[0][0]
    first_insts = pt.search_for(first_row_name)
    first_y0 = min(i.y0 for i in first_insts) - 1 if first_insts else 0.0
    if pax_hdr_y1 and first_y0 < pax_hdr_y1:
        first_y0 = pax_hdr_y1 + 1

    row_spans = []
    t_insts = src_pt.search_for(target_ticket)
    if t_insts:
        first_t = min(t_insts, key=lambda r: r.y0)
        ry0, ry1 = first_t.y0, first_t.y1
        for b in src_pt.get_text("dict")["blocks"]:
            for line in b.get("lines", []):
                for s in line["spans"]:
                    if s["bbox"][1] < ry1 and s["bbox"][3] > ry0 and s["text"].strip():
                        row_spans.append(s)

    name_lines, checkin_lines = [], []
    for b in src_pt.get_text("dict")["blocks"]:
        for line in b.get("lines", []):
            line_y0 = min(s["bbox"][1] for s in line["spans"])
            txt = "".join(s["text"] for s in line["spans"])
            if "»" in txt or "\u00bb" in txt:
                name_lines.append((line_y0, line["spans"]))
            elif "Check-In" in txt or "Required" in txt:
                checkin_lines.append((line_y0, line["spans"]))

    name_lines.sort(key=lambda x: x[0])
    checkin_lines.sort(key=lambda x: x[0])
    existing_pos = {(round(s["bbox"][0]), round(s["bbox"][1])) for s in row_spans}

    for lines_list in (name_lines, checkin_lines):
        if pax_idx < len(lines_list):
            _, line_spans = lines_list[pax_idx]
            for s in line_spans:
                pos = (round(s["bbox"][0]), round(s["bbox"][1]))
                if pos not in existing_pos:
                    row_spans.append(s)
                    existing_pos.add(pos)

    pt.add_redact_annot(fitz.Rect(0, first_y0, pw, pt.rect.height))
    pt.apply_redactions()

    if row_spans and pax_hdr_y1 is not None:
        orig_y0 = min(s["bbox"][1] for s in row_spans)
        y_delta = (pax_hdr_y1 + 2) - orig_y0
        for s in row_spans:
            x = s["bbox"][0]
            y = s["bbox"][3] + y_delta
            sz = s["size"]
            c = s.get("color", 0)
            rgb = (
                ((c >> 16) & 0xFF) / 255.0,
                ((c >> 8) & 0xFF) / 255.0,
                (c & 0xFF) / 255.0,
            )
            pt.insert_text((x, y), s["text"], fontsize=sz, color=rgb)

    src_doc.close()

    buf = doc.tobytes()
    doc.close()

    comp = fitz.open("pdf", buf)
    h_idx = segment.header_page
    d_idx = segment.detail_page
    if h_idx >= comp.page_count or d_idx >= comp.page_count:
        comp.close()
        return None

    ph2 = comp[h_idx]
    pt2 = comp[d_idx]
    pw = ph2.rect.width

    half_h = pt2.rect.height / 2
    dep_insts = pt2.search_for("DEPARTURE")
    dep_y0 = (
        min(r.y0 for r in dep_insts if r.y0 < half_h) - 5
        if dep_insts
        else 0.0
    )

    pax_insts = pt2.search_for(target_ticket)
    content_bottom = (
        min(r.y1 for r in pax_insts) + 15 if pax_insts else dep_y0 + 300.0
    )

    return PreparedSegment(
        comp=comp,
        header_page=h_idx,
        detail_page=d_idx,
        pw=pw,
        header_h=HEADER_CLIP_Y1,
        dep_y0=dep_y0,
        content_bottom=content_bottom,
    )


def render_segment_to_page(
    out_page: fitz.Page,
    prep: PreparedSegment,
    start_y: float,
    scale: float,
) -> float:
    """
    在 out_page 的 y=start_y 處繪製本航段（標頭裁切 + 明細裁切）。
    回傳此航段在輸出座標中佔用的總高度（已乘 scale），不含與下一段的間距。
    """
    hh = prep.header_h * scale
    dh = prep.detail_h * scale
    pw = prep.pw
    clip_h = fitz.Rect(0, 0, pw, prep.header_h)
    clip_d = fitz.Rect(0, prep.dep_y0, pw, prep.content_bottom)

    dest_h = fitz.Rect(0, start_y, pw, start_y + hh)
    dest_d = fitz.Rect(0, start_y + hh, pw, start_y + hh + dh)

    out_page.show_pdf_page(dest_h, prep.comp, prep.header_page, clip=clip_h)
    out_page.show_pdf_page(dest_d, prep.comp, prep.detail_page, clip=clip_d)

    return hh + dh


def build_one_passenger_single_a4(
    input_pdf: str,
    sorted_records: Sequence[PaxSegmentRecord],
    segments: Sequence[Segment],
    name_key: str,
    name_stem: str,
) -> Optional[Tuple[bytes, int]]:
    prepared: List[PreparedSegment] = []
    try:
        for rec in sorted_records:
            seg = segments[rec.segment_index]
            prep = prepare_passenger_segment(input_pdf, seg, name_key, seg.passengers)
            if prep is None:
                LOG.warning(
                    "  旅客 %s 於航段（標頭第 %d 頁／明細第 %d 頁）無法對應列，略過該段。",
                    name_stem,
                    seg.header_page + 1,
                    seg.detail_page + 1,
                )
                continue
            prepared.append(prep)

        if not prepared:
            return None

        out_w = max(p.pw for p in prepared)
        out_w = max(out_w, A4_WIDTH)

        block_heights = [p.header_h + p.detail_h for p in prepared]
        n_gap = max(0, len(prepared) - 1)
        raw_total = sum(block_heights) + n_gap * SEGMENT_GAP
        usable = A4_HEIGHT - PAGE_TOP_MARGIN - PAGE_BOTTOM_MARGIN
        scale = 1.0 if raw_total <= 0 else min(1.0, usable / raw_total)

        if scale < 1.0:
            LOG.info(
                "  旅客 %s：垂直內容高度 %.0f，縮放為 %.1f%% 以置入單頁 A4。",
                name_stem,
                raw_total,
                scale * 100.0,
            )

        out = fitz.open()
        page = out.new_page(width=out_w, height=A4_HEIGHT)
        cur_y = PAGE_TOP_MARGIN
        gap_scaled = SEGMENT_GAP * scale

        for i, prep in enumerate(prepared):
            used = render_segment_to_page(page, prep, cur_y, scale)
            cur_y += used
            if i < len(prepared) - 1:
                cur_y += gap_scaled

        data = out.tobytes()
        out.close()
        return data, len(prepared)
    finally:
        for p in prepared:
            p.comp.close()


# ── 主流程 ────────────────────────────────────────────────────────────────────


def process_pdf(input_path: str, out_dir: str) -> int:
    input_path = os.path.abspath(input_path)
    if not os.path.isfile(input_path):
        LOG.error("找不到檔案：%s", input_path)
        return 1

    stem = os.path.splitext(os.path.basename(input_path))[0]
    os.makedirs(out_dir, exist_ok=True)

    doc = fitz.open(input_path)
    try:
        segments = discover_segments(doc)
        LOG.info(
            "找到 %d 個航段（每段：PREPARED FOR 頁 + 其後第一個旅客明細頁；已解析出發日期供排序）。",
            len(segments),
        )

        if not segments:
            LOG.error("無法辨識任何航段，請確認為泰獅航多人機票格式。")
            return 1

        by_name, first_seen_order = build_passenger_segment_map(segments)
        LOG.info(
            "找到 %d 位旅客（以正規化姓名合併；各航段依日期由早到晚垂直排列於單頁）。",
            len(by_name),
        )

        for nk in first_seen_order:
            if nk not in by_name:
                continue
            records = by_name[nk]
            display_name = records[0].display_name
            name_stem = normalize_passenger_stem(display_name)
            out_name = f"{stem}_{name_stem}.pdf"
            out_path = os.path.join(out_dir, out_name)

            result = build_one_passenger_single_a4(
                input_path,
                records,
                segments,
                nk,
                name_stem,
            )
            if result is None:
                LOG.warning("  旅客 %s 無可輸出內容，略過。", name_stem)
                continue

            blob, n_merged = result
            with open(out_path, "wb") as f:
                f.write(blob)
            LOG.info(
                "輸出完成：%s（1 頁，合併 %d 個航段）",
                out_name,
                n_merged,
            )

    finally:
        doc.close()

    LOG.info("全部處理結束，輸出目錄：%s", out_dir)
    return 0


def parse_args(argv: Optional[Sequence[str]] = None) -> argparse.Namespace:
    p = argparse.ArgumentParser(
        description="將泰國獅子航空多人機票 PDF 依旅客拆成獨立 PDF（多航段單頁垂直合併）。",
    )
    p.add_argument(
        "pdf",
        metavar="INPUT.pdf",
        help="輸入的 PDF 檔案路徑",
    )
    p.add_argument(
        "-o",
        "--out",
        dest="out_dir",
        default=None,
        help="輸出資料夾（預設：<原檔名>_split）",
    )
    return p.parse_args(argv)


def main(argv: Optional[Sequence[str]] = None) -> int:
    _setup_logging()
    args = parse_args(argv)
    pdf = args.pdf
    stem = os.path.splitext(os.path.basename(pdf))[0]
    out_dir = args.out_dir if args.out_dir else os.path.join(
        os.path.dirname(os.path.abspath(pdf)),
        f"{stem}_split",
    )
    return process_pdf(pdf, out_dir)


if __name__ == "__main__":
    sys.exit(main())
