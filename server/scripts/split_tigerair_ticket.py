#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
台灣虎航 (Tigerair) 團體行程表 PDF 拆分工具

Clip1：頁首至「訂位代號」列 y1+10（挖除總名單）。Clip2／Clip3：保留來源水平全寬，貼上時 x=0 與頁首區對齊。
Clip3：詳細訊息以相鄰姓名錨點 Y 中點分群單字；列間縫隙中點下刀；可跨頁續貼。
"""

from __future__ import annotations

import argparse
import math
import re
from pathlib import Path
from typing import List, Optional, Set, Tuple

import fitz  # PyMuPDF

# ── 錨點關鍵字 ─────────────────────────────────────────────────────────────
MARKER_PNR_LABEL = "訂位代號"
MARKER_FLIGHT = "航班資訊"
MARKER_DETAIL = "詳細訊息"

# ── 頁尾廣告／聲明：向下掃描時遇見即硬性截止（忽略大小寫）──────────────────
FOOTER_MARKERS = (
    "CHUBB",
    "安達產物",
    "KLOOK",
    "LINE GO",
    "BREEZE",
    "重要資訊",
    "請務必準時抵達機場",
)

# ── 稱謂：normalize 時剔除 ─────────────────────────────────────────────────
TITLE_PREFIXES = (
    "MR", "MRS", "MS", "MISS", "MSTR", "MASTER", "DR",
    "先生", "女士", "小姐",
)

# ── 排版 ───────────────────────────────────────────────────────────────────
A4_W = 595.28
A4_H = 841.89
PAGE_SAFE_Y = 800.0
PAGE_TOP_Y = 36.0
CLIP_GAP = 14.0
# Clip2：航班資訊／詳細訊息標題上下預留（與需求一致 ±10）
FLIGHT_CLIP_PAD = 10.0
DETAIL_PAD_TOP = 2.0
DETAIL_PAD_BOTTOM = 6.0
PNR_BOTTOM_PAD = 10.0

CONTENT_TOP_PAD = 2.0
CONTENT_LAST_ROW_BOTTOM_PAD = 4.0
LAST_ROW_BELOW_FALLBACK = 45.0
# get_text("words")：過濾異常高大單字（直排／合併假象），避免當成列高
WORD_MAX_HEIGHT_PT = 30.0
WORD_MIN_HEIGHT_PT = 0.5
# 同一列內多個姓名 y0 僅差 1pt 內視為同一水平列
NAME_Y0_SAME_ROW_MERGE_PT = 1.0
# ── 正則 ─────────────────────────────────────────────────────────────────────
RE_MULTI_SPACE = re.compile(r"\s+")
# 名單列常見：MS JIALING WU / MR CHEN DA WEI
RE_LIST_NAME_LINE = re.compile(
    r"^(?:(?:MR|MRS|MS|MISS|MSTR|MASTER|DR)\s+)+(.+)$",
    re.I,
)
# 詳細訊息區內帶稱謂之英文姓名（忽略大小寫；姓名本體允許多個單字）
RE_DETAILS_NAME = re.compile(
    r"\b(MR|MRS|MS|MISS|MSTR|MASTER|DR)\s+([A-Za-z]+(?:\s+[A-Za-z]+)*)\b",
    re.I,
)


# ═══════════════════════════════════════════════════════════════════════════
# 基礎工具
# ═══════════════════════════════════════════════════════════════════════════

def normalize_spaces(s: str) -> str:
    return RE_MULTI_SPACE.sub(" ", (s or "")).strip()


def _finite_rect(r: fitz.Rect) -> bool:
    return all(math.isfinite(v) for v in (r.x0, r.y0, r.x1, r.y1))


def _ensure_valid_clip(r: fitz.Rect, page_rect: fitz.Rect, min_h: float = 5.0) -> fitz.Rect:
    """
    將裁切矩形強制限制在 page_rect 內，寬高至少 5（高度至少 min_h），避免 show_pdf_page 拋錯。
    """
    rr = fitz.Rect(
        max(page_rect.x0, r.x0),
        max(page_rect.y0, r.y0),
        min(page_rect.x1, r.x1),
        min(page_rect.y1, r.y1),
    )
    if (not _finite_rect(rr)) or rr.is_empty or rr.width < 5 or rr.height < max(5.0, min_h):
        raise ValueError(f"Invalid clip: {rr}")
    return rr


def page_lines(page: fitz.Page) -> List[Tuple[fitz.Rect, str]]:
    """擷取頁面文字行 (bbox, text)，依閱讀順序排序。"""
    d = page.get_text("dict")
    lines: List[Tuple[fitz.Rect, str]] = []
    for b in d.get("blocks", []):
        if b.get("type") != 0:
            continue
        for ln in b.get("lines", []):
            txt = "".join(s.get("text", "") for s in ln.get("spans", [])).strip()
            if txt:
                lines.append((fitz.Rect(ln["bbox"]), txt))
    lines.sort(key=lambda x: (x[0].y0, x[0].x0))
    return lines


def line_hits_footer(txt: str) -> bool:
    """中英混合：中文關鍵字用原樣比對，英文用大小寫不敏感。"""
    t = normalize_spaces(txt)
    tu = t.upper()
    for m in FOOTER_MARKERS:
        if re.search(r"[\u4e00-\u9fff]", m):
            if m in t:
                return True
        elif m.upper() in tu:
            return True
    return False


def normalize_name_tokens(name_str: str) -> Set[str]:
    """
    過濾稱謂、轉大寫、拆成 token 集合。
    用於比對「名單 MS JIALING WU」與「表格 MS WU JIALING」是否為同一人。
    """
    s = normalize_spaces(name_str)
    s = s.upper()
    parts = s.split()
    out: List[str] = []
    for p in parts:
        p2 = re.sub(r"^[^A-Z0-9\u4E00-\u9FFF]+|[,，．.]+$", "", p)
        if not p2:
            continue
        if p2 in TITLE_PREFIXES:
            continue
        out.append(p2)
    return set(out)


def tokens_match_passenger(a: Set[str], b: Set[str]) -> bool:
    """兩組 token 互相包含（相等）即視為同一人。"""
    if not a or not b:
        return False
    return a <= b and b <= a


def safe_filename_component(s: str) -> str:
    s = normalize_spaces(s).upper()
    s = re.sub(r"[^A-Z0-9\u4E00-\u9FFF]+", "_", s).strip("_")
    return s[:120] if len(s) > 120 else s


def find_first_line_global(
    doc: fitz.Document, needle: str, casefold: bool = True
) -> Optional[Tuple[int, fitz.Rect, str]]:
    n = needle if not casefold else needle.upper()
    for pno in range(len(doc)):
        for bbox, txt in page_lines(doc[pno]):
            t = normalize_spaces(txt)
            if casefold:
                t = t.upper()
                if n in t:
                    return pno, bbox, txt
            else:
                if needle in t:
                    return pno, bbox, txt
    return None


def clip_rect_trim_left_blank(page: fitz.Page, r: fitz.Rect) -> fitz.Rect:
    """保留來源水平佈局，不動態左裁（與頁首／航班區 x 對齊一致）。"""
    return r


def clip_rect_trim_left_blank_detail(page: fitz.Page, r: fitz.Rect) -> fitz.Rect:
    """
    精準平移對齊：將該列資料的最左側有效文字，精準平移到 X 座標 166.0
    （即上方「出發」與「機場名稱」的垂直對齊線）。
    """
    pr = page.rect
    words = page.get_text("words") or []
    min_x0: Optional[float] = None

    for w in words:
        if len(w) < 5:
            continue
        wx0, wy0, wx1, wy1 = float(w[0]), float(w[1]), float(w[2]), float(w[3])
        text = str(w[4]).strip()

        if text in ("去程", "回程", "詳細訊息") or not text:
            continue

        if wy1 <= r.y0 or wy0 >= r.y1:
            continue

        min_x0 = wx0 if min_x0 is None else min(min_x0, wx0)

    if min_x0 is not None:
        target_x = 166.0
        new_x0 = min_x0 - target_x
        return fitz.Rect(new_x0, r.y0, pr.width + new_x0, r.y1)

    return r


def find_footer_y_on_page(page: fitz.Page, y_start: float, y_hard: float) -> float:
    """在 [y_start, y_hard] 內若出現 FOOTER_MARKERS，回傳該行 y0；否則 y_hard。"""
    for bbox, txt in page_lines(page):
        if bbox.y0 <= y_start + 0.5 or bbox.y0 >= y_hard - 0.5:
            continue
        if line_hits_footer(txt):
            return bbox.y0
    return y_hard


# ═══════════════════════════════════════════════════════════════════════════
# 旅客清單：頂部名單（可選）＋詳細訊息表格（全域後備）
# ═══════════════════════════════════════════════════════════════════════════

def parse_passengers_from_list(
    doc: fitz.Document,
    pnr_pno: int,
    pnr_y0: float,
    flight_pno: int,
    flight_y0: float,
) -> List[Tuple[str, Set[str]]]:
    """
    在「訂位代號」列之下至「航班資訊」列之上，嘗試掃描頂部總名單（易飛等變體）。
    找不到任何列時回傳空串列，不拋錯（改由詳細訊息區萃取）。
    """
    seen: List[Tuple[str, Set[str]]] = []
    seen_sets: List[Set[str]] = []

    for pno in range(pnr_pno, flight_pno + 1):
        y_lo = pnr_y0 if pno == pnr_pno else 0.0
        for bbox, txt in page_lines(doc[pno]):
            if bbox.y0 < y_lo - 1:
                continue
            if pno == flight_pno and bbox.y0 >= flight_y0 - 3:
                break
            t = normalize_spaces(txt)
            if MARKER_FLIGHT in t:
                break
            m = RE_LIST_NAME_LINE.match(t)
            if not m:
                continue
            inner = normalize_spaces(m.group(1))
            toks = normalize_name_tokens(inner)
            if len(toks) < 2:
                continue
            dup = False
            for prev in seen_sets:
                if tokens_match_passenger(toks, prev):
                    dup = True
                    break
            if dup:
                continue
            seen.append((t, toks))
            seen_sets.append(toks)

    return seen


def _iter_detail_section_lines(
    doc: fitz.Document, detail_pno: int, detail_y0: float
) -> List[Tuple[int, str]]:
    """詳細訊息錨點以下、頁尾標記以前之文字行（頁碼, 文字）。"""
    out: List[Tuple[int, str]] = []
    for pno in range(detail_pno, len(doc)):
        page = doc[pno]
        pr = page.rect
        y_lo = detail_y0 - 2 if pno == detail_pno else 0.0
        footer_cut = find_footer_y_on_page(page, y_lo, pr.height - 5)
        for bbox, txt in page_lines(page):
            if bbox.y0 < y_lo - 1:
                continue
            if bbox.y0 >= footer_cut - 1:
                break
            if line_hits_footer(txt):
                break
            out.append((pno, txt))
    return out


def parse_passengers_from_details_table(
    doc: fitz.Document, detail_pno: int, detail_y0: float
) -> List[Tuple[str, Set[str]]]:
    """
    自「詳細訊息」以下全文掃描，以 RE_DETAILS_NAME 擷取稱謂＋姓名；
    以 normalize 後之 token 集合去重，保留首次出現順序（KKDAY 等無頂部名單變體）。
    """
    seen: List[Tuple[str, Set[str]]] = []
    seen_key: List[Set[str]] = []

    for _pno, txt in _iter_detail_section_lines(doc, detail_pno, detail_y0):
        t = normalize_spaces(txt)
        for m in RE_DETAILS_NAME.finditer(t):
            title = m.group(1).upper()
            body = normalize_spaces(m.group(2))
            if not body:
                continue
            toks = normalize_name_tokens(body)
            if len(toks) < 2:
                continue
            dup = False
            for prev in seen_key:
                if tokens_match_passenger(toks, prev):
                    dup = True
                    break
            if dup:
                continue
            display = f"{title} {body}"
            seen.append((display, toks))
            seen_key.append(toks)

    return seen


def list_order_to_filename_stem(display_name_line: str) -> str:
    """
    名單為 First Last（如 MS JIALING WU）→ 輸出檔名 姓氏_名字 = WU_JIALING。
    """
    t = normalize_spaces(display_name_line)
    m = RE_LIST_NAME_LINE.match(t)
    body = m.group(1).strip() if m else t
    toks = [x for x in body.split() if x.upper() not in TITLE_PREFIXES]
    if len(toks) >= 2:
        first_parts = toks[:-1]
        last = toks[-1]
        first = "_".join(safe_filename_component(x) for x in first_parts)
        return f"{safe_filename_component(last)}_{first}"
    return safe_filename_component(body)


def details_table_to_filename_stem(display_name_line: str) -> str:
    """
    詳細訊息表格多為「姓 名」（如 MS WU JIALING）→ 檔名 WU_JIALING。
    """
    t = normalize_spaces(display_name_line)
    toks_raw = [x for x in t.split() if x.upper() not in TITLE_PREFIXES]
    if len(toks_raw) >= 2:
        last = toks_raw[0]
        first = "_".join(safe_filename_component(x) for x in toks_raw[1:])
        return f"{safe_filename_component(last)}_{first}"
    return safe_filename_component(" ".join(toks_raw))


def passenger_filename_stem(display: str, source: str) -> str:
    """source 為 'list' 或 'details'。"""
    if source == "list":
        return list_order_to_filename_stem(display)
    return details_table_to_filename_stem(display)


# ═══════════════════════════════════════════════════════════════════════════
# Clip 1：僅頁首～「訂位代號」列底 + 10px（總名單在來源上被跳過）
# ═══════════════════════════════════════════════════════════════════════════

def build_pnr_top_clip(
    doc: fitz.Document, pnr_pno: int, pnr_line_bbox: fitz.Rect
) -> List[Tuple[int, fitz.Rect]]:
    pr = doc[pnr_pno].rect
    y1 = min(pr.height - 1.0, pnr_line_bbox.y1 + PNR_BOTTOM_PAD)
    return [
        (
            pnr_pno,
            _ensure_valid_clip(fitz.Rect(0, 0, pr.width, y1), pr, min_h=5.0),
        )
    ]


def apply_top_block_white_masks(
    out_page: fitz.Page,
    src_page: fitz.Page,
    y_offset: float,
    clip_local: fitz.Rect,
) -> None:
    """
    放棄所有白色遮蔽，100% 保留原始頂部黑框與資訊的完整性。
    """
    return


# ═══════════════════════════════════════════════════════════════════════════
# Clip 2：航班資訊（至詳細訊息標題上緣）
# ═══════════════════════════════════════════════════════════════════════════

def collect_flight_info_segments(
    doc: fitz.Document,
    flight_pno: int,
    flight_y0: float,
    detail_pno: int,
    detail_y0: float,
) -> List[Tuple[int, fitz.Rect]]:
    """「航班資訊」y0-10 至「詳細訊息」y0-10，支援跨頁。"""
    segs: List[Tuple[int, fitz.Rect]] = []
    y0_first = max(0.0, flight_y0 - FLIGHT_CLIP_PAD)
    y1_last = max(5.0, detail_y0 - FLIGHT_CLIP_PAD)

    if flight_pno == detail_pno:
        pr = doc[flight_pno].rect
        raw = fitz.Rect(0, y0_first, pr.width, y1_last)
        cr = clip_rect_trim_left_blank(doc[flight_pno], raw)
        segs.append(
            (flight_pno, _ensure_valid_clip(cr, pr, min_h=30.0)),
        )
        return segs

    pr0 = doc[flight_pno].rect
    raw0 = fitz.Rect(0, y0_first, pr0.width, pr0.height)
    cr0 = clip_rect_trim_left_blank(doc[flight_pno], raw0)
    segs.append(
        (flight_pno, _ensure_valid_clip(cr0, pr0, min_h=20.0)),
    )
    for pno in range(flight_pno + 1, detail_pno):
        pm = doc[pno].rect
        rawm = fitz.Rect(0, 0, pm.width, pm.height)
        crm = clip_rect_trim_left_blank(doc[pno], rawm)
        segs.append(
            (pno, _ensure_valid_clip(crm, pm, min_h=20.0)),
        )
    prd = doc[detail_pno].rect
    rawd = fitz.Rect(0, 0, prd.width, y1_last)
    crd = clip_rect_trim_left_blank(doc[detail_pno], rawd)
    segs.append(
        (detail_pno, _ensure_valid_clip(crd, prd, min_h=15.0)),
    )
    return segs


# ═══════════════════════════════════════════════════════════════════════════
# Clip 3：詳細訊息 — 依 Token 找列、頁尾截止、跨頁 spillover
# ═══════════════════════════════════════════════════════════════════════════

def line_name_tokens_if_any(txt: str) -> Optional[Set[str]]:
    """
    若此行含稱謂＋英文姓名（KKDAY／易飛常見前綴航班號，如 IT551 MR CHANG YUHAO）則回傳 token set。
    使用 RE_DETAILS_NAME 全域搜尋，不以行首 ^ 限制。
    """
    t = normalize_spaces(txt)
    if len(t) < 5:
        return None
    m = RE_DETAILS_NAME.search(t)
    if m:
        frag = f"{m.group(1)} {m.group(2)}"
        toks = normalize_name_tokens(frag)
        return toks if len(toks) >= 2 else None
    # 無稱謂時，若整行幾乎都是字母且 2–6 token 也可能為姓名（表格內偶發）
    if re.match(r"^[A-Za-z\s]{6,}$", t) and 2 <= len(t.split()) <= 6:
        toks = normalize_name_tokens(t)
        return toks if len(toks) >= 2 else None
    return None


def collect_detail_lines_ordered(
    doc: fitz.Document,
    detail_pno: int,
    detail_y0: float,
) -> List[Tuple[int, fitz.Rect, str, Optional[Set[str]]]]:
    """
    自「詳細訊息」以下至文件尾（遇頁尾標記則該頁截斷），收集所有文字行。
    每項：(頁碼, bbox, 原文, 若為姓名錨點則 token set 否則 None)
    """
    rows: List[Tuple[int, fitz.Rect, str, Optional[Set[str]]]] = []
    for pno in range(detail_pno, len(doc)):
        page = doc[pno]
        pr = page.rect
        y_lo = detail_y0 - 2 if pno == detail_pno else 0.0
        footer_cut = find_footer_y_on_page(page, y_lo, pr.height - 5)
        for bbox, txt in page_lines(page):
            if bbox.y0 < y_lo - 1:
                continue
            if bbox.y0 >= footer_cut - 1:
                break
            if line_hits_footer(txt):
                break
            nt = line_name_tokens_if_any(txt)
            rows.append((pno, bbox, txt, nt))
    return rows


def _name_owner_index(
    nt: Optional[Set[str]], all_passenger_token_sets: List[Set[str]]
) -> Optional[int]:
    if nt is None:
        return None
    for j, pts in enumerate(all_passenger_token_sets):
        if tokens_match_passenger(nt, pts):
            return j
    return None


def _sorted_passenger_name_y0s_on_page(
    page: fitz.Page,
    y_lo: float,
    hard_y: float,
    all_passenger_token_sets: List[Set[str]],
) -> List[float]:
    """
    該頁詳細區 [y_lo, hard_y] 內，凡辨識為「已知旅客」之姓名列，收集 y0 後由小到大排序。
    不依賴餐點／行李 bbox，僅姓名行 y0 作為列分界。
    """
    ys: List[float] = []
    for bbox, txt in page_lines(page):
        if bbox.y0 < y_lo - 1:
            continue
        if bbox.y0 >= hard_y - 0.5:
            break
        if line_hits_footer(txt):
            break
        nt = line_name_tokens_if_any(txt)
        if nt is None or _name_owner_index(nt, all_passenger_token_sets) is None:
            continue
        ys.append(bbox.y0)
    ys.sort()
    return ys


def _name_y0_row_levels(ys_sorted: List[float]) -> List[float]:
    """將同一列上多個姓名 y0 合併為單一水平層級，再與下列區隔。"""
    if not ys_sorted:
        return []
    levels = [ys_sorted[0]]
    for y in ys_sorted[1:]:
        if y - levels[-1] > NAME_Y0_SAME_ROW_MERGE_PT:
            levels.append(y)
    return levels


def _page_words_valid(page: fitz.Page) -> List[Tuple[float, float, float, float, str, int, int]]:
    """
    全頁單字 (x0,y0,x1,y1,text,block,line)；高度須 < WORD_MAX_HEIGHT_PT。
    PyMuPDF 合併區塊在 line 層會變巨型 bbox，單字層可拆開。
    """
    raw = page.get_text("words") or []
    out: List[Tuple[float, float, float, float, str, int, int]] = []
    for w in raw:
        if len(w) < 5:
            continue
        x0, y0, x1, y1 = float(w[0]), float(w[1]), float(w[2]), float(w[3])
        if not all(math.isfinite(v) for v in (x0, y0, x1, y1)):
            continue
        h = y1 - y0
        if h >= WORD_MAX_HEIGHT_PT or h < WORD_MIN_HEIGHT_PT:
            continue
        t = str(w[4])
        bno = int(w[5]) if len(w) > 5 else 0
        lno = int(w[6]) if len(w) > 6 else 0
        out.append((x0, y0, x1, y1, t, bno, lno))
    out.sort(key=lambda r: (r[1], r[0]))
    return out


def _nearest_row_level_index(y_coords: List[float], target_y0: float) -> int:
    """y_coords 為已排序之列級 y0；回傳與 target_y0 最接近之索引。"""
    if not y_coords:
        return -1
    return min(range(len(y_coords)), key=lambda i: abs(y_coords[i] - target_y0))


def _get_row_extents(
    words: List[Tuple[float, float, float, float, str, int, int]],
    y_coords: List[float],
    idx: int,
    y_lo: float,
    footer_y: float,
) -> Tuple[float, float]:
    """利用上下姓名錨點將單字分群。加入動態行距判定以適應緊密排版。"""
    n = len(y_coords)

    if idx == 0:
        top_bound = y_lo
    else:
        gap_above = y_coords[idx] - y_coords[idx - 1]
        top_bound = (
            y_coords[idx] - 4.0
            if gap_above < 18.0
            else (y_coords[idx - 1] + y_coords[idx]) / 2.0
        )

    if idx == n - 1:
        bot_bound = footer_y
    else:
        gap_below = y_coords[idx + 1] - y_coords[idx]
        bot_bound = (
            y_coords[idx + 1] - 4.0
            if gap_below < 18.0
            else (y_coords[idx] + y_coords[idx + 1]) / 2.0
        )

    min_y0: Optional[float] = None
    max_y1: Optional[float] = None
    for _x0, y0, _x1, y1, _t, _bno, _lno in words:
        if y0 >= top_bound and y0 < bot_bound:
            if y1 > footer_y:
                continue
            min_y0 = y0 if min_y0 is None else min(min_y0, y0)
            max_y1 = y1 if max_y1 is None else max(max_y1, y1)

    if min_y0 is None or max_y1 is None:
        return y_coords[idx], y_coords[idx] + 15.0
    return min_y0, max_y1


def _content_aware_vertical_band_for_row(
    page: fitz.Page,
    y_coords: List[float],
    idx: int,
    target_y0: float,
    y_lo: float,
    hard: float,
    footer_y: float,
    passenger_tokens: Set[str],
    all_passenger_token_sets: List[Set[str]],
) -> Tuple[float, float]:
    _ = (passenger_tokens, all_passenger_token_sets)
    if idx < 0 or idx >= len(y_coords):
        return max(y_lo, target_y0 - CONTENT_TOP_PAD), min(
            target_y0 + LAST_ROW_BELOW_FALLBACK, hard
        )

    words = _page_words_valid(page)
    cur_min_y0, cur_max_y1 = _get_row_extents(words, y_coords, idx, y_lo, footer_y)

    if idx > 0:
        _, prev_max_y1 = _get_row_extents(words, y_coords, idx - 1, y_lo, footer_y)
        gap_top = cur_min_y0 - prev_max_y1
        if gap_top > 0:
            y_start = prev_max_y1 + gap_top / 2.0
        else:
            gap_above = y_coords[idx] - y_coords[idx - 1]
            y_start = (
                y_coords[idx] - 4.0
                if gap_above < 18.0
                else (y_coords[idx - 1] + y_coords[idx]) / 2.0
            )
    else:
        y_start = max(y_lo, cur_min_y0 - CONTENT_TOP_PAD)

    if idx < len(y_coords) - 1:
        next_min_y0, _ = _get_row_extents(words, y_coords, idx + 1, y_lo, footer_y)
        gap_bot = next_min_y0 - cur_max_y1
        if gap_bot > 0:
            y_end = cur_max_y1 + gap_bot / 2.0
        else:
            gap_below = y_coords[idx + 1] - y_coords[idx]
            y_end = (
                y_coords[idx + 1] - 4.0
                if gap_below < 18.0
                else (y_coords[idx] + y_coords[idx + 1]) / 2.0
            )
    else:
        y_end = cur_max_y1 + CONTENT_LAST_ROW_BOTTOM_PAD

    return max(y_lo, y_start), min(y_end, hard, footer_y)


def _content_aware_spill_band(
    page: fitz.Page,
    y_lo: float,
    hard: float,
    footer_y: float,
    y_coords: List[float],
    passenger_tokens: Set[str],
    all_passenger_token_sets: List[Set[str]],
) -> Tuple[float, float]:
    _ = (passenger_tokens, all_passenger_token_sets)
    if not y_coords:
        return y_lo, min(y_lo + LAST_ROW_BELOW_FALLBACK, hard, footer_y)

    words = _page_words_valid(page)
    next_min_y0, _ = _get_row_extents(words, y_coords, 0, y_lo, footer_y)

    spill_min_y0: Optional[float] = None
    spill_max_y1: Optional[float] = None
    for x0, y0, x1, y1, _t, _bno, _lno in words:
        if y0 >= y_lo and y0 < next_min_y0 - 1.0:
            spill_min_y0 = y0 if spill_min_y0 is None else min(spill_min_y0, y0)
            spill_max_y1 = y1 if spill_max_y1 is None else max(spill_max_y1, y1)

    if spill_min_y0 is None or spill_max_y1 is None:
        return y_lo, min(max(y_lo + 5.0, next_min_y0 - 2.0), hard, footer_y)

    y_start = max(y_lo, spill_min_y0 - CONTENT_TOP_PAD)
    gap = next_min_y0 - spill_max_y1
    if gap > 0:
        y_end = spill_max_y1 + gap / 2.0
    else:
        y_end = spill_max_y1 + CONTENT_LAST_ROW_BOTTOM_PAD
    return y_start, min(y_end, hard, footer_y)


def _append_detail_row_clip(
    doc: fitz.Document,
    detail_pno: int,
    detail_y0: float,
    pno: int,
    band_y0: float,
    band_y1: float,
) -> Optional[Tuple[int, fitz.Rect]]:
    """依 [band_y0, band_y1] 裁切（全頁寬 x0=0..page.width，不動態左裁）。"""
    page = doc[pno]
    pr = page.rect
    y_lo = detail_y0 - 2 if pno == detail_pno else 0.0
    footer_y = find_footer_y_on_page(page, y_lo, pr.height)
    eff_footer = min(footer_y - 1.0, pr.height - 1.0)
    b0 = max(band_y0, y_lo)
    b1 = min(band_y1, eff_footer)
    if b1 <= b0 + 2:
        return None
    try:
        raw = fitz.Rect(0, b0, pr.width, b1)
        trimmed = clip_rect_trim_left_blank_detail(page, raw)
        return (pno, _ensure_valid_clip(trimmed, pr, min_h=5.0))
    except ValueError:
        return None


def _clips_row_based_for_detail_anchor(
    doc: fitz.Document,
    detail_pno: int,
    detail_y0: float,
    lines: List[Tuple[int, fitz.Rect, str, Optional[Set[str]]]],
    anchor_i: int,
    all_passenger_token_sets: List[Set[str]],
    passenger_tokens: Set[str],
) -> List[Tuple[int, fitz.Rect]]:
    """
    詳細訊息列：Y 以姓名錨點中點分群單字，列間縫隙中點下刀；跨頁延續用 spill 帶。
    """
    p_anchor, bb_anchor, _, _nt = lines[anchor_i]
    ty0 = bb_anchor.y0
    clips: List[Tuple[int, fitz.Rect]] = []
    p_cur = p_anchor
    on_anchor_page = True

    while p_cur < len(doc):
        page = doc[p_cur]
        y_lo = detail_y0 - 2 if p_cur == detail_pno else 0.0
        footer_y = find_footer_y_on_page(page, y_lo, page.rect.height)
        hard = min(footer_y - 1.0, page.rect.height - 1.0)
        raw_ys = _sorted_passenger_name_y0s_on_page(
            page, y_lo, hard, all_passenger_token_sets
        )
        y_coords = _name_y0_row_levels(raw_ys)

        if on_anchor_page:
            idx = _nearest_row_level_index(y_coords, ty0)
            band_y0, band_y1 = _content_aware_vertical_band_for_row(
                page,
                y_coords,
                idx,
                ty0,
                y_lo,
                hard,
                footer_y,
                passenger_tokens,
                all_passenger_token_sets,
            )
            on_anchor_page = False
        else:
            band_y0, band_y1 = _content_aware_spill_band(
                page,
                y_lo,
                hard,
                footer_y,
                y_coords,
                passenger_tokens,
                all_passenger_token_sets,
            )

        c = _append_detail_row_clip(
            doc, detail_pno, detail_y0, p_cur, band_y0, band_y1
        )
        if c:
            clips.append(c)

        if band_y1 < hard - 0.5:
            break
        p_cur += 1

    return clips


def passenger_detail_clip_segments(
    doc: fitz.Document,
    detail_pno: int,
    detail_y0: float,
    passenger_tokens: Set[str],
    all_passenger_token_sets: List[Set[str]],
) -> List[Tuple[int, fitz.Rect]]:
    """
    目標旅客在詳細訊息區的「每一筆」姓名錨點各產生一段或多段 clip；
    每段垂直範圍為 Y 中點分群列帶；全頁寬 clip，貼上時 x=0（paste_segments_vertical）。
    """
    lines = collect_detail_lines_ordered(doc, detail_pno, detail_y0)
    if not lines:
        return []

    if not any(
        tokens_match_passenger(passenger_tokens, pts) for pts in all_passenger_token_sets
    ):
        return []

    out: List[Tuple[int, fitz.Rect]] = []
    for i, (_pno, _bb, _txt, nt) in enumerate(lines):
        if nt is None or not tokens_match_passenger(nt, passenger_tokens):
            continue
        out.extend(
            _clips_row_based_for_detail_anchor(
                doc,
                detail_pno,
                detail_y0,
                lines,
                i,
                all_passenger_token_sets,
                passenger_tokens,
            )
        )
    return out


# ═══════════════════════════════════════════════════════════════════════════
# 產出驗證（自我監督）
# ═══════════════════════════════════════════════════════════════════════════

def _token_as_word_pattern(tok: str) -> str:
    return rf"(?<![A-Za-z]){re.escape(tok)}(?![A-Za-z])"


def verify_output_pdf(
    pdf_path: Path,
    target_passenger_name: str,
    *,
    target_tokens: Set[str],
    all_passenger_token_sets: List[Set[str]],
    top_band_y: float = 200.0,
    max_names_in_top_band: int = 3,
) -> None:
    """
    斷言 1：全文不得同時出現「其他旅客」之姓名 token（皆命中視為夾帶）。
    斷言 2：第一頁 y<top_band_y 內，稱謂＋姓名樣式列所對應之獨立人數不得超過 max。
    target_passenger_name 僅供錯誤訊息顯示。
    """
    d = fitz.open(pdf_path)
    try:
        full_text = "".join(p.get_text() for p in d)
        lines_all = full_text.splitlines()
        others = [
            s
            for s in all_passenger_token_sets
            if not tokens_match_passenger(s, target_tokens)
        ]
        # 航班區可佔大量行數；僅掃描最末約 12% 行（多為詳細訊息貼上區）做夾帶檢查
        tail_start = max(0, int(len(lines_all) * 0.88))
        detail_tail_lines = lines_all[tail_start:]
        for o in others:
            if len(o) < 2:
                continue
            # 與目標共用 token（同姓多欄列）時不檢：避免 KKDAY 併列姓名誤判；仍會抓完全無交集之他人列
            if not target_tokens.isdisjoint(o):
                continue
            for line in detail_tail_lines:
                lu = line.upper()
                if all(
                    re.search(_token_as_word_pattern(t), lu, re.I) for t in o
                ):
                    raise ValueError(
                        f"夾帶其他旅客! 目標={target_passenger_name!r} 偵測到他人 token={sorted(o)} 於列: {line[:80]!r}"
                    )

        if len(d) < 1:
            return
        p0 = d[0]
        dct = p0.get_text("dict")
        top_name_keys: Set[Tuple[str, ...]] = set()
        for b in dct.get("blocks", []):
            if b.get("type") != 0:
                continue
            for ln in b.get("lines", []):
                bbox = fitz.Rect(ln["bbox"])
                if bbox.y0 >= top_band_y:
                    continue
                txt = "".join(s.get("text", "") for s in ln.get("spans", []))
                for m in RE_DETAILS_NAME.finditer(txt):
                    body = normalize_spaces(m.group(2))
                    toks = normalize_name_tokens(body)
                    if len(toks) >= 2:
                        top_name_keys.add(tuple(sorted(toks)))
        if len(top_name_keys) > max_names_in_top_band:
            raise ValueError(
                f"未成功挖除總名單! 頁首 y<{top_band_y} 內辨識到 {len(top_name_keys)} 組姓名: {target_passenger_name!r}"
            )
    finally:
        d.close()


def paste_segments_vertical(
    out_doc: fitz.Document,
    src_doc: fitz.Document,
    segs: List[Tuple[int, fitz.Rect]],
    start_yc: float,
    is_detail: bool = False,
) -> float:
    """
    將多段 clip 垂直貼上；tgt x 置於 0。
    若為詳細訊息區 (is_detail=True)，主動塗白殘留的「去程」、「回程」等標籤。
    """
    yc = start_yc
    page = out_doc[-1]

    for pno, clip in segs:
        w, h = clip.width, clip.height
        if yc + h > PAGE_SAFE_Y and yc > PAGE_TOP_Y - 0.1:
            out_doc.new_page(width=A4_W, height=A4_H)
            page = out_doc[-1]
            yc = PAGE_TOP_Y

        tgt = fitz.Rect(0, yc, w, yc + h)
        page.show_pdf_page(tgt, src_doc, pno, clip=clip)

        if is_detail:
            src_page = src_doc[pno]
            for kw in ("去程", "回程", "詳細訊息"):
                for inst in src_page.search_for(kw):
                    if inst.y1 > clip.y0 and inst.y0 < clip.y1:
                        rx0 = inst.x0 - 2
                        ry0 = yc + (max(inst.y0, clip.y0) - clip.y0) - 2
                        rx1 = inst.x1 + 2
                        ry1 = yc + (min(inst.y1, clip.y1) - clip.y0) + 2
                        page.draw_rect(
                            fitz.Rect(rx0, ry0, rx1, ry1),
                            color=(1, 1, 1),
                            fill=(1, 1, 1),
                            width=0,
                        )

        yc += h + CLIP_GAP
    return yc


# ═══════════════════════════════════════════════════════════════════════════
# 主流程
# ═══════════════════════════════════════════════════════════════════════════

def split_one_pdf(
    pdf_path: Path, out_dir: Path, *, run_verify: bool = True
) -> int:
    doc = fitz.open(pdf_path)
    stem_prefix = pdf_path.stem
    try:
        hit_pnr = find_first_line_global(doc, MARKER_PNR_LABEL, casefold=False)
        if not hit_pnr:
            raise RuntimeError(f"找不到「{MARKER_PNR_LABEL}」。")
        pnr_pno, pnr_bbox, _ = hit_pnr

        hit_flight = find_first_line_global(doc, MARKER_FLIGHT, casefold=False)
        if not hit_flight:
            raise RuntimeError(f"找不到「{MARKER_FLIGHT}」。")
        flight_pno, flight_bbox, _ = hit_flight

        hit_detail = find_first_line_global(doc, MARKER_DETAIL, casefold=False)
        if not hit_detail:
            raise RuntimeError(f"找不到「{MARKER_DETAIL}」。")
        detail_pno, detail_bbox, _ = hit_detail

        if (flight_pno, flight_bbox.y0) < (pnr_pno, pnr_bbox.y0):
            raise RuntimeError("版型異常：航班資訊早於訂位代號區。")
        if (detail_pno, detail_bbox.y0) < (flight_pno, flight_bbox.y0):
            raise RuntimeError("版型異常：詳細訊息早於航班資訊。")

        passengers_list = parse_passengers_from_list(
            doc, pnr_pno, pnr_bbox.y0, flight_pno, flight_bbox.y0
        )
        if passengers_list:
            passengers = [(d, t, "list") for d, t in passengers_list]
        else:
            passengers_details = parse_passengers_from_details_table(
                doc, detail_pno, detail_bbox.y0
            )
            if not passengers_details:
                raise RuntimeError(
                    "無法從頂部名單或「詳細訊息」區解析任何旅客姓名（預期含 MR/MS… 之英文姓名）。"
                )
            passengers = [(d, t, "details") for d, t in passengers_details]

        all_sets = [s for _, s, _ in passengers]

        top_segs = build_pnr_top_clip(doc, pnr_pno, pnr_bbox)
        flight_segs = collect_flight_info_segments(
            doc, flight_pno, flight_bbox.y0, detail_pno, detail_bbox.y0
        )

        out_dir.mkdir(parents=True, exist_ok=True)
        n_out = 0

        for display_line, ptoks, src in passengers:
            fname = passenger_filename_stem(display_line, src)
            out_name = f"{stem_prefix}_{fname}.pdf" if stem_prefix else f"{fname}.pdf"
            out_path = out_dir / out_name

            detail_segs = passenger_detail_clip_segments(
                doc, detail_pno, detail_bbox.y0, ptoks, all_sets
            )

            out = fitz.open()
            out.new_page(width=A4_W, height=A4_H)
            cur = out[-1]
            yc = 0.0

            # Clip 1：Top block（1:1 寬高，禁止拉寬至 A4_W）
            for tpno, tclip in top_segs:
                w1, h1 = tclip.width, tclip.height
                if yc + h1 > PAGE_SAFE_Y and yc > PAGE_TOP_Y - 0.1:
                    out.new_page(width=A4_W, height=A4_H)
                    cur = out[-1]
                    yc = PAGE_TOP_Y
                cur.show_pdf_page(fitz.Rect(0, yc, w1, yc + h1), doc, tpno, clip=tclip)
                if tpno == pnr_pno:
                    apply_top_block_white_masks(cur, doc[tpno], yc, tclip)
                yc += h1 + CLIP_GAP

            # Clip 2：航班資訊
            yc = paste_segments_vertical(out, doc, flight_segs, yc, is_detail=False)

            # Clip 3：詳細訊息（該旅客）
            if detail_segs:
                yc = paste_segments_vertical(out, doc, detail_segs, yc, is_detail=True)
            else:
                # 仍產檔但註記：表格內未找到該姓名 token
                pass

            out.save(out_path)
            out.close()
            if run_verify:
                verify_output_pdf(
                    out_path,
                    display_line,
                    target_tokens=ptoks,
                    all_passenger_token_sets=all_sets,
                )
            n_out += 1

        return n_out
    finally:
        doc.close()


def main() -> None:
    ap = argparse.ArgumentParser(
        description="台灣虎航團體行程表依旅客拆分 PDF（Top+航班+詳細訊息，遮蔽 PNR 右側）"
    )
    ap.add_argument(
        "pdf",
        type=Path,
        nargs="+",
        help="一個或多個來源 PDF",
    )
    ap.add_argument(
        "-o",
        "--out",
        type=Path,
        default=None,
        help="輸出目錄（單檔時預設 <stem>_split；多檔時預設目前目錄下各 stem_split）",
    )
    ap.add_argument(
        "--no-verify",
        action="store_true",
        help="略過產出 PDF 之 verify_output_pdf 檢查",
    )
    args = ap.parse_args()
    paths: List[Path] = args.pdf

    total = 0
    if len(paths) == 1:
        p = paths[0]
        out_dir = args.out if args.out else p.with_suffix("").with_name(p.stem + "_split")
        n = split_one_pdf(p, out_dir, run_verify=not args.no_verify)
        total += n
        print(f"完成。共輸出 {n} 個檔案至：{out_dir}")
    else:
        for p in paths:
            out_dir = (
                args.out / p.stem
                if args.out
                else p.with_suffix("").with_name(p.stem + "_split")
            )
            n = split_one_pdf(p, out_dir, run_verify=not args.no_verify)
            total += n
            print(f"  {p.name} → {n} 個檔案 → {out_dir}")
        print(f"全部完成，共 {total} 個檔案。")


if __name__ == "__main__":
    main()
