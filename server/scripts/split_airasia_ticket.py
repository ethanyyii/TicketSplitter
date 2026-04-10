#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
AirAsia 團體行程表 PDF 拆解工具 — 還原「先航班、後行李」閱讀順序

  Page 1：僅 Header + Flight Summary「去程段」（不含任何 Add-ons）。
  Page 2 起：yc=40 → 回程 FS → 去程 Add-ons 標題與內容 → 回程 Add-ons 標題與內容；
            區塊超出底部安全線時自動換頁（標題帶 orphan 緩衝防孤行）。

Flight Summary 在「Return: …」日期行前 10px 切分去程／回程。Booking no. 白塊遮蔽。
"""

import argparse
import math
import re
from pathlib import Path
from typing import List, Optional, Tuple

import fitz  # PyMuPDF

# ── 錨點文字 ─────────────────────────────────────────────────────────────────
MARKER_GUEST_DETAILS  = "GUEST DETAILS"
MARKER_FLIGHT_SUMMARY = "FLIGHT SUMMARY"
MARKER_ADD_ONS        = "ADD-ONS"

# ── 頁尾截止標記（防止幽靈文字）
FOOTER_MARKERS = ("VIEW TERMS", "DOWNLOAD THE BEST", "GO FURTHER")

# ── 排版常數 ─────────────────────────────────────────────────────────────────
# Header 截止 y 由 find_header_y1() 動態計算（Booking Confirmed / Booking date）
FS_PAD       = 35.0   # Flight summary 上下留白（含頂部灰框）
SECTION_GAP  = 20.0   # Header 與去程 FS 之間
ADDON_GAP    = 30.0   # 區塊間可選間距參考
MARGIN_LEFT  = 36.0   # Add-ons 統一左邊距

A4_W           = 595.28
A4_H           = 841.89
PAGE_SAFE_Y    = 800.0   # 換頁判斷：yc + 高度 不可超過此底線
PAGE_TOP_Y     = 40.0    # 新頁／第二段起筆 y
HDR_ORPHAN_BUF = 80.0    # 路由標題換頁時額外緩衝（避免標題與行李分頁）

# ── RE ───────────────────────────────────────────────────────────────────────
GUEST_NAME_LINE_RE = re.compile(r"^(.+?)\s*\((?:adult|child|infant)\)\s*$", re.I)
PAREN_ROLE_RE      = re.compile(r"\s*\((?:adult|child|infant)\)\s*", re.I)
SEAT_RE            = re.compile(r"^Seat\s*[:\s]*[A-Z0-9]+$", re.I)
# 回程詳細日期標題（Flight Summary 內切分錨點）
RETURN_DATE_LINE_RE = re.compile(
    r"Return:\s+[A-Za-z]+day,\s+\d+", re.I)

ADDON_NAME_BAD = (
    "CHECKED", "CARRY", "SEAT", "BAGGAGE", "ADD-ONS",
    "MAC", "CHEESE", "OMELETTE", "STUFFED", "NON-STOP",
    "ECONOMY", "GMT", "DEPART", "RETURN", "AIRASIA", "THAI",
    "FLIGHT", "SUMMARY", "TERMINAL", "AIRPORT", "HTTP",
    "VIEW TERMS", "DOWNLOAD", "KIX", "TPE", "NRT", "KHH",
)


# ════════════════════════════════════════════════════════════════════════════
# 基礎工具
# ════════════════════════════════════════════════════════════════════════════

def normalize_spaces(s: str) -> str:
    return re.sub(r"\s+", " ", (s or "")).strip()


def passenger_name_only(full: str) -> str:
    s = normalize_spaces(full)
    s = re.sub(r"^(MR|MS|MRS|MISS|MSTR)\s+", "", s, flags=re.I).strip()
    s = PAREN_ROLE_RE.sub(" ", s)
    return normalize_spaces(s)


def safe_filename_component(s: str) -> str:
    s = normalize_spaces(s).upper()
    s = re.sub(r"[^A-Z0-9]+", "_", s).strip("_")
    return s[:120] if len(s) > 120 else s


def find_header_y1(page: fitz.Page) -> float:
    """
    動態 Header 底線：掃描 page 0，找「Booking Confirmed」或「Booking date」
    任一行的 y1，取較深者（涵蓋兩行皆存在時），再 +5px，緊貼裁掉下方灰線與留白。
    """
    best = 100.0
    for bbox, txt in page_lines(page):
        if bbox.y0 > 200:
            break
        u = txt.upper()
        if "BOOKING CONFIRMED" in u or "BOOKING DATE" in u:
            best = max(best, bbox.y1)
    return best + 5.0


def _finite_rect(r: fitz.Rect) -> bool:
    return all(math.isfinite(v) for v in (r.x0, r.y0, r.x1, r.y1))


def _ensure_valid_clip(r: fitz.Rect, page_rect: fitz.Rect, min_h: float = 5.0) -> fitz.Rect:
    rr = fitz.Rect(
        max(page_rect.x0, r.x0), max(page_rect.y0, r.y0),
        min(page_rect.x1, r.x1), min(page_rect.y1, r.y1),
    )
    if (not _finite_rect(rr)) or rr.is_empty or rr.width < 5 or rr.height < min_h:
        raise ValueError(f"Invalid clip: {rr}")
    return rr


def page_lines(page: fitz.Page) -> List[Tuple[fitz.Rect, str]]:
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


def find_first_line_global(
    doc: fitz.Document, needle_upper: str
) -> Optional[Tuple[int, fitz.Rect, str]]:
    for pno in range(len(doc)):
        for bbox, txt in page_lines(doc[pno]):
            if needle_upper in normalize_spaces(txt).upper():
                return pno, bbox, txt
    return None


# ════════════════════════════════════════════════════════════════════════════
# 錨點搜尋
# ════════════════════════════════════════════════════════════════════════════

def find_first_add_ons_after(
    doc: fitz.Document, min_pno: int, min_y0: float
) -> Tuple[int, fitz.Rect, str]:
    for pno in range(min_pno, len(doc)):
        for bbox, txt in page_lines(doc[pno]):
            if pno == min_pno and bbox.y0 <= min_y0 + 1:
                continue
            if MARKER_ADD_ONS in normalize_spaces(txt).upper():
                return pno, bbox, txt
    raise RuntimeError("找不到「Add-ons」區塊。")


# ════════════════════════════════════════════════════════════════════════════
# Part 2：Flight Summary 片段（跨頁：動態邊界 + 白邊上限）
# ════════════════════════════════════════════════════════════════════════════

# 無文字錨點時的後備白邊；有文字時以 bbox ±5px 為準，並與此上限取捨
_PAGE_MARGIN = 35.0


def collect_flight_summary_segs(
    doc: fitz.Document,
    fs_pno: int, fs_y0: float,
    ad_pno: int, ad_y0: float,
    header_y1: float,
    guest_bottom_y1: float,
) -> List[Tuple[int, fitz.Rect]]:
    """
    從「Flight summary」上方 FS_PAD px 起，至「Add-ons」上方 FS_PAD px 止，
    收集 Flight Summary 片段。跨頁時主頁底／次頁頂／中間頁上下界改為：
    掃描 FS 帶內文字 bbox，底端至少含 max(y1)+5、頂端至少從 min(y0)-5 起，
    並與「頁高 ± _PAGE_MARGIN」取 max/min，避免水平腰斬文字。不使用 get_drawings()。

      · 主 FS 頁頂部：fs_y0 - FS_PAD + guest／header 防護
      · 主 FS 頁底部：max(ph - _PAGE_MARGIN, 最底行 y1 + 5)，上限 ph
      · 中間頁：上下界同樣依該頁 FS 文字動態調整
      · 尾頁 sliver 頂：min(_PAGE_MARGIN, 最頂行 y0 - 5)（無文字則 _PAGE_MARGIN）
      · 尾頁 sliver 底：ad_y0 - FS_PAD
    """
    segs: List[Tuple[int, fitz.Rect]] = []

    # 主 FS 頁頂部：往上多留 FS_PAD，確保頂部水平框線完整
    y0_start = max(0.0, fs_y0 - FS_PAD)
    # 防護：FS 與 Header 同頁（page 0）時，避免往上蓋到 Header
    if fs_pno == 0:
        y0_start = max(y0_start, header_y1)
    y0_start = max(y0_start, guest_bottom_y1 + 5.0)

    if fs_pno == ad_pno:
        # 單頁情形：直接從 fs 頂到 add-ons 前
        pr = doc[fs_pno].rect
        y1 = max(y0_start + 60.0, ad_y0 - FS_PAD)
        segs.append((fs_pno, _ensure_valid_clip(
            fitz.Rect(0, y0_start, pr.width, y1), pr, min_h=40)))
    else:
        # ① 主 FS 頁：底部不低於 FS 最底行 y1+5，並至少保留 _PAGE_MARGIN 白邊意圖
        prf = doc[fs_pno].rect
        _, main_bot = _fs_text_y_extents_in_band(
            doc, fs_pno, fs_pno, fs_y0, ad_pno, ad_y0, header_y1, guest_bottom_y1)
        if main_bot is not None:
            clip_y1 = min(
                prf.height,
                max(prf.height - _PAGE_MARGIN, main_bot + 5.0),
            )
        else:
            clip_y1 = prf.height - _PAGE_MARGIN
        segs.append((fs_pno, _ensure_valid_clip(
            fitz.Rect(0, y0_start, prf.width, clip_y1), prf, min_h=40)))

        # ② 中間整頁（若有）：上下界依該頁 FS 文字動態調整
        for pno in range(fs_pno + 1, ad_pno):
            pm = doc[pno].rect
            mid_top, mid_bot = _fs_text_y_extents_in_band(
                doc, pno, fs_pno, fs_y0, ad_pno, ad_y0, header_y1, guest_bottom_y1)
            seg_y0 = _PAGE_MARGIN
            if mid_top is not None:
                seg_y0 = min(_PAGE_MARGIN, max(0.0, mid_top - 5.0))
            seg_y1 = pm.height - _PAGE_MARGIN
            if mid_bot is not None:
                seg_y1 = min(
                    pm.height,
                    max(pm.height - _PAGE_MARGIN, mid_bot + 5.0),
                )
            segs.append((pno, _ensure_valid_clip(
                fitz.Rect(0, seg_y0, pm.width, seg_y1), pm, min_h=20)))

        # ③ 尾頁 sliver：頂部不裁穿 FS 首行；底部到 Add-ons 前
        y1_ad = max(5.0, ad_y0 - FS_PAD)
        sliver_top = _PAGE_MARGIN
        sliver_top_scan, _ = _fs_text_y_extents_in_band(
            doc, ad_pno, fs_pno, fs_y0, ad_pno, ad_y0, header_y1, guest_bottom_y1)
        if sliver_top_scan is not None:
            sliver_top = min(_PAGE_MARGIN, max(0.0, sliver_top_scan - 5.0))
        if sliver_top < y1_ad:           # 有實際內容（非空白）
            pr_ad = doc[ad_pno].rect
            try:
                segs.append((ad_pno, _ensure_valid_clip(
                    fitz.Rect(0, sliver_top, pr_ad.width, y1_ad),
                    pr_ad, min_h=5)))
            except ValueError:
                pass
        # 若 sliver_top >= y1_ad（OSA：35 > 6.7）→ 跳過，不產生空白片段

    return segs


def _fs_line_top_y(
    fs_pno: int,
    fs_y0: float,
    ad_pno: int,
    ad_y0: float,
    header_y1: float,
    guest_bottom_y1: float,
) -> float:
    """Flight Summary 區在 fs_pno 上的內容頂緣（與 collect 一致）。"""
    y0_start = max(0.0, fs_y0 - FS_PAD)
    if fs_pno == 0:
        y0_start = max(y0_start, header_y1)
    y0_start = max(y0_start, guest_bottom_y1 + 5.0)
    return y0_start


def line_in_flight_summary_band(
    pno: int,
    bbox: fitz.Rect,
    fs_pno: int,
    fs_y0: float,
    ad_pno: int,
    ad_y0: float,
    header_y1: float,
    guest_bottom_y1: float,
) -> bool:
    """文字行是否落在 Flight Summary 垂直範圍內（不含 Add-ons）。"""
    y = bbox.y0
    if pno < fs_pno or pno > ad_pno:
        return False
    if pno == fs_pno:
        top = _fs_line_top_y(
            fs_pno, fs_y0, ad_pno, ad_y0, header_y1, guest_bottom_y1)
        if y < top - 1:
            return False
    if pno == ad_pno:
        if y >= ad_y0 - FS_PAD - 1:
            return False
    return True


def _fs_text_y_extents_in_band(
    doc: fitz.Document,
    pno: int,
    fs_pno: int,
    fs_y0: float,
    ad_pno: int,
    ad_y0: float,
    header_y1: float,
    guest_bottom_y1: float,
) -> Tuple[Optional[float], Optional[float]]:
    """
    掃描該頁落在 Flight Summary 垂直帶內的所有文字行，
    回傳 (最小 y0, 最大 y1)；若無任何行則 (None, None)。
    """
    min_y0: Optional[float] = None
    max_y1: Optional[float] = None
    for bbox, _txt in page_lines(doc[pno]):
        if not line_in_flight_summary_band(
            pno,
            bbox,
            fs_pno,
            fs_y0,
            ad_pno,
            ad_y0,
            header_y1,
            guest_bottom_y1,
        ):
            continue
        min_y0 = bbox.y0 if min_y0 is None else min(min_y0, bbox.y0)
        max_y1 = bbox.y1 if max_y1 is None else max(max_y1, bbox.y1)
    return min_y0, max_y1


def find_return_cut_y(
    doc: fitz.Document,
    fs_pno: int,
    fs_y0: float,
    ad_pno: int,
    ad_y0: float,
    header_y1: float,
    guest_bottom_y1: float,
) -> Tuple[int, float]:
    """
    在 Flight Summary 內尋找「Return: Sunday, 5 April 2026」類標題，
    回傳 (頁碼, CUT_Y)；CUT_Y = 該行 y0 - 10（保留上方灰底邊距）。
    """
    for pno in range(fs_pno, ad_pno + 1):
        pr = doc[pno].rect
        for bbox, txt in page_lines(doc[pno]):
            if not line_in_flight_summary_band(
                pno,
                bbox,
                fs_pno,
                fs_y0,
                ad_pno,
                ad_y0,
                header_y1,
                guest_bottom_y1,
            ):
                continue
            if RETURN_DATE_LINE_RE.search(txt):
                cut_y = max(pr.y0, bbox.y0 - 10.0)
                return pno, cut_y
    raise RuntimeError(
        "找不到 Flight Summary 內的回程日期標題（預期如 Return: Sunday, 5 April 2026）。"
    )


def split_fs_segs_outbound_return(
    doc: fitz.Document,
    fs_segs: List[Tuple[int, fitz.Rect]],
    cut_pno: int,
    cut_y: float,
) -> Tuple[List[Tuple[int, fitz.Rect]], List[Tuple[int, fitz.Rect]]]:
    """
    依 CUT_Y 將 fs_segs 分為去程段（頂端至 cut_y）與回程段（cut_y 至底端）。
    cut_y 為來源頁面座標；僅在 cut_pno 上對該頁片段做水平切開。
    """
    out_segs: List[Tuple[int, fitz.Rect]] = []
    ret_segs: List[Tuple[int, fitz.Rect]] = []

    for pno, r in fs_segs:
        pr = doc[pno].rect
        if pno < cut_pno:
            out_segs.append((pno, r))
        elif pno > cut_pno:
            ret_segs.append((pno, r))
        else:
            if cut_y > r.y0 + 0.5:
                try:
                    out_segs.append(
                        (
                            pno,
                            _ensure_valid_clip(
                                fitz.Rect(r.x0, r.y0, r.x1, min(r.y1, cut_y)),
                                pr,
                                min_h=5.0,
                            ),
                        )
                    )
                except ValueError:
                    pass
            if cut_y < r.y1 - 0.5:
                try:
                    ret_segs.append(
                        (
                            pno,
                            _ensure_valid_clip(
                                fitz.Rect(r.x0, max(r.y0, cut_y), r.x1, r.y1),
                                pr,
                                min_h=5.0,
                            ),
                        )
                    )
                except ValueError:
                    pass

    if not out_segs:
        raise RuntimeError("Flight Summary 去程段裁切結果為空。")
    if not ret_segs:
        raise RuntimeError("Flight Summary 回程段裁切結果為空。")
    return out_segs, ret_segs


# ════════════════════════════════════════════════════════════════════════════
# 路由標題搜尋（Add-ons 段落分隔）
# ════════════════════════════════════════════════════════════════════════════

def is_route_segment_header(txt: str) -> bool:
    t = normalize_spaces(txt)
    if len(t) < 14:
        return False
    u = t.upper()
    if " TO " not in u:
        return False
    if "HTTP" in u or "ENTRY REQUIREMENT" in u or "FLIGHT TO:" in u:
        return False
    if "TERMINAL" in u and "AIRPORT" in u:
        return False
    if not re.search(r"[A-Z][A-Z\- ]+\s+TO\s+[A-Z]", u):
        return False
    return True


def find_route_headers_after_add_ons(
    doc: fitz.Document,
    ad_pno: int, ad_y0: float,
    stop_pno: int, stop_y0: float,
) -> List[Tuple[int, fitz.Rect, str]]:
    collected: List[Tuple[int, fitz.Rect, str]] = []
    seen_y: List[Tuple[int, float]] = []

    def past_stop(pno: int, bbox: fitz.Rect) -> bool:
        return pno > stop_pno or (pno == stop_pno and bbox.y0 >= stop_y0 - 2)

    def push(pno: int, bbox: fitz.Rect, txt: str) -> None:
        for sp, sy in seen_y:
            if sp == pno and abs(sy - bbox.y0) < 2:
                return
        seen_y.append((pno, bbox.y0))
        collected.append((pno, bbox, txt))

    for pno in range(ad_pno, len(doc)):
        for bbox, txt in page_lines(doc[pno]):
            if past_stop(pno, bbox):
                return collected[:2]
            if pno == ad_pno and bbox.y0 < ad_y0 - 1:
                continue
            if is_route_segment_header(txt):
                push(pno, bbox, txt)
            if len(collected) >= 2:
                return collected[:2]
    return collected


# ════════════════════════════════════════════════════════════════════════════
# 旅客清單
# ════════════════════════════════════════════════════════════════════════════

def in_guest_band(
    pno: int, line_y0: float,
    g_pno: int, g_y0: float,
    fs_pno: int, fs_y0: float,
) -> bool:
    if pno < g_pno or pno > fs_pno:
        return False
    if pno == g_pno and line_y0 < g_y0 - 2:
        return False
    if pno == fs_pno and line_y0 >= fs_y0 - 2:
        return False
    return True


def guest_details_bottom_y1_on_fs_page(
    doc: fitz.Document,
    fs_pno: int,
    g_pno: int,
    g_y0: float,
    fs_y0: float,
) -> float:
    """
    在「Flight summary」所在頁，掃描 Guest details 垂直帶內所有文字行的最大 y1。
    若該頁無名單行，回傳 -inf（後續 max 不影響 y0_start）。
    """
    best = float("-inf")
    for bbox, _txt in page_lines(doc[fs_pno]):
        if not in_guest_band(fs_pno, bbox.y0, g_pno, g_y0, fs_pno, fs_y0):
            continue
        best = max(best, bbox.y1)
    return best


def extract_guest_keys(
    doc: fitz.Document,
    g_pno: int, g_y0: float,
    fs_pno: int, fs_y0: float,
) -> List[str]:
    keys: List[str] = []
    seen: set = set()
    for pno in range(len(doc)):
        for bbox, txt in page_lines(doc[pno]):
            if not in_guest_band(pno, bbox.y0, g_pno, g_y0, fs_pno, fs_y0):
                continue
            m = GUEST_NAME_LINE_RE.match(txt.strip())
            if not m:
                continue
            key = passenger_name_only(m.group(1).strip()).upper()
            if key and key not in seen:
                seen.add(key)
                keys.append(key)
    if not keys:
        raise RuntimeError("無法從 Guest details 區辨識旅客（預期為「姓名 (adult)」格式）。")
    return keys


# ════════════════════════════════════════════════════════════════════════════
# Part 3：Add-ons 雙欄裁切
# ════════════════════════════════════════════════════════════════════════════

def column_side(page: fitz.Page, bbox: fitz.Rect) -> str:
    return "left" if (bbox.x0 + bbox.x1) / 2.0 < page.rect.width / 2.0 else "right"


def column_x_bounds(page: fitz.Page, side: str) -> Tuple[float, float]:
    w, mid = page.rect.width, page.rect.width / 2.0
    if side == "left":
        return 0.0, max(mid - 1.0, 5.0)
    return min(mid + 1.0, w - 5.0), w


def is_likely_addon_name(txt: str) -> bool:
    t = normalize_spaces(txt)
    if len(t) < 4:
        return False
    u = t.upper()
    for bad in ADDON_NAME_BAD:
        if bad in u:
            return False
    if any(c in t for c in ("(", ")", ":", ",", ";")):
        return False
    if re.search(r"\d", t):
        return False
    tokens = u.split()
    if len(tokens) < 2 or len(tokens) > 6:
        return False
    return all(re.match(r"^[A-Z][A-Z0-9\-']*$", tok) for tok in tokens)


def extract_addon_names_page(page: fitz.Page) -> List[Tuple[str, fitz.Rect]]:
    res = [
        (normalize_spaces(txt), bbox)
        for bbox, txt in page_lines(page)
        if is_likely_addon_name(txt)
    ]
    res.sort(key=lambda x: x[1].y0)
    return res


def lowest_seat_bottom_in_band(
    page: fitz.Page, y0: float, y1: float, x0: float, x1: float
) -> Optional[float]:
    best: Optional[float] = None
    for bbox, txt in page_lines(page):
        if bbox.y0 < y0 - 3 or bbox.y1 > y1 + 3:
            continue
        if bbox.x0 < x0 - 8 or bbox.x1 > x1 + 8:
            continue
        if SEAT_RE.match(normalize_spaces(txt)):
            best = bbox.y1 if best is None else max(best, bbox.y1)
    return best


def find_footer_y_on_page(page: fitz.Page, y_start: float, y_hard: float) -> float:
    """
    在 y=[y_start, y_hard] 範圍內尋找頁尾標記文字（View terms、Download 等）。
    若找到，回傳該行的 y0（作為 Add-ons 的硬性上界）；否則回傳 y_hard。
    這可防止旅客 clip 的底部延伸到 "View terms" / 廣告等頁尾內容。
    """
    for bbox, txt in page_lines(page):
        if bbox.y0 <= y_start or bbox.y0 >= y_hard:
            continue
        t = normalize_spaces(txt).upper()
        if any(m in t for m in FOOTER_MARKERS):
            return bbox.y0  # 找到頁尾，截在這裡
    return y_hard


def passenger_addon_clip(
    page: fitz.Page,
    name_bb: fitz.Rect,
    region_y0: float,
    region_y1: float,
    all_names: List[Tuple[str, fitz.Rect]],
) -> fitz.Rect:
    """
    雙欄精確裁切：
    - X：以頁寬中線分左/右欄，clip 寬度 ≈ page_width/2。
    - Y：自姓名頂端起，以「下一旅客頂端 - 2」或「Seat 行底端 + 18」收斂。
         若兩者均無效，使用名字 y1 + 160（保守值）作為 fallback。
         最後以 find_footer_y_on_page 防止越界進入頁尾。
    回傳的 clip.x0 可能為 0（左欄）或 mid（右欄）；
    呼叫端以 target.x0=MARGIN_LEFT 渲染，右欄自動平移至左側靠齊。
    """
    pr   = page.rect
    side = column_side(page, name_bb)
    x0, x1 = column_x_bounds(page, side)

    col_boxes = sorted(
        [bb for _nm, bb in all_names
         if bb.y0 < region_y1 - 1 and bb.y1 > region_y0 + 1
            and column_side(page, bb) == side],
        key=lambda b: b.y0,
    )

    y_start = max(region_y0, name_bb.y0 - 4)
    prev_bottoms = [b.y1 for b in col_boxes if b.y0 < name_bb.y0 - 0.5]
    if prev_bottoms:
        y_start = max(y_start, max(prev_bottoms) + 2)

    # 以同欄下一旅客為硬性底部（最精確邊界）
    hard_bottom = region_y1
    next_tops = [b.y0 for b in col_boxes if b.y0 > name_bb.y0 + 2]
    if next_tops:
        hard_bottom = min(hard_bottom, min(next_tops) - 2)

    # 以「View terms」等頁尾標記進一步收緊底部（防幽靈文字）
    hard_bottom = find_footer_y_on_page(page, y_start, hard_bottom)

    seat_bot = lowest_seat_bottom_in_band(page, y_start, hard_bottom, x0, x1)
    if seat_bot:
        y_end = min(hard_bottom, seat_bot + 18)
    else:
        # 無座位資訊時保守估計：名字高度 + 160px（≈8行附加資訊）
        y_end = min(hard_bottom, name_bb.y1 + 160)
    if y_end <= y_start + 12:
        y_end = min(hard_bottom, y_start + 130)

    return _ensure_valid_clip(fitz.Rect(x0, y_start, x1, y_end), pr, min_h=18.0)


def pick_addon_occ_in_band(
    doc: fitz.Document,
    key_upper: str,
    p0: int, y0: float,
    p1: int, y1: float,
) -> Optional[Tuple[int, fitz.Rect]]:
    best: Optional[Tuple[int, fitz.Rect]] = None
    for pno in range(p0, p1 + 1):
        for nm, bb in extract_addon_names_page(doc[pno]):
            if passenger_name_only(nm).upper() != key_upper:
                continue
            if pno == p0 and bb.y0 < y0 - 2:
                continue
            if pno == p1 and bb.y0 >= y1 - 2:
                continue
            cand = (pno, bb)
            if best is None or (pno, bb.y0) < (best[0], best[1].y0):
                best = cand
    return best


def _addon_last_in_column_on_page(
    page: fitz.Page,
    name_bb: fitz.Rect,
    side: str,
    region_y0: float,
    region_y1: float,
    all_names: List[Tuple[str, fitz.Rect]],
) -> bool:
    """該欄在 name_bb 下方是否已無其他旅客姓名（可視為本頁該欄最後一人）。"""
    col_boxes = sorted(
        [
            bb
            for _nm, bb in all_names
            if bb.y0 < region_y1 - 1
            and bb.y1 > region_y0 + 1
            and column_side(page, bb) == side
        ],
        key=lambda b: b.y0,
    )
    next_tops = [b.y0 for b in col_boxes if b.y0 > name_bb.y0 + 2]
    return len(next_tops) == 0


def _build_spillover_addon_clip(
    page: fitz.Page,
    side: str,
    key_upper: str,
    pno: int,
    p0: int,
    y0: float,
    p1: int,
    y1: float,
) -> Optional[fitz.Rect]:
    """
    次頁同欄、航段帶內：自帶頂至「下一旅客姓名頂」或頁尾（含 find_footer 防護），
    若有 Seat／行李等文字則包成 clip；否則 None。
    """
    pr = page.rect
    x0, x1 = column_x_bounds(page, side)
    region_y0 = y0 if pno == p0 else 0.0
    region_y1 = y1 if pno == p1 else pr.height

    names_here = extract_addon_names_page(page)
    hard_bottom = region_y1
    for nm, bb in sorted(names_here, key=lambda t: t[1].y0):
        if bb.y0 < region_y0 - 1:
            continue
        if column_side(page, bb) != side:
            continue
        if passenger_name_only(nm).upper() == key_upper:
            continue
        hard_bottom = min(hard_bottom, bb.y0 - 2.0)
        break

    hard_bottom = find_footer_y_on_page(page, region_y0, hard_bottom)

    min_y0: Optional[float] = None
    max_y1: Optional[float] = None
    for bbox, txt in page_lines(page):
        if bbox.y0 < region_y0 - 2 or bbox.y0 >= hard_bottom - 1:
            continue
        cx = (bbox.x0 + bbox.x1) / 2.0
        if cx < x0 - 10 or cx > x1 + 10:
            continue
        t = normalize_spaces(txt).upper()
        if any(m in t for m in FOOTER_MARKERS):
            continue
        if is_likely_addon_name(txt):
            other = passenger_name_only(txt).upper()
            if other and other != key_upper:
                continue
        min_y0 = bbox.y0 if min_y0 is None else min(min_y0, bbox.y0)
        max_y1 = bbox.y1 if max_y1 is None else max(max_y1, bbox.y1)

    if min_y0 is None or max_y1 is None:
        return None
    pad_top, pad_bot = 4.0, 18.0
    try:
        return _ensure_valid_clip(
            fitz.Rect(x0, min_y0 - pad_top, x1, max_y1 + pad_bot),
            pr,
            min_h=12.0,
        )
    except ValueError:
        return None


def get_passenger_addon_clips(
    doc: fitz.Document,
    key_upper: str,
    p0: int,
    y0: float,
    p1: int,
    y1: float,
) -> List[Tuple[int, fitz.Rect]]:
    """
    擷取某旅客在航段帶內的 Add-ons 裁切框，支援跨頁 spillover。
    回傳 [(頁碼, clip), ...]；找不到姓名則 []。
    """
    occ = pick_addon_occ_in_band(doc, key_upper, p0, y0, p1, y1)
    if not occ:
        return []

    occ_pno, name_bb = occ
    page0 = doc[occ_pno]
    yt0 = y0 if occ_pno == p0 else 0.0
    yb0 = y1 if occ_pno == p1 else page0.rect.height
    names0 = extract_addon_names_page(page0)
    try:
        main_clip = passenger_addon_clip(page0, name_bb, yt0, yb0, names0)
    except ValueError:
        return []

    side = column_side(page0, name_bb)
    out: List[Tuple[int, fitz.Rect]] = [(occ_pno, main_clip)]

    cur = occ_pno
    while cur < p1:
        pg = doc[cur]
        yt_c = y0 if cur == p0 else 0.0
        yb_c = y1 if cur == p1 else pg.rect.height
        nms = extract_addon_names_page(pg)

        if cur == occ_pno:
            if not _addon_last_in_column_on_page(
                pg, name_bb, side, yt_c, yb_c, nms
            ):
                break
        else:
            _prev_pno, prev_clip = out[-1]
            if _prev_pno != cur or prev_clip.y1 < pg.rect.height - 55.0:
                break

        sp_pno = cur + 1
        spill = _build_spillover_addon_clip(
            doc[sp_pno], side, key_upper, sp_pno, p0, y0, p1, y1
        )
        if spill is None or spill.is_empty:
            break
        out.append((sp_pno, spill))
        cur = sp_pno

    return out


# ════════════════════════════════════════════════════════════════════════════
# 主程式
# ════════════════════════════════════════════════════════════════════════════

def main() -> None:
    ap = argparse.ArgumentParser(
        description="AirAsia 團體行程表依旅客拆分 PDF（雙頁 A4：去程／回程分頁）")
    ap.add_argument("pdf", type=Path)
    ap.add_argument("-o", "--out", type=Path, default=None)
    args = ap.parse_args()

    pdf_path: Path = args.pdf
    out_dir: Path = (
        args.out if args.out
        else pdf_path.with_suffix("").with_name(pdf_path.stem + "_split")
    )
    out_dir.mkdir(parents=True, exist_ok=True)

    doc = fitz.open(pdf_path)
    try:
        w0 = doc[0].rect.width   # ≈ 595.3（A4 寬）

        # ── 錨點 ─────────────────────────────────────────────────────────
        g_hit = find_first_line_global(doc, MARKER_GUEST_DETAILS)
        if not g_hit:
            raise RuntimeError("找不到「Guest details」。")
        g_pno, g_bbox, _ = g_hit
        g_y0 = g_bbox.y0

        fs_hit = find_first_line_global(doc, MARKER_FLIGHT_SUMMARY)
        if not fs_hit:
            raise RuntimeError("找不到「Flight summary」。")
        fs_pno, fs_bbox, _ = fs_hit
        fs_y0 = fs_bbox.y0

        if fs_pno < g_pno or (fs_pno == g_pno and fs_y0 < g_y0):
            raise RuntimeError("Flight summary 早於 Guest details，版型異常。")

        ad_pno, ad_bbox, _ = find_first_add_ons_after(doc, fs_pno, fs_y0)
        ad_y0 = ad_bbox.y0

        # Guest details 在 FS 所在頁的最底 y1：作為 FS 往上裁切的硬性天花板
        guest_bottom_y1 = guest_details_bottom_y1_on_fs_page(
            doc, fs_pno, g_pno, g_y0, fs_y0)

        # ── Header 截止 y（須先於 FS 收集，供 page0 防重疊）──────────────
        hdr_y1 = find_header_y1(doc[0])
        _p0 = doc[0]
        for _bb, _tx in page_lines(_p0):
            if _bb.y0 > 230:
                break
            if "BOOKING NO" in _tx.upper():
                _p0.draw_rect(
                    fitz.Rect(0, _bb.y0 - 3, _p0.rect.width, _bb.y1 + 22),
                    color=(1, 1, 1), fill=(1, 1, 1))
                break

        # ── 共用：Flight Summary 全段 → 再切成去程 / 回程兩段 ─────────────
        fs_segs = collect_flight_summary_segs(
            doc, fs_pno, fs_y0, ad_pno, ad_y0, hdr_y1, guest_bottom_y1)
        cut_pno, cut_y = find_return_cut_y(
            doc, fs_pno, fs_y0, ad_pno, ad_y0, hdr_y1, guest_bottom_y1)
        fs_out_segs, fs_ret_segs = split_fs_segs_outbound_return(
            doc, fs_segs, cut_pno, cut_y)

        # 路由標題搜尋上界（沿用 Flight summary 末頁底部，不再依賴頁尾錨點）
        stop_pno = len(doc) - 1
        stop_y0  = doc[stop_pno].rect.height

        routes = find_route_headers_after_add_ons(doc, ad_pno, ad_y0, stop_pno, stop_y0)
        if len(routes) < 2:
            raise RuntimeError("無法辨識去程與回程兩段路由標題。")
        out_pno, out_hdr_bb, _ = routes[0]
        ret_pno, ret_hdr_bb, _ = routes[1]

        out_band_p0, out_band_y0 = ad_pno, ad_y0
        out_band_p1, out_band_y1 = ret_pno, ret_hdr_bb.y0
        ret_band_p0, ret_band_y0 = ret_pno, ret_hdr_bb.y0
        ret_band_p1, ret_band_y1 = len(doc) - 1, doc[len(doc) - 1].rect.height

        # ── 預先建立路由標題 clip（限制 x1 = 頁寬/2 + 50，所有旅客共用）──────
        # Bug #2 防護：路由標題右側可能有 "Xh Xm Non-stop" 飛行時數文字；
        # 將 x1 限定在左半頁 + 50px，確保右側飛行時數不被截入。
        # 渲染時使用 paste_hdr（1:1，無縮放），不同於全寬的 paste_full。
        out_pg = doc[out_pno]
        out_hdr_x1 = min(out_pg.rect.width / 2 + 50, out_pg.rect.width)
        try:
            out_hdr_clip = _ensure_valid_clip(
                fitz.Rect(0, out_hdr_bb.y0 - 8, out_hdr_x1, out_hdr_bb.y1 + 10),
                out_pg.rect, min_h=8)
        except ValueError:
            out_hdr_clip = None

        ret_pg = doc[ret_pno]
        ret_hdr_x1 = min(ret_pg.rect.width / 2 + 50, ret_pg.rect.width)
        try:
            ret_hdr_clip = _ensure_valid_clip(
                fitz.Rect(0, ret_hdr_bb.y0 - 8, ret_hdr_x1, ret_hdr_bb.y1 + 10),
                ret_pg.rect, min_h=8)
        except ValueError:
            ret_hdr_clip = None

        hdr_clip = _ensure_valid_clip(
            fitz.Rect(0, 0, w0, hdr_y1), doc[0].rect, min_h=50)

        # ── 旅客清單 ─────────────────────────────────────────────────────
        keys     = extract_guest_keys(doc, g_pno, g_y0, fs_pno, fs_y0)
        stem     = pdf_path.stem
        exported = 0

        for key in sorted(keys):
            # ── 計算 Add-ons clips（每位旅客各不同，支援跨頁 spillover）────
            clips_o = get_passenger_addon_clips(
                doc, key, out_band_p0, out_band_y0, out_band_p1, out_band_y1)
            clips_r = get_passenger_addon_clips(
                doc, key, ret_band_p0, ret_band_y0, ret_band_p1, ret_band_y1)

            # ── Page 1：僅 Header + 去程 FS（不含 Add-ons）───────────────
            out = fitz.open()
            cur_page = out.new_page(width=A4_W, height=A4_H)
            yc = 0.0

            def paste_full_here(pno: int, clip: fitz.Rect, y: float) -> None:
                tgt = fitz.Rect(0, y, A4_W, y + clip.height)
                cur_page.show_pdf_page(tgt, doc, pno, clip=clip)

            paste_full_here(0, hdr_clip, yc)
            yc += hdr_clip.height + SECTION_GAP
            for pno, clip in fs_out_segs:
                paste_full_here(pno, clip, yc)
                yc += clip.height

            # ── Page 2 起：回程 FS → 去程／回程 Add-ons + 動態換頁 ────────
            out.new_page(width=A4_W, height=A4_H)
            cur_page = out[-1]
            yc = PAGE_TOP_Y

            def _maybe_new_page(clip_h: float, orphan_buf: float = 0.0) -> None:
                nonlocal cur_page, yc
                if yc + clip_h + orphan_buf > PAGE_SAFE_Y:
                    out.new_page(width=A4_W, height=A4_H)
                    cur_page = out[-1]
                    yc = PAGE_TOP_Y

            def paste_left_here(pno: int, clip: fitz.Rect, y: float) -> None:
                tgt = fitz.Rect(
                    MARGIN_LEFT, y,
                    MARGIN_LEFT + clip.width, y + clip.height)
                cur_page.show_pdf_page(tgt, doc, pno, clip=clip)

            def paste_hdr_here(pno: int, clip: fitz.Rect, y: float) -> None:
                tgt = fitz.Rect(0, y, clip.width, y + clip.height)
                cur_page.show_pdf_page(tgt, doc, pno, clip=clip)

            for pno, clip in fs_ret_segs:
                _maybe_new_page(clip.height)
                paste_full_here(pno, clip, yc)
                yc += clip.height
            yc += 20.0  # 回程 FS 與後續 Add-ons 區呼吸距離

            if out_hdr_clip:
                _maybe_new_page(out_hdr_clip.height, HDR_ORPHAN_BUF)
                paste_hdr_here(out_pno, out_hdr_clip, yc)
                yc += out_hdr_clip.height + 4.0
            if clips_o:
                for pno, clip in clips_o:
                    _maybe_new_page(clip.height)
                    paste_left_here(pno, clip, yc)
                    yc += clip.height + 2.0
                yc += 14.0

            if ret_hdr_clip:
                _maybe_new_page(ret_hdr_clip.height, HDR_ORPHAN_BUF)
                paste_hdr_here(ret_pno, ret_hdr_clip, yc)
                yc += ret_hdr_clip.height + 4.0
            if clips_r:
                for pno, clip in clips_r:
                    _maybe_new_page(clip.height)
                    paste_left_here(pno, clip, yc)
                    yc += clip.height + 2.0
                yc += 14.0

            # ── 檔名：姓氏在前，名字在後（CHENGHSUN HSIEH → HSIEH_CHENGHSUN）──
            name_parts = key.split()
            if len(name_parts) > 1:
                reordered = " ".join([name_parts[-1]] + name_parts[:-1])
            else:
                reordered = key
            out_name = f"{stem}_{safe_filename_component(reordered)}.pdf"

            out.save(out_dir / out_name)
            out.close()
            exported += 1

        print(f"完成。共輸出 {exported} 個檔案至：{out_dir}")
    finally:
        doc.close()


if __name__ == "__main__":
    main()
