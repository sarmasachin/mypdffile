from __future__ import annotations

import math
import re
from typing import Any

def _bbox_list_finite(b: list[float] | None) -> bool:
    if not b or len(b) != 4:
        return False
    try:
        return all(math.isfinite(float(x)) for x in b)
    except (TypeError, ValueError):
        return False


def _looks_like_noise_text(text: str) -> bool:
    t = (text or "").strip()
    if not t:
        return True
    if len(t) == 1 and not t.isalnum():
        return True
    # Drop OCR crumbs like isolated punctuation or border artifacts.
    if sum(ch.isalnum() for ch in t) == 0 and len(t) <= 3:
        return True
    return False


def _expand_bbox_for_readability(
    bbox: list[float],
    text: str,
    size: float,
    page_width: float,
    page_height: float,
) -> list[float]:
    x0, y0, x1, y1 = [float(v) for v in bbox]
    text_lines = [ln for ln in str(text or "").splitlines() if ln.strip()] or [str(text or "")]
    longest = max((len(ln) for ln in text_lines), default=1)
    line_count = max(1, len(text_lines))
    width = max(1.0, x1 - x0)
    height = max(1.0, y1 - y0)

    est_char_w = max(5.5, size * 0.56)
    target_width = max(width, longest * est_char_w + size * 0.9)
    target_height = max(height, line_count * max(14.0, size * 1.28) + 4.0)

    grow_x = min(page_width * 0.12, max(0.0, target_width - width))
    grow_y = min(page_height * 0.04, max(0.0, target_height - height))

    nx0 = max(0.0, x0 - min(10.0, grow_x * 0.18))
    nx1 = min(page_width, x1 + grow_x)
    ny0 = max(0.0, y0 - min(4.0, grow_y * 0.2))
    ny1 = min(page_height, y1 + grow_y)
    return [nx0, ny0, nx1, ny1]


def _cleanup_ocr_items(
    items: list[dict[str, Any]],
    page_width: float,
    page_height: float,
) -> list[dict[str, Any]]:
    cleaned: list[dict[str, Any]] = []
    for item in items:
        text = re.sub(r"\s+", " ", str(item.get("text") or "")).strip()
        bbox = item.get("bbox")
        if _looks_like_noise_text(text) or not _bbox_list_finite(bbox):
            continue
        x0, y0, x1, y1 = [float(v) for v in bbox]
        width = x1 - x0
        height = y1 - y0
        if width < 6 or height < 8:
            continue
        size = float(item.get("size") or 11.0)
        if len(text) <= 2 and width < size * 0.85:
            continue
        expanded = _expand_bbox_for_readability([x0, y0, x1, y1], text, size, page_width, page_height)
        cleaned.append(
            {
                **item,
                "text": text,
                "bbox": expanded,
                "size": size,
            }
        )

    cleaned.sort(key=lambda x: (x["page"], x["bbox"][1], x["bbox"][0]))
    cleaned = _merge_nearby_line_items(cleaned)
    cleaned = _merge_vertical_column_items(cleaned)
    cleaned = _apply_invoice_column_heuristics(cleaned, page_width)
    return cleaned


def cleanup_ocr_items_for_editor(
    items: list[dict[str, Any]],
    page_width: float,
    page_height: float,
) -> list[dict[str, Any]]:
    return _cleanup_ocr_items(items, page_width, page_height)


def _merge_nearby_line_items(items: list[dict[str, Any]]) -> list[dict[str, Any]]:
    merged: list[dict[str, Any]] = []
    for item in items:
        if not merged:
            merged.append(item)
            continue
        prev = merged[-1]
        if prev["page"] != item["page"]:
            merged.append(item)
            continue
        px0, py0, px1, py1 = [float(v) for v in prev["bbox"]]
        cx0, cy0, cx1, cy1 = [float(v) for v in item["bbox"]]
        ph = max(1.0, py1 - py0)
        ch = max(1.0, cy1 - cy0)
        prev_center = (py0 + py1) / 2.0
        curr_center = (cy0 + cy1) / 2.0
        same_line = abs(prev_center - curr_center) <= max(ph, ch) * 0.38
        gap = cx0 - px1
        small_gap = -4.0 <= gap <= max(18.0, min(float(prev["size"]), float(item["size"])) * 0.95)
        similar_height = min(ph, ch) / max(ph, ch) >= 0.62
        if not (same_line and small_gap and similar_height):
            merged.append(item)
            continue
        sep = "" if re.match(r"^[\)\]\}\.,:%/-]", item["text"]) else " "
        prev["text"] = f"{prev['text']}{sep}{item['text']}".strip()
        prev["bbox"] = [min(px0, cx0), min(py0, cy0), max(px1, cx1), max(py1, cy1)]
        prev["size"] = max(float(prev["size"]), float(item["size"]))
    return merged


def _vertical_merge_score(a: dict[str, Any], b: dict[str, Any]) -> float:
    ax0, ay0, ax1, ay1 = [float(v) for v in a["bbox"]]
    bx0, by0, bx1, by1 = [float(v) for v in b["bbox"]]
    aw = max(1.0, ax1 - ax0)
    bw = max(1.0, bx1 - bx0)
    ah = max(1.0, ay1 - ay0)
    bh = max(1.0, by1 - by0)
    a_center_x = (ax0 + ax1) / 2.0
    b_center_x = (bx0 + bx1) / 2.0
    x_center_diff = abs(a_center_x - b_center_x)
    x_overlap = max(0.0, min(ax1, bx1) - max(ax0, bx0)) / min(aw, bw)
    gap_y = max(0.0, by0 - ay1)
    similar_width = min(aw, bw) / max(aw, bw)
    similar_height = min(ah, bh) / max(ah, bh)
    atext = str(a.get("text") or "").strip()
    btext = str(b.get("text") or "").strip()
    small_fragment = len(atext) <= 5 or len(btext) <= 5
    numeric_fragment = bool(re.fullmatch(r"[\d,./%-]+", atext)) or bool(re.fullmatch(r"[\d,./%-]+", btext))
    if x_overlap < 0.7 and x_center_diff > min(aw, bw) * 0.18:
        return -1.0
    if gap_y > max(8.0, min(ah, bh) * 0.32):
        return -1.0
    if not (small_fragment or numeric_fragment):
        return -1.0
    return (x_overlap * 2.2) + (similar_width * 1.3) + (similar_height * 0.9) - (gap_y / max(8.0, max(ah, bh)))


def _merge_vertical_column_items(items: list[dict[str, Any]]) -> list[dict[str, Any]]:
    if len(items) < 2:
        return items
    source = sorted(items, key=lambda x: (x["page"], x["bbox"][0], x["bbox"][1]))
    used = [False] * len(source)
    out: list[dict[str, Any]] = []
    for i, item in enumerate(source):
        if used[i]:
            continue
        current = dict(item)
        used[i] = True
        changed = True
        while changed:
            changed = False
            best_idx = None
            best_score = -1.0
            for j, cand in enumerate(source):
                if used[j] or cand["page"] != current["page"]:
                    continue
                score = _vertical_merge_score(current, cand)
                if score > best_score:
                    best_idx = j
                    best_score = score
            if best_idx is None or best_score < 3.0:
                break
            cand = source[best_idx]
            used[best_idx] = True
            atext = str(current["text"]).strip()
            btext = str(cand["text"]).strip()
            joiner = ""
            if re.fullmatch(r"[\d,./%-]+", atext) and re.fullmatch(r"[\d,./%-]+", btext):
                joiner = ""
            elif len(atext) <= 4 or len(btext) <= 4:
                joiner = ""
            else:
                joiner = " "
            current["text"] = f"{atext}{joiner}{btext}".strip()
            ax0, ay0, ax1, ay1 = [float(v) for v in current["bbox"]]
            bx0, by0, bx1, by1 = [float(v) for v in cand["bbox"]]
            current["bbox"] = [min(ax0, bx0), min(ay0, by0), max(ax1, bx1), max(ay1, by1)]
            current["size"] = max(float(current["size"]), float(cand["size"]))
            changed = True
        out.append(current)
    out.sort(key=lambda x: (x["page"], x["bbox"][1], x["bbox"][0]))
    return out


def _looks_numericish_text(text: str) -> bool:
    t = (text or "").strip()
    if not t:
        return False
    if re.fullmatch(r"[\d,./:%\-]+", t):
        return True
    return any(ch.isdigit() for ch in t)


def _looks_codeish_text(text: str) -> bool:
    t = (text or "").strip()
    if len(t) < 4:
        return False
    return bool(re.fullmatch(r"[A-Za-z0-9:/#\-_]+", t)) and any(ch.isdigit() for ch in t)


def _apply_invoice_column_heuristics(
    items: list[dict[str, Any]],
    page_width: float,
) -> list[dict[str, Any]]:
    if len(items) < 4:
        return items

    header_keywords = {
        "price": ("price", "listprice", "mrp", "rate"),
        "qty": ("qty", "quantity"),
        "amount": ("amount", "amt", "total"),
        "code": ("code", "redeemcode", "uin", "hsn"),
        "discount": ("discount", "disc"),
    }

    by_page: dict[int, list[dict[str, Any]]] = {}
    for item in items:
        by_page.setdefault(int(item["page"]), []).append(item)

    final_items: list[dict[str, Any]] = []
    for page, page_items in by_page.items():
        headers: list[dict[str, Any]] = []
        for item in page_items:
            txt = re.sub(r"\s+", "", str(item["text"]).lower())
            for col_name, keys in header_keywords.items():
                if any(k in txt for k in keys):
                    x0, y0, x1, y1 = [float(v) for v in item["bbox"]]
                    headers.append(
                        {
                            "name": col_name,
                            "x0": x0,
                            "x1": x1,
                            "y0": y0,
                            "y1": y1,
                            "center": (x0 + x1) / 2.0,
                        }
                    )
                    break

        if not headers:
            final_items.extend(page_items)
            continue

        headers.sort(key=lambda x: x["center"])
        header_row_y = min(h["y0"] for h in headers)
        column_specs: list[dict[str, Any]] = []
        for idx, head in enumerate(headers):
            left_bound = 0.0 if idx == 0 else (headers[idx - 1]["center"] + head["center"]) / 2.0
            right_bound = (
                page_width if idx == len(headers) - 1 else (head["center"] + headers[idx + 1]["center"]) / 2.0
            )
            column_specs.append(
                {
                    "name": head["name"],
                    "left": max(0.0, left_bound),
                    "right": min(page_width, right_bound),
                    "center": head["center"],
                    "header_y": head["y1"],
                }
            )

        adjusted: list[dict[str, Any]] = []
        for item in page_items:
            txt = str(item["text"]).strip()
            x0, y0, x1, y1 = [float(v) for v in item["bbox"]]
            if y1 <= header_row_y:
                adjusted.append(item)
                continue

            item_center = (x0 + x1) / 2.0
            item_width = max(1.0, x1 - x0)
            best_col = None
            best_score = -1.0
            for col in column_specs:
                if y0 < col["header_y"] - 4:
                    continue
                inside = col["left"] <= item_center <= col["right"]
                dist = abs(item_center - col["center"])
                score = (1.0 if inside else 0.0) + max(0.0, 1.0 - (dist / max(28.0, item_width * 1.2)))
                if score > best_score:
                    best_score = score
                    best_col = col

            if best_col is None:
                adjusted.append(item)
                continue

            should_snap = False
            if best_col["name"] in {"price", "qty", "amount", "discount"} and _looks_numericish_text(txt):
                should_snap = True
            if best_col["name"] == "code" and (_looks_codeish_text(txt) or len(txt) >= 6):
                should_snap = True

            if not should_snap:
                adjusted.append(item)
                continue

            target_left = max(best_col["left"] + 2.0, x0)
            target_right = min(best_col["right"] - 2.0, max(x1, best_col["center"] + item_width * 0.55))
            if target_right <= target_left + 10:
                adjusted.append(item)
                continue

            adjusted.append(
                {
                    **item,
                    "bbox": [target_left, y0, target_right, y1],
                }
            )

        final_items.extend(adjusted)

    final_items.sort(key=lambda x: (x["page"], x["bbox"][1], x["bbox"][0]))
    return final_items
