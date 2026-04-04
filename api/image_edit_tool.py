from __future__ import annotations

import io
import json
import math
import re
import uuid
from pathlib import Path
from typing import Any

import fitz
import numpy as np
from fastapi import APIRouter, File, HTTPException, UploadFile
from fastapi.responses import FileResponse, HTMLResponse
from fastapi.templating import Jinja2Templates
from PIL import Image, ImageOps
from pydantic import BaseModel
from starlette.requests import Request

_PREVIEW_HEADERS = {"Cache-Control": "no-store, max-age=0"}
_rapid_ocr_engine: Any = None


class ImageEditItem(BaseModel):
    id: str
    text: str
    page: int
    bbox: list[float]
    font: str = "helv"
    size: float = 11.0
    color: str = "#111111"
    align: str = "left"


class ImageEditRequest(BaseModel):
    file_id: str
    edits: list[ImageEditItem]


def _bbox_list_finite(b: list[float] | None) -> bool:
    if not b or len(b) != 4:
        return False
    try:
        return all(math.isfinite(float(x)) for x in b)
    except (TypeError, ValueError):
        return False


def _clip_rect_to_page(page: fitz.Page, rect: fitz.Rect) -> fitz.Rect | None:
    try:
        for v in (rect.x0, rect.y0, rect.x1, rect.y1):
            if not math.isfinite(float(v)):
                return None
        if rect.is_empty or rect.width <= 0 or rect.height <= 0:
            return None
        clipped = rect & page.rect
        if clipped.is_empty or clipped.width < 1 or clipped.height < 1:
            return None
        return clipped
    except Exception:
        return None


def _median_int(vals: list[int]) -> int:
    if not vals:
        return 255
    s = sorted(vals)
    n = len(s)
    mid = n // 2
    if n % 2:
        return s[mid]
    return (s[mid - 1] + s[mid]) // 2


def _robust_background_rgb01(samples: list[tuple[int, int, int]]) -> tuple[float, float, float]:
    if not samples:
        return (1.0, 1.0, 1.0)
    med = (
        _median_int([p[0] for p in samples]),
        _median_int([p[1] for p in samples]),
        _median_int([p[2] for p in samples]),
    )
    kept = [
        p
        for p in samples
        if abs(p[0] - med[0]) <= 28 and abs(p[1] - med[1]) <= 28 and abs(p[2] - med[2]) <= 28
    ]
    source = kept if len(kept) >= max(6, len(samples) // 5) else samples
    return (
        _median_int([p[0] for p in source]) / 255.0,
        _median_int([p[1] for p in source]) / 255.0,
        _median_int([p[2] for p in source]) / 255.0,
    )


def _sample_background_fill_rgb(page: fitz.Page, inner: fitz.Rect) -> tuple[float, float, float]:
    clip = inner & page.rect
    if clip.is_empty:
        return (1.0, 1.0, 1.0)
    try:
        pix = page.get_pixmap(matrix=fitz.Matrix(3, 3), clip=clip, alpha=False)
    except Exception:
        return (1.0, 1.0, 1.0)
    if pix.width < 3 or pix.height < 3 or not pix.samples:
        return (1.0, 1.0, 1.0)

    data = pix.samples
    stride = pix.n
    width = pix.width
    height = pix.height

    def get_px(x: int, y: int) -> tuple[int, int, int]:
        idx = (y * width + x) * stride
        return (data[idx], data[idx + 1], data[idx + 2])

    samples: list[tuple[int, int, int]] = []
    for x in range(width):
        samples.append(get_px(x, 0))
        samples.append(get_px(x, height - 1))
    for y in range(height):
        samples.append(get_px(0, y))
        samples.append(get_px(width - 1, y))
    return _robust_background_rgb01(samples)


def _map_font_for_fitz(font_name: str | None) -> str:
    f = (font_name or "").lower()
    if not f or "symbol" in f or "dingbat" in f:
        return "helv"
    is_bold = "bold" in f
    is_italic = "italic" in f or "oblique" in f
    if "times" in f:
        if is_bold and is_italic:
            return "tibi"
        if is_bold:
            return "tibo"
        if is_italic:
            return "tiit"
        return "times-roman"
    if "courier" in f or "cour" in f:
        if is_bold and is_italic:
            return "cobi"
        if is_bold:
            return "cobo"
        if is_italic:
            return "coit"
        return "cour"
    if is_bold and is_italic:
        return "hebi"
    if is_bold:
        return "hebo"
    if is_italic:
        return "heit"
    return "helv"


def _insert_textbox_fit(
    page: fitz.Page,
    rect: fitz.Rect,
    text: str,
    font_name: str,
    start_size: float,
    color_hex: str = "#111111",
    align: str = "left",
) -> None:
    color_hex = color_hex.lstrip("#")
    if len(color_hex) == 6:
        r, g, b = tuple(int(color_hex[i : i + 2], 16) / 255.0 for i in (0, 2, 4))
    else:
        r, g, b = (0.07, 0.07, 0.07)
    if align == "center":
        align_code = fitz.TEXT_ALIGN_CENTER
    elif align == "right":
        align_code = fitz.TEXT_ALIGN_RIGHT
    else:
        align_code = fitz.TEXT_ALIGN_LEFT
    expanded_rect = fitz.Rect(rect.x0, rect.y0, rect.x1, rect.y1 + max(16, start_size * 0.8))
    page.insert_textbox(
        expanded_rect,
        text,
        fontsize=float(start_size),
        fontname=font_name,
        color=(r, g, b),
        align=align_code,
    )


def _get_rapid_ocr_engine() -> Any:
    global _rapid_ocr_engine
    if _rapid_ocr_engine is not None:
        return _rapid_ocr_engine
    try:
        from rapidocr_onnxruntime import RapidOCR
    except Exception:
        return None
    try:
        _rapid_ocr_engine = RapidOCR()
    except Exception:
        return None
    return _rapid_ocr_engine


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


def _ocr_items_for_image(img: Image.Image, page_index: int) -> list[dict[str, Any]]:
    engine = _get_rapid_ocr_engine()
    if engine is None:
        return []

    rgb = np.array(img.convert("RGB"))
    height, width = rgb.shape[0], rgb.shape[1]
    try:
        ocr_out, _ = engine(rgb)
    except Exception:
        return []

    items: list[dict[str, Any]] = []
    for i, row in enumerate(ocr_out or []):
        if not row or len(row) < 2:
            continue
        box, text = row[0], row[1]
        txt = str(text or "").strip()
        if not txt:
            continue
        try:
            xs = [float(p[0]) for p in box]
            ys = [float(p[1]) for p in box]
        except Exception:
            continue
        x0 = max(0.0, min(xs))
        x1 = min(float(width), max(xs))
        y0 = max(0.0, min(ys))
        y1 = min(float(height), max(ys))
        bh = max(12.0, y1 - y0)
        size = max(10.0, min(72.0, bh * 0.78))
        items.append(
            {
                "id": f"p{page_index}-ocr-{i}",
                "page": page_index + 1,
                "text": txt,
                "bbox": [x0, y0, x1, y1],
                "font": "helv",
                "size": size,
                "color": "#111111",
                "align": "left",
            }
        )
    return _cleanup_ocr_items(items, float(width), float(height))


def _ocr_items_for_pdf_page(page: fitz.Page, page_index: int) -> list[dict[str, Any]]:
    engine = _get_rapid_ocr_engine()
    if engine is None:
        return []

    mat = fitz.Matrix(2, 2)
    pix = page.get_pixmap(matrix=mat, alpha=False)
    h, w = pix.height, pix.width
    if h < 2 or w < 2:
        return []

    rgb = np.frombuffer(pix.samples, dtype=np.uint8).reshape(h, w, pix.n)
    if pix.n == 4:
        rgb = rgb[:, :, :3]

    try:
        ocr_out, _ = engine(rgb)
    except Exception:
        return []

    page_w = float(page.rect.width)
    page_h = float(page.rect.height)
    items: list[dict[str, Any]] = []
    for i, row in enumerate(ocr_out or []):
        if not row or len(row) < 2:
            continue
        box, text = row[0], row[1]
        txt = str(text or "").strip()
        if not txt:
            continue
        try:
            xs = [float(p[0]) for p in box]
            ys = [float(p[1]) for p in box]
        except Exception:
            continue
        x0 = min(xs) / float(w) * page_w
        x1 = max(xs) / float(w) * page_w
        y0 = min(ys) / float(h) * page_h
        y1 = max(ys) / float(h) * page_h
        bh = max(10.0, y1 - y0)
        items.append(
            {
                "id": f"p{page_index}-ocr-{i}",
                "page": page_index + 1,
                "text": txt,
                "bbox": [x0, y0, x1, y1],
                "font": "helv",
                "size": max(9.0, min(72.0, bh * 0.78)),
                "color": "#111111",
                "align": "left",
            }
        )
    return _cleanup_ocr_items(items, page_w, page_h)


def create_image_edit_router(base_dir: Path, work_dir: Path) -> APIRouter:
    router = APIRouter()
    templates = Jinja2Templates(directory=str(base_dir / "templates"))
    tool_root = work_dir / "image-editable-pdf"
    tool_root.mkdir(parents=True, exist_ok=True)

    def _folder(file_id: str) -> Path:
        return tool_root / file_id

    def _input_pdf(file_id: str) -> Path:
        return _folder(file_id) / "input.pdf"

    def _edited_pdf(file_id: str) -> Path:
        return _folder(file_id) / "edited.pdf"

    def _preview_png(file_id: str, page_number: int) -> Path:
        return _folder(file_id) / f"preview-{page_number}.png"

    def _ocr_json(file_id: str) -> Path:
        return _folder(file_id) / "ocr.json"

    def _resolve_pdf(file_id: str) -> Path:
        edited = _edited_pdf(file_id)
        if edited.exists():
            return edited
        source = _input_pdf(file_id)
        if source.exists():
            return source
        raise HTTPException(status_code=404, detail="File not found")

    def _main_folder(file_id: str) -> Path:
        return work_dir / file_id

    def _main_input_pdf(file_id: str) -> Path:
        return _main_folder(file_id) / "input.pdf"

    def _main_output_pdf(file_id: str) -> Path:
        return _main_folder(file_id) / "edited.pdf"

    def _resolve_main_pdf(file_id: str) -> Path:
        out = _main_output_pdf(file_id)
        if out.exists():
            return out
        src = _main_input_pdf(file_id)
        if src.exists():
            return src
        raise HTTPException(status_code=404, detail="File not found")

    @router.get("/image-to-editable-pdf", response_class=HTMLResponse)
    async def image_to_editable_pdf_page(request: Request) -> HTMLResponse:
        return templates.TemplateResponse("image_to_editable_pdf.html", {"request": request})

    @router.post("/image_to_pdf_ocr/upload")
    async def upload_images_for_editable_pdf(files: list[UploadFile] = File(...)) -> dict[str, Any]:
        if not files:
            raise HTTPException(status_code=400, detail="Please upload at least one image")

        file_id = str(uuid.uuid4())
        folder = _folder(file_id)
        folder.mkdir(parents=True, exist_ok=True)
        pdf_path = _input_pdf(file_id)

        doc = fitz.open()
        pages: list[dict[str, Any]] = []
        items: list[dict[str, Any]] = []
        names: list[str] = []
        try:
            for page_index, file in enumerate(files):
                raw = await file.read()
                if len(raw) < 32:
                    continue
                try:
                    img = Image.open(io.BytesIO(raw))
                    img = ImageOps.exif_transpose(img).convert("RGB")
                except Exception as e:
                    raise HTTPException(status_code=400, detail=f"Unsupported image: {file.filename}") from e

                names.append(file.filename or f"image-{page_index + 1}.png")
                width, height = img.size
                page = doc.new_page(width=width, height=height)

                img_buf = io.BytesIO()
                img.save(img_buf, format="PNG")
                page.insert_image(page.rect, stream=img_buf.getvalue(), keep_proportion=False)

                pages.append({"page": page_index + 1, "width": float(width), "height": float(height)})
                items.extend(_ocr_items_for_image(img, page_index))

            if doc.page_count == 0:
                raise HTTPException(status_code=400, detail="No valid images were uploaded")

            doc.save(str(pdf_path), garbage=3, deflate=True)
        finally:
            doc.close()

        payload = {"file_id": file_id, "pages": pages, "items": items, "filenames": names}
        _ocr_json(file_id).write_text(json.dumps(payload, ensure_ascii=True), encoding="utf-8")
        return payload

    @router.get("/image_to_pdf_ocr/preview/{file_id}/{page_number}")
    async def preview_image_editable_pdf(file_id: str, page_number: int) -> FileResponse:
        pdf_path = _resolve_pdf(file_id)
        out_path = _preview_png(file_id, page_number)
        out_path.parent.mkdir(parents=True, exist_ok=True)

        doc = fitz.open(pdf_path)
        try:
            if page_number < 1 or page_number > doc.page_count:
                raise HTTPException(status_code=404, detail="Page not found")
            page = doc[page_number - 1]
            pix = page.get_pixmap(matrix=fitz.Matrix(1.5, 1.5), alpha=False)
            pix.save(str(out_path))
        finally:
            doc.close()

        return FileResponse(out_path, media_type="image/png", headers=_PREVIEW_HEADERS)

    @router.post("/image_to_pdf_ocr/edit")
    async def edit_image_based_pdf(payload: ImageEditRequest) -> dict[str, str]:
        source = _input_pdf(payload.file_id)
        if not source.exists():
            raise HTTPException(status_code=404, detail="File not found")

        edits = [
            item
            for item in payload.edits
            if item.text is not None and _bbox_list_finite(item.bbox)
        ]
        doc = fitz.open(source)
        try:
            by_page: dict[int, list[ImageEditItem]] = {}
            for item in edits:
                by_page.setdefault(item.page - 1, []).append(item)

            for page_index, page_edits in by_page.items():
                if page_index < 0 or page_index >= doc.page_count:
                    continue
                page = doc[page_index]
                valid_ops: list[tuple[ImageEditItem, fitz.Rect]] = []
                for edit in page_edits:
                    rr = _clip_rect_to_page(page, fitz.Rect(edit.bbox))
                    if rr is None:
                        continue
                    rr = fitz.Rect(rr.x0 - 2, rr.y0 - 2, rr.x1 + 2, rr.y1 + 2) & page.rect
                    valid_ops.append((edit, rr))

                for _edit, rr in valid_ops:
                    fill_rgb = _sample_background_fill_rgb(page, rr)
                    page.add_redact_annot(rr, fill=fill_rgb)
                if valid_ops:
                    page.apply_redactions(
                        images=fitz.PDF_REDACT_IMAGE_PIXELS,
                        text=fitz.PDF_REDACT_TEXT_REMOVE,
                    )

                for edit, rr in valid_ops:
                    txt = (edit.text or "").strip()
                    if not txt:
                        continue
                    _insert_textbox_fit(
                        page,
                        rr,
                        txt,
                        _map_font_for_fitz(edit.font),
                        float(edit.size or 11),
                        edit.color or "#111111",
                        edit.align or "left",
                    )

            out_path = _edited_pdf(payload.file_id)
            out_path.parent.mkdir(parents=True, exist_ok=True)
            doc.save(str(out_path), garbage=3, deflate=True)
        finally:
            doc.close()

        return {"download_url": f"/image_to_pdf_ocr/download/{payload.file_id}"}

    @router.get("/image_to_pdf_ocr/download/{file_id}")
    async def download_image_editable_pdf(file_id: str) -> FileResponse:
        pdf_path = _resolve_pdf(file_id)
        return FileResponse(
            pdf_path,
            media_type="application/pdf",
            filename="image-editable.pdf",
        )

    @router.get("/image_ocr_tool/analyze/{file_id}")
    async def analyze_existing_pdf_for_image_ocr(file_id: str) -> dict[str, Any]:
        pdf_path = _resolve_main_pdf(file_id)
        doc = fitz.open(pdf_path)
        try:
            pages: list[dict[str, Any]] = []
            items: list[dict[str, Any]] = []
            for page_index in range(doc.page_count):
                page = doc[page_index]
                pages.append(
                    {
                        "page": page_index + 1,
                        "width": float(page.rect.width),
                        "height": float(page.rect.height),
                    }
                )
                items.extend(_ocr_items_for_pdf_page(page, page_index))
            return {"file_id": file_id, "pages": pages, "items": items}
        finally:
            doc.close()

    @router.post("/image_ocr_tool/edit")
    async def edit_existing_pdf_with_image_ocr(payload: ImageEditRequest) -> dict[str, str]:
        source = _resolve_main_pdf(payload.file_id)
        out_path = _main_output_pdf(payload.file_id)
        out_path.parent.mkdir(parents=True, exist_ok=True)

        edits = [
            item
            for item in payload.edits
            if item.text is not None and _bbox_list_finite(item.bbox)
        ]
        doc = fitz.open(source)
        try:
            by_page: dict[int, list[ImageEditItem]] = {}
            for item in edits:
                by_page.setdefault(item.page - 1, []).append(item)

            for page_index, page_edits in by_page.items():
                if page_index < 0 or page_index >= doc.page_count:
                    continue
                page = doc[page_index]
                valid_ops: list[tuple[ImageEditItem, fitz.Rect]] = []
                for edit in page_edits:
                    rr = _clip_rect_to_page(page, fitz.Rect(edit.bbox))
                    if rr is None:
                        continue
                    rr = fitz.Rect(rr.x0 - 2, rr.y0 - 2, rr.x1 + 2, rr.y1 + 2) & page.rect
                    valid_ops.append((edit, rr))

                for _edit, rr in valid_ops:
                    fill_rgb = _sample_background_fill_rgb(page, rr)
                    page.add_redact_annot(rr, fill=fill_rgb)
                if valid_ops:
                    page.apply_redactions(
                        images=fitz.PDF_REDACT_IMAGE_PIXELS,
                        text=fitz.PDF_REDACT_TEXT_REMOVE,
                    )

                for edit, rr in valid_ops:
                    txt = (edit.text or "").strip()
                    if not txt:
                        continue
                    _insert_textbox_fit(
                        page,
                        rr,
                        txt,
                        _map_font_for_fitz(edit.font),
                        float(edit.size or 11),
                        edit.color or "#111111",
                        edit.align or "left",
                    )

            doc.save(str(out_path), garbage=3, deflate=True)
        finally:
            doc.close()

        return {"download_url": f"/download/{payload.file_id}"}

    return router
