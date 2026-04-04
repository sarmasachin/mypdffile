from __future__ import annotations

import shutil
import uuid
from pathlib import Path
from collections import defaultdict
from typing import Any

import zipfile
import io
import math
import os
import re
import fitz  # PyMuPDF
from fastapi import FastAPI, File, Form, HTTPException, UploadFile
from fastapi.responses import FileResponse, HTMLResponse, Response, StreamingResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates
from api.image_edit_tool import create_image_edit_router, cleanup_ocr_items_for_editor
from pydantic import BaseModel, Field
from starlette.requests import Request
from uvicorn.middleware.proxy_headers import ProxyHeadersMiddleware

BASE_DIR = Path(__file__).resolve().parent
WORK_DIR = BASE_DIR / "work"
WORK_DIR.mkdir(exist_ok=True)
STAMP_DIR = WORK_DIR / "stamps"
STAMP_DIR.mkdir(exist_ok=True)
# Saved beside each watermarked export so /remove_watermark can restore the PDF before watermark.
NO_WM_BACKUP_NAME = "_no_wm.pdf"


def _carry_no_wm_backup_if_present(src_file_id: str, dest_dir: Path) -> None:
    """Keep remove-watermark working after compress/reorder/etc. create a new workspace id."""
    if not src_file_id:
        return
    src = WORK_DIR / src_file_id / NO_WM_BACKUP_NAME
    if src.is_file():
        dest_dir.mkdir(parents=True, exist_ok=True)
        shutil.copy2(src, dest_dir / NO_WM_BACKUP_NAME)

app = FastAPI(title="PDF Editor Tool")
# Behind Render / other reverse proxies so request.url_for / schemes stay correct.
app.add_middleware(ProxyHeadersMiddleware, trusted_hosts="*")
templates = Jinja2Templates(directory=str(BASE_DIR / "templates"))
app.include_router(create_image_edit_router(WORK_DIR))


class EditItem(BaseModel):
    id: str
    text: str
    page: int
    bbox: list[float]
    original_bbox: list[float] | None = None
    font: str = "helv"
    size: float = 11.0
    color: str = "#000000"
    align: str = "left"
    is_underline: bool = False
    is_strike: bool = False


class EditRequest(BaseModel):
    file_id: str
    edits: list[EditItem]


class PasswordRequest(BaseModel):
    file_id: str
    password: str


class CombineRequest(BaseModel):
    file_ids: list[str]


class SplitRequest(BaseModel):
    file_id: str
    page_indices: list[int]


class RotateRequest(BaseModel):
    file_id: str
    angle: int
    pages: list[int] | None = None


class CropRequest(BaseModel):
    file_id: str
    page: int
    left: float = 0.0
    top: float = 0.0
    right: float = 0.0
    bottom: float = 0.0
    all_pages: bool = False


class CropNormRequest(BaseModel):
    """Crop rectangle in normalized coordinates (0–1) relative to each page's mediabox."""

    file_id: str
    x0: float
    y0: float
    x1: float
    y1: float
    page: int = 1
    all_pages: bool = False


class RemoveWatermarkRequest(BaseModel):
    file_id: str


class WatermarkRequest(BaseModel):
    file_id: str
    text: str = "Draft"
    opacity: float = Field(0.25, ge=0.05, le=0.95)
    font_size: float = Field(48.0, ge=8.0, le=200.0, description="Font size in PDF points")
    position: str = Field(
        "center",
        description="e.g. center, diagonal, top_left, top_center, bottom_right, …",
    )


class MetadataUpdateRequest(BaseModel):
    file_id: str
    title: str | None = None
    author: str | None = None
    strip: bool = False


class ApplyStampRequest(BaseModel):
    file_id: str
    stamp_id: str
    page: int
    x0: float
    y0: float
    x1: float
    y1: float


def input_path(file_id: str) -> Path:
    return WORK_DIR / file_id / "input.pdf"


def ensure_output_path(file_id: str) -> Path:
    p = output_path(file_id)
    p.parent.mkdir(parents=True, exist_ok=True)
    return p


def output_path(file_id: str) -> Path:
    return WORK_DIR / file_id / "edited.pdf"


def preview_path(file_id: str, page_number: int) -> Path:
    return WORK_DIR / file_id / f"preview-{page_number}.png"


def preview_input_path(file_id: str, page_number: int) -> Path:
    """Editor overlay: always from input.pdf so bboxes match /analyze (never edited.pdf)."""
    return WORK_DIR / file_id / f"preview_input-{page_number}.png"


def preview_edited_path(file_id: str, page_number: int) -> Path:
    return WORK_DIR / file_id / f"preview_edited-{page_number}.png"


# /preview aur /preview_edited dono isi scale par — warna save ke baad alag zoom se "dusri image" jaisa lage.
_PREVIEW_MATRIX = fitz.Matrix(2, 2)


def _preview_png_stale(png: Path, pdf: Path) -> bool:
    """True if png is missing or older than the PDF (need to re-render)."""
    try:
        if not pdf.is_file():
            return True
        if not png.is_file():
            return True
        return pdf.stat().st_mtime > png.stat().st_mtime
    except OSError:
        return True


def _write_workspace_page_previews(file_id: str, doc: fitz.Document) -> None:
    """After save: refresh preview-*.png and preview_edited-*.png so home + post-save preview match edited.pdf."""
    out_pdf = output_path(file_id)
    out_pdf.parent.mkdir(parents=True, exist_ok=True)
    for pg_num in range(1, doc.page_count + 1):
        page = doc[pg_num - 1]
        pix = page.get_pixmap(matrix=_PREVIEW_MATRIX, alpha=False)
        pix.save(str(preview_edited_path(file_id, pg_num)))
        pix.save(str(preview_path(file_id, pg_num)))


_PREVIEW_CACHE_HEADERS = {"Cache-Control": "no-store, max-age=0"}


def fitz_open_workspace_pdf(path: str | Path) -> fitz.Document:
    """Open a workspace PDF. Passwords are not read from disk; use unlock flow for encrypted files."""
    return fitz.open(path)


def _safe_fitz_close(doc: fitz.Document | None) -> None:
    """Avoid ValueError('document closed or encrypted') from double-close after some save() paths."""
    if doc is None:
        return
    try:
        if getattr(doc, "is_closed", False):
            return
        doc.close()
    except Exception:
        pass


def _workspace_user_password_path(file_id: str) -> Path:
    """Deprecated: do not persist passwords on disk."""
    return WORK_DIR / file_id / ".pdf_user_pw"


def ensure_file(file_id: str) -> Path:
    path = input_path(file_id)
    if not path.exists():
        raise HTTPException(status_code=404, detail="PDF file not found")
    return path


def resolve_pdf_path(file_id: str) -> Path:
    out = output_path(file_id)
    if out.exists():
        return out
    inp = input_path(file_id)
    if not inp.exists():
        raise HTTPException(status_code=404, detail="PDF file not found")
    return inp


def _new_work_file_id() -> tuple[str, Path]:
    new_id = str(uuid.uuid4())
    out_dir = WORK_DIR / new_id
    out_dir.mkdir(parents=True, exist_ok=True)
    return new_id, out_dir / "input.pdf"


_WATERMARK_POSITIONS = frozenset(
    {
        "top_left",
        "top_center",
        "top_right",
        "middle_left",
        "center",
        "middle_right",
        "bottom_left",
        "bottom_center",
        "bottom_right",
        "diagonal",
        "four_corners",
        "perimeter",
    }
)


def _normalize_watermark_position(raw: str) -> str:
    s = (raw or "center").strip().lower().replace("-", "_").replace(" ", "_")
    aliases = {
        "centre": "center",
        "middle": "center",
        "mid": "center",
        "tl": "top_left",
        "tc": "top_center",
        "tr": "top_right",
        "ml": "middle_left",
        "mr": "middle_right",
        "bl": "bottom_left",
        "bc": "bottom_center",
        "br": "bottom_right",
        "diag": "diagonal",
        "border": "perimeter",
        "frame": "perimeter",
        "corners": "four_corners",
        "four_corner": "four_corners",
        "repeat": "perimeter",
        "around": "perimeter",
    }
    s = aliases.get(s, s)
    if s not in _WATERMARK_POSITIONS:
        raise ValueError(
            "Invalid position. Use one of: top_left, top_center, top_right, middle_left, center, "
            "middle_right, bottom_left, bottom_center, bottom_right, diagonal, four_corners, perimeter"
        )
    return s


def _watermark_gray_color(opacity: float) -> tuple[float, float, float]:
    o = float(min(max(opacity, 0.05), 1.0))
    gray = max(0.72, min(0.92, 0.78 + (1.0 - o) * 0.12))
    return (gray, gray, gray)


def _add_diagonal_center_watermark(page: fitz.Page, text: str, opacity: float, font_size: float) -> None:
    """Classic single watermark: centered, tilted along page diagonal (bottom-left → top-right)."""
    r = page.rect
    color = _watermark_gray_color(opacity)
    t = (text or "").strip() or "Draft"
    fs = float(max(12.0, min(120.0, font_size)))
    pw, ph = r.width, r.height
    cx = (r.x0 + r.x1) / 2
    cy = (r.y0 + r.y1) / 2
    twidth = fitz.get_text_length(t, fontname="helv", fontsize=fs)
    origin = fitz.Point(cx - twidth / 2, cy + fs * 0.35)
    twt = fitz.TextWriter(r)
    twt.append(origin, t, fontsize=fs)
    pivot = fitz.Point(cx, cy)
    scale = max(0.75, min(1.4, fs / 48.0))
    angle_deg = math.degrees(math.atan2(-ph, pw))
    mat = fitz.Matrix(scale, scale).prerotate(angle_deg)
    twt.write_text(page, color=color, morph=(pivot, mat), overlay=0)


def _add_subtle_four_corner_watermark(page: fitz.Page, text: str, opacity: float, font_size: float) -> None:
    """
    One small label in each corner only (readable but not covering the page body).
    Font size is capped ~7–11pt; drawn behind page content (overlay=False).
    """
    r = page.rect
    color = _watermark_gray_color(opacity)
    t = (text or "").strip() or "Draft"
    fs = max(7.0, min(11.0, float(font_size) * 0.22 + 4.5))
    tw = fitz.get_text_length(t, fontname="helv", fontsize=fs)
    pw, ph = r.width, r.height
    m = min(pw, ph) * 0.022
    y_top = r.y0 + m + fs * 0.82
    y_bot = r.y1 - m - fs * 0.12
    for px, py in (
        (r.x0 + m, y_top),
        (r.x1 - m - tw, y_top),
        (r.x0 + m, y_bot),
        (r.x1 - m - tw, y_bot),
    ):
        page.insert_text(
            fitz.Point(px, py),
            t,
            fontsize=fs,
            fontname="helv",
            color=color,
            overlay=False,
        )


def _add_perimeter_tiled_small(page: fitz.Page, text: str, opacity: float, font_size: float) -> None:
    """
    Repeat small text along all four edges (top L→R, right T→B, bottom R→L, left B→T)
    as many times as fit. Font stays ~7–11pt so the page is not covered like large tiles.
    """
    r = page.rect
    color = _watermark_gray_color(opacity)
    t = (text or "").strip() or "Draft"
    fs = max(7.0, min(11.0, float(font_size) * 0.22 + 4.5))
    tw = fitz.get_text_length(t, fontname="helv", fontsize=fs)
    pw, ph = r.width, r.height
    m = min(pw, ph) * 0.03
    gap = fs * 0.12
    chunk = tw + gap

    x0, y0, x1, y1 = r.x0 + m, r.y0 + m, r.x1 - m, r.y1 - m
    step_v = max(fs * 1.02, tw * 0.42)

    y_base = y0 + fs * 0.72
    x = x0
    while x + tw <= x1:
        page.insert_text(
            fitz.Point(x, y_base),
            t,
            fontsize=fs,
            fontname="helv",
            color=color,
            overlay=False,
        )
        x += chunk

    x_r = x1 - fs * 0.22
    y = y0 + fs * 0.35
    while y + tw <= y1:
        page.insert_text(
            fitz.Point(x_r, y),
            t,
            fontsize=fs,
            fontname="helv",
            color=color,
            rotate=90,
            overlay=False,
        )
        y += step_v

    y_bot = y1 - fs * 0.15
    x = x1 - tw
    while x >= x0:
        page.insert_text(
            fitz.Point(x, y_bot),
            t,
            fontsize=fs,
            fontname="helv",
            color=color,
            overlay=False,
        )
        x -= chunk

    x_l = x0 + fs * 0.22
    y = y1 - fs * 0.2
    while y - tw >= y0:
        page.insert_text(
            fitz.Point(x_l, y),
            t,
            fontsize=fs,
            fontname="helv",
            color=color,
            rotate=270,
            overlay=False,
        )
        y -= step_v


def _add_watermark(page: fitz.Page, text: str, opacity: float, font_size: float, position: str) -> None:
    """Place watermark using user-chosen size and position (normalized key).

    Drawn with overlay=0 so it sits *behind* existing page content (text/images drawn
    on top remain readable). Pure image-only pages may hide the mark where pixels are opaque.
    """
    r = page.rect
    color = _watermark_gray_color(opacity)
    fs = float(max(8.0, min(200.0, font_size)))

    if position == "perimeter":
        _add_perimeter_tiled_small(page, text, opacity, fs)
        return

    if position == "diagonal":
        _add_diagonal_center_watermark(page, text, opacity, fs)
        return

    if position == "four_corners":
        _add_subtle_four_corner_watermark(page, text, opacity, fs)
        return

    pw, ph = r.width, r.height
    margin = min(pw, ph) * 0.035
    # Single placement modes: cap size so one watermark does not cover the whole page
    fs = min(fs, 30.0)
    tw = fitz.get_text_length(text, fontname="helv", fontsize=fs)
    th = fs * 1.45
    box_w = min(max(tw + fs * 0.6, fs * 2.5), pw - 2 * margin)
    box_h = min(th + fs * 0.3, ph - 2 * margin)

    centers: dict[str, tuple[float, float]] = {
        "top_left": (r.x0 + margin + box_w / 2, r.y0 + margin + box_h / 2),
        "top_center": (r.x0 + pw / 2, r.y0 + margin + box_h / 2),
        "top_right": (r.x1 - margin - box_w / 2, r.y0 + margin + box_h / 2),
        "middle_left": (r.x0 + margin + box_w / 2, r.y0 + ph / 2),
        "center": (r.x0 + pw / 2, r.y0 + ph / 2),
        "middle_right": (r.x1 - margin - box_w / 2, r.y0 + ph / 2),
        "bottom_left": (r.x0 + margin + box_w / 2, r.y1 - margin - box_h / 2),
        "bottom_center": (r.x0 + pw / 2, r.y1 - margin - box_h / 2),
        "bottom_right": (r.x1 - margin - box_w / 2, r.y1 - margin - box_h / 2),
    }
    cx, cy = centers[position]
    rect = fitz.Rect(cx - box_w / 2, cy - box_h / 2, cx + box_w / 2, cy + box_h / 2)
    rect = rect & r
    page.insert_textbox(
        rect,
        text,
        fontsize=fs,
        fontname="helv",
        color=color,
        align=fitz.TEXT_ALIGN_CENTER,
        overlay=False,
    )


def _insert_textbox_fit(
    page: fitz.Page,
    rect: fitz.Rect,
    text: str,
    font_name: str,
    start_size: float,
    color_hex: str = "#000000",
    align: str = "left",
) -> None:
    # Parse hex color safely
    color_hex = color_hex.lstrip("#")
    if len(color_hex) == 6:
        r, g, b = tuple(int(color_hex[i:i+2], 16) / 255.0 for i in (0, 2, 4))
    else:
        r, g, b = (0.0, 0.0, 0.0)

    if align == "center":
        align_code = fitz.TEXT_ALIGN_CENTER
    elif align == "right":
        align_code = fitz.TEXT_ALIGN_RIGHT
    else:
        align_code = fitz.TEXT_ALIGN_LEFT

    # Do not extend x1: +400 made the layout box almost full-line wide so wrapped lines broke
    # at wrong places vs the table column (preview looked "left" / jagged vs Description header).
    expanded_rect = fitz.Rect(rect.x0, rect.y0, rect.x1, rect.y1 + 400)

    page.insert_textbox(
        expanded_rect,
        text,
        fontsize=float(start_size),
        fontname=font_name,
        color=(r, g, b),
        align=align_code,
    )


# PyMuPDF short names sometimes fail on specific PDFs; try PDF standard names before dropping bold.
_FITZ_FONT_FALLBACKS: dict[str, list[str]] = {
    "hebo": ["Helvetica-Bold"],
    "hebi": ["Helvetica-BoldOblique"],
    "heit": ["Helvetica-Oblique"],
    "tibo": ["Times-Bold"],
    "tibi": ["Times-BoldItalic"],
    "tiit": ["Times-Italic"],
    "cobo": ["Courier-Bold"],
    "cobi": ["Courier-BoldOblique"],
    "coit": ["Courier-Oblique"],
}


def _insert_font_size_for_rect(edit: EditItem, rect: fitz.Rect) -> float:
    """OCR `size` can be low vs cell height; scale up so replaced text matches scan neighbors."""
    fs = float(edit.size or 11)
    h = max(0.1, float(rect.height))
    from_h = h / 1.2
    return max(fs, min(72.0, from_h * 0.94))


def _insert_textbox_fit_try_font_chain(
    page: fitz.Page,
    rect: fitz.Rect,
    text: str,
    mapped_font: str,
    start_size: float,
    color_hex: str,
    align: str,
) -> None:
    seen: set[str] = set()
    chain = [mapped_font] + _FITZ_FONT_FALLBACKS.get(mapped_font, [])
    chain.append("helv")
    last_err: Exception | None = None
    for fn in chain:
        if fn in seen:
            continue
        seen.add(fn)
        try:
            _insert_textbox_fit(page, rect, text, fn, start_size, color_hex, align)
            return
        except Exception as e:
            last_err = e
            continue
    if last_err:
        raise last_err


def _min_insert_width_pt(edit: EditItem) -> float:
    """Minimum box width so insert_textbox does not wrap one glyph per line (table/unify bug)."""
    fs = float(edit.size or 11)
    t = (edit.text or "").replace("\r\n", "\n")
    lines = [ln for ln in t.split("\n") if ln.strip()] or [t or " "]
    max_len = max(len(ln) for ln in lines)
    est = max_len * fs * 0.52
    return max(fs * 3.0, min(est, 520.0))


# Invoice/table: many rows share similar x0/x1; paragraph-unify must not merge them all.
_MAX_LINES_UNIFY_AS_ONE_PARAGRAPH = 12


def _map_font_for_fitz(font_name: str | None) -> str:
    f = (font_name or "").lower()
    if not f or "symbol" in f or "dingbat" in f:
        return "helv"
        
    is_bold = "bold" in f
    is_italic = "italic" in f or "oblique" in f

    if "times" in f:
        if is_bold and is_italic: return "tibi"
        if is_bold: return "tibo"
        if is_italic: return "tiit"
        return "times-roman"
        
    if "courier" in f or "cour" in f:
        if is_bold and is_italic: return "cobi"
        if is_bold: return "cobo"
        if is_italic: return "coit"
        return "cour"
        
    if is_bold and is_italic: return "hebi"
    if is_bold: return "hebo"
    if is_italic: return "heit"
    return "helv"


def _bbox_list_finite(b: list[float] | None) -> bool:
    if not b or len(b) != 4:
        return False
    try:
        return all(math.isfinite(float(x)) for x in b)
    except (TypeError, ValueError):
        return False


def _clip_rect_to_page(page: fitz.Page, rect: fitz.Rect) -> fitz.Rect | None:
    """Intersect with page mediabox; None if rect is invalid or unusable (avoids PyMuPDF errors)."""
    try:
        for v in (rect.x0, rect.y0, rect.x1, rect.y1):
            if not math.isfinite(float(v)):
                return None
        if rect.is_empty or rect.width <= 0 or rect.height <= 0:
            return None
        clipped = rect & page.rect
        if clipped.is_empty or clipped.width < 0.5 or clipped.height < 0.5:
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


def _median_rgb01(samples: list[tuple[int, int, int]]) -> tuple[float, float, float]:
    if not samples:
        return (1.0, 1.0, 1.0)
    r = _median_int([p[0] for p in samples])
    g = _median_int([p[1] for p in samples])
    b = _median_int([p[2] for p in samples])
    return (
        min(1.0, max(0.0, r / 255.0)),
        min(1.0, max(0.0, g / 255.0)),
        min(1.0, max(0.0, b / 255.0)),
    )


def _dist_to_inner_rect_px(x: int, y: int, ix0: int, iy0: int, ix1: int, iy1: int) -> float:
    """Half-open rect [ix0, ix1) x [iy0, iy1) in pixel space; 0 if inside."""
    if ix1 <= ix0 or iy1 <= iy0:
        return 1e9
    cx = max(ix0, min(x, ix1 - 1))
    cy = max(iy0, min(y, iy1 - 1))
    return math.hypot(float(x - cx), float(y - cy))


def _robust_background_rgb01(samples: list[tuple[int, int, int]]) -> tuple[float, float, float]:
    """
    Pehle per-channel median, phir usse bahut alag pixels hatao (text/edge),
    phir dubara median — mean se zyada stable, patch kam dikhta hai.
    """
    if not samples:
        return (1.0, 1.0, 1.0)
    m0 = (
        _median_int([p[0] for p in samples]),
        _median_int([p[1] for p in samples]),
        _median_int([p[2] for p in samples]),
    )
    thresh_sq = 65 * 65
    kept = [
        p
        for p in samples
        if (p[0] - m0[0]) ** 2 + (p[1] - m0[1]) ** 2 + (p[2] - m0[2]) ** 2 <= thresh_sq
    ]
    if len(kept) >= max(6, len(samples) // 5):
        return _median_rgb01(kept)
    return _median_rgb01(samples)


def _sample_background_fill_rgb(page: fitz.Page, inner: fitz.Rect) -> tuple[float, float, float]:
    """
    Redaction fill: background jaisa RGB. Poora pixmap average mat lo — design me safed
    hissa mil kar fill phir safed ho jata hai. Sirf margin / pixmap border / corners.
    """
    margin_pt = 28.0
    clip = fitz.Rect(inner)
    clip.x0 -= margin_pt
    clip.y0 -= margin_pt
    clip.x1 += margin_pt
    clip.y1 += margin_pt
    clip &= page.rect
    if clip.is_empty or clip.width < 2 or clip.height < 2:
        return (1.0, 1.0, 1.0)
    try:
        pix = page.get_pixmap(matrix=fitz.Matrix(4, 4), clip=clip, alpha=False)
    except Exception:
        return (1.0, 1.0, 1.0)
    if pix.width < 2 or pix.height < 2 or not pix.samples:
        return (1.0, 1.0, 1.0)
    n = pix.n
    if n < 3:
        return (1.0, 1.0, 1.0)
    rx0 = inner.x0 - clip.x0
    ry0 = inner.y0 - clip.y0
    rx1 = inner.x1 - clip.x0
    ry1 = inner.y1 - clip.y0
    sx = pix.width / clip.width
    sy = pix.height / clip.height
    ix0 = int(rx0 * sx)
    iy0 = int(ry0 * sy)
    ix1 = int(math.ceil(rx1 * sx))
    iy1 = int(math.ceil(ry1 * sy))
    ix0 = max(0, min(pix.width - 1, ix0))
    iy0 = max(0, min(pix.height - 1, iy0))
    ix1 = max(ix0 + 1, min(pix.width, ix1))
    iy1 = max(iy0 + 1, min(pix.height, iy1))
    stride = pix.width * n
    data = pix.samples
    samples: list[tuple[int, int, int]] = []

    def get_px(x: int, y: int) -> tuple[int, int, int]:
        i = y * stride + x * n
        return (data[i], data[i + 1], data[i + 2])

    # 1) Pehle patli ring: sirf text box ke bilkul bahar ke pixels (solid red / card
    #    jaisa rang). Poora donut se stadium / doosre panel ka rang kam aata tha.
    w_in = ix1 - ix0
    h_in = iy1 - iy0
    base_ring = max(4.0, min(14.0, 0.14 * float(max(w_in, h_in, 1))))
    samples = []
    for widen in (1.0, 1.7, 2.4):
        rg = base_ring * widen
        samples = []
        for y in range(pix.height):
            for x in range(pix.width):
                d = _dist_to_inner_rect_px(x, y, ix0, iy0, ix1, iy1)
                if 0 < d <= rg:
                    samples.append(get_px(x, y))
        if len(samples) >= 10:
            break

    # 2) Ring se kam mila to poora donut (clip minus inner)
    if len(samples) < 8:
        samples = []
        for y in range(pix.height):
            for x in range(pix.width):
                if ix0 <= x < ix1 and iy0 <= y < iy1:
                    continue
                samples.append(get_px(x, y))

    # 3) Kam pixel hon to pixmap ke bahri dhari
    if len(samples) < 8:
        border_px = max(2, min(10, pix.width // 12, pix.height // 12))
        samples = []
        for y in range(pix.height):
            for x in range(pix.width):
                if x < border_px or x >= pix.width - border_px or y < border_px or y >= pix.height - border_px:
                    samples.append(get_px(x, y))

    # 4) Char kon
    if len(samples) < 8:
        csz = max(4, min(14, pix.width // 6, pix.height // 6))
        samples = []
        for y0, x0 in ((0, 0), (0, max(0, pix.width - csz)), (max(0, pix.height - csz), 0), (max(0, pix.height - csz), max(0, pix.width - csz))):
            for dy in range(min(csz, pix.height - y0)):
                for dx in range(min(csz, pix.width - x0)):
                    samples.append(get_px(x0 + dx, y0 + dy))

    return _robust_background_rgb01(samples)


@app.get("/", response_class=HTMLResponse)
async def home(request: Request) -> HTMLResponse:
    return templates.TemplateResponse("index.html", {"request": request})


@app.post("/reorder")
async def reorder_pages(req: dict = {"file_id": "", "page_indices": []}) -> dict[str, Any]:
    file_id = req.get("file_id")
    indices = req.get("page_indices", [])
    
    if not file_id or not indices:
        raise HTTPException(status_code=400, detail="Missing data")
        
    source = output_path(file_id)
    if not source.exists():
        source = input_path(file_id)
        
    if not source.exists():
        raise HTTPException(status_code=404, detail="File not found")
        
    new_id = str(uuid.uuid4())
    out_dir = WORK_DIR / new_id
    out_dir.mkdir(parents=True, exist_ok=True)
    out_pdf = out_dir / "input.pdf"
    _carry_no_wm_backup_if_present(file_id, out_dir)

    try:
        doc = fitz_open_workspace_pdf(source)
        # Select pages in the new order (this also deletes any pages not in the list)
        doc.select(indices)
        doc.save(str(out_pdf), garbage=3, deflate=True)
        doc.close()
        
        return {"file_id": new_id, "status": "success", "size": os.path.getsize(out_pdf)}
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Reorder failed: {str(e)}")


@app.post("/upload")
async def upload_pdf(file: UploadFile = File(...)) -> dict[str, Any]:
    file_id = str(uuid.uuid4())
    folder = WORK_DIR / file_id
    folder.mkdir(parents=True, exist_ok=True)
    destination = folder / "input.pdf"
    
    temp_path = folder / f"temp_{file.filename}"
    with temp_path.open("wb") as buffer:
        shutil.copyfileobj(file.file, buffer)
        
    is_encrypted = False
    try:
        # Check if it's already a PDF
        if file.filename.lower().endswith(".pdf"):
            shutil.move(str(temp_path), str(destination))
            doc = fitz_open_workspace_pdf(destination)
            is_encrypted = doc.is_encrypted
            doc.close()
        else:
            # Efficient image to PDF conversion (avoids size bloat)
            img_doc = fitz_open_workspace_pdf(temp_path)
            # Create a point-based page size from image pixels
            img_rect = img_doc[0].rect
            
            pdf_doc = fitz.open() # blank PDF
            page = pdf_doc.new_page(width=img_rect.width, height=img_rect.height)
            # Insert original image bytes without re-compressing
            page.insert_image(img_rect, filename=str(temp_path))
            
            pdf_doc.save(str(destination), garbage=3, deflate=True)
            pdf_doc.close()
            img_doc.close()
            # Cleanup temp
            if temp_path.exists(): os.remove(temp_path)
    except Exception as e:
        if temp_path.exists(): os.remove(temp_path)
        raise HTTPException(status_code=400, detail=f"Unsupported file format: {str(e)}")

    return {"file_id": file_id, "filename": file.filename, "needs_password": is_encrypted}


@app.post("/unlock")
async def unlock_pdf(req: PasswordRequest) -> dict[str, Any]:
    # Reuse PasswordRequest as it contains file_id and password
    path = input_path(req.file_id)
    if not path.exists():
        raise HTTPException(status_code=404, detail="File not found")

    try:
        doc = fitz.open(path)
        if not doc.is_encrypted:
            doc.close()
            return {"status": "success", "message": "File is not encrypted"}
            
        success = doc.authenticate(req.password)
        if not success:
            doc.close()
            raise HTTPException(status_code=401, detail="Invalid password")
        
        # Save decrypted versions so preview/edit can work without storing password.
        tmp_path = path.with_suffix(".unlocked.pdf")
        doc.save(tmp_path)
        doc.close()

        shutil.move(str(tmp_path), str(path))

        outp = output_path(req.file_id)
        if outp.exists():
            try:
                d2 = fitz.open(outp)
                if d2.is_encrypted:
                    ok2 = d2.authenticate(req.password)
                    if ok2:
                        tmp2 = outp.with_suffix(".unlocked.pdf")
                        d2.save(tmp2)
                        d2.close()
                        shutil.move(str(tmp2), str(outp))
                    else:
                        d2.close()
            except Exception:
                pass

        pwp = _workspace_user_password_path(req.file_id)
        if pwp.exists():
            pwp.unlink()
        return {"status": "success", "message": "PDF unlocked successfully"}
    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Unlock failed: {str(e)}")


@app.get("/source/{file_id}")
async def source_pdf(file_id: str) -> FileResponse:
    path = ensure_file(file_id)
    return FileResponse(path=path, filename="source.pdf", media_type="application/pdf")


@app.get("/preview/{file_id}/{page_number}")
async def page_preview(
    file_id: str,
    page_number: int,
    source: str | None = None,
) -> FileResponse:
    """
    Default: latest workspace PDF (edited if present) — thumbnails / read-only browse.
    source=input: original upload only — must match /analyze bboxes in the text editor.
    """
    if (source or "").strip().lower() == "input":
        path = ensure_file(file_id)
        out_path = preview_input_path(file_id, page_number)
    else:
        path = resolve_pdf_path(file_id)
        out_path = preview_path(file_id, page_number)
    doc = fitz.open(path)

    if page_number < 1 or page_number > doc.page_count:
        doc.close()
        raise HTTPException(status_code=404, detail="Page not found")

    if _preview_png_stale(out_path, path):
        out_path.parent.mkdir(parents=True, exist_ok=True)
        page = doc[page_number - 1]
        pix = page.get_pixmap(matrix=_PREVIEW_MATRIX, alpha=False)
        pix.save(str(out_path))

    doc.close()
    return FileResponse(
        path=out_path,
        media_type="image/png",
        headers=_PREVIEW_CACHE_HEADERS,
    )


@app.get("/preview_edited/{file_id}/{page_number}")
async def page_preview_edited(file_id: str, page_number: int) -> FileResponse:
    path = resolve_pdf_path(file_id)
    out_path = preview_edited_path(file_id, page_number)
    if not _preview_png_stale(out_path, path):
        return FileResponse(
            path=out_path,
            media_type="image/png",
            headers=_PREVIEW_CACHE_HEADERS,
        )

    try:
        doc = fitz.open(path)
        if page_number < 1 or page_number > doc.page_count:
            doc.close()
            raise HTTPException(status_code=404, detail="Page not found")
        out_path.parent.mkdir(parents=True, exist_ok=True)
        page = doc[page_number - 1]
        pix = page.get_pixmap(matrix=_PREVIEW_MATRIX, alpha=False)
        pix.save(out_path)
        doc.close()
    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Preview generation failed: {e}")

    return FileResponse(
        path=out_path,
        media_type="image/png",
        headers=_PREVIEW_CACHE_HEADERS,
    )


def _extract_items_from_text_dict(page_index: int, text_dict: dict[str, Any]) -> list[dict[str, Any]]:
    """Build editable items from PyMuPDF text dict (vector text or OCR textpage dict)."""
    out: list[dict[str, Any]] = []
    for block_index, block in enumerate(text_dict.get("blocks", [])):
        if block.get("type") != 0:
            continue

        for line_index, line in enumerate(block.get("lines", [])):
            spans = line.get("spans", [])
            if not spans:
                continue

            chunks: list[list[Any]] = []
            current_chunk: list[Any] = []

            for span in spans:
                t = (span.get("text") or "").strip()
                if not t:
                    continue

                if not current_chunk:
                    current_chunk.append(span)
                else:
                    prev_span = current_chunk[-1]
                    prev_x1 = float(prev_span.get("bbox", [0, 0, 0, 0])[2])
                    curr_x0 = float(span.get("bbox", [0, 0, 0, 0])[0])
                    if curr_x0 - prev_x1 > 12.0:
                        chunks.append(current_chunk)
                        current_chunk = [span]
                    else:
                        current_chunk.append(span)

            if current_chunk:
                chunks.append(current_chunk)

            for chunk_index, chunk in enumerate(chunks):
                chunk_text = " ".join(
                    (s.get("text") or "").strip() for s in chunk if (s.get("text") or "").strip()
                )
                if not chunk_text:
                    continue

                rects: list[fitz.Rect] = []
                for s in chunk:
                    sb = s.get("bbox")
                    if sb and len(sb) == 4:
                        try:
                            rects.append(fitz.Rect([float(v) for v in sb]))
                        except Exception:
                            pass

                if not rects:
                    continue

                x0 = min(r.x0 for r in rects)
                y0 = min(r.y0 for r in rects)
                x1 = max(r.x1 for r in rects)
                y1 = max(r.y1 for r in rects)

                bbox_list = [float(x0), float(y0), float(x1), float(y1)]
                item_id = f"p{page_index}-b{block_index}-l{line_index}-c{chunk_index}"

                first_span = chunk[0]
                font_name = first_span.get("font", "helv") or "helv"
                size = float(first_span.get("size", 11.0))
                # Span "flags" mark synthetic bold/italic even when the font name omits "Bold"/"Italic".
                any_bold = any(
                    int(s.get("flags", 0)) & int(fitz.TEXT_FONT_BOLD) for s in chunk
                )
                any_italic = any(
                    int(s.get("flags", 0)) & int(fitz.TEXT_FONT_ITALIC) for s in chunk
                )
                fl = font_name.lower()
                if any_bold and "bold" not in fl:
                    font_name = f"{font_name}-bold"
                    fl = font_name.lower()
                if any_italic and "italic" not in fl and "oblique" not in fl:
                    font_name = f"{font_name}-italic"

                out.append(
                    {
                        "id": item_id,
                        "page": page_index + 1,
                        "text": chunk_text,
                        "bbox": bbox_list,
                        "font": font_name,
                        "size": size,
                    }
                )
    return out


_rapid_ocr_engine: Any = None


def _ocr_items_rapid(page: fitz.Page, page_index: int) -> list[dict[str, Any]]:
    """Fallback OCR for scanned/photo pages (no vector text). Uses ONNX models via RapidOCR."""
    global _rapid_ocr_engine
    try:
        import numpy as np
        from rapidocr_onnxruntime import RapidOCR
    except Exception:
        return []

    if _rapid_ocr_engine is None:
        try:
            _rapid_ocr_engine = RapidOCR()
        except Exception:
            return []

    mat = fitz.Matrix(2, 2)
    pix = page.get_pixmap(matrix=mat, alpha=False)
    h, w = pix.height, pix.width
    if h < 2 or w < 2:
        return []

    img = np.frombuffer(pix.samples, dtype=np.uint8).reshape(h, w, pix.n)
    if pix.n == 4:
        img = img[:, :, :3]

    try:
        ocr_out, _ = _rapid_ocr_engine(img)
    except Exception:
        return []

    if not ocr_out:
        return []

    page_w = float(page.rect.width)
    page_h = float(page.rect.height)
    items: list[dict[str, Any]] = []
    for i, row in enumerate(ocr_out):
        if not row or len(row) < 2:
            continue
        box, text = row[0], row[1]
        if not text or not str(text).strip():
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
        bh = max(0.1, y1 - y0)
        size = max(8.0, min(72.0, bh * 0.82))
        items.append(
            {
                "id": f"p{page_index}-ocr-{i}",
                "page": page_index + 1,
                "text": str(text).strip(),
                "bbox": [x0, y0, x1, y1],
                "font": "helv",
                "size": size,
            }
        )
    return cleanup_ocr_items_for_editor(items, page_w, page_h)


@app.get("/analyze/{file_id}")
async def analyze_pdf(file_id: str) -> dict[str, Any]:
    # Always original upload — editor uses /preview?source=input; /edit applies onto this file.
    path = ensure_file(file_id)
    try:
        doc = fitz.open(path)
        if doc.is_encrypted:
            doc.close()
            raise HTTPException(status_code=401, detail="PDF is password-protected. Unlock it first.")
    except HTTPException:
        raise
    except Exception:
        raise HTTPException(status_code=401, detail="PDF is password-protected. Unlock it first.")
    items: list[dict[str, Any]] = []
    pages: list[dict[str, float]] = []

    for page_index in range(doc.page_count):
        page = doc[page_index]
        pages.append(
            {
                "page": page_index + 1,
                "width": float(page.rect.width),
                "height": float(page.rect.height),
            }
        )

        text_dict = page.get_text("dict")
        page_items = _extract_items_from_text_dict(page_index, text_dict)
        if not page_items:
            try:
                tp = page.get_textpage_ocr(full=True, dpi=150, language="eng")
                ocr_dict = page.get_text("dict", textpage=tp)
                page_items = _extract_items_from_text_dict(page_index, ocr_dict)
                page_items = cleanup_ocr_items_for_editor(
                    page_items,
                    float(page.rect.width),
                    float(page.rect.height),
                )
            except Exception:
                pass
        if not page_items:
            page_items = _ocr_items_rapid(page, page_index)
        items.extend(page_items)

    doc.close()
    return {"file_id": file_id, "pages": pages, "items": items}


def _unify_paragraph_left_x0_for_insert(edits: list[EditItem]) -> dict[str, float]:
    """
    OCR often splits one paragraph into several boxes with different bbox.x0; insert then
    starts each box at its own left edge so line 2+ looks "outdented" vs line 1.

    Group boxes that belong to the same paragraph (stacked lines in the same column, or
    strong horizontal overlap) and use max(x0) for insert. Redaction still uses each
    block's original bbox.

    Note: If the UI sends only ONE edit for a whole cell, this returns {} — nothing to unify.
    """
    n = len(edits)
    if n < 2:
        return {}
    parent = list(range(n))

    def find(i: int) -> int:
        while parent[i] != i:
            parent[i] = parent[parent[i]]
            i = parent[i]
        return i

    def union(i: int, j: int) -> None:
        pi, pj = find(i), find(j)
        if pi != pj:
            parent[pi] = pj

    def same_cluster(a: EditItem, b: EditItem) -> bool:
        if a.page != b.page:
            return False
        ax0, ay0, ax1, ay1 = a.bbox
        bx0, by0, bx1, by1 = b.bbox
        aw = max(0.01, ax1 - ax0)
        bw = max(0.01, bx1 - bx0)
        inter = max(0.0, min(ax1, bx1) - max(ax0, bx0))
        overlap_ratio = inter / min(aw, bw)
        ah = max(0.01, ay1 - ay0)
        bh = max(0.01, by1 - by0)

        # Same table column: left/right edges line up (paragraph lines stacked vertically).
        same_column = abs(ax0 - bx0) <= 24.0 and abs(ax1 - bx1) <= 24.0
        # Vertical: stacked lines in one cell (small gap), not a full table row skip.
        y_overlap = min(ay1, by1) - max(ay0, by0)
        if y_overlap > 0:
            vertically_adjacent = True
        else:
            gap = max(ay0, by0) - min(ay1, by1)
            max_h = max(ah, bh)
            vertically_adjacent = gap <= max_h * 0.55 + 3.5

        if same_column and vertically_adjacent:
            return True

        # Original rule: overlapping in x and on the same "row band" (side-by-side merge).
        if overlap_ratio < 0.22:
            return False
        acy = (ay0 + ay1) / 2.0
        bcy = (by0 + by1) / 2.0
        if abs(acy - bcy) > max(ah, bh) * 2.2:
            return False
        return True

    for i in range(n):
        for j in range(i + 1, n):
            if same_cluster(edits[i], edits[j]):
                union(i, j)

    clusters: dict[int, list[int]] = defaultdict(list)
    for i in range(n):
        clusters[find(i)].append(i)

    out: dict[str, float] = {}
    for _root, idxs in clusters.items():
        if len(idxs) < 2:
            continue
        if len(idxs) > _MAX_LINES_UNIFY_AS_ONE_PARAGRAPH:
            continue
        mx = max(edits[i].bbox[0] for i in idxs)
        for i in idxs:
            e = edits[i]
            if e.bbox[0] < mx - 1e-6:
                x1 = float(e.bbox[2])
                if x1 - mx >= _min_insert_width_pt(e) - 1e-3:
                    out[e.id] = mx
    return out


@app.post("/edit")
async def edit_pdf(payload: EditRequest) -> dict[str, str]:
    # Client bboxes are always in input.pdf space (matches /analyze). Rebuild edited.pdf from input each save.
    path = ensure_file(payload.file_id)

    normalized_edits = [
        item
        for item in payload.edits
        if item.text is not None
        and _bbox_list_finite(item.bbox)
        and (item.original_bbox is None or _bbox_list_finite(item.original_bbox))
    ]

    if not normalized_edits:
        shutil.copy(path, ensure_output_path(payload.file_id))
        try:
            preview_doc = fitz_open_workspace_pdf(ensure_output_path(payload.file_id))
            _write_workspace_page_previews(payload.file_id, preview_doc)
            preview_doc.close()
        except Exception:
            pass
        return {"download_url": f"/download/{payload.file_id}"}

    doc: fitz.Document | None = None
    try:
        try:
            doc = fitz.open(path)
        except Exception:
            raise HTTPException(
                status_code=401,
                detail="PDF is password-protected. Unlock it first.",
            ) from None
        if doc.is_encrypted:
            raise HTTPException(
                status_code=401,
                detail="PDF is password-protected. Unlock it first.",
            )

        meta_backup = dict(doc.metadata)

        by_page: dict[int, list[EditItem]] = defaultdict(list)
        for item in normalized_edits:
            by_page[item.page - 1].append(item)

        for page_index, edits in by_page.items():
            if page_index < 0 or page_index >= doc.page_count:
                continue
            page = doc[page_index]

            insert_left_unify = _unify_paragraph_left_x0_for_insert(edits)

            def _edit_for_insert_text(e: EditItem) -> EditItem:
                mx = insert_left_unify.get(e.id)
                if mx is None:
                    return e
                b = list(e.bbox)
                x0_o, _, x1_o, _ = (float(b[0]), float(b[1]), float(b[2]), float(b[3]))
                min_w = _min_insert_width_pt(e)
                max_x0 = x1_o - min_w
                if max_x0 <= x0_o + 1e-3:
                    return e
                tx = min(float(mx), max_x0)
                if tx <= x0_o + 1e-6:
                    return e
                b[0] = tx
                return e.model_copy(update={"bbox": b})

            def _padded_rect(edit_item: EditItem, mode: str) -> fitz.Rect:
                """
                Text boxes ka bbox extracted tight hota hai.
                Double text/cut se bachne ke liye redaction rect ko text insert rect se
                thoda bada rakhte hain.
                """
                if mode == "redact":
                    # Union tight OCR + editor bbox so bitmap ink is fully cleared; tight-only
                    # redaction left old scan pixels and new text drew on top (double print).
                    parts: list[fitz.Rect] = []
                    if edit_item.original_bbox and _bbox_list_finite(
                        edit_item.original_bbox
                    ):
                        parts.append(
                            fitz.Rect([float(v) for v in edit_item.original_bbox])
                        )
                    if _bbox_list_finite(edit_item.bbox):
                        parts.append(fitz.Rect([float(v) for v in edit_item.bbox]))
                    if parts:
                        raw = parts[0]
                        for r in parts[1:]:
                            raw |= r
                    else:
                        raw = fitz.Rect(edit_item.bbox)
                else:
                    raw = fitz.Rect(edit_item.bbox)
                    if mode == "text" and edit_item.original_bbox:
                        try:
                            ob = edit_item.original_bbox
                            if len(ob) == 4:
                                o = fitz.Rect([float(v) for v in ob])
                                r = raw
                                # Save path recomputes bbox from DOM; tiny drift used to break the old
                                # "both edges within 2.5pt" rule → insert used shifted x0 (first word jumped left).
                                tol = 8.0
                                x0 = o.x0 if abs(r.x0 - o.x0) <= tol else r.x0
                                x1 = o.x1 if abs(r.x1 - o.x1) <= tol else r.x1
                                if x1 > x0 + 0.5:
                                    raw = fitz.Rect(x0, r.y0, x1, r.y1)
                        except Exception:
                            pass

                base = float(edit_item.size or 11)
                w = max(0.1, float(raw.x1 - raw.x0))
                h = max(0.1, float(raw.y1 - raw.y0))

                txt = (edit_item.text or "").strip()
                lines = [ln for ln in txt.splitlines() if ln.strip() != ""]
                line_count = max(1, len(lines))
                max_line = max(lines, key=len) if lines else txt

                if mode == "redact":
                    pad_x = max(0.6, base * 0.15, w * 0.06)
                    pad_y = max(0.6, base * 0.20, h * 0.08)
                    est_char_w = base * 0.55
                    est_line_h = base * 1.05
                    cap_x = 10.0
                    cap_y = 10.0
                    mult_x = 0.22
                    mult_y = 0.18
                else:
                    pad_x = max(0.4, base * 0.10, w * 0.04)
                    pad_y = max(0.4, base * 0.14, h * 0.06)
                    est_char_w = base * 0.50
                    est_line_h = base * 1.02
                    cap_x = 7.0
                    cap_y = 7.0
                    mult_x = 0.18
                    mult_y = 0.14

                est_text_w = len(max_line) * est_char_w
                est_text_h = line_count * est_line_h

                extra_x = max(0.0, (est_text_w - w) / 2.0)
                extra_y = max(0.0, (est_text_h - h) / 2.0)

                pad_x = min(pad_x + extra_x * mult_x, cap_x)
                pad_y = min(pad_y + extra_y * mult_y, cap_y)

                # Narrow table cells: horizontal redact padding was still eating the vertical grid
                # between columns (e.g. List Price | Discount). Keep pad_x tiny vs cell width.
                if mode == "redact" and w < 130 and h < 56:
                    pad_x = min(
                        pad_x,
                        max(0.08, min(0.5, w * 0.014, base * 0.04)),
                    )
                    pad_y = min(
                        pad_y,
                        max(0.18, min(0.95, h * 0.1, base * 0.075)),
                    )
                if mode == "redact" and w < 75:
                    pad_x = min(pad_x, max(0.06, w * 0.012))

                # Tall OCR/table boxes were creating oversized white redaction strips in preview.
                # Keep redact/text padding tight for column-like boxes so the original background stays intact.
                if h > max(42.0, w * 1.7):
                    if mode == "redact":
                        pad_x = min(pad_x, max(1.2, w * 0.025))
                        pad_y = min(pad_y, max(1.0, base * 0.10, h * 0.025))
                    else:
                        pad_x = min(pad_x, max(0.8, w * 0.02))
                        pad_y = min(pad_y, max(0.8, base * 0.08, h * 0.02))

                if mode == "text":
                    return fitz.Rect(raw.x0, raw.y0 - pad_y, raw.x1 + pad_x, raw.y1 + pad_y)
                return fitz.Rect(raw.x0 - pad_x, raw.y0 - pad_y, raw.x1 + pad_x, raw.y1 + pad_y)

            valid_ops: list[tuple[EditItem, fitz.Rect, fitz.Rect | None]] = []
            for edit in edits:
                rr = _clip_rect_to_page(page, _padded_rect(edit, "redact"))
                if rr is None:
                    continue
                text = (edit.text or "").strip()
                ir: fitz.Rect | None = None
                if text:
                    ins_e = _edit_for_insert_text(edit)
                    ir = _clip_rect_to_page(page, _padded_rect(ins_e, "text"))
                    if ir is None:
                        continue
                valid_ops.append((edit, rr, ir))

            for _edit, rr, _ir in valid_ops:
                fill_rgb = _sample_background_fill_rgb(page, rr)
                page.add_redact_annot(rr, fill=fill_rgb)

            page.apply_redactions(
                images=fitz.PDF_REDACT_IMAGE_PIXELS,
                text=fitz.PDF_REDACT_TEXT_REMOVE,
            )

            for edit, _rr, ir in valid_ops:
                text = (edit.text or "").strip()
                if not text or ir is None:
                    continue

                rect = ir
                text_rect = rect
                font_name = _map_font_for_fitz(edit.font)
                font_size = _insert_font_size_for_rect(edit, ir)
                safe_color = (edit.color or "#000000").strip() or "#000000"
                _insert_textbox_fit_try_font_chain(
                    page, rect, text, font_name, font_size, safe_color, edit.align
                )

                if edit.is_underline or edit.is_strike:
                    ch = safe_color.lstrip("#")
                    if len(ch) == 6:
                        line_color = tuple(int(ch[i : i + 2], 16) / 255.0 for i in (0, 2, 4))
                    else:
                        line_color = (0.0, 0.0, 0.0)

                    orig = text_rect
                    font_sz = font_size
                    font_name = _map_font_for_fitz(edit.font)
                    text_content = (edit.text or "").rstrip("\n")

                    line_height = font_sz * 1.05
                    line_w = max(0.8, font_sz * 0.06)

                    lines = text_content.splitlines() or [""]
                    for idx, line in enumerate(lines):
                        try:
                            line_width = fitz.get_text_length(
                                line, fontname=font_name, fontsize=font_sz
                            )
                        except Exception:
                            line_width = min(orig.width, font_sz * max(1, len(line)) * 0.55)

                        if edit.align == "right":
                            lx0 = orig.x1 - line_width
                        elif edit.align == "center":
                            lx0 = orig.x0 + max(0, (orig.width - line_width) / 2)
                        else:
                            lx0 = orig.x0
                        lx1 = lx0 + line_width

                        baseline_y = orig.y0 + (idx + 1) * line_height

                        if edit.is_underline:
                            y_pos = baseline_y + 1.0
                            page.draw_line(
                                fitz.Point(lx0, y_pos),
                                fitz.Point(lx1, y_pos),
                                color=line_color,
                                width=line_w,
                            )

                        if edit.is_strike:
                            y_pos = baseline_y - (font_sz * 0.45)
                            page.draw_line(
                                fitz.Point(lx0, y_pos),
                                fitz.Point(lx1, y_pos),
                                color=line_color,
                                width=line_w,
                            )

        out_path = output_path(payload.file_id)
        try:
            doc.set_metadata(meta_backup)
        except Exception:
            pass
        ensure_output_path(payload.file_id)
        doc.save(out_path)

    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Save failed: {e!s}") from e
    finally:
        if doc is not None:
            try:
                doc.close()
            except Exception:
                pass

    try:
        preview_doc = fitz_open_workspace_pdf(output_path(payload.file_id))
        _write_workspace_page_previews(payload.file_id, preview_doc)
        preview_doc.close()
    except Exception:
        pass

    return {"download_url": f"/download/{payload.file_id}"}


@app.post("/set_password")
async def set_pdf_password(req: PasswordRequest) -> dict[str, str]:
    pw = (req.password or "").strip()
    if not pw:
        raise HTTPException(status_code=400, detail="Password is required")
    path = output_path(req.file_id)
    if not path.exists():
        path = input_path(req.file_id)
        if not path.exists():
            raise HTTPException(status_code=404, detail="File not found")

    try:
        try:
            doc = fitz.open(path)
        except Exception:
            try:
                doc = fitz.open(path, password=pw)
            except Exception as open_err:
                raise HTTPException(
                    status_code=401,
                    detail="Cannot open PDF. If it is already password-protected, enter the current password you used before.",
                ) from open_err
        if doc.is_encrypted and not doc.authenticate(pw):
            _safe_fitz_close(doc)
            raise HTTPException(
                status_code=401,
                detail="PDF is already password-protected. Enter the same password you used before, or unlock the file first.",
            )
        tmp_path = path.with_suffix(".tmp.pdf")
        try:
            doc.save(
                tmp_path,
                encryption=fitz.PDF_ENCRYPT_AES_256,
                owner_pw=pw,
                user_pw=pw,
                garbage=4,
                deflate=True,
            )
        except Exception as save_err:
            _safe_fitz_close(doc)
            err_msg = str(save_err).lower()
            if "signature" in err_msg or "signed" in err_msg:
                raise HTTPException(
                    status_code=400,
                    detail="This PDF is digitally signed or restricted; encryption cannot be applied. Try a copy without signatures, or use another PDF.",
                ) from save_err
            # PyMuPDF often raises this vague message on signed / restricted PDFs — not that the user "lied" about encryption.
            if "document closed or encrypted" in err_msg:
                raise HTTPException(
                    status_code=400,
                    detail="This PDF cannot be saved with a new password (often digitally signed or permission-locked). Use Print to PDF or Save As a copy, then set the password on that copy.",
                ) from save_err
            raise HTTPException(status_code=500, detail=str(save_err)) from save_err
        _safe_fitz_close(doc)
        out = output_path(req.file_id)
        shutil.move(str(tmp_path), str(out))
        # Previously only edited.pdf was encrypted; input.pdf stayed open — preview looked "unlocked".
        shutil.copy2(str(out), str(input_path(req.file_id)))
        # Do NOT persist password; require user to unlock to edit/preview.
        pwp = _workspace_user_password_path(req.file_id)
        if pwp.exists():
            pwp.unlink()
    except HTTPException:
        raise
    except Exception as e:
        el = str(e).lower()
        if "document closed or encrypted" in el:
            raise HTTPException(
                status_code=400,
                detail="This PDF cannot be saved with a new password (often digitally signed or permission-locked). Use Print to PDF or Save As a copy, then set the password on that copy.",
            ) from e
        raise HTTPException(status_code=500, detail=str(e))

    return {"status": "success", "message": "Password protected successfully"}


def _safe_download_filename(name: str | None, default: str = "edited.pdf") -> str:
    if not name or not str(name).strip():
        return default
    base = Path(str(name).strip()).name
    if not base.lower().endswith(".pdf"):
        base = f"{base}.pdf"
    # Windows-forbidden + path separators
    for c in '<>:"/\\|?*\x00':
        base = base.replace(c, "_")
    return base[:200] if len(base) > 200 else base


@app.get("/download/{file_id}")
async def download_pdf(file_id: str, filename: str | None = None) -> FileResponse:
    path = output_path(file_id)
    if not path.exists():
        path = input_path(file_id)
    if not path.exists():
        raise HTTPException(status_code=404, detail="Edited PDF not found")
    fname = _safe_download_filename(filename)
    return FileResponse(path=path, filename=fname, media_type="application/pdf")


@app.delete("/file/{file_id}")
async def delete_file_workspace(file_id: str) -> dict[str, Any]:
    """Remove the server-side workspace folder for this upload (PDFs, previews, stamps)."""
    if not re.fullmatch(
        r"[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}",
        file_id,
    ):
        raise HTTPException(status_code=400, detail="Invalid file id")
    d = (WORK_DIR / file_id).resolve()
    if WORK_DIR.resolve() != d.parent:
        raise HTTPException(status_code=400, detail="Invalid file id")
    if d.is_dir():
        shutil.rmtree(d, ignore_errors=True)
    return {"ok": True}


def _pdf_to_image_zip_bytes(file_id: str, kind: str) -> tuple[bytes, str, str]:
    """kind: jpeg | png | webp — returns (zip_bytes, zip_filename, media_type)."""
    path = output_path(file_id)
    if not path.exists():
        path = input_path(file_id)
    if not path.exists():
        raise HTTPException(status_code=404, detail="File not found")

    ext_map = {"jpeg": "jpg", "png": "png", "webp": "webp"}
    label_map = {"jpeg": "jpeg", "png": "png", "webp": "webp"}
    if kind not in ext_map:
        raise HTTPException(status_code=400, detail="Invalid image format")

    doc = fitz_open_workspace_pdf(path)
    zip_buffer = io.BytesIO()

    try:
        with zipfile.ZipFile(zip_buffer, "w") as zf:
            for i in range(len(doc)):
                page = doc[i]
                pix = page.get_pixmap(matrix=fitz.Matrix(4, 4))
                if kind == "jpeg":
                    img_data = pix.tobytes("jpg")
                elif kind == "png":
                    img_data = pix.tobytes("png")
                else:
                    from PIL import Image

                    png_bytes = pix.tobytes("png")
                    im = Image.open(io.BytesIO(png_bytes))
                    if im.mode != "RGB":
                        im = im.convert("RGB")
                    out = io.BytesIO()
                    im.save(out, format="WEBP", quality=85)
                    img_data = out.getvalue()

                zf.writestr(f"page_{i + 1}.{ext_map[kind]}", img_data)
    finally:
        doc.close()

    zip_buffer.seek(0)
    zip_name = f"pages_{label_map[kind]}_{file_id[:8]}.zip"
    return zip_buffer.getvalue(), zip_name, "application/x-zip-compressed"


@app.get("/convert_to_jpeg/{file_id}")
async def convert_to_jpeg(file_id: str):
    data, zip_name, media = _pdf_to_image_zip_bytes(file_id, "jpeg")
    return StreamingResponse(
        io.BytesIO(data),
        media_type=media,
        headers={"Content-Disposition": f'attachment; filename="{zip_name}"'},
    )


@app.get("/convert_to_png/{file_id}")
async def convert_to_png(file_id: str):
    data, zip_name, media = _pdf_to_image_zip_bytes(file_id, "png")
    return StreamingResponse(
        io.BytesIO(data),
        media_type=media,
        headers={"Content-Disposition": f'attachment; filename="{zip_name}"'},
    )


@app.get("/convert_to_webp/{file_id}")
async def convert_to_webp(file_id: str):
    data, zip_name, media = _pdf_to_image_zip_bytes(file_id, "webp")
    return StreamingResponse(
        io.BytesIO(data),
        media_type=media,
        headers={"Content-Disposition": f'attachment; filename="{zip_name}"'},
    )


@app.post("/combine")
async def combine_pdfs(req: CombineRequest) -> dict[str, Any]:
    try:
        if not req.file_ids:
            return {"status": "error", "detail": "No files provided"}
            
        new_id = str(uuid.uuid4())
        out_dir = WORK_DIR / new_id
        out_dir.mkdir(parents=True, exist_ok=True)
        out_pdf = out_dir / "input.pdf"
        for fid in req.file_ids:
            if (WORK_DIR / fid / NO_WM_BACKUP_NAME).is_file():
                _carry_no_wm_backup_if_present(fid, out_dir)
                break

        result_doc = fitz.open()
        for fid in req.file_ids:
            # Check for existing combined output or fresh input
            found = False
            for path in [output_path(fid), input_path(fid)]:
                if path.exists():
                    try:
                        next_doc = fitz_open_workspace_pdf(path)
                        result_doc.insert_pdf(next_doc)
                        next_doc.close()
                        found = True
                        break
                    except Exception as e:
                        print(f"Error opening {fid}: {e}")
                        continue
            if not found:
                print(f"Warning: File ID {fid} not found on disk.")
                
        if result_doc.page_count == 0:
            result_doc.close()
            raise HTTPException(status_code=400, detail="Merging failed: No valid PDF content found in chosen files.")

        result_doc.save(str(out_pdf))
        result_doc.close()
        
        return {"file_id": new_id, "status": "success", "size": os.path.getsize(out_pdf)}
    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Combine error: {str(e)}")


@app.post("/compress")
async def compress_pdf(req: dict = {"file_id": "", "quality": 60}) -> dict[str, Any]:
    file_id = req.get("file_id")
    quality = req.get("quality", 60)
    
    if not file_id:
        raise HTTPException(status_code=400, detail="Missing file_id")
        
    source = output_path(file_id)
    if not source.exists():
        source = input_path(file_id)
        
    if not source.exists():
        raise HTTPException(status_code=404, detail="File not found")
        
    new_id = str(uuid.uuid4())
    out_dir = WORK_DIR / new_id
    out_dir.mkdir(parents=True, exist_ok=True)
    out_pdf = out_dir / "input.pdf"
    _carry_no_wm_backup_if_present(file_id, out_dir)

    try:
        doc = fitz_open_workspace_pdf(source)
        
        # Intense Optimization: Only process images if they actually exist and have real size
        for page in doc:
            img_list = page.get_images()
            for img in img_list:
                xref = img[0]
                try:
                    orig_img = doc.extract_image(xref)
                    if not orig_img: continue
                    orig_size = orig_img["size"]
                    
                    if orig_size < 5120: # skip very tiny icons (< 5KB)
                        continue
                        
                    pix = fitz.Pixmap(doc, xref)
                    # Use passed quality value (progressive reduction)
                    img_data = pix.tobytes("jpeg", quality)
                    
                    if len(img_data) < orig_size:
                        page.replace_image(xref, stream=img_data)
                    pix = None
                except Exception as e:
                    print(f"Skipping image {xref}: {e}")
                    continue
        
        # Save with garbage collection
        doc.save(str(out_pdf), garbage=4, deflate=True, clean=True)
        doc.close()
        
        # FINAL SIZE GUARD: If compressed file is larger than original, revert!
        final_size = os.path.getsize(out_pdf)
        original_size = os.path.getsize(source)
        
        if final_size > original_size:
            # Revert to original file bytes to ensure size never increases
            shutil.copy2(str(source), str(out_pdf))
            final_size = original_size
        
        return {"file_id": new_id, "status": "success", "size": final_size}
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Compression failed: {str(e)}")


@app.post("/upload_multiple")
async def upload_multiple(files: list[UploadFile] = File(...)) -> dict[str, Any]:
    file_id = str(uuid.uuid4())
    folder = WORK_DIR / file_id
    folder.mkdir(parents=True, exist_ok=True)
    destination = folder / "input.pdf"
    
    doc = fitz.open()
    
    for file in files:
        temp_img = folder / f"temp_{uuid.uuid4()}_{file.filename}"
        with temp_img.open("wb") as buffer:
            shutil.copyfileobj(file.file, buffer)
            
        try:
            # Check if it's an image
            img_doc = fitz_open_workspace_pdf(temp_img)
            pdf_bytes = img_doc.convert_to_pdf()
            img_doc.close()
            
            img_page = fitz.open("pdf", pdf_bytes)
            doc.insert_pdf(img_page)
            img_page.close()
        except Exception:
            # If it's already a PDF, just insert it
            if file.filename.lower().endswith(".pdf"):
                pdf_doc = fitz_open_workspace_pdf(temp_img)
                doc.insert_pdf(pdf_doc)
                pdf_doc.close()
        
        # Cleanup temp image
        if temp_img.exists():
            os.remove(temp_img)
            
    doc.save(destination)
    doc.close()
    
    return {"file_id": file_id, "filename": "Combined.pdf", "needs_password": False}


@app.post("/split_pages")
async def split_pages(req: SplitRequest) -> dict[str, Any]:
    path = resolve_pdf_path(req.file_id)
    doc = fitz_open_workspace_pdf(path)
    try:
        n = doc.page_count
        if not req.page_indices:
            raise HTTPException(status_code=400, detail="No pages selected")
        indices: list[int] = []
        for p in req.page_indices:
            if p < 1 or p > n:
                raise HTTPException(status_code=400, detail=f"Invalid page number: {p}")
            indices.append(p - 1)
        new_id, out_pdf = _new_work_file_id()
        _carry_no_wm_backup_if_present(req.file_id, out_pdf.parent)
        new_doc = fitz.open()
        for i in indices:
            new_doc.insert_pdf(doc, from_page=i, to_page=i)
        new_doc.save(str(out_pdf), garbage=4, deflate=True)
        new_doc.close()
    finally:
        doc.close()
    return {"file_id": new_id, "status": "success", "size": os.path.getsize(out_pdf)}


@app.post("/rotate_pages")
async def rotate_pages(req: RotateRequest) -> dict[str, Any]:
    if req.angle not in (90, 180, 270):
        raise HTTPException(status_code=400, detail="angle must be 90, 180, or 270")
    path = resolve_pdf_path(req.file_id)
    doc = fitz_open_workspace_pdf(path)
    try:
        n = doc.page_count
        if req.pages is None or len(req.pages) == 0:
            targets = list(range(n))
        else:
            targets = []
            for p in req.pages:
                if p < 1 or p > n:
                    raise HTTPException(status_code=400, detail=f"Invalid page: {p}")
                targets.append(p - 1)
        for pi in targets:
            pg = doc[pi]
            cur = int(pg.rotation)
            pg.set_rotation((cur + req.angle) % 360)
        new_id, out_pdf = _new_work_file_id()
        _carry_no_wm_backup_if_present(req.file_id, out_pdf.parent)
        doc.save(str(out_pdf), garbage=4, deflate=True)
    finally:
        doc.close()
    return {"file_id": new_id, "status": "success", "size": os.path.getsize(out_pdf)}


@app.post("/crop_page")
async def crop_page(req: CropRequest) -> dict[str, Any]:
    path = resolve_pdf_path(req.file_id)
    doc = fitz_open_workspace_pdf(path)
    try:
        n = doc.page_count
        if req.page < 1 or req.page > n:
            raise HTTPException(status_code=400, detail="Invalid page")
        pages_to_crop = list(range(n)) if req.all_pages else [req.page - 1]
        for pi in pages_to_crop:
            page = doc[pi]
            m = page.mediabox
            r = fitz.Rect(m)
            r.x0 += float(req.left)
            r.y0 += float(req.top)
            r.x1 -= float(req.right)
            r.y1 -= float(req.bottom)
            if r.width < 12 or r.height < 12:
                raise HTTPException(status_code=400, detail="Crop removes too much; reduce margins")
            page.set_cropbox(r)
            page.set_mediabox(r)
        new_id, out_pdf = _new_work_file_id()
        _carry_no_wm_backup_if_present(req.file_id, out_pdf.parent)
        doc.save(str(out_pdf), garbage=4, deflate=True)
    finally:
        doc.close()
    return {"file_id": new_id, "status": "success", "size": os.path.getsize(out_pdf)}


@app.post("/crop_page_norm")
async def crop_page_norm(req: CropNormRequest) -> dict[str, Any]:
    for v in (req.x0, req.y0, req.x1, req.y1):
        if v < 0.0 or v > 1.0:
            raise HTTPException(status_code=400, detail="Crop values must be between 0 and 1")
    if req.x0 >= req.x1 or req.y0 >= req.y1:
        raise HTTPException(status_code=400, detail="Invalid crop rectangle")
    if (req.x1 - req.x0) < 0.03 or (req.y1 - req.y0) < 0.03:
        raise HTTPException(status_code=400, detail="Crop area too small")

    path = resolve_pdf_path(req.file_id)
    doc = fitz_open_workspace_pdf(path)
    try:
        n = doc.page_count
        if req.page < 1 or req.page > n:
            raise HTTPException(status_code=400, detail="Invalid page")
        pages_to_crop = list(range(n)) if req.all_pages else [req.page - 1]
        for pi in pages_to_crop:
            page = doc[pi]
            m = page.mediabox
            r = fitz.Rect(
                m.x0 + req.x0 * m.width,
                m.y0 + req.y0 * m.height,
                m.x0 + req.x1 * m.width,
                m.y0 + req.y1 * m.height,
            )
            if r.width < 12 or r.height < 12:
                raise HTTPException(status_code=400, detail="Crop removes too much on a page")
            page.set_cropbox(r)
            page.set_mediabox(r)
        new_id, out_pdf = _new_work_file_id()
        _carry_no_wm_backup_if_present(req.file_id, out_pdf.parent)
        doc.save(str(out_pdf), garbage=4, deflate=True)
    finally:
        doc.close()
    return {"file_id": new_id, "status": "success", "size": os.path.getsize(out_pdf)}


def _pristine_pdf_for_watermark_backup(path: Path) -> Path:
    """Prefer longest-lived pre-watermark bytes in this workspace folder."""
    candidate = path.parent / NO_WM_BACKUP_NAME
    if candidate.exists():
        return candidate
    return path


@app.post("/watermark")
async def watermark_pdf(req: WatermarkRequest) -> dict[str, Any]:
    text = (req.text or "Draft").strip() or "Draft"
    try:
        pos = _normalize_watermark_position(req.position)
    except ValueError as e:
        raise HTTPException(status_code=400, detail=str(e)) from e
    path = resolve_pdf_path(req.file_id)
    backup_source = _pristine_pdf_for_watermark_backup(path)
    doc = fitz_open_workspace_pdf(path)
    try:
        fs = float(req.font_size)
        op = float(req.opacity)
        for i in range(len(doc)):
            _add_watermark(doc[i], text, op, fs, pos)
        new_id, out_pdf = _new_work_file_id()
        out_dir = WORK_DIR / new_id
        doc.save(str(out_pdf), garbage=4, deflate=True)
        shutil.copy2(backup_source, out_dir / NO_WM_BACKUP_NAME)
    finally:
        doc.close()
    return {"file_id": new_id, "status": "success", "size": os.path.getsize(out_pdf)}


@app.post("/remove_watermark")
async def remove_watermark(req: RemoveWatermarkRequest) -> dict[str, Any]:
    folder = WORK_DIR / req.file_id
    clean = folder / NO_WM_BACKUP_NAME
    if not clean.exists():
        raise HTTPException(
            status_code=404,
            detail=(
                "No backup copy found. Open Watermark and tap Apply once (any text) to create a backup, "
                "then Remove will work. Or upload the PDF again if it was never watermarked on this server."
            ),
        )
    inp = input_path(req.file_id)
    outp = output_path(req.file_id)
    shutil.copy2(clean, inp)
    shutil.copy2(clean, outp)
    for pattern in ("preview-*.png", "preview_edited-*.png"):
        for p in folder.glob(pattern):
            try:
                p.unlink()
            except OSError:
                pass
    sz = os.path.getsize(inp)
    return {"file_id": req.file_id, "status": "success", "size": sz}


@app.get("/pdf_metadata/{file_id}")
async def get_pdf_metadata(file_id: str) -> dict[str, Any]:
    path = resolve_pdf_path(file_id)
    doc = fitz_open_workspace_pdf(path)
    try:
        meta = dict(doc.metadata)
    finally:
        doc.close()
    return {"file_id": file_id, "metadata": meta}


@app.post("/pdf_metadata")
async def update_pdf_metadata(req: MetadataUpdateRequest) -> dict[str, Any]:
    path = resolve_pdf_path(req.file_id)
    doc = fitz_open_workspace_pdf(path)
    try:
        if req.strip:
            doc.set_metadata({})
        else:
            m = dict(doc.metadata)
            if req.title is not None:
                m["title"] = req.title
            if req.author is not None:
                m["author"] = req.author
            doc.set_metadata(m)
        new_id, out_pdf = _new_work_file_id()
        _carry_no_wm_backup_if_present(req.file_id, out_pdf.parent)
        doc.save(str(out_pdf), garbage=4, deflate=True)
    finally:
        doc.close()
    return {"file_id": new_id, "status": "success", "size": os.path.getsize(out_pdf)}


@app.get("/export_page_image/{file_id}/{page_number}")
async def export_page_image(file_id: str, page_number: int, format: str = "png"):
    fmt = (format or "png").lower()
    if fmt not in ("png", "jpeg", "jpg"):
        raise HTTPException(status_code=400, detail="format must be png or jpeg")
    path = resolve_pdf_path(file_id)
    doc = fitz_open_workspace_pdf(path)
    try:
        if page_number < 1 or page_number > doc.page_count:
            raise HTTPException(status_code=404, detail="Page not found")
        page = doc[page_number - 1]
        pix = page.get_pixmap(matrix=fitz.Matrix(3, 3), alpha=False)
        if fmt in ("jpeg", "jpg"):
            data = pix.tobytes("jpg")
            media = "image/jpeg"
            ext = "jpg"
        else:
            data = pix.tobytes("png")
            media = "image/png"
            ext = "png"
    finally:
        doc.close()
    fname = f"page_{page_number}.{ext}"
    return Response(
        content=data,
        media_type=media,
        headers={"Content-Disposition": f'attachment; filename="{fname}"'},
    )


@app.get("/export_text/{file_id}")
async def export_text(file_id: str):
    path = resolve_pdf_path(file_id)
    doc = fitz_open_workspace_pdf(path)
    try:
        parts: list[str] = []
        for i in range(len(doc)):
            parts.append(doc[i].get_text())
    finally:
        doc.close()
    body = "\n\n".join(parts).strip() + "\n"
    safe = _safe_download_filename(None, default="export.txt")
    return Response(
        content=body.encode("utf-8"),
        media_type="text/plain; charset=utf-8",
        headers={"Content-Disposition": f'attachment; filename="{safe}"'},
    )


@app.get("/export_docx/{file_id}")
async def export_docx(file_id: str):
    from docx import Document as DocxDocument

    path = resolve_pdf_path(file_id)
    doc = fitz_open_workspace_pdf(path)
    try:
        full_text: list[str] = []
        for i in range(len(doc)):
            t = doc[i].get_text().strip()
            if t:
                full_text.append(t)
    finally:
        doc.close()
    merged = "\n\n".join(full_text)
    d = DocxDocument()
    for block in merged.split("\n\n"):
        if block.strip():
            d.add_paragraph(block)
    buf = io.BytesIO()
    d.save(buf)
    buf.seek(0)
    return StreamingResponse(
        buf,
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        headers={"Content-Disposition": 'attachment; filename="export.docx"'},
    )


def _signature_bg_remove_pil(raw: bytes) -> bytes:
    """Lightweight background removal for signatures (white / light paper)."""
    import numpy as np
    from PIL import Image

    im = Image.open(io.BytesIO(raw)).convert("RGBA")
    arr = np.asarray(im).copy()
    rgb = arr[:, :, :3].astype(np.float32)
    edge = np.concatenate(
        [rgb[0, :, :], rgb[-1, :, :], rgb[:, 0, :], rgb[:, -1, :]], axis=0
    )
    bg = np.median(edge.reshape(-1, 3), axis=0)
    d = np.sqrt(np.sum((rgb - bg.reshape(1, 1, 3)) ** 2, axis=2))
    lum = np.mean(rgb, axis=2)
    paper = (d < 45.0) & (lum > 188.0)
    near_white = (arr[:, :, 0] > 247) & (arr[:, :, 1] > 247) & (arr[:, :, 2] > 247)
    arr[:, :, 3] = np.where(paper | near_white, 0, 255).astype(np.uint8)
    out = Image.fromarray(arr, "RGBA")
    buf = io.BytesIO()
    out.save(buf, format="PNG", optimize=True)
    return buf.getvalue()


def remove_signature_background(raw: bytes, *, fast: bool = False) -> bytes:
    """If ``fast`` is True, use PIL-only removal (no rembg). Otherwise prefer rembg, then PIL."""
    if fast:
        return _signature_bg_remove_pil(raw)
    try:
        from rembg import remove as rembg_remove

        out = rembg_remove(raw)
        if out and len(out) > 64:
            return out
    except Exception:
        pass
    return _signature_bg_remove_pil(raw)


@app.post("/upload_stamp")
async def upload_stamp(
    file: UploadFile = File(...),
    fast: str = Form("false"),
) -> dict[str, Any]:
    from PIL import Image

    raw = await file.read()
    if len(raw) > 12 * 1024 * 1024:
        raise HTTPException(status_code=400, detail="Image too large (max 12 MB)")
    if len(raw) < 32:
        raise HTTPException(status_code=400, detail="Invalid image")

    fast_bg = fast.lower() in ("1", "true", "yes", "on")

    try:
        processed = remove_signature_background(raw, fast=fast_bg)
    except Exception as e:
        raise HTTPException(status_code=400, detail=f"Could not process image: {e}") from e

    stamp_id = str(uuid.uuid4())
    out_path = STAMP_DIR / f"{stamp_id}.png"
    out_path.write_bytes(processed)

    im = Image.open(io.BytesIO(processed))
    w, h = im.size
    return {"stamp_id": stamp_id, "width": w, "height": h}


@app.get("/stamp_preview/{stamp_id}")
async def stamp_preview(stamp_id: str) -> FileResponse:
    p = STAMP_DIR / f"{stamp_id}.png"
    if not p.exists():
        raise HTTPException(status_code=404, detail="Stamp not found")
    return FileResponse(p, media_type="image/png", filename="stamp.png")


def _insert_image_preserving_alpha(
    page: fitz.Page,
    rect: fitz.Rect,
    image_path: Path,
    *,
    keep_proportion: bool = True,
) -> None:
    """
    Insert PNG/JPEG via Pixmap so RGBA alpha becomes a proper PDF soft mask.

    Using insert_image(filename=...) alone can drop or mishandle transparency; many
    viewers then show transparent pixels as solid black.
    """
    pix = fitz.Pixmap(str(image_path))
    try:
        page.insert_image(rect, pixmap=pix, keep_proportion=keep_proportion)
    finally:
        pix = None


@app.post("/apply_stamp")
async def apply_stamp(req: ApplyStampRequest) -> dict[str, Any]:
    for v in (req.x0, req.y0, req.x1, req.y1):
        if v < 0.0 or v > 1.0:
            raise HTTPException(status_code=400, detail="Stamp box must use values 0–1")
    if req.x0 >= req.x1 or req.y0 >= req.y1:
        raise HTTPException(status_code=400, detail="Invalid stamp rectangle")
    if (req.x1 - req.x0) < 0.02 or (req.y1 - req.y0) < 0.02:
        raise HTTPException(status_code=400, detail="Stamp area too small")

    stamp_path = STAMP_DIR / f"{req.stamp_id}.png"
    if not stamp_path.exists():
        raise HTTPException(status_code=404, detail="Stamp expired or not found")

    path = resolve_pdf_path(req.file_id)
    doc = fitz_open_workspace_pdf(path)
    try:
        n = doc.page_count
        if req.page < 1 or req.page > n:
            raise HTTPException(status_code=400, detail="Invalid page")
        page = doc[req.page - 1]
        m = page.mediabox
        r = fitz.Rect(
            m.x0 + req.x0 * m.width,
            m.y0 + req.y0 * m.height,
            m.x0 + req.x1 * m.width,
            m.y0 + req.y1 * m.height,
        )
        _insert_image_preserving_alpha(page, r, stamp_path, keep_proportion=True)
        new_id, out_pdf = _new_work_file_id()
        _carry_no_wm_backup_if_present(req.file_id, out_pdf.parent)
        doc.save(str(out_pdf), garbage=4, deflate=True)
    finally:
        doc.close()
    return {"file_id": new_id, "status": "success", "size": os.path.getsize(out_pdf)}


async def _sign_pdf_bytes_pkcs12(pdf_bytes: bytes, p12_bytes: bytes, password: str) -> bytes:
    """Add a CMS/PKCS#7 digital signature using a PKCS#12 identity (requires pyhanko)."""
    from io import BytesIO

    try:
        from pyhanko.pdf_utils.incremental_writer import IncrementalPdfFileWriter
        from pyhanko.sign import signers
        from pyhanko.sign.fields import SigFieldSpec
        from pyhanko.sign.signers.pdf_signer import PdfSignatureMetadata, PdfSigner
    except ImportError as e:
        raise HTTPException(
            status_code=501,
            detail="Certificate signing requires pyhanko. Install: pip install pyhanko",
        ) from e

    passphrase = password.encode("utf-8") if password else None
    try:
        signer = signers.SimpleSigner.load_pkcs12_data(p12_bytes, (), passphrase)
    except Exception as e:
        raise HTTPException(status_code=400, detail=f"Could not load PKCS#12: {e}") from e

    field_name = "PdfEditorApprovalSig"
    meta = PdfSignatureMetadata(
        field_name=field_name,
        reason="Approved with certificate",
        location="PDF Editor Tool",
    )
    field_spec = SigFieldSpec(sig_field_name=field_name, on_page=0)

    bio = BytesIO(pdf_bytes)
    writer = IncrementalPdfFileWriter(bio, strict=False)
    pdf_signer = PdfSigner(meta, signer, new_field_spec=field_spec)
    try:
        out_stream = await pdf_signer.async_sign_pdf(writer, existing_fields_only=False)
    except Exception as e:
        raise HTTPException(status_code=400, detail=f"Signing failed: {e}") from e

    out_stream.seek(0)
    return out_stream.read()


@app.post("/sign_pkcs12")
async def sign_pkcs12(
    file_id: str = Form(...),
    password: str = Form(""),
    p12: UploadFile = File(...),
) -> dict[str, Any]:
    """Sign the current PDF with a PKCS#12 (.p12 / .pfx) certificate; returns a new ``file_id``."""
    path = resolve_pdf_path(file_id)
    pdf_bytes = path.read_bytes()
    p12_bytes = await p12.read()
    if len(p12_bytes) < 32:
        raise HTTPException(status_code=400, detail="Invalid PKCS#12 file")

    signed = await _sign_pdf_bytes_pkcs12(pdf_bytes, p12_bytes, password)
    new_id, out_path = _new_work_file_id()
    _carry_no_wm_backup_if_present(file_id, out_path.parent)
    out_path.write_bytes(signed)
    return {"file_id": new_id, "size": out_path.stat().st_size, "status": "success"}


# Mount static files last so /static/* is never shadowed by route registration quirks on Linux hosts.
_static_root = BASE_DIR / "static"
app.mount(
    "/static",
    StaticFiles(directory=str(_static_root), check_dir=True),
    name="static",
)
