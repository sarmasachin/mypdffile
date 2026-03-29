from __future__ import annotations

import shutil
import uuid
from pathlib import Path
from collections import defaultdict
from typing import Any

import zipfile
import io
import os
import fitz  # PyMuPDF
from fastapi import FastAPI, File, Form, HTTPException, UploadFile
from fastapi.responses import FileResponse, HTMLResponse, Response, StreamingResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates
from pydantic import BaseModel
from starlette.middleware.proxy_headers import ProxyHeadersMiddleware
from starlette.requests import Request

BASE_DIR = Path(__file__).resolve().parent
WORK_DIR = BASE_DIR / "work"
WORK_DIR.mkdir(exist_ok=True)
STAMP_DIR = WORK_DIR / "stamps"
STAMP_DIR.mkdir(exist_ok=True)

app = FastAPI(title="PDF Editor Tool")
# Behind Render / other reverse proxies so request.url_for / schemes stay correct.
app.add_middleware(ProxyHeadersMiddleware, trusted_hosts="*")
templates = Jinja2Templates(directory=str(BASE_DIR / "templates"))


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


class WatermarkRequest(BaseModel):
    file_id: str
    text: str = "Draft"
    opacity: float = 0.25


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

def preview_edited_path(file_id: str, page_number: int) -> Path:
    return WORK_DIR / file_id / f"preview_edited-{page_number}.png"


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


def _add_diagonal_watermark(page: fitz.Page, text: str, opacity: float) -> None:
    r = page.rect
    c = fitz.Point((r.x0 + r.x1) / 2, (r.y0 + r.y1) / 2)
    gray = max(0.72, min(0.92, 0.78 + (1.0 - min(max(opacity, 0.05), 1.0)) * 0.12))
    color = (gray, gray, gray)
    tw = fitz.TextWriter(r)
    tw.append(c, text, fontsize=44)
    tw.write_text(page, color=color, morph=(c, fitz.Matrix(1.15, 1.15).prerotate(45)))


def _insert_textbox_fit(
    page: fitz.Page,
    rect: fitz.Rect,
    text: str,
    font_name: str,
    start_size: float,
    color_hex: str = "#000000",
    align: str = "left",
) -> None:
    import copy
    
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

    expanded_rect = fitz.Rect(rect.x0, rect.y0, rect.x1 + 400, rect.y1 + 400)
    
    page.insert_textbox(
        expanded_rect,
        text,
        fontsize=float(start_size),
        fontname=font_name,
        color=(r, g, b),
        align=align_code,
    )


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
    
    try:
        doc = fitz.open(source)
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
            doc = fitz.open(destination)
            is_encrypted = doc.is_encrypted
            doc.close()
        else:
            # Efficient image to PDF conversion (avoids size bloat)
            img_doc = fitz.open(temp_path)
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
        
        # Save a decrypted version to overwrite input.pdf
        tmp_path = path.with_suffix('.unlocked.pdf')
        doc.save(tmp_path)
        doc.close()
        
        shutil.move(str(tmp_path), str(path))
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
async def page_preview(file_id: str, page_number: int) -> FileResponse:
    path = ensure_file(file_id)
    doc = fitz.open(path)

    if page_number < 1 or page_number > doc.page_count:
        doc.close()
        raise HTTPException(status_code=404, detail="Page not found")

    out_path = preview_path(file_id, page_number)
    if not out_path.exists():
        out_path.parent.mkdir(parents=True, exist_ok=True)
        page = doc[page_number - 1]
        pix = page.get_pixmap(matrix=fitz.Matrix(2, 2), alpha=False)
        pix.save(str(out_path))

    doc.close()
    return FileResponse(path=out_path, media_type="image/png")


@app.get("/preview_edited/{file_id}/{page_number}")
async def page_preview_edited(file_id: str, page_number: int) -> FileResponse:
    out_path = preview_edited_path(file_id, page_number)
    if out_path.exists():
        return FileResponse(path=out_path, media_type="image/png")
    
    # Fallback: generate from edited PDF if image not cached yet
    path = output_path(file_id)
    if not path.exists():
        raise HTTPException(status_code=404, detail="Edited PDF not found")
    
    try:
        doc = fitz.open(path)
        if page_number < 1 or page_number > doc.page_count:
            doc.close()
            raise HTTPException(status_code=404, detail="Page not found")
        page = doc[page_number - 1]
        pix = page.get_pixmap(matrix=fitz.Matrix(4, 4), alpha=False)
        pix.save(out_path)
        doc.close()
    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Preview generation failed: {e}")
    
    return FileResponse(path=out_path, media_type="image/png")

@app.get("/analyze/{file_id}")
async def analyze_pdf(file_id: str) -> dict[str, Any]:
    path = ensure_file(file_id)
    doc = fitz.open(path)
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
        for block_index, block in enumerate(text_dict.get("blocks", [])):
            if block.get("type") != 0:
                continue
            
            for line_index, line in enumerate(block.get("lines", [])):
                spans = line.get("spans", [])
                if not spans:
                    continue
                
                # Group text spans that are horizontally close into chunks.
                # If there's a layout gap (like between a label and its value in a table),
                # it separates them into isolated editable boxes, preserving original coordinates accurately!
                chunks = []
                current_chunk = []
                
                for span in spans:
                    t = (span.get("text") or "").strip()
                    if not t:
                        continue
                        
                    if not current_chunk:
                        current_chunk.append(span)
                    else:
                        prev_span = current_chunk[-1]
                        prev_x1 = float(prev_span.get("bbox", [0,0,0,0])[2])
                        curr_x0 = float(span.get("bbox", [0,0,0,0])[0])
                        
                        # Gap of > 12 points usually indicates a structural separation (columns, tabs)
                        if curr_x0 - prev_x1 > 12.0:
                            chunks.append(current_chunk)
                            current_chunk = [span]
                        else:
                            current_chunk.append(span)
                            
                if current_chunk:
                    chunks.append(current_chunk)
                
                for chunk_index, chunk in enumerate(chunks):
                    chunk_text = " ".join((s.get("text") or "").strip() for s in chunk if (s.get("text") or "").strip())
                    if not chunk_text:
                        continue
                        
                    # Calculate strict tight bounding box for the chunk
                    rects = []
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
                    font_name = first_span.get("font", "helv")
                    size = float(first_span.get("size", 11.0))
                    
                    items.append(
                        {
                            "id": item_id,
                            "page": page_index + 1,
                            "text": chunk_text,
                            "bbox": bbox_list,
                            "font": font_name,
                            "size": size,
                        }
                    )

    doc.close()
    return {"file_id": file_id, "pages": pages, "items": items}


@app.post("/edit")
async def edit_pdf(payload: EditRequest) -> dict[str, str]:
    path = ensure_file(payload.file_id)
    doc = fitz.open(path)

    normalized_edits = [
        item
        for item in payload.edits
        if item.text is not None
    ]

    if not normalized_edits:
        shutil.copy(path, ensure_output_path(payload.file_id))
        # Generate previews from original too
        try:
            preview_doc = fitz.open(ensure_output_path(payload.file_id))
            for pg_num in range(1, preview_doc.page_count + 1):
                pg = preview_doc[pg_num - 1]
                pix = pg.get_pixmap(matrix=fitz.Matrix(4, 4), alpha=False)
                pix.save(preview_edited_path(payload.file_id, pg_num))
            preview_doc.close()
        except Exception:
            pass
        return {"download_url": f"/download/{payload.file_id}"}

    by_page: dict[int, list[EditItem]] = defaultdict(list)
    for item in normalized_edits:
        by_page[item.page - 1].append(item)

    for page_index, edits in by_page.items():
        if page_index < 0 or page_index >= doc.page_count:
            continue
        page = doc[page_index]

        def _padded_rect(edit_item: EditItem, mode: str) -> fitz.Rect:
            """
            Text boxes ka bbox extracted tight hota hai.
            Double text/cut se bachne ke liye redaction rect ko text insert rect se
            thoda bada rakhte hain.
            """
            if mode == "redact" and edit_item.original_bbox:
                raw = fitz.Rect(edit_item.original_bbox)
            else:
                raw = fitz.Rect(edit_item.bbox)
                
            base = float(edit_item.size or 11)
            # Using both bbox geometry and detected font size makes padding work
            # even when OCR/text extraction bbox is slightly off.
            w = max(0.1, float(raw.x1 - raw.x0))
            h = max(0.1, float(raw.y1 - raw.y0))

            txt = (edit_item.text or "").strip()
            lines = [ln for ln in txt.splitlines() if ln.strip() != ""]
            line_count = max(1, len(lines))
            max_line = max(lines, key=len) if lines else txt

            if mode == "redact":
                # Redaction ko minimal rakhein, bas descenders/underlines ko cover karne ke liye.
                pad_x = max(0.6, base * 0.15, w * 0.06)
                pad_y = max(0.6, base * 0.20, h * 0.08)
                est_char_w = base * 0.55
                est_line_h = base * 1.05
                cap_x = 10.0
                cap_y = 10.0
                mult_x = 0.22
                mult_y = 0.18
            else:
                # Insert rect ko tight rakhein (neighbor text disturb na ho).
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

            return fitz.Rect(raw.x0 - pad_x, raw.y0 - pad_y, raw.x1 + pad_x, raw.y1 + pad_y)

        # 1) Redaction (purana text fully remove karne ke liye bada rect)
        for edit in edits:
            rect = _padded_rect(edit, "redact")
            page.add_redact_annot(rect, fill=(1, 1, 1))

        # Purana text/image ko redaction se remove karo,
        # warna "double/ghost" text insert ke baad bhi rahega.
        page.apply_redactions(
            images=fitz.PDF_REDACT_IMAGE_REMOVE,
            text=fitz.PDF_REDACT_TEXT_REMOVE,
        )

        # 2) Insert (text ko bbox ke andar tight rakhein)
        for edit in edits:
            text = (edit.text or "").strip()
            if not text:
                continue
                
            rect = _padded_rect(edit, "text")
            text_rect = rect  # use same rect for decorations so lines align with inserted text
            font_name = _map_font_for_fitz(edit.font)
            font_size = float(edit.size or 11)
            try:
                _insert_textbox_fit(page, rect, text, font_name, font_size, edit.color, edit.align)
            except Exception as e:
                _insert_textbox_fit(page, rect, text, "helv", font_size, edit.color, edit.align)
                
            # Render box-level text decorations (keep lines tight to actual text)
            if edit.is_underline or edit.is_strike:
                # Color parsing
                ch = edit.color.lstrip("#")
                if len(ch) == 6:
                    line_color = tuple(int(ch[i:i+2], 16) / 255.0 for i in (0, 2, 4))
                else:
                    line_color = (0.0, 0.0, 0.0)

                # Use the text insertion rect so underline width matches new text
                orig = text_rect
                font_sz = float(edit.size or 11)
                font_name = _map_font_for_fitz(edit.font)
                text_content = (edit.text or "").rstrip("\n")

                line_height = font_sz * 1.05
                line_w = max(0.8, font_sz * 0.06)

                # Draw per-line to avoid spilling outside text bounds
                lines = text_content.splitlines() or [""]
                for idx, line in enumerate(lines):
                    # Compute line width using font metrics
                    try:
                        line_width = fitz.get_text_length(line, fontname=font_name, fontsize=font_sz)
                    except Exception:
                        line_width = min(orig.width, font_sz * max(1, len(line)) * 0.55)

                    # Align start based on text alignment
                    if edit.align == "right":
                        lx0 = orig.x1 - line_width
                    elif edit.align == "center":
                        lx0 = orig.x0 + max(0, (orig.width - line_width) / 2)
                    else:
                        lx0 = orig.x0
                    lx1 = lx0 + line_width

                    # Vertical positions
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
    doc.save(out_path)
    doc.close()

    # Generate preview images immediately after saving
    try:
        preview_doc = fitz.open(out_path)
        for pg_num in range(1, preview_doc.page_count + 1):
            pg = preview_doc[pg_num - 1]
            pix = pg.get_pixmap(matrix=fitz.Matrix(4, 4), alpha=False)
            img_path = preview_edited_path(payload.file_id, pg_num)
            pix.save(img_path)
        preview_doc.close()
    except Exception:
        pass  # Preview generation failure should not block the response

    return {"download_url": f"/download/{payload.file_id}"}


@app.post("/set_password")
async def set_pdf_password(req: PasswordRequest) -> dict[str, str]:
    path = output_path(req.file_id)
    if not path.exists():
        path = input_path(req.file_id)
        if not path.exists():
            raise HTTPException(status_code=404, detail="File not found")
            
    try:
        doc = fitz.open(path)
        tmp_path = path.with_suffix('.tmp.pdf')
        doc.save(
            tmp_path, 
            encryption=fitz.PDF_ENCRYPT_AES_256, 
            owner_pw=req.password, 
            user_pw=req.password
        )
        doc.close()
        shutil.move(str(tmp_path), str(output_path(req.file_id)))
    except Exception as e:
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

    doc = fitz.open(path)
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
        
        result_doc = fitz.open()
        for fid in req.file_ids:
            # Check for existing combined output or fresh input
            found = False
            for path in [output_path(fid), input_path(fid)]:
                if path.exists():
                    try:
                        next_doc = fitz.open(str(path))
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
    
    try:
        doc = fitz.open(source)
        
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
            img_doc = fitz.open(temp_img)
            pdf_bytes = img_doc.convert_to_pdf()
            img_doc.close()
            
            img_page = fitz.open("pdf", pdf_bytes)
            doc.insert_pdf(img_page)
            img_page.close()
        except Exception:
            # If it's already a PDF, just insert it
            if file.filename.lower().endswith(".pdf"):
                pdf_doc = fitz.open(temp_img)
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
    doc = fitz.open(path)
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
    doc = fitz.open(path)
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
        doc.save(str(out_pdf), garbage=4, deflate=True)
    finally:
        doc.close()
    return {"file_id": new_id, "status": "success", "size": os.path.getsize(out_pdf)}


@app.post("/crop_page")
async def crop_page(req: CropRequest) -> dict[str, Any]:
    path = resolve_pdf_path(req.file_id)
    doc = fitz.open(path)
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
    doc = fitz.open(path)
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
        doc.save(str(out_pdf), garbage=4, deflate=True)
    finally:
        doc.close()
    return {"file_id": new_id, "status": "success", "size": os.path.getsize(out_pdf)}


@app.post("/watermark")
async def watermark_pdf(req: WatermarkRequest) -> dict[str, Any]:
    text = (req.text or "Draft").strip() or "Draft"
    path = resolve_pdf_path(req.file_id)
    doc = fitz.open(path)
    try:
        for i in range(len(doc)):
            _add_diagonal_watermark(doc[i], text, float(req.opacity))
        new_id, out_pdf = _new_work_file_id()
        doc.save(str(out_pdf), garbage=4, deflate=True)
    finally:
        doc.close()
    return {"file_id": new_id, "status": "success", "size": os.path.getsize(out_pdf)}


@app.get("/pdf_metadata/{file_id}")
async def get_pdf_metadata(file_id: str) -> dict[str, Any]:
    path = resolve_pdf_path(file_id)
    doc = fitz.open(path)
    try:
        meta = dict(doc.metadata)
    finally:
        doc.close()
    return {"file_id": file_id, "metadata": meta}


@app.post("/pdf_metadata")
async def update_pdf_metadata(req: MetadataUpdateRequest) -> dict[str, Any]:
    path = resolve_pdf_path(req.file_id)
    doc = fitz.open(path)
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
    doc = fitz.open(path)
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
    doc = fitz.open(path)
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
    doc = fitz.open(path)
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
    doc = fitz.open(path)
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
        page.insert_image(r, filename=str(stamp_path), keep_proportion=True)
        new_id, out_pdf = _new_work_file_id()
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
    out_path.write_bytes(signed)
    return {"file_id": new_id, "size": out_path.stat().st_size, "status": "success"}


# Mount static files last so /static/* is never shadowed by route registration quirks on Linux hosts.
_static_root = BASE_DIR / "static"
app.mount(
    "/static",
    StaticFiles(directory=str(_static_root), check_dir=True),
    name="static",
)
