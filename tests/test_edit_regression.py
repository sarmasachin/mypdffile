"""
Regression checks for POST /edit (redact + insert). Run from project root:

  .venv\\Scripts\\python.exe tests\\test_edit_regression.py

These mirror the user flow: upload -> analyze -> edit -> read edited.pdf text.
"""
from __future__ import annotations

import io
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

import fitz  # noqa: E402
from fastapi.testclient import TestClient  # noqa: E402

from app import WORK_DIR, app  # noqa: E402


def _pdf_two_lines() -> bytes:
    doc = fitz.open()
    page = doc.new_page(width=400, height=400)
    page.insert_text((50, 80), "AAA First line", fontsize=14)
    page.insert_text((50, 120), "BBB Second line", fontsize=14)
    buf = io.BytesIO()
    doc.save(buf)
    doc.close()
    return buf.getvalue()


def _pdf_tight_lines() -> bytes:
    doc = fitz.open()
    page = doc.new_page(width=400, height=400)
    page.insert_text((50, 72), "LineOne", fontsize=12)
    page.insert_text((50, 100), "LineTwo", fontsize=12)
    buf = io.BytesIO()
    doc.save(buf)
    doc.close()
    return buf.getvalue()


def _assert(cond: bool, msg: str) -> None:
    if not cond:
        raise AssertionError(msg)


def main() -> None:
    client = TestClient(app)

    # --- Case 1: replace first line with longer string (was failing: insert_textbox rc < 0) ---
    fid = client.post(
        "/upload", files={"file": ("t.pdf", _pdf_two_lines(), "application/pdf")}
    ).json()["file_id"]
    items = client.get(f"/analyze/{fid}").json()["items"]
    line1 = items[0]
    r = client.post(
        "/edit",
        json={
            "file_id": fid,
            "edits": [
                {
                    "id": line1["id"],
                    "text": "REPLACED_LINE_1",
                    "page": line1["page"],
                    "bbox": line1["bbox"],
                    "original_bbox": line1["bbox"],
                    "font": line1.get("font", "helv"),
                    "size": line1.get("size", 11.0),
                    "color": "#000000",
                    "align": "left",
                    "is_underline": False,
                    "is_strike": False,
                }
            ],
        },
    )
    _assert(r.status_code == 200, r.text)
    doc = fitz.open(WORK_DIR / fid / "edited.pdf")
    t = doc[0].get_text("text")
    doc.close()
    _assert("REPLACED_LINE_1" in t, f"replacement missing: {t!r}")
    _assert("BBB Second line" in t, f"unedited line lost: {t!r}")

    # --- Case 2: multiline replacement in first box ---
    fid2 = client.post(
        "/upload", files={"file": ("t2.pdf", _pdf_two_lines(), "application/pdf")}
    ).json()["file_id"]
    items2 = client.get(f"/analyze/{fid2}").json()["items"]
    a = items2[0]
    multi = "Row A\nRow B\nRow C"
    r2 = client.post(
        "/edit",
        json={
            "file_id": fid2,
            "edits": [
                {
                    "id": a["id"],
                    "text": multi,
                    "page": a["page"],
                    "bbox": a["bbox"],
                    "original_bbox": a["bbox"],
                    "font": a.get("font", "helv"),
                    "size": a.get("size", 11.0),
                    "color": "#000000",
                    "align": "left",
                    "is_underline": False,
                    "is_strike": False,
                }
            ],
        },
    )
    _assert(r2.status_code == 200, r2.text)
    doc2 = fitz.open(WORK_DIR / fid2 / "edited.pdf")
    t2 = doc2[0].get_text("text")
    doc2.close()
    t2n = t2.replace("\xa0", " ")
    for part in ("Row A", "Row B", "Row C"):
        _assert(part in t2n, f"expected {part!r} in {t2!r}")

    # --- Case 3: tight spacing; edit line 1; neighbor should remain ---
    fid3 = client.post(
        "/upload", files={"file": ("t3.pdf", _pdf_tight_lines(), "application/pdf")}
    ).json()["file_id"]
    items3 = client.get(f"/analyze/{fid3}").json()["items"]
    lo = items3[0]
    r3 = client.post(
        "/edit",
        json={
            "file_id": fid3,
            "edits": [
                {
                    "id": lo["id"],
                    "text": "ONE_OK",
                    "page": lo["page"],
                    "bbox": lo["bbox"],
                    "original_bbox": lo["bbox"],
                    "font": lo.get("font", "helv"),
                    "size": lo.get("size", 11.0),
                    "color": "#000000",
                    "align": "left",
                    "is_underline": False,
                    "is_strike": False,
                }
            ],
        },
    )
    _assert(r3.status_code == 200, r3.text)
    doc3 = fitz.open(WORK_DIR / fid3 / "edited.pdf")
    t3 = doc3[0].get_text("text")
    doc3.close()
    _assert("ONE_OK" in t3, f"short replacement missing: {t3!r}")
    _assert("LineTwo" in t3, f"neighbor missing: {t3!r}")

    print("edit regression: all checks passed")


if __name__ == "__main__":
    main()
