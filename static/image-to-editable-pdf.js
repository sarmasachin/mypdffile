(function () {
  const imageInput = document.getElementById("imageInput");
  const chooseImagesBtn = document.getElementById("chooseImagesBtn");
  const savePdfBtn = document.getElementById("savePdfBtn");
  const downloadPdfBtn = document.getElementById("downloadPdfBtn");
  const emptyState = document.getElementById("emptyState");
  const editorCanvas = document.getElementById("editorCanvas");
  const statusBar = document.getElementById("statusBar");

  const state = {
    fileId: null,
    pages: [],
    items: [],
  };

  function setStatus(message, isError) {
    statusBar.textContent = message || "";
    statusBar.classList.toggle("hidden", !message);
    statusBar.style.color = isError ? "#b42318" : "#0a4f8c";
  }

  function setDownloadState(enabled) {
    downloadPdfBtn.classList.toggle("disabled", !enabled);
    downloadPdfBtn.setAttribute("aria-disabled", enabled ? "false" : "true");
  }

  function pageItems(pageNumber) {
    return state.items.filter((item) => item.page === pageNumber);
  }

  function fitTextAreaHeight(el) {
    const minHeight = parseFloat(el.dataset.minHeight || "16");
    const manualLines = Math.max(1, String(el.value || "").split("\n").length);
    const lineHeight = parseFloat(el.dataset.lineHeight || "16");
    const neededHeight = Math.max(minHeight, manualLines * lineHeight + 6);
    el.style.height = `${neededHeight}px`;
  }

  function placeBoxes(card) {
    const pageNumber = Number(card.dataset.pageNumber);
    const img = card.querySelector("img");
    const layer = card.querySelector(".ocr-layer");
    const naturalPage = state.pages.find((p) => p.page === pageNumber);
    if (!img || !layer || !naturalPage || !img.clientWidth) return;

    layer.style.width = `${img.clientWidth}px`;
    layer.style.height = `${img.clientHeight}px`;

    const scaleX = img.clientWidth / naturalPage.width;
    const scaleY = img.clientHeight / naturalPage.height;

    Array.from(layer.querySelectorAll(".ocr-box")).forEach((box) => {
      const item = pageItems(pageNumber).find((entry) => entry.id === box.dataset.id);
      if (!item) return;
      const [x0, y0, x1, y1] = item.bbox;
      const left = x0 * scaleX;
      const top = y0 * scaleY;
      const width = Math.max(72, (x1 - x0) * scaleX);
      const height = Math.max(22, (y1 - y0) * scaleY);
      box.style.left = `${left}px`;
      box.style.top = `${top}px`;
      box.style.width = `${width}px`;
      box.style.height = `${height}px`;
      const fontPx = Math.max(11, item.size * scaleY * 0.9);
      const linePx = Math.max(14, fontPx * 1.12);
      box.style.fontSize = `${fontPx}px`;
      box.style.lineHeight = `${linePx}px`;
      box.dataset.minHeight = `${height}`;
      box.dataset.lineHeight = `${linePx}`;
      fitTextAreaHeight(box);
    });
  }

  function createPageCard(page) {
    const items = pageItems(page.page);
    const card = document.createElement("section");
    card.className = "page-card";
    card.dataset.pageNumber = String(page.page);
    card.innerHTML = `
      <div class="page-header">
        <div>
          <strong>Page ${page.page}</strong>
          <div class="page-meta">${items.length} OCR text block${items.length === 1 ? "" : "s"}</div>
        </div>
        <div class="page-meta">${Math.round(page.width)} × ${Math.round(page.height)} px</div>
      </div>
      <div class="page-stage">
        <img alt="Preview for page ${page.page}">
        <div class="ocr-layer"></div>
      </div>
      <div class="ocr-note">Edit the boxes directly on the page, then save the PDF.</div>
    `;

    const img = card.querySelector("img");
    const layer = card.querySelector(".ocr-layer");
    img.src = `/image_to_pdf_ocr/preview/${state.fileId}/${page.page}?v=${Date.now()}`;
    img.addEventListener("load", function () {
      placeBoxes(card);
    });

    items.forEach((item) => {
      const textarea = document.createElement("textarea");
      textarea.className = "ocr-box";
      textarea.value = item.text || "";
      textarea.dataset.id = item.id;
      textarea.setAttribute("spellcheck", "false");
      textarea.setAttribute("wrap", "off");
      textarea.setAttribute("rows", "1");
      textarea.addEventListener("input", function () {
        const target = state.items.find((entry) => entry.id === item.id);
        if (target) {
          target.text = textarea.value;
        }
        fitTextAreaHeight(textarea);
      });
      layer.appendChild(textarea);
    });

    return card;
  }

  function renderEditor() {
    emptyState.classList.add("hidden");
    editorCanvas.classList.remove("hidden");
    editorCanvas.innerHTML = "";
    state.pages.forEach((page) => {
      editorCanvas.appendChild(createPageCard(page));
    });
    savePdfBtn.disabled = !state.fileId;
    setDownloadState(!!state.fileId);
    if (state.fileId) {
      downloadPdfBtn.href = `/image_to_pdf_ocr/download/${state.fileId}`;
    }
  }

  async function uploadImages(files) {
    if (!files || !files.length) return;
    const form = new FormData();
    Array.from(files).forEach((file) => form.append("files", file));

    setStatus("Uploading images and detecting text...");
    chooseImagesBtn.disabled = true;
    savePdfBtn.disabled = true;
    setDownloadState(false);

    try {
      const res = await fetch("/image_to_pdf_ocr/upload", { method: "POST", body: form });
      const json = await res.json();
      if (!res.ok) {
        throw new Error(json.detail || "Upload failed");
      }
      state.fileId = json.file_id;
      state.pages = json.pages || [];
      state.items = json.items || [];
      renderEditor();
      setStatus(`Ready. ${state.items.length} editable OCR block${state.items.length === 1 ? "" : "s"} detected.`);
    } catch (err) {
      setStatus(err.message || String(err), true);
    } finally {
      chooseImagesBtn.disabled = false;
    }
  }

  async function savePdf() {
    if (!state.fileId) return;
    savePdfBtn.disabled = true;
    setStatus("Saving edited PDF...");
    try {
      const edits = state.items.map((item) => ({
        id: item.id,
        page: item.page,
        text: item.text || "",
        bbox: item.bbox,
        font: item.font || "helv",
        size: item.size || 11,
        color: item.color || "#111111",
        align: item.align || "left",
      }));

      const res = await fetch("/image_to_pdf_ocr/edit", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ file_id: state.fileId, edits }),
      });
      const json = await res.json();
      if (!res.ok) {
        throw new Error(json.detail || "Save failed");
      }
      setStatus("PDF saved successfully. Download the updated file when ready.");
      downloadPdfBtn.href = json.download_url;
      setDownloadState(true);
      document.querySelectorAll(".page-card").forEach((card) => {
        const img = card.querySelector("img");
        if (!img) return;
        const pageNumber = card.dataset.pageNumber;
        img.src = `/image_to_pdf_ocr/preview/${state.fileId}/${pageNumber}?v=${Date.now()}`;
      });
    } catch (err) {
      setStatus(err.message || String(err), true);
    } finally {
      savePdfBtn.disabled = false;
    }
  }

  function openPicker() {
    imageInput.click();
  }

  chooseImagesBtn?.addEventListener("click", openPicker);
  savePdfBtn?.addEventListener("click", savePdf);
  imageInput?.addEventListener("change", function () {
    uploadImages(imageInput.files);
    imageInput.value = "";
  });

  document.querySelectorAll('[data-action="pick-images"]').forEach((btn) => {
    btn.addEventListener("click", openPicker);
  });

  window.addEventListener("resize", function () {
    document.querySelectorAll(".page-card").forEach((card) => placeBoxes(card));
  });
})();
