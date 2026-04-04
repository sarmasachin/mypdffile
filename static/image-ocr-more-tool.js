(function () {
  function bridge() {
    return window.pdfEditorAppBridge || null;
  }

  const state = { fileId: null, pages: [], items: [], drag: null };

  function ensureView() {
    let view = document.getElementById("imageOcrMoreView");
    if (view) return view;
    view = document.createElement("div");
    view.id = "imageOcrMoreView";
    view.className = "view";
    view.innerHTML = `
      <div class="image-ocr-shell">
        <div class="image-ocr-head">
          <div>
            <div class="image-ocr-title">Image OCR Text Edit</div>
            <div class="image-ocr-sub">More menu se selected file ko OCR based edit mode me kholo. Text edit karo, boxes drag/resize karo, phir save.</div>
          </div>
          <div class="image-ocr-actions">
            <button type="button" class="image-ocr-btn image-ocr-btn-ghost" id="imageOcrMoreBackBtn">Back</button>
            <button type="button" class="image-ocr-btn image-ocr-btn-secondary" id="imageOcrMoreSaveBtn" disabled>Save PDF</button>
            <a id="imageOcrMoreDownloadBtn" class="image-ocr-btn image-ocr-btn-primary is-disabled" href="#" aria-disabled="true">Download</a>
          </div>
        </div>
        <div id="imageOcrMoreStatus" class="image-ocr-status is-hidden"></div>
        <div id="imageOcrMorePages" class="image-ocr-pages"></div>
      </div>
    `;
    document.body.appendChild(view);
    view.querySelector("#imageOcrMoreBackBtn")?.addEventListener("click", () => bridge()?.showHome?.());
    view.querySelector("#imageOcrMoreSaveBtn")?.addEventListener("click", saveEdits);
    bindGlobalPointerHandlers();
    return view;
  }

  function setStatus(message, isError) {
    const el = document.getElementById("imageOcrMoreStatus");
    if (!el) return;
    if (!message) {
      el.classList.add("is-hidden");
      el.textContent = "";
      return;
    }
    el.classList.remove("is-hidden");
    el.textContent = message;
    el.style.color = isError ? "#b42318" : "#0a4f8c";
  }

  function setDownload(url) {
    const el = document.getElementById("imageOcrMoreDownloadBtn");
    if (!el) return;
    if (!url) {
      el.href = "#";
      el.classList.add("is-disabled");
      el.setAttribute("aria-disabled", "true");
      return;
    }
    el.href = url;
    el.classList.remove("is-disabled");
    el.setAttribute("aria-disabled", "false");
  }

  function findItem(id) {
    return state.items.find((item) => item.id === id) || null;
  }

  function fitInput(input) {
    const minHeight = parseFloat(input.dataset.minHeight || "18");
    const lineHeight = parseFloat(input.dataset.lineHeight || "16");
    const lines = Math.max(1, String(input.value || "").split("\n").length);
    input.style.height = `${Math.max(minHeight, lines * lineHeight + 6)}px`;
  }

  function getScale(card) {
    const img = card.querySelector("img");
    const page = state.pages.find((x) => x.page === Number(card.dataset.pageNumber));
    if (!img || !img.clientWidth || !page) return null;
    return { scaleX: img.clientWidth / page.width, scaleY: img.clientHeight / page.height, page };
  }

  function placeItems(card) {
    const scale = getScale(card);
    if (!scale) return;
    const layer = card.querySelector(".image-ocr-layer");
    if (!layer) return;
    layer.style.width = `${scale.page.width * scale.scaleX}px`;
    layer.style.height = `${scale.page.height * scale.scaleY}px`;
    layer.querySelectorAll(".image-ocr-item").forEach((wrapper) => {
      const item = findItem(wrapper.dataset.id);
      if (!item) return;
      const [x0, y0, x1, y1] = item.bbox;
      const left = x0 * scale.scaleX;
      const top = y0 * scale.scaleY;
      const width = Math.max(72, (x1 - x0) * scale.scaleX);
      const height = Math.max(28, (y1 - y0) * scale.scaleY);
      wrapper.style.left = `${left}px`;
      wrapper.style.top = `${top}px`;
      wrapper.style.width = `${width}px`;
      wrapper.style.height = `${height}px`;
      const input = wrapper.querySelector(".image-ocr-input");
      if (input) {
        const fontPx = Math.max(11, item.size * scale.scaleY * 0.92);
        const linePx = Math.max(14, fontPx * 1.12);
        input.style.fontSize = `${fontPx}px`;
        input.style.lineHeight = `${linePx}px`;
        input.dataset.lineHeight = `${linePx}`;
        input.dataset.minHeight = `${Math.max(18, height - 12)}`;
        fitInput(input);
      }
    });
  }

  function createItem(item) {
    const wrapper = document.createElement("div");
    wrapper.className = "image-ocr-item";
    wrapper.dataset.id = item.id;
    wrapper.innerHTML = `
      <div class="image-ocr-drag" title="Drag box"></div>
      <textarea class="image-ocr-input" spellcheck="false" wrap="off" rows="1"></textarea>
      <div class="image-ocr-resize" title="Resize box"></div>
    `;
    const input = wrapper.querySelector(".image-ocr-input");
    input.value = item.text || "";
    input.addEventListener("focus", () => wrapper.classList.add("is-active"));
    input.addEventListener("blur", () => wrapper.classList.remove("is-active"));
    input.addEventListener("input", () => {
      const target = findItem(item.id);
      if (target) target.text = input.value;
      fitInput(input);
    });
    wrapper.querySelector(".image-ocr-drag")?.addEventListener("pointerdown", (e) => startDrag(e, item.id, "move"));
    wrapper.querySelector(".image-ocr-resize")?.addEventListener("pointerdown", (e) => startDrag(e, item.id, "resize"));
    return wrapper;
  }

  function createPageCard(page) {
    const items = state.items.filter((item) => item.page === page.page);
    const card = document.createElement("section");
    card.className = "image-ocr-card";
    card.dataset.pageNumber = String(page.page);
    card.innerHTML = `
      <div class="image-ocr-card-head">
        <div>
          <strong>Page ${page.page}</strong>
          <div class="image-ocr-meta">${items.length} OCR text block${items.length === 1 ? "" : "s"}</div>
        </div>
        <div class="image-ocr-meta">${Math.round(page.width)} × ${Math.round(page.height)} px</div>
      </div>
      <div class="image-ocr-stage">
        <img alt="Preview page ${page.page}">
        <div class="image-ocr-layer"></div>
      </div>
      <div class="image-ocr-help">Drag top strip to move. Blue handle se resize karo. Text seedha box ke andar edit karo.</div>
    `;
    const img = card.querySelector("img");
    const layer = card.querySelector(".image-ocr-layer");
    img.src = `/preview/${state.fileId}/${page.page}?source=input&v=${Date.now()}`;
    img.addEventListener("load", () => placeItems(card));
    items.forEach((item) => layer.appendChild(createItem(item)));
    return card;
  }

  function render() {
    const view = ensureView();
    const pagesWrap = view.querySelector("#imageOcrMorePages");
    pagesWrap.innerHTML = "";
    state.pages.forEach((page) => pagesWrap.appendChild(createPageCard(page)));
    view.querySelector("#imageOcrMoreSaveBtn").disabled = !state.fileId;
    setDownload(state.fileId ? `/download/${state.fileId}` : "");
    document.querySelectorAll(".view").forEach((el) => el.classList.remove("active"));
    view.classList.add("active");
  }

  function startDrag(event, itemId, mode) {
    event.preventDefault();
    event.stopPropagation();
    const card = event.currentTarget.closest(".image-ocr-card");
    const scale = card ? getScale(card) : null;
    const item = findItem(itemId);
    if (!card || !scale || !item) return;
    state.drag = {
      mode,
      itemId,
      startX: event.clientX,
      startY: event.clientY,
      startBBox: item.bbox.slice(),
      card,
      scaleX: scale.scaleX,
      scaleY: scale.scaleY,
    };
  }

  function bindGlobalPointerHandlers() {
    if (window.__imageOcrMoreToolBound) return;
    window.__imageOcrMoreToolBound = true;
    window.addEventListener("pointermove", (e) => {
      const drag = state.drag;
      if (!drag) return;
      const item = findItem(drag.itemId);
      if (!item) return;
      const dx = (e.clientX - drag.startX) / drag.scaleX;
      const dy = (e.clientY - drag.startY) / drag.scaleY;
      const [x0, y0, x1, y1] = drag.startBBox;
      if (drag.mode === "move") {
        item.bbox = [x0 + dx, y0 + dy, x1 + dx, y1 + dy];
      } else {
        item.bbox = [x0, y0, Math.max(x0 + 36, x1 + dx), Math.max(y0 + 18, y1 + dy)];
      }
      placeItems(drag.card);
    });
    window.addEventListener("pointerup", () => {
      state.drag = null;
    });
    window.addEventListener("pointercancel", () => {
      state.drag = null;
    });
    window.addEventListener("resize", () => {
      document.querySelectorAll("#imageOcrMoreView .image-ocr-card").forEach((card) => placeItems(card));
    });
  }

  async function openTool() {
    const api = bridge();
    const selected = api?.getSelectedMoreFile?.();
    if (!selected || !selected.fileObj || !selected.fileObj.id) {
      api?.showAlert?.("Demo File", "Please select a real uploaded file first.", false);
      return;
    }
    api?.hideMoreOptions?.();
    api?.setBusy?.("Detecting OCR text...");
    try {
      const res = await fetch(`/image_ocr_tool/analyze/${selected.fileObj.id}`);
      const json = await res.json();
      if (!res.ok) throw new Error(json.detail || "Could not analyze file.");
      state.fileId = selected.fileObj.id;
      state.pages = json.pages || [];
      state.items = json.items || [];
      render();
      setStatus(`Ready. ${state.items.length} editable OCR block${state.items.length === 1 ? "" : "s"} detected.`);
    } catch (err) {
      api?.showAlert?.("OCR Failed", err.message || String(err), false);
    } finally {
      api?.clearBusy?.();
    }
  }

  async function saveEdits() {
    const api = bridge();
    if (!state.fileId) return;
    api?.setBusy?.("Saving OCR edits...");
    try {
      const res = await fetch("/image_ocr_tool/edit", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ file_id: state.fileId, edits: state.items }),
      });
      const json = await res.json();
      if (!res.ok) throw new Error(json.detail || "Save failed.");
      setDownload(json.download_url || `/download/${state.fileId}`);
      setStatus("OCR edits saved. Download or continue refining boxes.");
      api?.setFileThumb?.(state.fileId, `/preview_edited/${state.fileId}/1?v=${Date.now()}`);
      api?.refreshFiles?.();
      document.querySelectorAll("#imageOcrMoreView .image-ocr-card img").forEach((img) => {
        const page = img.closest(".image-ocr-card")?.dataset.pageNumber;
        if (page) img.src = `/preview_edited/${state.fileId}/${page}?v=${Date.now()}`;
      });
    } catch (err) {
      api?.showAlert?.("Save Failed", err.message || String(err), false);
    } finally {
      api?.clearBusy?.();
    }
  }

  document.addEventListener("DOMContentLoaded", () => {
    ensureView();
    document.getElementById("imageOcrTextEditBtn")?.addEventListener("click", () => {
      void openTool();
    });
  });
})();
