const homeView = document.getElementById("homeView");
const editorView = document.getElementById("editorView");
const pdfFile = document.getElementById("pdfFile");
const loadingOverlay = document.getElementById("loadingOverlay");
const uploadStatus = document.getElementById("uploadStatus");
const pageEditor = document.getElementById("pageEditor");
const pdfBlockFloatMenu = document.getElementById("pdfBlockFloatMenu");
/** @type {HTMLElement | null} */
let pdfFloatMenuTargetBox = null;
const saveBtn = document.getElementById("saveBtn");
const closeEditorBtn = document.getElementById("closeEditorBtn");
const recentFilesList = document.getElementById("recentFilesList");

let currentFileId = null;
let originalItems = [];
let pagesMeta = [];
/** @type {HTMLTextAreaElement[]} */
let textAreas = [];
let selectedMoreFile = null; // Track which file's "More Options" is active
let selectedCombineFiles = []; // Temporary list for combine workflow
/** @type {{ id?: string, title: string, [k: string]: any } | null} */
let shareContextFile = null;
/** True while the formatting toolbar is being used — blur handlers must not steal focus. */
let isFormatting = false;

/**
 * PDF text overlay (blue box): typography tuned only here so save/bbox math stays consistent.
 * 1.1 was too tight — ascenders/descenders looked clipped inside the wrapper.
 */
const PDF_OVERLAY_LINE_HEIGHT = 1.25;
const PDF_OVERLAY_MIN_BOX_HT_EM = 1.34;
/** Horizontal slack beyond PDF bbox (pt → px); keeps long lines slightly less cramped. */
const PDF_OVERLAY_WIDTH_EXTRA_PX = 12;
/** Extra px after canvas measureText so real glyphs are not clipped vs canvas metrics. */
const PDF_OVERLAY_WIDTH_MEASURE_PAD_PX = 8;

/** Min pixel width for one logical line (canvas + padding); used to grow the box to the right while typing. */
function pdfOverlayUnwrappedContentWidthPx(area) {
  try {
    const st = getComputedStyle(area);
    const ctx = document.createElement("canvas").getContext("2d");
    if (!ctx) return 0;
    ctx.font = st.font || "16px sans-serif";
    const padL = parseFloat(st.paddingLeft) || 0;
    const padR = parseFloat(st.paddingRight) || 0;
    let max = 0;
    const lines = area.value.split("\n");
    for (let i = 0; i < lines.length; i++) {
      const lineW = ctx.measureText(lines[i] || " ").width;
      if (lineW > max) max = lineW;
    }
    return Math.ceil(max + padL + padR + PDF_OVERLAY_WIDTH_MEASURE_PAD_PX);
  } catch (_) {
    return 0;
  }
}

function sanitizePdfFilename(name) {
  if (!name || typeof name !== "string") return "document.pdf";
  let s = name.trim().replace(/[/\\?%*:|"<>]/g, "_").replace(/\s+/g, " ");
  if (!s.toLowerCase().endsWith(".pdf")) s = s ? `${s}.pdf` : "document.pdf";
  return s.length > 200 ? s.slice(0, 200) : s;
}

function getDownloadFilenameForFileId(fileId) {
  if (!fileId) return "document.pdf";
  const f = mockFiles.find((x) => x.id === fileId);
  return sanitizePdfFilename(f?.title || "edited_document.pdf");
}

function getShareDownloadUrl() {
  if (!shareContextFile || !shareContextFile.id) return "";
  const name = getDownloadFilenameForFileId(shareContextFile.id);
  const u = new URL(`${window.location.origin}/download/${shareContextFile.id}`);
  u.searchParams.set("filename", name);
  return u.toString();
}

function getShareBodyText() {
  const url = getShareDownloadUrl();
  const title = (shareContextFile && shareContextFile.title) || "PDF";
  return url ? `Download: ${title}\n${url}` : "";
}

function requireShareFileOrAlert() {
  if (shareContextFile && shareContextFile.id) return true;
  showCustomAlert("No file", "Upload a real PDF first, then open Share on that file.", false);
  return false;
}

function updateSharingLinkDisplay() {
  const el = document.getElementById("sharingLinkText");
  if (el) el.textContent = getShareDownloadUrl() || "—";
}

function updateShareSheetUI() {
  const ok = !!(shareContextFile && shareContextFile.id);
  const row = document.querySelector("#shareSheetOverlay .share-apps-row");
  if (row) {
    row.style.opacity = ok ? "1" : "0.45";
    row.style.pointerEvents = ok ? "auto" : "none";
  }
  [document.getElementById("shareSendCopyBtn"), document.getElementById("shareSendCompressedBtn")].forEach((btn) => {
    if (!btn) return;
    btn.disabled = !ok;
    btn.style.opacity = ok ? "1" : "0.55";
    btn.style.cursor = ok ? "pointer" : "not-allowed";
  });
  updateSharingLinkDisplay();
}

/** Mobile browsers often block programmatic <a download> clicks (especially after await). Full navigation to GET /download/… is reliable. */
function isMobileDownloadEnvironment() {
  const coarse =
    typeof window.matchMedia === "function" && window.matchMedia("(pointer: coarse)").matches;
  const ua = /Android|webOS|iPhone|iPad|iPod|BlackBerry|IEMobile|Opera Mini/i.test(navigator.userAgent);
  return coarse || ua;
}

function buildDownloadUrl(fileId, filename) {
  const u = new URL(`${window.location.origin}/download/${fileId}`);
  u.searchParams.set("filename", filename);
  return u.toString();
}

/**
 * @param {string} fileId
 * @param {string} filename suggested filename (query + download attr)
 * @returns {boolean} true if navigation/new-tab path (no in-page success toast needed)
 */
function downloadPdfByFileId(fileId, filename) {
  const url = buildDownloadUrl(fileId, filename);
  if (isMobileDownloadEnvironment()) {
    const w = window.open(url, "_blank", "noopener,noreferrer");
    if (!w || typeof w.closed === "undefined") {
      window.location.assign(url);
    }
    return true;
  }
  const link = document.createElement("a");
  link.href = url;
  link.download = filename;
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
  return false;
}

function triggerSharePdfDownload() {
  if (!requireShareFileOrAlert()) return;
  const name = getDownloadFilenameForFileId(shareContextFile.id);
  downloadPdfByFileId(shareContextFile.id, name);
}

function wireShareSheetActions() {
  const copyUrlToClipboard = async () => {
    if (!requireShareFileOrAlert()) return;
    const url = getShareDownloadUrl();
    try {
      await navigator.clipboard.writeText(url);
      showCustomAlert("Copied", "Link copied to clipboard.", true);
    } catch {
      window.prompt("Copy this link:", url);
    }
  };

  document.getElementById("shareRowCopyLink")?.addEventListener("click", copyUrlToClipboard);
  document.getElementById("shareRowCopyLink")?.addEventListener("keydown", (e) => {
    if (e.key === "Enter" || e.key === " ") {
      e.preventDefault();
      copyUrlToClipboard();
    }
  });

  document.getElementById("shareRowWhatsApp")?.addEventListener("click", () => {
    if (!requireShareFileOrAlert()) return;
    window.open(`https://wa.me/?text=${encodeURIComponent(getShareBodyText())}`, "_blank", "noopener,noreferrer");
  });

  document.getElementById("shareRowGmail")?.addEventListener("click", () => {
    if (!requireShareFileOrAlert()) return;
    const subj = encodeURIComponent(shareContextFile.title || "PDF");
    const body = encodeURIComponent(getShareBodyText());
    window.location.href = `mailto:?subject=${subj}&body=${body}`;
  });

  document.getElementById("shareRowMessages")?.addEventListener("click", () => {
    if (!requireShareFileOrAlert()) return;
    const body = encodeURIComponent(getShareBodyText());
    window.location.href = `sms:?&body=${body}`;
  });

  document.getElementById("shareRowShareVia")?.addEventListener("click", () => {
    if (!requireShareFileOrAlert()) return;
    updateSharingLinkDisplay();
    shareSheetOverlay?.classList.add("hidden");
    document.getElementById("sharingLinkSheetOverlay")?.classList.remove("hidden");
  });

  document.getElementById("shareSendCopyBtn")?.addEventListener("click", () => {
    if (!requireShareFileOrAlert()) return;
    triggerSharePdfDownload();
  });

  document.getElementById("shareSendCompressedBtn")?.addEventListener("click", async () => {
    if (!requireShareFileOrAlert()) return;
    // Open a tab synchronously (still part of the tap). After await(fetch) the browser may block
    // window.open/assign — this preserves a valid target for the download URL.
    let downloadTab = null;
    if (isMobileDownloadEnvironment()) {
      try {
        downloadTab = window.open("about:blank", "_blank", "noopener,noreferrer");
      } catch {
        /* ignore */
      }
    }
    const fid = shareContextFile.id;
    const idx = mockFiles.findIndex((f) => f.id === fid);
    shareSheetOverlay?.classList.add("hidden");
    loadingOverlay.classList.remove("hidden");
    uploadStatus.textContent = "Compressing PDF...";
    const currentCount = idx > -1 ? mockFiles[idx].compressionCount || 0 : 0;
    const qualities = [70, 45, 25, 10];
    const targetQuality = qualities[currentCount] || 5;
    try {
      const res = await fetch("/compress", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ file_id: fid, quality: targetQuality }),
      });
      if (!res.ok) {
        let detail = "Compression failed.";
        try {
          const j = await res.json();
          if (j.detail) detail = typeof j.detail === "string" ? j.detail : JSON.stringify(j.detail);
        } catch {
          /* ignore */
        }
        throw new Error(detail);
      }
      const json = await res.json();
      if (idx > -1) {
        mockFiles[idx].id = json.file_id;
        mockFiles[idx].size = json.size;
        mockFiles[idx].compressionCount = (mockFiles[idx].compressionCount || 0) + 1;
        shareContextFile = mockFiles[idx];
      }
      renderMockFiles();
      persistRecentFiles();
      const name = getDownloadFilenameForFileId(json.file_id);
      const dl = buildDownloadUrl(json.file_id, name);
      if (downloadTab && !downloadTab.closed) {
        downloadTab.location.href = dl;
      } else {
        const navigatedAway = downloadPdfByFileId(json.file_id, name);
        if (!navigatedAway) {
          showCustomAlert("Success", "Compressed PDF downloaded.", true);
        }
      }
    } catch (err) {
      if (downloadTab && !downloadTab.closed) {
        try {
          downloadTab.close();
        } catch {
          /* ignore */
        }
      }
      showCustomAlert("Failed", err.message || String(err), false);
    } finally {
      loadingOverlay.classList.add("hidden");
    }
  });

  document.getElementById("sharingLinkCopyBtn")?.addEventListener("click", copyUrlToClipboard);

  document.getElementById("quickShareRowBtn")?.addEventListener("click", async () => {
    if (!requireShareFileOrAlert()) return;
    const url = getShareDownloadUrl();
    const title = shareContextFile.title || "PDF";
    if (navigator.share) {
      try {
        await navigator.share({ title, text: getShareBodyText(), url });
        return;
      } catch (e) {
        if (e && e.name === "AbortError") return;
      }
    }
    await copyUrlToClipboard();
  });

  document.getElementById("secondSheetWhatsAppBtn")?.addEventListener("click", () => {
    if (!requireShareFileOrAlert()) return;
    window.open(`https://wa.me/?text=${encodeURIComponent(getShareBodyText())}`, "_blank", "noopener,noreferrer");
  });

  document.getElementById("secondSheetMessagesBtn")?.addEventListener("click", () => {
    if (!requireShareFileOrAlert()) return;
    window.location.href = `sms:?&body=${encodeURIComponent(getShareBodyText())}`;
  });

  document.getElementById("secondSheetChromeBtn")?.addEventListener("click", () => {
    if (!requireShareFileOrAlert()) return;
    window.open(getShareDownloadUrl(), "_blank", "noopener,noreferrer");
  });

  document.getElementById("secondSheetGmailBtn")?.addEventListener("click", () => {
    if (!requireShareFileOrAlert()) return;
    const subj = encodeURIComponent(shareContextFile.title || "PDF");
    const body = encodeURIComponent(getShareBodyText());
    window.location.href = `mailto:?subject=${subj}&body=${body}`;
  });

  document.querySelectorAll("#sharingLinkSheetOverlay .contacts-row .share-app-item").forEach((el) => {
    el.setAttribute("role", "button");
    el.setAttribute("tabindex", "0");
    el.addEventListener("click", () => {
      if (!requireShareFileOrAlert()) return;
      window.open(`https://wa.me/?text=${encodeURIComponent(getShareBodyText())}`, "_blank", "noopener,noreferrer");
    });
  });
}

const RECENT_FILES_KEY = "pdfEditorRecentFiles_v1";

function loadRecentFilesFromStorage() {
  try {
    const raw = localStorage.getItem(RECENT_FILES_KEY);
    if (!raw) return [];
    const arr = JSON.parse(raw);
    if (!Array.isArray(arr)) return [];
    const seen = new Set();
    const out = [];
    for (const x of arr) {
      if (!x || !x.id || seen.has(x.id)) continue;
      seen.add(x.id);
      out.push(x);
    }
    return out;
  } catch {
    return [];
  }
}

function persistRecentFiles() {
  try {
    const payload = mockFiles
      .filter((f) => f.id)
      .map((f) => ({
        id: f.id,
        title: f.title,
        size: f.size,
        date: f.date,
        thumb: f.thumb,
        compressionCount: f.compressionCount,
      }));
    localStorage.setItem(RECENT_FILES_KEY, JSON.stringify(payload));
  } catch {
    /* quota / private mode */
  }
}

async function executeDeleteFile(file) {
  moreOptionsSheetOverlay?.classList.add("hidden");
  loadingOverlay.classList.remove("hidden");
  uploadStatus.textContent = "Deleting file...";
  try {
    const res = await fetch(`${window.location.origin}/file/${encodeURIComponent(file.id)}`, {
      method: "DELETE",
    });
    if (!res.ok && res.status !== 404) {
      let msg = "Delete failed.";
      try {
        const j = await res.json();
        if (j.detail) msg = typeof j.detail === "string" ? j.detail : JSON.stringify(j.detail);
      } catch {
        /* ignore */
      }
      throw new Error(msg);
    }
  } catch (e) {
    showCustomAlert("Failed", e.message || String(e), false);
    loadingOverlay.classList.add("hidden");
    return;
  }
  loadingOverlay.classList.add("hidden");

  if (currentFileId === file.id) {
    currentFileId = null;
    pageEditor.innerHTML = "";
    textAreas = [];
    pagesMeta = [];
    originalItems = [];
    editorHistory = [];
    editorHistoryIndex = -1;
    editorView.classList.remove("active");
    homeView.classList.add("active");
  }

  const ix = mockFiles.findIndex((f) => f.id === file.id);
  if (ix > -1) mockFiles.splice(ix, 1);
  if (shareContextFile && shareContextFile.id === file.id) shareContextFile = null;
  if (selectedMoreFile && selectedMoreFile.fileObj && selectedMoreFile.fileObj.id === file.id) {
    selectedMoreFile = null;
  }
  persistRecentFiles();
  renderMockFiles();
  showCustomAlert("Deleted", "File removed.", true);
}

function confirmAndDeleteFile(file) {
  if (!file || !file.id) {
    showCustomAlert("Demo file", "Upload a real PDF first, then you can delete it from the list.", false);
    return;
  }
  showCustomConfirm(
    "Delete file?",
    `Remove "${file.title}" from this list and delete its copy on the server? This cannot be undone.`,
    () => {
      void executeDeleteFile(file);
    }
  );
}

let mockFiles = [...loadRecentFilesFromStorage()];

const MAX_EDITOR_HISTORY = 50;
let editorHistory = [];
let editorHistoryIndex = -1;
let isApplyingHistory = false;
let historyInputTimer = null;

function normalizeEditorText(s) {
  return String(s || "").replace(/\r\n/g, "\n");
}

function editorItemsSnapshot() {
  return textAreas.map((a) => ({
    id: a.dataset.id,
    text: a.value,
    page: Number(a.dataset.page),
    bbox: JSON.parse(a.dataset.bbox || "[0,0,0,0]"),
    original_bbox: JSON.parse(a.dataset.originalBbox || a.dataset.bbox || "[0,0,0,0]"),
    font: a.dataset.font || "helv",
    size: parseFloat(a.dataset.size || 11),
    color: a.dataset.color || "#000000",
    align: a.dataset.align || "left",
    originalText: a.dataset.originalText ?? "",
    wasFormatted: a.dataset.wasFormatted,
    style: {
      fontWeight: a.style.fontWeight || "",
      fontStyle: a.style.fontStyle || "",
      textDecoration: a.style.textDecoration || "",
      color: a.style.color || "",
      textAlign: a.style.textAlign || "",
      fontSize: a.style.fontSize || "",
      lineHeight: a.style.lineHeight || "",
    },
  }));
}

function snapshotEditorState() {
  return {
    pagesMeta: JSON.parse(JSON.stringify(pagesMeta)),
    items: editorItemsSnapshot(),
  };
}

function pushEditorHistory() {
  if (isApplyingHistory) return;
  if (!editorView.classList.contains("active")) return;
  if (!pagesMeta.length) return;
  const snap = snapshotEditorState();
  if (
    editorHistoryIndex >= 0 &&
    editorHistory[editorHistoryIndex] &&
    JSON.stringify(editorHistory[editorHistoryIndex]) === JSON.stringify(snap)
  ) {
    return;
  }
  editorHistory = editorHistory.slice(0, editorHistoryIndex + 1);
  editorHistory.push(snap);
  if (editorHistory.length > MAX_EDITOR_HISTORY) {
    editorHistory.shift();
  }
  editorHistoryIndex = editorHistory.length - 1;
  updateUndoRedoUI();
}

function updateUndoRedoUI() {
  const u = document.getElementById("undoBtn");
  const r = document.getElementById("redoBtn");
  if (u) {
    u.disabled = editorHistoryIndex <= 0;
    u.style.opacity = u.disabled ? "0.45" : "1";
  }
  if (r) {
    r.disabled = editorHistoryIndex >= editorHistory.length - 1;
    r.style.opacity = r.disabled ? "0.45" : "1";
  }
}

function initEditorHistory() {
  editorHistory = [];
  editorHistoryIndex = -1;
  pushEditorHistory();
}

async function applyEditorSnapshot(snap) {
  isApplyingHistory = true;
  try {
    pagesMeta = snap.pagesMeta;
    const itemsForRender = snap.items.map((it) => {
      const { style: _st, wasFormatted: _wf, ...rest } = it;
      return rest;
    });
    originalItems = itemsForRender.map((it) => ({ ...it }));
    await renderEditor(pagesMeta, itemsForRender);
    applyStylesFromSnapshot(snap.items);
    lastActiveArea = textAreas[0] || null;
  } finally {
    isApplyingHistory = false;
    updateUndoRedoUI();
  }
}

function applyStylesFromSnapshot(items) {
  const byId = new Map(items.map((it) => [it.id, it]));
  for (const a of textAreas) {
    const it = byId.get(a.dataset.id);
    if (!it) continue;
    const st = it.style;
    if (st) {
      if (st.fontWeight) a.style.fontWeight = st.fontWeight;
      if (st.fontStyle) a.style.fontStyle = st.fontStyle;
      if (st.textDecoration) a.style.textDecoration = st.textDecoration;
      if (st.color) {
        a.style.color = st.color;
        a.dataset.color = it.color || st.color;
      }
      if (st.textAlign) {
        a.style.textAlign = st.textAlign;
        a.dataset.align = it.align || st.textAlign;
      }
      if (st.fontSize) a.style.fontSize = st.fontSize;
      if (st.lineHeight) a.style.lineHeight = st.lineHeight;
    }
    if (it.wasFormatted) a.dataset.wasFormatted = it.wasFormatted;
    else delete a.dataset.wasFormatted;
    const box = a.closest(".pdf-box-wrapper");
    if (box) {
      if (normalizeEditorText(a.value) !== normalizeEditorText(a.dataset.originalText)) box.classList.add("edited");
      else box.classList.remove("edited");
    }
  }
}

function undoEditor() {
  if (editorHistoryIndex <= 0) return;
  editorHistoryIndex--;
  void applyEditorSnapshot(editorHistory[editorHistoryIndex]).catch((err) => {
    console.error(err);
    showCustomAlert("Undo failed", err.message || String(err), false);
  });
}

function redoEditor() {
  if (editorHistoryIndex >= editorHistory.length - 1) return;
  editorHistoryIndex++;
  void applyEditorSnapshot(editorHistory[editorHistoryIndex]).catch((err) => {
    console.error(err);
    showCustomAlert("Redo failed", err.message || String(err), false);
  });
}

pageEditor.addEventListener("input", (e) => {
  if (!e.target.classList.contains("pdf-text-input")) return;
  clearTimeout(historyInputTimer);
  historyInputTimer = setTimeout(() => pushEditorHistory(), 450);
});

let editorToolbarHistoryTimer = null;
document.querySelector(".editor-toolbar")?.addEventListener("click", () => {
  clearTimeout(editorToolbarHistoryTimer);
  editorToolbarHistoryTimer = setTimeout(() => pushEditorHistory(), 80);
});

function hidePdfBlockFloatMenu() {
  pdfFloatMenuTargetBox = null;
  if (pdfBlockFloatMenu) {
    pdfBlockFloatMenu.classList.add("hidden");
    pdfBlockFloatMenu.setAttribute("aria-hidden", "true");
  }
}

/**
 * After zoom (scale on .page-stage), scroll #pageEditor so the active text box sits in the
 * middle of the visible workspace (corner / left / right / middle — same target).
 * Multiple passes fix residual offset after scroll reflow.
 */
function centerEditableBoxInEditor(boxWrapper) {
  const scroller = document.getElementById("pageEditor");
  if (!scroller || !boxWrapper) return;
  const passes = 5;
  const eps = 0.75;
  for (let p = 0; p < passes; p++) {
    const sr = scroller.getBoundingClientRect();
    const br = boxWrapper.getBoundingClientRect();
    const boxCx = br.left + br.width / 2;
    const boxCy = br.top + br.height / 2;
    const viewCx = sr.left + sr.width / 2;
    const viewCy = sr.top + sr.height / 2;
    const dX = boxCx - viewCx;
    const dY = boxCy - viewCy;
    if (Math.abs(dX) < eps && Math.abs(dY) < eps) break;
    let nextL = scroller.scrollLeft + dX;
    let nextT = scroller.scrollTop + dY;
    const maxL = Math.max(0, scroller.scrollWidth - scroller.clientWidth);
    const maxT = Math.max(0, scroller.scrollHeight - scroller.clientHeight);
    nextL = Math.max(0, Math.min(nextL, maxL));
    nextT = Math.max(0, Math.min(nextT, maxT));
    scroller.scrollLeft = nextL;
    scroller.scrollTop = nextT;
  }
}

const PDF_EDIT_ZOOM_SCALE = 1.6;

function getPdfStageScale(stage) {
  if (!stage) return 1;
  const inline = (stage.style.transform || "").trim();
  const m = inline.match(/scale\(([^)]+)\)/);
  if (m) {
    const v = parseFloat(m[1]);
    if (!Number.isNaN(v)) return v;
  }
  const cs = getComputedStyle(stage).transform;
  if (cs && cs !== "none") {
    const mm = cs.match(/matrix\(([^)]+)\)/);
    if (mm) {
      const a = parseFloat(mm[1].split(/[,\s]+/)[0]);
      if (!Number.isNaN(a)) return a;
    }
  }
  return 1;
}

function isPdfStageZoomed(stage) {
  return getPdfStageScale(stage) > 1.01;
}

function resetPdfEditZoomStage(stage) {
  if (!stage) return;
  stage.style.removeProperty("transition");
  stage.style.transform = "scale(1)";
  stage.style.width = "100%";
  stage.style.height = "auto";
  const pw = stage.closest(".page-wrapper");
  if (pw) {
    pw.style.width = "100%";
    pw.style.height = "auto";
    pw.style.zIndex = "1";
  }
}

function resetAllPdfEditZoom() {
  if (!pageEditor) return;
  pageEditor.querySelectorAll(".page-stage").forEach(resetPdfEditZoomStage);
}

function resetOtherPdfEditZoomStages(exceptStage) {
  if (!pageEditor) return;
  pageEditor.querySelectorAll(".page-stage").forEach((s) => {
    if (s !== exceptStage) resetPdfEditZoomStage(s);
  });
}

let _applyZoomLastBox = null;
let _applyZoomLastAt = 0;
const APPLY_ZOOM_DEDUPE_MS = 120;

function applyZoomToBoxContext(boxWrapper) {
  if (!boxWrapper) return;
  const t = performance.now();
  if (_applyZoomLastBox === boxWrapper && t - _applyZoomLastAt < APPLY_ZOOM_DEDUPE_MS) {
    return;
  }
  _applyZoomLastBox = boxWrapper;
  _applyZoomLastAt = t;

  const pageWrapper = boxWrapper.closest(".page-wrapper");
  const stage = pageWrapper ? pageWrapper.querySelector(".page-stage") : null;
  if (!pageWrapper || !stage) return;

  /* Avoid "double zoom": CSS transition animated scale(1)→scale(1.6) as two steps. Measure with transitions off, then one animated jump to 1.6. */
  stage.style.transition = "none";
  stage.style.transform = "scale(1)";
  void stage.offsetWidth;

  const unscaledW = stage.offsetWidth;
  const unscaledH = stage.offsetHeight;

  stage.style.width = `${unscaledW}px`;
  stage.style.height = `${unscaledH}px`;

  stage.style.transformOrigin = "0 0";
  stage.style.transform = "scale(1)";
  pageWrapper.style.zIndex = "50";
  pageWrapper.style.width = `${unscaledW * PDF_EDIT_ZOOM_SCALE}px`;
  pageWrapper.style.height = `${unscaledH * PDF_EDIT_ZOOM_SCALE}px`;

  requestAnimationFrame(() => {
    stage.style.removeProperty("transition");
    requestAnimationFrame(() => {
      stage.style.transform = `scale(${PDF_EDIT_ZOOM_SCALE})`;
    });
  });

  const centerAfterZoom = () => centerEditableBoxInEditor(boxWrapper);
  setTimeout(() => {
    centerAfterZoom();
    requestAnimationFrame(() => {
      centerAfterZoom();
      setTimeout(centerAfterZoom, 50);
    });
  }, 320);
}

/**
 * First focus on a page runs zoom; further focuses while that page stays zoomed only scroll-center the box.
 * Focusing a box on another page resets zoom on all other pages, then zooms that page if needed.
 */
function focusPdfEditBoxZoomOrCenter(boxWrapper) {
  if (!boxWrapper) return;
  const pageWrapper = boxWrapper.closest(".page-wrapper");
  const stage = pageWrapper ? pageWrapper.querySelector(".page-stage") : null;
  if (!pageWrapper || !stage) return;

  resetOtherPdfEditZoomStages(stage);

  if (isPdfStageZoomed(stage)) {
    const runCenter = () => centerEditableBoxInEditor(boxWrapper);
    runCenter();
    requestAnimationFrame(() => {
      runCenter();
      setTimeout(runCenter, 50);
    });
    return;
  }

  applyZoomToBoxContext(boxWrapper);
}

function updatePdfBlockFloatMenuPosition() {
  if (!pdfBlockFloatMenu || pdfBlockFloatMenu.classList.contains("hidden")) return;
  const box = pdfFloatMenuTargetBox;
  if (!box) return;
  const rect = box.getBoundingClientRect();
  const menu = pdfBlockFloatMenu;
  const h = menu.offsetHeight || 44;
  const w = menu.offsetWidth || 280;
  let top = rect.top - h - 8;
  if (top < 8) top = rect.bottom + 8;
  const cx = rect.left + rect.width / 2;
  const left = Math.min(Math.max(cx, w / 2 + 8), window.innerWidth - w / 2 - 8);
  menu.style.left = `${left}px`;
  menu.style.top = `${top}px`;
  menu.style.transform = "translateX(-50%)";
}

function showPdfBlockFloatMenu(boxWrapper) {
  pdfFloatMenuTargetBox = boxWrapper;
  if (!pdfBlockFloatMenu) return;
  pdfBlockFloatMenu.classList.remove("hidden");
  pdfBlockFloatMenu.setAttribute("aria-hidden", "false");
  updatePdfBlockFloatMenuPosition();
  requestAnimationFrame(() => updatePdfBlockFloatMenuPosition());
}

function clearPdfBlockSelection() {
  hidePdfBlockFloatMenu();
  document.querySelectorAll(".pdf-box-wrapper.block-selected").forEach((w) => {
    w.classList.remove("block-selected");
    const a = w.querySelector(".pdf-text-input");
    if (a) a.removeAttribute("readonly");
  });
}

function selectPdfBlock(boxWrapper) {
  const area = boxWrapper.querySelector(".pdf-text-input");
  if (!area) return;
  clearPdfBlockSelection();
  boxWrapper.classList.add("block-selected");
  area.setAttribute("readonly", "readonly");
  showPdfBlockFloatMenu(boxWrapper);
}

function getFloatMenuTargetArea() {
  if (!pdfFloatMenuTargetBox) return null;
  return pdfFloatMenuTargetBox.querySelector(".pdf-text-input");
}

/** Assign id for CSS only on the one `.active` edit textarea — see `#pdf-active-edit-input` in styles.css */
const PDF_ACTIVE_EDIT_INPUT_ID = "pdf-active-edit-input";

function syncPdfActiveEditInputId() {
  const prev = document.getElementById(PDF_ACTIVE_EDIT_INPUT_ID);
  if (prev) prev.removeAttribute("id");
  const ta = document.querySelector(".pdf-box-wrapper.active .pdf-text-input");
  if (ta) ta.id = PDF_ACTIVE_EDIT_INPUT_ID;
}

function syncFormattingToolbarFromArea(area) {
  const boldBtn = document.getElementById("formatBoldBtn");
  const italicBtn = document.getElementById("formatItalicBtn");
  const underlineBtn = document.getElementById("formatUnderlineBtn");
  const strikeBtn = document.getElementById("formatStrikeBtn");
  const isBold = area.style.fontWeight === "bold";
  const isItalic = area.style.fontStyle === "italic";
  const u = area.style.textDecoration || "";
  const isUnderline = u.includes("underline");
  const isStrike = u.includes("line-through");
  boldBtn?.classList.toggle("toolbar-btn-active", isBold);
  italicBtn?.classList.toggle("toolbar-btn-active", isItalic);
  underlineBtn?.classList.toggle("toolbar-btn-active", isUnderline);
  strikeBtn?.classList.toggle("toolbar-btn-active", isStrike);
}

function enterPdfBlockEditMode(boxWrapper) {
  const area = boxWrapper.querySelector(".pdf-text-input");
  if (!area) return;
  clearPdfBlockSelection();
  document.querySelectorAll(".pdf-box-wrapper").forEach((w) => w.classList.remove("active"));
  boxWrapper.classList.add("active");
  area.removeAttribute("readonly");
  lastActiveArea = area;
  syncFormattingToolbarFromArea(area);
  area.focus();
  syncPdfActiveEditInputId();
}

function wirePdfBlockFloatMenuOnce() {
  if (!pdfBlockFloatMenu || pdfBlockFloatMenu.dataset.wired === "1") return;
  pdfBlockFloatMenu.dataset.wired = "1";

  pdfBlockFloatMenu.addEventListener("mousedown", (e) => e.stopPropagation());
  pdfBlockFloatMenu.addEventListener("click", (e) => {
    e.stopPropagation();
    const btn = e.target.closest("[data-action]");
    if (!btn) return;
    const action = btn.getAttribute("data-action");
    if (action === "edit") {
      const box = pdfFloatMenuTargetBox;
      if (box) enterPdfBlockEditMode(box);
      return;
    }
    if (action === "select") {
      const box = pdfFloatMenuTargetBox;
      if (box) {
        enterPdfBlockEditMode(box);
        const a = box.querySelector(".pdf-text-input");
        if (a) requestAnimationFrame(() => { a.focus(); a.select(); });
      }
      return;
    }
    if (action === "copy") {
      const a = getFloatMenuTargetArea();
      if (a) {
        void navigator.clipboard.writeText(a.value || "").catch(() => {});
      }
      return;
    }
    if (action === "delete") {
      const a = getFloatMenuTargetArea();
      if (!a) return;
      const box = a.closest(".pdf-box-wrapper");
      a.value = "";
      if (box) box.classList.add("edited");
      clearPdfBlockSelection();
      pushEditorHistory();
    }
  });

  if (pageEditor && !pageEditor.dataset.pdfFloatScrollBound) {
    pageEditor.dataset.pdfFloatScrollBound = "1";
    pageEditor.addEventListener(
      "scroll",
      () => {
        if (pdfBlockFloatMenu && !pdfBlockFloatMenu.classList.contains("hidden")) {
          updatePdfBlockFloatMenuPosition();
        }
      },
      { passive: true }
    );
  }
  window.addEventListener("resize", () => {
    if (pdfBlockFloatMenu && !pdfBlockFloatMenu.classList.contains("hidden")) {
      updatePdfBlockFloatMenuPosition();
    }
  });
}

wirePdfBlockFloatMenuOnce();

/**
 * Server says PDF is encrypted (/analyze 401 or upload needs_password). Prompt for password and POST /unlock.
 * @returns {Promise<boolean>} true if unlocked, false if user cancelled
 */
async function promptUnlockUntilSuccess(fileId) {
  let success = false;
  while (!success) {
    loadingOverlay.classList.add("hidden");
    const pw = await showCustomPrompt("This PDF is protected", "", "Enter password to unlock", true);
    if (pw === null) return false;
    loadingOverlay.classList.remove("hidden");
    uploadStatus.textContent = "Unlocking PDF...";
    try {
      const unlockRes = await fetch("/unlock", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ file_id: fileId, password: pw }),
      });
      if (unlockRes.ok) {
        success = true;
      } else {
        loadingOverlay.classList.add("hidden");
        showCustomAlert("Wrong password", "Incorrect password. Please try again.", false);
      }
    } catch {
      loadingOverlay.classList.add("hidden");
      showCustomAlert("Unlock failed", "Could not reach the server.", false);
      return false;
    }
  }
  return true;
}

/** Thumbnail / title: open Preview (page images), not the editor. */
async function openPreviewFromHome(file) {
  if (!file?.id) {
    showCustomAlert("Demo", "Upload a real PDF first.", false);
    return;
  }
  loadingOverlay.classList.remove("hidden");
  uploadStatus.textContent = "Loading preview...";
  try {
    let res = await fetch(`/analyze/${file.id}`);
    if (res.status === 401) {
      const unlocked = await promptUnlockUntilSuccess(file.id);
      if (!unlocked) return;
      res = await fetch(`/analyze/${file.id}`);
      if (res.status === 401) {
        showCustomAlert("Password required", "Could not unlock this PDF.", false);
        return;
      }
    }
    if (!res.ok) {
      let msg = "Could not load preview.";
      try {
        const j = await res.json();
        if (j.detail) msg = typeof j.detail === "string" ? j.detail : JSON.stringify(j.detail);
      } catch {
        /* ignore */
      }
      throw new Error(msg);
    }
    const json = await res.json();
    const pages = json.pages || [];
    const previewView = document.getElementById("preview-view");
    const previewContainer = document.getElementById("preview-container");
    if (!previewView || !previewContainer) return;
    previewContainer.innerHTML = "";
    pages.forEach((p) => {
      const img = document.createElement("img");
      img.src = `/preview/${file.id}/${p.page}?v=${Date.now()}`;
      img.style.width = "100%";
      img.style.marginBottom = "16px";
      img.style.boxShadow = "0 2px 6px rgba(0,0,0,0.15)";
      img.style.backgroundColor = "white";
      img.alt = `Page ${p.page}`;
      previewContainer.appendChild(img);
    });
    currentFileId = file.id;
    homeView.classList.remove("active");
    editorView.classList.remove("active");
    previewView.classList.add("active");

    const finalDlBtns = [document.getElementById("final-download-btn"), document.getElementById("final-download-btn-bottom")];
    finalDlBtns.forEach((btn) => {
      if (!btn) return;
      const newBtn = btn.cloneNode(true);
      btn.replaceWith(newBtn);
      newBtn.addEventListener("click", () => {
        const fname = getDownloadFilenameForFileId(file.id);
        downloadPdfByFileId(file.id, fname);
      });
    });

    const backBtn = document.getElementById("back-to-edit-btn");
    if (backBtn) {
      const newBack = backBtn.cloneNode(true);
      backBtn.replaceWith(newBack);
      newBack.addEventListener("click", () => {
        previewView.classList.remove("active");
        homeView.classList.add("active");
      });
    }
  } catch (e) {
    showCustomAlert("Preview failed", e.message || String(e), false);
  } finally {
    loadingOverlay.classList.add("hidden");
  }
}

function renderMockFiles() {
  recentFilesList.innerHTML = "";
  if (mockFiles.length === 0) {
    recentFilesList.innerHTML =
      '<p class="file-list-empty-hint" style="text-align:center;padding:32px 20px;color:#666;font-size:15px;line-height:1.5;">No PDFs yet. Tap <strong style="color:#0070d6;">+</strong> below to upload.</p>';
    persistRecentFiles();
    return;
  }
  mockFiles.forEach((file, index) => {
    const card = document.createElement("div");
    card.className = "file-card";
    const bgStyle = file.thumb ? `style="background-image: url('${file.thumb}');"` : ``;
    
    if (index === 0) {
      // Expanded layout for the top file
      card.innerHTML = `
        <div class="fc-container">
          <div class="fc-thumb top-thumb" ${bgStyle}></div>
          <div class="fc-content">
            <div class="fc-header">
              <div class="fc-title">${file.title}</div>
              <div class="fc-date"><i class="fa-solid fa-user" style="font-size:10px; margin-right:4px;"></i>Today</div>
            </div>
            <div class="fc-vertical-actions">
              <button class="action-row-btn" aria-label="Share">
                 <i class="fa-solid fa-share-nodes"></i> <span>Share</span>
              </button>
              <button class="action-row-btn edit-file-btn" aria-label="Edit text">
                 <i class="fa-solid fa-pen-to-square"></i> <span>Edit text</span> <i class="fa-solid fa-star blue-star" style="font-size:10px; color:#0070d6; margin-left:4px;"></i>
              </button>
              <button type="button" class="action-row-btn" aria-label="Delete file">
                 <i class="fa-solid fa-trash-can"></i> <span>Delete</span>
              </button>
              <button class="action-row-btn" aria-label="More Options">
                 <i class="fa-solid fa-ellipsis-vertical"></i> <span>More</span>
              </button>
            </div>
          </div>
        </div>
      `;
    } else {
      // Compact layout for the rest
      card.className += " compact-card";
      card.innerHTML = `
        <div class="fc-container compact">
          <div class="fc-thumb compact-thumb" ${bgStyle}></div>
          <div class="fc-content compact-content">
            <div class="fc-header compact-header">
              <div class="fc-title">${file.title}</div>
              <div class="fc-date">${file.date || "Yesterday"}</div>
            </div>
            <div class="fc-horizontal-actions">
              <button class="action-icon-btn" aria-label="Share">
                 <i class="fa-solid fa-share-nodes"></i>
              </button>
              <button class="action-icon-btn edit-file-btn" aria-label="Edit text">
                 <i class="fa-solid fa-pen-to-square" style="position:relative;">
                   <i class="fa-solid fa-star" style="position:absolute; top:-2px; right:-6px; font-size:8px; color:#0070d6;"></i>
                 </i>
              </button>
              <button type="button" class="action-icon-btn" aria-label="Delete file">
                 <i class="fa-solid fa-trash-can"></i>
              </button>
              <button class="action-icon-btn" aria-label="More Options">
                 <i class="fa-solid fa-ellipsis-vertical"></i>
              </button>
            </div>
          </div>
        </div>
      `;
    }
    
    // Add logic to "edit", "share", and "more" buttons
    const editBtn = card.querySelector('.edit-file-btn');
    const shareBtn = card.querySelector('[aria-label="Share"]');
    const deleteBtn = card.querySelector('[aria-label="Delete file"]');
    const moreBtn = card.querySelector('[aria-label="More Options"]');
    const shareSheetOverlay = document.getElementById("shareSheetOverlay");
    const moreOptionsSheetOverlay = document.getElementById("moreOptionsSheetOverlay");
    
    const openEditor = async () => {
      if (file.id) {
         currentFileId = file.id;
         loadingOverlay.classList.remove("hidden");
         uploadStatus.textContent = "Loading file...";
         try {
             let res = await fetch(`/analyze/${file.id}`);
             if (res.status === 401) {
               const unlocked = await promptUnlockUntilSuccess(file.id);
               if (!unlocked) return;
               res = await fetch(`/analyze/${file.id}`);
               if (res.status === 401) {
                 showCustomAlert("Password required", "Could not unlock this PDF.", false);
                 return;
               }
             }
             if (!res.ok) throw new Error("Load failed");
             const json = await res.json();
             originalItems = json.items || [];
             pagesMeta = json.pages || [];
             await renderEditor(pagesMeta, originalItems);
             homeView.classList.remove('active');
             editorView.classList.add('active');
             initEditorHistory();
         } catch (e) {
             alert(e.message);
         } finally {
             loadingOverlay.classList.add("hidden");
         }
      } else {
         alert("This is a demo item. Please upload a real PDF.");
      }
    };

    const openShare = () => {
      shareContextFile = file;
      updateShareSheetUI();
      shareSheetOverlay.classList.remove('hidden');
    };
    
    const openMoreOptions = () => {
      selectedMoreFile = { fileObj: file, cardEl: card };
      const thumb = document.getElementById("moreOptionsThumb");
      if (thumb) thumb.style.backgroundImage = file.thumb ? `url('${file.thumb}')` : 'none';
      
      const titleEl = document.getElementById("moreOptionsTitle");
      if (titleEl) titleEl.textContent = file.title;
      
      const formatSize = (bytes) => {
        if (!bytes) return "323.2 KB";
        if (bytes < 1024) return bytes + " B";
        else if (bytes < 1048576) return (bytes / 1024).toFixed(1) + " KB";
        else return (bytes / 1048576).toFixed(1) + " MB";
      };
      
      const metaEl = document.getElementById("moreOptionsMeta");
      if (metaEl) metaEl.innerHTML = `<i class="fa-solid fa-user"></i> ${file.date || "Yesterday"} &bull; ${formatSize(file.size)}`;
      
      moreOptionsSheetOverlay?.classList.remove('hidden');
    };

    editBtn?.addEventListener('click', openEditor);
    shareBtn?.addEventListener('click', openShare);
    deleteBtn?.addEventListener('click', (e) => {
      e.stopPropagation();
      confirmAndDeleteFile(file);
    });
    moreBtn?.addEventListener('click', openMoreOptions);

    const thumbEl = card.querySelector(".fc-thumb");
    const titleEl = card.querySelector(".fc-title");
    const openPreviewFromThumbOrTitle = (e) => {
      e.preventDefault();
      e.stopPropagation();
      void openPreviewFromHome(file);
    };
    thumbEl?.addEventListener("click", openPreviewFromThumbOrTitle);
    titleEl?.addEventListener("click", openPreviewFromThumbOrTitle);
    thumbEl?.setAttribute("role", "button");
    thumbEl?.setAttribute("tabindex", "0");
    thumbEl?.setAttribute("aria-label", "Preview");
    titleEl?.setAttribute("role", "button");
    titleEl?.setAttribute("tabindex", "0");
    titleEl?.setAttribute("aria-label", "Preview");
    const keyOpen = (e) => {
      if (e.key === "Enter" || e.key === " ") {
        e.preventDefault();
        void openPreviewFromHome(file);
      }
    };
    thumbEl?.addEventListener("keydown", keyOpen);
    titleEl?.addEventListener("keydown", keyOpen);

    recentFilesList.appendChild(card);
  });
  persistRecentFiles();
}
renderMockFiles();

// Global listeners for modals
const shareSheetOverlay = document.getElementById("shareSheetOverlay");
const sharingLinkSheetOverlay = document.getElementById("sharingLinkSheetOverlay");
const moreOptionsSheetOverlay = document.getElementById("moreOptionsSheetOverlay");

wireShareSheetActions();

sharingLinkSheetOverlay?.addEventListener("click", (e) => {
  if (e.target === sharingLinkSheetOverlay) sharingLinkSheetOverlay.classList.add("hidden");
});

function showCustomAlert(title, message, isSuccess = true) {
    customAlertTitle.textContent = title;
    customAlertMessage.textContent = message;
    
    if (isSuccess) {
        alertIconBox.className = "alert-icon-box success";
        alertIconBox.innerHTML = '<i class="fa-solid fa-check"></i>';
    } else {
        alertIconBox.className = "alert-icon-box error";
        alertIconBox.innerHTML = '<i class="fa-solid fa-xmark"></i>';
    }
    
    customAlertOverlay.classList.remove("hidden");
}

customAlertOk.onclick = () => {
    customAlertOverlay.classList.add("hidden");
};

const customConfirmOverlay = document.getElementById("customConfirmOverlay");
const customConfirmTitle = document.getElementById("customConfirmTitle");
const customConfirmMessage = document.getElementById("customConfirmMessage");
const customConfirmOk = document.getElementById("customConfirmOk");
const customConfirmCancel = document.getElementById("customConfirmCancel");

let onConfirmOk = null;
function showCustomConfirm(title, message, onOk) {
    customConfirmTitle.textContent = title;
    customConfirmMessage.textContent = message;
    onConfirmOk = onOk;
    customConfirmOverlay.classList.remove("hidden");
}

customConfirmOk.onclick = () => {
    customConfirmOverlay.classList.add("hidden");
    if (onConfirmOk) onConfirmOk();
};
customConfirmCancel.onclick = () => {
    customConfirmOverlay.classList.add("hidden");
    onConfirmOk = null;
};

// Compress PDF Global Logic (Updated for Custom Alert & Confirm)
document.getElementById("compressPdfBtn")?.addEventListener('click', async () => {
    if (!selectedMoreFile || !selectedMoreFile.fileObj.id) {
        showCustomAlert("Demo File", "This is a demo item. Please upload a real PDF.", false);
        return;
    }

    const performCompression = async () => {
        moreOptionsSheetOverlay?.classList.add('hidden');
        loadingOverlay.classList.remove("hidden");
        uploadStatus.textContent = "Compressing PDF...";
        
        // Calculate progressive quality based on previous attempts
        const currentCount = selectedMoreFile.fileObj.compressionCount || 0;
        const qualities = [70, 45, 25, 10]; // Progressive reduction
        const targetQuality = qualities[currentCount] || 5; // Extreme for 5th round+

        try {
            const res = await fetch("/compress", {
                method: "POST",
                headers: { "Content-Type": "application/json" },
                body: JSON.stringify({ 
                    file_id: selectedMoreFile.fileObj.id,
                    quality: targetQuality
                })
            });
            
            if (!res.ok) throw new Error("Compression failed.");
            const json = await res.json();
            
            // Update the item in the list
            const idx = mockFiles.findIndex(f => f.id === selectedMoreFile.fileObj.id);
            if (idx > -1) {
                mockFiles[idx].id = json.file_id;
                mockFiles[idx].size = json.size;
                mockFiles[idx].compressionCount = (mockFiles[idx].compressionCount || 0) + 1;
            }
            
            renderMockFiles();
            const formatSize = (bytes) => {
                if (bytes < 1024) return bytes + " B";
                if (bytes < 1024 * 1024) return (bytes / 1024).toFixed(1) + " KB";
                return (bytes / (1024 * 1024)).toFixed(1) + " MB";
            };
            showCustomAlert("Success!", `PDF Compressed! New size: ${formatSize(json.size)}`, true);
        } catch (err) {
            showCustomAlert("Failed!", err.message, false);
        } finally {
            loadingOverlay.classList.add("hidden");
        }
    };

    const count = selectedMoreFile.fileObj.compressionCount || 0;
    if (count >= 1) {
        showCustomConfirm("Warning", "The document is already compressed. Compressing again might reduce quality. Proceed?", performCompression);
    } else {
        await performCompression();
    }
});

document.getElementById("deleteFileMoreBtn")?.addEventListener("click", () => {
  if (!selectedMoreFile || !selectedMoreFile.fileObj) return;
  confirmAndDeleteFile(selectedMoreFile.fileObj);
});

shareSheetOverlay?.addEventListener('click', (e) => {
    if (e.target === shareSheetOverlay) shareSheetOverlay.classList.add('hidden');
});

moreOptionsSheetOverlay?.addEventListener('click', (e) => {
    const isInsideSheet = e.target.closest('.share-sheet');
    const isOptionBtn = e.target.closest('.more-option-btn');
    if (e.target === moreOptionsSheetOverlay || isOptionBtn) {
        moreOptionsSheetOverlay.classList.add('hidden');
    }
});

// Custom Prompt Helper
function showCustomPrompt(title, defaultValue = "", placeholder = "", isPassword = false) {
    return new Promise((resolve) => {
        const overlay = document.getElementById("customPromptOverlay");
        const titleEl = document.getElementById("customPromptTitle");
        const inputEl = document.getElementById("customPromptInput");
        const cancelBtn = document.getElementById("customPromptCancel");
        const confirmBtn = document.getElementById("customPromptConfirm");

        titleEl.textContent = title;
        inputEl.type = isPassword ? "password" : "text";
        inputEl.value = defaultValue;
        inputEl.placeholder = placeholder;
        
        overlay.classList.remove("hidden");
        // Focus the input safely
        setTimeout(() => {
            inputEl.focus();
            if(!isPassword && inputEl.value) inputEl.setSelectionRange(0, inputEl.value.length);
        }, 50);

        const cleanup = () => {
            overlay.classList.add("hidden");
            cancelBtn.removeEventListener("click", onCancel);
            confirmBtn.removeEventListener("click", onConfirm);
            inputEl.removeEventListener("keydown", onKey);
        };

        const onCancel = () => { cleanup(); resolve(null); };
        const onConfirm = () => { cleanup(); resolve(inputEl.value); };
        const onKey = (e) => { if (e.key === "Enter") onConfirm(); };

        cancelBtn.addEventListener("click", onCancel);
        confirmBtn.addEventListener("click", onConfirm);
        inputEl.addEventListener("keydown", onKey);
    });
}

// Rename Logic
document.getElementById('renameOptionBtn')?.addEventListener('click', async () => {
    if (!selectedMoreFile) return;
    document.getElementById("moreOptionsSheetOverlay")?.classList.add("hidden");
    
    const newName = await showCustomPrompt("Rename file", selectedMoreFile.fileObj.title, "Enter new name");
    if (newName && newName.trim()) {
        const cleanName = newName.trim();
        const titleText = cleanName + (cleanName.toLowerCase().endsWith('.pdf') ? '' : '.pdf');
        selectedMoreFile.fileObj.title = titleText;
        const titleEl = selectedMoreFile.cardEl.querySelector('.fc-title');
        if (titleEl) titleEl.textContent = titleText;
        if (shareContextFile && shareContextFile.id === selectedMoreFile.fileObj.id) {
          shareContextFile.title = titleText;
        }
        persistRecentFiles();
    }
});

// Set Password Logic
document.getElementById('setPasswordBtn')?.addEventListener('click', async () => {
    if (!selectedMoreFile) return;
    const file = selectedMoreFile.fileObj;
    
    document.getElementById("moreOptionsSheetOverlay")?.classList.add("hidden");
    
    if (!file.id) {
        alert("This is a demo file. Please upload a real PDF to set a password.");
        return;
    }
    
    const pw = await showCustomPrompt("Set password", "", "Enter a strong password", true);
    if (!pw) return;
    
    loadingOverlay.classList.remove("hidden");
    uploadStatus.textContent = "Encrypting document...";
    
    try {
        const res = await fetch("/set_password", {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify({ file_id: file.id, password: pw })
        });
        
        if (!res.ok) {
            let msg = "Failed to set password.";
            try {
                const data = await res.json();
                if (data.detail != null) {
                    msg = typeof data.detail === "string" ? data.detail : JSON.stringify(data.detail);
                }
            } catch (_) { /* keep generic */ }
            throw new Error(msg);
        }
        alert("Password protected successfully! Anyone opening this will now need the password.");
    } catch (err) {
        alert("Error: " + err.message);
    } finally {
        loadingOverlay.classList.add("hidden");
    }
});

// Reorder & Delete Pages (More menu → Pages view)
const reorderView = document.getElementById("reorder-view");
const reorderGrid = document.getElementById("reorderGrid");
const reorderBackBtn = document.getElementById("reorderBackBtn");
const reorderSaveBtn = document.getElementById("reorderSaveBtn");

/** @type {{ fileId: string | null, order: number[] }} order = 0-based original page indices in display order */
let reorderState = { fileId: null, order: [] };

/** Pointer-drag: works on touch + mouse (HTML5 DnD is unreliable on mobile). */
let reorderDragActive = null; // { from: number, pointerId: number, thumb: HTMLElement } | null
let reorderDragHover = null; // number | null

function reorderClearDragVisuals() {
  reorderGrid?.querySelectorAll(".reorder-item").forEach((el) => {
    el.classList.remove("reorder-item--dragging", "reorder-item--drop-target");
  });
}

function reorderIndexFromPoint(clientX, clientY) {
  const items = reorderGrid ? [...reorderGrid.querySelectorAll(".reorder-item")] : [];
  if (!items.length) return 0;
  for (let i = 0; i < items.length; i++) {
    const r = items[i].getBoundingClientRect();
    if (clientX >= r.left && clientX <= r.right && clientY >= r.top && clientY <= r.bottom) {
      return i;
    }
  }
  let best = 0;
  let bestD = Infinity;
  items.forEach((el, i) => {
    const r = el.getBoundingClientRect();
    const cx = (r.left + r.right) / 2;
    const cy = (r.top + r.bottom) / 2;
    const d = (clientX - cx) ** 2 + (clientY - cy) ** 2;
    if (d < bestD) {
      bestD = d;
      best = i;
    }
  });
  return best;
}

function reorderApplyMove(from, to) {
  if (from === to || from == null || to == null) return;
  const arr = reorderState.order;
  const [moved] = arr.splice(from, 1);
  arr.splice(to, 0, moved);
  renderReorderGrid();
}

function renderReorderGrid() {
  if (!reorderGrid || !reorderState.fileId) return;
  reorderGrid.innerHTML = "";
  const fid = reorderState.fileId;
  const reorderCols =
    typeof window !== "undefined" &&
    window.matchMedia &&
    window.matchMedia("(max-width: 480px)").matches
      ? 1
      : 2;
  const n = reorderState.order.length;

  reorderState.order.forEach((zeroBasedIdx, position) => {
    const pageNum = zeroBasedIdx + 1;
    const item = document.createElement("div");
    item.className = "reorder-item";
    item.dataset.listIndex = String(position);

    const delBtn = document.createElement("div");
    delBtn.className = "reorder-delete-btn";
    delBtn.setAttribute("role", "button");
    delBtn.innerHTML = '<i class="fa-solid fa-times"></i>';
    delBtn.addEventListener("pointerdown", (e) => e.stopPropagation());
    delBtn.addEventListener("click", (e) => {
      e.stopPropagation();
      e.preventDefault();
      if (reorderState.order.length <= 1) {
        showCustomAlert("Cannot delete", "A PDF must keep at least one page.", false);
        return;
      }
      reorderState.order.splice(position, 1);
      renderReorderGrid();
    });

    const thumb = document.createElement("img");
    thumb.className = "reorder-thumb";
    thumb.draggable = false;
    thumb.loading = "lazy";
    thumb.decoding = "async";
    thumb.alt = `Page ${pageNum}`;
    thumb.src = `/preview/${fid}/${pageNum}?v=${Date.now()}`;

    const onPointerDown = (e) => {
      if (e.button != null && e.button !== 0) return;
      reorderDragActive = { from: position, pointerId: e.pointerId, thumb };
      reorderDragHover = position;
      try {
        thumb.setPointerCapture(e.pointerId);
      } catch (_) {
        /* ignore */
      }
      item.classList.add("reorder-item--dragging");
    };

    const onPointerMove = (e) => {
      if (!reorderDragActive || reorderDragActive.pointerId !== e.pointerId) return;
      e.preventDefault();
      const hi = reorderIndexFromPoint(e.clientX, e.clientY);
      reorderDragHover = hi;
      reorderGrid.querySelectorAll(".reorder-item").forEach((el, i) => {
        el.classList.toggle("reorder-item--drop-target", i === hi && hi !== reorderDragActive.from);
      });
    };

    const finishPointerDrag = (e, commit) => {
      if (!reorderDragActive || reorderDragActive.pointerId !== e.pointerId) return;
      const from = reorderDragActive.from;
      const to = reorderDragHover;
      reorderDragActive = null;
      reorderDragHover = null;
      reorderClearDragVisuals();
      try {
        thumb.releasePointerCapture(e.pointerId);
      } catch (_) {
        /* ignore */
      }
      if (commit && to != null && from !== to) reorderApplyMove(from, to);
    };

    thumb.addEventListener("pointerdown", onPointerDown);
    thumb.addEventListener("pointermove", onPointerMove, { passive: false });
    thumb.addEventListener("pointerup", (e) => finishPointerDrag(e, true));
    thumb.addEventListener("pointercancel", (e) => finishPointerDrag(e, false));

    thumb.addEventListener("dragstart", (e) => e.preventDefault());

    const moveRow = document.createElement("div");
    moveRow.className = "reorder-move-row";

    const swapOrder = (a, b) => {
      const arr = reorderState.order;
      [arr[a], arr[b]] = [arr[b], arr[a]];
      renderReorderGrid();
    };

    const btnLeft = document.createElement("button");
    btnLeft.type = "button";
    btnLeft.className = "reorder-move-btn";
    btnLeft.setAttribute("aria-label", "Move page left");
    btnLeft.innerHTML = '<i class="fa-solid fa-chevron-left"></i>';
    if (reorderCols === 2) {
      btnLeft.disabled = position % 2 === 0;
    } else {
      btnLeft.disabled = position === 0;
    }
    btnLeft.addEventListener("click", (e) => {
      e.stopPropagation();
      if (reorderCols === 2) {
        if (position % 2 !== 1) return;
        swapOrder(position, position - 1);
      } else {
        if (position <= 0) return;
        swapOrder(position, position - 1);
      }
    });

    const btnUp = document.createElement("button");
    btnUp.type = "button";
    btnUp.className = "reorder-move-btn";
    btnUp.setAttribute("aria-label", "Move page up");
    btnUp.innerHTML = '<i class="fa-solid fa-chevron-up"></i>';
    btnUp.disabled = position === 0;
    btnUp.addEventListener("click", (e) => {
      e.stopPropagation();
      if (position <= 0) return;
      swapOrder(position, position - 1);
    });

    const btnDown = document.createElement("button");
    btnDown.type = "button";
    btnDown.className = "reorder-move-btn";
    btnDown.setAttribute("aria-label", "Move page down");
    btnDown.innerHTML = '<i class="fa-solid fa-chevron-down"></i>';
    btnDown.disabled = position === n - 1;
    btnDown.addEventListener("click", (e) => {
      e.stopPropagation();
      if (position >= n - 1) return;
      swapOrder(position, position + 1);
    });

    const btnRight = document.createElement("button");
    btnRight.type = "button";
    btnRight.className = "reorder-move-btn";
    btnRight.setAttribute("aria-label", "Move page right");
    btnRight.innerHTML = '<i class="fa-solid fa-chevron-right"></i>';
    if (reorderCols === 2) {
      btnRight.disabled = position % 2 === 1 || position + 1 >= n;
    } else {
      btnRight.disabled = position >= n - 1;
    }
    btnRight.addEventListener("click", (e) => {
      e.stopPropagation();
      if (reorderCols === 2) {
        if (position % 2 !== 0 || position + 1 >= n) return;
        swapOrder(position, position + 1);
      } else {
        if (position >= n - 1) return;
        swapOrder(position, position + 1);
      }
    });

    moveRow.appendChild(btnLeft);
    moveRow.appendChild(btnUp);
    moveRow.appendChild(btnDown);
    moveRow.appendChild(btnRight);

    const label = document.createElement("div");
    label.className = "reorder-page-num";
    label.textContent = `Page ${pageNum}`;

    item.appendChild(delBtn);
    item.appendChild(thumb);
    item.appendChild(moveRow);
    item.appendChild(label);
    reorderGrid.appendChild(item);
  });
}

function openReorderForFileId(fileId) {
  reorderState.fileId = fileId;
  reorderState.order = [];
}

document.getElementById("reorderOptionBtn")?.addEventListener("click", async () => {
  if (!selectedMoreFile || !selectedMoreFile.fileObj.id) {
    showCustomAlert("Demo file", "Upload a real PDF to reorder or delete pages.", false);
    return;
  }

  const fid = selectedMoreFile.fileObj.id;
  loadingOverlay.classList.remove("hidden");
  uploadStatus.textContent = "Loading pages...";

  try {
    const res = await fetch(`/analyze/${fid}`);
    if (!res.ok) throw new Error("Could not load PDF pages.");
    const json = await res.json();
    const pages = json.pages || [];
    if (!pages.length) throw new Error("No pages in this PDF.");

    openReorderForFileId(fid);
    reorderState.order = pages.map((_, i) => i);
    renderReorderGrid();

    homeView.classList.remove("active");
    editorView.classList.remove("active");
    combineView?.classList.remove("active");
    selectFilesView?.classList.remove("active");
    document.getElementById("preview-view")?.classList.remove("active");
    reorderView?.classList.add("active");
  } catch (err) {
    showCustomAlert("Error", err.message || String(err), false);
  } finally {
    loadingOverlay.classList.add("hidden");
  }
});

reorderBackBtn?.addEventListener("click", () => {
  reorderView?.classList.remove("active");
  homeView.classList.add("active");
  reorderState = { fileId: null, order: [] };
});

reorderSaveBtn?.addEventListener("click", async () => {
  if (!reorderState.fileId || !reorderState.order.length) return;

  loadingOverlay.classList.remove("hidden");
  uploadStatus.textContent = "Saving page order...";

  try {
    const res = await fetch("/reorder", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        file_id: reorderState.fileId,
        page_indices: reorderState.order,
      }),
    });

    if (!res.ok) {
      const err = await res.json().catch(() => ({}));
      throw new Error(err.detail || "Reorder failed.");
    }

    const json = await res.json();
    const oldId = reorderState.fileId;
    const idx = mockFiles.findIndex((f) => f.id === oldId);
    if (idx > -1) {
      mockFiles[idx].id = json.file_id;
      mockFiles[idx].size = json.size ?? mockFiles[idx].size;
      mockFiles[idx].thumb = `/preview/${json.file_id}/1?v=${Date.now()}`;
    }

    renderMockFiles();
    reorderView?.classList.remove("active");
    homeView.classList.add("active");
    reorderState = { fileId: null, order: [] };
    showCustomAlert("Saved", "Page order and deletions were applied.", true);
  } catch (err) {
    showCustomAlert("Failed", err.message || String(err), false);
  } finally {
    loadingOverlay.classList.add("hidden");
  }
});

// Save as JPEG / PNG / WebP (ZIP)
function wireSavePagesAsFormat(btnId, path, statusText) {
  document.getElementById(btnId)?.addEventListener("click", async () => {
    if (!selectedMoreFile) return;
    const file = selectedMoreFile.fileObj;

    document.getElementById("moreOptionsSheetOverlay")?.classList.add("hidden");

    if (!file.id) {
      alert("This is a demo file. Please upload a real PDF to convert.");
      return;
    }

    loadingOverlay.classList.remove("hidden");
    uploadStatus.textContent = statusText;

    try {
      window.location.href = `${path}/${file.id}`;
      setTimeout(() => loadingOverlay.classList.add("hidden"), 3000);
    } catch (err) {
      alert("Error: " + err.message);
      loadingOverlay.classList.add("hidden");
    }
  });
}

wireSavePagesAsFormat("saveAsJpegBtn", "/convert_to_jpeg", "Converting pages to JPEG…");
wireSavePagesAsFormat("saveAsPngBtn", "/convert_to_png", "Converting pages to PNG…");
wireSavePagesAsFormat("saveAsWebpBtn", "/convert_to_webp", "Converting pages to WebP…");

// Combine Workflow UI Logic
const combineView = document.getElementById("combineView");
const selectFilesView = document.getElementById("selectFilesView");
const combineList = document.getElementById("combineList");
const selectFilesList = document.getElementById("selectFilesList");
const selectedCount = document.getElementById("selectedCount");
const finalCombineBtn = document.getElementById("finalCombineBtn");

function renderCombineList() {
    combineList.innerHTML = "";
    if (selectedCombineFiles.length === 0) {
        combineList.innerHTML = `
            <div style="text-align:center; padding: 40px 20px; color:#999;">
                <i class="fa-solid fa-file-pdf" style="font-size:48px; margin-bottom:16px; opacity:0.3;"></i>
                <p>No files added to combine yet.</p>
            </div>
        `;
    }
    
    selectedCombineFiles.forEach((file, index) => {
        const item = document.createElement("div");
        item.className = "combine-item";
        const thumbUrl = file.thumb || "";
        const thumbStyle = thumbUrl ? `background-image:url('${thumbUrl}')` : `background-color:#f8f8f8; display:flex; align-items:center; justify-content:center;`;
        const thumbContent = thumbUrl ? "" : `<i class="fa-solid fa-file-pdf" style="color:#ff4d4d; font-size:18px;"></i>`;

        item.innerHTML = `
            <div style="color:#ff4d4d; font-size:22px; margin-right:8px; cursor:pointer;" onclick="removeFromCombine(${index})">
                <i class="fa-solid fa-circle-minus"></i>
            </div>
            <div class="combine-item-thumb" style="${thumbStyle}">${thumbContent}</div>
            <div class="combine-item-info">
                <div class="combine-item-title">${file.title}</div>
                <div class="combine-item-meta">${file.date || "Today"}</div>
            </div>
            <div style="color:#999; font-size:18px;"><i class="fa-solid fa-bars"></i></div>
        `;
        combineList.appendChild(item);
    });
    
    selectedCount.textContent = `${selectedCombineFiles.length} file${selectedCombineFiles.length !== 1 ? 's' : ''}`;
    finalCombineBtn.style.opacity = selectedCombineFiles.length >= 1 ? "1" : "0.5";
    
    // Auto-scroll to bottom so user sees newly added items and buttons
    const container = document.getElementById("combineListContainer");
    if (container) {
        setTimeout(() => {
            container.scrollTop = container.scrollHeight;
        }, 50);
    }
}

window.removeFromCombine = (index) => {
    selectedCombineFiles.splice(index, 1);
    renderCombineList();
};

function renderSelectFilesList() {
    selectFilesList.innerHTML = "";
    mockFiles.forEach(file => {
        const isSelected = selectedCombineFiles.some(f => f.id === file.id && f.id !== undefined);
        const item = document.createElement("div");
        item.className = `selectable-file-item ${isSelected ? 'selected' : ''}`;
        
        const formatSize = (bytes) => {
          if (!bytes) return "323.2 KB";
          if (bytes < 1024) return bytes + " B";
          else if (bytes < 1048576) return (bytes / 1024).toFixed(1) + " KB";
          else return (bytes / 1048576).toFixed(1) + " MB";
        };

        const thumbUrl = file.thumb || "";
        const thumbStyle = thumbUrl ? `background-image:url('${thumbUrl}')` : `background-color:#f8f8f8; display:flex; align-items:center; justify-content:center;`;
        const thumbContent = thumbUrl ? "" : `<i class="fa-solid fa-file-pdf" style="color:#ff4d4d; font-size:24px;"></i>`;

        item.innerHTML = `
            <div class="custom-checkbox"></div>
            <div class="combine-item-thumb" style="width:44px; height:58px; ${thumbStyle}">${thumbContent}</div>
            <div class="combine-item-info">
                <div class="combine-item-title" style="font-size:15px; color:#222;">${file.title}</div>
                <div class="combine-item-meta" style="font-size:13px; color:#777;">${file.date || "Today"} &bull; ${formatSize(file.size)}</div>
            </div>
        `;
        
        item.onclick = () => {
            const idx = selectedCombineFiles.findIndex(f => f.id === file.id);
            if (idx > -1) {
                selectedCombineFiles.splice(idx, 1);
                item.classList.remove('selected');
            } else {
                selectedCombineFiles.push(file);
                item.classList.add('selected');
            }
        };
        selectFilesList.appendChild(item);
    });
}

document.getElementById('combineFilesBtn')?.addEventListener('click', () => {
    document.getElementById("moreOptionsSheetOverlay")?.classList.add("hidden");
    if (selectedMoreFile) {
        selectedCombineFiles = [selectedMoreFile.fileObj];
        document.getElementById("combineFileName").value = selectedMoreFile.fileObj.title.replace(/\.pdf$/i, "");
    }
    homeView.classList.remove("active");
    combineView.classList.add("active");
    renderCombineList();
});

document.getElementById("closeCombineBtn")?.addEventListener("click", () => {
    combineView.classList.remove("active");
    homeView.classList.add("active");
});

document.getElementById("addExistingFilesBtn")?.addEventListener("click", () => {
    combineView.classList.remove("active");
    selectFilesView.classList.add("active");
    renderSelectFilesList();
});

document.getElementById("addFilesBtn")?.addEventListener("click", () => {
    document.getElementById("combineNewFilesInput").click();
});

// Allow adding brand new files from Phone in Selection logic
const combineNewFilesInput = document.getElementById("combineNewFilesInput");
combineNewFilesInput?.addEventListener("change", async (e) => {
    const files = e.target.files;
    if (!files || files.length === 0) return;
    
    const uploadEach = async (file) => {
        const d = new FormData();
        d.append("file", file);
        const r = await fetch("/upload", { method: "POST", body: d });
        if (!r.ok) {
            const errJson = await r.json();
            alert(`File ${file.name} failed: ${errJson.detail || 'Unknown error'}`);
            return null;
        }
        const j = await r.json();
        return {
            id: j.file_id,
            title: file.name,
            size: file.size,
            date: "Today",
            thumb: `/preview/${j.file_id}/1?v=${Date.now()}`
        };
    };

    loadingOverlay.classList.remove("hidden");

    try {
        for (let i = 0; i < files.length; i++) {
            uploadStatus.textContent = `Uploading file ${i+1} of ${files.length}...`;
            const newItem = await uploadEach(files[i]);
            if (newItem) {
                selectedCombineFiles.push(newItem);
                mockFiles.push(newItem);
            }
            renderCombineList();
        }
    } catch (err) {
        alert(err.message);
    } finally {
        loadingOverlay.classList.add("hidden");
        combineNewFilesInput.value = "";
    }
});

document.getElementById("backToCombineBtn")?.addEventListener("click", () => {
    selectFilesView.classList.remove("active");
    combineView.classList.add("active");
    renderCombineList();
});

document.getElementById("nextToCombineBtn")?.addEventListener("click", () => {
    selectFilesView.classList.remove("active");
    combineView.classList.add("active");
    renderCombineList();
});

// Clear combine file name button logic
document.querySelector(".combine-name-container .fa-xmark")?.addEventListener("click", () => {
    document.getElementById("combineFileName").value = "";
    document.getElementById("combineFileName").focus();
});

finalCombineBtn?.addEventListener('click', async () => {
    if (selectedCombineFiles.length < 1) return;
    
    loadingOverlay.classList.remove("hidden");
    uploadStatus.textContent = "Combining documents...";
    
    try {
        const validIds = selectedCombineFiles.filter(f => f.id).map(f => f.id);
        if (validIds.length < 2) {
            throw new Error("Pase select at least 2 real PDF files. Mock files cannot be combined.");
        }

        const res = await fetch("/combine", {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify({ file_ids: validIds })
        });
        
        if (!res.ok) {
            let errMsg = "Combining failed.";
            try {
                const contentType = res.headers.get('content-type');
                if (contentType && contentType.includes('application/json')) {
                    const errJson = await res.json();
                    errMsg = errJson.detail || errMsg;
                } else {
                    const rawText = await res.text();
                    errMsg = rawText || errMsg;
                }
            } catch (inner) {
                console.error("Error parsing backend error:", inner);
            }
            throw new Error(errMsg);
        }
        const json = await res.json();
        
        const combinedTitle = document.getElementById("combineFileName").value.trim() || "Combined PDF";
        const finalTitle = combinedTitle.toLowerCase().endsWith(".pdf") ? combinedTitle : combinedTitle + ".pdf";

        const newItem = {
            id: json.file_id,
            title: finalTitle,
            size: json.size || 0,
            date: "Today",
            thumb: `/preview/${json.file_id}/1?v=${Date.now()}`
        };
        
        mockFiles.unshift(newItem);
        selectedCombineFiles = []; // Reset combine list
        renderMockFiles(); // Update Home in background
        
        // DIRECTLY OPEN IN EDITOR
        currentFileId = json.file_id;
        uploadStatus.textContent = "Opening combined document...";
        const analyzeRes = await fetch(`/analyze/${currentFileId}`);
        if (!analyzeRes.ok) throw new Error("Could not analyze combined PDF.");
        const analyzeJson = await analyzeRes.json();
        
        originalItems = analyzeJson.items || [];
        pagesMeta = analyzeJson.pages || [];
        await renderEditor(pagesMeta, originalItems);
        
        combineView.classList.remove("active");
        editorView.classList.add("active");
        initEditorHistory();
    } catch (err) {
        alert(err.message);
    } finally {
        loadingOverlay.classList.add("hidden");
    }
});

// Upload logic mapping Home View to Backend Edit logic
pdfFile.addEventListener("change", async (e) => {
  const files = e.target.files;
  if (!files || files.length === 0) return;

  loadingOverlay.classList.remove("hidden");
  uploadStatus.textContent = files.length > 1 ? `Processing ${files.length} files...` : "Uploading document...";

  const data = new FormData();
  for (let i = 0; i < files.length; i++) {
    data.append("files", files[i]);
  }

  try {
    const uploadRes = await fetch("/upload_multiple", { method: "POST", body: data });
    if (!uploadRes.ok) throw new Error("Upload failed.");
    const uploadJson = await uploadRes.json();
    currentFileId = uploadJson.file_id;
    const firstFile = files[0];
    const totalSize = Array.from(files).reduce((acc, f) => acc + f.size, 0);

    // Password Check logic (same unlock prompt as preview / edit from home)
    if (uploadJson.needs_password) {
        const unlocked = await promptUnlockUntilSuccess(currentFileId);
        if (!unlocked) throw new Error("Upload cancelled (PDF is encrypted)");
    }

    uploadStatus.textContent = "Analyzing structure...";
    const analyzeRes = await fetch(`/analyze/${currentFileId}`);
    if (!analyzeRes.ok) throw new Error("Could not analyze PDF.");
    const analyzeJson = await analyzeRes.json();
    
    // Create new entry for uploaded file
    const today = new Date();
    const mm = String(today.getMonth() + 1).padStart(2, '0');
    const dd = String(today.getDate()).padStart(2, '0');
    const yy = String(today.getFullYear()).slice(-2);
    
    mockFiles.unshift({
      id: currentFileId,
      title: files.length > 1 ? `Combined (${files.length} images)` : firstFile.name,
      size: totalSize,
      date: `${dd}/${mm}/${yy}`,
      thumb: `/preview/${currentFileId}/1?v=${Date.now()}`
    });
    renderMockFiles();
    
    // Store original text state
    originalItems = analyzeJson.items || [];
    pagesMeta = analyzeJson.pages || [];
    
    // Prepare editor view
    await renderEditor(pagesMeta, originalItems);
    
    homeView.classList.remove("active");
    editorView.classList.add("active");
    initEditorHistory();

  } catch (err) {
    alert(err.message);
  } finally {
    loadingOverlay.classList.add("hidden");
    pdfFile.value = ''; // Reset input to allow re-uploading the same file
  }
});

// Render the editor content dynamically as an overlay on the PDF image
async function renderEditor(pages, items) {
  pageEditor.innerHTML = "";
  textAreas = []; // Reset global registry before re-renders

  if (!pages.length) {
    pageEditor.innerHTML = "<p style='text-align:center; padding: 20px; font-weight:500; color:#555;'>No pages found.</p>";
    return;
  }

  const itemsByPage = new Map();
  for (const item of items) {
    if (!itemsByPage.has(item.page)) {
      itemsByPage.set(item.page, []);
    }
    itemsByPage.get(item.page).push(item);
  }

  for (const page of pages) {
    const pageWrapper = document.createElement("div");
    pageWrapper.className = "page-wrapper";

    const stage = document.createElement("div");
    stage.className = "page-stage";
    stage.style.aspectRatio = `${page.width} / ${page.height}`;

    const img = document.createElement("img");
    img.className = "page-image";
    img.src = `/preview/${currentFileId}/${page.page}`;
    img.alt = `Page ${page.page}`;

    const overlay = document.createElement("div");
    overlay.className = "overlay";
    overlay.dataset.page = String(page.page);
    const boxWrappers = []; // Locals stay local for per-page events

    function mapFont(fontName) {
      const f = (fontName || "").toLowerCase();
      if (f.includes("times")) return '"Times New Roman", Times, serif';
      if (f.includes("symbol") || f.includes("dingbat")) return 'serif';
      return 'sans-serif';
    }

    const pageItems = itemsByPage.get(page.page) || [];
    for (const item of pageItems) {
      const [x0, y0, x1, y1] = item.bbox;
      const left = (x0 / page.width) * 100;
      const top = (y0 / page.height) * 100;
      const width = Math.max(((x1 - x0) / page.width) * 100, 1) + 2; 
      const height = Math.max(((y1 - y0) / page.height) * 100, 1) + 1;

      // Create interactive wrapper box
      const boxWrapper = document.createElement("div");
      boxWrapper.className = "pdf-box-wrapper";
      boxWrapper.style.left = `${left}%`;
      boxWrapper.style.top = `${top}%`;
      boxWrapper.style.width = `${width}%`;
      boxWrapper.style.height = `${height}%`;

      const area = document.createElement("textarea");
      area.className = "pdf-text-input";
      area.value = item.text;
      area.dataset.originalText = item.text;
      area.dataset.id = item.id;
      area.dataset.page = String(item.page);
      area.dataset.originalBbox = JSON.stringify(item.original_bbox ?? item.bbox);
      area.dataset.bbox = JSON.stringify(item.bbox);
      area.dataset.x0 = String(x0);
      area.dataset.y0 = String(y0);
      area.dataset.x1 = String(x1);
      area.dataset.y1 = String(y1);
      area.dataset.font = item.font || "helv";
      area.dataset.size = String(item.size || 11);

      area.style.fontFamily = mapFont(item.font);
      area.style.lineHeight = String(PDF_OVERLAY_LINE_HEIGHT);

      boxWrapper.appendChild(area);

      // Add 4 mobile-friendly drag resize handles + 1 move handle
      const corners = ['top-left', 'top-right', 'bottom-left', 'bottom-right', 'move'];
      corners.forEach(corner => {
        const handle = document.createElement('div');
        if (corner === 'move') {
            handle.className = `move-handle`;
            handle.innerHTML = '<i class="fa-solid fa-up-down-left-right"></i>';
        } else {
            handle.className = `resize-handle rh-${corner}`;
        }
        boxWrapper.appendChild(handle);
        
        let isDragging = false;
        let startX, startY, startW, startH, startL, startT;

        const startDrag = (e) => {
            isDragging = true;
            e.stopPropagation();
            e.preventDefault();
            const evt = e.touches ? e.touches[0] : e;
            startX = evt.clientX;
            startY = evt.clientY;
            
            // Use DOM metrics instead of style.left/top because style values are in percentages (%)
            // parseFloat("40%") returns 40, which completely corrupts pixel math.
            // offsetLeft/Top always returns exact layout pixels safely.
            startW = boxWrapper.offsetWidth;
            startH = boxWrapper.offsetHeight;
            startL = boxWrapper.offsetLeft;
            startT = boxWrapper.offsetTop;
            
            let currentScale = 1;
            const match = stage.style.transform.match(/scale\(([^)]+)\)/);
            if (match) currentScale = parseFloat(match[1]);

            const onDrag = (e) => {
                if (!isDragging) return;
                e.preventDefault();
                const evt = e.touches ? e.touches[0] : e;
                
                // Adjust deltas by the active zoom scale
                const dx = (evt.clientX - startX) / currentScale;
                const dy = (evt.clientY - startY) / currentScale;
                
                let newW = startW, newH = startH, newL = startL, newT = startT;
                
                if (corner === 'move') {
                    newL = startL + dx;
                    newT = startT + dy;
                } else {
                    if (corner.includes('right')) newW = Math.max(10, startW + dx);
                    if (corner.includes('bottom')) newH = Math.max(10, startH + dy);
                    if (corner.includes('left')) { newW = Math.max(10, startW - dx); newL = startL + dx; }
                    if (corner.includes('top')) { newH = Math.max(10, startH - dy); newT = startT + dy; }
                }
                
                boxWrapper.style.width = `${newW}px`;
                boxWrapper.style.height = `${newH}px`;
                boxWrapper.style.left = `${newL}px`;
                boxWrapper.style.top = `${newT}px`;

                // Calculate back to PDF points
                const stageRectInner = stage.getBoundingClientRect();
                const pxPerPoint = (stageRectInner.height / currentScale) / page.height; 
                
                const ptX0 = newL / pxPerPoint;
                const ptY0 = newT / pxPerPoint;
                const ptX1 = (newL + newW) / pxPerPoint;
                const ptY1 = (newT + newH) / pxPerPoint;
                
                area.dataset.bbox = JSON.stringify([ptX0, ptY0, ptX1, ptY1]);
                area.dataset.x0 = ptX0;
                area.dataset.y0 = ptY0;
                area.dataset.x1 = ptX1;
                area.dataset.y1 = ptY1;
            };

            const stopDrag = () => {
                isDragging = false;
                document.removeEventListener('mousemove', onDrag);
                document.removeEventListener('touchmove', onDrag);
                document.removeEventListener('mouseup', stopDrag);
                document.removeEventListener('touchend', stopDrag);
            };

            document.addEventListener('mousemove', onDrag, {passive: false});
            document.addEventListener('touchmove', onDrag, {passive: false});
            document.addEventListener('mouseup', stopDrag);
            document.addEventListener('touchend', stopDrag);
        };

        handle.addEventListener('mousedown', startDrag);
        handle.addEventListener('touchstart', startDrag, {passive: false});
      });

      boxWrapper.addEventListener('mousedown', (e) => {
        if (e.target.closest('.resize-handle')) return;
        if (e.target.closest('.move-handle')) return;
        if (boxWrapper.classList.contains('active')) return;
        e.preventDefault();
      }, true);

      boxWrapper.addEventListener('click', (e) => {
        if (e.target.closest('.resize-handle')) return;
        if (e.target.closest('.move-handle')) return;
        if (boxWrapper.classList.contains('active')) return;
        e.stopPropagation();
        selectPdfBlock(boxWrapper);
      });

      area.addEventListener("focus", () => {
        lastActiveArea = area;
        document.querySelectorAll('.pdf-box-wrapper').forEach(w => w.classList.remove('active'));
        boxWrapper.classList.remove('block-selected');
        boxWrapper.classList.add('active');
        area.removeAttribute('readonly');
        syncFormattingToolbarFromArea(area);
        focusPdfEditBoxZoomOrCenter(boxWrapper);
        syncPdfActiveEditInputId();
      });
      
      area.addEventListener("blur", (e) => {
        if (isFormatting || (e.relatedTarget && e.relatedTarget.closest('.editor-bottom-wrap'))) {
            setTimeout(() => { area.focus(); isFormatting = false; }, 50);
            return;
        }
        
        setTimeout(() => {
           if(document.activeElement !== area) {
               boxWrapper.classList.remove('active');
               syncPdfActiveEditInputId();
           }
        }, 150);
      });

      overlay.appendChild(boxWrapper);
      boxWrappers.push({box: boxWrapper, area: area});
      textAreas.push(area);
    }

    stage.appendChild(img);
    stage.appendChild(overlay);
    pageWrapper.appendChild(stage);
    pageEditor.appendChild(pageWrapper);

    // Previous misplaced compression logic removed from here


    // Initial resize to convert layout to exact absolute pixels
    requestAnimationFrame(() => {
      const stageRect = stage.getBoundingClientRect();
      if (!stageRect.height || stageRect.height <= 0) return;
      
      // Since it runs once initially without scale, stageRect is accurate unscaled size
      const pxPerPoint = stageRect.height / page.height;

      function pxPerPointLive() {
        const m = stage.style.transform.match(/scale\(([^)]+)\)/);
        const sc = m ? parseFloat(m[1]) : 1;
        const sr = stage.getBoundingClientRect();
        return (sr.height / sc) / page.height;
      }

      for (const obj of boxWrappers) {
        const area = obj.area;
        const box = obj.box;
        
        const sz = Number(area.dataset.size || 11);
        let fontPx = sz * pxPerPoint;
        const linePx = fontPx * PDF_OVERLAY_LINE_HEIGHT;
        area.style.fontSize = `${fontPx}px`;
        area.style.lineHeight = `${linePx}px`;

        const x0 = Number(area.dataset.x0);
        const y0 = Number(area.dataset.y0);
        const x1 = Number(area.dataset.x1);
        const y1 = Number(area.dataset.y1);
        
        const leftPx = x0 * pxPerPoint;
        const topPx = y0 * pxPerPoint;
        const baseWidthPx = (x1 - x0) * pxPerPoint + PDF_OVERLAY_WIDTH_EXTRA_PX;
        const baseHeightPx = Math.max(
          (y1 - y0) * pxPerPoint,
          fontPx * PDF_OVERLAY_MIN_BOX_HT_EM
        );

        box.style.left = `${leftPx}px`;
        box.style.top = `${topPx}px`;
        box.style.width = `${baseWidthPx}px`;
        box.style.height = `${baseHeightPx}px`;

        let prevScrollH = area.scrollHeight;
        let prevNewlineCount = (area.value.match(/\n/g) || []).length;
        let prevValueLen = area.value.length;

        area.addEventListener("input", () => {
           // Toggle edited state class for white background (hiding original PDF text behind it)
           if (area.value !== area.dataset.originalText) {
               box.classList.add('edited');
           } else {
               box.classList.remove('edited');
           }

           const stageMaxW = Math.max(baseWidthPx, stage.clientWidth - leftPx - 4);
           const needW = pdfOverlayUnwrappedContentWidthPx(area);
           const targetW = Math.min(Math.max(baseWidthPx, needW), stageMaxW);
           const widthCapped = needW > stageMaxW - 0.5;
           const currentW = parseFloat(box.style.width);
           if (Math.abs(targetW - currentW) > 0.5) {
             box.style.width = `${targetW}px`;
             const pppW = pxPerPointLive();
             const innerW = Math.max(0.1, targetW - PDF_OVERLAY_WIDTH_EXTRA_PX);
             const newX1 = Number(area.dataset.x0) + innerW / pppW;
             area.dataset.x1 = String(newX1);
             area.dataset.bbox = JSON.stringify([
               Number(area.dataset.x0),
               Number(area.dataset.y0),
               newX1,
               Number(area.dataset.y1),
             ]);
             void area.offsetHeight;
           }

           const currentH = parseFloat(box.style.height);
           const scrollH = area.scrollHeight;
           const linePxNum = parseFloat(area.style.lineHeight) || linePx;
           const nlc = (area.value.match(/\n/g) || []).length;
           const vlen = area.value.length;
           const enterAdded = nlc > prevNewlineCount;
           const wrapGain = scrollH - prevScrollH >= linePxNum * 0.65;
           const lineRemoved = nlc < prevNewlineCount;
           const textShortened = vlen < prevValueLen;

           let targetH = currentH;
           if (scrollH > currentH && (enterAdded || (wrapGain && widthCapped))) {
             targetH = scrollH;
           } else if (scrollH < currentH || lineRemoved || textShortened) {
             /* Tall box + height:100% textarea: scrollHeight often stays ~box height, so shrink never ran. Collapse to min height, read true content scrollHeight, restore. */
             const savedH = box.style.height;
             box.style.height = `${baseHeightPx}px`;
             void area.offsetHeight;
             const measured = area.scrollHeight;
             box.style.height = savedH;
             void area.offsetHeight;
             targetH = Math.max(baseHeightPx, measured);
           }

           if (Math.abs(targetH - currentH) > 0.5) {
             box.style.height = `${targetH}px`;
             const ppp = pxPerPointLive();
             const ptY1 = Number(area.dataset.y0) + (targetH / ppp);
             area.dataset.y1 = String(ptY1);
             area.dataset.bbox = JSON.stringify([
               Number(area.dataset.x0),
               Number(area.dataset.y0),
               Number(area.dataset.x1),
               ptY1,
             ]);
           }

           /* Programmatic width/height + overflow:hidden leaves stale scrollLeft; start of line looks "gone". */
           area.scrollLeft = 0;

           prevScrollH = area.scrollHeight;
           prevNewlineCount = nlc;
           prevValueLen = vlen;
        });
      }
    });

    stage.addEventListener('mousedown', (e) => {
        if(e.target === overlay || e.target === img) {
            document.querySelectorAll('.pdf-box-wrapper').forEach(w => {
              w.classList.remove('active');
              w.classList.remove('block-selected');
            });
            document.querySelectorAll('.overlay .pdf-text-input').forEach((a) => a.removeAttribute('readonly'));
            hidePdfBlockFloatMenu();
            syncPdfActiveEditInputId();
            resetAllPdfEditZoom();
        }
    });
    stage.addEventListener('touchstart', (e) => {
        if(e.target === overlay || e.target === img) {
            document.querySelectorAll('.pdf-box-wrapper').forEach(w => {
              w.classList.remove('active');
              w.classList.remove('block-selected');
            });
            document.querySelectorAll('.overlay .pdf-text-input').forEach((a) => a.removeAttribute('readonly'));
            hidePdfBlockFloatMenu();
            syncPdfActiveEditInputId();
            resetAllPdfEditZoom();
        }
    }, {passive: true});
  }
}

  const handleAddText = () => {
    const firstOverlay = document.querySelector('.overlay');
    if (!firstOverlay) return;
    
    const stage = firstOverlay.closest('.page-stage');
    const pageObj = pagesMeta?.find((p) => String(p.page) === firstOverlay.dataset.page) || null;
    const pageWidth = pageObj ? pageObj.width : stage.offsetWidth;
    const pageHeight = pageObj ? pageObj.height : stage.offsetHeight;
    
    // Default size and coords for a new block (e.g., 20% down, 20% right, 150pt wide)
    const pw = 150;
    const ph = 30;
    const px = pageWidth * 0.2;
    const py = pageHeight * 0.2;
    const bbox = [px, py, px + pw, py + ph];
    
    const newItem = {
      id: "inserted_text_" + Date.now(),
      page: parseInt(firstOverlay.dataset.page || "1"),
      text: "",
      bbox: bbox,
      font: "helv",
      size: 16.0
    };
    
    const left = (px / pageWidth) * 100;
    const top = (py / pageHeight) * 100;
    const width = (pw / pageWidth) * 100; 
    const height = (ph / pageHeight) * 100;

    const boxWrapper = document.createElement("div");
    boxWrapper.className = "pdf-box-wrapper active";
    boxWrapper.style.left = `${left}%`;
    boxWrapper.style.top = `${top}%`;
    boxWrapper.style.width = `${width}%`;
    boxWrapper.style.height = `${height}%`;

    const area = document.createElement("textarea");
    area.className = "pdf-text-input";
    area.value = "";
    area.dataset.originalText = "";
    area.dataset.id = newItem.id;
    area.dataset.page = String(newItem.page);
    area.dataset.originalBbox = JSON.stringify(bbox);
    area.dataset.bbox = JSON.stringify(bbox);
    area.dataset.x0 = String(bbox[0]);
    area.dataset.y0 = String(bbox[1]);
    area.dataset.x1 = String(bbox[2]);
    area.dataset.y1 = String(bbox[3]);
    area.dataset.font = newItem.font;
    area.dataset.size = String(newItem.size);
    area.dataset.color = "#000000";
    area.dataset.align = "left";
    area.style.color = "#000000";
    area.style.fontFamily = 'sans-serif';
    area.style.lineHeight = String(PDF_OVERLAY_LINE_HEIGHT);

    boxWrapper.appendChild(area);

    const corners = ['top-left', 'top-right', 'bottom-left', 'bottom-right', 'move'];
    corners.forEach(corner => {
      const handle = document.createElement('div');
      if (corner === 'move') {
          handle.className = `move-handle`;
          handle.innerHTML = '<i class="fa-solid fa-up-down-left-right"></i>';
      } else {
          handle.className = `resize-handle rh-${corner}`;
      }
      boxWrapper.appendChild(handle);
      
      let isDragging = false;
      let startX, startY, startW, startH, startL, startT;

      const startDrag = (e) => {
          isDragging = true;
          e.stopPropagation();
          e.preventDefault();
          const evt = e.touches ? e.touches[0] : e;
          startX = evt.clientX;
          startY = evt.clientY;
          startW = boxWrapper.offsetWidth;
          startH = boxWrapper.offsetHeight;
          startL = boxWrapper.offsetLeft;
          startT = boxWrapper.offsetTop;
          
          let currentScale = 1;
          const match = stage.style.transform.match(/scale\(([^)]+)\)/);
          if (match) currentScale = parseFloat(match[1]);

          const onDrag = (e) => {
              if (!isDragging) return;
              e.preventDefault();
              const evt = e.touches ? e.touches[0] : e;
              const dx = (evt.clientX - startX) / currentScale;
              const dy = (evt.clientY - startY) / currentScale;
              let newW = startW, newH = startH, newL = startL, newT = startT;
              
              if (corner === 'move') {
                  newL = startL + dx;
                  newT = startT + dy;
              } else {
                  if (corner.includes('right')) newW = Math.max(10, startW + dx);
                  if (corner.includes('bottom')) newH = Math.max(10, startH + dy);
                  if (corner.includes('left')) { newW = Math.max(10, startW - dx); newL = startL + dx; }
                  if (corner.includes('top')) { newH = Math.max(10, startH - dy); newT = startT + dy; }
              }
              
              boxWrapper.style.width = `${newW}px`;
              boxWrapper.style.height = `${newH}px`;
              boxWrapper.style.left = `${newL}px`;
              boxWrapper.style.top = `${newT}px`;

              const stageRectInner = stage.getBoundingClientRect();
              const pxPerPoint = (stageRectInner.height / currentScale) / pageHeight; 
              const ptX0 = newL / pxPerPoint;
              const ptY0 = newT / pxPerPoint;
              const ptX1 = (newL + newW) / pxPerPoint;
              const ptY1 = (newT + newH) / pxPerPoint;
              
              area.dataset.bbox = JSON.stringify([ptX0, ptY0, ptX1, ptY1]);
          };

          const stopDrag = () => {
              isDragging = false;
              document.removeEventListener('mousemove', onDrag);
              document.removeEventListener('touchmove', onDrag);
              document.removeEventListener('mouseup', stopDrag);
              document.removeEventListener('touchend', stopDrag);
          };

          document.addEventListener('mousemove', onDrag, {passive: false});
          document.addEventListener('touchmove', onDrag, {passive: false});
          document.addEventListener('mouseup', stopDrag);
          document.addEventListener('touchend', stopDrag);
      };

      handle.addEventListener('mousedown', startDrag);
      handle.addEventListener('touchstart', startDrag, {passive: false});
    });

    boxWrapper.addEventListener('mousedown', (e) => {
      if (e.target.closest('.resize-handle')) return;
      if (e.target.closest('.move-handle')) return;
      if (boxWrapper.classList.contains('active')) return;
      e.preventDefault();
    }, true);

    boxWrapper.addEventListener('click', (e) => {
      if (e.target.closest('.resize-handle')) return;
      if (e.target.closest('.move-handle')) return;
      if (boxWrapper.classList.contains('active')) return;
      e.stopPropagation();
      selectPdfBlock(boxWrapper);
    });

    area.addEventListener("focus", () => {
      lastActiveArea = area;
      document.querySelectorAll('.pdf-box-wrapper').forEach(w => w.classList.remove('active'));
      boxWrapper.classList.remove('block-selected');
      boxWrapper.classList.add('active');
      area.removeAttribute('readonly');
      syncFormattingToolbarFromArea(area);
      focusPdfEditBoxZoomOrCenter(boxWrapper);
      syncPdfActiveEditInputId();
    });
    
    area.addEventListener("blur", (e) => {
      if (isFormatting || (e.relatedTarget && e.relatedTarget.closest('.editor-bottom-wrap'))) {
          setTimeout(() => { area.focus(); isFormatting = false; }, 50);
          return;
      }
        
      setTimeout(() => {
         if(document.activeElement !== area) {
             boxWrapper.classList.remove('active');
             syncPdfActiveEditInputId();
         }
      }, 150);
    });

    firstOverlay.appendChild(boxWrapper);
    syncPdfActiveEditInputId();
    
    // Register for saving
    textAreas.push(area);
    
    // Focus automatically
    area.focus();
    lastActiveArea = area;
    setTimeout(() => pushEditorHistory(), 120);
  };

  document.getElementById("addTextBtn")?.addEventListener("click", handleAddText);
  document.getElementById("addTextBtnBottom")?.addEventListener("click", handleAddText);

  // Register formatting click interception to bypass blur bugs
  const flagFormat = () => isFormatting = true;
  document.querySelector(".editor-toolbar")?.addEventListener("mousedown", flagFormat);
  document.querySelector(".editor-toolbar")?.addEventListener("touchstart", flagFormat, {passive: true});

  document.getElementById("undoBtn")?.addEventListener("click", () => {
    undoEditor();
  });
  document.getElementById("redoBtn")?.addEventListener("click", () => {
    redoEditor();
  });

  // Formatting Toolbar Listeners
  let activeAlignmentIndex = 0;
  const alignments = ['left', 'center', 'right'];
  const alignIcons = ['fa-align-left', 'fa-align-center', 'fa-align-right'];

  document.getElementById("formatSizeBtn")?.addEventListener("click", () => {
      if (!lastActiveArea) return;
      lastActiveArea.dataset.wasFormatted = '1';
      const currentSize = parseFloat(lastActiveArea.dataset.size) || 11;
      const sizes = [11, 14, 18, 24, 32];
      let nextIndex = sizes.indexOf(currentSize) + 1;
      if (nextIndex >= sizes.length || nextIndex === 0) nextIndex = 0;
      
      const newSize = sizes[nextIndex];
      lastActiveArea.dataset.size = newSize;
      
      const currentPx = parseFloat(lastActiveArea.style.fontSize) || 11;
      const ratio = newSize / currentSize;
      lastActiveArea.style.fontSize = `${currentPx * ratio}px`;
      lastActiveArea.style.lineHeight = `${(currentPx * ratio) * PDF_OVERLAY_LINE_HEIGHT}px`;
  });

  document.getElementById("formatBoldBtn")?.addEventListener("click", () => {
      if (!lastActiveArea) return;
      lastActiveArea.dataset.wasFormatted = '1';
      const isBold = lastActiveArea.style.fontWeight === "bold";
      lastActiveArea.style.fontWeight = isBold ? "normal" : "bold";
      
      let f = (lastActiveArea.dataset.font || "helv").toLowerCase();
      if (!isBold) {
          if (!f.includes("bold")) f += "-bold";
      } else {
          f = f.replace("-bold", "");
      }
      lastActiveArea.dataset.font = f;
      syncFormattingToolbarFromArea(lastActiveArea);
  });

  document.getElementById("formatItalicBtn")?.addEventListener("click", () => {
      if (!lastActiveArea) return;
      lastActiveArea.dataset.wasFormatted = '1';
      const isItalic = lastActiveArea.style.fontStyle === "italic";
      lastActiveArea.style.fontStyle = isItalic ? "normal" : "italic";
      
      let f = (lastActiveArea.dataset.font || "helv").toLowerCase();
      if (!isItalic) {
          if (!f.includes("italic")) f += "-italic";
      } else {
          f = f.replace("-italic", "");
      }
      lastActiveArea.dataset.font = f;
      syncFormattingToolbarFromArea(lastActiveArea);
  });

  document.getElementById("formatUnderlineBtn")?.addEventListener("click", () => {
      if (!lastActiveArea) return;
      lastActiveArea.dataset.wasFormatted = '1';
      const isUnderlined = lastActiveArea.style.textDecoration === "underline";
      lastActiveArea.style.textDecoration = isUnderlined ? "none" : "underline";
      syncFormattingToolbarFromArea(lastActiveArea);
  });

  document.getElementById("formatStrikeBtn")?.addEventListener("click", () => {
      if (!lastActiveArea) return;
      lastActiveArea.dataset.wasFormatted = '1';
      const isStruck = lastActiveArea.style.textDecoration === "line-through";
      lastActiveArea.style.textDecoration = isStruck ? "none" : "line-through";
      syncFormattingToolbarFromArea(lastActiveArea);
  });

  document.getElementById("formatFontBtn")?.addEventListener("click", () => {
      if (!lastActiveArea) return;
      lastActiveArea.dataset.wasFormatted = '1';
      let f = (lastActiveArea.dataset.font || "helv").toLowerCase();
      let base = "helv";
      if (f.includes("helv")) base = "times";
      else if (f.includes("times")) base = "courier";
      
      if (f.includes("bold")) base += "-bold";
      if (f.includes("italic")) base += "-italic";
      
      lastActiveArea.dataset.font = base;
      if (base.includes("times")) lastActiveArea.style.fontFamily = '"Times New Roman", Times, serif';
      else if (base.includes("courier")) lastActiveArea.style.fontFamily = '"Courier New", Courier, monospace';
      else lastActiveArea.style.fontFamily = 'sans-serif';
  });

  const formatColorPicker = document.getElementById("formatColorPicker");
  const formatColorIcon = document.getElementById("formatColorIcon");
  
  formatColorPicker?.addEventListener("input", (e) => {
      const hex = e.target.value;
      if (formatColorIcon) formatColorIcon.style.color = hex;
      if (!lastActiveArea) return;
      lastActiveArea.dataset.wasFormatted = '1';
      lastActiveArea.style.color = hex;
      lastActiveArea.style.setProperty("color", hex, "important");
      lastActiveArea.dataset.color = hex;
  });

  document.getElementById("formatAlignBtn")?.addEventListener("click", () => {
      if (!lastActiveArea) return;
      lastActiveArea.dataset.wasFormatted = '1';
      activeAlignmentIndex = (activeAlignmentIndex + 1) % alignments.length;
      const align = alignments[activeAlignmentIndex];
      const icon = alignIcons[activeAlignmentIndex];
      
      const alignIconEl = document.getElementById("formatAlignIcon");
      if (alignIconEl) alignIconEl.className = `fa-solid ${icon}`;
      
      lastActiveArea.style.textAlign = align;
      lastActiveArea.dataset.align = align;
  });

  document.getElementById("closeEditorBtn")?.addEventListener("click", () => {
    hidePdfBlockFloatMenu();
    document.querySelectorAll(".pdf-box-wrapper").forEach((w) => {
      w.classList.remove("block-selected");
      w.classList.remove("active");
    });
    document.querySelectorAll(".pdf-text-input").forEach((a) => a.removeAttribute("readonly"));
    syncPdfActiveEditInputId();
    resetAllPdfEditZoom();
    editorView.classList.remove("active");
    homeView.classList.add("active");
  });

/** Reset mobile zoom so layout offsets match PDF math; must run before reading bboxes for save. */
function normalizeEditorLayoutForSave() {
  document.querySelectorAll(".page-stage").forEach((stage) => {
    stage.style.transform = "scale(1)";
    stage.style.transformOrigin = "0 0";
    stage.style.width = "100%";
    stage.style.height = "auto";
    const pw = stage.closest(".page-wrapper");
    if (pw) {
      pw.style.width = "100%";
      pw.style.height = "auto";
      pw.style.zIndex = "1";
    }
  });
  document.querySelectorAll(".pdf-box-wrapper").forEach((w) => {
    w.classList.remove("active");
    w.classList.remove("block-selected");
  });
  document.querySelectorAll(".pdf-text-input").forEach((a) => a.removeAttribute("readonly"));
  hidePdfBlockFloatMenu();
  syncPdfActiveEditInputId();
  const ae = document.activeElement;
  if (ae && ae.classList && ae.classList.contains("pdf-text-input")) {
    ae.blur();
  }
}

/** Recompute dataset bbox from box positions (PDF points) after zoom is 1 — fixes mobile edit drift. */
function syncBboxesFromDOMToPdfPoints() {
  const pagesByNum = new Map(pagesMeta.map((p) => [p.page, p]));
  document.querySelectorAll(".page-wrapper").forEach((pageWrapper) => {
    const stage = pageWrapper.querySelector(".page-stage");
    const overlay = pageWrapper.querySelector(".overlay");
    if (!stage || !overlay) return;
    const pageNum = parseInt(overlay.dataset.page || "1", 10);
    const pageMeta = pagesByNum.get(pageNum);
    if (!pageMeta) return;
    const ph = pageMeta.height;
    const stageH = stage.offsetHeight;
    if (!stageH) return;
    const pxPerPoint = stageH / ph;
    overlay.querySelectorAll(".pdf-text-input").forEach((area) => {
      const box = area.closest(".pdf-box-wrapper");
      if (!box) return;
      const newL = box.offsetLeft;
      const newT = box.offsetTop;
      const newW = box.offsetWidth;
      const newH = box.offsetHeight;
      const ptX0 = newL / pxPerPoint;
      const ptY0 = newT / pxPerPoint;
      const ptX1 = (newL + newW) / pxPerPoint;
      const ptY1 = (newT + newH) / pxPerPoint;
      area.dataset.bbox = JSON.stringify([ptX0, ptY0, ptX1, ptY1]);
      area.dataset.x0 = String(ptX0);
      area.dataset.y0 = String(ptY0);
      area.dataset.x1 = String(ptX1);
      area.dataset.y1 = String(ptY1);
    });
  });
}

function isTextareaVisuallyBold(area) {
  const w = String(area.style.fontWeight || "").toLowerCase();
  if (w === "bold" || w === "bolder") return true;
  let n = parseInt(w, 10);
  if (!Number.isNaN(n) && n >= 600) return true;
  try {
    const cw = String(getComputedStyle(area).fontWeight || "").toLowerCase();
    if (cw === "bold" || cw === "bolder") return true;
    n = parseInt(cw, 10);
    if (!Number.isNaN(n) && n >= 600) return true;
  } catch (_) {}
  return false;
}

/** PDF save: bold face only when the textarea is actually bold; else strip "bold" from font names. */
function fontPayloadForPdfSave(area) {
  const raw = (area.dataset.font || "helv").trim() || "helv";
  if (isTextareaVisuallyBold(area)) {
    return raw;
  }
  if (!raw.toLowerCase().includes("bold")) {
    return raw;
  }
  let f = raw.replace(/bold/gi, "");
  f = f.replace(/[-_]{2,}/g, (m) => m[0]).replace(/^[-_]+|[-_]+$/g, "");
  return f || "helv";
}

// Save changes to backend
saveBtn.addEventListener("click", async () => {
  if (!currentFileId) return;

  normalizeEditorLayoutForSave();
  await new Promise((resolve) =>
    requestAnimationFrame(() => requestAnimationFrame(resolve))
  );
  syncBboxesFromDOMToPdfPoints();

  const normalize = (s) => String(s || "").replace(/\r\n/g, "\n");

  const edits = textAreas
    .filter((area) => {
      if (!area) return false;
      // Include if text changed OR if any formatting was applied
      const original = normalize(area.dataset.originalText).trim();
      const current = normalize(area.value).trim();
      return current !== original || area.dataset.wasFormatted === '1';
    })
    .map((area) => ({
      id: area.dataset.id,
      text: area.value,
      page: Number(area.dataset.page),
      bbox: JSON.parse(area.dataset.bbox || "[]"),
      original_bbox: JSON.parse(area.dataset.originalBbox || area.dataset.bbox || "[]"),
      font: fontPayloadForPdfSave(area),
      size: parseFloat(area.dataset.size || 11.0),
      color: area.dataset.color || "#000000",
      align: area.dataset.align || "left",
      is_underline: (area.style.textDecoration || "").includes("underline"),
      is_strike: (area.style.textDecoration || "").includes("line-through")
    }));

  loadingOverlay.classList.remove("hidden");
  uploadStatus.textContent = `Applying edits (Detected ${edits.length} changes)...`;

  try {
    const res = await fetch("/edit", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ file_id: currentFileId, edits }),
    });

    if (!res.ok) {
      const err = await res.json().catch(() => ({}));
      throw new Error(err.detail || "Server failed to save edits.");
    }

    const { download_url } = await res.json();
    
    // Setup and show preview view instead of downloading immediately
    const previewView = document.getElementById("preview-view");
    const previewContainer = document.getElementById("preview-container");
    
    // Fill container with images of edited pages
    previewContainer.innerHTML = "";
    pagesMeta.forEach(p => {
        const img = document.createElement("img");
        img.src = `/preview_edited/${currentFileId}/${p.page}?_t=${Date.now()}`;
        img.style.width = "100%";
        img.style.marginBottom = "16px";
        img.style.boxShadow = "0 2px 6px rgba(0,0,0,0.15)";
        img.style.backgroundColor = "white";
        previewContainer.appendChild(img);
    });
    
    editorView.classList.remove("active");
    previewView.classList.add("active");

    // Setup download button in preview
    const finalDlBtns = [document.getElementById("final-download-btn"), document.getElementById("final-download-btn-bottom")];
    finalDlBtns.forEach(btn => {
      // Clean up old listeners
      const newBtn = btn.cloneNode(true);
      btn.replaceWith(newBtn);
      
      newBtn.addEventListener("click", () => {
        const link = document.createElement("a");
        const fname = getDownloadFilenameForFileId(currentFileId);
        link.href = `${download_url}?filename=${encodeURIComponent(fname)}`;
        link.download = fname;
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
      });
    });

    // Setup back button
    const backBtn = document.getElementById("back-to-edit-btn");
    const newBack = backBtn.cloneNode(true);
    backBtn.replaceWith(newBack);
    newBack.addEventListener("click", () => {
      previewView.classList.remove("active");
      editorView.classList.add("active");
    });

  } catch (err) {
    alert("Error saving: " + err.message);
  } finally {
    loadingOverlay.classList.add("hidden");
  }
});

// --- More menu: PDF tools (split, rotate, crop, watermark, metadata, exports) ---
const pdfToolModal = document.getElementById("pdfToolModal");
const pdfToolModalTitle = document.getElementById("pdfToolModalTitle");
const pdfToolModalBody = document.getElementById("pdfToolModalBody");
const pdfToolModalCancel = document.getElementById("pdfToolModalCancel");
const pdfToolModalApply = document.getElementById("pdfToolModalApply");

function closePdfToolModal() {
  pdfToolModal?.classList.add("hidden");
  if (pdfToolModalApply) pdfToolModalApply.onclick = null;
}

function parsePageSpec(str, maxPage) {
  const out = new Set();
  const parts = String(str || "")
    .split(/[,\s]+/)
    .map((p) => p.trim())
    .filter(Boolean);
  for (const part of parts) {
    if (part.includes("-")) {
      const [a, b] = part.split("-").map((x) => parseInt(x.trim(), 10));
      if (Number.isNaN(a) || Number.isNaN(b)) continue;
      const lo = Math.min(a, b);
      const hi = Math.max(a, b);
      for (let i = lo; i <= hi; i++) if (i >= 1 && i <= maxPage) out.add(i);
    } else {
      const n = parseInt(part, 10);
      if (!Number.isNaN(n) && n >= 1 && n <= maxPage) out.add(n);
    }
  }
  return [...out].sort((a, b) => a - b);
}

function finishPdfToolOperation(json, _titleSuffix) {
  const oldId = selectedMoreFile?.fileObj?.id;
  const idx = mockFiles.findIndex((f) => f.id === oldId);
  if (idx > -1 && json && json.file_id) {
    const keptTitle = mockFiles[idx].title;
    mockFiles[idx].id = json.file_id;
    mockFiles[idx].size = json.size ?? mockFiles[idx].size;
    mockFiles[idx].thumb = `/preview/${json.file_id}/1?v=${Date.now()}`;
    mockFiles[idx].title = keptTitle;
    if (selectedMoreFile && selectedMoreFile.fileObj) selectedMoreFile.fileObj = mockFiles[idx];
  }
  renderMockFiles();
}

function requireMoreMenuFile() {
  const id = selectedMoreFile?.fileObj?.id;
  if (!id) {
    showCustomAlert("No file", "Upload a real PDF first.", false);
    return null;
  }
  return id;
}

function openPdfToolModal(title, html, onApply) {
  if (!pdfToolModal || !pdfToolModalTitle || !pdfToolModalBody) return;
  pdfToolModalTitle.textContent = title;
  pdfToolModalBody.innerHTML = html;
  pdfToolModal.classList.remove("hidden");
  pdfToolModalApply.onclick = async () => {
    if (typeof onApply === "function") await onApply();
  };
}

pdfToolModalCancel?.addEventListener("click", closePdfToolModal);
pdfToolModal?.addEventListener("click", (e) => {
  if (e.target === pdfToolModal) closePdfToolModal();
});

document.getElementById("splitPdfBtn")?.addEventListener("click", async () => {
  moreOptionsSheetOverlay?.classList.add("hidden");
  const fid = requireMoreMenuFile();
  if (!fid) return;
  let maxP = 1;
  try {
    const ar = await fetch(`/analyze/${fid}`);
    if (!ar.ok) throw new Error("Could not read PDF");
    const j = await ar.json();
    maxP = (j.pages && j.pages.length) || 1;
  } catch (e) {
    showCustomAlert("Error", e.message || String(e), false);
    return;
  }
  openPdfToolModal(
    "Split PDF",
    `<p style="font-size:13px;color:#555;margin:0 0 10px;">Enter pages to keep in the new file (e.g. <code>1, 3-5</code>). Max page: ${maxP}.</p>
     <input type="text" id="splitPageSpec" class="prompt-input" style="width:100%;box-sizing:border-box;" placeholder="1, 3-5" autocomplete="off" />`,
    async () => {
      const spec = document.getElementById("splitPageSpec")?.value || "";
      const pages = parsePageSpec(spec, maxP);
      if (!pages.length) {
        showCustomAlert("Invalid", "Enter at least one valid page number.", false);
        return;
      }
      loadingOverlay.classList.remove("hidden");
      uploadStatus.textContent = "Splitting PDF…";
      try {
        const res = await fetch("/split_pages", {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({ file_id: fid, page_indices: pages }),
        });
        if (!res.ok) {
          const err = await res.json().catch(() => ({}));
          throw new Error(err.detail || "Split failed");
        }
        const json = await res.json();
        finishPdfToolOperation(json, "_split");
        closePdfToolModal();
        showCustomAlert("Done", "New PDF created from selected pages.", true);
      } catch (e) {
        showCustomAlert("Failed", e.message || String(e), false);
      } finally {
        loadingOverlay.classList.add("hidden");
      }
    }
  );
});

document.getElementById("rotatePagesBtn")?.addEventListener("click", async () => {
  moreOptionsSheetOverlay?.classList.add("hidden");
  const fid = requireMoreMenuFile();
  if (!fid) return;
  let maxP = 1;
  try {
    const ar = await fetch(`/analyze/${fid}`);
    const j = await ar.json();
    maxP = (j.pages && j.pages.length) || 1;
  } catch {
    showCustomAlert("Error", "Could not read PDF.", false);
    return;
  }
  openPdfToolModal(
    "Rotate pages",
    `<p style="font-size:13px;color:#555;margin:0 0 10px;">Leave <b>Pages</b> empty to rotate all pages. Example: <code>1,2</code> Max: ${maxP}.</p>
     <label style="display:block;font-size:13px;margin-bottom:4px;">Angle</label>
     <select id="rotAngle" class="prompt-input" style="width:100%;margin-bottom:12px;">
       <option value="90">90° clockwise</option>
       <option value="180">180°</option>
       <option value="270">270° clockwise (90° CCW)</option>
     </select>
     <label style="display:block;font-size:13px;margin-bottom:4px;">Pages (optional)</label>
     <input type="text" id="rotPagesSpec" class="prompt-input" style="width:100%;box-sizing:border-box;" placeholder="empty = all pages" />`,
    async () => {
      const angle = parseInt(document.getElementById("rotAngle")?.value || "90", 10);
      const spec = document.getElementById("rotPagesSpec")?.value?.trim() || "";
      let pages = null;
      if (spec) {
        pages = parsePageSpec(spec, maxP);
        if (!pages.length) {
          showCustomAlert("Invalid", "No valid page numbers.", false);
          return;
        }
      }
      loadingOverlay.classList.remove("hidden");
      uploadStatus.textContent = "Rotating…";
      try {
        const res = await fetch("/rotate_pages", {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({ file_id: fid, angle, pages }),
        });
        if (!res.ok) {
          const err = await res.json().catch(() => ({}));
          throw new Error(err.detail || "Rotate failed");
        }
        const json = await res.json();
        finishPdfToolOperation(json, "_rotated");
        closePdfToolModal();
        showCustomAlert("Done", "Pages rotated.", true);
      } catch (e) {
        showCustomAlert("Failed", e.message || String(e), false);
      } finally {
        loadingOverlay.classList.add("hidden");
      }
    }
  );
});

const cropView = document.getElementById("cropView");
const cropVisualState = {
  fileId: null,
  pageCount: 1,
  currentPage: 1,
  norm: { x0: 0.05, y0: 0.05, x1: 0.95, y1: 0.95 },
  drag: null,
};

function cropNormFromClient(clientX, clientY) {
  const el = document.getElementById("cropInteract");
  if (!el) return { x: 0, y: 0 };
  const r = el.getBoundingClientRect();
  if (r.width < 1 || r.height < 1) return { x: 0, y: 0 };
  return {
    x: Math.min(1, Math.max(0, (clientX - r.left) / r.width)),
    y: Math.min(1, Math.max(0, (clientY - r.top) / r.height)),
  };
}

function renderCropVisualBox() {
  const b = document.getElementById("cropBox");
  if (!b) return;
  const { norm } = cropVisualState;
  b.style.left = `${norm.x0 * 100}%`;
  b.style.top = `${norm.y0 * 100}%`;
  b.style.width = `${(norm.x1 - norm.x0) * 100}%`;
  b.style.height = `${(norm.y1 - norm.y0) * 100}%`;
}

function updateCropPageNav() {
  const label = document.getElementById("cropPageLabel");
  const prev = document.getElementById("cropPrevPage");
  const next = document.getElementById("cropNextPage");
  const n = cropVisualState.pageCount;
  const p = cropVisualState.currentPage;
  if (label) label.textContent = `Page ${p} / ${n}`;
  if (prev) prev.disabled = p <= 1;
  if (next) next.disabled = p >= n;
}

async function loadCropPageImage() {
  const img = document.getElementById("cropPageImg");
  const fid = cropVisualState.fileId;
  if (!img || !fid) return;
  img.decoding = "async";
  img.src = `/preview/${fid}/${cropVisualState.currentPage}?v=${Date.now()}`;
  await new Promise((res, rej) => {
    img.onload = res;
    img.onerror = rej;
  }).catch(() => {});
  renderCropVisualBox();
}

async function openCropVisualView(fid) {
  cropVisualState.fileId = fid;
  cropVisualState.norm = { x0: 0.05, y0: 0.05, x1: 0.95, y1: 0.95 };
  cropVisualState.drag = null;
  try {
    const ar = await fetch(`/analyze/${fid}`);
    if (!ar.ok) throw new Error("Could not read PDF");
    const j = await ar.json();
    cropVisualState.pageCount = (j.pages && j.pages.length) || 1;
    cropVisualState.currentPage = 1;
  } catch (e) {
    showCustomAlert("Error", e.message || String(e), false);
    return;
  }
  moreOptionsSheetOverlay?.classList.add("hidden");
  homeView?.classList.remove("active");
  editorView?.classList.remove("active");
  combineView?.classList.remove("active");
  selectFilesView?.classList.remove("active");
  document.getElementById("reorder-view")?.classList.remove("active");
  document.getElementById("preview-view")?.classList.remove("active");
  cropView?.classList.add("active");
  updateCropPageNav();
  loadingOverlay.classList.remove("hidden");
  uploadStatus.textContent = "Loading preview…";
  try {
    await loadCropPageImage();
  } finally {
    loadingOverlay.classList.add("hidden");
  }
}

function closeCropVisualView() {
  cropView?.classList.remove("active");
  cropVisualState.drag = null;
  homeView?.classList.add("active");
}

let cropPointerBound = false;
function bindCropVisualPointers() {
  if (cropPointerBound) return;
  cropPointerBound = true;
  const MIN = 0.05;

  const onMove = (e) => {
    const d = cropVisualState.drag;
    if (!d) return;
    e.preventDefault();
    const p = cropNormFromClient(e.clientX, e.clientY);

    if (d.type === "move") {
      const r = document.getElementById("cropInteract")?.getBoundingClientRect();
      if (!r || r.width < 1) return;
      const dx = (e.clientX - d.sx) / r.width;
      const dy = (e.clientY - d.sy) / r.height;
      const w = d.startNorm.x1 - d.startNorm.x0;
      const h = d.startNorm.y1 - d.startNorm.y0;
      let mx0 = d.startNorm.x0 + dx;
      let my0 = d.startNorm.y0 + dy;
      if (mx0 < 0) mx0 = 0;
      if (my0 < 0) my0 = 0;
      if (mx0 + w > 1) mx0 = 1 - w;
      if (my0 + h > 1) my0 = 1 - h;
      cropVisualState.norm = { x0: mx0, y0: my0, x1: mx0 + w, y1: my0 + h };
    } else {
      const sn = d.startNorm;
      const x0 = sn.x0;
      const y0 = sn.y0;
      const x1 = sn.x1;
      const y1 = sn.y1;
      if (d.type === "nw") {
        cropVisualState.norm = {
          x0: Math.min(p.x, x1 - MIN),
          y0: Math.min(p.y, y1 - MIN),
          x1,
          y1,
        };
      } else if (d.type === "ne") {
        cropVisualState.norm = {
          x0,
          y0: Math.min(p.y, y1 - MIN),
          x1: Math.max(p.x, x0 + MIN),
          y1,
        };
      } else if (d.type === "sw") {
        cropVisualState.norm = {
          x0: Math.min(p.x, x1 - MIN),
          y0,
          x1,
          y1: Math.max(p.y, y0 + MIN),
        };
      } else if (d.type === "se") {
        cropVisualState.norm = {
          x0,
          y0,
          x1: Math.max(p.x, x0 + MIN),
          y1: Math.max(p.y, y0 + MIN),
        };
      }
    }
    renderCropVisualBox();
  };

  const onUp = () => {
    cropVisualState.drag = null;
  };

  document.addEventListener("pointermove", onMove, { passive: false });
  document.addEventListener("pointerup", onUp);
  document.addEventListener("pointercancel", onUp);

  const cropMove = document.getElementById("cropMove");
  cropMove?.addEventListener("pointerdown", (e) => {
    if (e.button != null && e.button !== 0) return;
    e.preventDefault();
    e.stopPropagation();
    try {
      e.currentTarget.setPointerCapture(e.pointerId);
    } catch (_) {
      /* ignore */
    }
    cropVisualState.drag = {
      type: "move",
      sx: e.clientX,
      sy: e.clientY,
      startNorm: { ...cropVisualState.norm },
    };
  });

  document.querySelectorAll("#cropBox .crop-handle").forEach((h) => {
    h.addEventListener("pointerdown", (e) => {
      e.preventDefault();
      e.stopPropagation();
      const corner = h.getAttribute("data-corner");
      try {
        h.setPointerCapture(e.pointerId);
      } catch (_) {
        /* ignore */
      }
      cropVisualState.drag = {
        type: corner,
        startNorm: { ...cropVisualState.norm },
      };
    });
  });
}

document.getElementById("cropPageBtn")?.addEventListener("click", async () => {
  const fid = requireMoreMenuFile();
  if (!fid) return;
  bindCropVisualPointers();
  await openCropVisualView(fid);
});

document.getElementById("cropBackBtn")?.addEventListener("click", () => closeCropVisualView());

document.getElementById("cropPrevPage")?.addEventListener("click", async () => {
  if (cropVisualState.currentPage <= 1) return;
  cropVisualState.currentPage -= 1;
  updateCropPageNav();
  loadingOverlay.classList.remove("hidden");
  uploadStatus.textContent = "Loading…";
  try {
    await loadCropPageImage();
  } finally {
    loadingOverlay.classList.add("hidden");
  }
});

document.getElementById("cropNextPage")?.addEventListener("click", async () => {
  if (cropVisualState.currentPage >= cropVisualState.pageCount) return;
  cropVisualState.currentPage += 1;
  updateCropPageNav();
  loadingOverlay.classList.remove("hidden");
  uploadStatus.textContent = "Loading…";
  try {
    await loadCropPageImage();
  } finally {
    loadingOverlay.classList.add("hidden");
  }
});

document.getElementById("cropApplyBtn")?.addEventListener("click", async () => {
  const fid = cropVisualState.fileId;
  if (!fid) return;
  const { norm } = cropVisualState;
  const allPages = !!document.getElementById("cropApplyAllPages")?.checked;
  loadingOverlay.classList.remove("hidden");
  uploadStatus.textContent = "Cropping…";
  try {
    const res = await fetch("/crop_page_norm", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        file_id: fid,
        x0: norm.x0,
        y0: norm.y0,
        x1: norm.x1,
        y1: norm.y1,
        page: cropVisualState.currentPage,
        all_pages: allPages,
      }),
    });
    if (!res.ok) {
      const err = await res.json().catch(() => ({}));
      throw new Error(err.detail || "Crop failed");
    }
    const json = await res.json();
    if (!selectedMoreFile?.fileObj?.id) {
      const f = mockFiles.find((x) => x.id === fid);
      if (f) selectedMoreFile = { fileObj: f, cardEl: null };
    }
    finishPdfToolOperation(json, "_cropped");
    closeCropVisualView();
    showCustomAlert("Done", "Crop applied.", true);
  } catch (e) {
    showCustomAlert("Failed", e.message || String(e), false);
  } finally {
    loadingOverlay.classList.add("hidden");
  }
});

const stampView = document.getElementById("stampView");
const stampVisualState = {
  pdfId: null,
  stampId: null,
  pageCount: 1,
  currentPage: 1,
  norm: { x0: 0.12, y0: 0.68, x1: 0.88, y1: 0.93 },
  drag: null,
};
let stampPendingPdfId = null;

function stampNormFromClient(clientX, clientY) {
  const el = document.getElementById("stampInteract");
  if (!el) return { x: 0, y: 0 };
  const r = el.getBoundingClientRect();
  if (r.width < 1 || r.height < 1) return { x: 0, y: 0 };
  return {
    x: Math.min(1, Math.max(0, (clientX - r.left) / r.width)),
    y: Math.min(1, Math.max(0, (clientY - r.top) / r.height)),
  };
}

function renderStampBox() {
  const b = document.getElementById("stampBox");
  if (!b) return;
  const { norm } = stampVisualState;
  b.style.left = `${norm.x0 * 100}%`;
  b.style.top = `${norm.y0 * 100}%`;
  b.style.width = `${(norm.x1 - norm.x0) * 100}%`;
  b.style.height = `${(norm.y1 - norm.y0) * 100}%`;
}

function updateStampPageNav() {
  const label = document.getElementById("stampPageLabel");
  const prev = document.getElementById("stampPrevPage");
  const next = document.getElementById("stampNextPage");
  const n = stampVisualState.pageCount;
  const p = stampVisualState.currentPage;
  if (label) label.textContent = `Page ${p} / ${n}`;
  if (prev) prev.disabled = p <= 1;
  if (next) next.disabled = p >= n;
}

async function loadStampPageImage() {
  const img = document.getElementById("stampPageImg");
  const fid = stampVisualState.pdfId;
  if (!img || !fid) return;
  img.decoding = "async";
  img.src = `/preview/${fid}/${stampVisualState.currentPage}?v=${Date.now()}`;
  await new Promise((res, rej) => {
    img.onload = res;
    img.onerror = rej;
  }).catch(() => {});
  renderStampBox();
}

function closeStampView() {
  stampView?.classList.remove("active");
  stampVisualState.drag = null;
  stampVisualState.stampId = null;
  stampVisualState.pdfId = null;
  homeView?.classList.add("active");
}

async function openStampViewAfterUpload(stampId, pdfId) {
  stampVisualState.stampId = stampId;
  stampVisualState.pdfId = pdfId;
  stampVisualState.norm = { x0: 0.12, y0: 0.68, x1: 0.88, y1: 0.93 };
  stampVisualState.drag = null;
  const inner = document.getElementById("stampInnerImg");
  if (inner) inner.src = `/stamp_preview/${stampId}?v=${Date.now()}`;
  try {
    const ar = await fetch(`/analyze/${pdfId}`);
    if (!ar.ok) throw new Error("Could not read PDF");
    const j = await ar.json();
    stampVisualState.pageCount = (j.pages && j.pages.length) || 1;
    stampVisualState.currentPage = 1;
  } catch (e) {
    showCustomAlert("Error", e.message || String(e), false);
    return;
  }
  moreOptionsSheetOverlay?.classList.add("hidden");
  homeView?.classList.remove("active");
  editorView?.classList.remove("active");
  combineView?.classList.remove("active");
  selectFilesView?.classList.remove("active");
  cropView?.classList.remove("active");
  document.getElementById("reorder-view")?.classList.remove("active");
  document.getElementById("preview-view")?.classList.remove("active");
  stampView?.classList.add("active");
  updateStampPageNav();
  loadingOverlay.classList.remove("hidden");
  uploadStatus.textContent = "Loading…";
  try {
    await loadStampPageImage();
  } finally {
    loadingOverlay.classList.add("hidden");
  }
}

let stampPointerBound = false;
function bindStampPointers() {
  if (stampPointerBound) return;
  stampPointerBound = true;
  const MIN = 0.05;

  const onMove = (e) => {
    const d = stampVisualState.drag;
    if (!d) return;
    e.preventDefault();
    const p = stampNormFromClient(e.clientX, e.clientY);

    if (d.type === "move") {
      const r = document.getElementById("stampInteract")?.getBoundingClientRect();
      if (!r || r.width < 1) return;
      const dx = (e.clientX - d.sx) / r.width;
      const dy = (e.clientY - d.sy) / r.height;
      const w = d.startNorm.x1 - d.startNorm.x0;
      const h = d.startNorm.y1 - d.startNorm.y0;
      let mx0 = d.startNorm.x0 + dx;
      let my0 = d.startNorm.y0 + dy;
      if (mx0 < 0) mx0 = 0;
      if (my0 < 0) my0 = 0;
      if (mx0 + w > 1) mx0 = 1 - w;
      if (my0 + h > 1) my0 = 1 - h;
      stampVisualState.norm = { x0: mx0, y0: my0, x1: mx0 + w, y1: my0 + h };
    } else {
      const sn = d.startNorm;
      const x0 = sn.x0;
      const y0 = sn.y0;
      const x1 = sn.x1;
      const y1 = sn.y1;
      if (d.type === "nw") {
        stampVisualState.norm = {
          x0: Math.min(p.x, x1 - MIN),
          y0: Math.min(p.y, y1 - MIN),
          x1,
          y1,
        };
      } else if (d.type === "ne") {
        stampVisualState.norm = {
          x0,
          y0: Math.min(p.y, y1 - MIN),
          x1: Math.max(p.x, x0 + MIN),
          y1,
        };
      } else if (d.type === "sw") {
        stampVisualState.norm = {
          x0: Math.min(p.x, x1 - MIN),
          y0,
          x1,
          y1: Math.max(p.y, y0 + MIN),
        };
      } else if (d.type === "se") {
        stampVisualState.norm = {
          x0,
          y0,
          x1: Math.max(p.x, x0 + MIN),
          y1: Math.max(p.y, y0 + MIN),
        };
      }
    }
    renderStampBox();
  };

  const onUp = () => {
    stampVisualState.drag = null;
  };

  document.addEventListener("pointermove", onMove, { passive: false });
  document.addEventListener("pointerup", onUp);
  document.addEventListener("pointercancel", onUp);

  document.getElementById("stampMove")?.addEventListener("pointerdown", (e) => {
    if (e.button != null && e.button !== 0) return;
    e.preventDefault();
    e.stopPropagation();
    try {
      e.currentTarget.setPointerCapture(e.pointerId);
    } catch (_) {
      /* ignore */
    }
    stampVisualState.drag = {
      type: "move",
      sx: e.clientX,
      sy: e.clientY,
      startNorm: { ...stampVisualState.norm },
    };
  });

  document.querySelectorAll("#stampBox .stamp-handle").forEach((h) => {
    h.addEventListener("pointerdown", (e) => {
      e.preventDefault();
      e.stopPropagation();
      const corner = h.getAttribute("data-corner");
      try {
        h.setPointerCapture(e.pointerId);
      } catch (_) {
        /* ignore */
      }
      stampVisualState.drag = {
        type: corner,
        startNorm: { ...stampVisualState.norm },
      };
    });
  });
}

document.getElementById("signatureStampBtn")?.addEventListener("click", () => {
  moreOptionsSheetOverlay?.classList.add("hidden");
  const fid = requireMoreMenuFile();
  if (!fid) return;
  stampPendingPdfId = fid;
  bindStampPointers();
  document.getElementById("stampFileInput")?.click();
});

document.getElementById("stampFileInput")?.addEventListener("change", async (e) => {
  const file = e.target.files && e.target.files[0];
  e.target.value = "";
  if (!file || !stampPendingPdfId) return;
  const fd = new FormData();
  fd.append("file", file);
  const fastSig = document.getElementById("stampFastMode")?.checked;
  fd.append("fast", fastSig ? "true" : "false");
  loadingOverlay.classList.remove("hidden");
  uploadStatus.textContent = fastSig ? "Processing signature (fast)…" : "Removing background…";
  try {
    const res = await fetch("/upload_stamp", { method: "POST", body: fd });
    if (!res.ok) {
      const err = await res.json().catch(() => ({}));
      throw new Error(err.detail || "Upload failed");
    }
    const j = await res.json();
    await openStampViewAfterUpload(j.stamp_id, stampPendingPdfId);
  } catch (err) {
    showCustomAlert("Failed", err.message || String(err), false);
  } finally {
    loadingOverlay.classList.add("hidden");
  }
  stampPendingPdfId = null;
});

document.getElementById("stampBackBtn")?.addEventListener("click", () => closeStampView());

document.getElementById("stampPrevPage")?.addEventListener("click", async () => {
  if (stampVisualState.currentPage <= 1) return;
  stampVisualState.currentPage -= 1;
  updateStampPageNav();
  loadingOverlay.classList.remove("hidden");
  uploadStatus.textContent = "Loading…";
  try {
    await loadStampPageImage();
  } finally {
    loadingOverlay.classList.add("hidden");
  }
});

document.getElementById("stampNextPage")?.addEventListener("click", async () => {
  if (stampVisualState.currentPage >= stampVisualState.pageCount) return;
  stampVisualState.currentPage += 1;
  updateStampPageNav();
  loadingOverlay.classList.remove("hidden");
  uploadStatus.textContent = "Loading…";
  try {
    await loadStampPageImage();
  } finally {
    loadingOverlay.classList.add("hidden");
  }
});

document.getElementById("stampApplyBtn")?.addEventListener("click", async () => {
  const { pdfId, stampId, norm, currentPage } = stampVisualState;
  if (!pdfId || !stampId) return;
  loadingOverlay.classList.remove("hidden");
  uploadStatus.textContent = "Placing signature…";
  try {
    const res = await fetch("/apply_stamp", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        file_id: pdfId,
        stamp_id: stampId,
        page: currentPage,
        x0: norm.x0,
        y0: norm.y0,
        x1: norm.x1,
        y1: norm.y1,
      }),
    });
    if (!res.ok) {
      const err = await res.json().catch(() => ({}));
      throw new Error(err.detail || "Could not apply stamp");
    }
    const json = await res.json();
    if (!selectedMoreFile?.fileObj?.id) {
      const f = mockFiles.find((x) => x.id === pdfId);
      if (f) selectedMoreFile = { fileObj: f, cardEl: null };
    }
    finishPdfToolOperation(json, "_signed");
    closeStampView();
    showCustomAlert("Done", "Signature placed on the PDF.", true);
  } catch (err) {
    showCustomAlert("Failed", err.message || String(err), false);
  } finally {
    loadingOverlay.classList.add("hidden");
  }
});

document.getElementById("watermarkPdfBtn")?.addEventListener("click", () => {
  moreOptionsSheetOverlay?.classList.add("hidden");
  const fid = requireMoreMenuFile();
  if (!fid) return;
  openPdfToolModal(
    "Watermark",
    `<p style="font-size:13px;color:#555;margin:0 0 10px;line-height:1.45;">Text on <strong>every page</strong>. Har style alag need ke liye hai — zyada protection, classic look, halka mark, ya sirf ek jagah label.</p>
     <label style="font-size:13px;">Text</label>
     <input type="text" id="wmText" class="prompt-input" style="width:100%;margin-bottom:10px;box-sizing:border-box;" value="Draft" maxlength="100" />
     <label style="font-size:13px;">Size hint (edge/corner modes = chhota text; single spot = max ~30pt)</label>
     <input type="number" id="wmSize" class="prompt-input" style="width:100%;margin-bottom:10px;box-sizing:border-box;" min="8" max="200" step="1" value="48" />
     <label style="font-size:13px;">Position — choose by your use case</label>
     <select id="wmPos" class="prompt-input" style="width:100%;margin-bottom:10px;box-sizing:border-box;font-size:14px;">
       <optgroup label="Strong coverage (offices, legal, confidential drafts)">
         <option value="perimeter" selected>Repeat on all edges — small text, many copies; hardest to ignore or crop out</option>
       </optgroup>
       <optgroup label="Bold & obvious (classic “DRAFT” look)">
         <option value="diagonal">Center diagonal — one large tilted line through the middle</option>
       </optgroup>
       <optgroup label="Light & readable (students, sharing, less clutter)">
         <option value="four_corners">Four corners only — small label each corner; content stays clear</option>
       </optgroup>
       <optgroup label="Single spot (logo line, name, one stamp)">
         <option value="center">Center</option>
         <option value="top_center">Top center</option>
         <option value="bottom_center">Bottom center</option>
         <option value="top_left">Top left</option>
         <option value="top_right">Top right</option>
         <option value="middle_left">Middle left</option>
         <option value="middle_right">Middle right</option>
         <option value="bottom_left">Bottom left</option>
         <option value="bottom_right">Bottom right</option>
       </optgroup>
     </select>
     <label style="font-size:13px;">Opacity (0.05–0.95, higher = darker)</label>
     <input type="number" id="wmOp" class="prompt-input" style="width:100%" min="0.05" max="0.95" step="0.05" value="0.25" />`,
    async () => {
      const text = document.getElementById("wmText")?.value?.trim() || "Draft";
      let fontSize = parseFloat(document.getElementById("wmSize")?.value || "48");
      if (Number.isNaN(fontSize)) fontSize = 48;
      fontSize = Math.min(200, Math.max(8, fontSize));
      const position = document.getElementById("wmPos")?.value || "center";
      let opacity = parseFloat(document.getElementById("wmOp")?.value || "0.25");
      if (Number.isNaN(opacity)) opacity = 0.25;
      opacity = Math.min(0.95, Math.max(0.05, opacity));
      loadingOverlay.classList.remove("hidden");
      uploadStatus.textContent = "Adding watermark…";
      try {
        const res = await fetch("/watermark", {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({ file_id: fid, text, opacity, font_size: fontSize, position }),
        });
        if (!res.ok) {
          const err = await res.json().catch(() => ({}));
          throw new Error(err.detail || "Watermark failed");
        }
        const json = await res.json();
        finishPdfToolOperation(json, "_watermarked");
        closePdfToolModal();
        showCustomAlert("Done", "Watermark added.", true);
      } catch (e) {
        showCustomAlert("Failed", e.message || String(e), false);
      } finally {
        loadingOverlay.classList.add("hidden");
      }
    }
  );
});

document.getElementById("removeWatermarkBtn")?.addEventListener("click", async () => {
  moreOptionsSheetOverlay?.classList.add("hidden");
  const fid = requireMoreMenuFile();
  if (!fid) return;
  loadingOverlay.classList.remove("hidden");
  uploadStatus.textContent = "Removing watermark…";
  try {
    const res = await fetch("/remove_watermark", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ file_id: fid }),
    });
    if (!res.ok) {
      const err = await res.json().catch(() => ({}));
      throw new Error(err.detail || "Could not remove watermark");
    }
    const json = await res.json();
    const idx = mockFiles.findIndex((f) => f.id === fid);
    if (idx > -1) {
      mockFiles[idx].size = json.size ?? mockFiles[idx].size;
      mockFiles[idx].thumb = `/preview/${fid}/1?v=${Date.now()}`;
    }
    renderMockFiles();
    showCustomAlert("Done", "Watermark removed.", true);
  } catch (e) {
    showCustomAlert("Failed", e.message || String(e), false);
  } finally {
    loadingOverlay.classList.add("hidden");
  }
});

document.getElementById("pdfMetadataBtn")?.addEventListener("click", async () => {
  moreOptionsSheetOverlay?.classList.add("hidden");
  const fid = requireMoreMenuFile();
  if (!fid) return;
  let meta = {};
  try {
    const res = await fetch(`/pdf_metadata/${fid}`);
    if (!res.ok) throw new Error("Could not read metadata");
    const j = await res.json();
    meta = j.metadata || {};
  } catch (e) {
    showCustomAlert("Error", e.message || String(e), false);
    return;
  }
  openPdfToolModal(
    "PDF metadata",
    `<p style="font-size:13px;color:#555;margin:0 0 10px;">Edit standard title/author, or strip all metadata for privacy.</p>
     <label style="font-size:13px;">Title</label>
     <input type="text" id="metaTitle" class="prompt-input" style="width:100%;margin-bottom:8px;box-sizing:border-box;" autocomplete="off" />
     <label style="font-size:13px;">Author</label>
     <input type="text" id="metaAuthor" class="prompt-input" style="width:100%;margin-bottom:12px;box-sizing:border-box;" autocomplete="off" />
     <button type="button" id="stripMetaBtn" class="prompt-btn btn-cancel" style="width:100%;margin-bottom:8px;">Strip all metadata</button>`,
    async () => {
      const title = document.getElementById("metaTitle")?.value ?? "";
      const author = document.getElementById("metaAuthor")?.value ?? "";
      loadingOverlay.classList.remove("hidden");
      uploadStatus.textContent = "Saving metadata…";
      try {
        const res = await fetch("/pdf_metadata", {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({ file_id: fid, title, author, strip: false }),
        });
        if (!res.ok) {
          const err = await res.json().catch(() => ({}));
          throw new Error(err.detail || "Save failed");
        }
        const json = await res.json();
        finishPdfToolOperation(json, "_meta");
        closePdfToolModal();
        showCustomAlert("Done", "Metadata saved.", true);
      } catch (e) {
        showCustomAlert("Failed", e.message || String(e), false);
      } finally {
        loadingOverlay.classList.add("hidden");
      }
    }
  );
  setTimeout(() => {
    const mt = document.getElementById("metaTitle");
    const ma = document.getElementById("metaAuthor");
    if (mt) mt.value = meta.title || "";
    if (ma) ma.value = meta.author || "";
  }, 0);
  document.getElementById("stripMetaBtn")?.addEventListener(
    "click",
    async () => {
      loadingOverlay.classList.remove("hidden");
      uploadStatus.textContent = "Removing metadata…";
      try {
        const res = await fetch("/pdf_metadata", {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({ file_id: fid, strip: true }),
        });
        if (!res.ok) {
          const err = await res.json().catch(() => ({}));
          throw new Error(err.detail || "Strip failed");
        }
        const json = await res.json();
        finishPdfToolOperation(json, "_stripped");
        closePdfToolModal();
        showCustomAlert("Done", "Metadata removed.", true);
      } catch (e) {
        showCustomAlert("Failed", e.message || String(e), false);
      } finally {
        loadingOverlay.classList.add("hidden");
      }
    },
    { once: true }
  );
});

document.getElementById("downloadPageImageBtn")?.addEventListener("click", async () => {
  moreOptionsSheetOverlay?.classList.add("hidden");
  const fid = requireMoreMenuFile();
  if (!fid) return;
  let maxP = 1;
  try {
    const ar = await fetch(`/analyze/${fid}`);
    const j = await ar.json();
    maxP = (j.pages && j.pages.length) || 1;
  } catch {
    showCustomAlert("Error", "Could not read PDF.", false);
    return;
  }
  openPdfToolModal(
    "Download one page as image",
    `<p style="font-size:13px;color:#555;margin:0 0 10px;">Single page only (not a ZIP). Max page: ${maxP}.</p>
     <label style="font-size:13px;">Page #</label>
     <input type="number" id="onePageImgNum" class="prompt-input" style="width:100%;margin-bottom:10px;" min="1" max="${maxP}" value="1" />
     <label style="font-size:13px;">Format</label>
     <select id="onePageImgFmt" class="prompt-input" style="width:100%">
       <option value="png">PNG</option>
       <option value="jpeg">JPEG</option>
     </select>`,
    async () => {
      const p = parseInt(document.getElementById("onePageImgNum")?.value || "1", 10);
      const fmt = document.getElementById("onePageImgFmt")?.value || "png";
      if (p < 1 || p > maxP) {
        showCustomAlert("Invalid", `Page must be 1–${maxP}.`, false);
        return;
      }
      closePdfToolModal();
      window.location.href = `/export_page_image/${fid}/${p}?format=${encodeURIComponent(fmt)}`;
    }
  );
});

document.getElementById("exportTxtBtn")?.addEventListener("click", () => {
  moreOptionsSheetOverlay?.classList.add("hidden");
  const fid = requireMoreMenuFile();
  if (!fid) return;
  window.location.href = `/export_text/${fid}`;
});

document.getElementById("exportDocxBtn")?.addEventListener("click", () => {
  moreOptionsSheetOverlay?.classList.add("hidden");
  const fid = requireMoreMenuFile();
  if (!fid) return;
  window.location.href = `/export_docx/${fid}`;
});

document.getElementById("signPkcs12Btn")?.addEventListener("click", () => {
  moreOptionsSheetOverlay?.classList.add("hidden");
  const fid = requireMoreMenuFile();
  if (!fid) return;
  openPdfToolModal(
    "Digital signature (PKCS#12)",
    `<p style="font-size:13px;color:#555;margin:0 0 10px;">Choose a PKCS#12 file (.p12 / .pfx) with your signing certificate. A new signed copy is created.</p>
     <label style="display:block;font-size:13px;margin-bottom:6px;">Archive password (if any)</label>
     <input type="password" id="p12Password" class="prompt-input" style="width:100%;box-sizing:border-box;margin-bottom:12px;" autocomplete="off" />
     <label style="display:block;font-size:13px;margin-bottom:6px;">Certificate file</label>
     <input type="file" id="p12File" accept=".p12,.pfx,application/x-pkcs12" class="prompt-input" style="width:100%;box-sizing:border-box;" />`,
    async () => {
      const pwEl = document.getElementById("p12Password");
      const fEl = document.getElementById("p12File");
      const file = fEl?.files?.[0];
      if (!file) {
        showCustomAlert("Required", "Choose a .p12 / .pfx file.", false);
        return;
      }
      const fd = new FormData();
      fd.append("file_id", fid);
      fd.append("password", pwEl?.value || "");
      fd.append("p12", file);
      loadingOverlay.classList.remove("hidden");
      uploadStatus.textContent = "Signing PDF…";
      try {
        const res = await fetch("/sign_pkcs12", { method: "POST", body: fd });
        if (!res.ok) {
          const err = await res.json().catch(() => ({}));
          throw new Error(err.detail || "Signing failed");
        }
        const json = await res.json();
        finishPdfToolOperation(json, "_signed");
        closePdfToolModal();
        showCustomAlert("Signed", "Digital signature applied. Your list now points to the signed PDF.", true);
      } catch (e) {
        showCustomAlert("Signing failed", e.message || String(e), false);
      } finally {
        loadingOverlay.classList.add("hidden");
      }
    }
  );
});
