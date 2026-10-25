/**
 * PPT Unify — Selection-Based UI Frontend Application Logic
 * Supports uploading up to 100 PPT/PPTX files error-free,
 * line-by-line queue arrangement, sequential & interleaved merge,
 * and dual export to PPTX & PDF.
 */

let fileList = []; // Array of file objects in current line order
let isUploading = false;
let isMerging = false;
let activePollInterval = null;

// DOM Elements
const dropZone = document.getElementById("dropZone");
const fileInput = document.getElementById("fileInput");
const browseBtn = document.getElementById("browseBtn");
const fileTableBody = document.getElementById("fileTableBody");
const emptyRow = document.getElementById("emptyRow");
const fileCountBadge = document.getElementById("fileCountBadge");
const totalSlidesBadge = document.getElementById("totalSlidesBadge");
const systemStateBadge = document.getElementById("systemStateBadge");
const mergeBtn = document.getElementById("mergeBtn");
const sortNaturalBtn = document.getElementById("sortNaturalBtn");
const sortAlphaBtn = document.getElementById("sortAlphaBtn");
const reverseOrderBtn = document.getElementById("reverseOrderBtn");
const clearAllBtn = document.getElementById("clearAllBtn");

const uploadProgressContainer = document.getElementById("uploadProgressContainer");
const uploadProgressBar = document.getElementById("uploadProgressBar");
const uploadProgressText = document.getElementById("uploadProgressText");
const uploadProgressPct = document.getElementById("uploadProgressPct");

const progressPanel = document.getElementById("progressPanel");
const progressTitle = document.getElementById("progressTitle");
const progressPct = document.getElementById("progressPct");
const mergeProgressBar = document.getElementById("mergeProgressBar");
const progressStatusMsg = document.getElementById("progressStatusMsg");

const downloadCard = document.getElementById("downloadCard");
const downloadFileName = document.getElementById("downloadFileName");
const downloadMeta = document.getElementById("downloadMeta");
const downloadPptxBtn = document.getElementById("downloadPptxBtn");
const downloadPdfBtn = document.getElementById("downloadPdfBtn");
const downloadZipBtn = document.getElementById("downloadZipBtn");

const themeDarkBtn = document.getElementById("themeDarkBtn");
const themeLightBtn = document.getElementById("themeLightBtn");

// Initialize Event Listeners
function init() {
  initTheme();
  browseBtn.addEventListener("click", () => fileInput.click());
  fileInput.addEventListener("change", handleFileSelect);

  // Drag & drop
  ["dragenter", "dragover"].forEach(event => {
    dropZone.addEventListener(event, (e) => {
      e.preventDefault();
      dropZone.classList.add("dragover");
    });
  });

  ["dragleave", "drop"].forEach(event => {
    dropZone.addEventListener(event, (e) => {
      e.preventDefault();
      dropZone.classList.remove("dragover");
    });
  });

  dropZone.addEventListener("drop", handleDrop);

  // Tools
  sortNaturalBtn.addEventListener("click", sortByNatural);
  sortAlphaBtn.addEventListener("click", sortByAlpha);
  reverseOrderBtn.addEventListener("click", reverseOrder);
  clearAllBtn.addEventListener("click", clearAllFiles);

  // Radio cards UI sync
  document.querySelectorAll('input[name="mergeMode"]').forEach(radio => {
    radio.addEventListener("change", updateRadioCardStyles);
  });

  document.querySelectorAll('input[name="exportFormat"]').forEach(radio => {
    radio.addEventListener("change", updateChipStyles);
  });

  // Merge trigger
  mergeBtn.addEventListener("click", startMerge);
}

// Theme Management (Dark & White)
function initTheme() {
  const savedTheme = localStorage.getItem("ppt_theme") || "dark";
  applyTheme(savedTheme);

  if (themeDarkBtn) {
    themeDarkBtn.addEventListener("click", () => applyTheme("dark"));
  }
  if (themeLightBtn) {
    themeLightBtn.addEventListener("click", () => applyTheme("light"));
  }
}

function applyTheme(theme) {
  document.documentElement.setAttribute("data-theme", theme);
  localStorage.setItem("ppt_theme", theme);
  if (themeDarkBtn && themeLightBtn) {
    themeDarkBtn.classList.toggle("active", theme === "dark");
    themeLightBtn.classList.toggle("active", theme === "light");
  }
}

// Handle Drag & Drop Files
function handleDrop(e) {
  const dt = e.dataTransfer;
  const files = Array.from(dt.files);
  if (files.length > 0) {
    processIncomingFiles(files);
  }
}

function handleFileSelect(e) {
  const files = Array.from(e.target.files);
  if (files.length > 0) {
    processIncomingFiles(files);
  }
  fileInput.value = ""; // Reset
}

// Process up to 100 files in batches error-free
async function processIncomingFiles(files) {
  if (isUploading) return;

  const validExtensions = [".pptx", ".ppt"];
  const pptFiles = files.filter(f => {
    // Ignore Windows temporary / hidden lock files like ~$Presentation.pptx
    if (f.name.startsWith("~$")) return false;
    const ext = f.name.substring(f.name.lastIndexOf(".")).toLowerCase();
    return validExtensions.includes(ext);
  });

  if (pptFiles.length === 0) {
    alert("Please upload PowerPoint files (.pptx or .ppt).");
    return;
  }

  const remainingSlots = 100 - fileList.length;
  if (remainingSlots <= 0) {
    alert("Maximum limit of 100 presentations reached. Please clear or remove some files.");
    return;
  }

  const filesToUpload = pptFiles.slice(0, remainingSlots);
  if (pptFiles.length > remainingSlots) {
    alert(`Only ${remainingSlots} presentations could be added (max 100 limit).`);
  }

  isUploading = true;
  setSystemState("UPLOADING", true);
  uploadProgressContainer.style.display = "block";
  uploadProgressBar.style.width = "5%";
  if (uploadProgressPct) uploadProgressPct.textContent = "5%";
  uploadProgressText.textContent = `PREPARING ${filesToUpload.length} PRESENTATIONS...`;

  // Upload one presentation at a time for maximum reliability, per-file error isolation,
  // and staying strictly under serverless payload limits
  let processedCount = 0;

  for (let i = 0; i < filesToUpload.length; i++) {
    const file = filesToUpload[i];
    const truncName = file.name.length > 25 ? file.name.substring(0, 22) + "..." : file.name;
    uploadProgressText.textContent = `UPLOADING [${i + 1}/${filesToUpload.length}] ${truncName}`;

    // On Vercel free tier, serverless requests cannot exceed 4.5MB
    if (file.size > 4.5 * 1024 * 1024 && window.location.hostname.includes("vercel.app")) {
      fileList.push({
        id: "err_" + Math.random().toString(36).substring(7),
        original_name: file.name,
        size_bytes: file.size,
        slide_count: 0,
        is_valid: false,
        error_message: "File exceeds 4.5MB Vercel serverless limit. Please run locally for unlimited file sizes.",
        aspect_ratio: "—",
        dimensions: "—"
      });
      processedCount++;
      const pct = Math.round((processedCount / filesToUpload.length) * 100);
      uploadProgressBar.style.width = `${pct}%`;
      if (uploadProgressPct) uploadProgressPct.textContent = `${pct}%`;
      renderTable();
      continue;
    }

    const formData = new FormData();
    formData.append("files", file);

    try {
      const response = await fetch("/api/upload", {
        method: "POST",
        body: formData
      });

      if (!response.ok) {
        throw new Error(`Upload failed with status ${response.status}`);
      }

      const result = await response.json();
      if (result.success && result.files) {
        fileList.push(...result.files);
      } else {
        fileList.push({
          id: "err_" + Math.random().toString(36).substring(7),
          original_name: file.name,
          size_bytes: file.size,
          slide_count: 0,
          is_valid: false,
          error_message: result.error || "Validation failed",
          aspect_ratio: "—",
          dimensions: "—"
        });
      }
    } catch (err) {
      console.error("Upload error for file:", file.name, err);
      fileList.push({
        id: "err_" + Math.random().toString(36).substring(7),
        original_name: file.name,
        size_bytes: file.size,
        slide_count: 0,
        is_valid: false,
        error_message: `Upload error: ${err.message}`,
        aspect_ratio: "—",
        dimensions: "—"
      });
    }

    processedCount++;
    const pct = Math.round((processedCount / filesToUpload.length) * 100);
    uploadProgressBar.style.width = `${pct}%`;
    if (uploadProgressPct) uploadProgressPct.textContent = `${pct}%`;
    renderTable();
  }

  setTimeout(() => {
    uploadProgressContainer.style.display = "none";
    isUploading = false;
    setSystemState(fileList.length > 0 ? "READY" : "IDLE", false);
    renderTable();
  }, 400);
}

// Render the interactive File Table
function renderTable() {
  fileTableBody.innerHTML = "";

  if (fileList.length === 0) {
    fileTableBody.appendChild(emptyRow);
    emptyRow.style.display = "";
    mergeBtn.disabled = true;
    setSystemState("IDLE", false);
    updateBadges();
    return;
  }

  emptyRow.style.display = "none";
  let hasValidFiles = false;

  fileList.forEach((file, index) => {
    if (file.is_valid) hasValidFiles = true;

    const tr = document.createElement("tr");
    tr.dataset.index = index;
    tr.draggable = true;

    // Drag and drop row listeners
    tr.addEventListener("dragstart", handleRowDragStart);
    tr.addEventListener("dragover", handleRowDragOver);
    tr.addEventListener("drop", handleRowDrop);
    tr.addEventListener("dragend", handleRowDragEnd);

    // Formatted size
    const sizeStr = formatBytes(file.size_bytes);
    const lineIndexStr = `[${String(index + 1).padStart(2, "0")}]`;

    tr.innerHTML = `
      <td><span class="line-num-mono">${lineIndexStr}</span></td>
      <td>
        <div class="file-name-wrapper" title="${escapeHtml(file.original_name)}">
          <span class="file-name-text">${escapeHtml(file.original_name)}</span>
        </div>
      </td>
      <td>
        <span class="slide-count-mono">${file.is_valid ? String(file.slide_count).padStart(2, "0") + ' SLIDES' : '—'}</span>
      </td>
      <td>
        <span class="file-size-mono">${sizeStr}</span>
      </td>
      <td>
        ${file.is_valid 
          ? `<span class="status-pill status-ready">READY [${file.aspect_ratio}]</span>` 
          : `<span class="status-pill status-corrupt" title="${escapeHtml(file.error_message || 'Invalid presentation')}">ERROR</span>`}
      </td>
      <td>
        <div class="order-actions">
          <button class="order-btn" onclick="moveRow(${index}, -1)" title="Move Up" ${index === 0 ? 'disabled' : ''}>&uarr;</button>
          <button class="order-btn" onclick="moveRow(${index}, 1)" title="Move Down" ${index === fileList.length - 1 ? 'disabled' : ''}>&darr;</button>
          <button class="order-btn del-btn" onclick="removeFile(${index})" title="Remove line">&times;</button>
        </div>
      </td>
    `;
    fileTableBody.appendChild(tr);
  });

  mergeBtn.disabled = !hasValidFiles || isMerging;
  updateBadges();
}

// Update Header Statistics
function updateBadges() {
  const currentCount = String(fileList.length).padStart(2, "0");
  fileCountBadge.textContent = `${currentCount} / 100`;

  const totalSlides = fileList.reduce((acc, f) => acc + (f.is_valid ? f.slide_count : 0), 0);
  totalSlidesBadge.textContent = String(totalSlides).padStart(3, "0");
}

function setSystemState(stateText, isActive) {
  if (!systemStateBadge) return;
  systemStateBadge.textContent = stateText;
  if (isActive) {
    systemStateBadge.className = "telemetry-value state-active";
  } else {
    systemStateBadge.className = "telemetry-value state-idle";
  }
}

// Row Movement
window.moveRow = function(index, direction) {
  const newIndex = index + direction;
  if (newIndex < 0 || newIndex >= fileList.length) return;
  const item = fileList.splice(index, 1)[0];
  fileList.splice(newIndex, 0, item);
  renderTable();
};

window.removeFile = function(index) {
  fileList.splice(index, 1);
  renderTable();
};

// Drag and drop row reordering
let draggedRowIndex = null;

function handleRowDragStart(e) {
  draggedRowIndex = parseInt(this.dataset.index);
  this.classList.add("dragging");
  e.dataTransfer.effectAllowed = "move";
}

function handleRowDragOver(e) {
  e.preventDefault();
  e.dataTransfer.dropEffect = "move";
}

function handleRowDrop(e) {
  e.preventDefault();
  const targetIndex = parseInt(this.dataset.index);
  if (draggedRowIndex !== null && draggedRowIndex !== targetIndex) {
    const item = fileList.splice(draggedRowIndex, 1)[0];
    fileList.splice(targetIndex, 0, item);
    renderTable();
  }
}

function handleRowDragEnd() {
  this.classList.remove("dragging");
  draggedRowIndex = null;
}

// Sorting Functions
function sortByNatural() {
  const collator = new Intl.Collator(undefined, { numeric: true, sensitivity: 'base' });
  fileList.sort((a, b) => collator.compare(a.original_name, b.original_name));
  renderTable();
}

function sortByAlpha() {
  fileList.sort((a, b) => a.original_name.localeCompare(b.original_name));
  renderTable();
}

function reverseOrder() {
  fileList.reverse();
  renderTable();
}

async function clearAllFiles() {
  if (fileList.length === 0) return;
  if (confirm("Remove all uploaded files from the queue?")) {
    fileList = [];
    renderTable();
    try {
      await fetch("/api/clear", { method: "POST" });
    } catch (e) {}
  }
}

// Sync Form Styles
function updateRadioCardStyles() {
  const val = document.querySelector('input[name="mergeMode"]:checked').value;
  document.getElementById("labelSequential").classList.toggle("selected", val === "sequential");
  document.getElementById("labelInterleaved").classList.toggle("selected", val === "interleaved");
}

function updateChipStyles() {
  const val = document.querySelector('input[name="exportFormat"]:checked').value;
  document.getElementById("chipBoth").classList.toggle("selected", val === "both");
  document.getElementById("chipPptx").classList.toggle("selected", val === "pptx");
  document.getElementById("chipPdf").classList.toggle("selected", val === "pdf");
}

// Start Merge & Export Process
async function startMerge() {
  const validFiles = fileList.filter(f => f.is_valid);
  if (validFiles.length === 0) {
    alert("No valid presentations to merge.");
    return;
  }

  isMerging = true;
  mergeBtn.disabled = true;
  setSystemState("PROCESSING", true);
  downloadCard.style.display = "none";
  progressPanel.style.display = "block";
  progressTitle.textContent = "INITIALIZING";
  progressPct.textContent = "0%";
  mergeProgressBar.style.width = "0%";
  progressStatusMsg.textContent = "Connecting to merge engine...";

  const mode = document.querySelector('input[name="mergeMode"]:checked').value;
  const exportFormat = document.querySelector('input[name="exportFormat"]:checked').value;
  const outputName = document.getElementById("outputFilename").value.trim() || "Unified_Presentation";

  try {
    const payload = {
      file_ids: validFiles.map(f => f.id),
      mode: mode,
      export_format: exportFormat,
      output_name: outputName
    };

    const res = await fetch("/api/merge", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(payload)
    });

    const data = await res.json();
    if (!data.success) {
      throw new Error(data.error || "Failed to start merge task");
    }

    pollTaskStatus(data.taskId);
  } catch (err) {
    console.error("Merge trigger failed:", err);
    alert(`Could not start merge: ${err.message}`);
    isMerging = false;
    mergeBtn.disabled = false;
    setSystemState("READY", false);
    progressPanel.style.display = "none";
  }
}

// Poll Task Status
function pollTaskStatus(taskId) {
  if (activePollInterval) clearInterval(activePollInterval);

  activePollInterval = setInterval(async () => {
    try {
      const res = await fetch(`/api/task-status?taskId=${taskId}`);
      const task = await res.json();

      if (task.status === "processing" || task.status === "queued") {
        const pct = task.progress || 0;
        mergeProgressBar.style.width = `${pct}%`;
        progressPct.textContent = `${pct}%`;
        progressTitle.textContent = pct >= 90 ? "PDF EXPORT" : "MERGING SLIDES";
        progressStatusMsg.textContent = task.message || "Processing presentation lines...";
      } else if (task.status === "completed") {
        clearInterval(activePollInterval);
        mergeProgressBar.style.width = "100%";
        progressPct.textContent = "100%";
        progressTitle.textContent = "COMPLETED";
        progressStatusMsg.textContent = "All files successfully unified and ready for download.";

        showDownloadCard(task);
        isMerging = false;
        mergeBtn.disabled = false;
        setSystemState("COMPLETE", true);
      } else if (task.status === "failed") {
        clearInterval(activePollInterval);
        progressTitle.textContent = "FAILED";
        progressStatusMsg.textContent = task.message || task.error || "An error occurred.";
        alert(`Merge error: ${task.message || task.error}`);
        isMerging = false;
        mergeBtn.disabled = false;
        setSystemState("ERROR", false);
      }
    } catch (err) {
      console.error("Polling error:", err);
    }
  }, 750);
}

// Display Ready Download Card
function showDownloadCard(task) {
  downloadCard.style.display = "block";
  const outputName = document.getElementById("outputFilename").value.trim() || "Unified_Presentation";
  downloadFileName.textContent = outputName;
  downloadMeta.textContent = `TOTAL SLIDES: ${task.total_slides || 0} • DECKS MERGED: ${task.files_count || 0}`;

  const downloads = task.downloads || {};

  if (downloads.pptx) {
    const pptxName = downloads.pptx.split("/").pop();
    downloadPptxBtn.href = downloads.pptx;
    downloadPptxBtn.setAttribute("download", pptxName);
    downloadPptxBtn.onclick = (e) => {
      e.preventDefault();
      downloadFileViaBlob(downloads.pptx, pptxName, downloadPptxBtn);
    };
    downloadPptxBtn.style.display = "flex";
  } else {
    downloadPptxBtn.style.display = "none";
  }

  if (downloads.pdf) {
    const pdfName = downloads.pdf.split("/").pop();
    downloadPdfBtn.href = downloads.pdf;
    downloadPdfBtn.setAttribute("download", pdfName);
    downloadPdfBtn.onclick = (e) => {
      e.preventDefault();
      downloadFileViaBlob(downloads.pdf, pdfName, downloadPdfBtn);
    };
    downloadPdfBtn.style.display = "flex";
  } else {
    downloadPdfBtn.style.display = "none";
  }

  if (downloads.zip) {
    const zipName = downloads.zip.split("/").pop();
    downloadZipBtn.href = downloads.zip;
    downloadZipBtn.setAttribute("download", zipName);
    downloadZipBtn.onclick = (e) => {
      e.preventDefault();
      downloadFileViaBlob(downloads.zip, zipName, downloadZipBtn);
    };
    downloadZipBtn.style.display = "flex";
  } else {
    downloadZipBtn.style.display = "none";
  }

  downloadCard.scrollIntoView({ behavior: "smooth", block: "nearest" });
}

// Guaranteed Complete Download via Local Blob
async function downloadFileViaBlob(url, filename, tileElement) {
  const actionSpan = tileElement.querySelector(".tile-action");
  const origText = actionSpan ? actionSpan.textContent : "DOWNLOAD";
  if (actionSpan) actionSpan.textContent = "SAVING...";
  tileElement.style.opacity = "0.7";
  tileElement.style.pointerEvents = "none";

  try {
    const res = await fetch(url);
    if (!res.ok) throw new Error(`HTTP ${res.status}`);
    const blob = await res.blob();
    const blobUrl = window.URL.createObjectURL(blob);
    const tempLink = document.createElement("a");
    tempLink.href = blobUrl;
    tempLink.download = filename;
    document.body.appendChild(tempLink);
    tempLink.click();

    setTimeout(() => {
      window.URL.revokeObjectURL(blobUrl);
      tempLink.remove();
    }, 1200);

    if (actionSpan) actionSpan.textContent = "SAVED ✓";
    setTimeout(() => {
      if (actionSpan) actionSpan.textContent = origText;
    }, 2000);
  } catch (err) {
    console.warn("Direct blob download failed, falling back to window location:", err);
    window.location.href = url;
    if (actionSpan) actionSpan.textContent = origText;
  } finally {
    tileElement.style.opacity = "1";
    tileElement.style.pointerEvents = "auto";
  }
}

// Utilities
function formatBytes(bytes) {
  if (!bytes || bytes === 0) return "0 B";
  const k = 1024;
  const sizes = ["B", "KB", "MB", "GB"];
  const i = Math.floor(Math.log(bytes) / Math.log(k));
  return parseFloat((bytes / Math.pow(k, i)).toFixed(1)) + " " + sizes[i];
}

function escapeHtml(str) {
  if (!str) return "";
  return str.replace(/[&<>"']/g, function(m) {
    return {
      '&': '&amp;',
      '<': '&lt;',
      '>': '&gt;',
      '"': '&quot;',
      "'": '&#039;'
    }[m];
  });
}

// Launch
window.addEventListener("DOMContentLoaded", init);
