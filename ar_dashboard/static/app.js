/* ar_dashboard/static/app.js */

// ── Toast ─────────────────────────────────────────────────────────────────
const toast = document.getElementById("toast");
let toastTimer = null;

function showToast(msg, type = "success") {
  toast.textContent = msg;
  toast.className   = `show ${type}`;
  clearTimeout(toastTimer);
  toastTimer = setTimeout(() => { toast.className = ""; }, 4000);
}

// ── Tab switching ─────────────────────────────────────────────────────────
document.querySelectorAll(".nav-tab").forEach(btn => {
  btn.addEventListener("click", () => {
    document.querySelectorAll(".nav-tab").forEach(t => t.classList.remove("active"));
    document.querySelectorAll(".tab-panel").forEach(p => p.classList.remove("active"));
    btn.classList.add("active");
    const panel = document.getElementById("panel-" + btn.dataset.tab);
    if (panel) panel.classList.add("active");
  });
});

// ── Run automation modal ──────────────────────────────────────────────────
const btnRun     = document.getElementById("btn-run");
const overlay    = document.getElementById("run-overlay");
const modalTitle = document.getElementById("modal-title");
const modalMsg   = document.getElementById("modal-msg");
const output     = document.getElementById("run-output");
const btnClose   = document.getElementById("btn-close");
const btnRefresh = document.getElementById("btn-refresh");

function openModal(running) {
  overlay.classList.add("active");
  output.style.display = "none";
  output.textContent   = "";
  btnClose.style.display   = running ? "none" : "inline-block";
  btnRefresh.style.display = "none";
  if (running) {
    modalTitle.textContent = "Running Automation…";
    modalMsg.innerHTML = `<span class="spinner"></span> Processing AR data and generating email drafts. This may take up to 30 seconds.`;
  }
}

function resolveModal(ok, outputText) {
  modalTitle.textContent   = ok ? "Automation Complete" : "Automation Failed";
  modalMsg.textContent     = ok
    ? "AR data has been refreshed and email drafts have been updated."
    : "The automation script encountered an error.";
  output.textContent       = outputText || "";
  output.style.display     = "block";
  btnClose.style.display   = "inline-block";
  btnRefresh.style.display = ok ? "inline-block" : "none";
}

btnRun.addEventListener("click", async () => {
  btnRun.disabled = true;
  openModal(true);
  try {
    const resp = await fetch("/run-automation", { method: "POST" });
    const data = await resp.json();
    resolveModal(data.ok, data.output);
    showToast(data.ok ? "Automation completed — data refreshed." : "Automation failed. See output for details.", data.ok ? "success" : "error");
  } catch (err) {
    resolveModal(false, String(err));
    showToast("Network error — could not reach the server.", "error");
  } finally {
    btnRun.disabled = false;
  }
});

btnClose.addEventListener("click", () => overlay.classList.remove("active"));
btnRefresh.addEventListener("click", () => window.location.reload());
overlay.addEventListener("click", e => { if (e.target === overlay) overlay.classList.remove("active"); });

// ── Scan Inbox ────────────────────────────────────────────────────────────
const btnScanInbox    = document.getElementById("btn-scan-inbox");
const scanOverlay     = document.getElementById("scan-overlay");
const scanModalTitle  = document.getElementById("scan-modal-title");
const scanModalMsg    = document.getElementById("scan-modal-msg");
const scanResultsList = document.getElementById("scan-results-list");
const btnScanClose    = document.getElementById("btn-scan-close");
const btnScanRefresh  = document.getElementById("btn-scan-refresh");

function escHtml(str) {
  return String(str || "")
    .replace(/&/g, "&amp;").replace(/</g, "&lt;")
    .replace(/>/g, "&gt;").replace(/"/g, "&quot;");
}

function openScanModal() {
  scanOverlay.classList.add("active");
  scanResultsList.innerHTML    = "";
  btnScanClose.style.display   = "none";
  btnScanRefresh.style.display = "none";
  scanModalTitle.textContent   = "Scanning Inbox\u2026";
  scanModalMsg.innerHTML       = `<span class="spinner"></span> Checking for customer replies. This may take a moment.`;
}

function resolveScanModal(data) {
  if (!data.ok) {
    scanModalTitle.textContent = "Scan Failed";
    scanModalMsg.textContent   = data.error || "An error occurred.";
    btnScanClose.style.display = "inline-block";
    return;
  }

  const n  = data.new_count        || 0;
  const s  = data.skipped          || 0;
  const sf = data.skipped_filtered || 0;

  scanModalTitle.textContent = n > 0
    ? `${n} new ${n === 1 ? "reply" : "replies"} logged`
    : "No new replies found";

  let subMsg = "";
  if (sf > 0) subMsg += `${sf} filtered out (internal / automated). `;
  if (s  > 0) subMsg += `${s} already logged (skipped). `;
  if (n === 0) subMsg += "No new customer replies detected in the last 7 days.";
  scanModalMsg.textContent = subMsg.trim();

  if (n > 0 && data.items && data.items.length) {
    let html = '<div class="scan-results-list">';
    for (const item of data.items) {
      const inv = item.invoice
        ? `<span class="reply-invoice">inv ${escHtml(item.invoice)}</span>`
        : "";
      html += `
        <div class="scan-result-item">
          <div class="scan-result-meta">
            <span class="reply-customer">${escHtml(item.customer)}</span>
            ${inv}
            <span class="reply-date">${escHtml(item.date)}</span>
          </div>
          <div class="reply-notes">${escHtml(item.summary)}</div>
        </div>`;
    }
    html += "</div>";
    scanResultsList.innerHTML = html;
  }

  btnScanClose.style.display   = "inline-block";
  btnScanRefresh.style.display = n > 0 ? "inline-block" : "none";
}

btnScanInbox.addEventListener("click", async () => {
  btnScanInbox.disabled = true;
  openScanModal();
  try {
    const resp = await fetch("/scan-inbox", { method: "POST" });
    const data = await resp.json();
    resolveScanModal(data);
    if (data.ok) {
      const n = data.new_count || 0;
      showToast(
        n > 0
          ? `${n} new ${n === 1 ? "reply" : "replies"} logged from inbox.`
          : "Inbox scanned — no new replies found.",
        "success"
      );
    } else {
      showToast(data.error || "Scan failed.", "error");
    }
  } catch (err) {
    resolveScanModal({ ok: false, error: String(err) });
    showToast("Network error — " + err, "error");
  } finally {
    btnScanInbox.disabled = false;
  }
});

btnScanClose.addEventListener("click", () => scanOverlay.classList.remove("active"));
btnScanRefresh.addEventListener("click", () => window.location.reload());
scanOverlay.addEventListener("click", e => {
  if (e.target === scanOverlay) scanOverlay.classList.remove("active");
});

// ── Upload AR Export tab ──────────────────────────────────────────────────
const uploadZone     = document.getElementById("upload-zone");
const fileInput      = document.getElementById("file-input");
const filePreview    = document.getElementById("upload-file-preview");
const fileNameEl     = document.getElementById("upload-file-name");
const uploadActions  = document.getElementById("upload-actions");
const btnUpload      = document.getElementById("btn-upload");
const uploadClear    = document.getElementById("upload-clear");
const uploadStatus   = document.getElementById("upload-status");

let selectedFile = null;

function setUploadFile(file) {
  if (!file || !file.name.toLowerCase().endsWith(".csv")) {
    showUploadStatus("Please select a .csv file.", "error");
    return;
  }
  selectedFile = file;
  fileNameEl.textContent       = file.name;
  filePreview.style.display    = "flex";
  uploadActions.style.display  = "flex";
  uploadStatus.innerHTML       = "";
}

function clearUploadFile() {
  selectedFile                 = null;
  fileInput.value              = "";
  filePreview.style.display    = "none";
  uploadActions.style.display  = "none";
  uploadStatus.innerHTML       = "";
}

function showUploadStatus(msg, type) {
  uploadStatus.innerHTML = `<div class="upload-status-msg ${type}">${msg}</div>`;
}

uploadZone.addEventListener("click", () => fileInput.click());

uploadZone.addEventListener("dragover", e => {
  e.preventDefault();
  uploadZone.classList.add("drag-over");
});

uploadZone.addEventListener("dragleave", () => uploadZone.classList.remove("drag-over"));

uploadZone.addEventListener("drop", e => {
  e.preventDefault();
  uploadZone.classList.remove("drag-over");
  const file = e.dataTransfer.files[0];
  if (file) setUploadFile(file);
});

fileInput.addEventListener("change", () => {
  if (fileInput.files[0]) setUploadFile(fileInput.files[0]);
});

uploadClear.addEventListener("click", clearUploadFile);

btnUpload.addEventListener("click", async () => {
  if (!selectedFile) return;
  btnUpload.disabled = true;
  btnUpload.textContent = "Uploading…";

  const formData = new FormData();
  formData.append("file", selectedFile);

  try {
    const resp = await fetch("/upload-ar", { method: "POST", body: formData });
    const data = await resp.json();
    if (data.ok) {
      showUploadStatus("✓ Uploaded successfully — exports/ar_aging.csv has been replaced.", "success");
      showToast("AR export uploaded.", "success");
      clearUploadFile();
    } else {
      showUploadStatus(data.error || "Upload failed.", "error");
    }
  } catch (err) {
    showUploadStatus("Network error — " + err, "error");
  } finally {
    btnUpload.disabled    = false;
    btnUpload.textContent = "Upload File";
  }
});

// ── Weekly Report — Send Report (Graph API draft) ─────────────────────────
const btnSendScott      = document.getElementById("btn-send-scott");
const sendReportStatus  = document.getElementById("send-report-status");

if (btnSendScott) {
  btnSendScott.addEventListener("click", async () => {
    btnSendScott.disabled    = true;
    btnSendScott.textContent = "Creating draft…";
    if (sendReportStatus) sendReportStatus.innerHTML = "";

    try {
      const resp = await fetch("/send-report", { method: "POST" });
      const data = await resp.json();

      if (data.ok) {
        if (sendReportStatus) {
          sendReportStatus.innerHTML =
            `<span style="color:#2e7d32">&#10003; Draft saved to Outlook — ready to review and send.</span>`;
        }
        showToast("Report draft created in Outlook.", "success");
      } else {
        if (sendReportStatus) {
          sendReportStatus.innerHTML =
            `<span style="color:#c0392b">&#10005; ${data.error || "Failed to create draft."}</span>`;
        }
        showToast(data.error || "Failed to create draft.", "error");
      }
    } catch (err) {
      if (sendReportStatus) {
        sendReportStatus.innerHTML =
          `<span style="color:#c0392b">&#10005; Network error — ${err}</span>`;
      }
      showToast("Network error — " + err, "error");
    } finally {
      btnSendScott.disabled    = false;
      btnSendScott.textContent = "✉ Send Report";
    }
  });
}

// ── Customer Replies tab ──────────────────────────────────────────────────
const btnGenerate   = document.getElementById("btn-generate");
const summaryCard   = document.getElementById("summary-card");
const summaryText   = document.getElementById("summary-text");
const btnRedo       = document.getElementById("btn-redo");
const btnSave       = document.getElementById("btn-save");
const replyStatus   = document.getElementById("reply-status");

function showReplyStatus(msg, type) {
  replyStatus.innerHTML = `<div class="reply-status-msg ${type}">${msg}</div>`;
}

async function generateSummary() {
  const customer = document.getElementById("reply-customer").value.trim();
  const invoice  = document.getElementById("reply-invoice").value.trim();
  const reply    = document.getElementById("reply-text").value.trim();

  if (!invoice || !reply) {
    showReplyStatus("Invoice number and reply text are required.", "error");
    return;
  }

  btnGenerate.disabled    = true;
  btnGenerate.textContent = "Generating…";
  summaryCard.style.display = "none";
  replyStatus.innerHTML   = "";

  try {
    const resp = await fetch("/summarize-reply", {
      method:  "POST",
      headers: { "Content-Type": "application/json" },
      body:    JSON.stringify({ customer, invoice, reply }),
    });
    const data = await resp.json();

    if (data.ok) {
      summaryText.textContent    = data.summary;
      summaryCard.style.display  = "block";
      // store for save step
      btnSave.dataset.customer = customer;
      btnSave.dataset.invoice  = invoice;
      btnSave.dataset.summary  = data.summary;
    } else {
      showReplyStatus(data.error || "Failed to generate summary.", "error");
    }
  } catch (err) {
    showReplyStatus("Network error — " + err, "error");
  } finally {
    btnGenerate.disabled    = false;
    btnGenerate.textContent = "Generate Summary";
  }
}

btnGenerate.addEventListener("click", generateSummary);
btnRedo.addEventListener("click", generateSummary);

btnSave.addEventListener("click", async () => {
  const customer = btnSave.dataset.customer || "";
  const invoice  = btnSave.dataset.invoice  || "";
  const summary  = btnSave.dataset.summary  || "";

  if (!invoice || !summary) return;

  btnSave.disabled    = true;
  btnSave.textContent = "Saving…";

  try {
    const resp = await fetch("/log-reply", {
      method:  "POST",
      headers: { "Content-Type": "application/json" },
      body:    JSON.stringify({ customer, invoice, summary }),
    });
    const data = await resp.json();

    if (data.ok) {
      summaryCard.style.display = "none";
      showReplyStatus("✓ Reply logged to customer_interactions.csv.", "success");
      showToast("Reply logged successfully.", "success");
      // reset form
      document.getElementById("reply-customer").value = "";
      document.getElementById("reply-invoice").value  = "";
      document.getElementById("reply-text").value     = "";
    } else {
      showReplyStatus(data.error || "Failed to save.", "error");
    }
  } catch (err) {
    showReplyStatus("Network error — " + err, "error");
  } finally {
    btnSave.disabled    = false;
    btnSave.textContent = "✓ Save to Log";
  }
});

// ── Delete logged reply ────────────────────────────────────────────────────
document.addEventListener("click", async (e) => {
  const btn = e.target.closest(".btn-delete-reply");
  if (!btn) return;

  const row = btn.closest(".logged-reply-item");
  btn.disabled = true;
  btn.textContent = "…";

  try {
    const resp = await fetch("/delete-reply", {
      method:  "POST",
      headers: { "Content-Type": "application/json" },
      body:    JSON.stringify({
        raw_date: btn.dataset.rawDate,
        invoice:  btn.dataset.invoice,
        customer: btn.dataset.customer,
        notes:    btn.dataset.notes,
      }),
    });
    const data = await resp.json();
    if (data.ok) {
      row.style.transition = "opacity 0.2s";
      row.style.opacity = "0";
      setTimeout(() => row.remove(), 200);
      showToast("Entry deleted.", "success");
    } else {
      btn.disabled = false;
      btn.textContent = "✕";
      showToast(data.error || "Could not delete entry.", "error");
    }
  } catch (err) {
    btn.disabled = false;
    btn.textContent = "✕";
    showToast("Network error — " + err, "error");
  }
});
