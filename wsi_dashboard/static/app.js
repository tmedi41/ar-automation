/* wsi_dashboard/static/app.js */

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
function activateTab(tabName) {
  document.querySelectorAll(".nav-tab").forEach(t => t.classList.remove("active"));
  document.querySelectorAll(".tab-panel").forEach(p => p.classList.remove("active"));
  const btn = document.querySelector(`.nav-tab[data-tab="${tabName}"]`);
  const panel = document.getElementById("panel-" + tabName);
  if (btn) btn.classList.add("active");
  if (panel) panel.classList.add("active");
  if (tabName === "payment") initPaymentCharts();
}

document.querySelectorAll(".nav-tab").forEach(btn => {
  btn.addEventListener("click", () => {
    const tabName = btn.dataset.tab;
    location.hash = tabName;
    activateTab(tabName);
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
      const n = data.logged || 1;
      const entryLabel = n === 1 ? "1 entry" : `${n} entries`;
      showReplyStatus(`✓ Reply logged (${entryLabel} saved).`, "success");
      showToast(`Reply logged — ${entryLabel} saved.`, "success");
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

// ── Payment History charts ────────────────────────────────────────────────
let phChartsInit = false;

function initPaymentCharts() {
  if (phChartsInit) return;
  phChartsInit = true;

  const BEIGE_GRID = "rgba(0,0,0,0.07)";
  const TICK_COLOR = "#888780";
  const FONT       = { family: "'DM Sans', sans-serif", size: 12 };

  const baseScales = {
    x: { grid: { color: BEIGE_GRID }, ticks: { color: TICK_COLOR, font: FONT } },
    y: { grid: { color: BEIGE_GRID }, ticks: { color: TICK_COLOR, font: FONT }, beginAtZero: true },
  };

  // Past Due Trend — bar chart (8 weeks)
  const trendEl = document.getElementById("chart-past-due-trend");
  if (trendEl && typeof PH_WEEKLY !== "undefined" && PH_WEEKLY.length) {
    new Chart(trendEl, {
      type: "bar",
      data: {
        labels:   PH_WEEKLY.map(d => d.label),
        datasets: [{
          label:           "Past Due Contacts",
          data:            PH_WEEKLY.map(d => d.count),
          backgroundColor: PH_WEEKLY.map((d, i) =>
            i === PH_WEEKLY.length - 1 ? "#D85A30" : "rgba(216,90,48,0.35)"
          ),
          borderRadius:    4,
          borderSkipped:   false,
        }],
      },
      options: {
        responsive:          true,
        maintainAspectRatio: false,
        plugins: {
          legend: { display: false },
          tooltip: {
            callbacks: { label: ctx => ` ${ctx.parsed.y} contact${ctx.parsed.y !== 1 ? "s" : ""}` },
          },
        },
        scales: {
          x: { ...baseScales.x, grid: { display: false } },
          y: { ...baseScales.y, ticks: { ...baseScales.y.ticks, stepSize: 1 } },
        },
      },
    });
  }

  // Top Overdue Accounts — horizontal bar chart
  const overdueEl = document.getElementById("chart-top-overdue");
  if (overdueEl && typeof PH_OVERDUE !== "undefined" && PH_OVERDUE.length) {
    const labels = PH_OVERDUE.map(d =>
      d.customer.length > 22 ? d.customer.slice(0, 22) + "\u2026" : d.customer
    );
    const colors = PH_OVERDUE.map(d =>
      d.type === "ESCALATION" ? "#D85A30" : "#E8A838"
    );
    new Chart(overdueEl, {
      type: "bar",
      data: {
        labels,
        datasets: [{
          label:           "Balance Owed",
          data:            PH_OVERDUE.map(d => d.balance),
          backgroundColor: colors,
          borderRadius:    4,
          borderSkipped:   false,
        }],
      },
      options: {
        indexAxis:           "y",
        responsive:          true,
        maintainAspectRatio: false,
        plugins: {
          legend: { display: false },
          tooltip: {
            callbacks: {
              label: ctx => ` $${ctx.parsed.x.toLocaleString("en-US", { minimumFractionDigits: 2 })}`,
            },
          },
        },
        scales: {
          x: {
            ...baseScales.x,
            ticks: {
              ...baseScales.x.ticks,
              callback: v => "$" + (v >= 1000 ? (v / 1000).toFixed(0) + "k" : v),
            },
          },
          y: { ...baseScales.y, grid: { display: false } },
        },
      },
    });
  }
}


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

// ── Restore active tab from URL hash on page load ─────────────────────────
(function () {
  const hash = location.hash.replace("#", "");
  const valid = ["overview", "upload", "replies", "weekly", "payment"];
  if (!hash || !valid.includes(hash)) return;

  if (hash === "payment") {
    // Activate the panel immediately so the canvas is in the DOM and visible,
    // but defer chart init by 100ms to ensure the canvas has rendered dimensions
    // before Chart.js tries to measure it.
    document.querySelectorAll(".nav-tab").forEach(t => t.classList.remove("active"));
    document.querySelectorAll(".tab-panel").forEach(p => p.classList.remove("active"));
    const btn   = document.querySelector('.nav-tab[data-tab="payment"]');
    const panel = document.getElementById("panel-payment");
    if (btn)   btn.classList.add("active");
    if (panel) panel.classList.add("active");
    setTimeout(initPaymentCharts, 100);
  } else {
    activateTab(hash);
  }
})();
