// DataCleanup Pro — Main Application
// Uses Pyodide for in-browser Python (pandas) processing
// Rate limit via localStorage

const RATE_LIMIT = 3;
const STORAGE_KEY = "dc_rate_limits";
const PYODIDE_INDEX_URL = "https://cdn.jsdelivr.net/pyodide/v0.24.1/full/";

let pyodide = null;
let currentFile = null;
let originalData = null;
let cleanedData = null;
let changes = [];

// ========== DOM Elements ==========
const $ = (sel) => document.querySelector(sel);
const $$ = (sel) => document.querySelectorAll(sel);

const emailGate = $("#email-gate");
const startCleaning = $("#start-cleaning");
const rateMsg = $("#rate-msg");
const uploadArea = $("#upload-area");
const dropZone = $("#drop-zone");
const fileInput = $("#file-input");
const fileInfo = $("#file-info");
const fileName = $("#file-name");
const processBtn = $("#process-btn");
const processingArea = $("#processing-area");
const processingStatus = $("#processing-status");
const resultsArea = $("#results-area");
const statsGrid = $("#stats-grid");
const changesLog = $("#changes-log");
const tabBefore = $("#tab-before");
const tabAfter = $("#tab-after");
const previewTable = $("#preview-table");
const downloadBtn = $("#download-btn");
const resetBtn = $("#reset-btn");
const outOfCredits = $("#out-of-credits");
const creditsBadge = $("#credits-badge");

// ========== Rate Limiting ==========
function getTodayKey() {
  return new Date().toISOString().slice(0, 10);
}

function getRateLimit() {
  const data = JSON.parse(localStorage.getItem(STORAGE_KEY) || "{}");
  const today = getTodayKey();
  if (!data.device || data.device.date !== today) {
    data.device = { date: today, count: 0 };
    localStorage.setItem(STORAGE_KEY, JSON.stringify(data));
  }
  return data.device;
}

function canUse() {
  return getRateLimit().count < RATE_LIMIT;
}

function recordUse() {
  const rl = getRateLimit();
  rl.count += 1;
  const data = JSON.parse(localStorage.getItem(STORAGE_KEY) || "{}");
  data.device = rl;
  localStorage.setItem(STORAGE_KEY, JSON.stringify(data));
  return rl.count;
}

// ========== Pyodide Setup ==========
async function loadPyodide() {
  processingStatus.textContent = "Loading Python runtime...";
  try {
    pyodide = await loadPyodide({ indexURL: PYODIDE_INDEX_URL });
    processingStatus.textContent = "Installing pandas & openpyxl...";
    await pyodide.loadPackage(["pandas", "openpyxl"]);
    processingStatus.textContent = "Ready to clean!";
    return true;
  } catch (err) {
    console.error("Pyodide load failed:", err);
    processingStatus.textContent = "Failed to load Python runtime. Please refresh.";
    return false;
  }
}

// ========== Data Processing (Python) ==========
const PYTHON_CLEANUP_SCRIPT = `
import pandas as pd
import re
import json
from io import StringIO, BytesIO
import base64

def clean_dataframe(df):
    changes = []
    original_shape = df.shape
    
    # 1. Trim whitespace from column names
    df.columns = [str(c).strip() for c in df.columns]
    
    # 2. Trim whitespace from all string cells
    str_cols = df.select_dtypes(include=['object']).columns
    for col in str_cols:
        df[col] = df[col].astype(str).str.strip()
        # Fix 'nan' strings back
        df[col] = df[col].replace('nan', '')
    
    # 3. Remove exact duplicate rows
    before_dedup = len(df)
    df = df.drop_duplicates().reset_index(drop=True)
    dupes_removed = before_dedup - len(df)
    if dupes_removed > 0:
        changes.append({"action": "Removed duplicates", "count": int(dupes_removed)})
    
    # 4. Standardize date columns
    date_patterns = [
        (r'^\\d{1,2}/\\d{1,2}/\\d{2,4}$', 'MDY'),
        (r'^\\d{1,2}-\\d{1,2}-\\d{2,4}$', 'MDY2'),
        (r'^\\d{4}-\\d{2}-\\d{2}$', 'YMD'),
        (r'^\\d{4}/\\d{2}/\\d{2}$', 'YMD2'),
    ]
    
    date_keywords = ['date', 'time', 'dob', 'birth', 'start', 'end', 'created', 'updated']
    for col in df.columns:
        col_lower = col.lower()
        is_date_col = any(kw in col_lower for kw in date_keywords)
        
        if is_date_col or df[col].dtype == 'object':
            try:
                parsed = pd.to_datetime(df[col], infer_datetime_format=True, errors='coerce')
                non_null = parsed.notna().sum()
                if non_null > len(df) * 0.5 and non_null > 0:
                    original_vals = df[col].copy()
                    df[col] = parsed.dt.strftime('%Y-%m-%d')
                    changed = (original_vals != df[col]).sum()
                    if changed > 0:
                        changes.append({"action": f"Standardized dates in '{col}'", "count": int(changed)})
                    continue
            except:
                pass
    
    # 5. Standardize phone numbers
    phone_keywords = ['phone', 'tel', 'mobile', 'cell', 'fax', 'contact']
    phone_pattern = re.compile(r'[\\d\\(\\)\\+\\-\\s\\.]{7,}')
    
    for col in df.columns:
        col_lower = col.lower()
        is_phone_col = any(kw in col_lower for kw in phone_keywords)
        
        if is_phone_col:
            def format_phone(val):
                if pd.isna(val) or str(val).strip() == '':
                    return val
                digits = re.sub(r'\\D', '', str(val))
                if len(digits) == 10:
                    return f"({digits[:3]}) {digits[3:6]}-{digits[6:]}"
                elif len(digits) == 11 and digits[0] == '1':
                    return f"({digits[1:4]}) {digits[4:7]}-{digits[7:]}"
                return val
            
            original_vals = df[col].copy()
            df[col] = df[col].apply(format_phone)
            changed = (original_vals != df[col]).sum()
            if changed > 0:
                changes.append({"action": f"Formatted phones in '{col}'", "count": int(changed)})
    
    # 6. Fix casing in name fields
    name_keywords = ['name', 'first', 'last', 'full_name', 'fullname', 'company', 'org']
    for col in df.columns:
        col_lower = col.lower().replace(' ', '_')
        is_name_col = any(kw in col_lower for kw in name_keywords)
        
        if is_name_col:
            original_vals = df[col].copy()
            # Title case names (but keep all-caps acronyms if short)
            def fix_case(val):
                if pd.isna(val) or str(val).strip() == '':
                    return val
                s = str(val).strip()
                # If all caps and more than 3 chars, title case
                if s == s.upper() and len(s) > 3 and s != s.lower():
                    return s.title()
                # If all lowercase, title case
                if s == s.lower() and len(s) > 1:
                    return s.title()
                return s
            
            df[col] = df[col].apply(fix_case)
            changed = (original_vals != df[col]).sum()
            if changed > 0:
                changes.append({"action": f"Fixed casing in '{col}'", "count": int(changed)})
    
    # 7. Standardize address abbreviations
    addr_keywords = ['address', 'street', 'city', 'state', 'zip']
    addr_replacements = {
        r'\\bSt\\.?\\b': 'Street',
        r'\\bAve\\.?\\b': 'Avenue',
        r'\\bBlvd\\.?\\b': 'Boulevard',
        r'\\bDr\\.?\\b': 'Drive',
        r'\\bLn\\.?\\b': 'Lane',
        r'\\bRd\\.?\\b': 'Road',
        r'\\bCt\\.?\\b': 'Court',
        r'\\bPl\\.?\\b': 'Place',
        r'\\bSq\\.?\\b': 'Square',
        r'\\bCir\\.?\\b': 'Circle',
        r'\\bPky\\.?\\b': 'Parkway',
        r'\\bHwy\\.?\\b': 'Highway',
        r'\\bSte\\.?\\b': 'Suite',
        r'\\bApt\\.?\\b': 'Apartment',
        r'\\bN\\.?\\b': 'North',
        r'\\bS\\.?\\b': 'South',
        r'\\bE\\.?\\b': 'East',
        r'\\bW\\.?\\b': 'West',
    }
    
    for col in df.columns:
        col_lower = col.lower()
        is_addr_col = any(kw in col_lower for kw in addr_keywords)
        
        if is_addr_col:
            original_vals = df[col].copy()
            for pattern, replacement in addr_replacements.items():
                df[col] = df[col].str.replace(pattern, replacement, regex=True)
            changed = (original_vals != df[col]).sum()
            if changed > 0:
                changes.append({"action": f"Standardized addresses in '{col}'", "count": int(changed)})
    
    # 8. Remove completely empty rows
    before_empty = len(df)
    df = df.dropna(how='all').reset_index(drop=True)
    empty_removed = before_empty - len(df)
    if empty_removed > 0:
        changes.append({"action": "Removed empty rows", "count": int(empty_removed)})
    
    # Summary
    summary = {
        "original_rows": int(original_shape[0]),
        "original_cols": int(original_shape[1]),
        "cleaned_rows": int(df.shape[0]),
        "cleaned_cols": int(df.shape[1]),
        "rows_removed": int(original_shape[0] - df.shape[0]),
        "duplicates_removed": int(dupes_removed),
    }
    
    return df, json.dumps(changes), json.dumps(summary)

def process_file(file_content_b64, file_type, filename):
    content_bytes = base64.b64decode(file_content_b64)
    
    if file_type == 'csv':
        # Try different encodings
        for enc in ['utf-8', 'latin-1', 'cp1252']:
            try:
                df = pd.read_csv(BytesIO(content_bytes), encoding=enc)
                break
            except:
                continue
        else:
            df = pd.read_csv(BytesIO(content_bytes), encoding='utf-8', errors='replace')
    elif file_type in ('xlsx', 'xls'):
        df = pd.read_excel(BytesIO(content_bytes))
    else:
        raise ValueError(f"Unsupported file type: {file_type}")
    
    cleaned_df, changes_json, summary_json = clean_dataframe(df)
    
    # Generate CSV output
    csv_output = cleaned_df.to_csv(index=False)
    
    return csv_output, changes_json, summary_json
`;

async function processFile(file) {
  if (!pyodide) {
    const loaded = await loadPyodide();
    if (!loaded) return;
  }

  processingStatus.textContent = "Reading file...";
  
  // Read file as base64
  const reader = new FileReader();
  const fileContent = new Promise((resolve) => {
    reader.onload = () => {
      const base64 = reader.result.split(",")[1];
      resolve(base64);
    };
    reader.readAsDataURL(file);
  });

  const fileB64 = await fileContent;
  const ext = file.name.split(".").pop().toLowerCase();

  processingStatus.textContent = "Cleaning data with Python...";

  try {
    // Set up Python function
    await pyodide.runPythonAsync(PYTHON_CLEANUP_SCRIPT);
    
    // Call the processing function
    const result = await pyodide.runPythonAsync(`
process_file(
    ${JSON.stringify(fileB64)},
    ${JSON.stringify(ext)},
    ${JSON.stringify(file.name)}
)
    `);

    // Pyodide returns a proxy, we need to convert
    const csvOutput = await result[0];
    const changesJson = await result[1];
    const summaryJson = await result[2];

    const changesData = JSON.parse(changesJson);
    const summaryData = JSON.parse(summaryJson);

    return {
      csv: csvOutput,
      changes: changesData,
      summary: summaryData,
    };
  } catch (err) {
    console.error("Processing error:", err);
    throw err;
  }
}

// ========== File Reading for Preview ==========
function readFileForPreview(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    const ext = file.name.split(".").pop().toLowerCase();

    reader.onload = async (e) => {
      try {
        if (!pyodide) {
          await loadPyodide();
        }
        
        const content = e.target.result;
        
        if (ext === "csv") {
          const text = new TextDecoder("utf-8").decode(content);
          resolve(parseCSV(text));
        } else {
          // For Excel, use pyodide
          const pyodide = await loadPyodide();
          // Convert to base64 for pyodide
          const base64 = btoa(
            new Uint8Array(content).reduce(
              (data, byte) => data + String.fromCharCode(byte),
              ""
            )
          );
          await pyodide.runPythonAsync(`
import pandas as pd
from io import BytesIO
import base64

def read_excel_preview(b64, filename):
    content_bytes = base64.b64decode(b64)
    df = pd.read_excel(BytesIO(content_bytes))
    return df.to_csv(index=False)

_csv_result = read_excel_preview(${JSON.stringify(base64)}, ${JSON.stringify(file.name)})
`);
          const csvText = pyodide.globals.get("_csv_result");
          resolve(parseCSV(csvText));
        }
      } catch (err) {
        reject(err);
      }
    };

    if (ext === "csv") {
      reader.readAsText(file);
    } else {
      reader.readAsArrayBuffer(file);
    }
  });
}

function parseCSV(text) {
  // Simple CSV parser
  const lines = text.trim().split(/\r?\n/);
  if (lines.length === 0) return { headers: [], rows: [] };

  // Detect delimiter
  const firstLine = lines[0];
  const hasTab = firstLine.includes("\t");
  const hasSemicolon = firstLine.includes(";");
  const delimiter = hasTab ? "\t" : hasSemicolon ? ";" : ",";

  const parseLine = (line) => {
    const result = [];
    let current = "";
    let inQuotes = false;

    for (let i = 0; i < line.length; i++) {
      const char = line[i];
      if (char === '"') {
        if (inQuotes && line[i + 1] === '"') {
          current += '"';
          i++;
        } else {
          inQuotes = !inQuotes;
        }
      } else if (char === delimiter && !inQuotes) {
        result.push(current.trim());
        current = "";
      } else {
        current += char;
      }
    }
    result.push(current.trim());
    return result;
  };

  const headers = parseLine(lines[0]);
  const rows = lines.slice(1).map(parseLine);

  return { headers, rows };
}

// ========== UI Functions ==========
function showSection(section) {
  [emailGate, uploadArea, processingArea, resultsArea, outOfCredits].forEach(
    (el) => el && el.classList.add("hidden")
  );
  if (section) section.classList.remove("hidden");
}

function showError(msg) {
  rateMsg.textContent = msg;
  rateMsg.className = "rate-msg error";
  rateMsg.classList.remove("hidden");
}

function showSuccess(msg) {
  rateMsg.textContent = msg;
  rateMsg.className = "rate-msg success";
  rateMsg.classList.remove("hidden");
}

function renderPreview(headers, rows, tableElement, highlightChanges = false, originalHeaders = null) {
  let html = "<thead><tr>";
  headers.forEach((h) => {
    html += `<th>${escapeHtml(h)}</th>`;
  });
  html += "</tr></thead><tbody>";

  const maxRows = Math.min(rows.length, 50);
  for (let i = 0; i < maxRows; i++) {
    html += "<tr>";
    rows[i].forEach((cell, j) => {
      const cellStr = String(cell ?? "");
      html += `<td>${escapeHtml(cellStr)}</td>`;
    });
    html += "</tr>";
  }

  if (rows.length > maxRows) {
    html += `<tr><td colspan="${headers.length}" style="text-align:center;color:var(--text-muted)">... and ${rows.length - maxRows} more rows</td></tr>`;
  }

  html += "</tbody>";
  tableElement.innerHTML = html;
}

function escapeHtml(str) {
  const div = document.createElement("div");
  div.textContent = str;
  return div.innerHTML;
}

function renderStats(summary, changes) {
  statsGrid.innerHTML = `
    <div class="stat-box">
      <div class="stat-number">${summary.original_rows}</div>
      <div class="stat-label">Original Rows</div>
    </div>
    <div class="stat-box">
      <div class="stat-number">${summary.cleaned_rows}</div>
      <div class="stat-label">Cleaned Rows</div>
    </div>
    <div class="stat-box">
      <div class="stat-number">${summary.rows_removed}</div>
      <div class="stat-label">Rows Removed</div>
    </div>
    <div class="stat-box">
      <div class="stat-number">${summary.duplicates_removed}</div>
      <div class="stat-label">Duplicates</div>
    </div>
  `;

  // Changes log
  let changesHtml = "<h4>Changes Applied</h4>";
  if (changes.length === 0) {
    changesHtml += `<p class="muted">No significant issues found — your data was already clean!</p>`;
  } else {
    changes.forEach((c) => {
      const icon = getChangeIcon(c.action);
      changesHtml += `
        <div class="change-item">
          <span class="change-icon">${icon}</span>
          <span>${escapeHtml(c.action)}</span>
          <span class="change-count">×${c.count}</span>
        </div>
      `;
    });
  }
  changesLog.innerHTML = changesHtml;
}

function getChangeIcon(action) {
  const lower = action.toLowerCase();
  if (lower.includes("dup")) return "🗑️";
  if (lower.includes("date")) return "📅";
  if (lower.includes("phone")) return "📞";
  if (lower.includes("address")) return "📍";
  if (lower.includes("casing")) return "🔤";
  if (lower.includes("empty")) return "🧹";
  return "✨";
}

function updateCreditsBadge() {
  const rl = getRateLimit();
  const remaining = RATE_LIMIT - rl.count;
  creditsBadge.textContent = `${remaining} cleanup${remaining !== 1 ? "s" : ""} left today`;
  creditsBadge.classList.toggle("low", remaining <= 1);
}

// ========== Event Handlers ==========
startCleaning.addEventListener("click", () => {
  if (!canUse()) {
    showError("You've reached the device's daily limit.");
    showSection(outOfCredits);
    return;
  }
  sessionStorage.setItem("dc_started", "true");
  showSection(uploadArea);
  updateCreditsBadge();
  showSuccess(`You have ${RATE_LIMIT - getRateLimit().count} cleanups remaining today.`);
});

dropZone.addEventListener("click", () => fileInput.click());

dropZone.addEventListener("dragover", (e) => {
  e.preventDefault();
  dropZone.classList.add("drag-over");
});

dropZone.addEventListener("dragleave", () => {
  dropZone.classList.remove("drag-over");
});

dropZone.addEventListener("drop", (e) => {
  e.preventDefault();
  dropZone.classList.remove("drag-over");
  const files = e.dataTransfer.files;
  if (files.length > 0) handleFileSelect(files[0]);
});

fileInput.addEventListener("change", (e) => {
  if (e.target.files.length > 0) handleFileSelect(e.target.files[0]);
});

function handleFileSelect(file) {
  const maxSize = 10 * 1024 * 1024; // 10MB
  if (file.size > maxSize) {
    alert("File too large. Maximum size is 10MB.");
    return;
  }

  const validTypes = ["csv", "xlsx", "xls"];
  const ext = file.name.split(".").pop().toLowerCase();
  if (!validTypes.includes(ext)) {
    alert("Unsupported file type. Please upload a CSV or Excel file.");
    return;
  }

  currentFile = file;
  fileName.textContent = `${file.name} (${(file.size / 1024).toFixed(1)} KB)`;
  fileInfo.classList.remove("hidden");
}

processBtn.addEventListener("click", async () => {
  if (!currentFile) return;

  if (!sessionStorage.getItem("dc_started") || !canUse()) {
    showSection(outOfCredits);
    return;
  }

  // Show processing
  showSection(processingArea);
  uploadArea.classList.add("hidden");

  try {
    // Also get original data for preview comparison
    processingStatus.textContent = "Reading file for preview...";
    const originalPreview = await readFileForPreview(currentFile);

    // Process the file
    const result = await processFile(currentFile);

    // Parse cleaned CSV for preview
    const cleanedPreview = parseCSV(result.csv);

    originalData = originalPreview;
    cleanedData = cleanedPreview;
    changes = result.changes;

    // Record the use
    const usedCount = recordUse();

    // Show results
    showSection(resultsArea);
    renderStats(result.summary, result.changes);

    // Render "before" preview by default
    renderBeforePreview();

    // Set up download
    const blob = new Blob([result.csv], { type: "text/csv" });
    const url = URL.createObjectURL(blob);
    downloadBtn.href = url;
    downloadBtn.download = `cleaned_${currentFile.name.replace(/\.[^.]+$/, "")}.csv`;

    updateCreditsBadge();

    // Warn if running low
    const remaining = RATE_LIMIT - getRateLimit().count;
    if (remaining <= 0) {
      setTimeout(() => alert("⚠️ You've used all your free cleanups for today! Visit our pricing section to hire us for unlimited access."), 500);
    }
  } catch (err) {
    console.error(err);
    alert("Error processing file: " + err.message);
    showSection(uploadArea);
    uploadArea.classList.remove("hidden");
  }
});

tabBefore.addEventListener("click", () => {
  tabBefore.classList.add("active");
  tabAfter.classList.remove("active");
  renderBeforePreview();
});

tabAfter.addEventListener("click", () => {
  tabAfter.classList.add("active");
  tabBefore.classList.remove("active");
  renderAfterPreview();
});

function renderBeforePreview() {
  if (originalData) {
    renderPreview(originalData.headers, originalData.rows, previewTable);
  }
}

function renderAfterPreview() {
  if (cleanedData) {
    renderPreview(cleanedData.headers, cleanedData.rows, previewTable, true, originalData?.headers);
  }
}

resetBtn.addEventListener("click", () => {
  currentFile = null;
  originalData = null;
  cleanedData = null;
  changes = [];
  fileInfo.classList.add("hidden");
  fileInput.value = "";

  if (sessionStorage.getItem("dc_started") && canUse()) {
    showSection(uploadArea);
  } else {
    showSection(outOfCredits);
  }
});

// ========== Initialize ==========
async function init() {
  // Free quota is per device; no email collection is required.
  if (sessionStorage.getItem("dc_started")) {
    if (canUse()) {
      showSection(uploadArea);
      updateCreditsBadge();
    } else {
      showSection(outOfCredits);
    }
  } else {
    showSection(emailGate);
  }

  // Preload Pyodide in background
  loadPyodide();
}

init();
