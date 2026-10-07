---
name: ewa-dashboard
description: >-
  Generates a customer-ready SAP EWA service review dashboard from .doc/.docx EWA reports. Handles Word 2003 XML .doc (native SAP ECS format), OOXML .docx with SAP purl.oclc.org namespace, standard OOXML .docx. Parser v3.4 auto-extracts customer name and cloud provider (AWS/GCP/Azure) from Hardware Configuration. Self-contained HTML dashboard: 12 tabs, RAG indicators, Chart.js charts, PDF export, Action Tracker CSV/Excel, 4-EWA trend panel, 12-month Capacity trends — DB size history (Excel or me.sap.com seed) and HANA memory trend (localStorage auto-accumulation). Works on Windows and macOS. Use when: analyze EWA, prepare EWA dashboard, EWA service review, Early Watch Alert report, EWA analysis, prepare for customer review, service review dashboard, EWA presentation, generate dashboard from EWA, upload EWA report.
metadata:
  author: Ravikiran Chikati
  version: 2.3.0
  tags: ewa sap dashboard service-review ecs hana reporting capacity memory-trend local-storage cross-platform
---

# EWA Service Review Dashboard — v2.3

Generate a customer-ready SAP EWA dashboard from a `.doc` or `.docx` EWA report. Handles all three SAP EWA file formats automatically:

| Format | Extension | Description |
|---|---|---|
| Word 2003 XML | `.doc` | Native SAP ECS EWA format — plain XML, no conversion needed |
| SAP OOXML | `.docx` | ZIP-based OOXML with SAP purl.oclc.org namespace — auto-fixed |
| Standard OOXML | `.docx` | Standard ZIP-based OOXML — opened directly |

> **Old binary `.doc`** (OLE2, Word 97-2003 binary) is the only format that cannot be parsed. If you encounter this, open in Word and save as `.docx`.

## When to Activate

Activate when the user:
- Asks to "analyze EWA", "prepare EWA dashboard", "generate EWA report"
- Mentions "Early Watch Alert", "service review dashboard", "EWA presentation"
- Provides a `.doc` or `.docx` file and asks for analysis

---

## Step 1 — Identify the Input File

If the user has not provided a file path, ask:
> "Please provide the path to your EWA file (`.doc` or `.docx`)."
>
> **Windows example:** `C:\Reports\EWA_PRD_2026.doc`
> **macOS example:** `/Users/yourname/Documents/EWA_PRD_2026.doc`

If the user provides a file name without a full path, check the working directory using `ls` to locate it.

Accept `.doc` and `.docx`. The parser auto-detects the format — no renaming or pre-processing required.

If the file is an old binary `.doc` (Word 97-2003), tell the user:
> "This appears to be an old binary Word document. Please open it in Word and save as `.docx` (File → Save As → Word Document), then re-run."

---

## Step 2 — Install Dependencies

Run in a **separate terminal call** before executing any script:

```bash
pip install python-docx lxml
```

`lxml` is used for the Word 2003 XML parser (SAP `.doc` format) and for recovery mode on malformed XML. `python-docx` handles OOXML `.docx` files.

Do NOT combine the install with the script execution.

---

## Step 3 — Verify Template

Check that `ewa_template.html` exists in the working directory:

```bash
ls .
```

If `ewa_template.html` is NOT there, it needs to be regenerated once from the reference HTML (see skill notes below).

---

## Step 4 — Parse the EWA Report

The `/skills/ewa-dashboard/` path is resolved automatically by Joule Desktop on both Windows and macOS.

```bash
python "/skills/ewa-dashboard/scripts/parse_ewa.py" \
  --input "PATH_TO_EWA.doc" \
  --output "ewa_data.json" \
  --indent 2
```

Replace `PATH_TO_EWA.doc` with the actual file path.

> **macOS note:** If `python` is not found, use `python3` instead:
> ```bash
> python3 "/skills/ewa-dashboard/scripts/parse_ewa.py" --input "..." --output "ewa_data.json" --indent 2
> ```

**What the parser extracts automatically (v3.4):**
- `customer` — from EWA cover page (e.g. `Cardinal Health Inc.`)
- `cloud` — from Hardware Configuration section (`AWS`, `GCP`, `Azure`, `On-Prem`)
  - Detects: `Amazon Web Services`, `Amazon EC2`, `Google Compute Engine`, `Google Cloud`, `Microsoft Azure`
  - Uses hardware-context tracking to avoid false positives

---

## Step 5 — Generate the Dashboard

```bash
python "/skills/ewa-dashboard/scripts/generate_dashboard.py" \
  --input "ewa_data.json" \
  --output "ewa-dashboard.html"
```

> **macOS note:** Use `python3` if `python` is not found.

---

## Step 6 — Report to User

Tell the user:

> ✅ Dashboard ready: **ewa-dashboard.html** is in your working folder. Open it in any browser.
>
> **Header shows:** `<CustomerName> · System <SID> · <Product> · <Cloud>` — all extracted automatically from the EWA report.
>
> **Tabs:** Executive Summary, System Health, Performance, Capacity, Database, Security & Config, Operations, SQL Analysis, Top Risks, Action Tracker, Technical Debt, Management Summary
>
> **Standard features:** Traffic light RAG indicators, Chart.js visualizations, per-section PDF export, Action Tracker with CSV/Excel export, Upload panel for trend analysis across up to 4 EWA reports.
>
> **Capacity Tab — v2.2+ features:**
> - **DB Size on Disk — 12-Month History**: Three seeding methods:
>   1. Open the direct link embedded in the EWA report ("click here" in Database Growth section) → `me.sap.com/ewa/dashboard/dbSizeDetail/<SID>_<installation>_<session>`
>   2. Export from `me.sap.com` as Excel → share the file → convert serial dates with: `from datetime import datetime, timedelta; dt = datetime(1899,12,30) + timedelta(days=serial)`
>   3. Run the browser console script on the loaded page (see Notes below)
> - **HANA Memory — 12-Month Trend**: Auto-accumulates from first EWA upload via browser `localStorage` key `EWA_memHistory_<SID>`. Zero manual input.

---

## Notes

### Cross-platform compatibility
- The `/skills/ewa-dashboard/` path is a Joule Desktop virtual path — resolved automatically on both **Windows** and **macOS**.
- On **macOS**, Python may be invoked as `python3`. If `python` gives a "command not found" error, substitute `python3` in all commands.
- The `parse_ewa.py` script is fully cross-platform (no OS-specific code, pure Python stdlib + python-docx + lxml).
- `pip install python-docx lxml` works identically on Windows and macOS.

### Template
- `ewa_template.html` must be in the working directory. It is generated once from the reference HTML (`EWA_Dashboard_PS4_v11.html`) and reused for all subsequent EWA reports.

### Parser output fields (v3.4 additions)
| Field | Source | Example |
|---|---|---|
| `customer` | EWA cover page, before "Session No." | `Cardinal Health Inc.` |
| `cloud` | Hardware Configuration → Manufacturer column | `GCP`, `AWS`, `Azure` |

Both fields appear in the generated dashboard header: `<customer> · System <sid> · <product> · <cloud>`.

### DB size history — Excel seeding method
When the user provides a `Database_Size_on_Disk.xlsx` file:
1. Read date (Excel serial) and size (GB) columns
2. Convert serials: `from datetime import datetime, timedelta; date = datetime(1899,12,30) + timedelta(days=int(serial))`
3. Build JSON array: `[{"date": "DD.MM.YYYY", "size": XX.XX}, ...]`
4. Inject as `dbSizeHistory` in the dashboard `DEFAULT_DATA`

### DB size history — console script method
Open the EWA DB size detail URL from me.sap.com, then run in browser DevTools → Console:
```javascript
(function() {
  const rows = [];
  document.querySelectorAll('table tr').forEach(tr => {
    const cells = [...tr.querySelectorAll('td, th')].map(c => c.innerText.trim()).filter(Boolean);
    if (cells.length) rows.push(cells.join(' | '));
  });
  console.log(rows.join('\n'));
})();
```

### HANA version guidance
Always refer to SAP Note 2196476 for ECS-supported HANA database releases.

### localStorage persistence
HANA memory trend data is stored under `EWA_memHistory_<SID>` in the browser. Persists between sessions. Each EWA upload appends one new weekly reading automatically via `saveMemReading()` / `loadMemHistory()`. No user action beyond regular EWA upload.