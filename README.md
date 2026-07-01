# 🌐 Browser History Analyzer v2.0

> A forensics-grade, PyQt6-powered GUI for extracting, previewing, and exporting **Edge / Chrome browsing and download history** — with all timestamps in **UTC** and full download-to-history correlation.

---

## ✨ Features

| Feature | Details |
|---|---|
| 🔍 **On-screen Preview** | Full browse history + download list rendered in-app before any export |
| 🕐 **UTC Timestamps** | Every timestamp converted from Chromium's WebKit epoch (µs since 1601-01-01) to UTC — no local time zone pollution |
| 🔗 **Download Correlation** | Each history entry shows what files were downloaded from that URL — exact URL match + optional hostname match |
| 📊 **XLSX Export** | One-click export to Excel — separate `History` and `Downloads` sheets, all UTC headers |
| 🌙 **Dark / Light Theme** | iOS-style toggle — matches the MultiOSINT tool suite aesthetic |
| 🔎 **Live Filtering** | Real-time search bars on both History and Downloads tabs; "Downloads Only" toggle |
| 🗂️ **Multi-file Support** | Select multiple `History` DB files at once; export all in one run |
| 🔒 **Lock-safe DB access** | Copies the live SQLite DB to a temp file before opening — works even while the browser is running |

---

## 🖥️ Screenshots

> **Dark Mode — History Tab**
> History rows with green ✅ indicators for entries with correlated downloads

> **Dark Mode — Downloads Tab**
> Color-coded State column: `Complete` (green) · `In Progress` (orange) · `Cancelled / Interrupted` (red)

---

## 📋 Requirements

```
Python 3.9+
PyQt6
openpyxl
```

Install dependencies:

```bash
pip install PyQt6 openpyxl
```

---

## 🚀 Usage

```bash
python BrowsingHistoryGUI.pyw
```

### Step-by-step

1. **Select History File(s)** — Click `📂 Select History File(s)…` and pick one or more Chromium `History` SQLite files
2. **Set Output Folder** — Browse to your export destination (defaults to Desktop)
3. **Set Subject Label** — Optional prefix for the output filename (e.g. `CaseID`, `username`, `hostname`)
4. **Preview Data** — Click `🔍 Preview Data` to load and display history + downloads in the table view
5. **Export** — Once preview is loaded, click `📊 Export to XLSX` — the Export button is locked until a preview has been run

---

## 📁 Where is the History file?

| Browser | Default path |
|---|---|
| **Microsoft Edge** | `%LOCALAPPDATA%\Microsoft\Edge\User Data\Default\History` |
| **Google Chrome** | `%LOCALAPPDATA%\Google\Chrome\User Data\Default\History` |
| **Chromium** | `%LOCALAPPDATA%\Chromium\User Data\Default\History` |

> ⚠️ Close the browser before selecting the file, or the tool will automatically copy the locked DB to a temp location.

---

## 📊 Output: Excel Workbook

### Sheet 1 — `History`

| Column | Description |
|---|---|
| `last_visit_time_utc` | Last visited timestamp in UTC |
| `url` | Full page URL |
| `title` | Page title |
| `downloaded` | `Yes` / `No` — whether a file was downloaded from this URL |
| `downloaded_files` | Semicolon-separated list of downloaded file names |

### Sheet 2 — `Downloads`

| Column | Description |
|---|---|
| `id` | Download record ID |
| `file_name` | Downloaded file name |
| `source_url` | URL the download was initiated from |
| `final_url` | Final redirect URL (from `downloads_url_chains`) |
| `url_chain` | Full redirect chain |
| `start_time_utc` | Download start in UTC |
| `end_time_utc` | Download completion in UTC |
| `state` | `Complete` / `Cancelled` / `Interrupted` / `In Progress` |
| `size` | File size (auto-formatted: B / KB / MB / GB) |

---

## ⚙️ Options

### Hostname Matching
When enabled (default: ✅), the tool also correlates downloads to history entries by **hostname** — not just exact URL. Useful when download links differ slightly from the page URL (CDN redirects, tracking parameters, etc.).

---

## 🔧 Architecture

```
BrowsingHistoryGUI.pyw
├── THEME + build_stylesheet()       iOS-style dark/light CSS for PyQt6
├── webkit_ts_to_utc()               Chromium µs epoch → UTC string
├── load_data()                      SQLite → dicts (safe temp-copy, url_chains)
├── export_xlsx()                    openpyxl workbook builder
├── LoadWorker(QThread)              Background loading, emits finished/error
└── HistoryApp(QMainWindow)
    ├── _build_header()              Logo + theme toggle
    ├── _build_sidebar()             File picker, output config, actions
    └── _build_main() / QTabWidget
        ├── History tab              UTC Time · URL · Title · Downloaded · Files
        └── Downloads tab            ID · File · Start/End UTC · State · Size · URLs
```

---

## 🛡️ DFIR / OSINT Notes

- All timestamps are converted **directly from the WebKit epoch** (`1601-01-01 00:00:00 UTC`) — no system time zone involved
- The `downloads_url_chains` table is parsed to reconstruct the **full redirect chain** — critical for tracing CDN/tracking redirects back to the originating domain
- Download state codes: `0` = In Progress, `1` = Complete, `2` = Cancelled, `3/4` = Interrupted
- Error log written to `%TEMP%\history_export_error.log` on any DB or export failure

---

## 🧩 Part of the MultiOSINT Tool Suite

This tool shares the same iOS-inspired dark/light UI framework as [`MultiOSINTv11`](../MultiOSINTv11.pyw) — consistent theme, sidebar layout, and keyboard-friendly design.

---

## 📝 Changelog

### v2.0 — 2026-06-03
- Complete rewrite from **Tkinter → PyQt6**
- On-screen preview before export (History + Downloads tabs)
- All timestamps converted to **UTC** (was local time in v1)
- Download ↔ History correlation with hostname-match option
- Live filter bars with "Downloads Only" toggle
- Multi-file selection and batch export
- iOS dark/light theme with toggle button
- Background loading thread — UI stays responsive on large DBs

### v1.0
- Tkinter-based GUI
- Local timestamp display
- Direct XLSX export without preview

---

## 📄 License

MIT — use freely, attribution appreciated.
