# Changelog

All notable changes to the **ExcelClone** project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

---

## [2.0.0] - 2026-11-13

### Highlights
- **100% Free & Permissive Open Source Migration:** Fully replaced proprietary Syncfusion Essential Studio (`SfSpreadsheet` and `XlsIO`) with **ReoGrid V3 Community Edition (`unvell.ReoGridWPF.dll` v3.3.0, MIT License)** and **ClosedXML (v0.104.2, MIT License)**.
- **Zero License Keys & Popups:** Permanently eliminated all trial validation prompts, registration popups, and evaluation sheet watermarks.
- **83% Smaller Distribution Footprint:** Portable ZIP size reduced from ~27 MB to **4.55 MB**.
- **Comprehensive Test Suite:** 34 / 34 passing automated tests covering core calculation, end-to-end user workflows, external XLSX compatibility, and visual UI operations.

### Added
- **Find and Replace Studio (`FindReplaceDialog`):** Full multi-criteria cell searching and replacing with Match Case, Match Entire Cell, Find Next, Replace, Replace All, and cross-sheet / workbook-wide scanning (`Ctrl+F`, `Ctrl+H`).
- **Interactive Visual Charts:** Integrated native vector charts into worksheets supporting **Column Charts**, **Bar Charts**, **Line Charts**, and **Pie Charts** linked to active cell data via `WorksheetChartDataSource`.
- **PivotTable Builder (`PivotDialog`):** Multi-dimensional aggregation engine supporting Row and Column groupings with `Sum`, `Count`, `Average`, `Max`, and `Min` summaries and linked PivotChart generation.
- **Direct Printing:** Added `BtnPrint_Click` using `CreatePrintSession().Print()` for direct Windows page printing.
- **Interactive Viewport Zooming:** Added Zoom In, Zoom Out, and Zoom 100% Reset via native ReoGrid scale controls.
- **Automatic Session Recovery:** Startup detection and prompt to restore uncommitted work from auto-saved snapshots (`AutoRecovery_Current.xlsx`).
- **Data Transformation Studio Integration:** Connected `DataTransformationWindow` to live active worksheet data with direct reload back to sheet.
- **Enhanced Keyboard Shortcuts:** Added `Ctrl+Z` (Undo), `Ctrl+Y` (Redo), `Ctrl+A` (Select All), `Ctrl+F` (Find), `Ctrl+H` (Replace), `Ctrl+P` (Print), `Ctrl+O` (Open), `Ctrl+N` (New).

---

## [1.5.0-RC] - 2026-11-13

### Highlights
- First public Release Candidate of **ExcelClone**.
- Completely standalone, offline-capable desktop spreadsheet for Windows 10 & 11.
- All cloud backends, mandatory logins, and online requirements removed.
- 31 passing automated tests covering formula calculation, cross-sheet references, file round-trips, styles, and external OpenXML compatibility.

### Added
- **Full Ribbon Formatting Wiring:** Wired active selection styling directly to `SfSpreadsheet` (Font Family, Font Size, Bold, Italic, Underline, Alignment, Fill Color, and Borders).
- **Global Keyboard Shortcuts:**
  - `Ctrl+B` — Toggle Bold on selected range.
  - `Ctrl+I` — Toggle Italic on selected range.
  - `Ctrl+U` — Toggle Underline on selected range.
  - `Ctrl+S` — Save active workbook.
  - `F1` — Open Keyboard Shortcuts reference dialog.
- **Draw Tab InkCanvas Overlay:** Integrated freehand ink canvas overlay on `MainWindow.xaml` with Black Pen, Highlighter, Clear Canvas, and clean toggling back to standard cell selection via "Select (Grid)".
- **Formula Auditing:**
  - *Trace Precedents:* Inspects cell formulas and highlights precedent cell coordinates.
  - *Trace Dependents:* Scans active worksheet to discover all cells dependent on the selected cell.
- **Name Manager:**
  - *Define Name:* Dialog to assign friendly names to cell ranges.
  - *Name Manager:* Dialog to view, list, and verify all workbook named ranges.
- **Data Tools:**
  - *AutoFilter:* Attach/detach column filters to active range.
  - *Sort A to Z / Z to A:* Quick alphanumeric sorting across selected ranges.
  - *Remove Duplicates:* Detects and purges duplicate rows from selected ranges while preserving unique records and headers.
  - *Data Validation:* User list-based data validation with prompt boxes.
- **Cell Comments:**
  - *New Comment / Edit Comment:* Attach text notes to specific cells.
  - *Delete Comment:* Remove comments from cells.
  - *Show Comments:* Dialog enumerating all worksheet comments with cell addresses.
- **Protection Tools:**
  - *Protect / Unprotect Sheet:* Password protection locking worksheet content.
  - *Protect / Unprotect Workbook:* Protects workbook structure against adding or deleting sheets.
- **License Management:**
  - Early license key registration in `App.OnStartup()` before UI loads.
  - Settings dialog field to enter and persist Syncfusion Community or Commercial license key.
  - Environment variable (`SYNCFUSION_LICENSE_KEY`) and file-based (`syncfusion_license.txt`) discovery.
- **Automated Verification:** Expanded test suite from 16 to 31 tests, adding end-to-end user workflows A–F and external OpenXML compatibility tests.

### Changed
- Refactored `SfSpreadsheetRibbon` data binding from `DataContext` to `SfSpreadsheet` property so ribbon buttons remain active and bind directly to selected grid cells.
- Standardized default UI theme to `Office2019DarkGray`.
- Title bar format streamlined to `{WorkbookName} - Excel`.
- Consolidated distribution into portable directory (`Distribution/ExcelClone_Portable/`) and clean PowerShell installation and uninstallation scripts.

### Security & Privacy
- Completely removed all external Supabase / cloud authentication dependencies.
- Local document parsing and saving with zero telemetry of spreadsheet data.
- VBA `.xlsm` macro storage retained for lossless round-tripping without executing untrusted script payloads.

---

## [1.4.0] - 2026-11-12
### Added
- Goal Seek numerical solver.
- Descriptive Statistics and Linear Regression analysis tools.
- Data Transformation Studio.
- CSV and delimited text file import wizard.
- Starter workbook templates.

---

## [1.0.0] - 2026-11-10
### Initial Release
- Initial prototype and core spreadsheet engine integration using Syncfusion SfSpreadsheet on .NET 10.

