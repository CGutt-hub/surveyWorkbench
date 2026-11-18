# Survey Workbench

A desktop workbench for study coordinators and researchers. Prepare standardized participant folders with fillable PDF questionnaires, prefill initial values, and extract completed survey answers directly into Apache Parquet, Excel workbooks, CSV, TSV, or JSON masterfiles.

---

## 1. Overview & Key Features

- **Multi-Format Masterfile Exports:**
  - **Microsoft Excel (`.xlsx`, `.xls`):** Dedicated tabs per questionnaire type, placing each participant at their preassigned row (`x + 1`) and keeping schemas clean and isolated.
  - **Apache Parquet (`.parquet`):** High-speed columnar storage optimized for data science pipelines (Python, R, Polars, DuckDB, Julia).
  - **Delimited Text (`.csv`, `.tsv`, `.tab`):** Standard comma- or tab-separated tables with zero third-party dependencies.
  - **Structured JSON (`.json`):** Human-readable, structured data arrays.
- **Canonical Grammar Consolidation:** Automatically groups checkbox arrays (like `Check1_1`, `Check1_2`) into a single column (`Check1`) containing the chosen option identifier `n`.
- **Ignore / NOT Operator (`_`):** Any field prefixed with an underscore is omitted from masterfile exports.
- **Condition Ordering (`con` Field):** Supports counterbalanced / mixed-order repeated surveys via a dedicated `con` (or `_con`) PDF field and **File -> Save/Edit Condition Ordering** (Numerical as standard, or custom mappings & meanings saved to `survey_workbench.config`).
- **Filename-Driven Sheet Routing:** Automatically derives the masterfile sheet name directly from the PDF filename (e.g., `WaS3_001_nasa_1.pdf` -> `nasa`, `WaS3_002_you_are_an_idiot_34.pdf` -> `you_are_an_idiot`) and appends numbered copies (`nasa_1`, `nasa_2`, `nasa_3`, ...) horizontally in the same participant row.
- **Preassigned Excel Row Placement (`x + 1`):** In Excel sheets, participant `x` (e.g., `1`, `P001`, `WaS3_001`) is always placed in row `x + 1` (row 2 for participant 1, row 3 for participant 2, etc.) rather than simply appending to the bottom.
- **Consolidated Batch Processing & Interactive Selection:** Enter single IDs, paste multiple IDs, or import participant lists from text or CSV files. During extraction, an interactive checklist presents **all** participants in the batch (showing `New` vs `Already in Masterfile (will overwrite)`) so you can untick anyone you do not want to export.
- **Unified Batch Preview Table:** Both folder generation and masterfile extraction open a single, consolidated table displaying all participants at once. Inspect and edit friendly header names and questionnaire answers before committing any changes.
- **Safe Staging:** Prefill values and survey folders are validated in memory before any directory is modified on disk. If you cancel, nothing is touched.
- **Collapsible Workspace & Dynamic Resizing:** Curator settings stay tucked inside a collapsible drawer, leaving a clean, compact window that automatically resizes to fit active study configurations.
- **Cross-Project Field Mapping Import:** Instantly copy friendly column header names from any previously saved study profile via the `File -> Import Field Mapping` submenu.
- **Natural Alphanumeric Sorting:** Columns and previews are sorted naturally (`Text1`, `Text2`, ..., `Text10`) instead of alphabetical order (`Text1`, `Text10`, `Text2`).
- **Project Snapshotting:** Save and reload complete project setups (paths, templates, copy counts, custom column names, and condition orderings) to and from `survey_workbench.config`.
- **Pre-Compiled & Cross-Platform:** Standalone Windows executable (.exe) ready to download without setup; native Linux binary easily built with Nuitka.
- **Built-in Interactive Help:** Hover tooltips, a "What's This?" interactive click mode, and an in-app viewer for this documentation.

---

## 2. Canonical Form Field Grammar

PDF form fields must adhere to the following specification:

`(_){Check, Text}<i>(_n)` and optionally `(_)con(_n)`

### Breakdown of Grammar Elements:

| Element | Description | Masterfile & Workflow Behavior |
| :--- | :--- | :--- |
| **`_`** | **NOT Operator** (optional leading underscore) | If present, the content of this field **does NOT get exported into the masterfile**. Ideal for administrative cover-page notes, visual participant ID labels, print-only metadata, or routing-only `_con` fields. |
| **`{Check, Text}`** | **Entry Type** | Declares the input type. `Text` fields capture verbatim strings. `Check` fields represent boolean, binary, or nominal selection controls. |
| **`<i>`** | **Field / Question Index** | Index number of the form question or condition. If option buttons belong to the same question, `i` stays constant while `n` changes. |
| **`_n`** | **Option Identifier** (optional) | Specifies the choice identifier. In the masterfile, the cell shows **`n`** for column `Check<i>` if this option is chosen to be true. |
| **`(_)con(_n)`** | **Condition Field** (optional, case-insensitive) | Declares the experimental condition of a repeated survey (text field, dropdown, radio group, or `con_<opt>` checkboxes) so mixed-order PDFs align cleanly in the masterfile row. Use `_con` to route without exporting a `con` column, or `con` to both route and export `<prefix>con`. |

### Concrete Examples:

1. **Nominal / Likert Scales (`Check1_1` to `Check1_5`):**
   - Options are named `Check1_1`, `Check1_2`, `Check1_3`, `Check1_4`, `Check1_5`.
   - If the participant selects the 3rd box (`Check1_3`), the masterfile stores `3` in the column **`Check1`**.
   - No individual `Check1_1`, `Check1_2`, etc. columns clutter your masterfile.
2. **Binary Conditions (`Check2_0` and `Check2_1`):**
   - Canonically, `0` represents False/No and `1` represents True/Yes.
   - If the participant checks `Check2_0`, the masterfile records **`0`** in column **`Check2`**.
   - If the participant checks `Check2_1`, the masterfile records **`1`** in column **`Check2`**.
3. **Standalone Checkbox (`Check3`):**
   - A single boolean checkbox without an `_n` suffix.
   - If checked, the cell records **`1`**; if unchecked, it records **`0`**.
4. **Ignored Cover Fields (`_Text13` or `_Check5_1`):**
   - Prefixed with the leading `_` NOT operator.
   - Can be prefilled with the participant ID during folder generation so that it is printed on the physical survey form, but is completely omitted during extraction into the masterfile.
5. **Counterbalanced Condition Routing (`con` / `_con` / `con_sf`):**
   - Suppose Participant 1 receives `WaS3_001_nasa_1.pdf` (`con = 2` or `con = sf`) and `WaS3_001_nasa_2.pdf` (`con = 1` or `con = bl`).
   - With numerical `con` (or `bl` = Order `1`, `sf` = Order `2` configured in **File -> Save/Edit Condition Ordering**), `nasa_2.pdf` is routed into `nasa_1_*` (first block) and `nasa_1.pdf` is routed into `nasa_2_*` (second block), keeping every participant aligned by condition in the masterfile.
   - If prefixed with `_` (`_con`), the field is used strictly for ordering and omitted from the exported columns.

### Zero-Redundancy in Pop-Up Tables:
All pop-up tables (Save/Edit Field Mapping, Batch Generation Prefill, and Extraction Review) automatically consolidate all `Check<i>_n` fields into a single `Check<i>` entry. Defining names or mappings for every individual option `n` is completely eliminated.

---

## 3. Quick Start Guide

### Step 1: Declare Participant(s)
- **Single participant:** Type the identifier (e.g., `WaS3_001` or `P001`) into the Participant ID(s) box.
- **Multiple participants:** Check **Batch Mode**, then paste IDs (one per line or comma-separated) or click **Import List...** to load a `.txt` or `.csv` list.

### Step 2: Configure the Study (First-Time Setup)
Click **Configuration Settings** to open the setup panels:
1. **Target folder PATH:** Choose where new participant folders will be created.
2. **Questionnaires:** Enter the number of survey types, click Confirm, then give each survey a name, select its PDF template, and specify how many copies each participant gets.
3. **Source folder PATH:** Choose where completed participant folders are saved.
4. **Masterfile PATH:** Choose your destination file (`.parquet`, `.xlsx`, `.csv`, `.tsv`, or `.json`).
5. **Condition Ordering (optional, for mixed conditions):** Go to **File -> Save/Edit Condition Ordering**. **Numerical** is enabled by default (using numbers in the `con` field directly). If your PDFs use text codes like `sf` or `bl`, add their order (`1, 2, 3...`) and meaning (or click **Scan Source PDFs for 'con'**).

Tip: Go to **File -> Save Configuration** to save this setup (including field mappings and condition orderings) under a profile name so you can reload it anytime.

### Step 3: Run Actions
- **Generate Participant Folder(s):** Checks for existing folders, opens the prefill preview table, and generates the participant survey copies.
- **Extract Data to Masterfile:**
  1. Opens the **Select Participants to Export** table showing all participants in your batch and their masterfile status (`New` vs `Already in Masterfile`), allowing you to untick anyone you want to skip.
  2. If any unmapped text `con` codes (such as `sf`) are encountered, prompts you to assign their order and meaning.
  3. Opens the consolidated extraction preview table where you can inspect/edit column names and values, then writes all selected participants into their sheets (at row `x + 1` in Excel).

---

## 4. Interactive Help in the App

- **Hover:** Move your mouse over a control to see its tooltip.
- **What's This? Mode:** Press Shift + F1 (or go to HELP -> What's This?). Click any control to read an explanation of its function.
- **Read Manual:** Go to HELP -> User Manual to open this documentation directly inside the program.

---

## 5. Masterfile Formats, Sheet Naming & Row Placement

- **Sheet / File Base Naming:** Sheet names are derived directly from the PDF filename after stripping the participant ID prefix and any trailing copy indices (e.g., `WaS3_001_nasa_1.pdf`, `WaS3_001_nasa_2.pdf` -> sheet `nasa`; `WaS3_002_you_are_an_idiot_34.pdf` -> sheet `you_are_an_idiot`).
- **Multi-Copy Horizontal Appending:** Multiple PDFs sharing the same base name (`nasa_1`, `nasa_2`, `nasa_3`, ...) are appended horizontally in the **same row** on the same sheet—never as separate sheets—ordered by `con` (when present) or copy number.
- **Excel (`.xlsx`, `.xls`):**
  - Each questionnaire base name is matched case-insensitively to an existing worksheet (or creates a new one if absent).
  - Participant `x` (determined by the trailing/single number in the ID, e.g., `1`, `P001`, `WaS3_001` -> `1`) is written directly into their preassigned row **`x + 1`** (leaving skipped participant rows empty instead of shifting rows up). Non-numeric IDs are appended after the last occupied row.
- **Parquet (`.parquet`):** Automatically creates partitioned parquet tables named `<masterfile>_<survey_name>.parquet`.
- **CSV / TSV (`.csv`, `.tsv`):** Automatically writes partitioned files named `<masterfile>_<survey_name>.csv` / `.tsv`.
- **JSON (`.json`):** Automatically writes arrays named `<masterfile>_<survey_name>.json`.

---

## 6. Technical Specifications & Dependencies

- **Supported OS:** Windows 10/11, Linux, macOS
- **Python Runtime:** Python 3.9 through 3.13
- **GUI Framework:** PyQt5
- **PDF Processing:** pypdf (AcroForm, widget annotations, indirect object trees)
- **Masterfile Formats:** Apache Parquet (`.parquet`), Excel (`.xlsx`, `.xls`), CSV (`.csv`), TSV (`.tsv`), JSON (`.json`)
- **Parquet Engine:** `pyarrow` (recommended) or `pandas`
- **Configuration File:** `survey_workbench.config` (stored in application root)
- **Standalone Builds:** Nuitka (portable .exe on Windows / ELF binary on Linux)

To install all dependencies:
```bash
pip install pyarrow xlwings pypdf PyQt5
```

---

## 7. Compilation with Nuitka

### Windows Build Command:
```cmd
python -m nuitka --standalone --onefile --enable-plugin=pyqt5 --include-data-file=README.md=README.md --windows-console-mode=disable --output-filename=survey_workbench.exe survey_workbench_v2.2.py
```

### Linux Build Command:
```bash
python3 -m nuitka --standalone --onefile --enable-plugin=pyqt5 --include-data-file=README.md=README.md --output-filename=survey_workbench survey_workbench_v2.2.py
```