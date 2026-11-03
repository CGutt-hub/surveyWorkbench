# Survey Workbench

A desktop workbench for study coordinators and researchers. Prepare standardized participant folders with fillable PDF questionnaires, prefill initial values, and extract completed survey answers directly into Apache Parquet, Excel workbooks, CSV, TSV, or JSON masterfiles.

---

## 1. Overview & Key Features

- **Multi-Format Masterfile Exports:**
  - **Apache Parquet (`.parquet`):** High-speed columnar storage optimized for data science pipelines (Python, R, Polars, DuckDB, Julia).
  - **Microsoft Excel (`.xlsx`, `.xls`):** Dedicated tabs per questionnaire type, keeping schemas clean and isolated.
  - **Delimited Text (`.csv`, `.tsv`, `.tab`):** Standard comma- or tab-separated tables with zero third-party dependencies.
  - **Structured JSON (`.json`):** Human-readable, structured data arrays.
- **Canonical Grammar Consolidation:** Automatically groups checkbox arrays (like `Check1_1`, `Check1_2`) into a single column (`Check1`) containing the chosen option identifier `n`.
- **Ignore / NOT Operator (`_`):** Any field prefixed with an underscore is omitted from masterfile exports.
- **Consolidated Batch Processing:** Enter single IDs, paste multiple IDs, or import participant lists from text or CSV files. Generate folders or extract data in a single pass.
- **Interactive Overwrite Resolution:** If folders or masterfile records already exist, an interactive checklist lets you select exactly which participants to overwrite (with one-click *Select All* / *Deselect All*), avoiding accidental data loss.
- **Unified Batch Preview Table:** Both folder generation and masterfile extraction open a single, consolidated table displaying all participants at once. Inspect and edit friendly header names and questionnaire answers before committing any changes.
- **Safe Staging:** Prefill values and survey folders are validated in memory before any directory is modified on disk. If you cancel, nothing is touched.
- **Collapsible Workspace & Dynamic Resizing:** Curator settings stay tucked inside a collapsible drawer, leaving a clean, compact window that automatically resizes to fit active study configurations.
- **Cross-Project Field Mapping Import:** Instantly copy friendly column header names from any previously saved study profile via the `File -> Import Field Mapping` submenu.
- **Natural Alphanumeric Sorting:** Columns and previews are sorted naturally (`Text1`, `Text2`, ..., `Text10`) instead of alphabetical order (`Text1`, `Text10`, `Text2`).
- **Project Snapshotting:** Save and reload complete project setups (paths, templates, copy counts, and custom column names) to and from `survey_workbench.config`.
- **Pre-Compiled & Cross-Platform:** Standalone Windows executable (.exe) ready to download without setup; native Linux binary easily built with Nuitka.
- **Built-in Interactive Help:** Instant hover tooltips on every control, a "What's This?" interactive click mode, and an in-app viewer for this documentation.

---

## 2. Canonical Form Field Grammar

PDF form fields must adhere to the following mathematical specification:

$$\mathbf{(\_)\{Check, Text\}<i>(\_n)}$$

### Breakdown of Grammar Elements:

| Element | Description | Masterfile & Workflow Behavior |
| :--- | :--- | :--- |
| **`_`** | **NOT Operator** (optional leading underscore) | If present, the content of this field **does NOT get exported into the masterfile**. It simply gets ignored during extraction. Ideal for administrative cover-page notes, visual participant ID labels, or print-only metadata. |
| **`{Check, Text}`** | **Entry Type** | Declares the input type. `Text` fields capture verbatim strings. `Check` fields represent boolean, binary, or nominal selection controls. |
| **`<i>`** | **Field / Question Index** | Index number of the form question or condition. If option buttons belong to the same question, $i$ stays constant while $n$ changes. |
| **`_n`** | **Option Identifier** (optional) | Specifies the choice identifier. In the masterfile, the cell shows **`n`** for column `Check<i>` if this option is chosen to be true. |

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

### Zero-Redundancy in Pop-Up Tables:
All pop-up tables (Save/Edit Field Mapping, Batch Generation Prefill, and Extraction Review) automatically consolidate all `Check<i>_n` fields into a single `Check<i>` entry. Defining names or mappings for every individual option `n` is completely eliminated.

---

## 3. Quick Start Guide

### Step 1: Declare Participant(s)
- **Single participant:** Type the identifier (e.g., `P001`) into the Participant ID(s) box.
- **Multiple participants:** Check **Batch Mode**, then paste IDs (one per line or comma-separated) or click **Import List...** to load a `.txt` or `.csv` list.

### Step 2: Configure the Study (First-Time Setup)
Click ▶ **Configuration Settings** to open the setup panels:
1. **Target folder PATH:** Choose where new participant folders will be created.
2. **Questionnaires:** Enter the number of survey types, click Confirm, then give each survey a name, select its PDF template, and specify how many copies each participant gets.
3. **Source folder PATH:** Choose where completed participant folders are saved.
4. **Masterfile PATH:** Choose your destination file (`.parquet`, `.xlsx`, `.csv`, `.tsv`, or `.json`).

Tip: Go to **File -> Save Configuration** to save this setup under a profile name so you can reload it anytime without re-entering paths.

### Step 3: Run Actions
- **Generate Participant Folder(s):** Checks for existing folders, opens the prefill preview table, and generates the participant survey copies.
- **Extract Data to Masterfile:** Checks for existing spreadsheet records, opens the consolidated extraction review table, and commits answers into your masterfile in one single pass.

---

## 4. Interactive Help in the App

- **Hover:** Move your mouse over any button or text box to see a brief tooltip.
- **What's This? Mode:** Press Shift + F1 (or go to HELP -> What's This?). Click any control to read a detailed explanation of its function.
- **Read Manual:** Go to HELP -> User Manual to open this documentation directly inside the program.

---

## 5. Masterfile Formats & Output Partitioning

- **Excel (`.xlsx`, `.xls`):** Each questionnaire type is saved as a dedicated worksheet / tab.
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