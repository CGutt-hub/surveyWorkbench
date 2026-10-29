# Survey Workbench

A desktop workbench for study coordinators and researchers. Prepare standardized participant folders with fillable PDF questionnaires, prefill initial values, and extract completed survey answers directly into Apache Parquet, Excel workbooks, CSV, TSV, or JSON masterfiles.

---

## 1. Overview & Key Features

- **Multi-Format Masterfile Exports:**
  - **Apache Parquet (`.parquet`):** High-speed columnar storage optimized for data science pipelines (Python, R, Polars, DuckDB, Julia).
  - **Microsoft Excel (`.xlsx`, `.xls`):** Dedicated tabs per questionnaire type, keeping schemas clean and isolated.
  - **Delimited Text (`.csv`, `.tsv`, `.tab`):** Standard comma- or tab-separated tables with zero third-party dependencies.
  - **Structured JSON (`.json`):** Human-readable, structured data arrays.
- **Consolidated Batch Processing:** Enter single IDs, paste multiple IDs, or import participant lists from text or CSV files. Generate folders or extract data in a single pass.
- **Interactive Overwrite Resolution:** If folders or masterfile records already exist, an interactive checklist lets you select exactly which participants to overwrite (with one-click *Select All* / *Deselect All*), avoiding accidental data loss.
- **Unified Batch Preview Table:** Both folder generation and masterfile extraction open a single, consolidated table displaying all participants at once. Inspect and edit friendly header names and questionnaire answers before committing any changes.
- **Safe Staging:** Prefill values and survey folders are validated in memory before any directory is modified on disk. If you cancel, nothing is touched.
- **Collapsible Workspace & Dynamic Resizing:** Curator settings stay tucked inside a collapsible drawer, leaving a clean, compact window that automatically resizes to fit active study configurations.
- **Cross-Project Field Mapping Import:** Instantly copy friendly column header names from any previously saved study profile via the `File -> Import Field Mapping` submenu.
- **Smart Checkbox & Matrix Grouping:** Automatically groups checkbox arrays (like Check1_1, Check1_2) into a single column (Check1) containing the selected option number.
- **Natural Alphanumeric Sorting:** Columns and previews are sorted naturally (Text1, Text2, ..., Text10) instead of alphabetical order (Text1, Text10, Text2).
- **Project Snapshotting:** Save and reload complete project setups (paths, templates, copy counts, and custom column names) to and from `survey_workbench.config`.
- **Pre-Compiled & Cross-Platform:** Standalone Windows executable (.exe) ready to download without setup; native Linux binary easily built with Nuitka.
- **Built-in Interactive Help:** Instant hover tooltips on every control, a "What's This?" interactive click mode, and an in-app viewer for this documentation.

---

## 2. Quick Start Guide

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

## 3. Interactive Help in the App

If you are ever unsure what a control does:
- **Hover:** Move your mouse over any button or text box to see a brief tooltip.
- **What's This? Mode:** Press Shift + F1 (or go to HELP -> What's This?). Your mouse cursor turns into a question mark. Click any button, label, or input field to read a detailed explanation of its function.
- **Read Manual:** Go to HELP -> User Manual to open this documentation directly inside the program.

---

## 4. Step-by-Step Instructions

### 4.1 Generating Folders & Prefilling Data
1. Enter your participant ID(s).
2. Click **Generate Participant Folder(s)**.
3. **Overwrite Check:** If any participant folders already exist, a resolution window prompts you to tick which folders to overwrite (or skip).
4. **Batch Prefill Table:** A consolidated preview window displays all detected form fields across all participants. Any ID fields (such as `participant_id` or `Text13_n`) are automatically prefilled.
5. Review or edit values directly in the table, then click **Confirm and Generate All**.
6. Each participant receives a dedicated folder containing the prepared PDFs ready for data collection.

### 4.2 Extracting Data to Masterfile
1. Ensure your **Source folder** points to the completed participant folders and **Masterfile** points to your target file (`.parquet`, `.xlsx`, `.csv`, `.tsv`, or `.json`).
2. Click **Extract Data to Masterfile**:
   - **Overwrite Check:** If any participant already exists in your masterfile, a pop-up checklist lets you choose which participant rows to update.
   - **Batch Extraction Review Table:** A consolidated table (`Participant ID`, `Sheet`, `Field`, `Column Name`, `Data`) opens with all participant data loaded into memory. You can inspect or correct values directly.
3. Click **Confirm and Extract All**. The masterfile is written in a single pass without repeatedly launching office applications.

### 4.3 Output File Partitioning
- **Excel (`.xlsx`, `.xls`):** Each questionnaire type is saved as a dedicated worksheet / tab.
- **Parquet (`.parquet`):** Automatically creates partitioned parquet tables named `<masterfile>_<survey_name>.parquet`.
- **CSV / TSV (`.csv`, `.tsv`):** Automatically writes partitioned files named `<masterfile>_<survey_name>.csv` / `.tsv`.
- **JSON (`.json`):** Automatically writes arrays named `<masterfile>_<survey_name>.json`.

### 4.4 Renaming Columns & Importing Field Mappings
To replace raw PDF identifiers with friendly column names:
1. Open **File -> Save/Edit Field Mapping**.
2. Type custom titles into the *Masterfile Column Name* column (e.g., change `Demographics_Text1` to `Age`).
3. Click **Save Field Mapping**. The new names will be used on all subsequent data exports.
4. **Importing Mappings from Another Project:** Go to **File -> Import Field Mapping ▶** and choose any saved configuration from the submenu list to borrow column names immediately.

---

## 5. Technical Specifications & Dependencies

- **Supported OS:** Windows 10/11, Linux (Ubuntu/Debian/Fedora/Arch), macOS
- **Python Runtime:** Python 3.9 through 3.13
- **GUI Framework:** PyQt5
- **PDF Processing:** pypdf (AcroForm, widget annotations, indirect object trees)
- **Masterfile Formats:** Apache Parquet (`.parquet`), Excel (`.xlsx`, `.xls`), CSV (`.csv`), TSV (`.tsv`), JSON (`.json`)
- **Parquet Engine:** `pyarrow` (recommended) or `pandas`
- **Configuration File:** `survey_workbench.config` (stored in application root)
- **Standalone Builds:** Nuitka (portable .exe on Windows / ELF binary on Linux)

To install all supported format engines:
```bash
pip install pyarrow xlwings pypdf PyQt5
```

---

## 6. Compilation with Nuitka

### Windows Build Command:
```cmd
python -m nuitka --standalone --onefile --enable-plugin=pyqt5 --include-data-file=README.md=README.md --windows-console-mode=disable --output-filename=survey_workbench.exe survey_workbench_v2.2.py
```

### Linux Build Command:
```bash
python3 -m nuitka --standalone --onefile --enable-plugin=pyqt5 --include-data-file=README.md=README.md --output-filename=survey_workbench survey_workbench_v2.2.py
```