# Survey Workbench

A desktop workbench for study coordinators and researchers. Prepare standardized participant folders with fillable PDF questionnaires, prefill initial values, and extract completed survey answers directly into multi-sheet Excel workbooks or CSV files.

---

## 1. Overview & Key Features

- **Collapsible Workspace:** Curator settings stay tucked inside a collapsible drawer, leaving a clean, compact window for day-to-day folder generation and data extraction.
- **Dynamic Window Resizing:** The window automatically expands or shrinks to snugly fit whatever content or questionnaire rows are active.
- **Multi-Survey & Multi-Trial Support:** Configure unlimited questionnaire types per study, with custom repeat copy counts per participant (e.g., Trial 1, Trial 2).
- **Consolidated Batch Processing:** Enter single IDs, paste multiple IDs, or import participant lists from text or CSV files. Generate folders or extract data in a single pass.
- **Cross-Project Field Mapping Import:** Instantly copy friendly column header names from any previously saved study profile via the `File -> Import Field Mapping` submenu (just like Load and Delete Configuration), avoiding manual retyping.
- **Smart Checkbox & Matrix Grouping:** Automatically groups checkbox arrays (like Check1_1, Check1_2) into a single column (Check1) containing the selected option number.
- **Natural Alphanumeric Sorting:** Columns and previews are sorted naturally (Text1, Text2, ..., Text10) instead of alphabetical order (Text1, Text10, Text2).
- **Dedicated Excel Sheets:** In Excel masterfiles, each survey type automatically writes to its own tab, keeping Column A reserved for participant_id.
- **Friendly Column Mapping:** Rename technical PDF identifiers (e.g., Text1 -> Age) using an interactive visual editor.
- **Project Snapshotting:** Save and reload complete project setups (paths, templates, copy counts, and custom column names) to and from `survey_workbench.config`.
- **Pre-Compiled & Cross-Platform:** Standalone Windows executable (.exe) ready to download without setup; native Linux binary easily built with Nuitka.
- **Built-in Interactive Help:** Instant hover tooltips on every control, a "What's This?" interactive click mode, and an in-app viewer for this documentation.

---

## 2. Quick Start Guide

### Step 1: Declare Participant(s)
- **Single participant:** Type the identifier (e.g., P001) into the Participant ID(s) box.
- **Multiple participants:** Check Batch Mode, then paste IDs (one per line or comma-separated) or click Import List... to load a .txt or .csv file.

### Step 2: Configure the Study (First-Time Setup)
Click ▶ Configuration Settings to open the setup panels:
1. **Target folder PATH:** Choose where new participant folders will be created.
2. **Questionnaires:** Enter the number of survey types, click Confirm, then give each survey a name, select its PDF template, and specify how many copies each participant gets.
3. **Source folder PATH:** Choose where completed participant folders are saved.
4. **Masterfile PATH:** Choose your destination .xlsx, .xls, or .csv file.

Tip: Go to **File -> Save Configuration** to save this setup under a profile name so you can reload it anytime without re-entering paths.

### Step 3: Run Actions
- **Generate Participant Folder(s):** Creates the folders, copies and renames the survey PDFs, and opens a review window to check prefilled values.
- **Extract Data to Masterfile:** Scans completed PDFs from the source folder and writes all answers cleanly into your masterfile.

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
2. Click Generate Participant Folder(s).
3. A Batch Preview & Generation table appears, listing all detected form fields across the templates.
4. Any ID fields (such as participant_id or Text13_n) are automatically prefilled with the participant ID.
5. Review or edit values directly in the table, then click Confirm and Generate All.
6. Each participant will receive a dedicated folder containing the prepared PDFs ready for data collection.

### 4.2 Extracting Data to Masterfile
1. Make sure your Source folder points to the completed participant folders.
2. Make sure your Masterfile points to your target spreadsheet.
3. Click Extract Data to Masterfile:
   - Single participant: A 4-column review dialog (Sheet, Field, Name, Data) allows you to inspect and modify extracted values before committing them.
   - Batch mode: All participants are parsed in memory and written into the masterfile in a single session without repeated file-opening delays.
4. Excel exports create dedicated sheets per survey name, ensuring headers and schemas remain cleanly separated.

### 4.3 Renaming Columns & Importing Field Mappings
To replace raw PDF identifiers with friendly column names:
1. Open **File -> Save/Edit Field Mapping**.
2. Type custom titles into the *Masterfile Column Name* column (e.g., change `Demographics_Text1` to `Age`).
3. Click **Save Field Mapping**. The new names will be used on all subsequent data exports.
4. **Importing Mappings from Another Project:** Go to **File -> Import Field Mapping ▶** and choose any saved configuration from the submenu list (just like Load and Delete Configuration). The column renamings will immediately apply to your current session.

### 4.4 Managing Study Profiles
- **Save Configuration:** Go to **File -> Save Configuration** and choose a profile name. All paths, questionnaire templates, counts, and column renamings are saved to `survey_workbench.config`.
- **Load Configuration:** Go to **File -> Load Configuration ▶** and click your saved profile name to restore everything instantly.
- **Delete Configuration:** Go to **File -> Delete Configuration ▶** to remove obsolete profiles.
- **Import Field Mapping:** Go to **File -> Import Field Mapping ▶** to borrow column names from another profile without overwriting your current folders and paths.

---

## 5. Technical Specifications

| Parameter | Specification |
|---|---|
| Supported OS | Windows 10/11, Linux (Ubuntu/Debian/Fedora/Arch), macOS |
| Python Runtime | Python 3.9 through 3.13 |
| GUI Framework | PyQt5 |
| PDF Processing | pypdf (AcroForm, widget annotations, indirect object trees) |
| Masterfile Formats | CSV (.csv), Excel (.xlsx, .xls) |
| Configuration File | `survey_workbench.config` (stored in application root) |
| Standalone Builds | Nuitka (portable .exe on Windows / ELF binary on Linux) |

Note on Operating System Differences:
- CSV masterfiles work 100% natively on both Linux and Windows without requiring any external office suite.
- Excel masterfiles (.xlsx, .xls) automated via xlwings rely on Microsoft Excel (standard on Windows). On Linux systems without MS Excel, selecting .csv as the masterfile format is recommended.

---

## 6. Download & Installation

### 6.1 Windows (Pre-Compiled Executable)
Windows users do not need to install Python, Git, or any C compiler.
The Windows edition is already compiled with Nuitka into a portable standalone executable:
1. Download `survey_workbench.exe` directly from the repository / releases.
2. Place it in your desired working directory and double-click to launch.
   All GUI modules, PDF engines, configuration management, and this documentation are bundled inside.

---

### 6.2 Linux Compilation (Native ELF Binary)
Linux users can compile a standalone native ELF binary directly from source using Nuitka:

```bash
# 1. Install build tools (Debian/Ubuntu)
sudo apt update && sudo apt install build-essential python3-dev -y
pip install -r requirements.txt
pip install nuitka zstandard

# 2. Compile native executable
python3 -m nuitka --standalone --onefile --enable-plugin=pyqt5 --include-data-file=README.md=README.md --output-filename=survey_workbench survey_workbench_v2.2.py

# 3. Run
chmod +x survey_workbench
./survey_workbench