# Survey Workbench v2.1

A comprehensive tool for managing participant folders and extracting survey data from PDF questionnaires with automated field mapping and Excel integration.

## Overview

Survey Workbench is designed to streamline the workflow of:
1. **Generating** standardized participant folders with multiple questionnaire types
2. **Prefilling** PDF forms with default values before participant handover
3. **Extracting** completed survey data directly into Excel/CSV masterfiles
4. **Managing** field mappings for user-friendly column names

### Key Features

- **Dynamic Questionnaire Configuration** — Support unlimited questionnaire types per participant
- **Auto-Generated Field Mapping** — Automatically scans form templates and generates field mappings
- **Non-Participant-Specific Columns** — Multiple participants in a single masterfile without naming conflicts
- **Grouped Checkboxes** — Automatically collapses checkbox arrays (Check1_1, Check1_2, etc.) into single columns
- **Prefill Dialog** — Populate PDF form fields before participant handover with a graphical interface
- **Multiple Export Formats** — Support for CSV, XLS, and XLSX masterfiles
- **Configuration Persistence** — Save and load complete project configurations including field mappings
- **Type-Safe Codebase** — Full Python type hints for reliability and IDE support

## System Requirements

- **OS:** Windows 10/11
- **Python:** 3.10+ (if running from source)
- **Excel:** Microsoft Excel (for XLS/XLSX export)
- **RAM:** 2GB minimum
- **Storage:** Sufficient space for participant folders and questionnaire PDFs

## Installation

### Using the Executable

1. Download `survey_workbench_v2.1.exe` from the `dist/` folder
2. Run the executable — no installation needed
3. The first launch will create a `config.ini` file in the same directory

### From Source

```bash
# Clone or extract the repository
cd survey_workbench - v2.1

# Install dependencies
pip install -r requirements.txt

# Run the application
python survey_workbench_v2.1.py
```

**Dependencies:**
- PyQt5 (GUI framework)
- xlwings (Excel integration)
- pypdf (PDF processing)

## Quick Start

### 1. Generate Participant Folders

1. **Configure Questionnaires**
   - Enter participant ID (e.g., `P_001`)
   - Select target folder for participant directories
   - Add questionnaire types (name, PDF template, copy count)

2. **Generate**
   - Click "Generate Participant Folder"
   - System creates organized folder structure
   - Prefill dialog appears for optional value population

3. **Prefill (Optional)**
   - Edit field values in the dialog
   - Check/uncheck options for checkboxes
   - Click "Confirm" to save prefill values to PDFs

### 2. Extract Survey Data

1. **Configure Extraction**
   - Select source folder (contains completed participant folders)
   - Select masterfile (CSV, XLS, or XLSX)
   - Enter participant ID to extract

2. **Extract**
   - Click "Extract Data to Masterfile"
   - System automatically detects file format
   - Data is appended as new row with proper field mapping

3. **Verify**
   - Check masterfile for new row with mapped column names
   - Checkpoint values appear as option numbers (1, 2, 3, etc.)

### 3. Manage Field Mappings

1. **Edit Field Mapping**
   - Click "File → Save/Edit Field Mapping"
   - System auto-scans form templates and shows detected fields
   - Edit "Name" column to set user-friendly column headers
   - Click "Save" to persist to project configuration

2. **Save Configuration**
   - Click "File → Save Configuration"
   - Enter configuration name
   - Settings, questionnaires, AND field mapping are saved together

3. **Load Configuration**
   - Click "File → Load Configuration" (from menu)
   - Select previous configuration
   - Everything restores: questionnaires, paths, AND field mapping

## Data Organization

### Column Naming

- **Previous (v2.0):** `P_001_Demographics_Age` (participant-specific)
- **Current (v2.1):** `Age` (mapped friendly name) or `Demographics_Age` (system name)
- **Benefit:** Multiple participants can coexist in one masterfile

### Checkbox Handling

Checkboxes in PDFs use naming convention:
- `Check1_1`, `Check1_2`, `Check1_3` (individual options)
- Grouped as: `Check1` (single column)
- Value: option number that was checked (e.g., `1`, `2`, or `3`)
- System translates to/from PDF values (`/Yes`, `/Off`, `/On`) automatically

### Example Data Row

| participant_id | Age | Gender | Mental_Demand | Score |
|---|---|---|---|---|
| P_001 | 28 | 1 | 7.5 | 68 |
| P_002 | 35 | 2 | 5.2 | 72 |

## Configuration Files

### config.ini

Stores all project configurations. Each `[section]` is a named configuration:

```ini
[MyProject]
target_path = C:\Studies\MyProject\Participants
source_path = C:\Studies\MyProject\Completed
excel_path = C:\Studies\MyProject\masterfile.xlsx
quest_count = 3
quest_0_name = Demographics
quest_0_path = C:\Templates\demographics.pdf
quest_0_count = 1
field_mapping_json = {"Demographics_Age": "Age", "Check1": "Gender"}
```

Each configuration saves:
- Questionnaire setup (names, paths, counts)
- Source/target/excel paths
- **Field mapping** (system names → friendly names)

## Workflow Example

```
1. Create configuration "MyStudy_2026"
   ├─ Add questionnaires: demographics.pdf, survey-tlx.pdf
   ├─ Set target folder: C:\Studies\MyStudy\Participants
   └─ Save configuration

2. Generate participant folders
   ├─ ID: P_001
   ├─ Creates: P_001/
   │   ├─ P_001_demographics.pdf
   │   └─ P_001_survey-tlx.pdf
   └─ Prefill optional values

3. Participant completes questionnaires
   └─ Returns completed PDFs

4. Edit field mapping (optional)
   ├─ "Demographics_Age" → "Age"
   ├─ "Check1" → "Gender"
   └─ Save to configuration

5. Extract data
   ├─ Source: C:\Studies\MyStudy\Completed\P_001/
   ├─ Masterfile: C:\Studies\MyStudy\masterfile.xlsx
   └─ Appends row with mapped column names
```

## Menu Reference

### File Menu
- **Save Configuration** — Save current setup with field mappings
- **Load Configuration** — Load previous project setup
- **Delete Configuration** — Remove saved configuration
- **Save/Edit Field Mapping** — Edit field-to-column-name mappings (auto-scans templates)
- **Exit** — Close the application

### Generation Section
- **Generate Participant Folder** — Create folders with questionnaires and optional prefill

### Extraction Section
- **Extract Data to Masterfile** — Append completed survey data to masterfile

## Technical Details

### Version: 2.1 (June 2026)

**Changes from v2.0:**
- Field mapping auto-generated from form templates
- Non-participant-specific column names
- Integrated field mapping with project configuration
- Removed redundant load/delete field mapping operations
- Enhanced type hints and code quality

### Technology Stack

- **Python:** 3.13.7
- **GUI:** PyQt5 5.15.11
- **PDF Processing:** pypdf 4.0+
- **Excel Integration:** xlwings 0.33.20
- **Packaging:** PyInstaller 6.14.2

### Supported File Formats

- **PDF:** Standard AcroForm fields (text, checkboxes, radio buttons)
- **CSV:** UTF-8 encoded, comma-delimited
- **Excel:** XLS (2003), XLSX (2007+)

## Troubleshooting

### "No fields found in configured templates"
- Verify template PDF paths are correct
- Ensure PDFs have form fields (AcroForm structure)
- Check file permissions

### "Field mapping does not match configured forms"
- Forms have changed since configuration was created
- Click "Save/Edit Field Mapping" to regenerate from current templates
- Review and adjust friendly names, then save

### Excel file not updating
- Ensure masterfile is not open in Excel
- Verify path is correct and file is accessible
- Check that sheet named "Data" exists (or first sheet is used)

### Checkbox values not appearing
- Verify checkbox field names follow pattern: `Check{N}_{option}`
- Ensure option numbers are sequential (1, 2, 3...)
- Check that at least one checkbox is checked in the PDF

## Documentation

- **USER_MANUAL.pdf** — Comprehensive user guide with screenshots and step-by-step instructions
- **README.md** — This file, quick reference and overview

## Support

For issues, feature requests, or questions:
1. Check USER_MANUAL.pdf for detailed guidance
2. Review configuration in config.ini
3. Verify file paths and permissions
4. Consult "Troubleshooting" section above

## License

This project is licensed under the **MIT License** — See [LICENSE](LICENSE) file for details.

**Copyright © 2026 Cagatay Özcan Jagiello Gutt (Original Creator and Developer)**

The MIT License permits free use, modification, and distribution while preserving this copyright attribution. Your creative work and engineering effort are permanently credited as the original creator of this tool.

---

**Survey Workbench v2.1** | June 2026 | Open-Source PDF Form Extraction Tool
