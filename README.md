# Survey Workbench v2.0

> A comprehensive participant data management system for survey research with dynamic questionnaire configuration and batch processing capabilities.

[![Python](https://img.shields.io/badge/Python-3.8+-blue.svg)](https://www.python.org/)
[![PyQt5](https://img.shields.io/badge/GUI-PyQt5-green.svg)](https://www.riverbankcomputing.com/software/pyqt/)
[![License](https://img.shields.io/badge/License-Internal_Use-red.svg)]()

## Overview

Survey Workbench is a desktop application designed to streamline the management of participant folders and extraction of survey data from questionnaires. Built with PyQt5, it provides an intuitive graphical interface for researchers and data managers to efficiently organize and process survey data.

### Key Features

- **🔧 Dynamic Questionnaire Configuration**: Support for unlimited questionnaire types per participant with flexible template management
- **📦 Batch Processing**: Generate and extract data for multiple participants simultaneously
- **📥 Participant Import**: Import participant lists from .txt or .csv files
- **📋 Template Bundles**: Create and reuse questionnaire configuration bundles across projects
- **🔍 Duplicate Detection**: Automatic masterfile checking (supports CSV and Excel formats) to prevent duplicate entries
- **✅ Data Completeness Verification**: Validate all required data before extraction
- **👁️ Preview Dialog**: Review extracted data before finalizing
- **📊 Missing Data Report**: Generate quality control reports for incomplete data
- **💾 Configuration Management**: Save, load, and manage multiple configurations with an intuitive submenu interface
- **❓ Interactive Help System**: Built-in tooltips and "What's This?" mode for user assistance
- **📑 Auto-Format Detection**: Automatically detect masterfile format (CSV, XLS, XLSX)

## System Requirements

- **Operating System**: Windows 10/11, macOS 10.14+, or Linux
- **Python**: 3.8 or higher
- **Microsoft Excel**: Required for Excel file operations (via xlwings)
- **Memory**: 4GB RAM minimum (8GB recommended for large datasets)
- **Storage**: 100MB free space minimum

## Installation

### Prerequisites

Ensure you have Python 3.8+ installed on your system. You can download it from [python.org](https://www.python.org/downloads/).

### Install Dependencies

```bash
# Clone the repository
git clone https://github.com/CGutt-hub/surveyWorkbench.git
cd surveyWorkbench

# Install required Python packages
pip install PyQt5 xlwings configparser
```

### Additional Setup

For Excel integration (xlwings), you may need to install the Excel add-in:

```bash
xlwings addin install
```

## Quick Start

### Running the Application

```bash
python survey_workbench.py
```

Or run the compiled executable (if available):

```bash
./survey_workbench_v2.0  # On Linux/macOS
survey_workbench_v2.0.exe  # On Windows
```

### Basic Workflow

1. **Configure Questionnaires**: Set up your questionnaire templates and target folders
2. **Generate Participant Folders**: Create participant-specific folders with questionnaire templates
3. **Fill Out Questionnaires**: Have participants complete their questionnaires
4. **Extract Data**: Collect and consolidate data from completed questionnaires into a masterfile

## Usage

### Generate Participant Folders

1. Select template files for each questionnaire type
2. Specify the target folder where participant folders will be created
3. Enter participant IDs (manual entry or import from file)
4. Click "Generate Participant Folder" to create the folder structure

**Batch Mode**: Enable batch mode to process multiple participants at once by importing a list from .txt or .csv files.

### Extract Data

1. Select the source folder containing participant folders
2. Choose the masterfile (CSV or Excel) where data will be extracted
3. Configure questionnaire-specific extraction settings:
   - Excel sheet names
   - Column filters
   - Multiple questionnaire copies
4. Click "Extract Data" to consolidate participant data

**Features**:
- **Duplicate Detection**: Automatically checks if participant data already exists in the masterfile
- **Data Completeness Check**: Verifies all required questionnaires are present before extraction
- **Preview Dialog**: Review data before final extraction
- **Missing Data Report**: Generate reports for participants with incomplete data

### Configuration Management

Save and load configurations to quickly switch between different project setups:

- **Save Configuration**: Store your current questionnaire setup and settings
- **Load Configuration**: Quickly restore a previously saved configuration
- **Delete Configuration**: Remove outdated configurations
- **Recent Configurations**: Access recently used configurations from the menu

### Template Bundles

Create reusable template bundles for standardized project setups:

1. Configure all questionnaires and settings
2. Select "Save Template Bundle" from the menu
3. Load the bundle in future projects to instantly apply the same configuration

## File Structure

```
surveyWorkbench/
├── survey_workbench.py      # Main application source code
├── survey_workbench.spec    # PyInstaller build specification
├── config.ini               # Configuration storage file
├── USER_MANUAL.pdf          # Comprehensive user manual
├── USER_MANUAL.tex          # LaTeX source for user manual
└── README.md                # This file
```

## Technology Stack

- **GUI Framework**: PyQt5 - Cross-platform graphical user interface
- **Excel Integration**: xlwings - Python library for Excel automation
- **Configuration**: ConfigParser - INI file handling for settings persistence
- **Build Tool**: PyInstaller - Executable packaging (see survey_workbench.spec)
- **Type Hints**: Full type annotation support for better code maintainability

## Documentation

For detailed documentation, including screenshots and step-by-step guides, please refer to the [USER_MANUAL.pdf](USER_MANUAL.pdf) included in this repository.

## Troubleshooting

### Common Issues

- **Excel not found**: Ensure Microsoft Excel is installed and xlwings is properly configured
- **Configuration not saving**: Check write permissions for config.ini file
- **Import errors**: Verify all dependencies are installed with `pip list`
- **Template files not copying**: Ensure source template files exist and have read permissions

For more troubleshooting tips, consult the USER_MANUAL.pdf.

## Version History

### Version 2.0 (February 2026)
- Dynamic questionnaire support with unlimited types
- Enhanced batch processing capabilities
- Template bundle system
- Improved duplicate detection
- Data completeness verification
- Preview dialog for data extraction
- Missing data reporting
- Interactive help system

### Version 1.0 (April 2024)
- Initial release
- Basic participant folder generation
- Simple data extraction

## Author

**Cagatay Gutt**
- Created: April 15, 2024
- Last Updated: February 4, 2026

## License

This software is for internal use only. All rights reserved.

## Support

For questions, issues, or feature requests, please contact the project maintainer or refer to the comprehensive USER_MANUAL.pdf for detailed guidance.

---

*Survey Workbench - Streamlining survey data management for research excellence.*