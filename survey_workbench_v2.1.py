# -*- coding: utf-8 -*-
"""
Survey Workbench - Enhanced Participant Data Management System

Author: Cagatay Gutt
Created: April 15, 2024
Last Updated: February 4, 2026

Features:
- Dynamic questionnaire configuration with unlimited questionnaire types
- Batch participant processing (generation and extraction)
- Participant list import from .txt/.csv files
- Template bundle system for reusable questionnaire configurations
- Duplicate detection with masterfile checking (CSV and Excel)
- Data completeness verification before extraction
- Preview dialog before data extraction
- Missing data report for quality control
- Interactive help system with tooltips and What's This mode
- Configuration save/load with submenu interface
- Auto-detection of masterfile format (CSV, XLS, XLSX)
"""

import sys
import os
import shutil
import csv
import json
import re
from typing import Optional, List, cast, Dict, Any
from PyQt5.QtWidgets import (QApplication, QMainWindow, QPushButton, 
                             QGridLayout, QWidget, QLineEdit, QLabel,
                             QFileDialog, QMessageBox, QAction, QStatusBar,
                             QScrollArea, QVBoxLayout, QFrame,
                             QMenu, QDialog, QWhatsThis, QTextEdit, QTableWidget,
                             QTableWidgetItem, QHeaderView, QCheckBox, QInputDialog)
from configparser import ConfigParser
from PyQt5.QtCore import pyqtSignal, Qt
import xlwings as xlw  # type: ignore[import]
from pypdf import PdfReader
from pypdf import PdfWriter
from pypdf.generic import NameObject, create_string_object

FIELD_NAME_PATTERN = re.compile(r'^(Text|Check)(\d+)(?:_(\d+))?(?:_n)?$')
TEXT_EXPORT_PATTERN = re.compile(r'^Text(\d+)$')

# PyQt5 has incomplete type stubs - some overloads contain Unknown types
# This is a known limitation of PyQt5's typing support

class SaveConfigWindow(QDialog):
    signal = pyqtSignal(str)
    
    def __init__(self, parent: Optional[QWidget] = None) -> None:
        super().__init__(parent)
        self.setWindowTitle("Save Configuration")
        self.setModal(True)
        layout: QGridLayout = QGridLayout()
        self.title3: QLabel = QLabel("Set configuration name:")
        layout.addWidget(self.title3, 0, 0)
        self.configset: QLineEdit = QLineEdit()
        layout.addWidget(self.configset, 0, 1)
        self.save: QPushButton = QPushButton('Save configuration')
        layout.addWidget(self.save, 1, 1)
        self.setLayout(layout)
        
        self.save.clicked.connect(self.troubleshoot)
        
    def troubleshoot(self) -> None:
        self.name = self.configset.text()
        self.config = ConfigParser()
        self.config.read('config.ini')
        (self.areYouSure if self.config.has_section(self.name) else self.confirm)()
    
    def areYouSure(self) -> None:
        dlg = QMessageBox(self)
        dlg.setWindowTitle("Override Configuration")
        dlg.setIcon(QMessageBox.Icon.Question)
        dlg.setStandardButtons(cast(QMessageBox.StandardButton, QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No))
        dlg.setText("This configuration exists already.\nDo you want to override this configuration permanently?")
        
        if dlg.exec() == QMessageBox.StandardButton.Yes:
            self.config.remove_section(self.name)
            with open('config.ini', 'w') as f:
                self.config.write(f)
            self.confirm()
        
    def confirm(self) -> None:
        self.signal.emit(self.name)
        self.close()

class QuestionnaireRow:
    """Represents a single questionnaire configuration row"""
    def __init__(self, index: int) -> None:
        self.index: int = index
        self.name_label: QLabel = QLabel(f"Survey {index + 1} name:")
        self.name_edit: QLineEdit = QLineEdit()
        self.path_label: QLabel = QLabel("Template PATH:")
        self.path_display: QLabel = QLabel()
        self.path_button: QPushButton = QPushButton("Search")
        self.count_label: QLabel = QLabel("Number of copies:")
        self.count_edit: QLineEdit = QLineEdit("1")
        self.template_path: str = ""
        
    def get_widgets(self) -> List[QWidget]:
        """Returns list of all widgets for this row"""
        return [
            self.name_label, self.name_edit,
            self.path_label, self.path_display, self.path_button,
            self.count_label, self.count_edit
        ]

class PrefillPresetsDialog(QDialog):
    """Dialog for editing PDF field prefill default value presets."""

    def __init__(self, parent: Optional[QWidget] = None, pdf_fields: Optional[List[str]] = None) -> None:
        super().__init__(parent)
        self.setWindowTitle("Prefill Presets")
        self.setModal(True)
        self.resize(600, 400)
        self.presets: Dict[str, str] = {}

        config = ConfigParser()
        config.read('config.ini')
        existing: Dict[str, str] = dict(config.items('PrefillDefaults')) if config.has_section('PrefillDefaults') else {}
        fields = pdf_fields or []

        self._table = QTableWidget()
        self._table.setColumnCount(2)
        self._table.setHorizontalHeaderLabels(["Field", "Default Value"])
        header = self._table.horizontalHeader()
        if header is not None:
            header.setStretchLastSection(True)
            header.setSectionResizeMode(0, QHeaderView.ResizeMode.ResizeToContents)
        self._table.setRowCount(len(fields))
        for i, field in enumerate(fields):
            fi = QTableWidgetItem(field)
            fi.setFlags(Qt.ItemFlag.ItemIsEnabled)
            self._table.setItem(i, 0, fi)
            self._table.setItem(i, 1, QTableWidgetItem(existing.get(field, '')))

        save_btn = QPushButton("Save")
        cancel_btn = QPushButton("Cancel")
        save_btn.clicked.connect(self._collect_and_accept)
        cancel_btn.clicked.connect(self.reject)
        btn_layout = QVBoxLayout()
        btn_layout.addWidget(save_btn)
        btn_layout.addWidget(cancel_btn)
        layout = QVBoxLayout()
        layout.addWidget(QLabel("Set default prefill values for PDF fields:"))
        layout.addWidget(self._table)
        layout.addLayout(btn_layout)
        self.setLayout(layout)

    def _collect_and_accept(self) -> None:
        self.presets = {}
        for i in range(self._table.rowCount()):
            fi = self._table.item(i, 0)
            vi = self._table.item(i, 1)
            if fi is not None and vi is not None:
                self.presets[fi.text()] = vi.text()
        self.accept()


class MainWindow(QMainWindow):
    def __init__(self, parent: Optional[QWidget] = None) -> None:
        super().__init__(parent)
        self.setWindowTitle("Survey Workbench")
        
        # Initialize questionnaire rows list
        self.questionnaire_rows: List[QuestionnaireRow] = []
        self.recent_configs_menu: Optional[QMenu] = None
        self.delete_configs_menu: Optional[QMenu] = None
        
        # Main widget and layout
        main_widget = QWidget()
        main_layout = QVBoxLayout()
        
        # Menu setup
        button_action: QAction = QAction("&Save Configuration", self)
        button_action.triggered.connect(self.showSaveConfigWindow)
        
        button_action3: QAction = QAction("User Manual", self)
        button_action3.triggered.connect(self.onMyToolBarButtonClick3)
        
        whats_this_action: QAction = QAction("What's This? (Click for Help)", self)
        whats_this_action.triggered.connect(self.enterWhatsThisMode)
        whats_this_action.setShortcut("Shift+F1")
        
        self.setStatusBar(QStatusBar(self))

        menubar = self.menuBar()
        assert menubar is not None
        file_menu = menubar.addMenu("&File")
        assert file_menu is not None
        file_menu.addAction(button_action)  # type: ignore[misc]
        
        # Load Configuration submenu
        self.recent_configs_menu = file_menu.addMenu("&Load Configuration")
        assert self.recent_configs_menu is not None
        
        # Delete Configuration submenu
        self.delete_configs_menu = file_menu.addMenu("&Delete Configuration")
        assert self.delete_configs_menu is not None
        
        file_menu.aboutToShow.connect(self.updateRecentConfigsMenu)  # type: ignore[misc]

        edit_mapping_action: QAction = QAction("&Save/Edit Field Mapping", self)
        edit_mapping_action.triggered.connect(self.editFieldMappingFromExtraction)
        edit_mapping_action.setToolTip("Save or edit the field name mapping for the active configuration")
        edit_mapping_action.setToolTip("Edit the field name mapping for the active configuration")
        file_menu.addAction(edit_mapping_action)  # type: ignore[misc]
        
        help_menu = menubar.addMenu("&HELP")
        assert help_menu is not None
        help_menu.addAction(whats_this_action)  # type: ignore[misc]
        help_menu.addAction(button_action3)  # type: ignore[misc]
        
        # ========== GENERATION SECTION ==========
        gen_frame: QFrame = QFrame()
        gen_frame.setFrameStyle(QFrame.Shape.Box | QFrame.Shadow.Raised)
        gen_layout = QGridLayout()
        
        gen_title = QLabel("Generate Participant Folder")
        gen_title.setStyleSheet("font-weight: bold; font-size: 14px;")
        
        part_label = QLabel("Participant ID:")
        part_label.setStyleSheet("font-weight: bold")
        self.nameset = QLineEdit()
        self.nameset.setToolTip("Enter the unique participant identifier (e.g., P001)")
        self.nameset.setWhatsThis("The Participant ID is a unique identifier for each participant. This will be used as the folder name and in the data extraction. Example: P001")
        
        # Batch mode checkbox
        self.batch_mode_gen = QCheckBox("Batch Mode (multiple IDs)")
        self.batch_mode_gen.setToolTip("Enable to generate folders for multiple participants at once")
        self.batch_mode_gen.toggled.connect(self.toggleBatchModeGeneration)
        
        # Batch text area (hidden by default)
        self.batch_text_gen = QTextEdit()
        self.batch_text_gen.setPlaceholderText("Enter participant IDs (one per line or comma-separated):\nP001\nP002\nP003")
        self.batch_text_gen.setVisible(False)
        self.batch_text_gen.setMaximumHeight(100)
        self.batch_text_gen.setToolTip("Enter multiple participant IDs, one per line or comma-separated")
        
        target_label = QLabel("Target folder PATH:")
        self.target_pathset = QLabel()
        self.target_pathset.setToolTip("Location where participant folders will be created")
        target_pathbut = QPushButton("Search")
        target_pathbut.clicked.connect(self.select_target_folder)
        target_pathbut.setToolTip("Click to browse and select the target folder")
        target_pathbut.setWhatsThis("Select the folder where you want to create participant folders. Each participant will get a subfolder containing their questionnaire templates.")
        
        # Questionnaire configuration
        quest_count_label = QLabel("Number of questionnaire types:")
        self.quest_count_edit = QLineEdit()
        self.quest_count_edit.setToolTip("How many different questionnaire types will this participant complete?")
        self.quest_count_edit.setWhatsThis("Enter the number of different questionnaire types (e.g., if the participant will complete NASA-TLX at 8 different times, enter 8). After confirming, you can specify the template file and number of copies for each type.")
        quest_count_button = QPushButton("Confirm")
        quest_count_button.clicked.connect(self.create_questionnaire_rows)
        quest_count_button.setToolTip("Click to create configuration rows for each questionnaire type")
        
        gen_layout.addWidget(gen_title, 0, 0, 1, 3)
        gen_layout.addWidget(part_label, 1, 0)
        gen_layout.addWidget(self.nameset, 1, 1)
        gen_layout.addWidget(self.batch_mode_gen, 1, 2)
        
        # Import participant list button
        import_list_btn_gen = QPushButton("Import List...")
        import_list_btn_gen.clicked.connect(lambda: self.importParticipantList(self.batch_text_gen))
        import_list_btn_gen.setToolTip("Import participant IDs from .txt or .csv file")
        import_list_btn_gen.setWhatsThis("Import a text or CSV file containing participant IDs (one per line or comma-separated) to populate the batch mode text area.")
        gen_layout.addWidget(import_list_btn_gen, 2, 2)
        
        gen_layout.addWidget(self.batch_text_gen, 2, 0, 1, 2)
        gen_layout.addWidget(target_label, 3, 0)
        gen_layout.addWidget(self.target_pathset, 3, 1)
        gen_layout.addWidget(target_pathbut, 3, 2)
        gen_layout.addWidget(quest_count_label, 4, 0)
        gen_layout.addWidget(self.quest_count_edit, 4, 1)
        gen_layout.addWidget(quest_count_button, 4, 2)
        
        # Template Bundle buttons
        save_bundle_btn = QPushButton("Save Template Bundle")
        save_bundle_btn.clicked.connect(self.saveTemplateBundle)
        save_bundle_btn.setToolTip("Save current questionnaire configuration as a reusable template bundle")
        save_bundle_btn.setWhatsThis("Saves the current questionnaire templates, names, and copy counts as a reusable bundle that can be loaded later for other participants.")
        gen_layout.addWidget(save_bundle_btn, 5, 0)
        
        load_bundle_btn = QPushButton("Load Template Bundle")
        load_bundle_btn.clicked.connect(self.loadTemplateBundle)
        load_bundle_btn.setToolTip("Load a previously saved template bundle")
        load_bundle_btn.setWhatsThis("Loads a saved template bundle to quickly configure questionnaires with the same templates, names, and copy counts.")
        gen_layout.addWidget(load_bundle_btn, 5, 1)
        
        gen_frame.setLayout(gen_layout)
        main_layout.addWidget(gen_frame)
        
        # Scrollable area for questionnaire rows
        self.scroll_area = QScrollArea()
        self.scroll_area.setWidgetResizable(True)
        self.scroll_widget = QWidget()
        self.scroll_layout = QGridLayout()
        self.scroll_widget.setLayout(self.scroll_layout)
        self.scroll_area.setWidget(self.scroll_widget)
        self.scroll_area.setVisible(False)
        main_layout.addWidget(self.scroll_area)
        
        # Generate button
        self.generate_button = QPushButton("Generate Participant Folder")
        self.generate_button.setStyleSheet("font-weight: bold; padding: 10px;")
        self.generate_button.clicked.connect(self.generate_participant_folder)
        self.generate_button.setToolTip("Generate participant folder with questionnaires")
        self.generate_button.setWhatsThis("Creates a folder for the participant with all configured questionnaires. The folder will be saved in the target path.")
        main_layout.addWidget(self.generate_button)
        

        
        # ========== EXTRACTION SECTION ==========
        extract_frame = QFrame()
        extract_frame.setFrameStyle(QFrame.Shape.Box | QFrame.Shadow.Raised)
        extract_layout = QGridLayout()
        
        extract_title = QLabel("Extract Data to Excel")
        extract_title.setStyleSheet("font-weight: bold; font-size: 14px;")
        
        source_label = QLabel("Source folder PATH:")
        self.source_pathset = QLabel()
        self.source_pathset.setToolTip("Folder containing participant data to extract")
        source_pathbut = QPushButton("Search")
        source_pathbut.clicked.connect(self.select_source_folder)
        source_pathbut.setToolTip("Click to browse and select the source folder")
        source_pathbut.setWhatsThis("Select the folder containing participant subfolders with completed questionnaire data. The tool will extract data from CSV files in the participant's folder.")
        
        excel_label = QLabel("Masterfile PATH:")
        self.excel_pathset = QLabel()
        self.excel_pathset.setToolTip("Masterfile where extracted data will be appended (.csv, .xls, or .xlsx)")
        excel_pathbut = QPushButton("Search")
        excel_pathbut.clicked.connect(self.select_excel_file)
        excel_pathbut.setToolTip("Click to browse and select the masterfile")
        excel_pathbut.setWhatsThis("Select the masterfile where participant data will be appended. Supported formats: .csv (text), .xls/.xlsx (Excel). The tool automatically detects the format and extracts data accordingly.")
        
        extract_layout.addWidget(extract_title, 0, 0, 1, 3)
        extract_layout.addWidget(source_label, 1, 0)
        extract_layout.addWidget(self.source_pathset, 1, 1)
        extract_layout.addWidget(source_pathbut, 1, 2)
        extract_layout.addWidget(excel_label, 2, 0)
        extract_layout.addWidget(self.excel_pathset, 2, 1)
        extract_layout.addWidget(excel_pathbut, 2, 2)
        
        # Batch mode for extraction
        self.batch_mode_extract = QCheckBox("Batch Mode (multiple IDs)")
        self.batch_mode_extract.setToolTip("Enable to extract data for multiple participants at once")
        self.batch_mode_extract.toggled.connect(self.toggleBatchModeExtraction)
        extract_layout.addWidget(self.batch_mode_extract, 3, 0)
        
        # Import participant list button for extraction
        import_list_btn_extract = QPushButton("Import List...")
        import_list_btn_extract.clicked.connect(lambda: self.importParticipantList(self.batch_text_extract))
        import_list_btn_extract.setToolTip("Import participant IDs from .txt or .csv file")
        extract_layout.addWidget(import_list_btn_extract, 3, 2)
        
        self.batch_text_extract = QTextEdit()
        self.batch_text_extract.setPlaceholderText("Enter participant IDs (one per line or comma-separated):\nP001\nP002\nP003")
        self.batch_text_extract.setVisible(False)
        self.batch_text_extract.setMaximumHeight(100)
        self.batch_text_extract.setToolTip("Enter multiple participant IDs to extract")
        extract_layout.addWidget(self.batch_text_extract, 4, 0, 1, 2)
        
        # Missing Data Report button
        missing_data_btn = QPushButton("Missing Data Report")
        missing_data_btn.clicked.connect(self.generateMissingDataReport)
        missing_data_btn.setToolTip("Scan source folder for participants with incomplete questionnaires")
        missing_data_btn.setWhatsThis("Generates a report showing which participants in the source folder have missing or incomplete questionnaire data files.")
        extract_layout.addWidget(missing_data_btn, 5, 2)
        
        extract_frame.setLayout(extract_layout)
        main_layout.addWidget(extract_frame)
        
        # Single extract button with auto-detection
        self.extract_button = QPushButton("Extract Data to Masterfile")
        self.extract_button.setStyleSheet("font-weight: bold; padding: 10px;")
        self.extract_button.clicked.connect(self.extract_data)
        self.extract_button.setToolTip("Extract participant data from CSV files to the masterfile (auto-detects format)")
        self.extract_button.setWhatsThis("Extracts data from all 'Extract Data.csv' files in the participant's folder and appends it to the masterfile. The tool automatically detects whether to use CSV or Excel format based on the masterfile extension.")
        main_layout.addWidget(self.extract_button)
        

        
        main_widget.setLayout(main_layout)
        self.setCentralWidget(main_widget)
        
        # Window properties
        self.target_path: str = ""
        self.source_path: str = ""
        self.excel_path: str = ""
        self.active_config_name: str = ""
        self.field_mapping: Dict[str, str] = {}
        
        self.resize(700, 600)

    def showPrefillPresetsDialog(self) -> None:
        """Open the Prefill Presets dialog to define default values for PDF fields."""
        # Collect all PDF fields from configured templates
        pdf_fields: List[str] = []
        for row in self.questionnaire_rows:
            template = row.template_path.strip()
            if not template or not template.lower().endswith('.pdf'):
                continue
            fields = self._extract_pdf_form_fields(template)
            pdf_fields.extend(fields.keys())
        
        # Remove duplicates
        pdf_fields = list(dict.fromkeys(pdf_fields))
        
        # Open dialog
        dialog = PrefillPresetsDialog(self, pdf_fields)
        if dialog.exec() == QDialog.DialogCode.Accepted:
            # Save presets to config.ini
            config = ConfigParser()
            config.read('config.ini')
            
            if not config.has_section('PrefillDefaults'):
                config.add_section('PrefillDefaults')
            
            for field, value in dialog.presets.items():
                config.set('PrefillDefaults', field, value)
            
            with open('config.ini', 'w') as f:
                config.write(f)
            
            QMessageBox.information(self, "Success", "Prefill presets saved successfully!")

    def create_questionnaire_rows(self) -> None:
        """Create dynamic rows for questionnaires"""
        try:
            count: int = int(self.quest_count_edit.text())
            if count <= 0:
                raise ValueError
            
            # Clear existing rows
            for row in self.questionnaire_rows:
                for widget in row.get_widgets():
                    widget.deleteLater()
            self.questionnaire_rows.clear()
            
            # Create new rows
            for i in range(count):
                row: QuestionnaireRow = QuestionnaireRow(i)
                self.questionnaire_rows.append(row)
                
                # Add widgets to layout
                base_row = i * 3
                self.scroll_layout.addWidget(row.name_label, base_row, 0)
                self.scroll_layout.addWidget(row.name_edit, base_row, 1, 1, 2)
                
                self.scroll_layout.addWidget(row.path_label, base_row + 1, 0)
                self.scroll_layout.addWidget(row.path_display, base_row + 1, 1)
                self.scroll_layout.addWidget(row.path_button, base_row + 1, 2)
                
                self.scroll_layout.addWidget(row.count_label, base_row + 2, 0)
                self.scroll_layout.addWidget(row.count_edit, base_row + 2, 1)
                
                # Add separator line
                if i < count - 1:
                    line = QFrame()
                    line.setFrameShape(QFrame.Shape.HLine)
                    line.setFrameShadow(QFrame.Shadow.Sunken)
                    self.scroll_layout.addWidget(line, base_row + 3, 0, 1, 3)
                
                # Connect path button (capture row in closure)
                row.path_button.clicked.connect(
                    (lambda r: lambda: self.select_template_file(r))(row)  # type: ignore[misc]
                )
            
            self.scroll_area.setVisible(True)
            
        except ValueError:
            self.error_window("Please enter a valid positive number!")
    
    def select_template_file(self, row: QuestionnaireRow) -> None:
        """Select template file for a questionnaire row"""
        file_path = QFileDialog.getOpenFileName(
            self, "Select Template File", "", "PDF Files (*.pdf);;All Files (*.*)"
        )[0]
        if file_path:
            row.template_path = file_path
            row.path_display.setText(file_path)
    
    def select_target_folder(self) -> None:
        """Select target folder for generation"""
        if folder_path := QFileDialog.getExistingDirectory(self, "Select Target Folder"):
            self.target_path = folder_path
            self.target_pathset.setText(folder_path)
    
    def select_source_folder(self) -> None:
        """Select source folder for extraction"""
        if folder_path := QFileDialog.getExistingDirectory(self, "Select Source Folder"):
            self.source_path = folder_path
            self.source_pathset.setText(folder_path)
    
    def select_excel_file(self) -> None:
        """Select masterfile for extraction (CSV, XLS, or XLSX)"""
        if file_path := QFileDialog.getOpenFileName(
            self, "Select Masterfile", "", 
            "All Supported (*.csv *.xls *.xlsx);;CSV Files (*.csv);;Excel Files (*.xls *.xlsx);;All Files (*.*)"
        )[0]:
            self.excel_path = file_path
            self.excel_pathset.setText(file_path)
    
    def generate_participant_folder(self) -> None:
        """Generate participant folder with questionnaires"""
        try:
            # Handle batch mode
            if self.batch_mode_gen.isChecked():
                participant_ids = self.parseParticipantIDs(self.batch_text_gen.toPlainText())
                if not participant_ids:
                    return self.error_window("Please enter at least one participant ID!")
                
                success_count = 0
                failed: List[str] = []
                
                for pid in participant_ids:
                    try:
                        self._generate_single(pid)
                        success_count += 1
                    except Exception as e:
                        failed.append(f"{pid}: {str(e)}")
                
                result_msg = f"Generated {success_count} of {len(participant_ids)} folders successfully!"
                if failed:
                    result_msg += f"\n\nFailed:\n" + "\n".join(failed)
                
                QMessageBox.information(self, "Batch Generation Complete", result_msg)
            else:
                participant_id = self.nameset.text().strip()
                if not participant_id:
                    return self.error_window("Please enter a participant ID!")
                self._generate_single(participant_id)
                QMessageBox.information(
                    self, "Success",
                    f"Participant folder created successfully!\n{os.path.join(self.target_path, participant_id)}"
                )
            
        except Exception as e:
            self.error_window(f"Error generating folder: {str(e)}")
    
    def _generate_single(self, participant_id: str) -> None:
        """Generate folder for a single participant"""
        if not self.target_path:
            raise ValueError("Please select a target folder!")
        if not self.questionnaire_rows:
            raise ValueError("Please configure questionnaires!")
        
        # Create participant folder
        participant_folder = os.path.join(self.target_path, participant_id)
        if os.path.exists(participant_folder):
            shutil.rmtree(participant_folder)
        os.makedirs(participant_folder)

        generated_pdf_paths: List[str] = []
        
        # Copy questionnaires
        for row in self.questionnaire_rows:
            if not row.template_path:
                continue
            
            survey_name = row.name_edit.text().strip()
            if not survey_name:
                survey_name = f"survey_{row.index + 1}"
            
            try:
                copy_count = int(row.count_edit.text())
            except ValueError:
                copy_count = 1
            
            # Get original filename and extension
            original_filename: str = os.path.basename(row.template_path)
            _, ext = os.path.splitext(original_filename)
            
            # Copy files
            for i in range(copy_count):
                if copy_count == 1:
                    new_filename = f"{participant_id}_{survey_name}{ext}"
                else:
                    new_filename = f"{participant_id}_{survey_name}{i + 1}{ext}"
                
                dest_path = os.path.join(participant_folder, new_filename)
                shutil.copy2(row.template_path, dest_path)
                if ext.lower() == '.pdf':
                    generated_pdf_paths.append(dest_path)

        if generated_pdf_paths:
            self.field_mapping = self._ensure_or_create_field_mapping()
            apply_prefill, prefill_map = self.showGenerationPrefillDialog(participant_id, generated_pdf_paths)
            if not apply_prefill:
                return
            for pdf_path, field_values in prefill_map.items():
                if field_values:
                    self._write_pdf_form_fields(pdf_path, field_values)

    def _write_pdf_form_fields(self, pdf_path: str, field_values: Dict[str, Any]) -> None:
        """Write form field values into a PDF, preserving the full AcroForm structure."""
        # Merge pre-fill presets (only fill empty fields)
        config = ConfigParser()
        config.read('config.ini')
        if config.has_section('PrefillDefaults'):
            for field, value in config.items('PrefillDefaults'):
                if field not in field_values or not field_values[field]:
                    field_values[field] = value

        reader = PdfReader(pdf_path)
        writer = PdfWriter()
        # append() clones the full document including the /AcroForm catalog,
        # which is essential — add_page() alone drops the form structure.
        writer.append(reader)

        def _update_field_tree(fields_array: Any) -> None:
            """Walk the AcroForm /Fields tree and update /V for matching fields."""
            for ref in fields_array:
                try:
                    obj: Any = ref.get_object() if hasattr(ref, 'get_object') else ref
                    name_raw: Any = obj.get('/T')
                    if name_raw is not None:
                        name_str = str(name_raw)
                        if name_str in field_values:
                            new_val = str(field_values[name_str])
                            ft = str(obj.get('/FT', ''))
                            if ft == '/Btn':
                                # Checkbox/radio: translate 1/0 to /Yes or /Off
                                if new_val == '1':
                                    v = NameObject('/Yes')
                                elif new_val == '0':
                                    v = NameObject('/Off')
                                else:
                                    v = NameObject(f'/{new_val}' if not new_val.startswith('/') else new_val)
                                obj[NameObject('/V')] = v
                                obj[NameObject('/AS')] = v
                            else:
                                obj[NameObject('/V')] = create_string_object(new_val)
                    kids: Any = obj.get('/Kids')
                    if kids is not None:
                        _update_field_tree(kids)
                except Exception:
                    continue

        # Primary: walk AcroForm field tree directly.
        # This reliably covers LaTeX/hyperref PDFs where /Fields is the
        # canonical location for form field data.
        try:
            root_obj: Any = cast(Any, writer).trailer['/Root'].get_object()
            acroform_ref: Any = root_obj.get('/AcroForm')
            if acroform_ref is not None:
                acroform: Any = acroform_ref.get_object() if hasattr(acroform_ref, 'get_object') else acroform_ref
                fields_raw: Any = acroform.get('/Fields')
                if fields_raw is not None:
                    _update_field_tree(fields_raw)
        except Exception:
            pass

        # Secondary: per-page annotation path (covers Wondershare and similar).
        for page in writer.pages:
            try:
                writer.update_page_form_field_values(page, field_values)
            except Exception:
                pass

        # NeedAppearances signals viewers to regenerate field appearance streams.
        writer.set_need_appearances_writer()

        with open(pdf_path, "wb") as output_file:
            writer.write(output_file)

    def _build_prefill_defaults(self, participant_id: str, fields: Dict[str, Any]) -> Dict[str, Any]:
        """Build default prefill values for known administrative fields."""
        defaults: Dict[str, Any] = dict(fields)
        participant_aliases = {"participant_id", "participant-id", "Text13_n", "Text13"}
        for field_name in fields.keys():
            if field_name in participant_aliases and str(fields.get(field_name, "")).strip() == "":
                defaults[field_name] = participant_id
        return defaults

    def showGenerationPrefillDialog(self, participant_id: str, generated_pdf_paths: List[str]) -> tuple[bool, Dict[str, Dict[str, Any]]]:
        """Show editable prefill grid for all generated PDFs before participant handover."""
        dialog: QDialog = QDialog(self)
        dialog.setWindowTitle("Preview Data Extraction")
        dialog.setModal(True)
        dialog.resize(800, 600)

        layout = QVBoxLayout()
        label = QLabel(f"Preview data for participant: {participant_id}")
        label.setStyleSheet("font-weight: bold; font-size: 14px;")
        layout.addWidget(label)

        table = QTableWidget()
        table.setColumnCount(3)
        table.setHorizontalHeaderLabels(["Field", "Name", "Data"])

        rows: List[tuple[str, str, str]] = []
        file_to_path: Dict[str, str] = {os.path.basename(p): p for p in generated_pdf_paths}
        file_to_extracted_fields: Dict[str, Dict[str, str]] = {}
        survey_type: str = ''
        for pdf_path in sorted(generated_pdf_paths):
            file_label = os.path.basename(pdf_path)
            survey_type = os.path.splitext(file_label)[0].replace(f'{participant_id}_', '', 1)
            extracted_fields = self._extract_pdf_form_fields(pdf_path)
            if not extracted_fields:
                continue
            defaults = self._build_prefill_defaults(participant_id, extracted_fields)
            file_to_extracted_fields[file_label] = extracted_fields
            checkbox_groups: Dict[str, Dict[str, str]] = {}
            for field_name in sorted(defaults.keys()):
                if field_name.startswith('Check') and '_' in field_name:
                    base_name, option = field_name.split('_', 1)
                    checkbox_groups.setdefault(base_name, {})[option] = str(defaults[field_name])
                else:
                    system_key = f"{survey_type}_{field_name}"
                    rows.append((f"{file_label} :: {field_name}", self.field_mapping.get(system_key, field_name), str(defaults[field_name])))
            for base_name, options in checkbox_groups.items():
                system_key = f"{survey_type}_{base_name}"
                checked = next((opt for opt, val in options.items() if val == '1'), '')
                rows.append((f"{file_label} :: {base_name}", self.field_mapping.get(system_key, base_name), checked))

        table.setRowCount(len(rows))
        for idx, (combined_field, display_name, value) in enumerate(rows):
            field_item = QTableWidgetItem(combined_field)
            field_item.setFlags(field_item.flags() & Qt.ItemFlag.NoItemFlags)
            name_item = QTableWidgetItem(display_name)
            name_item.setFlags(name_item.flags() | Qt.ItemFlag.ItemIsEditable)
            table.setItem(idx, 0, field_item)
            table.setItem(idx, 1, name_item)
            table.setItem(idx, 2, QTableWidgetItem(value))

        table.horizontalHeader().setStretchLastSection(True)  # type: ignore[misc]
        table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeMode.ResizeToContents)  # type: ignore[misc]
        table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeMode.ResizeToContents)  # type: ignore[misc]
        layout.addWidget(table)

        button_layout = QVBoxLayout()
        confirm_btn = QPushButton("Confirm and Generate")
        cancel_btn = QPushButton("Cancel")
        button_layout.addWidget(confirm_btn)
        button_layout.addWidget(cancel_btn)
        layout.addLayout(button_layout)

        dialog.setLayout(layout)

        result: Dict[str, Any] = {"action": "cancel", "values": {}}

        def collect_values() -> Dict[str, Dict[str, Any]]:
            grouped: Dict[str, Dict[str, Any]] = {}
            for pdf_path in generated_pdf_paths:
                grouped[pdf_path] = {}

            for row_index in range(table.rowCount()):
                combined_item = table.item(row_index, 0)
                value_item = table.item(row_index, 2)
                if combined_item is None:
                    continue
                combined = combined_item.text()
                value = value_item.text() if value_item else ""
                if " :: " not in combined:
                    continue
                file_name, field_name = combined.split(" :: ", 1)
                pdf_path = file_to_path.get(file_name)
                if pdf_path:
                    grouped[pdf_path][field_name] = value
            return grouped

        def confirm_and_close() -> None:
            # Persist any Name edits back to self.field_mapping
            prefill_values: Dict[str, Dict[str, Any]] = {}
            for row_index in range(table.rowCount()):
                combined_item = table.item(row_index, 0)
                name_item = table.item(row_index, 1)
                value_item = table.item(row_index, 2)
                if combined_item is None or name_item is None:
                    continue
                combined = combined_item.text()
                if " :: " not in combined:
                    continue
                file_name, field_name = combined.split(" :: ", 1)
                pdf_path_for_key = file_to_path.get(file_name)
                if pdf_path_for_key:
                    file_base = os.path.splitext(file_name)[0]
                    survey_type_key = file_base.replace(f'{participant_id}_', '', 1)
                    system_key = f"{survey_type_key}_{field_name}"
                    new_name = name_item.text().strip()
                    if new_name:
                        self.field_mapping[system_key] = new_name
                    # Store prefill value
                    if value_item:
                        value = value_item.text()
                        if field_name.startswith('Check') and '_' not in field_name:
                            # Handle checkbox grouping: Check1=1 → Check1_1=1, Check1_2=0, etc.
                            base_name = field_name
                            for pdf_path, fields in file_to_extracted_fields.items():
                                for field in fields:
                                    if field.startswith(f"{base_name}_"):
                                        if pdf_path_for_key == pdf_path:
                                            if pdf_path_for_key not in prefill_values:
                                                prefill_values[pdf_path_for_key] = {}
                                            prefill_values[pdf_path_for_key][field] = '1' if field.endswith(f"_{value}") else '0'
                        else:
                            if pdf_path_for_key not in prefill_values:
                                prefill_values[pdf_path_for_key] = {}
                            prefill_values[pdf_path_for_key][field_name] = value
            
            # Save prefill values to config.ini under the active configuration
            if self.active_config_name and prefill_values:
                config = ConfigParser()
                config.read('config.ini')
                if not config.has_section(self.active_config_name):
                    config.add_section(self.active_config_name)
                config.set(self.active_config_name, 'prefill_values_json', json.dumps(prefill_values, ensure_ascii=False))
                with open('config.ini', 'w', encoding='utf-8') as f:
                    config.write(f)
            
            result["action"] = "apply"
            result["values"] = collect_values()
            dialog.accept()

        confirm_btn.clicked.connect(confirm_and_close)
        cancel_btn.clicked.connect(dialog.reject)

        dialog.exec()

        action = result["action"]
        if action == "apply":
            return True, cast(Dict[str, Dict[str, Any]], result["values"])
        raise ValueError("Generation cancelled by user")
    
    def extract_data(self) -> None:
        """Extract data with auto-detection of masterfile format"""
        if not self.excel_path:
            return self.error_window("Please select a masterfile!")

        self.field_mapping = self._load_field_mapping_for_extraction()
        
        # Handle batch mode
        if self.batch_mode_extract.isChecked():
            participant_ids = self.parseParticipantIDs(self.batch_text_extract.toPlainText())
            if not participant_ids:
                return self.error_window("Please enter at least one participant ID!")
            
            success_count = 0
            failed: List[str] = []
            duplicates: List[str] = []
            incomplete: List[str] = []
            
            for pid in participant_ids:
                # Check for duplicate
                if self.checkDuplicate(pid):
                    duplicates.append(pid)
                    continue
                
                # Check data completeness
                is_complete, issues = self.checkDataCompleteness(pid)
                if not is_complete:
                    incomplete.append(f"{pid}: {', '.join(issues)}")
                    continue
                
                try:
                    self._readout_single(pid)
                    success_count += 1
                except Exception as e:
                    failed.append(f"{pid}: {str(e)}")
            
            result_msg = f"Extracted {success_count} of {len(participant_ids)} participants successfully!"
            if duplicates:
                result_msg += f"\n\nSkipped (already in masterfile):\n" + "\n".join(duplicates)
            if incomplete:
                result_msg += f"\n\nIncomplete data:\n" + "\n".join(incomplete)
            if failed:
                result_msg += f"\n\nFailed:\n" + "\n".join(failed)
            
            QMessageBox.information(self, "Batch Extraction Complete", result_msg)
        else:
            # Single participant mode
            participant_id = self.nameset.text().strip()
            if not participant_id:
                return self.error_window("Please enter a participant ID!")
            
            # Check for duplicate
            if self.checkDuplicate(participant_id):
                response = QMessageBox.question(
                    self, "Duplicate Detected",
                    f"Participant {participant_id} already exists in masterfile.\n\nExtract anyway (will create duplicate entry)?",
                    cast(QMessageBox.StandardButton, QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
                )
                if response == QMessageBox.StandardButton.No:
                    return
            
            # Check data completeness
            is_complete, issues = self.checkDataCompleteness(participant_id)
            if not is_complete:
                response = QMessageBox.question(
                    self, "Incomplete Data",
                    f"Data completeness issues:\n" + "\n".join(issues) + "\n\nExtract anyway?",
                    cast(QMessageBox.StandardButton, QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
                )
                if response == QMessageBox.StandardButton.No:
                    return
            
            # Extract data directly to masterfile (CSV or Excel)
            self.readout()
    
    
    def readout(self) -> None:
        """Extract data from participant folder to Excel with preview"""
        try:
            self.field_mapping = self._load_field_mapping_for_extraction()
            participant_id = self.nameset.text().strip()
            all_data = self._prepare_data_for_extraction(participant_id)
            
            # Apply field mapping to all_data for preview
            mapped_data = {}
            for key, value in all_data.items():
                if key == 'participant_id':
                    mapped_data[key] = value
                else:
                    mapped_key = self.field_mapping.get(key, key)
                    mapped_data[mapped_key] = value
            
            # Show editable preview
            confirmed, edited_data = self.showPreviewDialog(mapped_data)  # type: ignore
            if not confirmed:
                return  # User cancelled
            
            # Reverse field mapping for export
            reverse_mapping = {v: k for k, v in self.field_mapping.items()}
            export_data = {}
            for key, value in edited_data.items():
                if key == 'participant_id':
                    export_data[key] = value
                else:
                    export_key = reverse_mapping.get(key, key)
                    export_data[export_key] = value
            
            # Extract
            self._readout_single(participant_id, export_data)  # type: ignore
            
            QMessageBox.information(self, "Success", f"Data extracted successfully for {participant_id}!")
            
        except Exception as e:
            self.error_window(f"Error extracting data: {str(e)}")

    def _extract_pdf_form_fields(self, pdf_path: str) -> Dict[str, str]:
        """Extract form fields from PDF via AcroForm, with fallback widget annotation scan.

        Primary path uses get_fields() (standard AcroForm).
        Fallback scans widget annotations directly, covering Wondershare PDFs,
        LaTeX/hyperref PDFs and other non-standard producers.
        """
        try:
            reader = PdfReader(pdf_path)
        except Exception:
            return {}

        extracted: Dict[str, str] = {}

        def normalize_value(raw: Any) -> str:
            if raw is None:
                return ''
            text = str(raw)
            return text[1:] if text.startswith('/') else text

        def store(name: Any, value: Any) -> None:
            if not name:
                return
            key = str(name)
            if key == 'File' or key in extracted:
                return
            
            # Translate checkbox values (/Yes, /On, /Off) to 1 or 0
            normalized_val = normalize_value(value)
            if normalized_val in ('/Yes', '/On'):
                extracted[key] = '1'
            elif normalized_val == '/Off':
                extracted[key] = '0'
            else:
                extracted[key] = normalized_val

        # Primary: standard AcroForm field tree
        try:
            acro_fields = reader.get_fields() or {}
            for fname, fmeta in acro_fields.items():
                if isinstance(fmeta, dict):
                    fmeta_typed: Dict[str, Any] = cast(Dict[str, Any], fmeta)
                    store(fname, fmeta_typed.get('/V', ''))
        except Exception:
            pass

        # Fallback: walk every page's widget annotations directly.
        # Covers Wondershare, LaTeX/hyperref, and other non-standard producers
        # whose fields are not reachable via get_fields().
        if not extracted:
            try:
                for page in reader.pages:
                    annots: List[Any] = list(page.get('/Annots') or [])
                    for ref in annots:
                        try:
                            annot: Any = ref.get_object()
                            if str(annot.get('/Subtype', '')) != '/Widget':
                                continue
                            # Field name: try own /T, then parent /T
                            fname_raw: Any = annot.get('/T')
                            value_raw: Any = annot.get('/V')
                            parent_ref: Any = annot.get('/Parent')
                            if parent_ref is not None:
                                parent: Any = parent_ref.get_object()
                                if fname_raw is None:
                                    fname_raw = parent.get('/T')
                                if value_raw is None:
                                    value_raw = parent.get('/V')
                            store(fname_raw, value_raw)
                        except Exception:
                            continue
            except Exception:
                pass

        return extracted

    def _inspect_pdf_form_model(self, pdf_path: str) -> Dict[str, Any]:
        """Return lightweight diagnostics about available PDF form structures."""
        info: Dict[str, Any] = {
            'acroform_fields': 0,
            'widget_candidates': 0,
            'xfa_present': False,
            'read_error': False,
        }
        try:
            reader = PdfReader(pdf_path)
        except Exception:
            info['read_error'] = True
            return info

        try:
            fields = reader.get_fields() or {}
            info['acroform_fields'] = len(fields.keys())
        except Exception:
            info['acroform_fields'] = 0

        try:
            widget_count = 0
            for page in reader.pages:
                annots: List[Any] = list(page.get('/Annots') or [])
                for annot_ref in annots:
                    try:
                        annot: Any = annot_ref.get_object()
                    except Exception:
                        continue
                    has_self_name: bool = annot.get('/T') is not None
                    has_parent_name = False
                    parent_ref: Any = annot.get('/Parent')
                    if parent_ref is not None:
                        try:
                            parent: Any = parent_ref.get_object()
                            has_parent_name = parent.get('/T') is not None
                        except Exception:
                            has_parent_name = False
                    if has_self_name or has_parent_name:
                        widget_count += 1
            info['widget_candidates'] = widget_count
        except Exception:
            info['widget_candidates'] = 0

        try:
            trailer_root = reader.trailer.get('/Root', {})
            root_dict: Dict[str, Any] = cast(Dict[str, Any], trailer_root) if isinstance(trailer_root, dict) else {}
            acroform_raw = root_dict.get('/AcroForm', {})
            acroform_dict: Dict[str, Any] = cast(Dict[str, Any], acroform_raw) if isinstance(acroform_raw, dict) else {}
            info['xfa_present'] = acroform_dict.get('/XFA') is not None
        except Exception:
            info['xfa_present'] = False

        return info

    def _validate_extracted_fields(self, survey_type: str, field_data: Dict[str, Any]) -> List[str]:
        """Validate field completeness. Returns advisory warnings only — never blocks extraction."""
        if not field_data:
            return [f"{survey_type}: No fields extracted"]

        issues: List[str] = []

        # Skip field name pattern validation to allow arbitrary field names
        export_fields = {k: v for k, v in field_data.items() if not k.endswith('_n')}
        text_fields = sorted([name for name in export_fields.keys() if TEXT_EXPORT_PATTERN.match(name)])
        if text_fields:
            text_indices = sorted([int(cast(re.Match[str], TEXT_EXPORT_PATTERN.match(name)).group(1)) for name in text_fields])
            expected = set(range(text_indices[0], text_indices[-1] + 1))
            missing = sorted(expected - set(text_indices))
            if missing:
                issues.append(
                    f"{survey_type}: Missing expected text fields: "
                    + ", ".join([f"Text{i}" for i in missing])
                )
            empty_text = [name for name in text_fields if str(export_fields.get(name, '')).strip() == '']
            if empty_text:
                issues.append(f"{survey_type}: Empty text fields: " + ", ".join(empty_text))

        return issues

    def _load_mapping_from_config(self, config_name: str) -> Dict[str, str]:
        """Load field mapping JSON from config section."""
        if not config_name:
            return {}

        config = ConfigParser()
        config.read('config.ini')
        if not config.has_section(config_name):
            return {}

        raw = config.get(config_name, 'field_mapping_json', fallback='')
        if not raw.strip():
            return {}

        try:
            loaded = json.loads(raw)
            if isinstance(loaded, dict):
                loaded_dict = cast(Dict[str, Any], loaded)
                return {str(k): str(v) for k, v in loaded_dict.items()}
        except Exception:
            return {}
        return {}

    def _save_mapping_to_config(self, config_name: str, mapping: Dict[str, str]) -> None:
        """Save field mapping JSON into config section."""
        if not config_name:
            raise ValueError("No active configuration selected")

        config = ConfigParser()
        config.read('config.ini')

        if not config.has_section(config_name):
            config.add_section(config_name)

        config.set(config_name, 'field_mapping_json', json.dumps(mapping, ensure_ascii=False))

        with open('config.ini', 'w', encoding='utf-8') as f:
            config.write(f)

    def _get_expected_system_fields(self) -> set[str]:
        """Build expected system keys from configured generation forms (field names only, no survey_type prefix)."""
        expected: set[str] = set()

        for row in self.questionnaire_rows:
            template = row.template_path.strip()
            if not template or not template.lower().endswith('.pdf'):
                continue

            template_fields = self._extract_pdf_form_fields(template)
            if not template_fields:
                continue

            for field_name in template_fields.keys():
                expected.add(field_name)

        return expected

    def _ensure_or_create_field_mapping(self) -> Dict[str, str]:
        """Ensure mapping exists and matches configured forms, or create it if missing.
        Used during generation (templates must be loaded).
        """
        expected = self._get_expected_system_fields()
        if not expected:
            raise ValueError("No fields found in configured generation forms")

        if not self.active_config_name:
            # No saved config — build a fresh mapping from the current form setup
            generated = {field: field for field in sorted(expected)}
            self.field_mapping = generated
            return generated

        existing = self._load_mapping_from_config(self.active_config_name)
        if not existing:
            generated = {field: field for field in sorted(expected)}
            self._save_mapping_to_config(self.active_config_name, generated)
            self.field_mapping = generated
            return generated

        existing_keys = set(existing.keys())
        if existing_keys != expected:
            missing = sorted(expected - existing_keys)
            extra = sorted(existing_keys - expected)
            parts: List[str] = ["Field mapping does not match configured forms."]
            if missing:
                parts.append("Missing mappings: " + ", ".join(missing))
            if extra:
                parts.append("Unexpected mappings: " + ", ".join(extra))
            raise ValueError("\n".join(parts))

        self.field_mapping = existing
        return existing

    def _load_field_mapping_for_extraction(self) -> Dict[str, str]:
        """Load the saved field mapping for extraction purposes.
        Does not require templates to be configured — loads directly from config.
        Falls back to any already-loaded in-memory mapping.
        """
        if self.active_config_name:
            stored = self._load_mapping_from_config(self.active_config_name)
            if stored:
                print(f"DEBUG: Loaded field mapping from config: {stored}")  # Debug logging
                self.field_mapping = stored
                return stored
        # No config loaded or nothing stored yet — use whatever is in memory
        print(f"DEBUG: Using in-memory field mapping: {self.field_mapping}")  # Debug logging
        return self.field_mapping

    def _generate_field_mapping_from_forms(self) -> Dict[str, str]:
        """Generate field mapping by scanning configured form templates.
        Returns a mapping of field_name -> field_name for all fields found in templates.
        """
        fields: set[str] = set()
        
        for row in self.questionnaire_rows:
            template = row.template_path.strip()
            if not template or not template.lower().endswith('.pdf'):
                continue
            
            try:
                template_fields = self._extract_pdf_form_fields(template)
                if template_fields:
                    fields.update(template_fields.keys())
            except Exception:
                # Silently skip templates that fail to read
                continue
        
        # Generate mapping: field -> field (identity mapping)
        return {field: field for field in sorted(fields)}

    def showFieldMappingDialog(self, mapping: Dict[str, str]) -> tuple[bool, Dict[str, str]]:
        """Open editable mapping list with columns Field | Name."""
        dialog: QDialog = QDialog(self)
        dialog.setWindowTitle("Edit Field Mapping")
        dialog.setModal(True)
        dialog.resize(800, 600)

        layout = QVBoxLayout()
        label = QLabel("Field mapping for active configuration")
        label.setStyleSheet("font-weight: bold; font-size: 14px;")
        layout.addWidget(label)

        table = QTableWidget()
        table.setColumnCount(2)
        table.setHorizontalHeaderLabels(["Field", "Name"])

        ordered_fields = sorted(mapping.keys())
        table.setRowCount(len(ordered_fields))
        for row_index, field_key in enumerate(ordered_fields):
            field_item = QTableWidgetItem(field_key)
            field_item.setFlags(Qt.ItemFlag.ItemIsEnabled)
            table.setItem(row_index, 0, field_item)
            name_item = QTableWidgetItem(mapping.get(field_key, field_key))
            name_item.setFlags(name_item.flags() | Qt.ItemFlag.ItemIsEditable)
            table.setItem(row_index, 1, name_item)

        table.horizontalHeader().setStretchLastSection(True)  # type: ignore[misc]
        table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeMode.ResizeToContents)  # type: ignore[misc]
        layout.addWidget(table)

        button_layout = QVBoxLayout()
        save_btn = QPushButton("Save")
        cancel_btn = QPushButton("Cancel")
        button_layout.addWidget(save_btn)
        button_layout.addWidget(cancel_btn)
        layout.addLayout(button_layout)

        dialog.setLayout(layout)

        result = [False]
        updated: Dict[str, str] = {}

        def confirm_and_close() -> None:
            for row_index in range(table.rowCount()):
                key_item = table.item(row_index, 0)
                name_item = table.item(row_index, 1)
                if key_item is None:
                    continue
                key = key_item.text()
                name = name_item.text().strip() if name_item else key
                updated[key] = name or key
            result[0] = True
            dialog.accept()

        save_btn.clicked.connect(confirm_and_close)
        cancel_btn.clicked.connect(dialog.reject)
        dialog.exec()

        if result[0]:
            return True, updated
        return False, mapping

    def editFieldMappingFromExtraction(self) -> None:
        """Open mapping editor from extraction area with auto-generated mapping from form templates."""
        try:
            # Generate mapping from currently configured forms
            generated_mapping = self._generate_field_mapping_from_forms()
            if not generated_mapping:
                self.error_window("No form fields found in configured templates. Please check your form templates.")
                return
            
            # Load any existing mapping from config and merge with generated
            mapping = dict(generated_mapping)  # Start with generated mapping
            if self.active_config_name:
                stored = self._load_mapping_from_config(self.active_config_name)
                if stored:
                    # For fields that exist in both, use the saved (user-customized) name
                    for field in mapping.keys():
                        if field in stored:
                            mapping[field] = stored[field]
            
            # Open dialog to edit the merged mapping
            confirmed, updated = self.showFieldMappingDialog(mapping)
            if not confirmed:
                return
            
            self.field_mapping = updated
            if self.active_config_name:
                self._save_mapping_to_config(self.active_config_name, updated)
                QMessageBox.information(self, "Success", "Field mapping saved to configuration.")
            else:
                QMessageBox.information(self, "Success", "Field mapping updated (no configuration loaded — changes are not persisted).")
        except Exception as e:
            self.error_window(str(e))

    def _collect_participant_source_data(self, participant_id: str, participant_folder: str) -> Dict[str, Dict[str, Any]]:
        """Collect questionnaire data from direct PDF fields only."""
        survey_rows: Dict[str, Dict[str, Any]] = {}
        validation_issues: List[str] = []

        # Single extraction path: read fields directly from PDF form files.
        pdf_files = sorted([f for f in os.listdir(participant_folder) if f.lower().endswith('.pdf')])
        for pdf_file in pdf_files:
            # Use the full filename (without extension) as the survey_type to avoid assumptions about naming
            survey_type = os.path.splitext(pdf_file)[0]
            field_data = self._extract_pdf_form_fields(os.path.join(participant_folder, pdf_file))
            if field_data:
                survey_rows[survey_type] = field_data
                validation_issues.extend(self._validate_extracted_fields(survey_type, field_data))

        return survey_rows
    
    def _prepare_data_for_extraction(self, participant_id: str) -> Dict[str, Any]:
        """Prepare data dictionary from participant folder."""
        if not participant_id:
            raise ValueError("Please enter a participant ID!")
        if not self.source_path:
            raise ValueError("Please select a source folder!")
        if not self.excel_path:
            raise ValueError("Please select a masterfile!")
        
        # Check if the source_path already ends with the participant_id
        normalized_source_path = os.path.normpath(self.source_path)
        normalized_participant_id = os.path.normpath(participant_id)
        if normalized_source_path.endswith(normalized_participant_id):
            participant_folder = self.source_path
        else:
            participant_folder = os.path.join(self.source_path, participant_id)
        
        if not os.path.exists(participant_folder):
            raise ValueError(f"Participant folder not found: {participant_folder}")

        survey_rows = self._collect_participant_source_data(participant_id, participant_folder)
        if not survey_rows:
            raise ValueError(
                "No readable questionnaire data found. "
                "Expected fillable PDFs with AcroForm fields."
            )

        # Build flattened dictionary for masterfile export.
        all_data: Dict[str, Any] = {'participant_id': participant_id}
        
        for _, row in survey_rows.items():
            checkbox_groups: Dict[str, Dict[str, str]] = {}
            for k, v in row.items():
                if k.startswith('Check') and '_' in k:
                    base_name, option = k.split('_', 1)
                    checkbox_groups.setdefault(base_name, {})[option] = v
                else:
                    all_data[k] = v
            for base_name, options in checkbox_groups.items():
                checked = next((opt for opt, val in options.items() if val == '1'), None)
                if checked:
                    all_data[base_name] = checked
        
        return all_data
    
    def _readout_single(self, participant_id: str, all_data: Optional[Dict[str, Any]] = None) -> None:  # all_data: {field_name: field_value}
        """Extract data for single participant to masterfile (CSV or Excel, no preview)"""
        if all_data is None:
            all_data = self._prepare_data_for_extraction(participant_id)

        # Map PDF fields to masterfile columns using field_mapping (set by Edit Field Mapping)
        export_data: Dict[str, Any] = {'participant_id': participant_id}
        for key, value in all_data.items():
            if key == 'participant_id':
                continue
            column_name = self.field_mapping.get(key, key)
            export_data[column_name] = value
        
        # Detect masterfile format from file extension
        file_ext = self.excel_path.lower().split('.')[-1]
        
        if file_ext in ['xls', 'xlsx']:
            # Open Excel workbook
            wb: xlw.Book = xlw.Book(self.excel_path)
            
            try:
                # Attempt to find a sheet named 'Data' or use the first sheet
                sheet_names: List[str] = [cast(str, s.name) for s in wb.sheets]  # type: ignore[misc]
                sheet: xlw.Sheet = cast(xlw.Sheet, wb.sheets['Data' if 'Data' in sheet_names else 0])
                
                # Find next empty row (data starts at row 2; row 1 is headers)
                next_row = 2
                while sheet.range(f'A{next_row}').value is not None:  # type: ignore[misc]
                    next_row += 1
                
                # Write headers in row 1 if not already present
                if sheet.range('A1').value is None:  # type: ignore[misc]
                    headers = list(export_data.keys())
                    for col, header in enumerate(headers, start=1):
                        sheet.range(1, col).value = header  # type: ignore[misc]
                
                # Write data to next empty row
                for col, key in enumerate(export_data.keys(), start=1):
                    sheet.range(next_row, col).value = export_data[key]  # type: ignore[misc]
                
                # Save but leave the workbook open so the user can verify the result.
                cast(xlw.Book, wb).save()  # type: ignore[misc]
                
            except Exception:
                if 'wb' in locals():
                    wb.close()  # type: ignore[misc]
                raise
        elif file_ext == 'csv':
            # Write to CSV
            csv_path = self.excel_path
            file_exists = os.path.exists(csv_path) and os.path.getsize(csv_path) > 0
            
            # Read existing headers if file exists
            existing_headers: List[str] = []
            if file_exists:
                with open(csv_path, 'r', newline='', encoding='utf-8') as f:
                    reader = csv.DictReader(f)
                    existing_headers = list(reader.fieldnames) if reader.fieldnames else []
            
            # Merge headers
            all_headers = list(dict.fromkeys(existing_headers + list(export_data.keys())))
            
            # Append to CSV
            with open(csv_path, 'a', newline='', encoding='utf-8') as f:
                writer = csv.DictWriter(f, fieldnames=all_headers)
                if not file_exists:
                    writer.writeheader()
                writer.writerow(export_data)
        else:
            raise ValueError(f"Unsupported masterfile format: {file_ext}")
    

    
    
    def updateRecentConfigsMenu(self) -> None:
        """Update recent configurations and field mapping menus with available configs"""
        if self.recent_configs_menu is not None:
            self.recent_configs_menu.clear()
        if self.delete_configs_menu is not None:
            self.delete_configs_menu.clear()
        config = ConfigParser()
        config.read('config.ini')
        sections = config.sections()
        
        if not sections:
            if self.recent_configs_menu is not None:
                no_config_action = QAction("(No configurations available)", self)
                no_config_action.setEnabled(False)
                self.recent_configs_menu.addAction(no_config_action)  # type: ignore[misc]
            
            if self.delete_configs_menu is not None:
                no_delete_action = QAction("(No configurations available)", self)
                no_delete_action.setEnabled(False)
                self.delete_configs_menu.addAction(no_delete_action)  # type: ignore[misc]
        else:
            for section in sections:
                # Recent configurations - load on click
                if self.recent_configs_menu is not None:
                    load_action = QAction(section, self)
                    load_action.triggered.connect(lambda checked, name=section: self.ConfigLoad(name))  # type: ignore[misc]
                    self.recent_configs_menu.addAction(load_action)  # type: ignore[misc]
                
                # Delete configurations - delete on click with confirmation
                if self.delete_configs_menu is not None:
                    delete_action = QAction(section, self)
                    delete_action.triggered.connect(lambda checked, name=section: self.confirmDeleteConfig(name))  # type: ignore[misc]
                    self.delete_configs_menu.addAction(delete_action)  # type: ignore[misc]
    
    def showSaveConfigWindow(self) -> None:
        """Show save configuration dialog"""
        self.save_win = SaveConfigWindow(self)
        self.save_win.signal.connect(self.ConfigGen)
        self.save_win.exec()
    
    def confirmDeleteConfig(self, name: str) -> None:
        """Show confirmation dialog and delete configuration"""
        dlg = QMessageBox(self)
        dlg.setWindowTitle("Confirm Delete")
        dlg.setIcon(QMessageBox.Icon.Question)
        dlg.setStandardButtons(cast(QMessageBox.StandardButton, QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No))
        dlg.setText(f"Are you sure you want to delete configuration '{name}'?")
        
        if dlg.exec() == QMessageBox.StandardButton.Yes:
            try:
                config = ConfigParser()
                config.read('config.ini')
                config.remove_section(name)
                with open('config.ini', 'w') as f:
                    config.write(f)
                QMessageBox.information(self, "Success", f"Configuration '{name}' deleted successfully!")
            except Exception as e:
                self.error_window(f"Error deleting configuration: {str(e)}")
    
    def ConfigGen(self, name: str) -> None:
        """Save current configuration"""
        if not name:
            return self.error_window("No configuration name provided!")
        
        try:
            config = ConfigParser()
            config.read('config.ini')
            
            # Remove section if it exists, then add it fresh
            if config.has_section(name):
                config.remove_section(name)
            config.add_section(name)
            
            # Save generation settings
            config.set(name, 'target_path', self.target_path)
            config.set(name, 'quest_count', str(len(self.questionnaire_rows)))
            
            for i, row in enumerate(self.questionnaire_rows):
                config.set(name, f'quest_{i}_name', row.name_edit.text())
                config.set(name, f'quest_{i}_path', row.template_path)
                config.set(name, f'quest_{i}_count', row.count_edit.text())
            
            # Save extraction settings
            config.set(name, 'source_path', self.source_path)
            config.set(name, 'excel_path', self.excel_path)

            expected_fields = self._get_expected_system_fields()
            if expected_fields:
                if set(self.field_mapping.keys()) == expected_fields:
                    mapping_to_save = self.field_mapping
                else:
                    mapping_to_save = {field: field for field in sorted(expected_fields)}
                    self.field_mapping = mapping_to_save
                config.set(name, 'field_mapping_json', json.dumps(mapping_to_save, ensure_ascii=False))
            
            with open('config.ini', 'w', encoding='utf-8') as f:
                config.write(f)

            self.active_config_name = name
            
            QMessageBox.information(self, "Success", f"Configuration '{name}' saved successfully!")
            
        except Exception as e:
            self.error_window(f"Error saving configuration: {str(e)}")
    
    def ConfigLoad(self, name: str) -> None:
        """Load configuration"""
        try:
            config = ConfigParser()
            config.read('config.ini')
            
            if not config.has_section(name):
                self.error_window(f"Configuration '{name}' not found!")
                return
            
            # Load generation settings
            self.target_path = config.get(name, 'target_path', fallback='')
            self.target_pathset.setText(self.target_path)
            
            quest_count = config.getint(name, 'quest_count', fallback=0)
            if quest_count > 0:
                self.quest_count_edit.setText(str(quest_count))
                self.create_questionnaire_rows()
                
                for i, row in enumerate(self.questionnaire_rows):
                    row.name_edit.setText(config.get(name, f'quest_{i}_name', fallback=''))
                    row.template_path = config.get(name, f'quest_{i}_path', fallback='')
                    row.path_display.setText(row.template_path)
                    row.count_edit.setText(config.get(name, f'quest_{i}_count', fallback='1'))
            
            # Load extraction settings
            self.source_path = config.get(name, 'source_path', fallback='')
            self.source_pathset.setText(self.source_path)
            self.excel_path = config.get(name, 'excel_path', fallback='')
            self.excel_pathset.setText(self.excel_path)

            self.active_config_name = name
            # Load mapping directly from config — don't require templates to be present
            stored_mapping = self._load_mapping_from_config(name)
            if stored_mapping:
                self.field_mapping = stored_mapping
            
            QMessageBox.information(self, "Success", f"Configuration '{name}' loaded successfully!")
            
        except Exception as e:
            self.error_window(f"Error loading configuration: {str(e)}")
    
    def enterWhatsThisMode(self) -> None:
        """Enter What's This help mode"""
        QWhatsThis.enterWhatsThisMode()
    
    def onMyToolBarButtonClick3(self) -> None:
        """Open user manual"""
        manual_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'USER_MANUAL.pdf')
        if os.path.exists(manual_path):
            os.startfile(manual_path)
        else:
            self.error_window("User manual not found!\\n\\nExpected location:\\n" + manual_path)
    
    def toggleBatchModeGeneration(self, checked: bool) -> None:
        """Toggle batch mode for generation"""
        self.nameset.setVisible(not checked)
        self.batch_text_gen.setVisible(checked)
    
    def toggleBatchModeExtraction(self, checked: bool) -> None:
        """Toggle batch mode for extraction"""
        self.nameset.setVisible(not checked)
        self.batch_text_extract.setVisible(checked)
    
    def importParticipantList(self, target_text_edit: QTextEdit) -> None:
        """Import participant IDs from a text or CSV file"""
        file_path = QFileDialog.getOpenFileName(
            self, "Import Participant List", "",
            "Text Files (*.txt);;CSV Files (*.csv);;All Files (*.*)"
        )[0]
        
        if not file_path:
            return
        
        try:
            participant_ids: List[str] = []
            
            # Read file based on extension
            if file_path.lower().endswith('.csv'):
                with open(file_path, 'r', encoding='utf-8') as f:
                    reader = csv.reader(f)
                    for row in reader:
                        # Take first column or all values in row
                        participant_ids.extend([cell.strip() for cell in row if cell.strip()])
            else:
                # Text file
                with open(file_path, 'r', encoding='utf-8') as f:
                    for line in f:
                        # Support comma-separated or one per line
                        for pid in line.replace(',', '\n').split('\n'):
                            if pid.strip():
                                participant_ids.append(pid.strip())
            
            if participant_ids:
                # Remove duplicates while preserving order
                seen: set[str] = set()
                unique_ids: List[str] = []
                for pid in participant_ids:
                    if pid not in seen:
                        seen.add(pid)
                        unique_ids.append(pid)
                
                target_text_edit.setPlainText('\n'.join(unique_ids))
                QMessageBox.information(
                    self, "Import Successful",
                    f"Imported {len(unique_ids)} unique participant IDs from:\n{os.path.basename(file_path)}"
                )
            else:
                self.error_window("No participant IDs found in the file!")
                
        except Exception as e:
            self.error_window(f"Error importing participant list: {str(e)}")
    
    def saveTemplateBundle(self) -> None:
        """Save questionnaire configuration as a template bundle"""
        if not self.questionnaire_rows:
            return self.error_window("No questionnaire configuration to save!")
        
        # Ask for bundle name
        bundle_name, ok = QInputDialog.getText(
            self, "Save Template Bundle",
            "Enter a name for this template bundle:"
        )
        
        if not ok or not bundle_name.strip():
            return
        
        bundle_name = bundle_name.strip()
        
        try:
            # Create bundles directory if it doesn't exist
            bundles_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'template_bundles')
            os.makedirs(bundles_dir, exist_ok=True)
            
            # Save bundle as JSON
            bundle_data: Dict[str, Any] = {
                'name': bundle_name,
                'questionnaire_count': len(self.questionnaire_rows),
                'questionnaires': []
            }
            
            for i, row in enumerate(self.questionnaire_rows):
                bundle_data['questionnaires'].append({
                    'index': i,
                    'name': row.name_edit.text(),
                    'template_path': row.template_path,
                    'copy_count': row.count_edit.text()
                })
            
            bundle_file = os.path.join(bundles_dir, f"{bundle_name}.json")
            
            # Check if file exists
            if os.path.exists(bundle_file):
                response = QMessageBox.question(
                    self, "Overwrite Bundle",
                    f"Template bundle '{bundle_name}' already exists. Overwrite?",
                    cast(QMessageBox.StandardButton, QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
                )
                if response == QMessageBox.StandardButton.No:
                    return
            
            with open(bundle_file, 'w', encoding='utf-8') as f:
                json.dump(bundle_data, f, indent=2)
            
            QMessageBox.information(
                self, "Success",
                f"Template bundle '{bundle_name}' saved successfully!"
            )
            
        except Exception as e:
            self.error_window(f"Error saving template bundle: {str(e)}")
    
    def loadTemplateBundle(self) -> None:
        """Load a template bundle"""
        try:
            bundles_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'template_bundles')
            
            if not os.path.exists(bundles_dir):
                return self.error_window("No template bundles found!")
            
            # Get list of available bundles
            bundle_files = [f for f in os.listdir(bundles_dir) if f.endswith('.json')]
            
            if not bundle_files:
                return self.error_window("No template bundles found!")
            
            # Show selection dialog
            bundle_names = [f.replace('.json', '') for f in bundle_files]
            
            bundle_name, ok = QInputDialog.getItem(
                self, "Load Template Bundle",
                "Select a template bundle to load:",
                bundle_names, 0, False
            )
            
            if not ok:
                return
            
            # Load bundle
            bundle_file = os.path.join(bundles_dir, f"{bundle_name}.json")
            
            with open(bundle_file, 'r', encoding='utf-8') as f:
                bundle_data = json.load(f)
            
            # Set questionnaire count and create rows
            quest_count = bundle_data['questionnaire_count']
            self.quest_count_edit.setText(str(quest_count))
            self.create_questionnaire_rows()
            
            # Populate questionnaire rows
            for q_data in bundle_data['questionnaires']:
                idx = q_data['index']
                if idx < len(self.questionnaire_rows):
                    row = cast(QuestionnaireRow, self.questionnaire_rows[idx])
                    row.name_edit.setText(q_data['name'])
                    row.template_path = q_data['template_path']
                    row.path_display.setText(q_data['template_path'])
                    row.count_edit.setText(q_data['copy_count'])
            
            QMessageBox.information(
                self, "Success",
                f"Template bundle '{bundle_name}' loaded successfully!"
            )
            
        except Exception as e:
            self.error_window(f"Error loading template bundle: {str(e)}")
    
    def parseParticipantIDs(self, text: str) -> List[str]:
        """Parse participant IDs from text (comma or newline separated)"""
        ids: List[str] = []
        for line in text.replace(',', '\n').split('\n'):
            if pid := line.strip():
                ids.append(pid)
        return ids
    
    def checkDuplicate(self, participant_id: str) -> bool:
        """Check if participant ID already exists in masterfile"""
        if not os.path.exists(self.excel_path):
            return False

        file_ext = self.excel_path.lower().split('.')[-1]

        try:
            if file_ext == 'csv':
                with open(self.excel_path, 'r', encoding='utf-8') as f:
                    reader = csv.DictReader(f)
                    return any(row.get('participant_id') == participant_id for row in reader)
            elif file_ext in ['xls', 'xlsx']:
                # Read the first column via xlwings and close immediately.
                # app.visible=False prevents Excel from appearing on screen during the check.
                app = xlw.App(visible=False, add_book=False)
                try:
                    wb_check = cast(xlw.Book, app.books.open(self.excel_path))  # type: ignore[misc]
                    try:
                        ws_check = wb_check.sheets[0]  # type: ignore[misc]
                        existing_ids = ws_check.range('A2:A1000').value  # type: ignore[misc]
                        id_list: list[Any] = list(existing_ids) if existing_ids else []  # type: ignore[arg-type]
                        return participant_id in [str(v).strip() for v in id_list if v is not None]
                    finally:
                        wb_check.close()  # type: ignore[misc]
                finally:
                    app.quit()
        except Exception:
            return False

        return False
    
    def checkDataCompleteness(self, participant_id: str) -> tuple[bool, List[str]]:
        """Check if participant folder has readable PDF questionnaire sources."""
        participant_folder = os.path.join(self.source_path, participant_id)
        print(f"DEBUG: Checking data completeness for participant {participant_id} in folder {participant_folder}")  # Debug logging
        if not os.path.exists(participant_folder):
            return False, [f"Folder not found: {participant_folder}"]

        pdf_files = sorted([f for f in os.listdir(participant_folder) if f.lower().endswith('.pdf')])
        print(f"DEBUG: Found PDF files: {pdf_files}")  # Debug logging

        # Single extraction path: direct PDF form fields.
        usable_pdf_count = 0
        validation_issues: List[str] = []
        unreadable_pdf_details: List[str] = []
        for pdf_file in pdf_files:
            pdf_path = os.path.join(participant_folder, pdf_file)
            survey_type = os.path.splitext(pdf_file)[0].replace(f'{participant_id}_', '', 1)
            print(f"DEBUG: Processing PDF file: {pdf_file}, derived survey_type: {survey_type}")  # Debug logging
            field_data = self._extract_pdf_form_fields(pdf_path)
            print(f"DEBUG: Extracted fields from {pdf_file}: {field_data}")  # Debug logging
            if field_data:
                usable_pdf_count += 1
                validation_issues.extend(self._validate_extracted_fields(survey_type, field_data))
            else:
                model_info = self._inspect_pdf_form_model(pdf_path)
                if model_info['read_error']:
                    unreadable_pdf_details.append(
                        f"{pdf_file}: could not be read (file may be corrupted, locked, or encrypted)"
                    )
                else:
                    unreadable_pdf_details.append(
                        f"{pdf_file}: unsupported form format for this extraction path "
                        f"(requires AcroForm/XFA with addressable field widgets). "
                        f"Detected "
                        f"(AcroForm={model_info['acroform_fields']}, "
                        f"Widgets={model_info['widget_candidates']}, "
                        f"XFA={'yes' if model_info['xfa_present'] else 'no'})"
                    )

        if usable_pdf_count == 0:
            base = [
                "No compatible PDF form fields found for extraction. "
                "At least one source PDF must expose fields via AcroForm/XFA widgets."
            ]
            return False, base + unreadable_pdf_details

        # Check for common incomplete patterns
        issues: List[str] = []
        expected_count = len(self.questionnaire_rows) if self.questionnaire_rows else None
        print(f"DEBUG: Expected questionnaire count: {expected_count}, Usable PDF count: {usable_pdf_count}")  # Debug logging

        found_count = usable_pdf_count
        found_label = "PDF forms"

        # Only compare counts if expected_count is set and non-zero
        if expected_count and expected_count > 0 and found_count < expected_count:
            issues.append(f"Expected {expected_count} questionnaires, found {found_count} {found_label}")
            issues.extend(unreadable_pdf_details)

        if issues:
            return False, issues

        return True, [f"PDF source detected ({usable_pdf_count} files with readable form fields)"]
    
    def showPreviewDialog(self, data: Dict[str, Any]) -> tuple[bool, Dict[str, Any]]:  # data: {field_name: field_value}
        """Show editable preview dialog and return confirmation flag plus edited data."""
        dialog: QDialog = QDialog(self)
        dialog.setWindowTitle("Preview Data Extraction")
        dialog.setModal(True)
        dialog.resize(800, 600)
        
        layout = QVBoxLayout()
        
        label = QLabel(f"Preview data for participant: {data.get('participant_id', 'Unknown')}")
        label.setStyleSheet("font-weight: bold; font-size: 14px;")
        layout.addWidget(label)
        
        # Load prefill values from config.ini
        prefill_values: Dict[str, Dict[str, Any]] = {}
        if self.active_config_name:
            config = ConfigParser()
            config.read('config.ini')
            if config.has_option(self.active_config_name, 'prefill_values_json'):
                try:
                    prefill_values = json.loads(config.get(self.active_config_name, 'prefill_values_json', fallback='{}'))
                except Exception:
                    prefill_values = {}
        
        # Create table
        table = QTableWidget()
        table.setColumnCount(3)
        table.setHorizontalHeaderLabels(["Field", "Name", "Data"])
        table.setRowCount(len(data))
        
        ordered_items = list(data.items())
        for i, (key, value) in enumerate(ordered_items):
            key_item = QTableWidgetItem(str(key))
            key_item.setFlags(Qt.ItemFlag.ItemIsEnabled)
            mapped_name = 'Participant ID' if str(key) == 'participant_id' else self.field_mapping.get(str(key), str(key))
            name_item = QTableWidgetItem(mapped_name)
            name_item.setFlags(name_item.flags() | Qt.ItemFlag.ItemIsEditable)
            
            # Reapply prefill value if available
            prefill_value = str(value)
            if str(key) != 'participant_id':
                # Extract field_name from key (e.g., "survey_type_field_name")
                field_name = str(key).split('_', 1)[-1]
                for fields in prefill_values.values():
                    if field_name in fields:
                        prefill_value = str(fields[field_name])
                        break
            
            table.setItem(i, 0, key_item)
            table.setItem(i, 1, name_item)
            table.setItem(i, 2, QTableWidgetItem(prefill_value))
        
        table.horizontalHeader().setStretchLastSection(True)  # type: ignore[misc]
        table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeMode.ResizeToContents)  # type: ignore[misc]
        table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeMode.ResizeToContents)  # type: ignore[misc]
        layout.addWidget(table)
        
        # Buttons
        button_layout = QVBoxLayout()
        confirm_btn = QPushButton("Confirm and Extract")
        cancel_btn = QPushButton("Cancel")
        button_layout.addWidget(confirm_btn)
        button_layout.addWidget(cancel_btn)
        layout.addLayout(button_layout)
        
        dialog.setLayout(layout)
        
        result = [False]
        edited_data: Dict[str, Any] = {}
        
        def confirm_and_close() -> None:
            for row_index in range(table.rowCount()):
                key_item = table.item(row_index, 0)
                name_item = table.item(row_index, 1)
                value_item = table.item(row_index, 2)
                if key_item is None:
                    continue
                key = key_item.text()
                value = value_item.text() if value_item else ""
                edited_data[key] = value
                # Persist Name edits back to field_mapping
                if name_item is not None:
                    new_name = name_item.text().strip()
                    if new_name and key != 'participant_id':
                        self.field_mapping[key] = new_name
            result[0] = True
            dialog.accept()
        
        confirm_btn.clicked.connect(confirm_and_close)
        cancel_btn.clicked.connect(dialog.reject)
        
        dialog.exec()
        return result[0], edited_data if result[0] else data
    
    def generateMissingDataReport(self) -> None:
        """Generate report of participants with incomplete data"""
        if not self.source_path:
            return self.error_window("Please select a source folder!")
        
        if not os.path.exists(self.source_path):
            return self.error_window(f"Source folder not found: {self.source_path}")
        
        # Scan all participant folders
        all_folders = [f for f in os.listdir(self.source_path) 
                      if os.path.isdir(os.path.join(self.source_path, f))]
        
        if not all_folders:
            return self.error_window("No participant folders found in source directory!")
        
        report: List[str] = []
        complete_count = 0
        incomplete_count = 0
        
        for folder in sorted(all_folders):
            is_complete, details = self.checkDataCompleteness(folder)
            if is_complete:
                complete_count += 1
                report.append(f"✓ {folder}: Complete ({len(details)} files)")
            else:
                incomplete_count += 1
                report.append(f"✗ {folder}: INCOMPLETE - {', '.join(details)}")
        
        # Show report dialog
        dialog = QDialog(self)
        dialog.setWindowTitle("Missing Data Report")
        dialog.setModal(True)
        dialog.resize(700, 500)
        
        layout = QVBoxLayout()
        
        summary = QLabel(f"Summary: {complete_count} complete, {incomplete_count} incomplete (out of {len(all_folders)} total)")
        summary.setStyleSheet("font-weight: bold; font-size: 14px;")
        layout.addWidget(summary)
        
        text_area = QTextEdit()
        text_area.setReadOnly(True)
        text_area.setText('\n'.join(report))
        layout.addWidget(text_area)
        
        close_btn = QPushButton("Close")
        close_btn.clicked.connect(dialog.accept)
        layout.addWidget(close_btn)
        
        dialog.setLayout(layout)
        dialog.exec()
    
    def error_window(self, message: str) -> None:
        """Display error message"""
        QMessageBox.warning(self, "ERROR", message)


if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())
