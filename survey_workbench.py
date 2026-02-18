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
from typing import Optional, List, cast, Dict, Any
from PyQt5.QtWidgets import (QApplication, QMainWindow, QPushButton, 
                             QGridLayout, QWidget, QLineEdit, QLabel,
                             QFileDialog, QMessageBox, QAction, QStatusBar,
                             QComboBox, QScrollArea, QVBoxLayout, QFrame,
                             QMenu, QDialog, QWhatsThis, QTextEdit, QTableWidget,
                             QTableWidgetItem, QHeaderView, QCheckBox, QInputDialog)
from configparser import ConfigParser
from PyQt5.QtCore import pyqtSignal
import xlwings as xlw  # type: ignore[import]

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

class LoadConfigWindow(QWidget):
    signal = pyqtSignal(str)
    
    def __init__(self, parent: Optional[QWidget] = None) -> None:
        super().__init__(parent)
        
        self.config: ConfigParser = ConfigParser()
        self.setWindowTitle("Survey Workbench")
        layout: QGridLayout = QGridLayout()
        self.label: QLabel = QLabel("Choose configuration:")
        layout.addWidget(self.label, 0, 0)
        self.choose: QComboBox = QComboBox()
        layout.addWidget(self.choose, 0, 1)
        self.delete: QPushButton = QPushButton('Delete configuration')
        layout.addWidget(self.delete, 1, 1)
        self.load: QPushButton = QPushButton('Load configuration')
        layout.addWidget(self.load, 1, 0)
        self.setLayout(layout)
        
        self.config.read('config.ini')
        self.choose.addItems(list(self.config.sections()))
        
        self.delete.clicked.connect(self.areYouSure)
        self.load.clicked.connect(self.Load)
        
    def areYouSure(self) -> None:
        dlg = QMessageBox(self)
        dlg.setWindowTitle("Delete Configuration")
        dlg.setIcon(QMessageBox.Icon.Question)
        dlg.setStandardButtons(cast(QMessageBox.StandardButton, QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No))
        dlg.setText("Are you sure you want to delete this configuration permanently?")
        
        if dlg.exec() == QMessageBox.StandardButton.Yes:
            self.Delete()
        
    def Delete(self) -> None:
        self.config.remove_section(self.choose.currentText())
        with open('config.ini', 'w') as f:
            self.config.write(f)
        self.choose.clear()
        self.choose.addItems(self.config.sections())
        
    def Load(self) -> None:
        self.signal.emit(self.choose.currentText())
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

class MainWindow(QMainWindow):
    def __init__(self, parent: Optional[QWidget] = None) -> None:
        super().__init__(parent)
        self.setWindowTitle("Survey Workbench")
        
        # Initialize questionnaire rows list
        self.questionnaire_rows: List[QuestionnaireRow] = []
        self.recent_configs_menu: QMenu
        self.delete_configs_menu: QMenu
        
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
        file_menu.aboutToShow.connect(self.updateRecentConfigsMenu)  # type: ignore[misc]
        
        # Delete Configuration submenu
        self.delete_configs_menu = file_menu.addMenu("&Delete Configuration")
        assert self.delete_configs_menu is not None
        
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
        self.generate_button.clicked.connect(self.generate)
        self.generate_button.setToolTip("Create participant folder with configured questionnaire templates")
        self.generate_button.setWhatsThis("Generates a folder structure for the participant with all specified questionnaire templates. Each template will be copied the specified number of times with appropriate naming.")
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
        
        self.resize(700, 600)
        
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
    
    def generate(self) -> None:
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
    
    def extract_data(self) -> None:
        """Extract data with auto-detection of masterfile format"""
        if not self.excel_path:
            return self.error_window("Please select a masterfile!")
        
        # Handle batch mode
        if self.batch_mode_extract.isChecked():
            participant_ids = self.parseParticipantIDs(self.batch_text_extract.toPlainText())
            if not participant_ids:
                return self.error_window("Please enter at least one participant ID!")
            
            # Detect format from file extension
            file_ext = self.excel_path.lower().split('.')[-1]
            
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
                    if file_ext == 'csv':
                        self._readout_csv_single(pid)
                    elif file_ext in ['xls', 'xlsx']:
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
            
            # Detect format from file extension
            file_ext = self.excel_path.lower().split('.')[-1]
            
            if file_ext == 'csv':
                self.readout_csv()
            elif file_ext in ['xls', 'xlsx']:
                self.readout()
            else:
                self.error_window("Unsupported file format! Please use .csv, .xls, or .xlsx")
    
    
    def readout(self) -> None:
        """Extract data from participant folder to Excel with preview"""
        try:
            participant_id = self.nameset.text().strip()
            all_data = self._prepare_data_for_extraction(participant_id)
            
            # Show preview
            if not self.showPreviewDialog(all_data):
                return  # User cancelled
            
            # Extract
            self._readout_single(participant_id, all_data)
            
            QMessageBox.information(self, "Success", f"Data extracted successfully for {participant_id}!")
            
        except Exception as e:
            self.error_window(f"Error extracting data: {str(e)}")
    
    def _prepare_data_for_extraction(self, participant_id: str) -> Dict[str, Any]:
        """Prepare data dictionary from participant folder"""
        if not participant_id:
            raise ValueError("Please enter a participant ID!")
        if not self.source_path:
            raise ValueError("Please select a source folder!")
        if not self.excel_path:
            raise ValueError("Please select a masterfile!")
        
        participant_folder = os.path.join(self.source_path, participant_id)
        if not os.path.exists(participant_folder):
            raise ValueError(f"Participant folder not found: {participant_folder}")
        
        # Find all Extract Data CSV files
        csv_files = sorted([f for f in os.listdir(participant_folder) if f.endswith('_Extract Data.csv')])
        if not csv_files:
            raise ValueError("No Extract Data CSV files found in participant folder!")
        
        # Process CSV files and build data dictionary
        all_data: Dict[str, Any] = {'participant_id': participant_id}
        for csv_file in csv_files:
            survey_type = csv_file.replace(f'{participant_id}_', '').replace('_Extract Data.csv', '')
            with open(os.path.join(participant_folder, csv_file), 'r', encoding='utf-8') as f:
                for row in csv.DictReader(f):
                    all_data.update({f"{survey_type}_{k}": v for k, v in row.items() if k != 'File'})
        
        return all_data
    
    def _readout_single(self, participant_id: str, all_data: Optional[Dict[str, Any]] = None) -> None:
        """Extract data for single participant to Excel (no preview)"""
        if all_data is None:
            all_data = self._prepare_data_for_extraction(participant_id)
        
        # Open Excel workbook
        wb: xlw.Book = xlw.Book(self.excel_path)
        
        try:
            # Attempt to find a sheet named 'Data' or use the first sheet
            sheet_names: List[str] = [cast(str, s.name) for s in wb.sheets]  # type: ignore[misc]
            sheet: xlw.Sheet = cast(xlw.Sheet, wb.sheets['Data' if 'Data' in sheet_names else 0])
            
            # Find next empty row
            next_row = 1
            while sheet.range(f'A{next_row}').value is not None:  # type: ignore[misc]
                next_row += 1
            
            # Write data (headers on first row, values on next_row)
            sorted_data = sorted((k, v) for k, v in all_data.items() if k != 'participant_id')
            sheet.range(f'A{next_row}').value = participant_id  # type: ignore[misc]
            for col, (key, value) in enumerate(sorted_data, start=2):
                if next_row == 1:
                    sheet.range(1, col).value = key  # type: ignore[misc]
                sheet.range(next_row, col).value = value  # type: ignore[misc]
            
            # Save as XLS format
            if not self.excel_path.lower().endswith('.xls'):
                xls_path = self.excel_path.rsplit('.', 1)[0] + '.xls'
                cast(xlw.Book, wb).save(xls_path)  # type: ignore[misc]
                wb.close()  # type: ignore[misc]
                self.excel_path = xls_path
                self.excel_pathset.setText(xls_path)
            else:
                cast(xlw.Book, wb).save()  # type: ignore[misc]
                
        finally:
            if 'wb' in locals():
                wb.close()  # type: ignore[misc]
    
    def readout_csv(self) -> None:
        """Extract data from participant folder to CSV with preview"""
        try:
            participant_id = self.nameset.text().strip()
            all_data = self._prepare_data_for_extraction(participant_id)
            
            # Show preview
            if not self.showPreviewDialog(all_data):
                return  # User cancelled
            
            # Extract
            self._readout_csv_single(participant_id, all_data)
            
            QMessageBox.information(self, "Success", f"Data extracted successfully for {participant_id}!")
            
        except Exception as e:
            self.error_window(f"Error extracting data: {str(e)}")
    
    def _readout_csv_single(self, participant_id: str, all_data: Optional[Dict[str, Any]] = None) -> None:
        """Extract data for single participant to CSV (no preview)"""
        if all_data is None:
            all_data = self._prepare_data_for_extraction(participant_id)
        
        csv_path = self.excel_path  # Use the pre-selected masterfile path
        
        # Read existing CSV to get all fieldnames
        existing_fieldnames: List[str] = []
        file_exists = os.path.exists(csv_path) and os.path.getsize(csv_path) > 0
        if file_exists:
            with open(csv_path, 'r', newline='', encoding='utf-8') as f:
                reader = csv.DictReader(f)
                existing_fieldnames = list(reader.fieldnames) if reader.fieldnames else []
        
        # Merge fieldnames
        all_fieldnames = list(dict.fromkeys(existing_fieldnames + list(all_data.keys())))
        
        # Append to CSV
        with open(csv_path, 'a', newline='', encoding='utf-8') as f:
            writer = csv.DictWriter(f, fieldnames=all_fieldnames)
            if not file_exists:
                writer.writeheader()
            writer.writerow(all_data)
    
    
    def updateRecentConfigsMenu(self) -> None:
        """Update recent configurations menu with available configs"""
        self.recent_configs_menu.clear()
        self.delete_configs_menu.clear()
        
        config = ConfigParser()
        config.read('config.ini')
        sections = config.sections()
        
        if not sections:
            no_config_action = QAction("(No configurations available)", self)
            no_config_action.setEnabled(False)
            self.recent_configs_menu.addAction(no_config_action)  # type: ignore[misc]
            
            no_delete_action = QAction("(No configurations available)", self)
            no_delete_action.setEnabled(False)
            self.delete_configs_menu.addAction(no_delete_action)  # type: ignore[misc]
        else:
            for section in sections:
                # Recent configurations - load on click
                load_action = QAction(section, self)
                load_action.triggered.connect(lambda checked, name=section: self.ConfigLoad(name))  # type: ignore[misc]
                self.recent_configs_menu.addAction(load_action)  # type: ignore[misc]
                
                # Delete configurations - delete on click with confirmation
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
            
            with open('config.ini', 'w') as f:
                config.write(f)
            
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
                wb: xlw.Book = xlw.Book(self.excel_path)
                try:
                    ws = wb.sheets[0]  # type: ignore[misc]
                    existing_ids = ws.range('A2:A1000').value  # type: ignore[misc]
                    return participant_id in [str(cast(Any, id)).strip() for id in existing_ids if id]  # type: ignore[misc]
                finally:
                    wb.close()  # type: ignore[misc]
        except Exception:
            return False
        
        return False
    
    def checkDataCompleteness(self, participant_id: str) -> tuple[bool, List[str]]:
        """Check if participant folder has all expected CSV files"""
        participant_folder = os.path.join(self.source_path, participant_id)
        if not os.path.exists(participant_folder):
            return False, [f"Folder not found: {participant_folder}"]
        
        csv_files = sorted([f for f in os.listdir(participant_folder) if f.endswith('_Extract Data.csv')])
        
        if not csv_files:
            return False, ["No Extract Data CSV files found"]
        
        # Check for common incomplete patterns
        issues: List[str] = []
        expected_count = len(self.questionnaire_rows) if self.questionnaire_rows else None
        
        if expected_count and len(csv_files) < expected_count:
            issues.append(f"Expected {expected_count} CSV files, found {len(csv_files)}")
        
        return len(issues) == 0, issues if issues else csv_files
    
    def showPreviewDialog(self, data: Dict[str, Any]) -> bool:
        """Show preview dialog with data to be extracted. Returns True if user confirms."""
        dialog = QDialog(self)
        dialog.setWindowTitle("Preview Data Extraction")
        dialog.setModal(True)
        dialog.resize(800, 600)
        
        layout = QVBoxLayout()
        
        label = QLabel(f"Preview data for participant: {data.get('participant_id', 'Unknown')}")
        label.setStyleSheet("font-weight: bold; font-size: 14px;")
        layout.addWidget(label)
        
        # Create table
        table = QTableWidget()
        table.setColumnCount(2)
        table.setHorizontalHeaderLabels(["Field", "Value"])
        table.setRowCount(len(data))
        
        for i, (key, value) in enumerate(data.items()):
            table.setItem(i, 0, QTableWidgetItem(str(key)))
            table.setItem(i, 1, QTableWidgetItem(str(value)))
        
        table.horizontalHeader().setStretchLastSection(True)  # type: ignore[misc]
        table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeMode.ResizeToContents)  # type: ignore[misc]
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
        
        def confirm_and_close() -> None:
            result[0] = True
            dialog.accept()
        
        confirm_btn.clicked.connect(confirm_and_close)
        cancel_btn.clicked.connect(dialog.reject)
        
        dialog.exec()
        return result[0]
    
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
