# -*- coding: utf-8 -*-
"""
Survey Workbench - Enhanced Participant Data Management System
Version: 2.1 (Configuration Snapshot & Scoped Field Mapping Edition)

Author: Cagatay Gutt
Last Updated: 2026

Features:
- Config snapshotting: Saves template paths, copy counts, paths, AND complete column remappings into config.ini.
- Fully qualified mapping keys ({survey}_{field}) preventing cross-sheet name collisions (e.g. demographics_Check1 vs protocol_Check1).
- Dedicated sheets per root PDF in Excel with Column A as 'participant_id'.
- Natural numerical sorting (Text1, Text2... Text10) across all sheets and dialogs.
"""

import sys
import os
import shutil
import csv
import json
import re
from typing import Optional, List, cast, Dict, Any, Tuple
from PyQt5.QtWidgets import (QApplication, QMainWindow, QPushButton, 
                             QGridLayout, QWidget, QLineEdit, QLabel,
                             QFileDialog, QMessageBox, QAction, QStatusBar,
                             QScrollArea, QVBoxLayout, QFrame,
                             QMenu, QDialog, QWhatsThis, QTextEdit, QTableWidget,
                             QTableWidgetItem, QHeaderView, QCheckBox, QInputDialog)
from configparser import ConfigParser
from PyQt5.QtCore import pyqtSignal, Qt
import xlwings as xlw  # type: ignore[import]
from pypdf import PdfReader, PdfWriter
from pypdf.generic import NameObject, create_string_object, ArrayObject

# Canonical grammar for all survey presentations (Section 7.4)
CANONICAL_PATTERN = re.compile(r'^(Text|Check)(\d+)(?:_(\d+))?(_n)?$', re.IGNORECASE)


def natural_sort_key(s: Any) -> List[Any]:
    """Naturally sort alphanumeric strings so Text2 comes before Text10."""
    return [int(text) if text.isdigit() else text.lower() for text in re.split(r'(\d+)', str(s))]


def resolve_pdf_object(obj: Any) -> Any:
    """Recursively resolve IndirectObject references from Acrobat, Word, and LaTeX forms."""
    max_depth = 10
    depth = 0
    while hasattr(obj, 'get_object') and depth < max_depth:
        try:
            resolved = obj.get_object()
            if resolved is None or resolved is obj:
                break
            obj = resolved
            depth += 1
        except Exception:
            break
    return obj


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
        self.name = self.configset.text().strip()
        if not self.name:
            return
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
        return [
            self.name_label, self.name_edit,
            self.path_label, self.path_display, self.path_button,
            self.count_label, self.count_edit
        ]


class MainWindow(QMainWindow):
    def __init__(self, parent: Optional[QWidget] = None) -> None:
        super().__init__(parent)
        self.setWindowTitle("Survey Workbench v2.1")
        
        self.questionnaire_rows: List[QuestionnaireRow] = []
        self.recent_configs_menu: Optional[QMenu] = None
        self.delete_configs_menu: Optional[QMenu] = None
        
        main_widget = QWidget()
        main_layout = QVBoxLayout()
        
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
        file_menu.addAction(button_action)
        
        self.recent_configs_menu = file_menu.addMenu("&Load Configuration")
        self.delete_configs_menu = file_menu.addMenu("&Delete Configuration")
        file_menu.aboutToShow.connect(self.updateRecentConfigsMenu)

        edit_mapping_action: QAction = QAction("&Save/Edit Field Mapping", self)
        edit_mapping_action.triggered.connect(self.editFieldMappingFromExtraction)
        edit_mapping_action.setToolTip("Edit and save the column name remapping for the active configuration")
        file_menu.addAction(edit_mapping_action)
        
        help_menu = menubar.addMenu("&HELP")
        assert help_menu is not None
        help_menu.addAction(whats_this_action)
        help_menu.addAction(button_action3)
        
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
        
        self.batch_mode_gen = QCheckBox("Batch Mode (multiple IDs)")
        self.batch_mode_gen.toggled.connect(self.toggleBatchModeGeneration)
        
        self.batch_text_gen = QTextEdit()
        self.batch_text_gen.setPlaceholderText("Enter participant IDs (one per line or comma-separated):\nP001\nP002\nP003")
        self.batch_text_gen.setVisible(False)
        self.batch_text_gen.setMaximumHeight(100)
        
        target_label = QLabel("Target folder PATH:")
        self.target_pathset = QLabel()
        target_pathbut = QPushButton("Search")
        target_pathbut.clicked.connect(self.select_target_folder)
        
        quest_count_label = QLabel("Number of questionnaire types:")
        self.quest_count_edit = QLineEdit()
        quest_count_button = QPushButton("Confirm")
        quest_count_button.clicked.connect(self.create_questionnaire_rows)
        
        gen_layout.addWidget(gen_title, 0, 0, 1, 3)
        gen_layout.addWidget(part_label, 1, 0)
        gen_layout.addWidget(self.nameset, 1, 1)
        gen_layout.addWidget(self.batch_mode_gen, 1, 2)
        
        import_list_btn_gen = QPushButton("Import List...")
        import_list_btn_gen.clicked.connect(lambda: self.importParticipantList(self.batch_text_gen))
        gen_layout.addWidget(import_list_btn_gen, 2, 2)
        
        gen_layout.addWidget(self.batch_text_gen, 2, 0, 1, 2)
        gen_layout.addWidget(target_label, 3, 0)
        gen_layout.addWidget(self.target_pathset, 3, 1)
        gen_layout.addWidget(target_pathbut, 3, 2)
        gen_layout.addWidget(quest_count_label, 4, 0)
        gen_layout.addWidget(self.quest_count_edit, 4, 1)
        gen_layout.addWidget(quest_count_button, 4, 2)
        
        save_bundle_btn = QPushButton("Save Template Bundle")
        save_bundle_btn.clicked.connect(self.saveTemplateBundle)
        gen_layout.addWidget(save_bundle_btn, 5, 0)
        
        load_bundle_btn = QPushButton("Load Template Bundle")
        load_bundle_btn.clicked.connect(self.loadTemplateBundle)
        gen_layout.addWidget(load_bundle_btn, 5, 1)
        
        gen_frame.setLayout(gen_layout)
        main_layout.addWidget(gen_frame)
        
        self.scroll_area = QScrollArea()
        self.scroll_area.setWidgetResizable(True)
        self.scroll_widget = QWidget()
        self.scroll_layout = QGridLayout()
        self.scroll_widget.setLayout(self.scroll_layout)
        self.scroll_area.setWidget(self.scroll_widget)
        self.scroll_area.setVisible(False)
        main_layout.addWidget(self.scroll_area)
        
        self.generate_button = QPushButton("Generate Participant Folder")
        self.generate_button.setStyleSheet("font-weight: bold; padding: 10px;")
        self.generate_button.clicked.connect(self.generate_participant_folder)
        main_layout.addWidget(self.generate_button)
        
        # ========== EXTRACTION SECTION ==========
        extract_frame = QFrame()
        extract_frame.setFrameStyle(QFrame.Shape.Box | QFrame.Shadow.Raised)
        extract_layout = QGridLayout()
        
        extract_title = QLabel("Extract Data to Excel")
        extract_title.setStyleSheet("font-weight: bold; font-size: 14px;")
        
        source_label = QLabel("Source folder PATH:")
        self.source_pathset = QLabel()
        source_pathbut = QPushButton("Search")
        source_pathbut.clicked.connect(self.select_source_folder)
        
        excel_label = QLabel("Masterfile PATH:")
        self.excel_pathset = QLabel()
        excel_pathbut = QPushButton("Search")
        excel_pathbut.clicked.connect(self.select_excel_file)
        
        extract_layout.addWidget(extract_title, 0, 0, 1, 3)
        extract_layout.addWidget(source_label, 1, 0)
        extract_layout.addWidget(self.source_pathset, 1, 1)
        extract_layout.addWidget(source_pathbut, 1, 2)
        extract_layout.addWidget(excel_label, 2, 0)
        extract_layout.addWidget(self.excel_pathset, 2, 1)
        extract_layout.addWidget(excel_pathbut, 2, 2)
        
        self.batch_mode_extract = QCheckBox("Batch Mode (multiple IDs)")
        self.batch_mode_extract.toggled.connect(self.toggleBatchModeExtraction)
        extract_layout.addWidget(self.batch_mode_extract, 3, 0)
        
        import_list_btn_extract = QPushButton("Import List...")
        import_list_btn_extract.clicked.connect(lambda: self.importParticipantList(self.batch_text_extract))
        extract_layout.addWidget(import_list_btn_extract, 3, 2)
        
        self.batch_text_extract = QTextEdit()
        self.batch_text_extract.setPlaceholderText("Enter participant IDs (one per line or comma-separated):\nP001\nP002\nP003")
        self.batch_text_extract.setVisible(False)
        self.batch_text_extract.setMaximumHeight(100)
        extract_layout.addWidget(self.batch_text_extract, 4, 0, 1, 2)
        
        missing_data_btn = QPushButton("Missing Data Report")
        missing_data_btn.clicked.connect(self.generateMissingDataReport)
        extract_layout.addWidget(missing_data_btn, 5, 2)
        
        extract_frame.setLayout(extract_layout)
        main_layout.addWidget(extract_frame)
        
        self.extract_button = QPushButton("Extract Data to Masterfile")
        self.extract_button.setStyleSheet("font-weight: bold; padding: 10px;")
        self.extract_button.clicked.connect(self.extract_data)
        main_layout.addWidget(self.extract_button)
        
        main_widget.setLayout(main_layout)
        self.setCentralWidget(main_widget)
        
        self.target_path: str = ""
        self.source_path: str = ""
        self.excel_path: str = ""
        self.active_config_name: str = ""
        self.field_mapping: Dict[str, str] = {}
        
        self.resize(720, 620)

    def get_configured_root_order(self) -> List[str]:
        """Return root questionnaires in the exact order configured in the UI."""
        order = []
        for r in self.questionnaire_rows:
            name = r.name_edit.text().strip()
            if name and name not in order:
                order.append(name)
        return order

    def _get_root_survey_name(self, clean_filename: str) -> str:
        """Derive the root questionnaire type (e.g. 'nasa1' -> 'nasa', 'protocol' -> 'protocol')."""
        clean = clean_filename.strip()
        for s_name in self.get_configured_root_order():
            if clean == s_name or clean.startswith(s_name):
                return s_name
        m = re.match(r'^(.*?)\d+$', clean)
        if m and m.group(1):
            return m.group(1).rstrip('_')
        return clean

    def create_questionnaire_rows(self) -> None:
        try:
            count: int = int(self.quest_count_edit.text())
            if count <= 0:
                raise ValueError
            
            for row in self.questionnaire_rows:
                for widget in row.get_widgets():
                    widget.deleteLater()
            self.questionnaire_rows.clear()
            
            for i in range(count):
                row: QuestionnaireRow = QuestionnaireRow(i)
                self.questionnaire_rows.append(row)
                
                base_row = i * 3
                self.scroll_layout.addWidget(row.name_label, base_row, 0)
                self.scroll_layout.addWidget(row.name_edit, base_row, 1, 1, 2)
                
                self.scroll_layout.addWidget(row.path_label, base_row + 1, 0)
                self.scroll_layout.addWidget(row.path_display, base_row + 1, 1)
                self.scroll_layout.addWidget(row.path_button, base_row + 1, 2)
                
                self.scroll_layout.addWidget(row.count_label, base_row + 2, 0)
                self.scroll_layout.addWidget(row.count_edit, base_row + 2, 1)
                
                if i < count - 1:
                    line = QFrame()
                    line.setFrameShape(QFrame.Shape.HLine)
                    line.setFrameShadow(QFrame.Shadow.Sunken)
                    self.scroll_layout.addWidget(line, base_row + 3, 0, 1, 3)
                
                row.path_button.clicked.connect(
                    (lambda r: lambda: self.select_template_file(r))(row)
                )
            
            self.scroll_area.setVisible(True)
        except ValueError:
            self.error_window("Please enter a valid positive number!")
    
    def select_template_file(self, row: QuestionnaireRow) -> None:
        file_path = QFileDialog.getOpenFileName(
            self, "Select Template File", "", "PDF Files (*.pdf);;All Files (*.*)"
        )[0]
        if file_path:
            row.template_path = file_path
            row.path_display.setText(file_path)
    
    def select_target_folder(self) -> None:
        if folder_path := QFileDialog.getExistingDirectory(self, "Select Target Folder"):
            self.target_path = folder_path
            self.target_pathset.setText(folder_path)
    
    def select_source_folder(self) -> None:
        if folder_path := QFileDialog.getExistingDirectory(self, "Select Source Folder"):
            self.source_path = folder_path
            self.source_pathset.setText(folder_path)
    
    def select_excel_file(self) -> None:
        if file_path := QFileDialog.getOpenFileName(
            self, "Select Masterfile", "", 
            "All Supported (*.csv *.xls *.xlsx);;CSV Files (*.csv);;Excel Files (*.xls *.xlsx);;All Files (*.*)"
        )[0]:
            self.excel_path = file_path
            self.excel_pathset.setText(file_path)
    
    def generate_participant_folder(self) -> None:
        try:
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
        if not self.target_path:
            raise ValueError("Please select a target folder!")
        if not self.questionnaire_rows:
            raise ValueError("Please configure questionnaires!")
        
        participant_folder = os.path.join(self.target_path, participant_id)
        if os.path.exists(participant_folder):
            shutil.rmtree(participant_folder)
        os.makedirs(participant_folder)

        generated_pdf_paths: List[str] = []
        for row in self.questionnaire_rows:
            if not row.template_path:
                continue
            
            survey_name = row.name_edit.text().strip() or f"survey_{row.index + 1}"
            try:
                copy_count = int(row.count_edit.text())
            except ValueError:
                copy_count = 1
            
            _, ext = os.path.splitext(os.path.basename(row.template_path))
            for i in range(copy_count):
                new_filename = f"{participant_id}_{survey_name}{ext}" if copy_count == 1 else f"{participant_id}_{survey_name}{i + 1}{ext}"
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
        reader = PdfReader(pdf_path)
        writer = PdfWriter()
        writer.append(reader)

        def _update_field_tree(fields_array: Any) -> None:
            for ref in fields_array:
                try:
                    obj = resolve_pdf_object(ref)
                    name_raw = resolve_pdf_object(obj.get('/T'))
                    if name_raw is not None:
                        name_str = str(name_raw)
                        if name_str in field_values:
                            new_val = str(field_values[name_str])
                            ft = str(resolve_pdf_object(obj.get('/FT', '')))
                            if ft == '/Btn':
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
                    kids = resolve_pdf_object(obj.get('/Kids'))
                    if isinstance(kids, list):
                        _update_field_tree(kids)
                except Exception:
                    continue

        try:
            root_obj = resolve_pdf_object(cast(Any, writer).trailer['/Root'])
            acroform_ref = root_obj.get('/AcroForm')
            if acroform_ref is not None:
                acroform = resolve_pdf_object(acroform_ref)
                fields_raw = resolve_pdf_object(acroform.get('/Fields'))
                if isinstance(fields_raw, list):
                    _update_field_tree(fields_raw)
        except Exception:
            pass

        for page in writer.pages:
            try:
                writer.update_page_form_field_values(page, field_values)
            except Exception:
                pass

        writer.set_need_appearances_writer()
        with open(pdf_path, "wb") as output_file:
            writer.write(output_file)

    def _extract_pdf_form_fields(self, pdf_path: str) -> Dict[str, str]:
        """Extract all form fields from any PDF, resolving matrices, radio buttons, and text cells."""
        try:
            reader = PdfReader(pdf_path)
        except Exception:
            return {}

        extracted: Dict[str, str] = {}

        def clean_value(raw: Any) -> str:
            resolved = resolve_pdf_object(raw)
            if resolved is None:
                return ''
            if isinstance(resolved, ArrayObject):
                return ', '.join([clean_value(x) for x in resolved if x is not None])
            text = str(resolved).strip()
            if text.startswith('/'):
                unslashed = text[1:].strip()
                if unslashed.lower() in ('yes', 'on', 'true', '1'):
                    return '1'
                elif unslashed.lower() in ('off', 'false', '0'):
                    return '0'
                return unslashed
            return text

        def store_field(name: Any, value: Any) -> None:
            if not name:
                return
            key = str(name).strip()
            if not key or key == 'File':
                return
            val_str = clean_value(value)
            if key not in extracted:
                extracted[key] = val_str
            elif extracted[key] in ('', '0') and val_str not in ('', '0'):
                extracted[key] = val_str

        def walk_acroform_fields(fields_list: Any, parent_name: str = "") -> None:
            if not isinstance(fields_list, list):
                return
            for field_ref in fields_list:
                try:
                    f = resolve_pdf_object(field_ref)
                    if not isinstance(f, dict):
                        continue
                    local_name = resolve_pdf_object(f.get('/T'))
                    name = str(local_name).strip() if local_name else parent_name
                    
                    val = resolve_pdf_object(f.get('/V'))
                    if val is None or str(val).strip() in ('', '/Off'):
                        val = resolve_pdf_object(f.get('/AS'))
                    
                    if name:
                        store_field(name, val)
                    
                    kids = resolve_pdf_object(f.get('/Kids'))
                    if isinstance(kids, list):
                        walk_acroform_fields(kids, name)
                except Exception:
                    continue

        try:
            root = resolve_pdf_object(reader.trailer.get('/Root', {}))
            acro_ref = root.get('/AcroForm')
            if acro_ref is not None:
                acro = resolve_pdf_object(acro_ref)
                fields_raw = resolve_pdf_object(acro.get('/Fields'))
                if isinstance(fields_raw, list):
                    walk_acroform_fields(fields_raw)
        except Exception:
            pass

        try:
            acro_fields = reader.get_fields() or {}
            for fname, fmeta in acro_fields.items():
                if isinstance(fmeta, dict):
                    val = resolve_pdf_object(fmeta.get('/V'))
                    if val is None or str(val).strip() in ('', '/Off'):
                        val = resolve_pdf_object(fmeta.get('/AS'))
                    store_field(fname, val)
        except Exception:
            pass

        try:
            for page in reader.pages:
                annots = resolve_pdf_object(page.get('/Annots'))
                if not isinstance(annots, list):
                    continue
                for ref in annots:
                    try:
                        annot = resolve_pdf_object(ref)
                        if not isinstance(annot, dict) or str(annot.get('/Subtype', '')) != '/Widget':
                            continue
                        fname_raw = resolve_pdf_object(annot.get('/T'))
                        value_raw = resolve_pdf_object(annot.get('/V'))
                        if value_raw is None or str(value_raw).strip() in ('', '/Off'):
                            value_raw = resolve_pdf_object(annot.get('/AS'))
                        
                        parent = resolve_pdf_object(annot.get('/Parent'))
                        if isinstance(parent, dict):
                            if fname_raw is None:
                                fname_raw = resolve_pdf_object(parent.get('/T'))
                            if value_raw is None or str(value_raw).strip() in ('', '/Off'):
                                value_raw = resolve_pdf_object(parent.get('/V')) or resolve_pdf_object(parent.get('/AS'))
                        
                        store_field(fname_raw, value_raw)
                    except Exception:
                        continue
        except Exception:
            pass

        return extracted

    def _build_prefill_defaults(self, participant_id: str, fields: Dict[str, Any]) -> Dict[str, Any]:
        defaults: Dict[str, Any] = dict(fields)
        participant_aliases = {"participant_id", "participant-id", "Text13_n", "Text13"}
        for field_name in fields.keys():
            if field_name in participant_aliases and str(fields.get(field_name, "")).strip() == "":
                defaults[field_name] = participant_id
        return defaults

    def showGenerationPrefillDialog(self, participant_id: str, generated_pdf_paths: List[str]) -> tuple[bool, Dict[str, Dict[str, Any]]]:
        dialog = QDialog(self)
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
        
        root_order = self.get_configured_root_order()
        sorted_pdf_paths = sorted(
            generated_pdf_paths,
            key=lambda p: (
                root_order.index(self._get_root_survey_name(os.path.splitext(os.path.basename(p))[0].replace(f'{participant_id}_', '', 1)))
                if self._get_root_survey_name(os.path.splitext(os.path.basename(p))[0].replace(f'{participant_id}_', '', 1)) in root_order
                else 9999,
                natural_sort_key(os.path.basename(p))
            )
        )

        for pdf_path in sorted_pdf_paths:
            file_label = os.path.basename(pdf_path)
            survey_type = os.path.splitext(file_label)[0].replace(f'{participant_id}_', '', 1)
            extracted_fields = self._extract_pdf_form_fields(pdf_path)
            if not extracted_fields:
                continue
            defaults = self._build_prefill_defaults(participant_id, extracted_fields)
            file_to_extracted_fields[file_label] = extracted_fields

            preferred_fields: Dict[str, str] = {}
            for k, v in defaults.items():
                if k.endswith('_n'):
                    base = k[:-2]
                    if base not in defaults:
                        preferred_fields[base] = str(v)
                else:
                    preferred_fields[k] = str(v)

            check_matrix_groups: Dict[str, Dict[str, str]] = {}
            text_fields: Dict[str, str] = {}
            
            for field_name in sorted(preferred_fields.keys(), key=natural_sort_key):
                val = preferred_fields[field_name]
                m = CANONICAL_PATTERN.match(field_name)
                if m:
                    kind, group_x, opt_n = m.group(1).capitalize(), m.group(2), m.group(3)
                    if kind == 'Check' and opt_n is not None:
                        base_name = f"Check{group_x}"
                        check_matrix_groups.setdefault(base_name, {})[opt_n] = val
                    else:
                        text_fields[field_name] = val
                else:
                    text_fields[field_name] = val

            for f_name, f_val in sorted(text_fields.items(), key=lambda it: natural_sort_key(it[0])):
                system_key = f"{survey_type}_{f_name}"
                display_name = self.field_mapping.get(system_key, self.field_mapping.get(f_name, f_name))
                rows.append((f"{file_label} :: {f_name}", display_name, f_val))

            for base_name, options in sorted(check_matrix_groups.items(), key=lambda it: natural_sort_key(it[0])):
                system_key = f"{survey_type}_{base_name}"
                display_name = self.field_mapping.get(system_key, self.field_mapping.get(base_name, base_name))
                checked_opt = next((opt for opt, cval in sorted(options.items(), key=lambda it: natural_sort_key(it[0])) if cval == '1'), '')
                rows.append((f"{file_label} :: {base_name}", display_name, checked_opt))

        table.setRowCount(len(rows))
        for idx, (combined_field, display_name, value) in enumerate(rows):
            field_item = QTableWidgetItem(combined_field)
            field_item.setFlags(field_item.flags() & Qt.ItemFlag.NoItemFlags)
            name_item = QTableWidgetItem(display_name)
            name_item.setFlags(name_item.flags() | Qt.ItemFlag.ItemIsEditable)
            table.setItem(idx, 0, field_item)
            table.setItem(idx, 1, name_item)
            table.setItem(idx, 2, QTableWidgetItem(value))

        table.horizontalHeader().setStretchLastSection(True)
        table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeMode.ResizeToContents)
        table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeMode.ResizeToContents)
        layout.addWidget(table)

        button_layout = QVBoxLayout()
        confirm_btn = QPushButton("Confirm and Generate")
        cancel_btn = QPushButton("Cancel")
        button_layout.addWidget(confirm_btn)
        button_layout.addWidget(cancel_btn)
        layout.addLayout(button_layout)
        dialog.setLayout(layout)

        result: Dict[str, Any] = {"action": "cancel", "values": {}}

        def confirm_and_close() -> None:
            prefill_values: Dict[str, Dict[str, Any]] = {p: {} for p in generated_pdf_paths}
            
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
                pdf_path = file_to_path.get(file_name)
                val = value_item.text().strip() if value_item else ""
                
                survey_type = os.path.splitext(file_name)[0].replace(f'{participant_id}_', '', 1)
                new_name = name_item.text().strip()
                if new_name:
                    self.field_mapping[f"{survey_type}_{field_name}"] = new_name

                if not pdf_path:
                    continue

                if field_name.startswith('Check') and '_' not in field_name:
                    base_name = field_name
                    raw_fields = file_to_extracted_fields.get(file_name, {})
                    for raw_f in raw_fields:
                        m = CANONICAL_PATTERN.match(raw_f)
                        if m and f"Check{m.group(2)}" == base_name and m.group(3) is not None:
                            col_opt = m.group(3)
                            prefill_values[pdf_path][raw_f] = '1' if col_opt == val else '0'
                else:
                    prefill_values[pdf_path][field_name] = val

            # Persist updated mapping if active config is loaded
            if self.active_config_name:
                self._save_mapping_to_config(self.active_config_name, self.field_mapping)

            result["action"] = "apply"
            result["values"] = prefill_values
            dialog.accept()

        confirm_btn.clicked.connect(confirm_and_close)
        cancel_btn.clicked.connect(dialog.reject)
        dialog.exec()

        if result["action"] == "apply":
            return True, cast(Dict[str, Dict[str, Any]], result["values"])
        raise ValueError("Generation cancelled by user")

    def _prepare_data_for_extraction(self, participant_id: str) -> Dict[str, Dict[str, Any]]:
        """Extract data segregated by root sheet name, preserving root sequence and numerical order."""
        normalized_source = os.path.normpath(self.source_path)
        participant_folder = self.source_path if normalized_source.endswith(participant_id) else os.path.join(self.source_path, participant_id)
        
        if not os.path.exists(participant_folder):
            raise ValueError(f"Participant folder not found: {participant_folder}")

        survey_files: List[str] = sorted([f for f in os.listdir(participant_folder) if f.lower().endswith('.pdf')], key=natural_sort_key)
        if not survey_files:
            raise ValueError("No readable questionnaire data found.")

        sheet_data: Dict[str, Dict[str, Any]] = {}

        for pdf_file in survey_files:
            clean_type = os.path.splitext(pdf_file)[0].replace(f"{participant_id}_", "", 1)
            fields = self._extract_pdf_form_fields(os.path.join(participant_folder, pdf_file))
            if not fields:
                continue

            root_sheet = self._get_root_survey_name(clean_type)
            if root_sheet not in sheet_data:
                sheet_data[root_sheet] = {'participant_id': participant_id}

            has_multiple_copies = any(
                clean_type != self._get_root_survey_name(os.path.splitext(f)[0].replace(f"{participant_id}_", "", 1))
                for f in survey_files if self._get_root_survey_name(os.path.splitext(f)[0].replace(f"{participant_id}_", "", 1)) == root_sheet
            ) or clean_type != root_sheet

            prefix = f"{clean_type}_" if has_multiple_copies else ""

            preferred_fields: Dict[str, str] = {}
            for k, v in fields.items():
                if k.endswith('_n'):
                    base = k[:-2]
                    if base not in fields:
                        preferred_fields[base] = str(v)
                else:
                    preferred_fields[k] = str(v)

            check_groups: Dict[str, Dict[str, str]] = {}
            
            for raw_k, v in sorted(preferred_fields.items(), key=lambda it: natural_sort_key(it[0])):
                val = str(v).strip()
                m = CANONICAL_PATTERN.match(raw_k)
                if m:
                    kind = m.group(1).capitalize()
                    group_x = m.group(2)
                    opt_n = m.group(3)

                    if kind == 'Check' and opt_n is not None:
                        base_name = f"Check{group_x}"
                        check_groups.setdefault(base_name, {})[opt_n] = val
                    else:
                        sheet_data[root_sheet][f"{prefix}{raw_k}"] = val
                else:
                    sheet_data[root_sheet][f"{prefix}{raw_k}"] = val

            for base_name, options in sorted(check_groups.items(), key=lambda it: natural_sort_key(it[0])):
                checked_opt = next((opt for opt, cval in sorted(options.items(), key=lambda it: natural_sort_key(it[0])) if str(cval).strip() == '1'), '')
                sheet_data[root_sheet][f"{prefix}{base_name}"] = checked_opt

        return sheet_data

    def _readout_single(self, participant_id: str, sheet_data: Optional[Dict[str, Dict[str, Any]]] = None) -> None:
        """Route each root PDF's extracted data to its own dedicated worksheet in root order."""
        if sheet_data is None:
            sheet_data = self._prepare_data_for_extraction(participant_id)

        file_ext = self.excel_path.lower().split('.')[-1]
        root_order = self.get_configured_root_order()

        ordered_sheet_names = sorted(
            sheet_data.keys(),
            key=lambda s: (root_order.index(s) if s in root_order else 9999, natural_sort_key(s))
        )

        if file_ext in ['xls', 'xlsx']:
            wb: xlw.Book = xlw.Book(self.excel_path)
            try:
                for sheet_name in ordered_sheet_names:
                    raw_data = sheet_data[sheet_name]
                    export_data: Dict[str, Any] = {'participant_id': participant_id}
                    
                    sorted_field_keys = sorted(
                        [k for k in raw_data.keys() if k != 'participant_id'],
                        key=natural_sort_key
                    )

                    for key in sorted_field_keys:
                        value = raw_data[key]
                        # Look up candidate keys: exact scoped key, sheet_scoped key, or base key
                        col_name = self.field_mapping.get(key)
                        if not col_name:
                            col_name = self.field_mapping.get(f"{sheet_name}_{key}")
                        if not col_name:
                            for map_k, map_v in self.field_mapping.items():
                                if key.endswith(f"_{map_k}") or map_k.endswith(f"_{key}"):
                                    col_name = map_v
                                    break
                        export_data[col_name or key] = value

                    sheet_names = [cast(str, s.name) for s in wb.sheets]
                    if sheet_name in sheet_names:
                        sheet: xlw.Sheet = cast(xlw.Sheet, wb.sheets[sheet_name])
                    else:
                        if len(wb.sheets) == 1 and wb.sheets[0].range('A1').value is None:
                            sheet = cast(xlw.Sheet, wb.sheets[0])
                            sheet.name = sheet_name
                        else:
                            sheet = cast(xlw.Sheet, wb.sheets.add(name=sheet_name, after=wb.sheets[-1]))

                    first_row = sheet.range('1:1').value
                    header_map: Dict[str, int] = {}
                    if first_row:
                        row_vals = first_row if isinstance(first_row, list) else [first_row]
                        for c_idx, val in enumerate(row_vals, start=1):
                            if val is not None and str(val).strip():
                                header_map[str(val).strip()] = c_idx

                    if not header_map:
                        headers = list(export_data.keys())
                        for idx, header in enumerate(headers, start=1):
                            sheet.range(1, idx).value = header
                        header_map = {h: idx for idx, h in enumerate(headers, start=1)}
                    else:
                        next_col = max(header_map.values(), default=0) + 1
                        for key in export_data.keys():
                            if key not in header_map:
                                sheet.range(1, next_col).value = key
                                header_map[key] = next_col
                                next_col += 1

                    next_row = 2
                    while sheet.range(f'A{next_row}').value is not None:
                        next_row += 1

                    for key, val in export_data.items():
                        col_idx = header_map.get(key)
                        if col_idx:
                            sheet.range(next_row, col_idx).value = val

                cast(xlw.Book, wb).save()
            except Exception:
                if 'wb' in locals():
                    wb.close()
                raise
        elif file_ext == 'csv':
            base_dir = os.path.dirname(self.excel_path)
            base_name = os.path.splitext(os.path.basename(self.excel_path))[0]
            
            for sheet_name in ordered_sheet_names:
                raw_data = sheet_data[sheet_name]
                target_csv = os.path.join(base_dir, f"{base_name}_{sheet_name}.csv")
                file_exists = os.path.exists(target_csv) and os.path.getsize(target_csv) > 0

                export_data = {'participant_id': participant_id}
                sorted_field_keys = sorted(
                    [k for k in raw_data.keys() if k != 'participant_id'],
                    key=natural_sort_key
                )
                for k in sorted_field_keys:
                    mapped_k = self.field_mapping.get(k, self.field_mapping.get(f"{sheet_name}_{k}", k))
                    export_data[mapped_k] = raw_data[k]

                existing_headers: List[str] = []
                existing_rows: List[Dict[str, Any]] = []
                if file_exists:
                    with open(target_csv, 'r', newline='', encoding='utf-8') as f:
                        reader = csv.DictReader(f)
                        existing_headers = list(reader.fieldnames) if reader.fieldnames else []
                        existing_rows = list(reader)

                new_headers = [k for k in export_data.keys() if k not in existing_headers]
                all_headers = existing_headers + new_headers if existing_headers else list(export_data.keys())

                if new_headers and file_exists:
                    with open(target_csv, 'w', newline='', encoding='utf-8') as f:
                        writer = csv.DictWriter(f, fieldnames=all_headers)
                        writer.writeheader()
                        writer.writerows(existing_rows)
                        writer.writerow(export_data)
                else:
                    with open(target_csv, 'a', newline='', encoding='utf-8') as f:
                        writer = csv.DictWriter(f, fieldnames=all_headers)
                        if not file_exists:
                            writer.writeheader()
                        writer.writerow(export_data)
        else:
            raise ValueError(f"Unsupported masterfile format: {file_ext}")

    def extract_data(self) -> None:
        if not self.excel_path:
            return self.error_window("Please select a masterfile!")

        self.field_mapping = self._load_field_mapping_for_extraction()
        
        if self.batch_mode_extract.isChecked():
            participant_ids = self.parseParticipantIDs(self.batch_text_extract.toPlainText())
            if not participant_ids:
                return self.error_window("Please enter at least one participant ID!")
            
            success_count = 0
            failed: List[str] = []
            duplicates: List[str] = []
            
            for pid in participant_ids:
                if self.checkDuplicate(pid):
                    duplicates.append(pid)
                    continue
                try:
                    self._readout_single(pid)
                    success_count += 1
                except Exception as e:
                    failed.append(f"{pid}: {str(e)}")
            
            msg = f"Extracted {success_count} of {len(participant_ids)} participants successfully!"
            if duplicates:
                msg += f"\n\nSkipped duplicates:\n" + "\n".join(duplicates)
            if failed:
                msg += f"\n\nFailed:\n" + "\n".join(failed)
            QMessageBox.information(self, "Batch Extraction Complete", msg)
        else:
            participant_id = self.nameset.text().strip()
            if not participant_id:
                return self.error_window("Please enter a participant ID!")
            
            if self.checkDuplicate(participant_id):
                if QMessageBox.question(self, "Duplicate Detected", "Extract anyway?",
                                        cast(QMessageBox.StandardButton, QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)) == QMessageBox.StandardButton.No:
                    return
            self.readout()
    
    def readout(self) -> None:
        try:
            self.field_mapping = self._load_field_mapping_for_extraction()
            participant_id = self.nameset.text().strip()
            sheet_data = self._prepare_data_for_extraction(participant_id)
            
            confirmed, edited_sheet_data = self.showPreviewDialog(participant_id, sheet_data)
            if not confirmed:
                return
            
            self._readout_single(participant_id, edited_sheet_data)
            QMessageBox.information(self, "Success", f"Data extracted successfully into sheets for {participant_id}!")
        except Exception as e:
            self.error_window(f"Error extracting data: {str(e)}")

    def showPreviewDialog(self, participant_id: str, sheet_data: Dict[str, Dict[str, Any]]) -> Tuple[bool, Dict[str, Dict[str, Any]]]:
        """4-column preview ordered first by root-PDF sequence, then numerically."""
        dialog = QDialog(self)
        dialog.setWindowTitle("Preview Data Extraction")
        dialog.setModal(True)
        dialog.resize(850, 600)
        
        layout = QVBoxLayout()
        label = QLabel(f"Preview data for participant: {participant_id}")
        label.setStyleSheet("font-weight: bold; font-size: 14px;")
        layout.addWidget(label)
        
        root_order = self.get_configured_root_order()
        sorted_sheet_names = sorted(
            sheet_data.keys(),
            key=lambda s: (root_order.index(s) if s in root_order else 9999, natural_sort_key(s))
        )

        flat_rows: List[Tuple[str, str, str, str]] = []
        for sheet_name in sorted_sheet_names:
            data = sheet_data[sheet_name]
            sorted_fields = sorted([k for k in data.keys() if k != 'participant_id'], key=natural_sort_key)
            for k in sorted_fields:
                v = data[k]
                mapped_name = self.field_mapping.get(k, self.field_mapping.get(f"{sheet_name}_{k}", k))
                flat_rows.append((sheet_name, k, mapped_name, str(v)))

        table = QTableWidget()
        table.setColumnCount(4)
        table.setHorizontalHeaderLabels(["Sheet", "Field", "Name", "Data"])
        table.setRowCount(len(flat_rows))
        
        for idx, (s_name, f_key, m_name, val) in enumerate(flat_rows):
            si = QTableWidgetItem(s_name)
            si.setFlags(Qt.ItemFlag.ItemIsEnabled)
            fi = QTableWidgetItem(f_key)
            fi.setFlags(Qt.ItemFlag.ItemIsEnabled)
            ni = QTableWidgetItem(m_name)
            ni.setFlags(ni.flags() | Qt.ItemFlag.ItemIsEditable)
            di = QTableWidgetItem(val)
            di.setFlags(di.flags() | Qt.ItemFlag.ItemIsEditable)

            table.setItem(idx, 0, si)
            table.setItem(idx, 1, fi)
            table.setItem(idx, 2, ni)
            table.setItem(idx, 3, di)
        
        table.horizontalHeader().setStretchLastSection(True)
        table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeMode.ResizeToContents)
        table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeMode.ResizeToContents)
        table.horizontalHeader().setSectionResizeMode(2, QHeaderView.ResizeMode.ResizeToContents)
        layout.addWidget(table)
        
        btn_layout = QVBoxLayout()
        c_btn = QPushButton("Confirm and Extract")
        can_btn = QPushButton("Cancel")
        btn_layout.addWidget(c_btn)
        btn_layout.addWidget(can_btn)
        layout.addLayout(btn_layout)
        dialog.setLayout(layout)
        
        result = [False]
        edited_output: Dict[str, Dict[str, Any]] = {s: {'participant_id': participant_id} for s in sheet_data}
        
        def save_and_close() -> None:
            for r in range(table.rowCount()):
                s = table.item(r, 0).text()
                k = table.item(r, 1).text()
                new_n = table.item(r, 2).text().strip() if table.item(r, 2) else ""
                val = table.item(r, 3).text() if table.item(r, 3) else ""
                
                edited_output.setdefault(s, {'participant_id': participant_id})[k] = val
                if new_n:
                    self.field_mapping[k] = new_n
                    self.field_mapping[f"{s}_{k}"] = new_n

            # Auto-persist updated mapping to active config
            if self.active_config_name:
                self._save_mapping_to_config(self.active_config_name, self.field_mapping)

            result[0] = True
            dialog.accept()
        
        c_btn.clicked.connect(save_and_close)
        can_btn.clicked.connect(dialog.reject)
        dialog.exec()
        return result[0], edited_output if result[0] else sheet_data

    def _get_expected_system_fields(self) -> set[str]:
        """Auto-generate expected field mapping with fully qualified keys: {survey_instance}_{canonical_field}."""
        expected: set[str] = set()

        for row in self.questionnaire_rows:
            template = row.template_path.strip()
            if not template or not template.lower().endswith('.pdf'):
                continue

            template_fields = self._extract_pdf_form_fields(template)
            if not template_fields:
                continue

            survey_name = row.name_edit.text().strip() or f"survey_{row.index + 1}"
            try:
                copy_count = int(row.count_edit.text())
            except ValueError:
                copy_count = 1

            for copy_i in range(copy_count):
                clean_type = survey_name if copy_count == 1 else f"{survey_name}{copy_i + 1}"
                prefix = f"{clean_type}_"

                check_matrix_bases: set[str] = set()
                for field_name in template_fields.keys():
                    m = CANONICAL_PATTERN.match(field_name)
                    if m and m.group(1).capitalize() == 'Check' and m.group(3) is not None:
                        check_matrix_bases.add(f"Check{m.group(2)}")
                    else:
                        clean_f = field_name[:-2] if field_name.endswith('_n') else field_name
                        expected.add(f"{prefix}{clean_f}")

                for base_name in check_matrix_bases:
                    expected.add(f"{prefix}{base_name}")

        return expected

    def _ensure_or_create_field_mapping(self) -> Dict[str, str]:
        expected = self._get_expected_system_fields()
        if not expected:
            return self.field_mapping

        # Start with any stored mapping from config
        stored = self._load_mapping_from_config(self.active_config_name) if self.active_config_name else {}
        
        mapping = {}
        for f in sorted(expected, key=natural_sort_key):
            # Prioritize: 1) stored in config, 2) in-memory field_mapping, 3) identity default
            mapping[f] = stored.get(f, self.field_mapping.get(f, f))

        self.field_mapping = mapping
        return mapping

    def _load_field_mapping_for_extraction(self) -> Dict[str, str]:
        if self.active_config_name:
            stored = self._load_mapping_from_config(self.active_config_name)
            if stored:
                self.field_mapping.update(stored)
                return self.field_mapping
        return self.field_mapping

    def _load_mapping_from_config(self, config_name: str) -> Dict[str, str]:
        if not config_name:
            return {}
        cfg = ConfigParser()
        cfg.read('config.ini')
        if not cfg.has_section(config_name):
            return {}
        raw = cfg.get(config_name, 'field_mapping_json', fallback='')
        if raw.strip():
            try:
                return dict(json.loads(raw))
            except Exception:
                return {}
        return {}

    def _save_mapping_to_config(self, config_name: str, mapping: Dict[str, str]) -> None:
        if not config_name:
            return
        cfg = ConfigParser()
        cfg.read('config.ini')
        if not cfg.has_section(config_name):
            cfg.add_section(config_name)
        cfg.set(config_name, 'field_mapping_json', json.dumps(mapping, ensure_ascii=False))
        with open('config.ini', 'w', encoding='utf-8') as f:
            cfg.write(f)

    def editFieldMappingFromExtraction(self) -> None:
        """Edit and save the column name remapping for the active configuration."""
        expected = self._get_expected_system_fields()
        root_order = self.get_configured_root_order()
        
        sorted_expected = sorted(
            expected,
            key=lambda f: (
                root_order.index(self._get_root_survey_name(f)) if self._get_root_survey_name(f) in root_order else 9999,
                natural_sort_key(f)
            )
        )
        
        stored = self._load_mapping_from_config(self.active_config_name) if self.active_config_name else {}
        mapping = {f: stored.get(f, self.field_mapping.get(f, f)) for f in sorted_expected}

        dialog = QDialog(self)
        dialog.setWindowTitle("Save/Edit Field Mapping")
        dialog.resize(750, 520)
        lay = QVBoxLayout()
        label = QLabel("Map system field identifiers to user-friendly column headers for Excel/CSV:")
        label.setStyleSheet("font-weight: bold;")
        lay.addWidget(label)

        table = QTableWidget(len(mapping), 2)
        table.setHorizontalHeaderLabels(["System Field", "Masterfile Column Name"])
        for idx, (k, v) in enumerate(mapping.items()):
            item_k = QTableWidgetItem(k)
            item_k.setFlags(Qt.ItemFlag.ItemIsEnabled)
            table.setItem(idx, 0, item_k)
            table.setItem(idx, 1, QTableWidgetItem(v))
        table.horizontalHeader().setStretchLastSection(True)
        table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeMode.ResizeToContents)
        lay.addWidget(table)
        
        btn = QPushButton("Save Field Mapping")
        def save():
            updated = {}
            for r in range(table.rowCount()):
                k = table.item(r, 0).text()
                val = table.item(r, 1).text().strip()
                updated[k] = val or k
            self.field_mapping.update(updated)
            if self.active_config_name:
                self._save_mapping_to_config(self.active_config_name, self.field_mapping)
                QMessageBox.information(self, "Success", f"Field mapping saved to active configuration '{self.active_config_name}'.")
            else:
                QMessageBox.information(self, "Success", "Field mapping saved in memory. Click 'Save Configuration' to create a permanent snapshot.")
            dialog.accept()
        btn.clicked.connect(save)
        lay.addWidget(btn)
        dialog.setLayout(lay)
        dialog.exec()

    def updateRecentConfigsMenu(self) -> None:
        if self.recent_configs_menu:
            self.recent_configs_menu.clear()
        if self.delete_configs_menu:
            self.delete_configs_menu.clear()
        cfg = ConfigParser()
        cfg.read('config.ini')
        for s in cfg.sections():
            if s == 'PrefillDefaults':
                continue
            if self.recent_configs_menu:
                act = QAction(s, self)
                act.triggered.connect(lambda chk, n=s: self.ConfigLoad(n))
                self.recent_configs_menu.addAction(act)
            if self.delete_configs_menu:
                act_d = QAction(s, self)
                act_d.triggered.connect(lambda chk, n=s: self.confirmDeleteConfig(n))
                self.delete_configs_menu.addAction(act_d)

    def showSaveConfigWindow(self) -> None:
        self.save_win = SaveConfigWindow(self)
        self.save_win.signal.connect(self.ConfigGen)
        self.save_win.exec()

    def confirmDeleteConfig(self, name: str) -> None:
        if QMessageBox.question(self, "Confirm Delete", f"Delete configuration '{name}'?",
                                cast(QMessageBox.StandardButton, QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)) == QMessageBox.StandardButton.Yes:
            cfg = ConfigParser()
            cfg.read('config.ini')
            cfg.remove_section(name)
            with open('config.ini', 'w') as f:
                cfg.write(f)

    def ConfigGen(self, name: str) -> None:
        """Create a complete snapshot including folders, questionnaires, and name remapping."""
        if not name:
            return self.error_window("No name provided!")
        cfg = ConfigParser()
        cfg.read('config.ini')
        if cfg.has_section(name):
            cfg.remove_section(name)
        cfg.add_section(name)
        
        cfg.set(name, 'target_path', self.target_path)
        cfg.set(name, 'quest_count', str(len(self.questionnaire_rows)))
        for i, row in enumerate(self.questionnaire_rows):
            cfg.set(name, f'quest_{i}_name', row.name_edit.text())
            cfg.set(name, f'quest_{i}_path', row.template_path)
            cfg.set(name, f'quest_{i}_count', row.count_edit.text())
        cfg.set(name, 'source_path', self.source_path)
        cfg.set(name, 'excel_path', self.excel_path)
        
        # Ensure mapping contains all expected fields and current edits
        self.field_mapping = self._ensure_or_create_field_mapping()
        cfg.set(name, 'field_mapping_json', json.dumps(self.field_mapping, ensure_ascii=False))
        
        with open('config.ini', 'w', encoding='utf-8') as f:
            cfg.write(f)
        
        self.active_config_name = name
        QMessageBox.information(self, "Success", f"Configuration snapshot '{name}' (with name remapping) saved successfully!")

    def ConfigLoad(self, name: str) -> None:
        """Load a complete snapshot, restoring all forms and column name remappings."""
        cfg = ConfigParser()
        cfg.read('config.ini')
        if not cfg.has_section(name):
            return self.error_window("Configuration not found!")
        
        self.target_path = cfg.get(name, 'target_path', fallback='')
        self.target_pathset.setText(self.target_path)
        
        qc = cfg.getint(name, 'quest_count', fallback=0)
        if qc > 0:
            self.quest_count_edit.setText(str(qc))
            self.create_questionnaire_rows()
            for i, row in enumerate(self.questionnaire_rows):
                row.name_edit.setText(cfg.get(name, f'quest_{i}_name', fallback=''))
                row.template_path = cfg.get(name, f'quest_{i}_path', fallback='')
                row.path_display.setText(row.template_path)
                row.count_edit.setText(cfg.get(name, f'quest_{i}_count', fallback='1'))
        
        self.source_path = cfg.get(name, 'source_path', fallback='')
        self.source_pathset.setText(self.source_path)
        self.excel_path = cfg.get(name, 'excel_path', fallback='')
        self.excel_pathset.setText(self.excel_path)
        
        self.active_config_name = name
        self.field_mapping = self._load_mapping_from_config(name)
        
        QMessageBox.information(self, "Success", f"Configuration snapshot '{name}' (with column name remapping) loaded successfully!")

    def enterWhatsThisMode(self) -> None:
        QWhatsThis.enterWhatsThisMode()

    def onMyToolBarButtonClick3(self) -> None:
        manual_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'USER_MANUAL.pdf')
        if os.path.exists(manual_path):
            os.startfile(manual_path)
        else:
            self.error_window("Manual not found:\n" + manual_path)

    def toggleBatchModeGeneration(self, checked: bool) -> None:
        self.nameset.setVisible(not checked)
        self.batch_text_gen.setVisible(checked)

    def toggleBatchModeExtraction(self, checked: bool) -> None:
        self.nameset.setVisible(not checked)
        self.batch_text_extract.setVisible(checked)

    def importParticipantList(self, target_text_edit: QTextEdit) -> None:
        path = QFileDialog.getOpenFileName(self, "Import Participant List", "", "Text/CSV (*.txt *.csv);;All (*.*)")[0]
        if not path:
            return
        pids = []
        with open(path, 'r', encoding='utf-8') as f:
            for line in f:
                for p in line.replace(',', '\n').split('\n'):
                    if p.strip():
                        pids.append(p.strip())
        unique = list(dict.fromkeys(pids))
        target_text_edit.setPlainText('\n'.join(unique))

    def saveTemplateBundle(self) -> None:
        name, ok = QInputDialog.getText(self, "Save Bundle", "Bundle name:")
        if not ok or not name.strip():
            return
        d = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'template_bundles')
        os.makedirs(d, exist_ok=True)
        data = [{'name': r.name_edit.text(), 'path': r.template_path, 'count': r.count_edit.text()} for r in self.questionnaire_rows]
        with open(os.path.join(d, f"{name.strip()}.json"), 'w', encoding='utf-8') as f:
            json.dump({'name': name.strip(), 'items': data}, f, indent=2)

    def loadTemplateBundle(self) -> None:
        d = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'template_bundles')
        if not os.path.exists(d):
            return self.error_window("No bundles found!")
        files = [f.replace('.json', '') for f in os.listdir(d) if f.endswith('.json')]
        if not files:
            return self.error_window("No bundles found!")
        choice, ok = QInputDialog.getItem(self, "Load Bundle", "Select bundle:", files, 0, False)
        if not ok:
            return
        with open(os.path.join(d, f"{choice}.json"), 'r', encoding='utf-8') as f:
            bundle = json.load(f)
        items = bundle.get('items', [])
        self.quest_count_edit.setText(str(len(items)))
        self.create_questionnaire_rows()
        for idx, it in enumerate(items):
            if idx < len(self.questionnaire_rows):
                self.questionnaire_rows[idx].name_edit.setText(it.get('name', ''))
                self.questionnaire_rows[idx].template_path = it.get('path', '')
                self.questionnaire_rows[idx].path_display.setText(it.get('path', ''))
                self.questionnaire_rows[idx].count_edit.setText(it.get('count', '1'))

    def parseParticipantIDs(self, text: str) -> List[str]:
        return [line.strip() for line in text.replace(',', '\n').split('\n') if line.strip()]

    def checkDuplicate(self, participant_id: str) -> bool:
        if not os.path.exists(self.excel_path):
            return False
        ext = self.excel_path.lower().split('.')[-1]
        try:
            if ext == 'csv':
                with open(self.excel_path, 'r', encoding='utf-8') as f:
                    return any(row.get('participant_id') == participant_id for row in csv.DictReader(f))
            elif ext in ['xls', 'xlsx']:
                app = xlw.App(visible=False, add_book=False)
                try:
                    wb = app.books.open(self.excel_path)
                    try:
                        for sheet in wb.sheets:
                            col_a = sheet.range('A2:A1000').value
                            ids = [str(x).strip() for x in (col_a if isinstance(col_a, list) else [col_a]) if x is not None]
                            if participant_id in ids:
                                return True
                    finally:
                        wb.close()
                finally:
                    app.quit()
        except Exception:
            return False
        return False

    def generateMissingDataReport(self) -> None:
        if not self.source_path or not os.path.exists(self.source_path):
            return self.error_window("Select a valid source folder!")
        folders = [f for f in os.listdir(self.source_path) if os.path.isdir(os.path.join(self.source_path, f))]
        report = []
        root_order = self.get_configured_root_order()
        for folder in sorted(folders, key=natural_sort_key):
            try:
                sheet_data = self._prepare_data_for_extraction(folder)
                total_fields = sum(len(d) - 1 for d in sheet_data.values())
                ordered_sheets = sorted(sheet_data.keys(), key=lambda s: (root_order.index(s) if s in root_order else 9999, natural_sort_key(s)))
                report.append(f"✓ {folder}: Complete ({len(sheet_data)} sheets: {', '.join(ordered_sheets)}, {total_fields} fields)")
            except Exception as e:
                report.append(f"✗ {folder}: INCOMPLETE - {str(e)}")
        
        dialog = QDialog(self)
        dialog.setWindowTitle("Missing Data Report")
        dialog.resize(700, 500)
        lay = QVBoxLayout()
        t = QTextEdit()
        t.setReadOnly(True)
        t.setText('\n'.join(report))
        lay.addWidget(t)
        dialog.setLayout(lay)
        dialog.exec()

    def error_window(self, message: str) -> None:
        QMessageBox.warning(self, "ERROR", message)


if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())