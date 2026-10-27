# -*- coding: utf-8 -*-
"""
Survey Workbench - Enhanced Participant Data Management System
Version: 2.6 (Dynamic Sizing & Adaptive Layout Edition)

Author: Cagatay Gutt
Last Updated: 2026

Features:
- Dynamic window resizing that automatically adapts and shrinks/grows based on presented content and expansion state.
- Compact participant declaration field with optimized sizing.
- True collapsible curator configuration container (fully hidden/collapsed by default).
- Parallelized Batch Mode with single consolidated preview and batch Excel writing.
- Config snapshotting: Saves template paths, copy counts, paths, AND complete column remappings into config.ini[cite: 1].
- Fully qualified mapping keys ({survey}_{field}) preventing cross-sheet name collisions[cite: 1].
- Dedicated sheets per root PDF in Excel with Column A as 'participant_id'[cite: 1].
- Natural numerical sorting (Text1, Text2... Text10) across all sheets and dialogs[cite: 1].
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
                             QTableWidgetItem, QHeaderView, QCheckBox, QInputDialog, QHBoxLayout,
                             QToolButton)
from configparser import ConfigParser
from PyQt5.QtCore import pyqtSignal, Qt, QTimer
import xlwings as xlw  # type: ignore[import]
from pypdf import PdfReader, PdfWriter
from pypdf.generic import NameObject, create_string_object, ArrayObject

# Canonical grammar for all survey presentations (Section 7.4)[cite: 1]
CANONICAL_PATTERN = re.compile(r'^(Text|Check)(\d+)(?:_(\d+))?(_n)?$', re.IGNORECASE)


def natural_sort_key(s: Any) -> List[Any]:
    """Naturally sort alphanumeric strings so Text2 comes before Text10[cite: 1]."""
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
        self.setWindowTitle("Survey Workbench v2.6")
        
        self.questionnaire_rows: List[QuestionnaireRow] = []
        self.recent_configs_menu: Optional[QMenu] = None
        self.delete_configs_menu: Optional[QMenu] = None
        
        main_widget = QWidget()
        main_layout = QVBoxLayout()
        main_layout.setSizeConstraint(QVBoxLayout.SetFixedSize)  # Enables dynamic resizing based on contents
        
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
        
        # ========== PARTICIPANT DECLARATION SECTION (UNIFIED & COMPACT) ==========
        part_frame: QFrame = QFrame()
        part_frame.setFrameStyle(QFrame.Shape.Box | QFrame.Shadow.Raised)
        part_layout = QGridLayout()
        
        part_title = QLabel("Participant Declaration")
        part_title.setStyleSheet("font-weight: bold; font-size: 13px;")
        
        part_label = QLabel("Participant ID(s):")
        part_label.setStyleSheet("font-weight: bold")
        self.nameset = QLineEdit()
        self.nameset.setToolTip("Enter single participant ID (e.g., P001)")
        
        self.batch_mode_global = QCheckBox("Batch Mode (multiple IDs)")
        self.batch_mode_global.toggled.connect(self.toggleBatchModeGlobal)
        
        import_list_btn = QPushButton("Import List...")
        import_list_btn.clicked.connect(lambda: self.importParticipantList(self.batch_text_global))
        
        self.batch_text_global = QTextEdit()
        self.batch_text_global.setPlaceholderText("Enter participant IDs (one per line or comma-separated):\nP001\nP002\nP003")
        self.batch_text_global.setVisible(False)
        self.batch_text_global.setMaximumHeight(80)
        
        part_layout.addWidget(part_title, 0, 0, 1, 3)
        part_layout.addWidget(part_label, 1, 0)
        part_layout.addWidget(self.nameset, 1, 1)
        part_layout.addWidget(self.batch_mode_global, 1, 2)
        part_layout.addWidget(import_list_btn, 2, 2)
        part_layout.addWidget(self.batch_text_global, 2, 0, 1, 2)
        
        part_frame.setLayout(part_layout)
        main_layout.addWidget(part_frame)

        # ========== COLLAPSIBLE CONFIGURATION CONTAINER (HIDDEN BY DEFAULT) ==========
        self.toggle_config_btn = QToolButton()
        self.toggle_config_btn.setText("▶ Configuration Settings")
        self.toggle_config_btn.setStyleSheet("font-weight: bold; border: none; text-align: left; padding: 5px;")
        self.toggle_config_btn.setCheckable(True)
        self.toggle_config_btn.setChecked(False)
        self.toggle_config_btn.setToolButtonStyle(Qt.ToolButtonTextBesideIcon)
        self.toggle_config_btn.clicked.connect(self.toggleConfigurationContainer)
        main_layout.addWidget(self.toggle_config_btn)

        self.config_container = QWidget()
        config_container_layout = QVBoxLayout()
        config_container_layout.setContentsMargins(0, 0, 0, 0)

        # Generation Sub-section
        gen_frame: QFrame = QFrame()
        gen_frame.setFrameStyle(QFrame.Shape.Box | QFrame.Shadow.Raised)
        gen_layout = QGridLayout()
        
        gen_title = QLabel("Generate Participant Folder(s)")
        gen_title.setStyleSheet("font-weight: bold; font-size: 13px;")
        
        target_label = QLabel("Target folder PATH:")
        self.target_pathset = QLabel()
        target_pathbut = QPushButton("Search")
        target_pathbut.clicked.connect(self.select_target_folder)
        
        quest_count_label = QLabel("Number of questionnaire types:")
        self.quest_count_edit = QLineEdit()
        quest_count_button = QPushButton("Confirm")
        quest_count_button.clicked.connect(self.create_questionnaire_rows)
        
        gen_layout.addWidget(gen_title, 0, 0, 1, 3)
        gen_layout.addWidget(target_label, 1, 0)
        gen_layout.addWidget(self.target_pathset, 1, 1)
        gen_layout.addWidget(target_pathbut, 1, 2)
        gen_layout.addWidget(quest_count_label, 2, 0)
        gen_layout.addWidget(self.quest_count_edit, 2, 1)
        gen_layout.addWidget(quest_count_button, 2, 2)
        gen_frame.setLayout(gen_layout)
        config_container_layout.addWidget(gen_frame)
        
        self.scroll_area = QScrollArea()
        self.scroll_area.setWidgetResizable(True)
        self.scroll_widget = QWidget()
        self.scroll_layout = QGridLayout()
        self.scroll_widget.setLayout(self.scroll_layout)
        self.scroll_area.setWidget(self.scroll_widget)
        self.scroll_area.setVisible(False)
        config_container_layout.addWidget(self.scroll_area)

        # Extraction Sub-section
        extract_frame = QFrame()
        extract_frame.setFrameStyle(QFrame.Shape.Box | QFrame.Shadow.Raised)
        extract_layout = QGridLayout()
        
        extract_title = QLabel("Extract Data to Excel")
        extract_title.setStyleSheet("font-weight: bold; font-size: 13px;")
        
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
        
        extract_frame.setLayout(extract_layout)
        config_container_layout.addWidget(extract_frame)

        self.config_container.setLayout(config_container_layout)
        self.config_container.setVisible(False)  # Collapsed by default for end-users
        main_layout.addWidget(self.config_container)
        
        # ========== ACTION BUTTONS (ALWAYS VISIBLE FOR END-USERS) ==========
        action_layout = QHBoxLayout()
        self.generate_button = QPushButton("Generate Participant Folder(s)")
        self.generate_button.setStyleSheet("font-weight: bold; padding: 10px;")
        self.generate_button.clicked.connect(self.generate_participant_folder)
        
        self.extract_button = QPushButton("Extract Data to Masterfile")
        self.extract_button.setStyleSheet("font-weight: bold; padding: 10px;")
        self.extract_button.clicked.connect(self.extract_data)

        action_layout.addWidget(self.generate_button)
        action_layout.addWidget(self.extract_button)
        main_layout.addLayout(action_layout)
        
        main_widget.setLayout(main_layout)
        self.setCentralWidget(main_widget)
        
        self.target_path: str = ""
        self.source_path: str = ""
        self.excel_path: str = ""
        self.active_config_name: str = ""
        self.field_mapping: Dict[str, str] = {}
        
        # Initial compact sizing and adjustment
        self.adjustSize()

    def toggleConfigurationContainer(self, checked: bool) -> None:
        """Toggles the visibility of the configuration container and dynamically adjusts window size."""
        self.config_container.setVisible(checked)
        if checked:
            self.toggle_config_btn.setText("▼ Configuration Settings")
        else:
            self.toggle_config_btn.setText("▶ Configuration Settings")
        
        # Dynamically resize window to snugly fit current contents
        QTimer.singleShot(10, self.adjustSize)

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
            QTimer.singleShot(10, self.adjustSize)
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

    def get_active_participant_ids(self) -> List[str]:
        if self.batch_mode_global.isChecked():
            return self.parseParticipantIDs(self.batch_text_global.toPlainText())
        else:
            pid = self.nameset.text().strip()
            return [pid] if pid else []
    
    def generate_participant_folder(self) -> None:
        try:
            participant_ids = self.get_active_participant_ids()
            if not participant_ids:
                return self.error_window("Please enter at least one participant ID!")
            if not self.target_path:
                return self.error_window("Please select a target folder! (Expand 'Configuration Settings' to configure paths)")
            if not self.questionnaire_rows:
                return self.error_window("Please configure questionnaires! (Expand 'Configuration Settings' to configure templates)")

            all_generated_pdf_paths: Dict[str, List[str]] = {}
            for pid in participant_ids:
                participant_folder = os.path.join(self.target_path, pid)
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
                        new_filename = f"{pid}_{survey_name}{ext}" if copy_count == 1 else f"{pid}_{survey_name}{i + 1}{ext}"
                        dest_path = os.path.join(participant_folder, new_filename)
                        shutil.copy2(row.template_path, dest_path)
                        if ext.lower() == '.pdf':
                            generated_pdf_paths.append(dest_path)
                
                if generated_pdf_paths:
                    all_generated_pdf_paths[pid] = generated_pdf_paths

            if all_generated_pdf_paths:
                self.field_mapping = self._ensure_or_create_field_mapping()
                apply_prefill, batch_prefill_map = self.showBatchGenerationPrefillDialog(all_generated_pdf_paths)
                if not apply_prefill:
                    return
                for pdf_path, field_values in batch_prefill_map.items():
                    if field_values:
                        self._write_pdf_form_fields(pdf_path, field_values)

            QMessageBox.information(
                self, "Success",
                f"Successfully generated folders and questionnaires for {len(participant_ids)} participant(s)!"
            )
        except Exception as e:
            self.error_window(f"Error generating folder(s): {str(e)}")

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

    def showBatchGenerationPrefillDialog(self, all_generated_pdf_paths: Dict[str, List[str]]) -> tuple[bool, Dict[str, Dict[str, Any]]]:
        dialog = QDialog(self)
        dialog.setWindowTitle("Consolidated Batch Preview & Generation")
        dialog.setModal(True)
        dialog.resize(900, 650)

        layout = QVBoxLayout()
        label = QLabel(f"Preview & Edit Data for {len(all_generated_pdf_paths)} Participant(s)")
        label.setStyleSheet("font-weight: bold; font-size: 14px;")
        layout.addWidget(label)

        table = QTableWidget()
        table.setColumnCount(4)
        table.setHorizontalHeaderLabels(["Participant ID", "File", "Field", "Data"])

        rows: List[tuple[str, str, str, str, str]] = []
        file_to_path: Dict[str, str] = {}
        file_to_extracted_fields: Dict[str, Dict[str, str]] = {}
        
        root_order = self.get_configured_root_order()

        for participant_id, pdf_paths in all_generated_pdf_paths.items():
            sorted_pdf_paths = sorted(
                pdf_paths,
                key=lambda p: (
                    root_order.index(self._get_root_survey_name(os.path.splitext(os.path.basename(p))[0].replace(f'{participant_id}_', '', 1)))
                    if self._get_root_survey_name(os.path.splitext(os.path.basename(p))[0].replace(f'{participant_id}_', '', 1)) in root_order
                    else 9999,
                    natural_sort_key(os.path.basename(p))
                )
            )

            for pdf_path in sorted_pdf_paths:
                file_label = os.path.basename(pdf_path)
                file_to_path[f"{participant_id}::{file_label}"] = pdf_path
                survey_type = os.path.splitext(file_label)[0].replace(f'{participant_id}_', '', 1)
                extracted_fields = self._extract_pdf_form_fields(pdf_path)
                if not extracted_fields:
                    continue
                defaults = self._build_prefill_defaults(participant_id, extracted_fields)
                file_to_extracted_fields[f"{participant_id}::{file_label}"] = extracted_fields

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
                    rows.append((participant_id, file_label, f_name, display_name, f_val))

                for base_name, options in sorted(check_matrix_groups.items(), key=lambda it: natural_sort_key(it[0])):
                    system_key = f"{survey_type}_{base_name}"
                    display_name = self.field_mapping.get(system_key, self.field_mapping.get(base_name, base_name))
                    checked_opt = next((opt for opt, cval in sorted(options.items(), key=lambda it: natural_sort_key(it[0])) if cval == '1'), '')
                    rows.append((participant_id, file_label, base_name, display_name, checked_opt))

        table.setRowCount(len(rows))
        for idx, (pid, file_label, f_name, display_name, value) in enumerate(rows):
            pid_item = QTableWidgetItem(pid)
            pid_item.setFlags(pid_item.flags() & ~Qt.ItemFlag.ItemIsEditable)
            file_item = QTableWidgetItem(file_label)
            file_item.setFlags(file_item.flags() & ~Qt.ItemFlag.ItemIsEditable)
            name_item = QTableWidgetItem(display_name)
            name_item.setFlags(name_item.flags() | Qt.ItemFlag.ItemIsEditable)
            
            table.setItem(idx, 0, pid_item)
            table.setItem(idx, 1, file_item)
            table.setItem(idx, 2, name_item)
            table.setItem(idx, 3, QTableWidgetItem(value))

        table.horizontalHeader().setStretchLastSection(True)
        table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeMode.ResizeToContents)
        table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeMode.ResizeToContents)
        table.horizontalHeader().setSectionResizeMode(2, QHeaderView.ResizeMode.ResizeToContents)
        layout.addWidget(table)

        button_layout = QHBoxLayout()
        confirm_btn = QPushButton("Confirm and Generate All")
        cancel_btn = QPushButton("Cancel")
        button_layout.addWidget(confirm_btn)
        button_layout.addWidget(cancel_btn)
        layout.addLayout(button_layout)
        dialog.setLayout(layout)

        result: Dict[str, Any] = {"action": "cancel", "values": {}}

        def confirm_and_close() -> None:
            prefill_values: Dict[str, Dict[str, Any]] = {}
            for paths in all_generated_pdf_paths.values():
                for p in paths:
                    prefill_values[p] = {}
            
            for row_index in range(table.rowCount()):
                pid_item = table.item(row_index, 0)
                file_item = table.item(row_index, 1)
                name_item = table.item(row_index, 2)
                value_item = table.item(row_index, 3)
                if pid_item is None or file_item is None or name_item is None:
                    continue
                pid = pid_item.text()
                file_name = file_item.text()
                pdf_path = file_to_path.get(f"{pid}::{file_name}")
                underlying_field_name = rows[row_index][2]
                val = value_item.text().strip() if value_item else ""
                
                survey_type = os.path.splitext(file_name)[0].replace(f'{pid}_', '', 1)
                new_name = name_item.text().strip()
                if new_name:
                    self.field_mapping[f"{survey_type}_{underlying_field_name}"] = new_name

                if not pdf_path:
                    continue

                if underlying_field_name.startswith('Check') and '_' not in underlying_field_name:
                    base_name = underlying_field_name
                    raw_fields = file_to_extracted_fields.get(f"{pid}::{file_name}", {})
                    for raw_f in raw_fields:
                        m = CANONICAL_PATTERN.match(raw_f)
                        if m and f"Check{m.group(2)}" == base_name and m.group(3) is not None:
                            col_opt = m.group(3)
                            prefill_values[pdf_path][raw_f] = '1' if col_opt == val else '0'
                else:
                    prefill_values[pdf_path][underlying_field_name] = val

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
        """Extract data segregated by root sheet name, preserving root sequence and numerical order[cite: 1]."""
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

    def _readout_batch(self, participant_ids: List[str]) -> None:
        """Route all participants' extracted data into masterfile sheets in a single session without repeated file openings."""
        if not self.excel_path:
            raise ValueError("Please select a masterfile! (Expand 'Configuration Settings' to select masterfile)")

        file_ext = self.excel_path.lower().split('.')[-1]
        root_order = self.get_configured_root_order()

        all_participants_sheet_data: Dict[str, Dict[str, Dict[str, Any]]] = {}
        for pid in participant_ids:
            all_participants_sheet_data[pid] = self._prepare_data_for_extraction(pid)

        if file_ext in ['xls', 'xlsx']:
            wb: xlw.Book = xlw.Book(self.excel_path)
            try:
                for pid in participant_ids:
                    sheet_data = all_participants_sheet_data.get(pid, {})
                    ordered_sheet_names = sorted(
                        sheet_data.keys(),
                        key=lambda s: (root_order.index(s) if s in root_order else 9999, natural_sort_key(s))
                    )

                    for sheet_name in ordered_sheet_names:
                        raw_data = sheet_data[sheet_name]
                        export_data: Dict[str, Any] = {'participant_id': pid}
                        
                        sorted_field_keys = sorted(
                            [k for k in raw_data.keys() if k != 'participant_id'],
                            key=natural_sort_key
                        )

                        for key in sorted_field_keys:
                            value = raw_data[key]
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
                            
                            first_row = sheet.range('1:1').value
                            header_map: Dict[str, int] = {}
                            if first_row:
                                row_vals = first_row if isinstance(first_row, list) else [first_row]
                                for c_idx, val in enumerate(row_vals, start=1):
                                    if val is not None and str(val).strip():
                                        header_map[str(val).strip()] = c_idx

                            existing_cols = set(header_map.keys())
                            incoming_cols = set(export_data.keys())
                            
                            if existing_cols != incoming_cols:
                                missing = existing_cols - incoming_cols
                                extra = incoming_cols - existing_cols
                                msg = (f"Schema Mismatch Detected in Sheet '{sheet_name}' for participant '{pid}'!\n")
                                if missing:
                                    msg += f"\nMissing: {list(missing)}"
                                if extra:
                                    msg += f"\nExtra: {list(extra)}"
                                msg += "\n\nProceed anyway?"
                                
                                if QMessageBox.question(self, "Schema Mismatch Warning", msg,
                                                        cast(QMessageBox.StandardButton, QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No),
                                                        QMessageBox.StandardButton.No) == QMessageBox.StandardButton.No:
                                    raise ValueError(f"Extraction aborted for '{pid}' due to schema mismatch in '{sheet_name}'.")

                            next_col = max(header_map.values(), default=0) + 1
                            for key in export_data.keys():
                                if key not in header_map:
                                    sheet.range(1, next_col).value = key
                                    header_map[key] = next_col
                                    next_col += 1

                            target_row = None
                            r_idx = 2
                            while True:
                                val = sheet.range(f'A{r_idx}').value
                                if val is None:
                                    break
                                if str(val).strip() == str(pid):
                                    target_row = r_idx
                                    break
                                r_idx += 1

                            if target_row is not None:
                                next_row = target_row
                            else:
                                next_row = 2
                                while sheet.range(f'A{next_row}').value is not None:
                                    next_row += 1

                        else:
                            if len(wb.sheets) == 1 and wb.sheets[0].range('A1').value is None:
                                sheet = cast(xlw.Sheet, wb.sheets[0])
                                sheet.name = sheet_name
                            else:
                                sheet = cast(xlw.Sheet, wb.sheets.add(name=sheet_name, after=wb.sheets[-1]))

                            headers = list(export_data.keys())
                            header_map = {}
                            for idx, header in enumerate(headers, start=1):
                                sheet.range(1, idx).value = header
                                header_map[header] = idx

                            next_row = 2

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
            
            for pid in participant_ids:
                sheet_data = all_participants_sheet_data.get(pid, {})
                ordered_sheet_names = sorted(
                    sheet_data.keys(),
                    key=lambda s: (root_order.index(s) if s in root_order else 9999, natural_sort_key(s))
                )

                for sheet_name in ordered_sheet_names:
                    raw_data = sheet_data[sheet_name]
                    target_csv = os.path.join(base_dir, f"{base_name}_{sheet_name}.csv")
                    file_exists = os.path.exists(target_csv) and os.path.getsize(target_csv) > 0

                    export_data = {'participant_id': pid}
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

                    existing_index = -1
                    for idx, row in enumerate(existing_rows):
                        if row.get('participant_id') == pid:
                            existing_index = idx
                            break

                    if existing_index != -1:
                        existing_rows[existing_index] = export_data
                    else:
                        existing_rows.append(export_data)

                    new_headers = [k for k in export_data.keys() if k not in existing_headers]
                    all_headers = existing_headers + new_headers if existing_headers else list(export_data.keys())

                    with open(target_csv, 'w', newline='', encoding='utf-8') as f:
                        writer = csv.DictWriter(f, fieldnames=all_headers)
                        writer.writeheader()
                        writer.writerows(existing_rows)
        else:
            raise ValueError(f"Unsupported masterfile format: {file_ext}")

    def extract_data(self) -> None:
        if not self.excel_path:
            return self.error_window("Please select a masterfile! (Expand 'Configuration Settings' to select masterfile)")
        if not self.source_path:
            return self.error_window("Please select a source folder! (Expand 'Configuration Settings' to select source folder)")

        self.field_mapping = self._load_field_mapping_for_extraction()
        participant_ids = self.get_active_participant_ids()
        if not participant_ids:
            return self.error_window("Please enter at least one participant ID!")

        try:
            if len(participant_ids) > 1 or self.batch_mode_global.isChecked():
                self._readout_batch(participant_ids)
                QMessageBox.information(self, "Batch Extraction Complete", f"Successfully extracted {len(participant_ids)} participant(s) into masterfile!")
            else:
                participant_id = participant_ids[0]
                sheet_data = self._prepare_data_for_extraction(participant_id)
                confirmed, edited_sheet_data = self.showPreviewDialog(participant_id, sheet_data)
                if not confirmed:
                    return
                self._readout_batch([participant_id])
                QMessageBox.information(self, "Success", f"Data extracted successfully into sheets for {participant_id}!")
        except Exception as e:
            self.error_window(f"Error extracting data: {str(e)}")

    def showPreviewDialog(self, participant_id: str, sheet_data: Dict[str, Dict[str, Any]]) -> Tuple[bool, Dict[str, Dict[str, Any]]]:
        """4-column preview ordered first by root-PDF sequence, then numerically[cite: 1]."""
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
            si.setFlags(si.flags() & ~Qt.ItemFlag.ItemIsEditable)
            fi = QTableWidgetItem(f_key)
            fi.setFlags(fi.flags() & ~Qt.ItemFlag.ItemIsEditable)
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
        
        btn_layout = QHBoxLayout()
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

            if self.active_config_name:
                self._save_mapping_to_config(self.active_config_name, self.field_mapping)

            result[0] = True
            dialog.accept()
        
        c_btn.clicked.connect(save_and_close)
        can_btn.clicked.connect(dialog.reject)
        dialog.exec()
        return result[0], edited_output if result[0] else sheet_data

    def _get_expected_system_fields(self) -> set[str]:
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

        stored = self._load_mapping_from_config(self.active_config_name) if self.active_config_name else {}
        
        mapping = {}
        for f in sorted(expected, key=natural_sort_key):
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
            item_k.setFlags(item_k.flags() & ~Qt.ItemFlag.ItemIsEditable)
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
        """Create a complete snapshot including folders, questionnaires, and name remapping[cite: 1]."""
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

    def toggleBatchModeGlobal(self, checked: bool) -> None:
        self.nameset.setVisible(not checked)
        self.batch_text_global.setVisible(checked)
        QTimer.singleShot(10, self.adjustSize)

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

    def parseParticipantIDs(self, text: str) -> List[str]:
        return [line.strip() for line in text.replace(',', '\n').split('\n') if line.strip()]

    def error_window(self, message: str) -> None:
        QMessageBox.warning(self, "ERROR", message)


if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())