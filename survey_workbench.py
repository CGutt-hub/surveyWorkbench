import os
import sys
from typing import Any, Dict, List, Optional

from PyQt5.QtCore import Qt, QTimer
from PyQt5.QtWidgets import (
    QAction, QApplication, QCheckBox, QFileDialog, QFrame, QGridLayout,
    QHBoxLayout, QLabel, QLayout, QLineEdit, QMainWindow, QMenu,
    QMessageBox, QPushButton, QScrollArea, QStatusBar, QTextEdit,
    QToolButton, QVBoxLayout, QWhatsThis, QWidget
)

from config_utils import (
    INVALID_ID_CHARS_PATTERN,
    default_condition_ordering,
    ensure_config_file_exists,
    get_config_path,
    load_condition_ordering_from_config,
    load_mapping_from_config,
    new_config_parser,
    save_mapping_to_config,
    write_config_atomically,
)
from dialogs import MarkdownManualDialog, PdfInspectorDialog, QuestionnaireRow, SaveConfigWindow
from mapping_dialogs import edit_condition_ordering_dialog, edit_field_mapping_dialog
from study_workflows import (
    get_expected_system_fields,
    load_config_profile,
    run_extract_data,
    run_generate_folders,
    save_config_profile,
)


class MainWindow(QMainWindow):
    def __init__(self, parent: Optional[QWidget] = None) -> None:
        super().__init__(parent)
        self.setWindowTitle("Survey Workbench")
        ensure_config_file_exists()

        self.questionnaire_rows: List[QuestionnaireRow] = []
        self.recent_configs_menu: Optional[QMenu] = None
        self.delete_configs_menu: Optional[QMenu] = None
        self.import_mapping_menu: Optional[QMenu] = None

        self.target_path: str = ""
        self.source_path: str = ""
        self.excel_path: str = ""
        self.active_config_name: str = ""
        self.field_mapping: Dict[str, str] = {}
        self.condition_ordering: Dict[str, Any] = default_condition_ordering()

        main_widget = QWidget()
        main_layout = QVBoxLayout()
        main_layout.setSizeConstraint(QLayout.SizeConstraint.SetFixedSize)

        button_action = QAction("&Save Configuration", self)
        button_action.triggered.connect(self.showSaveConfigWindow)

        button_action3 = QAction("User Manual", self)
        button_action3.triggered.connect(self.openUserManual)

        whats_this_action = QAction("What's This? (Click for Help)", self)
        whats_this_action.triggered.connect(self.enterWhatsThisMode)
        whats_this_action.setShortcut("Shift+F1")

        inspect_pdf_action = QAction("Inspect PDF Form Fields...", self)
        inspect_pdf_action.triggered.connect(self.inspectPdfFieldsDialog)

        self.setStatusBar(QStatusBar(self))
        menubar = self.menuBar()
        if menubar is not None:
            file_menu = menubar.addMenu("&File")
            if file_menu is not None:
                file_menu.addAction(button_action)
                self.recent_configs_menu = file_menu.addMenu("&Load Configuration")
                self.delete_configs_menu = file_menu.addMenu("&Delete Configuration")
                self.import_mapping_menu = file_menu.addMenu("&Import Field Mapping")
                file_menu.aboutToShow.connect(self.updateRecentConfigsMenu)

                edit_mapping_action = QAction("&Save/Edit Field Mapping", self)
                edit_mapping_action.triggered.connect(self.editFieldMappingFromExtraction)
                file_menu.addAction(edit_mapping_action)

                edit_condition_order_action = QAction("&Save/Edit Condition Ordering", self)
                edit_condition_order_action.triggered.connect(self.openConditionOrderingDialog)
                file_menu.addAction(edit_condition_order_action)

            help_menu = menubar.addMenu("&HELP")
            if help_menu is not None:
                help_menu.addAction(whats_this_action)
                help_menu.addAction(button_action3)
                help_menu.addAction(inspect_pdf_action)

        # Participant Declaration
        part_frame = QFrame()
        part_frame.setFrameShape(QFrame.Shape.Box)
        part_frame.setFrameShadow(QFrame.Shadow.Raised)
        part_layout = QGridLayout()

        part_title = QLabel("Participant Declaration")
        part_title.setStyleSheet("font-weight: bold; font-size: 13px;")

        part_label = QLabel("Participant ID(s):")
        part_label.setStyleSheet("font-weight: bold")
        self.nameset = QLineEdit()
        self.nameset.setToolTip("Enter a single participant ID (e.g., P001).")

        self.batch_mode_global = QCheckBox("Batch Mode (multiple IDs)")
        self.batch_mode_global.toggled.connect(self.toggleBatchModeGlobal)

        import_list_btn = QPushButton("Import List...")
        import_list_btn.clicked.connect(lambda _checked=False: self.importParticipantList(self.batch_text_global))

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

        # Configuration Container
        self.toggle_config_btn = QToolButton()
        self.toggle_config_btn.setText("▶ Configuration Settings")
        self.toggle_config_btn.setStyleSheet("font-weight: bold; border: none; text-align: left; padding: 5px;")
        self.toggle_config_btn.setCheckable(True)
        self.toggle_config_btn.setChecked(False)
        self.toggle_config_btn.setToolButtonStyle(Qt.ToolButtonStyle.ToolButtonTextBesideIcon)
        self.toggle_config_btn.clicked.connect(self.toggleConfigurationContainer)
        main_layout.addWidget(self.toggle_config_btn)

        self.config_container = QWidget()
        config_container_layout = QVBoxLayout()
        config_container_layout.setContentsMargins(0, 0, 0, 0)

        gen_frame = QFrame()
        gen_frame.setFrameShape(QFrame.Shape.Box)
        gen_frame.setFrameShadow(QFrame.Shadow.Raised)
        gen_layout = QGridLayout()

        gen_title = QLabel("Generate Participant Folder(s)")
        gen_title.setStyleSheet("font-weight: bold; font-size: 13px;")
        self.target_pathset = QLabel()
        target_pathbut = QPushButton("Search")
        target_pathbut.clicked.connect(self.select_target_folder)

        self.quest_count_edit = QLineEdit()
        quest_count_button = QPushButton("Confirm")
        quest_count_button.clicked.connect(self.create_questionnaire_rows)

        gen_layout.addWidget(gen_title, 0, 0, 1, 3)
        gen_layout.addWidget(QLabel("Target folder PATH:"), 1, 0)
        gen_layout.addWidget(self.target_pathset, 1, 1)
        gen_layout.addWidget(target_pathbut, 1, 2)
        gen_layout.addWidget(QLabel("Number of questionnaire types:"), 2, 0)
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

        extract_frame = QFrame()
        extract_frame.setFrameShape(QFrame.Shape.Box)
        extract_frame.setFrameShadow(QFrame.Shadow.Raised)
        extract_layout = QGridLayout()

        extract_title = QLabel("Extract Data to Masterfile")
        extract_title.setStyleSheet("font-weight: bold; font-size: 13px;")
        self.source_pathset = QLabel()
        source_pathbut = QPushButton("Search")
        source_pathbut.clicked.connect(self.select_source_folder)

        self.excel_pathset = QLabel()
        excel_pathbut = QPushButton("Search")
        excel_pathbut.clicked.connect(self.select_excel_file)

        extract_layout.addWidget(extract_title, 0, 0, 1, 3)
        extract_layout.addWidget(QLabel("Source folder PATH:"), 1, 0)
        extract_layout.addWidget(self.source_pathset, 1, 1)
        extract_layout.addWidget(source_pathbut, 1, 2)
        extract_layout.addWidget(QLabel("Masterfile PATH:"), 2, 0)
        extract_layout.addWidget(self.excel_pathset, 2, 1)
        extract_layout.addWidget(excel_pathbut, 2, 2)
        extract_frame.setLayout(extract_layout)
        config_container_layout.addWidget(extract_frame)

        self.config_container.setLayout(config_container_layout)
        self.config_container.setVisible(False)
        main_layout.addWidget(self.config_container)

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
        self.adjustSize()

    def toggleConfigurationContainer(self, checked: bool) -> None:
        self.config_container.setVisible(checked)
        self.toggle_config_btn.setText("▼ Configuration Settings" if checked else "▶ Configuration Settings")
        QTimer.singleShot(10, self.adjustSize)

    def openUserManual(self) -> None:
        MarkdownManualDialog(self).exec()

    def inspectPdfFieldsDialog(self) -> None:
        pdf_path = QFileDialog.getOpenFileName(self, "Inspect PDF Form Fields", "", "PDF Files (*.pdf);;All Files (*.*)")[0]
        if not pdf_path:
            return
        try:
            PdfInspectorDialog(pdf_path, self).exec()
        except Exception as e:
            self.error_window(f"Failed to read PDF:\n{e}")

    def get_configured_root_order(self) -> List[str]:
        order: List[str] = []
        for r in self.questionnaire_rows:
            name = r.name_edit.text().strip()
            if name and name not in order:
                order.append(name)
        return order

    def get_configured_copy_counts(self) -> Dict[str, int]:
        counts: Dict[str, int] = {}
        for r in self.questionnaire_rows:
            s_name = r.name_edit.text().strip()
            if s_name:
                try:
                    c = max(1, int(r.count_edit.text()))
                except ValueError:
                    c = 1
                counts[s_name.lower()] = max(counts.get(s_name.lower(), 1), c)
        return counts

    def openConditionOrderingDialog(self) -> None:
        saved, updated = edit_condition_ordering_dialog(
            parent=self,
            source_path=self.source_path,
            configured_names=self.get_configured_root_order(),
            active_config_name=self.active_config_name,
            condition_ordering=self.condition_ordering
        )
        if saved:
            self.condition_ordering = updated

    def _clear_questionnaire_rows(self) -> None:
        while self.scroll_layout.count():
            item = self.scroll_layout.takeAt(0)
            if item is not None:
                widget = item.widget()
                if widget is not None:
                    widget.deleteLater()
        self.questionnaire_rows.clear()

    def create_questionnaire_rows(self) -> None:
        try:
            count = int(self.quest_count_edit.text())
            if count <= 0:
                raise ValueError
            self._clear_questionnaire_rows()
            for i in range(count):
                row = QuestionnaireRow(i)
                self.questionnaire_rows.append(row)
                base_row = i * 4
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
                row.path_button.clicked.connect(lambda _checked=False, r=row: self.select_template_file(r))
            self.scroll_area.setVisible(True)
            QTimer.singleShot(10, self.adjustSize)
        except ValueError:
            self.error_window("Please enter a valid positive number!")

    def select_template_file(self, row: QuestionnaireRow) -> None:
        file_path = QFileDialog.getOpenFileName(self, "Select Template File", "", "PDF Files (*.pdf);;All Files (*.*)")[0]
        if file_path:
            row.template_path = file_path
            row.path_display.setText(file_path)

    def select_target_folder(self) -> None:
        folder_path = QFileDialog.getExistingDirectory(self, "Select Target Folder")
        if folder_path:
            self.target_path = folder_path
            self.target_pathset.setText(folder_path)

    def select_source_folder(self) -> None:
        folder_path = QFileDialog.getExistingDirectory(self, "Select Source Folder")
        if folder_path:
            self.source_path = folder_path
            self.source_pathset.setText(folder_path)

    def select_excel_file(self) -> None:
        file_filter = (
            "All Supported (*.parquet *.xlsx *.xls *.csv *.tsv *.json);;"
            "Parquet Files (*.parquet);;Excel Files (*.xlsx *.xls);;CSV Files (*.csv);;"
            "TSV Files (*.tsv *.tab);;JSON Files (*.json);;All Files (*.*)"
        )
        file_path = QFileDialog.getOpenFileName(self, "Select Masterfile", "", file_filter)[0]
        if file_path:
            self.excel_path = file_path
            self.excel_pathset.setText(file_path)

    def get_active_participant_ids(self) -> List[str]:
        if self.batch_mode_global.isChecked():
            return self.parseParticipantIDs(self.batch_text_global.toPlainText())
        pid = self.nameset.text().strip()
        return [pid] if pid else []

    def _validate_participant_ids(self, participant_ids: List[str]) -> Optional[str]:
        for pid in participant_ids:
            if pid in ('.', '..') or INVALID_ID_CHARS_PATTERN.search(pid):
                return f"Invalid participant ID '{pid}'.\nParticipant IDs cannot be '.' or '..' or contain: < > : \" / \\ | ? *"
        return None

    def generate_participant_folder(self) -> None:
        try:
            participant_ids = self.get_active_participant_ids()
            if not participant_ids:
                return self.error_window("Please enter at least one participant ID!")
            err = self._validate_participant_ids(participant_ids)
            if err:
                return self.error_window(err)
            if not self.target_path:
                return self.error_window("Please select a target folder! (Expand 'Configuration Settings' to set paths)")
            if not os.path.isdir(self.target_path):
                return self.error_window(f"Target folder does not exist:\n{self.target_path}")
            if not self.questionnaire_rows:
                return self.error_window("Please configure questionnaires! (Expand 'Configuration Settings' to set templates)")
            missing = [r.template_path for r in self.questionnaire_rows if r.template_path and not os.path.isfile(r.template_path)]
            if missing:
                return self.error_window(f"Template file not found:\n{missing[0]}")
            if not any(r.template_path for r in self.questionnaire_rows):
                return self.error_window("Please select at least one questionnaire template file!")

            run_generate_folders(
                self, participant_ids, self.target_path, self.questionnaire_rows,
                self.get_configured_root_order(), self.active_config_name, self.field_mapping
            )
        except Exception as e:
            self.error_window(f"Error generating folder(s): {str(e)}")

    def extract_data(self) -> None:
        if not self.excel_path:
            return self.error_window("Please select a masterfile! (Expand 'Configuration Settings' to select masterfile)")
        if not self.source_path:
            return self.error_window("Please select a source folder! (Expand 'Configuration Settings' to select source folder)")
        if not os.path.isdir(self.source_path):
            return self.error_window(f"Source folder does not exist:\n{self.source_path}")

        if self.active_config_name:
            stored_map = load_mapping_from_config(self.active_config_name)
            if stored_map:
                self.field_mapping.update(stored_map)
            stored_co = load_condition_ordering_from_config(self.active_config_name)
            if stored_co.get("mappings") or stored_co.get("mode") != "numerical":
                self.condition_ordering = stored_co

        participant_ids = self.get_active_participant_ids()
        if not participant_ids:
            return self.error_window("Please enter at least one participant ID!")
        err = self._validate_participant_ids(participant_ids)
        if err:
            return self.error_window(err)

        try:
            self.condition_ordering = run_extract_data(
                self, participant_ids, self.source_path, self.excel_path,
                self.get_configured_root_order(), self.get_configured_copy_counts(),
                self.active_config_name, self.field_mapping, self.condition_ordering
            )
        except Exception as e:
            self.error_window(f"Error extracting data: {str(e)}")

    def importFieldMapping(self, source_config_name: str) -> None:
        source_mapping = load_mapping_from_config(source_config_name)
        if not source_mapping:
            return self.error_window(f"Configuration '{source_config_name}' does not contain any saved field mappings!")
        matched = 0
        for k, v in source_mapping.items():
            if v and v != k:
                self.field_mapping[k] = v
                matched += 1
        if self.active_config_name:
            save_mapping_to_config(self.active_config_name, self.field_mapping)
        QMessageBox.information(
            self, "Field Mapping Imported",
            f"Successfully imported {matched} field mapping(s) from '{source_config_name}'."
            + (f"\nSaved into active configuration '{self.active_config_name}'." if self.active_config_name else "\nRetained in memory.")
        )

    def editFieldMappingFromExtraction(self) -> None:
        expected = get_expected_system_fields(self.questionnaire_rows)
        edit_field_mapping_dialog(self, expected, self.get_configured_root_order(), self.active_config_name, self.field_mapping)

    def updateRecentConfigsMenu(self) -> None:
        for menu in (self.recent_configs_menu, self.delete_configs_menu, self.import_mapping_menu):
            if menu is not None:
                menu.clear()
        cfg = new_config_parser()
        cfg.read(get_config_path(), encoding='utf-8')
        for s in cfg.sections():
            if s == 'PrefillDefaults':
                continue
            if self.recent_configs_menu is not None:
                act = QAction(s, self)
                act.triggered.connect(lambda _checked=False, n=s: self.ConfigLoad(n))
                self.recent_configs_menu.addAction(act)
            if self.delete_configs_menu is not None:
                act_d = QAction(s, self)
                act_d.triggered.connect(lambda _checked=False, n=s: self.confirmDeleteConfig(n))
                self.delete_configs_menu.addAction(act_d)
            if self.import_mapping_menu is not None:
                act_m = QAction(s, self)
                act_m.setToolTip(f"Import friendly column mappings from '{s}' into current session")
                act_m.triggered.connect(lambda _checked=False, n=s: self.importFieldMapping(n))
                self.import_mapping_menu.addAction(act_m)

    def showSaveConfigWindow(self) -> None:
        save_win = SaveConfigWindow(self)
        save_win.signal.connect(self.ConfigGen)
        save_win.exec()

    def confirmDeleteConfig(self, name: str) -> None:
        dlg = QMessageBox(self)
        dlg.setWindowTitle("Confirm Delete")
        dlg.setIcon(QMessageBox.Icon.Question)
        dlg.setText(f"Delete configuration '{name}'?")
        dlg.addButton(QMessageBox.StandardButton.Yes)
        dlg.addButton(QMessageBox.StandardButton.No)
        dlg.setDefaultButton(QMessageBox.StandardButton.No)
        if dlg.exec() == QMessageBox.StandardButton.Yes:
            cfg = new_config_parser()
            cfg.read(get_config_path(), encoding='utf-8')
            cfg.remove_section(name)
            write_config_atomically(cfg, get_config_path())

    def ConfigGen(self, name: str) -> None:
        if not name:
            return self.error_window("No name provided!")
        save_config_profile(
            name, self.target_path, self.questionnaire_rows,
            self.source_path, self.excel_path, self.field_mapping, self.condition_ordering
        )
        self.active_config_name = name
        QMessageBox.information(self, "Success", f"Configuration snapshot '{name}' saved successfully!")

    def ConfigLoad(self, name: str) -> None:
        data = load_config_profile(name)
        if data is None:
            return self.error_window("Configuration not found!")
        self.target_path = data["target_path"]
        self.target_pathset.setText(self.target_path)

        rows_meta = data["rows"]
        if rows_meta:
            self.quest_count_edit.setText(str(len(rows_meta)))
            self.create_questionnaire_rows()
            for i, row in enumerate(self.questionnaire_rows):
                r_name, r_path, r_count = rows_meta[i]
                row.name_edit.setText(r_name)
                row.template_path = r_path
                row.path_display.setText(r_path)
                row.count_edit.setText(r_count)
        else:
            self.quest_count_edit.clear()
            self._clear_questionnaire_rows()
            self.scroll_area.setVisible(False)
            QTimer.singleShot(10, self.adjustSize)

        self.source_path = data["source_path"]
        self.source_pathset.setText(self.source_path)
        self.excel_path = data["excel_path"]
        self.excel_pathset.setText(self.excel_path)
        self.active_config_name = name
        self.field_mapping = data["field_mapping"]
        self.condition_ordering = data["condition_ordering"]
        QMessageBox.information(self, "Success", f"Configuration snapshot '{name}' loaded successfully!")

    def enterWhatsThisMode(self) -> None:
        QWhatsThis.enterWhatsThisMode()

    def toggleBatchModeGlobal(self, checked: bool) -> None:
        self.nameset.setVisible(not checked)
        self.batch_text_global.setVisible(checked)
        QTimer.singleShot(10, self.adjustSize)

    def importParticipantList(self, target_text_edit: QTextEdit) -> None:
        path = QFileDialog.getOpenFileName(self, "Import Participant List", "", "Text/CSV (*.txt *.csv);;All (*.*)")[0]
        if not path:
            return
        with open(path, 'r', encoding='utf-8') as f:
            pids = [p.strip() for line in f for p in line.replace(',', '\n').split('\n') if p.strip()]
        unique = list(dict.fromkeys(pids))
        target_text_edit.setPlainText('\n'.join(unique))
        if unique and not self.batch_mode_global.isChecked():
            self.batch_mode_global.setChecked(True)

    def parseParticipantIDs(self, text: str) -> List[str]:
        return list(dict.fromkeys(line.strip() for line in text.replace(',', '\n').split('\n') if line.strip()))

    def error_window(self, message: str) -> None:
        QMessageBox.warning(self, "ERROR", message)


if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())