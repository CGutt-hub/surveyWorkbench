import sys
import os
import io
import importlib
import shutil
import csv
import json
import re
import tempfile
from typing import Optional, List, Dict, Any, Tuple, Set
from configparser import ConfigParser

from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QPushButton,
    QGridLayout, QWidget, QLineEdit, QLabel,
    QFileDialog, QMessageBox, QAction, QStatusBar,
    QScrollArea, QVBoxLayout, QFrame, QLayout,
    QMenu, QDialog, QWhatsThis, QTextEdit, QTableWidget,
    QTableWidgetItem, QHeaderView, QCheckBox, QHBoxLayout,
    QToolButton, QTextBrowser
)
from PyQt5.QtCore import pyqtSignal, Qt, QTimer
import xlwings as xlw
from pypdf import PdfReader, PdfWriter
from pypdf.generic import NameObject, create_string_object, ArrayObject

# Optional Parquet Engines (PyArrow / Pandas)
pa: Any = None
pq: Any = None
HAS_PYARROW: bool = False
try:
    import pyarrow as _pa
    import pyarrow.parquet as _pq
    pa = _pa
    pq = _pq
    HAS_PYARROW = True
except ImportError:
    pass

pd: Any = None
HAS_PANDAS: bool = False
try:
    pd = importlib.import_module("pandas")
    HAS_PANDAS = True
except ImportError:
    pass

# ==============================================================================
# CANONICAL FORM FIELD GRAMMAR:
#   (_){Check, Text}i(_n)
#
#   (_)           : NOT operator. If present (leading '_'), the field is strictly
#                   ignored and NOT exported to the masterfile.
#   {Check, Text} : Declares entry type. 'Text' is string; 'Check' is boolean/nominal.
#   i             : Index of the condition/question (stays constant across options n).
#   (_n)          : Option index. When selected, the masterfile cell for column
#                   'Check<i>' displays 'n' instead of the internal PDF value.
#                   - Nominal/Likert: Check1_1 to Check1_5 -> selected Check1_3 yields '3'
#                   - Binary: Check2_0 (False/No) and Check2_1 (True/Yes) yield '0' or '1'
# ==============================================================================
CANONICAL_PATTERN = re.compile(r'^(_)?(Text|Check)(\d+)(?:[._]([a-zA-Z0-9]+))?$', re.IGNORECASE)
INVALID_ID_CHARS_PATTERN = re.compile(r'[<>:"/\\|?*\x00-\x1f]')
CONFIG_FILENAME = "survey_workbench.config"


def new_config_parser() -> ConfigParser:
    """Create a ConfigParser without '%' interpolation so paths/headers with '%' never crash."""
    return ConfigParser(interpolation=None)


def set_item_read_only(item: QTableWidgetItem) -> None:
    """Configure a QTableWidgetItem as selectable and enabled, but read-only."""
    item.setFlags(Qt.ItemFlags(Qt.ItemFlag.ItemIsSelectable) | Qt.ItemFlag.ItemIsEnabled)


def get_config_path() -> str:
    """Return path to survey_workbench.config alongside executable or script."""
    if "NUITKA_ONEFILE_DIRECTORY" in os.environ:
        base_dir = os.environ["NUITKA_ONEFILE_DIRECTORY"]
    elif getattr(sys, 'frozen', False) or "__compiled__" in globals():
        exe_path = sys.argv[0] if sys.argv and sys.argv[0] else sys.executable
        base_dir = os.path.dirname(os.path.abspath(exe_path))
    else:
        base_dir = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base_dir, CONFIG_FILENAME)


def ensure_config_file_exists() -> str:
    """Ensure config file exists next to the executable/script."""
    cfg_path = get_config_path()
    if not os.path.exists(cfg_path):
        bundled_cfg = os.path.join(os.path.dirname(os.path.abspath(__file__)), CONFIG_FILENAME)
        if os.path.exists(bundled_cfg) and os.path.abspath(bundled_cfg) != os.path.abspath(cfg_path):
            try:
                shutil.copy2(bundled_cfg, cfg_path)
            except Exception:
                pass
        else:
            try:
                with open(cfg_path, 'w', encoding='utf-8'):
                    pass
            except Exception:
                pass
    return cfg_path


def write_config_atomically(cfg: ConfigParser, cfg_path: str) -> None:
    """Write configuration safely using a temporary file and atomic replace."""
    cfg_dir = os.path.dirname(os.path.abspath(cfg_path)) or "."
    fd, tmp_path = tempfile.mkstemp(prefix=".cfg_", suffix=".tmp", dir=cfg_dir)
    try:
        with os.fdopen(fd, 'w', encoding='utf-8') as f:
            cfg.write(f)
        os.replace(tmp_path, cfg_path)
    except Exception:
        if os.path.exists(tmp_path):
            try:
                os.remove(tmp_path)
            except Exception:
                pass
        raise


def natural_sort_key(s: Any) -> List[Any]:
    """Naturally sort alphanumeric strings so Text2 comes before Text10."""
    return [int(text) if text.isdigit() else text.lower() for text in re.split(r'(\d+)', str(s))]


def match_canonical_field(field_name: str) -> Optional[re.Match[str]]:
    """Match CANONICAL_PATTERN against full name first, then dot-separated suffixes."""
    m = CANONICAL_PATTERN.match(field_name)
    if m:
        return m
    parts = field_name.split('.')
    for idx in range(1, len(parts)):
        candidate = '.'.join(parts[idx:])
        m = CANONICAL_PATTERN.match(candidate)
        if m:
            return m
    return None


def resolve_pdf_object(obj: Any) -> Any:
    """Recursively resolve IndirectObject references from PDF forms."""
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


def find_readme_path() -> Optional[str]:
    """Resolve README.md location across development scripts and compiled binaries."""
    candidates = [
        os.path.join(os.path.dirname(os.path.abspath(__file__)), 'README.md'),
        os.path.join(os.path.dirname(os.path.abspath(sys.executable)), 'README.md'),
        os.path.join(os.getcwd(), 'README.md')
    ]
    if "NUITKA_ONEFILE_DIRECTORY" in os.environ:
        candidates.insert(0, os.path.join(os.environ["NUITKA_ONEFILE_DIRECTORY"], 'README.md'))
    for p in candidates:
        if os.path.exists(p) and os.path.isfile(p):
            return p
    return None


class MarkdownManualDialog(QDialog):
    """In-app viewer that renders README.md directly as the manual."""
    def __init__(self, parent: Optional[QWidget] = None) -> None:
        super().__init__(parent)
        self.setWindowTitle("Survey Workbench - Documentation & User Manual")
        self.resize(880, 680)

        layout = QVBoxLayout()
        self.browser = QTextBrowser()
        self.browser.setOpenExternalLinks(True)

        readme_path = find_readme_path()
        content = ""
        if readme_path:
            try:
                with open(readme_path, 'r', encoding='utf-8') as f:
                    content = f.read()
            except Exception as e:
                content = f"# Error Loading Manual\n\nCould not read file `{readme_path}`: {e}"
        else:
            content = (
                "# Survey Workbench Documentation\n\n"
                "**`README.md` was not found in the application directory.**\n\n"
                "Please ensure `README.md` is placed next to the executable."
            )

        self.browser.setMarkdown(content)
        layout.addWidget(self.browser)

        close_btn = QPushButton("Close")
        close_btn.clicked.connect(self.accept)
        layout.addWidget(close_btn, alignment=Qt.AlignmentFlag.AlignRight)
        self.setLayout(layout)


class OverwriteResolutionDialog(QDialog):
    """Pops up when existing participants are detected, letting users select which to overwrite."""
    def __init__(self, existing_pids: List[str], title: str, prompt: str, parent: Optional[QWidget] = None) -> None:
        super().__init__(parent)
        self.setWindowTitle(title)
        self.setModal(True)
        self.resize(520, 420)

        self.checkboxes: Dict[str, QCheckBox] = {}
        layout = QVBoxLayout()

        prompt_label = QLabel(prompt)
        prompt_label.setWordWrap(True)
        prompt_label.setStyleSheet("font-weight: bold; font-size: 13px; margin-bottom: 5px;")
        layout.addWidget(prompt_label)

        btn_bar = QHBoxLayout()
        select_all_btn = QPushButton("Select All")
        deselect_all_btn = QPushButton("Deselect All")
        select_all_btn.clicked.connect(self.select_all)
        deselect_all_btn.clicked.connect(self.deselect_all)
        btn_bar.addWidget(select_all_btn)
        btn_bar.addWidget(deselect_all_btn)
        btn_bar.addStretch()
        layout.addLayout(btn_bar)

        unique_pids = sorted(dict.fromkeys(existing_pids), key=natural_sort_key)
        table = QTableWidget(len(unique_pids), 2)
        table.setHorizontalHeaderLabels(["Overwrite?", "Participant ID"])
        header = table.horizontalHeader()
        if header is not None:
            header.setStretchLastSection(True)
            header.setSectionResizeMode(0, QHeaderView.ResizeMode.ResizeToContents)

        for idx, pid in enumerate(unique_pids):
            cb = QCheckBox()
            cb.setChecked(True)
            self.checkboxes[pid] = cb

            cb_container = QWidget()
            cb_layout = QHBoxLayout(cb_container)
            cb_layout.addWidget(cb)
            cb_layout.setAlignment(Qt.AlignmentFlag.AlignCenter)
            cb_layout.setContentsMargins(0, 0, 0, 0)

            table.setCellWidget(idx, 0, cb_container)
            pid_item = QTableWidgetItem(pid)
            set_item_read_only(pid_item)
            table.setItem(idx, 1, pid_item)

        layout.addWidget(table)

        action_layout = QHBoxLayout()
        confirm_btn = QPushButton("Confirm Selection")
        confirm_btn.setStyleSheet("font-weight: bold;")
        cancel_btn = QPushButton("Cancel")
        confirm_btn.clicked.connect(self.accept)
        cancel_btn.clicked.connect(self.reject)

        action_layout.addWidget(confirm_btn)
        action_layout.addWidget(cancel_btn)
        layout.addLayout(action_layout)
        self.setLayout(layout)

    def select_all(self) -> None:
        for cb in self.checkboxes.values():
            cb.setChecked(True)

    def deselect_all(self) -> None:
        for cb in self.checkboxes.values():
            cb.setChecked(False)

    def get_selected_pids(self) -> List[str]:
        return [pid for pid, cb in self.checkboxes.items() if cb.isChecked()]


class SaveConfigWindow(QDialog):
    signal = pyqtSignal(str)

    def __init__(self, parent: Optional[QWidget] = None) -> None:
        super().__init__(parent)
        self.setWindowTitle("Save Configuration")
        self.setModal(True)

        self.name: str = ""
        self.config: ConfigParser = new_config_parser()

        layout = QGridLayout()
        self.title3 = QLabel("Set configuration name:")
        layout.addWidget(self.title3, 0, 0)

        self.configset = QLineEdit()
        self.configset.setToolTip("Enter a unique profile name (e.g., 'Pilot_Study')")
        self.configset.setWhatsThis(f"Enter a profile name. All settings will be saved in {CONFIG_FILENAME}.")
        layout.addWidget(self.configset, 0, 1)

        self.save = QPushButton('Save configuration')
        layout.addWidget(self.save, 1, 1)
        self.setLayout(layout)
        self.save.clicked.connect(self.troubleshoot)

    def troubleshoot(self, checked: bool = False) -> None:
        self.name = self.configset.text().strip()
        if not self.name:
            return
        self.config = new_config_parser()
        self.config.read(get_config_path(), encoding='utf-8')
        if self.config.has_section(self.name):
            self.areYouSure()
        else:
            self.confirm()

    def areYouSure(self) -> None:
        dlg = QMessageBox(self)
        dlg.setWindowTitle("Override Configuration")
        dlg.setIcon(QMessageBox.Icon.Question)
        dlg.setText("This configuration exists already.\nDo you want to override this configuration permanently?")
        dlg.addButton(QMessageBox.StandardButton.Yes)
        dlg.addButton(QMessageBox.StandardButton.No)
        dlg.setDefaultButton(QMessageBox.StandardButton.No)

        if dlg.exec() == QMessageBox.StandardButton.Yes:
            self.config.remove_section(self.name)
            write_config_atomically(self.config, get_config_path())
            self.confirm()

    def confirm(self) -> None:
        self.signal.emit(self.name)
        self.accept()


class QuestionnaireRow:
    """Represents a single questionnaire configuration row."""
    def __init__(self, index: int) -> None:
        self.index: int = index
        self.name_label: QLabel = QLabel(f"Survey {index + 1} name:")
        self.name_edit: QLineEdit = QLineEdit()
        self.name_edit.setToolTip("Short name for this survey (e.g., 'Demographics', 'NASA_TLX')")

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

        main_widget = QWidget()
        main_layout = QVBoxLayout()
        main_layout.setSizeConstraint(QLayout.SizeConstraint.SetFixedSize)

        # Menu Actions
        button_action = QAction("&Save Configuration", self)
        button_action.triggered.connect(self.showSaveConfigWindow)

        button_action3 = QAction("User Manual", self)
        button_action3.triggered.connect(self.openUserManual)

        whats_this_action = QAction("What's This? (Click for Help)", self)
        whats_this_action.triggered.connect(self.enterWhatsThisMode)
        whats_this_action.setShortcut("Shift+F1")

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

            help_menu = menubar.addMenu("&HELP")
            if help_menu is not None:
                help_menu.addAction(whats_this_action)
                help_menu.addAction(button_action3)

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
        self.batch_text_global.setPlaceholderText(
            "Enter participant IDs (one per line or comma-separated):\nP001\nP002\nP003"
        )
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

        # Generation Sub-section
        gen_frame = QFrame()
        gen_frame.setFrameShape(QFrame.Shape.Box)
        gen_frame.setFrameShadow(QFrame.Shadow.Raised)
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
        extract_frame.setFrameShape(QFrame.Shape.Box)
        extract_frame.setFrameShadow(QFrame.Shadow.Raised)
        extract_layout = QGridLayout()

        extract_title = QLabel("Extract Data to Masterfile")
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
        self.config_container.setVisible(False)
        main_layout.addWidget(self.config_container)

        # Action Buttons
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
        dlg = MarkdownManualDialog(self)
        dlg.exec()

    def get_configured_root_order(self) -> List[str]:
        order: List[str] = []
        for r in self.questionnaire_rows:
            name = r.name_edit.text().strip()
            if name and name not in order:
                order.append(name)
        return order

    def _get_root_survey_name(self, clean_filename: str) -> str:
        clean = clean_filename.strip()
        for s_name in sorted(self.get_configured_root_order(), key=len, reverse=True):
            if clean == s_name or clean.startswith(s_name):
                return s_name
        m = re.match(r'^(.*?)\d+$', clean)
        if m and m.group(1):
            return m.group(1).rstrip('_')
        return clean

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
            "Parquet Files (*.parquet);;"
            "Excel Files (*.xlsx *.xls);;"
            "CSV Files (*.csv);;"
            "TSV Files (*.tsv *.tab);;"
            "JSON Files (*.json);;"
            "All Files (*.*)"
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
                return (
                    f"Invalid participant ID '{pid}'.\n"
                    "Participant IDs cannot be '.' or '..' or contain characters: < > : \" / \\ | ? *"
                )
        return None

    def _normalize_excel_pid(self, cell_val: Any) -> str:
        if cell_val is None:
            return ''
        if isinstance(cell_val, float) and cell_val.is_integer():
            return str(int(cell_val))
        return str(cell_val).strip()

    def _sanitize_excel_sheet_name(self, name: str) -> str:
        """Sanitize and truncate sheet names to obey Excel's 31-char and character rules."""
        cleaned = re.sub(r'[\\/?*\[\]:]', '_', name).strip()
        if not cleaned:
            cleaned = "Sheet"
        return cleaned[:31]

    def _get_masterfile_siblings(self, base_dir: str, base_name: str, valid_exts: Tuple[str, ...]) -> List[str]:
        if not os.path.isdir(base_dir):
            return []
        candidates: List[str] = []
        for f in os.listdir(base_dir):
            stem, ext = os.path.splitext(f)
            if ext.lower() in valid_exts and (stem == base_name or stem.startswith(f"{base_name}_")):
                candidates.append(os.path.join(base_dir, f))
        return candidates

    def _get_existing_masterfile_pids(self) -> Set[str]:
        """Detect existing participant IDs already in masterfile formats."""
        if not self.excel_path or not os.path.exists(self.excel_path):
            return set()

        file_ext = self.excel_path.lower().split('.')[-1]
        existing_pids: Set[str] = set()
        base_dir = os.path.dirname(self.excel_path)
        base_name = os.path.splitext(os.path.basename(self.excel_path))[0]

        if file_ext in ['xls', 'xlsx']:
            wb: Optional[xlw.Book] = None
            try:
                wb = xlw.Book(self.excel_path)
                for sheet in wb.sheets:
                    r = 2
                    while True:
                        cell_val = sheet.range(f'A{r}').value
                        if cell_val is None:
                            break
                        cell_str = self._normalize_excel_pid(cell_val)
                        if cell_str:
                            existing_pids.add(cell_str)
                        r += 1
            except Exception:
                pass
            finally:
                if wb is not None:
                    try:
                        wb.close()
                    except Exception:
                        pass

        elif file_ext == 'parquet':
            candidate_files = self._get_masterfile_siblings(base_dir, base_name, ('.parquet',))
            if not candidate_files and os.path.exists(self.excel_path):
                candidate_files = [self.excel_path]

            for c_path in candidate_files:
                try:
                    if HAS_PYARROW and pq is not None:
                        table = pq.read_table(c_path, columns=['participant_id'])
                        for val in table['participant_id'].to_pylist():
                            if val is not None and str(val).strip():
                                existing_pids.add(str(val).strip())
                    elif HAS_PANDAS and pd is not None:
                        df = pd.read_parquet(c_path, columns=['participant_id'])
                        for val in df['participant_id'].dropna().astype(str):
                            if val.strip():
                                existing_pids.add(val.strip())
                except Exception:
                    pass

        elif file_ext in ['csv', 'tsv', 'tab']:
            delim = '\t' if file_ext in ['tsv', 'tab'] else ','
            candidate_files = self._get_masterfile_siblings(base_dir, base_name, (f'.{file_ext}',))
            if not candidate_files and os.path.exists(self.excel_path):
                candidate_files = [self.excel_path]

            for c_path in candidate_files:
                try:
                    with open(c_path, 'r', newline='', encoding='utf-8') as f:
                        reader = csv.DictReader(f, delimiter=delim)
                        if reader.fieldnames and 'participant_id' in reader.fieldnames:
                            for row in reader:
                                pid = str(row.get('participant_id', '')).strip()
                                if pid:
                                    existing_pids.add(pid)
                except Exception:
                    pass

        elif file_ext == 'json':
            candidate_files = self._get_masterfile_siblings(base_dir, base_name, ('.json',))
            if not candidate_files and os.path.exists(self.excel_path):
                candidate_files = [self.excel_path]

            for c_path in candidate_files:
                try:
                    with open(c_path, 'r', encoding='utf-8') as f:
                        raw = json.load(f)
                        if isinstance(raw, list):
                            for item in raw:
                                if isinstance(item, dict) and 'participant_id' in item:
                                    pid = str(item['participant_id']).strip()
                                    if pid:
                                        existing_pids.add(pid)
                except Exception:
                    pass

        return existing_pids

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

            missing_templates = [
                row.template_path for row in self.questionnaire_rows
                if row.template_path and not os.path.isfile(row.template_path)
            ]
            if missing_templates:
                return self.error_window(f"Template file not found:\n{missing_templates[0]}")

            if not any(row.template_path for row in self.questionnaire_rows):
                return self.error_window("Please select at least one questionnaire template file!")

            existing_pids = [
                pid for pid in participant_ids
                if os.path.exists(os.path.join(self.target_path, pid))
            ]

            pids_to_process = list(participant_ids)
            if existing_pids:
                dlg = OverwriteResolutionDialog(
                    existing_pids=existing_pids,
                    title="Participant Folder(s) Already Exist",
                    prompt="The following participant folder(s) already exist in the target directory.\n"
                           "Tick the participants you wish to overwrite (unticked will be skipped):",
                    parent=self
                )
                if dlg.exec() != QDialog.DialogCode.Accepted:
                    return

                approved_overwrite = set(dlg.get_selected_pids())
                pids_to_process = [
                    pid for pid in participant_ids
                    if (pid not in existing_pids) or (pid in approved_overwrite)
                ]

            if not pids_to_process:
                QMessageBox.information(self, "No Changes", "No participant folders selected to generate.")
                return

            all_generated_pdf_meta: Dict[str, List[Tuple[str, str]]] = {}
            for pid in pids_to_process:
                surveys_meta: List[Tuple[str, str]] = []
                for row in self.questionnaire_rows:
                    if not row.template_path:
                        continue
                    survey_name = row.name_edit.text().strip() or f"survey_{row.index + 1}"
                    try:
                        copy_count = max(1, int(row.count_edit.text()))
                    except ValueError:
                        copy_count = 1

                    _, ext = os.path.splitext(os.path.basename(row.template_path))
                    for i in range(copy_count):
                        filename = f"{pid}_{survey_name}{ext}" if copy_count == 1 else f"{pid}_{survey_name}{i + 1}{ext}"
                        surveys_meta.append((filename, row.template_path))
                if surveys_meta:
                    all_generated_pdf_meta[pid] = surveys_meta

            if not all_generated_pdf_meta:
                return self.error_window("No valid questionnaires found to generate.")

            self.field_mapping = self._ensure_or_create_field_mapping()
            apply_prefill, batch_prefill_map = self.showBatchGenerationPrefillDialogMeta(all_generated_pdf_meta)
            if not apply_prefill:
                return

            for pid in pids_to_process:
                surveys_for_pid = all_generated_pdf_meta.get(pid, [])
                if not surveys_for_pid:
                    continue
                participant_folder = os.path.join(self.target_path, pid)
                if os.path.exists(participant_folder):
                    shutil.rmtree(participant_folder)
                os.makedirs(participant_folder, exist_ok=True)

                for filename, template_path in surveys_for_pid:
                    dest_path = os.path.join(participant_folder, filename)
                    shutil.copy2(template_path, dest_path)

                    lookup_key = f"{pid}::{filename}"
                    field_vals = batch_prefill_map.get(lookup_key, {})
                    if field_vals:
                        self._write_pdf_form_fields(dest_path, field_vals)

            QMessageBox.information(
                self, "Success",
                f"Successfully generated folders and questionnaires for {len(pids_to_process)} participant(s)!"
            )
        except Exception as e:
            self.error_window(f"Error generating folder(s): {str(e)}")

    def _get_writer_root_object(self, writer: PdfWriter) -> Any:
        root_candidate = getattr(writer, 'root_object', None)
        if root_candidate is None:
            root_candidate = getattr(writer, '_root_object', None)
        if root_candidate is None and hasattr(writer, 'trailer'):
            trailer_obj = getattr(writer, 'trailer')
            if isinstance(trailer_obj, dict):
                root_candidate = trailer_obj.get('/Root')
        return resolve_pdf_object(root_candidate)

    def _write_pdf_form_fields(self, pdf_path: str, field_values: Dict[str, Any]) -> None:
        """Write values into PDF form fields and widget appearances safely and atomically."""
        with open(pdf_path, "rb") as src_f:
            pdf_bytes = io.BytesIO(src_f.read())

        reader = PdfReader(pdf_bytes)
        writer = PdfWriter()
        writer.append(reader)

        def _update_tree(obj_ref: Any, inherited_name: str = "", inherited_ft: str = "") -> None:
            obj = resolve_pdf_object(obj_ref)
            if not isinstance(obj, dict):
                return
            local_t = resolve_pdf_object(obj.get('/T'))
            local_name = str(local_t).strip() if local_t else ""
            if inherited_name and local_name:
                field_name = f"{inherited_name}.{local_name}"
            else:
                field_name = local_name or inherited_name

            local_ft = resolve_pdf_object(obj.get('/FT'))
            ft = str(local_ft).strip() if local_ft else inherited_ft

            lookup_name: Optional[str] = None
            if field_name in field_values:
                lookup_name = field_name
            elif local_name and local_name in field_values:
                lookup_name = local_name

            if lookup_name is not None:
                val_to_set = str(field_values[lookup_name]).strip()
                is_button = (ft == '/Btn') or (ft != '/Tx' and '/AS' in obj)

                if is_button:
                    ap_dict = resolve_pdf_object(obj.get('/AP'))
                    ap_n = resolve_pdf_object(ap_dict.get('/N', {})) if isinstance(ap_dict, dict) else {}
                    valid_states = [str(k) for k in ap_n.keys()] if isinstance(ap_n, dict) else []

                    if val_to_set == '1':
                        on_state = next((k for k in valid_states if k != '/Off'), '/Yes')
                        v = NameObject(on_state if on_state.startswith('/') else f'/{on_state}')
                    elif val_to_set in ('0', ''):
                        v = NameObject('/Off')
                    else:
                        candidate = f'/{val_to_set}' if not val_to_set.startswith('/') else val_to_set
                        if valid_states and candidate not in valid_states and '/Off' in valid_states:
                            v = NameObject('/Off') if '/Kids' not in obj else NameObject(candidate)
                        else:
                            v = NameObject(candidate)

                    if '/Kids' in obj or '/T' in obj:
                        obj[NameObject('/V')] = v
                    if '/Subtype' in obj or '/AP' in obj or '/AS' in obj:
                        obj[NameObject('/AS')] = v
                else:
                    obj[NameObject('/V')] = create_string_object(val_to_set)

            kids = resolve_pdf_object(obj.get('/Kids'))
            if isinstance(kids, list):
                for k_ref in kids:
                    _update_tree(k_ref, field_name, ft)

        try:
            root_obj = self._get_writer_root_object(writer)
            if isinstance(root_obj, dict):
                acroform_ref = root_obj.get('/AcroForm')
                if acroform_ref is not None:
                    acroform = resolve_pdf_object(acroform_ref)
                    if isinstance(acroform, dict):
                        fields_raw = resolve_pdf_object(acroform.get('/Fields'))
                        if isinstance(fields_raw, list):
                            for f in fields_raw:
                                _update_tree(f)
        except Exception:
            pass

        for page in writer.pages:
            try:
                annots = resolve_pdf_object(page.get('/Annots'))
                if isinstance(annots, list):
                    for a_ref in annots:
                        _update_tree(a_ref)
            except Exception:
                pass

        writer.set_need_appearances_writer()

        pdf_dir = os.path.dirname(os.path.abspath(pdf_path)) or "."
        fd, tmp_path = tempfile.mkstemp(prefix=".pdf_", suffix=".tmp", dir=pdf_dir)
        try:
            with os.fdopen(fd, "wb") as output_file:
                writer.write(output_file)
            os.replace(tmp_path, pdf_path)
        except Exception:
            if os.path.exists(tmp_path):
                try:
                    os.remove(tmp_path)
                except Exception:
                    pass
            raise

    def _extract_pdf_form_fields(self, pdf_path: str) -> Dict[str, str]:
        """Deep extract form fields, radio groups, /Opt values, and appearance states."""
        try:
            reader = PdfReader(pdf_path)
        except Exception:
            return {}

        extracted: Dict[str, str] = {}

        def resolve_opt_value(raw: Any, opt_list: Optional[List[Any]]) -> Any:
            if opt_list and raw is not None:
                raw_str = str(raw).strip()
                if raw_str.startswith('/'):
                    idx_str = raw_str[1:].strip()
                    if idx_str.isdigit():
                        idx = int(idx_str)
                        if 0 <= idx < len(opt_list):
                            opt_item = resolve_pdf_object(opt_list[idx])
                            if isinstance(opt_item, list) and opt_item:
                                return resolve_pdf_object(opt_item[0])
                            return opt_item
            return raw

        def clean_value(raw: Any, is_button: bool = False) -> str:
            resolved = resolve_pdf_object(raw)
            if resolved is None:
                return ''
            if isinstance(resolved, ArrayObject):
                parts = [clean_value(x, is_button=is_button) for x in resolved if x is not None]
                return ', '.join([p for p in parts if p != ''])
            text = str(resolved).strip()
            if text.startswith('/'):
                unslashed = text[1:].strip()
                if unslashed.lower() in ('yes', 'on', 'true'):
                    return '1'
                if unslashed.lower() in ('off', 'false'):
                    return ''
                return unslashed
            if is_button and text.lower() in ('off', 'false'):
                return ''
            if is_button and text.lower() in ('yes', 'on', 'true'):
                return '1'
            return text

        def store_field(
            name: Any,
            value: Any,
            is_button: bool = False,
            opt_list: Optional[List[Any]] = None
        ) -> None:
            if not name:
                return
            key = str(name).strip()
            if not key or key == 'File':
                return
            resolved_val = resolve_opt_value(value, opt_list) if is_button else value
            val_str = clean_value(resolved_val, is_button=is_button)
            if key not in extracted:
                extracted[key] = val_str
            elif extracted[key] == '' and val_str != '':
                extracted[key] = val_str

        def walk_field_hierarchy(
            obj_ref: Any,
            inherited_name: str = "",
            inherited_ft: str = "",
            inherited_opt: Optional[List[Any]] = None
        ) -> None:
            obj = resolve_pdf_object(obj_ref)
            if not isinstance(obj, dict):
                return
            local_t = resolve_pdf_object(obj.get('/T'))
            local_name = str(local_t).strip() if local_t else ""

            if inherited_name and local_name:
                field_name = f"{inherited_name}.{local_name}"
            else:
                field_name = local_name or inherited_name

            local_ft = resolve_pdf_object(obj.get('/FT'))
            ft = str(local_ft).strip() if local_ft else inherited_ft

            local_opt = resolve_pdf_object(obj.get('/Opt'))
            opt_list = local_opt if isinstance(local_opt, list) else inherited_opt
            is_button = (ft == '/Btn') or ('/AS' in obj)

            raw_v = resolve_pdf_object(obj.get('/V'))
            raw_as = resolve_pdf_object(obj.get('/AS'))

            if raw_as is not None and str(raw_as).strip() not in ('', '/Off'):
                store_field(field_name, raw_as, is_button=True, opt_list=opt_list)
            elif raw_v is not None and str(raw_v).strip() not in ('', '/Off'):
                store_field(field_name, raw_v, is_button=is_button, opt_list=opt_list)
            else:
                store_field(field_name, '', is_button=is_button)

            kids = resolve_pdf_object(obj.get('/Kids'))
            if isinstance(kids, list):
                for k_ref in kids:
                    walk_field_hierarchy(k_ref, field_name, ft, opt_list)

        try:
            root = resolve_pdf_object(reader.trailer.get('/Root', {}))
            if isinstance(root, dict):
                acro_ref = root.get('/AcroForm')
                if acro_ref is not None:
                    acro = resolve_pdf_object(acro_ref)
                    if isinstance(acro, dict):
                        fields_raw = resolve_pdf_object(acro.get('/Fields'))
                        if isinstance(fields_raw, list):
                            for f in fields_raw:
                                walk_field_hierarchy(f)
        except Exception:
            pass

        try:
            acro_fields = reader.get_fields() or {}
            for fname, fmeta in acro_fields.items():
                if isinstance(fmeta, dict):
                    ft = str(resolve_pdf_object(fmeta.get('/FT', '')))
                    opt_raw = resolve_pdf_object(fmeta.get('/Opt'))
                    opt_list = opt_raw if isinstance(opt_raw, list) else None
                    is_btn = ft == '/Btn'
                    v = resolve_pdf_object(fmeta.get('/V'))
                    if v is not None and str(v).strip() not in ('', '/Off'):
                        store_field(fname, v, is_button=is_btn, opt_list=opt_list)
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
                        if not isinstance(annot, dict):
                            continue
                        subtype = str(resolve_pdf_object(annot.get('/Subtype', '')))
                        if subtype != '/Widget':
                            continue
                        fname_raw = resolve_pdf_object(annot.get('/T'))
                        parent = resolve_pdf_object(annot.get('/Parent'))
                        if fname_raw is None and isinstance(parent, dict):
                            fname_raw = resolve_pdf_object(parent.get('/T'))
                        if fname_raw is not None:
                            ft_raw = resolve_pdf_object(annot.get('/FT'))
                            if not ft_raw and isinstance(parent, dict):
                                ft_raw = resolve_pdf_object(parent.get('/FT'))
                            ft = str(ft_raw) if ft_raw else ''

                            opt_raw = resolve_pdf_object(annot.get('/Opt'))
                            if not isinstance(opt_raw, list) and isinstance(parent, dict):
                                opt_raw = resolve_pdf_object(parent.get('/Opt'))
                            opt_list = opt_raw if isinstance(opt_raw, list) else None
                            is_btn = (ft == '/Btn') or ('/AS' in annot)

                            val_as = resolve_pdf_object(annot.get('/AS'))
                            val_v = resolve_pdf_object(annot.get('/V'))
                            if val_as is not None and str(val_as).strip() not in ('', '/Off'):
                                store_field(fname_raw, val_as, is_button=True, opt_list=opt_list)
                            elif val_v is not None and str(val_v).strip() not in ('', '/Off'):
                                store_field(fname_raw, val_v, is_button=is_btn, opt_list=opt_list)
                    except Exception:
                        continue
        except Exception:
            pass

        return extracted

    def _build_prefill_defaults(self, participant_id: str, fields: Dict[str, Any]) -> Dict[str, Any]:
        defaults: Dict[str, Any] = dict(fields)
        participant_aliases = {"participant_id", "participant-id", "Text13", "_Text13", "Text1", "_Text1"}
        for field_name in fields.keys():
            clean_name = field_name.split('.')[-1].strip()
            if (clean_name in participant_aliases or "participant" in clean_name.lower()) and str(fields.get(field_name, "")).strip() == "":
                defaults[field_name] = participant_id
        return defaults

    def showBatchGenerationPrefillDialogMeta(
        self,
        all_generated_pdf_meta: Dict[str, List[Tuple[str, str]]]
    ) -> Tuple[bool, Dict[str, Dict[str, Any]]]:
        """Presents consolidated prefill table reading directly from templates."""
        dialog = QDialog(self)
        dialog.setWindowTitle("Consolidated Batch Preview & Generation")
        dialog.setModal(True)
        dialog.resize(920, 650)

        layout = QVBoxLayout()
        label = QLabel(f"Preview & Edit Data for {len(all_generated_pdf_meta)} Participant(s)")
        label.setStyleSheet("font-weight: bold; font-size: 14px;")
        layout.addWidget(label)

        table = QTableWidget()
        table.setColumnCount(4)
        table.setHorizontalHeaderLabels(["Participant ID", "File", "Field", "Data"])

        rows: List[Tuple[str, str, str, str, str]] = []
        file_to_extracted_fields: Dict[str, Dict[str, str]] = {}
        template_cache: Dict[str, Dict[str, str]] = {}
        root_order = self.get_configured_root_order()

        for participant_id, meta_list in all_generated_pdf_meta.items():
            sorted_meta = sorted(
                meta_list,
                key=lambda m: (
                    root_order.index(self._get_root_survey_name(os.path.splitext(m[0])[0].replace(f'{participant_id}_', '', 1)))
                    if self._get_root_survey_name(os.path.splitext(m[0])[0].replace(f'{participant_id}_', '', 1)) in root_order
                    else 9999,
                    natural_sort_key(m[0])
                )
            )

            for file_label, template_path in sorted_meta:
                survey_type = os.path.splitext(file_label)[0].replace(f'{participant_id}_', '', 1)

                if template_path not in template_cache:
                    template_cache[template_path] = self._extract_pdf_form_fields(template_path)
                extracted_fields = dict(template_cache[template_path])

                if not extracted_fields:
                    continue

                defaults = self._build_prefill_defaults(participant_id, extracted_fields)
                file_to_extracted_fields[f"{participant_id}::{file_label}"] = extracted_fields

                check_matrix_groups: Dict[str, Dict[str, str]] = {}
                text_fields: Dict[str, str] = {}

                for field_name in sorted(defaults.keys(), key=natural_sort_key):
                    val = str(defaults[field_name])
                    m = match_canonical_field(field_name)
                    if m:
                        is_ignored = bool(m.group(1))
                        kind = m.group(2).capitalize()
                        group_i = m.group(3)
                        opt_n = m.group(4)
                        if kind == 'Check' and opt_n is not None:
                            base_name = f"{'_' if is_ignored else ''}Check{group_i}"
                            group_opts = check_matrix_groups.setdefault(base_name, {})
                            if val != '' or group_opts.get(opt_n, '') == '':
                                group_opts[opt_n] = val
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
                    checked_opt = next((
                        opt for opt, cval in sorted(options.items(), key=lambda it: natural_sort_key(it[0]))
                        if str(cval).strip() != '' and str(cval).strip().lower() not in ('off', 'false')
                    ), '')
                    rows.append((participant_id, file_label, base_name, display_name, checked_opt))

        table.setRowCount(len(rows))
        original_displayed: Dict[Tuple[str, str], str] = {}
        for idx, (pid, file_label, f_name, display_name, value) in enumerate(rows):
            survey_type = os.path.splitext(file_label)[0].replace(f'{pid}_', '', 1)
            original_displayed[(survey_type, f_name)] = display_name

            pid_item = QTableWidgetItem(pid)
            set_item_read_only(pid_item)
            file_item = QTableWidgetItem(file_label)
            set_item_read_only(file_item)
            name_item = QTableWidgetItem(display_name)
            name_item.setFlags(name_item.flags() | Qt.ItemFlag.ItemIsEditable)

            table.setItem(idx, 0, pid_item)
            table.setItem(idx, 1, file_item)
            table.setItem(idx, 2, name_item)
            table.setItem(idx, 3, QTableWidgetItem(value))

        header = table.horizontalHeader()
        if header is not None:
            header.setStretchLastSection(True)
            header.setSectionResizeMode(0, QHeaderView.ResizeMode.ResizeToContents)
            header.setSectionResizeMode(1, QHeaderView.ResizeMode.ResizeToContents)
            header.setSectionResizeMode(2, QHeaderView.ResizeMode.ResizeToContents)
        layout.addWidget(table)

        button_layout = QHBoxLayout()
        confirm_btn = QPushButton("Confirm and Generate All")
        cancel_btn = QPushButton("Cancel")
        button_layout.addWidget(confirm_btn)
        button_layout.addWidget(cancel_btn)
        layout.addLayout(button_layout)
        dialog.setLayout(layout)

        confirmed_flag = [False]
        prefill_values_result: Dict[str, Dict[str, Any]] = {}

        def confirm_and_close() -> None:
            prefill_values: Dict[str, Dict[str, Any]] = {}
            for pid_key, meta_list in all_generated_pdf_meta.items():
                for f_label, _ in meta_list:
                    prefill_values[f"{pid_key}::{f_label}"] = {}

            for row_index in range(table.rowCount()):
                pid_item = table.item(row_index, 0)
                file_item = table.item(row_index, 1)
                name_item = table.item(row_index, 2)
                value_item = table.item(row_index, 3)
                if pid_item is None or file_item is None or name_item is None:
                    continue
                pid = pid_item.text()
                file_name = file_item.text()
                key = f"{pid}::{file_name}"
                underlying_field_name = rows[row_index][2]
                val = value_item.text().strip() if value_item is not None else ""

                survey_type = os.path.splitext(file_name)[0].replace(f'{pid}_', '', 1)
                new_name = name_item.text().strip()
                if new_name and not underlying_field_name.startswith('_') and not underlying_field_name.split('.')[-1].startswith('_'):
                    sys_map_key = f"{survey_type}_{underlying_field_name}"
                    if new_name != original_displayed.get((survey_type, underlying_field_name)) or sys_map_key not in self.field_mapping:
                        self.field_mapping[sys_map_key] = new_name

                raw_fields = file_to_extracted_fields.get(key, {})
                matched_matrix_fields: List[Tuple[str, str]] = []
                clean_underlying = underlying_field_name.lstrip('_')
                if clean_underlying.startswith('Check') and ('_' not in clean_underlying and '.' not in clean_underlying):
                    for raw_f in raw_fields:
                        m = match_canonical_field(raw_f)
                        if m and m.group(4) is not None:
                            f_base = f"{'_' if m.group(1) else ''}Check{m.group(3)}"
                            if f_base == underlying_field_name:
                                matched_matrix_fields.append((raw_f, m.group(4)))

                if matched_matrix_fields:
                    selected_opts = {s.strip() for s in val.split(',') if s.strip()} if val else set()
                    for raw_f, col_opt in matched_matrix_fields:
                        prefill_values[key][raw_f] = '1' if col_opt in selected_opts else '0'
                else:
                    prefill_values[key][underlying_field_name] = val

            if self.active_config_name:
                self._save_mapping_to_config(self.active_config_name, self.field_mapping)

            prefill_values_result.update(prefill_values)
            confirmed_flag[0] = True
            dialog.accept()

        confirm_btn.clicked.connect(confirm_and_close)
        cancel_btn.clicked.connect(dialog.reject)
        dialog.exec()

        if confirmed_flag[0]:
            return True, prefill_values_result
        return False, {}

    def _prepare_data_for_extraction(self, participant_id: str) -> Dict[str, Dict[str, Any]]:
        normalized_source = os.path.normpath(self.source_path)
        candidate_sub = os.path.join(self.source_path, participant_id)
        if os.path.isdir(candidate_sub):
            participant_folder = candidate_sub
        elif os.path.basename(normalized_source) == participant_id:
            participant_folder = self.source_path
        else:
            participant_folder = candidate_sub

        if not os.path.exists(participant_folder):
            raise ValueError(f"Participant folder not found: {participant_folder}")

        survey_files: List[str] = sorted(
            [f for f in os.listdir(participant_folder) if f.lower().endswith('.pdf')],
            key=natural_sort_key
        )
        if not survey_files:
            raise ValueError(f"No readable questionnaire data found for participant '{participant_id}'.")

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
                for f in survey_files
                if self._get_root_survey_name(os.path.splitext(f)[0].replace(f"{participant_id}_", "", 1)) == root_sheet
            ) or clean_type != root_sheet

            prefix = f"{clean_type}_" if has_multiple_copies else ""
            check_groups: Dict[str, Dict[str, str]] = {}

            for raw_k, v in sorted(fields.items(), key=lambda it: natural_sort_key(it[0])):
                if raw_k.startswith('_') or raw_k.split('.')[-1].startswith('_'):
                    continue

                val = str(v).strip()
                m = match_canonical_field(raw_k)
                if m:
                    is_ignored = bool(m.group(1))
                    if is_ignored:
                        continue
                    kind = m.group(2).capitalize()
                    group_i = m.group(3)
                    opt_n = m.group(4)

                    if kind == 'Check':
                        if opt_n is not None:
                            base_name = f"Check{group_i}"
                            group_opts = check_groups.setdefault(base_name, {})
                            existing_v = group_opts.get(opt_n, '')
                            if val != '' or existing_v == '':
                                group_opts[opt_n] = val
                        else:
                            val_lower = val.lower()
                            if val == '' or val_lower in ('off', 'false'):
                                final_val = '0'
                            elif val_lower in ('1', 'yes', 'on', 'true') or val_lower == f"check{group_i}":
                                final_val = '1'
                            else:
                                choice_m = re.search(
                                    r'(?:Check\d+[._]|Choice|Option|Auswahl|Button|Radio|_)+([a-zA-Z0-9]+)$',
                                    val,
                                    re.IGNORECASE
                                )
                                final_val = choice_m.group(1) if choice_m else val

                            target_key = f"{prefix}Check{group_i}"
                            existing_check = sheet_data[root_sheet].get(target_key, '')
                            if final_val != '0' or existing_check == '' or val == '0':
                                sheet_data[root_sheet][target_key] = final_val
                    else:
                        target_key = f"{prefix}Text{group_i}"
                        if val != '' or target_key not in sheet_data[root_sheet]:
                            sheet_data[root_sheet][target_key] = val
                else:
                    if raw_k != 'participant_id':
                        target_key = f"{prefix}{raw_k}"
                        if val != '' or target_key not in sheet_data[root_sheet]:
                            sheet_data[root_sheet][target_key] = val

            for base_name, options in sorted(check_groups.items(), key=lambda it: natural_sort_key(it[0])):
                checked_opts = [
                    opt for opt, cval in sorted(options.items(), key=lambda it: natural_sort_key(it[0]))
                    if str(cval).strip() != '' and str(cval).strip().lower() not in ('off', 'false')
                ]
                checked_val = ', '.join(checked_opts) if checked_opts else ''
                sheet_data[root_sheet][f"{prefix}{base_name}"] = checked_val

        return sheet_data

    def _resolve_mapped_column_name(self, sheet_name: str, sys_key: str) -> str:
        """Resolve a system key to its mapped masterfile column without cross-survey collisions."""
        scoped_key = f"{sheet_name}_{sys_key}"
        if scoped_key in self.field_mapping and self.field_mapping[scoped_key]:
            return self.field_mapping[scoped_key]
        if sys_key in self.field_mapping and self.field_mapping[sys_key]:
            return self.field_mapping[sys_key]
        if '_' in sys_key:
            copy_prefix, bare_field = sys_key.split('_', 1)
            if self._get_root_survey_name(copy_prefix) == sheet_name:
                root_scoped = f"{sheet_name}_{bare_field}"
                if root_scoped in self.field_mapping and self.field_mapping[root_scoped]:
                    return f"{copy_prefix}_{self.field_mapping[root_scoped]}"
        return sys_key

    def showBatchExtractionPreviewDialog(
        self,
        all_sheet_data: Dict[str, Dict[str, Dict[str, Any]]]
    ) -> Tuple[bool, Dict[str, Dict[str, Dict[str, Any]]]]:
        """Displays extracted participants in one unified editable preview dialog before writing."""
        dialog = QDialog(self)
        dialog.setWindowTitle("Preview Masterfile Extraction")
        dialog.setModal(True)
        dialog.resize(920, 650)

        layout = QVBoxLayout()
        label = QLabel(f"Preview & Verify Extraction Data for {len(all_sheet_data)} Participant(s)")
        label.setStyleSheet("font-weight: bold; font-size: 14px;")
        layout.addWidget(label)

        root_order = self.get_configured_root_order()
        flat_rows: List[Tuple[str, str, str, str, str]] = []

        for pid in sorted(all_sheet_data.keys(), key=natural_sort_key):
            sheets_map = all_sheet_data[pid]
            sorted_sheet_names = sorted(
                sheets_map.keys(),
                key=lambda s: (root_order.index(s) if s in root_order else 9999, natural_sort_key(s))
            )
            for sheet_name in sorted_sheet_names:
                data = sheets_map[sheet_name]
                sorted_fields = sorted([k for k in data.keys() if k != 'participant_id'], key=natural_sort_key)
                for k in sorted_fields:
                    v = data[k]
                    mapped_name = self._resolve_mapped_column_name(sheet_name, k)
                    flat_rows.append((pid, sheet_name, k, mapped_name, str(v)))

        table = QTableWidget(len(flat_rows), 5)
        table.setHorizontalHeaderLabels(["Participant ID", "Sheet", "Field", "Column Name", "Data"])

        original_displayed: Dict[Tuple[str, str], str] = {}
        for idx, (pid, s_name, f_key, m_name, val) in enumerate(flat_rows):
            original_displayed[(s_name, f_key)] = m_name

            pi = QTableWidgetItem(pid)
            set_item_read_only(pi)
            si = QTableWidgetItem(s_name)
            set_item_read_only(si)
            fi = QTableWidgetItem(f_key)
            set_item_read_only(fi)
            ni = QTableWidgetItem(m_name)
            ni.setFlags(ni.flags() | Qt.ItemFlag.ItemIsEditable)
            di = QTableWidgetItem(val)
            di.setFlags(di.flags() | Qt.ItemFlag.ItemIsEditable)

            table.setItem(idx, 0, pi)
            table.setItem(idx, 1, si)
            table.setItem(idx, 2, fi)
            table.setItem(idx, 3, ni)
            table.setItem(idx, 4, di)

        header = table.horizontalHeader()
        if header is not None:
            header.setStretchLastSection(True)
            header.setSectionResizeMode(0, QHeaderView.ResizeMode.ResizeToContents)
            header.setSectionResizeMode(1, QHeaderView.ResizeMode.ResizeToContents)
            header.setSectionResizeMode(2, QHeaderView.ResizeMode.ResizeToContents)
            header.setSectionResizeMode(3, QHeaderView.ResizeMode.ResizeToContents)
        layout.addWidget(table)

        btn_layout = QHBoxLayout()
        c_btn = QPushButton("Confirm and Extract All")
        c_btn.setStyleSheet("font-weight: bold;")
        can_btn = QPushButton("Cancel")
        btn_layout.addWidget(c_btn)
        btn_layout.addWidget(can_btn)
        layout.addLayout(btn_layout)
        dialog.setLayout(layout)

        confirmed_flag: List[bool] = [False]
        edited_output: Dict[str, Dict[str, Dict[str, Any]]] = {}
        for pid, s_map in all_sheet_data.items():
            edited_output[pid] = {s: {'participant_id': pid} for s in s_map}

        def save_and_close() -> None:
            for r in range(table.rowCount()):
                pid_item = table.item(r, 0)
                sheet_item = table.item(r, 1)
                field_item = table.item(r, 2)
                name_item = table.item(r, 3)
                val_item = table.item(r, 4)
                if pid_item is None or sheet_item is None or field_item is None:
                    continue

                pid = pid_item.text()
                s = sheet_item.text()
                k = field_item.text()
                new_n = name_item.text().strip() if name_item is not None else ""
                val = val_item.text() if val_item is not None else ""

                edited_output.setdefault(pid, {}).setdefault(s, {'participant_id': pid})[k] = val
                if new_n:
                    scoped_key = f"{s}_{k}"
                    if new_n != original_displayed.get((s, k)) or scoped_key not in self.field_mapping:
                        self.field_mapping[scoped_key] = new_n

            if self.active_config_name:
                self._save_mapping_to_config(self.active_config_name, self.field_mapping)

            confirmed_flag[0] = True
            dialog.accept()

        c_btn.clicked.connect(save_and_close)
        can_btn.clicked.connect(dialog.reject)
        dialog.exec()
        return confirmed_flag[0], edited_output if confirmed_flag[0] else all_sheet_data

    def _build_export_row(self, pid: str, sheet_name: str, raw_data: Dict[str, Any]) -> Tuple[Dict[str, Any], Dict[str, str]]:
        export_data: Dict[str, Any] = {'participant_id': pid}
        system_to_target_header: Dict[str, str] = {'participant_id': 'participant_id'}

        sorted_field_keys = sorted(
            [k for k in raw_data.keys() if k != 'participant_id'],
            key=natural_sort_key
        )
        for sys_key in sorted_field_keys:
            final_col = self._resolve_mapped_column_name(sheet_name, sys_key)
            export_data[final_col] = raw_data[sys_key]
            system_to_target_header[sys_key] = final_col

        return export_data, system_to_target_header

    def _export_excel_batch(
        self,
        participant_ids: List[str],
        all_participants_sheet_data: Dict[str, Dict[str, Dict[str, Any]]],
        root_order: List[str]
    ) -> None:
        wb: Optional[xlw.Book] = None
        try:
            if os.path.exists(self.excel_path):
                wb = xlw.Book(self.excel_path)
            else:
                wb = xlw.Book()
                wb.save(self.excel_path)

            for pid in participant_ids:
                sheet_data = all_participants_sheet_data.get(pid, {})
                pid_sheet_names = sorted(
                    sheet_data.keys(),
                    key=lambda s: (root_order.index(s) if s in root_order else 9999, natural_sort_key(s))
                )

                for sheet_name in pid_sheet_names:
                    raw_data = sheet_data[sheet_name]
                    export_data, system_to_target_header = self._build_export_row(pid, sheet_name, raw_data)

                    excel_sheet_name = self._sanitize_excel_sheet_name(sheet_name)
                    existing_sheet_map: Dict[str, xlw.Sheet] = {
                        str(s.name).strip().lower(): s for s in wb.sheets
                    }
                    matched_sheet = existing_sheet_map.get(excel_sheet_name.lower())
                    header_to_col: Dict[str, int] = {}
                    sheet: xlw.Sheet

                    if matched_sheet is not None:
                        sheet = matched_sheet
                        first_row = sheet.range('1:1').value
                        if first_row:
                            row_vals = first_row if isinstance(first_row, list) else [first_row]
                            last_non_empty = -1
                            for idx, v in enumerate(row_vals):
                                if v is not None and str(v).strip():
                                    last_non_empty = idx
                            if last_non_empty >= 0:
                                for idx in range(last_non_empty + 1):
                                    v = row_vals[idx]
                                    if v is not None and str(v).strip():
                                        header_to_col[str(v).strip()] = idx + 1

                        for sys_k, target_h in system_to_target_header.items():
                            if target_h not in header_to_col:
                                if sys_k in header_to_col:
                                    col_idx = header_to_col[sys_k]
                                    sheet.range((1, col_idx)).value = target_h
                                    del header_to_col[sys_k]
                                    header_to_col[target_h] = col_idx
                                else:
                                    next_col = max(header_to_col.values(), default=0) + 1
                                    sheet.range((1, next_col)).value = target_h
                                    header_to_col[target_h] = next_col

                        target_row: Optional[int] = None
                        r_idx = 2
                        while True:
                            cell_val = sheet.range(f'A{r_idx}').value
                            if cell_val is None:
                                break
                            if self._normalize_excel_pid(cell_val) == str(pid):
                                target_row = r_idx
                                break
                            r_idx += 1

                        write_row = target_row if target_row is not None else r_idx
                    else:
                        if (
                            len(wb.sheets) == 1
                            and wb.sheets[0].range('A1').value is None
                            and wb.sheets[0].used_range.value is None
                        ):
                            sheet = wb.sheets[0]
                            sheet.name = excel_sheet_name
                        else:
                            sheet = wb.sheets.add(name=excel_sheet_name, after=wb.sheets[-1])

                        headers = list(export_data.keys())
                        for idx, header_name in enumerate(headers, start=1):
                            sheet.range((1, idx)).value = header_name
                            header_to_col[header_name] = idx

                        write_row = 2

                    for col_header, val in export_data.items():
                        c_idx = header_to_col.get(col_header)
                        if c_idx is not None:
                            sheet.range((write_row, c_idx)).value = val

            wb.save()
        finally:
            if wb is not None:
                try:
                    wb.close()
                except Exception:
                    pass

    def _export_parquet_batch(
        self,
        participant_ids: List[str],
        all_participants_sheet_data: Dict[str, Dict[str, Dict[str, Any]]],
        ordered_sheet_names: List[str]
    ) -> None:
        if not HAS_PYARROW and not HAS_PANDAS:
            raise ValueError("Exporting to Parquet format requires 'pyarrow' (or 'pandas').")

        base_dir = os.path.dirname(os.path.abspath(self.excel_path)) or "."
        base_name = os.path.splitext(os.path.basename(self.excel_path))[0]

        for sheet_name in ordered_sheet_names:
            target_parquet = os.path.join(base_dir, f"{base_name}_{sheet_name}.parquet")
            existing_rows: List[Dict[str, Any]] = []

            if os.path.exists(target_parquet):
                try:
                    if HAS_PYARROW and pq is not None:
                        existing_rows = pq.read_table(target_parquet).to_pylist()
                    elif HAS_PANDAS and pd is not None:
                        existing_rows = pd.read_parquet(target_parquet).to_dict(orient='records')
                except Exception:
                    existing_rows = []

            for pid in participant_ids:
                raw_data = all_participants_sheet_data.get(pid, {}).get(sheet_name)
                if not raw_data:
                    continue

                export_data, _ = self._build_export_row(pid, sheet_name, raw_data)

                existing_index = -1
                for idx, row in enumerate(existing_rows):
                    if str(row.get('participant_id', '')).strip() == str(pid):
                        existing_index = idx
                        break

                if existing_index != -1:
                    existing_rows[existing_index].update(export_data)
                else:
                    existing_rows.append(export_data)

            if existing_rows:
                all_cols = ['participant_id']
                for r in existing_rows:
                    for c in r.keys():
                        if c not in all_cols:
                            all_cols.append(c)

                normalized_rows = [
                    {c: str(r.get(c, '')) if r.get(c) is not None else '' for c in all_cols}
                    for r in existing_rows
                ]

                fd, tmp_path = tempfile.mkstemp(prefix=".pq_", suffix=".tmp", dir=base_dir)
                os.close(fd)
                try:
                    if HAS_PYARROW and pa is not None and pq is not None:
                        pa_table = pa.Table.from_pylist(normalized_rows)
                        pq.write_table(pa_table, tmp_path)
                    elif HAS_PANDAS and pd is not None:
                        df = pd.DataFrame(normalized_rows)
                        df.to_parquet(tmp_path, index=False)
                    os.replace(tmp_path, target_parquet)
                except Exception:
                    if os.path.exists(tmp_path):
                        try:
                            os.remove(tmp_path)
                        except Exception:
                            pass
                    raise

    def _export_delimited_batch(
        self,
        participant_ids: List[str],
        all_participants_sheet_data: Dict[str, Dict[str, Dict[str, Any]]],
        ordered_sheet_names: List[str],
        file_ext: str
    ) -> None:
        delim = '\t' if file_ext in ['tsv', 'tab'] else ','
        base_dir = os.path.dirname(os.path.abspath(self.excel_path)) or "."
        base_name = os.path.splitext(os.path.basename(self.excel_path))[0]

        for sheet_name in ordered_sheet_names:
            target_file = os.path.join(base_dir, f"{base_name}_{sheet_name}.{file_ext}")
            file_exists = os.path.exists(target_file) and os.path.getsize(target_file) > 0

            existing_headers: List[str] = []
            existing_rows: List[Dict[str, Any]] = []
            if file_exists:
                with open(target_file, 'r', newline='', encoding='utf-8') as f:
                    reader = csv.DictReader(f, delimiter=delim)
                    existing_headers = list(reader.fieldnames) if reader.fieldnames else []
                    existing_rows = list(reader)

            all_headers: List[str] = list(existing_headers)

            for pid in participant_ids:
                raw_data = all_participants_sheet_data.get(pid, {}).get(sheet_name)
                if not raw_data:
                    continue

                export_data, _ = self._build_export_row(pid, sheet_name, raw_data)

                existing_index = -1
                for idx, row in enumerate(existing_rows):
                    if str(row.get('participant_id', '')).strip() == str(pid):
                        existing_index = idx
                        break

                if existing_index != -1:
                    existing_rows[existing_index].update(export_data)
                else:
                    existing_rows.append(export_data)

                for k in export_data.keys():
                    if k not in all_headers:
                        all_headers.append(k)

            if existing_rows:
                if 'participant_id' in all_headers:
                    all_headers = ['participant_id'] + [h for h in all_headers if h != 'participant_id']

                fd, tmp_path = tempfile.mkstemp(prefix=".csv_", suffix=".tmp", dir=base_dir)
                try:
                    with os.fdopen(fd, 'w', newline='', encoding='utf-8') as f:
                        writer = csv.DictWriter(f, fieldnames=all_headers, delimiter=delim, restval='')
                        writer.writeheader()
                        writer.writerows(existing_rows)
                    os.replace(tmp_path, target_file)
                except Exception:
                    if os.path.exists(tmp_path):
                        try:
                            os.remove(tmp_path)
                        except Exception:
                            pass
                    raise

    def _export_json_batch(
        self,
        participant_ids: List[str],
        all_participants_sheet_data: Dict[str, Dict[str, Dict[str, Any]]],
        ordered_sheet_names: List[str]
    ) -> None:
        base_dir = os.path.dirname(os.path.abspath(self.excel_path)) or "."
        base_name = os.path.splitext(os.path.basename(self.excel_path))[0]

        for sheet_name in ordered_sheet_names:
            target_json = os.path.join(base_dir, f"{base_name}_{sheet_name}.json")

            existing_rows: List[Dict[str, Any]] = []
            if os.path.exists(target_json):
                try:
                    with open(target_json, 'r', encoding='utf-8') as f:
                        loaded = json.load(f)
                        if isinstance(loaded, list):
                            existing_rows = [r for r in loaded if isinstance(r, dict)]
                except Exception:
                    existing_rows = []

            for pid in participant_ids:
                raw_data = all_participants_sheet_data.get(pid, {}).get(sheet_name)
                if not raw_data:
                    continue

                export_data, _ = self._build_export_row(pid, sheet_name, raw_data)

                existing_index = -1
                for idx, row in enumerate(existing_rows):
                    if str(row.get('participant_id', '')).strip() == str(pid):
                        existing_index = idx
                        break

                if existing_index != -1:
                    existing_rows[existing_index].update(export_data)
                else:
                    existing_rows.append(export_data)

            if existing_rows:
                fd, tmp_path = tempfile.mkstemp(prefix=".json_", suffix=".tmp", dir=base_dir)
                try:
                    with os.fdopen(fd, 'w', encoding='utf-8') as f:
                        json.dump(existing_rows, f, indent=2, ensure_ascii=False)
                    os.replace(tmp_path, target_json)
                except Exception:
                    if os.path.exists(tmp_path):
                        try:
                            os.remove(tmp_path)
                        except Exception:
                            pass
                    raise

    def _readout_batch(
        self,
        participant_ids: List[str],
        all_participants_sheet_data: Dict[str, Dict[str, Dict[str, Any]]]
    ) -> None:
        """Writes all processed participants in a single session across supported formats."""
        if not self.excel_path:
            raise ValueError("Please select a masterfile! (Expand 'Configuration Settings' to select masterfile)")

        file_ext = self.excel_path.lower().split('.')[-1]
        root_order = self.get_configured_root_order()

        all_sheet_names: Set[str] = set()
        for pid in participant_ids:
            all_sheet_names.update(all_participants_sheet_data.get(pid, {}).keys())

        ordered_sheet_names = sorted(
            all_sheet_names,
            key=lambda s: (root_order.index(s) if s in root_order else 9999, natural_sort_key(s))
        )

        if file_ext in ['xls', 'xlsx']:
            self._export_excel_batch(participant_ids, all_participants_sheet_data, root_order)
        elif file_ext == 'parquet':
            self._export_parquet_batch(participant_ids, all_participants_sheet_data, ordered_sheet_names)
        elif file_ext in ['csv', 'tsv', 'tab']:
            self._export_delimited_batch(participant_ids, all_participants_sheet_data, ordered_sheet_names, file_ext)
        elif file_ext == 'json':
            self._export_json_batch(participant_ids, all_participants_sheet_data, ordered_sheet_names)
        else:
            raise ValueError(f"Unsupported masterfile format: {file_ext}")

    def extract_data(self) -> None:
        if not self.excel_path:
            return self.error_window("Please select a masterfile! (Expand 'Configuration Settings' to select masterfile)")
        if not self.source_path:
            return self.error_window("Please select a source folder! (Expand 'Configuration Settings' to select source folder)")
        if not os.path.isdir(self.source_path):
            return self.error_window(f"Source folder does not exist:\n{self.source_path}")

        self.field_mapping = self._load_field_mapping_for_extraction()
        participant_ids = self.get_active_participant_ids()
        if not participant_ids:
            return self.error_window("Please enter at least one participant ID!")
        err = self._validate_participant_ids(participant_ids)
        if err:
            return self.error_window(err)

        try:
            existing_masterfile_pids = self._get_existing_masterfile_pids()
            conflicting_pids = [pid for pid in participant_ids if pid in existing_masterfile_pids]

            pids_to_extract = list(participant_ids)
            if conflicting_pids:
                dlg = OverwriteResolutionDialog(
                    existing_pids=conflicting_pids,
                    title="Masterfile Records Already Exist",
                    prompt="The following participant(s) already exist in your masterfile.\n"
                           "Tick the participants you wish to overwrite (unticked will be skipped):",
                    parent=self
                )
                if dlg.exec() != QDialog.DialogCode.Accepted:
                    return

                approved_overwrite = set(dlg.get_selected_pids())
                pids_to_extract = [
                    pid for pid in participant_ids
                    if (pid not in conflicting_pids) or (pid in approved_overwrite)
                ]

            if not pids_to_extract:
                QMessageBox.information(self, "No Changes", "No participants selected to extract.")
                return

            all_sheet_data: Dict[str, Dict[str, Dict[str, Any]]] = {}
            for pid in pids_to_extract:
                all_sheet_data[pid] = self._prepare_data_for_extraction(pid)

            confirmed, edited_data = self.showBatchExtractionPreviewDialog(all_sheet_data)
            if not confirmed:
                return

            self._readout_batch(pids_to_extract, edited_data)
            QMessageBox.information(
                self, "Extraction Complete",
                f"Successfully extracted {len(pids_to_extract)} participant(s) into masterfile!"
            )
        except Exception as e:
            self.error_window(f"Error extracting data: {str(e)}")

    def _get_expected_system_fields(self) -> Set[str]:
        expected: Set[str] = set()

        for row in self.questionnaire_rows:
            template = row.template_path.strip()
            if not template or not template.lower().endswith('.pdf') or not os.path.isfile(template):
                continue

            template_fields = self._extract_pdf_form_fields(template)
            if not template_fields:
                continue

            survey_name = row.name_edit.text().strip() or f"survey_{row.index + 1}"
            try:
                copy_count = max(1, int(row.count_edit.text()))
            except ValueError:
                copy_count = 1

            for copy_i in range(copy_count):
                clean_type = survey_name if copy_count == 1 else f"{survey_name}{copy_i + 1}"
                prefix = f"{clean_type}_"

                check_matrix_bases: Set[str] = set()
                for field_name in template_fields.keys():
                    if field_name.startswith('_') or field_name.split('.')[-1].startswith('_'):
                        continue

                    m = match_canonical_field(field_name)
                    if m:
                        is_ignored = bool(m.group(1))
                        if is_ignored:
                            continue
                        kind = m.group(2).capitalize()
                        group_i = m.group(3)

                        if kind == 'Check':
                            check_matrix_bases.add(f"Check{group_i}")
                        else:
                            expected.add(f"{prefix}Text{group_i}")
                    else:
                        if field_name != 'participant_id':
                            expected.add(f"{prefix}{field_name}")

                for base_name in check_matrix_bases:
                    expected.add(f"{prefix}{base_name}")

        return expected

    def _ensure_or_create_field_mapping(self) -> Dict[str, str]:
        expected = self._get_expected_system_fields()
        if not expected:
            return self.field_mapping

        stored = self._load_mapping_from_config(self.active_config_name) if self.active_config_name else {}
        mapping = dict(self.field_mapping)
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
        cfg = new_config_parser()
        cfg.read(get_config_path(), encoding='utf-8')
        if not cfg.has_section(config_name):
            return {}
        raw = cfg.get(config_name, 'field_mapping_json', fallback='')
        if raw.strip():
            try:
                loaded = json.loads(raw)
                if isinstance(loaded, dict):
                    return {str(k): str(v) for k, v in loaded.items()}
            except Exception:
                return {}
        return {}

    def _save_mapping_to_config(self, config_name: str, mapping: Dict[str, str]) -> None:
        if not config_name:
            return
        cfg = new_config_parser()
        cfg.read(get_config_path(), encoding='utf-8')
        if not cfg.has_section(config_name):
            cfg.add_section(config_name)
        cfg.set(config_name, 'field_mapping_json', json.dumps(mapping, ensure_ascii=False))
        write_config_atomically(cfg, get_config_path())

    def importFieldMapping(self, source_config_name: str) -> None:
        source_mapping = self._load_mapping_from_config(source_config_name)
        if not source_mapping:
            return self.error_window(f"Configuration '{source_config_name}' does not contain any saved field mappings!")

        matched_count = 0
        for k, v in source_mapping.items():
            if v and v != k:
                self.field_mapping[k] = v
                matched_count += 1

        if self.active_config_name:
            self._save_mapping_to_config(self.active_config_name, self.field_mapping)

        QMessageBox.information(
            self, "Field Mapping Imported",
            f"Successfully imported {matched_count} field mapping(s) from '{source_config_name}'."
            + (f"\nSaved into active configuration '{self.active_config_name}'." if self.active_config_name else "\nRetained in memory.")
        )

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
        label = QLabel("Map system field identifiers to user-friendly column headers:")
        label.setStyleSheet("font-weight: bold;")
        lay.addWidget(label)

        table = QTableWidget(len(mapping), 2)
        table.setHorizontalHeaderLabels(["System Field", "Masterfile Column Name"])
        for idx, (k, v) in enumerate(mapping.items()):
            item_k = QTableWidgetItem(k)
            set_item_read_only(item_k)
            table.setItem(idx, 0, item_k)
            table.setItem(idx, 1, QTableWidgetItem(v))
        header = table.horizontalHeader()
        if header is not None:
            header.setStretchLastSection(True)
            header.setSectionResizeMode(0, QHeaderView.ResizeMode.ResizeToContents)
        lay.addWidget(table)

        btn = QPushButton("Save Field Mapping")

        def save() -> None:
            updated: Dict[str, str] = {}
            for r in range(table.rowCount()):
                key_item = table.item(r, 0)
                val_item = table.item(r, 1)
                if key_item is None:
                    continue
                k = key_item.text()
                val = val_item.text().strip() if val_item is not None else ""
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
        if self.recent_configs_menu is not None:
            self.recent_configs_menu.clear()
        if self.delete_configs_menu is not None:
            self.delete_configs_menu.clear()
        if self.import_mapping_menu is not None:
            self.import_mapping_menu.clear()

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
        cfg = new_config_parser()
        cfg.read(get_config_path(), encoding='utf-8')
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

        write_config_atomically(cfg, get_config_path())

        self.active_config_name = name
        QMessageBox.information(self, "Success", f"Configuration snapshot '{name}' saved successfully!")

    def ConfigLoad(self, name: str) -> None:
        cfg = new_config_parser()
        cfg.read(get_config_path(), encoding='utf-8')
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
        else:
            self.quest_count_edit.clear()
            self._clear_questionnaire_rows()
            self.scroll_area.setVisible(False)
            QTimer.singleShot(10, self.adjustSize)

        self.source_path = cfg.get(name, 'source_path', fallback='')
        self.source_pathset.setText(self.source_path)
        self.excel_path = cfg.get(name, 'excel_path', fallback='')
        self.excel_pathset.setText(self.excel_path)

        self.active_config_name = name
        self.field_mapping = self._load_mapping_from_config(name)
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
        pids: List[str] = []
        with open(path, 'r', encoding='utf-8') as f:
            for line in f:
                for p in line.replace(',', '\n').split('\n'):
                    if p.strip():
                        pids.append(p.strip())
        unique = list(dict.fromkeys(pids))
        target_text_edit.setPlainText('\n'.join(unique))
        if unique and not self.batch_mode_global.isChecked():
            self.batch_mode_global.setChecked(True)

    def parseParticipantIDs(self, text: str) -> List[str]:
        raw = [line.strip() for line in text.replace(',', '\n').split('\n') if line.strip()]
        return list(dict.fromkeys(raw))

    def error_window(self, message: str) -> None:
        QMessageBox.warning(self, "ERROR", message)


if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())