import os
from configparser import ConfigParser
from typing import Dict, List, Optional, Set

from PyQt5.QtCore import Qt, pyqtSignal
from PyQt5.QtWidgets import (
    QCheckBox, QDialog, QGridLayout, QHBoxLayout, QHeaderView,
    QLabel, QLineEdit, QMessageBox, QPushButton, QTableWidget,
    QTableWidgetItem, QTextBrowser, QVBoxLayout, QWidget
)

from config_utils import (
    CONFIG_FILENAME,
    find_readme_path,
    get_config_path,
    natural_sort_key,
    new_config_parser,
    set_item_read_only,
    write_config_atomically,
)
from pdf_handler import inspect_pdf_raw_fields


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


class BatchParticipantSelectionDialog(QDialog):
    """Presents all batch participants with checkboxes so users can untick any they do not want to export."""
    def __init__(
        self,
        all_pids: List[str],
        conflicting_pids: Set[str],
        parent: Optional[QWidget] = None
    ) -> None:
        super().__init__(parent)
        self.setWindowTitle("Select Participants to Export")
        self.setModal(True)
        self.resize(580, 460)

        self.checkboxes: Dict[str, QCheckBox] = {}
        self.conflicting_pids: Set[str] = conflicting_pids
        layout = QVBoxLayout()

        prompt_label = QLabel(
            "Select which participant(s) to export to the masterfile.\n"
            "Untick any participant you do not want to export from this batch:"
        )
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
        if conflicting_pids:
            select_new_btn = QPushButton("Select New Only")
            select_new_btn.clicked.connect(self.select_new_only)
            btn_bar.addWidget(select_new_btn)
        btn_bar.addStretch()
        layout.addLayout(btn_bar)

        unique_pids = sorted(dict.fromkeys(all_pids), key=natural_sort_key)
        table = QTableWidget(len(unique_pids), 3)
        table.setHorizontalHeaderLabels(["Export?", "Participant ID", "Masterfile Status"])
        header = table.horizontalHeader()
        if header is not None:
            header.setSectionResizeMode(0, QHeaderView.ResizeMode.ResizeToContents)
            header.setSectionResizeMode(1, QHeaderView.ResizeMode.ResizeToContents)
            header.setStretchLastSection(True)

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

            status_text = "Already in Masterfile (will overwrite)" if pid in conflicting_pids else "New"
            status_item = QTableWidgetItem(status_text)
            set_item_read_only(status_item)
            table.setItem(idx, 2, status_item)

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

    def select_new_only(self) -> None:
        for pid, cb in self.checkboxes.items():
            cb.setChecked(pid not in self.conflicting_pids)

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


class PdfInspectorDialog(QDialog):
    """Diagnostic dialog to inspect raw /FT, /V, /AS, /AP states and extracted values of any PDF."""
    def __init__(self, pdf_path: str, parent: Optional[QWidget] = None) -> None:
        super().__init__(parent)
        self.setWindowTitle(f"PDF Field Inspector - {os.path.basename(pdf_path)}")
        self.resize(900, 560)

        rows = inspect_pdf_raw_fields(pdf_path)
        lay = QVBoxLayout()
        lay.addWidget(QLabel(f"File: {pdf_path}"))

        table = QTableWidget(len(rows), 6)
        table.setHorizontalHeaderLabels(["Field Name", "/FT", "/V (raw)", "/AS (raw)", "/AP /N States", "Extracted Value"])
        for idx, row_vals in enumerate(rows):
            for c_i, txt in enumerate(row_vals):
                it = QTableWidgetItem(txt)
                set_item_read_only(it)
                table.setItem(idx, c_i, it)

        hdr = table.horizontalHeader()
        if hdr is not None:
            hdr.setStretchLastSection(True)
            for c_i in range(5):
                hdr.setSectionResizeMode(c_i, QHeaderView.ResizeMode.ResizeToContents)
        lay.addWidget(table)

        close_btn = QPushButton("Close")
        close_btn.clicked.connect(self.accept)
        lay.addWidget(close_btn, alignment=Qt.AlignmentFlag.AlignRight)
        self.setLayout(lay)