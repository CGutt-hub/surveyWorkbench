import os
from typing import Any, Dict, List, Optional, Tuple

from PyQt5.QtCore import Qt
from PyQt5.QtWidgets import (
    QDialog, QHBoxLayout, QHeaderView, QLabel, QPushButton,
    QTableWidget, QTableWidgetItem, QVBoxLayout, QWidget
)

from config_utils import (
    derive_root_survey_name,
    is_ignored_field_name,
    is_unchecked_button_value,
    match_canonical_field,
    natural_sort_key,
    save_mapping_to_config,
    set_item_read_only,
    strip_pid_from_stem,
)
from data_extraction import resolve_mapped_column_name
from pdf_handler import build_prefill_defaults, extract_pdf_form_fields


def show_batch_generation_prefill_dialog(
    parent: Optional[QWidget],
    all_generated_pdf_meta: Dict[str, List[Tuple[str, str]]],
    root_order: List[str],
    active_config_name: str,
    field_mapping: Dict[str, str]
) -> Tuple[bool, Dict[str, Dict[str, Any]]]:
    """Presents consolidated prefill table reading directly from templates."""
    dialog = QDialog(parent)
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

    for participant_id, meta_list in all_generated_pdf_meta.items():
        sorted_meta = sorted(
            meta_list,
            key=lambda m: (
                root_order.index(derive_root_survey_name(strip_pid_from_stem(os.path.splitext(m[0])[0], participant_id), root_order))
                if derive_root_survey_name(strip_pid_from_stem(os.path.splitext(m[0])[0], participant_id), root_order) in root_order
                else 9999,
                natural_sort_key(m[0])
            )
        )

        for file_label, template_path in sorted_meta:
            survey_type = strip_pid_from_stem(os.path.splitext(file_label)[0], participant_id)

            if template_path not in template_cache:
                template_cache[template_path] = extract_pdf_form_fields(template_path)
            extracted_fields = dict(template_cache[template_path])
            if not extracted_fields:
                continue

            defaults = build_prefill_defaults(participant_id, extracted_fields)
            file_to_extracted_fields[f"{participant_id}::{file_label}"] = extracted_fields

            check_matrix_groups: Dict[str, Dict[str, str]] = {}
            text_fields: Dict[str, str] = {}

            for field_name in sorted(defaults.keys(), key=natural_sort_key):
                val = str(defaults[field_name])
                m = match_canonical_field(field_name)
                if m:
                    is_ignored = bool(m.group(1)) or is_ignored_field_name(field_name)
                    kind = m.group(2).capitalize()
                    group_i = m.group(3)
                    opt_n = m.group(4)
                    if kind == 'Check' and opt_n is not None:
                        base_name = f"{'_' if is_ignored else ''}Check{group_i}"
                        group_opts = check_matrix_groups.setdefault(base_name, {})
                        if not is_unchecked_button_value(val) or group_opts.get(opt_n, '') == '':
                            group_opts[opt_n] = val
                    else:
                        text_fields[field_name] = val
                else:
                    text_fields[field_name] = val

            for f_name, f_val in sorted(text_fields.items(), key=lambda it: natural_sort_key(it[0])):
                system_key = f"{survey_type}_{f_name}"
                display_name = field_mapping.get(system_key, field_mapping.get(f_name, f_name))
                rows.append((participant_id, file_label, f_name, display_name, f_val))

            for base_name, options in sorted(check_matrix_groups.items(), key=lambda it: natural_sort_key(it[0])):
                system_key = f"{survey_type}_{base_name}"
                display_name = field_mapping.get(system_key, field_mapping.get(base_name, base_name))
                checked_opt = next((
                    opt for opt, cval in sorted(options.items(), key=lambda it: natural_sort_key(it[0]))
                    if not is_unchecked_button_value(cval)
                ), '')
                rows.append((participant_id, file_label, base_name, display_name, checked_opt))

    table.setRowCount(len(rows))
    original_displayed: Dict[Tuple[str, str], str] = {}
    for idx, (pid, file_label, f_name, display_name, value) in enumerate(rows):
        survey_type = strip_pid_from_stem(os.path.splitext(file_label)[0], pid)
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
        prefill_values: Dict[str, Dict[str, Any]] = {
            f"{pid_key}::{f_label}": {}
            for pid_key, meta_list in all_generated_pdf_meta.items()
            for f_label, _ in meta_list
        }

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

            survey_type = strip_pid_from_stem(os.path.splitext(file_name)[0], pid)
            new_name = name_item.text().strip()
            if new_name and not is_ignored_field_name(underlying_field_name):
                sys_map_key = f"{survey_type}_{underlying_field_name}"
                if new_name != original_displayed.get((survey_type, underlying_field_name)) or sys_map_key not in field_mapping:
                    field_mapping[sys_map_key] = new_name

            raw_fields = file_to_extracted_fields.get(key, {})
            matched_matrix_fields: List[Tuple[str, str]] = []
            clean_underlying = underlying_field_name.lstrip('_')
            if clean_underlying.startswith('Check') and ('_' not in clean_underlying and '.' not in clean_underlying):
                for raw_f in raw_fields:
                    m = match_canonical_field(raw_f)
                    if m and m.group(4) is not None:
                        is_raw_ignored = bool(m.group(1)) or is_ignored_field_name(raw_f)
                        f_base = f"{'_' if is_raw_ignored else ''}Check{m.group(3)}"
                        if f_base == underlying_field_name:
                            matched_matrix_fields.append((raw_f, m.group(4)))

            if matched_matrix_fields:
                selected_opts = {s.strip() for s in val.split(',') if s.strip()} if val else set()
                for raw_f, col_opt in matched_matrix_fields:
                    prefill_values[key][raw_f] = '1' if col_opt in selected_opts else '0'
            else:
                prefill_values[key][underlying_field_name] = val

        if active_config_name:
            save_mapping_to_config(active_config_name, field_mapping)

        prefill_values_result.update(prefill_values)
        confirmed_flag[0] = True
        dialog.accept()

    confirm_btn.clicked.connect(confirm_and_close)
    cancel_btn.clicked.connect(dialog.reject)
    dialog.exec()

    return (True, prefill_values_result) if confirmed_flag[0] else (False, {})


def show_batch_extraction_preview_dialog(
    parent: Optional[QWidget],
    all_sheet_data: Dict[str, Dict[str, Dict[str, Any]]],
    root_order: List[str],
    active_config_name: str,
    field_mapping: Dict[str, str]
) -> Tuple[bool, Dict[str, Dict[str, Dict[str, Any]]]]:
    """Displays extracted participants in one unified editable preview dialog before writing."""
    dialog = QDialog(parent)
    dialog.setWindowTitle("Preview Masterfile Extraction")
    dialog.setModal(True)
    dialog.resize(920, 650)

    layout = QVBoxLayout()
    label = QLabel(f"Preview & Verify Extraction Data for {len(all_sheet_data)} Participant(s)")
    label.setStyleSheet("font-weight: bold; font-size: 14px;")
    layout.addWidget(label)

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
                mapped_name = resolve_mapped_column_name(sheet_name, k, field_mapping)
                flat_rows.append((pid, sheet_name, k, mapped_name, str(data[k])))

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
        for col in range(4):
            header.setSectionResizeMode(col, QHeaderView.ResizeMode.ResizeToContents)
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
    edited_output: Dict[str, Dict[str, Dict[str, Any]]] = {
        pid: {s: {'participant_id': pid} for s in s_map}
        for pid, s_map in all_sheet_data.items()
    }

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
                if new_n != original_displayed.get((s, k)) or scoped_key not in field_mapping:
                    field_mapping[scoped_key] = new_n

        if active_config_name:
            save_mapping_to_config(active_config_name, field_mapping)

        confirmed_flag[0] = True
        dialog.accept()

    c_btn.clicked.connect(save_and_close)
    can_btn.clicked.connect(dialog.reject)
    dialog.exec()
    return confirmed_flag[0], edited_output if confirmed_flag[0] else all_sheet_data