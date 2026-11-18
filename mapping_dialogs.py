import os
from typing import Any, Dict, List, Optional, Set, Tuple

from PyQt5.QtWidgets import (
    QButtonGroup, QDialog, QFrame, QHBoxLayout, QHeaderView,
    QLabel, QMessageBox, QPushButton, QRadioButton, QTableWidget,
    QTableWidgetItem, QVBoxLayout, QWidget
)

from config_utils import (
    load_mapping_from_config,
    natural_sort_key,
    save_condition_ordering_to_config,
    save_mapping_to_config,
    set_item_read_only,
)
from data_extraction import scan_source_for_conditions


def edit_field_mapping_dialog(
    parent: Optional[QWidget],
    expected: Set[str],
    root_order: List[str],
    active_config_name: str,
    field_mapping: Dict[str, str]
) -> None:
    def _mapping_sort_key(field_str: str) -> Tuple[int, List[Any]]:
        matched_idx = 9999
        for s_name in sorted(root_order, key=len, reverse=True):
            if field_str.lower().startswith(s_name.lower()):
                matched_idx = root_order.index(s_name)
                break
        return (matched_idx, natural_sort_key(field_str))

    sorted_expected = sorted(expected, key=_mapping_sort_key)
    stored = load_mapping_from_config(active_config_name) if active_config_name else {}
    mapping = {f: stored.get(f, field_mapping.get(f, f)) for f in sorted_expected}

    dialog = QDialog(parent)
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
        field_mapping.update(updated)
        if active_config_name:
            save_mapping_to_config(active_config_name, field_mapping)
            QMessageBox.information(parent, "Success", f"Field mapping saved to active configuration '{active_config_name}'.")
        else:
            QMessageBox.information(parent, "Success", "Field mapping saved in memory. Click 'Save Configuration' to create a permanent snapshot.")
        dialog.accept()

    btn.clicked.connect(save)
    lay.addWidget(btn)
    dialog.setLayout(lay)
    dialog.exec()


def edit_condition_ordering_dialog(
    parent: Optional[QWidget],
    source_path: str,
    configured_names: List[str],
    active_config_name: str,
    condition_ordering: Dict[str, Any],
    prepopulate_items: Optional[List[Tuple[str, str]]] = None,
    prompt_reason: Optional[str] = None
) -> Tuple[bool, Dict[str, Any]]:
    """Open popup dialog to configure numerical vs custom condition ordering and meanings."""
    dialog = QDialog(parent)
    dialog.setWindowTitle("Save/Edit Condition Ordering")
    dialog.setModal(True)
    dialog.resize(780, 520)

    layout = QVBoxLayout()

    if prompt_reason:
        warn_lbl = QLabel(prompt_reason)
        warn_lbl.setWordWrap(True)
        warn_lbl.setStyleSheet("font-weight: bold; color: #b35900; margin-bottom: 4px;")
        layout.addWidget(warn_lbl)

    mode_frame = QFrame()
    mode_frame.setFrameShape(QFrame.Shape.Box)
    mode_layout = QVBoxLayout()
    mode_title = QLabel("Condition Ordering Mode (from PDF field 'con'):")
    mode_title.setStyleSheet("font-weight: bold;")
    mode_layout.addWidget(mode_title)

    rb_numerical = QRadioButton("Numerical (Standard — numeric 'con' values 1, 2, 3... directly set column order)")
    rb_custom = QRadioButton("Custom Condition Ordering (use defined table order & meanings for codes)")
    mode_group = QButtonGroup(dialog)
    mode_group.addButton(rb_numerical)
    mode_group.addButton(rb_custom)

    current_mode = str(condition_ordering.get("mode", "numerical")).lower()
    rb_custom.setChecked(current_mode == "custom")
    rb_numerical.setChecked(current_mode != "custom")

    mode_layout.addWidget(rb_numerical)
    mode_layout.addWidget(rb_custom)
    mode_frame.setLayout(mode_layout)
    layout.addWidget(mode_frame)

    info_lbl = QLabel(
        "Define non-numeric condition codes (e.g., 'sf', 'bl', 'hf') or custom overrides below.\n"
        "Codes listed here are remembered in the project config and work in both modes:"
    )
    info_lbl.setWordWrap(True)
    layout.addWidget(info_lbl)

    btn_bar = QHBoxLayout()
    add_btn = QPushButton("Add Row")
    remove_btn = QPushButton("Remove Selected")
    scan_btn = QPushButton("Scan Source PDFs for 'con'")
    btn_bar.addWidget(add_btn)
    btn_bar.addWidget(remove_btn)
    btn_bar.addWidget(scan_btn)
    btn_bar.addStretch()
    layout.addLayout(btn_bar)

    existing_mappings: List[Dict[str, Any]] = [
        dict(m) for m in condition_ordering.get("mappings", []) if isinstance(m, dict)
    ]

    if prepopulate_items:
        used_orders_by_survey: Dict[str, Set[int]] = {}
        existing_codes_by_survey: Dict[str, Set[str]] = {}
        for m in existing_mappings:
            sv = str(m.get("survey", "*")).strip().lower() or "*"
            try:
                used_orders_by_survey.setdefault(sv, set()).add(int(m.get("order", 0)))
            except (ValueError, TypeError):
                pass
            for c in str(m.get("code", "")).split(","):
                if c.strip():
                    existing_codes_by_survey.setdefault(sv, set()).add(c.strip().lower())

        for sv_raw, code_raw in prepopulate_items:
            sv_key = sv_raw.strip().lower() or "*"
            code_l = code_raw.strip().lower()
            if code_l in existing_codes_by_survey.get(sv_key, set()) or code_l in existing_codes_by_survey.get("*", set()):
                continue
            used = used_orders_by_survey.setdefault(sv_key, set())
            next_order = int(code_raw) if code_raw.strip().isdigit() and int(code_raw) >= 1 else 1
            while next_order in used:
                next_order += 1
            used.add(next_order)
            existing_codes_by_survey.setdefault(sv_key, set()).add(code_l)
            existing_mappings.append({
                "survey": sv_raw.strip() or "*",
                "code": code_raw.strip(),
                "order": next_order,
                "meaning": ""
            })

    table = QTableWidget(len(existing_mappings), 4)
    table.setHorizontalHeaderLabels([
        "Survey (* = All)", "Condition Code ('con')", "Order (1, 2, 3...)", "Meaning / Description"
    ])
    header = table.horizontalHeader()
    if header is not None:
        header.setSectionResizeMode(0, QHeaderView.ResizeMode.ResizeToContents)
        header.setSectionResizeMode(1, QHeaderView.ResizeMode.ResizeToContents)
        header.setSectionResizeMode(2, QHeaderView.ResizeMode.ResizeToContents)
        header.setStretchLastSection(True)

    def populate_row(r_idx: int, survey: str, code: str, order: str, meaning: str) -> None:
        table.setItem(r_idx, 0, QTableWidgetItem(survey))
        table.setItem(r_idx, 1, QTableWidgetItem(code))
        table.setItem(r_idx, 2, QTableWidgetItem(order))
        table.setItem(r_idx, 3, QTableWidgetItem(meaning))

    for idx, item in enumerate(existing_mappings):
        populate_row(
            idx,
            str(item.get("survey", "*")),
            str(item.get("code", "")),
            str(item.get("order", "")),
            str(item.get("meaning", ""))
        )

    layout.addWidget(table)

    def add_row() -> None:
        r_idx = table.rowCount()
        table.insertRow(r_idx)
        populate_row(r_idx, "*", "", str(r_idx + 1), "")

    def remove_selected() -> None:
        selected_rows = sorted({idx.row() for idx in table.selectedIndexes()}, reverse=True)
        for r_idx in selected_rows:
            table.removeRow(r_idx)

    def scan_source_pdfs() -> None:
        if not source_path or not os.path.isdir(source_path):
            QMessageBox.warning(dialog, "ERROR", "Please select a valid Source folder first! (Expand 'Configuration Settings')")
            return
        discovered = scan_source_for_conditions(source_path, configured_names)
        if not discovered:
            QMessageBox.information(dialog, "Scan Complete", "No 'con' fields with values were found in the source folder.")
            return

        existing_pairs: Set[Tuple[str, str]] = set()
        used_orders: Dict[str, Set[int]] = {}
        for r_idx in range(table.rowCount()):
            sv_item = table.item(r_idx, 0)
            cd_item = table.item(r_idx, 1)
            or_item = table.item(r_idx, 2)
            sv_val = sv_item.text().strip().lower() if sv_item else "*"
            cd_val = cd_item.text().strip().lower() if cd_item else ""
            for c in cd_val.split(","):
                if c.strip():
                    existing_pairs.add((sv_val, c.strip()))
            if or_item and or_item.text().strip().isdigit():
                used_orders.setdefault(sv_val, set()).add(int(or_item.text().strip()))

        added_count = 0
        for sv_raw, cd_raw in discovered:
            sv_l = sv_raw.lower()
            cd_l = cd_raw.lower()
            if (sv_l, cd_l) in existing_pairs or ("*", cd_l) in existing_pairs:
                continue
            used = used_orders.setdefault(sv_l, set())
            ord_num = int(cd_raw) if cd_raw.isdigit() and int(cd_raw) >= 1 else 1
            while ord_num in used:
                ord_num += 1
            used.add(ord_num)
            existing_pairs.add((sv_l, cd_l))

            r_idx = table.rowCount()
            table.insertRow(r_idx)
            populate_row(r_idx, sv_raw, cd_raw, str(ord_num), "")
            added_count += 1

        QMessageBox.information(
            dialog, "Scan Complete",
            f"Found {len(discovered)} unique condition value(s); added {added_count} new row(s)."
        )

    add_btn.clicked.connect(add_row)
    remove_btn.clicked.connect(remove_selected)
    scan_btn.clicked.connect(scan_source_pdfs)

    action_layout = QHBoxLayout()
    save_btn = QPushButton("Save Condition Ordering")
    save_btn.setStyleSheet("font-weight: bold;")
    cancel_btn = QPushButton("Cancel")
    action_layout.addWidget(save_btn)
    action_layout.addWidget(cancel_btn)
    layout.addLayout(action_layout)
    dialog.setLayout(layout)

    result: List[Any] = [False, condition_ordering]

    def save_and_close() -> None:
        new_mode = "custom" if rb_custom.isChecked() else "numerical"
        new_mappings: List[Dict[str, Any]] = []
        seen_codes: Set[Tuple[str, str]] = set()
        seen_orders: Set[Tuple[str, int]] = set()

        for r_idx in range(table.rowCount()):
            sv_item = table.item(r_idx, 0)
            cd_item = table.item(r_idx, 1)
            or_item = table.item(r_idx, 2)
            mn_item = table.item(r_idx, 3)

            survey = (sv_item.text().strip() if sv_item else "") or "*"
            code = cd_item.text().strip() if cd_item else ""
            order_txt = or_item.text().strip() if or_item else ""
            meaning = mn_item.text().strip() if mn_item else ""

            if not code and not order_txt and not meaning:
                continue
            if not code:
                QMessageBox.warning(dialog, "ERROR", f"Row {r_idx + 1}: Condition Code ('con') cannot be empty.")
                return
            if not order_txt.isdigit() or int(order_txt) < 1:
                QMessageBox.warning(dialog, "ERROR", f"Row {r_idx + 1}: Order must be a positive integer (1, 2, 3...).")
                return

            order_val = int(order_txt)
            sv_l = survey.lower()
            if (sv_l, order_val) in seen_orders:
                QMessageBox.warning(
                    dialog, "ERROR",
                    f"Duplicate order '{order_val}' for survey '{survey}'.\n"
                    "Use comma-separated codes in one row if you want aliases for the same slot."
                )
                return
            seen_orders.add((sv_l, order_val))

            for single_c in [c.strip().lower() for c in code.split(",") if c.strip()]:
                if (sv_l, single_c) in seen_codes:
                    QMessageBox.warning(dialog, "ERROR", f"Duplicate condition code '{single_c}' for survey '{survey}'.")
                    return
                seen_codes.add((sv_l, single_c))

            new_mappings.append({
                "survey": survey,
                "code": code,
                "order": order_val,
                "meaning": meaning
            })

        new_mappings.sort(key=lambda it: (str(it["survey"]).lower(), int(it["order"])))
        updated_ordering: Dict[str, Any] = {"mode": new_mode, "mappings": new_mappings}

        if active_config_name:
            save_condition_ordering_to_config(active_config_name, updated_ordering)
            if not prompt_reason:
                QMessageBox.information(
                    parent, "Success",
                    f"Condition ordering saved to active configuration '{active_config_name}'."
                )
        elif not prompt_reason:
            QMessageBox.information(
                parent, "Success",
                "Condition ordering saved in memory. Click 'Save Configuration' to create a permanent snapshot."
            )

        result[0] = True
        result[1] = updated_ordering
        dialog.accept()

    save_btn.clicked.connect(save_and_close)
    cancel_btn.clicked.connect(dialog.reject)
    dialog.exec()
    return bool(result[0]), dict(result[1])