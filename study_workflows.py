import os
import json
import shutil
from typing import Any, Dict, List, Optional, Set, Tuple

from PyQt5.QtWidgets import QDialog, QMessageBox, QWidget

from config_utils import (
    default_condition_ordering,
    get_config_path,
    is_ignored_field_name,
    load_condition_ordering_from_config,
    load_mapping_from_config,
    match_canonical_field,
    match_con_field,
    natural_sort_key,
    new_config_parser,
    write_config_atomically,
)
from data_extraction import find_unmapped_conditions, prepare_data_for_extraction
from dialogs import BatchParticipantSelectionDialog, OverwriteResolutionDialog, QuestionnaireRow
from exporters import get_existing_masterfile_pids, readout_batch
from mapping_dialogs import edit_condition_ordering_dialog
from pdf_handler import extract_pdf_form_fields, write_pdf_form_fields
from preview_dialogs import show_batch_extraction_preview_dialog, show_batch_generation_prefill_dialog


def get_expected_system_fields(questionnaire_rows: List[QuestionnaireRow]) -> Set[str]:
    expected: Set[str] = set()
    for row in questionnaire_rows:
        template = row.template_path.strip()
        if not template or not template.lower().endswith('.pdf') or not os.path.isfile(template):
            continue
        template_fields = extract_pdf_form_fields(template)
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
            has_unignored_con = False

            for field_name in template_fields.keys():
                if is_ignored_field_name(field_name):
                    continue
                cm = match_con_field(field_name)
                if cm:
                    if not cm.group(1):
                        has_unignored_con = True
                    continue
                m = match_canonical_field(field_name)
                if m:
                    if bool(m.group(1)):
                        continue
                    kind = m.group(2).capitalize()
                    group_i = m.group(3)
                    if kind == 'Check':
                        check_matrix_bases.add(f"Check{group_i}")
                    else:
                        expected.add(f"{prefix}Text{group_i}")
                elif field_name != 'participant_id':
                    expected.add(f"{prefix}{field_name}")

            if has_unignored_con:
                expected.add(f"{prefix}con")
            for base_name in check_matrix_bases:
                expected.add(f"{prefix}{base_name}")

    return expected


def ensure_or_create_field_mapping(
    questionnaire_rows: List[QuestionnaireRow],
    active_config_name: str,
    field_mapping: Dict[str, str]
) -> Dict[str, str]:
    expected = get_expected_system_fields(questionnaire_rows)
    if not expected:
        return field_mapping
    stored = load_mapping_from_config(active_config_name) if active_config_name else {}
    mapping = dict(field_mapping)
    for f in sorted(expected, key=natural_sort_key):
        mapping[f] = stored.get(f, field_mapping.get(f, f))
    field_mapping.clear()
    field_mapping.update(mapping)
    return field_mapping


def run_generate_folders(
    parent: Optional[QWidget],
    participant_ids: List[str],
    target_path: str,
    questionnaire_rows: List[QuestionnaireRow],
    root_order: List[str],
    active_config_name: str,
    field_mapping: Dict[str, str]
) -> None:
    existing_pids = [pid for pid in participant_ids if os.path.exists(os.path.join(target_path, pid))]
    pids_to_process = list(participant_ids)

    if existing_pids:
        dlg = OverwriteResolutionDialog(
            existing_pids=existing_pids,
            title="Participant Folder(s) Already Exist",
            prompt="The following participant folder(s) already exist in the target directory.\n"
                   "Tick the participants you wish to overwrite (unticked will be skipped):",
            parent=parent
        )
        if dlg.exec() != QDialog.DialogCode.Accepted:
            return
        approved = set(dlg.get_selected_pids())
        pids_to_process = [pid for pid in participant_ids if (pid not in existing_pids) or (pid in approved)]

    if not pids_to_process:
        QMessageBox.information(parent, "No Changes", "No participant folders selected to generate.")
        return

    all_generated_pdf_meta: Dict[str, List[Tuple[str, str]]] = {}
    for pid in pids_to_process:
        surveys_meta: List[Tuple[str, str]] = []
        for row in questionnaire_rows:
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
        QMessageBox.warning(parent, "ERROR", "No valid questionnaires found to generate.")
        return

    ensure_or_create_field_mapping(questionnaire_rows, active_config_name, field_mapping)
    apply_prefill, batch_prefill_map = show_batch_generation_prefill_dialog(
        parent, all_generated_pdf_meta, root_order, active_config_name, field_mapping
    )
    if not apply_prefill:
        return

    for pid in pids_to_process:
        surveys_for_pid = all_generated_pdf_meta.get(pid, [])
        if not surveys_for_pid:
            continue
        participant_folder = os.path.join(target_path, pid)
        if os.path.exists(participant_folder):
            shutil.rmtree(participant_folder)
        os.makedirs(participant_folder, exist_ok=True)

        for filename, template_path in surveys_for_pid:
            dest_path = os.path.join(participant_folder, filename)
            shutil.copy2(template_path, dest_path)
            field_vals = batch_prefill_map.get(f"{pid}::{filename}", {})
            if field_vals:
                write_pdf_form_fields(dest_path, field_vals)

    QMessageBox.information(
        parent, "Success",
        f"Successfully generated folders and questionnaires for {len(pids_to_process)} participant(s)!"
    )


def run_extract_data(
    parent: Optional[QWidget],
    participant_ids: List[str],
    source_path: str,
    excel_path: str,
    root_order: List[str],
    configured_copy_counts: Dict[str, int],
    active_config_name: str,
    field_mapping: Dict[str, str],
    condition_ordering: Dict[str, Any]
) -> Dict[str, Any]:
    existing_masterfile_pids = get_existing_masterfile_pids(excel_path)
    conflicting_pids = {
        pid for pid in participant_ids
        if pid in existing_masterfile_pids or (pid.isdigit() and str(int(pid)) in existing_masterfile_pids)
    }

    dlg = BatchParticipantSelectionDialog(all_pids=participant_ids, conflicting_pids=conflicting_pids, parent=parent)
    if dlg.exec() != QDialog.DialogCode.Accepted:
        return condition_ordering

    selected_set = set(dlg.get_selected_pids())
    pids_to_extract = [pid for pid in participant_ids if pid in selected_set]
    if not pids_to_extract:
        QMessageBox.information(parent, "No Changes", "No participants selected to extract.")
        return condition_ordering

    pdf_fields_cache: Dict[str, Dict[str, str]] = {}
    unmapped = find_unmapped_conditions(source_path, pids_to_extract, root_order, condition_ordering, pdf_fields_cache)
    if unmapped:
        codes_str = ", ".join(f"'{c}' ({sv})" for sv, c in unmapped)
        saved_co, updated_co = edit_condition_ordering_dialog(
            parent=parent,
            source_path=source_path,
            configured_names=root_order,
            active_config_name=active_config_name,
            condition_ordering=condition_ordering,
            prepopulate_items=unmapped,
            prompt_reason=(
                f"Unmapped condition value(s) detected in PDFs: {codes_str}.\n"
                "Please assign each condition its column order (1, 2, 3...) and optional meaning before exporting:"
            )
        )
        if not saved_co:
            return condition_ordering
        condition_ordering = updated_co

    all_sheet_data: Dict[str, Dict[str, Dict[str, Any]]] = {}
    for pid in pids_to_extract:
        all_sheet_data[pid] = prepare_data_for_extraction(
            source_path=source_path,
            participant_id=pid,
            configured_names=root_order,
            configured_copy_counts=configured_copy_counts,
            condition_ordering=condition_ordering,
            fields_cache=pdf_fields_cache
        )

    confirmed, edited_data = show_batch_extraction_preview_dialog(
        parent, all_sheet_data, root_order, active_config_name, field_mapping
    )
    if not confirmed:
        return condition_ordering

    readout_batch(excel_path, pids_to_extract, edited_data, root_order, field_mapping)
    QMessageBox.information(
        parent, "Extraction Complete",
        f"Successfully extracted {len(pids_to_extract)} participant(s) into masterfile!"
    )
    return condition_ordering


def save_config_profile(
    name: str,
    target_path: str,
    questionnaire_rows: List[QuestionnaireRow],
    source_path: str,
    excel_path: str,
    field_mapping: Dict[str, str],
    condition_ordering: Dict[str, Any]
) -> None:
    cfg = new_config_parser()
    cfg.read(get_config_path(), encoding='utf-8')
    if cfg.has_section(name):
        cfg.remove_section(name)
    cfg.add_section(name)

    cfg.set(name, 'target_path', target_path)
    cfg.set(name, 'quest_count', str(len(questionnaire_rows)))
    for i, row in enumerate(questionnaire_rows):
        cfg.set(name, f'quest_{i}_name', row.name_edit.text())
        cfg.set(name, f'quest_{i}_path', row.template_path)
        cfg.set(name, f'quest_{i}_count', row.count_edit.text())
    cfg.set(name, 'source_path', source_path)
    cfg.set(name, 'excel_path', excel_path)

    ensure_or_create_field_mapping(questionnaire_rows, name, field_mapping)
    cfg.set(name, 'field_mapping_json', json.dumps(field_mapping, ensure_ascii=False))
    cfg.set(name, 'condition_ordering_json', json.dumps(condition_ordering, ensure_ascii=False))
    write_config_atomically(cfg, get_config_path())


def load_config_profile(name: str) -> Optional[Dict[str, Any]]:
    cfg = new_config_parser()
    cfg.read(get_config_path(), encoding='utf-8')
    if not cfg.has_section(name):
        return None
    qc = cfg.getint(name, 'quest_count', fallback=0)
    rows_meta = [
        (
            cfg.get(name, f'quest_{i}_name', fallback=''),
            cfg.get(name, f'quest_{i}_path', fallback=''),
            cfg.get(name, f'quest_{i}_count', fallback='1')
        )
        for i in range(max(0, qc))
    ]
    return {
        "target_path": cfg.get(name, 'target_path', fallback=''),
        "rows": rows_meta,
        "source_path": cfg.get(name, 'source_path', fallback=''),
        "excel_path": cfg.get(name, 'excel_path', fallback=''),
        "field_mapping": load_mapping_from_config(name),
        "condition_ordering": load_condition_ordering_from_config(name) or default_condition_ordering(),
    }