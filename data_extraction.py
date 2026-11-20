import os
import re
from typing import Any, Dict, List, Optional, Set, Tuple

from config_utils import (
    derive_root_survey_name,
    is_ignored_field_name,
    is_unchecked_button_value,
    match_canonical_field,
    match_con_field,
    natural_sort_key,
    strip_pid_from_stem,
)
from pdf_handler import extract_condition_from_fields, extract_pdf_form_fields


def get_participant_folder(source_path: str, participant_id: str) -> str:
    normalized_source = os.path.normpath(source_path)
    candidate_sub = os.path.join(source_path, participant_id)
    if os.path.isdir(candidate_sub):
        return candidate_sub
    if os.path.basename(normalized_source) == participant_id:
        return source_path
    return candidate_sub


def resolve_condition_slot(root_sheet: str, raw_con: str, condition_ordering: Dict[str, Any]) -> Optional[int]:
    """Resolve raw 'con' value to a 1-based condition slot order, or None if unmapped."""
    con_clean = str(raw_con).strip()
    if not con_clean:
        return None

    con_lower = con_clean.lower()
    sheet_lower = root_sheet.strip().lower()
    mode = str(condition_ordering.get("mode", "numerical")).lower()
    mappings = condition_ordering.get("mappings", [])

    for target_scope in (sheet_lower, "*"):
        for item in mappings:
            if not isinstance(item, dict):
                continue
            item_survey = str(item.get("survey", "*")).strip().lower() or "*"
            if item_survey != target_scope:
                continue
            codes = [c.strip().lower() for c in str(item.get("code", "")).split(",") if c.strip()]
            if con_lower in codes:
                try:
                    order_val = int(item.get("order", 0))
                    if order_val >= 1:
                        return order_val
                except (ValueError, TypeError):
                    pass

    if mode == "numerical" and con_clean.isdigit():
        order_val = int(con_clean)
        if order_val >= 1:
            return order_val

    return None


def scan_source_for_conditions(source_path: str, configured_names: List[str]) -> List[Tuple[str, str]]:
    """Scan PDFs in source_path for unique (root_sheet, raw_con) pairs."""
    if not source_path or not os.path.isdir(source_path):
        return []
    found: List[Tuple[str, str]] = []
    seen: Set[Tuple[str, str]] = set()

    for root_dir, _dirs, files in os.walk(source_path):
        folder_name = os.path.basename(root_dir)
        for pdf_file in sorted(files, key=natural_sort_key):
            if not pdf_file.lower().endswith('.pdf'):
                continue
            stem = os.path.splitext(pdf_file)[0]
            clean_type = strip_pid_from_stem(stem, folder_name)
            root_sheet = derive_root_survey_name(clean_type, configured_names)
            fields = extract_pdf_form_fields(os.path.join(root_dir, pdf_file))
            raw_con = extract_condition_from_fields(fields)
            if raw_con:
                key = (root_sheet.lower(), raw_con.lower())
                if key not in seen:
                    seen.add(key)
                    found.append((root_sheet, raw_con))
    return found


def find_unmapped_conditions(
    source_path: str,
    participant_ids: List[str],
    configured_names: List[str],
    condition_ordering: Dict[str, Any],
    fields_cache: Dict[str, Dict[str, str]]
) -> List[Tuple[str, str]]:
    """Scan selected participants' PDFs and return any (root_sheet, raw_con) lacking a slot order."""
    unmapped: List[Tuple[str, str]] = []
    seen: Set[Tuple[str, str]] = set()

    for pid in participant_ids:
        participant_folder = get_participant_folder(source_path, pid)
        if not os.path.isdir(participant_folder):
            continue
        survey_files = sorted(
            [f for f in os.listdir(participant_folder) if f.lower().endswith('.pdf')],
            key=natural_sort_key
        )
        for pdf_file in survey_files:
            pdf_path = os.path.join(participant_folder, pdf_file)
            if pdf_path not in fields_cache:
                fields_cache[pdf_path] = extract_pdf_form_fields(pdf_path)
            fields = fields_cache[pdf_path]
            if not fields:
                continue

            stem = os.path.splitext(pdf_file)[0]
            clean_type = strip_pid_from_stem(stem, pid)
            root_sheet = derive_root_survey_name(clean_type, configured_names)

            raw_con = extract_condition_from_fields(fields)
            if not raw_con:
                continue
            if resolve_condition_slot(root_sheet, raw_con, condition_ordering) is None:
                key = (root_sheet.lower(), raw_con.lower())
                if key not in seen:
                    seen.add(key)
                    unmapped.append((root_sheet, raw_con))

    return unmapped


def prepare_data_for_extraction(
    source_path: str,
    participant_id: str,
    configured_names: List[str],
    configured_copy_counts: Dict[str, int],
    condition_ordering: Dict[str, Any],
    fields_cache: Optional[Dict[str, Dict[str, str]]] = None
) -> Dict[str, Dict[str, Any]]:
    participant_folder = get_participant_folder(source_path, participant_id)
    if not os.path.exists(participant_folder):
        raise ValueError(f"Participant folder not found: {participant_folder}")

    survey_files: List[str] = sorted(
        [f for f in os.listdir(participant_folder) if f.lower().endswith('.pdf')],
        key=natural_sort_key
    )
    if not survey_files:
        raise ValueError(f"No readable questionnaire data found for participant '{participant_id}'.")

    parsed_files: List[Tuple[str, str, str, Dict[str, str], str, Optional[int]]] = []
    root_counts: Dict[str, int] = {}
    seen_slots: Dict[Tuple[str, int], str] = {}

    for pdf_file in survey_files:
        pdf_path = os.path.join(participant_folder, pdf_file)
        if fields_cache is not None and pdf_path in fields_cache:
            fields = fields_cache[pdf_path]
        else:
            fields = extract_pdf_form_fields(pdf_path)
            if fields_cache is not None:
                fields_cache[pdf_path] = fields

        if not fields:
            continue

        stem = os.path.splitext(pdf_file)[0]
        clean_type = strip_pid_from_stem(stem, participant_id)
        root_sheet = derive_root_survey_name(clean_type, configured_names)

        raw_con = extract_condition_from_fields(fields)
        slot = resolve_condition_slot(root_sheet, raw_con, condition_ordering) if raw_con else None

        if slot is not None:
            slot_key = (root_sheet.lower(), slot)
            if slot_key in seen_slots:
                raise ValueError(
                    f"Duplicate condition slot {slot} (con='{raw_con}') for participant '{participant_id}' "
                    f"in survey '{root_sheet}':\n'{seen_slots[slot_key]}' and '{pdf_file}'."
                )
            seen_slots[slot_key] = pdf_file

        parsed_files.append((pdf_file, clean_type, root_sheet, fields, raw_con, slot))
        root_counts[root_sheet] = root_counts.get(root_sheet, 0) + 1

    if not parsed_files:
        raise ValueError(f"No readable PDF form fields found for participant '{participant_id}'.")

    parsed_files.sort(key=lambda item: (item[5] if item[5] is not None else 999999, natural_sort_key(item[0])))

    sheet_data: Dict[str, Dict[str, Any]] = {}

    for _pdf_file, clean_type, root_sheet, fields, raw_con, slot in parsed_files:
        if root_sheet not in sheet_data:
            sheet_data[root_sheet] = {'participant_id': participant_id}

        if slot is not None:
            prefix = f"{root_sheet}_{slot}_"
        else:
            has_multiple_copies = (
                root_counts.get(root_sheet, 1) > 1
                or configured_copy_counts.get(root_sheet.lower(), 1) > 1
                or clean_type.lower() != root_sheet.lower()
            )
            prefix = f"{clean_type}_" if has_multiple_copies else ""

        check_groups: Dict[str, Dict[str, str]] = {}
        has_unignored_con = False

        for raw_k, v in sorted(fields.items(), key=lambda it: natural_sort_key(it[0])):
            if is_ignored_field_name(raw_k):
                continue

            cm = match_con_field(raw_k)
            if cm:
                if not cm.group(1):
                    has_unignored_con = True
                continue

            val = str(v).strip()
            m = match_canonical_field(raw_k)
            if m:
                if bool(m.group(1)):
                    continue
                kind = m.group(2).capitalize()
                group_i = m.group(3)
                opt_n = m.group(4)

                if kind == 'Check':
                    if opt_n is not None:
                        base_name = f"Check{group_i}"
                        group_opts = check_groups.setdefault(base_name, {})
                        existing_v = group_opts.get(opt_n, '')
                        if not is_unchecked_button_value(val) or existing_v == '':
                            group_opts[opt_n] = '' if is_unchecked_button_value(val) else val
                    else:
                        if is_unchecked_button_value(val):
                            final_val = '0'
                        elif val.lower() in ('1', 'yes', 'on', 'true', 'ja', 'ein', 'wahr') or val.lower() == f"check{group_i}":
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
                    target_key = f"{prefix}Text{group_i}_{opt_n}" if opt_n is not None else f"{prefix}Text{group_i}"
                    if val != '' or target_key not in sheet_data[root_sheet]:
                        sheet_data[root_sheet][target_key] = val
            else:
                if raw_k != 'participant_id':
                    target_key = f"{prefix}{raw_k}"
                    if val != '' or target_key not in sheet_data[root_sheet]:
                        sheet_data[root_sheet][target_key] = val

        if has_unignored_con and raw_con:
            sheet_data[root_sheet][f"{prefix}con"] = raw_con

        for base_name, options in sorted(check_groups.items(), key=lambda it: natural_sort_key(it[0])):
            checked_opts = [
                opt for opt, cval in sorted(options.items(), key=lambda it: natural_sort_key(it[0]))
                if not is_unchecked_button_value(cval)
            ]
            checked_val = ', '.join(checked_opts) if checked_opts else ''
            sheet_data[root_sheet][f"{prefix}{base_name}"] = checked_val

    return sheet_data


def resolve_mapped_column_name(sheet_name: str, sys_key: str, field_mapping: Dict[str, str]) -> str:
    """Resolve a system key to its mapped masterfile column without cross-survey collisions."""
    scoped_key = f"{sheet_name}_{sys_key}"
    if scoped_key in field_mapping and field_mapping[scoped_key]:
        return field_mapping[scoped_key]
    if sys_key in field_mapping and field_mapping[sys_key]:
        return field_mapping[sys_key]
    if sys_key.lower().startswith(sheet_name.lower()):
        m = re.match(rf'^({re.escape(sheet_name)}[\s_-]*\d+)_(.+)$', sys_key, re.IGNORECASE)
        if m:
            copy_prefix, bare_field = m.group(1), m.group(2)
            root_scoped = f"{sheet_name}_{bare_field}"
            if root_scoped in field_mapping and field_mapping[root_scoped]:
                return f"{copy_prefix}_{field_mapping[root_scoped]}"
    return sys_key


def build_export_row(
    pid: str,
    sheet_name: str,
    raw_data: Dict[str, Any],
    field_mapping: Dict[str, str]
) -> Tuple[Dict[str, Any], Dict[str, str]]:
    export_data: Dict[str, Any] = {'participant_id': pid}
    system_to_target_header: Dict[str, str] = {'participant_id': 'participant_id'}

    sorted_field_keys = sorted(
        [k for k in raw_data.keys() if k != 'participant_id'],
        key=natural_sort_key
    )
    for sys_key in sorted_field_keys:
        final_col = resolve_mapped_column_name(sheet_name, sys_key, field_mapping)
        export_data[final_col] = raw_data[sys_key]
        system_to_target_header[sys_key] = final_col

    return export_data, system_to_target_header