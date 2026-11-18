import os
import re
import csv
import json
import importlib
import tempfile
from typing import Any, Dict, List, Optional, Set, Tuple

import xlwings as xlw

from config_utils import natural_sort_key
from data_extraction import build_export_row

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


def normalize_excel_pid(cell_val: Any) -> str:
    if cell_val is None:
        return ''
    if isinstance(cell_val, float) and cell_val.is_integer():
        return str(int(cell_val))
    return str(cell_val).strip()


def sanitize_excel_sheet_name(name: str) -> str:
    cleaned = re.sub(r'[\\/?*\[\]:]', '_', name).strip()
    return (cleaned or "Sheet")[:31]


def get_masterfile_siblings(base_dir: str, base_name: str, valid_exts: Tuple[str, ...]) -> List[str]:
    if not os.path.isdir(base_dir):
        return []
    return [
        os.path.join(base_dir, f)
        for f in os.listdir(base_dir)
        for stem, ext in [os.path.splitext(f)]
        if ext.lower() in valid_exts and (stem == base_name or stem.startswith(f"{base_name}_"))
    ]


def get_preassigned_excel_row(pid: str) -> Optional[int]:
    """Return preassigned Excel row (x + 1) using trailing/single participant index x, or None."""
    s = str(pid).strip()
    if not s:
        return None
    m_trailing = re.search(r'(\d+)$', s)
    if m_trailing:
        x = int(m_trailing.group(1))
        return x + 1 if 1 <= x <= 100000 else None
    nums = re.findall(r'\d+', s)
    if len(nums) == 1:
        x = int(nums[0])
        return x + 1 if 1 <= x <= 100000 else None
    return None


def _upsert_row(existing_rows: List[Dict[str, Any]], pid: str, export_data: Dict[str, Any]) -> None:
    for row in existing_rows:
        if str(row.get('participant_id', '')).strip() == str(pid):
            row.update(export_data)
            return
    existing_rows.append(export_data)


def get_existing_masterfile_pids(excel_path: str) -> Set[str]:
    """Detect existing participant IDs across all supported masterfile formats."""
    if not excel_path or not os.path.exists(excel_path):
        return set()

    file_ext = excel_path.lower().split('.')[-1]
    existing_pids: Set[str] = set()
    base_dir = os.path.dirname(excel_path)
    base_name = os.path.splitext(os.path.basename(excel_path))[0]

    if file_ext in ['xls', 'xlsx']:
        wb: Optional[xlw.Book] = None
        try:
            wb = xlw.Book(excel_path)
            for sheet in wb.sheets:
                last_used_row = sheet.used_range.last_cell.row
                if last_used_row >= 2:
                    col_a_vals = sheet.range((2, 1), (last_used_row, 1)).value
                    if not isinstance(col_a_vals, list):
                        col_a_vals = [col_a_vals]
                    for cell_val in col_a_vals:
                        cell_str = normalize_excel_pid(cell_val)
                        if cell_str:
                            existing_pids.add(cell_str)
        except Exception:
            pass
        finally:
            if wb is not None:
                try:
                    wb.close()
                except Exception:
                    pass
        return existing_pids

    if file_ext == 'parquet':
        candidates = get_masterfile_siblings(base_dir, base_name, ('.parquet',)) or [excel_path]
        for c_path in candidates:
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
        candidates = get_masterfile_siblings(base_dir, base_name, (f'.{file_ext}',)) or [excel_path]
        for c_path in candidates:
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
        candidates = get_masterfile_siblings(base_dir, base_name, ('.json',)) or [excel_path]
        for c_path in candidates:
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


def export_excel_batch(
    excel_path: str,
    participant_ids: List[str],
    all_participants_sheet_data: Dict[str, Dict[str, Dict[str, Any]]],
    root_order: List[str],
    field_mapping: Dict[str, str]
) -> None:
    wb: Optional[xlw.Book] = None
    try:
        wb = xlw.Book(excel_path) if os.path.exists(excel_path) else xlw.Book()
        if not os.path.exists(excel_path):
            wb.save(excel_path)

        for pid in participant_ids:
            sheet_data = all_participants_sheet_data.get(pid, {})
            pid_sheet_names = sorted(
                sheet_data.keys(),
                key=lambda s: (root_order.index(s) if s in root_order else 9999, natural_sort_key(s))
            )

            for sheet_name in pid_sheet_names:
                raw_data = sheet_data[sheet_name]
                export_data, system_to_target_header = build_export_row(pid, sheet_name, raw_data, field_mapping)

                excel_sheet_name = sanitize_excel_sheet_name(sheet_name)
                existing_sheet_map: Dict[str, xlw.Sheet] = {str(s.name).strip().lower(): s for s in wb.sheets}
                matched_sheet = existing_sheet_map.get(excel_sheet_name.lower())
                header_to_col: Dict[str, int] = {}
                sheet: xlw.Sheet

                if matched_sheet is not None:
                    sheet = matched_sheet
                    first_row = sheet.range('1:1').value
                    if first_row:
                        row_vals = first_row if isinstance(first_row, list) else [first_row]
                        last_non_empty = max(
                            (idx for idx, v in enumerate(row_vals) if v is not None and str(v).strip()),
                            default=-1
                        )
                        for idx in range(last_non_empty + 1):
                            v = row_vals[idx]
                            if v is not None and str(v).strip():
                                header_to_col[str(v).strip()] = idx + 1

                    if 'participant_id' not in header_to_col:
                        col1_hdr = next((h for h, c in header_to_col.items() if c == 1), None)
                        if col1_hdr and col1_hdr.lower().replace(' ', '').replace('_', '').replace('-', '') == 'participantid':
                            header_to_col['participant_id'] = 1

                    for sys_k, target_h in system_to_target_header.items():
                        if target_h not in header_to_col:
                            if sys_k in header_to_col:
                                col_idx = header_to_col.pop(sys_k)
                                sheet.range((1, col_idx)).value = target_h
                                header_to_col[target_h] = col_idx
                            else:
                                next_col = max(header_to_col.values(), default=0) + 1
                                sheet.range((1, next_col)).value = target_h
                                header_to_col[target_h] = next_col

                    row_pid_map: Dict[int, str] = {}
                    pid_to_row: Dict[str, int] = {}
                    last_non_empty_row = 1
                    last_used_row = sheet.used_range.last_cell.row
                    if last_used_row >= 2:
                        col_a_vals = sheet.range((2, 1), (last_used_row, 1)).value
                        if not isinstance(col_a_vals, list):
                            col_a_vals = [col_a_vals]
                        for offset, cell_val in enumerate(col_a_vals):
                            r_idx = 2 + offset
                            norm_val = normalize_excel_pid(cell_val)
                            if norm_val:
                                row_pid_map[r_idx] = norm_val
                                pid_to_row.setdefault(norm_val, r_idx)
                                last_non_empty_row = r_idx

                    norm_pid = normalize_excel_pid(pid)
                    target_row = pid_to_row.get(norm_pid)
                    if target_row is None and norm_pid.isdigit():
                        target_row = pid_to_row.get(str(int(norm_pid)))

                    if target_row is not None:
                        write_row = target_row
                    else:
                        preassigned = get_preassigned_excel_row(norm_pid)
                        write_row = preassigned if (preassigned is not None and preassigned not in row_pid_map) else (last_non_empty_row + 1)
                else:
                    if len(wb.sheets) == 1 and wb.sheets[0].range('A1').value is None and wb.sheets[0].used_range.value is None:
                        sheet = wb.sheets[0]
                        sheet.name = excel_sheet_name
                    else:
                        sheet = wb.sheets.add(name=excel_sheet_name, after=wb.sheets[-1])

                    for idx, header_name in enumerate(export_data.keys(), start=1):
                        sheet.range((1, idx)).value = header_name
                        header_to_col[header_name] = idx

                    preassigned = get_preassigned_excel_row(normalize_excel_pid(pid))
                    write_row = preassigned if preassigned is not None else 2

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


def export_parquet_batch(
    excel_path: str,
    participant_ids: List[str],
    all_participants_sheet_data: Dict[str, Dict[str, Dict[str, Any]]],
    ordered_sheet_names: List[str],
    field_mapping: Dict[str, str]
) -> None:
    if not HAS_PYARROW and not HAS_PANDAS:
        raise ValueError("Exporting to Parquet format requires 'pyarrow' (or 'pandas').")

    base_dir = os.path.dirname(os.path.abspath(excel_path)) or "."
    base_name = os.path.splitext(os.path.basename(excel_path))[0]

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
            if raw_data:
                export_data, _ = build_export_row(pid, sheet_name, raw_data, field_mapping)
                _upsert_row(existing_rows, pid, export_data)

        if existing_rows:
            all_cols = ['participant_id']
            for r in existing_rows:
                for c in r.keys():
                    if c not in all_cols:
                        all_cols.append(c)
            normalized_rows = [{c: str(r.get(c, '')) if r.get(c) is not None else '' for c in all_cols} for r in existing_rows]

            fd, tmp_path = tempfile.mkstemp(prefix=".pq_", suffix=".tmp", dir=base_dir)
            os.close(fd)
            try:
                if HAS_PYARROW and pa is not None and pq is not None:
                    pq.write_table(pa.Table.from_pylist(normalized_rows), tmp_path)
                elif HAS_PANDAS and pd is not None:
                    pd.DataFrame(normalized_rows).to_parquet(tmp_path, index=False)
                os.replace(tmp_path, target_parquet)
            except Exception:
                if os.path.exists(tmp_path):
                    try:
                        os.remove(tmp_path)
                    except Exception:
                        pass
                raise


def export_delimited_batch(
    excel_path: str,
    participant_ids: List[str],
    all_participants_sheet_data: Dict[str, Dict[str, Dict[str, Any]]],
    ordered_sheet_names: List[str],
    file_ext: str,
    field_mapping: Dict[str, str]
) -> None:
    delim = '\t' if file_ext in ['tsv', 'tab'] else ','
    base_dir = os.path.dirname(os.path.abspath(excel_path)) or "."
    base_name = os.path.splitext(os.path.basename(excel_path))[0]

    for sheet_name in ordered_sheet_names:
        target_file = os.path.join(base_dir, f"{base_name}_{sheet_name}.{file_ext}")
        existing_headers: List[str] = []
        existing_rows: List[Dict[str, Any]] = []
        if os.path.exists(target_file) and os.path.getsize(target_file) > 0:
            with open(target_file, 'r', newline='', encoding='utf-8') as f:
                reader = csv.DictReader(f, delimiter=delim)
                existing_headers = list(reader.fieldnames) if reader.fieldnames else []
                existing_rows = list(reader)

        all_headers: List[str] = list(existing_headers)
        for pid in participant_ids:
            raw_data = all_participants_sheet_data.get(pid, {}).get(sheet_name)
            if not raw_data:
                continue
            export_data, _ = build_export_row(pid, sheet_name, raw_data, field_mapping)
            _upsert_row(existing_rows, pid, export_data)
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


def export_json_batch(
    excel_path: str,
    participant_ids: List[str],
    all_participants_sheet_data: Dict[str, Dict[str, Dict[str, Any]]],
    ordered_sheet_names: List[str],
    field_mapping: Dict[str, str]
) -> None:
    base_dir = os.path.dirname(os.path.abspath(excel_path)) or "."
    base_name = os.path.splitext(os.path.basename(excel_path))[0]

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
            if raw_data:
                export_data, _ = build_export_row(pid, sheet_name, raw_data, field_mapping)
                _upsert_row(existing_rows, pid, export_data)

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


def readout_batch(
    excel_path: str,
    participant_ids: List[str],
    all_participants_sheet_data: Dict[str, Dict[str, Dict[str, Any]]],
    root_order: List[str],
    field_mapping: Dict[str, str]
) -> None:
    """Write all processed participants in a single session across supported formats."""
    if not excel_path:
        raise ValueError("Please select a masterfile! (Expand 'Configuration Settings' to select masterfile)")

    file_ext = excel_path.lower().split('.')[-1]
    all_sheet_names: Set[str] = set()
    for pid in participant_ids:
        all_sheet_names.update(all_participants_sheet_data.get(pid, {}).keys())

    ordered_sheet_names = sorted(
        all_sheet_names,
        key=lambda s: (root_order.index(s) if s in root_order else 9999, natural_sort_key(s))
    )

    if file_ext in ['xls', 'xlsx']:
        export_excel_batch(excel_path, participant_ids, all_participants_sheet_data, root_order, field_mapping)
    elif file_ext == 'parquet':
        export_parquet_batch(excel_path, participant_ids, all_participants_sheet_data, ordered_sheet_names, field_mapping)
    elif file_ext in ['csv', 'tsv', 'tab']:
        export_delimited_batch(excel_path, participant_ids, all_participants_sheet_data, ordered_sheet_names, file_ext, field_mapping)
    elif file_ext == 'json':
        export_json_batch(excel_path, participant_ids, all_participants_sheet_data, ordered_sheet_names, field_mapping)
    else:
        raise ValueError(f"Unsupported masterfile format: {file_ext}")