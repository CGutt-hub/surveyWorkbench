import os
import io
import tempfile
from typing import Any, Dict, List, Optional, Set, Tuple

from pypdf import PdfReader, PdfWriter
from pypdf.generic import ArrayObject, NameObject, create_string_object

from config_utils import (
    UNCHECKED_BUTTON_VALUES,
    decode_pdf_scalar,
    is_unchecked_button_value,
    match_con_field,
    natural_sort_key,
    resolve_pdf_object,
)


def _get_writer_root_object(writer: PdfWriter) -> Any:
    root_candidate = getattr(writer, 'root_object', None)
    if root_candidate is None:
        root_candidate = getattr(writer, '_root_object', None)
    if root_candidate is None and hasattr(writer, 'trailer'):
        trailer_obj = getattr(writer, 'trailer')
        if isinstance(trailer_obj, dict):
            root_candidate = trailer_obj.get('/Root')
    return resolve_pdf_object(root_candidate)


def write_pdf_form_fields(pdf_path: str, field_values: Dict[str, Any]) -> None:
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
        local_name = decode_pdf_scalar(local_t)
        if inherited_name and local_name:
            field_name = f"{inherited_name}.{local_name}"
        else:
            field_name = local_name or inherited_name

        local_ft = resolve_pdf_object(obj.get('/FT'))
        ft = decode_pdf_scalar(local_ft) or inherited_ft

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
                    on_state = next((k for k in valid_states if not is_unchecked_button_value(k)), '/Yes')
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
        root_obj = _get_writer_root_object(writer)
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


def extract_pdf_form_fields(pdf_path: str) -> Dict[str, str]:
    """Deep extract form fields, radio groups, /Opt values, and widget /AS states."""
    try:
        reader = PdfReader(pdf_path)
        if getattr(reader, "is_encrypted", False):
            try:
                reader.decrypt("")
            except Exception:
                return {}
    except Exception:
        return {}

    need_appearances = False
    try:
        root = resolve_pdf_object(reader.trailer.get('/Root', {}))
        if isinstance(root, dict):
            acro = resolve_pdf_object(root.get('/AcroForm'))
            if isinstance(acro, dict):
                na_val = resolve_pdf_object(acro.get('/NeedAppearances'))
                need_appearances = bool(na_val)
    except Exception:
        pass

    field_records: Dict[str, Dict[str, Any]] = {}
    visited_obj_keys: Set[Any] = set()
    kept_refs: List[Any] = []

    def _mark_visited(obj_ref: Any, resolved_obj: Any) -> bool:
        iref = getattr(obj_ref, 'indirect_reference', None) or getattr(resolved_obj, 'indirect_reference', None)
        if iref is not None:
            marker: Any = ('iref', iref.idnum, iref.generation)
        else:
            marker = ('id', id(resolved_obj))
            kept_refs.append(resolved_obj)
        if marker in visited_obj_keys:
            return False
        visited_obj_keys.add(marker)
        return True

    def _get_ap_states(obj: Dict[str, Any]) -> Set[str]:
        ap_dict = resolve_pdf_object(obj.get('/AP'))
        if not isinstance(ap_dict, dict):
            return set()
        ap_n = resolve_pdf_object(ap_dict.get('/N'))
        if not isinstance(ap_n, dict):
            return set()
        return {str(k).lstrip('/').strip().lower() for k in ap_n.keys()}

    def _ensure_record(key: str, is_button: bool) -> Dict[str, Any]:
        rec = field_records.get(key)
        if rec is None:
            rec = {
                "is_button": is_button,
                "v_entries": [],
                "as_entries": [],
                "ap_states": set(),
            }
            field_records[key] = rec
        elif is_button:
            rec["is_button"] = True
        return rec

    def _record_node(
        key: str,
        is_button: bool,
        raw_v: Any,
        raw_as: Any,
        opt_list: Optional[List[Any]],
        ap_states: Set[str]
    ) -> None:
        if not key or key == 'File':
            return
        rec = _ensure_record(key, is_button)
        rec["ap_states"].update(ap_states)
        if raw_v is not None:
            rec["v_entries"].append((raw_v, opt_list, ap_states))
        if raw_as is not None:
            rec["as_entries"].append((raw_as, opt_list, ap_states))

    def _resolve_opt_index(raw: Any, opt_list: Optional[List[Any]]) -> Tuple[bool, Any]:
        if opt_list and raw is not None:
            raw_str = decode_pdf_scalar(raw)
            if raw_str.startswith('/'):
                idx_str = raw_str[1:].strip()
                if idx_str.isdigit():
                    idx = int(idx_str)
                    if 0 <= idx < len(opt_list):
                        opt_item = resolve_pdf_object(opt_list[idx])
                        if isinstance(opt_item, list) and opt_item:
                            return True, resolve_pdf_object(opt_item[0])
                        return True, opt_item
        return False, raw

    def _parse_button_val(
        raw: Any,
        opt_list: Optional[List[Any]],
        ap_states: Set[str]
    ) -> Tuple[bool, bool, str]:
        resolved = resolve_pdf_object(raw)
        if resolved is None:
            return False, False, ''
        if isinstance(resolved, ArrayObject):
            on_parts: List[str] = []
            saw_off = False
            for item in resolved:
                if item is None:
                    continue
                is_on, is_off, txt = _parse_button_val(item, opt_list, ap_states)
                if is_on and txt:
                    on_parts.append(txt)
                elif is_off:
                    saw_off = True
            if on_parts:
                return True, False, ', '.join(on_parts)
            return False, saw_off, ''

        via_opt, maybe_opt_val = _resolve_opt_index(resolved, opt_list)
        if via_opt:
            opt_txt = decode_pdf_scalar(maybe_opt_val)
            if opt_txt.startswith('/'):
                opt_txt = opt_txt[1:].strip()
            if opt_txt and opt_txt.lower() != 'off':
                return True, False, opt_txt
            return False, True, ''

        text = decode_pdf_scalar(resolved)
        unslashed = text[1:].strip() if text.startswith('/') else text
        ul = unslashed.lower()

        if ul in ('', 'off', 'none', 'null'):
            return False, True, ''

        if ap_states and ul in ap_states and ul != 'off':
            if ul in ('yes', 'on', 'true', 'ja', 'ein', 'wahr'):
                return True, False, '1'
            return True, False, unslashed

        if ul in UNCHECKED_BUTTON_VALUES:
            return False, True, ''

        if ul in ('yes', 'on', 'true', 'ja', 'ein', 'wahr'):
            return True, False, '1'

        return True, False, unslashed

    def _clean_text_val(raw: Any) -> str:
        resolved = resolve_pdf_object(raw)
        if resolved is None:
            return ''
        if isinstance(resolved, ArrayObject):
            parts = [_clean_text_val(x) for x in resolved if x is not None]
            return ', '.join(p for p in parts if p != '')
        text = decode_pdf_scalar(resolved)
        if text.startswith('/'):
            return text[1:].strip()
        return text

    def walk_field_hierarchy(
        obj_ref: Any,
        inherited_name: str = "",
        inherited_ft: str = "",
        inherited_opt: Optional[List[Any]] = None,
        inherited_ff: int = 0
    ) -> Set[str]:
        obj = resolve_pdf_object(obj_ref)
        if not isinstance(obj, dict):
            return set()
        if not _mark_visited(obj_ref, obj):
            return set()

        local_t = resolve_pdf_object(obj.get('/T'))
        local_name = decode_pdf_scalar(local_t)
        field_name = f"{inherited_name}.{local_name}" if (inherited_name and local_name) else (local_name or inherited_name)

        local_ft = resolve_pdf_object(obj.get('/FT'))
        ft = decode_pdf_scalar(local_ft) or inherited_ft

        local_ff_raw = resolve_pdf_object(obj.get('/Ff'))
        try:
            ff = int(local_ff_raw) if local_ff_raw is not None else inherited_ff
        except (TypeError, ValueError):
            ff = inherited_ff

        if ft == '/Btn' and (ff & 65536):
            return set()

        local_opt = resolve_pdf_object(obj.get('/Opt'))
        opt_list = local_opt if isinstance(local_opt, list) else inherited_opt
        is_button = (ft == '/Btn') or ('/AS' in obj)

        ap_states = _get_ap_states(obj)
        kids = resolve_pdf_object(obj.get('/Kids'))
        if isinstance(kids, list):
            for k_ref in kids:
                ap_states.update(walk_field_hierarchy(k_ref, field_name, ft, opt_list, ff))

        raw_as = resolve_pdf_object(obj.get('/AS'))
        raw_v = resolve_pdf_object(obj.get('/V'))
        _record_node(field_name, is_button, raw_v, raw_as, opt_list, ap_states)
        return ap_states

    def get_annot_full_name_and_meta(
        annot: Dict[str, Any]
    ) -> Tuple[str, str, Optional[List[Any]], int, Any]:
        names: List[str] = []
        ft_found = ''
        opt_found: Optional[List[Any]] = None
        ff_found = 0
        v_found: Any = None
        curr: Any = annot
        visited: Set[int] = set()
        depth = 0
        while isinstance(curr, dict) and depth < 15:
            obj_id = id(curr)
            if obj_id in visited:
                break
            visited.add(obj_id)

            t_val = resolve_pdf_object(curr.get('/T'))
            t_str = decode_pdf_scalar(t_val)
            if t_str:
                names.append(t_str)
            if not ft_found:
                ft_found = decode_pdf_scalar(resolve_pdf_object(curr.get('/FT')))
            if opt_found is None:
                opt_val = resolve_pdf_object(curr.get('/Opt'))
                if isinstance(opt_val, list):
                    opt_found = opt_val
            if ff_found == 0:
                ff_raw = resolve_pdf_object(curr.get('/Ff'))
                if ff_raw is not None:
                    try:
                        ff_found = int(ff_raw)
                    except (TypeError, ValueError):
                        pass
            if v_found is None:
                v_cand = resolve_pdf_object(curr.get('/V'))
                if v_cand is not None:
                    v_found = v_cand

            curr = resolve_pdf_object(curr.get('/Parent'))
            depth += 1

        names.reverse()
        return '.'.join(names), ft_found, opt_found, ff_found, v_found

    try:
        root = resolve_pdf_object(reader.trailer.get('/Root', {}))
        if isinstance(root, dict):
            acro = resolve_pdf_object(root.get('/AcroForm'))
            if isinstance(acro, dict):
                fields_raw = resolve_pdf_object(acro.get('/Fields'))
                if isinstance(fields_raw, list):
                    for f in fields_raw:
                        walk_field_hierarchy(f)
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
                    if not _mark_visited(ref, annot):
                        continue
                    if decode_pdf_scalar(resolve_pdf_object(annot.get('/Subtype', ''))) != '/Widget':
                        continue
                    full_name, ft, opt_list, ff, inherited_v = get_annot_full_name_and_meta(annot)
                    if not full_name or (ft == '/Btn' and (ff & 65536)):
                        continue
                    is_btn = (ft == '/Btn') or ('/AS' in annot)
                    ap_states = _get_ap_states(annot)
                    val_as = resolve_pdf_object(annot.get('/AS'))
                    val_v = resolve_pdf_object(annot.get('/V'))
                    if val_v is None:
                        val_v = inherited_v
                    _record_node(full_name, is_btn, val_v, val_as, opt_list, ap_states)
                except Exception:
                    continue
    except Exception:
        pass

    extracted: Dict[str, str] = {}
    for key, rec in field_records.items():
        if rec["is_button"]:
            combined_ap: Set[str] = rec["ap_states"]
            v_on_vals: List[str] = []
            v_explicit_off = False
            for raw_v, opt_list, local_ap in rec["v_entries"]:
                is_on, is_off, txt = _parse_button_val(raw_v, opt_list, local_ap or combined_ap)
                if is_on and txt:
                    v_on_vals.append(txt)
                elif is_off:
                    v_explicit_off = True

            as_on_vals: List[str] = []
            for raw_as, opt_list, local_ap in rec["as_entries"]:
                is_on, _is_off, txt = _parse_button_val(raw_as, opt_list, local_ap or combined_ap)
                if is_on and txt:
                    as_on_vals.append(txt)

            if need_appearances:
                if v_on_vals:
                    extracted[key] = v_on_vals[0]
                elif as_on_vals and not v_explicit_off:
                    extracted[key] = as_on_vals[0]
                else:
                    extracted[key] = ''
            else:
                if as_on_vals:
                    extracted[key] = as_on_vals[0]
                elif v_on_vals:
                    extracted[key] = v_on_vals[0]
                else:
                    extracted[key] = ''
        else:
            val_found = ''
            for raw_v, _opt_list, _local_ap in rec["v_entries"]:
                cand = _clean_text_val(raw_v)
                if cand != '':
                    val_found = cand
                    break
            extracted[key] = val_found

    try:
        acro_fields = reader.get_fields() or {}
        for fname, fmeta in acro_fields.items():
            key = decode_pdf_scalar(fname)
            if not key or key == 'File' or not isinstance(fmeta, dict):
                continue
            if key in extracted and extracted[key] != '':
                continue
            ft = decode_pdf_scalar(resolve_pdf_object(fmeta.get('/FT', '')))
            opt_raw = resolve_pdf_object(fmeta.get('/Opt'))
            opt_list = opt_raw if isinstance(opt_raw, list) else None
            is_btn = ft == '/Btn'
            v = resolve_pdf_object(fmeta.get('/V'))
            if v is not None:
                if is_btn:
                    is_on, _is_off, txt = _parse_button_val(v, opt_list, set())
                    if is_on and txt:
                        extracted[key] = txt
                    elif key not in extracted:
                        extracted[key] = ''
                else:
                    txt = _clean_text_val(v)
                    if txt != '' or key not in extracted:
                        extracted[key] = txt
            elif key not in extracted:
                extracted[key] = ''
    except Exception:
        pass

    return extracted


def extract_condition_from_fields(fields: Dict[str, str]) -> str:
    """Extract the condition declaration ('con' / '_con' / 'con_n' / 'con_<opt>') from a PDF's fields."""
    for raw_k, raw_v in sorted(fields.items(), key=lambda it: natural_sort_key(it[0])):
        m = match_con_field(raw_k)
        if m and m.group(2) is None:
            val = str(raw_v).strip()
            if val:
                return val

    for raw_k, raw_v in sorted(fields.items(), key=lambda it: natural_sort_key(it[0])):
        m = match_con_field(raw_k)
        if m and m.group(2) is not None and not is_unchecked_button_value(raw_v):
            return m.group(2).strip()
    return ""


def build_prefill_defaults(participant_id: str, fields: Dict[str, Any]) -> Dict[str, Any]:
    defaults: Dict[str, Any] = dict(fields)
    participant_aliases = {
        "participant_id", "participant-id",
        "Text13", "_Text13", "Text13_n",
        "Text1", "_Text1", "Text1_n",
    }
    for field_name in fields.keys():
        clean_name = field_name.split('.')[-1].strip()
        if (clean_name in participant_aliases or "participant" in clean_name.lower()) and str(fields.get(field_name, "")).strip() == "":
            defaults[field_name] = participant_id
    return defaults


def inspect_pdf_raw_fields(pdf_path: str) -> List[Tuple[str, str, str, str, str, str]]:
    """Return raw diagnostic tuples (name, /FT, /V, /AS, /AP /N keys, extracted) for a PDF."""
    reader = PdfReader(pdf_path)
    if getattr(reader, "is_encrypted", False):
        try:
            reader.decrypt("")
        except Exception:
            pass
    extracted = extract_pdf_form_fields(pdf_path)
    rows: List[Tuple[str, str, str, str, str, str]] = []

    def walk_diag(obj_ref: Any, inherited_name: str = "", inherited_ft: str = "") -> None:
        obj = resolve_pdf_object(obj_ref)
        if not isinstance(obj, dict):
            return
        local_name = decode_pdf_scalar(resolve_pdf_object(obj.get('/T')))
        field_name = f"{inherited_name}.{local_name}" if (inherited_name and local_name) else (local_name or inherited_name)
        ft = decode_pdf_scalar(resolve_pdf_object(obj.get('/FT'))) or inherited_ft

        raw_v = resolve_pdf_object(obj.get('/V'))
        raw_as = resolve_pdf_object(obj.get('/AS'))
        ap_dict = resolve_pdf_object(obj.get('/AP'))
        ap_n = resolve_pdf_object(ap_dict.get('/N')) if isinstance(ap_dict, dict) else None
        ap_keys = ", ".join(str(k) for k in ap_n.keys()) if isinstance(ap_n, dict) else ""

        kids = resolve_pdf_object(obj.get('/Kids'))
        if isinstance(kids, list) and kids:
            if raw_v is not None or raw_as is not None:
                name_key = field_name or "(group)"
                rows.append((name_key, ft, decode_pdf_scalar(raw_v), decode_pdf_scalar(raw_as), ap_keys, extracted.get(name_key, "")))
            for k_ref in kids:
                walk_diag(k_ref, field_name, ft)
        else:
            name_key = field_name or "(unnamed)"
            rows.append((name_key, ft, decode_pdf_scalar(raw_v), decode_pdf_scalar(raw_as), ap_keys, extracted.get(name_key, "")))

    root = resolve_pdf_object(reader.trailer.get('/Root', {}))
    if isinstance(root, dict):
        acro = resolve_pdf_object(root.get('/AcroForm'))
        if isinstance(acro, dict):
            fields_raw = resolve_pdf_object(acro.get('/Fields'))
            if isinstance(fields_raw, list):
                for f in fields_raw:
                    walk_diag(f)
    return rows