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
    except Exception:
        return {}

    extracted: Dict[str, str] = {}
    explicit_on_buttons: Set[str] = set()
    explicit_off_buttons: Set[str] = set()
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

    def resolve_opt_value(raw: Any, opt_list: Optional[List[Any]], allow_zero_index: bool = False) -> Any:
        if opt_list and raw is not None:
            raw_str = decode_pdf_scalar(raw)
            if raw_str.startswith('/'):
                idx_str = raw_str[1:].strip()
                if idx_str.isdigit():
                    idx = int(idx_str)
                    if idx == 0 and not allow_zero_index:
                        return raw
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
        text = decode_pdf_scalar(resolved)
        if text.startswith('/'):
            unslashed = text[1:].strip()
            if is_button and unslashed.lower() in UNCHECKED_BUTTON_VALUES:
                return ''
            if unslashed.lower() in ('yes', 'on', 'true', 'ja', 'ein', 'wahr'):
                return '1'
            if unslashed.lower() in ('off', 'false'):
                return ''
            return unslashed
        if is_button and text.lower() in UNCHECKED_BUTTON_VALUES:
            return ''
        if is_button and text.lower() in ('yes', 'on', 'true', 'ja', 'ein', 'wahr'):
            return '1'
        return text

    def store_field(
        name: Any,
        value: Any,
        is_button: bool = False,
        opt_list: Optional[List[Any]] = None,
        allow_zero_opt: bool = False,
        from_as: bool = False
    ) -> None:
        if not name:
            return
        key = decode_pdf_scalar(name)
        if not key or key == 'File':
            return

        if is_button and is_unchecked_button_value(value):
            if from_as or need_appearances:
                explicit_off_buttons.add(key)
                if not from_as:
                    explicit_on_buttons.discard(key)
            if key not in extracted:
                extracted[key] = ''
            elif (from_as and key not in explicit_on_buttons) or (not from_as and need_appearances):
                extracted[key] = ''
            return

        resolved_val = resolve_opt_value(value, opt_list, allow_zero_index=allow_zero_opt) if is_button else value
        val_str = clean_value(resolved_val, is_button=is_button)

        if is_button and val_str == '':
            if from_as:
                explicit_off_buttons.add(key)
            if key not in extracted:
                extracted[key] = ''
            return

        if is_button and from_as:
            if need_appearances and key in explicit_off_buttons:
                return
            explicit_on_buttons.add(key)

        if is_button and not from_as and not need_appearances and key in explicit_off_buttons and key not in explicit_on_buttons:
            return

        if key not in extracted:
            extracted[key] = val_str
        elif extracted[key] == '' and val_str != '':
            extracted[key] = val_str

    def walk_field_hierarchy(
        obj_ref: Any,
        inherited_name: str = "",
        inherited_ft: str = "",
        inherited_opt: Optional[List[Any]] = None,
        inherited_ff: int = 0
    ) -> None:
        obj = resolve_pdf_object(obj_ref)
        if not isinstance(obj, dict):
            return
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
            return

        local_opt = resolve_pdf_object(obj.get('/Opt'))
        opt_list = local_opt if isinstance(local_opt, list) else inherited_opt
        is_button = (ft == '/Btn') or ('/AS' in obj)

        kids = resolve_pdf_object(obj.get('/Kids'))
        has_kids = isinstance(kids, list) and len(kids) > 0
        allow_zero_opt = isinstance(opt_list, list) and len(opt_list) > 1 and has_kids

        if isinstance(kids, list):
            for k_ref in kids:
                walk_field_hierarchy(k_ref, field_name, ft, opt_list, ff)

        raw_as = resolve_pdf_object(obj.get('/AS'))
        raw_v = resolve_pdf_object(obj.get('/V'))

        if need_appearances and raw_v is not None:
            store_field(field_name, raw_v, is_button=is_button, opt_list=opt_list, allow_zero_opt=allow_zero_opt, from_as=False)
        elif raw_as is not None:
            store_field(field_name, raw_as, is_button=True, opt_list=opt_list, allow_zero_opt=allow_zero_opt, from_as=True)
        elif raw_v is not None:
            store_field(field_name, raw_v, is_button=is_button, opt_list=opt_list, allow_zero_opt=allow_zero_opt, from_as=False)
        elif field_name and field_name not in extracted:
            extracted[field_name] = ''

    def get_annot_full_name_and_meta(annot: Dict[str, Any]) -> Tuple[str, str, Optional[List[Any]], int]:
        names: List[str] = []
        ft_found = ''
        opt_found: Optional[List[Any]] = None
        ff_found = 0
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

            curr = resolve_pdf_object(curr.get('/Parent'))
            depth += 1

        names.reverse()
        return '.'.join(names), ft_found, opt_found, ff_found

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
                    if decode_pdf_scalar(resolve_pdf_object(annot.get('/Subtype', ''))) != '/Widget':
                        continue
                    full_name, ft, opt_list, ff = get_annot_full_name_and_meta(annot)
                    if not full_name or (ft == '/Btn' and (ff & 65536)):
                        continue
                    is_btn = (ft == '/Btn') or ('/AS' in annot)
                    allow_zero_opt = isinstance(opt_list, list) and len(opt_list) > 1

                    val_as = resolve_pdf_object(annot.get('/AS'))
                    val_v = resolve_pdf_object(annot.get('/V'))
                    if need_appearances and val_v is not None:
                        store_field(full_name, val_v, is_button=is_btn, opt_list=opt_list, allow_zero_opt=allow_zero_opt, from_as=False)
                    elif val_as is not None:
                        store_field(full_name, val_as, is_button=True, opt_list=opt_list, allow_zero_opt=allow_zero_opt, from_as=True)
                    elif val_v is not None:
                        store_field(full_name, val_v, is_button=is_btn, opt_list=opt_list, allow_zero_opt=allow_zero_opt, from_as=False)
                    elif full_name not in extracted:
                        extracted[full_name] = ''
                except Exception:
                    continue
    except Exception:
        pass

    if not extracted:
        try:
            acro_fields = reader.get_fields() or {}
            for fname, fmeta in acro_fields.items():
                if isinstance(fmeta, dict):
                    ft = decode_pdf_scalar(resolve_pdf_object(fmeta.get('/FT', '')))
                    opt_raw = resolve_pdf_object(fmeta.get('/Opt'))
                    opt_list = opt_raw if isinstance(opt_raw, list) else None
                    is_btn = ft == '/Btn'
                    v = resolve_pdf_object(fmeta.get('/V'))
                    if v is not None:
                        store_field(fname, v, is_button=is_btn, opt_list=opt_list)
                    else:
                        extracted.setdefault(str(fname).strip(), '')
        except Exception:
            pass

    return extracted


def extract_condition_from_fields(fields: Dict[str, str]) -> str:
    """Extract the condition declaration ('con' / '_con' / 'con_<opt>') from a PDF's fields."""
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
    participant_aliases = {"participant_id", "participant-id", "Text13", "_Text13", "Text1", "_Text1"}
    for field_name in fields.keys():
        clean_name = field_name.split('.')[-1].strip()
        if (clean_name in participant_aliases or "participant" in clean_name.lower()) and str(fields.get(field_name, "")).strip() == "":
            defaults[field_name] = participant_id
    return defaults


def inspect_pdf_raw_fields(pdf_path: str) -> List[Tuple[str, str, str, str, str, str]]:
    """Return raw diagnostic tuples (name, /FT, /V, /AS, /AP /N keys, extracted) for a PDF."""
    reader = PdfReader(pdf_path)
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