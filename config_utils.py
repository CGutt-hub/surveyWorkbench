import os
import sys
import re
import json
import shutil
import tempfile
from typing import Optional, List, Dict, Any, Set
from configparser import ConfigParser

from PyQt5.QtWidgets import QTableWidgetItem
from PyQt5.QtCore import Qt

# ==============================================================================
# CANONICAL FORM FIELD GRAMMAR:
#   (_){Check, Text}i(_n)
#
#   (_) / _n suffix : NOT operator. If present (leading '_' or '_n' operator),
#                     the field is strictly ignored and NOT exported.
#   {Check, Text}   : Declares entry type. 'Text' is string; 'Check' is boolean/nominal.
#   i               : Index of the condition/question (stays constant across options n).
#   (_n)            : Option index. When selected, the masterfile cell for column
#                     'Check<i>' displays 'n' instead of the internal PDF value.
#
#   (_)con(_n)      : Optional condition declaration field used to order mixed-condition
#                     surveys (e.g., nasa_1, nasa_2, nasa_3) consistently in the row.
# ==============================================================================
CANONICAL_PATTERN = re.compile(r'^(_)?(Text|Check)(\d+)(?:[._]([a-zA-Z0-9-]+))?$', re.IGNORECASE)
CON_FIELD_PATTERN = re.compile(r'^(_)?con(?:_n)?(?:[._](?!n$)([a-zA-Z0-9-]+?))?(?:_n)?$', re.IGNORECASE)
IGNORED_MID_N_PATTERN = re.compile(
    r'^(?:text|check|con)(?:\d*_n(?:[._-].*)?|_n\d+(?:[._-].*)?)$',
    re.IGNORECASE
)
INVALID_ID_CHARS_PATTERN = re.compile(r'[<>:"/\\|?*\x00-\x1f]')
CONFIG_FILENAME = "survey_workbench.config"
UNCHECKED_BUTTON_VALUES: Set[str] = {
    '', 'off', 'false', '0', '0.0', 'no', 'n', 'unchecked', 'none', 'null',
    'aus', 'nein', 'falsch', 'ungeprüft', 'ungeprueft', 'leer'
}


def new_config_parser() -> ConfigParser:
    """Create a ConfigParser without '%' interpolation so paths/headers with '%' never crash."""
    return ConfigParser(interpolation=None)


def set_item_read_only(item: QTableWidgetItem) -> None:
    """Configure a QTableWidgetItem as selectable and enabled, but read-only."""
    item.setFlags(Qt.ItemFlags(Qt.ItemFlag.ItemIsSelectable) | Qt.ItemFlag.ItemIsEnabled)


def decode_pdf_scalar(val: Any) -> str:
    """Safely convert a resolved PDF scalar (including bytes/UTF-16/ByteStringObject) to str."""
    if val is None:
        return ''
    if isinstance(val, bytes):
        if val.startswith(b'\xfe\xff'):
            text = val[2:].decode('utf-16-be', errors='ignore')
        elif val.startswith(b'\xff\xfe'):
            text = val[2:].decode('utf-16-le', errors='ignore')
        else:
            try:
                text = val.decode('utf-8')
            except UnicodeDecodeError:
                text = val.decode('latin-1', errors='ignore')
        return text.replace('\x00', '').strip()
    return str(val).replace('\x00', '').strip()


def is_ignored_field_name(field_name: str) -> bool:
    """Return True if the field or any hierarchy segment uses the '_' or '_n' NOT operator."""
    cleaned = str(field_name).strip().lstrip('\ufeff')
    if not cleaned:
        return False
    for part in cleaned.split('.'):
        p = part.strip().lstrip('\ufeff')
        if not p:
            continue
        if p.startswith('_') or p.lower().endswith('_n') or IGNORED_MID_N_PATTERN.match(p):
            return True
    return False


def is_unchecked_button_value(val: Any) -> bool:
    """Return True if a button/checkbox value represents an unchecked/off state."""
    if val is None:
        return True
    s = decode_pdf_scalar(val)
    if s.startswith('/'):
        s = s[1:].strip()
    return s.lower() in UNCHECKED_BUTTON_VALUES


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


def match_con_field(field_name: str) -> Optional[re.Match[str]]:
    """Match CON_FIELD_PATTERN against full name first, then dot-separated suffixes."""
    m = CON_FIELD_PATTERN.match(field_name)
    if m:
        return m
    parts = field_name.split('.')
    for idx in range(1, len(parts)):
        candidate = '.'.join(parts[idx:])
        m = CON_FIELD_PATTERN.match(candidate)
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


def strip_pid_from_stem(stem: str, participant_id: str) -> str:
    pid_esc = re.escape(participant_id)
    cleaned = re.sub(rf'^{pid_esc}_', '', stem, count=1, flags=re.IGNORECASE)
    if cleaned == stem:
        cleaned = re.sub(rf'_{pid_esc}$', '', stem, count=1, flags=re.IGNORECASE)
    return cleaned


def derive_root_survey_name(clean_filename: str, configured_names: List[str]) -> str:
    """Derive the root sheet name directly from the given filename, stripping copy indices."""
    clean = clean_filename.strip()
    if not clean:
        return clean
    clean_lower = clean.lower()

    for s_name in sorted(configured_names, key=len, reverse=True):
        s_lower = s_name.lower()
        if clean_lower == s_lower:
            return clean
        if re.match(rf'^{re.escape(s_lower)}(?:[\s_-]*\d+|[\s_-]*copy(?:\s*\(?\d+\)?)?)$', clean_lower):
            return clean[:len(s_name)]

    stripped = re.sub(r'[\s_-]*copy(?:\s*\(?\d+\)?)?$', '', clean, flags=re.IGNORECASE).strip('_- ')
    stripped = re.sub(r'[\s_-]*\d+$', '', stripped).strip('_- ')
    return stripped if stripped else clean


def default_condition_ordering() -> Dict[str, Any]:
    return {"mode": "numerical", "mappings": []}


def load_condition_ordering_from_config(config_name: str) -> Dict[str, Any]:
    if not config_name:
        return default_condition_ordering()
    cfg = new_config_parser()
    cfg.read(get_config_path(), encoding='utf-8')
    if not cfg.has_section(config_name):
        return default_condition_ordering()
    raw = cfg.get(config_name, 'condition_ordering_json', fallback='')
    if raw.strip():
        try:
            loaded = json.loads(raw)
            if isinstance(loaded, dict):
                mode = str(loaded.get("mode", "numerical")).lower()
                if mode not in ("numerical", "custom"):
                    mode = "numerical"
                raw_mappings = loaded.get("mappings", [])
                cleaned_mappings: List[Dict[str, Any]] = []
                if isinstance(raw_mappings, list):
                    for item in raw_mappings:
                        if not isinstance(item, dict):
                            continue
                        code = str(item.get("code", "")).strip()
                        survey = str(item.get("survey", "*")).strip() or "*"
                        meaning = str(item.get("meaning", "")).strip()
                        try:
                            order_val = int(item.get("order", 0))
                        except (ValueError, TypeError):
                            order_val = 0
                        if code and order_val >= 1:
                            cleaned_mappings.append({
                                "survey": survey,
                                "code": code,
                                "order": order_val,
                                "meaning": meaning
                            })
                return {"mode": mode, "mappings": cleaned_mappings}
        except Exception:
            return default_condition_ordering()
    return default_condition_ordering()


def save_condition_ordering_to_config(config_name: str, ordering: Dict[str, Any]) -> None:
    if not config_name:
        return
    cfg = new_config_parser()
    cfg.read(get_config_path(), encoding='utf-8')
    if not cfg.has_section(config_name):
        cfg.add_section(config_name)
    cfg.set(config_name, 'condition_ordering_json', json.dumps(ordering, ensure_ascii=False))
    write_config_atomically(cfg, get_config_path())


def load_mapping_from_config(config_name: str) -> Dict[str, str]:
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


def save_mapping_to_config(config_name: str, mapping: Dict[str, str]) -> None:
    if not config_name:
        return
    cfg = new_config_parser()
    cfg.read(get_config_path(), encoding='utf-8')
    if not cfg.has_section(config_name):
        cfg.add_section(config_name)
    cfg.set(config_name, 'field_mapping_json', json.dumps(mapping, ensure_ascii=False))
    write_config_atomically(cfg, get_config_path())