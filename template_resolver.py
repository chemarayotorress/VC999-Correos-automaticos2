from __future__ import annotations

import os
import re
from functools import lru_cache
from pathlib import Path
from typing import Dict, List, Optional

DOCX_EXTENSION_RE = re.compile(r"\.docx$", re.IGNORECASE)
DEFAULT_TEMPLATE_PATTERN = re.compile(r"^(CM|TS)[A-Z0-9]+$")
EXCLUDED_TEMPLATE_STEMS = {"COTIZACION MATERIALS"}


def _app_dir() -> Path:
    return Path(__file__).resolve().parent


def normalize_model(value: Optional[str]) -> str:
    if value is None:
        return ""
    normalized = str(value).strip()
    normalized = DOCX_EXTENSION_RE.sub("", normalized)
    normalized = re.sub(r"\s+", "", normalized)
    return normalized.upper()


@lru_cache(maxsize=1)
def get_templates_map() -> Dict[str, Path]:
    root = _app_dir()
    templates: Dict[str, Path] = {}

    for docx_file in root.glob("*.docx"):
        stem_normalized = normalize_model(docx_file.stem)
        if not stem_normalized:
            continue
        if stem_normalized in EXCLUDED_TEMPLATE_STEMS:
            continue
        if not DEFAULT_TEMPLATE_PATTERN.match(stem_normalized):
            continue
        templates[stem_normalized] = docx_file.resolve()

    return templates


def resolve_template_path(model: Optional[str]) -> Optional[Path]:
    normalized = normalize_model(model)
    if not normalized:
        return None
    return get_templates_map().get(normalized)


def list_available_models(limit: Optional[int] = None) -> List[str]:
    models = sorted(get_templates_map().keys())
    if limit is not None and limit > 0:
        return models[:limit]
    return models


def clear_templates_cache() -> None:
    get_templates_map.cache_clear()
