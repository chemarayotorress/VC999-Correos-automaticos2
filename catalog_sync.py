from __future__ import annotations

import csv
import io
import json
import logging
import os
import threading
import time
import urllib.parse
import urllib.request
from datetime import datetime, timezone
from decimal import Decimal
from typing import Any, Dict, List, Optional, Tuple

from machine_catalog import load_catalog, set_runtime_catalog

logger = logging.getLogger(__name__)

BASE_STEP_NAMES = {"", "base", "baseprice", "precio_base", "preciobase", "base_price", "pricebase"}


class CatalogSyncError(RuntimeError):
    pass


class CatalogSyncManager:
    def __init__(self):
        self._lock = threading.RLock()
        self._last_sync_monotonic: float = 0.0
        self._last_status: Dict[str, Any] = {
            "ok": False,
            "source": "none",
            "updated_at": None,
            "error": "Never synced",
        }

    def sync(self, force: bool = False, persist_cache: bool = True) -> Dict[str, Any]:
        ttl_seconds = max(0, int(os.getenv("CATALOG_SYNC_TTL_SECONDS", "300") or "300"))
        with self._lock:
            if not force and self._last_sync_monotonic > 0:
                age = time.monotonic() - self._last_sync_monotonic
                if age < ttl_seconds:
                    return dict(self._last_status)

            updated_at = datetime.now(timezone.utc).isoformat()
            try:
                rows_machines, rows_prices, source = _load_sheet_rows()
                catalog = _build_catalog(rows_machines, rows_prices)
                if not catalog:
                    raise CatalogSyncError("Google Sheets returned an empty catalog")
                set_runtime_catalog(catalog, persist=persist_cache)
                logger.info("Catalog synced from Google Sheets")
                self._last_sync_monotonic = time.monotonic()
                self._last_status = {
                    "ok": True,
                    "source": "sheets",
                    "mode": source,
                    "updated_at": updated_at,
                    "items": len(catalog),
                }
                return dict(self._last_status)
            except Exception as exc:
                logger.exception("Catalog sync failed, using local machines.json cache")
                fallback = load_catalog(force_disk=True)
                set_runtime_catalog(fallback, persist=False)
                logger.warning("Fallback to local machines.json")
                self._last_sync_monotonic = time.monotonic()
                self._last_status = {
                    "ok": False,
                    "source": "local",
                    "updated_at": updated_at,
                    "error": str(exc),
                    "items": len(fallback),
                }
                return dict(self._last_status)


_MANAGER = CatalogSyncManager()


def sync_catalog(force: bool = False, persist_cache: bool = True) -> Dict[str, Any]:
    return _MANAGER.sync(force=force, persist_cache=persist_cache)


def _normalize_key(key: Any) -> str:
    return str(key or "").strip().lower().replace(" ", "").replace("_", "").replace("-", "")


def _row_value(row: Dict[str, Any], aliases: List[str]) -> Any:
    if not row:
        return None
    keys = { _normalize_key(k): k for k in row.keys() }
    for alias in aliases:
        normalized = _normalize_key(alias)
        if normalized in keys:
            return row.get(keys[normalized])
    for alias in aliases:
        normalized = _normalize_key(alias)
        for key in row.keys():
            if normalized and normalized in _normalize_key(key):
                return row.get(key)
    return None


def _to_decimal(value: Any) -> Decimal:
    if value is None:
        return Decimal("0")
    text = str(value).strip()
    if not text:
        return Decimal("0")
    cleaned = "".join(ch for ch in text if ch.isdigit() or ch in {".", "-"})
    if not cleaned or cleaned in {".", "-", "-."}:
        return Decimal("0")
    try:
        return Decimal(cleaned)
    except Exception:
        return Decimal("0")


def _is_base_step(step: Any) -> bool:
    return _normalize_key(step) in BASE_STEP_NAMES


def _sheet_gid_url(sheet_id: str, sheet_name: str) -> str:
    return (
        f"https://docs.google.com/spreadsheets/d/{urllib.parse.quote(sheet_id)}"
        f"/gviz/tq?tqx=out:csv&sheet={urllib.parse.quote(sheet_name)}"
    )


def _read_csv_tab(sheet_id: str, sheet_name: str) -> List[Dict[str, Any]]:
    url = _sheet_gid_url(sheet_id, sheet_name)
    timeout = float(os.getenv("CATALOG_SYNC_TIMEOUT_SECONDS", "15") or "15")
    req = urllib.request.Request(url, headers={"User-Agent": "vc999-catalog-sync/1.0"})
    with urllib.request.urlopen(req, timeout=timeout) as resp:
        raw = resp.read().decode("utf-8-sig")
    reader = csv.DictReader(io.StringIO(raw))
    return [dict(row) for row in reader]


def _load_sheet_rows_csv(sheet_id: str) -> Tuple[List[Dict[str, Any]], List[Dict[str, Any]]]:
    machines = _read_csv_tab(sheet_id, "DB_Maquinas")
    prices = _read_csv_tab(sheet_id, "DB_Precios")
    return machines, prices


def _load_sheet_rows_service_account(sheet_id: str, credentials_path: str) -> Tuple[List[Dict[str, Any]], List[Dict[str, Any]]]:
    try:
        from google.oauth2.service_account import Credentials
        from googleapiclient.discovery import build
    except Exception as exc:  # pragma: no cover - optional deps
        raise CatalogSyncError(
            "google-api-python-client/google-auth are required for service account mode"
        ) from exc

    scopes = ["https://www.googleapis.com/auth/spreadsheets.readonly"]
    credentials = Credentials.from_service_account_file(credentials_path, scopes=scopes)
    service = build("sheets", "v4", credentials=credentials, cache_discovery=False)
    sheet = service.spreadsheets()

    def read_range(range_name: str) -> List[Dict[str, Any]]:
        result = sheet.values().get(spreadsheetId=sheet_id, range=range_name).execute()
        values = result.get("values", [])
        if not values:
            return []
        headers = [str(h).strip() for h in values[0]]
        rows: List[Dict[str, Any]] = []
        for row in values[1:]:
            payload = {headers[idx]: (row[idx] if idx < len(row) else "") for idx in range(len(headers))}
            rows.append(payload)
        return rows

    machines = read_range("DB_Maquinas!A:ZZ")
    prices = read_range("DB_Precios!A:ZZ")
    return machines, prices


def _load_sheet_rows() -> Tuple[List[Dict[str, Any]], List[Dict[str, Any]], str]:
    sheet_id = (os.getenv("GOOGLE_SHEET_ID") or "").strip()
    credentials_path = (os.getenv("GOOGLE_APPLICATION_CREDENTIALS") or "").strip()

    if credentials_path:
        if not sheet_id:
            raise CatalogSyncError("GOOGLE_SHEET_ID is required with service account mode")
        machines, prices = _load_sheet_rows_service_account(sheet_id, credentials_path)
        return machines, prices, "sheets_service_account"

    if sheet_id:
        machines, prices = _load_sheet_rows_csv(sheet_id)
        return machines, prices, "sheets_csv"

    raise CatalogSyncError("Missing GOOGLE_SHEET_ID (CSV mode) or GOOGLE_APPLICATION_CREDENTIALS (service mode)")


def _build_catalog(rows_machines: List[Dict[str, Any]], rows_prices: List[Dict[str, Any]]) -> Dict[str, Dict[str, Any]]:
    model_meta: Dict[str, Dict[str, Any]] = {}

    for row in rows_machines:
        model = _row_value(row, ["modelo", "model", "maquina", "maquina_id", "id"])
        if not model:
            continue
        model_str = str(model).strip()
        template = _row_value(row, ["plantilla", "template", "docx", "archivo"])
        template_str = str(template).strip() if template else model_str
        if template_str and not template_str.lower().endswith(".docx"):
            template_str = f"{template_str}.docx"
        base_price = _to_decimal(_row_value(row, ["precio_base", "base_price", "precio", "price", "costo"]))
        model_meta[model_str] = {
            "template": template_str,
            "base": base_price,
            "options": {},
        }

    grouped_options: Dict[Tuple[str, str], Dict[str, Any]] = {}

    for row in rows_prices:
        model = _row_value(row, ["modelo", "model", "maquina", "maquina_id", "id"])
        if not model:
            continue
        model_str = str(model).strip()
        step = _row_value(row, ["paso", "paso_id", "step", "step_id", "pregunta", "question"])
        step_str = str(step).strip() if step else ""
        label = _row_value(row, ["opcion", "opcion_label", "label", "nombre_opcion", "opcion_valor", "value"])
        if label is None:
            continue
        label_str = str(label).strip()
        if not label_str:
            continue

        price = _to_decimal(_row_value(row, ["precio", "price", "extra", "costo", "precio_extra_us"]))
        opt_type_raw = _row_value(row, ["tipo", "type", "input_type", "control", "tipo_control"])
        opt_type = _normalize_key(opt_type_raw)

        if model_str not in model_meta:
            fallback_template = model_str if model_str.lower().endswith(".docx") else f"{model_str}.docx"
            model_meta[model_str] = {"template": fallback_template, "base": Decimal("0"), "options": {}}

        if _is_base_step(step_str):
            if price > 0:
                model_meta[model_str]["base"] = price
            continue

        option_name = step_str or "Option"
        key = (model_str, option_name)
        bucket = grouped_options.setdefault(
            key,
            {"type": "checkbox" if opt_type in {"checkbox", "bool", "boolean", "check"} else "select", "choices": []},
        )
        bucket["choices"].append({"label": label_str, "price": float(price)})
        if bucket["type"] != "checkbox" and opt_type in {"checkbox", "bool", "boolean", "check"}:
            bucket["type"] = "checkbox"

    for (model_str, option_name), payload in grouped_options.items():
        meta = model_meta[model_str]
        choices = payload["choices"]
        if payload["type"] == "checkbox":
            max_price = max((Decimal(str(item["price"])) for item in choices), default=Decimal("0"))
            meta["options"][option_name] = {"type": "checkbox", "price": float(max_price)}
        else:
            meta["options"][option_name] = {"type": "select", "choices": choices}

    catalog: Dict[str, Dict[str, Any]] = {}
    for model_str, meta in model_meta.items():
        template = meta["template"]
        catalog[template] = {
            "base": float(meta["base"]),
            "options": meta["options"],
        }

    return catalog
