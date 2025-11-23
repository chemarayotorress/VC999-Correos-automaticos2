"""Client utilities to talk with the optional multiuser backend."""
from __future__ import annotations

import json
import os
import ssl
import urllib.error
import urllib.parse
import urllib.request
from typing import Any, Dict, Iterable, Optional


class BackendError(RuntimeError):
    """Base error for backend operations."""


class BackendUnavailable(BackendError):
    """Raised when the backend cannot be reached."""


class UnauthorizedError(BackendError):
    """Raised when the backend rejects the provided credentials."""


class BackendClient:
    """Simple HTTP client for the optional quotation backend service."""

    def __init__(self, base_url: Optional[str], token: Optional[str] = None, timeout: int = 8):
        self._base_url = (base_url or "").rstrip("/")
        self._token = token or ""
        self._timeout = timeout
        self._ssl_context = ssl.create_default_context() if hasattr(ssl, "create_default_context") else None

    @property
    def base_url(self) -> str:
        return self._base_url

    def is_enabled(self) -> bool:
        return bool(self._base_url)

    def set_token(self, token: Optional[str]):
        self._token = token or ""

    def _request(self, method: str, path: str, payload: Optional[Dict[str, Any]] = None) -> Dict[str, Any]:
        if not self._base_url:
            raise BackendUnavailable("Backend URL is not configured")
        url = f"{self._base_url}{path}"
        data: Optional[bytes] = None
        headers = {"Accept": "application/json"}
        if payload is not None:
            data = json.dumps(payload).encode("utf-8")
            headers["Content-Type"] = "application/json"
        if self._token:
            headers["Authorization"] = f"Bearer {self._token}"
        req = urllib.request.Request(url, data=data, headers=headers, method=method)
        try:
            if self._ssl_context is not None and url.lower().startswith("https"):
                response = urllib.request.urlopen(req, timeout=self._timeout, context=self._ssl_context)
            else:
                response = urllib.request.urlopen(req, timeout=self._timeout)
            raw = response.read().decode("utf-8")
            if not raw:
                return {}
            return json.loads(raw)
        except urllib.error.HTTPError as exc:  # pragma: no cover - network is optional
            if exc.code in (401, 403):
                raise UnauthorizedError(f"Backend rejected credentials ({exc.code})") from exc
            raise BackendError(f"Backend error {exc.code}: {exc.read().decode('utf-8', 'ignore')}") from exc
        except urllib.error.URLError as exc:  # pragma: no cover - network is optional
            raise BackendUnavailable(str(exc)) from exc

    # --- Auth ---
    def login(self, username: str, password: str, device_id: str, license_key: Optional[str] = None) -> Dict[str, Any]:
        payload = {
            "username": username,
            "password": password,
            "device_id": device_id,
        }
        if license_key:
            payload["license_key"] = license_key
        return self._request("POST", "/api/auth/login", payload)

    def validate_token(self) -> Dict[str, Any]:
        return self._request("GET", "/api/auth/token")

    # --- Quotes ---
    def list_quotes(self, kind: str) -> Iterable[Dict[str, Any]]:
        data = self._request("GET", f"/api/quotes?type={urllib.parse.quote(kind)}")
        return data.get("items", []) if isinstance(data, dict) else []

    def create_quote(self, kind: str, quote: Dict[str, Any]) -> Dict[str, Any]:
        payload = {"type": kind, "quote": quote}
        return self._request("POST", "/api/quotes", payload)

    def delete_quote(self, kind: str, quote_id: str) -> Dict[str, Any]:
        return self._request("DELETE", f"/api/quotes/{urllib.parse.quote(quote_id)}?type={urllib.parse.quote(kind)}")

    def metrics(self) -> Dict[str, Any]:
        return self._request("GET", "/api/metrics")


def backend_from_config(cfg: Dict[str, Any]) -> BackendClient:
    base_url = cfg.get("backend_url") if isinstance(cfg, dict) else None
    token = cfg.get("auth_token") if isinstance(cfg, dict) else None
    return BackendClient(base_url=base_url, token=token)
