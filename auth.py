"""Authentication and licensing utilities for the VC999 app."""
from __future__ import annotations

import base64
import hashlib
import os
import time
import uuid
from dataclasses import dataclass
from typing import Any, Callable, Dict, Optional

try:
    import tkinter as tk
    from tkinter import ttk, messagebox
except Exception:  # pragma: no cover - headless environments
    tk = None  # type: ignore
    ttk = None  # type: ignore
    messagebox = None  # type: ignore

from backend_client import BackendClient, BackendUnavailable, UnauthorizedError


@dataclass
class AuthToken:
    token: str
    username: str
    license_key: str
    expires_at: float


class AuthError(RuntimeError):
    pass


class AuthManager:
    """Coordinates login and token storage with optional backend validation."""

    def __init__(self, backend: BackendClient, cfg: Dict[str, Any], save_cfg: Callable[[Dict[str, Any]], None]):
        self.backend = backend
        self.cfg = cfg
        self._save_cfg = save_cfg
        self.token: Optional[AuthToken] = None
        stored = cfg.get("auth_token")
        if isinstance(stored, str):
            try:
                self.token = self._decode_token(stored)
                if backend.is_enabled():
                    backend.set_token(self.token.token)
            except Exception:
                self.token = None

    # --- helpers ---
    @staticmethod
    def _machine_id() -> str:
        node = uuid.getnode()
        host = os.uname().nodename if hasattr(os, "uname") else os.getenv("COMPUTERNAME", "unknown")
        return f"{node:x}-{host}".lower()

    def _cipher_key(self) -> bytes:
        return hashlib.sha256(self._machine_id().encode("utf-8")).digest()

    def _encode_token(self, token: AuthToken) -> str:
        payload = "|".join([
            token.token,
            token.username,
            token.license_key,
            f"{int(token.expires_at)}",
        ]).encode("utf-8")
        key = self._cipher_key()
        encoded = bytes(b ^ key[i % len(key)] for i, b in enumerate(payload))
        return base64.urlsafe_b64encode(encoded).decode("ascii")

    def _decode_token(self, value: str) -> AuthToken:
        data = base64.urlsafe_b64decode(value.encode("ascii"))
        key = self._cipher_key()
        payload = bytes(b ^ key[i % len(key)] for i, b in enumerate(data)).decode("utf-8")
        token, username, license_key, expires_raw = payload.split("|", 3)
        return AuthToken(token=token, username=username, license_key=license_key, expires_at=float(expires_raw))

    def _persist_token(self):
        if self.token:
            self.cfg["auth_token"] = self._encode_token(self.token)
        else:
            self.cfg.pop("auth_token", None)
        self._save_cfg(self.cfg)

    # --- public API ---
    def ensure_login(self) -> bool:
        if self.token and self.token.expires_at > time.time():
            if self.backend.is_enabled():
                try:
                    self.backend.set_token(self.token.token)
                    self.backend.validate_token()
                    return True
                except UnauthorizedError:
                    self.logout()
                except BackendUnavailable:
                    # allow offline access while token valid
                    return True
            else:
                return True
        if not self.backend.is_enabled():
            return True
        return self._login_flow()

    def _login_flow(self) -> bool:
        if tk is None:
            raise AuthError("Tkinter no disponible para autenticación")
        root = tk.Tk()
        root.withdraw()
        dialog = LoginDialog(root, self.backend)
        result = dialog.show()
        root.destroy()
        if not result:
            return False
        token = AuthToken(
            token=result["token"],
            username=result.get("username", ""),
            license_key=result.get("license_key", ""),
            expires_at=float(result.get("expires_at", time.time() + 3600)),
        )
        self.token = token
        self.backend.set_token(token.token)
        self._persist_token()
        return True

    def logout(self):
        self.token = None
        self.backend.set_token(None)
        self._persist_token()


class LoginDialog:
    """Modal dialog that handles login/activation."""

    def __init__(self, master, backend: BackendClient):
        if tk is None:
            raise RuntimeError("Tkinter is required for LoginDialog")
        self.backend = backend
        self.master = master
        self.top = tk.Toplevel(master)
        self.top.title("Acceso al cotizador VC999")
        self.top.geometry("420x260")
        self.top.grab_set()
        self.result: Optional[Dict[str, Any]] = None

        tk.Label(self.top, text="Correo electrónico").pack(pady=(18, 2))
        self.var_user = tk.StringVar()
        tk.Entry(self.top, textvariable=self.var_user, width=40).pack()

        tk.Label(self.top, text="Contraseña").pack(pady=(12, 2))
        self.var_pass = tk.StringVar()
        tk.Entry(self.top, textvariable=self.var_pass, show="*", width=40).pack()

        tk.Label(self.top, text="Clave de licencia").pack(pady=(12, 2))
        self.var_license = tk.StringVar()
        tk.Entry(self.top, textvariable=self.var_license, width=40).pack()

        self.var_status = tk.StringVar(value="Introduce tus credenciales")
        tk.Label(self.top, textvariable=self.var_status, fg="#666").pack(pady=(12, 0))

        btns = tk.Frame(self.top)
        btns.pack(pady=16)
        ttk.Button(btns, text="Cancelar", command=self._on_cancel).pack(side="right", padx=6)
        ttk.Button(btns, text="Iniciar sesión", command=self._on_submit).pack(side="right", padx=6)

        self.top.protocol("WM_DELETE_WINDOW", self._on_cancel)
        self.top.bind("<Return>", lambda *_: self._on_submit())

    def _on_cancel(self):
        self.result = None
        self.top.destroy()

    def _on_submit(self):
        username = (self.var_user.get() or "").strip()
        password = (self.var_pass.get() or "").strip()
        license_key = (self.var_license.get() or "").strip()
        if not username or not password:
            self.var_status.set("Usuario y contraseña requeridos")
            return
        if not self.backend.is_enabled():
            self.var_status.set("Configura el backend en app_config.json")
            return
        self.var_status.set("Validando...")
        self.top.update_idletasks()
        try:
            payload = self.backend.login(username, password, AuthManager._machine_id(), license_key)
            token = payload.get("token")
            if not token:
                raise AuthError("El backend no devolvió token")
            self.result = {
                "token": token,
                "username": payload.get("username", username),
                "license_key": payload.get("license_key", license_key),
                "expires_at": payload.get("expires_at", time.time() + 3600),
            }
            self.top.destroy()
        except UnauthorizedError as exc:
            self.var_status.set(f"Acceso denegado: {exc}")
        except BackendUnavailable as exc:
            self.var_status.set(f"Backend no disponible: {exc}")
        except Exception as exc:  # pragma: no cover - defensive
            if messagebox:
                messagebox.showerror("Error", f"No fue posible iniciar sesión:\n{exc}")
            self.var_status.set("Error inesperado")

    def show(self) -> Optional[Dict[str, Any]]:
        self.master.wait_window(self.top)
        return self.result
