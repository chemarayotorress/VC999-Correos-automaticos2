"""Machine catalog persistence and editor."""
from __future__ import annotations

import copy
import json
import os
import re
import threading
import uuid
from dataclasses import dataclass
from decimal import Decimal
from typing import Dict, List, Optional, Tuple

try:
    import tkinter as tk
    from tkinter import ttk, messagebox
except Exception:  # pragma: no cover - headless
    tk = None  # type: ignore
    ttk = None  # type: ignore
    messagebox = None  # type: ignore

CATALOG_FILE = "machines.json"
BACKUP_DIR = "respaldos"


@dataclass
class MachineOptionChoice:
    label: str
    price: Decimal


@dataclass
class MachineOption:
    name: str
    opt_type: str  # "checkbox" or "select"
    price: Decimal = Decimal("0")
    choices: Optional[List[MachineOptionChoice]] = None


@dataclass
class MachineConfig:
    template: str
    base_price: Decimal
    options: List[MachineOption]


DEFAULT_MACHINE_CATALOG: Dict[str, Dict] = {}
_RUNTIME_CATALOG: Optional[Dict[str, Dict]] = None
_CATALOG_LOCK = threading.RLock()


def _read_catalog_file(path: str) -> Optional[Dict[str, Dict]]:
    try:
        with open(path, "r", encoding="utf-8") as f:
            raw = f.read()
    except Exception:
        return None
    for candidate in (raw, re.sub(r"(?<=\d),(?=\d)", "", raw)):
        try:
            data = json.loads(candidate)
            if isinstance(data, dict):
                return data
        except Exception:
            continue
    return None


def _default_catalog() -> Dict[str, Dict]:
    global DEFAULT_MACHINE_CATALOG
    if DEFAULT_MACHINE_CATALOG:
        return copy.deepcopy(DEFAULT_MACHINE_CATALOG)
    path = os.path.join(os.path.dirname(os.path.abspath(__file__)), CATALOG_FILE)
    if os.path.exists(path):
        data = _read_catalog_file(path)
        if isinstance(data, dict):
            DEFAULT_MACHINE_CATALOG = data
            return copy.deepcopy(DEFAULT_MACHINE_CATALOG)
    return {}


def load_catalog(force_disk: bool = False) -> Dict[str, Dict]:
    global _RUNTIME_CATALOG
    with _CATALOG_LOCK:
        if _RUNTIME_CATALOG is not None and not force_disk:
            return copy.deepcopy(_RUNTIME_CATALOG)

        path = os.path.join(os.path.dirname(os.path.abspath(__file__)), CATALOG_FILE)
        if os.path.exists(path):
            data = _read_catalog_file(path)
            if isinstance(data, dict):
                _RUNTIME_CATALOG = copy.deepcopy(data)
                return copy.deepcopy(_RUNTIME_CATALOG)

        fallback = _default_catalog()
        _RUNTIME_CATALOG = copy.deepcopy(fallback)
        return copy.deepcopy(_RUNTIME_CATALOG)


def set_runtime_catalog(catalog: Dict[str, Dict], persist: bool = False):
    global _RUNTIME_CATALOG
    with _CATALOG_LOCK:
        _RUNTIME_CATALOG = copy.deepcopy(catalog or {})
        if persist:
            save_catalog(_RUNTIME_CATALOG)


def save_catalog(catalog: Dict[str, Dict]):
    appdir = os.path.dirname(os.path.abspath(__file__))
    path = os.path.join(appdir, CATALOG_FILE)
    os.makedirs(appdir, exist_ok=True)
    # Backup previous version
    try:
        if os.path.exists(path):
            os.makedirs(os.path.join(appdir, BACKUP_DIR), exist_ok=True)
            backup = os.path.join(appdir, BACKUP_DIR, f"machines_{uuid.uuid4().hex}.json")
            with open(path, "rb") as src, open(backup, "wb") as dst:
                dst.write(src.read())
    except Exception:
        pass
    with open(path, "w", encoding="utf-8") as f:
        json.dump(catalog, f, ensure_ascii=False, indent=2)

    with _CATALOG_LOCK:
        global _RUNTIME_CATALOG
        _RUNTIME_CATALOG = copy.deepcopy(catalog or {})


# --- Tk editor -----------------------------------------------------------
class MachineCatalogEditor:
    def __init__(self, master, catalog: Dict[str, Dict], on_save):
        if tk is None:
            raise RuntimeError("Tkinter no disponible")
        self.master = master
        self.on_save = on_save
        self.catalog = copy.deepcopy(catalog)
        self.top = tk.Toplevel(master)
        self.top.title("Catálogo de máquinas")
        self.top.geometry("720x500")
        self.top.grab_set()

        self.tree = ttk.Treeview(self.top, columns=("base",), show="headings")
        self.tree.heading("base", text="Precio base")
        self.tree.pack(fill="both", expand=True, padx=10, pady=10)

        btns = tk.Frame(self.top)
        btns.pack(fill="x", padx=10, pady=(0, 10))
        ttk.Button(btns, text="Agregar", command=self._add_machine).pack(side="left")
        ttk.Button(btns, text="Editar", command=self._edit_machine).pack(side="left", padx=6)
        ttk.Button(btns, text="Eliminar", command=self._delete_machine).pack(side="left")
        ttk.Button(btns, text="Guardar", command=self._save).pack(side="right")
        ttk.Button(btns, text="Cancelar", command=self._cancel).pack(side="right", padx=6)

        self._refresh_tree()
        self.top.protocol("WM_DELETE_WINDOW", self._cancel)

    def _refresh_tree(self):
        for item in self.tree.get_children():
            self.tree.delete(item)
        for template, data in sorted(self.catalog.items()):
            base = data.get("base", 0)
            self.tree.insert("", "end", iid=template, values=(f"US${base:,.2f}" if isinstance(base, (int, float)) else base,))

    def _add_machine(self):
        MachineDialog(self.top, None, self.catalog, self._refresh_tree)

    def _edit_machine(self):
        sel = self.tree.selection()
        if not sel:
            return
        template = sel[0]
        MachineDialog(self.top, template, self.catalog, self._refresh_tree)

    def _delete_machine(self):
        sel = self.tree.selection()
        if not sel:
            return
        template = sel[0]
        if messagebox and not messagebox.askyesno("Eliminar", f"¿Eliminar {template}?"):
            return
        self.catalog.pop(template, None)
        self._refresh_tree()

    def _save(self):
        try:
            self.on_save(self.catalog)
            self.top.destroy()
        except Exception as exc:
            if messagebox:
                messagebox.showerror("Error", f"No se pudo guardar:\n{exc}")

    def _cancel(self):
        self.top.destroy()


class MachineDialog:
    def __init__(self, master, template: Optional[str], catalog: Dict[str, Dict], on_change):
        if tk is None:
            raise RuntimeError("Tkinter no disponible")
        self.catalog = catalog
        self.on_change = on_change
        self.template = template
        data = catalog.get(template, {}) if template else {}

        self.top = tk.Toplevel(master)
        self.top.title("Editar máquina" if template else "Agregar máquina")
        self.top.geometry("620x460")
        self.top.grab_set()

        tk.Label(self.top, text="Archivo plantilla (.docx)").pack(pady=(12, 4))
        self.var_template = tk.StringVar(value=template or "")
        tk.Entry(self.top, textvariable=self.var_template, width=50).pack()

        tk.Label(self.top, text="Precio base (USD)").pack(pady=(12, 4))
        self.var_base = tk.StringVar(value=str(data.get("base", "0")))
        tk.Entry(self.top, textvariable=self.var_base, width=20).pack()

        tk.Label(self.top, text="Opciones").pack(pady=(12, 4))
        self.opt_tree = ttk.Treeview(self.top, columns=("tipo", "detalle"), show="headings", height=8)
        self.opt_tree.heading("tipo", text="Tipo")
        self.opt_tree.heading("detalle", text="Detalle")
        self.opt_tree.pack(fill="both", expand=True, padx=10, pady=4)

        opt_frame = tk.Frame(self.top)
        opt_frame.pack(pady=6)
        ttk.Button(opt_frame, text="Agregar opción", command=self._add_option).pack(side="left")
        ttk.Button(opt_frame, text="Editar opción", command=self._edit_option).pack(side="left", padx=6)
        ttk.Button(opt_frame, text="Eliminar opción", command=self._delete_option).pack(side="left")

        self.options = copy.deepcopy(data.get("options", {}))
        self._refresh_options()

        btns = tk.Frame(self.top)
        btns.pack(fill="x", pady=10)
        ttk.Button(btns, text="Cancelar", command=self._cancel).pack(side="right", padx=6)
        ttk.Button(btns, text="Aceptar", command=self._accept).pack(side="right")

        self.top.protocol("WM_DELETE_WINDOW", self._cancel)

    def _refresh_options(self):
        for item in self.opt_tree.get_children():
            self.opt_tree.delete(item)
        for name, data in self.options.items():
            opt_type = data.get("type", "select")
            if opt_type == "checkbox":
                detail = f"Checkbox - US${data.get('price', 0):,.2f}"
            else:
                choices = data.get("choices", [])
                detail = f"Select ({len(choices)} opciones)"
            self.opt_tree.insert("", "end", iid=name, values=(opt_type, detail))

    def _add_option(self):
        OptionDialog(self.top, None, self.options, self._refresh_options)

    def _edit_option(self):
        sel = self.opt_tree.selection()
        if not sel:
            return
        OptionDialog(self.top, sel[0], self.options, self._refresh_options)

    def _delete_option(self):
        sel = self.opt_tree.selection()
        if not sel:
            return
        self.options.pop(sel[0], None)
        self._refresh_options()

    def _accept(self):
        template = (self.var_template.get() or "").strip()
        if not template or not template.lower().endswith(".docx"):
            if messagebox:
                messagebox.showerror("Datos inválidos", "La plantilla debe terminar en .docx")
            return
        try:
            base = float(self.var_base.get().replace(",", ""))
        except Exception:
            if messagebox:
                messagebox.showerror("Datos inválidos", "Precio base no válido")
            return
        self.catalog.pop(self.template, None)
        self.catalog[template] = {
            "base": base,
            "options": copy.deepcopy(self.options),
        }
        self.on_change()
        self.top.destroy()

    def _cancel(self):
        self.top.destroy()


class OptionDialog:
    def __init__(self, master, name: Optional[str], options: Dict[str, Dict], on_change):
        if tk is None:
            raise RuntimeError("Tkinter no disponible")
        self.options = options
        self.on_change = on_change
        data = options.get(name, {}) if name else {}

        self.top = tk.Toplevel(master)
        self.top.title("Editar opción" if name else "Agregar opción")
        self.top.geometry("520x420")
        self.top.grab_set()

        tk.Label(self.top, text="Nombre de la opción").pack(pady=(10, 4))
        self.var_name = tk.StringVar(value=name or "")
        tk.Entry(self.top, textvariable=self.var_name, width=45).pack()

        tk.Label(self.top, text="Tipo").pack(pady=(12, 4))
        self.var_type = tk.StringVar(value=data.get("type", "select"))
        type_cb = ttk.Combobox(self.top, textvariable=self.var_type, values=["select", "checkbox"], state="readonly")
        type_cb.pack()
        type_cb.bind("<<ComboboxSelected>>", lambda *_: self._toggle_state())

        self.frame_checkbox = tk.Frame(self.top)
        tk.Label(self.frame_checkbox, text="Precio adicional (USD)").pack(pady=(8, 4))
        self.var_price = tk.StringVar(value=str(data.get("price", "0")))
        tk.Entry(self.frame_checkbox, textvariable=self.var_price, width=18).pack()

        self.frame_select = tk.Frame(self.top)
        self.choice_tree = ttk.Treeview(self.frame_select, columns=("precio",), show="headings", height=6)
        self.choice_tree.heading("precio", text="Precio")
        self.choice_tree.pack(fill="both", expand=True, padx=6, pady=4)
        c_btns = tk.Frame(self.frame_select)
        c_btns.pack(pady=4)
        ttk.Button(c_btns, text="Agregar", command=self._add_choice).pack(side="left")
        ttk.Button(c_btns, text="Editar", command=self._edit_choice).pack(side="left", padx=6)
        ttk.Button(c_btns, text="Eliminar", command=self._delete_choice).pack(side="left")

        self.choices = copy.deepcopy(data.get("choices", []))
        self._refresh_choices()
        self._toggle_state()

        btns = tk.Frame(self.top)
        btns.pack(fill="x", pady=10)
        ttk.Button(btns, text="Cancelar", command=self._cancel).pack(side="right", padx=6)
        ttk.Button(btns, text="Aceptar", command=self._accept).pack(side="right")

        self.top.protocol("WM_DELETE_WINDOW", self._cancel)

    def _toggle_state(self):
        mode = self.var_type.get()
        if mode == "checkbox":
            self.frame_select.pack_forget()
            self.frame_checkbox.pack(fill="x", pady=6)
        else:
            self.frame_checkbox.pack_forget()
            self.frame_select.pack(fill="both", expand=True, padx=6, pady=6)

    def _refresh_choices(self):
        for item in self.choice_tree.get_children():
            self.choice_tree.delete(item)
        for choice in self.choices:
            label = choice.get("label", "")
            price = choice.get("price", 0)
            self.choice_tree.insert("", "end", iid=label, values=(f"US${price:,.2f}",))

    def _add_choice(self):
        ChoiceDialog(self.top, None, self.choices, self._refresh_choices)

    def _edit_choice(self):
        sel = self.choice_tree.selection()
        if not sel:
            return
        ChoiceDialog(self.top, sel[0], self.choices, self._refresh_choices)

    def _delete_choice(self):
        sel = self.choice_tree.selection()
        if not sel:
            return
        label = sel[0]
        self.choices = [c for c in self.choices if c.get("label") != label]
        self._refresh_choices()

    def _accept(self):
        name = (self.var_name.get() or "").strip()
        if not name:
            if messagebox:
                messagebox.showerror("Datos inválidos", "La opción necesita un nombre")
            return
        opt_type = self.var_type.get()
        if opt_type == "checkbox":
            try:
                price = float(self.var_price.get().replace(",", ""))
            except Exception:
                if messagebox:
                    messagebox.showerror("Datos inválidos", "Precio no válido")
                return
            self.options[name] = {"type": "checkbox", "price": price}
        else:
            if not self.choices:
                if messagebox:
                    messagebox.showerror("Datos inválidos", "Agrega al menos una opción")
                return
            self.options[name] = {"type": "select", "choices": copy.deepcopy(self.choices)}
        self.on_change()
        self.top.destroy()

    def _cancel(self):
        self.top.destroy()


class ChoiceDialog:
    def __init__(self, master, label: Optional[str], choices: List[Dict[str, float]], on_change):
        if tk is None:
            raise RuntimeError("Tkinter no disponible")
        self.choices = choices
        self.on_change = on_change
        existing = next((c for c in choices if c.get("label") == label), None)

        self.top = tk.Toplevel(master)
        self.top.title("Editar valor" if label else "Agregar valor")
        self.top.geometry("420x200")
        self.top.grab_set()

        tk.Label(self.top, text="Etiqueta").pack(pady=(12, 4))
        self.var_label = tk.StringVar(value=existing.get("label", "") if existing else "")
        tk.Entry(self.top, textvariable=self.var_label, width=40).pack()

        tk.Label(self.top, text="Precio (USD)").pack(pady=(12, 4))
        self.var_price = tk.StringVar(value=str(existing.get("price", "0")) if existing else "0")
        tk.Entry(self.top, textvariable=self.var_price, width=20).pack()

        btns = tk.Frame(self.top)
        btns.pack(fill="x", pady=12)
        ttk.Button(btns, text="Cancelar", command=self._cancel).pack(side="right", padx=6)
        ttk.Button(btns, text="Aceptar", command=self._accept).pack(side="right")

        self.top.protocol("WM_DELETE_WINDOW", self._cancel)

    def _accept(self):
        label = (self.var_label.get() or "").strip()
        if not label:
            if messagebox:
                messagebox.showerror("Datos inválidos", "Especifica una etiqueta")
            return
        try:
            price = float(self.var_price.get().replace(",", ""))
        except Exception:
            if messagebox:
                messagebox.showerror("Datos inválidos", "Precio no válido")
            return
        found = False
        for choice in self.choices:
            if choice.get("label") == label:
                choice["price"] = price
                found = True
                break
        if not found:
            self.choices.append({"label": label, "price": price})
        self.on_change()
        self.top.destroy()

    def _cancel(self):
        self.top.destroy()
