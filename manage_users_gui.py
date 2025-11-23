"""Utility GUI to manage application users via the shared backend database."""
from __future__ import annotations

import sqlite3
import tkinter as tk
from tkinter import messagebox

from backend_service import DB_PATH, _ensure_admin, _hash_password, init_db


def _ensure_schema() -> bool:
    try:
        init_db()
        _ensure_admin()
    except sqlite3.Error as exc:
        messagebox.showerror(
            "Error",
            f"No se pudo inicializar la base de datos: {exc}",
        )
        return False
    return True


class UserManagerApp:
    """Simple Tkinter GUI to create/update users and revoke sessions."""

    def __init__(self, master: tk.Tk) -> None:
        self.master = master
        master.title("VC999 User Manager")
        master.resizable(False, False)

        if not _ensure_schema():
            master.destroy()
            raise SystemExit(1)

        self._build_form()

    def _build_form(self) -> None:
        padding = {"padx": 10, "pady": 5, "sticky": "w"}

        tk.Label(self.master, text="Correo del usuario").grid(row=0, column=0, **padding)
        self.email_entry = tk.Entry(self.master, width=40)
        self.email_entry.grid(row=0, column=1, **padding)

        tk.Label(self.master, text="Contraseña").grid(row=1, column=0, **padding)
        self.password_entry = tk.Entry(self.master, show="*", width=40)
        self.password_entry.grid(row=1, column=1, **padding)

        tk.Label(self.master, text="Licencia").grid(row=2, column=0, **padding)
        self.license_entry = tk.Entry(self.master, width=40)
        self.license_entry.grid(row=2, column=1, **padding)

        save_btn = tk.Button(self.master, text="Crear/Actualizar usuario", command=self.save_user)
        save_btn.grid(row=3, column=0, columnspan=2, pady=(10, 5))

        revoke_btn = tk.Button(self.master, text="Revocar sesiones", command=self.revoke_sessions)
        revoke_btn.grid(row=4, column=0, columnspan=2, pady=(0, 10))

    def save_user(self) -> None:
        username = (self.email_entry.get() or "").strip()
        password = self.password_entry.get() or ""
        license_key = (self.license_entry.get() or "").strip()

        if not username:
            messagebox.showerror("Error", "El correo es obligatorio.")
            return
        if not password:
            messagebox.showerror("Error", "La contraseña es obligatoria.")
            return
        if not license_key:
            messagebox.showerror("Error", "La licencia es obligatoria.")
            return

        hashed = _hash_password(password)

        try:
            with sqlite3.connect(DB_PATH) as conn:
                cur = conn.cursor()
                cur.execute(
                    "SELECT id FROM users WHERE lower(username)=?",
                    (username.lower(),),
                )
                row = cur.fetchone()
                if row:
                    cur.execute(
                        "UPDATE users SET username=?, password_hash=?, license_key=?, active=1 WHERE id=?",
                        (username, hashed, license_key, row[0]),
                    )
                    action = "actualizado"
                else:
                    cur.execute(
                        "INSERT INTO users(username, password_hash, license_key, active) VALUES (?,?,?,1)",
                        (username, hashed, license_key),
                    )
                    action = "creado"
                conn.commit()
        except sqlite3.Error as exc:
            messagebox.showerror("Error", f"No se pudo guardar el usuario: {exc}")
            return

        messagebox.showinfo("Éxito", f"Usuario {action} correctamente.")
        self.password_entry.delete(0, tk.END)

    def revoke_sessions(self) -> None:
        username = (self.email_entry.get() or "").strip()

        try:
            with sqlite3.connect(DB_PATH) as conn:
                cur = conn.cursor()
                if username:
                    cur.execute(
                        "DELETE FROM tokens WHERE user_id IN (SELECT id FROM users WHERE lower(username)=?)",
                        (username.lower(),),
                    )
                else:
                    cur.execute("DELETE FROM tokens")
                affected = cur.rowcount
                conn.commit()
        except sqlite3.Error as exc:
            messagebox.showerror("Error", f"No se pudieron revocar las sesiones: {exc}")
            return

        if username:
            message = f"Se revocaron {affected} sesiones para {username}."
        else:
            message = f"Se revocaron {affected} sesiones en total."
        messagebox.showinfo("Sesiones revocadas", message)


def main() -> None:
    root = tk.Tk()
    app = UserManagerApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
