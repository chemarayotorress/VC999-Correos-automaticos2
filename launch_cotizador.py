"""Asistente para preparar e iniciar el cotizador sin generar ejecutables."""
from __future__ import annotations

import os
import subprocess
import textwrap
from pathlib import Path
import venv

ROOT = Path(__file__).resolve().parent
VENV_DIR = ROOT / ".venv"
REQUIREMENTS_FILE = ROOT / "requirements.txt"
MAIN_SCRIPT = ROOT / "Cotizador2_FINAL_MATERIALS_FLETE_ONLY_FIX_ROBUST2_BACKUP_SEARCH_OPERACION.py"
ENV_FLAG = "VC999_ENV_READY"


def message(text: str) -> None:
    print(textwrap.dedent(text))


def ensure_venv() -> None:
    if not VENV_DIR.exists():
        message("\nCreando entorno virtual local (.venv)...")
        venv.EnvBuilder(with_pip=True).create(VENV_DIR)


def python_from_venv() -> Path:
    if os.name == "nt":
        return VENV_DIR / "Scripts" / "python.exe"
    return VENV_DIR / "bin" / "python"


def install_dependencies(python_path: Path) -> None:
    message("\nActualizando pip dentro del entorno virtual...")
    subprocess.run([str(python_path), "-m", "pip", "install", "--upgrade", "pip"], check=True)

    if REQUIREMENTS_FILE.exists():
        message("\nInstalando dependencias del proyecto (requirements.txt)...")
        subprocess.run(
            [str(python_path), "-m", "pip", "install", "--upgrade", "-r", str(REQUIREMENTS_FILE)],
            check=True,
        )
    else:
        message("\nNo se encontró requirements.txt; instalando dependencias básicas...")
        subprocess.run(
            [
                str(python_path),
                "-m",
                "pip",
                "install",
                "--upgrade",
                "flask>=2.2,<3.0",
                "python-docx>=0.8.11",
                "docx2pdf>=0.1.8",
            ],
            check=True,
        )


def launch_app(python_path: Path) -> int:
    if not MAIN_SCRIPT.exists():
        raise SystemExit(
            "No se encontró el archivo principal del cotizador. "
            "Comprueba que copiaste toda la carpeta del proyecto."
        )

    message("\nIniciando el cotizador...")
    env = os.environ.copy()
    env[ENV_FLAG] = "1"
    process = subprocess.run([str(python_path), str(MAIN_SCRIPT)], env=env)
    return process.returncode


def main() -> None:
    message(
        """
        ========== Asistente de lanzamiento del VC999 Packaging+ Cotizador ==========

        Este asistente prepara automáticamente el entorno (sin crear ejecutables) e
        inicia la aplicación original en cuanto todo está listo.
        """
    )

    ensure_venv()
    python_path = python_from_venv()
    if not python_path.exists():
        raise SystemExit(
            "No se encontró el intérprete de Python dentro de .venv. "
            "Verifica que tengas permisos para crear carpetas en esta ubicación."
        )
    install_dependencies(python_path)

    return_code = launch_app(python_path)
    if return_code == 0:
        message(
            "\nEl cotizador se cerró correctamente. "
            "Puedes volver a ejecutar este archivo cuando necesites abrirlo de nuevo."
        )
    else:
        message(
            f"\nEl cotizador terminó con código {return_code}. "
            "Revisa los mensajes anteriores para identificar el problema."
        )

    input("\nPresiona Enter para cerrar esta ventana...")


if __name__ == "__main__":
    try:
        main()
    except subprocess.CalledProcessError as exc:
        message("\nOcurrió un error al instalar dependencias:")
        message(str(exc))
        input("\nPresiona Enter para cerrar esta ventana...")
        raise SystemExit(1)
    except Exception as exc:  # pylint: disable=broad-except
        message("\nOcurrió un error inesperado:")
        message(str(exc))
        input("\nPresiona Enter para cerrar esta ventana...")
        raise SystemExit(1)
