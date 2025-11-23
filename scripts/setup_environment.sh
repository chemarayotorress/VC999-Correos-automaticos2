#!/usr/bin/env bash
set -euo pipefail

PROJECT_ROOT="$(cd "$(dirname "${BASH_SOURCE[0]}")"/.. && pwd)"

if [ ! -f "$PROJECT_ROOT/requirements.txt" ]; then
  echo "No se encontró requirements.txt en $PROJECT_ROOT" >&2
  exit 1
fi

if [ ! -d "$PROJECT_ROOT/.venv" ]; then
  echo "[1/3] Creando entorno virtual..."
  python3 -m venv "$PROJECT_ROOT/.venv"
fi

source "$PROJECT_ROOT/.venv/bin/activate"

echo "[2/3] Actualizando pip..."
python -m pip install --upgrade pip

echo "[3/3] Instalando dependencias del proyecto..."
pip install -r "$PROJECT_ROOT/requirements.txt"

echo
cat <<MSG
Entorno preparado correctamente. Para usarlo ejecuta:
  source .venv/bin/activate

Una vez activado puedes iniciar el backend con:
  python backend_service.py

o lanzar la interfaz del gestor de usuarios con:
  python manage_users_gui.py
MSG
