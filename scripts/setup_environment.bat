@echo off
setlocal ENABLEDELAYEDEXPANSION

REM Ruta del repositorio (carpeta donde está este script)
set SCRIPT_DIR=%~dp0
set PROJECT_ROOT=%SCRIPT_DIR%..

if not exist "%PROJECT_ROOT%\requirements.txt" (
    echo No se encontró requirements.txt en %PROJECT_ROOT%.
    exit /b 1
)

if not exist "%PROJECT_ROOT%\.venv" (
    echo [1/3] Creando entorno virtual...
    python -m venv "%PROJECT_ROOT%\.venv"
    if errorlevel 1 (
        echo Error al crear el entorno virtual.
        exit /b 1
    )
) else (
    echo Reutilizando entorno virtual existente en .venv
)

call "%PROJECT_ROOT%\.venv\Scripts\activate.bat"
if errorlevel 1 (
    echo No se pudo activar el entorno virtual.
    exit /b 1
)

echo [2/3] Actualizando pip...
python -m pip install --upgrade pip
if errorlevel 1 (
    echo No se pudo actualizar pip.
    exit /b 1
)

echo [3/3] Instalando dependencias del proyecto...
pip install -r "%PROJECT_ROOT%\requirements.txt"
if errorlevel 1 (
    echo La instalación de dependencias falló.
    exit /b 1
)

echo.
echo Entorno preparado correctamente. Para usarlo ejecuta:
echo   call .venv\Scripts\activate.bat

echo Una vez activado puedes iniciar el backend con:
echo   python backend_service.py

echo o lanzar la interfaz del gestor de usuarios con:
echo   python manage_users_gui.py

echo.
endlocal
