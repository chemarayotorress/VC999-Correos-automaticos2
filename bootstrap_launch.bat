@echo off
setlocal EnableExtensions EnableDelayedExpansion

rem Determine the directory of this script to locate launch_cotizador.py
set "SCRIPT_DIR=%~dp0"

rem Try to locate an existing Python interpreter
call :find_python
if defined PY_CMD goto :launch

echo [INFO] No se detecto una instalacion de Python. Se intentara instalar Python 3.11 para el usuario actual.
call :install_python
if errorlevel 1 goto :fail_install

rem Reintentar deteccion despues de instalar
call :find_python
if not defined PY_CMD goto :fail_not_found

:launch
if /I "%PY_CMD%"=="py" (
    set "LAUNCH_CMD=py -3"
) else (
    set "LAUNCH_CMD=%PY_CMD%"
)

echo.
echo [INFO] Preparando dependencias de Python...
call %LAUNCH_CMD% -m ensurepip --upgrade >nul 2>&1
if errorlevel 1 (
    echo [WARN] No se pudo forzar ensurepip; se continuara con pip si estuviera disponible.
)
call %LAUNCH_CMD% -m pip install --upgrade pip --user
if errorlevel 1 goto :fail_deps

if exist "%SCRIPT_DIR%requirements.txt" (
    echo [INFO] Instalando dependencias desde requirements.txt...
    call %LAUNCH_CMD% -m pip install --user --no-warn-script-location -r "%SCRIPT_DIR%requirements.txt"
    if errorlevel 1 goto :fail_deps
) else (
    echo [WARN] No se encontro requirements.txt en %SCRIPT_DIR%.
)

echo [INFO] Lanzando el asistente del cotizador con %LAUNCH_CMD%...
"%COMSPEC%" /c %LAUNCH_CMD% "%SCRIPT_DIR%launch_cotizador.py"
set "EXIT_CODE=%ERRORLEVEL%"

if not "%EXIT_CODE%"=="0" (
    echo [WARN] El cotizador finalizo con codigo de salida %EXIT_CODE%.
)
exit /b %EXIT_CODE%

:find_python
set "PY_CMD="
for /f "delims=" %%i in ('where python 2^>nul') do (
    set "PY_CMD=python"
    goto :eof
)
for /f "delims=" %%i in ('where py 2^>nul') do (
    set "PY_CMD=py"
    goto :eof
)
:eof
exit /b 0

:install_python
set "INSTALLER_URL=https://www.python.org/ftp/python/3.11.6/python-3.11.6-amd64.exe"
set "INSTALLER_PATH=%TEMP%\python-installer.exe"

echo [INFO] Descargando Python desde %INSTALLER_URL%
powershell -NoProfile -ExecutionPolicy Bypass -Command ^
    "try { Invoke-WebRequest -UseBasicParsing -Uri '%INSTALLER_URL%' -OutFile '%INSTALLER_PATH%' -ErrorAction Stop } catch { Write-Error $_; exit 1 }"
if errorlevel 1 (
    echo [ERROR] No se pudo descargar el instalador de Python.
    exit /b 1
)

echo [INFO] Ejecutando instalador silencioso de Python...
"%INSTALLER_PATH%" /quiet InstallAllUsers=0 PrependPath=1 Include_test=0 SimpleInstall=1
if errorlevel 1 (
    echo [ERROR] La instalacion de Python fallo. Prueba a ejecutar manualmente %INSTALLER_PATH%.
    exit /b 1
)

echo [INFO] Python se instalo correctamente.
exit /b 0

:fail_install
echo [ERROR] No se pudo preparar Python automaticamente. Instala Python 3.11 manualmente y vuelve a intentarlo.
exit /b 1

:fail_deps
echo [ERROR] La instalacion de dependencias de Python fallo. Revisa la conexion a internet e intenta nuevamente.
exit /b 1

:fail_not_found
echo [ERROR] Python sigue sin estar disponible tras la instalacion automatica. Verifica los permisos e intenta de nuevo.
exit /b 1
