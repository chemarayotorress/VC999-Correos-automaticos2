# Guía para distribuir el cotizador mediante USB

Esta guía explica cómo preparar una memoria USB con todo lo necesario para que otras personas puedan ejecutar el **VC999 Packaging+ Cotizador** sin crear archivos ejecutables (`.exe`). Todo el flujo se basa en código Python y en un asistente que instala las dependencias automáticamente.

## 1. Preparar el contenido en tu computadora

1. Crea una carpeta temporal, por ejemplo `VC999_portable/`, y copia dentro **todo** el contenido del repositorio (archivos `.py`, plantillas `.docx`, imágenes, JSON y la carpeta `respaldos/`).
2. Asegúrate de incluir:
   - `bootstrap_launch.bat` (solo Windows; instala Python si falta y abre el cotizador).
   - `launch_cotizador.py` (asistente automático para instalar dependencias y abrir el cotizador).
   - `Cotizador2_FINAL_MATERIALS_FLETE_ONLY_FIX_ROBUST2_BACKUP_SEARCH_OPERACION.py` (aplicación principal).
   - `requirements.txt` y la carpeta `scripts/` (útiles si alguien prefiere una instalación manual).
   - Los archivos de soporte (`app_config.json`, `historial_*.json`, `machines.json`, plantillas `.docx`, imágenes, carpeta `respaldos/`, etc.).
3. Verifica que la estructura resulte similar a:
   ```text
   VC999_portable/
   ├── bootstrap_launch.bat
   ├── launch_cotizador.py
   ├── Cotizador2_FINAL_MATERIALS_FLETE_ONLY_FIX_ROBUST2_BACKUP_SEARCH_OPERACION.py
   ├── requirements.txt
   ├── scripts/
   │   ├── setup_environment.bat
   │   └── setup_environment.sh
   ├── app_config.json
   ├── advisors.json
   ├── asesores.json
   ├── asesores_materials.json
   ├── Cotizacion Materials.docx
   ├── historial_materials.json
   ├── historial_packaging.json
   ├── logo_materials.png
   ├── machines.json
   ├── vc999_logo.png
   ├── respaldos/
   └── ...
   ```
   > Añade cualquier otro recurso adicional que hayas incorporado al cotizador.
4. Copia la carpeta `VC999_portable/` a la memoria USB.

## 2. Pasos para quien recibe la memoria USB

### 2.1 Abrir el cotizador con el asistente automático (recomendado)

1. Copia la carpeta `VC999_portable/` desde la USB a una ubicación local del equipo (por ejemplo `C:\VC999_portable` o `~/VC999_portable`).
2. En **Windows**, si no sabes si el equipo tiene Python instalado, ejecuta `bootstrap_launch.bat`. El script descargará e insta
   lará Python 3.11 (modo usuario) si es necesario y después abrirá el cotizador automáticamente.
3. En cualquier sistema con Python 3.8 o superior disponible, ejecuta `launch_cotizador.py` con doble clic o desde una terminal:
   ```bash
   python launch_cotizador.py
   ```
   El asistente creará (o reutilizará) el entorno virtual `.venv/`, instalará las dependencias indicadas en `requirements.txt` y lanzará automáticamente el cotizador.
4. Una vez que cierres la ventana del cotizador, la consola mostrará un mensaje final y te pedirá presionar **Enter** para terminar.
5. Si necesitas generar PDFs con `docx2pdf`, asegúrate de tener Microsoft Word (Windows) o LibreOffice (macOS/Linux) instalado.

### 2.2 Preparación manual y ejecución directa (opcional)

Si prefieres controlar el proceso manualmente o ejecutar otros módulos (backend, gestor de usuarios), sigue estos pasos:

#### En Windows

1. Abre PowerShell o el Símbolo del sistema en la carpeta copiada.
2. Ejecuta el script automático para crear el entorno y las dependencias:
   ```bat
   scripts\setup_environment.bat
   ```
3. Activa el entorno e inicia el módulo que necesites:
   ```bat
   call .venv\Scripts\activate.bat
   python Cotizador2_FINAL_MATERIALS_FLETE_ONLY_FIX_ROBUST2_BACKUP_SEARCH_OPERACION.py
   ```
   o bien:
   ```bat
   python backend_service.py
   python manage_users_gui.py
   ```

#### En macOS o Linux

1. Abre una terminal en la carpeta copiada.
2. Ejecuta el script de preparación:
   ```bash
   ./scripts/setup_environment.sh
   ```
3. Activa el entorno y lanza el módulo deseado:
   ```bash
   source .venv/bin/activate
   python Cotizador2_FINAL_MATERIALS_FLETE_ONLY_FIX_ROBUST2_BACKUP_SEARCH_OPERACION.py
   ```
   o bien inicia el backend con `python backend_service.py`.

## 3. Consejos adicionales

- **Evita ejecutar el proyecto directamente desde la USB.** Cópialo siempre al disco local para que los antivirus no bloqueen la creación del entorno virtual.
- **Mantén la USB actualizada.** Cada vez que modifiques plantillas, archivos JSON o el código, vuelve a copiar la carpeta completa.
- **Documenta requisitos especiales.** Si la instalación depende de Microsoft Word o LibreOffice para la exportación a PDF, indícalo a los destinatarios.
- **No se generan binarios.** Todo se ejecuta desde Python, por lo que no hay hashes ni registros de integridad que compartir.

Con este flujo cualquier persona podrá instalar las dependencias necesarias y ejecutar el cotizador sin tener que compilar un `.exe` ni solicitar excepciones en el antivirus.
