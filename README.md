# VC999 App Utilities

## Gestor de usuarios

Se añadió la herramienta `manage_users_gui.py` para gestionar usuarios de la base de datos del backend. La utilidad comprueba
automáticamente que el esquema esté creado y genera el usuario administrador por defecto si aún no existe.

### Requisitos
- Python 3.8 o superior.
- Soporte de Tkinter instalado en el sistema (incluido por defecto en la mayoría de distribuciones de Python para Windows y macOS; en algunas distribuciones de Linux puede requerir el paquete `python3-tk`).

### Uso
1. Asegúrate de que la base de datos `backend.db` esté inicializada (se crea automáticamente al ejecutar el backend o la herramienta por primera vez).
2. Ejecuta la interfaz gráfica con:
   ```bash
   python manage_users_gui.py
   ```
3. Introduce el correo del usuario, la contraseña y la licencia.
4. Pulsa **Crear/Actualizar usuario** para guardar los cambios en la tabla `users`.
5. Pulsa **Revocar sesiones** para eliminar los tokens activos del usuario indicado (o todos, si se deja el correo en blanco).

La aplicación usa internamente `backend_service.DB_PATH` y `_hash_password` para garantizar que comparte la misma base de datos y algoritmo de hashing que el backend principal.

## Guía de despliegue del backend

Sigue estos pasos para poner en marcha el servicio del backend en un servidor dedicado (Linux o Windows) y mantenerlo asegurado:

1. **Preparar el servidor.** Provisiona una máquina con Windows o Linux que disponga de Python 3.8 o superior instalado y copia el contenido de este repositorio en una carpeta estable, por ejemplo `/opt/vc999` en Linux o `C:\vc999` en Windows.
2. **Crear el entorno virtual.** Desde dicha carpeta crea un entorno virtual de Python y activa la nueva instalación. Una vez activa ejecuta `pip install flask` para instalar la única dependencia externa del backend.
3. **Registrar el servicio.** Configura el sistema para ejecutar `python /opt/vc999/backend_service.py` (ajusta la ruta si usas Windows) durante el arranque. En Linux se recomienda un servicio `systemd`; en Windows puedes emplear el Programador de tareas o la utilidad de Servicios. La primera ejecución generará el fichero `backend.db` junto con el usuario `admin@vc999.com`.
4. **Asegurar la red.** Restringe el acceso al puerto `5000` del servidor mediante el firewall, limitándolo a tu red interna o VPN. Si requieres HTTPS expón el servicio detrás de un proxy inverso que proporcione el cifrado TLS.
5. **Gestión de usuarios.** Cuando necesites dar de alta o revocar usuarios conecta remotamente con el servidor y ejecuta `python manage_users_gui.py`. La interfaz gráfica operará directamente sobre la base de datos `backend.db` centralizada.

## Cómo ejecutar el cotizador sin generar ejecutables

El cotizador se distribuye ahora únicamente como código fuente para evitar falsos positivos de antivirus. Para abrirlo en cualquier computadora sigue estos pasos:

1. Copia la carpeta completa del proyecto (plantillas `.docx`, imágenes, archivos `.json`, etc.) al disco local del equipo destino.
2. Si estás en **Windows** y no sabes si tienes Python instalado, ejecuta primero `bootstrap_launch.bat`. El script descargará Python 3.11 en caso de que falte y luego abrirá el cotizador automáticamente.
3. En cualquier sistema con Python 3.8 o superior disponible, ejecuta `launch_cotizador.py` (doble clic o `python launch_cotizador.py`). El asistente:
   - crea o reutiliza el entorno virtual local `.venv/`;
   - instala las dependencias declaradas en `requirements.txt` (`Flask`, `python-docx`, `docx2pdf`);
   - lanza automáticamente `Cotizador2_FINAL_MATERIALS_FLETE_ONLY_FIX_ROBUST2_BACKUP_SEARCH_OPERACION.py` usando ese entorno preparado.
4. Cuando cierres la ventana del cotizador, la consola mostrará un resumen y te pedirá presionar **Enter** para finalizar.

### Alternativa manual

Si prefieres preparar el entorno manualmente (por ejemplo para usar el backend o el gestor de usuarios), puedes ejecutar los scripts de la carpeta `scripts/` y después iniciar el cotizador a mano:

- En **Windows**: `scripts\setup_environment.bat` seguido de `call .venv\Scripts\activate.bat` y `python Cotizador2_FINAL_MATERIALS_FLETE_ONLY_FIX_ROBUST2_BACKUP_SEARCH_OPERACION.py`.
- En **macOS/Linux**: `./scripts/setup_environment.sh`, luego `source .venv/bin/activate` y `python Cotizador2_FINAL_MATERIALS_FLETE_ONLY_FIX_ROBUST2_BACKUP_SEARCH_OPERACION.py`.

Recuerda que la exportación a PDF mediante `docx2pdf` requiere Microsoft Word en Windows o LibreOffice en macOS/Linux.

## Distribución en otras computadoras

Si necesitas compartir la aplicación mediante una memoria USB u otro medio externo, revisa la guía [`docs/USB_DEPLOYMENT.md`](docs/USB_DEPLOYMENT.md). Encontrarás un checklist de archivos a copiar y scripts para automatizar la instalación de dependencias en equipos Windows, macOS o Linux.

## API FastAPI: prueba manual para n8n (`/generar-cotizacion`)

Inicia el servicio con:

```bash
python -m uvicorn fastapi_app:app --host 0.0.0.0 --port 8000 --reload
```

Prueba el payload nuevo (n8n) con `selections` como lista y `customer` como objeto:

```bash
curl -X POST http://127.0.0.1:8000/generar-cotizacion \
  -H "Content-Type: application/json" \
  -o cotizacion.pdf \
  -d '{
    "machine": "CM640.docx",
    "basePrice": 17995,
    "totalPrice": 17995,
    "selections": [
      {"step":"Voltage","value":"208V_3PH_60HZ","price":0},
      {"step":"GasFlush","value":"NO","price":0},
      {"step":"PositiveAirSealer","value":"NO","price":0}
    ],
    "customer": {"name":"Chema","email":"chema@example.com"}
  }'
```

Ejemplo equivalente en PowerShell:

```powershell
$body = @{
  machine = "CM640.docx"
  basePrice = 17995
  totalPrice = 17995
  selections = @(
    @{ step = "Voltage"; value = "208V_3PH_60HZ"; price = 0 },
    @{ step = "GasFlush"; value = "NO"; price = 0 },
    @{ step = "PositiveAirSealer"; value = "NO"; price = 0 }
  )
  customer = @{ name = "Chema"; email = "chema@example.com" }
} | ConvertTo-Json -Depth 5

Invoke-RestMethod -Method Post -Uri "http://127.0.0.1:8000/generar-cotizacion" -ContentType "application/json" -Body $body -OutFile "cotizacion.pdf"
```

Validaciones esperadas:
- HTTP 200 y generación de PDF (sin error 422).
- El mismo resultado al enviar `machine: "CM640"` o `machine: "CM640.docx"`.
- El payload legado (campos `modelo`, `nombre_cliente`, etc.) sigue siendo aceptado.


## Sincronización de catálogo desde Google Sheets (`/sync-catalog`)

El backend ahora puede usar Google Sheets como **source of truth** para precios base y opciones, con fallback automático a `machines.json` si falla Sheets.

### Variables de entorno

- `CATALOG_SYNC_TOKEN`: token requerido por el endpoint `POST /sync-catalog` (header `X-VC999-TOKEN`).
- `CATALOG_SYNC_TTL_SECONDS` (opcional, default `300`): TTL de caché para evitar recargas repetidas.
- `CATALOG_SYNC_TIMEOUT_SECONDS` (opcional, default `15`): timeout de lectura a Sheets.

#### Modo A: CSV export (sin credenciales)

- `GOOGLE_SHEET_ID=<id_del_sheet>`
- Compartir/publicar el documento para que las pestañas sean accesibles.
- Se leen estas URLs:
  - `https://docs.google.com/spreadsheets/d/<ID>/gviz/tq?tqx=out:csv&sheet=DB_Maquinas`
  - `https://docs.google.com/spreadsheets/d/<ID>/gviz/tq?tqx=out:csv&sheet=DB_Precios`

#### Modo B: Service Account (sheet privado, recomendado)

- `GOOGLE_SHEET_ID=<id_del_sheet>`
- `GOOGLE_APPLICATION_CREDENTIALS=<ruta_al_json_de_service_account>`
- Compartir el Sheet con el correo de la service account (permiso lector).

### Forzar refresh inmediato

```bash
curl -X POST http://127.0.0.1:8000/sync-catalog   -H "X-VC999-TOKEN: $CATALOG_SYNC_TOKEN"
```

Respuesta esperada cuando todo está correcto:

```json
{
  "ok": true,
  "source": "sheets",
  "updated_at": "2026-01-01T12:00:00+00:00",
  "items": 8
}
```

Si falla Google Sheets, el backend hace fallback a `machines.json` y deja log: `Fallback to local machines.json`.

### Integración n8n para ejecutar sync en cada `/start`

En el workflow versionado (`json3.json`) se agregó un nodo HTTP antes de mostrar el menú inicial:

1. Nodo: **Sync Catalog Backend** (HTTP Request)
   - Method: `POST`
   - URL: `https://<tu-ngrok>/sync-catalog`
   - Header: `X-VC999-TOKEN: {{$env.CATALOG_SYNC_TOKEN}}`
2. Nodo IF: **¿Sync catálogo OK?**
   - Condición: `{{$json.ok}}` es `true`
3. Rama `true` -> continuar con **Enviar Menú**.
4. Rama `false` -> enviar mensaje de error al usuario y no continuar el flujo.

Con esto, cada vez que Telegram mande `/start`, n8n fuerza refresh del catálogo en backend antes de continuar con cotización.
