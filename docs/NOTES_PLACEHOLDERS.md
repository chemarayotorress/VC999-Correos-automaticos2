# Notas de placeholders Telegram/n8n -> DOCX/PDF

## Keys aceptadas en el backend
- Diccionario de placeholders: `placeholders`, `reemplazos`, `specs`, `vars`.
- Selecciones: `selections` como `dict` o `list`.
  - Si es lista, cada item puede usar `id`, `key`, `slug`, `step` o `name` como llave.
  - Para el valor visible del PDF se prioriza `label`, luego `value` (si `value` es objeto, se intenta `label/text/name/title/value`).
- El endpoint construye `replacements` normalizando keys a `snake_case` minúscula sin acentos.

## Prueba local reproducible
```bash
python scripts/test_replacements_e2e.py
```

La prueba:
1. Envía un payload con valores distintivos (`TEST_VOLT_123`, `TEST_OP_456`, `TEST_PUMP_789`).
2. Genera DOCX/PDF con el flujo real de backend.
3. Verifica que el DOCX final contiene esos valores.
4. Verifica que no queden tokens `{{...}}`.

## Warnings de placeholders no reemplazados
- Si quedan placeholders sin resolver en la plantilla, se registra warning en logs:
  - `Quedaron placeholders sin reemplazar en la plantilla: ...`
- El proceso ya no falla silenciosamente: deja evidencia explícita en el log para depuración del mapeo.
