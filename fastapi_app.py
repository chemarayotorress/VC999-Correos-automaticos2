from __future__ import annotations

import io
import logging
from pathlib import Path
from typing import Any, Dict, Optional

from fastapi import FastAPI, HTTPException
from fastapi.responses import StreamingResponse
from pydantic import BaseModel, ConfigDict

from cotizador_backend import generar_cotizacion_desde_json
from template_resolver import list_available_models, normalize_model, resolve_template_path


logger = logging.getLogger(__name__)
app = FastAPI(title="VC999 Packaging+ API")


class CotizacionRequest(BaseModel):
    model_config = ConfigDict(extra="allow")

    # Nuevo payload (n8n)
    machine: Optional[str] = None
    basePrice: Optional[float] = None
    totalPrice: Optional[float] = None
    selections: Optional[Dict[str, Any]] = None
    customer: Optional[str] = None

    # Compatibilidad con payload anterior
    modelo: Optional[str] = None
    plantilla: Optional[str] = None
    asesor: Optional[str] = None
    inyeccion_gas: Optional[str] = None
    nombre_cliente: Optional[str] = None
    email: Optional[str] = None
    precio_cambiado: Optional[float] = None
    validez_dias: Optional[int] = None
    numero_tapa: Optional[int] = None
    altura_tapa: Optional[int] = None
    operacion: Optional[str] = None
    opcion_bomba: Optional[str] = None
    kit_muestras: Optional[str] = None
    sistema_biactivo: Optional[str] = None
    aire_positivo: Optional[str] = None
    kit_muestras_nit: Optional[str] = None
    notas: Optional[str] = None
    tipo_moneda: Optional[str] = None
    flete_texto: Optional[str] = None
    contrato1_porcentaje: Optional[int] = None
    contrato1_condicion: Optional[str] = None
    contrato2_porcentaje: Optional[int] = None
    contrato2_condicion: Optional[str] = None
    contrato3_porcentaje: Optional[int] = None
    contrato3_condicion: Optional[str] = None


@app.get("/")
def healthcheck() -> Dict[str, str]:
    return {"status": "ok"}


@app.post("/generar-cotizacion")
def generar_cotizacion(req: CotizacionRequest):
    datos = req.model_dump(exclude_none=True)

    model_raw = datos.get("machine") or datos.get("modelo") or datos.get("plantilla")
    model_normalized = normalize_model(model_raw)
    if not model_normalized:
        raise HTTPException(status_code=400, detail="Debe enviar 'machine', 'modelo' o 'plantilla'.")

    template_path = resolve_template_path(model_normalized)
    logger.info(
        "Cotizacion request model_raw=%r model_normalized=%r template_path=%r",
        model_raw,
        model_normalized,
        str(template_path) if template_path else None,
    )

    if template_path is None:
        available = ", ".join(list_available_models(limit=10))
        raise HTTPException(
            status_code=400,
            detail=f"Modelo/plantilla no disponible: {model_normalized}. Disponibles: {available}",
        )

    # Prioridad al payload nuevo de n8n sin romper compatibilidad.
    datos["modelo"] = model_normalized
    if datos.get("customer") and not datos.get("nombre_cliente"):
        datos["nombre_cliente"] = str(datos["customer"]).strip()

    selections = datos.get("selections")
    if isinstance(selections, dict):
        for key, value in selections.items():
            datos.setdefault(key, value)

    try:
        pdf_path_str = generar_cotizacion_desde_json(datos)
        pdf_path = Path(pdf_path_str)
        if not pdf_path.exists():
            raise FileNotFoundError(f"No se encontró el PDF en {pdf_path}")
    except (ValueError, FileNotFoundError) as e:
        raise HTTPException(status_code=400, detail=str(e))
    except HTTPException:
        raise
    except Exception as e:  # pragma: no cover - capa de servicio
        raise HTTPException(status_code=500, detail=str(e))

    with pdf_path.open("rb") as f:
        pdf_bytes = f.read()

    return StreamingResponse(
        io.BytesIO(pdf_bytes),
        media_type="application/pdf",
        headers={"Content-Disposition": f'attachment; filename="{pdf_path.name}"'},
    )


if __name__ == "__main__":
    import uvicorn

    uvicorn.run(app, host="0.0.0.0", port=8000)
