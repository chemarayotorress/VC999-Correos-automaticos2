from __future__ import annotations

import io
import logging
import os
from pathlib import Path
from typing import Any, Dict, List, Optional, Union

from fastapi import FastAPI, Header, HTTPException
from fastapi.responses import StreamingResponse
from pydantic import BaseModel, ConfigDict, Field

from cotizador_backend import generar_cotizacion_desde_json
from template_resolver import list_available_models, normalize_model, resolve_template_path
from catalog_sync import sync_catalog


logger = logging.getLogger(__name__)
app = FastAPI(title="VC999 Packaging+ API")


class CotizacionRequest(BaseModel):
    model_config = ConfigDict(extra="allow")

    # Nuevo payload (n8n)
    class SelectionItem(BaseModel):
        step: str
        value: str
        price: float = 0

    class Customer(BaseModel):
        name: str
        email: str

    machine: Optional[str] = None
    basePrice: Optional[float] = Field(default=None, description="Precio base recibido desde n8n")
    totalPrice: Optional[float] = Field(default=None, description="Precio total final recibido desde n8n")
    selections: Optional[Union[List[SelectionItem], Dict[str, Any]]] = None
    customer: Optional[Union[Customer, str]] = None

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


@app.post("/sync-catalog")
def force_sync_catalog(x_vc999_token: Optional[str] = Header(default=None, alias="X-VC999-TOKEN")) -> Dict[str, Any]:
    expected_token = os.getenv("CATALOG_SYNC_TOKEN", "").strip()
    if not expected_token or x_vc999_token != expected_token:
        raise HTTPException(status_code=401, detail="Unauthorized")

    result = sync_catalog(force=True, persist_cache=True)
    if not result.get("ok"):
        raise HTTPException(status_code=503, detail=result)
    return result


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
    customer = datos.get("customer")
    if isinstance(customer, dict):
        if customer.get("name") and not datos.get("nombre_cliente"):
            datos["nombre_cliente"] = str(customer.get("name")).strip()
        if customer.get("email") and not datos.get("email"):
            datos["email"] = str(customer.get("email")).strip()
    elif customer and not datos.get("nombre_cliente"):
        datos["nombre_cliente"] = str(customer).strip()

    if datos.get("basePrice") is not None and datos.get("precio_cambiado") is None:
        datos["precio_cambiado"] = datos["basePrice"]

    selections = datos.get("selections")
    if isinstance(selections, list):
        normalized_selections: Dict[str, Any] = {}
        for item in selections:
            if not isinstance(item, dict):
                continue
            step = str(item.get("step") or "").strip()
            if not step:
                continue
            normalized_selections[step] = item.get("value")
        datos["selections"] = normalized_selections
        selections = normalized_selections

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
