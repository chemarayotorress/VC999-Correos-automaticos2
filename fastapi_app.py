from fastapi import FastAPI, HTTPException
from fastapi.responses import StreamingResponse
from pydantic import BaseModel
from pathlib import Path
import io
from typing import Optional

from cotizador_backend import generar_cotizacion_desde_json


app = FastAPI(title="VC999 Packaging+ API")


class CotizacionRequest(BaseModel):
    modelo: Optional[str] = None
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


@app.post("/generar-cotizacion")
def generar_cotizacion(req: CotizacionRequest):
    datos = req.dict()
    try:
        pdf_path_str = generar_cotizacion_desde_json(datos)
        pdf_path = Path(pdf_path_str)
        if not pdf_path.exists():
            raise FileNotFoundError(f"No se encontró el PDF en {pdf_path}")
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
