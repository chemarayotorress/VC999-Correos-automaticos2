"""
CLI para ejecutar el cotizador Packaging+ desde terminal o n8n (Execute Command).
"""
from __future__ import annotations

import argparse
import sys
from typing import Any

from cotizador_backend import generar_cotizacion_backend


def normaliza_bool(valor: str) -> bool:
    truthy = {"si", "sí", "true", "1", "on", "yes"}
    falsy = {"no", "false", "0", "off"}
    v = str(valor).strip().lower()
    if v in truthy:
        return True
    if v in falsy:
        return False
    raise ValueError(f"Valor booleano no reconocido: {valor}")


def _add_override(overrides: dict, key: str, value: Any) -> None:
    if value is None:
        return
    if isinstance(value, str) and value.strip() == "":
        return
    overrides[key] = value


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Generador de cotizaciones Packaging+ desde CLI")
    parser.add_argument("--modelo", required=True, help="Modelo/plantilla (ej. CM860)")
    parser.add_argument("--cliente", required=True, help="Nombre del cliente")
    parser.add_argument("--salida_pdf", required=True, help="Ruta del PDF de salida")
    parser.add_argument("--salida_word", help="Ruta opcional del DOCX de salida")

    parser.add_argument("--asesor")
    parser.add_argument("--fecha")
    parser.add_argument("--validez_dias", type=int)
    parser.add_argument("--moneda")
    parser.add_argument("--flete_texto")
    parser.add_argument("--flete_monto", type=float)
    parser.add_argument("--notas")

    parser.add_argument("--voltaje")
    parser.add_argument("--altura_tapa")
    parser.add_argument("--operacion")
    parser.add_argument("--opcion_bomba")
    parser.add_argument("--sistema_biactivo")
    parser.add_argument("--descarga_gas")
    parser.add_argument("--aire_positivo")
    parser.add_argument("--kit_muestras")

    parser.add_argument("--contrato1_porcentaje", type=int)
    parser.add_argument("--contrato1_condicion")
    parser.add_argument("--contrato2_porcentaje", type=int)
    parser.add_argument("--contrato2_condicion")
    parser.add_argument("--contrato3_porcentaje", type=int)
    parser.add_argument("--contrato3_condicion")
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    opciones_overrides: dict = {}

    for key in (
        "voltaje",
        "altura_tapa",
        "operacion",
        "opcion_bomba",
        "kit_muestras",
    ):
        _add_override(opciones_overrides, key, getattr(args, key))

    for key in ("sistema_biactivo", "descarga_gas", "aire_positivo"):
        val = getattr(args, key)
        if val is not None:
            opciones_overrides[key] = normaliza_bool(val)

    for idx in range(1, 4):
        _add_override(opciones_overrides, f"contrato{idx}_porcentaje", getattr(args, f"contrato{idx}_porcentaje"))
        _add_override(opciones_overrides, f"contrato{idx}_condicion", getattr(args, f"contrato{idx}_condicion"))

    try:
        resultado = generar_cotizacion_backend(
            modelo=args.modelo,
            cliente=args.cliente,
            asesor=args.asesor,
            fecha=args.fecha,
            validez_dias=args.validez_dias,
            moneda=args.moneda,
            flete_texto=args.flete_texto,
            flete_monto=args.flete_monto,
            notas=args.notas,
            opciones_overrides=opciones_overrides,
            ruta_salida_word=args.salida_word,
            ruta_salida_pdf=args.salida_pdf,
        )
        msg = "COTIZACION_OK;" + ";".join(
            [
                f"MODELO={resultado.get('modelo','')}",
                f"CLIENTE={resultado.get('cliente','')}",
                f"WORD={resultado.get('ruta_word','')}",
                f"PDF={resultado.get('ruta_pdf','')}",
                f"TOTAL={resultado.get('total','')}",
                f"MONEDA={resultado.get('moneda','')}",
            ]
        )
        print(msg)
        sys.exit(0)
    except Exception as exc:  # pragma: no cover - envolvente CLI
        print(f"COTIZACION_ERROR;mensaje={exc}")
        sys.exit(1)


if __name__ == "__main__":
    main()

