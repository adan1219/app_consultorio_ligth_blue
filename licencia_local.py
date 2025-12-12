from __future__ import annotations

import json
from datetime import date
from pathlib import Path
from typing import Dict, Any


LICENCIA_PATH = Path(__file__).resolve().parent / "licencia.json"


def cargar_licencia() -> Dict[str, Any]:
    if LICENCIA_PATH.exists():
        with LICENCIA_PATH.open("r", encoding="utf-8") as archivo:
            return json.load(archivo)

    licencia_predeterminada: Dict[str, Any] = {
        "modo": "prueba",
        "fecha_inicio_prueba": date.today().isoformat(),
        "dias_prueba": 15,
    }
    guardar_licencia(licencia_predeterminada)
    return licencia_predeterminada


def guardar_licencia(licencia: Dict[str, Any]) -> None:
    with LICENCIA_PATH.open("w", encoding="utf-8") as archivo:
        json.dump(licencia, archivo, indent=4, ensure_ascii=False)


def evaluar_licencia() -> Dict[str, Any]:
    licencia = cargar_licencia()

    if licencia.get("modo") == "pago":
        return {
            "estado": "LICENCIA_ACTIVA",
            "dias_restantes": None,
            "datos": licencia,
        }

    fecha_inicio_str = licencia.get("fecha_inicio_prueba", date.today().isoformat())
    fecha_inicio = date.fromisoformat(fecha_inicio_str)
    dias_prueba = int(licencia.get("dias_prueba", 0))

    dias_transcurridos = (date.today() - fecha_inicio).days
    dias_restantes = dias_prueba - dias_transcurridos

    if dias_restantes >= 0:
        return {
            "estado": "PRUEBA_ACTIVA",
            "dias_restantes": dias_restantes,
            "datos": licencia,
        }

    return {
        "estado": "PRUEBA_VENCIDA",
        "dias_restantes": 0,
        "datos": licencia,
    }
