"""Configuración editable del sistema de ventas.

Ajusta aquí los planes reales que vendes y las reglas de priorización.
"""

from pathlib import Path

RUTA_DB = Path(__file__).resolve().parent.parent / "data" / "ventas.db"

# Cuántas veces se intenta contactar un número que no contesta antes de cerrarlo.
MAX_INTENTOS_SIN_RESPUESTA = 3

# Planes sugeridos según el monto de recarga (CLP).
# Cada tupla: (recarga mínima, nombre del plan). Se elige el tramo más alto que cumpla.
# REEMPLAZA estos nombres por los planes reales de tu compañía.
PLANES_POR_RECARGA = [
    (0, "Plan Básico"),
    (5000, "Plan Medio"),
    (10000, "Plan Full"),
    (20000, "Plan Ilimitado"),
]

# Puntos extra de prioridad por subsegmento (los demás suman 0).
BONO_SUBSEGMENTO = {
    "PERSONA NUEVA": 10,
}

# Resultados que el vendedor puede registrar en cada gestión -> estado en que queda el cliente.
RESULTADOS = {
    "No contesta": "PENDIENTE",
    "Volver a llamar": "VOLVER_A_LLAMAR",
    "Interesado": "INTERESADO",
    "Venta cerrada": "VENTA",
    "No interesado": "NO_INTERESADO",
    "Número equivocado / no existe": "INVALIDO",
    "Pide no ser contactado": "NO_CONTACTAR",
}

# Estados en los que el cliente ya no aparece en la lista de trabajo.
ESTADOS_CERRADOS = {"VENTA", "NO_INTERESADO", "INVALIDO", "NO_CONTACTAR", "AGOTADO"}
