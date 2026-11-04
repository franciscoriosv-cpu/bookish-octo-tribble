"""Limpieza de la base, almacenamiento en SQLite y registro de gestiones."""

from __future__ import annotations

import re
import sqlite3
import unicodedata
from datetime import datetime
from pathlib import Path

import pandas as pd

from . import config

# ---------------------------------------------------------------------------
# Limpieza
# ---------------------------------------------------------------------------

# Nombre interno -> palabras con las que puede venir el encabezado en el Excel.
ALIAS_COLUMNAS = {
    "pais": ["PAIS"],
    "segmento": ["SEGMENTO"],
    "subsegmento": ["SUBSEGMENTO"],
    "producto": ["DESC_PRODUCTO", "PRODUCTO"],
    "monto_recarga": ["MONTO_RECARGA", "RECARGA", "MONTO"],
    "vendedor": ["VENDEDOR", "EJECUTIVO"],
    "telefono": ["CELULAR", "TELEFONO", "FONO", "MOVIL", "MSISDN", "NUMERO", "LINEA", "COMPANIA"],
}

PATRON_MOVIL_CHILE = re.compile(r"^569\d{8}$")


def _normalizar_texto(texto: str) -> str:
    """Mayúsculas, sin tildes ni espacios sobrantes: 'Compañía ' -> 'COMPANIA'."""
    sin_tildes = unicodedata.normalize("NFKD", str(texto)).encode("ascii", "ignore").decode()
    return re.sub(r"\s+", " ", sin_tildes).strip().upper()


def normalizar_telefono(valor) -> str | None:
    """Devuelve el móvil chileno como '569XXXXXXXX' o None si no es válido."""
    if valor is None or (isinstance(valor, float) and pd.isna(valor)):
        return None
    texto = str(valor).strip()
    if texto.endswith(".0"):  # Excel a veces lo lee como número decimal
        texto = texto[:-2]
    digitos = re.sub(r"\D", "", texto)
    if len(digitos) == 8:  # sin el 9 inicial (formato antiguo)
        digitos = "569" + digitos
    elif len(digitos) == 9 and digitos.startswith("9"):
        digitos = "56" + digitos
    return digitos if PATRON_MOVIL_CHILE.match(digitos) else None


def plan_sugerido(monto_recarga: float) -> str:
    elegido = config.PLANES_POR_RECARGA[0][1]
    for minimo, plan in config.PLANES_POR_RECARGA:
        if monto_recarga >= minimo:
            elegido = plan
    return elegido


def calcular_prioridad(monto_recarga: float, subsegmento: str) -> int:
    """Mayor recarga = más poder de compra = se llama primero."""
    base = min(int(monto_recarga // 1000), 50)
    return base + config.BONO_SUBSEGMENTO.get(_normalizar_texto(subsegmento), 0)


def _mapear_columnas(df: pd.DataFrame) -> dict[str, str]:
    """Encuentra qué columna del Excel corresponde a cada dato."""
    mapa: dict[str, str] = {}
    normalizadas = {col: _normalizar_texto(col) for col in df.columns}
    # 'SUBSEGMENTO' contiene 'SEGMENTO': se prueban primero las coincidencias exactas.
    for interno, alias in ALIAS_COLUMNAS.items():
        for col, norm in normalizadas.items():
            if col not in mapa.values() and norm in alias:
                mapa[interno] = col
                break
    for interno, alias in ALIAS_COLUMNAS.items():
        if interno in mapa:
            continue
        for col, norm in normalizadas.items():
            if col not in mapa.values() and any(a in norm for a in alias):
                mapa[interno] = col
                break

    if "telefono" not in mapa:
        # Último recurso: la columna donde la mayoría de los valores parecen móviles.
        mejor, mejor_ratio = None, 0.0
        for col in df.columns:
            if col in mapa.values():
                continue
            muestra = df[col].dropna().head(200)
            if muestra.empty:
                continue
            ratio = muestra.map(normalizar_telefono).notna().mean()
            if ratio > mejor_ratio:
                mejor, mejor_ratio = col, ratio
        if mejor is not None and mejor_ratio >= 0.8:
            mapa["telefono"] = mejor
    if "telefono" not in mapa:
        raise ValueError("No se encontró la columna de teléfono en el archivo.")
    return mapa


def leer_archivo(origen, nombre: str | None = None) -> pd.DataFrame:
    """Lee un .xlsx/.xls/.csv. Los teléfonos se leen como texto para no perder dígitos."""
    nombre = (nombre or str(origen)).lower()
    if nombre.endswith(".csv"):
        return pd.read_csv(origen, dtype=str, sep=None, engine="python", encoding="utf-8-sig")
    return pd.read_excel(origen, dtype=str)


def limpiar_base(df: pd.DataFrame) -> tuple[pd.DataFrame, dict]:
    """Normaliza la base. Devuelve (clientes limpios, resumen de la limpieza)."""
    mapa = _mapear_columnas(df)
    limpio = pd.DataFrame({interno: df[col] for interno, col in mapa.items()})
    for col in ALIAS_COLUMNAS:
        if col not in limpio:
            limpio[col] = ""

    total = len(limpio)
    limpio["telefono"] = limpio["telefono"].map(normalizar_telefono)
    invalidos = int(limpio["telefono"].isna().sum())
    limpio = limpio.dropna(subset=["telefono"])

    antes = len(limpio)
    limpio = limpio.drop_duplicates(subset="telefono", keep="first")
    duplicados = antes - len(limpio)

    for col in ["pais", "segmento", "subsegmento", "producto", "vendedor"]:
        limpio[col] = limpio[col].fillna("").astype(str).str.strip()
    limpio["vendedor"] = limpio["vendedor"].replace("", "SIN ASIGNAR")
    limpio["monto_recarga"] = (
        pd.to_numeric(limpio["monto_recarga"].astype(str).str.replace(r"[^\d.]", "", regex=True),
                      errors="coerce")
        .fillna(0)
    )
    limpio["plan_sugerido"] = limpio["monto_recarga"].map(plan_sugerido)
    limpio["prioridad"] = [
        calcular_prioridad(m, s) for m, s in zip(limpio["monto_recarga"], limpio["subsegmento"])
    ]

    resumen = {
        "filas_leidas": total,
        "telefonos_invalidos": invalidos,
        "duplicados": duplicados,
        "clientes_validos": len(limpio),
        "columnas_detectadas": mapa,
    }
    return limpio.reset_index(drop=True), resumen


# ---------------------------------------------------------------------------
# Base de datos
# ---------------------------------------------------------------------------

ESQUEMA = """
CREATE TABLE IF NOT EXISTS clientes (
    telefono TEXT PRIMARY KEY,
    pais TEXT, segmento TEXT, subsegmento TEXT, producto TEXT,
    monto_recarga REAL, vendedor TEXT, plan_sugerido TEXT, prioridad INTEGER,
    estado TEXT NOT NULL DEFAULT 'PENDIENTE',
    intentos INTEGER NOT NULL DEFAULT 0,
    ultima_gestion TEXT, proximo_contacto TEXT,
    fecha_carga TEXT
);
CREATE TABLE IF NOT EXISTS gestiones (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    telefono TEXT NOT NULL REFERENCES clientes(telefono),
    vendedor TEXT, fecha TEXT, resultado TEXT, plan_ofrecido TEXT, nota TEXT
);
CREATE TABLE IF NOT EXISTS exclusiones (
    telefono TEXT PRIMARY KEY,
    motivo TEXT, fecha TEXT
);
"""


def conectar(ruta: Path | str = config.RUTA_DB) -> sqlite3.Connection:
    if str(ruta) != ":memory:":
        Path(ruta).parent.mkdir(parents=True, exist_ok=True)
    con = sqlite3.connect(ruta, check_same_thread=False)
    con.row_factory = sqlite3.Row
    con.executescript(ESQUEMA)
    return con


def _ahora() -> str:
    return datetime.now().isoformat(timespec="seconds")


def guardar_clientes(con: sqlite3.Connection, clientes: pd.DataFrame) -> dict:
    """Inserta clientes nuevos y actualiza los datos de los existentes.

    Nunca pisa el estado ni el historial de un cliente que ya fue gestionado.
    """
    existentes = {r[0] for r in con.execute("SELECT telefono FROM clientes")}
    excluidos = {r[0] for r in con.execute("SELECT telefono FROM exclusiones")}
    nuevos = actualizados = 0
    ahora = _ahora()
    with con:
        for fila in clientes.itertuples(index=False):
            datos = (fila.pais, fila.segmento, fila.subsegmento, fila.producto,
                     float(fila.monto_recarga), fila.vendedor, fila.plan_sugerido,
                     int(fila.prioridad))
            if fila.telefono in existentes:
                con.execute(
                    """UPDATE clientes SET pais=?, segmento=?, subsegmento=?, producto=?,
                       monto_recarga=?, vendedor=?, plan_sugerido=?, prioridad=?
                       WHERE telefono=?""",
                    (*datos, fila.telefono),
                )
                actualizados += 1
            else:
                estado = "NO_CONTACTAR" if fila.telefono in excluidos else "PENDIENTE"
                con.execute(
                    """INSERT INTO clientes (pais, segmento, subsegmento, producto, monto_recarga,
                       vendedor, plan_sugerido, prioridad, telefono, estado, fecha_carga)
                       VALUES (?,?,?,?,?,?,?,?,?,?,?)""",
                    (*datos, fila.telefono, estado, ahora),
                )
                nuevos += 1
    return {"nuevos": nuevos, "actualizados": actualizados}


def cargar_exclusiones(con: sqlite3.Connection, telefonos, motivo: str = "Lista No Molestar") -> int:
    """Registra números que NO se deben contactar y bloquea a los clientes que ya estén cargados."""
    validos = {t for t in (normalizar_telefono(x) for x in telefonos) if t}
    ahora = _ahora()
    with con:
        con.executemany(
            "INSERT OR IGNORE INTO exclusiones (telefono, motivo, fecha) VALUES (?,?,?)",
            [(t, motivo, ahora) for t in validos],
        )
        con.executemany(
            "UPDATE clientes SET estado='NO_CONTACTAR' WHERE telefono=?", [(t,) for t in validos]
        )
    return len(validos)


def vendedores(con: sqlite3.Connection) -> list[str]:
    return [r[0] for r in con.execute("SELECT DISTINCT vendedor FROM clientes ORDER BY vendedor")]


def cola_de_trabajo(con: sqlite3.Connection, vendedor: str, ahora: str | None = None) -> pd.DataFrame:
    """Clientes que el vendedor debe contactar ahora, del más prioritario al menos."""
    ahora = ahora or _ahora()
    cerrados = tuple(config.ESTADOS_CERRADOS)
    marcas = ",".join("?" * len(cerrados))
    consulta = f"""
        SELECT * FROM clientes
        WHERE vendedor = ?
          AND estado NOT IN ({marcas})
          AND (proximo_contacto IS NULL OR proximo_contacto <= ?)
        ORDER BY (estado = 'INTERESADO') DESC, (estado = 'VOLVER_A_LLAMAR') DESC,
                 prioridad DESC, intentos ASC, telefono
    """
    return pd.read_sql_query(consulta, con, params=(vendedor, *cerrados, ahora))


def registrar_gestion(con: sqlite3.Connection, telefono: str, vendedor: str, resultado: str,
                      plan_ofrecido: str = "", nota: str = "",
                      proximo_contacto: str | None = None) -> str:
    """Guarda el resultado de un contacto y actualiza el estado del cliente. Devuelve el estado."""
    if resultado not in config.RESULTADOS:
        raise ValueError(f"Resultado desconocido: {resultado}")
    estado = config.RESULTADOS[resultado]
    ahora = _ahora()
    with con:
        intentos = con.execute(
            "SELECT intentos FROM clientes WHERE telefono=?", (telefono,)
        ).fetchone()[0] + 1
        if resultado == "No contesta" and intentos >= config.MAX_INTENTOS_SIN_RESPUESTA:
            estado = "AGOTADO"
        con.execute(
            "INSERT INTO gestiones (telefono, vendedor, fecha, resultado, plan_ofrecido, nota)"
            " VALUES (?,?,?,?,?,?)",
            (telefono, vendedor, ahora, resultado, plan_ofrecido, nota),
        )
        con.execute(
            "UPDATE clientes SET estado=?, intentos=?, ultima_gestion=?, proximo_contacto=?"
            " WHERE telefono=?",
            (estado, intentos, ahora, proximo_contacto, telefono),
        )
        if estado == "NO_CONTACTAR":
            con.execute(
                "INSERT OR IGNORE INTO exclusiones (telefono, motivo, fecha) VALUES (?,?,?)",
                (telefono, "Solicitud del cliente", ahora),
            )
    return estado


def resumen_por_vendedor(con: sqlite3.Connection) -> pd.DataFrame:
    return pd.read_sql_query(
        """
        SELECT vendedor,
               COUNT(*) AS asignados,
               SUM(intentos > 0) AS contactados,
               SUM(estado = 'INTERESADO') AS interesados,
               SUM(estado = 'VENTA') AS ventas,
               SUM(estado = 'NO_CONTACTAR') AS no_contactar,
               ROUND(100.0 * SUM(estado = 'VENTA') / NULLIF(SUM(intentos > 0), 0), 1)
                   AS conversion_pct
        FROM clientes GROUP BY vendedor ORDER BY ventas DESC, vendedor
        """,
        con,
    )


def ventas_por_plan(con: sqlite3.Connection) -> pd.DataFrame:
    return pd.read_sql_query(
        """
        SELECT COALESCE(NULLIF(plan_ofrecido, ''), 'Sin especificar') AS plan, COUNT(*) AS ventas
        FROM gestiones WHERE resultado = 'Venta cerrada'
        GROUP BY plan ORDER BY ventas DESC
        """,
        con,
    )
