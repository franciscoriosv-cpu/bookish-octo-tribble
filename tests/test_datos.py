import pandas as pd
import pytest

from ventas_planes import datos


def base_ejemplo():
    # Mismas columnas que la BBDD real (la del teléfono viene con el encabezado cortado).
    return pd.DataFrame({
        "PAIS": ["CHILE"] * 5,
        "SEGMENTO": ["Persona"] * 5,
        "SUBSEGMENTO": ["Persona Nueva", "Persona Nueva", "Persona", "Persona", "Persona"],
        "DESC_PRODUCTO": ["Kit Masivo Renueva"] * 5,
        "MONTO_RECARGA": ["2000", "2000", "12000", "5000", "1000"],
        "NUMERO CELULAR COMPAÑÍA": ["56992567990", "56991451504", "9 9123 4567",
                                     "56992567990", "123"],
        "VENDEDOR": ["VENDEDOR 1", "VENDEDOR 1", "VENDEDOR 2", "VENDEDOR 2", "VENDEDOR 2"],
    })


@pytest.mark.parametrize("entrada,esperado", [
    ("56992567990", "56992567990"),
    ("+56 9 9256 7990", "56992567990"),
    ("992567990", "56992567990"),
    ("92567990", "56992567990"),
    (56992567990.0, "56992567990"),
    ("56222222222", None),  # fijo, no móvil
    ("123", None),
    (None, None),
])
def test_normalizar_telefono(entrada, esperado):
    assert datos.normalizar_telefono(entrada) == esperado


def test_limpiar_base_detecta_columnas_y_limpia():
    clientes, resumen = datos.limpiar_base(base_ejemplo())
    assert resumen["columnas_detectadas"]["telefono"] == "NUMERO CELULAR COMPAÑÍA"
    assert resumen["columnas_detectadas"]["subsegmento"] == "SUBSEGMENTO"
    assert resumen["telefonos_invalidos"] == 1
    assert resumen["duplicados"] == 1
    assert len(clientes) == 3
    fila = clientes.set_index("telefono").loc["56991234567"]
    assert fila["plan_sugerido"] == "Plan Full"


def test_columna_telefono_sin_encabezado_reconocible():
    df = base_ejemplo().rename(columns={"NUMERO CELULAR COMPAÑÍA": "COL_X"})
    _, resumen = datos.limpiar_base(df.iloc[:3])
    assert resumen["columnas_detectadas"]["telefono"] == "COL_X"


def test_flujo_completo_gestion_y_exclusiones():
    con = datos.conectar(":memory:")
    datos.cargar_exclusiones(con, ["991451504"])
    clientes, _ = datos.limpiar_base(base_ejemplo())
    assert datos.guardar_clientes(con, clientes) == {"nuevos": 3, "actualizados": 0}

    # El número excluido no aparece en la cola del vendedor.
    cola = datos.cola_de_trabajo(con, "VENDEDOR 1")
    assert list(cola["telefono"]) == ["56992567990"]

    # 3 "no contesta" seguidos cierran el contacto.
    for _ in range(3):
        estado = datos.registrar_gestion(con, "56992567990", "VENDEDOR 1", "No contesta")
    assert estado == "AGOTADO"
    assert datos.cola_de_trabajo(con, "VENDEDOR 1").empty

    # Re-importar no pisa el estado.
    datos.guardar_clientes(con, clientes)
    assert datos.cola_de_trabajo(con, "VENDEDOR 1").empty

    # Pedir no ser contactado lo agrega a exclusiones.
    datos.registrar_gestion(con, "56991234567", "VENDEDOR 2", "Pide no ser contactado")
    assert con.execute("SELECT COUNT(*) FROM exclusiones").fetchone()[0] == 2

    datos.registrar_gestion(con, "56992567990", "VENDEDOR 1", "Venta cerrada", "Plan Básico")
    resumen = datos.resumen_por_vendedor(con).set_index("vendedor")
    assert resumen.loc["VENDEDOR 1", "ventas"] == 1


def test_volver_a_llamar_respeta_fecha():
    con = datos.conectar(":memory:")
    clientes, _ = datos.limpiar_base(base_ejemplo())
    datos.guardar_clientes(con, clientes)
    datos.registrar_gestion(con, "56991234567", "VENDEDOR 2", "Volver a llamar",
                            proximo_contacto="2099-01-01T10:00:00")
    assert datos.cola_de_trabajo(con, "VENDEDOR 2").empty
    assert len(datos.cola_de_trabajo(con, "VENDEDOR 2", ahora="2099-01-02T00:00:00")) == 1
