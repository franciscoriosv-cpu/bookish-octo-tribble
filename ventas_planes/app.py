"""App web para que los vendedores trabajen la base.

Ejecutar con:  streamlit run ventas_planes/app.py
"""

import sys
from datetime import datetime, time
from pathlib import Path

import streamlit as st

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from ventas_planes import config, datos  # noqa: E402

st.set_page_config(page_title="Ventas de planes", page_icon="📱", layout="wide")


@st.cache_resource
def obtener_conexion():
    return datos.conectar()


con = obtener_conexion()
pagina = st.sidebar.radio("Menú", ["Mi lista", "Cargar base", "Reportes"])


# ---------------------------------------------------------------------------
def pagina_mi_lista():
    st.title("📞 Mi lista de contactos")
    lista_vendedores = datos.vendedores(con)
    if not lista_vendedores:
        st.info("Todavía no hay clientes cargados. Ve a **Cargar base**.")
        return
    vendedor = st.sidebar.selectbox("Vendedor", lista_vendedores)
    cola = datos.cola_de_trabajo(con, vendedor)

    st.metric("Clientes por contactar ahora", len(cola))
    if cola.empty:
        st.success("¡No tienes contactos pendientes por ahora!")
        return

    cliente = cola.iloc[0]
    tel = cliente["telefono"]
    with st.container(border=True):
        c1, c2, c3 = st.columns(3)
        c1.subheader(f"+{tel}")
        c1.write(f"**Estado:** {cliente['estado']} · **Intentos:** {cliente['intentos']}")
        c2.write(f"**Segmento:** {cliente['segmento']} / {cliente['subsegmento']}")
        c2.write(f"**Producto actual:** {cliente['producto']}")
        c3.write(f"**Recarga:** ${cliente['monto_recarga']:,.0f}".replace(",", "."))
        c3.write(f"**Plan sugerido:** {cliente['plan_sugerido']}")
        st.markdown(f"[📞 Llamar](tel:+{tel}) &nbsp;&nbsp; [💬 Abrir WhatsApp](https://wa.me/{tel})")
        st.caption("Usa WhatsApp solo si el cliente aceptó ser contactado por ese medio.")

        historial = con.execute(
            "SELECT fecha, resultado, plan_ofrecido, nota FROM gestiones"
            " WHERE telefono=? ORDER BY fecha DESC", (tel,)
        ).fetchall()
        if historial:
            with st.expander(f"Historial ({len(historial)})"):
                st.table([dict(h) for h in historial])

    with st.form("gestion", clear_on_submit=True):
        resultado = st.radio("Resultado del contacto", list(config.RESULTADOS), horizontal=True)
        planes = [p for _, p in config.PLANES_POR_RECARGA]
        plan = st.selectbox("Plan ofrecido", planes, index=planes.index(cliente["plan_sugerido"]))
        nota = st.text_area("Nota")
        c1, c2 = st.columns(2)
        fecha = c1.date_input("Volver a llamar el (opcional)", value=None)
        hora = c2.time_input("Hora", value=time(10, 0))
        if st.form_submit_button("Guardar y pasar al siguiente", type="primary"):
            proximo = datetime.combine(fecha, hora).isoformat(timespec="seconds") if fecha else None
            estado = datos.registrar_gestion(con, tel, vendedor, resultado, plan, nota, proximo)
            st.toast(f"Guardado: +{tel} → {estado}")
            st.rerun()

    with st.expander("Ver toda mi cola"):
        st.dataframe(
            cola[["telefono", "estado", "subsegmento", "monto_recarga", "plan_sugerido",
                  "prioridad", "intentos", "proximo_contacto"]],
            hide_index=True, width="stretch",
        )


# ---------------------------------------------------------------------------
def pagina_cargar():
    st.title("📥 Cargar base")
    st.write("Sube el Excel o CSV de la BBDD (PAIS, SEGMENTO, SUBSEGMENTO, DESC_PRODUCTO, "
             "MONTO_RECARGA, número, VENDEDOR). Puedes volver a cargarla: los clientes ya "
             "gestionados conservan su estado e historial.")
    archivo = st.file_uploader("Base de clientes", type=["xlsx", "xls", "csv"])
    if archivo:
        try:
            clientes, resumen = datos.limpiar_base(datos.leer_archivo(archivo, archivo.name))
        except ValueError as e:
            st.error(str(e))
            return
        c = st.columns(4)
        c[0].metric("Filas leídas", resumen["filas_leidas"])
        c[1].metric("Números inválidos", resumen["telefonos_invalidos"])
        c[2].metric("Duplicados", resumen["duplicados"])
        c[3].metric("Clientes válidos", resumen["clientes_validos"])
        st.caption(f"Columnas detectadas: {resumen['columnas_detectadas']}")
        st.dataframe(clientes.head(50), hide_index=True, width="stretch")
        if st.button("Guardar en la base", type="primary"):
            r = datos.guardar_clientes(con, clientes)
            st.success(f"{r['nuevos']} clientes nuevos, {r['actualizados']} actualizados.")

    st.divider()
    st.subheader("🚫 Números que no se deben contactar")
    st.write("Sube la lista **No Molestar** del SERNAC o cualquier lista de bajas "
             "(un archivo con una columna de números).")
    excl = st.file_uploader("Lista de exclusión", type=["xlsx", "xls", "csv"], key="excl")
    if excl and st.button("Aplicar exclusiones"):
        df = datos.leer_archivo(excl, excl.name)
        n = datos.cargar_exclusiones(con, df.iloc[:, 0].tolist())
        st.success(f"{n} números marcados como NO CONTACTAR.")
    total = con.execute("SELECT COUNT(*) FROM exclusiones").fetchone()[0]
    st.caption(f"Números excluidos registrados: {total}")


# ---------------------------------------------------------------------------
def pagina_reportes():
    st.title("📊 Reportes")
    total = con.execute("SELECT COUNT(*) FROM clientes").fetchone()[0]
    if not total:
        st.info("Todavía no hay datos.")
        return
    estados = dict(con.execute("SELECT estado, COUNT(*) FROM clientes GROUP BY estado").fetchall())
    c = st.columns(4)
    c[0].metric("Clientes", total)
    c[1].metric("Ventas", estados.get("VENTA", 0))
    c[2].metric("Interesados", estados.get("INTERESADO", 0))
    c[3].metric("Pendientes", estados.get("PENDIENTE", 0) + estados.get("VOLVER_A_LLAMAR", 0))

    st.subheader("Por vendedor")
    st.dataframe(datos.resumen_por_vendedor(con), hide_index=True, width="stretch")
    c1, c2 = st.columns(2)
    c1.subheader("Clientes por estado")
    c1.bar_chart(estados, horizontal=True)
    c2.subheader("Ventas por plan")
    c2.dataframe(datos.ventas_por_plan(con), hide_index=True, width="stretch")

    gestiones = con.execute("SELECT * FROM gestiones ORDER BY fecha DESC").fetchall()
    if gestiones:
        import pandas as pd
        csv = pd.DataFrame([dict(g) for g in gestiones]).to_csv(index=False).encode("utf-8-sig")
        st.download_button("Descargar gestiones (CSV)", csv, "gestiones.csv", "text/csv")


{"Mi lista": pagina_mi_lista, "Cargar base": pagina_cargar, "Reportes": pagina_reportes}[pagina]()
