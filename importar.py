"""Importa la BBDD desde la terminal.

Uso:
    python importar.py base.xlsx
    python importar.py base.xlsx --excluir no_molestar.csv
"""

import argparse

from ventas_planes import datos


def main():
    p = argparse.ArgumentParser(description="Limpia e importa la base de clientes.")
    p.add_argument("archivo", help="Excel o CSV con la BBDD")
    p.add_argument("--excluir", help="Archivo con números que NO se deben contactar (1ª columna)")
    args = p.parse_args()

    con = datos.conectar()
    if args.excluir:
        lista = datos.leer_archivo(args.excluir).iloc[:, 0].tolist()
        print(f"Exclusiones registradas: {datos.cargar_exclusiones(con, lista)}")

    clientes, resumen = datos.limpiar_base(datos.leer_archivo(args.archivo))
    print(f"Columnas detectadas: {resumen['columnas_detectadas']}")
    print(f"Filas leídas: {resumen['filas_leidas']}  |  inválidos: {resumen['telefonos_invalidos']}"
          f"  |  duplicados: {resumen['duplicados']}  |  válidos: {resumen['clientes_validos']}")
    r = datos.guardar_clientes(con, clientes)
    print(f"Guardados: {r['nuevos']} nuevos, {r['actualizados']} actualizados.")
    print(datos.resumen_por_vendedor(con).to_string(index=False))


if __name__ == "__main__":
    main()
