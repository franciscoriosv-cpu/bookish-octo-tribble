# Ventas de planes: gestión de la base de clientes

Mini-CRM para trabajar la base de clientes prepago (recargas) y ofrecerles planes.
Cada vendedor ve **su** lista ordenada por prioridad, registra el resultado de cada
contacto y el sistema lleva el control de ventas, re-llamados y números que no se
deben contactar.

## Qué hace

1. **Carga y limpieza de la BBDD** (Excel o CSV). Reconoce las columnas `PAIS`, `SEGMENTO`,
   `SUBSEGMENTO`, `DESC_PRODUCTO`, `MONTO_RECARGA`, la del número (aunque el encabezado
   venga cortado) y `VENDEDOR`.
   - Normaliza los números al formato `569XXXXXXXX` y descarta los que no son móviles chilenos.
   - Elimina duplicados.
   - Calcula una **prioridad** (mayor recarga primero, bono para "Persona Nueva") y un
     **plan sugerido** según el monto de recarga.
2. **Exclusiones.** Sube la lista "No Molestar" del SERNAC u otra lista de bajas y esos números
   quedan bloqueados. Si un cliente pide no ser contactado, también queda bloqueado.
3. **Mi lista (vendedor).** Muestra al siguiente cliente a contactar, con botón para llamar
   y para abrir WhatsApp, su historial y un formulario para registrar el resultado
   (No contesta, Volver a llamar, Interesado, Venta cerrada, etc.).
   - Después de 3 "No contesta" el contacto se cierra.
   - "Volver a llamar" lo esconde hasta la fecha y hora indicadas.
4. **Reportes.** Muestra ventas, interesados y conversión por vendedor, ventas por plan y
   permite exportar las gestiones a CSV.

## Instalación

```bash
pip install -r requirements.txt
```

## Uso

```bash
# Opción A: todo desde la app web
streamlit run ventas_planes/app.py

# Opción B: importar la base desde la terminal
python importar.py base.xlsx --excluir no_molestar.csv
```

Los datos quedan en `data/ventas.db` (SQLite). **Esa carpeta y los Excel/CSV están en
`.gitignore`: nunca subas la base de clientes al repositorio.**

## Personalizar

Edita `ventas_planes/config.py`:

- `PLANES_POR_RECARGA`: pon aquí los planes reales que vendes y desde qué recarga sugerir cada uno.
- `BONO_SUBSEGMENTO`: define qué subsegmentos se llaman primero.
- `MAX_INTENTOS_SIN_RESPUESTA`: cuántas veces se insiste antes de cerrar un contacto.

## Cumplimiento (importante)

- Cruza siempre la base con el registro **No Molestar** antes de una campaña.
- Usa WhatsApp o SMS de marketing solo con clientes que aceptaron ese canal.
- Respeta de inmediato a quien pide no ser contactado (la app lo hace automáticamente).
- Trata la base como dato personal (Ley 19.628 / Ley 21.719): acceso solo para quienes
  la necesitan, sin copias en correos ni pendrives.

## Pruebas

```bash
python -m pytest -q
```
