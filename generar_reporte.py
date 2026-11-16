"""
Consolida las planillas de ventas de todas las sucursales y genera un reporte
en Excel listo para gerencia, en un solo paso.

Qué hace:
1. Lee todos los .xlsx de la carpeta de entrada.
2. Encuentra el encabezado aunque haya títulos o filas vacías arriba.
3. Unifica nombres de columnas distintos ("Importe", "MONTO ($)", "total"...).
4. Limpia filas vacías, filas de TOTAL y montos escritos como texto ("$ 1.234,50").
5. Valida los datos y registra cada problema en un log.
6. Genera un Excel con resumen, rankings, gráfico, datos consolidados y log.

Uso:
    python generar_reporte.py --entrada entrada --salida reporte_ventas.xlsx
"""
from __future__ import annotations

import argparse
import hashlib
import logging
import re
import unicodedata
from dataclasses import dataclass, field
from datetime import date, datetime
from pathlib import Path

import pandas as pd

log = logging.getLogger("reporte")

# Nombre estándar -> variantes posibles (ya normalizadas: minúsculas, sin acentos ni símbolos)
SINONIMOS = {
    "fecha": {"fecha", "fechaventa", "fechadeventa"},
    "vendedor": {"vendedor", "vendedora", "vendedores"},
    "producto": {"producto", "articulo", "item", "descripcion"},
    "cantidad": {"cantidad", "cant", "unidades", "cantidades"},
    "precio_unitario": {"preciounitario", "precio", "punit", "preciounit"},
    "total": {"total", "importe", "monto", "totalventa", "importetotal"},
}
COLUMNAS = list(SINONIMOS)
PATRON_ARCHIVO = re.compile(r"ventas_(?P<sucursal>.+)_(?P<anio>\d{4})_(?P<mes>\d{2})", re.IGNORECASE)
TOLERANCIA_TOTAL = 1.0  # pesos de diferencia aceptados entre total y precio x cantidad
# Nombre para mostrar de cada sucursal (los nombres de archivo no llevan acentos ni espacios).
# Si una sucursal no está acá, se usa el nombre del archivo cambiando "_" por espacios.
NOMBRES_SUCURSALES = {"Cordoba_Centro": "Córdoba Centro", "Villa_Maria": "Villa María",
                      "Rio_Cuarto": "Río Cuarto", "San_Francisco": "San Francisco"}
MILES_SIN_DECIMALES = re.compile(r"\d{1,3}(\.\d{3})+")  # "1.234" en formato argentino
FORMATOS_FECHA = ("%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y", "%d/%m/%y")
ORIGEN_EXCEL = pd.Timestamp("1899-12-30")  # día 0 de las fechas guardadas como número en Excel


@dataclass
class Registro:
    """Guarda lo que pasó con cada archivo para poder auditarlo después."""
    eventos: list[dict] = field(default_factory=list)

    def agregar(self, archivo: str, nivel: str, mensaje: str) -> None:
        self.eventos.append({"archivo": archivo, "nivel": nivel, "detalle": mensaje})
        getattr(log, "warning" if nivel == "AVISO" else "info")(f"{archivo}: {mensaje}")


# ---------------------------------------------------------------- lectura

def normalizar(texto) -> str:
    texto = unicodedata.normalize("NFKD", str(texto)).encode("ascii", "ignore").decode()
    return re.sub(r"[^a-z]", "", texto.lower())


def mapear_columna(nombre) -> str | None:
    clave = normalizar(nombre)
    for estandar, variantes in SINONIMOS.items():
        if clave in variantes:
            return estandar
    return None


def encontrar_encabezado(crudo: pd.DataFrame, max_filas: int = 15) -> int | None:
    """Devuelve el índice de la primera fila que tiene al menos 4 nombres de columna conocidos."""
    for i in range(min(max_filas, len(crudo))):
        reconocidas = {mapear_columna(v) for v in crudo.iloc[i] if pd.notna(v)} - {None}
        if len(reconocidas) >= 4:
            return i
    return None


def a_numero(valor) -> float:
    """Convierte 1234.5, '1234.5', '$ 1.234,50' o '$ 1.234' a float."""
    if pd.isna(valor):
        return float("nan")
    if isinstance(valor, (int, float)):
        return float(valor)
    texto = str(valor).replace("$", "").replace(" ", "").strip()
    if "," in texto:  # formato argentino: punto de miles y coma decimal
        texto = texto.replace(".", "").replace(",", ".")
    elif MILES_SIN_DECIMALES.fullmatch(texto):
        texto = texto.replace(".", "")
    try:
        return float(texto)
    except ValueError:
        return float("nan")


def a_fecha(valor) -> pd.Timestamp:
    """Acepta fechas de Excel (datetime), números de serie de Excel (45658) y texto ('15/03/2025')."""
    if valor is None or (isinstance(valor, float) and pd.isna(valor)):
        return pd.NaT
    if isinstance(valor, (datetime, date, pd.Timestamp)):
        return pd.Timestamp(valor)
    if isinstance(valor, (int, float)):
        return ORIGEN_EXCEL + pd.Timedelta(days=float(valor)) if 20000 < valor < 80000 else pd.NaT
    texto = str(valor).strip()
    for formato in FORMATOS_FECHA:
        try:
            return pd.to_datetime(texto, format=formato)
        except ValueError:
            continue
    return pd.NaT


def huella(ruta: Path) -> str:
    """Hash del contenido del archivo, para detectar el mismo archivo enviado con otro nombre."""
    return hashlib.md5(ruta.read_bytes()).hexdigest()


def leer_planilla(ruta: Path, registro: Registro) -> pd.DataFrame | None:
    nombre = ruta.name
    try:
        crudo = pd.read_excel(ruta, header=None, dtype=object)
    except Exception as e:  # noqa: BLE001  # archivo dañado, abierto, etc.
        registro.agregar(nombre, "AVISO", f"no se pudo abrir ({e}); archivo omitido")
        return None

    fila_enc = encontrar_encabezado(crudo)
    if fila_enc is None:
        registro.agregar(nombre, "AVISO", "no se encontró el encabezado; archivo omitido")
        return None

    df = crudo.iloc[fila_enc + 1:].copy()
    df.columns = [mapear_columna(c) for c in crudo.iloc[fila_enc]]
    faltantes = [c for c in COLUMNAS if c not in df.columns]
    if faltantes:
        registro.agregar(nombre, "AVISO", f"faltan columnas {faltantes}; archivo omitido")
        return None
    df = df[COLUMNAS]

    # Filas vacías y fila de TOTAL
    n0 = len(df)
    df = df.dropna(how="all")
    vacias = n0 - len(df)
    es_total = df["fecha"].astype(str).str.strip().str.lower().str.startswith("total")
    df = df[~es_total]

    # Tipos
    df["fecha"] = df["fecha"].map(a_fecha)
    for col in ["cantidad", "precio_unitario", "total"]:
        df[col] = df[col].map(a_numero)
    for col in ["vendedor", "producto"]:
        df[col] = df[col].where(df[col].notna(), "Sin dato").astype(str).str.strip()

    # Validaciones
    invalidas = df["fecha"].isna() | df["cantidad"].isna() | (df["cantidad"] <= 0) | df["total"].isna()
    if invalidas.any():
        registro.agregar(nombre, "AVISO", f"filas excluidas por fecha, cantidad o total inválidos: {int(invalidas.sum())}")
    df = df[~invalidas]

    diferencia = (df["total"] - df["cantidad"] * df["precio_unitario"]).abs()
    inconsistentes = diferencia > TOLERANCIA_TOTAL
    if inconsistentes.any():
        registro.agregar(nombre, "AVISO", f"filas donde total ≠ cantidad × precio: {int(inconsistentes.sum())} (se conserva el total informado)")

    # Sucursal y período desde el nombre del archivo
    m = PATRON_ARCHIVO.search(ruta.stem)
    if m:
        df["sucursal"] = NOMBRES_SUCURSALES.get(m["sucursal"], m["sucursal"].replace("_", " "))
    else:
        df["sucursal"] = ruta.stem
    df["archivo_origen"] = nombre
    if m:
        fuera = df["fecha"].dt.strftime("%Y_%m") != f"{m['anio']}_{m['mes']}"
        if fuera.any():
            registro.agregar(nombre, "AVISO", f"ventas con fecha fuera del mes del archivo: {int(fuera.sum())}")

    detalle = f"ventas leídas: {len(df)} (encabezado en la fila {fila_enc + 1}"
    detalle += f"; filas vacías ignoradas: {vacias})" if vacias else ")"
    registro.agregar(nombre, "OK", detalle)
    return df


def consolidar(carpeta: Path, registro: Registro) -> pd.DataFrame:
    if not carpeta.is_dir():
        raise SystemExit(f"No existe la carpeta {carpeta}")
    # "~$" son archivos temporales que Excel crea mientras una planilla está abierta
    # Los nombres tipo "(copia)" o "(1)" van al final, así ante un duplicado se conserva el original
    es_copia = re.compile(r"copia|copy|\(\d+\)", re.IGNORECASE)
    archivos = sorted((p for p in carpeta.glob("*.xlsx") if not p.name.startswith("~$")),
                      key=lambda p: (bool(es_copia.search(p.stem)), p.name))
    if not archivos:
        raise SystemExit(f"No hay archivos .xlsx en {carpeta}")

    partes = []
    vistos: dict[str, str] = {}        # huella del contenido -> archivo
    periodos: dict[tuple, str] = {}    # (sucursal, año, mes) -> archivo
    for archivo in archivos:
        # Un mismo archivo enviado dos veces (aunque tenga otro nombre) se carga una sola vez.
        # Se compara el archivo entero y no venta por venta, porque dos ventas idénticas el mismo día son posibles.
        h = huella(archivo)
        if h in vistos:
            registro.agregar(archivo.name, "AVISO", f"es una copia exacta de {vistos[h]}; archivo omitido")
            continue
        vistos[h] = archivo.name

        m = PATRON_ARCHIVO.search(archivo.stem)
        if m:
            clave = (m["sucursal"].lower(), m["anio"], m["mes"])
            if clave in periodos:
                registro.agregar(archivo.name, "AVISO", f"ese período ya se cargó desde {periodos[clave]}; "
                                                       "archivo omitido (revisar cuál es la versión correcta)")
                continue
            periodos[clave] = archivo.name

        df = leer_planilla(archivo, registro)
        if df is not None and len(df):
            partes.append(df)

    if not partes:
        raise SystemExit("Ningún archivo pudo procesarse. Revisá el log.")
    datos = pd.concat(partes, ignore_index=True)
    datos["mes"] = datos["fecha"].dt.to_period("M").astype(str)
    return datos.sort_values(["fecha", "sucursal"]).reset_index(drop=True)


# ---------------------------------------------------------------- reporte

def escribir_reporte(datos: pd.DataFrame, registro: Registro, salida: Path) -> None:
    por_suc_mes = datos.pivot_table(index="sucursal", columns="mes", values="total", aggfunc="sum", fill_value=0)
    por_suc_mes = por_suc_mes.loc[por_suc_mes.sum(axis=1).sort_values(ascending=False).index]

    sucursales = datos.groupby("sucursal").agg(facturacion=("total", "sum"), operaciones=("total", "size"),
                                               unidades=("cantidad", "sum")).sort_values("facturacion", ascending=False)
    sucursales["ticket_promedio"] = sucursales["facturacion"] / sucursales["operaciones"]
    sucursales["participacion"] = sucursales["facturacion"] / sucursales["facturacion"].sum()

    vendedores = (datos.groupby(["vendedor", "sucursal"]).agg(facturacion=("total", "sum"), operaciones=("total", "size"))
                  .sort_values("facturacion", ascending=False).reset_index())
    productos = (datos.groupby("producto").agg(unidades=("cantidad", "sum"), facturacion=("total", "sum"))
                 .sort_values("facturacion", ascending=False).reset_index())

    with pd.ExcelWriter(salida, engine="xlsxwriter", datetime_format="dd/mm/yyyy") as w:
        libro = w.book
        f_titulo = libro.add_format({"bold": True, "font_size": 16, "font_color": "#1c5cab"})
        f_sub = libro.add_format({"italic": True, "font_color": "#52514e"})
        f_kpi_lbl = libro.add_format({"font_color": "#52514e", "font_size": 9, "top": 2, "top_color": "#2a78d6"})
        f_kpi_val = libro.add_format({"bold": True, "font_size": 14, "num_format": "$ #,##0", "align": "left"})
        f_kpi_num = libro.add_format({"bold": True, "font_size": 14, "num_format": "#,##0", "align": "left"})
        f_enc = libro.add_format({"bold": True, "bg_color": "#1c5cab", "font_color": "white", "border": 1, "text_wrap": True, "valign": "vcenter"})
        f_pesos = libro.add_format({"num_format": "$ #,##0", "border": 1})
        f_num = libro.add_format({"num_format": "#,##0", "border": 1})
        f_pct = libro.add_format({"num_format": "0.0%", "border": 1})
        f_txt = libro.add_format({"border": 1})
        f_tot = libro.add_format({"bold": True, "num_format": "$ #,##0", "top": 2, "bg_color": "#eef4fc"})
        f_tot_txt = libro.add_format({"bold": True, "top": 2, "bg_color": "#eef4fc"})

        n_ok = sum(e["nivel"] == "OK" for e in registro.eventos)
        n_avisos = sum(e["nivel"] == "AVISO" for e in registro.eventos)

        # --- Hoja Resumen
        hoja = libro.add_worksheet("Resumen")
        hoja.hide_gridlines(2)
        hoja.set_landscape()
        hoja.fit_to_pages(1, 1)  # lista para imprimir en una hoja
        hoja.set_column("A:A", 22)
        hoja.set_column("B:H", 15)
        hoja.write("A1", "Reporte de ventas consolidado", f_titulo)
        hoja.write("A2", f"Período {datos['fecha'].min():%d/%m/%Y} – {datos['fecha'].max():%d/%m/%Y} · "
                         f"{n_ok} archivos procesados · generado {datetime.now():%d/%m/%Y %H:%M}", f_sub)

        kpis = [("Facturación total", datos["total"].sum(), f_kpi_val),
                ("Operaciones", len(datos), f_kpi_num),
                ("Unidades vendidas", datos["cantidad"].sum(), f_kpi_num),
                ("Ticket promedio", datos["total"].mean(), f_kpi_val)]
        for i, (lbl, val, fmt) in enumerate(kpis):
            col = i * 2
            hoja.merge_range(3, col, 3, col + 1, lbl, f_kpi_lbl)
            hoja.merge_range(4, col, 4, col + 1, float(val), fmt)

        fila0 = 7
        hoja.write(fila0 - 1, 0, "Facturación por sucursal y mes", libro.add_format({"bold": True, "font_size": 12}))
        hoja.write(fila0, 0, "Sucursal", f_enc)
        meses = list(por_suc_mes.columns)
        for j, mes in enumerate(meses):
            hoja.write(fila0, j + 1, _nombre_mes(mes), f_enc)
        hoja.write(fila0, len(meses) + 1, "Total", f_enc)
        for i, (suc, fila) in enumerate(por_suc_mes.iterrows(), start=1):
            hoja.write(fila0 + i, 0, suc, f_txt)
            for j, v in enumerate(fila):
                hoja.write_number(fila0 + i, j + 1, float(v), f_pesos)
            # Fórmulas reales, así el cliente puede auditar o modificar los números
            hoja.write_formula(fila0 + i, len(meses) + 1,
                               f"=SUM({_celda(fila0 + i, 1)}:{_celda(fila0 + i, len(meses))})", f_tot, float(fila.sum()))
        fila_tot = fila0 + len(por_suc_mes) + 1
        hoja.write(fila_tot, 0, "Total", f_tot_txt)
        for j in range(1, len(meses) + 2):
            hoja.write_formula(fila_tot, j, f"=SUM({_celda(fila0 + 1, j)}:{_celda(fila_tot - 1, j)})", f_tot,
                               float(por_suc_mes.iloc[:, j - 1].sum()) if j <= len(meses) else float(por_suc_mes.values.sum()))

        # Gráfico nativo de Excel (se actualiza si cambian los datos)
        graf = libro.add_chart({"type": "column", "subtype": "stacked"})
        paleta = ["#2a78d6", "#eb6834", "#1baf7a", "#eda100", "#e87ba4", "#008300", "#6250d6", "#e34948"]
        for i in range(len(por_suc_mes)):
            graf.add_series({
                "name": ["Resumen", fila0 + 1 + i, 0],
                "categories": ["Resumen", fila0, 1, fila0, len(meses)],
                "values": ["Resumen", fila0 + 1 + i, 1, fila0 + 1 + i, len(meses)],
                "fill": {"color": paleta[i % len(paleta)]}, "border": {"color": "white", "width": 1.5},
                "gap": 60,
            })
        graf.set_title({"name": "Facturación mensual por sucursal", "name_font": {"size": 12}})
        graf.set_y_axis({"num_format": "$ #,##0,,\" M\"", "major_gridlines": {"visible": True, "line": {"color": "#e4e3df"}}})
        graf.set_legend({"position": "bottom"})
        graf.set_size({"width": 760, "height": 340})
        hoja.insert_chart(fila_tot + 2, 0, graf)

        # --- Rankings
        def tabla(nombre, df, formatos, anchos):
            df.to_excel(w, sheet_name=nombre, index=False, startrow=1, header=False)
            h = w.sheets[nombre]
            h.add_table(0, 0, len(df), len(df.columns) - 1,
                        {"columns": [{"header": c} for c in df.columns], "style": "Table Style Medium 2"})
            for j, (fmt, ancho) in enumerate(zip(formatos, anchos)):
                h.set_column(j, j, ancho, fmt)
            h.freeze_panes(1, 0)

        suc_out = sucursales.reset_index().rename(columns={
            "sucursal": "Sucursal", "facturacion": "Facturación", "operaciones": "Operaciones",
            "unidades": "Unidades", "ticket_promedio": "Ticket promedio", "participacion": "% del total"})
        tabla("Sucursales", suc_out, [None, f_pesos, f_num, f_num, f_pesos, f_pct], [20, 16, 13, 12, 16, 12])

        ven_out = vendedores.rename(columns={"vendedor": "Vendedor", "sucursal": "Sucursal",
                                             "facturacion": "Facturación", "operaciones": "Operaciones"})
        ven_out.insert(0, "Puesto", range(1, len(ven_out) + 1))
        tabla("Vendedores", ven_out, [None, None, None, f_pesos, f_num], [8, 20, 18, 16, 13])

        prod_out = productos.rename(columns={"producto": "Producto", "unidades": "Unidades", "facturacion": "Facturación"})
        tabla("Productos", prod_out, [None, f_num, f_pesos], [22, 12, 16])

        datos_out = datos[["fecha", "sucursal", "vendedor", "producto", "cantidad", "precio_unitario", "total", "archivo_origen"]]
        datos_out = datos_out.rename(columns={"fecha": "Fecha", "sucursal": "Sucursal", "vendedor": "Vendedor",
                                              "producto": "Producto", "cantidad": "Cantidad", "precio_unitario": "Precio unitario",
                                              "total": "Total", "archivo_origen": "Archivo de origen"})
        tabla("Datos", datos_out, [None, None, None, None, f_num, f_pesos, f_pesos, None], [12, 16, 16, 20, 10, 15, 14, 36])

        log_out = pd.DataFrame(registro.eventos).rename(columns={"archivo": "Archivo", "nivel": "Estado", "detalle": "Detalle"})
        tabla("Log", log_out, [None, None, None], [36, 9, 80])
        hoja.write(fila_tot + 21, 0, f"Procesamiento: {n_ok} archivos leídos, {n_avisos} avisos. Ver hoja 'Log'.", f_sub)


MESES_ES = ["Ene", "Feb", "Mar", "Abr", "May", "Jun", "Jul", "Ago", "Sep", "Oct", "Nov", "Dic"]


def _nombre_mes(periodo: str) -> str:
    """'2025-03' -> 'Mar 2025'"""
    anio, mes = periodo.split("-")
    return f"{MESES_ES[int(mes) - 1]} {anio}"


def _celda(fila: int, col: int) -> str:
    letras = ""
    col += 1
    while col:
        col, r = divmod(col - 1, 26)
        letras = chr(65 + r) + letras
    return f"{letras}{fila + 1}"


# ---------------------------------------------------------------- main

def main(argv: list[str] | None = None) -> None:
    parser = argparse.ArgumentParser(description="Consolida planillas de ventas y genera un reporte en Excel.")
    parser.add_argument("--entrada", default="entrada", type=Path, help="carpeta con los .xlsx de las sucursales")
    parser.add_argument("--salida", default="reporte_ventas.xlsx", type=Path, help="archivo Excel a generar")
    args = parser.parse_args(argv)

    logging.basicConfig(level=logging.INFO, format="%(levelname)-7s %(message)s")
    registro = Registro()
    datos = consolidar(args.entrada, registro)
    escribir_reporte(datos, registro, args.salida)
    log.info(f"Listo: {len(datos)} ventas de {datos['sucursal'].nunique()} sucursales -> {args.salida}")


if __name__ == "__main__":
    main()
