"""Tests del consolidador de planillas. Ejecutar con:  pytest -q"""
import shutil
import sys
from datetime import datetime
from pathlib import Path

import openpyxl
import pandas as pd
import pytest

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from generar_reporte import (  # noqa: E402
    Registro, a_fecha, a_numero, consolidar, encontrar_encabezado, escribir_reporte, mapear_columna,
)


def crear_planilla(ruta: Path, columnas: list[str], filas: list[list], filas_titulo: int = 0, total: bool = True) -> None:
    """Crea una planilla como las que mandan las sucursales: título opcional arriba y fila TOTAL abajo."""
    wb = openpyxl.Workbook()
    hoja = wb.active
    for _ in range(filas_titulo):
        hoja.append(["Ventas de la sucursal"])
    hoja.append(columnas)
    for f in filas:
        hoja.append(f)
    if total:
        hoja.append([])
        hoja.append(["TOTAL", None, None, None, None, 999999])
    wb.save(ruta)


VENTA = [datetime(2025, 3, 10), "Ana", "Yerba x 10kg", 2, 1000, 2000]


# ---- conversiones

@pytest.mark.parametrize("valor, esperado", [
    (1500, 1500.0), ("1500.5", 1500.5), ("$ 1.234,50", 1234.5), ("$ 1.234", 1234.0), ("  $ 12.000,00 ", 12000.0),
])
def test_a_numero(valor, esperado):
    assert a_numero(valor) == pytest.approx(esperado)


@pytest.mark.parametrize("valor, esperado", [
    (datetime(2025, 3, 15), "2025-03-15"), ("15/03/2025", "2025-03-15"), ("2025-03-15", "2025-03-15"),
    (45658, "2025-01-01"),  # número de serie de Excel
])
def test_a_fecha(valor, esperado):
    assert a_fecha(valor) == pd.Timestamp(esperado)


def test_a_fecha_invalida():
    assert pd.isna(a_fecha("sin fecha"))
    assert pd.isna(a_fecha(None))


@pytest.mark.parametrize("nombre, esperado", [
    ("Importe", "total"), ("MONTO ($)", "total"), ("Total venta", "total"), ("fecha venta", "fecha"),
    ("Vendedor/a", "vendedor"), ("P. UNIT", "precio_unitario"), ("artículo", "producto"), ("Observaciones", None),
])
def test_mapear_columna(nombre, esperado):
    assert mapear_columna(nombre) == esperado


def test_encuentra_encabezado_debajo_del_titulo():
    crudo = pd.DataFrame([["Ventas marzo", None, None, None], [None] * 4,
                          ["Fecha", "Vendedor", "Producto", "Cantidad"], [VENTA[0], "Ana", "Yerba", 2]])
    assert encontrar_encabezado(crudo) == 2


# ---- consolidación

COLS_A = ["Fecha", "Vendedor", "Producto", "Cantidad", "Precio Unitario", "Importe"]
COLS_B = ["FECHA", "VENDEDOR", "ARTICULO", "UNIDADES", "P. UNIT", "MONTO ($)"]


def test_consolida_formatos_distintos_e_ignora_fila_total(tmp_path):
    crear_planilla(tmp_path / "ventas_Centro_2025_03.xlsx", COLS_A, [VENTA, VENTA])
    crear_planilla(tmp_path / "ventas_Norte_2025_03.xlsx", COLS_B,
                   [[VENTA[0], "Luis", "Azúcar x 50kg", 1, 5000, "$ 5.000,00"]], filas_titulo=3)
    datos = consolidar(tmp_path, Registro())
    assert len(datos) == 3                       # dos ventas iguales en el mismo archivo son legítimas
    assert datos["total"].sum() == 9000          # la fila TOTAL (999999) no se suma
    assert set(datos["sucursal"]) == {"Centro", "Norte"}


def test_omite_archivo_copiado(tmp_path):
    crear_planilla(tmp_path / "ventas_Centro_2025_03.xlsx", COLS_A, [VENTA])
    shutil.copy(tmp_path / "ventas_Centro_2025_03.xlsx", tmp_path / "ventas_Centro_2025_03 (copia).xlsx")
    registro = Registro()
    datos = consolidar(tmp_path, registro)
    assert len(datos) == 1
    assert datos["archivo_origen"].iloc[0] == "ventas_Centro_2025_03.xlsx"   # se conserva el original
    assert any("copia exacta" in e["detalle"] for e in registro.eventos)


def test_avisa_dos_versiones_del_mismo_periodo(tmp_path):
    crear_planilla(tmp_path / "ventas_Centro_2025_03.xlsx", COLS_A, [VENTA])
    otra = VENTA[:3] + [3, 1000, 3000]
    # Mismo período y sucursal, con contenido distinto (por ejemplo, una versión corregida)
    crear_planilla(tmp_path / "ventas_centro_2025_03_corregido.xlsx", COLS_A, [otra])
    registro = Registro()
    datos = consolidar(tmp_path, registro)
    assert len(datos) == 1
    assert any("ya se cargó" in e["detalle"] for e in registro.eventos)


def test_excluye_cantidades_invalidas_y_lo_registra(tmp_path):
    crear_planilla(tmp_path / "ventas_Centro_2025_03.xlsx", COLS_A, [VENTA, VENTA[:3] + [0, 1000, 0]])
    registro = Registro()
    datos = consolidar(tmp_path, registro)
    assert len(datos) == 1
    assert any(e["nivel"] == "AVISO" and "inválidos" in e["detalle"] for e in registro.eventos)


def test_genera_reporte_con_todas_las_hojas(tmp_path):
    entrada = tmp_path / "entrada"
    entrada.mkdir()
    crear_planilla(entrada / "ventas_Centro_2025_03.xlsx", COLS_A, [VENTA])
    crear_planilla(entrada / "ventas_Centro_2025_04.xlsx", COLS_A, [[datetime(2025, 4, 2)] + VENTA[1:]])
    registro = Registro()
    salida = tmp_path / "reporte.xlsx"
    escribir_reporte(consolidar(entrada, registro), registro, salida)
    wb = openpyxl.load_workbook(salida)
    assert wb.sheetnames == ["Resumen", "Sucursales", "Vendedores", "Productos", "Datos", "Log"]
    assert wb["Datos"].max_row == 3              # encabezado + 2 ventas
