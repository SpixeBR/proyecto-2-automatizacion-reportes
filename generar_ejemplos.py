"""
Crea archivos Excel de ejemplo que imitan lo que mandan las sucursales de
"Distribuidora del Centro" (empresa ficticia) cada mes.

Cada sucursal arma su planilla "a su manera", como pasa en la vida real:
- nombres de columnas distintos ("Fecha" / "fecha venta" / "FECHA")
- títulos y filas vacías antes del encabezado
- una fila de "TOTAL" al final
- montos como texto en algunas planillas

Uso:
    python generar_ejemplos.py            # crea la carpeta entrada/
"""
from pathlib import Path

import numpy as np
import pandas as pd

CARPETA = Path(__file__).parent / "entrada"
rng = np.random.default_rng(7)

PRODUCTOS = {"Harina 000 x 25kg": 21000, "Azúcar x 50kg": 52000, "Aceite x 12u": 46000,
             "Yerba x 10kg": 39000, "Arroz x 10kg": 16500, "Fideos x 20u": 24000}

# Cada sucursal tiene su propio "estilo" de planilla
SUCURSALES = {
    "Cordoba_Centro": {"cols": ["Fecha", "Vendedor", "Producto", "Cantidad", "Precio Unitario", "Importe"],
                       "filas_titulo": 0, "monto_texto": False, "vendedores": ["Lucía Ferreyra", "Martín Ríos", "Sofía Paz"]},
    "Villa_Maria": {"cols": ["fecha venta", "vendedor", "articulo", "cant", "precio", "total"],
                    "filas_titulo": 2, "monto_texto": True, "vendedores": ["Diego Luna", "Carla Sosa"]},
    "Rio_Cuarto": {"cols": ["FECHA", "VENDEDOR", "PRODUCTO", "UNIDADES", "P. UNIT", "MONTO ($)"],
                   "filas_titulo": 3, "monto_texto": False, "vendedores": ["Ana Godoy", "Pablo Vera", "Julia Mena"]},
    "San_Francisco": {"cols": ["Fecha", "Vendedor/a", "Producto", "Cantidad", "Precio", "Total venta"],
                      "filas_titulo": 1, "monto_texto": True, "vendedores": ["Tomás Ruiz", "Paula Díaz"]},
}
MESES = pd.period_range("2025-01", "2025-06", freq="M")


def a_texto_ar(valor: float) -> str:
    entero, dec = f"{valor:,.2f}".split(".")
    return "$ " + entero.replace(",", ".") + "," + dec


def main() -> None:
    CARPETA.mkdir(exist_ok=True)
    for f in CARPETA.glob("*.xlsx"):
        f.unlink()

    for suc, cfg in SUCURSALES.items():
        for i, mes in enumerate(MESES):
            n = int(rng.integers(35, 70))
            dias = rng.integers(1, mes.days_in_month + 1, n)
            filas = []
            for d in sorted(dias):
                prod = rng.choice(list(PRODUCTOS))
                precio = round(PRODUCTOS[prod] * (1.02 ** i) * rng.uniform(0.97, 1.03), 2)
                cant = int(rng.integers(1, 12))
                total = round(precio * cant, 2)
                filas.append([pd.Timestamp(year=mes.year, month=mes.month, day=int(d)),
                              rng.choice(cfg["vendedores"]), prod, cant, precio,
                              a_texto_ar(total) if cfg["monto_texto"] else total])
            df = pd.DataFrame(filas, columns=cfg["cols"])

            # Error de carga típico: una venta con cantidad 0
            if suc == "Rio_Cuarto" and mes.month == 3:
                df.iloc[10, 3] = 0

            # Algunas filas vacías en el medio (alguien borró contenido y no la fila)
            if rng.random() < 0.5:
                pos = int(rng.integers(5, len(df) - 5))
                df = pd.concat([df.iloc[:pos], pd.DataFrame([[None] * 6], columns=df.columns), df.iloc[pos:]])

            archivo = CARPETA / f"ventas_{suc}_{mes.strftime('%Y_%m')}.xlsx"
            with pd.ExcelWriter(archivo, engine="openpyxl") as w:
                df.to_excel(w, index=False, startrow=cfg["filas_titulo"], sheet_name="Ventas")
                hoja = w.sheets["Ventas"]
                if cfg["filas_titulo"]:
                    hoja.cell(row=1, column=1, value=f"Ventas sucursal {suc.replace('_', ' ')} - {mes.strftime('%m/%Y')}")
                # Fila de total al final, como hacen muchas planillas armadas a mano
                ultima = cfg["filas_titulo"] + len(df) + 2
                hoja.cell(row=ultima, column=1, value="TOTAL")
                total_mes = sum(f[5] if not isinstance(f[5], str) else 0 for f in filas)
                hoja.cell(row=ultima, column=6, value=total_mes if total_mes else "ver arriba")

    # Alguien mandó dos veces el mismo archivo con otro nombre
    import shutil
    shutil.copy(CARPETA / "ventas_Villa_Maria_2025_02.xlsx", CARPETA / "ventas_Villa_Maria_2025_02 (copia).xlsx")

    print(f"Creados {len(list(CARPETA.glob('*.xlsx')))} archivos en {CARPETA}")


if __name__ == "__main__":
    main()
