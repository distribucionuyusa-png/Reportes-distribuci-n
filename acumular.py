# -*- coding: utf-8 -*-
"""
acumular.py
-----------
Fusiona el corte diario de DispatchTrack (reportes/despachos_diarios.xls)
dentro de un maestro acumulado (historico/historico_despachos.xlsx),
usando el N. de Orden como llave unica (upsert):

  - Si la Orden ya existe en el maestro  -> se actualiza con el dato mas reciente.
  - Si la Orden es nueva                  -> se agrega.

Agrega la columna "Fecha de corte" (fecha local de Peru del momento en que corre)
y guarda el resultado como una Tabla de Excel con nombre fijo (tbl_historico),
en la hoja HISTORICO, conservando TODAS las columnas del export. Si algun dia el
export trae columnas nuevas, el maestro las incorpora sin perder las anteriores.
"""

import os
from datetime import datetime
from zoneinfo import ZoneInfo

import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo

SNAPSHOT = os.path.join("reportes", "despachos_diarios.xls")
MASTER_DIR = "historico"
MASTER = os.path.join(MASTER_DIR, "historico_despachos.xlsx")
SHEET = "HISTORICO"
TABLE = "tbl_historico"
KEY = "Orden"
CUT_COL = "Fecha de corte"


def _norm_key(s: pd.Series) -> pd.Series:
    """Orden como texto estable (sin decimales) para que sea una llave confiable."""
    num = pd.to_numeric(s, errors="coerce")
    out = num.astype("Int64").astype("string")
    # para valores no numericos, conserva el texto original
    out = out.where(num.notna(), s.astype("string"))
    return out.fillna("").str.strip()


def _unique_headers(cols):
    """Garantiza encabezados unicos y no vacios (requisito de una Tabla de Excel)."""
    seen, result = {}, []
    for i, c in enumerate(cols):
        name = str(c).strip() or f"Columna_{i+1}"
        if name in seen:
            seen[name] += 1
            name = f"{name}.{seen[name]}"
        else:
            seen[name] = 0
        result.append(name)
    return result


def main():
    if not os.path.exists(SNAPSHOT):
        raise SystemExit(f"No se encontro el corte diario: {SNAPSHOT}")

    # 1) Corte del dia
    snap = pd.read_excel(SNAPSHOT, sheet_name=0)
    snap.columns = _unique_headers(snap.columns)
    if KEY not in snap.columns:
        raise SystemExit(f"El corte no tiene la columna llave '{KEY}'.")
    snap[KEY] = _norm_key(snap[KEY])
    snap = snap[snap[KEY] != ""]                      # descarta filas sin Orden
    snap = snap.drop_duplicates(subset=KEY, keep="last")

    fecha_corte = datetime.now(ZoneInfo("America/Lima")).strftime("%Y-%m-%d")
    snap.insert(0, CUT_COL, fecha_corte)

    # 2) Maestro existente (si lo hay)
    if os.path.exists(MASTER):
        master = pd.read_excel(MASTER, sheet_name=SHEET)
        master.columns = _unique_headers(master.columns)
        if KEY in master.columns:
            master[KEY] = _norm_key(master[KEY])
    else:
        master = pd.DataFrame(columns=snap.columns)

    # 3) Union de columnas: conserva el orden del maestro y agrega nuevas al final
    cols = list(master.columns)
    for c in snap.columns:
        if c not in cols:
            cols.append(c)
    master = master.reindex(columns=cols)
    snap = snap.reindex(columns=cols)

    # 4) Upsert: el corte de hoy gana sobre lo que ya habia (keep='last')
    antes = master[KEY].nunique() if len(master) else 0
    combinado = pd.concat([master, snap], ignore_index=True)
    combinado = combinado.drop_duplicates(subset=KEY, keep="last").reset_index(drop=True)

    # orden de lectura: por fecha de corte y luego Orden
    sort_cols = [c for c in [CUT_COL, KEY] if c in combinado.columns]
    combinado = combinado.sort_values(sort_cols, na_position="last").reset_index(drop=True)

    nuevas = combinado[KEY].nunique() - antes
    actualizadas = len(snap) - max(nuevas, 0)

    # 5) Escribir maestro como Tabla de Excel
    os.makedirs(MASTER_DIR, exist_ok=True)
    tmp = MASTER + ".tmp.xlsx"
    with pd.ExcelWriter(tmp, engine="openpyxl") as xw:
        combinado.to_excel(xw, index=False, sheet_name=SHEET)

    wb = load_workbook(tmp)
    ws = wb[SHEET]
    nrows, ncols = combinado.shape
    ref = f"A1:{get_column_letter(ncols)}{nrows + 1}"
    tbl = Table(displayName=TABLE, ref=ref)
    tbl.tableStyleInfo = TableStyleInfo(
        name="TableStyleMedium2", showRowStripes=True, showColumnStripes=False
    )
    ws.add_table(tbl)

    hdr_fill = PatternFill("solid", fgColor="1F4E78")
    hdr_font = Font(name="Arial", bold=True, color="FFFFFF")
    for j in range(1, ncols + 1):
        c = ws.cell(1, j)
        c.fill, c.font = hdr_fill, hdr_font
        c.alignment = Alignment(vertical="center")
    for col_cells in ws.iter_cols(min_row=2, max_row=min(nrows + 1, 400)):
        for c in col_cells:
            c.font = Font(name="Arial", size=10)
    ws.freeze_panes = "A2"

    wb.save(MASTER)
    wb.close()
    os.remove(tmp)

    print(f"Fecha de corte : {fecha_corte}")
    print(f"Corte del dia  : {len(snap)} ordenes")
    print(f"Maestro antes  : {antes} ordenes")
    print(f"  nuevas       : {max(nuevas,0)}")
    print(f"  actualizadas : {max(actualizadas,0)}")
    print(f"Maestro ahora  : {combinado[KEY].nunique()} ordenes | {nrows} filas x {ncols} columnas")
    print(f"Guardado en    : {MASTER} (hoja {SHEET}, tabla {TABLE})")


if __name__ == "__main__":
    main()
