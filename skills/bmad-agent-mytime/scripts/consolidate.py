# /// script
# requires-python = ">=3.10"
# dependencies = ["pywin32"]
# ///
"""Converte i due .xls scaricati e crea il consolidato con la Power Query.

Richiede Microsoft Excel desktop. Legge risorse-raw.xls e wbs-raw.xls dalla
cartella downloads e scrive i tre .xlsx con i nomi previsti dalla skill.
"""

import argparse
import datetime
from pathlib import Path

import win32com.client

XL_OPENXML = 51
XL_SRC_RANGE = 1
XL_SRC_EXTERNAL = 0

M_QUERY = """let
    Risorse = Excel.CurrentWorkbook(){[Name="TabRisorse"]}[Content],
    Wbs = Excel.CurrentWorkbook(){[Name="TabWbs"]}[Content],
    Colonne = Table.ColumnNames(Risorse),
    WbsNuovi = Table.RemoveColumns(Table.NestedJoin(Wbs, Colonne, Table.Buffer(Risorse), Colonne, "Match", JoinKind.LeftAnti), {"Match"}, MissingField.Ignore),
    Normalizzato = Table.Combine({Risorse, WbsNuovi})
in
    Normalizzato"""


def compact(d):
    return datetime.datetime.strptime(d, "%d/%m/%Y").strftime("%Y%m%d")


def convert(excel, raw, target):
    wb = excel.Workbooks.Open(str(raw))
    wb.SaveAs(str(target), XL_OPENXML)
    wb.Close(False)


def add_data_sheet(excel, wb, source_xlsx, sheet_name, table_name):
    src = excel.Workbooks.Open(str(source_xlsx))
    ws = wb.Worksheets.Add(None, wb.Worksheets(wb.Worksheets.Count))
    ws.Name = sheet_name
    src.Worksheets(1).UsedRange.Copy(ws.Range("A1"))
    src.Close(False)
    table = ws.ListObjects.Add(XL_SRC_RANGE, ws.UsedRange, None, 1)
    table.Name = table_name
    return ws


def add_key_columns(table):
    data = 'IF(ISNUMBER([@Data]),[@Data],DATEVALUE(LEFT(TRIM([@Data]),10)))'
    keys = [
        ("K-RISORSA", 'SUBSTITUTE(TRIM([@Utente])," ","#")'),
        ("K-WBS", "TRIM([@Wbs])"),
        ("K-DATE", f'TEXT(YEAR({data}),"0000")&TEXT(MONTH({data}),"00")&TEXT(DAY({data}),"00")'),
        ("K-MONTH", f'TEXT(YEAR({data}),"0000")&TEXT(MONTH({data}),"00")'),
    ]
    for position, (name, formula) in enumerate(keys, start=1):
        column = table.ListColumns.Add(position)
        column.Name = name
        column.DataBodyRange.Formula = "=" + formula


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--dal", required=True, help="dd/mm/yyyy")
    ap.add_argument("--al", required=True, help="dd/mm/yyyy")
    ap.add_argument("--out", required=True, help="cartella di lavoro")
    ap.add_argument("--stamp", default=datetime.datetime.now().strftime("%Y%m%d%H%M%S"))
    args = ap.parse_args()

    downloads = Path(args.out) / "downloads"
    suffix = f"{compact(args.dal)}--{compact(args.al)}--{args.stamp}"
    risorse = downloads / f"monitoring-export-risorse-{suffix}.xlsx"
    wbs = downloads / f"monitoring-export-wbs-{suffix}.xlsx"
    consolidato = downloads / f"monitoring-export-{suffix}.xlsx"

    excel = win32com.client.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    try:
        convert(excel, downloads / "risorse-raw.xls", risorse)
        convert(excel, downloads / "wbs-raw.xls", wbs)

        wb = excel.Workbooks.Add()
        add_data_sheet(excel, wb, risorse, "Risorse", "TabRisorse")
        add_data_sheet(excel, wb, wbs, "Wbs", "TabWbs")

        wb.Queries.Add("QueryNormalizzazione", M_QUERY)
        cons = wb.Worksheets.Add(wb.Worksheets(1))
        cons.Name = "Consuntivazione"
        for ws in list(wb.Worksheets):
            if ws.Name.startswith(("Foglio", "Sheet")):
                ws.Delete()
        conn = ("OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;"
                'Location=QueryNormalizzazione;Extended Properties=""')
        table = cons.ListObjects.Add(XL_SRC_EXTERNAL, conn, None, 1, cons.Range("A1"))
        table.QueryTable.CommandType = 2
        table.QueryTable.CommandText = "SELECT * FROM [QueryNormalizzazione]"
        table.QueryTable.Refresh(False)

        add_key_columns(cons.ListObjects(1))

        n_risorse = wb.Worksheets("Risorse").ListObjects(1).ListRows.Count
        n_wbs = wb.Worksheets("Wbs").ListObjects(1).ListRows.Count
        n_cons = cons.ListObjects(1).ListRows.Count
        wb.SaveAs(str(consolidato), XL_OPENXML)
        wb.Close(False)
    finally:
        excel.Quit()

    print(f"Consolidato: {consolidato}")
    print(f"Risorse: {n_risorse}")
    print(f"Wbs: {n_wbs}")
    print(f"Consuntivazione: {n_cons}")
    print(f"Wbs duplicati non aggiunti: {n_risorse + n_wbs - n_cons}")


if __name__ == "__main__":
    main()
