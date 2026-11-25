"""Create the 20,000-row test workbook used by SheetTests.swift (needs: pip install openpyxl)."""
import datetime, pathlib, random, sys
import openpyxl

out = pathlib.Path(sys.argv[1]); out.mkdir(parents=True, exist_ok=True)
random.seed(7)
wb = openpyxl.Workbook(); ws = wb.active; ws.title = "Dati"
ws.append(["ID", "Nome", "Città", "Importo", "Data", "Percentuale", "Nota"])
for i in range(1, 20001):
    ws.append([i, f"Cliente {i}", random.choice(["Roma", "Milano", "Napoli"]), round(random.random() * 1000, 2),
               datetime.date(2026, 1, 1) + datetime.timedelta(days=i % 300), random.random(), "testo lungo " * 3])
for row in ws.iter_rows(min_row=2, min_col=4, max_col=6):
    row[0].number_format = '#,##0.00 €'; row[2].number_format = '0.0%'
ws.auto_filter.ref = "A1:G20001"; ws.freeze_panes = "A2"
summary = wb.create_sheet("Riepilogo"); summary["A1"] = "Totale"; summary["B1"] = "=SUM(Dati!D2:D20001)"
summary.merge_cells("A3:C3"); summary["A3"] = "Unito"
wb.save(out / "big.xlsx")
print(out / "big.xlsx")
