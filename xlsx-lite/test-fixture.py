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

# A small deterministic workbook for edge cases that should not change the big-file benchmark.
from zipfile import ZipFile, ZIP_DEFLATED
from xml.etree import ElementTree as ET
small = openpyxl.Workbook(); table = small.active; table.title = 'Sort'
for row in [('Key', 'Payload'), (3, 'three'), ('beta', 'text'), (1, 'one'), (None, 'empty'), ('Alpha', 'capital')]:
    table.append(row)
table.auto_filter.ref = 'A1:B6'
plain = small.create_sheet('Blocks')
for row in [('First', 'Value'), ('z', 2), ('a', 1), (None, None), ('Second', 'Value'), ('y', 4), ('b', 3)]:
    plain.append(row)
errors = small.create_sheet('Errors'); errors['A1'] = '=UNSUPPORTED(1)'; errors['A2'] = '=SUM(C1,1)'
errors['B1'] = '#N/A'; errors['B2'] = '#VALUE!'; errors['B3'] = '#DIV/0!'; errors['B4'] = '#REF!'
errors['C1'] = 123; errors['C1'].number_format = '0.00%'
errors['C2'].number_format = '0.00%'; errors['C2'].fill = openpyxl.styles.PatternFill('solid', fgColor='FFFFCC')
errors['D1'] = '=IF(FALSE,1,"")'
merged = small.create_sheet('Merged'); merged.append(['Header', 'Value']); merged.append([2, 'two']); merged.append([1, 'one']); merged.merge_cells('B2:C2')
protected = small.create_sheet('Protected'); protected['A1'] = 'locked'; protected.protection.sheet = True
small.save(out / 'features.xlsx')
# Give both a known and an unknown formula cached values, like a file last saved by Excel.
path = out / 'features.xlsx'
with ZipFile(path) as archive:
    parts = {name: archive.read(name) for name in archive.namelist()}
ns = {'s': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
ET.register_namespace('', ns['s'])
root = ET.fromstring(parts['xl/worksheets/sheet3.xml'])
for cell in root.findall('.//s:c', ns):
    if cell.attrib.get('r') in ('A1', 'A2'):
        cell.find('s:v', ns).text = '123' if cell.attrib['r'] == 'A1' else '124'
parts['xl/worksheets/sheet3.xml'] = ET.tostring(root, encoding='utf-8', xml_declaration=True)
parts['customXml/keep.xml'] = b'<preserved>untouched package part</preserved>'
with ZipFile(path, 'w', ZIP_DEFLATED) as archive:
    for name, data in parts.items():
        archive.writestr(name, data)
print(path)

# Valid workbook without a styles part: formatting must add its relationship and content type.
bare = openpyxl.Workbook(); bare.active['A1'] = 123; bare.save(out / 'styleless.xlsx')
with ZipFile(out / 'styleless.xlsx') as archive:
    bare_parts = {name: archive.read(name) for name in archive.namelist() if name != 'xl/styles.xml'}
for member, namespace, unwanted, attribute, suffix in [
    ('xl/_rels/workbook.xml.rels', 'http://schemas.openxmlformats.org/package/2006/relationships', 'Relationship', 'Type', '/styles'),
    ('[Content_Types].xml', 'http://schemas.openxmlformats.org/package/2006/content-types', 'Override', 'PartName', '/xl/styles.xml'),
]:
    tree = ET.fromstring(bare_parts[member])
    for child in list(tree):
        if child.tag.endswith('}' + unwanted) and child.attrib.get(attribute, '').endswith(suffix):
            tree.remove(child)
    ET.register_namespace('', namespace)
    bare_parts[member] = ET.tostring(tree, encoding='utf-8', xml_declaration=True)
with ZipFile(out / 'styleless.xlsx', 'w', ZIP_DEFLATED) as archive:
    for name, data in bare_parts.items():
        archive.writestr(name, data)
print(out / 'styleless.xlsx')
