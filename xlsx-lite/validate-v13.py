"""Validate --model-test v1.3 outputs with an independent OOXML reader."""
import pathlib
import sys
import warnings
from zipfile import ZipFile
from xml.etree import ElementTree as ET
import openpyxl

folder = pathlib.Path(sys.argv[1])
expected = {'v13-rows.xlsx': 10, 'v13-cols.xlsx': 10, 'v13-charts.xlsx': 14,
            'v13-chart-deleted.xlsx': 13, 'v13-live.xlsx': 14, 'v13-new.xlsx': 1, 'v13-empty.xlsx': 0}
for name, count in expected.items():
    path = folder / name
    with ZipFile(path) as package:
        assert package.testzip() is None, name
        for member in package.namelist():
            if member.endswith(('.xml', '.rels')) or member == '[Content_Types].xml':
                root = ET.fromstring(package.read(member))
                if member.startswith('xl/worksheets/') and member.endswith('.xml'):
                    for node in root.iter():
                        assert node.tag.startswith('{'), (name, member, 'missing namespace', node.tag)
        assert 'xl/calcChain.xml' not in package.namelist()
        if name not in ('v13-new.xlsx', 'v13-empty.xlsx'):
            assert package.read('customXml/keep.xml') == b'<keep>v13 untouched</keep>'
    with warnings.catch_warnings(record=True) as caught:
        warnings.simplefilter('always')
        workbook = openpyxl.load_workbook(path)
    assert not caught, [str(w.message) for w in caught]
    actual = sum(len(ws._charts) for ws in workbook)
    assert actual == count, (name, actual, count)
    if name == 'v13-rows.xlsx':
        assert workbook['Other']['A1'].value == "=SUM('Data sheet'!$B$2:$B$6)"
        assert workbook['Data sheet'].tables['SalesTable'].ref == 'A1:C6'
    if name == 'v13-cols.xlsx':
        assert workbook['Data sheet'].tables['SalesTable'].ref == 'A1:D4'
        assert len(workbook['Data sheet'].tables['SalesTable'].tableColumns) == 4
        assert workbook['Data sheet']['B1'].value.startswith(('Column', 'Colonna'))
    if name == 'v13-charts.xlsx':
        anchor = workbook['Data sheet']._charts[-1].anchor
        assert anchor._from.col >= 0 and anchor.to.col > anchor._from.col
    if name == 'v13-live.xlsx':
        series = workbook['Data sheet']._charts[0].series[0]
        assert series.val.numRef.numCache.pt[0].v == 99
    print(f'PASS: {name}: XML valid; openpyxl loaded {actual} charts without warnings')
print('ALL V1.3 PACKAGE CHECKS PASSED')
