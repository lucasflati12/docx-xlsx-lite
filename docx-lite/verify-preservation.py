"""Validate smoke-test output packages using only the Python standard library.

Usage: python3 verify-preservation.py /private/tmp/docxlite-fixtures [--python-docx]
"""
import collections
import pathlib
import posixpath
import sys
import xml.etree.ElementTree as ET
import zipfile

folder = pathlib.Path(sys.argv[1])
W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
R = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}'
WP = '{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}'

def signature(node):
    return node.tag, sorted(node.attrib.items()), (node.text or '').strip(), [signature(c) for c in node]

def package(path):
    with zipfile.ZipFile(path) as z:
        assert len(z.namelist()) == len(set(z.namelist())), 'Duplicate ZIP members'
        parts = {n: z.read(n) for n in z.namelist() if not n.endswith('/')}
    xml = {n: ET.fromstring(data) for n, data in parts.items() if n.endswith(('.xml', '.rels'))}
    defaults, overrides = {}, {}
    for node in xml['[Content_Types].xml']:
        if node.tag.endswith('Default'):
            key = node.attrib['Extension'].lower()
            assert key not in defaults, 'Duplicate content type default'
            defaults[key] = node.attrib['ContentType']
        else:
            key = node.attrib['PartName'].lstrip('/')
            assert key not in overrides and key in parts, 'Duplicate or dangling content type override'
            overrides[key] = node.attrib['ContentType']
    for name in parts:
        assert name == '[Content_Types].xml' or name in overrides or name.rsplit('.', 1)[-1].lower() in defaults, name
    for name, root in xml.items():
        if not name.endswith('.rels'):
            continue
        ids = [n.attrib['Id'] for n in root]
        assert len(ids) == len(set(ids)), 'Duplicate relationship IDs: ' + name
        base, filename = name.split('_rels/')
        owner = base + filename[:-5] if filename != '.rels' else None
        for rel in root:
            if rel.attrib.get('TargetMode') == 'External':
                continue
            target = rel.attrib['Target']
            resolved = target.lstrip('/') if target.startswith('/') else posixpath.normpath(posixpath.join(base, target))
            assert resolved in parts, (name, target)
        if owner in xml:
            for node in xml[owner].iter():
                for attr, value in node.attrib.items():
                    if attr.startswith(R):
                        assert value in ids, ('Dangling relationship attribute', owner, value)
    # Drawing IDs must not collide across retained headers and new body drawings.
    drawing_ids = [n.attrib['id'] for root in xml.values() for n in root.iter(WP + 'docPr')]
    assert len(drawing_ids) == len(set(drawing_ids)), 'Duplicate drawing IDs'
    return parts, xml

original_parts, original = package(folder / 'retained.docx')
original_sections = [signature(n) for n in original['word/document.xml'].iter(W + 'sectPr')]
for name, deleted, empty in [('edited', False, False), ('second', False, False), ('deleted', True, False), ('undone', False, False), ('partial', True, False), ('empty', True, True)]:
    parts, xml = package(folder / ('retained-' + name + '.docx'))
    document = xml['word/document.xml']
    assert set(original_parts) <= set(parts), 'Missing original package part'
    changed = {'word/document.xml', 'word/_rels/document.xml.rels', '[Content_Types].xml', 'word/styles.xml', 'word/numbering.xml'}
    if deleted:
        changed |= {'word/comments.xml', 'word/commentsExtended.xml', 'word/commentsIds.xml'}
    if name in ('deleted', 'empty'):
        changed.add('word/footnotes.xml')
    if empty:
        changed.add('word/endnotes.xml')
    for path in original_parts.keys() - changed:
        assert parts[path] == original_parts[path], ('Changed retained part', name, path)
    sections = [signature(n) for n in document.iter(W + 'sectPr')]
    assert sections == (original_sections[-1:] if empty else original_sections), 'Changed section properties'
    for kind, expected in [('commentReference', [] if empty else ['0'] if deleted else ['0', '1']), ('footnoteReference', [] if empty else ['1'] if name == 'deleted' else ['1', '2']), ('endnoteReference', [] if empty else ['1'])]:
        ids = [n.attrib[W + 'id'] for n in document.iter(W + kind)]
        assert collections.Counter(ids) == collections.Counter(expected), (name, kind, ids)
        plural = {'commentReference': 'comments', 'footnoteReference': 'footnotes', 'endnoteReference': 'endnotes'}[kind]
        existing = [n.attrib[W + 'id'] for n in xml['word/' + plural + '.xml']]
        assert len(existing) == len(set(existing)) and set(ids) <= set(existing)
        singular = plural[:-1]
        if name in ('deleted', 'empty') or kind == 'commentReference':
            assert set(existing) - {'-1', '0'} == set(expected) - {'0'}
        if kind == 'commentReference':
            for role in ['commentRangeStart', 'commentRangeEnd']:
                assert [n.attrib[W + 'id'] for n in document.iter(W + role)] == expected
    for role in ['bookmarkStart', 'bookmarkEnd']:
        assert [n.attrib[W + 'id'] for n in document.iter(W + role)] == ([] if empty else ['42'])
    if deleted:
        for path in ['word/commentsExtended.xml', 'word/commentsIds.xml']:
            assert len(xml[path]) == (0 if empty else 1), 'Dangling comment metadata'
    for part, tag, attr in [('styles', 'style', 'styleId'), ('numbering', 'abstractNum', 'abstractNumId'), ('numbering', 'num', 'numId')]:
        nodes = list(xml['word/' + part + '.xml'].iter(W + tag))
        ids = [n.attrib[W + attr] for n in nodes]
        assert len(ids) == len(set(ids)), ('Duplicate merged IDs', part, tag)
        originals = list(original['word/' + part + '.xml'].iter(W + tag))
        assert all(any(signature(n) == signature(old) for n in nodes) for old in originals), 'Changed original definition'
    if name == 'edited':
        merged_counts = [len(xml['word/' + part + '.xml']) for part in ['styles', 'numbering']]
    if name == 'second':
        assert merged_counts == [len(xml['word/' + part + '.xml']) for part in ['styles', 'numbering']], 'Repeated save grows merged definitions'
    nums = xml['word/numbering.xml']
    kinds = [n.tag for n in nums]
    assert max(i for i, tag in enumerate(kinds) if tag == W + 'abstractNum') < min(i for i, tag in enumerate(kinds) if tag == W + 'num')
    for tag, definitions, idattr in [('numId', 'num', 'numId'), ('abstractNumId', 'abstractNum', 'abstractNumId')]:
        ids = {n.attrib[W + idattr] for n in nums.iter(W + definitions)}
        for root in [document, nums]:
            assert all(n.attrib[W + 'val'] in ids or n.attrib[W + 'val'] == '0' for n in root.iter(W + tag))
    if not empty:
        assert 'Edited elsewhere.' in ''.join(document.itertext())
        assert len(list(document.iter(W + 'drawing'))) == 1
        assert not any('DOCXLITEKEEP' in (n.text or '') for n in document.iter())
    if '--python-docx' in sys.argv:
        try:
            from docx import Document
        except ImportError:
            print('SKIP: optional python-docx is not installed')
        else:
            doc = Document(folder / ('retained-' + name + '.docx'))
            assert len(doc.sections) == (1 if empty else 2)
    print('PASS: retained-' + name + ': parts, XML, rels, content types, sections, annotations, IDs')
print('ALL PACKAGE PRESERVATION CHECKS PASSED')
