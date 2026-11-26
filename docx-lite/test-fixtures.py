"""Create small, self-contained DOCX fixtures without third-party dependencies."""
import struct
import zlib
import pathlib
import sys
import zipfile

out = pathlib.Path(sys.argv[1])
out.mkdir(parents=True, exist_ok=True)
content_types = '''<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/>
<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
</Types>'''
root_rels = '''<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>'''
rels = '''<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/test.png"/>
<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.com" TargetMode="External"/>
</Relationships>'''
styles = '''<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Helvetica" w:hAnsi="Helvetica"/><w:sz w:val="24"/></w:rPr></w:rPrDefault></w:docDefaults></w:styles>'''
picture = '''<w:p><w:r><w:drawing><wp:inline><wp:extent cx="457200" cy="457200"/><wp:docPr id="1" name="Test image"/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:nvPicPr><pic:cNvPr id="0" name="test.png"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="rId2"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="457200" cy="457200"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>'''
body = '''<w:p><w:r><w:rPr><w:b/><w:sz w:val="40"/></w:rPr><w:t>Lettura essenziale</w:t></w:r></w:p>
<w:p><w:r><w:t>Caffè, CAFFE e caffè. Accenti, emoji 🌈 e testo Unicode.</w:t></w:r></w:p>
<w:p><w:r><w:rPr><w:i/></w:rPr><w:t>Apri, leggi, cerca. Tutto qui.</w:t></w:r></w:p>
<w:p><w:hyperlink r:id="rId3"><w:r><w:t>Un collegamento di prova</w:t></w:r></w:hyperlink></w:p>
<w:tbl><w:tblPr><w:tblW w:w="5000" w:type="pct"/><w:tblBorders><w:top w:val="single" w:sz="4"/><w:left w:val="single" w:sz="4"/><w:bottom w:val="single" w:sz="4"/><w:right w:val="single" w:sz="4"/><w:insideH w:val="single" w:sz="4"/><w:insideV w:val="single" w:sz="4"/></w:tblBorders></w:tblPr><w:tblGrid><w:gridCol w:w="4500"/><w:gridCol w:w="4500"/></w:tblGrid>
<w:tr><w:tc><w:p><w:r><w:t>Nome</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>Valore</w:t></w:r></w:p></w:tc></w:tr>
<w:tr><w:tc><w:p><w:r><w:t>Documento</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>Leggerissimo</w:t></w:r></w:p></w:tc></w:tr></w:tbl>'''
body += picture
body += ''.join(f'<w:p><w:r><w:t>Paragrafo {i}. Un lettore piccolo, rapido e semplice per i documenti di ogni giorno.</w:t></w:r></w:p>' for i in range(1, 46))
namespaces = '''xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"'''
section = '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1134" w:right="1134" w:bottom="1134" w:left="1134"/></w:sectPr>'
# A valid PNG with checked CRCs; the previous literal had a corrupt IDAT checksum.
def chunk(kind, data):
    return struct.pack('>I', len(data)) + kind + data + struct.pack('>I', zlib.crc32(kind + data))
pixel = (b'\x89PNG\r\n\x1a\n'
         + chunk(b'IHDR', struct.pack('>IIBBBBB', 32, 24, 8, 6, 0, 0, 0))
         + chunk(b'IDAT', zlib.compress((b'\x00' + b'\x30\x80\xe0\xff' * 32) * 24))
         + chunk(b'IEND', b''))
(out / 'image.png').write_bytes(pixel)
(out / 'legacy.rtf').write_bytes(b'{\\rtf1\\ansi Legacy \\b bold\\b0 document.}')

for name, paragraphs in [('rich.docx', body), ('empty.docx', '<w:p/>')]:
    with zipfile.ZipFile(out / name, 'w', zipfile.ZIP_DEFLATED) as z:
        z.writestr('[Content_Types].xml', content_types)
        z.writestr('_rels/.rels', root_rels)
        z.writestr('word/document.xml', f'<?xml version="1.0"?><w:document {namespaces}><w:body>{paragraphs}{section}</w:body></w:document>')
        z.writestr('word/styles.xml', styles)
        z.writestr('word/_rels/document.xml.rels', rels)
        z.writestr('word/media/test.png', pixel)
(out / 'broken.docx').write_bytes(b'This is not a DOCX document.')
print(out)

# Package-preservation fixture: no third-party dependencies, real OOXML parts.
from xml.sax.saxutils import escape
rel_ns = 'http://schemas.openxmlformats.org/package/2006/relationships'
office_rel = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/'
w_ns = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
w14 = 'http://schemas.microsoft.com/office/word/2010/wordml'
w15 = 'http://schemas.microsoft.com/office/word/2012/wordml'
w16cid = 'http://schemas.microsoft.com/office/word/2016/wordml/cid'


def relationship_file(items):
    return '<Relationships xmlns="' + rel_ns + '">' + ''.join(
        f'<Relationship Id="{id}" Type="{type}" Target="{target}"{mode}/>'
        for id, type, target, mode in items) + '</Relationships>'


def paragraph(text):
    return '<w:p><w:r><w:t>' + escape(text) + '</w:t></w:r></w:p>'


retained = {}
header_refs = ''.join(f'<w:headerReference w:type="{kind}" r:id="header{kind}"/>' for kind in ['default', 'first', 'even'])
footer_refs = ''.join(f'<w:footerReference w:type="{kind}" r:id="footer{kind}"/>' for kind in ['default', 'first', 'even'])
section_one = (f'<w:sectPr>{header_refs}{footer_refs}<w:type w:val="nextPage"/>'
               '<w:pgSz w:w="16838" w:h="11906" w:orient="landscape"/>'
               '<w:pgMar w:top="1000" w:right="1200" w:bottom="1400" w:left="720" w:header="360" w:footer="420" w:gutter="100"/>'
               '<w:pgNumType w:fmt="decimal" w:start="3"/><w:cols w:num="2" w:space="400" w:sep="1"/><w:titlePg/></w:sectPr>')
section_two = (f'<w:sectPr>{header_refs}{footer_refs}<w:type w:val="continuous"/>'
               '<w:pgSz w:w="16838" w:h="11906" w:orient="landscape"/>'
               '<w:pgMar w:top="1100" w:right="1300" w:bottom="1500" w:left="900" w:header="400" w:footer="500" w:gutter="0"/>'
               '<w:pgNumType w:fmt="lowerRoman" w:start="7"/><w:cols w:num="1" w:space="720"/></w:sectPr>')
annotation_body = (paragraph('Text before annotations.')
    + '<w:p><w:r><w:t xml:space="preserve">A </w:t></w:r><w:commentRangeStart w:id="0"/>'
      '<w:r><w:rPr><w:b/></w:rPr><w:t>bold comment</w:t></w:r><w:r><w:t xml:space="preserve"> span</w:t></w:r>'
      '<w:commentRangeEnd w:id="0"/><w:r><w:rPr><w:rStyle w:val="CommentReference"/></w:rPr><w:commentReference w:id="0"/></w:r>'
      '<w:r><w:t xml:space="preserve"> and note</w:t></w:r><w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:footnoteReference w:id="1"/></w:r><w:r><w:t>.</w:t></w:r></w:p>'
    + '<w:p><w:r><w:t xml:space="preserve">Delete annotated paragraph </w:t></w:r><w:commentRangeStart w:id="1"/>'
      '<w:r><w:t>comment two</w:t></w:r><w:commentRangeEnd w:id="1"/><w:r><w:commentReference w:id="1"/></w:r>'
      '<w:r><w:t xml:space="preserve"> note</w:t></w:r><w:r><w:footnoteReference w:id="2"/></w:r><w:r><w:t>.</w:t></w:r></w:p>'
    + '<w:p><w:bookmarkStart w:id="42" w:name="RetainedBookmark"/><w:r><w:t>Bookmarked text</w:t></w:r><w:bookmarkEnd w:id="42"/></w:p>'
    + '<w:p><w:pPr>' + section_one + '</w:pPr><w:r><w:t>Section one ends.</w:t></w:r></w:p>'
    + '<w:p><w:r><w:t>Second section has endnote</w:t></w:r><w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:endnoteReference w:id="1"/></w:r></w:p>'
    + paragraph('Editable tail.'))
retained['word/document.xml'] = f'<w:document {namespaces} xmlns:w14="{w14}"><w:body>{annotation_body}{section_two}</w:body></w:document>'
relationships = [('styles', office_rel + 'styles', 'styles.xml', ''), ('numbering', office_rel + 'numbering', 'numbering.xml', '')]
for kind, filename in [('comments', 'comments.xml'), ('footnotes', 'footnotes.xml'), ('endnotes', 'endnotes.xml'), ('settings', 'settings.xml'), ('theme', 'theme/theme1.xml'), ('fontTable', 'fontTable.xml')]:
    relationships.append((kind, office_rel + kind, filename, ''))
for kind, uri in [('commentsExtended', 'http://schemas.microsoft.com/office/2011/relationships/commentsExtended'), ('people', 'http://schemas.microsoft.com/office/2011/relationships/people'), ('commentsIds', 'http://schemas.microsoft.com/office/2016/09/relationships/commentsIds')]:
    relationships.append((kind, uri, kind + '.xml', ''))
for kind in ['default', 'first', 'even']:
    for part in ['header', 'footer']:
        relationships.append((part + kind, office_rel + part, part + '-' + kind + '.xml', ''))
relationships.append(('custom', office_rel + 'customXml', '../customXml/item1.xml', ''))
retained['word/_rels/document.xml.rels'] = relationship_file(relationships)
for kind in ['default', 'first', 'even']:
    retained['word/header-' + kind + '.xml'] = f'<w:hdr {namespaces}>' + paragraph(kind.title() + ' page header') + (picture.replace('rId2', 'logo') if kind == 'default' else '') + '</w:hdr>'
    retained['word/footer-' + kind + '.xml'] = f'<w:ftr xmlns:w="{w_ns}"><w:p><w:r><w:t xml:space="preserve">{kind.title()} page footer </w:t></w:r><w:fldSimple w:instr="PAGE"><w:r><w:t>3</w:t></w:r></w:fldSimple></w:p></w:ftr>'
retained['word/_rels/header-default.xml.rels'] = relationship_file([('logo', office_rel + 'image', 'media/docxlite-0.png', '')])
# Deliberate collision with the editor's usual first generated image filename.
retained['word/media/docxlite-0.png'] = (b'\x89PNG\r\n\x1a\n' + chunk(b'IHDR', struct.pack('>IIBBBBB', 8, 8, 8, 6, 0, 0, 0))
    + chunk(b'IDAT', zlib.compress((b'\x00' + b'\xe0\x20\x30\xff' * 8) * 8)) + chunk(b'IEND', b''))
retained['word/comments.xml'] = (f'<w:comments xmlns:w="{w_ns}" xmlns:w14="{w14}">'
    '<w:comment w:id="0" w:author="Reviewer" w:date="2026-11-25T12:00:00Z"><w:p w14:paraId="000000A0"><w:r><w:t>First comment on bold text.</w:t></w:r></w:p></w:comment>'
    '<w:comment w:id="1" w:author="Reviewer" w:date="2026-11-25T12:01:00Z"><w:p w14:paraId="000000A1"><w:r><w:t>Second comment to delete.</w:t></w:r></w:p></w:comment></w:comments>')
retained['word/commentsExtended.xml'] = f'<w15:commentsEx xmlns:w15="{w15}"><w15:commentEx w15:paraId="000000A0" w15:done="0"/><w15:commentEx w15:paraId="000000A1" w15:done="0"/></w15:commentsEx>'
retained['word/commentsIds.xml'] = f'<w16cid:commentsIds xmlns:w16cid="{w16cid}"><w16cid:commentId w16cid:paraId="000000A0" w16cid:durableId="00000100"/><w16cid:commentId w16cid:paraId="000000A1" w16cid:durableId="00000101"/></w16cid:commentsIds>'
retained['word/people.xml'] = f'<w15:people xmlns:w15="{w15}"><w15:person w15:author="Reviewer"><w15:presenceInfo w15:providerId="None" w15:userId="reviewer@example.test"/></w15:person></w15:people>'
for kind in ['footnotes', 'endnotes']:
    singular = kind[:-1]
    retained['word/' + kind + '.xml'] = (f'<w:{kind} xmlns:w="{w_ns}">'
        f'<w:{singular} w:id="-1" w:type="separator"><w:p><w:r><w:separator/></w:r></w:p></w:{singular}>'
        f'<w:{singular} w:id="0" w:type="continuationSeparator"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:{singular}>'
        f'<w:{singular} w:id="1">' + paragraph('First footnote retained.' if kind == 'footnotes' else 'Endnote retained.') + f'</w:{singular}>'
        + (f'<w:{singular} w:id="2">' + paragraph('Second footnote to delete.') + f'</w:{singular}>' if kind == 'footnotes' else '') + f'</w:{kind}>')
retained['word/settings.xml'] = f'<w:settings xmlns:w="{w_ns}"><w:evenAndOddHeaders/><w:updateFields w:val="true"/><w:compat/></w:settings>'
retained['word/fontTable.xml'] = f'<w:fonts xmlns:w="{w_ns}"><w:font w:name="Arial"><w:family w:val="swiss"/><w:charset w:val="00"/></w:font></w:fonts>'
retained['word/styles.xml'] = (f'<w:styles xmlns:w="{w_ns}"><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="24"/></w:rPr></w:style>'
    '<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/><w:pPr><w:outlineLvl w:val="0"/></w:pPr><w:rPr><w:color w:val="336699"/><w:sz w:val="28"/></w:rPr></w:style>'
    + ''.join(f'<w:style w:type="character" w:styleId="{name}"><w:name w:val="{name}"/><w:rPr><w:vertAlign w:val="superscript"/></w:rPr></w:style>' for name in ['FootnoteReference', 'EndnoteReference', 'CommentReference']) + '</w:styles>')
retained['word/numbering.xml'] = f'<w:numbering xmlns:w="{w_ns}"><w:abstractNum w:abstractNumId="1"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="lowerRoman"/><w:lvlText w:val="%1."/></w:lvl></w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="1"/></w:num></w:numbering>'
colors = [('dk1', '000000'), ('lt1', 'FFFFFF'), ('dk2', '222222'), ('lt2', 'EEEEEE')] + [(f'accent{i}', '4472C4') for i in range(1, 7)] + [('hlink', '0563C1'), ('folHlink', '954F72')]
retained['word/theme/theme1.xml'] = ('<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Retained theme"><a:themeElements><a:clrScheme name="Retained">'
    + ''.join(f'<a:{name}><a:srgbClr val="{color}"/></a:{name}>' for name, color in colors) + '</a:clrScheme><a:fontScheme name="Arial">'
    + ''.join(f'<a:{kind}Font><a:latin typeface="Arial"/><a:ea typeface=""/><a:cs typeface=""/></a:{kind}Font>' for kind in ['major', 'minor'])
    + '</a:fontScheme><a:fmtScheme name="Retained"><a:fillStyleLst>' + '<a:solidFill><a:schemeClr val="accent1"/></a:solidFill>' * 3 + '</a:fillStyleLst><a:lnStyleLst>'
    + '<a:ln w="9525"><a:solidFill><a:schemeClr val="accent1"/></a:solidFill><a:prstDash val="solid"/></a:ln>' * 3 + '</a:lnStyleLst><a:effectStyleLst>'
    + '<a:effectStyle><a:effectLst/></a:effectStyle>' * 3 + '</a:effectStyleLst><a:bgFillStyleLst>' + '<a:solidFill><a:schemeClr val="lt1"/></a:solidFill>' * 3 + '</a:bgFillStyleLst></a:fmtScheme></a:themeElements></a:theme>')
retained['customXml/item1.xml'] = '<retained xmlns="urn:docx-lite:test">Custom XML must stay unchanged</retained>'
retained['customXml/itemProps1.xml'] = '<ds:datastoreItem xmlns:ds="http://schemas.openxmlformats.org/officeDocument/2006/customXml" ds:itemID="{00000000-0000-0000-0000-000000000001}"><ds:schemaRefs/></ds:datastoreItem>'
retained['customXml/_rels/item1.xml.rels'] = relationship_file([('properties', office_rel + 'customXmlProps', 'itemProps1.xml', '')])
retained['docProps/core.xml'] = '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/"><dc:title>Retained title</dc:title><dc:creator>Fixture</dc:creator></cp:coreProperties>'
retained['docProps/app.xml'] = '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"><Application>Fixture generator</Application><Pages>2</Pages></Properties>'
retained['docProps/custom.xml'] = '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"><property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="2" name="Retained"><vt:lpwstr>Original property</vt:lpwstr></property></Properties>'
retained['_rels/.rels'] = relationship_file([('document', office_rel + 'officeDocument', 'word/document.xml', ''), ('core', 'http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties', 'docProps/core.xml', ''), ('app', office_rel + 'extended-properties', 'docProps/app.xml', ''), ('custom', office_rel + 'custom-properties', 'docProps/custom.xml', '')])
ct = {'word/document.xml': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml', 'customXml/itemProps1.xml': 'application/vnd.openxmlformats-officedocument.customXmlProperties+xml'}
for name in ['styles', 'numbering', 'comments', 'footnotes', 'endnotes', 'settings', 'fontTable']:
    ct['word/' + name + '.xml'] = 'application/vnd.openxmlformats-officedocument.wordprocessingml.' + name + '+xml'
for kind in ['default', 'first', 'even']:
    for name in ['header', 'footer']:
        ct[f'word/{name}-{kind}.xml'] = f'application/vnd.openxmlformats-officedocument.wordprocessingml.{name}+xml'
ct.update({'word/commentsExtended.xml': 'application/vnd.ms-word.commentsExtended+xml', 'word/commentsIds.xml': 'application/vnd.ms-word.commentsIds+xml', 'word/people.xml': 'application/vnd.ms-word.people+xml', 'word/theme/theme1.xml': 'application/vnd.openxmlformats-officedocument.theme+xml', 'docProps/core.xml': 'application/vnd.openxmlformats-package.core-properties+xml', 'docProps/app.xml': 'application/vnd.openxmlformats-officedocument.extended-properties+xml', 'docProps/custom.xml': 'application/vnd.openxmlformats-officedocument.custom-properties+xml'})
retained['[Content_Types].xml'] = '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="png" ContentType="image/png"/>' + ''.join(f'<Override PartName="/{path}" ContentType="{type}"/>' for path, type in ct.items()) + '</Types>'
for filename, unsupported in [('retained.docx', False), ('unsupported.docx', True)]:
    with zipfile.ZipFile(out / filename, 'w', zipfile.ZIP_DEFLATED) as archive:
        for path, data in retained.items():
            if unsupported and path == 'word/document.xml':
                extra = '<w:p><w:ins w:id="7" w:author="Fixture"><w:r><w:t>Tracked insertion</w:t></w:r></w:ins><w:sdt><w:sdtPr/><w:sdtContent><w:r><w:t>Content control</w:t></w:r></w:sdtContent></w:sdt><m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"><m:r><m:t>x</m:t></m:r></m:oMath></w:p>'
                data = data.replace(section_two, extra + section_two)
            archive.writestr(path, data)
print('retained.docx, unsupported.docx')

# styles.docx: formatting that lives only in styles.xml (Title, Heading 1, a list style, a bold paragraph
# style, a character style) plus switched-off toggles (<w:b w:val="0"/>), which AppKit's importer mishandles.
W_NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
def styled_p(text, style=None, run_props=''):
    ppr = f'<w:pPr><w:pStyle w:val="{style}"/></w:pPr>' if style else ''
    rpr = f'<w:rPr>{run_props}</w:rPr>' if run_props else ''
    return f'<w:p>{ppr}<w:r>{rpr}<w:t xml:space="preserve">{text}</w:t></w:r></w:p>'
styled_body = ''.join([
    styled_p('Styled title', 'Title'), styled_p('Styled heading', 'Heading1'), styled_p('Styled bullet', 'ListBullet'),
    styled_p('Not bold explicitly', None, '<w:b w:val="0"/>'), styled_p('Not italic explicitly', None, '<w:i w:val="false"/>'),
    '<w:p><w:pPr><w:pStyle w:val="BoldPara"/></w:pPr><w:r><w:rPr><w:b w:val="0"/></w:rPr><w:t xml:space="preserve">Override not bold </w:t></w:r><w:r><w:t>Inherited bold</w:t></w:r></w:p>',
    styled_p('Character style italic', None, '<w:rStyle w:val="Emphasis"/>'),
])
styled_styles = (f'<w:styles {W_NS}><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Helvetica" w:hAnsi="Helvetica"/><w:sz w:val="24"/></w:rPr></w:rPrDefault></w:docDefaults>'
    '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>'
    '<w:style w:type="paragraph" w:styleId="Title"><w:name w:val="Title"/><w:basedOn w:val="Normal"/><w:rPr><w:b/><w:color w:val="1F4EA0"/><w:sz w:val="56"/></w:rPr></w:style>'
    '<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/><w:pPr><w:keepNext/></w:pPr><w:rPr><w:b/><w:sz w:val="32"/></w:rPr></w:style>'
    '<w:style w:type="paragraph" w:styleId="ListBullet"><w:name w:val="List Bullet"/><w:basedOn w:val="Normal"/><w:pPr><w:numPr><w:numId w:val="1"/></w:numPr></w:pPr></w:style>'
    '<w:style w:type="paragraph" w:styleId="BoldPara"><w:name w:val="Bold Para"/><w:basedOn w:val="Normal"/><w:rPr><w:b/></w:rPr></w:style>'
    '<w:style w:type="character" w:styleId="Emphasis"><w:name w:val="Emphasis"/><w:rPr><w:i/></w:rPr></w:style></w:styles>')
styled_numbering = (f'<w:numbering {W_NS}><w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val="•"/></w:lvl></w:abstractNum>'
    '<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num></w:numbering>')
with zipfile.ZipFile(out / 'styles.docx', 'w', zipfile.ZIP_DEFLATED) as z:
    z.writestr('[Content_Types].xml', '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>'
               '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
               '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>'
               '<Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/></Types>')
    z.writestr('_rels/.rels', root_rels)
    z.writestr('word/_rels/document.xml.rels', '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/></Relationships>')
    z.writestr('word/document.xml', f'<w:document {W_NS}><w:body>{styled_body}<w:sectPr/></w:body></w:document>')
    z.writestr('word/styles.xml', styled_styles)
    z.writestr('word/numbering.xml', styled_numbering)
print('styles.docx')
