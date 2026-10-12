"""Click-and-type fields, click boxes and write-on lines for the forms
templates, written straight into a python-docx document's XML.

Each field's gray hint is also stored where Word keeps its own (a
placeholder entry in the file's glossary part), so a field someone
clears shows its hint again. Call finish(doc) once, just before saving.
"""
import uuid
from xml.sax.saxutils import escape

from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.opc.packuri import PackURI
from docx.opc.part import Part
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls

W14 = 'xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"'
GLOSSARY = ('application/vnd.openxmlformats-officedocument.'
            'wordprocessingml.document.glossary+xml')
BOX_FONT = 'Segoe UI Symbol'

# Word wants a paragraph's settings in a fixed order; a border goes
# before any of these.
_AFTER_BORDER = ['w:' + t for t in (
    'shd', 'tabs', 'suppressAutoHyphens', 'kinsoku', 'wordWrap',
    'overflowPunct', 'topLinePunct', 'autoSpaceDE', 'autoSpaceDN', 'bidi',
    'adjustRightInd', 'snapToGrid', 'spacing', 'ind', 'contextualSpacing',
    'mirrorIndents', 'suppressOverlap', 'jc', 'textDirection',
    'textAlignment', 'textboxTightWrap', 'outlineLvl', 'divId', 'cnfStyle',
    'rPr', 'sectPr', 'pPrChange')]

_hints = []     # (placeholder name, hint) for finish()
_next_id = [100]


def _new_id():
    _next_id[0] += 1
    return _next_id[0]


def field(paragraph, hint, title=None):
    """Append a click-and-type field (plain text content control)."""
    ident = _new_id()
    name = uuid.uuid5(uuid.NAMESPACE_URL,
                      'forms-templates/%d/%s' % (ident, hint)).hex.upper()
    _hints.append((name, hint))
    xml = (
        '<w:sdt {ns}><w:sdtPr>'
        '<w:alias w:val="{title}"/><w:id w:val="{id}"/>'
        '<w:placeholder><w:docPart w:val="{name}"/></w:placeholder>'
        '<w:showingPlcHdr/><w:text/>'
        '</w:sdtPr><w:sdtContent><w:r>'
        '<w:rPr><w:rStyle w:val="PlaceholderText"/></w:rPr>'
        '<w:t xml:space="preserve">{hint}</w:t>'
        '</w:r></w:sdtContent></w:sdt>'
    ).format(ns=nsdecls('w'), title=escape(title or hint, {'"': '&quot;'}),
             id=ident, name=name, hint=escape(hint))
    paragraph._p.append(parse_xml(xml))


def checkbox(paragraph, title):
    """Append a box that checks and unchecks with a click."""
    # The look is set twice: on the box as drawn, and on the control so a
    # click that swaps the symbol keeps the same size.
    look = ('<w:rPr><w:rFonts w:ascii="{f}" w:eastAsia="{f}" w:hAnsi="{f}" '
            'w:cs="{f}"/><w:sz w:val="24"/></w:rPr>').format(f=BOX_FONT)
    xml = (
        '<w:sdt {ns} {w14}><w:sdtPr>{look}'
        '<w:alias w:val="{title}"/><w:id w:val="{id}"/>'
        '<w14:checkbox><w14:checked w14:val="0"/>'
        '<w14:checkedState w14:val="2612" w14:font="{f}"/>'
        '<w14:uncheckedState w14:val="2610" w14:font="{f}"/>'
        '</w14:checkbox>'
        '</w:sdtPr><w:sdtContent><w:r>{look}'
        '<w:t>☐</w:t></w:r></w:sdtContent></w:sdt>'
    ).format(ns=nsdecls('w'), w14=W14, look=look, f=BOX_FONT,
             title=escape(title, {'"': '&quot;'}), id=_new_id())
    paragraph._p.append(parse_xml(xml))


def line_under(paragraph):
    """Draw a thin line under the paragraph, the line to write on."""
    border = parse_xml(
        '<w:pBdr {}><w:bottom w:val="single" w:sz="4" w:space="1" '
        'w:color="000000"/></w:pBdr>'.format(nsdecls('w')))
    paragraph._p.get_or_add_pPr().insert_element_before(border, *_AFTER_BORDER)


def finish(doc):
    """Store every field's hint as a Word placeholder. Call before saving."""
    entries = ''.join(
        '<w:docPart><w:docPartPr><w:name w:val="{name}"/>'
        '<w:category><w:name w:val="General"/>'
        '<w:gallery w:val="placeholder"/></w:category>'
        '<w:types><w:type w:val="bbPlcHdr"/></w:types>'
        '<w:behaviors><w:behavior w:val="content"/></w:behaviors>'
        '<w:guid w:val="{{{guid}}}"/>'
        '</w:docPartPr><w:docPartBody><w:p><w:r>'
        '<w:rPr><w:rStyle w:val="PlaceholderText"/></w:rPr>'
        '<w:t xml:space="preserve">{hint}</w:t>'
        '</w:r></w:p></w:docPartBody></w:docPart>'.format(
            name=name, guid=str(uuid.UUID(name)).upper(), hint=escape(hint))
        for name, hint in _hints)
    xml = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
           '<w:glossaryDocument {ns}><w:docParts>{entries}</w:docParts>'
           '</w:glossaryDocument>').format(ns=nsdecls('w'), entries=entries)
    part = Part(PackURI('/word/glossary/document.xml'), GLOSSARY,
                xml.encode('utf-8'), doc.part.package)
    doc.part.relate_to(part, RT.GLOSSARY_DOCUMENT)
