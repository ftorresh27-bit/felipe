#!/usr/bin/env python3
"""
Conversor Markdown -> .docx en Python puro (sin dependencias externas).
Soporta: encabezados (#..####), parrafos, **negrita**, `codigo`, listas con
vinetas (-), listas numeradas (1.), citas (>), reglas (---) y tablas (| .. |).

Uso:
    python3 build_docx.py <entrada.md> <salida.docx>
"""
import sys
import re
import zipfile
from xml.sax.saxutils import escape

NSW = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'


def esc(text):
    return escape(text, {'"': '&quot;'})


# ---------- Inline parsing: **bold**, `code` ----------
def parse_runs(text):
    """Devuelve lista de (texto, {bold, mono}) preservando formato inline."""
    runs = []
    # Tokeniza por **...** y `...`
    pattern = re.compile(r'(\*\*.+?\*\*|`.+?`)')
    pos = 0
    for m in pattern.finditer(text):
        if m.start() > pos:
            runs.append((text[pos:m.start()], {}))
        tok = m.group(0)
        if tok.startswith('**'):
            runs.append((tok[2:-2], {'bold': True}))
        elif tok.startswith('`'):
            runs.append((tok[1:-1], {'mono': True}))
        pos = m.end()
    if pos < len(text):
        runs.append((text[pos:], {}))
    return runs or [(text, {})]


def run_xml(text, bold=False, mono=False, color=None, sz=None):
    rpr = []
    if bold:
        rpr.append('<w:b/>')
    if mono:
        rpr.append('<w:rFonts w:ascii="Consolas" w:hAnsi="Consolas"/>')
    if color:
        rpr.append(f'<w:color w:val="{color}"/>')
    if sz:
        rpr.append(f'<w:sz w:val="{sz}"/>')
    rpr_xml = f'<w:rPr>{"".join(rpr)}</w:rPr>' if rpr else ''
    # xml:space preserve para no perder espacios
    return (f'<w:r>{rpr_xml}'
            f'<w:t xml:space="preserve">{esc(text)}</w:t></w:r>')


def runs_to_xml(text, extra=None):
    extra = dict(extra or {})
    extra_bold = extra.pop('bold', False)
    extra_mono = extra.pop('mono', False)
    out = []
    for t, fmt in parse_runs(text):
        out.append(run_xml(t,
                           bold=fmt.get('bold', False) or extra_bold,
                           mono=fmt.get('mono', False) or extra_mono,
                           **extra))
    return ''.join(out)


def para(text='', style=None, runs_extra=None, jc=None, spacing_after=None):
    ppr = []
    if style:
        ppr.append(f'<w:pStyle w:val="{style}"/>')
    if jc:
        ppr.append(f'<w:jc w:val="{jc}"/>')
    if spacing_after is not None:
        ppr.append(f'<w:spacing w:after="{spacing_after}"/>')
    ppr_xml = f'<w:pPr>{"".join(ppr)}</w:pPr>' if ppr else ''
    body = runs_to_xml(text, runs_extra) if text else ''
    return f'<w:p>{ppr_xml}{body}</w:p>'


def list_para(text, ordered=False, level=0):
    num_id = 2 if ordered else 1
    ppr = (f'<w:pPr><w:pStyle w:val="ListParagraph"/>'
           f'<w:numPr><w:ilvl w:val="{level}"/>'
           f'<w:numId w:val="{num_id}"/></w:numPr></w:pPr>')
    return f'<w:p>{ppr}{runs_to_xml(text)}</w:p>'


def heading(text, level):
    return para(text, style=f'Heading{level}')


def hr():
    return ('<w:p><w:pPr><w:pBdr>'
            '<w:bottom w:val="single" w:sz="6" w:space="1" w:color="999999"/>'
            '</w:pBdr></w:pPr></w:p>')


def quote(text):
    return para(text, style='Quote')


# ---------- Tables ----------
def table_xml(rows):
    """rows: lista de listas de celdas (texto md). Primera fila = encabezado."""
    ncols = max(len(r) for r in rows)
    grid = ''.join('<w:gridCol w:w="%d"/>' % (9000 // ncols) for _ in range(ncols))
    tbl_pr = ('<w:tblPr>'
              '<w:tblStyle w:val="TableGrid"/>'
              '<w:tblW w:w="5000" w:type="pct"/>'
              '<w:tblBorders>'
              '<w:top w:val="single" w:sz="4" w:space="0" w:color="BFBFBF"/>'
              '<w:left w:val="single" w:sz="4" w:space="0" w:color="BFBFBF"/>'
              '<w:bottom w:val="single" w:sz="4" w:space="0" w:color="BFBFBF"/>'
              '<w:right w:val="single" w:sz="4" w:space="0" w:color="BFBFBF"/>'
              '<w:insideH w:val="single" w:sz="4" w:space="0" w:color="BFBFBF"/>'
              '<w:insideV w:val="single" w:sz="4" w:space="0" w:color="BFBFBF"/>'
              '</w:tblBorders>'
              '<w:tblLook w:val="04A0" w:firstRow="1" w:lastRow="0" '
              'w:firstColumn="0" w:lastColumn="0" w:noHBand="0" w:noVBand="1"/>'
              '</w:tblPr>')
    out = [f'<w:tbl>{tbl_pr}<w:tblGrid>{grid}</w:tblGrid>']
    for i, row in enumerate(rows):
        is_header = (i == 0)
        cells = []
        cols = row + [''] * (ncols - len(row))
        for cell in cols:
            shade = ('<w:shd w:val="clear" w:color="auto" w:fill="1F3864"/>'
                     if is_header else '')
            tc_pr = f'<w:tcPr><w:tcW w:w="0" w:type="auto"/>{shade}</w:tcPr>'
            extra = {'bold': True, 'color': 'FFFFFF'} if is_header else {}
            cell_p = (f'<w:p><w:pPr><w:spacing w:after="20"/></w:pPr>'
                      f'{runs_to_xml(cell, extra)}</w:p>')
            cells.append(f'<w:tc>{tc_pr}{cell_p}</w:tc>')
        out.append(f'<w:tr>{"".join(cells)}</w:tr>')
    out.append('</w:tbl>')
    # parrafo vacio despues de la tabla (requisito OOXML)
    out.append(para(''))
    return ''.join(out)


def split_table_row(line):
    line = line.strip()
    if line.startswith('|'):
        line = line[1:]
    if line.endswith('|'):
        line = line[:-1]
    return [c.strip() for c in line.split('|')]


def is_separator_row(line):
    return bool(re.match(r'^\|?[\s:\-|]+\|?$', line)) and '-' in line


# ---------- Main markdown -> body ----------
def convert(md_text):
    lines = md_text.split('\n')
    body = []
    i = 0
    n = len(lines)
    while i < n:
        line = lines[i]
        stripped = line.strip()

        # Tabla
        if stripped.startswith('|') and i + 1 < n and is_separator_row(lines[i + 1].strip()):
            rows = [split_table_row(stripped)]
            i += 2  # saltar separador
            while i < n and lines[i].strip().startswith('|'):
                rows.append(split_table_row(lines[i].strip()))
                i += 1
            body.append(table_xml(rows))
            continue

        # Regla horizontal
        if stripped == '---':
            body.append(hr())
            i += 1
            continue

        # Encabezados
        m = re.match(r'^(#{1,6})\s+(.*)$', stripped)
        if m:
            level = min(len(m.group(1)), 4)
            body.append(heading(m.group(2), level))
            i += 1
            continue

        # Cita
        if stripped.startswith('>'):
            content = stripped.lstrip('>').strip()
            body.append(quote(content))
            i += 1
            continue

        # Lista numerada (con indentacion)
        m = re.match(r'^(\s*)(\d+)\.\s+(.*)$', line)
        if m:
            level = min(len(m.group(1)) // 2, 2)
            body.append(list_para(m.group(3), ordered=True, level=level))
            i += 1
            continue

        # Lista con vinetas (con indentacion)
        m = re.match(r'^(\s*)[-*]\s+(.*)$', line)
        if m:
            level = min(len(m.group(1)) // 2, 2)
            body.append(list_para(m.group(2), ordered=False, level=level))
            i += 1
            continue

        # Linea en blanco
        if stripped == '':
            i += 1
            continue

        # Parrafo normal
        body.append(para(stripped))
        i += 1

    return '\n'.join(body)


# ---------- OOXML boilerplate ----------
CONTENT_TYPES = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
<Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>
</Types>'''

RELS = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>'''

DOC_RELS = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/>
</Relationships>'''

NUMBERING = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:abstractNum w:abstractNumId="0">
<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val="\u2022"/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr></w:lvl>
<w:lvl w:ilvl="1"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val="\u25e6"/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="1440" w:hanging="360"/></w:pPr></w:lvl>
<w:lvl w:ilvl="2"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val="\u25aa"/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="2160" w:hanging="360"/></w:pPr></w:lvl>
</w:abstractNum>
<w:abstractNum w:abstractNumId="1">
<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr></w:lvl>
<w:lvl w:ilvl="1"><w:start w:val="1"/><w:numFmt w:val="lowerLetter"/><w:lvlText w:val="%2."/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="1440" w:hanging="360"/></w:pPr></w:lvl>
<w:lvl w:ilvl="2"><w:start w:val="1"/><w:numFmt w:val="lowerRoman"/><w:lvlText w:val="%3."/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="2160" w:hanging="360"/></w:pPr></w:lvl>
</w:abstractNum>
<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>
<w:num w:numId="2"><w:abstractNumId w:val="1"/></w:num>
</w:numbering>'''

STYLES = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:sz w:val="22"/></w:rPr></w:rPrDefault></w:docDefaults>
<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:pPr><w:spacing w:after="160" w:line="276" w:lineRule="auto"/></w:pPr></w:style>
<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/><w:pPr><w:keepNext/><w:spacing w:before="360" w:after="120"/><w:outlineLvl w:val="0"/></w:pPr><w:rPr><w:b/><w:color w:val="1F3864"/><w:sz w:val="36"/></w:rPr></w:style>
<w:style w:type="paragraph" w:styleId="Heading2"><w:name w:val="heading 2"/><w:basedOn w:val="Normal"/><w:pPr><w:keepNext/><w:spacing w:before="280" w:after="100"/><w:outlineLvl w:val="1"/></w:pPr><w:rPr><w:b/><w:color w:val="2E5496"/><w:sz w:val="30"/></w:rPr></w:style>
<w:style w:type="paragraph" w:styleId="Heading3"><w:name w:val="heading 3"/><w:basedOn w:val="Normal"/><w:pPr><w:keepNext/><w:spacing w:before="220" w:after="80"/><w:outlineLvl w:val="2"/></w:pPr><w:rPr><w:b/><w:color w:val="2E5496"/><w:sz w:val="26"/></w:rPr></w:style>
<w:style w:type="paragraph" w:styleId="Heading4"><w:name w:val="heading 4"/><w:basedOn w:val="Normal"/><w:pPr><w:keepNext/><w:spacing w:before="180" w:after="60"/><w:outlineLvl w:val="3"/></w:pPr><w:rPr><w:b/><w:i/><w:color w:val="404040"/><w:sz w:val="24"/></w:rPr></w:style>
<w:style w:type="paragraph" w:styleId="Quote"><w:name w:val="Quote"/><w:basedOn w:val="Normal"/><w:pPr><w:ind w:left="480"/><w:pBdr><w:left w:val="single" w:sz="18" w:space="8" w:color="2E5496"/></w:pBdr><w:spacing w:before="80" w:after="80"/></w:pPr><w:rPr><w:i/><w:color w:val="404040"/></w:rPr></w:style>
<w:style w:type="paragraph" w:styleId="ListParagraph"><w:name w:val="List Paragraph"/><w:basedOn w:val="Normal"/><w:pPr><w:spacing w:after="60"/><w:ind w:left="720"/></w:pPr></w:style>
<w:style w:type="paragraph" w:styleId="Title"><w:name w:val="Title"/><w:basedOn w:val="Normal"/><w:pPr><w:jc w:val="center"/><w:spacing w:after="80"/></w:pPr><w:rPr><w:b/><w:color w:val="1F3864"/><w:sz w:val="56"/></w:rPr></w:style>
<w:style w:type="paragraph" w:styleId="Subtitle"><w:name w:val="Subtitle"/><w:basedOn w:val="Normal"/><w:pPr><w:jc w:val="center"/><w:spacing w:after="240"/></w:pPr><w:rPr><w:i/><w:color w:val="2E5496"/><w:sz w:val="28"/></w:rPr></w:style>
<w:style w:type="table" w:styleId="TableGrid"><w:name w:val="Table Grid"/><w:basedOn w:val="TableNormal"/><w:tblPr><w:tblBorders><w:top w:val="single" w:sz="4" w:space="0" w:color="BFBFBF"/><w:left w:val="single" w:sz="4" w:space="0" w:color="BFBFBF"/><w:bottom w:val="single" w:sz="4" w:space="0" w:color="BFBFBF"/><w:right w:val="single" w:sz="4" w:space="0" w:color="BFBFBF"/><w:insideH w:val="single" w:sz="4" w:space="0" w:color="BFBFBF"/><w:insideV w:val="single" w:sz="4" w:space="0" w:color="BFBFBF"/></w:tblBorders></w:tblPr></w:style>
<w:style w:type="table" w:default="1" w:styleId="TableNormal"><w:name w:val="Normal Table"/></w:style>
</w:styles>'''


def build(md_path, out_path):
    with open(md_path, 'r', encoding='utf-8') as f:
        md = f.read()
    body = convert(md)
    sect_pr = ('<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>'
               '<w:pgMar w:top="1418" w:right="1418" w:bottom="1418" '
               'w:left="1418" w:header="708" w:footer="708" w:gutter="0"/>'
               '</w:sectPr>')
    document = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<w:document xmlns:w="{NSW}">'
        f'<w:body>{body}{sect_pr}</w:body></w:document>'
    )
    with zipfile.ZipFile(out_path, 'w', zipfile.ZIP_DEFLATED) as z:
        z.writestr('[Content_Types].xml', CONTENT_TYPES)
        z.writestr('_rels/.rels', RELS)
        z.writestr('word/document.xml', document)
        z.writestr('word/styles.xml', STYLES)
        z.writestr('word/numbering.xml', NUMBERING)
        z.writestr('word/_rels/document.xml.rels', DOC_RELS)
    print(f'OK -> {out_path}')


if __name__ == '__main__':
    if len(sys.argv) != 3:
        print('Uso: python3 build_docx.py <entrada.md> <salida.docx>')
        sys.exit(1)
    build(sys.argv[1], sys.argv[2])
