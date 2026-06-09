# -*- coding: utf-8 -*-
"""
Generador del Dashboard Ejecutivo LATAM "Mision Cielo Verde".
Construye un .xlsx (OOXML) desde cero usando solo la libreria estandar,
porque el sandbox no tiene pandas/xlsxwriter ni acceso a PyPI.

Incluye:
  - Hoja "Data Cruda": 60 filas correlacionadas (ruta Santiago-Miami).
  - Hoja "Calculos": agregaciones por semana con formulas SUMIF/AVERAGEIF.
  - Hoja "Dashboard": tema azul marino, KPIs gigantes, semaforo (icon set)
    forzado a ROJO en semana 3, validacion de datos (slicer visual),
    grafico de lineas (margen) y grafico de anillo/gauge (APU).
"""
import os, shutil, zipfile, datetime

OUT = "/projects/sandbox/LATAM_Cielo_Verde_Dashboard.xlsx"
BUILD = "/projects/sandbox/_build_xlsx"

# ----------------------------------------------------------------------
# 1) DATOS CORRELACIONADOS (sin azar: relacion matematica real)
# ----------------------------------------------------------------------
ECO_BASE = {1: 18, 2: 19, 3: 4, 4: 8, 5: 9, 6: 11, 7: 12, 8: 13, 9: 14}
APU_BASE = {1: 55, 2: 52, 3: 50, 4: 46, 5: 43, 6: 40, 7: 37, 8: 34, 9: 31}
WIGGLE = [-1, 0, 1, 0, 1, -1, 0]
CAP_ECO = 20                  # capacidad cabina ejecutiva
ECO_SURCHARGE = 1500          # sobreprecio USD por asiento Eco-Corporate
BASE_ECON_MARGIN = 9000       # contribucion economica base (USD)
SAF_EXTRA_COST = 1.4          # USD de sobrecosto por litro SAF

start = datetime.date(2026, 1, 1)
EPOCH = datetime.date(1899, 12, 30)   # serial Excel (sistema 1900)

rows = []   # cada fila: dict con campos
for i in range(60):
    d = start + datetime.timedelta(days=i)
    serial = (d - EPOCH).days
    week = i // 7 + 1
    eco = ECO_BASE[week] + WIGGLE[i % 7]
    eco = max(2, min(CAP_ECO, eco))
    saf = round(4200 + eco * 40 + (i % 5) * 25)
    apu = round(APU_BASE[week] + (i % 4 - 1) * 1.5, 1)
    retraso = 5 + (i % 6) * 3 + (8 if week == 3 else 0)
    margen = round(BASE_ECON_MARGIN + eco * ECO_SURCHARGE - saf * SAF_EXTRA_COST)
    vuelo = "LA{}".format(532 + (i % 2))
    rows.append(dict(serial=serial, week=week, vuelo=vuelo, saf=saf,
                     eco=eco, margen=margen, apu=apu, retraso=retraso))

# Agregados semanales (valores cacheados para las formulas)
weeks = list(range(1, 10))
agg = {}
for w in weeks:
    wr = [r for r in rows if r["week"] == w]
    n = len(wr)
    agg[w] = dict(
        eco_sum=sum(r["eco"] for r in wr),
        eco_avg=round(sum(r["eco"] for r in wr) / n, 1),
        margen_sum=sum(r["margen"] for r in wr),
        saf_sum=sum(r["saf"] for r in wr),
        apu_avg=round(sum(r["apu"] for r in wr) / n, 1),
        retraso_avg=round(sum(r["retraso"] for r in wr) / n, 1),
    )

margen_total = sum(r["margen"] for r in rows)
eco_total = sum(r["eco"] for r in rows)
adopcion = eco_total / (60 * CAP_ECO)
apu_global = round(sum(r["apu"] for r in rows) / 60, 1)
holgura = round(60 - apu_global, 1)
week3_eco = agg[3]["eco_avg"]

# ----------------------------------------------------------------------
# Utilidades XML
# ----------------------------------------------------------------------
def col_letter(c):
    s = ""
    while c > 0:
        c, r = divmod(c - 1, 26)
        s = chr(65 + r) + s
    return s

def ref(r, c):
    return "{}{}".format(col_letter(c), r)

def esc(t):
    return (str(t).replace("&", "&amp;").replace("<", "&lt;")
            .replace(">", "&gt;").replace('"', "&quot;"))

class Sheet:
    def __init__(self):
        self.cells = {}        # (r,c) -> xml fragment
        self.merges = []
        self.colw = []         # (min,max,width)
        self.rowh = {}         # row -> height
        self.show_grid = True
        self.tab_selected = False

    def num(self, r, c, v, s=0):
        self.cells[(r, c)] = '<c r="{}" s="{}"><v>{}</v></c>'.format(ref(r, c), s, v)

    def txt(self, r, c, v, s=0):
        self.cells[(r, c)] = '<c r="{}" s="{}" t="inlineStr"><is><t xml:space="preserve">{}</t></is></c>'.format(ref(r, c), s, esc(v))

    def formula(self, r, c, f, cached, s=0):
        self.cells[(r, c)] = '<c r="{}" s="{}"><f>{}</f><v>{}</v></c>'.format(ref(r, c), s, esc(f), cached)

    def blank(self, r, c, s=0):
        self.cells[(r, c)] = '<c r="{}" s="{}"/>'.format(ref(r, c), s)

    def merge(self, r1, c1, r2, c2):
        self.merges.append("{}:{}".format(ref(r1, c1), ref(r2, c2)))

    def to_xml(self, extra=""):
        # filas ordenadas
        by_row = {}
        for (r, c), frag in self.cells.items():
            by_row.setdefault(r, []).append((c, frag))
        rows_xml = []
        for r in sorted(by_row):
            cells = "".join(f for _, f in sorted(by_row[r], key=lambda x: x[0]))
            h = ' ht="{}" customHeight="1"'.format(self.rowh[r]) if r in self.rowh else ""
            rows_xml.append('<row r="{}"{}>{}</row>'.format(r, h, cells))
        cols_xml = ""
        if self.colw:
            cols_xml = "<cols>" + "".join(
                '<col min="{}" max="{}" width="{}" customWidth="1"/>'.format(a, b, w)
                for a, b, w in self.colw) + "</cols>"
        merges_xml = ""
        if self.merges:
            merges_xml = '<mergeCells count="{}">'.format(len(self.merges)) + \
                "".join('<mergeCell ref="{}"/>'.format(m) for m in self.merges) + "</mergeCells>"
        sv = '<sheetView{}{} workbookViewId="0"/>'.format(
            "" if self.show_grid else ' showGridLines="0"',
            ' tabSelected="1"' if self.tab_selected else "")
        return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
                'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
                '<sheetViews>' + sv + '</sheetViews>'
                '<sheetFormatPr defaultRowHeight="15"/>'
                + cols_xml +
                '<sheetData>' + "".join(rows_xml) + '</sheetData>'
                + merges_xml + extra + '</worksheet>')

# ======================================================================
# HOJA 1 : DATA CRUDA
# ======================================================================
s1 = Sheet()
s1.colw = [(1, 1, 13), (2, 2, 9), (3, 3, 11), (4, 4, 16),
           (5, 5, 22), (6, 6, 16), (7, 7, 18), (8, 8, 16)]
headers = ["Fecha", "Semana", "N Vuelo", "Litros SAF",
           "Asientos Eco-Corporate", "Margen Ruta USD",
           "Minutos APU en Tierra", "Minutos de retraso"]
for c, h in enumerate(headers, 1):
    s1.txt(1, c, h, 4)
s1.rowh[1] = 30
for idx, rdat in enumerate(rows):
    rr = idx + 2
    s1.num(rr, 1, rdat["serial"], 1)      # fecha
    s1.num(rr, 2, rdat["week"], 2)        # semana
    s1.txt(rr, 3, rdat["vuelo"], 2)       # vuelo
    s1.num(rr, 4, rdat["saf"], 5)         # litros SAF
    s1.num(rr, 5, rdat["eco"], 2)         # asientos eco
    s1.num(rr, 6, rdat["margen"], 3)      # margen USD
    s1.num(rr, 7, rdat["apu"], 20)        # APU min
    s1.num(rr, 8, rdat["retraso"], 2)     # retraso

# ======================================================================
# HOJA 2 : CALCULOS  (formulas reales sobre Data Cruda)
# ======================================================================
s2 = Sheet()
s2.colw = [(1, 1, 10), (2, 2, 20), (3, 3, 20), (4, 4, 20),
           (5, 5, 18), (6, 6, 18), (7, 7, 16)]
chead = ["Semana", "Asientos Eco (Suma)", "Asientos Eco (Prom)",
         "Margen USD (Suma)", "Litros SAF (Suma)", "APU min (Prom)", "Retraso (Prom)"]
for c, h in enumerate(chead, 1):
    s2.txt(1, c, h, 4)
s2.rowh[1] = 30
D = "'Data Cruda'"
for j, w in enumerate(weeks):
    rr = j + 2
    a = agg[w]
    s2.num(rr, 1, w, 2)
    s2.formula(rr, 2, "SUMIF({0}!$B$2:$B$61,A{1},{0}!$E$2:$E$61)".format(D, rr), a["eco_sum"], 2)
    s2.formula(rr, 3, "AVERAGEIF({0}!$B$2:$B$61,A{1},{0}!$E$2:$E$61)".format(D, rr), a["eco_avg"], 20)
    s2.formula(rr, 4, "SUMIF({0}!$B$2:$B$61,A{1},{0}!$F$2:$F$61)".format(D, rr), a["margen_sum"], 3)
    s2.formula(rr, 5, "SUMIF({0}!$B$2:$B$61,A{1},{0}!$D$2:$D$61)".format(D, rr), a["saf_sum"], 2)
    s2.formula(rr, 6, "AVERAGEIF({0}!$B$2:$B$61,A{1},{0}!$G$2:$G$61)".format(D, rr), a["apu_avg"], 20)
    s2.formula(rr, 7, "AVERAGEIF({0}!$B$2:$B$61,A{1},{0}!$H$2:$H$61)".format(D, rr), a["retraso_avg"], 20)

# Bloque para el Gauge (anillo) -> filas 14 y 15
s2.txt(13, 1, "Datos Gauge APU", 4)
s2.txt(14, 1, "APU Promedio", 0)
s2.formula(14, 2, "AVERAGE({0}!$G$2:$G$61)".format(D), apu_global, 20)
s2.txt(15, 1, "Holgura (Meta 60)", 0)
s2.formula(15, 2, "60-B14", holgura, 20)

# ======================================================================
# HOJA 3 : DASHBOARD
# ======================================================================
s3 = Sheet()
s3.show_grid = False
s3.tab_selected = True
NCOLS = 14
s3.colw = [(1, NCOLS, 11.5)]
# Fondo navy: rellenar toda la zona con estilo 6
for r in range(1, 47):
    for c in range(1, NCOLS + 1):
        s3.blank(r, c, 6)

# --- Titulo ---
s3.rowh[1] = 34
s3.rowh[2] = 26
s3.txt(1, 1, "LATAM AIRLINES  |  MISION CIELO VERDE", 23)
s3.merge(1, 1, 1, NCOLS)
s3.txt(2, 1, "Ruta Piloto: Santiago (SCL)  ->  Miami (MIA)   |   Programa Eco-Corporate Premium", 17)
s3.merge(2, 1, 2, NCOLS)

# --- Filtro (slicer visual con validacion de datos) ---
s3.rowh[3] = 22
s3.txt(3, 1, "SELECCIONE SEMANA A ANALIZAR:", 8)
s3.merge(3, 1, 3, 4)
s3.num(3, 5, 3, 16)            # celda con validacion (lista 1-9)
s3.merge(3, 5, 3, 6)
s3.txt(3, 8, "Periodo: 60 dias  /  9 semanas  /  vuelo diario", 17)
s3.merge(3, 8, 3, NCOLS)

# --- Banda de acento verde brillante (separador) ---
s3.rowh[4] = 7
for c in range(1, NCOLS + 1):
    s3.blank(4, c, 36)
s3.merge(4, 1, 4, NCOLS)

# --- KPIs (4 paneles con colores vivos) ---
def panel(r1, c1, r2, c2, label, blank_s, label_s, value_frag_setter):
    # pinta panel con su color, pone label arriba y valor abajo
    for r in range(r1, r2 + 1):
        for c in range(c1, c2 + 1):
            s3.blank(r, c, blank_s)
    s3.txt(r1, c1, label, label_s)
    s3.merge(r1, c1, r1, c2)
    value_frag_setter(r1 + 1, c1, c2)

s3.rowh[5] = 22
s3.rowh[6] = 46
# Panel 1: Margen total (VERDE)
panel(5, 1, 6, 4, "MARGEN TOTAL RUTA (USD)", 32, 24,
      lambda r, c1, c2: (s3.formula(r, c1, "SUM('Data Cruda'!F2:F61)", margen_total, 25), s3.merge(r, c1, r, c2)))
# Panel 2: Adopcion green % (CIAN)
panel(5, 5, 6, 8, "ADOPCION GREEN %", 33, 26,
      lambda r, c1, c2: (s3.formula(r, c1, "SUM('Data Cruda'!E2:E61)/1200", round(adopcion, 4), 27), s3.merge(r, c1, r, c2)))
# Panel 3: APU promedio (AMBAR - lead)
panel(5, 9, 6, 11, "APU PROM (min) LEAD", 34, 28,
      lambda r, c1, c2: (s3.formula(r, c1, "'Calculos'!B14", apu_global, 29), s3.merge(r, c1, r, c2)))
# Panel 4: Venta Eco S3 (ROJO - crisis -> semaforo rojo)
panel(5, 12, 6, 14, "VENTA ECO S3 (CRISIS)", 35, 30,
      lambda r, c1, c2: (s3.formula(r, c1, "'Calculos'!C4", week3_eco, 31), s3.merge(r, c1, r, c2)))
SEMAFORO_CRISIS = ref(6, 12)   # celda del semaforo individual forzado a rojo

# --- Tabla semanal con semaforo automatico (icon set) ---
s3.txt(8, 1, "SEMAFORO ADOPCION ECO-CORPORATE POR SEMANA  (rojo = critico)", 8)
s3.merge(8, 1, 8, NCOLS)
s3.txt(9, 1, "Semana", 13)
for j, w in enumerate(weeks):
    s3.num(9, 2 + j, w, 13)
s3.txt(10, 1, "Asientos Eco (prom)", 14)
for j, w in enumerate(weeks):
    s3.formula(10, 2 + j, "'Calculos'!C{}".format(2 + j), agg[w]["eco_avg"], 21)
ICON_RANGE = "{}:{}".format(ref(10, 2), ref(10, 10))   # B10:J10

# --- Encabezados de graficos ---
s3.txt(12, 1, "EVOLUCION DEL MARGEN DE RUTA (60 DIAS)", 8)
s3.merge(12, 1, 12, 8)
s3.txt(12, 9, "GAUGE: MINUTOS APU EN TIERRA", 8)
s3.merge(12, 9, 12, NCOLS)

# --- Banda accion agil ---
s3.txt(31, 1, "ACCION CORRECTIVA AGIL (48H) - Crisis Semana 3:", 8)
s3.merge(31, 1, 31, NCOLS)
s3.txt(32, 1, "Reasignar squad comercial: contactar Top 5 multinacionales y ofrecer upgrade Eco-Corporate de prueba + certificado Scope 3 piloto.", 17)
s3.merge(32, 1, 32, NCOLS)

# --- Formato condicional (semaforos) ---
cf = (
    # Semaforo automatico tabla semanal: rojo<10, amarillo<15, verde>=15
    '<conditionalFormatting sqref="{rng}">'
    '<cfRule type="iconSet" priority="1"><iconSet iconSet="3TrafficLights1">'
    '<cfvo type="percent" val="0"/><cfvo type="num" val="10"/><cfvo type="num" val="15"/>'
    '</iconSet></cfRule></conditionalFormatting>'
    # Semaforo individual KPI crisis: thresholds altos -> SIEMPRE rojo
    '<conditionalFormatting sqref="{crisis}">'
    '<cfRule type="iconSet" priority="2"><iconSet iconSet="3TrafficLights1">'
    '<cfvo type="percent" val="0"/><cfvo type="num" val="10"/><cfvo type="num" val="15"/>'
    '</iconSet></cfRule></conditionalFormatting>'
).format(rng=ICON_RANGE, crisis=SEMAFORO_CRISIS)

# --- Validacion de datos (slicer visual) ---
dv = ('<dataValidations count="1">'
      '<dataValidation type="list" allowBlank="1" showInputMessage="1" showErrorMessage="1" '
      'sqref="{}"><formula1>"1,2,3,4,5,6,7,8,9"</formula1></dataValidation>'
      '</dataValidations>').format(ref(3, 5))

# --- Drawing (graficos) ---
drawing_rel = '<drawing r:id="rId1"/>'
s3_extra = cf + dv + drawing_rel

# ======================================================================
# STYLES.XML
# ======================================================================
styles_xml = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<numFmts count="5">
<numFmt numFmtId="164" formatCode="dd/mm/yyyy"/>
<numFmt numFmtId="165" formatCode="&quot;$&quot;#,##0"/>
<numFmt numFmtId="166" formatCode="0.0%"/>
<numFmt numFmtId="167" formatCode="#,##0"/>
<numFmt numFmtId="168" formatCode="0.0"/>
</numFmts>
<fonts count="15">
<font><sz val="11"/><color rgb="FF000000"/><name val="Calibri"/></font>
<font><b/><sz val="11"/><color rgb="FFFFFFFF"/><name val="Calibri"/></font>
<font><sz val="11"/><color rgb="FFFFFFFF"/><name val="Calibri"/></font>
<font><b/><sz val="20"/><color rgb="FF00E676"/><name val="Calibri"/></font>
<font><b/><sz val="16"/><color rgb="FFFFFFFF"/><name val="Calibri"/></font>
<font><b/><sz val="26"/><color rgb="FFFFFFFF"/><name val="Calibri"/></font>
<font><b/><sz val="11"/><color rgb="FFFFFFFF"/><name val="Calibri"/></font>
<font><b/><sz val="12"/><color rgb="FF00E676"/><name val="Calibri"/></font>
<font><b/><sz val="13"/><color rgb="FF00E676"/><name val="Calibri"/></font>
<font><b/><sz val="12"/><color rgb="FF0B1F3A"/><name val="Calibri"/></font>
<font><b/><sz val="30"/><color rgb="FF00E676"/><name val="Calibri"/></font>
<font><sz val="10"/><color rgb="FFB0BEC5"/><name val="Calibri"/></font>
<font><b/><sz val="30"/><color rgb="FF18FFFF"/><name val="Calibri"/></font>
<font><b/><sz val="30"/><color rgb="FFFFC107"/><name val="Calibri"/></font>
<font><b/><sz val="30"/><color rgb="FFFF5252"/><name val="Calibri"/></font>
</fonts>
<fills count="14">
<fill><patternFill patternType="none"/></fill>
<fill><patternFill patternType="gray125"/></fill>
<fill><patternFill patternType="solid"><fgColor rgb="FF0B1F3A"/></patternFill></fill>
<fill><patternFill patternType="solid"><fgColor rgb="FF12305A"/></patternFill></fill>
<fill><patternFill patternType="solid"><fgColor rgb="FF1B5E20"/></patternFill></fill>
<fill><patternFill patternType="solid"><fgColor rgb="FFFFC107"/></patternFill></fill>
<fill><patternFill patternType="solid"><fgColor rgb="FF173E6B"/></patternFill></fill>
<fill><patternFill patternType="solid"><fgColor rgb="FF0E2747"/></patternFill></fill>
<fill><gradientFill degree="35"><stop position="0"><color rgb="FF00C853"/></stop><stop position="1"><color rgb="FF0B1F3A"/></stop></gradientFill></fill>
<fill><patternFill patternType="solid"><fgColor rgb="FF14502B"/></patternFill></fill>
<fill><patternFill patternType="solid"><fgColor rgb="FF0E4C5A"/></patternFill></fill>
<fill><patternFill patternType="solid"><fgColor rgb="FF5A3E0E"/></patternFill></fill>
<fill><patternFill patternType="solid"><fgColor rgb="FF5A1414"/></patternFill></fill>
<fill><patternFill patternType="solid"><fgColor rgb="FF00C853"/></patternFill></fill>
</fills>
<borders count="3">
<border><left/><right/><top/><bottom/><diagonal/></border>
<border><left style="thin"><color rgb="FF2E5C8A"/></left><right style="thin"><color rgb="FF2E5C8A"/></right><top style="thin"><color rgb="FF2E5C8A"/></top><bottom style="thin"><color rgb="FF2E5C8A"/></bottom></border>
<border><left style="thin"><color rgb="FFBFBFBF"/></left><right style="thin"><color rgb="FFBFBFBF"/></right><top style="thin"><color rgb="FFBFBFBF"/></top><bottom style="thin"><color rgb="FFBFBFBF"/></bottom></border>
</borders>
<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
<cellXfs count="37">
<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
<xf numFmtId="164" fontId="0" fillId="0" borderId="2" xfId="0" applyNumberFormat="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center"/></xf>
<xf numFmtId="0" fontId="0" fillId="0" borderId="2" xfId="0" applyBorder="1" applyAlignment="1"><alignment horizontal="center"/></xf>
<xf numFmtId="165" fontId="0" fillId="0" borderId="2" xfId="0" applyNumberFormat="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center"/></xf>
<xf numFmtId="0" fontId="6" fillId="6" borderId="2" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf>
<xf numFmtId="167" fontId="0" fillId="0" borderId="2" xfId="0" applyNumberFormat="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center"/></xf>
<xf numFmtId="0" fontId="2" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1"/>
<xf numFmtId="0" fontId="5" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1" applyAlignment="1"><alignment horizontal="left" vertical="center"/></xf>
<xf numFmtId="0" fontId="8" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1" applyAlignment="1"><alignment horizontal="left" vertical="center"/></xf>
<xf numFmtId="0" fontId="1" fillId="3" borderId="0" xfId="0" applyFont="1" applyFill="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>
<xf numFmtId="165" fontId="10" fillId="3" borderId="0" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>
<xf numFmtId="166" fontId="10" fillId="3" borderId="0" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>
<xf numFmtId="0" fontId="4" fillId="3" borderId="0" xfId="0" applyFont="1" applyFill="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>
<xf numFmtId="0" fontId="6" fillId="6" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>
<xf numFmtId="0" fontId="2" fillId="2" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>
<xf numFmtId="165" fontId="2" fillId="2" borderId="1" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>
<xf numFmtId="0" fontId="9" fillId="5" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>
<xf numFmtId="0" fontId="11" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1" applyAlignment="1"><alignment horizontal="left" vertical="center"/></xf>
<xf numFmtId="0" fontId="7" fillId="3" borderId="0" xfId="0" applyFont="1" applyFill="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>
<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
<xf numFmtId="168" fontId="0" fillId="0" borderId="2" xfId="0" applyNumberFormat="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center"/></xf>
<xf numFmtId="168" fontId="2" fillId="2" borderId="1" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>
<xf numFmtId="0" fontId="2" fillId="3" borderId="0" xfId="0" applyFont="1" applyFill="1"/>
<xf numFmtId="0" fontId="5" fillId="8" borderId="0" xfId="0" applyFont="1" applyFill="1" applyAlignment="1"><alignment horizontal="left" vertical="center"/></xf>
<xf numFmtId="0" fontId="1" fillId="9" borderId="0" xfId="0" applyFont="1" applyFill="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>
<xf numFmtId="165" fontId="10" fillId="9" borderId="0" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>
<xf numFmtId="0" fontId="1" fillId="10" borderId="0" xfId="0" applyFont="1" applyFill="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>
<xf numFmtId="166" fontId="12" fillId="10" borderId="0" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>
<xf numFmtId="0" fontId="1" fillId="11" borderId="0" xfId="0" applyFont="1" applyFill="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>
<xf numFmtId="168" fontId="13" fillId="11" borderId="0" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>
<xf numFmtId="0" fontId="1" fillId="12" borderId="0" xfId="0" applyFont="1" applyFill="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>
<xf numFmtId="168" fontId="14" fillId="12" borderId="0" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>
<xf numFmtId="0" fontId="2" fillId="9" borderId="0" xfId="0" applyFont="1" applyFill="1"/>
<xf numFmtId="0" fontId="2" fillId="10" borderId="0" xfId="0" applyFont="1" applyFill="1"/>
<xf numFmtId="0" fontId="2" fillId="11" borderId="0" xfId="0" applyFont="1" applyFill="1"/>
<xf numFmtId="0" fontId="2" fillId="12" borderId="0" xfId="0" applyFont="1" applyFill="1"/>
<xf numFmtId="0" fontId="2" fillId="13" borderId="0" xfId="0" applyFont="1" applyFill="1"/>
</cellXfs>
<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
</styleSheet>'''

# ======================================================================
# CHARTS
# ======================================================================
NS_C = 'xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'

def num_cache(values, fmt="General"):
    pts = "".join('<c:pt idx="{}"><c:v>{}</c:v></c:pt>'.format(i, v) for i, v in enumerate(values))
    return ('<c:numCache><c:formatCode>{}</c:formatCode><c:ptCount val="{}"/>{}</c:numCache>'
            .format(fmt, len(values), pts))

def str_cache(values):
    pts = "".join('<c:pt idx="{}"><c:v>{}</c:v></c:pt>'.format(i, esc(v)) for i, v in enumerate(values))
    return '<c:strCache><c:ptCount val="{}"/>{}</c:strCache>'.format(len(values), pts)

cache_cat1 = num_cache([r["serial"] for r in rows], "dd/mm")
cache_val1 = num_cache([r["margen"] for r in rows], '"$"#,##0')
cache_cat2 = str_cache(["APU Promedio", "Holgura (Meta 60)"])
cache_val2 = num_cache([apu_global, holgura], "0.0")

chart1 = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace {ns}>
<c:chart>
<c:title><c:tx><c:rich><a:bodyPr/><a:p><a:pPr><a:defRPr sz="1100" b="1"><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill></a:defRPr></a:pPr><a:r><a:rPr lang="es" sz="1100" b="1"><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill></a:rPr><a:t>Margen Ruta USD (Lag Measure)</a:t></a:r></a:p></c:rich></c:tx><c:overlay val="0"/></c:title>
<c:autoTitleDeleted val="0"/>
<c:plotArea><c:layout/>
<c:lineChart><c:grouping val="standard"/><c:varyColors val="0"/>
<c:ser><c:idx val="0"/><c:order val="0"/>
<c:tx><c:v>Margen USD</c:v></c:tx>
<c:spPr><a:ln w="28575"><a:solidFill><a:srgbClr val="00E676"/></a:solidFill></a:ln></c:spPr>
<c:marker><c:symbol val="none"/></c:marker>
<c:cat><c:numRef><c:f>'Data Cruda'!$A$2:$A$61</c:f>{cc1}</c:numRef></c:cat>
<c:val><c:numRef><c:f>'Data Cruda'!$F$2:$F$61</c:f>{cv1}</c:numRef></c:val>
<c:smooth val="0"/>
</c:ser>
<c:marker val="1"/>
<c:axId val="111111111"/><c:axId val="222222222"/>
</c:lineChart>
<c:catAx><c:axId val="111111111"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="b"/>
<c:numFmt formatCode="dd/mm" sourceLinked="0"/>
<c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="800"><a:solidFill><a:srgbClr val="B0BEC5"/></a:solidFill></a:defRPr></a:pPr><a:endParaRPr lang="es"/></a:p></c:txPr>
<c:crossAx val="222222222"/></c:catAx>
<c:valAx><c:axId val="222222222"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="l"/>
<c:majorGridlines><c:spPr><a:ln><a:solidFill><a:srgbClr val="1F3B63"/></a:solidFill></a:ln></c:spPr></c:majorGridlines>
<c:numFmt formatCode="&quot;$&quot;#,##0" sourceLinked="0"/>
<c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="800"><a:solidFill><a:srgbClr val="B0BEC5"/></a:solidFill></a:defRPr></a:pPr><a:endParaRPr lang="es"/></a:p></c:txPr>
<c:crossAx val="111111111"/></c:valAx>
<c:spPr><a:noFill/><a:ln><a:noFill/></a:ln></c:spPr>
</c:plotArea>
<c:plotVisOnly val="1"/><c:dispBlanksAs val="gap"/>
</c:chart>
<c:spPr><a:solidFill><a:srgbClr val="12305A"/></a:solidFill><a:ln><a:noFill/></a:ln></c:spPr>
</c:chartSpace>'''.format(ns=NS_C, cc1=cache_cat1, cv1=cache_val1)

chart2 = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace {ns}>
<c:chart>
<c:title><c:tx><c:rich><a:bodyPr/><a:p><a:pPr><a:defRPr sz="1100" b="1"><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill></a:defRPr></a:pPr><a:r><a:rPr lang="es" sz="1100" b="1"><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill></a:rPr><a:t>APU min en Tierra (Lead Measure)</a:t></a:r></a:p></c:rich></c:tx><c:overlay val="0"/></c:title>
<c:autoTitleDeleted val="0"/>
<c:plotArea><c:layout/>
<c:doughnutChart><c:varyColors val="1"/>
<c:ser><c:idx val="0"/><c:order val="0"/>
<c:dPt><c:idx val="0"/><c:bubble3D val="0"/><c:spPr><a:solidFill><a:srgbClr val="00E676"/></a:solidFill><a:ln><a:solidFill><a:srgbClr val="12305A"/></a:solidFill></a:ln></c:spPr></c:dPt>
<c:dPt><c:idx val="1"/><c:bubble3D val="0"/><c:spPr><a:solidFill><a:srgbClr val="55606E"/></a:solidFill><a:ln><a:solidFill><a:srgbClr val="12305A"/></a:solidFill></a:ln></c:spPr></c:dPt>
<c:cat><c:strRef><c:f>'Calculos'!$A$14:$A$15</c:f>{cc2}</c:strRef></c:cat>
<c:val><c:numRef><c:f>'Calculos'!$B$14:$B$15</c:f>{cv2}</c:numRef></c:val>
</c:ser>
<c:firstSliceAng val="270"/><c:holeSize val="62"/>
</c:doughnutChart>
<c:spPr><a:noFill/><a:ln><a:noFill/></a:ln></c:spPr>
</c:plotArea>
<c:legend><c:legendPos val="b"/><c:overlay val="0"/>
<c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="900"><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill></a:defRPr></a:pPr><a:endParaRPr lang="es"/></a:p></c:txPr>
</c:legend>
<c:plotVisOnly val="1"/>
</c:chart>
<c:spPr><a:solidFill><a:srgbClr val="12305A"/></a:solidFill><a:ln><a:noFill/></a:ln></c:spPr>
</c:chartSpace>'''.format(ns=NS_C, cc2=cache_cat2, cv2=cache_val2)

# ======================================================================
# DRAWING (anclaje de graficos en el dashboard)
# ======================================================================
NS_XDR = 'xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'

def anchor(c_from, r_from, c_to, r_to, rid, gid, name):
    return ('<xdr:twoCellAnchor editAs="oneCell">'
            '<xdr:from><xdr:col>{cf}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>{rf}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>'
            '<xdr:to><xdr:col>{ct}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>{rt}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>'
            '<xdr:graphicFrame macro="">'
            '<xdr:nvGraphicFramePr><xdr:cNvPr id="{gid}" name="{name}"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr>'
            '<xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm>'
            '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">'
            '<c:chart r:id="{rid}"/></a:graphicData></a:graphic>'
            '</xdr:graphicFrame><xdr:clientData/></xdr:twoCellAnchor>').format(
                cf=c_from, rf=r_from, ct=c_to, rt=r_to, rid=rid, gid=gid, name=name)

# Linea: cols A..H (0..8), filas 13..30 ; Anillo: cols I..N (8..14), filas 13..30
drawing_xml = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
               '<xdr:wsDr {ns}>'.format(ns=NS_XDR)
               + anchor(0, 12, 8, 30, "rId1", 2, "GraficoMargen")
               + anchor(8, 12, 14, 30, "rId2", 3, "GraficoGauge")
               + '</xdr:wsDr>')

drawing_rels = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart1.xml"/>'
                '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart2.xml"/>'
                '</Relationships>')

sheet3_rels = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
               '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
               '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/>'
               '</Relationships>')

# ======================================================================
# WORKBOOK + RELS + CONTENT TYPES + THEME
# ======================================================================
workbook_xml = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
                'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
                '<bookViews><workbookView activeTab="2"/></bookViews>'
                '<sheets>'
                '<sheet name="Data Cruda" sheetId="1" r:id="rId1"/>'
                '<sheet name="Calculos" sheetId="2" r:id="rId2"/>'
                '<sheet name="Dashboard" sheetId="3" r:id="rId3"/>'
                '</sheets>'
                '<calcPr fullCalcOnLoad="1"/>'
                '</workbook>')

workbook_rels = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                 '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                 '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>'
                 '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/>'
                 '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet3.xml"/>'
                 '<Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
                 '<Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>'
                 '</Relationships>')

root_rels = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
             '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
             '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>'
             '</Relationships>')

content_types = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                 '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
                 '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
                 '<Default Extension="xml" ContentType="application/xml"/>'
                 '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
                 '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
                 '<Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
                 '<Override PartName="/xl/worksheets/sheet3.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
                 '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'
                 '<Override PartName="/xl/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>'
                 '<Override PartName="/xl/drawings/drawing1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>'
                 '<Override PartName="/xl/charts/chart1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>'
                 '<Override PartName="/xl/charts/chart2.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>'
                 '</Types>')

theme_xml = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office Theme"><a:themeElements><a:clrScheme name="Office"><a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1><a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1><a:dk2><a:srgbClr val="1F497D"/></a:dk2><a:lt2><a:srgbClr val="EEECE1"/></a:lt2><a:accent1><a:srgbClr val="4F81BD"/></a:accent1><a:accent2><a:srgbClr val="C0504D"/></a:accent2><a:accent3><a:srgbClr val="9BBB59"/></a:accent3><a:accent4><a:srgbClr val="8064A2"/></a:accent4><a:accent5><a:srgbClr val="4BACC6"/></a:accent5><a:accent6><a:srgbClr val="F79646"/></a:accent6><a:hlink><a:srgbClr val="0000FF"/></a:hlink><a:folHlink><a:srgbClr val="800080"/></a:folHlink></a:clrScheme><a:fontScheme name="Office"><a:majorFont><a:latin typeface="Calibri Light"/><a:ea typeface=""/><a:cs typeface=""/></a:majorFont><a:minorFont><a:latin typeface="Calibri"/><a:ea typeface=""/><a:cs typeface=""/></a:minorFont></a:fontScheme><a:fmtScheme name="Office"><a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:tint val="50000"/><a:satMod val="300000"/></a:schemeClr></a:gs><a:gs pos="35000"><a:schemeClr val="phClr"><a:tint val="37000"/><a:satMod val="300000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:tint val="15000"/><a:satMod val="350000"/></a:schemeClr></a:gs></a:gsLst><a:lin ang="16200000" scaled="1"/></a:gradFill><a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:shade val="51000"/><a:satMod val="130000"/></a:schemeClr></a:gs><a:gs pos="80000"><a:schemeClr val="phClr"><a:shade val="93000"/><a:satMod val="130000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:shade val="94000"/><a:satMod val="135000"/></a:schemeClr></a:gs></a:gsLst><a:lin ang="16200000" scaled="0"/></a:gradFill></a:fillStyleLst><a:lnStyleLst><a:ln w="9525" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"><a:shade val="95000"/><a:satMod val="105000"/></a:schemeClr></a:solidFill><a:prstDash val="solid"/></a:ln><a:ln w="25400" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/></a:ln><a:ln w="38100" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/></a:ln></a:lnStyleLst><a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst><a:outerShdw blurRad="40000" dist="23000" dir="5400000" rotWithShape="0"><a:srgbClr val="000000"><a:alpha val="35000"/></a:srgbClr></a:outerShdw></a:effectLst></a:effectStyle></a:effectStyleLst><a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:tint val="40000"/><a:satMod val="350000"/></a:schemeClr></a:gs><a:gs pos="40000"><a:schemeClr val="phClr"><a:tint val="45000"/><a:shade val="99000"/><a:satMod val="350000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:shade val="20000"/><a:satMod val="255000"/></a:schemeClr></a:gs></a:gsLst><a:path path="circle"><a:fillToRect l="50000" t="-80000" r="50000" b="180000"/></a:path></a:gradFill></a:bgFillStyleLst></a:fmtScheme></a:themeElements></a:theme>'''

# ======================================================================
# ENSAMBLAR EL PAQUETE
# ======================================================================
if os.path.exists(BUILD):
    shutil.rmtree(BUILD)

files = {
    "[Content_Types].xml": content_types,
    "_rels/.rels": root_rels,
    "xl/workbook.xml": workbook_xml,
    "xl/_rels/workbook.xml.rels": workbook_rels,
    "xl/styles.xml": styles_xml,
    "xl/theme/theme1.xml": theme_xml,
    "xl/worksheets/sheet1.xml": s1.to_xml(),
    "xl/worksheets/sheet2.xml": s2.to_xml(),
    "xl/worksheets/sheet3.xml": s3.to_xml(s3_extra),
    "xl/worksheets/_rels/sheet3.xml.rels": sheet3_rels,
    "xl/drawings/drawing1.xml": drawing_xml,
    "xl/drawings/_rels/drawing1.xml.rels": drawing_rels,
    "xl/charts/chart1.xml": chart1,
    "xl/charts/chart2.xml": chart2,
}

# Validacion XML antes de empaquetar
import xml.dom.minidom as MD
for name, content in files.items():
    if name.endswith(".xml") or name.endswith(".rels"):
        try:
            MD.parseString(content.encode("utf-8"))
        except Exception as e:
            raise SystemExit("XML invalido en {}: {}".format(name, e))

if os.path.exists(OUT):
    os.remove(OUT)
with zipfile.ZipFile(OUT, "w", zipfile.ZIP_DEFLATED) as z:
    # Content_Types primero
    z.writestr("[Content_Types].xml", files.pop("[Content_Types].xml"))
    for name, content in files.items():
        z.writestr(name, content)

# Verificacion final del zip
with zipfile.ZipFile(OUT) as z:
    bad = z.testzip()
sz = os.path.getsize(OUT)
print("OK -> {}  ({} bytes)  testzip={}".format(OUT, sz, bad))
print("Margen Total USD:", margen_total)
print("Adopcion Green %: {:.1%}".format(adopcion))
print("APU global (min):", apu_global)
print("Eco-Corporate prom Semana 3 (ROJO):", week3_eco)
print("Eco prom por semana:", [agg[w]['eco_avg'] for w in weeks])
