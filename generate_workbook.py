import datetime
import os
import zipfile
from xml.sax.saxutils import escape

OUT_FILES = ["Cotizaciones_JMVISION_LIMPIO.xlsm", "Cotizaciones_JMVISION.xlsm"]
LOGO_PATH = "assets/JM-LOGO LARGO OFICIAL.png"


def col(n: int) -> str:
    s = ""
    while n:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


def cell(ref, value, t="str", style=None):
    s = f' s="{style}"' if style is not None else ""
    if value is None:
        return f'<c r="{ref}"{s}/>'
    if t == "str":
        return f'<c r="{ref}" t="inlineStr"{s}><is><t>{escape(str(value))}</t></is></c>'
    if t == "n":
        return f'<c r="{ref}"{s}><v>{value}</v></c>'
    if t == "f":
        return f'<c r="{ref}"{s}><f>{escape(value)}</f></c>'
    return ""


def sheet_xml(rows, cols=None, merges=None, dvals=None, page_setup=None, drawing=False, row_breaks=None):
    out = [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">',
    ]
    if cols:
        out += ["<cols>"] + cols + ["</cols>"]
    out += ["<sheetData>"] + rows + ["</sheetData>"]
    if merges:
        out.append(f'<mergeCells count="{len(merges)}">')
        out += [f'<mergeCell ref="{m}"/>' for m in merges]
        out.append("</mergeCells>")
    if dvals:
        out.append(f'<dataValidations count="{len(dvals)}">')
        out += dvals
        out.append("</dataValidations>")
    if row_breaks:
        out.append(f'<rowBreaks count="{len(row_breaks)}" manualBreakCount="{len(row_breaks)}">')
        out += [f'<brk id="{r}" man="1" max="16383" min="0"/>' for r in row_breaks]
        out.append("</rowBreaks>")
    if page_setup:
        out += page_setup
    if drawing:
        out.append('<drawing r:id="rId1"/>')
    out.append("</worksheet>")
    return "".join(out)


productos_headers = [
    "CODIGO",
    "NOMBRE",
    "DETALLE",
    "FOTO",
    "COSTO_MAYORISTA",
    "PRECIO_FINAL",
    "PRECIO_TECNICO",
    "ACTIVO",
]
productos = [
    ["CAM-2MP-DOMO", "Camara Domo 2MP", "Lente 2.8mm, IR 30m, metalica", "CAM-2MP-DOMO.jpg", 32, 58, 50, "S"],
    ["CAM-2MP-BALA", "Camara Bala 2MP", "IP66, IR 40m, lente 3.6mm", "CAM-2MP-BALA.jpg", 35, 62, 54, "S"],
    ["CAM-5MP-DOMO", "Camara Domo 5MP", "WDR, IR 30m, lente varifocal", "CAM-5MP-DOMO.jpg", 48, 84, 74, "S"],
    ["DVR-8CH-XM", "Grabador DVR 8 Canales", "H.265+, 1080N, HDMI/VGA", "DVR-8CH-XM.jpg", 75, 130, 115, "S"],
    ["NVR-8CH-POE", "Grabador NVR 8CH PoE", "PoE integrado, 4K, ONVIF", "NVR-8CH-POE.jpg", 110, 188, 168, "S"],
    ["HDD-1TB-WD", "Disco Duro 1TB", "Especial videovigilancia 24/7", "HDD-1TB-WD.jpg", 38, 62, 55, "S"],
    ["HDD-2TB-WD", "Disco Duro 2TB", "Especial videovigilancia 24/7", "HDD-2TB-WD.jpg", 58, 95, 84, "S"],
    ["FUENTE-12V10A", "Fuente 12V 10A", "Fuente metalica regulada", "FUENTE-12V10A.jpg", 18, 32, 28, "S"],
    ["CONECTOR-BNC", "Conector BNC", "BNC macho a tornillo", "CONECTOR-BNC.jpg", 0.35, 1, 0.8, "S"],
    ["CONECTOR-DC", "Conector DC", "Conector DC macho/hembra", "CONECTOR-DC.jpg", 0.3, 0.9, 0.7, "S"],
    ["UTP-CAT5E", "Cable UTP Cat5e", "Bobina 305m cobre CCA", "UTP-CAT5E.jpg", 55, 92, 80, "S"],
    ["CAJA-PASO", "Caja de paso", "Caja plastica 10x10 cm", "CAJA-PASO.jpg", 1.4, 3.5, 3, "S"],
    ["BALUN-HD", "Balun HD", "Transceptor pasivo HD", "BALUN-HD.jpg", 1.2, 3.2, 2.8, "S"],
    ["SWITCH-8P", "Switch 8 Puertos", "Fast Ethernet 10/100", "SWITCH-8P.jpg", 14, 27, 23, "S"],
    ["RACK-6U", "Rack mural 6U", "Gabinete mural con llave", "RACK-6U.jpg", 45, 80, 70, "S"],
    ["MO-PUNTO", "Mano de obra por punto", "Instalacion CCTV por punto", "MO-PUNTO.jpg", 0, 180, 0, "S"],
]

rows = [f'<row r="1">{cell("A1", "PRODUCTOS", "str", 1)}</row>']
rows.append('<row r="2">' + "".join(cell(f"{col(i+1)}2", h, "str", 2) for i, h in enumerate(productos_headers)) + "</row>")
for r, p in enumerate(productos, start=3):
    vals = []
    for i, v in enumerate(p, start=1):
        vals.append(cell(f"{col(i)}{r}", v, "n" if isinstance(v, (int, float)) else "str", 0))
    rows.append(f'<row r="{r}">' + "".join(vals) + "</row>")
productos_xml = sheet_xml(rows)

# CLIENTES
rows = [
    f'<row r="1">{cell("A1", "CLIENTES", "str", 1)}</row>',
    '<row r="2">' + "".join([
        cell("A2", "NOMBRE", "str", 2), cell("B2", "CONTACTO", "str", 2), cell("C2", "CI/NIT", "str", 2)
    ]) + '</row>',
    '<row r="3">' + "".join([
        cell("A3", "Jesus Tarqui", "str"), cell("B3", "67278181", "str"), cell("C3", "CI: 1781818", "str")
    ]) + '</row>',
]
clientes_xml = sheet_xml(rows)

# COTIZACION
crows = []
crows.append('<row r="1">' + cell("A1", " ", "str", 0) + cell("E1", "Documento", "str", 2) + cell("F1", "COTIZACIÓN", "str", 3) + cell("G1", "Nro", "str", 2) + cell("H1", "COT-0102", "str", 3) + cell("N1", "Final", "str", 3) + cell("O1", "TipoCliente interno", "str", 2) + '</row>')
crows.append('<row r="2">' + cell("E2", "Fecha", "str", 2) + cell("F2", datetime.date.today().isoformat(), "str", 3) + cell("G2", "Imprimir con fotos", "str", 2) + cell("H2", "Sí", "str", 3) + cell("N2", "C:\\JMVISION_FOTOS\\", "str", 3) + cell("O2", "RutaFotos", "str", 2) + '</row>')
crows.append('<row r="4">' + cell("A4", "CLIENTE | CONTACTO | CI/NIT", "str", 2) + '</row>')
crows.append('<row r="5">' + cell("A5", "Jesus Tarqui | 67278181 | CI: 1781818", "str", 3) + '</row>')

headers = ["ITEM", "CODIGO", "FOTO", "DESCRIPCION", "CANT", "PRECIO UNI", "TOTAL Bs", "COSTO UNI", "COSTO TOT", "GANANCIA", "MARGEN %", "FOTO_ARCH"]
crows.append('<row r="7">' + "".join(cell(f"{col(i)}7", h, "str", 2) for i, h in enumerate(headers, start=1)) + '</row>')

codes = [
    "CAM-2MP-DOMO", "CAM-2MP-BALA", "CAM-5MP-DOMO", "DVR-8CH-XM", "NVR-8CH-POE", "HDD-1TB-WD", "HDD-2TB-WD", "FUENTE-12V10A", "CONECTOR-BNC", "CONECTOR-DC", "UTP-CAT5E", "CAJA-PASO",
    "BALUN-HD", "SWITCH-8P", "RACK-6U", "CAM-2MP-DOMO", "CAM-2MP-BALA", "DVR-8CH-XM", "HDD-1TB-WD", "FUENTE-12V10A", "UTP-CAT5E", "CAJA-PASO", "MO-PUNTO", "MO-PUNTO"
]
qty = [2,2,2,1,1,1,1,1,20,20,1,4,8,1,1,2,2,1,1,1,1,4,4,4]

for idx in range(1, 41):
    r = 7 + idx
    code = codes[idx - 1] if idx <= len(codes) else ""
    q = qty[idx - 1] if idx <= len(qty) else ""
    row = [
        cell(f"A{r}", idx, "n", 0),
        cell(f"B{r}", code, "str", 6),
        cell(f"C{r}", "", "str", 0),
        cell(f"D{r}", f'IF(B{r}="","",INDEX(PRODUCTOS!$B:$B,MATCH(B{r},PRODUCTOS!$A:$A,0))&CHAR(10)&INDEX(PRODUCTOS!$C:$C,MATCH(B{r},PRODUCTOS!$A:$A,0)))', "f", 4),
        cell(f"E{r}", q if q != "" else None, "n" if q != "" else "str", 3 if q != "" else 0),
        cell(f"F{r}", f'IF(B{r}="","",IF($N$1="Final",INDEX(PRODUCTOS!$F:$F,MATCH(B{r},PRODUCTOS!$A:$A,0)),INDEX(PRODUCTOS!$G:$G,MATCH(B{r},PRODUCTOS!$A:$A,0))))', "f", 3),
        cell(f"G{r}", f"IFERROR(E{r}*F{r},0)", "f", 3),
        cell(f"H{r}", f'IF(B{r}="","",INDEX(PRODUCTOS!$E:$E,MATCH(B{r},PRODUCTOS!$A:$A,0)))', "f", 3),
        cell(f"I{r}", f"IFERROR(E{r}*H{r},0)", "f", 3),
        cell(f"J{r}", f"IFERROR(G{r}-I{r},0)", "f", 3),
        cell(f"K{r}", f"IFERROR(J{r}/G{r},0)", "f", 5),
        cell(f"L{r}", f'IF(B{r}="","",INDEX(PRODUCTOS!$D:$D,MATCH(B{r},PRODUCTOS!$A:$A,0)))', "f", 0),
    ]
    crows.append(f'<row r="{r}">' + "".join(row) + "</row>")

crows.append('<row r="49">' + cell("F49", "SUBTOTAL", "str", 2) + cell("G49", "SUM(G8:G47)", "f", 2) + '</row>')
crows.append('<row r="50">' + cell("F50", "IVA %", "str", 2) + cell("G50", 0, "n", 3) + '</row>')
crows.append('<row r="51">' + cell("F51", "TOTAL", "str", 1) + cell("G51", "G49*(1+G50)", "f", 1) + '</row>')

# Internos fuera de impresión

cols = [
    '<col min="1" max="1" width="5" customWidth="1"/>',
    '<col min="2" max="2" width="8" customWidth="1"/>',
    '<col min="3" max="3" width="10" customWidth="1"/>',
    '<col min="4" max="4" width="38" customWidth="1"/>',
    '<col min="5" max="5" width="7" customWidth="1"/>',
    '<col min="6" max="7" width="12" customWidth="1"/>',
    '<col min="8" max="12" width="11" hidden="1" customWidth="1"/>',
    '<col min="14" max="15" width="18" hidden="1" customWidth="1"/>',
]

dvals = [
    '<dataValidation type="list" allowBlank="1" sqref="F1"><formula1>"COTIZACIÓN,NOTA DE VENTA"</formula1></dataValidation>',
    '<dataValidation type="list" allowBlank="1" sqref="H2"><formula1>"Sí,No"</formula1></dataValidation>',
    '<dataValidation type="list" allowBlank="1" sqref="N1"><formula1>"Final,Técnico"</formula1></dataValidation>',
]

page_setup = [
    '<printOptions horizontalCentered="0" verticalCentered="0"/>',
    '<pageMargins left="0.25" right="0.25" top="0.35" bottom="0.4" header="0.2" footer="0.2"/>',
    '<pageSetup paperSize="1" orientation="portrait" fitToWidth="1" fitToHeight="0"/>',
]

cot_xml = sheet_xml(crows, cols=cols, merges=["A1:D2", "A4:G4", "A5:G5"], dvals=dvals, page_setup=page_setup, drawing=True, row_breaks=[33])

# HISTORICO
hrows = [
    f'<row r="1">{cell("A1", "HISTORICO_COTIZACIONES", "str", 1)}</row>',
    '<row r="2">' + "".join(cell(f"{col(i)}2", h, "str", 2) for i, h in enumerate(["FECHA", "NRO", "DOC", "CLIENTE", "ITEMS", "TOTAL"], start=1)) + "</row>",
    '<row r="3">' + "".join([cell("A3", datetime.date.today().isoformat(), "str"), cell("B3", "COT-0101", "str"), cell("C3", "COTIZACIÓN", "str"), cell("D3", "Jesus Tarqui", "str"), cell("E3", 12, "n"), cell("F3", 1840, "n")]) + "</row>",
    '<row r="4">' + "".join([cell("A4", datetime.date.today().isoformat(), "str"), cell("B4", "COT-0102", "str"), cell("C4", "COTIZACIÓN", "str"), cell("D4", "Jesus Tarqui", "str"), cell("E4", 24, "n"), cell("F4", "COTIZACION!G51", "f")]) + "</row>",
]
hist_xml = sheet_xml(hrows)

# VENTAS
vrows = [
    f'<row r="1">{cell("A1", "VENTAS", "str", 1)}</row>',
    '<row r="2">' + "".join(cell(f"{col(i)}2", h, "str", 2) for i, h in enumerate(["FECHA", "NRO", "CLIENTE", "TOTAL", "ESTADO"], start=1)) + "</row>",
    '<row r="3">' + "".join([cell("A3", datetime.date.today().isoformat(), "str"), cell("B3", "VTA-0042", "str"), cell("C3", "Cliente de ejemplo", "str"), cell("D3", 990, "n"), cell("E3", "Pagada", "str")]) + "</row>",
]
ventas_xml = sheet_xml(vrows)

# DASHBOARD
drows = [
    f'<row r="1">{cell("A1", "DASHBOARD", "str", 1)}</row>',
    '<row r="3">' + cell("A3", "Cotizaciones por mes", "str", 2) + cell("B3", 'COUNTIFS(HISTORICO_COTIZACIONES!A:A,">="&DATE(YEAR(TODAY()),MONTH(TODAY()),1))', "f", 3) + '</row>',
    '<row r="4">' + cell("A4", "Ventas por mes", "str", 2) + cell("B4", 'COUNTIFS(VENTAS!A:A,">="&DATE(YEAR(TODAY()),MONTH(TODAY()),1))', "f", 3) + '</row>',
    '<row r="5">' + cell("A5", "% Conversión", "str", 2) + cell("B5", "IFERROR(B4/B3,0)", "f", 5) + '</row>',
]
dash_xml = sheet_xml(drows)

styles = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<fonts count="4"><font><sz val="10"/><name val="Calibri"/></font><font><b/><sz val="14"/><name val="Calibri"/></font><font><b/><sz val="10"/><name val="Calibri"/></font><font><sz val="9"/><name val="Calibri"/></font></fonts>
<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>
<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
<cellXfs count="7">
<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
<xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/>
<xf numFmtId="0" fontId="2" fillId="0" borderId="0" xfId="0" applyFont="1"/>
<xf numFmtId="4" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>
<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment wrapText="1" vertical="top"/></xf>
<xf numFmtId="10" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>
<xf numFmtId="0" fontId="3" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment textRotation="90" horizontal="center" vertical="center"/></xf>
</cellXfs>
<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
</styleSheet>'''

content_types = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Default Extension="png" ContentType="image/png"/>
<Override PartName="/xl/workbook.xml" ContentType="application/vnd.ms-excel.sheet.macroEnabled.main+xml"/>
<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
<Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
<Override PartName="/xl/worksheets/sheet3.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
<Override PartName="/xl/worksheets/sheet4.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
<Override PartName="/xl/worksheets/sheet5.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
<Override PartName="/xl/worksheets/sheet6.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
<Override PartName="/xl/drawings/drawing1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>
</Types>'''

rels_root = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>'''

wb = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<sheets>
<sheet name="PRODUCTOS" sheetId="1" r:id="rId1"/>
<sheet name="CLIENTES" sheetId="2" r:id="rId2"/>
<sheet name="COTIZACION" sheetId="3" r:id="rId3"/>
<sheet name="HISTORICO_COTIZACIONES" sheetId="4" r:id="rId4"/>
<sheet name="VENTAS" sheetId="5" r:id="rId5"/>
<sheet name="DASHBOARD" sheetId="6" r:id="rId6"/>
</sheets>
<definedNames>
<definedName name="_xlnm.Print_Area" localSheetId="2">COTIZACION!$A$1:$G$51</definedName>
<definedName name="_xlnm.Print_Titles" localSheetId="2">COTIZACION!$7:$7</definedName>
</definedNames>
</workbook>'''

wb_rels = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/>
<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet3.xml"/>
<Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet4.xml"/>
<Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet5.xml"/>
<Relationship Id="rId6" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet6.xml"/>
<Relationship Id="rId7" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>'''

sheet3_rels = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/>
</Relationships>'''

drawing1 = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<xdr:twoCellAnchor editAs="oneCell">
<xdr:from><xdr:col>0</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>0</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>
<xdr:to><xdr:col>3</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>3</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>
<xdr:pic>
<xdr:nvPicPr><xdr:cNvPr id="2" name="JM-LOGO"/><xdr:cNvPicPr/></xdr:nvPicPr>
<xdr:blipFill><a:blip r:embed="rId1"/><a:stretch><a:fillRect/></a:stretch></xdr:blipFill>
<xdr:spPr><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr>
</xdr:pic>
<xdr:clientData/>
</xdr:twoCellAnchor>
</xdr:wsDr>'''

drawing_rels = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png"/>
</Relationships>'''

for out in OUT_FILES:
    with zipfile.ZipFile(out, "w", compression=zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", content_types)
        z.writestr("_rels/.rels", rels_root)
        z.writestr("xl/workbook.xml", wb)
        z.writestr("xl/_rels/workbook.xml.rels", wb_rels)
        z.writestr("xl/styles.xml", styles)
        z.writestr("xl/worksheets/sheet1.xml", productos_xml)
        z.writestr("xl/worksheets/sheet2.xml", clientes_xml)
        z.writestr("xl/worksheets/sheet3.xml", cot_xml)
        z.writestr("xl/worksheets/sheet4.xml", hist_xml)
        z.writestr("xl/worksheets/sheet5.xml", ventas_xml)
        z.writestr("xl/worksheets/sheet6.xml", dash_xml)
        z.writestr("xl/worksheets/_rels/sheet3.xml.rels", sheet3_rels)
        z.writestr("xl/drawings/drawing1.xml", drawing1)
        z.writestr("xl/drawings/_rels/drawing1.xml.rels", drawing_rels)

        if os.path.exists(LOGO_PATH):
            with open(LOGO_PATH, "rb") as f:
                z.writestr("xl/media/image1.png", f.read())

    print("generated", out)
