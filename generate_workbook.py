import zipfile, os, datetime
from xml.sax.saxutils import escape

OUT='Cotizaciones_JMVISION.xlsm'

# Helpers

def col(n):
    s=''
    while n:
        n, r = divmod(n-1, 26)
        s=chr(65+r)+s
    return s

def cell(ref, value, t='str', style=None):
    s = f' s="{style}"' if style is not None else ''
    if value is None:
        return f'<c r="{ref}"{s}/>'
    if t=='str':
        return f'<c r="{ref}" t="inlineStr"{s}><is><t>{escape(str(value))}</t></is></c>'
    if t=='n':
        return f'<c r="{ref}"{s}><v>{value}</v></c>'
    if t=='f':
        return f'<c r="{ref}"{s}><f>{escape(value)}</f></c>'
    return ''

def sheet_xml(name, rows, merges=None, cols=None, data_validations=None, page_setup=True):
    out=['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
    '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">']
    if cols:
        out.append('<cols>')
        out.extend(cols)
        out.append('</cols>')
    out.append('<sheetData>')
    out.extend(rows)
    out.append('</sheetData>')
    if merges:
        out.append(f'<mergeCells count="{len(merges)}">')
        out.extend([f'<mergeCell ref="{m}"/>' for m in merges])
        out.append('</mergeCells>')
    if data_validations:
        out.append(f'<dataValidations count="{len(data_validations)}">')
        out.extend(data_validations)
        out.append('</dataValidations>')
    if page_setup:
        out.append('<pageMargins left="0.3" right="0.3" top="0.5" bottom="0.5" header="0.3" footer="0.3"/>')
        out.append('<pageSetup paperSize="1" orientation="portrait" fitToWidth="1" fitToHeight="0"/>')
    out.append('</worksheet>')
    return ''.join(out)

# PRODUCTOS
productos_headers=['CODIGO','NOMBRE','DETALLE','COSTO_MAYORISTA','PRECIO_FINAL','PRECIO_TECNICO','ACTIVO']
productos=[
['CAM-2MP-DOMO','Camara Domo 2MP','Lente 2.8mm, IR 30m, metalica',32,58,50,'S'],
['CAM-2MP-BALA','Camara Bala 2MP','IP66, IR 40m, lente 3.6mm',35,62,54,'S'],
['CAM-5MP-DOMO','Camara Domo 5MP','WDR, IR 30m, lente varifocal',48,84,74,'S'],
['DVR-8CH-XM','Grabador DVR 8 Canales','H.265+, 1080N, salida HDMI/VGA',75,130,115,'S'],
['NVR-8CH-POE','Grabador NVR 8CH PoE','PoE integrado, 4K, ONVIF',110,188,168,'S'],
['HDD-1TB-WD','Disco Duro 1TB','Especial videovigilancia 24/7',38,62,55,'S'],
['HDD-2TB-WD','Disco Duro 2TB','Especial videovigilancia 24/7',58,95,84,'S'],
['FUENTE-12V10A','Fuente 12V 10A','Fuente metalica regulada',18,32,28,'S'],
['CONECTOR-BNC','Conector BNC','Conector BNC macho a tornillo',0.35,1,0.8,'S'],
['CONECTOR-DC','Conector DC','Conector DC macho/hembra',0.3,0.9,0.7,'S'],
['UTP-CAT5E','Cable UTP Cat5e','Bobina 305m cobre CCA',55,92,80,'S'],
['CAJA-PASO','Caja de paso','Caja plastica 10x10 cm',1.4,3.5,3,'S'],
['BALUN-HD','Balun HD','Transceptor pasivo HD CVI/TVI/AHD',1.2,3.2,2.8,'S'],
['SWITCH-8P','Switch 8 Puertos','Fast Ethernet 10/100',14,27,23,'S'],
['RACK-6U','Rack mural 6U','Gabinete mural con llave',45,80,70,'S'],
['MO-PUNTO','Mano de obra por punto','Instalacion de sistema CCTV (por punto instalado)',0,180,0,'S']
]
rows=[]
# row1 title
rows.append('<row r="1">'+cell('A1','PRODUCTOS JM-VISION','str',1)+'</row>')
rows.append('<row r="2">'+''.join(cell(f'{col(i+1)}2',h,'str',2) for i,h in enumerate(productos_headers))+'</row>')
for r,p in enumerate(productos, start=3):
    c=[]
    for i,v in enumerate(p, start=1):
        t='n' if isinstance(v,(int,float)) else 'str'
        c.append(cell(f'{col(i)}{r}',v,t,0))
    rows.append(f'<row r="{r}">'+''.join(c)+'</row>')
productos_xml=sheet_xml('PRODUCTOS',rows)

# CLIENTES
rows=['<row r="1">'+cell('A1','CLIENTES','str',1)+'</row>',
      '<row r="2">'+''.join([cell('A2','NOMBRE','str',2),cell('B2','NIT/CI','str',2),cell('C2','CIUDAD','str',2),cell('D2','CONTACTO','str',2),cell('E2','EMAIL','str',2)])+'</row>',
      '<row r="3">'+''.join([cell('A3','Cliente Demo','str'),cell('B3','1234567','str'),cell('C3','La Paz','str'),cell('D3','70000000','str'),cell('E3','demo@cliente.com','str')])+'</row>'
]
clientes_xml=sheet_xml('CLIENTES',rows)

# COTIZACION
rows=[]
rows.append('<row r="1">'+cell('A1','JM-VISION | COTIZACION','str',1)+cell('F1','Nro Cotizacion','str',2)+cell('G1','COT-0001','str',3)+'</row>')
rows.append('<row r="2">'+cell('A2','Cliente','str',2)+cell('B2','Cliente Demo','str',3)+cell('D2','NIT/CI','str',2)+cell('E2','1234567','str',3)+cell('F2','Fecha','str',2)+cell('G2',datetime.date.today().isoformat(),'str',3)+'</row>')
rows.append('<row r="3">'+cell('A3','Ciudad','str',2)+cell('B3','La Paz','str',3)+cell('D3','Forma de pago','str',2)+cell('E3','Contado','str',3)+cell('F3','TipoCliente','str',2)+cell('G3','Final','str',3)+'</row>')
rows.append('<row r="4">'+cell('A4','Contacto','str',2)+cell('B4','70000000','str',3)+cell('D4','Direccion','str',2)+cell('E4','Av. Ejemplo 123','str',3)+cell('F4','Email','str',2)+cell('G4','demo@cliente.com','str',3)+'</row>')
rows.append('<row r="6">'+''.join(cell(f'{c}6',v,'str',2) for c,v in zip(['A','B','C','D','E','F','G','H','I','J'],['ITEM','CODIGO','DESCRIPCION','CANTIDAD','PRECIO UNI','TOTAL Bs','COSTO UNI','COSTO TOTAL','GANANCIA','MARGEN %']))+'</row>')
start=7
sample_codes=['CAM-2MP-DOMO','CAM-2MP-BALA','DVR-8CH-XM','HDD-1TB-WD','FUENTE-12V10A','UTP-CAT5E','MO-PUNTO','CAJA-PASO']
sample_qty=[4,4,1,1,1,1,4,4]
for i in range(1,21):
    r=start+i-1
    code=sample_codes[i-1] if i<=len(sample_codes) else ''
    qty=sample_qty[i-1] if i<=len(sample_qty) else ''
    c=[cell(f'A{r}',i,'n',0), cell(f'B{r}',code,'str',3)]
    c.append(cell(f'C{r}',f'IF(B{r}="","",INDEX(PRODUCTOS!$B:$B,MATCH(B{r},PRODUCTOS!$A:$A,0))&CHAR(10)&INDEX(PRODUCTOS!$C:$C,MATCH(B{r},PRODUCTOS!$A:$A,0)))','f',4))
    c.append(cell(f'D{r}',qty if qty!='' else None,'n' if qty!='' else 'str',3 if qty!='' else 0))
    c.append(cell(f'E{r}',f'IF(B{r}="","",IF($G$3="Final",INDEX(PRODUCTOS!$E:$E,MATCH(B{r},PRODUCTOS!$A:$A,0)),INDEX(PRODUCTOS!$F:$F,MATCH(B{r},PRODUCTOS!$A:$A,0))))','f',3))
    c.append(cell(f'F{r}',f'IFERROR(D{r}*E{r},0)','f',3))
    c.append(cell(f'G{r}',f'IF(B{r}="","",INDEX(PRODUCTOS!$D:$D,MATCH(B{r},PRODUCTOS!$A:$A,0)))','f',3))
    c.append(cell(f'H{r}',f'IFERROR(D{r}*G{r},0)','f',3))
    c.append(cell(f'I{r}',f'IFERROR(F{r}-H{r},0)','f',3))
    c.append(cell(f'J{r}',f'IFERROR(I{r}/F{r},0)','f',5))
    rows.append(f'<row r="{r}">'+''.join(c)+'</row>')
rows.append('<row r="29">'+cell('E29','SUBTOTAL','str',2)+cell('F29','SUM(F7:F26)','f',2)+'</row>')
rows.append('<row r="30">'+cell('E30','IVA %','str',2)+cell('F30',0,'n',3)+'</row>')
rows.append('<row r="31">'+cell('E31','TOTAL','str',1)+cell('F31','F29*(1+F30)','f',1)+cell('H31','Costo total','str',2)+cell('I31','SUM(H7:H26)','f',2)+cell('H32','Ganancia total','str',2)+cell('I32','SUM(I7:I26)','f',2)+cell('H33','Margen total','str',2)+cell('I33','IFERROR(I32/F31,0)','f',2)+'</row>')
rows.append('<row r="35">'+cell('A35','[ BOTON ] NUEVA COTIZACION','str',1)+cell('C35','[ BOTON ] GUARDAR EN HISTORICO','str',1)+cell('E35','[ BOTON ] CONVERTIR A VENTA','str',1)+cell('G35','[ BOTON ] EXPORTAR PDF (CARTA)','str',1)+'</row>')
cols=[
'<col min="1" max="1" width="6" customWidth="1"/>',
'<col min="2" max="2" width="16" customWidth="1"/>',
'<col min="3" max="3" width="45" customWidth="1"/>',
'<col min="4" max="4" width="10" customWidth="1"/>',
'<col min="5" max="6" width="14" customWidth="1"/>',
'<col min="7" max="10" width="12" hidden="1" customWidth="1"/>'
]
dv=['<dataValidation type="list" allowBlank="1" sqref="G3"><formula1>"Final,Técnico"</formula1></dataValidation>']
cot_xml=sheet_xml('COTIZACION',rows,merges=['A1:E1'],cols=cols,data_validations=dv)

# HISTORICO
rows=['<row r="1">'+cell('A1','HISTORICO_COTIZACIONES','str',1)+'</row>',
'<row r="2">'+''.join(cell(f'{col(i)}2',h,'str',2) for i,h in enumerate(['FECHA','NRO','CLIENTE','TIPO','SUBTOTAL','TOTAL'],1))+'</row>',
'<row r="3">'+''.join([cell('A3',datetime.date.today().isoformat(),'str'),cell('B3','COT-0001','str'),cell('C3','Cliente Demo','str'),cell('D3','Final','str'),cell('E3','SUM(COTIZACION!F7:F26)','f'),cell('F3','COTIZACION!F31','f')])+'</row>']
hist_xml=sheet_xml('HISTORICO',rows)

# VENTAS
rows=['<row r="1">'+cell('A1','VENTAS','str',1)+'</row>',
'<row r="2">'+''.join(cell(f'{col(i)}2',h,'str',2) for i,h in enumerate(['FECHA','NRO','CLIENTE','TOTAL','ESTADO'],1))+'</row>',
'<row r="3">'+''.join([cell('A3',datetime.date.today().isoformat(),'str'),cell('B3','COT-0001','str'),cell('C3','Cliente Demo','str'),cell('D3','COTIZACION!F31','f'),cell('E3','Pendiente','str')])+'</row>']
ventas_xml=sheet_xml('VENTAS',rows)

# DASHBOARD
rows=['<row r="1">'+cell('A1','DASHBOARD','str',1)+'</row>',
'<row r="3">'+cell('A3','Cotizaciones por mes','str',2)+cell('B3','=COUNTIFS(HISTORICO!A:A,">="&DATE(YEAR(TODAY()),MONTH(TODAY()),1))','f',3)+'</row>',
'<row r="4">'+cell('A4','Ventas por mes','str',2)+cell('B4','=COUNTIFS(VENTAS!A:A,">="&DATE(YEAR(TODAY()),MONTH(TODAY()),1))','f',3)+'</row>',
'<row r="5">'+cell('A5','% Conversion','str',2)+cell('B5','=IFERROR(B4/B3,0)','f',5)+'</row>',
'<row r="7">'+cell('A7','Nota: Graficos simples pueden crearse en Excel seleccionando esta tabla resumen.','str',3)+'</row>']
dash_xml=sheet_xml('DASHBOARD',rows)

content_types='''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/xl/workbook.xml" ContentType="application/vnd.ms-excel.sheet.macroEnabled.main+xml"/>
<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
<Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
<Override PartName="/xl/worksheets/sheet3.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
<Override PartName="/xl/worksheets/sheet4.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
<Override PartName="/xl/worksheets/sheet5.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
<Override PartName="/xl/worksheets/sheet6.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
</Types>'''
rels='''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>'''
wb='''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
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
<definedName name="_xlnm.Print_Area" localSheetId="2">COTIZACION!$A$1:$F$33</definedName>
</definedNames>
</workbook>'''
wb_rels='''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/>
<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet3.xml"/>
<Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet4.xml"/>
<Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet5.xml"/>
<Relationship Id="rId6" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet6.xml"/>
<Relationship Id="rId7" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>'''
styles='''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<fonts count="3"><font><sz val="11"/><name val="Calibri"/></font><font><b/><sz val="14"/><name val="Calibri"/></font><font><b/><sz val="11"/><name val="Calibri"/></font></fonts>
<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>
<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
<cellXfs count="6">
<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
<xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/>
<xf numFmtId="0" fontId="2" fillId="0" borderId="0" xfId="0" applyFont="1"/>
<xf numFmtId="4" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>
<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment wrapText="1" vertical="top"/></xf>
<xf numFmtId="10" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>
</cellXfs>
<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
</styleSheet>'''

with zipfile.ZipFile(OUT,'w',compression=zipfile.ZIP_DEFLATED) as z:
    z.writestr('[Content_Types].xml',content_types)
    z.writestr('_rels/.rels',rels)
    z.writestr('xl/workbook.xml',wb)
    z.writestr('xl/_rels/workbook.xml.rels',wb_rels)
    z.writestr('xl/styles.xml',styles)
    z.writestr('xl/worksheets/sheet1.xml',productos_xml)
    z.writestr('xl/worksheets/sheet2.xml',clientes_xml)
    z.writestr('xl/worksheets/sheet3.xml',cot_xml)
    z.writestr('xl/worksheets/sheet4.xml',hist_xml)
    z.writestr('xl/worksheets/sheet5.xml',ventas_xml)
    z.writestr('xl/worksheets/sheet6.xml',dash_xml)

print('generated',OUT)
