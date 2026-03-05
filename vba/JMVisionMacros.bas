Attribute VB_Name = "JMVisionMacros"
Option Explicit

Private Const SHEET_COT As String = "COTIZACION"
Private Const SHEET_PROD As String = "PRODUCTOS"
Private Const SHEET_HIS As String = "HISTORICO_COTIZACIONES"
Private Const SHEET_VTA As String = "VENTAS"

Public Sub NuevaCotizacion()
    Dim ws As Worksheet
    Dim i As Long, pref As String
    Set ws = ThisWorkbook.Worksheets(SHEET_COT)

    pref = IIf(UCase$(Trim$(ws.Range("F1").Value2)) Like "*VENTA*", "VTA-", "COT-")
    ws.Range("H1").Value = pref & Format$(SiguienteCorrelativo(pref), "0000")
    ws.Range("F2").Value = Date

    For i = 8 To 47
        ws.Cells(i, "B").ClearContents
        ws.Cells(i, "C").ClearContents
        ws.Cells(i, "E").ClearContents
        ws.Cells(i, "C").ClearComments
    Next i

    LimpiarMiniaturas ws
    MsgBox "Nueva cotización preparada: " & ws.Range("H1").Value, vbInformation
End Sub

Public Sub GuardarEnHistorico()
    Dim wsC As Worksheet, wsH As Worksheet, nr As Long, items As Long
    Set wsC = ThisWorkbook.Worksheets(SHEET_COT)
    Set wsH = ThisWorkbook.Worksheets(SHEET_HIS)

    nr = wsH.Cells(wsH.Rows.Count, "A").End(xlUp).Row + 1
    items = Application.WorksheetFunction.CountA(wsC.Range("B8:B47"))

    wsH.Cells(nr, "A").Value = wsC.Range("F2").Value
    wsH.Cells(nr, "B").Value = wsC.Range("H1").Value
    wsH.Cells(nr, "C").Value = wsC.Range("F1").Value
    wsH.Cells(nr, "D").Value = wsC.Range("A5").Value
    wsH.Cells(nr, "E").Value = items
    wsH.Cells(nr, "F").Value = wsC.Range("G51").Value

    MsgBox "Guardado en histórico: fila " & nr, vbInformation
End Sub

Public Sub ConvertirAVenta()
    Dim wsC As Worksheet, wsV As Worksheet, nr As Long
    Set wsC = ThisWorkbook.Worksheets(SHEET_COT)
    Set wsV = ThisWorkbook.Worksheets(SHEET_VTA)

    wsC.Range("F1").Value = "NOTA DE VENTA"
    wsC.Range("H1").Value = "VTA-" & Format$(SiguienteCorrelativo("VTA-"), "0000")

    nr = wsV.Cells(wsV.Rows.Count, "A").End(xlUp).Row + 1
    wsV.Cells(nr, "A").Value = wsC.Range("F2").Value
    wsV.Cells(nr, "B").Value = wsC.Range("H1").Value
    wsV.Cells(nr, "C").Value = wsC.Range("A5").Value
    wsV.Cells(nr, "D").Value = wsC.Range("G51").Value
    wsV.Cells(nr, "E").Value = "Emitida"

    MsgBox "Convertida a venta: " & wsC.Range("H1").Value, vbInformation
End Sub

Public Sub ExportarPDFCarta()
    Dim ws As Worksheet, ruta As String, cliente As String, docCod As String, nombre As String
    Set ws = ThisWorkbook.Worksheets(SHEET_COT)

    ruta = ThisWorkbook.Path & Application.PathSeparator & "PDF"
    If Dir$(ruta, vbDirectory) = vbNullString Then MkDir ruta

    cliente = LimpiarNombreArchivo(Split(ws.Range("A5").Value, "|")(0))
    docCod = IIf(UCase$(Trim$(ws.Range("F1").Value2)) Like "*VENTA*", "VTA", "COT")

    Call AplicarModoFotos(ws)

    nombre = ruta & Application.PathSeparator & docCod & "-" & ws.Range("H1").Value & "-" & cliente & ".pdf"
    ws.ExportAsFixedFormat Type:=xlTypePDF, Filename:=nombre, Quality:=xlQualityStandard, _
                           IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=False

    MsgBox "PDF generado: " & nombre, vbInformation
End Sub


Public Sub ActualizarCodigosYFotos()
    Dim i As Long
    For i = 8 To 47
        RecalcularFila i
    Next i
    MsgBox "Se actualizaron códigos y miniaturas.", vbInformation
End Sub

Public Sub RecalcularFila(ByVal fila As Long)
    Dim ws As Worksheet, cod As String
    Set ws = ThisWorkbook.Worksheets(SHEET_COT)
    cod = Trim$(ws.Cells(fila, "B").Value2)

    If cod = vbNullString Then
        ws.Cells(fila, "D").ClearContents
        ws.Cells(fila, "F").ClearContents
        ws.Cells(fila, "G").ClearContents
        Exit Sub
    End If

    If Application.WorksheetFunction.CountIf(ThisWorkbook.Worksheets(SHEET_PROD).Range("A:A"), cod) = 0 Then
        ws.Cells(fila, "B").Interior.Color = RGB(255, 199, 206)
        MsgBox "Código no existe: " & cod, vbExclamation
        Exit Sub
    End If

    ws.Cells(fila, "B").Interior.Pattern = xlNone
    CargarMiniatura ws, fila, cod
End Sub

Private Sub AplicarModoFotos(ByVal ws As Worksheet)
    If UCase$(Trim$(ws.Range("H2").Value2)) = "NO" Then
        ws.Columns("C").Hidden = True
    Else
        ws.Columns("C").Hidden = False
    End If
End Sub

Private Sub CargarMiniatura(ByVal ws As Worksheet, ByVal fila As Long, ByVal codigo As String)
    Dim shp As Shape, rutaBase As String, foto As String, fullPath As String
    Dim celda As Range

    rutaBase = ws.Range("N2").Value2
    If Right$(rutaBase, 1) <> "\" Then rutaBase = rutaBase & "\"
    foto = BuscarFotoProducto(codigo)
    If foto = vbNullString Then Exit Sub

    fullPath = rutaBase & foto
    If Dir$(fullPath) = vbNullString Then Exit Sub

    On Error Resume Next
    ws.Shapes("IMG_" & fila).Delete
    On Error GoTo 0

    Set celda = ws.Cells(fila, "C")
    Set shp = ws.Shapes.AddPicture(fullPath, msoFalse, msoTrue, celda.Left + 2, celda.Top + 2, -1, -1)
    shp.Name = "IMG_" & fila
    shp.LockAspectRatio = msoTrue
    If shp.Width > celda.Width - 4 Then shp.Width = celda.Width - 4
    If shp.Height > celda.Height - 4 Then shp.Height = celda.Height - 4
    shp.Top = celda.Top + (celda.Height - shp.Height) / 2
    shp.Left = celda.Left + (celda.Width - shp.Width) / 2
End Sub

Private Sub LimpiarMiniaturas(ByVal ws As Worksheet)
    Dim shp As Shape
    For Each shp In ws.Shapes
        If Left$(shp.Name, 4) = "IMG_" Then shp.Delete
    Next shp
End Sub

Private Function BuscarFotoProducto(ByVal codigo As String) As String
    Dim ws As Worksheet, m As Variant
    Set ws = ThisWorkbook.Worksheets(SHEET_PROD)
    m = Application.Match(codigo, ws.Range("A:A"), 0)
    If IsError(m) Then
        BuscarFotoProducto = vbNullString
    Else
        BuscarFotoProducto = ws.Cells(CLng(m), "D").Value2
    End If
End Function

Private Function SiguienteCorrelativo(ByVal prefijo As String) As Long
    Dim wsH As Worksheet, wsV As Worksheet, maxNum As Long
    maxNum = MaxCorrelativoEnHoja(ThisWorkbook.Worksheets(SHEET_HIS), prefijo)
    maxNum = Application.Max(maxNum, MaxCorrelativoEnHoja(ThisWorkbook.Worksheets(SHEET_VTA), prefijo))
    SiguienteCorrelativo = maxNum + 1
End Function

Private Function MaxCorrelativoEnHoja(ByVal ws As Worksheet, ByVal pref As String) As Long
    Dim r As Long, v As String, n As Long, maxN As Long
    For r = 3 To ws.Cells(ws.Rows.Count, "B").End(xlUp).Row
        v = CStr(ws.Cells(r, "B").Value2)
        If Left$(v, Len(pref)) = pref Then
            n = Val(Mid$(v, Len(pref) + 1))
            If n > maxN Then maxN = n
        End If
    Next r
    MaxCorrelativoEnHoja = maxN
End Function

Private Function LimpiarNombreArchivo(ByVal txt As String) As String
    Dim x As Variant
    LimpiarNombreArchivo = Trim$(txt)
    For Each x In Array("\", "/", ":", "*", "?", Chr$(34), "<", ">", "|")
        LimpiarNombreArchivo = Replace(LimpiarNombreArchivo, CStr(x), "_")
    Next x
End Function
