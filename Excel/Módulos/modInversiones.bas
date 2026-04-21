Option Explicit

' =========================================================
'  Macro principal: versión productiva basada en Debug_V3
'  Tablas principales:
'   - Inversiones_Raw  (origen copiado del archivo)
'   - Inversiones      (salida de Power Query)
'  Vistas de trabajo:
'   - Hoja Datos, tabla Inversiones_Datos
'   - Hoja Acciones ETF, tabla Acciones ETF
' =========================================================
Public Sub ImportarInversionesDesdeXls()

    Dim ruta As Variant
    Dim wbSrc As Workbook
    Dim wsSrc As Worksheet
    Dim wbDest As Workbook
    Dim wsOrigen As Worksheet
    Dim wsInv As Worksheet
    Dim wsDatos As Worksheet
    Dim headerRow As Long
    Dim colFecha As Long
    Dim lastRowSrc As Long
    Dim lastColSrc As Long
    Dim rngOrigen As Range
    Dim loRaw As ListObject
    Dim loInv As ListObject
    Dim shp As Shape
    Dim shpNew As Shape
    Dim c As Long
    Dim r As Long
    Dim lastRowCopia As Long
    Dim lastColCopia As Long

    ruta = Application.GetOpenFilename( _
                "Archivos Excel (*.xls;*.xlsx;*.xlsm;*.xlsb),*.xls;*.xlsx;*.xlsm;*.xlsb", _
                title:="Selecciona el archivo de inversiones")
    If ruta = False Then Exit Sub

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False   ' <<< NUEVO

    Set wbDest = ThisWorkbook

    ' Abrir archivo origen
    On Error Resume Next
    Set wbSrc = Workbooks.Open( _
                    Filename:=ruta, _
                    ReadOnly:=True, _
                    UpdateLinks:=False, _
                    IgnoreReadOnlyRecommended:=True)
    On Error GoTo 0

    If wbSrc Is Nothing Then
        MsgBox "No se pudo abrir el archivo seleccionado.", vbExclamation
        GoTo Salir
    End If

    If EsTextoOHtml(wbSrc) Then
        MsgBox "El archivo seleccionado no es un libro Excel nativo.", vbExclamation
        GoTo Salir
    End If

    On Error Resume Next
    Set wsSrc = wbSrc.Worksheets(1)
    On Error GoTo 0

    If wsSrc Is Nothing Then
        MsgBox "El libro de origen no tiene hojas visibles.", vbExclamation
        GoTo Salir
    End If

    ' 1) Validar que el archivo tenga exactamente los encabezados esperados
    '    en A1:W1 o en A3:W3 (con tildes normalizadas y Pocentaje/Porcentaje Tasa)
    If Not ValidarEncabezadosInversiones(wsSrc, headerRow) Then
        MsgBox "El archivo seleccionado no cumple con el formato de encabezados requerido " & _
               "(Portafolio, Codigo de Orden, Fecha de Operacion, ..., Pocentaje/Porcentaje Tasa " & _
               "en A1 o A3).", vbCritical
        GoTo Salir
    End If

    ' 1.b) Limpiar consultas y hojas antes de reconstruir todo
    LimpiarConsultasYHojas wbDest

    ' 2) Copiar hoja origen completa a Origen_Inversiones (hoja puente)
    On Error Resume Next
    Set wsOrigen = wbDest.Worksheets("Origen_Inversiones")
    On Error GoTo 0

    If wsOrigen Is Nothing Then
        Set wsOrigen = wbDest.Worksheets.Add(After:=wbDest.Worksheets(wbDest.Worksheets.Count))
        wsOrigen.Name = "Origen_Inversiones"
    Else
        wsOrigen.Cells.Clear
        On Error Resume Next
        For Each shp In wsOrigen.Shapes
            shp.Delete
        Next shp
        On Error GoTo 0
    End If

    ' Copia cruda de toda la hoja origen (valores y formatos)
    wsSrc.Cells.Copy wsOrigen.Cells

    ' Eliminar imágenes que se hayan copiado junto con las celdas
    On Error Resume Next
    For Each shp In wsOrigen.Shapes
        shp.Delete
    Next shp
    On Error GoTo 0

    ' Copiar ancho de columnas
    lastColCopia = wsSrc.Cells(1, wsSrc.Columns.Count).End(xlToLeft).Column
    For c = 1 To lastColCopia
        wsOrigen.Columns(c).ColumnWidth = wsSrc.Columns(c).ColumnWidth
    Next c

    ' Copiar alto de filas
    lastRowCopia = wsSrc.Cells(wsSrc.Rows.Count, 1).End(xlUp).Row
    For r = 1 To lastRowCopia
        wsOrigen.Rows(r).RowHeight = wsSrc.Rows(r).RowHeight
    Next r

    ' Copiar shapes (logos, etc.) con mismo tamaño y posición
    For Each shp In wsSrc.Shapes
        shp.Copy
        wsOrigen.Paste
        Set shpNew = wsOrigen.Shapes(wsOrigen.Shapes.Count)
        With shpNew
            .LockAspectRatio = msoTrue
            .Width = shp.Width
            .Height = shp.Height
            .Top = shp.Top
            .Left = shp.Left
            .Placement = shp.Placement
        End With
    Next shp

    ' 3) Crear tabla Inversiones_Raw en Origen_Inversiones usando solo el bloque de datos
    On Error Resume Next
    Do While wsOrigen.ListObjects.Count > 0
        wsOrigen.ListObjects(1).Unlist
    Loop
    On Error GoTo 0

    EliminarTablaPorNombre wbDest, "Inversiones_Raw"

    lastRowSrc = wsOrigen.Cells(wsOrigen.Rows.Count, "A").End(xlUp).Row
    lastColSrc = wsOrigen.Cells(headerRow, wsOrigen.Columns.Count).End(xlToLeft).Column
    If lastColSrc < 23 Then lastColSrc = 23

    Set rngOrigen = wsOrigen.Range(wsOrigen.Cells(headerRow, 1), wsOrigen.Cells(lastRowSrc, lastColSrc))

    Set loRaw = wsOrigen.ListObjects.Add( _
                    SourceType:=xlSrcRange, _
                    Source:=rngOrigen, _
                    XlListObjectHasHeaders:=xlYes)

    On Error Resume Next
    loRaw.Name = "Inversiones_Raw"
    loRaw.TableStyle = "TableStyleLight8"
    loRaw.TableStyle = "Claro 8"
    On Error GoTo 0

    ' Forzar texto en columnas no numéricas ni de fecha en Inversiones_Raw
    ForzarTextoEnNoNumericas loRaw

    ' 4) Crear o actualizar consulta Power Query "Inversiones" desde "Inversiones_Raw"
    CrearOActualizarConsultaPQ_Inversiones "Inversiones", "Inversiones_Raw"

    ' 5) Cargar la consulta en hoja y tabla "Inversiones"
    AsegurarTablaDeConsulta "Inversiones", "Inversiones", "Inversiones"

    ' 6) Corregir fechas, limpiar espacios y formatear números en la tabla final
    On Error Resume Next
    Set wsInv = wbDest.Worksheets("Inversiones")
    On Error GoTo 0

    If Not wsInv Is Nothing Then
        On Error Resume Next
        Set loInv = wsInv.ListObjects("Inversiones")
        If loInv Is Nothing Then
            If wsInv.ListObjects.Count > 0 Then
                Set loInv = wsInv.ListObjects(1)
            End If
        End If
        On Error GoTo 0

        If Not loInv Is Nothing Then
            CorregirFechasTextoDDMM loInv, "Fecha de Operacion"
            CorregirFechasTextoDDMM loInv, "Fecha de Operación"
            CorregirFechasTextoDDMM loInv, "Fecha Liquidacion"
            CorregirFechasTextoDDMM loInv, "Fecha fin Contrato"

            If Not loInv.DataBodyRange Is Nothing Then
                RecortarEspaciosEnRango loInv.DataBodyRange
            End If

            FormatearTablaInversiones loInv
            ' Forzar texto en columnas no numéricas ni de fecha en Inversiones
            ForzarTextoEnNoNumericas loInv

            ConstruirActualizarHojaDatos loInv, wbDest
        End If
    End If

    ' 7) Acciones ETF: consulta, tabla y columnas adicionales
    CrearOActualizarConsultaPQ_AccionesETF
    AsegurarTablaDeConsulta "Acciones ETF", "Acciones ETF", "Acciones ETF"
    ConfigurarExtrasAccionesETF

    ' 8) AutoFit de columnas en todas las hojas del libro (excepto hoja Inicio)
    AutoFitTablasEnLibro wbDest

    ' 9) Reordenar hojas principales
    OrdenarHojasPrincipal wbDest

    ' 10) Activar vista final: Datos si existe, si no Inversiones
    On Error Resume Next
    Set wsDatos = wbDest.Worksheets("Datos")
    On Error GoTo 0

    MsgBox "Proceso de importación de inversiones terminado.", vbInformation

Salir:
    On Error Resume Next
    If Not wbSrc Is Nothing Then wbSrc.Close SaveChanges:=False
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True      ' <<< NUEVO
End Sub


' =========================================================
'  Helpers generales
' =========================================================

Private Function NormalizarTildes(ByVal s As String) As String
    Dim i As Long
    Dim result As String
    Dim ch As Long

    result = ""
    For i = 1 To Len(s)
        ch = AscW(Mid$(s, i, 1))
        Select Case ch
            Case 225: result = result & "a"  ' á  U+00E1
            Case 233: result = result & "e"  ' é  U+00E9
            Case 237: result = result & "i"  ' í  U+00ED
            Case 243: result = result & "o"  ' ó  U+00F3
            Case 250: result = result & "u"  ' ú  U+00FA
            Case 193: result = result & "A"  ' Á  U+00C1
            Case 201: result = result & "E"  ' É  U+00C9
            Case 205: result = result & "I"  ' Í  U+00CD
            Case 211: result = result & "O"  ' Ó  U+00D3
            Case 218: result = result & "U"  ' Ú  U+00DA
            Case Else: result = result & Mid$(s, i, 1)
        End Select
    Next i

    NormalizarTildes = result
End Function


Private Function CanonHeader(ByVal s As String) As String
    Dim res As String
    res = LCase$(NormalizarTildes(Trim$(CStr(s))))
    ' Acepta tanto "Pocentaje" como "Porcentaje"
    res = Replace(res, "pocentaje", "porcentaje")
    CanonHeader = res
End Function

Private Function ValidarEncabezadosInversiones( _
        ByVal ws As Worksheet, _
        ByRef headerRow As Long) As Boolean

    Dim expected(1 To 23) As String
    Dim baseRow As Variant
    Dim c As Long
    Dim okRow As Boolean

    ' Forma canónica de los encabezados (tildes ya normalizadas)
    expected(1) = "portafolio"
    expected(2) = "codigo de orden"
    expected(3) = "fecha de operacion"
    expected(4) = "fecha liquidacion"
    expected(5) = "fecha fin contrato"
    expected(6) = "codigo isin"
    expected(7) = "codigo sbs"
    expected(8) = "monto de operacion original"
    expected(9) = "monto de operacion ml"
    expected(10) = "cantidad"
    expected(11) = "precio"
    expected(12) = "codigo de emisor"
    expected(13) = "operacion"
    expected(14) = "moneda"
    expected(15) = "nemonico"
    expected(16) = "codigo de tercero"
    expected(17) = "tercero"
    expected(18) = "monto nominal operacion original"
    expected(19) = "monto nominal operacion ml"
    expected(20) = "total de comisiones"
    expected(21) = "plaza"
    expected(22) = "tipo tasa"
    expected(23) = "porcentaje tasa" ' aquí se aceptan "Pocentaje" o "Porcentaje"

    ' Solo se aceptan dos casos: encabezados en fila 1 o en fila 3
    For Each baseRow In Array(1, 3)
        okRow = True
        For c = 1 To 23
            If CanonHeader(ws.Cells(baseRow, c).Value) <> expected(c) Then
                okRow = False
                Exit For
            End If
        Next c

        If okRow Then
            headerRow = baseRow
            ValidarEncabezadosInversiones = True
            Exit Function
        End If
    Next baseRow
End Function

Private Function EncontrarPosicionFecha( _
        ByVal ws As Worksheet, _
        ByRef headerRow As Long, _
        ByRef colFecha As Long) As Boolean

    Dim r As Long, c As Long
    Dim valor As String
    Dim objetivo1 As String
    Dim objetivo2 As String
    Dim norm As String

    objetivo1 = LCase$(NormalizarTildes("Fecha de Operación"))
    objetivo2 = LCase$(NormalizarTildes("Fecha de Operacion"))

    For r = 1 To 5
        For c = 1 To 50
            valor = Trim$(CStr(ws.Cells(r, c).Value))
            If valor <> "" Then
                norm = LCase$(NormalizarTildes(valor))
                If norm = objetivo1 Or norm = objetivo2 Then
                    headerRow = r
                    colFecha = c
                    EncontrarPosicionFecha = True
                    Exit Function
                End If
            End If
        Next c
    Next r
End Function

Private Function EsTextoOHtml(wb As Workbook) As Boolean

    Select Case wb.FileFormat
        Case xlHtml, xlCurrentPlatformText, xlTextMac, xlTextWindows, xlTextMSDOS, xlTextPrinter, xlUnicodeText, _
             xlCSV, xlCSVWindows, xlCSVMac, xlCSVMSDOS, xlCSVUTF8
            EsTextoOHtml = True
        Case Else
            EsTextoOHtml = False
    End Select

End Function

Private Sub EliminarTablaPorNombre(ByVal wb As Workbook, ByVal nombreTabla As String)
    Dim ws As Worksheet
    Dim lo As ListObject
    For Each ws In wb.Worksheets
        For Each lo In ws.ListObjects
            If StrComp(lo.Name, nombreTabla, vbTextCompare) = 0 Then
                lo.Unlist
                Exit Sub
            End If
        Next lo
    Next ws
End Sub

' =========================================================
'  Power Query: Inversiones
' =========================================================

Private Sub CrearOActualizarConsultaPQ_Inversiones( _
        ByVal NombreConsulta As String, _
        ByVal nombreTabla As String)

    Dim mCode As String
    Dim q As WorkbookQuery

    mCode = ""
    mCode = mCode & "let" & vbCrLf
    mCode = mCode & "    Origen = Excel.CurrentWorkbook(){[Name=""" & nombreTabla & """]}[Content]," & vbCrLf
    mCode = mCode & "    #""Encabezados normalizados"" =" & vbCrLf
    mCode = mCode & "        Table.TransformColumnNames(" & vbCrLf
    mCode = mCode & "            Origen," & vbCrLf
    mCode = mCode & "            each" & vbCrLf
    mCode = mCode & "                let" & vbCrLf
    mCode = mCode & "                    t0  = Text.Trim(_)," & vbCrLf
    mCode = mCode & "                    t1  = Text.Replace(t0,  ""á"", ""a"")," & vbCrLf
    mCode = mCode & "                    t2  = Text.Replace(t1,  ""é"", ""e"")," & vbCrLf
    mCode = mCode & "                    t3  = Text.Replace(t2,  ""í"", ""i"")," & vbCrLf
    mCode = mCode & "                    t4  = Text.Replace(t3,  ""ó"", ""o"")," & vbCrLf
    mCode = mCode & "                    t5  = Text.Replace(t4,  ""ú"", ""u"")," & vbCrLf
    mCode = mCode & "                    t6  = Text.Replace(t5,  ""Á"", ""A"")," & vbCrLf
    mCode = mCode & "                    t7  = Text.Replace(t6,  ""É"", ""E"")," & vbCrLf
    mCode = mCode & "                    t8  = Text.Replace(t7,  ""Í"", ""I"")," & vbCrLf
    mCode = mCode & "                    t9  = Text.Replace(t8,  ""Ó"", ""O"")," & vbCrLf
    mCode = mCode & "                    t10 = Text.Replace(t9,  ""Ú"", ""U"")" & vbCrLf
    mCode = mCode & "                in" & vbCrLf
    mCode = mCode & "                    t10" & vbCrLf
    mCode = mCode & "        )," & vbCrLf
    mCode = mCode & "    #""Corregir Pocentaje"" =" & vbCrLf
    mCode = mCode & "        Table.TransformColumnNames(" & vbCrLf
    mCode = mCode & "            #""Encabezados normalizados""," & vbCrLf
    mCode = mCode & "            each Text.Replace(_, ""Pocentaje"", ""Porcentaje"")" & vbCrLf
    mCode = mCode & "        )," & vbCrLf
    mCode = mCode & "    #""Tipo cambiado"" =" & vbCrLf
    mCode = mCode & "        Table.TransformColumnTypes(" & vbCrLf
    mCode = mCode & "            #""Corregir Pocentaje""," & vbCrLf
    mCode = mCode & "            {" & vbCrLf
    mCode = mCode & "                {""Portafolio"", type text}," & vbCrLf
    mCode = mCode & "                {""Codigo de Orden"", type text}," & vbCrLf
    mCode = mCode & "                {""Fecha de Operacion"", type date}," & vbCrLf
    mCode = mCode & "                {""Fecha Liquidacion"", type date}," & vbCrLf
    mCode = mCode & "                {""Fecha fin Contrato"", type date}," & vbCrLf
    mCode = mCode & "                {""Codigo ISIN"", type text}," & vbCrLf
    mCode = mCode & "                {""Codigo SBS"", type text}," & vbCrLf
    mCode = mCode & "                {""Monto de Operacion Original"", type number}," & vbCrLf
    mCode = mCode & "                {""Monto de Operacion ML"", type number}," & vbCrLf
    mCode = mCode & "                {""Cantidad"", type number}," & vbCrLf
    mCode = mCode & "                {""Precio"", type number}," & vbCrLf
    mCode = mCode & "                {""Codigo de Emisor"", type text}," & vbCrLf
    mCode = mCode & "                {""Operacion"", type text}," & vbCrLf
    mCode = mCode & "                {""Moneda"", type text}," & vbCrLf
    mCode = mCode & "                {""Nemonico"", type text}," & vbCrLf
    mCode = mCode & "                {""Codigo de Tercero"", type text}," & vbCrLf
    mCode = mCode & "                {""Tercero"", type text}," & vbCrLf
    mCode = mCode & "                {""Monto Nominal Operacion Original"", type number}," & vbCrLf
    mCode = mCode & "                {""Monto Nominal Operacion ML"", type number}," & vbCrLf
    mCode = mCode & "                {""Total de Comisiones"", type number}," & vbCrLf
    mCode = mCode & "                {""Plaza"", type text}," & vbCrLf
    mCode = mCode & "                {""Tipo Tasa"", type text}," & vbCrLf
    mCode = mCode & "                {""Porcentaje Tasa"", type number}" & vbCrLf
    mCode = mCode & "            }," & vbCrLf
    mCode = mCode & "            ""es-PE""" & vbCrLf
    mCode = mCode & "        )" & vbCrLf
    mCode = mCode & "in" & vbCrLf
    mCode = mCode & "    #""Tipo cambiado"""

    On Error Resume Next
    Set q = ThisWorkbook.Queries(NombreConsulta)
    On Error GoTo 0

    If q Is Nothing Then
        ThisWorkbook.Queries.Add Name:=NombreConsulta, Formula:=mCode
    Else
        q.Formula = mCode
    End If
End Sub

' =========================================================
'  Power Query: Acciones ETF
' =========================================================

Private Sub CrearOActualizarConsultaPQ_AccionesETF()

    Dim mCode As String
    Dim q As WorkbookQuery

    mCode = ""
    mCode = mCode & "let" & vbCrLf
    mCode = mCode & "    Origen = Excel.CurrentWorkbook(){[Name=""Inversiones""]}[Content]," & vbCrLf
    mCode = mCode & "    #""Encabezados normalizados"" =" & vbCrLf
    mCode = mCode & "        Table.TransformColumnNames(" & vbCrLf
    mCode = mCode & "            Origen," & vbCrLf
    mCode = mCode & "            each" & vbCrLf
    mCode = mCode & "                let" & vbCrLf
    mCode = mCode & "                    t0  = Text.Trim(_)," & vbCrLf
    mCode = mCode & "                    t1  = Text.Replace(t0,  ""á"", ""a"")," & vbCrLf
    mCode = mCode & "                    t2  = Text.Replace(t1,  ""é"", ""e"")," & vbCrLf
    mCode = mCode & "                    t3  = Text.Replace(t2,  ""í"", ""i"")," & vbCrLf
    mCode = mCode & "                    t4  = Text.Replace(t3,  ""ó"", ""o"")," & vbCrLf
    mCode = mCode & "                    t5  = Text.Replace(t4,  ""ú"", ""u"")," & vbCrLf
    mCode = mCode & "                    t6  = Text.Replace(t5,  ""Á"", ""A"")," & vbCrLf
    mCode = mCode & "                    t7  = Text.Replace(t6,  ""É"", ""E"")," & vbCrLf
    mCode = mCode & "                    t8  = Text.Replace(t7,  ""Í"", ""I"")," & vbCrLf
    mCode = mCode & "                    t9  = Text.Replace(t8,  ""Ó"", ""O"")," & vbCrLf
    mCode = mCode & "                    t10 = Text.Replace(t9,  ""Ú"", ""U"")" & vbCrLf
    mCode = mCode & "                in" & vbCrLf
    mCode = mCode & "                    t10" & vbCrLf
    mCode = mCode & "        )," & vbCrLf
    mCode = mCode & "    #""Corregir Pocentaje"" =" & vbCrLf
    mCode = mCode & "        Table.TransformColumnNames(" & vbCrLf
    mCode = mCode & "            #""Encabezados normalizados""," & vbCrLf
    mCode = mCode & "            each Text.Replace(_, ""Pocentaje"", ""Porcentaje"")" & vbCrLf
    mCode = mCode & "        )," & vbCrLf
    mCode = mCode & "    #""Tipo cambiado"" =" & vbCrLf
    mCode = mCode & "        Table.TransformColumnTypes(" & vbCrLf
    mCode = mCode & "            #""Corregir Pocentaje""," & vbCrLf
    mCode = mCode & "            {" & vbCrLf
    mCode = mCode & "                {""Portafolio"", type text}," & vbCrLf
    mCode = mCode & "                {""Codigo de Orden"", type text}," & vbCrLf
    mCode = mCode & "                {""Fecha de Operacion"", type date}," & vbCrLf
    mCode = mCode & "                {""Fecha Liquidacion"", type date}," & vbCrLf
    mCode = mCode & "                {""Fecha fin Contrato"", type date}," & vbCrLf
    mCode = mCode & "                {""Codigo ISIN"", type text}," & vbCrLf
    mCode = mCode & "                {""Codigo SBS"", type text}," & vbCrLf
    mCode = mCode & "                {""Monto de Operacion Original"", type number}," & vbCrLf
    mCode = mCode & "                {""Monto de Operacion ML"", type number}," & vbCrLf
    mCode = mCode & "                {""Cantidad"", type number}," & vbCrLf
    mCode = mCode & "                {""Precio"", type number}," & vbCrLf
    mCode = mCode & "                {""Codigo de Emisor"", type text}," & vbCrLf
    mCode = mCode & "                {""Operacion"", type text}," & vbCrLf
    mCode = mCode & "                {""Moneda"", type text}," & vbCrLf
    mCode = mCode & "                {""Nemonico"", type text}," & vbCrLf
    mCode = mCode & "                {""Codigo de Tercero"", type text}," & vbCrLf
    mCode = mCode & "                {""Tercero"", type text}," & vbCrLf
    mCode = mCode & "                {""Monto Nominal Operacion Original"", type number}," & vbCrLf
    mCode = mCode & "                {""Monto Nominal Operacion ML"", type number}," & vbCrLf
    mCode = mCode & "                {""Total de Comisiones"", type number}," & vbCrLf
    mCode = mCode & "                {""Plaza"", type text}," & vbCrLf
    mCode = mCode & "                {""Tipo Tasa"", type text}," & vbCrLf
    mCode = mCode & "                {""Porcentaje Tasa"", type number}" & vbCrLf
    mCode = mCode & "            }," & vbCrLf
    mCode = mCode & "            ""es-PE""" & vbCrLf
    mCode = mCode & "        )," & vbCrLf
    mCode = mCode & "    #""Filtrado"" =" & vbCrLf
    mCode = mCode & "        Table.SelectRows(" & vbCrLf
    mCode = mCode & "            #""Tipo cambiado""," & vbCrLf
    mCode = mCode & "            each" & vbCrLf
    mCode = mCode & "                Text.Upper(Text.Trim([Plaza])) = ""NEW YORK""" & vbCrLf
    mCode = mCode & "                and ([Operacion] = ""Compra"" or [Operacion] = ""Venta"")" & vbCrLf
    mCode = mCode & "                and not Text.StartsWith([Nemonico], ""TBILL"")" & vbCrLf
    mCode = mCode & "                and not Text.StartsWith([Nemonico], ""SB"")" & vbCrLf
    mCode = mCode & "                and not Text.StartsWith([Nemonico], ""PERU"")" & vbCrLf
    mCode = mCode & "                and not Text.StartsWith([Nemonico], ""BSMXB"")" & vbCrLf
    mCode = mCode & "                and not Text.StartsWith([Nemonico], ""BANBOG"")" & vbCrLf
    mCode = mCode & "                and not Text.StartsWith([Nemonico], ""COLOM"")" & vbCrLf
    mCode = mCode & "                and not Text.StartsWith([Nemonico], ""ALIPE"")" & vbCrLf
    mCode = mCode & "                and not Text.StartsWith([Nemonico], ""CARLYLE"")" & vbCrLf
    mCode = mCode & "                and not Text.Contains(Text.Upper([Tercero]), ""SURA"")" & vbCrLf
    mCode = mCode & "                and not Text.Contains(Text.Upper([Tercero]), ""FUND"")" & vbCrLf
    mCode = mCode & "                and ([Precio] <> 0)" & vbCrLf
    mCode = mCode & "        )," & vbCrLf
    mCode = mCode & "    #""Columnas quitadas"" =" & vbCrLf
    mCode = mCode & "        Table.RemoveColumns(" & vbCrLf
    mCode = mCode & "            #""Filtrado""," & vbCrLf
    mCode = mCode & "            {" & vbCrLf
    mCode = mCode & "                ""Codigo de Orden""," & vbCrLf
    mCode = mCode & "                ""Fecha Liquidacion""," & vbCrLf
    mCode = mCode & "                ""Fecha fin Contrato""," & vbCrLf
    mCode = mCode & "                ""Codigo ISIN""," & vbCrLf
    mCode = mCode & "                ""Codigo SBS""," & vbCrLf
    mCode = mCode & "                ""Codigo de Tercero""," & vbCrLf
    mCode = mCode & "                ""Monto Nominal Operacion Original""," & vbCrLf
    mCode = mCode & "                ""Monto Nominal Operacion ML""," & vbCrLf
    mCode = mCode & "                ""Total de Comisiones""," & vbCrLf
    mCode = mCode & "                ""Plaza""," & vbCrLf
    mCode = mCode & "                ""Tipo Tasa""," & vbCrLf
    mCode = mCode & "                ""Porcentaje Tasa""" & vbCrLf
    mCode = mCode & "            }" & vbCrLf
    mCode = mCode & "        )," & vbCrLf
    mCode = mCode & "    #""Columnas reordenadas"" =" & vbCrLf
    mCode = mCode & "        Table.ReorderColumns(" & vbCrLf
    mCode = mCode & "            #""Columnas quitadas""," & vbCrLf
    mCode = mCode & "            {" & vbCrLf
    mCode = mCode & "                ""Portafolio""," & vbCrLf
    mCode = mCode & "                ""Fecha de Operacion""," & vbCrLf
    mCode = mCode & "                ""Codigo de Emisor""," & vbCrLf
    mCode = mCode & "                ""Nemonico""," & vbCrLf
    mCode = mCode & "                ""Tercero""," & vbCrLf
    mCode = mCode & "                ""Monto de Operacion Original""," & vbCrLf
    mCode = mCode & "                ""Monto de Operacion ML""," & vbCrLf
    mCode = mCode & "                ""Cantidad""," & vbCrLf
    mCode = mCode & "                ""Precio""," & vbCrLf
    mCode = mCode & "                ""Operacion""," & vbCrLf
    mCode = mCode & "                ""Moneda""" & vbCrLf
    mCode = mCode & "            }" & vbCrLf
    mCode = mCode & "        )" & vbCrLf
    mCode = mCode & "in" & vbCrLf
    mCode = mCode & "    #""Columnas reordenadas"""

    On Error Resume Next
    Set q = ThisWorkbook.Queries("Acciones ETF")
    On Error GoTo 0

    If q Is Nothing Then
        ThisWorkbook.Queries.Add Name:="Acciones ETF", Formula:=mCode
    Else
        q.Formula = mCode
    End If
End Sub


' =========================================================
'  Carga de consultas en hojas
' =========================================================

Private Sub AsegurarTablaDeConsulta( _
        ByVal queryName As String, _
        ByVal sheetName As String, _
        ByVal tableName As String)

    Dim wb As Workbook
    Dim ws As Worksheet
    Dim connString As String
    Dim lo As ListObject
    Dim qt As QueryTable
    Dim qtExisting As QueryTable
    Dim q As WorkbookQuery

    Set wb = ThisWorkbook

    ' Verificar que la consulta de Power Query exista
    On Error Resume Next
    Set q = wb.Queries(queryName)
    On Error GoTo 0

    If q Is Nothing Then
        MsgBox "No se encontró la consulta de Power Query '" & queryName & "'.", vbCritical
        Exit Sub
    End If

    ' Obtener o crear la hoja destino
    On Error Resume Next
    Set ws = wb.Worksheets(sheetName)
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        ws.Name = sheetName
    Else
        ' Quitar tablas anteriores
        On Error Resume Next
        Do While ws.ListObjects.Count > 0
            ws.ListObjects(1).Unlist
        Loop

        ' Quitar consultas anteriores de la hoja
        For Each qtExisting In ws.QueryTables
            qtExisting.Delete
        Next qtExisting

        ' Limpiar contenidos
        ws.Cells.ClearContents
        On Error GoTo 0
    End If

    ' Conexión a Power Query (Microsoft.Mashup.OleDb)
    connString = "OLEDB;Provider=Microsoft.Mashup.OleDb.1;" & _
                 "Data Source=$Workbook$;Location=" & queryName & _
                 ";Extended Properties="""";"

    ' SourceType = 0  (xlSrcExternal)
    On Error Resume Next
    Set lo = ws.ListObjects.Add( _
                SourceType:=0, _
                Source:=connString, _
                Destination:=ws.Range("A1"))
    If Err.Number <> 0 Then
        MsgBox "Error al crear la tabla para la consulta '" & queryName & "':" & vbCrLf & _
               "Error " & Err.Number & " - " & Err.Description, vbCritical
        Err.Clear
        On Error GoTo 0
        Exit Sub
    End If
    On Error GoTo 0

    If lo Is Nothing Then
        MsgBox "No se pudo crear la tabla para la consulta '" & queryName & "'.", vbCritical
        Exit Sub
    End If

    Set qt = lo.QueryTable

    ' Configurar correctamente el comando antes de refrescar
    On Error GoTo ErrRefresh
    With qt
        .CommandType = xlCmdSql
        .CommandText = Array("SELECT * FROM [" & queryName & "]")
        .PreserveFormatting = True
        .AdjustColumnWidth = False
        .RefreshStyle = xlInsertDeleteCells
        .Refresh BackgroundQuery:=False   ' espera a que termine
    End With
    On Error GoTo 0

    ' Renombrar la tabla
    On Error Resume Next
    lo.Name = tableName
    lo.DisplayName = tableName
    lo.TableStyle = "TableStyleLight8"
    lo.TableStyle = "Claro 8"
    On Error GoTo 0

    Exit Sub

ErrRefresh:
    MsgBox "No se pudo actualizar la consulta '" & queryName & "'." & vbCrLf & _
           "Error " & Err.Number & ": " & Err.Description, vbCritical
    On Error GoTo 0
End Sub




' =========================================================
'  Limpieza, formatos y fechas
' =========================================================

Private Sub RecortarEspaciosEnRango(ByVal rng As Range)

    Dim c As Range
    Dim s As String

    For Each c In rng.Cells
        If Not c.HasFormula Then
            If VarType(c.Value2) = vbString Then
                s = c.Value2
                If Len(s) > 0 Then
                    s = Replace(s, Chr(160), " ")
                    s = Trim$(s)
                    If s <> c.Value2 Then
                        c.Value2 = s
                    End If
                End If
            End If
        End If
    Next c

End Sub

Private Sub FormatearColumna(ByVal lo As ListObject, _
                             ByVal nombreColumna As String, _
                             ByVal formato As String)

    Dim lc As ListColumn

    On Error Resume Next
    Set lc = lo.ListColumns(nombreColumna)
    On Error GoTo 0

    If lc Is Nothing Then Exit Sub

    On Error Resume Next
    lc.Range.NumberFormat = formato
    If Not lc.DataBodyRange Is Nothing Then
        lc.DataBodyRange.NumberFormat = formato
    End If
    On Error GoTo 0

End Sub

Private Sub FormatearTablaInversiones(ByVal lo As ListObject)

    FormatearColumna lo, "Monto de Operacion Original", "#,##0.00"
    FormatearColumna lo, "Monto de Operacion ML", "#,##0.00"
    FormatearColumna lo, "Cantidad", "#,##0.0000"
    FormatearColumna lo, "Precio", "#,##0.0000"
    FormatearColumna lo, "Monto Nominal Operacion Original", "#,##0.00"
    FormatearColumna lo, "Monto Nominal Operacion ML", "#,##0.00"
    FormatearColumna lo, "Total de Comisiones", "#,##0.00"
    FormatearColumna lo, "Porcentaje Tasa", "0.0000%"

End Sub

Private Sub FormatearTablaAccionesETF(ByVal lo As ListObject)

    CorregirFechasTextoDDMM lo, "Fecha de Operacion"

    FormatearColumna lo, "Monto de Operacion Original", "#,##0.00"
    FormatearColumna lo, "Monto de Operacion ML", "#,##0.00"
    FormatearColumna lo, "Cantidad", "#,##0.0000"
    FormatearColumna lo, "Precio", "#,##0.0000"

End Sub

' Copia una imagen de la franja superior del origen (A1:B2 aprox.)
' y la pega en A1 del destino, dejando un solo logo pequeño.
Private Sub CopiarImagenA1(ByVal wsOrigen As Worksheet, ByVal wsDestino As Worksheet)

    Dim shp As Shape
    Dim shpLogo As Shape
    Dim shpNew As Shape
    Dim shpDest As Shape
    Dim limSuperior As Double
    Dim i As Long

    If wsOrigen Is Nothing Then Exit Sub
    If wsDestino Is Nothing Then Exit Sub

    ' Buscar en el origen la primera imagen en las filas superiores (1 y 2, columnas 1 a 3)
    For Each shp In wsOrigen.Shapes
        If shp.Type = msoPicture Or shp.Type = msoLinkedPicture Then
            If Not shp.TopLeftCell Is Nothing Then
                If shp.TopLeftCell.Row <= 2 And shp.TopLeftCell.Column <= 3 Then
                    Set shpLogo = shp
                    Exit For
                End If
            End If
        End If
    Next shp

    If shpLogo Is Nothing Then Exit Sub

    ' Borrar cualquier logo previo en la franja superior del destino
    limSuperior = wsDestino.Rows(3).Top   ' todo lo que esté por encima de la fila 3

    For i = wsDestino.Shapes.Count To 1 Step -1
        Set shpDest = wsDestino.Shapes(i)
        If shpDest.Type = msoPicture Or shpDest.Type = msoLinkedPicture Then
            If shpDest.Top < limSuperior Then
                shpDest.Delete
            End If
        End If
    Next i

    ' Copiar el logo desde el origen
    shpLogo.Copy
    wsDestino.Paste
    Set shpNew = wsDestino.Shapes(wsDestino.Shapes.Count)

    With shpNew
        .LockAspectRatio = msoTrue
        .Width = shpLogo.Width
        .Height = shpLogo.Height
        .Top = wsDestino.Range("A1").Top
        .Left = wsDestino.Range("A1").Left
        .Placement = xlMove   ' no cambia de tamaño cuando se ajustan filas/columnas
    End With

End Sub


Private Sub CorregirFechasTextoDDMM(ByVal lo As ListObject, ByVal nombreColumna As String)

    Dim lc As ListColumn
    Dim cel As Range
    Dim v As Variant
    Dim partes() As String
    Dim d As Long, m As Long, y As Long

    On Error Resume Next
    Set lc = lo.ListColumns(nombreColumna)
    On Error GoTo 0
    If lc Is Nothing Then Exit Sub
    If lc.DataBodyRange Is Nothing Then Exit Sub

    For Each cel In lc.DataBodyRange.Cells
        v = cel.Value
        If Not IsEmpty(v) And Not IsNull(v) Then
            If VarType(v) = vbString Then
                partes = Split(Replace(v, "-", "/"), "/")
                If UBound(partes) = 2 Then
                    If IsNumeric(partes(0)) And IsNumeric(partes(1)) And IsNumeric(partes(2)) Then
                        d = CLng(partes(0))
                        m = CLng(partes(1))
                        y = CLng(partes(2))
                        If d >= 1 And d <= 31 And m >= 1 And m <= 12 Then
                            cel.Value = DateSerial(y, m, d)
                        End If
                    End If
                End If
            ElseIf IsDate(v) Then
                cel.Value = CDate(v)
            End If
        End If
    Next cel

    lc.Range.NumberFormat = "dd/mm/yyyy"

End Sub

Private Sub ForzarTextoEnNoNumericas(ByVal lo As ListObject)

    Dim lc As ListColumn
    Dim h As String
    Dim hLower As String
    Dim esFecha As Boolean
    Dim esNumerica As Boolean
    Dim cel As Range

    If lo Is Nothing Then Exit Sub
    If lo.Range Is Nothing Then Exit Sub

    For Each lc In lo.ListColumns
        h = Trim$(lc.Name)
        hLower = LCase$(h)

        esFecha = (InStr(1, hLower, "fecha", vbTextCompare) > 0)

        esNumerica = _
            (Left$(hLower, 5) = "monto") Or _
            (Left$(hLower, 5) = "total") Or _
            (hLower = "cantidad") Or _
            (hLower = "precio") Or _
            (hLower = "porcentaje tasa")

        If Not esFecha And Not esNumerica Then
            On Error Resume Next
            lc.Range.NumberFormat = "@"
            If Not lc.DataBodyRange Is Nothing Then
                For Each cel In lc.DataBodyRange.Cells
                    If Not IsEmpty(cel.Value) And Not IsNull(cel.Value) Then
                        cel.Value = CStr(cel.Value)
                    End If
                Next cel
            End If
            On Error GoTo 0
        End If
    Next lc

End Sub


' =========================================================
'  Construcción de hoja Datos
' =========================================================

Private Sub ConstruirActualizarHojaDatos(loInv As ListObject, wbDest As Workbook)

    Dim wsDatos As Worksheet
    Dim wsInicio As Worksheet
    Dim wsOrigen As Worksheet
    Dim datosNueva As Boolean
    Dim lastRowDatos As Long
    Dim loDatos As ListObject
    Dim hdr As Range
    Dim destRange As Range
    Dim dataRange As Range
    Dim colFecha As Long
    Dim minDate As Variant
    Dim maxDate As Variant
    Dim lastRowOrigen As Long
    Dim lastColOrigen As Long
    Dim c As Long
    Dim i As Long
    Dim lastCellOrigen As Range
    Dim rngFmtSrc As Range
    Dim rngFmtDest As Range

    If loInv Is Nothing Then Exit Sub

    On Error Resume Next
    Set wsDatos = wbDest.Worksheets("Datos")
    Set wsInicio = wbDest.Worksheets("Inicio")
    Set wsOrigen = wbDest.Worksheets("Origen_Inversiones")
    On Error GoTo 0

    If wsDatos Is Nothing Then
        datosNueva = True
        If Not wsInicio Is Nothing Then
            Set wsDatos = wbDest.Worksheets.Add(After:=wsInicio)
        Else
            Set wsDatos = wbDest.Worksheets.Add(After:=wbDest.Worksheets(wbDest.Worksheets.Count))
        End If
        wsDatos.Name = "Datos"
    Else
        datosNueva = False
    End If

    ' Encabezado (ancho/alto, merges y textos principales)
    If Not wsOrigen Is Nothing Then

        lastColOrigen = wsOrigen.Cells(1, wsOrigen.Columns.Count).End(xlToLeft).Column
        lastRowOrigen = wsOrigen.Cells(wsOrigen.Rows.Count, 1).End(xlUp).Row

        ' Copiar anchos de columnas y altura de las 3 primeras filas
        For c = 1 To lastColOrigen
            wsDatos.Columns(c).ColumnWidth = wsOrigen.Columns(c).ColumnWidth
        Next c

        For i = 1 To 3
            wsDatos.Rows(i).RowHeight = wsOrigen.Rows(i).RowHeight
        Next i

        ' Combinar y centrar A1:W1
        With wsDatos.Range("A1:W1")
            .ClearContents
            .ClearFormats
            .Merge
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With

        ' Combinar y centrar E2:W2
        With wsDatos.Range("E2:W2")
            .ClearContents
            .ClearFormats
            .Merge
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With

        ' Copiar contenido + formato de A2:D2
        wsOrigen.Range("A2:D2").Copy
        wsDatos.Range("A2").PasteSpecial Paste:=xlPasteAll
        Application.CutCopyMode = False

        ' Copiar valor de A1
        wsDatos.Range("A1").Value = wsOrigen.Range("A1").Value

        ' Copiar formato de A1
        wsOrigen.Range("A1").Copy
        wsDatos.Range("A1").PasteSpecial Paste:=xlPasteFormats
        Application.CutCopyMode = False

        ' Formato fecha para B2 y D2
        wsDatos.Range("B2").NumberFormatLocal = "dd/mm/aaaa"
        wsDatos.Range("D2").NumberFormatLocal = "dd/mm/aaaa"

    End If

    ' Limpiar solo zona de datos (desde fila 3 hacia abajo) sin tocar cabecera/logo
    If Not datosNueva Then
        lastRowDatos = wsDatos.Cells(wsDatos.Rows.Count, "A").End(xlUp).Row
        If lastRowDatos >= 3 Then
            wsDatos.Range("A3:Z" & lastRowDatos).ClearContents
        End If
    End If

    ' Crear o actualizar tabla Inversiones_Datos
    On Error Resume Next
    Set loDatos = wsDatos.ListObjects("Inversiones_Datos")
    On Error GoTo 0

    If loDatos Is Nothing Then
        wsDatos.Range("A3").Resize(loInv.Range.Rows.Count, loInv.Range.Columns.Count).Value = loInv.Range.Value
        Set loDatos = wsDatos.ListObjects.Add( _
                        SourceType:=xlSrcRange, _
                        Source:=wsDatos.Range("A3").Resize(loInv.Range.Rows.Count, loInv.Range.Columns.Count), _
                        XlListObjectHasHeaders:=xlYes)
        loDatos.Name = "Inversiones_Datos"
        On Error Resume Next
        loDatos.TableStyle = "TableStyleLight8"
        loDatos.TableStyle = "Claro 8"
        On Error GoTo 0
    Else
        Set hdr = loDatos.HeaderRowRange
        Set destRange = hdr.Resize(loInv.Range.Rows.Count, loInv.Range.Columns.Count)
        loDatos.Resize destRange
        destRange.Value = loInv.Range.Value
    End If

    If loDatos Is Nothing Then Exit Sub
    If loDatos.DataBodyRange Is Nothing Then Exit Sub

    ' Limpieza de espacios
    RecortarEspaciosEnRango loDatos.DataBodyRange

    ' Corregir fechas y aplicar formatos numéricos
    CorregirFechasTextoDDMM loDatos, "Fecha de Operacion"
    CorregirFechasTextoDDMM loDatos, "Fecha de Operación"
    CorregirFechasTextoDDMM loDatos, "Fecha Liquidacion"
    CorregirFechasTextoDDMM loDatos, "Fecha fin Contrato"

    FormatearTablaInversiones loDatos

    ' Forzar texto en columnas no numéricas ni de fecha en Inversiones_Datos
    ForzarTextoEnNoNumericas loDatos

    ' Fechas mínima y máxima de Fecha de Operacion
    colFecha = 0
    On Error Resume Next
    colFecha = loDatos.ListColumns("Fecha de Operacion").Index
    If colFecha = 0 Then
        colFecha = loDatos.ListColumns("Fecha de Operación").Index
    End If
    On Error GoTo 0

    If colFecha > 0 Then
        Set dataRange = loDatos.DataBodyRange.Columns(colFecha)
        On Error Resume Next
        minDate = Application.WorksheetFunction.Min(dataRange)
        maxDate = Application.WorksheetFunction.Max(dataRange)
        On Error GoTo 0

        If Not IsError(minDate) And Not IsError(maxDate) Then
            wsDatos.Range("B2").Value = minDate
            wsDatos.Range("D2").Value = maxDate
            wsDatos.Range("B2").NumberFormatLocal = "dd/mm/aaaa"
            wsDatos.Range("D2").NumberFormatLocal = "dd/mm/aaaa"
        End If
    End If

    ' ==== FORMATO DE LA FILA DE ENCABEZADOS (FILA 3) IGUAL AL ORIGEN ====
    If Not wsOrigen Is Nothing Then
        Dim nColsHeader As Long
        nColsHeader = loDatos.HeaderRowRange.Columns.Count

        wsOrigen.Range(wsOrigen.Cells(3, 1), wsOrigen.Cells(3, nColsHeader)).Copy
        loDatos.HeaderRowRange.Cells(1, 1).PasteSpecial Paste:=xlPasteFormats
        Application.CutCopyMode = False
    End If

    ' Copiar formato de los registros (A4 hasta la última celda usada del origen)
    If Not wsOrigen Is Nothing Then
        On Error Resume Next
        Set lastCellOrigen = wsOrigen.Cells.SpecialCells(xlCellTypeLastCell)
        On Error GoTo 0

        If Not lastCellOrigen Is Nothing Then
            Set rngFmtSrc = wsOrigen.Range(wsOrigen.Range("A4"), lastCellOrigen)
            Set rngFmtDest = wsDatos.Range("A4").Resize(rngFmtSrc.Rows.Count, rngFmtSrc.Columns.Count)

            rngFmtSrc.Copy
            rngFmtDest.PasteSpecial Paste:=xlPasteFormats
            Application.CutCopyMode = False
        End If
    End If

    ' Logo al final
    If Not wsOrigen Is Nothing Then
        CopiarImagenA1 wsOrigen, wsDatos
    End If

    ' Copiar SOLO FORMATO A1:W3 del origen
    If Not wsOrigen Is Nothing Then
        wsOrigen.Range("A1:W3").Copy
        wsDatos.Range("A1").PasteSpecial Paste:=xlPasteFormats
        Application.CutCopyMode = False
    End If

    ' Formato de fecha para las columnas de fecha en la tabla Inversiones_Datos
    FormatearColumna loDatos, "Fecha de Operacion", "dd/mm/yyyy"
    FormatearColumna loDatos, "Fecha Liquidacion", "dd/mm/yyyy"
    FormatearColumna loDatos, "Fecha fin Contrato", "dd/mm/yyyy"

End Sub


' =========================================================
'  Extras para Acciones ETF
' =========================================================

Private Sub ConfigurarExtrasAccionesETF()

    Dim wb As Workbook
    Dim wsAcc As Worksheet
    Dim lo As ListObject
    Dim colMoneda As ListColumn
    Dim colCierre As ListColumn
    Dim colApertura As ListColumn
    Dim colAlto As ListColumn
    Dim colBajo As ListColumn
    Dim colResultado As ListColumn
    Dim colValidacion As ListColumn
    Dim fBase As String
    Dim fCierre As String
    Dim fApertura As String
    Dim fAlto As String
    Dim fBajo As String
    Dim rngRes As Range
    Dim fcNoInfo As FormatCondition
    Dim fcDentro As FormatCondition
    Dim fcFuera As FormatCondition
    Dim fcSinDatos As FormatCondition

    Set wb = ThisWorkbook

    On Error Resume Next
    Set wsAcc = wb.Worksheets("Acciones ETF")
    On Error GoTo 0
    If wsAcc Is Nothing Then Exit Sub

    On Error Resume Next
    Set lo = wsAcc.ListObjects(1)
    On Error GoTo 0
    If lo Is Nothing Then Exit Sub

    On Error Resume Next
    lo.TableStyle = "TableStyleMedium7"
    lo.TableStyle = "Medio 7"
    On Error GoTo 0

    FormatearTablaAccionesETF lo

    On Error Resume Next
    Set colMoneda = lo.ListColumns("Moneda")
    On Error GoTo 0
    If colMoneda Is Nothing Then Exit Sub

    On Error Resume Next
    Set colCierre = lo.ListColumns("Cierre")
    On Error GoTo 0
    If colCierre Is Nothing Then
        Set colCierre = lo.ListColumns.Add
        colCierre.Name = "Cierre"
    End If

    On Error Resume Next
    Set colApertura = lo.ListColumns("Apertura")
    On Error GoTo 0
    If colApertura Is Nothing Then
        Set colApertura = lo.ListColumns.Add
        colApertura.Name = "Apertura"
    End If

    On Error Resume Next
    Set colAlto = lo.ListColumns("Alto")
    On Error GoTo 0
    If colAlto Is Nothing Then
        Set colAlto = lo.ListColumns.Add
        colAlto.Name = "Alto"
    End If

    On Error Resume Next
    Set colBajo = lo.ListColumns("Bajo")
    On Error GoTo 0
    If colBajo Is Nothing Then
        Set colBajo = lo.ListColumns.Add
        colBajo.Name = "Bajo"
    End If

    On Error Resume Next
    Set colResultado = lo.ListColumns("Resultado")
    On Error GoTo 0
    If colResultado Is Nothing Then
        Set colResultado = lo.ListColumns.Add
        colResultado.Name = "Resultado"
    End If

    On Error Resume Next
    Set colValidacion = lo.ListColumns("Validación Manual")
    On Error GoTo 0
    If colValidacion Is Nothing Then
        Set colValidacion = lo.ListColumns.Add
        colValidacion.Name = "Validación Manual"
    End If

    If lo.DataBodyRange Is Nothing Then Exit Sub

    fBase = "=SI(SI.ERROR(ENCONTRAR("" LN"";ESPACIOS([@Nemonico]));0)>0;" & _
            "SI.ERROR(INDICE(HISTORIALCOTIZACIONES(""XLON:"" & SUSTITUIR(SUSTITUIR(ESPACIOS([@Nemonico]);"" LN"";"""");"".L"";"""");[@[Fecha de Operacion]];[@[Fecha de Operacion]];0;0;1;2;3;4);1;[IDX]);" & _
            "SI.ERROR(INDICE(HISTORIALCOTIZACIONES(""XRES:"" & SUSTITUIR(SUSTITUIR(ESPACIOS([@Nemonico]);"" LN"";"""");"".L"";"""");[@[Fecha de Operacion]];[@[Fecha de Operacion]];0;0;1;2;3;4);1;[IDX]);" & _
            "SI.ERROR(INDICE(HISTORIALCOTIZACIONES(SUSTITUIR(ESPACIOS([@Nemonico]);"" LN"";"".L"");[@[Fecha de Operacion]];[@[Fecha de Operacion]];0;0;1;2;3;4);1;[IDX]);" & _
            "INDICE(HISTORIALCOTIZACIONES([@Nemonico];[@[Fecha de Operacion]];[@[Fecha de Operacion]];0;0;1;2;3;4);1;[IDX]))));" & _
            "SI.ERROR(INDICE(HISTORIALCOTIZACIONES([@Nemonico];[@[Fecha de Operacion]];[@[Fecha de Operacion]];0;0;1;2;3;4);1;[IDX]);" & _
            "SI.ERROR(INDICE(HISTORIALCOTIZACIONES(""XLON:"" & SUSTITUIR([@Nemonico];"" LN"";"""");[@[Fecha de Operacion]];[@[Fecha de Operacion]];0;0;1;2;3;4);1;[IDX]);" & _
            "INDICE(HISTORIALCOTIZACIONES(""XNYS:"" & SUSTITUIR([@Nemonico];"" US"";"""");[@[Fecha de Operacion]];[@[Fecha de Operacion]];0;0;1;2;3;4);1;[IDX]))))"

    fCierre = Replace(fBase, "[IDX]", "1")
    fApertura = Replace(fBase, "[IDX]", "2")
    fAlto = Replace(fBase, "[IDX]", "3")
    fBajo = Replace(fBase, "[IDX]", "4")

    Set colCierre = lo.ListColumns("Cierre")
    Set colApertura = lo.ListColumns("Apertura")
    Set colAlto = lo.ListColumns("Alto")
    Set colBajo = lo.ListColumns("Bajo")

    If Not colCierre.DataBodyRange Is Nothing Then
        colCierre.DataBodyRange.FormulaLocal = fCierre
    End If

    If Not colApertura.DataBodyRange Is Nothing Then
        colApertura.DataBodyRange.FormulaLocal = fApertura
    End If

    If Not colAlto.DataBodyRange Is Nothing Then
        colAlto.DataBodyRange.FormulaLocal = fAlto
    End If

    If Not colBajo.DataBodyRange Is Nothing Then
        colBajo.DataBodyRange.FormulaLocal = fBajo
    End If

    If Not colResultado.DataBodyRange Is Nothing Then
        colResultado.DataBodyRange.FormulaLocal = _
            "=SI(ESNOD([@Cierre]);""No hay datos para esa fecha"";" & _
            "SI(ESERROR([@Cierre]);""No se pudo encontrar información"";" & _
            "SI(Y([@Precio]>=MIN([@Alto];[@Bajo]);[@Precio]<=MAX([@Alto];[@Bajo]));""Dentro del rango"";""Fuera del rango"")))"
    End If

    Set rngRes = Nothing
    On Error Resume Next
    Set rngRes = colResultado.DataBodyRange
    On Error GoTo 0

    If Not rngRes Is Nothing Then
        With rngRes
            .FormatConditions.Delete

            Set fcNoInfo = .FormatConditions.Add( _
                Type:=xlTextString, _
                String:="No se pudo encontrar información", _
                TextOperator:=xlContains)
            fcNoInfo.Interior.Color = RGB(202, 237, 251)
            fcNoInfo.Font.Color = RGB(54, 96, 146)

            Set fcDentro = .FormatConditions.Add( _
                Type:=xlTextString, _
                String:="Dentro del rango", _
                TextOperator:=xlContains)
            fcDentro.Interior.Color = RGB(198, 239, 206)
            fcDentro.Font.Color = RGB(0, 97, 0)

            Set fcFuera = .FormatConditions.Add( _
                Type:=xlTextString, _
                String:="Fuera del rango", _
                TextOperator:=xlContains)
            fcFuera.Interior.Color = RGB(255, 199, 206)
            fcFuera.Font.Color = RGB(156, 0, 6)

            Set fcSinDatos = .FormatConditions.Add( _
                Type:=xlTextString, _
                String:="No hay datos para esa fecha", _
                TextOperator:=xlContains)
            fcSinDatos.Interior.Color = RGB(255, 235, 156)
            fcSinDatos.Font.Color = RGB(156, 87, 0)
        End With
    End If

End Sub

' AutoFit para todas las hojas excepto Inicio.
Private Sub AutoFitTablasEnLibro(ByVal wb As Workbook)

    Dim ws As Worksheet

    On Error Resume Next
    For Each ws In wb.Worksheets
        If ws.Name <> "Inicio" Then
            ws.Cells.EntireColumn.AutoFit
        End If
    Next ws
    On Error GoTo 0

End Sub




Private Sub OrdenarHojasPrincipal(ByVal wb As Workbook)

    On Error Resume Next
    With wb
        .Worksheets("Inicio").Move Before:=.Worksheets(1)
        .Worksheets("ListaBancos").Move After:=.Worksheets("Inicio")
        .Worksheets("Datos").Move After:=.Worksheets("ListaBancos")
        .Worksheets("Acciones ETF").Move After:=.Worksheets("Datos")
        .Worksheets("Origen_Inversiones").Move After:=.Worksheets(.Worksheets.Count)
    End With
    On Error GoTo 0

End Sub



Private Sub LimpiarConsultasYHojas(ByVal wb As Workbook)

    Dim ws As Worksheet
    Dim qt As QueryTable
    Dim lo As ListObject

    On Error Resume Next

    ' Limpiar hoja Origen_Inversiones (se va a volver a llenar)
    Set ws = Nothing
    Set ws = wb.Worksheets("Origen_Inversiones")
    If Not ws Is Nothing Then
        ws.Cells.Clear
        For Each lo In ws.ListObjects
            lo.Unlist
        Next lo
        For Each qt In ws.QueryTables
            qt.Delete
        Next qt
    End If

    ' Limpiar hoja Inversiones (resultado de PQ)
    Set ws = Nothing
    Set ws = wb.Worksheets("Inversiones")
    If Not ws Is Nothing Then
        ws.Cells.Clear
        For Each lo In ws.ListObjects
            lo.Unlist
        Next lo
        For Each qt In ws.QueryTables
            qt.Delete
        Next qt
    End If

    ' Limpiar hoja Acciones ETF (resultado de PQ)
    Set ws = Nothing
    Set ws = wb.Worksheets("Acciones ETF")
    If Not ws Is Nothing Then
        ws.Cells.Clear
        For Each lo In ws.ListObjects
            lo.Unlist
        Next lo
        For Each qt In ws.QueryTables
            qt.Delete
        Next qt
    End If

    ' Limpiar hoja Datos (vista armada por VBA)
    Set ws = Nothing
    Set ws = wb.Worksheets("Datos")
    If Not ws Is Nothing Then
        ws.Cells.Clear
        For Each lo In ws.ListObjects
            lo.Unlist
        Next lo
        For Each qt In ws.QueryTables
            qt.Delete
        Next qt
    End If

    On Error GoTo 0

End Sub