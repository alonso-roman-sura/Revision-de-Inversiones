Option Explicit

Sub ImportarTasasPasivasTipoPersona()

    Dim ruta As Variant
    Dim wbSrc As Workbook
    Dim wbDest As Workbook
    Dim loSrc As ListObject
    Dim wsOrigen As Worksheet
    Dim wsTasas As Worksheet
    Dim dataRange As Range
    Dim lastRow As Long
    Dim lastCol As Long
    Dim loRaw As ListObject
    Dim loDest As ListObject
    Dim ws As Worksheet
    Dim lo As ListObject
    Dim sh As Shape

    ruta = Application.GetOpenFilename( _
                "Archivos Excel (*.xlsx;*.xlsm;*.xlsb),*.xlsx;*.xlsm;*.xlsb", _
                title:="Selecciona el archivo de Tasas Pasivas por Tipo de Persona")
    If ruta = False Then Exit Sub

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    Set wbDest = ThisWorkbook
    Set wbSrc = Workbooks.Open(Filename:=ruta, ReadOnly:=True, UpdateLinks:=False, IgnoreReadOnlyRecommended:=True)

    ' Tabla origen en el archivo seleccionado
    Set loSrc = ObtenerTablaPorNombreEnLibro(wbSrc, "TasaPasivaTipoPersona")
    If loSrc Is Nothing Then
        MsgBox "No se encontró la tabla 'TasaPasivaTipoPersona' en el archivo seleccionado.", vbExclamation
        GoTo Salir
    End If
    Set dataRange = loSrc.Range

    ' Hoja de origen interna (raw)
    On Error Resume Next
    Set wsOrigen = wbDest.Worksheets("Origen_TasasPasivasTipoPersona")
    On Error GoTo 0

    If wsOrigen Is Nothing Then
        Set wsOrigen = wbDest.Worksheets.Add(After:=wbDest.Worksheets(wbDest.Worksheets.Count))
        wsOrigen.Name = "Origen_TasasPasivasTipoPersona"
    End If

    ' Limpiar origen
    wsOrigen.Cells.Clear
    On Error Resume Next
    For Each sh In wsOrigen.Shapes
        sh.Delete
    Next sh
    For Each lo In wsOrigen.ListObjects
        lo.Unlist
    Next lo
    On Error GoTo 0

    ' Volcar los datos del archivo a la hoja de origen
    wsOrigen.Range("A1").Resize(dataRange.Rows.Count, dataRange.Columns.Count).Value = dataRange.Value

    lastRow = dataRange.Rows.Count
    lastCol = dataRange.Columns.Count

    RecortarEspaciosEnRango wsOrigen.Range(wsOrigen.Cells(1, 1), wsOrigen.Cells(lastRow, lastCol))

    ' Borrar cualquier tabla previa con ese nombre y recrearla como RAW
    EliminarTablaPorNombreEnLibro wbDest, "TasaPasivaTipoPersona_Raw"

    Set loRaw = wsOrigen.ListObjects.Add( _
                    SourceType:=xlSrcRange, _
                    Source:=wsOrigen.Range("A1").Resize(lastRow, lastCol), _
                    XlListObjectHasHeaders:=xlYes)
    loRaw.Name = "TasaPasivaTipoPersona_Raw"

    ' Crear / actualizar las queries que usan la tabla RAW
    CrearOActualizarPQTasasPasivasTipoPersona
    CrearOActualizarPQConstitucionTitulosUnicos

    ' Cargar query de tasas en hoja y tabla conectada
    AsegurarTablaDeConsulta_Tasas "Tasas Pasivas Tipo Persona", _
                                  "Tasas Pasivas Tipo Persona", _
                                  "Tasas_Pasivas_Tipo_Persona"

    ' Cargar query de Constitución Títulos Únicos
    AsegurarTablaDeConsulta_Tasas "Constitución Títulos Únicos", _
                                  "Constitución Títulos Únicos", _
                                  "Constitucion_Titulos_Unicos"

    ' Estilo de la tabla de tasas
    On Error Resume Next
    Set wsTasas = wbDest.Worksheets("Tasas Pasivas Tipo Persona")
    If Not wsTasas Is Nothing Then
        Set loDest = wsTasas.ListObjects("Tasas_Pasivas_Tipo_Persona")
        If Not loDest Is Nothing Then
            loDest.TableStyle = "TableStyleMedium7"
            loDest.TableStyle = "Medio 7"
        End If
    End If
    On Error GoTo 0

    ' Configura hoja Constitución Títulos Únicos (usa la tabla conectada)
    ConfigurarHojaConstitucionTitulosUnicos

    MsgBox "Proceso completado: Tasas pasivas e inversión en Constitución Títulos Únicos actualizados.", vbInformation

Salir:
    On Error Resume Next
    If Not wbSrc Is Nothing Then wbSrc.Close SaveChanges:=False
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.CutCopyMode = False

End Sub


Private Function ObtenerTablaPorNombreEnLibro(ByVal wb As Workbook, ByVal nombreTabla As String) As ListObject

    Dim ws As Worksheet
    Dim lo As ListObject

    If wb Is Nothing Then Exit Function

    For Each ws In wb.Worksheets
        For Each lo In ws.ListObjects
            If StrComp(lo.Name, nombreTabla, vbTextCompare) = 0 Then
                Set ObtenerTablaPorNombreEnLibro = lo
                Exit Function
            End If
        Next lo
    Next ws

End Function

Private Sub CrearOActualizarPQTasasPasivasTipoPersona()

    Dim mCode As String
    Dim q As WorkbookQuery

    mCode = "let" & vbCrLf
    mCode = mCode & "    Origen = Excel.CurrentWorkbook(){[Name=""TasaPasivaTipoPersona_Raw""]}[Content]," & vbCrLf

    ' Normaliza nombres por si vienen sin tildes
    mCode = mCode & "    #""Columnas normalizadas"" = Table.RenameColumns(Origen," & vbCrLf
    mCode = mCode & "        { " & vbCrLf
    mCode = mCode & "            {""Hasta 30 dias"", ""Hasta 30 días""}," & vbCrLf
    mCode = mCode & "            {""31-90 dias"", ""31-90 días""}," & vbCrLf
    mCode = mCode & "            {""91-180 dias"", ""91-180 días""}," & vbCrLf
    mCode = mCode & "            {""181-360 dias"", ""181-360 días""}," & vbCrLf
    mCode = mCode & "            {""Mas de 360 dias"", ""Más de 360 días""}" & vbCrLf
    mCode = mCode & "        }, MissingField.Ignore)," & vbCrLf

    mCode = mCode & "    #""Tipo cambiado"" = Table.TransformColumnTypes(" & vbCrLf
    mCode = mCode & "        #""Columnas normalizadas""," & vbCrLf
    mCode = mCode & "        { " & vbCrLf
    mCode = mCode & "            {""Hasta 30 días"", type number}," & vbCrLf
    mCode = mCode & "            {""31-90 días"", type number}," & vbCrLf
    mCode = mCode & "            {""91-180 días"", type number}," & vbCrLf
    mCode = mCode & "            {""181-360 días"", type number}," & vbCrLf
    mCode = mCode & "            {""Más de 360 días"", type number}" & vbCrLf
    mCode = mCode & "        }" & vbCrLf
    mCode = mCode & "    )" & vbCrLf
    mCode = mCode & "in" & vbCrLf
    mCode = mCode & "    #""Tipo cambiado"""

    On Error Resume Next
    Set q = ThisWorkbook.Queries("Tasas Pasivas Tipo Persona")
    On Error GoTo 0

    If q Is Nothing Then
        ThisWorkbook.Queries.Add Name:="Tasas Pasivas Tipo Persona", Formula:=mCode
    Else
        q.Formula = mCode
    End If

End Sub

Private Sub CrearOActualizarPQConstitucionTitulosUnicos()

    Dim mCode As String
    Dim q As WorkbookQuery

    mCode = "let" & vbCrLf
    mCode = mCode & "    Origen = Excel.CurrentWorkbook(){[Name=""Inversiones""]}[Content]," & vbCrLf
    mCode = mCode & "    #""Tipo cambiado"" = Table.TransformColumnTypes(" & vbCrLf
    mCode = mCode & "        Origen," & vbCrLf
    mCode = mCode & "        {" & vbCrLf
    mCode = mCode & "            {""Portafolio"", type text}," & vbCrLf
    mCode = mCode & "            {""Codigo de Orden"", type text}," & vbCrLf
    mCode = mCode & "            {""Fecha de Operacion"", type date}," & vbCrLf
    mCode = mCode & "            {""Fecha Liquidacion"", type date}," & vbCrLf
    mCode = mCode & "            {""Fecha fin Contrato"", type date}," & vbCrLf
    mCode = mCode & "            {""Codigo ISIN"", type text}," & vbCrLf
    mCode = mCode & "            {""Codigo SBS"", type text}," & vbCrLf
    mCode = mCode & "            {""Monto de Operacion Original"", type number}," & vbCrLf
    mCode = mCode & "            {""Monto de Operacion ML"", type number}," & vbCrLf
    mCode = mCode & "            {""Cantidad"", type text}," & vbCrLf
    mCode = mCode & "            {""Precio"", type text}," & vbCrLf
    mCode = mCode & "            {""Codigo de Emisor"", type text}," & vbCrLf
    mCode = mCode & "            {""Operacion"", type text}," & vbCrLf
    mCode = mCode & "            {""Moneda"", type text}," & vbCrLf
    mCode = mCode & "            {""Nemonico"", type text}," & vbCrLf
    mCode = mCode & "            {""Codigo de Tercero"", type text}," & vbCrLf
    mCode = mCode & "            {""Tercero"", type text}," & vbCrLf
    mCode = mCode & "            {""Monto Nominal Operacion Original"", type number}," & vbCrLf
    mCode = mCode & "            {""Monto Nominal Operacion ML"", type number}," & vbCrLf
    mCode = mCode & "            {""Total de Comisiones"", type text}," & vbCrLf
    mCode = mCode & "            {""Plaza"", type text}," & vbCrLf
    mCode = mCode & "            {""Tipo Tasa"", type text}," & vbCrLf
    mCode = mCode & "            {""Porcentaje Tasa"", type number}" & vbCrLf
    mCode = mCode & "        }" & vbCrLf
    mCode = mCode & "    )," & vbCrLf
    mCode = mCode & "    #""Filtrado Operacion"" = Table.SelectRows(" & vbCrLf
    mCode = mCode & "        #""Tipo cambiado""," & vbCrLf
    mCode = mCode & "        each [Operacion] = ""CONSTITUCION TITULOS UNICOS""" & vbCrLf
    mCode = mCode & "    )," & vbCrLf
    mCode = mCode & "    #""Columnas quitadas"" = Table.RemoveColumns(" & vbCrLf
    mCode = mCode & "        #""Filtrado Operacion""," & vbCrLf
    mCode = mCode & "        {""Codigo ISIN"", ""Cantidad"", ""Precio"", ""Operacion"", ""Total de Comisiones""}" & vbCrLf
    mCode = mCode & "    )," & vbCrLf
    mCode = mCode & "    #""Ordenadas"" = Table.Sort(" & vbCrLf
    mCode = mCode & "        #""Columnas quitadas""," & vbCrLf
    mCode = mCode & "        {{""Fecha de Operacion"", Order.Ascending}, {""Codigo de Orden"", Order.Ascending}}" & vbCrLf
    mCode = mCode & "    )" & vbCrLf
    mCode = mCode & "in" & vbCrLf
    mCode = mCode & "    #""Ordenadas"""

    On Error Resume Next
    Set q = ThisWorkbook.Queries("Constitución Títulos Únicos")
    On Error GoTo 0

    If q Is Nothing Then
        ThisWorkbook.Queries.Add Name:="Constitución Títulos Únicos", Formula:=mCode
    Else
        q.Formula = mCode
    End If

End Sub

Private Sub ActualizarConexionPorNombre(ByVal nombre As String)

    Dim conn As WorkbookConnection

    On Error Resume Next
    For Each conn In ThisWorkbook.Connections
        If StrComp(conn.Name, nombre, vbTextCompare) = 0 Or InStr(1, conn.Name, nombre, vbTextCompare) > 0 Then
            conn.Refresh
        End If
    Next conn
    On Error GoTo 0

End Sub

Private Sub ConfigurarHojaConstitucionTitulosUnicos()

    Dim ws As Worksheet
    Dim lo As ListObject
    Dim colDurDias As ListColumn
    Dim colDurInt As ListColumn
    Dim colTasa As ListColumn
    Dim colRes As ListColumn
    Dim rngCF As Range
    Dim firstCell As Range
    Dim addr As String
    Dim idxFechaFin As Long
    Dim idxFechaOp As Long
    Dim offsetFin As Long
    Dim offsetOp As Long

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("Constitución Títulos Únicos")
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub
    If ws.ListObjects.Count = 0 Then Exit Sub

    Set lo = ws.ListObjects(1)

    On Error Resume Next
    lo.TableStyle = "TableStyleMedium5"
    lo.TableStyle = "Medio 5"
    On Error GoTo 0

    ' Normalizar nombres de columnas creadas (sin tildes)
    On Error Resume Next
    lo.ListColumns("Duración Días").Name = "Duracion Dias"
    lo.ListColumns("Duración Intervalo").Name = "Duracion Intervalo"
    On Error GoTo 0

    ' Columna Duracion Dias
    On Error Resume Next
    Set colDurDias = lo.ListColumns("Duracion Dias")
    On Error GoTo 0
    If colDurDias Is Nothing Then
        Set colDurDias = lo.ListColumns.Add
        colDurDias.Name = "Duracion Dias"
    End If

    ' Columna Duracion Intervalo
    On Error Resume Next
    Set colDurInt = lo.ListColumns("Duracion Intervalo")
    On Error GoTo 0
    If colDurInt Is Nothing Then
        Set colDurInt = lo.ListColumns.Add
        colDurInt.Name = "Duracion Intervalo"
    End If

    ' Índices de fechas
    On Error Resume Next
    idxFechaFin = lo.ListColumns("Fecha fin Contrato").Index
    idxFechaOp = lo.ListColumns("Fecha de Operacion").Index
    If idxFechaOp = 0 Then
        idxFechaOp = lo.ListColumns("Fecha de Operación").Index
    End If
    On Error GoTo 0

    ' Duracion Dias = Fecha fin Contrato - Fecha de Operacion
    If Not colDurDias.DataBodyRange Is Nothing Then
        If idxFechaFin > 0 And idxFechaOp > 0 Then
            offsetFin = idxFechaFin - colDurDias.Index
            offsetOp = idxFechaOp - colDurDias.Index
            colDurDias.DataBodyRange.FormulaR1C1 = "=RC[" & offsetFin & "]-RC[" & offsetOp & "]"
        End If
    End If

    ' Duracion Intervalo
    If Not colDurInt.DataBodyRange Is Nothing Then
        colDurInt.DataBodyRange.FormulaLocal = _
            "=SI.CONJUNTO(" & _
            "[@[Duracion Dias]]<=30;""Hasta 30 días"";" & _
            "[@[Duracion Dias]]<=90;""31-90 días"";" & _
            "[@[Duracion Dias]]<=180;""91-180 días"";" & _
            "[@[Duracion Dias]]<=360;""181-360 días"";" & _
            "[@[Duracion Dias]]>360;""Más de 360 días"")"
    End If

    ' Columna Tasa SBS
    On Error Resume Next
    Set colTasa = lo.ListColumns("Tasa SBS")
    On Error GoTo 0
    If colTasa Is Nothing Then
        Set colTasa = lo.ListColumns.Add
        colTasa.Name = "Tasa SBS"
    End If

    If Not colTasa.DataBodyRange Is Nothing Then
        colTasa.DataBodyRange.FormulaLocal = _
            "=TasaSBS_Valor([@Tercero];[@Moneda];[@[Duracion Intervalo]];[@[Fecha de Operacion]])"
        colTasa.DataBodyRange.NumberFormat = "0.00"
    End If

    ' Columna Resultado
    On Error Resume Next
    Set colRes = lo.ListColumns("Resultado")
    On Error GoTo 0
    If colRes Is Nothing Then
        Set colRes = lo.ListColumns.Add
        colRes.Name = "Resultado"
    End If

    If Not colRes.DataBodyRange Is Nothing Then
        colRes.DataBodyRange.FormulaLocal = _
            "=ResultadoTasaSBS([@[Tasa SBS]];[@[Porcentaje Tasa]])"

        Set rngCF = colRes.DataBodyRange
        Set firstCell = rngCF.Cells(1, 1)
        addr = firstCell.Address(RowAbsolute:=False, ColumnAbsolute:=True)

        rngCF.FormatConditions.Delete

        ' El banco no es peruano: texto azul fuerte, relleno azul claro
        With rngCF.FormatConditions.Add(Type:=xlExpression, Formula1:="=" & addr & "=""El banco no es peruano""")
            .Font.Color = RGB(54, 96, 146)
            .Interior.Color = RGB(202, 237, 251)
        End With

        ' No encontrado
        With rngCF.FormatConditions.Add(Type:=xlExpression, Formula1:="=" & addr & "=""No encontrado""")
            .Font.Color = RGB(156, 87, 0)
            .Interior.Color = RGB(255, 235, 156)
        End With

        ' Dentro del rango
        With rngCF.FormatConditions.Add(Type:=xlExpression, Formula1:="=" & addr & "=""Dentro del rango""")
            .Font.Color = RGB(0, 97, 0)
            .Interior.Color = RGB(198, 239, 206)
        End With

        ' Fuera del rango
        With rngCF.FormatConditions.Add(Type:=xlExpression, Formula1:="=" & addr & "=""Fuera del rango""")
            .Font.Color = RGB(156, 0, 6)
            .Interior.Color = RGB(255, 199, 206)
        End With

        ' Autoajuste de columnas y posición final en R1
        ws.Columns.AutoFit
        ws.Activate
        ws.Range("R1").Select
    End If

End Sub

Private Function BancoNormalizado(ByVal Tercero As String) As String

    Dim tert As String
    tert = UCase$(Tercero)

    If InStr(tert, "BBVA") > 0 Then BancoNormalizado = "BBVA": Exit Function
    If InStr(tert, "BCI") > 0 Then BancoNormalizado = "BCI": Exit Function
    If InStr(tert, "CREDITO") > 0 Then BancoNormalizado = "Crédito": Exit Function
    If InStr(tert, "FALABELLA") > 0 Then BancoNormalizado = "Falabella": Exit Function
    If InStr(tert, "GNB") > 0 Then BancoNormalizado = "GNB": Exit Function
    If InStr(tert, "INTERAMERICANO") > 0 Then BancoNormalizado = "BIF": Exit Function
    If InStr(tert, "INTERBANK") > 0 Then BancoNormalizado = "Interbank": Exit Function
    If InStr(tert, "SANTANDER CONSUMER") > 0 Then BancoNormalizado = "Santander Cons. Bank": Exit Function
    If InStr(tert, "SANTANDER") > 0 Then BancoNormalizado = "Santander": Exit Function
    If InStr(tert, "ICBC") > 0 Then BancoNormalizado = "ICBC": Exit Function
    If InStr(tert, "MIBANCO") > 0 Then BancoNormalizado = "Mibanco": Exit Function
    If InStr(tert, "SCOTIA") > 0 Then BancoNormalizado = "Scotiabank": Exit Function
    If InStr(tert, "COMPARTAMOS") > 0 Then BancoNormalizado = "Compartamos": Exit Function

End Function

Public Function TasaSBS_Valor( _
        ByVal Tercero As String, _
        ByVal Moneda As String, _
        ByVal Duracion As Variant, _
        ByVal FechaOperacion As Variant) As Variant

    Dim bancoNorm As String
    Dim monNorm As String
    Dim ws As Worksheet
    Dim lo As ListObject
    Dim idxPersona As Long
    Dim idxMoneda As Long
    Dim idxBanco As Long
    Dim idxFecha As Long
    Dim idxHasta30 As Long
    Dim idx31a90 As Long
    Dim idx91a180 As Long
    Dim idx181a360 As Long
    Dim idxMas360 As Long
    Dim rngData As Range
    Dim fila As Range
    Dim filaMatch As Range
    Dim fechaCel As Variant
    Dim idxDur As Long

    bancoNorm = BancoNormalizado(Tercero)
    If bancoNorm = "" Then
        TasaSBS_Valor = "El banco no es peruano"
        Exit Function
    End If

    If Moneda = "NSOL" Then
        monNorm = "Moneda Nacional"
    ElseIf Moneda = "DOL" Then
        monNorm = "Moneda Extranjera"
    Else
        monNorm = Moneda
    End If

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("Tasas Pasivas Tipo Persona")
    On Error GoTo 0
    If ws Is Nothing Then
        TasaSBS_Valor = "No encontrado"
        Exit Function
    End If

    On Error Resume Next
    Set lo = ws.ListObjects("Tasas_Pasivas_Tipo_Persona")
    On Error GoTo 0
    If lo Is Nothing Then
        TasaSBS_Valor = "No encontrado"
        Exit Function
    End If

    On Error Resume Next
    idxPersona = lo.ListColumns("Tipo Persona").Index
    idxMoneda = lo.ListColumns("Tipo de Moneda").Index
    idxBanco = lo.ListColumns("Banco").Index
    idxFecha = lo.ListColumns("Fecha").Index
    idxHasta30 = lo.ListColumns("Hasta 30 días").Index
    idx31a90 = lo.ListColumns("31-90 días").Index
    idx91a180 = lo.ListColumns("91-180 días").Index
    idx181a360 = lo.ListColumns("181-360 días").Index
    idxMas360 = lo.ListColumns("Más de 360 días").Index
    On Error GoTo 0

    If idxPersona = 0 Or idxMoneda = 0 Or idxBanco = 0 Or idxFecha = 0 Then
        TasaSBS_Valor = "No encontrado"
        Exit Function
    End If

    Set rngData = lo.DataBodyRange
    If rngData Is Nothing Then
        TasaSBS_Valor = "No encontrado"
        Exit Function
    End If

    For Each fila In rngData.Rows
        If fila.Cells(idxPersona).Value = "Jurídica" _
           And fila.Cells(idxMoneda).Value = monNorm _
           And fila.Cells(idxBanco).Value = bancoNorm Then

            fechaCel = fila.Cells(idxFecha).Value
            If IsDate(fechaCel) And IsDate(FechaOperacion) Then
                If CDate(fechaCel) = CDate(FechaOperacion) Then
                    Set filaMatch = fila
                    Exit For
                End If
            End If
        End If
    Next fila

    If filaMatch Is Nothing Then
        TasaSBS_Valor = "No encontrado"
        Exit Function
    End If

    Select Case CStr(Duracion)
        Case "Hasta 30 días"
            idxDur = idxHasta30
        Case "31-90 días"
            idxDur = idx31a90
        Case "91-180 días"
            idxDur = idx91a180
        Case "181-360 días"
            idxDur = idx181a360
        Case "Más de 360 días"
            idxDur = idxMas360
        Case Else
            idxDur = 0
    End Select

    If idxDur = 0 Then
        TasaSBS_Valor = "No encontrado"
    Else
        TasaSBS_Valor = filaMatch.Cells(idxDur).Value
    End If

End Function

Public Function ResultadoTasaSBS( _
        ByVal tasa As Variant, _
        ByVal valor As Variant) As Variant

    If VarType(tasa) = vbString Then
        If tasa = "El banco no es peruano" Then
            ResultadoTasaSBS = "El banco no es peruano"
            Exit Function
        ElseIf tasa = "No encontrado" Then
            ResultadoTasaSBS = "No encontrado"
            Exit Function
        End If
    End If

    If IsError(tasa) Or IsError(valor) Then
        ResultadoTasaSBS = ""
        Exit Function
    End If

    If Not IsNumeric(tasa) Or Not IsNumeric(valor) Then
        ResultadoTasaSBS = ""
        Exit Function
    End If

    If CDbl(valor) >= CDbl(tasa) - 1 And CDbl(valor) <= CDbl(tasa) + 1 Then
        ResultadoTasaSBS = "Dentro del rango"
    Else
        ResultadoTasaSBS = "Fuera del rango"
    End If

End Function

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

Private Sub AsegurarTablaDeConsulta_Tasas( _
        ByVal queryName As String, _
        ByVal sheetName As String, _
        ByVal tableName As String)

    Dim wb As Workbook
    Dim ws As Worksheet
    Dim lo As ListObject
    Dim qt As QueryTable
    Dim connString As String
    Dim lc As ListColumn

    Set wb = ThisWorkbook

    On Error Resume Next
    Set ws = wb.Worksheets(sheetName)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        ws.Name = sheetName
    End If

    On Error Resume Next
    Set lo = ws.ListObjects(tableName)
    On Error GoTo 0

    connString = "OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=" & queryName & ";Extended Properties="""";"

    If lo Is Nothing Then
        Set lo = ws.ListObjects.Add(SourceType:=0, Source:=connString, Destination:=ws.Range("A1"))
        lo.Name = tableName
        Set qt = lo.QueryTable
    Else
        If lo.SourceType = xlSrcExternal Then
            Set qt = lo.QueryTable
            qt.Connection = connString
        Else
            lo.Unlist
            ws.Cells.Clear
            Set lo = ws.ListObjects.Add(SourceType:=0, Source:=connString, Destination:=ws.Range("A1"))
            lo.Name = tableName
            Set qt = lo.QueryTable
        End If
    End If

    If Not qt Is Nothing Then
        qt.CommandType = xlCmdSql
        qt.CommandText = Array("SELECT * FROM [" & queryName & "]")
        qt.Refresh BackgroundQuery:=False
    End If

    ' Formato de fecha para todas las columnas cuyo nombre contenga "Fecha"
    On Error Resume Next
    Set lo = ws.ListObjects(tableName)
    On Error GoTo 0
    If Not lo Is Nothing Then
        For Each lc In lo.ListColumns
            If InStr(1, lc.Name, "Fecha", vbTextCompare) > 0 Then
                If Not lc.DataBodyRange Is Nothing Then
                    lc.DataBodyRange.NumberFormat = "dd/mm/yyyy"
                End If
            End If
        Next lc
    End If

End Sub


Public Sub EliminarTablaPorNombreEnLibro( _
    ByVal wb As Workbook, _
    ByVal nombreTabla As String, _
    Optional ByVal limpiarRango As Boolean = False)

    Dim ws As Worksheet
    Dim lo As ListObject
    Dim rngTabla As Range

    If wb Is Nothing Then Exit Sub
    If Len(nombreTabla) = 0 Then Exit Sub

    For Each ws In wb.Worksheets
        For Each lo In ws.ListObjects
            If StrComp(lo.Name, nombreTabla, vbTextCompare) = 0 Then

                Set rngTabla = lo.Range

                ' Si es una tabla externa, intenta eliminar su QueryTable sin romper
                On Error Resume Next
                If lo.SourceType = xlSrcExternal Then
                    If Not lo.QueryTable Is Nothing Then
                        lo.QueryTable.CancelRefresh
                        lo.QueryTable.Delete
                    End If
                End If
                On Error GoTo 0

                ' Quitar la tabla (mantiene valores)
                lo.Unlist

                ' Opcional: limpiar el rango donde estaba la tabla
                If limpiarRango Then
                    rngTabla.Clear
                End If

                Exit Sub
            End If
        Next lo
    Next ws

End Sub