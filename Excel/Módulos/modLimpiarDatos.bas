Option Explicit

Public Sub LimpiarDatos()
    Const KEEP_1 As String = "Inicio"
    Const KEEP_2 As String = "ListaBancos"
    
    Dim wb As Workbook
    Dim respuesta As VbMsgBoxResult
    Dim calcMode As XlCalculation
    
    Set wb = ThisWorkbook
    
    If Not SheetExists(wb, KEEP_1) Or Not SheetExists(wb, KEEP_2) Then
        MsgBox "No se puede ejecutar la limpieza porque deben existir las hojas 'Inicio' y 'ListaBancos'.", vbExclamation, "Limpiar datos"
        Exit Sub
    End If
    
    respuesta = MsgBox( _
        "Se eliminarán todas las hojas, consultas y conexiones del libro, excepto 'Inicio' y 'ListaBancos'." & vbCrLf & vbCrLf & _
        "Esta acción no se puede deshacer." & vbCrLf & vbCrLf & _
        "¿Desea continuar?", _
        vbYesNo + vbQuestion + vbDefaultButton2, _
        "Confirmar limpieza")
    
    If respuesta <> vbYes Then Exit Sub
    
    On Error GoTo ErrHandler
    
    calcMode = Application.Calculation
    
    With Application
        .ScreenUpdating = False
        .EnableEvents = False
        .DisplayAlerts = False
        .Calculation = xlCalculationManual
        .StatusBar = "Limpiando datos del libro..."
    End With
    
    DeleteNonBaseSheets wb, KEEP_1, KEEP_2
    DeleteWorkbookQueries wb
    DeleteWorkbookConnections wb
    DeleteBrokenNames wb
    
    With Application
        .StatusBar = False
        .Calculation = calcMode
        .DisplayAlerts = True
        .EnableEvents = True
        .ScreenUpdating = True
    End With
    
    MsgBox "Limpieza completada. Se conservaron únicamente las hojas 'Inicio' y 'ListaBancos'.", vbInformation, "Limpiar datos"
    Exit Sub

ErrHandler:
    With Application
        .StatusBar = False
        .Calculation = calcMode
        .DisplayAlerts = True
        .EnableEvents = True
        .ScreenUpdating = True
    End With
    
    MsgBox "No se pudo completar la limpieza." & vbCrLf & vbCrLf & _
           "Detalle: " & Err.Description, vbCritical, "Limpiar datos"
End Sub

Private Sub DeleteNonBaseSheets(ByVal wb As Workbook, ByVal sheetName1 As String, ByVal sheetName2 As String)
    Dim i As Long
    
    For i = wb.Worksheets.Count To 1 Step -1
        If Not IsBaseSheet(wb.Worksheets(i).Name, sheetName1, sheetName2) Then
            wb.Worksheets(i).Delete
        End If
    Next i
End Sub

Private Sub DeleteWorkbookQueries(ByVal wb As Workbook)
    Dim i As Long
    
    On Error Resume Next
    For i = wb.Queries.Count To 1 Step -1
        wb.Queries(i).Delete
    Next i
    On Error GoTo 0
End Sub

Private Sub DeleteWorkbookConnections(ByVal wb As Workbook)
    Dim i As Long
    
    On Error Resume Next
    For i = wb.Connections.Count To 1 Step -1
        wb.Connections(i).Delete
    Next i
    On Error GoTo 0
End Sub

Private Sub DeleteBrokenNames(ByVal wb As Workbook)
    Dim i As Long
    Dim refersToText As String
    
    On Error Resume Next
    For i = wb.Names.Count To 1 Step -1
        refersToText = vbNullString
        refersToText = wb.Names(i).RefersTo
        
        If InStr(1, refersToText, "#REF!", vbTextCompare) > 0 Then
            wb.Names(i).Delete
        End If
    Next i
    On Error GoTo 0
End Sub

Private Function SheetExists(ByVal wb As Workbook, ByVal sheetName As String) As Boolean
    Dim ws As Worksheet
    
    On Error Resume Next
    Set ws = wb.Worksheets(sheetName)
    SheetExists = Not ws Is Nothing
    Set ws = Nothing
    On Error GoTo 0
End Function

Private Function IsBaseSheet(ByVal currentName As String, ByVal sheetName1 As String, ByVal sheetName2 As String) As Boolean
    IsBaseSheet = (StrComp(currentName, sheetName1, vbTextCompare) = 0 Or _
                   StrComp(currentName, sheetName2, vbTextCompare) = 0)
End Function