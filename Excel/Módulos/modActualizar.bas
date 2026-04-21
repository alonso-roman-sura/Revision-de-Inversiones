Sub ActualizarLibroActivo()

    Dim wb As Workbook
    Dim calcMode As XlCalculation
    
    ' Verifica que haya un libro activo
    If ActiveWorkbook Is Nothing Then Exit Sub
    Set wb = ActiveWorkbook
    
    On Error GoTo Salir
    
    ' Guarda el modo de cálculo actual
    calcMode = Application.Calculation
    
    ' Mejora rendimiento durante la actualización
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    
    ' Equivalente al botón "Actualizar todo"
    wb.RefreshAll
    
    ' Espera a que terminen consultas asíncronas (si aplica)
    On Error Resume Next
    Application.CalculateUntilAsyncQueriesDone
    On Error GoTo Salir

Salir:
    ' Restaura configuración de la aplicación
    Application.Calculation = calcMode
    Application.EnableEvents = True
    Application.ScreenUpdating = True

End Sub