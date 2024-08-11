Sub OptimizarRendimiento()
    ' Declarar variables
    Dim ws As Worksheet
    Dim startTime As Double
    Dim endTime As Double
    
    ' Registrar tiempo de inicio
    startTime = Timer
    
    ' Optimizaciones
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    Application.DisplayStatusBar = False
    ActiveSheet.DisplayPageBreaks = False
    
    ' Si estás trabajando con una hoja específica, usa:
    ' Set ws = ThisWorkbook.Worksheets("NombreHoja")
    
    ' Tu código principal aquí
    ' Por ejemplo:
    ' For Each cell In ws.UsedRange
    '     ' Procesar datos
    ' Next cell
    
    ' Restaurar configuraciones
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.DisplayStatusBar = True
    ActiveSheet.DisplayPageBreaks = True
    
    ' Forzar una actualización final
    Application.ScreenUpdating = True
    
    ' Registrar tiempo de finalización y mostrar duración
    endTime = Timer
    MsgBox "Macro completada en " & Format(endTime - startTime, "0.00") & " segundos."
End Sub
