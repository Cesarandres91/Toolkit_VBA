Sub RellenarVacios()
    Dim rng As Range
    Dim cell As Range
    Dim lastRow As Long
    Dim lastValue As Variant
    Dim col As Range
    
    ' Optimizaciones para mejorar el rendimiento
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    ' Por defecto, trabajamos con la columna seleccionada
    Set rng = Selection
    
    ' Comprobar si se ha seleccionado algo
    If rng Is Nothing Then
        MsgBox "Por favor, selecciona una columna o rango.", vbExclamation
        Exit Sub
    End If
    
    ' Para trabajar con múltiples columnas o un rango personalizado, descomenta la siguiente línea:
    ' Set rng = Application.InputBox("Selecciona el rango a rellenar", "Selección de Rango", Type:=8)
    
    ' Si solo quieres trabajar con una columna específica, usa:
    ' Set rng = ActiveSheet.Range("A:A")  ' Cambia A:A por la columna deseada
    
    ' Iterar por cada columna en el rango seleccionado
    For Each col In rng.Columns
        lastRow = col.Cells(col.Rows.Count, 1).End(xlUp).Row
        lastValue = col.Cells(1, 1).Value
        
        ' Iterar por cada celda en la columna
        For Each cell In col.Cells(1, 1).Resize(lastRow)
            If IsEmpty(cell) Then
                cell.Value = lastValue
            Else
                lastValue = cell.Value
            End If
        Next cell
    Next col
    
    ' Restaurar configuraciones
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    
    MsgBox "Proceso completado.", vbInformation
End Sub
