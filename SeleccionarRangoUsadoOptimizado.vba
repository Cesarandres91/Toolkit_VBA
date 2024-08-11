Sub SeleccionarRangoUsadoOptimizado()
    Dim ws As Worksheet
    Dim ultimaFila As Long
    Dim ultimaColumna As Long
    Dim rangoUsado As Range
    Dim fila As Long, columna As Long
    
    ' Optimizaciones para mejorar el rendimiento
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    ' Establecer la hoja de trabajo actual
    Set ws = ActiveSheet
    
    ' Encontrar la última columna con datos
    ultimaColumna = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    
    ' Encontrar la última fila con datos
    ultimaFila = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    ' Verificar otras columnas para asegurar que no perdemos ninguna fila
    For columna = 2 To ultimaColumna
        fila = ws.Cells(ws.Rows.Count, columna).End(xlUp).Row
        If fila > ultimaFila Then ultimaFila = fila
    Next columna
    
    ' Verificar la última fila para asegurar que no perdemos ninguna columna
    For columna = ultimaColumna + 1 To ws.Columns.Count
        If Not IsEmpty(ws.Cells(ultimaFila, columna)) Then
            ultimaColumna = columna
        Else
            ' Salir si encontramos 50 columnas vacías consecutivas
            If columna > ultimaColumna + 50 Then Exit For
        End If
    Next columna
    
    ' Definir el rango usado
    Set rangoUsado = ws.Range(ws.Cells(1, 1), ws.Cells(ultimaFila, ultimaColumna))
    
    ' Seleccionar el rango usado
    rangoUsado.Select
    
    ' Restaurar configuraciones
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    
    ' Mostrar un mensaje con la información del rango seleccionado
    MsgBox "Rango seleccionado: " & rangoUsado.Address, vbInformation
End Sub
