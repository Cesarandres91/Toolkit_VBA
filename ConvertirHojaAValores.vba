Sub ConvertirHojaAValores()
    Dim ultimaFila As Long
    Dim ultimaColumna As Long
    Dim rango As Range
    
    ' Determinar la última fila y columna con datos
    ultimaFila = ActiveSheet.Cells(Rows.Count, 1).End(xlUp).Row
    ultimaColumna = ActiveSheet.Cells(1, Columns.Count).End(xlToLeft).Column
    
    ' Definir el rango de toda la hoja con datos
    Set rango = ActiveSheet.Range(ActiveSheet.Cells(1, 1), ActiveSheet.Cells(ultimaFila, ultimaColumna))
    
    ' Convertir el rango a valores
    rango.Value = rango.Value
    
    ' Cambiar el formato de las celdas a General
    rango.NumberFormat = "General"
End Sub
