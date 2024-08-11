Sub ConvertirColumnaANumeros()
    Dim ultimaFila As Long
    Dim rango As Range
    
    ' Determinar la última fila con datos en la columna A
    ultimaFila = ActiveSheet.Cells(Rows.Count, "A").End(xlUp).Row
    
    ' Definir el rango de la columna A
    Set rango = ActiveSheet.Range("A1:A" & ultimaFila)
    
    ' Convertir el rango a valores
    rango.Value = rango.Value
    
    ' Cambiar el formato de las celdas a Número
    rango.NumberFormat = "General"
End Sub
