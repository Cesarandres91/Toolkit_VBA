Sub EliminarEspaciosColumnaA()
    Dim ultimaFila As Long
    Dim celda As Range
    
    ' Determinar la última fila con datos en la columna A
    ultimaFila = ActiveSheet.Cells(Rows.Count, "A").End(xlUp).Row
    
    ' Recorrer cada celda en la columna A
    For Each celda In ActiveSheet.Range("A1:A" & ultimaFila)
        ' Eliminar espacios al inicio y al final del texto
        celda.Value = Trim(celda.Value)
    Next celda
    
    MsgBox "Se han eliminado los espacios en blanco al inicio y al final de las celdas en la columna A.", vbInformation
End Sub
