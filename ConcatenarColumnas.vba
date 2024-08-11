Sub ConcatenarColumnas()
    Dim ws As Worksheet
    Dim ultimaFila As Long
    Dim columnas() As String
    Dim columnaDestino As String
    Dim separador As String
    Dim i As Long, j As Long
    Dim valorConcatenado As String
    
    ' Configurar la hoja de trabajo
    Set ws = ActiveSheet
    ultimaFila = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    ' Solicitar al usuario las columnas a concatenar
    columnas = Split(InputBox("Ingrese las letras de las columnas a concatenar, separadas por comas (por ejemplo, A,B,C):", "Columnas a Concatenar"), ",")
    
    ' Solicitar la columna de destino
    columnaDestino = InputBox("Ingrese la letra de la columna donde desea el resultado:", "Columna de Destino")
    
    ' Solicitar el separador
    separador = InputBox("Ingrese el separador que desea usar entre los valores (deje en blanco si no desea separador):", "Separador")
    
    ' Verificar si la columna de destino ya tiene datos
    If WorksheetFunction.CountA(ws.Columns(columnaDestino)) > 0 Then
        If MsgBox("La columna " & columnaDestino & " ya contiene datos. ¿Desea sobrescribirlos?", vbYesNo + vbQuestion) = vbNo Then
            Exit Sub
        End If
    End If
    
    ' Concatenar las columnas
    For i = 1 To ultimaFila
        valorConcatenado = ""
        For j = 0 To UBound(columnas)
            If j > 0 And Len(valorConcatenado) > 0 Then
                valorConcatenado = valorConcatenado & separador
            End If
            valorConcatenado = valorConcatenado & ws.Cells(i, Trim(columnas(j))).Value
        Next j
        ws.Cells(i, columnaDestino).Value = valorConcatenado
    Next i
    
    MsgBox "Concatenación completada en la columna " & columnaDestino & ".", vbInformation
End Sub
