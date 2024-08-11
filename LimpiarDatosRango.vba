Sub LimpiarDatosRango()
    Dim ws As Worksheet
    Dim rng As Range
    Dim cel As Range
    Dim primeraCol As String, ultimaCol As String
    
    ' Configurar la hoja de trabajo
    Set ws = ActiveSheet
    
    ' Solicitar al usuario el rango de columnas
    primeraCol = InputBox("Ingrese la letra de la primera columna del rango:", "Rango de limpieza", "A")
    ultimaCol = InputBox("Ingrese la letra de la última columna del rango:", "Rango de limpieza", "Z")
    
    ' Definir el rango
    Set rng = ws.Range(primeraCol & ":" & ultimaCol)
    
    ' Iniciar la limpieza de datos
    For Each cel In rng.Cells
        ' Reemplazar NULL por cadena vacía
        If IsNull(cel.Value) Then
            cel.Value = ""
        End If
        
        ' Eliminar espacios en blanco al inicio y al final
        If VarType(cel.Value) = vbString Then
            cel.Value = Trim(cel.Value)
        End If
    Next cel
    
    MsgBox "Limpieza de datos completada para el rango " & primeraCol & ":" & ultimaCol, vbInformation
End Sub
