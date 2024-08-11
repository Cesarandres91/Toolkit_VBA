Sub NormalizarDatos()
    Dim ws As Worksheet
    Dim lastRow As Long, lastCol As Long
    Dim i As Long, j As Long, k As Long
    Dim staticCols As Long
    Dim repeatingCols As Long
    Dim outputRow As Long
    Dim titles() As String
    
    ' Configurar la hoja de trabajo
    Set ws = ActiveSheet
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    
    ' Configuración personalizable
    staticCols = InputBox("Ingrese el número de columnas estáticas al inicio:", "Configuración", "2")
    repeatingCols = InputBox("Ingrese el número de columnas que se repiten por mes:", "Configuración", "2")
    
    ' Preparar la hoja de salida
    Dim wsOutput As Worksheet
    Set wsOutput = Worksheets.Add
    wsOutput.Name = "Datos Normalizados"
    
    ' Copiar encabezados estáticos
    For i = 1 To staticCols
        wsOutput.Cells(1, i).Value = ws.Cells(1, i).Value
    Next i
    
    ' Preparar títulos para columnas repetitivas
    ReDim titles(1 To repeatingCols)
    For i = 1 To repeatingCols
        titles(i) = InputBox("Ingrese el título para la columna repetitiva " & i & ":", "Títulos", "Columna" & i)
        wsOutput.Cells(1, staticCols + i).Value = titles(i)
    Next i
    
    wsOutput.Cells(1, staticCols + repeatingCols + 1).Value = "Mes"
    
    ' Proceso de normalización
    outputRow = 2
    For i = 2 To lastRow ' Para cada fila de datos
        For j = staticCols + 1 To lastCol Step repeatingCols ' Para cada grupo de columnas repetitivas
            ' Copiar datos estáticos
            For k = 1 To staticCols
                wsOutput.Cells(outputRow, k).Value = ws.Cells(i, k).Value
            Next k
            
            ' Copiar datos repetitivos
            For k = 1 To repeatingCols
                wsOutput.Cells(outputRow, staticCols + k).Value = ws.Cells(i, j + k - 1).Value
            Next k
            
            ' Agregar el mes
            wsOutput.Cells(outputRow, staticCols + repeatingCols + 1).Value = ws.Cells(1, j).Value
            
            outputRow = outputRow + 1
        Next j
    Next i
    
    ' Ajustar columnas
    wsOutput.Columns.AutoFit
    
    MsgBox "Normalización completada. Los datos se han guardado en una nueva hoja llamada 'Datos Normalizados'.", vbInformation
End Sub
