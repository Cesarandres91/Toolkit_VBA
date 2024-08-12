Sub ProcesarDatos()
    Dim wsBase As Worksheet
    Dim wsDatos1 As Worksheet
    Dim wbZZZ As Workbook
    Dim wsDatosZZZ As Worksheet
    Dim lista As Variant
    Dim ultimaFila As Long
    Dim rangoFiltrado As Range
    Dim rutaLibroZZZ As String
    
    ' Desactivar actualizaciones para optimizar
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    
    ' Configurar hoja base en el libro activo
    Set wsBase = ThisWorkbook.Sheets("base")
    
    ' Obtener la lista de valores
    ultimaFila = wsBase.Cells(wsBase.Rows.Count, "A").End(xlUp).Row
    lista = wsBase.Range("A2:A" & ultimaFila).Value
    
    ' Configurar ruta del libro ZZZ (asumiendo que está en la misma carpeta)
    rutaLibroZZZ = ThisWorkbook.Path & "\ZZZ.xlsx"
    
    ' Abrir libro ZZZ sin mostrarlo
    Set wbZZZ = Workbooks.Open(rutaLibroZZZ, UpdateLinks:=0, ReadOnly:=True)
    Set wsDatosZZZ = wbZZZ.Sheets("datos")
    
    ' Aplicar filtro
    With wsDatosZZZ
        .AutoFilterMode = False ' Eliminar cualquier filtro existente
        .Range("A1").AutoFilter
        .Range("A1").AutoFilter Field:=1, Criteria1:=lista, Operator:=xlFilterValues
        
        ' Verificar si hay filas filtradas (excluyendo la fila de encabezado)
        If Application.WorksheetFunction.Subtotal(103, .Range("A:A")) > 1 Then
            ' Copiar rango filtrado visible
            Set rangoFiltrado = .AutoFilter.Range.Offset(1).SpecialCells(xlCellTypeVisible)
            rangoFiltrado.Copy
            
            ' Crear o seleccionar hoja datos_1 en el libro original
            On Error Resume Next
            Set wsDatos1 = ThisWorkbook.Sheets("datos_1")
            On Error GoTo 0
            
            If wsDatos1 Is Nothing Then
                Set wsDatos1 = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
                wsDatos1.Name = "datos_1"
            Else
                wsDatos1.Cells.Clear ' Limpiar contenido existente
            End If
            
            ' Pegar datos filtrados
            wsDatos1.Range("A1").PasteSpecial xlPasteValues
            Application.CutCopyMode = False ' Limpiar portapapeles
        End If
        
        ' Quitar filtro
        .AutoFilterMode = False
    End With
    
    ' Cerrar libro ZZZ sin guardar cambios
    wbZZZ.Close SaveChanges:=False
    
    ' Reactivar configuraciones
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    
    MsgBox "Proceso completado", vbInformation
End Sub
