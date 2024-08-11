Sub LeftJoinHojas()
    Dim wsIzquierda As Worksheet, wsDerecha As Worksheet, wsResultado As Worksheet
    Dim hojaIzquierda As String, hojaDerecha As String
    Dim columnaClaveIzq As String, columnaClaveDer As String
    Dim ultimaFilaIzq As Long, ultimaFilaDer As Long, ultimaColumnaIzq As Long, ultimaColumnaDer As Long
    Dim i As Long, j As Long, k As Long
    Dim dicDerecha As Object
    Dim clave As Variant
    Dim encontrado As Boolean
    
    ' Solicitar nombres de hojas y columnas clave
    hojaIzquierda = InputBox("Ingrese el nombre de la primera hoja (izquierda del JOIN):", "Hoja Izquierda")
    columnaClaveIzq = InputBox("Ingrese la letra de la columna clave en la hoja " & hojaIzquierda & ":", "Columna Clave Izquierda")
    hojaDerecha = InputBox("Ingrese el nombre de la segunda hoja (derecha del JOIN):", "Hoja Derecha")
    columnaClaveDer = InputBox("Ingrese la letra de la columna clave en la hoja " & hojaDerecha & ":", "Columna Clave Derecha")
    
    ' Verificar si las hojas existen
    On Error Resume Next
    Set wsIzquierda = ThisWorkbook.Worksheets(hojaIzquierda)
    Set wsDerecha = ThisWorkbook.Worksheets(hojaDerecha)
    On Error GoTo 0
    
    If wsIzquierda Is Nothing Or wsDerecha Is Nothing Then
        MsgBox "Una o ambas hojas no existen. Por favor, verifique los nombres.", vbExclamation
        Exit Sub
    End If
    
    ' Crear nueva hoja para el resultado
    Set wsResultado = ThisWorkbook.Worksheets.Add
    wsResultado.Name = "JOIN_" & hojaIzquierda & "_" & hojaDerecha
    
    ' Obtener últimas filas y columnas
    ultimaFilaIzq = wsIzquierda.Cells(wsIzquierda.Rows.Count, columnaClaveIzq).End(xlUp).Row
    ultimaFilaDer = wsDerecha.Cells(wsDerecha.Rows.Count, columnaClaveDer).End(xlUp).Row
    ultimaColumnaIzq = wsIzquierda.Cells(1, wsIzquierda.Columns.Count).End(xlToLeft).Column
    ultimaColumnaDer = wsDerecha.Cells(1, wsDerecha.Columns.Count).End(xlToLeft).Column
    
    ' Copiar encabezados
    wsIzquierda.Rows(1).Copy Destination:=wsResultado.Rows(1)
    wsDerecha.Range(wsDerecha.Cells(1, 1), wsDerecha.Cells(1, ultimaColumnaDer)).Copy _
        Destination:=wsResultado.Cells(1, ultimaColumnaIzq + 1)
    
    ' Crear diccionario para la hoja derecha
    Set dicDerecha = CreateObject("Scripting.Dictionary")
    For i = 2 To ultimaFilaDer
        clave = wsDerecha.Cells(i, columnaClaveDer).Value
        If Not dicDerecha.Exists(clave) Then
            dicDerecha.Add clave, New Collection
        End If
        dicDerecha(clave).Add i
    Next i
    
    ' Realizar el LEFT JOIN
    k = 2 ' Fila inicial en la hoja de resultado
    For i = 2 To ultimaFilaIzq
        clave = wsIzquierda.Cells(i, columnaClaveIzq).Value
        If dicDerecha.Exists(clave) Then
            For Each j In dicDerecha(clave)
                ' Copiar fila de la hoja izquierda
                wsIzquierda.Rows(i).Copy Destination:=wsResultado.Rows(k)
                ' Copiar datos correspondientes de la hoja derecha
                wsDerecha.Range(wsDerecha.Cells(j, 1), wsDerecha.Cells(j, ultimaColumnaDer)).Copy _
                    Destination:=wsResultado.Cells(k, ultimaColumnaIzq + 1)
                k = k + 1
            Next j
        Else
            ' Si no hay coincidencia, copiar solo la fila de la hoja izquierda
            wsIzquierda.Rows(i).Copy Destination:=wsResultado.Rows(k)
            k = k + 1
        End If
    Next i
    
    ' Ajustar columnas
    wsResultado.Columns.AutoFit
    
    MsgBox "LEFT JOIN completado. Resultados en la hoja '" & wsResultado.Name & "'.", vbInformation
End Sub
