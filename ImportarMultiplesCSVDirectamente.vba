Sub ImportarMultiplesCSVDirectamente()
    Dim rutaArchivoActual As String
    Dim rutaCarpeta As String
    Dim archivoCSV As String
    Dim libroDestino As Workbook
    Dim hojaDestino As Worksheet
    Dim nombreArchivo As String
    Dim contadorArchivos As Integer
    Dim fso As Object
    Dim ts As Object
    Dim linea As String
    Dim datos() As String
    Dim fila As Long, columna As Long
    
    ' Optimizaciones
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    
    ' Obtener la ruta del libro actual
    rutaArchivoActual = ThisWorkbook.Path
    rutaCarpeta = rutaArchivoActual & "\"
    
    ' Referencia al libro de destino (libro actual)
    Set libroDestino = ThisWorkbook
    
    ' Inicializar contador
    contadorArchivos = 0
    
    ' Crear objeto FileSystemObject
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Buscar todos los archivos CSV en la carpeta
    archivoCSV = Dir(rutaCarpeta & "*.csv")
    
    Do While archivoCSV <> ""
        ' Incrementar contador
        contadorArchivos = contadorArchivos + 1
        
        ' Obtener el nombre del archivo sin extensión
        nombreArchivo = Left(archivoCSV, InStrRev(archivoCSV, ".") - 1)
        
        ' Crear una nueva hoja con el nombre del archivo
        On Error Resume Next
        Set hojaDestino = libroDestino.Worksheets(nombreArchivo)
        On Error GoTo 0
        
        If hojaDestino Is Nothing Then
            Set hojaDestino = libroDestino.Worksheets.Add(After:=libroDestino.Sheets(libroDestino.Sheets.Count))
            hojaDestino.Name = Left(nombreArchivo, 31) ' Excel limita los nombres de hoja a 31 caracteres
        Else
            hojaDestino.Cells.Clear
        End If
        
        ' Abrir el archivo CSV para lectura
        Set ts = fso.OpenTextFile(rutaCarpeta & archivoCSV, 1, False)
        
        ' Leer y copiar datos
        fila = 1
        Do While Not ts.AtEndOfStream
            linea = ts.ReadLine
            datos = Split(linea, ",") ' Asume que el separador es una coma
            For columna = 1 To UBound(datos) + 1
                hojaDestino.Cells(fila, columna).Value = datos(columna - 1)
            Next columna
            fila = fila + 1
        Loop
        
        ' Cerrar el archivo
        ts.Close
        
        ' Buscar el siguiente archivo CSV
        archivoCSV = Dir()
    Loop
    
    ' Restaurar configuraciones
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    
    If contadorArchivos = 0 Then
        MsgBox "No se encontraron archivos CSV en la carpeta.", vbExclamation
    Else
        MsgBox "Proceso completado. Se procesaron " & contadorArchivos & " archivos CSV.", vbInformation
    End If
End Sub
