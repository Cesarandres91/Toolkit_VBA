Sub ImportarMultiplesCSVyConvertir()
    Dim rutaArchivoActual As String
    Dim rutaCarpeta As String
    Dim archivoCSV As String
    Dim libroDestino As Workbook
    Dim libroOrigen As Workbook
    Dim hojaDestino As Worksheet
    Dim nombreArchivo As String
    Dim rangoOrigen As Range
    Dim contadorArchivos As Integer
    
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
    
    ' Buscar todos los archivos CSV en la carpeta
    archivoCSV = Dir(rutaCarpeta & "*.csv")
    
    Do While archivoCSV <> ""
        ' Incrementar contador
        contadorArchivos = contadorArchivos + 1
        
        ' Obtener el nombre del archivo sin extensión
        nombreArchivo = Left(archivoCSV, InStrRev(archivoCSV, ".") - 1)
        
        ' Abrir el archivo CSV
        Set libroOrigen = Workbooks.Open(rutaCarpeta & archivoCSV)
        
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
        
        ' Copiar todos los datos del CSV a la nueva hoja
        Set rangoOrigen = libroOrigen.Sheets(1).UsedRange
        rangoOrigen.Copy hojaDestino.Range("A1")
        
        ' Guardar el archivo CSV como XLSX
        libroOrigen.SaveAs Filename:=rutaCarpeta & nombreArchivo & ".xlsx", FileFormat:=xlOpenXMLWorkbook
        
        ' Cerrar el libro origen
        libroOrigen.Close SaveChanges:=False
        
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
