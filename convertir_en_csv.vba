Sub ConvertirCSVaXLSX()
    Dim rutaArchivoCSV As String
    Dim rutaArchivoXLSX As String
    Dim libroCSV As Workbook
    
    ' Solicitar al usuario que seleccione el archivo CSV
    rutaArchivoCSV = ObtenerRutaArchivo("CSV Files (*.csv), *.csv")
    If rutaArchivoCSV = "" Then Exit Sub
    
    ' Abrir el archivo CSV
    Set libroCSV = Workbooks.Open(rutaArchivoCSV)
    
    ' Generar la ruta para el nuevo archivo XLSX
    rutaArchivoXLSX = Replace(rutaArchivoCSV, ".csv", ".xlsx")
    
    ' Guardar como XLSX
    libroCSV.SaveAs Filename:=rutaArchivoXLSX, FileFormat:=xlOpenXMLWorkbook
    
    ' Cerrar el libro
    libroCSV.Close SaveChanges:=False
    
    MsgBox "Conversión completada. El archivo se ha guardado como: " & rutaArchivoXLSX, vbInformation
End Sub

Function ObtenerRutaArchivo(filtro As String) As String
    Dim dlg As FileDialog
    Set dlg = Application.FileDialog(msoFileDialogFilePicker)
    
    With dlg
        .Filters.Clear
        .Filters.Add "Archivos CSV", "*.csv"
        .Title = "Seleccione el archivo CSV a convertir"
        .AllowMultiSelect = False
        
        If .Show = -1 Then
            ObtenerRutaArchivo = .SelectedItems(1)
        Else
            ObtenerRutaArchivo = ""
        End If
    End With
End Function
