Sub EliminarFilasYColumnas()
    Dim ws As Worksheet
    Dim filasAEliminar As Long
    Dim columnasAEliminar As Long
    
    ' Configurar la hoja de trabajo
    Set ws = ActiveSheet
    
    ' Solicitar al usuario el número de filas y columnas a eliminar
    filasAEliminar = ObtenerNumero("Ingrese el número de filas a eliminar:", "Eliminar Filas")
    columnasAEliminar = ObtenerNumero("Ingrese el número de columnas a eliminar:", "Eliminar Columnas")
    
    ' Eliminar filas
    If filasAEliminar > 0 Then
        ws.Rows("1:" & filasAEliminar).Delete Shift:=xlUp
    End If
    
    ' Eliminar columnas
    If columnasAEliminar > 0 Then
        ws.Columns(1).Resize(, columnasAEliminar).Delete Shift:=xlToLeft
    End If
    
    MsgBox "Se han eliminado " & filasAEliminar & " filas y " & columnasAEliminar & " columnas.", vbInformation
End Sub

Function ObtenerNumero(prompt As String, title As String) As Long
    Dim inputValue As String
    Dim numero As Long
    
    Do
        inputValue = InputBox(prompt, title, "0")
        If inputValue = "" Then Exit Function ' El usuario canceló
        If IsNumeric(inputValue) Then
            numero = CLng(inputValue)
            If numero >= 0 Then
                ObtenerNumero = numero
                Exit Function
            End If
        End If
        MsgBox "Por favor, ingrese un número entero no negativo.", vbExclamation
    Loop
End Function
