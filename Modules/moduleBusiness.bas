Sub logout(form As UserForm) 'Cerrar sesion
    Dim response As Byte
    response = MsgBox("Deseas cerrar sesion?", vbYesNo + vbQuestion, "Sistema Kardex")
    Debug.Print "la respuesta fue: " & response
    If response = 6 Then
        Unload form
        'frmLogin.Show
    End If
End Sub

Sub backMenu(form As UserForm) ' Volver al Menu
    Unload form
    frmMenuPrincipal.Show
End Sub

Sub listProducts(comboboxP As combobox)
    Dim ultFila As Long
    ultFila = Sheets("PRODUCTOS").Range("A" & Rows.Count).End(xlUp).Row 'obteniendo la ultima fila de productos
    Set productos = Worksheets("PRODUCTOS").Range("A2:A" & ultFila) 'Rango de Productos
    comboboxP.Clear
    For Each celda In productos 'recorriendo el rango de productos y mostrandolos en el comboBox
        comboboxP.AddItem (celda.Value)
    Next celda
End Sub

Function FDU_filaNueva(hoja As String) As Long
   FDU_filaNueva = Sheets(hoja).Range("A" & Rows.Count).End(xlUp).Row + 1
End Function
