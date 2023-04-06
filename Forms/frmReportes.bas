
Private Sub cmdCerrarSesionR_Click()
    Call logout(Me)
End Sub

Private Sub cmdVolverMenuR_Click()
    Call backMenu(Me)
End Sub

Private Sub UserForm_Activate()
    'Propiedades y Header del ListBox
    lboxReportes.ColumnCount = 4
    lboxReportes.ColumnWidths = "150;100;100;150"
    lboxReportes.AddItem
    lboxReportes.List(0, 0) = "Producto"
    lboxReportes.List(0, 1) = "Stock"
    lboxReportes.List(0, 2) = "StockMin"
    lboxReportes.List(0, 3) = "Alerta"
    
    'Obtener los productos a mostrar
    Sheets("PRODUCTOS").Select
    Dim Data As Range
    Dim nRows As Integer
    nRows = Range("A" & Rows.Count).End(xlUp).Row
    'Seteando el rango que vamos a buscar
    Set Data = Range("A2", "A" & nRows)
    
    Dim i As Integer
    For i = 1 To nRows - 1
        Dim stock As Integer
        Dim stockM As Integer
        stock = CInt(Range("B" & i + 1).Value)
        stockM = CInt(Range("C" & i + 1).Value)
        lboxReportes.AddItem
        lboxReportes.List(i, 0) = Range("A" & i + 1).Value
        lboxReportes.List(i, 1) = CStr(stock)
        lboxReportes.List(i, 2) = CStr(stockM)
        
        'Mensaje Stock Estado
        If (stock < stockM) Then
            lboxReportes.List(i, 3) = "Abastercer"
        Else
            lboxReportes.List(i, 3) = "OK"
        End If
    
    Next
End Sub

Private Sub UserForm_Click()

End Sub