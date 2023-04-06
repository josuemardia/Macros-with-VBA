Dim filaProdu As Long 'variable que guarda la fila del producto seleccionado en el combobox
Private Sub cboxProducto_Change()
    If productos.Find(cboxProducto.Text) Is Nothing Then
        txtStockActual.Text = Empty
        txtCantidad.Text = Empty
        Exit Sub
    End If
    filaProdu = productos.Find(cboxProducto.Text).Row 'encontrando la fila del producto seleccionado
    txtStockActual.Text = Sheets("PRODUCTOS").Range("B" & filaProdu) 'mostrando el stock inicial que tiene el producto seleccionado
End Sub

Private Sub cmdCerrarSesion_Click()
    Call logout(Me)
End Sub

Private Sub cmdGuardar_Click()
    Dim stock As Integer
    
    If (txtCantidad.Text = Empty Or cboxProducto.Text = Empty Or cboxProducto.Text = Empty) Then 'validando si el campo esta vacio
        MsgBox "Complete todos los campos por favor", vbExclamation + vbOKOnly, Me.Caption
    Else
        If Not (VBA.IsNumeric(txtCantidad.Text)) Then
            MsgBox "La cantidad debe ser expresada en numeros", vbCritical + vbOKOnly, Me.Caption
            Exit Sub
        End If
        If txtCantidad.Text > 0 Then 'validando que la cantidad sea valida
            Select Case gTypeMovement
                Case "Ingresos":
                    Dim filNueva As Long
                    Sheets("PRODUCTOS").Range("B" & filaProdu).Value = Sheets("PRODUCTOS").Range("B" & filaProdu) + txtCantidad.Text
                    stock = Sheets("PRODUCTOS").Range("B" & filaProdu).Value
                    txtStockActual.Text = stock
                    filNueva = FDU_filaNueva("MOVIMIENTOS")
                    With Sheets("MOVIMIENTOS")
                        .Range("A" & filNueva).Value = cboxProducto.Text
                        .Range("B" & filNueva).Value = gTypeMovement
                        .Range("C" & filNueva).Value = txtCantidad.Text
                    End With
                    txtCantidad.Text = Empty
                    MsgBox "Insercion guardada con exito!", vbInformation + vbOKOnly, Me.Caption
                Case "Salidas":
                    If txtCantidad.Text < Sheets("PRODUCTOS").Range("B" & filaProdu).Value Then
                        Sheets("PRODUCTOS").Range("B" & filaProdu).Value = Sheets("PRODUCTOS").Range("B" & filaProdu) - txtCantidad.Text
                        stock = Sheets("PRODUCTOS").Range("B" & filaProdu).Value
                        txtStockActual.Text = stock
                        filNueva = FDU_filaNueva("MOVIMIENTOS")
                        With Sheets("MOVIMIENTOS")
                            .Range("A" & filNueva).Value = cboxProducto.Text
                            .Range("B" & filNueva).Value = gTypeMovement
                            .Range("C" & filNueva).Value = txtCantidad.Text
                        End With
                        txtCantidad.Text = Empty
                        MsgBox "Salida realizada con exito!", vbInformation + vbOKOnly, Me.Caption
                    Else
                        MsgBox "La cantidad supera al stock, salida invalida", vbCritical + vbOKOnly, Me.Caption
                    End If
            End Select
        Else
            MsgBox "Ingrese una cantidad valida por favor", vbExclamation + vbOKOnly, Me.Caption
        End If
    End If
End Sub

Private Sub cmdVolverMenu_Click()
    Call backMenu(Me)
End Sub


Private Sub UserForm_Activate()
    Me.Caption = gTypeMovement 'Dando nombre al titulo del form segun la seleccion del usuario: ingresos | salidas
    lblMovimientos.Caption = gTypeMovement 'nombrando al label: ingresos | salidas
    Call listProducts(cboxProducto) ' cargando la lista de productos de la bd del excel
End Sub

