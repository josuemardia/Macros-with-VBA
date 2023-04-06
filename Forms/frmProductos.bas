'Variables que solo se usaran en este form
Dim filaProdu As Long
Dim botonSelected As String
'Metodos personalizados que solo se usaran aqui----------------------------------------------------
Private Sub bloquearTexbox(estado As Boolean)
    txtStockInitial.Locked = estado
    txtStockLimit.Locked = estado
End Sub

Private Sub bloquearListaProdu(estado As Boolean)
    cboxDescription.Locked = estado
End Sub

Private Sub mostrarBotonesSave(estado As Boolean)
    cmdCancelar.Visible = estado
    cmdGuardarCambios.Visible = estado
End Sub

Private Sub bloquearBotones(estado As Boolean)
    cmdAgregar.Enabled = Not estado
    cmdEditar.Enabled = Not estado
    cmdEliminar.Enabled = Not estado
End Sub

Private Sub limpiarEntradas()
        txtDescription.Text = Empty
        cboxDescription.Text = Empty
        txtStockInitial.Text = Empty
        txtStockLimit.Text = Empty
End Sub


'''-------------------------------------------------------------------

Private Sub cboxDescription_Change()
    If productos.Find(cboxDescription.Text) Is Nothing Then
        txtStockInitial.Text = Empty
        txtStockLimit.Text = Empty
        Exit Sub
    End If
    filaProdu = productos.Find(cboxDescription.Text).Row 'encontrando la fila del producto seleccionado
    txtStockInitial.Text = Sheets("PRODUCTOS").Range("B" & filaProdu) 'mostrando el stock inicial que tiene el producto seleccionado
    txtStockLimit.Text = Sheets("PRODUCTOS").Range("C" & filaProdu) 'mostrando el stock limite que tiene el producto seleccionado
End Sub

Private Sub cmdAgregar_Click()
    txtStockInitial.Text = Empty
    txtStockLimit.Text = Empty
    Call bloquearTexbox(False)
    Call mostrarBotonesSave(True)
    txtDescription.Visible = True
    cboxDescription.Visible = False
    Call bloquearBotones(True)
    botonSelected = "agregar"

    
End Sub

Private Sub cmdCancelar_Click()
    Call mostrarBotonesSave(False)
    Call bloquearBotones(False)
    Call bloquearTexbox(True)
    Call limpiarEntradas
    txtDescription.Visible = False
    cboxDescription.Visible = True
    cboxDescription.Locked = False
End Sub

Private Sub cmdCerrarSesion_Click()
    Call logout(Me) 'cierra sesion - procedimiento en el modulo 1
End Sub

Private Sub cmdEditar_Click()
    If cboxDescription.Text = Empty Then
        MsgBox "No ha seleccionado un producto", vbExclamation + vbOKOnly, Me.Caption
    Else
        Call bloquearTexbox(False)
        Call mostrarBotonesSave(True)
        Call bloquearBotones(True)
        Call bloquearListaProdu(True)
        botonSelected = "editar"
    End If
End Sub

Private Sub cmdEliminar_Click()
    If cboxDescription.Text = Empty Then
        MsgBox "No ha seleccionado un producto", vbExclamation + vbOKOnly, Me.Caption
    Else
        Dim response As Byte
        response = MsgBox("Esta seguro de eliminar este producto: " & cboxDescription.Text, vbInformation + vbYesNo, Me.Caption)
        If response = 6 Then
            With Sheets("PRODUCTOS")
                .Rows(filaProdu & ":" & filaProdu).Select
                Selection.Delete Shift:=xlUp
            End With
            txtStockInitial.Text = Empty
            txtStockLimit.Text = Empty
            Call listProducts(cboxDescription)
            MsgBox "Producto eliminado exitosamente!", vbInformation + vbOKOnly, Me.Caption
        End If
    End If
End Sub

Private Sub cmdGuardarCambios_Click()
    If txtStockInitial.Text = Empty Or txtStockLimit.Text = Empty Then
        MsgBox "Complete todos los campos por favor", vbExclamation + vbOKOnly, Me.Caption
    Else
        If (Not VBA.IsNumeric(txtStockInitial.Text) Or Not VBA.IsNumeric(txtStockLimit.Text)) Then
            MsgBox "Solo se acepta numeros para los stocks", vbCritical + vbOKOnly, Me.Caption
            Exit Sub
        End If
        If txtStockInitial.Text > 0 And txtStockLimit.Text > 0 Then
            With Sheets("PRODUCTOS")
                Select Case botonSelected
                    Case "editar":
                        If cboxDescription.Text = Empty Then
                            MsgBox "Escriba el producto a editar", vbExclamation + vbOKOnly, Me.Caption
                            Exit Sub
                        End If
                        .Range("B" & filaProdu).Value = txtStockInitial.Text
                        .Range("C" & filaProdu).Value = txtStockLimit.Text
                        Call mostrarBotonesSave(False)
                        Call bloquearTexbox(True)
                        Call bloquearBotones(False)
                        Call listProducts(cboxDescription)
                        cboxDescription.Locked = False
                         MsgBox "Producto actualizado con exito!", vbInformation + vbOKOnly, Me.Caption
                    Case "agregar":
                        If txtDescription.Text = Empty Then
                            MsgBox "Escriba el producto a agregar", vbExclamation + vbOKOnly, Me.Caption
                            Exit Sub
                        End If
                        If VBA.IsNumeric(txtDescription.Text) Then
                            MsgBox "Los nombres de productos no pueden ser numeros", vbCritical + vbOKOnly, Me.Caption
                            Exit Sub
                        End If
                        If (VBA.UCase(txtDescription.Text) = txtDescription.Text) Then
                            If productos.Find(txtDescription.Text) Is Nothing Then
                                Dim filaNueva As Long
                                filaNueva = FDU_filaNueva("PRODUCTOS")
                                .Range("A" & filaNueva).Value = txtDescription.Text
                                .Range("B" & filaNueva).Value = txtStockInitial.Text
                                .Range("C" & filaNueva).Value = txtStockLimit.Text
                                Call mostrarBotonesSave(False)
                                Call bloquearTexbox(True)
                                Call bloquearBotones(False)
                                Call limpiarEntradas
                                txtDescription.Visible = False
                                cboxDescription.Visible = True
                                Call listProducts(cboxDescription)
                                MsgBox "Producto agregado con exito!", vbInformation + vbOKOnly, Me.Caption
                            Else
                                MsgBox "Este producto ya existe", vbCritical + vbOKOnly, Me.Caption
                            End If
                        Else
                         MsgBox "Los productos deben ser escritos en mayuscula", vbExclamation + vbOKOnly, Me.Caption
                        End If
                End Select
            End With
        Else
           MsgBox "Las cantidades de los stocks no son validas, no son mayores a 0", vbCritical + vbOKOnly, Me.Caption
        End If
    End If
End Sub

Private Sub cmdVolverMenu_Click()
    Call backMenu(Me) 'vuelve al home - procedimiento en el modulo 1
End Sub


Private Sub UserForm_Activate()
    Call listProducts(cboxDescription)
    Call bloquearTexbox(True)
    cmdGuardarCambios.Visible = False
    cmdCancelar.Visible = False
End Sub

