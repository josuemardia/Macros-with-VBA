
Private Sub cmdRegistrarCuenta_Click()
 If txtNuevoUsuario.Text = Empty Then
    MsgBox "iCompletar Campo Nombre de Usuario!", vbExclamation + vbOKOnly, Me.Caption
    txtNuevoUsuario.SetFocus
 Else
    If txtNuevoPassword.Text = Empty Then
        MsgBox "iCompletar Campo Password!", vbExclamation + vbOKOnly, Me.Caption
        txtNuevoPassword.SetFocus
    Else
        If cboTipoUsuario.ListIndex = -1 Then
            MsgBox "iSeleccione un perfil por favor!", vbExclamation + vbOKOnly, Me.Caption
            cboTipoUsuario.SetFocus
        Else
            If (UsuarioExistente(txtNuevoUsuario.Text)) Then
                'Usaurio Ya existe en la Base de Datos
                MsgBox "iEl Usuario: " & txtNuevoUsuario.Text & " ya existe!", vbCritical + vbOKOnly, Me.Caption
            Else
                'Registrar nuevo usuario
                Call GuardarNuevoUsuario(txtNuevoUsuario.Text, txtNuevoPassword.Text, cboTipoUsuario.Text)
                MsgBox "Usuario Registrado Correctamente!", vbInformation + vbOKOnly, Me.Caption
                Unload Me
                frmLogin.Show
            End If
            
        End If
    End If
 End If
End Sub

Private Sub GuardarNuevoUsuario(username As String, password As String, perfil As String)
    Dim FilaN As Long
    Sheets("USUARIOS").Select
    FilaN = Range("A" & Rows.Count).End(xlUp).Row + 1 'Fila nueva
    'Copiar datos a la hoja de calculo
    Cells(FilaN, 1).Value = username 'Nombre Usuario
    Cells(FilaN, 2).Value = password 'Clave
    Cells(FilaN, 3).Value = perfil 'Tipo Usuario

    'Acomodar Tabla Usuarios
    Range("A" & FilaN).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .ReadingOrder = xlContext
    End With
    Sheets("INICIO").Select
End Sub
Function UsuarioExistente(newuser As String) As Boolean
    'Rango de Usuarios Disponibles
    Sheets("USUARIOS").Select
    Dim Data As Range
    Dim nRows As Long
    nRows = Range("A" & Rows.Count).End(xlUp).Row
    'Seteando el rango que vamos a buscar
    Set Data = Range("A2", "A" & nRows)
    'Verificar si existe el Usuario en Data
    If Data.Find(newuser) Is Nothing Then
        UsuarioExistente = False
    Else
        UsuarioExistente = True
    End If
    Sheets("INICIO").Select
End Function

Private Sub UserForm_Activate()
    cboTipoUsuario.AddItem "Perfil 1"
    cboTipoUsuario.AddItem "Perfil 2"
    cboTipoUsuario.AddItem "Perfil 3"
End Sub
