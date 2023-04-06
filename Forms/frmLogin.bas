
Private Sub cmdIngresar_Click()
    'Rango de Usuarios Disponibles
    Sheets("USUARIOS").Select
    Dim Data As Range
    Dim nRows As Long
    nRows = Range("A" & Rows.Count).End(xlUp).Row
    
    'Seteando el rango que vamos a buscar
    Set Data = Range("A2", "A" & nRows)
    
    'Verificar si existe el Usuario en Data
    If Data.Find(txtNameUser.Text) Is Nothing Then
        MsgBox "!El Usuario Ingresado No existe!", vbExclamation + vbOKOnly, Me.Caption
    Else
        'Usuario Valido, ahora verificar su password
        Dim PasswordBD As String
        PasswordBD = Range("B" & Data.Find(txtNameUser.Text).Row).Value
        If (txtPasswordUser.Text = PasswordBD) Then
            'Setear Variables Globales
            gNameUser = txtNameUser.Text
            gTypeUser = Range("C" & Data.Find(txtNameUser.Text).Row).Value
            gCounterLogin = 0
            'Mensaje de Confirmacion
            MsgBox "Bienvenido: " & txtNameUser.Text, vbInformation + vbOKOnly, Me.Caption
            'Cargar Menu Principal
            Unload Me
            frmMenuPrincipal.Show
        Else
            gCounterLogin = gCounterLogin + 1
            If gCounterLogin = 3 Then
                MsgBox "Limite de Intentos Alcanzados", vbCritical + vbOKOnly, Me.Caption
                Unload Me
            Else
             MsgBox "!Password Incorrecto,Te quedan:" & (3 - gCounterLogin) & " intentos", vbCritical + vbOKOnly, Me.Caption
            End If
           
        End If
    End If
    Sheets("INICIO").Select
End Sub

Private Sub cmdNuevaCuenta_Click()
    Unload Me
    frmNuevaCuenta.Show
End Sub

Private Sub UserForm_Click()

End Sub

Private Sub UserForm_Initialize()
    gCounterLogin = 0
End Sub
