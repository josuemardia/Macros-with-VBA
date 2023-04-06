
Private Sub cmdCerrarSesion_Click()
    Call logout(Me)
End Sub

Private Sub cmdIngresos_Click()
    gTypeMovement = "Ingresos"
    Unload Me
    frmMovimientos.Show
End Sub

Private Sub cmdProductos_Click()
    Unload Me
    frmProductos.Show
End Sub

Private Sub activarControles(estado As Boolean)
    cmdProductos.Enabled = estado
    cmdIngresos.Enabled = estado
    cmdSalidas.Enabled = estado
    cmdReportes.Enabled = estado
End Sub

Private Sub cmdReportes_Click()
    Unload Me
    frmReportes.Show
End Sub

Private Sub cmdSalidas_Click()
    gTypeMovement = "Salidas"
    Unload Me
    frmMovimientos.Show
End Sub

Private Sub UserForm_Activate()
    'Mostrar Variables Globales en el label
    lblUserName.Caption = gNameUser
    lblUserPerfil.Caption = gTypeUser
    
    Select Case gTypeUser
        Case "Perfil 1": Call activarControles(True)
        Case "Perfil 2": Call activarControles(True): cmdReportes.Enabled = False
        Case "Perfil 3": Call activarControles(False): cmdReportes.Enabled = True
    End Select
End Sub


