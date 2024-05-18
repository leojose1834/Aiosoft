Attribute VB_Name = "modulo_proveedor"
Dim CnnProveedor As ADODB.Connection
Dim TablaVendedorProveedor As ADODB.Recordset
Global IdvendedorProveedor As String
Global DatosAlmacenadosProveedor As Boolean
Global ValorRifCedulaProveedor As String


Public Sub LimpiarTextBoxProveedor()
mantenimiento_proveedor.t_codigo.Text = ""
mantenimiento_proveedor.t_rifcedula.Text = ""
mantenimiento_proveedor.t_nombrers.Text = ""
mantenimiento_proveedor.t_direccion.Text = ""
mantenimiento_proveedor.t_telefono1.Text = ""
mantenimiento_proveedor.t_telefono2.Text = ""
mantenimiento_proveedor.t_limitecredito.Text = ""
mantenimiento_proveedor.c_tiempopago.Text = ""
mantenimiento_proveedor.c_vendedor.Text = ""
mantenimiento_proveedor.t_email.Text = ""
mantenimiento_proveedor.c_vendedor.Clear
mantenimiento_proveedor.c_tipodocumento.ListIndex = 0

End Sub

Public Sub DesabilitarTextboxProveedor()
mantenimiento_proveedor.t_rifcedula.Enabled = False
mantenimiento_proveedor.t_nombrers.Enabled = False
mantenimiento_proveedor.t_direccion.Enabled = False
mantenimiento_proveedor.t_telefono1.Enabled = False
mantenimiento_proveedor.t_telefono2.Enabled = False
mantenimiento_proveedor.t_limitecredito.Enabled = False
mantenimiento_proveedor.c_tiempopago.Enabled = False
mantenimiento_proveedor.c_vendedor.Enabled = False
mantenimiento_proveedor.c_tipodocumento.Enabled = False
mantenimiento_proveedor.t_email.Enabled = False
End Sub
Public Sub HabilitarTextBoxProveedor()
mantenimiento_proveedor.t_rifcedula.Enabled = True
mantenimiento_proveedor.t_nombrers.Enabled = True
mantenimiento_proveedor.t_direccion.Enabled = True
mantenimiento_proveedor.t_telefono1.Enabled = True
mantenimiento_proveedor.t_telefono2.Enabled = True
mantenimiento_proveedor.t_limitecredito.Enabled = True
mantenimiento_proveedor.c_tiempopago.Enabled = True
mantenimiento_proveedor.c_vendedor.Enabled = True
mantenimiento_proveedor.c_tipodocumento.Enabled = True
mantenimiento_proveedor.t_email.Enabled = True
End Sub
Public Function ValidarRucProveedor(Valor)
If Len(Valor) = 11 Then
    If Mid(Valor, 1, 2) = 10 Or Mid(Valor, 1, 2) = 20 Or Mid(Valor, 1, 2) = 15 Or Mid(Valor, 1, 2) = 17 Then
        ValidarRucProveedor = True
    Else
        MsgBox "Ingrese un formato de RUC válido."
        ValidarRucProveedor = False
    End If
Else
    ValidarRucProveedor = False
    MsgBox "El RUC debe tener 11 digitos."
End If

End Function

