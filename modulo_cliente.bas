Attribute VB_Name = "modulo_cliente"
Dim Cnn As ADODB.Connection
Dim TablaVendedor As ADODB.Recordset
Global Idvendedor As String
Global DatosAlmacenados As Boolean
Global BanderaCliente As Integer
Global ValorRifCedula As String


Public Sub LimpiarTextBoxCliente()
mantenimiento_cliente.t_codigo.Text = ""
mantenimiento_cliente.t_rifcedula.Text = ""
mantenimiento_cliente.t_nombrers.Text = ""
mantenimiento_cliente.t_direccion.Text = ""
mantenimiento_cliente.t_telefono1.Text = ""
mantenimiento_cliente.t_telefono2.Text = ""
mantenimiento_cliente.t_limitecredito.Text = ""
mantenimiento_cliente.c_tiempopago.Text = ""
mantenimiento_cliente.c_vendedor.Text = ""
mantenimiento_cliente.t_email.Text = ""
mantenimiento_cliente.c_vendedor.Clear
mantenimiento_cliente.c_tipodocumento.ListIndex = 0

End Sub

Public Sub DesabilitarTextboxCliente()
mantenimiento_cliente.t_rifcedula.Enabled = False
mantenimiento_cliente.t_nombrers.Enabled = False
mantenimiento_cliente.t_direccion.Enabled = False
mantenimiento_cliente.t_telefono1.Enabled = False
mantenimiento_cliente.t_telefono2.Enabled = False
mantenimiento_cliente.t_limitecredito.Enabled = False
mantenimiento_cliente.c_tiempopago.Enabled = False
mantenimiento_cliente.c_vendedor.Enabled = False
mantenimiento_cliente.c_tipodocumento.Enabled = False
mantenimiento_cliente.t_email.Enabled = False
End Sub
Public Sub HabilitarTextBoxCliente()
mantenimiento_cliente.t_rifcedula.Enabled = True
mantenimiento_cliente.t_nombrers.Enabled = True
mantenimiento_cliente.t_direccion.Enabled = True
mantenimiento_cliente.t_telefono1.Enabled = True
mantenimiento_cliente.t_telefono2.Enabled = True
mantenimiento_cliente.t_limitecredito.Enabled = True
mantenimiento_cliente.c_tiempopago.Enabled = True
mantenimiento_cliente.c_vendedor.Enabled = True
mantenimiento_cliente.c_tipodocumento.Enabled = True
mantenimiento_cliente.t_email.Enabled = True
End Sub
Public Function ValidarRuc(Valor)
If Len(Valor) = 11 Then
    If Mid(Valor, 1, 2) = 10 Or Mid(Valor, 1, 2) = 20 Or Mid(Valor, 1, 2) = 15 Or Mid(Valor, 1, 2) = 17 Then
        ValidarRuc = True
    Else
        MsgBox "Ingrese un formato de RUC válido."
        ValidarRuc = False
    End If
Else
    ValidarRuc = False
    MsgBox "El RUC debe tener 11 digitos."
End If

End Function
