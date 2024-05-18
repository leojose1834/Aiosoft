Attribute VB_Name = "modulo_vendedor"
Global IDVendedor2 As String
Global NuevoRegistro As Boolean
Global CnnVendedor As ADODB.Connection
Global RstVendedor As ADODB.Recordset
Public Sub HabilitarTextboxVendedor()
f_vendedores.t_idvendedor.Enabled = True
f_vendedores.t_nombres.Enabled = True
f_vendedores.t_apellidos.Enabled = True
f_vendedores.t_direccion.Enabled = True
f_vendedores.t_telefono.Enabled = True
f_vendedores.t_xcentajecomision.Enabled = True

End Sub

Public Sub DeshabilitarTextboxVendedor()
f_vendedores.t_idvendedor.Enabled = False
f_vendedores.t_nombres.Enabled = False
f_vendedores.t_apellidos.Enabled = False
f_vendedores.t_direccion.Enabled = False
f_vendedores.t_telefono.Enabled = False
f_vendedores.t_xcentajecomision.Enabled = False
End Sub

Public Sub LimpiarTextboxVendedor()
f_vendedores.t_idvendedor.Text = ""
f_vendedores.t_nombres.Text = ""
f_vendedores.t_apellidos.Text = ""
f_vendedores.t_direccion.Text = ""
f_vendedores.t_telefono.Text = ""
f_vendedores.t_xcentajecomision.Text = ""
End Sub
