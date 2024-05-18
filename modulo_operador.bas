Attribute VB_Name = "modulo_operador"
Global e_usuario As String
Global Bitacora(1, 3) As String
Global ValorLoginOperador As String
Global ValorCedulaOperador As String
Global BanderaOperador As Integer
Public Sub GuardarCambiosOperador()
Dim TempOperador As ADODB.Recordset
' la Variable BanderaOperador es igual a 1 si el operador a presionado
'el boton nuevo, entonces se procede a crear un nuevo registro
If BanderaOperador = 1 Then
    Call Conn_BDaiosoft
    Set TempOperador = New ADODB.Recordset
    TempOperador.Open "SELECT idoperador FROM operador WHERE tipodoc= '" & mantenimiento_usuarios.c_tipodoc.Text & "' AND nrodoc= '" & mantenimiento_usuarios.t_nrodoc.Text & "'", Conn_Mysqldb, adOpenDynamic, adLockBatchOptimistic
    If TempOperador.BOF = True And TempOperador.EOF = True Then
        TempOperador.Close
        TempOperador.Open "SELECT idoperador FROM operador WHERE login= '" & mantenimiento_usuarios.t_login & "'", Conn_Mysqldb, adOpenDynamic, adLockBatchOptimistic
        If TempOperador.BOF = True And TempOperador.EOF = True Then
            TempOperador.Close
            TempOperador.Open "SELECT idoperador FROM operador", Conn_Mysqldb, adOpenDynamic, adLockBatchOptimistic
            If TempOperador.BOF = False And TempOperador.EOF = False Then
                TempOperador.MoveLast
                VarIdOperador = TempOperador.Fields(0).Value + 1
            Else
                VarIdOperador = 1
            End If
            'alamcena los datos de operador
            Conn_Mysqldb.Execute "INSERT INTO operador SET idoperador = " & VarIdOperador & "," _
            & "tipodoc = '" & mantenimiento_usuarios.c_tipodoc.Text & "', nrodoc = '" & mantenimiento_usuarios.t_nrodoc.Text & "'," _
            & "apellidos = '" & mantenimiento_usuarios.t_apellidos.Text & "', nombres = '" & mantenimiento_usuarios.t_nombres.Text & "'," _
            & "direccion = '" & mantenimiento_usuarios.t_direccion.Text & "', telefono = '" & mantenimiento_usuarios.t_telefono.Text & "'," _
            & "perfil = '" & mantenimiento_usuarios.c_perfil.Text & "', login = '" & mantenimiento_usuarios.t_login.Text & "'," _
            & "password = '" & mantenimiento_usuarios.t_password & "'," _
            & "mnmantenimiento= '" & mantenimiento_usuarios.l_permisos.Selected(0) & "', subclientes= '" & mantenimiento_usuarios.l_permisos.Selected(1) & "'," _
            & "subproveedores= '" & mantenimiento_usuarios.l_permisos.Selected(2) & "', subinventario= '" & mantenimiento_usuarios.l_permisos.Selected(3) & "'," _
            & "subalmacen= '" & mantenimiento_usuarios.l_permisos.Selected(4) & "', mnventas= '" & mantenimiento_usuarios.l_permisos.Selected(5) & "'," _
            & "subfacturar= '" & mantenimiento_usuarios.l_permisos.Selected(6) & "', subpos= '" & mantenimiento_usuarios.l_permisos.Selected(7) & "'," _
            & "subpresupuesto= '" & mantenimiento_usuarios.l_permisos.Selected(8) & "', subctaxcobrar= '" & mantenimiento_usuarios.l_permisos.Selected(9) & "'," _
            & "subnotacredito= '" & mantenimiento_usuarios.l_permisos.Selected(10) & "', subcombaja= '" & mantenimiento_usuarios.l_permisos.Selected(11) & "'," _
            & "subanulaciones= '" & mantenimiento_usuarios.l_permisos.Selected(12) & "', mncompras= '" & mantenimiento_usuarios.l_permisos.Selected(13) & "'," _
            & "subregcompras= '" & mantenimiento_usuarios.l_permisos.Selected(14) & "', subordencompra= '" & mantenimiento_usuarios.l_permisos.Selected(15) & "'," _
            & "subctaporpagar= '" & mantenimiento_usuarios.l_permisos.Selected(16) & "', mnreportes= '" & mantenimiento_usuarios.l_permisos.Selected(17) & "'," _
            & "subventas= '" & mantenimiento_usuarios.l_permisos.Selected(18) & "', subcompras= '" & mantenimiento_usuarios.l_permisos.Selected(19) & "'," _
            & "subrepinventario= '" & mantenimiento_usuarios.l_permisos.Selected(20) & "', subrepctaxpagar= '" & mantenimiento_usuarios.l_permisos.Selected(21) & "'," _
            & "subrepctaxcobrar= '" & mantenimiento_usuarios.l_permisos.Selected(22) & "', mnajustes= '" & mantenimiento_usuarios.l_permisos.Selected(23) & "'," _
            & "subdatemp= '" & mantenimiento_usuarios.l_permisos.Selected(24) & "', subuser= '" & mantenimiento_usuarios.l_permisos.Selected(25) & "'," _
            & "mnsoporte= '" & mantenimiento_usuarios.l_permisos.Selected(26) & "', subservicios= '" & mantenimiento_usuarios.l_permisos.Selected(27) & "'," _
            & "subgarantias= '" & mantenimiento_usuarios.l_permisos.Selected(28) & "'"
          
            
            'habilita y desabilita botones
            mantenimiento_usuarios.b_nuevo.Enabled = True
            mantenimiento_usuarios.b_cancelar.Enabled = False
            mantenimiento_usuarios.b_guardar.Enabled = False
            
            ' limpia  las cajas de texto
            Call LimpiarTextBoxOperador
            
            'desabilita y habilita frames
            mantenimiento_usuarios.fra_buscaroperador.Enabled = True
            mantenimiento_usuarios.fra_datosoperador.Enabled = False
            mantenimiento_usuarios.fra_permisosoperador.Enabled = False
            
            'quita los cheks seleccionados en la lista permisos
            Call QuitarCheksEnListaPermisos
            
            MsgBox "Los datos han sido almacenados satisfacoriamente"
        Else
            MsgBox "El login que ingresó ya  está registrado, por favor intente con otro"
        End If
    Else
        MsgBox "El número de documento que ingresó ya  está registrado, por favor intente con otro"
    End If
End If

'la Variable BanderaOperador es igual a 2 si el operador a presionado
'el boton modificar, entonces se procede a editar el registro actual registro
If BanderaOperador = 2 Then

        Conn_Mysqldb.Execute "UPDATE operador SET apellidos = '" & mantenimiento_usuarios.t_apellidos.Text & "'," _
        & "nombres = '" & mantenimiento_usuarios.t_nombres.Text & "'," _
        & "direccion = '" & mantenimiento_usuarios.t_direccion.Text & "', telefono = '" & mantenimiento_usuarios.t_telefono.Text & "'," _
        & "perfil = '" & mantenimiento_usuarios.c_perfil.Text & "'," _
        & "password = '" & mantenimiento_usuarios.t_password & "'," _
        & "mnmantenimiento= '" & mantenimiento_usuarios.l_permisos.Selected(0) & "', subclientes= '" & mantenimiento_usuarios.l_permisos.Selected(1) & "'," _
        & "subproveedores= '" & mantenimiento_usuarios.l_permisos.Selected(2) & "', subinventario= '" & mantenimiento_usuarios.l_permisos.Selected(3) & "'," _
        & "subalmacen= '" & mantenimiento_usuarios.l_permisos.Selected(4) & "', mnventas= '" & mantenimiento_usuarios.l_permisos.Selected(5) & "'," _
        & "subfacturar= '" & mantenimiento_usuarios.l_permisos.Selected(6) & "', subpos= '" & mantenimiento_usuarios.l_permisos.Selected(7) & "'," _
        & "subpresupuesto= '" & mantenimiento_usuarios.l_permisos.Selected(8) & "', subctaxcobrar= '" & mantenimiento_usuarios.l_permisos.Selected(9) & "'," _
        & "subnotacredito= '" & mantenimiento_usuarios.l_permisos.Selected(10) & "', subcombaja= '" & mantenimiento_usuarios.l_permisos.Selected(11) & "'," _
        & "subanulaciones= '" & mantenimiento_usuarios.l_permisos.Selected(12) & "', mncompras= '" & mantenimiento_usuarios.l_permisos.Selected(13) & "'," _
        & "subregcompras= '" & mantenimiento_usuarios.l_permisos.Selected(14) & "', subordencompra= '" & mantenimiento_usuarios.l_permisos.Selected(15) & "'," _
        & "subctaporpagar= '" & mantenimiento_usuarios.l_permisos.Selected(16) & "', mnreportes= '" & mantenimiento_usuarios.l_permisos.Selected(17) & "'," _
        & "subventas= '" & mantenimiento_usuarios.l_permisos.Selected(18) & "', subcompras= '" & mantenimiento_usuarios.l_permisos.Selected(19) & "'," _
        & "subrepinventario= '" & mantenimiento_usuarios.l_permisos.Selected(20) & "', subrepctaxpagar= '" & mantenimiento_usuarios.l_permisos.Selected(21) & "'," _
        & "subrepctaxcobrar= '" & mantenimiento_usuarios.l_permisos.Selected(22) & "', mnajustes= '" & mantenimiento_usuarios.l_permisos.Selected(23) & "'," _
        & "subdatemp= '" & mantenimiento_usuarios.l_permisos.Selected(24) & "', subuser= '" & mantenimiento_usuarios.l_permisos.Selected(25) & "'," _
        & "mnsoporte= '" & mantenimiento_usuarios.l_permisos.Selected(26) & "', subservicios= '" & mantenimiento_usuarios.l_permisos.Selected(27) & "'," _
        & "subgarantias= '" & mantenimiento_usuarios.l_permisos.Selected(28) & "' WHERE tipodoc= '" & mantenimiento_usuarios.c_tipodoc.Text & "' AND nrodoc= '" & mantenimiento_usuarios.t_nrodoc.Text & "'"
            
        
        'habilita y desabilita botones
        mantenimiento_usuarios.b_nuevo.Enabled = True
        mantenimiento_usuarios.b_cancelar.Enabled = False
        mantenimiento_usuarios.b_guardar.Enabled = False
                
        'limpia  las cajas de texto
        Call LimpiarTextBoxOperador
                
        'desabilita y habilita frames
        mantenimiento_usuarios.fra_buscaroperador.Enabled = True
        mantenimiento_usuarios.fra_datosoperador.Enabled = False
        mantenimiento_usuarios.fra_permisosoperador.Enabled = False
                
        'quita los cheks seleccionados en la lista permisos
        Call QuitarCheksEnListaPermisos
                
        MsgBox "Los datos han sido almacenados satisfacoriamente"
  
End If

End Sub

Public Sub LimpiarTextBoxOperador()
mantenimiento_usuarios.t_nrodoc.Text = ""
mantenimiento_usuarios.t_apellidos.Text = ""
mantenimiento_usuarios.t_nombres.Text = ""
mantenimiento_usuarios.t_direccion.Text = ""
mantenimiento_usuarios.t_telefono.Text = ""
mantenimiento_usuarios.t_login.Text = ""
mantenimiento_usuarios.t_password.Text = ""
mantenimiento_usuarios.c_perfil.ListIndex = 0

End Sub
Public Sub QuitarCheksEnListaPermisos()
For x = 0 To (mantenimiento_usuarios.l_permisos.ListCount - 1)
    mantenimiento_usuarios.l_permisos.Selected(x) = False
Next x
End Sub
