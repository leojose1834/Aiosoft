Attribute VB_Name = "modulo_empleado"
'Global e_usuario As String
'Global Bitacora(1, 3) As String
'Global ValorLoginEmpleado As String
Global ValorCedulaEmpleado As String
Global BanderaEmpleado As Integer
Public Sub LimpiarTextBoxEmpleado()
    mantenimiento_empleado.t_nrodoc.Text = ""
    mantenimiento_empleado.t_apellidos.Text = ""
    mantenimiento_empleado.t_nombres.Text = ""
    mantenimiento_empleado.t_direccion.Text = ""
    mantenimiento_empleado.t_telefono.Text = ""
    mantenimiento_empleado.c_perfil.ListIndex = 0
    mantenimiento_empleado.c_ciclopago.ListIndex = 0
    mantenimiento_empleado.c_tipodoc.ListIndex = 0
    mantenimiento_empleado.fg_operador.Rows = 1
End Sub

Public Sub GuardarCambiosEmpleado()
Dim TempEmpleado As ADODB.Recordset
' la Variable BanderaOperador es igual a 1 si el operador a presionado
'el boton nuevo, entonces se procede a crear un nuevo registro
If BanderaEmpleado = 1 Then
    Call Conn_BDaiosoft
    Set TempEmpleado = New ADODB.Recordset
    TempEmpleado.Open "SELECT idempleado FROM empleado WHERE tipodoc= '" & mantenimiento_empleado.c_tipodoc.Text & "' AND nrodoc= '" & mantenimiento_empleado.t_nrodoc.Text & "'", Conn_Mysqldb, adOpenDynamic, adLockBatchOptimistic
    If TempEmpleado.BOF = True And TempEmpleado.EOF = True Then
        TempEmpleado.Close
            TempEmpleado.Open "SELECT idempleado FROM empleado", Conn_Mysqldb, adOpenDynamic, adLockBatchOptimistic
            If TempEmpleado.BOF = False And TempEmpleado.EOF = False Then
                TempEmpleado.MoveLast
                VarIdEmpleado = TempEmpleado.Fields(0).Value + 1
            Else
                VarIdEmpleado = 1
            End If
            'alamcena los datos de operador
            Conn_Mysqldb.Execute "INSERT INTO empleado SET idempleado = " & VarIdEmpleado & "," _
            & "tipodoc = '" & mantenimiento_empleado.c_tipodoc.Text & "', nrodoc = '" & mantenimiento_empleado.t_nrodoc.Text & "'," _
            & "apellidos = '" & mantenimiento_empleado.t_apellidos.Text & "', nombres = '" & mantenimiento_empleado.t_nombres.Text & "'," _
            & "direccion = '" & mantenimiento_empleado.t_direccion.Text & "', telefono = '" & mantenimiento_empleado.t_telefono.Text & "'," _
            & "cargo = '" & mantenimiento_empleado.c_perfil.Text & "', ciclopago = '" & mantenimiento_empleado.c_ciclopago.Text & "'"
            
            
            'habilita y desabilita botones
            mantenimiento_empleado.b_nuevo.Enabled = True
            mantenimiento_empleado.b_cancelar.Enabled = False
            mantenimiento_empleado.b_guardar.Enabled = False
            
            
            
            'desabilita y habilita frames
            mantenimiento_empleado.fra_buscaroperador.Enabled = True
            mantenimiento_empleado.fra_datosoperador.Enabled = False
            mantenimiento_empleado.fra_permisosoperador.Enabled = False
            
            MsgBox "Los datos han sido almacenados exitosamente."
            
            ' limpia  las cajas de texto
            Call LimpiarTextBoxEmpleado
            
    Else
        MsgBox "El número de documento que ingresó ya  está registrado, por favor intente con otro"
    End If
End If

'la Variable BanderaOperador es igual a 2 si el operador a presionado
'el boton modificar, entonces se procede a editar el registro actual registro
If BanderaEmpleado = 2 Then

        Conn_Mysqldb.Execute "UPDATE empleado SET apellidos = '" & mantenimiento_empleado.t_apellidos.Text & "'," _
        & "nombres = '" & mantenimiento_empleado.t_nombres.Text & "'," _
        & "direccion = '" & mantenimiento_empleado.t_direccion.Text & "', telefono = '" & mantenimiento_empleado.t_telefono.Text & "'," _
        & "cargo = '" & mantenimiento_empleado.c_perfil.Text & "', ciclopago = '" & mantenimiento_empleado.c_ciclopago.Text & "' WHERE idempleado= '" & VarIdEmpleado & "'"
        
        'habilita y desabilita botones
        mantenimiento_empleado.b_nuevo.Enabled = True
        mantenimiento_empleado.b_cancelar.Enabled = False
        mantenimiento_empleado.b_guardar.Enabled = False
                
        
                
        'desabilita y habilita frames
        mantenimiento_empleado.fra_buscaroperador.Enabled = True
        mantenimiento_empleado.fra_datosoperador.Enabled = False
        mantenimiento_empleado.fra_permisosoperador.Enabled = False
                           
        MsgBox "Los datos han sido actualizados exitosamente."
        
        'limpia  las cajas de texto
        Call LimpiarTextBoxEmpleado
  
End If

'codigo para actualizar  los conceptos del empleado
Conn_Mysqldb.Execute "DELETE FROM  empleado_concepto WHERE idempleado='" & VarIdEmpleado & "'"

If mantenimiento_empleado.fg_concepto.Rows > 1 Then
    For X = 1 To mantenimiento_empleado.fg_concepto.Rows - 1
        Conn_Mysqldb.Execute "INSERT INTO empleado_concepto SET idempleado='" & VarIdEmpleado & "'," _
        & "idconcepto='" & mantenimiento_empleado.fg_concepto.TextMatrix(X, 0) & "'," _
        & "tipovalor='" & mantenimiento_empleado.fg_concepto.TextMatrix(X, 3) & "'," _
        & "valor='" & mantenimiento_empleado.fg_concepto.TextMatrix(X, 4) & "'"
    Next X
mantenimiento_empleado.fg_concepto.Rows = 1
End If

End Sub

