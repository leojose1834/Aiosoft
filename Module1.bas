Attribute VB_Name = "Modulo_producto"
Global BanderaProducto As Boolean
Global NoEstaElemento As Boolean
Global contador As Integer
Global Contador_Texistencia As Integer
Global CodigoImpuesto As String
Global TasaImpuesto As String
Global cambio_foco_modificar As Boolean
Global auxunidadcaja As String
Global auxexistencia As String
Global auxexistenciaunidades As String
Global auxcodigo As String
Global auxdescripcion As String
Global ValidadoCodigoDescripcion As Boolean

Public Sub HabilitarTextbox() 'procedimiento para habilitar los  textbox
mantenimiento_producto.t_codigo.Enabled = True
mantenimiento_producto.t_descripcion.Enabled = True
mantenimiento_producto.t_existencia.Enabled = True
mantenimiento_producto.t_cantmax.Enabled = True
mantenimiento_producto.t_cantmin.Enabled = True
mantenimiento_producto.c_categoria.Enabled = True
mantenimiento_producto.t_precio.Enabled = True
mantenimiento_producto.t_codbarra.Enabled = True
mantenimiento_producto.c_impuesto.Enabled = True
mantenimiento_producto.dt_fvencimiento.Enabled = True
mantenimiento_producto.t_costo.Enabled = True
mantenimiento_producto.t_porcentaje.Enabled = True
mantenimiento_producto.OB_simple.Enabled = True
mantenimiento_producto.OB_estandar.Enabled = True
mantenimiento_producto.c_tributoisc.Enabled = True
mantenimiento_producto.c_otrotributo.Enabled = True
mantenimiento_producto.c_tributoicbper.Enabled = True
mantenimiento_producto.t_preciosinigv.Enabled = True
mantenimiento_producto.fr_tipo.Enabled = True
End Sub

Public Sub DesabilitarTextbox() 'procedimento para desabilitar los textbox
mantenimiento_producto.t_codigo.Enabled = False
mantenimiento_producto.t_descripcion.Enabled = False
mantenimiento_producto.t_existencia.Enabled = False
mantenimiento_producto.t_cantmax.Enabled = False
mantenimiento_producto.t_cantmin.Enabled = False
mantenimiento_producto.c_categoria.Enabled = False
mantenimiento_producto.t_precio.Enabled = False
mantenimiento_producto.t_codbarra.Enabled = False
mantenimiento_producto.c_impuesto.Enabled = False
mantenimiento_producto.dt_fvencimiento.Enabled = False
mantenimiento_producto.t_costo.Enabled = False
mantenimiento_producto.t_porcentaje.Enabled = False
mantenimiento_producto.OB_simple.Enabled = False
mantenimiento_producto.OB_estandar.Enabled = False
mantenimiento_producto.c_tributoisc.Enabled = False
mantenimiento_producto.c_otrotributo.Enabled = False
mantenimiento_producto.c_tributoicbper.Enabled = False
mantenimiento_producto.t_preciosinigv.Enabled = False
mantenimiento_producto.fr_tipo.Enabled = False
End Sub



Public Sub GuardarCambios()
Dim VarConBDAiosoft As Object

'If mantenimiento_producto.c_almacen.Text = "ACCESORIOS" Then
    Set VarConBDAiosoft = Conn_Mysqldb
'End If
'If mantenimiento_producto.c_almacen.Text = "REPUESTOS" Then
'    Set VarConBDAiosoft = Conn_MysqldbAlmRep
'End If

'procedimiento para guardar los cambios realizados en el registro actual
ValidadoCodigoDescripcion = False

Dim VarDia As String
Dim VarMes As String
Dim VarAnio As String
Dim VarFecha As String
Dim temproducto As ADODB.Recordset

VarDia = mantenimiento_producto.dt_fvencimiento.Day
VarMes = mantenimiento_producto.dt_fvencimiento.Month
VarAnio = mantenimiento_producto.dt_fvencimiento.Year
VarFecha = VarAnio + "-" + VarMes + "-" + VarDia


If mantenimiento_producto.op_producto.Value = True Then
    VarTipoPS = "p"
Else
    VarTipoPS = "s"
End If

' si bandera es verdadero signifia  que el  usuario preciono
'el boton modificar, se procede entonces a editar los datos
If BanderaProducto Then

    If mantenimiento_producto.OB_simple.Value = True Then
        VarTipoFormula = 1
    Else
        VarTipoFormula = 2
    End If
    
    If mantenimiento_producto.op_servicio.Value = True Then
        mantenimiento_producto.t_cantmax.Text = 0
        mantenimiento_producto.t_cantmin.Text = 0
        mantenimiento_producto.t_existencia.Text = 0
    End If
    
    If mantenimiento_producto.c_diasgarantia.Text = "" Then
        mantenimiento_producto.c_diasgarantia.Text = "0"
    End If
    
        VarConBDAiosoft.Execute "UPDATE producto SET codigo = '" & mantenimiento_producto.t_codigo.Text & "'," _
        & "categoria = '" & mantenimiento_producto.c_categoria.Text & "'," _
        & "descripcion = '" & mantenimiento_producto.t_descripcion.Text & "'," _
        & "cantidadmax = " & mantenimiento_producto.t_cantmax.Text & "," _
        & "cantidadmin = " & mantenimiento_producto.t_cantmin.Text & "," _
        & "existencia = " & mantenimiento_producto.t_existencia.Text & "," _
        & "impuestogeneral = '" & mantenimiento_producto.c_impuesto & "'," _
        & "impuestoisc = '" & mantenimiento_producto.c_tributoisc & "'," _
        & "otrosimpuestos = '" & mantenimiento_producto.c_otrotributo.Text & "'," _
        & "impuestoicbper = '" & mantenimiento_producto.c_tributoicbper.Text & "'," _
        & "precio = " & VarPrecioConIgv & "," _
        & "costo = " & VarCostoProducto & "," _
        & "codigobarra = '" & mantenimiento_producto.t_codbarra.Text & "'," _
        & "xcentajebeneficio = " & VarxcentajeBeneficio & "," _
        & "tipoformula = " & VarTipoFormula & "," _
        & "fvencimiento = '" & Format(VarFecha, "yyyy-mm-dd") & "'," _
        & "preciosinigv = " & VarPrecioSinIgv & "," _
        & "unidmed = '" & mantenimiento_producto.c_unidmed.Text & "'," _
        & "precio2 = " & VarPrecioConIgv2 & ", porcentaje2 = " & VarxcentajeBeneficio2 & "," _
        & "preciosinigv2 = " & VarPrecioSinIgv2 & "," _
        & "precio3 = " & VarPrecioConIgv3 & ", porcentaje3 = " & VarxcentajeBeneficio3 & "," _
        & "preciosinigv3 = " & VarPrecioSinIgv3 & "," _
        & "precio4 = " & VarPrecioConIgv4 & ", porcentaje4 = " & VarxcentajeBeneficio4 & "," _
        & "preciosinigv4 = " & VarPrecioSinIgv4 & "," _
        & "tipo = '" & VarTipoPS & "', diasgarantia='" & mantenimiento_producto.c_diasgarantia.Text & "', marca='" & mantenimiento_producto.t_marca.Text & "', modelo='" & mantenimiento_producto.t_modelo.Text & "', color='" & mantenimiento_producto.t_color.Text & "', idalmacen=1 WHERE codigo = " & mantenimiento_producto.t_codigo.Text & ""
        MsgBox "La información fue actualizada exitosamente"
        ValidadoCodigoDescripcion = True
Else

    ' si el valor de bandera es falso quiere decir que el usuario preciono
    ' el boton nuevo, se procede entonces a añadir los nuevos datos
    
    
    If mantenimiento_producto.OB_simple.Value = True Then
        VarTipoFormula = 1
    Else
        VarTipoFormula = 2
    End If
    
    If mantenimiento_producto.c_diasgarantia.Text = "" Then
        mantenimiento_producto.c_diasgarantia.Text = 0
    End If
    
    
    'validacion para no ingresar codigo de barra repetido
    Set temproducto = New ADODB.Recordset
    temproducto.Open "SELECT codigobarra FROM producto WHERE codigobarra= '" & mantenimiento_producto.t_codbarra.Text & "'", Conn_Mysqldb, adOpenDynamic, adLockBatchOptimistic
    If temproducto.BOF = True And temproducto.EOF = True Then

        
        VarConBDAiosoft.Execute "INSERT INTO producto SET codigo = '" & mantenimiento_producto.t_codigo.Text & "'," _
        & "categoria = '" & mantenimiento_producto.c_categoria.Text & "'," _
        & "descripcion = '" & mantenimiento_producto.t_descripcion.Text & "'," _
        & "cantidadmax = " & mantenimiento_producto.t_cantmax.Text & "," _
        & "cantidadmin = " & mantenimiento_producto.t_cantmin.Text & "," _
        & "existencia = " & mantenimiento_producto.t_existencia.Text & "," _
        & "impuestogeneral = '" & mantenimiento_producto.c_impuesto.Text & "'," _
        & "impuestoisc = '" & mantenimiento_producto.c_tributoisc.Text & "'," _
        & "otrosimpuestos = '" & mantenimiento_producto.c_otrotributo.Text & "'," _
        & "impuestoicbper = '" & mantenimiento_producto.c_tributoicbper.Text & "'," _
        & "precio = " & VarPrecioConIgv & "," _
        & "costo = " & VarCostoProducto & "," _
        & "codigobarra = '" & mantenimiento_producto.t_codbarra.Text & "'," _
        & "xcentajebeneficio = " & VarxcentajeBeneficio & "," _
        & "tipoformula =  " & VarTipoFormula & "," _
        & "fvencimiento = '" & Format(VarFecha, "yyyy-mm-dd") & "'," _
        & "preciosinigv = " & VarPrecioSinIgv & "," _
        & "unidmed = '" & mantenimiento_producto.c_unidmed.Text & "'," _
        & "precio2 = " & VarPrecioConIgv2 & ", porcentaje2 = " & VarxcentajeBeneficio2 & "," _
        & "preciosinigv2 = " & VarPrecioSinIgv2 & "," _
        & "precio3 = " & VarPrecioConIgv3 & ", porcentaje3 = " & VarxcentajeBeneficio3 & "," _
        & "preciosinigv3 = " & VarPrecioSinIgv3 & "," _
        & "precio4 = " & VarPrecioConIgv4 & ", porcentaje4 = " & VarxcentajeBeneficio4 & "," _
        & "preciosinigv4 = " & VarPrecioSinIgv4 & "," _
        & "tipo = '" & VarTipoPS & "', diasgarantia='" & mantenimiento_producto.c_diasgarantia.Text & "', marca='" & mantenimiento_producto.t_marca.Text & "', modelo='" & mantenimiento_producto.t_modelo.Text & "', color='" & mantenimiento_producto.t_color.Text & "', idalmacen=1"
        
        Contador_Texistencia = 0
        MsgBox "El nuevo producto ha sido agregado satisfactoriamente"
        ValidadoCodigoDescripcion = True
        contador = 0
    Else
        MsgBox "El código de barras ya se encuentra registrado."
        ValidadoCodigoDescripcion = False
    End If


End If
End Sub

Public Sub LimpiarTextbox()
mantenimiento_producto.t_codigo.Text = ""
mantenimiento_producto.t_descripcion.Text = ""
mantenimiento_producto.t_existencia.Text = ""
mantenimiento_producto.t_cantmax.Text = ""
mantenimiento_producto.t_cantmin.Text = ""
mantenimiento_producto.c_categoria.Text = ""
mantenimiento_producto.t_codbarra.Text = ""
mantenimiento_producto.t_costo.Text = ""
mantenimiento_producto.t_precio.Text = ""
mantenimiento_producto.t_porcentaje.Text = ""
mantenimiento_producto.t_preciosinigv.Text = ""

mantenimiento_producto.t_precio2.Text = ""
mantenimiento_producto.t_porcentaje2.Text = ""
mantenimiento_producto.t_preciosinigv2.Text = ""

mantenimiento_producto.t_precio3.Text = ""
mantenimiento_producto.t_porcentaje3.Text = ""
mantenimiento_producto.t_preciosinigv3.Text = ""

mantenimiento_producto.t_precio4.Text = ""
mantenimiento_producto.t_porcentaje4.Text = ""
mantenimiento_producto.t_preciosinigv4.Text = ""

mantenimiento_producto.dt_fvencimiento = Date
mantenimiento_producto.c_tributoisc.ListIndex = 0
mantenimiento_producto.c_otrotributo.ListIndex = 0
mantenimiento_producto.c_tributoicbper.ListIndex = 0
mantenimiento_producto.c_almacen.ListIndex = 0
mantenimiento_producto.c_impuesto.ListIndex = 2
mantenimiento_producto.c_unidmed.ListIndex = 0

mantenimiento_producto.c_diasgarantia.Text = ""
mantenimiento_producto.t_marca.Text = ""
mantenimiento_producto.t_modelo.Text = ""
mantenimiento_producto.t_color.Text = ""

End Sub


Public Sub CargarComboCategoria() 'procedimiento para cargar el combo categoria y eliminar elementos repetidos

Dim temproducto As New ADODB.Recordset

Call Conn_BDaiosoft
Set temproducto = New ADODB.Recordset
temproducto.Open "SELECT DISTINCT categoria FROM producto", Conn_Mysqldb, adOpenDynamic, adLockBatchOptimistic

If temproducto.BOF = False And temproducto.EOF = False Then
    temproducto.MoveFirst
    Do While Not temproducto.EOF
    
        mantenimiento_producto.c_categoria.AddItem temproducto.Fields(0).Value
        temproducto.MoveNext
    
    Loop

End If
        
End Sub




Public Sub CargarTablaImpuesto()
Dim tempimpuesto As New ADODB.Recordset

Call Conn_BDaiosoft
Set tempimpuesto = New ADODB.Recordset
    tempimpuesto.Open "SELECT * FROM impuesto", Conn_Mysqldb, adOpenDynamic, adLockBatchOptimistic

        
'condicional para verificar q la tabla no este vacia
If Not tempimpuesto.RecordCount = 0 Then
            
    'se mueve tempimpuesto al primer registro
    tempimpuesto.MoveFirst
            
    'se limpia el flexgrid impuesto
    mantenimiento_producto.fg_impuesto.Rows = 1
    
    'se da formato de encabezado al flexgrid impuesto
    mantenimiento_producto.fg_impuesto.FormatString = "Código       |Denominación          |% Tasa  "
            
    'mientras no sea fin de archivo tempimpuesto se procede a llenar
    'el flexgrid impuesto
    Do While Not tempimpuesto.EOF
        mantenimiento_producto.fg_impuesto.AddItem tempimpuesto.Fields(0).Value
        mantenimiento_producto.fg_impuesto.TextMatrix(mantenimiento_producto.fg_impuesto.Rows - 1, 1) = tempimpuesto.Fields(1).Value
        mantenimiento_producto.fg_impuesto.TextMatrix(mantenimiento_producto.fg_impuesto.Rows - 1, 2) = tempimpuesto.Fields(2).Value
        
                
        tempimpuesto.MoveNext
    Loop
        
Else
    mantenimiento_producto.fg_impuesto.Clear
    mantenimiento_producto.fg_impuesto.Rows = 1
    mantenimiento_producto.fg_impuesto.FormatString = "Código               |Denominación              |% Tasa      "
End If
End Sub

Public Sub CargarComboImpuesto()

Dim tempimpuesto As New ADODB.Recordset

Call Conn_BDaiosoft
Set tempimpuesto = New ADODB.Recordset
    tempimpuesto.Open "SELECT * FROM impuesto", Conn_Mysqldb, adOpenDynamic, adLockBatchOptimistic


If tempimpuesto.RecordCount = -1 Then
    tempimpuesto.MoveFirst
    Do While Not tempimpuesto.EOF
        AuxCodigoImpuesto = tempimpuesto.Fields(0).Value
        'AuxTasaImpuesto = mantenimiento_producto.impuesto.Recordset("tasa")
        'mantenimiento_producto.c_impuesto.AddItem AuxCodigoImpuesto + "/" + AuxTasaImpuesto + "%"
       
        If AuxCodigoImpuesto <> "9999" And AuxCodigoImpuesto <> "7152" And AuxCodigoImpuesto <> "2000" Then
            mantenimiento_producto.c_impuesto.AddItem AuxCodigoImpuesto
        End If
        If AuxCodigoImpuesto = "2000" Or AuxCodigoImpuesto = "01" Then
            mantenimiento_producto.c_tributoisc.AddItem AuxCodigoImpuesto
        End If
        If AuxCodigoImpuesto = "9999" Or AuxCodigoImpuesto = "01" Then
            mantenimiento_producto.c_otrotributo.AddItem AuxCodigoImpuesto
        End If
        If AuxCodigoImpuesto = "7152" Or AuxCodigoImpuesto = "01" Then
            mantenimiento_producto.c_tributoicbper.AddItem AuxCodigoImpuesto
        End If
        tempimpuesto.MoveNext
    Loop
End If
End Sub
