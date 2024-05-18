Attribute VB_Name = "modulo_compras"
Global VarTipoFormula As Integer
Global MatrizCostoPrecio(50, 14)
Global VarGlobalImpuesto As Double
Global ContadorCompra As Integer
Public Sub HabilitaEdicionFlexGridProveedor()
'se habilita t_editar en la columna 0 y en la fila disponible para que el cliente introduzca el codigo del articulo
        f_compras.fg_detallefactura.Row = 1
        f_compras.fg_detallefactura.Col = 0
        If f_compras.fg_detallefactura.TextMatrix(f_compras.fg_detallefactura.Row, 0) <> "" Then
            Do While f_compras.fg_detallefactura.TextMatrix(f_compras.fg_detallefactura.Row, 0) <> ""
                f_compras.fg_detallefactura.Row = f_compras.fg_detallefactura.Row + 1
            Loop
        End If
        f_compras.t_editar.Left = f_compras.fg_detallefactura.CellLeft + f_compras.fg_detallefactura.Left
        f_compras.t_editar.Top = f_compras.fg_detallefactura.CellTop + f_compras.fg_detallefactura.Top
        f_compras.t_editar.Width = f_compras.fg_detallefactura.CellWidth
        f_compras.t_editar.Height = f_compras.fg_detallefactura.CellHeight
        f_compras.t_editar.BorderStyle = 0
        f_compras.t_editar.FontName = f_compras.fg_detallefactura.CellFontName
        f_compras.t_editar.FontSize = f_compras.fg_detallefactura.CellFontSize
        f_compras.t_editar.FontBold = True
        f_compras.t_editar.Visible = True ' se coloca visible el t_editar
        f_compras.t_editar.Text = f_compras.fg_detallefactura.TextMatrix(f_compras.fg_detallefactura.Row, f_compras.fg_detallefactura.Col)
        f_compras.t_editar.SetFocus ' t_editar recibe el enfoque
End Sub

Public Sub ActivaEdicionItemProveedor()
        
        f_totalfactura.fg_tipopago.Col = 3
        f_totalfactura.t_detalle.Text = ""
        f_totalfactura.t_detalle.Left = f_totalfactura.fg_tipopago.CellLeft + f_totalfactura.fg_tipopago.Left
        f_totalfactura.t_detalle.Top = f_totalfactura.fg_tipopago.CellTop + f_totalfactura.fg_tipopago.Top
        f_totalfactura.t_detalle.Width = f_totalfactura.fg_tipopago.CellWidth
        f_totalfactura.t_detalle.BorderStyle = 0
        f_totalfactura.t_detalle.FontName = f_totalfactura.fg_tipopago.CellFontName
        f_totalfactura.t_detalle.Visible = True
        
        f_totalfactura.fg_tipopago.Col = 2
        f_totalfactura.p_banco.Left = f_totalfactura.fg_tipopago.CellLeft + f_totalfactura.fg_tipopago.Left
        f_totalfactura.p_banco.Top = f_totalfactura.fg_tipopago.CellTop + f_totalfactura.fg_tipopago.Top
        f_totalfactura.p_banco.Width = f_totalfactura.fg_tipopago.CellWidth
        f_totalfactura.c_banco.Width = f_totalfactura.p_banco.Width + 30
        f_totalfactura.p_banco.FontName = f_totalfactura.fg_tipopago.CellFontName
        f_totalfactura.c_banco.ListIndex = 0
        f_totalfactura.p_banco.Visible = True
        
        f_totalfactura.fg_tipopago.Col = 1
        f_totalfactura.p_tipopago.Left = f_totalfactura.fg_tipopago.CellLeft + f_totalfactura.fg_tipopago.Left
        f_totalfactura.p_tipopago.Top = f_totalfactura.fg_tipopago.CellTop + f_totalfactura.fg_tipopago.Top
        f_totalfactura.p_tipopago.Width = f_totalfactura.fg_tipopago.CellWidth
        f_totalfactura.c_tipopago.Width = f_totalfactura.p_tipopago.Width + 30
        f_totalfactura.p_tipopago.FontName = f_totalfactura.fg_tipopago.CellFontName
        f_totalfactura.c_tipopago.ListIndex = 0
        f_totalfactura.p_tipopago.Visible = True
        
        f_totalfactura.fg_tipopago.Col = 0
        f_totalfactura.t_monto.Text = ""
        f_totalfactura.t_monto.Left = f_totalfactura.fg_tipopago.CellLeft + f_totalfactura.fg_tipopago.Left
        f_totalfactura.t_monto.Top = f_totalfactura.fg_tipopago.CellTop + f_totalfactura.fg_tipopago.Top
        f_totalfactura.t_monto.Width = f_totalfactura.fg_tipopago.CellWidth
        f_totalfactura.t_monto.BorderStyle = 0
        f_totalfactura.t_monto.FontName = f_totalfactura.fg_tipopago.CellFontName
        f_totalfactura.t_monto.Visible = True
        f_totalfactura.t_monto.SetFocus
        
        'validacion para evitar error con los combobox al agregar nuevo item ya que estos solo aceptan valores que esten en su lista
        If f_totalfactura.fg_tipopago.TextMatrix(f_totalfactura.fg_tipopago.Row, 1) <> "" Then
            f_totalfactura.t_monto.Text = f_totalfactura.fg_tipopago.TextMatrix(f_totalfactura.fg_tipopago.Row, 0)
            f_totalfactura.c_tipopago.Text = f_totalfactura.fg_tipopago.TextMatrix(f_totalfactura.fg_tipopago.Row, 1)
            f_totalfactura.c_banco.Text = f_totalfactura.fg_tipopago.TextMatrix(f_totalfactura.fg_tipopago.Row, 2)
            f_totalfactura.t_detalle.Text = f_totalfactura.fg_tipopago.TextMatrix(f_totalfactura.fg_tipopago.Row, 3)
        End If
End Sub



Public Sub GenerarOrdenProveedor()

'##### AQUI INICIA EL CODIGO PARA ALMACENAR LOS DATOS DEL DOCUMENTO EN LA TABLA cabecera_doc #####
 Dim TempCabDoc As ADODB.Recordset
        '****codigo para cargar el numero de orden siguiente****
        Call Conn_BDaiosoft
        Set TempCabDoc = New ADODB.Recordset
        TempCabDoc.Open "SELECT nroorden FROM cabecera_doc  ORDER BY nroorden", Conn_Mysqldb, adOpenDynamic, adLockBatchOptimistic
        If TempCabDoc.BOF = False And TempCabDoc.EOF = False Then
            TempCabDoc.MoveLast
            numorden = TempCabDoc.Fields(0).Value + 1
            f_posrest.l_numorden.Caption = numorden
        Else
            f_posrest.l_numorden.Caption = 1
        End If
        '****fin de codigo para cargar el numero de orden siguiente****

            
            'consulta para la insercion de datos en la tabla cabecera_doc
            Conn_Mysqldb.Execute "INSERT INTO cabecera_doc SET serie= '0000'," _
            & "nroorden = " & f_posrest.l_numorden.Caption & "," _
            & "nrodoc = " & f_posrest.l_numorden.Caption & "," _
            & "tipodoc = '00'," _
            & "tipooper = '0101'," _
            & "firmadig = '-'," _
            & "fechaem = '" & Format(Date, "yyyy-mm-dd") & "'," _
            & "horaem = '" & Format(Time, "HH:MM:SS") & "'," _
            & "fechavenci = '" & Format(Date, "yyyy-mm-dd") & "'," _
            & "codlocalemisor = '0000'," _
            & "tipdocusuario = '" & f_posrest.c_tipodoc.Text & "'," _
            & "numdocusuario = '" & f_posrest.t_nrodoc.Text & "'," _
            & "nombrers = '" & f_posrest.t_nombrers.Text & "'," _
            & "tipmoneda = 'PEN'," _
            & "sumtottributos = 0," _
            & "sumtotvalventa = 0," _
            & "sumprecioventa = 0," _
            & "sumdesctotal = 0," _
            & "sumotroscargos = 0," _
            & "sumtotalanticipos = 0," _
            & "sumimpventa = 0," _
            & "ublversionld = '2.1'," _
            & "customizationld = '2.0'," _
            & "idvendedor = '.', abono = 0, restante = 0," _
            & "impreso = 's', estado = 'abierta', codcliente= " & f_posrest.l_idcliente.Caption & ", idoperador=" & f_principal.l_idoperador.Caption & ",idcaja= " & f_principal.l_idcaja.Caption & ""
               
            
               
            '##### AQUI FINALIZA EL CODIGO PARA ALMACENAR LOS DATOS DEL DOCUMENTO EN LA TABLA cabecera_doc #####
End Sub

Public Sub RegistrarProveedor()

If f_posrest.t_nrodoc.Text <> "" And f_posrest.t_nombrers.Text <> "" Then

    Dim tempcliente As New ADODB.Recordset
    Set tempcliente = New ADODB.Recordset
    Call Conn_BDaiosoft
    
    tempcliente.Open "SELECT codigo FROM cliente", Conn_Mysqldb, adOpenDynamic, adLockBatchOptimistic
    If tempcliente.EOF = False And tempcliente.BOF = False Then
        tempcliente.MoveLast
        CodigoCliente = tempcliente.Fields(0).Value + 1
    Else
        CodigoCliente = 1
    End If
    Conn_Mysqldb.Execute "INSERT INTO cliente SET codigo = " & CodigoCliente & "," _
    & "tipodocumento = '" & f_posrest.c_tipodoc.Text & "'," _
    & "nrodocumento = '" & f_posrest.t_nrodoc.Text & "'," _
    & "nombrers = '" & f_posrest.t_nombrers.Text & "'," _
    & "direccion = '" & f_posrest.t_direccion.Text & "'," _
    & "telefono1 = '" & f_posrest.t_telefono.Text & "'," _
    & "telefono2 = '" & f_posrest.t_telefono.Text & "'," _
    & "email = '" & f_posrest.t_email.Text & "'"
        
    f_posrest.l_idcliente.Caption = CodigoCliente
    f_posrest.c_tipodoc.Locked = True
    f_posrest.t_nrodoc.Locked = True
    f_posrest.t_nombrers.Locked = True
    f_posrest.t_direccion.Locked = True
    f_posrest.t_telefono.Locked = True
    f_posrest.t_email.Locked = True
    f_posrest.t_codpro.SetFocus
    
    Call GenerarOrden
Else
    MsgBox "Los campos número de documento y nombre o razón social no pueden estar en blanco."
End If
End Sub

