Attribute VB_Name = "Modulo_servgarantia"
Global CerrarVentanaCliente As Boolean
Global MatrizDetReparacion() As String
Public Sub ImprimirServGarant80mm(TituloDoc, VarTipoOrden, Fechaem, horaem, _
        VarNumDoc, VarNrodoc, VarNomCliente, VarAspectoDet, VarFalla, VarFila, VarColum, VarTotalDoc, _
        VarNomOperador, VarIdCaja, VarTipoDisp, VarMarcaDisp, VarModelo, VarColor, VarImeiSerial, VarTelefono)
'############ INICIO DE CODIGO PARA IMPRIMIR LA ORDEN ##############
            Dim TempDatosEmpresa As ADODB.Recordset
            Dim TempLogo As PictureBox
            
            If FormularioActual = "fgarantia" Then
                Set TempLogo = f_garantia.p_logo
            End If
            If FormularioActual = "servicios" Then
                Set TempLogo = f_servicios.p_logo
            End If
            
            Call Conn_BDaiosoft
            Set TempDatosEmpresa = New ADODB.Recordset
            TempDatosEmpresa.Open "SELECT ruc, nombre,direccionppal,telefonocelular,logo FROM datos_empresa", Conn_Mysqldb, adOpenDynamic, adLockBatchOptimistic
            If TempDatosEmpresa.BOF = False And TempDatosEmpresa.EOF = False Then
                TempDatosEmpresa.MoveFirst
                VarRuc = TempDatosEmpresa.Fields(0).Value
                VarNombre = TempDatosEmpresa.Fields(1).Value
                VarDireccion = TempDatosEmpresa.Fields(2).Value
                VarTelefonoEmpresa = TempDatosEmpresa.Fields(3).Value
                VarLogo = TempDatosEmpresa.Fields(4).Value
            Else
                VarRuc = " RUC DEMO"
                VarNombre = "NOMBRE DEMO"
                VarDireccion = "DIRECCION DEMO"
                VarTelefonoEmpresa = "TELEFONO DEMO"
                VarLogo = ""
            End If
            ' 567 twips equivalen a un centimetro
            Printer.Copies = 1
            For z = 1 To 2
            AnchoPapel = 4073
            On Error Resume Next
            Printer.PaintPicture TempLogo.Picture, 800, 400, 2402, 851
            Printer.CurrentY = 1418
            Printer.Font = "Courier New"
            Printer.FontSize = 8
            Printer.FontBold = True
            
            Printer.CurrentX = AnchoPapel / 2 - Printer.TextWidth(VarNombre) / 2
            Printer.Print VarNombre
            Printer.CurrentX = AnchoPapel / 2 - Printer.TextWidth("RUC: " & VarRuc) / 2
            Printer.Print "RUC: "; VarRuc
            Printer.CurrentX = AnchoPapel / 2 - Printer.TextWidth(VarDireccion) / 2
            Printer.Print VarDireccion
            Printer.CurrentX = AnchoPapel / 2 - Printer.TextWidth("TLF: " & VarTelefono) / 2
            Printer.Print "TLF: "; VarTelefonoEmpresa
            Printer.Print
            Printer.CurrentX = AnchoPapel / 2 - Printer.TextWidth(TituloDoc) / 2
            Printer.Print TituloDoc
            Printer.Print
            Printer.CurrentX = 1
            'inicio codigo para ordenar el nro orden de derecha a izquierda
            largonumorden = Len(VarNrodoc)
            cantespacio = 9 - largonumorden
            VarSerie = "0000"
            'fin codigo para ordenar el nro orden de derecha a izquierda
            Printer.Print "FECHA: "; Fechaem; Spc(10); "NRO: "; Spc(espacioderecha - 1); VarSerie; "-"; VarNrodoc
            Printer.CurrentX = 1
            Printer.Print "HORA: "; horaem
            Printer.CurrentX = 1
            Printer.Print "OPERADOR: "; VarNomOperador
            Printer.CurrentX = 1
            Printer.Print "CAJA: "; VarIdCaja
            Printer.CurrentX = 1
            Printer.Print "CLIENTE: "; VarNomCliente
            Printer.CurrentX = 1
            Printer.Print "CEL-TEL: "; VarTelefono
            Printer.CurrentX = 1
            Printer.Print
            Printer.CurrentX = 1
            Printer.Print "********DATOS DEL DISPOSITIVO********"
            Printer.Print
            Printer.CurrentX = 1
            Printer.Print "TIPO DISP: "; VarTipoDisp; " MARCA: "; VarMarcaDisp
            Printer.CurrentX = 1
            Printer.Print "MODELO: "; VarModelo; " COLOR: "; VarColor
            Printer.CurrentX = 1
            Printer.Print "IMEI / SERIAL: "; VarImeiSerial
            Printer.CurrentX = 1
            Printer.Print
            Printer.CurrentX = 1
            Printer.Print "*******EVALUACION DEL DISPOSITIVO*******"
            Printer.Print
            Printer.CurrentX = 1
            Printer.Print "FALLA PRESENTADA: "; VarFalla
            Printer.CurrentX = 1
            
            If Len(VarAspectoDet) <= 31 Then
                Printer.Print "OBS. DISP: "; VarAspectoDet
            End If
            If Len(VarAspectoDet) > 31 Then
                FILA1 = Mid(VarAspectoDet, 1, 31)
                FILA2 = Mid(VarAspectoDet, 32, 74)
                Printer.Print "OBS. DISP: "; FILA1
                Printer.CurrentX = 1
                Printer.Print FILA2
            End If
            
            Printer.Print
            Printer.CurrentX = 1
            Printer.Print "*********PRESUPUESTO ESTIMADO*********"
            
            Printer.CurrentX = 1
            Printer.Print "___________________________________________"
            Printer.CurrentX = 1
            Printer.Print "DESCRIPCION"; Spc(11); "P.U"; Spc(5); "CNT"; Spc(2); "TOTAL S/"
            Printer.CurrentX = 1
            Printer.Print "-------------------------------------------"
            For X = 1 To VarFila
            
                If Len(MatrizDetReparacion(X, 1)) <= 42 Then
                    FILA1 = Mid(MatrizDetReparacion(X, 1), 1, 42)
                    Printer.CurrentX = 1
                    Printer.Print FILA1
                End If
                If Len(MatrizDetReparacion(X, 1)) > 42 And Len(MatrizDetReparacion(X, 1)) < 87 Then
                    FILA1 = Mid(MatrizDetReparacion(X, 1), 1, 42)
                    FILA2 = Mid(MatrizDetReparacion(X, 1), 44, 42)
                    Printer.CurrentX = 1
                    Printer.Print FILA1
                    Printer.CurrentX = 1
                    Printer.Print FILA2
                End If
                
                'Printer.CurrentX = 1
                'Printer.Print Mid(MatrizDetReparacion(x, 1), 1, 42)
                
                anchopu = Len(MatrizDetReparacion(X, 2))
                ubipu = 1 + (28 - anchopu) * 96
                Printer.CurrentX = ubipu
                
                'CONDICIONAL A SOLICITUD DE CLIENTE
                'PARA NO MOSTRAR LOS PRECIOS SI LA CATEGORIA
                'DEL ARTICULO ES SERVICIO O REP
                If MatrizDetReparacion(X, 5) = "SERVICIO" Or MatrizDetReparacion(X, 5) = "REPUESTO" Then
                    MatrizDetReparacion(X, 2) = ""
                    MatrizDetReparacion(X, 3) = ""
                    MatrizDetReparacion(X, 4) = ""
                End If
                'FIN DE CONDICIONAL
                
                
                Printer.Print MatrizDetReparacion(X, 2)
                
                anchocant = Len(MatrizDetReparacion(X, 3))
                ubicant = 1 + (32 - anchocant) * 96
                Printer.CurrentY = Printer.CurrentY - 170
                Printer.CurrentX = ubicant
                Printer.Print MatrizDetReparacion(X, 3)
                
                anchototal = Len(MatrizDetReparacion(X, 4))
                ubitotal = 1 + (42 - anchototal) * 96
                Printer.CurrentY = Printer.CurrentY - 170
                Printer.CurrentX = ubitotal
                Printer.Print MatrizDetReparacion(X, 4)
            
            Next X
            Printer.CurrentX = 1
            Printer.Print "-------------------------------------------"
        
                
            Printer.CurrentX = 1 + 96 * (42 - Len(VarTotalDoc) - Len("TOTAL PR. ESTIMADO $: "))
            Printer.Print "TOTAL PR. ESTIMADO $: "
            Printer.CurrentX = 1 + 96 * (42 - Len(VarTotalDoc))
            Printer.CurrentY = Printer.CurrentY - 170
            Printer.Print VarTotalDoc
            
            Printer.Print ""
            Printer.Print ""
            Printer.Print ""
            Printer.Print "IMPORTANTE"
            Printer.Print ""
            Printer.Print "Pasados 7 días luego del envío de presu-"
            Printer.Print "puesto, este se entenderá como aprobado"
            Printer.Print "procediendose a la reparación. Pasado "
            Printer.Print "treinta (30)días luego de la notificación "
            Printer.Print "de equipo reparado o no, y este no sea"
            Printer.Print "retirado por el dueño tendrá un recargo "
            Printer.Print "monetario conforme a la ley, ademas por"
            Printer.Print "concepto de resguardo y almacenaje cuyo"
            Printer.Print "costo es de 0.04 UF diarios. Pasados"
            Printer.Print "noventa(90) días desde la fecha de "
            Printer.Print "recepción del equipo sin que el dueño lo "
            Printer.Print "retire, se considerará en estado de aban-"
            Printer.Print "dono por lo que usted autoriza al servicio"
            Printer.Print "técnico a enajenarlo a objeto de recuperar"
            Printer.Print "el valor invertido en su reparación junto "
            Printer.Print "con el monto generado por concepto de "
            Printer.Print "resguardo y almacenaje. La garantía solo"
            Printer.Print "cubre la reparación realizada según "
            Printer.Print "diagnostico previo y esta es de treinta"
            Printer.Print "(30) días para repuestos originales y"
            Printer.Print "diez (10) días para repuestos alternativos"
            Printer.Print "la cual comienza a hacerse efetiva al "
            Printer.Print "momento de la notificación de equipo"
            Printer.Print "reparado."
            Printer.Print ""
            Printer.Print ""
            Printer.Print ""
            Printer.Print ""
            Printer.Print "          __________________________"
            Printer.Print "              FIRMA DEL CLIENTE"
            
            Printer.EndDoc
            Next z
            
            '**************** FIN DE CODIGO DE IMPRESION DE ORDEN***************

End Sub
