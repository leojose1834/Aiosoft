Attribute VB_Name = "modulo_facturacion"
Global MatrizDetDoc() As String
Global MaTrizSeriales(1 To 50, 1 To 2) As String
Global ValorFilaMaTrizSeriales As Integer
Global Decena(1 To 10) As String
Global Unidad(9) As String
Global DecenaEsp(0 To 6) As String
Global Centena(1 To 9) As String
Global DecenaEsp2(10 To 16) As String
Global Var_Anulacion As String
Global Clik_Aceptar As Boolean

Public Function MontoEnLetras(ValorSinDecimal)
cifra = ValorSinDecimal



'******CONVERSION DE 8 CIFRAS********
If Len(ValorSinDecimal) = 8 Then

    vardecenamillon = (Mid(ValorSinDecimal, 1, 2))
    varcentenamiles = (Mid(ValorSinDecimal, 3, 3))
    valorcentena = Mid(ValorSinDecimal, 6, 3)

   If varcentenamiles = "000" And valorcentena = "000" Then
    MontoEnLetras = Cifra2ATexto(vardecenamillon) + " " + "MILLONES"
   End If
   If varcentenamiles = "000" And valorcentena <> "000" Then
    MontoEnLetras = Cifra2ATexto(vardecenamillon) + " " + "MILLONES" + " " + Cifra3ATexto(valorcentena)
   End If
   If varcentenamiles <> "000" And valorcentena <> "000" Then
        MontoEnLetras = Cifra2ATexto(vardecenamillon) + " " + "MILLONES" + " " + Cifra3ATexto(varcentenamiles) + " " + "MIL" + " " + Cifra3ATexto(valorcentena)
   End If
   If varcentenamiles <> "000" And valorcentena = "000" Then
        MontoEnLetras = Cifra2ATexto(vardecenamillon) + " " + "MILLONES" + " " + Cifra3ATexto(varcentenamiles) + " " + "MIL" + " " + Cifra3ATexto(valorcentena)
   End If

End If


'******CONVERSION DE 7 CIFRAS********
If Len(ValorSinDecimal) = 7 Then

    varmillon = Val(Mid(ValorSinDecimal, 1, 1))
    varcentenamiles = Val(Mid(ValorSinDecimal, 2, 3))
    valorcentena = Mid(ValorSinDecimal, 5, 3)

    If varmillon = 1 Then
    
        MontoEnLetras = "UN MILLON" + " " + Cifra3ATexto(varcentenamiles) + "MIL" + " " + Cifra3ATexto(valorcentena)
    End If
    If varmillon > 1 Then
        MontoEnLetras = Cifra1ATexto(varmillon) + "MILLONES" + " " + Cifra3ATexto(varcentenamiles) + "MIL" + " " + Cifra3ATexto(valorcentena)
    End If

End If



'******CONVERSION DE 6 CIFRAS********
If Len(ValorSinDecimal) = 6 Then


    varcentenamiles = Val(Mid(ValorSinDecimal, 1, 3))
    valorcentena = Mid(ValorSinDecimal, 4, 3)

    'If valormil = 1 Then
    
    '    MontoEnLetras = "MIL" + " " + Cifra3ATexto(valorcentena)
    'End If
    'If valormil > 1 Then
        MontoEnLetras = Cifra3ATexto(varcentenamiles) + " " + "MIL" + " " + Cifra3ATexto(valorcentena)
    'End If

End If



'******CONVERSION DE 5 CIFRAS********
If Len(ValorSinDecimal) = 5 Then


    vardecenamiles = Val(Mid(ValorSinDecimal, 1, 2))
    valorcentena = Mid(ValorSinDecimal, 3, 3)

    'If valormil = 1 Then
    
    '    MontoEnLetras = "MIL" + " " + Cifra3ATexto(valorcentena)
    'End If
    'If valormil > 1 Then
        MontoEnLetras = Cifra2ATexto(vardecenamiles) + " " + "MIL" + " " + Cifra3ATexto(valorcentena)
    'End If

End If

'********CONVERSION DE 4 CIFRAS************
If Len(ValorSinDecimal) = 4 Then
    
    varmiles = Val(Mid(ValorSinDecimal, 1, 1))
    valorcentena = Mid(ValorSinDecimal, 2, 4)

    If valormil = 1 Then
    
        MontoEnLetras = "MIL" + " " + Cifra3ATexto(valorcentena)
    End If
    If valormil > 1 Then
        MontoEnLetras = Cifra1ATexto(valormil) + " " + "MIL" + " " + Cifra3ATexto(valorcentena)
    End If
End If

'*********CONVERSION DE 3 CIFRAS**********
If Len(ValorSinDecimal) = 3 Then
    MontoEnLetras = Cifra3ATexto(cifra)
End If

'********CONVERSION DE DOS CIFRAS*********
If Len(ValorSinDecimal) = 2 Then
    MontoEnLetras = Cifra2ATexto(cifra)

End If

'********CONVERSION DE UNA CIFRA**********
If Len(ValorSinDecimal) = 1 Then
    MontoEnLetras = Cifra1ATexto(cifra)

End If

End Function

Public Function Cifra3ATexto(valor)

Unidad(0) = "CERO"
Unidad(1) = "UNO"
Unidad(2) = "DOS"
Unidad(3) = "TRES"
Unidad(4) = "CUATRO"
Unidad(5) = "CINCO"
Unidad(6) = "SEIS"
Unidad(7) = "SIETE"
Unidad(8) = "OCHO"
Unidad(9) = "NUEVE"

DecenaEsp(0) = "DIEZ"
DecenaEsp(1) = "ONCE"
DecenaEsp(2) = "DOCE"
DecenaEsp(3) = "TRECE"
DecenaEsp(4) = "CATORCE"
DecenaEsp(5) = "QUINCE"
DecenaEsp(6) = "DIECI"

DecenaEsp2(10) = "DIEZ"
DecenaEsp2(11) = "ONCE"
DecenaEsp2(12) = "DOCE"
DecenaEsp2(13) = "TRECE"
DecenaEsp2(14) = "CATORCE"
DecenaEsp2(15) = "QUINCE"
DecenaEsp2(16) = "DIECI"

Decena(1) = "DIEZ"
Decena(2) = "VEINTI"
Decena(3) = "TREINTA"
Decena(4) = "CUARENTA"
Decena(5) = "CINCUENTA"
Decena(6) = "SESENTA"
Decena(7) = "SETENTA"
Decena(8) = "OCHENTA"
Decena(9) = "NOVENTA"
Decena(10) = "VEINTE"

Centena(1) = "CIEN"
Centena(2) = "CIENTO"
Centena(3) = "CIENTOS"
Centena(5) = "QUINIENTOS"
Centena(7) = "SETECIENTOS"
Centena(9) = "NOVECIENTOS"


primerdigito = Val(Mid(valor, 1, 1))
segdigito = Val(Mid(valor, 2, 1))
tercerdigito = Val(Mid(valor, 3, 1))
vardecena = Val(Mid(valor, 2, 2))
varunidad = Val(Mid(valor, 3, 1))
        
        
If primerdigito = 0 Then
    Cifra3ATexto = Cifra2ATexto(vardecena)
End If
        
If primerdigito = 1 Then
    If segdigito = 0 And tercerdigito = 0 Then
        Cifra3ATexto = Centena(1)
    Else
        Cifra3ATexto = Centena(2) + " " + Cifra2ATexto(vardecena)
    End If
End If

If primerdigito = 5 Then
    Cifra3ATexto = Centena(5) + " " + Cifra2ATexto(vardecena)
End If
            
If primerdigito = 7 Then
    Cifra3ATexto = Centena(7) + " " + Cifra2ATexto(vardecena)
End If
            
If primerdigito = 9 Then
    Cifra3ATexto = Centena(9) + " " + Cifra2ATexto(vardecena)
End If
            
If primerdigito <> 0 And primerdigito <> 1 And primerdigito <> 5 And primerdigito <> 7 And primerdigito <> 9 Then
    Cifra3ATexto = Unidad(primerdigito) + Centena(3) + " " + Cifra2ATexto(vardecena)
End If
            
If primerdigito = 0 And segdigito = 0 Then
    Cifra3ATexto = Cifra1ATexto(tercerdigito)
End If
            
If primerdigito <> 0 And segdigito = 0 And tercerdigito <> 0 Then
    Cifra3ATexto = Centena(primerdigito) + " " + Cifra1ATexto(tercerdigito)
End If
            
If primerdigito = 0 And segdigito = 0 And tercerdigito = 0 Then
    Cifra3ATexto = " "
End If

End Function

Public Function Cifra2ATexto(valor)

Unidad(0) = "CERO"
Unidad(1) = "UNO"
Unidad(2) = "DOS"
Unidad(3) = "TRES"
Unidad(4) = "CUATRO"
Unidad(5) = "CINCO"
Unidad(6) = "SEIS"
Unidad(7) = "SIETE"
Unidad(8) = "OCHO"
Unidad(9) = "NUEVE"

DecenaEsp2(10) = "DIEZ"
DecenaEsp2(11) = "ONCE"
DecenaEsp2(12) = "DOCE"
DecenaEsp2(13) = "TRECE"
DecenaEsp2(14) = "CATORCE"
DecenaEsp2(15) = "QUINCE"
DecenaEsp2(16) = "DIECI"

Decena(1) = "DIEZ"
Decena(2) = "VEINTI"
Decena(3) = "TREINTA"
Decena(4) = "CUARENTA"
Decena(5) = "CINCUENTA"
Decena(6) = "SESENTA"
Decena(7) = "SETENTA"
Decena(8) = "OCHENTA"
Decena(9) = "NOVENTA"
Decena(10) = "VEINTE"

primerdigito = Val(Mid(valor, 1, 1))
segdigito = Val(Mid(valor, 2, 1))
Valor2 = Val(valor)

Select Case Valor2
            Case 10 To 15
                Cifra2ATexto = DecenaEsp2(Valor2)
            Case 16 To 19
                valor = Mid(valor, 2, 1)
                Valor2 = Val(valor)
                Cifra2ATexto = DecenaEsp2(16) + Unidad(Valor2)
            Case Is = 20
                Cifra2ATexto = Decena(10)
            Case 21 To 29
                valor = Mid(valor, 2, 1)
                Valor2 = Val(valor)
                Cifra2ATexto = Decena(2) + Unidad(Valor2)
                Case Is = 30
                Cifra2ATexto = Decena(3)
            Case 31 To 39
                valor = Mid(valor, 2, 1)
                Valor2 = Val(valor)
                Cifra2ATexto = Decena(3) + " Y " + Unidad(Valor2)
            Case Is = 40
                Cifra2ATexto = Decena(4)
            Case 41 To 49
                valor = Mid(valor, 2, 1)
                Valor2 = Val(valor)
                Cifra2ATexto = Decena(4) + " Y " + Unidad(Valor2)
            Case Is = 50
                Cifra2ATexto = Decena(5)
            Case 51 To 59
                valor = Mid(valor, 2, 1)
                Valor2 = Val(valor)
                Cifra2ATexto = Decena(5) + " Y " + Unidad(Valor2)
            Case Is = 60
                Cifra2ATexto = Decena(6)
            Case 61 To 69
                valor = Mid(valor, 2, 1)
                Valor2 = Val(valor)
                Cifra2ATexto = Decena(6) + " Y " + Unidad(Valor2)
            Case Is = 70
                Cifra2ATexto = Decena(7)
            Case 71 To 79
                valor = Mid(valor, 2, 1)
                Valor2 = Val(valor)
                Cifra2ATexto = Decena(7) + " Y " + Unidad(Valor2)
            Case Is = 80
                Cifra2ATexto = Decena(8)
            Case 81 To 89
                valor = Mid(valor, 2, 1)
                Valor2 = Val(valor)
                Cifra2ATexto = Decena(8) + " Y " + Unidad(Valor2)
            Case Is = 90
                Cifra2ATexto = Decena(9)
            Case 91 To 99
                valor = Mid(valor, 2, 1)
                Valor2 = Val(valor)
                Cifra2ATexto = Decena(9) + " Y " + Unidad(Valor2)
        End Select

End Function

Public Function Cifra1ATexto(valor)

Unidad(0) = "CERO"
Unidad(1) = "UNO"
Unidad(2) = "DOS"
Unidad(3) = "TRES"
Unidad(4) = "CUATRO"
Unidad(5) = "CINCO"
Unidad(6) = "SEIS"
Unidad(7) = "SIETE"
Unidad(8) = "OCHO"
Unidad(9) = "NUEVE"


Valor2 = valor
Cifra1ATexto = Unidad(Valor2)

End Function

Public Sub HabilitaEdicionFlexGrid()
'se habilita t_editar en la columna 0 y en la fila disponible para que el cliente introduzca el codigo del articulo
        f_detallefactura.fg_detallefactura.Row = 1
        f_detallefactura.fg_detallefactura.Col = 0
        If f_detallefactura.fg_detallefactura.TextMatrix(f_detallefactura.fg_detallefactura.Row, 0) <> "" Then
            Do While f_detallefactura.fg_detallefactura.TextMatrix(f_detallefactura.fg_detallefactura.Row, 0) <> ""
                f_detallefactura.fg_detallefactura.Row = f_detallefactura.fg_detallefactura.Row + 1
            Loop
        End If
        f_detallefactura.t_editar.Left = f_detallefactura.fg_detallefactura.CellLeft + f_detallefactura.fg_detallefactura.Left
        f_detallefactura.t_editar.Top = f_detallefactura.fg_detallefactura.CellTop + f_detallefactura.fg_detallefactura.Top
        f_detallefactura.t_editar.Width = f_detallefactura.fg_detallefactura.CellWidth
        f_detallefactura.t_editar.Height = f_detallefactura.fg_detallefactura.CellHeight
        f_detallefactura.t_editar.BorderStyle = 0
        f_detallefactura.t_editar.FontName = f_detallefactura.fg_detallefactura.CellFontName
        f_detallefactura.t_editar.FontSize = f_detallefactura.fg_detallefactura.CellFontSize
        f_detallefactura.t_editar.FontBold = True
        f_detallefactura.t_editar.Visible = True ' se coloca visible el t_editar
        f_detallefactura.t_editar.Text = f_detallefactura.fg_detallefactura.TextMatrix(f_detallefactura.fg_detallefactura.Row, f_detallefactura.fg_detallefactura.Col)
        f_detallefactura.t_editar.SetFocus ' t_editar recibe el enfoque
End Sub

Public Sub ActivaEdicionItem()
        
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

Public Sub GenerarArchivoCab(VarNombreArchivo, ParametroSerie, ParametroNroDoc)
'#### CONSULTA PARA GENERAR ARCHIVO .CAB ####
                    
VarNombreArchivoCab = VarNombreArchivo + ".CAB"

Dim ArchivoCab As New ADODB.Recordset
Set ArchivoCab = New ADODB.Recordset
                    
ArchivoCab.Open "SELECT tipooper, fechaem, horaem, fechavenci, codlocalemisor, LEFT(tipdocusuario,1), numdocusuario, nombrers, tipmoneda, sumtottributos, sumtotvalventa, sumprecioventa, sumdesctotal, sumotroscargos, sumtotalanticipos, sumimpventa, ublversionld, customizationld FROM cabecera_doc WHERE serie= '" & ParametroSerie & "' AND nrodoc= " & ParametroNroDoc & " INTO OUTFILE '" & VarNombreArchivoCab & "' FIELDS TERMINATED BY '|' LINES TERMINATED BY '|'", Conn_Mysqldb, adOpenDynamic, adLockBatchOptimistic
                  
'#### FIN DE CONSULTA PARA GENERAR ARCHIVO .CAB ####
End Sub

Public Sub GenerarArchivoDet(VarNombreArchivo, ParametroSerie, ParametroNroDoc)
'#### INICIO CODIGO PARA GENERAR ARCHIVO .DET ####
                  Dim camp1, camp2, camp3, camp4, camp5, camp6, camp7, camp8 As String
                  
                  VarNombreArchivoDet = VarNombreArchivo + ".DET"
                    
                  Dim fso, txtdet
                  Set fso = CreateObject("Scripting.FileSystemObject")
                  Set txtdet = fso.CreateTextFile(VarNombreArchivoDet, True)
                  Dim ArchivoDet As New ADODB.Recordset
                  Set ArchivoDet = New ADODB.Recordset
                  ArchivoDet.Open "SELECT * FROM det_doc WHERE serie= '" & ParametroSerie & "' AND nrodoc= '" & ParametroNroDoc & "'", Conn_Mysqldb, adOpenDynamic, adLockBatchOptimistic
                  If ArchivoDet.BOF = False And ArchivoDet.EOF = False Then
                      ArchivoDet.MoveFirst
                      
                      Do While Not ArchivoDet.EOF
                      
                          camp1 = ArchivoDet.Fields(4).Value
                          camp2 = Format(ArchivoDet.Fields(5).Value, "Standard")
                          camp3 = ArchivoDet.Fields(3).Value
                          camp4 = "-"
                          camp5 = ArchivoDet.Fields(6).Value
                          camp6 = Format(ArchivoDet.Fields(7).Value, "Standard")
                          camp7 = Format(ArchivoDet.Fields(8).Value, "Standard")
                          camp8 = ArchivoDet.Fields(9).Value
                          camp9 = Format(ArchivoDet.Fields(10).Value, "Standard")
                          camp10 = Format(ArchivoDet.Fields(11).Value, "Standard")
                          camp11 = ArchivoDet.Fields(12).Value
                          camp12 = ArchivoDet.Fields(13).Value
                          camp13 = ArchivoDet.Fields(14).Value
                          camp14 = Format(ArchivoDet.Fields(15).Value, "Standard")
                          If ArchivoDet.Fields(16).Value = "-" Then
                              camp15 = ArchivoDet.Fields(16).Value
                              camp16 = ""
                              camp17 = ""
                              camp18 = ""
                              camp19 = ""
                              camp20 = ""
                              camp21 = ""
                          Else
                              camp15 = ArchivoDet.Fields(16).Value
                              camp16 = ArchivoDet.Fields(17).Value
                              camp17 = ArchivoDet.Fields(18).Value
                              camp18 = ArchivoDet.Fields(19).Value
                              camp19 = ArchivoDet.Fields(20).Value
                              camp20 = ArchivoDet.Fields(21).Value
                              camp21 = ArchivoDet.Fields(22).Value
                          End If
                          If ArchivoDet.Fields(23).Value = "-" Then
                              camp22 = ArchivoDet.Fields(23).Value
                              camp23 = ""
                              camp24 = ""
                              camp25 = ""
                              camp26 = ""
                              camp27 = ""
                          Else
                              camp22 = ArchivoDet.Fields(23).Value
                              camp23 = ArchivoDet.Fields(24).Value
                              camp24 = ArchivoDet.Fields(25).Value
                              camp25 = ArchivoDet.Fields(26).Value
                              camp26 = ArchivoDet.Fields(27).Value
                              camp27 = ArchivoDet.Fields(28).Value
                          End If
                          If ArchivoDet.Fields(29).Value = "-" Then
                              camp28 = ArchivoDet.Fields(29).Value
                              camp29 = ""
                              camp30 = ""
                              camp31 = ""
                              camp32 = ""
                              camp33 = ""
                          Else
                              camp28 = ArchivoDet.Fields(29).Value
                              camp29 = Format(ArchivoDet.Fields(30).Value, "Standard")
                              camp30 = Format(ArchivoDet.Fields(31).Value, "General Number")
                              camp31 = ArchivoDet.Fields(32).Value
                              camp32 = ArchivoDet.Fields(33).Value
                              camp33 = Format(ArchivoDet.Fields(34).Value, "Standard")
                          End If
                          
                          camp34 = Format(ArchivoDet.Fields(35).Value, "Standard")
                          camp35 = Format(ArchivoDet.Fields(36).Value, "Standard")
                          camp36 = "0"
                          
                          txtdet.WriteLine (camp1 + "|" + camp2 + "|" + camp3 + "|" + camp4 + "|" + camp5 + "|" + camp6 + "|" + camp7 + "|" + camp8 + "|" + _
                          camp9 + "|" + camp10 + "|" + camp11 + "|" + camp12 + "|" + camp13 + "|" + camp14 + "|" + camp15 + "|" + camp16 + "|" + _
                          camp17 + "|" + camp18 + "|" + camp19 + "|" + camp20 + "|" + camp21 + "|" + camp22 + "|" + camp23 + "|" + camp24 + "|" + _
                          camp25 + "|" + camp26 + "|" + camp27 + "|" + camp28 + "|" + camp29 + "|" + camp30 + "|" + camp31 + "|" + camp32 + "|" + _
                          camp33 + "|" + camp34 + "|" + camp35 + "|" + camp36 + "|")
                          ArchivoDet.MoveNext
                      Loop
                  
                  
                      txtdet.Close
                  End If
                '#### FIN DE CODIGO  PARA GENERAR ARCHIVO .DET ####
End Sub

Public Sub GenerarArchivoLey(VarNombreArchivo, ParametroTotal)
'########## INICIO DE CODIGO PARA GENERAR ARCHIVO .LEY ##########
                  
VarNombreArchivoley = VarNombreArchivo + ".LEY"
Dim txtley
Dim CampoLey1, CampoLey2 As String
                  
CampoLey1 = "1000"
                  
FinValEntero = InStr(1, ParametroTotal, ".")
FinValEntero = FinValEntero - 1
InicioDecimal = FinValEntero + 2
ValorSinDecimal = Mid(ParametroTotal, 1, FinValEntero)
ValorDecimal = Mid(ParametroTotal, InicioDecimal, 2)
              
CampoLey2 = MontoEnLetras(ValorSinDecimal) + " SOLES CON " + ValorDecimal + "/100"
                  
Set fso = CreateObject("Scripting.FileSystemObject")
Set txtley = fso.CreateTextFile(VarNombreArchivoley, True)
                  
txtley.WriteLine (CampoLey1 + "|" + CampoLey2 + "|")
                  
txtley.Close
                  
'########## FIN DE CODIGO PARA GENERAR ARCHIVO .LEY #############
End Sub

Public Sub GenerarArchivoTri(VarNombreArchivo, ParametroSerie, ParametroNroDoc)
'########## INICIO DE CODIGO PARA GENERAR ARCHIVO .TRI ##########
Dim ArchivoTri As ADODB.Recordset
VarNombreArchivoTri = VarNombreArchivo + ".TRI"
                     
Set ArchivoTri = New ADODB.Recordset
ArchivoTri.Open "SELECT cabecera_impuesto.codigo, nombrenacional, codinternacional, montobase, montoimpuesto FROM cabecera_impuesto INNER JOIN tipo_tributos ON cabecera_impuesto.codigo = tipo_tributos.codigo WHERE serie = '" & ParametroSerie & "' AND nrodoc = " & ParametroNroDoc & " AND montobase !=0 INTO OUTFILE '" & VarNombreArchivoTri & "' FIELDS TERMINATED BY '|' LINES TERMINATED BY '\r\n'", Conn_Mysqldb, adOpenDynamic, adLockBatchOptimistic
                     
'########## FIN DE CODIGO PARA GENERAR ARCHIVO .TRI #############
                     
End Sub

Public Sub GenerarArchivoEmail(VarNombreArchivo, ParametroEmail)
'########## INICIO DE CODIGO PARA GENERAR ARCHIVO .EMAIL #############
                     
VarNombreArchivoEmail = VarNombreArchivo + ".EMAIL"
                       
Dim txtemail
Set fso = CreateObject("Scripting.FileSystemObject")
Set txtemail = fso.CreateTextFile(VarNombreArchivoEmail, True)
txtemail.WriteLine (ParametroEmail)
txtemail.Close
                      
'########## FIN DE CODIGO PARA GENERAR ARCHIVO .EMAIL #############
          
End Sub

Public Sub ImprimirFormato80mm(TituloDoc, FechaVentaImpr, HoraVenta, VarSerie, VarNrodoc, _
VarPagado, VarAbono, VarRestante, VarNomCliente, VarNumDoc, VarFila, VarColum, VarTotalDoc, _
VarOpGravada, VarExoGraIna, VarIcbper, VarIgv, VarDsctos, VarEfectivo, VarVuelto, VarNomOperador, VarIdCaja)
'############ INICIO DE CODIGO PARA IMPRIMIR LA ORDEN ##############
            Dim TempDatosEmpresa As ADODB.Recordset
            Call Conn_BDaiosoft
            Set TempDatosEmpresa = New ADODB.Recordset
            TempDatosEmpresa.Open "SELECT ruc, nombre,direccionppal,telefonocelular,logo,telefonofijo,paginaweb FROM datos_empresa", Conn_Mysqldb, adOpenDynamic, adLockBatchOptimistic
            If TempDatosEmpresa.BOF = False And TempDatosEmpresa.EOF = False Then
                TempDatosEmpresa.MoveFirst
                VarRuc = TempDatosEmpresa.Fields(0).Value
                VarNombre = TempDatosEmpresa.Fields(1).Value
                VarDireccion = TempDatosEmpresa.Fields(2).Value
                VarTelefono = TempDatosEmpresa.Fields(3).Value
                VarLogo = TempDatosEmpresa.Fields(4).Value
                VarTelefonoFijo = TempDatosEmpresa.Fields(5).Value
                VarPaginaWeb = TempDatosEmpresa.Fields(6).Value
            Else
                VarRuc = " RUT DEMO"
                VarNombre = "NOMBRE DEMO"
                VarDireccion = "DIRECCION DEMO"
                VarTelefono = "TELEFONO DEMO"
                VarLogo = ""
            End If
            ' 567 twips equivalen a un centimetro
            AnchoPapel = 4073
            On Error Resume Next
            Printer.PaintPicture f_posrest.p_logo.Picture, 800, 400, 2402, 851
            Printer.CurrentY = 1418
            Printer.Font = "Courier New"
            Printer.FontSize = 8
            Printer.CurrentX = AnchoPapel / 2 - Printer.TextWidth(VarNombre) / 2
            Printer.Print VarNombre
            Printer.CurrentX = AnchoPapel / 2 - Printer.TextWidth("RUT: " & VarRuc) / 2
            Printer.Print "RUT: "; VarRuc
            Printer.CurrentX = AnchoPapel / 2 - Printer.TextWidth(VarDireccion) / 2
            Printer.Print VarDireccion
            Printer.CurrentX = AnchoPapel / 2 - Printer.TextWidth("TLF: " & VarTelefono & "/" & VarTelefonoFijo) / 2
            Printer.Print "TLF: "; VarTelefono; "/"; VarTelefonoFijo
            Printer.CurrentX = AnchoPapel / 2 - Printer.TextWidth("SITIO WEB: " & VarPaginaWeb) / 2
            Printer.Print "SITIO WEB: "; VarPaginaWeb
            Printer.Print
            Printer.CurrentX = AnchoPapel / 2 - Printer.TextWidth(TituloDoc) / 2
            Printer.Print TituloDoc
            Printer.Print
            Printer.CurrentX = 1
            'inicio codigo para ordenar el nro orden de derecha a izquierda
            largonumorden = Len(VarNrodoc)
            cantespacio = 9 - largonumorden
            'fin codigo para ordenar el nro orden de derecha a izquierda
            Printer.Print "FECHA: "; FechaVentaImpr; Spc(10); "NRO: "; Spc(espacioderecha - 1); VarSerie; "-"; VarNrodoc
            Printer.CurrentX = 1
            Printer.Print "HORA: "; HoraVenta
            Printer.CurrentX = 1
            Printer.Print "OPERADOR: "; VarNomOperador
            Printer.CurrentX = 1
            Printer.Print "CAJA: "; VarIdCaja
            Printer.CurrentX = 1
            Printer.Print "CLIENTE: "; VarNomCliente
            Printer.CurrentX = 1
            Printer.Print "NRO.DOC: "; VarNumDoc
            Printer.CurrentX = 1
            Printer.Print "___________________________________________"
            Printer.CurrentX = 1
            Printer.Print "DESCRIPCION"; Spc(11); "P.U"; Spc(5); "CNT"; Spc(3); "TOTAL $"
            Printer.CurrentX = 1
            Printer.Print "-------------------------------------------"
            For X = 1 To VarFila
                
            
                
                If Len(MatrizDetDoc(X, 1)) <= 42 Then
                    FILA1 = Mid(MatrizDetDoc(X, 1), 1, 42)
                    Printer.CurrentX = 1
                    Printer.Print FILA1
                End If
                If Len(MatrizDetDoc(X, 1)) > 42 And Len(MatrizDetDoc(X, 1)) < 87 Then
                    FILA1 = Mid(MatrizDetDoc(X, 1), 1, 42)
                    FILA2 = Mid(MatrizDetDoc(X, 1), 44, 42)
                    Printer.CurrentX = 1
                    Printer.Print FILA1
                    Printer.CurrentX = 1
                    Printer.Print FILA2
                End If
                
                'Printer.CurrentX = 1
                'Printer.Print Mid(MatrizDetDoc(x, 1), 1, 42)
                
                anchopu = Len(MatrizDetDoc(X, 2))
                ubipu = 1 + (27 - anchopu) * 96
                Printer.CurrentX = ubipu
                
                'CONDICIONAL A SOLICITUD DE CLIENTE
                'PARA NO MOSTRAR LOS PRECIOS SI LA CATEGORIA
                'DEL ARTICULO ES SERVICIO O REP
                If MatrizDetDoc(X, 5) = "SERVICIO" Or MatrizDetDoc(X, 5) = "REPUESTO" Then
                    MatrizDetDoc(X, 2) = ""
                    MatrizDetDoc(X, 3) = ""
                    MatrizDetDoc(X, 4) = ""
                End If
                'FIN DE CONDICIONAL
                Printer.Print MatrizDetDoc(X, 2)
                
                anchocant = Len(MatrizDetDoc(X, 3))
                ubicant = 1 + (31 - anchocant) * 96
                Printer.CurrentY = Printer.CurrentY - 170
                Printer.CurrentX = ubicant
                Printer.Print MatrizDetDoc(X, 3)
                
                anchototal = Len(MatrizDetDoc(X, 4))
                ubitotal = 1 + (42 - anchototal) * 96
                Printer.CurrentY = Printer.CurrentY - 170
                Printer.CurrentX = ubitotal
                Printer.Print MatrizDetDoc(X, 4)
            
            Next X
            Printer.CurrentX = 1
            Printer.Print "-------------------------------------------"
            
            'Printer.CurrentX = 1 + 96 * (42 - Len(VarOpGravada) - Len("OP. GRAVADA $: "))
            'Printer.Print "OP. GRAVADA $: "
            'Printer.CurrentX = 1 + 96 * (42 - Len(VarOpGravada))
            'Printer.CurrentY = Printer.CurrentY - 170
            'Printer.Print VarOpGravada
            
            'Printer.CurrentX = 1 + 96 * (42 - Len(VarIgv) - Len("TOTAL I.V.A $: "))
            'Printer.Print "TOTAL I.V.A $: "
            'Printer.CurrentX = 1 + 96 * (42 - Len(VarIgv))
            'Printer.CurrentY = Printer.CurrentY - 170
            'Printer.Print VarIgv
            
            'Printer.CurrentX = 1 + 96 * (42 - Len(VarExoGraIna) - Len("OP. EXO-GRA-INA S/: "))
            'Printer.Print "OP. EXO-GRA-INA $: "
            'Printer.CurrentX = 1 + 96 * (42 - Len(VarExoGraIna))
            'Printer.CurrentY = Printer.CurrentY - 170
            'Printer.Print VarExoGraIna
            
            'Printer.CurrentX = 1 + 96 * (42 - Len(VarIcbper) - Len("TOTAL ICBPER S/: "))
            'Printer.Print "TOTAL ICBPER $: "
            'Printer.CurrentX = 1 + 96 * (42 - Len(VarIcbper))
            'Printer.CurrentY = Printer.CurrentY - 170
            'Printer.Print VarIcbper
            
            If VarDsctos = "" Then
                VarDsctos = "0.00"
            End If
            
            Printer.CurrentX = 1 + 96 * (42 - Len(VarDsctos) - Len("TOT. DSCTOS $: "))
            Printer.Print "TOT. DSCTOS $: "
            Printer.CurrentX = 1 + 96 * (42 - Len(VarDsctos))
            Printer.CurrentY = Printer.CurrentY - 170
            Printer.Print VarDsctos
            
            Printer.CurrentX = 1 + 96 * (42 - Len(VarTotalDoc) - Len("TOTAL A PAGAR $: "))
            Printer.Print "TOTAL A PAGAR $: "
            Printer.CurrentX = 1 + 96 * (42 - Len(VarTotalDoc))
            Printer.CurrentY = Printer.CurrentY - 170
            Printer.Print VarTotalDoc
            
            VarEfectivo = Format(VarEfectivo, "standard")
            Printer.CurrentX = 1 + 96 * (42 - Len(VarEfectivo) - Len("RECIBIDO $: "))
            Printer.Print "RECIBIDO $: "
            Printer.CurrentX = 1 + 96 * (42 - Len(VarEfectivo))
            Printer.CurrentY = Printer.CurrentY - 170
            Printer.Print VarEfectivo
            
            VarVuelto = Format(VarVuelto, "standard")
            VarVuelto2 = Format(VarVuelto, "standard")
          
            Printer.CurrentX = 1 + 96 * (42 - Len(VarVuelto2) - Len("VUELTO $: "))
            Printer.Print "VUELTO $: "
            Printer.CurrentX = 1 + 96 * (42 - Len(VarVuelto2))
            Printer.CurrentY = Printer.CurrentY - 170
            Printer.Print VarVuelto2

            If VarPagado = "s" Then
                Printer.CurrentX = 1
                Printer.Print "PAGADO"
            End If
            If VarPagado = "n" Then
                Printer.CurrentX = 1
                Printer.Print "ABONADO $:"; Spc(1); VarAbono
                Printer.CurrentX = 1
                Printer.Print "RESTANTE $:"; Spc(1); VarRestante
                Printer.CurrentX = 1
                Printer.Print "POR PAGAR"
            End If
            
            
            Printer.EndDoc
            
            '**************** FIN DE CODIGO DE IMPRESION DE ORDEN***************
End Sub

Public Sub GenerarOrden()
Dim UsarOrdenExistente As Boolean
UsarOrdenExistente = False

'##### AQUI INICIA EL CODIGO PARA ALMACENAR LOS DATOS DEL DOCUMENTO EN LA TABLA cabecera_doc #####
 Dim TempCabDoc As ADODB.Recordset
        'buscar si hay ordenes abiertas no utilizadas para reutilizarla
        Call Conn_BDaiosoft
        Set TempCabDoc = New ADODB.Recordset
        TempCabDoc.Open "SELECT nroorden FROM cabecera_doc WHERE estado='abierta' AND idoperador=" & f_principal.l_idoperador.Caption & " AND idcaja= " & f_principal.l_idcaja.Caption & "  ORDER BY nroorden", Conn_Mysqldb, adOpenDynamic, adLockBatchOptimistic
        If TempCabDoc.BOF = False And TempCabDoc.EOF = False Then
            numorden = TempCabDoc.Fields(0).Value
            f_posrest.l_numorden.Caption = numorden
            UsarOrdenExistente = True
            
            'consulta para actualizar la tabla cabecera_doc
            Conn_Mysqldb.Execute "UPDATE cabecera_doc SET serie= '0000'," _
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
            & "sumtottributos = 0, sumtotvalventa = 0, sumprecioventa = 0," _
            & "sumdesctotal = 0, sumotroscargos = 0, sumtotalanticipos = 0," _
            & "sumimpventa = 0, ublversionld = '2.1', customizationld = '2.0'," _
            & "idvendedor = '.', abono = 0, restante = 0, impreso = 's'," _
            & " estado = 'abierta', codcliente= " & f_posrest.l_idcliente.Caption & "," _
            & "idoperador=" & f_principal.l_idoperador.Caption & "," _
            & "idcaja= " & f_principal.l_idcaja.Caption & " WHERE nroorden= " & numorden & ""
               
            f_posrest.t_codpro.Enabled = True
            f_posrest.t_cantidad.Enabled = True
            
                
        Else
            UsarOrdenExistente = False
            TempCabDoc.Close
        End If
        If UsarOrdenExistente = False Then
            'si no hay ordenes abiertas crea una orden nueva
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
            
            f_posrest.t_codpro.Enabled = True
            f_posrest.t_cantidad.Enabled = True
            
        
        End If
            
               
            '##### AQUI FINALIZA EL CODIGO PARA ALMACENAR LOS DATOS DEL DOCUMENTO EN LA TABLA cabecera_doc #####

'##### AQUI INICIA EL CODIGO PARA ALMACENAR LOS DATOS DEL DOCUMENTO EN LA TABLA cabecera_doc #####
 'Dim TempCabDoc As ADODB.Recordset
 '       '****codigo para cargar el numero de orden siguiente****
 '       Call Conn_BDaiosoft
 '       Set TempCabDoc = New ADODB.Recordset
 '       TempCabDoc.Open "SELECT nroorden FROM cabecera_doc  ORDER BY nroorden", Conn_Mysqldb, adOpenDynamic, adLockBatchOptimistic
 '       If TempCabDoc.BOF = False And TempCabDoc.EOF = False Then
 '           TempCabDoc.MoveLast
 '           numorden = TempCabDoc.Fields(0).Value + 1
 '           f_posrest.l_numorden.Caption = numorden
 '       Else
 '           f_posrest.l_numorden.Caption = 1
 '       End If
        '****fin de codigo para cargar el numero de orden siguiente****

            
            'consulta para la insercion de datos en la tabla cabecera_doc
 '           Conn_Mysqldb.Execute "INSERT INTO cabecera_doc SET serie= '0000'," _
 '           & "nroorden = " & f_posrest.l_numorden.Caption & "," _
 '           & "nrodoc = " & f_posrest.l_numorden.Caption & "," _
 '           & "tipodoc = '00'," _
 '           & "tipooper = '0101'," _
 '           & "firmadig = '-'," _
 '           & "fechaem = '" & Format(Date, "yyyy-mm-dd") & "'," _
 '           & "horaem = '" & Format(Time, "HH:MM:SS") & "'," _
 '           & "fechavenci = '" & Format(Date, "yyyy-mm-dd") & "'," _
 '           & "codlocalemisor = '0000'," _
 '           & "tipdocusuario = '" & f_posrest.c_tipodoc.Text & "'," _
 '           & "numdocusuario = '" & f_posrest.t_nrodoc.Text & "'," _
 '           & "nombrers = '" & f_posrest.t_nombrers.Text & "'," _
 '           & "tipmoneda = 'PEN'," _
 '           & "sumtottributos = 0," _
 '           & "sumtotvalventa = 0," _
 '           & "sumprecioventa = 0," _
 '           & "sumdesctotal = 0," _
 '           & "sumotroscargos = 0," _
 '           & "sumtotalanticipos = 0," _
 '           & "sumimpventa = 0," _
 '           & "ublversionld = '2.1'," _
 '           & "customizationld = '2.0'," _
 '           & "idvendedor = '.', abono = 0, restante = 0," _
 '           & "impreso = 's', estado = 'abierta', codcliente= " & f_posrest.l_idcliente.Caption & ", idoperador=" & f_principal.l_idoperador.Caption & ",idcaja= " & f_principal.l_idcaja.Caption & ""
                             
            '##### AQUI FINALIZA EL CODIGO PARA ALMACENAR LOS DATOS DEL DOCUMENTO EN LA TABLA cabecera_doc #####
End Sub

Public Sub RegistrarCliente()

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
    'f_posrest.t_codpro.SetFocus
    
    Call GenerarOrden
Else
    MsgBox "Los campos número de documento y nombre o razón social no pueden estar en blanco."
End If
End Sub


