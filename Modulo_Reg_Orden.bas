Attribute VB_Name = "Modulo_Reg_Orden"

Public Function DevolverPrecio(Cod_Prod_Orden)

Dim temproducto As New ADODB.Recordset

Call conn_BDaiosoft
Set temproducto = New ADODB.Recordset
temproducto.Open "SELECT precio  FROM producto where codigo = '" & Trim(Cod_Prod_Orden) & "'", conn_mysqldb, adOpenDynamic, adLockBatchOptimistic
If temproducto.BOF = False And temproducto.EOF = False Then
    temproducto.MoveFirst

VarPrecio = temproducto.Fields(0).Value

VarPrecio = Format(temproducto.Fields(0).Value, "standard")
DevolverPrecio = VarPrecio

End If



End Function

Public Sub CambioColorBoton()


'codigo para cambiar la informacion de la etiqueta label9
If NombreBoton = "b_boton1" Then
    registro_orden.Label9.Caption = "Precio x KG"
Else
    registro_orden.Label9.Caption = "Precio x Pieza"
End If

'codigo paga cambiar el color del boton al haber sido precionado
If NombreBoton = "b_boton1" Then
    registro_orden.b_boton1.BackColor = &H808080
End If
If NombreBoton = "b_boton2" Then
    registro_orden.b_boton2.BackColor = &H808080
End If
If NombreBoton = "b_boton3" Then
    registro_orden.b_boton3.BackColor = &H808080
End If
If NombreBoton = "b_boton4" Then
    registro_orden.b_boton4.BackColor = &H808080
End If
If NombreBoton = "b_boton5" Then
    registro_orden.b_boton5.BackColor = &H808080
End If
If NombreBoton = "b_boton6" Then
    registro_orden.b_boton6.BackColor = &H808080
End If
If NombreBoton = "b_boton7" Then
    registro_orden.b_boton7.BackColor = &H808080
End If
If NombreBoton = "b_boton8" Then
    registro_orden.b_boton8.BackColor = &H808080
End If
If NombreBoton = "b_boton9" Then
    registro_orden.b_boton9.BackColor = &H808080
End If
If NombreBoton = "b_boton10" Then
    registro_orden.b_boton10.BackColor = &H808080
End If
If NombreBoton = "b_boton11" Then
    registro_orden.b_boton11.BackColor = &H808080
End If
If NombreBoton = "b_boton12" Then
    registro_orden.b_boton12.BackColor = &H808080
End If
If NombreBoton = "b_boton13" Then
    registro_orden.b_boton13.BackColor = &H808080
End If
If NombreBoton = "b_boton14" Then
    registro_orden.b_boton14.BackColor = &H808080
End If
If NombreBoton = "b_boton15" Then
    registro_orden.b_boton15.BackColor = &H808080
End If
If NombreBoton = "b_boton16" Then
    registro_orden.b_boton16.BackColor = &H808080
End If
If NombreBoton = "b_boton17" Then
    registro_orden.b_boton17.BackColor = &H808080
End If

If NombreBoton = "b_boton18" Then
    registro_orden.b_boton18.BackColor = &H808080
End If
If NombreBoton = "b_boton19" Then
    registro_orden.b_boton19.BackColor = &H808080
End If
If NombreBoton = "b_boton20" Then
    registro_orden.b_boton20.BackColor = &H808080
End If
If NombreBoton = "b_boton21" Then
    registro_orden.b_boton21.BackColor = &H808080
End If
If NombreBoton = "b_boton22" Then
    registro_orden.b_boton22.BackColor = &H808080
End If
If NombreBoton = "b_boton23" Then
    registro_orden.b_boton23.BackColor = &H808080
End If
If NombreBoton = "b_boton24" Then
    registro_orden.b_boton24.BackColor = &H808080
End If
If NombreBoton = "b_boton25" Then
    registro_orden.b_boton25.BackColor = &H808080
End If
If NombreBoton = "b_boton26" Then
    registro_orden.b_boton26.BackColor = &H808080
End If
If NombreBoton = "b_boton27" Then
    registro_orden.b_boton27.BackColor = &H808080
End If
If NombreBoton = "b_boton28" Then
    registro_orden.b_boton28.BackColor = &H808080
End If
If NombreBoton = "b_boton29" Then
    registro_orden.b_boton29.BackColor = &H808080
End If
If NombreBoton = "b_boton30" Then
    registro_orden.b_boton30.BackColor = &H808080
End If
If NombreBoton = "b_boton31" Then
    registro_orden.b_boton31.BackColor = &H808080
End If
If NombreBoton = "b_boton32" Then
    registro_orden.b_boton32.BackColor = &H808080
End If
If NombreBoton = "b_boton33" Then
    registro_orden.b_boton33.BackColor = &H808080
End If
If NombreBoton = "b_boton34" Then
    registro_orden.b_boton34.BackColor = &H808080
End If
If NombreBoton = "b_boton35" Then
    registro_orden.b_boton35.BackColor = &H808080
End If

End Sub

Public Sub HabilitarTexbox()
registro_orden.t_peso.Visible = False
registro_orden.Label6.Visible = False
registro_orden.b_agregar.Enabled = True
registro_orden.b_cancelar.Enabled = True
registro_orden.t_cantidad.Enabled = True
registro_orden.t_precio.Enabled = True
registro_orden.t_cantidad.SetFocus
End Sub

Public Sub CambioColorBotonOriginal()

'codigo para cambiar la informacion de la etiqueta label9
If NombreBoton = "b_boton1" Then
    registro_orden.Label9.Caption = "Precio x KG"
Else
    registro_orden.Label9.Caption = "Precio x Pieza"
End If

'codigo paga cambiar el color del boton al perder el foco

If NombreBoton = "b_boton1" Then
    registro_orden.b_boton1.BackColor = &HFFFFFF
End If
If NombreBoton = "b_boton2" Then
    registro_orden.b_boton2.BackColor = &HFFFFFF
End If
If NombreBoton = "b_boton3" Then
    registro_orden.b_boton3.BackColor = &HFFFFFF
End If
If NombreBoton = "b_boton4" Then
    registro_orden.b_boton4.BackColor = &HFFFFFF
End If
If NombreBoton = "b_boton5" Then
    registro_orden.b_boton5.BackColor = &HFFFFFF
End If
If NombreBoton = "b_boton6" Then
    registro_orden.b_boton6.BackColor = &HFFFFFF
End If
If NombreBoton = "b_boton7" Then
    registro_orden.b_boton7.BackColor = &HFFFFFF
End If
If NombreBoton = "b_boton8" Then
    registro_orden.b_boton8.BackColor = &HFFFFFF
End If
If NombreBoton = "b_boton9" Then
    registro_orden.b_boton9.BackColor = &HFFFFFF
End If
If NombreBoton = "b_boton10" Then
    registro_orden.b_boton10.BackColor = &HFFFFFF
End If
If NombreBoton = "b_boton11" Then
    registro_orden.b_boton11.BackColor = &HFFFFFF
End If
If NombreBoton = "b_boton12" Then
    registro_orden.b_boton12.BackColor = &HFFFFFF
End If
If NombreBoton = "b_boton13" Then
    registro_orden.b_boton13.BackColor = &HFFFFFF
End If
If NombreBoton = "b_boton14" Then
    registro_orden.b_boton14.BackColor = &HFFFFFF
End If
If NombreBoton = "b_boton15" Then
    registro_orden.b_boton15.BackColor = &HFFFFFF
End If
If NombreBoton = "b_boton16" Then
    registro_orden.b_boton16.BackColor = &HFFFFFF
End If
If NombreBoton = "b_boton17" Then
    registro_orden.b_boton17.BackColor = &HFFFFFF
End If
If NombreBoton = "b_boton18" Then
    registro_orden.b_boton18.BackColor = &HFFFFFF
End If
If NombreBoton = "b_boton19" Then
    registro_orden.b_boton19.BackColor = &HFFFFFF
End If
If NombreBoton = "b_boton20" Then
    registro_orden.b_boton20.BackColor = &HFFFFFF
End If
If NombreBoton = "b_boton21" Then
    registro_orden.b_boton21.BackColor = &HFFFFFF
End If
If NombreBoton = "b_boton22" Then
    registro_orden.b_boton22.BackColor = &HFFFFFF
End If
If NombreBoton = "b_boton23" Then
    registro_orden.b_boton23.BackColor = &HFFFFFF
End If
If NombreBoton = "b_boton24" Then
    registro_orden.b_boton24.BackColor = &HFFFFFF
End If
If NombreBoton = "b_boton25" Then
    registro_orden.b_boton25.BackColor = &HFFFFFF
End If
If NombreBoton = "b_boton26" Then
    registro_orden.b_boton26.BackColor = &HFFFFFF
End If
If NombreBoton = "b_boton27" Then
    registro_orden.b_boton27.BackColor = &HFFFFFF
End If
If NombreBoton = "b_boton28" Then
    registro_orden.b_boton28.BackColor = &HFFFFFF
End If
If NombreBoton = "b_boton29" Then
    registro_orden.b_boton29.BackColor = &HFFFFFF
End If
If NombreBoton = "b_boton30" Then
    registro_orden.b_boton30.BackColor = &HFFFFFF
End If
If NombreBoton = "b_boton31" Then
    registro_orden.b_boton31.BackColor = &HFFFFFF
End If
If NombreBoton = "b_boton32" Then
    registro_orden.b_boton32.BackColor = &HFFFFFF
End If
If NombreBoton = "b_boton33" Then
    registro_orden.b_boton33.BackColor = &HFFFFFF
End If
If NombreBoton = "b_boton34" Then
    registro_orden.b_boton34.BackColor = &HFFFFFF
End If
If NombreBoton = "b_boton35" Then
    registro_orden.b_boton35.BackColor = &HFFFFFF
End If

End Sub
