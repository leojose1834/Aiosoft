VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form registro_de_compras 
   Caption         =   "Registro de Compras"
   ClientHeight    =   3120
   ClientLeft      =   60
   ClientTop       =   420
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3120
   ScaleWidth      =   4680
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   8895
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   14535
      Begin VB.TextBox t_nrocompra 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   11160
         TabIndex        =   11
         Top             =   240
         Width           =   2055
      End
      Begin VB.CommandButton b_producto 
         Caption         =   "Abrir Lista de Productos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   5
         Top             =   1680
         Width           =   3255
      End
      Begin VB.TextBox t_rif 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   9240
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   960
         Width           =   1695
      End
      Begin VB.TextBox t_editar 
         BackColor       =   &H00C0FFFF&
         Height          =   375
         Left            =   2520
         TabIndex        =   8
         Top             =   3960
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.CommandButton b_factura 
         Caption         =   "Guardar Compra"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   5760
         Picture         =   "registro_de_compras.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   7200
         Width           =   3495
      End
      Begin MSFlexGridLib.MSFlexGrid fg_detallecompra 
         Height          =   4935
         Left            =   360
         TabIndex        =   6
         Top             =   2040
         Width           =   13935
         _ExtentX        =   24580
         _ExtentY        =   8705
         _Version        =   393216
         Rows            =   20
         Cols            =   8
         FixedCols       =   0
         FormatString    =   $"registro_de_compras.frx":119A
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.TextBox t_proveedor 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2520
         TabIndex        =   4
         Top             =   960
         Width           =   5415
      End
      Begin VB.CommandButton b_proveedor 
         Caption         =   "Proveedor:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   480
         TabIndex        =   3
         Top             =   960
         Width           =   1695
      End
      Begin MSComCtl2.DTPicker dt_fechacompra 
         Height          =   375
         Left            =   2160
         TabIndex        =   2
         Top             =   360
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
         _Version        =   393216
         Format          =   39911425
         CurrentDate     =   39185
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Nº de Compra:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   9480
         TabIndex        =   12
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Rif:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   8520
         TabIndex        =   9
         Top             =   960
         Width           =   495
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Fecha:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   480
         TabIndex        =   1
         Top             =   360
         Width           =   975
      End
   End
End
Attribute VB_Name = "registro_de_compras"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub b_factura_Click()
Dim datoscompra As Recordset
If t_nrocompra.Text = "" Or t_proveedor.Text = "" Then
    MsgBox "No puede dejar en blanco  los campos PROVEEDOR y NÚMERO DE COMPRA"
Else
    Set datoscompra = bdsistfact.OpenRecordset("select numerocompra from datoscompra")
    criterio = "trim(numerocompra)='" & Trim(t_nrocompra.Text) & "'"
    datoscompra.FindFirst criterio
    If datoscompra.NoMatch Then
        f_totalcompra.Show
    Else
        MsgBox "El numero de la compra ya existe por favor intente con otro"
    End If
End If
End Sub

Private Sub b_producto_Click()
'muestra el formulario lista de productos  cuando se haga clic sobre el boton
' "motrar listado de productos"
f_listadoproductos.Show
End Sub

Private Sub b_proveedor_Click()
'muestra el formulario lista de proveedores cuando se haga clic sobre el boton
' "proveedor"
f_listadoproveedor.Show
End Sub
Private Sub fg_detallecompra_Click()
If fg_detallecompra.Row >= 1 And fg_detallecompra.TextMatrix(fg_detallecompra.Row - 1, 0) <> "" Then
    If fg_detallecompra.Col = 0 Or fg_detallecompra.Col = 1 Then
    'determina si el valor de la columna del flexgrid es diferente de 5
    'o es diferente de 7  ya que esas columnas no se pueden editar
    'If fg_detallecompra.Col <> 5 And fg_detallecompra.Col <> 7 Then
    
        'se procede a dar al t_editar  todos los valores de propiedad de la
        'de la celda que tiene el enfoque
        t_editar.Left = fg_detallecompra.CellLeft + fg_detallecompra.Left
        t_editar.Top = fg_detallecompra.CellTop + fg_detallecompra.Top
        t_editar.Width = fg_detallecompra.CellWidth
        t_editar.Height = fg_detallecompra.CellHeight
        t_editar.BorderStyle = 0
        t_editar.FontName = fg_detallecompra.CellFontName
        t_editar.FontSize = fg_detallecompra.CellFontSize
        t_editar.FontBold = True
        t_editar.Visible = True ' se coloca visible el t_editar
        t_editar.SetFocus ' t_editar recibe el enfoque
        t_editar.Text = fg_detallecompra.TextMatrix(fg_detallecompra.Row, fg_detallecompra.Col)
        'se pasa al t_editar, el valor de la celda que tiene el enfoque
    
    End If
End If
If fg_detallecompra.TextMatrix(fg_detallecompra.Row, 0) <> "" Then

    'determina si el valor de la columna del flexgrid es diferente de 5
    'o es diferente de 7  ya que esas columnas no se pueden editar
    If fg_detallecompra.Col <> 5 And fg_detallecompra.Col <> 7 And fg_detallecompra.Col <> 2 Then
        'se procede a dar al t_editar  todos los valores de propiedad de la
        'de la celda que tiene el enfoque
        t_editar.Left = fg_detallecompra.CellLeft + fg_detallecompra.Left
        t_editar.Top = fg_detallecompra.CellTop + fg_detallecompra.Top
        t_editar.Width = fg_detallecompra.CellWidth
        t_editar.Height = fg_detallecompra.CellHeight
        t_editar.BorderStyle = 0
        t_editar.FontName = fg_detallecompra.CellFontName
        t_editar.FontSize = fg_detallecompra.CellFontSize
        t_editar.FontBold = True
        t_editar.Visible = True ' se coloca visible el t_editar
        t_editar.SetFocus ' t_editar recibe el enfoque
        t_editar.Text = fg_detallecompra.TextMatrix(fg_detallecompra.Row, fg_detallecompra.Col)
        'se pasa al t_editar, el valor de la celda que tiene el enfoque
    End If
End If
End Sub

Private Sub fg_detallecompra_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyDelete And fg_detallecompra.TextMatrix(fg_detallecompra.Row, 0) <> "" Then
    msg = MsgBox("¿Seguro desea quitar de la compra el producto  " + Trim(DescripcionCodigoProducto(fg_detallecompra.Row, 0)) + "?", vbQuestion + vbYesNo)
    If msg = vbYes Then
        For columna = 0 To 7
            fg_detallecompra.TextMatrix(fg_detallecompra.Row, columna) = ""
        Next columna
        DescripcionCodigoProducto(fg_detallecompra.Row, 0) = ""
        DescripcionCodigoProducto(fg_detallecompra.Row, 1) = ""
        filaenblanco = fg_detallecompra.Row
        Do While fg_detallecompra.TextMatrix(filaenblanco + 1, 0) <> ""
            For columna = 0 To 7
                fg_detallecompra.TextMatrix(filaenblanco, columna) = fg_detallecompra.TextMatrix(filaenblanco + 1, columna)
                fg_detallecompra.TextMatrix(filaenblanco + 1, columna) = ""
            Next columna
            DescripcionCodigoProducto(filaenblanco, 0) = fg_detallecompra.TextMatrix(filaenblanco, 0)
            DescripcionCodigoProducto(filaenblanco, 1) = fg_detallecompra.TextMatrix(filaenblanco, 1)
            filaenblanco = filaenblanco + 1
        Loop
    End If
    fg_detallecompra.SetFocus
End If
End Sub


Private Sub fg_detallecompra_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 And fg_detallecompra.Row <> 0 And fg_detallecompra.TextMatrix(fg_detallecompra.Row - 1, 0) <> "" Then
    If fg_detallecompra.Col = 0 Or fg_detallecompra.Col = 1 Then
    'determina si el valor de la columna del flexgrid es diferente de 5
    'o es diferente de 7  ya que esas columnas no se pueden editar
    'If fg_detallecompra.Col <> 5 And fg_detallecompra.Col <> 7 Then
    
        'se procede a dar al t_editar  todos los valores de propiedad de la
        'de la celda que tiene el enfoque
        t_editar.Left = fg_detallecompra.CellLeft + fg_detallecompra.Left
        t_editar.Top = fg_detallecompra.CellTop + fg_detallecompra.Top
        t_editar.Width = fg_detallecompra.CellWidth
        t_editar.Height = fg_detallecompra.CellHeight
        t_editar.BorderStyle = 0
        t_editar.FontName = fg_detallecompra.CellFontName
        t_editar.FontSize = fg_detallecompra.CellFontSize
        t_editar.FontBold = True
        t_editar.Visible = True ' se coloca visible el t_editar
        t_editar.SetFocus ' t_editar recibe el enfoque
        t_editar.Text = fg_detallecompra.TextMatrix(fg_detallecompra.Row, fg_detallecompra.Col)
        'se pasa al t_editar, el valor de la celda que tiene el enfoque
    End If
End If
If KeyAscii = 13 And fg_detallecompra.TextMatrix(fg_detallecompra.Row, 0) <> "" Then
    
    'determina si el valor de la columna del flexgrid es diferente de 5
    'o es diferente de 7  ya que esas columnas no se pueden editar
    If fg_detallecompra.Col <> 5 And fg_detallecompra.Col <> 7 And fg_detallecompra.Col <> 2 Then
        'se procede a dar al t_editar  todos los valores de propiedad de la
        'de la celda que tiene el enfoque
        t_editar.Left = fg_detallecompra.CellLeft + fg_detallecompra.Left
        t_editar.Top = fg_detallecompra.CellTop + fg_detallecompra.Top
        t_editar.Width = fg_detallecompra.CellWidth
        t_editar.Height = fg_detallecompra.CellHeight
        t_editar.BorderStyle = 0
        t_editar.FontName = fg_detallecompra.CellFontName
        t_editar.FontSize = fg_detallecompra.CellFontSize
        t_editar.FontBold = True
        t_editar.Visible = True ' se coloca visible el t_editar
        t_editar.SetFocus ' t_editar recibe el enfoque
        t_editar.Text = fg_detallecompra.TextMatrix(fg_detallecompra.Row, fg_detallecompra.Col)
        'se pasa al t_editar, el valor de la celda que tiene el enfoque
    End If
End If
End Sub

Private Sub Form_Activate()
dt_fechacompra = Date
End Sub

Private Sub Form_Unload(Cancel As Integer)
'cuando se descargue el formulario, se le dará valor 0 a la variable global
'FilaDetalleCompra, ya que esta va aumentando de uno en uno cuando se agrega un nuevo
'producto para que genere una nueva fila en el flexgrid detallecompra.
'Se da un valor cero, para que cuando se habra nuevamente el formulario,
'empiece a agregar productos a partir de la fila 2
FilaDetalleCompra = 0
End Sub

Private Sub t_editar_Change() 'codigo para calcular automaticamente el importe
ValorFila = fg_detallecompra.Row
ValorColumna = fg_detallecompra.Col
' que debe pagar cada producto
' determina si  el valor del flex grid en la fila que este y en la columna
' 0 es diferente de blanco (nada)
If fg_detallecompra.TextMatrix(fg_detallecompra.Row, 0) <> "" Then

    'determina que el valor del flexgrid en la fila y en la columna que este en ese
    'momento va a ser igual al valor del  t_editar
    
    fg_detallecompra.TextMatrix(fg_detallecompra.Row, fg_detallecompra.Col) = t_editar.Text
    
    ' determina si se esta en la columna 3 que es la que tiene el valor de la
    'cantidad de cada producto
    If fg_detallecompra.Col = 3 Then
        cantidad = Val(t_editar.Text)
        costounitario = Val(fg_detallecompra.TextMatrix(fg_detallecompra.Row, 4))
        descuento = Val(fg_detallecompra.TextMatrix(fg_detallecompra.Row, 6))
        fg_detallecompra.TextMatrix(fg_detallecompra.Row, 7) = cantidad * costounitario - ((costounitario * cantidad) * descuento / 100)
        Importe = fg_detallecompra.TextMatrix(fg_detallecompra.Row, 7)
        DescripcionCodigoProducto(registro_de_compras.fg_detallecompra.Row, 3) = cantidad
        DescripcionCodigoProducto(registro_de_compras.fg_detallecompra.Row, 7) = Importe
        'el valor del importe por cada producto va a ser igual a la multiplicacion del costo
        'unitario por la cantidad
    End If
    
    ' determina si se esta en la columna 3 que es la que tiene el valor del
    'costo unitario del producto producto
    If fg_detallecompra.Col = 4 Then
        cantidad = Val(fg_detallecompra.TextMatrix(fg_detallecompra.Row, 3))
        costounitario = Val(t_editar.Text)
        descuento = Val(fg_detallecompra.TextMatrix(fg_detallecompra.Row, 6))
        fg_detallecompra.TextMatrix(fg_detallecompra.Row, 7) = cantidad * costounitario - ((costounitario * cantidad) * descuento / 100)
        Importe = fg_detallecompra.TextMatrix(fg_detallecompra.Row, 7)
        DescripcionCodigoProducto(registro_de_compras.fg_detallecompra.Row, 4) = costounitario
        DescripcionCodigoProducto(registro_de_compras.fg_detallecompra.Row, 7) = Importe
        'el valor del importe por cada producto va a ser igual a la multiplicacion del costo
        'unitario por la cantidad
    End If
    If fg_detallecompra.Col = 6 Then
        cantidad = Val(fg_detallecompra.TextMatrix(fg_detallecompra.Row, 3))
        costounitario = Val(fg_detallecompra.TextMatrix(fg_detallecompra.Row, 4))
        descuento = Val(t_editar.Text)
        fg_detallecompra.TextMatrix(fg_detallecompra.Row, 7) = cantidad * costounitario - ((costounitario * cantidad) * descuento / 100) 'el valor del
        Importe = fg_detallecompra.TextMatrix(fg_detallecompra.Row, 7)
        DescripcionCodigoProducto(registro_de_compras.fg_detallecompra.Row, 6) = descuento
        DescripcionCodigoProducto(registro_de_compras.fg_detallecompra.Row, 7) = Importe
        'importe por cada producto va a ser igual a la multiplicacion del costo
        'unitario por la cantidad
    End If
End If
End Sub


Private Sub t_editar_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyDown Then

   If t_editar.Text = "" And fg_detallecompra.Col = 0 Then
      fg_detallecompra.TextMatrix(fg_detallecompra.Row, 0) = DescripcionCodigoProducto(fg_detallecompra.Row, 0)
    End If

   If t_editar.Text = "" And fg_detallecompra.Col = 1 Then
        fg_detallecompra.TextMatrix(fg_detallecompra.Row, 1) = DescripcionCodigoProducto(fg_detallecompra.Row, 1)
   End If
    'se va aumentando de uno en uno el valor de las columnas del flexgrid
    'para que se vaya desplazando a traves de las celdas
    fg_detallecompra.Row = fg_detallecompra.Row + 1
    fg_detallecompra.SetFocus
    t_editar.Visible = False
    End If

'si se preciona la tecla flecha arriba
If KeyCode = vbKeyUp And fg_detallecompra.Row > 1 Then
    If t_editar.Text = "" And fg_detallecompra.Col = 0 Then
       fg_detallecompra.TextMatrix(fg_detallecompra.Row, 0) = DescripcionCodigoProducto(fg_detallecompra.Row, 0)
    End If
    
    If t_editar.Text = "" And fg_detallecompra.Col = 1 Then
        fg_detallecompra.TextMatrix(fg_detallecompra.Row, 1) = DescripcionCodigoProducto(fg_detallecompra.Row, 1)
    End If

    'se va aumentando de uno en uno el valor de las columnas del flexgrid
    'para que se vaya desplazando a traves de las celdas
    fg_detallecompra.Row = fg_detallecompra.Row - 1
    fg_detallecompra.SetFocus
    t_editar.Visible = False
End If
End Sub

Private Sub t_editar_KeyPress(KeyAscii As Integer)
If KeyAscii = 32 And t_editar.Text = "" Then
    t_editar.Locked = True
Else
    t_editar.Locked = False
End If
'determina si se ha precionado la tecla enter, estando en blanco el flexgrid
'en la fila 1 columna 0, si es asi, se procede a mostrar el formulario
If KeyAscii = 13 And fg_detallecompra.TextMatrix(fg_detallecompra.Row, 0) = "" Then
    f_listadoproductos.Show
End If

'determina si el valor del flexgrid en la fila seleccionada y en la columna 0
'es diferente de blanco (nada)
If KeyAscii = 13 And fg_detallecompra.TextMatrix(fg_detallecompra.Row, 0) <> "" Then
    
    'determina si la columna en el flexgrid es menor o igual a 6
    If fg_detallecompra.Col <= 6 Then
        
        If fg_detallecompra.Col = 0 Then
            fg_detallecompra.Col = 1
        Else
            If fg_detallecompra.Col = 1 Then
                fg_detallecompra.Col = 3
            Else
                If fg_detallecompra.Col = 3 Then
                    fg_detallecompra.Col = 4
                Else
                    If fg_detallecompra.Col = 4 Then
                        fg_detallecompra.Col = 6
                    Else
                        If fg_detallecompra.Col = 6 Then
                            fg_detallecompra.Col = 7
                        End If
                    End If
                End If
            End If
        End If
    
        'determina si el valor de la columna del flexgrid es diferente de 5
        'o es diferente de 7  ya que esas columnas no se pueden editar
        If fg_detallecompra.Col <> 5 And fg_detallecompra.Col <> 7 And fg_detallecompra.Col <> 2 Then
        
            ' se procede a dar al t_editar  todos los valores de propiedad de la
            'de la celda que tiene el enfoque
            t_editar.Left = fg_detallecompra.CellLeft + fg_detallecompra.Left
            t_editar.Top = fg_detallecompra.CellTop + fg_detallecompra.Top
            t_editar.Width = fg_detallecompra.CellWidth
            t_editar.Height = fg_detallecompra.CellHeight
            t_editar.BorderStyle = 0
            t_editar.FontName = fg_detallecompra.CellFontName
            t_editar.FontSize = fg_detallecompra.CellFontSize
            t_editar.FontBold = True
            t_editar.Visible = True ' se coloca visible el t_editar
            t_editar.SetFocus ' t_editar recibe el enfoque
            t_editar.Text = fg_detallecompra.TextMatrix(fg_detallecompra.Row, fg_detallecompra.Col)
            ' se pasa al t_editar, el valor de la celda que tiene el enfoque
        End If
    End If
    
    'determina  si el  valor de la columna del flexgrid es  = 7 y  si  el valor
    'de la celda del flexgrid en la fila 1 columna 0 es diferente de blanco (nada)
    'nota: esto se hace para no agregar filas en blanco  en el flexgrid detalle compra
    If fg_detallecompra.Col = 7 And fg_detallecompra.TextMatrix(fg_detallecompra.Row, 0) <> "" Then
        
        fg_detallecompra.Row = fg_detallecompra.Row + 1
        fg_detallecompra.Col = 0
        ' se procede a dar al t_editar  todos los valores de propiedad de la
        'de la celda que tiene el enfoque
        t_editar.Left = fg_detallecompra.CellLeft + fg_detallecompra.Left
        t_editar.Top = fg_detallecompra.CellTop + fg_detallecompra.Top
        t_editar.Width = fg_detallecompra.CellWidth
        t_editar.Height = fg_detallecompra.CellHeight
        t_editar.BorderStyle = 0
        t_editar.FontName = fg_detallecompra.CellFontName
        t_editar.FontSize = fg_detallecompra.CellFontSize
        t_editar.FontBold = True
        t_editar.Visible = True ' se coloca visible el t_editar
        t_editar.SetFocus ' t_editar recibe el enfoque
        t_editar.Text = fg_detallecompra.TextMatrix(fg_detallecompra.Row, fg_detallecompra.Col)
        ' se pasa al t_editar, el valor de la celda que tiene el enfoque
   End If
End If
End Sub

Private Sub t_editar_LostFocus()
'codigo para no dejar en blanco la celda descripcion y codigo cuando el
'usuario decide borrarla
If ValorFila <> 0 And t_editar.Text = "" And ValorColumna = 0 Then
     fg_detallecompra.TextMatrix(ValorFila, 0) = DescripcionCodigoProducto(ValorFila, 0)
End If
If t_editar.Text = "" And ValorColumna = 1 Then
    fg_detallecompra.TextMatrix(ValorFila, 1) = DescripcionCodigoProducto(ValorFila, 1)
End If
t_editar.Visible = False
End Sub
