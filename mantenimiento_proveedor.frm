VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form mantenimiento_proveedor 
   BackColor       =   &H00404040&
   Caption         =   "Proveedores"
   ClientHeight    =   8070
   ClientLeft      =   60
   ClientTop       =   1455
   ClientWidth     =   14040
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "mantenimiento_proveedor.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   8070
   ScaleWidth      =   14040
   Begin VB.PictureBox Picture5 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   870
      Left            =   3960
      ScaleHeight     =   840
      ScaleWidth      =   1755
      TabIndex        =   41
      Top             =   120
      Width           =   1785
      Begin VB.CommandButton b_eliminar 
         BackColor       =   &H00FFFFFF&
         DisabledPicture =   "mantenimiento_proveedor.frx":058A
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   900
         Left            =   -30
         Picture         =   "mantenimiento_proveedor.frx":87CC
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Eliminar producto"
         Top             =   -30
         Width           =   1815
      End
   End
   Begin VB.PictureBox Picture4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   870
      Left            =   5880
      ScaleHeight     =   840
      ScaleWidth      =   1755
      TabIndex        =   40
      Top             =   120
      Width           =   1785
      Begin VB.CommandButton b_guardar 
         BackColor       =   &H00FFFFFF&
         DisabledPicture =   "mantenimiento_proveedor.frx":10A0E
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   900
         Left            =   -30
         Picture         =   "mantenimiento_proveedor.frx":19168
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Guardar cambios"
         Top             =   -30
         UseMaskColor    =   -1  'True
         Width           =   1815
      End
   End
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   870
      Left            =   7800
      ScaleHeight     =   840
      ScaleWidth      =   1755
      TabIndex        =   39
      Top             =   120
      Width           =   1785
      Begin VB.CommandButton b_cancelar 
         BackColor       =   &H00FFFFFF&
         DisabledPicture =   "mantenimiento_proveedor.frx":218C2
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   900
         Left            =   -30
         Picture         =   "mantenimiento_proveedor.frx":2A01C
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Cancelar"
         Top             =   -30
         Width           =   1815
      End
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   870
      Left            =   2040
      ScaleHeight     =   840
      ScaleWidth      =   1755
      TabIndex        =   38
      Top             =   120
      Width           =   1785
      Begin VB.CommandButton b_modificar 
         BackColor       =   &H00FFFFFF&
         DisabledPicture =   "mantenimiento_proveedor.frx":32776
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   900
         Left            =   -30
         Picture         =   "mantenimiento_proveedor.frx":3ACC4
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Modificar Producto"
         Top             =   -30
         Width           =   1815
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   870
      Left            =   120
      ScaleHeight     =   840
      ScaleWidth      =   1725
      TabIndex        =   37
      Top             =   120
      Width           =   1755
      Begin VB.CommandButton b_agregar 
         BackColor       =   &H00FFFFFF&
         DisabledPicture =   "mantenimiento_proveedor.frx":43212
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   900
         Left            =   -30
         Picture         =   "mantenimiento_proveedor.frx":4BA74
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Agregar Nuevo producto"
         Top             =   -30
         Width           =   1815
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Caption         =   "Datos de Crédito"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      TabIndex        =   32
      Top             =   7200
      Width           =   13335
      Begin VB.PictureBox Picture9 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   7680
         ScaleHeight     =   300
         ScaleWidth      =   2115
         TabIndex        =   45
         Top             =   0
         Width           =   2145
         Begin VB.ComboBox c_tiempopago 
            Appearance      =   0  'Flat
            BackColor       =   &H00404040&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0C0C0&
            Height          =   360
            Left            =   -30
            TabIndex        =   23
            Top             =   -30
            Width           =   2175
         End
      End
      Begin VB.TextBox t_limitecredito 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   375
         Left            =   2040
         TabIndex        =   22
         Top             =   0
         Width           =   2655
      End
      Begin VB.Label Label5 
         BackColor       =   &H00404040&
         Caption         =   "Monto Límite Crédito :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   375
         Left            =   120
         TabIndex        =   36
         Top             =   0
         Width           =   1935
      End
      Begin VB.Label Label6 
         BackColor       =   &H00404040&
         Caption         =   "Tiempo de Pago:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   375
         Left            =   6000
         TabIndex        =   35
         Top             =   0
         Width           =   1935
      End
      Begin VB.Label Label10 
         BackColor       =   &H00404040&
         Caption         =   "(opcional)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   255
         Left            =   4800
         TabIndex        =   34
         Top             =   0
         Width           =   975
      End
      Begin VB.Label Label11 
         BackColor       =   &H00404040&
         Caption         =   "(opcional)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   255
         Left            =   10080
         TabIndex        =   33
         Top             =   0
         Width           =   855
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Caption         =   "Datos Personales"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Left            =   240
      TabIndex        =   11
      Top             =   4200
      Width           =   13335
      Begin VB.PictureBox Picture8 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   9360
         ScaleHeight     =   300
         ScaleWidth      =   3795
         TabIndex        =   44
         Top             =   1440
         Width           =   3825
         Begin VB.ComboBox c_vendedor 
            Appearance      =   0  'Flat
            BackColor       =   &H00404040&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0C0C0&
            Height          =   360
            ItemData        =   "mantenimiento_proveedor.frx":542D6
            Left            =   -30
            List            =   "mantenimiento_proveedor.frx":542D8
            TabIndex        =   21
            Top             =   -30
            Width           =   3855
         End
      End
      Begin VB.PictureBox Picture7 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   1680
         ScaleHeight     =   300
         ScaleWidth      =   3075
         TabIndex        =   43
         Top             =   600
         Width           =   3105
         Begin VB.ComboBox c_tipodocumento 
            Appearance      =   0  'Flat
            BackColor       =   &H00404040&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0C0C0&
            Height          =   360
            Left            =   -30
            Style           =   2  'Dropdown List
            TabIndex        =   13
            Top             =   -30
            Width           =   3135
         End
      End
      Begin VB.TextBox t_rifcedula 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   375
         Left            =   1680
         MaxLength       =   12
         TabIndex        =   14
         Top             =   1080
         Width           =   2415
      End
      Begin VB.TextBox t_nombrers 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   375
         Left            =   1680
         TabIndex        =   15
         Top             =   1560
         Width           =   5175
      End
      Begin VB.TextBox t_direccion 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   375
         Left            =   1680
         TabIndex        =   16
         Top             =   2040
         Width           =   5175
      End
      Begin VB.TextBox t_telefono1 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   375
         Left            =   9360
         TabIndex        =   17
         Top             =   0
         Width           =   2415
      End
      Begin VB.TextBox t_telefono2 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   375
         Left            =   9360
         TabIndex        =   18
         Top             =   480
         Width           =   2415
      End
      Begin VB.TextBox t_codigo 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   375
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   0
         Width           =   1695
      End
      Begin VB.TextBox t_email 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   375
         Left            =   9360
         TabIndex        =   20
         Top             =   960
         Width           =   3735
      End
      Begin VB.Label Label1 
         BackColor       =   &H00404040&
         Caption         =   "Nro Documento:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   255
         Left            =   120
         TabIndex        =   31
         Top             =   1080
         Width           =   2055
      End
      Begin VB.Label Label2 
         BackColor       =   &H00404040&
         Caption         =   "Nombre o RS:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   255
         Left            =   360
         TabIndex        =   30
         Top             =   1560
         Width           =   1935
      End
      Begin VB.Label Label3 
         BackColor       =   &H00404040&
         Caption         =   "Dirección:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   255
         Left            =   720
         TabIndex        =   29
         Top             =   2040
         Width           =   1335
      End
      Begin VB.Label Label4 
         BackColor       =   &H00404040&
         Caption         =   "Teléfono Fijo:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   375
         Left            =   7920
         TabIndex        =   28
         Top             =   0
         Width           =   1695
      End
      Begin VB.Label Label9 
         BackColor       =   &H00404040&
         Caption         =   "Teléfono Celular:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   375
         Left            =   7680
         TabIndex        =   27
         Top             =   480
         Width           =   1935
      End
      Begin VB.Label Label14 
         BackColor       =   &H00404040&
         Caption         =   "Vendedor (a):"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   375
         Left            =   8040
         TabIndex        =   26
         Top             =   1440
         Width           =   1575
      End
      Begin VB.Label Label12 
         BackColor       =   &H00404040&
         Caption         =   "Tipo Documento:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   255
         Left            =   0
         TabIndex        =   25
         Top             =   600
         Width           =   2295
      End
      Begin VB.Label Label13 
         BackColor       =   &H00404040&
         Caption         =   "Código:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   255
         Left            =   840
         TabIndex        =   24
         Top             =   0
         Width           =   1095
      End
      Begin VB.Label Label15 
         BackColor       =   &H00404040&
         Caption         =   "E-mail:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   255
         Left            =   8640
         TabIndex        =   19
         Top             =   960
         Width           =   1215
      End
   End
   Begin VB.Data cliente 
      Caption         =   "cliente"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1080
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   9240
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Caption         =   "Búsqueda de Cliente"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Left            =   240
      TabIndex        =   0
      Top             =   1200
      Width           =   13455
      Begin VB.PictureBox Picture6 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   1440
         ScaleHeight     =   300
         ScaleWidth      =   2595
         TabIndex        =   42
         Top             =   0
         Width           =   2625
         Begin VB.ComboBox c_busqueda 
            Appearance      =   0  'Flat
            BackColor       =   &H00404040&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0C0C0&
            Height          =   360
            Left            =   -30
            Style           =   2  'Dropdown List
            TabIndex        =   6
            Top             =   -30
            Width           =   2655
         End
      End
      Begin MSFlexGridLib.MSFlexGrid fg_listacliente 
         Height          =   1935
         Left            =   120
         TabIndex        =   8
         Top             =   480
         Width           =   13215
         _ExtentX        =   23310
         _ExtentY        =   3413
         _Version        =   393216
         Cols            =   11
         FixedCols       =   0
         BackColor       =   4210752
         ForeColor       =   12632256
         BackColorFixed  =   4210752
         ForeColorFixed  =   12632256
         BackColorSel    =   3026478
         ForeColorSel    =   12632256
         BackColorBkg    =   4210752
         GridColor       =   8421504
         GridColorFixed  =   8421504
         GridLinesFixed  =   1
         SelectionMode   =   1
         AllowUserResizing=   1
         Appearance      =   0
         FormatString    =   $"mantenimiento_proveedor.frx":542DA
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.TextBox t_busqueda 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   405
         Left            =   6840
         TabIndex        =   7
         Top             =   0
         Width           =   4095
      End
      Begin VB.Label Label8 
         BackColor       =   &H00404040&
         Caption         =   " Buscar:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   255
         Left            =   5280
         TabIndex        =   10
         Top             =   0
         Width           =   855
      End
      Begin VB.Label Label7 
         BackColor       =   &H00404040&
         Caption         =   "Buscar Por:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   0
         Width           =   1695
      End
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H00808080&
      Height          =   855
      Left            =   120
      Top             =   7080
      Width           =   13695
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00808080&
      BorderColor     =   &H00808080&
      Height          =   2775
      Left            =   120
      Top             =   1080
      Width           =   13695
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00808080&
      BorderColor     =   &H00808080&
      Height          =   2775
      Left            =   120
      Top             =   4080
      Width           =   13695
   End
End
Attribute VB_Name = "mantenimiento_proveedor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public BanderaProveedor As Integer
Private Sub b_agregar_Click()

Call HabilitarTextBoxProveedor
Call LimpiarTextBoxProveedor
b_agregar.Enabled = False
b_modificar.Enabled = False
b_eliminar.Enabled = False
b_guardar.Enabled = True
b_cancelar.Enabled = True
Frame1.Enabled = False
BanderaProveedor = 1


Dim TemProveedor As New ADODB.Recordset
Dim TempCodigo As Integer
Call Conn_BDaiosoft
Set TemProveedor = New ADODB.Recordset
TemProveedor.Open "SELECT codigo FROM proveedor ORDER BY codigo", Conn_Mysqldb, adOpenDynamic, adLockBatchOptimistic
If TemProveedor.BOF = False And TemProveedor.EOF = False Then
    TemProveedor.MoveLast
    TempCodigo = TemProveedor.Fields(0).Value
    TempCodigo = TempCodigo + 1
    t_codigo.Text = TempCodigo
Else
    t_codigo.Text = 1
End If

End Sub

Private Sub b_cancelar_Click()
Call DesabilitarTextboxProveedor
Call LimpiarTextBoxProveedor

b_cancelar.Enabled = False
b_guardar.Enabled = False
b_modificar.Enabled = False
b_eliminar.Enabled = False
b_agregar.Enabled = True
Frame1.Enabled = True
End Sub

Private Sub b_eliminar_Click()


    msg = MsgBox("¿Seguro desea eliminar el registro seleccionado?", vbQuestion + vbYesNo)
    If msg = vbYes Then
        Call Conn_BDaiosoft
        Conn_Mysqldb.Execute "DELETE FROM proveedor WHERE tipodocumento = '" & c_tipodocumento.Text & "' AND nrodocumento= '" & t_rifcedula.Text & "'"
        MsgBox "El registro ha sido eliminado con éxito"
        Call DesabilitarTextboxProveedor
        Call LimpiarTextBoxProveedor
        b_guardar.Enabled = False
        b_cancelar.Enabled = False
        b_agregar.Enabled = True
        Frame1.Enabled = True
        t_busqueda.Text = ""
        t_busqueda.SetFocus
        
    End If


End Sub

Private Sub b_guardar_Click()

Dim TemProveedor As New ADODB.Recordset


If t_rifcedula.Text = "" Or t_nombrers.Text = "" Or t_direccion.Text = "" Then
    MsgBox "Hay uno o más campos obligatorios vacios, por favor verifique"
Else
    msg = MsgBox("¿Está de acuerdo con la información suministrada?", vbQuestion + vbYesNo)
    If msg = vbYes Then
    
    If t_limitecredito.Text = "" Then
        t_limitecredito.Text = 0
    End If
                 
    If c_tiempopago.Text = "" Then
        c_tiempopago.Text = 0
    End If
                
    If c_vendedor.Text = "" Then
        c_vendedor.Text = 0
    End If
       
       
    If BanderaProveedor = 1 Then
        
        Call Conn_BDaiosoft
        Set TemProveedor = New ADODB.Recordset
        TemProveedor.Open "SELECT tipodocumento, nrodocumento FROM proveedor WHERE tipodocumento = '" & c_tipodocumento.Text & "' AND nrodocumento = '" & t_rifcedula.Text & "' ", Conn_Mysqldb, adOpenDynamic, adLockBatchOptimistic
            
        If TemProveedor.BOF = True And TemProveedor.EOF = True Then
           
            Call Conn_BDaiosoft
            
            Conn_Mysqldb.Execute "INSERT INTO proveedor SET codigo = " & t_codigo.Text & "," _
            & "tipodocumento = '" & c_tipodocumento.Text & "'," _
            & "nrodocumento = '" & t_rifcedula.Text & "'," _
            & "nombrers = '" & t_nombrers.Text & "'," _
            & "direccion = '" & t_direccion.Text & "'," _
            & "telefono1 = '" & t_telefono1.Text & "'," _
            & "telefono2 = '" & t_telefono2.Text & "'," _
            & "limitecredito = " & t_limitecredito.Text & "," _
            & "tiempocobro = " & Mid(c_tiempopago.Text, 1, 2) & "," _
            & "idvendedor = " & Mid(c_vendedor.Text, 1, 8) & "," _
            & "email= '" & t_email.Text & "'"
            DatosAlmacenadosProveedor = True
        Else
            MsgBox "El Nro de documento ya se encuentra registrado. Intente con otro."
        End If
            
    End If
        
        If BanderaProveedor = 2 Then
         
            Call Conn_BDaiosoft
            
            Conn_Mysqldb.Execute "UPDATE proveedor SET nombrers = '" & t_nombrers.Text & "'," _
            & "direccion = '" & t_direccion.Text & "'," _
            & "telefono1 = '" & t_telefono1.Text & "'," _
            & "telefono2 = '" & t_telefono2.Text & "'," _
            & "limitecredito = " & t_limitecredito.Text & "," _
            & "tiempocobro = " & Mid(c_tiempopago.Text, 1, 2) & "," _
            & "idvendedor = " & Mid(c_vendedor.Text, 1, 8) & "," _
            & "email= '" & t_email.Text & "' WHERE tipodocumento = '" & c_tipodocumento.Text & "' AND nrodocumento = '" & t_rifcedula.Text & "'"
            
            DatosAlmacenadosProveedor = True
         
         
        End If
        
        
        If DatosAlmacenadosProveedor Then
            Call DesabilitarTextboxProveedor
            Call LimpiarTextBoxProveedor
            b_guardar.Enabled = False
            b_cancelar.Enabled = False
            b_agregar.Enabled = True
            Frame1.Enabled = True
            t_busqueda.Text = ""
            t_busqueda.SetFocus
            MsgBox "Los datos fueron almacenados exitosamente"
        End If
    End If
End If
End Sub
Private Sub b_modificar_Click()
Call HabilitarTextBoxProveedor
c_tipodocumento.Enabled = False
t_rifcedula.Enabled = False

b_agregar.Enabled = False
b_modificar.Enabled = False
b_eliminar.Enabled = False
b_guardar.Enabled = True
b_cancelar.Enabled = True
ValorRifCedula = t_rifcedula.Text
BanderaProveedor = 2
Frame1.Enabled = False
End Sub

Private Sub c_vendedor_Click()
Dim Posicion As Integer


Posicion = InStr(c_vendedor.Text, " ")

Idvendedor = Mid(c_vendedor.Text, 1, Posicion - 1)

End Sub

Private Sub c_vendedor_KeyPress(KeyAscii As Integer)
If KeyAscii > 1 Then
        If KeyAscii <> 8 Then
            KeyAscii = 0
            Beep
        End If
    End If
End Sub

Private Sub fg_listacliente_DblClick()

Dim VarTipoDoc As String
b_modificar.Enabled = True
b_eliminar.Enabled = True
                

t_codigo.Text = fg_listacliente.TextMatrix(fg_listacliente.Row, 0)
c_tipodocumento.Text = fg_listacliente.TextMatrix(fg_listacliente.Row, 1)
t_rifcedula.Text = fg_listacliente.TextMatrix(fg_listacliente.Row, 2)
t_nombrers.Text = fg_listacliente.TextMatrix(fg_listacliente.Row, 3)
t_direccion.Text = fg_listacliente.TextMatrix(fg_listacliente.Row, 4)
t_telefono1.Text = fg_listacliente.TextMatrix(fg_listacliente.Row, 5)
t_telefono2.Text = fg_listacliente.TextMatrix(fg_listacliente.Row, 6)
t_limitecredito.Text = fg_listacliente.TextMatrix(fg_listacliente.Row, 7)
c_tiempopago.Text = fg_listacliente.TextMatrix(fg_listacliente.Row, 8)
c_vendedor.Text = fg_listacliente.TextMatrix(fg_listacliente.Row, 9)
t_email.Text = fg_listacliente.TextMatrix(fg_listacliente.Row, 10)
End Sub

Private Sub Form_Load()

'codigo para definir el tamaño del formulario
mantenimiento_proveedor.Height = 8610
mantenimiento_proveedor.Width = 14280
'codigo para posicionar el formulario en el centro de la pantala
mantenimiento_proveedor.Left = f_principal.ScaleWidth / 2 - mantenimiento_proveedor.ScaleWidth / 2
mantenimiento_proveedor.Top = f_principal.ScaleHeight / 2 - mantenimiento_proveedor.ScaleHeight / 2


Dim a As String
Dim TipoDocIdentidad As New ADODB.Recordset


'codigo para cargar el combobox tipo de documento de identidad del cliente
Call Conn_BDaiosoft
Set TipoDocIdentidad = New ADODB.Recordset
TipoDocIdentidad.Open "SELECT codigo, descripcion FROM tipo_doc_identidad", Conn_Mysqldb, adOpenDynamic, adLockBatchOptimistic
      
If TipoDocIdentidad.BOF = False And TipoDocIdentidad.EOF = False Then
    
    TipoDocIdentidad.MoveFirst
    'Do While Not TipoDocIdentidad.EOF
     For x = 1 To 3
        c_tipodocumento.AddItem TipoDocIdentidad.Fields(0) + " " + TipoDocIdentidad.Fields(1)
        TipoDocIdentidad.MoveNext
     Next x
    'Loop

End If

Call DesabilitarTextboxProveedor
c_busqueda.AddItem "Nombre o RS"
c_busqueda.AddItem "Nro Documento"
c_busqueda.Text = c_busqueda.List(0)

For x = 1 To 31
    a = x
    c_tiempopago.AddItem a + " Días"
Next x
c_tiempopago.Text = c_tiempopago.List(6)

End Sub

Private Sub t_busqueda_Change()
Dim templista As New ADODB.Recordset
If Not t_busqueda.Text = "" Then

    fg_listacliente.Rows = 1
    If fg_listacliente.Rows = 1 Then
        If c_busqueda.Text = "Nombre o RS" Then
            Call Conn_BDaiosoft
            Set templista = New ADODB.Recordset
            templista.Open "SELECT * FROM proveedor where nombrers like '%" & Trim(t_busqueda.Text) & "%'", Conn_Mysqldb, adOpenDynamic, adLockBatchOptimistic
            'Set templista = bdsistfact.OpenRecordset("SELECT * FROM cliente where trim(nombrers) like '" & Trim(t_busqueda.Text) & "*'", dbOpenDynaset)
                End If
        If c_busqueda.Text = "Nro Documento" Then
            Call Conn_BDaiosoft
            Set templista = New ADODB.Recordset
            templista.Open "SELECT * FROM proveedor where trim(nrodocumento) like '%" & Trim(t_busqueda.Text) & "%'", Conn_Mysqldb, adOpenDynamic, adLockBatchOptimistic
            'Set templista = bdsistfact.OpenRecordset("SELECT * FROM cliente where trim(rifcedula) like '" & Trim(t_busqueda.Text) & "*'", dbOpenDynaset)
        End If
        If templista.BOF = False And templista.EOF = False Then
            templista.MoveFirst
            fg_listacliente.Clear
            fg_listacliente.FormatString = "   Código   |   Tipo Documento   |Nro Documento| Nombre o Rs                    |Dirección                              |Teléfono 1          |Teléfono 2          |Límite Crédito               |Tiempo de pago|ID Vendedor|E-mail                        "
            Do While Not templista.EOF
                fg_listacliente.AddItem templista.Fields(0)
                fg_listacliente.TextMatrix(fg_listacliente.Rows - 1, 1) = templista.Fields(1).Value
                fg_listacliente.TextMatrix(fg_listacliente.Rows - 1, 2) = templista.Fields(2).Value
                fg_listacliente.TextMatrix(fg_listacliente.Rows - 1, 3) = templista.Fields(3).Value
                fg_listacliente.TextMatrix(fg_listacliente.Rows - 1, 4) = templista.Fields(4).Value
                fg_listacliente.TextMatrix(fg_listacliente.Rows - 1, 5) = templista.Fields(5).Value
                fg_listacliente.TextMatrix(fg_listacliente.Rows - 1, 6) = templista.Fields(6).Value
                fg_listacliente.TextMatrix(fg_listacliente.Rows - 1, 7) = templista.Fields(7).Value
                fg_listacliente.TextMatrix(fg_listacliente.Rows - 1, 8) = templista.Fields(8).Value
                fg_listacliente.TextMatrix(fg_listacliente.Rows - 1, 9) = templista.Fields(9).Value
                fg_listacliente.TextMatrix(fg_listacliente.Rows - 1, 10) = templista.Fields(10).Value

                templista.MoveNext

            Loop
        Else
            fg_listacliente.Clear
            fg_listacliente.FormatString = "   Código   |   Tipo Documento   |Nro Documento| Nombre o Rs                    |Dirección                              |Teléfono 1          |Teléfono 2          |Límite Crédito               |Tiempo de pago|ID Vendedor|E-mail                        "
        End If
    End If
Else
    fg_listacliente.Clear
    fg_listacliente.Rows = 1
    fg_listacliente.FormatString = "   Código   |   Tipo Documento   |Nro Documento| Nombre o Rs                    |Dirección                              |Teléfono 1          |Teléfono 2          |Límite Crédito               |Tiempo de pago|ID Vendedor|E-mail                        "
End If
End Sub

Private Sub t_busqueda_KeyPress(KeyAscii As Integer)
If KeyAscii = 39 Then
    KeyAscii = 0
End If
End Sub

Private Sub t_direccion_KeyPress(KeyAscii As Integer)
If KeyAscii = 39 Then
    KeyAscii = 0
End If
End Sub

Private Sub t_email_KeyPress(KeyAscii As Integer)
If KeyAscii = 39 Then
    KeyAscii = 0
End If
End Sub

Private Sub t_limitecredito_KeyPress(KeyAscii As Integer)
If KeyAscii < 48 Or KeyAscii > 57 Then
    If KeyAscii <> 8 And KeyAscii <> 44 And KeyAscii <> 46 Then
        KeyAscii = 0
        Beep
    End If
End If
End Sub

Private Sub t_nombrers_KeyPress(KeyAscii As Integer)
If KeyAscii = 39 Then
    KeyAscii = 0
End If
End Sub

Private Sub t_rifcedula_KeyPress(KeyAscii As Integer)
If KeyAscii < 48 Or KeyAscii > 57 Then
    If KeyAscii <> 8 And KeyAscii <> 45 Then
        KeyAscii = 0
        Beep
    End If
End If
End Sub

Private Sub t_telefono1_KeyPress(KeyAscii As Integer)
If KeyAscii < 48 Or KeyAscii > 57 Then
        If KeyAscii <> 8 And KeyAscii <> 44 And KeyAscii <> 46 Then
            KeyAscii = 0
            Beep
        End If
End If
End Sub

Private Sub t_telefono1_LostFocus()
If t_telefono1.Text <> "" Then
    'If Len(t_telefono1.Text) < 7 Then
    '    MsgBox "El número de teléfono debe tener al menos 7 dígitos"
    '    t_telefono1.SetFocus
    'End If
End If
End Sub

Private Sub t_telefono2_KeyPress(KeyAscii As Integer)
If KeyAscii < 48 Or KeyAscii > 57 Then
    If KeyAscii <> 8 And KeyAscii <> 44 And KeyAscii <> 46 Then
        KeyAscii = 0
        Beep
    End If
End If
End Sub

Private Sub t_telefono2_LostFocus()
If t_telefono2.Text <> "" Then
    'If Len(t_telefono2.Text) < 7 Then
    '    MsgBox "El número de teléfono debe tener al menos 7 digitos"
    '    t_telefono2.SetFocus
    'End If
End If
End Sub
