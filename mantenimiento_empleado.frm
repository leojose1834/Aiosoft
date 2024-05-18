VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form mantenimiento_empleado 
   BackColor       =   &H00404040&
   Caption         =   "Empleados"
   ClientHeight    =   9825
   ClientLeft      =   60
   ClientTop       =   420
   ClientWidth     =   14100
   Icon            =   "mantenimiento_empleado.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   9825
   ScaleWidth      =   14100
   Begin VB.PictureBox Picture8 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   7320
      ScaleHeight     =   825
      ScaleWidth      =   1725
      TabIndex        =   47
      Top             =   120
      Width           =   1755
      Begin VB.CommandButton b_cancelar 
         DisabledPicture =   "mantenimiento_empleado.frx":058A
         Height          =   900
         Left            =   -30
         Picture         =   "mantenimiento_empleado.frx":8CE4
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   -30
         Width           =   1785
      End
   End
   Begin VB.PictureBox Picture7 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   5520
      ScaleHeight     =   825
      ScaleWidth      =   1725
      TabIndex        =   46
      Top             =   120
      Width           =   1755
      Begin VB.CommandButton b_guardar 
         DisabledPicture =   "mantenimiento_empleado.frx":1143E
         Height          =   900
         Left            =   -30
         Picture         =   "mantenimiento_empleado.frx":19B98
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   -30
         Width           =   1785
      End
   End
   Begin VB.PictureBox Picture6 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   3720
      ScaleHeight     =   825
      ScaleWidth      =   1725
      TabIndex        =   45
      Top             =   120
      Width           =   1755
      Begin VB.CommandButton b_eliminar 
         Appearance      =   0  'Flat
         DisabledPicture =   "mantenimiento_empleado.frx":222F2
         Height          =   900
         Left            =   -30
         Picture         =   "mantenimiento_empleado.frx":2A534
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   -30
         Width           =   1785
      End
   End
   Begin VB.PictureBox Picture5 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   1920
      ScaleHeight     =   825
      ScaleWidth      =   1725
      TabIndex        =   44
      Top             =   120
      Width           =   1755
      Begin VB.CommandButton b_modificar 
         DisabledPicture =   "mantenimiento_empleado.frx":32776
         Height          =   900
         Left            =   -30
         Picture         =   "mantenimiento_empleado.frx":3ACC4
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   -30
         Width           =   1785
      End
   End
   Begin VB.PictureBox Picture4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   870
      Left            =   120
      ScaleHeight     =   840
      ScaleWidth      =   1725
      TabIndex        =   43
      Top             =   120
      Width           =   1755
      Begin VB.CommandButton b_nuevo 
         DisabledPicture =   "mantenimiento_empleado.frx":43212
         Height          =   900
         Left            =   -30
         Picture         =   "mantenimiento_empleado.frx":4BA74
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   -30
         Width           =   1785
      End
   End
   Begin VB.Frame fra_permisosoperador 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Caption         =   "Permisos de Operador"
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
      TabIndex        =   34
      Top             =   6960
      Width           =   13455
      Begin VB.TextBox t_tipoconcepto 
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
         MaxLength       =   100
         TabIndex        =   22
         Top             =   1080
         Width           =   3255
      End
      Begin VB.PictureBox Picture14 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   1680
         ScaleHeight     =   300
         ScaleWidth      =   3225
         TabIndex        =   57
         Top             =   1560
         Width           =   3255
         Begin VB.ComboBox c_tipovalor 
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
            TabIndex        =   23
            Top             =   -30
            Width           =   3285
         End
      End
      Begin VB.TextBox t_concepto 
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
         MaxLength       =   100
         TabIndex        =   21
         Top             =   600
         Width           =   3255
      End
      Begin VB.TextBox t_valorconcepto 
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
         TabIndex        =   24
         Top             =   2040
         Width           =   3255
      End
      Begin VB.PictureBox Picture10 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         ForeColor       =   &H80000008&
         Height          =   500
         Left            =   0
         ScaleHeight     =   465
         ScaleWidth      =   1140
         TabIndex        =   53
         Top             =   0
         Width           =   1170
         Begin VB.CommandButton b_buscar 
            BackColor       =   &H002E2E2E&
            DisabledPicture =   "mantenimiento_empleado.frx":542D6
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   670
            Left            =   -30
            Picture         =   "mantenimiento_empleado.frx":56C58
            Style           =   1  'Graphical
            TabIndex        =   17
            Top             =   -60
            Width           =   1200
         End
      End
      Begin VB.PictureBox Picture11 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         ForeColor       =   &H80000008&
         Height          =   500
         Left            =   1260
         ScaleHeight     =   465
         ScaleWidth      =   1140
         TabIndex        =   52
         Top             =   0
         Width           =   1170
         Begin VB.CommandButton b_agregar 
            BackColor       =   &H002E2E2E&
            DisabledPicture =   "mantenimiento_empleado.frx":595DA
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   670
            Left            =   -30
            Picture         =   "mantenimiento_empleado.frx":61E3C
            Style           =   1  'Graphical
            TabIndex        =   18
            Top             =   -30
            Width           =   1200
         End
      End
      Begin VB.PictureBox Picture12 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         ForeColor       =   &H80000008&
         Height          =   500
         Left            =   3840
         ScaleHeight     =   465
         ScaleWidth      =   1140
         TabIndex        =   51
         Top             =   0
         Width           =   1170
         Begin VB.CommandButton b_eliminar2 
            BackColor       =   &H002E2E2E&
            DisabledPicture =   "mantenimiento_empleado.frx":6A69E
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   670
            Left            =   -30
            Picture         =   "mantenimiento_empleado.frx":6F1E0
            Style           =   1  'Graphical
            TabIndex        =   20
            Top             =   -30
            Width           =   1200
         End
      End
      Begin VB.PictureBox Picture13 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         ForeColor       =   &H80000008&
         Height          =   500
         Left            =   2550
         ScaleHeight     =   465
         ScaleWidth      =   1140
         TabIndex        =   50
         Top             =   0
         Width           =   1170
         Begin VB.CommandButton b_modificar2 
            BackColor       =   &H002E2E2E&
            DisabledPicture =   "mantenimiento_empleado.frx":77A42
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   670
            Left            =   -30
            Picture         =   "mantenimiento_empleado.frx":802A4
            Style           =   1  'Graphical
            TabIndex        =   19
            Top             =   -30
            Width           =   1200
         End
      End
      Begin MSFlexGridLib.MSFlexGrid fg_concepto 
         Height          =   2535
         Left            =   5160
         TabIndex        =   25
         Top             =   0
         Width           =   8295
         _ExtentX        =   14631
         _ExtentY        =   4471
         _Version        =   393216
         Rows            =   1
         Cols            =   5
         FixedCols       =   0
         BackColor       =   4210752
         ForeColor       =   12632256
         BackColorFixed  =   2434341
         ForeColorFixed  =   12632256
         BackColorSel    =   12632064
         ForeColorSel    =   4210752
         BackColorBkg    =   4210752
         GridColor       =   8421504
         GridColorFixed  =   8421504
         GridLinesFixed  =   1
         SelectionMode   =   1
         AllowUserResizing=   1
         Appearance      =   0
         FormatString    =   "ID|Descripción                                                          |Tipo Concepto|Tipo Valor|Valor                  "
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
      Begin VB.Label Label17 
         BackColor       =   &H00404040&
         Caption         =   "Tipo Concepto:"
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
         TabIndex        =   59
         Top             =   1080
         Width           =   1935
      End
      Begin VB.Label Label16 
         BackColor       =   &H00404040&
         Caption         =   "Concepto:"
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
         TabIndex        =   56
         Top             =   600
         Width           =   1935
      End
      Begin VB.Label Label14 
         BackColor       =   &H00404040&
         Caption         =   "Tipo de Valor:"
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
         TabIndex        =   55
         Top             =   1560
         Width           =   2655
      End
      Begin VB.Label Label9 
         BackColor       =   &H00404040&
         Caption         =   "Valor Concepto:"
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
         TabIndex        =   54
         Top             =   2040
         Width           =   1815
      End
   End
   Begin VB.Frame fra_datosoperador 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Caption         =   "Datos de Operador"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   240
      TabIndex        =   26
      Top             =   4080
      Width           =   13335
      Begin VB.PictureBox Picture9 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   11040
         ScaleHeight     =   300
         ScaleWidth      =   2115
         TabIndex        =   49
         Top             =   1320
         Width           =   2145
         Begin VB.ComboBox c_ciclopago 
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
            TabIndex        =   16
            Top             =   -30
            Width           =   2175
         End
      End
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   11040
         ScaleHeight     =   300
         ScaleWidth      =   2115
         TabIndex        =   40
         Top             =   720
         Width           =   2145
         Begin VB.ComboBox c_perfil 
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
            TabIndex        =   15
            Top             =   -30
            Width           =   2175
         End
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   2880
         ScaleHeight     =   300
         ScaleWidth      =   2835
         TabIndex        =   39
         Top             =   90
         Width           =   2865
         Begin VB.ComboBox c_tipodoc 
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
            TabIndex        =   9
            Top             =   -30
            Width           =   2895
         End
      End
      Begin VB.TextBox t_nombres 
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
         Left            =   2880
         MaxLength       =   30
         TabIndex        =   12
         Top             =   1560
         Width           =   3615
      End
      Begin VB.TextBox t_telefono 
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
         Left            =   11040
         MaxLength       =   19
         TabIndex        =   14
         Top             =   120
         Width           =   2175
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
         Left            =   2880
         MaxLength       =   100
         TabIndex        =   13
         Top             =   2040
         Width           =   5775
      End
      Begin VB.TextBox t_apellidos 
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
         Left            =   2880
         MaxLength       =   30
         TabIndex        =   11
         Top             =   1080
         Width           =   3615
      End
      Begin VB.TextBox t_nrodoc 
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
         Left            =   2880
         MaxLength       =   15
         TabIndex        =   10
         Top             =   600
         Width           =   2895
      End
      Begin VB.Label Label8 
         BackColor       =   &H00404040&
         Caption         =   "Ciclo de pago:"
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
         Left            =   9480
         TabIndex        =   48
         Top             =   1320
         Width           =   1455
      End
      Begin VB.Label Label15 
         BackColor       =   &H00404040&
         Caption         =   "Tipo documento:"
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
         Left            =   600
         TabIndex        =   36
         Top             =   120
         Width           =   1575
      End
      Begin VB.Label Label12 
         BackColor       =   &H00404040&
         Caption         =   "Nombres:"
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
         Left            =   600
         TabIndex        =   35
         Top             =   1560
         Width           =   1815
      End
      Begin VB.Label Label10 
         BackColor       =   &H00404040&
         Caption         =   "Cargo:"
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
         Left            =   9840
         TabIndex        =   33
         Top             =   720
         Width           =   615
      End
      Begin VB.Label Label7 
         BackColor       =   &H00404040&
         Caption         =   "Teléfono:"
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
         Left            =   9480
         TabIndex        =   32
         Top             =   120
         Width           =   1935
      End
      Begin VB.Label Label6 
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
         Height          =   375
         Left            =   600
         TabIndex        =   31
         Top             =   2040
         Width           =   1935
      End
      Begin VB.Label Label5 
         BackColor       =   &H00404040&
         Caption         =   "Apellidos:"
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
         Left            =   600
         TabIndex        =   30
         Top             =   1080
         Width           =   2655
      End
      Begin VB.Label Label4 
         BackColor       =   &H00404040&
         Caption         =   "Nro documento:"
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
         Left            =   600
         TabIndex        =   29
         Top             =   600
         Width           =   1935
      End
   End
   Begin VB.Frame fra_buscaroperador 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Caption         =   "Búsqueda de Operador"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   240
      TabIndex        =   0
      Top             =   1320
      Width           =   13455
      Begin VB.PictureBox Picture3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   1440
         ScaleHeight     =   300
         ScaleWidth      =   2355
         TabIndex        =   41
         Top             =   240
         Width           =   2385
         Begin VB.ComboBox c_buscar 
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
            Width           =   2415
         End
      End
      Begin VB.TextBox t_buscar 
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
         Left            =   5640
         TabIndex        =   7
         Top             =   240
         Width           =   2415
      End
      Begin MSFlexGridLib.MSFlexGrid fg_operador 
         Height          =   1575
         Left            =   0
         TabIndex        =   8
         Top             =   720
         Width           =   13455
         _ExtentX        =   23733
         _ExtentY        =   2778
         _Version        =   393216
         Rows            =   1
         Cols            =   9
         FixedCols       =   0
         BackColor       =   4210752
         ForeColor       =   12632256
         BackColorFixed  =   2434341
         ForeColorFixed  =   12632256
         BackColorSel    =   12632064
         ForeColorSel    =   4210752
         BackColorBkg    =   4210752
         GridColor       =   8421504
         GridColorFixed  =   8421504
         GridLinesFixed  =   1
         SelectionMode   =   1
         AllowUserResizing=   1
         Appearance      =   0
         FormatString    =   $"mantenimiento_empleado.frx":88B06
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
      Begin VB.Label Label3 
         BackColor       =   &H00404040&
         Caption         =   "Buscar:"
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
         Left            =   4800
         TabIndex        =   28
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label2 
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
         Height          =   375
         Left            =   360
         TabIndex        =   27
         Top             =   240
         Width           =   2295
      End
   End
   Begin VB.Label l_idconcepto 
      Caption         =   "l_idconcepto"
      Height          =   255
      Left            =   1680
      TabIndex        =   58
      Top             =   9600
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label Label13 
      BackColor       =   &H00404040&
      Caption         =   "Buscar Empleado."
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
      Left            =   240
      TabIndex        =   42
      Top             =   1080
      Width           =   1695
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H00808080&
      Height          =   2535
      Left            =   120
      Top             =   1200
      Width           =   13695
   End
   Begin VB.Label Label1 
      BackColor       =   &H00404040&
      Caption         =   "Datos Empleado."
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
      TabIndex        =   38
      Top             =   3840
      Width           =   1935
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00808080&
      Height          =   2655
      Left            =   120
      Top             =   3960
      Width           =   13695
   End
   Begin VB.Label Label11 
      BackColor       =   &H00404040&
      Caption         =   "Conceptos."
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
      TabIndex        =   37
      Top             =   6720
      Width           =   1095
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00808080&
      Height          =   2775
      Left            =   120
      Top             =   6840
      Width           =   13695
   End
End
Attribute VB_Name = "mantenimiento_empleado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub b_agregar_Click()
If t_concepto.Text <> "" Then
    fg_concepto.AddItem l_idconcepto.Caption
    fg_concepto.TextMatrix(fg_concepto.Rows - 1, 1) = t_concepto.Text
    fg_concepto.TextMatrix(fg_concepto.Rows - 1, 2) = t_tipoconcepto.Text
    fg_concepto.TextMatrix(fg_concepto.Rows - 1, 3) = c_tipovalor.Text
    fg_concepto.TextMatrix(fg_concepto.Rows - 1, 4) = t_valorconcepto.Text
    
    t_concepto.Text = ""
    t_tipoconcepto.Text = ""
    c_tipovalor.Text = c_tipovalor.List(0)
    t_valorconcepto.Text = ""
    b_agregar.Enabled = False
    b_buscar.SetFocus
End If
End Sub

Private Sub b_buscar_Click()
f_dialogconcepto.Show vbModal
End Sub

Private Sub b_cancelar_Click()
'habilita y desabilita los botones correspondientes
b_modificar.Enabled = False
b_nuevo.Enabled = True
b_eliminar.Enabled = False
b_cancelar.Enabled = False
b_guardar.Enabled = False

'habilita y desabilita frames
fra_buscaroperador.Enabled = True
fra_datosoperador.Enabled = False
fra_permisosoperador.Enabled = False

'limpa las cajas de texto
Call LimpiarTextBoxEmpleado


End Sub

Private Sub b_eliminar_Click()
msg = MsgBox("¿Seguro desea elimar el registro seleccionado?", vbQuestion + vbYesNo)
    If msg = vbYes Then
        Call Conn_BDaiosoft
        Conn_Mysqldb.Execute "DELETE FROM empleado WHERE tipodoc = '" & c_tipodoc.Text & "' AND nrodoc= '" & t_nrodoc.Text & "'"
        MsgBox "El registro ha sido eliminado con éxito"
        Call LimpiarTextBoxEmpleado
        b_eliminar.Enabled = False
        b_modificar.Enabled = False
        t_buscar.Text = ""
    End If
End Sub

Private Sub b_eliminar2_Click()
If fg_concepto.Rows = 2 Then
    fg_concepto.Rows = 1
End If
If fg_concepto.Rows > 2 Then
    fg_concepto.RemoveItem (fg_concepto.Row)
    b_buscar.SetFocus
End If

End Sub

Private Sub b_guardar_Click()
Dim PermisosSeleccionados As Boolean
'determina si  exite algun campo en blanco antes de guardar los datos
If t_nrodoc.Text = "" Or t_apellidos.Text = "" Or t_nombres.Text = "" Or t_direccion.Text = "" Or t_telefono.Text = "" Then
    MsgBox "No puede dejar ningún campo en blanco"
Else
    
        Call GuardarCambiosEmpleado
        
    
End If

End Sub

Private Sub b_modificar_Click()
'da un valor 2 a la variable BanderaOperador  para indicar que se desea
'editar  el registro actual del operador operador
BanderaEmpleado = 2

'habilita y desabilita los botones correspondientes
b_modificar.Enabled = False
b_nuevo.Enabled = False
b_eliminar.Enabled = False
b_cancelar.Enabled = True
b_guardar.Enabled = True
c_tipodoc.Locked = True
t_nrodoc.Locked = True

'habilita y desabilita frames
If c_perfil.Text = "Administrador" Then
    fra_permisosoperador.Enabled = True
    fra_buscaroperador.Enabled = False
    fra_datosoperador.Enabled = True
Else
    fra_buscaroperador.Enabled = False
    fra_datosoperador.Enabled = True
    fra_permisosoperador.Enabled = True
End If

' la variable ValorCedulaOperador toma el valor del campo cedula para
'para utilizarlo a la hora de modificar los datos del registro actual
ValorCedulaEmpleado = t_nrodoc.Text
'ValorLoginOperador = t_login.Text
End Sub

Private Sub b_modificar2_Click()
If fg_concepto.Rows > 1 Then
    f_dialogmodificarconcepto.Show vbModal
End If
End Sub

Private Sub b_nuevo_Click()

    fra_buscaroperador.Enabled = False
    fra_datosoperador.Enabled = True
    fra_permisosoperador.Enabled = True
    b_nuevo.Enabled = False
    b_cancelar.Enabled = True
    b_modificar.Enabled = False
    b_guardar.Enabled = True
    b_eliminar.Enabled = False
    Call LimpiarTextBoxEmpleado
    c_tipodoc.Locked = False
    t_nrodoc.Locked = False
    
    'da un valor 1 a la variable BanderaOperador  para indicar que se desea
    'crear  un registro para nuevo operador
    BanderaEmpleado = 1

End Sub

Private Sub c_perfil_Click()

If c_perfil.Text = "Cajero" Then
    For X = 0 To 25
        If X = 5 Or X = 7 Then
            l_permisos.Selected(X) = True
        Else
            l_permisos.Selected(X) = False
        End If
    Next X
End If
If c_perfil.Text = "Ventas" Then
    For X = 0 To 25
        If X = 5 Or X = 6 Then
            l_permisos.Selected(X) = True
        Else
            l_permisos.Selected(X) = False
        End If
    Next X
End If
If c_perfil.Text = "Compras" Then
    For X = 0 To 25
        If X = 13 Or X = 14 Or X = 15 Then
            l_permisos.Selected(X) = True
        Else
            l_permisos.Selected(X) = False
        End If
    Next X
End If
If c_perfil.Text = "Cobranza" Then
    For X = 0 To 25
        If X = 5 Or X = 9 Or X = 10 Or X = 11 Or X = 12 Then
            l_permisos.Selected(X) = True
        Else
            l_permisos.Selected(X) = False
        End If
    Next X
End If
If c_perfil.Text = "Administrador" Then
    For X = 0 To 28
        l_permisos.Selected(X) = True
    Next X
End If
If c_perfil.Text = "Personalizado" Then
    For X = 0 To 28
        l_permisos.Selected(X) = False
    Next X
End If



End Sub




Private Sub fg_concepto_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    f_dialogmodificarconcepto.Show vbModal
End If
End Sub

Private Sub fg_operador_Click()
If fg_operador.Rows > 1 Then
Dim TempEmpleadoConcepto As ADODB.Recordset

VarIdEmpleado = fg_operador.TextMatrix(fg_operador.Row, 0)
c_tipodoc.Text = fg_operador.TextMatrix(fg_operador.Row, 1)
t_nrodoc.Text = fg_operador.TextMatrix(fg_operador.Row, 2)
t_apellidos.Text = fg_operador.TextMatrix(fg_operador.Row, 3)
t_nombres.Text = fg_operador.TextMatrix(fg_operador.Row, 4)
t_direccion.Text = fg_operador.TextMatrix(fg_operador.Row, 5)
t_telefono.Text = fg_operador.TextMatrix(fg_operador.Row, 6)
c_perfil.Text = fg_operador.TextMatrix(fg_operador.Row, 7)

fg_concepto.Rows = 1
Call Conn_BDaiosoft
Set TempEmpleadoConcepto = New ADODB.Recordset
TempEmpleadoConcepto.Open "SELECT empleado_concepto.idconcepto, descripcion, tipo, tipovalor, valor  FROM empleado_concepto INNER JOIN concepto ON empleado_concepto.idconcepto= concepto.idconcepto WHERE empleado_concepto.idempleado=" & VarIdEmpleado & "", Conn_Mysqldb, adOpenDynamic, adLockBatchOptimistic
If TempEmpleadoConcepto.BOF = False And TempEmpleadoConcepto.EOF = False Then
    TempEmpleadoConcepto.MoveFirst
    Do While TempEmpleadoConcepto.EOF = False
        fg_concepto.AddItem TempEmpleadoConcepto.Fields(0)
        fg_concepto.TextMatrix(fg_concepto.Rows - 1, 1) = TempEmpleadoConcepto.Fields(1)
        fg_concepto.TextMatrix(fg_concepto.Rows - 1, 2) = TempEmpleadoConcepto.Fields(2)
        fg_concepto.TextMatrix(fg_concepto.Rows - 1, 3) = TempEmpleadoConcepto.Fields(3)
        fg_concepto.TextMatrix(fg_concepto.Rows - 1, 4) = TempEmpleadoConcepto.Fields(4)
        TempEmpleadoConcepto.MoveNext
    Loop
    
End If



b_modificar.Enabled = True
b_eliminar.Enabled = True
End If
End Sub

Private Sub Form_Load()
'carga el combo buscar
c_buscar.AddItem "Nro Documento"
c_buscar.AddItem "Nombre"
c_buscar.AddItem "Apellido"
c_buscar.Text = c_buscar.List(1)

'carga el combo tipo valor
c_tipovalor.AddItem "PORCENTAJE"
c_tipovalor.AddItem "MONTO"
c_tipovalor.Text = c_tipovalor.List(0)


'establece el tamaño del formulario
mantenimiento_empleado.Width = 14205
mantenimiento_empleado.Height = 10395
mantenimiento_empleado.Left = 2000

'habilita y desabilita los botones correspondientes
b_modificar.Enabled = False
b_nuevo.Enabled = True
b_eliminar.Enabled = False
b_cancelar.Enabled = False
b_guardar.Enabled = False

'habilita y desabilita frames
fra_buscaroperador.Enabled = True
fra_datosoperador.Enabled = False
fra_permisosoperador.Enabled = False

'codigo para cargar el combobox c_perfil

c_perfil.AddItem "VENDEDOR"
c_perfil.AddItem "TECNICO"
c_perfil.Text = c_perfil.List(1)

c_ciclopago.AddItem "SEMANAL"
c_ciclopago.AddItem "QUINCENAL"
c_ciclopago.AddItem "MENSUAL"
c_ciclopago.Text = c_ciclopago.List(0)


'codigo para cargar el combobox tipo de documento del operador
Call Conn_BDaiosoft
Set TipoDocIdentidad = New ADODB.Recordset
TipoDocIdentidad.Open "SELECT codigo, descripcion FROM tipo_doc_identidad", Conn_Mysqldb, adOpenDynamic, adLockBatchOptimistic
If TipoDocIdentidad.BOF = False And TipoDocIdentidad.EOF = False Then
    TipoDocIdentidad.MoveFirst
    Do While Not TipoDocIdentidad.EOF
        c_tipodoc.AddItem TipoDocIdentidad.Fields(0) + " " + TipoDocIdentidad.Fields(1)
        TipoDocIdentidad.MoveNext
    Loop
End If
c_tipodoc.ListIndex = 0
End Sub

Private Sub t_apellidos_KeyPress(KeyAscii As Integer)
If KeyAscii = 39 Then
    KeyAscii = 0
End If
If KeyAscii = 13 Then
    t_nombres.SetFocus
End If
End Sub


Private Sub t_buscar_Change()
Dim TempEmpleado As ADODB.Recordset
Call Conn_BDaiosoft
Set TempEmpleado = New ADODB.Recordset

If Not t_buscar.Text = "" Then
    fg_operador.Rows = 1
    If fg_operador.Rows = 1 Then
        If c_buscar.Text = "Nombre" Then
            TempEmpleado.Open "SELECT * FROM empleado WHERE  nombres LIKE '%" & t_buscar.Text & "%'", Conn_Mysqldb, adOpenDynamic, adLockBatchOptimistic
        End If
        If c_buscar.Text = "Apellido" Then
            TempEmpleado.Open "SELECT * FROM empleado WHERE  apellidos LIKE '%" & t_buscar.Text & "%'", Conn_Mysqldb, adOpenDynamic, adLockBatchOptimistic
        End If
        If c_buscar.Text = "Nro Documento" Then
             TempEmpleado.Open "SELECT * FROM empleado WHERE  nrodoc LIKE '%" & t_buscar.Text & "%'", Conn_Mysqldb, adOpenDynamic, adLockBatchOptimistic
        End If
        If TempEmpleado.BOF = False And TempEmpleado.EOF = False Then
            TempEmpleado.MoveFirst
            fg_operador.Clear
            fg_operador.FormatString = "ID|Tipo Doc   |Nro Documento         | Apellidos                    |Nombres                              |Dirección          |Teléfono          |Cargo           |Ciclo pago"
            Do While TempEmpleado.EOF = False
                fg_operador.AddItem TempEmpleado.Fields(0)
                fg_operador.TextMatrix(fg_operador.Rows - 1, 1) = TempEmpleado.Fields(1).Value
                fg_operador.TextMatrix(fg_operador.Rows - 1, 2) = TempEmpleado.Fields(2).Value
                fg_operador.TextMatrix(fg_operador.Rows - 1, 3) = TempEmpleado.Fields(3).Value
                fg_operador.TextMatrix(fg_operador.Rows - 1, 4) = TempEmpleado.Fields(4).Value
                fg_operador.TextMatrix(fg_operador.Rows - 1, 5) = TempEmpleado.Fields(5).Value
                fg_operador.TextMatrix(fg_operador.Rows - 1, 6) = TempEmpleado.Fields(6).Value
                fg_operador.TextMatrix(fg_operador.Rows - 1, 7) = TempEmpleado.Fields(7).Value
                fg_operador.TextMatrix(fg_operador.Rows - 1, 8) = TempEmpleado.Fields(8).Value
 
                TempEmpleado.MoveNext
            Loop
        Else
            fg_operador.Clear
            fg_operador.FormatString = "ID|Tipo Doc   |Nro Documento         | Apellidos                    |Nombres                              |Dirección          |Teléfono          |Cargo           |Ciclo pago"
        End If
    End If
Else
    fg_operador.Clear
    fg_operador.Rows = 1
    fg_operador.FormatString = "ID|Tipo Doc   |Nro Documento         | Apellidos                    |Nombres                              |Dirección          |Teléfono          |Cargo           |Ciclo pago"
End If
End Sub

Private Sub t_buscar_KeyPress(KeyAscii As Integer)
If KeyAscii = 39 Then
    KeyAscii = 0
End If
End Sub

Private Sub t_nrodoc_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    t_apellidos.SetFocus
Else
    If KeyAscii < 48 Or KeyAscii > 57 Then
        If KeyAscii <> 8 And KeyAscii <> 44 And KeyAscii <> 46 Then
            KeyAscii = 0
            Beep
        End If
    End If
End If
End Sub

Private Sub t_codarea_Change()
If Len(t_codarea.Text) = 4 And fra_datosoperador.Enabled = True Then
    t_telefono.SetFocus
End If
End Sub

Private Sub t_codarea_KeyPress(KeyAscii As Integer)
If KeyAscii < 48 Or KeyAscii > 57 Then
        If KeyAscii <> 8 And KeyAscii <> 44 And KeyAscii <> 46 Then
            KeyAscii = 0
            Beep
        End If
    End If
End Sub

Private Sub t_direccion_KeyPress(KeyAscii As Integer)
If KeyAscii = 39 Then
    KeyAscii = 0
End If
If KeyAscii = 13 Then
    t_telefono.SetFocus
End If
End Sub


Private Sub t_login_KeyPress(KeyAscii As Integer)
If KeyAscii = 39 Then
    KeyAscii = 0
End If
If KeyAscii = 13 Then
    t_password.SetFocus
End If
End Sub


Private Sub t_nombres_KeyPress(KeyAscii As Integer)
If KeyAscii = 39 Then
    KeyAscii = 0
End If
If KeyAscii = 13 Then
    t_direccion.SetFocus
End If
End Sub


Private Sub t_password_KeyPress(KeyAscii As Integer)
If KeyAscii = 39 Then
    KeyAscii = 0
End If
If KeyAscii = 13 And fra_permisosoperador.Enabled = False Then
    b_guardar.SetFocus
End If
End Sub


Private Sub t_telefono_KeyPress(KeyAscii As Integer)

    If KeyAscii < 48 Or KeyAscii > 57 Then
        If KeyAscii <> 8 And KeyAscii <> 44 And KeyAscii <> 46 Then
            KeyAscii = 0
            Beep
        End If
    End If

End Sub


Private Sub t_valorconcepto_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    b_agregar.Value = True
End If
End Sub
