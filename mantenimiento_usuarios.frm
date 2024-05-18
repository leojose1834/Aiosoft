VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form mantenimiento_usuarios 
   BackColor       =   &H00404040&
   Caption         =   "Operadores del Sistema"
   ClientHeight    =   9375
   ClientLeft      =   60
   ClientTop       =   420
   ClientWidth     =   14370
   Icon            =   "mantenimiento_usuarios.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   9375
   ScaleWidth      =   14370
   Begin VB.PictureBox Picture8 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   7320
      ScaleHeight     =   825
      ScaleWidth      =   1725
      TabIndex        =   41
      Top             =   120
      Width           =   1755
      Begin VB.CommandButton b_cancelar 
         DisabledPicture =   "mantenimiento_usuarios.frx":058A
         Height          =   900
         Left            =   -30
         Picture         =   "mantenimiento_usuarios.frx":8CE4
         Style           =   1  'Graphical
         TabIndex        =   42
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
      TabIndex        =   39
      Top             =   120
      Width           =   1755
      Begin VB.CommandButton b_guardar 
         DisabledPicture =   "mantenimiento_usuarios.frx":1143E
         Height          =   900
         Left            =   -30
         Picture         =   "mantenimiento_usuarios.frx":19B98
         Style           =   1  'Graphical
         TabIndex        =   40
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
      TabIndex        =   37
      Top             =   120
      Width           =   1755
      Begin VB.CommandButton b_eliminar 
         Appearance      =   0  'Flat
         DisabledPicture =   "mantenimiento_usuarios.frx":222F2
         Height          =   900
         Left            =   -30
         Picture         =   "mantenimiento_usuarios.frx":2A534
         Style           =   1  'Graphical
         TabIndex        =   38
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
      TabIndex        =   35
      Top             =   120
      Width           =   1755
      Begin VB.CommandButton b_modificar 
         DisabledPicture =   "mantenimiento_usuarios.frx":32776
         Height          =   900
         Left            =   -30
         Picture         =   "mantenimiento_usuarios.frx":3ACC4
         Style           =   1  'Graphical
         TabIndex        =   36
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
      TabIndex        =   33
      Top             =   120
      Width           =   1755
      Begin VB.CommandButton b_nuevo 
         DisabledPicture =   "mantenimiento_usuarios.frx":43212
         Height          =   900
         Left            =   -30
         Picture         =   "mantenimiento_usuarios.frx":4BA74
         Style           =   1  'Graphical
         TabIndex        =   34
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
      Height          =   2415
      Left            =   240
      TabIndex        =   21
      Top             =   6960
      Width           =   13335
      Begin VB.ListBox l_permisos 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         Columns         =   2
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
         Height          =   2190
         Left            =   120
         Style           =   1  'Checkbox
         TabIndex        =   10
         Top             =   120
         Width           =   13215
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
      TabIndex        =   11
      Top             =   4080
      Width           =   13335
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   11040
         ScaleHeight     =   300
         ScaleWidth      =   2115
         TabIndex        =   28
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
            TabIndex        =   29
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
         TabIndex        =   26
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
            TabIndex        =   27
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
         TabIndex        =   5
         Top             =   1560
         Width           =   3615
      End
      Begin VB.TextBox t_password 
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
         IMEMode         =   3  'DISABLE
         Left            =   11040
         MaxLength       =   15
         PasswordChar    =   "*"
         TabIndex        =   9
         Top             =   1680
         Width           =   2175
      End
      Begin VB.TextBox t_login 
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
         MaxLength       =   15
         TabIndex        =   8
         Top             =   1200
         Width           =   2175
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
         TabIndex        =   7
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
         TabIndex        =   6
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
         TabIndex        =   4
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
         TabIndex        =   3
         Top             =   600
         Width           =   2895
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
         TabIndex        =   23
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
         TabIndex        =   22
         Top             =   1560
         Width           =   1815
      End
      Begin VB.Label Label10 
         BackColor       =   &H00404040&
         Caption         =   "Perfil:"
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
         TabIndex        =   20
         Top             =   720
         Width           =   615
      End
      Begin VB.Label Label9 
         BackColor       =   &H00404040&
         Caption         =   "Password:"
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
         TabIndex        =   19
         Top             =   1680
         Width           =   1095
      End
      Begin VB.Label Label8 
         BackColor       =   &H00404040&
         Caption         =   "Login:"
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
         TabIndex        =   18
         Top             =   1200
         Width           =   1095
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
         TabIndex        =   17
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
         TabIndex        =   16
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
         TabIndex        =   15
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
         TabIndex        =   14
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
         TabIndex        =   30
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
            TabIndex        =   31
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
         TabIndex        =   1
         Top             =   240
         Width           =   2415
      End
      Begin MSFlexGridLib.MSFlexGrid fg_operador 
         Height          =   1575
         Left            =   0
         TabIndex        =   2
         Top             =   720
         Width           =   13455
         _ExtentX        =   23733
         _ExtentY        =   2778
         _Version        =   393216
         Rows            =   1
         Cols            =   10
         FixedCols       =   0
         BackColor       =   4210752
         ForeColor       =   12632256
         BackColorFixed  =   4210752
         ForeColorFixed  =   12632256
         BackColorSel    =   2434341
         ForeColorSel    =   12632256
         BackColorBkg    =   4210752
         GridColor       =   8421504
         GridColorFixed  =   8421504
         GridLinesFixed  =   1
         SelectionMode   =   1
         AllowUserResizing=   1
         Appearance      =   0
         FormatString    =   $"mantenimiento_usuarios.frx":542D6
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
         TabIndex        =   13
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
         TabIndex        =   12
         Top             =   240
         Width           =   2295
      End
   End
   Begin VB.Label Label13 
      BackColor       =   &H00404040&
      Caption         =   "Buscar Operador."
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
      TabIndex        =   32
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
      Caption         =   "Datos de Operador."
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
      TabIndex        =   25
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
      Caption         =   "Para agregar o quitar permisos, active o desactive la casilla deseada."
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
      TabIndex        =   24
      Top             =   6720
      Width           =   6375
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00808080&
      Height          =   2655
      Left            =   120
      Top             =   6840
      Width           =   13695
   End
End
Attribute VB_Name = "mantenimiento_usuarios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

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
Call LimpiarTextBoxOperador

'quita los cheks seleccionados de la lista permisos
Call QuitarCheksEnListaPermisos

End Sub

Private Sub b_eliminar_Click()
msg = MsgBox("¿Seguro desea elimar el registro seleccionado?", vbQuestion + vbYesNo)
    If msg = vbYes Then
        Call Conn_BDaiosoft
        Conn_Mysqldb.Execute "DELETE FROM operador WHERE tipodoc = '" & c_tipodoc.Text & "' AND nrodoc= '" & t_nrodoc.Text & "'"
        MsgBox "El registro ha sido eliminado con éxito"
        Call LimpiarTextBoxOperador
        Call QuitarCheksEnListaPermisos
        b_eliminar.Enabled = False
        b_modificar.Enabled = False
        t_buscar.Text = ""
    End If
End Sub

Private Sub b_guardar_Click()
Dim PermisosSeleccionados As Boolean
'determina si  exite algun campo en blanco antes de guardar los datos
If t_nrodoc.Text = "" Or t_apellidos.Text = "" Or t_nombres.Text = "" Or t_direccion.Text = "" Or t_telefono.Text = "" Or t_login.Text = "" Or t_password.Text = "" Then
    MsgBox "No puede dejar ningún campo en blanco"
Else
    'determina si no se ha seleccionado ningun permiso para el operador.
    'Si no se a asigando ningun permiso, se le informa al  operador que se
    'le debe asignar  almenos un permiso a cada operador
    For x = 0 To (l_permisos.ListCount - 1)
        If l_permisos.Selected(x) = False Then
            PermisosSeleccionados = False
        Else
            PermisosSeleccionados = True
            Exit For
        End If
    Next x
    If Not PermisosSeleccionados Then
        MsgBox "Debe asignar almenos un permiso al operador"
    Else
        Call GuardarCambiosOperador
        
    End If
End If

End Sub

Private Sub b_modificar_Click()
'da un valor 2 a la variable BanderaOperador  para indicar que se desea
'editar  el registro actual del operador operador
BanderaOperador = 2

'habilita y desabilita los botones correspondientes
b_modificar.Enabled = False
b_nuevo.Enabled = False
b_eliminar.Enabled = False
b_cancelar.Enabled = True
b_guardar.Enabled = True
c_tipodoc.Locked = True
t_nrodoc.Locked = True
t_login.Locked = True
TempLoginOperador = t_login.Text

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
ValorCedulaOperador = t_nrodoc.Text
ValorLoginOperador = t_login.Text
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
    Call QuitarCheksEnListaPermisos
    Call LimpiarTextBoxOperador
    c_tipodoc.Locked = False
    t_nrodoc.Locked = False
    t_login.Locked = False
    
    'da un valor 1 a la variable BanderaOperador  para indicar que se desea
    'crear  un registro para nuevo operador
    BanderaOperador = 1

End Sub

Private Sub c_perfil_Click()

If c_perfil.Text = "Cajero" Then
    For x = 0 To 25
        If x = 5 Or x = 7 Then
            l_permisos.Selected(x) = True
        Else
            l_permisos.Selected(x) = False
        End If
    Next x
End If
If c_perfil.Text = "Ventas" Then
    For x = 0 To 25
        If x = 5 Or x = 6 Then
            l_permisos.Selected(x) = True
        Else
            l_permisos.Selected(x) = False
        End If
    Next x
End If
If c_perfil.Text = "Compras" Then
    For x = 0 To 25
        If x = 13 Or x = 14 Or x = 15 Then
            l_permisos.Selected(x) = True
        Else
            l_permisos.Selected(x) = False
        End If
    Next x
End If
If c_perfil.Text = "Cobranza" Then
    For x = 0 To 25
        If x = 5 Or x = 9 Or x = 10 Or x = 11 Or x = 12 Then
            l_permisos.Selected(x) = True
        Else
            l_permisos.Selected(x) = False
        End If
    Next x
End If
If c_perfil.Text = "Administrador" Then
    For x = 0 To 28
        l_permisos.Selected(x) = True
    Next x
End If
If c_perfil.Text = "Personalizado" Then
    For x = 0 To 28
        l_permisos.Selected(x) = False
    Next x
End If



End Sub


Private Sub fg_operador_Click()
If fg_operador.Rows > 1 Then
Dim TempOperador As ADODB.Recordset
Dim VarIdOperador As Integer

VarIdOperador = fg_operador.TextMatrix(fg_operador.Row, 0)
c_tipodoc.Text = fg_operador.TextMatrix(fg_operador.Row, 1)
t_nrodoc.Text = fg_operador.TextMatrix(fg_operador.Row, 2)
t_apellidos.Text = fg_operador.TextMatrix(fg_operador.Row, 3)
t_nombres.Text = fg_operador.TextMatrix(fg_operador.Row, 4)
t_direccion.Text = fg_operador.TextMatrix(fg_operador.Row, 5)
t_telefono.Text = fg_operador.TextMatrix(fg_operador.Row, 6)
c_perfil.Text = fg_operador.TextMatrix(fg_operador.Row, 7)
t_login.Text = fg_operador.TextMatrix(fg_operador.Row, 8)
t_password.Text = fg_operador.TextMatrix(fg_operador.Row, 9)

Call Conn_BDaiosoft
Set TempOperador = New ADODB.Recordset
TempOperador.Open "SELECT * from operador WHERE idoperador= " & VarIdOperador & "", Conn_Mysqldb, adOpenDynamic, adLockBatchOptimistic
If TempOperador.BOF = False And TempOperador.EOF = False Then
    TempOperador.MoveFirst
    IndiceCampos = 10
    For x = 0 To 28
        l_permisos.Selected(x) = TempOperador.Fields(IndiceCampos).Value
        IndiceCampos = IndiceCampos + 1
    Next x
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


'carga la lista permisos
l_permisos.AddItem "Mantenimiento"
l_permisos.AddItem "Mantenimiento: Clientes"
l_permisos.AddItem "Mantenimiento: Proveedores"
l_permisos.AddItem "Mantenimiento: Inventario"
l_permisos.AddItem "Mantenimiento: Almacén"
l_permisos.AddItem "Ventas"
l_permisos.AddItem "Ventas: Facturar"
l_permisos.AddItem "Ventas: Punto de venta"
l_permisos.AddItem "Ventas: Presupuestos"
l_permisos.AddItem "Ventas: Cuentas por cobrar"
l_permisos.AddItem "Ventas: Notas de crédito"
l_permisos.AddItem "Ventas: Comunicación de baja"
l_permisos.AddItem "Ventas: Anulaciones"
l_permisos.AddItem "Compras"
l_permisos.AddItem "Compras: Reg compras"
l_permisos.AddItem "Compras: Orden de compras"
l_permisos.AddItem "Compras: Cuentas por pagar"
l_permisos.AddItem "Reportes"
l_permisos.AddItem "Reportes: Ventas"
l_permisos.AddItem "Reportes: Compras"
l_permisos.AddItem "Reportes: Inventario"
l_permisos.AddItem "Reportes: Cuentas por pagar"
l_permisos.AddItem "Reportes: Cuentas por cobrar"
l_permisos.AddItem "Ajustes"
l_permisos.AddItem "Ajustes: Datos de la empresa"
l_permisos.AddItem "Ajustes: Usuarios del sistema"
l_permisos.AddItem "Soporte"
l_permisos.AddItem "Soporte: Servicios"
l_permisos.AddItem "Soporte: Garantías"


'establece el tamaño del formulario
mantenimiento_usuarios.Width = 14205
mantenimiento_usuarios.Height = 9945
mantenimiento_usuarios.Left = 2000

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

c_perfil.AddItem "Cajero"
c_perfil.AddItem "Ventas"
c_perfil.AddItem "Compras"
c_perfil.AddItem "Cobranza"
c_perfil.AddItem "Administrador"
c_perfil.AddItem "Personalizado"


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
Dim TempOperador As ADODB.Recordset
Call Conn_BDaiosoft
Set TempOperador = New ADODB.Recordset

If Not t_buscar.Text = "" Then
    fg_operador.Rows = 1
    If fg_operador.Rows = 1 Then
        If c_buscar.Text = "Nombre" Then
            TempOperador.Open "SELECT * FROM operador WHERE  nombres LIKE '%" & t_buscar.Text & "%'", Conn_Mysqldb, adOpenDynamic, adLockBatchOptimistic
        End If
        If c_buscar.Text = "Apellido" Then
            TempOperador.Open "SELECT * FROM operador WHERE  apellidos LIKE '%" & t_buscar.Text & "%'", Conn_Mysqldb, adOpenDynamic, adLockBatchOptimistic
        End If
        If c_buscar.Text = "Nro Documento" Then
             TempOperador.Open "SELECT * FROM operador WHERE  nrodoc LIKE '%" & t_buscar.Text & "%'", Conn_Mysqldb, adOpenDynamic, adLockBatchOptimistic
        End If
        If TempOperador.BOF = False And TempOperador.EOF = False Then
            TempOperador.MoveFirst
            fg_operador.Clear
            fg_operador.FormatString = "ID|Tipo Doc   |Nro Documento         | Apellidos                    |Nombres                              |Dirección          |Teléfono          |Tipo Operador           |Login  |Password "
            Do While TempOperador.EOF = False
                fg_operador.AddItem TempOperador.Fields(0)
                fg_operador.TextMatrix(fg_operador.Rows - 1, 1) = TempOperador.Fields(1).Value
                fg_operador.TextMatrix(fg_operador.Rows - 1, 2) = TempOperador.Fields(2).Value
                fg_operador.TextMatrix(fg_operador.Rows - 1, 3) = TempOperador.Fields(3).Value
                fg_operador.TextMatrix(fg_operador.Rows - 1, 4) = TempOperador.Fields(4).Value
                fg_operador.TextMatrix(fg_operador.Rows - 1, 5) = TempOperador.Fields(5).Value
                fg_operador.TextMatrix(fg_operador.Rows - 1, 6) = TempOperador.Fields(6).Value
                fg_operador.TextMatrix(fg_operador.Rows - 1, 7) = TempOperador.Fields(7).Value
                fg_operador.TextMatrix(fg_operador.Rows - 1, 8) = TempOperador.Fields(8).Value
                fg_operador.TextMatrix(fg_operador.Rows - 1, 9) = TempOperador.Fields(9).Value
 
                TempOperador.MoveNext
            Loop
        Else
            fg_operador.Clear
            fg_operador.FormatString = "ID|Tipo Doc   |Nro Documento         | Apellidos                    |Nombres                              |Dirección          |Teléfono          |Tipo Operador           |Login  |Password "
        End If
    End If
Else
    fg_operador.Clear
    fg_operador.Rows = 1
    fg_operador.FormatString = "ID|Tipo Doc   |Nro Documento         | Apellidos                    |Nombres                              |Dirección          |Teléfono          |Tipo Operador           |Login  |Password "
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


