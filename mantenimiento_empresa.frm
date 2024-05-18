VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form mantenimiento_empresa 
   Appearance      =   0  'Flat
   BackColor       =   &H00404040&
   Caption         =   "Información de la empresa."
   ClientHeight    =   5790
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10680
   Icon            =   "mantenimiento_empresa.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5790
   ScaleWidth      =   10680
   Begin VB.PictureBox Picture5 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   870
      Left            =   7800
      ScaleHeight     =   840
      ScaleWidth      =   1755
      TabIndex        =   23
      Top             =   120
      Width           =   1785
      Begin VB.CommandButton b_cancelar 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         DisabledPicture =   "mantenimiento_empresa.frx":058A
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   900
         Left            =   -30
         MaskColor       =   &H000000FF&
         Picture         =   "mantenimiento_empresa.frx":8CE4
         Style           =   1  'Graphical
         TabIndex        =   28
         ToolTipText     =   "Cancelar"
         Top             =   -30
         UseMaskColor    =   -1  'True
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
      TabIndex        =   22
      Top             =   120
      Width           =   1785
      Begin VB.CommandButton b_guardar 
         BackColor       =   &H00404040&
         DisabledPicture =   "mantenimiento_empresa.frx":1143E
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
         Picture         =   "mantenimiento_empresa.frx":19B98
         Style           =   1  'Graphical
         TabIndex        =   27
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
      Left            =   3960
      ScaleHeight     =   840
      ScaleWidth      =   1755
      TabIndex        =   21
      Top             =   120
      Width           =   1785
      Begin VB.CommandButton b_eliminar 
         BackColor       =   &H00404040&
         DisabledPicture =   "mantenimiento_empresa.frx":222F2
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
         Picture         =   "mantenimiento_empresa.frx":2A534
         Style           =   1  'Graphical
         TabIndex        =   26
         ToolTipText     =   "Eliminar producto"
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
      TabIndex        =   20
      Top             =   120
      Width           =   1785
      Begin VB.CommandButton b_modificar 
         BackColor       =   &H00404040&
         DisabledPicture =   "mantenimiento_empresa.frx":32776
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
         Picture         =   "mantenimiento_empresa.frx":3ACC4
         Style           =   1  'Graphical
         TabIndex        =   25
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
      ScaleWidth      =   1755
      TabIndex        =   19
      Top             =   120
      Width           =   1785
      Begin VB.CommandButton b_agregar 
         BackColor       =   &H00404040&
         DisabledPicture =   "mantenimiento_empresa.frx":43212
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
         Picture         =   "mantenimiento_empresa.frx":4BA74
         Style           =   1  'Graphical
         TabIndex        =   24
         ToolTipText     =   "Agregar Nuevo producto"
         Top             =   -30
         Width           =   1815
      End
   End
   Begin VB.Frame fr_datosempresa 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Caption         =   "Datos Empresa."
      Enabled         =   0   'False
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
      Height          =   4095
      Left            =   240
      TabIndex        =   0
      Top             =   1440
      Width           =   10215
      Begin VB.TextBox t_vencilicencia 
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
         Left            =   7560
         Locked          =   -1  'True
         TabIndex        =   31
         Top             =   2280
         Width           =   2235
      End
      Begin VB.TextBox t_fechainiperiodo 
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
         Left            =   7560
         Locked          =   -1  'True
         TabIndex        =   30
         Top             =   1680
         Width           =   2235
      End
      Begin MSComDlg.CommonDialog CD_logoempresa 
         Left            =   6840
         Top             =   3360
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.PictureBox p_logo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   975
         Left            =   7560
         Picture         =   "mantenimiento_empresa.frx":542D6
         ScaleHeight     =   945
         ScaleWidth      =   2385
         TabIndex        =   18
         Top             =   3000
         Width           =   2415
      End
      Begin VB.TextBox t_sitioweb 
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
         Left            =   7560
         TabIndex        =   13
         Top             =   1080
         Width           =   2235
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
         Left            =   7560
         TabIndex        =   12
         Top             =   480
         Width           =   2235
      End
      Begin VB.TextBox t_telefonocelular 
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
         TabIndex        =   11
         Top             =   3480
         Width           =   2235
      End
      Begin VB.TextBox t_telefonofijo 
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
         TabIndex        =   5
         Top             =   2880
         Width           =   2235
      End
      Begin VB.TextBox t_direccion2 
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
         TabIndex        =   4
         Top             =   2280
         Width           =   3915
      End
      Begin VB.TextBox t_direccion1 
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
         TabIndex        =   3
         Top             =   1680
         Width           =   3915
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
         TabIndex        =   2
         Top             =   1080
         Width           =   3915
      End
      Begin VB.TextBox t_ruc 
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
         TabIndex        =   1
         Top             =   480
         Width           =   2235
      End
      Begin VB.Label Label12 
         BackColor       =   &H00404040&
         Caption         =   "Vencimiento Lic:"
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
         TabIndex        =   33
         Top             =   2280
         Width           =   1575
      End
      Begin VB.Label Label11 
         BackColor       =   &H00404040&
         Caption         =   "Ini. Periodo Fisc:"
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
         TabIndex        =   32
         Top             =   1680
         Width           =   1575
      End
      Begin VB.Label Label9 
         BackColor       =   &H00404040&
         Caption         =   "Logo:"
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
         Left            =   6960
         TabIndex        =   17
         Top             =   3000
         Width           =   1095
      End
      Begin VB.Label Label8 
         BackColor       =   &H00404040&
         Caption         =   "Sitio Web:"
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
         Left            =   6480
         TabIndex        =   16
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label Label7 
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
         Left            =   6840
         TabIndex        =   15
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label Label6 
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
         Height          =   255
         Left            =   0
         TabIndex        =   14
         Top             =   3480
         Width           =   1935
      End
      Begin VB.Label Label5 
         BackColor       =   &H00404040&
         Caption         =   "Teléfono fijo:"
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
         TabIndex        =   10
         Top             =   2880
         Width           =   1695
      End
      Begin VB.Label Label4 
         BackColor       =   &H00404040&
         Caption         =   "Otra dirección:"
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
         TabIndex        =   9
         Top             =   2280
         Width           =   1575
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
         Left            =   600
         TabIndex        =   8
         Top             =   1680
         Width           =   1335
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
         Left            =   120
         TabIndex        =   7
         Top             =   1080
         Width           =   1695
      End
      Begin VB.Label Label1 
         BackColor       =   &H00404040&
         Caption         =   "RUT:"
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
         Left            =   1080
         TabIndex        =   6
         Top             =   480
         Width           =   1575
      End
   End
   Begin VB.Label Label10 
      BackColor       =   &H00404040&
      Caption         =   "Datos de empresa."
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
      TabIndex        =   29
      Top             =   1200
      Width           =   1815
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00808080&
      Height          =   4335
      Left            =   120
      Top             =   1320
      Width           =   10455
   End
End
Attribute VB_Name = "mantenimiento_empresa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'variable bandera para determinar si se debe agregar un registro nuevo
'o se debe actualizar uno existente a la tabla datos_empresa
Public bandera As String
Public NuevaRuta As String

Private Sub b_agregar_Click()
fr_datosempresa.Enabled = True
b_agregar.Enabled = False
b_modificar.Enabled = False
b_cancelar.Enabled = True
b_guardar.Enabled = True
bandera = "agregar"
End Sub

Private Sub b_cancelar_Click()

t_ruc.Text = ""
t_nombrers.Text = ""
t_direccion1.Text = ""
t_direccion2.Text = ""
t_telefonofijo.Text = ""
t_telefonocelular.Text = ""
t_sitioweb.Text = ""
t_email.Text = ""
p_logo.Picture = LoadPicture
b_guardar.Enabled = False
b_cancelar.Enabled = False

Call Conn_BDaiosoft
Dim TempDatosEmpresa As ADODB.Recordset
Set TempDatosEmpresa = New ADODB.Recordset
TempDatosEmpresa.Open "SELECT * FROM datos_empresa", Conn_Mysqldb, adOpenDynamic, adLockBatchOptimistic
If TempDatosEmpresa.EOF = False And TempDatosEmpresa.BOF = False Then
    TempDatosEmpresa.MoveFirst
    t_ruc.Text = TempDatosEmpresa.Fields(0).Value
    t_nombrers.Text = TempDatosEmpresa.Fields(1).Value
    t_direccion1.Text = TempDatosEmpresa.Fields(2).Value
    t_direccion2.Text = TempDatosEmpresa.Fields(3).Value
    t_telefonofijo.Text = TempDatosEmpresa.Fields(4).Value
    t_telefonocelular.Text = TempDatosEmpresa.Fields(5).Value
    t_sitioweb.Text = TempDatosEmpresa.Fields(6).Value
    t_email.Text = TempDatosEmpresa.Fields(7).Value
    RutaLogo = TempDatosEmpresa.Fields(8).Value
    If RutaLogo <> "" Then
        p_logo.Picture = LoadPicture(RutaLogo, vbLPLarge, vbLPColor)
    Else
        p_logo.Picture = LoadPicture
    End If
    b_modificar.Enabled = True
    b_eliminar.Enabled = True
Else
    b_agregar.Enabled = True
    b_modificar.Enabled = False
    b_eliminar.Enabled = False
End If
fr_datosempresa.Enabled = False

End Sub

Private Sub b_eliminar_Click()

msg = MsgBox("¿Seguro desea eliminar?", vbQuestion + vbYesNo)
If msg = vbYes Then
    Call Conn_BDaiosoft
    Conn_Mysqldb.Execute "DELETE FROM datos_empresa"
    
    MsgBox "El registro ha sido eliminado"
    t_ruc.Text = ""
    t_nombrers.Text = ""
    t_direccion1.Text = ""
    t_direccion2.Text = ""
    t_telefonofijo.Text = ""
    t_telefonocelular.Text = ""
    t_sitioweb.Text = ""
    t_email.Text = ""
    p_logo.Picture = LoadPicture
    
    b_eliminar.Enabled = False
    b_modificar.Enabled = False
    b_agregar.Enabled = True
    
End If

End Sub

Private Sub b_guardar_Click()

If t_ruc.Text = "" Or t_nombrers.Text = "" Then
    MsgBox "Debe Ingresar RUC y nombre de empresa."
Else
    msg = MsgBox("¿Está de acuerdo con los cambios?", vbQuestion + vbYesNo)
    If msg = vbYes Then
    
        If bandera = "agregar" Then

            Call Conn_BDaiosoft
            
            Conn_Mysqldb.Execute "INSERT INTO datos_empresa SET ruc = '" & t_ruc.Text & "'," _
            & "nombre = '" & t_nombrers.Text & "'," _
            & "direccionppal = '" & t_direccion1.Text & "'," _
            & "direccionsuc = '" & t_direccion2.Text & "'," _
            & "telefonofijo = '" & t_telefonofijo.Text & "'," _
            & "telefonocelular = '" & t_telefonocelular.Text & "'," _
            & "paginaweb = '" & t_sitioweb.Text & "'," _
            & "email = '" & t_email.Text & "'," _
            & "logo = '" & NuevaRuta & "'," _
            & "fechainiperfiscal = '" & Format(t_fechainiperiodo.Text, "yyyy-mm-dd") & "'," _
            & "fechavencilic = '" & Format(t_vencilicencia.Text, "yyyy-mm-dd") & "'," _
           
            
            MsgBox "La inforamción se guardó exitosamente"
            fr_datosempresa.Enabled = False
            b_guardar.Enabled = False
            b_cancelar.Enabled = False
            b_eliminar.Enabled = True
            b_modificar.Enabled = True
            
        End If
        If bandera = "modificar" Then

            Call Conn_BDaiosoft
            
            Conn_Mysqldb.Execute "UPDATE datos_empresa SET ruc = '" & t_ruc.Text & "'," _
            & "nombre = '" & t_nombrers.Text & "'," _
            & "direccionppal = '" & t_direccion1.Text & "'," _
            & "direccionsuc = '" & t_direccion2.Text & "'," _
            & "telefonofijo = '" & t_telefonofijo.Text & "'," _
            & "telefonocelular = '" & t_telefonocelular.Text & "'," _
            & "paginaweb = '" & t_sitioweb.Text & "'," _
            & "email = '" & t_email.Text & "'," _
            & "fechainiperfiscal = '" & Format(t_fechainiperiodo.Text, "yyyy-mm-dd") & "'," _
            & "fechavencilic = '" & Format(t_vencilicencia.Text, "yyyy-mm-dd") & "'," _
            & "logo = '" & NuevaRuta & "'"
            
            MsgBox "La inforamción se actualizó exitosamente"
            fr_datosempresa.Enabled = False
            b_guardar.Enabled = False
            b_cancelar.Enabled = False
            b_eliminar.Enabled = True
            b_modificar.Enabled = True
            
        End If
        
        
    End If
End If



End Sub

Private Sub b_modificar_Click()
bandera = "modificar"
b_agregar.Enabled = False
b_modificar.Enabled = False
b_eliminar.Enabled = False
b_cancelar.Enabled = True
b_guardar.Enabled = True
fr_datosempresa.Enabled = True
End Sub

Private Sub b_nuevo_Click()

End Sub

Private Sub Form_Load()

'codigo para dar las dimenciones de ancho y alto al formulario
mantenimiento_empresa.Width = 10920
mantenimiento_empresa.Height = 6360

'codigo para posicionar el formulario en el centro de la pantala
mantenimiento_empresa.Left = f_principal.ScaleWidth / 2 - mantenimiento_empresa.ScaleWidth / 2
mantenimiento_empresa.Top = f_principal.ScaleHeight / 2 - mantenimiento_empresa.ScaleHeight / 2

Call Conn_BDaiosoft
Dim TempDatosEmpresa As ADODB.Recordset
Set TempDatosEmpresa = New ADODB.Recordset
TempDatosEmpresa.Open "SELECT * FROM datos_empresa", Conn_Mysqldb, adOpenDynamic, adLockBatchOptimistic
If TempDatosEmpresa.EOF = False And TempDatosEmpresa.BOF = False Then
    TempDatosEmpresa.MoveFirst
    t_ruc.Text = TempDatosEmpresa.Fields(0).Value
    t_nombrers.Text = TempDatosEmpresa.Fields(1).Value
    t_direccion1.Text = TempDatosEmpresa.Fields(2).Value
    t_direccion2.Text = TempDatosEmpresa.Fields(3).Value
    t_telefonofijo.Text = TempDatosEmpresa.Fields(4).Value
    t_telefonocelular.Text = TempDatosEmpresa.Fields(5).Value
    t_sitioweb.Text = TempDatosEmpresa.Fields(6).Value
    t_email.Text = TempDatosEmpresa.Fields(7).Value
    t_fechainiperiodo.Text = TempDatosEmpresa.Fields(9).Value
    t_vencilicencia.Text = TempDatosEmpresa.Fields(10).Value
    RutaLogo = TempDatosEmpresa.Fields(8).Value
    If RutaLogo <> "" Then
        On Error GoTo linea32
        p_logo.Picture = LoadPicture(RutaLogo, vbLPLarge, vbLPColor)
    Else
linea32:
        p_logo.Picture = LoadPicture
    End If
    b_modificar.Enabled = True
    b_eliminar.Enabled = True

Else
    b_agregar.Enabled = True
End If


End Sub

Private Sub p_logo_DblClick()
CD_logoempresa.ShowOpen
RutaImagen = CD_logoempresa.FileName
p_logo.Picture = LoadPicture(RutaImagen, vbLPLarge, vbLPColor)

'codigo para colocar doble slash invertido a la ruta de la imagen
'para que mysql la guarde correctamente ya que si se coloca un
'solo slash invetido mysql lo omite
tamaño = Len(RutaImagen)
NuevaRuta = ""
For x = 1 To tamaño
    If Mid(RutaImagen, x, 1) = "\" Then
        NuevaRuta = NuevaRuta & Mid(RutaImagen, x, 1) & "\"
    Else
        NuevaRuta = NuevaRuta & Mid(RutaImagen, x, 1)
    End If
Next x
'fin de codigo de formato de ruta

End Sub

Private Sub t_direccion1_KeyPress(KeyAscii As Integer)
If KeyAscii = 39 Then
    KeyAscii = 0
End If
End Sub

Private Sub t_direccion2_KeyPress(KeyAscii As Integer)
If KeyAscii = 39 Then
    KeyAscii = 0
End If
End Sub

Private Sub t_email_KeyPress(KeyAscii As Integer)
If KeyAscii = 39 Then
    KeyAscii = 0
End If
End Sub

Private Sub t_fechainiperiodo_DblClick()
If t_fechainiperiodo.Locked = True Then
    varclave = InputBox("Ingrese credencial para modificar este campo", "Validacion")
    If varclave = "Aio2019Pass" Then
        t_fechainiperiodo.Locked = False
    End If
End If
End Sub

Private Sub t_fechainiperiodo_KeyPress(KeyAscii As Integer)
If KeyAscii = 39 Then
    KeyAscii = 0
End If
End Sub

Private Sub t_nombrers_KeyPress(KeyAscii As Integer)
If KeyAscii = 39 Then
    KeyAscii = 0
End If
End Sub

Private Sub t_ruc_KeyPress(KeyAscii As Integer)
If KeyAscii < 48 Or KeyAscii > 57 Then
    If KeyAscii <> 8 And KeyAscii <> 13 Then
        KeyAscii = 0
        Beep
    End If
End If
'If KeyAscii = 13 Then
'    If ValidarRuc(t_ruc.Text) = True Then
'        t_nombrers.SetFocus
'    Else
'        t_ruc.SetFocus
'    End If
'End If
End Sub
Private Sub t_ruc_LostFocus()
'If ValidarRuc(t_ruc.Text) = True Then
'    t_nombrers.SetFocus
'Else
'    t_ruc.SetFocus
'End If
End Sub

Private Sub t_sitioweb_KeyPress(KeyAscii As Integer)
If KeyAscii = 39 Then
    KeyAscii = 0
End If
End Sub

Private Sub t_telefonocelular_KeyPress(KeyAscii As Integer)
If KeyAscii = 39 Then
    KeyAscii = 0
End If
End Sub

Private Sub t_telefonofijo_KeyPress(KeyAscii As Integer)
If KeyAscii = 39 Then
    KeyAscii = 0
End If
End Sub

Private Sub t_vencilicencia_DblClick()
If t_vencilicencia.Locked = True Then
    varclave = InputBox("Ingrese credencial para modificar este campo", "Validacion")
    If varclave = "Aio2019Pass" Then
        t_vencilicencia.Locked = False
    End If
End If
End Sub

Private Sub t_vencilicencia_KeyPress(KeyAscii As Integer)
If KeyAscii = 39 Then
    KeyAscii = 0
End If
End Sub
