VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form registro_orden 
   BackColor       =   &H00FFC0C0&
   Caption         =   "Registro de Orden"
   ClientHeight    =   10950
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   18960
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10950
   ScaleWidth      =   18960
   WindowState     =   2  'Maximized
   Begin VB.CommandButton b_anular 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Anular"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   11640
      Picture         =   "registro_orden.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   108
      Top             =   9720
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.CommandButton b_volver 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Volver"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   14520
      Picture         =   "registro_orden.frx":0DB0
      Style           =   1  'Graphical
      TabIndex        =   107
      Top             =   9960
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.CommandButton b_reimprimir 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Reimprimir"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   11040
      Picture         =   "registro_orden.frx":1BE6
      Style           =   1  'Graphical
      TabIndex        =   106
      Top             =   9240
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.TextBox t_restante 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   101
      Top             =   10440
      Width           =   2055
   End
   Begin VB.TextBox t_abono 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   100
      Top             =   9720
      Width           =   2055
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   7680
      Top             =   10080
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   7080
      Top             =   10080
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Registro de Orden"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   240
      TabIndex        =   61
      Top             =   7080
      Width           =   18495
      Begin VB.Frame Frame4 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Forma de pago"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1815
         Left            =   0
         TabIndex        =   109
         Top             =   240
         Width           =   2535
         Begin VB.OptionButton o_abono 
            BackColor       =   &H00FFC0C0&
            Caption         =   "Abono"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   240
            TabIndex        =   113
            Top             =   1440
            Width           =   1575
         End
         Begin VB.OptionButton o_alrecojo 
            BackColor       =   &H00FFC0C0&
            Caption         =   "Pago al Recoger"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   112
            Top             =   1080
            Width           =   2175
         End
         Begin VB.OptionButton o_prepagado 
            BackColor       =   &H00FFC0C0&
            Caption         =   "Prepagado"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   240
            TabIndex        =   111
            Top             =   360
            Value           =   -1  'True
            Width           =   1695
         End
         Begin VB.OptionButton o_pagoexacto 
            BackColor       =   &H00FFC0C0&
            Caption         =   "Pago Exacto"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   110
            Top             =   720
            Width           =   1695
         End
      End
      Begin VB.TextBox t_mtodesc 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   6600
         TabIndex        =   82
         Top             =   1560
         Width           =   2055
      End
      Begin VB.TextBox t_xcentajedesc 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   6600
         TabIndex        =   80
         Top             =   600
         Width           =   2055
      End
      Begin VB.CommandButton b_buscar 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Buscar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1695
         Left            =   16080
         Picture         =   "registro_orden.frx":79E8
         Style           =   1  'Graphical
         TabIndex        =   68
         Top             =   360
         Width           =   2175
      End
      Begin VB.CommandButton b_factura 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Factura"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1695
         Left            =   13680
         Picture         =   "registro_orden.frx":92B0
         Style           =   1  'Graphical
         TabIndex        =   67
         Top             =   360
         Width           =   2175
      End
      Begin VB.CommandButton b_boleta 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Boleta"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1695
         Left            =   11280
         Picture         =   "registro_orden.frx":A6FC
         Style           =   1  'Graphical
         TabIndex        =   66
         Top             =   360
         Width           =   2175
      End
      Begin VB.TextBox t_vuelto 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4320
         Locked          =   -1  'True
         TabIndex        =   64
         Top             =   1560
         Width           =   2055
      End
      Begin VB.TextBox t_efectivo 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4320
         TabIndex        =   63
         Top             =   960
         Width           =   2055
      End
      Begin VB.TextBox t_total 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4320
         Locked          =   -1  'True
         TabIndex        =   62
         Top             =   360
         Width           =   2055
      End
      Begin VB.CommandButton b_orden 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Procesar orden"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1695
         Left            =   8880
         Picture         =   "registro_orden.frx":BB54
         Style           =   1  'Graphical
         TabIndex        =   65
         Top             =   360
         Width           =   2175
      End
      Begin VB.Label l_restante 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Restante:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   103
         Top             =   1560
         Width           =   1695
      End
      Begin VB.Label l_abono 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Abono:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   102
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Label l_mtodesc 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Mto Dscto."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6720
         TabIndex        =   81
         Top             =   1200
         Width           =   2055
      End
      Begin VB.Label l_xcentajedesc 
         BackColor       =   &H00FFC0C0&
         Caption         =   "% Dscto."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   6720
         TabIndex        =   79
         Top             =   240
         Width           =   1935
      End
      Begin VB.Label l_vuelto 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Vuelto:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3000
         TabIndex        =   71
         Top             =   1560
         Width           =   1215
      End
      Begin VB.Label l_efectivo 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Efectivo:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2760
         TabIndex        =   70
         Top             =   960
         Width           =   1575
      End
      Begin VB.Label l_total 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Total:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3240
         TabIndex        =   69
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.Frame fr_prendasvestir 
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4815
      Left            =   9960
      TabIndex        =   44
      Top             =   -1200
      Visible         =   0   'False
      Width           =   6855
      Begin VB.CommandButton b_atras3 
         Height          =   615
         Left            =   5400
         Picture         =   "registro_orden.frx":CC7D
         Style           =   1  'Graphical
         TabIndex        =   96
         Top             =   2880
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.CommandButton b_boton36 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Polera"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   5160
         Style           =   1  'Graphical
         TabIndex        =   94
         Top             =   3120
         Width           =   1575
      End
      Begin VB.CommandButton b_boton35 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Gorra"
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
         Left            =   3480
         Style           =   1  'Graphical
         TabIndex        =   92
         Top             =   4080
         Width           =   1575
      End
      Begin VB.CommandButton b_boton34 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Jumper Mujer"
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
         Left            =   1800
         Style           =   1  'Graphical
         TabIndex        =   91
         Top             =   4080
         Width           =   1575
      End
      Begin VB.CommandButton b_boton33 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Jumper Niña"
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
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   90
         Top             =   4080
         Width           =   1575
      End
      Begin VB.CommandButton b_boton31 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Terno (saco pantalon chaleco)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   3480
         Style           =   1  'Graphical
         TabIndex        =   89
         Top             =   3120
         Width           =   1575
      End
      Begin VB.CommandButton b_boton30 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Terno (saco pantalon)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   1800
         Style           =   1  'Graphical
         TabIndex        =   88
         Top             =   3120
         Width           =   1575
      End
      Begin VB.CommandButton b_boton29 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Saco Masculino"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   87
         Top             =   3120
         Width           =   1575
      End
      Begin VB.CommandButton b_atras 
         Height          =   615
         Left            =   5280
         Picture         =   "registro_orden.frx":1358A
         Style           =   1  'Graphical
         TabIndex        =   58
         Top             =   4080
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.CommandButton b_sig 
         BackColor       =   &H00FFFFFF&
         Height          =   615
         Left            =   6120
         Picture         =   "registro_orden.frx":19E97
         Style           =   1  'Graphical
         TabIndex        =   57
         Top             =   4080
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.CommandButton b_boton28 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Saco Dama Grande"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   5160
         Style           =   1  'Graphical
         TabIndex        =   56
         Top             =   2160
         Width           =   1575
      End
      Begin VB.CommandButton b_boton25 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Vestido Niña"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   55
         Top             =   2160
         Width           =   1575
      End
      Begin VB.CommandButton b_boton27 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Saco Dama Mediano"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   3480
         Style           =   1  'Graphical
         TabIndex        =   54
         Top             =   2160
         Width           =   1575
      End
      Begin VB.CommandButton b_boton26 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Vestido Dama"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   1800
         Style           =   1  'Graphical
         TabIndex        =   53
         Top             =   2160
         Width           =   1575
      End
      Begin VB.CommandButton b_boton22 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Polo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   1800
         Style           =   1  'Graphical
         TabIndex        =   52
         Top             =   1200
         Width           =   1575
      End
      Begin VB.CommandButton b_boton19 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Casaca"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   3480
         Style           =   1  'Graphical
         TabIndex        =   51
         Top             =   240
         Width           =   1575
      End
      Begin VB.CommandButton b_boton24 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Mochila"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   5160
         Style           =   1  'Graphical
         TabIndex        =   50
         Top             =   1200
         Width           =   1575
      End
      Begin VB.CommandButton b_boton23 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Chompa"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   3480
         OLEDropMode     =   1  'Manual
         Style           =   1  'Graphical
         TabIndex        =   49
         Top             =   1200
         Width           =   1575
      End
      Begin VB.CommandButton b_boton21 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Bermuda"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   48
         Top             =   1200
         Width           =   1575
      End
      Begin VB.CommandButton b_boton18 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Pantalon"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   1800
         Style           =   1  'Graphical
         TabIndex        =   47
         Top             =   240
         Width           =   1575
      End
      Begin VB.CommandButton b_boton20 
         BackColor       =   &H00FFFFFF&
         Caption         =   "short"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   5160
         Style           =   1  'Graphical
         TabIndex        =   46
         Top             =   240
         Width           =   1575
      End
      Begin VB.CommandButton b_boton17 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Camisa"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   45
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.Frame fr_prendas 
      BackColor       =   &H00FFC0C0&
      Height          =   4815
      Left            =   10440
      TabIndex        =   27
      Top             =   -480
      Visible         =   0   'False
      Width           =   6855
      Begin VB.CommandButton b_sig2 
         BackColor       =   &H00FFFFFF&
         Height          =   615
         Left            =   6000
         Picture         =   "registro_orden.frx":206BA
         Style           =   1  'Graphical
         TabIndex        =   95
         Top             =   4080
         Width           =   615
      End
      Begin VB.CommandButton b_boton32 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Toalla"
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
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   93
         Top             =   4080
         Width           =   1575
      End
      Begin VB.CommandButton b_boton16 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Peluche Gigante"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   5160
         Style           =   1  'Graphical
         TabIndex        =   86
         Top             =   3120
         Width           =   1575
      End
      Begin VB.CommandButton b_boton15 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Peluche Grande"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   3480
         Style           =   1  'Graphical
         TabIndex        =   85
         Top             =   3120
         Width           =   1575
      End
      Begin VB.CommandButton b_boton14 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Peluche Mediano"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   1800
         Style           =   1  'Graphical
         TabIndex        =   84
         Top             =   3120
         Width           =   1575
      End
      Begin VB.CommandButton b_boton13 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Zapatillas Blancas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   5160
         Style           =   1  'Graphical
         TabIndex        =   83
         Top             =   2160
         Width           =   1575
      End
      Begin VB.CommandButton b_atras2 
         Height          =   615
         Left            =   5280
         Picture         =   "registro_orden.frx":26EDD
         Style           =   1  'Graphical
         TabIndex        =   59
         Top             =   4080
         Width           =   615
      End
      Begin VB.CommandButton b_boton12 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Juego de Sabanas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   1800
         Style           =   1  'Graphical
         TabIndex        =   39
         Top             =   2160
         Width           =   1575
      End
      Begin VB.CommandButton b_boton11 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Peluche Pequeño"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   38
         Top             =   3120
         Width           =   1575
      End
      Begin VB.CommandButton b_boton10 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Zapatillas Color"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   3480
         Style           =   1  'Graphical
         TabIndex        =   37
         Top             =   2160
         Width           =   1575
      End
      Begin VB.CommandButton b_boton9 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Almohada"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   36
         Top             =   2160
         Width           =   1575
      End
      Begin VB.CommandButton b_boton8 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Alfombra"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   5160
         Style           =   1  'Graphical
         TabIndex        =   35
         Top             =   1200
         Width           =   1575
      End
      Begin VB.CommandButton b_boton7 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Cortina por Pliege"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   3480
         Style           =   1  'Graphical
         TabIndex        =   34
         Top             =   1200
         Width           =   1575
      End
      Begin VB.CommandButton b_boton6 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Juego Muebles"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   1800
         Style           =   1  'Graphical
         TabIndex        =   33
         Top             =   1200
         Width           =   1575
      End
      Begin VB.CommandButton b_boton5 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Edredon Polar 2 PLazas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   32
         Top             =   1200
         Width           =   1575
      End
      Begin VB.CommandButton b_boton4 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Edredon Polar plaza y media"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   5160
         Style           =   1  'Graphical
         TabIndex        =   31
         Top             =   240
         Width           =   1575
      End
      Begin VB.CommandButton b_boton3 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Frazada 2 Plazas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   3480
         Style           =   1  'Graphical
         TabIndex        =   30
         Top             =   240
         Width           =   1575
      End
      Begin VB.CommandButton b_boton2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Frazada Plaza y Media"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   1800
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   240
         Width           =   1575
      End
      Begin VB.CommandButton b_boton1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Prendas por KG"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Cliente"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   240
      TabIndex        =   2
      Top             =   0
      Width           =   18495
      Begin VB.ComboBox c_tipodoc 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   4680
         Style           =   2  'Dropdown List
         TabIndex        =   99
         Top             =   1080
         Width           =   2775
      End
      Begin VB.TextBox t_direccion 
         Appearance      =   0  'Flat
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
         Left            =   14520
         TabIndex        =   8
         Top             =   1080
         Width           =   3135
      End
      Begin VB.TextBox t_email 
         Appearance      =   0  'Flat
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
         Left            =   14520
         TabIndex        =   6
         Top             =   360
         Width           =   3135
      End
      Begin VB.TextBox t_telefono 
         Appearance      =   0  'Flat
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
         Left            =   9600
         TabIndex        =   5
         Top             =   360
         Width           =   2415
      End
      Begin VB.TextBox t_nrodoc 
         Appearance      =   0  'Flat
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
         Left            =   9600
         TabIndex        =   7
         Top             =   1080
         Width           =   2415
      End
      Begin VB.TextBox t_nombrers 
         Appearance      =   0  'Flat
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
         Left            =   4680
         TabIndex        =   4
         Top             =   360
         Width           =   2775
      End
      Begin VB.CommandButton b_cliente 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Cliente"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1575
         Left            =   120
         Picture         =   "registro_orden.frx":2D7EA
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   240
         Width           =   1935
      End
      Begin VB.Label Label11 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Dirección:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   12240
         TabIndex        =   43
         Top             =   1080
         Width           =   2055
      End
      Begin VB.Label Label10 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Correo electrónico:"
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
         Left            =   12240
         TabIndex        =   42
         Top             =   360
         Width           =   2295
      End
      Begin VB.Label Label5 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Teléfono:"
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
         Left            =   7920
         TabIndex        =   18
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Nro documento:"
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
         Left            =   7920
         TabIndex        =   17
         Top             =   1080
         Width           =   1695
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Tipo documento:"
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
         Left            =   2160
         TabIndex        =   16
         Top             =   1080
         Width           =   1935
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Nombre o Razon Social:"
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
         Left            =   2160
         TabIndex        =   15
         Top             =   360
         Width           =   2535
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Registro de Servicio"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5175
      Left            =   240
      TabIndex        =   1
      Top             =   1920
      Width           =   18495
      Begin VB.TextBox t_editar 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   11760
         TabIndex        =   97
         Top             =   1440
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.TextBox t_importe1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   13440
         Locked          =   -1  'True
         TabIndex        =   74
         Top             =   4440
         Width           =   1575
      End
      Begin VB.TextBox t_igv1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   15000
         Locked          =   -1  'True
         TabIndex        =   73
         Top             =   4440
         Width           =   1575
      End
      Begin VB.TextBox t_total1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   16560
         Locked          =   -1  'True
         TabIndex        =   72
         Top             =   4440
         Width           =   1575
      End
      Begin VB.CommandButton b_otros 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Otros"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1815
         Left            =   5040
         Picture         =   "registro_orden.frx":34D5C
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   2760
         Width           =   1935
      End
      Begin VB.TextBox t_precio 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   7200
         Locked          =   -1  'True
         TabIndex        =   41
         Top             =   3120
         Width           =   2775
      End
      Begin VB.CommandButton Command14 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Lavado Centrifugado y Secado"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1815
         Left            =   480
         Picture         =   "registro_orden.frx":3C2CE
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   480
         Width           =   1935
      End
      Begin VB.CommandButton Command13 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Secado"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1815
         Left            =   2760
         Picture         =   "registro_orden.frx":3EC94
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   2760
         Width           =   1935
      End
      Begin VB.CommandButton Command12 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Teñidos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1815
         Left            =   480
         Picture         =   "registro_orden.frx":44136
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   2760
         Width           =   1935
      End
      Begin VB.CommandButton Command11 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Lavado al Seco"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1815
         Left            =   5040
         Picture         =   "registro_orden.frx":4ABB8
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   480
         Width           =   1935
      End
      Begin VB.CommandButton Command10 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Lavado y Centrifugado"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1815
         Left            =   2760
         Picture         =   "registro_orden.frx":4BB88
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   480
         Width           =   1935
      End
      Begin VB.CommandButton b_cancelar 
         BackColor       =   &H000000FF&
         Caption         =   "Limpiar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   8760
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   3840
         Width           =   1215
      End
      Begin VB.CommandButton b_agregar 
         BackColor       =   &H00008000&
         Caption         =   "Agregar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   7200
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   3840
         Width           =   1215
      End
      Begin VB.TextBox t_cantidad 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   7200
         TabIndex        =   23
         Top             =   1920
         Width           =   2775
      End
      Begin VB.TextBox t_peso 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   7200
         TabIndex        =   21
         Top             =   720
         Visible         =   0   'False
         Width           =   2775
      End
      Begin MSFlexGridLib.MSFlexGrid fg_prod_orden 
         Height          =   3735
         Left            =   10200
         TabIndex        =   19
         Top             =   360
         Width           =   8175
         _ExtentX        =   14420
         _ExtentY        =   6588
         _Version        =   393216
         Rows            =   1
         Cols            =   6
         FixedCols       =   0
         Appearance      =   0
         FormatString    =   "Código| Descripción                     |Precio U.|Cant.| Peso|    Precio  "
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label17 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFC0C0&
         Caption         =   "Importe"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   13440
         TabIndex        =   77
         Top             =   4080
         Width           =   1575
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFC0C0&
         Caption         =   "IGV 18%"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   15000
         TabIndex        =   76
         Top             =   4080
         Width           =   1575
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFC0C0&
         Caption         =   "Total"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   16560
         TabIndex        =   75
         Top             =   4080
         Width           =   1575
      End
      Begin VB.Label Label9 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Precio x Pieza:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   7200
         TabIndex        =   40
         Top             =   2640
         Width           =   3135
      End
      Begin VB.Label Label7 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Cant. Piezas:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   7200
         TabIndex        =   22
         Top             =   1440
         Width           =   2415
      End
      Begin VB.Label Label6 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Peso KG:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   7200
         TabIndex        =   20
         Top             =   240
         Visible         =   0   'False
         Width           =   2535
      End
   End
   Begin VB.Label l_hora2 
      BackColor       =   &H00FFC0C0&
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
      Left            =   14160
      TabIndex        =   105
      Top             =   9480
      Width           =   2175
   End
   Begin VB.Label l_hora 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Hora:"
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
      Left            =   13440
      TabIndex        =   104
      Top             =   9480
      Width           =   735
   End
   Begin VB.Label l_aviso 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "P O R  P A G A R"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   4440
      TabIndex        =   98
      Top             =   9240
      Visible         =   0   'False
      Width           =   7575
   End
   Begin VB.Label l_fecha 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFC0C0&
      Caption         =   "Label18"
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
      Left            =   14280
      TabIndex        =   78
      Top             =   9240
      Width           =   1575
   End
   Begin VB.Label l_numorden 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
      Caption         =   "l_numorden"
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
      Left            =   17280
      TabIndex        =   60
      Top             =   9240
      Width           =   1335
   End
   Begin VB.Label Label8 
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
      Left            =   13440
      TabIndex        =   26
      Top             =   9240
      Width           =   975
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Orden N°: "
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
      Left            =   16080
      TabIndex        =   0
      Top             =   9240
      Width           =   1215
   End
End
Attribute VB_Name = "registro_orden"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public TipoServicio As String
Public CodProOrden As Integer
Private Sub b_agregar_Click()

Dim temproducto As New ADODB.Recordset
Dim CantPeso As Double
Dim CantPieza As Integer
Dim VarTasaIgv As Double
    
CantPeso = Val(t_peso.Text)
CantPieza = Val(t_cantidad.Text)
    
If (t_peso.Visible = True And CantPeso <> 0 And t_peso.Text <> "" And t_cantidad.Text <> "" And CantPieza <> 0) Or (t_cantidad.Text <> "" And CantPieza <> 0) Then
    
    Call Conn_BDaiosoft
    Set temproducto = New ADODB.Recordset
    
    temproducto.Open "SELECT producto.codigo, descripcion, precio, tasa  FROM producto INNER JOIN impuesto ON producto.impuestogeneral= impuesto.codigo where producto.codigo = '" & Trim(CodProOrden) & "'", Conn_Mysqldb, adOpenDynamic, adLockBatchOptimistic
 
    
    If temproducto.BOF = False And temproducto.EOF = False Then
        temproducto.MoveFirst
        
        fg_prod_orden.AddItem temproducto.Fields(0).Value
        fg_prod_orden.TextMatrix(fg_prod_orden.Rows - 1, 1) = temproducto.Fields(1).Value
        fg_prod_orden.TextMatrix(fg_prod_orden.Rows - 1, 2) = Format(t_precio.Text, "standard")
        fg_prod_orden.TextMatrix(fg_prod_orden.Rows - 1, 3) = t_cantidad.Text
        
        If t_peso.Text <> "" Then
            VarPrecio = Val(t_peso.Text) * Val(t_precio.Text)
            fg_prod_orden.TextMatrix(fg_prod_orden.Rows - 1, 4) = t_peso.Text
        Else
            VarPrecio = Val(t_cantidad.Text) * Val(t_precio.Text)
        End If
        fg_prod_orden.TextMatrix(fg_prod_orden.Rows - 1, 5) = Format(VarPrecio, "standard")
            
        VarTasaIgv = temproducto.Fields(3).Value
        
    End If
    
    
    t_peso.Text = ""
    t_cantidad.Text = ""
    t_precio.Text = ""
    t_peso.Enabled = False
    t_peso.Visible = False
    Label6.Visible = False
    t_cantidad.Enabled = False
    t_precio.Enabled = False
    b_agregar.Enabled = False
    'b_cancelar.Enabled = False
    Call CambioColorBotonOriginal
    
    'totalizacion de la orden
    
    For x = 1 To fg_prod_orden.Rows - 1
        VarAcum = VarAcum + Val(fg_prod_orden.TextMatrix(x, 5))
    Next x
    t_total1.Text = Format(VarAcum, "standard")
    t_total.Text = Format(VarAcum, "standard")
    VarImporte = Format(VarAcum * 100 / (VarTasaIgv + 100), "standard")
    VarIgv = VarAcum - VarImporte
    
    t_igv1.Text = VarIgv
    t_importe1.Text = VarImporte
    o_prepagado.SetFocus
Else
    MsgBox "Ingrese un valor válido."
    
    If t_peso.Visible = True Then
        t_peso.Text = ""
        t_peso.SetFocus
    Else
        t_cantidad.SetFocus
    End If
End If

End Sub

Private Sub b_atras_Click()
fr_prendasvestir.Visible = False
fr_prendas.Visible = True
'b_atras2.Visible = True
'b_sig.Visible = False

t_peso.Text = ""
t_cantidad.Text = ""
t_precio.Text = ""
t_peso.Enabled = False
t_peso.Visible = False
Label6.Visible = False
t_cantidad.Enabled = False
t_precio.Enabled = False
b_agregar.Enabled = False
b_cancelar.Enabled = False

Call CambioColorBotonOriginal

End Sub

Private Sub b_atras2_Click()
fr_prendas.Visible = False

t_peso.Text = ""
t_cantidad.Text = ""
t_precio.Text = ""
t_peso.Enabled = False
t_peso.Visible = False
Label6.Visible = False
t_cantidad.Enabled = False
t_precio.Enabled = False
b_agregar.Enabled = False

Call CambioColorBotonOriginal

End Sub

Private Sub b_atras3_Click()
fr_prendasvestir.Visible = False
b_atras3.Visible = False
b_sig.Visible = False

t_peso.Text = ""
t_cantidad.Text = ""
t_precio.Text = ""
t_peso.Enabled = False
t_peso.Visible = False
Label6.Visible = False
t_cantidad.Enabled = False
t_precio.Enabled = False
b_agregar.Enabled = False

Call CambioColorBotonOriginal

End Sub

Private Sub b_boleta_Click()
Dim TempCabDoc As New ADODB.Recordset
Dim VarSerie As String
Dim VarNrodoc As Long

Dim VarFechaActual As String
Dim VarDia As String
Dim VarMes As String
Dim VarAnio As String
Dim VarAnioImpr As String
Dim FechaVenta As String
Dim FechaVentaImpr As String
Dim HoraVenta As String
Dim AnchoPapel As Integer
Dim VarRuc As String
Dim VarNombre As String
Dim VarDireccion As String
Dim VarTelefono As String
Dim VarLogo As String
Dim ArchivoDet As New ADODB.Recordset
Dim ArchivoCab As New ADODB.Recordset
Dim Temp_cod_afect_Igv As ADODB.Recordset
Dim TempTipoTributos As ADODB.Recordset
Dim temproducto As ADODB.Recordset
Dim tempimpuesto As ADODB.Recordset
Dim VarNroDocumento As Long
Dim VarCodigo As String
Dim VarAbono As Currency
Dim VarRestante As Currency
Dim VarPagado As String
Dim VarImpreso As String
Dim VarsumTotTributosItem As Currency
Dim CodImpGeneral As String
Dim CodImpIsc As String
Dim CodOtImp As String
Dim VarCodAfecIgv As String
Dim VarTasaImp As Double
Dim VarMontoBaseIgvItem As Double
Dim ValidacionCorrecta As Boolean
Dim TempDatosEmpresa As ADODB.Recordset
sumTotTributosItem = 0


'***** CODIGO PARA ESTABLECER EL CORRELATIVO SIGUIENTE DE LA BOLETA *********
If ActualizarOrden = False Then
    Call Conn_BDaiosoft
    Set TempCabDoc = New ADODB.Recordset
    TempCabDoc.Open "SELECT serie, nrodoc FROM cabecera_doc where serie = 'B001' ORDER BY nrodoc", Conn_Mysqldb, adOpenDynamic, adLockBatchOptimistic
    If TempCabDoc.BOF = False And TempCabDoc.EOF = False Then
        
        TempCabDoc.MoveLast
        VarNrodoc = TempCabDoc.Fields(1).Value + 1
        VarSerie = TempCabDoc.Fields(0).Value
    Else
        VarSerie = "B001"
        VarNrodoc = 1
    End If
Else
    Call Conn_BDaiosoft
    Set TempCabDoc = New ADODB.Recordset
    TempCabDoc.Open "SELECT serie, nrodoc FROM cabecera_doc where nroorden = '" & l_numorden.Caption & "' ORDER BY nrodoc", Conn_Mysqldb, adOpenDynamic, adLockBatchOptimistic
    If TempCabDoc.BOF = False And TempCabDoc.EOF = False Then
        TempCabDoc.MoveFirst
        VarNrodoc = TempCabDoc.Fields(1).Value
        VarSerie = TempCabDoc.Fields(0).Value
    End If
End If
'***** FIN CODIGO PARA ESTABLECER EL CORRELATIVO SIGUIENTE DE LA BOLETA *********


'***** CONDICIONAL PARA DETERMINAR SI ES UNA ORDEN NUEVA O ACTUALIZACION DE UNA EXISTENTE ********
If ActualizarOrden = True Then

    If Val(t_efectivo.Text) < Val(t_restante.Text) Then
            
        MsgBox " Efectivo insuficiente."
            
    Else

        '*********** INICIO CONSULTA PARA ACTUALIZAR ORDEN ********************
        
        Call Conn_BDaiosoft
        Conn_Mysqldb.Execute "UPDATE cabecera_doc SET pagado = 's', abono = " & t_total.Text & ", restante = 0 WHERE nroorden = " & l_numorden.Caption & ""
        
        '*********** FIN DE CONSULTA PARA ACTUALIZAR ORDEN ********************
        TituloDoc = "BOLETA DE VENTA ELECTRONICA"
        VarPagado = "s"
        Call ImprimirFormato80mm(TituloDoc, FechaVentaImpr, HoraVenta, VarSerie, VarNrodoc, VarPagado, VarAbono, VarRestante)
            
        '*********** INICIO DE CODIGO PARA IMPRIMIR BOLETA***********
            
         
            
        '**************** FIN DE CODIGO DE IMPRESION DE BOLETA***************
        
            '******* INICIO DE CODIGO PARA LIMPIAR FORMULARIO *****************
            
            fg_prod_orden.Clear
            fg_prod_orden.FormatString = "Código|Descripción                 |Precio U.|Cant.| Peso|    Precio  "
            fg_prod_orden.Rows = 1
            VarAcum = 0
            
            c_tipodoc.ListIndex = 0
            t_nombrers.Text = ""
            t_telefono.Text = ""
            t_email.Text = ""
            t_nrodoc.Text = ""
            t_direccion.Text = ""
            t_peso.Text = ""
            t_cantidad.Text = ""
            t_precio.Text = ""
            t_importe1.Text = ""
            t_total1.Text = ""
            t_igv1.Text = ""
            t_total.Text = ""
            t_efectivo.Text = ""
            t_vuelto.Text = ""
            t_xcentajedesc.Text = ""
            t_mtodesc.Text = ""
            o_prepagado.Value = True
            Frame1.Enabled = True
            Frame2.Enabled = True
            l_abono.Visible = False
            t_abono.Visible = False
            l_restante.Visible = False
            t_restante.Visible = False
            Frame4.Left = 360
            Frame4.Top = 7320
            Frame4.Visible = True
            l_total.Top = 360
            l_efectivo.Top = 960
            l_vuelto.Top = 1560
            l_total.Left = 3240
            l_efectivo.Left = 2760
            l_vuelto.Left = 3000
            t_total.Left = 4320
            t_efectivo.Left = 4320
            t_vuelto.Left = 4320
            t_total.Top = 360
            t_efectivo.Top = 960
            t_vuelto.Top = 1560
            l_xcentajedesc.Visible = True
            t_xcentajedesc.Visible = True
            l_mtodesc.Visible = True
            t_mtodesc.Visible = True
            Timer1.Enabled = False
            Timer2.Enabled = False
            l_aviso.Visible = False
            l_hora2.Caption = ""
            l_fecha.Caption = Mid(Now, 1, 10)
            ActualizarOrden = False
            b_cliente.SetFocus
            '******* FIN DE CODIGO PARA LIMPIAR FORMULARIO *****************
            
            
             '**** INICIO CODIGO PARA CARGAR EL NUMERO DE ORDEN SIGUIENTE****
            Call Conn_BDaiosoft
            Set TempCabDoc = New ADODB.Recordset
            TempCabDoc.Open "SELECT nroorden FROM cabecera_doc  ORDER BY nroorden", Conn_Mysqldb, adOpenDynamic, adLockBatchOptimistic
            If TempCabDoc.BOF = False And TempCabDoc.EOF = False Then
            TempCabDoc.MoveLast
            numorden = TempCabDoc.Fields(0).Value + 1
            l_numorden.Caption = numorden
            Else
    
                l_numorden.Caption = 1
            End If
            '**** FIN CODIGO PARA CARGAR EL NUMERO DE ORDEN SIGUIENTE****
    
    End If

Else

If t_nombrers.Text = "" Or t_telefono.Text = "" Then
    MsgBox "debe completar los datos del cliente para procesar la orden"
Else
    If t_total.Text = "" Or t_total.Text = "0" Then
        MsgBox "debe agregar al menos un producto para procesar la orden"
     Else
        'If op_efectivo.Value = False Or op_cheque.Value = False Or op_transferencia.Value = False Or op_puntodeventa.
        'MsgBox "Debe seleccionar una forma de pago"
        'Else
        msg = MsgBox("¿Esta de acuerdo con los cambios?", vbQuestion + vbYesNo)
        If msg = vbYes Then
        
       '*****VALIDACION DE LAS FORMAS DE PAGO, PREPAGADO, ABONO, PAGO AL RECOGER
        If o_prepagado.Value = True Then
            If Val(t_efectivo.Text) < Val(t_total.Text) Then
                ValidacionCorrecta = False
                MsgBox " Efectivo insuficiente."
            Else
                ValidacionCorrecta = True
                VarPagado = "s"
                VarAbono = Val(t_total.Text)
                VarRestante = 0
            End If
        End If
        
        If o_abono.Value = True Then
            If t_efectivo.Text = "" Or Val(t_efectivo.Text) = 0 Then
                ValidacionCorrecta = False
                MsgBox "Ingrese el pago parcial."
            Else
                ValidacionCorrecta = True
                VarAbono = Val(t_efectivo.Text)
                VarRestante = Val(t_total.Text) - Val(t_efectivo.Text)
                VarPagado = "n"
            End If
        End If
        
        If o_alrecojo.Value = True Then
            ValidacionCorrecta = True
            VarPagado = "n"
            VarAbono = 0
            VarRestante = Val(t_total.Text)
        End If
                                                            
       '***** FIN DE CODIGO PARA VALIDACION DE LAS FORMAS DE PAGO, PREPAGADO, ABONO, PAGO AL RECOGER
       
       
        If ValidacionCorrecta = True Then
        
            '##### AQUI INICIA EL CODIGO PARA ALMACENAR EL DETALLE DEL DOCUMENTO EN LA TABLA det_doc #####
            For x = 1 To fg_prod_orden.Rows - 1
    
                'consulta para recuperar los codigos de tipos de impuesto IGV, ISC, otros impuestos
                Call Conn_BDaiosoft
                Set temproducto = New ADODB.Recordset
                temproducto.Open "SELECT * FROM producto where codigo = '" & fg_prod_orden.TextMatrix(x, 0) & "'", Conn_Mysqldb, adOpenDynamic, adLockBatchOptimistic
                If temproducto.BOF = False And temproducto.EOF = False Then
                    temproducto.MoveFirst
                    CodImpGeneral = temproducto.Fields(6).Value
                    CodImpIsc = temproducto.Fields(7).Value
                    CodOtImp = temproducto.Fields(8).Value
                    CodIcbper = temproducto.Fields(9).Value
                End If
    
        
                If CodImpIsc = "01" Then
            
                    VarCodImpIsc = "-"
                    VarnomTributoIscItem = ""
                    Varmontoiscitem = 0
                    Varmontobaseiscitem = 0
                    Varcodtiptributoiscitem = ""
                    Vartipsisisc = ""
                    Varporcentajeiscitem = 0
                Else
            
                End If
        
        
                If CodOtImp = "01" Then
        
                    VarcodTriOtroItem = "-"
                    VarmtoTriOtroItem = 0
                    VarmtoBaseTriOtroItem = 0
                    VarnomTributoIOtroItem = ""
                    VarcodTipTributoIOtroItem = ""
                    VarporTriOtroItem = "0"
                Else
        
        
                End If
        
        
                '****seccion de codigo para realizar los calculos del impueso ICBPER por item****
                If CodIcbper = "01" Then
        
                    Varcodtriicbper = "-"
                    Varmtotriicbperitem = 0
                    Varcantboltriicbperitem = 0
                    Varnomtriicbperitem = ""
                    Varcodtiptributoicbperitem = ""
                    Varmtotributoicbperunidad = 0
                Else
            
                    'Varcodtriicbper = CodIcbper
                    'consulta para extraer nombre de impuesto  y tasa
                    'Set tempimpuesto = New ADODB.Recordset
                    'tempimpuesto.Open "SELECT denominacion, tasa FROM impuesto WHERE codigo = '" & CodIcbper & "'", conn_mysqldb, adOpenDynamic, adLockBatchOptimistic
                    'If tempimpuesto.BOF = False And tempimpuesto.EOF = False Then
                    '    tempimpuesto.MoveFirst
                    '    Varnomtriicbperitem = tempimpuesto.Fields(0).Value
                    '    Varmtotributoicbperunidad = tempimpuesto.Fields(1).Value
                    'End If
                    'Varcantboltriicbperitem = Val(fg_detallefactura.TextMatrix(contador, 2))
                    ' Varmtotriicbperitem = Val(Varmtotributoicbperunidad) * Val(fg_detallefactura.TextMatrix(contador, 2))
                    
                    'consulta para extraer de la tabla tipo_tributos el codigo internacional del impuesto
                    'Set TempTipoTributos = New ADODB.Recordset
                    'TempTipoTributos.Open "SELECT codinternacional FROM tipo_tributos WHERE codigo = '" & CodIcbper & "'", conn_mysqldb, adOpenDynamic, adLockBatchOptimistic
                    'If TempTipoTributos.BOF = False And TempTipoTributos.EOF = False Then
                    '    TempTipoTributos.MoveFirst
                    '    Varcodtiptributoicbperitem = TempTipoTributos.Fields(0).Value
                    'End If
        
                End If
                '****fin de seccion de codigo para realizar los calculos del impueso ICBPER por item****
        
        
                'consulta para recuperar el codigo de tipo de afectacion del IGV
                'Call conn_BDaiosoft
                Set Temp_cod_afect_Igv = New ADODB.Recordset
                 Temp_cod_afect_Igv.Open "SELECT codigo FROM cod_afectacion_igv WHERE codigo_tributo = '" & CodImpGeneral & "'", Conn_Mysqldb, adOpenDynamic, adLockBatchOptimistic
                If Temp_cod_afect_Igv.BOF = False And Temp_cod_afect_Igv.EOF = False Then
                    Temp_cod_afect_Igv.MoveFirst
                    VarCodAfecIgv = Temp_cod_afect_Igv.Fields(0).Value
                End If
       
                'consulta para recuper la tasa del impuesto correspondiente y nombre de impuesto (IGV,EXP,INA,GRA,IVAP)
                Set tempimpuesto = New ADODB.Recordset
                tempimpuesto.Open "SELECT denominacion, tasa FROM impuesto WHERE codigo = '" & CodImpGeneral & "'", Conn_Mysqldb, adOpenDynamic, adLockBatchOptimistic
                If tempimpuesto.BOF = False And tempimpuesto.EOF = False Then
                    tempimpuesto.MoveFirst
                    VarDenominacion = tempimpuesto.Fields(0).Value
                    VarTasaImp = tempimpuesto.Fields(1).Value
                End If
       
                'consulta para extraer de la tabla tipo_tributos el codigo internacional del impuesto
                Set TempTipoTributos = New ADODB.Recordset
                TempTipoTributos.Open "SELECT codinternacional FROM tipo_tributos WHERE codigo = '" & CodImpGeneral & "'", Conn_Mysqldb, adOpenDynamic, adLockBatchOptimistic
                If TempTipoTributos.BOF = False And TempTipoTributos.EOF = False Then
                    TempTipoTributos.MoveFirst
                    VarCodInternacional = TempTipoTributos.Fields(0).Value
                End If
       
       
       
                'calculo de la sumatoria de todos los tributos por item
                VarsumTotTributosItem = Val(fg_prod_orden.TextMatrix(x, 5)) * 0.18
                VarMontoBaseIgvItem = Val(fg_prod_orden.TextMatrix(x, 5)) * 100 / 118
                
                If fg_prod_orden.TextMatrix(x, 4) = "" Then
                    VarPeso = 0
                Else
                    VarPeso = fg_prod_orden.TextMatrix(x, 4)
                End If
                
                'consulta para insertar el detalle del documento de venta en la tabla det_doc
                'Call conn_BDaiosoft

                Conn_Mysqldb.Execute "INSERT INTO det_doc SET serie = '" & VarSerie & "'," _
                & "nroorden = " & l_numorden.Caption & ", nrodoc = " & VarNrodoc & ", codproducto = '" & fg_prod_orden.TextMatrix(x, 0) & "'," _
                & "codunidadmedida = 'NIU', cantidaditem = " & fg_prod_orden.TextMatrix(x, 3) & "," _
                & "peso = " & VarPeso & "," _
                & "descripitem = '" & fg_prod_orden.TextMatrix(x, 1) & "', preciounititem = " & fg_prod_orden.TextMatrix(x, 2) & "," _
                & "sumtottributositem = " & VarsumTotTributosItem & ", codtriigv = '" & CodImpGeneral & "'," _
                & "montoigvitem = " & VarsumTotTributosItem & ", montobaseigvitem = " & VarMontoBaseIgvItem & "," _
                & "nombretributoigvxitem = '" & VarDenominacion & "', codtiptribigvitem = '" & VarCodInternacional & "'," _
                & "tipafeigv = '" & VarCodAfecIgv & "', porcentajeigv = " & VarTasaImp & "," _
                & "codtriisc = '" & VarCodImpIsc & "', montoiscitem = " & Varmontoiscitem & "," _
                & "montobaseiscitem = " & Varmontobaseiscitem & ", nombretributoiscxitem = '" & VarnomTributoIscItem & "'," _
                & "codtiptributoiscitem = '" & Varcodtiptributoiscitem & "', tipsisisc = '" & Vartipsisisc & "'," _
                & "porcentajeiscitem = " & Varporcentajeiscitem & ", codotrostributos = '" & VarcodTriOtroItem & "'," _
                & "montootrotriitem = " & VarmtoTriOtroItem & ", montobaseotrotriitem = " & VarmtoBaseTriOtroItem & "," _
                & "nombreotrotributoxitem = '" & VarnomTributoIOtroItem & "', codtipotrotributositem = '" & VarcodTipTributoIOtroItem & "'," _
                & "porcentajeotrotribitem = " & VarporTriOtroItem & "," _
                & "codtriicbper = '" & Varcodtriicbper & "'," _
                & "mtotriicbperitem = " & Varmtotriicbperitem & "," _
                & "cantboltriicbperitem = " & Varcantboltriicbperitem & "," _
                & "nomtriicbperitem = '" & Varnomtriicbperitem & "'," _
                & "codtiptributoicbperitem = '" & Varcodtiptributoicbperitem & "'," _
                & "mtotributoicbperunidad = " & Varmtotributoicbperunidad & "," _
                & "montototalxitem = " & VarMontoBaseIgvItem & "," _
                & "montototalxitemconimpuestos = " & fg_prod_orden.TextMatrix(x, 5) & ""
                
                
            Next x
            '##### AQUI FINALIZA EL CODIGO PARA ALMACENAR EL DETALLE DEL DOCUMENTO EN LA TABLA det_doc #####
    
    
            '##### AQUI INICIA EL CODIGO PARA ALMACENAR LOS DATOS DEL DOCUMENTO EN LA TABLA cabecera_doc #####
       
            'codigo para construir la fecha y hora actual
    
            VarFechaActual = Now
            VarDia = Mid(VarFechaActual, 1, 2)
            VarMes = Mid(VarFechaActual, 4, 2)
            VarAnio = Mid(VarFechaActual, 7, 4)
            VarAnioImpr = Mid(VarFechaActual, 9, 2)
    
            FechaVenta = VarAnio + "-" + VarMes + "-" + VarDia
            FechaVentaImpr = VarDia + "/" + VarMes + "/" + VarAnioImpr
            mihoracompleta = Now
            HoraVenta = Mid(mihoracompleta, 12, 8)
            
    
            'codigo para determinar el tipo de documento a almacenar
            'If c_tipodocumento.Text = "BOLETA" Then
            '    Vartipodoc = "03"
            'End If
            'If c_tipodocumento.Text = "FACTURA" Then
            '    Vartipodoc = "01"
            'End If
            
            
            If t_xcentajedesc.Text = "" Then
                VarXcentDesc = 0
            Else
                VarXcentDesc = Val(t_xcentajedesc.Text)
            End If
    
            'consulta para la insercion de datos en la tabla cabecera_doc
            Conn_Mysqldb.Execute "INSERT INTO cabecera_doc SET serie= '" & VarSerie & "'," _
            & "nroorden = " & l_numorden.Caption & "," _
            & "nrodoc = " & VarNrodoc & "," _
            & "tipodoc = '00'," _
            & "tipooper = '0101'," _
            & "firmadig = '-'," _
            & "fechaem = '" & FechaVenta & "'," _
            & "horaem = '" & HoraVenta & "'," _
            & "fechavenci = '" & FechaVenta & "'," _
            & "codlocalemisor = '0000'," _
            & "tipdocusuario = '" & c_tipodoc.Text & "'," _
            & "numdocusuario = '" & t_nrodoc.Text & "'," _
            & "nombrers = '" & t_nombrers.Text & "'," _
            & "tipmoneda = 'PEN'," _
            & "sumtottributos = " & t_igv1.Text & "," _
            & "sumtotvalventa = " & t_importe1.Text & "," _
            & "sumprecioventa = " & t_total1.Text & "," _
            & "sumdesctotal = " & VarXcentDesc & "," _
            & "sumotroscargos = 0," _
            & "sumtotalanticipos = 0," _
            & "sumimpventa = " & t_total1.Text & "," _
            & "ublversionld = '2.1'," _
            & "customizationld = '2.0'," _
            & "pagado = '" & VarPagado & "', abono = " & VarAbono & ", restante = " & VarRestante & "," _
            & "impreso = 's'"
    
    
            '##### AQUI FINALIZA EL CODIGO PARA ALMACENAR LOS DATOS DEL DOCUMENTO EN LA TABLA cabecera_doc #####
            
            
            '**********INICIO DE CODIGO PARA IMPRIMIR LA BOLETA***********
            
            TituloDoc = "BOLETA DE VENTA ELECTRONICA"

            Call ImprimirFormato80mm(TituloDoc, FechaVentaImpr, HoraVenta, VarSerie, VarNrodoc, VarPagado, VarAbono, VarRestante)

            
            '**************** FIN DE CODIGO DE IMPRESION DE BOLETA********
            
            
             '**** CODIGO PARA CARGAR EL NUMERO DE ORDEN SIGUIENTE****
            Call Conn_BDaiosoft
            Set TempCabDoc = New ADODB.Recordset
            TempCabDoc.Open "SELECT nroorden FROM cabecera_doc  ORDER BY nroorden", Conn_Mysqldb, adOpenDynamic, adLockBatchOptimistic
            If TempCabDoc.BOF = False And TempCabDoc.EOF = False Then
            TempCabDoc.MoveLast
            numorden = TempCabDoc.Fields(0).Value + 1
            l_numorden.Caption = numorden
            Else
    
                l_numorden.Caption = 1
            End If
            '**** FIN CODIGO PARA CARGAR EL NUMERO DE ORDEN SIGUIENTE****
            
            
            '******* INICIO DE CODIGO PARA LIMPIAR FORMULARIO *****************
            
            fg_prod_orden.Clear
            fg_prod_orden.FormatString = "Código|Descripción                 |Precio U.|Cant.| Peso|    Precio  "
            fg_prod_orden.Rows = 1
            VarAcum = 0
            
            c_tipodoc.ListIndex = 0
            t_nombrers.Text = ""
            t_telefono.Text = ""
            t_email.Text = ""
            t_nrodoc.Text = ""
            t_direccion.Text = ""
            t_peso.Text = ""
            t_cantidad.Text = ""
            t_precio.Text = ""
            t_importe1.Text = ""
            t_total1.Text = ""
            t_igv1.Text = ""
            t_total.Text = ""
            t_efectivo.Text = ""
            t_vuelto.Text = ""
            t_xcentajedesc.Text = ""
            t_mtodesc.Text = ""
            o_prepagado.Value = True
            t_peso.Enabled = False
            t_peso.Visible = False
            Label6.Visible = False
            t_cantidad.Enabled = False
            t_precio.Enabled = False
            b_agregar.Enabled = False
            'b_cancelar.Enabled = False
            Call CambioColorBotonOriginal
            b_cliente.SetFocus
           '******* FIN DE CODIGO PARA LIMPIAR FORMULARIO *****************
            
            
            
        End If
        End If
    
    End If
End If
End If


End Sub

Private Sub b_boton1_Click()
NombreBoton = b_boton1.Name
If TipoServicio = "lavado" Then
     CodProOrden = 1
End If
If TipoServicio = "secado" Then
    CodProOrden = 2
End If

If TipoServicio = "lavadoycentrifugado" Then
    CodProOrden = 63
End If

Call CambioColorBoton
Call HabilitarTexbox
t_precio.Text = DevolverPrecio(CodProOrden)
t_peso.Visible = True
Label6.Visible = True
t_peso.Enabled = True
t_peso.SetFocus
End Sub

Private Sub b_boton1_GotFocus()
Call CambioColorBotonOriginal

End Sub

Private Sub b_boton10_Click()
NombreBoton = b_boton10.Name

If TipoServicio = "lavado" Then
     CodProOrden = 20
End If
If TipoServicio = "secado" Then
    CodProOrden = 76
End If

If TipoServicio = "lavadoycentrifugado" Then
    CodProOrden = 47
End If

Call CambioColorBoton
Call HabilitarTexbox
t_precio.Text = DevolverPrecio(CodProOrden)
End Sub

Private Sub b_boton10_GotFocus()
Call CambioColorBotonOriginal
End Sub

Private Sub b_boton11_Click()
NombreBoton = b_boton11.Name

If TipoServicio = "lavado" Then
     CodProOrden = 11
End If
If TipoServicio = "secado" Then
    CodProOrden = 78
End If

If TipoServicio = "lavadoycentrifugado" Then
    CodProOrden = 49
End If

Call CambioColorBoton
Call HabilitarTexbox
t_precio.Text = DevolverPrecio(CodProOrden)
End Sub

Private Sub b_boton11_GotFocus()
Call CambioColorBotonOriginal
End Sub

Private Sub b_boton12_Click()
NombreBoton = b_boton12.Name


If TipoServicio = "lavado" Then
     CodProOrden = 15
End If
If TipoServicio = "secado" Then
    CodProOrden = 83
End If

If TipoServicio = "lavadoycentrifugado" Then
    CodProOrden = 82
End If

Call CambioColorBoton
Call HabilitarTexbox
t_precio.Text = DevolverPrecio(CodProOrden)
End Sub

Private Sub b_boton12_GotFocus()
Call CambioColorBotonOriginal
End Sub

Private Sub b_boton13_Click()
NombreBoton = b_boton13.Name

If TipoServicio = "lavado" Then
     CodProOrden = 21
End If
If TipoServicio = "secado" Then
    CodProOrden = 77
End If

If TipoServicio = "lavadoycentrifugado" Then
    CodProOrden = 48
End If

Call CambioColorBoton
Call HabilitarTexbox
t_precio.Text = DevolverPrecio(CodProOrden)
End Sub

Private Sub b_boton13_GotFocus()
Call CambioColorBotonOriginal
End Sub

Private Sub b_boton14_Click()
NombreBoton = b_boton14.Name

If TipoServicio = "lavado" Then
     CodProOrden = 12
End If
If TipoServicio = "secado" Then
    CodProOrden = 79
End If

If TipoServicio = "lavadoycentrifugado" Then
    CodProOrden = 50
End If

Call CambioColorBoton
Call HabilitarTexbox
t_precio.Text = DevolverPrecio(CodProOrden)
End Sub

Private Sub b_boton14_GotFocus()
Call CambioColorBotonOriginal
End Sub

Private Sub b_boton15_Click()
NombreBoton = b_boton15.Name

If TipoServicio = "lavado" Then
     CodProOrden = 13
End If
If TipoServicio = "secado" Then
    CodProOrden = 80
End If

If TipoServicio = "lavadoycentrifugado" Then
    CodProOrden = 51
End If

Call CambioColorBoton
Call HabilitarTexbox
t_precio.Text = DevolverPrecio(CodProOrden)
End Sub

Private Sub b_boton15_GotFocus()
Call CambioColorBotonOriginal
End Sub

Private Sub b_boton16_Click()

NombreBoton = b_boton16.Name

If TipoServicio = "lavado" Then
     CodProOrden = 34
End If
If TipoServicio = "secado" Then
    CodProOrden = 81
End If

If TipoServicio = "lavadoycentrifugado" Then
    CodProOrden = 52
End If

Call CambioColorBoton
Call HabilitarTexbox
t_precio.Text = DevolverPrecio(CodProOrden)

End Sub

Private Sub b_boton16_GotFocus()
Call CambioColorBotonOriginal
End Sub

Private Sub b_boton17_Click()

NombreBoton = b_boton17.Name

If TipoServicio = "lavado" Then
     CodProOrden = 85
End If
If TipoServicio = "secado" Then
    CodProOrden = 65
End If

If TipoServicio = "lavadoycentrifugado" Then
    CodProOrden = 36
End If

If TipoServicio = "lavadoseco" Then
    CodProOrden = 53
End If

If TipoServicio = "teñido" Then
    CodProOrden = 27
End If

Call CambioColorBoton
Call HabilitarTexbox
t_precio.Text = DevolverPrecio(CodProOrden)

End Sub

Private Sub b_boton17_GotFocus()
Call CambioColorBotonOriginal
End Sub

Private Sub b_boton18_Click()

NombreBoton = b_boton18.Name

If TipoServicio = "lavado" Then
     CodProOrden = 84
End If
If TipoServicio = "secado" Then
    CodProOrden = 64
End If

If TipoServicio = "lavadoycentrifugado" Then
    CodProOrden = 35
End If

If TipoServicio = "lavadoseco" Then
    CodProOrden = 54
End If

If TipoServicio = "teñido" Then
    CodProOrden = 26
End If

Call CambioColorBoton
Call HabilitarTexbox
t_precio.Text = DevolverPrecio(CodProOrden)

End Sub

Private Sub b_boton18_GotFocus()
Call CambioColorBotonOriginal
End Sub

Private Sub b_boton19_Click()

NombreBoton = b_boton19.Name

If TipoServicio = "lavado" Then
     CodProOrden = 91
End If
If TipoServicio = "secado" Then
    CodProOrden = 106
End If

If TipoServicio = "lavadoycentrifugado" Then
    CodProOrden = 99
End If

If TipoServicio = "lavadoseco" Then
    CodProOrden = 18
End If

If TipoServicio = "teñido" Then
    CodProOrden = 33
End If

Call CambioColorBoton
Call HabilitarTexbox
t_precio.Text = DevolverPrecio(CodProOrden)

End Sub

Private Sub b_boton19_GotFocus()
Call CambioColorBotonOriginal
End Sub

Private Sub b_boton2_Click()

NombreBoton = b_boton2.Name

If TipoServicio = "lavado" Then
     CodProOrden = 5
End If
If TipoServicio = "secado" Then
    CodProOrden = 70
End If

If TipoServicio = "lavadoycentrifugado" Then
    CodProOrden = 41
End If


Call CambioColorBoton
Call HabilitarTexbox
t_precio.Text = DevolverPrecio(CodProOrden)

End Sub

Private Sub b_boton2_GotFocus()
Call CambioColorBotonOriginal
End Sub

Private Sub b_boton20_Click()

NombreBoton = b_boton20.Name

If TipoServicio = "lavado" Then
     CodProOrden = 89
End If
If TipoServicio = "secado" Then
    CodProOrden = 104
End If

If TipoServicio = "lavadoycentrifugado" Then
    CodProOrden = 97
End If

If TipoServicio = "lavadoseco" Then
    CodProOrden = 56
End If

If TipoServicio = "teñido" Then
    CodProOrden = 31
End If

Call CambioColorBoton
Call HabilitarTexbox
t_precio.Text = DevolverPrecio(CodProOrden)

End Sub

Private Sub b_boton20_GotFocus()
Call CambioColorBotonOriginal
End Sub

Private Sub b_boton21_Click()

NombreBoton = b_boton21.Name

If TipoServicio = "lavado" Then
     CodProOrden = 90
End If
If TipoServicio = "secado" Then
    CodProOrden = 105
End If

If TipoServicio = "lavadoycentrifugado" Then
    CodProOrden = 98
End If

If TipoServicio = "lavadoseco" Then
    CodProOrden = 57
End If

If TipoServicio = "teñido" Then
    CodProOrden = 32
End If

Call CambioColorBoton
Call HabilitarTexbox
t_precio.Text = DevolverPrecio(CodProOrden)

End Sub

Private Sub b_boton21_GotFocus()
Call CambioColorBotonOriginal
End Sub

Private Sub b_boton22_Click()

NombreBoton = b_boton22.Name

If TipoServicio = "lavadoseco" Then
    CodProOrden = 58
End If

If TipoServicio = "lavadoycentrifugado" Then
    CodProOrden = 38
End If

If TipoServicio = "lavado" Then
     CodProOrden = 86
End If

If TipoServicio = "secado" Then
    CodProOrden = 67
End If

If TipoServicio = "teñido" Then
    CodProOrden = 28
End If

Call CambioColorBoton
Call HabilitarTexbox
t_precio.Text = DevolverPrecio(CodProOrden)

End Sub

Private Sub b_boton22_GotFocus()
Call CambioColorBotonOriginal
End Sub

Private Sub b_boton23_Click()

NombreBoton = b_boton23.Name

If TipoServicio = "lavadoseco" Then
    CodProOrden = 109
End If

If TipoServicio = "lavadoycentrifugado" Then
    CodProOrden = 37
End If

If TipoServicio = "lavado" Then
     CodProOrden = 94
End If

If TipoServicio = "secado" Then
    CodProOrden = 66
End If

If TipoServicio = "teñido" Then
    CodProOrden = 110
End If

Call CambioColorBoton
Call HabilitarTexbox
t_precio.Text = DevolverPrecio(CodProOrden)

End Sub

Private Sub b_boton23_GotFocus()
Call CambioColorBotonOriginal
End Sub

Private Sub b_boton24_Click()

NombreBoton = b_boton24.Name

If TipoServicio = "lavadoseco" Then
   CodProOrden = 25
End If

If TipoServicio = "lavadoycentrifugado" Then
    CodProOrden = 116
End If

If TipoServicio = "lavado" Then
     CodProOrden = 117
End If

If TipoServicio = "secado" Then
    CodProOrden = 118
End If

If TipoServicio = "teñido" Then
    CodProOrden = 119
End If

Call CambioColorBoton
Call HabilitarTexbox
t_precio.Text = DevolverPrecio(CodProOrden)

End Sub

Private Sub b_boton24_GotFocus()
Call CambioColorBotonOriginal
End Sub

Private Sub b_boton25_Click()

NombreBoton = b_boton25.Name

If TipoServicio = "lavadoseco" Then
   CodProOrden = 24
End If

If TipoServicio = "lavadoycentrifugado" Then
    CodProOrden = 100
End If

If TipoServicio = "lavado" Then
     CodProOrden = 92
End If

If TipoServicio = "secado" Then
    CodProOrden = 107
End If

If TipoServicio = "teñido" Then
    CodProOrden = 61
End If

Call CambioColorBoton
Call HabilitarTexbox
t_precio.Text = DevolverPrecio(CodProOrden)

End Sub

Private Sub b_boton25_GotFocus()
Call CambioColorBotonOriginal
End Sub

Private Sub b_boton26_Click()

NombreBoton = b_boton26.Name

If TipoServicio = "lavadoseco" Then
   CodProOrden = 59
End If

If TipoServicio = "lavadoycentrifugado" Then
    CodProOrden = 101
End If

If TipoServicio = "lavado" Then
     CodProOrden = 93
End If

If TipoServicio = "secado" Then
    CodProOrden = 108
End If

If TipoServicio = "teñido" Then
    CodProOrden = 62
End If

Call CambioColorBoton
Call HabilitarTexbox
t_precio.Text = DevolverPrecio(CodProOrden)

End Sub

Private Sub b_boton26_GotFocus()
Call CambioColorBotonOriginal
End Sub

Private Sub b_boton27_Click()

NombreBoton = b_boton27.Name

If TipoServicio = "lavadoseco" Then
   CodProOrden = 16
End If

Call CambioColorBoton
Call HabilitarTexbox
t_precio.Text = DevolverPrecio(CodProOrden)

End Sub

Private Sub b_boton27_GotFocus()
Call CambioColorBotonOriginal
End Sub

Private Sub b_boton28_Click()

NombreBoton = b_boton28.Name

If TipoServicio = "lavadoseco" Then
   CodProOrden = 17
End If

Call CambioColorBoton
Call HabilitarTexbox
t_precio.Text = DevolverPrecio(CodProOrden)

End Sub

Private Sub b_boton28_GotFocus()
Call CambioColorBotonOriginal
End Sub

Private Sub b_boton29_Click()
NombreBoton = b_boton29.Name

If TipoServicio = "lavadoseco" Then
   CodProOrden = 60
End If

Call CambioColorBoton
Call HabilitarTexbox
t_precio.Text = DevolverPrecio(CodProOrden)
End Sub

Private Sub b_boton29_GotFocus()
Call CambioColorBotonOriginal
End Sub

Private Sub b_boton3_Click()
NombreBoton = b_boton3.Name

If TipoServicio = "lavado" Then
     CodProOrden = 6
End If
If TipoServicio = "secado" Then
    CodProOrden = 71
End If

If TipoServicio = "lavadoycentrifugado" Then
    CodProOrden = 42
End If

Call CambioColorBoton
Call HabilitarTexbox
t_precio.Text = DevolverPrecio(CodProOrden)
End Sub

Private Sub b_boton3_GotFocus()
Call CambioColorBotonOriginal
End Sub

Private Sub b_boton30_Click()

NombreBoton = b_boton30.Name

If TipoServicio = "lavadoseco" Then
   CodProOrden = 23
End If

Call CambioColorBoton
Call HabilitarTexbox
t_precio.Text = DevolverPrecio(CodProOrden)

End Sub

Private Sub b_boton30_GotFocus()
Call CambioColorBotonOriginal
End Sub

Private Sub b_boton31_Click()
NombreBoton = b_boton31.Name

If TipoServicio = "lavadoseco" Then
   CodProOrden = 55
End If

Call CambioColorBoton
Call HabilitarTexbox
t_precio.Text = DevolverPrecio(CodProOrden)
End Sub

Private Sub b_boton31_GotFocus()
Call CambioColorBotonOriginal
End Sub

Private Sub b_boton32_Click()
NombreBoton = b_boton32.Name

If TipoServicio = "lavado" Then
    CodProOrden = 14
End If

Call CambioColorBoton
Call HabilitarTexbox
t_precio.Text = DevolverPrecio(CodProOrden)
End Sub

Private Sub b_boton32_GotFocus()
Call CambioColorBotonOriginal
End Sub

Private Sub b_boton33_Click()

NombreBoton = b_boton33.Name

If TipoServicio = "lavadoseco" Then
    CodProOrden = 121
End If

If TipoServicio = "lavadoycentrifugado" Then
    CodProOrden = 95
End If

If TipoServicio = "lavado" Then
     CodProOrden = 87
End If

If TipoServicio = "secado" Then
    CodProOrden = 102
End If

If TipoServicio = "teñido" Then
    CodProOrden = 29
End If

Call CambioColorBoton
Call HabilitarTexbox
t_precio.Text = DevolverPrecio(CodProOrden)

End Sub

Private Sub b_boton33_GotFocus()
Call CambioColorBotonOriginal
End Sub

Private Sub b_boton34_Click()

NombreBoton = b_boton34.Name

If TipoServicio = "lavadoseco" Then
    CodProOrden = 120
End If

If TipoServicio = "lavadoycentrifugado" Then
    CodProOrden = 96
End If

If TipoServicio = "lavado" Then
     CodProOrden = 88
End If

If TipoServicio = "secado" Then
    CodProOrden = 103
End If

If TipoServicio = "teñido" Then
    CodProOrden = 30
End If

Call CambioColorBoton
Call HabilitarTexbox
t_precio.Text = DevolverPrecio(CodProOrden)

End Sub

Private Sub b_boton34_GotFocus()
Call CambioColorBotonOriginal
End Sub

Private Sub b_boton35_Click()
NombreBoton = b_boton35.Name

If TipoServicio = "lavado" Then
    CodProOrden = 22
End If

Call CambioColorBoton
Call HabilitarTexbox
t_precio.Text = DevolverPrecio(CodProOrden)

End Sub

Private Sub b_boton35_GotFocus()
Call CambioColorBotonOriginal
End Sub

Private Sub b_boton36_Click()

NombreBoton = b_boton36.Name

If TipoServicio = "lavadoseco" Then
    CodProOrden = 111
End If

If TipoServicio = "lavadoycentrifugado" Then
    CodProOrden = 112
End If

If TipoServicio = "lavado" Then
     CodProOrden = 113
End If

If TipoServicio = "secado" Then
    CodProOrden = 114
End If

If TipoServicio = "teñido" Then
    CodProOrden = 115
End If

Call CambioColorBoton
Call HabilitarTexbox
t_precio.Text = DevolverPrecio(CodProOrden)

End Sub

Private Sub b_boton36_GotFocus()
Call CambioColorBotonOriginal
End Sub

Private Sub b_boton4_Click()
NombreBoton = b_boton4.Name

If TipoServicio = "lavado" Then
     CodProOrden = 3
End If
If TipoServicio = "secado" Then
    CodProOrden = 68
End If

If TipoServicio = "lavadoycentrifugado" Then
    CodProOrden = 39
End If

Call CambioColorBoton
Call HabilitarTexbox
t_precio.Text = DevolverPrecio(CodProOrden)
End Sub

Private Sub b_boton4_GotFocus()
Call CambioColorBotonOriginal
End Sub

Private Sub b_boton5_Click()
NombreBoton = b_boton5.Name

If TipoServicio = "lavado" Then
     CodProOrden = 4
End If
If TipoServicio = "secado" Then
    CodProOrden = 69
End If

If TipoServicio = "lavadoycentrifugado" Then
    CodProOrden = 40
End If

Call CambioColorBoton
Call HabilitarTexbox
t_precio.Text = DevolverPrecio(CodProOrden)
End Sub

Private Sub b_boton5_GotFocus()
Call CambioColorBotonOriginal
End Sub

Private Sub b_boton6_Click()
NombreBoton = b_boton6.Name

If TipoServicio = "lavado" Then
     CodProOrden = 7
End If
If TipoServicio = "secado" Then
    CodProOrden = 72
End If

If TipoServicio = "lavadoycentrifugado" Then
    CodProOrden = 43
End If

Call CambioColorBoton
Call HabilitarTexbox
t_precio.Text = DevolverPrecio(CodProOrden)
End Sub

Private Sub b_boton6_GotFocus()
Call CambioColorBotonOriginal
End Sub

Private Sub b_boton7_Click()
NombreBoton = b_boton7.Name

If TipoServicio = "lavado" Then
     CodProOrden = 8
End If
If TipoServicio = "secado" Then
    CodProOrden = 73
End If

If TipoServicio = "lavadoycentrifugado" Then
    CodProOrden = 44
End If

Call CambioColorBoton
Call HabilitarTexbox
t_precio.Text = DevolverPrecio(CodProOrden)
End Sub

Private Sub b_boton7_GotFocus()
Call CambioColorBotonOriginal
End Sub

Private Sub b_boton8_Click()
NombreBoton = b_boton8.Name

If TipoServicio = "lavado" Then
     CodProOrden = 9
End If
If TipoServicio = "secado" Then
    CodProOrden = 74
End If

If TipoServicio = "lavadoycentrifugado" Then
    CodProOrden = 45
End If

Call CambioColorBoton
Call HabilitarTexbox
t_precio.Text = DevolverPrecio(CodProOrden)
End Sub

Private Sub b_boton8_GotFocus()
Call CambioColorBotonOriginal
End Sub

Private Sub b_boton9_Click()
NombreBoton = b_boton9.Name

If TipoServicio = "lavado" Then
     CodProOrden = 10
End If
If TipoServicio = "secado" Then
    CodProOrden = 75
End If

If TipoServicio = "lavadoycentrifugado" Then
    CodProOrden = 46
End If

Call CambioColorBoton
Call HabilitarTexbox
t_precio.Text = DevolverPrecio(CodProOrden)
End Sub

Private Sub b_boton9_GotFocus()
Call CambioColorBotonOriginal
End Sub

Private Sub b_buscar_Click()
fr_buscarorden.Show


            
End Sub

Private Sub b_cancelar_Click()

msg = MsgBox("¿Seguro de sea borrar este pedido?", vbQuestion + vbYesNo)

If msg = vbYes Then

    fg_prod_orden.Clear
    fg_prod_orden.FormatString = "Código|Descripción                 |Precio U.|Cant.| Peso|    Precio  "
    fg_prod_orden.Rows = 1
    VarAcum = 0
    
    
    t_peso.Text = ""
    t_cantidad.Text = ""
    t_precio.Text = ""
    t_importe1.Text = ""
    t_total1.Text = ""
    t_igv1.Text = ""
    t_total.Text = ""
    t_efectivo.Text = ""
    t_vuelto.Text = ""
    t_xcentajedesc.Text = ""
    t_mtodesc.Text = ""
    
    t_nombrers.Text = ""
    t_nrodoc.Text = ""
    c_tipodoc.ListIndex = 0
    t_telefono.Text = ""
    t_direccion.Text = ""
    t_email.Text = ""
    
    t_peso.Enabled = False
    t_peso.Visible = False
    Label6.Visible = False
    t_cantidad.Enabled = False
    t_precio.Enabled = False
    b_agregar.Enabled = False
    'b_cancelar.Enabled = False
    Call CambioColorBotonOriginal
End If

End Sub

Private Sub b_cliente_Click()
FormularioActual = "registroorden"
f_listacliente.Show
End Sub

Private Sub b_orden_Click()
Dim VarFechaActual As String
Dim VarDia As String
Dim VarMes As String
Dim VarAnio As String
Dim VarAnioImpr As String
Dim FechaVenta As String
Dim FechaVentaImpr As String
Dim HoraVenta As String
Dim AnchoPapel As Integer
Dim VarRuc As String
Dim VarNombre As String
Dim VarDireccion As String
Dim VarTelefono As String
Dim VarLogo As String
Dim ArchivoDet As New ADODB.Recordset
Dim ArchivoCab As New ADODB.Recordset
Dim Temp_cod_afect_Igv As ADODB.Recordset
Dim TempTipoTributos As ADODB.Recordset
Dim temproducto As ADODB.Recordset
Dim tempimpuesto As ADODB.Recordset
Dim VarNroDocumento As Long
Dim VarCodigo As String
Dim VarAbono As Currency
Dim VarRestante As Currency
Dim VarPagado As String
Dim VarImpreso As String
Dim VarsumTotTributosItem As Currency
Dim CodImpGeneral As String
Dim CodImpIsc As String
Dim CodOtImp As String
Dim VarCodAfecIgv As String
Dim VarTasaImp As Double
Dim VarMontoBaseIgvItem As Double
Dim ValidacionCorrecta As Boolean
Dim TempDatosEmpresa As ADODB.Recordset
sumTotTributosItem = 0


'***** CONDICIONAL PARA DETERMINAR SI ES UNA ORDEN NUEVA O ACTUALIZACION DE UNA EXISTENTE ********
If ActualizarOrden = True Then

    If Val(t_efectivo.Text) < Val(t_restante.Text) Then
            
        MsgBox " Efectivo insuficiente."
            
    Else

        '*********** INICIO CONSULTA PARA ACTUALIZAR ORDEN ********************
        
        Call Conn_BDaiosoft
        Conn_Mysqldb.Execute "UPDATE cabecera_doc SET pagado = 's', abono = " & t_total.Text & ", restante = 0 WHERE nroorden = " & l_numorden.Caption & ""
        
        '*********** FIN DE CONSULTA PARA ACTUALIZAR ORDEN ********************
        
        '*********** INICIO DE CODIGO PARA IMPRIMIR LA ORDEN*********
        
        TituloDoc = "ORDEN DE SERVICIO"
        VarSerie = "0000"
        VarPagado = "s"
        VarNrodoc = l_numorden.Caption
        Call ImprimirFormato80mm(TituloDoc, FechaVentaImpr, HoraVenta, VarSerie, VarNrodoc, VarPagado, VarAbono, VarRestante)
             
        '**************** FIN DE CODIGO DE IMPRESION DE ORDEN********
        
            '******* INICIO DE CODIGO PARA LIMPIAR FORMULARIO *****************
            
            fg_prod_orden.Clear
            fg_prod_orden.FormatString = "Código|Descripción                 |Precio U.|Cant.| Peso|    Precio  "
            fg_prod_orden.Rows = 1
            VarAcum = 0
            
            c_tipodoc.ListIndex = 0
            t_nombrers.Text = ""
            t_telefono.Text = ""
            t_email.Text = ""
            t_nrodoc.Text = ""
            t_direccion.Text = ""
            t_peso.Text = ""
            t_cantidad.Text = ""
            t_precio.Text = ""
            t_importe1.Text = ""
            t_total1.Text = ""
            t_igv1.Text = ""
            t_total.Text = ""
            t_efectivo.Text = ""
            t_vuelto.Text = ""
            t_xcentajedesc.Text = ""
            t_mtodesc.Text = ""
            o_prepagado.Value = True
            Frame1.Enabled = True
            Frame2.Enabled = True
            l_abono.Visible = False
            t_abono.Visible = False
            l_restante.Visible = False
            t_restante.Visible = False
            Frame4.Left = 360
            Frame4.Top = 7320
            Frame4.Visible = True
            l_total.Top = 360
            l_efectivo.Top = 960
            l_vuelto.Top = 1560
            l_total.Left = 3240
            l_efectivo.Left = 2760
            l_vuelto.Left = 3000
            t_total.Left = 4320
            t_efectivo.Left = 4320
            t_vuelto.Left = 4320
            t_total.Top = 360
            t_efectivo.Top = 960
            t_vuelto.Top = 1560
            l_xcentajedesc.Visible = True
            t_xcentajedesc.Visible = True
            l_mtodesc.Visible = True
            t_mtodesc.Visible = True
            Timer1.Enabled = False
            Timer2.Enabled = False
            l_aviso.Visible = False
            l_hora2.Caption = ""
            l_fecha.Caption = Mid(Now, 1, 10)
            ActualizarOrden = False
            b_cliente.SetFocus
            '******* FIN DE CODIGO PARA LIMPIAR FORMULARIO *****************
            
            
             '**** INICIO CODIGO PARA CARGAR EL NUMERO DE ORDEN SIGUIENTE****
            Call Conn_BDaiosoft
            Set TempCabDoc = New ADODB.Recordset
            TempCabDoc.Open "SELECT nroorden FROM cabecera_doc  ORDER BY nroorden", Conn_Mysqldb, adOpenDynamic, adLockBatchOptimistic
            If TempCabDoc.BOF = False And TempCabDoc.EOF = False Then
            TempCabDoc.MoveLast
            numorden = TempCabDoc.Fields(0).Value + 1
            l_numorden.Caption = numorden
            Else
    
                l_numorden.Caption = 1
            End If
            '**** FIN CODIGO PARA CARGAR EL NUMERO DE ORDEN SIGUIENTE****
    
    End If

Else

If t_nombrers.Text = "" Or t_telefono.Text = "" Then
    MsgBox "debe completar los datos del cliente para procesar la orden"
Else
    If t_total.Text = "" Or t_total.Text = "0" Then
        MsgBox "debe agregar al menos un producto para procesar la orden"
     Else
        'If op_efectivo.Value = False Or op_cheque.Value = False Or op_transferencia.Value = False Or op_puntodeventa.
        'MsgBox "Debe seleccionar una forma de pago"
        'Else
        msg = MsgBox("¿Esta de acuerdo con los cambios?", vbQuestion + vbYesNo)
        If msg = vbYes Then
        
       '*****VALIDACION DE LAS FORMAS DE PAGO, PREPAGADO, ABONO, PAGO AL RECOGER
        If o_prepagado.Value = True Then
            If Val(t_efectivo.Text) < Val(t_total.Text) Then
                ValidacionCorrecta = False
                MsgBox " Efectivo insuficiente."
            Else
                ValidacionCorrecta = True
                VarPagado = "s"
                VarAbono = Val(t_total.Text)
                VarRestante = 0
            End If
        End If
        
        If o_abono.Value = True Then
            If t_efectivo.Text = "" Or Val(t_efectivo.Text) = 0 Then
                ValidacionCorrecta = False
                MsgBox "Ingrese el pago parcial."
            Else
                ValidacionCorrecta = True
                VarAbono = Val(t_efectivo.Text)
                VarRestante = Val(t_total.Text) - Val(t_efectivo.Text)
                VarPagado = "n"
            End If
        End If
        
        If o_alrecojo.Value = True Then
            ValidacionCorrecta = True
            VarPagado = "n"
            VarAbono = 0
            VarRestante = Val(t_total.Text)
        End If
                                                            
       '***** FIN DE CODIGO PARA VALIDACION DE LAS FORMAS DE PAGO, PREPAGADO, ABONO, PAGO AL RECOGER
       
       
        If ValidacionCorrecta = True Then
        
            '##### AQUI INICIA EL CODIGO PARA ALMACENAR EL DETALLE DEL DOCUMENTO EN LA TABLA det_doc #####
            For x = 1 To fg_prod_orden.Rows - 1
    
                'consulta para recuperar los codigos de tipos de impuesto IGV, ISC, otros impuestos
                Call Conn_BDaiosoft
                Set temproducto = New ADODB.Recordset
                temproducto.Open "SELECT * FROM producto where codigo = '" & fg_prod_orden.TextMatrix(x, 0) & "'", Conn_Mysqldb, adOpenDynamic, adLockBatchOptimistic
                If temproducto.BOF = False And temproducto.EOF = False Then
                    temproducto.MoveFirst
                    CodImpGeneral = temproducto.Fields(6).Value
                    CodImpIsc = temproducto.Fields(7).Value
                    CodOtImp = temproducto.Fields(8).Value
                    CodIcbper = temproducto.Fields(9).Value
                End If
    
        
                If CodImpIsc = "01" Then
            
                    VarCodImpIsc = "-"
                    VarnomTributoIscItem = ""
                    Varmontoiscitem = 0
                    Varmontobaseiscitem = 0
                    Varcodtiptributoiscitem = ""
                    Vartipsisisc = ""
                    Varporcentajeiscitem = 0
                Else
            
                End If
        
        
                If CodOtImp = "01" Then
        
                    VarcodTriOtroItem = "-"
                    VarmtoTriOtroItem = 0
                    VarmtoBaseTriOtroItem = 0
                    VarnomTributoIOtroItem = ""
                    VarcodTipTributoIOtroItem = ""
                    VarporTriOtroItem = "0"
                Else
        
        
                End If
        
        
                '****seccion de codigo para realizar los calculos del impueso ICBPER por item****
                If CodIcbper = "01" Then
        
                    Varcodtriicbper = "-"
                    Varmtotriicbperitem = 0
                    Varcantboltriicbperitem = 0
                    Varnomtriicbperitem = ""
                    Varcodtiptributoicbperitem = ""
                    Varmtotributoicbperunidad = 0
                Else
            
                    'Varcodtriicbper = CodIcbper
                    'consulta para extraer nombre de impuesto  y tasa
                    'Set tempimpuesto = New ADODB.Recordset
                    'tempimpuesto.Open "SELECT denominacion, tasa FROM impuesto WHERE codigo = '" & CodIcbper & "'", conn_mysqldb, adOpenDynamic, adLockBatchOptimistic
                    'If tempimpuesto.BOF = False And tempimpuesto.EOF = False Then
                    '    tempimpuesto.MoveFirst
                    '    Varnomtriicbperitem = tempimpuesto.Fields(0).Value
                    '    Varmtotributoicbperunidad = tempimpuesto.Fields(1).Value
                    'End If
                    'Varcantboltriicbperitem = Val(fg_detallefactura.TextMatrix(contador, 2))
                    ' Varmtotriicbperitem = Val(Varmtotributoicbperunidad) * Val(fg_detallefactura.TextMatrix(contador, 2))
                    
                    'consulta para extraer de la tabla tipo_tributos el codigo internacional del impuesto
                    'Set TempTipoTributos = New ADODB.Recordset
                    'TempTipoTributos.Open "SELECT codinternacional FROM tipo_tributos WHERE codigo = '" & CodIcbper & "'", conn_mysqldb, adOpenDynamic, adLockBatchOptimistic
                    'If TempTipoTributos.BOF = False And TempTipoTributos.EOF = False Then
                    '    TempTipoTributos.MoveFirst
                    '    Varcodtiptributoicbperitem = TempTipoTributos.Fields(0).Value
                    'End If
        
                End If
                '****fin de seccion de codigo para realizar los calculos del impueso ICBPER por item****
        
        
                'consulta para recuperar el codigo de tipo de afectacion del IGV
                'Call conn_BDaiosoft
                Set Temp_cod_afect_Igv = New ADODB.Recordset
                 Temp_cod_afect_Igv.Open "SELECT codigo FROM cod_afectacion_igv WHERE codigo_tributo = '" & CodImpGeneral & "'", Conn_Mysqldb, adOpenDynamic, adLockBatchOptimistic
                If Temp_cod_afect_Igv.BOF = False And Temp_cod_afect_Igv.EOF = False Then
                    Temp_cod_afect_Igv.MoveFirst
                    VarCodAfecIgv = Temp_cod_afect_Igv.Fields(0).Value
                End If
       
                'consulta para recuper la tasa del impuesto correspondiente y nombre de impuesto (IGV,EXP,INA,GRA,IVAP)
                Set tempimpuesto = New ADODB.Recordset
                tempimpuesto.Open "SELECT denominacion, tasa FROM impuesto WHERE codigo = '" & CodImpGeneral & "'", Conn_Mysqldb, adOpenDynamic, adLockBatchOptimistic
                If tempimpuesto.BOF = False And tempimpuesto.EOF = False Then
                    tempimpuesto.MoveFirst
                    VarDenominacion = tempimpuesto.Fields(0).Value
                    VarTasaImp = tempimpuesto.Fields(1).Value
                End If
       
                'consulta para extraer de la tabla tipo_tributos el codigo internacional del impuesto
                Set TempTipoTributos = New ADODB.Recordset
                TempTipoTributos.Open "SELECT codinternacional FROM tipo_tributos WHERE codigo = '" & CodImpGeneral & "'", Conn_Mysqldb, adOpenDynamic, adLockBatchOptimistic
                If TempTipoTributos.BOF = False And TempTipoTributos.EOF = False Then
                    TempTipoTributos.MoveFirst
                    VarCodInternacional = TempTipoTributos.Fields(0).Value
                End If
       
       
       
                'calculo de la sumatoria de todos los tributos por item
                VarsumTotTributosItem = Val(fg_prod_orden.TextMatrix(x, 5)) * 0.18
                VarMontoBaseIgvItem = Val(fg_prod_orden.TextMatrix(x, 5)) * 100 / 118
                
                If fg_prod_orden.TextMatrix(x, 4) = "" Then
                    VarPeso = 0
                Else
                    VarPeso = fg_prod_orden.TextMatrix(x, 4)
                End If
                
                'consulta para insertar el detalle del documento de venta en la tabla det_doc
                'Call conn_BDaiosoft

                Conn_Mysqldb.Execute "INSERT INTO det_doc SET serie = '0000'," _
                & "nroorden = " & l_numorden.Caption & ", nrodoc = " & l_numorden.Caption & ", codproducto = '" & fg_prod_orden.TextMatrix(x, 0) & "'," _
                & "codunidadmedida = 'NIU', cantidaditem = " & fg_prod_orden.TextMatrix(x, 3) & "," _
                & "peso = " & VarPeso & "," _
                & "descripitem = '" & fg_prod_orden.TextMatrix(x, 1) & "', preciounititem = " & fg_prod_orden.TextMatrix(x, 2) & "," _
                & "sumtottributositem = " & VarsumTotTributosItem & ", codtriigv = '" & CodImpGeneral & "'," _
                & "montoigvitem = " & VarsumTotTributosItem & ", montobaseigvitem = " & VarMontoBaseIgvItem & "," _
                & "nombretributoigvxitem = '" & VarDenominacion & "', codtiptribigvitem = '" & VarCodInternacional & "'," _
                & "tipafeigv = '" & VarCodAfecIgv & "', porcentajeigv = " & VarTasaImp & "," _
                & "codtriisc = '" & VarCodImpIsc & "', montoiscitem = " & Varmontoiscitem & "," _
                & "montobaseiscitem = " & Varmontobaseiscitem & ", nombretributoiscxitem = '" & VarnomTributoIscItem & "'," _
                & "codtiptributoiscitem = '" & Varcodtiptributoiscitem & "', tipsisisc = '" & Vartipsisisc & "'," _
                & "porcentajeiscitem = " & Varporcentajeiscitem & ", codotrostributos = '" & VarcodTriOtroItem & "'," _
                & "montootrotriitem = " & VarmtoTriOtroItem & ", montobaseotrotriitem = " & VarmtoBaseTriOtroItem & "," _
                & "nombreotrotributoxitem = '" & VarnomTributoIOtroItem & "', codtipotrotributositem = '" & VarcodTipTributoIOtroItem & "'," _
                & "porcentajeotrotribitem = " & VarporTriOtroItem & "," _
                & "codtriicbper = '" & Varcodtriicbper & "'," _
                & "mtotriicbperitem = " & Varmtotriicbperitem & "," _
                & "cantboltriicbperitem = " & Varcantboltriicbperitem & "," _
                & "nomtriicbperitem = '" & Varnomtriicbperitem & "'," _
                & "codtiptributoicbperitem = '" & Varcodtiptributoicbperitem & "'," _
                & "mtotributoicbperunidad = " & Varmtotributoicbperunidad & "," _
                & "montototalxitem = " & VarMontoBaseIgvItem & "," _
                & "montototalxitemconimpuestos = " & fg_prod_orden.TextMatrix(x, 5) & ""
                
                
            Next x
            '##### AQUI FINALIZA EL CODIGO PARA ALMACENAR EL DETALLE DEL DOCUMENTO EN LA TABLA det_doc #####
    
    
            '##### AQUI INICIA EL CODIGO PARA ALMACENAR LOS DATOS DEL DOCUMENTO EN LA TABLA cabecera_doc #####
       
            VarFechaActual = Now
            VarDia = Mid(VarFechaActual, 1, 2)
            VarMes = Mid(VarFechaActual, 4, 2)
            VarAnio = Mid(VarFechaActual, 7, 4)
            VarAnioImpr = Mid(VarFechaActual, 9, 2)
    
            FechaVenta = VarAnio + "-" + VarMes + "-" + VarDia
            FechaVentaImpr = VarDia + "/" + VarMes + "/" + VarAnioImpr
            mihoracompleta = Now
            HoraVenta = Mid(mihoracompleta, 12, 8)
            
            
            If t_xcentajedesc.Text = "" Then
                VarXcentDesc = 0
            Else
                VarXcentDesc = Val(t_xcentajedesc.Text)
            End If
    
            'consulta para la insercion de datos en la tabla cabecera_doc
            Conn_Mysqldb.Execute "INSERT INTO cabecera_doc SET serie= '0000'," _
            & "nroorden = " & l_numorden.Caption & "," _
            & "nrodoc = " & l_numorden.Caption & "," _
            & "tipodoc = '00'," _
            & "tipooper = '0101'," _
            & "firmadig = '-'," _
            & "fechaem = '" & FechaVenta & "'," _
            & "horaem = '" & HoraVenta & "'," _
            & "fechavenci = '" & FechaVenta & "'," _
            & "codlocalemisor = '0000'," _
            & "tipdocusuario = '" & c_tipodoc.Text & "'," _
            & "numdocusuario = '" & t_nrodoc.Text & "'," _
            & "nombrers = '" & t_nombrers.Text & "'," _
            & "tipmoneda = 'PEN'," _
            & "sumtottributos = " & t_igv1.Text & "," _
            & "sumtotvalventa = " & t_importe1.Text & "," _
            & "sumprecioventa = " & t_total1.Text & "," _
            & "sumdesctotal = " & VarXcentDesc & "," _
            & "sumotroscargos = 0," _
            & "sumtotalanticipos = 0," _
            & "sumimpventa = " & t_total1.Text & "," _
            & "ublversionld = '2.1'," _
            & "customizationld = '2.0'," _
            & "pagado = '" & VarPagado & "', abono = " & VarAbono & ", restante = " & VarRestante & "," _
            & "impreso = 's'"
    
    
            '##### AQUI FINALIZA EL CODIGO PARA ALMACENAR LOS DATOS DEL DOCUMENTO EN LA TABLA cabecera_doc #####
            
            
            '############ INICIO DE CODIGO PARA IMPRIMIR LA ORDEN ##############
                
            TituloDoc = "ORDEN DE SERVICIO"
            VarSerie = "0000"
            VarNrodoc = l_numorden.Caption
            Call ImprimirFormato80mm(TituloDoc, FechaVentaImpr, HoraVenta, VarSerie, VarNrodoc, VarPagado, VarAbono, VarRestante)
           
            
            '############ FIN DE CODIGO DE IMPRESION DE ORDEN ############
            
            
             '**** CODIGO PARA CARGAR EL NUMERO DE ORDEN SIGUIENTE****
            Call Conn_BDaiosoft
            Set TempCabDoc = New ADODB.Recordset
            TempCabDoc.Open "SELECT nroorden FROM cabecera_doc  ORDER BY nroorden", Conn_Mysqldb, adOpenDynamic, adLockBatchOptimistic
            If TempCabDoc.BOF = False And TempCabDoc.EOF = False Then
            TempCabDoc.MoveLast
            numorden = TempCabDoc.Fields(0).Value + 1
            l_numorden.Caption = numorden
            Else
    
                l_numorden.Caption = 1
            End If
            '**** FIN CODIGO PARA CARGAR EL NUMERO DE ORDEN SIGUIENTE****
            
            
             '******* INICIO DE CODIGO PARA LIMPIAR FORMULARIO *****************
            
            fg_prod_orden.Clear
            fg_prod_orden.FormatString = "Código|Descripción                 |Precio U.|Cant.| Peso|    Precio  "
            fg_prod_orden.Rows = 1
            VarAcum = 0
            
            c_tipodoc.ListIndex = 0
            t_nombrers.Text = ""
            t_telefono.Text = ""
            t_email.Text = ""
            t_nrodoc.Text = ""
            t_direccion.Text = ""
            t_peso.Text = ""
            t_cantidad.Text = ""
            t_precio.Text = ""
            t_importe1.Text = ""
            t_total1.Text = ""
            t_igv1.Text = ""
            t_total.Text = ""
            t_efectivo.Text = ""
            t_vuelto.Text = ""
            t_xcentajedesc.Text = ""
            t_mtodesc.Text = ""
            o_prepagado.Value = True
            t_peso.Enabled = False
            t_peso.Visible = False
            Label6.Visible = False
            t_cantidad.Enabled = False
            t_precio.Enabled = False
            b_agregar.Enabled = False
            'b_cancelar.Enabled = False
            Call CambioColorBotonOriginal
            b_cliente.SetFocus
           '******* FIN DE CODIGO PARA LIMPIAR FORMULARIO *****************
            
            
            
        End If
        End If
    
    End If
End If
End If

End Sub

Private Sub b_otros_Click()
'f_listaproductoventa.Show
End Sub

Private Sub b_prendaxpeso_GotFocus()
Call CambioColorBotonOriginal
End Sub

Private Sub b_reimprimir_Click()
'*********** INICIO DE CODIGO PARA IMPRIMIR LA ORDEN***********
Call Conn_BDaiosoft
            Set TempDatosEmpresa = New ADODB.Recordset
            TempDatosEmpresa.Open "SELECT ruc, nombre,direccionppal,telefonocelular,logo FROM datos_empresa", Conn_Mysqldb, adOpenDynamic, adLockBatchOptimistic
            If TempDatosEmpresa.BOF = False And TempDatosEmpresa.EOF = False Then
                TempDatosEmpresa.MoveFirst
               VarRuc = TempDatosEmpresa.Fields(0).Value
                VarNombre = TempDatosEmpresa.Fields(1).Value
                VarDireccion = TempDatosEmpresa.Fields(2).Value
                VarTelefono = TempDatosEmpresa.Fields(3).Value
                VarLogo = TempDatosEmpresa.Fields(4).Value
            Else
                VarRuc = " RUC DEMO"
                VarNombre = "NOMBRE DEMO"
                VarDireccion = "DIRECCION DEMO"
                VarTelefono = "TELEFONO DEMO"
                VarLogo = ""
            End If
            ' 567 twips equivalen a un centimetro
            AnchoPapel = 4536
            Printer.CurrentY = 567
            Printer.Font = "Courier New"
            Printer.FontSize = 8
            Printer.CurrentX = AnchoPapel / 2 - Printer.TextWidth(VarNombre) / 2
            Printer.Print VarNombre
            Printer.CurrentX = AnchoPapel / 2 - Printer.TextWidth("RUC: " & VarRuc) / 2
            Printer.Print "RUC: "; VarRuc
            Printer.CurrentX = AnchoPapel / 2 - Printer.TextWidth(VarDireccion) / 2
            Printer.Print VarDireccion
            Printer.CurrentX = AnchoPapel / 2 - Printer.TextWidth("TLF: " & VarTelefono) / 2
            Printer.Print "TLF: "; VarTelefono
            Printer.Print
            Printer.CurrentX = AnchoPapel / 2 - Printer.TextWidth("REIMPR. ORD. SERV.") / 2
            Printer.Print "REIMPR. ORD. SERV."
            Printer.Print
            Printer.Print
            Printer.CurrentX = 284
            'inicio codigo para ordenar el nro orden de derecha a izquierda
            largonumorden = Len(l_numorden.Caption)
            cantespacio = 9 - largonumorden
            'fin codigo para ordenar el nro orden de derecha a izquierda
            Printer.Print "FECHA: "; FechaVentaImpr; Spc(10); "NRO ORD: "; Spc(cantespacio); l_numorden.Caption
            Printer.CurrentX = 284
            Printer.Print "HORA: "; HoraVenta
            Printer.CurrentX = 284
            Printer.Print "CLIENTE: "; t_nombrers.Text
            Printer.CurrentX = 284
            Printer.Print "___________________________________________"
            Printer.CurrentX = 284
            Printer.Print "DESCRIPCION"; Spc(8); "P.U"; Spc(5); "KG"; Spc(1); "CNT"; Spc(2); "TOTAL S/"
            Printer.CurrentX = 284
            Printer.Print "-------------------------------------------"
            For x = 1 To fg_prod_orden.Rows - 1
                Printer.CurrentX = 284
                Printer.Print Mid(fg_prod_orden.TextMatrix(x, 1), 1, 43)
                
                anchopu = Len(fg_prod_orden.TextMatrix(x, 2))
                ubipu = 284 + (22 - anchopu) * 96
                Printer.CurrentX = ubipu
                Printer.Print fg_prod_orden.TextMatrix(x, 2)
                
                anchopeso = Len(fg_prod_orden.TextMatrix(x, 4))
                ubipeso = 284 + (29 - anchopeso) * 96
                Printer.CurrentY = Printer.CurrentY - 170
                Printer.CurrentX = ubipeso
                Printer.Print fg_prod_orden.TextMatrix(x, 4)
                
                anchocant = Len(fg_prod_orden.TextMatrix(x, 3))
                ubicant = 284 + (33 - anchocant) * 96
                Printer.CurrentY = Printer.CurrentY - 170
                Printer.CurrentX = ubicant
                Printer.Print fg_prod_orden.TextMatrix(x, 3)
                
                anchototal = Len(fg_prod_orden.TextMatrix(x, 5))
                ubitotal = 284 + (43 - anchototal) * 96
                Printer.CurrentY = Printer.CurrentY - 170
                Printer.CurrentX = ubitotal
                Printer.Print fg_prod_orden.TextMatrix(x, 5)
                       
            Next x
            Printer.CurrentX = 284
            Printer.Print "-------------------------------------------"
            Printer.CurrentX = 284 + 96 * (43 - Len(t_total.Text) - Len("TOTAL A PAGAR S/: "))
            Printer.Print "TOTAL A PAGAR S/: "
            Printer.CurrentX = 284 + 96 * (43 - Len(t_total.Text))
            Printer.CurrentY = Printer.CurrentY - 170
            Printer.Print t_total.Text
            'If VarPagado = "s" Then
                Printer.CurrentX = 284
                Printer.Print "PAGADO"
            'End If
            'If VarPagado = "n" Then
            '    Printer.CurrentX = 284
            '    Printer.Print "ABONADO S/"; Spc(1); VarAbono
            '    Printer.CurrentX = 284
            '    Printer.Print "RESTANTE S/"; Spc(1); VarRestante
            '    Printer.CurrentX = 284
            '    Printer.Print "POR PAGAR"
            'End If
            
            
            Printer.EndDoc
           
            
'**************** FIN DE CODIGO DE IMPRESION DE ORDEN***************
End Sub

Private Sub b_sig_Click()
fr_prendasvestir.Visible = False
fr_prendas.Visible = True
fr_prendas.Top = 2160
fr_prendas.Left = 360
b_atras2.Visible = True

t_peso.Text = ""
t_cantidad.Text = ""
t_precio.Text = ""
t_peso.Enabled = False
t_peso.Visible = False
Label6.Visible = False
t_cantidad.Enabled = False
t_precio.Enabled = False
b_agregar.Enabled = False
b_cancelar.Enabled = False
Call CambioColorBotonOriginal

End Sub

Private Sub b_sig2_Click()
fr_prendasvestir.Visible = True
fr_prendasvestir.Top = 2160
fr_prendasvestir.Left = 360
b_atras.Visible = True

t_peso.Text = ""
t_cantidad.Text = ""
t_precio.Text = ""
t_peso.Enabled = False
t_peso.Visible = False
Label6.Visible = False
t_cantidad.Enabled = False
t_precio.Enabled = False
b_agregar.Enabled = False
b_cancelar.Enabled = False
Call CambioColorBotonOriginal

End Sub



Private Sub b_volver_Click()
  '******* INICIO DE CODIGO PARA LIMPIAR FORMULARIO *****************
            
            fg_prod_orden.Clear
            fg_prod_orden.FormatString = "Código|Descripción                 |Precio U.|Cant.| Peso|    Precio  "
            fg_prod_orden.Rows = 1
            VarAcum = 0
            
            c_tipodoc.ListIndex = 0
            t_nombrers.Text = ""
            t_telefono.Text = ""
            t_email.Text = ""
            t_nrodoc.Text = ""
            t_direccion.Text = ""
            t_peso.Text = ""
            t_cantidad.Text = ""
            t_precio.Text = ""
            t_importe1.Text = ""
            t_total1.Text = ""
            t_igv1.Text = ""
            t_total.Text = ""
            t_efectivo.Text = ""
            t_vuelto.Text = ""
            t_xcentajedesc.Text = ""
            t_mtodesc.Text = ""
            o_prepagado.Value = True
            Frame1.Enabled = True
            Frame2.Enabled = True
            l_abono.Visible = False
            t_abono.Visible = False
            l_restante.Visible = False
            t_restante.Visible = False
            Frame4.Left = 360
            Frame4.Top = 7320
            Frame4.Visible = True
            Frame4.Enabled = True
            l_total.Top = 360
            l_efectivo.Top = 960
            l_vuelto.Top = 1560
            l_total.Left = 3240
            l_efectivo.Left = 2760
            l_vuelto.Left = 3000
            t_total.Left = 4320
            t_efectivo.Left = 4320
            t_vuelto.Left = 4320
            t_total.Top = 360
            t_efectivo.Top = 960
            t_vuelto.Top = 1560
            l_xcentajedesc.Visible = True
            t_xcentajedesc.Visible = True
            l_mtodesc.Visible = True
            t_mtodesc.Visible = True
            Timer1.Enabled = False
            Timer2.Enabled = False
            l_aviso.Visible = False
            l_hora2.Caption = ""
            l_fecha.Caption = Mid(Now, 1, 10)
            ActualizarOrden = False
            b_reimprimir.Visible = False
            b_anular.Visible = False
            b_volver.Visible = False
            b_orden.Visible = True
            b_boleta.Visible = True
            b_factura.Visible = True
            t_efectivo.Locked = False
            b_buscar.Visible = True
            b_cliente.SetFocus
            '******* FIN DE CODIGO PARA LIMPIAR FORMULARIO *****************
            
            
             '**** INICIO CODIGO PARA CARGAR EL NUMERO DE ORDEN SIGUIENTE****
            Call Conn_BDaiosoft
            Set TempCabDoc = New ADODB.Recordset
            TempCabDoc.Open "SELECT nroorden FROM cabecera_doc  ORDER BY nroorden", Conn_Mysqldb, adOpenDynamic, adLockBatchOptimistic
            If TempCabDoc.BOF = False And TempCabDoc.EOF = False Then
            TempCabDoc.MoveLast
            numorden = TempCabDoc.Fields(0).Value + 1
            l_numorden.Caption = numorden
            Else
    
                l_numorden.Caption = 1
            End If
            '**** FIN CODIGO PARA CARGAR EL NUMERO DE ORDEN SIGUIENTE****
    
End Sub

Private Sub Command10_Click()

TipoServicio = "lavadoycentrifugado"

fr_prendas.Visible = True
fr_prendas.Top = 2160
fr_prendas.Left = 360


b_boton27.Enabled = False
b_boton28.Enabled = False
b_boton29.Enabled = False
b_boton30.Enabled = False
b_boton31.Enabled = False
b_boton32.Enabled = False

b_boton24.Enabled = True
b_boton35.Enabled = False

End Sub

Private Sub Command11_Click()
fr_prendasvestir.Visible = True
fr_prendasvestir.Top = 2160
fr_prendasvestir.Left = 360

b_atras3.Visible = True
b_atras3.Left = 5280
b_atras3.Top = 4080

b_sig.Visible = False
b_atras.Visible = False
TipoServicio = "lavadoseco"


b_boton27.Enabled = True
b_boton28.Enabled = True
b_boton29.Enabled = True
b_boton30.Enabled = True
b_boton31.Enabled = True

b_boton24.Enabled = False
b_boton35.Enabled = False



End Sub

Private Sub Command12_Click()
fr_prendasvestir.Visible = True
fr_prendasvestir.Top = 2160
fr_prendasvestir.Left = 360

b_atras3.Visible = True
b_atras3.Left = 5280
b_atras3.Top = 4080

b_sig.Visible = False
b_atras.Visible = False
TipoServicio = "teñido"

b_boton35.Enabled = False
b_boton27.Enabled = False
b_boton28.Enabled = False
b_boton29.Enabled = False
b_boton30.Enabled = False
b_boton31.Enabled = False

End Sub

Private Sub Command13_Click()
fr_prendas.Visible = True
fr_prendas.Top = 2160
fr_prendas.Left = 360
TipoServicio = "secado"

b_boton27.Enabled = False
b_boton28.Enabled = False
b_boton29.Enabled = False
b_boton30.Enabled = False
b_boton31.Enabled = False
b_boton32.Enabled = False

b_boton24.Enabled = True
b_boton35.Enabled = False

End Sub

Private Sub Command14_Click()
fr_prendas.Visible = True
fr_prendas.Top = 2160
fr_prendas.Left = 360
'b_atras3.Visible = True

b_boton27.Enabled = False
b_boton28.Enabled = False
b_boton29.Enabled = False
b_boton30.Enabled = False
b_boton31.Enabled = False

b_boton24.Enabled = True
b_boton35.Enabled = True
b_boton32.Enabled = True
b_boton1.SetFocus

TipoServicio = "lavado"

End Sub

Private Sub b_prendaxpeso_Click()
    
End Sub

Private Sub Command22_Click()
fr_prendas.Visible = False
fr_prendas.Top = 2640
fr_prendas.Left = 480
End Sub

Private Sub Command30_Click()
fr_prendasvestir.Visible = False
End Sub

Private Sub Commandb_boton23_Click()

NombreBoton = b_boton23

If TipoServicio = "lavadoseco" Then
    CodProOrden = 109
End If

If TipoServicio = "lavadoycentrifugado" Then
    CodProOrden = 37
End If

If TipoServicio = "lavado" Then
     CodProOrden = 94
End If

If TipoServicio = "secado" Then
    CodProOrden = 66
End If

If TipoServicio = "teñido" Then
    CodProOrden = 110
End If

Call CambioColorBoton
Call HabilitarTexbox
t_precio.Text = DevolverPrecio(CodProOrden)

End Sub


Private Sub fg_prod_orden_Click()
If fg_prod_orden.Rows > 1 Then

    ' se habilita la edicion solo para las columnas 3 y 4 que admiten modificacion
    If fg_prod_orden.Col = 3 Or fg_prod_orden.Col = 4 And fg_prod_orden.TextMatrix(fg_prod_orden.Row, 4) <> "" Then
        'se procede a dar al t_editar  todos los valores de propiedad de la
        'de la celda que tiene el enfoque
        t_editar.Left = fg_prod_orden.CellLeft + fg_prod_orden.Left
        t_editar.Top = fg_prod_orden.CellTop + fg_prod_orden.Top
        t_editar.Width = fg_prod_orden.CellWidth
        t_editar.Height = fg_prod_orden.CellHeight
        t_editar.BorderStyle = 0
        t_editar.FontName = fg_prod_orden.CellFontName
        t_editar.FontSize = fg_prod_orden.CellFontSize
        t_editar.FontBold = True
        t_editar.Visible = True ' se coloca visible el t_editar
        t_editar.SetFocus ' t_editar recibe el enfoque
        t_editar.Text = fg_prod_orden.TextMatrix(fg_prod_orden.Row, fg_prod_orden.Col)
        'se pasa al t_editar, el valor de la celda que tiene el enfoque
    End If
End If
End Sub

Private Sub fg_prod_orden_DblClick()
msg = MsgBox(" Desea eliminar el registro actual?", vbQuestion + vbYesNo)
If msg = vbYes Then

    If fg_prod_orden.Rows > 2 Then
        fg_prod_orden.RemoveItem (fg_prod_orden.Row)
    Else
        fg_prod_orden.Clear
        fg_prod_orden.FormatString = "Código|Descripción                 |Precio U.|Cant.| Peso|    Precio  "
        fg_prod_orden.Rows = 1
        VarAcum = 0
        t_total1.Text = ""
        t_igv1.Text = ""
        t_importe1.Text = ""
        t_total.Text = ""
        t_efectivo.Text = ""
        t_vuelto.Text = ""
        t_xcentajedesc.Text = ""
        t_mtodesc.Text = ""
        
    End If

    If fg_prod_orden.Rows >= 2 Then
        For x = 1 To fg_prod_orden.Rows - 1
            VarAcum = VarAcum + Val(fg_prod_orden.TextMatrix(x, 5))
            t_total1.Text = VarAcum
        Next x
        t_total1.Text = VarAcum
        t_total.Text = VarAcum
        VarIgv = Round(VarAcum * 0.18, 2)
        VarImporte = Round(VarAcum - VarIgv, 2)
        t_igv1.Text = VarIgv
        t_importe1.Text = VarImporte
    End If

End If

End Sub

Private Sub Form_Load()
l_fecha.Caption = Mid(Now, 1, 10)
b_agregar.Enabled = False

l_abono.Visible = False
l_restante.Visible = False
t_abono.Visible = False
t_restante.Visible = False

ActualizarOrden = False


'****codigo para cargar el numero de orden siguiente****
Call Conn_BDaiosoft
    Set TempCabDoc = New ADODB.Recordset
    TempCabDoc.Open "SELECT nroorden FROM cabecera_doc  ORDER BY nroorden", Conn_Mysqldb, adOpenDynamic, adLockBatchOptimistic
    If TempCabDoc.BOF = False And TempCabDoc.EOF = False Then
        TempCabDoc.MoveLast
        numorden = TempCabDoc.Fields(0).Value + 1
        l_numorden.Caption = numorden
    Else
    
        l_numorden.Caption = 1
    End If
'****fin de codigo para cargar el numero de orden siguiente****

'codigo para cargar el combobox tipo de documento de identidad del cliente
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

Private Sub o_abono_Click()
t_efectivo.Enabled = True
End Sub

Private Sub o_abono_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    t_efectivo.Enabled = True
    t_efectivo.SetFocus
End If
End Sub

Private Sub o_alrecojo_Click()
t_efectivo.Enabled = False
t_efectivo.Text = ""
t_vuelto.Text = ""

End Sub

Private Sub o_alrecojo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    t_efectivo.Locked = True
    t_efectivo.Text = ""
    t_vuelto.Text = ""
    b_orden.SetFocus
End If
End Sub

Private Sub o_pagoexacto_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    t_efectivo.Text = t_total.Text
    t_efectivo.Locked = True
    b_orden.SetFocus
End If
End Sub

Private Sub o_prepagado_Click()
t_efectivo.Enabled = True
t_efectivo.SetFocus
End Sub
Private Sub o_prepagado_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    t_efectivo.Enabled = True
    t_efectivo.SetFocus
End If
End Sub

Private Sub t_cantidad_KeyPress(KeyAscii As Integer)

If KeyAscii < 48 Or KeyAscii > 57 Then
    If KeyAscii <> 8 And KeyAscii <> 13 Then
        KeyAscii = 0
        Beep
    End If
End If
If KeyAscii = 13 Then
    b_agregar.Value = True
End If

End Sub

Private Sub t_editar_Change()


fg_prod_orden.TextMatrix(fg_prod_orden.Row, fg_prod_orden.Col) = t_editar.Text

If fg_prod_orden.Col = 4 And fg_prod_orden.TextMatrix(fg_prod_orden.Row, 4) <> "" Then

    If TipoServicio = "lavado" And Val(t_editar.Text) >= 3 Then
    
        fg_prod_orden.TextMatrix(fg_prod_orden.Row, 2) = 3
        fg_prod_orden.TextMatrix(fg_prod_orden.Row, 5) = Val(t_editar.Text) * Val(fg_prod_orden.TextMatrix(fg_prod_orden.Row, 2))
    
    End If
    If TipoServicio = "lavado" And Val(t_editar.Text) < 3 Then
    
        fg_prod_orden.TextMatrix(fg_prod_orden.Row, 2) = 4
        fg_prod_orden.TextMatrix(fg_prod_orden.Row, 5) = Val(t_editar.Text) * Val(fg_prod_orden.TextMatrix(fg_prod_orden.Row, 2))
    
    End If

End If

If fg_prod_orden.Col = 3 And fg_prod_orden.TextMatrix(fg_prod_orden.Row, 4) = "" Then

    fg_prod_orden.TextMatrix(fg_prod_orden.Row, 5) = Val(t_editar.Text) * Val(fg_prod_orden.TextMatrix(fg_prod_orden.Row, 2))

End If

If fg_prod_orden.Rows >= 2 Then
        For x = 1 To fg_prod_orden.Rows - 1
            VarAcum = VarAcum + Val(fg_prod_orden.TextMatrix(x, 5))
            t_total1.Text = VarAcum
        Next x
        t_total1.Text = VarAcum
        t_total.Text = VarAcum
        VarIgv = Round(VarAcum * 0.18, 2)
        VarImporte = Round(VarAcum - VarIgv, 2)
        t_igv1.Text = VarIgv
        t_importe1.Text = VarImporte
    End If


End Sub

Private Sub t_editar_KeyPress(KeyAscii As Integer)

If fg_prod_orden.Col = 4 Then

    If KeyAscii < 48 Or KeyAscii > 57 Then
        If KeyAscii <> 8 And KeyAscii <> 46 Then
            KeyAscii = 0
            Beep
        End If
    End If
    
End If


If fg_prod_orden.Col = 3 Then

    If KeyAscii < 48 Or KeyAscii > 57 Then
        If KeyAscii <> 8 Then
            KeyAscii = 0
            Beep
        End If
    End If
    
End If
    
End Sub

Private Sub t_editar_LostFocus()
t_editar.Visible = False
End Sub

Private Sub t_efectivo_Change()
If ActualizarOrden = True Then
    t_vuelto = Val(t_efectivo.Text) - Val(t_restante.Text)
Else
    t_vuelto.Text = Val(t_efectivo.Text) - Val(t_total.Text)
End If
End Sub

Private Sub t_efectivo_KeyPress(KeyAscii As Integer)
If KeyAscii < 48 Or KeyAscii > 57 Then
    If KeyAscii <> 8 And KeyAscii <> 46 And KeyAscii <> 13 Then
        KeyAscii = 0
        Beep
    End If
End If
If KeyAscii = 13 Then
    b_orden.Value = True
End If

End Sub

Private Sub t_mtodesc_Change()
If Val(t_total.Text) > 0 Then
    For x = 1 To fg_prod_orden.Rows - 1
        VarTotalOrden = VarTotalOrden + Val(fg_prod_orden.TextMatrix(x, 5))
    Next x
    
    t_total1.Text = VarTotalOrden - Val(t_mtodesc.Text)
    t_total.Text = VarTotalOrden - Val(t_mtodesc.Text)
    
    t_igv1.Text = Val(t_total.Text) * 0.18
    t_importe1.Text = Val(t_total.Text) - Val(t_igv1.Text)
    
    
    VarXcentajePrecio = Val(t_total.Text) * 100 / VarTotalOrden
    VarXcentajeDesc = 100 - VarXcentajePrecio
    
    t_xcentajedesc.Text = VarXcentajeDesc
Else
    t_mtodesc.Text = ""
End If
End Sub

Private Sub t_mtodesc_KeyPress(KeyAscii As Integer)
If KeyAscii < 48 Or KeyAscii > 57 Then
    If KeyAscii <> 8 And KeyAscii <> 46 Then
        KeyAscii = 0
        Beep
    End If
End If
End Sub

Private Sub t_peso_Change()

If Val(t_peso.Text) >= 3 And TipoServicio = "lavado" Then

    t_precio.Text = 3
Else
    If Val(t_peso.Text) < 3 And TipoServicio = "lavado" Then
        t_precio.Text = 4
    End If
End If

End Sub

Private Sub t_peso_KeyPress(KeyAscii As Integer)
If KeyAscii < 48 Or KeyAscii > 57 Then
    If KeyAscii <> 8 And KeyAscii <> 46 And KeyAscii <> 13 Then
        KeyAscii = 0
        Beep
    End If
End If
If KeyAscii = 13 Then
    t_cantidad.SetFocus
End If

End Sub

Private Sub t_xcentajedesc_Change()
If Val(t_total.Text) > 0 Then
    For x = 1 To fg_prod_orden.Rows - 1
        VarTotalOrden = VarTotalOrden + Val(fg_prod_orden.TextMatrix(x, 5))
    Next x
    
    If Val(t_xcentajedesc.Text) >= 1 Then
        VarNuevoTotal = VarTotalOrden - VarTotalOrden * (Val(t_xcentajedesc.Text) / 100)
        t_total1.Text = VarNuevoTotal
        t_total.Text = VarNuevoTotal
        t_mtodesc.Text = VarTotalOrden - VarNuevoTotal
        t_igv1.Text = VarNuevoTotal * 0.18
        t_importe1.Text = VarNuevoTotal - Val(t_igv1.Text)
    Else
        t_total1.Text = VarTotalOrden
        t_total.Text = VarTotalOrden
        t_igv1.Text = VarTotalOrden * 0.18
        t_importe1.Text = VarTotalOrden - Val(t_igv1.Text)
        t_mtodesc.Text = 0
    End If
Else
    t_xcentajedesc.Text = ""
End If
End Sub

Private Sub t_xcentajedesc_KeyPress(KeyAscii As Integer)
If KeyAscii < 48 Or KeyAscii > 57 Then
    If KeyAscii <> 8 And KeyAscii <> 46 Then
        KeyAscii = 0
        Beep
    End If
End If
End Sub

Private Sub Timer1_Timer()


    l_aviso.BackColor = &HFF&
    l_aviso.ForeColor = &HFFFFFF


End Sub

Private Sub Timer2_Timer()

    l_aviso.BackColor = &HFFFFFF
    l_aviso.ForeColor = &H80000008

End Sub
