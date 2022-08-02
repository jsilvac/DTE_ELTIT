VERSION 5.00
Object = "{4777436C-EB8C-4596-98A8-EBCF98683969}#1.0#0"; "FlexCell.ocx"
Object = "{085FBDF3-00C5-421B-B762-1D57299A2B89}#1.0#0"; "CLBUTN.OCX"
Object = "{ADBBAED6-B16F-40DC-80DF-B44910CBA76C}#1.0#0"; "Frame-Xp.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form form3323 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Generacion Formulario 3323"
   ClientHeight    =   9735
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   15270
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9735
   ScaleWidth      =   15270
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin XPFrame.FrameXp FrameQuickMenu 
      Height          =   615
      Left            =   11800
      TabIndex        =   53
      Top             =   45
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   1085
      BackColor       =   16744576
      Caption         =   " Mis Datos"
      BackColor       =   16744576
      BordeColor      =   4194304
      ColorBarraArriba=   4194304
      ColorBarraAbajo =   4194304
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   1
      Begin VB.CommandButton botonmisaccesos 
         Caption         =   "Permisos Modulo"
         Height          =   255
         Left            =   120
         TabIndex        =   55
         Top             =   280
         Width           =   1455
      End
      Begin VB.CommandButton botonmisfavoritos 
         Caption         =   "Mis Favoritos"
         Height          =   255
         Left            =   1800
         TabIndex        =   54
         Top             =   280
         Width           =   1335
      End
   End
   Begin XPFrame.FrameXp VIGENCIA 
      Height          =   7350
      Left            =   90
      TabIndex        =   7
      Top             =   2340
      Width           =   15000
      _ExtentX        =   26458
      _ExtentY        =   12965
      BackColor       =   16711680
      Caption         =   "PROCESO VIGENTE EN BASE DE DATOS"
      CaptionEstilo3D =   1
      BackColor       =   16711680
      ForeColor       =   65535
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      HabilitarArrastre=   -1  'True
      Begin VB.CommandButton Command2 
         BackColor       =   &H0000FF00&
         Caption         =   "GENERA FORM3323 SII"
         Height          =   240
         Left            =   11115
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   405
         Width           =   3660
      End
      Begin FlexCell.Grid Grid1 
         Height          =   3975
         Left            =   90
         TabIndex        =   9
         Top             =   720
         Width           =   14820
         _ExtentX        =   26141
         _ExtentY        =   7011
         Cols            =   5
         DefaultFontSize =   8.25
         Rows            =   30
      End
      Begin VB.Label Label22 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "TOTAL CASOS INFORMADOS"
         Height          =   285
         Left            =   2790
         TabIndex        =   51
         Top             =   6840
         Width           =   2895
      End
      Begin VB.Label Label21 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "TOTAL CASOS C/M/F/N/A"
         Height          =   285
         Left            =   2790
         TabIndex        =   50
         Top             =   6525
         Width           =   2895
      End
      Begin VB.Label TOTAL 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   285
         Index           =   20
         Left            =   5760
         TabIndex        =   49
         Top             =   6840
         Width           =   1905
      End
      Begin VB.Label TOTAL 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   285
         Index           =   19
         Left            =   5760
         TabIndex        =   48
         Top             =   6525
         Width           =   1905
      End
      Begin VB.Label TOTAL 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   285
         Index           =   11
         Left            =   13005
         TabIndex        =   47
         Top             =   4725
         Width           =   1905
      End
      Begin VB.Label TOTAL 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   285
         Index           =   12
         Left            =   13005
         TabIndex        =   46
         Top             =   5040
         Width           =   1905
      End
      Begin VB.Label Label20 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "TOTAL IVA COMPRAS SEM.ANTER."
         Height          =   285
         Left            =   10035
         TabIndex        =   45
         Top             =   4725
         Width           =   2895
      End
      Begin VB.Label Label19 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "TOTAL IVA COMPRAS SEM.SIGUI."
         Height          =   285
         Left            =   10035
         TabIndex        =   44
         Top             =   5040
         Width           =   2895
      End
      Begin VB.Label TOTAL 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   285
         Index           =   18
         Left            =   13005
         TabIndex        =   43
         Top             =   6930
         Width           =   1905
      End
      Begin VB.Label Label18 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "TOTAL VENTAS NO AFECTAS IVA"
         Height          =   285
         Left            =   10035
         TabIndex        =   42
         Top             =   6930
         Width           =   2895
      End
      Begin VB.Label TOTAL 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   285
         Index           =   13
         Left            =   13005
         TabIndex        =   41
         Top             =   5355
         Width           =   1905
      End
      Begin VB.Label TOTAL 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   285
         Index           =   14
         Left            =   13005
         TabIndex        =   40
         Top             =   5670
         Width           =   1905
      End
      Begin VB.Label TOTAL 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   285
         Index           =   15
         Left            =   13005
         TabIndex        =   39
         Top             =   5985
         Width           =   1905
      End
      Begin VB.Label TOTAL 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   285
         Index           =   16
         Left            =   13005
         TabIndex        =   38
         Top             =   6300
         Width           =   1905
      End
      Begin VB.Label TOTAL 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   285
         Index           =   17
         Left            =   13005
         TabIndex        =   37
         Top             =   6615
         Width           =   1905
      End
      Begin VB.Label Label17 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "TOTAL ANUAL IVA NO UTILIZADO"
         Height          =   285
         Left            =   10035
         TabIndex        =   36
         Top             =   5355
         Width           =   2895
      End
      Begin VB.Label Label16 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "TOTAL IVA COMUN"
         Height          =   285
         Left            =   10035
         TabIndex        =   35
         Top             =   5670
         Width           =   2895
      End
      Begin VB.Label Label15 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "TOTAL COMPRAS EXENTAS IVA"
         Height          =   285
         Left            =   10035
         TabIndex        =   34
         Top             =   5985
         Width           =   2895
      End
      Begin VB.Label Label14 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "TOTAL VENTAS EXENTAS IVA"
         Height          =   285
         Left            =   10035
         TabIndex        =   33
         Top             =   6300
         Width           =   2895
      End
      Begin VB.Label Label13 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "TOTAL COMPRAS NO AFECTAS IVA"
         Height          =   285
         Left            =   10035
         TabIndex        =   32
         Top             =   6615
         Width           =   2895
      End
      Begin VB.Label Label12 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "TOTAL NOTAS CREDITO EMITIDAS"
         Height          =   285
         Left            =   5085
         TabIndex        =   31
         Top             =   5985
         Width           =   2895
      End
      Begin VB.Label Label11 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "TOTAL IVA NOTA CREDITO"
         Height          =   285
         Left            =   5085
         TabIndex        =   30
         Top             =   5670
         Width           =   2895
      End
      Begin VB.Label Label10 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "TOTAL DOCUMENTOS EMITIDOS"
         Height          =   285
         Left            =   5085
         TabIndex        =   29
         Top             =   5355
         Width           =   2895
      End
      Begin VB.Label Label9 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "TOTAL IVA NO RETENIDO VENTAS"
         Height          =   285
         Left            =   5085
         TabIndex        =   28
         Top             =   5040
         Width           =   2895
      End
      Begin VB.Label Label8 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "TOTAL IVA VENTAS"
         Height          =   285
         Left            =   5085
         TabIndex        =   27
         Top             =   4725
         Width           =   2895
      End
      Begin VB.Label TOTAL 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   285
         Index           =   10
         Left            =   8055
         TabIndex        =   26
         Top             =   5985
         Width           =   1905
      End
      Begin VB.Label TOTAL 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   285
         Index           =   9
         Left            =   8055
         TabIndex        =   25
         Top             =   5670
         Width           =   1905
      End
      Begin VB.Label TOTAL 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   285
         Index           =   8
         Left            =   8055
         TabIndex        =   24
         Top             =   5355
         Width           =   1905
      End
      Begin VB.Label TOTAL 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   285
         Index           =   7
         Left            =   8055
         TabIndex        =   23
         Top             =   5040
         Width           =   1905
      End
      Begin VB.Label TOTAL 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   285
         Index           =   6
         Left            =   8055
         TabIndex        =   22
         Top             =   4725
         Width           =   1905
      End
      Begin VB.Label TOTAL 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   285
         Index           =   5
         Left            =   3060
         TabIndex        =   21
         Top             =   5985
         Width           =   1905
      End
      Begin VB.Label TOTAL 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   285
         Index           =   4
         Left            =   3060
         TabIndex        =   20
         Top             =   5670
         Width           =   1905
      End
      Begin VB.Label TOTAL 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   285
         Index           =   3
         Left            =   3060
         TabIndex        =   19
         Top             =   5355
         Width           =   1905
      End
      Begin VB.Label TOTAL 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   285
         Index           =   2
         Left            =   3060
         TabIndex        =   18
         Top             =   5040
         Width           =   1905
      End
      Begin VB.Label TOTAL 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   285
         Index           =   1
         Left            =   3060
         TabIndex        =   17
         Top             =   4725
         Width           =   1905
      End
      Begin VB.Label Label7 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "TOTAL NOTAS CREDITO RECIBIDAS"
         Height          =   285
         Left            =   90
         TabIndex        =   16
         Top             =   5985
         Width           =   2895
      End
      Begin VB.Label Label6 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "TOTAL IVA NOTA CREDITO"
         Height          =   285
         Left            =   90
         TabIndex        =   15
         Top             =   5670
         Width           =   2895
      End
      Begin VB.Label Label5 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "TOTAL DOCUMENTOS RECIBIDOS IVA COMPRAS"
         Height          =   285
         Left            =   90
         TabIndex        =   14
         Top             =   5355
         Width           =   2895
      End
      Begin VB.Label Label4 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "TOTAL IVA RETENIDO COMPRAS"
         Height          =   285
         Left            =   90
         TabIndex        =   13
         Top             =   5040
         Width           =   2895
      End
      Begin VB.Label Label3 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "TOTAL IVA COMPRAS"
         Height          =   285
         Left            =   90
         TabIndex        =   12
         Top             =   4725
         Width           =   2895
      End
      Begin VB.Label Label2 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   3735
         TabIndex        =   11
         Top             =   360
         Width           =   3570
      End
      Begin VB.Label Label1 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   90
         TabIndex        =   10
         Top             =   360
         Width           =   3570
      End
   End
   Begin XPFrame.FrameXp FrameXp2 
      Height          =   900
      Left            =   90
      TabIndex        =   1
      Top             =   1350
      Width           =   14985
      _ExtentX        =   26432
      _ExtentY        =   1588
      BackColor       =   16744576
      Caption         =   "PROCESO FORMULARIO"
      CaptionEstilo3D =   1
      BackColor       =   16744576
      ForeColor       =   65535
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   1
      Begin MSComctlLib.ProgressBar Barra 
         Height          =   375
         Left            =   135
         TabIndex        =   2
         Top             =   360
         Width           =   14745
         _ExtentX        =   26009
         _ExtentY        =   661
         _Version        =   393216
         Appearance      =   1
      End
   End
   Begin XPFrame.FrameXp FRMPROCESO 
      Height          =   1230
      Left            =   90
      TabIndex        =   0
      Top             =   45
      Width           =   14985
      _ExtentX        =   26432
      _ExtentY        =   2170
      BackColor       =   16744576
      Caption         =   "FORMULARIO 3323"
      CaptionEstilo3D =   1
      BackColor       =   16744576
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   1
      Begin VB.TextBox varipaso 
         Height          =   240
         Left            =   8040
         TabIndex        =   52
         Text            =   "Text1"
         Top             =   630
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00FF8080&
         Caption         =   "SEGUNDO SEMESTRE"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   3375
         TabIndex        =   6
         Top             =   495
         Width           =   3705
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00FF8080&
         Caption         =   "PRIMER SEMESTRE"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   45
         TabIndex        =   5
         Top             =   495
         Width           =   4830
      End
      Begin CoolButtons.cool_Button command1 
         Height          =   375
         Left            =   9240
         TabIndex        =   3
         Top             =   600
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         Caption         =   "COMIENZA ACTUALIZACION"
      End
      Begin VB.Label actualiza 
         BackColor       =   &H00FF8080&
         Height          =   465
         Left            =   1350
         TabIndex        =   4
         Top             =   630
         Width           =   3750
      End
   End
End
Attribute VB_Name = "form3323"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim debe(12) As Double
Dim haber(12) As Double
Dim totales(20) As Double
Dim año3323 As String
Dim periodo As String
Dim semestre As String

Dim FORMATOGRILLA(20, 20) As String

 

Private Sub Command1_Click()
If Grid1.Rows > 1 Then
    If MsgBox("DESEA BLANQUEAR EL PERIODO ACTUAL ", vbYesNo) = vbYes Then
        ELIMINAFORM3323
        genercompras
        generatrigo
        generaventas
    
    End If
End If
If Grid1.Rows = 1 Then
        genercompras
        generatrigo
        generaventas
End If

LEERFORM3323

End Sub


Sub genercompras()
LEERcompras
End Sub
Sub generatrigo()
'lEERTRIGO

End Sub
Sub generaventas()
LEERVENTAS
Rem LEERLIQUIDACIONES
End Sub
Sub LEERVENTAS()

Dim resultados As rdoResultset
    Dim csql As New rdoQuery
    Dim fecha1 As String
    Dim fecha2 As String
    Dim mes1 As String
    Dim mes2 As String
    
    If Option1.Value = True Then
    fecha1 = Format(fechasistema, "yyyy") + "-01-01"
    fecha2 = Format(fechasistema, "yyyy") + "-06-30"
    mes1 = "01"
    mes2 = "06"
    
    End If
    If Option2.Value = True Then
    fecha1 = Format(fechasistema, "yyyy") + "-07-01"
    fecha2 = Format(fechasistema, "yyyy") + "-12-31"
    mes1 = "07"
    mes2 = "12"
    
    End If

        Set csql.ActiveConnection = contadb
        
        csql.sql = "SELECT fv.tipo,fv.rut,sum(fv.iva),count(fv.numero) "
        csql.sql = csql.sql + "from facturasdeventas AS fv "
        csql.sql = csql.sql + "WHERE fv.FECHA>='" + fecha1 + "' AND fv.fecha<='" + fecha2 + "' and fv.iva<>'0' AND (tipo='1' OR TIPO='4') and fv.rut<>'0888888888' "
        csql.sql = csql.sql + "GROUP BY fv.rut,fv.TIPO "

        csql.Execute
        
        If csql.RowsAffected > 0 Then
        barra.Max = csql.RowsAffected + 1
        barra.Value = 0
        
        Set resultados = csql.OpenResultset
         While Not resultados.EOF
          barra.Value = barra.Value + 1
          If resultados(0) = "1" Then
          Call graba3323(resultados(1), "+", "0", "0", "0", "0", "0", resultados(2), "0", resultados(3), "0", "0")
          Else
          Call graba3323(resultados(1), "+", "0", "0", "0", "0", "0", "0", "0", "0", resultados(2), resultados(3))
          End If
          resultados.MoveNext
           
         Wend
          resultados.Close
            Set resultados = Nothing

        End If
        
End Sub
Sub LEERLIQUIDACIONES()

Dim resultados As rdoResultset
    Dim csql As New rdoQuery
    Dim fecha1 As String
    Dim fecha2 As String
    Dim mes1 As String
    Dim mes2 As String
    
    If Option1.Value = True Then
    fecha1 = Format(fechasistema, "yyyy") + "-01-01"
    fecha2 = Format(fechasistema, "yyyy") + "-06-30"
    mes1 = "01"
    mes2 = "06"
    
    End If
    If Option2.Value = True Then
    fecha1 = Format(fechasistema, "yyyy") + "-07-01"
    fecha2 = Format(fechasistema, "yyyy") + "-12-31"
    mes1 = "07"
    mes2 = "12"
    
    End If

        Set csql.ActiveConnection = contadb
        
        csql.sql = "SELECT FV.TIPO,FV.rut,fV.fecha,sum(FV.iva),SUM(FV.IVARETENIDO),SUM(FV.IVACOMISION),count(FV.numero) "
        csql.sql = csql.sql + "from MALI AS FV "
        csql.sql = csql.sql + "WHERE FV.FECHA>='" + fecha1 + "' AND FV.FECHA<='" + fecha2 + "' and fv.iva<>'0' "
        csql.sql = csql.sql + "GROUP BY FV.rut,FV.TIPO "

        csql.Execute
        
        If csql.RowsAffected > 0 Then
        barra.Max = csql.RowsAffected + 1
        barra.Value = 0
        
        Set resultados = csql.OpenResultset
         While Not resultados.EOF
          barra.Value = barra.Value + 1
          
          Call graba3323(resultados(1), "+", resultados(3), resultados(4), resultados(6), "0", "0", resultados(5), "0", resultados(6), "0", "0")
          resultados.MoveNext
           
         Wend
          resultados.Close
            Set resultados = Nothing

        End If
        
End Sub

Sub LEERVENTASEXENTAS()

Dim resultados As rdoResultset
    Dim csql As New rdoQuery
    Dim fecha1 As String
    Dim fecha2 As String
    Dim mes1 As String
    Dim mes2 As String
    Dim sumas As Double
    Dim montos As Double
    
    
    If Option1.Value = True Then
    fecha1 = Format(fechasistema, "yyyy") + "-01-01"
    fecha2 = Format(fechasistema, "yyyy") + "-06-30"
    mes1 = "01"
    mes2 = "06"
    
    End If
    If Option2.Value = True Then
    fecha1 = Format(fechasistema, "yyyy") + "-07-01"
    fecha2 = Format(fechasistema, "yyyy") + "-12-31"
    mes1 = "07"
    mes2 = "12"
    
    End If

        Set csql.ActiveConnection = contadb
        
        
        csql.sql = "SELECT fv.tipo,fv.rut,sum(fv.neto),count(fv.numero) "
        csql.sql = csql.sql + "from facturasdeventas AS fv "
        csql.sql = csql.sql + "WHERE fv.fecha>='" + fecha1 + "' AND fv.fecha<='" + fecha2 + "'  AND iva='0' "
        csql.sql = csql.sql + "GROUP BY fv.rut,fv.tipo "

        csql.Execute
        sumas = 0
        If csql.RowsAffected > 0 Then
        
        Set resultados = csql.OpenResultset
         While Not resultados.EOF
         montos = resultados(2)
         sumas = sumas + montos
          
          resultados.MoveNext
           
         Wend
          resultados.Close
            Set resultados = Nothing

        End If
        totales(16) = sumas
        TOTAL(16).Caption = Format(totales(16), "#,###,###,###")
        
          
        
End Sub


Private Sub COMMAND2_Click()
generaform3323


End Sub
Sub generaform3323()
Dim cadena As String
Dim vari1 As String
Dim AÑOPROCESO As String
Dim MESPROCESO As String
Dim CORRELATIVO As String
Dim RUTEMPRESAPROCESO As String
Dim TIPOINFORME As String
Dim DIGITOVERIFICADOR As String
Dim NOMBREEMPRESAPROCESO As String * 30
Dim monto As String
Dim RUTCLIENTE As String
Dim i As Double

Close 20
Open App.path + "\form3323.txt" For Output As #20
AÑOPROCESO = Format(fechasistema, "YYYY")
RUTEMPRESAPROCESO = Mid(rutempresa, 1, 8)
DIGITOVERIFICADOR = Mid(rutempresa, 10, 1)
TIPOINFORME = "O"
NOMBREEMPRESAPROCESO = nombreempresa

CORRELATIVO = "00"

If Option1.Value = True Then MESPROCESO = "01" Else MESPROCESO = "02"
cadena = "11" + AÑOPROCESO + MESPROCESO + "3323I" + RUTEMPRESAPROCESO + CORRELATIVO + TIPOINFORME + "000000000 0000000000      " + RUTEMPRESAPROCESO + DIGITOVERIFICADOR + NOMBREEMPRESAPROCESO + String(117, 32)
Print #20, cadena
cadena = "12" + String(23, 32) + "  00000000000000000" + String(162, 32)
Print #20, cadena
For k = 1 To Grid1.Rows - 1
cadena = "213323I" + RUTEMPRESAPROCESO + CORRELATIVO + "O" + AÑOPROCESO + MESPROCESO + RUTEMPRESAPROCESO + DIGITOVERIFICADOR
RUTCLIENTE = Mid(Grid1.Cell(k, 1).text, 2, 8)
RUTCLIENTE = RUTCLIENTE + rut("0" + RUTCLIENTE)
cadena = cadena + RUTCLIENTE
For i = 3 To 12
varipaso.MaxLength = 15
varipaso.text = Grid1.Cell(k, i).text
Call ceros(varipaso)
cadena = cadena + varipaso.text
Next i
cadena = cadena + " 0000000000000"

Print #20, cadena
s = Len(cadena)
Next k

cadena = "313323I" + RUTEMPRESAPROCESO + CORRELATIVO + AÑOPROCESO + MESPROCESO + RUTEMPRESAPROCESO + DIGITOVERIFICADOR
For k = 1 To 10
varipaso.MaxLength = 18
If k = 3 Then varipaso.MaxLength = 16
If k = 5 Then varipaso.MaxLength = 16
If k = 8 Then varipaso.MaxLength = 16
If k = 10 Then varipaso.MaxLength = 16

varipaso.text = Replace(TOTAL(k), ".", "")
Call ceros(varipaso)
cadena = cadena + varipaso.text

Next k
cadena = cadena + String(2, 32)

Print #20, cadena

cadena = "32"
Rem CASOS CMF
varipaso.MaxLength = 7
varipaso.text = Replace(TOTAL(19), ".", "")
Call ceros(varipaso)
cadena = cadena + varipaso.text
Rem CASOS INFORMADOS
varipaso.MaxLength = 7
varipaso.text = Replace(TOTAL(20), ".", "")
Call ceros(varipaso)
cadena = cadena + varipaso.text
Rem DE LOS IVAS ANTERIOR Y SIGUIENTE ABAJO

For k = 11 To 18

varipaso.MaxLength = 18
varipaso.text = Replace(TOTAL(k), ".", "")
Call ceros(varipaso)
cadena = cadena + varipaso.text
Next k
cadena = cadena + "00000000" + String(38, 32)
Print #20, cadena
Close 20
Unload Me
Shell "NOTEPAD " + App.path + "\form3323.TXT"

End Sub
Private Sub Form_Load()
   
     
    Call Conectar_BD
    Call Conectarconta(Servidor, clientesistema + "conta", Usuario, password)
   Rem  Call Conectarventas(servidor, "molino_" + "ventas00", usuario, password)
    
    FRMPROCESO.Caption = "GENERA FORMULARIO 3323 AÑO:" + Format(fechasistema, "YYYY") + " " + nombreempresa
    Option1.Value = True
    VIGENCIA.Visible = False
    
    CARGAGRILLA
    LEERFORM3323
    
    
End Sub
Sub LEERFORM3323()
Dim SUMAR As Double

Dim resultados As rdoResultset
    Dim csql As New rdoQuery
       
        Set csql.ActiveConnection = contadb
        
        csql.sql = "SELECT * from form3323 order by rut "
        csql.Execute
        If csql.RowsAffected > 0 Then
        Grid1.Rows = 1
        Grid1.AutoRedraw = False
        
        
        VIGENCIA.Visible = True
        Set resultados = csql.OpenResultset
        año3323 = resultados(13)
        semestre = resultados(14)
        For k = 1 To 20
        totales(k) = 0
        Next k
        
         While Not resultados.EOF
           For k = 3 To 12
           SUMAR = SUMAR + resultados(k)
           Next k
           If SUMAR <> 0 Then
          
          Grid1.Rows = Grid1.Rows + 1
          Grid1.Cell(Grid1.Rows - 1, 1).text = resultados(1)
          Grid1.Cell(Grid1.Rows - 1, 2).text = LEERNOMBREPROVEEDOR(resultados(1))
          
          Grid1.Cell(Grid1.Rows - 1, 3).text = resultados(3)
          Grid1.Cell(Grid1.Rows - 1, 4).text = resultados(4)
          Grid1.Cell(Grid1.Rows - 1, 5).text = resultados(5)
          Grid1.Cell(Grid1.Rows - 1, 6).text = resultados(6)
          Grid1.Cell(Grid1.Rows - 1, 7).text = resultados(7)
          Grid1.Cell(Grid1.Rows - 1, 8).text = resultados(8)
          Grid1.Cell(Grid1.Rows - 1, 9).text = resultados(9)
          Grid1.Cell(Grid1.Rows - 1, 10).text = resultados(10)
          Grid1.Cell(Grid1.Rows - 1, 11).text = resultados(11)
          Grid1.Cell(Grid1.Rows - 1, 12).text = resultados(12)
          For k = 1 To 10
          
          totales(k) = totales(k) + resultados(k + 2)
          Next k
          totales(20) = totales(20) + 1
          End If
          
          resultados.MoveNext
           
         Wend
          resultados.Close
            Set resultados = Nothing

        End If
        For k = 1 To 20
        TOTAL(k).Caption = Format(totales(k), "#,###,###,###")
        Next k
        Label1.Caption = "PERIODO AÑO :" + año3323
        If semestre = "1" Then
        Label2.Caption = "PRIMER SEMESTRE"
        Else
        Label2.Caption = "SEGUNDO SEMESTRE"
        
        End If
    LEERanteriores
    LEERSIGUIENTES
    LEERCOMPRASEXENTAS
    LEERivasnousados
    LEERVENTASEXENTAS
    
    Grid1.AutoRedraw = True
    Grid1.Refresh
        
End Sub

Sub LEERcompras()

Dim resultados As rdoResultset
    Dim csql As New rdoQuery
    Dim fecha1 As String
    Dim fecha2 As String
    Dim mes1 As String
    Dim mes2 As String
    
    If Option1.Value = True Then
    fecha1 = Format(fechasistema, "yyyy") + "-01-01"
    fecha2 = Format(fechasistema, "yyyy") + "-06-30"
    mes1 = "01"
    mes2 = "06"
    
    End If
    If Option2.Value = True Then
    fecha1 = Format(fechasistema, "yyyy") + "-07-01"
    fecha2 = Format(fechasistema, "yyyy") + "-12-31"
    mes1 = "07"
    mes2 = "12"
    
    End If

        Set csql.ActiveConnection = contadb
        
        csql.sql = "SELECT fc.tipo,fc.rut,fc.fecha,sum(fc.iva),sum(fc.retencion),count(fc.numero) "
        csql.sql = csql.sql + "from facturasdecompras AS fc "
        csql.sql = csql.sql + "WHERE fc.fecha>='" + fecha1 + "' AND fc.fecha<='" + fecha2 + "' and fc.añocontable='" + Format(fechasistema, "yyyy") + "' and fc.mescontable>='" + mes1 + "' and fc.mescontable<='" + mes2 + "' and fc.iva<>'0' "
        csql.sql = csql.sql + "GROUP BY fc.rut,fc.tipo "

        csql.Execute
        
        If csql.RowsAffected > 0 Then
        barra.Max = csql.RowsAffected + 1
        barra.Value = 0
        
        Set resultados = csql.OpenResultset
         While Not resultados.EOF
          barra.Value = barra.Value + 1
          If resultados(0) <> "3" And resultados(0) <> "6" Then
          Call graba3323(resultados(1), "+", resultados(3), resultados(4), resultados(5), "0", "0", "0", "0", "0", "0", "0")
          Else
          Call graba3323(resultados(1), "+", "0", "0", "0", resultados(3), resultados(5), "0", "0", "0", "0", "0")
          End If
          resultados.MoveNext
         Wend
          resultados.Close
            Set resultados = Nothing

        End If
        
End Sub
Sub LEERTRIGO()

Dim resultados As rdoResultset
    Dim csql As New rdoQuery
    Dim fecha1 As String
    Dim fecha2 As String
    Dim mes1 As String
    Dim mes2 As String
    
    If Option1.Value = True Then
    fecha1 = Format(fechasistema, "yyyy") + "-01-01"
    fecha2 = Format(fechasistema, "yyyy") + "-06-30"
    mes1 = "01"
    mes2 = "06"
    
    End If
    If Option2.Value = True Then
    fecha1 = Format(fechasistema, "yyyy") + "-07-01"
    fecha2 = Format(fechasistema, "yyyy") + "-12-31"
    mes1 = "07"
    mes2 = "12"
    
    End If

        Set csql.ActiveConnection = contadb
        
        csql.sql = "SELECT FC.TIPO,MID(FC.PROVEEDOR,1,10),fc.fechaemision,sum(FC.IVARETENIDO)+SUM(FC.IVANORETENIDO),sum(FC.IVAretenido),count(FC.numero) "
        csql.sql = csql.sql + "from molino_trigo.FACTURAS AS FC "
        csql.sql = csql.sql + "WHERE FC.FECHAemision>='" + fecha1 + "' AND FC.FECHAemision<='" + fecha2 + "' and mid(fc.proveedor,2,5)<>'88888' "
        csql.sql = csql.sql + "GROUP BY FC.proveedor,FC.TIPO "
        csql.sql = csql.sql + " UNION "
        csql.sql = csql.sql + "SELECT FC.TIPO,mid(FC.PROVEEDOR,1,10),fc.fechaemision,sum(FC.IVARETENIDO)+SUM(FC.IVANORETENIDO),sum(FC.IVAretenido),count(FC.numero) "
        csql.sql = csql.sql + "from ALLIPEN_trigo.FACTURAS AS FC "
        csql.sql = csql.sql + "WHERE FC.FECHAemision>='" + fecha1 + "' AND FC.FECHAemision<='" + fecha2 + "' and mid(fc.proveedor,2,5)<>'88888' "
        csql.sql = csql.sql + "GROUP BY FC.proveedor,FC.TIPO "


        csql.Execute
        
        If csql.RowsAffected > 0 Then
        barra.Max = csql.RowsAffected + 1
        barra.Value = 0
        
        Set resultados = csql.OpenResultset
         While Not resultados.EOF
          barra.Value = barra.Value + 1
          If resultados(0) <> "3" Then
          Call graba3323(resultados(1), "+", resultados(3), resultados(4), resultados(5), "0", "0", "0", "0", "0", "0", "0")
          Else
          Call graba3323(resultados(1), "+", "0", "0", "0", resultados(3), resultados(5), "0", "0", "0", "0", "0")
          End If
          resultados.MoveNext
           
         Wend
          resultados.Close
            Set resultados = Nothing

        End If
        
End Sub


Sub LEERCOMPRASEXENTAS()

Dim resultados As rdoResultset
    Dim csql As New rdoQuery
    Dim fecha1 As String
    Dim fecha2 As String
    Dim mes1 As String
    Dim mes2 As String
    Dim sumas As Double
    Dim montos As Double
    
    If Option1.Value = True Then
    fecha1 = Format(fechasistema, "yyyy") + "-01-01"
    fecha2 = Format(fechasistema, "yyyy") + "-06-30"
    mes1 = "01"
    mes2 = "06"
    
    End If
    If Option2.Value = True Then
    fecha1 = Format(fechasistema, "yyyy") + "-07-01"
    fecha2 = Format(fechasistema, "yyyy") + "-12-31"
    mes1 = "07"
    mes2 = "12"
    
    End If

        Set csql.ActiveConnection = contadb
        
        csql.sql = "SELECT fc.tipo,fc.rut,fc.fecha,sum(fc.iva),sum(fc.exento),count(fc.numero) "
        csql.sql = csql.sql + "from facturasdecompras AS fc "
        csql.sql = csql.sql + "WHERE fc.fecha>='" + fecha1 + "' AND fc.fecha<='" + fecha2 + "' and fc.añocontable='" + Format(fechasistema, "yyyy") + "' and fc.mescontable>='" + mes1 + "' and fc.mescontable<='" + mes2 + "' and iva='0' "
        csql.sql = csql.sql + "GROUP BY fc.tipo "

        csql.Execute
        sumas = 0
        If csql.RowsAffected > 0 Then
        
        Set resultados = csql.OpenResultset
         While Not resultados.EOF
         montos = resultados(4)
         If resultados(0) = "NC" Then
         montos = resultados(4) * -1
         End If
         sumas = sumas + montos
          
          
          resultados.MoveNext
           
         Wend
          resultados.Close
            Set resultados = Nothing

        End If
        totales(15) = sumas
        TOTAL(15).Caption = Format(totales(15), "#,###,###,###")
        
End Sub
Sub LEERivasnousados()

Dim resultados As rdoResultset
    Dim csql As New rdoQuery
    Dim fecha1 As String
    Dim fecha2 As String
    Dim mes1 As String
    Dim mes2 As String
    Dim sumas As Double
    Dim montos As Double
    Dim NETO As Double
    
    If Option1.Value = True Then
    fecha1 = Format(fechasistema, "yyyy") + "-01-01"
    fecha2 = Format(fechasistema, "yyyy") + "-06-30"
    mes1 = "01"
    mes2 = "06"
    
    End If
    If Option2.Value = True Then
    fecha1 = Format(fechasistema, "yyyy") + "-07-01"
    fecha2 = Format(fechasistema, "yyyy") + "-12-31"
    mes1 = "07"
    mes2 = "12"
    
    End If

        Set csql.ActiveConnection = contadb
        
        csql.sql = "SELECT fc.tipo,fc.rut,fc.fecha,sum(fc.neto),sum(fc.exento),count(fc.numero) "
        csql.sql = csql.sql + "from facturasdecompras AS fc "
        csql.sql = csql.sql + "WHERE fc.fecha>='" + fecha1 + "' AND fc.fecha<='" + fecha2 + "' and fc.añocontable='" + Format(fechasistema, "yyyy") + "' and fc.mescontable>='" + mes1 + "' and fc.mescontable<='" + mes2 + "' and iva='0' and neto<>'0' "
        csql.sql = csql.sql + "GROUP BY fc.tipo "

        csql.Execute
        
        If csql.RowsAffected > 0 Then
        sumas = 0
        Set resultados = csql.OpenResultset
         While Not resultados.EOF
         NETO = resultados(3) / 1.19
         montos = Round(resultados(3) - NETO)
         If resultados(0) = "NC" Then
         montos = resultados(3) * -1
         End If
         sumas = sumas + montos
          
          
          resultados.MoveNext
           
         Wend
          resultados.Close
            Set resultados = Nothing

        End If
        totales(13) = sumas
        TOTAL(13).Caption = Format(totales(13), "#,###,###,###")
        
End Sub


Sub LEERanteriores()
    
Dim resultados As rdoResultset
    Dim csql As New rdoQuery
    Dim fecha1 As String
    Dim fecha2 As String
    Dim mes1 As String
    Dim mes2 As String
    Dim sumas As Double
    Dim montos As Double
    
    If semestre = "1" Then Option1.Value = True Else Option2.Value = True
    
    If Option1.Value = True Then
    fecha1 = Format(fechasistema, "yyyy") + "-01-01"
    fecha2 = Format(fechasistema, "yyyy") + "-06-31"
    mes1 = "01"
    mes2 = "06"
    
    End If
    If Option2.Value = True Then
    fecha1 = Format(fechasistema, "yyyy") + "-07-01"
    fecha2 = Format(fechasistema, "yyyy") + "-12-31"
    mes1 = "07"
    mes2 = "12"
    
    End If

        Set csql.ActiveConnection = contadb
        
        csql.sql = "SELECT fc.tipo,fc.rut,fc.fecha,sum(fc.iva),sum(fc.retencion),count(fc.numero) "
        csql.sql = csql.sql + "from facturasdecompras AS fc "
        csql.sql = csql.sql + "WHERE fc.FECHA<'" + fecha1 + "' and fc.añocontable='" + Format(fechasistema, "yyyy") + "' and fc.mescontable>='" + mes1 + "' and fc.mescontable<='" + mes2 + "' and iva<>'0' "
        csql.sql = csql.sql + "GROUP BY fc.TIPO "
        csql.Execute
        sumas = 0
        If csql.RowsAffected > 0 Then
        Set resultados = csql.OpenResultset
         While Not resultados.EOF
         montos = resultados(3)
         If resultados(0) = "NC" Then
         montos = resultados(3) * -1
         End If
         sumas = sumas + montos
         resultados.MoveNext
         Wend
          resultados.Close
            Set resultados = Nothing
        End If
        totales(11) = sumas
        TOTAL(11).Caption = Format(totales(11), "#,###,###,###")
        
End Sub
Sub LEERSIGUIENTES()
    
Dim resultados As rdoResultset
    Dim csql As New rdoQuery
    Dim fecha1 As String
    Dim fecha2 As String
    Dim mes1 As String
    Dim mes2 As String
    Dim sumas As Double
    Dim montos As Double
    
    If semestre = "1" Then Option1.Value = True Else Option2.Value = True
    
    If Option2.Value = True Then
    fecha1 = Format(fechasistema, "yyyy") + "-01-01"
    fecha2 = Format(fechasistema, "yyyy") + "-06-31"
    mes1 = "01"
    mes2 = "06"
    
    End If
    If Option1.Value = True Then
    fecha1 = Format(fechasistema, "yyyy") + "-07-01"
    fecha2 = Format(fechasistema, "yyyy") + "-12-31"
    mes1 = "07"
    mes2 = "12"
    
    End If

        Set csql.ActiveConnection = contadb
        
        csql.sql = "SELECT fc.tipo,fc.rut,fc.fecha,sum(fc.iva),sum(fc.retencion),count(fc.numero) "
        csql.sql = csql.sql + "from facturasdecompras AS fc "
        csql.sql = csql.sql + "WHERE fc.fecha<'" + fecha1 + "' and fc.añocontable='" + Format(fechasistema, "yyyy") + "' and fc.mescontable>='" + mes1 + "' and fc.mescontable<='" + mes2 + "' and iva<>'0' "
        csql.sql = csql.sql + "GROUP BY fc.tipo "
        csql.Execute
        sumas = 0
        If csql.RowsAffected > 0 Then
        Set resultados = csql.OpenResultset
         While Not resultados.EOF
         montos = resultados(3)
         If resultados(0) = "NC" Then
         montos = resultados(3) * -1
         End If
         sumas = sumas + montos
         resultados.MoveNext
         Wend
          resultados.Close
            Set resultados = Nothing
        End If
        
        
        totales(12) = sumas
        TOTAL(12).Caption = Format(totales(12), "#,###,###,###")
        
End Sub

Sub ELIMINAFORM3323()

Dim resultados As rdoResultset
    Dim csql As New rdoQuery
   
        Set csql.ActiveConnection = contadb
        
        csql.sql = "DELETE FROM form3323 "
        
        csql.Execute
        Call sincronizadatos(csql.sql, contadb, "")
        
        
        
End Sub


Sub graba3323(rut, tipo, com_fc_iva, com_fc_ivaretenido, com_fc_cantidad, com_nc_iva, com_nc_cantidad, ven_fv_iva, ven_fv_ivanoretenido, ven_fv_cantidad, ven_nc_iva, ven_nc_cantidad)
   
    

    campos(0, 0) = "rut"
    campos(1, 0) = "tipo"
    campos(2, 0) = "com_fc_iva"
    campos(3, 0) = "com_fc_ivaretenido"
    campos(4, 0) = "com_fc_cantidad"
    campos(5, 0) = "com_nc_iva"
    campos(6, 0) = "com_nc_cantidad"
    campos(7, 0) = "ven_fv_iva"
    campos(8, 0) = "ven_fv_ivanoretenido"
    campos(9, 0) = "ven_fv_cantidad"
    campos(10, 0) = "ven_nc_iva"
    campos(11, 0) = "ven_nc_cantidad"
    campos(12, 0) = "año"
    campos(13, 0) = "semestre"
    
    
    campos(14, 0) = ""
    campos(0, 2) = "form3323"
    condicion = "rut=" + "'" + rut + "' and tipo='" + tipo + "'"
    op = 5
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)
    If sqlconta.status = 4 Then
    condicion = ""
    op = 2
    Else
    op = 3
    End If
    
    campos(0, 1) = rut
    campos(1, 1) = tipo
    campos(2, 1) = sqlconta.response(2, 3) + CDbl(com_fc_iva)
    campos(3, 1) = sqlconta.response(3, 3) + CDbl(com_fc_ivaretenido)
    campos(4, 1) = sqlconta.response(4, 3) + CDbl(com_fc_cantidad)
    campos(5, 1) = sqlconta.response(5, 3) + CDbl(com_nc_iva)
    campos(6, 1) = sqlconta.response(6, 3) + CDbl(com_nc_cantidad)
    campos(7, 1) = sqlconta.response(7, 3) + CDbl(ven_fv_iva)
    campos(8, 1) = sqlconta.response(8, 3) + CDbl(ven_fv_ivanoretenido)
    campos(9, 1) = sqlconta.response(9, 3) + CDbl(ven_fv_cantidad)
    campos(10, 1) = sqlconta.response(10, 3) + CDbl(ven_nc_iva)
    campos(11, 1) = sqlconta.response(11, 3) + CDbl(ven_nc_cantidad)
    campos(12, 1) = Format(fechasistema, "yyyy")
    If Option1.Value = True Then
    campos(13, 1) = "1"
    Else
    campos(13, 1) = "2"
    End If
    
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)
    
    

End Sub

Sub CARGAGRILLA()
Rem DATOS DE LA COLUMNA
    Grid1.DefaultFont.Size = 7
    FORMATOGRILLA(1, 1) = "RUT"
    FORMATOGRILLA(1, 2) = "NOMBRE"
    FORMATOGRILLA(1, 3) = "IVA FC "
    FORMATOGRILLA(1, 4) = "IVA RET FC"
    FORMATOGRILLA(1, 5) = "CANT.FC"
    FORMATOGRILLA(1, 6) = "IVA NC "
    FORMATOGRILLA(1, 7) = "CANT.NC"
    FORMATOGRILLA(1, 8) = "IVA FV"
    FORMATOGRILLA(1, 9) = "IVA NO RET FV"
    FORMATOGRILLA(1, 10) = "CANT.FV"
    FORMATOGRILLA(1, 11) = "IVA NC"
    FORMATOGRILLA(1, 12) = "CANT.NC"
     
    Rem LARGO DE LOS DATOS
    
    FORMATOGRILLA(2, 1) = "9"
    FORMATOGRILLA(2, 2) = "30"
    FORMATOGRILLA(2, 3) = "10"
    FORMATOGRILLA(2, 4) = "10"
    FORMATOGRILLA(2, 5) = "10"
    FORMATOGRILLA(2, 6) = "10"
    FORMATOGRILLA(2, 7) = "10"
    FORMATOGRILLA(2, 8) = "10"
    FORMATOGRILLA(2, 9) = "10"
    FORMATOGRILLA(2, 10) = "10"
    FORMATOGRILLA(2, 11) = "10"
    FORMATOGRILLA(2, 12) = "10"

    
    Rem TIPO DE DATOS C=CEROS,S=STRING,N=NUMERICO,D=DATE,
    FORMATOGRILLA(3, 1) = "S"
    FORMATOGRILLA(3, 2) = "S"
    FORMATOGRILLA(3, 3) = "N"
    FORMATOGRILLA(3, 4) = "N"
    FORMATOGRILLA(3, 5) = "N"
    FORMATOGRILLA(3, 6) = "N"
    FORMATOGRILLA(3, 7) = "N"
    FORMATOGRILLA(3, 8) = "N"
    FORMATOGRILLA(3, 9) = "N"
    FORMATOGRILLA(3, 10) = "N"
    FORMATOGRILLA(3, 11) = "N"
    FORMATOGRILLA(3, 12) = "N"
    
    
    Rem FORMATO GRILLA
    For k = 3 To 12
    
    FORMATOGRILLA(4, k) = "###,###,###"
    Next k
    
    
    
    Rem LOCCKED
    FORMATOGRILLA(5, 1) = "TRUE"
    FORMATOGRILLA(5, 2) = "TRUE"
    FORMATOGRILLA(5, 3) = "TRUE"
    FORMATOGRILLA(5, 4) = "TRUE"
    FORMATOGRILLA(5, 5) = "TRUE"
    FORMATOGRILLA(5, 6) = "TRUE"
    FORMATOGRILLA(5, 7) = "TRUE"
    FORMATOGRILLA(5, 8) = "TRUE"
    FORMATOGRILLA(5, 9) = "TRUE"
    FORMATOGRILLA(5, 10) = "TRUE"
    FORMATOGRILLA(5, 11) = "TRUE"
    FORMATOGRILLA(5, 12) = "TRUE"
    FORMATOGRILLA(5, 13) = "TRUE"
    FORMATOGRILLA(5, 14) = "TRUE"
    FORMATOGRILLA(5, 15) = "TRUE"
    
    Grid1.Cols = 13
    Grid1.Rows = 1
    
     'infogrilla.grid1.AllowUserResizing = False
    Grid1.DisplayFocusRect = False
    'infogrilla.grid1.ExtendLastCol = True
    Grid1.BoldFixedCell = False
    
    Grid1.DrawMode = cellOwnerDraw
    
    Grid1.Appearance = Flat
    Grid1.ScrollBarStyle = Flat
    Grid1.FixedRowColStyle = Flat
    
   'infogrilla.grid1.BackColorFixed = RGB(90, 158, 214)
   ' infogrilla.grid1.BackColorFixedSel = RGB(110, 180, 230)
   ' infogrilla.grid1.BackColorBkg = RGB(90, 158, 214)
   ' infogrilla.grid1.BackColorScrollBar = RGB(231, 235, 247)
   ' infogrilla.grid1.BackColor1 = RGB(231, 235, 247)
   ' infogrilla.grid1.BackColor2 = RGB(239, 243, 255)
   ' infogrilla.grid1.GridColor = RGB(148, 190, 231)
    Grid1.Column(0).Width = 0
    
    For k = 1 To Grid1.Cols - 1
        
        Grid1.Cell(0, k).text = FORMATOGRILLA(1, k)
        Grid1.Column(k).Width = Val(FORMATOGRILLA(2, k)) * Grid1.DefaultFont.Size
       
        Grid1.Column(k).MaxLength = Val(FORMATOGRILLA(2, k))
        Grid1.Column(k).FormatString = FORMATOGRILLA(4, k)
        Grid1.Column(k).Locked = FORMATOGRILLA(5, k)
        If FORMATOGRILLA(3, k) = "N" Then Grid1.Column(k).Alignment = cellRightCenter
        If FORMATOGRILLA(3, k) = "D" Then Grid1.Column(k).CellType = cellCalendar
        
    Next k
End Sub

Private Function LEERNOMBREPROVEEDOR(rut) As String



    campos(0, 0) = "nombre"
    campos(1, 0) = ""
    campos(0, 2) = "cuentascorrientes"
    condicion = "rut=" + "'" + rut + "'"
    op = 5
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)
    If sqlconta.status = 0 Then
    LEERNOMBREPROVEEDOR = sqlconta.response(0, 3)
    End If
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)
    
    

End Function
Private Sub botonmisaccesos_Click()
programafiltro = Me.Caption
misaccesos.Show
End Sub


Private Sub botonmisfavoritos_Click()
Call AgregaFavoritos(USUARIOSISTEMA, App.EXEName, Me.Name, Me.Caption)
End Sub
