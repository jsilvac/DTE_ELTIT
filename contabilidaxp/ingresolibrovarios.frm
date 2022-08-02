VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash.ocx"
Object = "{4777436C-EB8C-4596-98A8-EBCF98683969}#1.0#0"; "FlexCell.ocx"
Object = "{ADBBAED6-B16F-40DC-80DF-B44910CBA76C}#1.0#0"; "Frame-Xp.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form publi0009 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   10125
   ClientLeft      =   45
   ClientTop       =   1110
   ClientWidth     =   14280
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   10125
   ScaleWidth      =   14280
   ShowInTaskbar   =   0   'False
   Begin XPFrame.FrameXp frmTipo 
      Height          =   2280
      Left            =   9120
      TabIndex        =   64
      Top             =   3840
      Visible         =   0   'False
      Width           =   2715
      _ExtentX        =   4789
      _ExtentY        =   4022
      BackColor       =   8421631
      Caption         =   "Tipo Documento"
      CaptionEstilo3D =   1
      BackColor       =   8421631
      ColorBarraArriba=   12632319
      ColorBarraAbajo =   128
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
      Begin VB.Label Label14 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " 5 - Nota de Debito"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   120
         TabIndex        =   84
         Top             =   1920
         Width           =   2475
      End
      Begin VB.Label Label13 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " 4 - Factura Exenta"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   120
         TabIndex        =   83
         Top             =   1560
         Width           =   2475
      End
      Begin VB.Label Label12 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " 2 - Boleta"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   135
         TabIndex        =   67
         Top             =   840
         Width           =   2475
      End
      Begin VB.Label lbl24 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " 3 - Nota de Credito"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   120
         TabIndex        =   66
         Top             =   1200
         Width           =   2475
      End
      Begin VB.Label lbl23 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " 1 - Factura"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   135
         TabIndex        =   65
         Top             =   495
         Width           =   2475
      End
   End
   Begin XPFrame.FrameXp frmdatos 
      Height          =   645
      Left            =   7560
      TabIndex        =   68
      Top             =   3600
      Visible         =   0   'False
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   1138
      BackColor       =   255
      Caption         =   "DATOS DOCUMENTO"
      CaptionEstilo3D =   1
      BackColor       =   255
      ForeColor       =   8438015
      ColorBarraArriba=   16744576
      ColorBarraAbajo =   16711680
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.TextBox dato14 
         BackColor       =   &H00E1FFFD&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   960
         MaxLength       =   2
         TabIndex        =   73
         Tag             =   "codigo"
         Top             =   240
         Width           =   375
      End
      Begin VB.TextBox dato17 
         BackColor       =   &H00E1FFFD&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   3120
         MaxLength       =   4
         TabIndex        =   72
         Tag             =   "fecha"
         Top             =   240
         Width           =   615
      End
      Begin VB.TextBox dato16 
         BackColor       =   &H00E1FFFD&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   2760
         MaxLength       =   2
         TabIndex        =   71
         Tag             =   "fecha"
         Top             =   240
         Width           =   375
      End
      Begin VB.TextBox dato15 
         BackColor       =   &H00E1FFFD&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   2400
         MaxLength       =   2
         TabIndex        =   70
         Tag             =   "fecha"
         Top             =   240
         Width           =   375
      End
      Begin VB.TextBox dato18 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E1FFFD&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   4920
         MaxLength       =   10
         TabIndex        =   69
         Tag             =   "numero"
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   " FECHA"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   1440
         TabIndex        =   76
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   " NUMERO"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   3840
         TabIndex        =   75
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "TIPO"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   120
         TabIndex        =   74
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.CommandButton CmdFavoritos 
      BackColor       =   &H0080FF80&
      Caption         =   "Agregar a Favoritos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   12240
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   0
      Width           =   1935
   End
   Begin FlexCell.Grid Grid4 
      Height          =   105
      Left            =   1305
      TabIndex        =   25
      Top             =   0
      Visible         =   0   'False
      Width           =   465
      _ExtentX        =   820
      _ExtentY        =   185
      Cols            =   5
      DefaultFontSize =   9.75
      Rows            =   30
   End
   Begin FlexCell.Grid Grid3 
      Height          =   240
      Left            =   0
      TabIndex        =   24
      Top             =   0
      Visible         =   0   'False
      Width           =   135
      _ExtentX        =   238
      _ExtentY        =   423
      BackColor1      =   12648447
      BackColor2      =   12648447
      BackColorActiveCellSel=   16777088
      BackColorBkg    =   16761024
      BackColorFixedSel=   16761024
      BackColorScrollBar=   16744576
      BorderColor     =   16744576
      CellBorderColor =   16744576
      CellBorderColorFixed=   16744576
      SelectionBorderColor=   16744576
      Cols            =   5
      DefaultFontName =   "Arial"
      DefaultFontSize =   8.25
      DisplayRowIndex =   -1  'True
      ForeColorFixed  =   8388608
      GridColor       =   16744576
      Rows            =   30
      DateFormat      =   2
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FF8080&
      Caption         =   "GRABAR"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   11160
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   9360
      Width           =   2310
   End
   Begin VB.TextBox PIVOTE4 
      Height          =   285
      Left            =   9600
      MaxLength       =   9
      TabIndex        =   15
      Text            =   "Text1"
      Top             =   9600
      Visible         =   0   'False
      Width           =   615
   End
   Begin XPFrame.FrameXp impuestos 
      Height          =   375
      Left            =   14280
      TabIndex        =   6
      Top             =   6840
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   661
      BackColor       =   49344
      Caption         =   "IMPUESTOS"
      CaptionEstilo3D =   1
      BackColor       =   49344
      ForeColor       =   8438015
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin FlexCell.Grid Grid2 
         Height          =   3975
         Left            =   0
         TabIndex        =   7
         Top             =   240
         Width           =   6735
         _ExtentX        =   11880
         _ExtentY        =   7011
         Cols            =   5
         DefaultFontName =   "Arial"
         DefaultFontSize =   8.25
         ExtendLastCol   =   -1  'True
         Rows            =   30
      End
   End
   Begin VB.PictureBox MANUAL 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF8080&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   135
      Left            =   0
      ScaleHeight     =   135
      ScaleWidth      =   615
      TabIndex        =   5
      Top             =   0
      Width           =   615
   End
   Begin VB.TextBox LINEAS 
      Height          =   285
      Left            =   9600
      MaxLength       =   3
      TabIndex        =   4
      Top             =   9000
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox pivote2 
      Height          =   285
      Left            =   9600
      MaxLength       =   10
      TabIndex        =   3
      Top             =   9000
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox PIVOTE 
      Height          =   285
      Left            =   9480
      MaxLength       =   8
      TabIndex        =   1
      Top             =   9240
      Visible         =   0   'False
      Width           =   255
   End
   Begin XPFrame.FrameXp FrameXp3 
      Height          =   30
      Left            =   6960
      TabIndex        =   8
      Top             =   8640
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   53
      BackColor       =   16761024
      Caption         =   "VALORES DEL COMPROBANTE"
      CaptionEstilo3D =   1
      BackColor       =   16761024
      ForeColor       =   8438015
      ColorBarraAbajo =   16711680
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   1
      Begin XPFrame.FrameXp FrameXp4 
         Height          =   615
         Left            =   480
         TabIndex        =   9
         Top             =   360
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   1085
         BackColor       =   16761024
         Caption         =   "INGRESOS"
         CaptionEstilo3D =   1
         BackColor       =   16761024
         ForeColor       =   8438015
         ColorBarraAbajo =   16711680
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Begin VB.Label debe 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFDF2&
            BackStyle       =   0  'Transparent
            Caption         =   " "
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   255
            Left            =   0
            TabIndex        =   10
            Top             =   240
            Width           =   1575
         End
      End
      Begin XPFrame.FrameXp FrameXp5 
         Height          =   615
         Left            =   2760
         TabIndex        =   11
         Top             =   360
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   1085
         BackColor       =   16761024
         Caption         =   "TOTAL"
         CaptionEstilo3D =   1
         BackColor       =   16761024
         ForeColor       =   8438015
         ColorBarraAbajo =   16711680
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Begin VB.Label haber 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFDF2&
            BackStyle       =   0  'Transparent
            Caption         =   " "
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   255
            Left            =   90
            TabIndex        =   12
            Top             =   240
            Width           =   1695
         End
      End
      Begin XPFrame.FrameXp FrameXp7 
         Height          =   615
         Left            =   4920
         TabIndex        =   13
         Top             =   360
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   1085
         BackColor       =   16761024
         Caption         =   "SALDO"
         CaptionEstilo3D =   1
         BackColor       =   16761024
         ForeColor       =   8438015
         ColorBarraAbajo =   16711680
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Begin VB.Label saldo 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFDF2&
            BackStyle       =   0  'Transparent
            Caption         =   " "
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   255
            Left            =   240
            TabIndex        =   14
            Top             =   240
            Width           =   1695
         End
      End
   End
   Begin XPFrame.FrameXp detalle 
      Height          =   4335
      Left            =   120
      TabIndex        =   16
      Top             =   4320
      Width           =   14055
      _ExtentX        =   24791
      _ExtentY        =   7646
      BackColor       =   16744576
      Caption         =   "Comprobante Glosa"
      CaptionEstilo3D =   1
      BackColor       =   16744576
      ForeColor       =   8438015
      ColorBarraAbajo =   16711680
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.TextBox txtitemfactura 
         Height          =   375
         Left            =   120
         MaxLength       =   40
         TabIndex        =   27
         Top             =   240
         Width           =   13815
      End
      Begin FlexCell.Grid grid1 
         Height          =   2775
         Left            =   45
         TabIndex        =   17
         Top             =   1440
         Width           =   13950
         _ExtentX        =   24606
         _ExtentY        =   4895
         AllowUserResizing=   0   'False
         BackColor1      =   12648447
         BackColor2      =   12648447
         BackColorActiveCellSel=   16777088
         BackColorBkg    =   16744576
         BackColorFixedSel=   16761024
         BackColorScrollBar=   16744576
         BorderColor     =   16744576
         CellBorderColor =   16744576
         CellBorderColorFixed=   16744576
         SelectionBorderColor=   16744576
         Cols            =   5
         DefaultFontName =   "Arial"
         DefaultFontSize =   8.25
         DisplayRowIndex =   -1  'True
         ForeColorFixed  =   8388608
         GridColor       =   16744576
         Rows            =   30
         DateFormat      =   2
      End
      Begin XPFrame.FrameXp frmexento 
         Height          =   675
         Left            =   0
         TabIndex        =   62
         Top             =   720
         Visible         =   0   'False
         Width           =   13935
         _ExtentX        =   24580
         _ExtentY        =   1191
         BackColor       =   16761024
         Caption         =   "GLOSA EXENTO"
         CaptionEstilo3D =   1
         BackColor       =   16761024
         ForeColor       =   8438015
         ColorBarraArriba=   16744576
         ColorBarraAbajo =   16711680
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Begin VB.TextBox txtitemfacturaEXENTO 
            Height          =   375
            Left            =   0
            MaxLength       =   40
            TabIndex        =   63
            Top             =   240
            Width           =   13935
         End
      End
   End
   Begin XPFrame.FrameXp CRCC 
      Height          =   645
      Left            =   135
      TabIndex        =   18
      Top             =   3600
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   1138
      BackColor       =   16761024
      Caption         =   "Nombre del Centro de Costo"
      CaptionEstilo3D =   1
      BackColor       =   16761024
      ForeColor       =   8438015
      ColorBarraArriba=   16744576
      ColorBarraAbajo =   16711680
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.TextBox DATO21 
         BackColor       =   &H00E1FFFD&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   1305
         MaxLength       =   2
         TabIndex        =   20
         Tag             =   "codigo"
         Top             =   240
         Width           =   375
      End
      Begin VB.TextBox DATO22 
         BackColor       =   &H00E1FFFD&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   1800
         MaxLength       =   2
         TabIndex        =   19
         Top             =   240
         Width           =   375
      End
      Begin VB.Label nombrecrcc 
         BackColor       =   &H00FFC0C0&
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
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   2385
         TabIndex        =   22
         Top             =   270
         Width           =   3855
      End
      Begin VB.Label Label7 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Codigo Cuenta"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   240
         Width           =   1095
      End
   End
   Begin XPFrame.FrameXp cabeza 
      Height          =   3435
      Left            =   120
      TabIndex        =   28
      Top             =   120
      Width           =   14055
      _ExtentX        =   24791
      _ExtentY        =   6059
      BackColor       =   16744576
      Caption         =   "Datos Documentos"
      CaptionEstilo3D =   1
      BackColor       =   16744576
      ForeColor       =   8438015
      ColorBarraAbajo =   16711680
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.OptionButton Option2 
         BackColor       =   &H00FF8080&
         Caption         =   "Publicidad"
         Height          =   375
         Left            =   10440
         TabIndex        =   61
         Top             =   720
         Value           =   -1  'True
         Width           =   3255
      End
      Begin VB.TextBox txtfolio 
         Alignment       =   1  'Right Justify
         BackColor       =   &H0080FF80&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   8000
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   59
         Top             =   240
         Width           =   1500
      End
      Begin VB.TextBox dato9 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E1FFFD&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   1680
         MaxLength       =   9
         TabIndex        =   48
         Tag             =   "rut"
         Top             =   945
         Width           =   1095
      End
      Begin VB.TextBox dato2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E1FFFD&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   1665
         MaxLength       =   10
         TabIndex        =   0
         Tag             =   "numero"
         Top             =   600
         Width           =   1455
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00FF8080&
         Caption         =   "Ventas"
         Height          =   375
         Left            =   10440
         TabIndex        =   30
         Top             =   240
         Width           =   3255
      End
      Begin VB.ComboBox COMBOTIPO 
         Height          =   315
         Left            =   1680
         Style           =   2  'Dropdown List
         TabIndex        =   29
         Top             =   240
         Width           =   3015
      End
      Begin XPFrame.FrameXp FrameXp8 
         Height          =   975
         Left            =   90
         TabIndex        =   31
         Top             =   2385
         Width           =   13695
         _ExtentX        =   24156
         _ExtentY        =   1720
         BackColor       =   16761024
         Caption         =   "Valores Documentos"
         CaptionEstilo3D =   1
         BackColor       =   16761024
         ForeColor       =   8438015
         ColorBarraAbajo =   16711680
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Begin VB.CheckBox chk1 
            BackColor       =   &H00C0FFC0&
            Caption         =   "OC"
            Height          =   195
            Left            =   9360
            TabIndex        =   82
            Top             =   360
            Width           =   735
         End
         Begin XPFrame.FrameXp frmorden 
            Height          =   975
            Left            =   10200
            TabIndex        =   77
            Top             =   0
            Visible         =   0   'False
            Width           =   3495
            _ExtentX        =   6165
            _ExtentY        =   1720
            BackColor       =   12648384
            Caption         =   "DATOS ORDEN DE COMPRA"
            CaptionEstilo3D =   1
            BackColor       =   12648384
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Alignment       =   1
            Begin VB.TextBox dato13b 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E1FFFD&
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   285
               Left            =   120
               MaxLength       =   10
               TabIndex        =   79
               Tag             =   "exento"
               Text            =   "0"
               Top             =   600
               Width           =   1695
            End
            Begin MSComCtl2.DTPicker fechaoc 
               Height          =   285
               Left            =   2040
               TabIndex        =   78
               Top             =   600
               Width           =   1335
               _ExtentX        =   2355
               _ExtentY        =   503
               _Version        =   393216
               Format          =   160038913
               CurrentDate     =   42110
            End
            Begin VB.Label Label9 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               BorderStyle     =   1  'Fixed Single
               Caption         =   " OC PROVEEDOR"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00800000&
               Height          =   255
               Left            =   120
               TabIndex        =   81
               Top             =   360
               Width           =   1695
            End
            Begin VB.Label Label10 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               BorderStyle     =   1  'Fixed Single
               Caption         =   " FECHA OC"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00800000&
               Height          =   255
               Left            =   2040
               TabIndex        =   80
               Top             =   360
               Width           =   1335
            End
         End
         Begin VB.TextBox dato5 
            BackColor       =   &H00E1FFFD&
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   840
            MaxLength       =   4
            TabIndex        =   41
            Tag             =   "fecha"
            Top             =   480
            Width           =   615
         End
         Begin VB.TextBox dato4 
            BackColor       =   &H00E1FFFD&
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   480
            MaxLength       =   2
            TabIndex        =   40
            Tag             =   "fecha"
            Top             =   480
            Width           =   375
         End
         Begin VB.TextBox dato3 
            BackColor       =   &H00E1FFFD&
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   120
            MaxLength       =   2
            TabIndex        =   39
            Tag             =   "fecha"
            Top             =   480
            Width           =   375
         End
         Begin VB.TextBox dato6 
            BackColor       =   &H00E1FFFD&
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   1680
            MaxLength       =   2
            TabIndex        =   38
            Tag             =   "fechavencimiento"
            Top             =   480
            Width           =   375
         End
         Begin VB.TextBox dato7 
            BackColor       =   &H00E1FFFD&
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   2040
            MaxLength       =   2
            TabIndex        =   37
            Tag             =   "fecha"
            Top             =   480
            Width           =   375
         End
         Begin VB.TextBox dato8 
            BackColor       =   &H00E1FFFD&
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   2400
            MaxLength       =   4
            TabIndex        =   36
            Tag             =   "fecha"
            Top             =   480
            Width           =   615
         End
         Begin VB.TextBox dato11 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E1FFFD&
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   3105
            MaxLength       =   15
            TabIndex        =   35
            Tag             =   "neto"
            Text            =   "0"
            Top             =   480
            Width           =   1455
         End
         Begin VB.TextBox dato12 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E1FFFD&
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   4680
            MaxLength       =   15
            TabIndex        =   34
            Tag             =   "iva"
            Text            =   "0"
            Top             =   480
            Width           =   1455
         End
         Begin VB.TextBox dato13 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E1FFFD&
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   6210
            MaxLength       =   15
            TabIndex        =   33
            Tag             =   "exento"
            Text            =   "0"
            Top             =   480
            Width           =   1455
         End
         Begin VB.TextBox total 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E1FFFD&
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   7800
            MaxLength       =   15
            TabIndex        =   32
            Tag             =   "monto"
            Text            =   "0"
            Top             =   480
            Width           =   1455
         End
         Begin VB.Label Label6 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   " FECHA"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   255
            Left            =   120
            TabIndex        =   47
            Top             =   240
            Width           =   1335
         End
         Begin VB.Label Label5 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "VENCIMIENTO"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   255
            Left            =   1680
            TabIndex        =   46
            Top             =   240
            Width           =   1335
         End
         Begin VB.Label Label11 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   " NETO"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   255
            Left            =   3120
            TabIndex        =   45
            Top             =   240
            Width           =   1455
         End
         Begin VB.Label Label17 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   " I.V.A"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   255
            Left            =   4680
            TabIndex        =   44
            Top             =   240
            Width           =   1455
         End
         Begin VB.Label Label26 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   " EXENTO"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   255
            Left            =   6210
            TabIndex        =   43
            Top             =   240
            Width           =   1455
         End
         Begin VB.Label Label28 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   " TOTAL"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   255
            Left            =   7785
            TabIndex        =   42
            Top             =   240
            Width           =   1455
         End
      End
      Begin VB.Label lbl_folio 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   " FOLIO SII"
         Height          =   315
         Left            =   6720
         TabIndex        =   60
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label DV 
         BackColor       =   &H00DAF9FE&
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
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   2880
         TabIndex        =   58
         Top             =   960
         Width           =   255
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   " TIPO"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   120
         TabIndex        =   57
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   " NUMERO"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   120
         TabIndex        =   56
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label Label16 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   " CLIENTE"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   120
         TabIndex        =   55
         Top             =   960
         Width           =   1455
      End
      Begin VB.Label nombreproveedor 
         BackColor       =   &H0080FFFF&
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
         Height          =   255
         Left            =   3240
         TabIndex        =   54
         Top             =   960
         Width           =   6255
      End
      Begin VB.Label tipodocumento 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   2040
         TabIndex        =   53
         Top             =   240
         Width           =   3135
      End
      Begin VB.Label lbldireccion 
         BackColor       =   &H0080FFFF&
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
         Height          =   255
         Left            =   135
         TabIndex        =   52
         Top             =   1305
         Width           =   9360
      End
      Begin VB.Label lblcomuna 
         BackColor       =   &H0080FFFF&
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
         Height          =   255
         Left            =   135
         TabIndex        =   51
         Top             =   1665
         Width           =   4500
      End
      Begin VB.Label lblciudad 
         BackColor       =   &H0080FFFF&
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
         Height          =   255
         Left            =   4770
         TabIndex        =   50
         Top             =   1665
         Width           =   4725
      End
      Begin VB.Label lblgiro 
         BackColor       =   &H0080FFFF&
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
         Height          =   255
         Left            =   135
         TabIndex        =   49
         Top             =   2025
         Width           =   9360
      End
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash opciones 
      Height          =   1455
      Left            =   120
      TabIndex        =   2
      Top             =   8640
      Width           =   6855
      _cx             =   12091
      _cy             =   2566
      FlashVars       =   ""
      Movie           =   "c:\barra_opciones.swf"
      Src             =   "c:\barra_opciones.swf"
      WMode           =   "Transparent"
      Play            =   "0"
      Loop            =   "-1"
      Quality         =   "High"
      SAlign          =   ""
      Menu            =   "-1"
      Base            =   ""
      AllowScriptAccess=   "always"
      Scale           =   "ShowAll"
      DeviceFont      =   "-1"
      EmbedMovie      =   "0"
      BGColor         =   ""
      SWRemote        =   ""
      MovieData       =   ""
      SeamlessTabbing =   "1"
      Profile         =   "0"
      ProfileAddress  =   ""
      ProfilePort     =   "0"
      AllowNetworking =   "all"
      AllowFullScreen =   "false"
   End
End
Attribute VB_Name = "publi0009"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
     Private GRABACON As Boolean
     Private publicidad As Boolean
     Private empresarelacionada As Boolean
     
     Private lin As Double
     Private tipocuenta As String
     Private cc As Integer
     Private FORMATOGRILLA(100, 20)
     Private formatogrilla2(100, 20)
     Private cdi As Integer
     Private CANDO As Integer
     Private existe As String
     Private MODIFI As Integer
     Private canli As Double
     Private AUXILIAR(1000, 3) As String
     Private respu As String
     Private tipoctacte As String
     Private nlineas As Double
     Private DOCU(10) As String
     Private grilladetalle(1000, 14) As String
     Private SALDOPE As Double
     Private NETO As Double
     Private CUENTAMAYOR(999) As String
     Private TIENECTACTE(999) As String
     Private TIENECRCC(999) As String
     Private TIENEBANCO(999) As String
     Private TIENEILA(999) As String
     Private TIENEICA(999) As String
     Private TIENEIHA(999) As String
     Private TIENEACTIVO(999) As String
     
     
 

 
Private Sub chk1_Click()
    If chk1.Value = 1 Then
        frmorden.Visible = True
    Else
        frmorden.Visible = False
    End If
End Sub

Private Sub CmdFavoritos_Click()
Call AgregaFavoritos(USUARIOSISTEMA, App.EXEName, Me.Name, Me.Caption)
End Sub

Private Sub Command1_Click()
    grabafactura
    dato9.SetFocus
End Sub
 
 
 



Private Sub dato13b_GotFocus()
Call cargatexto(dato13b)
End Sub

Private Sub dato13b_KeyDown(KeyCode As Integer, Shift As Integer)
Call flechas(dato13, dato13b, KeyCode)
End Sub

Private Sub dato13b_KeyPress(KeyAscii As Integer)
snum = 0: KeyAscii = esNumero(KeyAscii)
If KeyAscii = 13 Then
Call ceros(dato13b)
fechaoc.Enabled = True
fechaoc.SetFocus
End If
End Sub

Private Sub dato2_GotFocus()


'If Val(dato1.text) < 1 Or Val(dato1.text) > CANDO Then dato1.text = "": dato1.SetFocus:
'tipodocumento.Caption = GRILLATIPO.TextMatrix(Val(dato1.text) - 1, 1)
dato2.text = LeerUltimo(Val(Mid(COMBOTIPO.text, 1, 3)))
Call cargatexto(dato2)
End Sub


Private Sub dato3_Change()
If Val(dato3.text) > 31 Then dato3.text = ""
End Sub
Private Sub dato4_Change()
If Val(dato4.text) > 12 Or Val(dato4.text) < 1 Then dato4.text = ""
End Sub

Private Sub dato4_LostFocus()
Call ceros(dato4)

End Sub

Private Sub dato5_LostFocus()
If dato5.text < "1900" Or dato5.text <> Format(fechasistema, "YYYY") Then dato5.text = ""

End Sub

Private Sub dato8_LostFocus()
If dato8.text < dato5.text Then dato8.text = ""

End Sub

Private Sub dato6_Change()
If Val(dato6.text) > 31 Then dato6.text = ""
End Sub
Private Sub dato7_Change()
If Val(dato7.text) > 12 Or Val(dato7.text) < 1 Then dato7.text = ""
End Sub


Private Sub dato3_GotFocus()
If Val(dato9.text) = 0 Then dato9.SetFocus: GoTo no

If tipocuenta <> "00" Then
    DV.Caption = rut(dato9.text)
    pivote2.text = dato9.text + DV.Caption
    
    leectacte
'    leectacte_FINAL
        If Val(dato9.text) = 0 Then dato9.SetFocus: GoTo no:
    If cierrect = "-" Then cierrect = "": dato9.SetFocus: GoTo no:
End If


Rem If tipocuenta <> "00" And Val(dato9.text) = 0 Then dato9.SetFocus: GoTo no:
no:
End Sub
'Sub leectacte(rut)
'    campos(0, 0) = "rut"
'    campos(1, 0) = "nombre"
'    campos(2, 0) = ""
'    campos(0, 2) = "cuentascorrientes"
'    condicion = "tipo=" + "'" + cuentacliente + "' and rut=" + "'" + rut + "' and año='" + año + "'"
'    op = 5
'    sqlconta.response = campos
'    Set sqlconta.conexion = contadb
'    Call sqlconta.sqlconta(op, condicion)
'    If sqlconta.status = 4 Then
'    Call crearcuentacorriente(rut)
'    End If
'
'End Sub

Sub leectacte()
Dim cuentaclientes As String
    If Val(Mid(COMBOTIPO.text, 1, 3)) = 46 Then ' FACTURA DE COMPRA
           cuentaclientes = leerdatos(conta, "maestroempresas", "cuentaproveedor", "codigoempresa='" + empresaactiva + "' ")
        Else
           cuentaclientes = leerdatos(conta, "maestroempresas", "cuentaclientes", "codigoempresa='" + empresaactiva + "' ")
        End If
        


    campos(0, 0) = "rut"
    campos(1, 0) = "nombre"
    campos(2, 0) = "direccion"
    campos(3, 0) = "comuna"
    campos(4, 0) = "ciudad"
    campos(5, 0) = "giro"
    campos(6, 0) = ""
    
    
   campos(0, 2) = "cuentascorrientes"
    condicion = "tipo=" + "'" + cuentaclientes + "' and rut=" + "'" + pivote2.text + "' and año='" + año + "'"
    op = 5
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)
    If sqlconta.status = 4 Then
                maestro02.dato1.Enabled = True
                maestro02.dato2.Enabled = True
                maestro02.DV.Caption = True
                maestro02.dato1.text = cuentaclientes
                maestro02.dato2.text = dato9.text
                maestro02.DV.Caption = DV.Caption
                cierrect = "S"
                maestro02.Show
                
                GoTo no:
    
    End If
    
    nombreproveedor.Caption = sqlconta.response(1, 3)
    lbldireccion.Caption = sqlconta.response(2, 3)
    lblcomuna.Caption = sqlconta.response(3, 3)
    lblciudad.Caption = sqlconta.response(4, 3)
    lblgiro.Caption = sqlconta.response(5, 3)
    
     Call crearcliente(dato9.text & DV.Caption, "0", nombreproveedor.Caption, lbldireccion.Caption, lblcomuna.Caption, lblciudad.Caption, lblgiro.Caption)
     
    dato3.Enabled = True
    dato3.SetFocus
no:

End Sub

Sub crearcliente(RUTCLIENTE, suc, NOMBRE, direccion, comuna, ciudad, giro)
        Dim csql As New rdoQuery
        Dim resultados As rdoResultset
        Set csql.ActiveConnection = contadb
    
            csql.sql = "replace INTO " & clientesistema & "ventas.sv_maestroclientes   "
            csql.sql = csql.sql & "(rut,sucursal,nombre,direccion,comuna,ciudad,giro) "
            csql.sql = csql.sql & "value ('" + RUTCLIENTE + "','" + suc + "','" & NOMBRE & "','" & direccion & "','" & comuna & "','" & ciudad & "','" & giro & "') "
            csql.Execute
            
            csql.Close
            Set csql = Nothing
            
End Sub



Sub crearcuentacorriente(rut)
    Dim resultados As rdoResultset
    Dim csql As New rdoQuery
    
        Set csql.ActiveConnection = gestion

            csql.sql = "INSERT INTO " + clientesistema + "conta" + empresaactiva + ".cuentascorrientes "
            csql.sql = csql.sql & "(año,tipo,rut,nombre,direccion,comuna,ciudad,giro,fono) "
            csql.sql = csql.sql & "SELECT '" + año + "','" + cuentacliente + "',mc.rut,mc.nombre,mc.direccion,mc.comuna,mc.ciudad,mc.giro,mc.fono1 "
            csql.sql = csql.sql & "FROM " & clientesistema & "ventas.sv_maestroclientes as mc "
            csql.sql = csql.sql & "WHERE mc.rut = '" & rut & "' AND mc.sucursal ='0'"
            
            csql.Execute
            Call sincronizadatos(csql.sql, gestion, "")
            
            
            csql.sql = "INSERT INTO " + clientesistema + "conta" + empresaactiva + ".saldosctacte "
            csql.sql = csql.sql & "(año,tipo,rut) "
            csql.sql = csql.sql & "SELECT '" + año + "','" + cuentacliente + "',mc.rut "
            csql.sql = csql.sql & "FROM " & clientesistema & "ventas.sv_maestroclientes as mc "
            csql.sql = csql.sql & "WHERE mc.rut = '" & rut & "' AND mc.sucursal ='0'"
            
            csql.Execute
            Call sincronizadatos(csql.sql, gestion, "")
            


End Sub



Private Sub dato4_GotFocus()
If dato3.text = "00" Then dato4.Enabled = True: dato5.Enabled = True: dato6.Enabled = True: dato3.text = Mid(fechasistema, 1, 2): dato4.text = Mid(fechasistema, 4, 2): dato5.text = Mid(fechasistema, 7, 4): dato6.SetFocus
Call cargatexto(dato4)
End Sub

Private Sub dato5_GotFocus()
Call cargatexto(dato5)
End Sub




Private Sub dato6_GotFocus()

Call cargatexto(dato6)
If IsDate(dato3.text + "-" + dato4.text + "-" + dato5.text) = False Then dato3.text = "": dato4.text = "": dato5.text = "": dato3.SetFocus

End Sub


Private Sub dato7_GotFocus()
If dato6.text = "00" Then dato6.Enabled = True: dato7.Enabled = True: dato8.Enabled = True: dato6.text = Mid(fechasistema, 1, 2): dato7.text = Mid(fechasistema, 4, 2): dato8.text = Mid(fechasistema, 7, 4): dato11.Enabled = True: dato11.SetFocus
Call cargatexto(dato7)
End Sub
Private Sub dato8_GotFocus()

Call cargatexto(dato8)
End Sub
Private Sub dato9_GotFocus()
Dim tipo As String
    Dim tIPOrEF As String
    Dim fecharef As String
    Dim folioref As String
    Dim localref As String
leefactura
If Mid(COMBOTIPO.text, 1, 3) = "033" Then
    dato13b.text = LeedatoOCproveedor("FV", dato2.text, CONFI_EMPRESAFAE, "ordencompra", "99")
    fechaoc.Value = LeedatoOCproveedor("FV", dato2.text, CONFI_EMPRESAFAE, "ifnull(fechaoc,'0000-00-00')", "99")
    If dato13b.text <> "0000000000" Then frmorden.Visible = True
End If

If sqlconta.status = 0 Then
   carga
   
   If Val(Mid(COMBOTIPO.text, 1, 3)) = 33 Then ' factura
        tipo = 6
    End If
    
    If Val(Mid(COMBOTIPO.text, 1, 3)) = 34 Then ' factura extenta
        tipo = 0
    End If
    
    If Val(Mid(COMBOTIPO.text, 1, 3)) = 56 Then ' nota debito
        tipo = 7
        
        frmdatos.Visible = True
        Option1.Value = True
        tIPOrEF = leerdatos(conta, clientesistema & "ventas" & CONFI_EMPRESAFAE & ".sv_otros_documento_detalle_" & CONFI_EMPRESAFAE, "tipodocumento", "tipo='ND' and numero=lpad('" & dato2.text & "',10,0) and fecha='" & dato5.text + "-" + dato4.text + "-" + dato3.text & "' and caja='99' and tipodocumento<>'' ")
        fecharef = leerdatos(conta, clientesistema & "ventas" & CONFI_EMPRESAFAE & ".sv_otros_documento_detalle_" & CONFI_EMPRESAFAE, "fechadocumento", "tipo='ND' and numero=lpad('" & dato2.text & "',10,0) and fecha='" & dato5.text + "-" + dato4.text + "-" + dato3.text & "' and caja='99' and tipodocumento<>'' ")
        folioref = leerdatos(conta, clientesistema & "ventas" & CONFI_EMPRESAFAE & ".sv_otros_documento_detalle_" & CONFI_EMPRESAFAE, "numerodocumento", "tipo='ND' and numero=lpad('" & dato2.text & "',10,0) and fecha='" & dato5.text + "-" + dato4.text + "-" + dato3.text & "' and caja='99' and tipodocumento<>'' ")
        localref = empresaactiva
        If IsDate(fecharef) = True Then
            dato15.text = Format(fecharef, "dd")
            dato16.text = Format(fecharef, "mm")
            dato17.text = Format(fecharef, "yyyy")
        End If
        dato14 = tIPOrEF
        
        dato18 = folioref
    End If
    
    If Val(Mid(COMBOTIPO.text, 1, 3)) = 61 Then ' nota credito
         tipo = 8
        frmdatos.Visible = True
'        Option1.Value = True
        tIPOrEF = leerdatos(conta, clientesistema & "ventas" & CONFI_EMPRESAFAE & ".sv_otros_documento_detalle_" & CONFI_EMPRESAFAE, "tipodocumento", "tipo='NF' and numero=lpad('" & dato2.text & "',10,0) and fecha='" & dato5.text + "-" + dato4.text + "-" + dato3.text & "' and caja='99' and tipodocumento<>'' ")
        fecharef = leerdatos(conta, clientesistema & "ventas" & CONFI_EMPRESAFAE & ".sv_otros_documento_detalle_" & CONFI_EMPRESAFAE, "fechadocumento", "tipo='NF' and numero=lpad('" & dato2.text & "',10,0) and fecha='" & dato5.text + "-" + dato4.text + "-" + dato3.text & "' and caja='99' and tipodocumento<>'' ")
        folioref = leerdatos(conta, clientesistema & "ventas" & CONFI_EMPRESAFAE & ".sv_otros_documento_detalle_" & CONFI_EMPRESAFAE, "numerodocumento", "tipo='NF' and numero=lpad('" & dato2.text & "',10,0) and fecha='" & dato5.text + "-" + dato4.text + "-" + dato3.text & "' and caja='99' and tipodocumento<>'' ")
        localref = empresaactiva ' leerdatos(conta, clientesistema & "ventas" & CONFI_EMPRESAFAE & ".sv_otros_documento_detalle_" & CONFI_EMPRESAFAE, "localdocumento", "tipo='NF' and numero=lpad('" & dato2.text & "',10,0) and fecha='" & dato5.text + "-" + dato4.text + "-" + dato3.text & "' and caja='99' and tipodocumento<>'' ")
        


'        If tiporef = "FC" Then
''            LBLRETENCION.Caption = 0
'            dato12.text = Round(dato11.text * 1.19) - dato11
'            LBLRETENCION.Caption = dato12
'        End If
        
        If IsDate(fecharef) = True Then
            dato15.text = Format(fecharef, "dd")
            dato16.text = Format(fecharef, "mm")
            dato17.text = Format(fecharef, "yyyy")
        End If
        dato14 = tIPOrEF
        
        dato18 = folioref
    End If
    If Val(Mid(COMBOTIPO.text, 1, 3)) = 46 Then ' nota debito
        tipo = 7
    End If
 
   Call leecomprobante(tipo)

      If nlineas <> 0 And txtfolio.text <> "NO FISCAL" Then
      opciones.Visible = True
      opciones.SetFocus
      detalle.Enabled = False
           
           
           GoTo no:
        
      End If
'      If txtfolio.text = "NO FISCAL" Then
        opciones.Visible = True: grid1.Enabled = False: opciones.SetFocus
'      End If
End If

If Val(dato2.text) = 0 Then dato2.text = "": dato2.Enabled = True: dato2.SetFocus
Call cargatexto(dato3)
no:

Call cargatexto(dato9)





End Sub

Private Sub dato11_GotFocus()
Rem If IsDate(dato6.text + "-" + dato7.text + "-" + dato8.text) = False Then dato6.text = "": dato7.text = "": dato8.text = "": dato6.SetFocus



Call cargatexto(dato11)

no:
End Sub

Private Sub dato12_GotFocus()

If Mid(COMBOTIPO.text, 1, 3) <> "034" Then
    sumador = Round(CDbl(Replace(dato11.text, ",", "")) * iva / 100)
    dato12.text = Format(sumador, "#,###,###,##0")
Else
    dato12.text = Format(0, "#,###,###,##0")
End If
    totalfactura
    Call cargatexto(dato12)
End Sub
Private Sub dato13_GotFocus()
totalfactura
Call cargatexto(dato13)
End Sub
 

Private Sub fechaoc_Click()
DATO21.SetFocus
End Sub

Private Sub fechaoc_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
DATO21.SetFocus
End If
End Sub

Private Sub Form_Activate()
leeCUENTA
End Sub
Sub leeCUENTA()

    campos(0, 0) = "codigo"
    campos(1, 0) = ""
    
    campos(0, 2) = "cuentasdelmayor"
    condicion = "codigo=" + "'" + cuentacliente + "' and año='" + Format(fechasistema, "yyyy") + "' order by codigo"
    
    op = 5
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)
    If sqlconta.status = 4 Then
    MsgBox ("CUENTA CLIENTE NUMERO " + cuentacliente + " NO EXISTE EN MAESTRO DEL MAYOR CONFIGURE SISTEMA EN MAESTRO EMPRESAS ")
   Unload Me
    
    
    End If
    
   
    
End Sub

Private Sub Form_Load()
CENTRAR Me
iva = 19
    Call Conectar_BD
    Call Conectarconta(Servidor, clientesistema + "conta", Usuario, password)
    sc = 0
    opciones.Visible = False
'GRILLATIPOS
cargatipos
Call CARGAGRILLA2(2, 12)

Call CARGAGRILLA(20, 2)
Call CARGAGRILLAexento

DOCU(0) = "EX "
DOCU(1) = "FA "
DOCU(2) = "ND "
DOCU(3) = "NB "
DOCU(4) = "NF "
DOCU(5) = "FE"
DOCU(6) = "FAE "
DOCU(7) = "NDE "
DOCU(8) = "NCE "
DOCU(9) = "FX "

impuestos.Visible = False



End Sub
'Sub GRILLATIPOS()
''GRILLATIPO.Cols = 2
''GRILLATIPO.Rows = 1
''GRILLATIPO.ColWidth(0) = 200 * 2
''GRILLATIPO.ColWidth(1) = 200 * 10
''
''GRILLATIPO.TextMatrix(0, 0) = "1"
''
''GRILLATIPO.TextMatrix(0, 1) = "FACTURA"
''
''CANDO = 1
'
'
'GRILLATIPO.Cols = 2
'GRILLATIPO.Rows = 5
'GRILLATIPO.ColWidth(0) = 200 * 2
'GRILLATIPO.ColWidth(1) = 200 * 10
'
'GRILLATIPO.TextMatrix(0, 0) = "1"
'GRILLATIPO.TextMatrix(1, 0) = "2"
'GRILLATIPO.TextMatrix(2, 0) = "3"
'GRILLATIPO.TextMatrix(3, 0) = "4"
'GRILLATIPO.TextMatrix(4, 0) = "5"
'
'GRILLATIPO.TextMatrix(0, 1) = "FACTURA"
'GRILLATIPO.TextMatrix(1, 1) = "NOTA DE DEBITO"
'GRILLATIPO.TextMatrix(2, 1) = "NC. BOLETA    "
'GRILLATIPO.TextMatrix(3, 1) = "NC. FACTURA   "
'GRILLATIPO.TextMatrix(4, 1) = "FAC.EXPORTACION"
'
'CANDO = 5
'
' GRILLATIPO.RowHeight(1) = 0
'GRILLATIPO.RowHeight(2) = 0
'GRILLATIPO.RowHeight(4) = 0
'
'
'
'End Sub

 
    Sub cargatipos()
        Dim csql As New rdoQuery
        Dim resultados As rdoResultset
        Set csql.ActiveConnection = contadb
        csql.sql = "select concat(lpad(tipo,3,0),' - ',nombre) "
        csql.sql = csql.sql & "from " & cliente_sql & "fae.sv_fae_tiposdocumentos "
        csql.sql = csql.sql & " where visible='1' order by tipo"
        csql.Execute
        COMBOTIPO.Clear
        If csql.RowsAffected > 0 Then
            Set resultados = csql.OpenResultset
            While Not resultados.EOF
                COMBOTIPO.AddItem (resultados(0))
                resultados.MoveNext
            Wend
        End If
      COMBOTIPO.ListIndex = 0
    End Sub


 
Private Sub dato2_KeyDown(KeyCode As Integer, Shift As Integer)
    Call flechas(dato2, dato9, KeyCode)
End Sub
 Private Sub dato9_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF2 Then Call ayudactacte2(dato3)
    Call flechas(dato2, dato3, KeyCode)
End Sub
Private Sub dato3_KeyDown(KeyCode As Integer, Shift As Integer)
    Call flechas(dato9, dato4, KeyCode)
End Sub
 
 Private Sub dato4_KeyDown(KeyCode As Integer, Shift As Integer)
    Call flechas(dato3, dato5, KeyCode)
End Sub
Private Sub dato5_KeyDown(KeyCode As Integer, Shift As Integer)
    Call flechas(dato4, dato6, KeyCode)
End Sub
Private Sub dato6_KeyDown(KeyCode As Integer, Shift As Integer)
        Call flechas(dato5, dato7, KeyCode)
End Sub
Private Sub dato7_KeyDown(KeyCode As Integer, Shift As Integer)
    Call flechas(dato6, dato8, KeyCode)
End Sub
Private Sub dato8_KeyDown(KeyCode As Integer, Shift As Integer)
    Call flechas(dato7, dato11, KeyCode)
End Sub
Private Sub dato11_KeyDown(KeyCode As Integer, Shift As Integer)
        'If KeyCode = vbKeyF2 Then Call ayudacrcc(dato12)
    Call flechas(dato8, dato12, KeyCode)
End Sub
Private Sub dato12_KeyDown(KeyCode As Integer, Shift As Integer)

    Call flechas(dato11, dato13, KeyCode)
End Sub
Private Sub dato13_KeyDown(KeyCode As Integer, Shift As Integer)
    Call flechas(dato12, dato13, KeyCode)
End Sub

 


 
Private Sub dato2_KeyPress(KeyAscii As Integer)
    snum = 0: KeyAscii = esNumero(KeyAscii)
    If KeyAscii = 13 Then Call ceros(dato2): Call Pregunta(dato2, dato9)

End Sub
Private Sub dato9_KeyPress(KeyAscii As Integer)
    snum = 0: KeyAscii = esNumero(KeyAscii)
    If KeyAscii = 13 Then Call ceros(dato9): Call Pregunta(dato9, dato3)
End Sub

Private Sub dato3_KeyPress(KeyAscii As Integer)
    snum = 0: KeyAscii = esNumero(KeyAscii)
    If KeyAscii = 13 Then Call ceros(dato3): Call Pregunta(dato3, dato4)
End Sub
Private Sub dato4_KeyPress(KeyAscii As Integer)
    snum = 0: KeyAscii = esNumero(KeyAscii)
    If KeyAscii = 13 Then Call ceros(dato4): Call Pregunta(dato4, dato5)
End Sub
Private Sub dato5_KeyPress(KeyAscii As Integer)
    snum = 0: KeyAscii = esNumero(KeyAscii)
    If KeyAscii = 13 Then Call ceros(dato5): Call Pregunta(dato5, dato6)
End Sub
Private Sub dato6_KeyPress(KeyAscii As Integer)
    ' If KeyAscii = 42 And SUMADEBE = SUMAHABER Then grabarcomprobante:retorno: dato3.Enabled = True: dato3.SetFocus
    snum = 0: KeyAscii = esNumero(KeyAscii)
    If KeyAscii = 13 Then Call ceros(dato6): Call Pregunta(dato6, dato7)
no:
End Sub
Private Sub dato7_KeyPress(KeyAscii As Integer)
    snum = 0: KeyAscii = esNumero(KeyAscii)
    If KeyAscii = 13 Then Call ceros(dato7): Call Pregunta(dato7, dato8)
End Sub
Private Sub dato8_KeyPress(KeyAscii As Integer)
    snum = 0: KeyAscii = esNumero(KeyAscii)
    If KeyAscii = 13 Then Call ceros(dato8): Call Pregunta(dato8, dato11)
End Sub
Private Sub dato11_KeyPress(KeyAscii As Integer)
    snum = 0: KeyAscii = esNumero(KeyAscii)
    If KeyAscii = 13 Then
        If Mid(COMBOTIPO.text, 1, 3) <> "034" Then
            If dato11.text <> "" And dato11.text <> "0" Then
                Call formato(dato11, 0)
                Call Pregunta(dato11, dato12)
            End If
        End If
        If Mid(COMBOTIPO.text, 1, 3) = "034" Then
            Call formato(dato11, 0)
            Call Pregunta(dato11, dato13)
        End If
        If Mid(COMBOTIPO.text, 1, 3) = "061" And dato11.text = "0" Then
            Call formato(dato11, 0)
            Call Pregunta(dato11, DATO21)
        End If
    End If
    
End Sub
Private Sub dato12_KeyPress(KeyAscii As Integer)
    snum = 0: KeyAscii = esNumero(KeyAscii)
    
    If KeyAscii = 13 And dato12.text <> "" Then
    If Val(dato12.text) <> 0 And Mid(COMBOTIPO.text, 1, 3) = "034" Then dato12.text = "0"
    Call formato(dato12, 0)
    Call Pregunta(dato12, dato13)
    End If
    
End Sub

Private Sub dato13_KeyPress(KeyAscii As Integer)
    snum = 0: KeyAscii = esNumero(KeyAscii)
    If KeyAscii = 13 Then
       detalle.Enabled = True
       If Mid(COMBOTIPO.text, 1, 3) = "033" Then
        If Val(Replace(dato11.text, ".", "")) <= 5 And Val(Replace(dato13.text, ".", "")) > 0 Then
            dato11.text = "5"
            dato12.text = "1"
        End If
        If frmorden.Visible = True Then
        dato13b.Enabled = True
        dato13b.SetFocus
        Else
        Call Pregunta(dato13, DATO21)
        End If
        Else
          Call Pregunta(dato13, DATO21)
       End If
        
    End If
End Sub

 
Sub carga()
    Dim k As Double
    Dim tipo As String
    
    disponible (True)
'    For k = 0 To COMBOTIPO.ListCount - 1
'        If sqlconta.response(0, 3) = Val(Mid(COMBOTIPO.text, 1, 3)) Then
'            COMBOTIPO.ListIndex = k
'            Exit For
'        End If
'    Next k
    
    dato2.text = sqlconta.response(1, 3)
    dato3.text = Mid(sqlconta.response(2, 3), 1, 2)
    dato4.text = Mid(sqlconta.response(2, 3), 4, 2)
    dato5.text = Mid(sqlconta.response(2, 3), 7, 4)
    dato6.text = Mid(sqlconta.response(3, 3), 1, 2)
    dato7.text = Mid(sqlconta.response(3, 3), 4, 2)
    dato8.text = Mid(sqlconta.response(3, 3), 7, 4)
    dato9.text = Mid(sqlconta.response(4, 3), 1, 9)
    DV.Caption = Mid(sqlconta.response(4, 3), 10, 1)
    dato11.text = Format(sqlconta.response(5, 3), "##,###,###,##0")
    dato12.text = Format(sqlconta.response(6, 3), "##,###,###,##0")
    dato13.text = Format(sqlconta.response(7, 3), "##,###,###,##0")
    
    total.text = Format(sqlconta.response(9, 3), "##,###,###,##0")
    
     If sqlconta.response(10, 3) = 1 Then
        Option1.Value = True
    End If
     If sqlconta.response(10, 3) = 2 Then
        Option2.Value = True
    End If
     
    txtitemfactura.text = sqlconta.response(11, 3)
    
    
    
   
    totalfactura
        DV.Caption = rut(dato9.text)
    pivote2.text = dato9.text + DV.Caption
    
    
     If Val(Mid(COMBOTIPO.text, 1, 3)) = 33 Then
        tipo = "FV"
    End If

  If Val(Mid(COMBOTIPO.text, 1, 3)) = 34 Then ' factura exenta
         tipo = "EX"
    End If
    
    If Val(Mid(COMBOTIPO.text, 1, 3)) = 56 Then ' nota debito
         tipo = "ND"
    End If
     If Val(Mid(COMBOTIPO.text, 1, 3)) = 46 Then ' FACTURA DE COMPRA
         tipo = "FC"
    End If
    
    
    If Val(Mid(COMBOTIPO.text, 1, 3)) = "61" Then
        tipo = "NF"
    End If
    
     If documentocreado(tipo, "99", CONFI_EMPRESAFAE, dato2.text, dato5.text & "-" & dato4.text & "-" & dato3.text) = True Then
            txtfolio.text = NUMERODOCUMENTO_DTE
            Else
            txtfolio.text = "NO FISCAL"
            End If
            
            
    
        leectacte_FINAL
        
fin:
End Sub


Sub flechas(ByRef caja As TextBox, ByRef caja1 As TextBox, ByRef codigo As Integer)
    If codigo = 38 And caja.Enabled = True Then caja.SetFocus: caja.SelStart = 0
    If codigo = 40 And caja1.Enabled = True Then caja1.SetFocus: caja.SelStart = 0
End Sub
Sub GRABADETALLEIMPUESTOS()
'    Dim j As Double
'    For j = 1 To Grid2.Rows - 2
'
'    If Val(Grid2.Cell(LINEAS, 3).text) <> 0 Then
'    campos(0, 0) = "tipo"
'    campos(1, 0) = "numero"
'    campos(2, 0) = "rut"
'    campos(3, 0) = "Cuenta"
'    campos(4, 0) = "Monto"
'    campos(5, 0) = ""
'    campos(0, 1) = dato1.text
'    campos(1, 1) = dato2.text
'    campos(2, 1) = dato9.text + DV.Caption
'    campos(3, 1) = Grid2.Cell(LINEAS, 1).text
'    campos(4, 1) = Grid1.Cell(LINEAS, 3).text
'    campos(5, 1) = ""
'
'    campos(0, 2) = "facturasdeventas_impuestos"
'
'    sqlconta.response = campos
'    Set sqlconta.conexion = contadb
'    Call sqlconta.sqlconta(op, condicion)
'    End If
'
'    Next j
End Sub

Sub grabafactura()
    Dim netos As Double
    Dim DH As String
    Dim tipofactura As Double
    Dim tipo As String
    Dim loc As String
    Dim tipoconta As String
    
    If Mid(COMBOTIPO.text, 1, 3) = "033" Then
        If Val(Replace(dato11.text, ".", "")) <= 5 And Val(Replace(dato13.text, ".", "")) > 0 Then
            dato11.text = 5
            dato12.text = 1
            totalfactura
        End If
    End If
    
    
    tipofactura = 0
    If Option1.Value = True Then
        tipofactura = 1
    End If
    If Option2.Value = True Then
        tipofactura = 2
    End If
    
    
    campos(0, 0) = "tipo"
    campos(1, 0) = "numero"
    campos(2, 0) = "fecha"
    campos(3, 0) = "fechavencimiento"
    campos(4, 0) = "rut"
    campos(5, 0) = "neto"
    campos(6, 0) = "iva"
    campos(7, 0) = "exento"
    campos(8, 0) = "total"
    campos(9, 0) = "fechadigitacion"
    campos(10, 0) = "crcc"
    campos(11, 0) = "tipofactura"
    campos(12, 0) = "itemdte"
    campos(13, 0) = ""
    
    
         
    campos(0, 1) = Val(Mid(COMBOTIPO.text, 1, 3))
    campos(1, 1) = dato2.text
    campos(2, 1) = dato5.text + "-" + dato4.text + "-" + dato3.text
    campos(3, 1) = dato8.text + "-" + dato7.text + "-" + dato6.text
    campos(4, 1) = dato9.text + DV.Caption
    campos(5, 1) = Replace(dato11.text, ".", "")
    campos(6, 1) = Replace(dato12.text, ".", "")
    campos(7, 1) = Replace(dato13.text, ".", "")
    campos(8, 1) = Replace(total.text, ".", "")
    campos(9, 1) = fechasistema
    campos(10, 1) = DATO21.text & DATO22.text
    campos(11, 1) = tipofactura
    campos(12, 1) = txtitemfactura.text
    
    condicion = ""
    campos(0, 2) = "facturasvarias"
    op = 2
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)
    
    For k = 1 To grid1.Rows - 1
    campos(0, 0) = "tipo"
    campos(1, 0) = "numero"
    campos(2, 0) = "linea"
    campos(3, 0) = "glosa"
    campos(4, 0) = ""
    
    campos(0, 1) = Val(Mid(COMBOTIPO.text, 1, 3))
    campos(1, 1) = dato2.text
    campos(2, 1) = k
    campos(3, 1) = grid1.Cell(k, 1).text
    
    condicion = ""
    campos(0, 2) = "facturasvarias_glosa"
    op = 2
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)
    
    Next k
    
    

    If Val(Mid(COMBOTIPO.text, 1, 3)) = 33 Then
        tipo = "FV"
        tipoconta = 6
    End If

    If Val(Mid(COMBOTIPO.text, 1, 3)) = 34 Then ' factura extenta
        tipo = "EX"
        tipoconta = 0
    End If
      
    If Val(Mid(COMBOTIPO.text, 1, 3)) = 56 Then ' nota debito
        tipo = "ND"
        tipoconta = 7
    End If
    If Val(Mid(COMBOTIPO.text, 1, 3)) = "46" Then ' factura de compra
        tipo = "FC"
        tipoconta = 10
    End If
    
    If Val(Mid(COMBOTIPO.text, 1, 3)) = "61" Then 'nota de credito
        tipo = "NF"
        tipoconta = 8
    End If
    
    
     

  loc = CONFI_EMPRESAFAE

 Call grabardte(tipo, dato2.text, 1, dato5.text & "-" & dato4.text & "-" & dato3.text, dato8.text & "-" & dato7.text & "-" & dato6.text, dato9.text & DV.Caption, "0000000000100", txtitemfactura.text, 1, total.text, 0, total.text, "", 0, "00", "0", "99", 0, "", "", dato11.text, dato12.text, dato13.text, loc)
 
End Sub
Sub grabarcontable()
        Dim tipoconta As String
        Dim campos(40, 3) As String
        Dim condicion As String
        Dim op As Integer
        Dim tIPOrEF As String
        
    
    
    If Val(Mid(COMBOTIPO.text, 1, 3)) = 33 Then ' factura
        tipoconta = 6
    End If
    
    If Val(Mid(COMBOTIPO.text, 1, 3)) = 34 Then ' factura extenta
        tipoconta = 0
    End If
    
    If Val(Mid(COMBOTIPO.text, 1, 3)) = 56 Then ' nota debito
        tipoconta = 7
    End If
    
    If Val(Mid(COMBOTIPO.text, 1, 3)) = 61 Then ' nota credito
        tipoconta = 8
    End If
    
    If Val(Mid(COMBOTIPO.text, 1, 3)) = 46 Then ' FACTURA DE COMPRA
        tipoconta = 7
    End If
    
    
    
    campos(0, 0) = "tipo"
    campos(1, 0) = "numero"
    campos(2, 0) = "fecha"
    campos(3, 0) = "fechavencimiento"
    campos(4, 0) = "rut"
    campos(5, 0) = "neto"
    campos(6, 0) = "iva"
    campos(7, 0) = "exento"
    campos(8, 0) = "total"
    campos(9, 0) = "fechadigitacion"
    campos(10, 0) = "crcc"
    campos(11, 0) = "itemdte"
    campos(12, 0) = "añocontable"
    campos(13, 0) = "mescontable"
    
    If Val(Mid(COMBOTIPO.text, 1, 3)) = 46 Then
        campos(14, 0) = ""
    Else
        campos(14, 0) = "caja"
        campos(15, 0) = ""
    End If
    
    If Val(Mid(COMBOTIPO.text, 1, 3)) = 61 Then
        campos(14, 0) = "caja"
        campos(15, 0) = "tnc"
        campos(16, 0) = ""
        campos(17, 0) = ""
    End If
    
    
    campos(0, 1) = tipoconta
    campos(1, 1) = Format(txtfolio.text, "0000000000")
    campos(2, 1) = dato5.text + "-" + dato4.text + "-" + dato3.text
    campos(3, 1) = dato8.text + "-" + dato7.text + "-" + dato6.text
    campos(4, 1) = dato9.text + DV.Caption
    campos(5, 1) = Replace(dato11.text, ".", "")
    campos(6, 1) = Replace(dato12.text, ".", "")
    campos(7, 1) = Replace(dato13.text, ".", "")
    campos(8, 1) = Replace(total.text, ".", "")
    campos(9, 1) = fechasistema
    campos(10, 1) = DATO21.text & DATO22.text
    campos(11, 1) = txtitemfactura.text
    campos(12, 1) = dato5.text
    campos(13, 1) = dato4.text
    If Val(Mid(COMBOTIPO.text, 1, 3)) = 46 Then
        campos(14, 1) = ""
    Else
        campos(14, 1) = "99"
    End If
    
    
    
  
    If Val(Mid(COMBOTIPO.text, 1, 3)) = 46 Then
        campos(0, 2) = "facturasdecompras"
    Else
        campos(0, 2) = "facturasdeventas"
    End If
    
    If Val(Mid(COMBOTIPO.text, 1, 3)) = 61 Then
        tIPOrEF = leerdatos(conta, clientesistema & "ventas" & CONFI_EMPRESAFAE & ".sv_otros_documento_detalle_" & CONFI_EMPRESAFAE, "tipodocumento", "tipo='NF' and numero=lpad('" & txtfolio.text & "',10,0) and fecha='" & dato5.text + "-" + dato4.text + "-" + dato3.text & "' and caja='99' and tipodocumento<>'' ")
        If tIPOrEF = "" Then tIPOrEF = "F"
     campos(14, 1) = "99"
     campos(15, 1) = Mid(tIPOrEF, 1, 1)
    End If
    
    
    condicion = ""
    op = 2
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    
    
    Call sqlconta.sqlconta(op, condicion)
    
Rem GRABADETALLEIMPUESTOS
Call grabardetallefactura(tipoconta)


grabar2

End Sub
Private Sub dato14_GotFocus()
     frmTipo.Visible = True
End Sub

Private Sub dato14_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
            Case 97, 49
                dato14.text = "FV"
            Case 98, 50
                dato14.text = "BV"
            Case 99, 51
                dato14.text = "NC"
             Case 100, 52
                dato14.text = "EX"
            Case 101, 53
                dato14.text = "ND"
            Case Else
             
        End Select
End Sub

Private Sub dato14_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And dato14.text <> "" Then
        dato15.Enabled = True
        dato15.SetFocus
        frmTipo.Visible = False
    End If
End Sub

Private Sub dato15_GotFocus()
    Call cargatexto(dato15)
End Sub

Private Sub dato15_KeyPress(KeyAscii As Integer)
    KeyAscii = esNumero(KeyAscii)
    If KeyAscii = 13 Then
        Call ceros(dato15)
        If dato15.text = "00" Then dato15.text = Format(fechasistema, "dd")
        dato16.Enabled = True
        dato16.SetFocus
    End If
End Sub

Private Sub DATO16_GotFocus()
    Call cargatexto(dato16)
End Sub

Private Sub dato16_KeyPress(KeyAscii As Integer)
     KeyAscii = esNumero(KeyAscii)
    If KeyAscii = 13 Then
        Call ceros(dato16)
        If dato16.text = "00" Then dato16.text = Format(fechasistema, "mm")
            dato17.Enabled = True
            
        dato17.SetFocus
    End If
End Sub

Private Sub dato17_GotFocus()
     Call cargatexto(dato17)
End Sub

Private Sub dato17_KeyPress(KeyAscii As Integer)
     KeyAscii = esNumero(KeyAscii)
    If KeyAscii = 13 Then
        Call ceros(dato17)
        If dato17.text = "0000" Then dato17.text = Format(fechasistema, "yyyy")
        If IsDate(dato15.text & "-" & dato16.text & "-" & dato17.text) = True Then
            dato18.SetFocus
        Else
            MsgBox "FECHA NO ES VALIDA", vbInformation, "ATENCION"
            dato15.text = ""
            dato16.text = ""
            dato17.text = ""
            dato15.SetFocus
        End If
         
    End If
End Sub

Private Sub dato18_GotFocus()
     Call cargatexto(dato18)
End Sub

Private Sub dato18_KeyPress(KeyAscii As Integer)
    KeyAscii = esNumero(KeyAscii)
    If KeyAscii = "13" Then
        Call ceros(dato18)
           txtitemfactura.SetFocus
           totalfactura
    End If
End Sub

Sub grabardte(tipo, numero, LINEA, fecha, vencimiento, rut, codigo, descripcion, Cantidad, Precio, descuento, total, Vendedor, pcosto, bodega, SUCURSAL, caja, descuentopesos, tipodespacho, despacho, NETO, iva, EXENTO, loc)
      'detalles
      
      campos(0, 0) = "local"
      campos(1, 0) = "tipo"
      campos(2, 0) = "numero"
      campos(3, 0) = "linea"
      campos(4, 0) = "fecha"
      campos(5, 0) = "rut"
      campos(6, 0) = "codigo"
      campos(7, 0) = "descripcion"
      campos(8, 0) = "cantidad"
      campos(9, 0) = "precio"
      campos(10, 0) = "descuento"
      campos(11, 0) = "total"
      campos(12, 0) = "vendedor"
      campos(13, 0) = "pcosto"
      campos(14, 0) = "bodega"
      campos(15, 0) = "sucursal"
      campos(16, 0) = "caja"
      campos(17, 0) = "descuentopesos"
      campos(18, 0) = "tipodespacho"
      campos(19, 0) = "despachado"
      campos(20, 0) = "impuesto"
      campos(21, 0) = "tipodocumento"
      campos(22, 0) = "numerodocumento"
      campos(23, 0) = "fechadocumento"
      campos(24, 0) = ""
      
    
            campos(0, 1) = loc
            campos(1, 1) = tipo
            campos(2, 1) = Format(numero, "0000000000")
            campos(3, 1) = Format(LINEA, "000")
            campos(4, 1) = Format(fecha, "yyyy-mm-dd")
            campos(5, 1) = rut
            campos(6, 1) = codigo
            campos(7, 1) = descripcion
            campos(8, 1) = Cantidad
            campos(9, 1) = Replace(Replace(Precio, ".", ""), ",", ".")
            campos(10, 1) = Replace(descuento, ".", "")
            campos(11, 1) = Replace(Replace(total, ".", ""), ",", ".")
            campos(12, 1) = Vendedor
            campos(13, 1) = pcosto
            campos(14, 1) = bodega
            campos(15, 1) = SUCURSAL
            campos(16, 1) = caja
            campos(17, 1) = descuentopesos
            campos(18, 1) = tipodespacho
            campos(19, 1) = 1
            campos(20, 1) = "00000"
            If tipo = "EX" Then
                campos(20, 1) = "00008"
            End If
            If tipo = "FC" Then
                campos(20, 1) = "00015"
            End If
            If tipo = "NF" And dato14.text = "EX" Then
                campos(20, 1) = "00008"
            End If
            campos(21, 1) = dato14.text
            campos(22, 1) = Format(dato18.text, "0000000000")
            campos(23, 1) = dato17.text & "-" & dato16.text & "-" & dato15.text
            
            
            campos(0, 2) = clientesistema & "ventas" & loc & ".sv_otros_documento_detalle_" & loc ' & "_fc"
            condicion = ""
            op = 2
            sqlconta.response = campos
            Set sqlconta.conexion = contadb
            Call sqlconta.sqlconta(op, condicion)
     
   
      
      
      If tipo = "FV" And EXENTO > 0 Then
         LINEA = LINEA + 1
         campos(0, 1) = loc
         campos(1, 1) = tipo
         campos(2, 1) = Format(numero, "0000000000")
         campos(3, 1) = Format(LINEA, "000")
         campos(4, 1) = Format(fecha, "yyyy-mm-dd")
         campos(5, 1) = rut
         campos(6, 1) = "0000000000101"
         campos(7, 1) = Mid(txtitemfacturaEXENTO.text, 1, 50)
         campos(8, 1) = Cantidad
         campos(9, 1) = Replace(Replace(EXENTO, ".", ""), ",", ".")
         campos(10, 1) = Replace(0, ".", "")
         campos(11, 1) = Replace(Replace(EXENTO, ".", ""), ",", ".")
         campos(12, 1) = Vendedor
         campos(13, 1) = pcosto
         campos(14, 1) = bodega
         campos(15, 1) = SUCURSAL
         campos(16, 1) = caja
         campos(17, 1) = descuentopesos
         campos(18, 1) = tipodespacho
         campos(19, 1) = 1
         campos(20, 1) = "00008"
        
         
         campos(0, 2) = clientesistema & "ventas" & loc & ".sv_otros_documento_detalle_" & loc
         condicion = ""
         op = 2
         sqlconta.response = campos
         Set sqlconta.conexion = contadb
         Call sqlconta.sqlconta(op, condicion)
      End If
          
      
      
      'cabeza
      campos(0, 0) = "local"
      campos(1, 0) = "tipo"
      campos(2, 0) = "numero"
      campos(3, 0) = "fecha"
      campos(4, 0) = "plazo"
      campos(5, 0) = "vencimiento"
      campos(6, 0) = "rut"
      campos(7, 0) = "cajera"
      campos(8, 0) = "notapedido"
      campos(9, 0) = "notaventa"
      campos(10, 0) = "ordencompra"
      campos(11, 0) = "neto"
      campos(12, 0) = "iva"
      campos(13, 0) = "impuestoharina"
      campos(14, 0) = "impuestoila"
      campos(15, 0) = "impuestoespecifico"
      campos(16, 0) = "exento"
      campos(17, 0) = "retencionparcial"
      campos(18, 0) = "retenciontotal"
      campos(19, 0) = "total"
      campos(20, 0) = "abono"
      campos(21, 0) = "pagado"
      campos(22, 0) = "caja"
      campos(23, 0) = "horaventas"
      campos(24, 0) = "subtotal"
      campos(25, 0) = "descuento"
      campos(26, 0) = "foliosii"
      campos(27, 0) = "vendedor"
      campos(28, 0) = "contabilizado"
      campos(29, 0) = "sucursal"
      campos(30, 0) = "glosafactura"
      'fechaoc
      campos(31, 0) = "fechaoc"
      campos(32, 0) = ""
      
      
      If tipo = "NF" And dato14.text = "EX" Then
        EXENTO = NETO
        NETO = 0
        iva = 0
        
      End If
      
      campos(0, 1) = loc
      campos(1, 1) = tipo
      campos(2, 1) = Format(numero, "0000000000")
      campos(3, 1) = Format(fecha, "yyyy-mm-dd")
      campos(4, 1) = "000"
      campos(5, 1) = Format(fecha, "yyyy-mm-dd")
      campos(6, 1) = rut
      campos(7, 1) = "000000019"
      campos(8, 1) = "0000000000"
      campos(9, 1) = "0000000000"
      campos(10, 1) = dato13b.text
      campos(11, 1) = Replace(Replace(NETO, ".", ""), ",", ".")
      campos(12, 1) = Replace(Replace(iva, ".", ""), ",", ".")
      campos(13, 1) = "0"
      campos(14, 1) = "0"
      campos(15, 1) = "0"
      campos(16, 1) = Replace(Replace(EXENTO, ".", ""), ",", ".")
      campos(17, 1) = "0"
      campos(18, 1) = "0"
      campos(19, 1) = Replace(Replace(total, ".", ""), ",", ".")
      campos(20, 1) = Replace(Replace(total, ".", ""), ",", ".")
      campos(21, 1) = "S"
      campos(22, 1) = caja
      campos(23, 1) = Time
      campos(24, 1) = Replace(Replace(total, ".", ""), ",", ".")
      campos(25, 1) = "0"
      campos(26, 1) = Format(numero, "0000000000")
      campos(27, 1) = ""
      campos(28, 1) = ""
      campos(29, 1) = SUCURSAL
      campos(30, 1) = grid1.Cell(1, 1).text
      campos(31, 1) = Format(fechaoc.Value, "yyyy-mm-dd")
      
      campos(0, 2) = clientesistema & "ventas" & loc & ".sv_otros_documento_cabeza_" & loc ' & "_fc"
      condicion = ""
      op = 2
      sqlconta.response = campos
      Set sqlconta.conexion = contadb
      Call sqlconta.sqlconta(op, condicion)
      
      'pagos
        campos(0, 0) = "local"
        campos(1, 0) = "tipo"
        campos(2, 0) = "numero"
        campos(3, 0) = "lineapago"
        campos(4, 0) = "fecha"
        campos(5, 0) = "tipopago"
        campos(6, 0) = "cuentacorriente"
        campos(7, 0) = "banco"
        campos(8, 0) = "plaza"
        campos(9, 0) = "numerodocumento"
        campos(10, 0) = "monto"
        campos(11, 0) = "vencimiento"
        campos(12, 0) = "rut"
        campos(13, 0) = "glosa"
        campos(14, 0) = "pagoenlazado"
        campos(15, 0) = "localdocumento"
        campos(16, 0) = "foliofiscal"
        campos(17, 0) = "cuotas"
        campos(18, 0) = "montocuotas"
        campos(19, 0) = "rutcredito"
        campos(20, 0) = "primervencimiento"
        campos(21, 0) = "caja"
        campos(22, 0) = "rutadicional"
        campos(23, 0) = ""
        
        campos(0, 1) = loc
        campos(1, 1) = tipo
        campos(2, 1) = Format(numero, "0000000000")
        campos(3, 1) = Format(LINEA, "000")
        campos(4, 1) = Format(fecha, "yyyy-mm-dd")
        campos(5, 1) = "1"
        campos(6, 1) = ""
        campos(7, 1) = ""
        campos(8, 1) = ""
        campos(9, 1) = ""
        campos(10, 1) = Replace(total, ".", "")
        campos(11, 1) = Format(fecha, "yyyy-mm-dd")
        campos(12, 1) = rut
        campos(13, 1) = ""
        campos(14, 1) = ""
        campos(15, 1) = ""
        campos(16, 1) = Format(numero, "0000000000")
        campos(17, 1) = ""
        campos(18, 1) = ""
        campos(19, 1) = ""
        campos(20, 1) = ""
        campos(22, 1) = ""
        campos(21, 1) = caja
        
       
        campos(0, 2) = clientesistema & "ventas" & loc & ".sv_otros_documento_pagos_" & loc ' & "_fc"
        condicion = ""
        op = 2
        sqlconta.response = campos
        Set sqlconta.conexion = contadb
        Call sqlconta.sqlconta(op, condicion)
        
                
End Sub

Sub grabar2()
    Dim tipo As String
    
    If Val(Mid(COMBOTIPO.text, 1, 3)) = 33 Then ' factura
        tipo = 6
    End If
    
    If Val(Mid(COMBOTIPO.text, 1, 3)) = 34 Then ' factura extenta
        tipo = 0
    End If
    
    If Val(Mid(COMBOTIPO.text, 1, 3)) = 56 Then ' nota debito
        tipo = 7
    End If
    
    If Val(Mid(COMBOTIPO.text, 1, 3)) = 61 Then ' nota credito
        tipo = 8
    End If
    If Val(Mid(COMBOTIPO.text, 1, 3)) = 46 Then ' factura de compra
        tipo = 10
    End If
    
Call leecomprobante(tipo)
opciones.Visible = True
opciones.SetFocus
detalle.Enabled = False
'If GRABACON = True Then
Call GRABARCOMPROBANTE(tipo, Format(txtfolio.text, "0000000000"))
'End If
End Sub
Sub ELIMINAR()
    Dim TIPOCON As String
    Dim MENSA As String
    Dim tipo As String
    Dim tipoconta As String
    
    
    tipo = Val(Mid(COMBOTIPO.text, 1, 3))
    
    If Val(Mid(COMBOTIPO.text, 1, 3)) = 33 Then ' factura
        tipoconta = 6
    End If
    
    If Val(Mid(COMBOTIPO.text, 1, 3)) = 34 Then ' factura extenta
        tipoconta = 0
    End If
    
    If Val(Mid(COMBOTIPO.text, 1, 3)) = 56 Then ' nota debito
        tipoconta = 7
    End If
    
    If Val(Mid(COMBOTIPO.text, 1, 3)) = 61 Then ' nota credito
        tipoconta = 8
    End If
    
    
    
    
    Call ACTUALIZADOCUMENTO("-", tipoconta)

    campos(0, 2) = "facturasvarias"
    condicion = "tipo=" + "'" + tipo + "'" + " and numero=" + "'" + dato2.text + "'"
    op = 4
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)
    
    campos(0, 2) = "facturasvarias_glosa"
    condicion = "tipo=" + "'" + tipo + "'" + " and numero=" + "'" + dato2.text + "'"
    op = 4
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)
    
    
    
    campos(0, 2) = "facturasdeventas_impuestos"
    condicion = "tipo=" + "'" + tipoconta + "'" + " and numero=" + "'" + dato2.text + "'"
    op = 4
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)
    
    campos(0, 2) = "facturasdeventas"
    condicion = "tipo=" + "'" + tipoconta + "'" + " and numero=" + "'" + dato2.text + "'"
    op = 4
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)
    
  
    
     If tipoconta = "9" Then TIPOCON = "FX":
    If tipoconta = "1" Or tipoconta = "6" Then TIPOCON = "FA":
    If tipoconta = "2" Or tipoconta = "7" Then TIPOCON = "ND":
'    If tipo = "3" Or tipo = "6" Then TIPOCON = "NB":
    If tipoconta = "4" Or tipoconta = "8" Then TIPOCON = "NF":
    If tipoconta = "5" Then TIPOCON = "FE":
    If tipoconta = "0" Then TIPOCON = "EX":
    
    
    campos(0, 2) = "movimientoscontables"
    
    condicion = "tipo=" + "'" + TIPOCON + "'" + " and numero=" + "'" + dato2.text + "'"
    op = 4
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)
     
     campos(0, 2) = "facturasdeventas_detalle"
    condicion = "tipo=" + "'" + tipoconta + "'" + " and numero=" + "'" + dato2.text + "'"
    op = 4
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)
  '  If sqlconta.status = 4 Then Stop
    campos(0, 2) = clientesistema & "ventas" & CONFI_EMPRESAFAE & ".sv_otros_documento_detalle_" & CONFI_EMPRESAFAE
    condicion = "tipo=" + "'" + "FV" + "'" + " and numero=" + "'" + dato2.text + "'"
    op = 4
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)
    'clientesistema & "ventas" & loc & ".sv_otros_documento_detalle_" & loc ' & "_fc"
    '  If sqlconta.status = 4 Then Stop
    campos(0, 2) = clientesistema & "ventas" & CONFI_EMPRESAFAE & ".sv_otros_documento_cabeza_" & CONFI_EMPRESAFAE
    condicion = "tipo=" + "'" + "FV" + "'" + " and numero=" + "'" + dato2.text + "'"
    op = 4
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)
    
    campos(0, 2) = clientesistema & "ventas" & CONFI_EMPRESAFAE & ".sv_otros_documento_pagos_" & CONFI_EMPRESAFAE
    condicion = "tipo=" + "'" + "FV" + "'" + " and numero=" + "'" + dato2.text + "'"
    op = 4
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)
  
  
no:



End Sub


Private Sub glosafactura_Change()

End Sub



Private Sub MSHFlexGrid1_Click()

End Sub


Private Sub Grid1_GotFocus()

Rem If dato3.text + dato4.text <> Format(fechasistema, "mm") + Format(fechasistema, "yyyy") Then dato2.text = "": dato3.text = "": dato4.text = "": dato2.SetFocus

End Sub


Private Sub opciones_FSCommand(ByVal command As String, ByVal args As String)
Dim tipo As String

If command = "retorno" And MODIFI = 0 Then retorno
If command = "retorno" And MODIFI = 1 Then grabafactura: retorno
If command = "modifica" Then
  MsgBox "IMPOSIBLE MODIFICAR SOLO ELIMINAR "
End If
  
'ELIMINAR
'dato2.Enabled = True
'dato2.SetFocus
'End If
If command = "elimina" Then
    If txtfolio.text = "NO FISCAL" Then
        If Verifica_Permiso(Me.Caption, "elimina") = True Then
            ELIMINAR
            retorno
        Else
          MsgBox mensaje_nopermiso, vbCritical + vbOKOnly, "Permiso Denegado"
        End If
    Else
          MsgBox "IMPOSIBLE ELIMINAR DOCUMENTOS ELECTRONICOS "
    End If
End If
If command = "imprime" Then
'imprime_factura


  If Val(Mid(COMBOTIPO.text, 1, 3)) = 33 Then
        tipo = "FV"
  End If

  If Val(Mid(COMBOTIPO.text, 1, 3)) = 34 Then ' factura extenta
         tipo = "EX"
  End If
    
  If Val(Mid(COMBOTIPO.text, 1, 3)) = 56 Then ' nota debito
         tipo = "ND"
    End If
    
    
    If Val(Mid(COMBOTIPO.text, 1, 3)) = "61" Then
        tipo = "NF"
    End If
    
     If Val(Mid(COMBOTIPO.text, 1, 3)) = "46" Then
        tipo = "FC"
    End If
    
    

    
     If documentocreado(tipo, "99", CONFI_EMPRESAFAE, dato2.text, dato5.text & "-" & dato4.text & "-" & dato3.text) = False Then
        Call GENERARdocumento(tipo, "99", CONFI_EMPRESAFAE, dato2.text, dato5.text & "-" & dato4.text & "-" & dato3.text)
        
        MsgBox "PRESIONE ENTER PARA GENERAR DOCUMENTO ELECTRONICO DEMORA 20 SEGUNDOS APROX."
        Sleep (10000)
    End If
    
    If documentocreado(tipo, "99", CONFI_EMPRESAFAE, dato2.text, dato5.text & "-" & dato4.text & "-" & dato3.text) = True Then
        
        txtfolio.text = NUMERODOCUMENTO_DTE
        Rem If documento_dte_impreso = False Then
        MsgBox "PRESIONE ENTER PARA COMENZAR IMPRESION DEMORA 10 SEGUNDOS APROX."
            If documento_dte_impreso = False Then
                If tipo <> "FC" Then
                    Call grabarcontable
                End If
            End If
            Call Cargarpdf(Val(Mid(COMBOTIPO.text, 1, 3)), txtfolio.text, dato5.text & "-" & dato4.text & "-" & dato3.text, Val(dato9.text) & DV.Caption, 0)
            Call Sleep(25000)
            Call Cargarpdf(Val(Mid(COMBOTIPO.text, 1, 3)), txtfolio.text, dato5.text & "-" & dato4.text & "-" & dato3.text, Val(dato9.text) & DV.Caption, 1)
            Call modificaimpresa(Val(Mid(COMBOTIPO.text, 1, 3)), Val(txtfolio.text))
        Rem Else
            Rem MsgBox "IMPOSIBLE REIMPRIMIR DOCUMENTO SOLICITA AUTORIZACION "
        Rem End If
     Else
        MsgBox "REINTENTE DOCUMENTO AUN NO GENERADO"
     End If
                                
                                
    
   
    End If
End Sub

Public Function Cargarpdf(tipo, numero, fecha, RUTCLIENTE, hoja) As String
Dim Tamaño As Double
Dim cn As ADODB.Connection
Dim Rs As ADODB.Recordset
Dim mstream As ADODB.Stream
Dim pdfpath, pdfpath1 As String
Dim pdffile As ADODB.Stream
Rem On Error GoTo error:
'If tipo = "1" Then
'    tipo = "33"
'End If
'If tipo = "4" Then
'    tipo = "61"
'End If

Dim ImgTemporal As String
ImgTemporal = "C:\tmp_pdf" & hoja & ".pdf"
If ExisteArchivo(ImgTemporal) = True Then Kill ImgTemporal

Set cn = New ADODB.Connection
cn.Open "DRIVER={MySQL ODBC 3.51 Driver}; SERVER=" & Servidor & "; DATABASE=" & clientesistema & "ventas" & ";PWD=" & password & "; UID=" & Usuario & ";OPTION=3"
cn.CursorLocation = adUseClient


Set Rs = New ADODB.Recordset
'Rs.Open " select * from pdf where pdfid='" & txtid.text & "' and pdfname='" & txtname.text & "'", cn, adOpenKeyset, adLockOptimistic
Rs.Open "Select * from " & cliente_sql & "fae" & CONFI_EMPRESAFAE & ".sv_dtepdf_" & CONFI_EMPRESAFAE & " where tipo='" & tipo & "' and numero='" & numero & "' and rut = '" & RUTCLIENTE & "' and  fecha ='" & Format(fecha, "yyyy-mm-dd") & "' and cedible='" & hoja & "' limit 0,1 ", cn, adOpenKeyset, adLockOptimistic

If Not Rs.EOF Then
Set pdffile = New ADODB.Stream
pdffile.Type = adTypeBinary
pdffile.Open
If IsNull(Rs.Fields("pdf")) = False Then
pdffile.Write Rs.Fields("pdf").Value
'Dim pdfnme As String
'pdfnme = txtid.text & txtname.text
'pdffile.SaveToFile "" & App.Path & "\reports\" & pdfnme & ".pdf", adSaveCreateOverWrite
pdffile.SaveToFile ImgTemporal, adSaveCreateOverWrite
pdffile.SaveToFile ImgTemporal, adSaveCreateOverWrite
pdffile.Close
Set pdffile = Nothing
'ShellExecute Me.hwnd, "print", ImgTemporal, vbNullString, App.path, 0
 ShellExecute Me.hwnd, "open", ImgTemporal, vbNullString, App.path, 0

'Shell ImgTemporal
'Shell "C:\Archivos de programa\Adobe\Reader 10.0\Reader\AcroRd32.exe " & ImgTemporal

'Call ShellAndWait
'MsgBox "pdf file downloaded"
Else
MsgBox "NO SE HA ENCONTRADO EL ARCHIVO", vbCritical, "ATENCION"
Rs.Close
Set Rs = Nothing
End If
End If
Exit Function
error:
MsgBox "debe tener creada la carpeta C:\Archivos de programa\Adobe\Reader 10.0\Reader\AcroRd32.exe "
End Function


Sub retorno()



grid1.Rows = 1


opciones.Visible = False
limpia
disponible (False)

 
dato2.Enabled = True
dato2.SetFocus

End Sub


Sub limpia()

    nombreproveedor.Caption = ""
    lbldireccion.Caption = ""
    lblcomuna.Caption = ""
    lblciudad.Caption = ""
    lblgiro.Caption = ""
    
    
    'dato1.text = ""
    dato2.text = ""
    dato3.text = ""
    dato4.text = ""
    dato5.text = ""
    dato6.text = ""
    dato7.text = ""
    dato8.text = ""
    dato9.text = ""
    DV.Caption = ""
    dato11.text = "0"
    dato12.text = "0"
    dato13.text = "0"
    dato13b.text = "0000000000"
    fechaoc.Value = Now()
    
    dato14.text = ""
    dato15.text = ""
    dato16.text = ""
    dato17.text = ""
    dato18.text = ""
    
    txtitemfactura.text = ""
    txtitemfacturaEXENTO.text = ""
    total.text = "0"
    txtfolio.text = ""
    frmexento.Visible = False
    
    LINEAS.text = "001"
grid1.Rows = 1
NETO = 0
SUMAR
no:
End Sub
Sub ayudactacte2(ByRef caja As TextBox)
    Dim campos As Variant
    Dim cfijo As Variant
    Dim largo As Variant
    campos = Array("rut", "nombre")
    largo = Array("12n", "40s")
    cfijo = "sucursal='0' "
    cabezas = Array("RUT", "NOMBRE")
    mensajeAyuda = "Ayuda Cuentas Corrientes"
    pivote2.MaxLength = 10
    
    Call cargaAyudaT(Servidor, basebus, Usuario, password, "eltit_ventas.sv_maestroclientes", pivote2, campos, cfijo, largo, 2)
    If Val(pivote2.text) = 0 Then dato9.SetFocus: GoTo no
    dato9.text = Mid(pivote2.text, 1, 9)
    DV.Caption = Mid(pivote2.text, 10, 1)
    caja.Enabled = True
    caja.SetFocus
no:

End Sub



Sub Pregunta(ByRef caja As TextBox, ByRef caja1 As TextBox)
    If caja.text = "" And sc = 0 Then caja.SetFocus: caja.SelStart = 0
    If caja.text <> "" Or sc = 1 Then caja1.Enabled = True: caja1.SetFocus
    sc = 0
End Sub

Sub leemayor(cuenta)
    campos(0, 0) = "codigo"
    campos(1, 0) = "nombre"
    campos(2, 0) = "ctacte"
    campos(3, 0) = "crcc"
    campos(4, 0) = ""
    campos(0, 2) = "cuentasdelmayor"
    
    condicion = "codigo=" + "'" + cuenta + "' and año='" + año + "'"
    op = 5
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)

    If sqlconta.status = 4 Then
    
    End If
    If Val(sqlconta.response(2, 3)) <> 0 Then tipocuenta = sqlconta.response(2, 3)
    tipocentro = sqlconta.response(3, 3)

no:

End Sub
Sub leectacte_FINAL()
Dim cuentapublicidad As String
Rem cuentapublicidad = leerdatos(conta, "maestroempresas", "cuentapublicidad", "codigoempresa='" + empresaactiva + "' ")

    campos(0, 0) = "rut"
    campos(1, 0) = "nombre"
    campos(2, 0) = "direccion"
    campos(3, 0) = "comuna"
    campos(4, 0) = "ciudad"
    campos(5, 0) = "giro"
    campos(6, 0) = ""
    
    
   campos(0, 2) = "eltit_ventas.sv_maestroclientes "
    condicion = "rut=" + "'" + pivote2.text + "' and sucursal='0' "
    op = 5
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)
    If sqlconta.status = 0 Then
    nombreproveedor.Caption = sqlconta.response(1, 3)
    lbldireccion.Caption = sqlconta.response(2, 3)
    lblcomuna.Caption = sqlconta.response(3, 3)
    lblciudad.Caption = sqlconta.response(4, 3)
    lblgiro.Caption = sqlconta.response(5, 3)
    
    dato3.Enabled = True
    dato3.SetFocus
Else
MsgBox "rut de cliente no existe "
dato2.SetFocus

End If
no:

End Sub


Sub leetipos()
    
    campos(0, 0) = "tipos"
    campos(1, 0) = "nombredocumento"
    campos(2, 0) = ""
    
    campos(0, 2) = "maestrotipodedocumentos"
    condicion = "tipos=" + "'" + dato13.text + "'"
    op = 5
    sqlconta.response = campos
    Set sqlconta.conexion = conta
    Call sqlconta.sqlconta(op, condicion)

    If sqlconta.status = 4 Then dato13.text = "": dato13.SetFocus:  GoTo no:
    varipaso = "S"
    

no:

End Sub
Sub cargatexto(ByRef caja As TextBox)


caja.SelStart = 0: caja.SelLength = Len(caja.text)

End Sub






Sub totalfactura()
sumador = CDbl(Replace(dato11.text, ",", "")) + CDbl(Replace(dato12.text, ",", "")) + CDbl(Replace(dato13.text, ",", ""))
total.text = Format(sumador, "###,###,###,##0")
NETO = CDbl(Replace(dato11.text, ",", "")) + CDbl(Replace(dato13.text, ",", ""))
debe.Caption = Format(NETO, "###,###,##0")
End Sub
Sub leefactura()
    Dim tipo As String
    
    campos(0, 0) = "tipo"
    campos(1, 0) = "numero"
    campos(2, 0) = "fecha"
    campos(3, 0) = "fechavencimiento"
    campos(4, 0) = "rut"
    campos(5, 0) = "neto"
    campos(6, 0) = "iva"
    campos(7, 0) = "exento"
    campos(8, 0) = "retencion"
    campos(9, 0) = "total"
    campos(10, 0) = "tipofactura"
    campos(11, 0) = "itemdte"
    campos(12, 0) = ""
    tipo = Val(Mid(COMBOTIPO.text, 1, 3))
    
    campos(0, 2) = "facturasvarias"
    condicion = "tipo=" & "'" + tipo + "'" & " and numero=" + "'" + dato2.text + "'"

    op = 5
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)
    leedetalle
    
    Rem If sqlconta.status = 0 Then modifi = 1: carga: opciones.Visible = True: disponible (True): habilita (True): opciones.SetFocus


End Sub
Sub grabardetallefactura(tipo)
    Dim TIPOCON As String
    Dim tipo2 As String
    Dim j As Integer
    Dim lin As Integer
    
    
    campos(0, 0) = "tipo"
    campos(1, 0) = "numero"
    campos(2, 0) = "linea"
    campos(3, 0) = "rut"
    campos(4, 0) = "cuentadelmayor"
    campos(5, 0) = "glosa"
    campos(6, 0) = "monto"
    campos(7, 0) = "dh"
    campos(8, 0) = "centrodecosto"
    campos(9, 0) = "rutctacte"
    campos(10, 0) = ""
    publicidad = False
    empresarelacionada = False
    
    GRABACON = False
    
Rem graba detalle factura
    Rem error aca lo hace pero parce esta listo
    
    LINEAS.text = "1"
    Call ceros(LINEAS)
    campos(0, 1) = tipo
    campos(1, 1) = Format(txtfolio.text, "0000000000")
    campos(2, 1) = LINEAS.text
    campos(3, 1) = dato9.text + DV.Caption
    If Option1.Value = True And tipo <> "0" Then campos(4, 1) = "35100001": campos(5, 1) = "INGRESOS POR VENTAS"
    If Option1.Value = True And tipo = "0" Then campos(4, 1) = "35100002": campos(5, 1) = "ARRIENDO EXENTO"
    If Option2.Value = True Then campos(4, 1) = "35150001": campos(5, 1) = "INGRESOS POR PUBLICIDAD"
    If tipo <> "0" Then
        campos(6, 1) = Replace(dato11.text, ".", "")
    Else
        campos(6, 1) = Replace(dato13.text, ".", "")
    End If
    campos(7, 1) = "H"
    campos(8, 1) = DATO21.text & DATO22.text
    campos(9, 1) = ""
    If Val(Mid(COMBOTIPO.text, 1, 3)) = 46 Then
       campos(0, 2) = "facturasdecompras_detalle"
    Else
        campos(0, 2) = "facturasdeventas_detalle"
    End If
    
 
    condicion = ""
    op% = 2
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)
    publicidad = True
    GRABACON = True
    
    
    
    
End Sub

'
''
''Sub grabardetallefactura()
''    Dim tipocon As String
''    Dim tipo2 As String
''    Dim j As Integer
''    Dim lin As Integer
''
''
''    campos(0, 0) = "tipo"
''    campos(1, 0) = "numero"
''    campos(2, 0) = "linea"
''    campos(3, 0) = "fecha"
''    campos(4, 0) = "codigocuenta"
''    campos(5, 0) = "tipoctacte"
''    campos(6, 0) = "rutctacte"
''    campos(7, 0) = "centrocosto"
''    campos(8, 0) = "glosacontable"
''    campos(9, 0) = "tipodocumento"
''    campos(10, 0) = "numerodocumento"
''    campos(11, 0) = "fechadocumento"
''    campos(12, 0) = "fechavencimiento"
''    campos(13, 0) = "monto"
''    campos(14, 0) = "dh"
''    campos(15, 0) = "creadopor"
''    campos(16, 0) = "mes"
''    campos(17, 0) = "año"
''    campos(18, 0) = "fechacreacion"
''    campos(19, 0) = "horacreacion"
''    campos(20, 0) = "rutproveedor"
''    campos(21, 0) = ""
''    Rem cuenta proveedores
''    If dato1.text = "1" Then tipocon = "FV"
''    If dato1.text = "2" Then tipocon = "DV"
''    If dato1.text = "3" Then tipocon = "NV"
''    If dato1.text = "4" Then tipocon = "NB"
''
''
''Rem graba impuestos
''    lin = 0
''    For j = 1 To Grid2.Rows - 1
''
''    If Val(Grid2.Cell(j, 3).text) <> 0 Then
''    lin = lin + 1
''    LINEAS.text = lin
''    Call ceros(LINEAS)
''
''    campos(0, 1) = tipocon
''    campos(1, 1) = dato2.text
''    campos(2, 1) = LINEAS.text
''    campos(3, 1) = dato5.text + "-" + dato4.text + "-" + dato3.text
''    campos(4, 1) = Grid2.Cell(j, 1).text
''    campos(5, 1) = ""
''    campos(6, 1) = dato9.text + DV.Caption
''
''    campos(7, 1) = ""
''    campos(8, 1) = Grid2.Cell(j, 2).text
''    campos(9, 1) = DOCU(Val(dato1.text))
''    campos(10, 1) = dato2.text
''    campos(11, 1) = dato5.text + "-" + dato4.text + "-" + dato3.text
''    campos(12, 1) = dato5.text + "-" + dato4.text + "-" + dato3.text
''    campos(13, 1) = Grid2.Cell(j, 3).text
''
''    campos(14, 1) = "D"
''    campos(15, 1) = USUARIOSISTEMA
''    campos(16, 1) = MES
''    campos(17, 1) = año
''    campos(18, 1) = Format(Date$, "yyyy") + "-" + Format(Date$, "mm") + "-" + Format(Date$, "dd")
''    campos(19, 1) = Time$
''    campos(20, 1) = campos(6, 1)
''
''    campos(0, 2) = "facturasdeventas_detalle"
''    condicion = ""
''    sqlconta.response = campos
''    Set sqlconta.conexion = contadb
''    Call sqlconta.sqlconta(op, condicion)
''    End If
''
''    Next j
'
'
'
'    For K = 1 To Grid1.Rows - 2
'    lin = lin + 1
'    LINEAS.text = lin
'    Call ceros(LINEAS)
'
'    campos(0, 1) = tipocon
'    campos(1, 1) = dato2.text
'    campos(2, 1) = LINEAS.text
'    campos(3, 1) = dato5.text + "-" + dato4.text + "-" + dato3.text
'    campos(4, 1) = Grid1.Cell(K, 1).text + Grid1.Cell(K, 2).text + Grid1.Cell(K, 3).text
'    campos(5, 1) = ""
'    campos(6, 1) = Grid1.Cell(K, 10).text
'    campos(7, 1) = Grid1.Cell(K, 11).text
'    campos(8, 1) = Grid1.Cell(K, 4).text
'    campos(9, 1) = DOCU(Val(dato1.text))
'    campos(10, 1) = dato2.text
'    campos(11, 1) = dato5.text + "-" + dato4.text + "-" + dato3.text
'    campos(12, 1) = dato8.text + "-" + dato7.text + "-" + dato6.text
'    campos(13, 1) = Grid1.Cell(K, 5).text
'    campos(14, 1) = Grid1.Cell(K, 6).text
'    campos(15, 1) = USUARIOSISTEMA
'    campos(16, 1) = MES
'    campos(17, 1) = año
'    campos(18, 1) = Format(Date$, "yyyy") + "-" + Format(Date$, "mm") + "-" + Format(Date$, "dd")
'    campos(19, 1) = Time$
'    campos(20, 1) = dato9.text + DV.Caption
'
'    campos(0, 2) = "facturasdeventas_detalle"
'    condicion = ""
'
'    op = 2
'    sqlconta.response = campos
'    Set sqlconta.conexion = contadb
'    Call sqlconta.sqlconta(op, condicion)
'    Next K
'
'End Sub
'



Sub GRABARCOMPROBANTE(tipo, numero)
    Dim TIPOCON As String
    Dim tipo2 As String
    Dim j As Integer
    Dim lin As Integer
    Dim HD1 As String
    Dim HD2 As String
    Dim cuentafinal As String
    
  
    
    
    campos(0, 0) = "tipo"
    campos(1, 0) = "numero"
    campos(2, 0) = "linea"
    campos(3, 0) = "fecha"
    campos(4, 0) = "codigocuenta"
    campos(5, 0) = "tipoctacte"
    campos(6, 0) = "rutctacte"
    campos(7, 0) = "centrocosto"
    campos(8, 0) = "glosacontable"
    campos(9, 0) = "tipodocumento"
    campos(10, 0) = "numerodocumento"
    campos(11, 0) = "fechadocumento"
    campos(12, 0) = "fechavencimiento"
    campos(13, 0) = "monto"
    campos(14, 0) = "dh"
    campos(15, 0) = "creadopor"
    campos(16, 0) = "mes"
    campos(17, 0) = "año"
    campos(18, 0) = "fechacreacion"
    campos(19, 0) = "horacreacion"
    campos(20, 0) = ""
    Rem cuenta proveedores
    
'    If tipo = "9" Then TIPOCON = "FX": HD1 = "D": HD2 = "H"
'    If tipo = "1" Or tipo = "6" Then TIPOCON = "FA": HD1 = "D": HD2 = "H"
'    If tipo = "2" Or tipo = "7" Then TIPOCON = "ND": HD1 = "D": HD2 = "H"
''    If tipo = "3" Or tipo = "6" Then TIPOCON = "NB": HD1 = "H": HD2 = "D"
'    If tipo = "4" Or tipo = "8" Then TIPOCON = "NF": HD1 = "H": HD2 = "D"
'    If tipo = "5" Then TIPOCON = "FE": HD1 = "D": HD2 = "H"
'    If tipo = "0" Then TIPOCON = "EX": HD1 = "D": HD2 = "H"
'    If tipo = "10" Then TIPOCON = "FC": HD1 = "D": HD2 = "H"
    
    
    If tipo = "9" Then TIPOCON = "FX": HD1 = "D": HD2 = "H"
    If tipo = "1" Then TIPOCON = "FA": HD1 = "D": HD2 = "H"
    If tipo = "2" Then TIPOCON = "ND": HD1 = "D": HD2 = "H"
    If tipo = "3" Then TIPOCON = "NB": HD1 = "H": HD2 = "D"
    If tipo = "4" Then TIPOCON = "NF": HD1 = "H": HD2 = "D"
    If tipo = "6" Then TIPOCON = "EF": HD1 = "D": HD2 = "H"
    If tipo = "7" Then TIPOCON = "ED": HD1 = "D": HD2 = "H"
    If tipo = "8" Then TIPOCON = "EC": HD1 = "H": HD2 = "D"
    If tipo = "0" Then TIPOCON = "EX": HD1 = "D": HD2 = "H"
    If tipo = "10" Then TIPOCON = "FC": HD1 = "D": HD2 = "H"
    
      
      If Option1.Value = True And tipo = "4" Then ' si es nota de credito y no es publicidad no se contabiliza 4-07-2019
        Exit Sub
      End If
    
'    If empresaactiva = "29" Then
        If Option1.Value = True Then
           cuentafinal = "11200029"
        End If
            
        If Option2.Value = True Then
            cuentafinal = "11200028"
        End If
'    Else
'        cuentafinal = cuentacliente
'    End If
    
    
    If Option2.Value = True And tipo = "4" Then
        cuentafinal = "11200028"
    End If
    
    campos(0, 1) = TIPOCON
    campos(1, 1) = Format(numero, "0000000000")
    campos(2, 1) = "001"
    campos(3, 1) = dato5.text + "-" + dato4.text + "-" + dato3.text
    campos(4, 1) = cuentafinal
    campos(5, 1) = tipocuenta
    campos(6, 1) = dato9.text + DV.Caption
    campos(7, 1) = ""
    campos(8, 1) = "CONTABILIZACION " + DOCU$(Val(tipo)) + " " + nombreproveedor.Caption
    campos(9, 1) = DOCU$(Val(tipo))
    campos(10, 1) = numero
    campos(11, 1) = dato5.text + "-" + dato4.text + "-" + dato3.text
    campos(12, 1) = dato8.text + "-" + dato7.text + "-" + dato6.text
    campos(13, 1) = Replace(total.text, ".", "")

    campos(14, 1) = HD1
    campos(15, 1) = USUARIOSISTEMA
    campos(16, 1) = dato4.text
    campos(17, 1) = dato5.text
    campos(18, 1) = Format(Date$, "yyyy") + "-" + Format(Date$, "mm") + "-" + Format(Date$, "dd")
    campos(19, 1) = Time$

    campos(0, 2) = "movimientoscontables"
    condicion = ""

    op = 2
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)
    
    
    Rem cuenta I.V.A
    campos(0, 1) = TIPOCON
    campos(1, 1) = numero
    campos(2, 1) = "002"
    campos(3, 1) = dato5.text + "-" + dato4.text + "-" + dato3.text
    campos(4, 1) = ivadebito
    campos(5, 1) = ""
    campos(6, 1) = ""
    campos(7, 1) = ""
    campos(8, 1) = "CONTABILIZACION I.V.A " + DOCU(Val(tipo))
    campos(9, 1) = DOCU(Val(tipo))
    campos(10, 1) = numero
    campos(11, 1) = dato5.text + "-" + dato4.text + "-" + dato3.text
    campos(12, 1) = dato5.text + "-" + dato4.text + "-" + dato3.text
    campos(13, 1) = Replace(dato12.text, ".", "")
    
    campos(14, 1) = HD2
    campos(15, 1) = USUARIOSISTEMA
    campos(16, 1) = MES
    campos(17, 1) = año
    campos(0, 2) = "movimientoscontables"
    condicion = ""
    op = 2
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)
    
    
   
    
Rem graba impuestos
    lin = 3
    For j = 1 To Grid2.Rows - 1
   
    If Val(Grid2.Cell(j, 3).text) <> 0 Then
    lin = lin + 1
    LINEAS.text = lin
    Call ceros(LINEAS)
    
    campos(0, 1) = TIPOCON
    campos(1, 1) = numero
    campos(2, 1) = LINEAS.text
    campos(3, 1) = dato5.text + "-" + dato4.text + "-" + dato3.text
    campos(4, 1) = Grid2.Cell(j, 1).text
    campos(5, 1) = ""
    campos(6, 1) = ""
    campos(7, 1) = ""
    campos(8, 1) = Grid2.Cell(j, 2).text
    campos(9, 1) = DOCU(Val(tipo))
    campos(10, 1) = dato2.text
    campos(11, 1) = dato5.text + "-" + dato4.text + "-" + dato3.text
    campos(12, 1) = dato5.text + "-" + dato4.text + "-" + dato3.text
    campos(13, 1) = Grid2.Cell(j, 3).text
    
    campos(14, 1) = "H"
    campos(15, 1) = USUARIOSISTEMA
    campos(16, 1) = dato4.text
    campos(17, 1) = dato5.text
    campos(0, 2) = "movimientoscontables"
    condicion = ""
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)
    End If
    
    Next j

    
    
    For k = 1 To Grid3.Rows - 2
    lin = lin + 1
    LINEAS.text = lin
    Call ceros(LINEAS)
    
    campos(0, 1) = TIPOCON
    campos(1, 1) = numero
    campos(2, 1) = LINEAS.text
    campos(3, 1) = dato5.text + "-" + dato4.text + "-" + dato3.text
    campos(4, 1) = Grid3.Cell(k, 1).text + Grid3.Cell(k, 2).text + Grid3.Cell(k, 3).text
    campos(5, 1) = ""
    campos(6, 1) = Grid3.Cell(k, 10).text
    campos(7, 1) = Grid3.Cell(k, 11).text
    campos(8, 1) = Grid3.Cell(k, 4).text
    campos(9, 1) = DOCU(Val(tipo))
    campos(10, 1) = numero
    campos(11, 1) = dato5.text + "-" + dato4.text + "-" + dato3.text
    campos(12, 1) = dato8.text + "-" + dato7.text + "-" + dato6.text
    campos(13, 1) = Grid3.Cell(k, 5).text
    campos(14, 1) = HD2
    campos(15, 1) = USUARIOSISTEMA
    campos(16, 1) = dato4.text
    campos(17, 1) = dato5.text
    campos(0, 2) = "movimientoscontables"
    condicion = ""


    op = 2
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)
    Next k
    
     Rem cuenta exento
     
   
    
    
    Call ACTUALIZADOCUMENTO("+", tipo)
   
End Sub




Sub ELIMINA()
    Dim tipoconta As String
 
     If Val(Mid(COMBOTIPO.text, 1, 3)) = 33 Then ' factura
        tipoconta = 6
    End If
    
    If Val(Mid(COMBOTIPO.text, 1, 3)) = 34 Then ' factura extenta
        tipoconta = 0
    End If
    
    If Val(Mid(COMBOTIPO.text, 1, 3)) = 56 Then ' nota debito
        tipoconta = 7
    End If
    
    If Val(Mid(COMBOTIPO.text, 1, 3)) = 61 Then ' nota credito
        tipoconta = 8
    End If
    
    
Call ACTUALIZADOCUMENTO("-", tipoconta)
End Sub



Sub ayudamayor(row As Long, col As Long)
    Dim campos As Variant
    Dim cfijo As Variant
    Dim largo As Variant
    campos = Array("codigo", "nombre")
    cabezas = Array("codigo", "nombre")
    largo = Array("8n", "40s")
    mensajeAyuda = "Ayuda Cuentas del Mayor"
    cfijo = "año='" + Format(fechasistema, "yyyy") + "'"

    Call cargaAyudaT(Servidor, basebus, Usuario, password, "cuentasdelmayor", pivote2, campos, cfijo, largo, 2)
    grid1.Cell(row, col).text = Mid(pivote2.text, 1, 2)
    grid1.Cell(row, col + 1).text = Mid(pivote2.text, 3, 2)
    grid1.Cell(row, col + 2).text = Mid(pivote2.text, 5, 4)
    Rem Call leermayor(row, col)
    respu = ""
    If pivote2.text <> "" Then Call leermayor(row, col): respu = "S"
    pivote2.text = ""
    
End Sub

Sub disponible(ByVal condicion As Boolean)
    
    
    dato2.Enabled = condicion
    dato3.Enabled = condicion
    dato4.Enabled = condicion
    dato5.Enabled = condicion
    dato6.Enabled = condicion
    dato7.Enabled = condicion
    dato8.Enabled = condicion
    dato9.Enabled = condicion
    
    dato11.Enabled = condicion
    dato12.Enabled = condicion
    dato13.Enabled = condicion
    dato13b.Enabled = condicion
    fechaoc.Enabled = condicion
   
    total.Enabled = condicion
   
    
    
End Sub


Sub ACTUALIZADOCUMENTO(COMANDO As String, tipo)
    Dim lin As Integer
    Dim TIPOCON As String
    Dim resultados As rdoResultset
    Dim csql As New rdoQuery
    Dim rut As String
    Dim TIPOFA As String
    
'    If tipo = "9" Then TIPOCON = "FX":
'    If tipo = "1" Or tipo = "6" Then TIPOCON = "FA":
'    If tipo = "2" Or tipo = "7" Then TIPOCON = "ND":
''    If tipo = "3" Or tipo = "6" Then TIPOCON = "NB":
'    If tipo = "4" Or tipo = "8" Then TIPOCON = "NF":
'    If tipo = "5" Then TIPOCON = "FE":
'    If tipo = "0" Then TIPOCON = "EX":
'    If tipo = "10" Then TIPOCON = "FC":
    
    If tipo = "9" Then TIPOCON = "FX"
    If tipo = "1" Then TIPOCON = "FA"
    If tipo = "2" Then TIPOCON = "ND"
    If tipo = "3" Then TIPOCON = "NB"
    If tipo = "4" Then TIPOCON = "NF"
    If tipo = "6" Then TIPOCON = "EF"
    If tipo = "7" Then TIPOCON = "ED"
    If tipo = "8" Then TIPOCON = "EC"
    If tipo = "0" Then TIPOCON = "EX"
    If tipo = "10" Then TIPOCON = "FC"
    
    
    
      
        Set csql.ActiveConnection = contadb
       csql.sql = "SELECT tipo,numero,linea,fecha,codigocuenta,tipoctacte,rutctacte,centrocosto,glosacontable,tipodocumento,numerodocumento,fechavencimiento,monto,dh "
      csql.sql = csql.sql + "FROM movimientoscontables "
     
            
        csql.sql = csql.sql + "WHERE tipo='" + TIPOCON + "' and numero='" & Format(txtfolio.text, "0000000000") & "'and año='" + año + "' and mes='" + MES + "' order by linea"
        csql.Execute


        If csql.RowsAffected > 0 Then
            Set resultados = csql.OpenResultset
            While Not resultados.EOF
                
                Call actualizamayor(COMANDO, resultados(4), resultados(12), resultados(13), resultados(5), resultados(6), resultados(7), MES, año)
                
                resultados.MoveNext
            Wend
            
            resultados.Close
            Set resultados = Nothing
        End If
   
End Sub

Private Sub opciones_GotFocus()
MANUAL.SetFocus

End Sub

Sub leeFOLIO()

End Sub
Sub CARGAGRILLAexento()
    formatogrilla2(1, 1) = "CODIGO"
    formatogrilla2(1, 2) = "IMPUESTO"
    formatogrilla2(1, 3) = "MONTO"
    Rem LARGO DE LOS DATOS
    formatogrilla2(2, 1) = "10"
    
    formatogrilla2(2, 2) = "30"
    formatogrilla2(2, 3) = "10"
    
    Rem TIPO DE DATOS C=CEROS,S=STRING,N=NUMERICO,D=DATE,
    formatogrilla2(3, 1) = "C"
    formatogrilla2(3, 2) = "C"
    formatogrilla2(3, 3) = "N"
    Rem FORMATO GRILLA
    
    formatogrilla2(4, 3) = " ###,###,##0"
    Rem LOCCKED
    formatogrilla2(5, 1) = "TRUE"
    
    formatogrilla2(5, 2) = "TRUE"
    formatogrilla2(5, 3) = "FALSE"
    Rem VALOR MAXIMO
    formatogrilla2(7, 3) = "999999999"
    Grid2.Cols = 4
    Grid2.Rows = 1
    Grid2.AllowUserResizing = False
    Grid2.DisplayFocusRect = False
    Grid2.ExtendLastCol = True
    Grid2.BoldFixedCell = False
    Grid2.DrawMode = cellOwnerDraw
    Grid2.Appearance = Flat
    Grid2.ScrollBarStyle = Flat
    Grid2.FixedRowColStyle = Flat
    Grid2.BackColorFixed = RGB(90, 158, 214)
    Grid2.BackColorFixedSel = RGB(110, 180, 230)
    Grid2.BackColorBkg = RGB(90, 158, 214)
    Grid2.BackColorScrollBar = RGB(231, 235, 247)
    Grid2.BackColor1 = RGB(231, 235, 247)
    Grid2.BackColor2 = RGB(239, 243, 255)
    Grid2.GridColor = RGB(148, 190, 231)
    Grid2.Column(0).Width = 0
    
    For k = 1 To Grid2.Cols - 1
        Grid2.Cell(0, k).text = formatogrilla2(1, k)
        If k < 5 Then Grid2.Column(k).Width = Val(formatogrilla2(2, k)) * 8
        
        
        Rem Grid1.Column(K).Width = Val(formatogrilla(2, K)) * 9
        Grid2.Column(k).MaxLength = Val(formatogrilla2(2, k))
        Grid2.Column(k).FormatString = formatogrilla2(4, k)
        Grid2.Column(k).Locked = formatogrilla2(5, k)
        If formatogrilla2(3, k) = "N" Then Grid2.Column(k).Alignment = cellRightCenter
        If formatogrilla2(3, k) = "D" Then Grid2.Column(k).CellType = cellCalendar
        
    Next k
   Call leecuentas
    
    End Sub

Sub leecuentas()

Dim resultados2 As rdoResultset
    Dim csql2 As New rdoQuery
    Dim rut As String
    Dim PASO As String
    Dim LINEAS As Double
    
       
        
        Set csql2.ActiveConnection = contadb
        csql2.sql = "SELECT codigo,nombre "
        csql2.sql = csql2.sql + "FROM cuentasdelmayor where ila<>'0' or iha<>'0' or ica<>'0' and año='" + Format(fechasistema, "yyyy") + "' "
       
        csql2.sql = csql2.sql + "order by codigo"
        csql2.Execute
        
        LINEAS = 0
        If csql2.RowsAffected > 0 Then
        Set resultados2 = csql2.OpenResultset
        While Not resultados2.EOF
        
        LINEAS = LINEAS + 1
        Grid2.Rows = Grid2.Rows + 1
        Grid2.Cell(LINEAS, 1).text = resultados2(0)
        Grid2.Cell(LINEAS, 2).text = resultados2(1)
        
        resultados2.MoveNext
        Wend
          
          resultados2.Close
            Set resultados2 = Nothing

        End If
      
cdi = LINEAS

End Sub
Sub SumaImpuestos()
Dim valor As Double
valor = 0
For k = 1 To Grid2.Rows - 1
valor = valor + Val(Grid2.Cell(k, 3).text)

Next k
dato13.text = valor

End Sub

Sub CREARCTACTE(row)
maestro02.dato1.text = grid1.Cell(row, 1).text + grid1.Cell(row, 2).text + grid1.Cell(row, 3).text
maestro02.dato2.text = Mid(grid1.Cell(row, 4).text, 1, 9)
maestro02.Show


End Sub


Sub leermayor(row As Long, col As Long)
    TIENECTACTE(row) = "0"
    TIENECRCC(row) = "0"
    TIENEBANCO(row) = "0"
    TIENEILA(row) = "0"
    TIENEICA(row) = "0"
    TIENEIHA(row) = "0"
    TIENEACTIVO(row) = "0"
    CUENTAMAYOR(row) = "0"
    
    campos(0, 0) = "codigo"
    campos(1, 0) = "nombre"
    campos(2, 0) = "ctacte"
    campos(3, 0) = "crcc"
    campos(4, 0) = "banco"
    campos(5, 0) = "ila"
    campos(6, 0) = "ica"
    campos(7, 0) = "iha"
    campos(8, 0) = "activo"
    
    campos(9, 0) = ""
    campos(0, 2) = "cuentasdelmayor"
    
condicion = "codigo=" + "'" + grid1.Cell(row, 1).text + grid1.Cell(row, 2).text + grid1.Cell(row, 3).text + "' and año='" + Format(fechasistema, "yyyy") + "'"
    
    op = 5
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)
 
    If sqlconta.status = 4 Or grid1.Cell(row, 3).text = "0000" Or sqlconta.response(2, 3) = "1" Then
    grid1.Cell(row, 1).text = ""
    grid1.Cell(row, 2).text = ""
    grid1.Cell(row, 3).text = ""
    grid1.Cell(row, 1).SetFocus
    
    respuesta = "N"
    Else
    respuesta = "S"
    If grid1.Cell(row, 7).text <> "" Then
    grid1.Cell(row, 7).text = sqlconta.response(1, 3)

    grid1.Cell(row, 4).text = sqlconta.response(1, 3)
    End If
    TIENECTACTE(row) = sqlconta.response(2, 3)
    TIENECRCC(row) = sqlconta.response(3, 3)
    TIENEBANCO(row) = sqlconta.response(4, 3)
    TIENEILA(row) = sqlconta.response(5, 3)
    TIENEICA(row) = sqlconta.response(6, 3)
    TIENEIHA(row) = sqlconta.response(7, 3)
    TIENEACTIVO(row) = sqlconta.response(8, 3)
    CUENTAMAYOR(row) = sqlconta.response(0, 3)
    If TIENECRCC(row) = "1" And col <> 9999 Then
        lin = row
        CRCC.Enabled = True
        DATO21.Enabled = True
        cabeza.Enabled = False
        detalle.Enabled = False
        DATO21.SetFocus
        DATO21.Tag = row
        
    End If

End If
End Sub

    

Sub leercrcc()
    campos(0, 0) = "codigo"
    campos(1, 0) = "nombre"
    campos(2, 0) = ""
    campos(0, 2) = "centrosdecosto"
    condicion = "codigo=" + "'" + DATO21.text + DATO22.text + "' and año='" + Format(fechasistema, "yyyy") + "'"
    op = 5
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)
    If sqlconta.status = 4 Or DATO22.text = "00" Then DATO21.text = "": DATO22.text = "": DATO21.SetFocus: GoTo no:
    DATO21.text = Mid(sqlconta.response(0, 3), 1, 2)
    DATO22.text = Mid(sqlconta.response(0, 3), 3, 2)
    nombrecrcc.Caption = sqlconta.response(1, 3)
    
'    If col <> 9999 Then
'    DATO21.text = ""
'    DATO22.text = ""
'
    'comentado 18-07-2014
     cabeza.Enabled = True
'    detalle.Enabled = True
'    grid1.Enabled = True
'    grid1.Cell(1, 1).SetFocus
    ' comentado 18-07-2014
    ' cambiado a txtitemfactura
    
    
      If Val(Mid(COMBOTIPO.text, 1, 3)) = 61 Or Val(Mid(COMBOTIPO.text, 1, 3)) = 56 Then
            frmdatos.Visible = True
            dato14.SetFocus
        Else
            txtitemfactura.SetFocus
            totalfactura
        End If
        
        
    
    
'  End If

no:
End Sub
Sub SUMAR()
Dim o As Integer
Dim sumadebe As Double
Dim sumahaber As Double

sumadebe = NETO
sumahaber = 0
SALDOPE = 0
For o = 1 To grid1.Rows - 1
If grid1.Cell(o, 6).text = "D" Then sumadebe = sumadebe + Val(grid1.Cell(o, 5).text)
If grid1.Cell(o, 6).text = "H" Then sumahaber = sumahaber + Val(grid1.Cell(o, 5).text)
Next o
debe.Caption = Format(sumadebe, "###,###,###,##0")
haber.Caption = Format(sumahaber, "###,###,###,##0")

saldo.Caption = Format(sumadebe - sumahaber, "###,###,###,##0")
SALDOPE = sumadebe - sumahaber
End Sub

Sub CARGAGRILLA(row, col)
    Rem DATOS DE LA COLUMNA
    FORMATOGRILLA(1, 1) = "DETALLE DE LA FACTURA"
    
    Rem LARGO DE LOS DATOS
    
    FORMATOGRILLA(2, 1) = "150"
    
    Rem TIPO DE DATOS C=CEROS,S=STRING,N=NUMERICO,D=DATE,
    FORMATOGRILLA(3, 1) = "S"
    
    Rem FORMATO GRILLA
    FORMATOGRILLA(4, 1) = ""
    Rem LOCCKED
    FORMATOGRILLA(5, 1) = "false"
    
    grid1.Cols = col
    grid1.Rows = row
    grid1.AllowUserResizing = False
    grid1.DisplayFocusRect = False
    grid1.ExtendLastCol = True
    grid1.BoldFixedCell = False
    grid1.DrawMode = cellOwnerDraw
    grid1.Appearance = Flat
    grid1.ScrollBarStyle = Flat
    grid1.FixedRowColStyle = Flat
    grid1.Column(0).Width = 0
    
    
    For k = 1 To col - 1
        grid1.Cell(0, k).text = FORMATOGRILLA(1, k)
        
        grid1.Column(k).MaxLength = Val(FORMATOGRILLA(2, k))
        grid1.Column(k).FormatString = FORMATOGRILLA(4, k)
        grid1.Column(k).Locked = FORMATOGRILLA(5, k)
        If FORMATOGRILLA(3, k) = "N" Then
            grid1.Column(k).Alignment = cellRightCenter
            grid1.Column(k).Mask = cellNumeric
        End If
        If FORMATOGRILLA(3, k) = "S" Then
            grid1.Column(k).Alignment = cellLeftCenter
''            Grid1.Column(k).Mask = cellUpper
'             Grid1.Column(k).Mask = cellLetter
        End If
        If FORMATOGRILLA(3, k) = "D" Then
            grid1.Column(k).CellType = cellCalendar
            grid1.Column(k).Mask = cellNumeric
        End If
        
        'Grid1.Column(7).CellType = cellComboBox
    Next k
    grid1.Range(0, 1, 0, 1).FontSize = 12
    grid1.Range(0, 1, 0, 1).FontBold = True
    grid1.Range(0, 1, 0, 1).Alignment = cellCenterCenter
    
    
    
    grid1.Column(1).Width = 600
    
End Sub


Private Sub DATO21_GotFocus()
DATO21.text = "01"
DATO22.text = "01"

Call cargatexto(DATO21)
End Sub
Private Sub dato22_GotFocus()
Call cargatexto(DATO22)
End Sub

Private Sub dato21_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF2 Then Call ayudacrcc(DATO21.Tag, 11)
    Call flechas(DATO21, DATO22, KeyCode)
no:
End Sub

Private Sub dato22_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Call flechas(DATO21, DATO22, KeyCode)
End Sub

Private Sub dato21_KeyPress(KeyAscii As Integer)
    snum = 0: KeyAscii = esNumero(KeyAscii)
    If KeyAscii = 13 Then Call ceros(DATO21): Call Pregunta(DATO21, DATO22)
End Sub

Private Sub dato22_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then
    Call ceros(DATO22)
    Call leercrcc
    
    
    End If
End Sub

Sub leecomprobante(tipo)
    Dim lin As Integer
    Dim TIPOFA As String
    Dim resultados As rdoResultset
    Dim csql As New rdoQuery
    Dim rut As String
     
    
        Set csql.ActiveConnection = contadb
        csql.sql = "SELECT tipo,numero,linea,cuentadelmayor,glosa,monto,dh,rutctacte,centrodecosto "
        If Val(Mid(COMBOTIPO.text, 1, 3)) = 46 Then ' FACTURA DE COMPRA
            csql.sql = csql.sql + "FROM facturasdecompras_detalle "
        Else
            csql.sql = csql.sql + "FROM facturasdeventas_detalle "
        End If
        csql.sql = csql.sql + "WHERE tipo='" + tipo + "' and numero='" & Format(txtfolio.text, "0000000000") & "' and rut='" & dato9.text & DV.Caption & "' order by linea"
        csql.Execute

        canli = 0
        If csql.RowsAffected > 0 Then
            Set resultados = csql.OpenResultset
            While Not resultados.EOF
             canli = canli + 1
                rut = resultados(2)
                
                grilladetalle(canli, 1) = Mid(resultados(3), 1, 2)
                grilladetalle(canli, 2) = Mid(resultados(3), 3, 2)
                grilladetalle(canli, 3) = Mid(resultados(3), 5, 4)
                grilladetalle(canli, 4) = resultados(4)
                grilladetalle(canli, 5) = resultados(5)
                grilladetalle(canli, 6) = resultados(6)
                grilladetalle(canli, 7) = ""
                grilladetalle(canli, 8) = ""
                grilladetalle(canli, 9) = ""
                grilladetalle(canli, 10) = resultados(7)
                grilladetalle(canli, 11) = resultados(8)
                
       
                resultados.MoveNext
            Wend
            cargadorcomprobante
            resultados.Close
            Set resultados = Nothing
        End If
    
'leerglosa
   If csql.RowsAffected > 0 Then opciones.Visible = True: grid1.Enabled = False: opciones.SetFocus

no:
End Sub
Sub leedetalle()
    Dim lin As Integer
    Dim TIPOFA As String
    Dim resultados As rdoResultset
    Dim csql As New rdoQuery
    Dim rut As String
     TIPOFA = Val(Mid(COMBOTIPO.text, 1, 3))
    
        Set csql.ActiveConnection = contadb
        csql.sql = "SELECT tipo,numero,linea,glosa "
        csql.sql = csql.sql + "FROM facturasvarias_glosa "
        csql.sql = csql.sql + "WHERE tipo='" + TIPOFA + "' and numero='" & dato2.text & "' order by linea"
        csql.Execute

        canli = 0
        grid1.Rows = 20
        If csql.RowsAffected > 0 Then
            Set resultados = csql.OpenResultset
            While Not resultados.EOF
             canli = canli + 1
                grid1.Cell(canli, 1).text = resultados(3)
                
                resultados.MoveNext
            Wend
            resultados.Close
            Set resultados = Nothing
        End If
    

no:
End Sub


Sub cargadorcomprobante()
    Dim LINEA As Long
    Grid3.AutoRedraw = False


    Grid3.Rows = canli + 2

    For k = 1 To canli
    CUENTAMAYOR(k) = grilladetalle(k, 1)
    Grid3.Cell(k, 1).text = grilladetalle(k, 1)
    Grid3.Cell(k, 2).text = grilladetalle(k, 2)
    Grid3.Cell(k, 3).text = grilladetalle(k, 3)
    Grid3.Cell(k, 4).text = grilladetalle(k, 4)
    Grid3.Cell(k, 5).text = grilladetalle(k, 5)
    Grid3.Cell(k, 6).text = grilladetalle(k, 6)
    Grid3.Cell(k, 7).text = ""
    Grid3.Cell(k, 8).text = ""
    Grid3.Cell(k, 9).text = ""

    Grid3.Cell(k, 10).text = grilladetalle(k, 10)
    Grid3.Cell(k, 11).text = grilladetalle(k, 11)

    DATO21.text = Mid(grilladetalle(k, 11), 1, 2)
    DATO22.text = Mid(grilladetalle(k, 11), 3, 2)

    LINEA = k

'    Call leermayor(linea, 9999)
'
'    If Val(grilladetalle(linea, 11)) <> 0 Then Call leercrcc(linea, 9999)

    

    Next k
    Grid3.AutoRedraw = True
    Grid3.Refresh

    
End Sub
                

Sub ayudacrcc(row As Long, col As Long)
    Dim campos As Variant
    Dim cfijo As Variant
    Dim largo As Variant
    campos = Array("codigo", "nombre")
    cabezas = Array("codigo", "nombre")
    largo = Array("8n", "40s")
    mensajeAyuda = "Ayuda Centros de costo"
    cfijo = "año='" + año + "'"
    pivote2.MaxLength = 4
    Call cargaAyudaT(Servidor, basebus, Usuario, password, "centrosdecosto", pivote2, campos, cfijo, largo, 2)
    DATO21.text = Mid(pivote2.text, 1, 2)
    DATO22.text = Mid(pivote2.text, 3, 2)
    
    pivote2.text = ""
End Sub

Sub modificafactura(tipo, numero, caja)
    Dim campos(10, 10) As String
    Dim condicion As String
    
    Dim netos As Double
    Dim DH As String
    campos(0, 0) = "caja"
    campos(1, 0) = ""
    campos(0, 1) = caja
    
    
    
    condicion = "tipo='" + tipo + "' and numero='" + numero + "' "
    campos(0, 2) = "facturasdeventas"
    op = 3
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)
    
End Sub

Sub CARGAGRILLA2(row, col)
    Rem DATOS DE LA COLUMNA
    FORMATOGRILLA(1, 1) = "C1"
    FORMATOGRILLA(1, 2) = "C2"
    FORMATOGRILLA(1, 3) = "C3"
    FORMATOGRILLA(1, 4) = "GLOSA"
    FORMATOGRILLA(1, 5) = "MONTO"
    FORMATOGRILLA(1, 6) = "D/H"
    FORMATOGRILLA(1, 7) = "MAYOR"
    FORMATOGRILLA(1, 8) = "CTACTE"
    FORMATOGRILLA(1, 9) = "CRCC"
    FORMATOGRILLA(1, 10) = "RUT"
    FORMATOGRILLA(1, 11) = "CRCC"
    
    Rem LARGO DE LOS DATOS
    
    FORMATOGRILLA(2, 1) = "2"
    FORMATOGRILLA(2, 2) = "2"
    FORMATOGRILLA(2, 3) = "4"
    FORMATOGRILLA(2, 4) = "60"
    FORMATOGRILLA(2, 5) = "12"
    FORMATOGRILLA(2, 6) = "3"
    FORMATOGRILLA(2, 7) = "15"
    FORMATOGRILLA(2, 8) = "15"
    FORMATOGRILLA(2, 9) = "15"
    FORMATOGRILLA(2, 10) = "10"
    FORMATOGRILLA(2, 11) = "4"
    
    Rem TIPO DE DATOS C=CEROS,S=STRING,N=NUMERICO,D=DATE,
    FORMATOGRILLA(3, 1) = "C"
    FORMATOGRILLA(3, 2) = "C"
    FORMATOGRILLA(3, 3) = "C"
    FORMATOGRILLA(3, 4) = "S"
    FORMATOGRILLA(3, 5) = "N"
    FORMATOGRILLA(3, 6) = "S"
    FORMATOGRILLA(3, 7) = "S"
    FORMATOGRILLA(3, 8) = "S"
    FORMATOGRILLA(3, 9) = "S"
    
    Rem FORMATO GRILLA
    FORMATOGRILLA(4, 1) = ""
    FORMATOGRILLA(4, 2) = ""
    FORMATOGRILLA(4, 3) = ""
    FORMATOGRILLA(4, 4) = ""
    FORMATOGRILLA(4, 5) = "$ ###,###,##0"
    FORMATOGRILLA(4, 6) = "H"
    FORMATOGRILLA(4, 7) = ""
    Rem LOCCKED
    FORMATOGRILLA(5, 1) = "FALSE"
    FORMATOGRILLA(5, 2) = "FALSE"
    FORMATOGRILLA(5, 3) = "FALSE"
    FORMATOGRILLA(5, 4) = "FALSE"
    FORMATOGRILLA(5, 5) = "FALSE"
    FORMATOGRILLA(5, 6) = "TRUE"
    FORMATOGRILLA(5, 7) = "TRUE"
    FORMATOGRILLA(5, 8) = "TRUE"
    FORMATOGRILLA(5, 9) = "TRUE"
    FORMATOGRILLA(5, 10) = "TRUE"
    FORMATOGRILLA(5, 11) = "TRUE"
    
    Grid3.Cols = col
    Grid3.Rows = row
    Grid3.AllowUserResizing = False
    Grid3.DisplayFocusRect = False
    Grid3.ExtendLastCol = True
    Grid3.BoldFixedCell = False
    Grid3.DrawMode = cellOwnerDraw
    Grid3.Appearance = Flat
    Grid3.ScrollBarStyle = Flat
    Grid3.FixedRowColStyle = Flat
    Grid3.Column(0).Width = 4 * 8.8
    Grid3.Column(1).Width = 2 * 10
    Grid3.Column(2).Width = 2 * 10
    Grid3.Column(3).Width = 4 * 10
    Grid3.Column(4).Width = 40 * 9
    Grid3.Column(5).Width = 12 * 9
    Grid3.Column(6).Width = 3 * 9
    Grid3.Column(7).Width = 100
    Grid3.Column(8).Width = 40
    Grid3.Column(9).Width = 100
    Grid3.Column(10).Width = 40
    
    
    For k = 1 To col - 1
        Grid3.Cell(0, k).text = FORMATOGRILLA(1, k)
        
        Grid3.Column(k).MaxLength = Val(FORMATOGRILLA(2, k))
        Grid3.Column(k).FormatString = FORMATOGRILLA(4, k)
        Grid3.Column(k).Locked = FORMATOGRILLA(5, k)
        If FORMATOGRILLA(3, k) = "N" Then
            Grid3.Column(k).Alignment = cellRightCenter
            Grid3.Column(k).Mask = cellNumeric
        End If
        If FORMATOGRILLA(3, k) = "S" Then
            Grid3.Column(k).Alignment = cellLeftCenter
            Grid3.Column(k).Mask = cellUpper
        End If
        If FORMATOGRILLA(3, k) = "D" Then
            Grid3.Column(k).CellType = cellCalendar
            Grid3.Column(k).Mask = cellNumeric
        End If
        
        'grid3.Column(7).CellType = cellComboBox
    Next k
    Grid3.Range(0, 1, 0, 3).Merge
    Grid3.Cell(0, 1).text = "CUENTA"
    Grid3.Range(0, 0, 0, Grid3.Cols - 1).Alignment = cellCenterCenter

End Sub

Public Sub imprime_factura()
    Dim SS As String
    Dim i As Integer
    Dim k As Integer
    Dim cad As String
    Dim totalprod As String
    Dim descuento As String
    Dim NETO As String
    Dim piva As String
    Dim piha As String
    Dim LINEAS As Double
    Dim fecha As String
    Dim vencimiento As String
    Dim Vendedor As String
    Dim notapedido As String
    Dim NOMBRE As String
    Dim rut As String
    Dim direccion As String * 30
    Dim ciudad As String
    Dim comuna As String * 40
    Dim giro As String * 35
    Dim fono As String * 10
    Dim o As Integer
    Dim dia As String
    Dim MES As String
    Dim ano As String
    Dim nvalor As Long
    Dim codigo As String
    Dim tiposDePago As String
    Dim razon As String
    Dim DIFE As Double
    Dim impuesto(10) As Double
    Dim TAZAS(10) As Double
    Dim EXENTO As Double
    Dim iva As Double
    Dim codigoimpuesto As String
    Dim tazaiva As Double
    Dim Precio As Double
    Dim netoprecio As Double
    Dim tazaimpuesto As Double
    Dim CAMPO1(10) As String * 15
    Dim TOTA1 As String
    Dim TOTA2 As String
    Dim TOTA3 As String
    Dim TOTA4 As String
    Dim TOTA5 As String
    Dim PALABRAIMPUESTO As String
    Dim esretenedor As Boolean
    
    Dim LINEA As Double
    
 
    Grid4.Rows = 1
    Grid4.Cols = 6
    Grid4.Rows = 64
   
    
    Grid4.DefaultFont.Size = 10
    Grid4.DefaultFont.Bold = False
    
    Grid4.Column(0).Width = 0
    Grid4.Column(1).Width = 160
    Grid4.Column(2).Width = 320
    Grid4.Column(3).Width = 50
    Grid4.Column(4).Width = 100
    Grid4.Column(5).Width = 100
    
    Grid4.Column(1).Alignment = cellLeftCenter
    Grid4.Column(2).Alignment = cellLeftCenter '/**/
    Grid4.Column(3).Alignment = cellCenterCenter
    Grid4.Column(4).Alignment = cellRightCenter
    Grid4.Column(5).Alignment = cellRightCenter

    'grid4.Column(7).Alignment = cellRightCenter
    '
    Grid4.DefaultRowHeight = 15
    Grid4.PageSetup.PrintGridlines = False
    Grid4.AutoRedraw = False
        
    'CABEZA
    
       NOMBRE = nombreproveedor.Caption
       rut = dato9.text + "-" + DV.Caption
       direccion = lbldireccion.Caption
       ciudad = lblciudad.Caption
       comuna = lblcomuna.Caption
       giro = lblgiro.Caption
       razon = nombreproveedor.Caption
    
        
         Grid4.Cell(3, 5).text = dato2.text
        Grid4.Range(9, 2, 9, 3).Merge
        Grid4.Range(9, 2, 9, 3).Alignment = cellCenterCenter
        Grid4.Cell(9, 2).text = nombreempresa
 
    
 '    'SEÑORES
        Grid4.Range(12, 1, 12, 2).Merge
        Grid4.Range(12, 1, 12, 2).Alignment = cellLeftCenter
        Grid4.Cell(12, 1).text = "      " & razon
        
     'RUT
        Grid4.Cell(12, 3).Alignment = cellLeftCenter
        Grid4.Range(12, 3, 12, 4).Merge
        Grid4.Cell(12, 3).text = "     " + Format(Left(rut, 9), "###,###,###") & "-" & Right(rut, 1)
 
'    'SEÑORES
'       Grid4.Range(9, 2, 9, 3).Merge
'       Grid4.Range(9, 2, 9, 3).Alignment = cellLeftCenter
'       Grid4.Cell(9, 2).text = razon
    
    ' fecha
'       FECHA = dato3.text + "-" + dato4.text + "-" + DATO5.text
''      Grid4.Cell(9, 5).text = fecha
'       Grid4.Cell(6, 5).text = FECHA
       
       fecha = "   " & dato3.text + "   " + MonthName(dato4.text) + "       " + Mid(dato5.text, 3, 2)
       Grid4.Range(9, 4, 9, 5).Merge
        Grid4.Cell(9, 4).text = fecha
    
'    'DIRECCION
'       Grid4.Range(10, 1, 10, 5).Merge
'       Grid4.Range(10, 1, 10, 5).Alignment = cellLeftCenter
'       Grid4.Cell(10, 1).text = direccion & "              " & ciudad & "              " & giro & "        " & fono
      
       'DIRECCION 'COMUNA'GIRO'FONO
    Grid4.Range(15, 1, 15, 5).Merge
    Grid4.Range(15, 1, 15, 5).Alignment = cellLeftCenter
    Grid4.Cell(15, 1).text = "      " & direccion & comuna & giro & fono
      
'    'GIRO
'      Grid4.Range(12, 2, 12, 3).Merge
'      Grid4.Range(12, 2, 12, 3).Alignment = cellLeftCenter
'      Grid4.Cell(12, 2).text = giro
    
    'CIUDAD
    
''     Grid4.Cell(12, 5).Alignment = cellLeftCenter
'      Grid4.Cell(12, 5).text = "     " + ciudad
    'DETALLE
        
        LINEAS = 19
        LINEA = 0
        EXENTO = 0
        NETO = 0
        total = CDbl(total.text)
        EXENTO = CDbl(dato13.text)
        NETO = CDbl(dato11.text)
        
        For k = 1 To grid1.Rows - 1
            LINEA = LINEA + 1
            LINEAS = LINEAS + 1
            Grid4.Cell(LINEAS, 2).text = grid1.Cell(k, 1).text
            If LINEA = 1 Then
'                Grid4.Cell(LINEAS, 3).text = "1"
                Grid4.Cell(LINEAS, 4).text = Format(NETO, " $ ###,###,###")
                Grid4.Cell(LINEAS, 5).text = Format(NETO, " $ ###,###,###")
            End If
        Next k
    
        
    
    
    
    
    
    
    iva = CDbl(dato12.text)
'    Grid4.Cell(49, 4).Alignment = cellLeftCenter
'    Grid4.Cell(49, 4).text = "  NETO"
    Grid4.Cell(43, 5).text = Format(NETO, "###,###,##0")
'    Grid4.Cell(50, 4).Alignment = cellLeftCenter
'    Grid4.Cell(50, 4).text = "  IVA"
    Grid4.Cell(45, 5).text = Format(iva, "###,###,##0")
'    Grid4.Cell(51, 4).Alignment = cellLeftCenter
'    Grid4.Cell(51, 4).text = "  OTROS IMPUESTOS"
'    Grid4.Cell(51, 5).text = Format(EXENTO, "###,###,##0")
'    Grid4.Cell(52, 4).Alignment = cellLeftCenter
'    Grid4.Cell(52, 4).text = "  TOTAL"
    Grid4.Cell(47, 5).text = Format(total, "###,###,##0")
    nvalor = Format(CDbl(total.text))
    SS = Numero_Texto(nvalor)
'    Grid4.Range(51, 1, 51, 3).Merge
'    Grid4.Range(51, 1, 51, 3).Alignment = cellLeftCenter
'    Grid4.Cell(51, 1).text = "        " + SS
    
    
        Grid4.Range(49, 1, 50, 3).Merge
        Grid4.Range(49, 1, 50, 3).Alignment = cellLeftCenter
        Grid4.Cell(49, 1).text = "                    " + SS
        Grid4.Range(49, 1, 50, 3).WrapText = True
    
'    Grid4.Range(50, 1, 50, 3).WrapText = True
'    Grid4.Range(56, 2, 56, 5).Merge
'    Grid4.Range(56, 2, 56, 5).Alignment = cellLeftCenter
'    Grid4.Range(56, 2, 56, 5).FontBold = True
'    Grid4.Range(57, 2, 57, 5).Merge
    Grid4.AutoRedraw = True
    Grid4.Refresh
    
    Grid4.PageSetup.LeftMargin = 0.25
    Grid4.PageSetup.RightMargin = 0
    Grid4.PageSetup.TopMargin = 2.5
    Grid4.PageSetup.BottomMargin = 0
    
'    For i = 1 To Grid4.PageSetup.PaperSizes.Count
'        If UCase(Grid4.PageSetup.PaperSizes.Item(i).PaperName) = "CARTA" Then
'            Grid4.PageSetup.PaperSize = Grid4.PageSetup.PaperSizes.Item(i).Kind
'            Exit For
'        End If
'    Next i
    
    Grid4.PageSetup.PrintGridlines = False
    'grid4.DirectPrint
    Grid4.PrintPreview
End Sub


Public Function Numero_Texto(nvalor As Long) As String
    
    Dim Mon_Esc, QueES As String
    Dim k As String
    ReDim uni(15) As String
    ReDim Dec(9) As String
    Dim Z, Num, var As Variant
    Dim c, D, U, v, i As Integer
    Dim textnum As Long
    If Len(nvalor) = 0 Then                        'Si no se ingresa Valor se Devuelve Vacío
        textnum = "": Exit Function
    End If
    If nvalor = 0 Or nvalor > 1E+17 Then
       Mon_Esc = IIf(nvalor = 0, "CERO", "*")
    End If
    ' ------------ UNIDADES ----------------------------------
    uni(1) = "UN"
    uni(2) = "DOS"
    uni(3) = "TRES"
    uni(4) = "CUATRO"
    uni(5) = "CINCO"
    uni(6) = "SEIS"
    uni(7) = "SIETE"
    uni(8) = "OCHO"
    uni(9) = "NUEVE"
    uni(10) = "DIEZ"
    uni(11) = "ONCE"
    uni(12) = "DOCE"
    uni(13) = "TRECE"
    uni(14) = "CATORCE"
    uni(15) = "QUINCE"
    ' ------------ DECENAS ----------------------------------
    Dec(3) = "TREINTA"
    Dec(4) = "CUARENTA"
    Dec(5) = "CINCUENTA"
    Dec(6) = "SESENTA"
    Dec(7) = "SETENTA"
    Dec(8) = "OCHENTA"
    Dec(9) = "NOVENTA"
    
    Num = String$(19 - Len(Str(Trim(nvalor))), Space(1))
    Num = Num + Trim(Str(nvalor))
    i = 1
    Z = ""
    
    Do While True
       k = Mid(Num, 18 - (i * 3 - 1), 3)
    
       If k = Space(3) Then
          Exit Do
       End If
    
       c = Val(Mid(k, 1, 1))
       D = Val(Mid(k, 2, 1))
       U = Val(Mid(k, 3, 1))
       v = Val(Mid(k, 2, 2))
    
       If i > 1 Then
          If (i = 2 Or i = 4) And Val(k) > 0 Then
             Z = " MIL " + Z
          End If
          If i = 3 And Val(Mid(Num, 7, 6)) > 0 Then
             If Val(k) = 1 Then
                Z = " MILLON " + Z
             Else
                Z = " MILLONES " + Z
             End If
          End If
          If i = 5 And Val(k) > 0 Then
             If Val(k) = 1 Then
                Z = " BILLON " + Z
             Else
                Z = " BILLONES " + Z
             End If
          End If
       End If
    
       If v > 0 Then
          Select Case v
                 Case 0 To 15
                      Z = uni(v) + Z
                 Case 0 To 19
                      Z = " DIECI" + uni(U) + Z
                 Case 20
                      Z = " VEINTE " + Z
                 Case 0 To 29
                      Z = " VEINTI" + uni(U) + Z
                 Case Else
                      If U = 0 Then
                         Z = Dec(D) + Z
                      Else
                         Z = Dec(D) + " Y " + uni(U) + Z
                      End If
          End Select
       End If
    
       If c > 0 Then
          If c = 1 Then
             If v = 0 Then
                Z = " CIEN " + Z
             Else
                Z = " CIENTO " + Z
             End If
          End If
          If c = 2 Or c = 3 Or c = 4 Or c = 6 Or c = 8 Then
             Z = uni(c) + "CIENTOS " + Z
          End If
          If c = 5 Then
             Z = " QUINIENTOS " + Z
          End If
          If c = 7 Then
             Z = " SETECIENTOS " + Z
          End If
          If c = 9 Then
             Z = " NOVECIENTOS " + Z
          End If
       End If
    
       i = i + 1
    Loop
    
    Mon_Esc = Trim(Z)
    ' CAMBIA "UNO MIL ..." POR "MIL..."
    If Mid(Mon_Esc, 1, 7) = "UN MIL " Then
        Mon_Esc = "MIL " + Trim(Mid(Mon_Esc, 7, Len(Mon_Esc)))
    End If
    Numero_Texto = Mon_Esc + " PESOS" + QueES
End Function

Private Sub Text1_Change()

End Sub

Private Sub txtitemfactura_GotFocus()
    Call cargatexto(txtitemfactura)
End Sub

Private Sub txtitemfactura_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        If Val(dato13.text) > 0 And Mid(COMBOTIPO.text, 1, 3) = "033" Then
            frmexento.Visible = True
            txtitemfacturaEXENTO.SetFocus
        Else
            grid1.Enabled = True
            grid1.Cell(1, 1).SetFocus
        End If
        
    End If
End Sub

Private Sub txtitemfacturaEXENTO_GotFocus()
    Call cargatexto(txtitemfactura)
End Sub

Private Sub txtitemfacturaEXENTO_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then
            grid1.Enabled = True
            grid1.Cell(1, 1).SetFocus
    End If
End Sub
Function LeerUltimo(tipo) As String
    Dim csql As New rdoQuery
    Dim resultados As rdoResultset
    Set csql.ActiveConnection = contadb
    csql.sql = "select ifnull(max(numero),0)+1 from facturasvarias "
    csql.sql = csql.sql & " where tipo='" & tipo & "' and fecha>='2014-07-04' "
    csql.sql = csql.sql & "group by tipo "
    csql.Execute
    LeerUltimo = "0000000001"
    If csql.RowsAffected > 0 Then
        Set resultados = csql.OpenResultset
        LeerUltimo = Format(resultados(0), "0000000000")
    End If
    csql.Close
    Set csql = Nothing
    
End Function


'CONFI_EMPRESAFAE

Function LeedatoOCproveedor(tipo, numero, loc, campo, cajadedocumento) As String
    Dim csql As New rdoQuery
    Dim resultados As rdoResultset
    Set csql.ActiveConnection = contadb
    csql.sql = "select " & campo & " from " & clientesistema & "ventas" & loc & ".sv_otros_documento_cabeza_" & loc
    csql.sql = csql.sql & " where tipo='" & tipo & "' and numero='" & numero & "' and caja='" & cajadedocumento & "'"
    csql.Execute
    LeedatoOCproveedor = ""
    If csql.RowsAffected > 0 Then
        Set resultados = csql.OpenResultset
        LeedatoOCproveedor = resultados(0)
    End If
    csql.Close
    Set csql = Nothing
    If campo = "ifnull(fechaoc,'0000-00-00')" And (LeedatoOCproveedor = "" Or LeedatoOCproveedor = "0000-00-00") Then LeedatoOCproveedor = Now()
End Function



