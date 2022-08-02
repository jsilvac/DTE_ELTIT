VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{4777436C-EB8C-4596-98A8-EBCF98683969}#1.0#0"; "FlexCell.ocx"
Object = "{ADBBAED6-B16F-40DC-80DF-B44910CBA76C}#1.0#0"; "Frame-Xp.ocx"
Begin VB.Form ingreso01 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Ingreso de Comprobantes Contables"
   ClientHeight    =   9645
   ClientLeft      =   45
   ClientTop       =   1305
   ClientWidth     =   15060
   FillColor       =   &H00FF8080&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9645
   ScaleWidth      =   15060
   ShowInTaskbar   =   0   'False
   Begin XPFrame.FrameXp frmimprimir 
      Height          =   4335
      Left            =   765
      TabIndex        =   33
      Top             =   3660
      Width           =   3480
      _ExtentX        =   6138
      _ExtentY        =   7646
      BackColor       =   16744576
      Caption         =   "IMPRESION DE COMPROBANTE"
      CaptionEstilo3D =   1
      BackColor       =   16744576
      ForeColor       =   8438015
      BordeColor      =   -2147483635
      ColorBarraArriba=   4194304
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.CommandButton Command7 
         BackColor       =   &H000080FF&
         Caption         =   "Cheque Nuevo Formato Santander"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   64
         Top             =   2880
         Width           =   3075
      End
      Begin VB.CommandButton Command6 
         BackColor       =   &H000080FF&
         Caption         =   "Cheque Formato Estandar"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   63
         Top             =   3360
         Width           =   3075
      End
      Begin VB.CommandButton Command10 
         BackColor       =   &H000080FF&
         Caption         =   "Cheque Suelto Estandar"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   62
         Top             =   3840
         Width           =   3075
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Cheque Banco &Santander Masivo"
         Height          =   375
         Left            =   240
         TabIndex        =   48
         Top             =   2280
         Width           =   3075
      End
      Begin VB.CommandButton Command1 
         Caption         =   "NO IMPRIMIR"
         Height          =   375
         Left            =   225
         TabIndex        =   38
         Top             =   1755
         Width           =   3075
      End
      Begin VB.CommandButton imprime3 
         Caption         =   "&Solo Comprobante "
         Height          =   375
         Left            =   225
         TabIndex        =   36
         Top             =   1305
         Width           =   3075
      End
      Begin VB.CommandButton imprime2 
         Caption         =   "Cheque Suelto"
         Height          =   375
         Left            =   225
         TabIndex        =   35
         Top             =   855
         Width           =   3075
      End
      Begin VB.CommandButton imprime1 
         Caption         =   "Cheque Banco &Santander"
         Height          =   375
         Left            =   240
         TabIndex        =   34
         Top             =   405
         Width           =   3075
      End
   End
   Begin XPFrame.FrameXp frmImportar 
      Height          =   1455
      Left            =   8160
      TabIndex        =   55
      Top             =   0
      Visible         =   0   'False
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   2566
      BackColor       =   16744576
      Caption         =   "OPCIONES DE IMPORTACION"
      CaptionEstilo3D =   1
      BackColor       =   16744576
      ForeColor       =   8438015
      BordeColor      =   -2147483635
      ColorBarraArriba=   4194304
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
      Begin Contabilidadxp.BotonMyERP CmdImportar 
         Height          =   255
         Left            =   720
         TabIndex        =   61
         Top             =   1080
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   450
         Caption         =   "IMPORTAR CSV"
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Contabilidadxp.BotonMyERP BotonMyERP1 
         Height          =   255
         Left            =   3120
         TabIndex        =   60
         Top             =   1080
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   450
         Caption         =   "CANCELAR"
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.CheckBox invertir 
         BackColor       =   &H00FF8080&
         Caption         =   "INVERTIR CUENTAS"
         Height          =   255
         Left            =   3120
         TabIndex        =   58
         Top             =   360
         Width           =   2775
      End
      Begin VB.OptionButton delimitador 
         BackColor       =   &H00FF8080&
         Caption         =   "USAR COMO DELIMITADOR  "","""
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   57
         Top             =   720
         Value           =   -1  'True
         Width           =   2775
      End
      Begin VB.OptionButton delimitador 
         BackColor       =   &H00FF8080&
         Caption         =   "USARCOMO  DELIMITADOR "";"""
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   56
         Top             =   360
         Width           =   2775
      End
   End
   Begin FlexCell.Grid grillatraspaso 
      Height          =   255
      Left            =   7800
      TabIndex        =   54
      Top             =   8280
      Visible         =   0   'False
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   450
      Cols            =   5
      DefaultFontSize =   8.25
      Rows            =   30
   End
   Begin VB.CommandButton Command4 
      Caption         =   "TRASPASAR A EMPRESA RELACIONADA"
      Height          =   375
      Left            =   10800
      TabIndex        =   53
      Top             =   8160
      Width           =   3675
   End
   Begin XPFrame.FrameXp frmcorrelativo 
      Height          =   2535
      Left            =   4440
      TabIndex        =   39
      Top             =   2520
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   4471
      BackColor       =   16744576
      Caption         =   "IMPRESION MASIVA DE CHEQUES"
      CaptionEstilo3D =   2
      BackColor       =   16744576
      ForeColor       =   8438015
      BordeColor      =   -2147483635
      ColorBarraArriba=   16744576
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.OptionButton Option2 
         BackColor       =   &H00FF8080&
         Caption         =   "FORMATO NUEVO"
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
         Left            =   2880
         TabIndex        =   66
         Top             =   1320
         Width           =   2295
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00FF8080&
         Caption         =   "FORMATO ANTIGUO"
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
         Left            =   2880
         TabIndex        =   65
         Top             =   960
         Value           =   -1  'True
         Width           =   2295
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00FF8080&
         Caption         =   "Vista Previa"
         Height          =   375
         Left            =   240
         TabIndex        =   49
         Top             =   1920
         Width           =   1215
      End
      Begin VB.CommandButton Command3 
         Caption         =   "IMPRIMIR AHORA"
         Height          =   375
         Left            =   1800
         TabIndex        =   47
         Top             =   1920
         Width           =   3075
      End
      Begin VB.TextBox Text3 
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
         Left            =   1320
         MaxLength       =   10
         TabIndex        =   45
         Top             =   1200
         Width           =   1425
      End
      Begin VB.TextBox Text2 
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
         Left            =   1320
         MaxLength       =   10
         TabIndex        =   43
         Top             =   840
         Width           =   1425
      End
      Begin VB.TextBox Text1 
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
         Left            =   720
         MaxLength       =   2
         TabIndex        =   40
         Top             =   480
         Width           =   420
      End
      Begin VB.Label Label8 
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " HASTA"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   120
         TabIndex        =   46
         Top             =   1200
         Width           =   1095
      End
      Begin VB.Label Label4 
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " DESDE"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   120
         TabIndex        =   44
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " TIPO"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   120
         TabIndex        =   42
         Top             =   480
         Width           =   495
      End
      Begin VB.Label Label2 
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1200
         TabIndex        =   41
         Top             =   480
         Width           =   3855
      End
   End
   Begin XPFrame.FrameXp MODIFICANDO 
      Height          =   855
      Left            =   900
      TabIndex        =   31
      Top             =   8235
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   1508
      BackColor       =   49344
      Caption         =   "MODIFICACION"
      CaptionEstilo3D =   1
      BackColor       =   49344
      ForeColor       =   8438015
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
      Begin VB.CommandButton BOTONMODIFICA 
         Caption         =   "FIN MODIFICACION"
         Height          =   375
         Left            =   120
         TabIndex        =   32
         Top             =   360
         Width           =   4575
      End
   End
   Begin XPFrame.FrameXp glosafinal 
      Height          =   2535
      Left            =   3720
      TabIndex        =   28
      Top             =   5160
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   4471
      BackColor       =   16761024
      Caption         =   "Glosas Comprobante"
      CaptionEstilo3D =   1
      BackColor       =   16761024
      ForeColor       =   8438015
      BordeColor      =   -2147483635
      ColorBarraArriba=   16744576
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.CommandButton grabar2 
         Caption         =   "Grabar"
         Height          =   375
         Left            =   1800
         TabIndex        =   30
         Top             =   2040
         Width           =   2175
      End
      Begin VB.TextBox glosa 
         BackColor       =   &H00E0E0E0&
         Height          =   1575
         Left            =   120
         TabIndex        =   29
         Top             =   360
         Width           =   5535
      End
   End
   Begin XPFrame.FrameXp FrameXp2 
      Height          =   975
      Left            =   7800
      TabIndex        =   21
      Top             =   8640
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   1720
      BackColor       =   16744576
      Caption         =   "VALORES DEL COMPROBANTE"
      CaptionEstilo3D =   2
      BackColor       =   16744576
      ForeColor       =   8438015
      BordeColor      =   -2147483635
      ColorBarraArriba=   16744576
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
      Begin XPFrame.FrameXp FrameXp3 
         Height          =   615
         Left            =   480
         TabIndex        =   22
         Top             =   240
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   1085
         BackColor       =   16761024
         Caption         =   "DEBE"
         CaptionEstilo3D =   2
         BackColor       =   16761024
         ForeColor       =   8438015
         BordeColor      =   -2147483635
         ColorBarraArriba=   16744576
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
            TabIndex        =   25
            Top             =   240
            Width           =   1575
         End
      End
      Begin XPFrame.FrameXp FrameXp4 
         Height          =   615
         Left            =   2760
         TabIndex        =   23
         Top             =   240
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   1085
         BackColor       =   16761024
         Caption         =   "HABER"
         CaptionEstilo3D =   2
         BackColor       =   16761024
         ForeColor       =   8438015
         BordeColor      =   -2147483635
         ColorBarraArriba=   16744576
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
            Left            =   120
            TabIndex        =   26
            Top             =   240
            Width           =   1695
         End
      End
      Begin XPFrame.FrameXp FrameXp5 
         Height          =   615
         Left            =   4920
         TabIndex        =   24
         Top             =   240
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   1085
         BackColor       =   16761024
         Caption         =   "SALDO"
         CaptionEstilo3D =   2
         BackColor       =   16761024
         ForeColor       =   8438015
         BordeColor      =   -2147483635
         ColorBarraArriba=   16744576
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
            TabIndex        =   27
            Top             =   240
            Width           =   1695
         End
      End
   End
   Begin XPFrame.FrameXp comprobante 
      Height          =   7095
      Left            =   0
      TabIndex        =   18
      Top             =   960
      Width           =   15015
      _ExtentX        =   26485
      _ExtentY        =   12515
      BackColor       =   16761024
      Caption         =   "Comprobante Contable"
      CaptionEstilo3D =   2
      BackColor       =   16761024
      ForeColor       =   8438015
      BordeColor      =   -2147483635
      ColorBarraArriba=   16744576
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin FlexCell.Grid Grid1 
         Height          =   6735
         Left            =   0
         TabIndex        =   19
         Top             =   240
         Width           =   15015
         _ExtentX        =   26485
         _ExtentY        =   11880
         BackColor1      =   16761024
         BackColor2      =   16761024
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
      Begin FlexCell.Grid Grid2 
         Height          =   375
         Left            =   3240
         TabIndex        =   20
         Top             =   1080
         Visible         =   0   'False
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   661
         Cols            =   5
         DefaultFontSize =   8.25
         Rows            =   30
      End
   End
   Begin XPFrame.FrameXp cabeza 
      Height          =   915
      Left            =   0
      TabIndex        =   13
      Top             =   0
      Width           =   15015
      _ExtentX        =   26485
      _ExtentY        =   1614
      BackColor       =   16744576
      Caption         =   "Datos del Comprobante"
      CaptionEstilo3D =   2
      BackColor       =   16744576
      ForeColor       =   8438015
      BordeColor      =   -2147483635
      ColorBarraArriba=   16744576
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.CommandButton Command5 
         Caption         =   "IMPORTAR"
         Height          =   255
         Left            =   10920
         TabIndex        =   59
         Top             =   360
         Width           =   1095
      End
      Begin MSComDlg.CommonDialog archivo 
         Left            =   12240
         Top             =   720
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.TextBox DATO1 
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
         Left            =   6600
         MaxLength       =   10
         TabIndex        =   1
         Top             =   360
         Width           =   1425
      End
      Begin VB.TextBox DATO4 
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
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   10200
         MaxLength       =   4
         TabIndex        =   4
         Tag             =   "fecha"
         Top             =   360
         Width           =   615
      End
      Begin VB.TextBox DATO3 
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
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   9840
         MaxLength       =   2
         TabIndex        =   3
         Tag             =   "fecha"
         Top             =   360
         Width           =   375
      End
      Begin VB.TextBox DATO2 
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
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   9480
         MaxLength       =   2
         TabIndex        =   2
         Tag             =   "fecha"
         Top             =   360
         Width           =   375
      End
      Begin VB.TextBox dato0 
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
         Left            =   720
         MaxLength       =   2
         TabIndex        =   0
         Top             =   360
         Width           =   420
      End
      Begin XPFrame.FrameXp FrameXp1 
         Height          =   615
         Left            =   12120
         TabIndex        =   50
         Top             =   120
         Width           =   2895
         _ExtentX        =   5106
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
         Begin VB.CommandButton botonmisaccesos 
            Caption         =   "Permisos Modulo"
            Height          =   255
            Left            =   120
            TabIndex        =   52
            Top             =   280
            Width           =   1455
         End
         Begin VB.CommandButton botonmisfavoritos 
            Caption         =   "Mis Favoritos"
            Height          =   255
            Left            =   1560
            TabIndex        =   51
            Top             =   280
            Width           =   1215
         End
      End
      Begin VB.Label centro 
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
         Height          =   375
         Left            =   11040
         TabIndex        =   37
         Top             =   360
         Width           =   3315
      End
      Begin VB.Label Label3 
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " NUMERO"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   5280
         TabIndex        =   17
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label6 
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " FECHA :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   8400
         TabIndex        =   16
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label5 
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " TIPO"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   120
         TabIndex        =   15
         Top             =   360
         Width           =   495
      End
      Begin VB.Label tipocompro 
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1200
         TabIndex        =   14
         Top             =   360
         Width           =   3855
      End
   End
   Begin VB.TextBox general 
      Height          =   285
      Left            =   6720
      TabIndex        =   12
      Text            =   "Text1"
      Top             =   9120
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox pivote4 
      Height          =   285
      Left            =   0
      TabIndex        =   11
      Top             =   1080
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox pivote3 
      Height          =   285
      Left            =   0
      MaxLength       =   12
      TabIndex        =   10
      Top             =   1440
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox pivote2 
      Height          =   375
      Left            =   0
      MaxLength       =   8
      TabIndex        =   9
      Top             =   1800
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox lineas 
      Height          =   285
      Left            =   0
      MaxLength       =   3
      TabIndex        =   6
      Text            =   "Text1"
      Top             =   2400
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox PIVOTE 
      Height          =   285
      Left            =   6480
      MaxLength       =   13
      TabIndex        =   5
      Top             =   9090
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.PictureBox MANUAL 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   0
      ScaleHeight     =   255
      ScaleWidth      =   855
      TabIndex        =   8
      Top             =   0
      Width           =   855
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash opciones 
      Height          =   1215
      Left            =   120
      TabIndex        =   7
      Top             =   8400
      Width           =   6735
      _cx             =   11880
      _cy             =   2143
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
      DeviceFont      =   "0"
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
Attribute VB_Name = "ingreso01"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Private digita As Integer
    Private lin As Double
    Private fechacon As Date
    Private PAGARFACTURA As Boolean
    Private rutprove As String
    Private VISTA As Boolean
    
    Private MODIFI As Integer
    Private FORMATOGRILLA(10, 20) As String
    Private sg As String
    Private tipoctacte As String
    Private sumadebe As Double
    Private sumahaber As Double
    Private canli As Integer
    Private grilladetalle(10000, 13) As String
    Private o As Integer
    Private ef As String
    Private cc As String
    Private respu As String
    Private AUXILIAR(10000, 3) As String
    Private CUENTAMAYOR(10000)  As String
    Private cierrect As String
    Private TIENECTACTE(10000) As String
    Private TIENECRCC(10000) As String
    Private TIENEBANCO(10000) As String
    Private TIENEILA(10000) As String
    Private TIENEICA(10000) As String
    Private TIENEIHA(10000) As String
    Private TIENEACTIVO(10000) As String
    Private existe As String
    Private chequecreado As String
    Private facturano As String
    Private boletano As String
    Private fila As Long
    Private columna As Long
    Private vacio As Boolean

Private Sub BotonMyERP1_Click()
frmImportar.Visible = False
Command5.Visible = True
End Sub

Private Sub CmdImportar_Click()
frmImportar.Visible = False
If IsDate(dato2.text & "-" & dato3.text & "-" & dato4.text) = False Then
    dato2.text = Format(fechasistema, "dd")
    dato3.text = Format(fechasistema, "mm")
    dato4.text = Format(fechasistema, "yyyy")
End If
Call ImportaCSV
End Sub

Private Sub Command10_Click()
Grid2.ReportTitles.Clear
 
 Call ChequeEstandar2(1, True)
End Sub

Private Sub COMMAND2_Click()
frmcorrelativo.Visible = True
Text1.SetFocus

End Sub

Private Sub Command3_Click()
If Check1.Value = "1" Then
VISTA = True
Else
VISTA = False
End If

Dim j As Integer
If Text2.text <> "" And Text3.text <> "" Then
'desbordamiento
For j = CDbl(Text2.text) To CDbl(Text3.text)
dato0.text = Text1.text
dato1.text = Format(j, "0000000000")
leecomprobante
    If Option1.Value = True Then
        Call imprimir(1)
    End If
    If Option2.Value = True Then
        Call imprimirSTANDAR(1)
    End If
Next j
VISTA = True
frmcorrelativo.Visible = False
End If
End Sub

Private Sub BLANQUEA_Click()

End Sub

Private Sub BOTONMODIFICA_Click()
    If Val(saldo.Caption) = 0 Then
        MODIFICANDO.Visible = False
        graba
        
        
    End If
End Sub

Private Sub Command1_Click()
frmimprimir.Visible = False
retorno

dato0.SetFocus

End Sub

Private Sub crccok_Click()
    CABEZA.Enabled = True
    comprobante.Enabled = True
    
    CABEZA.Enabled = True
    comprobante.Enabled = True
    
    
    Grid1.Cell(Grid1.ActiveCell.row, 5).SetFocus

End Sub

Private Sub Command4_Click()
      Dim nuevaempresa As String
      Dim rutprove As String
      Dim i As Double
      Dim tiponuevo As String
      
    If dato0.text = "PA" Or dato0.text = "CE" Or dato0.text = "NG" Then
        If dato0.text = "PA" Then tiponuevo = "TP"
        If dato0.text = "CE" Then tiponuevo = "TC"
        If dato0.text = "NG" Then tiponuevo = "TN"
        For i = 1 To Grid1.Rows - 1
          If Grid1.Cell(i, 13).text <> "" Then
              rutprove = Grid1.Cell(i, 13).text
              nuevaempresa = leerempresaproveedor(rutprove)
              Exit For
          End If
        Next i
        If nuevaempresa <> "" Then
            If verificacomprobante(nuevaempresa, empresaactiva & Mid(dato1, 3, 8), tiponuevo, dato4.text & "-" & dato3.text & "-" & dato2.text) = False Then
                Call generarcomprobante(empresaactiva, dato0.text, dato1.text, Grid1, dato4.text & "-" & dato3.text & "-" & dato2.text)
                MsgBox "GRABADO EXITOSAMENTE NUMERO " & empresaactiva & Mid(dato1, 3, 8) & " ", vbInformation, "ATENCION"
            Else
                MsgBox "COMPROBANTE " & empresaactiva & Mid(dato1, 3, 8) & " YA EXISTE", vbCritical, "ATENCION"
            End If
        Else
            MsgBox "PROVEEDOR NO ES EMPRESA RELACIONADA", vbCritical, "ATENCION"
        End If
    End If
End Sub

 
Private Sub Command5_Click()
If IsDate(dato2.text & "-" & dato3.text & "-" & dato4.text) = False Then
    dato2.text = Format(fechasistema, "dd")
    dato3.text = Format(fechasistema, "mm")
    dato4.text = Format(fechasistema, "yyyy")
End If
frmImportar.Visible = True

End Sub

Private Sub Command6_Click()
 Call ChequeEstandar2(1, False)
End Sub

Private Sub Command7_Click()
    Call imprimirSTANDAR(1)
End Sub

Private Sub dato0_GotFocus()
Call cargatexto(dato0)

End Sub

Private Sub dato0_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 38 Then Unload Me: GoTo no:
    If KeyCode = vbKeyF2 Then Call ayudatipos(dato0)
no:
End Sub

Private Sub text1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 38 Then Unload Me: GoTo no:
    If KeyCode = vbKeyF2 Then Call ayudatipos(Text1)
no:
End Sub

Private Sub dato1_GotFocus()
ultimo
Call cargatexto(dato1)
Rem DATO2.Enabled = True: DATO2.SetFocus
Call leetipos(dato0)

End Sub

Private Sub dato1_LostFocus()

leecomprobante
End Sub

Private Sub dato2_GotFocus()
    
    Call cargatexto(dato2)
End Sub

Private Sub DATO2_LostFocus()
    
    If dato2.text = "00" Then dato3.Enabled = True: dato4.Enabled = True: dato2.text = Format(fechasistema, "dd"): dato3.text = Format(fechasistema, "mm"): dato4.text = Format(fechasistema, "yyyy"): Grid1.Enabled = True: Grid1.Cell(1, 1).SetFocus

End Sub

Private Sub dato3_GotFocus()
    Call cargatexto(dato3)
End Sub
Private Sub dato4_GotFocus()
    
    Call cargatexto(dato4)
End Sub

Private Sub dato4_LostFocus()
    Call esfecha(dato2.text, dato3.text, dato4.text)
    If ef = "N" Then dato2.SetFocus
    If ef = "S" And MODIFI = 0 Then Call CARGAGRILLA(2, 16): Grid1.Enabled = True:
    'Grid1.Cell(1, 1).SetFocus
    

End Sub


Private Sub dato1_KeyDown(KeyCode As Integer, Shift As Integer)
    Call flechas(dato0, dato2, KeyCode)
End Sub

Private Sub dato2_KeyDown(KeyCode As Integer, Shift As Integer)
    Call flechas(dato1, dato3, KeyCode)
End Sub

Private Sub dato3_KeyDown(KeyCode As Integer, Shift As Integer)
    Call flechas(dato2, dato4, KeyCode)
End Sub

Private Sub dato4_KeyDown(KeyCode As Integer, Shift As Integer)
    Call flechas(dato3, dato4, KeyCode)
End Sub
Private Sub dato0_KeyPress(KeyAscii As Integer)
    
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 And frmimprimir.Visible = False Then
    
    Call Pregunta(dato0, dato1)
    End If
    
End Sub
Private Sub text1_KeyPress(KeyAscii As Integer)
    
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then
    
    Call Pregunta(Text1, Text2)
    Call leetipos2(Text1)
    End If
    
End Sub

Private Sub text2_KeyPress(KeyAscii As Integer)
    snum = 1: KeyAscii = esNumero(KeyAscii)
    If KeyAscii = 13 Then Call ceros(Text2): Call Pregunta(Text1, Text3)
End Sub

Private Sub text3_KeyPress(KeyAscii As Integer)
    snum = 1: KeyAscii = esNumero(KeyAscii)
    If KeyAscii = 13 Then
    Call ceros(Text3)
    If Text2.text > Text3.text Then
    Text2.SetFocus
    Else
    
    End If
    
    End If
    
End Sub

Private Sub dato1_KeyPress(KeyAscii As Integer)
    snum = 1: KeyAscii = esNumero(KeyAscii)
    If KeyAscii = 13 Then Call ceros(dato1): Call Pregunta(dato1, dato2)
End Sub

Private Sub dato2_KeyPress(KeyAscii As Integer)
    snum = 1: KeyAscii = esNumero(KeyAscii)
    If KeyAscii = 13 Then Call ceros(dato2): Call Pregunta(dato2, dato3)
End Sub

Private Sub dato3_KeyPress(KeyAscii As Integer)
    snum = 1: KeyAscii = esNumero(KeyAscii)
    If KeyAscii = 13 Then Call ceros(dato3): Call Pregunta(dato3, dato4)
End Sub
Private Sub dato4_KeyPress(KeyAscii As Integer)
    snum = 1: KeyAscii = esNumero(KeyAscii)

    If KeyAscii = 13 Then Grid1.Enabled = True: Grid1.Cell(1, 1).SetFocus
    
    
    
End Sub





Private Sub Form_Load()
    Call CENTRAR(Me)
    Call Conectar_BD
    Call Conectarconta(Servidor, clientesistema + "conta", Usuario, password)
    Dim SS As String
    Dim margen As Double
    Dim suma As Double
    Call CARGAGRILLA(2, 17)
    frmcorrelativo.Visible = False
    
    VISTA = True
    
    opciones.Visible = False
    Grid1.Enabled = False
    glosafinal.Visible = False
    

    MODIFICANDO.Visible = False
    frmimprimir.Visible = False
End Sub

Sub CARGAGRILLA(row, col)
    Rem DATOS DE LA COLUMNA
    FORMATOGRILLA(1, 1) = "C1"
    FORMATOGRILLA(1, 2) = "C2"
    FORMATOGRILLA(1, 3) = "C3"
    FORMATOGRILLA(1, 4) = "GLOSA"
    FORMATOGRILLA(1, 5) = "TP"
    FORMATOGRILLA(1, 6) = "NUMERO"
    FORMATOGRILLA(1, 7) = "F.VENCI."
    FORMATOGRILLA(1, 8) = "MONTO"
    FORMATOGRILLA(1, 9) = "D/H"
    FORMATOGRILLA(1, 10) = "MAYOR"
    FORMATOGRILLA(1, 11) = "CTACTE"
    FORMATOGRILLA(1, 12) = "CRCC"
    FORMATOGRILLA(1, 13) = "RUT"
    FORMATOGRILLA(1, 14) = "CRCC"
    FORMATOGRILLA(1, 15) = "CENTRO GASTO"
    FORMATOGRILLA(1, 16) = "ANALISIS"
    
    Rem LARGO DE LOS DATOS
    
    FORMATOGRILLA(2, 1) = "2"
    FORMATOGRILLA(2, 2) = "2"
    FORMATOGRILLA(2, 3) = "4"
    FORMATOGRILLA(2, 4) = "40"
    FORMATOGRILLA(2, 5) = "2"
    FORMATOGRILLA(2, 6) = "10"
    FORMATOGRILLA(2, 7) = "10"
    FORMATOGRILLA(2, 8) = "12"
    FORMATOGRILLA(2, 9) = "3"
    FORMATOGRILLA(2, 10) = "15"
    FORMATOGRILLA(2, 11) = "15"
    FORMATOGRILLA(2, 12) = "15"
    FORMATOGRILLA(2, 13) = "10"
    FORMATOGRILLA(2, 14) = "4"
    FORMATOGRILLA(2, 15) = "20"
    
    Rem TIPO DE DATOS C=CEROS,S=STRING,N=NUMERICO,D=DATE,
    FORMATOGRILLA(3, 1) = "C"
    FORMATOGRILLA(3, 2) = "C"
    FORMATOGRILLA(3, 3) = "C"
    FORMATOGRILLA(3, 4) = "S"
    FORMATOGRILLA(3, 5) = "S"
    FORMATOGRILLA(3, 6) = "C"
    FORMATOGRILLA(3, 7) = "D"
    FORMATOGRILLA(3, 8) = "N"
    FORMATOGRILLA(3, 9) = "S"
    FORMATOGRILLA(3, 10) = "S"
    FORMATOGRILLA(3, 11) = "S"
    FORMATOGRILLA(3, 12) = "S"
    FORMATOGRILLA(3, 13) = "S"
    FORMATOGRILLA(3, 14) = "S"
    FORMATOGRILLA(3, 15) = "S"
    FORMATOGRILLA(3, 16) = "S"
    
    Rem FORMATO GRILLA
    FORMATOGRILLA(4, 1) = ""
    FORMATOGRILLA(4, 2) = ""
    FORMATOGRILLA(4, 3) = ""
    FORMATOGRILLA(4, 4) = ""
    FORMATOGRILLA(4, 5) = ""
    FORMATOGRILLA(4, 6) = ""
    FORMATOGRILLA(4, 7) = ""
    FORMATOGRILLA(4, 8) = "$ ###,###,##0"
    FORMATOGRILLA(4, 9) = ""
    FORMATOGRILLA(4, 10) = ""
    Rem LOCCKED
    FORMATOGRILLA(5, 1) = "FALSE"
    FORMATOGRILLA(5, 2) = "FALSE"
    FORMATOGRILLA(5, 3) = "FALSE"
    FORMATOGRILLA(5, 4) = "FALSE"
    FORMATOGRILLA(5, 5) = "FALSE"
    FORMATOGRILLA(5, 6) = "FALSE"
    FORMATOGRILLA(5, 7) = "FALSE"
    FORMATOGRILLA(5, 8) = "FALSE"
    FORMATOGRILLA(5, 9) = "FALSE"
    FORMATOGRILLA(5, 10) = "TRUE"
    FORMATOGRILLA(5, 11) = "TRUE"
    FORMATOGRILLA(5, 12) = "TRUE"
    FORMATOGRILLA(5, 13) = "TRUE"
    FORMATOGRILLA(5, 14) = "TRUE"
    FORMATOGRILLA(5, 15) = "TRUE"
    FORMATOGRILLA(5, 16) = "TRUE"
    col = 17
    Grid1.Cols = col
    Grid1.Rows = row
    Grid1.AllowUserResizing = False
    Grid1.DisplayFocusRect = False
    Grid1.ExtendLastCol = True
    Grid1.BoldFixedCell = False
    Grid1.DrawMode = cellOwnerDraw
    Grid1.Appearance = Flat
    Grid1.ScrollBarStyle = Flat
    Grid1.FixedRowColStyle = Flat
    Grid1.Column(0).Width = 4 * 8.8
    Grid1.Column(1).Width = 2 * 10
    Grid1.Column(2).Width = 2 * 10
    Grid1.Column(3).Width = 4 * 10
    Grid1.Column(4).Width = 30 * 9
    Grid1.Column(5).Width = 3 * 9
    Grid1.Column(6).Width = 8 * 9
    Grid1.Column(7).Width = 8 * 9
    Grid1.Column(8).Width = 12 * 9
    Grid1.Column(9).Width = 3 * 9
    Grid1.Column(13).Width = 100
    Grid1.Column(14).Width = 40
    Grid1.Column(15).Width = 200
    Grid1.Column(16).Width = 200
    
    For k = 1 To col - 1
        Grid1.Cell(0, k).text = FORMATOGRILLA(1, k)
        
        Grid1.Column(k).MaxLength = Val(FORMATOGRILLA(2, k))
        Grid1.Column(k).FormatString = FORMATOGRILLA(4, k)
        Grid1.Column(k).Locked = FORMATOGRILLA(5, k)
        If FORMATOGRILLA(3, k) = "N" Then
            Grid1.Column(k).Alignment = cellRightCenter
            Grid1.Column(k).Mask = cellNumeric
        End If
        If FORMATOGRILLA(3, k) = "S" Then
            Grid1.Column(k).Alignment = cellLeftCenter
            'Grid1.Column(K).Mask = cellUpper
        End If
        If FORMATOGRILLA(3, k) = "D" Then
            Grid1.Column(k).CellType = cellCalendar
            Grid1.Column(k).Mask = cellNumeric
        End If
        
        'Grid1.Column(7).CellType = cellComboBox
    Next k
    Grid1.Range(0, 1, 0, 3).Merge
    Grid1.Cell(0, 1).text = "CUENTA"
    Grid1.Range(0, 0, 0, Grid1.Cols - 1).Alignment = cellCenterCenter

End Sub

Private Sub Form_Unload(Cancel As Integer)
Cancel = 0
If MODIFI = 1 Then Cancel = 1

End Sub

Private Sub glosa_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then grabar2.SetFocus

End Sub

Private Sub grabar2_Click()
    
        If Verifica_Permiso(Me.Caption, "agrega") = True Then
            If MODIFI = 1 Then
                Call ELIMINA
               
            End If
           
            Call GRABAR3
            MODIFI = 0
        
        Else
            MsgBox mensaje_nopermiso, vbCritical + vbOKOnly, "Permiso Denegado"
        End If
          
     End Sub

Sub GRABAR3()
    Dim conta
    'verifica si existe en comprobante
   If MODIFI = 0 Then
    ultimo
    End If
    rutprove = ""
    For k = 1 To Grid1.Rows - 1
        If dato0.text = "DB" And Grid1.Cell(k, 1).text + Grid1.Cell(k, 2).text + Grid1.Cell(k, 3).text = CUENTAPROVEEDOR Then
        rutprove = Grid1.Cell(k, 13).text
        End If
        
        If dato0.text = "NG" Then
             rutprove = Grid1.Cell(k, 13).text
        End If
        
        
        
        If Grid1.Cell(k, 9).text <> "" Then Call grabar(k)
    Next k
    ACTUALIZADOCUMENTO ("+")
    grabarglosa
    glosafinal.Visible = False
    Call opciones_FSCommand("imprime", "")
    'retorno
End Sub
Sub GRABAR4(grilla)
    Dim conta
    'verifica si existe en comprobante
   If MODIFI = 0 Then
    ultimo
    End If
    rutprove = ""
    For k = 1 To Grid1.Rows - 1
        If dato0.text = "DB" And Grid1.Cell(k, 1).text + Grid1.Cell(k, 2).text + Grid1.Cell(k, 3).text = CUENTAPROVEEDOR Then
        rutprove = Grid1.Cell(k, 13).text
        End If
        
        If Grid1.Cell(k, 9).text <> "" Then Call grabar(k)
    Next k
    ACTUALIZADOCUMENTO ("+")
    grabarglosa
    glosafinal.Visible = False
    Call opciones_FSCommand("imprime", "")
    'retorno
End Sub

Private Sub Grid1_GotFocus()
    glosafinal.Visible = False
    If dato3.text & dato4.text <> Format(fechasistema, "mm") & Format(fechasistema, "yyyy") Then
    dato2.text = ""
    dato3.text = "":
    dato4.text = "":
    dato2.SetFocus
End If
End Sub

Private Sub Grid1_KeyDown(KeyCode As Integer, ByVal Shift As Integer)
    Static numcheque As String
    Dim cuenta As String
    Dim fila As Long
    fila = Grid1.ActiveCell.row
    If KeyCode = 35 And MODIFI = 0 And Grid1.ActiveCell.col = 1 And Val(saldo.Caption) = 0 And Grid1.ActiveCell.row <> 1 Then Grid1.Cell(Grid1.ActiveCell.row, Grid1.ActiveCell.col).text = "": graba
    Rem If KeyCode = 38 And Grid1.ActiveCell.row = Grid1.Rows - 1 Then SG = "S" Else SG = "N"
    If Grid1.ActiveCell.col = 1 And KeyCode = vbKeyF2 Then Call ayudamayor(Grid1.ActiveCell.row, Grid1.ActiveCell.col)
    If Grid1.ActiveCell.col = 6 And Grid1.Cell(Grid1.ActiveCell.row, 1).text + Grid1.Cell(Grid1.ActiveCell.row, 2).text + Grid1.Cell(Grid1.ActiveCell.row, 3).text = CUENTAPROVEEDOR And KeyCode = vbKeyF2 Then Call ayudacompras(Grid1.ActiveCell.row, Grid1.ActiveCell.col)
    If MODIFI = 1 Then
        If Grid1.ActiveCell.col = 6 And Grid1.Cell(fila, 5).text = "CH" And numcheque = "" Then
            numcheque = Grid1.ActiveCell.text
            cuenta = Grid1.Cell(fila, 1).text & Grid1.Cell(fila, 2).text & Grid1.Cell(fila, 3).text
            Call eliminarCheque(cuenta, numcheque)
        End If
        If Grid1.ActiveCell.col = 6 And Grid1.Cell(fila, 5).text = "CH" And KeyCode = 13 Then
            numcheque = ""
            cuenta = ""
        End If
    End If
    End Sub

Private Sub Grid1_KeyPress(KeyAscii As Integer)
    Dim cadena As String
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    'If Grid1.ActiveCell.Col = 11 And Chr(KeyAscii) <> "D" And Chr(KeyAscii) <> "H" Then KeyAscii = 0
    If Grid1.ActiveCell.col = 1 And Chr(KeyAscii) = "*" And Val(saldo.Caption) = 0 And Grid1.ActiveCell.row <> 1 Then graba
    If Grid1.ActiveCell.col = 1 And Chr(KeyAscii) = "*" And Val(saldo.Caption) <> 0 And Grid1.ActiveCell.row <> 1 Then MsgBox ("comprobante descuadrado"): Grid1.Cell(Grid1.ActiveCell.row, 1).SetFocus
    
    
    Rem If formatogrilla(3, Grid1.ActiveCell.col) = "S" Then Grid1.Cell(Grid1.ActiveCell.row, Grid1.ActiveCell.col).text = UCase(Grid1.ActiveCell.text)
    If FORMATOGRILLA(3, Grid1.ActiveCell.col) = "N" Then snum = 1: KeyAscii = esNumero(KeyAscii)
    If FORMATOGRILLA(3, Grid1.ActiveCell.col) = "C" Then snum = 1: KeyAscii = esNumero(KeyAscii)
'    If KeyAscii = 13 Then
'        If fila > 1 Then
'            cadena = Grid1.Cell(fila - 1, columna).text
'            If Grid1.ActiveCell.text = "" Then
'                'Grid1.Cell(fila, columna).text = cadena
'                Grid1.ActiveCell.text = cadena
'                'If Grid1.ActiveCell.Col < Grid1.Cols - 1 Then
'                '    Call Grid1_LeaveCell(Grid1.ActiveCell.Row, Grid1.ActiveCell.Col, Grid1.ActiveCell.Row, Grid1.ActiveCell.Row + 1, False)
'                    'Grid1.Cell(Grid1.ActiveCell.Row, Grid1.ActiveCell.Col + 1).SetFocus
'                'Else
'                '    Grid1.Cell(Grid1.ActiveCell.Row, 1).SetFocus
'                'End If
'                'Grid1.Cell(Grid1.ActiveCell.Row, Grid1.ActiveCell.Col + 1).text = ""
'            End If
'        End If
'    End If
'    If columna = 9 Then
'        Grid1.Cell(fila, columna).text = Format(fechasistema, "dd-mm-yyyy")
'    End If
    If Grid1.ActiveCell.col = Grid1.Cols - 1 Then
        If KeyAscii <> 68 And KeyAscii <> 72 And KeyAscii <> 8 Then
           KeyAscii = 0
       End If
    End If
    If Grid1.ActiveCell.col = 5 Then
        If KeyAscii = 13 And Grid1.ActiveCell.text = "" Then
            Grid1.ActiveCell.text = dato0.text
        End If
    End If
End Sub

Sub graba()
    glosafinal.Visible = True
    glosa.Enabled = True
    glosa.SetFocus
End Sub

Sub retorno()
    MODIFI = 0
    Grid1.AutoRedraw = False
    CABEZA.Enabled = True
    Grid1.Enabled = True
    Grid1.Rows = 2
    dato1.Enabled = True
    dato1.SetFocus
    dato2.text = dia
    dato3.text = MES
    dato4.text = año
    
    For k = 1 To 14
    Grid1.Cell(1, k).text = ""
    Next k
    opciones.Visible = False
    glosafinal.Visible = False
    Grid1.AutoRedraw = True
    Grid1.Refresh
    debe.Caption = ""
    haber.Caption = ""
    saldo.Caption = ""
    Grid1.SelectionMode = cellSelectionFree
    Command5.Visible = True
    CmdImportar.Visible = True
    frmImportar.Visible = False
    frmimprimir.Visible = False
    centro.Caption = ""
End Sub

'Private Sub Grid1_LeaveCell(ByVal Row As Long, ByVal Col As Long, NewRow As Long, NewCol As Long, Cancel As Boolean)
'    Dim i As Long
'    If vacio = True Then
'        If Col < NewCol Then
'            Select Case Col
'                Case 4
'                    If TIENECTACTE(Row) = "0" Then
'                        NewCol = 5
'                    End If
'                    If TIENECRCC(Row) = "0" Then
'                        NewCol = 6
'                    End If
'                Case 5
'                    If TIENECRCC(Row) = "0" Then
'                        NewCol = 6
'                    End If
'                Case Else
'                    NewRow = fila
'                    NewCol = columna
'            End Select
'        End If
'        If Col = 11 And NewCol = 1 Then
'            NewRow = fila
'            NewCol = columna
'        End If
'    Else
'        Select Case Col
'            Case 1
'                PIVOTE.MaxLength = 2: PIVOTE.text = Grid1.Cell(Row, Col).text: Call ceros(PIVOTE): Grid1.Cell(Row, Col).text = PIVOTE.text
'                If Grid1.Cell(Row, Col).text = "00" And NewCol > Col Then
'                    Col = 1: NewCol = 1
'                End If
'            Case 2
'                PIVOTE.MaxLength = 2: PIVOTE.text = Grid1.Cell(Row, Col).text: Call ceros(PIVOTE): Grid1.Cell(Row, Col).text = PIVOTE.text
'                If Grid1.Cell(Row, Col).text = "00" And NewCol > Col Then
'                    Col = 2: NewCol = 2
'                End If
'            Case 3
'                PIVOTE.MaxLength = 4: PIVOTE.text = Grid1.Cell(Row, Col).text: Call ceros(PIVOTE): Grid1.Cell(Row, Col).text = PIVOTE.text
'                If Grid1.Cell(Row, Col).text = "0000" And NewCol > Col Then
'                    Col = 3: NewCol = 3
'                End If
'                Call leermayor(Row, 1)
'                If TIENECTACTE(Row) = "0" Then
'                    If Col < NewCol Then
'                        NewCol = 5
'                    End If
'                End If
'                If TIENECRCC(Row) = "0" Then
'                    If Col < NewCol Then
'                        NewCol = 6
'                    End If
'                End If
'
'
'            Case 4
'                If TIENECTACTE(Row) = "0" Then
'                    If Col < NewCol Then
'                        NewCol = 5
'                    End If
'                End If
'                If TIENECRCC(Row) = "0" Then
'                    If Col < NewCol Then
'                        NewCol = 6
'                    End If
'                End If
'            Case 5
'                If TIENECRCC(Row) = "0" Then
'                    If Col < NewCol Then
'                        NewCol = 6
'                    End If
'                End If
'
'            Case 11
'                If NewCol = 1 Then
'                    If Row = Grid1.Rows - 1 Then
'                        Grid1.AddItem "", True
'                        NewRow = Row + 1
'                    End If
'                End If
'        End Select
'        If Row <> NewRow Then
'            For i = 1 To 11
'                If Grid1.Cell(NewRow, i).text = "" Then
'                    Select Case i
'                        Case 4
'                            If TIENECTACTE(Row) <> "0" Then
'                                NewCol = i
'                                Exit For
'                            End If
'                        Case 5
'                            If TIENECRCC(Row) <> "0" Then
'                                NewCol = i
'                                Exit For
'                            End If
'                        Case Else
'                            NewCol = i
'                            Exit For
'                    End Select
'                End If
'            Next i
'        End If
'    End If
'End Sub
'

Private Sub Grid1_LeaveCell(ByVal row As Long, ByVal col As Long, NewRow As Long, NewCol As Long, Cancel As Boolean)
    Dim TEXTO As String
    If Grid1.SelectionMode <> cellSelectionByRow Then

    If row = 0 And col = 0 And Grid1.Rows - 1 = 1 Then NewRow = 1: NewCol = 1: GoTo no:
    If row = 0 And col = 0 And Grid1.Rows - 1 > 1 Then GoTo no:
    If NewCol = 10 And row = Grid1.Rows - 1 Then Grid1.Rows = Grid1.Rows + 1: NewRow = Grid1.Rows - 1
    If NewCol = 10 And row < Grid1.Rows - 1 Then NewCol = 1: NewRow = row + 1
    If NewCol > 9 Then NewCol = 9
    TEXTO = Grid1.Cell(row, col).text
    For k = 1 To 10
        If col = k And row > 1 And Grid1.Cell(row, col).text = "" And col < NewCol Then
            Grid1.Cell(row, col).text = Grid1.Cell(row - 1, col).text
        End If
    Next k

    lin = row
   'If Col = 1 And Row = Grid1.Rows - 1 And NewRow < Row Then GoTo paso2:

    If col = 1 Then
        pivote.MaxLength = 2: pivote.text = Grid1.Cell(row, col).text: Call ceros(pivote): Grid1.Cell(row, col).text = pivote.text
        If Grid1.Cell(row, col).text = "00" And NewCol > col Then
            Grid1.Cell(row, col).text = ""
            col = 1: NewCol = 1
        End If
    End If

    If col = 2 Then
        pivote.MaxLength = 2: pivote.text = Grid1.Cell(row, col).text: Call ceros(pivote): Grid1.Cell(row, col).text = pivote.text
        If Grid1.Cell(row, col).text = "00" And NewCol > col Then
            Grid1.Cell(row, col).text = ""
            col = 2: NewCol = 2
        End If
    End If
    
    If col = 3 Then
        pivote.MaxLength = 4: pivote.text = Grid1.Cell(row, col).text: Call ceros(pivote): Grid1.Cell(row, col).text = pivote.text
        If Grid1.Cell(row, col).text = "0000" And NewCol > col Then
            Grid1.Cell(row, col).text = ""
            col = 3: NewCol = 4
        End If
        If dato0.text = "CE" Or dato0.text = "DB" Or dato0.text = "PA" Then
            If Grid1.Cell(row, 1).text & Grid1.Cell(row, 2).text & Grid1.Cell(row, 3).text = "11120001" Then
                Grid1.Cell(Grid1.ActiveCell.row, 9).text = "H"
                Grid1.Cell(Grid1.ActiveCell.row, 9).Locked = True
            End If
        End If
    End If

'    If NewCol = 4 Or (NewRow <> Row And Col < 4) And Row < Grid1.Rows - 1 Then
'        Call leermayor(Row, Col)
'        If RESPUESTA = "N" Then
'            NewCol = Col
'            NewRow = Row
'        End If
'    End If
    
    If (col = 3 And NewCol = 4) Or (NewRow <> row And col < 4) And row < Grid1.Rows - 1 Then
        Call leermayor(row, col)
        If respuesta = "N" Then
            NewCol = 1
            NewRow = row
        End If
    End If
    

'    If (col = 3 And NewCol = 4) And row = 1 Then
'        Call leermayor(row, col)
'        If RESPUESTA = "N" Then
'            NewCol = col
'            NewRow = row
'        End If
'    End If
    
    If col = 7 And Grid1.Cell(row, col).text = "" Then Grid1.Cell(row, col).text = Format(fechasistema, "dd-mm-yyyy")
    If col = 9 And Grid1.Cell(row, col).text = "" And Grid1.Cell(row - 1, col).text = "D" And Grid1.Rows - 1 > 2 Then Grid1.Cell(row, col).text = "H"
    If col = 9 And Grid1.Cell(row, col).text = "" And Grid1.Cell(row - 1, col).text = "H" And Grid1.Rows - 1 > 2 Then Grid1.Cell(row, col).text = "D"
    If col = 6 Then
        pivote.MaxLength = 10
        pivote.text = Grid1.Cell(row, 6).text
        Call ceros(pivote)
        'If pivote.text = "0000000000" Then
        '    Grid1.Cell(Row, Col).text = ""
        'Else
            Grid1.Cell(row, col).text = pivote.text
        'End If
    End If
    If col = 9 Then SUMAR
10:

    If NewRow = row And NewCol > col Then
        If Grid1.Cell(row, 9).text <> "D" And Grid1.Cell(row, 9).text <> "H" Then NewCol = 9
        If Grid1.Cell(row, 8).text = "" Then NewCol = 8
        If Grid1.Cell(row, 7).text = "" Then NewCol = 7
        If Grid1.Cell(row, 6).text = "" Then NewCol = 6
        If Grid1.Cell(row, 5).text = "" Then NewCol = 5
        If Grid1.Cell(row, 4).text = "" Then NewCol = 4
        If Val(Grid1.Cell(row, 3).text) = 0 Then NewCol = 3
        If Val(Grid1.Cell(row, 2).text) = 0 Then NewCol = 2
        If Val(Grid1.Cell(row, 1).text) = 0 Then NewCol = 1
        NewRow = row
    End If
    Rem cuando cae

    If NewRow = Grid1.Rows - 1 And col < NewCol Then
        If Grid1.Cell(NewRow, 9).text <> "D" And Grid1.Cell(NewRow, 9).text <> "H" Then NewCol = 9
        If Grid1.Cell(NewRow, 8).text = "" Then NewCol = 8
        If Grid1.Cell(NewRow, 7).text = "" Then NewCol = 7
        If Grid1.Cell(NewRow, 6).text = "" Then NewCol = 6
        If Grid1.Cell(NewRow, 5).text = "" Then NewCol = 5
        If Grid1.Cell(NewRow, 4).text = "" Then NewCol = 4
        If Val(Grid1.Cell(NewRow, 3).text) = 0 Then NewCol = 3
        If Val(Grid1.Cell(NewRow, 2).text) = 0 Then NewCol = 2
        If Val(Grid1.Cell(NewRow, 1).text) = 0 Then NewCol = 1
    End If
   
    'If NewCol = 7 Then
   
   If TIENEBANCO(row) = "0" And Grid1.Cell(row, 5).text = "CH" Then
        MsgBox "NO PUEDE AGREGAR UN CHEQUE SI LA CUENTA NO ES BANCO", vbCritical, "ATENCION"
        Grid1.Cell(row, 5).text = ""
        NewRow = row
        NewCol = 1
        Exit Sub
   End If
   
    If TIENEBANCO(row) = "1" And Grid1.Cell(row, 5).text = "CH" Then
        If MODIFI = 0 Then
            Call leercheque(row)
            If chequecreado = "S" Then
                MsgBox ("cheque ya esta girado")
                NewRow = row
                NewCol = 6
                col = 6
            End If
        End If
    End If

    If col = 6 And NewCol > 6 Then
        If (Grid1.Cell(row, 5).text = "FC" Or Grid1.Cell(row, 5).text = "EE") And Grid1.Cell(row, 1).text + Grid1.Cell(row, 2).text + Grid1.Cell(row, 3).text = CUENTAPROVEEDOR Then
            Call leefacturadecompra(row, col)
            If facturano = "N" Then
                MsgBox ("factura no corresponde a proveedor")
                Grid1.ActiveCell.text = ""
                
                NewRow = row
                NewCol = 6
                col = 6
                
            End If
        
            If facturano = "R" Then
                Grid1.ActiveCell.text = ""
                Grid1.Cell(Grid1.ActiveCell.row, 8).text = ""
                NewRow = row
                NewCol = 6
                col = 6
                
            End If
        
        
        If PAGARFACTURA = False And facturano = "S" Then
                MsgBox ("factura ya se encuentra cancelada")
                Grid1.ActiveCell.text = ""
                NewRow = row
                NewCol = 6
                col = 6
                
                
                
            End If
                
'                If saldocuentacorriente("23100027", DATO20.text + DV.Caption, dato4.text, empresaactiva) <> 0 Then
'                      If MsgBox("PROVEEDOR TIENE UN ANTICIPO PENDIENTE POR REBAJAR RECTIFIQUE ") = vbYes Then
'                      End If
'                End If
        Rem  Grid1.Cell(Grid1.ActiveCell.row, 8).Locked = True
        Else
        Grid1.Cell(Grid1.ActiveCell.row, 8).Locked = False
        Grid1.Cell(Grid1.ActiveCell.row, 8).Locked = False
        End If
        
        
        If Grid1.Cell(row, 5).text = "BH" Then
            Call leeboletadehonorarios(row, col)
            If boletano = "N" Then
                MsgBox ("boleta no corresponde a prestador")
                Grid1.ActiveCell.text = ""
                NewRow = row
                NewCol = 6
                col = 6
            End If
        If PAGARFACTURA = False Then
                MsgBox ("boleta honorarios ya se encuentra cancelada")
                Grid1.ActiveCell.text = ""
                NewRow = row
                NewCol = 6
                col = 6
            End If
        
        
        End If
    End If
    'End If

    If NewRow <> row And Grid1.Rows - 1 <> row Then
        If Grid1.Cell(row, col).text = "" Then
            NewRow = row
            col = NewCol
        End If
    End If
If NewRow <> row And Grid1.Rows - 1 <> row Then
    If Grid1.Cell(row, 9).text <> "D" And Grid1.Cell(row, 9).text <> "H" Then NewCol = 9: NewRow = row
    
End If
    If NewRow = Grid1.Rows - 1 And Grid1.Rows > 2 And row < NewRow Then NewCol = 1
    
  

no:
End If
End Sub

Sub CREARCTACTE(row)
    maestro02.dato1.text = Grid1.Cell(row, 1).text + Grid1.Cell(row, 2).text + Grid1.Cell(row, 3).text
    maestro02.dato2.text = Mid(Grid1.Cell(row, 5).text, 1, 9)
    maestro02.Show
End Sub

Sub flechas(ByRef caja As TextBox, ByRef caja1 As TextBox, ByRef codigo As Integer)
    If codigo = 38 And caja.Enabled = True Then caja.SetFocus
    If codigo = 40 And caja1.Enabled = True Then caja1.SetFocus
End Sub

Sub Pregunta(ByRef caja As TextBox, ByRef caja1 As TextBox)
    If caja.text = "" And sc = 0 Then caja.SetFocus
    If caja.text <> "" Or sc = 1 Then caja1.Enabled = True: caja1.SetFocus
    sc = 0
End Sub


Sub leeultima()
End Sub

Sub grabar(row)
    Dim w As Long
    Dim tipo2 As String
    
    LINEAS.text = row
    w = row
    Call ceros(LINEAS)
    campos(0, 0) = "tipo"
    campos(1, 0) = "numero"
    campos(2, 0) = "linea"
    campos(3, 0) = "fecha"
    campos(4, 0) = "codigocuenta"
    campos(5, 0) = "glosacontable"
    campos(6, 0) = "tipodocumento"
    campos(7, 0) = "numerodocumento"
    campos(8, 0) = "fechadocumento"
    campos(9, 0) = "fechavencimiento"
    campos(10, 0) = "monto"
    campos(11, 0) = "dh"
    campos(12, 0) = "creadopor"
    campos(13, 0) = "mes"
    campos(14, 0) = "año"
    campos(15, 0) = "rutctacte"
    campos(16, 0) = "centrocosto"
    campos(17, 0) = "fechacreacion"
    campos(18, 0) = "horacreacion"
    campos(19, 0) = "rutproveedor"
    campos(20, 0) = "cuenta_presupuesto"
    campos(21, 0) = "centro_gastos"
    campos(22, 0) = ""
    
    campos(0, 1) = dato0.text
    campos(1, 1) = dato1.text
    campos(2, 1) = LINEAS.text
    campos(3, 1) = dato4.text + "-" + dato3.text + "-" + dato2.text
    campos(4, 1) = Grid1.Cell(w, 1).text + Grid1.Cell(w, 2).text + Grid1.Cell(w, 3).text
    campos(5, 1) = Grid1.Cell(w, 4).text
    campos(6, 1) = Grid1.Cell(w, 5).text
    campos(7, 1) = Grid1.Cell(w, 6).text
    campos(8, 1) = campos(3, 1)
    campos(9, 1) = Mid(Grid1.Cell(w, 7).text, 7, 4) + "-" + Mid(Grid1.Cell(w, 7).text, 4, 2) + "-" + Mid(Grid1.Cell(w, 7).text, 1, 2)
    campos(10, 1) = Replace(Grid1.Cell(w, 8).text, ",", ".")
    If Grid1.Cell(w, 5).text = "FC" Then Call abonofactura(w, 8, False)
    If Grid1.Cell(w, 5).text = "EE" Then Call abonofactura(w, 8, False)
    If Grid1.Cell(w, 5).text = "BH" Then Call abonoboletadehonorarios(w, 8, False)
    
    campos(11, 1) = Grid1.Cell(w, 9).text
    campos(12, 1) = USUARIOSISTEMA
    campos(13, 1) = dato3.text
    campos(14, 1) = dato4.text
    If TIENECTACTE(w) = "0" Then Grid1.Cell(w, 13).text = ""
    
    campos(15, 1) = Grid1.Cell(w, 13).text
    campos(16, 1) = Grid1.Cell(w, 14).text
    campos(17, 1) = Format(Date$, "yyyy") + "-" + Format(Date$, "mm") + "-" + Format(Date$, "dd")
    campos(18, 1) = Time$
    campos(19, 1) = rutprove
    campos(20, 1) = Mid(Grid1.Cell(w, 15).text, 1, 4)
    campos(21, 1) = Mid(Grid1.Cell(w, 16).text, 1, 4)
    
    campos(0, 2) = "movimientoscontables"
    condicion = ""
    op = 2
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)
    If TIENEBANCO(k) = "1" And Grid1.Cell(w, 5).text = "CH" And Grid1.Cell(w, 9).text = "H" Then Call grabacheque(w)
End Sub

Sub grabarglosa()
    campos(0, 0) = "tipo"
    campos(1, 0) = "numero"
    campos(2, 0) = "fecha"
    campos(3, 0) = "glosa"
    campos(4, 0) = ""
    campos(0, 1) = dato0.text
    campos(1, 1) = dato1.text
    campos(2, 1) = dato4.text + "-" + dato3.text + "-" + dato2.text
    campos(3, 1) = glosa.text
    
    campos(0, 2) = "movimientos_glosa"
    condicion = ""
    op = 2
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)
    
    
End Sub
Sub leerglosa()
    campos(0, 0) = "tipo"
    campos(1, 0) = "numero"
    campos(2, 0) = "fecha"
    campos(3, 0) = "glosa"
    campos(4, 0) = ""
    
    campos(0, 2) = "movimientos_glosa"
    condicion = "tipo='" + dato0.text + "' and numero='" + dato1.text + "' and fecha='" + dato4.text + "-" + dato3.text + "-" + dato2.text + "'"
    op = 5
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)
    If IsNull(sqlconta.response(3, 3)) = False Then glosa.text = sqlconta.response(3, 3)
    
End Sub


Sub leecomprobante()
    Dim lin As Integer
    
    Dim resultados As rdoResultset
    Dim csql As New rdoQuery
    Dim rut2 As String

    
        Set csql.ActiveConnection = contadb
        csql.sql = "SELECT tipo,numero,linea,fecha,codigocuenta,glosacontable,tipodocumento,numerodocumento,fechavencimiento,monto,dh,rutctacte,centrocosto,cuenta_presupuesto,centro_gastos,rutproveedor "
        csql.sql = csql.sql + "FROM movimientoscontables "
            
        csql.sql = csql.sql + "WHERE tipo='" + dato0.text + "' and numero='" & dato1.text & "'and año='" + Format(fechasistema, "yyyy") + "' and mes='" + Format(fechasistema, "mm") + "' order by linea"
        csql.Execute

        canli = 0
        If csql.RowsAffected > 0 Then
        CmdImportar.Visible = False
        Command5.Visible = False
            Set resultados = csql.OpenResultset
            While Not resultados.EOF
             fechacon = resultados(3)
             canli = canli + 1
                rut2 = resultados(2)
                
                dato2.text = Mid(resultados(3), 1, 2)
                dato3.text = Mid(resultados(3), 4, 2)
                dato4.text = Mid(resultados(3), 7, 4)
                
                
                grilladetalle(canli, 1) = Mid(resultados(4), 1, 2)
                grilladetalle(canli, 2) = Mid(resultados(4), 3, 2)
                grilladetalle(canli, 3) = Mid(resultados(4), 5, 4)
                grilladetalle(canli, 4) = resultados(5)
                grilladetalle(canli, 5) = resultados(6)
                grilladetalle(canli, 6) = resultados(7)
                grilladetalle(canli, 7) = resultados(8)
                grilladetalle(canli, 8) = resultados(9)
                grilladetalle(canli, 9) = resultados(10)
                If resultados(4) = "11130001" And resultados(0) = "NG" Then
                    grilladetalle(canli, 10) = resultados(15)
                Else
                    grilladetalle(canli, 10) = resultados(11)
                End If
                grilladetalle(canli, 11) = resultados(12)
              
                
                grilladetalle(canli, 12) = resultados(13)
                grilladetalle(canli, 13) = resultados(14)
                
                If resultados(6) = "FA" And resultados(7) <> "0000000000" And resultados(10) = "D" Then
                Call leecentrofactura(resultados(6), resultados(7), resultados(11))
                End If
                

                resultados.MoveNext

            Wend
            cargadorcomprobante
            resultados.Close
            Set resultados = Nothing
                            Else
                CmdImportar.Visible = True
                Command5.Visible = True
        End If
    
leerglosa
   If csql.RowsAffected > 0 Then
   opciones.Visible = True
   Grid1.Enabled = True
   CABEZA.Enabled = False
    Grid1.SelectionMode = cellSelectionByRow
    
    
   opciones.SetFocus
End If

no:
End Sub
Sub cargadorcomprobante()
    Dim LINEA As Long
    Grid1.AutoRedraw = False
    Grid1.Rows = canli + 1
    For k = 1 To canli
    Grid1.Cell(k, 1).text = grilladetalle(k, 1)
    Grid1.Cell(k, 2).text = grilladetalle(k, 2)
    Grid1.Cell(k, 3).text = grilladetalle(k, 3)
    Grid1.Cell(k, 4).text = grilladetalle(k, 4)
    Grid1.Cell(k, 5).text = grilladetalle(k, 5)
    Grid1.Cell(k, 6).text = grilladetalle(k, 6)
    Grid1.Cell(k, 7).text = grilladetalle(k, 7)
    Grid1.Cell(k, 8).text = grilladetalle(k, 8)
    Grid1.Cell(k, 9).text = grilladetalle(k, 9)
    Grid1.Cell(k, 10).text = leerNombreMayor(grilladetalle(k, 1) + grilladetalle(k, 2) + grilladetalle(k, 3))
    Grid1.Cell(k, 11).text = leerNombrerut(grilladetalle(k, 1) + grilladetalle(k, 2) + grilladetalle(k, 3), grilladetalle(k, 10))
    Grid1.Cell(k, 12).text = leerNOMBREcrcc(grilladetalle(k, 11))
    Grid1.Cell(k, 13).text = grilladetalle(k, 10)
    Grid1.Cell(k, 14).text = grilladetalle(k, 11)
    If grilladetalle(k, 12) <> "" Then
    Grid1.Cell(k, 16).text = grilladetalle(k, 12) + ":" + leernombreanalisis(grilladetalle(k, 1) + grilladetalle(k, 2) + grilladetalle(k, 3), grilladetalle(k, 12))
    
    Else
    Grid1.Cell(k, 16).text = ""
    
    End If
    
    If grilladetalle(k, 13) <> "" Then
    Grid1.Cell(k, 15).text = grilladetalle(k, 13) + ":" + leerNOMBREgastos(Mid(grilladetalle(k, 13), 1, 4))
    
    Else
    Grid1.Cell(k, 15).text = ""
    
    End If
    
    
    LINEA = k
    
    Next k
    SUMAR
    Grid1.AutoRedraw = True
    Grid1.Refresh
End Sub
                

Private Sub imprime1_Click()
    
    Call imprimir(1)
End Sub

Private Sub imprime2_Click()
Grid2.ReportTitles.Clear
    
    Call IMPRIMIR2
End Sub

Private Sub imprime3_Click()
    Call imprimir(3)
End Sub

Private Sub MANUAL_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then Call opciones_FSCommand("retorno", "")
    If UCase(Chr(KeyAscii)) = "M" Then Call opciones_FSCommand("modifica", "")
    If UCase(Chr(KeyAscii)) = "E" Then Call opciones_FSCommand("elimina", "")
    If UCase(Chr(KeyAscii)) = "S" Then Call opciones_FSCommand("siguiente", "")
    If UCase(Chr(KeyAscii)) = "A" Then Call opciones_FSCommand("anterior", "")
    If UCase(Chr(KeyAscii)) = "R" Then Call opciones_FSCommand("retorno", "")
    If UCase(Chr(KeyAscii)) = "I" Then Call opciones_FSCommand("imprime", "")
End Sub

Private Sub opciones_FSCommand(ByVal command As String, ByVal args As String)
    If command = "retorno" Then retorno
    If command = "modifica" Then
    If comprobantedigitable(dato0.text, Me.Caption) = True Then
     If Verifica_Permiso(Me.Caption, "modifica") = True Then
       
            Grid1.SelectionMode = cellSelectionFree
            'Grid1.SelectionMode = grid1.cel
             modifica
        Else
            MsgBox mensaje_nopermiso, vbCritical + vbOKOnly, "Permiso Denegado"
            
        End If
     Else
        MsgBox "COMPROBANTE BLOQUEADO, IMPOSIBLE MODIFICAR", vbCritical + vbOKOnly, "Permiso Denegado"
    End If
    
    End If
    
    If command = "elimina" Then
    If comprobantedigitable(dato0.text, Me.Caption) = True Then
        If Verifica_Permiso(Me.Caption, "elimina") = True Then
            If MsgBox("ESTA SEGURO DE ELIMINAR ", vbYesNo) = vbYes Then
                ELIMINA
                retorno
            End If
        Else
         MsgBox mensaje_nopermiso, vbCritical + vbOKOnly, "Permiso Denegado"
        End If
    Else
        MsgBox "COMPROBANTE BLOQUEADO, IMPOSIBLE MODIFICAR", vbCritical + vbOKOnly, "Permiso Denegado"
    End If
    
    End If
    
    If command = "siguiente" Then SIGUIENTE
    If command = "anterior" Then anterior
    If command = "imprime" Then imprimirtipo
End Sub

Sub modifica()
    MODIFICANDO.Visible = True
    opciones.Visible = False
    MODIFI = 1
    CABEZA.Enabled = True
    Grid1.Enabled = True
    Grid1.Rows = Grid1.Rows + 1
    Grid1.Cell(Grid1.Rows - 1, 1).SetFocus
End Sub

Sub imprimirtipo()
    If dato0.text <> "CE" And dato0.text <> "PA" And dato0.text <> "DB" And dato0.text <> "AG" And dato0.text <> "PS" Then
        imprime1.Visible = False
        imprime2.Visible = False
        Command2.Visible = False
    Else
        imprime1.Visible = True
        imprime2.Visible = True
        Command2.Visible = True
    End If
    frmimprimir.Visible = True
End Sub

Sub imprimir(documento As Integer)
    Dim j As Integer
    Dim n As Integer
    Dim a As Integer
    Dim TITU As String
    Dim montocheque As Double
    Dim numerocheque As String
    Dim fechacheque As Date
    Dim CHEQUEGIRADO As String
    Dim IMPRIMECHEQUE As Boolean
    Dim largo As Integer
    Dim monto As String
    Dim calzador As Integer
    Dim FINAL As Double
    
    Rem DATOS DE LA COLUMNA
    FORMATOGRILLA(1, 1) = "CUENTA"
    FORMATOGRILLA(1, 2) = "L"
    FORMATOGRILLA(1, 3) = "CODIGO"
    FORMATOGRILLA(1, 4) = "TP"
    FORMATOGRILLA(1, 5) = "NUMERO"
    FORMATOGRILLA(1, 6) = "F.VENCI."
    FORMATOGRILLA(1, 7) = "GLOSA"
    FORMATOGRILLA(1, 8) = "DEBE"
    FORMATOGRILLA(1, 9) = "HABER"
    FORMATOGRILLA(1, 10) = "ANALISIS"
    
    
    Rem LARGO DE LOS DATOS
    
    FORMATOGRILLA(2, 1) = "19"
    FORMATOGRILLA(2, 2) = "2"
    FORMATOGRILLA(2, 3) = "8"
    FORMATOGRILLA(2, 4) = "3"
    FORMATOGRILLA(2, 5) = "9"
    FORMATOGRILLA(2, 6) = "8"
    FORMATOGRILLA(2, 7) = "30"
    FORMATOGRILLA(2, 8) = "9"
    FORMATOGRILLA(2, 9) = "9"
    FORMATOGRILLA(2, 10) = "9"
    FORMATOGRILLA(2, 11) = "9"
    
    Rem TIPO DE DATOS C=CEROS,S=STRING,N=NUMERICO,D=DATE,
    FORMATOGRILLA(3, 1) = "S"
    FORMATOGRILLA(3, 2) = "N"
    FORMATOGRILLA(3, 3) = "S"
    FORMATOGRILLA(3, 4) = "S"
    FORMATOGRILLA(3, 5) = "C"
    FORMATOGRILLA(3, 6) = "D"
    FORMATOGRILLA(3, 7) = "S"
    FORMATOGRILLA(3, 8) = "N"
    FORMATOGRILLA(3, 9) = "N"
    FORMATOGRILLA(3, 10) = "S"
    
    Rem FORMATO GRILLA
    FORMATOGRILLA(4, 1) = ""
    FORMATOGRILLA(4, 2) = ""
    FORMATOGRILLA(4, 3) = ""
    FORMATOGRILLA(4, 4) = ""
    FORMATOGRILLA(4, 5) = ""
    FORMATOGRILLA(4, 6) = "dd-mm-yyyy"
    FORMATOGRILLA(4, 7) = ""
    FORMATOGRILLA(4, 8) = " ###,###,##0"
    FORMATOGRILLA(4, 9) = " ###,###,##0"
    FORMATOGRILLA(4, 10) = ""
    FORMATOGRILLA(4, 11) = ""
    
    Rem LOCCKED
    FORMATOGRILLA(5, 1) = "FALSE"
    FORMATOGRILLA(5, 2) = "FALSE"
    FORMATOGRILLA(5, 3) = "FALSE"
    FORMATOGRILLA(5, 4) = "FALSE"
    FORMATOGRILLA(5, 5) = "FALSE"
    FORMATOGRILLA(5, 6) = "FALSE"
    FORMATOGRILLA(5, 7) = "FALSE"
    FORMATOGRILLA(5, 8) = "FALSE"
    FORMATOGRILLA(5, 9) = "FALSE"
    FORMATOGRILLA(5, 10) = "FALSE"
    
    Rem VALOR MINIMO
    FORMATOGRILLA(6, 1) = ""
    FORMATOGRILLA(6, 2) = ""
    FORMATOGRILLA(6, 3) = ""
    FORMATOGRILLA(6, 4) = ""
    FORMATOGRILLA(6, 5) = ""
    FORMATOGRILLA(6, 6) = ""
    FORMATOGRILLA(6, 7) = ""
    FORMATOGRILLA(6, 8) = ""
    FORMATOGRILLA(6, 9) = ""
    
    Rem VALOR MAXIMO
    FORMATOGRILLA(7, 1) = ""
    FORMATOGRILLA(7, 2) = ""
    FORMATOGRILLA(7, 3) = ""
    FORMATOGRILLA(7, 4) = ""
    FORMATOGRILLA(7, 5) = ""
    FORMATOGRILLA(7, 6) = ""
    FORMATOGRILLA(7, 7) = ""
    FORMATOGRILLA(7, 8) = ""
    FORMATOGRILLA(7, 9) = ""
    Grid2.Cols = 11
    Grid2.Rows = 1
    FINAL = Grid1.Rows
    
    If FINAL > 14 And documento = 1 Then FINAL = 15
    
    
    Grid2.Rows = FINAL
    Grid2.Range(1, 1, Grid2.Rows - 1, Grid2.Cols - 1).ClearText
    Grid2.Column(10).Alignment = cellRightGeneral
    
    
    For k = 1 To Grid2.Cols - 1
        Grid2.Cell(0, k).text = FORMATOGRILLA(1, k)
        Grid2.Column(k).MaxLength = Val(FORMATOGRILLA(2, k))
        Grid2.Column(k).FormatString = FORMATOGRILLA(4, k)
        Rem Grid2.Column(K).Locked = FORMATOGRILLA(5, K)
        If FORMATOGRILLA(3, k) = "N" Then Grid2.Column(k).Alignment = cellRightCenter
        If FORMATOGRILLA(3, k) = "D" Then Grid2.Column(k).CellType = cellCalendar
        'Grid1.Column(7).CellType = cellComboBox
    Next k
    
    Grid2.DefaultFont.Size = 6.5
    For k = 1 To Grid2.Cols - 1
        Grid2.Column(k).Width = Val(FORMATOGRILLA(2, k)) * (Grid2.DefaultFont.Size)
    Next k

    Grid2.PageSetup.Orientation = cellPortrait
    Grid2.PageSetup.PrintFixedRow = True
    Grid2.PageSetup.PrintGridlines = False

    Grid2.PageSetup.BlackAndWhite = True
    Grid2.PageSetup.BottomMargin = 1
    Grid2.PageSetup.TopMargin = 1
    Grid2.PageSetup.LeftMargin = 1.9
    Grid2.PageSetup.RightMargin = 0.5
    Grid2.Range(0, 1, 0, Grid2.Cols - 1).FontBold = True
    Grid2.Range(0, 1, 0, Grid2.Cols - 1).Borders(cellEdgeTop) = cellThin
    Grid2.Range(0, 1, 0, Grid2.Cols - 1).Borders(cellEdgeBottom) = cellThin
    'Grid2.Range(0, 1, 0, 3).Merge
    'Grid2.Cell(0, 1).text = "CUENTA"
    TITU = dato0.text + " " + leerNOMBREcomprobantes(dato0.text)
    
    
    Call cabezERA(TITU + " NUMERO " + dato1.text, dato1.text, fechacon)
    IMPRIMECHEQUE = False
    
    For j = 1 To Grid1.Rows - 1
                
       If FINAL = 15 And documento = 1 Then GoTo PASO:
        Grid2.Cell(j, 1).text = Grid1.Cell(j, 10).text
        Grid2.Cell(j, 2).text = j
        Grid2.Cell(j, 3).text = Grid1.Cell(j, 1).text + "." + Grid1.Cell(j, 2).text + "." + Grid1.Cell(j, 3).text
        Grid2.Cell(j, 4).text = Grid1.Cell(j, 5).text
        Grid2.Cell(j, 5).text = Grid1.Cell(j, 6).text
        Grid2.Cell(j, 6).text = Grid1.Cell(j, 7).text
        Grid2.Cell(j, 7).text = Grid1.Cell(j, 4).text
        If Grid1.Cell(j, 9).text = "D" Then Grid2.Cell(j, 8).text = Grid1.Cell(j, 8).text
        If Grid1.Cell(j, 9).text = "H" Then Grid2.Cell(j, 9).text = Grid1.Cell(j, 8).text
        If Grid1.Cell(j, 13).text <> "" Then Grid2.Cell(j, 10).text = Mid(Grid1.Cell(j, 13).text, 1, 9) + "-" + Mid(Grid1.Cell(j, 13).text, 10, 1)
        If Grid1.Cell(j, 14).text <> "" Then Grid2.Cell(j, 10).text = Grid1.Cell(j, 14).text
    
       
PASO:
        If Grid1.Cell(j, 5).text = "CH" Then
            montocheque = Grid1.Cell(j, 8).text
            fechacheque = Grid1.Cell(j, 7).text
            numerocheque = Grid1.Cell(j, 6).text
            CHEQUEGIRADO = Grid1.Cell(j, 4).text
            IMPRIMECHEQUE = True
        End If
    Next j
    
    
    
    If FINAL = 15 And documento = 1 Then
    
    Grid2.Range(1, 1, 1, 7).Merge
    Grid2.Range(1, 1, 1, 7).Alignment = cellCenterCenter
    Grid2.Range(1, 1, 1, 7).FontSize = 12
    Grid2.Range(1, 1, 1, 7).FontBold = True
    
    
    Grid2.Cell(1, 1).text = "**** IMPRIMIR COMPROBANTE MANUAL ******"
    
    
    End If
    
    
    For a = FINAL To 14
        Grid2.Rows = Grid2.Rows + 1
    Next a
    If documento <> 1 Then Grid2.Rows = Grid2.Rows + 1
    Grid2.Range(Grid2.Rows - 1, 1, Grid2.Rows - 1, Grid2.Cols - 1).Borders(cellEdgeTop) = cellThin
    Grid2.Cell(Grid2.Rows - 1, 7).text = "TOTAL COMPROBANTE CONTABLE"
    Grid2.Cell(Grid2.Rows - 1, 8).text = debe.Caption
    Grid2.Cell(Grid2.Rows - 1, 9).text = haber.Caption

    Grid2.AddItem ""
    Grid2.Range(Grid2.Rows - 1, 1, Grid2.Rows - 1, Grid2.Cols - 7).Merge
    Grid2.Cell(Grid2.Rows - 1, 1).text = "GLOSA :" + glosa.text

    Grid2.AddItem ""
    Grid2.AddItem ""
    Grid2.AddItem ""

    Grid2.Range(Grid2.Rows - 1, 1, Grid2.Rows - 1, Grid2.Cols - 1).Merge
    Grid2.Cell(Grid2.Rows - 1, 1).text = "              _______________________                       ______________________                           _______________________             ____________________"
    Grid2.AddItem ""

    Grid2.Range(Grid2.Rows - 1, 1, Grid2.Rows - 1, Grid2.Cols - 1).Merge
    Grid2.Cell(Grid2.Rows - 1, 1).text = "                  VºBº Gerencia                                            VºBº Contabilidad                                      Recibo Conforme                           Confec. " + USUARIOSISTEMA

    Rem CHEQUE SANTANDER
    
    
    
    If documento = 1 Then
        Grid2.PageSetup.LeftMargin = 0.5
        Grid2.Column(7).Width = Grid2.Column(7).Width - 20
        For k = 1 To 3
            Grid2.AddItem ""
        Next k
         Grid2.AddItem ""
        Grid2.Range(Grid2.Rows - 2, 7, Grid2.Rows - 1, 7).Merge
        Grid2.Range(Grid2.Rows - 2, 7, Grid2.Rows - 1, 7).CellType = cellTextBox
        Grid2.Range(Grid2.Rows - 2, 7, Grid2.Rows - 1, 7).FontBold = True
        Grid2.Range(Grid2.Rows - 2, 7, Grid2.Rows - 1, 7).FontSize = 12
        Grid2.Range(Grid2.Rows - 2, 7, Grid2.Rows - 1, 7).Alignment = cellLeftCenter
        
        Grid2.Cell(Grid2.Rows - 2, 7).text = "  " + Format(montocheque, "##,###,###,###") & ".="
    
        Grid2.AddItem ""
           Grid2.RowHeight(Grid2.Rows - 1) = 22
        Grid2.AddItem ""
        
        Grid2.Range(Grid2.Rows - 1, 3, Grid2.Rows - 1, 10).FontBold = True
        Grid2.Range(Grid2.Rows - 1, 3, Grid2.Rows - 1, 10).FontSize = 9
        Grid2.Range(Grid2.Rows - 1, 5, Grid2.Rows - 1, 6).Merge
        
        Grid2.Cell(Grid2.Rows - 1, 3).text = Format(fechacheque, "dd")
        Grid2.Cell(Grid2.Rows - 1, 3).Alignment = cellRightBottom
        
        Grid2.Cell(Grid2.Rows - 1, 5).text = MonthName(Format(fechacheque, "mm"))
        Grid2.Cell(Grid2.Rows - 1, 5).Alignment = cellLeftBottom
        Grid2.Cell(Grid2.Rows - 1, 7).text = Format(fechacheque, "yyyy")
        Grid2.Cell(Grid2.Rows - 1, 7).Alignment = cellRightBottom
        Grid2.RowHeight(Grid2.Rows - 1) = 22
        
        
        
        Grid2.AddItem ""
        
        Grid2.Range(Grid2.Rows - 1, 1, Grid2.Rows - 1, 8).FontBold = True
        Grid2.Range(Grid2.Rows - 1, 1, Grid2.Rows - 1, 8).FontSize = 9
        Grid2.Range(Grid2.Rows - 1, 1, Grid2.Rows - 1, 8).Merge
        Grid2.Cell(Grid2.Rows - 1, 1).text = "       " + CHEQUEGIRADO

    
        Grid2.AddItem ""
        Grid2.AddItem ""
        
        'numToLet(MONTOCHEQUE, "PESO", "", "CENTAVO", "CENTAVOS", 0)
    
        
        Grid2.Range(Grid2.Rows - 1, 1, Grid2.Rows - 1, 7).FontBold = True
        Grid2.Range(Grid2.Rows - 1, 1, Grid2.Rows - 1, 7).FontSize = 9
        Grid2.Range(Grid2.Rows - 1, 1, Grid2.Rows - 1, 7).Merge
        monto = WORDNUM(montocheque, "PESO", "", "CENTAVO", "CENTAVOS", 0)
        monto = UCase(monto)
        'ariel revisa largo cheques
        largo = Len(monto)
        calzador = largo
        If largo > 60 Then
            For k = 60 To 1 Step -1
                If Asc(Mid(monto, k, 1)) <> 32 Then calzador = 60 - (60 - k): Exit For
            Next k
            Grid2.Cell(Grid2.Rows - 1, 1).text = "      " + Mid(monto, 1, calzador) + "-"
            Grid2.AddItem ""
            Grid2.Range(Grid2.Rows - 1, 1, Grid2.Rows - 1, 7).FontBold = True
            Grid2.Range(Grid2.Rows - 1, 1, Grid2.Rows - 1, 7).FontSize = 9
            Grid2.Range(Grid2.Rows - 1, 1, Grid2.Rows - 1, 7).Merge
            Grid2.Cell(Grid2.Rows - 1, 1).text = Mid(monto, calzador + 1, largo)
        Else
            Grid2.Cell(Grid2.Rows - 1, 1).text = "      " + monto
            Grid2.Range(Grid2.Rows - 1, 1, Grid2.Rows - 1, 7).Merge
        End If
        Grid2.Cell(Grid2.Rows - 1, 1).text = Grid2.Cell(Grid2.Rows - 1, 1).text & "**************************************************"
        
        Grid2.Rows = 49
Grid2.Cell(32, 1).Font.Bold = True
Grid2.Cell(32, 1).Font.Size = 10
Grid2.Cell(32, 1).text = "CH:" + numerocheque

        Grid2.Range(38, 1, 48, 8).FontBold = True
        Grid2.Range(38, 1, 48, 8).FontSize = 8
   
   Grid2.Range(38, 1, 38, 7).Merge
   Grid2.RowHeight(38) = 22
   Grid2.Cell(38, 1).text = "                       " & Format(fechacon, "dd") + "            " + MonthName(Format(fechacon, "mm")) + "    " + Format(fechacon, "yyyy")
   Grid2.Cell(40, 7).Alignment = cellRightCenter
   Grid2.Cell(40, 7).text = numerocheque
   Grid2.Cell(41, 1).text = "       Santander"
   Grid2.Range(41, 5, 41, 6).Merge
   
   Grid2.Cell(41, 5).text = Format(montocheque, "###,###,###")
   
   Grid2.Range(46, 1, 42, 9).Merge
   Grid2.Range(47, 1, 43, 9).Merge
   Grid2.Range(48, 1, 44, 9).Merge
   Grid2.Cell(46, 1).text = "SR : PROVEEDOR. CUALQUIER DUDA O ALCANCE SOBRE ESTE PAGO. HACERLO"
   Grid2.Cell(47, 1).text = "EN UN PLAZO MAXIMO DE 30 DIAS POSTERIORES A LA FECHA DE PAGO"
   Grid2.Cell(48, 1).text = "PASANDO DICHO PLAZO LA EMPRESA NO ACEPTARA RECLAMOS"


    
    
    
    
    
    
    
    
    
    
    
    
    End If



'    Rem BANCO CHILE
'    If documento = 2 Then
'        For K = 1 To 4
'            Grid2.AddItem ""
'        Next K
'
'        Grid2.Range(Grid2.Rows - 2, 9, Grid2.Rows - 1, 10).Merge
'        Grid2.Range(Grid2.Rows - 2, 9, Grid2.Rows - 1, 10).CellType = cellTextBox
'        Grid2.Range(Grid2.Rows - 2, 9, Grid2.Rows - 1, 10).FontBold = True
'        Grid2.Range(Grid2.Rows - 2, 9, Grid2.Rows - 1, 10).FontSize = 12
'        Grid2.Range(Grid2.Rows - 2, 9, Grid2.Rows - 1, 10).Alignment = cellRightCenter
'        Grid2.Cell(Grid2.Rows - 2, 9).text = Format(MONTOCHEQUE, "  ###,###,###,###")
'        Grid2.AddItem ""
'        Grid2.AddItem ""
'        Grid2.Range(Grid2.Rows - 1, 8, Grid2.Rows - 1, 10).FontBold = True
'        Grid2.Range(Grid2.Rows - 1, 8, Grid2.Rows - 1, 10).FontSize = 8
'        Grid2.Range(Grid2.Rows - 1, 8, Grid2.Rows - 1, 10).Merge
'        Grid2.Cell(Grid2.Rows - 1, 8).text = Format(fechacheque, "dd") & "  de  " & MonthName(Format(fechacheque, "mm")) & "  de  " & Format(fechacheque, "yyyy")
'        Grid2.Cell(Grid2.Rows - 1, 8).Alignment = cellRightTop
'
'        Grid2.AddItem ""
'        Grid2.AddItem ""
'
'        Grid2.Range(Grid2.Rows - 1, 5, Grid2.Rows - 1, 10).FontBold = True
'        Grid2.Range(Grid2.Rows - 1, 5, Grid2.Rows - 1, 10).FontSize = 9
'        Grid2.Range(Grid2.Rows - 1, 5, Grid2.Rows - 1, 10).Merge
'        Grid2.Range(Grid2.Rows - 1, 5, Grid2.Rows - 1, 10).Alignment = cellLeftCenter
'        Grid2.Cell(Grid2.Rows - 1, 5).text = "               " & CHEQUEGIRADO
'
'        Grid2.AddItem ""
'        Grid2.AddItem ""
'
'        Grid2.Range(Grid2.Rows - 1, 5, Grid2.Rows - 1, 10).FontBold = True
'        Grid2.Range(Grid2.Rows - 1, 5, Grid2.Rows - 1, 10).FontSize = 9
'        Grid2.Range(Grid2.Rows - 1, 5, Grid2.Rows - 1, 10).Merge
'        monto = WORDNUM(MONTOCHEQUE, "PESO", "", "CENTAVO", "CENTAVOS", 0)
'        monto = UCase(monto)
'
'        largo = Len(monto)
'        calzador = largo
'        If largo > 58 Then
'            For K = 58 To 1 Step -1
'                If Asc(Mid(monto, K, 1)) <> 32 Then calzador = 58 - (58 - K): Exit For
'            Next K
'            Grid2.Cell(Grid2.Rows - 1, 5).text = "               " & Mid(monto, 1, calzador) & "-"
'            Grid2.AddItem ""
'            Grid2.Range(Grid2.Rows - 1, 5, Grid2.Rows - 1, 10).FontBold = True
'            Grid2.Range(Grid2.Rows - 1, 5, Grid2.Rows - 1, 10).FontSize = 9
'            Grid2.Range(Grid2.Rows - 1, 5, Grid2.Rows - 1, 10).Merge
'            Grid2.Cell(Grid2.Rows - 1, 5).text = "               " & Mid(monto, calzador + 1, largo)
'        Else
'            Grid2.Cell(Grid2.Rows - 1, 5).text = "               " & monto
'            Grid2.Range(Grid2.Rows - 1, 5, Grid2.Rows - 1, 10).Merge
'        End If
'        Grid2.Cell(Grid2.Rows - 1, 5).text = Grid2.Cell(Grid2.Rows - 1, 5).text & "**************************************************"
'    End If
    'Grid2.PageSetup.PrintGridlines = True
    'Grid2.Range(40, 1, Grid2.Rows - 1, Grid2.Cols - 1).Borders(cellInsideHorizontal) = cellThin
    'Grid2.Range(40, 1, Grid2.Rows - 1, Grid2.Cols - 1).Borders(cellInsideVertical) = cellThin
    If documento = 1 Then
    Grid2.PageSetup.PaperWidth = 21
    Grid2.PageSetup.PaperHeight = 27
    End If

    If documento = 2 Then
    Grid2.PageSetup.PaperWidth = 21
    Grid2.PageSetup.PaperHeight = 20.3
    End If
    If documento = 3 Then
    Grid2.PageSetup.PaperWidth = 21
    Grid2.PageSetup.PaperHeight = 14
    End If
    
'
     For k = 1 To Grid2.PageSetup.PaperSizes.Count
            If UCase(Grid2.PageSetup.PaperSizes.item(k).PaperName) = "CARTA" Or UCase(Grid2.PageSetup.PaperSizes.item(k).PaperName) = "LETTER" Then
                Grid2.PageSetup.PaperSize = Grid2.PageSetup.PaperSizes.item(k).Kind
                Exit For
            End If
        Next k
    Grid2.PageSetup.BottomMargin = 1
    Grid2.PageSetup.TopMargin = 1
    Grid2.PageSetup.LeftMargin = 1.9
    Grid2.PageSetup.RightMargin = 0.5
    
  

   
    
    
    
    
    Grid2.PageSetup.PrintGridlines = False
    
    
    
    
    
    
    
 
    If VISTA = True Then
    Grid2.PrintPreview
    Else
    Grid2.DirectPrint
    End If
     
     frmimprimir.Visible = False
    Call retorno
End Sub

Sub imprimirSTANDAR(documento As Integer)
    Dim j As Integer
    Dim n As Integer
    Dim a As Integer
    Dim TITU As String
    Dim montocheque As Double
    Dim numerocheque As String
    Dim fechacheque As Date
    Dim CHEQUEGIRADO As String
    Dim IMPRIMECHEQUE As Boolean
    Dim largo As Integer
    Dim monto As String
    Dim calzador As Integer
    Dim FINAL As Double
    
    Rem DATOS DE LA COLUMNA
    FORMATOGRILLA(1, 1) = "CUENTA"
    FORMATOGRILLA(1, 2) = "L"
    FORMATOGRILLA(1, 3) = "CODIGO"
    FORMATOGRILLA(1, 4) = "TP"
    FORMATOGRILLA(1, 5) = "NUMERO"
    FORMATOGRILLA(1, 6) = "F.VENCI."
    FORMATOGRILLA(1, 7) = "GLOSA"
    FORMATOGRILLA(1, 8) = "DEBE"
    FORMATOGRILLA(1, 9) = "HABER"
    FORMATOGRILLA(1, 10) = "ANALISIS"
    
    
    Rem LARGO DE LOS DATOS
    
    FORMATOGRILLA(2, 1) = "19"
    FORMATOGRILLA(2, 2) = "2"
    FORMATOGRILLA(2, 3) = "8"
    FORMATOGRILLA(2, 4) = "3"
    FORMATOGRILLA(2, 5) = "9"
    FORMATOGRILLA(2, 6) = "8"
    FORMATOGRILLA(2, 7) = "30"
    FORMATOGRILLA(2, 8) = "9"
    FORMATOGRILLA(2, 9) = "9"
    FORMATOGRILLA(2, 10) = "9"
    FORMATOGRILLA(2, 11) = "9"
    
    Rem TIPO DE DATOS C=CEROS,S=STRING,N=NUMERICO,D=DATE,
    FORMATOGRILLA(3, 1) = "S"
    FORMATOGRILLA(3, 2) = "N"
    FORMATOGRILLA(3, 3) = "S"
    FORMATOGRILLA(3, 4) = "S"
    FORMATOGRILLA(3, 5) = "C"
    FORMATOGRILLA(3, 6) = "D"
    FORMATOGRILLA(3, 7) = "S"
    FORMATOGRILLA(3, 8) = "N"
    FORMATOGRILLA(3, 9) = "N"
    FORMATOGRILLA(3, 10) = "S"
    
    Rem FORMATO GRILLA
    FORMATOGRILLA(4, 1) = ""
    FORMATOGRILLA(4, 2) = ""
    FORMATOGRILLA(4, 3) = ""
    FORMATOGRILLA(4, 4) = ""
    FORMATOGRILLA(4, 5) = ""
    FORMATOGRILLA(4, 6) = "dd-mm-yyyy"
    FORMATOGRILLA(4, 7) = ""
    FORMATOGRILLA(4, 8) = " ###,###,##0"
    FORMATOGRILLA(4, 9) = " ###,###,##0"
    FORMATOGRILLA(4, 10) = ""
    FORMATOGRILLA(4, 11) = ""
    
    Rem LOCCKED
    FORMATOGRILLA(5, 1) = "FALSE"
    FORMATOGRILLA(5, 2) = "FALSE"
    FORMATOGRILLA(5, 3) = "FALSE"
    FORMATOGRILLA(5, 4) = "FALSE"
    FORMATOGRILLA(5, 5) = "FALSE"
    FORMATOGRILLA(5, 6) = "FALSE"
    FORMATOGRILLA(5, 7) = "FALSE"
    FORMATOGRILLA(5, 8) = "FALSE"
    FORMATOGRILLA(5, 9) = "FALSE"
    FORMATOGRILLA(5, 10) = "FALSE"
    
    Rem VALOR MINIMO
    FORMATOGRILLA(6, 1) = ""
    FORMATOGRILLA(6, 2) = ""
    FORMATOGRILLA(6, 3) = ""
    FORMATOGRILLA(6, 4) = ""
    FORMATOGRILLA(6, 5) = ""
    FORMATOGRILLA(6, 6) = ""
    FORMATOGRILLA(6, 7) = ""
    FORMATOGRILLA(6, 8) = ""
    FORMATOGRILLA(6, 9) = ""
    
    Rem VALOR MAXIMO
    FORMATOGRILLA(7, 1) = ""
    FORMATOGRILLA(7, 2) = ""
    FORMATOGRILLA(7, 3) = ""
    FORMATOGRILLA(7, 4) = ""
    FORMATOGRILLA(7, 5) = ""
    FORMATOGRILLA(7, 6) = ""
    FORMATOGRILLA(7, 7) = ""
    FORMATOGRILLA(7, 8) = ""
    FORMATOGRILLA(7, 9) = ""
    Grid2.Cols = 11
    Grid2.Rows = 1
    FINAL = Grid1.Rows
    
    If FINAL > 14 And documento = 1 Then FINAL = 15
    
    
    Grid2.Rows = FINAL
    Grid2.Range(1, 1, Grid2.Rows - 1, Grid2.Cols - 1).ClearText
    Grid2.Column(10).Alignment = cellRightGeneral
    
    
    For k = 1 To Grid2.Cols - 1
        Grid2.Cell(0, k).text = FORMATOGRILLA(1, k)
        Grid2.Column(k).MaxLength = Val(FORMATOGRILLA(2, k))
        Grid2.Column(k).FormatString = FORMATOGRILLA(4, k)
        Rem Grid2.Column(K).Locked = FORMATOGRILLA(5, K)
        If FORMATOGRILLA(3, k) = "N" Then Grid2.Column(k).Alignment = cellRightCenter
        If FORMATOGRILLA(3, k) = "D" Then Grid2.Column(k).CellType = cellCalendar
        'Grid1.Column(7).CellType = cellComboBox
    Next k
    
    Grid2.DefaultFont.Size = 6.5
    For k = 1 To Grid2.Cols - 1
        Grid2.Column(k).Width = Val(FORMATOGRILLA(2, k)) * (Grid2.DefaultFont.Size)
    Next k

    Grid2.PageSetup.Orientation = cellPortrait
    Grid2.PageSetup.PrintFixedRow = True
    Grid2.PageSetup.PrintGridlines = False

    Grid2.PageSetup.BlackAndWhite = True
    Grid2.PageSetup.BottomMargin = 1
    Grid2.PageSetup.TopMargin = 1
    Grid2.PageSetup.LeftMargin = 1.9
    Grid2.PageSetup.RightMargin = 0.5
    Grid2.Range(0, 1, 0, Grid2.Cols - 1).FontBold = True
    Grid2.Range(0, 1, 0, Grid2.Cols - 1).Borders(cellEdgeTop) = cellThin
    Grid2.Range(0, 1, 0, Grid2.Cols - 1).Borders(cellEdgeBottom) = cellThin
    'Grid2.Range(0, 1, 0, 3).Merge
    'Grid2.Cell(0, 1).text = "CUENTA"
    TITU = dato0.text + " " + leerNOMBREcomprobantes(dato0.text)
    
    
    Call cabezERA(TITU + " NUMERO " + dato1.text, dato1.text, fechacon)
    IMPRIMECHEQUE = False
    
    For j = 1 To Grid1.Rows - 1
                
       If FINAL = 15 And documento = 1 Then GoTo PASO:
        Grid2.Cell(j, 1).text = Grid1.Cell(j, 10).text
        Grid2.Cell(j, 2).text = j
        Grid2.Cell(j, 3).text = Grid1.Cell(j, 1).text + "." + Grid1.Cell(j, 2).text + "." + Grid1.Cell(j, 3).text
        Grid2.Cell(j, 4).text = Grid1.Cell(j, 5).text
        Grid2.Cell(j, 5).text = Grid1.Cell(j, 6).text
        Grid2.Cell(j, 6).text = Grid1.Cell(j, 7).text
        Grid2.Cell(j, 7).text = Grid1.Cell(j, 4).text
        If Grid1.Cell(j, 9).text = "D" Then Grid2.Cell(j, 8).text = Grid1.Cell(j, 8).text
        If Grid1.Cell(j, 9).text = "H" Then Grid2.Cell(j, 9).text = Grid1.Cell(j, 8).text
        If Grid1.Cell(j, 13).text <> "" Then Grid2.Cell(j, 10).text = Mid(Grid1.Cell(j, 13).text, 1, 9) + "-" + Mid(Grid1.Cell(j, 13).text, 10, 1)
        If Grid1.Cell(j, 14).text <> "" Then Grid2.Cell(j, 10).text = Grid1.Cell(j, 14).text
    
       
PASO:
        If Grid1.Cell(j, 5).text = "CH" Then
            montocheque = Grid1.Cell(j, 8).text
            fechacheque = Grid1.Cell(j, 7).text
            numerocheque = Grid1.Cell(j, 6).text
            CHEQUEGIRADO = Grid1.Cell(j, 4).text
            IMPRIMECHEQUE = True
        End If
    Next j
    
    
    
    If FINAL = 15 And documento = 1 Then
    
    Grid2.Range(1, 1, 1, 7).Merge
    Grid2.Range(1, 1, 1, 7).Alignment = cellCenterCenter
    Grid2.Range(1, 1, 1, 7).FontSize = 12
    Grid2.Range(1, 1, 1, 7).FontBold = True
    
    
    Grid2.Cell(1, 1).text = "**** IMPRIMIR COMPROBANTE MANUAL ******"
    
    
    End If
    
    
    For a = FINAL To 14
        Grid2.Rows = Grid2.Rows + 1
    Next a
    If documento <> 1 Then Grid2.Rows = Grid2.Rows + 1
    Grid2.Range(Grid2.Rows - 1, 1, Grid2.Rows - 1, Grid2.Cols - 1).Borders(cellEdgeTop) = cellThin
    Grid2.Cell(Grid2.Rows - 1, 7).text = "TOTAL COMPROBANTE CONTABLE"
    Grid2.Cell(Grid2.Rows - 1, 8).text = debe.Caption
    Grid2.Cell(Grid2.Rows - 1, 9).text = haber.Caption

    Grid2.AddItem ""
    Grid2.Range(Grid2.Rows - 1, 1, Grid2.Rows - 1, Grid2.Cols - 7).Merge
    Grid2.Cell(Grid2.Rows - 1, 1).text = "GLOSA :" + glosa.text

    Grid2.AddItem ""
    Grid2.AddItem ""
    Grid2.AddItem ""

    Grid2.Range(Grid2.Rows - 1, 1, Grid2.Rows - 1, Grid2.Cols - 1).Merge
    Grid2.Cell(Grid2.Rows - 1, 1).text = "              _______________________                       ______________________                           _______________________             ____________________"
    Grid2.AddItem ""

    Grid2.Range(Grid2.Rows - 1, 1, Grid2.Rows - 1, Grid2.Cols - 1).Merge
    Grid2.Cell(Grid2.Rows - 1, 1).text = "                  VºBº Gerencia                                            VºBº Contabilidad                                      Recibo Conforme                           Confec. " + USUARIOSISTEMA

    Rem CHEQUE SANTANDER
    
    
    
    If documento = 1 Then
        Grid2.PageSetup.LeftMargin = 0.5
        Grid2.Column(7).Width = Grid2.Column(7).Width - 20
        For k = 1 To 5
            Grid2.AddItem ""
        Next k
         Grid2.AddItem ""
        Grid2.Range(Grid2.Rows - 2, 7, Grid2.Rows - 1, 7).Merge
        Grid2.Range(Grid2.Rows - 2, 7, Grid2.Rows - 1, 7).CellType = cellTextBox
        Grid2.Range(Grid2.Rows - 2, 7, Grid2.Rows - 1, 7).FontBold = True
        Grid2.Range(Grid2.Rows - 2, 7, Grid2.Rows - 1, 7).FontSize = 12
        Grid2.Range(Grid2.Rows - 2, 7, Grid2.Rows - 1, 7).Alignment = cellLeftCenter
        
'        Grid2.Cell(Grid2.Rows - 2, 7).text = "  " + Format(montocheque, "##,###,###,###")
        
        Grid2.Cell(Grid2.Rows - 2, 7).text = Format(montocheque, "#  #  #  #  #  #  #  #  #  #  #  #  0") 'Format(montocheque, "##,###,###,###")
        Grid2.Cell(Grid2.Rows - 2, 7).Alignment = cellRightGeneral
        
    
        Grid2.AddItem ""
           Grid2.RowHeight(Grid2.Rows - 1) = 22
      
        
        Grid2.Range(Grid2.Rows - 1, 3, Grid2.Rows - 1, 10).FontBold = True
        Grid2.Range(Grid2.Rows - 1, 3, Grid2.Rows - 1, 10).FontSize = 12
        Grid2.Range(Grid2.Rows - 1, 5, Grid2.Rows - 1, 6).Merge
        Grid2.Cell(Grid2.Rows - 1, 5).Alignment = cellRightCenter
        Grid2.Cell(Grid2.Rows - 1, 5).text = "PUCON"
        Grid2.Cell(Grid2.Rows - 1, 7).text = Format(Format(fechacheque, "dd"), "0  #") & "    " & Format(Format(fechacheque, "mm"), "0  #") & "   " & Format(Format(fechacheque, "yyyy"), " # # # #")
        Grid2.Cell(Grid2.Rows - 1, 7).Alignment = cellRightBottom
        
        
        Grid2.RowHeight(Grid2.Rows - 1) = 22
        
'        Grid2.Cell(Grid2.Rows - 1, 5).text = " " & Format(Format(fechacheque, "dd"), "0  #")
'        Grid2.Cell(Grid2.Rows - 1, 6).text = " " & Format(Format(fechacheque, "mm"), "0  #")
'        Grid2.Cell(Grid2.Rows - 1, 8).text = Format(Mid(Format(fechacheque, "yyyy"), 3, 2), "0  #")
        
        
        
        
          Grid2.AddItem ""
        Grid2.AddItem ""
        
        Grid2.Range(Grid2.Rows - 1, 1, Grid2.Rows - 1, 8).FontBold = True
        Grid2.Range(Grid2.Rows - 1, 1, Grid2.Rows - 1, 8).FontSize = 9
        Grid2.Range(Grid2.Rows - 1, 1, Grid2.Rows - 1, 8).Merge
        Grid2.Cell(Grid2.Rows - 1, 1).text = "       " + CHEQUEGIRADO

    
        Grid2.AddItem ""
     
        
        'numToLet(MONTOCHEQUE, "PESO", "", "CENTAVO", "CENTAVOS", 0)
    
        
        Grid2.Range(Grid2.Rows - 1, 1, Grid2.Rows - 1, 7).FontBold = True
        Grid2.Range(Grid2.Rows - 1, 1, Grid2.Rows - 1, 7).FontSize = 9
        Grid2.Range(Grid2.Rows - 1, 1, Grid2.Rows - 1, 7).Merge
        monto = WORDNUM(montocheque, "PESO", "", "CENTAVO", "CENTAVOS", 0)
        monto = UCase(monto)
        'ariel revisa largo cheques
        largo = Len(monto)
        calzador = largo
        If largo > 60 Then
            For k = 60 To 1 Step -1
                If Asc(Mid(monto, k, 1)) <> 32 Then calzador = 60 - (60 - k): Exit For
            Next k
            Grid2.Cell(Grid2.Rows - 1, 1).text = "      " + Mid(monto, 1, calzador) + "-"
            Grid2.AddItem ""
            Grid2.Range(Grid2.Rows - 1, 1, Grid2.Rows - 1, 7).FontBold = True
            Grid2.Range(Grid2.Rows - 1, 1, Grid2.Rows - 1, 7).FontSize = 9
            Grid2.Range(Grid2.Rows - 1, 1, Grid2.Rows - 1, 7).Merge
            Grid2.Cell(Grid2.Rows - 1, 1).text = Mid(monto, calzador + 1, largo)
        Else
            Grid2.Cell(Grid2.Rows - 1, 1).text = "      " + monto
            Grid2.Range(Grid2.Rows - 1, 1, Grid2.Rows - 1, 7).Merge
        End If
        Grid2.Cell(Grid2.Rows - 1, 1).text = Grid2.Cell(Grid2.Rows - 1, 1).text & "**************************************************"
        
        Grid2.Rows = 49
        
        
Grid2.Cell(34, 1).Font.Bold = True
Grid2.Cell(34, 1).Font.Size = 10
Grid2.Cell(34, 1).text = empresaactiva & "  CH:" & Val(numerocheque)

        Grid2.Range(39, 1, 39, 8).FontBold = True
        Grid2.Range(39, 1, 39, 8).FontSize = 8
   
   Grid2.Range(39, 1, 39, 7).Merge
   Grid2.RowHeight(41) = 22
   Grid2.Cell(39, 1).text = "                       " & Format(fechacon, "dd") + "            " + MonthName(Format(fechacon, "mm")) + "    " + Format(fechacon, "yyyy")
   Grid2.Cell(41, 7).Alignment = cellRightCenter
   Grid2.Cell(41, 7).text = numerocheque
   Grid2.Cell(42, 1).text = "       Santander"
   Grid2.Range(42, 5, 42, 6).Merge
   
   Grid2.Cell(42, 5).Alignment = cellRightCenter
   Grid2.Cell(42, 5).text = Format(montocheque, "###,###,###") & "   "
   
'   Grid2.Range(46, 1, 42, 9).Merge
'   Grid2.Range(47, 1, 43, 9).Merge
'   Grid2.Range(48, 1, 44, 9).Merge
'   Grid2.Cell(46, 1).text = "SR : PROVEEDOR. CUALQUIER DUDA O ALCANCE SOBRE ESTE PAGO. HACERLO"
'   Grid2.Cell(47, 1).text = "EN UN PLAZO MAXIMO DE 30 DIAS POSTERIORES A LA FECHA DE PAGO"
'   Grid2.Cell(48, 1).text = "PASANDO DICHO PLAZO LA EMPRESA NO ACEPTARA RECLAMOS"
    End If
 
    If documento = 1 Then
    Grid2.PageSetup.PaperWidth = 21
    Grid2.PageSetup.PaperHeight = 27
    End If

    If documento = 2 Then
    Grid2.PageSetup.PaperWidth = 21
    Grid2.PageSetup.PaperHeight = 20.3
    End If
    If documento = 3 Then
    Grid2.PageSetup.PaperWidth = 21
    Grid2.PageSetup.PaperHeight = 14
    End If
    
'
     For k = 1 To Grid2.PageSetup.PaperSizes.Count
            If UCase(Grid2.PageSetup.PaperSizes.item(k).PaperName) = "CARTA" Or UCase(Grid2.PageSetup.PaperSizes.item(k).PaperName) = "LETTER" Then
                Grid2.PageSetup.PaperSize = Grid2.PageSetup.PaperSizes.item(k).Kind
                Exit For
            End If
        Next k
    Grid2.PageSetup.BottomMargin = 1
    Grid2.PageSetup.TopMargin = 1
    Grid2.PageSetup.LeftMargin = 1.9
    Grid2.PageSetup.RightMargin = 0.5
    
  

   
    
    
    
    
    Grid2.PageSetup.PrintGridlines = False
    
    
    
    
    
    
    
 
    If VISTA = True Then
    Grid2.PrintPreview
    Else
    Grid2.DirectPrint
    End If
     
     frmimprimir.Visible = False
    Call retorno
End Sub

Sub IMPRIMIR2()
    Dim j As Integer
    Dim n As Integer
    Dim a As Integer
    Dim TITU As String
    Dim montocheque As Double
    Dim numerocheque As String
    Dim fechacheque As Date
    Dim CHEQUEGIRADO As String
    Dim IMPRIMECHEQUE As Boolean
    Dim largo As Integer
    Dim monto As String
    Dim calzador As Integer
    Dim FINAL As Double
    Dim lin2 As Double
Grid2.ReportTitles.Clear
   
    For j = 1 To Grid1.Rows - 1
                
        If Grid1.Cell(j, 5).text = "CH" Then
            montocheque = Grid1.Cell(j, 8).text
            fechacheque = Grid1.Cell(j, 7).text
            numerocheque = Grid1.Cell(j, 6).text
            CHEQUEGIRADO = Grid1.Cell(j, 4).text
            IMPRIMECHEQUE = True
        End If
    Next j
    
   
    Rem CHEQUE SANTANDER suelto
    
         Grid2.ReportTitles.Clear
         
         
        Grid2.PageSetup.TopMargin = 0
        
        Grid2.PageSetup.LeftMargin = 0
        
        Grid2.Cols = 5
        Grid2.Column(1).Width = 300
        Grid2.Column(2).Width = 100
        Grid2.Column(3).Width = 80
        Grid2.Column(4).Width = 80
        
        Grid2.Rows = 1
        Grid2.Rows = 10
        Grid2.AddItem ""
        Grid2.Range(1, 3, 1, 4).Merge
        Grid2.Range(1, 3, 1, 4).CellType = cellTextBox
        Grid2.Range(1, 3, 1, 4).FontBold = True
        Grid2.Range(1, 3, 1, 4).FontSize = 12
        Grid2.Range(1, 3, 1, 4).Alignment = cellLeftCenter
        
        Grid2.Cell(1, 3).text = "  " + Format(montocheque, "##,###,###,###") & ".="
    
        Grid2.Range(2, 1, 10, 4).FontBold = True
        Grid2.Range(2, 1, 10, 4).FontSize = 9
        
        
        Grid2.Cell(4, 2).text = Format(fechacheque, "dd") + " "
        Grid2.Cell(4, 2).Alignment = cellLeftCenter
        
        
        Grid2.Cell(4, 3).text = MonthName(Format(fechacheque, "mm"))
        Grid2.Cell(4, 3).Alignment = cellLeftCenter
        
        
        Grid2.Cell(4, 4).text = Format(fechacheque, "yyyy")
        Grid2.Cell(4, 4).Alignment = cellRightBottom
        
        Grid2.Range(5, 1, 5, 4).FontBold = True
        Grid2.Range(5, 1, 5, 4).FontSize = 9
        Grid2.Range(5, 1, 5, 4).Merge
        Grid2.Cell(5, 1).text = "                " + CHEQUEGIRADO

    
        
        'numToLet(MONTOCHEQUE, "PESO", "", "CENTAVO", "CENTAVOS", 0)
    
        Grid2.RowHeight(6) = 20
        
        Grid2.Range(6, 1, 6, 4).FontBold = True
        Grid2.Range(6, 1, 6, 4).FontSize = 9
        Grid2.Range(6, 1, 6, 4).Merge
        monto = WORDNUM(montocheque, "PESO", "", "CENTAVO", "CENTAVOS", 0)
        monto = UCase(monto)
        
        largo = Len(monto)
        calzador = largo
        If largo > 60 Then
            For k = 60 To 1 Step -1
                If Asc(Mid(monto, k, 1)) <> 32 Then calzador = 60 - (60 - k): Exit For
            Next k
            Grid2.Cell(6, 1).text = "              " + Mid(monto, 1, calzador) + "-"
          
            Grid2.Range(7, 1, 7, 4).FontBold = True
            Grid2.Range(7, 1, 7, 4).FontSize = 9
            Grid2.Range(7, 1, 7, 4).Merge
            Grid2.Cell(7, 1).text = Mid(monto, calzador + 1, largo)
            lin2 = 7
        Else
            Grid2.Cell(6, 1).text = "              " + monto
            Grid2.Range(6, 1, 7, 3).Merge
            lin2 = 6
        End If
        
        Grid2.Cell(lin2, 1).text = "             " + Grid2.Cell(lin2, 1).text & "**************************************************"
        
         Grid2.Cell(9, 1).Font.Bold = True
         Grid2.Cell(9, 1).Font.Size = 10
         Grid2.Cell(9, 1).text = "           " & empresaactiva + "  CH:" + numerocheque

    Grid2.Cell(0, 1).text = ""
    Grid2.Cell(0, 2).text = ""
    Grid2.Cell(0, 3).text = ""
    Grid2.Cell(0, 4).text = ""

    
    Grid2.PageSetup.PrintGridlines = False
    
    
    
    
    
    
    
 
    
    Grid2.PrintPreview
     frmimprimir.Visible = False
    Call retorno
End Sub
 
Sub cabezERA(titulo As String, numero As String, fecha As Date)
Dim objReportTitle As FlexCell.ReportTitle
Grid2.ReportTitles.Clear


Set objReportTitle = New FlexCell.ReportTitle
    objReportTitle.text = titulo
    objReportTitle.Font.Name = "verdana"
    objReportTitle.Font.Size = 9
    objReportTitle.PrintOnAllPages = True
    Grid2.ReportTitles.Add objReportTitle
    
    'Report Title 1
    For k = 1 To 5
    Set objReportTitle = New FlexCell.ReportTitle
    If k < 5 Then objReportTitle.text = DATOSEMPRESA(k)
    
    
    If k = 5 Then objReportTitle.text = DATOSEMPRESA(k) + "       FECHA COMPROBANTE:" & fechacon & "                                                                           EMITIDO   :" & Str(fechasistema)
    objReportTitle.Font.Name = "verdana"
    objReportTitle.Font.Size = 7
    objReportTitle.Font.Italic = True
    objReportTitle.PrintOnAllPages = True
    objReportTitle.color = RGB(128, 0, 0)
    objReportTitle.Align = CellLeft
    Grid2.ReportTitles.Add objReportTitle
    Next k
'    Set objReportTitle = New FlexCell.ReportTitle
'    objReportTitle.text = "CENTRO DE COSTO " + centro.Caption
'
'
'    objReportTitle.Font.Name = "verdana"
'    objReportTitle.Font.Size = 10
'    objReportTitle.Font.Italic = True
'    objReportTitle.PrintOnAllPages = True
'    objReportTitle.Color = RGB(128, 0, 0)
'    objReportTitle.Align = cellCenter
'
'    Grid2.ReportTitles.Add objReportTitle
'
End Sub

Sub ELIMINA()
    Dim tipo2 As String
    Dim i As Long
    Call ACTUALIZADOCUMENTO("-")
    For i = 1 To Grid1.Rows - 1
        If Grid1.Cell(i, 5).text = "FC" Then Call abonofactura(i, 8, True)
        If Grid1.Cell(i, 5).text = "EE" Then Call abonofactura(i, 8, True)
        If Grid1.Cell(i, 5).text = "BH" Then Call abonoboletadehonorarios(i, 8, True)
        If Grid1.Cell(i, 5).text = "DM" Then Call abonoGUIADEVOLUCION(Grid1.Cell(i, 5).text, Grid1.Cell(i, 6).text, "", "", "", "0")
        If Grid1.Cell(i, 5).text = "D1" Then Call abonoGUIADEVOLUCION(Grid1.Cell(i, 5).text, Grid1.Cell(i, 6).text, "", "", "", "0")
        
        If Grid1.Cell(i, 5).text = "FP" Then
        Call abonopublicidad("1", Grid1.Cell(i, 6).text, Grid1.Cell(i, 8).text)
        End If
        
        If Grid1.Cell(i, 5).text = "FA" And Grid1.Cell(i, 1).text + Grid1.Cell(i, 2).text + Grid1.Cell(i, 3).text = "11200028" And Grid1.Cell(i, 9).text = "H" Then
        Call abonopublicidad("2", Grid1.Cell(i, 6).text, Grid1.Cell(i, 8).text)
        End If
        
        
        If Mid(Grid1.Cell(i, 4).text, 1, 14) = "CANC. DONACION" Then
'            Stop
             If MODIFI = 0 Then
                Call actualizaliquidacion(Grid1.Cell(i, 13).text, Grid1.Cell(i, 8).text, dato3.text, dato4.text, "00097")
             End If
        End If
        
        '''''''''''''''''''''''''''''''''''''''''''''
        ''''''''''ESTO ESTABA CMENTADO,
        ''''''''''LO DESCOMENTE Y AJUSTE LA CONDICION
        ''''''''''PARA QUE ELIMINE LOS CHEQUES
        If Grid1.Cell(i, 5).text = "CH" Then
            campos(0, 0) = ""
            campos(0, 2) = "chequesdocumento"
            condicion = "cuenta = '" & Grid1.Cell(i, 1).text & Grid1.Cell(i, 2).text & Grid1.Cell(i, 3).text & "' AND numero = '" & Grid1.Cell(i, 6).text & "' AND tipocomprobante = '" & dato0.text & "' AND numerocomprobante = '" & dato1.text & "'"
            op = 4
            sqlconta.response = campos
            Set sqlconta.conexion = contadb
            Call sqlconta.sqlconta(op, condicion)
        End If
        ''''''''''ESTO ESTABA CMENTADO,
        ''''''''''LO DESCOMENTE Y AJUSTE LA CONDICION
        ''''''''''PARA QUE ELIMINE LOS CHEQUES
        '''''''''''''''''''''''''''''''''''''''''''''
    Next i
    campos(0, 0) = ""
    campos(0, 2) = "movimientoscontables"
    condicion = "tipo = '" + dato0.text + "' and numero = '" + dato1.text + "' and año = '" + Format(fechasistema, "yyyy") + "' and mes = '" + Format(fechasistema, "mm") + "' order by numero desc"
    op = 4
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)
    
    campos(0, 0) = ""
    campos(0, 2) = "movimientos_glosa"
    condicion = "tipo = '" + dato0.text + "' and numero = '" + dato1.text + "' and fecha = '" + dato4.text + "-" + dato3.text + "-" + dato2.text + "'"
    op = 4
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)
    
    If empresaactiva = "28" Then
        If dato0.text = "CD" Then
            Call actualizacastigados
        End If
    End If
    'If sqlconta.status = 4 Then Stop
End Sub
Sub actualizacastigados()
     Dim csql As New rdoQuery
     Dim resultados As rdoResultset
     Set csql.ActiveConnection = contadb
     csql.sql = "update " & cliente_sql & "ventas.sv_cuotas_detalle "
     csql.sql = csql.sql & "set fecha_castigo='0000-00-00' where fecha_castigo like '" & dato4.text & "-" & dato3.text & "%' "
     csql.Execute
     
     Call sincronizadatos(csql.sql, conta, Servidor)
     csql.sql = "delete from " & cliente_sql & "ventas.sv_castigos_tributarios "
     csql.sql = csql.sql & " where fecha_castigo like '" & dato4.text & "-" & dato3.text & "%'"
     csql.Execute
     Call sincronizadatos(csql.sql, conta, Servidor)
     csql.Close
     Set csql = Nothing
     
End Sub
Private Sub actualizaliquidacion(ruttrabajador, monto, MES, año, codigo)
     Dim csql As New rdoQuery
     Dim resultados As rdoResultset
     Dim csql1 As New rdoQuery
     
     Set csql.ActiveConnection = contadb
     csql.sql = "update " & cliente_sql & "remu" & empresaactiva & ".liquidacionhd "
     csql.sql = csql.sql & "set monto=monto-'" & Replace(Replace(monto, ".", ""), ",", ".") & "' "
     csql.sql = csql.sql & " where rut='" & ruttrabajador & "' and mes='" & MES & "' and año='" & año & "' and codtablacalculo='" & codigo & "' "
     csql.Execute
     
     csql.sql = "select monto from " & cliente_sql & "remu" & empresaactiva & ".liquidacionhd "
    csql.sql = csql.sql & " where rut='" & ruttrabajador & "' and mes='" & MES & "' and año='" & año & "' and codtablacalculo='" & codigo & "' "
    csql.Execute
    If csql.RowsAffected > 0 Then
        Set resultados = csql.OpenResultset
        If resultados(0) <= 0 Then
            Set csql1.ActiveConnection = contadb
            csql1.sql = "delete from " & cliente_sql & "remu" & empresaactiva & ".liquidacionhd "
            csql1.sql = csql1.sql & " where rut='" & ruttrabajador & "' and mes='" & MES & "' and año='" & año & "' and codtablacalculo='" & codigo & "' "
            csql1.Execute
            csql1.Close
            Set csql1 = Nothing
        End If
    End If
    csql.Close
    Set csql = Nothing
    
 End Sub


Private Sub eliminarCheque(ByVal cuenta As String, CHEQUE As String)
    campos(0, 0) = ""
    campos(0, 2) = "chequesdocumento"
    condicion = "cuenta = '" & cuenta & "' AND numero = '" & CHEQUE & "' "
    op = 4
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)
End Sub

Sub esfecha(ByRef dias As Integer, ByRef MES As Integer, ByRef ano As Integer)
    
    If dias < 1 Or dias > 31 Then ef = "N": GoTo no
    If MES < 1 Or MES > 12 Then ef = "N": GoTo no
    If ano < 1999 Then ef = "N": GoTo no:
    ef = "S"
no:
End Sub
Sub leermayor(row As Long, col As Long)
    Dim cuenta As String
    
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
    cuenta = Grid1.Cell(row, 1).text + Grid1.Cell(row, 2).text + Grid1.Cell(row, 3).text
    condicion = "codigo=" + "'" + Grid1.Cell(row, 1).text + Grid1.Cell(row, 2).text + Grid1.Cell(row, 3).text + "' and año='" + Format(fechasistema, "yyyy") + "'"
    op = 5
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)
    
        If PermisosCuentasDelMayor(USUARIOSISTEMA, Format(Grid1.Cell(row, 1).text + Grid1.Cell(row, 2).text + Grid1.Cell(row, 3).text, "00000000")) = False Then
    MsgBox "USTED NO TIENE PRIVILEGIOS PARA ACCEDER A ESTA CUENTA", vbCritical, "ATENCION"
    Grid1.Cell(row, 1).text = ""
    Grid1.Cell(row, 2).text = ""
    Grid1.Cell(row, 3).text = ""
    
    Grid1.Cell(row, col).SetFocus
    
    
    Exit Sub
    End If
    
    
 
    If sqlconta.status = 4 Or Mid(sqlconta.response(0, 3), 5, 4) = "0000" Then
    
    Grid1.Cell(row, 1).text = ""
    Grid1.Cell(row, 2).text = ""
    Grid1.Cell(row, 3).text = ""
    
    respuesta = "N"
    Else
    respuesta = "S"
    Grid1.Cell(row, 10).text = sqlconta.response(1, 3)
    TIENECTACTE(row) = sqlconta.response(2, 3)
    TIENECRCC(row) = sqlconta.response(3, 3)
    TIENEBANCO(row) = sqlconta.response(4, 3)
    TIENEILA(row) = sqlconta.response(5, 3)
    TIENEICA(row) = sqlconta.response(6, 3)
    TIENEIHA(row) = sqlconta.response(7, 3)
    TIENEACTIVO(row) = sqlconta.response(8, 3)
    CUENTAMAYOR(row) = sqlconta.response(0, 3)
        
    If TIENECTACTE(row) = "1" Then
       digitarut.lblcuenta.Caption = sqlconta.response(0, 3)
        digitarut.lblglosa.Caption = sqlconta.response(1, 3)
        If digitarut.Visible = False Then
        If MODIFI = 1 Then
        digitarut.DATO20.text = Grid1.Cell(row, 13).text
        End If
        'Ariel envia tipo de comprobante al  form digitarut
        digitarut.tipo.Caption = dato0.text '<-----    aqui
        'UNICO CAMBIO
        Load digitarut
        digitarut.DATO20.text = Grid1.Cell(Grid1.ActiveCell.row, 13).text
        If MODIFI = 1 And digitarut.DATO20.text <> "" Then
            Call digitarut.DATO20_KeyPress(13)
        End If
        digitarut.Show vbModal
        Else
        digitarut.SetFocus
        
        End If
        If DIGITA_RUT_RUT = "" Then MsgBox "DEBE INGRESEAR RUT DE CLIENTE "
        Grid1.Cell(row, 11).text = DIGITA_RUT_NOMBRE
        Grid1.Cell(row, 13).text = DIGITA_RUT_RUT
    Else
        Grid1.Cell(row, 11).text = ""
        Grid1.Cell(row, 13).text = ""
   
    End If
    
    If TIENECRCC(row) = "1" Then
        If MODIFI = 1 Then
        digitacrcc.DATO21.text = Mid(Grid1.Cell(row, 12).text, 1, 2)
        digitacrcc.DATO22.text = Mid(Grid1.Cell(row, 12).text, 3, 2)
        
        End If
       
        digitacrcc.Show vbModal
        
        Grid1.Cell(row, 12).text = DIGITA_CRCC_NOMBRE
        Grid1.Cell(row, 14).text = DIGITA_CRCC_CODIGO
    Else
        Grid1.Cell(row, 12).text = ""
        Grid1.Cell(row, 14).text = ""
    End If
    
    
    If TIENEANALISIS(sqlconta.response(0, 3)) = True Then
      
    digitaanalisis.lblcuenta.Caption = sqlconta.response(0, 3)
    digitaanalisis.lblglosa.Caption = sqlconta.response(1, 3)
    
    digitaanalisis.Show vbModal
    Grid1.Cell(row, 15).text = DIGITA_ANALISIS_CODIGO + ":" + DIGITA_ANALISIS_NOMBRE
    Grid1.Cell(row, 16).text = DIGITA_CENTROS_CODIGO + ":" + DIGITA_CENTROS_NOMBRE
    
    Else
    Grid1.Cell(row, 15).text = ""
    Grid1.Cell(row, 16).text = ""
    
    End If
    
    


End If
End Sub


Sub SIGUIENTE()
    campos(0, 0) = "tipo"
    campos(1, 0) = "numero"
    campos(2, 0) = ""
    campos(0, 2) = "movimientoscontables"
    condicion = "tipo='" + dato0.text + "' and numero>'" + dato1.text + "' and año='" + Format(fechasistema, "yyyy") + "' and mes='" + Format(fechasistema, "mm") + "' order by numero asc"
    op = 5
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)
    If sqlconta.status = 0 Then dato0.text = sqlconta.response(0, 3): dato1.text = sqlconta.response(1, 3): leecomprobante
End Sub

Sub anterior()
    campos(0, 0) = "tipo"
    campos(1, 0) = "numero"
    campos(2, 0) = ""
    campos(0, 2) = "movimientoscontables"
    condicion = "tipo='" + dato0.text + "' and numero<'" + dato1.text + "' and año='" + Format(fechasistema, "yyyy") + "' and mes='" + Format(fechasistema, "mm") + "' order by numero desc"
    '" + DATO1.text + "' order by numero desc"
    op = 5
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)
    If sqlconta.status = 0 Then dato0.text = sqlconta.response(0, 3): dato1.text = sqlconta.response(1, 3): leecomprobante
    
End Sub

Private Sub opciones_GotFocus()
   Rem MANUAL.SetFocus
End Sub

Sub ayudamayor(row As Long, col As Long)
    Dim campos As Variant
    Dim cfijo As Variant
    Dim largo As Variant
    campos = Array("codigo", "nombre")
    cabezas = Array("codigo", "nombre")
    largo = Array("8n", "40s")
    mensajeAyuda = "Ayuda Cuentas del Mayor"
    cfijo = "año='" + dato4.text + "'"
    pivote2.MaxLength = 8
    
    Call cargaAyudaT(Servidor, basebus, Usuario, password, "cuentasdelmayor", pivote2, campos, cfijo, largo, 2)
    Grid1.Cell(row, col).text = Mid(pivote2.text, 1, 2)
    Grid1.Cell(row, col + 1).text = Mid(pivote2.text, 3, 2)
    Grid1.Cell(row, col + 2).text = Mid(pivote2.text, 5, 4)
    Call leermayor(row, col)
    

End Sub

Sub SUMAR()
sumadebe = 0
sumahaber = 0

For o = 1 To Grid1.Rows - 1
If Grid1.Cell(o, 9).text = "D" Then sumadebe = sumadebe + CDbl(Grid1.Cell(o, 8).text)
If Grid1.Cell(o, 9).text = "H" Then sumahaber = sumahaber + CDbl(Grid1.Cell(o, 8).text)
Next o
debe.Caption = Format(sumadebe, "###,###,###,##0")
haber.Caption = Format(sumahaber, "###,###,###,##0")
saldo.Caption = Format(sumadebe - sumahaber, "###,###,###,##0")
End Sub

Sub ayudatipos(ByRef caja As TextBox)
    Dim campos As Variant
    Dim cfijo As Variant
    Dim largo As Variant
    campos = Array("tipos", "nombredocumento")
    cabezas = Array("TIPOS", "NOMBRE")
    largo = Array("4N", "40s")
    mensajeAyuda = "Ayuda Tipos de Documentos"
    cfijo = "no"
    
    Call cargaAyudaT(Servidor, clientesistema + "conta", Usuario, password, "maestrotipodedocumentos", caja, campos, cfijo, largo, 2)
    If caja.text = "" Then caja.SetFocus: GoTo no
    Call leetipos(caja)
    caja.Enabled = True
    caja.SetFocus


no:

End Sub

Sub leetipos(caja As TextBox)
    
    campos(0, 0) = "tipos"
    campos(1, 0) = "nombredocumento"
    campos(2, 0) = ""
    
    campos(0, 2) = "maestrotipodedocumentos"
    condicion = "tipos=" + "'" + caja.text + "'"
    op = 5
    sqlconta.response = campos
    Set sqlconta.conexion = conta
    Call sqlconta.sqlconta(op, condicion)

    If sqlconta.status = 4 Then caja.text = "": caja.SetFocus:  GoTo no:
    tipocompro.Caption = sqlconta.response(1, 3)
    
    

no:

End Sub


Sub leetipos2(caja As TextBox)
    
    campos(0, 0) = "tipos"
    campos(1, 0) = "nombredocumento"
    campos(2, 0) = ""
    
    campos(0, 2) = "maestrotipodedocumentos"
    condicion = "tipos=" + "'" + caja.text + "'"
    op = 5
    sqlconta.response = campos
    Set sqlconta.conexion = conta
    Call sqlconta.sqlconta(op, condicion)

    If sqlconta.status = 4 Then caja.text = "": caja.SetFocus:  GoTo no:
    Label2.Caption = sqlconta.response(1, 3)
    
    

no:

End Sub
Sub ultimo()
    campos(0, 0) = "tipo"
    campos(1, 0) = "MAX(numero)"
    campos(2, 0) = ""
    campos(0, 2) = "movimientoscontables"
    condicion = "tipo='" + dato0.text + "' and año='" + Format(fechasistema, "yyyy") + "' and mes='" + Format(fechasistema, "mm") + "' "
   op = 5
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)
    dato1.text = sqlconta.response(1, 3) + 1
    Call ceros(dato1)
      
    Rem aca
    End Sub

Sub grabacheque(row)
    campos(0, 0) = "cuenta"
    campos(1, 0) = "numero"
    campos(2, 0) = "emision"
    campos(3, 0) = "monto"
    campos(4, 0) = "vencimiento"
    campos(5, 0) = "tipocomprobante"
    campos(6, 0) = "numerocomprobante"
    campos(7, 0) = "giradoa"
    campos(8, 0) = "ubicacion"
    campos(9, 0) = "tipodocumento"
    campos(10, 0) = "fechamovimiento"
    campos(11, 0) = ""
    campos(0, 1) = Grid1.Cell(k, 1).text + Grid1.Cell(k, 2).text + Grid1.Cell(k, 3).text
    campos(1, 1) = Grid1.Cell(k, 6).text
    campos(2, 1) = dato4.text + dato3.text + dato2.text
    campos(3, 1) = Replace(Grid1.Cell(k, 8).text, ",", ".")
    campos(4, 1) = Format(Grid1.Cell(k, 7).text, "yyyy-mm-dd")
    campos(5, 1) = dato0.text
    campos(6, 1) = dato1.text
    campos(7, 1) = Grid1.Cell(k, 4).text
    campos(8, 1) = "0"
    campos(9, 1) = "CH"
    campos(10, 1) = Format(fechasistema, "yyyy-mm-dd")
    
    campos(0, 2) = "chequesdocumento"
       
    op = 2
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)
End Sub

Sub leercheque(k)
    campos(0, 0) = "cuenta"
    campos(1, 0) = "numero"
    campos(2, 0) = ""
    campos(0, 1) = Grid1.Cell(k, 1).text + Grid1.Cell(k, 2).text + Grid1.Cell(k, 3).text
    campos(1, 1) = Grid1.Cell(k, 6).text
    campos(0, 2) = "chequesdocumento"
    condicion = "cuenta='" + campos(0, 1) + "' and  numero='" + campos(1, 1) + "'"
    op = 5
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)
    If sqlconta.status = 0 Then chequecreado = "S" Else chequecreado = "N"
End Sub


Sub ayudacompras(row As Long, col As Long)
    Dim campos As Variant
    Dim cfijo As Variant
    Dim largo As Variant
    campos = Array("numero", "tipo", "fecha", "fechavencimiento", "total", "abono")
    cabezas = Array("numero", "tipo", "fecha", "vencimiento", "total", "abono")
    largo = Array("10n", "5n", "10d", "10d", "10n", "10n")
    mensajeAyuda = "Ayuda Facturas de Compras"
    cfijo = "rut ='" + Grid1.Cell(row, 13).text + "' and total<>abono"

    Call cargaAyudaT(Servidor, basebus, Usuario, password, "facturasdecompras", general, campos, cfijo, largo, 6)
    If Val(general.text) = 0 Then general.text = "0000000000"
    Grid1.Cell(row, col).text = general.text
    
End Sub
Sub leefacturadecompra(row As Long, col As Long)
    Dim ABONO As Double
    Dim Cantidad As Double
    Dim tpdoc As String
    Dim k As Integer
    Dim NoRevisaFactorizada As Boolean
    PAGARFACTURA = False
    
     
    
 Rem   campos(0, 0) = "total-abono"
    campos(0, 0) = "total"
    campos(1, 0) = "tipo"
    campos(2, 0) = "fecha"
    campos(3, 0) = ""
    campos(0, 2) = "facturasdecompras"
    condicion = "(tipo='1' or tipo='4' or tipo='7' or tipo='0') and rut='" + Grid1.Cell(Grid1.ActiveCell.row, 13).text + "' and numero='" + Grid1.Cell(Grid1.ActiveCell.row, 6).text + "'  order by fecha desc"
    op = 5
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)
    If sqlconta.status = 0 Then
    
    facturano = "S"
    Else
    facturano = "N"
    End If
    
    If sqlconta.response(1, 3) = "1" Then tpdoc = "FC"
    If sqlconta.response(1, 3) = "2" Then tpdoc = "DC"
    If sqlconta.response(1, 3) = "3" Then tpdoc = "NC"
    If sqlconta.response(1, 3) = "4" Then tpdoc = "FC"
    If sqlconta.response(1, 3) = "5" Then tpdoc = "DC"
    If sqlconta.response(1, 3) = "6" Then tpdoc = "NC"
    If sqlconta.response(1, 3) = "7" Then tpdoc = "FP"
    If sqlconta.response(1, 3) = "8" Then tpdoc = "IM"
    If sqlconta.response(1, 3) = "0" Then tpdoc = "EE"
    If sqlconta.response(1, 3) = "9" Then tpdoc = "EN"
    
    If Verifica_Permiso(Me.Caption, "autoriza") = True Then
        NoRevisaFactorizada = True
    End If
    
    Grid1.Cell(Grid1.ActiveCell.row, 8).text = sqlconta.response(0, 3) - leerabonofactura(sqlconta.response(1, 3), tpdoc, Grid1.Cell(Grid1.ActiveCell.row, 6).text, Grid1.Cell(Grid1.ActiveCell.row, 13).text, CUENTAPROVEEDOR, "D", sqlconta.response(2, 3), NoRevisaFactorizada)
    Grid1.Cell(Grid1.ActiveCell.row, 9).text = "D"
    Grid1.Cell(Grid1.ActiveCell.row, 8).Locked = True
    Grid1.Cell(Grid1.ActiveCell.row, 9).Locked = True
    
    
    If sqlconta.response(0, 3) = 0 Then
    PAGARFACTURA = False
    Else
    PAGARFACTURA = True
    
    End If
    Cantidad = 0
    For k = 1 To Grid1.Rows - 1
    If Grid1.Cell(Grid1.ActiveCell.row, 6).text = Grid1.Cell(k, 6).text Then
    Cantidad = Cantidad + 1
    End If
    
    Next k
    If Cantidad > 1 Then
    MsgBox ("factura ya esta cancelada en este comprobante ")
    facturano = "R"
    End If
    
End Sub


'
'Sub leefacturadecompra(Row As Long, col As Long)
'    Dim ABONO As Double
'    Dim Cantidad As Double
'    Dim TP2 As String
'
'    Dim k As Integer
'
'    PAGARFACTURA = False
'
'    campos(0, 0) = "total-abono"
'    campos(1, 0) = ""
'    campos(0, 2) = "facturasdecompras"
'    If Grid1.Cell(Row, 5).text = "FC" Then TP2 = "(TIPO='1' OR TIPO='4' OR TIPO='7')"
'    If Grid1.Cell(Row, 5).text = "EE" Then TP2 = "TIPO='0'"
'
'
'
'    condicion = TP2 + " and rut='" + Grid1.Cell(Row, 13).text + "' and numero='" + Grid1.Cell(Row, 6).text + "' order by numero desc"
'    op = 5
'    sqlconta.response = campos
'    Set sqlconta.conexion = contadb
'    Call sqlconta.sqlconta(op, condicion)
'    If sqlconta.status = 0 Then
'    facturano = "S"
'    Else
'    facturano = "N"
'    End If
'    Grid1.Cell(Grid1.ActiveCell.Row, 8).text = sqlconta.response(0, 3)
'    Grid1.Cell(Grid1.ActiveCell.Row, 9).text = "D"
'    Grid1.Cell(Grid1.ActiveCell.Row, 8).Locked = True
'    Grid1.Cell(Grid1.ActiveCell.Row, 9).Locked = True
'
'
'    If sqlconta.response(0, 3) = 0 Then
'    PAGARFACTURA = False
'    Else
'    PAGARFACTURA = True
'
'    End If
'    Cantidad = 0
'    For k = 1 To Grid1.Rows - 1
'    If Grid1.Cell(Grid1.ActiveCell.Row, 6).text = Grid1.Cell(k, 6).text Then
'    Cantidad = Cantidad + 1
'    End If
'
'    Next k
'    If Cantidad > 1 Then
'    MsgBox ("factura ya esta cancelada en este comprobante ")
'    facturano = "R"
'    End If
'
'End Sub

Sub leecentrofactura(tipo, numero, rut)
    campos(0, 0) = "centrodecosto"
    campos(1, 0) = ""
    campos(0, 2) = "facturasdecompras_detalle"
    condicion = "tipo='" + "1" + "' and rut='" + rut + "' and numero='" + numero + "' order by numero desc"
    op = 5
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)
    If sqlconta.status = 0 Then
    centro.Caption = leerNOMBREcrcc(sqlconta.response(0, 3))
    Else
    centro.Caption = ""
    End If
End Sub


Sub abonofactura(row As Long, col As Long, ByVal ELIMINA As Boolean)
    Dim csql As rdoQuery
    Set csql = New rdoQuery
    Set csql.ActiveConnection = contadb
    Dim TP2 As String
    
    If ELIMINA = True Then
        csql.sql = "update facturasdecompras set abono = abono + " & -1 * CDbl(Grid1.Cell(row, col).text) & " "
    Else
        csql.sql = "update facturasdecompras set abono = abono + " & CDbl(Grid1.Cell(row, col).text) & " "
    End If
    If Grid1.Cell(row, 5).text = "FC" Then TP2 = "(TIPO='1' OR TIPO='4' OR TIPO='7')"
    If Grid1.Cell(row, 5).text = "EE" Then TP2 = "TIPO='0'"
    
    csql.sql = csql.sql & "where " + TP2 + " and rut='" + Grid1.Cell(row, 13).text + "' and numero='" + Grid1.Cell(row, 6).text + "'"
    csql.Execute
'    Call sincronizadatos(csql.sql, contadb, "")
    
    csql.Close
    Set csql = Nothing

    Set csql = New rdoQuery
    Set csql.ActiveConnection = contadb
    If ELIMINA = False Then
        csql.sql = "update " + clientesistema + "consumos_basicos.detalle_servicios set tipocomprobante = '" + dato0.text + "',numerocomprobante='" + dato1.text + "' "
    Else
        csql.sql = "update " + clientesistema + "consumos_basicos.detalle_servicios set tipocomprobante = '',numerocomprobante='' "
    End If
    csql.sql = csql.sql & "where tipodocumento='F' and rut='" + Grid1.Cell(row, 13).text + "' and numerodocumento='" + Grid1.Cell(row, 6).text + "'"
    csql.Execute
    Call sincronizadatos(csql.sql, contadb, "")
    
    csql.Close
    Set csql = Nothing






End Sub

Sub leeboletadehonorarios(row As Long, col As Long)
    PAGARFACTURA = False
    
    campos(0, 0) = "liquido-abono"
    campos(1, 0) = ""
    campos(0, 2) = "boletasdehonorarios"
    condicion = "tipo='" + "1" + "' and rut='" + Grid1.Cell(Grid1.ActiveCell.row, 13).text + "' and numero='" + Grid1.Cell(Grid1.ActiveCell.row, 6).text + "' order by numero desc"
    op = 5
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)
    If sqlconta.status = 0 Then boletano = "S" Else boletano = "N"
    Grid1.Cell(Grid1.ActiveCell.row, 8).text = sqlconta.response(0, 3)
    If sqlconta.response(0, 3) = 0 Then
    PAGARFACTURA = False
    Else
    PAGARFACTURA = True
    
    End If
    
End Sub

Sub abonoboletadehonorarios(row As Long, col As Long, ByVal ELIMINA As Boolean)
    Dim csql As rdoQuery
    Set csql = New rdoQuery
    Set csql.ActiveConnection = contadb
    If ELIMINA = True Then
        csql.sql = "update boletasdehonorarios set abono = abono + " & -1 * CDbl(Grid1.Cell(row, col).text) & " "
    Else
        csql.sql = "update boletasdehonorarios set abono=abono+" & CDbl(Grid1.Cell(row, col).text) & " "
    End If
    csql.sql = csql.sql & "where tipo='" + "1" + "' and rut='" + Grid1.Cell(row, 13).text + "' and numero='" + Grid1.Cell(row, 6).text + "'"
    csql.Execute
    Call sincronizadatos(csql.sql, contadb, "")
    
    csql.Close
    Set csql = Nothing
End Sub

Sub ACTUALIZADOCUMENTO(COMANDO As String)
    Dim lin As Integer
    Dim resultados As rdoResultset
    Dim csql As New rdoQuery
    Dim rut As String
    Dim TIPOFA As String
        Set csql.ActiveConnection = contadb
        csql.sql = "SELECT tipo, numero, linea, fecha, codigocuenta, tipoctacte, rutctacte, centrocosto, glosacontable, tipodocumento, numerodocumento, fechavencimiento, monto, dh "
        csql.sql = csql.sql + "FROM movimientoscontables "
        csql.sql = csql.sql + "WHERE tipo = '" + dato0.text + "' and numero = '" & dato1.text & "' and año = '" + dato4.text + "' and mes = '" + dato3.text + "' order by linea"
        csql.Execute
        If csql.RowsAffected > 0 Then
            Set resultados = csql.OpenResultset
            While Not resultados.EOF
                Call actualizamayor(COMANDO, resultados(4), resultados(12), resultados(13), resultados(5), resultados(6), resultados(7), dato3.text, dato4.text)
                resultados.MoveNext
            Wend
            resultados.Close
            Set resultados = Nothing
        End If
End Sub
Sub abonoGUIADEVOLUCION(tipo, numero, TIPOCO, NUMEROCO, fechaco, montoco)
    Dim csql As rdoQuery
    Set csql = New rdoQuery
    Set csql.ActiveConnection = contadb
    csql.sql = "update devoluciones_proveedores set tipoco='" & TIPOCO & "',numeroco='" & NUMEROCO & "',fechaco='" & Format(fechaco, "yyyy-mm-dd") & "',montoco='" & montoco & "' "
    csql.sql = csql.sql & "where tipo='" + tipo + "' and numero='" + numero + "'"
    csql.Execute
    Call sincronizadatos(csql.sql, contadb, "")
    
    csql.Close
    Set csql = Nothing
End Sub



Sub abonopublicidad(tipo, numero, montoco)
    Dim csql As rdoQuery
    Set csql = New rdoQuery
    Set csql.ActiveConnection = contadb
    
    
    csql.sql = "update facturasdepublicidad set abono=abono-'" & montoco & "' "
    csql.sql = csql.sql & "where tipo='" + tipo + "' and numero='" + numero + "'"
    csql.Execute
    Call sincronizadatos(csql.sql, contadb, "")
    
    csql.Close
    Set csql = Nothing
End Sub

Private Sub botonmisaccesos_Click()
programafiltro = Me.Caption
misaccesos.Show
End Sub


Private Sub botonmisfavoritos_Click()
Call AgregaFavoritos(USUARIOSISTEMA, App.EXEName, Me.Name, Me.Caption)

End Sub

Private Sub ImportaCSV()
Dim RUTA As String
If IsDate(dato2.text & "-" & dato3.text & "-" & dato4.text) = False Then
    dato2.text = Format(fechasistema, "dd")
    dato3.text = Format(fechasistema, "mm")
    dato4.text = Format(fechasistema, "yyyy")
End If

RUTA = Environ("UserProfile") & "\escritorio"
If ExisteCarpeta(RUTA) = False Then
    RUTA = Environ("UserProfile") & "\desktop"
End If
If ExisteCarpeta(RUTA) = False Then
    RUTA = "c:"
End If

 ARCHIVO.InitDir = RUTA
    ARCHIVO.Filter = "Archivos CSV|*.csv"
    ARCHIVO.ShowOpen
    
    
    If ARCHIVO.FileName <> "" And ExisteArchivo(ARCHIVO.FileName) = True Then Call TRASPASADATOS

End Sub

Private Sub TRASPASADATOS()
Dim cad As String
Dim leerachivo As Boolean
Dim impresion As Grid
Dim cadena As Variant
Dim xls_cuenta As String
Dim xls_glosa As String
Dim xls_tipo As String
Dim xls_nume As String
Dim xls_venc As String
Dim xls_monto As Double
Dim xls_dh As String
Dim xls_rutcta As String
Dim dato4 As String
'LEERARCHIVO = False
Dim n As Long
Dim k As Long
Dim leer As String
leer = ARCHIVO.FileName
Open leer For Input As #1
n = 0
k = 0
On Error GoTo error
Grid1.Rows = 1
Grid1.AutoRedraw = False

While Not EOF(1)
Line Input #1, cad

If delimitador(0).Value = True Then
    cadena = Split(cad, ";")
Else
    cadena = Split(cad, ",")
End If

n = n + 1
If Val(cadena(0)) = 0 Then GoTo sig
    xls_cuenta = cadena(0)
    xls_glosa = cadena(1)
    xls_tipo = cadena(2)
    xls_nume = Format(cadena(3), "0000000000")
    xls_venc = cadena(4)
    xls_monto = cadena(5)
    xls_dh = cadena(6)
    xls_rutcta = Format(cadena(7), "0000000000")
If invertir.Value = 1 Then
      If xls_dh = "D" Then
        xls_dh = "H"
      Else
        xls_dh = "D"
      End If
End If

If cadena(0) <> "" And cadena(6) > 0 Then
    Grid1.AddItem "", True
    Grid1.Cell(Grid1.Rows - 1, 1).text = Mid(xls_cuenta, 1, 2)
    Grid1.Cell(Grid1.Rows - 1, 2).text = Mid(xls_cuenta, 4, 2)
    Grid1.Cell(Grid1.Rows - 1, 3).text = Mid(xls_cuenta, 7, 4)
    
    Grid1.Cell(Grid1.Rows - 1, 4).text = xls_glosa
    Grid1.Cell(Grid1.Rows - 1, 5).text = xls_tipo
    Grid1.Cell(Grid1.Rows - 1, 6).text = xls_nume
    Grid1.Cell(Grid1.Rows - 1, 7).text = xls_venc
    Grid1.Cell(Grid1.Rows - 1, 8).text = xls_monto
    Grid1.Cell(Grid1.Rows - 1, 9).text = xls_dh
    Grid1.Cell(Grid1.Rows - 1, 10).text = leerNombreMayor(Replace(xls_cuenta, ".", ""))
    Grid1.Cell(Grid1.Rows - 1, 13).text = xls_rutcta
    Grid1.Cell(Grid1.Rows - 1, 11).text = leerNombrerut(Replace(xls_cuenta, ".", ""), xls_rutcta)
    
End If
sig:
 
Wend
Close #1
Grid1.AutoRedraw = True
Grid1.Refresh
 Screen.MousePointer = vbHourglass

    Grid1.AddItem "", True
    CABEZA.Enabled = True
    Grid1.Enabled = True

Grid1.Cell(Grid1.Rows - 1, 1).EnsureVisible
Grid1.Cell(Grid1.Rows - 1, 1).SetFocus
Screen.MousePointer = vbNormal
Call SUMAR
 
Exit Sub
error:
Close #1
End Sub


    Function ExisteCarpeta(RUTA As String) As Boolean
        Dim carpeta As String
        On Error Resume Next
         carpeta = Dir(RUTA, vbDirectory)
        If carpeta = "" Then
            
            ExisteCarpeta = False
        Else
            ExisteCarpeta = True
        End If
    End Function


Sub ChequeEstandar2(documento As Integer, Optional ByVal suelto As Boolean)
    Dim j As Integer
    Dim n As Integer
    Dim a As Integer
    Dim TITU As String
    Dim montocheque As Double
    Dim numerocheque As String
    Dim fechacheque As Date
    Dim CHEQUEGIRADO As String
    Dim IMPRIMECHEQUE As Boolean
    Dim largo As Integer
    Dim monto As String
    Dim calzador As Integer

    Rem DATOS DE LA COLUMNA
    FORMATOGRILLA(1, 1) = "CUENTA"
    FORMATOGRILLA(1, 2) = "L"
    FORMATOGRILLA(1, 3) = "CODIGO"
    FORMATOGRILLA(1, 4) = "TP"
    FORMATOGRILLA(1, 5) = "NUMERO"
    FORMATOGRILLA(1, 6) = "F.VENCI."
    FORMATOGRILLA(1, 7) = "GLOSA"
    FORMATOGRILLA(1, 8) = "DEBE"
    FORMATOGRILLA(1, 9) = "HABER"
    FORMATOGRILLA(1, 10) = "ANALISIS"
    
    Rem LARGO DE LOS DATOS
    
    FORMATOGRILLA(2, 1) = "19"
    FORMATOGRILLA(2, 2) = "2"
    FORMATOGRILLA(2, 3) = "8"
    FORMATOGRILLA(2, 4) = "3"
    FORMATOGRILLA(2, 5) = "8"
    FORMATOGRILLA(2, 6) = "8"
    FORMATOGRILLA(2, 7) = "30"
    FORMATOGRILLA(2, 8) = "9"
    FORMATOGRILLA(2, 9) = "9"
    FORMATOGRILLA(2, 10) = "9"
    
    Rem TIPO DE DATOS C=CEROS,S=STRING,N=NUMERICO,D=DATE,
    FORMATOGRILLA(3, 1) = "S"
    FORMATOGRILLA(3, 2) = "N"
    FORMATOGRILLA(3, 3) = "S"
    FORMATOGRILLA(3, 4) = "S"
    FORMATOGRILLA(3, 5) = "C"
    FORMATOGRILLA(3, 6) = "D"
    FORMATOGRILLA(3, 7) = "S"
    FORMATOGRILLA(3, 8) = "N"
    FORMATOGRILLA(3, 9) = "N"
    FORMATOGRILLA(3, 10) = "S"
    
    Rem FORMATO GRILLA
    FORMATOGRILLA(4, 1) = ""
    FORMATOGRILLA(4, 2) = ""
    FORMATOGRILLA(4, 3) = ""
    FORMATOGRILLA(4, 4) = ""
    FORMATOGRILLA(4, 5) = ""
    FORMATOGRILLA(4, 6) = ""
    FORMATOGRILLA(4, 7) = ""
    FORMATOGRILLA(4, 8) = ""
    FORMATOGRILLA(4, 9) = ""
    FORMATOGRILLA(4, 10) = ""
    
    Rem LOCCKED
    FORMATOGRILLA(5, 1) = "FALSE"
    FORMATOGRILLA(5, 2) = "FALSE"
    FORMATOGRILLA(5, 3) = "FALSE"
    FORMATOGRILLA(5, 4) = "FALSE"
    FORMATOGRILLA(5, 5) = "FALSE"
    FORMATOGRILLA(5, 6) = "FALSE"
    FORMATOGRILLA(5, 7) = "FALSE"
    FORMATOGRILLA(5, 8) = "FALSE"
    FORMATOGRILLA(5, 9) = "FALSE"
    FORMATOGRILLA(5, 10) = "FALSE"
    
    Rem VALOR MINIMO
    FORMATOGRILLA(6, 1) = ""
    FORMATOGRILLA(6, 2) = ""
    FORMATOGRILLA(6, 3) = ""
    FORMATOGRILLA(6, 4) = ""
    FORMATOGRILLA(6, 5) = ""
    FORMATOGRILLA(6, 6) = ""
    FORMATOGRILLA(6, 7) = ""
    FORMATOGRILLA(6, 8) = ""
    FORMATOGRILLA(6, 9) = ""
    
    Rem VALOR MAXIMO
    FORMATOGRILLA(7, 1) = ""
    FORMATOGRILLA(7, 2) = ""
    FORMATOGRILLA(7, 3) = ""
    FORMATOGRILLA(7, 4) = ""
    FORMATOGRILLA(7, 5) = ""
    FORMATOGRILLA(7, 6) = ""
    FORMATOGRILLA(7, 7) = ""
    FORMATOGRILLA(7, 8) = ""
    FORMATOGRILLA(7, 9) = ""
    Grid2.Cols = 11
    Grid2.Rows = Grid1.Rows + 1
    Grid2.Range(1, 1, Grid2.Rows - 1, Grid2.Cols - 1).ClearText
    Grid2.Column(10).Alignment = cellRightGeneral
    
    
    For k = 1 To Grid2.Cols - 1
        Grid2.Cell(0, k).text = FORMATOGRILLA(1, k)
        Grid2.Column(k).MaxLength = Val(FORMATOGRILLA(2, k))
        Grid2.Column(k).FormatString = FORMATOGRILLA(4, k)
        Rem Grid2.Column(K).Locked = FORMATOGRILLA(5, K)
        If FORMATOGRILLA(3, k) = "N" Then Grid2.Column(k).Alignment = cellRightCenter
        If FORMATOGRILLA(3, k) = "D" Then Grid2.Column(k).CellType = cellCalendar
        'Grid1.Column(7).CellType = cellComboBox
    Next k
    
    Grid2.DefaultFont.Size = 6.5
    For k = 1 To Grid2.Cols - 1
        Grid2.Column(k).Width = Val(FORMATOGRILLA(2, k)) * (Grid2.DefaultFont.Size)
    Next k

    Grid2.PageSetup.Orientation = cellPortrait
    Grid2.PageSetup.PrintFixedRow = True
    Grid2.PageSetup.PrintGridlines = False

    Grid2.PageSetup.BlackAndWhite = True
    Grid2.PageSetup.BottomMargin = 1
    Grid2.PageSetup.TopMargin = 1
    Grid2.PageSetup.LeftMargin = 1
    Grid2.PageSetup.RightMargin = 0.5
    Grid2.Range(0, 1, 0, Grid2.Cols - 1).FontBold = True
    Grid2.Range(0, 1, 0, Grid2.Cols - 1).Borders(cellEdgeTop) = cellThin
    Grid2.Range(0, 1, 0, Grid2.Cols - 1).Borders(cellEdgeBottom) = cellThin
    
    If dato0.text = "CI" Then TITU = "COMPROBANTE DE INGRESO"
    If dato0.text = "CE" Then TITU = "COMPROBANTE DE EGRESO"
    If dato0.text = "CT" Then TITU = "COMPROBANTE DE TRASPASO"

  If suelto = False Then Call cabezERA(TITU + " NUMERO " + dato1.text, dato1.text, fechacon)
    IMPRIMECHEQUE = False

    For j = 1 To Grid1.Rows - 1
        Grid2.Cell(j, 1).text = Grid1.Cell(j, 10).text
        Grid2.Cell(j, 2).text = j
        Grid2.Cell(j, 3).text = Grid1.Cell(j, 1).text + "." + Grid1.Cell(j, 2).text + "." + Grid1.Cell(j, 3).text
        Grid2.Cell(j, 4).text = Grid1.Cell(j, 5).text
        Grid2.Cell(j, 5).text = Grid1.Cell(j, 6).text
        Grid2.Cell(j, 6).text = Grid1.Cell(j, 7).text
        Grid2.Cell(j, 7).text = Grid1.Cell(j, 4).text
        If Grid1.Cell(j, 9).text = "D" Then Grid2.Cell(j, 8).text = Grid1.Cell(j, 8).text
        If Grid1.Cell(j, 9).text = "H" Then Grid2.Cell(j, 9).text = Grid1.Cell(j, 8).text
        If Grid1.Cell(j, 13).text <> "" Then Grid2.Cell(j, 10).text = Mid(Grid1.Cell(j, 13).text, 1, 9) + "-" + Mid(Grid1.Cell(j, 13).text, 10, 1)
        If Grid1.Cell(j, 14).text <> "" Then Grid2.Cell(j, 10).text = Grid1.Cell(j, 14).text
        
        If Grid1.Cell(j, 5).text = "CH" Then
            If Grid2.Cell(j, 9).text <> "" Then
            montocheque = Grid2.Cell(j, 9).text
            Else
            montocheque = Grid2.Cell(j, 8).text
            End If
            fechacheque = Grid2.Cell(j, 6).text
            numerocheque = Grid2.Cell(j, 5).text
            CHEQUEGIRADO = Grid2.Cell(j, 7).text
            IMPRIMECHEQUE = True
        End If
    Next j
    
    For a = j To 12
        Grid2.Rows = Grid2.Rows + 1
    Next a
    
    Grid2.Range(Grid2.Rows - 1, 1, Grid2.Rows - 1, Grid2.Cols - 1).Borders(cellEdgeTop) = cellThin
    Grid2.Cell(Grid2.Rows - 1, 7).text = "TOTAL COMPROBANTE CONTABLE"
    Grid2.Cell(Grid2.Rows - 1, 8).text = debe.Caption
    Grid2.Cell(Grid2.Rows - 1, 9).text = haber.Caption

    Grid2.AddItem ""
    Grid2.AddItem ""
    Grid2.Range(Grid2.Rows - 2, 1, Grid2.Rows - 1, Grid2.Cols - 1).Merge
    Grid2.Range(Grid2.Rows - 2, 1, Grid2.Rows - 1, Grid2.Cols - 1).WrapText = True
'    Grid2.Cell(Grid2.rows - 1, 1).text = "GLOSA :" + glosa.text
    Grid2.Cell(Grid2.Rows - 2, 1).text = "GLOSA :  " & glosa.text

    Grid2.AddItem ""
    Grid2.AddItem ""
    Grid2.AddItem ""
    Grid2.AddItem ""
    Grid2.AddItem ""
    Grid2.Range(Grid2.Rows - 1, 1, Grid2.Rows - 1, Grid2.Cols - 1).Merge
    Grid2.Cell(Grid2.Rows - 1, 1).text = "              _______________________                       ______________________                           _______________________             ____________________"
    Grid2.AddItem ""

    Grid2.Range(Grid2.Rows - 1, 1, Grid2.Rows - 1, Grid2.Cols - 1).Merge
    Grid2.Cell(Grid2.Rows - 1, 1).text = "                  VºBº Gerencia                                            VºBº Contabilidad                                      Recibo Conforme                           Confec. " + USUARIOSISTEMA
    For n = 0 To Grid2.Rows - 1
            Grid2.RowHeight(n) = 0
        Next n
        
    Rem CHEQUE SANTANDER
    If documento = 1 Then
        Grid2.PageSetup.LeftMargin = 1
        Grid1.RowHeight(Grid1.Rows - 1) = 10
        Grid2.AddItem "", True
         Grid2.AddItem "", True
          Grid2.AddItem "", True
          
          
'        For k = 1 To 18
'            Grid2.AddItem ""
'        Next k
        Grid2.Column(1).Width = 95
        Grid2.Column(2).Width = 240
        Grid2.Column(3).Width = 40
        Grid2.Column(4).Width = 43
        Grid2.Column(5).Width = 50
        Grid2.Column(6).Width = 40
        Grid2.Column(7).Width = 40
        Grid2.Column(8).Width = 40

       
       
        Grid2.Range(Grid2.Rows - 2, 4, Grid2.Rows - 1, 8).Merge
        Grid2.Range(Grid2.Rows - 2, 4, Grid2.Rows - 1, 8).CellType = cellTextBox
        Grid2.Range(Grid2.Rows - 2, 4, Grid2.Rows - 1, 8).FontBold = True
        Grid2.Range(Grid2.Rows - 2, 4, Grid2.Rows - 1, 8).FontSize = 12
        Grid2.Range(Grid2.Rows - 2, 4, Grid2.Rows - 1, 8).Alignment = cellRightTop
        Grid2.RowHeight(Grid2.Rows - 3) = 20
        
        Grid2.Cell(Grid2.Rows - 2, 4).text = Format(montocheque, "#  #  #  #  #  #  #  #  #  #  #  #  0") 'Format(montocheque, "##,###,###,###")
        Grid2.Cell(Grid2.Rows - 2, 4).Alignment = cellRightTop
        Grid2.AddItem ""
       
        
        Grid2.Range(Grid2.Rows - 1, 5, Grid2.Rows - 1, 8).FontBold = True
        Grid2.Range(Grid2.Rows - 1, 5, Grid2.Rows - 1, 8).FontSize = 10
        
        Grid2.Cell(Grid2.Rows - 1, 5).text = " " & Format(Format(fechacheque, "dd"), "0  #")
        Grid2.Cell(Grid2.Rows - 1, 6).text = " " & Format(Format(fechacheque, "mm"), "0  #")
        Grid2.Cell(Grid2.Rows - 1, 8).text = Format(Mid(Format(fechacheque, "yyyy"), 3, 2), "0  #")
        
        Grid2.Range(Grid2.Rows - 1, 3, Grid2.Rows - 1, 4).Merge
        Grid2.Range(Grid2.Rows - 1, 3, Grid2.Rows - 1, 4).FontSize = 10
        Grid2.Range(Grid2.Rows - 1, 3, Grid2.Rows - 1, 4).FontBold = True
        
        Grid2.Cell(Grid2.Rows - 1, 3).text = "PUCON"
        Grid2.AddItem ""
        Grid2.AddItem ""
        
        Grid2.Range(Grid2.Rows - 1, 2, Grid2.Rows - 1, 8).FontBold = True
        Grid2.Range(Grid2.Rows - 1, 2, Grid2.Rows - 1, 8).FontSize = 9
        Grid2.Range(Grid2.Rows - 1, 2, Grid2.Rows - 1, 8).Merge
        Grid2.Range(Grid2.Rows - 1, 2, Grid2.Rows - 1, 8).Alignment = cellLeftCenter
        
        
        Grid2.Cell(Grid2.Rows - 1, 2).text = CHEQUEGIRADO

        Grid2.AddItem ""
      
        Grid2.Range(Grid2.Rows - 1, 2, Grid2.Rows - 1, 8).FontBold = True
        Grid2.Range(Grid2.Rows - 1, 2, Grid2.Rows - 1, 8).FontSize = 9
        Grid2.Range(Grid2.Rows - 1, 2, Grid2.Rows - 1, 8).Merge
        Grid2.Range(Grid2.Rows - 1, 2, Grid2.Rows - 1, 8).Alignment = cellLeftCenter
        
        monto = WORDNUM(montocheque, "PESO", "", "CENTAVO", "CENTAVOS", 0)
        monto = UCase(monto)
        
        largo = Len(monto)
        calzador = largo
        If largo > 60 Then
            For k = 60 To 1 Step -1
                If Asc(Mid(monto, k, 1)) <> 32 Then calzador = 60 - (60 - k): Exit For
            Next k
            Grid2.Cell(Grid2.Rows - 1, 2).text = Mid(monto, 1, calzador) + "-"
            Grid2.AddItem ""
            Grid2.Range(Grid2.Rows - 1, 2, Grid2.Rows - 1, 8).FontBold = True
            Grid2.Range(Grid2.Rows - 1, 2, Grid2.Rows - 1, 8).FontSize = 9
            Grid2.Range(Grid2.Rows - 1, 2, Grid2.Rows - 1, 8).Merge
            Grid2.Range(Grid2.Rows - 1, 2, Grid2.Rows - 1, 8).Alignment = cellLeftCenter
            
            Grid2.Cell(Grid2.Rows - 1, 2).text = Mid(monto, calzador + 1, largo)
        Else
            Grid2.Cell(Grid2.Rows - 1, 2).text = monto
            Grid2.Range(Grid2.Rows - 1, 2, Grid2.Rows - 1, 8).Merge
        End If
        Grid2.Range(Grid2.Rows - 1, 2, Grid2.Rows - 1, 8).FontBold = True
        Grid2.Range(Grid2.Rows - 1, 2, Grid2.Rows - 1, 8).FontSize = 9
        Grid2.Range(Grid2.Rows - 1, 2, Grid2.Rows - 1, 8).Merge
        Grid2.Range(Grid2.Rows - 1, 2, Grid2.Rows - 1, 8).Alignment = cellLeftCenter
    
        Grid2.Cell(Grid2.Rows - 1, 2).text = Grid2.Cell(Grid2.Rows - 1, 2).text & ".__________________________"
           
    
            '----
    Grid2.AddItem "", True
    Grid2.AddItem "", True
        
    Grid2.Cell(Grid2.Rows - 1, 2).text = empresaactiva & " " & "CH:" & numerocheque
    Grid2.Cell(Grid2.Rows - 1, 2).Alignment = cellRightCenter
    
    Grid2.AddItem "", True
    '---
    
    
    End If
    If suelto = True Then
        Grid2.PageSetup.LeftMargin = 0
        
        Grid2.PageSetup.TopMargin = 0.2
        Grid2.PageSetup.RightMargin = 0
         Grid2.PageSetup.PrintGridlines = False
'          Grid2.PageSetup.PrintGridlines = True
       
        
        Grid2.Column(9).Width = 0
        Grid2.Column(10).Width = 0
        
    Else
        Grid2.PageSetup.LeftMargin = 1.481667
        Grid2.Column(10).Width = 100
        Grid2.PageSetup.RightMargin = 0
                Grid2.PageSetup.PrintGridlines = False
    End If

        Grid2.PrintPreview ' 100
        frmimprimir.Visible = False
          Call retorno
    
End Sub

