VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash9b.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{4777436C-EB8C-4596-98A8-EBCF98683969}#1.0#0"; "FlexCell.ocx"
Object = "{ADBBAED6-B16F-40DC-80DF-B44910CBA76C}#1.0#0"; "Frame-Xp.ocx"
Begin VB.Form ingreso06 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Ingreso Facturas de Compras Propias"
   ClientHeight    =   8250
   ClientLeft      =   45
   ClientTop       =   1110
   ClientWidth     =   15045
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8250
   ScaleWidth      =   15045
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox PIVOTE4 
      Height          =   285
      Left            =   6240
      MaxLength       =   9
      TabIndex        =   60
      Text            =   "Text1"
      Top             =   6240
      Visible         =   0   'False
      Width           =   615
   End
   Begin XPFrame.FrameXp impuestos 
      Height          =   375
      Left            =   14280
      TabIndex        =   8
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
         TabIndex        =   9
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
   Begin XPFrame.FrameXp DETALLE 
      Height          =   5055
      Left            =   6000
      TabIndex        =   35
      Top             =   120
      Width           =   9015
      _ExtentX        =   15901
      _ExtentY        =   8916
      BackColor       =   16744576
      Caption         =   "Detalle de Gastos"
      CaptionEstilo3D =   1
      BackColor       =   16744576
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
      Begin FlexCell.Grid Grid1 
         Height          =   3375
         Left            =   0
         TabIndex        =   50
         Top             =   240
         Width           =   9015
         _ExtentX        =   15901
         _ExtentY        =   5953
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
         DefaultFontSize =   8.25
         DefaultFontBold =   -1  'True
         ForeColorFixed  =   8388608
         GridColor       =   16744576
         Rows            =   30
         DateFormat      =   2
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "NOMBRE CUENTA CONTABLE"
         Height          =   255
         Left            =   120
         TabIndex        =   41
         Top             =   3720
         Width           =   4575
      End
      Begin VB.Label nombremayor 
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
         Height          =   255
         Left            =   120
         TabIndex        =   40
         Top             =   3960
         Width           =   4575
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "NOMBRE CENTRO DE COSTO"
         Height          =   255
         Left            =   4800
         TabIndex        =   39
         Top             =   4320
         Width           =   4215
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
         Height          =   255
         Left            =   4800
         TabIndex        =   38
         Top             =   4560
         Width           =   4215
      End
      Begin VB.Label nombrectacte 
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
         Height          =   255
         Left            =   120
         TabIndex        =   37
         Top             =   4560
         Width           =   4575
      End
      Begin VB.Label Label14 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "NOMBRE CUENTA CORRIENTE"
         Height          =   255
         Left            =   120
         TabIndex        =   36
         Top             =   4320
         Width           =   4575
      End
   End
   Begin XPFrame.FrameXp FrameXp1 
      Height          =   7575
      Left            =   120
      TabIndex        =   10
      Top             =   120
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   13361
      BackColor       =   16744576
      Caption         =   "DATOS DOCUMENTO"
      CaptionEstilo3D =   1
      BackColor       =   16744576
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
      Begin VB.TextBox DATO15 
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
         Left            =   1800
         MaxLength       =   1
         TabIndex        =   62
         Tag             =   "tipo"
         Text            =   "N"
         ToolTipText     =   "(S)i o (N)o"
         Top             =   4920
         Width           =   255
      End
      Begin VB.TextBox DATO16 
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
         Left            =   1800
         MaxLength       =   1
         TabIndex        =   51
         Tag             =   "tipo"
         Text            =   "N"
         ToolTipText     =   "(S)i o (N)o"
         Top             =   5280
         Width           =   255
      End
      Begin XPFrame.FrameXp FrameXp2 
         Height          =   1815
         Left            =   120
         TabIndex        =   48
         Top             =   5640
         Width           =   5535
         _ExtentX        =   9763
         _ExtentY        =   3201
         BackColor       =   16761024
         Caption         =   "GLOSA DOCUMENTO"
         CaptionEstilo3D =   1
         BackColor       =   16761024
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
         Begin VB.TextBox DATO17 
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
            ForeColor       =   &H80000002&
            Height          =   975
            Left            =   120
            MaxLength       =   255
            MultiLine       =   -1  'True
            TabIndex        =   49
            ToolTipText     =   "ingreso detalle informativo de la Factura"
            Top             =   360
            Width           =   5295
         End
      End
      Begin XPFrame.FrameXp TIPOS 
         Height          =   1215
         Left            =   2280
         TabIndex        =   42
         Top             =   960
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   2143
         BackColor       =   16761024
         Caption         =   "TIPOS DE DOCUMENTOS"
         CaptionEstilo3D =   1
         BackColor       =   16761024
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
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid GRILLATIPO 
            Height          =   855
            Left            =   120
            TabIndex        =   43
            Top             =   240
            Width           =   3015
            _ExtentX        =   5318
            _ExtentY        =   1508
            _Version        =   393216
            BackColor       =   16107953
            ForeColor       =   16711680
            Rows            =   3
            FixedRows       =   0
            FixedCols       =   0
            ForeColorFixed  =   16777152
            BackColorBkg    =   16761024
            GridColor       =   16744576
            GridColorFixed  =   14282751
            GridColorUnpopulated=   14282751
            _NumberOfBands  =   1
            _Band(0).Cols   =   2
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
         Left            =   2520
         MaxLength       =   4
         TabIndex        =   21
         Tag             =   "fecha"
         Top             =   2400
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
         Left            =   2160
         MaxLength       =   2
         TabIndex        =   20
         Tag             =   "fecha"
         Top             =   2400
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
         Left            =   1800
         MaxLength       =   2
         TabIndex        =   19
         Tag             =   "fecha"
         Top             =   2400
         Width           =   375
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
         Left            =   1800
         MaxLength       =   9
         TabIndex        =   2
         Tag             =   "rut"
         Top             =   1680
         Width           =   1095
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
         Left            =   1800
         MaxLength       =   15
         TabIndex        =   18
         Tag             =   "neto"
         Text            =   "0"
         Top             =   3120
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
         Left            =   1800
         MaxLength       =   15
         TabIndex        =   17
         Tag             =   "iva"
         Text            =   "0"
         Top             =   3480
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
         Left            =   1800
         MaxLength       =   15
         TabIndex        =   16
         Tag             =   "exento"
         Text            =   "0"
         Top             =   3840
         Width           =   1455
      End
      Begin VB.TextBox dato14 
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
         Left            =   1800
         MaxLength       =   15
         TabIndex        =   15
         Tag             =   "retencion"
         Text            =   "0"
         Top             =   4200
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
         Left            =   1800
         MaxLength       =   15
         TabIndex        =   14
         Tag             =   "monto"
         Text            =   "0"
         Top             =   4560
         Width           =   1455
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
         Left            =   1800
         MaxLength       =   2
         TabIndex        =   13
         Tag             =   "fechavencimiento"
         Top             =   2760
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
         Left            =   2160
         MaxLength       =   2
         TabIndex        =   12
         Tag             =   "fecha"
         Top             =   2760
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
         Left            =   2520
         MaxLength       =   4
         TabIndex        =   11
         Tag             =   "fecha"
         Top             =   2760
         Width           =   615
      End
      Begin VB.TextBox dato2 
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
         Left            =   1800
         MaxLength       =   10
         TabIndex        =   1
         Tag             =   "numero"
         Top             =   1320
         Width           =   1455
      End
      Begin VB.TextBox dato1 
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
         Left            =   1800
         MaxLength       =   1
         TabIndex        =   0
         Tag             =   "tipo"
         Top             =   960
         Width           =   255
      End
      Begin XPFrame.FrameXp FrameXp6 
         Height          =   615
         Left            =   120
         TabIndex        =   44
         Top             =   240
         Width           =   5415
         _ExtentX        =   9551
         _ExtentY        =   1085
         BackColor       =   16744576
         Caption         =   "MES                                                       AÑO                        FOLIOS"
         CaptionEstilo3D =   1
         BackColor       =   16744576
         ForeColor       =   65535
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Begin VB.TextBox folios 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000F&
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
            Left            =   3960
            Locked          =   -1  'True
            MaxLength       =   8
            TabIndex        =   47
            Tag             =   "tipo"
            Top             =   240
            Width           =   1335
         End
         Begin VB.ComboBox COMBOAÑO 
            Height          =   315
            Left            =   2640
            TabIndex        =   46
            Top             =   240
            Width           =   1215
         End
         Begin VB.ComboBox COMBOMES 
            Height          =   315
            Left            =   0
            TabIndex        =   45
            Top             =   240
            Width           =   2655
         End
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   " ACTIVO"
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
         TabIndex        =   63
         Top             =   4920
         Width           =   1575
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
         Left            =   2280
         TabIndex        =   61
         Top             =   960
         Width           =   3135
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "ELECTRONICA"
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
         TabIndex        =   52
         Top             =   5280
         Width           =   1575
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
         Left            =   120
         TabIndex        =   34
         Top             =   2040
         Width           =   5175
      End
      Begin VB.Label Label16 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   " PROVEEDOR"
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
         TabIndex        =   33
         Top             =   1680
         Width           =   1455
      End
      Begin VB.Label Label11 
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
         Left            =   120
         TabIndex        =   32
         Top             =   3120
         Width           =   1575
      End
      Begin VB.Label Label17 
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
         Left            =   120
         TabIndex        =   31
         Top             =   3480
         Width           =   1575
      End
      Begin VB.Label Label26 
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
         Left            =   120
         TabIndex        =   30
         Top             =   3840
         Width           =   1575
      End
      Begin VB.Label Label27 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   " RETENCION"
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
         TabIndex        =   29
         Top             =   4200
         Width           =   1575
      End
      Begin VB.Label Label28 
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
         Left            =   120
         TabIndex        =   28
         Top             =   4560
         Width           =   1575
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
         TabIndex        =   27
         Top             =   1320
         Width           =   1455
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
         TabIndex        =   26
         Top             =   960
         Width           =   1455
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   " VENCIMIENTO"
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
         TabIndex        =   25
         Top             =   2760
         Width           =   1575
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   " FECHA EMISION"
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
         TabIndex        =   24
         Top             =   2400
         Width           =   1575
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
         Left            =   3000
         TabIndex        =   23
         Top             =   1680
         Width           =   255
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "DETALLE DE FACTURA DE COMPRR (FIN ACCESO DETALLE)"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   6120
         Width           =   5175
      End
   End
   Begin VB.PictureBox MANUAL 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   135
      Left            =   0
      ScaleHeight     =   135
      ScaleWidth      =   615
      TabIndex        =   7
      Top             =   0
      Width           =   615
   End
   Begin VB.TextBox LINEAS 
      Height          =   285
      Left            =   5760
      MaxLength       =   3
      TabIndex        =   6
      Top             =   7560
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox pivote2 
      Height          =   285
      Left            =   7440
      MaxLength       =   10
      TabIndex        =   5
      Top             =   7560
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox PIVOTE 
      Height          =   285
      Left            =   6240
      MaxLength       =   8
      TabIndex        =   3
      Top             =   7560
      Visible         =   0   'False
      Width           =   1215
   End
   Begin XPFrame.FrameXp FrameXp3 
      Height          =   1095
      Left            =   7680
      TabIndex        =   53
      Top             =   5280
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   1931
      BackColor       =   16744576
      Caption         =   "VALORES DEL COMPROBANTE"
      CaptionEstilo3D =   1
      BackColor       =   16744576
      ForeColor       =   8438015
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
         TabIndex        =   54
         Top             =   360
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   1085
         BackColor       =   16761024
         Caption         =   "GASTOS"
         CaptionEstilo3D =   1
         BackColor       =   16761024
         ForeColor       =   8438015
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
            TabIndex        =   55
            Top             =   240
            Width           =   1575
         End
      End
      Begin XPFrame.FrameXp FrameXp5 
         Height          =   615
         Left            =   2760
         TabIndex        =   56
         Top             =   360
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   1085
         BackColor       =   16761024
         Caption         =   "TOTAL"
         CaptionEstilo3D =   1
         BackColor       =   16761024
         ForeColor       =   8438015
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
            TabIndex        =   57
            Top             =   240
            Width           =   1695
         End
      End
      Begin XPFrame.FrameXp FrameXp7 
         Height          =   615
         Left            =   4920
         TabIndex        =   58
         Top             =   360
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   1085
         BackColor       =   16761024
         Caption         =   "SALDO"
         CaptionEstilo3D =   1
         BackColor       =   16761024
         ForeColor       =   8438015
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
            TabIndex        =   59
            Top             =   240
            Width           =   1695
         End
      End
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash opciones 
      Height          =   1455
      Left            =   5880
      TabIndex        =   4
      Top             =   6600
      Width           =   8775
      _cx             =   15478
      _cy             =   2566
      FlashVars       =   ""
      Movie           =   "c:\barra_opciones.swf"
      Src             =   "c:\barra_opciones.swf"
      WMode           =   "Transparent"
      Play            =   0   'False
      Loop            =   -1  'True
      Quality         =   "High"
      SAlign          =   ""
      Menu            =   -1  'True
      Base            =   ""
      AllowScriptAccess=   "always"
      Scale           =   "ShowAll"
      DeviceFont      =   0   'False
      EmbedMovie      =   0   'False
      BGColor         =   ""
      SWRemote        =   ""
      MovieData       =   ""
      SeamlessTabbing =   -1  'True
      Profile         =   0   'False
      ProfileAddress  =   ""
      ProfilePort     =   0
      AllowNetworking =   "all"
      AllowFullScreen =   "false"
   End
End
Attribute VB_Name = "ingreso06"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Private tipocuenta As String
    Private cc As Integer
    Private FORMATOGRILLA(100, 20)
    Private formatogrilla2(100, 20)
    Private cdi As Integer
    Private CANDO As Integer
    Private EXISTE As String
    
    Private AUXILIAR(1000, 3) As String
    
    Private respu As String
    Private tipoctacte As String
    Private nlineas As Double
    Private DOCU(6) As String
    Private grilladetalle(1000, 13) As String
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
     Private mes As String
     Private año As String
     
     
    
Private Sub Command2_Click()

End Sub







Private Sub COMBOAÑO_Click()
año = COMBOAÑO.text
mes = COMBOMES.ListIndex + 1
If Val(mes) < 10 Then mes = "0" + Mid(Str(mes), 2, 1) Else mes = Mid(Str(mes), 2, 2)

End Sub


Private Sub COMBOMES_Click()
año = COMBOAÑO.text
mes = COMBOMES.ListIndex + 1
If Val(mes) < 10 Then mes = "0" + Mid(Str(mes), 2, 1) Else mes = Mid(Str(mes), 2, 2)

End Sub

Private Sub dato1_Change()
If Val(dato1.text) < 1 Or Val(dato1.text) > CANDO Then dato1.Enabled = True: dato1.text = "": dato1.SetFocus
End Sub

Private Sub dato1_lostFocus()
TIPOS.Visible = False
leeFOLIO
End Sub
Private Sub DATO1_GotFocus()

Call cargatexto(dato1)
TIPOS.Visible = True
End Sub




Private Sub DATO15_GotFocus()
Call cargatexto(DATO15)

totalfactura
End Sub

Private Sub DATO16_GotFocus()
Call cargatexto(DATO16)

End Sub

Private Sub DATO17_KeyPress(KeyAscii As Integer)
    'If KeyAscii = 13 Then DETALLE.Enabled = True: Grid1.Rows = 2: Grid1.Cell(1, 1).SetFocus

End Sub

Private Sub dato2_GotFocus()

Call cargatexto(dato2)
If Val(dato1.text) < 1 Or Val(dato1.text) > CANDO Then dato1.text = "": dato1.SetFocus:
tipodocumento.Caption = GRILLATIPO.TextMatrix(Val(dato1.text) - 1, 1)

End Sub


Private Sub dato3_Change()
If Val(dato3.text) > 31 Then dato3.text = ""
End Sub
Private Sub dato4_Change()
If Val(dato4.text) > 12 Or Val(dato4.text) < 1 Then dato4.text = ""
End Sub

Private Sub dato5_LostFocus()
If dato5.text < "1900" Or dato5.text > Format(Now, "YYYY") Then dato5.text = ""

End Sub

Private Sub dato8_LostFocus()
If dato8.text < "1900" Or dato8.text > Format(Now, "YYYY") Then dato8.text = ""

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
        If Val(dato9.text) = 0 Then dato9.SetFocus: GoTo no:
    If cierrect = "-" Then cierrect = "": dato9.SetFocus: GoTo no:
End If

Rem If tipocuenta <> "00" And Val(dato9.text) = 0 Then dato9.SetFocus: GoTo no:

leefactura
If SQLUTIL.estado = 0 Then
   carga
   LEERMOVIMIENTOS

      If nlineas <> 0 Then
      opciones.Visible = True
      opciones.SetFocus
      detalle.Enabled = False
           
           
           GoTo no:
      End If
End If

If Val(dato2.text) = 0 Then dato2.text = "": dato2.Enabled = True: dato2.SetFocus
Call cargatexto(dato3)
no:
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

Call cargatexto(dato9)





End Sub

Private Sub dato11_GotFocus()
Rem If IsDate(dato6.text + "-" + dato7.text + "-" + dato8.text) = False Then dato6.text = "": dato7.text = "": dato8.text = "": dato6.SetFocus



Call cargatexto(dato11)

no:
End Sub

Private Sub dato12_GotFocus()
SUMADOR = Int((CDbl(Replace(dato11.text, ",", "")) * iva / 100) + 0.5)
dato12.text = Format(SUMADOR, "#,###,###,##0")
totalfactura
Call cargatexto(dato12)
End Sub
Private Sub dato13_GotFocus()
'Grid2.Cell(1, 1).SetFocus

totalfactura
Call cargatexto(dato13)
End Sub

Private Sub dato14_GotFocus()
totalfactura
Call cargatexto(dato14)

End Sub



Private Sub Form_Load()
CENTRAR Me
iva = 19
    Call Conectar_BD
    Call Conectarconta(servidor, "conta", USUARIO, password)
    sc = 0
    opciones.Visible = False
GRILLATIPOS

Call CARGAGRILLA
Call CARGAGRILLAexento
DOCU(1) = "FA "
DOCU(2) = "ND "
DOCU(3) = "NC "

For K = 1 To 12
COMBOMES.AddItem MonthName(K)
Next K
COMBOMES.ListIndex = CDbl(Format(fechasistema, "mm") - 1)
For K = 2000 To Val(Format(fechasistema, "yyyy"))
COMBOAÑO.AddItem K
Next K
COMBOAÑO.ListIndex = K - 2001

año = COMBOAÑO.text
mes = COMBOMES.ListIndex + 1
If Val(mes) < 10 Then mes = "0" + Mid(Str(mes), 2, 1) Else mes = Mid(Str(mes), 2, 2)
impuestos.Visible = False


End Sub
Sub GRILLATIPOS()
GRILLATIPO.Cols = 2
GRILLATIPO.Rows = 3
GRILLATIPO.ColWidth(0) = 200 * 2
GRILLATIPO.ColWidth(1) = 200 * 10

GRILLATIPO.TextMatrix(0, 0) = "1"
GRILLATIPO.TextMatrix(1, 0) = "2"
GRILLATIPO.TextMatrix(2, 0) = "3"

GRILLATIPO.TextMatrix(0, 1) = "FACTURA"
GRILLATIPO.TextMatrix(1, 1) = "NOTA DE DEBITO"
GRILLATIPO.TextMatrix(2, 1) = "NOTA DE CREDITO"
CANDO = 3



End Sub


Private Sub dato1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 38 Then Unload Me: GoTo no:
    Call flechas(dato1, dato2, KeyCode)
no:
End Sub
Private Sub dato2_KeyDown(KeyCode As Integer, Shift As Integer)
    Call flechas(dato1, dato9, KeyCode)
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
    ' If KeyCode = vbKeyF2 Then Call ayudatipos(dato14)
    Call flechas(dato12, dato14, KeyCode)
End Sub
Private Sub dato14_KeyDown(KeyCode As Integer, Shift As Integer)
    Call flechas(dato13, DATO15, KeyCode)
End Sub
Private Sub dato15_KeyDown(KeyCode As Integer, Shift As Integer)
    Call flechas(dato14, DATO16, KeyCode)
End Sub
Private Sub dato16_KeyDown(KeyCode As Integer, Shift As Integer)
    Call flechas(DATO15, DATO17, KeyCode)
End Sub

Private Sub DATO17_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 35 Then detalle.Enabled = True: Grid1.Rows = 2: Grid1.Cell(1, 1).SetFocus


    Call flechas(DATO16, DATO17, KeyCode)
End Sub

Private Sub DATO1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then Unload Me
    snum = 0: KeyAscii = esNumero(KeyAscii)
    If KeyAscii = 13 Then Call Pregunta(dato1, dato2)
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
    If KeyAscii = 13 And Val(dato11.text) <> 0 Then Call formato(dato11, 0): Call Pregunta(dato11, dato12)
    
End Sub
Private Sub dato12_KeyPress(KeyAscii As Integer)
    snum = 0: KeyAscii = esNumero(KeyAscii)
    If KeyAscii = 13 Then Call formato(dato12, 0): Call Pregunta(dato12, dato13)
    
End Sub
Private Sub dato13_KeyPress(KeyAscii As Integer)
    snum = 0: KeyAscii = esNumero(KeyAscii)
    If KeyAscii = 13 Then Call formato(dato13, 0): Call Pregunta(dato13, dato14)

End Sub
Private Sub dato14_KeyPress(KeyAscii As Integer)
    snum = 0: KeyAscii = esNumero(KeyAscii)
    If KeyAscii = 13 Then Call formato(dato14, 0): Call Pregunta(dato14, DATO15)

End Sub
Private Sub dato15_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    
    If KeyAscii = 13 Then Call Pregunta(DATO15, DATO16)
 If Chr(KeyAscii) <> "S" And Chr(KeyAscii) <> "N" Then KeyAscii = 0
End Sub
Private Sub dato16_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    
    If KeyAscii = 13 Then Call Pregunta(DATO16, DATO17)
 If Chr(KeyAscii) <> "S" And Chr(KeyAscii) <> "N" Then KeyAscii = 0

End Sub


Sub carga()
    disponible (True)
    
    dato1.text = SQLUTIL.datos(0, 3)
    dato2.text = SQLUTIL.datos(1, 3)
    dato3.text = Mid(SQLUTIL.datos(2, 3), 1, 2)
    dato4.text = Mid(SQLUTIL.datos(2, 3), 4, 2)
    dato5.text = Mid(SQLUTIL.datos(2, 3), 7, 4)
    dato6.text = Mid(SQLUTIL.datos(3, 3), 1, 2)
    dato7.text = Mid(SQLUTIL.datos(3, 3), 4, 2)
    dato8.text = Mid(SQLUTIL.datos(3, 3), 7, 4)
    dato9.text = Mid(SQLUTIL.datos(4, 3), 1, 9)
    DV.Caption = Mid(SQLUTIL.datos(4, 3), 10, 1)
    dato11.text = Format(SQLUTIL.datos(5, 3), "##,###,###,##0")
    dato12.text = Format(SQLUTIL.datos(6, 3), "##,###,###,##0")
    dato13.text = Format(SQLUTIL.datos(7, 3), "##,###,###,##0")
    dato14.text = Format(SQLUTIL.datos(8, 3), "##,###,###,##0")
    
    total.text = Format(SQLUTIL.datos(9, 3), "##,###,###,##0")
    mescontabilizado = SQLUTIL.datos(10, 3)
    añocontabilizado = SQLUTIL.datos(11, 3)
    DATO15.text = SQLUTIL.datos(13, 3)
    DATO16.text = SQLUTIL.datos(12, 3)
    
    
fin:
End Sub


Sub flechas(ByRef caja As TextBox, ByRef caja1 As TextBox, ByRef codigo As Integer)
    If codigo = 38 And caja.Enabled = True Then caja.SetFocus: caja.SelStart = 0
    If codigo = 40 And caja1.Enabled = True Then caja1.SetFocus: caja.SelStart = 0
End Sub
Sub GRABADETALLEFACTURA()
    Dim J As Double
    For J = 1 To Grid1.Rows - 2
    Dim lin As String
    LINEAS.text = J
    Call ceros(LINEAS)
    
    campos(0, 0) = "tipo"
    campos(1, 0) = "numero"
    campos(2, 0) = "linea"
    campos(3, 0) = "rut"
    campos(4, 0) = "cuentadelmayor"
    campos(5, 0) = "glosa"
    campos(6, 0) = "monto"
    campos(7, 0) = "dh"
    campos(8, 0) = "centrodecosto"
    campos(9, 0) = "tipoctacte"
    campos(10, 0) = "rutctacte"
    campos(11, 0) = ""
    
    campos(0, 1) = dato1.text
    campos(1, 1) = dato2.text
    campos(2, 1) = LINEAS.text
    campos(3, 1) = dato9.text + DV.Caption
    campos(4, 1) = Grid1.Cell(LINEAS, 1).text + Grid1.Cell(LINEAS, 2).text + Grid1.Cell(LINEAS, 3).text
    campos(5, 1) = Grid1.Cell(LINEAS, 6).text
    campos(6, 1) = Grid1.Cell(LINEAS, 7).text
    campos(7, 1) = Grid1.Cell(LINEAS, 8).text
    campos(8, 1) = Grid1.Cell(LINEAS, 5).text
    campos(10, 1) = Grid1.Cell(LINEAS, 4).text
    
    campos(9, 1) = ""
   
    campos(0, 2) = "detallefacturasdecompra"
    If modifi = 1 Then condicion = "TIPO=" + "'" + dato1.text + "' AND NUMERO=" + "'" + dato2.text + "' AND LINEA=" + "'" + LINEAS.text + "'"
   
    SQLUTIL.datos = campos
    Set SQLUTIL.conexion = db
    Call SQLUTIL.SQLUTIL(op, condicion)

Next J
End Sub
Sub GRABADETALLEIMPUESTOS()
    Dim J As Double
    For J = 1 To Grid2.Rows - 2
   
    If Val(Grid2.Cell(LINEAS, 3).text) <> 0 Then
    campos(0, 0) = "tipo"
    campos(1, 0) = "numero"
    campos(2, 0) = "rut"
    campos(3, 0) = "Cuenta"
    campos(4, 0) = "Monto"
    campos(5, 0) = ""
    campos(0, 1) = dato1.text
    campos(1, 1) = dato2.text
    campos(2, 1) = dato9.text + DV.Caption
    campos(3, 1) = Grid2.Cell(LINEAS, 1).text
    campos(4, 1) = Grid1.Cell(LINEAS, 3).text
    campos(5, 1) = ""
   
    campos(0, 2) = "facturasdecompraspropias_impuestos"
   
    SQLUTIL.datos = campos
    Set SQLUTIL.conexion = db
    Call SQLUTIL.SQLUTIL(op, condicion)
    End If
    
    Next J
End Sub

Sub grabafactura()
    Dim NETOS As Double
    Dim DH As String
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
    campos(10, 0) = "añocontable"
    campos(11, 0) = "mescontable"
    campos(12, 0) = "comentario"
    campos(13, 0) = "electronica"
    campos(14, 0) = "activo"
    campos(15, 0) = "fechadigitacion"
    campos(16, 0) = "folio"
    campos(17, 0) = ""
    campos(0, 1) = dato1.text
    campos(1, 1) = dato2.text
    campos(2, 1) = dato5.text + "-" + dato4.text + "-" + dato3.text
    campos(3, 1) = dato8.text + "-" + dato7.text + "-" + dato6.text
    campos(4, 1) = dato9.text + DV.Caption
    campos(5, 1) = Replace(dato11.text, ".", "")
    campos(6, 1) = Replace(dato12.text, ".", "")
    campos(7, 1) = Replace(dato13.text, ".", "")
    campos(8, 1) = Replace(dato14.text, ".", "")
    campos(9, 1) = Replace(total.text, ".", "")
    campos(10, 1) = año
    campos(11, 1) = mes
    campos(12, 1) = DATO17.text
    campos(13, 1) = DATO16.text
    campos(14, 1) = DATO15.text
    campos(15, 1) = fechasistema
    
    campos(16, 1) = folios.text
    
    If dato1.text = "1" Then campos(15, 1) = "FP"
    If dato1.text = "2" Then campos(15, 1) = "DC"
    If dato1.text = "3" Then campos(15, 1) = "NC"
    
    condicion = ""
    campos(0, 2) = "facturasdecompraspropias"
    op = 2
    SQLUTIL.datos = campos
    Set SQLUTIL.conexion = db
    Call SQLUTIL.SQLUTIL(op, condicion)
    

grabar2

End Sub


Sub grabar2()
GRABADETALLEFACTURA
GRABADETALLEIMPUESTOS
grabarcomprobante

LEERMOVIMIENTOS
opciones.Visible = True
opciones.SetFocus
detalle.Enabled = False
End Sub
Sub ELIMINAR()
    Dim tipocon As String
    
    Call ACTUALIZADOCUMENTO("-")

    
    campos(0, 2) = "detallefacturasdecompra"
    condicion = "tipo=" + "'" + dato1.text + "'" + " and numero=" + "'" + dato2.text + "' order by linea"
    op = 4
    SQLUTIL.datos = campos
    Set SQLUTIL.conexion = db
    Call SQLUTIL.SQLUTIL(op, condicion)
    
    campos(0, 2) = "facturasdecompraspropias_impuestos"
    condicion = "tipo=" + "'" + dato1.text + "'" + " and numero=" + "'" + dato2.text + "'" + " and rut=" + "'" + dato9.text + DV.Caption + "'"
    op = 4
    SQLUTIL.datos = campos
    Set SQLUTIL.conexion = db
    Call SQLUTIL.SQLUTIL(op, condicion)
    
    campos(0, 2) = "facturasdecompraspropias"
    condicion = "tipo=" + "'" + dato1.text + "'" + " and numero=" + "'" + dato2.text + "'" + " and rut=" + "'" + dato9.text + DV.Caption + "'"
    op = 4
    SQLUTIL.datos = campos
    Set SQLUTIL.conexion = db
    Call SQLUTIL.SQLUTIL(op, condicion)
    campos(0, 2) = "movimientoscontables"
    If dato1.text = "1" Then tipocon = "FP"
    If dato1.text = "2" Then tipocon = "DC"
    If dato1.text = "3" Then tipocon = "NC"
     
    condicion = "tipo=" + "'" + tipocon + "'" + " and numero=" + "'" + dato2.text + "'"
    op = 4
    SQLUTIL.datos = campos
    Set SQLUTIL.conexion = db
    Call SQLUTIL.SQLUTIL(op, condicion)
no:
End Sub


Private Sub glosafactura_Change()

End Sub



Private Sub MSHFlexGrid1_Click()

End Sub



Private Sub Grid1_LeaveCell(ByVal Row As Long, ByVal Col As Long, NewRow As Long, NewCol As Long, Cancel As Boolean)
    Dim TEXTO As String
    Dim DV As String

    If Row = 0 And Col = 0 Then NewRow = 1: NewCol = 1: GoTo no:
    If Col = 1 And NewCol > 11 Then NewCol = 11
    If Col = 12 And NewCol > Col Then NewCol = 1
    TEXTO = Grid1.Cell(Row, Col).text
    For K = 1 To 11
    If Col = K And Row > 1 And Grid1.Cell(Row, Col).text = "" Then Grid1.Cell(Row, Col).text = Grid1.Cell(Row - 1, Col).text
    Next K

    If Col = 12 And Row > 1 And Grid1.Cell(Row - 1, Col).text = "D" And Grid1.Cell(Row, Col).text = "" Then Grid1.Cell(Row, Col).text = "H"
    If Col = 12 And Row > 1 And Grid1.Cell(Row - 1, Col).text = "H" And Grid1.Cell(Row, Col).text = "" Then Grid1.Cell(Row, Col).text = "D"

    If Col = 1 And Row = Grid1.Rows - 1 And NewRow < Row Then GoTo paso2:

    If Col = 1 Then
        pivote.MaxLength = 2: pivote.text = Grid1.Cell(Row, Col).text: Call ceros(pivote): Grid1.Cell(Row, Col).text = pivote.text
        If Grid1.Cell(Row, Col).text = "00" And NewCol > Col Then
        Col = 1: NewCol = 1
        End If
    End If

    If Col = 2 Then
        pivote.MaxLength = 2: pivote.text = Grid1.Cell(Row, Col).text: Call ceros(pivote): Grid1.Cell(Row, Col).text = pivote.text
        If Grid1.Cell(Row, Col).text = "00" And NewCol > Col Then
        Col = 2: NewCol = 2
        End If
    End If
    If Col = 3 Then
        pivote.MaxLength = 4: pivote.text = Grid1.Cell(Row, Col).text: Call ceros(pivote): Grid1.Cell(Row, Col).text = pivote.text
        If Grid1.Cell(Row, Col).text = "0000" And NewCol > Col Then
        Col = 3: NewCol = 3
        End If
    End If


    If Col = 3 And NewCol = 4 Then Call leermayor(Row, Col)



10:


    'If NewRow <> Row And tienecheque(Row) = "1" Then Call grabacheque(Row)

    If Col < 7 And NewRow <> Row Then NewRow = Row

    If NewCol = 4 And Col = 4 And NewRow <> Row Then NewCol = 4: NewRow = Row
    If NewCol = 5 And Col = 5 And NewRow <> Row Then NewCol = 4: NewRow = Row
    If NewCol = 6 And Col = 6 And NewRow <> Row Then NewCol = 4: NewRow = Row
'    If NewCol = 4 And Col = 4 And NewRow <> Row Then NewCol = 5: NewRow = Row
'    If NewCol = 5 And Col = 5 And NewRow <> Row Then NewCol = 5: NewRow = Row
'
'    If NewCol = 4 And Col < NewCol And TIENECTACTE(Row) = "0" Then NewCol = 5
'    If NewCol = 5 And Col < NewCol And TIENECRCC(Row) = "0" Then NewCol = 6
'    If NewCol = 5 And Col > NewCol And TIENECRCC(Row) = "0" Then NewCol = 6
'    If NewCol = 4 And Col > NewCol And TIENECTACTE(Row) = "0" Then NewCol = 3
 
    
    If NewCol = 4 And Col < NewCol And TIENECTACTE(Row) = "0" Then NewCol = 5
    If NewCol = 5 And Col < NewCol And TIENECRCC(Row) = "0" Then NewCol = 6
    If NewCol = 5 And Col > NewCol And TIENECRCC(Row) = "0" Then NewCol = 4
    If NewCol = 4 And Col > NewCol And TIENECTACTE(Row) = "0" Then NewCol = 3

    If NewCol = 4 And Col > NewCol Then NewCol = 3
    If Col = 4 And Val(Grid1.Cell(Row, Col).text) = 0 And NewCol > Col Then Col = 4: NewCol = 4: NewRow = Row: GoTo 20
         If Col = 4 Then
         pivote.MaxLength = 9
         pivote.text = Grid1.Cell(Row, Col).text
         Call ceros(pivote)
         Grid1.Cell(Row, Col).text = pivote.text
         End If

    If Col = 4 Then

            DV = rut(Grid1.Cell(Row, Col).text)
            Grid1.Cell(Row, Col).text = pivote.text + DV
    End If

5:    If Col = 4 And Val(Mid(Grid1.Cell(Row, Col).text, 1, 9)) <> 0 Then
         Call leerctacte(Row, Col)
         If EXISTE = "N" Then
         Call CREARCTACTE(Row)
         NewCol = 4: Col = 4
         End If

    End If

    If Col = 4 And NewCol = 6 And respu = "N" Then NewCol = 5: NewRow = Row
20: If NewCol = 5 And Col < NewCol And TIENECRCC(Row) = "0" Then NewCol = 6: GoTo paso2:
    If Col = 5 Then pivote.MaxLength = 4: pivote.text = Grid1.Cell(Row, Col).text: Call ceros(pivote): Grid1.Cell(Grid1.ActiveCell.Row, Grid1.ActiveCell.Col).text = pivote.text
    If NewCol = 6 And Col = 5 Then Call leercrcc(Row, Col)
paso2:
    If Col > 7 Then SUMAR
    If Row = Grid1.Rows - 1 And Col = Grid1.Cols - 1 And NewCol = 1 Then Grid1.Rows = Grid1.Rows + 1: NewRow = Grid1.Rows - 1
    For K = 1 To NewCol - 1
    'If Grid1.Cell(Row, K).text = String(Grid1.Column(K).MaxLength, "0") And K <> 4 Then NewCol = K: Exit For

    If Grid1.Cell(Row, K).text = "" Then NewCol = K: Exit For
    Next K
 If NewRow <> Row Then
    If Grid1.Cell(Row, 8).text = "" Then NewRow = Row: NewCol = 7



End If


no:
End Sub
'Private Sub Grid1_LeaveCell(ByVal Row As Long, ByVal Col As Long, NewRow As Long, NewCol As Long, Cancel As Boolean)
'    Dim TEXTO As String
'    Dim DV As String
'    If Row = 0 And Col = 0 Then NewRow = 1: NewCol = 1: GoTo no:
'    If Col = 1 And NewCol > 11 Then NewCol = 11
'    If Col = 12 And NewCol > Col Then NewCol = 1
'    TEXTO = Grid1.Cell(Row, Col).text
'    For K = 1 To 11
'    If Col = K And Row > 1 And Grid1.Cell(Row, Col).text = "" Then Grid1.Cell(Row, Col).text = Grid1.Cell(Row - 1, Col).text
'    Next K
'    If Col = 12 And Row > 1 And Grid1.Cell(Row - 1, Col).text = "D" And Grid1.Cell(Row, Col).text = "" Then Grid1.Cell(Row, Col).text = "H"
'    If Col = 12 And Row > 1 And Grid1.Cell(Row - 1, Col).text = "H" And Grid1.Cell(Row, Col).text = "" Then Grid1.Cell(Row, Col).text = "D"
'    If Col = 1 And Row = Grid1.Rows - 1 And NewRow < Row Then GoTo paso2:
'
'    If Col = 1 Then
'        PIVOTE.MaxLength = 2: PIVOTE.text = Grid1.Cell(Row, Col).text: Call ceros(PIVOTE): Grid1.Cell(Row, Col).text = PIVOTE.text
'        If Grid1.Cell(Row, Col).text = "00" Then
'        Col = 1: NewCol = 1
'        End If
'    End If
'
'    If Col = 2 Then
'        PIVOTE.MaxLength = 2: PIVOTE.text = Grid1.Cell(Row, Col).text: Call ceros(PIVOTE): Grid1.Cell(Row, Col).text = PIVOTE.text
'        If Grid1.Cell(Row, Col).text = "00" Then
'        Col = 2: NewCol = 2
'        End If
'    End If
'    If Col = 3 Then
'        PIVOTE.MaxLength = 4: PIVOTE.text = Grid1.Cell(Row, Col).text: Call ceros(PIVOTE): Grid1.Cell(Row, Col).text = PIVOTE.text
'        If Grid1.Cell(Row, Col).text = "0000" Then
'        Col = 3: NewCol = 3
'        End If
'    End If
'
'
'    If Col = 3 And NewCol = 4 Then Call leermayor(Row, Col)
'
'10:
'    If NewRow <> Row Then Call leermayor(NewRow, 3)
'    If Col < 7 And NewRow <> Row Then NewRow = Row
'
'    If NewCol = 4 And Col = 4 And NewRow <> Row Then NewCol = 5: NewRow = Row
'    If NewCol = 5 And Col = 5 And NewRow <> Row Then NewCol = 5: NewRow = Row
'
'    If NewCol = 4 And Col < NewCol And TIENECTACTE(Row) = "0" Then NewCol = 5
'    If NewCol = 5 And Col < NewCol And TIENECRCC(Row) = "0" Then NewCol = 6
'    If NewCol = 5 And Col > NewCol And TIENECRCC(Row) = "0" Then NewCol = 6
'    If NewCol = 4 And Col > NewCol And TIENECTACTE(Row) = "0" Then NewCol = 3
'
'    If NewCol = 4 And Col = 4 And NewRow <> Row Then NewCol = 5: NewRow = Row
'    If NewCol = 5 And Col = 5 And NewRow <> Row Then NewCol = 5: NewRow = Row
'
'    If NewCol = 4 And Col < NewCol And TIENECTACTE(Row) = "0" Then NewCol = 5
'    If NewCol = 5 And Col < NewCol And TIENECRCC(Row) = "0" Then NewCol = 6
'    If NewCol = 5 And Col > NewCol And TIENECRCC(Row) = "0" Then NewCol = 6
'    If NewCol = 4 And Col > NewCol And TIENECTACTE(Row) = "0" Then NewCol = 3


'
'    If Col = 4 Then PIVOTE.MaxLength = 9: PIVOTE.text = Grid1.Cell(Row, Col).text: Call ceros(PIVOTE): Grid1.Cell(Row, Col).text = PIVOTE.text
'    If Col = 4 Then: DV = rut(Grid1.Cell(Row, Col).text): Grid1.Cell(Row, Col).text = PIVOTE.text + DV
'
'    If Col = 4 And NewCol = 6 Then Call leerctacte(Row, Col)
'20: If NewCol = 5 And Col < NewCol And TIENECRCC(Row) = "0" Then NewCol = 6: GoTo 30:
'    If Col = 5 Then PIVOTE.MaxLength = 4: PIVOTE.text = Grid1.Cell(Row, Col).text: Call ceros(PIVOTE): Grid1.Cell(Grid1.ActiveCell.Row, Grid1.ActiveCell.Col).text = PIVOTE.text
'
'    If NewCol = 6 And Col = 5 Then Call leercrcc(Row, Col)
'paso2:
'30: If Col > 7 Then SUMAR
'    If Row = Grid1.Rows - 1 And Col = Grid1.Cols - 1 And NewCol = 1 Then Grid1.Rows = Grid1.Rows + 1: NewRow = Grid1.Rows - 1
'    For K = 1 To NewCol - 1
'    'If Grid1.Cell(Row, K).text = String(Grid1.Column(K).MaxLength, "0") And K <> 4 Then NewCol = K: Exit For
'
'    If Grid1.Cell(Row, K).text = "" And K > 4 Then NewCol = K: Exit For
'    Next K
'no:
'End Sub

Private Sub Grid2_LeaveCell(ByVal Row As Long, ByVal Col As Long, NewRow As Long, NewCol As Long, Cancel As Boolean)
If NewRow = cdi And Row = cdi Then impuestos.Visible = False: sumaimpuestos: dato14.Enabled = True: dato14.SetFocus

If NewCol <> 3 Then NewCol = 3

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
If command = "retorno" And modifi = 0 Then retorno
If command = "retorno" And modifi = 1 Then grabafactura: retorno
'If command = "modifica" Then ELIMINAR: dato2.Enabled = True: dato2.SetFocus

If command = "elimina" Then ELIMINAR: retorno

End Sub


Sub retorno()



Call CARGAGRILLA


opciones.Visible = False
limpia
disponible (False)

dato1.Enabled = True
dato2.Enabled = True
dato2.SetFocus

End Sub


Sub limpia()

    nombreproveedor.Caption = ""
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
    dato14.text = "0"
    DATO15.text = "N"
    total.text = "0"
    DATO16.text = "N"
    DATO17.text = ""
    
    LINEAS.text = "001"
SUMAR
no:
End Sub
Sub ayudactacte2(ByRef caja As TextBox)
    Dim campos As Variant
    Dim cfijo As Variant
    Dim largo As Variant
    campos = Array("rut", "nombre")
    largo = Array("12n", "40s")
    cfijo = "tipo='" & cuentaproveedor & "' and año='" + año + "'"
    CABEZAS = Array("RUT", "NOMBRE")
    mensajeAyuda = "Ayuda Cuentas Corrientes"
    Call cargaAyudaT(servidor, basebus, USUARIO, password, "cuentascorrientes", pivote2, campos, cfijo, largo, 2)
    If Val(pivote2.text) = 0 Then dato9.SetFocus: GoTo no
    dato9.text = Mid(pivote2.text, 1, 9)
    DV.Caption = Mid(pivote2.text, 10, 1)
    caja.Enabled = True
    caja.SetFocus
no:

End Sub


Sub LEERMOVIMIENTOS()
    Dim resultados As rdoResultset
    Dim cSql As New rdoQuery
    Dim rut As String
    Dim SALDOS As Double
    Dim sumadebe As Double
    Dim sumahaber As Double
    nlineas = 0
    
        Set cSql.ActiveConnection = db
        cSql.SQL = "SELECT linea,cuentadelmayor,rutctacte,centrodecosto,glosa,monto,dh "
        cSql.SQL = cSql.SQL + "FROM detallefacturasdecompra"
        cSql.SQL = cSql.SQL + " where tipo=" + "'" + dato1.text + "'" + " and numero=" + "'" + dato2.text + "' and rut='" + dato9.text + DV.Caption + "' order by linea"
        ' cSql.SQL = cSql.SQL + " where tipo=1 and numero=0000000005 order by linea"
        cSql.Execute
        
        linea = 0: SUMADOR = 0

        If cSql.RowsAffected > 0 Then
            
            Set resultados = cSql.OpenResultset
            sumadebe = 0
            SALDOS = 0
            sumahaber = 0
            Grid1.Rows = 1
            While Not resultados.EOF
                linea = linea + 1
                
                Grid1.Rows = Grid1.Rows + 1
                K = resultados(0)
            
                Grid1.Cell(linea, 1).text = Mid(resultados(1), 1, 2)
                Grid1.Cell(linea, 2).text = Mid(resultados(1), 3, 2)
                Grid1.Cell(linea, 3).text = Mid(resultados(1), 5, 4)
                Grid1.Cell(linea, 4).text = resultados(2)
                Grid1.Cell(linea, 5).text = resultados(3)
                Grid1.Cell(linea, 6).text = resultados(4)
                Grid1.Cell(linea, 7).text = resultados(5)
                Grid1.Cell(linea, 8).text = resultados(6)
                If resultados(6) = "D" Then sumadebe = sumadebe + CDbl(resultados(5))
                If resultados(6) = "H" Then sumahaber = sumahaber + CDbl(resultados(5))
                resultados.MoveNext
            Wend
            resultados.Close
            Set resultados = Nothing

        End If
    nlineas = cSql.RowsAffected
    LINEAS.text = CDbl(linea) + 1
    Call ceros(LINEAS)
    Rem If VARIPASO <> "0" Then CARGADATAFIELD
    saldo = sumadebe - NETO
    
    debe.Caption = Format(sumadebe - sumahaber, "##,###,###,##0")
    haber.Caption = Format(NETO, "##,###,###,##0")
    saldo.Caption = Format(saldo, "##,###,###,##0")
    
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
    SQLUTIL.datos = campos
    Set SQLUTIL.conexion = db
    Call SQLUTIL.SQLUTIL(op, condicion)

    If SQLUTIL.estado = 4 Then
    nombremayor.Caption = SQLUTIL.datos(1, 3)
    End If
    If Val(SQLUTIL.datos(2, 3)) <> 0 Then tipocuenta = SQLUTIL.datos(2, 3)
    tipocentro = SQLUTIL.datos(3, 3)

no:

End Sub
Sub leectacte()
    campos(0, 0) = "rut"
    campos(1, 0) = "nombre"
    campos(2, 0) = ""
    campos(0, 2) = "cuentascorrientes"
    condicion = "tipo=" + "'" + cuentaproveedor + "' and rut=" + "'" + pivote2.text + "' and año='" + año + "'"
    op = 5
    SQLUTIL.datos = campos
    Set SQLUTIL.conexion = db
    Call SQLUTIL.SQLUTIL(op, condicion)
    If SQLUTIL.estado = 4 Then
                maestro02.dato1.Enabled = True
                maestro02.dato2.Enabled = True
                maestro02.DV.Caption = True
                maestro02.dato1.text = cuentaproveedor
                maestro02.dato2.text = dato9.text
                maestro02.DV.Caption = DV.Caption
                cierrect = "S"
                maestro02.Show
                
                GoTo no:
    
    End If
    
    nombreproveedor.Caption = SQLUTIL.datos(1, 3)
    dato3.Enabled = True
    dato3.SetFocus
no:

End Sub


Sub leetipos()
    
    campos(0, 0) = "tipos"
    campos(1, 0) = "nombredocumento"
    campos(2, 0) = ""
    
    campos(0, 2) = "maestrotipodedocumentos"
    condicion = "tipos=" + "'" + dato13.text + "'"
    op = 5
    SQLUTIL.datos = campos
    Set SQLUTIL.conexion = conta
    Call SQLUTIL.SQLUTIL(op, condicion)

    If SQLUTIL.estado = 4 Then dato13.text = "": dato13.SetFocus:  GoTo no:
    VARIPASO = "S"
    

no:

End Sub
Sub cargatexto(ByRef caja As TextBox)


caja.SelStart = 0: caja.SelLength = Len(caja.text)

End Sub






Sub totalfactura()
SUMADOR = CDbl(Replace(dato11.text, ",", "")) + CDbl(Replace(dato12.text, ",", "")) + CDbl(Replace(dato13.text, ",", "") - CDbl(Replace(dato14.text, ",", "")))
total.text = Format(SUMADOR, "###,###,###,##0")
NETO = CDbl(Replace(dato11.text, ",", "")) + CDbl(Replace(dato13.text, ",", ""))
haber.Caption = Format(NETO, "###,###,##0")
End Sub
Sub leefactura()
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
    campos(10, 0) = "añocontable"
    campos(11, 0) = "mescontable"
    campos(12, 0) = "comentario"
    campos(13, 0) = "electronica"
    campos(14, 0) = "activo"
    campos(15, 0) = "fechadigitacion"
    campos(16, 0) = ""
    campos(0, 2) = "facturasdecompraspropias"
    condicion = "tipo=" + "'" + dato1.text + "'" + " and numero=" + "'" + dato2.text + "'" + " and rut=" + "'" + dato9.text + DV.Caption + "'"

    op = 5
    SQLUTIL.datos = campos
    Set SQLUTIL.conexion = db
    Call SQLUTIL.SQLUTIL(op, condicion)
    Rem If SQLUTIL.estado = 0 Then modifi = 1: carga: opciones.Visible = True: disponible (True): habilita (True): opciones.SetFocus


End Sub




Sub grabarcomprobante()
    Dim tipocon As String
    Dim tipo2 As String
    Dim J As Integer
    Dim lin As Integer
    
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
    campos(18, 0) = ""
    Rem cuenta proveedores
    If dato1.text = "1" Then tipocon = "FP"
    If dato1.text = "2" Then tipocon = "DC"
    If dato1.text = "3" Then tipocon = "NC"
    
    campos(0, 1) = tipocon
    campos(1, 1) = dato2.text
    campos(2, 1) = "001"
    campos(3, 1) = dato5.text + "-" + dato4.text + "-" + dato3.text
    campos(4, 1) = cuentaproveedor
    campos(5, 1) = tipocuenta
    campos(6, 1) = dato9.text + DV.Caption
    campos(7, 1) = ""
    campos(8, 1) = "CONTABILIZACION " + DOCU$(Val(dato1.text)) + " " + nombreproveedor.Caption
    campos(9, 1) = DOCU$(Val(dato1.text))
    campos(10, 1) = dato2.text
    campos(11, 1) = dato5.text + "-" + dato4.text + "-" + dato3.text
    campos(12, 1) = dato8.text + "-" + dato7.text + "-" + dato6.text
    campos(13, 1) = Replace(total.text, ".", "")

    campos(14, 1) = "H"
    campos(15, 1) = USUARIOSISTEMA
    campos(16, 1) = mes
    campos(17, 1) = año

    campos(0, 2) = "movimientoscontables"
    condicion = ""

    op = 2
    SQLUTIL.datos = campos
    Set SQLUTIL.conexion = db
    Call SQLUTIL.SQLUTIL(op, condicion)
    
    
    Rem cuenta I.V.A
    
    
    campos(0, 1) = tipocon
    campos(1, 1) = dato2.text
    campos(2, 1) = "002"
    campos(3, 1) = dato5.text + "-" + dato4.text + "-" + dato3.text
    campos(4, 1) = ivacredito
    campos(5, 1) = ""
    campos(6, 1) = ""
    campos(7, 1) = ""
    campos(8, 1) = "CONTABILIZACION I.V.A " + DOCU(Val(dato1.text))
    campos(9, 1) = DOCU(Val(dato1.text))
    campos(10, 1) = dato2.text
    campos(11, 1) = dato5.text + "-" + dato4.text + "-" + dato3.text
    campos(12, 1) = dato5.text + "-" + dato4.text + "-" + dato3.text
    campos(13, 1) = Replace(dato12.text, ".", "")
    
    campos(14, 1) = "D"
    campos(15, 1) = USUARIOSISTEMA
    campos(16, 1) = mes
    campos(17, 1) = año
    campos(0, 2) = "movimientoscontables"
    condicion = ""

    op = 2
    SQLUTIL.datos = campos
    Set SQLUTIL.conexion = db
    Call SQLUTIL.SQLUTIL(op, condicion)
Rem graba impuestos
    lin = 2
    For J = 1 To Grid2.Rows - 1
   
    If Val(Grid2.Cell(J, 3).text) <> 0 Then
    lin = lin + 1
    LINEAS.text = lin
    Call ceros(LINEAS)
    
    campos(0, 1) = tipocon
    campos(1, 1) = dato2.text
    campos(2, 1) = LINEAS.text
    campos(3, 1) = dato5.text + "-" + dato4.text + "-" + dato3.text
    campos(4, 1) = Grid2.Cell(J, 1).text
    campos(5, 1) = ""
    campos(6, 1) = ""
    campos(7, 1) = ""
    campos(8, 1) = Grid2.Cell(J, 2).text
    campos(9, 1) = DOCU(Val(dato1.text))
    campos(10, 1) = dato2.text
    campos(11, 1) = dato5.text + "-" + dato4.text + "-" + dato3.text
    campos(12, 1) = dato5.text + "-" + dato4.text + "-" + dato3.text
    campos(13, 1) = Grid2.Cell(J, 3).text
    
    campos(14, 1) = "D"
    campos(15, 1) = USUARIOSISTEMA
    campos(16, 1) = mes
    campos(17, 1) = año
    campos(0, 2) = "movimientoscontables"
    condicion = ""
    SQLUTIL.datos = campos
    Set SQLUTIL.conexion = db
    Call SQLUTIL.SQLUTIL(op, condicion)
    End If
    
    Next J

    
    
    For K = 1 To Grid1.Rows - 2
    lin = lin + 1
    LINEAS.text = lin
    Call ceros(LINEAS)
    campos(0, 1) = tipocon
    campos(1, 1) = dato2.text
    campos(2, 1) = LINEAS.text
    campos(3, 1) = dato5.text + "-" + dato4.text + "-" + dato3.text
    campos(4, 1) = Grid1.Cell(K, 1).text + Grid1.Cell(K, 2).text + Grid1.Cell(K, 3).text
    campos(5, 1) = ""
    campos(6, 1) = Grid1.Cell(K, 4).text
    campos(7, 1) = Grid1.Cell(K, 5).text
    campos(8, 1) = Grid1.Cell(K, 6).text
    campos(9, 1) = DOCU(Val(dato1.text))
    campos(10, 1) = dato2.text
    campos(11, 1) = dato5.text + "-" + dato4.text + "-" + dato3.text
    campos(12, 1) = dato8.text + "-" + dato7.text + "-" + dato6.text
    campos(13, 1) = Grid1.Cell(K, 7).text
    campos(14, 1) = Grid1.Cell(K, 8).text
    campos(15, 1) = USUARIOSISTEMA
    campos(16, 1) = mes
    campos(17, 1) = año
    campos(0, 2) = "movimientoscontables"
    condicion = ""

    op = 2
    SQLUTIL.datos = campos
    Set SQLUTIL.conexion = db
    Call SQLUTIL.SQLUTIL(op, condicion)
    Next K
    Call ACTUALIZADOCUMENTO("+")
   
End Sub




Sub elimina()
Call ACTUALIZADOCUMENTO("-")
End Sub






Sub CARGAGRILLA()

    FORMATOGRILLA(1, 1) = "C1"
    FORMATOGRILLA(1, 2) = "C2"
    FORMATOGRILLA(1, 3) = "C3"
    FORMATOGRILLA(1, 4) = "RUT"
    FORMATOGRILLA(1, 5) = "CRCC"
    FORMATOGRILLA(1, 6) = "GLOSA"
    FORMATOGRILLA(1, 7) = "MONTO"
    FORMATOGRILLA(1, 8) = "D/H"
    Rem LARGO DE LOS DATOS
    FORMATOGRILLA(2, 1) = "2"
    FORMATOGRILLA(2, 2) = "2"
    FORMATOGRILLA(2, 3) = "4"
    FORMATOGRILLA(2, 4) = "10"
    FORMATOGRILLA(2, 5) = "4"
    FORMATOGRILLA(2, 6) = "27"
    FORMATOGRILLA(2, 7) = "11"
    FORMATOGRILLA(2, 8) = "1"
    FORMATOGRILLA(2, 9) = "0"
    
    Rem TIPO DE DATOS C=CEROS,S=STRING,N=NUMERICO,D=DATE,
    FORMATOGRILLA(3, 1) = "C"
    FORMATOGRILLA(3, 2) = "C"
    FORMATOGRILLA(3, 3) = "C"
    FORMATOGRILLA(3, 4) = "C"
    FORMATOGRILLA(3, 5) = "C"
    FORMATOGRILLA(3, 6) = "S"
    FORMATOGRILLA(3, 7) = "N"
    FORMATOGRILLA(3, 8) = "S"
    Rem FORMATO GRILLA
    FORMATOGRILLA(4, 7) = " ###,###,##0"
    Rem LOCCKED
    FORMATOGRILLA(5, 1) = "FALSE"
    FORMATOGRILLA(5, 2) = "FALSE"
    FORMATOGRILLA(5, 3) = "FALSE"
    FORMATOGRILLA(5, 4) = "FALSE"
    FORMATOGRILLA(5, 5) = "FALSE"
    FORMATOGRILLA(5, 6) = "FALSE"
    FORMATOGRILLA(5, 7) = "FALSE"
    FORMATOGRILLA(5, 8) = "FALSE"
    
    Rem VALOR MAXIMO
    FORMATOGRILLA(7, 6) = "999999999"
    Grid1.Cols = 9
    Grid1.Rows = 1
    Grid1.AllowUserResizing = False
    Grid1.DisplayFocusRect = False
    Grid1.ExtendLastCol = False
    Grid1.BoldFixedCell = False
    Grid1.DrawMode = cellOwnerDraw
    Grid1.Appearance = Flat
    Grid1.ScrollBarStyle = Flat
    Grid1.FixedRowColStyle = Flat
    Grid1.Column(0).Width = 0
    For K = 1 To Grid1.Cols - 1
        Grid1.Cell(0, K).text = FORMATOGRILLA(1, K)
        If K < 5 Then Grid1.Column(K).Width = Val(FORMATOGRILLA(2, K)) * 10.5
        If K > 4 Then Grid1.Column(K).Width = Val(FORMATOGRILLA(2, K)) * 8.8
        
        
        Rem Grid1.Column(K).Width = Val(formatogrilla(2, K)) * 9
        Grid1.Column(K).MaxLength = Val(FORMATOGRILLA(2, K))
        Grid1.Column(K).FormatString = FORMATOGRILLA(4, K)
        Grid1.Column(K).Locked = FORMATOGRILLA(5, K)
        If FORMATOGRILLA(3, K) = "N" Then Grid1.Column(K).Alignment = cellRightCenter
        If FORMATOGRILLA(3, K) = "D" Then Grid1.Column(K).CellType = cellCalendar
        
    Next K
    Grid1.Column(8).Width = 3 * 9
    Grid1.Range(0, 1, 0, 3).Merge
    Grid1.Cell(0, 1).text = "CUENTA"
    detalle.Enabled = False
    
    
    
    End Sub



Sub ayudamayor(Row As Long, Col As Long)
    Dim campos As Variant
    Dim cfijo As Variant
    Dim largo As Variant
    campos = Array("codigo", "nombre")
    CABEZAS = Array("codigo", "nombre")
    largo = Array("8n", "40s")
    mensajeAyuda = "Ayuda Cuentas del Mayor"
    cfijo = "no"
    Call cargaAyudaT(servidor, basebus, USUARIO, password, "cuentasdelmayor", pivote2, campos, cfijo, largo, 2)
    Grid1.Cell(Row, Col).text = Mid(pivote2.text, 1, 2)
    Grid1.Cell(Row, Col + 1).text = Mid(pivote2.text, 3, 2)
    Grid1.Cell(Row, Col + 2).text = Mid(pivote2.text, 5, 4)
    Rem Call leermayor(row, col)
    respu = ""
    If pivote2.text <> "" Then Call leermayor(Row, Col): respu = "S"
    pivote2.text = ""
    
End Sub
Sub ayudacrcc(Row As Long, Col As Long)
    Dim campos As Variant
    Dim cfijo As Variant
    Dim largo As Variant
    campos = Array("codigo", "nombre")
    CABEZAS = Array("codigo", "nombre")
    largo = Array("8n", "40s")
    mensajeAyuda = "Ayuda Centros de costo"
    cfijo = "no"
    pivote2.MaxLength = 4
    Call cargaAyudaT(servidor, basebus, USUARIO, password, "centrosdecosto", pivote2, campos, cfijo, largo, 2)
    Grid1.Cell(Row, Col).text = pivote2.text

    pivote2.text = ""
End Sub
Private Sub Grid1_KeyDown(KeyCode As Integer, ByVal Shift As Integer)
    
    
    If KeyCode = 35 And Grid1.ActiveCell.Col = 1 And SALDOPE = 0 Then grabafactura: Stop
   
    Rem If KeyCode = 38 And Grid1.ActiveCell.row = Grid1.Rows - 1 Then SG = "S" Else SG = "N"
    If Grid1.ActiveCell.Col = "1" And KeyCode = vbKeyF2 Then Call ayudamayor(Grid1.ActiveCell.Row, Grid1.ActiveCell.Col)
    If Grid1.ActiveCell.Col = "4" And KeyCode = vbKeyF2 Then Call ayudactacte(Grid1.ActiveCell.Row, Grid1.ActiveCell.Col)
    If Grid1.ActiveCell.Col = "5" And KeyCode = vbKeyF2 Then Call ayudacrcc(Grid1.ActiveCell.Row, Grid1.ActiveCell.Col)
   
    End Sub

Private Sub Grid1_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    
    If Grid1.ActiveCell.Col = 1 And Chr(KeyAscii) = "*" And SALDOPE = 0 Then grabafactura
    If Grid1.ActiveCell.Col = 8 And Chr(KeyAscii) <> "D" And Chr(KeyAscii) <> "H" Then KeyAscii = 0
 
    If FORMATOGRILLA(3, Grid1.ActiveCell.Col) = "N" Then snum = 1: KeyAscii = esNumero(KeyAscii)
    If FORMATOGRILLA(3, Grid1.ActiveCell.Col) = "C" Then snum = 1: KeyAscii = esNumero(KeyAscii)
End Sub
Sub leermayor(Row As Long, Col As Long)
    TIENECTACTE(Row) = "0"
    TIENECRCC(Row) = "0"
    TIENEBANCO(Row) = "0"
    TIENEILA(Row) = "0"
    TIENEICA(Row) = "0"
    TIENEIHA(Row) = "0"
    TIENEACTIVO(Row) = "0"
    CUENTAMAYOR(Row) = "0"
    
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
    condicion = "codigo=" + "'" + Grid1.Cell(Row, 1).text + Grid1.Cell(Row, 2).text + Grid1.Cell(Row, 3).text + "'"
    op = 5
    SQLUTIL.datos = campos
    Set SQLUTIL.conexion = db
    Call SQLUTIL.SQLUTIL(op, condicion)
 
    If SQLUTIL.estado = 4 Or Grid1.Cell(Row, 3).text = "0000" Then
    RESPUESTA = "N"
    Grid1.Cell(Row, Col).text = ""
    Grid1.Cell(Row, Col + 1).text = ""
    Grid1.Cell(Row, Col + 2).text = ""
    Grid1.Cell(Row, Col + 5).text = ""
    Else
    RESPUESTA = "S"
    nombremayor.Caption = SQLUTIL.datos(1, 3)
    Rem If Grid1.Cell(Row, 5).text = "" Then Grid1.Cell(Row, 5).text = String(10, 32)
    Rem If Grid1.Cell(Row, 6).text = "" Then Grid1.Cell(Row, 6).text = String(4, 32)
    Rem If Grid1.Cell(Row, 4).text = "" Then Grid1.Cell(Row, 4).text = String(2, 32)
    TIENECTACTE(Row) = SQLUTIL.datos(2, 3)
    TIENECRCC(Row) = SQLUTIL.datos(3, 3)
    TIENEBANCO(Row) = SQLUTIL.datos(4, 3)
    TIENEILA(Row) = SQLUTIL.datos(5, 3)
    TIENEICA(Row) = SQLUTIL.datos(6, 3)
    TIENEIHA(Row) = SQLUTIL.datos(7, 3)
    TIENEACTIVO(Row) = SQLUTIL.datos(8, 3)
    CUENTAMAYOR(Row) = SQLUTIL.datos(0, 3)
     
    If TIENECTACTE(Row) = "0" Then Grid1.Cell(Row, 4).text = String(2, 32): Grid1.Cell(Row, 5).text = String(10, 32)
    If TIENECRCC(Row) = "0" Then Grid1.Cell(Row, 6).text = String(4, 32)
    
    End If

End Sub


Sub leercrcc(Row As Long, Col As Long)
    campos(0, 0) = "codigo"
    campos(1, 0) = "nombre"
    campos(2, 0) = ""
    campos(0, 2) = "centrosdecosto"
    condicion = "codigo=" + "'" + Grid1.Cell(Row, 5).text + "'"
    op = 5
    SQLUTIL.datos = campos
    Set SQLUTIL.conexion = db
    Call SQLUTIL.SQLUTIL(op, condicion)
   
    If SQLUTIL.estado = 4 Then Grid1.Cell(Row, 5).text = "": GoTo no:
    nombrecrcc.Caption = SQLUTIL.datos(1, 3)
no:

End Sub
Sub SUMAR()
Dim o As Integer
Dim sumadebe As Double
Dim sumahaber As Double

sumadebe = 0
sumahaber = NETO
SALDOPE = 0
For o = 1 To Grid1.Rows - 1
If Grid1.Cell(o, 8).text = "D" Then sumadebe = sumadebe + Grid1.Cell(o, 7).text
If Grid1.Cell(o, 8).text = "H" Then sumahaber = sumahaber + Grid1.Cell(o, 7).text
Next o
debe.Caption = Format(sumadebe, "###,###,###,##0")
haber.Caption = Format(sumahaber, "###,###,###,##0")
saldo.Caption = Format(sumadebe - sumahaber, "###,###,###,##0")
SALDOPE = sumadebe - sumahaber
End Sub

Sub disponible(ByVal condicion As Boolean)
    
    dato1.Enabled = condicion
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
    dato14.Enabled = condicion
    DATO15.Enabled = condicion
    
    total.Enabled = condicion
    DATO16.Enabled = condicion
    
    
End Sub


Sub ACTUALIZADOCUMENTO(COMANDO As String)
    Dim lin As Integer
  
    Dim resultados As rdoResultset
    Dim cSql As New rdoQuery
    Dim rut As String
    Dim TIPOFA As String
    
        If dato1.text = "1" Then TIPOFA = "FP"
        If dato1.text = "2" Then TIPOFA = "DC"
        If dato1.text = "3" Then TIPOFA = "NC"
        
        Set cSql.ActiveConnection = db
       cSql.SQL = "SELECT tipo,numero,linea,fecha,codigocuenta,tipoctacte,rutctacte,centrocosto,glosacontable,tipodocumento,numerodocumento,fechavencimiento,monto,dh "
      cSql.SQL = cSql.SQL + "FROM movimientoscontables "
     
            
        cSql.SQL = cSql.SQL + "WHERE tipo='" + TIPOFA + "' and numero='" & dato2.text & "'and año='" + año + "' and mes='" + mes + "' order by linea"
        cSql.Execute


        If cSql.RowsAffected > 0 Then
            Set resultados = cSql.OpenResultset
            While Not resultados.EOF
                
                Call actualizamayor(COMANDO, resultados(4), resultados(12), resultados(13), resultados(5), resultados(6), resultados(7), mes, año)
                
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
    campos(0, 0) = "folio"
    campos(1, 0) = ""
    campos(0, 2) = "facturasdecompraspropias"
    condicion = "mescontable=" + "'" + mes + "'" + " and añocontable=" + "'" + año + "' order by folio desc"
    op = 5
    SQLUTIL.datos = campos
    Set SQLUTIL.conexion = db
    Call SQLUTIL.SQLUTIL(op, condicion)
    If SQLUTIL.estado = 0 Then K = Val(SQLUTIL.datos(0, 3)) Else K = 0
    folios.text = K + 1
    
    Call ceros(folios)



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
    
    For K = 1 To Grid2.Cols - 1
        Grid2.Cell(0, K).text = formatogrilla2(1, K)
        If K < 5 Then Grid2.Column(K).Width = Val(formatogrilla2(2, K)) * 8
        
        
        Rem Grid1.Column(K).Width = Val(formatogrilla(2, K)) * 9
        Grid2.Column(K).MaxLength = Val(formatogrilla2(2, K))
        Grid2.Column(K).FormatString = formatogrilla2(4, K)
        Grid2.Column(K).Locked = formatogrilla2(5, K)
        If formatogrilla2(3, K) = "N" Then Grid2.Column(K).Alignment = cellRightCenter
        If formatogrilla2(3, K) = "D" Then Grid2.Column(K).CellType = cellCalendar
        
    Next K
   Call leecuentas
    
    End Sub

Sub leecuentas()

Dim resultados2 As rdoResultset
    Dim cSql2 As New rdoQuery
    Dim rut As String
    Dim PASO As String
    Dim LINEAS As Double
    
       
        
        Set cSql2.ActiveConnection = db
        cSql2.SQL = "SELECT codigo,nombre "
        cSql2.SQL = cSql2.SQL + "FROM cuentasdelmayor where ila<>'0' or iha<>'0' or ica<>'0'"
       
        cSql2.SQL = cSql2.SQL + "order by codigo"
        cSql2.Execute
        
        LINEAS = 0
        If cSql2.RowsAffected > 0 Then
        Set resultados2 = cSql2.OpenResultset
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
Sub sumaimpuestos()
Dim valor As Double
valor = 0
For K = 1 To Grid2.Rows - 1
valor = valor + Val(Grid2.Cell(K, 3).text)

Next K
dato13.text = valor

End Sub
Sub leerctacte(Row As Long, Col As Long)
    campos(0, 0) = "rut"
    campos(1, 0) = "nombre"
    campos(2, 0) = ""
    campos(0, 2) = "cuentascorrientes"
10: condicion = "tipo=" + "'" + CUENTAMAYOR(Row) + "' and rut=" + "'" + Grid1.Cell(Row, Col).text + "' and año='" + Format(fechasistema, "yyyy") + "'"
    op = 5
    SQLUTIL.datos = campos
    Set SQLUTIL.conexion = db
    Call SQLUTIL.SQLUTIL(op, condicion)
    If SQLUTIL.estado = 4 Then
    EXISTE = "N"
   
        Else
    EXISTE = "S"
    nombrectacte.Caption = SQLUTIL.datos(1, 3)
End If

no:

End Sub

Sub ayudactacte(Row As Long, Col As Long)
    Dim campos As Variant
    Dim cfijo As Variant
    Dim largo As Variant
    campos = Array("rut", "nombre")
    CABEZAS = Array("rut", "nombre")
    largo = Array("12n", "40s")
    mensajeAyuda = "Ayuda Cuentas Corrientes"
    cfijo = "tipo='" & CUENTAMAYOR(Row) & "'"
    pivote2.MaxLength = 10
    Call cargaAyudaT(servidor, basebus, USUARIO, password, "cuentascorrientes", pivote2, campos, cfijo, largo, 2)
    Grid1.Cell(Row, Col).text = pivote2.text
    pivote2.text = ""
End Sub

Sub CREARCTACTE(Row)
maestro02.dato1.text = Grid1.Cell(Row, 1).text + Grid1.Cell(Row, 2).text + Grid1.Cell(Row, 3).text
maestro02.dato2.text = Mid(Grid1.Cell(Row, 4).text, 1, 9)
maestro02.Show


End Sub

