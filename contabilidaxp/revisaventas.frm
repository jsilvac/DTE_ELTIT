VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{4777436C-EB8C-4596-98A8-EBCF98683969}#1.0#0"; "FlexCell.ocx"
Object = "{ADBBAED6-B16F-40DC-80DF-B44910CBA76C}#1.0#0"; "Frame-Xp.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form revisaventas 
   BackColor       =   &H00FF8080&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Revisa Factura de Venta"
   ClientHeight    =   7830
   ClientLeft      =   45
   ClientTop       =   1110
   ClientWidth     =   14160
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7830
   ScaleWidth      =   14160
   ShowInTaskbar   =   0   'False
   Begin XPFrame.FrameXp FrameQuickMenu 
      Height          =   615
      Left            =   10920
      TabIndex        =   70
      Top             =   0
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   1085
      BackColor       =   16761024
      Caption         =   " Mis Datos"
      BackColor       =   16761024
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
         TabIndex        =   72
         Top             =   280
         Width           =   1455
      End
      Begin VB.CommandButton botonmisfavoritos 
         Caption         =   "Mis Favoritos"
         Height          =   255
         Left            =   1800
         TabIndex        =   71
         Top             =   280
         Width           =   1335
      End
   End
   Begin XPFrame.FrameXp FrameXp1 
      Height          =   7890
      Left            =   135
      TabIndex        =   54
      Top             =   120
      Width           =   14145
      _ExtentX        =   24950
      _ExtentY        =   13917
      BackColor       =   16761024
      Caption         =   "REVISA LIBRO DE VENTAS"
      CaptionEstilo3D =   1
      BackColor       =   16761024
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
      Begin XPFrame.FrameXp FrameXp6 
         Height          =   660
         Left            =   8040
         TabIndex        =   60
         Top             =   840
         Width           =   5595
         _ExtentX        =   9869
         _ExtentY        =   1164
         BackColor       =   16744576
         Caption         =   "MES                                                          A�O                        FOLIOS"
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
         Begin VB.ComboBox COMBOMES 
            Height          =   315
            Left            =   45
            Locked          =   -1  'True
            TabIndex        =   63
            Top             =   270
            Width           =   2655
         End
         Begin VB.ComboBox COMBOA�O 
            Height          =   315
            Left            =   2745
            Locked          =   -1  'True
            TabIndex        =   62
            Top             =   270
            Width           =   1215
         End
         Begin VB.TextBox folios 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
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
            Height          =   330
            Left            =   4095
            Locked          =   -1  'True
            MaxLength       =   8
            TabIndex        =   61
            Tag             =   "tipo"
            Top             =   270
            Width           =   1335
         End
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00FF8080&
         Caption         =   "Ventas"
         Height          =   375
         Left            =   10560
         TabIndex        =   74
         Top             =   480
         Width           =   3255
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00FF8080&
         Caption         =   "Publicidad"
         Height          =   375
         Left            =   10560
         TabIndex        =   73
         Top             =   960
         Value           =   -1  'True
         Width           =   3255
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
         Left            =   1770
         MaxLength       =   1
         TabIndex        =   66
         Tag             =   "tipo"
         Top             =   360
         Width           =   255
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
         Left            =   1770
         MaxLength       =   10
         TabIndex        =   65
         Tag             =   "numero"
         Top             =   720
         Width           =   1455
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
         Left            =   1755
         MaxLength       =   9
         TabIndex        =   64
         Tag             =   "rut"
         Top             =   1065
         Width           =   1095
      End
      Begin MSComctlLib.ProgressBar BARRA 
         Height          =   285
         Left            =   135
         TabIndex        =   57
         Top             =   3870
         Width           =   13830
         _ExtentX        =   24395
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Revisa Contabilizaciones"
         Height          =   285
         Left            =   5445
         TabIndex        =   56
         Top             =   4275
         Width           =   2310
      End
      Begin VB.ListBox List1 
         BackColor       =   &H00FFFFFF&
         Height          =   1230
         Left            =   180
         TabIndex        =   55
         Top             =   1845
         Width           =   13695
      End
      Begin XPFrame.FrameXp detalle 
         Height          =   3255
         Left            =   0
         TabIndex        =   58
         Top             =   4635
         Width           =   14055
         _ExtentX        =   24791
         _ExtentY        =   5741
         BackColor       =   16761024
         Caption         =   "Comprobante Contable"
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
         Begin FlexCell.Grid Grid1 
            Height          =   2895
            Left            =   90
            TabIndex        =   59
            Top             =   270
            Width           =   13860
            _ExtentX        =   24448
            _ExtentY        =   5106
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
         Left            =   225
         TabIndex        =   69
         Top             =   360
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
         Left            =   225
         TabIndex        =   68
         Top             =   720
         Width           =   1455
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
         Left            =   225
         TabIndex        =   67
         Top             =   1080
         Width           =   1455
      End
   End
   Begin VB.TextBox PIVOTE4 
      Height          =   285
      Left            =   120
      MaxLength       =   9
      TabIndex        =   17
      Text            =   "Text1"
      Top             =   7680
      Visible         =   0   'False
      Width           =   615
   End
   Begin XPFrame.FrameXp cabeza 
      Height          =   2040
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   14055
      _ExtentX        =   24791
      _ExtentY        =   3598
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
      Begin XPFrame.FrameXp TIPOS 
         Height          =   2160
         Left            =   2025
         TabIndex        =   8
         Top             =   270
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   3810
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
            Height          =   1800
            Left            =   120
            TabIndex        =   9
            Top             =   240
            Width           =   3015
            _ExtentX        =   5318
            _ExtentY        =   3175
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
      Begin XPFrame.FrameXp FrameXp8 
         Height          =   975
         Left            =   135
         TabIndex        =   19
         Top             =   1350
         Width           =   13695
         _ExtentX        =   24156
         _ExtentY        =   1720
         BackColor       =   16761024
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
         Begin VB.TextBox DATO25 
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
            Left            =   9315
            MaxLength       =   15
            TabIndex        =   49
            Tag             =   "retencion"
            Text            =   "0"
            Top             =   495
            Width           =   1455
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
            Left            =   810
            MaxLength       =   4
            TabIndex        =   47
            Tag             =   "fecha"
            Top             =   510
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
            Left            =   450
            MaxLength       =   2
            TabIndex        =   46
            Tag             =   "fecha"
            Top             =   510
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
            Left            =   90
            MaxLength       =   2
            TabIndex        =   45
            Tag             =   "fecha"
            Top             =   510
            Width           =   375
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
            Left            =   13365
            MaxLength       =   1
            TabIndex        =   37
            Tag             =   "tipo"
            Text            =   "N"
            ToolTipText     =   "(S)i o (N)o"
            Top             =   585
            Width           =   255
         End
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
            Left            =   13365
            MaxLength       =   1
            TabIndex        =   36
            Tag             =   "tipo"
            Text            =   "N"
            ToolTipText     =   "(S)i o (N)o"
            Top             =   270
            Width           =   255
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
            Left            =   10800
            MaxLength       =   15
            TabIndex        =   35
            Tag             =   "monto"
            Text            =   "0"
            Top             =   495
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
            Left            =   7785
            MaxLength       =   15
            TabIndex        =   34
            Tag             =   "retencion"
            Text            =   "0"
            Top             =   495
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
            Top             =   495
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
            TabIndex        =   32
            Tag             =   "iva"
            Text            =   "0"
            Top             =   480
            Width           =   1455
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
            Left            =   3120
            MaxLength       =   15
            TabIndex        =   31
            Tag             =   "neto"
            Text            =   "0"
            Top             =   480
            Width           =   1455
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
            TabIndex        =   30
            Tag             =   "fecha"
            Top             =   480
            Width           =   615
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
            TabIndex        =   29
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
            TabIndex        =   28
            Tag             =   "fechavencimiento"
            Top             =   480
            Width           =   375
         End
         Begin VB.Label Label8 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "IMPTO.DIESEL"
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
            Left            =   9315
            TabIndex        =   50
            Top             =   225
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
            Left            =   90
            TabIndex        =   48
            Top             =   270
            Width           =   1335
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "F.ELEC."
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
            Left            =   12420
            TabIndex        =   27
            Top             =   270
            Width           =   840
         End
         Begin VB.Label Label4 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "ACTIVO "
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
            Left            =   12420
            TabIndex        =   26
            Top             =   585
            Width           =   855
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
            Height          =   300
            Left            =   10800
            TabIndex        =   25
            Top             =   240
            Width           =   1455
         End
         Begin VB.Label Label27 
            Alignment       =   2  'Center
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
            Left            =   7800
            TabIndex        =   24
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
            TabIndex        =   23
            Top             =   225
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
            TabIndex        =   22
            Top             =   240
            Width           =   1455
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
            TabIndex        =   21
            Top             =   240
            Width           =   1455
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
            TabIndex        =   20
            Top             =   240
            Width           =   1335
         End
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
         TabIndex        =   18
         Top             =   240
         Width           =   3135
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
         TabIndex        =   7
         Top             =   960
         Width           =   5175
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
         TabIndex        =   6
         Top             =   960
         Width           =   255
      End
   End
   Begin VB.PictureBox MANUAL 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF8080&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   990
      Left            =   0
      ScaleHeight     =   990
      ScaleWidth      =   2460
      TabIndex        =   4
      Top             =   720
      Width           =   2460
   End
   Begin VB.TextBox LINEAS 
      Height          =   285
      Left            =   120
      MaxLength       =   3
      TabIndex        =   3
      Top             =   7080
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox pivote2 
      Height          =   285
      Left            =   120
      MaxLength       =   10
      TabIndex        =   2
      Top             =   7080
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox PIVOTE 
      Height          =   285
      Left            =   0
      MaxLength       =   8
      TabIndex        =   0
      Top             =   7320
      Visible         =   0   'False
      Width           =   255
   End
   Begin XPFrame.FrameXp FrameXp3 
      Height          =   1095
      Left            =   6960
      TabIndex        =   10
      Top             =   6960
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
         TabIndex        =   11
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
            TabIndex        =   12
            Top             =   240
            Width           =   1575
         End
      End
      Begin XPFrame.FrameXp FrameXp5 
         Height          =   615
         Left            =   2760
         TabIndex        =   13
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
            TabIndex        =   14
            Top             =   240
            Width           =   1695
         End
      End
      Begin XPFrame.FrameXp FrameXp7 
         Height          =   615
         Left            =   4920
         TabIndex        =   15
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
            TabIndex        =   16
            Top             =   240
            Width           =   1695
         End
      End
   End
   Begin XPFrame.FrameXp CRCC 
      Height          =   735
      Left            =   180
      TabIndex        =   38
      Top             =   2745
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   1296
      BackColor       =   16744576
      Caption         =   "Nombre del Centro de Costo"
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
         Left            =   1320
         MaxLength       =   2
         TabIndex        =   40
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
         TabIndex        =   39
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
         Left            =   2400
         TabIndex        =   42
         Top             =   240
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
         TabIndex        =   41
         Top             =   240
         Width           =   1095
      End
   End
   Begin XPFrame.FrameXp glosafinal 
      Height          =   2775
      Left            =   3285
      TabIndex        =   51
      Top             =   4545
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   4895
      BackColor       =   16761024
      Caption         =   "Glosas Comprobante"
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
      Begin VB.TextBox dato17 
         BackColor       =   &H00E0E0E0&
         Height          =   1575
         Left            =   120
         TabIndex        =   53
         Top             =   360
         Width           =   5535
      End
      Begin VB.CommandButton grabar20 
         Caption         =   "Grabar"
         Height          =   375
         Left            =   1935
         TabIndex        =   52
         Top             =   2205
         Width           =   2175
      End
   End
   Begin XPFrame.FrameXp impuestos 
      Height          =   4245
      Left            =   4005
      TabIndex        =   43
      Top             =   2565
      Width           =   6780
      _ExtentX        =   11959
      _ExtentY        =   7488
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
         Left            =   90
         TabIndex        =   44
         Top             =   315
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
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash opciones 
      Height          =   1455
      Left            =   90
      TabIndex        =   1
      Top             =   6960
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
Attribute VB_Name = "revisaventas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
     Private ivalegal As Double
     Private ACTIVAIMPUESTO As Boolean
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
     Private FECHAMAXIMA As String
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
     Private MES As String
     Private a�o As String
         
Private Sub COMMAND2_Click()

End Sub

Private Sub comboa�o_Click()
a�o = COMBOA�O.text
MES = COMBOMES.ListIndex + 1
If Val(MES) < 10 Then MES = "0" + Mid(Str(MES), 2, 1) Else MES = Mid(Str(MES), 2, 2)

End Sub


Private Sub COMBOMES_Click()
a�o = COMBOA�O.text
MES = COMBOMES.ListIndex + 1
If Val(MES) < 10 Then MES = "0" + Mid(Str(MES), 2, 1) Else MES = Mid(Str(MES), 2, 2)

End Sub

Private Sub CARGAFECHA()
For k = 1 To 12
COMBOMES.AddItem MonthName(k)
Next k
COMBOMES.ListIndex = CDbl(Format(fechasistema, "mm") - 1)
For k = 2000 To Val(Format(fechasistema, "yyyy"))
COMBOA�O.AddItem k
Next k
COMBOA�O.ListIndex = k - 2001

a�o = COMBOA�O.text
MES = COMBOMES.ListIndex + 1
If Val(MES) < 10 Then MES = "0" + Mid(Str(MES), 2, 1) Else MES = Mid(Str(MES), 2, 2)

End Sub
Sub Command1_Click()
borratipos

Dim resultados As rdoResultset
    Dim csql As New rdoQuery
        
        Set csql.ActiveConnection = contadb
        csql.sql = "SELECT '',fc.tipo,numero,fecha,fc.rut,cc.nombre,neto,iva,exento,total,caja "
        csql.sql = csql.sql + "FROM facturasdeventas as fc,cuentascorrientes as cc "
        csql.sql = csql.sql + "where fc.rut=cc.rut and cc.a�o='" + COMBOA�O.text + "' and "
        csql.sql = csql.sql & "(cc.tipo='" + cuentacliente + "' or cc.tipo='" + leerdatos(conta, "maestroempresas", "cuentapublicidad", "codigoempresa='" & empresaactiva & "'") & "') and mid(fecha,1,7)='" + Format(fechasistema, "yyyy-mm") + "' "
        csql.sql = csql.sql & "and caja>'90' "
        csql.sql = csql.sql + "group by tipo,numero,fecha,rut order by tipo,fecha "
        csql.Execute
        If csql.RowsAffected > 0 Then
            BARRA.Max = csql.RowsAffected + 1
            
            BARRA.Value = 0
            Set resultados = csql.OpenResultset
             While Not resultados.EOF
                 If resultados(9) <> 0 Then
                    BARRA.Value = BARRA.Value + 1
                    a�o = Format(resultados(3), "yyyy")
                    MES = Format(resultados(3), "mm")
                    dato1.text = resultados(1)
                    dato2.text = resultados(2)
                    dato9.text = Mid(resultados(4), 1, 9)
                    DV.Caption = Mid(resultados(4), 10, 1)
                    publicidad = False
                    empresarelacionada = False
                    If resultados(10) = "98" Then publicidad = True
                    If resultados(10) = "99" Then empresarelacionada = True
                    
                    If resultados(1) = "8" Then
                        publicidad = espublicidad(resultados(1), resultados(2), resultados(4))
                        If publicidad = True Then
                            empresarelacionada = False
                        End If
                    End If
                    
                    leefactura
                    carga
                    leecomprobante
                    GRABARCOMPROBANTE
                 End If
             resultados.MoveNext
         Wend
          
          resultados.Close
            Set resultados = Nothing

        End If
Unload Me

proceso02.Show



End Sub
Function espublicidad(tipo, numero, rutprove) As Boolean
    Dim csql As New rdoQuery
    Dim resultados As rdoResultset
     If tipo = "8" Then tipo = "NF"
    Set csql.ActiveConnection = contadb
    empresa_fae = leerdatos(conta, "maestroempresas", "empresafae", "codigoempresa='" + empresaactiva + "' ")
    csql.sql = "SELECT dd.tipodocumento,numerodocumento,fechadocumento "
    csql.sql = csql.sql & "FROM " & cliente_sql & "ventas" & empresa_fae & ".sv_otros_documento_cabeza_" & empresa_fae & " AS dc "
    csql.sql = csql.sql & "INNER JOIN " & cliente_sql & "ventas" & empresa_fae & ".sv_otros_documento_detalle_" & empresa_fae & " AS dd ON "
    csql.sql = csql.sql & "dc.tipo = dd.tipo And dc.numero = dd.numero And dc.local = dd.local And dc.fecha = dd.fecha "
    csql.sql = csql.sql & "WHERE dc.foliosii='" & numero & "' AND dc.tipo='" & tipo & "' "
    csql.Execute
    If csql.RowsAffected > 0 Then
        Set resultados = csql.OpenResultset
        espublicidad = leerfacturapubli(rutprove, resultados(1), resultados(2))
    End If
    csql.Close
    Set csql = Nothing
    
    
End Function
Function leerfacturapubli(rutprove, numero, fecha) As Boolean
    Dim csql As New rdoQuery
    Dim resultados As rdoResultset
    Set csql.ActiveConnection = contadb
    csql.sql = " SELECT numero FROM facturasdepublicidad WHERE  fecha='" & Format(fecha, "yyyy-mm-dd") & "' AND rut='" & rutprove & "' AND foliosii='" & numero & "' "
    csql.Execute
        leerfacturapubli = False
    If csql.RowsAffected > 0 Then
        leerfacturapubli = True
    End If
End Function

Private Sub borratipos()
Dim resultados As rdoResultset
    Dim csql As New rdoQuery
        
        Set csql.ActiveConnection = contadb
        csql.sql = "delete from  movimientoscontables where (TIPO='FX' OR TIPO='FA' OR TIPO='ND' OR TIPO='NB' OR TIPO= 'NF' OR TIPO='EF' OR TIPO='ED' OR  TIPO='EC' OR TIPO='EX' OR TIPO='') "
        csql.sql = csql.sql + " and fecha like '" + Format(fechasistema, "yyyy-mm") + "%' "
        
        csql.Execute
        Call sincronizadatos(csql.sql, contadb, "")

        

End Sub

Private Sub dato1_Change()
If Val(dato1.text) < 0 Or Val(dato1.text) > CANDO Then dato1.Enabled = True: dato1.text = "": dato1.SetFocus
End Sub

Private Sub dato1_LostFocus()
    TIPOS.Visible = False
   
End Sub
Private Sub dato1_GotFocus()
CARGAFECHA

Call cargatexto(dato1)
TIPOS.Visible = True
End Sub


Private Sub dato12_LostFocus()
If ivalegal <> CDbl(dato12.text) Then
MsgBox ("IVA DIGITADO NO CORRESPONDE AL IVA REAL DE " & Format(ivalegal, "###,###,###"))

End If

If dato5.text + "-" + dato4.text + "-" + dato3.text < FECHAMAXIMA And CDbl(dato12.text) <> 0 Then
dato12.text = "0"
MsgBox ("FACTURA YA PERDIO DERECHO A CREDITO FISCAL ")

End If

End Sub

Private Sub dato15_GotFocus()
If NETO = 0 Then dato14.SetFocus

Call cargatexto(dato15)

totalfactura

End Sub

Private Sub DATO16_GotFocus()

Call cargatexto(dato16)

End Sub

Private Sub dato16_LostFocus()
If dato16.text <> "N" And dato16.text <> "S" Then dato1.SetFocus

End Sub

Private Sub dato17_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then grabar20.SetFocus
    

End Sub

Private Sub dato2_GotFocus()
CARGAFECHA
leeFOLIO

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
If dato5.text < "1900" Or (dato4.text + dato5.text > Format(fechasistema, "MMYYYY")) Then dato5.text = ""

End Sub

Private Sub dato8_LostFocus()
If dato8.text < "1900" Or dato8.text > Format(fechasistema, "YYYY") Then dato8.text = ""

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
    If sqlconta.status = 0 Then
       carga
       leecomprobante
    
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
If CDbl(dato2.text) = 0 Then dato2.SetFocus

Call cargatexto(dato9)

End Sub

Private Sub dato11_GotFocus()
Rem If IsDate(dato6.text + "-" + dato7.text + "-" + dato8.text) = False Then dato6.text = "": dato7.text = "": dato8.text = "": dato6.SetFocus



Call cargatexto(dato11)

no:
End Sub

Private Sub dato12_GotFocus()
sumador = Int((CDbl(Replace(dato11.text, ",", "")) * iva / 100) + 0.5)
ivalegal = sumador

dato12.text = Format(sumador, "#,###,###,##0")
If dato1.text = "7" Then dato14.text = Format(sumador, "#,###,###,##0")

totalfactura
Call cargatexto(dato12)
End Sub
Private Sub dato13_GotFocus()
totalfactura
Call cargatexto(dato13)


End Sub

Private Sub dato14_GotFocus()
totalfactura
Call cargatexto(dato14)

End Sub
Private Sub dato25_GotFocus()
totalfactura
Call cargatexto(DATO25)

End Sub





Private Sub folios_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Call ceros(folios)
End If

End Sub

Private Sub Form_Activate()
leeCUENTA

End Sub

Private Sub Form_Load()
Dim mesante As Double
Dim A�OANTE As Double

CENTRAR Me
iva = 19
    Call Conectar_BD
    Call Conectarconta(Servidor, clientesistema + "conta", Usuario, password)
    sc = 0
    opciones.Visible = False
GRILLATIPOS
ACTIVAIMPUESTO = False

Call CARGAGRILLA(2, 12)
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

For k = 1 To 12
COMBOMES.AddItem MonthName(k)
Next k
COMBOMES.ListIndex = CDbl(Format(fechasistema, "mm") - 1)
For k = 2000 To Val(Format(fechasistema, "yyyy"))
COMBOA�O.AddItem k
Next k
COMBOA�O.ListIndex = k - 2001

a�o = COMBOA�O.text
MES = COMBOMES.ListIndex + 1
mesante = MES
A�OANTE = a�o

If Val(MES) < 10 Then MES = "0" + Mid(Str(MES), 2, 1) Else MES = Mid(Str(MES), 2, 2)

For k = 1 To 2
mesante = mesante - 1
If mesante = 0 Then mesante = 1: A�OANTE = A�OANTE - 1
Next k

FECHAMAXIMA = "01" & "-" & mesante & "-" & A�OANTE
FECHAMAXIMA = Format(FECHAMAXIMA, "yyyy-mm-dd")
impuestos.Visible = False

glosafinal.Visible = False

CARGAGRILLAexento
folios.Enabled = True
CRCC.Enabled = True
'If Verifica_Permiso(Me.Caption, "todas") = True Then
'Command1.Visible = True
'
'End If


End Sub
Sub leeCUENTA()

    campos(0, 0) = "codigo"
    campos(1, 0) = ""
    
    campos(0, 2) = "cuentasdelmayor"
    condicion = "codigo=" + "'" + CUENTAPROVEEDOR + "' and a�o='" + Format(fechasistema, "yyyy") + "' order by codigo"
    
    op = 5
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)
    If sqlconta.status = 4 Then
    MsgBox ("CUENTA PROVEEDOR NUMERO " + CUENTAPROVEEDOR + " NO EXISTE EN MAESTRO DEL MAYOR CONFIGURE SISTEMA EN MAESTRO EMPRESAS ")
   Unload Me
    
    
    End If
    
   
    
End Sub


Sub GRILLATIPOS()
GRILLATIPO.Cols = 2
GRILLATIPO.Rows = 10
GRILLATIPO.ColWidth(0) = 200 * 2
GRILLATIPO.ColWidth(1) = 220 * 10

GRILLATIPO.TextMatrix(0, 0) = "1"
GRILLATIPO.TextMatrix(1, 0) = "2"
GRILLATIPO.TextMatrix(2, 0) = "3"
GRILLATIPO.TextMatrix(3, 0) = "4"
GRILLATIPO.TextMatrix(4, 0) = "5"
GRILLATIPO.TextMatrix(5, 0) = "6"
GRILLATIPO.TextMatrix(6, 0) = "7"
GRILLATIPO.TextMatrix(7, 0) = "8"
GRILLATIPO.TextMatrix(8, 0) = "9"
GRILLATIPO.TextMatrix(9, 0) = "0"

GRILLATIPO.TextMatrix(0, 1) = "FACTURA"
GRILLATIPO.TextMatrix(1, 1) = "NOTA DE DEBITO"
GRILLATIPO.TextMatrix(2, 1) = "NC. BOLETA    "
GRILLATIPO.TextMatrix(3, 1) = "NC. FACTURA   "
GRILLATIPO.TextMatrix(4, 1) = "FAC.EXPORTACION"
GRILLATIPO.TextMatrix(5, 1) = "FACTURA ELECTRONICA   "
GRILLATIPO.TextMatrix(6, 1) = "ND. ELECTRONICA   "
GRILLATIPO.TextMatrix(7, 1) = "NC. ELECTRONICA "
GRILLATIPO.TextMatrix(8, 1) = "F.EXENTA "
GRILLATIPO.TextMatrix(9, 1) = "F.EXENTA ELECTRONICA "

CANDO = 10



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
    Call flechas(dato13, dato15, KeyCode)
End Sub
Private Sub dato25_KeyDown(KeyCode As Integer, Shift As Integer)
    Call flechas(dato14, dato15, KeyCode)
End Sub

Private Sub dato15_KeyDown(KeyCode As Integer, Shift As Integer)
    Call flechas(DATO25, dato16, KeyCode)
End Sub
Private Sub dato16_KeyDown(KeyCode As Integer, Shift As Integer)
    Call flechas(dato15, dato16, KeyCode)
End Sub



Private Sub dato1_KeyPress(KeyAscii As Integer)
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
    If KeyAscii = 13 Then Call formato(dato11, 0): Call Pregunta(dato11, dato12)
    
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
    If KeyAscii = 13 Then Call formato(dato14, 0): Call Pregunta(dato14, DATO25)

End Sub
Private Sub dato25_KeyPress(KeyAscii As Integer)
    snum = 0: KeyAscii = esNumero(KeyAscii)
    If KeyAscii = 13 Then Call formato(DATO25, 0): Call Pregunta(DATO25, dato15)

End Sub

Private Sub dato15_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    
    If KeyAscii = 13 Then Call Pregunta(dato15, dato16)
 If Chr(KeyAscii) <> "S" And Chr(KeyAscii) <> "N" Then KeyAscii = 0
End Sub
Private Sub dato16_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    
    If KeyAscii = 13 Then detalle.Enabled = True: grid1.Enabled = True: grid1.Rows = 2: grid1.Cell(1, 1).SetFocus
    If Chr(KeyAscii) <> "S" And Chr(KeyAscii) <> "N" Then KeyAscii = 0
    

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




Sub flechas(ByRef caja As TextBox, ByRef caja1 As TextBox, ByRef codigo As Integer)
    If codigo = 38 And caja.Enabled = True Then caja.SetFocus: caja.SelStart = 0
    If codigo = 40 And caja1.Enabled = True Then caja1.SetFocus: caja.SelStart = 0
End Sub
Sub GRABADETALLEIMPUESTOS()
'    Dim j As Double
'    For j = 1 To Grid2.Rows - 2
'
'    If Grid2.Cell(LINEAS, 3).text <> "" Then
'    campos(0, 0) = "tipo"
'    campos(1, 0) = "numero"
'    campos(2, 0) = "rut"
'    campos(3, 0) = "Cuenta"
'    campos(4, 0) = "Monto"
'    campos(5, 0) = ""
'    campos(0, 1) = DATO1.text
'    campos(1, 1) = dato2.text
'    campos(2, 1) = dato9.text + DV.Caption
'    campos(3, 1) = Grid2.Cell(LINEAS, 1).text
'    campos(4, 1) = Grid1.Cell(LINEAS, 3).text
'    campos(5, 1) = ""
'
'    campos(0, 2) = "facturasdecompras_impuestos"
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
    campos(10, 0) = "a�ocontable"
    campos(11, 0) = "mescontable"
    campos(12, 0) = "comentario"
    campos(13, 0) = "electronica"
    campos(14, 0) = "activo"
    campos(15, 0) = "fechadigitacion"
    campos(16, 0) = "folio"
    campos(17, 0) = "impuestoespecifico"
    campos(18, 0) = ""
    
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
    campos(10, 1) = a�o
    campos(11, 1) = MES
    campos(12, 1) = dato17.text
    campos(13, 1) = dato15.text
    campos(14, 1) = dato16.text
    campos(15, 1) = Format(Date, "yyyy") + "-" + Format(Date, "mm") + "-" + Format(Date, "dd")
    campos(16, 1) = folios.text
    campos(17, 1) = Replace(DATO25.text, ".", "")
    
    
    
    condicion = ""
    campos(0, 2) = "facturasdecompras"
    op = 2
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)
    
GRABADETALLEIMPUESTOS
grabardetallefactura


grabar2

End Sub


Sub grabar2()
    leecomprobante
    opciones.Visible = True
    opciones.SetFocus
    detalle.Enabled = False
    If ivacredito <> 99999999 Then
        GRABARCOMPROBANTE
    End If

End Sub
Sub ELIMINAR()
    Dim TIPOCON As String
     
    
    Call ACTUALIZADOCUMENTO("-")

  
    
    campos(0, 2) = "facturasdecompras_impuestos"
    condicion = "tipo=" + "'" + dato1.text + "'" + " and numero=" + "'" + dato2.text + "'" + " and rut=" + "'" + dato9.text + DV.Caption + "'"
    op = 4
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)
    
    campos(0, 2) = "facturasdecompras"
    condicion = "tipo=" + "'" + dato1.text + "'" + " and numero=" + "'" + dato2.text + "'" + " and rut=" + "'" + dato9.text + DV.Caption + "'"
    op = 4
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)
    campos(0, 2) = "movimientoscontables"
    If dato1.text = "9" Then TIPOCON = "FX"
    If dato1.text = "1" Then TIPOCON = "FA"
    If dato1.text = "2" Then TIPOCON = "ND"
    If dato1.text = "3" Then TIPOCON = "NB"
    If dato1.text = "4" Then TIPOCON = "NF"
    If dato1.text = "5" Then TIPOCON = "FE"
    
    condicion = "tipo=" + "'" + TIPOCON + "'" + " and numero=" + "'" + dato2.text + "' and rutproveedor='" + dato9.text + DV.Caption + "'"
    
    op = 4
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)
  
    campos(0, 2) = "facturasdecompras_detalle"
    condicion = "tipo=" + "'" + dato1.text + "'" + " and numero=" + "'" + dato2.text + "'" + " and rut=" + "'" + dato9.text + DV.Caption + "'"
    op = 4
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)
    'If sqlconta.status = 4 Then Stop
no:

End Sub


Private Sub glosafactura_Change()

End Sub



Private Sub MSHFlexGrid1_Click()

End Sub


Private Sub Form_Unload(Cancel As Integer)
If MODIFI = 1 Then
   grabafactura
    retorno
End If
End Sub

Private Sub grabar20_Click()
If Verifica_Permiso(Me.Caption, "agrega") = True Then

grabafactura
glosafinal.Visible = False
Else
  MsgBox mensaje_nopermiso, vbCritical + vbOKOnly, "Permiso Denegado"
End If

End Sub

Private Sub Grid1_GotFocus()
glosafinal.Visible = False

Rem If dato3.text + dato4.text <> Format(fechasistema, "mm") + Format(fechasistema, "yyyy") Then dato2.text = "": dato3.text = "": dato4.text = "": dato2.SetFocus

End Sub

Private Sub Grid1_KeyDown(KeyCode As Integer, ByVal Shift As Integer)
    If KeyCode = 35 And grid1.ActiveCell.col = 1 And Val(saldo.Caption) = 0 And grid1.ActiveCell.row <> 1 Then grid1.Cell(grid1.ActiveCell.row, grid1.ActiveCell.col).text = "": datosglosa
    Rem If KeyCode = 38 And Grid1.ActiveCell.row = Grid1.Rows - 1 Then SG = "S" Else SG = "N"
    If grid1.ActiveCell.col = "1" And KeyCode = vbKeyF2 Then Call ayudamayor(grid1.ActiveCell.row, grid1.ActiveCell.col)
    End Sub

Private Sub Grid1_KeyPress(KeyAscii As Integer)
    Dim cadena As String
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    'If Grid1.ActiveCell.Col = 11 And Chr(KeyAscii) <> "D" And Chr(KeyAscii) <> "H" Then KeyAscii = 0
    If grid1.ActiveCell.col = 1 And Chr(KeyAscii) = "*" And Val(saldo.Caption) = 0 And grid1.ActiveCell.row <> 1 Then datosglosa
    If grid1.ActiveCell.col = 1 And Chr(KeyAscii) = "*" And Val(saldo.Caption) <> 0 And grid1.ActiveCell.row <> 1 Then MsgBox ("comprobante descuadrado"): grid1.Cell(grid1.ActiveCell.row, 1).SetFocus
    
    
    Rem If formatogrilla(3, Grid1.ActiveCell.col) = "S" Then Grid1.Cell(Grid1.ActiveCell.row, Grid1.ActiveCell.col).text = UCase(Grid1.ActiveCell.text)
    If FORMATOGRILLA(3, grid1.ActiveCell.col) = "N" Then snum = 1: KeyAscii = esNumero(KeyAscii)
    If FORMATOGRILLA(3, grid1.ActiveCell.col) = "C" Then snum = 1: KeyAscii = esNumero(KeyAscii)
    If grid1.ActiveCell.col = grid1.Cols - 1 Then
        If KeyAscii <> 68 And KeyAscii <> 72 And KeyAscii <> 8 Then
           KeyAscii = 0
       End If
    End If
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
If command = "retorno" And MODIFI = 0 Then retorno
If command = "retorno" And MODIFI = 1 Then
    grabafactura
    retorno
End If

If command = "modifica" Then
If Verifica_Permiso(Me.Caption, "modifica") = True Then
    MODIFI = 1
    ELIMINAR
    dato2.Enabled = True
    dato2.SetFocus
    Else
    MsgBox mensaje_nopermiso, vbCritical + vbOKOnly, "Permiso Denegado"
    End If
    
End If
If command = "elimina" Then
If Verifica_Permiso(Me.Caption, "elimina") = True Then
 If MsgBox("DESEA REALMENTE ELIMINAR ", vbYesNo) = vbYes Then
    ELIMINAR
    retorno
End If
Else
  MsgBox mensaje_nopermiso, vbCritical + vbOKOnly, "Permiso Denegado"
  End If
  
End If
End Sub


Sub retorno()



grid1.Rows = 1


opciones.Visible = False
limpia
disponible (False)
MODIFI = 0
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
    DATO25.text = "0"
    
    dato15.text = "N"
    total.text = "0"
    dato16.text = "N"
    dato17.text = ""
    
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
    cfijo = "tipo='" & CUENTAPROVEEDOR & "' and a�o='" + a�o + "'"
    cabezas = Array("RUT", "NOMBRE")
    mensajeAyuda = "Ayuda Cuentas Corrientes"
    pivote2.MaxLength = 10
    
    Call cargaAyudaT(Servidor, basebus, Usuario, password, "cuentascorrientes", pivote2, campos, cfijo, largo, 2)
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
    
    condicion = "codigo=" + "'" + cuenta + "' and a�o='" + a�o + "'"
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
Sub leectacte()
    campos(0, 0) = "rut"
    campos(1, 0) = "nombre"
    campos(2, 0) = ""
    campos(0, 2) = "cuentascorrientes"
    condicion = "tipo=" + "'" + cuentacliente + "' and rut=" + "'" + pivote2.text + "' and a�o='" + a�o + "'"
    op = 5
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)
    
    nombreproveedor.Caption = sqlconta.response(1, 3)
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
sumador = CDbl(Replace(dato11.text, ",", "")) + CDbl(Replace(dato12.text, ",", "")) + CDbl(Replace(dato13.text, ",", "") + CDbl(Replace(DATO25.text, ",", "")) - CDbl(Replace(dato14.text, ",", "")))
total.text = Format(sumador, "###,###,###,##0")
NETO = CDbl(Replace(dato11.text, ",", "")) + CDbl(Replace(dato13.text, ",", "")) + CDbl(Replace(DATO25.text, ",", ""))
haber.Caption = Format(NETO, "###,###,##0")
End Sub



Sub grabardetallefactura()
    Dim TIPOCON As String
    Dim tipo2 As String
    Dim j As Integer
    Dim lin As Integer
    Dim monto As Double
    
    
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
    campos(10, 0) = "fechacreacion"
    campos(11, 0) = ""
    
    
Rem graba detalle factura
    lin = 0
    For j = 1 To grid1.Rows - 2
    lin = lin + 1
    LINEAS.text = lin
    Call ceros(LINEAS)
    campos(0, 1) = dato1.text
    campos(1, 1) = dato2.text
    campos(2, 1) = LINEAS.text
    campos(3, 1) = dato9.text + DV.Caption
    campos(4, 1) = grid1.Cell(j, 1).text + grid1.Cell(j, 2).text + grid1.Cell(j, 3).text
    campos(5, 1) = grid1.Cell(j, 4).text
    campos(6, 1) = grid1.Cell(j, 5).text
    campos(7, 1) = grid1.Cell(j, 6).text
    campos(8, 1) = grid1.Cell(j, 11).text
    campos(9, 1) = grid1.Cell(j, 10).text
    campos(10, 1) = Format(fechasistema, "yyyy-mm-dd")
    campos(0, 2) = "facturasdecompras_detalle"
    condicion = ""
    op% = 2
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)
    
    Next j

    
    
End Sub





Sub ELIMINA()
Call ACTUALIZADOCUMENTO("-")
End Sub









Sub ayudamayor(row As Long, col As Long)
    Dim campos As Variant
    Dim cfijo As Variant
    Dim largo As Variant
    campos = Array("codigo", "nombre")
    cabezas = Array("codigo", "nombre")
    largo = Array("8n", "40s")
    mensajeAyuda = "Ayuda Cuentas del Mayor"
    cfijo = "no"
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
    dato15.Enabled = condicion
    DATO25.Enabled = condicion
    
    total.Enabled = condicion
    dato16.Enabled = condicion
    
    
End Sub


Sub ACTUALIZADOCUMENTO(COMANDO As String)
    Dim lin As Integer
  
    Dim resultados As rdoResultset
    Dim csql As New rdoQuery
    Dim rut As String
    Dim TIPOFA As String
  
        If dato1.text = "1" Then TIPOFA = "FC"
        If dato1.text = "2" Then TIPOFA = "DC"
        If dato1.text = "3" Then TIPOFA = "NC"
        If dato1.text = "4" Then TIPOFA = "FC"
        If dato1.text = "5" Then TIPOFA = "DC"
        If dato1.text = "6" Then TIPOFA = "NC"
        If dato1.text = "7" Then TIPOFA = "FP"
        
        Set csql.ActiveConnection = contadb
        csql.sql = "SELECT tipo,numero,linea,fecha,codigocuenta,tipoctacte,rutctacte,centrocosto,glosacontable,tipodocumento,numerodocumento,fechavencimiento,monto,dh "
        csql.sql = csql.sql + "FROM movimientoscontables "
        csql.sql = csql.sql + "WHERE tipo='" + TIPOFA + "' and numero='" & dato2.text & "'and a�o='" + a�o + "' and mes='" + MES + "' order by linea"
        csql.Execute


        If csql.RowsAffected > 0 Then
            Set resultados = csql.OpenResultset
            While Not resultados.EOF
                
                Call actualizamayor(COMANDO, resultados(4), resultados(12), resultados(13), resultados(5), resultados(6), resultados(7), MES, a�o)
                
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
    If MODIFI = 0 Then
    campos(0, 0) = "folio"
    campos(1, 0) = ""
    campos(0, 2) = "facturasdecompras"
    condicion = "mescontable = '" & MES & "' AND a�ocontable = '" & a�o & "' ORDER BY folio DESC "
    op = 5
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)
    
    If sqlconta.status = 0 Then
        k = Val(sqlconta.response(0, 3))
    Else
        k = 0
    End If
    folios.text = k + 1
    
    Call ceros(folios)
End If
End Sub

Sub leecuentas()

Dim resultados2 As rdoResultset
    Dim csql2 As New rdoQuery
    Dim rut As String
    Dim PASO As String
    Dim LINEAS As Double
    
       
        
        Set csql2.ActiveConnection = contadb
        csql2.sql = "SELECT codigo,nombre "
        csql2.sql = csql2.sql + "FROM cuentasdelmayor where ila<>'0' or iha<>'0' or ica<>'0' and a�o='" + Format(fechasistema, "yyyy") + "' "
       
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

Private Sub Grid1_LeaveCell(ByVal row As Long, ByVal col As Long, NewRow As Long, NewCol As Long, Cancel As Boolean)
    Dim TEXTO As String
    Dim DV As String

    If row = 0 And col = 0 Then NewRow = 1: NewCol = 1: GoTo no:
    If NewCol = 7 And row = grid1.Rows - 1 Then grid1.Rows = grid1.Rows + 1: NewRow = grid1.Rows - 1
    If NewCol = 7 And row < grid1.Rows - 1 Then NewCol = 1: NewRow = row + 1
    If NewCol > 6 Then NewCol = 6
    TEXTO = grid1.Cell(row, col).text
    For k = 1 To 7
        If col = k And row > 1 And grid1.Cell(row, col).text = "" And col < NewCol Then
            grid1.Cell(row, col).text = grid1.Cell(row - 1, col).text
        End If
    Next k

    lin = row
   'If Col = 1 And Row = Grid1.Rows - 1 And NewRow < Row Then GoTo paso2:

    If col = 1 Then
        PIVOTE.MaxLength = 2: PIVOTE.text = grid1.Cell(row, col).text: Call ceros(PIVOTE): grid1.Cell(row, col).text = PIVOTE.text
        If grid1.Cell(row, col).text = "00" And NewCol > col Then
        grid1.Cell(row, col).text = ""
        
        col = 1: NewCol = 1
        
        End If
    End If

    If col = 2 Then
        PIVOTE.MaxLength = 2: PIVOTE.text = grid1.Cell(row, col).text: Call ceros(PIVOTE): grid1.Cell(row, col).text = PIVOTE.text
        If grid1.Cell(row, col).text = "00" And NewCol > col Then
        grid1.Cell(row, col).text = ""
        
        col = 2: NewCol = 2
        End If
    End If
    If col = 3 Then
        PIVOTE.MaxLength = 4: PIVOTE.text = grid1.Cell(row, col).text: Call ceros(PIVOTE): grid1.Cell(row, col).text = PIVOTE.text
        If grid1.Cell(row, col).text = "0000" And NewCol > col Then
        grid1.Cell(row, col).text = ""
        
        col = 3: NewCol = 4
        End If
    End If

    If NewCol = 4 Or (NewRow <> row And col < 4) And row < grid1.Rows - 1 Then
    
    Call leermayor(row, col)
            If respuesta = "N" Then
            NewCol = col
            NewRow = row
            
            End If
    End If
    

    If (col = 3 And NewCol = 4) And row = 1 Then
    
    Call leermayor(row, col)
            If respuesta = "N" Then
            NewCol = col
            NewRow = row
                        End If
    End If
    

     grid1.Cell(row, 6).text = "D"
    If col = 6 Then SUMAR
    If NewCol = 5 And Val(grid1.Cell(row, 5).text) = 0 Then grid1.Cell(row, 5).text = NETO
10:


   If NewRow = row And NewCol > col Then
        If grid1.Cell(row, 6).text <> "D" And grid1.Cell(row, 6).text <> "H" Then NewCol = 6: NewRow = row
        If grid1.Cell(row, 5).text = "" Then NewCol = 5: NewRow = row
        If grid1.Cell(row, 4).text = "" Then NewCol = 4: NewRow = row
        If Val(grid1.Cell(row, 3).text) = 0 Then NewCol = 3: NewRow = row
        If Val(grid1.Cell(row, 2).text) = 0 Then NewCol = 2: NewRow = row
        If Val(grid1.Cell(row, 1).text) = 0 Then NewCol = 1: NewRow = row
   End If
   Rem cuando cae

   If NewRow = grid1.Rows - 1 And col < NewCol Then
        If grid1.Cell(NewRow, 6).text <> "D" And grid1.Cell(NewRow, 6).text <> "H" Then NewCol = 6
        If grid1.Cell(NewRow, 5).text = "" Then NewCol = 5
        If grid1.Cell(NewRow, 4).text = "" Then NewCol = 4
        If Val(grid1.Cell(NewRow, 3).text) = 0 Then NewCol = 3
        If Val(grid1.Cell(NewRow, 2).text) = 0 Then NewCol = 2
        If Val(grid1.Cell(NewRow, 1).text) = 0 Then NewCol = 1
   End If
   
   
   
If NewRow <> row And grid1.Rows - 1 <> row Then
    If grid1.Cell(row, col).text = "" Then
    NewRow = row: col = NewCol
    End If
End If
If NewRow <> row And grid1.Rows - 1 <> row Then
    If grid1.Cell(row, 6).text <> "D" And grid1.Cell(row, 6).text <> "H" Then NewCol = 6: NewRow = row
    
End If
    If NewRow = grid1.Rows - 1 And grid1.Rows > 2 And row < NewRow Then NewCol = 1
    

no:
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
    condicion = "codigo=" + "'" + grid1.Cell(row, 1).text + grid1.Cell(row, 2).text + grid1.Cell(row, 3).text + "' and a�o='" + Format(fechasistema, "yyyy") + "'"
    
    op = 5
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)
 
    If sqlconta.status = 4 Or grid1.Cell(row, 3).text = "0000" Then
    grid1.Cell(row, 1).text = ""
    grid1.Cell(row, 2).text = ""
    grid1.Cell(row, 3).text = ""
    grid1.Cell(row, 1).SetFocus
    
    respuesta = "N"
    Else
    respuesta = "S"
    grid1.Cell(row, 7).text = sqlconta.response(1, 3)
    If grid1.Cell(row, 4).text = "" Or MODIFI = 1 Then
    grid1.Cell(row, 4).text = sqlconta.response(1, 3)
    End If
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
     If TIENECTACTE(row) = "1" Then
     grid1.Cell(row, 10).text = dato9.text + DV.Caption
     
     End If
     

End Sub


    

Sub leercrcc(row As Long, col As Long)
    campos(0, 0) = "codigo"
    campos(1, 0) = "nombre"
    campos(2, 0) = ""
    campos(0, 2) = "centrosdecosto"
    condicion = "codigo=" + "'" + DATO21.text + DATO22.text + "' and a�o='" + Format(fechasistema, "yyyy") + "'"
    op = 5
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)
    If sqlconta.status = 4 Or DATO22.text = "00" Then DATO21.text = "": DATO22.text = "": DATO21.SetFocus: GoTo no:
    grid1.Cell(row, 9).text = sqlconta.response(1, 3)
    grid1.Cell(row, 11).text = sqlconta.response(0, 3)
    nombrecrcc.Caption = sqlconta.response(1, 3)
    If col <> 9999 Then
    DATO21.text = ""
    DATO22.text = ""
    
    cabeza.Enabled = True
    detalle.Enabled = True
    CRCC.Enabled = False
    
    grid1.Cell(row, 5).SetFocus
  End If

no:
End Sub
Sub SUMAR()
Dim o As Integer
Dim sumadebe As Double
Dim sumahaber As Double

sumadebe = 0
sumahaber = NETO
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
    FORMATOGRILLA(4, 6) = "D"
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
    grid1.Column(0).Width = 4 * 8.8
    grid1.Column(1).Width = 2 * 10
    grid1.Column(2).Width = 2 * 10
    grid1.Column(3).Width = 4 * 10
    grid1.Column(4).Width = 40 * 9
    grid1.Column(5).Width = 12 * 9
    grid1.Column(6).Width = 3 * 9
    grid1.Column(7).Width = 100
    grid1.Column(8).Width = 40
    grid1.Column(9).Width = 100
    grid1.Column(10).Width = 40
    
    
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
            grid1.Column(k).Mask = cellUpper
        End If
        If FORMATOGRILLA(3, k) = "D" Then
            grid1.Column(k).CellType = cellCalendar
            grid1.Column(k).Mask = cellNumeric
        End If
        
        'Grid1.Column(7).CellType = cellComboBox
    Next k
    grid1.Range(0, 1, 0, 3).Merge
    grid1.Cell(0, 1).text = "CUENTA"
    grid1.Range(0, 0, 0, grid1.Cols - 1).Alignment = cellCenterCenter

End Sub


Private Sub DATO21_GotFocus()
DATO21.text = Mid(grid1.Cell(lin, 11).text, 1, 2)
DATO22.text = Mid(grid1.Cell(lin, 11).text, 3, 4)

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
    Call leercrcc(DATO21.Tag, 12)
    End If
End Sub


Sub cargadorcomprobante()
    Dim LINEA As Long
    grid1.AutoRedraw = False
    
    grid1.Rows = 1
    grid1.Rows = canli + 2
    
    For k = 1 To canli
    CUENTAMAYOR(k) = grilladetalle(k, 1)
    grid1.Cell(k, 1).text = grilladetalle(k, 1)
    grid1.Cell(k, 2).text = grilladetalle(k, 2)
    grid1.Cell(k, 3).text = grilladetalle(k, 3)
    grid1.Cell(k, 4).text = grilladetalle(k, 4)
    grid1.Cell(k, 5).text = grilladetalle(k, 5)
    grid1.Cell(k, 6).text = grilladetalle(k, 6)
    grid1.Cell(k, 7).text = ""
    grid1.Cell(k, 8).text = ""
    grid1.Cell(k, 9).text = ""
    
    grid1.Cell(k, 10).text = grilladetalle(k, 10)
    grid1.Cell(k, 11).text = grilladetalle(k, 11)
    
    DATO21.text = Mid(grilladetalle(k, 11), 1, 2)
    DATO22.text = Mid(grilladetalle(k, 11), 3, 2)
    
    LINEA = k
    
    Call leermayor(LINEA, 9999)
    
    If Val(grilladetalle(LINEA, 11)) <> 0 Then Call leercrcc(LINEA, 9999)

    SUMAR

    Next k
    grid1.AutoRedraw = True
    grid1.Refresh
    
    
End Sub
                

Sub ayudacrcc(row As Long, col As Long)
    Dim campos As Variant
    Dim cfijo As Variant
    Dim largo As Variant
    campos = Array("codigo", "nombre")
    cabezas = Array("codigo", "nombre")
    largo = Array("8n", "40s")
    mensajeAyuda = "Ayuda Centros de costo"
    cfijo = "a�o='" + a�o + "'"
    pivote2.MaxLength = 4
    Call cargaAyudaT(Servidor, basebus, Usuario, password, "centrosdecosto", pivote2, campos, cfijo, largo, 2)
    DATO21.text = Mid(pivote2.text, 1, 2)
    DATO22.text = Mid(pivote2.text, 3, 2)
    
    pivote2.text = ""
End Sub

Sub datosglosa()
glosafinal.Visible = True
dato17.SetFocus


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
    campos(10, 0) = ""
    campos(0, 2) = "facturasdeventas"
    condicion = "tipo=" + "'" + dato1.text + "'" + " and numero=" + "'" + dato2.text + "'"

    op = 5
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)
    
    Rem If sqlconta.status = 0 Then modifi = 1: carga: opciones.Visible = True: disponible (True): habilita (True): opciones.SetFocus


End Sub

Sub leervarios()
    Dim csql As New rdoQuery
    Dim tipoconta As String
    
    Dim resultados As rdoResultset
    Set csql.ActiveConnection = contadb
    
     If dato1.text = 6 Then ' factura
        tipoconta = 33
    End If
    
    If dato1.text = 0 Then  ' factura extenta
        tipoconta = 34
    End If
    
    If dato1.text = 7 Then ' nota debito
        tipoconta = 56
    End If
    
    If dato1.text = 8 Then  ' nota credito
        tipoconta = 61
    End If
    
    
    
    csql.sql = "select tipofactura "
    csql.sql = csql.sql & "from facturasvarias "
    csql.sql = csql.sql & "where tipo=" + "'" + tipoconta + "'" + " and numero=" + "'" + dato2.text + "'"
    csql.Execute
    If csql.RowsAffected > 0 Then
        Set resultados = csql.OpenResultset
        If resultados(0) = "1" Then
           Option1.Value = True
        End If
        If resultados(0) = "2" Then
             Option2.Value = True
        End If
    End If

End Sub

Sub carga()
    disponible (True)
    
    dato1.text = sqlconta.response(0, 3)
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
   
    totalfactura
        DV.Caption = rut(dato9.text)
    pivote2.text = dato9.text + DV.Caption
    
        leectacte
fin:
End Sub
Sub leecomprobante()
    Dim lin As Integer
    Dim TIPOFA As String
    Dim resultados As rdoResultset
    Dim csql As New rdoQuery
    Dim rut As String
     
    
        Set csql.ActiveConnection = contadb
        csql.sql = "SELECT tipo,numero,linea,cuentadelmayor,glosa,monto,dh,rutctacte,centrodecosto "
        csql.sql = csql.sql + "FROM facturasdeventas_detalle "
        csql.sql = csql.sql + "WHERE tipo='" + dato1.text + "' and numero='" & dato2.text & "' order by linea"
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

Sub GRABARCOMPROBANTE()
    Dim TIPOCON As String
    Dim tipo2 As String
    Dim j As Integer
    Dim lin As Integer
    Dim HD1 As String
    Dim HD2 As String
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
    campos(17, 0) = "a�o"
    campos(18, 0) = "fechacreacion"
    campos(19, 0) = "horacreacion"
    campos(20, 0) = ""
    Rem cuenta proveedores
    If dato1.text = "9" Then TIPOCON = "FX": HD1 = "D": HD2 = "H"
    If dato1.text = "1" Then TIPOCON = "FA": HD1 = "D": HD2 = "H"
    If dato1.text = "2" Then TIPOCON = "ND": HD1 = "D": HD2 = "H"
    If dato1.text = "3" Then TIPOCON = "NB": HD1 = "H": HD2 = "D"
    If dato1.text = "4" Then TIPOCON = "NF": HD1 = "H": HD2 = "D"
    If dato1.text = "6" Then TIPOCON = "EF": HD1 = "D": HD2 = "H"
    If dato1.text = "7" Then TIPOCON = "ED": HD1 = "D": HD2 = "H"
    If dato1.text = "8" Then TIPOCON = "EC": HD1 = "H": HD2 = "D"
    If dato1.text = "0" Then TIPOCON = "EX": HD1 = "D": HD2 = "H"
    
'      If empresaactiva = "29" Then
'        If Option1.Value = True Then
'           cuentafinal = "11200029"
'        End If
'
'        If Option2.Value = True Then
'            cuentafinal = "11200028"
'        End If
'    Else
'        cuentafinal = cuentacliente
'    End If
'
    
    
    campos(0, 1) = TIPOCON
    campos(1, 1) = dato2.text
    campos(2, 1) = "001"
    campos(3, 1) = dato5.text + "-" + dato4.text + "-" + dato3.text
    campos(4, 1) = cuentacliente
    
    If publicidad = True Then
        campos(4, 1) = leerdatos(conta, "maestroempresas", "cuentapublicidad", "codigoempresa='" + empresaactiva + "' ")
    End If
    If empresarelacionada = True Then
        campos(4, 1) = leerdatos(conta, "maestroempresas", "cuentacreditoer", "codigoempresa='" + empresaactiva + "' ")
    End If
        
        
    
    
    campos(5, 1) = tipocuenta
    campos(6, 1) = dato9.text + DV.Caption
    campos(7, 1) = ""
    campos(8, 1) = "CONTABILIZACION " + DOCU$(Val(dato1.text)) + " " + nombreproveedor.Caption
    campos(9, 1) = DOCU$(Val(dato1.text))
    campos(10, 1) = dato2.text
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
    campos(1, 1) = dato2.text
    campos(2, 1) = "002"
    campos(3, 1) = dato5.text + "-" + dato4.text + "-" + dato3.text
    campos(4, 1) = ivadebito
    campos(5, 1) = ""
    campos(6, 1) = ""
    campos(7, 1) = ""
    campos(8, 1) = "CONTABILIZACION I.V.A " + DOCU(Val(dato1.text))
    campos(9, 1) = DOCU(Val(dato1.text))
    campos(10, 1) = dato2.text
    campos(11, 1) = dato5.text + "-" + dato4.text + "-" + dato3.text
    campos(12, 1) = dato5.text + "-" + dato4.text + "-" + dato3.text
    campos(13, 1) = Replace(dato12.text, ".", "")
    
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
Rem graba impuestos
    lin = 2
    For j = 1 To Grid2.Rows - 1
   
    If Val(Grid2.Cell(j, 3).text) <> 0 Then
    lin = lin + 1
    LINEAS.text = lin
    Call ceros(LINEAS)
    
    campos(0, 1) = TIPOCON
    campos(1, 1) = dato2.text
    campos(2, 1) = LINEAS.text
    campos(3, 1) = dato5.text + "-" + dato4.text + "-" + dato3.text
    campos(4, 1) = Grid2.Cell(j, 1).text
    campos(5, 1) = ""
    campos(6, 1) = ""
    campos(7, 1) = ""
    campos(8, 1) = Grid2.Cell(j, 2).text
    campos(9, 1) = DOCU(Val(dato1.text))
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

    
    
    For k = 1 To grid1.Rows - 2
    lin = lin + 1
    LINEAS.text = lin
    Call ceros(LINEAS)
    
    campos(0, 1) = TIPOCON
    campos(1, 1) = dato2.text
    campos(2, 1) = LINEAS.text
    campos(3, 1) = dato5.text + "-" + dato4.text + "-" + dato3.text
    campos(4, 1) = grid1.Cell(k, 1).text + grid1.Cell(k, 2).text + grid1.Cell(k, 3).text
    campos(5, 1) = ""
    campos(6, 1) = grid1.Cell(k, 10).text
    campos(7, 1) = grid1.Cell(k, 11).text
    campos(8, 1) = grid1.Cell(k, 4).text
    campos(9, 1) = DOCU(Val(dato1.text))
    campos(10, 1) = dato2.text
    campos(11, 1) = dato5.text + "-" + dato4.text + "-" + dato3.text
    campos(12, 1) = dato8.text + "-" + dato7.text + "-" + dato6.text
    campos(13, 1) = grid1.Cell(k, 5).text
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
    
'     If Val(Replace(dato13.text, ".", "")) > 0 And (dato1.text = "1" Or dato1.text = "6") Then
'        lin = lin + 1
'        LINEAS.text = lin
'        Call ceros(LINEAS)
'
'        campos(0, 1) = TIPOCON
'        campos(1, 1) = numero
'        campos(2, 1) = lin
'        campos(3, 1) = dato5.text + "-" + dato4.text + "-" + dato3.text
'        campos(4, 1) = "35100002"
'        campos(5, 1) = ""
'        campos(6, 1) = ""
'        campos(7, 1) = ""
'        campos(8, 1) = "CONTABILIZACION EXENTO " + DOCU(Val(tipo))
'        campos(9, 1) = DOCU(Val(tipo))
'        campos(10, 1) = numero
'        campos(11, 1) = dato5.text + "-" + dato4.text + "-" + dato3.text
'        campos(12, 1) = dato5.text + "-" + dato4.text + "-" + dato3.text
'        campos(13, 1) = Replace(dato13.text, ".", "")
'
'        campos(14, 1) = HD2
'        campos(15, 1) = USUARIOSISTEMA
'        campos(16, 1) = MES
'        campos(17, 1) = a�o
'        campos(0, 2) = "movimientoscontables"
'        condicion = ""
'        op = 2
'        sqlconta.response = campos
'        Set sqlconta.conexion = contadb
'        Call sqlconta.sqlconta(op, condicion)
'    End If
'
    Rem Call ACTUALIZADOCUMENTO("+")
   
End Sub

Private Sub botonmisaccesos_Click()
programafiltro = Me.Caption
misaccesos.Show
End Sub


Private Sub botonmisfavoritos_Click()
Call AgregaFavoritos(USUARIOSISTEMA, App.EXEName, Me.Name, Me.Caption)
End Sub
