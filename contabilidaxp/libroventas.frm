VERSION 5.00
Object = "{4777436C-EB8C-4596-98A8-EBCF98683969}#1.0#0"; "FlexCell.ocx"
Object = "{085FBDF3-00C5-421B-B762-1D57299A2B89}#1.0#0"; "CLBUTN.OCX"
Object = "{ADBBAED6-B16F-40DC-80DF-B44910CBA76C}#1.0#0"; "Frame-Xp.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form auxiliar44 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "gril"
   ClientHeight    =   11580
   ClientLeft      =   435
   ClientTop       =   825
   ClientWidth     =   8175
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   11580
   ScaleWidth      =   8175
   Begin XPFrame.FrameXp FrameXp11 
      Height          =   1215
      Left            =   240
      TabIndex        =   42
      Top             =   10080
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   2143
      BackColor       =   49344
      CaptionEstilo3D =   1
      BackColor       =   49344
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
         Height          =   975
         Left            =   0
         TabIndex        =   43
         Top             =   240
         Width           =   7455
         _ExtentX        =   13150
         _ExtentY        =   1720
         Cols            =   5
         DefaultFontSize =   8.25
         Rows            =   30
      End
   End
   Begin XPFrame.FrameXp fechas 
      Height          =   1620
      Left            =   1320
      TabIndex        =   20
      Top             =   8400
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   2858
      BackColor       =   14737632
      Caption         =   "Rangos de Fecha"
      CaptionEstilo3D =   1
      BackColor       =   14737632
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
      Begin CoolButtons.cool_Button command8 
         Height          =   375
         Left            =   1800
         TabIndex        =   21
         Top             =   1125
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         SkinId          =   "13"
         Caption         =   "Cambia Fecha"
      End
      Begin VB.Label hastafecha 
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   2520
         TabIndex        =   25
         Top             =   720
         Width           =   1935
      End
      Begin VB.Label desdefecha 
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   360
         TabIndex        =   24
         Top             =   720
         Width           =   1935
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Hasta Fecha"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   2520
         TabIndex        =   23
         Top             =   360
         Width           =   1935
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Desde Fecha"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   360
         TabIndex        =   22
         Top             =   360
         Width           =   1935
      End
   End
   Begin XPFrame.FrameXp OPCIONES 
      Height          =   8175
      Left            =   90
      TabIndex        =   0
      Top             =   120
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   14420
      BackColor       =   16761024
      Caption         =   "Lista Libro de Ventas"
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
      Begin VB.TextBox txtcodigoautorizacion 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   3120
         MaxLength       =   10
         TabIndex        =   48
         Top             =   7320
         Width           =   2535
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Formato Para SII XML"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   375
         Left            =   3480
         TabIndex        =   41
         Top             =   6480
         Value           =   1  'Checked
         Width           =   3975
      End
      Begin CoolButtons.cool_Button COMMAND2 
         Height          =   360
         Left            =   4200
         TabIndex        =   12
         Top             =   6000
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   635
         Caption         =   "Genera Informe"
      End
      Begin XPFrame.FrameXp FrameXp2 
         Height          =   1095
         Left            =   240
         TabIndex        =   2
         Top             =   360
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   1931
         BackColor       =   16761024
         Caption         =   "Datos"
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
         Begin VB.OptionButton datos2 
            BackColor       =   &H00FFC0C0&
            Caption         =   "Rango Fecha"
            Height          =   375
            Left            =   360
            TabIndex        =   19
            Top             =   600
            Width           =   2055
         End
         Begin VB.OptionButton datos1 
            BackColor       =   &H00FFC0C0&
            Caption         =   "Mensual"
            Height          =   375
            Left            =   360
            TabIndex        =   18
            Top             =   240
            Width           =   2055
         End
      End
      Begin MSComctlLib.ProgressBar barra 
         Height          =   255
         Left            =   90
         TabIndex        =   1
         Top             =   6885
         Width           =   7575
         _ExtentX        =   13361
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
      End
      Begin XPFrame.FrameXp FrameXp3 
         Height          =   1095
         Left            =   240
         TabIndex        =   3
         Top             =   1560
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   1931
         BackColor       =   16761024
         Caption         =   "Resumen"
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
         Begin VB.OptionButton RESUMEN2 
            BackColor       =   &H00FFC0C0&
            Caption         =   "Resumido"
            Height          =   375
            Left            =   480
            TabIndex        =   5
            Top             =   600
            Width           =   2055
         End
         Begin VB.OptionButton RESUMEN1 
            BackColor       =   &H00FFC0C0&
            Caption         =   "Detallado"
            Height          =   375
            Left            =   480
            TabIndex        =   4
            Top             =   240
            Width           =   2055
         End
      End
      Begin XPFrame.FrameXp FrameXp4 
         Height          =   4215
         Left            =   3240
         TabIndex        =   6
         Top             =   360
         Width           =   4575
         _ExtentX        =   8070
         _ExtentY        =   7435
         BackColor       =   16761024
         Caption         =   "Configuracion"
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
         Begin XPFrame.FrameXp FrameXp5 
            Height          =   855
            Left            =   120
            TabIndex        =   7
            Top             =   360
            Width           =   4335
            _ExtentX        =   7646
            _ExtentY        =   1508
            BackColor       =   16744576
            Caption         =   "EMPRESA"
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
            Begin VB.TextBox DATO1 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   240
               TabIndex        =   17
               Text            =   "01"
               Top             =   360
               Width           =   375
            End
            Begin VB.Label empresanombre 
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
               Left            =   840
               TabIndex        =   16
               Top             =   360
               Width           =   3255
            End
         End
         Begin XPFrame.FrameXp FrameXp6 
            Height          =   855
            Left            =   120
            TabIndex        =   8
            Top             =   1320
            Width           =   4335
            _ExtentX        =   7646
            _ExtentY        =   1508
            BackColor       =   16744576
            Caption         =   "MES"
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
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Left            =   240
               TabIndex        =   10
               Top             =   360
               Width           =   3855
            End
         End
         Begin XPFrame.FrameXp FrameXp7 
            Height          =   855
            Left            =   120
            TabIndex        =   9
            Top             =   2280
            Width           =   4335
            _ExtentX        =   7646
            _ExtentY        =   1508
            BackColor       =   16744576
            Caption         =   "AÑO"
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
            Begin VB.ComboBox COMBOAÑO 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Left            =   240
               TabIndex        =   11
               Top             =   360
               Width           =   3855
            End
         End
         Begin XPFrame.FrameXp FrameXp9 
            Height          =   855
            Left            =   120
            TabIndex        =   30
            Top             =   3240
            Width           =   4335
            _ExtentX        =   7646
            _ExtentY        =   1508
            BackColor       =   16744576
            Caption         =   "Centros de Costo"
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
            Begin VB.ComboBox Combocrcc 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Left            =   240
               TabIndex        =   31
               Top             =   360
               Width           =   3855
            End
         End
      End
      Begin XPFrame.FrameXp FrameXp1 
         Height          =   1005
         Left            =   240
         TabIndex        =   13
         Top             =   2760
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   1773
         BackColor       =   16761024
         Caption         =   "Detalle Imputaciones"
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
         Begin VB.OptionButton DETALLE1 
            BackColor       =   &H00FFC0C0&
            Caption         =   "Si"
            Height          =   375
            Left            =   480
            TabIndex        =   15
            Top             =   240
            Width           =   2055
         End
         Begin VB.OptionButton DETALLE2 
            BackColor       =   &H00FFC0C0&
            Caption         =   "No"
            Height          =   375
            Left            =   480
            TabIndex        =   14
            Top             =   600
            Width           =   2055
         End
      End
      Begin XPFrame.FrameXp FrameXp8 
         Height          =   855
         Left            =   3510
         TabIndex        =   26
         Top             =   4635
         Width           =   3630
         _ExtentX        =   6403
         _ExtentY        =   1508
         BackColor       =   16761024
         Caption         =   "TIPO DE IMPRESION"
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
         Begin VB.OptionButton original 
            BackColor       =   &H00FFC0C0&
            Caption         =   "Imprime Original"
            Height          =   255
            Left            =   90
            TabIndex        =   29
            Top             =   270
            Width           =   1575
         End
         Begin VB.OptionButton timbrado 
            BackColor       =   &H00FFC0C0&
            Caption         =   "Imprime Timbrado"
            Height          =   255
            Left            =   90
            TabIndex        =   28
            Top             =   540
            Width           =   1695
         End
         Begin VB.TextBox FOLIO 
            Height          =   285
            Left            =   1980
            MaxLength       =   8
            TabIndex        =   27
            Top             =   315
            Width           =   1455
         End
      End
      Begin XPFrame.FrameXp FrameXp10 
         Height          =   2895
         Left            =   270
         TabIndex        =   32
         Top             =   3825
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   5106
         BackColor       =   16761024
         Caption         =   "FILTROS DE IMPRESION"
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
         Begin VB.OptionButton Option9 
            BackColor       =   &H00FFC0C0&
            Caption         =   "Facturas Exenta"
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
            Left            =   45
            TabIndex        =   44
            Top             =   2160
            Width           =   2595
         End
         Begin VB.OptionButton Option8 
            BackColor       =   &H00FFC0C0&
            Caption         =   "Facturas Exportacion"
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
            Left            =   45
            TabIndex        =   40
            Top             =   1920
            Width           =   2595
         End
         Begin VB.OptionButton Option7 
            BackColor       =   &H00FFC0C0&
            Caption         =   "Notas de Credito Factura"
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
            Left            =   45
            TabIndex        =   39
            Top             =   1680
            Width           =   2685
         End
         Begin VB.OptionButton Option6 
            BackColor       =   &H00FFC0C0&
            Caption         =   "Notas de Credito Boleta"
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
            Left            =   45
            TabIndex        =   38
            Top             =   1440
            Width           =   2685
         End
         Begin VB.OptionButton Option5 
            BackColor       =   &H00FFC0C0&
            Caption         =   "Notas de debito"
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
            Left            =   45
            TabIndex        =   37
            Top             =   1200
            Width           =   2055
         End
         Begin VB.OptionButton Option4 
            BackColor       =   &H00FFC0C0&
            Caption         =   "Facturas Empresa Relac."
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
            Left            =   45
            TabIndex        =   36
            Top             =   960
            Width           =   2685
         End
         Begin VB.OptionButton Option3 
            BackColor       =   &H00FFC0C0&
            Caption         =   "Facturas Publicidad"
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
            Left            =   45
            TabIndex        =   35
            Top             =   720
            Width           =   2055
         End
         Begin VB.OptionButton Option2 
            BackColor       =   &H00FFC0C0&
            Caption         =   "Facturas"
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
            Left            =   45
            TabIndex        =   34
            Top             =   480
            Width           =   2055
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00FFC0C0&
            Caption         =   "Todas"
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
            Left            =   45
            TabIndex        =   33
            Top             =   225
            Width           =   2055
         End
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " COD RECTIFICACION"
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   120
         TabIndex        =   49
         Top             =   7320
         Width           =   3015
      End
   End
   Begin XPFrame.FrameXp FrameQuickMenu 
      Height          =   615
      Left            =   4800
      TabIndex        =   45
      Top             =   9360
      Width           =   3135
      _ExtentX        =   5530
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
      Begin VB.CommandButton botonmisfavoritos 
         Caption         =   "Mis Favoritos"
         Height          =   255
         Left            =   1680
         TabIndex        =   47
         Top             =   280
         Width           =   1335
      End
      Begin VB.CommandButton botonmisaccesos 
         Caption         =   "Permisos Modulo"
         Height          =   255
         Left            =   120
         TabIndex        =   46
         Top             =   280
         Width           =   1455
      End
   End
End
Attribute VB_Name = "auxiliar44"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private FORMATOGRILLA(20, 20)
Private lin As Double
Private tipoprove As String
Private plan(2000, 3) As Variant
Private canplan As Double
Private total(20) As Double
Private detalle(20, 20) As Double
Private TIPOS(12) As String
Private MES As String
Private año As String
Private centro As String
Private refrescos As String
Private licores As String
Private vinos As String
Private cerveza As String
Private HARINA As String
Private CARNE As String
Private SINAZUCAR As String
Private netocom As String
Private ivacom As String
Private totcom As String
Private ivacigarro As Double
Private netocigarro As Double




Private Sub COMMAND2_Click()
    Dim TIMBRA As String
    xmlventa = True

    If original.Value = True Then TIMBRA = "N" Else TIMBRA = "S"

    Dim infogrilla As grillainformes
    Set infogrilla = New grillainformes

    Call Conectartemporal(Servidor, clientesistema + "conta" + dato1.text, Usuario, password)
    centro = Mid(Combocrcc.text, 1, 2) + Mid(Combocrcc.text, 4, 2)
    año = COMBOAÑO.text
    MES = COMBOMES.ListIndex + 1
    If Val(MES) < 10 Then MES = "0" + Mid(Str(MES), 2, 1) Else MES = Mid(Str(MES), 2, 2)
    FECHALV = año + MES

    CARGAmayor
    leermayor
    
'     Call AgregaTotalBoletas
     
     Call AgregaTotalBoletasafectas
     Call AgregaTotalBoletasexentas
     
    If Check2.Value = "1" Then
        Call CARGAGRILLA2(infogrilla)
        Call Consulta_Informe2(infogrilla)
        
    Else
        Call CARGAGRILLA(infogrilla)
        Call Consulta_Informe(infogrilla)
    End If
    
    For k = 1 To 15
        detalle(k, 1) = 0
        detalle(k, 2) = 0
        detalle(k, 3) = 0
        detalle(k, 4) = 0
        detalle(k, 5) = 0
        detalle(k, 6) = 0
        detalle(k, 7) = 0
        detalle(k, 8) = 0
        detalle(k, 9) = 0
        detalle(k, 10) = 0
        detalle(k, 11) = 0
        detalle(k, 12) = 0
    
    Next k
   
    infogrilla.Visible = True
    infogrilla.Caption = "LIBRO DE VENTAS " + Combocrcc.text: grillainformes.Tag = "auxiliar44" & TIMBRA & FOLIO.text
    infogrilla.CABEZA.Caption = "LISTADO DE VENTAS " + Combocrcc.text
    infogrilla.Show
    D35_cantidad = Val(Grid1.Cell(1, 2).text)
    D35_neto = Val(Grid1.Cell(1, 3).text)
    D35_iva = Val(Grid1.Cell(1, 4).text)
    D35_total = Val(Grid1.Cell(1, 5).text)
    
    D38_cantidad = Val(Grid1.Cell(2, 2).text)
    D38_neto = Val(Grid1.Cell(2, 3).text)
    D38_iva = Val(Grid1.Cell(2, 4).text)
    D38_total = Val(Grid1.Cell(2, 5).text)
    
    D48_cantidad = Val(Grid1.Cell(3, 2).text)
    D48_neto = Val(Grid1.Cell(3, 3).text)
    D48_iva = Val(Grid1.Cell(3, 4).text)
    D48_total = Val(Grid1.Cell(3, 5).text)
    
    CODAUTREC = txtcodigoautorizacion.text
End Sub


Sub AgregaTotalBoletasafectas()
Dim csql As New rdoQuery
Dim resultados  As rdoResultset
Dim FECHACONSULTA As String

FECHACONSULTA = Format(fechasistema, "yyyy-mm")
Set csql.ActiveConnection = contadb
 

csql.sql = "SELECT  IFNULL(SUM(boletafinal-(boletainicial-1)),0) AS emitidas, "
csql.sql = csql.sql & "IFNULL(ROUND(SUM(monto) / 1.19),0) AS NETO ,"
csql.sql = csql.sql & "IFNULL(SUM(monto) -ROUND(SUM(monto)/1.19),0) AS iva ,"
csql.sql = csql.sql & "IFNULL(SUM(total),0) AS total,estbk,cigarro,IFNULL(SUM(monto),0) "
csql.sql = csql.sql & "FROM boletasdeventa WHERE fecha LIKE '" & FECHACONSULTA & "%' and monto>0 "
csql.sql = csql.sql & "GROUP BY estbk,cigarro "


csql.Execute

netocigarro = 0
ivacigarro = 0
If csql.RowsAffected > 0 Then
    Set resultados = csql.OpenResultset
    While resultados.EOF = False
   Call leercigarros(FECHACONSULTA)
    If resultados(4) <> "1" Then
        Grid1.Cell(1, 2).text = resultados("emitidas")
        Grid1.Cell(1, 3).text = resultados(1) - netocigarro
        Grid1.Cell(1, 4).text = resultados(2) - ivacigarro
        Grid1.Cell(1, 5).text = resultados("total") - (netocigarro + ivacigarro)
    End If
    If resultados(4) = "1" Then
        Grid1.Cell(3, 2).text = resultados("emitidas")
        Grid1.Cell(3, 3).text = resultados("neto")
        Grid1.Cell(3, 4).text = resultados("iva")
        Grid1.Cell(3, 5).text = resultados("total")
    End If
    resultados.MoveNext
    Wend
End If
Grid1.Refresh
End Sub
Sub leercigarros(FECHACONSULTA)
    Dim csql As New rdoQuery
    Dim resultados  As rdoResultset
    
    
    Set csql.ActiveConnection = contadb
 

csql.sql = "SELECT  "
csql.sql = csql.sql & "IFNULL(ROUND(SUM(exento) / 1.19),0) AS NETO ,"
csql.sql = csql.sql & "IFNULL(SUM(exento) -ROUND(SUM(exento)/1.19),0) AS iva  "
csql.sql = csql.sql & "FROM boletasdeventa WHERE fecha LIKE '" & FECHACONSULTA & "%' and cigarro='1'  "
 

csql.Execute

netocigarro = 0
ivacigarro = 0
If csql.RowsAffected > 0 Then
    Set resultados = csql.OpenResultset
    While resultados.EOF = False
        netocigarro = netocigarro + resultados(0)
        ivacigarro = ivacigarro + resultados(1)
        resultados.MoveNext
    Wend
End If
 csql.Close
 Set csql = Nothing


End Sub
Sub AgregaTotalBoletasexentas()
Dim csql As New rdoQuery
Dim resultados  As rdoResultset
Dim FECHACONSULTA As String
FECHACONSULTA = Format(fechasistema, "yyyy-mm")
Set csql.ActiveConnection = contadb
 

csql.sql = "SELECT  IFNULL(SUM(boletafinal-(boletainicial-1)),0) AS emitidas, "
csql.sql = csql.sql & "IFNULL(ROUND(SUM(exento)),0) AS NETO , '0' AS iva ,"
csql.sql = csql.sql & "IFNULL(SUM(total),0) AS total,SUM(exento),estbk,cigarro  "
csql.sql = csql.sql & "FROM boletasdeventa WHERE fecha LIKE '" & FECHACONSULTA & "%' AND exento>0 "
csql.sql = csql.sql & "AND cigarro='0'  GROUP BY estbk,cigarro "


csql.Execute
If csql.RowsAffected > 0 Then
    Set resultados = csql.OpenResultset
    While resultados.EOF = False
    
   
        Grid1.Cell(2, 2).text = resultados("emitidas")
        Grid1.Cell(2, 3).text = resultados("neto")
        Grid1.Cell(2, 4).text = resultados("iva")
        Grid1.Cell(2, 5).text = resultados("total")
    
    resultados.MoveNext
    Wend
End If
Grid1.Refresh
End Sub



Private Sub command8_Click()
Call retornofecha(desdefecha, hastafecha)
End Sub

Private Sub dato1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF2 Then Call ayudaempresa(dato1)
    
End Sub

Sub leer()
    campos(0, 0) = "codigoempresa"
    campos(1, 0) = "nombre"
    campos(2, 0) = ""
    campos(0, 2) = "maestroempresas"
    condicion = "codigoempresa=" + "'" + dato1.text + "'"
    op = 5
    sqlconta.response = campos
    Set sqlconta.conexion = conta
    Call sqlconta.sqlconta(op, condicion)
    If sqlconta.status = 4 Then dato1.SetFocus: GoTo no:
    COMBOMES.SetFocus
    empresanombre.Caption = sqlconta.response(1, 3)
no:
End Sub

Private Sub datos1_Click()
If datos2.Value = True Then fechas.Visible = True
If datos2.Value = False Then fechas.Visible = False

End Sub

Private Sub datos2_Click()
If datos2.Value = True Then fechas.Visible = True
If datos2.Value = False Then fechas.Visible = False

End Sub

Private Sub Form_Load()

CENTRAR Me

Dim i As Integer
Dim k As Integer

TIPOS(1) = "FACTURAS "
TIPOS(2) = "NOTAS DE DEBITO"
TIPOS(3) = "NOTAS DE CREDITO BOLETAS"
TIPOS(4) = "NOTAS DE CREDITO FACTURAS"
TIPOS(5) = "FACTURAS EXPORTACION"

TIPOS(6) = "FACTURAS ELECTRONICAS"
TIPOS(7) = "NOTAS DE DEBITO ELECTRONICAS"
TIPOS(8) = "N.CREDITO ELECTRONICAS FACTURA"
TIPOS(9) = "N.CREDITO ELECTRONICAS BOLETA"
TIPOS(10) = "FACTURAS EXENTAS"
TIPOS(11) = "FACTURAS EXENTAS ELECTRONICAS"
TIPOS(12) = "LIQUIDACION-FACTURAS ELECTRONICAS"
    
    Call Conectar_BD
    Call Conectarconta(Servidor, clientesistema + "conta", Usuario, password)
For i = 1 To 12
For k = 1 To 12
detalle(k, i) = 0
Next k

Next i
opciones.Visible = True

original.Value = True
Option1.Value = True

For k = 1 To 12
COMBOMES.AddItem MonthName(k)
Next k
COMBOMES.ListIndex = CDbl(Format(fechasistema, "mm") - 1)
For k = 2000 To Val(Format(fechasistema, "yyyy"))
COMBOAÑO.AddItem k
Next k
COMBOAÑO.ListIndex = k - 2001
dato1.text = empresaactiva
empresanombre.Caption = nombreempresa
datos1.Value = True
RESUMEN1.Value = True
DETALLE1.Value = True
    desdefecha.Caption = fechasistema
    hastafecha.Caption = fechasistema

fechas.Visible = False
CARGAcrcc
Grid1.Rows = 4
Grid1.Cols = 6
Grid1.Cell(0, 1).text = "TIPO"
Grid1.Cell(0, 2).text = "CANTIDAD"
Grid1.Cell(0, 3).text = "NETO"
Grid1.Cell(0, 4).text = "IVA"
Grid1.Cell(0, 5).text = "TOTAL"

Grid1.Column(0).MaxLength = 0
Grid1.Column(1).MaxLength = 3
Grid1.Column(2).MaxLength = 15
Grid1.Column(3).MaxLength = 15
Grid1.Column(4).MaxLength = 15
Grid1.Column(5).MaxLength = 15

Grid1.Cell(1, 1).text = "35"
Grid1.Cell(2, 1).text = "38"
Grid1.Cell(3, 1).text = "48"


End Sub


    
Sub Consulta_Informe(infogrilla As grillainformes)
Dim resultados As rdoResultset
    Dim csql As New rdoQuery
    Dim rut As String
    Dim multi As Double
    Dim FOLIO As Double
    Dim PASO As String
    Dim fecha1 As String
    Dim fecha2 As String
    
    If datos1.Value = True Then
    fecha1 = año + "-" + MES + "-" + "01"
    fecha2 = año + "-" + MES + "-" + "31"
    Else
    fecha1 = Format(desdefecha.Caption, "yyyy-mm-dd")
    fecha2 = Format(hastafecha.Caption, "yyyy-mm-dd")
    
    End If
    
        Set csql.ActiveConnection = contadb
        csql.sql = "SELECT fc.tipo,numero,fecha,fc.rut,ifnull(cc.nombre,''),neto,iva,exento,total,nula,IF(fc.tipo='8',IF(tnc='','F',tnc),tnc) AS tnc "
'        csql.sql = csql.sql + "FROM facturasdeventas as fc,cuentascorrientes as cc "
         csql.sql = csql.sql + "FROM facturasdeventas as fc left join cuentascorrientes as cc on fc.rut=cc.rut and cc.año='" + COMBOAÑO.text + "' and (cc.tipo='" + cuentacliente + "' or cc.tipo='11200028') "
        If Option1.Value = True Then csql.sql = csql.sql + "WHERE " 'fc.tipo<>'' and "
        If Option2.Value = True Then csql.sql = csql.sql + "WHERE (fc.tipo='1' or fc.tipo='6') and fc.caja<'90' and "
        If Option3.Value = True Then csql.sql = csql.sql + "WHERE (fc.tipo='1' or fc.tipo='6') and fc.caja='98' and "
        If Option4.Value = True Then csql.sql = csql.sql + "WHERE (fc.tipo='1' or fc.tipo='6') and fc.caja='99' and "
        If Option5.Value = True Then csql.sql = csql.sql + "WHERE (fc.tipo='2' or fc.tipo='7') and "
        If Option6.Value = True Then csql.sql = csql.sql + "WHERE fc.tipo='3'  and "
        If Option7.Value = True Then csql.sql = csql.sql + "WHERE (fc.tipo='4' or fc.tipo='8') and "
        If Option8.Value = True Then csql.sql = csql.sql + "WHERE fc.tipo='5' and "
        If Option9.Value = True Then csql.sql = csql.sql + "WHERE (fc.tipo='9' or fc.tipo='0') and "
        
        If Option1.Value = False And Option2.Value = False And Option3.Value = False And Option4.Value = False And Option5.Value = False And Option6.Value = False And Option7.Value = False And Option8.Value = False And Option9.Value = False Then
            csql.sql = csql.sql & " and "
        End If
        
        If centro = "9999" Then csql.sql = csql.sql + "fc.rut=cc.rut and cc.año='" + COMBOAÑO.text + "' and (cc.tipo='" + cuentacliente + "' or cc.tipo='11200028') and "
        If centro <> "9999" Then csql.sql = csql.sql + "fc.crcc='" + centro + "' and fc.rut=cc.rut and (cc.tipo='" + cuentacliente + "' or cc.tipo='11200028') and cc.año='" + COMBOAÑO.text + "' and "
        
        csql.sql = csql.sql + " fecha >= '" + fecha1 + "' and fecha <= '" + fecha2 + "' "
        
        'fecha >= '" + año + "/" + mes + "/" + "01" + "' and fecha <= '" + año + "/" + mes + "/" + "31' order by fecha,tipo,numero "
        csql.sql = csql.sql + " GROUP by tipo,numero order by fecha,tipo,numero "
        csql.Execute
        infogrilla.Grid1.AutoRedraw = False
        total(1) = 0
        total(2) = 0
        total(3) = 0
        total(4) = 0
        refrescos = "0"
        licores = "0"
        vinos = "0"
        cerveza = "0"
        HARINA = "0"
        CARNE = "0"
    
        
        If csql.RowsAffected > 0 Then
        barra.Max = csql.RowsAffected + 1
        Set resultados = csql.OpenResultset
        lin = 0
        FOLIO = CDbl(resultados(1))
     
        While Not resultados.EOF
        
        
         If RESUMEN1.Value = True Then
             barra.Value = lin
             lin = lin + 1
             infogrilla.Grid1.Rows = infogrilla.Grid1.Rows + 1
             infogrilla.Grid1.Cell(lin, 0).text = resultados("nula")
             For k = 0 To 8
             infogrilla.Grid1.Cell(lin, k + 1).text = resultados(k)
             Next k
             If FOLIO <> CDbl(resultados(1)) Then infogrilla.Grid1.Cell(lin, 10).text = "***": FOLIO = CDbl(resultados(1))
             
             FOLIO = FOLIO + 1
             
             multi = 1
                        
                If resultados(0) = "1" Then infogrilla.Grid1.Cell(lin, 1).text = "FA": multi = 1
                If resultados(0) = "2" Then infogrilla.Grid1.Cell(lin, 1).text = "ND": multi = 1
                If resultados(0) = "3" Then infogrilla.Grid1.Cell(lin, 1).text = "NB": multi = -1
                If resultados(0) = "4" Then infogrilla.Grid1.Cell(lin, 1).text = "NF": multi = -1
                If resultados(0) = "5" Then infogrilla.Grid1.Cell(lin, 1).text = "FE": multi = 1
                
                If resultados(0) = "6" Then infogrilla.Grid1.Cell(lin, 1).text = "FAE": multi = 1
                If resultados(0) = "7" Then infogrilla.Grid1.Cell(lin, 1).text = "NDE": multi = 1
                
                If resultados(0) = "8" And resultados("tnc") = "F" Then
                infogrilla.Grid1.Cell(lin, 1).text = "NCE": multi = -1
                End If
                If resultados(0) = "8" And resultados("tnc") = "B" Then infogrilla.Grid1.Cell(lin, 1).text = "NCE": multi = -1
                
                If resultados(0) = "9" Then infogrilla.Grid1.Cell(lin, 1).text = "FX": multi = 1
                If resultados(0) = "0" Then infogrilla.Grid1.Cell(lin, 1).text = "FEE": multi = 1
                
                infogrilla.Grid1.Cell(lin, 6).text = resultados(5) * multi
                infogrilla.Grid1.Cell(lin, 7).text = (resultados(6) - refrescos - licores - vinos - cerveza - HARINA - CARNE) * multi
                infogrilla.Grid1.Cell(lin, 8).text = resultados(7) * multi
                infogrilla.Grid1.Cell(lin, 9).text = resultados(8) * multi
                infogrilla.Grid1.Cell(lin, 4).text = Mid(resultados(3), 1, 9) + "-" + Mid(resultados(3), 10, 1)

         
         End If
             If (resultados(0) > "2" And resultados(0) < "5") Or resultados(0) = "8" Then multi = -1 Else multi = 1
             total(1) = total(1) + resultados(5) * multi
             total(2) = total(2) + resultados(6) * multi
             total(3) = total(3) + resultados(7) * multi
             total(4) = total(4) + resultados(8) * multi
             
             If resultados(0) = "1" Then detalle(1, 1) = detalle(1, 1) + 1: detalle(1, 2) = detalle(1, 2) + resultados(5): detalle(1, 3) = detalle(1, 3) + resultados(6):: detalle(1, 4) = detalle(1, 4) + resultados(7):: detalle(1, 5) = detalle(1, 5) + resultados(8)
             If resultados(0) = "2" Then detalle(2, 1) = detalle(2, 1) + 1: detalle(2, 2) = detalle(2, 2) + resultados(5): detalle(2, 3) = detalle(2, 3) + resultados(6):: detalle(2, 4) = detalle(2, 4) + resultados(7):: detalle(2, 5) = detalle(2, 5) + resultados(8)
             If resultados(0) = "3" Then detalle(3, 1) = detalle(3, 1) + 1: detalle(3, 2) = detalle(3, 2) + resultados(5): detalle(3, 3) = detalle(3, 3) + resultados(6):: detalle(3, 4) = detalle(3, 4) + resultados(7):: detalle(3, 5) = detalle(3, 5) + resultados(8)
             If resultados(0) = "4" Then detalle(4, 1) = detalle(4, 1) + 1: detalle(4, 2) = detalle(4, 2) + resultados(5): detalle(4, 3) = detalle(4, 3) + resultados(6):: detalle(4, 4) = detalle(4, 4) + resultados(7):: detalle(4, 5) = detalle(4, 5) + resultados(8)
             If resultados(0) = "5" Then detalle(5, 1) = detalle(5, 1) + 1: detalle(5, 2) = detalle(5, 2) + resultados(5): detalle(5, 3) = detalle(5, 3) + resultados(6):: detalle(5, 4) = detalle(5, 4) + resultados(7):: detalle(5, 5) = detalle(5, 5) + resultados(8)
             If resultados(0) = "6" Then detalle(6, 1) = detalle(6, 1) + 1: detalle(6, 2) = detalle(6, 2) + resultados(5): detalle(6, 3) = detalle(6, 3) + resultados(6):: detalle(6, 4) = detalle(6, 4) + resultados(7):: detalle(6, 5) = detalle(6, 5) + resultados(8)
             If resultados(0) = "7" Then detalle(7, 1) = detalle(7, 1) + 1: detalle(7, 2) = detalle(7, 2) + resultados(5): detalle(7, 3) = detalle(7, 3) + resultados(6):: detalle(7, 4) = detalle(7, 4) + resultados(7):: detalle(7, 5) = detalle(7, 5) + resultados(8)
             If resultados(0) = "8" And resultados("tnc") = "F" Then detalle(8, 1) = detalle(8, 1) + 1: detalle(8, 2) = detalle(8, 2) + resultados(5): detalle(8, 3) = detalle(8, 3) + resultados(6):: detalle(8, 4) = detalle(8, 4) + resultados(7):: detalle(8, 5) = detalle(8, 5) + resultados(8)
             If resultados(0) = "8" And resultados("tnc") = "B" Then detalle(9, 1) = detalle(9, 1) + 1: detalle(9, 2) = detalle(9, 2) + resultados(5): detalle(9, 3) = detalle(9, 3) + resultados(6):: detalle(9, 4) = detalle(9, 4) + resultados(7):: detalle(9, 5) = detalle(9, 5) + resultados(8)
             
Rem              If resultados(0) = "9" Then detalle(10, 1) = detalle(10, 1) + 1: detalle(10, 2) = detalle(10, 2) + resultados(5): detalle(10, 3) = detalle(10, 3) + resultados(6): detalle(10, 4) = detalle(10, 4) + EXENTO:: detalle(10, 5) = detalle(10, 5) + resultados(8): detalle(10, 6) = detalle(10, 6) + refrescos: detalle(10, 7) = detalle(10, 7) + licores: detalle(10, 8) = detalle(10, 8) + vinos: detalle(10, 9) = detalle(10, 9) + cerveza: detalle(10, 10) = detalle(10, 10) + HARINA: detalle(10, 11) = detalle(10, 11) + CARNE
Rem If resultados(0) = "0" Then detalle(11, 1) = detalle(11, 1) + 1: detalle(11, 2) = detalle(11, 2) + resultados(5): detalle(11, 3) = detalle(11, 3) + resultados(6): detalle(11, 4) = detalle(11, 4) + EXENTO:: detalle(11, 5) = detalle(11, 5) + resultados(8): detalle(11, 6) = detalle(11, 6) + refrescos: detalle(11, 7) = detalle(11, 7) + licores: detalle(11, 8) = detalle(11, 8) + vinos: detalle(11, 9) = detalle(11, 9) + cerveza: detalle(11, 10) = detalle(11, 10) + HARINA: detalle(11, 11) = detalle(11, 11) + CARNE
             If resultados(0) = "9" Then detalle(10, 1) = detalle(10, 1) + 1: detalle(10, 2) = detalle(10, 2) + resultados(5): detalle(10, 3) = detalle(10, 3) + resultados(6):: detalle(10, 4) = detalle(10, 4) + resultados(7):: detalle(10, 5) = detalle(10, 5) + resultados(8)
             If resultados(0) = "0" Then detalle(11, 1) = detalle(11, 1) + 1: detalle(11, 2) = detalle(11, 2) + resultados(5): detalle(11, 3) = detalle(11, 3) + resultados(6):: detalle(11, 4) = detalle(11, 4) + resultados(7):: detalle(11, 5) = detalle(11, 5) + resultados(8)


             
              Call Consultadetalle(resultados(0), resultados(1), resultados(2), infogrilla)
PASO:
             resultados.MoveNext


           
         Wend
          
          resultados.Close
            Set resultados = Nothing

        End If
     
Call totallibro(infogrilla)
barra.Max = 1
infogrilla.Grid1.AutoRedraw = True
infogrilla.Grid1.Refresh
fechas.Visible = False

End Sub
Sub Consulta_Informe2(infogrilla As grillainformes)
Dim resultados As rdoResultset
    Dim csql As New rdoQuery
    Dim rut As String
    Dim multi As Double
    Dim FOLIO As Double
    Dim PASO As String
    Dim fecha1 As String
    Dim fecha2 As String
    Dim EXENTO As Double
    
    If datos1.Value = True Then
    fecha1 = año + "-" + MES + "-" + "01"
    fecha2 = año + "-" + MES + "-" + "31"
    Else
    fecha1 = Format(desdefecha.Caption, "yyyy-mm-dd")
    fecha2 = Format(hastafecha.Caption, "yyyy-mm-dd")
    
    End If
    
        Set csql.ActiveConnection = contadb
        csql.sql = "SELECT fc.tipo,numero,fecha,fc.rut,cc.nombre,neto,iva,exento,total,nula,IF(fc.tipo='8',IF(tnc='','F',tnc),tnc) AS tnc,crcc "
        csql.sql = csql.sql + "FROM facturasdeventas as fc left join cuentascorrientes as cc on fc.rut=cc.rut and cc.año='" + COMBOAÑO.text + "' and (cc.tipo='" + cuentacliente + "' or cc.tipo='11200028')  "
        If Option1.Value = True Then csql.sql = csql.sql + "WHERE " 'fc.tipo<>'' and "
        If Option2.Value = True Then csql.sql = csql.sql + "WHERE (fc.tipo='1' or fc.tipo='6') and fc.caja<'90' and "
        If Option3.Value = True Then csql.sql = csql.sql + "WHERE (fc.tipo='1' or fc.tipo='6') and fc.caja='98' and "
        
        If Option4.Value = True Then csql.sql = csql.sql + "WHERE (fc.tipo='1' or fc.tipo='6') and fc.caja='99' and "
        If Option5.Value = True Then csql.sql = csql.sql + "WHERE (fc.tipo='2' or fc.tipo='7') and "
        
        If Option6.Value = True Then csql.sql = csql.sql + "WHERE fc.tipo='3'  and "
        If Option7.Value = True Then csql.sql = csql.sql + "WHERE (fc.tipo='4' or fc.tipo='8') and "
        If Option8.Value = True Then csql.sql = csql.sql + "WHERE fc.tipo='5' and "
        If Option9.Value = True Then csql.sql = csql.sql + "WHERE (fc.tipo='9' or fc.tipo='0') and "
        
        
        If centro <> "9999" Then csql.sql = csql.sql + "fc.crcc='" + centro + "' and "
        
        csql.sql = csql.sql + " fecha >= '" + fecha1 + "' and fecha <= '" + fecha2 + "' "
        
        'fecha >= '" + año + "/" + mes + "/" + "01" + "' and fecha <= '" + año + "/" + mes + "/" + "31' order by fecha,tipo,numero "
        csql.sql = csql.sql + "GROUP by tipo,numero order by fecha,tipo,numero "
        csql.Execute
        infogrilla.Grid1.AutoRedraw = False
        total(1) = 0
        total(2) = 0
        total(3) = 0
        total(4) = 0
        total(5) = 0
        total(6) = 0
        total(7) = 0
        total(8) = 0
        total(9) = 0
        total(10) = 0
        total(11) = 0
        refrescos = "0"
        licores = "0"
        vinos = "0"
        cerveza = "0"
        HARINA = "0"
        CARNE = "0"
        SINAZUCAR = "0"
        netocom = "0"
        ivacom = "0"
        totcom = "0"
        
    
        
        If csql.RowsAffected > 0 Then
        barra.Max = csql.RowsAffected + 1
        Set resultados = csql.OpenResultset
        lin = 0
        FOLIO = CDbl(resultados(1))
        nogenerar = False
        While Not resultados.EOF
        
        
         If RESUMEN1.Value = True Then
             barra.Value = lin
             lin = lin + 1
             infogrilla.Grid1.Rows = infogrilla.Grid1.Rows + 1
             infogrilla.Grid1.Cell(lin, 0).text = resultados("nula")
             If fechasistema > "2016-12-31" Then
             refrescos = leerimpuestov(resultados(0), resultados(1), resultados(4), "23300010")
             licores = leerimpuestov(resultados(0), resultados(1), resultados(4), "23300013")
             vinos = leerimpuestov(resultados(0), resultados(1), resultados(4), "23300011")
             cerveza = leerimpuestov(resultados(0), resultados(1), resultados(4), "23300014")
             HARINA = leerimpuestov(resultados(0), resultados(1), resultados(4), "23200005")
             CARNE = leerimpuestov(resultados(0), resultados(1), resultados(4), "23200009")
             SINAZUCAR = leerimpuestov(resultados(0), resultados(1), resultados(4), "23300017")
             Else
             
             refrescos = leerimpuestov(resultados(0), resultados(1), resultados(4), "11400010")
             licores = leerimpuestov(resultados(0), resultados(1), resultados(4), "11400013")
             vinos = leerimpuestov(resultados(0), resultados(1), resultados(4), "11400011")
             cerveza = leerimpuestov(resultados(0), resultados(1), resultados(4), "11400014")
             HARINA = leerimpuestov(resultados(0), resultados(1), resultados(4), "23200005")
             CARNE = leerimpuestov(resultados(0), resultados(1), resultados(4), "23200009")
             SINAZUCAR = leerimpuestov(resultados(0), resultados(1), resultados(4), "11400017")
             End If
             
             
             'comision liquidacion
             netocom = "0"
             ivacom = "0"
             totcom = "0"
             If resultados(0) = "L" Then
                netocom = leedatocomision(resultados(0), resultados(1), resultados(3), "comision+otros")
                ivacom = leedatocomision(resultados(0), resultados(1), resultados(3), "iva")
                totcom = leedatocomision(resultados(0), resultados(1), resultados(3), "total")
             End If
             
            
             
             
             
             For k = 0 To 8
               If IsNull(resultados(k)) = False Then
                infogrilla.Grid1.Cell(lin, k + 1).text = resultados(k)
                End If
             
             Next k
             infogrilla.Grid1.Cell(lin, 4).text = Replace(infogrilla.Grid1.Cell(lin, 4).text, ".", "0")
             infogrilla.Grid1.Cell(lin, 11).text = refrescos
             infogrilla.Grid1.Cell(lin, 12).text = licores
             infogrilla.Grid1.Cell(lin, 13).text = vinos
             infogrilla.Grid1.Cell(lin, 14).text = cerveza
             infogrilla.Grid1.Cell(lin, 15).text = HARINA
             infogrilla.Grid1.Cell(lin, 16).text = CARNE
             infogrilla.Grid1.Cell(lin, 17).text = SINAZUCAR
             
             infogrilla.Grid1.Cell(lin, 18).text = netocom
             infogrilla.Grid1.Cell(lin, 19).text = ivacom
             infogrilla.Grid1.Cell(lin, 20).text = totcom
             
             
             
             If FOLIO <> CDbl(resultados(1)) Then infogrilla.Grid1.Cell(lin, 10).text = "***": FOLIO = CDbl(resultados(1))
             
             FOLIO = FOLIO + 1
             
             multi = 1
                        
                If resultados(0) = "1" Then infogrilla.Grid1.Cell(lin, 1).text = "FA": multi = 1
                If resultados(0) = "2" Then infogrilla.Grid1.Cell(lin, 1).text = "ND": multi = 1
                If resultados(0) = "3" Then infogrilla.Grid1.Cell(lin, 1).text = "NB": multi = -1
                If resultados(0) = "4" Then infogrilla.Grid1.Cell(lin, 1).text = "NF": multi = -1
                If resultados(0) = "5" Then infogrilla.Grid1.Cell(lin, 1).text = "FE": multi = 1
                If resultados(0) = "6" Then infogrilla.Grid1.Cell(lin, 1).text = "FAE": multi = 1
                If resultados(0) = "7" Then infogrilla.Grid1.Cell(lin, 1).text = "NDE": multi = 1
                
                
                
                If resultados(0) = "8" And resultados("tnc") = "F" Then infogrilla.Grid1.Cell(lin, 1).text = "NCE": multi = -1
                If resultados(0) = "8" And resultados("tnc") = "B" Then infogrilla.Grid1.Cell(lin, 1).text = "NCE": multi = -1
                If resultados(0) = "9" Then infogrilla.Grid1.Cell(lin, 1).text = "FX": multi = 1
                If resultados(0) = "0" Then infogrilla.Grid1.Cell(lin, 1).text = "FEE": multi = 1
                If resultados(0) = "L" Then infogrilla.Grid1.Cell(lin, 1).text = "LFE": multi = 1
                
                infogrilla.Grid1.Cell(lin, 6).text = resultados(5) * multi
                infogrilla.Grid1.Cell(lin, 7).text = resultados(6) * multi
                infogrilla.Grid1.Cell(lin, 8).text = (resultados(7) - refrescos - licores - vinos - cerveza - HARINA - CARNE - SINAZUCAR) * multi
                infogrilla.Grid1.Cell(lin, 9).text = resultados(8) * multi
                infogrilla.Grid1.Cell(lin, 4).text = Mid(resultados(3), 1, 9) + "-" + Mid(resultados(3), 10, 1)
                infogrilla.Grid1.Cell(lin, 11).text = refrescos * multi
                infogrilla.Grid1.Cell(lin, 12).text = licores * multi
                infogrilla.Grid1.Cell(lin, 13).text = vinos * multi
                infogrilla.Grid1.Cell(lin, 14).text = cerveza * multi
                infogrilla.Grid1.Cell(lin, 15).text = HARINA * multi
                infogrilla.Grid1.Cell(lin, 16).text = CARNE * multi
                infogrilla.Grid1.Cell(lin, 17).text = SINAZUCAR * multi
                
                infogrilla.Grid1.Cell(lin, 18).text = netocom * multi
                infogrilla.Grid1.Cell(lin, 19).text = ivacom * multi
                infogrilla.Grid1.Cell(lin, 20).text = totcom * multi
                If resultados(0) > 5 Or (resultados(0) = "L" Or resultados(0) = "0") Then
                    If verificaenviado(resultados(0), resultados(1), resultados(2), resultados(11)) = False Then
                         infogrilla.Grid1.Range(lin, 1, lin, infogrilla.Grid1.Cols - 1).BackColor = vbRed
                         nogenerar = True
                    End If
                End If
             
         
         End If
          
             
             If (resultados(0) > "2" And resultados(0) < "5") Or resultados(0) = "8" Then multi = -1 Else multi = 1
             total(1) = total(1) + resultados(5) * multi
             total(2) = total(2) + resultados(6) * multi
             total(3) = total(3) + (resultados(7) - refrescos - licores - vinos - cerveza - HARINA - CARNE - SINAZUCAR) * multi
             total(4) = total(4) + resultados(8) * multi
             total(5) = total(5) + refrescos * multi
             total(6) = total(6) + licores * multi
             total(7) = total(7) + vinos * multi
             total(8) = total(8) + cerveza * multi
             total(9) = total(9) + HARINA * multi
             total(10) = total(10) + CARNE * multi
             total(11) = total(11) + SINAZUCAR * multi
             
    EXENTO = resultados(7) - refrescos - licores - vinos - cerveza - HARINA - CARNE - SINAZUCAR
             
             
             
             If resultados(0) = "1" Then detalle(1, 1) = detalle(1, 1) + 1: detalle(1, 2) = detalle(1, 2) + resultados(5): detalle(1, 3) = detalle(1, 3) + resultados(6): detalle(1, 4) = detalle(1, 4) + EXENTO: detalle(1, 5) = detalle(1, 5) + resultados(8): detalle(1, 6) = detalle(1, 6) + refrescos: detalle(1, 7) = detalle(1, 7) + licores: detalle(1, 8) = detalle(1, 8) + vinos: detalle(1, 9) = detalle(1, 9) + cerveza: detalle(1, 10) = detalle(1, 10) + HARINA: detalle(1, 11) = detalle(1, 11) + CARNE: detalle(1, 12) = detalle(1, 12) + SINAZUCAR
             If resultados(0) = "2" Then detalle(2, 1) = detalle(2, 1) + 1: detalle(2, 2) = detalle(2, 2) + resultados(5): detalle(2, 3) = detalle(2, 3) + resultados(6): detalle(2, 4) = detalle(2, 4) + EXENTO:: detalle(2, 5) = detalle(2, 5) + resultados(8): detalle(2, 6) = detalle(2, 6) + refrescos: detalle(2, 7) = detalle(2, 7) + licores: detalle(2, 8) = detalle(2, 8) + vinos: detalle(2, 9) = detalle(2, 9) + cerveza: detalle(2, 10) = detalle(2, 10) + HARINA: detalle(2, 11) = detalle(2, 11) + CARNE: detalle(2, 12) = detalle(2, 12) + SINAZUCAR
             If resultados(0) = "3" Then detalle(3, 1) = detalle(3, 1) + 1: detalle(3, 2) = detalle(3, 2) + resultados(5): detalle(3, 3) = detalle(3, 3) + resultados(6): detalle(3, 4) = detalle(3, 4) + EXENTO:: detalle(3, 5) = detalle(3, 5) + resultados(8): detalle(3, 6) = detalle(3, 6) + refrescos: detalle(3, 7) = detalle(3, 7) + licores: detalle(3, 8) = detalle(3, 8) + vinos: detalle(3, 9) = detalle(3, 9) + cerveza: detalle(3, 10) = detalle(3, 10) + HARINA: detalle(3, 11) = detalle(3, 11) + CARNE: detalle(3, 12) = detalle(3, 12) + SINAZUCAR
             If resultados(0) = "4" Then detalle(4, 1) = detalle(4, 1) + 1: detalle(4, 2) = detalle(4, 2) + resultados(5): detalle(4, 3) = detalle(4, 3) + resultados(6): detalle(4, 4) = detalle(4, 4) + EXENTO:: detalle(4, 5) = detalle(4, 5) + resultados(8): detalle(4, 6) = detalle(4, 6) + refrescos: detalle(4, 7) = detalle(4, 7) + licores: detalle(4, 8) = detalle(4, 8) + vinos: detalle(4, 9) = detalle(4, 9) + cerveza: detalle(4, 10) = detalle(4, 10) + HARINA: detalle(4, 11) = detalle(4, 11) + CARNE: detalle(4, 12) = detalle(4, 12) + SINAZUCAR
             If resultados(0) = "5" Then detalle(5, 1) = detalle(5, 1) + 1: detalle(5, 2) = detalle(5, 2) + resultados(5): detalle(5, 3) = detalle(5, 3) + resultados(6): detalle(5, 4) = detalle(5, 4) + EXENTO:: detalle(5, 5) = detalle(5, 5) + resultados(8): detalle(5, 6) = detalle(5, 6) + refrescos: detalle(5, 7) = detalle(5, 7) + licores: detalle(5, 8) = detalle(5, 8) + vinos: detalle(5, 9) = detalle(5, 9) + cerveza: detalle(5, 10) = detalle(5, 10) + HARINA: detalle(5, 11) = detalle(5, 11) + CARNE: detalle(5, 12) = detalle(5, 12) + SINAZUCAR
             If resultados(0) = "6" Then detalle(6, 1) = detalle(6, 1) + 1: detalle(6, 2) = detalle(6, 2) + resultados(5): detalle(6, 3) = detalle(6, 3) + resultados(6): detalle(6, 4) = detalle(6, 4) + EXENTO:: detalle(6, 5) = detalle(6, 5) + resultados(8): detalle(6, 6) = detalle(6, 6) + refrescos: detalle(6, 7) = detalle(6, 7) + licores: detalle(6, 8) = detalle(6, 8) + vinos: detalle(6, 9) = detalle(6, 9) + cerveza: detalle(6, 10) = detalle(6, 10) + HARINA: detalle(6, 11) = detalle(6, 11) + CARNE: detalle(6, 12) = detalle(6, 12) + SINAZUCAR
             If resultados(0) = "7" Then detalle(7, 1) = detalle(7, 1) + 1: detalle(7, 2) = detalle(7, 2) + resultados(5): detalle(7, 3) = detalle(7, 3) + resultados(6): detalle(7, 4) = detalle(7, 4) + EXENTO:: detalle(7, 5) = detalle(7, 5) + resultados(8): detalle(7, 6) = detalle(7, 6) + refrescos: detalle(7, 7) = detalle(7, 7) + licores: detalle(7, 8) = detalle(7, 8) + vinos: detalle(7, 9) = detalle(7, 9) + cerveza: detalle(7, 10) = detalle(7, 10) + HARINA: detalle(7, 11) = detalle(7, 11) + CARNE: detalle(7, 12) = detalle(7, 12) + SINAZUCAR
             If resultados(0) = "8" And resultados("tnc") = "F" Then detalle(8, 1) = detalle(8, 1) + 1: detalle(8, 2) = detalle(8, 2) + resultados(5): detalle(8, 3) = detalle(8, 3) + resultados(6): detalle(8, 4) = detalle(8, 4) + EXENTO:: detalle(8, 5) = detalle(8, 5) + resultados(8): detalle(8, 6) = detalle(8, 6) + refrescos: detalle(8, 7) = detalle(8, 7) + licores: detalle(8, 8) = detalle(8, 8) + vinos: detalle(8, 9) = detalle(8, 9) + cerveza: detalle(8, 10) = detalle(8, 10) + HARINA: detalle(8, 11) = detalle(8, 11) + CARNE: detalle(8, 12) = detalle(8, 12) + SINAZUCAR
             If resultados(0) = "8" And resultados("tnc") = "B" Then detalle(9, 1) = detalle(9, 1) + 1: detalle(9, 2) = detalle(9, 2) + resultados(5): detalle(9, 3) = detalle(9, 3) + resultados(6): detalle(9, 4) = detalle(9, 4) + EXENTO:: detalle(9, 5) = detalle(9, 5) + resultados(8): detalle(9, 6) = detalle(9, 6) + refrescos: detalle(9, 7) = detalle(9, 7) + licores: detalle(9, 8) = detalle(9, 8) + vinos: detalle(9, 9) = detalle(9, 9) + cerveza: detalle(9, 10) = detalle(9, 10) + HARINA: detalle(9, 11) = detalle(9, 11) + CARNE: detalle(9, 12) = detalle(9, 12) + SINAZUCAR
             If resultados(0) = "9" Then detalle(10, 1) = detalle(10, 1) + 1: detalle(10, 2) = detalle(10, 2) + resultados(5): detalle(10, 3) = detalle(10, 3) + resultados(6): detalle(10, 4) = detalle(10, 4) + EXENTO:: detalle(10, 5) = detalle(10, 5) + resultados(8): detalle(10, 6) = detalle(10, 6) + refrescos: detalle(10, 7) = detalle(10, 7) + licores: detalle(10, 8) = detalle(10, 8) + vinos: detalle(10, 9) = detalle(10, 9) + cerveza: detalle(10, 10) = detalle(10, 10) + HARINA: detalle(10, 11) = detalle(10, 11) + CARNE: detalle(10, 12) = detalle(10, 12) + SINAZUCAR
             If resultados(0) = "0" Then detalle(11, 1) = detalle(11, 1) + 1: detalle(11, 2) = detalle(11, 2) + resultados(5): detalle(11, 3) = detalle(11, 3) + resultados(6): detalle(11, 4) = detalle(11, 4) + EXENTO:: detalle(11, 5) = detalle(11, 5) + resultados(8): detalle(11, 6) = detalle(11, 6) + refrescos: detalle(11, 7) = detalle(11, 7) + licores: detalle(11, 8) = detalle(11, 8) + vinos: detalle(11, 9) = detalle(11, 9) + cerveza: detalle(11, 10) = detalle(11, 10) + HARINA: detalle(11, 11) = detalle(11, 11) + CARNE: detalle(11, 12) = detalle(11, 12) + SINAZUCAR
             If resultados(0) = "L" Then detalle(12, 1) = detalle(12, 1) + 1: detalle(12, 2) = detalle(12, 2) + resultados(5): detalle(12, 3) = detalle(12, 3) + resultados(6): detalle(12, 4) = detalle(12, 4) + EXENTO:: detalle(12, 5) = detalle(12, 5) + resultados(8): detalle(12, 6) = detalle(12, 6) + refrescos: detalle(12, 7) = detalle(12, 7) + licores: detalle(12, 8) = detalle(12, 8) + vinos: detalle(12, 9) = detalle(12, 9) + cerveza: detalle(12, 10) = detalle(12, 10) + HARINA: detalle(12, 11) = detalle(12, 11) + CARNE: detalle(12, 12) = detalle(12, 12) + SINAZUCAR: detalle(12, 13) = detalle(12, 13) + netocom: detalle(12, 14) = detalle(12, 14) + ivacom: detalle(12, 15) = detalle(12, 15) + totcom

             
             Rem  Call Consultadetalle(resultados(0), resultados(1), resultados(2), infogrilla)
PASO:
             resultados.MoveNext


           
         Wend
          If verificadtes(fecha1, fecha2) = True Then
             nogenerar = True
          End If
          resultados.Close
            Set resultados = Nothing

        End If
     
Call totallibro2(infogrilla)
barra.Max = 1
infogrilla.Grid1.AutoRedraw = True
infogrilla.Grid1.Refresh
fechas.Visible = False

End Sub
  Function verificadtes(desde, hasta) As Boolean
    Dim csql As New rdoQuery
    Dim csql2 As New rdoQuery
    Dim contador As Double
    
    Dim resultados As rdoResultset
    Set csql.ActiveConnection = contadb
    Set csql2.ActiveConnection = contadb
    csql.sql = "select codigo from " & cliente_sql & "gestion.g_maestroempresas "
    csql.sql = csql.sql & " where codigocontable='" & empresaactiva & "'  and facturadorelectronico='1' "
    csql.Execute
    
    If csql.RowsAffected > 0 Then
        Set resultados = csql.OpenResultset
        contador = 0
        csql2.sql = ""
        While Not resultados.EOF
            
            csql2.sql = csql2.sql & "SELECT numero FROM " & cliente_sql & "fae" & resultados(0) & ".sv_dte" & resultados(0) & " AS dte "
            csql2.sql = csql2.sql & "WHERE NOT EXISTS (SELECT 1 FROM  eltit_conta" & empresaactiva & ".facturasdeventas AS fv  WHERE  "
            csql2.sql = csql2.sql & "fecha BETWEEN '" & Format(desde, "yyyy-mm-dd") & "' AND '" & Format(hasta, "yyyy-mm-dd") & "' AND fv.numero = LPAD(dte.numero,10,'0') "
            csql2.sql = csql2.sql & "AND fv.fecha=dte.fecha AND fv.rut=LPAD(REPLACE(dte.rut,'-',''),10,0))  "
            csql2.sql = csql2.sql & "AND fecha BETWEEN '" & Format(desde, "yyyy-mm-dd") & "' AND '" & Format(hasta, "yyyy-mm-dd") & "' AND tipo<>'52' "
            
            resultados.MoveNext
            If resultados.EOF = False Then
                csql2.sql = csql2.sql & " union  "
            End If
            contador = contador + 1
        Wend
        csql2.Execute
            verificadtes = False
        If csql2.RowsAffected > 0 Then
            verificadtes = True
        Else
             verificadtes = False
         End If
              
         
    End If
    
    
  End Function
Function verificaenviado(tipo, numero, fecha, centrocosto) As Boolean
    Dim csql As New rdoQuery
    Dim resultados As rdoResultset
    Dim csql2 As New rdoQuery
    Dim TIPOS As String
    Dim loc As String
    Dim resultado2 As rdoResultset
    
    
    If tipo = "6" Then
        TIPOS = "33"
    End If
    
     If tipo = "7" Then
        TIPOS = "56"
    End If
    
    If tipo = "8" Or tipo = "9" Then
        TIPOS = "61"
       
    End If
    
    If tipo = "0" Then
        TIPOS = "34"
    End If
    
    If tipo = "L" Then
        TIPOS = "46"
    End If
    
    verificaenviado = False
    
    Set csql.ActiveConnection = contadb
    Set csql2.ActiveConnection = contadb
    If centrocosto = "" Then centrocosto = "0101"
    
    csql.sql = "select codigo from " & cliente_sql & "gestion.g_maestroempresas "
    csql.sql = csql.sql & " where codigocontable='" & empresaactiva & "' and codigocrcc='" & centrocosto & "' and facturadorelectronico='1'"
    csql.Execute
    
    If csql.RowsAffected > 0 Then
        Set resultados = csql.OpenResultset
        While Not resultados.EOF
            csql2.sql = "select glosa_sii from " & cliente_sql & "fae" & resultados(0) & ".sv_dte" & resultados(0)
            csql2.sql = csql2.sql & " where tipo='" & TIPOS & "' and fecha='" & Format(fecha, "yyyy-mm-dd") & "' and numero='" & Val(numero) & "' "
            csql2.Execute
            If csql2.RowsAffected > 0 Then
                Set resultado2 = csql2.OpenResultset
                If resultado2(0) = "DTE Recibido" Then
                    verificaenviado = True
                Else
                    verificaenviado = False
                End If
                Exit Function
            End If
            resultados.MoveNext
        Wend
    End If
    
    csql.Close
    Set csql = Nothing
    
    
End Function
Sub totallibro2(infogrilla As grillainformes)
    
    Dim TOTALge As Double
      
        lin = lin + 1
        infogrilla.Grid1.Rows = lin + 1
        infogrilla.Grid1.Range(lin, 6, lin, 17).Borders(cellEdgeTop) = cellThin
        infogrilla.Grid1.Cell(lin, 5).text = "TOTALES"
        infogrilla.Grid1.Cell(lin, 6).text = total(1)
        infogrilla.Grid1.Cell(lin, 7).text = total(2)
        infogrilla.Grid1.Cell(lin, 8).text = total(3)
        infogrilla.Grid1.Cell(lin, 9).text = total(4)
        
        infogrilla.Grid1.Cell(lin, 11).text = total(5)
        infogrilla.Grid1.Cell(lin, 12).text = total(6)
        infogrilla.Grid1.Cell(lin, 13).text = total(7)
        infogrilla.Grid1.Cell(lin, 14).text = total(8)
        infogrilla.Grid1.Cell(lin, 15).text = total(9)
        infogrilla.Grid1.Cell(lin, 16).text = total(10)
        infogrilla.Grid1.Cell(lin, 17).text = total(11)
      
    
    TOTALge = 0
    lin = lin + 2
    infogrilla.Grid1.Rows = infogrilla.Grid1.Rows + 2
    infogrilla.Grid1.Rows = infogrilla.Grid1.Rows + 13
    infogrilla.Grid1.Range(lin, 4, lin + 12, 20).Borders(cellEdgeTop) = cellThin
    infogrilla.Grid1.Range(lin, 4, lin + 12, 20).Borders(cellEdgeLeft) = cellThin
    infogrilla.Grid1.Range(lin, 4, lin + 12, 20).Borders(cellEdgeRight) = cellThin
    infogrilla.Grid1.Range(lin, 4, lin + 12, 20).Borders(cellEdgeBottom) = cellThin
    infogrilla.Grid1.Range(lin, 4, lin + 12, 20).Borders(cellInsideHorizontal) = cellThin
    infogrilla.Grid1.Range(lin, 4, lin + 12, 20).Borders(cellInsideVertical) = cellThin
    
    infogrilla.Grid1.Cell(lin, 4).text = "Cant."
    infogrilla.Grid1.Cell(lin, 5).text = "Documentos"
    infogrilla.Grid1.Cell(lin, 6).text = "Neto"
    infogrilla.Grid1.Cell(lin, 7).text = "i.v.a"
    infogrilla.Grid1.Cell(lin, 8).text = "exento"
    infogrilla.Grid1.Cell(lin, 9).text = "total"
    infogrilla.Grid1.Cell(lin, 11).text = "Azucarados"
    infogrilla.Grid1.Cell(lin, 12).text = "Licores"
    infogrilla.Grid1.Cell(lin, 13).text = "Vinos"
    infogrilla.Grid1.Cell(lin, 14).text = "Cerveza"
    infogrilla.Grid1.Cell(lin, 15).text = "Harina"
    infogrilla.Grid1.Cell(lin, 16).text = "Carne"
    infogrilla.Grid1.Cell(lin, 17).text = "Sin Azucar"
    infogrilla.Grid1.Cell(lin, 18).text = "Neto Comision"
    infogrilla.Grid1.Cell(lin, 19).text = "Iva Comision"
    infogrilla.Grid1.Cell(lin, 20).text = "Total Comision"
    
    For k = 1 To 12
        lin = lin + 1
        infogrilla.Grid1.Rows = infogrilla.Grid1.Rows + 1
        infogrilla.Grid1.Cell(lin, 5).text = TIPOS(k)
        infogrilla.Grid1.Cell(lin, 4).text = Format(detalle(k, 1), "###,###,##0")
        infogrilla.Grid1.Cell(lin, 6).text = Format(detalle(k, 2), "###,###,##0")
        infogrilla.Grid1.Cell(lin, 7).text = Format(detalle(k, 3), "###,###,##0")
        infogrilla.Grid1.Cell(lin, 8).text = Format(detalle(k, 4), "###,###,##0")
        infogrilla.Grid1.Cell(lin, 9).text = Format(detalle(k, 5), "###,###,##0")
        infogrilla.Grid1.Cell(lin, 11).text = Format(detalle(k, 6), "###,###,##0")
        infogrilla.Grid1.Cell(lin, 12).text = Format(detalle(k, 7), "###,###,##0")
        infogrilla.Grid1.Cell(lin, 13).text = Format(detalle(k, 8), "###,###,##0")
        infogrilla.Grid1.Cell(lin, 14).text = Format(detalle(k, 9), "###,###,##0")
        infogrilla.Grid1.Cell(lin, 15).text = Format(detalle(k, 10), "###,###,##0")
        infogrilla.Grid1.Cell(lin, 16).text = Format(detalle(k, 11), "###,###,##0")
        infogrilla.Grid1.Cell(lin, 17).text = Format(detalle(k, 12), "###,###,##0")
        infogrilla.Grid1.Cell(lin, 18).text = Format(detalle(k, 13), "###,###,##0")
        infogrilla.Grid1.Cell(lin, 19).text = Format(detalle(k, 14), "###,###,##0")
        infogrilla.Grid1.Cell(lin, 20).text = Format(detalle(k, 15), "###,###,##0")
    
    Next k
    infogrilla.Grid1.Rows = infogrilla.Grid1.Rows + 2
    lin = lin + 2
    For k = 1 To canplan
    If plan(k, 3) <> 0 Then
        lin = lin + 1
        infogrilla.Grid1.Rows = infogrilla.Grid1.Rows + 1
        infogrilla.Grid1.Cell(lin, 4).text = plan(k, 1)
        infogrilla.Grid1.Cell(lin, 5).text = plan(k, 2)
        infogrilla.Grid1.Cell(lin, 6).text = plan(k, 3)
        TOTALge = TOTALge + plan(k, 3)
        End If
    Next k
        lin = lin + 1
        infogrilla.Grid1.Rows = infogrilla.Grid1.Rows + 1
        infogrilla.Grid1.Cell(lin, 5).text = "TOTAL DETALLE"
        infogrilla.Grid1.Range(lin, 6, lin, 6).Borders(cellEdgeTop) = cellThin
        infogrilla.Grid1.Cell(lin, 6).text = TOTALge
               
    End Sub
Sub totallibro(infogrilla As grillainformes)
    
    Dim TOTALge As Double
      
        lin = lin + 1
        infogrilla.Grid1.Rows = lin + 1
        infogrilla.Grid1.Range(lin, 6, lin, 9).Borders(cellEdgeTop) = cellThin
        infogrilla.Grid1.Cell(lin, 5).text = "TOTALES"
        infogrilla.Grid1.Cell(lin, 6).text = total(1)
        infogrilla.Grid1.Cell(lin, 7).text = total(2)
        infogrilla.Grid1.Cell(lin, 8).text = total(3)
        infogrilla.Grid1.Cell(lin, 9).text = total(4)
      
    
    TOTALge = 0
    lin = lin + 2
    infogrilla.Grid1.Rows = infogrilla.Grid1.Rows + 2
    infogrilla.Grid1.Rows = infogrilla.Grid1.Rows + 11
    infogrilla.Grid1.Range(lin, 4, lin + 11, 9).Borders(cellEdgeTop) = cellThin
    infogrilla.Grid1.Range(lin, 4, lin + 11, 9).Borders(cellEdgeLeft) = cellThin
    infogrilla.Grid1.Range(lin, 4, lin + 11, 9).Borders(cellEdgeRight) = cellThin
    infogrilla.Grid1.Range(lin, 4, lin + 11, 9).Borders(cellEdgeBottom) = cellThin
    infogrilla.Grid1.Range(lin, 4, lin + 11, 9).Borders(cellInsideHorizontal) = cellThin
    infogrilla.Grid1.Range(lin, 4, lin + 11, 9).Borders(cellInsideVertical) = cellThin
    
    infogrilla.Grid1.Cell(lin, 4).text = "Cant."
    infogrilla.Grid1.Cell(lin, 5).text = "Documentos"
    infogrilla.Grid1.Cell(lin, 6).text = "Neto"
    infogrilla.Grid1.Cell(lin, 7).text = "i.v.a"
    infogrilla.Grid1.Cell(lin, 8).text = "exento"
    infogrilla.Grid1.Cell(lin, 9).text = "total"
    
    For k = 1 To 11
        lin = lin + 1
        infogrilla.Grid1.Rows = infogrilla.Grid1.Rows + 1
        infogrilla.Grid1.Cell(lin, 5).text = TIPOS(k)
        infogrilla.Grid1.Cell(lin, 4).text = Format(detalle(k, 1), "###,###,##0")
        infogrilla.Grid1.Cell(lin, 6).text = Format(detalle(k, 2), "###,###,##0")
        infogrilla.Grid1.Cell(lin, 7).text = Format(detalle(k, 3), "###,###,##0")
        infogrilla.Grid1.Cell(lin, 8).text = Format(detalle(k, 4), "###,###,##0")
        infogrilla.Grid1.Cell(lin, 9).text = Format(detalle(k, 5), "###,###,##0")
    
    Next k
    infogrilla.Grid1.Rows = infogrilla.Grid1.Rows + 2
    lin = lin + 2
    For k = 1 To canplan
    If plan(k, 3) <> 0 Then
        lin = lin + 1
        infogrilla.Grid1.Rows = infogrilla.Grid1.Rows + 1
        infogrilla.Grid1.Cell(lin, 4).text = plan(k, 1)
        infogrilla.Grid1.Cell(lin, 5).text = plan(k, 2)
        infogrilla.Grid1.Cell(lin, 6).text = plan(k, 3)
        TOTALge = TOTALge + plan(k, 3)
        End If
    Next k
        lin = lin + 1
        infogrilla.Grid1.Rows = infogrilla.Grid1.Rows + 1
        infogrilla.Grid1.Cell(lin, 5).text = "TOTAL DETALLE"
        infogrilla.Grid1.Range(lin, 6, lin, 6).Borders(cellEdgeTop) = cellThin
        infogrilla.Grid1.Cell(lin, 6).text = TOTALge
               
    End Sub
    





Sub CARGAGRILLA(infogrilla As grillainformes)
Rem DATOS DE LA COLUMNA
    infogrilla.Grid1.DefaultFont.Size = 7.5
    
    
    FORMATOGRILLA(1, 1) = "TP"
    FORMATOGRILLA(1, 2) = "NUMERO"
    FORMATOGRILLA(1, 3) = "FECHA"
    FORMATOGRILLA(1, 4) = "RUT"
    FORMATOGRILLA(1, 5) = "PROVEEDOR"
    FORMATOGRILLA(1, 6) = "NETO"
    FORMATOGRILLA(1, 7) = "IVA"
    FORMATOGRILLA(1, 8) = "EXENTO"
    FORMATOGRILLA(1, 9) = "TOTAL"
    FORMATOGRILLA(1, 10) = "FOLIO"
     
    Rem LARGO DE LOS DATOS
    
    FORMATOGRILLA(2, 1) = "3"
    FORMATOGRILLA(2, 2) = "10"
    FORMATOGRILLA(2, 3) = "10"
    FORMATOGRILLA(2, 4) = "10"
    FORMATOGRILLA(2, 5) = "30"
    FORMATOGRILLA(2, 6) = "9"
    FORMATOGRILLA(2, 7) = "9"
    FORMATOGRILLA(2, 8) = "9"
    FORMATOGRILLA(2, 9) = "9"
    FORMATOGRILLA(2, 10) = "4"
    
    
    Rem TIPO DE DATOS C=CEROS,S=STRING,N=NUMERICO,D=DATE,
    FORMATOGRILLA(3, 1) = "S"
    FORMATOGRILLA(3, 2) = "S"
    FORMATOGRILLA(3, 3) = "S"
    FORMATOGRILLA(3, 4) = "S"
    FORMATOGRILLA(3, 5) = "S"
    FORMATOGRILLA(3, 6) = "N"
    FORMATOGRILLA(3, 7) = "N"
    FORMATOGRILLA(3, 8) = "N"
    FORMATOGRILLA(3, 9) = "N"
    FORMATOGRILLA(3, 10) = "S"
    
    Rem FORMATO GRILLA
    FORMATOGRILLA(4, 6) = "###,###,###"
    FORMATOGRILLA(4, 7) = "###,###,###"
    FORMATOGRILLA(4, 8) = "###,###,###"
    FORMATOGRILLA(4, 9) = "###,###,###"
    
    Rem LOCCKED
    For k = 1 To 10
    FORMATOGRILLA(5, k) = "TRUE"
    Next k
    
    infogrilla.Grid1.Cols = 11
    infogrilla.Grid1.Rows = 2
    
     'infogrilla.grid1.AllowUserResizing = False
    infogrilla.Grid1.DisplayFocusRect = False
    'infogrilla.grid1.ExtendLastCol = True
    infogrilla.Grid1.BoldFixedCell = False
    
    infogrilla.Grid1.DrawMode = cellOwnerDraw
    
    infogrilla.Grid1.Appearance = Flat
    infogrilla.Grid1.ScrollBarStyle = Flat
    infogrilla.Grid1.FixedRowColStyle = Flat
    
   'infogrilla.grid1.BackColorFixed = RGB(90, 158, 214)
   ' infogrilla.grid1.BackColorFixedSel = RGB(110, 180, 230)
   ' infogrilla.grid1.BackColorBkg = RGB(90, 158, 214)
   ' infogrilla.grid1.BackColorScrollBar = RGB(231, 235, 247)
   ' infogrilla.grid1.BackColor1 = RGB(231, 235, 247)
   ' infogrilla.grid1.BackColor2 = RGB(239, 243, 255)
   ' infogrilla.grid1.GridColor = RGB(148, 190, 231)
    infogrilla.Grid1.Column(0).Width = 0
    
    For k = 1 To infogrilla.Grid1.Cols - 1
        
        infogrilla.Grid1.Cell(0, k).text = FORMATOGRILLA(1, k)
        infogrilla.Grid1.Column(k).Width = Val(FORMATOGRILLA(2, k)) * infogrilla.Grid1.DefaultFont.Size
        
        
        infogrilla.Grid1.Column(k).MaxLength = Val(FORMATOGRILLA(2, k))
        infogrilla.Grid1.Column(k).FormatString = FORMATOGRILLA(4, k)
        infogrilla.Grid1.Column(k).Locked = FORMATOGRILLA(5, k)
        If FORMATOGRILLA(3, k) = "N" Then infogrilla.Grid1.Column(k).Alignment = cellRightCenter
        If FORMATOGRILLA(3, k) = "D" Then infogrilla.Grid1.Column(k).CellType = cellCalendar
        
    Next k
End Sub
Sub CARGAGRILLA2(infogrilla As grillainformes)
Rem DATOS DE LA COLUMNA
    infogrilla.Grid1.DefaultFont.Size = 7.5
    
    
    FORMATOGRILLA(1, 1) = "TP"
    FORMATOGRILLA(1, 2) = "NUMERO"
    FORMATOGRILLA(1, 3) = "FECHA"
    FORMATOGRILLA(1, 4) = "RUT"
    FORMATOGRILLA(1, 5) = "PROVEEDOR"
    FORMATOGRILLA(1, 6) = "NETO"
    FORMATOGRILLA(1, 7) = "IVA"
    FORMATOGRILLA(1, 8) = "EXENTO"
    FORMATOGRILLA(1, 9) = "TOTAL"
    FORMATOGRILLA(1, 10) = "FOLIO"
    FORMATOGRILLA(1, 11) = "AZUCARADOS"
    FORMATOGRILLA(1, 12) = "LICORES"
    FORMATOGRILLA(1, 13) = "VINOS"
    FORMATOGRILLA(1, 14) = "CERVEZAS"
    FORMATOGRILLA(1, 15) = "HARINA"
    FORMATOGRILLA(1, 16) = "CARNE"
    FORMATOGRILLA(1, 17) = "SIN AZUCAR"
    
    FORMATOGRILLA(1, 18) = "NETO COMISION"
    FORMATOGRILLA(1, 19) = "IVA COMISION"
    FORMATOGRILLA(1, 20) = "TOTAL COMISION"
     
    Rem LARGO DE LOS DATOS
    
    FORMATOGRILLA(2, 1) = "3"
    FORMATOGRILLA(2, 2) = "10"
    FORMATOGRILLA(2, 3) = "10"
    FORMATOGRILLA(2, 4) = "10"
    FORMATOGRILLA(2, 5) = "30"
    FORMATOGRILLA(2, 6) = "9"
    FORMATOGRILLA(2, 7) = "9"
    FORMATOGRILLA(2, 8) = "9"
    FORMATOGRILLA(2, 9) = "9"
    FORMATOGRILLA(2, 10) = "4"
    FORMATOGRILLA(2, 11) = "9"
    FORMATOGRILLA(2, 12) = "9"
    FORMATOGRILLA(2, 13) = "9"
    FORMATOGRILLA(2, 14) = "9"
    FORMATOGRILLA(2, 15) = "9"
    FORMATOGRILLA(2, 16) = "9"
    FORMATOGRILLA(2, 17) = "9"
    
    FORMATOGRILLA(2, 18) = "9"
    FORMATOGRILLA(2, 19) = "9"
    FORMATOGRILLA(2, 20) = "9"
    
    
    Rem TIPO DE DATOS C=CEROS,S=STRING,N=NUMERICO,D=DATE,
    FORMATOGRILLA(3, 1) = "S"
    FORMATOGRILLA(3, 2) = "S"
    FORMATOGRILLA(3, 3) = "S"
    FORMATOGRILLA(3, 4) = "S"
    FORMATOGRILLA(3, 5) = "S"
    FORMATOGRILLA(3, 6) = "N"
    FORMATOGRILLA(3, 7) = "N"
    FORMATOGRILLA(3, 8) = "N"
    FORMATOGRILLA(3, 9) = "N"
    FORMATOGRILLA(3, 10) = "S"
    FORMATOGRILLA(3, 11) = "N"
    FORMATOGRILLA(3, 12) = "N"
    FORMATOGRILLA(3, 13) = "N"
    FORMATOGRILLA(3, 14) = "N"
    FORMATOGRILLA(3, 15) = "N"
    FORMATOGRILLA(3, 16) = "N"
    FORMATOGRILLA(3, 17) = "N"
    
    FORMATOGRILLA(3, 18) = "N"
    FORMATOGRILLA(3, 19) = "N"
    FORMATOGRILLA(3, 20) = "N"
    
    Rem FORMATO GRILLA
    FORMATOGRILLA(4, 6) = "###,###,###"
    FORMATOGRILLA(4, 7) = "###,###,###"
    FORMATOGRILLA(4, 8) = "###,###,###"
    FORMATOGRILLA(4, 9) = "###,###,###"
    FORMATOGRILLA(4, 11) = "###,###,###"
    FORMATOGRILLA(4, 12) = "###,###,###"
    FORMATOGRILLA(4, 13) = "###,###,###"
    FORMATOGRILLA(4, 14) = "###,###,###"
    FORMATOGRILLA(4, 15) = "###,###,###"
    FORMATOGRILLA(4, 16) = "###,###,###"
    FORMATOGRILLA(4, 17) = "###,###,###"
    
    FORMATOGRILLA(4, 18) = "###,###,###"
    FORMATOGRILLA(4, 19) = "###,###,###"
    FORMATOGRILLA(4, 20) = "###,###,###"
    
    Rem LOCCKED
    For k = 1 To 20
    FORMATOGRILLA(5, k) = "TRUE"
    Next k
    
    infogrilla.Grid1.Cols = 21
    infogrilla.Grid1.Rows = 2
    
     'infogrilla.grid1.AllowUserResizing = False
    infogrilla.Grid1.DisplayFocusRect = False
    'infogrilla.grid1.ExtendLastCol = True
    infogrilla.Grid1.BoldFixedCell = False
    
    infogrilla.Grid1.DrawMode = cellOwnerDraw
    
    infogrilla.Grid1.Appearance = Flat
    infogrilla.Grid1.ScrollBarStyle = Flat
    infogrilla.Grid1.FixedRowColStyle = Flat
    
   'infogrilla.grid1.BackColorFixed = RGB(90, 158, 214)
   ' infogrilla.grid1.BackColorFixedSel = RGB(110, 180, 230)
   ' infogrilla.grid1.BackColorBkg = RGB(90, 158, 214)
   ' infogrilla.grid1.BackColorScrollBar = RGB(231, 235, 247)
   ' infogrilla.grid1.BackColor1 = RGB(231, 235, 247)
   ' infogrilla.grid1.BackColor2 = RGB(239, 243, 255)
   ' infogrilla.grid1.GridColor = RGB(148, 190, 231)
    infogrilla.Grid1.Column(0).Width = 0
    
    For k = 1 To infogrilla.Grid1.Cols - 1
        
        infogrilla.Grid1.Cell(0, k).text = FORMATOGRILLA(1, k)
        infogrilla.Grid1.Column(k).Width = Val(FORMATOGRILLA(2, k)) * infogrilla.Grid1.DefaultFont.Size
        
        
        infogrilla.Grid1.Column(k).MaxLength = Val(FORMATOGRILLA(2, k))
        infogrilla.Grid1.Column(k).FormatString = FORMATOGRILLA(4, k)
        infogrilla.Grid1.Column(k).Locked = FORMATOGRILLA(5, k)
        If FORMATOGRILLA(3, k) = "N" Then infogrilla.Grid1.Column(k).Alignment = cellRightCenter
        If FORMATOGRILLA(3, k) = "D" Then infogrilla.Grid1.Column(k).CellType = cellCalendar
        
    Next k
End Sub

Sub leermayor()
    tipoprove = CUENTAPROVEEDOR
    

    
End Sub

'Sub Consultadetalle(MES As String, año As String)
'Dim multi As Integer
'
'Dim resultados2 As rdoResultset
'    Dim cSql2 As New rdoQuery
'        Set cSql2.ActiveConnection = contadb
'        cSql2.SQL = "SELECT cuentadelmayor,dfc.tipo,sum(dfc.monto)"
'        cSql2.SQL = cSql2.SQL + "FROM facturasdecompras as fc,detallefacturasdecompra as dfc "
'        cSql2.SQL = cSql2.SQL + "where añocontable='" + año + "' and mescontable='" + MES + "'"
'        cSql2.SQL = cSql2.SQL + " and fc.tipo=dfc.tipo and fc.numero=dfc.numero and fc.rut=dfc.rut"
'        cSql2.SQL = cSql2.SQL + " group by cuentadelmayor,dfc.tipo "
'
'        cSql2.Execute
'
'
'        If cSql2.RowsAffected > 0 Then
'        Set resultados2 = cSql2.OpenResultset
'
'         While Not resultados2.EOF
'         For K = 1 To canplan
'         If resultados2(1) = "3" Then multi = -1 Else multi = 1
'         If resultados2(0) = plan(K, 1) Then plan(K, 3) = plan(K, 3) + (resultados2(2) * multi): infogrilla.Grid1.Cell(lin, 11).text = plan(K, 2): K = canplan + 1
'         Next K
'          resultados2.MoveNext
'
'
'         Wend
'
'          resultados2.Close
'
'        End If
'
'End Sub
Sub CARGAmayor()
    Dim resultados As rdoResultset
    Dim csql As New rdoQuery
    Dim LINEAS As Double
    
   
        Set csql.ActiveConnection = contadb
        csql.sql = "SELECT codigo,nombre,tipo "
        csql.sql = csql.sql + "FROM cuentasdelmayor where año='" + COMBOAÑO.text + "' "
        csql.sql = csql.sql + " order by codigo"
        csql.Execute
        LINEA = 0
        If csql.RowsAffected > 0 Then
            Set resultados = csql.OpenResultset
             While Not resultados.EOF
             LINEA = LINEA + 1
             plan(LINEA, 1) = resultados(0)
             plan(LINEA, 2) = resultados(1)
             plan(LINEA, 3) = 0

            resultados.MoveNext
            Wend
        End If
canplan = LINEA
   

End Sub

Sub Consultadetalle(tipo, numero, fecha As Date, infogrilla As grillainformes)
Dim multi As Integer

Dim resultados2 As rdoResultset
    Dim csql2 As New rdoQuery
       
        
        Set csql2.ActiveConnection = contadb
        csql2.sql = "SELECT cuentadelmayor,monto "
        csql2.sql = csql2.sql + "FROM facturasdeventas_detalle "
        csql2.sql = csql2.sql + "where tipo='" & tipo & "' and numero='" + numero + "'  order by linea "
        csql2.Execute

        If csql2.RowsAffected > 0 Then
        Set resultados2 = csql2.OpenResultset

        While Not resultados2.EOF
          For k = 1 To canplan
          If tipo = "3" Or tipo = "4" Or tipo = "8" Then multi = -1 Else multi = 1
          If resultados2(0) = plan(k, 1) Then plan(k, 3) = plan(k, 3) + (resultados2(1) * multi)
          Rem  If resultados2(0) = plan(K, 1) And DETALLE1.Value = True Then infogrilla.Grid1.Cell(lin, 10).text = plan(K, 2): K = canplan + 1

          Next k
          resultados2.MoveNext


         Wend

          resultados2.Close

        End If

End Sub

Sub ayudaempresa(ByRef caja As TextBox)
    Dim campos As Variant
    Dim cfijo As Variant
    Dim largo As Variant
    campos = Array("codigoempresa", "nombre")
    largo = Array("6s", "40s")
    cfijo = "no"
    basebus = clientesistema + "conta"
    Call cargaAyudaT(Servidor, basebus, Usuario, password, "maestroempresas", dato1, campos, cfijo, largo, 2)
    caja.Enabled = True
    caja.SetFocus
    leer
End Sub

Sub CARGAcrcc()
    Dim resultados As rdoResultset
    Dim csql As New rdoQuery
    Dim LINEAS As Double
    
   
        Set csql.ActiveConnection = contadb
        csql.sql = "SELECT codigo,nombre "
        csql.sql = csql.sql + "FROM centrosdecosto where año='" + COMBOAÑO.text + "' "
        csql.sql = csql.sql + "order by codigo"
        csql.Execute
        LINEA = 0
        If csql.RowsAffected > 0 Then
            Set resultados = csql.OpenResultset
             While Not resultados.EOF
             LINEA = LINEA + 1
             Combocrcc.AddItem (Mid(resultados(0), 1, 2) + "." + Mid(resultados(0), 3, 2) + " " + resultados(1))
             
            resultados.MoveNext
            Wend
        End If
        Combocrcc.AddItem ("99.99" + " " + "TODOS")
            
        Combocrcc.text = Combocrcc.List(LINEA)
        
   

End Sub

Public Function leerimpuestov(tipo, numero, rut, cuenta)
Dim multi As Integer

Dim resultados2 As rdoResultset
    Dim csql2 As New rdoQuery
        Dim linpaso As Integer
        
        Set csql2.ActiveConnection = contadb
        csql2.sql = "SELECT monto "
        csql2.sql = csql2.sql + "FROM facturasdeventas_detalle "
        csql2.sql = csql2.sql + "where tipo='" & tipo & "' and numero='" + numero + "' and cuentadelmayor='" + cuenta + "' "
        csql2.Execute
        leerimpuestov = 0
        If csql2.RowsAffected > 0 Then
        
        Set resultados2 = csql2.OpenResultset
        linpaso = 0
        While Not resultados2.EOF
          
        leerimpuestov = resultados2(0)
        resultados2.MoveNext
        Wend

        resultados2.Close
        Else
        leerimpuestov = "0"

        End If

End Function

Public Function leedatocomision(tipo, numero, rut, campo)
Dim multi As Integer

Dim resultados2 As rdoResultset
    Dim csql2 As New rdoQuery
        Dim linpaso As Integer
        
        Set csql2.ActiveConnection = contadb
        csql2.sql = "SELECT " & campo & " "
        csql2.sql = csql2.sql + "FROM facturasdeventas_liquidacion "
        csql2.sql = csql2.sql + "where tipo='" & tipo & "' and numero='" + numero + "' and rut='" & rut & "' "
        csql2.Execute
        leedatocomision = 0
        If csql2.RowsAffected > 0 Then
        
        Set resultados2 = csql2.OpenResultset
        linpaso = 0
        While Not resultados2.EOF
          
        leedatocomision = resultados2(0)
        resultados2.MoveNext
        Wend

        resultados2.Close
        Else
        leedatocomision = "0"

        End If

End Function


Private Sub botonmisaccesos_Click()
programafiltro = Me.Caption
misaccesos.Show
End Sub


Private Sub botonmisfavoritos_Click()
Call AgregaFavoritos(USUARIOSISTEMA, App.EXEName, Me.Name, Me.Caption)

End Sub

