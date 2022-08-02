VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash9e.ocx"
Object = "{4777436C-EB8C-4596-98A8-EBCF98683969}#1.0#0"; "Flexcell.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{ADBBAED6-B16F-40DC-80DF-B44910CBA76C}#1.0#0"; "Frame-Xp.ocx"
Begin VB.Form MClientes 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Maestro de Clientes"
   ClientHeight    =   10110
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12585
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10110
   ScaleWidth      =   12585
   Begin XPFrame.FrameXp repa 
      Height          =   1635
      Left            =   360
      TabIndex        =   117
      Top             =   6960
      Width           =   6405
      _ExtentX        =   11298
      _ExtentY        =   2884
      BackColor       =   16761024
      Caption         =   "REPACTACION"
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
      Begin VB.TextBox CUOTASREPACTACION 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2925
         MaxLength       =   3
         TabIndex        =   124
         Top             =   1170
         Width           =   645
      End
      Begin VB.CommandButton cmd40 
         BackColor       =   &H000080FF&
         Caption         =   "Retorno"
         Height          =   375
         Left            =   4320
         Style           =   1  'Graphical
         TabIndex        =   123
         Top             =   1080
         Width           =   1860
      End
      Begin VB.CheckBox lbl200 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         Caption         =   "Repactacion Autorizada"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   3735
         TabIndex        =   122
         Top             =   360
         Width           =   2445
      End
      Begin VB.TextBox CONDO1 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2925
         MaxLength       =   3
         TabIndex        =   119
         Top             =   315
         Width           =   645
      End
      Begin VB.TextBox CONDO2 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2925
         MaxLength       =   3
         TabIndex        =   118
         Top             =   720
         Width           =   645
      End
      Begin VB.Label lbl222 
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "CUOTAS REPACTACION"
         Height          =   330
         Left            =   180
         TabIndex        =   125
         Top             =   1170
         Width           =   2670
      End
      Begin VB.Label lbl220 
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "CONDONACION PRONTO PAGO %"
         Height          =   330
         Left            =   180
         TabIndex        =   121
         Top             =   315
         Width           =   2670
      End
      Begin VB.Label lbl221 
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "CONDONACION INTERES MORA%"
         Height          =   330
         Left            =   180
         TabIndex        =   120
         Top             =   720
         Width           =   2670
      End
   End
   Begin XPFrame.FrameXp FRMTIPO 
      Height          =   3495
      Left            =   5520
      TabIndex        =   109
      Top             =   960
      Width           =   4230
      _ExtentX        =   7461
      _ExtentY        =   6165
      BackColor       =   12648384
      Caption         =   "TIPOS DE CUENTAS"
      CaptionEstilo3D =   1
      BackColor       =   12648384
      ColorBarraArriba=   12648384
      ColorBarraAbajo =   32768
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   1
      Begin FlexCell.Grid TIPOCLIENTE 
         Height          =   3030
         Left            =   45
         TabIndex        =   110
         Top             =   315
         Width           =   4065
         _ExtentX        =   7170
         _ExtentY        =   5345
         DefaultFontSize =   8.25
         Rows            =   5
         SelectionMode   =   1
      End
      Begin MSAdodcLib.Adodc data 
         Height          =   375
         Left            =   5940
         Top             =   1500
         Visible         =   0   'False
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   661
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   3
         LockType        =   3
         CommandType     =   8
         CursorOptions   =   0
         CacheSize       =   50
         MaxRecords      =   0
         BOFAction       =   0
         EOFAction       =   0
         ConnectStringType=   1
         Appearance      =   1
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Orientation     =   0
         Enabled         =   -1
         Connect         =   ""
         OLEDBString     =   ""
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   ""
         Caption         =   "Adodc1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
   End
   Begin XPFrame.FrameXp FRMCARGOS 
      Height          =   5100
      Left            =   4560
      TabIndex        =   93
      Top             =   720
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   8996
      BackColor       =   15967352
      Caption         =   "AGREGAR CARGOS"
      CaptionEstilo3D =   1
      BackColor       =   15967352
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin MSComctlLib.ProgressBar barra 
         Height          =   285
         Left            =   90
         TabIndex        =   116
         Top             =   4725
         Width           =   5775
         _ExtentX        =   10186
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.CommandButton cmd4 
         BackColor       =   &H00C68851&
         Caption         =   "RETORNO"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   495
         Style           =   1  'Graphical
         TabIndex        =   105
         Top             =   4230
         Width           =   2535
      End
      Begin VB.CommandButton cmd5 
         BackColor       =   &H00C68851&
         Caption         =   "GENERAR CARGOS"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   3195
         Style           =   1  'Graphical
         TabIndex        =   100
         Top             =   4230
         Width           =   2535
      End
      Begin VB.TextBox CA2 
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
         Left            =   2205
         MaxLength       =   10
         TabIndex        =   98
         Top             =   1350
         Width           =   2445
      End
      Begin VB.TextBox CA1 
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
         Left            =   2205
         MaxLength       =   30
         TabIndex        =   96
         Top             =   765
         Width           =   4110
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00F3A478&
         Caption         =   "Todos Los Rut"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2790
         TabIndex        =   95
         Top             =   315
         Width           =   1860
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00F3A478&
         Caption         =   "Solo Rut Seleccionado"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   90
         TabIndex        =   94
         Top             =   315
         Width           =   3120
      End
      Begin XPFrame.FrameXp frmmes 
         Height          =   855
         Left            =   1170
         TabIndex        =   101
         Top             =   1935
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
            TabIndex        =   102
            Top             =   360
            Width           =   3855
         End
      End
      Begin XPFrame.FrameXp frmano 
         Height          =   1095
         Left            =   1170
         TabIndex        =   103
         Top             =   2895
         Width           =   4335
         _ExtentX        =   7646
         _ExtentY        =   1931
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
            TabIndex        =   104
            Top             =   360
            Width           =   3855
         End
      End
      Begin VB.Label lbl52 
         BackStyle       =   0  'Transparent
         Caption         =   "MONTO DEL CARGO"
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
         Left            =   135
         TabIndex        =   99
         Top             =   1485
         Width           =   1995
      End
      Begin VB.Label lbl51 
         BackStyle       =   0  'Transparent
         Caption         =   "GLOSA DEL CARGO"
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
         Left            =   135
         TabIndex        =   97
         Top             =   855
         Width           =   2265
      End
   End
   Begin XPFrame.FrameXp FrameXp4 
      Height          =   4335
      Left            =   6975
      TabIndex        =   78
      Top             =   5310
      Width           =   5460
      _ExtentX        =   9631
      _ExtentY        =   7646
      BackColor       =   8454016
      Caption         =   "OBSERVACIONES"
      BackColor       =   8454016
      ForeColor       =   8388608
      ColorBarraAbajo =   12648384
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   1
      ColorTextShadow =   8454016
      Begin VB.CommandButton cmd3 
         BackColor       =   &H000080FF&
         Caption         =   "MODIFICAR OBSERVACION"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1395
         Style           =   1  'Graphical
         TabIndex        =   82
         Top             =   3915
         Visible         =   0   'False
         Width           =   2850
      End
      Begin VB.TextBox observa 
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
         Height          =   3435
         Left            =   180
         MaxLength       =   200
         MultiLine       =   -1  'True
         TabIndex        =   79
         Top             =   360
         Width           =   5100
      End
   End
   Begin XPFrame.FrameXp FrameXp1 
      Height          =   8655
      Left            =   135
      TabIndex        =   13
      Top             =   45
      Width           =   6720
      _ExtentX        =   11853
      _ExtentY        =   15266
      BackColor       =   16744576
      Caption         =   "Datos del Cliente"
      CaptionEstilo3D =   1
      BackColor       =   16744576
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   1
      Begin VB.CheckBox lbl78 
         BackColor       =   &H00FF8080&
         Caption         =   "Constructora"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4080
         TabIndex        =   130
         Top             =   4440
         Width           =   2175
      End
      Begin XPFrame.FrameXp datoscredito 
         Height          =   3375
         Left            =   135
         TabIndex        =   59
         Top             =   5220
         Width           =   6495
         _ExtentX        =   11456
         _ExtentY        =   5953
         BackColor       =   8454016
         Caption         =   "Datos de Crédito"
         CaptionEstilo3D =   1
         BackColor       =   8454016
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Begin VB.CommandButton cmd17 
            BackColor       =   &H0000FF00&
            Caption         =   "CONTRATOS"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   525
            Left            =   4095
            Style           =   1  'Graphical
            TabIndex        =   128
            Top             =   1920
            Width           =   2040
         End
         Begin VB.CheckBox lbl300 
            Appearance      =   0  'Flat
            BackColor       =   &H0080FF80&
            Caption         =   "Carta Cobranza"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   3915
            TabIndex        =   127
            Top             =   360
            Width           =   1905
         End
         Begin VB.CheckBox lbl1000 
            Appearance      =   0  'Flat
            BackColor       =   &H0080FF80&
            Caption         =   "Bloqueado"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   330
            Left            =   2250
            TabIndex        =   115
            Top             =   360
            Width           =   1905
         End
         Begin VB.CommandButton cmd15 
            BackColor       =   &H0000FF00&
            Caption         =   "REPACTACION"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   4095
            Style           =   1  'Graphical
            TabIndex        =   108
            Top             =   1230
            Width           =   2040
         End
         Begin VB.CommandButton cmd6 
            BackColor       =   &H0000FF00&
            Caption         =   "AGREGA CARGOS A CUENTA"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   465
            Left            =   4080
            Style           =   1  'Graphical
            TabIndex        =   92
            Top             =   720
            Width           =   2040
         End
         Begin VB.TextBox dato18 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
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
            Left            =   1845
            MaxLength       =   7
            TabIndex        =   84
            Tag             =   "proveedor"
            Top             =   765
            Width           =   1830
         End
         Begin VB.TextBox dato17 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
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
            Left            =   1485
            MaxLength       =   3
            TabIndex        =   83
            Tag             =   "proveedor"
            Top             =   360
            Width           =   390
         End
         Begin XPFrame.FrameXp frmFinancieros 
            Height          =   345
            Left            =   225
            TabIndex        =   111
            Top             =   2565
            Width           =   2775
            _ExtentX        =   4895
            _ExtentY        =   609
            BackColor       =   49344
            Caption         =   "Datos Financieros"
            CaptionEstilo3D =   1
            BackColor       =   49344
            ColorBarraArriba=   12632319
            ColorBarraAbajo =   128
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Alignment       =   1
         End
         Begin XPFrame.FrameXp frmLaborales 
            Height          =   345
            Left            =   3555
            TabIndex        =   112
            Top             =   2565
            Width           =   2775
            _ExtentX        =   4895
            _ExtentY        =   609
            BackColor       =   49344
            Caption         =   "Datos Laborales"
            CaptionEstilo3D =   1
            BackColor       =   49344
            ColorBarraArriba=   12632319
            ColorBarraAbajo =   128
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Alignment       =   1
         End
         Begin XPFrame.FrameXp frmAdicionales 
            Height          =   345
            Left            =   225
            TabIndex        =   113
            Top             =   2925
            Width           =   2775
            _ExtentX        =   4895
            _ExtentY        =   609
            BackColor       =   49344
            Caption         =   "Datos Adicionales"
            CaptionEstilo3D =   1
            BackColor       =   49344
            ColorBarraArriba=   12632319
            ColorBarraAbajo =   128
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Alignment       =   1
         End
         Begin XPFrame.FrameXp frmCuentas 
            Height          =   345
            Left            =   3555
            TabIndex        =   114
            Top             =   2925
            Width           =   2775
            _ExtentX        =   4895
            _ExtentY        =   609
            BackColor       =   49344
            Caption         =   "Cuentas Adicionales"
            CaptionEstilo3D =   1
            BackColor       =   49344
            ColorBarraArriba=   12632319
            ColorBarraAbajo =   128
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Alignment       =   1
         End
         Begin VB.Label lblSaldo 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
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
            ForeColor       =   &H0080FFFF&
            Height          =   285
            Left            =   1845
            TabIndex        =   65
            Top             =   1485
            Width           =   1830
         End
         Begin VB.Label lblUsado 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
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
            ForeColor       =   &H0080FFFF&
            Height          =   285
            Left            =   1845
            TabIndex        =   64
            Top             =   1125
            Width           =   1830
         End
         Begin VB.Label lbl19 
            Appearance      =   0  'Flat
            BackColor       =   &H00F5C9B1&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " Disponible       ($)"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   90
            TabIndex        =   63
            Top             =   1485
            Width           =   1665
         End
         Begin VB.Label lbl18 
            Appearance      =   0  'Flat
            BackColor       =   &H00F5C9B1&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " Cupo Usado     ($)"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   90
            TabIndex        =   62
            Top             =   1125
            Width           =   1665
         End
         Begin VB.Label lbl17 
            Appearance      =   0  'Flat
            BackColor       =   &H00F5C9B1&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " Cupo Otorgado ($)"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   90
            TabIndex        =   61
            Top             =   765
            Width           =   1665
         End
         Begin VB.Label lbl16 
            Appearance      =   0  'Flat
            BackColor       =   &H00F5C9B1&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " Dia de Pago"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   90
            TabIndex        =   60
            Top             =   360
            Width           =   1335
         End
      End
      Begin VB.CheckBox lbl77 
         BackColor       =   &H00FF8080&
         Caption         =   "Tercera Edad"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2400
         TabIndex        =   126
         Top             =   4440
         Width           =   2175
      End
      Begin VB.TextBox dato16 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
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
         Left            =   1530
         MaxLength       =   9
         TabIndex        =   106
         Tag             =   "proveedor"
         Top             =   4770
         Visible         =   0   'False
         Width           =   1500
      End
      Begin VB.TextBox dato15 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
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
         Locked          =   -1  'True
         MaxLength       =   1
         TabIndex        =   87
         Tag             =   "proveedor"
         Top             =   4365
         Visible         =   0   'False
         Width           =   525
      End
      Begin VB.TextBox dato14 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
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
         Left            =   1440
         MaxLength       =   2
         TabIndex        =   85
         Tag             =   "proveedor"
         Top             =   3960
         Width           =   390
      End
      Begin VB.TextBox dato13 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
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
         Left            =   1455
         MaxLength       =   30
         TabIndex        =   80
         Tag             =   "proveedor"
         Top             =   3645
         Width           =   4890
      End
      Begin VB.TextBox dato1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
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
         Left            =   1440
         MaxLength       =   9
         TabIndex        =   0
         Tag             =   "proveedor"
         Top             =   360
         Width           =   1500
      End
      Begin VB.TextBox dato2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
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
         Left            =   5625
         MaxLength       =   1
         TabIndex        =   1
         Tag             =   "proveedor"
         Top             =   360
         Width           =   495
      End
      Begin VB.TextBox dato3 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
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
         Left            =   1440
         MaxLength       =   50
         TabIndex        =   2
         Tag             =   "proveedor"
         Top             =   720
         Width           =   4890
      End
      Begin VB.TextBox dato4 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
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
         Left            =   1440
         MaxLength       =   50
         TabIndex        =   3
         Tag             =   "proveedor"
         Top             =   1080
         Width           =   4890
      End
      Begin VB.TextBox dato7 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
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
         Left            =   1440
         MaxLength       =   10
         TabIndex        =   6
         Tag             =   "proveedor"
         Top             =   2160
         Width           =   1215
      End
      Begin VB.TextBox dato8 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
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
         Left            =   4080
         MaxLength       =   10
         TabIndex        =   7
         Tag             =   "proveedor"
         Top             =   2160
         Width           =   2250
      End
      Begin VB.TextBox dato9 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
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
         Left            =   1440
         MaxLength       =   10
         TabIndex        =   8
         Tag             =   "proveedor"
         Top             =   2520
         Width           =   1215
      End
      Begin VB.TextBox dato10 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
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
         Left            =   4080
         MaxLength       =   10
         TabIndex        =   9
         Tag             =   "proveedor"
         Top             =   2520
         Width           =   2250
      End
      Begin VB.TextBox dato11 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
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
         Left            =   1440
         MaxLength       =   30
         TabIndex        =   10
         Tag             =   "proveedor"
         Top             =   2880
         Width           =   4890
      End
      Begin VB.TextBox dato12 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
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
         Left            =   1440
         MaxLength       =   30
         TabIndex        =   11
         Tag             =   "proveedor"
         Top             =   3240
         Width           =   4890
      End
      Begin VB.TextBox dato5 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
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
         Left            =   1440
         MaxLength       =   30
         TabIndex        =   4
         Tag             =   "proveedor"
         Top             =   1440
         Width           =   4890
      End
      Begin VB.TextBox dato6 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
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
         Left            =   1440
         MaxLength       =   30
         TabIndex        =   5
         Tag             =   "proveedor"
         Top             =   1800
         Width           =   4890
      End
      Begin VB.Label lblDV2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
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
         ForeColor       =   &H0080FFFF&
         Height          =   285
         Left            =   2970
         TabIndex        =   107
         Top             =   4770
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label LBLTC 
         BackStyle       =   0  'Transparent
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
         Height          =   375
         Left            =   1935
         TabIndex        =   91
         Top             =   3960
         Width           =   3570
      End
      Begin VB.Label lblVendedor 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
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
         ForeColor       =   &H0080FFFF&
         Height          =   285
         Left            =   3690
         TabIndex        =   90
         Top             =   4770
         Visible         =   0   'False
         Width           =   2940
      End
      Begin VB.Label lbl50 
         Appearance      =   0  'Flat
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Vendedor"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   135
         TabIndex        =   89
         Top             =   4770
         Visible         =   0   'False
         Width           =   1380
      End
      Begin VB.Label lbl15 
         Appearance      =   0  'Flat
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Descuento"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   135
         TabIndex        =   88
         Top             =   4365
         Visible         =   0   'False
         Width           =   1365
      End
      Begin VB.Label Lbl14 
         Appearance      =   0  'Flat
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Tipo Cliente"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   135
         TabIndex        =   86
         Top             =   4005
         Width           =   1200
      End
      Begin VB.Label Lbl13 
         Appearance      =   0  'Flat
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " E-mail"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   135
         TabIndex        =   81
         Top             =   3645
         Width           =   1215
      End
      Begin VB.Label lblDV 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
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
         ForeColor       =   &H0080FFFF&
         Height          =   285
         Left            =   2880
         TabIndex        =   27
         Top             =   360
         Width           =   495
      End
      Begin VB.Label lbl8 
         Appearance      =   0  'Flat
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Fono2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   2760
         TabIndex        =   26
         Top             =   2160
         Width           =   1215
      End
      Begin VB.Label lbl7 
         Appearance      =   0  'Flat
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Fono1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   120
         TabIndex        =   25
         Top             =   2160
         Width           =   1215
      End
      Begin VB.Label lbl1 
         Appearance      =   0  'Flat
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Rut"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   120
         TabIndex        =   24
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label lbl3 
         Appearance      =   0  'Flat
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Nombre"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   120
         TabIndex        =   23
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label lbl4 
         Appearance      =   0  'Flat
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Dirección"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   120
         TabIndex        =   22
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label lbl5 
         Appearance      =   0  'Flat
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Comuna"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   120
         TabIndex        =   21
         Top             =   1440
         Width           =   1215
      End
      Begin VB.Label lbl6 
         Appearance      =   0  'Flat
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Ciudad"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   120
         TabIndex        =   20
         Top             =   1800
         Width           =   1215
      End
      Begin VB.Label lbl9 
         Appearance      =   0  'Flat
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Fax"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   120
         TabIndex        =   19
         Top             =   2520
         Width           =   1215
      End
      Begin VB.Label lbl11 
         Appearance      =   0  'Flat
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Giro"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   120
         TabIndex        =   18
         Top             =   2880
         Width           =   1215
      End
      Begin VB.Label lbl12 
         Appearance      =   0  'Flat
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Contacto"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   120
         TabIndex        =   17
         Top             =   3240
         Width           =   1215
      End
      Begin VB.Label lbl2 
         Appearance      =   0  'Flat
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Sucursal Cliente"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   3555
         TabIndex        =   16
         Top             =   360
         Width           =   1890
      End
      Begin VB.Label lbl10 
         Appearance      =   0  'Flat
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Celular"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   2760
         TabIndex        =   15
         Top             =   2520
         Width           =   1215
      End
   End
   Begin MSAdodcLib.Adodc Clientes 
      Height          =   330
      Left            =   120
      Top             =   0
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin XPFrame.FrameXp FrameXp3 
      Height          =   5160
      Left            =   7020
      TabIndex        =   14
      Top             =   225
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   9102
      BackColor       =   8454016
      Caption         =   "Resumen Cuenta"
      CaptionEstilo3D =   1
      BackColor       =   8454016
      ForeColor       =   65535
      ColorBarraArriba=   8454016
      ColorBarraAbajo =   16384
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   1
      Begin VB.CommandButton cmd16 
         BackColor       =   &H0000FF00&
         Caption         =   "EVENTOS COBRANZA"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1920
         Style           =   1  'Graphical
         TabIndex        =   129
         Top             =   4800
         Width           =   2040
      End
      Begin VB.CommandButton Command2 
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
         Left            =   4995
         Picture         =   "MClientes.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   72
         Top             =   4395
         Width           =   315
      End
      Begin VB.CommandButton Command1 
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
         Left            =   4995
         Picture         =   "MClientes.frx":0F72
         Style           =   1  'Graphical
         TabIndex        =   66
         Top             =   3720
         Width           =   315
      End
      Begin VB.CommandButton cmdCheques 
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
         Left            =   4980
         Picture         =   "MClientes.frx":1EE4
         Style           =   1  'Graphical
         TabIndex        =   52
         Top             =   3060
         Width           =   315
      End
      Begin VB.CommandButton cmdFacturas 
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
         Left            =   4980
         Picture         =   "MClientes.frx":2E56
         Style           =   1  'Graphical
         TabIndex        =   51
         Top             =   2460
         Width           =   315
      End
      Begin VB.CommandButton cmdBoletas 
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
         Left            =   4980
         Picture         =   "MClientes.frx":3DC8
         Style           =   1  'Graphical
         TabIndex        =   50
         Top             =   1860
         Width           =   315
      End
      Begin VB.CommandButton cmdProrrogas 
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
         Left            =   4980
         Picture         =   "MClientes.frx":4D3A
         Style           =   1  'Graphical
         TabIndex        =   49
         Top             =   1260
         Width           =   315
      End
      Begin VB.CommandButton cmdProtestos 
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
         Left            =   4980
         Picture         =   "MClientes.frx":5CAC
         Style           =   1  'Graphical
         TabIndex        =   48
         Top             =   660
         Width           =   315
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
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
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   4035
         TabIndex        =   77
         Top             =   4395
         Width           =   915
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
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
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   1215
         TabIndex        =   76
         Top             =   4395
         Width           =   1635
      End
      Begin VB.Label lbl42 
         Appearance      =   0  'Flat
         BackColor       =   &H0000C000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Cantidad"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   2955
         TabIndex        =   75
         Top             =   4395
         Width           =   1035
      End
      Begin VB.Label lbl44 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00008000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "CUOTAS MOROSAS"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   285
         Left            =   135
         TabIndex        =   74
         Top             =   4095
         Width           =   5175
      End
      Begin VB.Label lbl41 
         Appearance      =   0  'Flat
         BackColor       =   &H0000C000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Monto"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   135
         TabIndex        =   73
         Top             =   4395
         Width           =   1035
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
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
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   4035
         TabIndex        =   71
         Top             =   3720
         Width           =   915
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
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
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   1215
         TabIndex        =   70
         Top             =   3720
         Width           =   1635
      End
      Begin VB.Label lbl37 
         Appearance      =   0  'Flat
         BackColor       =   &H0000C000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Cantidad"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   2955
         TabIndex        =   69
         Top             =   3720
         Width           =   1035
      End
      Begin VB.Label lbl43 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00008000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "CUOTAS VIGENTES"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   285
         Left            =   135
         TabIndex        =   68
         Top             =   3420
         Width           =   5175
      End
      Begin VB.Label lbl36 
         Appearance      =   0  'Flat
         BackColor       =   &H0000C000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Monto"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
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
         Top             =   3720
         Width           =   1035
      End
      Begin VB.Label lbl33 
         Appearance      =   0  'Flat
         BackColor       =   &H0000C000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Monto"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   120
         TabIndex        =   57
         Top             =   3060
         Width           =   1035
      End
      Begin VB.Label lbl27 
         Appearance      =   0  'Flat
         BackColor       =   &H0000C000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Monto"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   120
         TabIndex        =   56
         Top             =   1860
         Width           =   1035
      End
      Begin VB.Label lbl30 
         Appearance      =   0  'Flat
         BackColor       =   &H0000C000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Monto"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   120
         TabIndex        =   55
         Top             =   2460
         Width           =   1035
      End
      Begin VB.Label lbl21 
         Appearance      =   0  'Flat
         BackColor       =   &H0000C000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Monto"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   120
         TabIndex        =   54
         Top             =   660
         Width           =   1035
      End
      Begin VB.Label lbl24 
         Appearance      =   0  'Flat
         BackColor       =   &H0000C000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Monto"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   120
         TabIndex        =   53
         Top             =   1260
         Width           =   1035
      End
      Begin VB.Label lbl25 
         Appearance      =   0  'Flat
         BackColor       =   &H0000C000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Cantidad"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   2940
         TabIndex        =   47
         Top             =   1260
         Width           =   1035
      End
      Begin VB.Label lbl23 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00008000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "PRORROGAS"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   285
         Left            =   120
         TabIndex        =   46
         Top             =   960
         Width           =   5175
      End
      Begin VB.Label lbl20 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00008000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "PROTESTOS"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   285
         Left            =   120
         TabIndex        =   45
         Top             =   360
         Width           =   5175
      End
      Begin VB.Label lbl22 
         Appearance      =   0  'Flat
         BackColor       =   &H0000C000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Cantidad"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   2940
         TabIndex        =   44
         Top             =   660
         Width           =   1035
      End
      Begin VB.Label lbl31 
         Appearance      =   0  'Flat
         BackColor       =   &H0000C000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Cantidad"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   2940
         TabIndex        =   43
         Top             =   2460
         Width           =   1035
      End
      Begin VB.Label lbl29 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00008000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "FACTURAS"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   285
         Left            =   120
         TabIndex        =   42
         Top             =   2160
         Width           =   5175
      End
      Begin VB.Label lbl26 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00008000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "BOLETAS"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   285
         Left            =   120
         TabIndex        =   41
         Top             =   1560
         Width           =   5175
      End
      Begin VB.Label lbl28 
         Appearance      =   0  'Flat
         BackColor       =   &H0000C000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Cantidad"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   2940
         TabIndex        =   40
         Top             =   1860
         Width           =   1035
      End
      Begin VB.Label lbl32 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00008000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "CHEQUES EN CARTERA"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   285
         Left            =   120
         TabIndex        =   39
         Top             =   2760
         Width           =   5175
      End
      Begin VB.Label lbl34 
         Appearance      =   0  'Flat
         BackColor       =   &H0000C000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Cantidad"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   2940
         TabIndex        =   38
         Top             =   3060
         Width           =   1035
      End
      Begin VB.Label lblProtestos 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
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
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   1200
         TabIndex        =   37
         Top             =   660
         Width           =   1635
      End
      Begin VB.Label lblCantProtestos 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
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
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   4020
         TabIndex        =   36
         Top             =   660
         Width           =   915
      End
      Begin VB.Label lblProrrogas 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
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
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   1200
         TabIndex        =   35
         Top             =   1260
         Width           =   1635
      End
      Begin VB.Label lblCantProrrogas 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
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
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   4020
         TabIndex        =   34
         Top             =   1260
         Width           =   915
      End
      Begin VB.Label lblBoletas 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
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
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   1200
         TabIndex        =   33
         Top             =   1860
         Width           =   1635
      End
      Begin VB.Label lblCantBoletas 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
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
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   4020
         TabIndex        =   32
         Top             =   1860
         Width           =   915
      End
      Begin VB.Label lblFacturas 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
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
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   1200
         TabIndex        =   31
         Top             =   2460
         Width           =   1635
      End
      Begin VB.Label lblCantFacturas 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
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
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   4020
         TabIndex        =   30
         Top             =   2460
         Width           =   915
      End
      Begin VB.Label lblCheques 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
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
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   1200
         TabIndex        =   29
         Top             =   3060
         Width           =   1635
      End
      Begin VB.Label lblCantCheques 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
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
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   4020
         TabIndex        =   28
         Top             =   3060
         Width           =   915
      End
   End
   Begin VB.PictureBox manual 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   0
      ScaleHeight     =   375
      ScaleWidth      =   555
      TabIndex        =   58
      Top             =   0
      Width           =   555
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash opciones 
      Height          =   1155
      Left            =   180
      TabIndex        =   12
      Top             =   8775
      Width           =   6840
      _cx             =   12065
      _cy             =   2037
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
Attribute VB_Name = "MClientes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Private c As Cliente
    Private c2 As Clientes
    Private modifica As Boolean
    Private cargo As Boolean
        
Private Sub CA1_KeyPress(KeyAscii As Integer)
  KeyAscii = Asc(UCase(Chr(KeyAscii)))
  If KeyAscii = 13 Then
  CA2.SetFocus
  
  End If
  
End Sub

Private Sub CA2_KeyPress(KeyAscii As Integer)
 KeyAscii = esNumero(KeyAscii)
If KeyAscii = 13 And CA2.text <> "" And CA2.text <> "0" Then
cmd5.SetFocus

End If

End Sub

Private Sub lbl78_Click()
If dato1.text <> "" Then
Call modificaconstructora(lbl78.Value, dato1.text + lblDV.Caption)


End If



End Sub

Private Sub cmd16_Click()
If dato1.text <> "" Then
    With tmpvistagestion
        .lblRut.Caption = Format(dato1.text, "###,###,###") & "-" & lblDV.Caption
        .lblNombre.Caption = dato3.text
        .cargavista
        .Show
    End With
End If

End Sub
 
Private Sub cmd3_Click()
If cmd3.Caption = "MODIFICAR OBSERVACION" Then
   observa.Enabled = True
   observa.SetFocus
   cmd3.Caption = "GRABAR OBSERVACION"
   cmd3.Enabled = True
Else
    c.observaciones = observa.text
    Call grabarCliente(c, modifica)
    cmd3.Caption = "MODIFICAR OBSERVACION"
    observa.Enabled = False
End If
 
End Sub

Private Sub cmd4_Click()
FRMCARGOS.Visible = False
dato1.SetFocus
End Sub

Private Sub cmd6_Click()
FRMCARGOS.Visible = True
CA1.SetFocus

End Sub

Private Sub cmd5_Click()
Dim mes As String
Dim año As String
Dim FECHA As String
If CA1.text <> "" And CA2.text <> "" Then
If MsgBox("ESTA SEGURO DE GENERAR LOS CARGOS CORRESPONDIENTES", vbYesNo) = vbYes Then
    
    If Option1.Value = True And (dato14.text = "02" Or dato14.text = "03" Or dato14.text = "04" Or dato14.text = "06" Or dato15.text = "07") Then
    año = COMBOAÑO.text
    mes = Format(COMBOMES.ListIndex + 1, "00")
    
    If mes = "02" And CDbl(dato17.text) >= 28 Then
       FECHA = año + "-" + mes + "-" + "28"
    End If

    If mes = "02" And CDbl(dato17.text) >= 29 And año = "2008" Or año = "2012" Or año = "2016" Then
       FECHA = año + "-" + mes + "-" + "29"
    End If
 
    If mes = "02" And CDbl(dato17.text) < 28 Then
        FECHA = año + "-" + mes + "-" + dato17.text
    End If

    If mes <> "02" Then
        FECHA = año + "-" + mes + "-" + dato17.text
    End If
    
    
      If existecargo(dato1.text + lblDV.Caption, FECHA, CA2.text) = False Then
      
       Call grabarcuotas(dato1.text + lblDV.Caption, FECHA)
     Else
     MsgBox ("cargo para este mes ya generado")
     End If
       
    End If

    If Option2.Value = True Then
    cargatodo
    
    End If

FRMCARGOS.Visible = False

End If

Else
If CA2.text = "" Then CA2.SetFocus
If CA1.text = "" Then CA1.SetFocus

End If

End Sub
Public Sub grabarcuotas(rut, FECHA)
        
        Dim campos(13, 3) As String
        Dim op As Integer
        Dim K As Integer
       
        Set sql = New sqlventas.sqlventa
        campos(0, 0) = "local"
        campos(1, 0) = "tipo"
        campos(2, 0) = "numero"
        campos(3, 0) = "rut"
        campos(4, 0) = "numerocuota"
        campos(5, 0) = "vencimientooriginal"
        campos(6, 0) = "vencimientoactual"
        campos(7, 0) = "montocuota"
        campos(8, 0) = "cantidadcuotas"
        campos(9, 0) = "capitalcuota"
        campos(10, 0) = "montocredito"
        campos(11, 0) = "fechacompra"
        campos(12, 0) = "glosacompra"
        campos(13, 0) = ""
        
        campos(0, 1) = "07"
        campos(1, 1) = "CA"
        campos(2, 1) = Folioingresomanual
        campos(3, 1) = rut
        campos(4, 1) = "1"
        campos(5, 1) = FECHA
        campos(6, 1) = FECHA
        campos(7, 1) = Replace(CA2.text, ".", "")
        campos(8, 1) = "1"
        campos(9, 1) = Replace(CA2.text, ".", "")
        campos(10, 1) = Replace(CA2.text, ".", "")
        campos(11, 1) = Format(fechasistema, "yyyy-mm-dd")
        campos(12, 1) = CA1.text
        
        campos(0, 2) = "sv_cuotas_detalle"
        condicion = ""
        op = 2
        sql.response = campos
        Set sql.conexion = ventas
        Call sql.sqlventas(op, condicion)
        
    End Sub
Private Sub Cmd40_Click()
repa.Visible = False
ctrltostruct
End Sub

Private Sub Cmd17_Click()
contratos.dato6.text = dato1.text
contratos.lblDV.Caption = lblDV.Caption
contratos.lblNombre.Caption = dato3.text
contratos.lblDireccion.Caption = dato4.text
contratos.lblCiudad.Caption = dato6.text
contratos.LBLFONO.Caption = dato7.text
contratos.BTNIMPRIME.Enabled = True

Load contratos
contratos.Show
End Sub

Private Sub CONDO1_GotFocus()
Call cargatexto(CONDO1)

End Sub
Private Sub CONDO2_GotFocus()
Call cargatexto(CONDO2)

End Sub

Private Sub CONDO1_KeyPress(KeyAscii As Integer)
KeyAscii = esNumero(KeyAscii)
If KeyAscii = 13 Then
CONDO2.SetFocus

End If

End Sub

Private Sub CONDO1_LostFocus()
If CONDO1.text = "" Then CONDO1.text = "0"
If CDbl(CONDO1.text) > 100 Then
CONDO1.text = "0"
CONDO1.SetFocus

End If

End Sub
Private Sub CONDO2_LostFocus()
If CONDO2.text = "" Then CONDO2.text = "0"
If CDbl(CONDO2.text) > 100 Then
CONDO2.text = "0"
CONDO2.SetFocus

End If

End Sub


Private Sub CONDO2_KeyPress(KeyAscii As Integer)
KeyAscii = esNumero(KeyAscii)
If KeyAscii = 13 Then
CUOTASREPACTACION.SetFocus

End If

End Sub

Private Sub CUOTASREPACTACION_KeyPress(KeyAscii As Integer)
KeyAscii = esNumero(KeyAscii)
If KeyAscii = 13 And CUOTASREPACTACION.text <> "" And CUOTASREPACTACION.text <> "0" And CUOTASREPACTACION.text < "19" Then
CONDO1.SetFocus

End If

End Sub

'============================================================
'MANEJO DE LOS CONTOLES
'============================================================
    '========================================================
    'GotFocus
    '========================================================
    Private Sub dato1_GotFocus()
        Call VerificarCajas(Me, dato1)
        Call selecciona(dato1)
        Principal.barraEstado.Panels(2).text = "F2: Ayuda Cliente"
    End Sub
    


    Private Sub dato2_GotFocus()
        Call VerificarCajas(Me, dato2)
        Call selecciona(dato2)
    End Sub
    
Private Sub dato2_LostFocus()
observa.Enabled = False

End Sub

    Private Sub dato3_GotFocus()
        Call VerificarCajas(Me, dato3)
        Call selecciona(dato3)
    End Sub
    
    Private Sub dato4_GotFocus()
        Call VerificarCajas(Me, dato4)
        Call selecciona(dato4)
    End Sub
    
    Private Sub dato5_GotFocus()
        Call VerificarCajas(Me, dato5)
        Call selecciona(dato5)
    End Sub
    
    Private Sub dato6_GotFocus()
        Call VerificarCajas(Me, dato6)
        Call selecciona(dato6)
    End Sub
    
    Private Sub dato7_GotFocus()
       Call VerificarCajas(Me, dato7)
        Call selecciona(dato7)
    End Sub
    
    Private Sub dato8_GotFocus()
        Call VerificarCajas(Me, dato8)
        Call selecciona(dato8)
    End Sub
    
    Private Sub dato9_GotFocus()
        Call VerificarCajas(Me, dato9)
        Call selecciona(dato9)
    End Sub
    
    Private Sub dato10_GotFocus()
        Call VerificarCajas(Me, dato10)
        Call selecciona(dato10)
    End Sub
    
    Private Sub dato11_GotFocus()
        Call VerificarCajas(Me, dato11)
        Call selecciona(dato11)
    End Sub
    
    Private Sub dato12_GotFocus()
        Call VerificarCajas(Me, dato12)
        Call selecciona(dato12)
    End Sub
    
    Private Sub dato13_GotFocus()
       Call VerificarCajas(Me, dato13)
       Call selecciona(dato13)
    End Sub
    
    Private Sub dato14_GotFocus()
       FRMTIPO.Visible = True
        Call VerificarCajas(Me, dato14)
        Call selecciona(dato14)
    End Sub
    
    Private Sub dato15_GotFocus()
        Call VerificarCajas(Me, dato15)
        Call selecciona(dato15)
    End Sub
    
    Private Sub dato16_GotFocus()
        Call VerificarCajas(Me, dato16)
        Call selecciona(dato16)
    End Sub
    
    Private Sub dato17_GotFocus()
        Call VerificarCajas(Me, dato17)
        Call selecciona(dato17)
    End Sub
    '========================================================
    'GotFocus
    '========================================================
    
    '========================================================
    'KeyDown
    '========================================================
    Private Sub dato1_KeyDown(KeyCode As Integer, Shift As Integer)
        If KeyCode = vbKeyF2 Then
            Call ayudaCliente(dato1, dato2, lblDV)
        Else
            Call Flechas(KeyCode, dato1)
        End If
    End Sub
    
    Private Sub dato2_KeyDown(KeyCode As Integer, Shift As Integer)
        Call Flechas(KeyCode, dato1)
    End Sub
    
    Private Sub dato3_KeyDown(KeyCode As Integer, Shift As Integer)
        Call Flechas(KeyCode, dato2)
    End Sub
    
    Private Sub dato4_KeyDown(KeyCode As Integer, Shift As Integer)
        Call Flechas(KeyCode, dato3)
    End Sub
    
    Private Sub dato5_KeyDown(KeyCode As Integer, Shift As Integer)
        Call Flechas(KeyCode, dato4)
    End Sub
    
    Private Sub dato6_KeyDown(KeyCode As Integer, Shift As Integer)
        Call Flechas(KeyCode, dato5)
    End Sub
    
    Private Sub dato7_KeyDown(KeyCode As Integer, Shift As Integer)
        Call Flechas(KeyCode, dato6)
    End Sub
    
    Private Sub dato8_KeyDown(KeyCode As Integer, Shift As Integer)
        Call Flechas(KeyCode, dato7)
    End Sub
    
    Private Sub dato9_KeyDown(KeyCode As Integer, Shift As Integer)
        Call Flechas(KeyCode, dato8)
    End Sub
    
    Private Sub dato10_KeyDown(KeyCode As Integer, Shift As Integer)
        Call Flechas(KeyCode, dato9)
    End Sub
    
    Private Sub dato11_KeyDown(KeyCode As Integer, Shift As Integer)
        Call Flechas(KeyCode, dato10)
    End Sub
    
    Private Sub dato12_KeyDown(KeyCode As Integer, Shift As Integer)
        Call Flechas(KeyCode, dato11)
    End Sub
    
    Private Sub dato13_KeyDown(KeyCode As Integer, Shift As Integer)
        Call Flechas(KeyCode, dato12)
    End Sub
    
    Private Sub dato14_KeyDown(KeyCode As Integer, Shift As Integer)
        Call Flechas(KeyCode, dato13)
    End Sub
    
    Private Sub dato15_KeyDown(KeyCode As Integer, Shift As Integer)
        Call Flechas(KeyCode, dato14)
    End Sub
    
    Private Sub dato16_KeyDown(KeyCode As Integer, Shift As Integer)
         If KeyCode = vbKeyF2 Then
            Call ayudaVendedores(dato16)
        Else
            Call Flechas(KeyCode, dato15)
    End If
    End Sub
    
    Private Sub dato17_KeyDown(KeyCode As Integer, Shift As Integer)
        Call Flechas(KeyCode, dato16)
    End Sub
    '========================================================
    'KeyDown
    '========================================================
    
    '========================================================
    'KeyPress
    '========================================================
    Private Sub dato1_KeyPress(KeyAscii As Integer)
        KeyAscii = esNumero(KeyAscii)
        If KeyAscii = 13 Then
            dato1.text = ceros(dato1)
            lblDV.Caption = rut(dato1.text)
            rut_cliente = dato1.text & lblDV.Caption
            SendKeys "{Tab}"
        End If
    End Sub
    
    Private Sub dato2_KeyPress(KeyAscii As Integer)
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If KeyAscii = 13 Then
            dato2.text = ceros(dato2)
            c2.cc.rut = dato1.text & lblDV.Caption
            c2.cc.sucursal = dato2.text
            
            
            If leerCliente(c, dato1.text & lblDV.Caption, dato2.text, "=") = True Then
                Call structtoctrl
                cargo = True
                Call leerClienteadicional(c2, "=")
            Else
            If Verifica_Permiso(Me.Caption, "agrega") = True Then
                Call HabilitarCajas(Me, modifica)
                 SendKeys "{Tab}"
            Else
                MsgBox mensaje_nopermiso, vbCritical + vbOKOnly, "Permiso Denegado"
                dato1.SelStart = 0
                dato1.SelLength = Len(dato1.text)
                dato1.SetFocus
            End If
            
                
            End If
          
           
        End If
    End Sub
    
    Private Sub dato3_KeyPress(KeyAscii As Integer)
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If KeyAscii = 13 Then
            SendKeys "{Tab}"
        End If
    End Sub
    
    Private Sub dato4_KeyPress(KeyAscii As Integer)
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If KeyAscii = 13 Then
            SendKeys "{Tab}"
        End If
    End Sub
    
    Private Sub dato5_KeyPress(KeyAscii As Integer)
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If KeyAscii = 13 Then
            SendKeys "{Tab}"
        End If
    End Sub
    
    Private Sub dato6_KeyPress(KeyAscii As Integer)
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If KeyAscii = 13 Then
            SendKeys "{Tab}"
        End If
    End Sub
    
    Private Sub dato7_KeyPress(KeyAscii As Integer)
      
        If KeyAscii = 13 Then
            SendKeys "{Tab}"
        End If
    End Sub
    
    Private Sub dato8_KeyPress(KeyAscii As Integer)
      
        If KeyAscii = 13 Then
            SendKeys "{Tab}"
        End If
    End Sub
    
    Private Sub dato9_KeyPress(KeyAscii As Integer)
      
        If KeyAscii = 13 Then
            SendKeys "{Tab}"
        End If
    End Sub
    
    
    Private Sub dato10_KeyPress(KeyAscii As Integer)
       
        If KeyAscii = 13 Then
            SendKeys "{Tab}"
        End If
    End Sub
    
    Private Sub dato11_KeyPress(KeyAscii As Integer)
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If KeyAscii = 13 Then
            SendKeys "{Tab}"
        End If
    End Sub
    
    Private Sub dato12_KeyPress(KeyAscii As Integer)
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If KeyAscii = 13 Then
            
            dato13.SetFocus
            'SendKeys "{Tab}"
            
        End If
       
    End Sub
    
    Private Sub dato13_KeyPress(KeyAscii As Integer)
     KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If KeyAscii = 13 Then
            'SendKeys "{Tab}"
            dato14.SetFocus
        End If
    End Sub
    
    Private Sub dato14_KeyPress(KeyAscii As Integer)
        KeyAscii = esNumero(KeyAscii)
        If KeyAscii = 13 Then
        If dato1.text <> "08" Then dato15.text = "0"
        If leertipocli(dato14.text) <> "" Then
        LBLTC.Caption = leertipocli(dato14.text)
          
          
          If escredito(dato14.text) = True Then
          datoscredito.Visible = True
          dato17.SetFocus
          
          Else
        
                 If dato14.text = "08" Then
                    lbl15.Visible = True
                    dato15.Visible = True
                    dato15.text = "2"
                    dato15.Enabled = True
                    dato15.SetFocus
                Else
         
                  Call ctrltostruct
                End If
                
          End If
        Else
        
        
                    dato14.SetFocus
           
            
        End If
    End If
    
    End Sub
    
    Private Sub dato15_KeyPress(KeyAscii As Integer)
        KeyAscii = esNumero(KeyAscii)
     
        If KeyAscii = 13 Then
         Call ctrltostruct
        End If
        
    End Sub
    
    Private Sub dato16_KeyPress(KeyAscii As Integer)
        KeyAscii = esNumero(KeyAscii)
          
        If KeyAscii = 13 And dato16 <> "" Then
        If dato14.text = "S" Or dato14.text = "T" Or dato14.text = "M" Then
            'SendKeys "{Tab}"
            dato16.text = ceros(dato16)
            lblDV2.Caption = rut(dato16.text)
            lblVendedor.Caption = leerNombreVendedor(dato16.text & lblDV2.Caption)
            If lblVendedor.Caption <> "" Then
            dato17.SetFocus
            Else
            dato16.text = ""
            dato16.SetFocus
            End If
            
            End If
        End If
        If KeyAscii = 13 And dato14.text <> "S" And dato14.text <> "T" And dato14.text <> "M" Then
        Call ctrltostruct
        End If
    End Sub
    
    Private Sub dato17_KeyPress(KeyAscii As Integer)
        KeyAscii = esNumero(KeyAscii)
        If KeyAscii = 13 And dato17.text > "0" And dato17.text < "31" Then
            dato18.SetFocus
        End If
    End Sub
Private Sub dato18_KeyPress(KeyAscii As Integer)
        KeyAscii = esNumero(KeyAscii)
        If KeyAscii = 13 And dato18.text <> "" And dato18.text <> "0" Then
            If dato14.text = "06" Then
           Rem CALLl modificadueñacasa(dato18.text)
            
            End If
            
            Call ctrltostruct
        End If
    End Sub
    
    '========================================================
    'KeyPress
    '========================================================
    
    '========================================================
    'LostFocus
    '========================================================
    Private Sub dato1_LostFocus()
        Call limpiaBarra(2)
    End Sub
    
    Private Sub dato14_LostFocus()
        FRMTIPO.Visible = False
    End Sub
    '========================================================
    'LostFocus
    '========================================================
'============================================================
'MANEJO DE LOS CONTOLES
'============================================================

    Private Sub Form_Activate()
        Principal.barraEstado.Panels(1).text = UCase(Me.Caption)
        sqlventas.audit = True: sqlventas.programaactivo = Me.Caption
           
        
        'If segurity = True Then
        '    Seguridad.Show vbModal
            'segurity = False
        'End If
    End Sub
    
    Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
        If KeyCode = 27 Then
            Unload Me
        End If
        If KeyCode = 38 Then
            If Screen.ActiveForm.ActiveControl.Name = "dato1" Then
                Unload Me
            End If
        End If
    End Sub

    Private Sub Form_Load()
        Dim K As Integer
        repa.Visible = False
        
        'segurity = Not Verificar(usuarioSistema, passwordSistema)
        modifica = False
        cargo = False
        Call Centrar(Me)
        observa.Enabled = False
        For K = 1 To 12
    COMBOMES.AddItem MonthName(K)
    Next K
    COMBOMES.ListIndex = CDbl(Format(fechasistema, "mm") - 1)
    For K = 2000 To Val(Format(fechasistema, "yyyy")) + 2
    COMBOAÑO.AddItem K
    Next K
    COMBOAÑO.ListIndex = K - 2003
      FRMCARGOS.Visible = False
      Option1.Value = True
      CA1.text = ""
      CA2.text = ""
    FRMTIPO.Visible = False
    TIPOCLIENTE.Cols = 3
    TIPOCLIENTE.Rows = 1
    TIPOCLIENTE.Column(1).Width = 40
    TIPOCLIENTE.Column(2).Width = 300
    TIPOCLIENTE.Cell(0, 1).text = "CODIGO"
    TIPOCLIENTE.Cell(0, 2).text = "DETALLE"
    TIPOCLIENTE.Column(0).Width = 0
    datoscredito.Visible = False
    lbl1000.Value = "0"
    
    Call leertiposdeclientes(Me, TIPOCLIENTE)
    End Sub
    
    Private Sub Form_Unload(Cancel As Integer)
        Principal.barraEstado.Panels(1).text = UCase(Principal.Caption)
        Call limpiaBarra(2)
    End Sub

'=============================================================================
'PASA LOS DATOS DE LOS CONTROLES A LA ESTRUCTURA
'=============================================================================
    Private Sub ctrltostruct()
        c.rut = dato1.text & lblDV.Caption
        c.sucursal = dato2.text
        c.nombre = dato3.text
        c.direccion = dato4.text
        c.comuna = dato5.text
        c.ciudad = dato6.text
        c.fono1 = dato7.text
        c.fono2 = dato8.text
        c.fax = dato9.text
        c.celular = dato10.text
        c.giro = dato11.text
        c.contacto = dato12.text
        c.email = dato13.text
       'c.plazo = 0
        c.TIPOCLIENTE = dato14.text
        ' c.listaprecios = dato15.text
        c.DIAPAGO = dato17.text
        c.Descuento = dato15.text
        c.cupodirecto = dato18.text
        c.vendedor = dato16.text & lblDV2.Caption
        c.bloqueoTMP = lbl1000.Value
        c.rebajainterespp = CONDO1.text
        c.rebajainteresmora = CONDO2.text
        c.repactacion = lbl200.Value
        c.CUOTASREPACTACION = CUOTASREPACTACION.text
        c.terceraedad = lbl77.Value
        If observa.text = "" Then
        c.observaciones = "Sin observaciones"
        End If
               
        Call grabarCliente(c, modifica)
        Call retorno
    End Sub
    
      Private Sub strcttoctrl_personales()
        With MCDatosAdicionales
            .dato1.text = Format(c2.cp.fechanacimiento, "dd")
            .dato2.text = Format(c2.cp.fechanacimiento, "mm")
            .dato3.text = Format(c2.cp.fechanacimiento, "yyyy")
            .dato4.text = c2.cp.sexo
            If .dato4.text = "M" Then
                .lblSexo.Caption = "MASCULINO"
            End If
            If .dato4.text = "F" Then
                .lblSexo.Caption = "FEMENINO"
            End If
            .dato5.text = c2.cp.nacionalidad
            .dato6.text = c2.cp.estadocivil
            .dato7.text = c2.cp.rutconyuge
            .lblDV.Caption = rut(.dato7.text)
            .dato8.text = c2.cp.nombreconyuge
        End With
    End Sub
    
    Private Sub strcttoctrl_laborales()
        With MCDatosLaborales
            .dato1.text = c2.cl.rutempleador
            .lblDV.Caption = rut(.dato1.text)
            .dato2.text = c2.cl.nombre
            .dato3.text = c2.cl.direccion
            .dato4.text = c2.cl.comuna
            .dato5.text = c2.cl.ciudad
            .dato6.text = c2.cl.fono
            .dato7.text = c2.cl.labor
            .dato8.text = c2.cl.antiguedad
            If c2.cl.codeudor = "" Then c2.cl.codeudor = 0
            .chkCodeudor.Value = c2.cl.codeudor
        End With
    End Sub
    
    Private Sub strcttoctrl_financieros()
        With MCDatosFinancieros
            .dato1.text = c2.cf.ingresomensual
            .dato2.text = c2.cf.pagoscasascomerciales
            .dato3.text = c2.cf.tipovivienda
            .dato4.text = c2.cf.tasacionvivienda
            .dato5.text = c2.cf.arriendo
            .dato6.text = c2.cf.vehiculos
            .dato7.text = c2.cf.tasacionvehiculos
            .dato8.text = c2.cf.cuentacorriente
            .dato9.text = c2.cf.Banco
            .dato10.text = c2.cf.numerocuenta
            .dato11.text = c2.cf.antuguedad
            .dato12.text = c2.cf.otrastarjetas
            .dato13.text = c2.cf.otratarjeta1
            .dato14.text = c2.cf.otratarjetacupo1
            .dato15.text = c2.cf.otratarjeta2
            .dato16.text = c2.cf.otratarjetacupo2
            .dato17.text = c2.cf.otratarjeta3
            .dato18.text = c2.cf.otratarjetacupo3
            If c2.cf.otratarjetacupo1 = "" Then c2.cf.otratarjetacupo1 = 0: If c2.cf.otratarjetacupo2 = "" Then c2.cf.otratarjetacupo2 = 0: If c2.cf.otratarjetacupo3 = "" Then c2.cf.otratarjetacupo3 = 0
            .lblTotal.Caption = Format(CDbl(c2.cf.otratarjetacupo1) + CDbl(c2.cf.otratarjetacupo2) + CDbl(c2.cf.otratarjetacupo3), "$ ###,###,##0")
            
            .lblDiaImpresion.Caption = Format(c2.cf.fechaimpresionpagare, "dd")
            .lblMesImpresion.Caption = Format(c2.cf.fechaimpresionpagare, "mm")
            .lblAñoImpresion.Caption = Format(c2.cf.fechaimpresionpagare, "yyyy")
            
            .lblDiaCredito = Format(c2.cf.fechaautorizacioncredito, "dd")
            .lblMesCredito = Format(c2.cf.fechaautorizacioncredito, "mm")
            .lblAñoCredito = Format(c2.cf.fechaautorizacioncredito, "yyyy")
            
            .lblAutorizador.Caption = c2.cf.autorizador
            
            .lblDiaTarjeta.Caption = Format(c2.cf.fechaentregatarjeta, "dd")
            .lblMesTarjeta.Caption = Format(c2.cf.fechaentregatarjeta, "mm")
            .lblAñoTarjeta.Caption = Format(c2.cf.fechaentregatarjeta, "yyyy")
        End With
    End Sub
'=============================================================================
'PASA LOS DATOS DE LOS CONTROLES A LA ESTRUCTURA
'=============================================================================

'=============================================================================
'PASA LOS DATOS DE LA ESTRUCTURA A LOS CONTROLES
'=============================================================================
    Private Sub structtoctrl()
        Dim cad As String
        Dim cadena As String
        cad = c.rut
        dato1.text = Left(cad, Len(cad) - 1)
        lblDV.Caption = Right(cad, 1)
        dato2.text = c.sucursal
        dato3.text = c.nombre
        dato4.text = c.direccion
        dato5.text = c.comuna
        dato6.text = c.ciudad
        dato7.text = c.fono1
        dato8.text = c.fono2
        dato9.text = c.fax
        dato10.text = c.celular
        dato11.text = c.giro
        dato12.text = c.contacto
        dato13.text = c.email
        'dato13.text = c.plazo
        dato14.text = c.TIPOCLIENTE
        LBLTC.Caption = leertipocli(dato14.text)
        lbl77.Value = c.terceraedad
        lbl78.Value = c.constructora
        
        If escredito(dato14.text) = True Then
        datoscredito.Visible = True
        lbl1000.Value = c.bloqueoTMP
        lbl200.Value = c.repactacion
        CONDO1.text = c.rebajainterespp
        CONDO2.text = c.rebajainteresmora
        CUOTASREPACTACION.text = c.CUOTASREPACTACION
        
        
        End If
        
        
        
        
        dato15.text = c.Descuento
        
        dato16.text = c.vendedor
        If dato14.text = "08" Then
        dato15.Visible = True
        lbl15.Visible = True
        dato15.text = c.Descuento
        
        End If
        
        lblDV2.Caption = Mid(c.vendedor, 10, 1)
        dato17.text = c.DIAPAGO
        dato18.text = c.cupodirecto
        observa.text = c.observaciones
        lblVendedor.Caption = leerNombreVendedor(dato16.text & lblDV2.Caption)
        If c.CREDITO = "S" Then LBLTC.Caption = "CREDITO DIRECTO"
        If c.CREDITO = "T" Then LBLTC.Caption = "CREDITO TMP"
        If c.CREDITO = "B" Then LBLTC.Caption = "BLOQUEADO"
        If c.CREDITO = "N" Then LBLTC.Caption = "            "
        cadena = leerPesosProtesto(dato1.text & lblDV.Caption, dato2.text)
        cad = Left(cadena, InStr(1, cadena, "/", vbBinaryCompare) - 1)
        lblProtestos.Caption = Format(cad, "$ ###,###,##0")
        cad = Right(cadena, Len(cadena) - InStr(1, cadena, "/", vbBinaryCompare))
        lblCantProtestos.Caption = cad
        
        cadena = leerPesosProrroga(dato1.text & lblDV.Caption, dato2.text)
        cad = Left(cadena, InStr(1, cadena, "/", vbBinaryCompare) - 1)
        lblProrrogas.Caption = Format(cad, "$ ###,###,##0")
        cad = Right(cadena, Len(cadena) - InStr(1, cadena, "/", vbBinaryCompare))
        lblCantProrrogas.Caption = cad
        
        cadena = leerPesosBoleta(dato1.text & lblDV.Caption, dato2.text)
        cad = Left(cadena, InStr(1, cadena, "/", vbBinaryCompare) - 1)
        lblBoletas.Caption = Format(cad, "$ ###,###,##0")
        cad = Right(cadena, Len(cadena) - InStr(1, cadena, "/", vbBinaryCompare))
        lblCantBoletas.Caption = cad
        
        cadena = leerPesosFactura(dato1.text & lblDV.Caption, dato2.text)
        cad = Left(cadena, InStr(1, cadena, "/", vbBinaryCompare) - 1)
        lblFacturas.Caption = Format(cad, "$ ###,###,##0")
        cad = Right(cadena, Len(cadena) - InStr(1, cadena, "/", vbBinaryCompare))
        lblCantFacturas.Caption = cad
        
        'revisar
        cadena = leerPesosCheque(dato1.text & lblDV.Caption, dato2.text)
'        cad = Left(cadena, InStr(1, cadena, "/", vbBinaryCompare) - 1)
'        lblCheques.Caption = Format(cad, "$ ###,###,##0")
'        cad = Right(cadena, Len(cadena) - InStr(1, cadena, "/", vbBinaryCompare))
'        lblCantCheques.Caption = cad
        'revisar
        
        Call DeshabilitarCajas(Me)
    End Sub
'=============================================================================
'PASA LOS DATOS DE LA ESTRUCTURA A LOS CONTROLES
'=============================================================================

Private Sub frmPrecios_BarClick()


    Load Especiales
    Especiales.lblRut.Caption = dato1.text
    Especiales.lblDV.Caption = lblDV.Caption
    Especiales.suc = dato2.text
    Especiales.lblNombre.Caption = dato3.text
    Especiales.Show vbModal
End Sub

Private Sub frmPrecios_BarMouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
'
'    frmPrecios.CaptionEstilo3D = Raised
End Sub

'===============================================================
'PROTESTOS
'===============================================================
    Private Sub cmdProtestos_Click()
        Dim tabla As String
        If cargo = True Then
            tabla = "SELECT CONCAT(cheque, '" & vbTab & "', IFNULL(DATE_FORMAT(fechacheque,'%d-%m-%Y'),''), '" & vbTab & "', IFNULL(DATE_FORMAT(fechaprotesto,'%d-%m-%Y'),''), '" & vbTab & "', CONCAT('$ ', FORMAT(monto,0)), '" & vbTab & "', IFNULL(DATE_FORMAT(cancelado,'%d-%m-%Y'),'')) AS item "
            tabla = tabla & "FROM sv_protesto_" & empresaActiva & " "
            tabla = tabla & "WHERE rut = '" & rut_cliente & "' "
            tabla = tabla & "ORDER BY cheque ASC"
            Load listaDocumentos
            listaDocumentos.formulario = "clientes"
            listaDocumentos.datos = "protestos"
            listaDocumentos.tabla = tabla
            listaDocumentos.Show vbModal
        End If
    End Sub
'===============================================================
'PROTESTOS
'===============================================================

'===============================================================
'PRORROGAS
'===============================================================
    Private Sub cmdProrrogas_Click()
        Dim tabla As String
        If cargo = True Then
            tabla = "SELECT CONCAT(numero, '" & vbTab & "', cheque, '" & vbTab & "', IFNULL(DATE_FORMAT(fcheque,'%d-%m-%Y'),''), '" & vbTab & "', CONCAT('$ ', FORMAT(monto,0)), '" & vbTab & "', IFNULL(DATE_FORMAT(fprorroga,'%d-%m-%Y'),'')) AS item "
            tabla = tabla & "FROM sv_prorroga_" & empresaActiva & " "
            tabla = tabla & "WHERE rut = '" & rut_cliente & "' "
            tabla = tabla & "ORDER BY numero ASC"
            Load listaDocumentos
            listaDocumentos.formulario = "clientes"
            listaDocumentos.datos = "prorrogas"
            listaDocumentos.tabla = tabla
            listaDocumentos.Show vbModal
        End If
    End Sub
'===============================================================
'PRORROGAS
'===============================================================

'===============================================================
'BOLETAS
'===============================================================
    Private Sub cmdBoletas_Click()
        Dim TIPO As String
        Dim tabla As String
        If cargo = True Then
            TIPO = "BV"
            tabla = "SELECT CONCAT(CONCAT(dc.tipo, ' ', dc.numero), '" & vbTab & "', DATE_FORMAT(dc.fecha,'%d-%m-%Y'), '" & vbTab & "', dc.rut) AS item1, CONCAT(CONCAT('$ ', FORMAT(dc.descuento,0)), '" & vbTab & "', CONCAT('$ ', FORMAT(dc.neto,0)), '" & vbTab & "', CONCAT('$ ', FORMAT(dc.iva,0)), '" & vbTab & "', CONCAT('$ ', FORMAT(IF(dc.retencionparcial=0,dc.retenciontotal,dc.retencionparcial),0)), '" & vbTab & "', CONCAT('$ ', FORMAT(dc.total,0))) AS item2, dc.rut, dc.descuento, dc.neto, dc.iva, IF(dc.retencionparcial=0, dc.retenciontotal, dc.retencionparcial) AS retencion, dc.total, dc.tipo "
            tabla = tabla & "FROM sv_documento_cabeza_" + empresaActiva + " AS dc "
            tabla = tabla & "WHERE local = '" & empresaActiva & "' AND tipo = '" & TIPO & "' AND nula = 'N' AND rut = '" & rut_cliente & "' AND sucursal = '" & dato2.text & "' ORDER BY dc.numero ASC"
            Load listaDocumentos
            listaDocumentos.formulario = "clientes"
            listaDocumentos.datos = "ventas"
            listaDocumentos.tabla = tabla
            listaDocumentos.Show vbModal
        End If
    End Sub
'===============================================================
'BOLETAS
'===============================================================

'===============================================================
'FACTURAS
'===============================================================
    Private Sub cmdFacturas_Click()
        Dim TIPO As String
        Dim tabla As String
        If cargo = True Then
            TIPO = "FV"
            tabla = "SELECT CONCAT(CONCAT(dc.tipo, ' ', dc.numero), '" & vbTab & "', DATE_FORMAT(dc.fecha,'%d-%m-%Y'), '" & vbTab & "', dc.rut) AS item1, CONCAT(CONCAT('$ ', FORMAT(dc.descuento,0)), '" & vbTab & "', CONCAT('$ ', FORMAT(dc.neto,0)), '" & vbTab & "', CONCAT('$ ', FORMAT(dc.iva,0)), '" & vbTab & "', CONCAT('$ ', FORMAT(IF(dc.retencionparcial=0,dc.retenciontotal,dc.retencionparcial),0)), '" & vbTab & "', CONCAT('$ ', FORMAT(dc.total,0))) AS item2, dc.rut, dc.descuento, dc.neto, dc.iva, IF(dc.retencionparcial=0, dc.retenciontotal, dc.retencionparcial) AS retencion, dc.total, dc.tipo "
            tabla = tabla & "FROM sv_documento_cabeza_" + empresaActiva + " AS dc "
            tabla = tabla & "WHERE local = '" & empresaActiva & "' AND tipo = '" & TIPO & "' AND nula = 'N' AND rut = '" & rut_cliente & "' AND sucursal = '" & dato2.text & "' ORDER BY dc.numero ASC"
            Load listaDocumentos
            listaDocumentos.formulario = "clientes"
            listaDocumentos.datos = "ventas"
            listaDocumentos.tabla = tabla
            listaDocumentos.Show vbModal
        End If
    End Sub
'===============================================================
'FACTURAS
'===============================================================

'===============================================================
'CHEQUES EN CARTERA
'===============================================================
    Private Sub cmdCheques_Click()
        Dim tabla As String
        If cargo = True Then
            tabla = "SELECT CONCAT(c.numerocheque, '" & vbTab & "', c.banco,' ', IFNULL(mb.nombre,'BANCO NO ENCONTRADO'), '" & vbTab & "', IFNULL(DATE_FORMAT(c.fecharecepcion,'%d-%m-%Y'),''), '" & vbTab & "', CONCAT('$ ', FORMAT(c.monto,0)), '" & vbTab & "', IFNULL(DATE_FORMAT(c.fechavencimiento,'%d-%m-%Y'),''), '" & vbTab & "', c.tipodocumento, ' ', c.numero) AS item "
            tabla = tabla & "FROM sv_carteracheques AS c LEFT JOIN " & baseVentas & ".sv_maestrobancos AS mb ON c.banco = mb.codigobanco "
            tabla = tabla & "WHERE rut = '" & rut_cliente & "' AND sucursal = '" & dato2.text & "' AND c.fechavencimiento >= '" & fechasistema & "' "
            tabla = tabla & "ORDER BY c.fechavencimiento ASC"
            Load listaDocumentos
            listaDocumentos.formulario = "clientes"
            listaDocumentos.datos = "cheques"
            listaDocumentos.tabla = tabla
            listaDocumentos.Show vbModal
        End If
    End Sub
'===============================================================
'CHEQUES EN CARTERA
'===============================================================







Private Sub lbl39_Click()

End Sub

Private Sub lbl40_Click()
End Sub

Private Sub manual_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 27, Asc("r"), Asc("R")
            Call retorno
        Case Asc("a"), Asc("A"), 37
            Call anterior
        Case Asc("s"), Asc("S"), 39
            Call siguiente
        Case Asc("m"), Asc("M")
            Call modificar
        Case Asc("e"), Asc("E"), 46
            Call ELIMINAR
    End Select
End Sub

Private Sub observa_KeyPress(KeyAscii As Integer)
 KeyAscii = Asc(UCase(Chr(KeyAscii)))
 
End Sub

'=============================================================================
'OPCIONES
'=============================================================================
    Private Sub opciones_FSCommand(ByVal command As String, ByVal args As String)
        Select Case command
            Case "modifica"
             If Verifica_Permiso(Me.Caption, "modifica") = True Then
                Call modificar
             Else
                MsgBox mensaje_nopermiso, vbCritical + vbOKOnly, "Permiso Denegado"
             End If
            Case "elimina"
            If Verifica_Permiso(Me.Caption, "elimina") = True Then
                If MsgBox("DESEA REALMENTE ELIMINAR Si / No", vbYesNo) = vbYes Then
                   Call ELIMINAR
                End If
            Else
                MsgBox mensaje_nopermiso, vbCritical + vbOKOnly, "Permiso Denegado"
            End If
            Case "imprime"
            LClientes.Show
            
            Case "movimientos"
                tmplistado7.rut1.text = dato1.text
                tmplistado7.lblDV.Caption = lblDV.Caption
                tmplistado7.CARGARDESDEAFUERAtmp
                tmplistado7.Show
                
            Case "historico"
            LibroVentasclientes.rut1.text = dato1.text
            LibroVentasclientes.lblDV.Caption = lblDV.Caption
            LibroVentasclientes.CARGARDESDEAFUERA
            
             LibroVentasclientes.Show
            Case "retorno"
                 If modifica = True Then
                 Call ctrltostruct
                End If
                Call retorno
            Case "anterior"
                Call anterior
            Case "siguiente"
                Call siguiente
        End Select
    End Sub
    
    Private Sub modificar()
        modifica = True
        Call HabilitarCajas(Me, modifica)
        dato1.Enabled = False
        dato2.Enabled = False
        dato3.SetFocus
        cmd3.Visible = True
                
    End Sub
    
    Private Sub ELIMINAR()
        frmglosaeliminacion.Show vbModal
        
        
        Call eliminarCliente(c)
        
        
        Call retorno
        Call HabilitarCajas(Me, modifica)
        dato1.SetFocus
    End Sub

    Private Sub retorno()
        Call LimpiarCajas(MClientes)
        Call LimpiarLabels(MClientes)
        
        modifica = False
        cargo = False
        Call DeshabilitarCajas(Me)
        cmd3.Visible = False
        dato1.SetFocus
        FRMTIPO.Visible = False
        datoscredito.Visible = False
        lbl1000.Value = "0"
        CONDO1.text = "0"
        CONDO2.text = "0"
        lbl200.Value = "0"
        CUOTASREPACTACION.text = "0"
        lbl77.Value = 0
        lbl15.Visible = False
        dato15.Visible = False
        lbl78.Value = 0
    End Sub
    
    
    Private Sub anterior()
        If leerCliente(c, dato1.text & lblDV.Caption, dato2.text, "<") = True Then
            structtoctrl
        End If
    End Sub
    
    Private Sub siguiente()
        If leerCliente(c, dato1.text & lblDV.Caption, dato2.text, ">") = True Then
            structtoctrl
        End If
    End Sub
'=============================================================================
'OPCIONES
'=============================================================================

    




Public Function Folioingresomanual() As String
    
    Dim op As Integer
    Dim sql As New sqlventas.sqlventa
    Dim campos(3, 3) As String
    
    campos(0, 0) = "IFNULL(MAX(numero) + 1,'0000000001')"
    campos(1, 0) = ""
    campos(0, 2) = "sv_cuotas_detalle"
    condicion = "tipo='CA'"
    op = 5
    sql.response = campos
    Set sql.conexion = ventas
    Call sql.sqlventas(op, condicion)
    If sql.Status = 0 Then
        If sql.response(0, 3) <> "" And sql.response(0, 3) <> "0" Then
            Folioingresomanual = Format(sql.response(0, 3), "0000000000")
        Else
            Folioingresomanual = "0000000001"
        End If
    End If
End Function

Public Function existecargo(rut, FECHA, MONTO) As Boolean

        Dim csql As rdoQuery
        Dim resultado As rdoResultset
        
        Set csql = New rdoQuery
        Set csql.ActiveConnection = ventas

        csql.sql = "SELECT * "
        csql.sql = csql.sql & "FROM sv_cuotas_detalle "
        csql.sql = csql.sql & "WHERE rut = '" & rut & "' and vencimientooriginal='" & FECHA & "' and montocuota='" + MONTO + "' and tipo='CA' "
        csql.Execute
        
        If csql.RowsAffected > 0 Then
        existecargo = True
        Else
        existecargo = False
        End If
        
        Set resultado = Nothing
        csql.Close
        Set csql = Nothing
    End Function
Private Sub frmAdicionales_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
        Call cambiaColor(frmAdicionales)
        frmAdicionales.CaptionEstilo3D = Raised
    End Sub

    Private Sub frmAdicionales_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
        Call cambiaColor(frmAdicionales)
        frmAdicionales.CaptionEstilo3D = Inserted
        If cargo = True Then
            Load MCDatosAdicionales
            Call strcttoctrl_personales
            MCDatosAdicionales.Show vbModal
            Call leerClienteadicional(c2, "=")
        End If
    End Sub

    Private Sub frmCuentas_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
        Call cambiaColor(frmCuentas)
        frmCuentas.CaptionEstilo3D = Raised
    End Sub

    Private Sub frmCuentas_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
        Call cambiaColor(frmCuentas)
        frmCuentas.CaptionEstilo3D = Inserted
        If cargo = True Then
            Load MCCuentasAdicionales
            MCCuentasAdicionales.lblCupo.Caption = c2.cc.cupodirecto
            MCCuentasAdicionales.Show vbModal
            
        End If
    End Sub

    Private Sub frmFinancieros_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
        Call cambiaColor(frmFinancieros)
        frmFinancieros.CaptionEstilo3D = Raised
    End Sub

    Private Sub frmFinancieros_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
        Call cambiaColor(frmFinancieros)
        frmFinancieros.CaptionEstilo3D = Inserted
        If cargo = True Then
            Load MCDatosFinancieros
            Call strcttoctrl_financieros
            MCDatosFinancieros.Show vbModal
            Call leerClienteadicional(c2, "=")
        End If
    End Sub
 Private Sub frmLaborales_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
        Call cambiaColor(frmLaborales)
        frmLaborales.CaptionEstilo3D = Raised
    End Sub

    Private Sub frmLaborales_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
        Call cambiaColor(frmLaborales)
        frmLaborales.CaptionEstilo3D = Inserted
        If cargo = True Then
            Load MCDatosLaborales
            Call strcttoctrl_laborales
            MCDatosLaborales.Show vbModal
            Call leerClienteadicional(c2, "=")
        End If
    End Sub
Private Sub cmd15_Click()
repa.Visible = True
CONDO1.SetFocus

End Sub
Private Sub modificadueñacasa(cupo)

        Dim csql As rdoQuery
        Dim resultado As rdoResultset
        
        Set csql = New rdoQuery
        Set csql.ActiveConnection = ventas

        csql.sql = "update sv_maestroclientes set cupodirecto='" + cupo + "' "
        csql.sql = csql.sql & "WHERE tipocliente='06' and cupodirecto>'30000' "
        csql.Execute
            Call sincronizadatos(csql.sql, ventas)
        
    
    End Sub
Private Sub modificaconstructora(dato, rut)

        Dim csql As rdoQuery
        Dim resultado As rdoResultset
        
        Set csql = New rdoQuery
        Set csql.ActiveConnection = ventas

        csql.sql = "update sv_maestroclientes set constructora='" & dato & "' "
        csql.sql = csql.sql & "WHERE rut='" + rut + "' "
        csql.Execute
            Call sincronizadatos(csql.sql, ventas)
        
    
    End Sub

Sub cargatodo()

        Dim csql As rdoQuery
        Dim resultado As rdoResultset
        Dim mes As String
        Dim año As String
        Dim FECHA As String
        
        Set csql = New rdoQuery
        Set csql.ActiveConnection = ventas

        csql.sql = "SELECT mc.rut,mc.nombre,mc.cupodirecto,sum(cd.montocuota-cd.abono) as saldo,mc.diapago "
        csql.sql = csql.sql + "FROM sv_maestroclientes as mc inner join sv_cuotas_detalle as cd on (cd.rut=mc.rut) where mc.diapago<>'00' "
        csql.sql = csql.sql + "group by cd.rut order by mc.nombre "
        csql.Execute
        If csql.RowsAffected > 0 Then
        barra.Max = csql.RowsAffected + 1
        
        barra.Value = 0
        
        Set resultado = csql.OpenResultset
        
        mes = Format(COMBOMES.ListIndex + 1, "00")
        año = COMBOAÑO.text
        
    
     
        
        
        While Not resultado.EOF
        barra.Value = barra.Value + 1
        barra.Refresh
        
              If CDbl(resultado(3)) > 2200 Then
    If mes = "02" And CDbl(resultado(4)) >= 28 Then
       FECHA = año + "-" + mes + "-" + "28"
    End If

    If mes = "02" And CDbl(resultado(4)) >= 29 And año = "2008" Or año = "2012" Or año = "2016" Then
       FECHA = año + "-" + mes + "-" + "29"
    End If

    If mes = "02" And CDbl(resultado(4)) < 28 Then
        FECHA = año + "-" + mes + "-" + resultado(4)
    End If

    If mes <> "02" Then
        FECHA = año + "-" + mes + "-" + resultado(4)
    End If
              
       
     If existecargo(resultado(0), FECHA, CA2.text) = False Then
      
       Call grabarcuotas(resultado(0), FECHA)
     Else
     MsgBox ("cargo para este mes ya generado")
     End If
                  
           
              End If
              
            
            resultado.MoveNext
            Wend
        End If
        
    End Sub

