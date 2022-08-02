VERSION 5.00
Object = "{4777436C-EB8C-4596-98A8-EBCF98683969}#1.0#0"; "FlexCell.ocx"
Object = "{ADBBAED6-B16F-40DC-80DF-B44910CBA76C}#1.0#0"; "Frame-Xp.ocx"
Begin VB.Form creditoTMP 
   BackColor       =   &H00AE1118&
   BorderStyle     =   0  'None
   Caption         =   "PANTALLA GENERA CUOTAS"
   ClientHeight    =   9705
   ClientLeft      =   210
   ClientTop       =   1710
   ClientWidth     =   15360
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   9705
   ScaleWidth      =   15360
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin XPFrame.FrameXp FrameXp1 
      Height          =   2940
      Left            =   10920
      TabIndex        =   56
      Top             =   6240
      Width           =   4155
      _ExtentX        =   7329
      _ExtentY        =   5186
      BackColor       =   49344
      Caption         =   "VENCIMIENTO PROXIMAS CUOTAS"
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
      Begin FlexCell.Grid GRID3 
         Height          =   2580
         Left            =   45
         TabIndex        =   57
         Top             =   270
         Width           =   4065
         _ExtentX        =   7170
         _ExtentY        =   4551
         Cols            =   5
         DefaultFontSize =   8.25
         Rows            =   30
      End
   End
   Begin VB.CommandButton VERMOROSOS 
      BackColor       =   &H000000FF&
      Caption         =   "MOROSIDAD"
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
      Left            =   7965
      Style           =   1  'Graphical
      TabIndex        =   51
      Top             =   3105
      Width           =   2040
   End
   Begin VB.CommandButton retorno 
      BackColor       =   &H0000FF00&
      Caption         =   "RETORNO"
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
      Left            =   8280
      Style           =   1  'Graphical
      TabIndex        =   50
      Top             =   5400
      Width           =   1410
   End
   Begin VB.CommandButton ELIMINAR 
      BackColor       =   &H000000FF&
      Caption         =   "Eliminar Credito"
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
      Left            =   3555
      Style           =   1  'Graphical
      TabIndex        =   49
      Top             =   5400
      Visible         =   0   'False
      Width           =   3120
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "X"
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
      Left            =   14715
      Style           =   1  'Graphical
      TabIndex        =   41
      Top             =   135
      UseMaskColor    =   -1  'True
      Width           =   465
   End
   Begin VB.CommandButton imprimir 
      BackColor       =   &H0000FF00&
      Caption         =   "IMPRIMIR"
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
      Left            =   6840
      Style           =   1  'Graphical
      TabIndex        =   40
      Top             =   5400
      Visible         =   0   'False
      Width           =   1290
   End
   Begin XPFrame.FrameXp datosboleta 
      Height          =   960
      Left            =   45
      TabIndex        =   28
      Top             =   90
      Width           =   15270
      _ExtentX        =   26935
      _ExtentY        =   1693
      BackColor       =   16744576
      Caption         =   "DETALLES"
      CaptionEstilo3D =   1
      BackColor       =   16744576
      ForeColor       =   1563884
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.TextBox cajadoc 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   2520
         MaxLength       =   2
         TabIndex        =   45
         Text            =   "01"
         Top             =   495
         Width           =   465
      End
      Begin VB.TextBox NUMERO 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   5175
         MaxLength       =   10
         TabIndex        =   1
         Top             =   495
         Width           =   1635
      End
      Begin VB.TextBox TIPO 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   720
         MaxLength       =   2
         TabIndex        =   0
         Text            =   "VC"
         Top             =   495
         Width           =   465
      End
      Begin VB.Label lblLabels 
         BackStyle       =   0  'Transparent
         Caption         =   "CAJA"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   10
         Left            =   1560
         TabIndex        =   46
         Top             =   495
         Width           =   885
      End
      Begin VB.Label lblLabels 
         BackStyle       =   0  'Transparent
         Caption         =   "MONTO CREDITO"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   8
         Left            =   7320
         TabIndex        =   33
         Top             =   495
         Width           =   2940
      End
      Begin VB.Label MONTODOCUMENTO 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
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
         ForeColor       =   &H80000008&
         Height          =   420
         Left            =   10320
         TabIndex        =   32
         Top             =   495
         Width           =   2280
      End
      Begin VB.Label lblLabels 
         BackStyle       =   0  'Transparent
         Caption         =   "NUMERO"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   7
         Left            =   3615
         TabIndex        =   30
         Top             =   495
         Width           =   1635
      End
      Begin VB.Label lblLabels 
         BackStyle       =   0  'Transparent
         Caption         =   "TP"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   6
         Left            =   90
         TabIndex        =   29
         Top             =   495
         Width           =   525
      End
   End
   Begin VB.Timer tmrBlink 
      Interval        =   500
      Left            =   0
      Top             =   0
   End
   Begin XPFrame.FrameXp DATOSCREDITO 
      Height          =   8445
      Left            =   90
      TabIndex        =   5
      Top             =   1170
      Width           =   15225
      _ExtentX        =   26855
      _ExtentY        =   14896
      BackColor       =   16744576
      Caption         =   "Venta a Crédito"
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
      Begin VB.TextBox PIVOTE 
         Height          =   330
         Left            =   600
         MaxLength       =   2
         TabIndex        =   39
         Top             =   120
         Visible         =   0   'False
         Width           =   690
      End
      Begin VB.TextBox SUCU 
         Height          =   330
         Left            =   225
         TabIndex        =   34
         Text            =   "Text1"
         Top             =   90
         Visible         =   0   'False
         Width           =   420
      End
      Begin VB.TextBox rut2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   2070
         MaxLength       =   9
         TabIndex        =   2
         Top             =   495
         Width           =   1530
      End
      Begin XPFrame.FrameXp FRMCUOTAS 
         Height          =   1905
         Left            =   90
         TabIndex        =   19
         Top             =   2835
         Width           =   10635
         _ExtentX        =   18759
         _ExtentY        =   3360
         BackColor       =   8454016
         Caption         =   "Calcula Cuotas"
         CaptionEstilo3D =   1
         BackColor       =   8454016
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Begin VB.TextBox CREDITO 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   4230
            Locked          =   -1  'True
            TabIndex        =   54
            Top             =   855
            Width           =   1680
         End
         Begin VB.TextBox PIE 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   2340
            TabIndex        =   52
            Top             =   855
            Width           =   1680
         End
         Begin VB.TextBox VALORCUOTA 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   7200
            TabIndex        =   38
            Top             =   855
            Width           =   1545
         End
         Begin VB.TextBox AÑOC 
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
            Height          =   375
            Left            =   9720
            MaxLength       =   4
            TabIndex        =   37
            Top             =   855
            Width           =   780
         End
         Begin VB.TextBox MESC 
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
            Height          =   375
            Left            =   9270
            MaxLength       =   2
            TabIndex        =   36
            Top             =   855
            Width           =   420
         End
         Begin VB.TextBox DIAC 
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
            Height          =   375
            Left            =   8820
            MaxLength       =   2
            TabIndex        =   35
            Top             =   855
            Width           =   420
         End
         Begin VB.TextBox MONTO 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   180
            TabIndex        =   31
            Top             =   855
            Width           =   1860
         End
         Begin VB.CommandButton ACEPTAR 
            BackColor       =   &H0000FF00&
            Caption         =   "ACEPTAR Y GRABAR"
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
            Left            =   225
            Style           =   1  'Graphical
            TabIndex        =   24
            Top             =   1395
            Width           =   2775
         End
         Begin VB.TextBox CUOTAS 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   6075
            TabIndex        =   4
            Top             =   855
            Width           =   1005
         End
         Begin VB.CheckBox Check1 
            BackColor       =   &H0080FF80&
            Caption         =   "PRECIO CONTADO"
            Height          =   330
            Left            =   180
            TabIndex        =   42
            Top             =   1305
            Visible         =   0   'False
            Width           =   2715
         End
         Begin VB.Label Label6 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FF8080&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "MONTO CREDITO"
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
            Left            =   4275
            TabIndex        =   55
            Top             =   495
            Width           =   1680
         End
         Begin VB.Label Label5 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FF8080&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "PIE"
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
            Left            =   2340
            TabIndex        =   53
            Top             =   495
            Width           =   1680
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FF8080&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "CUOTAS"
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
            Left            =   6120
            TabIndex        =   20
            Top             =   495
            Width           =   1005
         End
         Begin VB.Label Label4 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FF8080&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "MONTO VENTA"
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
            Left            =   180
            TabIndex        =   23
            Top             =   495
            Width           =   1860
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FF8080&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "PRIMER VENC."
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
            Left            =   8820
            TabIndex        =   22
            Top             =   495
            Width           =   1725
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FF8080&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "MONTO CUOTA"
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
            Left            =   7200
            TabIndex        =   21
            Top             =   495
            Width           =   1545
         End
      End
      Begin XPFrame.FrameXp FrameXp2 
         Height          =   3480
         Left            =   90
         TabIndex        =   17
         Top             =   4770
         Width           =   15045
         _ExtentX        =   26538
         _ExtentY        =   6138
         BackColor       =   16761024
         Caption         =   "CUOTAS"
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
            TabIndex        =   18
            Top             =   225
            Width           =   10500
            _ExtentX        =   18521
            _ExtentY        =   5106
            Cols            =   5
            DefaultFontSize =   8.25
            Rows            =   30
         End
      End
      Begin XPFrame.FrameXp FrameXp4 
         Height          =   4335
         Left            =   10845
         TabIndex        =   25
         Top             =   405
         Width           =   4335
         _ExtentX        =   7646
         _ExtentY        =   7646
         BackColor       =   16777152
         Caption         =   "SIMULADOR DE CUOTAS"
         CaptionEstilo3D =   1
         BackColor       =   16777152
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
         Begin FlexCell.Grid Grid2 
            Height          =   4065
            Left            =   90
            TabIndex        =   26
            Top             =   270
            Width           =   4200
            _ExtentX        =   7408
            _ExtentY        =   7170
            BackColor1      =   12648384
            Cols            =   5
            DefaultFontSize =   8.25
            Rows            =   30
         End
      End
      Begin VB.TextBox DIAPAGO 
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
         Left            =   2025
         MaxLength       =   2
         TabIndex        =   3
         Top             =   1485
         Width           =   465
      End
      Begin VB.Label lbltc 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
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
         ForeColor       =   &H80000008&
         Height          =   420
         Left            =   3420
         TabIndex        =   48
         Top             =   1485
         Width           =   465
      End
      Begin VB.Label lbltipocliente 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
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
         ForeColor       =   &H80000008&
         Height          =   420
         Left            =   3960
         TabIndex        =   47
         Top             =   1485
         Width           =   5460
      End
      Begin VB.Label LBLMOROSO 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "$ 0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   420
         Left            =   7920
         TabIndex        =   44
         Top             =   2295
         Width           =   1875
      End
      Begin VB.Label lblLabels 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H000000C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "MOROSO"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   390
         Index           =   9
         Left            =   7875
         TabIndex        =   43
         Top             =   1935
         Width           =   2010
      End
      Begin VB.Label lbldv 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
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
         ForeColor       =   &H80000008&
         Height          =   420
         Left            =   3600
         TabIndex        =   27
         Top             =   495
         Width           =   285
      End
      Begin VB.Label lblLabels 
         BackStyle       =   0  'Transparent
         Caption         =   "Rut Cliente"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   0
         Left            =   120
         TabIndex        =   16
         Top             =   480
         Width           =   1680
      End
      Begin VB.Label lblNombre 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
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
         ForeColor       =   &H80000008&
         Height          =   420
         Left            =   3915
         TabIndex        =   15
         Top             =   495
         Width           =   5970
      End
      Begin VB.Label lblLabels 
         BackStyle       =   0  'Transparent
         Caption         =   "Dirección"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   1
         Left            =   120
         TabIndex        =   14
         Top             =   960
         Width           =   1680
      End
      Begin VB.Label lblDireccion 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
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
         ForeColor       =   &H80000008&
         Height          =   420
         Left            =   2025
         TabIndex        =   13
         Top             =   990
         Width           =   7845
      End
      Begin VB.Label lblLabels 
         BackStyle       =   0  'Transparent
         Caption         =   "Día de Pago"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   2
         Left            =   120
         TabIndex        =   12
         Top             =   1440
         Width           =   1800
      End
      Begin VB.Label lblLabels 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Crédito Autorizado"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   390
         Index           =   3
         Left            =   75
         TabIndex        =   11
         Top             =   1920
         Width           =   2730
      End
      Begin VB.Label lblCupo 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "$ 0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   420
         Left            =   120
         TabIndex        =   10
         Top             =   2280
         Width           =   2640
      End
      Begin VB.Label lblLabels 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Crédito Utilizado"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   390
         Index           =   4
         Left            =   2880
         TabIndex        =   9
         Top             =   1935
         Width           =   2625
      End
      Begin VB.Label lblUtilizado 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "$ 0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   420
         Left            =   2925
         TabIndex        =   8
         Top             =   2295
         Width           =   2535
      End
      Begin VB.Label lblLabels 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Disponible"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   390
         Index           =   5
         Left            =   5580
         TabIndex        =   7
         Top             =   1935
         Width           =   2235
      End
      Begin VB.Label lblDisponible 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "$ 0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   420
         Left            =   5625
         TabIndex        =   6
         Top             =   2295
         Width           =   2145
      End
   End
End
Attribute VB_Name = "creditoTMP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim fechas(100) As String
Dim fechacompra As String
Dim moroso As Double
Dim rutprincipal As String
Dim rutadicional As String
Dim adicional As String
Dim clientebloqueado As Boolean
Dim mensajecredito As String

Private Sub ACEPTAR_Click()
If CREDITO.text <> "" And Cuotas.text <> "" And PIE.text <> "" And CREDITO.text <> "0" And Cuotas.text <> "0" And VALORCUOTA.text <> "0" And VALORCUOTA.text <> "" Then

    If MsgBox("ESTA SEGURO DE REALIZAR EL CARGO DEL CREDITO ", vbYesNo) = vbYes Then
            CREDITO.text = Format(CDbl(MONTO.text) - CDbl(PIE.text), "$ #,###,##0")
            CALCULAR_Click
            grabarcuotas
            Call NUMERO_KeyPress(13)
            Call imprimir_Click

      
    Else
            Cuotas.SetFocus

    End If
Else
    CREDITO.SetFocus
End If
End Sub

Private Sub AÑOC_LostFocus()
Call esfecha(DIAC, MESC, AÑOC, "yyyy")

End Sub



Private Sub cajadoc_KeyPress(KeyAscii As Integer)
  KeyAscii = esNumero(KeyAscii)
If KeyAscii = 13 And cajadoc.text <> "" Then
   NUMERO.SetFocus
End If
End Sub

Private Sub CALCULAR_Click()
If Cuotas.text <> "" Then
   If CDbl(lblCupo.Caption) > CDbl(lblUtilizado.Caption + CDbl(CREDITO.text)) Then
      CALCULACUOTA
   
   Else
     If MsgBox("CUPO NO ES SUFICIENTE", vbOKOnly, "ATENCION") = vbOK Then
       
      MONTO.SetFocus
     End If
   End If
Else

End If
End Sub

Private Sub Check1_Click()
CALCULACUOTA

End Sub

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub CUOTAS_GotFocus()

Call cargatexto(Cuotas)
End Sub

Private Sub CUOTAS_KeyPress(KeyAscii As Integer)
KeyAscii = esNumero(KeyAscii)
If Cuotas.text <> "" Then
If KeyAscii = 13 And Cuotas.text <> "" And Cuotas.text <> "0" And CDbl(Cuotas.text) <= leefactorMAXIMO Then
CALCULAR_Click
ACEPTAR.SetFocus

End If


End If

End Sub

Private Sub DIAC_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
MESC.SetFocus

End If


End Sub

Private Sub DIAC_LostFocus()
Call esfecha(DIAC, MESC, AÑOC, "dd")

End Sub

Private Sub DIAPAGO_GotFocus()
Call selecciona(DIAPAGO)
End Sub

Private Sub DIAPAGO_KeyPress(KeyAscii As Integer)
KeyAscii = esNumero(KeyAscii)
If KeyAscii = 13 Then

MONTO.SetFocus

End If

End Sub

Private Sub ELIMINAR_Click()
        
        Dim campos(18, 3) As String
        Dim op As Integer
        Dim K As Integer
         frmglosaeliminacion.Show vbModal
        Set sql = New sqlventas.sqlventa
        campos(0, 2) = "sv_cuotas_detalle"
        condicion = "tipo='" + TIPO.text + "' and numero='" + NUMERO.text + "' and rut='" + rut2.text + lblDV.Caption + "' "
        op = 4
        sql.response = campos
        Set sql.conexion = ventas
        sql.audit = True: sql.programaactivo = Me.Caption
        Set sql.conauditoria = conauditoria: sql.usuarioauditoria = usuarioSistema
        sql.glosaeliminacion = glosaeliminacionsistema
        sql.solicitoeliminacion = solicitaeliminacion
        Call sql.sqlventas(op, condicion)
       Call retorno_Click
       
        

End Sub

Private Sub Form_Load()
CARGAGRILLA
CARGAGRILLA2
CARGAGRILLA3

DATOSCREDITO.Enabled = False
FRMCUOTAS.Enabled = False
MONTO.Locked = True
datosboleta.Caption = "CAJERO :" + nombrecajero
ELIMINAR.Visible = False



End Sub

Sub CARGAGRILLA()
    Grid1.Cols = 11
    
    Grid1.Column(0).Width = 0
    Grid1.Column(1).Width = 0
    Grid1.Column(2).Width = 80
    Grid1.Column(3).Width = 80
    Grid1.Column(4).Width = 80
    Grid1.Column(5).Width = 80
    Grid1.Column(6).Width = 80
    Grid1.Column(7).Width = 80
    Grid1.Column(8).Width = 80
    Grid1.Column(9).Width = 80
    Grid1.Column(10).Width = 80
    
    
    Grid1.Column(0).Locked = True
    Grid1.Column(1).Locked = True
    Grid1.Column(2).Locked = True
    Grid1.Column(3).Locked = True
    Grid1.Column(4).Locked = True
    Grid1.Column(5).Locked = True
    Grid1.Column(6).Locked = True
    Grid1.Column(7).Locked = True
    Grid1.Column(8).Locked = True
    Grid1.Column(9).Locked = True
    Grid1.Column(10).Locked = True
    
    Grid1.Cell(0, 1).text = "TD"
    Grid1.Cell(0, 2).text = "NUMERO"
    Grid1.Cell(0, 3).text = "FECHA"
    Grid1.Cell(0, 4).text = "CUOTA"
    Grid1.Cell(0, 5).text = "ABONO"
    Grid1.Cell(0, 6).text = "SALDO"
    Grid1.Cell(0, 7).text = "MOROSIDAD"
    Grid1.Cell(0, 8).text = "INTERES"
    Grid1.Cell(0, 9).text = "TOTAL"
    
    Grid1.Column(4).Alignment = cellRightTop
    Grid1.Column(5).Alignment = cellRightTop
    Grid1.Column(6).Alignment = cellRightTop
    Grid1.Column(7).Alignment = cellRightTop
    Grid1.Column(8).Alignment = cellRightTop
    Grid1.Column(9).Alignment = cellRightTop
    
    
    
    
    
    Me.Left = (Screen.Width - Me.Width) / 2
    Me.Top = (Screen.Height - Me.Height) / 2 - 950
Grid1.Rows = 1
    
   
End Sub
Sub CARGAGRILLA2()
    Grid2.Cols = 5
    Grid2.Column(0).Width = 0
    Grid2.Column(1).Width = 70
    Grid2.Column(2).Width = 70
    Grid2.Column(3).Width = 70
    Grid2.Column(4).Width = 0
    
    Grid2.Column(0).Locked = True
    Grid2.Column(1).Locked = True
    Grid2.Column(2).Locked = True
    Grid2.Column(3).Locked = True
    Grid2.Column(4).Locked = True
    
    Grid2.Cell(0, 1).text = "CUOTAS"
    Grid2.Cell(0, 2).text = "VALOR"
    Grid2.Cell(0, 3).text = "TOTAL"
    Grid2.Cell(0, 4).text = "VENCIMIENTO"
    
    Grid2.Column(1).Alignment = cellRightTop
    Grid2.Column(2).Alignment = cellRightTop
    Grid2.Column(3).Alignment = cellRightTop
    Grid2.Column(4).Alignment = cellRightTop
    
    Grid2.Rows = leefactorMAXIMO + 1
    Grid2.DefaultFont.Size = 10
    Grid2.DefaultFont.Bold = True
    
    
    
    

End Sub

Sub CALCULACUOTA()
Dim FACTOR As Double
Dim cuota As Double
Dim interes As Double

If Cuotas.text <> "" Then

FACTOR = leefactor(1) / 100
If lbltc.Caption = "03" And CDbl(Cuotas.text) < 5 Then FACTOR = 0
If lbltc.Caption = "04" And CDbl(Cuotas.text) = 1 Then FACTOR = 0
If lbltc.Caption = "07" And CDbl(Cuotas.text) < 4 Then FACTOR = 0


cuota = Int((CDbl(Replace(CREDITO.text, ",", "")) / CDbl(Cuotas.text)) + 0.5)

interes = 0
For K = 1 To CDbl(Cuotas.text)
interes = interes + (cuota * (K * FACTOR))
Next K

End If
cuota = Round((CDbl(CREDITO.text) + interes) / CDbl(Cuotas), 0)



'cuota = Int((CDbl(Replace(credito.text, ",", "")) * FACTOR / CDbl(CUOTAS.text)) + 0.5)

VALORCUOTA.text = Format(cuota, "###,###,###")
'     If (CDbl(VALORCUOTA.text) * CDbl(CUOTAS.text)) > CDbl(lblDisponible.Caption) Then
'       If MsgBox("CUPO NO ES SUFICIENTE PARA CUBRIR LOS INTERESES", vbOKOnly, "ATENCION") = vbOK Then
'       VALORCUOTA.text = ""
'       CUOTAS.SetFocus
'       End If
'     Else
     Rem VALORCUOTA.SetFocus
        CALCULATODASCUOTAS
'End If


End Sub
Sub CALCULATODASCUOTAS()
Dim K As Integer
Dim cuota As Double
Dim fin As Double
Dim FECHA As String
Dim mesinicio As Double
Dim añoinicio As Double
Dim s As Integer
CARGAGRILLA2

FECHA = DIAC.text + "/" + MESC.text + "/" + AÑOC.text
fin = leefactorMAXIMO
mesinicio = MESC.text
añoinicio = AÑOC.text
For K = 1 To fin
FACTOR = leefactor(1) / 100

If lbltc.Caption = "03" And K < 5 Then FACTOR = 0
If lbltc.Caption = "04" And K = 1 Then FACTOR = 0
If lbltc.Caption = "07" And K < 4 Then FACTOR = 0



cuota = Int((CDbl(Replace(CREDITO.text, ",", "")) / K) + 0.5)
interes = 0
For s = 1 To K
interes = interes + (cuota * (s * FACTOR))
Next s


cuota = Round((CDbl(CREDITO.text) + interes) / K, 0)

Grid2.Cell(K, 1).text = K

Grid2.Cell(K, 2).text = Format(cuota, "###,###,##0")
Grid2.Cell(K, 3).text = Format(cuota * K, "###,###,##0")
Grid2.Cell(K, 4).text = Format(FECHA, "dd-mm-yyyy")
mesinicio = mesinicio + 1
If mesinicio > 12 Then
mesinicio = 1
añoinicio = añoinicio + 1

End If

pivote.MaxLength = 2
pivote.text = mesinicio
pivote.text = ceros(pivote)


If pivote = "02" And CDbl(DIAC.text) >= 28 Then
FECHA = "28" & "-" & pivote.text & "-" & añoinicio
End If

If pivote = "02" And CDbl(DIAC.text) >= 29 And añoinicio = "2008" Or añoinicio = "2012" Or añoinicio = "2016" Then
FECHA = "29" & "-" & pivote.text & "-" & añoinicio
End If

If pivote = "02" And CDbl(DIAC.text) < 28 Then
FECHA = DIAC.text & "-" & pivote.text & "-" & añoinicio
End If

If pivote <> "02" Then
FECHA = DIAC.text & "-" & pivote.text & "-" & añoinicio
End If

Next K

End Sub
Sub CALCULAPRIMERVENCIMIENTO()
Dim mes As Double
Dim año As Double
Dim dia As Double

Dim sumar As Double
dia = Format(fechasistema, "dd")
mes = Format(fechasistema, "mm")
año = Format(fechasistema, "yyyy")
If dia <= CDbl(DIAPAGO.text) Then sumar = 1 Else sumar = 2
DIAC.text = DIAPAGO.text
mes = mes + sumar
If mes > 12 Then
mes = mes - 12
año = año + 1
End If
pivote.text = mes
pivote.MaxLength = 2
pivote.text = ceros(pivote)
MESC = pivote.text
AÑOC = año

If mes = 2 And DIAC.text > "28" Then
dia = 28
DIAC.text = "28"

End If

End Sub

Private Sub imprimir_Click()
MsgBox ("VERIFIQUE IMPRESORA E IMPRIMA EL VALE ")
 IMPRIMEcredito

End Sub
Private Sub Ingreso_Click()
        Load creditoTMPmanual
        
        creditoTMPmanual.Show
End Sub

Private Sub MESC_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
AÑOC.SetFocus

End If

End Sub

Private Sub MESC_LostFocus()
Call esfecha(DIAC, MESC, AÑOC, "mm")

End Sub

Private Sub MONTO_GotFocus()
DIAPAGO.text = ceros(DIAPAGO)
If DIAPAGO.text = "00" Or DIAPAGO > "31" Then
DIAPAGO.SetFocus
Else
CALCULAPRIMERVENCIMIENTO


End If


End Sub

Private Sub MONTO_KeyPress(KeyAscii As Integer)
KeyAscii = esNumero(KeyAscii)

If KeyAscii = 13 And MONTO.text <> "" Then
    If CDbl(MONTO.text) <> 0 Then
    MONTO.text = Format(MONTO.text, "$ ###,###,###")
        
        PIE.SetFocus
    End If
End If
End Sub


Private Sub NUMERO_GotFocus()
Call selecciona(NUMERO)

End Sub

Private Sub NUMERO_KeyPress(KeyAscii As Integer)
    KeyAscii = esNumero(KeyAscii)
    If KeyAscii = 13 And NUMERO.text <> "" And NUMERO.text <> "000000000" Then
        NUMERO.text = ceros(NUMERO)
        'ZURITA
        'LE AGREGUE QUE HAGA LA BUSQUEDA POR VENTA BOLETA Y FACTURA
        'If TIPO.text = "VC" Or TIPO.text = "VR"  Then
      If TIPO.text = "VC" Or TIPO.text = "VR" Or TIPO = "BV" Or TIPO = "FV" Then

'        If LEERDOCUMENTO(TIPO.text, NUMERO.text, cajadoc.text) = True Then
'        DATOSCREDITO.Enabled = True
'        rut2.SetFocus
'        MONTO.Locked = True
'
'
'        Else
'        MsgBox ("NUMERO DE DOCUMENTO NO EXISTE ")
'        NUMERO.SetFocus
'
'
'
'        End If
'        Else
        DATOSCREDITO.Enabled = True
        rut2.SetFocus
        MONTO.Locked = False
       End If

        
    End If

End Sub
Public Function LEERDOCUMENTO(TIPO, NUMERO, CAJA, FECHA) As Boolean

        Dim csql As rdoQuery
        Dim resultado As rdoResultset
        Dim i As Integer
        Set csql = New rdoQuery
        Set csql.ActiveConnection = ventasRubro
        
        csql.sql = "SELECT dc.rut,dc.total,mc.diapago,mc.cupodirecto,mc.cupoutilizadotmp,mc.nombre,mc.direccion,dc.abono2,dc.fecha "
        csql.sql = csql.sql & "FROM sv_documento_cabeza_" & empresaActiva & " as dc," + baseVentas + ".sv_maestroclientes as mc "
        csql.sql = csql.sql & "WHERE tipo = '" & TIPO & "' and foliosii='" & NUMERO & "' and caja='" & CAJA & "' and dc.rut=mc.rut and local='" + empresaActiva + "' "
        csql.Execute
        
        If csql.RowsAffected > 0 Then
       
            Set resultado = csql.OpenResultset
            While Not resultado.EOF
        MONTODOCUMENTO.Caption = Format(resultado(1), "###,###,##0")
     
        fechacompra = Format(resultado(8), "yyyy-mm-dd")
        
        DIAPAGO.text = resultado(2)
        lblCupo.Caption = Format(resultado(3) * (1 + toleranciacredito / 100), "###,###,##0")
        lblUtilizado.Caption = Format(resultado(4), "###,###,##0")
        lblDisponible.Caption = Format(resultado(3) - resultado(4), "###,###,##0")
        CREDITO.text = resultado(1) - resultado(7)
        rut2.text = Mid(resultado(0), 1, 9)
        lblDV.Caption = Mid(resultado(0), 10, 1)
        lblNombre.Caption = resultado(5)
        lblDireccion.Caption = resultado(6)
        
        LEERDOCUMENTO = True
        
            resultado.MoveNext
            Wend
        Else
        LEERDOCUMENTO = False
        
        End If
        Set resultado = Nothing
        csql.Close
        Set csql = Nothing
    End Function
                                                            'ZURITA
'Public Function LEErcuotas(TIPO, NUMERO) As Boolean         <- CADENA DE CONSULTA ORIGINA
Public Function LEErcuotas(CAJA, TIPO, NUMERO, FECHA) As Boolean  '<- SE LE AGREGO LA CAJA

        Dim csql As rdoQuery
        Dim resultado As rdoResultset
        Dim i As Integer
        Dim diasmora As Double
        Dim saldo As Double
        Dim interes As Double
        
        Set csql = New rdoQuery
        Set csql.ActiveConnection = ventas

        csql.sql = "SELECT dc.rut,dc.montocredito,mc.diapago,mc.cupodirecto,mc.cupoutilizadotmp,mc.nombre,mc.direccion,dc.cantidadcuotas,dc.montocuota,dc.vencimientooriginal,dc.tipo,dc.numero,dc.abono,dc.numerocuota,dc.vencimientoactual,dc.pie,dc.montoventa "
        csql.sql = csql.sql & "FROM sv_cuotas_detalle as dc," + baseVentas + ".sv_maestroclientes as mc "
        'ZURITA - CADENA DE CONSULTA ORIGINAL
        'csql.sql = csql.sql & "WHERE tipo = '" & TIPO & "' and numero='" & NUMERO & "' and dc.rut=mc.rut  "
        csql.sql = csql.sql & "WHERE tipo = '" & TIPO & "' and numero='" & NUMERO & "' and caja = '" & CAJA & "' and fechacompra = '" & FECHA & "' and dc.rut=mc.rut  "
        csql.sql = csql.sql & "order by vencimientooriginal "
        csql.Execute
        
        If csql.RowsAffected > 0 Then

            Set resultado = csql.OpenResultset
            
        MONTODOCUMENTO.Caption = Format(resultado(1), "###,###,##0")
        DIAPAGO.text = resultado(2)
        lblCupo.Caption = Format(resultado(3), "###,###,##0")
        lblUtilizado.Caption = Format(resultado(4), "###,###,##0")
        lblDisponible.Caption = Format(resultado(3) - resultado(4), "###,###,##0")
        CREDITO.text = resultado(1)
        MONTO.text = resultado(16)
        PIE.text = resultado(15)
        rut2.text = Mid(resultado(0), 1, 9)
        lblDV.Caption = Mid(resultado(0), 10, 1)
        lblNombre.Caption = resultado(5)
        lblDireccion.Caption = resultado(6)
        Cuotas.text = resultado(7)
        VALORCUOTA.text = resultado(8)
        DIAC.text = Format(resultado(9), "dd")
        MESC.text = Format(resultado(9), "mm")
        AÑOC.text = Format(resultado(9), "yyyy")
        
        LEErcuotas = True
        Grid1.Rows = 1
        While Not resultado.EOF
        Grid1.Rows = Grid1.Rows + 1
        Grid1.Cell(Grid1.Rows - 1, 2).text = resultado(13) & " / " & resultado(7)
        Grid1.Cell(Grid1.Rows - 1, 3).text = Format(resultado(14), "dd/mm/yyyy")
        Grid1.Cell(Grid1.Rows - 1, 4).text = Format(resultado(8), "###,###,###")
        Grid1.Cell(Grid1.Rows - 1, 5).text = Format(resultado(12), "###,###,###")
        saldo = resultado(8) - resultado(12)
        diasmora = 0
        interes = 0
        total = saldo + interes
        
        Grid1.Cell(Grid1.Rows - 1, 6).text = Format(saldo, "###,###,###")
        Grid1.Cell(Grid1.Rows - 1, 7).text = diasmora
        Grid1.Cell(Grid1.Rows - 1, 8).text = interes
        Grid1.Cell(Grid1.Rows - 1, 9).text = total
        
            
            
            resultado.MoveNext
            Wend
        Else
        LEErcuotas = False

        End If
        Set resultado = Nothing
        csql.Close
        Set csql = Nothing
    End Function

Public Function leerCliente(RUT) As Boolean

        Dim csql As rdoQuery
        Dim resultado As rdoResultset
        Dim i As Integer
        Dim cSql1 As New rdoQuery
        Dim resultados As rdoResultset
        
        
        
        
        Set cSql1.ActiveConnection = ventas
        cSql1.sql = "select * from sv_maestroclientes_adicionales "
        cSql1.sql = cSql1.sql & "where rutadicional='" & RUT & "' "
        cSql1.Execute
        If cSql1.RowsAffected > 0 Then
        Set resultados = cSql1.OpenResultset
        rutprincipal = resultados("rut")
        adicional = resultados("nombre")
        rutadicional = RUT
        Else
        rutadicional = ""
        adicional = ""
        rutprincipal = ""
        End If
        cSql1.Close
        Set cSql1 = Nothing
        Set resultados = Nothing
        
        
        
        
        Set csql = New rdoQuery
        Set csql.ActiveConnection = ventas
        csql.sql = "SELECT mc.diapago,mc.cupodirecto,mc.cupoutilizadodirecto,mc.nombre,mc.direccion,mc.tipocliente,mc.bloqueotmp "
        csql.sql = csql.sql & "FROM sv_maestroclientes as mc  "
        If rutprincipal = "" Then
        csql.sql = csql.sql & "WHERE mc.rut='" + RUT + "' "
        Else
        csql.sql = csql.sql & "WHERE mc.rut='" + rutprincipal + "' "
        End If
        
        csql.Execute
        If csql.RowsAffected > 0 Then
            Set resultado = csql.OpenResultset
            While Not resultado.EOF
        
        If escredito(resultado(5)) = True Then
        
                     DIAPAGO.text = resultado(0)
                    lblCupo.Caption = Format(resultado(1) * (1 + toleranciacredito / 100), "###,###,##0")
                    lblUtilizado.Caption = LEErcreditoutilizado(RUT, "9999-99-99")
                    lblDisponible.Caption = Format(CDbl(lblCupo.Caption) - CDbl(lblUtilizado.Caption), "###,###,##0")
                    lbltipocliente.Caption = leertipocli(resultado(5))
                    lbltc.Caption = resultado(5)
                    If resultado(6) = "1" Then
                        clientebloqueado = True
                        mensajecredito = "CLIENTE BLOQUEADO PASAR A OFICINA COMERCIAL"
                        Else
                        clientebloqueado = False
        
                    End If
      
        
                    If adicional = "" Then
                        lblNombre.Caption = resultado(3)
                    Else
                        lblNombre.Caption = adicional & " Adicional de: " & resultado(3)
                    End If
        
                    lblDireccion.Caption = resultado(4)
                    moroso = LEErcreditoutilizado(RUT, Format(fechasistema, "yyyy-mm-dd"))
                    LBLMOROSO.Caption = Format(moroso, "###,###,##0")
                    leerCliente = True
        End If
        
            resultado.MoveNext
            Wend
        Else
        leerCliente = False
        
        End If
        Set resultado = Nothing
        csql.Close
        Set csql = Nothing
    End Function


Private Sub PIE_KeyPress(KeyAscii As Integer)
KeyAscii = esNumero(KeyAscii)

If KeyAscii = 13 Then
        If PIE.text = "" Then
        PIE.text = "$0"
        End If
If CDbl(PIE.text) < CDbl(MONTO.text) Then
PIE.text = Format(PIE.text, "$ ###,###,##0")
CREDITO.text = Format(CDbl(MONTO.text) - CDbl(PIE.text), "$ #,###,##0")

Cuotas.SetFocus
Else
MsgBox ("MONTO DEL PIE DEBE SER INFERIOR AL MONTO DE LA COMPRA")
PIE.text = "0"
PIE.SetFocus

End If

End If

End Sub

Private Sub retorno_Click()
NUMERO.text = ""
MONTODOCUMENTO.Caption = ""
rut2.text = ""
lblDV.Caption = ""
lblNombre.Caption = ""
lblDireccion = ""
DIAPAGO.text = ""
lblCupo.Caption = "$ 0"
lblUtilizado.Caption = "$ 0"
lblDisponible.Caption = "$ 0"
MONTO.text = ""
Cuotas.text = ""
VALORCUOTA.text = ""
DIAC.text = ""
MESC.text = ""
AÑOC.text = ""
Grid1.Rows = 1
Grid2.Rows = 1
Check1.Value = 0
CARGAGRILLA2
ELIMINAR.Visible = False
PIE.text = ""
CREDITO.text = ""
CARGAGRILLA3

TIPO.SetFocus

End Sub

Private Sub rut2_GotFocus()

Call selecciona(rut2)
'ZURITA - LE INGRESÉ LA CAJA, NO SE PODIA REIMPRIMIR UN VALE HECHO EN UNA CAJA
'If LEErcuotas(TIPO.text, NUMERO.text) = True Then  '<- ORIGINAL
If LEErcuotas(cajadoc, TIPO.text, NUMERO.text, (Format(fechasistema, "yyyy-mm-dd"))) = True Then
FRMCUOTAS.Enabled = True
imprimir.Visible = True

ACEPTAR.Visible = False
imprimir.Visible = True
If Verifica_Permiso(Me.Caption, "autoriza") = True Then
ELIMINAR.Visible = True
  Call leerCliente(rut2.text + lblDV.Caption)

End If
DATOSCREDITO.Enabled = False


retorno.Visible = True
Else
DATOSCREDITO.Enabled = False

imprimir.Visible = False
retorno.Visible = False
FRMCUOTAS.Enabled = False

ACEPTAR.Visible = True

End If
retorno.Visible = True

End Sub

Private Sub rut2_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF2 Then
            
            DATOSCREDITO.Enabled = True
            
            
            Call ayudaCliente(rut2, SUCU, lblDV)
  End If
End Sub

Private Sub rut2_KeyPress(KeyAscii As Integer)
          KeyAscii = esNumero(KeyAscii)
           If KeyAscii = 13 And rut2.text <> "" And Val(rut2.text) <> 0 Then
             rut2.text = ceros(rut2)
             lblDV.Caption = RUT(rut2.text)
             
             If leerCliente(rut2.text + lblDV.Caption) = True Then
             Call LEErcuotasACUMULADAS(rut2.text + lblDV.Caption)
             
             
             If clientebloqueado = False Then
                
                If CDbl(lblCupo.Caption) > CDbl(lblUtilizado.Caption) Then
                        If CDbl(LBLMOROSO.Caption) = 0 Then
                                        DATOSCREDITO.Enabled = True
 
                            FRMCUOTAS.Enabled = True
                            MONTO.SetFocus
                        Else
                        MsgBox ("CLIENTE CON CUOTAS MOROSAS")
                        End If
                        
                Else
                        If MsgBox("CLIENTE CON CUPO INSUFICIENTE", vbOKCancel, "ATENCION") = vbOK Then
             DATOSCREDITO.Enabled = True
                        
                        rut2.SetFocus
                        End If
             
                End If
             Else
             MsgBox ("CLIENTE BLOQUEADO PASAR A OFICINA COMERCIAL POR FAVOR")
             DATOSCREDITO.Enabled = True
             
             rut2.SetFocus
             
             End If
             
             
             
             Else
             MsgBox ("CLIENTE NO CORRESPONDE A CLIENTE A CREDITO ")
             DATOSCREDITO.Enabled = True
             
             rut2.SetFocus
             End If
             
           End If
End Sub

Private Sub TIPO_GotFocus()
Call selecciona(TIPO)

End Sub

Private Sub TIPO_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
If KeyAscii = 13 Then
If TIPO.text <> "BV" And TIPO.text <> "FV" And TIPO.text <> "ZE" And TIPO.text <> "VC" And TIPO.text <> "" Then
TIPO.SetFocus
Else
cajadoc.SetFocus
End If
End If
End Sub

 Public Sub grabarcuotas()
        
        Dim campos(20, 3) As String
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
        campos(13, 0) = "rutadicional"
        campos(14, 0) = "nombreadicional"
        campos(15, 0) = "caja"
        campos(16, 0) = "cajera"
        campos(17, 0) = "pie"
        campos(18, 0) = "montoventa"
        campos(19, 0) = ""
        campos(0, 1) = empresaActiva
        campos(1, 1) = TIPO.text
        campos(2, 1) = NUMERO.text
        If rutprincipal = "" Then
        campos(3, 1) = rut2.text + lblDV.Caption
        Else
        campos(3, 1) = rutprincipal
        End If
        For K = 1 To CDbl(Cuotas.text)
        campos(4, 1) = K
        campos(5, 1) = Format(Grid2.Cell(K, 4).text, "yyyy-mm-dd")
        campos(6, 1) = Format(Grid2.Cell(K, 4).text, "yyyy-mm-dd")
        campos(7, 1) = Replace(VALORCUOTA.text, ".", "")
        campos(8, 1) = Cuotas.text
        campos(9, 1) = CDbl(CREDITO.text) / CDbl(Cuotas.text)
        campos(10, 1) = Replace(CDbl(CREDITO.text), ".", "")
        campos(11, 1) = Format(fechasistema, "yyyy-mm-dd")
        If TIPO.text <> "VC" Then
        campos(12, 1) = LEERGLOSA(TIPO.text, NUMERO.text, cajadoc.text)
        Else
        campos(12, 1) = "COMPRA " + nombreempresa
        
        
        End If
        
        campos(13, 1) = rutadicional
        campos(14, 1) = adicional
        campos(15, 1) = cajadoc.text
        campos(16, 1) = codigocajero
        campos(17, 1) = Replace(CDbl(PIE.text), ".", "")
        campos(18, 1) = Replace(CDbl(MONTO.text), ".", "")
        
        
        campos(0, 2) = "sv_cuotas_detalle"
        condicion = ""
        op = 2
        sql.response = campos
        Set sql.conexion = ventas
        sql.audit = True: sql.programaactivo = Me.Caption
        Set sql.conauditoria = conauditoria: sql.usuarioauditoria = usuarioSistema
        Call sql.sqlventas(op, condicion)
        Next K
    Call modificarut(TIPO.text, NUMERO.text, rut2.text + lblDV.Caption, cajadoc.text)
    End Sub
Public Sub IMPRIMEcredito()
    Dim K As Integer
    
    ''''''''''''''''''
    Numfic = 20
    If impresoracredito = "0" Then
    Open "impresion.txt" For Output As #Numfic
    End If
    If impresoracredito = "1" Then
    Open "COM1:4800,N,8,1,CD0,CS0,DS0,OP0,RS,TB100,RB100" For Output As #Numfic
    End If
    If impresoracredito = "2" Then
    Open "LPT1" For Output As #Numfic
    End If
    
    
    ''''''''''''''''''

    '''''''''''''''''''''
    'EMPAQUE
    '''''''''''''''''''''
    For K = 1 To 2
    Print #Numfic, Chr$(27); Chr$(64) '
    Print #Numfic, nombreempresa
    
    
    Print #Numfic, ""
    Print #Numfic, "          VALE DE CREDITO         "
    Print #Numfic, "          ================         "
    Print #Numfic,
    Print #Numfic, "TIPO    :"; TIPO.text
    Print #Numfic, "NUMERO  :"; NUMERO.text
    Print #Numfic, "FECHA   :"; Format(fechasistema, "dd-mm-yyyy")
    Print #Numfic, "CLIENTE :"; rut2.text + "-" + lblDV.Caption
    Print #Numfic, "NOMBRE  :"; lblNombre.Caption
    Print #Numfic, "DIREC.  :"; lblDireccion.Caption
    Print #Numfic,
    Print #Numfic, "MONTO VENTA   :"; Format(MONTO.text, " $ ###,###,###")
    Print #Numfic, "MONTO PIE     :"; Format(PIE.text, " $ ###,###,###")
    
    Print #Numfic, "MONTO CREDITO :"; Format(CREDITO.text, " $ ###,###,###")
    
    Print #Numfic, "YO AUTORIZO SEGUN CONTRATO PALGUIN LTDA "
    Print #Numfic, "CARGAR A MI CUENTA "
    
    Print #Numfic, Cuotas.text + " CUOTAS de " + Format(VALORCUOTA.text, "$ ##,###,###")
    Print #Numfic, "Primer vencimiento:"; DIAC + "-" + MESC + "-" + AÑOC
    Print #Numfic,
    Print #Numfic,
    Print #Numfic,
    Print #Numfic, "              _______________             "
    Print #Numfic, "               FIRMA CLIENTE             "
    Print #Numfic,
    Print #Numfic, "CAJERA(o):" + nombrecajero
    
    Print #Numfic,
    Print #Numfic,
    Print #Numfic,
    Print #Numfic,
    
    Print #Numfic, Chr(27); "i"
    Next K
    Close #Numfic
    If impresoracredito = "0" Then Shell "notepad impresion.txt"


End Sub


Public Function LEERGLOSA(TIPO, NUMERO, CAJA) As String

        Dim csql As rdoQuery
        Dim resultado As rdoResultset
        Dim i As Integer
        Dim GLOSA As String
        
        Set csql = New rdoQuery
        Set csql.ActiveConnection = ventasRubro
        
        csql.sql = "SELECT sdd.cantidad,sdd.descripcion "
        csql.sql = csql.sql & "FROM sv_documento_detalle_" & empresaActiva & " as sdd, sv_documento_cabeza_" & empresaActiva & " as sdc "
        csql.sql = csql.sql & " WHERE sdd.tipo='" + TIPO + "' and sdc.foliosii='" + NUMERO + "' and sdd.local='" + empresaActiva + "' and sdc.tipo=sdd.tipo and sdc.numero=sdd.numero and sdc.caja=sdd.caja and sdc.caja='" & CAJA & "'"
        csql.Execute
        GLOSA = ""
        If csql.RowsAffected > 0 Then
            Set resultado = csql.OpenResultset
            While Not resultado.EOF
            GLOSA = GLOSA & "/" & resultado(0) & ":" & resultado(1)
            
            resultado.MoveNext
            Wend
        End If
        LEERGLOSA = GLOSA
        Set resultado = Nothing
        csql.Close
        Set csql = Nothing
    End Function

Sub modificarut(TIPO, NUMERO, RUT, CAJA)

        Dim csql As rdoQuery
        Dim resultado As rdoResultset
        Dim i As Integer
        Dim GLOSA As String
        Dim cSql2 As rdoQuery
        
        Set cSql2 = New rdoQuery
        Set cSql2.ActiveConnection = ventasRubro
        
        cSql2.sql = "update sv_documento_cabeza_" & empresaActiva & " set rut='" + RUT + "' "
        cSql2.sql = cSql2.sql & "WHERE tipo='" + TIPO + "' and foliosii='" + NUMERO + "' and local='" + empresaActiva + "' and caja='" & CAJA & "' "
        cSql2.Execute
            Call sincronizadatos(cSql2.sql, ventasRubro)
        
        Set csql = New rdoQuery
        Set csql.ActiveConnection = ventasRubro
        
'        cSql.sql = "update sv_documento_detalle_" & empresaActiva & " set rut='" + rut + "' "
'        cSql.sql = cSql.sql & "WHERE tipo='" + TIPO + "' and numero='" + NUMERO + "' and local='" + empresaActiva + "' "
'        cSql.Execute
        
'        cSql.sql = "update sv_documento_cabeza_" & empresaActiva & " as sdc, sv_documento_detalle_" & empresaActiva & " as sdd set sdd.rut='" & rut & "' "
'        cSql.sql = cSql.sql & " where sdc.foliosii='" & NUMERO & "' and sdc.tipo='" & TIPO & "' and sdc.caja='" & caja & "' and sdc.numero=sdd.numero and sdc.tipo=sdd.tipo and sdc.caja=sdd.caja"
'        cSql.Execute
'
        
        
End Sub


Private Sub VERMOROSOS_Click()
creditoPAGOSTMP.rut2 = rut2.text
creditoPAGOSTMP.lblDV = lblDV.Caption

creditoPAGOSTMP.Show

End Sub
Sub CARGAGRILLA3()
    GRID3.Cols = 3
    GRID3.Column(0).Width = 0
    GRID3.Column(1).Width = 100
    GRID3.Column(2).Width = 100
    
    GRID3.Column(0).Locked = True
    GRID3.Column(1).Locked = True
    Grid2.Column(2).Locked = True
    
    GRID3.Cell(0, 1).text = "FECHA"
    GRID3.Cell(0, 2).text = "MONTO"
    
    GRID3.Column(1).Alignment = cellRightTop
    GRID3.Column(2).Alignment = cellRightTop
    
    GRID3.Rows = 1
    GRID3.DefaultFont.Bold = True
    GRID3.DefaultFont.Size = 10
    GRID3.ExtendLastCol = True
    
    
    
    

End Sub


Sub LEErcuotasACUMULADAS(RUT)

        Dim csql As rdoQuery
        Dim resultado As rdoResultset
        
        Set csql = New rdoQuery
        Set csql.ActiveConnection = ventas

        csql.sql = "SELECT vencimientoactual,sum(montocuota-abono) "
        csql.sql = csql.sql & "FROM sv_cuotas_detalle "
        csql.sql = csql.sql & "WHERE rut='" + RUT + "'  and ( (montocuota-abono)>0 or ((interesmora+montocuota)-abono)>0) "
        csql.sql = csql.sql & "group by vencimientoactual order by vencimientoactual "
        csql.Execute
        If csql.RowsAffected > 0 Then

            Set resultado = csql.OpenResultset
            
        GRID3.Rows = 1
        GRID3.AutoRedraw = False
        
        While Not resultado.EOF
        GRID3.Rows = GRID3.Rows + 1
        GRID3.Cell(GRID3.Rows - 1, 1).text = Format(resultado(0), "dd-mm-yyyy")
        GRID3.Cell(GRID3.Rows - 1, 2).text = Format(resultado(1), "###,###,###")
        resultado.MoveNext
        Wend
        Else
       
        End If
        
       
       
        Set resultado = Nothing
        csql.Close
        Set csql = Nothing
        GRID3.AutoRedraw = True
        GRID3.Refresh
    End Sub


