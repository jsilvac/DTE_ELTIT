VERSION 5.00
Object = "{4777436C-EB8C-4596-98A8-EBCF98683969}#1.0#0"; "FlexCell.ocx"
Object = "{ADBBAED6-B16F-40DC-80DF-B44910CBA76C}#1.0#0"; "Frame-Xp.ocx"
Begin VB.Form repactacion 
   BackColor       =   &H00AE1118&
   BorderStyle     =   0  'None
   Caption         =   "Crédito"
   ClientHeight    =   8700
   ClientLeft      =   210
   ClientTop       =   1710
   ClientWidth     =   14640
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8700
   ScaleWidth      =   14640
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrBlink 
      Interval        =   500
      Left            =   0
      Top             =   0
   End
   Begin XPFrame.FrameXp DATOSCREDITO 
      Height          =   8445
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   14460
      _ExtentX        =   25506
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
         Left            =   14040
         Style           =   1  'Graphical
         TabIndex        =   44
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   465
      End
      Begin VB.CommandButton actualiza_cliente 
         BackColor       =   &H0000FF00&
         Caption         =   "ACTUALIZAR CLIENTE"
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
         Left            =   6930
         Style           =   1  'Graphical
         TabIndex        =   43
         Top             =   1485
         Width           =   2640
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00FF8080&
         Caption         =   "DIRECTO"
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
         Left            =   4860
         TabIndex        =   42
         Top             =   1530
         Width           =   1410
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00FF8080&
         Caption         =   "MAAT"
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
         Left            =   3375
         TabIndex        =   41
         Top             =   1530
         Width           =   1230
      End
      Begin VB.TextBox CUPOAUTORIZADO 
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
         Height          =   420
         Left            =   135
         TabIndex        =   40
         Top             =   2280
         Width           =   3120
      End
      Begin VB.TextBox PIVOTE 
         Height          =   330
         Left            =   600
         MaxLength       =   2
         TabIndex        =   28
         Top             =   120
         Visible         =   0   'False
         Width           =   690
      End
      Begin VB.TextBox SUCU 
         Height          =   330
         Left            =   225
         TabIndex        =   23
         Text            =   "Text1"
         Top             =   90
         Visible         =   0   'False
         Width           =   420
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
         TabIndex        =   1
         Top             =   1485
         Width           =   465
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
         Left            =   2040
         MaxLength       =   9
         TabIndex        =   0
         Top             =   495
         Width           =   1410
      End
      Begin XPFrame.FrameXp FRMCUOTAS 
         Height          =   2025
         Left            =   120
         TabIndex        =   13
         Top             =   2715
         Width           =   9555
         _ExtentX        =   16854
         _ExtentY        =   3572
         BackColor       =   49344
         Caption         =   "Calcula Cuotas"
         CaptionEstilo3D =   1
         BackColor       =   49344
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
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
            Left            =   7680
            Style           =   1  'Graphical
            TabIndex        =   46
            Top             =   1440
            Width           =   1335
         End
         Begin VB.CommandButton Command2 
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
            Left            =   6120
            Style           =   1  'Graphical
            TabIndex        =   45
            Top             =   1440
            Width           =   1335
         End
         Begin VB.TextBox numerocuota 
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
            Height          =   435
            Left            =   2475
            TabIndex        =   39
            Top             =   840
            Width           =   2220
         End
         Begin VB.TextBox VALORCUOTA 
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
            Left            =   4830
            TabIndex        =   27
            Top             =   840
            Width           =   2130
         End
         Begin VB.TextBox AÑOC 
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
            Left            =   8400
            MaxLength       =   4
            TabIndex        =   26
            Top             =   855
            Width           =   780
         End
         Begin VB.TextBox MESC 
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
            Left            =   7740
            MaxLength       =   2
            TabIndex        =   25
            Top             =   855
            Width           =   420
         End
         Begin VB.TextBox DIAC 
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
            Left            =   7200
            MaxLength       =   2
            TabIndex        =   24
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
            Left            =   195
            TabIndex        =   22
            Top             =   840
            Width           =   2085
         End
         Begin VB.CommandButton ACEPTAR 
            BackColor       =   &H0000FF00&
            Caption         =   "ACEPTAR"
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
            Style           =   1  'Graphical
            TabIndex        =   18
            Top             =   1440
            Width           =   1335
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "TOTAL DEUDA $"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   240
            TabIndex        =   48
            Top             =   1440
            Width           =   2415
         End
         Begin VB.Label TOTALREPACTACION 
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   2640
            TabIndex        =   47
            Top             =   1440
            Width           =   1815
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
            Left            =   2475
            TabIndex        =   14
            Top             =   495
            Width           =   2220
         End
         Begin VB.Label Label4 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FF8080&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "ABONO"
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
            TabIndex        =   17
            Top             =   495
            Width           =   2130
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FF8080&
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
            ForeColor       =   &H80000008&
            Height          =   330
            Left            =   7065
            TabIndex        =   16
            Top             =   495
            Width           =   2130
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
            Left            =   4815
            TabIndex        =   15
            Top             =   495
            Width           =   2130
         End
      End
      Begin XPFrame.FrameXp FrameXp4 
         Height          =   4200
         Left            =   9990
         TabIndex        =   19
         Top             =   450
         Width           =   4335
         _ExtentX        =   7646
         _ExtentY        =   7408
         BackColor       =   49344
         Caption         =   "SIMULADOR DE CUOTAS"
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
         Begin FlexCell.Grid Grid2 
            Height          =   3840
            Left            =   45
            TabIndex        =   20
            Top             =   225
            Width           =   4200
            _ExtentX        =   7408
            _ExtentY        =   6773
            Cols            =   5
            DefaultFontSize =   8.25
            Rows            =   30
         End
      End
      Begin XPFrame.FrameXp FrameXp2 
         Height          =   3555
         Left            =   0
         TabIndex        =   29
         Top             =   4800
         Width           =   14280
         _ExtentX        =   25188
         _ExtentY        =   6271
         BackColor       =   12582912
         Caption         =   "CUOTAS "
         CaptionEstilo3D =   1
         BackColor       =   12582912
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
            Height          =   2175
            Left            =   0
            TabIndex        =   30
            Top             =   240
            Width           =   14145
            _ExtentX        =   24950
            _ExtentY        =   3836
            Cols            =   5
            DefaultFontSize =   8.25
            Rows            =   30
         End
         Begin VB.Label lblLabels 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFC0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Total Mora"
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
            Height          =   390
            Index           =   13
            Left            =   7830
            TabIndex        =   38
            Top             =   2685
            Width           =   2130
         End
         Begin VB.Label totalmora 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "$ 0"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FF00&
            Height          =   375
            Left            =   7830
            TabIndex        =   37
            Top             =   3045
            Width           =   2130
         End
         Begin VB.Label totalcuotas 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "$ 0"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FF00&
            Height          =   375
            Left            =   5490
            TabIndex        =   36
            Top             =   3045
            Width           =   2130
         End
         Begin VB.Label lblLabels 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFC0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Total Cuotas"
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
            Height          =   390
            Index           =   11
            Left            =   5490
            TabIndex        =   35
            Top             =   2685
            Width           =   2130
         End
         Begin VB.Label moroso 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "$ 0"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FF00&
            Height          =   375
            Left            =   3060
            TabIndex        =   34
            Top             =   3045
            Width           =   2325
         End
         Begin VB.Label lblLabels 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFC0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Cuotas morosas"
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
            Height          =   390
            Index           =   10
            Left            =   3060
            TabIndex        =   33
            Top             =   2685
            Width           =   2325
         End
         Begin VB.Label lblLabels 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFC0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Total Deuda"
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
            Height          =   390
            Index           =   12
            Left            =   10125
            TabIndex        =   32
            Top             =   2685
            Width           =   2130
         End
         Begin VB.Label totaldeuda 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "$ 0"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FF00&
            Height          =   375
            Left            =   10125
            TabIndex        =   31
            Top             =   3045
            Width           =   2130
         End
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
         Left            =   3555
         TabIndex        =   21
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
         TabIndex        =   12
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
         TabIndex        =   11
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
         TabIndex        =   10
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
         TabIndex        =   9
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
         TabIndex        =   8
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
         Left            =   120
         TabIndex        =   7
         Top             =   1920
         Width           =   3135
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
         Left            =   3240
         TabIndex        =   6
         Top             =   1920
         Width           =   3255
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
         Left            =   3240
         TabIndex        =   5
         Top             =   2280
         Width           =   3255
      End
      Begin VB.Label lblLabels 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Crédito Disponible"
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
         Left            =   6480
         TabIndex        =   4
         Top             =   1920
         Width           =   3135
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
         Left            =   6480
         TabIndex        =   3
         Top             =   2280
         Width           =   3135
      End
   End
End
Attribute VB_Name = "repactacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim fechas(100) As String
Public numerorepa As String



Private Sub actualiza_cliente_Click()
Dim rut As String
Dim tp As String
rut = rut2.text + lbldv.Caption
If Option2.Value = True Then tp = "T"
If Option1.Value = True Then tp = "M"
Call actualizacliente(rut, DIAPAGO.text, Replace(CUPOAUTORIZADO.text, ".", ""), tp)
End Sub
Sub actualizacliente(rut, DIAPAGO, cupo, TIPOCLIENTE)
Dim csql As rdoQuery
Dim resultado As rdoResultset
Dim i As Integer
Set csql = New rdoQuery
Set csql.ActiveConnection = ventas
csql.sql = "UPDATE sv_maestroclientes AS MC set mc.diapago='" + DIAPAGO + "',mc.cupodirecto='" + cupo + "',credito='" + TIPOCLIENTE + "' "
csql.sql = csql.sql & "WHERE mc.rut='" + rut + "' "
csql.Execute
    Call sincronizadatos(csql.sql, ventas)
csql.Close
Set csql = Nothing
End Sub
Private Sub AÑOC_KeyPress(KeyAscii As Integer)
KeyAscii = esNumero(KeyAscii)
If KeyAscii = 13 And AÑOC.text <> "" Then
CALCULATODASCUOTAS
End If
End Sub

Private Sub Command2_Click()
If MsgBox("SEGURO QUE DESEA IMPRIMIR COMPROBANTE", vbOKCancel, "ATENCION") = vbOK Then

repactacionimprimir = "si"
sw = True

Load impresionboleta

End If
End Sub

Private Sub CUPOAUTORIZADO_KeyPress(KeyAscii As Integer)
KeyAscii = esNumero(KeyAscii)
If KeyAscii = 13 And CUPOAUTORIZADO.text <> "" Then
FRMCUOTAS.Enabled = True
MONTO.SetFocus
End If
End Sub
Private Sub VALORCUOTA_GotFocus()
Call cargatexto(VALORCUOTA)
End Sub

Private Sub VALORCUOTA_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 And VALORCUOTA.text <> "" Then
TOTALREPACTACION.Caption = Str(CDbl(VALORCUOTA.text) * CDbl(numerocuota.text))
ACEPTAR.SetFocus
End If
End Sub

Private Sub ACEPTAR_Click()
eliminarcuotas
CALCULATODASCUOTAS
CALCULAR_Click
grabarcuotas
actualiza_cliente_Click
LEErcuotascliente (rut2.text + lbldv.Caption)
End Sub

Private Sub AÑOC_LostFocus()
Call esfecha(DIAC, MESC, AÑOC, "yyyy")
End Sub

Sub CALCULAR_Click()
If numerocuota.text <> "" Then
CALCULACUOTA
End If
End Sub
Private Sub Command1_Click()
Unload Me
End Sub

Private Sub numerocuota_GotFocus()
Call cargatexto(numerocuota)
End Sub

Private Sub numerocuota_KeyPress(KeyAscii As Integer)
KeyAscii = esNumero(KeyAscii)
If numerocuota.text <> "" Then
If KeyAscii = 13 And numerocuota.text <> "" And numerocuota.text <> "0" And CDbl(numerocuota.text) <= leefactorMAXIMO Then
CALCULAR_Click
End If
End If
End Sub
Private Sub DIAC_KeyPress(KeyAscii As Integer)
KeyAscii = esNumero(KeyAscii)
If KeyAscii = 13 And DIAC.text <> "" Then
DIAC.text = ceros(DIAC)
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

Private Sub Form_Load()
CARGAGRILLA
CARGAGRILLA2
Option2.Value = True
End Sub

Sub CARGAGRILLA()
    Grid1.Cols = 12
    Grid1.Column(0).Width = 0
    Grid1.Column(1).Width = 100
    Grid1.Column(2).Width = 80
    Grid1.Column(3).Width = 100
    Grid1.Column(4).Width = 60
    Grid1.Column(5).Width = 60
    Grid1.Column(6).Width = 60
    Grid1.Column(7).Width = 90
    Grid1.Column(8).Width = 60
    Grid1.Column(9).Width = 60
    Grid1.Column(10).Width = 90
    Grid1.Column(11).Width = 300
   
    
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
    Grid1.Column(11).Locked = True
    
    Grid1.Cell(0, 1).text = "DOCUMENTO"
    Grid1.Cell(0, 2).text = "N.CUOTA"
    Grid1.Cell(0, 3).text = "VENCIMIENTO"
    Grid1.Cell(0, 4).text = "CUOTA"
    Grid1.Cell(0, 5).text = "ABONO"
    Grid1.Cell(0, 6).text = "SALDO"
    Grid1.Cell(0, 7).text = "DIAS MORA"
    Grid1.Cell(0, 8).text = "INTERES"
    Grid1.Cell(0, 9).text = "TOTAL"
    Grid1.Cell(0, 10).text = "CANCELADO"
    Grid1.Cell(0, 11).text = "DETALLE COMPRAS "
    
    Grid1.Column(4).Alignment = cellRightTop
    Grid1.Column(5).Alignment = cellRightTop
    Grid1.Column(6).Alignment = cellRightTop
    Grid1.Column(7).Alignment = cellRightTop
    Grid1.Column(8).Alignment = cellRightTop
    Grid1.Column(9).Alignment = cellRightTop
    Grid1.Column(10).Alignment = cellRightTop
    
    Grid1.Column(11).Alignment = cellLeftCenter
    
    Me.Left = (Screen.Width - Me.Width) / 2
    Me.Top = (Screen.Height - Me.Height) / 2 - 950
    Grid1.Rows = 1
   
End Sub
Sub CARGAGRILLA2()
    Grid2.Cols = 5
    Grid2.Column(0).Width = 0
    Grid2.Column(1).Width = 60
    Grid2.Column(2).Width = 60
    Grid2.Column(3).Width = 60
    Grid2.Column(4).Width = 70
    
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
     
End Sub
Sub CALCULACUOTA()
Dim FACTOR As Double
Dim cuota As Double
If numerocuota.text <> "" Then
FACTOR = leefactor(numerocuota.text)
cuota = Int((CDbl(Replace(TOTALREPACTACION.Caption, ",", "")) * FACTOR) + 0.5)
VALORCUOTA.text = Format(cuota, "###,###,###")
ACEPTAR.SetFocus
CALCULATODASCUOTAS
End If

End Sub
Sub CALCULATODASCUOTAS()
Dim K As Integer
Dim cuota As Double
Dim fin As Double
Dim fecha As String
Dim mesinicio As Double
Dim añoinicio As Double

fecha = DIAC.text + "/" + MESC.text + "/" + AÑOC.text
fin = leefactorMAXIMO
mesinicio = MESC.text
añoinicio = AÑOC.text
Grid2.Rows = fin + 1
For K = 1 To fin

FACTOR = leefactor(Str(K))
cuota = Int((CDbl(Replace(TOTALREPACTACION.Caption, ",", "")) * FACTOR) + 0.5)
Grid2.Cell(K, 1).text = K

Grid2.Cell(K, 2).text = Format(cuota, "###,###,##0")
Grid2.Cell(K, 3).text = Format(cuota * K, "###,###,##0")
Grid2.Cell(K, 4).text = Format(fecha, "dd-mm-yyyy")
mesinicio = mesinicio + 1
If mesinicio > 12 Then
mesinicio = 1
añoinicio = añoinicio + 1

End If
pivote.MaxLength = 2
pivote.text = mesinicio
pivote.text = ceros(pivote)

If pivote = "02" And CDbl(DIAC.text) >= 28 Then
fecha = "28" & "-" & pivote.text & "-" & añoinicio
End If

If pivote = "02" And CDbl(DIAC.text) >= 29 And añoinicio = "2008" Or añoinicio = "2012" Or añoinicio = "2016" Then
fecha = "29" & "-" & pivote.text & "-" & añoinicio
End If

If pivote = "02" And CDbl(DIAC.text) < 28 Then
fecha = DIAC.text & "-" & pivote.text & "-" & añoinicio
End If

If pivote <> "02" Then
fecha = DIAC.text & "-" & pivote.text & "-" & añoinicio
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
mes = sumar
año = año + 1
End If
pivote.text = mes
pivote.MaxLength = 2
pivote.text = ceros(pivote)
MESC = pivote.text
AÑOC = año
End Sub
Private Sub MESC_KeyPress(KeyAscii As Integer)
KeyAscii = esNumero(KeyAscii)
If KeyAscii = 13 And MESC.text <> "" Then
MESC.text = ceros(MESC)
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
If MONTO.text <> "" Then

If KeyAscii = 13 And CDbl(MONTO.text) > 0 Then
TOTALREPACTACION.Caption = CDbl(TOTALREPACTACION.Caption) - CDbl(MONTO.text)
TOTALREPACTACION.Caption = Format(TOTALREPACTACION.Caption, "###,###,##0")
numerocuota.SetFocus
End If
End If

End Sub

Public Function LEERDOCUMENTO(TIPO, NUMERO) As Boolean

        Dim csql As rdoQuery
        Dim resultado As rdoResultset
        Dim i As Integer
        Set csql = New rdoQuery
        Set csql.ActiveConnection = ventasRubro
        
        csql.sql = "SELECT dc.rut,dc.total,mc.diapago,mc.cupodirecto,mc.cupoutilizadotmp,mc.nombre,mc.direccion,dc.abono2 "
        csql.sql = csql.sql & "FROM sv_documento_cabeza as dc," + baseVentas + ".sv_maestroclientes as mc "
        csql.sql = csql.sql & "WHERE tipo = '" & TIPO & "' and numero='" & NUMERO & "' and dc.rut=mc.rut and local='00' "
        csql.Execute
        
        If csql.RowsAffected > 0 Then
            
            Set resultado = csql.OpenResultset
            While Not resultado.EOF
        MONTODOCUMENTO.Caption = Format(resultado(1), "###,###,##0")
        abono.text = Format(resultado(7), "###,###,##0")
        DIAPAGO.text = resultado(2)
        lblCupo.Caption = Format(resultado(3), "###,###,##0")
        lblUtilizado.Caption = Format(resultado(4), "###,###,##0")
        lblDisponible.Caption = Format(resultado(3) - resultado(4), "###,###,##0")
        TOTALREPACTACION.Caption = resultado(1) - resultado(7)
        rut2.text = Mid(resultado(0), 1, 9)
        lbldv.Caption = Mid(resultado(0), 10, 1)
        lblnombre.Caption = resultado(5)
        LBLDIRECCION.Caption = resultado(6)
        
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
Public Function LEErcuotas(TIPO, NUMERO) As Boolean

        Dim csql As rdoQuery
        Dim resultado As rdoResultset
        Dim i As Integer
        Dim diasmora As Double
        Dim saldo As Double
        Dim interes As Double
        Dim montototal As Double
        
        
        Set csql = New rdoQuery
        Set csql.ActiveConnection = ventas

        csql.sql = "SELECT dc.rut,dc.montocredito,mc.diapago,mc.cupodirecto,mc.cupoutilizadotmp,mc.nombre,mc.direccion,dc.cantidadcuotas,dc.montocuota,dc.vencimientooriginal,dc.tipo,dc.numero,dc.abono,dc.numerocuota,dc.vencimientoactual "
        csql.sql = csql.sql & "FROM sv_cuotas_detalle as dc," + baseVentas + ".sv_maestroclientes as mc "
        csql.sql = csql.sql & "WHERE tipo = '" & TIPO & "' and numero='" & NUMERO & "' and dc.rut=mc.rut  "
        csql.sql = csql.sql & "order by vencimientooriginal "
        csql.Execute
        
        If csql.RowsAffected > 0 Then

            Set resultado = csql.OpenResultset
            
       
        DIAPAGO.text = resultado(2)
        CUPOAUTORIZADO.text = Format(resultado(3), "###,###,##0")
        lblUtilizado.Caption = Format(resultado(4), "###,###,##0")
        lblDisponible.Caption = Format(resultado(3) - resultado(4), "###,###,##0")
        TOTALREPACTACION.Caption = resultado(1)
        rut2.text = Mid(resultado(0), 1, 9)
        lbldv.Caption = Mid(resultado(0), 10, 1)
        lblnombre.Caption = resultado(5)
        LBLDIRECCION.Caption = resultado(6)
        numerocuota.text = resultado(7)
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
        montototal = montototal + total
        
        Grid1.Cell(Grid1.Rows - 1, 6).text = Format(saldo, "###,###,###")
        Grid1.Cell(Grid1.Rows - 1, 7).text = diasmora
        Grid1.Cell(Grid1.Rows - 1, 8).text = interes
        Grid1.Cell(Grid1.Rows - 1, 9).text = total
        
            
            
            resultado.MoveNext
            Wend
        Else
        LEErcuotas = False
        TOTALREPACTACION.Caption = Format(montototal, "###,###,##0")
        End If
        Set resultado = Nothing
        csql.Close
        Set csql = Nothing
    End Function

Public Function LEERCLIENTE(rut) As Boolean

        Dim csql As rdoQuery
        Dim resultado As rdoResultset
        Dim i As Integer
        Set csql = New rdoQuery
        Set csql.ActiveConnection = ventas
        
        csql.sql = "SELECT mc.diapago,mc.cupodirecto,mc.cupoutilizadodirecto,mc.nombre,mc.direccion "
        csql.sql = csql.sql & "FROM sv_maestroclientes as mc "
        csql.sql = csql.sql & "WHERE mc.rut='" + rut + "' "
        csql.Execute
        If csql.RowsAffected > 0 Then
            Set resultado = csql.OpenResultset
            While Not resultado.EOF
        DIAPAGO.text = resultado(0)
        CUPOAUTORIZADO.text = Format(resultado(1), "###,###,##0")
        lblUtilizado.Caption = Format(resultado(2), "###,###,##0")
        lblDisponible.Caption = Format(resultado(1) - resultado(2), "###,###,##0")
        lblnombre.Caption = resultado(3)
        LBLDIRECCION.Caption = resultado(4)
        
        If resultado(1) > 0 Then
                LEERCLIENTE = True
                
        Else
                LEERCLIENTE = True
        
        End If
        
            resultado.MoveNext
            Wend
        Else
        LEERCLIENTE = False
        
        End If
        Set resultado = Nothing
        csql.Close
        Set csql = Nothing
    End Function


Private Sub retorno_Click()
rut2.text = ""
lbldv.Caption = ""
lblnombre.Caption = ""
LBLDIRECCION = ""
DIAPAGO.text = ""
CUPOAUTORIZADO.text = "$ 0"
lblUtilizado.Caption = "$ 0"
lblDisponible.Caption = "$ 0"
MONTO.text = ""
numerocuota.text = ""
numerocuota.text = ""
VALORCUOTA.text = ""
DIAC.text = ""
MESC.text = ""
AÑOC.text = ""
Grid1.Rows = 1
Grid2.Rows = 1
TOTALREPACTACION.Caption = ""

End Sub

Private Sub rut2_GotFocus()
Call selecciona(rut2)
End Sub

Private Sub rut2_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF2 Then
            Call ayudaCliente(rut2, SUCU, lbldv)
  End If
End Sub

Private Sub rut2_KeyPress(KeyAscii As Integer)
          KeyAscii = esNumero(KeyAscii)
           If KeyAscii = 13 And rut2.text <> "" And Val(rut2.text) <> 0 Then
             rut2.text = ceros(rut2)
             lbldv.Caption = rut(rut2.text)
             If LEERCLIENTE(rut2.text + lbldv.Caption) = True Then
             LEErcuotascliente (rut2.text + lbldv.Caption)
             FRMCUOTAS.Enabled = True
             MONTO.SetFocus
             End If
             End If
             
            
        
End Sub

 Public Sub grabarcuotas()
        
        Dim CAMPOS(12, 3) As String
        Dim op As Integer
        Dim K As Integer
       
        Set sql = New sqlventas.sqlventa
        CAMPOS(0, 0) = "local"
        CAMPOS(1, 0) = "tipo"
        CAMPOS(2, 0) = "numero"
        CAMPOS(3, 0) = "rut"
        CAMPOS(4, 0) = "numerocuota"
        CAMPOS(5, 0) = "vencimientooriginal"
        CAMPOS(6, 0) = "vencimientoactual"
        CAMPOS(7, 0) = "montocuota"
        CAMPOS(8, 0) = "cantidadcuotas"
        CAMPOS(9, 0) = "capitalcuota"
        CAMPOS(10, 0) = "montocredito"
        CAMPOS(11, 0) = "glosacompra"
        CAMPOS(12, 0) = "caja"
        CAMPOS(13, 0) = "cajera"
        CAMPOS(14, 0) = ""
        numerorepa = Folioingresomanual
        CAMPOS(0, 1) = empresaActiva
        CAMPOS(1, 1) = "IM"
        CAMPOS(2, 1) = Folioingresomanual
        CAMPOS(3, 1) = rut2.text + lbldv.Caption
        For K = 1 To CDbl(numerocuota.text)
        CAMPOS(4, 1) = K
        CAMPOS(5, 1) = Format(Grid2.Cell(K, 4).text, "yyyy-mm-dd")
        CAMPOS(6, 1) = Format(Grid2.Cell(K, 4).text, "yyyy-mm-dd")
        CAMPOS(7, 1) = Replace(VALORCUOTA.text, ".", "")
        CAMPOS(8, 1) = numerocuota.text
        CAMPOS(9, 1) = CDbl(TOTALREPACTACION.Caption) / CDbl(numerocuota.text)
        CAMPOS(10, 1) = Replace(TOTALREPACTACION.Caption, ".", "")
        CAMPOS(11, 1) = "REPACTACION"
        CAMPOS(0, 2) = "sv_cuotas_detalle"
        condicion = ""
        op = 2
        sql.response = CAMPOS
        Set sql.conexion = ventas
        Call sql.sqlventas(op, condicion)
        Next K
    End Sub


Sub LEErcuotascliente(rut)

        Dim csql As rdoQuery
        Dim resultado As rdoResultset
        Dim i As Integer
        Dim diasmora As Double
        Dim saldo As Double
        Dim interes As Double
        Dim t1 As Double
        Dim t2 As Double
        Dim totalusado As Double
        Dim totalizado As Double
        
        Set csql = New rdoQuery
        Set csql.ActiveConnection = ventas

        csql.sql = "SELECT *  "
        csql.sql = csql.sql & "FROM sv_cuotas_detalle "
        csql.sql = csql.sql & "WHERE rut='" + rut + "'  and ( (montocuota-abono)>0 or ((interesmora+montocuota)-abono)>0) "
        csql.sql = csql.sql & "order by vencimientoactual "
        csql.Execute
        
        If csql.RowsAffected > 0 Then

            Set resultado = csql.OpenResultset
            
        Grid1.Rows = 1
       Grid1.AutoRedraw = False
        
        totalusado = 0
        moratotal = 0
        While Not resultado.EOF
        Grid1.Rows = Grid1.Rows + 1
        Grid1.Cell(Grid1.Rows - 1, 1).text = resultado(1) & " " & resultado(2)
        Grid1.Cell(Grid1.Rows - 1, 2).text = resultado(4) & " / " & resultado(12)
        Grid1.Cell(Grid1.Rows - 1, 3).text = Format(resultado(6), "dd/mm/yyyy")
        Grid1.Cell(Grid1.Rows - 1, 4).text = Format(resultado(7), "###,###,###")
        Grid1.Cell(Grid1.Rows - 1, 5).text = Format(resultado(8), "###,###,###")
        saldo = resultado(7) - resultado(8)
        tazainteresmora = leerInteresMora("00")
        diasmora = DateDiff("d", resultado(6), fechasistema)
        
        
        If diasmora <= diasgracia Then diasmora = 0
       
        
        interes = Round(saldo * ((tazainteresmora * diasmora) / 100), 0)
        
        t1 = t1 + saldo
        t2 = t2 + interes
        
        total = saldo + interes
        
        Grid1.Cell(Grid1.Rows - 1, 6).text = Format(saldo, "###,###,###")
        
        Grid1.Cell(Grid1.Rows - 1, 7).text = diasmora
        Grid1.Cell(Grid1.Rows - 1, 8).text = interes
        Grid1.Cell(Grid1.Rows - 1, 9).text = Format(total, "###,###,###")
        Grid1.Cell(Grid1.Rows - 1, 10).text = "0"
        
        Grid1.Cell(Grid1.Rows - 1, 11).text = resultado(13)
        
        totalusado = totalusado + total
        If interes <> 0 Then moratotal = moratotal + total
        
            
            resultado.MoveNext
            Wend
        Else
       
        End If
        Set resultado = Nothing
        csql.Close
        Set csql = Nothing
        Grid1.AutoRedraw = True
        Grid1.Refresh
'        If totalusado = 0 Then
'        lblUtilizado.Caption = 0
'        lblDisponible.Caption = 0
'        totaldeuda.Caption = 0
'        moroso.Caption = 0
'        Else
        
        lblUtilizado.Caption = Format(totalusado, "###,###,##0")
        lblDisponible.Caption = Format(CDbl(CUPOAUTORIZADO.text) - totalusado, "###,###,##0")
        totaldeuda.Caption = Format(totalusado, "###,###,##0")
        moroso.Caption = Format(moratotal, "###,###,##0")
'        End If
        
        TOTALREPACTACION.Caption = Format(totalusado, "###,###,##0")
        totalcuotas.Caption = Format(t1, "###,###,##0")
        totalmora.Caption = Format(t2, "###,###,##0")
        
    End Sub

Public Function Folioingresomanual() As String
    
    Dim op As Integer
    Dim sql As New sqlventas.sqlventa
    Dim CAMPOS(3, 3) As String
    
    CAMPOS(0, 0) = "IFNULL(MAX(numero) + 1,'0000000001')"
    CAMPOS(1, 0) = ""
    CAMPOS(0, 2) = "sv_cuotas_detalle"
    condicion = " local='00' and tipo='IM'"
    op = 5
    sql.response = CAMPOS
    Set sql.conexion = ventas
    Call sql.sqlventas(op, condicion)
    If sql.Status = 0 Then
        If sql.response(0, 3) <> "" And sql.response(0, 3) <> "0" Then
            Folioingresomanual = sql.response(0, 3)
        Else
            Folioingresomanual = "0000000001"
        End If
    End If
End Function
Sub eliminarcuotas()
Dim csql As New rdoQuery
Dim tabla As String

Set csql.ActiveConnection = ventas
tabla = "update sv_cuotas_detalle set tipo='NL' where rut='" & rut2.text & lbldv.Caption & "' "
csql.sql = tabla
csql.Execute
    Call sincronizadatos(csql.sql, ventas)

End Sub
