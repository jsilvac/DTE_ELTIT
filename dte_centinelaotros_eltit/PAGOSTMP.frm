VERSION 5.00
Object = "{4777436C-EB8C-4596-98A8-EBCF98683969}#1.0#0"; "Flexcell.ocx"
Object = "{ADBBAED6-B16F-40DC-80DF-B44910CBA76C}#1.0#0"; "Frame-Xp.ocx"
Begin VB.Form creditoPAGOSTMP 
   BackColor       =   &H00800000&
   BorderStyle     =   0  'None
   Caption         =   "Crédito"
   ClientHeight    =   9705
   ClientLeft      =   210
   ClientTop       =   1710
   ClientWidth     =   14640
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9705
   ScaleWidth      =   14640
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin XPFrame.FrameXp FRMCUOTAS 
      Height          =   2670
      Left            =   1080
      TabIndex        =   57
      Top             =   2880
      Width           =   10815
      _ExtentX        =   19076
      _ExtentY        =   4710
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
      Begin VB.CommandButton Command4 
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
         Left            =   5310
         Style           =   1  'Graphical
         TabIndex        =   77
         Top             =   1980
         Width           =   2775
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
         TabIndex        =   73
         Top             =   495
         Width           =   465
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
         Locked          =   -1  'True
         TabIndex        =   66
         Top             =   1305
         Width           =   1005
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
         Left            =   2205
         Style           =   1  'Graphical
         TabIndex        =   65
         Top             =   1980
         Width           =   2775
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
         Left            =   135
         Locked          =   -1  'True
         TabIndex        =   64
         Top             =   1305
         Width           =   1995
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
         Height          =   420
         Left            =   8775
         Locked          =   -1  'True
         MaxLength       =   2
         TabIndex        =   63
         Top             =   1305
         Width           =   420
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
         Height          =   420
         Left            =   9225
         Locked          =   -1  'True
         MaxLength       =   2
         TabIndex        =   62
         Top             =   1305
         Width           =   420
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
         Height          =   420
         Left            =   9675
         Locked          =   -1  'True
         MaxLength       =   4
         TabIndex        =   61
         Top             =   1305
         Width           =   780
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
         Left            =   7155
         Locked          =   -1  'True
         TabIndex        =   60
         Top             =   1305
         Width           =   1545
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
         Left            =   2295
         MaxLength       =   6
         TabIndex        =   59
         Top             =   1305
         Visible         =   0   'False
         Width           =   1680
      End
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
         TabIndex        =   58
         Top             =   1305
         Width           =   1680
      End
      Begin FlexCell.Grid Grid4 
         Height          =   60
         Left            =   180
         TabIndex        =   78
         Top             =   1755
         Visible         =   0   'False
         Width           =   10500
         _ExtentX        =   18521
         _ExtentY        =   106
         Cols            =   5
         DefaultFontSize =   8.25
         Rows            =   30
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
         Left            =   3420
         TabIndex        =   76
         Top             =   495
         Width           =   5460
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
         Left            =   2880
         TabIndex        =   75
         Top             =   495
         Width           =   465
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
         Index           =   14
         Left            =   180
         TabIndex        =   74
         Top             =   450
         Width           =   1800
      End
      Begin VB.Label Label9 
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
         Left            =   7155
         TabIndex        =   72
         Top             =   945
         Width           =   1545
      End
      Begin VB.Label Label8 
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
         Left            =   8775
         TabIndex        =   71
         Top             =   945
         Width           =   1725
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FF8080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "REPACTACION"
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
         Left            =   135
         TabIndex        =   70
         Top             =   945
         Width           =   2040
      End
      Begin VB.Label Label4 
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
         Left            =   6075
         TabIndex        =   69
         Top             =   945
         Width           =   1005
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
         Left            =   2295
         TabIndex        =   68
         Top             =   945
         Width           =   1680
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
         Left            =   4230
         TabIndex        =   67
         Top             =   945
         Width           =   1680
      End
   End
   Begin VB.CommandButton cancelar 
      BackColor       =   &H0000FF00&
      Caption         =   "PAGAR CUOTAS"
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
      Left            =   180
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   8190
      Width           =   2085
   End
   Begin VB.Timer tmrBlink 
      Interval        =   500
      Left            =   0
      Top             =   0
   End
   Begin XPFrame.FrameXp DATOSCREDITO 
      Height          =   9525
      Left            =   45
      TabIndex        =   1
      Top             =   45
      Width           =   14460
      _ExtentX        =   25506
      _ExtentY        =   16801
      BackColor       =   16744576
      Caption         =   "Cancelacion de Cuotas"
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
      Begin XPFrame.FrameXp frmcancelado 
         Height          =   2670
         Left            =   5895
         TabIndex        =   17
         Top             =   405
         Width           =   8565
         _ExtentX        =   15108
         _ExtentY        =   4710
         BackColor       =   16744576
         Caption         =   "CANCELACION"
         CaptionEstilo3D =   1
         BackColor       =   16744576
         ForeColor       =   1563884
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
         Begin VB.CheckBox Check1 
            BackColor       =   &H00FF8080&
            Caption         =   "Vuelto"
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
            Left            =   45
            TabIndex        =   51
            Top             =   2790
            Width           =   1095
         End
         Begin XPFrame.FrameXp Vuelto 
            Height          =   2040
            Left            =   -90
            TabIndex        =   47
            Top             =   2745
            Visible         =   0   'False
            Width           =   2895
            _ExtentX        =   5106
            _ExtentY        =   3598
            BackColor       =   49344
            Caption         =   "Vuelto"
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
            Alignment       =   1
            Begin VB.TextBox efectivo 
               Alignment       =   1  'Right Justify
               BackColor       =   &H80000007&
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H0000FF00&
               Height          =   375
               Left            =   90
               TabIndex        =   52
               Top             =   675
               Width           =   2670
            End
            Begin VB.TextBox vueltototal 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00000000&
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H0000FF00&
               Height          =   375
               Left            =   45
               Locked          =   -1  'True
               TabIndex        =   49
               Top             =   1530
               Width           =   2715
            End
            Begin VB.Label lblLabels 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00FFC0C0&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "Vuelto"
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
               Left            =   45
               TabIndex        =   50
               Top             =   1125
               Width           =   2715
            End
            Begin VB.Label lblLabels 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00FFC0C0&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "Cancelado Con"
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
               Left            =   45
               TabIndex        =   48
               Top             =   270
               Width           =   2715
            End
         End
         Begin VB.CommandButton Command3 
            BackColor       =   &H0080FFFF&
            Caption         =   "RETORNO SIN CANCELAR"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   360
            MaskColor       =   &H0080C0FF&
            Style           =   1  'Graphical
            TabIndex        =   42
            Top             =   2250
            UseMaskColor    =   -1  'True
            Width           =   3750
         End
         Begin VB.TextBox PAGOCUOTAS 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00000000&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   375
            Left            =   1440
            Locked          =   -1  'True
            TabIndex        =   26
            Top             =   1800
            Width           =   2715
         End
         Begin VB.TextBox PAGOMORA 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00000000&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   375
            Left            =   4275
            Locked          =   -1  'True
            TabIndex        =   25
            Top             =   1800
            Width           =   2715
         End
         Begin VB.CommandButton Command2 
            BackColor       =   &H0080FFFF&
            Caption         =   "ACEPTAR E IMPRIMIR RECIBO"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   4320
            MaskColor       =   &H0080C0FF&
            Style           =   1  'Graphical
            TabIndex        =   24
            Top             =   2250
            UseMaskColor    =   -1  'True
            Width           =   3750
         End
         Begin VB.TextBox diferencia 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00000000&
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000C000&
            Height          =   375
            Left            =   7065
            TabIndex        =   22
            Top             =   2925
            Visible         =   0   'False
            Width           =   2715
         End
         Begin VB.TextBox SELECCIONADO 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00000000&
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000C000&
            Height          =   375
            Left            =   4230
            TabIndex        =   20
            Top             =   2925
            Visible         =   0   'False
            Width           =   2715
         End
         Begin VB.TextBox CANCELADO 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00000000&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   24
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FF00&
            Height          =   690
            Left            =   765
            TabIndex        =   18
            Top             =   675
            Width           =   6675
         End
         Begin VB.Label lblLabels 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "TOTAL CUOTAS"
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
            Left            =   1440
            TabIndex        =   28
            Top             =   1395
            Width           =   2715
         End
         Begin VB.Label lblLabels 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "INTERES MORA"
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
            Index           =   9
            Left            =   4275
            TabIndex        =   27
            Top             =   1395
            Width           =   2715
         End
         Begin VB.Label lblLabels 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Diferencia"
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
            Index           =   8
            Left            =   7065
            TabIndex        =   23
            Top             =   2520
            Visible         =   0   'False
            Width           =   2715
         End
         Begin VB.Label lblLabels 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Seleccionado"
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
            Index           =   7
            Left            =   4230
            TabIndex        =   21
            Top             =   2520
            Visible         =   0   'False
            Width           =   2715
         End
         Begin VB.Label lblLabels 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Total a Cancelar"
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
            Height          =   300
            Index           =   6
            Left            =   180
            TabIndex        =   19
            Top             =   405
            Width           =   7935
         End
      End
      Begin XPFrame.FrameXp FrameXp4 
         Height          =   2625
         Left            =   6255
         TabIndex        =   15
         Top             =   450
         Width           =   8115
         _ExtentX        =   14314
         _ExtentY        =   4630
         BackColor       =   12582912
         Caption         =   "HISTORICO DE PAGOS"
         CaptionEstilo3D =   1
         BackColor       =   12582912
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Begin FlexCell.Grid Grid2 
            Height          =   2040
            Left            =   90
            TabIndex        =   16
            Top             =   405
            Width           =   7935
            _ExtentX        =   13996
            _ExtentY        =   3598
            Cols            =   5
            DefaultFontSize =   8.25
            Rows            =   30
         End
      End
      Begin VB.TextBox CONDO2 
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
         Left            =   5220
         Locked          =   -1  'True
         MaxLength       =   3
         TabIndex        =   46
         Top             =   2655
         Width           =   645
      End
      Begin VB.TextBox CONDO1 
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
         Left            =   5220
         Locked          =   -1  'True
         MaxLength       =   3
         TabIndex        =   45
         Top             =   2250
         Width           =   645
      End
      Begin VB.TextBox PIVOTE 
         Height          =   330
         Left            =   630
         MaxLength       =   2
         TabIndex        =   13
         Top             =   135
         Visible         =   0   'False
         Width           =   690
      End
      Begin VB.TextBox SUCU 
         Height          =   330
         Left            =   225
         TabIndex        =   12
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
         Left            =   1710
         MaxLength       =   9
         TabIndex        =   0
         Top             =   480
         Width           =   1410
      End
      Begin VB.CommandButton Command1 
         Caption         =   "X"
         Height          =   330
         Left            =   13950
         TabIndex        =   10
         Top             =   45
         Width           =   465
      End
      Begin XPFrame.FrameXp FrameXp2 
         Height          =   6315
         Left            =   45
         TabIndex        =   29
         Top             =   3105
         Width           =   14280
         _ExtentX        =   25188
         _ExtentY        =   11139
         BackColor       =   12582912
         Caption         =   "CUOTAS PENDIENTES"
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
         Begin VB.CommandButton repactar 
            BackColor       =   &H00FF8080&
            Caption         =   "REPACTACION"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   465
            Left            =   45
            Style           =   1  'Graphical
            TabIndex        =   56
            Top             =   5490
            Width           =   2085
         End
         Begin FlexCell.Grid GRID3 
            Height          =   1365
            Left            =   7830
            TabIndex        =   55
            Top             =   4815
            Width           =   6135
            _ExtentX        =   10821
            _ExtentY        =   2408
            Cols            =   5
            DefaultFontSize =   8.25
            Rows            =   30
         End
         Begin XPFrame.FrameXp FrameXp1 
            Height          =   1590
            Left            =   7830
            TabIndex        =   54
            Top             =   4545
            Width           =   6135
            _ExtentX        =   10821
            _ExtentY        =   2805
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
         End
         Begin VB.CommandButton RETORNAR 
            BackColor       =   &H000000FF&
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
            Left            =   90
            Style           =   1  'Graphical
            TabIndex        =   41
            Top             =   4590
            Width           =   2085
         End
         Begin FlexCell.Grid Grid1 
            Height          =   4290
            Left            =   90
            TabIndex        =   30
            Top             =   270
            Width           =   14145
            _ExtentX        =   24950
            _ExtentY        =   7567
            Cols            =   5
            DefaultFontSize =   8.25
            Rows            =   30
         End
         Begin VB.Label total4 
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
            Left            =   4815
            TabIndex        =   38
            Top             =   5805
            Width           =   2310
         End
         Begin VB.Label lblLabels 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFC0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Total General"
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
            Left            =   2430
            TabIndex        =   37
            Top             =   5805
            Width           =   2310
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
            Index           =   1
            Left            =   2430
            TabIndex        =   36
            Top             =   4590
            Width           =   2325
         End
         Begin VB.Label total1 
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
            Left            =   4815
            TabIndex        =   35
            Top             =   4590
            Width           =   2325
         End
         Begin VB.Label lblLabels 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFC0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Total Interes"
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
            Index           =   2
            Left            =   2430
            TabIndex        =   34
            Top             =   4995
            Width           =   2310
         End
         Begin VB.Label total2 
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
            Left            =   4815
            TabIndex        =   33
            Top             =   4995
            Width           =   2310
         End
         Begin VB.Label total3 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H000000FF&
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
            ForeColor       =   &H0000FFFF&
            Height          =   375
            Left            =   4815
            TabIndex        =   32
            Top             =   5400
            Width           =   2310
         End
         Begin VB.Label lbl11 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFC0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Total a Pagar"
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
            Left            =   2430
            TabIndex        =   31
            Top             =   5400
            Width           =   2310
         End
      End
      Begin VB.Label fechareal 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   630
         TabIndex        =   53
         Top             =   45
         Width           =   3975
      End
      Begin VB.Label Label3 
         BackColor       =   &H000080FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "CONDONACION INTERES MORA%"
         Height          =   330
         Left            =   2520
         TabIndex        =   44
         Top             =   2655
         Width           =   2670
      End
      Begin VB.Label Label2 
         BackColor       =   &H000080FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "CONDONACION PRONTO PAGO %"
         Height          =   330
         Left            =   2520
         TabIndex        =   43
         Top             =   2250
         Width           =   2670
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "FOLIO"
         Height          =   240
         Left            =   3690
         TabIndex        =   40
         Top             =   315
         Width           =   1950
      End
      Begin VB.Label folio 
         BackColor       =   &H80000007&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C000&
         Height          =   375
         Left            =   3690
         TabIndex        =   39
         Top             =   540
         Width           =   1950
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
         Left            =   3150
         TabIndex        =   11
         Top             =   450
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
         TabIndex        =   9
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
         Left            =   90
         TabIndex        =   8
         Top             =   945
         Width           =   5700
      End
      Begin VB.Label lblLabels 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Crédito Autorizado"
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
         Index           =   3
         Left            =   90
         TabIndex        =   7
         Top             =   1395
         Width           =   2325
      End
      Begin VB.Label lblCupo 
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
         Left            =   90
         TabIndex        =   6
         Top             =   1755
         Width           =   2325
      End
      Begin VB.Label lblLabels 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Crédito Utilizado"
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
         Index           =   4
         Left            =   90
         TabIndex        =   5
         Top             =   2250
         Width           =   2310
      End
      Begin VB.Label lblUtilizado 
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
         Left            =   90
         TabIndex        =   4
         Top             =   2610
         Width           =   2310
      End
      Begin VB.Label lblLabels 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Crédito Disponible"
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
         Index           =   5
         Left            =   2610
         TabIndex        =   3
         Top             =   1395
         Width           =   2460
      End
      Begin VB.Label lblDisponible 
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
         Height          =   420
         Left            =   2610
         TabIndex        =   2
         Top             =   1755
         Width           =   2460
      End
   End
End
Attribute VB_Name = "creditoPAGOSTMP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim fechas(100) As String
Dim totalusado As Double
Dim moratotal As Double
Dim deutatotal As Double
Dim tazainteresmora As Double
Dim mcancelado As Double
Dim mdiferencia As Double
Dim mseleccionado As Double
Dim REPACTACIONES As Boolean



Private Sub ACEPTAR_Click()
If VALORCUOTA.text <> "0" And VALORCUOTA.text <> "" Then
CALCULAPRIMERVENCIMIENTO
CALCULACUOTA

CUOTAS.SetFocus
GENERAPAGO
Call grabarcuotas(FOLIO.Caption)

End If

FRMCUOTAS.Visible = False

creditotmppago.Show

End Sub

Private Sub CANCELADO_KeyPress(KeyAscii As Integer)
KeyAscii = esNumero(KeyAscii)
    If KeyAscii = 13 And CANCELADO.text <> "" Then
        If CANCELADO.text = "" Then CANCELADO.text = "0"
        CANCELADO.text = Format(CANCELADO.text, "###,###,##0")
        If CDbl(CANCELADO.text) < CDbl(total2.Caption) Then
            MsgBox ("MONTO CANCELADO INFERIOR A LOS INTERESES " + Format(total2.Caption, "$ ###,###,###"))
            CANCELADO.SetFocus
            Exit Sub
        End If
        
        If CDbl(CANCELADO.text) <= CDbl(total4.Caption) Then
            seleccionacuotas
            Exit Sub
        End If
        If CDbl(CANCELADO.text) > CDbl(total4.Caption) Then
            MsgBox ("MONTO CANCELADO SUPERA LA DEUDA TOTAL " + Format(total4.Caption, "$ ###,###,###"))
            CANCELADO.SetFocus
        End If
        
End If

End Sub
Sub seleccionacuotas()
Dim monto2 As Double
Dim saldo2 As Double
mcancelado = CDbl(CANCELADO.text)

monto2 = CDbl(CANCELADO.text) - CDbl(total2.Caption)
saldo2 = monto2
For K = 1 To Grid1.Rows - 1
Grid1.Cell(K, 10).text = "0"

Next K

For K = 1 To Grid1.Rows - 1

    If saldo2 > 0 Then
        
            If saldo2 > CDbl(Grid1.Cell(K, 6).text) Then
                Grid1.Cell(K, 10).text = Grid1.Cell(K, 6).text
            Else
                Grid1.Cell(K, 10).text = saldo2
            End If
    saldo2 = saldo2 - CDbl(Grid1.Cell(K, 6).text)
    End If

Next K
SUMACANCELACION
End Sub
Sub SUMACANCELACION()
Dim PCUOTAS As Double
Dim PMORA As Double

mseleccionado = 0
PCUOTAS = 0
PMORA = 0
For K = 1 To Grid1.Rows - 1
mseleccionado = mseleccionado + CDbl(Grid1.Cell(K, 10).text)
    If CDbl(Grid1.Cell(K, 10).text) <> 0 Then
        If CDbl(Grid1.Cell(K, 10).text) <= CDbl(Grid1.Cell(K, 9).text) Then
        PCUOTAS = PCUOTAS + CDbl(Grid1.Cell(K, 10).text)
        Else
        PCUOTAS = PCUOTAS + CDbl(Grid1.Cell(K, 9).text)
        End If
        
       
    End If
    
Next K
mdiferencia = mcancelado - mseleccionado + CDbl(totalmora)
SELECCIONADO.text = Format(mseleccionado, "###,###,###")
diferencia.text = Format(mdiferencia, "###,###,###")
PAGOCUOTAS.text = Format(PCUOTAS, "###,###,###")
PAGOMORA.text = Format(total2.Caption, "###,###,###")

End Sub

Private Sub CANCELAR_Click()
FOLIO.Caption = leerUltimoFolioPAGO()
  
  If LEERCLIENTE(rut2.text + lbldv.Caption) = True Then
If total4.Caption <> "0" Then
REPACTACIONES = False


frmcancelado.Visible = True
CANCELADO.Enabled = True
CANCELADO.text = total3.Caption
CANCELADO.SetFocus
CANCELAR.Enabled = False
Grid1.Column(10).Locked = False


Else
rut2.SetFocus

End If
End If
End Sub

Private Sub Command1_Click()
Unload Me

End Sub




Private Sub Command2_Click()

If CDbl(CANCELADO.text) > CDbl(total4.Caption) Then
            MsgBox ("MONTO CANCELADO SUPERA LA DEUDA TOTAL " + Format(total4.Caption, "$ ###,###,###"))
            CANCELADO.SetFocus
Exit Sub
End If

        
    If PAGOCUOTAS.text <> "" Then
        seleccionacuotas
    If MsgBox("desea realmente realizar el pago", vbYesNo) = vbYes Then

        GENERAPAGO
        imprimepago
        Call RETORNAR_Click
    
    End If

    Else
    MsgBox "Agregar alguna Cuota para cancelar", vbOKOnly, "Atencion"
    End If




End Sub

Sub imprimepago()

creditotmppago.Show vbModal


End Sub
Sub GENERAPAGO()
  
        Dim CAMPOS(12, 3) As String
        Dim op As Integer
        Dim K As Integer
        Dim TIPO As String
        Dim NUMERO As String
        Dim cuota As String
        Dim abono As Double
        Dim intereses As Double
        Set sql = New sqlventas.sqlventa
        CAMPOS(0, 0) = "abono"
        CAMPOS(1, 0) = "interesmora"
        CAMPOS(2, 0) = "vencimientoactual"
        CAMPOS(3, 0) = ""
        For K = 1 To Grid1.Rows - 1
         If Grid1.Cell(K, 5).text = "" Then Grid1.Cell(K, 5).text = "0"
         abono = CDbl(Replace(Grid1.Cell(K, 10).text, ".", "")) + CDbl(Replace(Grid1.Cell(K, 5).text, ".", ""))
         If abono <> 0 Or Grid1.Cell(K, 7).text <> "0" Then
         intereses = CDbl(Replace(Grid1.Cell(K, 8).text, ".", ""))
         CAMPOS(0, 1) = abono
         CAMPOS(1, 1) = intereses
         TIPO = Mid(Grid1.Cell(K, 1).text, 1, 2)
         NUMERO = Mid(Grid1.Cell(K, 1).text, 4, 10)
         cuota = Mid(Grid1.Cell(K, 2).text, 1, 2)
         cuota = Replace(cuota, "/", "")
        
        If Grid1.Cell(K, 7).text <> "0" Then
        CAMPOS(2, 1) = Format(fechasistema, "yyyy-mm-dd")
        Else
        CAMPOS(2, 1) = Format(Grid1.Cell(K, 3).text, "yyyy-mm-dd")
        
        End If
        
        
        CAMPOS(0, 2) = "sv_cuotas_detalle"
        
        
        condicion = "tipo='" & TIPO & "' and numero='" & NUMERO & "' and numerocuota='" & cuota & "' and rut='" + rut2.text + lbldv.Caption + "' "
        op = 3
        sql.response = CAMPOS
        Set sql.conexion = ventas
        sql.audit = True: sql.programaactivo = Me.Caption
        sql.programaactivo = Me.Caption
        Set sql.conauditoria = conauditoria: sql.usuarioauditoria = usuarioSistema
        Call sql.sqlventas(op, condicion)
        End If
        Next K
        grabacabeza
        grabadetalle
        
End Sub
Sub grabacabeza()
        
        Dim CAMPOS(12, 3) As String
        Dim op As Integer
        Dim K As Integer
        Dim TIPO As String
        Dim NUMERO As String
        Dim cuota As String
        Dim abono As Double
        
        Set sql = New sqlventas.sqlventa
        CAMPOS(0, 0) = "local"
        CAMPOS(1, 0) = "numero"
        CAMPOS(2, 0) = "rut"
        CAMPOS(3, 0) = "fecha"
        CAMPOS(4, 0) = "monto"
        CAMPOS(5, 0) = "cajero"
        CAMPOS(6, 0) = "montocuotas"
        CAMPOS(7, 0) = "interesesmora"
        CAMPOS(8, 0) = "caja"
        CAMPOS(9, 0) = "tipopago"
        CAMPOS(10, 0) = ""
        
        CAMPOS(0, 1) = empresaActiva
        FOLIO.Caption = leerUltimoFolioPAGO()
        CAMPOS(1, 1) = FOLIO.Caption
        CAMPOS(2, 1) = rut2.text + lbldv.Caption
        CAMPOS(3, 1) = Format(fechasistema, "yyyy-mm-dd")
        CAMPOS(4, 1) = CDbl(Replace(CANCELADO.text, ",", "."))
        CAMPOS(5, 1) = codigocajero
        CAMPOS(6, 1) = CDbl(Replace(PAGOCUOTAS.text, ",", "."))
        If PAGOMORA.text = "" Then PAGOMORA.text = "0"
        CAMPOS(7, 1) = CDbl(Replace(PAGOMORA.text, ",", "."))
        CAMPOS(8, 1) = caja
        If REPACTACIONES = True Then
        CAMPOS(0, 1) = "07"
        CAMPOS(9, 1) = "1"
        Else
        CAMPOS(9, 1) = ""
        End If
        
        CAMPOS(0, 2) = "sv_cuotas_pago_cabeza"
        condicion = ""
        op = 2
        sql.response = CAMPOS
        Set sql.conexion = ventas
        Call sql.sqlventas(op, condicion)
       
     
        
End Sub

Sub grabadetalle()
        
        Dim CAMPOS(14, 3) As String
        Dim op As Integer
        Dim K As Integer
        Dim TIPO As String
        Dim NUMERO As String
        Dim cuota As String
        Dim abono As Double
        
        Set sql = New sqlventas.sqlventa
        CAMPOS(0, 0) = "local"
        CAMPOS(1, 0) = "numero"
        CAMPOS(2, 0) = "rut"
        CAMPOS(3, 0) = "fecha"
        CAMPOS(4, 0) = "localdocumento"
        CAMPOS(5, 0) = "tipodocumento"
        CAMPOS(6, 0) = "numerodocumento"
        CAMPOS(7, 0) = "numerocuota"
        CAMPOS(8, 0) = "monto"
        CAMPOS(9, 0) = "montoantespago"
        CAMPOS(10, 0) = "diasmora"
        CAMPOS(11, 0) = "interesmora"
        CAMPOS(12, 0) = "totalcuotas"
        CAMPOS(13, 0) = "vencimiento"
        CAMPOS(14, 0) = ""
        
        For K = 1 To Grid1.Rows - 1
        If Grid1.Cell(K, 10).text <> "0" Then
        If REPACTACIONES = False Then
        CAMPOS(0, 1) = empresaActiva
        Else
        CAMPOS(0, 1) = "07"
        End If
     
        CAMPOS(1, 1) = FOLIO.Caption
        CAMPOS(2, 1) = rut2.text + lbldv.Caption
        CAMPOS(3, 1) = Format(fechasistema, "yyyy-mm-dd")
        CAMPOS(4, 1) = Grid1.Cell(K, 0).text
        
        TIPO = Mid(Grid1.Cell(K, 1).text, 1, 2)
        cuota = Mid(Grid1.Cell(K, 2).text, 1, 2)
        cuota = Replace(cuota, "/", "")
        NUMERO = Mid(Grid1.Cell(K, 1).text, 4, 10)
        CAMPOS(5, 1) = TIPO
        CAMPOS(6, 1) = NUMERO
        CAMPOS(7, 1) = cuota
                  
        CAMPOS(8, 1) = CDbl(Replace(Grid1.Cell(K, 10).text, ".", ""))
        If Grid1.Cell(K, 5).text = "" Then Grid1.Cell(K, 5).text = "0"
        CAMPOS(9, 1) = CDbl(Replace(Grid1.Cell(K, 5).text, ".", ""))
        CAMPOS(10, 1) = CDbl(Replace(Grid1.Cell(K, 7).text, ".", ""))
        CAMPOS(11, 1) = CDbl(Replace(Grid1.Cell(K, 8).text, ".", ""))
        CAMPOS(12, 1) = Grid1.Cell(K, 2).text
        CAMPOS(13, 1) = Format(Grid1.Cell(K, 3).text, "yyyy-mm-dd")
        CAMPOS(0, 2) = "sv_cuotas_pago_detalle"
        condicion = ""
        op = 2
        sql.response = CAMPOS
        Set sql.conexion = ventas
        Call sql.sqlventas(op, condicion)
        End If
        Next K
     
        
End Sub

Private Sub Command3_Click()
frmcancelado.Visible = False
Call RETORNAR_Click

End Sub

Private Sub IMPRESORA_DragDrop(Source As Control, x As Single, Y As Single)

End Sub


Private Sub Command4_Click()
FRMCUOTAS.Visible = False

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
CANCELADO.SetFocus

End If

End Sub

Private Sub CONDO1_LostFocus()
If CONDO1.text = "" Then CONDO1.text = "0"
If CDbl(CONDO1.text) > 100 Then
CONDO1.text = "0"
CONDO1.SetFocus

End If
Call LEErcuotas(rut2.text + lbldv.Caption)

End Sub
Private Sub CONDO2_LostFocus()
If CONDO2.text = "" Then CONDO2.text = "0"
If CDbl(CONDO2.text) > 100 Then
CONDO2.text = "0"
CONDO2.SetFocus

End If
Call LEErcuotas(rut2.text + lbldv.Caption)

End Sub


Private Sub CONDO2_KeyPress(KeyAscii As Integer)
KeyAscii = esNumero(KeyAscii)
If KeyAscii = 13 Then
CANCELAR.SetFocus

End If

End Sub





Private Sub Check1_Click()

If Check1.Value = 0 Then
vueltototal.text = ""
efectivo.text = ""
Vuelto.Visible = False

End If

If Check1.Value = 1 Then
Vuelto.Visible = True
efectivo.SetFocus
End If


End Sub

Private Sub CUOTAS_GotFocus()

Call cargatexto(CUOTAS)
End Sub

Private Sub CUOTAS_KeyPress(KeyAscii As Integer)
KeyAscii = esNumero(KeyAscii)
If CUOTAS.text <> "" Then
If KeyAscii = 13 And CUOTAS.text <> "" And CUOTAS.text <> "0" And CDbl(CUOTAS.text) <= leefactorMAXIMO Then
CALCULAR_Click
ACEPTAR.SetFocus

End If


End If

End Sub

Private Sub efectivo_GotFocus()
Call selecciona(efectivo)
End Sub

Private Sub efectivo_KeyDown(KeyCode As Integer, Shift As Integer)
  Call Flechas(KeyCode, efectivo)
End Sub

Private Sub efectivo_KeyPress(KeyAscii As Integer)
            KeyAscii = esNumero(KeyAscii)
           If KeyAscii = 13 And efectivo.text <> "" Then
           If CDbl(CANCELADO.text) <= CDbl(efectivo) Then
            vueltototal.text = CDbl(efectivo.text) - CDbl(CANCELADO.text)
            Else
            vueltototal.text = ""
            
            End If
            
            End If
End Sub



Private Sub Grid1_DblClick()

If CONDO1.text = "0" Then
If frmcancelado.Visible = True Then
If Grid1.ActiveCell.col = 9 Then
Grid1.Cell(Grid1.ActiveCell.row, 10).text = Grid1.Cell(Grid1.ActiveCell.row, 6).text
End If
If Grid1.ActiveCell.col = 10 Then
Grid1.Cell(Grid1.ActiveCell.row, 10).text = "0"
End If
If Grid1.ActiveCell.col = 1 Then
Call verboleta(Grid1.Cell(Grid1.ActiveCell.row, 1).text)
End If
Else
Call verboleta(Grid1.Cell(Grid1.ActiveCell.row, 1).text)
End If


SUMAPAGOS
End If
End Sub
Sub SUMAPAGOS()
Dim suma As Double
Dim PCUOTAS As Double
Dim PMORA As Double
suma = 0
For K = 1 To Grid1.Rows - 1
suma = suma + CDbl(Grid1.Cell(K, 10).text)
If Grid1.Cell(K, 10).text <> "0" Then
  
  If CDbl(Grid1.Cell(K, 10).text) > CDbl(Grid1.Cell(K, 9).text) Then
  Grid1.Cell(K, 10).text = Grid1.Cell(K, 9).text
  End If
  
  If CDbl(Grid1.Cell(K, 6).text) <= Grid1.Cell(K, 10).text Then
  PCUOTAS = PCUOTAS + CDbl(Grid1.Cell(K, 6).text)
  End If
  
End If
Next K
CANCELADO.text = Format(suma + CDbl(total2.Caption), "###,###,###")

PAGOCUOTAS.text = Format(PCUOTAS, "###,###,###")
PAGOMORA.text = Format(total2.Caption, "###,###,###")
End Sub

Private Sub Grid1_EnterRow(ByVal row As Long)
SUMAPAGOS

End Sub

Private Sub GRID1_LeaveCell(ByVal row As Long, ByVal col As Long, NewRow As Long, NewCol As Long, Cancel As Boolean)
SUMAPAGOS

End Sub

Private Sub Grid2_DblClick()
FOLIO.Caption = Grid2.Cell(Grid2.ActiveCell.row, 1).text

creditotmppago.Show

End Sub



Private Sub PIE_KeyPress(KeyAscii As Integer)
KeyAscii = esNumero(KeyAscii)

If KeyAscii = 13 Then
        If PIE.text = "" Then
        PIE.text = "0"
        End If
If CDbl(PIE.text) < CDbl(MONTO.text) Then
PIE.text = Format(PIE.text, " ###,###,##0")
CREDITO.text = Format(CDbl(MONTO.text) - CDbl(PIE.text), " #,###,##0")

CUOTAS.SetFocus
Else
MsgBox ("MONTO DEL PIE DEBE SER INFERIOR AL MONTO DE LA COMPRA")
PIE.text = "0"
PIE.SetFocus

End If

End If

End Sub

Private Sub repactar_Click()

If total4.Caption <> "0" Then
REPACTACIONES = True

frmcancelado.Visible = True
CANCELADO.Locked = True
CANCELADO.text = total4.Caption
MONTO.text = CANCELADO.text
CANCELAR.Enabled = False
Grid1.Column(10).Locked = True
CALCULAPRIMERVENCIMIENTO
    seleccionacuotas
    FRMCUOTAS.Visible = True
PIE.text = "0"
CREDITO.text = MONTO.text

CALCULAR_Click
ACEPTAR.SetFocus



End If
End Sub

Private Sub RETORNAR_Click()
lblCupo.Caption = "0"
lblDisponible.Caption = "0"
lblUtilizado.Caption = "0"
frmcancelado.Visible = False
total1.Caption = "0"
total2.Caption = "0"
total3.Caption = "0"
total4.Caption = "0"
Grid1.Column(10).Locked = True
CANCELADO.text = "0"
SELECCIONADO.text = "0"
diferencia.text = "0"
PAGOCUOTAS.text = "0"
PAGOMORA.text = "0"
rut2.SetFocus
Grid1.Rows = 1
Grid2.Rows = 1
lblnombre.Caption = ""
CONDO1.text = "0"
CONDO2.text = "0"
repactar.Visible = False

End Sub

Private Sub Form_Load()
CARGAGRILLA
CARGAGRILLA2
CARGAGRILLA3

DATOSCREDITO.Enabled = True
frmcancelado.Visible = False
CONDO1.text = "0"
CONDO2.text = "0"
DATOSCREDITO.Caption = "CAJERO :" + nombrecajero
repactar.Visible = False
FRMCUOTAS.Visible = False

End Sub
Private Sub CALCULAR_Click()
If CUOTAS.text <> "" Then
      CALCULACUOTA
   
End If
End Sub

Sub CALCULACUOTA()
Dim FACTOR As Double
Dim cuota As Double
Dim interes As Double

If CUOTAS.text <> "" Then

FACTOR = 0
If lbltc.Caption = "03" And CUOTAS.text < "5" Then FACTOR = 0
If lbltc.Caption = "04" And CUOTAS.text = "1" Then FACTOR = 0
If lbltc.Caption = "07" And CUOTAS.text < "4" Then FACTOR = 0


cuota = Int((CDbl(Replace(CREDITO.text, ",", "")) / CDbl(CUOTAS.text)) + 0.5)

interes = 0
For K = 1 To CDbl(CUOTAS.text)
interes = interes + (cuota * (K * FACTOR))
Next K

End If
cuota = Round((CDbl(CREDITO.text) + interes) / CDbl(CUOTAS), 0)



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

Sub CARGAGRILLA()
    Grid1.Cols = 13
    Grid1.DefaultFont.Size = 10
    Grid1.DefaultFont.Bold = True
    
    
    
    Grid1.Column(0).Width = 20
    Grid1.Column(1).Width = 120
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
    Grid1.Column(12).Width = 60
   
    
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
    
    Grid1.Cell(0, 0).text = "LO"
    Grid1.Cell(0, 1).text = "DOCUMENTO"
    Grid1.Cell(0, 2).text = "N.CUOTA"
    Grid1.Cell(0, 3).text = "VENCIMIENTO"
    Grid1.Cell(0, 4).text = "CUOTA"
    Grid1.Cell(0, 5).text = "ABONO"
    Grid1.Cell(0, 6).text = "SALDO"
    Grid1.Cell(0, 7).text = "DIAS MORA"
    Grid1.Cell(0, 8).text = "INTERES"
    Grid1.Cell(0, 9).text = "TOTAL"
    Grid1.Cell(0, 10).text = "CANCELANDO"
    
    Grid1.Column(10).Mask = cellNumeric
    
    
    
    Grid1.Cell(0, 11).text = "DETALLE COMPRAS "
    Grid1.Cell(0, 12).text = "CAPITAL"
    
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
    Grid2.Cols = 7
    Grid2.Column(0).Width = 20
    Grid2.Column(1).Width = 80
    Grid2.Column(2).Width = 80
    Grid2.Column(3).Width = 80
    Grid2.Column(4).Width = 80
    Grid2.Column(5).Width = 80
    Grid2.Column(6).Width = 150
    
    Grid2.Column(0).Locked = True
    Grid2.Column(1).Locked = True
    Grid2.Column(2).Locked = True
    Grid2.Column(3).Locked = True
    Grid2.Column(4).Locked = True
    Grid2.Column(5).Locked = True
    Grid2.Column(6).Locked = True
    
    Grid2.Cell(0, 0).text = "TP"
    Grid2.Cell(0, 1).text = "NUMERO"
    Grid2.Cell(0, 2).text = "FECHA"
    Grid2.Cell(0, 3).text = "CUOTAS"
    Grid2.Cell(0, 4).text = "INT.MORA"
    Grid2.Cell(0, 5).text = "TOTAL "
    Grid2.Cell(0, 6).text = "CAJERA"

    Grid2.Column(1).Alignment = cellRightTop
    Grid2.Column(2).Alignment = cellRightTop
    Grid2.Column(3).Alignment = cellRightTop
    Grid2.Column(4).Alignment = cellRightTop
    Grid2.Column(5).Alignment = cellRightTop
    Grid2.Column(6).Alignment = cellRightTop
    
    Grid2.Rows = 1
    
    

End Sub
Sub CARGAGRILLA3()
    Grid3.Cols = 3
    Grid3.Column(0).Width = 0
    Grid3.Column(1).Width = 100
    Grid3.Column(2).Width = 100
    
    Grid3.Column(0).Locked = True
    Grid3.Column(1).Locked = True
    Grid2.Column(2).Locked = True
    
    Grid3.Cell(0, 1).text = "FECHA"
    Grid3.Cell(0, 2).text = "MONTO"
    
    Grid3.Column(1).Alignment = cellRightTop
    Grid3.Column(2).Alignment = cellRightTop
    
    Grid3.Rows = 1
    Grid3.DefaultFont.Bold = True
    Grid3.DefaultFont.Size = 10
    Grid3.ExtendLastCol = True
    
    
    
    

End Sub



Sub CALCULAPRIMERVENCIMIENTO()
Dim mes As Double
Dim año As Double
Dim dia As Double

Dim sumar As Double
dia = Format(fechasistema, "dd")
mes = Format(fechasistema, "mm")
año = Format(fechasistema, "yyyy")
sumar = 0
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
If mes = 2 And dia > 28 Then
dia = 28
DIAC.text = "28"
End If


End Sub

Private Sub imprimir_Click()
'IMPRIMEcredito



End Sub

Private Sub MESC_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
AÑOC.SetFocus

End If

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






Public Function LEERDOCUMENTO(TIPO, NUMERO) As Boolean

        Dim csql As rdoQuery
        Dim resultado As rdoResultset
        Dim i As Integer
        Set csql = New rdoQuery
        Set csql.ActiveConnection = ventasRubro
        
        csql.sql = "SELECT dc.rut,dc.total,mc.diapago,mc.cupodirecto,mc.cupoutilizadotmp,mc.nombre,mc.direccion "
        csql.sql = csql.sql & "FROM sv_documento_cabeza as dc," + baseVentas + ".sv_maestroclientes as mc "
        csql.sql = csql.sql & "WHERE tipo = '" & TIPO & "' and numero='" & NUMERO & "' and dc.rut=mc.rut and local='00' "
        csql.Execute
        
        If csql.RowsAffected > 0 Then
            
            Set resultado = csql.OpenResultset
            While Not resultado.EOF
        MONTODOCUMENTO.Caption = Format(resultado(1), "###,###,##0")
        DIAPAGO.text = resultado(2)
        lblCupo.Caption = Format(resultado(3) * (1 + toleranciacredito / 100), "###,###,##0")
        lblUtilizado.Caption = Format(resultado(4), "###,###,##0")
        lblDisponible.Caption = Format(resultado(3) * (1 + toleranciacredito / 100) - resultado(4), "###,###,##0")
        MONTO.text = resultado(1)
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
Sub LEErcuotas(rut)

        Dim csql As rdoQuery
        Dim resultado As rdoResultset
        Dim i As Integer
        Dim diasmora As Double
        Dim saldo As Double
        Dim interes As Double
        Dim t1 As Double
        Dim t2 As Double
        Dim total As Double
        Dim porcecondo1 As Double
        Dim porcecondo2 As Double
        Dim cuota As Double
        Dim interescuota As Double
        Dim capital As Double
        Dim cuotabase As Double
        
        Set csql = New rdoQuery
        Set csql.ActiveConnection = ventas

        csql.sql = "SELECT *  "
        csql.sql = csql.sql & "FROM sv_cuotas_detalle "
        csql.sql = csql.sql & "WHERE rut='" + rut + "'  and ( (montocuota-abono)>0 ) "
        csql.sql = csql.sql & "order by vencimientoactual,abono,local,tipo,numero asc "
        csql.Execute
        totalusado = 0
        moratotal = 0
        If csql.RowsAffected > 0 Then

            Set resultado = csql.OpenResultset
            
        Grid1.Rows = 1
        Grid1.AutoRedraw = False
        
        totalusado = 0
        moratotal = 0
        While Not resultado.EOF
        Grid1.Rows = Grid1.Rows + 1
        Grid1.Cell(Grid1.Rows - 1, 0).text = resultado(0)
        Grid1.Cell(Grid1.Rows - 1, 1).text = resultado(1) & " " & resultado(2)
        Grid1.Cell(Grid1.Rows - 1, 2).text = resultado(4) & " / " & resultado(12)
        Grid1.Cell(Grid1.Rows - 1, 3).text = Format(resultado(6), "dd/mm/yyyy")
        cuotabase = resultado(7)
        cuota = resultado(7)
    
        interescuota = resultado(7) - resultado("capitalcuota")
     
        
        capital = resultado("capitalcuota")
        If CDbl(CONDO1.text) > 0 Then
        porcecondo1 = 1 - (CDbl(CONDO1.text) / 100)
        interescuota = Round(interescuota * porcecondo1)
        cuota = capital + interescuota
        Else
        cuota = cuotabase
        End If
        
        
        Grid1.Cell(Grid1.Rows - 1, 4).text = Format(cuota, "###,###,###")
        Grid1.Cell(Grid1.Rows - 1, 5).text = Format(resultado(8), "###,###,###")
       
        saldo = (cuota) - resultado(8)
      
        
        tazainteresmora = leerInteresMora("00")
        porcecondo2 = 1 - (CDbl(CONDO2.text) / 100)
        tazainteresmora = tazainteresmora * porcecondo2
        
        diasmora = DateDiff("d", resultado(6), fechasistema)
      If diasmora > 0 Then
        Grid1.Range(Grid1.Rows - 1, 1, Grid1.Rows - 1, Grid1.Cols - 1).BackColor = &HFF&
      End If
      
      
        If diasmora <= diasgracia Then diasmora = 0
      
        
        
        interes = Round(saldo * ((tazainteresmora / 100 / 30) * diasmora), 0)
        
        total = saldo + interes
        If saldo = 0 Then
        Grid1.Cell(Grid1.Rows - 1, 6).text = "0"
        Else
         Grid1.Cell(Grid1.Rows - 1, 6).text = Format(saldo, "###,###,###")
        End If
        Grid1.Cell(Grid1.Rows - 1, 7).text = diasmora
        Grid1.Cell(Grid1.Rows - 1, 8).text = interes
        If total = 0 Then
        Grid1.Cell(Grid1.Rows - 1, 9).text = "0"
        Else
        Grid1.Cell(Grid1.Rows - 1, 9).text = Format(saldo, "###,###,###")
        End If
        Grid1.Cell(Grid1.Rows - 1, 10).text = "0"
        Grid1.Cell(Grid1.Rows - 1, 11).text = resultado(13)
        Grid1.Cell(Grid1.Rows - 1, 12).text = resultado("capitalcuota")
        
        
        
        totalusado = totalusado + total
        If interes <> 0 Then moratotal = moratotal + total
        If Format(resultado(6), "yyyy-mm-dd") <= Format(fechasistema, "yyyy-mm-dd") Or CONDO1.text <> "0" Then
        Grid1.Cell(Grid1.Rows - 1, 10).text = Grid1.Cell(Grid1.Rows - 1, 9).text
        t1 = t1 + saldo
        t2 = t2 + interes
               
        End If
            
            resultado.MoveNext
            Wend
        Else
       
        End If
        
       
       
        Set resultado = Nothing
        csql.Close
        Set csql = Nothing
        Grid1.AutoRedraw = True
        Grid1.Refresh
        lblUtilizado.Caption = Format(totalusado, "###,###,##0")
        
        lblDisponible.Caption = Format(CDbl(lblCupo.Caption) - totalusado, "###,###,##0")
        total4.Caption = Format(totalusado, "###,###,##0")
        total1.Caption = Format(t1, "###,###,##0")
        total2.Caption = Format(t2, "###,###,##0")
        total3.Caption = Format(t1 + t2, "###,###,##0")
    
    SUMAPAGOS
    Call LEErcuotasACUMULADAS(rut)
    End Sub
Sub LEErcuotasACUMULADAS(rut)

        Dim csql As rdoQuery
        Dim resultado As rdoResultset
        
        Set csql = New rdoQuery
        Set csql.ActiveConnection = ventas

        csql.sql = "SELECT vencimientoactual,sum(montocuota-abono) "
        csql.sql = csql.sql & "FROM sv_cuotas_detalle "
        csql.sql = csql.sql & "WHERE rut='" + rut + "'  and ( (montocuota-abono)>0 or ((interesmora+montocuota)-abono)>0) "
        csql.sql = csql.sql & "group by vencimientoactual order by vencimientoactual "
        csql.Execute
        If csql.RowsAffected > 0 Then

            Set resultado = csql.OpenResultset
            
        Grid3.Rows = 1
        Grid3.AutoRedraw = False
        
        While Not resultado.EOF
        If resultado(1) <> 0 Then
        Grid3.Rows = Grid3.Rows + 1
        Grid3.Cell(Grid3.Rows - 1, 1).text = Format(resultado(0), "dd-mm-yyyy")
        Grid3.Cell(Grid3.Rows - 1, 2).text = Format(resultado(1), "###,###,###")
        End If
        resultado.MoveNext
        Wend
        Else
       
        End If
        
       
       
        Set resultado = Nothing
        csql.Close
        Set csql = Nothing
        Grid3.AutoRedraw = True
        Grid3.Refresh
    End Sub


Sub LEErHISTORICOPAGOS(rut)

        Dim csql As rdoQuery
        Dim resultado As rdoResultset
        Dim i As Integer
        Dim diasmora As Double
        Dim saldo As Double
        Dim interes As Double
        Dim t1 As Double
        Dim t2 As Double
        Dim total As Double
        Dim PAGO As String
        Dim cajera As String
        Dim dv As String
        
        Set csql = New rdoQuery
        Set csql.ActiveConnection = ventas

        csql.sql = "SELECT *  "
        csql.sql = csql.sql & "FROM sv_cuotas_pago_cabeza "
        csql.sql = csql.sql & "WHERE rut='" + rut + "'   "
        csql.sql = csql.sql & "order by fecha desc "
        csql.Execute
        
        If csql.RowsAffected > 0 Then

            Set resultado = csql.OpenResultset
            
        Grid2.Rows = 1
       Grid2.AutoRedraw = False
        
        While Not resultado.EOF
        Grid2.Rows = Grid2.Rows + 1
        If resultado(9) = "1" Then
        PAGO = "CR"
        Else
        PAGO = "CP"
        End If
        
        Grid2.Cell(Grid2.Rows - 1, 0).text = PAGO
        
        Grid2.Cell(Grid2.Rows - 1, 1).text = resultado(1)
        Grid2.Cell(Grid2.Rows - 1, 2).text = Format(resultado(3), "dd/mm/yyyy")
        Grid2.Cell(Grid2.Rows - 1, 3).text = Format(resultado(6), "###,###,###")
        Grid2.Cell(Grid2.Rows - 1, 4).text = Format(resultado(7), "###,###,###")
        Grid2.Cell(Grid2.Rows - 1, 5).text = Format(resultado(4), "###,###,###")
         
        
        Grid2.Cell(Grid2.Rows - 1, 6).text = cajera
            
            resultado.MoveNext
            Wend
        Else
       
        End If
        Set resultado = Nothing
        csql.Close
        Set csql = Nothing
        Grid2.AutoRedraw = True
        Grid2.Refresh
        
        
        
    End Sub

Public Function LEERCLIENTE(rut) As Boolean

        Dim csql As rdoQuery
        Dim resultado As rdoResultset
        Dim i As Integer
        Set csql = New rdoQuery
        Set csql.ActiveConnection = ventas
        
        csql.sql = "SELECT mc.diapago,mc.cupodirecto,mc.cupoutilizadodirecto,mc.nombre,mc.direccion,mc.repactacion,mc.rebajainterespp,rebajainteresmora,mc.tipocliente,mc.cuotasrepactacion "
        csql.sql = csql.sql & "FROM sv_maestroclientes as mc "
        csql.sql = csql.sql & "WHERE mc.rut='" + rut + "' "
        csql.Execute
        If csql.RowsAffected > 0 Then
            Set resultado = csql.OpenResultset
            While Not resultado.EOF
        
        lblCupo.Caption = Format(resultado(1) * (1 + toleranciacredito / 100), "###,###,##0")
        DIAPAGO.text = resultado(0)
        lblnombre.Caption = resultado(3)
        CONDO1.text = resultado(6)
        CONDO2.text = resultado(7)
        lbltipocliente.Caption = leertipocli(resultado(8))
                    lbltc.Caption = resultado(8)
        If resultado(5) = "1" Then
        repactar.Visible = True
        Else
        repactar.Visible = False
        End If
        CUOTAS.text = resultado(9)
        
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


Private Sub rut2_GotFocus()
Grid1.Rows = 1
FOLIO.Caption = leerUltimoFolioPAGO()

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
             Call LEErHISTORICOPAGOS(rut2.text + lbldv.Caption)
             
             
             Call LEErcuotas(rut2.text + lbldv.Caption)
            
             If Grid1.Rows > 1 Then CANCELAR.Enabled = True
             
             
             Else
             MsgBox ("CLIENTE NO CORRESPONDE A CLIENTE A CREDITO O NO TIENE CUPO ASIGNADO")
             rut2.SetFocus
             
             End If
             
            
        End If
End Sub

Public Sub IMPRIMEcredito()
    Dim K As Integer
    
    ''''''''''''''''''
    numfic = 20
    
    
    Open "LPT1" For Output As #numfic
    ''''''''''''''''''

    '''''''''''''''''''''
    'EMPAQUE
    '''''''''''''''''''''
    For K = 1 To 1
    Print #numfic, Chr$(27); Chr$(64) '
    Print #numfic, "ALMACENES ELTIT   "
    Print #numfic, ""
    Print #numfic, "          COMPROBANTE DE CREDITO         "
    Print #numfic, "          ======================         "
    Print #numfic,
    Print #numfic, "CLIENTE :"; rut2.text + "-" + lbldv.Caption
    Print #numfic, "NOMBRE  :"; lblnombre.Caption
    Print #numfic, "DIREC.  :"; LBLDIRECCION.Caption
    Print #numfic,
    Print #numfic, "SEGUN CONTRATO MAAT AUTORIZO  CARGAR A MI "
    Print #numfic, "CUENTA "
    
    Print #numfic, CUOTAS.text + " CUOTAS de " + VALORCUOTA.text + " Pesos"
    Print #numfic, "A partir  "; DIAC + "-" + MESC + "-" + AÑOC
    Print #numfic,
    Print #numfic,
    Print #numfic,
    Print #numfic, "              _______________             "
    Print #numfic, "               FIRMA CLIENTE             "
    Print #numfic,
    Print #numfic,
    Print #numfic,
    Print #numfic,
    Print #numfic,
    Print #numfic,
    
    Print #numfic, Chr(27); "i"
    Next K
    Close #numfic



End Sub
Sub verboleta(CODIGO)
Dim boleta As String
Dim TIPO As String

boleta = Mid(CODIGO, 4, 10)
TIPO = Mid(CODIGO, 1, 2)

'If TIPO <> "VC" Then
'Load PVentas
'With PVentas
'.Show
'.dato1.text = TIPO
'.cargardeafuera2
'.dato2.text = boleta
'.cargardeafuera
'
'End With
'
'End If

End Sub

Private Sub tmrBlink_Timer()
fechareal.Caption = Format(fechasistema, "dd-mm-yyyy") & " " & Time
End Sub
Sub cargagrilla4()
    Grid4.Cols = 11
    
    Grid4.Column(0).Width = 0
    Grid4.Column(1).Width = 0
    Grid4.Column(2).Width = 80
    Grid4.Column(3).Width = 80
    Grid4.Column(4).Width = 80
    Grid4.Column(5).Width = 80
    Grid4.Column(6).Width = 80
    Grid4.Column(7).Width = 80
    Grid4.Column(8).Width = 80
    Grid4.Column(9).Width = 80
    Grid4.Column(10).Width = 80
    
    
    Grid4.Column(0).Locked = True
    Grid4.Column(1).Locked = True
    Grid4.Column(2).Locked = True
    Grid4.Column(3).Locked = True
    Grid4.Column(4).Locked = True
    Grid4.Column(5).Locked = True
    Grid4.Column(6).Locked = True
    Grid4.Column(7).Locked = True
    Grid4.Column(8).Locked = True
    Grid4.Column(9).Locked = True
    Grid4.Column(10).Locked = True
    
    Grid4.Cell(0, 1).text = "TD"
    Grid4.Cell(0, 2).text = "NUMERO"
    Grid4.Cell(0, 3).text = "FECHA"
    Grid4.Cell(0, 4).text = "CUOTA"
    Grid4.Cell(0, 5).text = "ABONO"
    Grid4.Cell(0, 6).text = "SALDO"
    Grid4.Cell(0, 7).text = "MOROSIDAD"
    Grid4.Cell(0, 8).text = "INTERES"
    Grid4.Cell(0, 9).text = "TOTAL"
    
    Grid4.Column(4).Alignment = cellRightTop
    Grid4.Column(5).Alignment = cellRightTop
    Grid4.Column(6).Alignment = cellRightTop
    Grid4.Column(7).Alignment = cellRightTop
    Grid4.Column(8).Alignment = cellRightTop
    Grid4.Column(9).Alignment = cellRightTop
    
    
    
    
    
    Me.Left = (Screen.Width - Me.Width) / 2
    Me.Top = (Screen.Height - Me.Height) / 2 - 950
Grid4.Rows = 1
    
   
End Sub

Sub CALCULATODASCUOTAS()
Dim K As Integer
Dim cuota As Double
Dim fin As Double
Dim fecha As String
Dim mesinicio As Double
Dim añoinicio As Double
Dim s As Integer
cargagrilla4

fecha = DIAC.text + "/" + MESC.text + "/" + AÑOC.text
fin = leefactorMAXIMO
mesinicio = MESC.text
añoinicio = AÑOC.text
Grid4.Rows = fin + 1
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

Grid4.Cell(K, 1).text = K

Grid4.Cell(K, 2).text = Format(cuota, "###,###,##0")
Grid4.Cell(K, 3).text = Format(cuota * K, "###,###,##0")
Grid4.Cell(K, 4).text = Format(fecha, "dd-mm-yyyy")
mesinicio = mesinicio + 1
If mesinicio > 12 Then
mesinicio = 1
añoinicio = añoinicio + 1

End If

pivote.MaxLength = 2
pivote.text = mesinicio
pivote.text = ceros(pivote)

'If DIAPAGO.text <> DIAC.text Then
'DIAC.text = DIAPAGO.text
'
'End If


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

Public Sub grabarcuotas(NUMERO)
        
        Dim CAMPOS(20, 3) As String
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
        CAMPOS(11, 0) = "fechacompra"
        CAMPOS(12, 0) = "glosacompra"
        CAMPOS(13, 0) = "rutadicional"
        CAMPOS(14, 0) = "nombreadicional"
        CAMPOS(15, 0) = "caja"
        CAMPOS(16, 0) = "cajera"
        CAMPOS(17, 0) = "pie"
        CAMPOS(18, 0) = "montoventa"
        CAMPOS(19, 0) = ""
 
        CAMPOS(0, 1) = "07"
 
        
        CAMPOS(1, 1) = "VR"
        CAMPOS(2, 1) = NUMERO
        CAMPOS(3, 1) = rut2.text + lbldv.Caption
        For K = 1 To CDbl(CUOTAS.text)
        CAMPOS(4, 1) = K
        CAMPOS(5, 1) = Format(Grid4.Cell(K, 4).text, "yyyy-mm-dd")
        CAMPOS(6, 1) = Format(Grid4.Cell(K, 4).text, "yyyy-mm-dd")
        CAMPOS(7, 1) = Replace(VALORCUOTA.text, ".", "")
        CAMPOS(8, 1) = CUOTAS.text
        CAMPOS(9, 1) = CDbl(CREDITO.text) / CDbl(CUOTAS.text)
        CAMPOS(10, 1) = Replace(CDbl(CREDITO.text), ".", "")
        CAMPOS(11, 1) = Format(fechasistema, "yyyy-mm-dd")
        CAMPOS(12, 1) = "REPACTACION " + nombreempresa
        
        
        
        
        CAMPOS(13, 1) = rutadicional
        CAMPOS(14, 1) = adicional
        CAMPOS(15, 1) = ""
        CAMPOS(16, 1) = cajera
        CAMPOS(17, 1) = Replace(CDbl(PIE.text), ".", "")
        CAMPOS(18, 1) = Replace(CDbl(MONTO.text), ".", "")
        
        
        CAMPOS(0, 2) = "sv_cuotas_detalle"
        condicion = ""
        op = 2
        sql.response = CAMPOS
        Set sql.conexion = ventas
        sql.audit = True: sql.programaactivo = Me.Caption
        Set sql.conauditoria = conauditoria: sql.usuarioauditoria = usuarioSistema
        Call sql.sqlventas(op, condicion)
        Next K
    
    End Sub
