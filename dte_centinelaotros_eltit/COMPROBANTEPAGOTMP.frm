VERSION 5.00
Object = "{4777436C-EB8C-4596-98A8-EBCF98683969}#1.0#0"; "FlexCell.ocx"
Object = "{ADBBAED6-B16F-40DC-80DF-B44910CBA76C}#1.0#0"; "Frame-Xp.ocx"
Begin VB.Form creditoPAGOSTMP 
   BackColor       =   &H00AE1118&
   BorderStyle     =   0  'None
   Caption         =   "Crédito"
   ClientHeight    =   9705
   ClientLeft      =   210
   ClientTop       =   1710
   ClientWidth     =   14640
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   9705
   ScaleWidth      =   14640
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
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
      Left            =   495
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   9045
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
      Top             =   90
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
         Left            =   6210
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
            Left            =   1440
            TabIndex        =   26
            Top             =   1800
            Width           =   2715
         End
         Begin VB.TextBox PAGOMORA 
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
            Left            =   4275
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
            Left            =   4275
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
            Left            =   5670
            TabIndex        =   22
            Top             =   900
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
            Left            =   2835
            TabIndex        =   20
            Top             =   900
            Width           =   2715
         End
         Begin VB.TextBox CANCELADO 
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
            TabIndex        =   18
            Top             =   900
            Width           =   2715
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
            Left            =   5670
            TabIndex        =   23
            Top             =   495
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
            Left            =   2835
            TabIndex        =   21
            Top             =   495
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
            Height          =   390
            Index           =   6
            Left            =   45
            TabIndex        =   19
            Top             =   495
            Width           =   2715
         End
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
         Top             =   450
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
            Left            =   405
            Style           =   1  'Graphical
            TabIndex        =   41
            Top             =   5355
            Width           =   2085
         End
         Begin FlexCell.Grid Grid1 
            Height          =   5055
            Left            =   0
            TabIndex        =   30
            Top             =   270
            Width           =   14145
            _ExtentX        =   24950
            _ExtentY        =   8916
            Cols            =   5
            DefaultFontSize =   8.25
            Rows            =   30
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
            TabIndex        =   38
            Top             =   5805
            Width           =   2130
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
            TabIndex        =   37
            Top             =   5445
            Width           =   2130
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
            Index           =   1
            Left            =   3060
            TabIndex        =   36
            Top             =   5445
            Width           =   2325
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
            TabIndex        =   35
            Top             =   5805
            Width           =   2325
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
            Index           =   2
            Left            =   5490
            TabIndex        =   34
            Top             =   5445
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
            TabIndex        =   33
            Top             =   5805
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
            TabIndex        =   32
            Top             =   5805
            Width           =   2130
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
            Index           =   11
            Left            =   7830
            TabIndex        =   31
            Top             =   5445
            Width           =   2130
         End
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
         Width           =   5970
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



Private Sub CANCELADO_KeyPress(KeyAscii As Integer)
KeyAscii = esNumero(KeyAscii)
    If KeyAscii = 13 And CANCELADO.text <> "" Then
        If CDbl(CANCELADO.text) <= CDbl(totaldeuda.Caption) And CDbl(totalmora.Caption) < CDbl(CANCELADO.text) Then
            seleccionacuotas
        Else
        If CDbl(CANCELADO.text) >= CDbl(totaldeuda.Caption) Then
            MsgBox ("MONTO CANCELADO SUPERA LA DEUDA TOTAL ")
            CANCELADO.SetFocus
        End If
        If CDbl(CANCELADO.text) < CDbl(totalmora.Caption) Then
            MsgBox ("MONTO CANCELADO INFERIOR A LOS INTERESES")
            CANCELADO.SetFocus
        End If
        End If
End If

End Sub
Sub seleccionacuotas()
Dim monto2 As Double
Dim saldo2 As Double
mcancelado = CDbl(CANCELADO.text)

monto2 = CDbl(CANCELADO.text)
saldo2 = monto2
For k = 1 To Grid1.Rows - 1
Grid1.Cell(k, 10).text = "0"

Next k

For k = 1 To Grid1.Rows - 1

    If saldo2 > 0 Then
        
            If saldo2 > CDbl(Grid1.Cell(k, 9).text) Then
                Grid1.Cell(k, 10).text = Grid1.Cell(k, 9).text
            Else
                Grid1.Cell(k, 10).text = saldo2
            End If
    saldo2 = saldo2 - CDbl(Grid1.Cell(k, 6).text)
    End If

Next k
SUMACANCELACION
End Sub
Sub SUMACANCELACION()
Dim PCUOTAS As Double
Dim PMORA As Double

mseleccionado = 0
PCUOTAS = 0
PMORA = 0
For k = 1 To Grid1.Rows - 1
mseleccionado = mseleccionado + CDbl(Grid1.Cell(k, 10).text)
    If CDbl(Grid1.Cell(k, 10).text) <> 0 Then
        If CDbl(Grid1.Cell(k, 10).text) <= CDbl(Grid1.Cell(k, 9).text) Then
        PCUOTAS = PCUOTAS + CDbl(Grid1.Cell(k, 10).text)
        Else
        PCUOTAS = PCUOTAS + CDbl(Grid1.Cell(k, 9).text)
        End If
        
        PMORA = PMORA + CDbl(Grid1.Cell(k, 8).text)
    End If
    
Next k
mdiferencia = mcancelado - mseleccionado + CDbl(totalmora)
SELECCIONADO.text = Format(mseleccionado, "###,###,###")
diferencia.text = Format(mdiferencia, "###,###,###")
PAGOCUOTAS.text = Format(PCUOTAS, "###,###,###")
PAGOMORA.text = Format(PMORA, "###,###,###")

End Sub

Private Sub CANCELAR_Click()
If totaldeuda.Caption <> "0" Then
frmcancelado.Visible = True
CANCELADO.SetFocus
cancelar.Enabled = False
Else
rut2.SetFocus

End If
End Sub

Private Sub Command1_Click()
Unload Me

End Sub




Private Sub Command2_Click()
GENERAPAGO
IMPRIME_PAGO

Call RETORNAR_Click

End Sub

Sub IMPRIMEPAGO()
IMPRESORA.Visible = True

End Sub
Sub GENERAPAGO()
  Dim condicion As String
        Dim campos(12, 3) As String
        Dim op As Integer
        Dim k As Integer
        Dim TIPO As String
        Dim NUMERO As String
        Dim cuota As String
        Dim abono As Double
        Set sql = New CSQLUtil
        campos(0, 0) = "abono"
        campos(1, 0) = ""
        For k = 1 To Grid1.Rows - 1
        If Grid1.Cell(k, 5).text = "" Then
         abono = CDbl(Replace(Grid1.Cell(k, 10).text, ".", "")) + 0
       Else
        abono = CDbl(Replace(Grid1.Cell(k, 10).text, ".", "")) + CDbl(Replace(Grid1.Cell(k, 5).text, ".", ""))
      
        End If
         campos(0, 1) = abono
        TIPO = Mid(Grid1.Cell(k, 1).text, 1, 2)
        NUMERO = Mid(Grid1.Cell(k, 1).text, 4, 10)
        cuota = Mid(Grid1.Cell(k, 2).text, 1, 2)
        cuota = Replace(cuota, "/", "")
        campos(0, 2) = "sv_cuotas_detalle"
        condicion = "tipo='" & TIPO & "' and numero='" & NUMERO & "' and numerocuota='" & cuota & "'"
        op = 3
        sql.datos = campos
        Set sql.conexion = ventas
        Call sql.SQLUTIL(op, condicion)
        Next k
      grabacabeza
        grabadetalle
        
End Sub
Sub grabacabeza()
        Dim condicion As String
        Dim campos(12, 3) As String
        Dim op As Integer
        Dim k As Integer
        Dim TIPO As String
        Dim NUMERO As String
        Dim cuota As String
        Dim abono As Double
        
        Set sql = New CSQLUtil
        campos(0, 0) = "local"
        campos(1, 0) = "numero"
        campos(2, 0) = "rut"
        campos(3, 0) = "fecha"
        campos(4, 0) = "monto"
        campos(5, 0) = "cajero"
        campos(6, 0) = "montocuotas"
        campos(7, 0) = "interesesmora"
        campos(8, 0) = ""
        
        campos(0, 1) = empresaActiva
        campos(1, 1) = folio.Caption
        campos(2, 1) = rut2.text + lbldv.Caption
        campos(3, 1) = Format(fechasistema, "yyyy-mm-dd")
        campos(4, 1) = CDbl(Replace(CANCELADO.text, ",", "."))
        campos(5, 1) = ""
        campos(6, 1) = CDbl(Replace(PAGOCUOTAS.text, ",", "."))
        If PAGOMORA.text = "" Then PAGOMORA.text = "0"
        campos(7, 1) = CDbl(Replace(PAGOMORA.text, ",", "."))
              
        campos(0, 2) = "sv_cuotas_pago_cabeza"
        condicion = ""
        op = 2
        sql.datos = campos
        Set sql.conexion = ventas
        Call sql.SQLUTIL(op, condicion)
       
     
        
End Sub

Sub grabadetalle()
        Dim condicion As String
        Dim campos(12, 3) As String
        Dim op As Integer
        Dim k As Integer
        Dim TIPO As String
        Dim NUMERO As String
        Dim cuota As String
        Dim abono As Double
        
        Set sql = New CSQLUtil
        campos(0, 0) = "local"
        campos(1, 0) = "numero"
        campos(2, 0) = "rut"
        campos(3, 0) = "fecha"
        campos(4, 0) = "localdocumento"
        campos(5, 0) = "tipodocumento"
        campos(6, 0) = "numerodocumento"
        campos(7, 0) = "numerocuota"
        campos(8, 0) = "monto"
        campos(9, 0) = "montoantespago"
        campos(10, 0) = "diasmora"
        campos(11, 0) = "interesmora"
        
        
        
        For k = 1 To Grid1.Rows - 1
        If Grid1.Cell(k, 10).text <> "0" Then
        
        campos(0, 1) = empresaActiva
        
        campos(1, 1) = folio.Caption
        campos(2, 1) = rut2.text + lbldv.Caption
        campos(3, 1) = Format(fechasistema, "yyyy-mm-dd")
        campos(4, 1) = Grid1.Cell(k, 0).text
        
        TIPO = Mid(Grid1.Cell(k, 1).text, 1, 2)
        cuota = Mid(Grid1.Cell(k, 2).text, 1, 2)
        cuota = Replace(cuota, "/", "")
        NUMERO = Mid(Grid1.Cell(k, 1).text, 4, 10)
        campos(5, 1) = TIPO
        campos(6, 1) = NUMERO
        campos(7, 1) = cuota
                  
        campos(8, 1) = CDbl(Replace(Grid1.Cell(k, 10).text, ".", ""))
        If Grid1.Cell(k, 5).text = "" Then Grid1.Cell(k, 5).text = "0"
        campos(9, 1) = CDbl(Replace(Grid1.Cell(k, 5).text, ".", ""))
        campos(10, 1) = CDbl(Replace(Grid1.Cell(k, 7).text, ".", ""))
        campos(11, 1) = CDbl(Replace(Grid1.Cell(k, 8).text, ".", ""))
              
        campos(0, 2) = "sv_cuotas_pago_detalle"
        condicion = ""
        op = 2
        sql.datos = campos
        Set sql.conexion = ventas
        Call sql.SQLUTIL(op, condicion)
        End If
        Next k
     
        
End Sub

Private Sub Command3_Click()
frmcancelado.Visible = False
Call RETORNAR_Click

End Sub

Private Sub IMPRESORA_DragDrop(Source As Control, X As Single, Y As Single)

End Sub

Private Sub RETORNAR_Click()
lblCupo.Caption = "0"
lblDisponible.Caption = "0"
lblUtilizado.Caption = "0"
frmcancelado.Visible = False
moroso.Caption = "0"
totalcuotas.Caption = "0"
totalmora.Caption = "0"
totaldeuda.Caption = "0"
CANCELADO.text = "0"
SELECCIONADO.text = "0"
diferencia.text = "0"
PAGOCUOTAS.text = "0"
PAGOMORA.text = "0"
rut2.SetFocus
Grid1.Rows = 1
Grid2.Rows = 1
lblNombre.Caption = ""
End Sub

Private Sub Form_Load()
CARGAGRILLA
CARGAGRILLA2
DATOSCREDITO.Enabled = True
frmcancelado.Visible = False


IMPRESORA.Visible = False


End Sub
Sub CARGAGRILLA()
    Grid1.Cols = 12
    
    Grid1.Column(0).Width = 20
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
    Grid2.Cols = 6
    Grid2.Column(0).Width = 0
    Grid2.Column(1).Width = 80
    Grid2.Column(2).Width = 80
    Grid2.Column(3).Width = 80
    Grid2.Column(4).Width = 80
    Grid2.Column(5).Width = 80
    
    Grid2.Column(0).Locked = True
    Grid2.Column(1).Locked = True
    Grid2.Column(2).Locked = True
    Grid2.Column(3).Locked = True
    Grid2.Column(4).Locked = True
    Grid2.Column(5).Locked = True
    
    Grid2.Cell(0, 1).text = "FECHA"
    Grid2.Cell(0, 2).text = "NUMERO"
    Grid2.Cell(0, 3).text = "CUOTAS"
    Grid2.Cell(0, 4).text = "INT.MORA"
    Grid2.Cell(0, 5).text = "TOTAL "

    Grid2.Column(1).Alignment = cellRightTop
    Grid2.Column(2).Alignment = cellRightTop
    Grid2.Column(3).Alignment = cellRightTop
    Grid2.Column(4).Alignment = cellRightTop
    Grid2.Column(5).Alignment = cellRightTop
    
    Grid2.Rows = 1
    
    

End Sub



Sub CALCULACUOTA()
Dim FACTOR As Double

Dim cuota As Double
FACTOR = leefactor(CUOTAS.text)
cuota = Int((CDbl(Replace(MONTO.text, ",", "")) * FACTOR / CDbl(CUOTAS.text)) + 0.5)

VALORCUOTA.text = Format(cuota, "###,###,###")

CALCULATODASCUOTAS

End Sub
Sub CALCULATODASCUOTAS()
Dim k As Integer
Dim cuota As Double
Dim fin As Double
Dim fecha As String
Dim mesinicio As Double
Dim añoinicio As Double

fecha = DIAC.text + "/" + MESC.text + "/" + AÑOC.text
fin = leefactorMAXIMO
mesinicio = MESC.text
añoinicio = AÑOC.text

For k = 1 To fin

FACTOR = leefactor(Str(k))
cuota = Int((CDbl(Replace(MONTO.text, ",", "")) * FACTOR / k) + 0.5)
Grid2.Cell(k, 1).text = k

Grid2.Cell(k, 2).text = Format(cuota, "###,###,##0")
Grid2.Cell(k, 3).text = Format(cuota * k, "###,###,##0")
Grid2.Cell(k, 4).text = Format(fecha, "dd-mm-yyyy")
mesinicio = mesinicio + 1
If mesinicio > 12 Then
mesinicio = 1
añoinicio = añoinicio + 1

End If
PIVOTE.MaxLength = 2
PIVOTE.text = mesinicio
PIVOTE.text = ceros(PIVOTE)
fecha = DIAC.text & "-" & PIVOTE.text & "-" & añoinicio


Next k

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
PIVOTE.text = mes
PIVOTE.MaxLength = 2
PIVOTE.text = ceros(PIVOTE)
MESC = PIVOTE.text
AÑOC = año




End Sub

Private Sub imprimir_Click()
'IMPRIMEcredito



End Sub

Private Sub MESC_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
AÑOC.SetFocus

End If

End Sub



Private Sub MONTO_KeyPress(KeyAscii As Integer)
KeyAscii = esNumero(KeyAscii)
If KeyAscii = 13 And CDbl(MONTO.text) <> 0 Then
    If CDbl(MONTO.text) < CDbl(MONTODOCUMENTO.Caption) Then
    MsgBox ("EL CREDITO QUE ESTA OTORGANDO ES MENOR A MONTO DE LA VENTA ")
    
    End If
    If CDbl(MONTO.text) > CDbl(MONTODOCUMENTO.Caption) Then
    MsgBox ("EL CREDITO QUE ESTA ORTONGANDO ES SUPERIOR A LA VENTA RECTIFIQUE ")
    MONTO.text = Str(CDbl(MONTODOCUMENTO.Caption))
    MONTO.SetFocus
    
    
    End If

CUOTAS.SetFocus


  
    
End If


End Sub


Public Function LEERDOCUMENTO(TIPO, NUMERO) As Boolean

        Dim cSql As rdoQuery
        Dim resultado As rdoResultset
        Dim i As Integer
        Set cSql = New rdoQuery
        Set cSql.ActiveConnection = ventasRubro
        
        cSql.sql = "SELECT dc.rut,dc.total,mc.diapago,mc.cupotmp,mc.cupoutilizadotmp,mc.nombre,mc.direccion "
        cSql.sql = cSql.sql & "FROM sv_documento_cabeza as dc," + baseVentas + ".sv_maestroclientes as mc "
        cSql.sql = cSql.sql & "WHERE tipo = '" & TIPO & "' and numero='" & NUMERO & "' and dc.rut=mc.rut and local='00' "
        cSql.Execute
        
        If cSql.RowsAffected > 0 Then
            
            Set resultado = cSql.OpenResultset
            While Not resultado.EOF
        MONTODOCUMENTO.Caption = Format(resultado(1), "###,###,##0")
        DIAPAGO.text = resultado(2)
        lblCupo.Caption = Format(resultado(3), "###,###,##0")
        lblUtilizado.Caption = Format(resultado(4), "###,###,##0")
        lblDisponible.Caption = Format(resultado(3) - resultado(4), "###,###,##0")
        MONTO.text = resultado(1)
        rut2.text = Mid(resultado(0), 1, 9)
        lbldv.Caption = Mid(resultado(0), 10, 1)
        lblNombre.Caption = resultado(5)
        lblDireccion.Caption = resultado(6)
        
        LEERDOCUMENTO = True
        
            resultado.MoveNext
            Wend
        Else
        LEERDOCUMENTO = False
        
        End If
        Set resultado = Nothing
        cSql.Close
        Set cSql = Nothing
    End Function
Sub LEErcuotas(rut)

        Dim cSql As rdoQuery
        Dim resultado As rdoResultset
        Dim i As Integer
        Dim diasmora As Double
        Dim saldo As Double
        Dim interes As Double
        Dim t1 As Double
        Dim t2 As Double
        Dim total As Double
        
        Set cSql = New rdoQuery
        Set cSql.ActiveConnection = ventas

        cSql.sql = "SELECT *  "
        cSql.sql = cSql.sql & "FROM sv_cuotas_detalle "
        cSql.sql = cSql.sql & "WHERE rut='" + rut + "'  and (montocuota-abono)>0 "
        cSql.sql = cSql.sql & "order by vencimientoactual "
        cSql.Execute
        
        If cSql.RowsAffected > 0 Then

            Set resultado = cSql.OpenResultset
            
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
        Grid1.Cell(Grid1.Rows - 1, 4).text = Format(resultado(7), "###,###,###")
        Grid1.Cell(Grid1.Rows - 1, 5).text = Format(resultado(8), "###,###,###")
        saldo = resultado(7) - resultado(8)
        tazainteresmora = leerInteresMora("00")
        diasmora = DateDiff("d", resultado(6), fechasistema)
        
        
        If diasmora <= diasgracia Then diasmora = 0
       
        
        interes = Round(saldo * ((tazainteresmora / 30 * diasmora) / 100), 0)
        
        t1 = t1 + saldo
        t2 = t2 + interes
        
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
        
        Grid1.Cell(Grid1.Rows - 1, 9).text = Format(total, "###,###,###")
        End If
        Grid1.Cell(Grid1.Rows - 1, 10).text = "0"
        
        Grid1.Cell(Grid1.Rows - 1, 11).text = resultado(13)
        
        totalusado = totalusado + total
        If interes <> 0 Then moratotal = moratotal + total
        
            
            resultado.MoveNext
            Wend
        Else
       
        End If
        Set resultado = Nothing
        cSql.Close
        Set cSql = Nothing
        Grid1.AutoRedraw = True
        Grid1.Refresh
        lblUtilizado.Caption = Format(totalusado, "###,###,##0")
        lblDisponible.Caption = Format(CDbl(lblCupo.Caption) - totalusado, "###,###,##0")
        totaldeuda.Caption = Format(totalusado, "###,###,##0")
        moroso.Caption = Format(moratotal, "###,###,##0")
        
        
        totalcuotas.Caption = Format(t1, "###,###,##0")
        totalmora.Caption = Format(t2, "###,###,##0")
        
    End Sub
Sub LEErHISTORICOPAGOS(rut)

        Dim cSql As rdoQuery
        Dim resultado As rdoResultset
        Dim i As Integer
        Dim diasmora As Double
        Dim saldo As Double
        Dim interes As Double
        Dim t1 As Double
        Dim t2 As Double
        Dim total As Double
        
        Set cSql = New rdoQuery
        Set cSql.ActiveConnection = ventas

        cSql.sql = "SELECT *  "
        cSql.sql = cSql.sql & "FROM sv_cuotas_pago_cabeza "
        cSql.sql = cSql.sql & "WHERE rut='" + rut + "'   "
        cSql.sql = cSql.sql & "order by fecha "
        cSql.Execute
        
        If cSql.RowsAffected > 0 Then

            Set resultado = cSql.OpenResultset
            
        Grid2.Rows = 1
       Grid2.AutoRedraw = False
        
        While Not resultado.EOF
        Grid2.Rows = Grid2.Rows + 1
        Grid2.Cell(Grid2.Rows - 1, 1).text = resultado(1)
        Grid2.Cell(Grid2.Rows - 1, 2).text = Format(resultado(3), "dd/mm/yyyy")
        Grid2.Cell(Grid2.Rows - 1, 3).text = Format(resultado(6), "###,###,###")
        Grid2.Cell(Grid2.Rows - 1, 4).text = Format(resultado(7), "###,###,###")
        Grid2.Cell(Grid2.Rows - 1, 5).text = Format(resultado(4), "###,###,###")
            resultado.MoveNext
            Wend
        Else
       
        End If
        Set resultado = Nothing
        cSql.Close
        Set cSql = Nothing
        Grid2.AutoRedraw = True
        Grid2.Refresh
        
        
        
    End Sub

Public Function LEERCLIENTE(rut) As Boolean

        Dim cSql As rdoQuery
        Dim resultado As rdoResultset
        Dim i As Integer
        Set cSql = New rdoQuery
        Set cSql.ActiveConnection = ventas
        
        cSql.sql = "SELECT mc.diapago,mc.cupodirecto,mc.cupoutilizadodirecto,mc.nombre,mc.direccion "
        cSql.sql = cSql.sql & "FROM sv_maestroclientes as mc "
        cSql.sql = cSql.sql & "WHERE mc.rut='" + rut + "' "
        cSql.Execute
        If cSql.RowsAffected > 0 Then
            Set resultado = cSql.OpenResultset
            While Not resultado.EOF
        
        
        lblCupo.Caption = Format(resultado(1), "###,###,##0")
        
        lblNombre.Caption = resultado(3)

        
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
        cSql.Close
        Set cSql = Nothing
    End Function


Private Sub rut2_GotFocus()
folio.Caption = leerUltimoFolioPAGO()

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
            
             If Grid1.Rows > 1 Then cancelar.Enabled = True
             
             
             Else
             MsgBox ("CLIENTE NO CORRESPONDE A CLIENTE A CREDITO O NO TIENE CUPO ASIGNADO")
             rut2.SetFocus
             
             End If
             
            
        End If
End Sub

Public Sub IMPRIMEcredito()
    Dim k As Integer
    
    ''''''''''''''''''
    numfic = 20
    
    
    Open "LPT1" For Output As #numfic
    ''''''''''''''''''

    '''''''''''''''''''''
    'EMPAQUE
    '''''''''''''''''''''
    For k = 1 To 1
    Print #numfic, Chr$(27); Chr$(64) '
    Print #numfic, "IMPORTADORA SKORPIOS   "
    Print #numfic, ""
    Print #numfic, "          COMPROBANTE DE CREDITO         "
    Print #numfic, "          ======================         "
    Print #numfic,
    Print #numfic, "CLIENTE :"; rut2.text + "-" + lbldv.Caption
    Print #numfic, "NOMBRE  :"; lblNombre.Caption
    Print #numfic, "DIREC.  :"; lblDireccion.Caption
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
    Next k
    Close #numfic



End Sub


