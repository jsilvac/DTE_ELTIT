VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash.ocx"
Object = "{4777436C-EB8C-4596-98A8-EBCF98683969}#1.0#0"; "FlexCell.ocx"
Object = "{ADBBAED6-B16F-40DC-80DF-B44910CBA76C}#1.0#0"; "Frame-Xp.ocx"
Begin VB.Form consumo04 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Maestro Consumos Basicos"
   ClientHeight    =   10425
   ClientLeft      =   2040
   ClientTop       =   1425
   ClientWidth     =   15240
   DrawStyle       =   5  'Transparent
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   695
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1016
   StartUpPosition =   2  'CenterScreen
   Begin XPFrame.FrameXp FrameQuickMenu 
      Height          =   615
      Left            =   11760
      TabIndex        =   72
      Top             =   4560
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
         TabIndex        =   74
         Top             =   280
         Width           =   1335
      End
      Begin VB.CommandButton botonmisaccesos 
         Caption         =   "Permisos Modulo"
         Height          =   255
         Left            =   120
         TabIndex        =   73
         Top             =   280
         Width           =   1455
      End
   End
   Begin VB.CommandButton cmdgrabar 
      BackColor       =   &H00FF8080&
      Caption         =   "GRABAR"
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
      Left            =   7080
      Style           =   1  'Graphical
      TabIndex        =   60
      Top             =   6120
      Visible         =   0   'False
      Width           =   1815
   End
   Begin XPFrame.FrameXp FrameXp2 
      Height          =   4380
      Left            =   9120
      TabIndex        =   18
      Top             =   120
      Width           =   6030
      _ExtentX        =   10636
      _ExtentY        =   7726
      BackColor       =   16761024
      Caption         =   "Detalle de Pagos"
      CaptionEstilo3D =   1
      BackColor       =   16761024
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin FlexCell.Grid Grid1 
         Height          =   3930
         Left            =   120
         TabIndex        =   19
         Top             =   270
         Width           =   5850
         _ExtentX        =   10319
         _ExtentY        =   6932
         Cols            =   5
         DefaultFontSize =   8.25
         Rows            =   30
      End
   End
   Begin VB.PictureBox MANUAL 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H80000002&
      ForeColor       =   &H80000008&
      Height          =   0
      Left            =   0
      ScaleHeight     =   0
      ScaleWidth      =   15210
      TabIndex        =   17
      Top             =   10425
      Width           =   15240
   End
   Begin XPFrame.FrameXp FrameXp3 
      Height          =   3810
      Left            =   90
      TabIndex        =   20
      Top             =   6600
      Width           =   14910
      _ExtentX        =   26300
      _ExtentY        =   6720
      BackColor       =   16761024
      Caption         =   "Servicios Vigentes"
      CaptionEstilo3D =   1
      BackColor       =   16761024
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
         Height          =   3465
         Left            =   90
         TabIndex        =   21
         Top             =   270
         Width           =   14775
         _ExtentX        =   26061
         _ExtentY        =   6112
         Cols            =   5
         DefaultFontSize =   8.25
         Rows            =   30
      End
   End
   Begin XPFrame.FrameXp FrameXp1 
      Height          =   6525
      Left            =   0
      TabIndex        =   22
      Top             =   120
      Width           =   9075
      _ExtentX        =   16007
      _ExtentY        =   11509
      BackColor       =   16761024
      Caption         =   " DATOS DOCUMENTOS"
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
      Begin VB.TextBox dato27 
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
         Left            =   7440
         MaxLength       =   4
         TabIndex        =   70
         Tag             =   "fecha"
         Top             =   3600
         Width           =   975
      End
      Begin VB.TextBox dato26 
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
         Left            =   6960
         MaxLength       =   2
         TabIndex        =   69
         Tag             =   "fecha"
         Top             =   3600
         Width           =   375
      End
      Begin VB.TextBox dato25 
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
         Left            =   6480
         MaxLength       =   2
         TabIndex        =   68
         Tag             =   "fecha"
         Top             =   3600
         Width           =   375
      End
      Begin VB.CommandButton cmdpagar 
         BackColor       =   &H00FF8080&
         Caption         =   "PAGAR"
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
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   67
         Top             =   6120
         Visible         =   0   'False
         Width           =   1215
      End
      Begin XPFrame.FrameXp FrameXp5 
         Height          =   1215
         Left            =   5040
         TabIndex        =   65
         Top             =   1800
         Width           =   3975
         _ExtentX        =   7011
         _ExtentY        =   2143
         BackColor       =   16744576
         Caption         =   "UBICACION"
         CaptionEstilo3D =   1
         BackColor       =   16744576
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Begin VB.Label lblubicacion 
            BackColor       =   &H80000007&
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
            ForeColor       =   &H0000FFFF&
            Height          =   930
            Left            =   0
            TabIndex        =   66
            Top             =   240
            Width           =   3945
         End
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00FFC0C0&
         Caption         =   "PAGAR"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   61
         Top             =   6120
         Visible         =   0   'False
         Width           =   975
      End
      Begin XPFrame.FrameXp FrameXp4 
         Height          =   2055
         Left            =   0
         TabIndex        =   53
         Top             =   3960
         Visible         =   0   'False
         Width           =   4545
         _ExtentX        =   8017
         _ExtentY        =   3625
         BackColor       =   16744576
         Caption         =   " INFORMACION ADICIONAL"
         CaptionEstilo3D =   1
         BackColor       =   16744576
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Begin VB.TextBox dato22 
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
            Left            =   2160
            TabIndex        =   56
            Tag             =   "fecha"
            Top             =   600
            Width           =   2175
         End
         Begin VB.TextBox dato23 
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
            Left            =   2160
            TabIndex        =   55
            Tag             =   "fecha"
            Top             =   240
            Width           =   2175
         End
         Begin VB.TextBox dato24 
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
            Left            =   2160
            TabIndex        =   54
            Tag             =   "fecha"
            Top             =   990
            Width           =   2175
         End
         Begin VB.Label Label9 
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   " LECTURA ACTUAL"
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
            TabIndex        =   59
            Top             =   600
            Width           =   2025
         End
         Begin VB.Label Label15 
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   " LECTURA ANTERIOR"
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
            TabIndex        =   58
            Top             =   270
            Width           =   2025
         End
         Begin VB.Label Label16 
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   " CONSUMO"
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
            Top             =   990
            Width           =   2025
         End
      End
      Begin XPFrame.FrameXp frmpagos 
         Height          =   2055
         Left            =   4560
         TabIndex        =   40
         Top             =   3960
         Visible         =   0   'False
         Width           =   4455
         _ExtentX        =   7858
         _ExtentY        =   3625
         BackColor       =   16761024
         Caption         =   " PAGOS"
         CaptionEstilo3D =   1
         BackColor       =   16761024
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Begin XPFrame.FrameXp tppago 
            Height          =   975
            Left            =   2280
            TabIndex        =   62
            Top             =   240
            Visible         =   0   'False
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   1720
            BackColor       =   16744576
            Caption         =   "Tipo Pago"
            CaptionEstilo3D =   1
            BackColor       =   16744576
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
            Begin VB.Label Label22 
               BackStyle       =   0  'Transparent
               BorderStyle     =   1  'Fixed Single
               Caption         =   " 1 - CHEQUE"
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
               Left            =   120
               TabIndex        =   64
               Top             =   600
               Width           =   1935
            End
            Begin VB.Label Label21 
               BackStyle       =   0  'Transparent
               BorderStyle     =   1  'Fixed Single
               Caption         =   " 0 - EFECTIVO"
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
               Left            =   120
               TabIndex        =   63
               Top             =   240
               Width           =   1935
            End
         End
         Begin VB.TextBox dato30 
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
            Left            =   2640
            MaxLength       =   20
            TabIndex        =   51
            Tag             =   "fecha"
            Top             =   1680
            Width           =   1695
         End
         Begin VB.TextBox dato29 
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
            Left            =   2640
            MaxLength       =   2
            TabIndex        =   49
            Tag             =   "fecha"
            Top             =   1320
            Width           =   495
         End
         Begin VB.TextBox dato20 
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
            Left            =   2640
            MaxLength       =   2
            TabIndex        =   45
            Tag             =   "fecha"
            Top             =   570
            Width           =   495
         End
         Begin VB.TextBox dato21 
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
            Left            =   2640
            MaxLength       =   10
            TabIndex        =   44
            Tag             =   "fecha"
            Top             =   960
            Width           =   1695
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
            Left            =   2640
            MaxLength       =   2
            TabIndex        =   43
            Tag             =   "fecha"
            Top             =   240
            Width           =   375
         End
         Begin VB.TextBox dato18 
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
            Left            =   3000
            MaxLength       =   2
            TabIndex        =   42
            Tag             =   "fecha"
            Top             =   240
            Width           =   375
         End
         Begin VB.TextBox dato19 
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
            Left            =   3360
            MaxLength       =   4
            TabIndex        =   41
            Tag             =   "fecha"
            Top             =   240
            Width           =   975
         End
         Begin VB.Label Label20 
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   " Nº CHEQUE"
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
            Left            =   0
            TabIndex        =   52
            Top             =   1680
            Width           =   2505
         End
         Begin VB.Label Label19 
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   " TIPO PAGO"
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
            Left            =   0
            TabIndex        =   50
            Top             =   1320
            Width           =   2505
         End
         Begin VB.Label Label7 
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   " TIPO COMPROBANTE"
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
            Left            =   0
            TabIndex        =   48
            Top             =   570
            Width           =   2505
         End
         Begin VB.Label Label8 
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   " NUMERO COMPROBANTE"
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
            Left            =   0
            TabIndex        =   47
            Top             =   960
            Width           =   2505
         End
         Begin VB.Label Label14 
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   " FECHA PAGO"
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
            Left            =   0
            TabIndex        =   46
            Top             =   240
            Width           =   2505
         End
      End
      Begin VB.TextBox dato28 
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
         Left            =   2160
         TabIndex        =   38
         Tag             =   "fecha"
         Top             =   3600
         Width           =   1935
      End
      Begin VB.TextBox dato10 
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
         Left            =   2160
         TabIndex        =   9
         Tag             =   "fecha"
         Top             =   2880
         Width           =   1935
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
         Left            =   2175
         MaxLength       =   10
         TabIndex        =   8
         Tag             =   "fecha"
         Top             =   2550
         Width           =   1935
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
         Left            =   3120
         MaxLength       =   4
         TabIndex        =   6
         Tag             =   "fecha"
         Top             =   1845
         Width           =   975
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
         Left            =   2640
         MaxLength       =   2
         TabIndex        =   5
         Tag             =   "fecha"
         Top             =   1845
         Width           =   375
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
         Left            =   2160
         MaxLength       =   2
         TabIndex        =   4
         Tag             =   "fecha"
         Top             =   1845
         Width           =   375
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
         Left            =   7440
         MaxLength       =   4
         TabIndex        =   15
         Tag             =   "fecha"
         Top             =   3240
         Width           =   975
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
         Left            =   6960
         MaxLength       =   2
         TabIndex        =   14
         Tag             =   "fecha"
         Top             =   3240
         Width           =   375
      End
      Begin VB.TextBox dato14 
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
         Left            =   6480
         MaxLength       =   2
         TabIndex        =   13
         Tag             =   "fecha"
         Top             =   3240
         Width           =   375
      End
      Begin VB.TextBox dato13 
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
         TabIndex        =   12
         Tag             =   "fecha"
         Top             =   3240
         Width           =   975
      End
      Begin VB.TextBox dato12 
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
         Left            =   2640
         MaxLength       =   2
         TabIndex        =   11
         Tag             =   "fecha"
         Top             =   3240
         Width           =   375
      End
      Begin VB.TextBox dato11 
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
         TabIndex        =   10
         Tag             =   "fecha"
         Top             =   3240
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
         Left            =   2160
         MaxLength       =   1
         TabIndex        =   7
         Tag             =   "fecha"
         Top             =   2205
         Width           =   375
      End
      Begin VB.TextBox dato3 
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
         Left            =   2160
         MaxLength       =   9
         TabIndex        =   2
         Top             =   1035
         Width           =   1275
      End
      Begin VB.TextBox dato1 
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
         Left            =   2160
         MaxLength       =   2
         TabIndex        =   0
         Tag             =   "fecha"
         Top             =   315
         Width           =   510
      End
      Begin VB.TextBox dato2 
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
         MaxLength       =   20
         TabIndex        =   1
         Tag             =   "fecha"
         Top             =   675
         Width           =   2670
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
         TabIndex        =   3
         Tag             =   "fecha"
         Top             =   1440
         Width           =   375
      End
      Begin VB.Label Label17 
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
         Left            =   4440
         TabIndex        =   71
         Top             =   3600
         Width           =   2025
      End
      Begin VB.Label Label18 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   " SALDO ANTERIOR$"
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
         Left            =   180
         TabIndex        =   39
         Top             =   3600
         Width           =   1905
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   " MONTO DOC.       $"
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
         Left            =   180
         TabIndex        =   37
         Top             =   2880
         Width           =   1905
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   " NUMERO DOC."
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
         Left            =   180
         TabIndex        =   36
         Top             =   2520
         Width           =   1905
      End
      Begin VB.Label Label2 
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
         Left            =   180
         TabIndex        =   35
         Top             =   1845
         Width           =   1905
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   " HASTA"
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
         Left            =   4440
         TabIndex        =   34
         Top             =   3240
         Width           =   1905
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   " DESDE"
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
         Left            =   180
         TabIndex        =   33
         Top             =   3240
         Width           =   1905
      End
      Begin VB.Label LBLFAC 
         BackStyle       =   0  'Transparent
         Caption         =   "F=FACTURA O B=BOLETA"
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
         Left            =   2640
         TabIndex        =   32
         Top             =   2250
         Width           =   2340
      End
      Begin VB.Label LBLDV 
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
         Height          =   330
         Left            =   3465
         TabIndex        =   31
         Top             =   1035
         Width           =   285
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   " TIPO DOCUMENTO"
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
         Left            =   180
         TabIndex        =   30
         Top             =   2205
         Width           =   1905
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   " NUMERO SERVICIO"
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
         Left            =   180
         TabIndex        =   29
         Top             =   675
         Width           =   1905
      End
      Begin VB.Label Label3 
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
         Left            =   180
         TabIndex        =   28
         Top             =   1035
         Width           =   1905
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   " EMPRESA"
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
         Left            =   180
         TabIndex        =   27
         Top             =   1440
         Width           =   1905
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   " TIPO CONSUMO"
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
         Left            =   180
         TabIndex        =   26
         Top             =   315
         Width           =   1905
      End
      Begin VB.Label LBLTIPO 
         BackColor       =   &H80000007&
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
         ForeColor       =   &H0000FFFF&
         Height          =   330
         Left            =   2835
         TabIndex        =   25
         Top             =   315
         Width           =   6165
      End
      Begin VB.Label LBLPROVEEDOR 
         BackColor       =   &H80000007&
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
         ForeColor       =   &H0000FFFF&
         Height          =   330
         Left            =   3735
         TabIndex        =   24
         Top             =   1035
         Width           =   5265
      End
      Begin VB.Label LBLEMPRESA 
         BackColor       =   &H80000007&
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
         ForeColor       =   &H0000FFFF&
         Height          =   330
         Left            =   2655
         TabIndex        =   23
         Top             =   1440
         Width           =   6345
      End
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash opciones 
      Height          =   1095
      Left            =   9000
      TabIndex        =   16
      Top             =   5040
      Width           =   6615
      _cx             =   11668
      _cy             =   1931
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
Attribute VB_Name = "consumo04"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public saldoglobal As Double
Private moneda As String
Private rutpropi As String
Private MODIFI As Integer

Private Sub cmdgrabar_Click()
grabar
retorno
End Sub

Private Sub Command1_Click()
    frmpagos.Visible = True
    dato17.Enabled = True
    dato17.Locked = False
    dato18.Enabled = True
    dato18.Locked = False
    dato19.Enabled = True
    dato19.Locked = False
    DATO20.Enabled = True
    DATO20.Locked = False
    DATO21.Enabled = True
    DATO21.Locked = False
    dato29.Enabled = True
    dato29.Locked = False
    dato30.Enabled = True
    dato30.Locked = False
    If MODIFI = "1" Then
        dato17.SetFocus
    End If
    Label20.Visible = False
    dato30.Visible = False
End Sub

Private Sub cmdpagar_Click()
Check1.Value = 1
If Check1.Value = 1 Then
    frmpagos.Visible = True
    dato17.Enabled = True
    dato17.Locked = False
    dato18.Enabled = True
    dato18.Locked = False
    dato19.Enabled = True
    dato19.Locked = False
    DATO20.Enabled = True
    DATO20.Locked = False
    DATO21.Enabled = True
    DATO21.Locked = False
    dato29.Enabled = True
    dato29.Locked = False
    dato30.Enabled = True
    dato30.Locked = False
    If MODIFI = "1" Then
        dato17.SetFocus
    End If
    Label20.Visible = False
    dato30.Visible = False
Else
    frmpagos.Visible = False
End If
End Sub
 
Private Sub Check1_Click()
If Check1.Value = 1 Then
    frmpagos.Visible = True
    dato17.Enabled = True
    dato17.Locked = False
    dato18.Enabled = True
    dato18.Locked = False
    dato19.Enabled = True
    dato19.Locked = False
    DATO20.Enabled = True
    DATO20.Locked = False
    DATO21.Enabled = True
    DATO21.Locked = False
    dato29.Enabled = True
    dato29.Locked = False
    dato30.Enabled = True
    dato30.Locked = False
    If MODIFI = "1" Then
        dato17.SetFocus
    End If
    Label20.Visible = False
    dato30.Visible = False
Else
    frmpagos.Visible = False
End If
End Sub

Private Sub dato10_KeyPress(KeyAscii As Integer)
    snum = 0: KeyAscii = esNumero(KeyAscii)
    If KeyAscii = 13 And dato10.text <> "" Then
        If CDbl(dato10.text) > 0 Then
            dato10.text = Format(dato10.text, "###,###,###")
            Call Pregunta(dato10, dato11)
        End If
    End If
End Sub
 

Private Sub dato11_KeyPress(KeyAscii As Integer)
    snum = 0: KeyAscii = esNumero(KeyAscii)
    If KeyAscii = 13 Then
        If dato11.text = "" Then dato11.text = Day(fechasistema)
        If CDbl(dato11.text) > 0 Then
            Call ceros(dato11)
            Call Pregunta(dato11, dato12)
        End If
    End If
End Sub

Private Sub dato11_LostFocus()
Call esfechareal(dato11, dato12, dato13, "dd")
End Sub
  
Private Sub dato12_KeyPress(KeyAscii As Integer)
    snum = 0: KeyAscii = esNumero(KeyAscii)
    If KeyAscii = 13 Then
        If dato12.text = "" Then dato12.text = Month(fechasistema)
        If CDbl(dato12.text) > 0 Then
            Call ceros(dato12)
            Call Pregunta(dato12, dato13)
        End If
    End If
End Sub

Private Sub dato12_LostFocus()
Call esfechareal(dato11, dato12, dato13, "mm")
End Sub

Private Sub dato13_KeyPress(KeyAscii As Integer)
    snum = 0: KeyAscii = esNumero(KeyAscii)
    If KeyAscii = 13 Then
        If dato13.text = "" Then dato13.text = Year(fechasistema)
        If CDbl(dato13.text) > 0 Then
            Call ceros(dato13)
            Call Pregunta(dato13, dato14)
        End If
    End If
End Sub

Private Sub dato13_LostFocus()
Call esfechareal(dato11, dato12, dato13, "yyyy")
End Sub

Private Sub dato14_KeyPress(KeyAscii As Integer)
    snum = 0: KeyAscii = esNumero(KeyAscii)
    If KeyAscii = 13 Then
        If dato14.text = "" Then dato14.text = Day(fechasistema)
        If CDbl(dato14.text) > 0 Then
            Call ceros(dato14)
            Call Pregunta(dato14, dato15)
        End If
    End If
End Sub

Private Sub dato14_LostFocus()
Call esfechareal(dato14, dato15, dato16, "dd")
End Sub

Private Sub dato15_KeyPress(KeyAscii As Integer)
    snum = 0: KeyAscii = esNumero(KeyAscii)
    If KeyAscii = 13 Then
        If dato15.text = "" Then dato15.text = Month(fechasistema)
        If CDbl(dato15.text) > 0 Then
            Call ceros(dato15)
            Call Pregunta(dato15, dato16)
        End If
    End If
End Sub

Private Sub dato15_LostFocus()
Call esfechareal(dato14, dato15, dato16, "mm")
End Sub

Private Sub dato16_KeyPress(KeyAscii As Integer)
    snum = 0: KeyAscii = esNumero(KeyAscii)
    If KeyAscii = 13 Then
        If dato16.text = "" Then dato16.text = Year(fechasistema)
        If CDbl(dato16.text) > 0 Then
            Call ceros(dato16)
            Call Pregunta(dato16, dato28)
        End If
    End If
End Sub

Private Sub dato16_LostFocus()
Call esfechareal(dato14, dato15, dato16, "yyyy")
End Sub

 

Private Sub dato17_KeyPress(KeyAscii As Integer)
    snum = 0: KeyAscii = esNumero(KeyAscii)
    If KeyAscii = 13 Then
        If dato17.text = "" Then dato17.text = Day(fechasistema)
        If CDbl(dato17.text) > 0 Then
            Call ceros(dato17)
            Call Pregunta(dato17, dato18)
        End If
    End If
End Sub

Private Sub dato17_LostFocus()
Call esfechareal(dato17, dato18, dato19, "dd")
End Sub
 
Private Sub dato18_KeyPress(KeyAscii As Integer)
    snum = 0: KeyAscii = esNumero(KeyAscii)
    If KeyAscii = 13 Then
        If dato18.text = "" Then dato18.text = Month(fechasistema)
        If CDbl(dato18.text) > 0 Then
            Call ceros(dato18)
            Call Pregunta(dato18, dato19)
        End If
    End If
End Sub

Private Sub dato18_LostFocus()
Call esfechareal(dato17, dato18, dato19, "mm")
End Sub

Private Sub dato19_KeyPress(KeyAscii As Integer)
    snum = 0: KeyAscii = esNumero(KeyAscii)
    If KeyAscii = 13 Then
        If dato19.text = "" Then dato19.text = Year(fechasistema)
        If CDbl(dato19.text) > 0 Then
            Call ceros(dato19)
            Call Pregunta(dato19, DATO20)
        End If
    End If
End Sub

Private Sub dato19_LostFocus()
Call esfechareal(dato17, dato18, dato19, "yyyy")
End Sub
 
Private Sub dato20_GotFocus()
DATO20.text = "CE"
End Sub

Private Sub DATO20_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        Call Pregunta(DATO20, DATO21)
    End If
End Sub

Private Sub dato21_KeyPress(KeyAscii As Integer)
    snum = 0: KeyAscii = esNumero(KeyAscii)
    If KeyAscii = 13 And DATO21.text <> "" Then
        If CDbl(DATO21.text) > 0 Then
            Call ceros(DATO21)
            Call Pregunta(DATO21, dato29)
        End If
    End If
End Sub
 

Private Sub dato22_KeyPress(KeyAscii As Integer)
    snum = 0: KeyAscii = esNumero(KeyAscii)
    If KeyAscii = 13 And DATO22.text <> "" Then
        If CDbl(DATO22.text) > 0 Then
            DATO22.text = Format(DATO22.text, "###,###,###")
            dato24.text = Format(CDbl(DATO22.text - dato23.text), "###,###,###")
            Call Pregunta(DATO22, dato24)
        End If
    End If
End Sub

Private Sub dato23_KeyPress(KeyAscii As Integer)
    snum = 0: KeyAscii = esNumero(KeyAscii)
    If KeyAscii = 13 And dato23.text <> "" Then
        If CDbl(dato23.text) > 0 Then
            dato23.text = Format(dato23.text, "###,###,###")
            Call Pregunta(dato23, DATO22)
        End If
    End If
End Sub

Private Sub dato24_KeyPress(KeyAscii As Integer)
    snum = 0: KeyAscii = esNumero(KeyAscii)
    If KeyAscii = 13 And dato24.text <> "" Then
        If CDbl(dato24.text) > 0 Then
           Call Pregunta(dato24, DATO25)
           If MODIFI = "1" Then
            DATO25.Enabled = True
            DATO25.Locked = False
            dato26.Enabled = True
            dato26.Locked = False
            dato27.Enabled = True
            dato27.Locked = False
           End If
           
           CmdGrabar.Visible = True
        End If
    End If
End Sub

Private Sub dato25_KeyPress(KeyAscii As Integer)
    snum = 0: KeyAscii = esNumero(KeyAscii)
    If KeyAscii = 13 Then
            Call ceros(DATO25)
            Call Pregunta(DATO25, dato26)
    End If
End Sub

Private Sub dato25_LostFocus()
If DATO25.text <> "00" Then
    Call esfechareal(DATO25, dato26, dato27, "dd")
End If

End Sub

Private Sub dato26_KeyPress(KeyAscii As Integer)
    snum = 0: KeyAscii = esNumero(KeyAscii)
    If KeyAscii = 13 Then
            Call ceros(dato26)
            Call Pregunta(dato26, dato27)
    End If
End Sub

Private Sub dato26_LostFocus()
If dato26.text <> "00" Then
    Call esfechareal(DATO25, dato26, dato27, "mm")
End If
End Sub

Private Sub dato27_KeyPress(KeyAscii As Integer)
    snum = 0: KeyAscii = esNumero(KeyAscii)
    If KeyAscii = 13 Then
            Call ceros(dato27)
    End If
End Sub

 

Private Sub dato29_GotFocus()
tppago.Visible = True
End Sub

Private Sub dato29_KeyPress(KeyAscii As Integer)
 snum = 0: KeyAscii = esNumero(KeyAscii)
    If KeyAscii = 13 And dato29.text <> "" Then
        If dato29.text = "0" Then
            CmdGrabar.Visible = True
            Label20.Visible = False
            dato30.Visible = False
        Else
            Label20.Visible = True
            dato30.Visible = True
            dato30.Enabled = True
            dato30.Locked = False
            
            dato30.SetFocus
        End If
    End If
End Sub

Private Sub dato29_LostFocus()
tppago.Visible = False
End Sub

Private Sub dato30_KeyPress(KeyAscii As Integer)
    snum = 0: KeyAscii = esNumero(KeyAscii)
    If KeyAscii = 13 And dato30.text <> "" Then
        If CDbl(dato30.text) > 0 Then
            Call ceros(dato30)
            CmdGrabar.Visible = True
        End If
    End If
End Sub

Private Sub dato27_LostFocus()
If dato27.text <> "0000" Then
    Call esfechareal(DATO25, dato26, dato27, "yyyy")
End If
End Sub
 

Private Sub dato28_KeyPress(KeyAscii As Integer)
    KeyAscii = esNumero(KeyAscii)
    If KeyAscii = 13 And dato28.text <> "" Then
        dato28.text = Format(dato28.text, "###,###,##0")
        If dato1.text = "01" Or dato1.text = "02" Then
            Call Pregunta(dato28, dato23)
        Else
            CmdGrabar.Visible = True
        End If
        
    End If
End Sub

Private Sub dato5_LostFocus()
 Call esfechareal(DATO5, dato6, dato7, "dd")
End Sub

Private Sub dato6_LostFocus()
 Call esfechareal(DATO5, dato6, dato7, "mm")
End Sub

Private Sub dato7_LostFocus()
Call esfechareal(DATO5, dato6, dato7, "yyyy")
End Sub

Private Sub Grid1_DblClick()
 If Grid1.Rows > 1 Then
    Call cargadocumento(dato1.text, dato2.text, Grid1.Cell(Grid1.ActiveCell.row, 1).text, Grid1.Cell(Grid1.ActiveCell.row, 2).text)
 End If
 
End Sub

Private Sub Grid2_DblClick()
dato1.text = Grid2.Cell(Grid2.ActiveCell.row, 1).text
dato2.text = Grid2.Cell(Grid2.ActiveCell.row, 2).text
Call dato2_KeyPress(13)
End Sub

 

 Private Sub MANUAL_KeyPress(KeyAscii As Integer)
If UCase(Chr(KeyAscii)) = "M" Then Call opciones_FSCommand("modifica", "")
If UCase(Chr(KeyAscii)) = "E" Then Call opciones_FSCommand("elimina", "")
If UCase(Chr(KeyAscii)) = "S" Then Call opciones_FSCommand("siguiente", "")
If UCase(Chr(KeyAscii)) = "A" Then Call opciones_FSCommand("anterior", "")
If UCase(Chr(KeyAscii)) = "R" Then Call opciones_FSCommand("retorno", "")
If UCase(Chr(KeyAscii)) = "I" Then Call opciones_FSCommand("imprime", "")
End Sub

Private Sub Form_Load()
Call CENTRAR(Me)
    Call Conectar_BD
    Rem Call Funciones_Forms_M_Productos.Conecta_Maestro_Productos
    sc = 0
    opciones.Visible = False
Rem Call RECUPERAFECHA

Call CARGAPERMISO(Me.Name)
Call CARGAGRILLA
Call CARGAGRILLA2
End Sub

Sub leer()
    campos(0, 0) = "tipo"
    campos(1, 0) = "numeroservicio"
    campos(2, 0) = "proveedor"
    campos(3, 0) = "empresacontable"
    campos(4, 0) = "diapago"
    campos(5, 0) = "documento"
    campos(6, 0) = "ubicacion"
    campos(7, 0) = ""
    campos(0, 2) = clientesistema & "consumos_basicos.maestro_unidades_consumo"
    condicion = "tipo='" + dato1.text + "' and numeroservicio='" + dato2.text + "' "
    op = 5
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)
    If sqlconta.status = 4 Then dato3.SetFocus: GoTo no:
    carga
    disponible (True)
    habilita (False)
        
no:
End Sub

Sub carga()
    habilita (False)
    dato1.text = sqlconta.response(0, 3)
    If sqlconta.response(0, 3) = "01" Or sqlconta.response(0, 3) = "02" Then
        FrameXp4.Visible = True
    Else
        FrameXp4.Visible = False
    End If
    
    dato2.text = sqlconta.response(1, 3)
    dato3.text = Mid(sqlconta.response(2, 3), 1, 9)
    LBLDV.Caption = Mid(sqlconta.response(2, 3), 10, 1)
    dato4.text = sqlconta.response(3, 3)
    dato17.text = sqlconta.response(4, 3)
    dato8.text = sqlconta.response(5, 3)
    lblubicacion.Caption = sqlconta.response(6, 3)
    LBLEMPRESA.Caption = leerempresa(dato4.text)
    LBLTIPO.Caption = leetipoconsumo(dato1.text)
    LBLPROVEEDOR.Caption = leerProveedor(dato3.text + LBLDV.Caption)
    Call leerpagosanteriores(dato1.text, dato2.text)
    
    
fin:
End Sub

Sub habilita(ByVal condicion As Boolean)
    
    dato1.Locked = condicion
    dato2.Locked = condicion
    dato3.Locked = condicion
    dato4.Locked = condicion
    DATO5.Locked = condicion
    dato6.Locked = condicion
    dato7.Locked = condicion
    dato8.Locked = condicion
    dato9.Locked = condicion
    dato10.Locked = condicion
    dato11.Locked = condicion
    dato12.Locked = condicion
    dato13.Locked = condicion
    dato14.Locked = condicion
    dato15.Locked = condicion
    dato16.Locked = condicion
    dato17.Locked = condicion
    dato18.Locked = condicion
    dato19.Locked = condicion
    DATO20.Locked = condicion
    DATO21.Locked = condicion
    DATO22.Locked = condicion
    dato23.Locked = condicion
    dato24.Locked = condicion
    DATO25.Locked = condicion
    dato26.Locked = condicion
    dato27.Locked = condicion
    dato28.Locked = condicion
    dato29.Locked = condicion
    dato30.Locked = condicion
    
End Sub
Sub disponible(ByVal condicion As Boolean)
    
    dato1.Enabled = condicion
    dato2.Enabled = condicion
    dato3.Enabled = condicion
    dato4.Enabled = condicion
    DATO5.Enabled = condicion
    dato6.Enabled = condicion
    dato7.Enabled = condicion
    dato8.Enabled = condicion
    dato9.Enabled = condicion
    dato10.Enabled = condicion
    dato11.Enabled = condicion
    dato12.Enabled = condicion
    dato13.Enabled = condicion
    dato14.Enabled = condicion
    dato15.Enabled = condicion
    dato16.Enabled = condicion
    dato17.Enabled = condicion
    dato18.Enabled = condicion
    dato19.Enabled = condicion
    DATO20.Enabled = condicion
    DATO21.Enabled = condicion
    DATO22.Enabled = condicion
    dato23.Enabled = condicion
    dato24.Enabled = condicion
    DATO25.Enabled = condicion
    dato26.Enabled = condicion
    dato27.Enabled = condicion
    dato28.Enabled = condicion
    dato29.Enabled = condicion
    dato30.Enabled = condicion
End Sub
Sub Pregunta(ByRef caja As TextBox, ByRef caja1 As TextBox)
    If caja.text = "" And sc = 0 Then caja.SetFocus
    If caja.text <> "" Or sc = 1 Then caja1.Enabled = True: caja1.SetFocus
    sc = 0
End Sub


Sub ayudaarrendadores(ByRef caja As TextBox)
    Dim campos As Variant
    Dim cfijo As Variant
    Dim largo As Variant
    campos = Array("rut", "nombre")
    largo = Array("11s", "40s")
    cfijo = "rut like '%%'"
    cabezas = Array("Rut", "Nombre")
    mensajeAyuda = "Ayuda de Arrendadores"
       
    Call cargaAyudaT(Servidor, clientesistema & "arriendos", Usuario, password, ".maestro_arrendadores", caja, campos, cfijo, largo, 2)
    caja.Enabled = True
    caja.SetFocus
End Sub
Sub ayudaarrendatarios(ByRef caja As TextBox)
    Dim campos As Variant
    Dim cfijo As Variant
    Dim largo As Variant
    campos = Array("rut", "nombre")
    largo = Array("11s", "40s")
    cfijo = "rut like '%%'"
    cabezas = Array("Rut", "Nombre")
    mensajeAyuda = "Ayuda de Arrendatarios"
       
    Call cargaAyudaT(Servidor, clientesistema & "arriendos", Usuario, password, ".maestro_arrendatarios", caja, campos, cfijo, largo, 2)
    caja.Enabled = True
    caja.SetFocus
End Sub
Sub ayudapropiedades(ByRef caja As TextBox)
    Dim campos As Variant
    Dim cfijo As Variant
    Dim largo As Variant
    campos = Array("codigopropiedad", "direccion")
    largo = Array("11s", "40s")
    cfijo = "codigopropiedad like '%%'"
    cabezas = Array("Codigo", "Nombre")
    mensajeAyuda = "Ayuda de Propiedades"
       
    Call cargaAyudaT(Servidor, clientesistema & "arriendos", Usuario, password, ".maestro_propiedades", caja, campos, cfijo, largo, 2)
    caja.Enabled = True
    caja.SetFocus
End Sub


Sub flechas(ByRef caja As TextBox, ByRef caja1 As TextBox, ByRef codigo As Integer)
    If codigo = 38 And caja.Enabled = True Then caja.SetFocus
    If codigo = 40 And caja1.Enabled = True Then caja1.SetFocus
End Sub

Sub grabar()
 Dim fecha As String
 Dim desde As String
 Dim hasta As String
 Dim corte As String
 Dim pago As String
 Dim proveedor As String
 
 fecha = dato7.text & "-" & dato6.text & "-" & DATO5.text
 desde = dato13.text & "-" & dato12.text & "-" & dato11.text
 hasta = dato16.text & "-" & dato15.text & "-" & dato14.text
 pago = dato19.text & "-" & dato18.text & "-" & dato17.text
 corte = dato27.text & "-" & dato26.text & "-" & DATO25.text
 
    campos(0, 0) = "tipo"
    campos(1, 0) = "numeroservicio"
    campos(2, 0) = "fecha"
    campos(3, 0) = "tipodocumento"
    campos(4, 0) = "numerodocumento"
    campos(5, 0) = "montodocumento"
    campos(6, 0) = "desde"
    campos(7, 0) = "hasta"
    campos(8, 0) = "fechapago"
    campos(9, 0) = "tipocomprobante"
    campos(10, 0) = "numerocomprobante"
    campos(11, 0) = "lecturaactual"
    campos(12, 0) = "lecturaanterior"
    campos(13, 0) = "consumo"
    campos(14, 0) = "saldoanterior"
    campos(15, 0) = "fechacorte"
    campos(16, 0) = "tipopago"
    campos(17, 0) = "numerocheque"
    campos(18, 0) = "rut"
    campos(19, 0) = ""
    campos(0, 1) = dato1.text
    campos(1, 1) = dato2.text
    campos(2, 1) = fecha
    campos(3, 1) = dato8.text
    campos(4, 1) = dato9.text
    campos(5, 1) = CDbl(dato10.text)
    campos(6, 1) = desde
    campos(7, 1) = hasta
    campos(8, 1) = pago
    campos(9, 1) = DATO20.text
    campos(10, 1) = DATO21.text
    campos(11, 1) = DATO22.text
    campos(12, 1) = dato23.text
    campos(13, 1) = dato24.text
    campos(14, 1) = CDbl(dato28.text)
    campos(15, 1) = corte
    campos(16, 1) = dato29.text
    campos(17, 1) = dato30.text
    campos(18, 1) = dato3.text + LBLDV.Caption
    
    
    campos(0, 2) = clientesistema & "consumos_basicos.detalle_servicios"
    
    If MODIFI = 1 Then
    condicion = "tipo='" + dato1.text + "' and numeroservicio='" + dato2.text + "' and numerodocumento='" & dato9.text & "' "
    
    End If
    
    If MODIFI = 1 Then op = 3 Else op = 2
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)
    
    End Sub
 
Sub ELIMINAR()
Dim fecha As String

    fecha = dato7.text & "-" & dato6.text & "-" & DATO5.text
    campos(0, 2) = clientesistema & "consumos_basicos.detalle_servicios "
    condicion = "tipo='" + dato1.text + "' and numeroservicio='" + dato2.text + "' and fecha='" & fecha & "' and numerodocumento='" & dato9.text & "' "
    
    op = 4
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)

    

End Sub
  

Private Sub opciones_FSCommand(ByVal command As String, ByVal args As String)

If command = "retorno" Then retorno
If command = "modifica" Then
    MODIFI = 1
    dato3.Enabled = True
    dato3.Locked = False
    
    dato3.SetFocus
    
    
End If

If command = "elimina" Then
    If Verifica_Permiso(Me.Caption, "elimina") = True Then
        ELIMINA
    Else
    MsgBox ("NO PUEDE ELIMINAR DOCUMENTOS ")
    End If
    
End If


 



End Sub
Sub ELIMINA()
 
disponible (True)
habilita (False)
ELIMINAR
limpia
opciones.Visible = False
dato1.SetFocus
 
End Sub

Sub modifica()
disponible (True)
habilita (False)
dato1.Enabled = False
dato2.SetFocus
MODIFI = 1

End Sub
Sub retorno()

disponible (True)
habilita (False)
limpia
opciones.Visible = False
dato1.Enabled = True
dato1.SetFocus
Grid1.Rows = 1
MODIFI = 0
no:
 
 
    
End Sub
Sub limpia()
    dato1.text = ""
    dato2.text = ""
    dato3.text = ""
    dato4.text = ""
    DATO5.text = ""
    dato6.text = ""
    dato7.text = ""
    dato8.text = ""
    dato9.text = ""
    dato10.text = ""
    dato11.text = ""
    dato12.text = ""
    dato13.text = ""
    dato14.text = ""
    dato15.text = ""
    dato16.text = ""
    dato17.text = ""
    dato18.text = ""
    dato19.text = ""
    DATO20.text = ""
    DATO21.text = ""
    DATO22.text = ""
    dato23.text = ""
    dato24.text = ""
    DATO25.text = ""
    dato26.text = ""
    dato27.text = ""
    dato28.text = ""
    dato29.text = ""
    dato30.text = ""
    lblubicacion.Caption = ""
    LBLTIPO.Caption = ""
    LBLPROVEEDOR.Caption = ""
    LBLEMPRESA.Caption = ""
    LBLDV.Caption = ""
    Check1.Value = 0
    frmpagos.Visible = False
    Grid1.Rows = 1
    Grid2.Rows = 1
    CmdGrabar.Visible = False
    cmdpagar.Visible = False
    
    
End Sub
 
Sub cargatexto(ByRef caja As TextBox)
caja.SelStart = 0: caja.SelLength = Len(caja.text)
End Sub

Private Sub opciones_GotFocus()
MANUAL.SetFocus
End Sub

 
Sub CARGAGRILLA()
    Dim formatogrilla2(10, 10)
    formatogrilla2(1, 1) = "FECHA"
    formatogrilla2(1, 2) = "NUMERO DOCUMENTO"
    formatogrilla2(1, 3) = "MONTO"
    formatogrilla2(1, 4) = "PAGADO"
  
    
    
    Rem LARGO DE LOS DATOS
    formatogrilla2(2, 1) = "8"
    formatogrilla2(2, 2) = "8"
    formatogrilla2(2, 3) = "8"
    formatogrilla2(2, 4) = "8"
 
    
    Rem TIPO DE DATOS C=CEROS,S=STRING,N=NUMERICO,D=DATE,
    formatogrilla2(3, 1) = "D"
    formatogrilla2(3, 2) = "N"
    formatogrilla2(3, 3) = "N"
    formatogrilla2(3, 4) = "S"
 
    
    Rem FORMATO GRILLA
    
    formatogrilla2(4, 1) = ""
    formatogrilla2(4, 2) = ""
    formatogrilla2(4, 3) = " ###,###,##0"
 
    
    Rem LOCCKED
    formatogrilla2(5, 1) = "TRUE"
    formatogrilla2(5, 2) = "TRUE"
    formatogrilla2(5, 3) = "TRUE"
    formatogrilla2(5, 4) = "TRUE"

    
    
    Rem VALOR MAXIMO
    
    Grid1.Cols = 5
    Grid1.Rows = 1
    Grid1.AllowUserResizing = False
    Grid1.DisplayFocusRect = False
    Grid1.ExtendLastCol = True
    Grid1.BoldFixedCell = False
    Grid1.DrawMode = cellOwnerDraw
    Grid1.Appearance = Flat
    Grid1.ScrollBarStyle = Flat
    Grid1.FixedRowColStyle = Flat
    Grid1.BackColorFixed = RGB(90, 158, 214)
    Grid1.BackColorFixedSel = RGB(110, 180, 230)
    Grid1.BackColorBkg = RGB(90, 158, 214)
    Grid1.BackColorScrollBar = RGB(231, 235, 247)
    Grid1.BackColor1 = RGB(231, 235, 247)
    Grid1.BackColor2 = RGB(239, 243, 255)
    Grid1.GridColor = RGB(148, 190, 231)
    Grid1.Column(0).Width = 0
    
    For k = 1 To Grid1.Cols - 1
        Grid1.Cell(0, k).text = formatogrilla2(1, k)
        Grid1.Column(k).Width = Val(formatogrilla2(2, k)) * 9
        Grid1.Column(k).MaxLength = Val(formatogrilla2(2, k))
        Grid1.Column(k).FormatString = formatogrilla2(4, k)
        Grid1.Column(k).Locked = formatogrilla2(5, k)
        If formatogrilla2(3, k) = "N" Then Grid1.Column(k).Alignment = cellRightCenter
        If formatogrilla2(3, k) = "S" Then Grid1.Column(k).Alignment = cellLeftTop
        If formatogrilla2(3, k) = "D" Then Grid1.Column(k).CellType = cellCalendar
        
    Next k
     Grid1.Column(4).CellType = cellCheckBox
     
     
     
    End Sub

Sub CARGAGRILLA2()
    Dim formatogrilla2(10, 12)
    formatogrilla2(1, 1) = "TIPO"
    formatogrilla2(1, 2) = "SERVICIO"
    formatogrilla2(1, 3) = "PROVEEDOR"
    formatogrilla2(1, 4) = "EMPRESA"
    formatogrilla2(1, 5) = "DIA PAGO  "
    formatogrilla2(1, 6) = "DOC"
    formatogrilla2(1, 7) = "UBICACION"
    formatogrilla2(1, 8) = "MEDIDOR "
    formatogrilla2(1, 9) = "TARIFA"
    
    Rem LARGO DE LOS DATOS
    formatogrilla2(2, 1) = "10"
    formatogrilla2(2, 2) = "10"
    formatogrilla2(2, 3) = "10"
    formatogrilla2(2, 4) = "10"
    formatogrilla2(2, 5) = "10"
    formatogrilla2(2, 6) = "10"
    formatogrilla2(2, 7) = "10"
    formatogrilla2(2, 8) = "7"
    formatogrilla2(2, 9) = "10"
    formatogrilla2(2, 10) = "10"
    formatogrilla2(2, 11) = "10"
    
    Rem TIPO DE DATOS C=CEROS,S=STRING,N=NUMERICO,D=DATE,
    formatogrilla2(3, 1) = "S"
    formatogrilla2(3, 2) = "S"
    formatogrilla2(3, 3) = "S"
    formatogrilla2(3, 4) = "S"
    formatogrilla2(3, 5) = "S"
    formatogrilla2(3, 6) = "S"
    formatogrilla2(3, 7) = "S"
    formatogrilla2(3, 8) = "S"
    formatogrilla2(3, 9) = "N"
    formatogrilla2(3, 10) = "S"
    formatogrilla2(3, 11) = "S"
    
    Rem FORMATO GRILLA
    
    formatogrilla2(4, 9) = " ###,###,##0.00"
    
    
    Rem LOCCKED
    formatogrilla2(5, 1) = "TRUE"
    formatogrilla2(5, 2) = "TRUE"
    formatogrilla2(5, 3) = "TRUE"
    formatogrilla2(5, 4) = "TRUE"
    formatogrilla2(5, 5) = "TRUE"
    formatogrilla2(5, 6) = "TRUE"
    formatogrilla2(5, 7) = "TRUE"
    formatogrilla2(5, 8) = "TRUE"
    formatogrilla2(5, 9) = "TRUE"
    formatogrilla2(5, 10) = "TRUE"
    formatogrilla2(5, 11) = "TRUE"
    
    
    Rem VALOR MAXIMO
    
    Grid2.Cols = 10
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
        Grid2.Column(k).Width = Val(formatogrilla2(2, k)) * 9
        Grid2.Column(k).MaxLength = Val(formatogrilla2(2, k))
        Grid2.Column(k).FormatString = formatogrilla2(4, k)
        Grid2.Column(k).Locked = formatogrilla2(5, k)
        If formatogrilla2(3, k) = "N" Then Grid2.Column(k).Alignment = cellRightCenter
        If formatogrilla2(3, k) = "S" Then Grid2.Column(k).Alignment = cellLeftTop
        
        
        If formatogrilla2(3, k) = "D" Then Grid2.Column(k).CellType = cellCalendar
        
    Next k
   
    
    
    
    
    End Sub
Public Sub leerpagosanteriores(tipo, numero)
Dim csql As New rdoQuery
 Dim resultados As rdoResultset

 Set csql.ActiveConnection = contadb
 csql.sql = "select fecha,numerodocumento,montodocumento,ifnull(fechapago,'0') from " & clientesistema & "consumos_basicos.detalle_servicios "
 csql.sql = csql.sql + "where tipo='" & tipo & "' and numeroservicio='" & numero & "' "
 csql.sql = csql.sql + "order by fecha "
 csql.Execute
 Grid1.Rows = 1
 If csql.RowsAffected > 0 Then
    Set resultados = csql.OpenResultset
    While resultados.EOF = False
        Grid1.Rows = Grid1.Rows + 1
        Grid1.Cell(Grid1.Rows - 1, 1).text = Format(resultados(0), "dd-mm-yyyy")
        Grid1.Cell(Grid1.Rows - 1, 2).text = resultados(1)
        Grid1.Cell(Grid1.Rows - 1, 3).text = resultados(2)
        If resultados(3) = "0000-00-00" Then
            Grid1.Cell(Grid1.Rows - 1, 4).text = 0
        Else
            Grid1.Cell(Grid1.Rows - 1, 4).text = 1
        End If
        resultados.MoveNext
    Wend
  End If
 csql.Close
 Set csql = Nothing
 Set resultados = Nothing
End Sub

 Public Sub leerCONSUMOS()
 Dim csql As New rdoQuery
 Dim resultados As rdoResultset

 Set csql.ActiveConnection = contadb
 csql.sql = "select * from " & clientesistema & "consumos_basicos.maestro_unidades_consumo "
 csql.sql = csql.sql + "where tipo='" + dato1.text + "' "
 csql.sql = csql.sql + "order by empresacontable "
 csql.Execute
 Grid2.Rows = 1
 If csql.RowsAffected > 0 Then
    Set resultados = csql.OpenResultset
    While resultados.EOF = False
    Grid2.Rows = Grid2.Rows + 1
    Grid2.Cell(Grid2.Rows - 1, 1).text = resultados(0) + "=" + leetipoconsumo(resultados(0))
    Grid2.Cell(Grid2.Rows - 1, 2).text = resultados(1)
    Grid2.Cell(Grid2.Rows - 1, 3).text = resultados(2)
    Grid2.Cell(Grid2.Rows - 1, 4).text = resultados(3) + "=" + leerempresa(resultados(3))
    Grid2.Cell(Grid2.Rows - 1, 5).text = resultados(4)
    Grid2.Cell(Grid2.Rows - 1, 6).text = resultados(5)
    Grid2.Cell(Grid2.Rows - 1, 7).text = resultados(6)
    Grid2.Cell(Grid2.Rows - 1, 8).text = resultados(7)
    Grid2.Cell(Grid2.Rows - 1, 9).text = resultados(8)
    




    resultados.MoveNext



    Wend


  End If
 csql.Close
 Set csql = Nothing
 Set resultados = Nothing

 End Sub

Private Function leemonedas(codigo) As String
Dim csql As New rdoQuery
Dim resultados As rdoResultset

Set csql.ActiveConnection = contadb

csql.sql = "select nombremoneda from " & clientesistema & "arriendos" & ".maestro_monedas where codigomoneda='" & codigo & "'"
csql.Execute
leemonedas = ""
If csql.RowsAffected > 0 Then
Set resultados = csql.OpenResultset
leemonedas = resultados(0)
End If
Set resultados = Nothing
csql.Close
Set csql = Nothing

End Function

Public Function LEERULTIMOFOLIOcontrato() As String

    Dim resultados As rdoResultset
    Dim csql As New rdoQuery
    
        Set csql.ActiveConnection = contadb

            csql.sql = "select IFNULL(max(numero),0) from " + clientesistema + "arriendos.contratos_arriendo"
            
            csql.Execute
    If csql.RowsAffected > 0 Then
    Set resultados = csql.OpenResultset
    
    
        LEERULTIMOFOLIOcontrato = Format(resultados(0) + 1, "0000000000")
    End If
    
End Function



Private Sub dato1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF2 Then Call ayudaTIPO(dato1)
    If KeyCode = 38 Then Unload Me: GoTo no:
    
    Call flechas(dato1, dato2, KeyCode)
no:
End Sub
Sub ayudaTIPO(ByRef caja As TextBox)
    Dim campos As Variant
    Dim cfijo As Variant
    Dim largo As Variant
    campos = Array("codigo", "nombre")
    cabezas = Array("CODIGO", "NOMBRE")
    largo = Array("4N", "40s")
    mensajeAyuda = "Ayuda Tipos de Consumos Basicos"
    cfijo = "no"
    
    Call cargaAyudaT(Servidor, clientesistema + "consumos_basicos", Usuario, password, "maestro_tipo_consumos", caja, campos, cfijo, largo, 2)
    If caja.text = "" Then caja.SetFocus: GoTo no
    caja.Enabled = True
    caja.SetFocus


no:

End Sub

Private Sub dato2_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF2 Then Call ayudaservicios(dato2)
    
    Call flechas(dato1, dato3, KeyCode)
    
End Sub
Sub ayudaservicios(ByRef caja As TextBox)
    Dim campos As Variant
    Dim cfijo As Variant
    Dim largo As Variant
    campos = Array("numeroservicio", "ubicacion")
    cabezas = Array("CODIGO", "NOMBRE")
    largo = Array("20N", "40s")
    mensajeAyuda = "Ayuda tipo de creditos"
    cfijo = "tipo='" + dato1.text + "'"
    
    Call cargaAyudaT(Servidor, clientesistema + "consumos_basicos", Usuario, password, "maestro_unidades_consumo", caja, campos, cfijo, largo, 2)
    If caja.text = "" Then caja.SetFocus: GoTo no
    caja.Enabled = True
    caja.SetFocus


no:

End Sub


Private Sub dato3_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF2 Then Call ayudaProveedor(dato3)
    
    Call flechas(dato2, dato4, KeyCode)
End Sub
Sub ayudaProveedor(ByRef caja As TextBox)
    Dim campos As Variant
    Dim cfijo As Variant
    Dim largo As Variant
    campos = Array("rut", "nombre")
    largo = Array("13n", "30s")
    cfijo = "no"
    mensajeAyuda = "Ayuda Proveedores"
    cabezas = Array("rut", "nombre")

    Call cargaAyudaT(Servidor, basedatos, Usuario, password, clientesistema + "consumos_basicos.proveedores", dato3, campos, cfijo, largo, 2)
    If dato3.text = "" Then dato3.SetFocus: GoTo no:
    caja.Enabled = True
    caja.SetFocus
no:
End Sub
    
 Private Sub dato4_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF2 Then Call ayudaempresa(dato4)
    
    Call flechas(dato3, DATO5, KeyCode)


End Sub
Sub ayudaempresa(ByRef caja As TextBox)
    Dim campos As Variant
    Dim cfijo As Variant
    Dim largo As Variant
    campos = Array("codigoempresa", "nombre")
    cabezas = Array("CODIGO", "NOMBRE")
    largo = Array("4N", "40s")
    mensajeAyuda = "Ayuda Empresas"
    cfijo = "no"
    
    Call cargaAyudaT(Servidor, clientesistema + "conta", Usuario, password, "maestroempresas", caja, campos, cfijo, largo, 2)
    If caja.text = "" Then caja.SetFocus: GoTo no
    caja.Enabled = True
    caja.SetFocus


no:

End Sub


Private Sub dato5_KeyDown(KeyCode As Integer, Shift As Integer)
    Call flechas(dato4, dato6, KeyCode)
End Sub
Private Sub dato6_KeyDown(KeyCode As Integer, Shift As Integer)
        Call flechas(DATO5, dato7, KeyCode)
End Sub
Private Sub dato7_KeyDown(KeyCode As Integer, Shift As Integer)
    Call flechas(dato6, dato8, KeyCode)
End Sub
Private Sub dato8_KeyDown(KeyCode As Integer, Shift As Integer)
    LBLFAC.Visible = True
    Call flechas(dato7, dato9, KeyCode)
End Sub
Private Sub dato9_KeyDown(KeyCode As Integer, Shift As Integer)
      Call flechas(dato8, dato9, KeyCode)
End Sub
Sub ayudatipomonedas(ByRef caja As TextBox)
    Dim campos As Variant
    Dim cfijo As Variant
    Dim largo As Variant
    campos = Array("codigo", "nombre")
    cabezas = Array("CODIGO", "NOMBRE")
    largo = Array("4N", "40s")
    mensajeAyuda = "Ayuda tipo de monedas"
    cfijo = "no"
    
    Call cargaAyudaT(Servidor, clientesistema + "creditos_bancarios", Usuario, password, "maestro_tipo_monedas", caja, campos, cfijo, largo, 2)
    If caja.text = "" Then caja.SetFocus: GoTo no
    caja.Enabled = True
    caja.SetFocus


no:

End Sub



Private Sub dato1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then Unload Me
    snum = 0: KeyAscii = esNumero(KeyAscii)
    If KeyAscii = 13 Then
        
    Call ceros(dato1)
    If leetipoconsumo(dato1.text) <> "" Then
    LBLTIPO.Caption = leetipoconsumo(dato1.text)
    dato2.Enabled = True
    leerCONSUMOS
    
    dato2.SetFocus
    Else
    dato1.SetFocus
    End If
    End If
    
    
    
    
End Sub
Private Sub dato2_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    
    If KeyAscii = 13 Then
    Call Pregunta(dato2, dato3)
    leer
    End If
    
End Sub

Private Sub dato3_KeyPress(KeyAscii As Integer)
    snum = 0: KeyAscii = esNumero(KeyAscii)
    If KeyAscii = 13 Then
    Call ceros(dato3)
    LBLDV.Caption = rut(dato3)
    If leerProveedor(dato3.text + LBLDV.Caption) <> "" Then
    LBLPROVEEDOR.Caption = leerProveedor(dato3.text + LBLDV.Caption)
    Call Pregunta(dato3, dato4)
    Else
    MsgBox ("PROVEEDOR NO EXISTE")
    dato3.SetFocus
    
    End If
End If
End Sub
Private Sub dato4_KeyPress(KeyAscii As Integer)
    snum = 0: KeyAscii = esNumero(KeyAscii)
    If KeyAscii = 13 Then
    Call ceros(dato4)
    If leerempresa(dato4.text) <> "" Then
    LBLEMPRESA.Caption = leerempresa(dato4.text)
    DATO5.Enabled = True
    
    DATO5.SetFocus
    Else
    dato4.SetFocus
    End If
    
    
    End If
    
End Sub
Private Sub dato5_KeyPress(KeyAscii As Integer)
    snum = 0: KeyAscii = esNumero(KeyAscii)
    If KeyAscii = 13 Then
        If DATO5.text = "" Then DATO5.text = Day(fechasistema)
        If CDbl(DATO5.text) > 0 Then
            Call ceros(DATO5)
            Call Pregunta(DATO5, dato6)
        End If
    End If
    
End Sub
Private Sub dato6_KeyPress(KeyAscii As Integer)
    snum = 0: KeyAscii = esNumero(KeyAscii)
    If KeyAscii = 13 Then
        If dato6.text = "" Then dato6.text = Month(fechasistema)
        If CDbl(dato6.text) > 0 Then
            Call ceros(dato6)
            Call Pregunta(dato6, dato7)
        End If
    End If
End Sub
Private Sub dato7_KeyPress(KeyAscii As Integer)
     snum = 0: KeyAscii = esNumero(KeyAscii)
    If KeyAscii = 13 Then
        If dato7.text = "" Then dato7.text = Year(fechasistema)
        If CDbl(dato7.text) > 0 Then
            Call ceros(dato7)
            Call Pregunta(dato7, dato8)
        End If
    End If
End Sub
Private Sub dato8_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 And (dato8.text = "F" Or dato8.text = "B") Then Call Pregunta(dato8, dato9)
End Sub
Private Sub dato9_KeyPress(KeyAscii As Integer)
    snum = 0: KeyAscii = esNumero(KeyAscii)
    If KeyAscii = 13 And dato9.text <> "" Then
        If CDbl(dato9.text) > 0 Then
            Call ceros(dato9)
            Call Pregunta(dato9, dato10)
            If MODIFI = "0" Then
                Call cargadocumento(dato1.text, dato2.text, dato7.text & "-" & dato6.text & "-" & DATO5.text, dato9.text)
            End If
            
        End If
        cmdpagar.Visible = True
    End If
End Sub

Private Sub cargadocumento(tipo, numero, fecha, numerodocumento)
Dim csql As New rdoQuery
Dim resultados As rdoResultset

Set csql.ActiveConnection = contadb

csql.sql = "select tipo,numeroservicio,ifnull(fecha,'0'),tipodocumento,numerodocumento,montodocumento,ifnull(desde,'0'), "
csql.sql = csql.sql & "ifnull(hasta,'0'),ifnull(fechapago,'0'),tipocomprobante,numerocomprobante,lecturaactual, "
csql.sql = csql.sql & "lecturaanterior,consumo,saldoanterior,ifnull(fechacorte,'0'),tipopago,numerocheque "
csql.sql = csql.sql & "from " & clientesistema & "consumos_basicos.detalle_servicios "
csql.sql = csql.sql & "where tipo='" & tipo & "' and numeroservicio='" & numero & "' "
csql.sql = csql.sql & " and fecha='" & Format(fecha, "yyyy-mm-dd") & "' and numerodocumento='" & numerodocumento & "' "
csql.Execute

If csql.RowsAffected > 0 Then
    Set resultados = csql.OpenResultset
     
        dato1.text = resultados(0)
        dato2.text = resultados(1)
        DATO5.text = Mid(resultados(2), 9, 2)
        dato6.text = Mid(resultados(2), 6, 2)
        dato7.text = Mid(resultados(2), 1, 4)
        dato8.text = resultados(3)
        dato9.text = resultados(4)
        dato10.text = resultados(5)
        
        dato11.text = Mid(resultados(6), 9, 2)
        dato12.text = Mid(resultados(6), 6, 2)
        dato13.text = Mid(resultados(6), 1, 4)
        
        dato14.text = Mid(resultados(7), 9, 2)
        dato15.text = Mid(resultados(7), 6, 2)
        dato16.text = Mid(resultados(7), 1, 4)
        
        DATO20.text = resultados(9)
        DATO21.text = resultados(10)
        DATO22.text = resultados(11)
        dato23.text = resultados(12)
        dato24.text = resultados(13)
        dato28.text = resultados(14)
        
        DATO25.text = Mid(resultados(15), 9, 2)
        dato26.text = Mid(resultados(15), 6, 2)
        dato27.text = Mid(resultados(15), 1, 2)
        
        dato29.text = resultados(16)
        dato30.text = resultados(17)
        
        If resultados(8) <> "0000-00-00" Then
            dato17.text = Mid(resultados(8), 9, 2)
            dato18.text = Mid(resultados(8), 6, 2)
            dato19.text = Mid(resultados(8), 1, 2)
            Check1.Value = 1
        Else
            cmdpagar.Visible = True
        End If
        
        
    habilita (False)
    disponible (False)
    opciones.Visible = True
    opciones.SetFocus
    
        
End If

End Sub
Private Sub botonmisaccesos_Click()
programafiltro = Me.Caption
misaccesos.Show
End Sub


Private Sub botonmisfavoritos_Click()
Call AgregaFavoritos(USUARIOSISTEMA, App.EXEName, Me.Name, Me.Caption)
End Sub
