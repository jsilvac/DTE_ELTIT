VERSION 5.00
Object = "{ADBBAED6-B16F-40DC-80DF-B44910CBA76C}#1.0#0"; "Frame-Xp.ocx"
Begin VB.Form MCDatosFinancieros 
   BackColor       =   &H00FF8080&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   7965
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7305
   LinkTopic       =   "Form1"
   ScaleHeight     =   7965
   ScaleWidth      =   7305
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin XPFrame.FrameXp FrameXp3 
      Height          =   7815
      Left            =   60
      TabIndex        =   18
      Top             =   60
      Width           =   7155
      _ExtentX        =   12621
      _ExtentY        =   13785
      BackColor       =   16761024
      Caption         =   "Información Financiera"
      CaptionEstilo3D =   1
      BackColor       =   16761024
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
      Begin XPFrame.FrameXp frmInfoCredito 
         Height          =   2055
         Left            =   60
         TabIndex        =   46
         Top             =   5640
         Width           =   7035
         _ExtentX        =   12409
         _ExtentY        =   3625
         BackColor       =   8438015
         Caption         =   "Información del Crédito"
         CaptionEstilo3D =   1
         BackColor       =   8438015
         ColorBarraArriba=   12640511
         ColorBarraAbajo =   16512
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
         Begin VB.Label lbl23 
            Appearance      =   0  'Flat
            BackColor       =   &H000040C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " Autorizador"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   60
            TabIndex        =   60
            Top             =   1200
            Width           =   2715
         End
         Begin VB.Label lblAutorizador 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
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
            ForeColor       =   &H0080FFFF&
            Height          =   285
            Left            =   2940
            TabIndex        =   59
            Top             =   1200
            Width           =   4020
         End
         Begin VB.Label lblAñoTarjeta 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
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
            ForeColor       =   &H0080FFFF&
            Height          =   285
            Left            =   5040
            TabIndex        =   58
            Top             =   1620
            Width           =   780
         End
         Begin VB.Label lblMesTarjeta 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
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
            ForeColor       =   &H0080FFFF&
            Height          =   285
            Left            =   4620
            TabIndex        =   57
            Top             =   1620
            Width           =   360
         End
         Begin VB.Label lbl24 
            Appearance      =   0  'Flat
            BackColor       =   &H000040C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " F. Entrega Tarjeta"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   1320
            TabIndex        =   56
            Top             =   1620
            Width           =   2715
         End
         Begin VB.Label lblDiaTarjeta 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
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
            ForeColor       =   &H0080FFFF&
            Height          =   285
            Left            =   4200
            TabIndex        =   55
            Top             =   1620
            Width           =   360
         End
         Begin VB.Label lblAñoCredito 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
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
            ForeColor       =   &H0080FFFF&
            Height          =   285
            Left            =   3780
            TabIndex        =   54
            Top             =   840
            Width           =   780
         End
         Begin VB.Label lblMesCredito 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
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
            ForeColor       =   &H0080FFFF&
            Height          =   285
            Left            =   3360
            TabIndex        =   53
            Top             =   840
            Width           =   360
         End
         Begin VB.Label lbl22 
            Appearance      =   0  'Flat
            BackColor       =   &H000040C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " F. Autorización Crédito"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   60
            TabIndex        =   52
            Top             =   840
            Width           =   2715
         End
         Begin VB.Label lblDiaCredito 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
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
            ForeColor       =   &H0080FFFF&
            Height          =   285
            Left            =   2940
            TabIndex        =   51
            Top             =   840
            Width           =   360
         End
         Begin VB.Label lblAñoImpresion 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
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
            ForeColor       =   &H0080FFFF&
            Height          =   285
            Left            =   5040
            TabIndex        =   50
            Top             =   360
            Width           =   780
         End
         Begin VB.Label lblMesImpresion 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
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
            ForeColor       =   &H0080FFFF&
            Height          =   285
            Left            =   4620
            TabIndex        =   49
            Top             =   360
            Width           =   360
         End
         Begin VB.Label lbl21 
            Appearance      =   0  'Flat
            BackColor       =   &H000040C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " F. Impresión Pagaré"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   1320
            TabIndex        =   48
            Top             =   360
            Width           =   2715
         End
         Begin VB.Label lblDiaImpresion 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
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
            ForeColor       =   &H0080FFFF&
            Height          =   285
            Left            =   4200
            TabIndex        =   47
            Top             =   360
            Width           =   360
         End
      End
      Begin XPFrame.FrameXp frmCasasComerciales 
         Height          =   1815
         Left            =   60
         TabIndex        =   36
         Top             =   3780
         Width           =   7035
         _ExtentX        =   12409
         _ExtentY        =   3201
         BackColor       =   8950508
         Caption         =   "Casas Comerciales"
         CaptionEstilo3D =   1
         BackColor       =   8950508
         ColorBarraArriba=   12632319
         ColorBarraAbajo =   128
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
         Begin VB.TextBox dato18 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0FF&
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
            Left            =   5820
            MaxLength       =   9
            TabIndex        =   17
            Tag             =   "proveedor"
            Top             =   1440
            Width           =   1140
         End
         Begin VB.TextBox dato17 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0FF&
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
            TabIndex        =   16
            Tag             =   "proveedor"
            Top             =   1440
            Width           =   2520
         End
         Begin VB.TextBox dato16 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0FF&
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
            Left            =   5820
            MaxLength       =   9
            TabIndex        =   15
            Tag             =   "proveedor"
            Top             =   1080
            Width           =   1140
         End
         Begin VB.TextBox dato15 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0FF&
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
            TabIndex        =   14
            Tag             =   "proveedor"
            Top             =   1080
            Width           =   2520
         End
         Begin VB.TextBox dato14 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0FF&
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
            Left            =   5820
            MaxLength       =   9
            TabIndex        =   13
            Tag             =   "proveedor"
            Top             =   720
            Width           =   1140
         End
         Begin VB.TextBox dato13 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0FF&
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
            TabIndex        =   12
            Tag             =   "proveedor"
            Top             =   720
            Width           =   2520
         End
         Begin VB.TextBox dato12 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0FF&
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
            TabIndex        =   11
            Tag             =   "proveedor"
            Top             =   360
            Width           =   420
         End
         Begin VB.Label lbl20 
            Appearance      =   0  'Flat
            BackColor       =   &H00000080&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " Cupo3"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   4800
            TabIndex        =   45
            Top             =   1440
            Width           =   915
         End
         Begin VB.Label lbl19 
            Appearance      =   0  'Flat
            BackColor       =   &H00000080&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " Tarjeta3"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   60
            TabIndex        =   44
            Top             =   1440
            Width           =   1935
         End
         Begin VB.Label lbl18 
            Appearance      =   0  'Flat
            BackColor       =   &H00000080&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " Cupo2"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   4800
            TabIndex        =   43
            Top             =   1080
            Width           =   915
         End
         Begin VB.Label lbl17 
            Appearance      =   0  'Flat
            BackColor       =   &H00000080&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " Tarjeta2"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   60
            TabIndex        =   42
            Top             =   1080
            Width           =   1935
         End
         Begin VB.Label lbl16 
            Appearance      =   0  'Flat
            BackColor       =   &H00000080&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " Cupo1"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   4800
            TabIndex        =   41
            Top             =   720
            Width           =   915
         End
         Begin VB.Label lbl15 
            Appearance      =   0  'Flat
            BackColor       =   &H00000080&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " Tarjeta1"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   60
            TabIndex        =   40
            Top             =   720
            Width           =   1935
         End
         Begin VB.Label lblTotal 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
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
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   4800
            TabIndex        =   39
            Top             =   360
            Width           =   2160
         End
         Begin VB.Label lbl14 
            Appearance      =   0  'Flat
            BackColor       =   &H00000080&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " Cupo Total"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   2760
            TabIndex        =   38
            Top             =   360
            Width           =   1935
         End
         Begin VB.Label lbl13 
            Appearance      =   0  'Flat
            BackColor       =   &H00000080&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " Otras Tarjetas"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   60
            TabIndex        =   37
            Top             =   360
            Width           =   1935
         End
      End
      Begin XPFrame.FrameXp frmBanco 
         Height          =   1095
         Left            =   60
         TabIndex        =   29
         Top             =   2640
         Width           =   7035
         _ExtentX        =   12409
         _ExtentY        =   1931
         BackColor       =   7727290
         Caption         =   "Banco"
         CaptionEstilo3D =   1
         BackColor       =   7727290
         ColorBarraArriba=   12648384
         ColorBarraAbajo =   32768
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
         Begin VB.TextBox dato11 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFC0&
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
            MaxLength       =   2
            TabIndex        =   10
            Tag             =   "proveedor"
            Top             =   720
            Width           =   480
         End
         Begin VB.TextBox dato10 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFC0&
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
            MaxLength       =   11
            TabIndex        =   9
            Tag             =   "proveedor"
            Top             =   720
            Width           =   1320
         End
         Begin VB.TextBox dato9 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFC0&
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
            Left            =   3660
            MaxLength       =   3
            TabIndex        =   8
            Tag             =   "proveedor"
            Top             =   360
            Width           =   540
         End
         Begin VB.TextBox dato8 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFC0&
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
            Tag             =   "proveedor"
            Top             =   360
            Width           =   420
         End
         Begin VB.Label lbl12 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   " Años"
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
            Height          =   285
            Left            =   6180
            TabIndex        =   35
            Top             =   720
            Width           =   780
         End
         Begin VB.Label lbl11 
            Appearance      =   0  'Flat
            BackColor       =   &H00008080&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Antigüedad"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   3600
            TabIndex        =   34
            Top             =   720
            Width           =   1935
         End
         Begin VB.Label lbl10 
            Appearance      =   0  'Flat
            BackColor       =   &H00008080&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " N° Cuenta"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   60
            TabIndex        =   33
            Top             =   720
            Width           =   1935
         End
         Begin VB.Label lblBanco 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
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
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   4260
            TabIndex        =   32
            Top             =   360
            Width           =   2700
         End
         Begin VB.Label lbl9 
            Appearance      =   0  'Flat
            BackColor       =   &H00008080&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " Banco"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   2700
            TabIndex        =   31
            Top             =   360
            Width           =   915
         End
         Begin VB.Label lbl8 
            Appearance      =   0  'Flat
            BackColor       =   &H00008080&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " Cuenta Corriente"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   60
            TabIndex        =   30
            Top             =   360
            Width           =   1935
         End
      End
      Begin XPFrame.FrameXp frmIngresoseGRESOS 
         Height          =   2175
         Left            =   60
         TabIndex        =   20
         Top             =   420
         Width           =   7035
         _ExtentX        =   12409
         _ExtentY        =   3836
         BackColor       =   11070967
         Caption         =   "Ingresos y Egresos"
         CaptionEstilo3D =   1
         BackColor       =   11070967
         ColorBarraArriba=   12648447
         ColorBarraAbajo =   32896
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
            Left            =   2160
            MaxLength       =   9
            TabIndex        =   6
            Tag             =   "proveedor"
            Top             =   1800
            Width           =   1320
         End
         Begin VB.TextBox dato5 
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
            Left            =   5640
            MaxLength       =   9
            TabIndex        =   4
            Tag             =   "proveedor"
            Top             =   1080
            Width           =   1320
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
            Left            =   5640
            MaxLength       =   9
            TabIndex        =   1
            Tag             =   "proveedor"
            Top             =   360
            Width           =   1320
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
            Left            =   2160
            MaxLength       =   50
            TabIndex        =   5
            Tag             =   "proveedor"
            Top             =   1440
            Width           =   4800
         End
         Begin VB.TextBox dato4 
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
            Left            =   2160
            MaxLength       =   9
            TabIndex        =   3
            Tag             =   "proveedor"
            Top             =   1080
            Width           =   1320
         End
         Begin VB.TextBox dato3 
            Alignment       =   2  'Center
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
            Left            =   2160
            MaxLength       =   1
            TabIndex        =   2
            Tag             =   "proveedor"
            Top             =   720
            Width           =   405
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
            Left            =   2160
            MaxLength       =   9
            TabIndex        =   0
            Tag             =   "proveedor"
            Top             =   360
            Width           =   1320
         End
         Begin VB.Label lbl7 
            Appearance      =   0  'Flat
            BackColor       =   &H00008000&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " Tas. Vehiculos"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   60
            TabIndex        =   28
            Top             =   1800
            Width           =   1935
         End
         Begin VB.Label lbl5 
            Appearance      =   0  'Flat
            BackColor       =   &H00004000&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " Arriendo"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   3600
            TabIndex        =   27
            Top             =   1080
            Width           =   1935
         End
         Begin VB.Label lbl2 
            Appearance      =   0  'Flat
            BackColor       =   &H00004000&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " P. Casas Comer."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   3600
            TabIndex        =   26
            Top             =   360
            Width           =   1935
         End
         Begin VB.Label lbl6 
            Appearance      =   0  'Flat
            BackColor       =   &H00008000&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " Vehiculos"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   60
            TabIndex        =   25
            Top             =   1440
            Width           =   1935
         End
         Begin VB.Label lbl4 
            Appearance      =   0  'Flat
            BackColor       =   &H00008000&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " Tas. Vivienda"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   60
            TabIndex        =   24
            Top             =   1080
            Width           =   1935
         End
         Begin VB.Label lblVivienda 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
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
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   2625
            TabIndex        =   23
            Top             =   720
            Width           =   4335
         End
         Begin VB.Label lbl3 
            Appearance      =   0  'Flat
            BackColor       =   &H00008000&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " Tipo Vivienda"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   60
            TabIndex        =   22
            Top             =   720
            Width           =   1935
         End
         Begin VB.Label lbl1 
            Appearance      =   0  'Flat
            BackColor       =   &H00008000&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " Ingreso Mensual"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   60
            TabIndex        =   21
            Top             =   360
            Width           =   1935
         End
      End
      Begin XPFrame.FrameXp frmCerrar 
         Height          =   285
         Left            =   6840
         TabIndex        =   19
         Top             =   30
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   503
         BackColor       =   49344
         Caption         =   "X"
         CaptionEstilo3D =   1
         BackColor       =   49344
         ColorBarraArriba=   8388608
         ColorBarraAbajo =   16761024
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
      End
   End
End
Attribute VB_Name = "MCDatosFinancieros"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Private cf As financieros
    Private modificar As Boolean
    
'==============================================================
'GOTFOCUS
'==============================================================
    Private Sub dato1_GotFocus()
        Call VerificarCajas(Me, dato1)
        Call selecciona(dato1)
    End Sub
    
    Private Sub dato2_GotFocus()
        Call VerificarCajas(Me, dato2)
        Call selecciona(dato2)
    End Sub
    
    Private Sub dato3_GotFocus()
        Call VerificarCajas(Me, dato3)
        Call selecciona(dato3)
        Principal.barraEstado.Panels(2).text = "F2: Ayuda Tipo Vivienda "
    End Sub
    


Private Sub dato3_LostFocus()
limpiaBarra (2)
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
    
    Private Sub dato18_GotFocus()
        Call VerificarCajas(Me, dato18)
        Call selecciona(dato18)
    End Sub
'==============================================================
'GOTFOCUS
'==============================================================

'==============================================================
'KEYDOWN
'==============================================================
    Private Sub dato1_KeyDown(KeyCode As Integer, Shift As Integer)
        Call Flechas(KeyCode, dato1)
    End Sub
    
    Private Sub dato2_KeyDown(KeyCode As Integer, Shift As Integer)
        Call Flechas(KeyCode, dato1)
    End Sub
    
    Private Sub dato3_KeyDown(KeyCode As Integer, Shift As Integer)
      If KeyCode = vbKeyF2 Then
            Call ayudavivienda(dato3)
        Else
            Call Flechas(KeyCode, dato2)
        End If
        
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
        Call Flechas(KeyCode, dato15)
    End Sub
    
    Private Sub dato17_KeyDown(KeyCode As Integer, Shift As Integer)
        Call Flechas(KeyCode, dato16)
    End Sub
    
    Private Sub dato18_KeyDown(KeyCode As Integer, Shift As Integer)
        Call Flechas(KeyCode, dato17)
    End Sub
'==============================================================
'KEYDOWN
'==============================================================

'==============================================================
'KEYPRESS
'==============================================================
    Private Sub dato1_KeyPress(KeyAscii As Integer)
        KeyAscii = esNumero(KeyAscii)
        If KeyAscii = 13 Then
            SendKeys "{Tab}"
        End If
    End Sub
    
    Private Sub dato2_KeyPress(KeyAscii As Integer)
        KeyAscii = esNumero(KeyAscii)
        If KeyAscii = 13 Then
            SendKeys "{Tab}"
        End If
    End Sub
    
    Private Sub dato3_KeyPress(KeyAscii As Integer)
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If KeyAscii = 13 And dato3.text <> "" Then
            lblVivienda.Caption = leerTipoVivienda(dato3.text)
            If lblVivienda.Caption <> "" Then
                SendKeys "{Tab}"
            End If
        End If
    End Sub
    
    Private Sub dato4_KeyPress(KeyAscii As Integer)
        KeyAscii = esNumero(KeyAscii)
        If KeyAscii = 13 Then
            SendKeys "{Tab}"
        End If
    End Sub
    
    Private Sub dato5_KeyPress(KeyAscii As Integer)
        KeyAscii = esNumero(KeyAscii)
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
        KeyAscii = esNumero(KeyAscii)
        If KeyAscii = 13 Then
            SendKeys "{Tab}"
        End If
    End Sub
    
    Private Sub dato8_KeyPress(KeyAscii As Integer)
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If Chr(KeyAscii) <> "S" And Chr(KeyAscii) <> "N" Then
            If KeyAscii = 13 And LTrim(dato8.text) <> "" Then
                Select Case dato8.text
                    Case "S"
                        dato9.Enabled = True
                        dato10.Enabled = True
                        dato11.Enabled = True
                    Case "N"
                        dato9.Enabled = False
                        dato10.Enabled = False
                        dato11.Enabled = False
                End Select
                SendKeys "{Tab}"
            Else
                KeyAscii = 0
            End If
        End If
    End Sub
    
    Private Sub dato9_KeyPress(KeyAscii As Integer)
        KeyAscii = esNumero(KeyAscii)
        If KeyAscii = 13 Then
            lblBanco.Caption = leerNombreBanco(dato9.text)
            If lblBanco.Caption <> "" Then
                SendKeys "{Tab}"
            End If
        End If
    End Sub
    
    Private Sub dato10_KeyPress(KeyAscii As Integer)
        KeyAscii = esNumero(KeyAscii)
        If KeyAscii = 13 Then
            SendKeys "{Tab}"
        End If
    End Sub
    
    Private Sub dato11_KeyPress(KeyAscii As Integer)
        KeyAscii = esNumero(KeyAscii)
        If KeyAscii = 13 Then
            SendKeys "{Tab}"
        End If
    End Sub
    
    Private Sub dato12_KeyPress(KeyAscii As Integer)
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If Chr(KeyAscii) <> "S" And Chr(KeyAscii) <> "N" Then
            If KeyAscii = 13 And LTrim(dato12.text) <> "" Then
                Select Case dato12.text
                    Case "S"
                        dato13.Enabled = True
                        dato14.Enabled = True
                        dato15.Enabled = True
                        dato16.Enabled = True
                        dato17.Enabled = True
                        dato18.Enabled = True
                    Case "N"
                        dato13.Enabled = False
                        dato14.Enabled = False
                        dato15.Enabled = False
                        dato16.Enabled = False
                        dato17.Enabled = False
                        dato18.Enabled = False
                        Call ctrltostruct
                End Select
                SendKeys "{Tab}"
            Else
                KeyAscii = 0
            End If
        End If
    End Sub
    
    Private Sub dato13_KeyPress(KeyAscii As Integer)
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If KeyAscii = 13 Then
            SendKeys "{Tab}"
        End If
    End Sub
    
    Private Sub dato14_KeyPress(KeyAscii As Integer)
        KeyAscii = esNumero(KeyAscii)
        If KeyAscii = 13 Then
            If dato14.text = "" Then
                dato14.text = "0"
            End If
            If dato16.text = "" Then
                dato16.text = "0"
            End If
            If dato18.text = "" Then
                dato18.text = "0"
            End If
            lblTotal.Caption = Format(CDbl(dato14.text) + CDbl(dato16.text) + CDbl(dato18.text), "$ ###,###,##0")
            SendKeys "{Tab}"
        End If
    End Sub
    
    Private Sub dato15_KeyPress(KeyAscii As Integer)
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If KeyAscii = 13 Then
            SendKeys "{Tab}"
        End If
    End Sub
    
    Private Sub dato16_KeyPress(KeyAscii As Integer)
        KeyAscii = esNumero(KeyAscii)
        If KeyAscii = 13 Then
            If dato14.text = "" Then
                dato14.text = "0"
            End If
            If dato16.text = "" Then
                dato16.text = "0"
            End If
            If dato18.text = "" Then
                dato18.text = "0"
            End If
            lblTotal.Caption = Format(CDbl(dato14.text) + CDbl(dato16.text) + CDbl(dato18.text), "$ ###,###,##0")
            SendKeys "{Tab}"
        End If
    End Sub
    
    Private Sub dato17_KeyPress(KeyAscii As Integer)
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If KeyAscii = 13 Then
            SendKeys "{Tab}"
        End If
    End Sub
    
    Private Sub dato18_KeyPress(KeyAscii As Integer)
        KeyAscii = esNumero(KeyAscii)
        If KeyAscii = 13 Then
            If dato14.text = "" Then
                dato14.text = "0"
            End If
            If dato16.text = "" Then
                dato16.text = "0"
            End If
            If dato18.text = "" Then
                dato18.text = "0"
            End If
            lblTotal.Caption = Format(CDbl(dato14.text) + CDbl(dato16.text) + CDbl(dato18.text), "$ ###,###,##0")
            Call ctrltostruct
        End If
    End Sub
'==============================================================
'KEYPRESS
'==============================================================

    Private Sub Form_Activate()
        modificar = False
        If dato1.text <> "" Then
            modificar = True
        End If
    End Sub

    Private Sub Form_Load()
        cf.rut = MClientes.dato1.text & MClientes.lblDV.Caption
        cf.sucursal = MClientes.dato2.text
    End Sub

    Private Sub frmCerrar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
        Call cambiaColor(frmCerrar)
        frmCerrar.CaptionEstilo3D = Raised
    End Sub

    Private Sub frmCerrar_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
        Call cambiaColor(frmCerrar)
        frmCerrar.CaptionEstilo3D = Inserted
        Unload Me
    End Sub

'==============================================================
'PASA LOS DATOS DE LOS CONTROLES A LA ESTRUCTURA
'==============================================================
    Private Sub ctrltostruct()
        cf.ingresomensual = dato1.text
        cf.pagoscasascomerciales = dato2.text
        cf.tipovivienda = dato3.text
        cf.tasacionvivienda = dato4.text
        cf.arriendo = dato5.text
        cf.vehiculos = dato6.text
        cf.tasacionvehiculos = dato7.text
        cf.cuentacorriente = dato8.text
        cf.Banco = dato9.text
        cf.numerocuenta = dato10.text
        cf.antuguedad = dato11.text
        cf.otrastarjetas = dato12.text
        cf.otratarjeta1 = dato13.text
        cf.otratarjetacupo1 = dato14.text
        cf.otratarjeta2 = dato15.text
        cf.otratarjetacupo2 = dato16.text
        cf.otratarjeta3 = dato17.text
        cf.otratarjetacupo3 = dato18.text
        
        Call grabarClienteFinancieros(cf, modificar)
        Unload Me
    End Sub
'==============================================================
'PASA LOS DATOS DE LOS CONTROLES A LA ESTRUCTURA
'==============================================================








