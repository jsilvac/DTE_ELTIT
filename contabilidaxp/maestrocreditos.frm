VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash.ocx"
Object = "{4777436C-EB8C-4596-98A8-EBCF98683969}#1.0#0"; "FlexCell.ocx"
Object = "{ADBBAED6-B16F-40DC-80DF-B44910CBA76C}#1.0#0"; "Frame-Xp.ocx"
Begin VB.Form prestamo02 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Contratos De Arriendo De Propiedades"
   ClientHeight    =   10275
   ClientLeft      =   2040
   ClientTop       =   1425
   ClientWidth     =   15240
   DrawStyle       =   5  'Transparent
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   685
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1016
   StartUpPosition =   2  'CenterScreen
   Begin XPFrame.FrameXp FrameQuickMenu 
      Height          =   615
      Left            =   11640
      TabIndex        =   55
      Top             =   9480
      Width           =   3255
      _ExtentX        =   5741
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
      Alignment       =   1
      Begin VB.CommandButton botonmisaccesos 
         Caption         =   "Permisos Modulo"
         Height          =   255
         Left            =   120
         TabIndex        =   57
         Top             =   280
         Width           =   1455
      End
      Begin VB.CommandButton botonmisfavoritos 
         Caption         =   "Mis Favoritos"
         Height          =   255
         Left            =   1800
         TabIndex        =   56
         Top             =   280
         Width           =   1335
      End
   End
   Begin XPFrame.FrameXp FrameXp2 
      Height          =   5055
      Left            =   8865
      TabIndex        =   9
      Top             =   0
      Width           =   6270
      _ExtentX        =   11060
      _ExtentY        =   8916
      BackColor       =   16761024
      Caption         =   "Detalle de Arriendos"
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
      Begin VB.CheckBox Check1 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Modifica Cuotas"
         Height          =   330
         Left            =   180
         TabIndex        =   54
         Top             =   4365
         Width           =   1500
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Pagar"
         Height          =   330
         Left            =   1890
         TabIndex        =   15
         Top             =   4635
         Visible         =   0   'False
         Width           =   1410
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Todos"
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
         Left            =   3645
         TabIndex        =   14
         Top             =   4635
         Width           =   1860
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Pendientes"
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
         Left            =   3645
         TabIndex        =   13
         Top             =   4275
         Value           =   -1  'True
         Width           =   1860
      End
      Begin FlexCell.Grid Grid1 
         Height          =   3930
         Left            =   45
         TabIndex        =   10
         Top             =   270
         Width           =   6090
         _ExtentX        =   10742
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
      TabIndex        =   8
      Top             =   10275
      Width           =   15240
   End
   Begin XPFrame.FrameXp FrameXp3 
      Height          =   4110
      Left            =   45
      TabIndex        =   11
      Top             =   5040
      Width           =   14910
      _ExtentX        =   26300
      _ExtentY        =   7250
      BackColor       =   16761024
      Caption         =   "Prestamos Vigentes"
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
      Begin FlexCell.Grid Grid2 
         Height          =   3705
         Left            =   90
         TabIndex        =   12
         Top             =   270
         Width           =   14775
         _ExtentX        =   26061
         _ExtentY        =   6535
         Cols            =   5
         DefaultFontSize =   8.25
         Rows            =   30
      End
   End
   Begin XPFrame.FrameXp FrameXp1 
      Height          =   4830
      Left            =   0
      TabIndex        =   16
      Top             =   0
      Width           =   8835
      _ExtentX        =   15584
      _ExtentY        =   8520
      BackColor       =   16744576
      Caption         =   "Numero Credito"
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
         MaxLength       =   10
         TabIndex        =   2
         Top             =   945
         Width           =   1455
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
         MaxLength       =   3
         TabIndex        =   0
         Tag             =   "fecha"
         Top             =   315
         Width           =   555
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
         MaxLength       =   2
         TabIndex        =   1
         Tag             =   "fecha"
         Top             =   630
         Width           =   375
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
         Top             =   1260
         Width           =   375
      End
      Begin XPFrame.FrameXp FrameXp4 
         Height          =   3165
         Left            =   90
         TabIndex        =   17
         Top             =   1620
         Width           =   8610
         _ExtentX        =   15187
         _ExtentY        =   5583
         BackColor       =   16761024
         Caption         =   "Comprobante Credito  Bancario"
         CaptionEstilo3D =   1
         BackColor       =   16761024
         ForeColor       =   14737632
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Begin VB.TextBox DATO16 
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
            Left            =   2610
            MaxLength       =   20
            TabIndex        =   47
            Top             =   2610
            Width           =   5775
         End
         Begin VB.TextBox DATO14 
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
            Left            =   2610
            MaxLength       =   20
            TabIndex        =   45
            Top             =   1980
            Width           =   1455
         End
         Begin VB.TextBox DATO9 
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
            Left            =   2610
            MaxLength       =   2
            TabIndex        =   36
            Top             =   945
            Width           =   465
         End
         Begin VB.TextBox DATO8 
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
            Left            =   2610
            MaxLength       =   20
            TabIndex        =   34
            Top             =   630
            Width           =   1455
         End
         Begin VB.TextBox dato5 
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
            Left            =   2610
            MaxLength       =   2
            TabIndex        =   4
            Tag             =   "fecha"
            Top             =   315
            Width           =   375
         End
         Begin VB.TextBox dato6 
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
            Left            =   3015
            MaxLength       =   2
            TabIndex        =   5
            Tag             =   "fecha"
            Top             =   315
            Width           =   375
         End
         Begin VB.TextBox dato7 
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
            Left            =   3465
            MaxLength       =   4
            TabIndex        =   6
            Tag             =   "fecha"
            Top             =   315
            Width           =   645
         End
         Begin VB.TextBox dato10 
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
            Left            =   2610
            MaxLength       =   2
            TabIndex        =   22
            Tag             =   "fecha"
            Top             =   1260
            Width           =   375
         End
         Begin VB.TextBox dato11 
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
            Left            =   3015
            MaxLength       =   2
            TabIndex        =   21
            Tag             =   "fecha"
            Top             =   1260
            Width           =   375
         End
         Begin VB.TextBox dato12 
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
            Left            =   3465
            MaxLength       =   4
            TabIndex        =   20
            Tag             =   "fecha"
            Top             =   1260
            Width           =   645
         End
         Begin VB.TextBox DATO13 
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
            Left            =   2610
            MaxLength       =   2
            TabIndex        =   19
            Tag             =   "fecha"
            Top             =   1620
            Width           =   510
         End
         Begin VB.TextBox dato15 
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
            Left            =   2610
            MaxLength       =   20
            TabIndex        =   18
            Top             =   2295
            Width           =   1455
         End
         Begin XPFrame.FrameXp FrameXp5 
            Height          =   2175
            Left            =   4500
            TabIndex        =   38
            Top             =   360
            Width           =   3975
            _ExtentX        =   7011
            _ExtentY        =   3836
            BackColor       =   16744576
            Caption         =   "RESUMEN DEL CREDITO"
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
            Begin VB.TextBox SALDO 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00000000&
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
               ForeColor       =   &H0000FFFF&
               Height          =   285
               Left            =   2160
               Locked          =   -1  'True
               MaxLength       =   20
               TabIndex        =   51
               Top             =   1755
               Width           =   1455
            End
            Begin VB.TextBox CANCELADO 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00000000&
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
               ForeColor       =   &H0000FFFF&
               Height          =   285
               Left            =   2160
               Locked          =   -1  'True
               MaxLength       =   20
               TabIndex        =   49
               Top             =   1440
               Width           =   1455
            End
            Begin VB.TextBox CAPITAL 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00000000&
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
               ForeColor       =   &H0000FFFF&
               Height          =   285
               Left            =   2160
               Locked          =   -1  'True
               MaxLength       =   20
               TabIndex        =   41
               Top             =   360
               Width           =   1455
            End
            Begin VB.TextBox PROYECTADO 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00000000&
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
               ForeColor       =   &H0000FFFF&
               Height          =   285
               Left            =   2160
               Locked          =   -1  'True
               MaxLength       =   20
               TabIndex        =   40
               Top             =   720
               Width           =   1455
            End
            Begin VB.TextBox SOBRE 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00000000&
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
               ForeColor       =   &H0000FFFF&
               Height          =   285
               Left            =   2160
               Locked          =   -1  'True
               MaxLength       =   20
               TabIndex        =   39
               Top             =   1080
               Width           =   1455
            End
            Begin VB.Label Label14 
               BackStyle       =   0  'Transparent
               BorderStyle     =   1  'Fixed Single
               Caption         =   " SALDO CREDITO"
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
               Left            =   45
               TabIndex        =   52
               Top             =   1755
               Width           =   2085
            End
            Begin VB.Label Label13 
               BackStyle       =   0  'Transparent
               BorderStyle     =   1  'Fixed Single
               Caption         =   " CANCELADO"
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
               Left            =   45
               TabIndex        =   50
               Top             =   1440
               Width           =   2085
            End
            Begin VB.Label Label17 
               BackStyle       =   0  'Transparent
               BorderStyle     =   1  'Fixed Single
               Caption         =   " CAPITAL"
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
               Left            =   45
               TabIndex        =   44
               Top             =   360
               Width           =   2085
            End
            Begin VB.Label Label18 
               BackStyle       =   0  'Transparent
               BorderStyle     =   1  'Fixed Single
               Caption         =   " CREDITO FINAL"
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
               Left            =   45
               TabIndex        =   43
               Top             =   720
               Width           =   2085
            End
            Begin VB.Label Label19 
               BackStyle       =   0  'Transparent
               BorderStyle     =   1  'Fixed Single
               Caption         =   " INTERESES"
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
               Left            =   45
               TabIndex        =   42
               Top             =   1080
               Width           =   2085
            End
         End
         Begin VB.Label LBLMONEDA 
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
            Height          =   285
            Left            =   3150
            TabIndex        =   53
            Top             =   945
            Width           =   1275
         End
         Begin VB.Label Label12 
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   " GLOSA"
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
            Left            =   45
            TabIndex        =   48
            Top             =   2610
            Width           =   2445
         End
         Begin VB.Label Label11 
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   " CANTIDAD CUOTAS"
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
            Left            =   45
            TabIndex        =   46
            Top             =   1980
            Width           =   2445
         End
         Begin VB.Label Label10 
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   " MONEDA"
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
            Left            =   45
            TabIndex        =   37
            Top             =   945
            Width           =   2445
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   " CAPITAL"
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
            Left            =   45
            TabIndex        =   35
            Top             =   630
            Width           =   2445
         End
         Begin VB.Label Label9 
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   " PRIMER VENCIMIENTO"
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
            Left            =   45
            TabIndex        =   26
            Top             =   1260
            Width           =   2445
         End
         Begin VB.Label Label8 
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
            Left            =   45
            TabIndex        =   25
            Top             =   315
            Width           =   2445
         End
         Begin VB.Label Label7 
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   " MONTO CUOTA"
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
            Left            =   45
            TabIndex        =   24
            Top             =   2295
            Width           =   2445
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   " DIAS TOLERANCIA"
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
            Left            =   45
            TabIndex        =   23
            Top             =   1620
            Width           =   2445
         End
      End
      Begin VB.Label Label1 
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
         Left            =   180
         TabIndex        =   33
         Top             =   630
         Width           =   1815
      End
      Begin VB.Label Label3 
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
         Left            =   180
         TabIndex        =   32
         Top             =   945
         Width           =   1815
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
         TabIndex        =   31
         Top             =   1260
         Width           =   1815
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   " BANCO"
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
         Top             =   315
         Width           =   1815
      End
      Begin VB.Label LBLBANCO 
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
         TabIndex        =   29
         Top             =   315
         Width           =   4785
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
         TabIndex        =   28
         Top             =   630
         Width           =   4785
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
         Left            =   2835
         TabIndex        =   27
         Top             =   1260
         Width           =   4785
      End
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash opciones 
      Height          =   1215
      Left            =   630
      TabIndex        =   7
      Top             =   9000
      Width           =   6735
      _cx             =   11880
      _cy             =   2143
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
Attribute VB_Name = "prestamo02"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public saldoglobal As Double
Private moneda As String
Private rutpropi As String

Private MODIFI As Integer


Public Sub grabar_meses()

Dim mesin As Double
Dim añoin As Double
Dim mesfin As Double
Dim añofin As Double
Dim dia As String
Dim FINAL As Boolean
FINAL = False

dia = dato10.text
añoin = CDbl(dato12.text)
mesin = CDbl(dato11.text)

For k = 1 To CDbl(dato14.text)

Call grabarmeses(dato1.text, dato2.text, dato3.text, dato4.text, dato15.text, Format(añoin, "0000") & "-" & Format(mesin, "00") & "-" & dia, k)
mesin = mesin + 1
If mesin = 13 Then añoin = añoin + 1: mesin = 1

Next k

Call leermensualidades(dato1.text, dato2.text, dato3.text, dato4.text)

End Sub

Private Sub Command1_Click()
grabar_meses
End Sub

Private Sub Check1_Click()
If Verifica_Permiso(Me.Caption, "autoriza") = False Then
Check1.Value = "0"
Else
Grid1.Column(2).Locked = False

End If
If Check1.Value = "0" And Grid1.Column(2).Locked = False Then
Grid1.Column(2).Locked = True


End If


End Sub

Private Sub COMMAND2_Click()

arriendo06.Show

End Sub


Private Sub dato5_GotFocus()
leer

End Sub

Private Sub Grid1_DblClick()
Dim pagado As String

If Grid1.ActiveCell.col = 3 Then
If Grid1.Cell(Grid1.ActiveCell.row, 3).text = "1" Then
pagado = "0"
Else
pagado = "1"
End If

Call modificapago(dato1.text, dato2.text, dato3.text, dato4.text, Format(Grid1.Cell(Grid1.ActiveCell.row, 1).text, "yyyy-mm-dd"), pagado)
Grid1.Cell(Grid1.ActiveCell.row, 3).text = pagado
End If




End Sub

Private Sub Grid1_LeaveCell(ByVal row As Long, ByVal col As Long, NewRow As Long, NewCol As Long, Cancel As Boolean)
If col = 2 And Check1.Value = "1" Then
Call modificamonto(dato1.text, dato2.text, dato3.text, dato4.text, Format(Grid1.Cell(Grid1.ActiveCell.row, 1).text, "yyyy-mm-dd"), Grid1.Cell(Grid1.ActiveCell.row, 2).text)
End If

End Sub

Private Sub Grid2_DblClick()
dato1.text = Grid2.Cell(Grid2.ActiveCell.row, 1).text
dato2.text = Grid2.Cell(Grid2.ActiveCell.row, 2).text
dato3.text = Grid2.Cell(Grid2.ActiveCell.row, 3).text
dato4.text = Grid2.Cell(Grid2.ActiveCell.row, 4).text
DATO5.SetFocus

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
    campos(0, 0) = "banco"
    campos(1, 0) = "tipo"
    campos(2, 0) = "numero"
    campos(3, 0) = "empresa"
    campos(4, 0) = "fecha"
    campos(5, 0) = "capital"
    campos(6, 0) = "moneda"
    campos(7, 0) = "impuestos"
    campos(8, 0) = "cuotas"
    campos(9, 0) = "primervencimiento"
    campos(10, 0) = "monto"
    campos(11, 0) = "diapago"
    campos(12, 0) = ""
    campos(0, 2) = clientesistema & "creditos_bancarios" & ".maestro_compromisos"
    
    
    condicion = "banco='" + dato1.text + "' and tipo='" + dato2.text + "' and numero ='" & dato3.text & "' and empresa='" + dato4.text + "' "
    
    op = 5
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)
    If sqlconta.status = 4 Then DATO5.SetFocus: GoTo no:
    carga
    opciones.Visible = True
    disponible (True)
    habilita (True)
    opciones.SetFocus
        
no:
End Sub

Sub carga()
    Dim TOTA1 As Double
    Dim TOTA2 As Double
    Dim TOTA3 As Double
    Dim TOTA4 As Double
    Dim TOTA5 As Double
    

    habilita (True)
    dato1.text = sqlconta.response(0, 3)
    dato2.text = sqlconta.response(1, 3)
    dato3.text = sqlconta.response(2, 3)
    dato4.text = sqlconta.response(3, 3)
    
    DATO5.text = Mid(sqlconta.response(4, 3), 1, 2)
    dato6.text = Mid(sqlconta.response(4, 3), 4, 2)
    dato7.text = Mid(sqlconta.response(4, 3), 7, 4)
    dato8.text = Format(sqlconta.response(5, 3), "###,###,###,##0")
    dato9.text = sqlconta.response(6, 3)
    dato10.text = Mid(sqlconta.response(9, 3), 1, 2)
    dato11.text = Mid(sqlconta.response(9, 3), 4, 2)
    dato12.text = Mid(sqlconta.response(9, 3), 7, 4)
    dato13.text = sqlconta.response(11, 3)
    dato14.text = Format(sqlconta.response(8, 3), "###,###,###,##0")
    dato15.text = Format(sqlconta.response(10, 3), "###,###,###,##0.00")
    TOTA1 = CDbl(dato8.text)
    TOTA2 = CDbl(dato14.text) * CDbl(dato15.text)
    TOTA3 = TOTA2 - TOTA1
    TOTA4 = leerpagado(dato1.text, dato2.text, dato3.text, dato4.text)
    
    
    TOTA5 = TOTA2 - TOTA4
    CAPITAL.text = Format(TOTA1, "###,###,###,##0.00")
    PROYECTADO.text = Format(TOTA2, "###,###,###,##0.00")
    SOBRE.text = Format(TOTA3, "###,###,###,##0.00")
    CANCELADO.text = Format(TOTA4, "###,###,###,##0.00")
    saldo.text = Format(TOTA5, "###,###,###,##0.00")
    lblBanco.Caption = leerbanco(dato1.text)
    LBLTIPO.Caption = leertipocredito(dato2.text)
    LBLEMPRESA.Caption = leerempresa(dato4.text)
    LBLMONEDA.Caption = leertipoMONEDA(dato9.text)

    Call leermensualidades(dato1.text, dato2.text, dato3.text, dato4.text)
Call leerCREDITOS
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
    campos(0, 0) = "banco"
    campos(1, 0) = "tipo"
    campos(2, 0) = "numero"
    campos(3, 0) = "empresa"
    campos(4, 0) = "fecha"
    campos(5, 0) = "capital"
    campos(6, 0) = "moneda"
    campos(7, 0) = "impuestos"
    campos(8, 0) = "cuotas"
    campos(9, 0) = "primervencimiento"
    campos(10, 0) = "monto"
    campos(11, 0) = "diapago"
    campos(12, 0) = "glosa"
    campos(13, 0) = ""
    campos(0, 1) = dato1.text
    campos(1, 1) = dato2.text
    campos(2, 1) = dato3.text
    campos(3, 1) = dato4.text
    campos(4, 1) = dato7.text & "-" & dato6.text & "-" & DATO5.text
    campos(5, 1) = dato8.text
    campos(6, 1) = dato9.text
    campos(7, 1) = "0"
    campos(8, 1) = dato14.text
    campos(9, 1) = dato12.text + "-" + dato11.text + "-" + dato10.text
    campos(10, 1) = Replace(dato15.text, ",", ".")
    campos(11, 1) = dato13.text
    campos(12, 1) = dato16.text
    campos(0, 2) = clientesistema & "creditos_bancarios" & ".maestro_compromisos"
    If MODIFI = 1 Then
    condicion = "banco='" + dato1.text + "' and tipo='" + dato2.text + "' and numero ='" & dato3.text & "' and empresa='" + dato4.text + "' "
    End If
    
    If MODIFI = 1 Then op = 3 Else op = 2
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)
    grabar_meses
    
    End Sub
 
Sub ELIMINAR()
    campos(0, 2) = clientesistema & "creditos_bancarios" & ".maestro_compromisos"
    condicion = "banco='" + dato1.text + "' and tipo='" + dato2.text + "' and numero ='" & dato3.text & "' and empresa='" + dato4.text + "' "
    
    op = 4
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)

    campos(0, 2) = clientesistema & "creditos_bancarios" & ".creditos_vencimientos"
    condicion = "banco='" + dato1.text + "' and tipo='" + dato2.text + "' and numero ='" & dato3.text & "' and empresa='" + dato4.text + "' "
    
    op = 4
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)


End Sub
  

Private Sub opciones_FSCommand(ByVal command As String, ByVal args As String)

If command = "retorno" Then retorno
If command = "modifica" Then
    MsgBox ("IMPOSIBLE MODIFICAR CREDITOS INGRESADOS ")
    
End If

If command = "elimina" Then
    If CDbl(CANCELADO.text) = 0 Then
    If Verifica_Permiso(Me.Caption, "elimina") = True Then
        ELIMINA
    End If
    Else
    MsgBox ("NO PUEDE ELIMINAR CREDITOS CON PAGOS ")
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
    formatogrilla2(1, 2) = "MONTO"
    formatogrilla2(1, 3) = "PAGADO"
    formatogrilla2(1, 4) = "CUOTA"
    formatogrilla2(1, 5) = "PESOS"
    
    
    Rem LARGO DE LOS DATOS
    formatogrilla2(2, 1) = "8"
    formatogrilla2(2, 2) = "8"
    formatogrilla2(2, 3) = "8"
    formatogrilla2(2, 4) = "8"
    formatogrilla2(2, 5) = "8"
    formatogrilla2(2, 6) = "10"
    formatogrilla2(2, 7) = "17"
    
    Rem TIPO DE DATOS C=CEROS,S=STRING,N=NUMERICO,D=DATE,
    formatogrilla2(3, 1) = "D"
    formatogrilla2(3, 2) = "N"
    formatogrilla2(3, 3) = "N"
    formatogrilla2(3, 4) = "N"
    formatogrilla2(3, 5) = "N"
    formatogrilla2(3, 6) = "N"
    formatogrilla2(3, 7) = "N"
    
    Rem FORMATO GRILLA
    
    formatogrilla2(4, 2) = " ###,###,##0.00"
    formatogrilla2(4, 3) = " "
    formatogrilla2(4, 4) = " ###,###,##0"
    formatogrilla2(4, 5) = " ###,###,##0"
    
    Rem LOCCKED
    formatogrilla2(5, 1) = "TRUE"
    formatogrilla2(5, 2) = "TRUE"
    formatogrilla2(5, 3) = "TRUE"
    formatogrilla2(5, 4) = "TRUE"
    formatogrilla2(5, 5) = "TRUE"
    formatogrilla2(5, 6) = "TRUE"
    formatogrilla2(5, 7) = "TRUE"
    
    
    Rem VALOR MAXIMO
    
    Grid1.Cols = 6
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
     Grid1.Column(3).CellType = cellCheckBox
     
     
     
    End Sub

Sub CARGAGRILLA2()
    Dim formatogrilla2(10, 12)
    formatogrilla2(1, 1) = "BANCO"
    formatogrilla2(1, 2) = "TIPO"
    formatogrilla2(1, 3) = "NUMERO"
    formatogrilla2(1, 4) = "EMPRESA"
    formatogrilla2(1, 5) = "GLOSA  "
    formatogrilla2(1, 6) = "EMISION"
    formatogrilla2(1, 7) = "CAPITAL"
    formatogrilla2(1, 8) = "TIPO "
    formatogrilla2(1, 9) = "MONTO  "
    formatogrilla2(1, 10) = "VENCIMIENTO"
    formatogrilla2(1, 11) = "N/CUOTA"
    
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
    
    Grid2.Cols = 12
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
   
    leerCREDITOS
    
    
    
    End Sub


 Public Sub leerCREDITOS()
 Dim csql As New rdoQuery
 Dim resultados As rdoResultset
 
 Set csql.ActiveConnection = contadb
 csql.sql = "select cb.banco,cb.tipo,cb.numero,cb.empresa,cb.glosa,cb.fecha,cb.capital,cb.moneda,cb.monto,cv.fecha,concat(cv.numerocuota,'/',cb.cuotas) from " & clientesistema & "creditos_bancarios" & ".maestro_compromisos as cb "
 csql.sql = csql.sql + " inner join " + clientesistema & "creditos_bancarios" & ".creditos_vencimientos as cv on (cv.numero=cb.numero and cv.tipo=cb.tipo and cv.banco=cb.banco and cb.empresa=cv.empresa ) "
 csql.sql = csql.sql + "where cv.fecha<'" + Format(fechasistema, "yyyy-mm-dd") + "' and cv.pagado='0' "
 csql.sql = csql.sql + "order by cb.empresa "
 csql.Execute
 Grid2.Rows = 1
 If csql.RowsAffected > 0 Then
    Set resultados = csql.OpenResultset
    While resultados.EOF = False
    Grid2.Rows = Grid2.Rows + 1
    lblBanco.Caption = leerbanco(dato1.text)
    LBLTIPO.Caption = leertipocredito(dato2.text)
    LBLEMPRESA.Caption = leerempresa(dato4.text)
    LBLMONEDA.Caption = leertipoMONEDA(dato9.text)
    Grid2.Cell(Grid2.Rows - 1, 1).text = resultados(0) + "=" + leerbanco(resultados(0))
    Grid2.Cell(Grid2.Rows - 1, 2).text = resultados(1) + "=" + leertipocredito(resultados(1))
    Grid2.Cell(Grid2.Rows - 1, 3).text = resultados(2)
    Grid2.Cell(Grid2.Rows - 1, 4).text = resultados(3) + "=" + leerempresa(resultados(3))
    
    Grid2.Cell(Grid2.Rows - 1, 5).text = resultados(4)
    Grid2.Cell(Grid2.Rows - 1, 6).text = resultados(5)
    Grid2.Cell(Grid2.Rows - 1, 7).text = resultados(6)
    Grid2.Cell(Grid2.Rows - 1, 8).text = resultados(7) + "=" + leertipoMONEDA(resultados(7))
    Grid2.Cell(Grid2.Rows - 1, 9).text = resultados(8)
    Grid2.Cell(Grid2.Rows - 1, 10).text = resultados(9)
    Grid2.Cell(Grid2.Rows - 1, 11).text = resultados(10)
    

'    If Format(resultados(6), "yyyy-mm-dd") < Format(fechasistema, "yyyy-mm-dd") Then
'    Grid2.Range(Grid2.Rows - 1, 1, Grid2.Rows - 1, 11).BackColor = &HFF&
    
'    End If
    
    
    
    
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

Sub grabarmeses(banco, tipo, numero, empresa, monto, fecha, cuota)
    campos(0, 0) = "banco"
    campos(1, 0) = "tipo"
    campos(2, 0) = "numero"
    campos(3, 0) = "empresa"
    campos(4, 0) = "monto"
    campos(5, 0) = "fecha"
    campos(6, 0) = "numerocuota"
    campos(7, 0) = ""
    
    campos(0, 1) = banco
    campos(1, 1) = tipo
    campos(2, 1) = numero
    campos(3, 1) = empresa
    campos(4, 1) = Replace(monto, ",", ".")
    campos(5, 1) = fecha
    campos(6, 1) = cuota
    campos(0, 2) = clientesistema & "creditos_bancarios" & ".creditos_vencimientos"
    op = 2
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)
    
    End Sub

Public Sub leermensualidades(banco, tipo, numero, empresa)
 Dim csql As New rdoQuery
 Dim resultados As rdoResultset
 
 Set csql.ActiveConnection = contadb
 csql.sql = "select fecha,monto,pagado,numerocuota from " & clientesistema & "creditos_bancarios" & ".creditos_vencimientos as mp where tipo='" + tipo + "' and banco='" + banco + "' and empresa='" + empresa + "' and numero='" + numero + "' "
 If Option1.Value = True Then
 csql.sql = csql.sql + " and pagado='0' "
 End If
 csql.sql = csql.sql + "order by fecha "
 
 csql.Execute
 Grid1.Rows = 1
 If csql.RowsAffected > 0 Then
    Set resultados = csql.OpenResultset
    While resultados.EOF = False
    Grid1.Rows = Grid1.Rows + 1
    Grid1.Cell(Grid1.Rows - 1, 1).text = resultados(0)
    Grid1.Cell(Grid1.Rows - 1, 2).text = resultados(1)
    Grid1.Cell(Grid1.Rows - 1, 3).text = resultados(2)
    Grid1.Cell(Grid1.Rows - 1, 4).text = resultados(3)
    
    resultados.MoveNext
    
    Wend
    
    
  End If
 csql.Close
 Set csql = Nothing
 Set resultados = Nothing
 
 End Sub
Public Function leerpagado(banco, tipo, numero, empresa) As Double

 Dim csql As New rdoQuery
 Dim resultados As rdoResultset
 
 Set csql.ActiveConnection = contadb
 csql.sql = "select IFNULL(sum(monto),0) from " & clientesistema & "creditos_bancarios" & ".creditos_vencimientos as mp where tipo='" + tipo + "' and banco='" + banco + "' and empresa='" + empresa + "' and numero='" + numero + "' "
 csql.sql = csql.sql + " and pagado='1' "
 csql.sql = csql.sql + "order by fecha "
 
 csql.Execute
 Grid1.Rows = 1
 If csql.RowsAffected > 0 Then
    Set resultados = csql.OpenResultset
 leerpagado = resultados(0)
 Else
 leerpagado = 0
 
  End If
 csql.Close
 Set csql = Nothing
 Set resultados = Nothing
 
 End Function


Sub modificapago(banco, tipo, numero, empresa, fecha, pagado)
    campos(0, 0) = "pagado"
    campos(1, 0) = ""
    campos(0, 1) = pagado
    campos(0, 2) = clientesistema & "creditos_bancarios" & ".creditos_vencimientos"
    condicion = "banco='" + dato1.text + "' and tipo='" + dato2.text + "' and numero ='" & dato3.text & "' and empresa='" + dato4.text + "' and fecha='" + fecha + "'  "
    
    op = 3
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)
    
    End Sub
Sub modificamonto(banco, tipo, numero, empresa, fecha, monto)
    campos(0, 0) = "monto"
    campos(1, 0) = ""
    campos(0, 1) = Replace(monto, ",", ".")
    campos(0, 2) = clientesistema & "creditos_bancarios" & ".creditos_vencimientos"
    condicion = "banco='" + dato1.text + "' and tipo='" + dato2.text + "' and numero ='" & dato3.text & "' and empresa='" + dato4.text + "' and fecha='" + fecha + "'  "
    
    op = 3
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)
    
    End Sub

 
Private Sub Option1_Click()
Call leermensualidades(dato1.text, dato2.text, dato3.text, dato4.text)

End Sub

Private Sub Option2_Click()
Call leermensualidades(dato1.text, dato2.text, dato3.text, dato4.text)

End Sub

Private Sub dato1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF2 Then Call ayudabancos(dato1)
    If KeyCode = 38 Then Unload Me: GoTo no:
    
    Call flechas(dato1, dato2, KeyCode)
no:
End Sub
Sub ayudabancos(ByRef caja As TextBox)
    Dim campos As Variant
    Dim cfijo As Variant
    Dim largo As Variant
    campos = Array("codigobanco", "nombre")
    cabezas = Array("CODIGO", "NOMBRE")
    largo = Array("4N", "40s")
    mensajeAyuda = "Ayuda Bancos"
    cfijo = "no"
    
    Call cargaAyudaT(Servidor, clientesistema + "conta", Usuario, password, "maestrobancos", caja, campos, cfijo, largo, 2)
    If caja.text = "" Then caja.SetFocus: GoTo no
    caja.Enabled = True
    caja.SetFocus


no:

End Sub

Private Sub dato2_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF2 Then Call ayudadepositos(dato2)
    
    Call flechas(dato1, dato3, KeyCode)
End Sub
Sub ayudadepositos(ByRef caja As TextBox)
    Dim campos As Variant
    Dim cfijo As Variant
    Dim largo As Variant
    campos = Array("codigo", "nombre")
    cabezas = Array("CODIGO", "NOMBRE")
    largo = Array("4N", "40s")
    mensajeAyuda = "Ayuda tipo de creditos"
    cfijo = "no"
    
    Call cargaAyudaT(Servidor, clientesistema + "creditos_bancarios", Usuario, password, "maestro_tipo_compromiso", caja, campos, cfijo, largo, 2)
    If caja.text = "" Then caja.SetFocus: GoTo no
    caja.Enabled = True
    caja.SetFocus


no:

End Sub


Private Sub dato3_KeyDown(KeyCode As Integer, Shift As Integer)
    Call flechas(dato2, dato4, KeyCode)
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
    Call flechas(dato7, dato9, KeyCode)
End Sub
Private Sub dato9_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF2 Then Call ayudatipomonedas(dato9)
    
    Call flechas(dato8, dato10, KeyCode)
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



Private Sub dato10_KeyDown(KeyCode As Integer, Shift As Integer)
    Call flechas(dato9, dato11, KeyCode)
End Sub
Private Sub dato11_KeyDown(KeyCode As Integer, Shift As Integer)
    Call flechas(dato10, dato12, KeyCode)
End Sub
Private Sub dato12_KeyDown(KeyCode As Integer, Shift As Integer)
    Call flechas(dato11, dato13, KeyCode)
End Sub
Private Sub dato13_KeyDown(KeyCode As Integer, Shift As Integer)
    Call flechas(dato12, dato14, KeyCode)
End Sub
Private Sub dato14_KeyDown(KeyCode As Integer, Shift As Integer)
    Call flechas(dato13, dato15, KeyCode)
End Sub
Private Sub dato15_KeyDown(KeyCode As Integer, Shift As Integer)
    Call flechas(dato14, dato15, KeyCode)
End Sub

Private Sub dato1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then Unload Me
    snum = 0: KeyAscii = esNumero(KeyAscii)
    If KeyAscii = 13 Then
        
    Call ceros(dato1)
    If leerbanco(dato1.text) <> "" Then
    lblBanco.Caption = leerbanco(dato1.text)
    dato2.Enabled = True
    
    dato2.SetFocus
    Else
    dato1.SetFocus
    End If
    End If
    
    
    
    
End Sub
Private Sub dato2_KeyPress(KeyAscii As Integer)
    snum = 0: KeyAscii = esNumero(KeyAscii)
    If KeyAscii = 13 Then
    Call ceros(dato2)

If leertipocredito(dato2.text) <> "" Then
    LBLTIPO.Caption = leertipocredito(dato2.text)
    dato3.Enabled = True
    
    dato3.SetFocus
    Else
    dato2.SetFocus
    End If
    
    End If
    
End Sub

Private Sub dato3_KeyPress(KeyAscii As Integer)
    snum = 0: KeyAscii = esNumero(KeyAscii)
    If KeyAscii = 13 Then Call ceros(dato3): Call Pregunta(dato3, dato4)
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
    If KeyAscii = 13 Then Call ceros(DATO5): Call Pregunta(DATO5, dato6)
End Sub
Private Sub dato6_KeyPress(KeyAscii As Integer)
    ' If KeyAscii = 42 And SUMADEBE = SUMAHABER Then grabarcomprobante:retorno: dato3.Enabled = True: dato3.SetFocus
    snum = 0: KeyAscii = esNumero(KeyAscii)
    If KeyAscii = 13 Then Call ceros(dato6): Call Pregunta(dato6, dato7)
no:
End Sub
Private Sub dato7_KeyPress(KeyAscii As Integer)
    Dim PASO As String
    snum = 0: KeyAscii = esNumero(KeyAscii)
    
    If KeyAscii = 13 Then
        Call ceros(dato7)
        PASO = DATO5.text & "-" & dato6.text & "-" & dato7.text
        If IsDate(PASO) = True Then
            Call Pregunta(dato7, dato8)
        Else
            MsgBox "FECHA INVALIDA", vbCritical, "ATENCION"
            DATO5.SetFocus
        End If
    End If
End Sub
Private Sub dato8_KeyPress(KeyAscii As Integer)
    snum = 0: KeyAscii = esNumero(KeyAscii)
    If KeyAscii = 13 Then Call Pregunta(dato8, dato9)
End Sub
Private Sub dato9_KeyPress(KeyAscii As Integer)
    snum = 0: KeyAscii = esNumero(KeyAscii)
    
    If KeyAscii = 13 Then
    Call ceros(dato9)
    If leertipoMONEDA(dato9.text) <> "" Then
    
    LBLMONEDA.Caption = leertipoMONEDA(dato9.text)
    dato10.Enabled = True
    dato10.SetFocus
    Else
    dato9.SetFocus
    End If

    
    End If
    
End Sub
Private Sub dato10_KeyPress(KeyAscii As Integer)
    snum = 0: KeyAscii = esNumero(KeyAscii)
    If KeyAscii = 13 Then Call ceros(dato10): Call Pregunta(dato10, dato11)
End Sub

Private Sub dato11_KeyPress(KeyAscii As Integer)
    snum = 1: KeyAscii = esNumero(KeyAscii)
    If KeyAscii = 13 Then Call ceros(dato11): Call Pregunta(dato11, dato12)
End Sub
Private Sub dato12_KeyPress(KeyAscii As Integer)
    snum = 1: KeyAscii = esNumero(KeyAscii)
       
    If KeyAscii = 13 Then
    Call Pregunta(dato12, dato13)
    End If
End Sub
Private Sub dato13_KeyPress(KeyAscii As Integer)
    snum = 1: KeyAscii = esNumero(KeyAscii)
    If KeyAscii = 13 Then
    Call Pregunta(dato13, dato14)
    End If
End Sub
Private Sub dato14_KeyPress(KeyAscii As Integer)
    snum = 1: KeyAscii = esNumero(KeyAscii)
    If KeyAscii = 13 Then
    Call Pregunta(dato14, dato15)
    End If
End Sub
Private Sub dato15_KeyPress(KeyAscii As Integer)
    snum = 1: KeyAscii = esNumero(KeyAscii)
    If KeyAscii = 13 Then
    Call Pregunta(dato15, dato16)
    

    End If
End Sub
Private Sub dato16_KeyPress(KeyAscii As Integer)
     KeyAscii = Asc(UCase(Chr(KeyAscii)))

    If KeyAscii = 13 Then
    grabar
    retorno
    

    End If
End Sub
Private Sub botonmisaccesos_Click()
programafiltro = Me.Caption
misaccesos.Show
End Sub


Private Sub botonmisfavoritos_Click()
Call AgregaFavoritos(USUARIOSISTEMA, App.EXEName, Me.Name, Me.Caption)
End Sub
