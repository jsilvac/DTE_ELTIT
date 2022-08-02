VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash.ocx"
Object = "{ADBBAED6-B16F-40DC-80DF-B44910CBA76C}#1.0#0"; "Frame-Xp.ocx"
Begin VB.Form inver03 
   BackColor       =   &H00FF8080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Ingreso Inversiones"
   ClientHeight    =   9225
   ClientLeft      =   45
   ClientTop       =   1230
   ClientWidth     =   9990
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9225
   ScaleWidth      =   9990
   Begin XPFrame.FrameXp FrameXp5 
      Height          =   615
      Left            =   6840
      TabIndex        =   62
      Top             =   8520
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
      Begin VB.CommandButton botonmisfavoritos 
         Caption         =   "Mis Favoritos"
         Height          =   255
         Left            =   1800
         TabIndex        =   64
         Top             =   280
         Width           =   1335
      End
      Begin VB.CommandButton botonmisaccesos 
         Caption         =   "Permisos Modulo"
         Height          =   255
         Left            =   120
         TabIndex        =   63
         Top             =   280
         Width           =   1455
      End
   End
   Begin XPFrame.FrameXp FrameXp1 
      Height          =   7260
      Left            =   180
      TabIndex        =   6
      Top             =   180
      Width           =   8835
      _ExtentX        =   15584
      _ExtentY        =   12806
      BackColor       =   16744576
      Caption         =   "Numero Inversion"
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
      Begin XPFrame.FrameXp FrameXp2 
         Height          =   2670
         Left            =   90
         TabIndex        =   12
         Top             =   1620
         Width           =   8610
         _ExtentX        =   15187
         _ExtentY        =   4710
         BackColor       =   16761024
         Caption         =   "Comprobante Bancario"
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
         Begin XPFrame.FrameXp FrameXp4 
            Height          =   2265
            Left            =   4500
            TabIndex        =   42
            Top             =   315
            Width           =   3975
            _ExtentX        =   7011
            _ExtentY        =   3995
            BackColor       =   16744576
            Caption         =   "INVERSIONES"
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
            Begin VB.TextBox MONTOPESOS 
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
               Left            =   2160
               MaxLength       =   20
               TabIndex        =   60
               Top             =   1440
               Width           =   1455
            End
            Begin VB.CommandButton Command3 
               Caption         =   "Grabar"
               Height          =   330
               Left            =   1170
               TabIndex        =   54
               Top             =   1890
               Width           =   1680
            End
            Begin VB.TextBox dato17 
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
               MaxLength       =   20
               TabIndex        =   48
               Top             =   1080
               Width           =   1455
            End
            Begin VB.TextBox dato16 
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
               MaxLength       =   20
               TabIndex        =   47
               Top             =   720
               Width           =   1455
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
               Left            =   2160
               MaxLength       =   20
               TabIndex        =   46
               Top             =   360
               Width           =   1455
            End
            Begin VB.Label Label20 
               BackStyle       =   0  'Transparent
               BorderStyle     =   1  'Fixed Single
               Caption         =   " MONTO PESOS"
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
               TabIndex        =   61
               Top             =   1485
               Width           =   1905
            End
            Begin VB.Label Label19 
               BackStyle       =   0  'Transparent
               BorderStyle     =   1  'Fixed Single
               Caption         =   " T/C EMISION"
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
               TabIndex        =   45
               Top             =   1125
               Width           =   1905
            End
            Begin VB.Label Label18 
               BackStyle       =   0  'Transparent
               BorderStyle     =   1  'Fixed Single
               Caption         =   " MONTO A COBRAR"
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
               TabIndex        =   44
               Top             =   720
               Width           =   1905
            End
            Begin VB.Label Label17 
               BackStyle       =   0  'Transparent
               BorderStyle     =   1  'Fixed Single
               Caption         =   " MONTO INTERES"
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
               TabIndex        =   43
               Top             =   360
               Width           =   1905
            End
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
            Left            =   2610
            MaxLength       =   20
            TabIndex        =   28
            Top             =   2160
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
            Left            =   2610
            MaxLength       =   20
            TabIndex        =   27
            Top             =   1800
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
            Left            =   2610
            MaxLength       =   20
            TabIndex        =   26
            Top             =   1440
            Width           =   1455
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
            Left            =   2610
            MaxLength       =   3
            TabIndex        =   25
            Tag             =   "fecha"
            Top             =   1080
            Width           =   510
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
            Left            =   3465
            MaxLength       =   4
            TabIndex        =   24
            Tag             =   "fecha"
            Top             =   675
            Width           =   645
         End
         Begin VB.TextBox dato9 
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
            TabIndex        =   23
            Tag             =   "fecha"
            Top             =   675
            Width           =   375
         End
         Begin VB.TextBox dato8 
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
            Top             =   675
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
            TabIndex        =   21
            Tag             =   "fecha"
            Top             =   315
            Width           =   645
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
            TabIndex        =   20
            Tag             =   "fecha"
            Top             =   315
            Width           =   375
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
            TabIndex        =   19
            Tag             =   "fecha"
            Top             =   315
            Width           =   375
         End
         Begin VB.Label Label10 
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   " INTERES PERIODO"
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
            TabIndex        =   18
            Top             =   2160
            Width           =   2445
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   " INTERES ANUAL"
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
            TabIndex        =   17
            Top             =   1800
            Width           =   2445
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   " DIAS"
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
            TabIndex        =   16
            Top             =   1080
            Width           =   2445
         End
         Begin VB.Label Label7 
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   " MONTO"
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
            TabIndex        =   15
            Top             =   1440
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
            TabIndex        =   14
            Top             =   315
            Width           =   2445
         End
         Begin VB.Label Label9 
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   " FECHA VENCIMIENTO"
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
            TabIndex        =   13
            Top             =   675
            Width           =   2445
         End
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
         TabIndex        =   11
         Tag             =   "fecha"
         Top             =   1260
         Width           =   375
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
      Begin XPFrame.FrameXp FrameXp3 
         Height          =   2760
         Left            =   135
         TabIndex        =   29
         Top             =   4410
         Width           =   8520
         _ExtentX        =   15028
         _ExtentY        =   4868
         BackColor       =   8454016
         Caption         =   "VALORES EN PESOS"
         CaptionEstilo3D =   1
         BackColor       =   8454016
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
         Begin VB.CommandButton Command1 
            Caption         =   "CLICK INGRESA T/Cambio"
            Height          =   285
            Left            =   45
            TabIndex        =   59
            Top             =   1440
            Width           =   2445
         End
         Begin VB.CommandButton anulav 
            Caption         =   "Anular"
            Height          =   285
            Left            =   6525
            TabIndex        =   56
            Top             =   1845
            Width           =   1005
         End
         Begin VB.CommandButton anulae 
            Caption         =   "Anular"
            Height          =   285
            Left            =   6480
            TabIndex        =   55
            Top             =   810
            Width           =   1005
         End
         Begin VB.CommandButton contav 
            Caption         =   "Contabilizar"
            Height          =   285
            Left            =   4905
            TabIndex        =   50
            Top             =   1845
            Width           =   1005
         End
         Begin VB.CommandButton contae 
            Caption         =   "Contabilizar"
            Height          =   285
            Left            =   4950
            TabIndex        =   49
            Top             =   810
            Width           =   1005
         End
         Begin VB.TextBox dato18 
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
            Height          =   285
            Left            =   2610
            Locked          =   -1  'True
            MaxLength       =   20
            TabIndex        =   35
            Tag             =   "fecha"
            Top             =   315
            Width           =   1455
         End
         Begin VB.TextBox dato19 
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
            Height          =   285
            Left            =   2610
            Locked          =   -1  'True
            MaxLength       =   20
            TabIndex        =   34
            Tag             =   "fecha"
            Top             =   675
            Width           =   1455
         End
         Begin VB.TextBox dato20 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
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
            Locked          =   -1  'True
            MaxLength       =   20
            TabIndex        =   33
            Tag             =   "fecha"
            Top             =   1080
            Width           =   1455
         End
         Begin VB.TextBox dato21 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
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
            Locked          =   -1  'True
            MaxLength       =   20
            TabIndex        =   32
            Top             =   1440
            Width           =   1455
         End
         Begin VB.TextBox dato22 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
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
            Locked          =   -1  'True
            MaxLength       =   20
            TabIndex        =   31
            Top             =   1800
            Width           =   1455
         End
         Begin VB.TextBox dato23 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
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
            Locked          =   -1  'True
            MaxLength       =   20
            TabIndex        =   30
            Top             =   2160
            Width           =   1455
         End
         Begin VB.Label CONTACOBRO 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "RESCATE NO CONTABILIZADO"
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
            Left            =   4590
            TabIndex        =   58
            Top             =   1395
            Width           =   3660
         End
         Begin VB.Label contadepo 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "DEPOSITO NO CONTABILIZADO"
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
            Left            =   4590
            TabIndex        =   57
            Top             =   360
            Width           =   3660
         End
         Begin VB.Label Label16 
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
            Height          =   255
            Left            =   45
            TabIndex        =   41
            Top             =   675
            Width           =   2445
         End
         Begin VB.Label Label15 
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   " MONTO"
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
            TabIndex        =   40
            Top             =   315
            Width           =   2445
         End
         Begin VB.Label Label14 
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   " T/C VENCIMIENTO"
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
            TabIndex        =   39
            Top             =   1440
            Width           =   2445
         End
         Begin VB.Label Label13 
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   " MONTO A COBRAR"
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
            TabIndex        =   38
            Top             =   1080
            Width           =   2445
         End
         Begin VB.Label Label12 
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   " MONTO CAMBIO"
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
            Top             =   1800
            Width           =   2445
         End
         Begin VB.Label Label11 
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   " DIFERENCIA CAMBIO"
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
            TabIndex        =   36
            Top             =   2160
            Width           =   2445
         End
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
         TabIndex        =   53
         Top             =   1260
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
         TabIndex        =   52
         Top             =   630
         Width           =   4785
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
         TabIndex        =   51
         Top             =   315
         Width           =   4785
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
         TabIndex        =   10
         Top             =   315
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
         TabIndex        =   9
         Top             =   1260
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
         TabIndex        =   8
         Top             =   945
         Width           =   1815
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
         TabIndex        =   7
         Top             =   630
         Width           =   1815
      End
   End
   Begin VB.PictureBox MANUAL 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   135
      Left            =   0
      ScaleHeight     =   135
      ScaleWidth      =   615
      TabIndex        =   5
      Top             =   0
      Width           =   615
   End
   Begin VB.TextBox pivote2 
      Height          =   285
      Left            =   7440
      MaxLength       =   10
      TabIndex        =   4
      Top             =   7560
      Visible         =   0   'False
      Width           =   1215
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash opciones 
      Height          =   1800
      Left            =   180
      TabIndex        =   3
      Top             =   7290
      Width           =   7065
      _cx             =   12462
      _cy             =   3175
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
Attribute VB_Name = "inver03"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Private tipocuenta As String
    Private cc As Integer
    Private FORMATOGRILLA(100, 20)
    Private formatogrilla2(100, 20)
    Private cdi As Integer
    Private CANDO As Integer
    Private existe As String
    Private AFECTO As Double
    Private EXENTO As Double
    Private MODIFI As String
        
    Private AUXILIAR(1000, 3) As String
    
    Private respu As String
    Private tipoctacte As String
    Private nlineas As Double
    Private DOCU(6) As String
    Private grilladetalle(1000, 13) As String
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
     Private año As String
     
     
    
Sub grabarcomprobantedeposito(empresa)
    Dim TIPOCON As String
    Dim tipo2 As String
    Dim j As Integer
    Dim lin As Integer
    Dim debehaber As String
    Dim fechacom As String
    Dim cuenta As String
    Dim CUENTA2 As String
    
    
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
    campos(17, 0) = "año"
    campos(18, 0) = "fechacreacion"
    campos(19, 0) = "horacreacion"
    campos(20, 0) = "rutproveedor"
    campos(21, 0) = ""
    Rem cuenta proveedores
    
    campos(0, 1) = "DP"
    campos(1, 1) = dato3.text
    
    campos(2, 1) = "001"
    campos(3, 1) = dato7.text + "-" + dato6.text + "-" + DATO5.text
    If dato2.text <> "05" Then
      cuenta = "11150003"
      Else
      cuenta = "11150001"
      End If
      
    If dato2.text <> "05" Then
    CUENTA2 = "11500180"
    Else
    CUENTA2 = "11500180"
    End If
    campos(3, 1) = dato7.text + "-" + dato6.text + "-" + DATO5.text
    campos(4, 1) = cuenta
    campos(5, 1) = ""
    campos(6, 1) = ""
    campos(7, 1) = ""
    If dato2.text = "05" Then
    campos(8, 1) = "CONTABILIZACION FONDO MUTUO"
    Else
    campos(8, 1) = "CONTABILIZACION DEPOSITO A PLAZO"
    
    End If
    
    campos(9, 1) = "DP"
    campos(10, 1) = dato3.text
    campos(11, 1) = dato7.text + "-" + dato6.text + "-" + DATO5.text
    campos(12, 1) = dato10.text + "-" + dato9.text + "-" + dato8.text
    campos(13, 1) = Replace(dato18.text, ".", "")
    campos(14, 1) = "D"
    campos(15, 1) = USUARIOSISTEMA
    campos(16, 1) = dato6.text
    campos(17, 1) = dato7.text
    campos(18, 1) = Format(Date$, "yyyy") + "-" + Format(Date$, "mm") + "-" + Format(Date$, "dd")
    campos(19, 1) = Time$
    campos(20, 1) = campos(6, 1)

    campos(0, 2) = clientesistema + "conta" + empresa + ".movimientoscontables"
    condicion = ""
    op = 2
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)
    
    campos(0, 1) = "DP"
    campos(1, 1) = dato3.text
    campos(2, 1) = "002"
    campos(4, 1) = CUENTA2
    campos(5, 1) = ""
    campos(6, 1) = ""
    campos(7, 1) = ""
    If dato2.text = "05" Then
    campos(8, 1) = "CONTABILIZACION FONDO MUTUO"
    Else
    campos(8, 1) = "CONTABILIZACION DEPOSITO A PLAZO"
    
    End If
    campos(9, 1) = "DP"
    campos(10, 1) = dato3.text
    campos(11, 1) = dato7.text + "-" + dato6.text + "-" + DATO5.text
    campos(12, 1) = dato10.text + "-" + dato9.text + "-" + dato8.text
    campos(13, 1) = Replace(dato18.text, ".", "")
    campos(14, 1) = "H"
    campos(15, 1) = USUARIOSISTEMA
    campos(16, 1) = dato6.text
    campos(17, 1) = dato7.text
    campos(18, 1) = Format(Date$, "yyyy") + "-" + Format(Date$, "mm") + "-" + Format(Date$, "dd")
    campos(19, 1) = Time$
    campos(20, 1) = campos(6, 1)

    campos(0, 2) = clientesistema + "conta" + empresa + ".movimientoscontables"
    condicion = ""
    op = 2
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)
    
    
    Rem  Call ACTUALIZADOCUMENTO("+")
   
End Sub

Sub grabarcomprobanteCOBRO(empresa)
    Dim TIPOCON As String
    Dim tipo2 As String
    Dim j As Integer
    Dim lin As Integer
    Dim debehaber As String
    Dim fechacom As String
    Dim cuenta As String
    Dim CUENTA2 As String
    Dim monto As Double
    Dim GLOSAREAJUSTE As String
    Dim DH As String
    
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
    campos(17, 0) = "año"
    campos(18, 0) = "fechacreacion"
    campos(19, 0) = "horacreacion"
    campos(20, 0) = "rutproveedor"
    campos(21, 0) = ""
    Rem DEPOSITO A PLAZO
    campos(0, 1) = "RD"
    campos(1, 1) = dato3.text
    
    campos(2, 1) = "001"
    campos(3, 1) = dato10.text + "-" + dato9.text + "-" + dato8.text
   
      
      If dato2.text <> "05" Then
      cuenta = "11150003"
      Else
      cuenta = "11150001"
    
      End If
    If dato2.text <> "05" Then
    CUENTA2 = "11500180"
    Else
    CUENTA2 = "11150001"
    End If
    campos(3, 1) = dato10.text + "-" + dato9.text + "-" + dato8.text
    campos(4, 1) = CUENTA2
    campos(5, 1) = ""
    campos(6, 1) = ""
    campos(7, 1) = ""
    campos(8, 1) = "RESCATE DEPOSITO A PLAZO"
    campos(9, 1) = "RD"
    campos(10, 1) = dato3.text
    campos(11, 1) = dato7.text + "-" + dato6.text + "-" + DATO5.text
    campos(12, 1) = dato10.text + "-" + dato9.text + "-" + dato8.text
    campos(13, 1) = Replace(DATO20.text, ".", "")
    campos(14, 1) = "D"
    campos(15, 1) = USUARIOSISTEMA
    campos(16, 1) = dato9.text
    campos(17, 1) = dato10.text
    campos(18, 1) = Format(Date$, "yyyy") + "-" + Format(Date$, "mm") + "-" + Format(Date$, "dd")
    campos(19, 1) = Time$
    campos(20, 1) = campos(6, 1)

    campos(0, 2) = clientesistema + "conta" + empresa + ".movimientoscontables"
    condicion = ""
    op = 2
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)
    
    
    
    Rem CUENTA DEPOSITOS
    campos(0, 1) = "RD"
    campos(1, 1) = dato3.text
    campos(2, 1) = "002"
    campos(4, 1) = cuenta
    campos(5, 1) = ""
    campos(6, 1) = ""
    campos(7, 1) = ""
    campos(8, 1) = "RESCATE DEPOSITO A PLAZO"
    campos(9, 1) = "RD"
    campos(10, 1) = dato3.text
    campos(11, 1) = dato7.text + "-" + dato6.text + "-" + DATO5.text
    campos(12, 1) = dato10.text + "-" + dato9.text + "-" + dato8.text
    campos(13, 1) = Replace(dato18.text, ".", "")
    campos(14, 1) = "H"
    campos(15, 1) = USUARIOSISTEMA
    campos(16, 1) = dato9.text
    campos(17, 1) = dato10.text
    campos(18, 1) = Format(Date$, "yyyy") + "-" + Format(Date$, "mm") + "-" + Format(Date$, "dd")
    campos(19, 1) = Time$
    campos(20, 1) = campos(6, 1)

    campos(0, 2) = clientesistema + "conta" + empresa + ".movimientoscontables"
    condicion = ""
    op = 2
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)
    
    Rem INTERES
    
    lin = 2
    If CDbl(dato19.text) <> 0 Then
    lin = lin + 1
    campos(0, 1) = "RD"
    campos(1, 1) = dato3.text
    campos(2, 1) = Format(lin, "000")
    campos(4, 1) = "35200014"
    campos(5, 1) = ""
    campos(6, 1) = ""
    campos(7, 1) = ""
    campos(8, 1) = "INTERESES DEPOSITO A PLAZO"
    campos(9, 1) = "RD"
    campos(10, 1) = dato3.text
    campos(11, 1) = dato7.text + "-" + dato6.text + "-" + DATO5.text
    campos(12, 1) = dato10.text + "-" + dato9.text + "-" + dato8.text
    campos(13, 1) = Replace(dato19.text, ".", "")
    campos(14, 1) = "H"
    campos(15, 1) = USUARIOSISTEMA
    campos(16, 1) = dato9.text
    campos(17, 1) = dato10.text
    campos(18, 1) = Format(Date$, "yyyy") + "-" + Format(Date$, "mm") + "-" + Format(Date$, "dd")
    campos(19, 1) = Time$
    campos(20, 1) = campos(6, 1)

    campos(0, 2) = clientesistema + "conta" + empresa + ".movimientoscontables"
    condicion = ""
    op = 2
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)
    End If
    Rem DIFERENCIA REAJUSTE
    If CDbl(dato23.text) <> 0 Then
    monto = CDbl(dato23.text)
    If monto < 1 Then
    monto = monto * -1: DH = "D"
    Else
    monto = monto * 1: DH = "H"
    
    End If
    If dato2.text = "01" Or dato2.text = "04" Then
    CUENTA2 = "35200001"
    GLOSAREAJUSTE = "REAJUSTE U.F RESCATE DEPOSITO A PLAZO"
    Else
    CUENTA2 = "35200001"
    GLOSAREAJUSTE = "DIFERENCIA CAMBIO RESCATE DEPOSITO A PLAZO"
    
    End If
    
    lin = lin + 1
    campos(0, 1) = "RD"
    campos(1, 1) = dato3.text
    campos(2, 1) = Format(lin, "000")
    campos(4, 1) = CUENTA2
    campos(5, 1) = ""
    campos(6, 1) = ""
    campos(7, 1) = ""
    campos(8, 1) = GLOSAREAJUSTE
    campos(9, 1) = "RD"
    campos(10, 1) = dato3.text
    campos(11, 1) = dato7.text + "-" + dato6.text + "-" + DATO5.text
    campos(12, 1) = dato10.text + "-" + dato9.text + "-" + dato8.text
    campos(13, 1) = Str(monto)
    campos(14, 1) = DH
    campos(15, 1) = USUARIOSISTEMA
    campos(16, 1) = dato9.text
    campos(17, 1) = dato10.text
    campos(18, 1) = Format(Date$, "yyyy") + "-" + Format(Date$, "mm") + "-" + Format(Date$, "dd")
    campos(19, 1) = Time$
    campos(20, 1) = campos(6, 1)

    campos(0, 2) = clientesistema + "conta" + empresa + ".movimientoscontables"
    condicion = ""
    op = 2
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)
    
    
    Rem  Call ACTUALIZADOCUMENTO("+")
   End If
End Sub


Private Sub anulae_Click()
Call eliminadepositocontabilizado(dato3.text, dato7.text + "-" + dato6.text + "-" + DATO5.text, dato4.text, "DP")
leer
End Sub

Private Sub anulav_Click()
Call eliminadepositocontabilizado(dato3.text, dato10.text + "-" + dato9.text + "-" + dato8.text, dato4.text, "RD")
leer

End Sub

Private Sub Command1_Click()
DATO21.Locked = False
DATO21.SetFocus

MODIFI = "1"

End Sub

Private Sub Command3_Click()
grabar

End Sub

Private Sub contae_Click()
If contadepo.Caption = "DEPOSITO NO CONTABILIZADO" Then
Call grabarcomprobantedeposito(dato4.text)
leer
End If

End Sub

Private Sub contav_Click()
If DATO21.text <> "" Then
If CDbl(DATO21.text) <> 0 Then

If CONTACOBRO.Caption = "RESCATE NO CONTABILIZADO" Then
Call grabarcomprobanteCOBRO(dato4.text)

leer
End If
Else
MsgBox ("debe ingresar tipo de cambio")
End If
End If
End Sub

Private Sub dato14_GotFocus()
dato14.text = (CDbl(dato13.text) / 360) * CDbl(dato11.text)

End Sub

Private Sub dato21_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 And DATO21.text <> "" And DATO21.text <> "0" Then
MODIFI = "1"
dato19.text = Format(CDbl(dato15.text) * CDbl(DATO21.text), "###,###,###.##0")
DATO21.text = Format(DATO21.text, "###,###,##0.00")
DATO22.text = Format(CDbl(dato12.text) * CDbl(DATO21.text), "###,###,###,##0")
dato23.text = Format((CDbl(DATO22.text) - (dato18.text)), "###,###,###,##0")
DATO20.text = Format(CDbl(DATO22.text) + CDbl(dato19.text), "###,###,###,##0")

grabar

End If

End Sub

Private Sub dato5_GotFocus()
leer
End Sub

Private Sub Form_Load()
CENTRAR Me
iva = 19
    Call Conectar_BD
    Call Conectarconta(Servidor, clientesistema + "conta", Usuario, password)
    sc = 0
    opciones.Visible = False

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
    mensajeAyuda = "Ayuda Depositos"
    cfijo = "no"
    
    Call cargaAyudaT(Servidor, clientesistema + "conta", Usuario, password, "maestrodepositos", caja, campos, cfijo, largo, 2)
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
    Call flechas(dato8, dato10, KeyCode)
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
    Call flechas(dato14, dato16, KeyCode)
End Sub
Private Sub dato16_KeyDown(KeyCode As Integer, Shift As Integer)
    Call flechas(dato15, dato17, KeyCode)
End Sub
Private Sub dato17_KeyDown(KeyCode As Integer, Shift As Integer)
    Call flechas(dato16, dato17, KeyCode)
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
    If leerdeposito(dato2.text) <> "" Then
    LBLTIPO.Caption = leerdeposito(dato2.text)
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
    If KeyAscii = 13 Then Call ceros(dato9): Call Pregunta(dato9, dato10)
End Sub
Private Sub dato10_KeyPress(KeyAscii As Integer)
Dim paso2 As String
        snum = 0: KeyAscii = esNumero(KeyAscii)
        If KeyAscii = 13 Then
            Call ceros(dato10)
            paso2 = dato8.text & "-" & dato9.text & "-" & dato10.text
            If IsDate(paso2) = True Then
                Call Pregunta(dato10, dato11)
            Else
                MsgBox "FECHA INVALIDA", vbCritical, "ATENCION"
                dato8.SetFocus
            End If
        End If
End Sub
Private Sub dato11_KeyPress(KeyAscii As Integer)
    snum = 1: KeyAscii = esNumero(KeyAscii)
    If KeyAscii = 13 Then Call ceros(dato11): Call Pregunta(dato11, dato12)
End Sub
Private Sub dato12_KeyPress(KeyAscii As Integer)
    snum = 1: KeyAscii = esNumero(KeyAscii)
    If KeyAscii = 13 Then
    Call Pregunta(dato12, dato13)
    If dato2.text = "01" Then
    dato12.text = Format(CDbl(dato12.text), "###,###,###,##0")
    Else
    dato12.text = Format(CDbl(dato12.text), "###,###,###,##0.0000")
    
    End If
    
    End If
    
End Sub
Private Sub dato13_KeyPress(KeyAscii As Integer)
    snum = 1: KeyAscii = esNumero(KeyAscii)
    If KeyAscii = 13 Then
        
    Call Pregunta(dato13, dato14)
    If dato2.text = "01" Then
    dato13.text = Format(CDbl(dato13.text), "###,###,###,##0.0000")
    Else
    dato13.text = Format(CDbl(dato13.text), "###,###,###,##0.0000")
    
    End If
    
End If

End Sub

Private Sub dato14_KeyPress(KeyAscii As Integer)
    snum = 1: KeyAscii = esNumero(KeyAscii)
    If KeyAscii = 13 And dato14.text <> "" And dato14.text <> "0" Then
    Call Pregunta(dato14, dato15)
    If dato2.text = "01" Then
    dato14.text = Format(CDbl(dato14.text), "###,###,###,##0.0000")
    dato15.text = Format((CDbl(dato12.text) * CDbl(dato14.text)) / 100, "###,###,###,##0")
    Else
    dato14.text = Format(CDbl(dato14.text), "###,###,###,##0.0000")
    dato15.text = Format((CDbl(dato12.text) * CDbl(dato14.text)) / 100, "###,###,###,##0.0000")
    End If
End If

End Sub

Private Sub dato15_KeyPress(KeyAscii As Integer)
    snum = 0: KeyAscii = esNumero(KeyAscii)
    If KeyAscii = 13 Then
    If dato2.text = "01" Then
    dato15.text = Format(CDbl(dato15.text), "###,###,###,##0")
    dato16.text = Format(CDbl(dato12.text) + CDbl(dato15.text), "###,###,###,##0")
    Else
    dato15.text = Format(CDbl(dato15.text), "###,###,###,##0.0000")
    dato16.text = Format(CDbl(dato12.text) + CDbl(dato15.text), "###,###,###,##0.0000")
    
    End If
    
    Call Pregunta(dato15, dato16)
    End If
End Sub
Private Sub dato16_KeyPress(KeyAscii As Integer)
    snum = 0: KeyAscii = esNumero(KeyAscii)
    If KeyAscii = 13 Then
    Call Pregunta(dato16, dato17)
    If dato2.text = "01" Then
    dato16.text = Format(CDbl(dato16.text), "###,###,###,##0")
    Else
    dato16.text = Format(CDbl(dato16.text), "###,###,###,##0.0000")
    End If
    End If

End Sub
Private Sub dato17_KeyPress(KeyAscii As Integer)
    snum = 1: KeyAscii = esNumero(KeyAscii)
    If KeyAscii = 13 And dato17.text <> "" Then
    MONTOPESOS.text = Format(CDbl(dato12.text) * CDbl(dato17.text), "###,###,###,##0")
    MONTOPESOS.SetFocus
    
    Else
    dato17.SetFocus
    
    End If
    
    
End Sub



Sub flechas(ByRef caja As TextBox, ByRef caja1 As TextBox, ByRef codigo As Integer)
    If codigo = 38 And caja.Enabled = True Then caja.SetFocus: caja.SelStart = 0
    If codigo = 40 And caja1.Enabled = True Then caja1.SetFocus: caja.SelStart = 0
End Sub
Sub grabar()
   
    campos(0, 0) = "banco"
    campos(1, 0) = "tipo"
    campos(2, 0) = "numero"
    campos(3, 0) = "empresa"
    campos(4, 0) = "fechaemision"
    campos(5, 0) = "fechavencimiento"
    campos(6, 0) = "dias"
    campos(7, 0) = "d_monto"
    campos(8, 0) = "d_tazaanual"
    campos(9, 0) = "d_tazaperiodo"
    campos(10, 0) = "d_interes"
    campos(11, 0) = "d_montoacobrar"
    campos(12, 0) = "d_tipocambio"
    campos(13, 0) = "p_monto"
    campos(14, 0) = "p_interes"
    campos(15, 0) = "p_montocobrar"
    campos(16, 0) = "p_tipocambio"
    campos(17, 0) = "p_montocambio"
    campos(18, 0) = "p_diferencia"
    campos(19, 0) = "contadeposito"
    campos(20, 0) = "contavencimiento"
    campos(21, 0) = ""
    
    
    
    campos(0, 1) = dato1.text
    campos(1, 1) = dato2.text
    campos(2, 1) = dato3.text
    campos(3, 1) = dato4.text
    campos(4, 1) = dato7.text + "-" + dato6.text + "-" + DATO5.text
    campos(5, 1) = dato10.text + "-" + dato9.text + "-" + dato8.text
    campos(6, 1) = dato11.text
    campos(7, 1) = Replace(CDbl(dato12.text), ",", ".")
    campos(8, 1) = Replace(CDbl(dato13.text), ",", ".")
    campos(9, 1) = Replace(CDbl(dato14.text), ",", ".")
    campos(10, 1) = Replace(CDbl(dato15.text), ",", ".")
    campos(11, 1) = Replace(CDbl(dato16.text), ",", ".")
    campos(12, 1) = Replace(CDbl(dato17.text), ",", ".")
    campos(13, 1) = Replace(CDbl(dato18.text), ",", ".")
    campos(14, 1) = Replace(CDbl(dato19.text), ",", ".")
    campos(15, 1) = Replace(CDbl(DATO20.text), ",", ".")
   If DATO21.text = "" Then DATO21.text = "0"
   If DATO22.text = "" Then DATO22.text = "0"
   If dato23.text = "" Then dato23.text = "0"
    
    campos(16, 1) = Replace(CDbl(DATO21.text), ",", ".")
   
    campos(17, 1) = Replace(CDbl(DATO22.text), ",", ".")
    campos(18, 1) = Replace(CDbl(dato23.text), ",", ".")
    campos(19, 1) = "0"
    campos(20, 1) = "0"
    
    If MODIFI = "1" Then
    condicion = "banco='" + dato1.text + "' and tipo='" + dato2.text + "' and numero='" + dato3.text + "' and empresa='" + dato4.text + "' "
    op = 3
    
    Else
    op = 2
    
    condicion = ""
    End If
    campos(0, 2) = "inver_movimientos"
    sqlconta.response = campos
    Set sqlconta.conexion = conta
    Call sqlconta.sqlconta(op, condicion)
    
MODIFI = "0"
leer

End Sub


Sub leer()

    campos(0, 0) = "banco"
    campos(1, 0) = "tipo"
    campos(2, 0) = "numero"
    campos(3, 0) = "empresa"
    campos(4, 0) = "fechaemision"
    campos(5, 0) = "fechavencimiento"
    campos(6, 0) = "dias"
    campos(7, 0) = "d_monto"
    campos(8, 0) = "d_tazaanual"
    campos(9, 0) = "d_tazaperiodo"
    campos(10, 0) = "d_interes"
    campos(11, 0) = "d_montoacobrar"
    campos(12, 0) = "d_tipocambio"
    campos(13, 0) = "p_monto"
    campos(14, 0) = "p_interes"
    campos(15, 0) = "p_montocobrar"
    campos(16, 0) = "p_tipocambio"
    campos(17, 0) = "p_montocambio"
    campos(18, 0) = "p_diferencia"
    campos(19, 0) = "contadeposito"
    campos(20, 0) = "contavencimiento"
    campos(21, 0) = ""
    
    
    
    condicion = "banco='" + dato1.text + "' and tipo='" + dato2.text + "' and numero='" + dato3.text + "' and empresa='" + dato4.text + "' "
    campos(0, 2) = "inver_movimientos"
    op = 5
    sqlconta.response = campos
    Set sqlconta.conexion = conta
    Call sqlconta.sqlconta(op, condicion)
        
    If sqlconta.status = 0 Then
    MODIFI = "1"
    carga
    opciones.Visible = True
    disponible (True)
    opciones.SetFocus
    contae.Enabled = True
    anulae.Enabled = False
    contadepo.Caption = "DEPOSITO NO CONTABILIZADO"
    If depositocontabilizado(dato3.text, dato7.text + "-" + dato6.text + "-" + DATO5.text, dato4.text, "DP") = True Then
    contadepo.Caption = "DEPOSITO CONTABILIZADO"
    contae.Enabled = False
    anulae.Enabled = True
    End If
    
    contav.Enabled = True
    anulav.Enabled = False
    CONTACOBRO.Caption = "RESCATE NO CONTABILIZADO"
    If depositocontabilizado(dato3.text, dato10.text + "-" + dato9.text + "-" + dato8.text, dato4.text, "RD") = True Then
    CONTACOBRO.Caption = "RESCATE CONTABILIZADO"
    contav.Enabled = False
    anulav.Enabled = True
    End If
    
    
    End If

End Sub
Sub carga()
    disponible (True)
'    campos(0, 1) = dato1.text
'    campos(1, 1) = dato2.text
'    campos(2, 1) = dato3.text
'    campos(3, 1) = dato4.text
'    campos(4, 1) = dato7.text + "-" + dato6.text + "-" + dato5.text
'    campos(5, 1) = dato10.text + "-" + dato9.text + "-" + dato8.text
'    campos(6, 1) = dato11.text
'    campos(7, 1) = Replace(CDbl(dato12.text), ",", ".")
'    campos(8, 1) = Replace(CDbl(dato13.text), ",", ".")
'    campos(9, 1) = Replace(CDbl(dato14.text), ",", ".")
'    campos(10, 1) = Replace(CDbl(dato15.text), ",", ".")
'    campos(11, 1) = Replace(CDbl(dato16.text), ",", ".")
'    campos(12, 1) = Replace(CDbl(dato17.text), ",", ".")
'    campos(13, 1) = Replace(CDbl(dato18.text), ",", ".")
'    campos(14, 1) = Replace(CDbl(dato19.text), ",", ".")
'    campos(15, 1) = Replace(CDbl(dato20.text), ",", ".")
'    campos(16, 1) = "0"
'    campos(17, 1) = "0"
'    campos(18, 1) = "0"
'    campos(19, 1) = "0"
'    campos(20, 1) = "0"
    
    dato1.text = sqlconta.response(0, 3)
    dato2.text = sqlconta.response(1, 3)
    dato3.text = sqlconta.response(2, 3)
    dato4.text = sqlconta.response(3, 3)
    DATO5.text = Mid(sqlconta.response(4, 3), 1, 2)
    dato6.text = Mid(sqlconta.response(4, 3), 4, 2)
    dato7.text = Mid(sqlconta.response(4, 3), 7, 4)
    dato8.text = Mid(sqlconta.response(5, 3), 1, 2)
    dato9.text = Mid(sqlconta.response(5, 3), 4, 2)
    dato10.text = Mid(sqlconta.response(5, 3), 7, 4)
    dato11.text = sqlconta.response(6, 3)
    If dato2.text = "01" Then
    dato12.text = Format(sqlconta.response(7, 3), "###,###,###")
    dato13.text = Format(sqlconta.response(8, 3), "###,###,##0.000000")
    dato14.text = Format(sqlconta.response(9, 3), "###,###,##0.000000")
    dato15.text = Format(sqlconta.response(10, 3), "###,###,###")
    dato16.text = Format(sqlconta.response(11, 3), "###,###,###")
    dato17.text = Format(sqlconta.response(12, 3), "###,###,###,##0")
    Else
    dato12.text = Format(sqlconta.response(7, 3), "###,###,##0.0000")
    dato13.text = Format(sqlconta.response(8, 3), "###,###,##0.0000")
    dato14.text = Format(sqlconta.response(9, 3), "###,###,##0.0000")
    dato15.text = Format(sqlconta.response(10, 3), "###,###,##0.0000")
    dato16.text = Format(sqlconta.response(11, 3), "###,###,##0.0000")
    dato17.text = Format(sqlconta.response(12, 3), "###,###,###,##0.0000")
    
    End If
    
    dato18.text = Format(sqlconta.response(13, 3), "###,###,###,##0")
    dato19.text = Format(sqlconta.response(14, 3), "###,###,###,##0")
    DATO20.text = Format(sqlconta.response(15, 3), "###,###,###,##0")
    DATO21.text = Format(sqlconta.response(16, 3), "###,###,##0.00")
    DATO22.text = Format(sqlconta.response(17, 3), "###,###,###,##0")
    dato23.text = Format(sqlconta.response(18, 3), "###,###,###,##0")
    
    
    Call leerbanco(dato1.text)
    Call leerdeposito(dato2.text)
    Call leerempresa(dato4.text)
    
End Sub



Sub ELIMINAR()
    condicion = "banco='" + dato1.text + "' and tipo='" + dato2.text + "' and numero='" + dato3.text + "' and empresa='" + dato4.text + "' "
    campos(0, 2) = "inver_movimientos"
    
    op = 4
    sqlconta.response = campos
    Set sqlconta.conexion = conta
    Call sqlconta.sqlconta(op, condicion)
no:
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


Private Sub MONTOPESOS_KeyPress(KeyAscii As Integer)
KeyAscii = esNumero(KeyAscii)

If KeyAscii = 13 Then

dato18.text = Format(MONTOPESOS.text, "###,###,###,###")
dato19.text = "0"
DATO20.text = "0"
    
    Command3.SetFocus
    

End If

End Sub

Private Sub opciones_FSCommand(ByVal command As String, ByVal args As String)
If command = "retorno" Then retorno

If command = "modifica" Then
    If contadepo.Caption <> "DEPOSITO CONTABILIZADO" And CONTACOBRO.Caption <> "RESCATE CONTABILIZADO" Then
    MODIFI = "1"
    dato1.Enabled = True
    dato1.SetFocus
Else
MsgBox ("imposible modificar deposito contabilizado")

    End If
End If

If command = "elimina" Then
If contadepo.Caption <> "DEPOSITO CONTABILIZADO" And CONTACOBRO.Caption <> "RESCATE CONTABILIZADO" Then
    
    ELIMINAR
    retorno
Else
MsgBox ("imposible eliminar deposito contabilizado")

End If

End If

End Sub


Sub retorno()


opciones.Visible = False
limpia
disponible (False)

dato1.Enabled = True
dato1.SetFocus

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
    DATO21.text = "0"
    DATO22.text = ""
    dato23.text = ""
    
    lblBanco.Caption = ""
    LBLEMPRESA.Caption = ""
    LBLTIPO.Caption = ""
    
    
End Sub


Sub Pregunta(ByRef caja As TextBox, ByRef caja1 As TextBox)
    If caja.text = "" And sc = 0 Then caja.SetFocus: caja.SelStart = 0
    If caja.text <> "" Or sc = 1 Then caja1.Enabled = True: caja1.SetFocus
    sc = 0
End Sub

Sub cargatexto(ByRef caja As TextBox)


caja.SelStart = 0: caja.SelLength = Len(caja.text)

End Sub















Sub ayudacrcc(row As Long, col As Long)
    Dim campos As Variant
    Dim cfijo As Variant
    Dim largo As Variant
    campos = Array("codigo", "nombre")
    cabezas = Array("codigo", "nombre")
    largo = Array("8n", "40s")
    mensajeAyuda = "Ayuda Centros de costo"
    cfijo = "no"
    pivote2.MaxLength = 4
    Call cargaAyudaT(Servidor, basebus, Usuario, password, "centrosdecosto", pivote2, campos, cfijo, largo, 2)
    
    pivote2.text = ""
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
    
    
End Sub



Private Sub opciones_GotFocus()
    MANUAL.SetFocus
End Sub


Private Sub botonmisaccesos_Click()
programafiltro = Me.Caption
misaccesos.Show
End Sub


Private Sub botonmisfavoritos_Click()
Call AgregaFavoritos(USUARIOSISTEMA, App.EXEName, Me.Name, Me.Caption)
End Sub
