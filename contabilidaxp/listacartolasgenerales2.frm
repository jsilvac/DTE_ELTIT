VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{085FBDF3-00C5-421B-B762-1D57299A2B89}#1.0#0"; "CLBUTN.OCX"
Object = "{ADBBAED6-B16F-40DC-80DF-B44910CBA76C}#1.0#0"; "Frame-Xp.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form informa44 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Lista Cartolas de Movimientos"
   ClientHeight    =   8160
   ClientLeft      =   435
   ClientTop       =   825
   ClientWidth     =   10785
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8160
   ScaleWidth      =   10785
   Begin MSComctlLib.ProgressBar barra 
      Height          =   240
      Left            =   765
      TabIndex        =   47
      Top             =   7650
      Width           =   8475
      _ExtentX        =   14949
      _ExtentY        =   423
      _Version        =   393216
      Appearance      =   1
   End
   Begin TabDlg.SSTab sbtab1 
      Height          =   5655
      Left            =   720
      TabIndex        =   1
      Top             =   1920
      Width           =   9615
      _ExtentX        =   16960
      _ExtentY        =   9975
      _Version        =   393216
      TabHeight       =   520
      BackColor       =   -2147483635
      ForeColor       =   16576
      TabCaption(0)   =   "CUENTAS DEL MAYOR"
      TabPicture(0)   =   "listacartolasgenerales2.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "opcion1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "CUENTAS CORRIENTES"
      TabPicture(1)   =   "listacartolasgenerales2.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "opcion2"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "CENTROS DE COSTO"
      TabPicture(2)   =   "listacartolasgenerales2.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "opcion3"
      Tab(2).ControlCount=   1
      Begin XPFrame.FrameXp opcion1 
         Height          =   5175
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   9330
         _ExtentX        =   16457
         _ExtentY        =   9128
         BackColor       =   16761024
         Caption         =   ""
         CaptionEstilo3D =   1
         BackColor       =   16761024
         ForeColor       =   8438015
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
         Begin VB.CommandButton ctacte 
            Caption         =   "Saldos"
            Height          =   495
            Left            =   6600
            TabIndex        =   68
            Top             =   4560
            Width           =   2175
         End
         Begin VB.CheckBox Check2 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
            Caption         =   "Conciliado"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   330
            Left            =   240
            TabIndex        =   48
            Top             =   4455
            Width           =   2265
         End
         Begin CoolButtons.cool_Button command5 
            Height          =   495
            Left            =   2880
            TabIndex        =   3
            Top             =   4560
            Width           =   2775
            _ExtentX        =   4895
            _ExtentY        =   873
            SkinId          =   "6"
            Caption         =   "Genera Informe"
         End
         Begin XPFrame.FrameXp cuentas 
            Height          =   1215
            Left            =   120
            TabIndex        =   4
            Top             =   1800
            Width           =   7815
            _ExtentX        =   13785
            _ExtentY        =   2143
            BackColor       =   16761024
            Caption         =   "CODIGO DE CUENTA"
            CaptionEstilo3D =   1
            BackColor       =   16761024
            ForeColor       =   8438015
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
            Begin VB.TextBox cmdato3 
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
               Left            =   2400
               MaxLength       =   4
               TabIndex        =   7
               Top             =   360
               Width           =   735
            End
            Begin VB.TextBox cmdato2 
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
               Left            =   1920
               MaxLength       =   2
               TabIndex        =   6
               Top             =   360
               Width           =   375
            End
            Begin VB.TextBox cmdato1 
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
               Left            =   1440
               MaxLength       =   2
               TabIndex        =   5
               Tag             =   "codigo"
               Top             =   360
               Width           =   375
            End
            Begin VB.Label cmnombre 
               BorderStyle     =   1  'Fixed Single
               Height          =   330
               Left            =   1440
               TabIndex        =   44
               Top             =   720
               Width           =   5640
            End
            Begin VB.Label Label1 
               BackColor       =   &H00FFC0C0&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "Codigo Cuenta"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00400000&
               Height          =   255
               Left            =   120
               TabIndex        =   9
               Top             =   360
               Width           =   1215
            End
            Begin VB.Label Label9 
               BackColor       =   &H00FFC0C0&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "Nombre"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00400000&
               Height          =   255
               Left            =   120
               TabIndex        =   8
               Top             =   720
               Width           =   1215
            End
         End
         Begin XPFrame.FrameXp FrameXp1 
            Height          =   1455
            Left            =   45
            TabIndex        =   10
            Top             =   315
            Width           =   2775
            _ExtentX        =   4895
            _ExtentY        =   2566
            BackColor       =   16761024
            Caption         =   "Opciones"
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
            Alignment       =   1
            Begin VB.OptionButton cmindi 
               BackColor       =   &H00FFC0C0&
               Caption         =   "Una Cuenta Individual"
               Height          =   375
               Left            =   315
               TabIndex        =   12
               Top             =   960
               Value           =   -1  'True
               Width           =   1935
            End
            Begin VB.OptionButton cmtoda 
               BackColor       =   &H00FFC0C0&
               Caption         =   "Todas las Cuentas"
               Height          =   495
               Left            =   315
               TabIndex        =   11
               Top             =   360
               Width           =   1935
            End
         End
         Begin XPFrame.FrameXp frm_crcc 
            Height          =   1215
            Left            =   120
            TabIndex        =   55
            Top             =   3120
            Width           =   7815
            _ExtentX        =   13785
            _ExtentY        =   2143
            BackColor       =   16761024
            Caption         =   "Filtro centro de costo"
            CaptionEstilo3D =   1
            BackColor       =   16761024
            ForeColor       =   8438015
            ColorBarraAbajo =   16711680
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Begin VB.CommandButton Command1 
               Caption         =   "blanquear"
               Height          =   255
               Left            =   5760
               TabIndex        =   60
               Top             =   360
               Width           =   1575
            End
            Begin VB.TextBox txt_crcc 
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
               Left            =   1440
               MaxLength       =   4
               TabIndex        =   56
               Tag             =   "codigo"
               Top             =   360
               Width           =   735
            End
            Begin VB.Label Label7 
               BackStyle       =   0  'Transparent
               BorderStyle     =   1  'Fixed Single
               Caption         =   "Codigo Cuenta"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00400000&
               Height          =   255
               Left            =   120
               TabIndex        =   59
               Top             =   360
               Width           =   1095
            End
            Begin VB.Label Label3 
               BackStyle       =   0  'Transparent
               BorderStyle     =   1  'Fixed Single
               Caption         =   "Nombre"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00400000&
               Height          =   255
               Left            =   120
               TabIndex        =   58
               Top             =   720
               Width           =   1095
            End
            Begin VB.Label lbl_crcc 
               BorderStyle     =   1  'Fixed Single
               Height          =   330
               Left            =   1440
               TabIndex        =   57
               Top             =   720
               Width           =   5910
            End
         End
      End
      Begin XPFrame.FrameXp opcion3 
         Height          =   5175
         Left            =   -74880
         TabIndex        =   13
         Top             =   360
         Width           =   9435
         _ExtentX        =   16642
         _ExtentY        =   9128
         BackColor       =   16761024
         Caption         =   ""
         CaptionEstilo3D =   1
         BackColor       =   16761024
         ForeColor       =   8438015
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
         Begin XPFrame.FrameXp FrameXp3 
            Height          =   1215
            Left            =   240
            TabIndex        =   14
            Top             =   1920
            Width           =   7815
            _ExtentX        =   13785
            _ExtentY        =   2143
            BackColor       =   16761024
            Caption         =   "CODIGO CENTRO COSTO"
            CaptionEstilo3D =   1
            BackColor       =   16761024
            ForeColor       =   8438015
            ColorBarraAbajo =   16711680
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Begin VB.TextBox ccdato2 
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
               Left            =   1920
               MaxLength       =   2
               TabIndex        =   16
               Top             =   360
               Width           =   375
            End
            Begin VB.TextBox ccdato1 
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
               Left            =   1440
               MaxLength       =   2
               TabIndex        =   15
               Tag             =   "codigo"
               Top             =   360
               Width           =   375
            End
            Begin VB.Label ccnombre 
               BorderStyle     =   1  'Fixed Single
               Height          =   330
               Left            =   1440
               TabIndex        =   43
               Top             =   720
               Width           =   5910
            End
            Begin VB.Label Label12 
               BackStyle       =   0  'Transparent
               BorderStyle     =   1  'Fixed Single
               Caption         =   "Nombre"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00400000&
               Height          =   255
               Left            =   120
               TabIndex        =   18
               Top             =   720
               Width           =   1095
            End
            Begin VB.Label Label13 
               BackStyle       =   0  'Transparent
               BorderStyle     =   1  'Fixed Single
               Caption         =   "Codigo Cuenta"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00400000&
               Height          =   255
               Left            =   120
               TabIndex        =   17
               Top             =   360
               Width           =   1095
            End
         End
         Begin XPFrame.FrameXp FrameXp2 
            Height          =   1335
            Left            =   225
            TabIndex        =   19
            Top             =   315
            Width           =   2535
            _ExtentX        =   4471
            _ExtentY        =   2355
            BackColor       =   16761024
            Caption         =   "Opciones"
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
            Begin VB.OptionButton ccindi 
               BackColor       =   &H00FFC0C0&
               Caption         =   "Una Cuenta Individual"
               Height          =   375
               Left            =   120
               TabIndex        =   21
               Top             =   840
               Width           =   1935
            End
            Begin VB.OptionButton cctoda 
               BackColor       =   &H00FFC0C0&
               Caption         =   "Todas las Cuentas"
               Height          =   495
               Left            =   120
               TabIndex        =   20
               Top             =   360
               Width           =   1935
            End
         End
         Begin CoolButtons.cool_Button command6 
            Height          =   375
            Left            =   3120
            TabIndex        =   22
            Top             =   4680
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   661
            Caption         =   "GENERA INFORME"
         End
         Begin XPFrame.FrameXp FrameXp6 
            Height          =   1215
            Left            =   240
            TabIndex        =   61
            Top             =   3240
            Width           =   7815
            _ExtentX        =   13785
            _ExtentY        =   2143
            BackColor       =   16761024
            Caption         =   "CODIGO DE CUENTA"
            CaptionEstilo3D =   1
            BackColor       =   16761024
            ForeColor       =   8438015
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
            Begin VB.TextBox txtdato1 
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
               Left            =   1440
               MaxLength       =   2
               TabIndex        =   64
               Tag             =   "codigo"
               Top             =   360
               Width           =   375
            End
            Begin VB.TextBox txtdato2 
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
               Left            =   1920
               MaxLength       =   2
               TabIndex        =   63
               Top             =   360
               Width           =   375
            End
            Begin VB.TextBox txtdato3 
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
               Left            =   2400
               MaxLength       =   4
               TabIndex        =   62
               Top             =   360
               Width           =   735
            End
            Begin VB.Label Label10 
               BackColor       =   &H00FFC0C0&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "Nombre"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00400000&
               Height          =   255
               Left            =   120
               TabIndex        =   67
               Top             =   720
               Width           =   1215
            End
            Begin VB.Label Label8 
               BackColor       =   &H00FFC0C0&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "Codigo Cuenta"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00400000&
               Height          =   255
               Left            =   120
               TabIndex        =   66
               Top             =   360
               Width           =   1215
            End
            Begin VB.Label lblnombrecuenta 
               BorderStyle     =   1  'Fixed Single
               Height          =   330
               Left            =   1440
               TabIndex        =   65
               Top             =   720
               Width           =   5640
            End
         End
      End
      Begin XPFrame.FrameXp opcion2 
         Height          =   5130
         Left            =   -74880
         TabIndex        =   23
         Top             =   360
         Width           =   9450
         _ExtentX        =   16669
         _ExtentY        =   9049
         BackColor       =   16761024
         Caption         =   ""
         CaptionEstilo3D =   1
         BackColor       =   16761024
         ForeColor       =   8438015
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
         Begin VB.CheckBox Check3 
            BackColor       =   &H00FFC0C0&
            Caption         =   "No Mostrar Saldos en Cero"
            Height          =   195
            Left            =   360
            TabIndex        =   51
            Top             =   4680
            Visible         =   0   'False
            Width           =   2295
         End
         Begin VB.OptionButton Option2 
            BackColor       =   &H00FFC0C0&
            Caption         =   "Orden x Numeo"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   5640
            TabIndex        =   50
            Top             =   4680
            Value           =   -1  'True
            Width           =   2535
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00FFC0C0&
            Caption         =   "Orden x Fecha"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   5640
            TabIndex        =   49
            Top             =   4320
            Width           =   2295
         End
         Begin VB.CheckBox Check1 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
            Caption         =   "Conciliado"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   330
            Left            =   360
            TabIndex        =   46
            Top             =   4335
            Width           =   2265
         End
         Begin XPFrame.FrameXp FrameXp5 
            Height          =   1815
            Left            =   90
            TabIndex        =   24
            Top             =   2385
            Width           =   7815
            _ExtentX        =   13785
            _ExtentY        =   3201
            BackColor       =   16761024
            Caption         =   "INGRESO DE RUT"
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
            Begin VB.TextBox ctdato1 
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
               Left            =   1560
               MaxLength       =   8
               TabIndex        =   27
               Tag             =   "tipo"
               Top             =   360
               Width           =   1095
            End
            Begin VB.TextBox ctdato2 
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
               Left            =   1575
               MaxLength       =   9
               TabIndex        =   26
               Tag             =   "rut"
               Top             =   720
               Width           =   1095
            End
            Begin VB.TextBox dv 
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
               Left            =   2700
               MaxLength       =   2
               TabIndex        =   25
               Tag             =   "tipo"
               Top             =   720
               Width           =   255
            End
            Begin VB.Label nombrectacte 
               BorderStyle     =   1  'Fixed Single
               Height          =   330
               Left            =   2700
               TabIndex        =   45
               Top             =   360
               Width           =   4695
            End
            Begin VB.Label ctnombre 
               BackStyle       =   0  'Transparent
               BorderStyle     =   1  'Fixed Single
               Height          =   285
               Left            =   1575
               TabIndex        =   42
               Top             =   1080
               Width           =   5865
            End
            Begin VB.Label Label6 
               BackColor       =   &H00FFC0C0&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "Rut"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00400000&
               Height          =   255
               Left            =   240
               TabIndex        =   30
               Top             =   720
               Width           =   1215
            End
            Begin VB.Label Label5 
               BackColor       =   &H00FFC0C0&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "Tipo Cuenta"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00400000&
               Height          =   255
               Left            =   240
               TabIndex        =   29
               Top             =   360
               Width           =   1215
            End
            Begin VB.Label Label4 
               BackColor       =   &H00FFC0C0&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "Nombre"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00400000&
               Height          =   255
               Left            =   240
               TabIndex        =   28
               Top             =   1080
               Width           =   1215
            End
         End
         Begin XPFrame.FrameXp FrameXp4 
            Height          =   1935
            Left            =   90
            TabIndex        =   31
            Top             =   315
            Width           =   2775
            _ExtentX        =   4895
            _ExtentY        =   3413
            BackColor       =   16761024
            Caption         =   "Opciones"
            CaptionEstilo3D =   1
            BackColor       =   16761024
            ForeColor       =   1563884
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Begin VB.OptionButton cttodatipo 
               BackColor       =   &H00FFC0C0&
               Caption         =   "Todas las Cuentas de Un tipo"
               Height          =   495
               Left            =   120
               TabIndex        =   34
               Top             =   840
               Width           =   2535
            End
            Begin VB.OptionButton ctindi 
               BackColor       =   &H00FFC0C0&
               Caption         =   "Una Cuenta corriente Individual"
               Height          =   375
               Left            =   90
               TabIndex        =   33
               Top             =   1320
               Width           =   2535
            End
            Begin VB.OptionButton cttoda 
               BackColor       =   &H00FFC0C0&
               Caption         =   "Todos los Tipos"
               Height          =   495
               Left            =   120
               TabIndex        =   32
               Top             =   360
               Width           =   1935
            End
         End
         Begin CoolButtons.cool_Button command2 
            Height          =   495
            Left            =   3000
            TabIndex        =   35
            Top             =   4410
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   873
            Caption         =   "GENERA INFORME"
         End
      End
   End
   Begin VB.TextBox pivote 
      Height          =   285
      Left            =   7560
      MaxLength       =   8
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   1440
      Visible         =   0   'False
      Width           =   1575
   End
   Begin XPFrame.FrameXp FrameXp8 
      Height          =   1695
      Left            =   2520
      TabIndex        =   36
      Top             =   120
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   2990
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
      Begin CoolButtons.cool_Button cool_Button3 
         Height          =   375
         Left            =   1680
         TabIndex        =   37
         Top             =   1200
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
         TabIndex        =   41
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
         TabIndex        =   40
         Top             =   720
         Width           =   1935
      End
      Begin VB.Label Label17 
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
         TabIndex        =   39
         Top             =   360
         Width           =   1935
      End
      Begin VB.Label Label16 
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
         TabIndex        =   38
         Top             =   360
         Width           =   1935
      End
   End
   Begin XPFrame.FrameXp FrameQuickMenu 
      Height          =   615
      Left            =   7500
      TabIndex        =   52
      Top             =   20
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
      Begin VB.CommandButton botonmisfavoritos 
         Caption         =   "Mis Favoritos"
         Height          =   255
         Left            =   1800
         TabIndex        =   54
         Top             =   280
         Width           =   1335
      End
      Begin VB.CommandButton botonmisaccesos 
         Caption         =   "Permisos Modulo"
         Height          =   255
         Left            =   120
         TabIndex        =   53
         Top             =   280
         Width           =   1455
      End
   End
End
Attribute VB_Name = "informa44"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private FORMATOGRILLA(20, 20)
Private lin As Double
Private saldo As Double
Private dedonde As Integer
Private tipoctacte As String





Private Sub busca_Click()

End Sub

Private Sub ccdato1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF2 Then Call ayudacrcc(ccdato1, ccdato2)
End Sub
Sub ayudacrcc(primero As TextBox, segundo As TextBox)
    Dim campos As Variant
    Dim cfijo As Variant
    Dim largo As Variant
    campos = Array("codigo", "nombre")
    cabezas = Array("codigo", "nombre")
    largo = Array("8n", "40s")
    mensajeAyuda = "Ayuda Centros de costo"
    cfijo = "año='" + año + "'"
    pivote.MaxLength = 4
    Call cargaAyudaT(Servidor, basebus, Usuario, password, "centrosdecosto", pivote, campos, cfijo, largo, 2)
    primero.text = Mid(pivote.text, 1, 2)
    segundo.text = Mid(pivote.text, 3, 2)
    pivote.text = ""
End Sub

Private Sub ccdato1_KeyPress(KeyAscii As Integer)
    KeyAscii = esNumero(KeyAscii)
    If KeyAscii = 13 Then
        Call ceros(ccdato1)
        ccdato2.SetFocus
    End If
End Sub

Private Sub ccdato2_KeyPress(KeyAscii As Integer)
    KeyAscii = esNumero(KeyAscii)
    If KeyAscii = 13 Then
        ccnombre.Caption = leerNOMBREcrcc(ccdato1.text & ccdato2.text)
        If ccnombre.Caption <> "" Then
            txtdato1.SetFocus
        End If
    End If
End Sub

Private Sub Check1_Click()
If Check1.Value = 1 Then
Check3.Visible = True
Else
Check3.Visible = False
End If
End Sub

Private Sub cmdato1_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = vbKeyF2 Then Call ayudamayor(cmdato1)

End Sub

Private Sub cmde01_Change()

End Sub

Private Sub cmde01_KeyDown(KeyCode As Integer, Shift As Integer)

End Sub

Private Sub cmdato1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Call ceros(cmdato1)
cmdato2.SetFocus
End If

End Sub
Private Sub cmdato2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Call ceros(cmdato2)
cmdato3.SetFocus
End If
End Sub
Private Sub cmdato3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Call ceros(cmdato3)
If leercuenta(cmdato1.text + cmdato2.text + cmdato3.text) <> "" Then
cmnombre.Caption = leercuenta(cmdato1.text + cmdato2.text + cmdato3.text)
Command5.SetFocus
If cmdato1.text > "29" Then
frm_crcc.Visible = True
Else
frm_crcc.Visible = False

End If

Else
MsgBox ("codigo de cuenta madre no existe")


cmdato3.SetFocus
End If
End If
End Sub

Private Sub cmindi_Click()
cuentas.Visible = True

End Sub

Private Sub cmtoda_Click()
cuentas.Visible = False

End Sub


Private Sub Command10_Click()
fechas.Visible = False



End Sub

Private Sub Command11_Click()
fechas.Visible = True

opcion1.Visible = True
opcion2.Visible = False
opcion3.Visible = False
End Sub
Private Sub Command12_Click()
fechas.Visible = True

opcion1.Visible = False
opcion2.Visible = True
opcion3.Visible = False
End Sub
Private Sub Command13_Click()
fechas.Visible = True

opcion1.Visible = False
opcion2.Visible = False
opcion3.Visible = True
End Sub

Private Sub Command1_Click()
txt_crcc.text = ""
lbl_crcc.Caption = ""
End Sub

Private Sub COMMAND2_Click()
lin = 0
dedonde = 2
Call ACEPTA(dedonde)
End Sub

Private Sub Command5_Click()
If cmindi = True And cmnombre.Caption <> "" Then
lin = 0
dedonde = 1
Call ACEPTA(dedonde)
End If
If cmindi = False Then
lin = 0
dedonde = 1
Call ACEPTA(dedonde)
End If





End Sub
Private Function leercuenta(cuenta) As String


If PermisosCuentasDelMayor(USUARIOSISTEMA, Format(cuenta, "00000000")) = False Then
    MsgBox "USTED NO TIENE PRIVILEGIOS PARA ACCEDER A ESTA CUENTA", vbCritical, "ATENCION"
cmdato1.SetFocus
 
  Exit Function
End If


      campos(0, 0) = "nombre"
      campos(1, 0) = ""
      campos(0, 2) = "cuentasdelmayor"
      condicion = "codigo=" + "'" + cuenta + "' and año='" + Format(fechasistema, "yyyy") + "'"
      op = 5
      sqlconta.response = campos
      Set sqlconta.conexion = contadb
      Call sqlconta.sqlconta(op, condicion)
        If sqlconta.status = 0 Then
        leercuenta = sqlconta.response(0, 3)
        
        Else
        leercuenta = ""
        
        End If
        

End Function
Private Function leerctacte(tipo, rut) As String

      campos(0, 0) = "nombre"
      campos(1, 0) = ""
      campos(0, 2) = "cuentascorrientes"
      condicion = "tipo=" + "'" + tipo + "' and rut='" + rut + "' and año='" + Format(fechasistema, "yyyy") + "'"
      op = 5
      sqlconta.response = campos
      Set sqlconta.conexion = contadb
      Call sqlconta.sqlconta(op, condicion)
        If sqlconta.status = 0 Then
        leerctacte = sqlconta.response(0, 3)
        
        Else
        leerctacte = ""
        
        End If
        

End Function

Sub ACEPTA(opcion)
Dim fecha1 As String
Dim fecha2 As String
Dim infogrilla As grillainformes
Set infogrilla = New grillainformes
Call CARGAGRILLA(infogrilla)
If opcion = 1 Then infogrilla.Caption = "CARTOLA CUENTAS DEL MAYOR"
If opcion = 2 Then infogrilla.Caption = "CARTOLA CUENTAS CORRIENTES"
If opcion = 3 Then infogrilla.Caption = "CARTOLA CENTROS DE COSTO"
If opcion = 1 Then Call leecuentas(infogrilla)
If opcion = 2 Then Call leercuentascorrientes(infogrilla)
If opcion = 3 Then Call leecrcc(infogrilla)
If opcion = 1 Then
fecha1 = Format(desdefecha.Caption, "dd-mm-yyyy")
fecha2 = Format(hastafecha.Caption, "dd-mm-yyyy")

infogrilla.Caption = "CARTOLA CUENTAS DEL MAYOR"
grillainformes.Tag = "CARTOLAMAYOR" & "0" & "0000000000"
infogrilla.cabeza.Caption = "CARTOLA CUENTA DEL MAYOR del " & fecha1 & " al " & fecha2 & " "
End If
If opcion = 2 Then
fecha1 = Format(desdefecha.Caption, "dd-mm-yyyy")
fecha2 = Format(hastafecha.Caption, "dd-mm-yyyy")
infogrilla.Caption = "CARTOLA CUENTAS CORRIENTES"
grillainformes.Tag = "CARTOLACTACTE" & "0" & "0000000000"
infogrilla.cabeza.Caption = "CARTOLA CUENTA CORRIENTE del " & fecha1 & " al " & fecha2 & " "
End If
infogrilla.Grid1.Visible = True
infogrilla.Show
End Sub

Private Sub Command6_Click()
lin = 0
dedonde = 3
Call ACEPTA(dedonde)
End Sub


Private Sub command8_Click()
Call retornofecha(desdefecha, hastafecha)
End Sub

Private Sub Command9_Click()
If dedonde = 1 Then
    If cmtoda.Value = True Then Call ACEPTA(1)
    If cmindi.Value = True Then Call ACEPTA(1)
End If
fechas.Visible = False
End Sub
Private Sub cool_Button3_Click()
Call retornofecha(desdefecha, hastafecha)
End Sub

Private Sub ctacte_Click()
If leertiene(cmdato1.text + cmdato2.text + cmdato3.text, 1) = True Then
informa01.dato1.text = cmdato1.text + cmdato2.text + cmdato3.text
informa01.Label1.Caption = cmnombre.Caption
informa01.Show
Else
MsgBox "CUENTA NO TIENE ANALISIS CUENTA CORRIENTE"
End If

End Sub

Private Sub ctdato1_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = vbKeyF2 Then Call ayudatipocuenta(ctdato1)

End Sub

Private Sub ctdato1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
If leercuenta(ctdato1.text) <> "" Then
nombrectacte.Caption = leercuenta(ctdato1.text)
ctdato2.SetFocus

Else
MsgBox ("codigo de cuenta no existe")
ctdato1.text = ""
ctdato1.SetFocus

End If
End If

End Sub

Private Sub ctdato2_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = vbKeyF2 Then Call ayudactacte(ctdato2)
End Sub

Private Sub ctdato2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Call ceros(ctdato2)
DV.text = rut(ctdato2.text)
If leerctacte(ctdato1.text, ctdato2.text + DV.text) <> "" Then
ctnombre.Caption = leerctacte(ctdato1.text, ctdato2.text + DV.text)
Command2.SetFocus

Else

MsgBox ("rut no existe")
ctdato2.text = ""
DV.text = ""

ctdato2.SetFocus

End If


End If


End Sub

Private Sub Form_Activate()
If Mid(grillainformes.cabeza.Caption, 1, 6) = "ESTADO" Then
Call Command5_Click

End If


End Sub

Private Sub Form_Load()
Call CENTRAR(Me)
frm_crcc.Visible = False

ctindi.Value = True
ccindi.Value = True
cmindi.Value = True
'fechas.Visible = True





    Call Conectar_BD
    Call Conectarconta(Servidor, clientesistema + "conta", Usuario, password)
    'CARGAGRILLA
    'leecuentas

desdefecha.Caption = "01-01-" + Format(fechasistema, "yyyy")
hastafecha.Caption = fechasistema

lin = 0

'If Mid(cartola, 1, 5) = "mayor" Then
'    cmdato1.text = Mid(cartola, 7, 2)
'    cmdato2.text = Mid(cartola, 9, 2)
'    cmdato3.text = Mid(cartola, 11, 4)
'
'    cmindi = True
'
''Call Command11_Click
'
'End If


End Sub


    
Sub LEERMOVIMIENTOS(infogrilla As grillainformes, cuenta, NOMBRE)

Dim resultados As rdoResultset
    Dim csql As New rdoQuery
    Dim rut As String
    Dim PASO As String
    Dim fecha1 As String
    Dim fecha2 As String
    
        fecha1 = Mid(desdefecha.Caption, 7, 4) + "-" + Mid(desdefecha.Caption, 4, 2) + "-" + Mid(desdefecha.Caption, 1, 2)
        fecha2 = Mid(hastafecha.Caption, 7, 4) + "-" + Mid(hastafecha.Caption, 4, 2) + "-" + Mid(hastafecha.Caption, 1, 2)
    
        Set csql.ActiveConnection = contadb
        If dedonde = 1 And Check2.Value = 1 Then
        csql.sql = "SELECT fecha,tipo,numero,linea,codigocuenta,glosacontable,tipodocumento,numerodocumento,fechadocumento,fechavencimiento,monto,dh,centrocosto,tipoctacte,rutctacte,count(monto) "
        Else
        csql.sql = "SELECT fecha,tipo,numero,linea,codigocuenta,glosacontable,tipodocumento,numerodocumento,fechadocumento,fechavencimiento,monto,dh,centrocosto,tipoctacte,rutctacte "
        End If
        
        If dedonde = 1 Then csql.sql = csql.sql + "FROM movimientoscontables where codigocuenta='" + cuenta + "' and fecha>='" + fecha1 + "' and fecha<='" + fecha2 + "' "
        If dedonde = 2 Then csql.sql = csql.sql + "FROM movimientoscontables where codigocuenta='" + tipoctacte + "' and rutctacte='" + cuenta + "' and fecha>='" + fecha1 + "' and fecha<='" + fecha2 + "' "
        If dedonde = 3 Then
        csql.sql = csql.sql + "FROM movimientoscontables where "
        csql.sql = csql.sql & "MID(codigocuenta,1,1)>2 and "
        csql.sql = csql.sql & "centrocosto='" + cuenta + "' and fecha>='" + fecha1 + "' and fecha<='" + fecha2 + "' "
        If LblnombreCuenta.Caption <> "" Then
            csql.sql = csql.sql & " and codigocuenta='" & txtdato1.text & txtdato2.text & txtdato3.text & "' "
        End If
        End If
        If lbl_crcc.Caption <> "" Then
        csql.sql = csql.sql + " and centrocosto='" + txt_crcc.text + "' "
        End If
        
        If dedonde = 1 And Check2.Value = 0 Then
        csql.sql = csql.sql + "order by codigocuenta,fecha,tipo,numero,linea"
        End If
        If dedonde = 1 And Check2.Value = 1 Then
        csql.sql = csql.sql + "group by monto having count(monto)<>'2' order by codigocuenta,fecha,tipo,numero,linea"
        End If
        
        If dedonde = 2 And Check1.Value = 0 Then
            
            If Option1.Value = True Then
            csql.sql = csql.sql + "order by codigocuenta,rutctacte,fecha,tipo,numero,linea"
            Else
            csql.sql = csql.sql + "order by numerodocumento"
            End If
            
        End If
        If dedonde = 2 And Check1.Value = 1 Then csql.sql = csql.sql + "group by tipodocumento,numerodocumento,monto,rutctacte,dh having count(numerodocumento)<>'2' order by tipoctacte,rutctacte,fecha,tipo,numero,linea"
        If dedonde = 3 Then csql.sql = csql.sql + "order by centrocosto,fecha,tipo,numero,linea"
        
        csql.Execute
        
        
        If dedonde <> 2 Then Call DATOSSALDOS(cuenta)
        If dedonde = 2 Then Call DATOSSALDOSctacte(cuenta)
        For k = 1 To 6
        infogrilla.Grid1.Column(k).Locked = False
        Next k
        
        
        If saldo <> 0 Or csql.RowsAffected <> 0 Then
        lin = lin + 1
        infogrilla.Grid1.Rows = infogrilla.Grid1.Rows + 1
                
        infogrilla.Grid1.Range(lin, 1, lin, 6).Merge
      
        infogrilla.Grid1.Cell(lin, 1).CellType = cellTextBox
        
        infogrilla.Grid1.Cell(lin, 10).CellType = cellTextBox
        If dedonde = 2 Then
        infogrilla.Grid1.Cell(lin, 1).text = Format(Mid(cuenta, 1, 9), "###,###,###") & "-" & Mid(cuenta, 10, 1) & " " + NOMBRE
        End If
        If dedonde = 1 Then
        infogrilla.Grid1.Cell(lin, 1).text = cuenta & " " + NOMBRE + "  " + lbl_crcc.Caption
        End If
        
        If dedonde = 2 Then infogrilla.Grid1.Cell(lin, 6).text = nombrectacte
        infogrilla.Grid1.Cell(lin, 10).text = "SALDO-->"
        
        infogrilla.Grid1.Cell(lin, 13).text = saldo
        infogrilla.Grid1.Range(lin, 0, lin, infogrilla.Grid1.Cols - 1).FontBold = True
        infogrilla.Grid1.Range(lin, 0, lin, infogrilla.Grid1.Cols - 1).FontUnderline = True
        
        
        End If
        
        If csql.RowsAffected > 0 Then
        
        
        Set resultados = csql.OpenResultset
        
         While Not resultados.EOF
        If dedonde = 1 And Check2.Value = 1 Then
        If resultados(15) > 2 Then GoTo dale:
        End If
          lin = lin + 1
             infogrilla.Grid1.Rows = infogrilla.Grid1.Rows + 1
             infogrilla.Grid1.Cell(lin, 0).text = resultados("rutctacte")
             For k = 0 To 9
             infogrilla.Grid1.Cell(lin, k + 1).text = resultados(k)
             Next k
             If resultados(11) = "D" Then infogrilla.Grid1.Cell(lin, 11).text = resultados(10): anted = anted + resultados(10): saldo = saldo + resultados(10)
             If resultados(11) = "H" Then infogrilla.Grid1.Cell(lin, 12).text = resultados(10): anteh = anteh + resultados(10): saldo = saldo - resultados(10)
             infogrilla.Grid1.Cell(lin, 5).text = Mid(resultados(4), 1, 2) + "." + Mid(resultados(4), 3, 2) + "." + Mid(resultados(4), 5, 4)
             infogrilla.Grid1.Cell(lin, 13).text = saldo
             
dale:             resultados.MoveNext
          
         Wend
          lin = lin + 1
             infogrilla.Grid1.Rows = infogrilla.Grid1.Rows + 1
         
         Call totalcomprobante(infogrilla, lin)
          resultados.Close
            Set resultados = Nothing

        End If
 For k = 1 To 6
        infogrilla.Grid1.Column(k).Locked = True
        
        Next k
End Sub
Sub LEERMOVIMIENTOSCONCILIADOS(infogrilla As grillainformes, cuenta, NOMBRE)

Dim resultados As rdoResultset
    Dim csql As New rdoQuery
    Dim rut As String
    Dim PASO As String
    Dim fecha1 As String
    Dim fecha2 As String
    
        fecha1 = Mid(desdefecha.Caption, 7, 4) + "-" + Mid(desdefecha.Caption, 4, 2) + "-" + Mid(desdefecha.Caption, 1, 2)
        fecha2 = Mid(hastafecha.Caption, 7, 4) + "-" + Mid(hastafecha.Caption, 4, 2) + "-" + Mid(hastafecha.Caption, 1, 2)
'
'    SELECT codigocuenta,rutctacte,tipodocumento,numerodocumento,SUM(IF(dh='D', monto,0)) as debe, SUM(IF(dh='H', monto,0)) as haber
'From eltit_conta08.movimientoscontables
'Where rutctacte='0885029000' and codigocuenta='23100026' and fecha >='2009-01-01' and fecha<='2009-07-31'
'GROUP BY tipodocumento,numerodocumento,rutctacte
'order by codigocuenta
        If LEERSALDOSCTACTEmovi(tipoctacte, cuenta, empresaactiva) = 0 And Check3.Value = 1 Then
        Exit Sub
        End If
        
        
    
        Set csql.ActiveConnection = contadb
        csql.sql = "SELECT fecha,tipo,numero,linea,codigocuenta,glosacontable,tipodocumento,numerodocumento,fechadocumento,fechavencimiento,monto,dh,centrocosto,tipoctacte,rutctacte,SUM(IF(dh='D', monto,0)) as debe, SUM(IF(dh='H', monto,0)) as haber "
        csql.sql = csql.sql + "FROM movimientoscontables where codigocuenta='" + tipoctacte + "' and rutctacte='" + cuenta + "' and fecha>='" + fecha1 + "' and fecha<='" + fecha2 + "' "
        csql.sql = csql.sql + "GROUP BY numerodocumento,monto HAVING debe <> haber "
        csql.sql = csql.sql + "order by tipoctacte,rutctacte,fecha,tipo,numero,linea"
        csql.Execute
        Call DATOSSALDOSctacte(cuenta)
        For k = 1 To 6
        infogrilla.Grid1.Column(k).Locked = False
        Next k
        If saldo <> 0 Or csql.RowsAffected <> 0 Then
        
        lin = lin + 1
        infogrilla.Grid1.Rows = infogrilla.Grid1.Rows + 1
        infogrilla.Grid1.Range(lin, 1, lin, 6).Merge
        infogrilla.Grid1.Cell(lin, 1).CellType = cellTextBox
        infogrilla.Grid1.Cell(lin, 10).CellType = cellTextBox
        infogrilla.Grid1.Cell(lin, 1).text = cuenta + " " + NOMBRE
        infogrilla.Grid1.Cell(lin, 6).text = nombrectacte
        infogrilla.Grid1.Cell(lin, 10).text = "SALDO-->"
        infogrilla.Grid1.Cell(lin, 13).text = saldo
        infogrilla.Grid1.Range(lin, 0, lin, infogrilla.Grid1.Cols - 1).FontBold = True
        infogrilla.Grid1.Range(lin, 0, lin, infogrilla.Grid1.Cols - 1).FontUnderline = True
        
        End If
        
        If csql.RowsAffected > 0 Then
        
        
        Set resultados = csql.OpenResultset
        
         While Not resultados.EOF
            If resultados(15) <> resultados(16) Then
             lin = lin + 1
             infogrilla.Grid1.Rows = infogrilla.Grid1.Rows + 1
             For k = 0 To 9
             infogrilla.Grid1.Cell(lin, k + 1).text = resultados(k)
             Next k
             infogrilla.Grid1.Cell(lin, 0).text = resultados("rutctacte")
             saldo = saldo + resultados(15) - resultados(16)
             Rem If resultados(11) = "H" Then infogrilla.Grid1.Cell(lin, 12).text = resultados(10): anteh = anteh + resultados(15) - resultados(16): saldo = saldo - resultados(15) + resultados(16)
             infogrilla.Grid1.Cell(lin, 11).text = resultados(15)
             infogrilla.Grid1.Cell(lin, 12).text = resultados(16)
             
             infogrilla.Grid1.Cell(lin, 5).text = Mid(resultados(4), 1, 2) + "." + Mid(resultados(4), 3, 2) + "." + Mid(resultados(4), 5, 4)
             infogrilla.Grid1.Cell(lin, 13).text = saldo
             If resultados(15) <> 0 And resultados(16) <> 0 Then
             infogrilla.Grid1.Range(lin, 1, lin, infogrilla.Grid1.Cols - 1).ForeColor = &HFF&
             
             End If
             
             
             End If
             
dale:             resultados.MoveNext
          
         Wend
          lin = lin + 1
             infogrilla.Grid1.Rows = infogrilla.Grid1.Rows + 1
         
         Call totalcomprobante(infogrilla, lin)
          resultados.Close
            Set resultados = Nothing

        End If
 For k = 1 To 6
        infogrilla.Grid1.Column(k).Locked = True
        
        Next k
End Sub

Sub totalcomprobante(infogrilla As grillainformes, row)
    infogrilla.Grid1.Range(row, 11, row, 12).Borders(cellEdgeTop) = cellThin
    infogrilla.Grid1.Range(row, 1, row, 12).FontBold = True
    infogrilla.Grid1.Range(row, 1, row, 12).FontUnderline = True
    
    
    infogrilla.Grid1.Cell(row, 10).CellType = cellTextBox
    infogrilla.Grid1.Cell(row, 10).text = "TOTAL "
    infogrilla.Grid1.Cell(row, 11).text = anted
    infogrilla.Grid1.Cell(row, 12).text = anteh
    lin = lin + 2
             infogrilla.Grid1.Rows = infogrilla.Grid1.Rows + 2
        
        anted = 0: anteh = 0: saldo = 0
    End Sub
    





Sub CARGAGRILLA(infogrilla As grillainformes)
Rem DATOS DE LA COLUMNA
    infogrilla.Grid1.DefaultFont.Size = 8
    
    
    FORMATOGRILLA(1, 1) = "FECHA"
    FORMATOGRILLA(1, 2) = "TP"
    FORMATOGRILLA(1, 3) = "NUMERO"
    FORMATOGRILLA(1, 4) = "LIN"
    FORMATOGRILLA(1, 5) = "CUENTA"
    FORMATOGRILLA(1, 6) = "GLOSA"
    FORMATOGRILLA(1, 7) = "TP"
    FORMATOGRILLA(1, 8) = "NUMERO"
    FORMATOGRILLA(1, 9) = "EMISION"
    FORMATOGRILLA(1, 10) = "VENCE"
    FORMATOGRILLA(1, 11) = "DEBE"
    FORMATOGRILLA(1, 12) = "HABER"
    FORMATOGRILLA(1, 13) = "SALDO"
    FORMATOGRILLA(1, 14) = "NOMBRE CUENTA"
    FORMATOGRILLA(1, 15) = "CUENTA CORRIENTE"
    FORMATOGRILLA(1, 16) = "CRCC"
     
    Rem LARGO DE LOS DATOS
    
    FORMATOGRILLA(2, 1) = "8"
    FORMATOGRILLA(2, 2) = "3"
    
    FORMATOGRILLA(2, 3) = "9"
    FORMATOGRILLA(2, 4) = "3"
    FORMATOGRILLA(2, 5) = "8"
    FORMATOGRILLA(2, 6) = "28"
    FORMATOGRILLA(2, 7) = "3"
    FORMATOGRILLA(2, 8) = "8"
    FORMATOGRILLA(2, 9) = "8"
    FORMATOGRILLA(2, 10) = "8"
    FORMATOGRILLA(2, 11) = "11"
    FORMATOGRILLA(2, 12) = "11"
    FORMATOGRILLA(2, 13) = "11"
    FORMATOGRILLA(2, 14) = "30"
    FORMATOGRILLA(2, 15) = "30"

    
    Rem TIPO DE DATOS C=CEROS,S=STRING,N=NUMERICO,D=DATE,
    FORMATOGRILLA(3, 1) = "D"
    FORMATOGRILLA(3, 2) = "S"
    FORMATOGRILLA(3, 3) = "S"
    FORMATOGRILLA(3, 4) = "S"
    FORMATOGRILLA(3, 5) = "S"
    FORMATOGRILLA(3, 6) = "S"
    FORMATOGRILLA(3, 7) = "S"
    FORMATOGRILLA(3, 8) = "S"
    FORMATOGRILLA(3, 9) = "D"
    FORMATOGRILLA(3, 10) = "D"
    FORMATOGRILLA(3, 11) = "N"
    FORMATOGRILLA(3, 12) = "N"
    FORMATOGRILLA(3, 13) = "N"
    FORMATOGRILLA(3, 14) = "S"
    FORMATOGRILLA(3, 15) = "S"
    
    
    Rem FORMATO GRILLA
    FORMATOGRILLA(4, 11) = "###,###,###,##0"
    FORMATOGRILLA(4, 12) = "###,###,###,##0"
    FORMATOGRILLA(4, 13) = "###,###,###,##0"
    Rem LOCCKED
    FORMATOGRILLA(5, 1) = "TRUE"
    FORMATOGRILLA(5, 2) = "TRUE"
    FORMATOGRILLA(5, 3) = "TRUE"
    FORMATOGRILLA(5, 4) = "TRUE"
    FORMATOGRILLA(5, 5) = "TRUE"
    FORMATOGRILLA(5, 6) = "TRUE"
    FORMATOGRILLA(5, 7) = "TRUE"
    FORMATOGRILLA(5, 8) = "TRUE"
    FORMATOGRILLA(5, 9) = "TRUE"
    FORMATOGRILLA(5, 10) = "TRUE"
    FORMATOGRILLA(5, 11) = "TRUE"
    FORMATOGRILLA(5, 12) = "TRUE"
    FORMATOGRILLA(5, 13) = "TRUE"
    FORMATOGRILLA(5, 14) = "TRUE"
    FORMATOGRILLA(5, 15) = "TRUE"
    
    infogrilla.Grid1.Cols = 15
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


Sub leecuentas(infogrilla As grillainformes)
Dim resultados2 As rdoResultset
    Dim csql2 As New rdoQuery
    Dim rut As String
    Dim PASO As String
    Dim LINEAS As Double
    
        
        infogrilla.Grid1.AutoRedraw = False
        
        Set csql2.ActiveConnection = contadb
        csql2.sql = "SELECT codigo,nombre "
        csql2.sql = csql2.sql + "FROM cuentasdelmayor where año='" + Format(fechasistema, "yyyy") + "' "
        If cmindi.Value = True Then
        csql2.sql = csql2.sql + "and codigo='" + cmdato1.text + cmdato2.text + cmdato3.text + "' "
        End If
        csql2.sql = csql2.sql + "order by codigo"
        csql2.Execute
        
        
        LINEAS = 0
        If csql2.RowsAffected > 0 Then
        barra.Max = csql2.RowsAffected + 1
        barra.Value = 0
        
        Set resultados2 = csql2.OpenResultset
        While Not resultados2.EOF
        LINEAS = LINEAS + 1
        barra.Value = barra.Value + 1
        
        
        If Mid(resultados2(0), 5, 4) <> "0000" Then Call LEERMOVIMIENTOS(infogrilla, resultados2(0), resultados2(1))
       
        resultados2.MoveNext
        Wend
          
          resultados2.Close
            Set resultados2 = Nothing

        End If
        infogrilla.Grid1.Column(8).Locked = True
        infogrilla.Grid1.Column(9).Locked = True
        infogrilla.Grid1.Column(10).Locked = True
        
infogrilla.Grid1.AutoRedraw = True
infogrilla.Grid1.Refresh

End Sub

Sub LEERSALDOS(cuenta)
Dim resultados3 As rdoResultset
    
    Dim mesin As String
    Dim añoin As String
    Dim cSql3 As New rdoQuery
    Dim rut As String
    Dim PASO As String
    Dim fecha1 As String
    Dim fecha2 As String
    Dim mesante As Integer
    
    campos(0, 0) = "codigo"
    campos(1, 0) = "año"
    campos(2, 0) = "debeanterior"
    campos(3, 0) = "haberanterior"
    campos(4, 0) = "debe01"
    campos(5, 0) = "debe02"
    campos(6, 0) = "debe03"
    campos(7, 0) = "debe04"
    campos(8, 0) = "debe05"
    campos(9, 0) = "debe06"
    campos(10, 0) = "debe07"
    campos(11, 0) = "debe08"
    campos(12, 0) = "debe09"
    campos(13, 0) = "debe10"
    campos(14, 0) = "debe11"
    campos(15, 0) = "debe12"
    campos(16, 0) = "haber01"
    campos(17, 0) = "haber02"
    campos(18, 0) = "haber03"
    campos(19, 0) = "haber04"
    campos(20, 0) = "haber05"
    campos(21, 0) = "haber06"
    campos(22, 0) = "haber07"
    campos(23, 0) = "haber08"
    campos(24, 0) = "haber09"
    campos(25, 0) = "HABER10"
    campos(26, 0) = "HABER11"
    campos(27, 0) = "HABER12"
    campos(28, 0) = ""
    
    If dedonde = 1 Then condicion = "codigo=" + "'" + cuenta + "' and año='" + Mid(desdefecha.Caption, 7, 4) + "' order by codigo"
    If dedonde = 3 Then condicion = "codigo=" + "'" + cuenta + "' and año='" + Mid(desdefecha.Caption, 7, 4) + "' order by codigo"
    
    If dedonde = 1 Then campos(0, 2) = "saldosdelmayor"
    If dedonde = 3 Then campos(0, 2) = "saldoscentrosdecosto"
    
    op = 5
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)
   ' If sqlconta.status = 4 Then Stop
    sumador = Val(sqlconta.response(2, 3)) - Val(sqlconta.response(3, 3))
  
    saldo = sumador
Rem acumula fecha
        fecha1 = Mid(desdefecha.Caption, 7, 4) + "-" + Mid(desdefecha.Caption, 4, 2) + "-" + Mid(desdefecha.Caption, 1, 2)

    
        
        Set cSql3.ActiveConnection = contadb
        cSql3.sql = "SELECT SUM(monto),dh "
        If dedonde = 1 Then cSql3.sql = cSql3.sql + "FROM movimientoscontables where codigocuenta='" + cuenta + "' and fecha<'" + fecha1 + "' and fecha>='" + Format(fechasistema, "yyyy") + "-01-01" + "' "
     
        If dedonde = 3 Then cSql3.sql = cSql3.sql + "FROM movimientoscontables where centrocosto='" + cuenta + "' and fecha<'" + fecha1 + "' and fecha>='" + Format(fechasistema, "yyyy") + "-01-01" + "' "
        
        cSql3.sql = cSql3.sql + "GROUP by DH"
        cSql3.Execute
        
        If cSql3.RowsAffected > 0 Then
        
        
        Set resultados3 = cSql3.OpenResultset
        
         While Not resultados3.EOF
         If resultados3(1) = "D" Then saldo = saldo + resultados3(0)
         If resultados3(1) = "H" Then saldo = saldo - resultados3(0)
         
             
             resultados3.MoveNext
           
         Wend
          resultados3.Close
            Set resultados3 = Nothing

        End If

End Sub
Sub DATOSSALDOS(cuenta)
Call LEERSALDOS(cuenta)






End Sub
Sub leecrcc(infogrilla As grillainformes)
Dim resultados2 As rdoResultset
    Dim csql2 As New rdoQuery
    Dim rut As String
    Dim PASO As String
    Dim LINEAS As Double
        Set csql2.ActiveConnection = contadb
        csql2.sql = "SELECT codigo,nombre "
        csql2.sql = csql2.sql + "FROM centrosdecosto "
        If ccindi.Value = True Then csql2.sql = csql2.sql + "where codigo='" + ccdato1.text + ccdato2.text + "' and año='" & Format(fechasistema, "yyyy") & "'"
        
        csql2.sql = csql2.sql + " group by codigo order by codigo"
         csql2.Execute
        LINEAS = 0
  
        If csql2.RowsAffected > 0 Then
     
        Set resultados2 = csql2.OpenResultset
        While Not resultados2.EOF
    
        If Mid(resultados2(0), 3, 2) <> "00" Then Call LEERMOVIMIENTOS(infogrilla, resultados2(0), resultados2(1))
        
        resultados2.MoveNext
        Wend
          
          resultados2.Close
            Set resultados2 = Nothing

        End If
        infogrilla.Grid1.Column(8).Locked = True
        infogrilla.Grid1.Column(9).Locked = True
        infogrilla.Grid1.Column(10).Locked = True
        

End Sub

Sub leercuentascorrientes(infogrilla As grillainformes)

Dim resultados2 As rdoResultset
    Dim csql2 As New rdoQuery
    Dim rut As String
    Dim PASO As String
    Dim LINEAS As Double
    infogrilla.Grid1.AutoRedraw = False
    
        Set csql2.ActiveConnection = contadb
        csql2.sql = "SELECT tipo,rut,nombre "
        If cttoda.Value = True Then csql2.sql = csql2.sql + "FROM cuentascorrientes where tipo>'00' and año='" + Format(fechasistema, "yyyy") + "' "
        If cttodatipo.Value = True Then csql2.sql = csql2.sql + "FROM cuentascorrientes where tipo='" + ctdato1.text + "' and año='" + Format(fechasistema, "yyyy") + "' "
        If ctindi.Value = True Then csql2.sql = csql2.sql + "FROM cuentascorrientes where tipo='" + ctdato1.text + "' and rut='" + ctdato2.text + DV.text + "' and año='" + Format(fechasistema, "yyyy") + "' "
        
        csql2.sql = csql2.sql + "order by tipo,nombre"
       
        csql2.Execute
        
        LINEAS = 0
        If csql2.RowsAffected > 0 Then
         barra.Max = csql2.RowsAffected
        barra.Value = 0
        Set resultados2 = csql2.OpenResultset
        While Not resultados2.EOF
        LINEAS = LINEAS + 1
        barra.Value = barra.Value + 1
        barra.Refresh
        
        tipoctacte = resultados2(0)
        If Check1.Value = "1" Then
        Call LEERMOVIMIENTOSCONCILIADOS(infogrilla, resultados2(1), resultados2(2))
        
        Else
        Call LEERMOVIMIENTOS(infogrilla, resultados2(1), resultados2(2))
        End If
        resultados2.MoveNext
        Wend
          
          resultados2.Close
            Set resultados2 = Nothing

        End If
        infogrilla.Grid1.Column(8).Locked = True
        infogrilla.Grid1.Column(9).Locked = True
        infogrilla.Grid1.Column(10).Locked = True
        
infogrilla.Grid1.AutoRedraw = True
infogrilla.Grid1.Refresh


End Sub

Sub LEERSALDOSCTACTE(cuenta)
   Dim resultados3 As rdoResultset
    Dim mesin As String
    Dim añoin As String
    Dim cSql3 As New rdoQuery
    Dim rut As String
    Dim PASO As String
    Dim fecha1 As String
    Dim fecha2 As String
    Dim mesante As Integer
    
    campos(0, 0) = "tipo"
    campos(1, 0) = "rut"
    campos(2, 0) = "año"
    campos(3, 0) = "debeanterior"
    campos(4, 0) = "haberanterior"
    campos(5, 0) = "debe01"
    campos(6, 0) = "debe02"
    campos(7, 0) = "debe03"
    campos(8, 0) = "debe04"
    campos(9, 0) = "debe05"
    campos(10, 0) = "debe06"
    campos(11, 0) = "debe07"
    campos(12, 0) = "debe08"
    campos(13, 0) = "debe09"
    campos(14, 0) = "debe10"
    campos(15, 0) = "debe11"
    campos(16, 0) = "debe12"
    campos(17, 0) = "haber01"
    campos(18, 0) = "haber02"
    campos(19, 0) = "haber03"
    campos(20, 0) = "haber04"
    campos(21, 0) = "haber05"
    campos(22, 0) = "haber06"
    campos(23, 0) = "haber07"
    campos(24, 0) = "haber08"
    campos(25, 0) = "haber09"
    campos(26, 0) = "haber10"
    campos(27, 0) = "haber11"
    campos(28, 0) = "haber12"
    campos(29, 0) = ""
    condicion = "tipo=" + "'" + tipoctacte + "' and rut='" + cuenta + "' and año='" + Mid(desdefecha.Caption, 7, 4) + "'"
    campos(0, 2) = "saldosctacte"
    op = 5
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)
   ' If sqlconta.status = 4 Then Stop
    sumador = Val(sqlconta.response(3, 3)) - Val(sqlconta.response(4, 3))
    saldo = sumador
    Rem acumula fecha
        fecha1 = Mid(desdefecha.Caption, 7, 4) + "-" + Mid(desdefecha.Caption, 4, 2) + "-" + Mid(desdefecha.Caption, 1, 2)
        Set cSql3.ActiveConnection = contadb
        cSql3.sql = "SELECT SUM(monto),dh "
        cSql3.sql = cSql3.sql + "FROM movimientoscontables where codigocuenta='" + tipoctacte + "' and rutctacte='" + cuenta + "' and fecha<'" + fecha1 + "' and fecha>='" + Format(fechasistema, "yyyy") + "-01-01" + "' "
        cSql3.sql = cSql3.sql + "GROUP BY DH"
        cSql3.Execute
        
        If cSql3.RowsAffected > 0 Then
        
        
        Set resultados3 = cSql3.OpenResultset
        
         While Not resultados3.EOF
         If resultados3(1) = "D" Then saldo = saldo + resultados3(0)
         If resultados3(1) = "H" Then saldo = saldo - resultados3(0)
         resultados3.MoveNext
           
         Wend
          resultados3.Close
            Set resultados3 = Nothing

        End If
End Sub

Sub DATOSSALDOSctacte(cuenta)

Call LEERSALDOSCTACTE(cuenta)

End Sub


Sub ayudamayor(ByRef caja As TextBox)
  
   
    
    Dim campos As Variant
    Dim cfijo As Variant
    Dim largo As Variant
    campos = Array("codigo", "nombre")
    largo = Array("12s", "40s")
    cfijo = "no"
    cabezas = Array("codigo", "nombre")
    mensajeAyuda = "Ayuda Cuentas del Mayor"
    
    Call cargaAyudaT(Servidor, basebus, Usuario, password, "cuentasdelmayor", pivote, campos, cfijo, largo, 2)
    

    
    
    If Val(pivote.text) = 0 Then cmdato1.SetFocus: GoTo no
    cmdato1.text = Mid(pivote.text, 1, 2)
    cmdato2.text = Mid(pivote.text, 3, 2)
    cmdato3.text = Mid(pivote.text, 5, 4)
    
    
    caja.Enabled = True
    caja.SetFocus
    caja.MaxLength = 2
no:
End Sub

Sub ayudamayor2(ByRef caja As TextBox)
  
   
    
    Dim campos As Variant
    Dim cfijo As Variant
    Dim largo As Variant
    campos = Array("codigo", "nombre")
    largo = Array("12s", "40s")
    cfijo = "mid(codigo,1,1)>2"
    cabezas = Array("codigo", "nombre")
    mensajeAyuda = "Ayuda Cuentas del Mayor"
    pivote.MaxLength = 8
    Call cargaAyudaT(Servidor, basebus, Usuario, password, "cuentasdelmayor", pivote, campos, cfijo, largo, 2)
    
    
    If Val(pivote.text) = 0 Then cmdato1.SetFocus: GoTo no
    txtdato1.text = Mid(pivote.text, 1, 2)
    txtdato2.text = Mid(pivote.text, 3, 2)
    txtdato3.text = Mid(pivote.text, 5, 4)
      pivote.MaxLength = 4
    
    caja.Enabled = True
    caja.SetFocus
    caja.MaxLength = 2
no:
End Sub
   
Private Sub ProgressBar1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

End Sub


Private Sub Label14_Click()

End Sub

Private Sub Image1_Click()

End Sub

Private Sub Text9_Change()

End Sub

Sub ayudatipocuenta(ByRef caja As TextBox)
    Dim campos As Variant
    Dim cfijo As Variant
    Dim largo As Variant
    campos = Array("codigo", "nombre")
    largo = Array("8s", "40s")
    cfijo = "CTACTE <> '0' and año='" + Format(fechasistema, "yyyy") + "'"
    cabezas = Array("cuenta", "nombre")
    mensajeAyuda = "Ayuda tipo de Cuentas Corrientes"
        
    Call cargaAyudaT(Servidor, basebus, Usuario, password, "cuentasdelmayor", caja, campos, cfijo, largo, 2)
    caja.Enabled = True
    caja.SetFocus
End Sub

Sub ayudactacte(ByRef caja As TextBox)
    Dim campos As Variant
    Dim cfijo As Variant
    Dim largo As Variant
    campos = Array("rut", "nombre")
    largo = Array("12n", "40s")
    cfijo = "tipo='" & ctdato1.text & "' and año='" + Format(fechasistema, "yyyy") + "'"
    cabezas = Array("rut", "nombre")
    mensajeAyuda = "Ayuda Cuentas Corrientes"
    
    Call cargaAyudaT(Servidor, basebus, Usuario, password, "cuentascorrientes", ctdato2, campos, cfijo, largo, 2)

    If Val(caja.text) = 0 Then ctdato2.SetFocus: GoTo no
   
    
    DV.text = rut(caja)
    caja.Enabled = True
    caja.SetFocus

no:

End Sub
Sub ayudacentrodecosto(ByRef caja As TextBox)
    Dim campos As Variant
    Dim cfijo As Variant
    Dim largo As Variant
    campos = Array("codigo", "nombre")
    largo = Array("8s", "40s")
    cfijo = "año='" + Format(fechasistema, "yyyy") + "'"
    cabezas = Array("cuenta", "nombre")
    mensajeAyuda = "Ayuda tipo de centros de costo "
        
    Call cargaAyudaT(Servidor, basebus, Usuario, password, "centrosdecosto", caja, campos, cfijo, largo, 2)
    caja.Enabled = True
    caja.SetFocus
End Sub


Private Sub botonmisaccesos_Click()
programafiltro = Me.Caption
misaccesos.Show
End Sub


Private Sub botonmisfavoritos_Click()
Call AgregaFavoritos(USUARIOSISTEMA, App.EXEName, Me.Name, Me.Caption)
End Sub

Private Sub txt_crcc_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Call ceros(txt_crcc)
If leercrcc(txt_crcc) <> "" Then
lbl_crcc.Caption = leercrcc(txt_crcc)


Else
MsgBox ("codigo de crcc no existe")


cmdato3.SetFocus
End If
End If

End Sub
Private Sub txt_crcc_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = vbKeyF2 Then Call ayudacentrodecosto(txt_crcc)

End Sub
Private Function leercrcc(cuenta) As String
      campos(0, 0) = "nombre"
      campos(1, 0) = ""
      campos(0, 2) = "centrosdecosto"
      condicion = "codigo=" + "'" + cuenta + "' and año='" + Format(fechasistema, "yyyy") + "'"
      op = 5
      sqlconta.response = campos
      Set sqlconta.conexion = contadb
      Call sqlconta.sqlconta(op, condicion)
        If sqlconta.status = 0 Then
            leercrcc = sqlconta.response(0, 3)
        Else
            leercrcc = ""
        End If
End Function

Private Sub txtdato1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF2 Then Call ayudamayor2(txtdato1)

End Sub

Private Sub txtdato1_KeyPress(KeyAscii As Integer)
    KeyAscii = esNumero(KeyAscii)
    If KeyAscii = 13 Then
        Call ceros(txtdato1)
        txtdato2.SetFocus
    End If
    
End Sub

Private Sub txtdato2_KeyPress(KeyAscii As Integer)
    KeyAscii = esNumero(KeyAscii)
    If KeyAscii = 13 Then
        Call ceros(txtdato2)
        txtdato3.SetFocus
    End If
End Sub

Private Sub txtdato3_KeyPress(KeyAscii As Integer)
    KeyAscii = esNumero(KeyAscii)
    If KeyAscii = 13 Then
        Call ceros(txtdato3)
        LblnombreCuenta.Caption = leerNombreCuentaMayor(txtdato1.text & txtdato2.text & txtdato3.text, 2)
        Command6.SetFocus
    End If
End Sub
