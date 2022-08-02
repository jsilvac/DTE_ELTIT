VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash.ocx"
Object = "{ADBBAED6-B16F-40DC-80DF-B44910CBA76C}#1.0#0"; "Frame-Xp.ocx"
Begin VB.Form confi02 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mestro de Empresas"
   ClientHeight    =   10035
   ClientLeft      =   630
   ClientTop       =   825
   ClientWidth     =   7755
   DrawStyle       =   5  'Transparent
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   669
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   517
   Begin XPFrame.FrameXp FrameXp2 
      Height          =   4815
      Left            =   240
      TabIndex        =   14
      Top             =   3240
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   8493
      BackColor       =   16744576
      Caption         =   "Datos Contables"
      CaptionEstilo3D =   1
      BackColor       =   16744576
      ForeColor       =   8438015
      BordeColor      =   -2147483635
      ColorBarraArriba=   4194304
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
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
         ForeColor       =   &H00FF8080&
         Height          =   285
         Left            =   2640
         MaxLength       =   4
         TabIndex        =   55
         Tag             =   "yyyy"
         Top             =   4320
         Width           =   735
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
         ForeColor       =   &H00FF8080&
         Height          =   285
         Left            =   2160
         MaxLength       =   2
         TabIndex        =   54
         Tag             =   "mm"
         Top             =   4320
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
         ForeColor       =   &H00FF8080&
         Height          =   285
         Left            =   1680
         MaxLength       =   2
         TabIndex        =   52
         Tag             =   "dd"
         Top             =   4320
         Width           =   375
      End
      Begin VB.TextBox DATO22 
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
         ForeColor       =   &H00FF8080&
         Height          =   285
         Left            =   1680
         MaxLength       =   4
         TabIndex        =   45
         Tag             =   "ppm"
         Top             =   3960
         Width           =   855
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
         ForeColor       =   &H00FF8080&
         Height          =   285
         Left            =   2040
         MaxLength       =   8
         TabIndex        =   29
         Tag             =   "cuentaproveedor"
         Top             =   360
         Width           =   1455
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
         ForeColor       =   &H00FF8080&
         Height          =   285
         Left            =   5400
         MaxLength       =   8
         TabIndex        =   28
         Tag             =   "cuentaclientes"
         Top             =   360
         Width           =   1455
      End
      Begin VB.TextBox dato9 
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
         ForeColor       =   &H00FF8080&
         Height          =   285
         Left            =   2040
         MaxLength       =   8
         TabIndex        =   27
         Tag             =   "ivacredito"
         Top             =   720
         Width           =   1455
      End
      Begin VB.TextBox dato10 
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
         ForeColor       =   &H00FF8080&
         Height          =   285
         Left            =   5400
         MaxLength       =   8
         TabIndex        =   26
         Tag             =   "ivadebito"
         Top             =   720
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
         ForeColor       =   &H00FF8080&
         Height          =   285
         Left            =   2040
         MaxLength       =   8
         TabIndex        =   25
         Tag             =   "cuentahonorarios"
         Top             =   1080
         Width           =   1455
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
         ForeColor       =   &H00FF8080&
         Height          =   285
         Left            =   5400
         MaxLength       =   10
         TabIndex        =   24
         Tag             =   "retencionhonorarios"
         Top             =   1080
         Width           =   1455
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
         ForeColor       =   &H00FF8080&
         Height          =   285
         Left            =   2040
         MaxLength       =   8
         TabIndex        =   23
         Tag             =   "cuentaperdida"
         Top             =   1440
         Width           =   1455
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
         ForeColor       =   &H00FF8080&
         Height          =   285
         Left            =   5400
         MaxLength       =   8
         TabIndex        =   22
         Tag             =   "cuentaganancia"
         Top             =   1440
         Width           =   1455
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
         ForeColor       =   &H00FF8080&
         Height          =   285
         Left            =   1680
         MaxLength       =   1
         TabIndex        =   21
         Tag             =   "seguridad"
         Top             =   2160
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
         ForeColor       =   &H00FF8080&
         Height          =   285
         Left            =   5400
         MaxLength       =   1
         TabIndex        =   20
         Tag             =   "auditoria"
         Top             =   2160
         Width           =   1455
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
         ForeColor       =   &H00FF8080&
         Height          =   285
         Left            =   1680
         MaxLength       =   40
         TabIndex        =   19
         Tag             =   "codigoae"
         Top             =   2520
         Width           =   1455
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
         ForeColor       =   &H00FF8080&
         Height          =   285
         Left            =   5400
         MaxLength       =   10
         TabIndex        =   18
         Tag             =   "clavesii"
         Top             =   2520
         Width           =   1455
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
         ForeColor       =   &H00FF8080&
         Height          =   285
         Left            =   1680
         MaxLength       =   10
         TabIndex        =   17
         Tag             =   "rutrepresentante"
         Top             =   2880
         Width           =   1455
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
         ForeColor       =   &H00FF8080&
         Height          =   285
         Left            =   1680
         MaxLength       =   30
         TabIndex        =   16
         Tag             =   "representantelegal"
         Top             =   3240
         Width           =   3855
      End
      Begin VB.TextBox dato21 
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
         ForeColor       =   &H00FF8080&
         Height          =   285
         Left            =   1680
         MaxLength       =   10
         TabIndex        =   15
         Tag             =   "emailcontable"
         Top             =   3600
         Width           =   3855
      End
      Begin VB.Label Label25 
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Fecha Flujo"
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
         Left            =   240
         TabIndex        =   53
         Top             =   4320
         Width           =   1335
      End
      Begin VB.Label Label14 
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "P.P.M"
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
         Left            =   240
         TabIndex        =   46
         Top             =   3960
         Width           =   1335
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Cuenta Proveedor"
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
         Left            =   240
         TabIndex        =   44
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label Label7 
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Cuenta Clientes"
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
         Left            =   3720
         TabIndex        =   43
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label Label8 
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Iva Credito"
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
         Left            =   240
         TabIndex        =   42
         Top             =   720
         Width           =   1695
      End
      Begin VB.Label Label9 
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Iva Debito"
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
         Left            =   3720
         TabIndex        =   41
         Top             =   720
         Width           =   1575
      End
      Begin VB.Label Label11 
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Cuenta Honorario"
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
         Left            =   240
         TabIndex        =   40
         Top             =   1080
         Width           =   1695
      End
      Begin VB.Label Label12 
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Retencion Honorarios"
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
         Left            =   3720
         TabIndex        =   39
         Top             =   1080
         Width           =   1575
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Cuenta Perdida"
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
         Left            =   240
         TabIndex        =   38
         Top             =   1440
         Width           =   1695
      End
      Begin VB.Label Label15 
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Seguridad"
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
         Left            =   240
         TabIndex        =   37
         Top             =   2160
         Width           =   1335
      End
      Begin VB.Label Label16 
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Auditoria"
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
         Left            =   3720
         TabIndex        =   36
         Top             =   2160
         Width           =   1575
      End
      Begin VB.Label Label17 
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "C.Acti.Econ."
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
         Left            =   240
         TabIndex        =   35
         Top             =   2520
         Width           =   1335
      End
      Begin VB.Label Label18 
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Clave Sii"
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
         Left            =   3720
         TabIndex        =   34
         Top             =   2520
         Width           =   1575
      End
      Begin VB.Label Label19 
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Cuenta Ganancia"
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
         Left            =   3720
         TabIndex        =   33
         Top             =   1440
         Width           =   1575
      End
      Begin VB.Label Label20 
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Rut Represen.Legal"
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
         Left            =   240
         TabIndex        =   32
         Top             =   2880
         Width           =   1335
      End
      Begin VB.Label Label21 
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Rep.Legal"
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
         Left            =   240
         TabIndex        =   31
         Top             =   3240
         Width           =   1335
      End
      Begin VB.Label Label22 
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Email"
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
         Left            =   240
         TabIndex        =   30
         Tag             =   "emailcontable"
         Top             =   3600
         Width           =   1335
      End
   End
   Begin XPFrame.FrameXp FrameXp1 
      Height          =   3015
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   5318
      BackColor       =   16761024
      Caption         =   "Maestro de Empresa"
      CaptionEstilo3D =   1
      BackColor       =   16761024
      ForeColor       =   8438015
      BordeColor      =   -2147483635
      ColorBarraArriba=   4194304
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.TextBox dato24 
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
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   1320
         MaxLength       =   3
         TabIndex        =   49
         Tag             =   "rut"
         Top             =   2520
         Width           =   735
      End
      Begin VB.TextBox dato23 
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
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   4800
         MaxLength       =   11
         TabIndex        =   47
         Tag             =   "rut"
         Top             =   2160
         Width           =   1695
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
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   1320
         MaxLength       =   2
         TabIndex        =   7
         Tag             =   "codigoempresa"
         Top             =   360
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
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   1320
         MaxLength       =   50
         TabIndex        =   6
         Tag             =   "nombre"
         Top             =   720
         Width           =   5175
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
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   1320
         MaxLength       =   50
         TabIndex        =   5
         Tag             =   "comuna"
         Top             =   1440
         Width           =   5175
      End
      Begin VB.TextBox dato3 
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
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   1320
         MaxLength       =   50
         TabIndex        =   4
         Tag             =   "direccion"
         Top             =   1080
         Width           =   5175
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
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   1320
         MaxLength       =   50
         TabIndex        =   3
         Tag             =   "ciudad"
         Top             =   1800
         Width           =   5175
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
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   1320
         MaxLength       =   10
         TabIndex        =   2
         Tag             =   "rut"
         Top             =   2160
         Width           =   1695
      End
      Begin VB.Label lblnombrebanco 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2160
         TabIndex        =   51
         Top             =   2520
         Width           =   4335
      End
      Begin VB.Label Label24 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Banco "
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
         Left            =   240
         TabIndex        =   50
         Top             =   2520
         Width           =   975
      End
      Begin VB.Label Label23 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Cuenta Bancaria"
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
         Left            =   3120
         TabIndex        =   48
         Top             =   2160
         Width           =   1575
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Codigo"
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
         Left            =   240
         TabIndex        =   13
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label5 
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
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Direccion"
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
         Left            =   240
         TabIndex        =   11
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Comuna"
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
         Left            =   240
         TabIndex        =   10
         Top             =   1440
         Width           =   975
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Ciudad"
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
         Left            =   240
         TabIndex        =   9
         Top             =   1800
         Width           =   975
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
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
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   2160
         Width           =   975
      End
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash opciones 
      Height          =   1695
      Left            =   360
      TabIndex        =   0
      Top             =   8160
      Width           =   7335
      _cx             =   12938
      _cy             =   2990
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
Attribute VB_Name = "confi02"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private MODIFI As Integer
Private con As rdoConnection


Private Sub Command1_Click()
crearTablascontables

End Sub

Private Sub COMMAND2_Click()
eliminatablas

End Sub

Private Sub dato1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF2 Then Call ayudaempresa(dato2)
    Call flechas(dato1, dato2, KeyCode)
End Sub

Private Sub dato1_LostFocus()
    If sl = 0 Then leer
sl = 0
End Sub

Private Sub dato10_GotFocus()
MsgBox ("SI NO DESEA CONTABILIZAR FACTURAS DE VENTA AUTOMATICAMENTE COLOCAR 99999999")
End Sub

Private Sub dato2_KeyDown(KeyCode As Integer, Shift As Integer)
    Call flechas(dato1, dato3, KeyCode)
End Sub

Private Sub dato22_KeyDown(KeyCode As Integer, Shift As Integer)
    Call flechas(DATO21, DATO25, KeyCode)
End Sub

Private Sub dato22_KeyPress(KeyAscii As Integer)
 KeyAscii = esNumero(KeyAscii)
 If KeyAscii = 13 Then
    Rem grabar:     retorno
     Call Pregunta(DATO22, DATO25)
End If
End Sub

Private Sub dato23_KeyPress(KeyAscii As Integer)
    KeyAscii = esNumero(KeyAscii)
    If KeyAscii = 13 Then
        Call ceros(dato23)
        Call Pregunta(dato23, dato24)
    End If
End Sub

 

Private Sub dato24_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF2 Then
        Call ayudabancos(dato24)
    End If
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
Private Sub dato24_KeyPress(KeyAscii As Integer)
    KeyAscii = esNumero(KeyAscii)
    If KeyAscii = 13 Then
        Call ceros(dato24)
        lblnombrebanco.Caption = leerbanco(dato24.text)
        If lblnombrebanco.Caption <> "" Then
            Call Pregunta(dato24, dato7)
        Else
            MsgBox "BANCO NO EXISTE POR FAVOR REVISAR", vbCritical, "ATENCION"
            dato24.SetFocus
        End If
    End If
End Sub

Private Sub dato25_GotFocus()
Call cargatexto(DATO25)
End Sub

Private Sub dato25_KeyDown(KeyCode As Integer, Shift As Integer)
    Call flechas(DATO22, dato26, KeyCode)
End Sub

Private Sub dato25_KeyPress(KeyAscii As Integer)
 KeyAscii = esNumero(KeyAscii)
 If KeyAscii = 13 Then
 If DATO25 = "" Then DATO25 = Format(fechasistema, "dd")
  Call ceros(DATO25)
  Call Pregunta(DATO25, dato26)
 End If
 'grabar: retorno
End Sub

Private Sub dato26_GotFocus()
Call cargatexto(dato26)
End Sub

Private Sub dato26_KeyDown(KeyCode As Integer, Shift As Integer)
    Call flechas(DATO25, dato27, KeyCode)
End Sub

Private Sub dato26_KeyPress(KeyAscii As Integer)
KeyAscii = esNumero(KeyAscii)
 If KeyAscii = 13 Then
 If dato26 = "" Then dato26 = Format(fechasistema, "mm")
  Call ceros(dato26)
  Call Pregunta(dato26, dato27)
 End If
End Sub

Private Sub dato27_GotFocus()
Call cargatexto(dato27)
End Sub

Private Sub dato27_KeyDown(KeyCode As Integer, Shift As Integer)
    Call flechas(dato26, dato27, KeyCode)
End Sub

Private Sub dato27_KeyPress(KeyAscii As Integer)
KeyAscii = esNumero(KeyAscii)
 If KeyAscii = 13 Then
 If dato27 = "" Then dato27 = Format(fechasistema, "yyyy")
  Call ceros(dato27)
  If IsDate(DATO25 & "-" & dato26 & "-" & dato27) = True Then
    Call grabar
    Call retorno
    Else
    MsgBox "DEBE INGRESAR UNA FACHA VALIDA"
    DATO25.SetFocus
  End If
 End If
End Sub

Private Sub dato3_KeyDown(KeyCode As Integer, Shift As Integer)
    Call flechas(dato2, dato4, KeyCode)
End Sub

Private Sub dato4_KeyDown(KeyCode As Integer, Shift As Integer)
    Call flechas(dato3, DATO5, KeyCode)
End Sub

Private Sub dato9_GotFocus()
MsgBox ("SI NO DESEA CONTABILIZAR FACTURAS DE COMPRA AUTOMATICAMENTE COLOCAR 99999999")
End Sub


Private Sub Form_Load()
    Call Conectarconta(Servidor, clientesistema + "conta", Usuario, password)
    sc = 0
    opciones.Visible = False

End Sub

Private Sub dato1_KeyPress(KeyAscii As Integer)
    snum = 1: KeyAscii = esNumero(KeyAscii)
    If KeyAscii = 13 Then
    Call ceros(dato1):
    Call Pregunta(dato1, dato2)
    End If
End Sub

Private Sub dato2_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then sc = 1: Call Pregunta(dato2, dato3)
End Sub

Private Sub dato3_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then Call Pregunta(dato3, dato4)
End Sub

Private Sub dato4_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then Call Pregunta(dato4, DATO5)
End Sub
Private Sub dato5_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then Call Pregunta(DATO5, dato6)
End Sub
Private Sub dato6_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then Call Pregunta(dato6, dato23)
End Sub
Private Sub dato7_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then Call Pregunta(dato7, dato8)
End Sub

Private Sub dato8_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then Call Pregunta(dato8, dato9)
End Sub
Private Sub dato9_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then Call Pregunta(dato9, dato10)
End Sub
Private Sub dato10_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then Call Pregunta(dato10, dato11)
End Sub
Private Sub dato11_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then Call Pregunta(dato11, dato12)
End Sub
Private Sub dato12_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then Call Pregunta(dato12, dato13)
End Sub
Private Sub dato13_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then Call Pregunta(dato13, dato14)
End Sub
Private Sub dato14_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then Call Pregunta(dato14, dato15)
End Sub
Private Sub dato15_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then Call Pregunta(dato15, dato16)
End Sub
Private Sub dato16_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then Call Pregunta(dato16, dato17)
End Sub
Private Sub dato17_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then Call Pregunta(dato17, dato18)
End Sub
Private Sub dato18_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then Call Pregunta(dato18, dato19)
End Sub
Private Sub dato19_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then Call Pregunta(dato19, DATO20)
End Sub
Private Sub DATO20_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then Call Pregunta(DATO20, DATO21)
End Sub
Private Sub dato21_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    
    If KeyAscii = 13 Then Call Pregunta(DATO20, DATO22)
End Sub



Sub leer()
    campos(0, 0) = dato1.Tag 'CODIGO
    campos(1, 0) = dato2.Tag 'BODEGA
    campos(2, 0) = dato3.Tag 'DIRECCION
    campos(3, 0) = dato4.Tag 'CIUDAD
    campos(4, 0) = DATO5.Tag 'OTROS
    campos(5, 0) = dato6.Tag 'rut
    campos(6, 0) = dato7.Tag 'cuentaprov
    campos(7, 0) = dato8.Tag 'cuenta cliente
    campos(8, 0) = dato9.Tag 'iva credito
    campos(9, 0) = dato10.Tag 'iva debito
    campos(10, 0) = dato11.Tag 'cuenta honorarios
    campos(11, 0) = dato12.Tag 'retencionhonorarios
    campos(12, 0) = dato13.Tag 'cuenta perdida
    campos(13, 0) = dato14.Tag 'cuenta ganancia
    campos(14, 0) = dato15.Tag 'seguridad
    campos(15, 0) = dato16.Tag 'codigo gae
    campos(16, 0) = dato17.Tag 'auditoria
    campos(17, 0) = dato18.Tag 'clavesii
    campos(18, 0) = dato19.Tag 'representantelegal
    campos(19, 0) = DATO20.Tag 'rut representante
    campos(20, 0) = DATO21.Tag 'mailcontable
    campos(21, 0) = DATO22.Tag
    campos(22, 0) = "cuentabancaria"
    campos(23, 0) = "bancocuenta"
    campos(24, 0) = "flujoactualizado"
    campos(25, 0) = ""
    
    campos(0, 2) = "maestroempresas"
    condicion = "codigoempresa=" + "'" + dato1.text + "'"
    op = 5
    
    
    sqlconta.response = campos
    
    Set sqlconta.conexion = conta
    Call sqlconta.sqlconta(op, condicion)
    If sqlconta.status = 0 Then carga: opciones.Visible = True: disponible (True): habilita (True): opciones.SetFocus
    
End Sub
Sub leersiguiente()
    campos(0, 0) = dato1.Tag 'CODIGO
    campos(1, 0) = dato2.Tag 'BODEGA
    campos(2, 0) = dato3.Tag 'DIRECCION
    campos(3, 0) = dato4.Tag 'CIUDAD
    campos(4, 0) = DATO5.Tag 'OTROS
    campos(5, 0) = dato6.Tag
    campos(6, 0) = dato7.Tag
    campos(7, 0) = dato8.Tag
    campos(8, 0) = dato9.Tag
    campos(9, 0) = dato10.Tag
    campos(10, 0) = dato11.Tag
    campos(11, 0) = dato12.Tag
    campos(12, 0) = dato13.Tag
    campos(13, 0) = dato14.Tag
    campos(14, 0) = dato15.Tag
    campos(15, 0) = dato16.Tag
    campos(16, 0) = dato17.Tag
    campos(17, 0) = dato18.Tag
    campos(18, 0) = dato19.Tag
    campos(19, 0) = DATO20.Tag
    campos(20, 0) = DATO21.Tag
    campos(21, 0) = DATO22.Tag
    campos(22, 0) = "cuentabancaria"
    campos(23, 0) = "bancocuenta"
    campos(24, 0) = ""
    
    campos(0, 2) = "maestroempresas"
    condicion = "codigoempresa>" + "'" + dato1.text + "' order by codigoempresa"
    op = 5
    sqlconta.response = campos
    Set sqlconta.conexion = conta
    Call sqlconta.sqlconta(op, condicion)
    If sqlconta.status = 0 Then carga: opciones.Visible = True: disponible (True): habilita (True): opciones.SetFocus
    
End Sub
Sub leeranterior()
    campos(0, 0) = dato1.Tag 'RUT
    campos(1, 0) = dato2.Tag 'NOMBRE
    campos(2, 0) = dato3.Tag 'DIRECCION
    campos(3, 0) = dato4.Tag 'COMUNA
    campos(4, 0) = DATO5.Tag '5
    campos(5, 0) = dato6.Tag
    campos(6, 0) = dato7.Tag
    campos(7, 0) = dato8.Tag
    campos(8, 0) = dato9.Tag
    campos(9, 0) = dato10.Tag
    campos(10, 0) = dato11.Tag
    campos(11, 0) = dato12.Tag
    campos(12, 0) = dato13.Tag
    campos(13, 0) = dato14.Tag
    campos(14, 0) = dato15.Tag
    campos(15, 0) = dato16.Tag
    campos(16, 0) = dato17.Tag
    campos(17, 0) = dato18.Tag
    campos(18, 0) = dato19.Tag
    campos(19, 0) = DATO20.Tag
    campos(20, 0) = DATO21.Tag
    campos(21, 0) = DATO22.Tag
    campos(22, 0) = "cuentabancaria"
    campos(23, 0) = "bancocuenta"
    campos(24, 0) = ""
    
    campos(0, 2) = "maestroempresas"
    condicion = "codigoempresa<" + "'" + dato1.text + "' order by codigoempresa"
    op = 5
    sqlconta.response = campos
    Set sqlconta.conexion = conta
    Call sqlconta.sqlconta(op, condicion)
    
    If sqlconta.status = 0 Then carga: opciones.Visible = True: disponible (True): habilita (True): opciones.SetFocus
 
End Sub

Sub carga()
    habilita (True)
    dato1.text = sqlconta.response(0, 3)
    dato2.text = sqlconta.response(1, 3)
    dato3.text = sqlconta.response(2, 3)
    dato4.text = sqlconta.response(3, 3)
    DATO5.text = sqlconta.response(4, 3)
    dato6.text = sqlconta.response(5, 3)
    dato7.text = sqlconta.response(6, 3)
    dato8.text = sqlconta.response(7, 3)
    dato9.text = sqlconta.response(8, 3)
    dato10.text = sqlconta.response(9, 3)
    dato11.text = sqlconta.response(10, 3)
    dato12.text = sqlconta.response(11, 3)
    dato13.text = sqlconta.response(12, 3)
    dato14.text = sqlconta.response(13, 3)
    dato15.text = sqlconta.response(14, 3)
    
    dato16.text = sqlconta.response(15, 3) '
    dato17.text = sqlconta.response(16, 3)
    dato18.text = sqlconta.response(17, 3)
    dato19.text = sqlconta.response(18, 3)
    DATO20.text = sqlconta.response(19, 3)
    DATO21.text = sqlconta.response(20, 3)
    DATO22.text = sqlconta.response(21, 3)
    dato23.text = sqlconta.response(22, 3)
    dato24.text = sqlconta.response(23, 3)
    DATO25.text = Format(sqlconta.response(24, 3), "dd")
    dato26.text = Format(sqlconta.response(24, 3), "mm")
    dato27.text = Format(sqlconta.response(24, 3), "yyyy")
    lblnombrebanco.Caption = leerbanco(sqlconta.response(23, 3))
    
    
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
    
    DATO25.Locked = condicion
    dato26.Locked = condicion
    dato27.Locked = condicion
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
    
    DATO25.Enabled = condicion
    dato26.Enabled = condicion
    dato27.Enabled = condicion
    
End Sub

Sub Pregunta(ByRef caja As TextBox, ByRef caja1 As TextBox)
    If caja.text = "" And sc = 0 Then caja.SetFocus
    If caja.text <> "" Or sc = 1 Then caja1.Enabled = True: caja1.SetFocus
    sc = 0
End Sub

Sub flechas(ByRef caja As TextBox, ByRef caja1 As TextBox, ByRef codigo As Integer)
    If codigo = 38 And caja.Enabled = True Then caja.SetFocus
    If codigo = 40 And caja1.Enabled = True Then caja1.SetFocus
End Sub
Sub grabar()

    campos(0, 0) = dato1.Tag 'CODIGO
    campos(1, 0) = dato2.Tag 'BODEGA
    campos(2, 0) = dato3.Tag 'DIRECCION
    campos(3, 0) = dato4.Tag 'CIUDAD
    campos(4, 0) = DATO5.Tag 'OTROS
    campos(5, 0) = dato6.Tag
    campos(6, 0) = dato7.Tag
    campos(7, 0) = dato8.Tag
    campos(8, 0) = dato9.Tag
    campos(9, 0) = dato10.Tag
    campos(10, 0) = dato11.Tag
    campos(11, 0) = dato12.Tag
    campos(12, 0) = dato13.Tag
    campos(13, 0) = dato14.Tag
    campos(14, 0) = dato15.Tag
    campos(15, 0) = dato16.Tag
    campos(16, 0) = dato17.Tag
    campos(17, 0) = dato18.Tag
    campos(18, 0) = dato19.Tag
    campos(19, 0) = DATO20.Tag
    campos(20, 0) = DATO21.Tag
    campos(21, 0) = DATO22.Tag
    campos(22, 0) = "cuentabancaria"
    campos(23, 0) = "bancocuenta"
    campos(24, 0) = "flujoactualizado"
    campos(25, 0) = ""
    
  
    campos(0, 1) = dato1.text 'CODIGO
    campos(1, 1) = dato2.text 'BODEGA
    campos(2, 1) = dato3.text 'DIRECCION
    campos(3, 1) = dato4.text 'CIUDAD
    campos(4, 1) = DATO5.text 'OTROS
    campos(5, 1) = dato6.text
    campos(6, 1) = dato7.text
    campos(7, 1) = dato8.text
    campos(8, 1) = dato9.text
    campos(9, 1) = dato10.text
    campos(10, 1) = dato11.text
    campos(11, 1) = dato12.text
    campos(12, 1) = dato13.text
    campos(13, 1) = dato14.text
    campos(14, 1) = dato15.text
    campos(15, 1) = dato16.text
    campos(16, 1) = dato17.text
    campos(17, 1) = dato18.text
    campos(18, 1) = dato19.text
    campos(19, 1) = DATO20.text
    campos(20, 1) = DATO21.text
    campos(21, 1) = DATO22.text
    campos(22, 1) = dato23.text
    campos(23, 1) = dato24.text
    campos(24, 1) = Format(DATO25 & "-" & dato26 & "-" & dato27, "yyyy-mm-dd")
    
    
    campos(0, 2) = "maestroempresas"
    If MODIFI = 1 Then condicion = "codigoempresa=" + "'" + dato1.text + "'"
    If MODIFI = 1 Then op = 3 Else op = 2
    MODIFI = 0
    sqlconta.response = campos
    Set sqlconta.conexion = conta
    
    Call sqlconta.sqlconta(op, condicion)
    status = sqlconta.status

     If op = 2 Then crearTablascontables
     


End Sub
Sub ELIMINAR()
    Dim respu As Boolean
    
    If MsgBox("ESTA SITUACION ES CRITICA PODRIA ESTAR ELIMINANDO TODA LA CONTABILIDAD DESEA ELIMINAR ", vbYesNo) = vbYes Then
    
    campos(0, 2) = "maestroempresas"
    condicion = "codigoempresa=" + "'" + dato1.text + "'"
    op = 4
    sqlconta.response = campos
    Set sqlconta.conexion = conta
   
    Call sqlconta.sqlconta(op, condicion)
    
    End If
    
End Sub


Private Sub opciones_FSCommand(ByVal command As String, ByVal args As String)
If command = "retorno" Then retorno
If command = "modifica" Then modifica

If command = "elimina" Then
If Verifica_Permiso(Me.Caption, "elimina") = True Then
                      
                   
disponible (True)
habilita (False)
ELIMINAR
limpia
opciones.Visible = False
dato1.SetFocus
 Else
                       MsgBox mensaje_nopermiso, vbCritical + vbOKOnly, "Permiso Denegado"
 End If
End If
If command = "siguiente" Then leersiguiente
If command = "anterior" Then leeranterior
End Sub
Sub modifica()
disponible (True)
habilita (False)
dato1.Enabled = False
MODIFI = 1

dato2.SetFocus
End Sub
Sub retorno()
disponible (True)
habilita (False)
limpia
opciones.Visible = False
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
    DATO21.text = ""
    dato23.text = ""
    dato24.text = ""
    DATO22.text = ""
    DATO25.text = ""
    dato26.text = ""
    dato27.text = ""
    lblnombrebanco.Caption = ""
    
    
    
    End Sub
Sub ayudaempresa(ByRef caja As TextBox)
    Dim campos As Variant
    Dim cfijo As Variant
    Dim largo As Variant
    campos = Array("codigoempresa", "nombre")
    largo = Array("6s", "40s")
    cfijo = "no"
    basebus = clientesistema + "conta"
    Call cargaAyudaT(Servidor, basebus, Usuario, password, "maestroempresas", dato1, campos, cfijo, largo, 2)
    caja.Enabled = True
    caja.SetFocus
    
End Sub

Private Sub crearTablascontables()
    Dim tabla As String
    Dim csql As New rdoQuery
    Dim database As String
    Exit Sub
    'GESTION
    Call ConectarCreacion
    Set csql.ActiveConnection = con
    
    csql.sql = "CREATE DATABASE IF NOT EXISTS " & clientesistema + "conta" & dato1.text & ""
    csql.Execute
    
    csql.sql = "CREATE TABLE IF NOT EXISTS " & clientesistema + "conta" & dato1.text & ".boletasdehonorarios "
    csql.sql = csql.sql & "(LIKE boletasdehonorarios)"
    csql.Execute
    
    csql.sql = "CREATE TABLE IF NOT EXISTS " & clientesistema + "conta" & dato1.text & ".boletasdehonorarios_detalle "
    csql.sql = csql.sql & "(LIKE boletasdehonorarios_detalle)"
    csql.Execute
    
    csql.sql = "CREATE TABLE IF NOT EXISTS " & clientesistema + "conta" & dato1.text & ".boletasdeventa "
    csql.sql = csql.sql & "(LIKE boletasdeventa)"
    csql.Execute
    
    csql.sql = "CREATE TABLE IF NOT EXISTS " & clientesistema + "conta" & dato1.text & ".centrosdecosto "
    csql.sql = csql.sql & "(LIKE centrosdecosto)"
    csql.Execute
    
    csql.sql = "CREATE TABLE IF NOT EXISTS " & clientesistema + "conta" & dato1.text & ".chequesdocumento "
    csql.sql = csql.sql & "(LIKE chequesdocumento)"
    csql.Execute
    
    csql.sql = "CREATE TABLE IF NOT EXISTS " & clientesistema + "conta" & dato1.text & ".chequesmovimiento "
    csql.sql = csql.sql & "(LIKE chequesmovimiento)"
    csql.Execute
    
    csql.sql = "CREATE TABLE IF NOT EXISTS " & clientesistema + "conta" & dato1.text & ".cuentascorrientes "
    csql.sql = csql.sql & "(LIKE cuentascorrientes)"
    csql.Execute
    
    csql.sql = "CREATE TABLE IF NOT EXISTS " & clientesistema + "conta" & dato1.text & ".cuentasdelmayor "
    csql.sql = csql.sql & "(LIKE cuentasdelmayor)"
    csql.Execute
    
    csql.sql = "CREATE TABLE IF NOT EXISTS " & clientesistema + "conta" & dato1.text & ".facturasdecompras "
    csql.sql = csql.sql & "(LIKE facturasdecompras)"
    csql.Execute
    
    csql.sql = "CREATE TABLE IF NOT EXISTS " & clientesistema + "conta" & dato1.text & ".facturasdecompras_impuestos "
    csql.sql = csql.sql & "(LIKE facturasdecompras_impuestos)"
    csql.Execute
    
    csql.sql = "CREATE TABLE IF NOT EXISTS " & clientesistema + "conta" & dato1.text & ".facturasdecompras_detalle "
    csql.sql = csql.sql & "(LIKE facturasdecompras_detalle)"
    csql.Execute
    
    csql.sql = "CREATE TABLE IF NOT EXISTS " & clientesistema + "conta" & dato1.text & ".facturasdeventas "
    csql.sql = csql.sql & "(LIKE facturasdeventas)"
    csql.Execute
    
    csql.sql = "CREATE TABLE IF NOT EXISTS " & clientesistema + "conta" & dato1.text & ".facturasdeventas_impuestos "
    csql.sql = csql.sql & "(LIKE facturasdeventas_impuestos)"
    csql.Execute
    
    csql.sql = "CREATE TABLE IF NOT EXISTS " & clientesistema + "conta" & dato1.text & ".facturasdeventas_detalle "
    csql.sql = csql.sql & "(LIKE facturasdeventas_detalle)"
    csql.Execute
    
    csql.sql = "CREATE TABLE IF NOT EXISTS " & clientesistema + "conta" & dato1.text & ".movimientos_glosa "
    csql.sql = csql.sql & "(LIKE movimientos_glosa)"
    csql.Execute
    
    csql.sql = "CREATE TABLE IF NOT EXISTS " & clientesistema + "conta" & dato1.text & ".movimientoscontables "
    csql.sql = csql.sql & "(LIKE movimientoscontables)"
    csql.Execute
    
    csql.sql = "CREATE TABLE IF NOT EXISTS " & clientesistema + "conta" & dato1.text & ".saldoscentrosdecosto "
    csql.sql = csql.sql & "(LIKE saldoscentrosdecosto)"
    csql.Execute
    
    csql.sql = "CREATE TABLE IF NOT EXISTS " & clientesistema + "conta" & dato1.text & ".saldosctacte "
    csql.sql = csql.sql & "(LIKE saldosctacte)"
    csql.Execute
    
    csql.sql = "CREATE TABLE IF NOT EXISTS " & clientesistema + "conta" & dato1.text & ".saldosdelmayor "
    csql.sql = csql.sql & "(LIKE saldosdelmayor)"
    csql.Execute
    
    
    con.Close
End Sub
Private Sub ConectarCreacion()
    Dim cadena_conexion As String
    
    cadena_conexion = "DRIVER={MySQL ODBC 3.51 Driver}; SERVER=" & Servidor & "; DATABASE=" & clientesistema + "conta08" & "; PWD=" & password & "; UID=" & Usuario & ";OPTION=3"
    Set con = New rdoConnection
    con.Connect = cadena_conexion
    con.CursorDriver = rdUseServer
    con.EstablishConnection
    
End Sub
Private Sub eliminatablas()
Dim tabla As String
    Dim csql As New rdoQuery
    Dim database As String
    
    'GESTION
    Call ConectarCreacion
    Set csql.ActiveConnection = con
    
    csql.sql = "DROP DATABASE IF EXISTS " & clientesistema + "conta" & dato1.text
    csql.Execute
       
End Sub
 
