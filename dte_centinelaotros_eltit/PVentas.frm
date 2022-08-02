VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash10d.ocx"
Object = "{4777436C-EB8C-4596-98A8-EBCF98683969}#1.0#0"; "Flexcell.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{ADBBAED6-B16F-40DC-80DF-B44910CBA76C}#1.0#0"; "Frame-Xp.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form PVentas 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pantalla de Ventas"
   ClientHeight    =   9945
   ClientLeft      =   240
   ClientTop       =   810
   ClientWidth     =   14610
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9945
   ScaleWidth      =   14610
   Begin FlexCell.Grid Grid1 
      Height          =   255
      Left            =   6840
      TabIndex        =   101
      Top             =   9360
      Visible         =   0   'False
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   450
      Cols            =   5
      DefaultFontSize =   8.25
      Rows            =   30
   End
   Begin XPFrame.FrameXp FrameXp1 
      Height          =   1575
      Left            =   5520
      TabIndex        =   81
      Top             =   7680
      Width           =   3915
      _ExtentX        =   6906
      _ExtentY        =   2778
      BackColor       =   8421504
      Caption         =   "IMPUESTOS ADICIONALES"
      CaptionEstilo3D =   1
      BackColor       =   8421504
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.TextBox dato15 
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
         Left            =   720
         MaxLength       =   9
         TabIndex        =   91
         Tag             =   "proveedor"
         Text            =   "0"
         Top             =   1080
         Width           =   1215
      End
      Begin VB.TextBox dato13 
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
         Left            =   720
         MaxLength       =   9
         TabIndex        =   85
         Tag             =   "proveedor"
         Text            =   "0"
         Top             =   360
         Width           =   1215
      End
      Begin VB.TextBox dato16 
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
         Left            =   2640
         MaxLength       =   20
         TabIndex        =   84
         Tag             =   "proveedor"
         Text            =   "0"
         Top             =   360
         Width           =   1215
      End
      Begin VB.TextBox dato17 
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
         Left            =   2640
         MaxLength       =   9
         TabIndex        =   83
         Tag             =   "proveedor"
         Text            =   "0"
         Top             =   720
         Width           =   1215
      End
      Begin VB.TextBox dato14 
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
         Left            =   720
         MaxLength       =   5
         TabIndex        =   82
         Tag             =   "proveedor"
         Text            =   "0"
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label lbl42 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Carne"
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
         Height          =   285
         Left            =   2040
         TabIndex        =   90
         Top             =   360
         Width           =   855
      End
      Begin VB.Label lbl43 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Vinos"
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
         Height          =   285
         Left            =   120
         TabIndex        =   89
         Top             =   720
         Width           =   855
      End
      Begin VB.Label lbl41 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Refre."
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
         Height          =   285
         Left            =   120
         TabIndex        =   88
         Top             =   360
         Width           =   855
      End
      Begin VB.Label lbl44 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Harina"
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
         Height          =   285
         Left            =   2040
         TabIndex        =   87
         Top             =   720
         Width           =   855
      End
      Begin VB.Label lbl40 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Licores"
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
         Height          =   285
         Left            =   90
         TabIndex        =   86
         Top             =   1080
         Width           =   855
      End
   End
   Begin XPFrame.FrameXp FRMPREVENTA 
      Height          =   1815
      Left            =   1395
      TabIndex        =   62
      Top             =   3870
      Width           =   9495
      _ExtentX        =   16748
      _ExtentY        =   3201
      BackColor       =   49344
      Caption         =   "FORMULARIO AGREGAR PREVENTAS"
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
      Begin VB.TextBox PREVENTA 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   6720
         MaxLength       =   10
         TabIndex        =   64
         Top             =   720
         Width           =   2535
      End
      Begin MSForms.CheckBox CheckBox1 
         Height          =   60
         Left            =   7830
         TabIndex        =   69
         Top             =   855
         Width           =   60
         BackColor       =   16761024
         ForeColor       =   -2147483630
         DisplayStyle    =   4
         Size            =   "106;106"
         Value           =   "0"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "INGRESE NUMERO DE PREVENTA"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   495
         TabIndex        =   63
         Top             =   765
         Width           =   6015
      End
   End
   Begin XPFrame.FrameXp frmTipo 
      Height          =   1920
      Left            =   2520
      TabIndex        =   57
      Top             =   360
      Visible         =   0   'False
      Width           =   2715
      _ExtentX        =   4789
      _ExtentY        =   3387
      BackColor       =   8421631
      Caption         =   "Tipo Documento"
      CaptionEstilo3D =   1
      BackColor       =   8421631
      ColorBarraArriba=   12632319
      ColorBarraAbajo =   128
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
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " 2 - Boleta"
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
         Height          =   285
         Left            =   135
         TabIndex        =   66
         Top             =   840
         Width           =   2475
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " 4 - NC Boleta"
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
         Height          =   285
         Left            =   120
         TabIndex        =   65
         Top             =   1560
         Width           =   2475
      End
      Begin VB.Label lbl24 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " 3 - NC Factura"
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
         Height          =   285
         Left            =   120
         TabIndex        =   59
         Top             =   1200
         Width           =   2475
      End
      Begin VB.Label lbl23 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " 1 - Factura"
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
         Height          =   285
         Left            =   135
         TabIndex        =   58
         Top             =   495
         Width           =   2475
      End
   End
   Begin XPFrame.FrameXp frmDatos 
      Height          =   3420
      Left            =   0
      TabIndex        =   20
      Top             =   0
      Width           =   14895
      _ExtentX        =   26273
      _ExtentY        =   6033
      BackColor       =   16744576
      Caption         =   "Datos de la Venta"
      CaptionEstilo3D =   1
      BackColor       =   16744576
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
      Begin VB.ComboBox vendedores 
         BackColor       =   &H00FF8080&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   105
         Top             =   2670
         Width           =   6615
      End
      Begin VB.TextBox DatoHora 
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
         Left            =   3840
         Locked          =   -1  'True
         MaxLength       =   8
         TabIndex        =   103
         Tag             =   "proveedor"
         Top             =   720
         Width           =   1215
      End
      Begin VB.TextBox dato50 
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
         Left            =   1080
         MaxLength       =   9
         TabIndex        =   10
         Tag             =   "proveedor"
         Top             =   3060
         Width           =   1455
      End
      Begin VB.TextBox dato27 
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
         Left            =   11880
         MaxLength       =   10
         TabIndex        =   5
         Tag             =   "proveedor"
         Top             =   720
         Width           =   1455
      End
      Begin VB.TextBox dato30 
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
         Left            =   5340
         MaxLength       =   2
         TabIndex        =   76
         Tag             =   "proveedor"
         Top             =   360
         Width           =   495
      End
      Begin VB.TextBox dato21 
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
         Left            =   8685
         MaxLength       =   20
         TabIndex        =   13
         Tag             =   "proveedor"
         Top             =   2685
         Width           =   4695
      End
      Begin VB.TextBox dato20 
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
         Left            =   8685
         MaxLength       =   20
         TabIndex        =   12
         Tag             =   "proveedor"
         Top             =   2350
         Width           =   4695
      End
      Begin VB.TextBox dato1 
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
         Left            =   1920
         Locked          =   -1  'True
         MaxLength       =   2
         TabIndex        =   0
         Tag             =   "proveedor"
         Top             =   360
         Width           =   495
      End
      Begin VB.TextBox dato8 
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
         Left            =   1935
         MaxLength       =   20
         TabIndex        =   8
         Tag             =   "proveedor"
         Top             =   2020
         Width           =   4830
      End
      Begin VB.TextBox dato6 
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
         Left            =   1935
         MaxLength       =   9
         TabIndex        =   6
         Tag             =   "proveedor"
         Top             =   1035
         Width           =   1455
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
         Left            =   1965
         MaxLength       =   4
         TabIndex        =   3
         Tag             =   "proveedor"
         Top             =   720
         Width           =   615
      End
      Begin VB.TextBox dato9 
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
         Left            =   8685
         MaxLength       =   20
         TabIndex        =   11
         Tag             =   "proveedor"
         Top             =   2020
         Width           =   4695
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
         Left            =   1605
         MaxLength       =   2
         TabIndex        =   2
         Tag             =   "proveedor"
         Top             =   720
         Width           =   375
      End
      Begin VB.TextBox dato3 
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
         Left            =   1245
         MaxLength       =   2
         TabIndex        =   1
         Tag             =   "proveedor"
         Top             =   720
         Width           =   375
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
         Left            =   9120
         MaxLength       =   10
         TabIndex        =   4
         Tag             =   "proveedor"
         Top             =   705
         Width           =   1455
      End
      Begin VB.TextBox dato10 
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
         Left            =   1080
         MaxLength       =   9
         TabIndex        =   9
         Tag             =   "proveedor"
         Top             =   2350
         Width           =   1455
      End
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
         Left            =   5625
         MaxLength       =   1
         TabIndex        =   7
         Tag             =   "proveedor"
         Top             =   1035
         Width           =   615
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " HORA"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   2640
         TabIndex        =   104
         Top             =   720
         Width           =   1020
      End
      Begin VB.Label Label10 
         Appearance      =   0  'Flat
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Cajera"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   120
         TabIndex        =   102
         Top             =   3060
         Width           =   885
      End
      Begin VB.Label lbldvcajera 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
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
         ForeColor       =   &H0080FFFF&
         Height          =   285
         Left            =   2565
         TabIndex        =   99
         Top             =   3060
         Width           =   375
      End
      Begin VB.Label lblcajera 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
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
         ForeColor       =   &H0080FFFF&
         Height          =   285
         Left            =   3000
         TabIndex        =   98
         Top             =   3060
         Width           =   3705
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "INTERNO"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   10800
         TabIndex        =   97
         Top             =   705
         Width           =   975
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "FISCAL"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   8400
         TabIndex        =   96
         Top             =   705
         Width           =   735
      End
      Begin VB.Label lbl35 
         Appearance      =   0  'Flat
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Caja"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   4710
         TabIndex        =   75
         Top             =   360
         Width           =   615
      End
      Begin VB.Label lblnombrecaja 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
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
         ForeColor       =   &H0080FFFF&
         Height          =   285
         Left            =   5880
         TabIndex        =   74
         Top             =   360
         Width           =   3300
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Cantidad de Bultos"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   6885
         TabIndex        =   68
         Top             =   2685
         Width           =   1695
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Revisado por"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   6885
         TabIndex        =   67
         Top             =   2350
         Width           =   1695
      End
      Begin VB.Label lblDocumento 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
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
         ForeColor       =   &H0080FFFF&
         Height          =   285
         Left            =   2430
         TabIndex        =   28
         Top             =   360
         Width           =   2235
      End
      Begin VB.Label lbl2 
         Appearance      =   0  'Flat
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Número:"
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   9240
         TabIndex        =   40
         Top             =   360
         Width           =   4095
      End
      Begin VB.Label lbl12 
         Appearance      =   0  'Flat
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Vendedor"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   120
         TabIndex        =   39
         Top             =   2350
         Width           =   885
      End
      Begin VB.Label lbl11 
         Appearance      =   0  'Flat
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Transporte"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   6885
         TabIndex        =   38
         Top             =   2020
         Width           =   1695
      End
      Begin VB.Label lbl9 
         Appearance      =   0  'Flat
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Condiciones de Pago"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   120
         TabIndex        =   37
         Top             =   2020
         Width           =   1695
      End
      Begin VB.Label lbl6 
         Appearance      =   0  'Flat
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Dirección"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   120
         TabIndex        =   36
         Top             =   1365
         Width           =   1695
      End
      Begin VB.Label lbl5 
         Appearance      =   0  'Flat
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Razón Social"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   6840
         TabIndex        =   35
         Top             =   1035
         Width           =   1815
      End
      Begin VB.Label lbl4 
         Appearance      =   0  'Flat
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Rut Cliente"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   120
         TabIndex        =   34
         Top             =   1035
         Width           =   1695
      End
      Begin VB.Label lbl3 
         Appearance      =   0  'Flat
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Fecha"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   120
         TabIndex        =   33
         Top             =   720
         Width           =   1020
      End
      Begin VB.Label lbl1 
         Appearance      =   0  'Flat
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Tipo"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   120
         TabIndex        =   32
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label lbl7 
         Appearance      =   0  'Flat
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Comuna"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   120
         TabIndex        =   31
         Top             =   1695
         Width           =   1695
      End
      Begin VB.Label lbl8 
         Appearance      =   0  'Flat
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Ciudad"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   6885
         TabIndex        =   30
         Top             =   1695
         Width           =   1695
      End
      Begin VB.Label lblDV 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
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
         ForeColor       =   &H0080FFFF&
         Height          =   285
         Left            =   3375
         TabIndex        =   29
         Top             =   1035
         Width           =   375
      End
      Begin VB.Label lblRazon 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
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
         ForeColor       =   &H0080FFFF&
         Height          =   285
         Left            =   8760
         TabIndex        =   27
         Top             =   1035
         Width           =   4575
      End
      Begin VB.Label lblDireccion 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
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
         ForeColor       =   &H0080FFFF&
         Height          =   285
         Left            =   1935
         TabIndex        =   26
         Top             =   1365
         Width           =   11415
      End
      Begin VB.Label lblComuna 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
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
         ForeColor       =   &H0080FFFF&
         Height          =   285
         Left            =   1935
         TabIndex        =   25
         Top             =   1695
         Width           =   4800
      End
      Begin VB.Label lblCiudad 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
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
         ForeColor       =   &H0080FFFF&
         Height          =   285
         Left            =   8685
         TabIndex        =   24
         Top             =   1695
         Width           =   4695
      End
      Begin VB.Label lblVendedor 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
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
         ForeColor       =   &H0080FFFF&
         Height          =   285
         Left            =   3000
         TabIndex        =   23
         Top             =   2355
         Width           =   3735
      End
      Begin VB.Label lbl22 
         Appearance      =   0  'Flat
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Sucursal"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   3870
         TabIndex        =   22
         Top             =   1035
         Width           =   1695
      End
      Begin VB.Label lblDVV 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
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
         ForeColor       =   &H0080FFFF&
         Height          =   285
         Left            =   2565
         TabIndex        =   21
         Top             =   2355
         Width           =   375
      End
      Begin VB.Label lblfoliofiscal 
         BackColor       =   &H80000007&
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
         ForeColor       =   &H0000FFFF&
         Height          =   285
         Left            =   14505
         TabIndex        =   80
         Top             =   600
         Visible         =   0   'False
         Width           =   1500
      End
      Begin VB.Label lblfoliointerno 
         BackColor       =   &H80000007&
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
         ForeColor       =   &H0000FFFF&
         Height          =   285
         Left            =   14505
         TabIndex        =   79
         Top             =   240
         Visible         =   0   'False
         Width           =   1500
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Folio Fiscal"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   13155
         TabIndex        =   78
         Top             =   600
         Visible         =   0   'False
         Width           =   1230
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Folio Interno"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   12675
         TabIndex        =   77
         Top             =   240
         Visible         =   0   'False
         Width           =   1710
      End
   End
   Begin XPFrame.FrameXp frmDetalle 
      Height          =   4215
      Left            =   0
      TabIndex        =   19
      Top             =   3480
      Width           =   14655
      _ExtentX        =   25850
      _ExtentY        =   7435
      BackColor       =   16744576
      Caption         =   "Lista de Productos"
      CaptionEstilo3D =   1
      BackColor       =   16744576
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
      Begin FlexCell.Grid detalle 
         Height          =   3705
         Left            =   0
         TabIndex        =   14
         Top             =   360
         Width           =   14535
         _ExtentX        =   25638
         _ExtentY        =   6535
         Cols            =   5
         DefaultFontSize =   8.25
         Rows            =   30
      End
   End
   Begin VB.TextBox pivote 
      Height          =   285
      Left            =   120
      MaxLength       =   13
      TabIndex        =   18
      Top             =   2760
      Visible         =   0   'False
      Width           =   735
   End
   Begin MSAdodcLib.Adodc data 
      Height          =   330
      Left            =   180
      Top             =   6720
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin XPFrame.FrameXp frmResumen 
      Height          =   2175
      Left            =   9480
      TabIndex        =   41
      Top             =   7680
      Width           =   5115
      _ExtentX        =   9022
      _ExtentY        =   3836
      BackColor       =   16744576
      Caption         =   "Resumen"
      CaptionEstilo3D =   1
      BackColor       =   16744576
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
      Begin VB.TextBox dato18 
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
         Left            =   3840
         MaxLength       =   7
         TabIndex        =   95
         Tag             =   "proveedor"
         Text            =   "0"
         Top             =   480
         Width           =   1215
      End
      Begin VB.TextBox dato26 
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
         Left            =   3840
         MaxLength       =   20
         TabIndex        =   94
         Tag             =   "proveedor"
         Text            =   "0"
         Top             =   1560
         Width           =   1215
      End
      Begin VB.TextBox dato19 
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
         Left            =   3840
         MaxLength       =   7
         TabIndex        =   93
         Tag             =   "proveedor"
         Text            =   "0"
         Top             =   840
         Width           =   1215
      End
      Begin VB.TextBox dato25 
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
         Left            =   3840
         MaxLength       =   9
         TabIndex        =   92
         Tag             =   "proveedor"
         Text            =   "0"
         Top             =   1200
         Width           =   1215
      End
      Begin VB.TextBox dato12 
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
         Left            =   1440
         MaxLength       =   9
         TabIndex        =   16
         Tag             =   "proveedor"
         Text            =   "0"
         Top             =   840
         Width           =   1215
      End
      Begin VB.TextBox dato11 
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
         Left            =   1440
         MaxLength       =   5
         TabIndex        =   15
         Tag             =   "proveedor"
         Text            =   "0"
         Top             =   480
         Width           =   1215
      End
      Begin VB.TextBox dato22 
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
         Left            =   1440
         MaxLength       =   20
         TabIndex        =   70
         Tag             =   "proveedor"
         Text            =   "0"
         Top             =   1200
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label lbl19 
         Appearance      =   0  'Flat
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Total"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   2760
         TabIndex        =   48
         Top             =   1560
         Width           =   975
      End
      Begin VB.Label lbl18 
         Appearance      =   0  'Flat
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "EXENTO"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   2760
         TabIndex        =   52
         Top             =   1200
         Width           =   975
      End
      Begin VB.Label lbl17 
         Appearance      =   0  'Flat
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " IVA"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   2760
         TabIndex        =   47
         Top             =   840
         Width           =   975
      End
      Begin VB.Label lbl16 
         Appearance      =   0  'Flat
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Neto"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   2760
         TabIndex        =   46
         Top             =   480
         Width           =   975
      End
      Begin VB.Label lbl15 
         Appearance      =   0  'Flat
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Descuento ($)"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   120
         TabIndex        =   45
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label lbl14 
         Appearance      =   0  'Flat
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Descuento (%)"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   120
         TabIndex        =   44
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label lbl27 
         Appearance      =   0  'Flat
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Abono ($)"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   120
         TabIndex        =   71
         Top             =   1200
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label lblTotal 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FFFF&
         Height          =   285
         Left            =   6420
         TabIndex        =   50
         Top             =   1080
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label lblIHA 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FFFF&
         Height          =   285
         Left            =   6420
         TabIndex        =   51
         Top             =   720
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label lblIVA 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FFFF&
         Height          =   285
         Left            =   6420
         TabIndex        =   49
         Top             =   360
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label lblNeto 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FFFF&
         Height          =   285
         Left            =   6420
         TabIndex        =   43
         Top             =   0
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label lblSub 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FFFF&
         Height          =   285
         Left            =   4080
         TabIndex        =   42
         Top             =   0
         Visible         =   0   'False
         Width           =   1215
      End
   End
   Begin FlexCell.Grid impresion 
      Height          =   495
      Left            =   1320
      TabIndex        =   53
      Top             =   6540
      Visible         =   0   'False
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   873
      Cols            =   5
      DefaultFontSize =   8.25
      Rows            =   30
   End
   Begin VB.PictureBox manual 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   0
      ScaleHeight     =   375
      ScaleWidth      =   555
      TabIndex        =   56
      Top             =   0
      Width           =   555
   End
   Begin XPFrame.FrameXp glosas 
      Height          =   975
      Left            =   3120
      TabIndex        =   72
      Top             =   3720
      Visible         =   0   'False
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   1720
      BackColor       =   16744576
      Caption         =   "Glosa"
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
      Begin VB.TextBox dato24 
         BackColor       =   &H00FAF0E7&
         Height          =   495
         Left            =   120
         TabIndex        =   73
         Text            =   "  "
         Top             =   300
         Width           =   6375
      End
   End
   Begin FlexCell.Grid impresionboleta 
      Height          =   495
      Left            =   0
      TabIndex        =   100
      Top             =   3960
      Visible         =   0   'False
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   873
      BackColorBkg    =   -2147483644
      BackColorSel    =   16777215
      Cols            =   5
      DefaultFontSize =   8.25
      Rows            =   30
   End
   Begin VB.Label lbl30 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " 'P' Ver Pago"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   120
      TabIndex        =   61
      Top             =   9120
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Label lbl26 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " * Fin Venta"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   120
      TabIndex        =   60
      Top             =   9480
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Label lbl20 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " - Elimina Linea"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   2040
      TabIndex        =   55
      Top             =   9480
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Label lblNulo 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   54
      Top             =   7560
      Width           =   7935
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash opciones 
      Height          =   1335
      Left            =   0
      TabIndex        =   17
      Top             =   7800
      Visible         =   0   'False
      Width           =   5535
      _cx             =   9763
      _cy             =   2355
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
Attribute VB_Name = "PVentas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Private c As Cliente
    Private v As venta
    Private p As pagos
    Private formatogrilla(10, 10) As String
    Private modifica As Boolean
    Private vacio As Boolean
    Private fila As Long
    Private columna As Long
    Private nula As Boolean
    Private lectura As Boolean
    Public imprimio As Boolean
    Public DESDE As String
    Public HASTA As String
    Public dire As Integer
    Private tipoprecio As String
    Private fecha As String
    Private cupo As Double
    Private costo(1000) As Double
    Private compara As Double
    Private compara2 As Double
    Private GLOSA As String
    Private ok As Integer
    Private caja As String
    Private i As Double
    

    'Private segurity As Boolean

'============================================================
'MANEJO DE LOS CONTOLES
'============================================================
    '========================================================
    'GotFocus
    '========================================================
    Private Sub dato1_GotFocus()
        lectura = False
        modifica = False
        
        frmTipo.Visible = True
        Call selecciona(dato1)
    End Sub

Private Sub dato10_Change()
            If dato10.text <> "" Then
            'dato10.text = ceros(dato10)
            lblDVV.Caption = rut(dato10.text)
            lblvendedor.Caption = leerNombreVendedor(dato10.text & lblDVV.Caption)
            If lblvendedor.Caption <> "" Then
               dato50.SetFocus
            End If
        End If
End Sub

  Private Sub dato13_GotFocus()
 Call VerificarCajas(Me, dato13)
 Call selecciona(dato13)
End Sub
Private Sub dato13_KeyDown(KeyCode As Integer, Shift As Integer)
Call Flechas(KeyCode, dato12)
End Sub

Private Sub dato13_KeyPress(KeyAscii As Integer)
  KeyAscii = esNumero(KeyAscii)
  If KeyAscii = 13 And dato13.text <> "" Then
  dato14.SetFocus
  End If
End Sub
Private Sub dato14_GotFocus()
 Call VerificarCajas(Me, dato14)
 Call selecciona(dato14)
End Sub
Private Sub dato14_KeyDown(KeyCode As Integer, Shift As Integer)
Call Flechas(KeyCode, dato13)
End Sub

Private Sub dato14_KeyPress(KeyAscii As Integer)
  KeyAscii = esNumero(KeyAscii)
  If KeyAscii = 13 And dato14.text <> "" Then
  dato15.SetFocus
  End If
End Sub
Private Sub dato15_GotFocus()
 Call VerificarCajas(Me, dato15)
 Call selecciona(dato15)
End Sub
Private Sub dato15_KeyDown(KeyCode As Integer, Shift As Integer)
Call Flechas(KeyCode, dato14)
End Sub

Private Sub dato15_KeyPress(KeyAscii As Integer)
  KeyAscii = esNumero(KeyAscii)
  If KeyAscii = 13 And dato15.text <> "" Then
  dato16.SetFocus
  End If
End Sub

Private Sub dato16_GotFocus()
 Call VerificarCajas(Me, dato16)
 Call selecciona(dato16)
End Sub
Private Sub dato16_KeyDown(KeyCode As Integer, Shift As Integer)
Call Flechas(KeyCode, dato15)
End Sub

Private Sub dato16_KeyPress(KeyAscii As Integer)
  KeyAscii = esNumero(KeyAscii)
  If KeyAscii = 13 And dato16.text <> "" Then
  dato17.SetFocus
  End If
End Sub
Private Sub dato17_GotFocus()
 Call VerificarCajas(Me, dato17)
 Call selecciona(dato17)
End Sub
Private Sub dato17_KeyDown(KeyCode As Integer, Shift As Integer)
Call Flechas(KeyCode, dato16)
End Sub

Private Sub dato17_KeyPress(KeyAscii As Integer)
KeyAscii = esNumero(KeyAscii)
If KeyAscii = 13 And dato17.text <> "" Then
 dato18.SetFocus
End If
End Sub
Private Sub dato18_GotFocus()
 Call VerificarCajas(Me, dato18)
 Call selecciona(dato18)
End Sub
Private Sub dato18_KeyDown(KeyCode As Integer, Shift As Integer)
Call Flechas(KeyCode, dato17)
End Sub

Private Sub dato18_KeyPress(KeyAscii As Integer)
KeyAscii = esNumero(KeyAscii)
If KeyAscii = 13 And dato18.text <> "" Then
 dato19.SetFocus
End If
End Sub

Private Sub dato19_GotFocus()
 Call VerificarCajas(Me, dato19)
 Call selecciona(dato19)
End Sub
Private Sub dato19_KeyDown(KeyCode As Integer, Shift As Integer)
Call Flechas(KeyCode, dato19)
End Sub

Private Sub dato19_KeyPress(KeyAscii As Integer)
KeyAscii = esNumero(KeyAscii)
If KeyAscii = 13 And dato19.text <> "" Then
 dato25.SetFocus
End If
End Sub
Private Sub dato25_GotFocus()
 Call VerificarCajas(Me, dato25)
 Call selecciona(dato25)
End Sub
Private Sub dato25_KeyDown(KeyCode As Integer, Shift As Integer)
Call Flechas(KeyCode, dato25)
End Sub

Private Sub dato25_KeyPress(KeyAscii As Integer)
KeyAscii = esNumero(KeyAscii)
If KeyAscii = 13 And dato25.text <> "" Then
 dato26.SetFocus
End If
End Sub

Private Sub dato26_GotFocus()
 Call VerificarCajas(Me, dato26)
 Call selecciona(dato26)
End Sub
Private Sub dato26_KeyDown(KeyCode As Integer, Shift As Integer)
Call Flechas(KeyCode, dato26)
End Sub
Private Sub dato26_KeyPress(KeyAscii As Integer)
KeyAscii = esNumero(KeyAscii)
If KeyAscii = 13 And dato26.text <> "" Then
modifica = False
            Call ctrltostruct(True)
            lbl26.Visible = False
            lbl20.Visible = False
            detallePagos.Show
End If
End Sub
    Private Sub dato2_GotFocus()
        Call VerificarCajas(Me, dato2)
        Call selecciona(dato2)
    End Sub
    
Private Sub dato22_GotFocus()
   Call VerificarCajas(Me, dato22)
   Call selecciona(dato22)
End Sub

Private Sub dato22_KeyDown(KeyCode As Integer, Shift As Integer)
 Call Flechas(KeyCode, dato11)
End Sub

Private Sub dato22_KeyPress(KeyAscii As Integer)
  KeyAscii = esNumero(KeyAscii)
   If KeyAscii = 13 And dato22.text <> "" Then
   dato12.SetFocus
   End If
End Sub

Private Sub dato24_KeyPress(KeyAscii As Integer)
 KeyAscii = Asc(UCase(Chr(KeyAscii)))
If KeyAscii = 13 And dato24.text <> "" Then
detalle.Cell(detalle.ActiveCell.row, 9).text = dato24.text
dato24.text = ""
glosas.Visible = False
dato10.SetFocus
' SendKeys "{Tab}"
' SendKeys "{Tab}"
' SendKeys "{Tab}"
' SendKeys "{Tab}"

detalle.Cell(detalle.Rows - 1, 3).SetFocus

End If

End Sub


Private Sub dato27_Change()

Call LeerVendedorDocumento(empresaActiva, dato1.text, dato27.text, dato30.text, dato5.text & "-" & dato4.text & "-" & dato3.text)
End Sub

Private Sub dato27_GotFocus()
'Call VerificarCajas(Me, dato27)
'        Call selecciona(dato27)
End Sub

Private Sub dato27_KeyDown(KeyCode As Integer, Shift As Integer)
  Call Flechas(KeyCode, dato2)
End Sub

Private Sub dato27_KeyPress(KeyAscii As Integer)
 KeyAscii = esNumero(KeyAscii)
If KeyAscii = 13 Then
dato27.text = ceros(dato27)
dato3.SetFocus
End If

End Sub

    Private Sub dato3_GotFocus()
        Call VerificarCajas(Me, dato3)
        Call selecciona(dato3)
    End Sub
Private Sub dato30_GotFocus()
'       Call VerificarCajas(Me, dato30)
'        Call selecciona(dato30)
        Principal.barraEstado.Panels(2).text = "F2: Ayuda Caja"
End Sub

    Private Sub dato4_GotFocus()
        Call VerificarCajas(Me, dato4)
        Call selecciona(dato4)
    End Sub
    
    Private Sub dato5_GotFocus()
        
        Call VerificarCajas(Me, dato5)
        Call selecciona(dato5)
    End Sub
 
Private Sub dato50_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF2 Then
            Call ayudacajera(dato50)
  End If
End Sub

Private Sub dato50_KeyPress(KeyAscii As Integer)
KeyAscii = esNumero(KeyAscii)
        If KeyAscii = 13 And dato50.text <> "" Then
            dato50.text = ceros(dato50)
            lbldvcajera.Caption = rut(dato50.text)
            lblcajera.Caption = leerNombreCajera(dato50.text & lbldvcajera.Caption)
            If lblcajera.Caption <> "" Then
'               dato9.SetFocus
            End If
        End If
End Sub

    Private Sub dato6_GotFocus()
'        If dato5.text + "-" + dato4.text + "-" + dato3.text <> fechasistema And dato1.text <> "NP" And dato1.text <> "CO" Then
'        MsgBox ("imposible digitar facturas con fecha anterior")
'
'        dato3.SetFocus
'
'        End If
        
        Call VerificarCajas(Me, dato6)
        Call selecciona(dato6)
        Principal.barraEstado.Panels(2).text = "F2: Ayuda Cliente"
    End Sub
    
    Private Sub dato7_GotFocus()
        Call VerificarCajas(Me, dato7)
        Call selecciona(dato7)
        Principal.barraEstado.Panels(2).text = "F2: Ayuda Sucursal"
    End Sub
    
    Private Sub dato9_GotFocus()
        Call VerificarCajas(Me, dato9)
        Call selecciona(dato9)
'        Principal.barraEstado.Panels(2).text = "F2: Ayuda Nota de Pedido"
    End Sub
    
    Private Sub dato10_GotFocus()
'        Call VerificarCajas(Me, dato10)
'        Call selecciona(dato10)
        Principal.barraEstado.Panels(2).text = "F2: Ayuda Vendedor"
    End Sub
    
    Private Sub dato11_GotFocus()
        Static cont As Integer
        Call VerificarCajas(Me, dato11)
        Call selecciona(dato11)
        cont = cont + 1
        If Descuento.Visible = False Then
            If cont = 1 Then
                Load Descuento
                Descuento.MONTO = CDbl(lblSub.Caption)
                Descuento.Show vbModal
                If dire = 38 Then
                    detalle.SetFocus
                End If
                If dire = 40 Then
                    dato12.SetFocus
                End If
            Else
                cont = 0
            End If
        Else
            cont = 0
        End If
    End Sub
    
    Private Sub dato12_GotFocus()
          Call VerificarCajas(Me, dato12)
          Call selecciona(dato12)
    End Sub
    
    Private Sub dato8_GotFocus()
        Call VerificarCajas(Me, dato8)
        Call selecciona(dato8)
        If dato1.text = "ZE" Then
            dato8.Locked = True
        Else
            dato8.Locked = False
        End If
    End Sub
    '========================================================
    'GotFocus
    '========================================================
    '========================================================
    'KeyDown
    '========================================================
    Private Sub dato1_KeyDown(KeyCode As Integer, Shift As Integer)
        Select Case KeyCode
            Case 97, 49
                dato1.text = "FV"
                tipoprecio = "01"
            Case 98, 50
                dato1.text = "BV"
                tipoprecio = "01"
            Case 99, 51
                dato1.text = "NF"
                tipoprecio = "01"
            Case 100, 52
                dato1.text = "NB"
                tipoprecio = "01"
'           Case 99, 53
'                dato1.text = "CO"
'                tipoprecio = "01"
            Case Else
                Call Flechas(KeyCode, dato1)
        End Select
    End Sub
       Private Sub dato30_KeyDown(KeyCode As Integer, Shift As Integer)
       If KeyCode = vbKeyF2 Then
            Call ayudaCaja(dato30, empresaActiva)
        Else
            Call Flechas(KeyCode, dato1)
        End If
   
       End Sub
    
    Private Sub dato2_KeyDown(KeyCode As Integer, Shift As Integer)
        Call Flechas(KeyCode, dato30)
    End Sub
 
    
    Private Sub dato3_KeyDown(KeyCode As Integer, Shift As Integer)
        Call Flechas(KeyCode, dato2)
    End Sub
    
    Private Sub dato4_KeyDown(KeyCode As Integer, Shift As Integer)
        Call Flechas(KeyCode, dato3)
    End Sub
    
    Private Sub dato5_KeyDown(KeyCode As Integer, Shift As Integer)
        Call Flechas(KeyCode, dato4)
    End Sub
    
    Private Sub dato6_KeyDown(KeyCode As Integer, Shift As Integer)
        If KeyCode = vbKeyF2 Then
            Call ayudaCliente(dato6, dato7, lbldv)
        Else
            Call Flechas(KeyCode, dato5)
        End If
    End Sub
    
    Private Sub dato7_KeyDown(KeyCode As Integer, Shift As Integer)
        Call Flechas(KeyCode, dato6)
    End Sub
    
    Private Sub dato9_KeyDown(KeyCode As Integer, Shift As Integer)
        Call Flechas(KeyCode, dato8)
    End Sub
    
    Private Sub dato10_KeyDown(KeyCode As Integer, Shift As Integer)
        If KeyCode = vbKeyF2 Then
            Call ayudaVendedores(dato10)
        Else
            Call Flechas(KeyCode, dato9)
        End If
    End Sub
    
    Private Sub dato11_KeyDown(KeyCode As Integer, Shift As Integer)
        Call Flechas(KeyCode, dato10)
    End Sub
    
    Private Sub dato12_KeyDown(KeyCode As Integer, Shift As Integer)
        Call Flechas(KeyCode, dato11)
    End Sub
    
    Private Sub dato8_KeyDown(KeyCode As Integer, Shift As Integer)
        Call Flechas(KeyCode, dato7)
    End Sub
    Private Sub dato20_KeyDown(KeyCode As Integer, Shift As Integer)
        Call Flechas(KeyCode, dato9)
    End Sub
    Private Sub dato21_KeyDown(KeyCode As Integer, Shift As Integer)
        Call Flechas(KeyCode, dato20)
    End Sub
    
    '========================================================
    'KeyDown
    '========================================================
    
    '========================================================
    'KeyPress
    '========================================================
    Private Sub dato1_KeyPress(KeyAscii As Integer)
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If KeyAscii = 13 Then
            lblDocumento.Caption = leerNombreDocumento(dato1.text)
            If lblDocumento.Caption <> "" Then
'                Select Case dato1.text
'
'                    Case Else
'
'                End Select
               dato30.Enabled = True
               dato30.SetFocus
'              SendKeys "{Tab}"
            End If
            tipo_doc = dato1.text
        End If
    End Sub
    Private Sub dato30_KeyPress(KeyAscii As Integer)
    KeyAscii = esNumero(KeyAscii)
    If KeyAscii = 13 And dato30.text <> "" Then
    dato30.text = ceros(dato30)
    lblnombrecaja.Caption = leerNombreCaja(dato30.text)
    If lblDocumento.Caption <> "" And lblnombrecaja.Caption <> "" Then
                Select Case dato30.text

                    Case Else
      
'           dato2.text = leerfoliocaja(dato1.text, dato30.text)
'
'           dato27.text = leerUltimofoliocaja(dato1.text, dato30.text)
'
       
       dato2.text = ceros(dato2)
                End Select
               dato3.SetFocus
            End If
       caja = dato30.text
    End If
  
End Sub
   
    Private Sub dato2_KeyPress(KeyAscii As Integer)
        KeyAscii = esNumero(KeyAscii)
        If KeyAscii = 13 Then
            dato2.text = ceros(dato2)
            If Val(dato2.text) = 0 Then
                Call dato1_KeyPress(13)
                Exit Sub
            End If
            
            If leerVenta(v, dato1.text, dato2.text, "=", data, detalle, dato30.text, dato5.text & "-" & dato4.text & "-" & dato3.text) = True Then
            
            
                lectura = True
                Call structtoctrl
                nula = leerDocumentoNulo(dato1.text, dato2.text)
                
                If nula = True Then
                    lblNulo.Caption = "DOCUMENTO ANULADO"
                Else
                    lblNulo.Caption = ""
                End If
                
                lbl30.Visible = True
                lbl20.Visible = False
                lbl26.Visible = False
                opciones.Visible = True
                opciones.SetFocus
                 
            Else
            
             If Verifica_Permiso(Me.Caption, "agrega") = True Then
               lectura = False
                detalle.SelectionMode = cellSelectionFree
                'If detalle.Rows <= 1 Then
                    detalle.Rows = 1
                    detalle.AddItem vbTab & vbTab & "1" & vbTab & "0" & vbTab & "0" & vbTab & "0" & vbTab & "0", True
                'End If
                Call HabilitarCajas(Me, modifica)
                lbl30.Visible = False
                lbl20.Visible = True
                lbl26.Visible = True
                dato27.SetFocus
            Else
                MsgBox mensaje_nopermiso, vbCritical + vbOKOnly, "Permiso Denegado"
                dato2.SelStart = 0
                dato2.SelLength = Len(dato2.text)
                dato2.SetFocus
            End If
            
            
                
            End If
            numero_doc = dato2.text
        End If
    End Sub
    
    Private Sub dato3_KeyPress(KeyAscii As Integer)
        KeyAscii = esNumero(KeyAscii)
        If KeyAscii = 13 Then
            dato3.text = ceros(dato3)
            If dato3.text = "00" Then
                dato3.text = Format(fechasistema, "dd")
                dato4.text = Format(fechasistema, "mm")
                dato5.text = Format(fechasistema, "yyyy")
                'dato4.Enabled = True
                'dato5.Enabled = True
                fecha = dato3.text & "-" & dato4.text & "-" & dato5.text
               
                    dato5.SetFocus
                
            End If
            SendKeys "{Tab}"
        End If
    End Sub
    
    Private Sub dato4_KeyPress(KeyAscii As Integer)
        KeyAscii = esNumero(KeyAscii)
        If KeyAscii = 13 Then
            dato4.text = ceros(dato4)
            If dato4.text = "00" Then
                dato4.text = Format(fechasistema, "mm")
            End If
            SendKeys "{Tab}"
        End If
    End Sub
    
    Private Sub dato5_KeyPress(KeyAscii As Integer)
        KeyAscii = esNumero(KeyAscii)
        If KeyAscii = 13 Then
            dato5.text = ceros(dato5)
            If dato5.text = "0000" Then
                dato5.text = Format(fechasistema, "yyyy")
            End If
            fecha = dato3.text & "-" & dato4.text & "-" & dato5.text
           
             SendKeys "{Tab}"
            
        End If
    End Sub
    
    Private Sub dato6_KeyPress(KeyAscii As Integer)
        KeyAscii = esNumero(KeyAscii)
        If KeyAscii = 13 And dato6.text <> "" And Val(dato6.text) <> 0 Then
            dato6.text = ceros(dato6)
            lbldv.Caption = rut(dato6.text)
            rut_cliente = dato6.text
            SendKeys "{Tab}"
        End If
    End Sub
    
    Private Sub dato7_KeyPress(KeyAscii As Integer)
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If KeyAscii = 13 Then
            dato7.text = ceros(dato7)
            If leerCliente(c, dato6.text & lbldv.Caption, dato7.text, "=") = True Then
                structtoctrlCliente
                rut_cliente = dato6.text & lbldv.Caption
                sucursal_cliente = dato7.text
                If lectura = False Then
                    If autorizado = False Then
                        If verificarCupoCliente(dato6.text & lbldv.Caption, dato7.text) = False Then
                            Call enviarInformacion(rut_cliente, sucursal_cliente, dato1.text, dato2.text, "0", "CUPO INSUFICINTE")
                            'Call mensaje.mostrarMensaje("Información Crédito Cliente", "El cliente " & dato6.text & "-" & lblDV.Caption & " no posee cupo suficiente para realizar la compra.", "Solicite autorizaión")
                        End If
                    End If
                End If
'                If empresaActiva = "01" Then dato10.text = c.vendedor
                
                SendKeys "{Tab}"
            Else
                If MsgBox("El rut ingresado no se encuentra. ¿Desea crearlo?", vbYesNo, "Mensaje") = vbYes Then
                    Load MClientes
                    MClientes.dato1.text = dato6.text
                    MClientes.lbldv.Caption = lbldv.Caption
                    MClientes.dato2.text = dato7.text
                    MClientes.Show
                End If
            End If
        End If
    End Sub
    
    Private Sub dato9_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If KeyAscii = 13 Then
          
            SendKeys "{Tab}"
        End If
    End Sub
    
    Private Sub dato20_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If KeyAscii = 13 Then
          
            SendKeys "{Tab}"
        End If
    End Sub
    Private Sub dato21_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
        
        If KeyAscii = 13 Then
                vacio = False
                
                detalle.Enabled = True
               
                    detalle.Cell(1, 1).SetFocus
        End If
    End Sub
    
    Private Sub dato10_KeyPress(KeyAscii As Integer)
        KeyAscii = esNumero(KeyAscii)
        If KeyAscii = 13 And dato10.text <> "" Then
            dato10.text = ceros(dato10)
            lblDVV.Caption = rut(dato10.text)
            lblvendedor.Caption = leerNombreVendedor(dato10.text & lblDVV.Caption)
            If lblvendedor.Caption <> "" Then
               dato50.SetFocus
            End If
        End If
    End Sub
    
    Private Sub dato11_KeyPress(KeyAscii As Integer)
        KeyAscii = esNumeroDecimal(dato11, KeyAscii)
        If KeyAscii = 13 And dato11.text <> "" Then
            SendKeys "{Tab}"
        End If
        If dato11.text = "" Then
            dato11.text = "0"
        End If
    End Sub
    
    Private Sub dato12_KeyPress(KeyAscii As Integer)
        KeyAscii = esNumero(KeyAscii)
        
        If KeyAscii = 13 And dato12.text <> "" Then
'            Call ctrltostruct(True)
'            lbl26.Visible = False
'            lbl20.Visible = False
            
            If dato1.text <> "GD" And dato1.text <> "GM" Then
            compara = CDbl(dato12.text)
            compara2 = CDbl(lblSub.Caption)
            
               If compara > compara2 Then
                  If MsgBox("El descuento es mayor al total", vbOKOnly) = vbOK Then
                     Descuento.Show
                  End If
               Else
            
             '   detallePagos.Show vbModal
               End If
                
            Else
                Rem If MsgBox("DESEA IMPRIMIR COMPROBANTE ", vbYesNo) = vbYes Then
                    Call imprimir
                    
                    
                Rem End If
                imprimio = True
            End If
            
            If imprimio = True Then
                Call retorno
            Else
                modifica = True
            End If
              dato13.SetFocus
        End If
        If dato12.text = "" Then
            dato11.text = "0"
            dato12.text = "0"
        End If
        
        
        
        
    End Sub
    
    Private Sub dato8_KeyPress(KeyAscii As Integer)
        Dim fechven As String
       KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If KeyAscii = 13 Then
'            If dato8.text <> "" Then
'                If Val(dato8.text) > 90 Then
'                    dato8.text = "90"
'                End If
'                fechven = DateAdd("d", dato8.text, fecha)
'            Else
'                dato8.text = "0"
'                fechven = fecha
'            End If
'            lblDia.Caption = Format(fechven, "dd")
'            lblMes.Caption = Format(fechven, "mm")
'            lblAño.Caption = Format(fechven, "yyyy")
          dato10.SetFocus
        End If
    End Sub
    '========================================================
    'KeyPress
    '========================================================
    
    '========================================================
    'KeyUp
    '========================================================
    Private Sub dato2_KeyUp(KeyCode As Integer, Shift As Integer)
        Call seleccionaUno(KeyCode, dato2)
    End Sub
    
'    Private Sub dato3_KeyUp(KeyCode As Integer, Shift As Integer)
'        If Len(dato3.text) = dato3.MaxLength Then
'            Call dato3_KeyPress(13)
'        End If
'    End Sub
'
'    Private Sub dato4_KeyUp(KeyCode As Integer, Shift As Integer)
'        If Len(dato4.text) = dato4.MaxLength Then
'            Call dato4_KeyPress(13)
'        End If
'    End Sub
'
'    Private Sub dato5_KeyUp(KeyCode As Integer, Shift As Integer)
'        If Len(dato5.text) = dato5.MaxLength Then
'            Call dato5_KeyPress(13)
'        End If
'    End Sub
    
    Private Sub dato11_KeyUp(KeyCode As Integer, Shift As Integer)
        Dim desc As Double
        Dim total As Double
        Dim cadena As String
        Dim deci As String
        Dim i As Long
'        'total = Round(CDbl(lblNeto.Caption) + CDbl(lblIVA.Caption) + CDbl(lblIHA.Caption), 0)
        total = CDbl(lblSub.Caption)
        If dato11.text <> "" Then
            desc = total * CDbl(dato11.text) / 100
            cadena = Format(desc, "########0.0")
            deci = Right(cadena, 1)
            If deci >= 5 Then
                deci = 1
            Else
                deci = 0
            End If
            dato12.text = Val(desc) + CDbl(deci)
'            lblNeto.Caption = Round(total - CDbl(dato12.text), 0)
            For i = 1 To detalle.Rows - 1
                detalle.Cell(i, 8).text = dato11.text
                Rem detalle.Cell(i, 7).text = CDbl(detalle.Cell(i, 3).text) * CDbl(detalle.Cell(i, 5).text) - (CDbl(detalle.Cell(i, 5).text) * CDbl(detalle.Cell(i, 6).text) / 100) * CDbl(detalle.Cell(i, 3).text)
            Next i
        Else
            dato11.text = "0"
            dato12.text = "0"
'            lblTotal.Caption = Round(total - CDbl(dato12.text), 0)
        End If
        Call sumaGrilla(detalle)
    End Sub
    
    
'
'      Private Sub dato11_KeyUp(KeyCode As Integer, Shift As Integer)
'        Dim desc As Double
'        Dim total As Double
'        Dim cadena As String
'        Dim deci As String
'        Dim i As Long
''        'total = Round(CDbl(lblNeto.Caption) + CDbl(lblIVA.Caption) + CDbl(lblIHA.Caption), 0)
'        total = CDbl(lblSub.Caption)
'        If dato11.text <> "" Then
'            desc = total * CDbl(dato11.text) / 100
'            cadena = Format(desc, "########0.0")
'            deci = Right(cadena, 1)
'            If deci >= 5 Then
'                deci = 1
'            Else
'                deci = 0
'            End If
'            dato12.text = Val(desc) + CDbl(deci)
''            lblNeto.Caption = Round(total - CDbl(dato12.text), 0)
'            For i = 1 To detalle.Rows - 1
'                detalle.Cell(i, 6).text = dato11.text
'                detalle.Cell(i, 7).text = CDbl(detalle.Cell(i, 3).text) * CDbl(detalle.Cell(i, 5).text) - (CDbl(detalle.Cell(i, 5).text) * CDbl(detalle.Cell(i, 6).text) / 100) * CDbl(detalle.Cell(i, 3).text)
'            Next i
'        Else
'            dato11.text = "0"
'            dato12.text = "0"
''            lblTotal.Caption = Round(total - CDbl(dato12.text), 0)
'        End If
'        Call sumaGrilla(detalle)
'    End Sub

    Private Sub dato12_KeyUp(KeyCode As Integer, Shift As Integer)
        Dim desc As Double
'        Dim neto As Double
        Dim total As Double
        Dim i As Long
'        If lblNeto.Caption <> "" And lblIVA.Caption <> "" And lblIHA.Caption <> "" Then
'            total = Round(CDbl(lblNeto.Caption) + CDbl(lblIVA.Caption) + CDbl(lblIHA.Caption), 0)
'        End If
        total = CDbl(lblSub.Caption)
        If dato12.text <> "" And dato12.text <> "0" Then
            desc = CDbl(dato12.text)
            desc = desc * 100 / total
            dato11.text = Round(desc, 2)
'            lblTotal.Caption = Round(total - CDbl(dato12.text), 0)
            For i = 1 To detalle.Rows - 1
                detalle.Cell(i, 8).text = dato11.text
                Rem detalle.Cell(i, 7).text = CDbl(detalle.Cell(i, 3).text) * CDbl(detalle.Cell(i, 5).text) - (CDbl(detalle.Cell(i, 5).text) * CDbl(detalle.Cell(i, 6).text) / 100) * CDbl(detalle.Cell(i, 3).text)
            Next i
        Else
            dato11.text = "0"
            dato12.text = "0"
'            lblTotal.Caption = Round(total - CDbl(dato12.text), 0)
        End If
        Call sumaGrilla(detalle)
    End Sub
    '========================================================
    'KeyUp
    '========================================================

    '========================================================
    'LostFocus
    '========================================================
    Private Sub dato1_LostFocus()
        frmTipo.Visible = False
        lbl20.Visible = True
        lbl26.Visible = True
        
        
    If dato1.text = "BV" Or dato1.text = "ZE" And dato6.text = "" Then
    CARGADATOSBOLETA
    
    End If
    
    End Sub
    
    Private Sub dato3_LostFocus()
    Call esfecha(dato3, dato4, dato5, "dd")
    End Sub
    Private Sub dato4_LostFocus()
    Call esfecha(dato3, dato4, dato5, "mm")
    End Sub
    Private Sub dato5_LostFocus()
    Call esfecha(dato3, dato4, dato5, "yyyy")
    End Sub

    Private Sub dato6_LostFocus()
        Call limpiaBarra(2)
    End Sub
    Private Sub dato30_LostFocus()
    Call limpiaBarra(2)
    End Sub
    
    Private Sub dato7_LostFocus()
        Call limpiaBarra(2)
    End Sub
    
    Private Sub dato9_LostFocus()
'        Call limpiaBarra(2)
    End Sub
    
    Private Sub dato10_LostFocus()
        Call limpiaBarra(2)
        vend = dato10.text
    End Sub
    '========================================================
    'LostFocus
    '========================================================
'============================================================
'MANEJO DE LOS CONTOLES
'============================================================

Private Sub detalle_KeyPress(KeyAscii As Integer)
    
    If detalle.ActiveCell.col = 1 And KeyAscii > 65 And detalle.ActiveCell.text = "" Then
    detalle.ActiveCell.text = leeletra(Chr(KeyAscii))
    End If
    
End Sub

    Private Sub Form_Activate()
        If segurity = True Then
            seguridad.Show vbModal
            segurity = False
        End If
        Principal.barraEstado.Panels(1).text = UCase(Me.Caption)
    End Sub
    Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
        Dim tabla As String
        Dim TIPO As String
        Dim csql As New rdoQuery
        Dim resultados As rdoResultset
        Set csql.ActiveConnection = ventasRubro
        If KeyCode = 27 And Screen.ActiveForm.ActiveControl.Name = "dato1" Then
            
            Unload Me
        End If
        If KeyCode = 80 Then
            If lectura = True Then
                Load detallePagos
                With detallePagos
                    .lectura = True
                    .pagos.Rows = 1
                    .pagos.AutoRedraw = False
                    
                    csql.sql = "SELECT tipopago, monto, CONCAT('" & vbTab & "', monto, '" & vbTab & "', numerodocumento, '" & vbTab & "', banco, '" & vbTab & "', cuentacorriente, '" & vbTab & "', IF(vencimiento <> '00-00-0000',CONCAT(DATE_FORMAT(vencimiento,'%d'), '" & vbTab & "', DATE_FORMAT(vencimiento,'%m'), '" & vbTab & "', DATE_FORMAT(vencimiento,'%Y')),'')) AS item "
                    csql.sql = csql.sql & "FROM sv_documento_pagos_" + empresaActiva + " "
                    csql.sql = csql.sql & "WHERE local = '" & empresaActiva & "' AND tipo = '" & dato1.text & "' AND numero = '" & dato27.text & "'"
                    csql.sql = csql.sql & "ORDER BY lineapago ASC"
                    csql.Execute
                    'Call ConectarControlData(.data, servidor, baseVentas & rubro, usuario, password, tabla)
                    If csql.RowsAffected > 0 Then
'                        .data.Recordset.MoveFirst
                          Set resultados = csql.OpenResultset
                        While Not resultados.EOF
                        pivote.MaxLength = 2
                        pivote.text = resultados("tipopago")
                        pivote.text = ceros(pivote)
                        
                            TIPO = leerTipopagos(pivote.text)
'                            Select Case tipo
'                                Case "1"
'                                    If Val(resultados("monto")) <= 0 Then
'                                        tipo = "7 - Vuelto"
'                                    Else
'                                        tipo = "1 - Efectivo"
'                                    End If
'                                Case "2"
'                                    tipo = "2 - Cheque Propio"
'                                Case "3"
'                                    tipo = "3 - Cheque Tercero"
'                                Case "4"
'                                    tipo = "4 - Crédito Directo"
'                                'Case "5"
'                                '    tipo = "5 - "
'                                'Case "6"
'                                '    tipo = "6 - "
'                            End Select
                            .pagos.AddItem TIPO & resultados("item"), True
                            resultados.MoveNext
                        Wend
                        Set csql = Nothing
                        csql.Close
                        Set resultados = Nothing
                        
                    End If
                    '.pagos.Range(1, 1, .pagos.Rows - 1, .pagos.Cols - 1).Locked = True
                    .pagos.AutoRedraw = True
                    .pagos.Refresh
                    .pagos.SelectionMode = cellSelectionByRow
                    detallePagos.Show vbModal
                    Call retorno
                     
                End With
            End If
        End If
        If KeyCode = 38 Then
            If Screen.ActiveForm.ActiveControl.Name = "dato1" Then
                Unload Me
            End If
        End If

          If KeyCode = 45 And dato1.text <> "" And dato2.text <> "" Then
        frmdatos.Enabled = True
        frmdatos.Enabled = False
        
        FRMPREVENTA.Visible = True
        PREVENTA.SetFocus
        
        End If
       

    End Sub
    
    Private Sub Form_Load()
        FRMPREVENTA.Visible = False
        
        titCaption = Me.Caption
        'segurity = Not Verificar(usuarioSistema, passwordSistema)
        modifica = False
        nula = False
        imprimio = False
        lectura = False
        autorizado = False
        Call Centrar(Me)
        Call CARGAGRILLA(1, 9)
        dato1.text = "FV"
'        iva = leerImpuesto("IVA")
    End Sub
    
'****************************************************************************
'Formato de la Grilla
'****************************************************************************
    Private Sub CARGAGRILLA(ByVal row As Integer, ByVal col As Integer)
        Dim i As Integer
        Rem DATOS DE LA COLUMNA
        formatogrilla(1, 0) = "LN"
        formatogrilla(1, 1) = "CODIGO"
        formatogrilla(1, 2) = "DESCRIPCION"
        formatogrilla(1, 3) = "CANTIDAD"
        formatogrilla(1, 4) = "     "
        formatogrilla(1, 5) = "PRECIO"
        formatogrilla(1, 6) = "DESC"
        formatogrilla(1, 7) = "TOTAL"
        formatogrilla(1, 8) = " "
        formatogrilla(1, 9) = ""
 
        
        Rem LARGO DE LOS DATOS
        formatogrilla(2, 1) = "15"
        formatogrilla(2, 2) = "40"
        formatogrilla(2, 3) = "10"
        formatogrilla(2, 4) = "9"
        formatogrilla(2, 5) = "9"
        formatogrilla(2, 6) = "2"
        formatogrilla(2, 7) = "9"
        formatogrilla(2, 8) = "0"
        formatogrilla(2, 9) = "0"
        
        
        Rem TIPO DE DATOS C=CEROS,S=STRING,N=NUMERICO,D=DATE,
        formatogrilla(3, 1) = "N"
        formatogrilla(3, 2) = "S"
        formatogrilla(3, 3) = "N"
        formatogrilla(3, 4) = "N"
        formatogrilla(3, 5) = "N"
        formatogrilla(3, 6) = "N"
        formatogrilla(3, 7) = "N"
        formatogrilla(3, 8) = "N"
        formatogrilla(3, 9) = "S"
       
        
        Rem FORMATO GRILLA
        formatogrilla(4, 1) = "0000000000000"
        formatogrilla(4, 2) = ""
        formatogrilla(4, 3) = "#,###,##0.00"
        formatogrilla(4, 4) = "###,###,##0"
        formatogrilla(4, 5) = "$ ###,###,##0.00"
        formatogrilla(4, 6) = "#0.00"
        formatogrilla(4, 7) = "$ ###,###,##0"
        formatogrilla(4, 8) = ""
        formatogrilla(4, 9) = ""
       
        Rem LOCCKED
        'FormatoGrilla(5, 0) = "TRUE"
        formatogrilla(5, 1) = "FALSE"
        formatogrilla(5, 2) = "TRUE"
        formatogrilla(5, 3) = "FALSE"
        formatogrilla(5, 4) = "TRUE"
        formatogrilla(5, 5) = "TRUE"
        formatogrilla(5, 6) = "FALSE"
        formatogrilla(5, 7) = "TRUE"
        formatogrilla(5, 8) = "TRUE"
        formatogrilla(5, 9) = "TRUE"
       
        Rem VALOR MINIMO
        formatogrilla(6, 1) = ""
        formatogrilla(6, 2) = ""
        formatogrilla(6, 3) = ""
        formatogrilla(6, 4) = ""
        formatogrilla(6, 5) = ""
        formatogrilla(6, 6) = ""
        formatogrilla(6, 7) = ""
        formatogrilla(6, 8) = ""
     
        
        Rem VALOR MAXIMO
        formatogrilla(7, 1) = ""
        formatogrilla(7, 2) = ""
        formatogrilla(7, 3) = ""
        formatogrilla(7, 4) = ""
        formatogrilla(7, 5) = ""
        formatogrilla(7, 6) = ""
        formatogrilla(7, 7) = ""
        formatogrilla(7, 8) = ""
    
        
        Rem ANCHO
        formatogrilla(8, 1) = "10"
        formatogrilla(8, 2) = "35"
        formatogrilla(8, 3) = "10"
        formatogrilla(8, 4) = "6"
        formatogrilla(8, 5) = "12"
        formatogrilla(8, 6) = "5"
        formatogrilla(8, 7) = "12"
        formatogrilla(8, 8) = "0"
        formatogrilla(8, 9) = "0"
       
            
        detalle.Cols = col + 1
        detalle.Rows = row
        detalle.AllowUserResizing = False
        detalle.DisplayFocusRect = False
        detalle.ExtendLastCol = False
        detalle.BoldFixedCell = False
        detalle.DisplayRowIndex = True
        detalle.DrawMode = cellOwnerDraw
        detalle.Appearance = Flat
        detalle.ScrollBarStyle = Flat
        detalle.FixedRowColStyle = Flat
        detalle.BackColorFixed = RGB(90, 158, 214)
        detalle.BackColorFixedSel = RGB(110, 180, 230)
        detalle.BackColorBkg = RGB(90, 158, 214)
        detalle.BackColorScrollBar = RGB(231, 235, 247)
        detalle.BackColor1 = RGB(231, 235, 247)
        detalle.BackColor2 = RGB(239, 243, 255)
        detalle.GridColor = RGB(148, 190, 231)
        'detalle.DefaultFont.Size = 8
        detalle.Column(8).Width = 0
        
        
        detalle.Cell(0, 0).text = formatogrilla(1, 0)
        For i = 1 To col - 1
            detalle.Cell(0, i).text = formatogrilla(1, i)
            detalle.Column(i).Width = Val(formatogrilla(8, i)) * (detalle.Cell(0, i).Font.Size + 1.25)
            detalle.Column(i).MaxLength = Val(formatogrilla(2, i))
            detalle.Column(i).FormatString = formatogrilla(4, i)
            detalle.Column(i).Locked = formatogrilla(5, i)
            If formatogrilla(3, i) = "N" Then
                detalle.Column(i).Alignment = cellRightCenter
                If i <> 5 And i <> 3 Then
                    detalle.Column(i).Mask = cellNumeric
                End If
            Else
                detalle.Column(i).Alignment = cellLeftCenter
                detalle.Column(i).Mask = cellUpper
            End If
        Next i
        detalle.Range(0, 0, 0, detalle.Cols - 1).Alignment = cellCenterCenter
        detalle.AddItem vbTab & vbTab & "1" & vbTab & "0" & vbTab & "0" & vbTab & "0" & vbTab & "0" & vbTab & "0" & vbTab & "", True
        detalle.Enabled = True
        
    
    End Sub
    
    
    Private Sub CARGAGRILLAboleta()
        Dim i As Integer
        Rem DATOS DE LA COLUMNA
  
        formatogrilla(1, 1) = ""
        formatogrilla(1, 2) = ""
        formatogrilla(1, 3) = ""
        
        Rem LARGO DE LOS DATOS
        formatogrilla(2, 1) = "4"
        formatogrilla(2, 2) = "20"
        formatogrilla(2, 3) = "9"
        Rem TIPO DE DATOS C=CEROS,S=STRING,N=NUMERICO,D=DATE,
        formatogrilla(3, 1) = "S"
        formatogrilla(3, 2) = "S"
        formatogrilla(3, 3) = "N"
        Rem FORMATO GRILLA
        formatogrilla(4, 1) = ""
        formatogrilla(4, 2) = ""
        formatogrilla(4, 3) = ""
        Rem LOCCKED
        'FormatoGrilla(5, 0) = "TRUE"
        formatogrilla(5, 1) = "FALSE"
        formatogrilla(5, 2) = "FALSE"
        formatogrilla(5, 3) = "FALSE"
        Rem ANCHO
        formatogrilla(8, 1) = "4"
        formatogrilla(8, 2) = "15"
        formatogrilla(8, 3) = "6"
        
        impresionboleta.Cols = 4
        impresionboleta.Rows = 1
        impresionboleta.AllowUserResizing = False
        impresionboleta.DisplayFocusRect = False
        impresionboleta.ExtendLastCol = False
        impresionboleta.BoldFixedCell = False
        impresionboleta.DisplayRowIndex = True
        impresionboleta.DrawMode = cellOwnerDraw
        impresionboleta.Appearance = Flat
        impresionboleta.ScrollBarStyle = Flat
        impresionboleta.FixedRowColStyle = Flat
 
        'detalle.DefaultFont.Size = 8
'        impresionboleta.Column(8).Width = 0
        
        
        impresionboleta.Cell(0, 0).text = formatogrilla(1, 0)
        For i = 1 To impresionboleta.Cols - 1
            impresionboleta.Cell(0, i).text = formatogrilla(1, i)
            impresionboleta.Column(i).Width = Val(formatogrilla(8, i)) * (detalle.Cell(0, i).Font.Size + 1.25)
            impresionboleta.Column(i).MaxLength = Val(formatogrilla(2, i))
            impresionboleta.Column(i).FormatString = formatogrilla(4, i)
            impresionboleta.Column(i).Locked = formatogrilla(5, i)
            If formatogrilla(3, i) = "N" Then
                detalle.Column(i).Alignment = cellRightCenter
            Else
                impresionboleta.Column(i).Alignment = cellLeftCenter
                impresionboleta.Column(i).Mask = cellUpper
            End If
        Next i
        impresionboleta.Range(0, 0, 0, impresionboleta.Cols - 1).Alignment = cellCenterCenter
        impresionboleta.Enabled = True
        
    
    End Sub
    
    
    Private Sub detalle_GotFocus()
        
        Call VerificarCajas(Me, detalle)
        If detalle.ActiveCell.col = 1 Then
            Principal.barraEstado.Panels(2).text = "F2 Ayuda - Producto"
        Else
            Principal.barraEstado.Panels(2).text = ""
        End If
    End Sub
    
    Private Sub detalle_KeyDown(KeyCode As Integer, ByVal Shift As Integer)
        fila = detalle.ActiveCell.row
        columna = detalle.ActiveCell.col
        If detalle.ActiveCell.col = 1 And KeyCode = vbKeyF2 Then Call ayudaProducto2(detalle, pivote): detalle.Cell(fila, columna).SetFocus
        Select Case KeyCode
            Case 13, 37, 38, 39, 40
                If detalle.ActiveCell.text <> "" Then
                    vacio = False
                Else
                    vacio = True
                End If
                If KeyCode = 38 And detalle.ActiveCell.row = 1 And detalle.ActiveCell.col = 1 Then
                    dato10.SetFocus
                End If
                If KeyCode = 40 And detalle.ActiveCell.row = detalle.Rows - 1 And detalle.ActiveCell.col = 1 Then
                    dato13.SetFocus
                End If
            Case 106, 35
                SendKeys "{Tab}"
                'dato13.SetFocus
            Case 109
                If fila > 1 Or detalle.Rows > 2 Then
                    detalle.RemoveItem (fila)
                    Call sumaGrilla(detalle)
                End If
        End Select
        
    End Sub
    
    Private Sub detalle_KeyUp(KeyCode As Integer, Shift As Integer)
        If detalle.ActiveCell.col = 3 Or detalle.ActiveCell.col = 5 Then
            pivote.text = detalle.ActiveCell.text
            If pivote.text <> "" Then
                'KeyCode = esNumeroDecimal(pivote, Asc(Right(pivote.text, 1)))
                KeyCode = esNumeroDecimal(pivote, Asc(pivote.text))
            End If
            If KeyCode = 0 Then
                pivote.text = Left(pivote.text, Len(pivote.text) - 1)
            End If
            If KeyCode = 44 Then
                pivote.text = Left(pivote.text, Len(pivote.text) - 1) & ","
            End If
            detalle.ActiveCell.text = pivote.text
        End If
    End Sub
    
    Private Sub detalle_Click()
        Dim i As Integer
        For i = 1 To detalle.ActiveCell.col
            If detalle.Cell(detalle.ActiveCell.row, i).text = "" Then
                detalle.Cell(detalle.ActiveCell.row, i).SetFocus
                Exit For
            End If
        Next i
    End Sub
    
    Private Sub detalle_LeaveCell(ByVal row As Long, ByVal col As Long, NewRow As Long, NewCol As Long, Cancel As Boolean)
        Dim i As Integer
        Dim linea As String
        Dim limite As Integer
        Dim descu As Double
        Dim precio As Double
        Dim descu2 As Double
        
        If detalle.Rows <= NewRow Then
            NewRow = row
        End If
       
       If vacio = True Then
            If NewRow <> row And row <> detalle.Rows - 1 Then
                NewRow = fila
                NewCol = columna
            Else
                If NewCol > col Then
                    NewRow = fila
                    NewCol = columna
                End If
            End If
        Else
'         If (NewCol <> col Or NewRow <> row) And col = 5 And detalle.Cell(col, row).text = "0" Then
'         NewCol = col
'         NewRow = row
'
'
'         End If
         
            If col = 6 And NewCol = 7 Then
                If detalle.ActiveCell.text <> "" Then
                    If row = detalle.Rows - 1 Then
                        Select Case dato1.text
                            Case "FV", "FE", "GD", "NP", "CO", "ZE"
                                limite = 35
                                
                            Case "BV"
                                limite = 100
                            
                            Case Else
                                limite = 0
                        End Select
                        If limite > 0 Then
                            If limite > detalle.Rows - 1 Then
                                detalle.AddItem vbTab & vbTab & "1" & vbTab & "0" & vbTab & "0" & vbTab & "0" & vbTab & "0" & vbTab & "", True
                                
                                NewRow = detalle.Rows - 1
                                NewCol = 1
                            Else
                                dato11.SetFocus
                            End If
                        Else
                            detalle.AddItem vbTab & vbTab & "1" & vbTab & "0" & vbTab & "0" & vbTab & "0" & vbTab & "0" & vbTab & "0" & vbTab & "", True
                            
                            NewRow = detalle.Rows - 1
                            NewCol = 1
                        End If
                    Else
                        NewCol = 1
                    End If
                Else
                    NewCol = col
                End If
            Else
                If col = 5 And NewCol = 4 Then
                    NewCol = 3
                End If
            End If
            If col = 1 And NewCol = detalle.Cols - 1 Then
                For i = 1 To detalle.Cols - 1
                    If detalle.Cell(NewRow, i).text = "" Then
                        NewCol = i
                        Exit For
                    End If
                Next i
            End If
            
            If col = 1 And NewCol < detalle.Cols - 2 And NewCol > 1 Then
                pivote.text = detalle.Cell(row, 1).text
                If pivote.text <> "0000000000000" Then
                pivote.text = ceros(pivote)
                If Val(pivote.text) = "0" Then
                    pivote.text = ""
                End If
                End If
                
                detalle.Cell(row, 1).text = pivote.text
                detalle.Cell(row, 1).text = leeralias(detalle.Cell(detalle.ActiveCell.row, 1).text)
                If leerCodigoProducto(detalle.Cell(row, 1).text) = False Then
                    detalle.Cell(row, 1).text = ""
                    detalle.Cell(row, 1).SetFocus
                    
                End If

'                If leerstock(detalle.Cell(Row, 1).text) = False Then
''                    detalle.Cell(Row, 1).text = ""
''                      If MsgBox("PRODUCTO SIN STOCK", vbOKOnly, "ATENCION") = vbOK Then
'''                        glosas.Visible = True
'''                         dato24.SetFocus
''                      End If
'               End If
                           Rem detalle.Cell(row, detalle.Cols - 1).text = leerCostoProducto(detalle.Cell(row, 1).text)
                           detalle.Cell(row, 2).text = leerNombreProducto(detalle.Cell(row, 1).text)
                           If detalle.Cell(row, 2).text <> "" Then
                              NewCol = 3
                           End If
                           If row > 0 Then
                              detalle.Cell(row, 4).text = "1"
                           End If
                           Rem If Val(detalle.Cell(row, 5).text) = 0 Then
                           detalle.Cell(row, 5).text = leerPrecioEspecial(detalle.Cell(row, 1).text)
                           Rem If Val(detalle.Cell(row, 5).text) = 0 Then
                           detalle.Cell(row, 5).text = leerPrecioProducto(detalle.Cell(row, 1).text, tipoprecio)
                           If leerPrecioProducto(detalle.Cell(row, 1).text, tipoprecio) = "0" Then detalle.Cell(row, 5).text = ""
                           Rem End If
                           Rem End If
                           If detalle.Cell(row, 3).text <> "" And detalle.Cell(row, 5).text <> "" And row > 0 Then
                    
                           Else
                               detalle.Cell(row, 7).text = "0"
                           End If
               Else
                    If col = 1 And NewCol = detalle.Cols - 2 Then
                       NewCol = 5
                    End If
            End If
            If col = 3 And NewCol <> col Then
                If detalle.Cell(row, 3).text <> "" And CDbl(detalle.Cell(row, 3).text) > 0 Then
                    If NewCol > col Then
                        NewCol = 5
                    End If
                    If NewCol < col Then
                        NewCol = 1
                    End If
                Else
                    NewCol = col
                End If
            End If
            If NewRow > row Then
                For i = 1 To detalle.Cols - 3
                    If detalle.Cell(row, i).text = "" Then
                        NewRow = row
                        NewCol = i
                        Exit For
                    End If
                Next i
                For i = 1 To detalle.Cols - 3
                    If detalle.Cell(NewRow, i).text = "" Then
                        NewCol = i
                        Exit For
                    End If
                Next i
            End If
            If row > 0 Then
                If detalle.Cell(row, 3).text <> "" And detalle.Cell(row, 5).text <> "" Then
                    If dato1.text <> "WW" Then
                        detalle.Cell(row, 7).text = Round(detalle.Cell(row, 3).text * detalle.Cell(row, 5).text + 0.1, 0)
                    Else
                        If Val(detalle.Cell(row, 7).text) <> Val(detalle.Cell(row, 3).text) * Val(detalle.Cell(row, 5).text) Then
                            detalle.Cell(row, 7).text = detalle.Cell(row, 5).text
                            detalle.Cell(row, 5).text = detalle.Cell(row, 7).text / detalle.Cell(row, 3).text
                        End If
                    End If
                    detalle.Cell(row, 4).text = "1"
                End If
            End If
'            If row > 0 Then
'                precio = Round(detalle.Cell(row, 3).text * detalle.Cell(row, 5).text + 0.1, 0)
'                descu = Int((detalle.Cell(row, 7).text * ((detalle.Cell(row, 6).text) / 100)) + 0.5)
'                detalle.Cell(row, 7).text = Str(precio - descu)
'            End If
            If NewCol = 5 Then
            filaprecio = row
            
            ingresaprecio.Show vbModal
            
            End If
            
            If col > 0 And row > 0 Then
                Call sumaGrilla(detalle)
            End If
        End If
'        End If
                    

    End Sub
    Private Sub sumaGrilla(ByRef lista As Grid)
        Dim cad As String
        Dim i As Integer
        Dim suma As Double
        Dim sumaIVA As Double
        Dim sumaIHA As Double
        Dim sumaEXENTO As Double
        Dim CODIGO As String
        Dim deci As String
        Dim precio As Double
        Dim descu As Double
        Dim descu2 As Double
        
        suma = 0
        sumaIVA = 0
        sumaIHA = 0
        sumaEXENTO = 0
        descu2 = 0
        For i = 1 To lista.Rows - 1
            CODIGO = lista.Cell(i, 1).text
            If CODIGO <> "" Then
                If Val(lista.Cell(i, 3).text) = 0 Then
                    lista.Cell(i, 3).text = lista.Cell(i, 3).text
                End If
                If Val(lista.Cell(i, 5).text) = 0 Then
                    lista.Cell(i, 5).text = "0"
                End If
                precio = Int((lista.Cell(i, 3).text * lista.Cell(i, 5).text) + 0.5)
                If lista.Cell(i, 6).text <> "" Then descu = Int((precio * ((lista.Cell(i, 6).text) / 100)) + 0.5)
                lista.Cell(i, 7).text = Str(precio - descu)
                precio = lista.Cell(i, 7).text
                descu2 = descu2 + Int(((precio * dato11.text / 100)) + 0.5)
                sumaIVA = sumaIVA + CDbl(lista.Cell(i, 7).text)
                
                                
                suma = suma + CDbl(lista.Cell(i, 7).text)
            End If
        Next i
        
        sumaIVA = sumaIVA - descu2
        cad = Format(suma, "########0.0")
        deci = Right(cad, 1)
        If deci >= 5 Then
            deci = 1
        Else
            deci = 0
        End If
        suma = Val(cad) + CDbl(deci)
        
        
        lblSub.Caption = suma
        dato26.text = suma
        Select Case dato1.text
            Case "BV", "ZE"
                lblneto.Caption = Round((CDbl(lblSub.Caption) - CDbl(dato12.text)) / (1 + iva / 100), 0)
                lbliva.Caption = Round(CDbl(lblSub.Caption) - CDbl(dato12.text) - CDbl(lblneto.Caption), 0)
                lblIHA.Caption = "0"
            Case "FE"
                lblneto.Caption = Round(CDbl(lblSub.Caption) - CDbl(dato12.text), 0)
                lbliva.Caption = "0"
                lblIHA.Caption = "0"
            Case "GD"
                lblneto.Caption = Round(CDbl(lblSub.Caption) - CDbl(dato12.text), 0)
                lbliva.Caption = Round(sumaIVA * iva / 100, 0)
                lblIHA.Caption = "0"
            Case Else
                lblneto.Caption = Round(CDbl(lblSub.Caption) - CDbl(dato12.text), 0)
                lbliva.Caption = Round(sumaIVA * iva / 100, 0)
                lblIHA.Caption = Round(sumaIHA * iha / 100, 0)
                dato18.text = Round(CDbl(dato26.text) - CDbl(dato12.text), 0)
                dato19.text = Round(sumaIVA * iva / 100, 0)
        End Select
        dato26.text = CDbl(dato18.text) + CDbl(dato19.text) + 0
        lbltotal.Caption = CDbl(lblneto.Caption) + CDbl(lbliva.Caption) + CDbl(lblIHA.Caption)
    End Sub

'****************************************************************************
'Formato de la Grilla
'****************************************************************************

'=============================================================================
'PASA LOS DATOS DE LOS CONTROLES A LA ESTRUCTURA
'=============================================================================
    Private Sub ctrltostruct(ByVal graba As Boolean)
        v.cabeza.loc = empresaActiva
        v.cabeza.TIPO = dato1.text
'        dato2.text = leerUltimoFolio(v.cabeza.TIPO)
'        v.cabeza.numero = ceros(dato2)
        v.cabeza.NUMERO = ceros(dato27)
        v.cabeza.foliosii = ceros(dato2)
        v.cabeza.fecha = dato5.text & "-" & dato4.text & "-" & dato3.text
        v.cabeza.condicionesdepago = dato8.text
        v.cabeza.vencimiento = dato5.text & "-" & dato4.text & "-" & dato3.text
        v.cabeza.rut = dato6.text & lbldv.Caption
        v.cabeza.sucursal = dato7.text
        v.cabeza.vendedor = dato10.text & lblDVV.Caption
        v.cabeza.transporte = dato9.text
        v.cabeza.revisado = dato20.text
        v.cabeza.bultos = dato21.text
        v.cabeza.cajera = dato50.text & lbldvcajera.Caption
    
        v.cabeza.notaventas = ""
        v.cabeza.ordencompra = ""
        
        v.cabeza.Descuento = Replace(dato12.text, ".", ",")
        v.cabeza.subtotal = Format(dato26.text, "########0")
        v.cabeza.neto = Format(dato18.text, "########0")
        v.cabeza.iva = Format(dato19.text, "########0")
        v.cabeza.total = Format(dato26.text, "########0")
     
        v.cabeza.impuestoIla13 = Format(dato13.text, "########0")
        v.cabeza.impuestoIla15 = Format(dato14.text, "########0")
        v.cabeza.impuestoIla27 = Format(dato15.text, "########0")
        v.cabeza.impuestoCarne = Format(dato16.text, "########0")
        v.cabeza.impuestoHarina = Format(dato17.text, "########0")
        v.cabeza.exento = Format(dato25.text, "########0")
        v.cabeza.impuestoila = ""
        
        v.cabeza.impuestoespecifico = ""
       
        v.cabeza.retencionparcial = ""
        v.cabeza.retenciontotal = ""
        
        v.cabeza.abono = ""
       
        v.cabeza.contabilizado = ""
        v.cabeza.PAGADO = ""
        v.cabeza.comision = ""
        v.cabeza.fechapagocomision = ""
        v.cabeza.nula = "N"
        v.cabeza.boletadesde = DESDE
        v.cabeza.boletahasta = HASTA
        v.cabeza.abono2 = Format(dato22.text, "########0")
        v.cabeza.caja = dato30.text
'        v.cabeza.subtotal = Format(lblSub.Caption, "########0")
'        v.cabeza.neto = Format(lblNeto.Caption, "########0")
'        v.cabeza.iva = Format(lblIVA.Caption, "########0")
   
        
        v.detalle.caja = dato30.text
        v.detalle.loc = empresaActiva
        v.detalle.TIPO = dato1.text
        v.detalle.NUMERO = v.cabeza.NUMERO
        v.detalle.linea = ""
        v.detalle.fecha = dato5.text & "-" & dato4.text & "-" & dato3.text
        v.detalle.rut = dato6.text & lbldv.Caption
        v.detalle.sucursal = dato7.text
        v.detalle.CODIGO = ""
        v.detalle.descripcion = ""
        v.detalle.cantidad = ""
        v.detalle.unidades = ""
        v.detalle.precio = ""
        v.detalle.Descuento = ""
        v.detalle.total = ""
        v.detalle.vendedor = dato10.text & lblDVV.Caption
        v.detalle.pcosto = ""
        v.detalle.bodega = bodega
        v.detalle.vencimiento = dato5.text & "-" & dato4.text & "-" & dato3.text
        
        If graba = True Then
            Call grabarVenta(v, modifica, detalle)
        End If
        'Call retorno
    End Sub
'=============================================================================
'PASA LOS DATOS DE LOS CONTROLES A LA ESTRUCTURA
'=============================================================================

'=============================================================================
'PASA LOS DATOS DE LA ESTRUCTURA A LOS CONTROLES
'=============================================================================
    Private Sub structtoctrlCliente()
        Dim fechven As Date
        lblRazon.Caption = c.nombre
        LBLDIRECCION.Caption = c.direccion
        lblComuna.Caption = c.comuna
        LBLCIUDAD.Caption = c.ciudad
'        dato8.text = c.plazo
'        If dato8.text <> "" Then
'            fechven = DateAdd("d", dato8.text, fecha)
'
'            fechven = fecha
'        End If
'        lblDia.Caption = Format(fechven, "dd")
'        lblMes.Caption = Format(fechven, "mm")
'        lblAño.Caption = Format(fechven, "yyyy")
    End Sub
    
    Private Sub structtoctrl()
        Dim subtotal As Double
        Dim Descuento As Double
        
       ' dato1.text = v.cabeza.TIPO
      '  dato2.text = v.cabeza.NUMERO
        fecha = Format(v.cabeza.fecha, "dd-mm-yyyy")
        dato3.text = Format(v.cabeza.fecha, "dd")
        dato4.text = Format(v.cabeza.fecha, "mm")
        dato5.text = Format(v.cabeza.fecha, "yyyy")
        dato8.text = v.cabeza.condicionesdepago
'        lblDia.Caption = Format(v.cabeza.vencimiento, "dd")
'        lblMes.Caption = Format(v.cabeza.vencimiento, "mm")
'        lblAño.Caption = Format(v.cabeza.vencimiento, "yyyy")
        dato6.text = v.cabeza.rut
        dato7.text = v.cabeza.sucursal
        dato10.text = v.cabeza.vendedor
        dato9.text = v.cabeza.transporte
        dato20.text = v.cabeza.revisado
        dato21.text = v.cabeza.bultos
        dato27.text = v.cabeza.NUMERO
        dato50.text = v.cabeza.cajera
        dato13.text = Format(v.cabeza.impuestoIla13, "##########0")
        dato14.text = Format(v.cabeza.impuestoIla15, "##########0")
        dato15.text = Format(v.cabeza.impuestoIla27, "##########0")
        dato16.text = Format(v.cabeza.impuestoCarne, "##########0")
        dato17.text = Format(v.cabeza.impuestoHarina, "##########0")
'        lblSub.Caption = v.cabeza.subtotal
        dato18.text = v.cabeza.neto
        dato19.text = v.cabeza.iva
        dato25.text = v.cabeza.exento
        dato26.text = v.cabeza.total
'        lblSub.Caption = v.cabeza.subtotal
'        lblNeto.Caption = v.cabeza.neto
'        lblIVA.Caption = v.cabeza.iva
'        lblIHA.Caption = v.cabeza.exento
'        lblTotal.Caption = v.cabeza.total
        If v.cabeza.Descuento <> 0 Then
        Descuento = CDbl(v.cabeza.Descuento) / CDbl(v.cabeza.subtotal) * 100
        dato11.text = Descuento
        End If
        dato12.text = Replace(v.cabeza.Descuento, ",", ".")
'        dato22.text = v.cabeza.abono2
         
        Call dato6_KeyPress(13)
        Call dato7_KeyPress(13)
        Call dato10_KeyPress(13)
        Call dato50_KeyPress(13)
        Call DeshabilitarCajas(Me)
       'Detalle.RemoveItem Detalle.Rows - 1
'        If detalle.Rows > 1 Then
'            detalle.SelectionMode = cellSelectionByRow
'        End If
    End Sub





'=============================================================================
'PASA LOS DATOS DE LA ESTRUCTURA A LOS CONTROLES
'=============================================================================
Private Sub manual_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 27, Asc("r"), Asc("R")
            Call retorno
        Case Asc("a"), Asc("A"), 37
            Call anterior
        Case Asc("s"), Asc("S"), 39
            Call siguiente
        Case Asc("m"), Asc("M")
            Call modificar
        Case Asc("e"), Asc("E"), 46
            Call ELIMINAR
        Case Asc("i"), Asc("I")
            Call imprimir
        Case Asc("r"), Asc("R"), 46
            Call retorno
    End Select
End Sub

Private Sub op1_Click()
If lblDocumento.Caption <> "" And lblnombrecaja.Caption <> "" Then
dato2.text = leerfoliocaja(dato1.text, dato30.text)
dato2.SetFocus
End If

End Sub

Private Sub op2_Click()
If lblDocumento.Caption <> "" And lblnombrecaja.Caption <> "" Then

dato2.text = leerUltimofoliocaja(dato1.text, dato30.text)
dato2.SetFocus
End If

End Sub

'=============================================================================
'OPCIONES
'=============================================================================
    Private Sub opciones_FSCommand(ByVal command As String, ByVal args As String)
        Select Case command
            Case "modifica"
               
                If Verifica_Permiso(Me.Caption, "modifica") = True Then
                    Call modificar
                Else
                    MsgBox mensaje_nopermiso, vbCritical + vbOKOnly, "Permiso Denegado"
                End If
            Case "elimina"
               If Verifica_Permiso(Me.Caption, "elimina") = True Then
                    If MsgBox("DESEA REALMENTE ELIMINAR Si / No", vbYesNo) = vbYes Then
                    Call ELIMINAR
                    End If
                Else
                    MsgBox mensaje_nopermiso, vbCritical + vbOKOnly, "Permiso Denegado"
                End If
            Case "imprime"
                Call imprimir
            Case "movimientos"
            Case "historico"
            Case "retorno"
                Call retorno
            Case "anterior"
                Call anterior
            Case "siguiente"
                Call siguiente
        End Select
    End Sub
    
    Private Sub modificar()
    If dato1.text = "NP" Then
'     modifica = True

     Call ctrltostruct(False)
     Call eliminarVenta(v, detalle)
     Call eliminarPagos(p.tipodocumento, p.numeroDocumento, Format(p.fecha, "yyyy-mm-dd"), v.cabeza.caja)
     Call HabilitarCajas(Me, modifica)
     detalle.Rows = detalle.Rows + 1
     detalle.Cell(detalle.Rows - 1, 1).SetFocus

    End If
    
    End Sub
    
    Private Sub ELIMINAR()
        frmglosaeliminacion.Show vbModal
        Select Case MsgBox("Si desea eliminar el documento presione SI" & vbCrLf & "Si desea anular el documento presione NO" & vbCrLf & "Presione CANCELAR para volver", vbYesNoCancel, "Alerta")
            Case vbYes
                Call Eliminarcredito
                Call ctrltostruct(False)
                Call eliminarVenta(v, detalle)
                Call eliminarPagos(p.tipodocumento, p.numeroDocumento, p.fecha, v.cabeza.caja)
                Call retorno
                Call HabilitarCajas(Me, modifica)
                dato1.SetFocus
               Case vbNo
                If lblNulo.Caption = "" Then
                    v.detalle.fecha = dato5.text & "-" & dato4.text & "-" & dato3.text
                    Call anularDocumento(dato1.text, dato2.text, detalle, v.detalle)
                    Call eliminarPagos(dato1.text, dato2.text, Format(v.cabeza.fecha, "yyyy-mm-dd"), v.cabeza.caja)
                    Call eliminarDocManual(dato1.text, dato2.text)
                    Eliminarcredito
                End If
                
                Call retorno
                Call HabilitarCajas(Me, modifica)
                dato1.SetFocus
        End Select
    End Sub

    Private Sub imprimir()
        If nula = False Then
            Select Case dato1.text
                Case "BV"
                    numeroboleta = dato2.text
                    imprimirboleta
'                    ImpresionBoleta.Show
                Case "FV", "FE"
'                    Call imprimeFactura(dato1.text, dato2.text, impresion, data)
                     Call imprimeFactura(dato2.text, impresion, data)
                Case "GD", "GM", "ZE"
                    Call imprimeGuia(dato1.text, dato2.text, impresion, data)
'                Case "ZE"
                
                Case "NP"
                    Call imprimeNOTAPEDIDO(dato1.text, dato2.text, impresion, data)
                Case "CO"
                    Call imprimeNOTAPEDIDO(dato1.text, dato2.text, impresion, data)
                Case "NB", "NF"
                    Call imprimenotadecredito2(dato2.text, impresion, data, dato1.text)
            End Select
        Else
            Call MsgBox("Documento nulo no imprimible", vbOKOnly, "Mensaje")
        End If
    End Sub
    
    Public Sub retorno()
        Call LimpiarCajas(Me)
'       Call LimpiarLabels(Me)
        Call CARGAGRILLA(1, 9)
        detalle.Rows = 1
        detalle.Rows = 2
        detalle.Cell(1, 6).text = "0"
        modifica = False
        nula = False
        imprimio = False
        lectura = False
        autorizado = False
        dato11.text = "0"
        dato12.text = "0"
        dato13.text = "0"
        dato14.text = "0"
        dato15.text = "0"
        dato16.text = "0"
        dato17.text = "0"
        dato18.text = "0"
        dato19.text = "0"
        
        dato22.text = "0"
        dato20.text = ""
        dato21.text = ""
        dato25.text = "0"
        dato26.text = "0"
        lblSub.Caption = "0"
        lblneto.Caption = "0"
        lbliva.Caption = "0"
        lblIHA.Caption = "0"
        lbltotal.Caption = "0"
        lblDocumento.Caption = ""
        lbldv.Caption = ""
        lblRazon.Caption = ""
        LBLDIRECCION.Caption = ""
        lblComuna.Caption = ""
        LBLCIUDAD.Caption = ""
        lblvendedor.Caption = ""
        dato8.text = ""
        dato6.text = ""
        lblnombrecaja.Caption = ""
        lblfoliofiscal.Caption = ""
        lblfoliointerno.Caption = ""
        lbldvcajera.Caption = ""
        lblDVV.Caption = ""
        lblcajera.Caption = ""
        Call HabilitarCajas(Me, modifica)
   
        lbl20.Visible = False
        
        lbl26.Visible = False
        lbl30.Visible = False
        opciones.Visible = False
        dato1.SetFocus
    
    End Sub
        
    Private Sub anterior()
        If leerVenta(v, dato1.text, dato2.text, "<", data, detalle) = True Then
            structtoctrl
            nula = leerDocumentoNulo(dato1.text, dato2.text)
            If nula = True Then
                lblNulo.Caption = "DOCUMENTO ANULADO"
            Else
                lblNulo.Caption = ""
            End If
        End If
    End Sub
    
    Private Sub siguiente()
        If leerVenta(v, dato1.text, dato2.text, ">", data, detalle) = True Then
            structtoctrl
            nula = leerDocumentoNulo(dato1.text, dato2.text)
            If nula = True Then
                lblNulo.Caption = "DOCUMENTO ANULADO"
            Else
                lblNulo.Caption = ""
            End If
        End If
    End Sub
'=============================================================================
'OPCIONES
'=============================================================================

    Private Function leerPrecioEspecial(ByVal CODIGO As String) As String
        
        Dim campos(10, 3) As String
        Dim op As Integer
        Set sql = New sqlventas.sqlventa
        campos(0, 0) = "precioespecial"
        campos(1, 0) = ""
        
        campos(0, 2) = "sv_maestroclientes_especiales"
        
        condicion = "rut = '" & dato6.text & lbldv.Caption & "' AND sucursal = '" & dato7.text & "' AND codigo = '" & CODIGO & "'"
        op = 5
        sql.response = campos
        Set sql.conexion = ventas
        Call sql.sqlventas(op, condicion)
        If sql.Status = 0 Then
            leerPrecioEspecial = sql.response(0, 3)
        Else
            leerPrecioEspecial = ""
        End If
    End Function

    Private Sub opciones_GotFocus()
        MANUAL.SetFocus
    End Sub


Private Sub PREVENTA_GotFocus()
PREVENTA.text = ""

End Sub

Private Sub PREVENTA_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
frmdatos.Enabled = True
frmDetalle.Enabled = True
FRMPREVENTA.Visible = False
dato3.SetFocus
PREVENTA.text = ceros(PREVENTA)
leepreventa
          
End If


End Sub
Sub leepreventa()
    Dim resultados As rdoResultset
    Dim csql As New rdoQuery
    Dim rut As String
    Dim linea As Double
        Set csql.ActiveConnection = ventasRubro
        csql.sql = "SELECT codigo,descripcion,cantidad,precio,descuento,total,descuento2,vendedor "
        csql.sql = csql.sql + "FROM sv_documento_detalle_" + empresaActiva
        csql.sql = csql.sql + " WHERE local='" + empresaActiva + "' and tipo='PV' and numero='" + PREVENTA.text + "' "
        
        csql.Execute
        linea = detalle.Rows - 2
        
        If csql.RowsAffected > 0 Then
            Set resultados = csql.OpenResultset
            dato10.text = resultados(7)
            lblvendedor.Caption = leerNombreVendedor(dato10.text)
            While Not resultados.EOF
            detalle.Rows = detalle.Rows + 1
            linea = linea + 1
            detalle.Cell(linea, 1).text = resultados(0)
            detalle.Cell(linea, 2).text = resultados(1)
            detalle.Cell(linea, 3).text = resultados(2)
            detalle.Cell(linea, 4).text = "1"
            detalle.Cell(linea, 5).text = resultados(3)
            detalle.Cell(linea, 6).text = resultados(4)
            detalle.Cell(linea, 7).text = resultados(5)
            detalle.Cell(linea, 8).text = "0"
            
            resultados.MoveNext
            Wend
            resultados.Close
        Set resultados = Nothing

        End If
        Call sumaGrilla(detalle)
        
        detalle.Cell(detalle.Rows - 1, 1).SetFocus
 'Borrapreventa

End Sub
Sub borrapreventa()
    Dim resultados As rdoResultset
    Dim csql As New rdoQuery
    Dim rut As String
    Dim linea As Double
        Set csql.ActiveConnection = ventasRubro
        csql.sql = "DELETE "
        csql.sql = csql.sql + "FROM sv_documento_detalle_" + empresaActiva + " "
        csql.sql = csql.sql + "WHERE local='" + empresaActiva + "' and tipo='PV' and numero='" + PREVENTA.text + "' "
        csql.Execute
            
            
            Set csql.ActiveConnection = ventasRubro
        csql.sql = "DELETE "
        csql.sql = csql.sql + "FROM sv_documento_cabeza_" + empresaActiva + " "
        csql.sql = csql.sql + "WHERE local='" + empresaActiva + "' and tipo='PV' and numero='" + PREVENTA.text + "' "
        csql.Execute
    
        
End Sub

Sub CARGADATOSBOLETA()
dato3.text = Format(fechasistema, "dd")
dato4.text = Format(fechasistema, "mm")
dato5.text = Format(fechasistema, "yyyy")
dato6.text = "000000001"
lbldv.Caption = "9"
dato7.text = "0"
dato8.text = "NO"
dato9.text = "NO"
dato20.text = "NO"
dato21.text = "NO"
dato30.SetFocus
End Sub
Private Function leerstock(CODIGO) As Boolean
    
    Dim op As Integer
    Dim cantidad As Integer
    Dim campos(3, 3) As String
        Dim sql As New sqlventas.sqlventa
    
    campos(0, 0) = "stockactual"
    campos(1, 0) = ""
    
    campos(0, 2) = "r_maestroproductos_stock_" & rubro
    
    condicion = "codigo = '" & CODIGO & "' and año='" + Format(fechasistema, "yyyy") + "' and Bodega='" + empresaActiva + "' "
    
    op = 5
    sql.response = campos
    Set sql.conexion = gestionRubro
    Call sql.sqlventas(op, condicion)
    If sql.Status = 0 Then
        cantidad = CDbl(Val(sql.response(0, 3)))
        If cantidad <= 0 Then
           If MsgBox("PRODUCTO SIN STOCK", vbOKOnly, "ATENCION") = vbOK Then
           End If
           leerstock = False
           Else
           leerstock = True
           
        End If
    
    End If
End Function
Public Sub cargardeafuera()
dato2_KeyPress (13)
End Sub
Public Sub cargardeafuera2()
dato1_KeyPress (13)
End Sub
Private Function leerTipopagos(tipopag) As String
        Dim csql As rdoQuery
        Dim resultado As rdoResultset
        Dim i As Integer
        Dim diasmora As Double
        Dim saldo As Double
        Dim interes As Double
        
        Set csql = New rdoQuery
        Set csql.ActiveConnection = ventas

        csql.sql = "SELECT codigo,nombre "
        csql.sql = csql.sql & "FROM sv_tiposdepagoclientes "
        csql.sql = csql.sql & "where codigo='" & tipopag & "' "
        csql.Execute
        
        If csql.RowsAffected > 0 Then
        Set resultado = csql.OpenResultset
        
         leerTipopagos = resultado(0) & "-" & resultado(1)
        Else
      
          leerTipopagos = ""
        End If
        Set resultado = Nothing
        csql.Close
        Set csql = Nothing
    End Function
Public Sub Eliminarcredito()
        Dim csql As rdoQuery
        Dim resultado As rdoResultset
        Dim i As Integer
        Dim diasmora As Double
        Dim saldo As Double
        Dim interes As Double
        
        Set csql = New rdoQuery
        Set csql.ActiveConnection = ventas

        csql.sql = "delete from sv_cuotas_detalle "
        
        csql.sql = csql.sql & "where tipo='" + dato1.text + "' and numero='" + dato2.text + "' and rut='" & dato6.text & lbldv.Caption + "' "
        csql.Execute
        
        csql.Close
        Set csql = Nothing
    End Sub

Sub imprimirboleta()
Dim K As Double


CARGAGRILLAboleta


impresionboleta.Rows = impresionboleta.Rows + 1
impresionboleta.Range(1, 2, 1, 3).Merge
impresionboleta.Cell(1, 2).Font.Bold = True
impresionboleta.Cell(1, 2).text = Replace(leerNombreEmpresa(empresaActiva), "LTDA", "")

impresionboleta.Rows = impresionboleta.Rows + 1
impresionboleta.Range(2, 2, 2, 3).Merge
impresionboleta.Cell(2, 2).text = leerDireccionEmpresa(empresaActiva)

impresionboleta.Rows = impresionboleta.Rows + 1
impresionboleta.Range(3, 1, 3, 3).Merge
 impresionboleta.Cell(3, 1).text = "NRO.FISCAL : 77575340-18403"

impresionboleta.Rows = impresionboleta.Rows + 1
impresionboleta.Range(4, 1, 4, 3).Merge
impresionboleta.Cell(4, 1).text = leerNombreEmpresa(empresaActiva)

impresionboleta.Rows = impresionboleta.Rows + 1
impresionboleta.Range(5, 1, 5, 3).Merge
impresionboleta.Cell(5, 1).text = "RUT: " & Format(Mid(leerRutEmpresa(empresaActiva), 1, 9), "###,###,###") & "-" & Mid(leerRutEmpresa(empresaActiva), 10, 1)

impresionboleta.Rows = impresionboleta.Rows + 1
impresionboleta.Range(6, 1, 6, 3).Merge
impresionboleta.Cell(6, 1).text = "GIRO: " & leerGiroEmpresa(empresaActiva)

impresionboleta.Rows = impresionboleta.Rows + 1
impresionboleta.Range(7, 1, 7, 3).Merge
impresionboleta.Cell(7, 1).text = leerDireccionEmpresa(empresaActiva)

impresionboleta.Rows = impresionboleta.Rows + 1
impresionboleta.Range(8, 1, 8, 3).Merge
impresionboleta.Cell(8, 1).text = "Res.SII Nro 29 del 4 de Junio del 2003"

impresionboleta.Rows = impresionboleta.Rows + 1
impresionboleta.Range(9, 2, 9, 3).Merge
impresionboleta.Cell(9, 2).text = "BOLETA AUTORIZADA POR S.I.I "

impresionboleta.Rows = impresionboleta.Rows + 2
impresionboleta.Range(11, 1, 11, 3).Merge
impresionboleta.Cell(11, 1).text = " " & dato3.text & "/" & dato4.text & "/" & dato5.text

impresionboleta.Rows = impresionboleta.Rows + 1
impresionboleta.Range(12, 1, 12, 3).Merge
impresionboleta.Cell(12, 1).text = "BOLETA NRO. : " & dato2.text & "     CAJA NRO. : " & dato30.text

impresionboleta.Rows = impresionboleta.Rows + 1
impresionboleta.Range(13, 1, 13, 3).Merge
impresionboleta.Cell(13, 1).text = "----------------------------------------------------------"
i = impresionboleta.Rows

For K = 1 To detalle.Rows - 1
impresionboleta.Rows = impresionboleta.Rows + 1
impresionboleta.Cell(i, 1).text = Format(detalle.Cell(K, 3).text, "#0.000")
impresionboleta.Cell(i, 2).text = Mid(detalle.Cell(K, 2).text, 1, 20)
impresionboleta.Cell(i, 3).Alignment = cellRightCenter
impresionboleta.Cell(i, 3).text = Format(detalle.Cell(K, 7).text, "###,###,###")
i = i + 1
Next K

i = i + 1
impresionboleta.Rows = impresionboleta.Rows + 2
impresionboleta.Range(i, 1, i, 2).Merge
impresionboleta.Cell(i, 1).Font.Bold = True
impresionboleta.Cell(i, 3).Font.Bold = True
impresionboleta.Cell(i, 1).text = "SUBTOTAL"
impresionboleta.Cell(i, 3).Alignment = cellRightCenter
'impresionboleta.Cell(i, 3).text = Format(CDbl(dato26.text) + CDbl(dato12.text), "###,###,###")
impresionboleta.Cell(i, 3).text = dato26.text 'Format(CDbl(dato26.text), "$ ##,###,##0")

'If CDbl(dato12.text) > 0 Then
'    i = i + 1
'    impresionboleta.Rows = impresionboleta.Rows + 1
'    impresionboleta.Range(i, 1, i, 3).Merge
'    impresionboleta.Cell(i, 1).text = "DESCUENTO SUBTOTAL"
'
'    i = i + 1
'    impresionboleta.Rows = impresionboleta.Rows + 1
'    impresionboleta.Range(i, 1, i, 2).Merge
'    impresionboleta.Cell(i, 1).Font.Bold = True
'    impresionboleta.Cell(i, 3).Font.Bold = True
'    impresionboleta.Cell(i, 1).text = "Descuento al Total"
'    impresionboleta.Cell(i, 3).Alignment = cellRightCenter
'    impresionboleta.Cell(i, 3).text = Format(dato12.text, "###,###,###") & "-"
'
'    i = i + 2
'    impresionboleta.Rows = impresionboleta.Rows + 2
'    impresionboleta.Range(i, 1, i, 2).Merge
'    impresionboleta.Cell(i, 1).Font.Bold = True
'    impresionboleta.Cell(i, 3).Font.Bold = True
'    impresionboleta.Cell(i, 1).text = "SUBTOTAL"
'    impresionboleta.Cell(i, 3).Alignment = cellRightCenter
'    impresionboleta.Cell(i, 3).text = Format(CDbl(dato26.text), "###,###,###")
'
'End If


Call leerpagosdocumento(dato1.text, dato30.text, dato27.text, dato5.text & "-" & dato4.text & "-" & dato3.text, i)

i = i + 3
impresionboleta.Rows = impresionboleta.Rows + 3
impresionboleta.Range(i, 1, i, 2).Merge
impresionboleta.Cell(i, 1).Font.Bold = True
impresionboleta.Cell(i, 1).Font.Size = 10
impresionboleta.Cell(i, 1).text = "TOTAL"
impresionboleta.Cell(i, 3).Font.Bold = True
impresionboleta.Cell(i, 3).Alignment = cellRightCenter
impresionboleta.Cell(i, 3).text = dato26.text 'Format(dato26.text, "$ ##,###,##0")

i = i + 2
impresionboleta.Rows = impresionboleta.Rows + 2
impresionboleta.Range(i, 1, i, 2).Merge
impresionboleta.Cell(i, 1).text = "AGRADECEMOS"

i = i + 1
impresionboleta.Rows = impresionboleta.Rows + 1
impresionboleta.Range(i, 1, i, 2).Merge
impresionboleta.Cell(i, 1).text = "SU PREFERENCIA"

i = i + 1
impresionboleta.Rows = impresionboleta.Rows + 1
impresionboleta.Range(i, 1, i, 3).Merge
impresionboleta.Cell(i, 1).text = "----------------------------------------------------------"

i = i + 1
impresionboleta.Rows = impresionboleta.Rows + 1
impresionboleta.Range(i, 1, i, 2).Merge
impresionboleta.Cell(i, 1).text = "TRANSACCION NRO.: " & dato27.text
impresionboleta.Cell(i, 3).Alignment = cellRightCenter
impresionboleta.Cell(i, 3).text = "V13"

i = i + 1
impresionboleta.Rows = impresionboleta.Rows + 1
impresionboleta.Range(i, 1, i, 2).Merge
impresionboleta.Cell(i, 1).text = " IB77001995"

i = i + 2
impresionboleta.Rows = impresionboleta.Rows + 2
impresionboleta.Range(i, 2, i, 3).Merge
impresionboleta.Cell(i, 2).text = "INICIO COMENTARIO"

i = i + 1
impresionboleta.Rows = impresionboleta.Rows + 1
impresionboleta.Range(i, 1, i, 2).Merge
impresionboleta.Cell(i, 1).text = "CAJERO(A): " & lblcajera.Caption

i = i + 1
impresionboleta.Rows = impresionboleta.Rows + 1
impresionboleta.Range(i, 1, i, 2).Merge
impresionboleta.Cell(i, 1).text = "NUMERO OFFSET:" & dato27.text

i = i + 1
impresionboleta.Rows = impresionboleta.Rows + 1
impresionboleta.Range(i, 1, i, 3).Merge
impresionboleta.Cell(i, 1).Font.Bold = True
impresionboleta.Cell(i, 1).text = "COPIA FIEL A LA ORIGINAL"

i = i + 1
impresionboleta.Rows = impresionboleta.Rows + 2
impresionboleta.Range(i, 2, i, 3).Merge
impresionboleta.Cell(i, 2).text = "FIN COMENTARIO"

impresionboleta.PrintPreview
End Sub
Sub leerpagosdocumento(TIPO, caja, NUMERO, fecha, contador)
Dim csql As New rdoQuery
Dim resultados As rdoResultset
Dim K As Double
Dim PAGADO As Double


Set csql.ActiveConnection = ventasRubro
csql.sql = "select  p.lineapago,p.tipopago,p.monto,tp.nombre "
csql.sql = csql.sql & "from sv_documento_pagos_" & empresaActiva & " as p , " & baseVentas & ".sv_tiposdepagoclientes as tp "
csql.sql = csql.sql & "where p.tipopago=tp.codigo "
csql.sql = csql.sql & "and  p.local='" & empresaActiva & "' and p.tipo='" & TIPO & "' and p.caja='" & caja & "' and p.numero='" & NUMERO & "' and p.fecha='" & fecha & "'"
csql.Execute

If csql.RowsAffected > 0 Then
    Set resultados = csql.OpenResultset
    While Not resultados.EOF
        contador = contador + 1
        impresionboleta.Rows = impresionboleta.Rows + 1
        impresionboleta.Range(contador, 1, contador, 2).Merge
'        impresionboleta.Cell(contador, 1).Font.Bold = True
        impresionboleta.Cell(contador, 1).Font.Size = 10
        impresionboleta.Cell(contador, 1).text = resultados(3)
        
        contador = contador + 1
        impresionboleta.Rows = impresionboleta.Rows + 1
        impresionboleta.Range(contador, 1, contador, 2).Merge
        impresionboleta.Cell(contador, 1).Font.Size = 10
        impresionboleta.Cell(contador, 1).text = ".................................."
        impresionboleta.Cell(contador, 3).Alignment = cellRightCenter
        impresionboleta.Cell(contador, 3).text = resultados(2)
        PAGADO = PAGADO + resultados(2)
        resultados.MoveNext
    Wend
    
    If PAGADO - CDbl(dato26.text) > 0 Then
        contador = contador + 2
        impresionboleta.Rows = impresionboleta.Rows + 2
        impresionboleta.Range(contador, 1, contador, 2).Merge
'        impresionboleta.Cell(contador, 1).Font.Bold = True
        impresionboleta.Cell(contador, 1).Font.Size = 10
        impresionboleta.Cell(contador, 1).text = "VUELTO"
        impresionboleta.Cell(contador, 3).Alignment = cellRightCenter
        impresionboleta.Cell(contador, 3).text = PAGADO - CDbl(dato26.text)
      
    End If
    
    i = contador
End If
If LEErcuotas(caja, dato1, dato2, (Format(fechasistema, "yyyy-mm-dd"))) = True Then
End If
End Sub

Public Function LEErcuotas(caja, TIPO, NUMERO, fecha) As Boolean  '<- SE LE AGREGO LA CAJA
        Dim csql As rdoQuery
        Dim resultado As rdoResultset
        Dim i As Integer
        Dim diasmora As Double
        Dim saldo As Double
        Dim interes As Double
        Set csql = New rdoQuery
        Set csql.ActiveConnection = ventas
        csql.sql = "SELECT dc.rut,dc.montocredito,mc.diapago,mc.cupodirecto,mc.cupoutilizadotmp,mc.nombre,mc.direccion,dc.cantidadcuotas,dc.montocuota,dc.vencimientooriginal,dc.tipo,dc.numero,dc.abono,dc.numerocuota,dc.vencimientoactual,dc.pie,dc.montoventa,dc.cajera "
        csql.sql = csql.sql & "FROM sv_cuotas_detalle as dc," + baseVentas + ".sv_maestroclientes as mc "
        csql.sql = csql.sql & "WHERE tipo = '" & TIPO & "' and numero='" & NUMERO & "' and caja = '" & caja & "' and fechacompra = '" & fecha & "' and dc.rut=mc.rut  "
        csql.sql = csql.sql & "order by vencimientooriginal "
        csql.Execute
        
        If csql.RowsAffected > 0 Then

            Set resultado = csql.OpenResultset
       
     '   total.Caption = Format(resultado(1), "###,###,##0")
     '   DIAPAGO.text = resultado(2)
     '   lblCupo.Caption = Format(resultado(3), "###,###,##0")
     '   lblUtilizado.Caption = Format(resultado(4), "###,###,##0")
     '   lblDisponible.Caption = Format(resultado(3) - resultado(4), "###,###,##0")
     '   CREDITO.text = resultado(1)
     '   MONTO.text = resultado(16)
     '   PIE.text = resultado(15)
     '   rut2.text = Mid(resultado(0), 1, 9)
     '   lblDV.Caption = Mid(resultado(0), 10, 1)
     '   lblNombre.Caption = resultado(5)
     '   lblDireccion.Caption = resultado(6)
     '   CUOTAS.text = resultado(7)
     '   VALORCUOTA.text = resultado(8)
     '   DIAC.text = Format(resultado(9), "dd")
     '   MESC.text = Format(resultado(9), "mm")
     '   AÑOC.text = Format(resultado(9), "yyyy")
         Call IMPRIMEcredito(dato1, dato2, fecha, Mid(resultado(0), 1, 9), resultado(5), resultado(6), dato26, resultado(1), resultado(7), resultado(8), resultado(15), Format(resultado(9), "dd"), Format(resultado(9), "mm"), Format(resultado(9), "yyyy"), (resultado(17) & rut(resultado(17))))
      
        
        LEErcuotas = True
        Grid1.Rows = 1
        While Not resultado.EOF
        Grid1.Rows = Grid1.Rows + 1
        Grid1.Cell(Grid1.Rows - 1, 2).text = resultado(13) & " / " & resultado(7)
        Grid1.Cell(Grid1.Rows - 1, 3).text = Format(resultado(14), "dd/mm/yyyy")
        Grid1.Cell(Grid1.Rows - 1, 4).text = Format(resultado(8), "###,###,###")
'        Grid1.Cell(Grid1.Rows - 1, 5).text = Format(resultado(12), "###,###,###")
        saldo = resultado(8) - resultado(12)
        diasmora = 0
        interes = 0
        'total = saldo + interes
        
        'Grid1.Cell(Grid1.Rows - 1, 6).text = Format(saldo, "###,###,###")
        'Grid1.Cell(Grid1.Rows - 1, 7).text = diasmora
        'Grid1.Cell(Grid1.Rows - 1, 8).text = interes
        'Grid1.Cell(Grid1.Rows - 1, 9).text = total
  
        resultado.MoveNext
        Wend
        Else
        LEErcuotas = False

        End If
        Set resultado = Nothing
        csql.Close
        Set csql = Nothing
    End Function

Public Sub IMPRIMEcredito(TIPO, NUMERO, fecha, ClienteRut, ClienteNombre, ClienteDireccion, totalventa, TotalCredito, CUOTAS, ValorCuotas, MontoPie, DIAC, MESC, AÑOC, CajeroRut)
Dim K As Integer
nombrecajero = leerNombreCajera(CajeroRut)
    Dim numfic As Long
    ''''''''''''''''''
    numfic = 20
    If impresoracredito = "0" Then
    Open "impresion.txt" For Output As #numfic
    End If
    If impresoracredito = "1" Then
    Open "COM1:4800,N,8,1,CD0,CS0,DS0,OP0,RS,TB100,RB100" For Output As #numfic
    End If
    If impresoracredito = "2" Then
    Open "LPT1" For Output As #numfic
    End If
    For K = 1 To 2
    Print #numfic, Chr$(27); Chr$(64) '
    Print #numfic, ""
    Print #numfic, "          VALE DE CREDITO         "
    Print #numfic, "          ================        "
    Print #numfic,
    Print #numfic, "TIPO    :"; dato1
    Print #numfic, "NUMERO  :"; dato2
    Print #numfic, "FECHA   :"; fecha
    Print #numfic, "CLIENTE :"; ClienteRut
    Print #numfic, "NOMBRE  :"; ClienteNombre
    Print #numfic, "DIREC.  :"; ClienteDireccion
    Print #numfic,
    Print #numfic, "MONTO VENTA   :"; Format(totalventa, " $ ###,###,###")
    Print #numfic, "MONTO PIE     :"; Format(MontoPie, " $ ###,###,###")
    Print #numfic, "MONTO CREDITO :"; Format(TotalCredito, " $ ###,###,###")
    Print #numfic, "YO AUTORIZO SEGUN CONTRATO PALGUIN LTDA "
    Print #numfic, "CARGAR A MI CUENTA "
    Print #numfic, CUOTAS & " CUOTAS de " & Format(ValorCuotas, "$ ###,###,###")
    Print #numfic, "Primer vencimiento:"; DIAC + "-" + MESC + "-" + AÑOC
    Print #numfic,
    Print #numfic,
    Print #numfic,
    Print #numfic, "              _______________             "
    Print #numfic, "               FIRMA CLIENTE             "
    Print #numfic,
    Print #numfic, "CAJERA(o):" + nombrecajero
    Print #numfic,
    Print #numfic,
    Print #numfic,
    Print #numfic,
    Print #numfic, Chr(27); "i"
    Next K
    Close #numfic
    If impresoracredito = "0" Then Shell "notepad impresion.txt"
End Sub
Private Sub LeerVendedorDocumento(ByVal loc, TIPO, NUMERO, caja, fecha)
        Dim tabla As String
        Dim csql As New rdoQuery
        Dim resultados As rdoResultset
        Set csql.ActiveConnection = ventasRubro
        csql.sql = "SELECT VENDEDOR FROM sv_documento_detalle_" + empresaActiva + " "
        csql.sql = csql.sql & "WHERE local = '" & loc & "' AND tipo = '" & TIPO & "' AND numero = '" & NUMERO & " '"
        csql.sql = csql.sql & " and caja='" + caja + "' and fecha='" + fecha + "' limit 1"
        csql.Execute
        If csql.RowsAffected > 0 Then
            Set resultados = csql.OpenResultset
            dato10 = resultados(0)
        resultados.Close
        csql.sql = " SELECT concat('Linea :', ' ',cv.linea,' - Vendedor : ',  mv.nombre) as vendedores from"
        csql.sql = csql.sql & " sv_documento_detalle_" & empresaActiva & " as cv, " & cliente_sql & "ventas.sv_maestrovendedores as mv"
        csql.sql = csql.sql & " WHERE cv.local = '" & loc & "' AND cv.tipo = '" & TIPO & "' AND cv.numero = '" & NUMERO & " '"
        csql.sql = csql.sql & " and cv.caja='" + caja + "' and cv.fecha='" + fecha + "' and cv.vendedor=mv.rut"
        csql.Execute
        
        If csql.RowsAffected > 0 Then
            Set resultados = csql.OpenResultset
         While Not resultados.EOF
            vendedores.AddItem resultados(0)
            
            
         resultados.MoveNext
         Wend
         resultados.Close
         csql.Close
        End If
        Else
        End If

End Sub

