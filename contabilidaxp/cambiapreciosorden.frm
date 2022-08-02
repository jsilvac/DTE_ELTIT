VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash9b.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{4777436C-EB8C-4596-98A8-EBCF98683969}#1.0#0"; "Flexcell.ocx"
Object = "{ADBBAED6-B16F-40DC-80DF-B44910CBA76C}#1.0#0"; "Frame-Xp.ocx"
Begin VB.Form preciosorden 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mestro de Productos"
   ClientHeight    =   9915
   ClientLeft      =   390
   ClientTop       =   -15
   ClientWidth     =   15120
   DrawStyle       =   5  'Transparent
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   661
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1008
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command6 
      Caption         =   "Command6"
      Height          =   375
      Left            =   120
      TabIndex        =   116
      Top             =   1080
      Visible         =   0   'False
      Width           =   615
   End
   Begin XPFrame.FrameXp sipack 
      Height          =   825
      Left            =   6795
      TabIndex        =   113
      Top             =   4455
      Width           =   8115
      _ExtentX        =   14314
      _ExtentY        =   1455
      BackColor       =   16744576
      Caption         =   "DATOS DEL PACK"
      CaptionEstilo3D =   1
      BackColor       =   16744576
      ForeColor       =   8438015
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.CommandButton Command5 
         BackColor       =   &H000000FF&
         Caption         =   "MAXIMIZAR"
         Height          =   285
         Left            =   6795
         Style           =   1  'Graphical
         TabIndex        =   115
         Top             =   0
         Width           =   1140
      End
      Begin FlexCell.Grid grillapack 
         Height          =   2175
         Left            =   45
         TabIndex        =   114
         Top             =   270
         Width           =   8025
         _ExtentX        =   14155
         _ExtentY        =   3836
         Cols            =   5
         DefaultFontSize =   8.25
         Rows            =   30
      End
   End
   Begin XPFrame.FrameXp Frmcodigobarra 
      Height          =   1545
      Left            =   3120
      TabIndex        =   98
      Top             =   240
      Visible         =   0   'False
      Width           =   2805
      _ExtentX        =   4948
      _ExtentY        =   2725
      BackColor       =   49344
      Caption         =   "CODIGO BARRA"
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
      Begin VB.CommandButton RETORNO_CODIGO 
         Caption         =   "RETORNO"
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
         Left            =   1395
         TabIndex        =   101
         Top             =   1170
         Width           =   1140
      End
      Begin VB.CommandButton CAMBIA_CODIGO_BARRA 
         Caption         =   "MODIFICA"
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
         Left            =   180
         TabIndex        =   100
         Top             =   1170
         Width           =   1095
      End
      Begin VB.TextBox dato30 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   135
         MaxLength       =   13
         TabIndex        =   99
         Top             =   630
         Width           =   2310
      End
      Begin VB.Label Label25 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "NUEVO CODIGO"
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
         Height          =   330
         Left            =   135
         TabIndex        =   102
         Top             =   360
         Width           =   2310
      End
   End
   Begin VB.TextBox codigoenvase 
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
      Left            =   7920
      MaxLength       =   13
      TabIndex        =   93
      Tag             =   "codigobarra"
      Top             =   720
      Width           =   1575
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   285
      Left            =   6300
      TabIndex        =   89
      Top             =   3555
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   555
      Left            =   6345
      TabIndex        =   85
      Top             =   5265
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.CommandButton finmodifica 
      BackColor       =   &H00FF8080&
      Caption         =   "Finalizar Modificacion"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   1800
      Style           =   1  'Graphical
      TabIndex        =   70
      Top             =   8775
      Width           =   3435
   End
   Begin XPFrame.FrameXp FrameXp4 
      Height          =   1500
      Left            =   6660
      TabIndex        =   67
      Top             =   8370
      Width           =   8415
      _ExtentX        =   14843
      _ExtentY        =   2646
      BackColor       =   16773879
      Caption         =   "Planilla"
      CaptionEstilo3D =   1
      BackColor       =   16773879
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.CommandButton historico 
         BackColor       =   &H0080FF80&
         Caption         =   "Historico Cambios de Precio"
         Height          =   345
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   75
         Top             =   1080
         Width           =   2265
      End
      Begin XPFrame.FrameXp frmLista 
         Height          =   375
         Left            =   2520
         TabIndex        =   68
         Top             =   1035
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   661
         BackColor       =   49344
         Caption         =   "Precio Ofertas"
         CaptionEstilo3D =   1
         BackColor       =   49344
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
      End
      Begin FlexCell.Grid lista 
         Height          =   720
         Left            =   90
         TabIndex        =   69
         Top             =   270
         Width           =   8115
         _ExtentX        =   14314
         _ExtentY        =   1270
         Cols            =   5
         DefaultFontName =   "Arial"
         DefaultFontSize =   8.25
         Rows            =   1
         SelectionMode   =   1
      End
   End
   Begin XPFrame.FrameXp FrameXp2 
      Height          =   3600
      Left            =   135
      TabIndex        =   54
      Top             =   4815
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   6350
      BackColor       =   16773879
      Caption         =   "Datos Logisticos"
      CaptionEstilo3D =   1
      BackColor       =   16773879
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.CheckBox noactualiza 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "No Actualize Precio al Recepcionar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   480
         Left            =   3240
         MaskColor       =   &H00FFFFFF&
         TabIndex        =   76
         Top             =   360
         Width           =   2580
      End
      Begin XPFrame.FrameXp FrameXp3 
         Height          =   2145
         Left            =   45
         TabIndex        =   59
         Top             =   1800
         Width           =   6000
         _ExtentX        =   10583
         _ExtentY        =   3784
         BackColor       =   16773879
         Caption         =   "Datos Adicionales"
         CaptionEstilo3D =   1
         BackColor       =   16773879
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Begin VB.TextBox dato17 
            BackColor       =   &H00C0FFFF&
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
            MaxLength       =   14
            TabIndex        =   17
            Tag             =   "dun14"
            Top             =   1440
            Width           =   1695
         End
         Begin VB.TextBox dato16 
            BackColor       =   &H00C0FFFF&
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
            TabIndex        =   16
            Tag             =   "referenciaproveedor"
            Top             =   1080
            Width           =   2295
         End
         Begin VB.TextBox dato15 
            BackColor       =   &H00C0FFFF&
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
            MaxLength       =   16
            TabIndex        =   15
            Tag             =   "glosaregistradoras"
            Top             =   720
            Width           =   2460
         End
         Begin VB.TextBox dato14 
            BackColor       =   &H00C0FFFF&
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
            MaxLength       =   30
            TabIndex        =   14
            Tag             =   "glosaflejes"
            Top             =   360
            Width           =   3495
         End
         Begin VB.CheckBox Check7 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            Caption         =   "Fleje Imprimible"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   255
            Left            =   4185
            TabIndex        =   60
            Top             =   1440
            Value           =   1  'Checked
            Width           =   1695
         End
         Begin VB.Label Label20 
            BackColor       =   &H00F5C9B1&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " Codigo dun 14"
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
            Height          =   285
            Left            =   180
            TabIndex        =   64
            Top             =   1440
            Width           =   1920
         End
         Begin VB.Label Label13 
            BackColor       =   &H00F5C9B1&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " Codigo Proveedor"
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
            Height          =   285
            Left            =   180
            TabIndex        =   63
            Top             =   1080
            Width           =   1920
         End
         Begin VB.Label Label12 
            BackColor       =   &H00F5C9B1&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " Glosa Registradoras"
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
            Height          =   285
            Left            =   180
            TabIndex        =   62
            Top             =   720
            Width           =   1965
         End
         Begin VB.Label Label7 
            BackColor       =   &H00F5C9B1&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " Glosa Flejes"
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
            Height          =   285
            Left            =   180
            TabIndex        =   61
            Top             =   360
            Width           =   1965
         End
      End
      Begin VB.TextBox dato13 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
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
         Left            =   2250
         MaxLength       =   7
         TabIndex        =   13
         Tag             =   "cantidadporembalaje"
         Text            =   "1"
         Top             =   1410
         Width           =   1215
      End
      Begin VB.TextBox dato12 
         BackColor       =   &H00C0FFFF&
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
         Left            =   1860
         MaxLength       =   10
         TabIndex        =   12
         Tag             =   "tipoembalaje"
         Top             =   1050
         Width           =   1215
      End
      Begin VB.TextBox dato10 
         BackColor       =   &H00C0FFFF&
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
         Left            =   1860
         MaxLength       =   3
         TabIndex        =   10
         Tag             =   "unidadmedida"
         Top             =   330
         Width           =   495
      End
      Begin VB.TextBox dato11 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
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
         Left            =   1860
         MaxLength       =   7
         TabIndex        =   11
         Tag             =   "contenido"
         Top             =   690
         Width           =   1215
      End
      Begin VB.Label FECHACREACION 
         BackStyle       =   0  'Transparent
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
         Left            =   3555
         TabIndex        =   86
         Top             =   1395
         Width           =   2355
      End
      Begin VB.Label Label58 
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Unidades x Embalaje"
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
         Height          =   285
         Left            =   180
         TabIndex        =   58
         Top             =   1410
         Width           =   2010
      End
      Begin VB.Label Label57 
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Tipo Embalaje"
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
         Height          =   285
         Left            =   180
         TabIndex        =   57
         Top             =   1050
         Width           =   1695
      End
      Begin VB.Label Label3 
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " U/M Envase"
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
         Height          =   285
         Left            =   180
         TabIndex        =   56
         Top             =   330
         Width           =   1695
      End
      Begin VB.Label Label39 
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Contenido x Envase"
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
         Height          =   285
         Left            =   180
         TabIndex        =   55
         Top             =   690
         Width           =   1635
      End
   End
   Begin XPFrame.FrameXp clasificacion 
      Height          =   3405
      Left            =   135
      TabIndex        =   21
      Top             =   1350
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   6006
      BackColor       =   16773879
      Caption         =   "Datos de Clasificacion"
      CaptionEstilo3D =   1
      BackColor       =   16773879
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.CheckBox descontinuado 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "Producto Descontinuado"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   240
         Left            =   3285
         MaskColor       =   &H00FFFFFF&
         TabIndex        =   74
         Top             =   2655
         Width           =   2580
      End
      Begin VB.TextBox dato7 
         BackColor       =   &H00C0FFFF&
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
         Left            =   1440
         MaxLength       =   5
         TabIndex        =   6
         Tag             =   "codigomarca"
         Top             =   1920
         Width           =   735
      End
      Begin VB.TextBox dato6 
         BackColor       =   &H00C0FFFF&
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
         Left            =   1440
         MaxLength       =   5
         TabIndex        =   5
         Tag             =   "codigoimpuesto"
         Top             =   1560
         Width           =   735
      End
      Begin VB.TextBox dato5 
         BackColor       =   &H00C0FFFF&
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
         Left            =   1440
         MaxLength       =   5
         TabIndex        =   4
         Tag             =   "codigolinea"
         Top             =   1200
         Width           =   735
      End
      Begin VB.TextBox dato4 
         BackColor       =   &H00C0FFFF&
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
         Left            =   1440
         MaxLength       =   5
         TabIndex        =   3
         Tag             =   "codigodepto"
         Top             =   840
         Width           =   735
      End
      Begin VB.TextBox dato3 
         BackColor       =   &H00C0FFFF&
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
         Left            =   1440
         MaxLength       =   5
         TabIndex        =   2
         Tag             =   "codigoseccion"
         Top             =   495
         Width           =   735
      End
      Begin VB.TextBox dato8 
         BackColor       =   &H00C0FFFF&
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
         Left            =   1440
         MaxLength       =   5
         TabIndex        =   7
         Tag             =   "codigotemporada"
         Top             =   2280
         Width           =   735
      End
      Begin VB.TextBox dato9 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
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
         Left            =   1440
         MaxLength       =   9
         TabIndex        =   8
         Tag             =   "proveedor"
         Top             =   2640
         Width           =   1215
      End
      Begin VB.Label dsctomarca 
         Alignment       =   1  'Right Justify
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
         ForeColor       =   &H8000000D&
         Height          =   285
         Left            =   5250
         TabIndex        =   53
         Top             =   1920
         Width           =   615
      End
      Begin VB.Label dsctoimpuesto 
         Alignment       =   1  'Right Justify
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
         ForeColor       =   &H8000000D&
         Height          =   285
         Left            =   5250
         TabIndex        =   52
         Top             =   1560
         Width           =   615
      End
      Begin VB.Label dsctolinea 
         Alignment       =   1  'Right Justify
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
         ForeColor       =   &H8000000D&
         Height          =   285
         Left            =   5250
         TabIndex        =   51
         Top             =   1200
         Width           =   615
      End
      Begin VB.Label dsctodpto 
         Alignment       =   1  'Right Justify
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
         ForeColor       =   &H8000000D&
         Height          =   285
         Left            =   5250
         TabIndex        =   50
         Top             =   840
         Width           =   615
      End
      Begin VB.Label dsctoseccion 
         Alignment       =   1  'Right Justify
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
         ForeColor       =   &H8000000D&
         Height          =   285
         Left            =   5250
         TabIndex        =   49
         Top             =   480
         Width           =   615
      End
      Begin VB.Label nombremarca 
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
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   2250
         TabIndex        =   48
         Top             =   1920
         Width           =   3015
      End
      Begin VB.Label nombreimpUESTO 
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
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   2250
         TabIndex        =   47
         Top             =   1560
         Width           =   3015
      End
      Begin VB.Label nombrelinea 
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
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   2340
         TabIndex        =   46
         Top             =   1170
         Width           =   3015
      End
      Begin VB.Label Label27 
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Marca"
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
         Height          =   285
         Left            =   60
         TabIndex        =   45
         Top             =   1920
         Width           =   1365
      End
      Begin VB.Label nombredepto 
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
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   2250
         TabIndex        =   44
         Top             =   840
         Width           =   3015
      End
      Begin VB.Label nombreseccion 
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
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   2250
         TabIndex        =   43
         Top             =   480
         Width           =   3015
      End
      Begin VB.Label Label10 
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Impuesto"
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
         Height          =   285
         Left            =   60
         TabIndex        =   42
         Top             =   1560
         Width           =   1365
      End
      Begin VB.Label Label6 
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Linea"
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
         Height          =   285
         Left            =   60
         TabIndex        =   41
         Top             =   1200
         Width           =   1365
      End
      Begin VB.Label Label5 
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Departamento"
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
         Height          =   285
         Left            =   60
         TabIndex        =   40
         Top             =   840
         Width           =   1365
      End
      Begin VB.Label Label4 
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Seccion"
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
         Height          =   285
         Left            =   60
         TabIndex        =   39
         Top             =   480
         Width           =   1365
      End
      Begin VB.Label nombretemporada 
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
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   2250
         TabIndex        =   38
         Top             =   2280
         Width           =   3015
      End
      Begin VB.Label Label28 
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Temporada"
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
         Height          =   285
         Left            =   60
         TabIndex        =   37
         Top             =   2280
         Width           =   1365
      End
      Begin VB.Label Label40 
         BackStyle       =   0  'Transparent
         Caption         =   "(%) dcto"
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
         Left            =   5190
         TabIndex        =   36
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label16 
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Proveedor"
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
         Height          =   285
         Left            =   60
         TabIndex        =   35
         Top             =   2640
         Width           =   1365
      End
      Begin VB.Label nombreproveedor 
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
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   120
         TabIndex        =   34
         Top             =   3000
         Width           =   5775
      End
      Begin VB.Label dsctotemporada 
         Alignment       =   1  'Right Justify
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
         ForeColor       =   &H8000000D&
         Height          =   285
         Left            =   5250
         TabIndex        =   33
         Top             =   2280
         Width           =   615
      End
      Begin VB.Label dv 
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
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
         Left            =   2700
         TabIndex        =   9
         Top             =   2640
         Width           =   285
      End
   End
   Begin XPFrame.FrameXp comercial 
      Height          =   4095
      Left            =   6705
      TabIndex        =   27
      Top             =   1305
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   7223
      BackColor       =   16773879
      Caption         =   "Datos Comercializacion Valores C/IVA"
      CaptionEstilo3D =   1
      BackColor       =   16773879
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.CommandButton CompraVenta 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Ver Ultimos Precios"
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
         Left            =   6255
         Style           =   1  'Graphical
         TabIndex        =   111
         Top             =   360
         Width           =   1920
      End
      Begin VB.CommandButton cmdpack 
         BackColor       =   &H0000FFFF&
         Caption         =   "VER PACK"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   6300
         Style           =   1  'Graphical
         TabIndex        =   106
         Top             =   765
         Visible         =   0   'False
         Width           =   1860
      End
      Begin VB.CheckBox checkpack 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFF2F7&
         Caption         =   "Venta x Pack"
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
         Left            =   4230
         MaskColor       =   &H80000010&
         TabIndex        =   105
         Top             =   540
         Width           =   1635
      End
      Begin VB.CheckBox envase 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFF2F7&
         Caption         =   "Es Envase"
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
         Left            =   4230
         MaskColor       =   &H80000010&
         TabIndex        =   97
         Top             =   270
         Width           =   1365
      End
      Begin VB.TextBox descuento 
         Alignment       =   1  'Right Justify
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
         Left            =   2460
         Locked          =   -1  'True
         MaxLength       =   2
         TabIndex        =   90
         Tag             =   "codigobarra"
         Text            =   "0"
         Top             =   1320
         Width           =   690
      End
      Begin XPFrame.FrameXp FRMPESABLE 
         Height          =   1860
         Left            =   180
         TabIndex        =   79
         Top             =   990
         Width           =   2130
         _ExtentX        =   3757
         _ExtentY        =   3281
         BackColor       =   12648384
         Caption         =   "TIPO VENTA"
         CaptionEstilo3D =   1
         BackColor       =   12648384
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Begin VB.Label Label18 
            BackStyle       =   0  'Transparent
            Caption         =   "3 = REQ.CANTIDAD"
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
            Left            =   135
            TabIndex        =   83
            Top             =   1440
            Width           =   1860
         End
         Begin VB.Label Label17 
            BackStyle       =   0  'Transparent
            Caption         =   "2 = REQ.PRECIO"
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
            Left            =   135
            TabIndex        =   82
            Top             =   1080
            Width           =   1860
         End
         Begin VB.Label Label15 
            BackStyle       =   0  'Transparent
            Caption         =   "1 = REQ.PESO MTS"
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
            Left            =   135
            TabIndex        =   81
            Top             =   720
            Width           =   1860
         End
         Begin VB.Label Label14 
            BackStyle       =   0  'Transparent
            Caption         =   "0 = NORMAL"
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
            Left            =   135
            TabIndex        =   80
            Top             =   360
            Width           =   1860
         End
      End
      Begin VB.TextBox dato20 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
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
         Left            =   2460
         MaxLength       =   1
         TabIndex        =   77
         Tag             =   "margen"
         Text            =   "0"
         Top             =   1035
         Width           =   1695
      End
      Begin VB.TextBox dato18 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
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
         Left            =   2460
         MaxLength       =   14
         TabIndex        =   18
         Tag             =   "pcosto"
         Top             =   360
         Width           =   1695
      End
      Begin VB.TextBox dato19 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
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
         Left            =   2460
         MaxLength       =   14
         TabIndex        =   19
         Tag             =   "margen"
         Text            =   "20"
         Top             =   720
         Width           =   1695
      End
      Begin VB.CheckBox Check1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFF2F7&
         Caption         =   "Calcula Precios Automaticos"
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
         Left            =   5220
         TabIndex        =   29
         Top             =   -45
         Value           =   1  'Checked
         Visible         =   0   'False
         Width           =   2775
      End
      Begin VB.CommandButton btn_cambioprecios 
         BackColor       =   &H0080C0FF&
         Caption         =   "Cambiar Precios"
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
         Height          =   375
         Left            =   6300
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   1215
         Width           =   1890
      End
      Begin FlexCell.Grid Grid1 
         Height          =   1500
         Left            =   135
         TabIndex        =   30
         Top             =   1665
         Width           =   8055
         _ExtentX        =   14208
         _ExtentY        =   2646
         Cols            =   5
         DefaultFontSize =   8.25
         Rows            =   30
      End
      Begin VB.CheckBox Sindescuento 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFF2F7&
         Caption         =   "Sin Descuento"
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
         Left            =   3600
         MaskColor       =   &H80000010&
         TabIndex        =   104
         Top             =   1350
         Width           =   1635
      End
      Begin VB.Label Label22 
         BackStyle       =   0  'Transparent
         Caption         =   "%"
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
         Left            =   3240
         TabIndex        =   92
         Top             =   1320
         Width           =   240
      End
      Begin VB.Label Label21 
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Descuento"
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
         Height          =   285
         Left            =   180
         TabIndex        =   91
         Top             =   1320
         Width           =   2295
      End
      Begin VB.Label lblpesable 
         BackStyle       =   0  'Transparent
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
         Left            =   4230
         TabIndex        =   84
         Top             =   1035
         Width           =   2040
      End
      Begin VB.Label Label11 
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Tipo de Venta"
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
         Height          =   285
         Left            =   180
         TabIndex        =   78
         Top             =   1030
         Width           =   2295
      End
      Begin VB.Label Label8 
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Precio Costo Con Iva"
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
         Height          =   285
         Left            =   180
         TabIndex        =   32
         Top             =   360
         Width           =   2295
      End
      Begin VB.Label Label9 
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Margen de Utilidad Publico"
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
         Height          =   285
         Left            =   180
         TabIndex        =   31
         Top             =   720
         Width           =   2295
      End
   End
   Begin XPFrame.FrameXp datospersonales 
      Height          =   1230
      Left            =   240
      TabIndex        =   24
      Top             =   0
      Width           =   14685
      _ExtentX        =   25903
      _ExtentY        =   2170
      BackColor       =   16761024
      Caption         =   "Datos"
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
      Begin VB.CommandButton cambia_codigo 
         BackColor       =   &H00FF8080&
         Caption         =   "Cambiar Codigo Barra"
         Height          =   375
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   103
         Top             =   720
         Visible         =   0   'False
         Width           =   1725
      End
      Begin VB.TextBox CODIGOINVEL 
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
         Left            =   3570
         Locked          =   -1  'True
         MaxLength       =   13
         TabIndex        =   87
         Tag             =   "codigobarra"
         Top             =   720
         Width           =   2250
      End
      Begin VB.CommandButton recupera 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Recuperar Codigo Disponible"
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
         Left            =   11025
         Style           =   1  'Graphical
         TabIndex        =   66
         Top             =   270
         Width           =   2685
      End
      Begin VB.TextBox DATO2 
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
         Left            =   4365
         MaxLength       =   70
         TabIndex        =   1
         Tag             =   "descripcion"
         Top             =   300
         Width           =   6435
      End
      Begin VB.TextBox dato1 
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
         Left            =   1320
         MaxLength       =   13
         TabIndex        =   0
         Tag             =   "codigobarra"
         Top             =   315
         Width           =   1575
      End
      Begin VB.Label lblenvase 
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
         ForeColor       =   &H00400000&
         Height          =   300
         Left            =   9480
         TabIndex        =   96
         Top             =   720
         Width           =   4935
      End
      Begin VB.Label Label24 
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Codigo Envase"
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
         Height          =   285
         Left            =   5880
         TabIndex        =   95
         Top             =   720
         Width           =   1575
      End
      Begin VB.Label Label19 
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Codigo Invel"
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
         Height          =   285
         Left            =   2130
         TabIndex        =   88
         Top             =   720
         Width           =   1365
      End
      Begin VB.Label ultimocodigo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
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
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   11610
         TabIndex        =   65
         Top             =   270
         Width           =   2580
      End
      Begin VB.Label Label2 
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Descripcion"
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
         Height          =   285
         Left            =   3060
         TabIndex        =   26
         Top             =   300
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Codigo barra"
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
         Height          =   285
         Left            =   120
         TabIndex        =   25
         Top             =   300
         Width           =   1095
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
      ScaleWidth      =   375
      TabIndex        =   22
      Top             =   240
      Width           =   375
   End
   Begin MSAdodcLib.Adodc mp 
      Height          =   375
      Left            =   13680
      Top             =   7020
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   1
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   3
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
   Begin FlexCell.Grid Grid3 
      Height          =   405
      Left            =   14130
      TabIndex        =   23
      Top             =   8325
      Visible         =   0   'False
      Width           =   870
      _ExtentX        =   1535
      _ExtentY        =   714
      Cols            =   6
      DefaultFontSize =   8.25
      Rows            =   30
   End
   Begin MSAdodcLib.Adodc data 
      Height          =   330
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   2143
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
   Begin XPFrame.FrameXp FrameXp1 
      Height          =   2895
      Left            =   6660
      TabIndex        =   71
      Top             =   5445
      Width           =   8385
      _ExtentX        =   14790
      _ExtentY        =   5106
      BackColor       =   8454016
      Caption         =   "Stock de Productos"
      CaptionEstilo3D =   1
      BackColor       =   8454016
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.CommandButton Command4 
         BackColor       =   &H00FF8080&
         Caption         =   "Ver todos  los Inventarios"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Left            =   6750
         Style           =   1  'Graphical
         TabIndex        =   112
         Top             =   2205
         Width           =   1500
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H0080C0FF&
         Caption         =   "MODIFICAR UBICACION"
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
         Left            =   180
         Style           =   1  'Graphical
         TabIndex        =   72
         Top             =   2160
         Width           =   2715
      End
      Begin FlexCell.Grid Grid2 
         Height          =   1755
         Left            =   90
         TabIndex        =   73
         Top             =   270
         Width           =   8235
         _ExtentX        =   14526
         _ExtentY        =   3096
         Cols            =   5
         DefaultFontSize =   8.25
         Rows            =   30
      End
      Begin VB.Label CFUI 
         Alignment       =   1  'Right Justify
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
         Height          =   330
         Left            =   5220
         TabIndex        =   110
         Top             =   2430
         Width           =   1500
      End
      Begin VB.Label FUI 
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
         Height          =   330
         Left            =   3150
         TabIndex        =   109
         Top             =   2430
         Width           =   2085
      End
      Begin VB.Label LABEL2000 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "CANTIDAD "
         Height          =   285
         Left            =   5220
         TabIndex        =   108
         Top             =   2160
         Width           =   1500
      End
      Begin VB.Label Label26 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "ULTIMO INVENTARIO"
         Height          =   285
         Left            =   3150
         TabIndex        =   107
         Top             =   2160
         Width           =   2085
      End
   End
   Begin VB.Label Label23 
      BackColor       =   &H00F5C9B1&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Codigo barra"
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
      Height          =   285
      Left            =   5760
      TabIndex        =   94
      Top             =   720
      Width           =   1095
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00808080&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FF8080&
      Height          =   3510
      Left            =   225
      Top             =   5040
      Width           =   6060
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash opciones 
      Height          =   1200
      Left            =   120
      TabIndex        =   20
      Top             =   8595
      Width           =   6780
      _cx             =   11959
      _cy             =   2117
      FlashVars       =   ""
      Movie           =   "c:\barra_opciones.swf"
      Src             =   "c:\barra_opciones.swf"
      WMode           =   "Transparent"
      Play            =   0   'False
      Loop            =   -1  'True
      Quality         =   "High"
      SAlign          =   ""
      Menu            =   -1  'True
      Base            =   ""
      AllowScriptAccess=   "always"
      Scale           =   "ShowAll"
      DeviceFont      =   0   'False
      EmbedMovie      =   0   'False
      BGColor         =   ""
      SWRemote        =   ""
      MovieData       =   ""
      SeamlessTabbing =   -1  'True
      Profile         =   0   'False
      ProfileAddress  =   ""
      ProfilePort     =   0
      AllowNetworking =   "all"
      AllowFullScreen =   "false"
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H00808080&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FF8080&
      Height          =   2970
      Left            =   6795
      Top             =   1575
      Width           =   8295
   End
   Begin VB.Shape Shape5 
      BackColor       =   &H00808080&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FF8080&
      Height          =   3390
      Left            =   180
      Top             =   1485
      Width           =   6060
   End
End
Attribute VB_Name = "preciosorden"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private PASO As String
Private bodegas(20) As String
Private PORCENTAJES(30) As Double
Private empresasrubro(20) As String
Private cer As Double
Private sc As Integer
Private MODIFI As Integer
Private apostrofe As String
Private digito As String
Private PRECIODIFERENCIADO As Boolean

Private Sub btn_cambioprecios_Click()
    If btn_cambioprecios.caption = "Cambiar Precios" Then
        btn_cambioprecios.caption = "Guardar Cambios"
        Grid1.Enabled = True
        Grid1.Column(7).Locked = False
        Grid1.Column(6).Locked = False
        Grid1.Cell(1, 7).SetFocus
    Else
        Check1.Value = 0
        Call grabarPrecios(dato1.text)
        Call leeprecios(dato1.text)
        
        btn_cambioprecios.caption = "Cambiar Precios"
        btn_cambioprecios.Enabled = False
        
        Grid1.Enabled = False
    End If
End Sub

Private Sub CAMBIA_CODIGO_BARRA_Click()
Call ceros(dato30)
lee_si_existe

End Sub

Private Sub cambia_codigo_Click()
If Verifica_Permiso(Me.caption, "autoriza") = True Then
Frmcodigobarra.Visible = True
dato30.SetFocus
Else
 MsgBox mensaje_nopermiso, vbCritical + vbOKOnly, "Permiso Denegado"
End If
End Sub

Private Sub Command6_Click()
TRASPASA.Show vbModal
End Sub

Private Sub checkpack_Click()
If checkpack.Value = "1" Then
    cmdpack.Visible = True
    End If
    
End Sub

Private Sub cmdpack_Click()
pack.Show vbModal

End Sub

Private Sub codigoenvase_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF2 Then
Call ayudaenvase(codigoenvase)
End If
End Sub

Private Sub codigoenvase_KeyPress(KeyAscii As Integer)
KeyAscii = esNumero(KeyAscii, "N")
If KeyAscii = 13 And codigoenvase.text <> "" Then
Call ceros(codigoenvase)
lblenvase.caption = leercodigoenvase(codigoenvase.text)
End If

End Sub



Private Sub Command2_Click()
  Dim resultados As rdoResultset
    Dim cSql As New rdoQuery
    
    Set cSql.ActiveConnection = GESTIONrubro
    cSql.sql = "SELECT * "
    cSql.sql = cSql.sql + "FROM r_maestroproductos_fijo_" & rubro & " where codigobarra like '0000078%' order by descripcion"
    cSql.Execute
    
    If cSql.RowsAffected > 0 Then
        Set resultados = cSql.OpenResultset
            
            While Not resultados.EOF
       
                Call modificacodigo(resultados(0), "0" + Mid(resultados(0), 1, 12))
                
                resultados.MoveNext
            Wend
            resultados.Close
        Set resultados = Nothing
        
    End If


End Sub
Private Sub modificacodigo(codigonuevo, codigoantiguo)
  Dim resultados As rdoResultset
    Dim cSql As New rdoQuery
    
    Set cSql.ActiveConnection = GESTIONrubro
    cSql.sql = "update l_movimientos_detalle_" + empresaactiva + " set codigo='" + codigonuevo + "'  where codigo='" + codigoantiguo + "' "
    
    cSql.Execute
    
End Sub


Private Sub Command1_Click()
If Command1.caption = "MODIFICAR UBICACION" Then
Command1.caption = "GRABAR UBICACION"
        Grid2.Enabled = True
        Grid2.Column(3).Locked = False
        Grid2.Cell(1, 3).SetFocus
Else
        Call grabarUBICACIONES(dato1.text)
        
        Command1.caption = "MODIFICAR UBICACION"
        Command1.Enabled = False
        Grid2.Enabled = False

End If

End Sub

Private Sub Command3_Click()
TRASPASA.Show

End Sub

Private Sub Command4_Click()
tomasInventario.dato1.text = dato1.text
tomasInventario.lblProducto = DATO2.text
tomasInventario.codigoToma = dato1.text




tomasInventario.Show vbModal

End Sub

Private Sub Command5_Click()
If sipack.Height = 180 Then
sipack.Height = 50
Else
sipack.Height = 180


End If

End Sub

Private Sub CompraVenta_Click()
    Load ultimos
    ultimos.codigoUltimo = dato1.text
    ultimos.Show vbModal
End Sub



Private Sub dato1_LostFocus()
'    Dim digi As String
'    If dato1.text <> "" And Len(Str(CDbl(dato1.text))) > 7 Then
'            digi = Validaciones.dv(Mid(dato1.text, 1, 12))
'            If Mid(dato1.text, 13, 1) = digi Then
            Call leer("=")
'            Else
'            MsgBox ("digito verificador erroneo el correcto para ese codigo es=" + digi + " revise")
'            dato1.SetFocus
'
'            End If
'    End If
'Call leer("=")
End Sub

Private Sub dato15_LostFocus()
    dato15.text = dato15.text + String(16 - Len(dato15.text), ".")

End Sub


Private Sub dato20_LostFocus()
FRMPESABLE.Visible = False

End Sub

Private Sub dato30_KeyPress(KeyAscii As Integer)
    KeyAscii = esNumero(KeyAscii, "N")
    If KeyAscii = 13 Then
    Call ceros(dato30)
    CAMBIA_CODIGO_BARRA.SetFocus
   End If
End Sub

Private Sub descuento_GotFocus()
Call cargatexto(descuento)
End Sub


Private Sub descuento_KeyPress(KeyAscii As Integer)
    KeyAscii = esNumero(KeyAscii, "N")
    If KeyAscii = 13 Then
    Call ceros(descuento)
    End If
End Sub

Private Sub finmodifica_Click()
If descuento.text = "" Then
descuento.text = 0
End If
If CDbl(descuento.text) <= 50 Then
codigoenvase_KeyPress (13)
GRABAR
retorno
finmodifica.Visible = False
Else

MsgBox "Atencion El Descuento No Puede Ser Mayor Que 50%", vbOKOnly, "Atencion"
descuento.SetFocus

End If


End Sub

Private Sub Form_Activate()

sqlgesti.audit = True
sqlgesti.programaactivo = Me.caption
Call dato1_KeyPress(13)


End Sub

Private Sub Form_Load()
    Me.Left = (Screen.Width - Me.Width) / 2
    Me.Top = (Screen.Height - Me.Height) / 2 - 1000
sipack.Visible = False

    sc = 0
    opciones.Visible = False
 Call CargaGrillaLista(1, 6)
    Call CARGAGRILLAPACK
    Call CARGAGRILLA
    Call empresasdelrubro
    Call CARGAGRILLAbodegas
    basebus = CLIENTE + "gestion" + rubro
    'Call ConectarControlData(mp, servidor, basedatos, USUARIO, password, "SELECT * from r_maestroproductos_fijo_" & rubro &" order by codigobarra ")
    btn_cambioprecios.caption = "Cambiar Precios"
finmodifica.Visible = False
Command1.Enabled = False
Check1.Value = 1
noactualiza.Value = 0
FRMPESABLE.Visible = False
sqlgesti.audit = True


End Sub

Private Sub codigo_Click()
    Call dato1_KeyDown(vbKeyF2, 0)
End Sub

Private Sub dato1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF2 Then Call ayudaProducto(dato1)
    Call flechas(dato1, DATO2, KeyCode)
    If KeyCode = 38 Then Unload maestro01 'Salida
    If KeyCode = 27 Then Unload maestro01 'Salida
End Sub

Private Sub dato1_GotFocus()
    
    Call cargatexto(dato1)
End Sub

Private Sub dato2_GotFocus()
            
           Call cargatexto(DATO2)
    
End Sub

Private Sub dato3_GotFocus()
     
           Call cargatexto(dato3)
        
End Sub

Private Sub dato4_GotFocus()
   
    Call cargatexto(dato4)
    If dato3.text <> "" Then Call leeSeccion
  
    
End Sub

Private Sub dato5_GotFocus()
 
    Call cargatexto(dato5)
  
    If dato4.text <> "" Then Call leeDepartamento
   
    
End Sub

Private Sub dato6_GotFocus()

    Call cargatexto(dato6)
    If dato5.text <> "" Then Call leeLinea
   
    
End Sub

Private Sub dato7_GotFocus()

    Call cargatexto(dato7)
    Call leeImpuesto
   
    
End Sub

Private Sub dato8_GotFocus()
    Call cargatexto(dato8)
    Call leeMarca
End Sub

Private Sub dato9_GotFocus()
    Call cargatexto(dato9)
    Call leeTemporada
End Sub

Private Sub dato10_GotFocus()
    dv.caption = rut(dato9)
    Call cargatexto(dato10)
    Call leeproveedor
End Sub

Private Sub dato11_GotFocus()
    Call cargatexto(dato11)
End Sub

Private Sub dato12_GotFocus()
    Call cargatexto(dato12)
End Sub

Private Sub dato13_GotFocus()
    Call cargatexto(dato13)
End Sub

Private Sub dato14_GotFocus()
    If Val(dato13.text) = 0 Then dato13.text = 1
    dato14.text = DATO2.text
    Call cargatexto(dato14)
End Sub

Private Sub dato15_GotFocus()
    Dim k As Double
    Dim cajas As String
    Dim CONTA As Double
    cajas = "": CONTA = 0
    For k = 1 To 50
    If CONTA > 4 And Mid(DATO2.text, k, 1) <> " " Then GoTo no:
    If Mid(DATO2.text, k, 1) = " " Then cajas = cajas + ".": CONTA = 0: GoTo no:
    cajas = cajas + Mid(DATO2.text, k, 1)
    CONTA = CONTA + 1
no:
    cajas = Mid(cajas, 1, 16)
    Next k
    
    dato15.text = cajas + String(16 - Len(cajas), ".")
    k = Len(dato15.text)
    Call cargatexto(dato15)
End Sub

Private Sub dato16_GotFocus()
    Call cargatexto(dato16)
End Sub

Private Sub dato17_GotFocus()
    Call cargatexto(dato17)
End Sub

Private Sub dato18_GotFocus()
    Call cargatexto(dato18)
End Sub

Private Sub dato19_GotFocus()
    Call cargatexto(dato19)
End Sub
Private Sub dato20_GotFocus()
 FRMPESABLE.Visible = True
    
    Call cargatexto(dato20)
End Sub
Private Sub dato2_KeyDown(KeyCode As Integer, Shift As Integer)
       Call flechas(dato1, dato3, KeyCode)
End Sub

Private Sub dato3_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF2 Then Call ayudaSeccion(dato4)
    Call flechas(DATO2, dato4, KeyCode)
End Sub

Private Sub dato4_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF2 Then Call ayudaDepto(dato5)
    Call flechas(dato3, dato5, KeyCode)
End Sub

Private Sub dato5_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF2 Then Call ayudaLinea(dato6)
    Call flechas(dato4, dato6, KeyCode)
End Sub

Private Sub dato6_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF2 Then Call ayudaImpuesto(dato7)
    Call flechas(dato5, dato7, KeyCode)
End Sub

Private Sub dato7_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyF2 Then Call ayudaMarca(dato8)
   Call flechas(dato6, dato8, KeyCode)
End Sub

Private Sub dato8_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF2 Then Call ayudaTemporada(dato9)
    Call flechas(dato7, dato9, KeyCode)
End Sub

Private Sub dato9_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF2 Then Call ayudaProveedor(dato10)
    Call flechas(dato8, dato10, KeyCode)
End Sub

Private Sub dato10_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF2 Then Call ayudaMedida(dato11)
    Call flechas(dato9, dato11, KeyCode)
End Sub

Private Sub dato11_KeyDown(KeyCode As Integer, Shift As Integer)
    Call flechas(dato10, dato12, KeyCode)
End Sub

Private Sub dato12_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF2 Then Call ayudaEmbalaje(dato13)
    Call flechas(dato11, dato13, KeyCode)
End Sub

Private Sub dato13_KeyDown(KeyCode As Integer, Shift As Integer)
    Call flechas(dato12, dato14, KeyCode)
End Sub

Private Sub dato14_KeyDown(KeyCode As Integer, Shift As Integer)
    Call flechas(dato13, dato15, KeyCode)
    'If KeyCode = 27 Then Unload maestro01
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

Private Sub dato18_KeyDown(KeyCode As Integer, Shift As Integer)
    Call flechas(dato17, dato19, KeyCode)
End Sub

Private Sub dato19_KeyDown(KeyCode As Integer, Shift As Integer)
    Call flechas(dato18, dato19, KeyCode)
End Sub

Private Sub dato1_KeyPress(KeyAscii As Integer)
    If empresaactiva <> "13" Then
    snum = 1: KeyAscii = esNumero(KeyAscii, "N")
    Else
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    
   End If
    
    If Len(dato1.text) = 12 And Chr(KeyAscii) = "0" And (Mid(dato1.text, 1, 7)) <> "0000000" Then
        KeyAscii = Asc(Validaciones.dv(Mid(dato1.text, 1, 12)))
        End If
    
    
    If KeyAscii = 13 Then
    If empresaactiva <> "13" Then
    Call ceros(dato1)
    End If
    Call Pregunta(dato1, DATO2)
    End If
    
End Sub

Private Sub dato2_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then sc = 0: Call Pregunta(DATO2, dato3)
End Sub

Private Sub dato3_keypress(KeyAscii As Integer)
    snum = 1: KeyAscii = esNumero(KeyAscii, "N")
    If KeyAscii = 13 Then Call ceros(dato3): Call Pregunta(dato3, dato4)
End Sub

Private Sub dato4_KeyPress(KeyAscii As Integer)
    snum = 1: KeyAscii = esNumero(KeyAscii, "N")
    If KeyAscii = 13 Then Call ceros(dato4): Call Pregunta(dato4, dato5)
End Sub

Private Sub dato5_KeyPress(KeyAscii As Integer)
    snum = 1: KeyAscii = esNumero(KeyAscii, "N")
    If KeyAscii = 13 Then Call ceros(dato5): Call Pregunta(dato5, dato6)
End Sub

Private Sub dato6_KeyPress(KeyAscii As Integer)
    snum = 1: KeyAscii = esNumero(KeyAscii, "N")
    If KeyAscii = 13 Then Call ceros(dato6): Call Pregunta(dato6, dato7)
End Sub

Private Sub dato7_KeyPress(KeyAscii As Integer)
    snum = 1: KeyAscii = esNumero(KeyAscii, "N")
    If KeyAscii = 13 Then Call ceros(dato7): Call Pregunta(dato7, dato8)
End Sub

Private Sub dato8_KeyPress(KeyAscii As Integer)
    KeyAscii = esNumero(KeyAscii, "N")
    If KeyAscii = 13 Then Call ceros(dato8): Call Pregunta(dato8, dato9)
End Sub

Private Sub dato9_KeyPress(KeyAscii As Integer)
    KeyAscii = esNumero(KeyAscii, "N")
    If KeyAscii = 13 Then Call ceros(dato9): Call Pregunta(dato9, dato10)
End Sub

Private Sub dato10_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 13 Then Call Pregunta(dato10, dato11)
End Sub

Private Sub dato11_KeyPress(KeyAscii As Integer)
    KeyAscii = esNumero(KeyAscii, "S")
    If KeyAscii = 13 Then Call FORMATO(dato11, 3): Call Pregunta(dato11, dato12)
End Sub

Private Sub dato12_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then Call Pregunta(dato12, dato13)
End Sub

Private Sub dato13_KeyPress(KeyAscii As Integer)
    KeyAscii = esNumero(KeyAscii, "N")
    If KeyAscii = 13 And Val(dato13.text) <> 0 Then Call FORMATO(dato13, 0): Call Pregunta(dato13, dato14)
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
    If KeyAscii = 13 Then sc = 1: Call Pregunta(dato16, dato17)
End Sub

Private Sub dato17_KeyPress(KeyAscii As Integer)
    KeyAscii = esNumero(KeyAscii, "N")
    If KeyAscii = 13 Then Call ceros(dato17): Call Pregunta(dato17, dato18)
End Sub

Private Sub dato18_KeyPress(KeyAscii As Integer)
    KeyAscii = esNumero(KeyAscii, "S")
    If KeyAscii = 13 Then Call FORMATO(dato18, 1): Call Pregunta(dato18, dato19)
End Sub

Private Sub dato19_KeyPress(KeyAscii As Integer)
    KeyAscii = esNumero(KeyAscii, "S")
    If KeyAscii = 13 Then Call Pregunta(dato19, dato20)
    ' If KeyAscii = 13 And Val(dato19.text) <> 0 Then Call calculaPrecios: Call GRABAR: Call leer("=")
    End Sub
Private Sub dato20_KeyPress(KeyAscii As Integer)
    KeyAscii = esNumero(KeyAscii, "N")
    If KeyAscii = 13 And Val(dato20.text) < 4 Then
    Call calculaPrecios: Call GRABAR: Call leer("=")
    FRMPESABLE.Visible = False
    End If
    
    End Sub
    
Sub calculaPrecios()
'    Dim PORCE As Double
'    Dim MARGENDEFINITIVO As Double
'    Dim precioventa As Double
'
'    If Check1.Value = 1 Then
'
'    For k = 1 To Grid1.Rows - 1
'        Grid1.Cell(k, 3).text = PORCENTAJES(k)
'        MARGENDEFINITIVO = dato19.text * PORCENTAJES(k) / 100
'        Grid1.Cell(k, 4).text = MARGENDEFINITIVO
'        Grid1.Cell(k, 5).text = Int((dato18.text * (1 + (Grid1.Cell(k, 4).text) / 100)) + 0.5)
'        'If modifi = 1 Then GoTo paso:  ''''' DESCOMENTAR
'        Grid1.Cell(k, 6).text = Int((Grid1.Cell(k, 5).text) + 0.5)
'        Grid1.Cell(k, 7).text = Grid1.Cell(k, 5).text
'        PORCE = Grid1.Cell(k, 5).text / dato18.text
'        'Grid1.Cell(k, 8).text = (PORCE - 1) * 100
'
'    Next k
'    Else
'    If MsgBox("Atencion el Producto se Creara con Precio 0 ¿ Desea Generar Precios Automaticos?", vbOKCancel, "Atencion") = vbOK Then
'    Check1.Value = 1
'    calculaPrecios
'    End If
'End If
End Sub

Sub leer(Opcion)
    campos(0, 0) = dato1.Tag
    campos(1, 0) = DATO2.Tag
    campos(2, 0) = dato3.Tag
    campos(3, 0) = dato4.Tag
    campos(4, 0) = dato5.Tag
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
    campos(19, 0) = "descontinuado"
    campos(20, 0) = "imprimefleje"
    campos(21, 0) = "noactualiza"
    campos(22, 0) = "pesable"
    campos(23, 0) = "fechacreacion"
    campos(24, 0) = "correlativoinvel"
    campos(25, 0) = "descuento"
    campos(26, 0) = "envase"
    campos(27, 0) = "codigoenvase"
    campos(28, 0) = "nodescuento"
    campos(29, 0) = "pack"
    campos(30, 0) = ""
    campos(0, 2) = "r_maestroproductos_fijo_" & rubro
    If Opcion = "=" Then condicion = "codigobarra = '" & dato1.text & "'  ORDER BY codigobarra limit 0,1"
    If Opcion = "-" Then condicion = "codigobarra < '" & dato1.text & "'  ORDER BY codigobarra DESC limit 0,1"
    If Opcion = "+" Then condicion = "codigobarra > '" & dato1.text & "'  ORDER BY codigobarra ASC limit 0,1"
    op = 5
    sqlgesti.response = campos
    Set sqlgesti.conexion = GESTIONrubro
    Call sqlgesti.sqlgesti(op, condicion)
    If sqlgesti.status = 0 Then
        Call carga
        MODIFI = 1
        opciones.Visible = True
        Call disponible(True)
        Call habilita(True)
        cambia_codigo.Visible = True
        opciones.SetFocus
    Rem Call GRABAR3(dato1.text, "00", empresaactiva)
    Else
        If Verifica_Permiso(Me.caption, "agrega") = True Then
            If Opcion = "=" Then
                CODIGOINVEL.text = leerultimocodigoinvel
                
                maestro01.DATO2.SetFocus
            Else
                'opciones.SetFocus
            End If
        Else
            MsgBox mensaje_nopermiso, vbCritical + vbOKOnly, "Permiso Denegado"
            dato1.SelStart = 0
            dato1.SelLength = Len(dato1.text)
            dato1.SetFocus
        End If
    End If
End Sub

Sub carga()
    Dim margen As Double
    
    habilita (True)
    dato1.text = sqlgesti.response(0, 3)
    DATO2.text = sqlgesti.response(1, 3)
    dato3.text = sqlgesti.response(2, 3)
    dato4.text = sqlgesti.response(3, 3)
    dato5.text = sqlgesti.response(4, 3)
    dato6.text = sqlgesti.response(5, 3)
    dato7.text = sqlgesti.response(6, 3)
    dato8.text = sqlgesti.response(7, 3)
    dato9.text = Mid(sqlgesti.response(8, 3), 1, 9)
    dv.caption = Mid(sqlgesti.response(8, 3), 10, 1)
    
    dato10.text = sqlgesti.response(9, 3)
    dato11.text = sqlgesti.response(10, 3)
    dato12.text = sqlgesti.response(11, 3)
    dato13.text = sqlgesti.response(12, 3)
    dato14.text = sqlgesti.response(13, 3)
    dato15.text = sqlgesti.response(14, 3)
    dato16.text = sqlgesti.response(15, 3)
    dato17.text = sqlgesti.response(16, 3)
    dato18.text = Format(sqlgesti.response(17, 3), "###,###,##0.00")
    dato19.text = Format(sqlgesti.response(18, 3), "###,0.00")
    FECHACREACION.caption = "CREADO:" + Format(sqlgesti.response(23, 3), "dd-mm-yyyy")
    
    descontinuado.Value = sqlgesti.response(19, 3)
    
    Check7.Value = Val(sqlgesti.response(20, 3))
    noactualiza.Value = sqlgesti.response(21, 3)
    dato20.text = sqlgesti.response(22, 3)
    If dato20.text = "0" Then lblpesable.caption = Label14.caption
    If dato20.text = "1" Then lblpesable.caption = Label15.caption
    If dato20.text = "2" Then lblpesable.caption = Label17.caption
    If dato20.text = "3" Then lblpesable.caption = Label18.caption
    CODIGOINVEL.text = sqlgesti.response(24, 3)
    descuento.text = sqlgesti.response(25, 3)
    envase.Value = sqlgesti.response(26, 3)
    codigoenvase.text = sqlgesti.response(27, 3)
    Sindescuento.Value = sqlgesti.response(28, 3)
    cmdpack.Visible = False
    
    checkpack.Value = Val(sqlgesti.response(29, 3))
    If checkpack.Value = "1" Then
    cmdpack.Visible = True
    End If
    
    lblenvase.caption = leercodigoenvase(codigoenvase.text)
    Call leeSeccion
    Call leeDepartamento
    Call leeLinea
    Call leeImpuesto
    Call leeMarca
    Call leeTemporada
    Call leeproveedor
    Call CARGASTOCKBODEGAS
    Call leeprecios(dato1)
    Call leerEspeciales
    Call LEErultimoinventario(dato1.text, empresaactiva)
    FUI.caption = Format(fechaui, "dd-mm-yyyy")
    CFUI.caption = Format(cantidadui, "###,##0.0")
    Call leerSOYPACK(dato1.text)
    
    
End Sub
Sub leerSOYPACK(codigo)
    Dim resultados As rdoResultset
    Dim cSql As New rdoQuery
    Dim rut As String
    Dim linea As Double
    
        Set cSql.ActiveConnection = GESTIONrubro
        cSql.sql = "SELECT codigopack,codigobarra,cantidad "
        cSql.sql = cSql.sql + "FROM r_pack_detalle_" + rubro + " "
        cSql.sql = cSql.sql + "where codigobarra='" + codigo + "' "
        cSql.Execute
        linea = 0
        grillapack.Rows = 1
   
        If cSql.RowsAffected > 0 Then
            Set resultados = cSql.OpenResultset
                sipack.Visible = True
                
            While Not resultados.EOF
                
                grillapack.Rows = grillapack.Rows + 1
                linea = linea + 1
                grillapack.Cell(linea, 1).text = resultados(0)
                grillapack.Cell(linea, 2).text = leerNombreProducto(resultados(0))
                grillapack.Cell(linea, 3).text = leerPrecioProducto(resultados(0))
                
                grillapack.Cell(linea, 4).text = resultados(2)
                
                resultados.MoveNext
            
            Wend
            resultados.Close
        Set resultados = Nothing

        End If
End Sub


Sub habilita(ByVal condicion As Boolean)
    dato1.Locked = condicion
    DATO2.Locked = condicion
    dato3.Locked = condicion
    dato4.Locked = condicion
    dato5.Locked = condicion
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
    dato20.Locked = condicion
    descuento.Locked = condicion
End Sub

Sub disponible(ByVal condicion As Boolean)
    dato1.Enabled = condicion
    DATO2.Enabled = condicion
    dato3.Enabled = condicion
    dato4.Enabled = condicion
    dato5.Enabled = condicion
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
    dato16.Enabled = condicion
    dato17.Enabled = condicion
    dato18.Enabled = condicion
    dato19.Enabled = condicion
    dato20.Enabled = condicion
    descuento.Enabled = condicion
    

End Sub



Sub Pregunta(ByRef caja As TextBox, ByRef caja1 As TextBox)
    If caja.text = "" And sc = 0 Then caja.SetFocus
    If caja.text <> "" Or sc = 1 Then caja1.Enabled = True: caja1.SetFocus
    sc = 0
End Sub

Sub ayudaProveedor(ByRef caja As TextBox)
    Dim campos As Variant
    Dim cfijo As Variant
    Dim largo As Variant
    campos = Array("rut", "nombre")
    largo = Array("10n", "60s")
    cfijo = "no"
    mensajeAyuda = "Ayuda proveedores"
    cabezas = Array("rut", "nombre")
    Call cargaAyudaT(servidor, CLIENTE + "gestion" + rubro, usuario, password, "r_maestroproveedores_" + rubro, dato9, campos, cfijo, largo, 2)
    caja.Enabled = True
    caja.SetFocus
End Sub
Sub ayudaenvase(ByRef caja As TextBox)
    Dim campos As Variant
    Dim cfijo As Variant
    Dim largo As Variant
    campos = Array("codigobarra", "descripcion")
    largo = Array("10n", "60s")
    cfijo = "envase='1'"
    mensajeAyuda = "Ayuda Envase"
    cabezas = Array("Codigo", "Descripcion")
    Call cargaAyudaT(servidor, CLIENTE + "gestion" + rubro, usuario, password, "r_maestroproductos_fijo_" + rubro, codigoenvase, campos, cfijo, largo, 2)
    caja.Enabled = True
    caja.SetFocus
End Sub


Sub ayudaSeccion(ByRef caja As TextBox)
    Dim campos As Variant
    Dim cfijo As Variant
    Dim largo As Variant
    campos = Array("codigo", "nombre")
    largo = Array("6n", "30s")
    cfijo = "rubro = '" & rubro & "'"
    mensajeAyuda = "Ayuda Secciones"
    cabezas = Array("codigo", "nombre")
    Call cargaAyudaT(servidor, CLIENTE + "gestion" & rubro, usuario, password, "r_maestrosecciones_" & rubro, dato3, campos, cfijo, largo, 2)
    caja.Enabled = True
    caja.SetFocus
    End Sub

Sub ayudaImpuesto(ByRef caja As TextBox)
    Dim campos As Variant
    Dim cfijo As Variant
    Dim largo As Variant
    campos = Array("codigo", "nombre")
    largo = Array("6n", "60s")
    cfijo = "no"
    
    mensajeAyuda = "Ayuda Impuestos"
    cabezas = Array("codigo", "nombre")
    
    Call cargaAyudaT(servidor, CLIENTE + "gestion", usuario, password, "g_maestroimpuestos", dato6, campos, cfijo, largo, 2)
    caja.Enabled = True
    caja.SetFocus
End Sub

Sub ayudaMedida(ByRef caja As TextBox)
    Dim campos As Variant
    Dim cfijo As Variant
    Dim largo As Variant
    campos = Array("unidad", "descripcion")
    largo = Array("10n", "30s")
    cfijo = "no"
    mensajeAyuda = "Ayuda Unidades"
    cabezas = Array("Unidad", "Detalle")
    Call cargaAyudaT(servidor, CLIENTE + "gestion", usuario, password, "g_maestromedidas", dato10, campos, cfijo, largo, 2)
    caja.Enabled = True
    caja.SetFocus
End Sub

Sub ayudaEmbalaje(ByRef caja As TextBox)
    Dim campos As Variant
    Dim cfijo As Variant
    Dim largo As Variant
    campos = Array("nombre", "material")
    largo = Array("10n", "30s")
    cfijo = "no"
    mensajeAyuda = "Ayuda Embalajes"
    cabezas = Array("Nombre", "Detalle")
    Call cargaAyudaT(servidor, CLIENTE + "gestion", usuario, password, "g_maestroembalajes", dato12, campos, cfijo, largo, 2)
    caja.Enabled = True
    caja.SetFocus
End Sub

Sub ayudaDepto(ByRef caja As TextBox)
    Dim campos As Variant
    Dim cfijo As Variant
    Dim largo As Variant
    campos = Array("codigodepto", "nombre")
    largo = Array("6n", "50s")
    cfijo = "codigoseccion='" + dato3.text + "'"

    mensajeAyuda = "Ayuda Departamentos"
    cabezas = Array("codigo", "nombre")
    
    Call cargaAyudaT(servidor, CLIENTE + "gestion" & rubro, usuario, password, "r_maestrodepartamentos_" & rubro, dato4, campos, cfijo, largo, 2)
    caja.Enabled = True
    caja.SetFocus
   
End Sub

Sub ayudaLinea(ByRef caja As TextBox)
    Dim campos As Variant
    Dim cfijo As Variant
    Dim largo As Variant
    campos = Array("codigolinea", "nombre")
    largo = Array("6n", "50s")
        
    cfijo = "codigoseccion='" + dato3.text + "' and codigodepto = '" & dato4.text & "'"

    mensajeAyuda = "Ayuda Lineas"
    cabezas = Array("codigo", "nombre")
    
    Call cargaAyudaT(servidor, CLIENTE + "gestion" & rubro, usuario, password, "r_maestrolineas_" & rubro, dato5, campos, cfijo, largo, 2)
    caja.Enabled = True
    caja.SetFocus
    
End Sub

Sub ayudaMarca(ByRef caja As TextBox)
    Dim campos As Variant
    Dim cfijo As Variant
    Dim largo As Variant
    campos = Array("codigo", "nombre")
    largo = Array("6n", "50s")
    cfijo = "rubro = '" & rubro & "'"
    mensajeAyuda = "Ayuda Marcas"
    cabezas = Array("codigo", "nombre")
    
    Call cargaAyudaT(servidor, CLIENTE + "gestion" & rubro, usuario, password, "r_maestromarcas_" & rubro, dato7, campos, cfijo, largo, 2)
    caja.Enabled = True
    caja.SetFocus
End Sub

Sub ayudaTemporada(ByRef caja As TextBox)
    Dim campos As Variant
    Dim cfijo As Variant
    Dim largo As Variant
    campos = Array("codigo", "nombre")
    largo = Array("6n", "50s")
    cfijo = "no"
    mensajeAyuda = "Ayuda temporadas"
    cabezas = Array("codigo", "nombre")
    
    Call cargaAyudaT(servidor, CLIENTE + "gestion" & rubro, usuario, password, "r_maestrotemporadas_" & rubro, dato8, campos, cfijo, largo, 2)
    caja.Enabled = True
    caja.SetFocus
End Sub

Sub ayudaProducto(ByRef caja As TextBox)
    Dim campos As Variant
    Dim cfijo As Variant
    Dim largo As Variant
    mensajeAyuda = "Ayuda Maestro de Productos - Código Producto"
    cabezas = Array("codigo producto", "descripcion", "precio venta")
    campos = Array("mpf.codigobarra", "mpf.descripcion", "mpp.preciopuntoventa")
    largo = Array("15n", "50s", "15n")
    cfijo = "mpf.codigobarra=mpp.codigo and mpp.local= " & empresaactiva & " and mpf.descontinuado='0' "
    Call cargaAyudaT(servidor, basedatos & rubro, usuario, password, "r_maestroproductos_fijo_" & rubro & " as mpf, r_maestroproductos_precios_" & rubro & " as mpp", dato1, campos, cfijo, largo, 3)

    caja.Enabled = True
    caja.SetFocus
End Sub

Sub flechas(ByRef caja As TextBox, ByRef caja1 As TextBox, ByRef codigo As Integer)
    If codigo = 38 And caja.Enabled = True Then caja.SetFocus
    If codigo = 40 And caja1.Enabled = True Then caja1.SetFocus
End Sub

Sub GRABAR()
If descuento.text = "" Then descuento.text = "0"
If CDbl(descuento.text) <= 50 Then
    campos(0, 0) = dato1.Tag
    campos(1, 0) = DATO2.Tag
    campos(2, 0) = dato3.Tag
    campos(3, 0) = dato4.Tag
    campos(4, 0) = dato5.Tag
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
    campos(19, 0) = "descontinuado"
    campos(20, 0) = "imprimefleje"
    campos(21, 0) = "noactualiza"
    campos(22, 0) = "pesable"
    campos(23, 0) = "correlativoinvel"
    campos(24, 0) = "descuento"
    campos(25, 0) = "envase"
    campos(26, 0) = "codigoenvase"
    campos(27, 0) = "nodescuento"
    campos(28, 0) = "pack"
    campos(29, 0) = "fechacreacion"
    If MODIFI = 1 Then campos(29, 0) = ""
    campos(30, 0) = ""
    
    
    campos(0, 1) = dato1.text
    campos(1, 1) = DATO2.text
    campos(2, 1) = dato3.text
    campos(3, 1) = dato4.text
    campos(4, 1) = dato5.text
    campos(5, 1) = dato6.text
    campos(6, 1) = dato7.text
    campos(7, 1) = dato8.text
    campos(8, 1) = dato9.text + dv.caption
    campos(9, 1) = dato10.text
    campos(10, 1) = Replace(dato11.text, ",", ".")
    campos(11, 1) = dato12.text
    campos(12, 1) = dato13.text
    campos(13, 1) = dato14.text
    campos(14, 1) = dato15.text
    campos(15, 1) = dato16.text
    campos(16, 1) = dato17.text
    dato18.text = Replace(dato18.text, ".", "")
    campos(17, 1) = Replace(dato18.text, ",", ".")
    campos(18, 1) = Replace(dato19.text, ",", ".")
    campos(19, 1) = descontinuado.Value
    campos(20, 1) = Check7.Value
    campos(21, 1) = noactualiza.Value
    campos(22, 1) = dato20.text
    campos(23, 1) = CODIGOINVEL.text
    campos(24, 1) = descuento.text
    campos(25, 1) = envase.Value
    If lblenvase.caption <> "" Then
    campos(26, 1) = codigoenvase.text
    Else
    campos(26, 1) = ""
    End If
    
    campos(27, 1) = Sindescuento.Value
    
    campos(28, 1) = checkpack.Value
    campos(29, 1) = Format(Date, "yyyy-mm-dd")
   
    campos(0, 2) = "r_maestroproductos_fijo_" & rubro
    If MODIFI = 1 Then condicion = "codigobarra= '" + dato1.text + "'"
    If MODIFI = 1 Then op = 3 Else op = 2
    sqlgesti.response = campos
    Set sqlgesti.conexion = GESTIONrubro
    Call sqlgesti.sqlgesti(op, condicion)
 
    If MODIFI = 0 Then
    Call grabaBodegas(dato1.text)
    Call grabarPrecios(dato1.text)
    Call calculaPrecios
    End If
    Check1.Value = 0
    
    
    MODIFI = 0
    finmodifica.Visible = False
    Else
    MsgBox "El Descuento No Puede Ser Mayor Que 50%", vbOKOnly, "Atencion"
    descuento.SetFocus
    End If
    
    
End Sub

Sub GRABAR2(codigo, bodega, locales)
    campos(0, 0) = "local"
    campos(1, 0) = "codigo"
    campos(2, 0) = "bodega"
    campos(3, 0) = "año"
    campos(4, 0) = ""
    
    campos(0, 1) = locales
    campos(1, 1) = codigo
    campos(2, 1) = bodega
    campos(3, 1) = "2007"
   
    campos(0, 2) = "r_maestroproductos_stock_" & rubro
    op = 2
    sqlgesti.response = campos
    Set sqlgesti.conexion = GESTIONrubro
    Call sqlgesti.sqlgesti(op, condicion)
End Sub

Sub GRABAR3(codigo, bodega, locales)
    campos(0, 0) = "local"
    campos(1, 0) = "codigo"
    campos(2, 0) = "año"
    campos(3, 0) = ""
    campos(0, 1) = locales
    campos(1, 1) = codigo
    campos(2, 1) = Mid(fechasistema, 7, 4)
   
    campos(0, 2) = "r_maestroproductos_estadistica_" & rubro
    op = 2
    sqlgesti.response = campos
    Set sqlgesti.conexion = GESTIONrubro
    Call sqlgesti.sqlgesti(op, condicion)

End Sub

Sub ELIMINAR()
'    If Verifica_Permiso(Me.caption, "elimina") = True Then
'        'MAESTRO PRODUCTOS FIJO
'        campos(0, 2) = "r_maestroproductos_fijo_" & rubro
'        condicion = "codigobarra = '" & dato1.text & "'"
'        op = 4
'        sqlgesti.response = campos
'        Set sqlgesti.conexion = GESTIONrubro
'        Call sqlgesti.sqlgesti(op, condicion)
'        op = sqlgesti.status
'
'        'MAESTRO PRODUCTOS ESTADISTICA
'        campos(0, 2) = "r_maestroproductos_estadistica_" & rubro
'        condicion = "codigo = '" & dato1.text & "'"
'        op = 4
'        sqlgesti.response = campos
'        Set sqlgesti.conexion = GESTIONrubro
'        Call sqlgesti.sqlgesti(op, condicion)
'        op = sqlgesti.status
'
'        'MAESTRO PRODUCTOS STOCK
'        campos(0, 2) = "r_maestroproductos_stock_" & rubro
'        condicion = "codigo = '" & dato1.text & "'"
'        op = 4
'        sqlgesti.response = campos
'        Set sqlgesti.conexion = GESTIONrubro
'        Call sqlgesti.sqlgesti(op, condicion)
'        op = sqlgesti.status
'        Call eliminaPrecios
'    Else
'        MsgBox mensaje_noelimina, vbCritical + vbOKOnly, "Permiso Denegado"
'        dato1.SelStart = 0
'        dato1.SelLength = Len(dato1.text)
'        dato1.SetFocus
'    End If
End Sub

Sub eliminaPrecios()
    campos(0, 2) = "r_maestroproductos_precios_" & rubro
    condicion = "codigo = '" & dato1.text & "'"
    op = 4
    sqlgesti.response = campos
    Set sqlgesti.conexion = GESTIONrubro
    Call sqlgesti.sqlgesti(op, condicion)
End Sub

    Private Sub frmLista_BarMouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
        Call cambiaColor(frmLista)
        frmLista.CaptionEstilo3D = Raised
    End Sub

    Private Sub frmLista_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
        Call cambiaColor(frmLista)
        frmLista.CaptionEstilo3D = Inserted
        Load ofertas
        Especiales.dato1.text = dato1.text
        Call ofertas.cargaLista
        ofertas.Show vbModal
    End Sub

Private Sub Grid1_LeaveCell(ByVal Row As Long, ByVal Col As Long, NewRow As Long, NewCol As Long, Cancel As Boolean)
If NewRow = 8 Then NewRow = 7
If NewRow = 5 Then NewRow = 6
End Sub

Private Sub historico_Click()
     codigocambioprecio = dato1.text
     cambiosdeprecio.Show vbModal

End Sub

Private Sub opciones_FSCommand(ByVal command As String, ByVal args As String)
    If command = "retorno" Then
    retorno
    End If
    
    If command = "modifica" Then
        If Verifica_Permiso(Me.caption, "modifica") = True Then
            opciones.Visible = False
            finmodifica.Visible = True
            
            disponible (True)
            habilita (False)
            dato1.Enabled = False
            btn_cambioprecios.Enabled = True
            DATO2.SetFocus
            Command1.Enabled = True
            MODIFI = 1
        Else
            MsgBox mensaje_nopermiso, vbCritical + vbOKOnly, "Permiso Denegado"
        End If
    End If
    If command = "elimina" Then
            If Verifica_Permiso(Me.caption, "elimina") = True Then
      
                If MsgBox("REALMENTE DESEA ELIMINAR", vbYesNo) = vbYes Then
        
                    If Consulta_Movimiento(dato1.text) = False Then 'NO PRESENTA MOVIMIENTOS
                    Call disponible(True)
                    Call habilita(False)
                    Call ELIMINAR
                    Call limpia
                    opciones.Visible = False
                    dato1.SetFocus
                    End If
                End If
            Else
            MsgBox mensaje_nopermiso, vbCritical + vbOKOnly, "Permiso Denegado"
        End If
    End If
    If command = "siguiente" Then Call leer("+")
    If command = "anterior" Then Call leer("-")
    If command = "imprime" Then Call IMPRIMIR
    If command = "historico" Then
       estadisticas.codigo.caption = dato1.text
       estadisticas.descripcion.caption = DATO2.text
       estadisticas.UXC.caption = dato13.text
       estadisticas.Show vbModal
       
       
    End If
    'If command = "movimientos" Then
    '    Load tomas
    '    tomas.codigoToma = dato1.text
    '    tomas.Show vbModal
    'End If
End Sub
Sub retorno()
disponible (True)
habilita (False)
limpia
finmodifica.Visible = False
opciones.Visible = False
Command1.Enabled = False
descontinuado.Value = False
cambia_codigo.Visible = False
Sindescuento.Value = 0
checkpack.Value = 0
cmdpack.Visible = False
sipack.Visible = False

MODIFI = 0
Unload Me


End Sub
'================================== CAMBIAR
Function Consulta_Movimiento(codigo As String) As Boolean
    Dim resultados As rdoResultset
    Dim cSql As New rdoQuery
    
    Set cSql.ActiveConnection = GESTIONrubro
    cSql.sql = "SELECT SUM(cantidadultimaventa) "
    cSql.sql = cSql.sql + "FROM r_maestroproductos_estadistica_" & rubro & " "
    cSql.sql = cSql.sql + "WHERE codigo='" & codigo & "' and local='" + empresaactiva + "' "
    cSql.Execute
    
    If cSql.RowsAffected > 0 Then
        Set resultados = cSql.OpenResultset
            If Not IsNull(resultados(0).Value) And (resultados(0).Value) > 0 Then
                Consulta_Movimiento = True
                resultados.Close
                Set resultados = Nothing
            Else
                Consulta_Movimiento = False
                resultados.Close
                Set resultados = Nothing
            End If
    End If

End Function

Sub limpia()
    dv.caption = ""
    dato1.text = ""
    DATO2.text = ""
    dato3.text = ""
    dato4.text = ""
    dato5.text = ""
    dato6.text = ""
    dato7.text = ""
    dato8.text = ""
    dato9.text = ""
    dato10.text = ""
    dato11.text = "0"
    dato12.text = ""
    dato13.text = "1"
    dato14.text = ""
    dato15.text = ""
    dato16.text = ""
    dato17.text = ""
    dato18.text = ""
    dato19.text = ""
    dato20.text = "0"
    CODIGOINVEL.text = ""
    nombreseccion.caption = ""
    nombremarca.caption = ""
    nombretemporada.caption = ""
    nombredepto.caption = ""
    nombrelinea.caption = ""
    nombreimpUESTO.caption = ""
    nombreproveedor.caption = ""
    Grid2.Rows = 1
    envase.Value = 0
    dsctodpto.caption = ""
    dsctolinea.caption = ""
    descuento.text = "0"
    codigoenvase = ""
    lblenvase.caption = ""
    FECHACREACION.caption = ""
    
    Check1.Value = 1
    
    For k = 1 To Grid1.Rows - 1
        Grid1.Cell(k, 4).text = ""
        Grid1.Cell(k, 5).text = ""
        Grid1.Cell(k, 6).text = ""
        Grid1.Cell(k, 7).text = ""
     
    Next k
    MODIFI = 0
    lblpesable.caption = ""
End Sub

Sub IMPRIMIR()

End Sub

'Sub grilla()
'    If lineas > LARGOPAGINA Then Call cabeza
'    PALABRA = ""
'    For k = 1 To cancolu
'    If tipodato(k) = "s" Or tipodato(k) = "S" Then dato(k) = dato(k) & String(colu(k) - Len(dato(k)), 32)
'    If tipodato(k) = "n" Or tipodato(k) = "N" Then dato(k) = String(colu(k) - Len(dato(k)), 32) & dato(k)
'    PALABRA = PALABRA & dato(k)
'    Next k
'
'    info.AddItem (PALABRA)
'    lineas = lineas + 1
'End Sub

'Sub cabeza()
'    informes.info.AddItem ("")
'    informes.info.AddItem ("")
'    pagina = pagina + 1
'    informes.info.AddItem ("NOMBRE EMPRESA                                                                                 PAGINA " + Str$(pagina))
'    informes.info.AddItem ("DIRECCION EMPRESA                                                                              ")
'    informes.info.AddItem ("                                " + tituloinforme)
'    informes.info.AddItem String(132, "=")
'    PALABRA = ""
'    For k = 1 To cancolu
'    titu(k) = titu(k) & String(colu(k) - Len(titu(k)), 32)
'    PALABRA = PALABRA & titu(k)
'    Next k
'    informes.info.AddItem (PALABRA)
'    informes.info.AddItem String(132, "=")
'lineas = 8
'End Sub

Sub grabaBodegas(codigo)
    Dim a As Integer
    Dim resultados As rdoResultset
    Dim cSql As New rdoQuery
    Dim rut As String
   
    For a = 1 To cer
        Set cSql.ActiveConnection = GESTIONrubro
        cSql.sql = "SELECT local,codigobodega "
        cSql.sql = cSql.sql + "FROM r_maestrobodegas_" & rubro & " "
        cSql.sql = cSql.sql + "WHERE local='" + empresasrubro(a) + "' ORDER BY codigobodega"
        cSql.Execute
      
        If cSql.RowsAffected > 0 Then
            Set resultados = cSql.OpenResultset
            While Not resultados.EOF
                Call GRABAR2(codigo, resultados(1), empresasrubro(a))
                resultados.MoveNext
            Wend
            resultados.Close
        Set resultados = Nothing

        End If
        Call GRABAR3(codigo, "00", empresasrubro(a))
    Next a

End Sub

Sub leeSeccion()
    campos(0, 0) = "codigo"
    campos(1, 0) = "nombre"
    campos(2, 0) = ""
    campos(0, 2) = "r_maestrosecciones_" & rubro
    condicion = "rubro = '" & rubro & "' AND codigo = '" & dato3.text & "' ORDER BY codigo,rubro LIMIT 0,1"
    op = 5
    sqlgesti.response = campos
    Set sqlgesti.conexion = GESTIONrubro
    
    Call sqlgesti.sqlgesti(op, condicion)
    If sqlgesti.status = 4 Then dato3.Enabled = True: dato3.SetFocus: GoTo no:
    nombreseccion.caption = sqlgesti.response(1, 3)
no:
End Sub

Sub leeDepartamento()
    campos(0, 0) = "codigodepto"
    campos(1, 0) = "nombre"
    campos(2, 0) = "preciodiferenciado"
    campos(3, 0) = "descuento"
    campos(4, 0) = ""
    
    campos(0, 2) = "r_maestrodepartamentos_" & rubro
    condicion = "codigoseccion='" & dato3.text & "' and  codigodepto = '" & dato4.text & "' ORDER BY codigodepto LIMIT 0,1"
    
    op = 5
    sqlgesti.response = campos
    Set sqlgesti.conexion = GESTIONrubro
    Call sqlgesti.sqlgesti(op, condicion)
    If sqlgesti.status = 4 Then dato4.Enabled = True: dato4.SetFocus: GoTo no:
    nombredepto.caption = sqlgesti.response(1, 3)
    PRECIODIFERENCIADO = sqlgesti.response(2, 3)
    dsctodpto.caption = sqlgesti.response(3, 3)
no:
End Sub

Sub leeLinea()
    campos(0, 0) = "codigolinea"
    campos(1, 0) = "nombre"
    campos(2, 0) = "descuento"
    campos(3, 0) = ""
    campos(0, 2) = "r_maestrolineas_" & rubro
    condicion = "codigoseccion='" & dato3.text & "' and codigodepto = '" & dato4.text & "' AND codigolinea = '" & dato5.text & "' ORDER BY codigolinea LIMIT 0,1"
    op = 5
    sqlgesti.response = campos
    Set sqlgesti.conexion = GESTIONrubro
    Call sqlgesti.sqlgesti(op, condicion)
    If sqlgesti.status = 4 Then dato5.Enabled = True: dato5.SetFocus: GoTo no:
    nombrelinea.caption = sqlgesti.response(1, 3)
    dsctolinea.caption = sqlgesti.response(2, 3)
no:
End Sub

Sub leeImpuesto()
    campos(0, 0) = "codigo"
    campos(1, 0) = "nombre"
    campos(2, 0) = ""
    campos(0, 2) = "g_maestroimpuestos"
    condicion = "codigo = '" & dato6.text & "' ORDER BY codigo LIMIT 0,1"
    op = 5
    Set sqlgesti.conexion = GESTION
    sqlgesti.response = campos
    Call sqlgesti.sqlgesti(op, condicion)
   'If sqlgesti.status = 4 Then dato6.SetFocus: GoTo no:
    nombreimpUESTO = sqlgesti.response(1, 3)
no:
End Sub

Sub leeMarca()
    campos(0, 0) = "codigo"
    campos(1, 0) = "nombre"
    campos(2, 0) = ""
    campos(0, 2) = "r_maestromarcas_" & rubro
    condicion = "rubro = '" & rubro & "' AND codigo = '" & dato7.text & "' ORDER BY codigo, rubro LIMIT 0,1"
    op = 5
    Set sqlgesti.conexion = GESTIONrubro
    sqlgesti.response = campos
    Call sqlgesti.sqlgesti(op, condicion)
    If sqlgesti.status = 4 Then dato7.Enabled = True: dato7.SetFocus: GoTo no:
    nombremarca = sqlgesti.response(1, 3)
no:
End Sub

Sub leeTemporada()
    campos(0, 0) = "codigo"
    campos(1, 0) = "nombre"
    campos(2, 0) = ""
    campos(0, 2) = "r_maestrotemporadas_" & rubro
    condicion = "codigo = '" & dato8.text & "' ORDER BY codigo LIMIT 0,1"
    op = 5
    Set sqlgesti.conexion = GESTIONrubro
    sqlgesti.response = campos
    Call sqlgesti.sqlgesti(op, condicion)
    If sqlgesti.status = 4 Then nombretemporada.caption = "TEMPORADA NO EXISTE": GoTo no:
    nombretemporada.caption = sqlgesti.response(1, 3)
no:
End Sub

Sub leeproveedor()
    campos(0, 0) = "rut"
    campos(1, 0) = "nombre"
    campos(2, 0) = ""
    campos(0, 2) = "r_maestroproveedores_" & rubro
    condicion = "rut = '" & dato9.text & dv.caption & "' ORDER BY rut LIMIT 0,1"
    op = 5
    Set sqlgesti.conexion = GESTIONrubro
    sqlgesti.response = campos
    Call sqlgesti.sqlgesti(op, condicion)
   If sqlgesti.status = 4 Then dato9.Enabled = True: dato9.SetFocus: GoTo no:
    nombreproveedor = sqlgesti.response(1, 3)
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

Private Sub opciones_GotFocus()
    MANUAL.SetFocus
End Sub

Sub CARGAGRILLA()
    
    
    Rem DATOS DE LA COLUMNA
    formatoGrilla(1, 1) = "COD"
    formatoGrilla(1, 2) = "TIPO PRECIO"
    formatoGrilla(1, 3) = "% REP."
    formatoGrilla(1, 4) = "MARGEN"
    formatoGrilla(1, 5) = "P.SUGERIDO "
    formatoGrilla(1, 6) = "P.SISTEMA   "
    formatoGrilla(1, 7) = "P.EN POS   "
    formatoGrilla(1, 8) = " % REAL"
    formatoGrilla(1, 9) = "LOCAL"
    
    
    Rem LARGO DE LOS DATOS
    formatoGrilla(2, 1) = "4"
    formatoGrilla(2, 2) = "15"
    formatoGrilla(2, 3) = "6"
    formatoGrilla(2, 4) = "6"
    formatoGrilla(2, 5) = "0"
    formatoGrilla(2, 6) = "9"
    formatoGrilla(2, 7) = "9"
    formatoGrilla(2, 8) = "5"
    formatoGrilla(2, 9) = "9"
    
    Rem TIPO DE DATOS C=CEROS,S=STRING,N=NUMERICO,D=DATE,
    formatoGrilla(3, 1) = "S"
    formatoGrilla(3, 2) = "S"
    formatoGrilla(3, 3) = "N"
    formatoGrilla(3, 4) = "N"
    formatoGrilla(3, 5) = "N"
    formatoGrilla(3, 6) = "N"
    formatoGrilla(3, 7) = "N"
    formatoGrilla(3, 8) = "N"
    formatoGrilla(3, 9) = "S"
    
    Rem FORMATO GRILLA
    formatoGrilla(4, 1) = ""
    formatoGrilla(4, 2) = ""
    formatoGrilla(4, 3) = "%##0.00"
    formatoGrilla(4, 4) = "%##0.00"
    formatoGrilla(4, 5) = "$ ###,###,##0"
    formatoGrilla(4, 6) = "$ ###,###,##0"
    formatoGrilla(4, 7) = "$ ###,###,##0"
    formatoGrilla(4, 8) = ""
    formatoGrilla(4, 9) = ""
    
    Rem LOCCKED
    formatoGrilla(5, 1) = "TRUE"
    formatoGrilla(5, 2) = "TRUE"
    formatoGrilla(5, 3) = "TRUE"
    formatoGrilla(5, 4) = "TRUE"
    formatoGrilla(5, 5) = "TRUE"
    formatoGrilla(5, 6) = "TRUE"
    formatoGrilla(5, 7) = "TRUE"
    formatoGrilla(5, 8) = "TRUE"
    formatoGrilla(5, 9) = "TRUE"
    
    Grid1.Cols = 10
    Grid1.Rows = 2
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
    For k = 1 To Grid1.Cols - 1
        Grid1.Cell(0, k).text = formatoGrilla(1, k)
        Grid1.Column(k).Width = Val(formatoGrilla(2, k)) * Grid1.DefaultFont.Size
        
        Grid1.Column(k).MaxLength = Val(formatoGrilla(2, k))
        Grid1.Column(k).FormatString = formatoGrilla(4, k)
        Grid1.Column(k).Locked = formatoGrilla(5, k)
        If formatoGrilla(3, k) = "N" Then Grid1.Column(k).Alignment = cellRightCenter
       
    Next k
    Grid1.Column(0).Width = 0
    Grid1.Range(0, 0, 0, Grid1.Cols - 1).Alignment = cellCenterCenter
    Call tiposdeprecios
    Grid1.Enabled = False
End Sub
Sub CARGAGRILLAPACK()
    
    
    Rem DATOS DE LA COLUMNA
    formatoGrilla(1, 1) = "CODIGO"
    formatoGrilla(1, 2) = "DESCRIPCION"
    formatoGrilla(1, 3) = "PRECIO"
    formatoGrilla(1, 4) = "UNIDADES"
    
    
    Rem LARGO DE LOS DATOS
    formatoGrilla(2, 1) = "10"
    formatoGrilla(2, 2) = "35"
    formatoGrilla(2, 3) = "8"
    formatoGrilla(2, 4) = "8"
    
    Rem TIPO DE DATOS C=CEROS,S=STRING,N=NUMERICO,D=DATE,
    formatoGrilla(3, 1) = "S"
    formatoGrilla(3, 2) = "S"
    formatoGrilla(3, 3) = "N"
    formatoGrilla(3, 4) = "N"
    
    Rem FORMATO GRILLA
    formatoGrilla(4, 1) = ""
    formatoGrilla(4, 2) = ""
    formatoGrilla(4, 3) = "$ ###,###,##0"
    formatoGrilla(4, 4) = "  ###,###,##0"
    
    Rem LOCCKED
    formatoGrilla(5, 1) = "TRUE"
    formatoGrilla(5, 2) = "TRUE"
    formatoGrilla(5, 3) = "TRUE"
    formatoGrilla(5, 4) = "TRUE"
    
    grillapack.Cols = 5
    grillapack.Rows = 2
    grillapack.AllowUserResizing = False
    grillapack.DisplayFocusRect = False
    grillapack.ExtendLastCol = True
    grillapack.BoldFixedCell = False
    grillapack.DrawMode = cellOwnerDraw
    grillapack.Appearance = Flat
    grillapack.ScrollBarStyle = Flat
    grillapack.FixedRowColStyle = Flat
    grillapack.BackColorFixed = RGB(90, 158, 214)
    grillapack.BackColorFixedSel = RGB(110, 180, 230)
    grillapack.BackColorBkg = RGB(90, 158, 214)
    grillapack.BackColorScrollBar = RGB(231, 235, 247)
    grillapack.BackColor1 = RGB(231, 235, 247)
    grillapack.BackColor2 = RGB(239, 243, 255)
    grillapack.GridColor = RGB(148, 190, 231)
    For k = 1 To grillapack.Cols - 1
        grillapack.Cell(0, k).text = formatoGrilla(1, k)
        grillapack.Column(k).Width = Val(formatoGrilla(2, k)) * grillapack.DefaultFont.Size
        
        grillapack.Column(k).MaxLength = Val(formatoGrilla(2, k))
        grillapack.Column(k).FormatString = formatoGrilla(4, k)
        grillapack.Column(k).Locked = formatoGrilla(5, k)
        If formatoGrilla(3, k) = "N" Then grillapack.Column(k).Alignment = cellRightCenter
       
    Next k
    grillapack.Column(0).Width = 0
    grillapack.Range(0, 0, 0, grillapack.Cols - 1).Alignment = cellCenterCenter


End Sub

Sub tiposdeprecios()
    Dim resultados As rdoResultset
    Dim cSql As New rdoQuery
    Dim rut As String
    Dim linea As Double
    
        Set cSql.ActiveConnection = GESTION
        cSql.sql = "SELECT tp.codigo,tp.nombre,tp.porcentajedelmargen,me.codigo,me.nombre "
        cSql.sql = cSql.sql + "FROM g_maestrodetiposdeprecios as tp,g_maestroempresas as me where me.rubro='" + rubro + "' ORDER BY me.codigo "
        'cSql.SQL = cSql.SQL + "WHERE local='" + codigoempresa + "' order by codigobodega"
        cSql.Execute
        linea = 0
        Grid1.Rows = 1
   
        If cSql.RowsAffected > 0 Then
            Set resultados = cSql.OpenResultset
            
            While Not resultados.EOF
                Grid1.Rows = Grid1.Rows + 1
                linea = linea + 1
                Grid1.Cell(linea, 1).text = resultados(0)
                Grid1.Cell(linea, 2).text = resultados(1)
                Grid1.Cell(linea, 3).text = resultados(2)
                Grid1.Cell(linea, 8).text = resultados(3)
                Grid1.Cell(linea, 9).text = resultados(4)
                
                PORCENTAJES(linea) = resultados(2)
                resultados.MoveNext
            
            Wend
            resultados.Close
        Set resultados = Nothing

        End If
End Sub

Sub empresasdelrubro()
    Dim resultados As rdoResultset
    Dim cSql As New rdoQuery
    Dim rut As String
    Dim linea As Double
    
        Set cSql.ActiveConnection = GESTION
        cSql.sql = "SELECT codigo,nombre "
        cSql.sql = cSql.sql + "FROM g_maestroempresas "
        cSql.sql = cSql.sql + "WHERE rubro='" & rubro & "' ORDER BY codigo"
        cSql.Execute
        linea = 0
      
        If cSql.RowsAffected > 0 Then
            Set resultados = cSql.OpenResultset
            While Not resultados.EOF
            linea = linea + 1
            empresasrubro(linea) = resultados(0)
            resultados.MoveNext
            Wend
            resultados.Close
        Set resultados = Nothing

        End If
cer = linea

End Sub

Sub CARGAGRILLAbodegas()
    Rem DATOS DE LA COLUMNA
    formatoGrilla(1, 1) = "LOCAL"
    formatoGrilla(1, 2) = "BODEGA"
    formatoGrilla(1, 3) = "UBICACION"
    formatoGrilla(1, 4) = "CAJAS "
    formatoGrilla(1, 5) = "UXC   "
    formatoGrilla(1, 6) = "UNIDADES"
    
    Rem LARGO DE LOS DATOS
    formatoGrilla(2, 1) = "15"
    formatoGrilla(2, 2) = "12"
    formatoGrilla(2, 3) = "12"
    formatoGrilla(2, 4) = "8"
    formatoGrilla(2, 5) = "8"
    formatoGrilla(2, 6) = "8"
    
    Rem TIPO DE DATOS C=CEROS,S=STRING,N=NUMERICO,D=DATE,
    formatoGrilla(3, 1) = "S"
    formatoGrilla(3, 2) = "S"
    formatoGrilla(3, 3) = "S"
    formatoGrilla(3, 4) = "N"
    formatoGrilla(3, 5) = "N"
    formatoGrilla(3, 6) = "N"
    
    Rem FORMATO GRILLA
    formatoGrilla(4, 1) = ""
    formatoGrilla(4, 2) = ""
    formatoGrilla(4, 3) = ""
    formatoGrilla(4, 4) = "#,###,##0.0"
    formatoGrilla(4, 5) = "#,###,##0"
    formatoGrilla(4, 6) = "#,###,##0.0"
    
    Rem LOCCKED
    formatoGrilla(5, 1) = "TRUE"
    formatoGrilla(5, 2) = "TRUE"
    formatoGrilla(5, 3) = "TRUE"
    formatoGrilla(5, 4) = "TRUE"
    formatoGrilla(5, 5) = "TRUE"
    formatoGrilla(5, 6) = "TRUE"
    
    Grid2.Cols = 7
    Grid2.Rows = 2
    
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
    For k = 1 To Grid2.Cols - 1
        Grid2.Cell(0, k).text = formatoGrilla(1, k)
        Grid2.Column(k).Width = Val(formatoGrilla(2, k)) * Grid2.DefaultFont.Size
        Grid2.Column(k).MaxLength = Val(formatoGrilla(2, k))
        Grid2.Column(k).FormatString = formatoGrilla(4, k)
        Grid2.Column(k).Locked = formatoGrilla(5, k)
        If formatoGrilla(3, k) = "N" Then Grid2.Column(k).Alignment = cellRightCenter
       
    Next k
    Grid2.Column(0).Width = 0
    Grid2.Range(0, 0, 0, Grid2.Cols - 1).Alignment = cellCenterCenter
    Grid2.Enabled = False
End Sub

'Sub CARGAGRILLAbodegas()
'    Rem DATOS DE LA COLUMNA
'    formatoGrilla(1, 1) = "LOCAL"
'    formatoGrilla(1, 2) = "BODEGA"
'    formatoGrilla(1, 3) = "STOCK "
'    formatoGrilla(1, 4) = "UNI. X CAJA"
'    formatoGrilla(1, 5) = " UNIDADES"
'
'
'    Rem LARGO DE LOS DATOS
'    formatoGrilla(2, 1) = "15"
'    formatoGrilla(2, 2) = "15"
'    formatoGrilla(2, 3) = "10"
'    formatoGrilla(2, 4) = "10"
'    formatoGrilla(2, 5) = "10"
'
'    Rem TIPO DE DATOS C=CEROS,S=STRING,N=NUMERICO,D=DATE,
'    formatoGrilla(3, 1) = "S"
'    formatoGrilla(3, 2) = "S"
'    formatoGrilla(3, 3) = "N"
'    formatoGrilla(3, 4) = "N"
'    formatoGrilla(3, 5) = "N"
'
'    Rem FORMATO GRILLA
'    formatoGrilla(4, 1) = ""
'    formatoGrilla(4, 2) = ""
'    formatoGrilla(4, 3) = "#,###,##0.0"
'    formatoGrilla(4, 4) = "#,###,##0.0"
'    formatoGrilla(4, 5) = "#,###,##0.0"
'
'    Rem LOCCKED
'    formatoGrilla(5, 1) = "TRUE"
'    formatoGrilla(5, 2) = "TRUE"
'    formatoGrilla(5, 3) = "TRUE"
'    formatoGrilla(5, 4) = "TRUE"
'    formatoGrilla(5, 5) = "TRUE"
'    Grid2.Cols = 6
'    Grid2.Rows = 2
'
'    Grid2.AllowUserResizing = False
'    Grid2.DisplayFocusRect = False
'    Grid2.ExtendLastCol = True
'    Grid2.BoldFixedCell = False
'    Grid2.DrawMode = cellOwnerDraw
'    Grid2.Appearance = Flat
'    Grid2.ScrollBarStyle = Flat
'    Grid2.FixedRowColStyle = Flat
'    Grid2.BackColorFixed = RGB(90, 158, 214)
'    Grid2.BackColorFixedSel = RGB(110, 180, 230)
'    Grid2.BackColorBkg = RGB(90, 158, 214)
'    Grid2.BackColorScrollBar = RGB(231, 235, 247)
'    Grid2.BackColor1 = RGB(231, 235, 247)
'    Grid2.BackColor2 = RGB(239, 243, 255)
'    Grid2.GridColor = RGB(148, 190, 231)
'    For k = 1 To Grid2.Cols - 1
'        Grid2.Cell(0, k).text = formatoGrilla(1, k)
'        Grid2.Column(k).Width = Val(formatoGrilla(2, k)) * Grid2.DefaultFont.Size
'        Grid2.Column(k).MaxLength = Val(formatoGrilla(2, k))
'        Grid2.Column(k).FormatString = formatoGrilla(4, k)
'        Grid2.Column(k).Locked = formatoGrilla(5, k)
'        If formatoGrilla(3, k) = "N" Then Grid2.Column(k).Alignment = cellRightCenter
'
'    Next k
'    Grid2.Column(0).Width = 0
'    Grid2.Range(0, 0, 0, Grid2.Cols - 1).Alignment = cellCenterCenter
'    Grid2.Enabled = False
'End Sub

Sub CARGASTOCKBODEGAS()
    Dim a As Integer
    Dim resultados As rdoResultset
    Dim cSql As New rdoQuery
    Dim rut As String
    Dim saldo As Double

        Set cSql.ActiveConnection = GESTIONrubro
        cSql.sql = "SELECT local,bodega,ubicacion,stockactual "
        cSql.sql = cSql.sql + "FROM r_maestroproductos_stock_" & rubro & " "
        cSql.sql = cSql.sql + "WHERE año='2007' AND codigo='" + dato1.text + "' order by local,bodega limit 0,10"
        cSql.Execute
        Grid2.Rows = 1
        Grid2.AutoRedraw = False
        If cSql.RowsAffected > 0 Then
            Set resultados = cSql.OpenResultset
            While Not resultados.EOF
                saldo = resultados(3)
                If dato13.text = "0" Then dato13.text = "1"
                Grid2.AddItem leelocal(resultados(0)) & vbTab & leebodega(resultados(1)) & vbTab & resultados(2) & vbTab & saldo / dato13.text & vbTab & dato13.text & vbTab & saldo, False
                resultados.MoveNext
            Wend
            resultados.Close
            Set resultados = Nothing
        End If
        Grid2.AutoRedraw = True
        Grid2.Refresh
        Grid2.Enabled = True
End Sub

Sub leerstock(locales, bodega, linea)
    Dim total As Double

    campos(0, 0) = "codigo"
    campos(1, 0) = "stockactual"
    campos(2, 0) = ""
    condicion = "local = '" & locales & "' AND bodega = '" & bodega & "' AND codigo = '" & dato1.text & "' and año='" + año + "' order by local asc"
    campos(0, 2) = "r_maestroproductos_stock_" & rubro
    op = 5
    sqlgesti.response = campos
    Set sqlgesti.conexion = GESTION
    Call sqlgesti.sqlgesti(op, condicion)
    If sqlgesti.status = 4 Then Grid2.Cell(linea, 3).text = 0: Grid2.Cell(linea, 4).text = dato13.text: Grid2.Cell(linea, 5).text = 0: GoTo no:
    Grid2.Cell(linea, 3).text = sqlgesti.response(1, 3)
    Grid2.Cell(linea, 4).text = dato13.text
    total = Grid2.Cell(linea, 3).text + Grid2.Cell(linea, 4).text
    Grid2.Cell(linea, 5).text = total

no:
End Sub

Function leelocal(codigo) As String
    campos(0, 0) = "nombre"
    campos(1, 0) = ""
    campos(0, 2) = "g_maestroempresas"
    condicion = "codigo = '" & codigo & "'"
    op = 5
    Set sqlgesti.conexion = GESTION
    sqlgesti.response = campos
    Call sqlgesti.sqlgesti(op, condicion)
    leelocal = sqlgesti.response(0, 3)
End Function

Function leebodega(codigo) As String
    campos(0, 0) = "nombre"
    campos(2, 0) = ""
    campos(0, 2) = "r_maestrobodegas_" & rubro
    condicion = "rubro = '" & rubro & "' AND codigobodega = '" & codigo & "'"
    op = 5
    Set sqlgesti.conexion = GESTIONrubro
    sqlgesti.response = campos
    Call sqlgesti.sqlgesti(op, condicion)
    leebodega = sqlgesti.response(0, 3)
End Function

Sub grabarPrecios(codigo)
    campos(0, 0) = "local"
    campos(1, 0) = "codigo"
    campos(2, 0) = "codigoprecio"
    campos(3, 0) = "precioautomatico"
    campos(4, 0) = "preciosistema"
    campos(5, 0) = "preciopuntoventa"
    campos(6, 0) = "preciocosto"
    campos(7, 0) = "margen"
    campos(8, 0) = "fechavigencia"
    campos(9, 0) = ""
  
    campos(1, 1) = dato1.text
    campos(8, 1) = fechasistema

    For k = 1 To Grid1.Rows - 1
     
      
       campos(0, 1) = Grid1.Cell(k, 8).text
       campos(2, 1) = Grid1.Cell(k, 1).text 'codigoprecio
       campos(3, 1) = Grid1.Cell(k, 5).text 'precioautomatico
       campos(4, 1) = Grid1.Cell(k, 6).text 'preciosistema
       campos(5, 1) = Grid1.Cell(k, 7).text 'preciopuntoventa
       'dato18.text = Replace(dato18.text, ".", "")
       campos(6, 1) = Replace(dato18.text, ",", ".") 'preciocosto
       campos(7, 1) = Replace(Grid1.Cell(k, 4).text, ",", ".") 'margen
    
       campos(0, 2) = "r_maestroproductos_precios_" & rubro
       condicion = ""
        If MODIFI = "0" Then
        op = 2
        Else
        op = 3
        condicion = "codigo='" + campos(1, 1) + "' and codigoprecio='" + campos(2, 1) + "' and local='" + campos(0, 1) + "' "
        End If
       
       
       Set sqlgesti.conexion = GESTIONrubro
       sqlgesti.response = campos
       Call sqlgesti.sqlgesti(op, condicion)
           
    Next k
End Sub

Sub leeprecios(codigo)
Dim real As String


    campos(0, 0) = "local"
    campos(1, 0) = "codigo"
    campos(2, 0) = "codigoprecio"
    campos(3, 0) = "precioautomatico"
    campos(4, 0) = "preciosistema"
    campos(5, 0) = "preciopuntoventa"
    campos(6, 0) = "preciocosto"
    campos(7, 0) = "margen"
    campos(8, 0) = "fechavigencia"
    campos(9, 0) = ""
    
    For k = 1 To Grid1.Rows - 1
        condicion = "codigo = '" + dato1.text + "' AND codigoprecio='" + Grid1.Cell(k, 1).text + "' and local='" + Grid1.Cell(k, 8).text + "' "
        campos(0, 2) = "r_maestroproductos_precios_" & rubro
        op = 5
        Set sqlgesti.conexion = GESTIONrubro
        sqlgesti.response = campos
        Call sqlgesti.sqlgesti(op, condicion)
        If sqlgesti.status = 0 Then
            If CDbl(dato18.text) <> 0 Then
            real = Round(((sqlgesti.response(5, 3) / CDbl(dato18.text)) - 1) * 100, 2)
            End If
            
            Grid1.Cell(k, 5).text = sqlgesti.response(3, 3)
            Grid1.Cell(k, 6).text = sqlgesti.response(4, 3)
            Grid1.Cell(k, 7).text = sqlgesti.response(5, 3)
            Grid1.Cell(k, 4).text = real
        
        Else
        
            Grid1.Cell(k, 5).text = "0"
            Grid1.Cell(k, 6).text = "0"
            Grid1.Cell(k, 7).text = "0"
            Grid1.Cell(k, 4).text = "0"
    End If
    
    Next k
End Sub

Sub LeeUltimo()
    campos(0, 0) = "codigobarra"
    campos(1, 0) = ""
    campos(0, 2) = "r_maestroproductos_fijo_" & rubro
    
    If rubro = "02" Then apostrofe = "19"
    
    
    condicion = "codigobarra < '" + apostrofe + "99999999999' order by codigobarra desc"
    op = 5
    sqlgesti.response = campos
    Set sqlgesti.conexion = GESTIONrubro
    Call sqlgesti.sqlgesti(op, condicion)
    If Mid(sqlgesti.response(0, 3), 1, 2) = "19" Then
    dato1.text = Mid(sqlgesti.response(0, 3), 1, 12) + 1
    
    Else
    dato1.text = "190000000001"
    End If
    
    dato1.text = dato1.text + Validaciones.dv(dato1.text)
    Call ceros(dato1)
End Sub

Private Sub pack_Click()


End Sub

Private Sub recupera_Click()
LeeUltimo
dato1.SetFocus

End Sub
    Private Sub CargaGrillaLista(ByVal Row As Integer, ByVal Col As Integer)
        Dim i As Integer
        Rem DATOS DE LA COLUMNA
        formatoGrilla(1, 1) = "CODIGO"
        formatoGrilla(1, 2) = "DESCRIPCION"
        formatoGrilla(1, 3) = "CANTIDAD"
        formatoGrilla(1, 4) = "P.ESPECIAL"
        formatoGrilla(1, 5) = "MARGEN"
        
        Rem LARGO DE LOS DATOS
        formatoGrilla(2, 1) = "13"
        formatoGrilla(2, 2) = "50"
        formatoGrilla(2, 3) = "9"
        formatoGrilla(2, 4) = "8"
        formatoGrilla(2, 5) = "5"
        Rem TIPO DE DATOS C=CEROS,S=STRING,N=NUMERICO,D=DATE,
        formatoGrilla(3, 1) = "N"
        formatoGrilla(3, 2) = "S"
        formatoGrilla(3, 3) = "N"
        formatoGrilla(3, 4) = "N"
        formatoGrilla(3, 5) = "N"
        Rem FORMATO GRILLA
        formatoGrilla(4, 1) = "0000000000000"
        formatoGrilla(4, 2) = ""
        formatoGrilla(4, 3) = "########0"
        formatoGrilla(4, 4) = "$ ###,###,##0"
        formatoGrilla(4, 5) = "% #,##0.0"
        Rem LOCCKED
        formatoGrilla(5, 1) = "TRUE"
        formatoGrilla(5, 2) = "TRUE"
        formatoGrilla(5, 3) = "TRUE"
        formatoGrilla(5, 4) = "TRUE"
        formatoGrilla(5, 5) = "TRUE"
        Rem VALOR MINIMO
        formatoGrilla(6, 1) = ""
        formatoGrilla(6, 2) = ""
        formatoGrilla(6, 3) = ""
        formatoGrilla(6, 4) = ""
        formatoGrilla(6, 5) = ""
        Rem VALOR MAXIMO
        formatoGrilla(7, 1) = ""
        formatoGrilla(7, 2) = ""
        formatoGrilla(7, 3) = ""
        formatoGrilla(7, 4) = ""
        formatoGrilla(7, 5) = ""
        Rem ANCHO
        formatoGrilla(8, 1) = "7"
        formatoGrilla(8, 2) = "25"
        formatoGrilla(8, 3) = "6"
        formatoGrilla(8, 4) = "10"
        formatoGrilla(8, 5) = "5"
        lista.Cols = Col
        lista.Rows = Row
        lista.AllowUserResizing = False
        lista.DisplayFocusRect = False
        lista.ExtendLastCol = True
        lista.BoldFixedCell = False
        lista.DrawMode = cellOwnerDraw
        lista.Appearance = Flat
        lista.ScrollBarStyle = Flat
        lista.FixedRowColStyle = Flat
        lista.BackColorFixed = RGB(90, 214, 158)
        lista.BackColorFixedSel = RGB(110, 230, 180)
        lista.BackColorBkg = RGB(90, 214, 158)
        lista.BackColorScrollBar = RGB(231, 247, 235)
        lista.BackColor1 = RGB(231, 247, 235)
        lista.BackColor2 = RGB(239, 255, 243)
        lista.GridColor = RGB(148, 231, 190)
        
        lista.Column(0).Width = 0
        For i = 1 To Col - 1
            lista.Cell(0, i).text = formatoGrilla(1, i)
            lista.Column(i).Width = Val(formatoGrilla(8, i)) * (lista.Cell(0, i).Font.Size + 1.25)
            lista.Column(i).MaxLength = Val(formatoGrilla(2, i))
            lista.Column(i).FormatString = formatoGrilla(4, i)
            lista.Column(i).Locked = formatoGrilla(5, i)
            If formatoGrilla(3, i) = "N" Then
                lista.Column(i).Alignment = cellRightCenter
            Else
                lista.Column(i).Alignment = cellLeftCenter
            End If
        Next i
        lista.Range(0, 1, 0, lista.Cols - 1).Alignment = cellCenterCenter
        lista.Enabled = True
    End Sub

    Private Sub leerEspeciales()
        Dim tabla As String
        Dim costo As Double
        Dim utilidad As Double
        
        
        tabla = "SELECT CONCAT(mppc.codigo, '" & vbTab & "', mpf.descripcion, '" & vbTab & "', mppc.cantidad, '" & vbTab & "', mppc.precio) AS item "
        tabla = tabla & "FROM r_maestroproductos_fijo_" & rubro & " As mpf INNER JOIN r_maestroproductos_precio_cantidad_" & rubro & " AS mppc ON mpf.codigobarra = mppc.codigo "
        tabla = tabla & "WHERE mpf.codigobarra = '" & dato1.text & "' ORDER BY mppc.cantidad ASC"
        Call ConectarControlData(data, servidor, basedatos & rubro, usuario, password, tabla)
        
        lista.Rows = 1
        lista.AutoRedraw = False
        If data.Recordset.RecordCount > 0 Then
            data.Recordset.MoveFirst
            While Not data.Recordset.EOF
            costo = CDbl(dato18.text)
              lista.AddItem data.Recordset.Fields("item"), True
            
            If costo <> 0 And lista.Rows > 1 Then
            
            utilidad = ((CDbl(lista.Cell(lista.Rows - 1, 4).text) / costo) - 1) * 100
            End If
            
              lista.Cell(lista.Rows - 1, 5).text = utilidad
                data.Recordset.MoveNext
            Wend
        lista.AutoRedraw = True
        lista.Refresh
        End If
    End Sub

Sub grabarUBICACIONES(codigo)
       Dim bodepaso As String
       
        campos(0, 0) = "ubicacion"
        campos(1, 0) = ""
        For k = 1 To Grid2.Rows - 1
        bodepaso = "0" + Mid(Str(k - 1), 2, 1)
       campos(0, 2) = "r_maestroproductos_stock_" & rubro
       condicion = "local = '" & empresaactiva & "' and codigo ='" + codigo + "' and año='" + Format(fechasistema, "yyyy") + "' and bodega='" + bodepaso + "'"
       op = 5
       Set sqlgesti.conexion = GESTIONrubro
       sqlgesti.response = campos
       Call sqlgesti.sqlgesti(op, condicion)
        If sqlgesti.status = 0 Then
        
       campos(0, 1) = Grid2.Cell(k, 3).text
       op = 3
       Set sqlgesti.conexion = GESTIONrubro
       sqlgesti.response = campos
       Call sqlgesti.sqlgesti(op, condicion)
        End If
    Next k
    End Sub
    Sub nuevo_codigo()
            
        'MAESTRO PRODUCTOS FIJO
        campos(0, 0) = "codigobarra"
        campos(1, 0) = ""
        
        campos(0, 1) = dato30.text
        campos(0, 2) = "r_maestroproductos_fijo_" & rubro
               
                condicion = "codigobarra = '" & dato1.text & "' "
                op = 3
                sqlgesti.response = campos
                Set sqlgesti.conexion = GESTIONrubro
                Call sqlgesti.sqlgesti(op, condicion)
        
        'MAESTRO PRODUCTOS ESTADISTICA
        campos(0, 0) = "codigo"
        campos(1, 0) = ""
        
        campos(0, 1) = dato30.text
        campos(0, 2) = "r_maestroproductos_estadistica_" & rubro
        condicion = "codigo = '" & dato1.text & "'"
        op = 3
        sqlgesti.response = campos
        Set sqlgesti.conexion = GESTIONrubro
        Call sqlgesti.sqlgesti(op, condicion)
        op = sqlgesti.status
        
        'MAESTRO PRODUCTOS STOCK
        campos(0, 0) = "codigo"
        campos(1, 0) = ""
        
        campos(0, 1) = dato30.text
        campos(0, 2) = "r_maestroproductos_stock_" & rubro
        condicion = "codigo = '" & dato1.text & "'"
        op = 3
        sqlgesti.response = campos
        Set sqlgesti.conexion = GESTIONrubro
        Call sqlgesti.sqlgesti(op, condicion)
        op = sqlgesti.status
        
         'MAESTRO PRODUCTOS PRECIOS
        campos(0, 0) = "codigo"
        campos(1, 0) = ""
        
        campos(0, 1) = dato30.text
        campos(0, 2) = "r_maestroproductos_precios_" & rubro
        condicion = "codigo = '" & dato1.text & "'"
        op = 3
        sqlgesti.response = campos
        Set sqlgesti.conexion = GESTIONrubro
        Call sqlgesti.sqlgesti(op, condicion)
        op = sqlgesti.status
        
         ' MOVIMIENTOS DETALLE
        campos(0, 0) = "codigo"
        campos(1, 0) = ""
        
        campos(0, 1) = dato30.text
        campos(0, 2) = "l_movimientos_detalle_" & rubro
        
               
        condicion = "codigo = '" & dato1.text & "' "
        op = 3
        sqlgesti.response = campos
        Set sqlgesti.conexion = GESTIONrubro
        Call sqlgesti.sqlgesti(op, condicion)


       'orden de compra DETALLE
        campos(0, 0) = "codigo"
        campos(1, 0) = ""
        
        campos(0, 1) = dato30.text
        campos(0, 2) = "l_ordendecompra_detalle_" & rubro
        
               
        condicion = "codigo = '" & dato1.text & "' "
        op = 3
        sqlgesti.response = campos
        Set sqlgesti.conexion = GESTIONrubro
        Call sqlgesti.sqlgesti(op, condicion)
        
          'VENTAS DOCUMENTOS DETALLE
        
        campos(0, 0) = "codigo"
        campos(1, 0) = ""

        campos(0, 1) = dato30.text
        campos(0, 2) = "sv_documento_detalle_" & empresaactiva
        condicion = "codigo = '" & dato1.text & "'"
        op = 3
        sqlgesti.response = campos
        Set sqlgesti.conexion = ventasrubro
        Call sqlgesti.sqlgesti(op, condicion)
        op = sqlgesti.status
        dato1.text = dato30.text
        dato30.text = ""
        Frmcodigobarra.Visible = False
        
        MsgBox "El Codigo se Cambiado Exitosamente", vbOKOnly, "Guardado"
        
        retorno
End Sub
Sub lee_si_existe()
    campos(0, 0) = "codigobarra"
    campos(1, 0) = ""
    
    campos(0, 2) = "r_maestroproductos_fijo_" & rubro
    condicion = "codigobarra = '" & dato30.text & "' "
     
    op = 5
    sqlgesti.response = campos
    Set sqlgesti.conexion = GESTIONrubro
    Call sqlgesti.sqlgesti(op, condicion)
    If sqlgesti.status = 0 Then
    
       If MsgBox("CODIGO EXISTE", vbOKOnly) = vbOK Then
         dato30.text = ""
         dato30.SetFocus
         End If
       Else
    
        If dato30.text = "0000000000000" Then
         If MsgBox("CODIGO INVALIDO", vbOKOnly) = vbOK Then
          dato30.text = ""
          dato30.SetFocus
         End If
        Else
        Call nuevo_codigo
      End If
    End If
    End Sub

Private Sub RETORNO_CODIGO_Click()
Frmcodigobarra.Visible = False
dato30.text = ""
DATO2.SetFocus
End Sub

