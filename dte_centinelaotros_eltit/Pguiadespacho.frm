VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash10a.ocx"
Object = "{4777436C-EB8C-4596-98A8-EBCF98683969}#1.0#0"; "Flexcell.ocx"
Object = "{ADBBAED6-B16F-40DC-80DF-B44910CBA76C}#1.0#0"; "Frame-Xp.ocx"
Begin VB.Form Pguiadespacho 
   Caption         =   "Guia de Despacho Fletes"
   ClientHeight    =   8580
   ClientLeft      =   60
   ClientTop       =   465
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   ScaleHeight     =   8580
   ScaleWidth      =   15240
   StartUpPosition =   2  'CenterScreen
   Begin XPFrame.FrameXp FOLI 
      Height          =   1680
      Left            =   4635
      TabIndex        =   48
      Top             =   3105
      Width           =   4110
      _ExtentX        =   7250
      _ExtentY        =   2963
      BackColor       =   16744576
      Caption         =   "INGRESE FOLIO"
      CaptionEstilo3D =   1
      BackColor       =   16744576
      ForeColor       =   8438015
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
      Begin VB.CommandButton GENERA 
         Caption         =   "GENERAR GUIA"
         Height          =   375
         Left            =   1215
         TabIndex        =   50
         Top             =   1080
         Width           =   1725
      End
      Begin VB.TextBox FOLIO 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Left            =   810
         MaxLength       =   10
         TabIndex        =   49
         Top             =   495
         Width           =   2400
      End
   End
   Begin XPFrame.FrameXp PANTALLA 
      Height          =   8550
      Left            =   0
      TabIndex        =   15
      Top             =   0
      Width           =   15255
      _ExtentX        =   26908
      _ExtentY        =   15081
      BackColor       =   12632256
      Caption         =   "Datos Documento"
      CaptionEstilo3D =   2
      BackColor       =   12632256
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
      Begin VB.TextBox numerointerno 
         Height          =   285
         Left            =   10680
         TabIndex        =   62
         Top             =   720
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.TextBox foliofiscal 
         Height          =   285
         Left            =   10680
         TabIndex        =   61
         Top             =   360
         Visible         =   0   'False
         Width           =   975
      End
      Begin XPFrame.FrameXp docu 
         Height          =   1110
         Left            =   135
         TabIndex        =   42
         Top             =   225
         Width           =   7050
         _ExtentX        =   12435
         _ExtentY        =   1958
         BackColor       =   16761024
         Caption         =   "documento"
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
         Begin VB.OptionButton Option4 
            BackColor       =   &H0000FF00&
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
            Height          =   195
            Left            =   4995
            TabIndex        =   56
            Top             =   855
            Width           =   1770
         End
         Begin VB.OptionButton Option3 
            BackColor       =   &H0000FF00&
            Caption         =   "Folios Offset"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   4995
            TabIndex        =   55
            Top             =   585
            Value           =   -1  'True
            Width           =   1770
         End
         Begin VB.TextBox ncaja 
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
            MaxLength       =   2
            TabIndex        =   53
            Tag             =   "proveedor"
            Top             =   270
            Width           =   435
         End
         Begin VB.TextBox dato23 
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
            Left            =   6510
            MaxLength       =   2
            TabIndex        =   52
            Tag             =   "proveedor"
            Top             =   270
            Width           =   435
         End
         Begin VB.TextBox dato20 
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
            Left            =   900
            MaxLength       =   2
            TabIndex        =   0
            Tag             =   "proveedor"
            Top             =   270
            Width           =   495
         End
         Begin VB.TextBox dato21 
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
            Left            =   3960
            MaxLength       =   10
            TabIndex        =   43
            Tag             =   "proveedor"
            Top             =   270
            Width           =   1395
         End
         Begin VB.Label lbllocal 
            Appearance      =   0  'Flat
            BackColor       =   &H00F5C9B1&
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
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   120
            TabIndex        =   57
            Top             =   720
            Width           =   4545
         End
         Begin VB.Label Label9 
            Appearance      =   0  'Flat
            BackColor       =   &H00F5C9B1&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " Caja"
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
            Left            =   1440
            TabIndex        =   54
            Top             =   270
            Width           =   570
         End
         Begin VB.Label Label8 
            Appearance      =   0  'Flat
            BackColor       =   &H00F5C9B1&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " Local"
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
            Left            =   5430
            TabIndex        =   51
            Top             =   270
            Width           =   975
         End
         Begin VB.Label Label6 
            Appearance      =   0  'Flat
            BackColor       =   &H00F5C9B1&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " Tipo"
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
            Left            =   45
            TabIndex        =   45
            Top             =   270
            Width           =   795
         End
         Begin VB.Label Label5 
            Appearance      =   0  'Flat
            BackColor       =   &H00F5C9B1&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " Numero"
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
            Left            =   2880
            TabIndex        =   44
            Top             =   270
            Width           =   975
         End
      End
      Begin VB.OptionButton Option2 
         Caption         =   " Por Documento"
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
         Left            =   7320
         TabIndex        =   41
         Top             =   450
         Value           =   -1  'True
         Width           =   2565
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Por Rut "
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
         Left            =   7320
         TabIndex        =   40
         Top             =   960
         Width           =   2520
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FF8080&
         Caption         =   "Retorno"
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
         Left            =   10755
         Style           =   1  'Graphical
         TabIndex        =   39
         Top             =   7155
         UseMaskColor    =   -1  'True
         Width           =   1815
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
         Left            =   1755
         MaxLength       =   9
         TabIndex        =   1
         Tag             =   "proveedor"
         Top             =   -180
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.TextBox dato13 
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
         Left            =   1815
         MaxLength       =   50
         TabIndex        =   11
         Tag             =   "proveedor"
         Top             =   2475
         Width           =   11655
      End
      Begin VB.CommandButton guardar 
         BackColor       =   &H00FF8080&
         Caption         =   "Genera guia"
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
         Left            =   12915
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   7155
         UseMaskColor    =   -1  'True
         Width           =   1815
      End
      Begin VB.TextBox dato9 
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
         Left            =   8655
         MaxLength       =   10
         TabIndex        =   8
         Tag             =   "proveedor"
         Top             =   1755
         Width           =   4815
      End
      Begin XPFrame.FrameXp frmTipo 
         Height          =   1200
         Left            =   6720
         TabIndex        =   23
         Top             =   -1440
         Width           =   2715
         _ExtentX        =   4789
         _ExtentY        =   2117
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
            TabIndex        =   27
            Top             =   495
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
            TabIndex        =   26
            Top             =   1200
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
            TabIndex        =   25
            Top             =   1560
            Width           =   2475
         End
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
            TabIndex        =   24
            Top             =   840
            Width           =   2475
         End
      End
      Begin VB.TextBox dato11 
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
         Left            =   8655
         MaxLength       =   50
         TabIndex        =   10
         Tag             =   "proveedor"
         Top             =   2145
         Width           =   4815
      End
      Begin VB.TextBox dato10 
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
         Left            =   1875
         MaxLength       =   50
         TabIndex        =   9
         Tag             =   "proveedor"
         Top             =   2115
         Width           =   4815
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
         Left            =   1875
         MaxLength       =   50
         TabIndex        =   7
         Tag             =   "proveedor"
         Top             =   1755
         Width           =   4815
      End
      Begin VB.TextBox dato7 
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
         Left            =   5940
         MaxLength       =   50
         TabIndex        =   6
         Tag             =   "proveedor"
         Top             =   1395
         Width           =   7305
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
         Left            =   1890
         MaxLength       =   9
         TabIndex        =   2
         Tag             =   "proveedor"
         Top             =   1395
         Width           =   1575
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
         Left            =   13680
         MaxLength       =   9
         TabIndex        =   12
         Tag             =   "proveedor"
         Top             =   2520
         Visible         =   0   'False
         Width           =   1455
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
         Left            =   13800
         MaxLength       =   2
         TabIndex        =   3
         Tag             =   "proveedor"
         Top             =   2130
         Visible         =   0   'False
         Width           =   375
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
         Left            =   14160
         MaxLength       =   2
         TabIndex        =   4
         Tag             =   "proveedor"
         Top             =   2130
         Visible         =   0   'False
         Width           =   375
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
         Left            =   14520
         MaxLength       =   4
         TabIndex        =   5
         Tag             =   "proveedor"
         Top             =   2130
         Visible         =   0   'False
         Width           =   615
      End
      Begin XPFrame.FrameXp frmDetalle 
         Height          =   4215
         Left            =   135
         TabIndex        =   28
         Top             =   2760
         Width           =   15090
         _ExtentX        =   26617
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
            Height          =   4185
            Left            =   0
            TabIndex        =   29
            Top             =   360
            Width           =   15015
            _ExtentX        =   26485
            _ExtentY        =   7382
            Cols            =   6
            DefaultFontSize =   6.75
            Rows            =   15
         End
      End
      Begin VB.TextBox pivote 
         Height          =   285
         Left            =   240
         TabIndex        =   34
         Text            =   "Text1"
         Top             =   -360
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Label lblnumeronota 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
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
         Left            =   7080
         TabIndex        =   60
         Top             =   7440
         Width           =   1815
      End
      Begin VB.Label Label11 
         BackColor       =   &H00C0C0C0&
         Caption         =   "NOTA CREDITO "
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
         Left            =   7680
         TabIndex        =   59
         Top             =   7200
         Width           =   1335
      End
      Begin VB.Label Label10 
         BackColor       =   &H000000FF&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   7080
         TabIndex        =   58
         Top             =   7200
         Width           =   615
      End
      Begin VB.Label LBLDESPACHO 
         BackColor       =   &H80000007&
         ForeColor       =   &H0000FF00&
         Height          =   330
         Left            =   12105
         TabIndex        =   47
         Top             =   540
         Width           =   2940
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackColor       =   &H00FF8080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "BODEGA DESPACHO"
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
         Left            =   12105
         TabIndex        =   46
         Top             =   270
         Width           =   2940
      End
      Begin VB.Label lbldespachados 
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   11400
         TabIndex        =   38
         Top             =   7920
         Width           =   1095
      End
      Begin VB.Label lbltotaldespachados 
         BackColor       =   &H00C0C0C0&
         Caption         =   "TOTAL DESPACHADOS :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   7515
         TabIndex        =   37
         Top             =   7920
         Width           =   3855
      End
      Begin VB.Label lblnombredocumento 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "GUIA DESPACHO FLETE"
         ForeColor       =   &H80000008&
         Height          =   165
         Left            =   3420
         TabIndex        =   36
         Top             =   -240
         Visible         =   0   'False
         Width           =   7215
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Giro"
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
         Left            =   135
         TabIndex        =   35
         Top             =   2475
         Width           =   1695
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Fono"
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
         Left            =   6855
         TabIndex        =   33
         Top             =   1755
         Width           =   1695
      End
      Begin ShockwaveFlashObjectsCtl.ShockwaveFlash opciones 
         Height          =   1395
         Left            =   135
         TabIndex        =   14
         Top             =   7065
         Visible         =   0   'False
         Width           =   6870
         _cx             =   12118
         _cy             =   2461
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
      Begin VB.Label lbl4 
         Appearance      =   0  'Flat
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Rut Cliente"
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
         Left            =   135
         TabIndex        =   32
         Top             =   1380
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
         Left            =   3600
         TabIndex        =   31
         Top             =   1395
         Width           =   375
      End
      Begin VB.Label lbl5 
         Appearance      =   0  'Flat
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Razón Social"
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
         Left            =   4125
         TabIndex        =   30
         Top             =   1395
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
         Left            =   9360
         TabIndex        =   22
         Top             =   2835
         Visible         =   0   'False
         Width           =   375
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
         Left            =   9810
         TabIndex        =   21
         Top             =   2835
         Visible         =   0   'False
         Width           =   3705
      End
      Begin VB.Label lbl8 
         Appearance      =   0  'Flat
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Ciudad"
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
         Left            =   6855
         TabIndex        =   20
         Top             =   2115
         Width           =   1695
      End
      Begin VB.Label lbl7 
         Appearance      =   0  'Flat
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Comuna"
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
         Left            =   135
         TabIndex        =   19
         Top             =   2115
         Width           =   1695
      End
      Begin VB.Label lbl3 
         Appearance      =   0  'Flat
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Fecha"
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
         Left            =   12735
         TabIndex        =   18
         Top             =   2130
         Visible         =   0   'False
         Width           =   1020
      End
      Begin VB.Label lbl6 
         Appearance      =   0  'Flat
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Dirección"
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
         Left            =   135
         TabIndex        =   17
         Top             =   1755
         Width           =   1695
      End
      Begin VB.Label lbl12 
         Appearance      =   0  'Flat
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Vendedor"
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
         Left            =   6795
         TabIndex        =   16
         Top             =   2925
         Visible         =   0   'False
         Width           =   1125
      End
   End
End
Attribute VB_Name = "Pguiadespacho"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
  Private modifica As Boolean
  Private lectura As Boolean
  Private tipoprecio As String
  Private fecha As String
  Private formatogrilla(11, 20) As String
  Private numeroDoc As String
  Private numerofle As String
  Private horaorden As String
  Private final As Double
  Private despachados As Double
  Private muestra As Double
  
  


Private Sub Command1_Click()
retorno

End Sub

Private Sub dato12_GotFocus()
 Principal.barraEstado.Panels(2).text = "F2: Ayuda Vendedor"
End Sub

Private Sub dato12_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = vbKeyF2 Then
            Call ayudaVendedores(dato12)
  End If
  
End Sub




Private Sub dato2_KeyPress(KeyAscii As Integer)
     KeyAscii = esNumero(KeyAscii)
        If KeyAscii = 13 Then
        
            dato2.text = ceros(dato2)
            If leerguia(dato2.text) = False Then
            dato6.SetFocus
            Else
            opciones.Visible = True
            opciones.SetFocus
            End If
            
        End If
   
End Sub

Private Sub dato20_GotFocus()
Call cargatexto(dato20)

End Sub

Private Sub dato20_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
If KeyAscii = 13 Then
    ncaja.SetFocus
End If

End Sub

Private Sub dato20_LostFocus()
    If dato20.text <> "FV" And dato20.text <> "BV" Then
        dato20.SetFocus
    End If
End Sub

Private Sub dato21_GotFocus()
Call cargatexto(dato21)
End Sub

Private Sub dato21_KeyPress(KeyAscii As Integer)
    KeyAscii = esNumero(KeyAscii)
    If KeyAscii = 13 Then
        dato21.text = Format(dato21.text, "0000000000")
        leerdetalleventas
    End If
End Sub
Private Sub dato23_LostFocus()
    If rubro <> leerRubro(dato23.text) Then
        dato23.text = empresaActiva
        lbllocal.Caption = leerNombreEmpresa(dato23.text)
    End If
End Sub

 Private Sub dato3_GotFocus()
        Call VerificarCajas(Me, dato3)
        Call selecciona(dato3)
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
    Private Sub dato13_GotFocus()
'        Call VerificarCajas(Me, dato13)
         Call selecciona(dato13)
    End Sub
  
  
     Private Sub dato3_KeyDown(KeyCode As Integer, Shift As Integer)
        Call Flechas(KeyCode, dato3)
    End Sub
    Private Sub dato23_KeyDown(KeyCode As Integer, Shift As Integer)
        If KeyCode = vbKeyF2 Then
           Call ayudaLocalesRubro(dato23)
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
   
    Private Sub dato23_KeyPress(KeyAscii As Integer)
            KeyAscii = esNumero(KeyAscii)
            If KeyAscii = 13 And dato23.text <> "" Then
                If rubro = leerRubro(dato23.text) Then
                    lbllocal.Caption = leerNombreEmpresa(dato23.text)
                    Call dato21_KeyPress(13)
                Else
                    dato23.SetFocus
                End If
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
              
               dato6.SetFocus
              
            End If
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
            dato6.SetFocus
        End If
    End Sub
    
    Private Sub dato6_KeyPress(KeyAscii As Integer)
        KeyAscii = esNumero(KeyAscii)
        If KeyAscii = 13 And dato6.text <> "" And Val(dato6.text) <> 0 Then
             dato6.text = ceros(dato6)
             lbldv.Caption = rut(dato6.text)
             rut_cliente = dato6.text
            
             muestra = 0
             If leerNombreCliente(dato6.text + lbldv.Caption) <> "" Then
                Call LeerClienteFlete(rut_cliente & lbldv.Caption, "0")
                dato3.text = Format(fechasistema, "dd")
                dato4.text = Format(fechasistema, "mm")
                dato5.text = Format(fechasistema, "yyyy")
                guardar.Visible = True
                leerdetalleventas
             Else
                dato6.SetFocus
             End If
'            SendKeys "{Tab}"
        End If
    End Sub
    
    Private Sub dato7_KeyPress(KeyAscii As Integer)
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If KeyAscii = 13 Then
            dato8.SetFocus
        End If
    End Sub
     Private Sub dato8_KeyPress(KeyAscii As Integer)
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If KeyAscii = 13 Then
            dato9.SetFocus
        End If
    End Sub
    Private Sub dato9_KeyPress(KeyAscii As Integer)
        KeyAscii = esNumero(KeyAscii)
        If KeyAscii = 13 Then
           dato10.SetFocus
        End If
    End Sub
     Private Sub dato10_KeyPress(KeyAscii As Integer)
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If KeyAscii = 13 Then
           dato11.SetFocus
        End If
    End Sub
     Private Sub dato11_KeyPress(KeyAscii As Integer)
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If KeyAscii = 13 Then
           dato13.SetFocus
        End If
    End Sub
    Private Sub dato12_KeyPress(KeyAscii As Integer)
       KeyAscii = esNumero(KeyAscii)
        If KeyAscii = 13 And dato12.text <> "" And Val(dato12.text) <> 0 Then
            dato12.text = ceros(dato12)
            lblDVV.Caption = rut(dato12.text)
'            rut_cliente = dato12.text
            lblvendedor.Caption = leerNombreVendedor(dato12.text & lblDVV.Caption)
        End If
End Sub
    Private Sub dato13_KeyPress(KeyAscii As Integer)
       KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If KeyAscii = 13 Then
           If dato6.text & lbldv.Caption <> "0000000019" Then
                guardar.Visible = True
                guardar.SetFocus
           End If
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

    Private Sub dato7_LostFocus()
        Call limpiaBarra(2)
    End Sub
    
    Private Sub dato9_LostFocus()
'        Call limpiaBarra(2)
    End Sub
    
    Private Sub dato10_LostFocus()
        Call limpiaBarra(2)
    End Sub
    
    
Sub leerdetalleventas()
   Dim csql As New rdoQuery
   Dim resultados As rdoResultset
   Dim tabla As String
   Dim linea As Double
   Dim DESPA As Double
   Dim cantidad  As Double
   Dim fechadespacho As String
  
    Call CARGAGRILLA(1, 14)
   despachados = 0
   Set csql.ActiveConnection = ventasRubro
   tabla = ""
   If Option2.Value = False Then
    tabla = "select dd.tipo,dc.foliosii,dd.linea,dd.fecha,dd.codigo,dd.descripcion,dd.cantidad,dd.precio,dd.total,dd.despachado,dd.numero,dc.caja,dc.rut,dc.local "
    tabla = tabla & "from " & baseVentas & dato23.text & ".sv_documento_detalle_" & dato23.text & " as dd inner join " & baseVentas & dato23.text & ".sv_documento_cabeza_" + dato23.text + " as dc on (dc.tipo=dd.tipo and dc.numero=dd.numero and dd.caja=dc.caja)  "
    tabla = tabla & " where  dc.rut='" + dato6.text + lbldv.Caption + "' and (dd.tipo='BV' OR dd.tipo='FV') and dd.tipodespacho>'01' and dc.caja='" + ncaja.text + "' "
    tabla = tabla & " order by dd.tipo,dd.numero,dd.linea"
   Else
    tabla = "select dd.tipo,dc.foliosii,dd.linea,dd.fecha,dd.codigo,dd.descripcion,dd.cantidad,dd.precio,dd.total,dd.despachado,dd.numero,dc.caja,dc.rut,dc.local "
    tabla = tabla & "from " & baseVentas & dato23.text & ".sv_documento_detalle_" & dato23.text & " as dd inner join " & baseVentas & dato23.text & ".sv_documento_cabeza_" + dato23.text + " as dc on (dc.tipo=dd.tipo and dc.numero=dd.numero and dd.caja=dc.caja)  "
    If Option3.Value = True Then tabla = tabla & " where dc.caja='" + ncaja.text + "' and  dc.TIPO='" + dato20.text + "' and dc.numero='" + dato21.text + "' and dd.tipodespacho>'01' "
    If Option4.Value = True Then tabla = tabla & " where dc.caja='" + ncaja.text + "' and  dc.TIPO='" + dato20.text + "' and dc.foliosii='" + dato21.text + "' and dd.tipodespacho>'01' "
    tabla = tabla & " order by dd.tipo,dd.numero,dd.linea"
   End If
   csql.sql = tabla
   csql.Execute
   detalle.Rows = 1
   linea = detalle.Rows
   If csql.RowsAffected > 0 Then
   Set resultados = csql.OpenResultset
   Call LeerClienteFlete(resultados(12), "0")
   foliofiscal.text = resultados(1)
   numerointerno.text = resultados(10)
   
  While Not resultados.EOF
        'RODRIGO ACA ESTA EL PROBLEMA CON LA FACTURA 336036
        fechadespacho = leerfechadespacho(resultados(0), resultados(10), resultados(4), resultados(6), "01", resultados(11))
        '0LINEA ORIGINAL GRANATE
        'cantidad = resultados(6) - leernotacredito(resultados(4), resultados(0), resultados(1), fechadespacho)
        
        'LINEA ARIEL
        cantidad = resultados(6) - leernotacredito(resultados(4), resultados(0), resultados(1), Format(fechasistema, "yyyy-mm-dd"))
        
        If cantidad > 0 Then
            detalle.Rows = detalle.Rows + 1
            If leernotacredito2(resultados(4), resultados(0), resultados(1)) <> 0 Then
                detalle.Range(detalle.Rows - 1, 1, detalle.Rows - 1, detalle.Cols - 1).BackColor = &HFF&
            End If
            DESPA = leerdespacho(resultados(0), resultados(10), resultados(4), resultados(6), "01", resultados(11), resultados(13)) + leerdespacho(resultados(0), resultados(10), resultados(4), resultados(6), "20", resultados(11), resultados(13))
            detalle.Cell(detalle.Rows - 1, 1).text = resultados(0)
            detalle.Cell(detalle.Rows - 1, 2).text = resultados(1)
            detalle.Cell(detalle.Rows - 1, 3).text = resultados(2)
            detalle.Cell(detalle.Rows - 1, 4).text = resultados(3)
            detalle.Cell(detalle.Rows - 1, 5).text = resultados(4)
            detalle.Cell(detalle.Rows - 1, 6).text = resultados(5)
            '  detalle.Cell(detalle.Rows - 1, 7).text = resultados(6)
            detalle.Cell(detalle.Rows - 1, 7).text = cantidad
            
            detalle.Cell(detalle.Rows - 1, 8).text = resultados(7)
            detalle.Cell(detalle.Rows - 1, 9).text = resultados(8)
'            DESPA = leerdespacho(resultados(0), resultados(10), resultados(4), resultados(6), "01", resultados(11)) + leerdespacho(resultados(0), resultados(10), resultados(4), resultados(6), "20", resultados(11))
'
            Rem DESPA = resultados(9)
            detalle.Cell(detalle.Rows - 1, 10).text = DESPA
            detalle.Cell(detalle.Rows - 1, 11).text = cantidad - DESPA
            detalle.Cell(detalle.Rows - 1, 12).text = "0"
            detalle.Cell(detalle.Rows - 1, 0).text = resultados(10)
            detalle.Cell(detalle.Rows - 1, 13).text = resultados(11)
  
        End If
        resultados.MoveNext
        linea = linea + 1
  Wend
   End If
   csql.Close
   Set resultados = Nothing
   Set csql = Nothing
 
    lbldespachados.Caption = despachados
    
    
   End Sub
 Private Sub CARGAGRILLA(ByVal row As Integer, ByVal col As Integer)
        Dim i As Integer
        Rem DATOS DE LA COLUMNA
        formatogrilla(1, 1) = "TD"
        formatogrilla(1, 2) = "NUMERO"
        formatogrilla(1, 3) = "LIN"
        formatogrilla(1, 4) = "FECHA "
        formatogrilla(1, 5) = "CODIGO"
        formatogrilla(1, 6) = "DESCRIPCION"
        formatogrilla(1, 7) = "CANTIDAD"
        formatogrilla(1, 8) = "PRECIO"
        formatogrilla(1, 9) = "TOTAL"
        formatogrilla(1, 10) = "DESPACHADO"
        formatogrilla(1, 11) = "PENDIENTE"
        formatogrilla(1, 12) = "DESPACHAR"
  
        
        Rem LARGO DE LOS DATOS
        formatogrilla(2, 1) = "10"
        formatogrilla(2, 2) = "10"
        formatogrilla(2, 3) = "10"
        formatogrilla(2, 4) = "10"
        formatogrilla(2, 5) = "10"
        formatogrilla(2, 6) = "40"
        formatogrilla(2, 7) = "10"
        formatogrilla(2, 8) = "10"
        formatogrilla(2, 9) = "10"
        formatogrilla(2, 10) = "10"
        formatogrilla(2, 11) = "10"
        formatogrilla(2, 12) = "10"
        
       
      
        
        
        Rem TIPO DE DATOS C=CEROS,S=STRING,N=NUMERICO,D=DATE,
        formatogrilla(3, 1) = "N"
        formatogrilla(3, 2) = "N"
        formatogrilla(3, 3) = "N"
        formatogrilla(3, 4) = "D"
        formatogrilla(3, 5) = "N"
        formatogrilla(3, 6) = "S"
        formatogrilla(3, 7) = "N"
        formatogrilla(3, 8) = "N"
        formatogrilla(3, 9) = "N"
        formatogrilla(3, 10) = "N"
        formatogrilla(3, 11) = "N"
        formatogrilla(3, 12) = "N"
       
              
        Rem FORMATO GRILLA
        formatogrilla(4, 1) = ""
        formatogrilla(4, 2) = ""
        formatogrilla(4, 3) = ""
        formatogrilla(4, 4) = ""
        formatogrilla(4, 5) = ""
        formatogrilla(4, 6) = ""
        formatogrilla(4, 7) = "###,###,##0.000"
        formatogrilla(4, 8) = "###,###,##0"
        formatogrilla(4, 9) = "###,###,##0"
        formatogrilla(4, 10) = "###,###,##0.000"
        formatogrilla(4, 11) = "###,###,##0.000"
        formatogrilla(4, 12) = "###,###,##0.000"
   
       
        Rem LOCCKED
        'FormatoGrilla(5, 0) = "TRUE"
        formatogrilla(5, 1) = "TRUE"
        formatogrilla(5, 2) = "TRUE"
        formatogrilla(5, 3) = "TRUE"
        formatogrilla(5, 4) = "TRUE"
        formatogrilla(5, 5) = "TRUE"
        formatogrilla(5, 6) = "TRUE"
        formatogrilla(5, 7) = "TRUE"
        formatogrilla(5, 8) = "TRUE"
        formatogrilla(5, 9) = "TRUE"
        formatogrilla(5, 10) = "TRUE"
        formatogrilla(5, 11) = "TRUE"
        formatogrilla(5, 12) = "FALSE"


      
       
        Rem VALOR MINIMO
        formatogrilla(6, 1) = ""
        formatogrilla(6, 2) = ""
        formatogrilla(6, 3) = ""
        formatogrilla(6, 4) = ""
        formatogrilla(6, 5) = ""
        formatogrilla(6, 6) = ""
        formatogrilla(6, 7) = ""
        formatogrilla(6, 8) = ""
        formatogrilla(6, 9) = ""
       
        
        Rem VALOR MAXIMO
        formatogrilla(7, 1) = ""
        formatogrilla(7, 2) = ""
        formatogrilla(7, 3) = ""
        formatogrilla(7, 4) = ""
        formatogrilla(7, 5) = ""
        formatogrilla(7, 6) = ""
        formatogrilla(7, 7) = ""
        formatogrilla(7, 8) = ""
        formatogrilla(7, 9) = ""
   
       
        Rem ANCHO
        formatogrilla(8, 1) = "2"
        formatogrilla(8, 2) = "8"
        formatogrilla(8, 3) = "3"
        formatogrilla(8, 4) = "7"
        formatogrilla(8, 5) = "10"
        formatogrilla(8, 6) = "30"
        formatogrilla(8, 7) = "7"
        formatogrilla(8, 8) = "7"
        formatogrilla(8, 9) = "0"
        formatogrilla(8, 10) = "8"
        formatogrilla(8, 11) = "8"
        formatogrilla(8, 12) = "7"
      
            
        detalle.Cols = col
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
        detalle.DefaultFont.Size = 9
'        detalle.Column(8).Width = 0
        detalle.Column(0).Width = 0
        detalle.Column(13).Width = 0
        
        
        detalle.Cell(0, 0).text = formatogrilla(1, 0)
        For i = 1 To col - 2
            detalle.Cell(0, i).text = formatogrilla(1, i)
            detalle.Column(i).Width = Val(formatogrilla(8, i)) * (detalle.Cell(0, i).Font.Size + 1.25)
            detalle.Column(i).MaxLength = Val(formatogrilla(2, i))
            detalle.Column(i).FormatString = formatogrilla(4, i)
            detalle.Column(i).Locked = formatogrilla(5, i)
            If formatogrilla(3, i) = "N" Then
                detalle.Column(i).Alignment = cellRightCenter
                If i <> 5 And i <> 3 Then
                    'detalle.Column(i).Mask = cellNumeric
                    'detalle.Column(i).Mask = cellValue
                     
                End If
            Else
                detalle.Column(i).Alignment = cellLeftCenter
                detalle.Column(i).Mask = cellUpper
            End If
        Next i
        detalle.Range(0, 0, 0, detalle.Cols - 1).Alignment = cellCenterCenter
'        detalle.AddItem vbTab & vbTab & "1" & vbTab & "0" & vbTab & "0" & vbTab & "0" & vbTab & "0" & vbTab & "0" & vbTab & "", True
        detalle.Enabled = True
'        detalle.ExtendLastCol = True
        
    
    End Sub
Sub LeerClienteFlete(rut, sucursal)


Dim op As Integer
Dim campos(10, 6) As String

    campos(0, 0) = "nombre"
    campos(1, 0) = "direccion"
    campos(2, 0) = "comuna"
    campos(3, 0) = "ciudad"
    campos(4, 0) = "fono1"
    campos(5, 0) = "giro"
    campos(6, 0) = "rut"
    campos(7, 0) = ""
    campos(0, 2) = "sv_maestroclientes"
    condicion = "rut='" & rut & "' and sucursal='" & sucursal & "'"
    op = 5
    Set sqlventas.conexion = ventas
    sqlventas.response = campos
    Call sqlventas.sqlventas(op, condicion)

    If sqlventas.Status = 0 Then
      dato6.text = Mid(sqlventas.response(6, 3), 1, 9)
      lbldv.Caption = Mid(sqlventas.response(6, 3), 10, 1)
      
      dato7.text = sqlventas.response(0, 3)
      dato8.text = sqlventas.response(1, 3)
      dato9.text = sqlventas.response(4, 3)
      dato10.text = sqlventas.response(2, 3)
      dato11.text = sqlventas.response(3, 3)
      dato13.text = sqlventas.response(5, 3)
      dato7.Locked = True
      dato8.Locked = True
      dato9.Locked = True
      dato10.Locked = True
      dato11.Locked = True
      dato13.Locked = True
      detalle.SetFocus
    
    Else
      dato7.SetFocus
    End If
    

End Sub

 

Private Sub detalle_DblClick()
    Dim csql As New rdoQuery
    Dim resultados As rdoResultset
    Set csql.ActiveConnection = gestion
    
    If detalle.Rows > 1 And detalle.ActiveCell.row > 0 Then
        If detalle.Cell(detalle.ActiveCell.row, 10).text > 0 Then
'            Call leerguiadespacho(detalle.Cell(detalle.ActiveCell.row, 1).text, detalle.Cell(detalle.ActiveCell.row, 2).text)
             Load cartoladespacho
             cartoladespacho.foliofiscal.text = foliofiscal.text
             cartoladespacho.numerointerno.text = numerointerno.text
             cartoladespacho.dato1.text = dato20.text
             cartoladespacho.dato2.text = dato21.text
             cartoladespacho.dato3.text = ncaja.text
             cartoladespacho.dato4.text = dato23.text
             cartoladespacho.dato5.text = Format(detalle.Cell(detalle.ActiveCell.row, 4).text, "dd")
             cartoladespacho.dato6.text = Format(detalle.Cell(detalle.ActiveCell.row, 4).text, "mm")
             cartoladespacho.dato7.text = Format(detalle.Cell(detalle.ActiveCell.row, 4).text, "yyyy")
             cartoladespacho.dato8.text = detalle.Cell(detalle.ActiveCell.row, 5).text
             cartoladespacho.dato9.text = detalle.Cell(detalle.ActiveCell.row, 6).text
             cartoladespacho.Show
             csql.sql = "select codigo from g_maestroempresas where "
             csql.sql = csql.sql & "rubro='" & rubro & "' order by codigo "
             csql.Execute
             If csql.RowsAffected > 0 Then
                Set resultados = csql.OpenResultset
                While Not resultados.EOF
                    Call cartoladespacho.cargadeafueraguias(resultados(0))
                    resultados.MoveNext
                Wend
             End If
        Else
            MsgBox "NO TIENE DESPACHO", vbCritical, "ATENCION"
        End If
    End If
End Sub

Private Sub detalle_KeyPress(KeyAscii As Integer)
    If KeyAscii <> 44 Then
       KeyAscii = esNumero(KeyAscii)
    End If



End Sub

Private Sub detalle_LeaveCell(ByVal row As Long, ByVal col As Long, NewRow As Long, NewCol As Long, Cancel As Boolean)
    NewCol = 12
    If NewRow <> row Then
        If detalle.Cell(row, 12).text = "" Then detalle.Cell(row, 12).text = "0"
        If CDbl(detalle.Cell(row, 12).text) > CDbl(detalle.Cell(row, 11).text) Then
            MsgBox ("cantidad excede despacho pendiente")
            NewRow = row
        End If
    End If

sumadespacho
End Sub
 
Sub GENERARGUIA()

End Sub

Private Sub FOLIO_GotFocus()
    If folio.text = "" Then
        folio.text = LEERULTIMAGUIA
    End If
    Call cargatexto(folio)


End Sub

Private Sub FOLIO_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then
        folio.Visible = False
        PANTALLA.Enabled = True
    End If

End Sub

Private Sub FOLIO_KeyPress(KeyAscii As Integer)
    KeyAscii = esNumero(KeyAscii)
    
    If KeyAscii = 13 Then
        folio.text = Format(folio.text, "0000000000")
        GENERA.SetFocus
    End If


End Sub
Public Function leerDocumentoexiste(ByVal NUMERO As String) As Boolean
        
        Dim op As Integer
        Dim campos(10, 10)
        campos(0, 0) = "numero"
        campos(1, 0) = ""
        campos(0, 2) = "sv_guia_despacho_entrega_" + empresaActiva
        condicion = " numero = '" & NUMERO & "'"
        op = 5
        sqlventas.response = campos
        Set sqlventas.conexion = ventasRubro
        Call sqlventas.sqlventas(op, condicion)
        If sqlventas.Status = 0 Then
            leerDocumentoexiste = True
         Else
            leerDocumentoexiste = False
        End If
    End Function

Private Sub Form_Load()
    Call CARGAGRILLA(1, 14)
    dato3.text = Format(fechasistema, "dd")
    dato4.text = Format(fechasistema, "mm")
    dato5.text = Format(fechasistema, "yyyy")
    docu.Visible = True
    dato23.text = empresaActiva
    lbllocal.Caption = leerNombreEmpresa(dato23.text)
    LBLDESPACHO.Caption = LEErBODEGADESPACHO(BODEGARETIRO)
    If LBLDESPACHO.Caption = "" Then
        MsgBox ("PUNTO NO HABILITADO PARA DESPACHOS DEBE INGRESAR EN CONFIGURACION ")
    End If
    FOLI.Visible = False

End Sub
Private Sub GENERA_Click()
If leerDocumentoexiste(folio.text) = True Then
    MsgBox ("DOCUMENTO YA ESTA CREADO")
    folio.text = ""
    folio.SetFocus
Else
    grabar
    Pdespachoflete.dato2.text = folio.text
    Pdespachoflete.Show
    Unload Me
End If



End Sub

Private Sub guardar_Click()
Dim K As Double
    Call sumadespacho
    If CDbl(lbldespachados.Caption) > 0 Then
        For K = 1 To detalle.Rows - 1
            If CDbl(detalle.Cell(K, 12).text) > CDbl(detalle.Cell(K, 11).text) And CDbl(detalle.Cell(K, 12).text) > 0 Then
                MsgBox "ESTA COMETIENDO UN ERROR CRITICO EN LA ENTREGA " & vbCrLf & "SUPERA LA ENTREGA PERMITIDA", vbCritical, "ATENCION"
                Unload Me
                Exit Sub
            End If
        Next K
        PANTALLA.Enabled = False
        FOLI.Visible = True
        folio.SetFocus
    Else
        MsgBox "DEBE AGREGAR UNA CANTIDAD DE DESPACHO", vbCritical, "ATENCION"
    End If
End Sub

 

 

Private Sub ncaja_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And ncaja.text <> "" Then
        ncaja.text = Format(ncaja.text, "00")
        dato21.Enabled = True
        dato21.SetFocus
    End If

End Sub

  Private Sub opciones_FSCommand(ByVal command As String, ByVal args As String)
        Select Case command
            Case "modifica"
'                Call modificar
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
'                Call anterior
            Case "siguiente"
'                Call siguiente
        End Select
    End Sub
 Sub imprimir()
    final = detalle.Rows
    
    detalle.Column(1).Locked = False
    detalle.Column(2).Locked = False
    detalle.Column(3).Locked = False
    detalle.Column(4).Locked = False
    detalle.Column(5).Locked = False
    detalle.Rows = detalle.Rows + 1
    detalle.Range(detalle.Rows - 1, 1, detalle.Rows - 1, 5).Merge
    detalle.Range(detalle.Rows - 1, 1, detalle.Rows - 1, 5).FontSize = 8
    detalle.Range(detalle.Rows - 1, 1, detalle.Rows - 1, 5).FontBold = True
    detalle.Range(detalle.Rows - 1, 1, detalle.Rows - 1, 5).Alignment = cellLeftGeneral
    detalle.Cell(detalle.Rows - 1, 1).text = " "
    
    detalle.Rows = detalle.Rows + 1
    detalle.Range(detalle.Rows - 1, 1, detalle.Rows - 1, 5).Merge
    detalle.Range(detalle.Rows - 1, 1, detalle.Rows - 1, 5).FontSize = 9
    detalle.Range(detalle.Rows - 1, 1, detalle.Rows - 1, 5).Alignment = cellCenterCenter
    detalle.Range(detalle.Rows - 1, 1, detalle.Rows - 1, 5).FontBold = True
    detalle.Range(detalle.Rows - 1, 1, detalle.Rows - 1, 2).Borders(cellEdgeTop) = cellThin
    detalle.Cell(detalle.Rows - 1, 1).text = "INFORMACION DE REPARTO  "
    
    detalle.Rows = detalle.Rows + 1
    detalle.Range(detalle.Rows - 1, 1, detalle.Rows - 1, 5).Merge
    detalle.Range(detalle.Rows - 1, 1, detalle.Rows - 1, 5).FontSize = 13
    detalle.Range(detalle.Rows - 1, 1, detalle.Rows - 1, 5).Alignment = cellCenterCenter
    detalle.Range(detalle.Rows - 1, 1, detalle.Rows - 1, 5).FontBold = True
    detalle.Cell(detalle.Rows - 1, 1).text = "Entregar entre las 08:00 y las 20:00  "
    
    detalle.Rows = detalle.Rows + 1
    detalle.Range(detalle.Rows - 1, 1, detalle.Rows - 1, 5).Merge
    detalle.Range(detalle.Rows - 1, 1, detalle.Rows - 1, 5).FontSize = 8
    detalle.Range(detalle.Rows - 1, 1, detalle.Rows - 1, 5).FontBold = True
    detalle.Range(detalle.Rows - 1, 1, detalle.Rows - 1, 5).Alignment = cellLeftGeneral
    detalle.Cell(detalle.Rows - 1, 1).text = " "
    
    detalle.Rows = detalle.Rows + 1
    detalle.Range(detalle.Rows - 1, 1, detalle.Rows - 1, 5).Merge
    detalle.Range(detalle.Rows - 1, 1, detalle.Rows - 1, 5).FontSize = 8
    detalle.Range(detalle.Rows - 1, 1, detalle.Rows - 1, 5).FontBold = False
    detalle.Range(detalle.Rows - 1, 1, detalle.Rows - 1, 5).Alignment = cellLeftCenter
    detalle.Cell(detalle.Rows - 1, 1).text = "Especificaciones de acceso al lugar "
    
    detalle.Rows = detalle.Rows + 1
    detalle.Range(detalle.Rows - 1, 1, detalle.Rows - 1, 5).Merge
    detalle.Range(detalle.Rows - 1, 1, detalle.Rows - 1, 5).FontSize = 8
    detalle.Range(detalle.Rows - 1, 1, detalle.Rows - 1, 5).FontBold = True
    detalle.Range(detalle.Rows - 1, 1, detalle.Rows - 1, 5).Alignment = cellLeftGeneral
    detalle.Cell(detalle.Rows - 1, 1).text = " "
    
    detalle.Rows = detalle.Rows + 1
    detalle.Range(detalle.Rows - 1, 1, detalle.Rows - 1, 5).Merge
    detalle.Range(detalle.Rows - 1, 1, detalle.Rows - 1, detalle.Cols - 1).Borders(cellEdgeTop) = cellThin
    detalle.Range(detalle.Rows - 1, 1, detalle.Rows - 1, detalle.Cols - 1).Borders(cellEdgeLeft) = cellThin
    detalle.Range(detalle.Rows - 1, 1, detalle.Rows - 1, detalle.Cols - 1).Borders(cellEdgeRight) = cellThin
    
    
    detalle.Rows = detalle.Rows + 1
    detalle.Range(detalle.Rows - 1, 1, detalle.Rows - 1, 5).Merge
    detalle.Range(detalle.Rows - 1, 1, detalle.Rows - 1, detalle.Cols - 1).Borders(cellEdgeLeft) = cellThin
    detalle.Range(detalle.Rows - 1, 1, detalle.Rows - 1, detalle.Cols - 1).Borders(cellEdgeRight) = cellThin
    
    detalle.Rows = detalle.Rows + 1
    detalle.Range(detalle.Rows - 1, 1, detalle.Rows - 1, 5).Merge
    detalle.Range(detalle.Rows - 1, 1, detalle.Rows - 1, detalle.Cols - 1).Borders(cellEdgeRight) = cellThin
    detalle.Range(detalle.Rows - 1, 1, detalle.Rows - 1, detalle.Cols - 1).Borders(cellEdgeLeft) = cellThin
    
    detalle.Rows = detalle.Rows + 1
    detalle.Range(detalle.Rows - 1, 1, detalle.Rows - 1, 5).Merge
    detalle.Range(detalle.Rows - 1, 1, detalle.Rows - 1, detalle.Cols - 1).Borders(cellEdgeLeft) = cellThin
    detalle.Range(detalle.Rows - 1, 1, detalle.Rows - 1, detalle.Cols - 1).Borders(cellEdgeRight) = cellThin
    
    detalle.Rows = detalle.Rows + 1
    detalle.Range(detalle.Rows - 1, 1, detalle.Rows - 1, 5).Merge
    detalle.Range(detalle.Rows - 1, 1, detalle.Rows - 1, detalle.Cols - 1).Borders(cellEdgeRight) = cellThin
    detalle.Range(detalle.Rows - 1, 1, detalle.Rows - 1, detalle.Cols - 1).Borders(cellEdgeLeft) = cellThin
    
    detalle.Rows = detalle.Rows + 1
    detalle.Range(detalle.Rows - 1, 1, detalle.Rows - 1, 5).Merge
    detalle.Range(detalle.Rows - 1, 1, detalle.Rows - 1, detalle.Cols - 1).Borders(cellEdgeLeft) = cellThin
    detalle.Range(detalle.Rows - 1, 1, detalle.Rows - 1, detalle.Cols - 1).Borders(cellEdgeRight) = cellThin
    
    detalle.Rows = detalle.Rows + 1
    detalle.Range(detalle.Rows - 1, 1, detalle.Rows - 1, 5).Merge
    detalle.Range(detalle.Rows - 1, 1, detalle.Rows - 1, detalle.Cols - 1).Borders(cellEdgeRight) = cellThin
    detalle.Range(detalle.Rows - 1, 1, detalle.Rows - 1, detalle.Cols - 1).Borders(cellEdgeLeft) = cellThin
    
    detalle.Rows = detalle.Rows + 1
    detalle.Range(detalle.Rows - 1, 1, detalle.Rows - 1, 5).Merge
    detalle.Range(detalle.Rows - 1, 1, detalle.Rows - 1, detalle.Cols - 1).Borders(cellEdgeBottom) = cellThin
    detalle.Range(detalle.Rows - 1, 1, detalle.Rows - 1, detalle.Cols - 1).Borders(cellEdgeRight) = cellThin
    detalle.Range(detalle.Rows - 1, 1, detalle.Rows - 1, detalle.Cols - 1).Borders(cellEdgeLeft) = cellThin
    
    Call Titulos("ORDEN DE FLETE")
    detalle.PageSetup.HeaderMargin = 0.5
    detalle.PageSetup.PrintFixedRow = True
    detalle.PageSetup.TopMargin = 1
    detalle.PageSetup.LeftMargin = 1
    detalle.PageSetup.RightMargin = 1
    detalle.PageSetup.BottomMargin = 3
    detalle.PageSetup.FooterMargin = 2
    detalle.PrintPreview
    detalle.Rows = final
    
 End Sub
    Sub retorno()
        dato2.text = ""
        dato3.text = ""
        dato4.text = ""
        dato5.text = ""
        dato6.text = ""
        dato8.text = ""
        dato9.text = ""
        dato10.text = ""
        dato7.text = ""
        dato11.text = ""
        dato12.text = ""
        dato13.text = ""
        dato23.text = empresaActiva
        lbllocal.Caption = leerNombreEmpresa(dato23.text)
        dato20.text = ""
        dato21.text = ""
        lblnumeronota.Caption = ""
        ncaja.text = ""
        detalle.Rows = 1
        lbldv.Caption = ""
        lblDVV.Caption = ""
        lblvendedor.Caption = ""
        Call CARGAGRILLA(1, 14)
        guardar.Visible = True
        despachados = 0
        lbldespachados.Caption = despachados
        If Option1.Value = True Then
            dato6.SetFocus
        Else
            dato20.SetFocus
        End If
    
    End Sub
    
    Sub Titulos(titulo1)

    Dim i As Integer
    Dim objReportTitle As FlexCell.ReportTitle
    
        detalle.FixedRowColStyle = Fixed3D
        detalle.CellBorderColorFixed = vbButtonShadow
        detalle.ShowResizeTips = False
        detalle.ReportTitles.Clear
        detalle.PageSetup.CenterHorizontally = True
        detalle.PageSetup.Orientation = cellPortrait
        detalle.PageSetup.BlackAndWhite = True
        detalle.Column(1).Width = 90
        detalle.Column(2).Width = 250
        detalle.Column(3).Width = 50
        detalle.Column(4).Width = 80
        detalle.Column(5).Width = 85
        detalle.PageSetup.PrintTitleRows = 1
    
    'Logo
'    Grid1.Images.Add App.path & "\Admin.gif", "Logo"
'    Set objReportTitle = New FlexCell.ReportTitle
'    objReportTitle.ImageKey = "Logo"
'    objReportTitle.Align = CellLeft
'    Grid1.ReportTitles.Add objReportTitle
    
    'ENCABEZADO DE PAGINA
        detalle.PageSetup.Header = nombreempresa & vbCrLf & direccionempresa & vbCrLf & comunaempresa
        detalle.PageSetup.HeaderAlignment = cellLeft
        detalle.PageSetup.HeaderFont.Name = "Verdana"
        detalle.PageSetup.HeaderFont.Size = 8
    
    'TITULOS DEL REPORTE
        Set objReportTitle = New FlexCell.ReportTitle
        objReportTitle.text = titulo1
        objReportTitle.Font.Name = "Verdana" '"Times New Roman"
        objReportTitle.Font.Size = 10
        objReportTitle.Font.Bold = True
        objReportTitle.Align = cellCenter
        objReportTitle.PrintOnAllPages = True
        detalle.ReportTitles.Add objReportTitle
    
        Set objReportTitle = New FlexCell.ReportTitle
        objReportTitle.text = "NUMERO  :  " & numerofle
        objReportTitle.Font.Name = "Verdana" '"Times New Roman"
        objReportTitle.Font.Size = 8
        objReportTitle.Font.Bold = True
        objReportTitle.Align = cellRight
        objReportTitle.PrintOnAllPages = True
        detalle.ReportTitles.Add objReportTitle
    
'    Set objReportTitle = New FlexCell.ReportTitle
'    objReportTitle.text = "EMITIDO  :  " & Format(fechasistema, "dd-mm-yyyy") & "    HORA  :" & horaorden
'    objReportTitle.Font.Name = "Verdana" '"Times New Roman"
'    objReportTitle.Font.Size = 8
'    objReportTitle.Font.Bold = True
'    objReportTitle.Align = cellRight
'    objReportTitle.PrintOnAllPages = True
'    detalle.ReportTitles.Add objReportTitle
    
        Set objReportTitle = New FlexCell.ReportTitle
        objReportTitle.text = "DOCUMENTO :  " & dato2.text
        objReportTitle.Font.Name = "Verdana" '"Times New Roman"
        objReportTitle.Font.Size = 8
        objReportTitle.Font.Bold = True
        objReportTitle.Align = cellLeft
        objReportTitle.PrintOnAllPages = True
        detalle.ReportTitles.Add objReportTitle

        
        Set objReportTitle = New FlexCell.ReportTitle
        objReportTitle.text = "RUT  :  " & dato6.text & "-" & lbldv.Caption
        objReportTitle.Font.Name = "Verdana" '"Times New Roman"
        objReportTitle.Font.Size = 8
        objReportTitle.Font.Bold = True
        objReportTitle.Align = cellLeft
        objReportTitle.PrintOnAllPages = True
        detalle.ReportTitles.Add objReportTitle
    
        Set objReportTitle = New FlexCell.ReportTitle
        objReportTitle.text = "CLIENTE  :  " & dato7.text
        objReportTitle.Font.Name = "Verdana" '"Times New Roman"
        objReportTitle.Font.Size = 8
        objReportTitle.Font.Bold = True
        objReportTitle.Align = cellLeft
        objReportTitle.PrintOnAllPages = True
        detalle.ReportTitles.Add objReportTitle
    
    
    
        Set objReportTitle = New FlexCell.ReportTitle
        objReportTitle.text = "DIRECCION  :  " & dato8.text
        objReportTitle.Font.Name = "Verdana" '"Times New Roman"
        objReportTitle.Font.Size = 8
        objReportTitle.Font.Bold = True
        objReportTitle.Align = cellLeft
        objReportTitle.PrintOnAllPages = True
        detalle.ReportTitles.Add objReportTitle
    
        Set objReportTitle = New FlexCell.ReportTitle
        objReportTitle.text = "FONO  :  " & dato9.text
        objReportTitle.Font.Name = "Verdana" '"Times New Roman"
        objReportTitle.Font.Size = 8
        objReportTitle.Font.Bold = True
        objReportTitle.Align = cellLeft
        objReportTitle.PrintOnAllPages = True
        detalle.ReportTitles.Add objReportTitle
    
        Set objReportTitle = New FlexCell.ReportTitle
        objReportTitle.text = "CIUDAD  :  " & dato10.text & "       CONTACTO  :  " & dato7.text
        objReportTitle.Font.Name = "Verdana" '"Times New Roman"
        objReportTitle.Font.Size = 8
        objReportTitle.Font.Bold = True
        objReportTitle.Align = cellLeft
        objReportTitle.PrintOnAllPages = True
        detalle.ReportTitles.Add objReportTitle
        Set objReportTitle = New FlexCell.ReportTitle
        objReportTitle.text = "CONTACTO  :  " & dato7.text
        objReportTitle.Font.Name = "Verdana" '"Times New Roman"
        objReportTitle.Font.Size = 8
        objReportTitle.Font.Bold = True
        objReportTitle.Align = cellLeft
        objReportTitle.PrintOnAllPages = True
        detalle.ReportTitles.Add objReportTitle
        
        Set objReportTitle = New FlexCell.ReportTitle
        objReportTitle.text = "VENDEDOR :  " & dato12.text & "-" & lblDVV.Caption & "     " & lblvendedor.Caption
        objReportTitle.Font.Name = "Verdana" '"Times New Roman"
        objReportTitle.Font.Size = 8
        objReportTitle.Font.Bold = True
        objReportTitle.Align = cellLeft
        objReportTitle.PrintOnAllPages = True
        detalle.ReportTitles.Add objReportTitle
    
        'PIE DE PAGINA
         detalle.PageSetup.Footer = "Pág &P de &N" & vbCrLf & "Fecha: &D" & vbCrLf & " Hora: " & Time & vbCrLf & "Usuario: " & usuarioSistema
         detalle.PageSetup.FooterAlignment = cellRight
         detalle.PageSetup.FooterFont.Name = "Verdana"
         detalle.PageSetup.FooterFont.Size = 7
         detalle.Range(0, 1, 0, detalle.Cols - 1).Borders(cellEdgeTop) = cellThin
         detalle.Range(0, 1, 0, detalle.Cols - 1).Borders(cellEdgeBottom) = cellThin
         detalle.Range(0, 1, 0, detalle.Cols - 1).Borders(cellEdgeLeft) = cellThin
         detalle.Range(0, 1, 0, detalle.Cols - 1).Borders(cellEdgeRight) = cellThin
         detalle.Range(0, 1, 0, detalle.Cols - 1).Borders(cellInsideHorizontal) = cellThin
         detalle.Range(0, 1, 0, detalle.Cols - 1).Borders(cellInsideVertical) = cellThin
         detalle.Range(0, 1, 0, detalle.Cols - 1).FontBold = True
    
End Sub

 Function LEERULTIMAGUIA() As String
  
    Dim csql As New rdoQuery
    Dim resultados As rdoResultset
    
    Set csql.ActiveConnection = ventasRubro
    
    csql.sql = " select IFNULL(MAX(numero) + 1,'0000000001')"
    csql.sql = csql.sql & " from sv_guia_despacho_entrega_" & empresaActiva
    csql.Execute
    If csql.RowsAffected > 0 Then
    Set resultados = csql.OpenResultset
              If resultados(0) <> "" And resultados(0) <> "0" Then
                 LEERULTIMAGUIA = resultados(0)
               Else
                 LEERULTIMAGUIA = "0000000001"
              End If
    End If
    Set resultados = Nothing
    csql.Close
    
End Function
Function grabar() As Boolean
Dim csql As New rdoQuery
Dim tabla As String
Dim K As Integer
Dim j As Double
Dim MONTO As Double

j = 0
numerofle = folio.text
pivote.MaxLength = 10
pivote.text = numerofle
numerofle = ceros(pivote)
Set csql.ActiveConnection = ventasRubro
horaorden = Time

MONTO = 0
For K = 1 To detalle.Rows - 1
tabla = ""
If detalle.Cell(K, 12).text <> "" And detalle.Cell(K, 12).text <> "0" Then
j = j + 1
If empresaActiva <> dato23.text Then
Call grabarDetalle("EL", numerofle, K, Format(fechasistema, "yyyy-mm-dd"), detalle.Cell(K, 5).text, detalle.Cell(K, 6).text, detalle.Cell(K, 12).text, detalle.Cell(K, 8).text, "0", CDbl(detalle.Cell(K, 12).text) * CDbl(detalle.Cell(K, 8).text), "00", "0", empresaActiva)
Call grabarDetalle("RL", numerofle, K, Format(fechasistema, "yyyy-mm-dd"), detalle.Cell(K, 5).text, detalle.Cell(K, 6).text, detalle.Cell(K, 12).text, detalle.Cell(K, 8).text, "0", CDbl(detalle.Cell(K, 12).text) * CDbl(detalle.Cell(K, 8).text), "00", "0", dato23.text)

End If

Call grabardespacho(numerofle, j, dato6.text + lbldv.Caption, Format(fechasistema, "yyyy-mm-dd"), detalle.Cell(K, 5).text, detalle.Cell(K, 6).text, detalle.Cell(K, 12).text, detalle.Cell(K, 8).text, detalle.Cell(K, 9).text, detalle.Cell(K, 1).text, detalle.Cell(K, 0).text, detalle.Cell(K, 3).text, detalle.Cell(K, 13).text, dato23.text)
Call modificardespacho(detalle.Cell(K, 1).text, detalle.Cell(K, 0).text, detalle.Cell(K, 3).text, detalle.Cell(K, 12).text, detalle.Cell(K, 13).text)

End If
Next K

If empresaActiva <> dato23.text Then

Call grabarCabeza("EL", numerofle, Format(fechasistema, "yyyy-mm-dd"), "00", empresaActiva, dato23.text, MONTO, empresaActiva)
Call grabarCabeza("RL", numerofle, Format(fechasistema, "yyyy-mm-dd"), "00", empresaActiva, empresaActiva, MONTO, dato23.text)
End If

csql.Close
Set csql = Nothing

grabar = True
End Function

Public Sub modificardespacho(TIPO, NUMERO, linea, cantidad, caja)
Dim csql As New rdoQuery
Dim tabla As String
Set csql.ActiveConnection = ventasRubro
cantidad = Replace(cantidad, ",", ".")
tabla = "update " + baseVentas + dato23.text + ".sv_documento_detalle_" & dato23.text & " set despachado=despachado+'" & cantidad & "' "
tabla = tabla & "where tipo='" & TIPO & "' and numero='" & NUMERO & "' and linea='" & linea & "' and caja='" + caja + "' "
csql.sql = tabla
csql.Execute
    Call sincronizadatos(csql.sql, ventasRubro)
csql.Close
Set csql = Nothing
End Sub


    
    
Sub ELIMINAR()

End Sub

Sub sumadespacho()
Dim K As Integer
despachados = 0
For K = 1 To detalle.Rows - 1
despachados = despachados + CDbl(detalle.Cell(K, 12).text)
Next K
lbldespachados.Caption = despachados

End Sub
Private Function leerguia(NUMERO) As Boolean
  Dim csql As New rdoQuery
  Dim resultados As rdoResultset
  Dim linea As Double
  
    Set csql.ActiveConnection = ventasRubro
    csql.sql = "select numero,linea,rut,fecha,codigo,descripcion,despachado,numeroflete from sv_guia_despacho_flete_" & empresaActiva
    csql.sql = csql.sql & " where  numero='" & NUMERO & "'"
    csql.Execute
    
    If csql.RowsAffected > 0 Then
    detalle.Rows = csql.RowsAffected + 1
    Set resultados = csql.OpenResultset
    leerguia = True
    dato6.text = resultados("rut")
    dato3.text = Format(resultados("fecha"), "dd")
    dato4.text = Format(resultados("fecha"), "mm")
    dato5.text = Format(resultados("fecha"), "yyyy")
    While Not resultados.EOF
    detalle.Cell(resultados("linea"), 1).text = resultados("numeroflete")
    detalle.Cell(resultados("linea"), 2).text = leerfechaflete(resultados("numeroflete"))
    detalle.Cell(resultados("linea"), 3).text = resultados("codigo")
    detalle.Cell(resultados("linea"), 4).text = resultados("descripcion")
    detalle.Cell(resultados("linea"), 5).text = resultados("despachado")
    resultados.MoveNext
    Wend
   
    Else
    leerguia = False
    
    End If
    Set resultados = Nothing
    csql.Close
    End Function
Function leerfechaflete(NUMERO) As String
    Dim csql As New rdoQuery
    Dim resultado As rdoResultset
        
    Set csql.ActiveConnection = ventasRubro
    csql.sql = "select fecha from sv_fletes_cabeza_" & empresaActiva
    csql.sql = csql.sql & " where numero='" & NUMERO & "' "
    csql.Execute
    If csql.RowsAffected > 0 Then
    Set resultado = csql.OpenResultset
       leerfechaflete = resultado(0)
    End If
    csql.Close
    Set resultado = Nothing
    Set csql = Nothing
    
    End Function

Private Sub ver_Click()
muestra = 1

End Sub

Private Sub Option1_Click()

dato6.SetFocus
docu.Enabled = False
End Sub

Private Sub Option2_Click()
dato20.text = ""
dato21.text = ""
docu.Enabled = True

dato20.SetFocus

End Sub

Private Function leernotacredito(CODIGO, TIPO, NUMERO, fecha) As Double
  Dim csql As New rdoQuery
    Dim resultado As rdoResultset
        
    Set csql.ActiveConnection = gestionRubro
    csql.sql = "select dd.cantidad from " & baseVentas & dato23.text & ".sv_documento_detalle_" & dato23.text & " as dd "
    csql.sql = csql.sql & " where  dd.numerodocumento='" & NUMERO & "' and tipodocumento='" & TIPO & "'  and codigo='" & CODIGO & "' and fecha< '" & fecha & "' "
    csql.Execute
    leernotacredito = 0
    If csql.RowsAffected > 0 Then
    
    Set resultado = csql.OpenResultset
       leernotacredito = resultado(0)
    End If
    csql.Close
    Set resultado = Nothing
    Set csql = Nothing

End Function
Private Function leernotacredito2(CODIGO, TIPO, NUMERO) As Double
  Dim csql As New rdoQuery
    Dim resultado As rdoResultset
        
    Set csql.ActiveConnection = gestionRubro
    csql.sql = "select dd.numero from " & baseVentas & dato23.text & ".sv_documento_detalle_" & dato23.text & " as dd "
    csql.sql = csql.sql & " where  dd.numerodocumento='" & NUMERO & "' and tipodocumento='" & TIPO & "'  and codigo='" & CODIGO & "' "
    csql.Execute
    leernotacredito2 = 0
    If csql.RowsAffected > 0 Then
    
    Set resultado = csql.OpenResultset
       leernotacredito2 = resultado(0)
       lblnumeronota.Caption = resultado(0)
    End If
    csql.Close
    Set resultado = Nothing
    Set csql = Nothing

End Function

 Public Sub grabardespacho(NUMERO, linea, rut, fecha, CODIGO, descripcion, cantidad, precio, total, tipodo, numerodo, lineado, caja, loc)
        
        Dim campos(14, 3) As String
        Dim op As Integer
        Dim K As Integer
       
        Set sql = New sqlventas.sqlventa
        campos(0, 0) = "numero"
        campos(1, 0) = "linea"
        campos(2, 0) = "rut"
        campos(3, 0) = "fecha"
        campos(4, 0) = "codigo"
        campos(5, 0) = "descripcion"
        campos(6, 0) = "cantidad"
        campos(7, 0) = "precio"
        campos(8, 0) = "total"
        campos(9, 0) = "tipodocumento"
        campos(10, 0) = "numerodocumento"
        campos(11, 0) = "lineadocumento"
        campos(12, 0) = "cajadocumento"
        campos(13, 0) = "localdocumento"
        campos(14, 0) = ""
        
        campos(0, 1) = NUMERO
        campos(1, 1) = linea
        campos(2, 1) = rut
        campos(3, 1) = fecha
        campos(4, 1) = CODIGO
        campos(5, 1) = descripcion
        campos(6, 1) = Replace(cantidad, ",", ".")
        campos(7, 1) = precio
        campos(8, 1) = CDbl(cantidad) * CDbl(precio)
        campos(9, 1) = tipodo
        campos(10, 1) = numerodo
        campos(11, 1) = lineado
        campos(12, 1) = caja
        campos(13, 1) = loc
        
        campos(0, 2) = "sv_guia_despacho_entrega_" + empresaActiva
        condicion = ""
        op = 2
        sql.response = campos
        Set sql.conexion = ventasRubro
        sql.audit = True: sql.programaactivo = Me.Caption
        Set sql.conauditoria = conauditoria: sql.usuarioauditoria = usuarioSistema
        Call sql.sqlventas(op, condicion)
     
    End Sub


    Private Sub grabarCabeza(TIPO, NUMERO, fecha, ORIGEN, localorigen, localdestino, MONTO, empresa)
        
        Dim op As Integer
        Dim campos(10, 10) As String
        
        campos(0, 0) = "tipo"
        campos(1, 0) = "numero"
        campos(2, 0) = "fecha"
        campos(3, 0) = "origen"
        campos(4, 0) = "localorigen"
        campos(5, 0) = "localdestino"
        campos(6, 0) = "monto"
        campos(7, 0) = ""
        
        campos(0, 1) = TIPO
        campos(1, 1) = NUMERO
        campos(2, 1) = fecha
        campos(3, 1) = ORIGEN
        campos(4, 1) = localorigen
        campos(5, 1) = localdestino
        campos(6, 1) = MONTO
        campos(0, 2) = "l_movimientos_cabeza_" & empresa
        condicion = ""
        op = 2
        sqlventas.response = campos
        Set sqlventas.conexion = gestionRubro
        sqlventas.audit = True: sqlventas.programaactivo = Me.Caption
        Set sqlventas.conauditoria = conauditoria: sqlventas.usuarioauditoria = usuarioSistema
        Call sqlventas.sqlventas(op, condicion)
        
    End Sub
    
    Private Sub grabarDetalle(TIPO, NUMERO, linea, fecha, CODIGO, descripcion, unidades, precio, Descuento, total, bodega, stocktransaccion, empresa)
        
        Dim op As Integer
        Dim campos(20, 10) As String
         
        campos(0, 0) = "tipo"
        campos(1, 0) = "numero"
        campos(2, 0) = "linea"
        campos(3, 0) = "fecha"
        campos(4, 0) = "codigo"
        campos(5, 0) = "descripcion"
        campos(6, 0) = "unidades"
        campos(7, 0) = "precio"
        campos(8, 0) = "descuento"
        campos(9, 0) = "total"
        campos(10, 0) = "bodega"
        campos(11, 0) = "stocktransaccion"
        campos(12, 0) = ""
        
        campos(0, 1) = TIPO
        campos(1, 1) = NUMERO
        campos(2, 1) = linea
        campos(3, 1) = fecha
        campos(4, 1) = CODIGO
        campos(5, 1) = descripcion
        campos(6, 1) = unidades
        campos(7, 1) = precio
        campos(8, 1) = Descuento
        campos(9, 1) = total
        campos(10, 1) = bodega
        campos(11, 1) = stocktransaccion
        campos(0, 2) = "l_movimientos_detalle_" & empresa
            
        condicion = ""
        op = 2
        Set sqlventas.conexion = gestionRubro
        sqlventas.audit = True: sqlventas.programaactivo = Me.Caption
        Set sqlventas.conauditoria = conauditoria: sqlventas.usuarioauditoria = usuarioSistema
            sqlventas.response = campos
            Call sqlventas.sqlventas(op, condicion)
            If TIPO = "EL" Then
             Call actualiza_stock("-", CODIGO, "N", "N", "00", Format(fecha, "YYYY"), unidades, precio, fecha, "0", empresaActiva)
            Else
             Call actualiza_stock("+", CODIGO, "N", "N", "00", Format(fecha, "YYYY"), unidades, precio, fecha, "0", dato23.text)
            End If
        
    End Sub

Function leerdespacho(TIPO, NUMERO, CODIGO, cantidad, loc, caja, LOCALdocumento) As Double
    Dim csql As New rdoQuery
    Dim resultados As rdoResultset
    Dim tabla As String
    
    Set csql.ActiveConnection = ventasRubro
    tabla = "select ifnull(sum(cantidad),'0') from " + clientesistema + "ventas" + loc + ".sv_guia_despacho_entrega_" & loc & "  "
    tabla = tabla & "where tipodocumento='" & TIPO & "' and numerodocumento='" & NUMERO & "' and codigo='" & CODIGO & "' and localdocumento ='" & LOCALdocumento & "' and cajadocumento='" + caja + "' "
    csql.sql = tabla
    csql.Execute
    leerdespacho = 0
    If csql.RowsAffected > 0 Then
        Set resultados = csql.OpenResultset
        leerdespacho = Replace(resultados(0), ".", ",")
    End If

End Function



Function leerfechadespacho(TIPO, NUMERO, CODIGO, cantidad, loc, caja) As String
    Dim csql As New rdoQuery
    Dim resultados As rdoResultset
    Dim tabla As String
 
            Set csql.ActiveConnection = ventasRubro
            tabla = "select ifnull(fecha,'') from " + clientesistema + "ventas" + loc + ".sv_guia_despacho_entrega_" & loc & "  "
            tabla = tabla & "where tipodocumento='" & TIPO & "' and numerodocumento='" & NUMERO & "' and codigo='" & CODIGO & "' and localdocumento ='" & loc & "' and cajadocumento='" + caja + "' "
            csql.sql = tabla
            csql.Execute
            leerfechadespacho = ""
            If csql.RowsAffected > 0 Then
                Set resultados = csql.OpenResultset
                leerfechadespacho = resultados(0)
            End If
           
End Function
 
