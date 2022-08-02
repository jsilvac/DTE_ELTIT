VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash10c.ocx"
Object = "{4777436C-EB8C-4596-98A8-EBCF98683969}#1.0#0"; "FlexCell.ocx"
Object = "{ADBBAED6-B16F-40DC-80DF-B44910CBA76C}#1.0#0"; "Frame-Xp.ocx"
Begin VB.Form Pdespachoflete 
   Caption         =   "Guia de Despacho Fletes"
   ClientHeight    =   8655
   ClientLeft      =   60
   ClientTop       =   465
   ClientWidth     =   13590
   LinkTopic       =   "Form1"
   ScaleHeight     =   8655
   ScaleWidth      =   13590
   StartUpPosition =   2  'CenterScreen
   Begin XPFrame.FrameXp FrameXp1 
      Height          =   8595
      Left            =   0
      TabIndex        =   12
      Top             =   0
      Width           =   15255
      _ExtentX        =   26908
      _ExtentY        =   15161
      BackColor       =   12648384
      Caption         =   "Datos Documento"
      CaptionEstilo3D =   2
      BackColor       =   12648384
      ForeColor       =   8438015
      BordeColor      =   16744576
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
         Left            =   8820
         Locked          =   -1  'True
         MaxLength       =   2
         TabIndex        =   42
         Tag             =   "proveedor"
         Top             =   270
         Width           =   375
      End
      Begin FlexCell.Grid Grid1 
         Height          =   135
         Left            =   600
         TabIndex        =   41
         Top             =   0
         Visible         =   0   'False
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   238
         Cols            =   5
         DefaultFontSize =   8.25
         DefaultRowHeight=   14
         Rows            =   30
      End
      Begin XPFrame.FrameXp FRMIMPRESION 
         Height          =   1635
         Left            =   4440
         TabIndex        =   36
         Top             =   5280
         Visible         =   0   'False
         Width           =   5820
         _ExtentX        =   10266
         _ExtentY        =   2884
         BackColor       =   8454016
         Caption         =   "IMPRESION DE DOCUMENTOS"
         CaptionEstilo3D =   1
         BackColor       =   8454016
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
         Begin VB.OptionButton Option1 
            BackColor       =   &H0080FF80&
            Caption         =   "IMPRIME GUIA CORTA"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   510
            Left            =   180
            TabIndex        =   40
            Top             =   405
            Width           =   3480
         End
         Begin VB.OptionButton Option2 
            BackColor       =   &H0080FF80&
            Caption         =   "IMPRIME GUIA LARGA"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   465
            Left            =   180
            TabIndex        =   39
            Top             =   945
            Value           =   -1  'True
            Width           =   3480
         End
         Begin VB.CommandButton Command1 
            BackColor       =   &H00FF8080&
            Caption         =   "IMPRIME"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   4140
            Style           =   1  'Graphical
            TabIndex        =   38
            Top             =   975
            Width           =   1590
         End
         Begin VB.CommandButton Command2 
            BackColor       =   &H00FF8080&
            Caption         =   "RETORNO"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   4140
            Style           =   1  'Graphical
            TabIndex        =   37
            Top             =   525
            Width           =   1590
         End
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
         Left            =   1740
         MaxLength       =   10
         TabIndex        =   33
         Tag             =   "proveedor"
         Top             =   600
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
         Left            =   1680
         MaxLength       =   50
         TabIndex        =   9
         Tag             =   "proveedor"
         Top             =   2040
         Width           =   11655
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
         Left            =   8520
         MaxLength       =   10
         TabIndex        =   6
         Tag             =   "proveedor"
         Top             =   1320
         Width           =   4815
      End
      Begin XPFrame.FrameXp frmTipo 
         Height          =   1200
         Left            =   6720
         TabIndex        =   20
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
            TabIndex        =   24
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
            TabIndex        =   23
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
            TabIndex        =   22
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
            TabIndex        =   21
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
         Left            =   8520
         MaxLength       =   50
         TabIndex        =   8
         Tag             =   "proveedor"
         Top             =   1680
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
         Left            =   1740
         MaxLength       =   50
         TabIndex        =   7
         Tag             =   "proveedor"
         Top             =   1680
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
         Left            =   1740
         MaxLength       =   50
         TabIndex        =   5
         Tag             =   "proveedor"
         Top             =   1320
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
         Left            =   5400
         MaxLength       =   50
         TabIndex        =   4
         Tag             =   "proveedor"
         Top             =   960
         Width           =   7935
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
         Left            =   1740
         MaxLength       =   9
         TabIndex        =   3
         Tag             =   "proveedor"
         Top             =   960
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
         Left            =   7635
         MaxLength       =   9
         TabIndex        =   10
         Tag             =   "proveedor"
         Top             =   2040
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
         Left            =   12000
         MaxLength       =   2
         TabIndex        =   0
         Tag             =   "proveedor"
         Top             =   600
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
         Left            =   12360
         MaxLength       =   2
         TabIndex        =   1
         Tag             =   "proveedor"
         Top             =   600
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
         Left            =   12720
         MaxLength       =   4
         TabIndex        =   2
         Tag             =   "proveedor"
         Top             =   600
         Width           =   615
      End
      Begin XPFrame.FrameXp frmDetalle 
         Height          =   4575
         Left            =   120
         TabIndex        =   25
         Top             =   2400
         Width           =   13335
         _ExtentX        =   23521
         _ExtentY        =   8070
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
            TabIndex        =   26
            Top             =   360
            Width           =   13215
            _ExtentX        =   23310
            _ExtentY        =   7382
            Cols            =   6
            DefaultFontSize =   8.25
            Rows            =   15
         End
      End
      Begin VB.TextBox pivote 
         Height          =   285
         Left            =   240
         TabIndex        =   31
         Text            =   "Text1"
         Top             =   -360
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Label lbllocal 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "   "
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
         Left            =   9240
         TabIndex        =   45
         Top             =   270
         Width           =   3975
      End
      Begin VB.Label corresponde 
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
         Height          =   915
         Left            =   8190
         TabIndex        =   44
         Top             =   7380
         Width           =   5190
      End
      Begin VB.Label Label6 
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
         Left            =   7740
         TabIndex        =   43
         Top             =   270
         Width           =   1020
      End
      Begin VB.Label lblnombredocumento 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "   GUIA DESPACHO FLETE"
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
         Left            =   3360
         TabIndex        =   35
         Top             =   600
         Width           =   7215
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Nº Guia"
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
         Left            =   0
         TabIndex        =   34
         Top             =   600
         Width           =   1695
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
         Left            =   0
         TabIndex        =   32
         Top             =   2040
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
         Left            =   6720
         TabIndex        =   30
         Top             =   1320
         Width           =   1695
      End
      Begin ShockwaveFlashObjectsCtl.ShockwaveFlash opciones 
         Height          =   1350
         Left            =   135
         TabIndex        =   11
         Top             =   7200
         Visible         =   0   'False
         Width           =   7455
         _cx             =   13150
         _cy             =   2381
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
         Left            =   0
         TabIndex        =   29
         Top             =   960
         Width           =   1695
      End
      Begin VB.Label lblDV 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
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
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   3240
         TabIndex        =   28
         Top             =   960
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
         Left            =   3765
         TabIndex        =   27
         Top             =   960
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
         Left            =   9120
         TabIndex        =   19
         Top             =   2040
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
         Left            =   9600
         TabIndex        =   18
         Top             =   2040
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
         Left            =   6720
         TabIndex        =   17
         Top             =   1680
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
         Left            =   0
         TabIndex        =   16
         Top             =   1680
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
         Left            =   10935
         TabIndex        =   15
         Top             =   600
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
         Left            =   0
         TabIndex        =   14
         Top             =   1320
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
         Left            =   6600
         TabIndex        =   13
         Top             =   2040
         Visible         =   0   'False
         Width           =   1125
      End
   End
End
Attribute VB_Name = "Pdespachoflete"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
  Private modifica As Boolean
  Private lectura As Boolean
  Private tipoprecio As String
  Private fecha As String
  Private formatogrilla(10, 10) As String
  Private numeroDoc As String
  Private numerofle As String
  Private horaorden As String
  Private final As Double
  Private DESPACHOTOTAL As String
  


Private Sub Command1_Click()
                    If Option2.Value = True Then
                         Grid1.Rows = 1
                        Call imprime_largo
                    End If

                    If Option1.Value = True Then
                         Grid1.Rows = 1
                        Call imprime_corto
                    End If
            FRMIMPRESION.Visible = False
            opciones.Visible = True
            detalle.Enabled = True
End Sub

Private Sub Command2_Click()
FRMIMPRESION.Visible = False
opciones.Visible = True
detalle.Enabled = True
End Sub

Private Sub dato2_KeyPress(KeyAscii As Integer)
     KeyAscii = esNumero(KeyAscii)
        If KeyAscii = 13 Then
            dato2.text = ceros(dato2)
            If leerguia(dato2.text, dato23.text) = False Then
            MsgBox ("guia de despacho no existe")
            dato2.SetFocus
            
           
            Else
            
            opciones.Visible = True
'            opciones.SetFocus
            End If
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
 Private Sub DATO7_GotFocus()
        Call VerificarCajas(Me, dato7)
        Call selecciona(dato7)
    End Sub
Private Sub DATO8_GotFocus()
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
    
    Private Sub dato4_KeyDown(KeyCode As Integer, Shift As Integer)
        Call Flechas(KeyCode, dato3)
    End Sub
    
    Private Sub dato5_KeyDown(KeyCode As Integer, Shift As Integer)
        Call Flechas(KeyCode, dato4)
    End Sub
    Private Sub dato6_KeyDown(KeyCode As Integer, Shift As Integer)
        Call Flechas(KeyCode, dato5)
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
                Call LeerClienteFlete(rut_cliente & lbldv.Caption, "0")
             rut_cliente = dato6.text
            'SendKeys "{Tab}"
        End If
    End Sub
    
    Private Sub DATO7_KeyPress(KeyAscii As Integer)
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If KeyAscii = 13 Then
        dato8.SetFocus
        End If
    End Sub
     Private Sub DATO8_KeyPress(KeyAscii As Integer)
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If KeyAscii = 13 Then
        dato9.SetFocus
        End If
    End Sub
    Private Sub DATO9_KeyPress(KeyAscii As Integer)
        KeyAscii = esNumero(KeyAscii)
        If KeyAscii = 13 Then
           dato10.SetFocus
        End If
    End Sub
     Private Sub DATO10_KeyPress(KeyAscii As Integer)
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
   
    Private Sub dato13_KeyPress(KeyAscii As Integer)
       KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If KeyAscii = 13 Then
                opciones.Visible = True
                opciones.SetFocus
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
     
    Private Sub dato10_LostFocus()
        Call limpiaBarra(2)
    End Sub
   
  
 Private Sub CARGAGRILLA(ByVal row As Integer, ByVal col As Integer)
        Dim i As Integer
        Rem DATOS DE LA COLUMNA
        formatogrilla(1, 0) = "LN"
        formatogrilla(1, 1) = "CODIGO"
        formatogrilla(1, 2) = "CANTIDAD"
        formatogrilla(1, 3) = "DESCRIPCION"
        formatogrilla(1, 4) = "PRECIO"
        formatogrilla(1, 5) = "TOTAL   "
      
  
        
        Rem LARGO DE LOS DATOS
        formatogrilla(2, 1) = "10"
        formatogrilla(2, 2) = "10"
        formatogrilla(2, 3) = "40"
        formatogrilla(2, 4) = "10"
        formatogrilla(2, 5) = "10"
        formatogrilla(2, 6) = "0"

       
      
        
        
        Rem TIPO DE DATOS C=CEROS,S=STRING,N=NUMERICO,D=DATE,
        formatogrilla(3, 1) = "N"
        formatogrilla(3, 2) = "N"
        formatogrilla(3, 3) = "S"
        formatogrilla(3, 4) = "N"
        formatogrilla(3, 5) = "N"

       
              
        Rem FORMATO GRILLA
        formatogrilla(4, 1) = "0000000000000"
        formatogrilla(4, 2) = ""
        formatogrilla(4, 3) = "0000000000000"
        formatogrilla(4, 4) = "###,###,##0"
        formatogrilla(4, 5) = "###,###,##0"
      
   
       
        Rem LOCCKED
        'FormatoGrilla(5, 0) = "TRUE"
        formatogrilla(5, 1) = "TRUE"
        formatogrilla(5, 2) = "TRUE"
        formatogrilla(5, 3) = "TRUE"
        formatogrilla(5, 4) = "TRUE"
        formatogrilla(5, 5) = "TRUE"
         formatogrilla(5, 6) = "TRUE"
     

      
       
        Rem VALOR MINIMO
        formatogrilla(6, 1) = ""
        formatogrilla(6, 2) = ""
        formatogrilla(6, 3) = ""
        formatogrilla(6, 4) = ""
        formatogrilla(6, 5) = ""
      
       
        
        Rem VALOR MAXIMO
        formatogrilla(7, 1) = ""
        formatogrilla(7, 2) = ""
        formatogrilla(7, 3) = ""
        formatogrilla(7, 4) = ""
        formatogrilla(7, 5) = ""
      
   
       
        Rem ANCHO
        formatogrilla(8, 1) = "10"
        formatogrilla(8, 2) = "10"
        formatogrilla(8, 3) = "46"
        formatogrilla(8, 4) = "10"
        formatogrilla(8, 5) = "10"
        formatogrilla(8, 6) = "0"
   
      
            
        detalle.Cols = 11
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
        detalle.Column(7).Width = 0
        detalle.Column(8).Width = 0
        detalle.Column(9).Width = 0
        detalle.Column(10).Width = 0
        
        
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
'        detalle.AddItem vbTab & vbTab & "1" & vbTab & "0" & vbTab & "0" & vbTab & "0" & vbTab & "0" & vbTab & "0" & vbTab & "", True
        detalle.Enabled = True
'        detalle.ExtendLastCol = True
        
    
    End Sub

 

Private Sub detalle_KeyPress(KeyAscii As Integer)
KeyAscii = esNumero(KeyAscii)
End Sub


Private Sub Form_Activate()
dato2.SetFocus
End Sub

Private Sub Form_Load()
Call CARGAGRILLA(1, 7)
dato3.text = Format(fechasistema, "dd")
dato4.text = Format(fechasistema, "mm")
dato5.text = Format(fechasistema, "yyyy")
End Sub



  Private Sub opciones_FSCommand(ByVal command As String, ByVal args As String)
        Select Case command
            Case "modifica"
'                Call modificar
            Case "elimina"
                 If Verifica_Permiso(Me.Caption, "elimina") = True Then
                    If MsgBox("DESEA REALMENTE ELIMINAR Si / No", vbYesNo) = vbYes Then
                    Call ELIMINAR
                    retorno
                    End If
                Else
                    MsgBox mensaje_nopermiso, vbCritical + vbOKOnly, "Permiso Denegado"
                End If
            Case "imprime"
               If dato2.text <> "" Then
                    dato6_KeyPress (13)
                    If detalle.Rows > 12 Then
                         Option2.Value = True
                         
                    Else
                        Option1.Value = True
                        
                    End If
                    opciones.Visible = False
                    detalle.Enabled = False
                    FRMIMPRESION.Visible = True
               End If
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
    detalle.Rows = 1
    dato23.text = ""
    lbllocal.Caption = ""
    lbldv.Caption = ""
    lblDVV.Caption = ""
    lblVendedor.Caption = ""
    opciones.Visible = False
    corresponde.Caption = ""
    Call CARGAGRILLA(1, 7)
    
    dato2.SetFocus
    End Sub

 Function LEERULTIMAGUIA() As String
    Dim csql As New rdoQuery
    Dim resultados As rdoResultset
    
    Set csql.ActiveConnection = ventasRubro
    
    csql.sql = " select IFNULL(MAX(numero) + 1,'0000000001')"
    csql.sql = csql.sql & " from sv_guia_despacho_flete_" & empresaActiva
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

Sub ELIMINAR()
Dim K As Integer
Call eliminarguia(dato2.text)

For K = 1 To detalle.Rows - 1
Call modificardespacho(detalle.Cell(K, 7).text, detalle.Cell(K, 8).text, detalle.Cell(K, 9).text, detalle.Cell(K, 2).text, detalle.Cell(K, 10).text)

Next K

End Sub
Private Sub eliminarguia(NUMERO)
  Dim csql As New rdoQuery
  
    Set csql.ActiveConnection = ventasRubro
    csql.sql = "delete from sv_guia_despacho_entrega_" & empresaActiva
    csql.sql = csql.sql & " where  numero='" & NUMERO & "'"
    csql.Execute
    Call eliminaDetalle(empresaActiva)
    
    Set csql.ActiveConnection = gestionRubro
    csql.sql = "delete from l_movimientos_detalle_" & empresaActiva
    csql.sql = csql.sql & " where  numero='" & NUMERO & "' and tipo='EL' "
    csql.Execute
    
    Set csql.ActiveConnection = gestionRubro
    csql.sql = "delete from l_movimientos_cabeza_" & empresaActiva
    csql.sql = csql.sql & " where  numero='" & NUMERO & "' and tipo='EL' "
    csql.Execute
    
    Call eliminaDetalle(dato23.text)
    
    Set csql.ActiveConnection = gestionRubro
    csql.sql = "delete from l_movimientos_detalle_" & dato23.text
    csql.sql = csql.sql & " where  numero='" & NUMERO & "' and tipo='RL' "
    csql.Execute
    
    Set csql.ActiveConnection = gestionRubro
    csql.sql = "delete from l_movimientos_cabeza_" & dato23.text
    csql.sql = csql.sql & " where  numero='" & NUMERO & "' and tipo='RL' "
    csql.Execute
    
    
    csql.Close
    End Sub

Private Function leerguia(NUMERO, LOCALdocumento) As Boolean
  Dim csql As New rdoQuery
  Dim resultados As rdoResultset
  Dim linea As Double
  Dim tipodo As String
  If LOCALdocumento = "" Then LOCALdocumento = empresaActiva
    Set csql.ActiveConnection = ventasRubro
    csql.sql = "select numero,linea,rut,fecha,codigo,descripcion,cantidad,precio,total,tipodocumento,numerodocumento,cajadocumento,lineadocumento,localdocumento "
    csql.sql = csql.sql & "from " & baseVentas & LOCALdocumento & ".sv_guia_despacho_entrega_" & LOCALdocumento
    csql.sql = csql.sql & " where  numero='" & NUMERO & "'"
    csql.Execute
    
    If csql.RowsAffected > 0 Then
    detalle.Rows = 1
    Set resultados = csql.OpenResultset
    leerguia = True
    dato23.text = resultados("localdocumento")
    lbllocal.Caption = leerNombreEmpresa(dato23.text)
    
    dato6.text = resultados("rut")
    dato6.Enabled = True
    dato6_KeyPress (13)
    dato3.text = Format(resultados("fecha"), "dd")
    dato4.text = Format(resultados("fecha"), "mm")
    dato5.text = Format(resultados("fecha"), "yyyy")
    If resultados("tipodocumento") = "FV" Then
    tipodo = "FACTURA"
    Else
    tipodo = "BOLETA"
    End If
     Call LeerClienteFlete(rut_cliente & lbldv.Caption, "0")
    
    DESPACHOTOTAL = tipodo + " " + leerfoliosii(resultados("tipodocumento"), resultados("numerodocumento"), resultados("cajadocumento"), resultados("localdocumento"))
    corresponde.Caption = DESPACHOTOTAL
    While Not resultados.EOF
    detalle.Rows = detalle.Rows + 1
    detalle.Cell(detalle.Rows - 1, 1).text = resultados("codigo")
    detalle.Cell(detalle.Rows - 1, 2).text = resultados("cantidad")
    detalle.Cell(detalle.Rows - 1, 3).text = resultados("descripcion")
    detalle.Cell(detalle.Rows - 1, 4).text = resultados("precio")
    detalle.Cell(detalle.Rows - 1, 5).text = resultados("total")
    detalle.Cell(detalle.Rows - 1, 7).text = resultados("tipodocumento")
    detalle.Cell(detalle.Rows - 1, 8).text = resultados("numerodocumento")
    detalle.Cell(detalle.Rows - 1, 9).text = resultados("cajadocumento")
    detalle.Cell(detalle.Rows - 1, 10).text = resultados("lineadocumento")
    
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
    csql.sql = "select fecha,tipodocumento,numerodocumento from sv_fletes_cabeza_" & empresaActiva
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
    Function leerfoliosii(TIPO, NUMERO, caja, EMPRESA) As String
    Dim csql As New rdoQuery
    Dim resultado As rdoResultset
        
    Set csql.ActiveConnection = ventasRubro
    
    csql.sql = "select foliosii from " + baseVentas + EMPRESA + ".sv_documento_cabeza_" & EMPRESA + " "
    csql.sql = csql.sql & "where tipo='" + TIPO + "' and numero='" + NUMERO + "' and caja='" + caja + "' "
    csql.Execute
    If csql.RowsAffected > 0 Then
    Set resultado = csql.OpenResultset
       leerfoliosii = resultado(0)
    End If
    csql.Close
    Set resultado = Nothing
    Set csql = Nothing
    
    End Function
    
    
    Public Sub imprime_largo()
    Dim SS As String
    Dim i As Integer
    Dim K As Integer
    Dim cad As String
    Dim totalprod As String
    Dim Descuento As String
    Dim neto As String
    Dim piva As String
    Dim piha As String
    Dim total As String
    Dim lineas As Integer
    Dim fecha As String
    Dim vencimiento As String
    Dim vendedor As String
    Dim notapedido As String
    Dim nombre As String
    Dim rut As String
    Dim direccion As String
    Dim ciudad As String
    Dim comuna As String
    Dim giro As String
    Dim fono As String
    Dim o As Integer
    Dim dia As String
    Dim mes As String
    Dim ano As String
    Dim nvalor As Long
    Dim CODIGO As String
    Dim tiposDePago As String
    Dim razon As String
    Dim dife As Double
    Dim impuesto(10) As Double
    Dim TAZAS(10) As Double
    Dim totalfinal As Double
    
    Grid1.Rows = 3
    Grid1.Cols = 6
    Grid1.Rows = 60
    Grid1.DefaultFont.Size = 10
    
    Grid1.DefaultFont.Bold = False
    
    Grid1.Column(0).Width = 0
    Grid1.Column(1).Width = 150
    Grid1.Column(2).Width = 90
    Grid1.Column(3).Width = 260
    Grid1.Column(4).Width = 100
    Grid1.Column(5).Width = 150
    
    Grid1.Column(1).Alignment = cellRightCenter
    Grid1.Column(2).Alignment = cellCenterCenter
    Grid1.Column(3).Alignment = cellLeftCenter '/**/
    Grid1.Column(4).Alignment = cellRightCenter
    Grid1.Column(5).Alignment = cellRightCenter

    'grid1.Column(7).Alignment = cellRightCenter
    
   
    
    Grid1.PageSetup.PrintGridlines = False
    Grid1.AutoRedraw = False
    
 
        
    'CABEZA
'    tabla = "SELECT nombre, rut, direccion, ciudad, comuna, giro,razonsocial "
'    tabla = tabla & "FROM g_maestroempresas "
'    tabla = tabla & "WHERE codigo = '" & dato6.text & "'"
'
'    Call ConectarControlData(rollo, servidor, basedatos, usuario, password, tabla)
    
'    If rollo.Recordset.RecordCount > 0 Then
'       rollo.Recordset.MoveFirst
       nombre = dato7.text
       rut = dato6.text & lbldv.Caption
       direccion = dato8.text
       ciudad = dato11.text
       comuna = dato10.text
        giro = dato13.text
        razon = dato7.text
'    End If
        
        
    
    Grid1.Cell(5, 5).text = dato2.text
        
    Grid1.Range(5, 2, 5, 3).Merge
    Grid1.Range(5, 2, 5, 3).Alignment = cellCenterCenter
    Grid1.Cell(5, 2).text = leerNombreEmpresa(empresaActiva)
    
    '
    
    'SEÑORES
    Grid1.Range(9, 2, 9, 3).Merge
    Grid1.Range(9, 2, 9, 3).Alignment = cellLeftCenter
    
    Grid1.Cell(9, 2).text = razon
    
    ' fecha
    fecha = dato3.text + "-" + dato4.text + "-" + dato5.text
    Grid1.Cell(9, 5).text = fecha
    
    
    'DIRECCION
    Grid1.Range(11, 2, 11, 3).Merge
    Grid1.Range(11, 2, 11, 3).Alignment = cellLeftCenter
    Grid1.Cell(11, 2).text = direccion
    
    'RUT
    'grid1.Range(9, 2, 9, 3).Merge
    Grid1.Cell(11, 5).Alignment = cellLeftCenter
    Grid1.Cell(11, 5).text = "     " + Format(Left(rut, 9), "###,###,###") & "-" & Right(rut, 1)
    
    'GIRO
    Grid1.Range(13, 2, 13, 3).Merge
    Grid1.Range(13, 2, 13, 3).Alignment = cellLeftCenter
    Grid1.Cell(13, 2).text = giro
    
    'CIUDAD
    
    Grid1.Cell(13, 5).Alignment = cellLeftCenter
    Grid1.Cell(13, 5).text = "     " + ciudad
        
        
    'grid1

        lineas = 19
     For K = 1 To detalle.Rows - 1
            lineas = lineas + 1
            Grid1.Cell(lineas, 1).text = detalle.Cell(K, 1).text
            Grid1.Cell(lineas, 2).text = Format(detalle.Cell(K, 2).text, "###,##0.00")
            Grid1.Cell(lineas, 3).text = detalle.Cell(K, 3).text
            Grid1.Cell(lineas, 4).text = Format(detalle.Cell(K, 4).text, "###,###,###")
            Grid1.Cell(lineas, 5).text = Format(detalle.Cell(K, 5).text, "###,###,###")
            totalfinal = totalfinal + CDbl(Grid1.Cell(lineas, 5).text)
           
            
      Next K
      
        Grid1.Range(47, 1, 47, 5).Merge
        Grid1.Range(47, 1, 47, 5).Alignment = cellCenterCenter
        Grid1.Range(47, 1, 47, 5).FontBold = True
        Grid1.Cell(47, 1).text = "NO CONSTITUYE VENTA, SOLO ENTREGA MERCADERIA."
        Grid1.Range(48, 1, 48, 5).Merge
        Grid1.Range(48, 1, 48, 5).Alignment = cellCenterCenter
        Grid1.Range(48, 1, 48, 5).FontBold = True
        
        Grid1.Cell(48, 1).text = "MERCADERIA CORRESPONDE " + DESPACHOTOTAL
    
    
    
    Grid1.Cell(51, 4).Alignment = cellLeftCenter
    Grid1.Cell(51, 4).text = "  NETO"
    Grid1.Cell(51, 5).text = Format(Round(CDbl(totalfinal / 1.19), 0), "###,###,##0")
    
    Grid1.Cell(52, 4).Alignment = cellLeftCenter
    Grid1.Cell(52, 4).text = "  IVA"
    Grid1.Cell(52, 5).text = Format(Round(totalfinal - totalfinal / 1.19, 0), "###,###,##0")
    
    Grid1.Cell(53, 4).Alignment = cellLeftCenter
    Grid1.Cell(53, 4).text = "  EXENTO"
    Grid1.Cell(53, 5).text = Format(0, "###,###,##0")
    
    Grid1.Cell(54, 4).Alignment = cellLeftCenter
    Grid1.Cell(54, 4).text = "  TOTAL"
    
    Grid1.Cell(54, 5).text = Format(totalfinal, "###,###,##0")
    
    nvalor = Format(totalfinal)
    SS = Numero_Texto(nvalor)
    Grid1.Range(51, 1, 53, 3).Merge
    Grid1.Range(51, 1, 53, 3).Alignment = cellLeftCenter
    Grid1.Cell(51, 1).text = "        " + SS
    Grid1.Range(51, 1, 53, 3).WrapText = True
    
    Grid1.AutoRedraw = True
    Grid1.Refresh
    
    Grid1.PageSetup.LeftMargin = 0.25
    Grid1.PageSetup.RightMargin = 0
    Grid1.PageSetup.TopMargin = 2.5
    Grid1.PageSetup.BottomMargin = 0
    
    For i = 1 To Grid1.PageSetup.PaperSizes.Count
        If UCase(Grid1.PageSetup.PaperSizes.Item(i).PaperName) = "CARTA" Then
            Grid1.PageSetup.PaperSize = Grid1.PageSetup.PaperSizes.Item(i).Kind
            Exit For
        End If
    Next i
    
    Grid1.PageSetup.PrintGridlines = False
    
    'grid1.DirectPrint
    Grid1.PrintPreview
End Sub
    
  Public Sub imprime_corto()
    Dim SS As String
    Dim i As Integer
    Dim K As Integer
    Dim cad As String
    Dim totalprod As String
    Dim Descuento As String
    Dim neto As String
    Dim piva As String
    Dim piha As String
    Dim total As String
    Dim lineas As Integer
    Dim fecha As String
    Dim vencimiento As String
    Dim vendedor As String
    Dim notapedido As String
    Dim nombre As String
    Dim rut As String
    Dim direccion As String
    Dim ciudad As String
    Dim comuna As String
    Dim giro As String
    Dim fono As String
    Dim o As Integer
    Dim dia As String
    Dim mes As String
    Dim ano As String
    Dim nvalor As Long
    Dim CODIGO As String
    Dim tiposDePago As String
    Dim razon As String
    Dim dife As Double
    Dim impuesto(10) As Double
    Dim TAZAS(10) As Double
    Dim totalfinal As Double
    
    Grid1.Rows = 3
    Grid1.Cols = 6
    Grid1.Rows = 60
    Grid1.DefaultFont.Size = 10
    
    Grid1.DefaultFont.Bold = False
    
    Grid1.Column(0).Width = 0
    Grid1.Column(1).Width = 98
    Grid1.Column(2).Width = 90
    Grid1.Column(3).Width = 260
    Grid1.Column(4).Width = 100
    Grid1.Column(5).Width = 150
    
    Grid1.Column(1).Alignment = cellRightCenter
    Grid1.Column(2).Alignment = cellCenterCenter
    Grid1.Column(3).Alignment = cellLeftCenter '/**/
    Grid1.Column(4).Alignment = cellRightCenter
    Grid1.Column(5).Alignment = cellRightCenter
    Grid1.DefaultRowHeight = 12
    
    
    
    'grid1.Column(7).Alignment = cellRightCenter
    
   
    
    Grid1.PageSetup.PrintGridlines = False
    Grid1.AutoRedraw = False
    
 
        
    'CABEZA
'    tabla = "SELECT nombre, rut, direccion, ciudad, comuna, giro,razonsocial "
'    tabla = tabla & "FROM g_maestroempresas "
'    tabla = tabla & "WHERE codigo = '" & dato6.text & "'"
'
'    Call ConectarControlData(rollo, servidor, basedatos, usuario, password, tabla)
    
'    If rollo.Recordset.RecordCount > 0 Then
'       rollo.Recordset.MoveFirst
       nombre = dato7.text
       rut = dato6.text & lbldv.Caption
       direccion = dato8.text
       ciudad = dato11.text
       comuna = dato10.text
        giro = dato13.text
        razon = dato7.text
'    End If
        
        
    
    Grid1.Cell(5, 5).text = dato2.text
        
    Grid1.Range(5, 2, 5, 3).Merge
    Grid1.Range(5, 2, 5, 3).Alignment = cellCenterCenter
    Grid1.Cell(5, 2).text = leerNombreEmpresa(empresaActiva)
    
    '
    
    'SEÑORES
    Grid1.Range(9, 2, 9, 3).Merge
    Grid1.Range(9, 2, 9, 3).Alignment = cellLeftCenter
    
    Grid1.Cell(9, 2).text = razon
    
    ' fecha
    fecha = dato3.text + "-" + dato4.text + "-" + dato5.text
    Grid1.Cell(9, 5).text = fecha
    
    
    'DIRECCION
    Grid1.Range(11, 2, 11, 3).Merge
    Grid1.Range(11, 2, 11, 3).Alignment = cellLeftCenter
    Grid1.Cell(11, 2).text = direccion
    
    'RUT
    'grid1.Range(9, 2, 9, 3).Merge
    Grid1.Cell(11, 5).Alignment = cellLeftCenter
    Grid1.Cell(11, 5).text = "     " + Format(Left(rut, 9), "###,###,###") & "-" & Right(rut, 1)
    
    'GIRO
    Grid1.Range(13, 2, 13, 3).Merge
    Grid1.Range(13, 2, 13, 3).Alignment = cellLeftCenter
    Grid1.Cell(13, 2).text = giro
    
    'CIUDAD
    
    Grid1.Cell(13, 5).Alignment = cellLeftCenter
    Grid1.Cell(13, 5).text = "     " + ciudad
        
        
    'grid1

        lineas = 19
     For K = 1 To detalle.Rows - 1
            lineas = lineas + 1
            Grid1.Cell(lineas, 1).text = detalle.Cell(K, 1).text
            Grid1.Cell(lineas, 2).text = Format(detalle.Cell(K, 2).text, "###,##0.00")
            Grid1.Cell(lineas, 3).text = detalle.Cell(K, 3).text
            Grid1.Cell(lineas, 4).text = Format(detalle.Cell(K, 4).text, "###,###,###")
            Grid1.Cell(lineas, 5).text = Format(detalle.Cell(K, 5).text, "###,###,###")
            totalfinal = totalfinal + CDbl(Grid1.Cell(lineas, 5).text)
           
            
      Next K
      
        Grid1.Range(36, 1, 36, 3).Merge
        Grid1.Range(36, 1, 36, 3).Alignment = cellCenterCenter
        Grid1.Range(36, 1, 36, 3).FontBold = True
        Grid1.Cell(36, 1).text = "NO CONSTITUYE VENTA, SOLO ENTREGA MERCADERIA."
        Grid1.Range(37, 1, 37, 3).Merge
        Grid1.Range(37, 1, 37, 3).Alignment = cellCenterCenter
        Grid1.Range(37, 1, 37, 3).FontBold = True
        
        Grid1.Cell(37, 1).text = "MERCADERIA CORRESPONDE " + DESPACHOTOTAL
    
    
    
    Grid1.Cell(36, 4).Alignment = cellLeftCenter
    Grid1.Cell(36, 4).text = "      "
    Grid1.Cell(36, 5).text = Format(Round(CDbl(totalfinal / 1.19), 0), "###,###,##0")
    
    Grid1.Cell(37, 4).Alignment = cellLeftCenter
    Grid1.Cell(37, 4).text = "     "
    Grid1.Cell(37, 5).text = Format(Round(totalfinal - totalfinal / 1.19, 0), "###,###,##0")
    
    Grid1.Cell(38, 4).Alignment = cellLeftCenter
    Grid1.Cell(38, 4).text = "       "
    
    Grid1.Cell(38, 5).text = Format(totalfinal, "###,###,##0")
    
'    nvalor = Format(totalfinal)
'    SS = Numero_Texto(nvalor)
'    Grid1.Range(51, 1, 53, 3).Merge
'    Grid1.Range(51, 1, 53, 3).Alignment = cellLeftCenter
'    Grid1.Cell(51, 1).text = "        " + SS
'    Grid1.Range(51, 1, 53, 3).WrapText = True
'
    Grid1.AutoRedraw = True
    Grid1.Refresh
    
    Grid1.PageSetup.LeftMargin = 0.25
    Grid1.PageSetup.RightMargin = 0
    Grid1.PageSetup.TopMargin = 2.2
    Grid1.PageSetup.BottomMargin = 0
    
    For i = 1 To Grid1.PageSetup.PaperSizes.Count
        If UCase(Grid1.PageSetup.PaperSizes.Item(i).PaperName) = "CARTA" Then
            Grid1.PageSetup.PaperSize = Grid1.PageSetup.PaperSizes.Item(i).Kind
            Exit For
        End If
    Next i
    
    Grid1.PageSetup.PrintGridlines = False
    
    
    
    'grid1.DirectPrint
    Grid1.PrintPreview
End Sub
    
Sub LeerClienteFlete(rut, sucursal)


Dim op As Integer
Dim CAMPOS(6, 6) As String

    CAMPOS(0, 0) = "nombre"
    CAMPOS(1, 0) = "direccion"
    CAMPOS(2, 0) = "comuna"
    CAMPOS(3, 0) = "ciudad"
    CAMPOS(4, 0) = "fono1"
    CAMPOS(5, 0) = "giro"
    
    CAMPOS(6, 0) = ""
    CAMPOS(0, 2) = "sv_maestroclientes"
    condicion = "rut='" & rut & "' and sucursal='" & sucursal & "'"
    op = 5
    Set sqlventas.conexion = ventas
    sqlventas.response = CAMPOS
    Call sqlventas.sqlventas(op, condicion)

    If sqlventas.Status = 0 Then
    
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
'    detalle.SetFocus
  
   
    End If
End Sub
Public Sub pdespachofletedeafuera()
dato2_KeyPress (13)


End Sub

Public Sub modificardespacho(TIPO, NUMERO, caja, cantidad, linea)
Dim csql As New rdoQuery
Dim tabla As String
Set csql.ActiveConnection = ventasRubro

tabla = "update " + baseVentas + dato23.text + ".sv_documento_detalle_" & dato23.text & " set despachado=despachado-'" & cantidad & "' "
tabla = tabla & "where tipo='" & TIPO & "' and numero='" & NUMERO & "' and linea='" & linea & "' and caja='" + caja + "' "
csql.sql = tabla
csql.Execute
    Call sincronizadatos(csql.sql, ventasRubro)
csql.Close
Set csql = Nothing
End Sub

Private Sub eliminaDetalle(loc)
    Dim i As Integer
    Dim cantidad As Double
    Dim CODIGO As String
    Dim fecha As Date
    
    For i = 1 To detalle.Rows - 1
        fecha = dato5.text & "-" & dato4.text & "-" & dato3.text
        CODIGO = detalle.Cell(i, 1).text
        cantidad = CDbl(detalle.Cell(i, 2).text)
        If loc = empresaActiva Then
        Call desactualiza_stock("-", CODIGO, "N", "N", "00", dato5.text, cantidad, "0", fecha, "0", loc)
        Else
        Call desactualiza_stock("+", CODIGO, "N", "N", "00", dato5.text, cantidad, "0", fecha, "0", loc)
        End If
    
    Next i
    
End Sub

Public Sub leerguiadeafuera()
    Call dato2_KeyPress(13)

End Sub
