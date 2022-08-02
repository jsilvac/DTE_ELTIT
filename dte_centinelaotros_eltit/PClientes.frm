VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash9f.ocx"
Object = "{4777436C-EB8C-4596-98A8-EBCF98683969}#1.0#0"; "Flexcell.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{ADBBAED6-B16F-40DC-80DF-B44910CBA76C}#1.0#0"; "Frame-Xp.ocx"
Begin VB.Form PClientes 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pagos Clientes"
   ClientHeight    =   8295
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12855
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8295
   ScaleWidth      =   12855
   Begin XPFrame.FrameXp frmDeposito 
      Height          =   1155
      Left            =   3420
      TabIndex        =   49
      Top             =   1920
      Visible         =   0   'False
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   2037
      BackColor       =   8454016
      Caption         =   "Fecha Deposito"
      CaptionEstilo3D =   1
      BackColor       =   8454016
      ColorBarraArriba=   12648384
      ColorBarraAbajo =   32768
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
      Begin VB.TextBox txtAño 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   405
         Left            =   1500
         MaxLength       =   4
         TabIndex        =   8
         Tag             =   "proveedor"
         Top             =   540
         Width           =   855
      End
      Begin VB.TextBox txtMes 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   405
         Left            =   900
         MaxLength       =   2
         TabIndex        =   7
         Tag             =   "proveedor"
         Top             =   540
         Width           =   495
      End
      Begin VB.TextBox txtDia 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   405
         Left            =   300
         MaxLength       =   2
         TabIndex        =   6
         Tag             =   "proveedor"
         Top             =   540
         Width           =   495
      End
   End
   Begin XPFrame.FrameXp frmTipo 
      Height          =   2415
      Left            =   1620
      TabIndex        =   44
      Top             =   5820
      Visible         =   0   'False
      Width           =   2715
      _ExtentX        =   4789
      _ExtentY        =   4260
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
      Begin VB.Label lbl30 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " 4 - Saldo Anterior"
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
         TabIndex        =   51
         Top             =   1560
         Width           =   2475
      End
      Begin VB.Label lbl25 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " 3 - Nota de Credito"
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
         TabIndex        =   48
         Top             =   1200
         Width           =   2475
      End
      Begin VB.Line Line1 
         X1              =   120
         X2              =   2610
         Y1              =   1920
         Y2              =   1920
      End
      Begin VB.Label lbl26 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " * - Grabar"
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
         TabIndex        =   47
         Top             =   2025
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
         Left            =   120
         TabIndex        =   46
         Top             =   480
         Width           =   2475
      End
      Begin VB.Label lbl24 
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
         Left            =   120
         TabIndex        =   45
         Top             =   840
         Width           =   2475
      End
   End
   Begin XPFrame.FrameXp frmGrabar 
      Height          =   375
      Left            =   8220
      TabIndex        =   15
      Top             =   7380
      Visible         =   0   'False
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   661
      BackColor       =   49344
      Caption         =   "G   R   A   B   A   R"
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
   Begin XPFrame.FrameXp FrameXp2 
      Height          =   2535
      Left            =   120
      TabIndex        =   17
      Top             =   120
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   4471
      BackColor       =   16744576
      Caption         =   "Datos del Pago"
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
         Left            =   1560
         MaxLength       =   10
         TabIndex        =   0
         Tag             =   "proveedor"
         Top             =   360
         Width           =   1695
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
         Left            =   1560
         MaxLength       =   9
         TabIndex        =   1
         Tag             =   "proveedor"
         Top             =   720
         Width           =   1215
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
         Left            =   1560
         MaxLength       =   1
         TabIndex        =   5
         Tag             =   "proveedor"
         Top             =   1800
         Width           =   375
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
         Left            =   1560
         MaxLength       =   9
         TabIndex        =   9
         Tag             =   "proveedor"
         Top             =   2160
         Width           =   1695
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
         Left            =   1560
         MaxLength       =   2
         TabIndex        =   2
         Tag             =   "proveedor"
         Top             =   1440
         Width           =   435
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
         Left            =   2040
         MaxLength       =   2
         TabIndex        =   3
         Tag             =   "proveedor"
         Top             =   1440
         Width           =   435
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
         Left            =   2520
         MaxLength       =   4
         TabIndex        =   4
         Tag             =   "proveedor"
         Top             =   1440
         Width           =   735
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
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   2760
         TabIndex        =   28
         Top             =   720
         Width           =   495
      End
      Begin VB.Label lbl1 
         Appearance      =   0  'Flat
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Número"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   120
         TabIndex        =   27
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label lbl3 
         Appearance      =   0  'Flat
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Forma de Pago"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   120
         TabIndex        =   26
         Top             =   1800
         Width           =   1335
      End
      Begin VB.Label lbl4 
         Appearance      =   0  'Flat
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Monto"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   120
         TabIndex        =   25
         Top             =   2160
         Width           =   1335
      End
      Begin VB.Label lbl5 
         Appearance      =   0  'Flat
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Rut"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   120
         TabIndex        =   24
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label lbl2 
         Appearance      =   0  'Flat
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Fecha"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   120
         TabIndex        =   23
         Top             =   1440
         Width           =   1335
      End
      Begin VB.Label lblNombre 
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
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   120
         TabIndex        =   22
         Top             =   1080
         Width           =   3135
      End
      Begin VB.Label lblForma 
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
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   1920
         TabIndex        =   21
         Top             =   1800
         Width           =   1335
      End
   End
   Begin XPFrame.FrameXp FrameXp1 
      Height          =   2535
      Left            =   3600
      TabIndex        =   18
      Top             =   120
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   4471
      BackColor       =   16744576
      Caption         =   "Cheques"
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
      Begin VB.TextBox pivote 
         Height          =   285
         Left            =   240
         MaxLength       =   3
         TabIndex        =   29
         Top             =   2160
         Visible         =   0   'False
         Width           =   975
      End
      Begin FlexCell.Grid Cheques 
         Height          =   1695
         Left            =   120
         TabIndex        =   12
         Top             =   360
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   2990
         DefaultFontSize =   8.25
         Rows            =   2
         DateFormat      =   2
      End
      Begin VB.Label lbl6 
         Appearance      =   0  'Flat
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Total"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   6360
         TabIndex        =   31
         Top             =   2160
         Width           =   975
      End
      Begin VB.Label lblTotalCheques 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "$ 0"
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
         Left            =   7440
         TabIndex        =   30
         Top             =   2160
         Width           =   1575
      End
   End
   Begin XPFrame.FrameXp FrameXp3 
      Height          =   3975
      Left            =   120
      TabIndex        =   19
      Top             =   2760
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   7011
      BackColor       =   16744576
      Caption         =   "Documentos por Pagar"
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
      Begin VB.Frame Frame1 
         BackColor       =   &H00FF8080&
         Height          =   675
         Left            =   60
         TabIndex        =   40
         Top             =   3240
         Width           =   6135
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
            Left            =   4620
            MaxLength       =   10
            TabIndex        =   11
            Tag             =   "proveedor"
            Top             =   240
            Width           =   1380
         End
         Begin VB.TextBox dato8 
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
            Left            =   960
            MaxLength       =   2
            TabIndex        =   10
            Tag             =   "proveedor"
            Top             =   240
            Width           =   435
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
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   1440
            TabIndex        =   43
            Top             =   240
            Width           =   2235
         End
         Begin VB.Label lbl28 
            Appearance      =   0  'Flat
            BackColor       =   &H00F5C9B1&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " Número"
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   3780
            TabIndex        =   42
            Top             =   240
            Width           =   795
         End
         Begin VB.Label lbl27 
            Appearance      =   0  'Flat
            BackColor       =   &H00F5C9B1&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " Tipo"
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   120
            TabIndex        =   41
            Top             =   240
            Width           =   795
         End
      End
      Begin MSAdodcLib.Adodc data 
         Height          =   330
         Left            =   180
         Top             =   2940
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
         LockType        =   -1
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
      Begin FlexCell.Grid Documentos 
         Height          =   2475
         Left            =   120
         TabIndex        =   13
         Top             =   360
         Width           =   6015
         _ExtentX        =   10610
         _ExtentY        =   4366
         DefaultFontSize =   8.25
         Rows            =   2
         SelectionMode   =   1
         DateFormat      =   2
      End
      Begin VB.Label lblTotalDocumentos 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "$ 0"
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
         Left            =   4560
         TabIndex        =   33
         Top             =   2940
         Width           =   1575
      End
      Begin VB.Label lbl7 
         Appearance      =   0  'Flat
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Total"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   3480
         TabIndex        =   32
         Top             =   2940
         Width           =   975
      End
   End
   Begin XPFrame.FrameXp FrameXp4 
      Height          =   3975
      Left            =   6480
      TabIndex        =   20
      Top             =   2760
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   7011
      BackColor       =   16744576
      Caption         =   "Documentos Pagados"
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
      Begin FlexCell.Grid Pagar 
         Height          =   3135
         Left            =   120
         TabIndex        =   14
         Top             =   360
         Width           =   6015
         _ExtentX        =   10610
         _ExtentY        =   5530
         DefaultFontSize =   8.25
         Rows            =   2
         DateFormat      =   2
      End
      Begin VB.Label lbl29 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Doble Click Eliminar Pago"
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
         TabIndex        =   50
         Top             =   3600
         Width           =   2895
      End
      Begin VB.Label lbl9 
         Appearance      =   0  'Flat
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Abono"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   3420
         TabIndex        =   37
         Top             =   3600
         Width           =   975
      End
      Begin VB.Label lbldiferencia 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "$ 0"
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
         Left            =   1200
         TabIndex        =   36
         Top             =   3300
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.Label lbltotalabonos 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "$ 0"
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
         Left            =   4500
         TabIndex        =   35
         Top             =   3600
         Width           =   1575
      End
      Begin VB.Label lbl8 
         Appearance      =   0  'Flat
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Diferencia"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   120
         TabIndex        =   34
         Top             =   3300
         Visible         =   0   'False
         Width           =   975
      End
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
      TabIndex        =   38
      Top             =   0
      Width           =   555
   End
   Begin FlexCell.Grid impresion 
      Height          =   675
      Left            =   5940
      TabIndex        =   39
      Top             =   6060
      Visible         =   0   'False
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   1191
      Cols            =   7
      DefaultFontSize =   8.25
      Rows            =   1
   End
   Begin MSAdodcLib.Adodc pendientes 
      Height          =   330
      Left            =   7740
      Top             =   5940
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
      LockType        =   -1
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
   Begin VB.CheckBox chk1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Mostrar Pendientes al Imprimir"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2160
      TabIndex        =   52
      Top             =   6840
      Value           =   1  'Checked
      Width           =   4095
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash opciones 
      Height          =   1335
      Left            =   240
      TabIndex        =   16
      Top             =   6960
      Width           =   6135
      _cx             =   10821
      _cy             =   2355
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
End
Attribute VB_Name = "PClientes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Private formatoGrilla(10, 10) As String
    Private pc As pagoCliente
    Private modifica As Boolean
    Private MONTO As Double
    Private vacio As Boolean
    Private fila As Long
    Private columna As Long
    Private totalpago As Double
    Private totalpagado As Double
    Private diferencia As Double
    Private fechadeposito As String
    Private escribir As Boolean
    'Private segurity As Boolean
    
Private Sub CDetalle_Click()
    Load DetalleDocumento
    DetalleDocumento.TIPO = Pagar.Cell(Pagar.ActiveCell.row, 1).text
    DetalleDocumento.NUMERO = Pagar.Cell(Pagar.ActiveCell.row, 2).text
    DetalleDocumento.fechaAudit = "2007-05-09"
    DetalleDocumento.Show
End Sub

    Private Sub Cheques_KeyDown(KeyCode As Integer, ByVal Shift As Integer)
        fila = Cheques.ActiveCell.row
        columna = Cheques.ActiveCell.col
        If KeyCode = vbKeyF2 And Cheques.ActiveCell.col = 1 Then
            Call ayudaBancos(pivote)
            Cheques.ActiveCell.text = pivote.text
            Cheques.Cell(Cheques.ActiveCell.row, 2).text = leerNombreBanco(pivote.text)
        End If
        Select Case KeyCode
            Case 13, 37, 38, 39, 40
                If Cheques.ActiveCell.text <> "" Then
                    vacio = False
                Else
                    vacio = True
                End If
                If KeyCode = 13 Then
                    Select Case Cheques.ActiveCell.col
                        Case 1
                            If Cheques.ActiveCell.text <> "" Then
                                pivote.MaxLength = 3
                                pivote.text = Cheques.ActiveCell.text
                                pivote.text = ceros(pivote)
                                Cheques.ActiveCell.text = pivote.text
                            End If
                        Case 3
                            If Cheques.ActiveCell.text <> "" Then
                                pivote.MaxLength = 7
                                pivote.text = Cheques.ActiveCell.text
                                pivote.text = ceros(pivote)
                                Cheques.ActiveCell.text = pivote.text
                            End If
                        Case 6
                            If Cheques.ActiveCell.text = "" Or Cheques.ActiveCell.text = "00" Then
                                Cheques.ActiveCell.text = Format(fechasistema, "dd")
                                vacio = False
                                'Call Cheques_LeaveCell(Cheques.ActiveCell.row, 6, Cheques.ActiveCell.row, 7, False)
                            End If
                        Case 7
                            If Cheques.ActiveCell.text = "" Or Cheques.ActiveCell.text = "00" Then
                                Cheques.ActiveCell.text = Format(fechasistema, "mm")
                                vacio = False
                                'Call Cheques_LeaveCell(Cheques.ActiveCell.row, 7, Cheques.ActiveCell.row, 8, False)
                            End If
                        Case 8
                            If Cheques.ActiveCell.text = "" Or Cheques.ActiveCell.text = "0000" Then
                                Cheques.ActiveCell.text = Format(fechasistema, "yyyy")
                                vacio = False
                                'Call Cheques_LeaveCell(Cheques.ActiveCell.row, 8, Cheques.ActiveCell.row, 1, False)
                            End If
                    End Select
                End If
            Case 106
                If Cheques.Rows > 2 Then
                    'dato7.Text=
                    dato8.SetFocus
                End If
        End Select
    End Sub

    Private Sub Cheques_LeaveCell(ByVal row As Long, ByVal col As Long, NewRow As Long, NewCol As Long, Cancel As Boolean)
        Dim i As Integer
        Dim linea As String
        If row > 0 Then
            If vacio = True Then
                If NewRow > row Then
                    NewRow = fila
                    NewCol = columna
                Else
                    If NewCol > col Then
                        NewRow = fila
                        NewCol = columna
                    End If
                End If
            Else
                If col = Cheques.Cols - 1 And NewCol = 1 Then
                    If row = Cheques.Rows - 1 Then
                        Cheques.AddItem "", True
                        NewRow = Cheques.Rows - 1
                    End If
                End If
                If col = 1 And NewCol = Cheques.Cols - 1 Then
                    For i = 1 To Cheques.Cols - 1
                        If Cheques.Cell(NewRow, i).text = "" Then
                            NewCol = i
                            Exit For
                        End If
                    Next i
                End If
                If col = 1 And NewCol <> col Then
                    pivote.text = Cheques.Cell(row, 1).text
                    pivote.text = ceros(pivote)
                    Cheques.Cell(row, 1).text = pivote.text
                    Cheques.Cell(row, 2).text = leerNombreBanco(Cheques.Cell(row, 1).text)
                End If
                If NewRow > row Then
                    For i = 1 To Cheques.Cols - 1
                        If Cheques.Cell(row, i).text = "" Then
                            NewRow = row
                            NewCol = i
                            Exit For
                        End If
                    Next i
                    For i = 1 To Cheques.Cols - 1
                        If Cheques.Cell(NewRow, i).text = "" Then
                            NewCol = i
                            Exit For
                        End If
                    Next i
                End If
                Call sumaGrilla
            End If
        End If
    End Sub

    Private Sub sumaGrilla()
        Dim i As Long
        Dim suma As Double
        suma = 0
        For i = 1 To Cheques.Rows - 1
            If Cheques.Cell(i, 5).text <> "" Then
                suma = suma + CDbl(Cheques.Cell(i, 5).text)
            End If
        Next i
        lblTotalCheques.Caption = Format(suma, "$ ###,###,##0")
        dato7.text = suma
    End Sub
    
'============================================================
'MANEJO DE LOS CONTOLES
'============================================================
    '========================================================
    'GotFocus
    '========================================================
    Private Sub dato1_GotFocus()
        Call VerificarCajas(Me, dato1)
        Call selecciona(dato1)
        Principal.barraEstado.Panels(2).text = "F2: Ayuda Numero de Pago"
    End Sub
    
    Private Sub dato2_GotFocus()
        Call VerificarCajas(Me, dato2)
        Call selecciona(dato2)
        Principal.barraEstado.Panels(2).text = "F2: Ayuda Cliente"
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
        If dato6.text = "3" Then
            frmDeposito.Visible = True
            If fechadeposito <> "" Then
                txtDia.text = Format(fechadeposito, "dd")
                txtMes.text = Format(fechadeposito, "mm")
                txtAño.text = Format(fechadeposito, "yyyy")
            End If
        End If
        Call VerificarCajas(Me, dato6)
        Call selecciona(dato6)
        Principal.barraEstado.Panels(2).text = "F2: Ayuda Tipo Pago"
    End Sub
    
    Private Sub dato7_GotFocus()
        Call VerificarCajas(Me, dato7)
        Call selecciona(dato7)
        If dato6.text = "2" Then
            dato7.Locked = True
        Else
            dato7.Locked = False
        End If
    End Sub
    
    Private Sub dato8_GotFocus()
        dato8.text = ""
        Call VerificarCajas(Me, dato8)
        frmTipo.Visible = True
        Call selecciona(dato8)
    End Sub
    
    Private Sub dato9_GotFocus()
        Call VerificarCajas(Me, dato9)
        Call selecciona(dato9)
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


    
    Private Sub Pagar_DblClick()
        Dim i As Integer
        Dim fila As Integer
        fila = Pagar.ActiveCell.row
        For i = 1 To Documentos.Rows - 1
            If Documentos.Cell(i, 1).text = Pagar.Cell(fila, 1).text Then
                If Documentos.Cell(i, 2).text = Pagar.Cell(fila, 2).text Then
                    Documentos.Range(i, 1, i, Documentos.Cols - 1).ForeColor = 0
                    Exit For
                End If
            End If
        Next i
        Pagar.RemoveItem (fila)
        Call sumarPagos
    End Sub

    Private Sub Pagar_KeyUp(KeyCode As Integer, Shift As Integer)
        If escribir = True Then
            If Pagar.ActiveCell.col = 4 Then
                If KeyCode = 109 Then
                    If Left(Pagar.ActiveCell.text, 1) = "-" And Len(Pagar.ActiveCell.text) > 1 Then
                        Pagar.ActiveCell.text = Left(Pagar.ActiveCell.text, Len(Pagar.ActiveCell.text) - 1)
                    End If
                Else
                    If KeyCode > 57 Then
                        KeyCode = KeyCode - 48
                    End If
                    If KeyCode <> 37 And KeyCode <> 38 And KeyCode <> 39 And KeyCode <> 40 Then
                        If IsNumeric(Chr(KeyCode)) = False And Len(Pagar.ActiveCell.text) > 0 Then
                            Pagar.ActiveCell.text = Left(Pagar.ActiveCell.text, Len(Pagar.ActiveCell.text) - 1)
                        End If
                    End If
                End If
            End If
        End If
        escribir = True
    End Sub

    Private Sub Pagar_LeaveCell(ByVal row As Long, ByVal col As Long, NewRow As Long, NewCol As Long, Cancel As Boolean)
        If col = 4 Then
            If row <> NewRow Then
                If Pagar.Cell(row, col).text = "" Then
                    NewCol = col
                    NewRow = row
                End If
            End If
            If col <> NewCol Then
                If Pagar.Cell(row, col).text = "" Then
                    NewCol = col
                    NewRow = row
                End If
            End If
        End If
    End Sub

    Private Sub txtDia_GotFocus()
        'Call VerificarCajas(Me, txtDia)
        Call selecciona(txtDia)
    End Sub
    


    Private Sub txtMes_GotFocus()
        'Call VerificarCajas(Me, txtMes)
        Call selecciona(txtMes)
    End Sub
    
    Private Sub txtAño_GotFocus()
        'Call VerificarCajas(Me, txtAño)
        Call selecciona(txtAño)
    End Sub
    '========================================================
    'GotFocus
    '========================================================
    



    '========================================================
    'KeyDown
    '========================================================
    Private Sub dato1_KeyDown(KeyCode As Integer, Shift As Integer)
        If KeyCode = vbKeyF2 Then
            'Call ayudaFolioPago(dato1)
        Else
            Call Flechas(KeyCode, dato1)
        End If
    End Sub
    
    Private Sub dato2_KeyDown(KeyCode As Integer, Shift As Integer)
        If KeyCode = vbKeyF2 Then
            Call ayudaClienteDeuda(dato2, lblDV)
        Else
            Call Flechas(KeyCode, dato1)
        End If
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
            Call ayudaTipoPagopago(dato6)
        Else
            Call Flechas(KeyCode, dato5)
        End If
    End Sub
    
    Private Sub dato7_KeyDown(KeyCode As Integer, Shift As Integer)
        Call Flechas(KeyCode, dato6)
    End Sub
    
    Private Sub dato8_KeyDown(KeyCode As Integer, Shift As Integer)
        Dim diferencia As Double
        Select Case KeyCode
            Case 97
                dato8.text = "FV"
            Case 98
                dato8.text = "BV"
            Case 99
                dato8.text = "NV"
            Case 100
                dato8.text = "SA"
            Case 101
            '    dato1.text = "FE"
            '    tipoprecio = "01"
            'Case 102
            '    dato1.text = "GD"
            '    tipoprecio = "01"
            'Case 103
            '    dato1.text = "GM"
            '    tipoprecio = "01"
            Case 106, 35
                If frmGrabar.Visible = True Then
                    diferencia = CDbl(lbltotalabonos.Caption) - CDbl(dato7.text)
                    If diferencia = 0 Then
                        Call ctrltostruct
                    Else
                        If diferencia < 0 Then
                            If MsgBox("¿Desea dejar el vuelto como saldo a favor?", vbYesNo, "Pregunta") = vbYes Then
                                Call grabarSaldo(diferencia)
                                Call ctrltostruct
                            Else
                                dato7.text = Format(lbltotalabonos.Caption, "########0")
                                Call ctrltostruct
                                Call MsgBox("Vuelto: " & Format(diferencia * -1, "$ ###,###,##0"), vbOKOnly, "Información")
                            End If
                        Else
                            Call MsgBox("la cantidad abonada es menor que la cantidad pagada" & vbCrLf & "Por favor ajuste el abono o el pago", vbOKOnly, "Error")
                            Pagar.Cell(1, 4).SetFocus
                        End If
                    End If
'                    If CDbl(dato7.text) = CDbl(lbltotalabonos.Caption) Then
'                        Call ctrltostruct
'                        If MsgBox("DESEA IMPRIMIR COMPROBANTE ", vbYesNo) = vbYes Then
'                            Call imprimir
'                        End If
'                        Call retorno
'                    Else
'                        Call MsgBox("las cantidades abonadas y pagadas son distintas" & vbCrLf & "Por favor ajuste los abonos", vbOKOnly, "Error")
'                        Pagar.Cell(1, 4).SetFocus
'                    End If
                Else
                    Call MsgBox("Imposible Grabar." & vbCrLf & "Primero sleccione los documentos a cancelar", vbOKOnly, "Error")
                    dato8.SetFocus
                    dato8.text = ""
                End If
            Case Else
                Call Flechas(KeyCode, dato7)
        End Select
    End Sub
    
    Private Sub dato9_KeyDown(KeyCode As Integer, Shift As Integer)
        Call Flechas(KeyCode, dato8)
    End Sub
    
    Private Sub txtDia_KeyDown(KeyCode As Integer, Shift As Integer)
        Call Flechas(KeyCode, dato6)
    End Sub
    
    Private Sub txtMes_KeyDown(KeyCode As Integer, Shift As Integer)
        Call Flechas(KeyCode, txtDia)
    End Sub
    
    Private Sub txtAño_KeyDown(KeyCode As Integer, Shift As Integer)
        Call Flechas(KeyCode, txtMes)
    End Sub
    '========================================================
    'KeyDown
    '========================================================
    
    '========================================================
    'KeyPress
    '========================================================
    Private Sub dato1_KeyPress(KeyAscii As Integer)
        KeyAscii = esNumero(KeyAscii)
        If KeyAscii = 13 And dato1.text <> "" Then
            dato1.text = ceros(dato1)
            If Val(dato1.text) = 0 Then
                dato1.text = leer_Ultimo_Folio("numero", "sv_pagos_cabeza_" & empresaActiva, dato1.MaxLength, ventasRubro, "local = '" & empresaActiva & "'")
                SendKeys "{Tab}"
            Else
                If leerPagoCliente(pc, dato1.text, "=", data, Cheques, Pagar, lblTotalCheques, lbltotalabonos, lbldiferencia) = True Then
                    Call structtoctrl
                    frmGrabar.Visible = False
                    If dato6.text = "3" Then
                        frmDeposito.Visible = True
                        txtDia.text = Format(fechadeposito, "dd")
                        txtMes.text = Format(fechadeposito, "mm")
                        txtAño.text = Format(fechadeposito, "yyyy")
                    Else
                        frmDeposito.Visible = False
                    End If
                    If Cheques.Rows > 1 Then
                        Cheques.Range(1, 1, Cheques.Rows - 1, Cheques.Cols - 1).Locked = True
                    End If
                    Pagar.Range(1, 1, Pagar.Rows - 1, Pagar.Cols - 1).Locked = True
                Else
                    Call HabilitarCajas(Me, modifica)
                    If Documentos.Rows > 1 Then
                        Documentos.Range(1, 0, Documentos.Rows - 1, Documentos.Cols - 1).ClearText
                        Documentos.Rows = 1
                    End If
                    lblTotalCheques.Caption = "$ 0"
                    If Pagar.Rows > 1 Then
                        Pagar.Range(1, 0, Pagar.Rows - 1, Pagar.Cols - 1).ClearText
                        Pagar.Rows = 1
                    End If
                    lblTotalDocumentos.Caption = "$ 0"
                    If Cheques.Rows > 1 Then
                        Cheques.Range(1, 0, Cheques.Rows - 1, Cheques.Cols - 1).ClearText
                        Cheques.Rows = 1
                    End If
                    lbldiferencia.Caption = "$ 0"
                    lbldiferencia.Caption = "$ 0"
                    SendKeys "{Tab}"
                End If
            End If
        End If
    End Sub
    
    Private Sub dato2_KeyPress(KeyAscii As Integer)
        KeyAscii = esNumero(KeyAscii)
        If KeyAscii = 13 Then
            dato2.text = ceros(dato2)
            lblDV.Caption = rut(dato2.text)
            lblNombre.Caption = leerNombreCliente(dato2.text & lblDV.Caption)
            If lblNombre.Caption <> "" Then
                Call leerDocumentos(data, dato2.text & lblDV.Caption, lblTotalDocumentos, Documentos)
                SendKeys "{Tab}"
            End If
        End If
    End Sub
    
    Private Sub dato3_KeyPress(KeyAscii As Integer)
        KeyAscii = esNumero(KeyAscii)
        If KeyAscii = 13 Then
            dato3.text = ceros(dato3)
            If dato3.text = "00" Then
                dato3.text = Format(fechasistema, "dd")
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
            SendKeys "{Tab}"
        End If
    End Sub
    
    Private Sub dato6_KeyPress(KeyAscii As Integer)
        KeyAscii = esNumero(KeyAscii)
        If KeyAscii = 13 Then
            dato6.text = ceros(dato6)
            lblForma.Caption = leerFormaPago2(dato6.text)
            If lblForma.Caption <> "" Then
                If lblForma.Caption = "CHEQUE" Then
                    If Cheques.Rows = 1 Then
                        Cheques.AddItem "", True
                    End If
                    Cheques.Cell(1, 1).SetFocus
                Else
                    If lblForma.Caption = "DEPOSITO" Then
                        frmDeposito.Visible = True
                    End If
                    SendKeys "{Tab}"
                End If
            End If
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
        If KeyAscii = 13 Then
            If dato8.text = "1" Then dato8.text = "FV"
            If dato8.text = "2" Then dato8.text = "BV"
            If dato8.text = "3" Then dato8.text = "ZE"
            
            lblDocumento.Caption = leerNombreDocumento(dato8.text)
            If lblDocumento.Caption <> "" Then
                SendKeys "{Tab}"
            End If
        End If
    End Sub
    
    Private Sub dato9_KeyPress(KeyAscii As Integer)
        KeyAscii = esNumero(KeyAscii)
        If KeyAscii = 13 Then
            dato9.text = ceros(dato9)
            If verificarDocumento = True Then
                dato9.text = ""
                Pagar.Cell(Pagar.Rows - 1, 4).SetFocus
                'dato8.SetFocus
            Else
                Call selecciona(dato9)
            End If
        End If
    End Sub
    
    Private Sub txtDia_KeyPress(KeyAscii As Integer)
        KeyAscii = esNumero(KeyAscii)
        If KeyAscii = 13 Then
            txtDia.text = ceros(txtDia)
            If txtDia.text = "00" Then
                txtDia.text = Format(fechasistema, "dd")
            End If
            SendKeys "{Tab}"
        End If
    End Sub
    
    Private Sub txtMes_KeyPress(KeyAscii As Integer)
        KeyAscii = esNumero(KeyAscii)
        If KeyAscii = 13 Then
            txtMes.text = ceros(txtMes)
            If txtMes.text = "00" Then
                txtMes.text = Format(fechasistema, "mm")
            End If
            SendKeys "{Tab}"
        End If
    End Sub
    
    Private Sub txtAño_KeyPress(KeyAscii As Integer)
        KeyAscii = esNumero(KeyAscii)
        If KeyAscii = 13 Then
            txtAño.text = ceros(txtAño)
            If txtAño.text = "0000" Then
                txtAño.text = Format(fechasistema, "yyyy")
            End If
            SendKeys "{Tab}"
        End If
    End Sub
    '========================================================
    'KeyPress
    '========================================================
    
    '========================================================
    'KeyUp
    '========================================================
    Private Sub dato1_KeyUp(KeyCode As Integer, Shift As Integer)
        Call seleccionaUno(KeyCode, dato1)
    End Sub
    
    '========================================================
    'KeyUp
    '========================================================
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
'
'    Private Sub txtDia_KeyUp(KeyCode As Integer, Shift As Integer)
'        If Len(txtDia.text) = txtDia.MaxLength Then
'            Call txtDia_KeyPress(13)
'        End If
'    End Sub
'
'    Private Sub txtMes_KeyUp(KeyCode As Integer, Shift As Integer)
'        If Len(txtMes.text) = txtMes.MaxLength Then
'            Call txtMes_KeyPress(13)
'        End If
'    End Sub
'
'    Private Sub txtAño_KeyUp(KeyCode As Integer, Shift As Integer)
'        If Len(txtAño.text) = txtAño.MaxLength Then
'            Call txtAño_KeyPress(13)
'        End If
'    End Sub
    
    Private Sub dato9_KeyUp(KeyCode As Integer, Shift As Integer)
        Call seleccionaUno(KeyCode, dato9)
    End Sub
    '========================================================
    'KeyUp
    '========================================================
    
    '========================================================
    'LostFocus
    '========================================================
    Private Sub dato1_LostFocus()
        Call limpiaBarra(2)
    End Sub
    
    Private Sub dato2_LostFocus()
        Call limpiaBarra(2)
    End Sub
    
    Private Sub dato6_LostFocus()
        Call limpiaBarra(2)
        If dato6.text <> "3" Then
            frmDeposito.Visible = False
        End If
    End Sub
    
    Private Sub dato8_LostFocus()
        frmTipo.Visible = False
    End Sub
    Private Sub txtdia_LostFocus()
Call esfecha(txtDia, txtMes, txtAño, "dd")
End Sub
Private Sub txtmes_LostFocus()
Call esfecha(txtDia, txtMes, txtAño, "mm")
End Sub

    Private Sub txtaño_LostFocus()
        Call esfecha(txtDia, txtMes, txtAño, "yyyy")
        fechadeposito = txtAño.text & "-" & txtMes.text & "-" & txtDia.text
        frmDeposito.Visible = False
    End Sub
    '========================================================
    'LostFocus
    '========================================================
'============================================================
'MANEJO DE LOS CONTOLES
'============================================================

    Private Sub Documentos_DblClick()
        Dim cadena As String
        Dim TIPO As String
        Dim i As Integer
        Dim montoNC As Integer
        
        If Documentos.Cell(Documentos.ActiveCell.row, 1).ForeColor <> vbRed Then
            If dato7.text <> "" Then
                TIPO = Documentos.Cell(Documentos.ActiveCell.row, 1).text
                
                cadena = Documentos.Cell(Documentos.ActiveCell.row, 1).text & vbTab
                cadena = cadena & Documentos.Cell(Documentos.ActiveCell.row, 2).text & vbTab
                cadena = cadena & Documentos.Cell(Documentos.ActiveCell.row, 3).text & vbTab
                
                MONTO = CDbl(dato7.text)
                cadena = cadena & Documentos.Cell(Documentos.ActiveCell.row, 4).text
                Documentos.Range(Documentos.ActiveCell.row, 1, Documentos.ActiveCell.row, Documentos.Cols - 1).ForeColor = vbRed
                Pagar.AddItem cadena, True
                Pagar.Cell(Pagar.Rows - 1, 0).text = Documentos.Cell(Documentos.ActiveCell.row, 0).text
                sumarPagos
                'totalabonado
                
                frmGrabar.Visible = True
            Else
                MsgBox "No puede cancelar documentos si no ha ingresado un monto", vbOKOnly, "Error"
            End If
        End If
    End Sub
    
    Sub sumarPagos()
        Dim k As Integer
        totalpagado = 0
        For k = 1 To Pagar.Rows - 1
            totalpagado = totalpagado + Pagar.Cell(k, 4).text
        Next k
        lbltotalabonos.Caption = Format(totalpagado, "###,###,###")
    End Sub

    Private Sub Form_Activate()
    sqlventas.audit = True: sqlventas.programaactivo = Me.Caption
    sqlventas.programaactivo = Me.Caption
        
        If segurity = True Then
            Seguridad.Show vbModal
            segurity = False
        End If
        Principal.barraEstado.Panels(1).text = UCase(Me.Caption)
    End Sub
    
    Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
        If KeyCode = 27 Then
            Unload Me
        End If
        If KeyCode = 38 Then
            If Screen.ActiveForm.ActiveControl.Name = "dato1" Then
                Unload Me
            End If
        End If
    End Sub
    
    Private Sub Form_Load()
        'segurity = Not Verificar(usuarioSistema, passwordSistema)
        modifica = False
        escribir = False
        Call Centrar(Me)
        Call CargaGrillaDocumentos(1, 5)
        Call CargaGrillaPagar(1, 5)
        Call CargaGrillaCheques(1, 9)
        dato1.text = leer_Ultimo_Folio("numero", "sv_pagos_cabeza_" & empresaActiva, dato1.MaxLength, ventasRubro, "local = '" & empresaActiva & "'")
    End Sub
    
    Private Sub Form_Unload(Cancel As Integer)
        Principal.barraEstado.Panels(1).text = UCase(Principal.Caption)
        Call limpiaBarra(2)
    End Sub

'****************************************************************************
'Formato de la Grilla Documentos
'****************************************************************************
    Private Sub CargaGrillaDocumentos(ByVal row As Integer, ByVal col As Integer)
        Dim i As Integer
        Rem DATOS DE LA COLUMNA
        formatoGrilla(1, 1) = "TIPO"
        formatoGrilla(1, 2) = "NUMERO"
        formatoGrilla(1, 3) = "MONTO"
        formatoGrilla(1, 4) = "SALDO"
        
        Rem LARGO DE LOS DATOS
        formatoGrilla(2, 1) = "2"
        formatoGrilla(2, 2) = "12"
        formatoGrilla(2, 3) = "9"
        formatoGrilla(2, 4) = "9"
        
        Rem TIPO DE DATOS C=CEROS,S=STRING,N=NUMERICO,D=DATE,
        formatoGrilla(3, 1) = "S"
        formatoGrilla(3, 2) = "N"
        formatoGrilla(3, 3) = "N"
        formatoGrilla(3, 4) = "N"
        
        Rem FORMATO GRILLA
        formatoGrilla(4, 1) = "00"
        formatoGrilla(4, 2) = "0000000000"
        formatoGrilla(4, 3) = "$ ###,###,##0"
        formatoGrilla(4, 4) = "$ ###,###,##0"
        
        Rem LOCCKED
        formatoGrilla(5, 1) = "TRUE"
        formatoGrilla(5, 2) = "TRUE"
        formatoGrilla(5, 3) = "TRUE"
        formatoGrilla(5, 4) = "TRUE"
        
        Rem VALOR MINIMO
        formatoGrilla(6, 1) = ""
        formatoGrilla(6, 2) = ""
        formatoGrilla(6, 3) = ""
        formatoGrilla(6, 4) = ""
        
        Rem VALOR MAXIMO
        formatoGrilla(7, 1) = ""
        formatoGrilla(7, 2) = ""
        formatoGrilla(7, 3) = ""
        formatoGrilla(7, 4) = ""
        
        Rem ANCHO
        formatoGrilla(8, 1) = "5"
        formatoGrilla(8, 2) = "10"
        formatoGrilla(8, 3) = "12"
        formatoGrilla(8, 4) = "12"
            
        Documentos.Cols = col
        Documentos.Rows = row
        Documentos.AllowUserResizing = False
        Documentos.DisplayFocusRect = False
        Documentos.ExtendLastCol = True
        Documentos.BoldFixedCell = False
        Documentos.DrawMode = cellOwnerDraw
        Documentos.Appearance = Flat
        Documentos.ScrollBarStyle = Flat
        Documentos.FixedRowColStyle = Flat
        Documentos.BackColorFixed = RGB(90, 158, 214)
        Documentos.BackColorFixedSel = RGB(110, 180, 230)
        Documentos.BackColorBkg = RGB(90, 158, 214)
        Documentos.BackColorScrollBar = RGB(231, 235, 247)
        Documentos.BackColor1 = RGB(231, 235, 247)
        Documentos.BackColor2 = RGB(239, 243, 255)
        Documentos.GridColor = RGB(148, 190, 231)
        
        Documentos.Column(0).Width = 0
        For i = 1 To col - 1
            Documentos.Cell(0, i).text = formatoGrilla(1, i)
            Documentos.Column(i).Width = Val(formatoGrilla(8, i)) * (Documentos.Cell(0, i).Font.Size + 1.25)
            Documentos.Column(i).MaxLength = Val(formatoGrilla(2, i))
            Documentos.Column(i).FormatString = formatoGrilla(4, i)
            Documentos.Column(i).Locked = formatoGrilla(5, i)
            If formatoGrilla(3, i) = "N" Then
                Documentos.Column(i).Alignment = cellRightCenter
                Documentos.Column(i).Mask = cellNumeric
            Else
                Documentos.Column(i).Alignment = cellLeftCenter
                Documentos.Column(i).Mask = cellUpper
            End If
        Next i
        Documentos.Range(0, 0, 0, Documentos.Cols - 1).Alignment = cellCenterCenter
        Documentos.Enabled = True
    End Sub
'****************************************************************************
'Formato de la Grilla Documentos
'****************************************************************************

'****************************************************************************
'Formato de la Grilla Pagar
'****************************************************************************
    Private Sub CargaGrillaPagar(ByVal row As Integer, ByVal col As Integer)
        Dim i As Integer
        Rem DATOS DE LA COLUMNA
        formatoGrilla(1, 1) = "TIPO"
        formatoGrilla(1, 2) = "NUMERO"
        formatoGrilla(1, 3) = "MONTO"
        formatoGrilla(1, 4) = "ABONO"
        
        Rem LARGO DE LOS DATOS
        formatoGrilla(2, 1) = "2"
        formatoGrilla(2, 2) = "12"
        formatoGrilla(2, 3) = "9"
        formatoGrilla(2, 4) = "9"
        
        Rem TIPO DE DATOS C=CEROS,S=STRING,N=NUMERICO,D=DATE,
        formatoGrilla(3, 1) = "S"
        formatoGrilla(3, 2) = "N"
        formatoGrilla(3, 3) = "N"
        formatoGrilla(3, 4) = "N"
        
        Rem FORMATO GRILLA
        formatoGrilla(4, 1) = "00"
        formatoGrilla(4, 2) = "0000000000"
        formatoGrilla(4, 3) = "$ ###,###,##0"
        formatoGrilla(4, 4) = "$ ###,###,##0"
        
        Rem LOCCKED
        formatoGrilla(5, 1) = "TRUE"
        formatoGrilla(5, 2) = "TRUE"
        formatoGrilla(5, 3) = "TRUE"
        formatoGrilla(5, 4) = "FALSE"
        
        Rem VALOR MINIMO
        formatoGrilla(6, 1) = ""
        formatoGrilla(6, 2) = ""
        formatoGrilla(6, 3) = ""
        formatoGrilla(6, 4) = ""
        
        Rem VALOR MAXIMO
        formatoGrilla(7, 1) = ""
        formatoGrilla(7, 2) = ""
        formatoGrilla(7, 3) = ""
        formatoGrilla(7, 4) = ""
        
        Rem ANCHO
        formatoGrilla(8, 1) = "5"
        formatoGrilla(8, 2) = "11"
        formatoGrilla(8, 3) = "11"
        formatoGrilla(8, 4) = "11"
            
        Pagar.Cols = col
        Pagar.Rows = row
        Pagar.AllowUserResizing = False
        Pagar.DisplayFocusRect = False
        Pagar.ExtendLastCol = True
        Pagar.BoldFixedCell = False
        Pagar.DrawMode = cellOwnerDraw
        Pagar.Appearance = Flat
        Pagar.ScrollBarStyle = Flat
        Pagar.FixedRowColStyle = Flat
        Pagar.BackColorFixed = RGB(90, 158, 214)
        Pagar.BackColorFixedSel = RGB(110, 180, 230)
        Pagar.BackColorBkg = RGB(90, 158, 214)
        Pagar.BackColorScrollBar = RGB(231, 235, 247)
        Pagar.BackColor1 = RGB(231, 235, 247)
        Pagar.BackColor2 = RGB(239, 243, 255)
        Pagar.GridColor = RGB(148, 190, 231)
        
        Pagar.Column(0).Width = 0
        For i = 1 To col - 1
            Pagar.Cell(0, i).text = formatoGrilla(1, i)
            Pagar.Column(i).Width = Val(formatoGrilla(8, i)) * (Pagar.Cell(0, i).Font.Size + 1.25)
            Pagar.Column(i).MaxLength = Val(formatoGrilla(2, i))
            Pagar.Column(i).FormatString = formatoGrilla(4, i)
            Pagar.Column(i).Locked = formatoGrilla(5, i)
            If formatoGrilla(3, i) = "N" Then
                Pagar.Column(i).Alignment = cellRightCenter
                If i <> 4 Then
                    Pagar.Column(i).Mask = cellNumeric
                End If
            Else
                Pagar.Column(i).Alignment = cellLeftCenter
                Pagar.Column(i).Mask = cellUpper
            End If
        Next i
        Pagar.Range(0, 0, 0, Pagar.Cols - 1).Alignment = cellCenterCenter
        Pagar.Enabled = True
    End Sub
'****************************************************************************
'Formato de la Grilla Pagar
'****************************************************************************

'****************************************************************************
'Formato de la Grilla Cheques
'****************************************************************************
    Private Sub CargaGrillaCheques(ByVal row As Integer, ByVal col As Integer)
        Dim i As Integer
        Rem DATOS DE LA COLUMNA
        formatoGrilla(1, 1) = "BANCO"
        formatoGrilla(1, 2) = "NOMBRE"
        formatoGrilla(1, 3) = "NUMERO"
        formatoGrilla(1, 4) = "CUENTA"
        formatoGrilla(1, 5) = "MONTO"
        formatoGrilla(1, 6) = "VENCIMIENTO"
        formatoGrilla(1, 7) = ""
        formatoGrilla(1, 8) = ""
        
        Rem LARGO DE LOS DATOS
        formatoGrilla(2, 1) = "3"
        formatoGrilla(2, 2) = "30"
        formatoGrilla(2, 3) = "7"
        formatoGrilla(2, 4) = "11"
        formatoGrilla(2, 5) = "9"
        formatoGrilla(2, 6) = "2"
        formatoGrilla(2, 7) = "2"
        formatoGrilla(2, 8) = "4"
        
        Rem TIPO DE DATOS C=CEROS,S=STRING,N=NUMERICO,D=DATE,
        formatoGrilla(3, 1) = "N"
        formatoGrilla(3, 2) = "S"
        formatoGrilla(3, 3) = "N"
        formatoGrilla(3, 4) = "N"
        formatoGrilla(3, 5) = "N"
        formatoGrilla(3, 6) = "N"
        formatoGrilla(3, 7) = "N"
        formatoGrilla(3, 8) = "N"
        
        Rem FORMATO GRILLA
        formatoGrilla(4, 1) = "000"
        formatoGrilla(4, 2) = ""
        formatoGrilla(4, 3) = "0000000"
        formatoGrilla(4, 4) = "00000000000"
        formatoGrilla(4, 5) = "$ ###,###,##0"
        formatoGrilla(4, 6) = "00"
        formatoGrilla(4, 7) = "00"
        formatoGrilla(4, 8) = "0000"
        
        Rem LOCCKED
        formatoGrilla(5, 1) = "FALSE"
        formatoGrilla(5, 2) = "TRUE"
        formatoGrilla(5, 3) = "FALSE"
        formatoGrilla(5, 4) = "FALSE"
        formatoGrilla(5, 5) = "FALSE"
        formatoGrilla(5, 6) = "FALSE"
        formatoGrilla(5, 7) = "FALSE"
        formatoGrilla(5, 8) = "FALSE"
        
        Rem VALOR MINIMO
        formatoGrilla(6, 1) = ""
        formatoGrilla(6, 2) = ""
        formatoGrilla(6, 3) = ""
        formatoGrilla(6, 4) = ""
        formatoGrilla(6, 5) = ""
        formatoGrilla(6, 6) = ""
        formatoGrilla(6, 7) = ""
        formatoGrilla(6, 8) = ""
        
        Rem VALOR MAXIMO
        formatoGrilla(7, 1) = ""
        formatoGrilla(7, 2) = ""
        formatoGrilla(7, 3) = ""
        formatoGrilla(7, 4) = ""
        formatoGrilla(7, 5) = ""
        formatoGrilla(7, 6) = ""
        formatoGrilla(7, 7) = ""
        formatoGrilla(7, 8) = ""
        
        Rem ANCHO
        formatoGrilla(8, 1) = "6"
        formatoGrilla(8, 2) = "15"
        formatoGrilla(8, 3) = "8"
        formatoGrilla(8, 4) = "10"
        formatoGrilla(8, 5) = "12"
        formatoGrilla(8, 6) = "3"
        formatoGrilla(8, 7) = "3"
        formatoGrilla(8, 8) = "4"
        
        Cheques.Cols = col
        Cheques.Rows = row
        Cheques.AllowUserResizing = False
        Cheques.DisplayFocusRect = False
        Cheques.ExtendLastCol = True
        Cheques.BoldFixedCell = False
        Cheques.DrawMode = cellOwnerDraw
        Cheques.Appearance = Flat
        Cheques.ScrollBarStyle = Flat
        Cheques.FixedRowColStyle = Flat
        Cheques.BackColorFixed = RGB(90, 158, 214)
        Cheques.BackColorFixedSel = RGB(110, 180, 230)
        Cheques.BackColorBkg = RGB(90, 158, 214)
        Cheques.BackColorScrollBar = RGB(231, 235, 247)
        Cheques.BackColor1 = RGB(231, 235, 247)
        Cheques.BackColor2 = RGB(239, 243, 255)
        Cheques.GridColor = RGB(148, 190, 231)
        
        Cheques.Column(0).Width = 0
        For i = 1 To col - 1
            Cheques.Cell(0, i).text = formatoGrilla(1, i)
            Cheques.Column(i).Width = Val(formatoGrilla(8, i)) * (Cheques.Cell(0, i).Font.Size + 1.25)
            Cheques.Column(i).MaxLength = Val(formatoGrilla(2, i))
            Cheques.Column(i).FormatString = formatoGrilla(4, i)
            Cheques.Column(i).Locked = formatoGrilla(5, i)
            If formatoGrilla(3, i) = "N" Then
                Cheques.Column(i).Mask = cellNumeric
                Cheques.Column(i).Alignment = cellRightCenter
            End If
            If formatoGrilla(3, i) = "S" Then
                Cheques.Column(i).Mask = cellUpper
                Cheques.Column(i).Alignment = cellLeftCenter
            End If
            If formatoGrilla(3, i) = "D" Then
                Cheques.Column(i).Alignment = cellRightCenter
                Cheques.Column(i).CellType = cellCalendar
            End If
        Next i
        Cheques.Range(0, 6, 0, Cheques.Cols - 1).Merge
        Cheques.Range(0, 0, 0, Cheques.Cols - 1).Alignment = cellCenterCenter
        Cheques.Enabled = True
    End Sub
'****************************************************************************
'Formato de la Grilla Cheques
'****************************************************************************

'=============================================================================
'PASA LOS DATOS DE LOS CONTROLES A LA ESTRUCTURA
'=============================================================================
    Private Sub ctrltostruct()
        pc.c.FOLIO = dato1.text
        pc.c.rut = dato2.text & lblDV.Caption
        pc.c.fecha = dato5.text & "-" & dato4.text & "-" & dato3.text
        pc.c.TIPO = dato6.text
        pc.c.MONTO = Format(dato7.text, "########0")
        pc.c.glosa = ""
        pc.c.fechadeposito = fechadeposito
    
        pc.D.FOLIO = pc.c.FOLIO
        pc.D.tipopago = pc.c.TIPO
        pc.D.rut = pc.c.rut
        pc.D.fecha = pc.c.fecha
        
        pc.ch.FOLIO = pc.c.FOLIO
        pc.ch.codigolocal = empresaActiva
        pc.ch.fecha = pc.c.fecha
        pc.ch.rut = pc.c.rut
        pc.ch.tipodocumento = "PA"
        pc.ch.cajera = ""
        Call grabarPagoCliente(pc, modifica, Cheques, Pagar)
        If MsgBox("DESEA IMPRIMIR COMPROBANTE ", vbYesNo) = vbYes Then
            Call imprimir
        End If
        Call retorno
    End Sub
'=============================================================================
'PASA LOS DATOS DE LOS CONTROLES A LA ESTRUCTURA
'=============================================================================

'=============================================================================
'PASA LOS DATOS DE LA ESTRUCTURA A LOS CONTROLES
'=============================================================================
    Private Sub structtoctrl()
        dato1.text = pc.c.FOLIO
        dato2.text = pc.c.rut
        lblDV.Caption = rut(dato2.text)
        lblNombre.Caption = leerNombreCliente(pc.c.rut)
        dato3.text = Format(pc.c.fecha, "dd")
        dato4.text = Format(pc.c.fecha, "mm")
        dato5.text = Format(pc.c.fecha, "yyyy")
        dato6.text = pc.c.TIPO
        lblForma.Caption = leerFormaPago(dato6.text)
        dato7.text = pc.c.MONTO
        fechadeposito = pc.c.fechadeposito
        Call DeshabilitarCajas(Me)
        dato1.SetFocus
    End Sub
'=============================================================================
'PASA LOS DATOS DE LA ESTRUCTURA A LOS CONTROLES
'=============================================================================

    Private Sub frmGrabar_BarMouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
        frmGrabar.ColorBarraAbajo = &HFFC0C0
        frmGrabar.ColorBarraArriba = &H800000
        frmGrabar.CaptionEstilo3D = Raised
    End Sub

    Private Sub frmGrabar_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
        Dim diferencia As Double
        frmGrabar.ColorBarraAbajo = &H800000
        frmGrabar.ColorBarraArriba = &HFFC0C0
        frmGrabar.CaptionEstilo3D = Inserted
        diferencia = CDbl(lbltotalabonos.Caption) - CDbl(dato7.text)
        If diferencia = 0 Then
            Call ctrltostruct
'            If MsgBox("DESEA IMPRIMIR COMPROBANTE ", vbYesNo) = vbYes Then
'                Call imprimir
'            End If
'            Call retorno
        Else
            If diferencia < 0 Then
                If MsgBox("¿Desea dejar el vuelto como saldo a favor?", vbYesNo, "Pregunta") = vbYes Then
                    Call grabarSaldo(diferencia)
                    Call ctrltostruct
                Else
                    dato7.text = Format(lbltotalabonos.Caption, "########0")
                    Call ctrltostruct
                    Call MsgBox("Vuelto: " & Format(diferencia * -1, "$ ###,###,##0"), vbOKOnly, "Información")
                End If
            Else
                Call MsgBox("la cantidad abonada es menor que la cantidad pagada" & vbCrLf & "Por favor ajuste el abono o el pago", vbOKOnly, "Error")
                Pagar.Cell(1, 4).SetFocus
            End If
        End If
    End Sub

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
    End Select
End Sub

'=============================================================================
'OPCIONES
'=============================================================================
    Private Sub opciones_FSCommand(ByVal command As String, ByVal args As String)
        Select Case command
            Case "modifica"
                Call modificar
            Case "elimina"
                Call ELIMINAR
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
        modifica = True
        frmGrabar.Visible = True
        'Call HabilitarCajas(Me, modifica)
        If Cheques.Rows > 1 Then
            Cheques.Range(1, 1, Cheques.Rows - 1, Cheques.Cols - 1).Locked = False
            Cheques.Cell(1, Cheques.Cols - 3).SetFocus
        End If
    End Sub
    
    Private Sub ELIMINAR()
        frmglosaeliminacion.Show vbModal
        Call eliminarPagoCliente(pc, Me.Pagar)
        Call eliminarSaldo
        Call retorno
        Call HabilitarCajas(Me, modifica)
        dato1.SetFocus
    End Sub

    Private Sub retorno()
        Call LimpiarCajas(Me)
        Call LimpiarLabels(Me)
        modifica = False
        escribir = False
        Call DeshabilitarCajas(Me)
        If Documentos.Rows > 1 Then
            Documentos.Range(1, 0, Documentos.Rows - 1, Documentos.Cols - 1).ClearText
            Documentos.Rows = 1
        End If
        lblTotalCheques.Caption = "$ 0"
        If Pagar.Rows > 1 Then
            Pagar.Range(1, 0, Pagar.Rows - 1, Pagar.Cols - 1).ClearText
            Pagar.Rows = 1
        End If
        lblTotalDocumentos.Caption = "$ 0"
        If Cheques.Rows > 1 Then
            Cheques.Range(1, 0, Cheques.Rows - 1, Cheques.Cols - 1).ClearText
            Cheques.Rows = 1
        End If
        lbldiferencia.Caption = "$ 0"
        lbldiferencia.Caption = "$ 0"
        dato1.text = leer_Ultimo_Folio("numero", "sv_pagos_cabeza_" & empresaActiva, dato1.MaxLength, ventasRubro, "local = '" & empresaActiva & "'")
        dato1.SetFocus
        frmGrabar.Visible = False
        fechadeposito = ""
    End Sub
        
    Private Sub anterior()
        If leerPagoCliente(pc, dato1.text, "<", data, Cheques, Pagar, lblTotalCheques, lbldiferencia, lbldiferencia) = True Then
            structtoctrl
            frmGrabar.Visible = False
            If dato6.text = "3" Then
                frmDeposito.Visible = True
                txtDia.text = Format(fechadeposito, "dd")
                txtMes.text = Format(fechadeposito, "mm")
                txtAño.text = Format(fechadeposito, "yyyy")
            Else
                frmDeposito.Visible = False
            End If
            If Cheques.Rows > 1 Then
                Cheques.Range(1, 1, Cheques.Rows - 1, Cheques.Cols - 1).Locked = True
            End If
            Pagar.Range(1, 1, Pagar.Rows - 1, Pagar.Cols - 1).Locked = True
            Call sumarPagos
        End If
    End Sub
    
    Private Sub siguiente()
        If leerPagoCliente(pc, dato1.text, ">", data, Cheques, Pagar, lblTotalCheques, lbldiferencia, lbldiferencia) = True Then
            structtoctrl
            frmGrabar.Visible = False
            If dato6.text = "3" Then
                frmDeposito.Visible = True
                txtDia.text = Format(fechadeposito, "dd")
                txtMes.text = Format(fechadeposito, "mm")
                txtAño.text = Format(fechadeposito, "yyyy")
            Else
                frmDeposito.Visible = False
            End If
            If Cheques.Rows > 1 Then
                Cheques.Range(1, 1, Cheques.Rows - 1, Cheques.Cols - 1).Locked = True
            End If
            Pagar.Range(1, 1, Pagar.Rows - 1, Pagar.Cols - 1).Locked = True
            Call sumarPagos
        End If
    End Sub
    
    Private Sub imprimir()
        Dim i As Long
        Dim tabla As String
        Dim saldo As Double

        impresion.Rows = 1
        impresion.AutoRedraw = False
        
        impresion.Column(0).Width = 0
        impresion.Column(1).Width = 80
        impresion.Column(2).Width = 90
        impresion.Column(3).Width = 70
        impresion.Column(4).Width = 80
        impresion.Column(5).Width = 80
        impresion.Column(6).Width = 80
        
        impresion.PageSetup.HeaderMargin = 1
        impresion.PageSetup.TopMargin = 1
        
        Call cargaCabeza("COMPROBANTE DE PAGO", empresaActiva, impresion)
        
        impresion.AddItem "NUMERO PAGO: " & vbTab & vbTab & dato1.text & vbTab & vbTab & "FECHA: " & vbTab & dato3.text & "-" & dato4.text & "-" & dato5.text, True
        impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, 2).Merge
        impresion.Cell(impresion.Rows - 1, 1).Font.Bold = True
        impresion.Cell(impresion.Rows - 1, 5).Font.Bold = True
        impresion.AddItem "", True
        
        impresion.AddItem "RUT: " & vbTab & dato2.text & "-" & lblDV.Caption & vbTab & vbTab & vbTab & "CUPO" & vbTab & Format(leerCupoCliente(dato2.text & lblDV.Caption), "$ ###,###,##0"), True
        impresion.Cell(impresion.Rows - 1, 1).Font.Bold = True
        impresion.Cell(impresion.Rows - 1, 5).Font.Bold = True
        impresion.AddItem leerNombreCliente(dato2.text & lblDV.Caption), True
        impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, 6).Merge
        impresion.AddItem "", True
        
        impresion.AddItem "FORMA DE PAGO: " & vbTab & vbTab & lblForma.Caption, True
        impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, 2).Merge
        impresion.Cell(impresion.Rows - 1, 1).Font.Bold = True
        impresion.AddItem "", True
        
        If dato6.text = "2" Then
            impresion.AddItem "LISTADO DE CHEQUES", True
            impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, 6).Merge
            impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, 6).Alignment = cellCenterCenter
            impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, 6).FontBold = True
            impresion.AddItem "NUMERO" & vbTab & vbTab & "MONTO" & vbTab & vbTab & "VENCIMIENTO", True
            impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, 2).Merge
            impresion.Range(impresion.Rows - 1, 3, impresion.Rows - 1, 4).Merge
            impresion.Range(impresion.Rows - 1, 5, impresion.Rows - 1, 6).Merge
            impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, 6).Alignment = cellCenterCenter
            impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, 6).Borders(cellEdgeTop) = cellThin
            impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, 6).Borders(cellEdgeBottom) = cellThin
            impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, 6).Borders(cellEdgeLeft) = cellThin
            impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, 6).Borders(cellEdgeRight) = cellThin
            impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, 6).Borders(cellInsideVertical) = cellThin
            impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, 6).FontBold = True
            For i = 1 To Cheques.Rows - 1
                impresion.AddItem Cheques.Cell(i, 3).text & vbTab & vbTab & Format(Cheques.Cell(i, 5).text, "$ ###,###,##0") & vbTab & vbTab & Cheques.Cell(i, 6).text & "-" & Cheques.Cell(i, 7).text & "-" & Cheques.Cell(i, 8).text, True
                impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, 2).Merge
                impresion.Range(impresion.Rows - 1, 3, impresion.Rows - 1, 4).Merge
                impresion.Range(impresion.Rows - 1, 5, impresion.Rows - 1, 6).Merge
                impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, 6).Alignment = cellRightCenter
            Next i
            impresion.AddItem "", True
        End If
        
        impresion.AddItem "LISTADO DE DOCUMENTOS CANCELADOS O ABONADOS", True
        impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, 6).Merge
        impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, 6).Alignment = cellCenterCenter
        impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, 6).FontBold = True
        impresion.AddItem "DOCUMENTO" & vbTab & "VENCIMIENTO" & vbTab & "MONTO" & vbTab & vbTab & "ABONO / DOCUMENTO", True
        'impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, 2).Merge
        impresion.Range(impresion.Rows - 1, 3, impresion.Rows - 1, 4).Merge
        impresion.Range(impresion.Rows - 1, 5, impresion.Rows - 1, 6).Merge
        impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, 6).Alignment = cellCenterCenter
        impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, 6).Borders(cellEdgeTop) = cellThin
        impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, 6).Borders(cellEdgeBottom) = cellThin
        impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, 6).Borders(cellEdgeLeft) = cellThin
        impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, 6).Borders(cellEdgeRight) = cellThin
        impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, 6).Borders(cellInsideVertical) = cellThin
        impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, 6).FontBold = True
        For i = 1 To Pagar.Rows - 1
            If Pagar.Cell(i, 1).text <> "NV" Then
                impresion.AddItem Pagar.Cell(i, 1).text & " " & Pagar.Cell(i, 2).text & vbTab & leerVencimiento(Pagar.Cell(i, 1).text, Pagar.Cell(i, 2).text) & vbTab & Format(Pagar.Cell(i, 3).text, "$ ###,###,##0") & vbTab & vbTab & Format(Pagar.Cell(i, 4).text, "$ ###,###,##0"), True
            Else
                impresion.AddItem Pagar.Cell(i, 1).text & " " & Pagar.Cell(i, 2).text & vbTab & leerVencimiento(Pagar.Cell(i, 1).text, Pagar.Cell(i, 2).text) & vbTab & Format(Pagar.Cell(i, 3).text, "$ ###,###,##0") & vbTab & vbTab & Format(Pagar.Cell(i, 4).text, "0000000000"), True
            End If
            'impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, 2).Merge
            impresion.Range(impresion.Rows - 1, 3, impresion.Rows - 1, 4).Merge
            impresion.Range(impresion.Rows - 1, 5, impresion.Rows - 1, 6).Merge
            impresion.Range(impresion.Rows - 1, 2, impresion.Rows - 1, 6).Alignment = cellRightCenter
        Next i
        
        impresion.AddItem vbTab & vbTab & vbTab & "TOTAL: " & vbTab & vbTab & Format(lbltotalabonos.Caption, "$ ###,###,##0")
        impresion.Range(impresion.Rows - 1, 3, impresion.Rows - 1, 6).Alignment = cellRightCenter
        impresion.Cell(impresion.Rows - 1, 4).Font.Bold = True
        impresion.Range(impresion.Rows - 1, 4, impresion.Rows - 1, 6).Borders(cellEdgeTop) = cellThin
        impresion.AddItem "", True
        impresion.AddItem "", True
        
        If chk1.Value = 1 Then
            tabla = "SELECT CONCAT(tipo, ' ', numero, '" & vbTab & "', DATE_FORMAT(vencimiento,'%d-%m-%Y'), '" & vbTab & "', CONCAT('$ ', FORMAT(monto,0)), '" & vbTab & vbTab & "', CONCAT('$ ', FORMAT(monto-abono,0))) AS item, monto - abono AS saldo "
            tabla = tabla & "FROM sv_documentos_cobranza_" & empresaActiva & " "
            tabla = tabla & "WHERE local = '" & empresaActiva & "' AND rut = '" & dato2.text & lblDV.Caption & "' AND monto > abono ORDER BY numero ASC"
            Call ConectarControlData(pendientes, servidor, baseVentas & empresaActiva, usuario, password, tabla)
            If pendientes.Recordset.RecordCount > 0 Then
                pendientes.Recordset.MoveFirst
                impresion.AddItem "LISTADO DE DOCUMENTOS PENDIENTES", True
                impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, 6).Merge
                impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, 6).Alignment = cellCenterCenter
                impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, 6).FontBold = True
                impresion.AddItem "DOCUMENTO" & vbTab & "VENCIMIENTO" & vbTab & "MONTO" & vbTab & vbTab & "SALDO", True
                'impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, 2).Merge
                impresion.Range(impresion.Rows - 1, 3, impresion.Rows - 1, 4).Merge
                impresion.Range(impresion.Rows - 1, 5, impresion.Rows - 1, 6).Merge
                impresion.Range(impresion.Rows - 1, 2, impresion.Rows - 1, 6).Alignment = cellCenterCenter
                impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, 6).Borders(cellEdgeTop) = cellThin
                impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, 6).Borders(cellEdgeBottom) = cellThin
                impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, 6).Borders(cellEdgeLeft) = cellThin
                impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, 6).Borders(cellEdgeRight) = cellThin
                impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, 6).Borders(cellInsideVertical) = cellThin
                impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, 6).FontBold = True
                saldo = 0
                While Not pendientes.Recordset.EOF
                    saldo = saldo + CDbl(pendientes.Recordset.Fields("saldo"))
                    impresion.AddItem Replace(pendientes.Recordset.Fields("item"), ",", "."), True
                    'impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, 2).Merge
                    impresion.Range(impresion.Rows - 1, 3, impresion.Rows - 1, 4).Merge
                    impresion.Range(impresion.Rows - 1, 5, impresion.Rows - 1, 6).Merge
                    impresion.Range(impresion.Rows - 1, 2, impresion.Rows - 1, 6).Alignment = cellRightCenter
                    pendientes.Recordset.MoveNext
                Wend
                impresion.AddItem vbTab & vbTab & vbTab & "TOTAL: " & vbTab & vbTab & Format(saldo, "$ ###,###,##0")
                impresion.Range(impresion.Rows - 1, 3, impresion.Rows - 1, 6).Alignment = cellRightCenter
                impresion.Cell(impresion.Rows - 1, 4).Font.Bold = True
                impresion.Range(impresion.Rows - 1, 4, impresion.Rows - 1, 6).Borders(cellEdgeTop) = cellThin
            End If
            impresion.AddItem "", True
            impresion.AddItem "", True
            impresion.AddItem "", True
            impresion.AddItem "", True
            'impresion.Images.Remove ("publicidad")
            'impresion.Images.Add ruta, "publicidad"
            'impresion.Cell(impresion.Rows - 1, 1).SetImage "publicidad"
            impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, 6).Merge
            impresion.RowHeight(impresion.Rows - 1) = 250
        End If
        
        impresion.AutoRedraw = True
        impresion.Refresh
        
        Call verificaImpresora(5, impresion)
    End Sub
'=============================================================================
'OPCIONES
'=============================================================================

Private Function verificarDocumento() As Boolean
    Dim i As Long
    Dim TIPO As String
    Dim NUMERO As String
    TIPO = dato8.text
    NUMERO = dato9.text
    For i = 1 To Documentos.Rows - 1
        If TIPO = Documentos.Cell(i, 1).text Then
            If NUMERO = Documentos.Cell(i, 2).text Then
                Documentos.Range(i, 1, i, Documentos.Cols - 1).Selected
                Call Documentos_DblClick
                verificarDocumento = True
                Exit For
            Else
                verificarDocumento = False
            End If
        End If
    Next i
End Function

Private Sub Pagar_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Pagar.ActiveCell.col = 1 Then
            sumarPagos
            dato8.SetFocus
        End If
    End If
End Sub

Private Sub grabarSaldo(ByVal diferencia As Double)
    Dim cSql As rdoQuery
    Set cSql = New rdoQuery
    Set cSql.ActiveConnection = ventas
    cSql.sql = "INSERT INTO sv_maestroclientes_saldos (rut, sucursal, numeropago, monto) "
    cSql.sql = cSql.sql & "VALUES('" & dato2.text & lblDV.Caption & "', '0', '" & dato1.text & "', '" & diferencia & "') "
    cSql.sql = cSql.sql & "ON DUPLICATE KEY UPDATE monto = '" & diferencia & "' "
    cSql.Execute
        Call sincronizadatos(cSql.sql, ventas)
    cSql.Close
    Set cSql = Nothing
End Sub

Private Sub eliminarSaldo()
    Dim cSql As rdoQuery
    Set cSql = New rdoQuery
    Set cSql.ActiveConnection = ventas
    cSql.sql = "DELETE FROM sv_maestroclientes_saldos "
    cSql.sql = cSql.sql & "WHERE rut = '" & dato2.text & lblDV.Caption & "' AND numeropago = '" & dato1.text & "'"
    cSql.Execute
    cSql.Close
    Set cSql = Nothing
End Sub
