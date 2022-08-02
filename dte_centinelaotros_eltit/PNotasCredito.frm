VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash10c.ocx"
Object = "{4777436C-EB8C-4596-98A8-EBCF98683969}#1.0#0"; "FlexCell.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{ADBBAED6-B16F-40DC-80DF-B44910CBA76C}#1.0#0"; "Frame-Xp.ocx"
Begin VB.Form PNotasCredito 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pantalla de Notas de Crédito"
   ClientHeight    =   9660
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13710
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9660
   ScaleWidth      =   13710
   StartUpPosition =   2  'CenterScreen
   Begin XPFrame.FrameXp frmTipo 
      Height          =   1155
      Left            =   3240
      TabIndex        =   60
      Top             =   -120
      Visible         =   0   'False
      Width           =   2715
      _ExtentX        =   4789
      _ExtentY        =   2037
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
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " 2 - NC Boleta"
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
         TabIndex        =   64
         Top             =   840
         Width           =   2475
      End
      Begin VB.Label lbl23 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " 1 - NC Factura"
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
         TabIndex        =   61
         Top             =   480
         Width           =   2475
      End
   End
   Begin XPFrame.FrameXp frmDatos 
      Height          =   3135
      Left            =   120
      TabIndex        =   15
      Top             =   120
      Width           =   13455
      _ExtentX        =   23733
      _ExtentY        =   5530
      BackColor       =   16744576
      Caption         =   "Datos de la Nota de Crédito"
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
         Top             =   480
         Width           =   495
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
         Left            =   1920
         MaxLength       =   9
         TabIndex        =   5
         Tag             =   "proveedor"
         Top             =   840
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
         Left            =   12720
         MaxLength       =   4
         TabIndex        =   4
         Tag             =   "proveedor"
         Top             =   480
         Width           =   615
      End
      Begin VB.TextBox dato8 
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
         Left            =   11520
         MaxLength       =   12
         TabIndex        =   7
         Tag             =   "proveedor"
         Top             =   1920
         Width           =   1815
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
         TabIndex        =   3
         Tag             =   "proveedor"
         Top             =   480
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
         Left            =   12000
         MaxLength       =   2
         TabIndex        =   2
         Tag             =   "proveedor"
         Top             =   480
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
         Left            =   8520
         MaxLength       =   10
         TabIndex        =   1
         Tag             =   "proveedor"
         Top             =   480
         Width           =   1455
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
         Left            =   1890
         MaxLength       =   2
         TabIndex        =   8
         Tag             =   "proveedor"
         Top             =   2280
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
         Left            =   5880
         MaxLength       =   1
         TabIndex        =   6
         Tag             =   "proveedor"
         Top             =   840
         Width           =   615
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
         Left            =   2460
         TabIndex        =   59
         Top             =   480
         Width           =   4035
      End
      Begin VB.Label lbl10 
         Appearance      =   0  'Flat
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Vencimiento"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   4920
         TabIndex        =   39
         Top             =   1920
         Width           =   1695
      End
      Begin VB.Label lbl2 
         Appearance      =   0  'Flat
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Número"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   6720
         TabIndex        =   38
         Top             =   480
         Width           =   1695
      End
      Begin VB.Label lbl12 
         Appearance      =   0  'Flat
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Vendedor"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   120
         TabIndex        =   37
         Top             =   2280
         Width           =   1695
      End
      Begin VB.Label lbl11 
         Appearance      =   0  'Flat
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Nota de Pedido"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   9720
         TabIndex        =   36
         Top             =   1920
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
         TabIndex        =   35
         Top             =   1920
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
         TabIndex        =   34
         Top             =   1200
         Width           =   1695
      End
      Begin VB.Label lbl5 
         Appearance      =   0  'Flat
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Razón Social"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   6720
         TabIndex        =   33
         Top             =   840
         Width           =   1695
      End
      Begin VB.Label lbl4 
         Appearance      =   0  'Flat
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Rut Cliente"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   120
         TabIndex        =   32
         Top             =   840
         Width           =   1695
      End
      Begin VB.Label lbl3 
         Appearance      =   0  'Flat
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Fecha"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   10200
         TabIndex        =   31
         Top             =   480
         Width           =   1695
      End
      Begin VB.Label lbl1 
         Appearance      =   0  'Flat
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Tipo"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   120
         TabIndex        =   30
         Top             =   480
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
         TabIndex        =   29
         Top             =   1560
         Width           =   1695
      End
      Begin VB.Label lbl8 
         Appearance      =   0  'Flat
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Ciudad"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   6840
         TabIndex        =   28
         Top             =   1560
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
         Left            =   3360
         TabIndex        =   27
         Top             =   840
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
         Left            =   8520
         TabIndex        =   26
         Top             =   840
         Width           =   4815
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
         Left            =   1920
         TabIndex        =   25
         Top             =   1200
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
         Left            =   1920
         TabIndex        =   24
         Top             =   1560
         Width           =   4695
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
         Left            =   8640
         TabIndex        =   23
         Top             =   1560
         Width           =   4695
      End
      Begin VB.Label lblDias 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Días"
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
         Left            =   1920
         TabIndex        =   22
         Top             =   1920
         Width           =   2295
      End
      Begin VB.Label lblDia 
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
         Left            =   6720
         TabIndex        =   21
         Top             =   1920
         Width           =   495
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
         Left            =   3510
         TabIndex        =   20
         Top             =   2280
         Width           =   9825
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Sucursal"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   4080
         TabIndex        =   19
         Top             =   840
         Width           =   1695
      End
      Begin VB.Label lblMes 
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
         Left            =   7320
         TabIndex        =   18
         Top             =   1920
         Width           =   495
      End
      Begin VB.Label lblAño 
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
         Left            =   7920
         TabIndex        =   17
         Top             =   1920
         Width           =   975
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
         Left            =   3360
         TabIndex        =   16
         Top             =   2280
         Visible         =   0   'False
         Width           =   375
      End
   End
   Begin XPFrame.FrameXp frmDetalle 
      Height          =   4215
      Left            =   120
      TabIndex        =   14
      Top             =   3360
      Width           =   13455
      _ExtentX        =   23733
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
         Height          =   3855
         Left            =   0
         TabIndex        =   9
         Top             =   315
         Width           =   13470
         _ExtentX        =   23760
         _ExtentY        =   6800
         Cols            =   5
         DefaultFontSize =   8.25
         Rows            =   30
      End
   End
   Begin VB.TextBox pivote 
      Height          =   285
      Left            =   120
      MaxLength       =   13
      TabIndex        =   13
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
      Height          =   1935
      Left            =   8160
      TabIndex        =   40
      Top             =   7620
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   3413
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
         TabIndex        =   11
         Tag             =   "proveedor"
         Text            =   "0"
         Top             =   1170
         Width           =   1215
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
         Left            =   1440
         MaxLength       =   5
         TabIndex        =   10
         Tag             =   "proveedor"
         Text            =   "0"
         Top             =   840
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
         Left            =   4080
         TabIndex        =   50
         Top             =   1560
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
         Left            =   4080
         TabIndex        =   51
         Top             =   1200
         Width           =   1215
      End
      Begin VB.Label lbl18 
         Appearance      =   0  'Flat
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " IHA"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   2760
         TabIndex        =   52
         Top             =   1200
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
         Left            =   4080
         TabIndex        =   49
         Top             =   840
         Width           =   1215
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
         Width           =   1215
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
         Width           =   1215
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
         Top             =   1200
         Width           =   1215
      End
      Begin VB.Label lbl13 
         Appearance      =   0  'Flat
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Sub Total"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   120
         TabIndex        =   44
         Top             =   480
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
         TabIndex        =   43
         Top             =   840
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
         Left            =   4080
         TabIndex        =   42
         Top             =   480
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
         Left            =   1440
         TabIndex        =   41
         Top             =   480
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
      TabIndex        =   58
      Top             =   0
      Width           =   555
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
      Left            =   6165
      TabIndex        =   63
      Top             =   7650
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Label lbl26 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " * Fin Nota credito"
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
      Height          =   690
      Left            =   6300
      TabIndex        =   62
      Top             =   8145
      Visible         =   0   'False
      Width           =   1800
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Para ver como se pagó el documento presione ""P"""
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1275
      Left            =   3690
      TabIndex        =   57
      Top             =   7590
      Visible         =   0   'False
      Width           =   1755
   End
   Begin VB.Label lblHasta 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Hasta: 1234567890"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   5580
      TabIndex        =   56
      Top             =   7920
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Label lblDesde 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Desde: 1234567890"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   2760
      TabIndex        =   55
      Top             =   7920
      Visible         =   0   'False
      Width           =   2475
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
      Left            =   120
      TabIndex        =   54
      Top             =   7260
      Width           =   7935
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash opciones 
      Height          =   1455
      Left            =   300
      TabIndex        =   12
      Top             =   8160
      Width           =   6375
      _cx             =   11245
      _cy             =   2566
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
Attribute VB_Name = "PNotasCredito"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Private c As Cliente
    Private v As venta
    Private p As pagos
    Private D As docManual
    Private formatogrilla(10, 10) As String
    Private modifica As Boolean
    Private vacio As Boolean
    Private fila As Long
    Private columna As Long
    Private nula As Boolean
    Private lectura As Boolean
    Public imprimio As Boolean
    Public desde As String
    Public hasta As String
    Private tipoprecio As String
    Private fecha As String
    Private formato As String
    'Private segurity As Boolean

'============================================================
'MANEJO DE LOS CONTOLES
'============================================================
    '========================================================
    'GotFocus
    '========================================================
    Private Sub dato1_GotFocus()
        frmTipo.Visible = True
        Call selecciona(dato1)
    End Sub
    
    Private Sub dato2_GotFocus()
        Call VerificarCajas(Me, dato2)
        Call selecciona(dato2)
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
        Principal.barraEstado.Panels(2).text = "F2: Ayuda Cliente"
    End Sub
    
    Private Sub DATO7_GotFocus()
        Call VerificarCajas(Me, dato7)
        Call selecciona(dato7)
        Principal.barraEstado.Panels(2).text = "F2: Ayuda Sucursal"
    End Sub
    
    Private Sub DATO8_GotFocus()
        Call VerificarCajas(Me, dato8)
        Call selecciona(dato8)
        Principal.barraEstado.Panels(2).text = "F2: Ayuda Nota de Pedido"
    End Sub
    
    Private Sub dato9_GotFocus()
        Call VerificarCajas(Me, dato9)
        Call selecciona(dato9)
        Principal.barraEstado.Panels(2).text = "F2: Ayuda Vendedor"
    End Sub
    
    Private Sub dato10_GotFocus()
        Call VerificarCajas(Me, dato10)
        Call selecciona(dato10)
    End Sub
    
    Private Sub dato11_GotFocus()
        Call VerificarCajas(Me, dato11)
        Call selecciona(dato11)
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
                dato1.text = "NF"
                tipoprecio = "01"
            Case 98, 50
                dato1.text = "NB"
                tipoprecio = "01"
            Case Else
                Call Flechas(KeyCode, dato1)
        End Select
    End Sub
    
    Private Sub dato2_KeyDown(KeyCode As Integer, Shift As Integer)
        Call Flechas(KeyCode, dato1)
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
    
    Private Sub dato8_KeyDown(KeyCode As Integer, Shift As Integer)
        Call Flechas(KeyCode, dato7)
    End Sub
    
    Private Sub dato9_KeyDown(KeyCode As Integer, Shift As Integer)
        If KeyCode = vbKeyF2 Then
            Call ayudaVendedores(dato9)
        Else
            Call Flechas(KeyCode, dato8)
        End If
    End Sub
    
    Private Sub dato10_KeyDown(KeyCode As Integer, Shift As Integer)
        Call Flechas(KeyCode, dato9)
       
    End Sub
    
    Private Sub dato11_KeyDown(KeyCode As Integer, Shift As Integer)
        Call Flechas(KeyCode, dato10)
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
            lbl20.Visible = True
            lbl26.Visible = True
            
            lblDocumento.Caption = leerNombreDocumento(dato1.text)
            If lblDocumento.Caption <> "" Then
                dato2.text = leerUltimoFolio(dato1.text)
                dato2.text = ceros(dato2)
                SendKeys "{Tab}"
            End If
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
            If leerVenta(v, dato1.text, dato2.text, "=", data, detalle) = True Then
                lectura = True
                Call structtoctrl
                nula = leerDocumentoNulo(dato1.text, dato2.text)
                If nula = True Then
                    lblNulo.Caption = "DOCUMENTO ANULADO"
                Else
                    lblNulo.Caption = ""
                End If
            Else
                lectura = False
                detalle.SelectionMode = cellSelectionFree
'                If detalle.Rows <= 1 Then
                    detalle.Rows = 1
                    formato = "0000000000"
                    detalle.AddItem vbTab & vbTab & "1" & vbTab & "0" & vbTab & "0" & vbTab & "0" & vbTab & "0", True
'                    detalle.AddItem vbTab & vbTab & "1" & vbTab & "0" & vbTab & "0" & vbTab & "0" & vbTab & "0", True
             
'                    detalle.AddItem " " & vbTab & " " & vbTab & "1" & vbTab & "0" & vbTab & "0" & vbTab & "0" & vbTab & "0", True
'                End If
                Call HabilitarCajas(Me, modifica)
            End If
            SendKeys "{Tab}"
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
            If dato1.text <> "ZE" Then
                SendKeys "{Tab}"
            Else
                DesdeHasta.Show vbModal
                dato6.text = "999999999"
                lbldv.Caption = "6"
                dato7.text = "0"
                dato8.text = "0000000000"
                dato9.text = "0000000000"
                Call DATO7_KeyPress(13)
                'lblVendedor.Caption = leerNombreVendedor(dato9.text & lblDVV.Caption)
                SendKeys "{Tab}"
                SendKeys "{Tab}"
                SendKeys "{Tab}"
            End If
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
    
    Private Sub DATO7_KeyPress(KeyAscii As Integer)
        KeyAscii = esNumero(KeyAscii)
        If KeyAscii = 13 Then
            dato7.text = ceros(dato7)
            If LEERCLIENTE(c, dato6.text & lbldv.Caption, dato7.text, "=") = True Then
                structtoctrlCliente
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
    
    Private Sub DATO8_KeyPress(KeyAscii As Integer)
        KeyAscii = esNumero(KeyAscii)
        If KeyAscii = 13 Then
            dato8.text = ceros(dato8)
            SendKeys "{Tab}"
        End If
    End Sub
    
    Private Sub DATO9_KeyPress(KeyAscii As Integer)
        KeyAscii = esNumero(KeyAscii)
        If KeyAscii = 13 And dato9.text <> "" Then
            dato9.text = ceros(dato9)
            'lblDVV.Caption = rut(dato9.text)
            lblVendedor.Caption = leerNombreVendedor(dato9.text)
            If lblVendedor.Caption <> "" Then
                detalle.Enabled = True
                detalle.SelectionMode = cellSelectionFree
                If detalle.Rows > 1 And lectura = False Then
                    detalle.Cell(detalle.Rows - 1, 1).SetFocus
                End If
            End If
        End If
    End Sub
    
    Private Sub DATO10_KeyPress(KeyAscii As Integer)
        Dim desc As Double
        Dim total As Double
        Dim cadena As String
        Dim deci As String
        Dim i As Long
         KeyAscii = esNumeroDecimal(dato10, KeyAscii)
         If KeyAscii = 13 And dato10.text <> "" Then
         dato10.text = dato10.text * -1
         total = CDbl(lblSub.Caption)
         
         If dato10.text <> "" Then
            
            desc = total * CDbl(dato10.text) / 100
            cadena = Format(desc, "########0.0")
            deci = Right(cadena, 1)
            If deci >= 5 Then
                deci = 1
            Else
                deci = 0
            End If
            dato11.text = Val(desc) + CDbl(deci)
            For i = 1 To detalle.Rows - 1
                detalle.Cell(i, 6).text = dato10.text * -1
                If detalle.Cell(i, 1).text <> "0000000000100" Then
                    detalle.Cell(i, 7).text = CDbl(detalle.Cell(i, 3).text) * CDbl(detalle.Cell(i, 5).text) - (CDbl(detalle.Cell(i, 5).text) * CDbl(detalle.Cell(i, 6).text) / 100) * CDbl(detalle.Cell(i, 3).text)
                Else
                    detalle.Cell(i, 7).text = CDbl(detalle.Cell(i, 5).text) - (CDbl(detalle.Cell(i, 5).text) * CDbl(detalle.Cell(i, 6).text) / 100)
                End If
            Next i
        Else
            dato10.text = "0"
            dato11.text = "0"
        End If
        Call sumaGrilla(detalle)
            SendKeys "{Tab}"
        End If
        If dato10.text = "" Then
            dato10.text = "0"
        End If
    End Sub
    
    Private Sub dato11_KeyPress(KeyAscii As Integer)
        KeyAscii = esNumero(KeyAscii)
        If KeyAscii = 13 And dato11.text <> "" Then
            lbl20.Visible = False
            lbl26.Visible = False
            
            Call ctrltostruct(True)
            Call ctrltostructCobranza
            If MsgBox("DESEA IMPRIMIR COMPROBANTE ", vbYesNo) = vbYes Then
                Call imprimeFactura(dato2.text, impresion, data)
            End If
            Call retorno
        End If
        If dato11.text = "" Then
            dato10.text = "0"
            dato11.text = "0"
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
    
    Private Sub dato10_KeyUp(KeyCode As Integer, Shift As Integer)
'        Dim desc As Double
'        Dim total As Double
'        Dim cadena As String
'        Dim deci As String
'        Dim i As Long
'
'        total = CDbl(lblSub.Caption)
'        If dato10.text <> "" Then
'
'            desc = total * CDbl(dato10.text * -1) / 100
'            cadena = Format(desc, "########0.0")
'            deci = Right(cadena, 1)
'            If deci >= 5 Then
'                deci = 1
'            Else
'                deci = 0
'            End If
'            dato11.text = Val(desc) + CDbl(deci)
'            For i = 1 To detalle.Rows - 1
'                detalle.Cell(i, 6).text = dato10.text * -1
'                If detalle.Cell(i, 1).text <> "0000000000100" Then
'                    detalle.Cell(i, 7).text = CDbl(detalle.Cell(i, 3).text) * CDbl(detalle.Cell(i, 5).text) - (CDbl(detalle.Cell(i, 5).text) * CDbl(detalle.Cell(i, 6).text) / 100) * CDbl(detalle.Cell(i, 3).text)
'                Else
'                    detalle.Cell(i, 7).text = CDbl(detalle.Cell(i, 5).text) - (CDbl(detalle.Cell(i, 5).text) * CDbl(detalle.Cell(i, 6).text) / 100)
'                End If
'            Next i
'        Else
'            dato10.text = "0"
'            dato11.text = "0"
'        End If
'        Call sumaGrilla(detalle)
    End Sub

    Private Sub dato11_KeyUp(KeyCode As Integer, Shift As Integer)
        Dim desc As Double
        Dim total As Double
        Dim i As Long
        
        total = CDbl(lblSub.Caption)
        If total = 0 Then
            total = 1
        End If
        If dato11.text <> "" Then
            desc = CDbl(dato11.text)
            desc = desc * 100 / total
            dato10.text = Round(desc, 2)
            For i = 1 To detalle.Rows - 1
                detalle.Cell(i, 6).text = dato10.text * -1
                If detalle.Cell(i, 1).text <> "0000000000100" Then
                    detalle.Cell(i, 7).text = CDbl(detalle.Cell(i, 3).text) * CDbl(detalle.Cell(i, 5).text) - (CDbl(detalle.Cell(i, 5).text) * CDbl(detalle.Cell(i, 6).text) / 100) * CDbl(detalle.Cell(i, 3).text)
                Else
                    detalle.Cell(i, 7).text = CDbl(detalle.Cell(i, 5).text) - (CDbl(detalle.Cell(i, 5).text) * CDbl(detalle.Cell(i, 6).text) / 100)
                End If
            Next i
        Else
            dato10.text = "0"
            dato11.text = "0"
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
    
    Private Sub dato7_LostFocus()
        Call limpiaBarra(2)
    End Sub
    
    Private Sub dato8_LostFocus()
        Call limpiaBarra(2)
    End Sub
    
    Private Sub dato9_LostFocus()
        Call limpiaBarra(2)
        vend = dato9.text
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
        If KeyCode = 27 Then
            Unload Me
        End If
'        If KeyCode = 80 Then
'            If lectura = True Then
'                Load detallePagos
'                With detallePagos
'                    .lectura = True
'                    .pagos.Rows = 1
'                    .pagos.AutoRedraw = False
'                    tabla = "SELECT tipopago, CONCAT('" & vbTab & "', monto, '" & vbTab & "', numerocheque, '" & vbTab & "', banco, '" & vbTab & "', cuentacorriente, '" & vbTab & "', IF(vencimiento <> '00-00-0000',DATE_FORMAT(vencimiento,'%d-%m-%Y'),'')) AS item "
'                    tabla = tabla & "FROM sv_documento_pagos "
'                    tabla = tabla & "WHERE local = '" & empresaActiva & "' AND tipo = '" & dato1.text & "' AND numero = '" & dato2.text & "'"
'                    tabla = tabla & "ORDER BY lineapago ASC"
'                    Call ConectarControlData(.data, servidor, baseVentas & rubro, usuario, password, tabla)
'                    If .data.Recordset.RecordCount > 0 Then
'                        .data.Recordset.MoveFirst
'                        While Not .data.Recordset.EOF
'                            tipo = .data.Recordset.Fields("tipopago")
'                            Select Case tipo
'                                Case "1"
'                                    tipo = "1 - Efectivo"
'                                Case "2"
'                                    tipo = "2 - Cheque"
'                                Case "3"
'                                    tipo = "3 - "
'                                Case "4"
'                                    tipo = "4 - "
'                                Case "5"
'                                    tipo = "5 - "
'                                Case "6"
'                                    tipo = "6 - Credito"
'                            End Select
'                            .pagos.AddItem tipo & .data.Recordset.Fields("item"), True
'                            .data.Recordset.MoveNext
'                        Wend
'                    End If
'                    '.pagos.Range(1, 1, .pagos.Rows - 1, .pagos.Cols - 1).Locked = True
'                    .pagos.AutoRedraw = True
'                    .pagos.Refresh
'                    .pagos.SelectionMode = cellSelectionByRow
'                    detallePagos.Show vbModal
'                End With
'            End If
'        End If
        If KeyCode = 38 Then
            If Screen.ActiveForm.ActiveControl.Name = "dato1" Then
                Unload Me
            End If
        End If
    End Sub
    
    Private Sub Form_Load()
        titCaption = Me.Caption
        'segurity = Not Verificar(usuarioSistema, passwordSistema)
        modifica = False
        nula = False
        imprimio = False
        lectura = False
        Call Centrar(Me)
        Call CARGAGRILLA(1, 9)
        dato1.text = "NF"
    End Sub
    
    Private Sub Form_Unload(Cancel As Integer)
        Principal.barraEstado.Panels(1).text = UCase(Principal.Caption)
        Call limpiaBarra(2)
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
        formatogrilla(1, 3) = "DOC. / CANT."
        formatogrilla(1, 4) = "     "
        formatogrilla(1, 5) = "MONTO / PRECIO"
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
        formatogrilla(4, 3) = ""
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
        formatogrilla(5, 5) = "FALSE"
        formatogrilla(5, 6) = "TRUE"
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
        
        detalle.Column(0).Width = 25
        detalle.Column(col).Width = 0
        detalle.Cell(0, col).text = "PCOSTO"
        detalle.Column(col).Locked = True
        
        detalle.Cell(0, 0).text = formatogrilla(1, 0)
        For i = 1 To col - 1
            detalle.Cell(0, i).text = formatogrilla(1, i)
            detalle.Column(i).Width = Val(formatogrilla(8, i)) * (detalle.Cell(0, i).Font.Size + 1.25)
            detalle.Column(i).MaxLength = Val(formatogrilla(2, i))
            detalle.Column(i).FormatString = formatogrilla(4, i)
            detalle.Column(i).Locked = formatogrilla(5, i)
            If formatogrilla(3, i) = "N" Then
                detalle.Column(i).Alignment = cellRightCenter
                If i <> 3 And i <> 5 Then
                    detalle.Column(i).Mask = cellNumeric
                End If
            Else
                detalle.Column(i).Alignment = cellLeftCenter
                detalle.Column(i).Mask = cellUpper
            End If
        Next i
        detalle.Range(0, 0, 0, detalle.Cols - 1).Alignment = cellCenterCenter
        'detalle.AddItem "0000000000100" & vbTab & vbTab & "1" & vbTab & "0" & vbTab & "0" & vbTab & "0" & vbTab & "0", True
        detalle.Enabled = True
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
        If detalle.ActiveCell.col = 1 And KeyCode = vbKeyF2 Then Call ayudaProducto(detalle, pivote): detalle.Cell(fila, columna).SetFocus
        Select Case KeyCode
            Case 13, 37, 38, 39, 40
                If detalle.ActiveCell.text <> "" Then
                    vacio = False
                    If detalle.ActiveCell.col = 5 Then
                        If CDbl(detalle.ActiveCell.text) > 0 Then
                            detalle.ActiveCell.text = detalle.ActiveCell.text
                        End If
                    End If
                Else
                    vacio = True
                End If
            Case 106
               dato10.SetFocus
            Case 109
                If fila > 1 Then
                    detalle.RemoveItem (fila)
                    Call sumaGrilla(detalle)
                End If
        End Select
        
    End Sub
    
    Private Sub detalle_KeyUp(KeyCode As Integer, Shift As Integer)
        If detalle.ActiveCell.col = 3 Or detalle.ActiveCell.col = 5 Then
            pivote.text = detalle.ActiveCell.text
            If pivote.text <> "" Then
                KeyCode = esNumeroDecimal(pivote, Asc(Right(pivote.text, 1)))
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
    
    Private Sub detalle_LeaveCell(ByVal row As Long, ByVal col As Long, NewRow As Long, NewCol As Long, Cancel As Boolean)
        Dim i As Integer
        Dim linea As String
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
            If col = 5 And NewCol = 6 Then
                If row = detalle.Rows - 1 Then
                    If dato1.text = "NV" Then
                        detalle.AddItem vbTab & vbTab & "1" & vbTab & "0" & vbTab & "0" & vbTab & "0" & vbTab & "0", True
                        NewRow = detalle.Rows - 1
                        NewCol = 1
                    End If
                Else
                    NewCol = 1
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
                pivote.MaxLength = 13
                pivote.text = detalle.Cell(row, 1).text
                pivote.text = ceros(pivote)
                If pivote.text = "0000000000100" Then
                    formato = "0000000000"
                Else
                    formato = "###,###,##0.00"
                End If
                detalle.Cell(row, 1).text = pivote.text
                detalle.Cell(row, detalle.Cols - 1).text = leerCostoProducto(detalle.Cell(row, 1).text)
                detalle.Cell(row, 2).text = leerNombreProducto(detalle.Cell(row, 1).text)
                If detalle.Cell(row, 2).text <> "" Then
                    NewCol = 3
                End If
                If detalle.Cell(row, 1).text <> "0000000000100" Then
                    detalle.Cell(row, 4).text = CDbl(leerUnidadesProducto(detalle.Cell(row, 1).text)) * CDbl(detalle.Cell(row, 3).text)
                Else
                    detalle.Cell(row, 4).text = leerUnidadesProducto(detalle.Cell(row, 1).text)
                End If
                detalle.Cell(row, 5).text = leerPrecioEspecial(detalle.Cell(row, 1).text)
                If Val(detalle.Cell(row, 5).text) = 0 Then
                    detalle.Cell(row, 5).text = leerPrecioProducto(detalle.Cell(row, 1).text, tipoprecio)
                End If
                If detalle.Cell(row, 3).text <> "" And detalle.Cell(row, 5).text <> "" Then
                    If detalle.Cell(row, 1).text <> "0000000000100" Then
                        detalle.Cell(row, 7).text = detalle.Cell(row, 3).text * detalle.Cell(row, 5).text
                    Else
                        detalle.Cell(row, 7).text = detalle.Cell(row, 5).text
                    End If
                Else
                    detalle.Cell(row, 7).text = "0"
                End If
            Else
                If col = 1 And NewCol = detalle.Cols - 2 Then
                    NewCol = 5
                End If
            End If
            If col = 3 And NewCol <> col Then
                If detalle.Cell(row, 3).text <> "" Then
                    pivote.text = detalle.Cell(row, 3).text
                    detalle.Cell(row, 3).text = Format(pivote.text, formato)
                    If NewCol > col Then
                        NewCol = 5
                    End If
                    If NewCol < col Then
                        NewCol = 1
                    End If
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
                    If detalle.Cell(row, 1).text <> "0000000000100" Then
                        detalle.Cell(row, 4).text = CDbl(leerUnidadesProducto(detalle.Cell(row, 1).text)) * CDbl(detalle.Cell(row, 3).text)
                        detalle.Cell(row, 7).text = detalle.Cell(row, 3).text * detalle.Cell(row, 5).text
                    Else
                        detalle.Cell(row, 4).text = leerUnidadesProducto(detalle.Cell(row, 1).text)
                        detalle.Cell(row, 7).text = detalle.Cell(row, 5).text
                    End If
                End If
            End If
            If col > 0 And row > 0 Then
                Call sumaGrilla(detalle)
            End If
        End If
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
        
        suma = 0
        sumaIVA = 0
        sumaIHA = 0
        sumaEXENTO = 0
        
        For i = 1 To detalle.Rows - 1
            CODIGO = lista.Cell(i, 1).text
            If CODIGO <> "" Then
                Select Case leerImpuestoProducto(CODIGO)
                    Case "IHA"
                        sumaIVA = sumaIVA + CDbl(detalle.Cell(i, 7).text)
                        sumaIHA = sumaIHA + CDbl(detalle.Cell(i, 7).text)
                    Case "IVA"
                        sumaIVA = sumaIVA + CDbl(detalle.Cell(i, 7).text)
                End Select
                If Val(detalle.Cell(i, 3).text) = 0 Then
                    detalle.Cell(i, 3).text = "0"
                End If
                If Val(detalle.Cell(i, 5).text) = 0 Then
                    detalle.Cell(i, 5).text = "0"
                End If
                If detalle.Cell(i, 1).text = "0000000000100" Then
                    suma = suma + CDbl(detalle.Cell(i, 7).text)
                Else
                    suma = suma + CDbl(detalle.Cell(i, 3).text) * CDbl(detalle.Cell(i, 5).text)
                End If
            End If
        Next i
        
        cad = Format(suma, "########0.0")
        deci = Right(cad, 1)
        If deci >= 5 Then
            deci = 1
        Else
            deci = 0
        End If
        suma = Val(cad) - CDbl(deci)
        
        
        lblSub.Caption = suma
        
        Select Case dato1.text
            Case "NB", "NZ"
                lblNeto.Caption = Round((CDbl(lblSub.Caption) - CDbl(dato11.text)) / (1 + iva / 100), 0)
                lblIVA.Caption = Round(CDbl(lblSub.Caption) - CDbl(dato11.text) - CDbl(lblNeto.Caption), 0)
                lblIHA.Caption = "0"
            Case Else
                lblNeto.Caption = Round(CDbl(lblSub.Caption) + CDbl(dato11.text), 0)
                lblIVA.Caption = Round(sumaIVA * iva / 100, 0)
                lblIHA.Caption = Round(sumaIHA * iha / 100, 0)
        End Select
        
        lblTotal.Caption = CDbl(lblNeto.Caption) + CDbl(lblIVA.Caption) + CDbl(lblIHA.Caption)
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
        v.cabeza.NUMERO = dato2.text
        v.cabeza.fecha = dato5.text & "-" & dato4.text & "-" & dato3.text
        v.cabeza.plazo = lblDias.Caption
        v.cabeza.vencimiento = lblAño.Caption & "-" & lblMes.Caption & "-" & lblDia.Caption
        v.cabeza.rut = dato6.text & lbldv.Caption
        v.cabeza.sucursal = dato7.text
        v.cabeza.cajera = dato9.text & lbldv.Caption
        v.cabeza.notapedido = dato8.text
        v.cabeza.notaventas = ""
        v.cabeza.ordencompra = ""
        v.cabeza.subtotal = Format(lblSub.Caption, "########0")
        v.cabeza.neto = Format(lblNeto.Caption, "########0")
        v.cabeza.iva = Format(lblIVA.Caption, "########0")
        v.cabeza.impuestoHarina = Format(lblIHA.Caption, "########0")
        v.cabeza.impuestoila = ""
        v.cabeza.impuestoespecifico = ""
        v.cabeza.exento = ""
        v.cabeza.retencionparcial = ""
        v.cabeza.retenciontotal = ""
        v.cabeza.total = Format(lblTotal.Caption, "########0")
        v.cabeza.abono = "0"
        v.cabeza.Descuento = Replace(dato11.text, ".", ",")
        v.cabeza.contabilizado = ""
        v.cabeza.PAGADO = ""
        v.cabeza.comision = ""
        v.cabeza.fechapagocomision = ""
        v.cabeza.nula = "N"
        v.cabeza.boletadesde = desde
        v.cabeza.boletahasta = hasta
        
        v.detalle.loc = empresaActiva
        v.detalle.TIPO = dato1.text
        v.detalle.NUMERO = dato2.text
        v.detalle.linea = ""
        v.detalle.fecha = dato5.text & "-" & dato4.text & "-" & dato3.text
        v.detalle.rut = dato6.text & lbldv.Caption
        v.detalle.sucursal = dato7.text
        v.detalle.CODIGO = ""
        v.detalle.descripcion = ""
        v.detalle.cantidad = ""
        v.detalle.unidades = ""
        v.detalle.PRECIO = ""
        v.detalle.Descuento = ""
        v.detalle.total = ""
        v.detalle.vendedor = dato9.text
        v.detalle.pcosto = ""
        v.detalle.bodega = bodega
        v.detalle.vencimiento = lblAño.Caption & "-" & lblMes.Caption & "-" & lblDia.Caption
        v.detalle.numerofactura = ""
        
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
        lblDias.Caption = c.plazo
        If lblDias.Caption <> "" Then
            fechven = DateAdd("d", lblDias.Caption, fecha)
        Else
            fechven = fecha
        End If
        lblDia.Caption = Format(fechven, "dd")
        lblMes.Caption = Format(fechven, "mm")
        lblAño.Caption = Format(fechven, "yyyy")
    End Sub
    
    Private Sub structtoctrl()
        dato1.text = v.cabeza.TIPO
        dato2.text = v.cabeza.NUMERO
        fecha = Format(v.cabeza.fecha, "dd-mm-yyyy")
        dato3.text = Format(v.cabeza.fecha, "dd")
        dato4.text = Format(v.cabeza.fecha, "mm")
        dato5.text = Format(v.cabeza.fecha, "yyyy")
        lblDias.Caption = v.cabeza.plazo
        lblDia.Caption = Format(v.cabeza.vencimiento, "dd")
        lblMes.Caption = Format(v.cabeza.vencimiento, "mm")
        lblAño.Caption = Format(v.cabeza.vencimiento, "yyyy")
        dato6.text = v.cabeza.rut
        dato7.text = v.cabeza.sucursal
        dato9.text = v.cabeza.cajera
        dato8.text = v.cabeza.notapedido
        lblSub.Caption = v.cabeza.subtotal
        lblNeto.Caption = v.cabeza.neto
        lblIVA.Caption = v.cabeza.iva
        lblIHA.Caption = v.cabeza.impuestoHarina
        lblTotal.Caption = v.cabeza.total
        dato10.text = CDbl(detalle.Cell(1, 6).text)
        dato11.text = Replace(v.cabeza.Descuento, ",", ".")
        If dato1.text = "ZE" Then
            desde = String(10 - Len(v.cabeza.boletadesde), "0") & v.cabeza.boletadesde
            hasta = String(10 - Len(v.cabeza.boletahasta), "0") & v.cabeza.boletahasta
            lblDesde.Caption = "Desde: " & desde
            lblHasta.Caption = "Hasta: " & hasta
            lblDesde.Visible = True
            lblHasta.Visible = True
        End If
        
        Call dato6_KeyPress(13)
        Call DATO7_KeyPress(13)
        Call DATO9_KeyPress(13)
        Call DeshabilitarCajas(Me)
       'Detalle.RemoveItem Detalle.Rows - 1
        If detalle.Rows > 1 Then
            detalle.SelectionMode = cellSelectionByRow
        End If
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
    End Select
End Sub

'=============================================================================
'OPCIONES
'=============================================================================
    Private Sub opciones_FSCommand(ByVal command As String, ByVal args As String)
        Select Case command
            Case "modifica"
                'Call modificar
            Case "elimina"
'                If MsgBox("DESEA REALMENTE ELIMINAR Si / No", vbYesNo) = vbYes Then
                Call ELIMINAR
'                End If
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
        Call HabilitarCajas(Me, modifica)
        dato1.Enabled = False
        dato2.SetFocus
    End Sub
    
    Private Sub ELIMINAR()
        Select Case MsgBox("Si desea eliminar el documento presione SI" & vbCrLf & "Si desea anular el documento presione NO" & vbCrLf & "Presione CANCELAR para volver", vbYesNoCancel, "Alerta")
            Case vbYes
                Call ctrltostruct(False)
                Call eliminarVenta(v, detalle)
                Call eliminarPagos(p.tipodocumento, p.numeroDocumento, Format(p.fecha, "yyyy-mm-dd"), v.cabeza.caja)
                Call retorno
                Call HabilitarCajas(Me, modifica)
                dato1.SetFocus
            Case vbNo
                'If lblNulo.Caption <> "" Then
                    v.detalle.fecha = dato5.text & "-" & dato4.text & "-" & dato3.text
                    Call anularDocumento(dato1.text, dato2.text, detalle, v.detalle)
                    Call eliminarPagos(p.tipodocumento, p.numeroDocumento, Format(p.fecha, "yyyy-mm-dd"), v.cabeza.caja)
                'End If
                Call retorno
                Call HabilitarCajas(Me, modifica)
                dato1.SetFocus
        End Select
    End Sub

    Private Sub imprimir()
        If nula = False Then
            Call imprimeFactura(dato2.text, impresion, data)
'            Select Case dato1.text
'                Case "BV"
'                    Call imprimeBoletaMatPun(dato2.text, impresion, data)
'                Case "FV"
'                    Call imprimeFactura(dato2.text, impresion, data)
'                Case "ZE"
'            End Select
        Else
            Call MsgBox("Documento nulo no imprimible", vbOKOnly, "Mensaje")
        End If
    End Sub
    
    Private Sub retorno()
        Call LimpiarCajas(Me)
        Call LimpiarLabels(Me)
        'Call cargaGrilla(1, 6)
        detalle.Rows = 1
        modifica = False
        nula = False
        imprimio = False
        lectura = False
        dato10.text = "0"
        dato11.text = "0"
        lblSub.Caption = "0"
        lblNeto.Caption = "0"
        lblIVA.Caption = "0"
        lblIHA.Caption = "0"
        lblTotal.Caption = "0"
        Call HabilitarCajas(Me, modifica)
        dato1.SetFocus
        lblDesde.Visible = False
        lblHasta.Visible = False
    End Sub
        
    Private Sub anterior()
        If leerVenta(v, dato1.text, dato2.text, "<", data, detalle) = True Then
            structtoctrl
        End If
    End Sub
    
    Private Sub siguiente()
        If leerVenta(v, dato1.text, dato2.text, ">", data, detalle) = True Then
            structtoctrl
        End If
    End Sub
'=============================================================================
'OPCIONES
'=============================================================================

    Private Function leerPrecioEspecial(ByVal CODIGO As String) As String
        
        Dim CAMPOS(10, 3) As String
        Dim op As Integer
        Set sql = New sqlventas.sqlventa
        CAMPOS(0, 0) = "precioespecial"
        CAMPOS(1, 0) = ""
        
        CAMPOS(0, 2) = "sv_maestroclientes_especiales"
        
        condicion = "rut = '" & dato6.text & lbldv.Caption & "' AND sucursal = '" & dato7.text & "' AND codigo = '" & CODIGO & "'"
        op = 5
        sql.response = CAMPOS
        Set sql.conexion = ventas
        Call sql.sqlventas(op, condicion)
        If sql.Status = 0 Then
            leerPrecioEspecial = sql.response(0, 3)
        Else
            leerPrecioEspecial = "0"
        End If
    End Function

    Private Sub ctrltostructCobranza()
        D.TIPO = dato1.text
        D.NUMERO = dato2.text
        D.fEmision = dato5.text & "-" & dato4.text & "-" & dato3.text
        D.fVencimiento = lblAño.Caption & "-" & lblMes.Caption & "-" & lblDia.Caption
        D.rut = dato6.text & lbldv.Caption
        D.sucursal = dato7.text
        D.cajera = "0000"
        D.MONTO = Format(lblTotal.Caption, "########0")
        D.abono = "0"
        D.obs = "GENERADO POR PANTALLA DE NOTAS DE CREDITO"
        Call grabarDocManual(D, modifica)
    End Sub


