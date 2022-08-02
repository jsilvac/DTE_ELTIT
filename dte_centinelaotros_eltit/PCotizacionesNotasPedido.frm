VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash9b.ocx"
Object = "{4777436C-EB8C-4596-98A8-EBCF98683969}#1.0#0"; "FlexCell.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{ADBBAED6-B16F-40DC-80DF-B44910CBA76C}#1.0#0"; "Frame-Xp.ocx"
Begin VB.Form PCotizacionesNotasPedido 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pantalla de Ventas"
   ClientHeight    =   9660
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13710
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9660
   ScaleWidth      =   13710
   Begin FlexCell.Grid impresion 
      Height          =   495
      Left            =   4680
      TabIndex        =   55
      Top             =   7560
      Visible         =   0   'False
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   873
      Cols            =   5
      DefaultFontSize =   8.25
      Rows            =   30
   End
   Begin XPFrame.FrameXp frmDatos 
      Height          =   2655
      Left            =   120
      TabIndex        =   16
      Top             =   120
      Width           =   13455
      _ExtentX        =   23733
      _ExtentY        =   4683
      BackColor       =   16744576
      Caption         =   "Datos de la Cotizacion"
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
      Begin VB.ComboBox cmbTipo 
         Appearance      =   0  'Flat
         BackColor       =   &H00E1FFFD&
         Height          =   315
         ItemData        =   "PCotizacionesNotasPedido.frx":0000
         Left            =   1920
         List            =   "PCotizacionesNotasPedido.frx":000A
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   480
         Width           =   1815
      End
      Begin VB.TextBox dato6 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         MaxLength       =   9
         TabIndex        =   5
         Tag             =   "proveedor"
         Top             =   840
         Width           =   1455
      End
      Begin VB.TextBox dato5 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Left            =   8520
         MaxLength       =   10
         TabIndex        =   1
         Tag             =   "proveedor"
         Top             =   480
         Width           =   1455
      End
      Begin VB.TextBox dato1 
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
         Left            =   1920
         MaxLength       =   2
         TabIndex        =   10
         Tag             =   "proveedor"
         Top             =   480
         Width           =   375
      End
      Begin VB.TextBox dato9 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Left            =   5880
         MaxLength       =   1
         TabIndex        =   6
         Tag             =   "proveedor"
         Top             =   840
         Width           =   615
      End
      Begin VB.Label lbl10 
         Appearance      =   0  'Flat
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Vencimiento"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   4920
         TabIndex        =   41
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
         TabIndex        =   40
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
         TabIndex        =   39
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
         TabIndex        =   38
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
         TabIndex        =   37
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
         TabIndex        =   36
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
         TabIndex        =   35
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
         TabIndex        =   34
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
         TabIndex        =   33
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
         TabIndex        =   32
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
         TabIndex        =   31
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
         TabIndex        =   30
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
         TabIndex        =   29
         Top             =   840
         Width           =   375
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
         Left            =   3840
         TabIndex        =   28
         Top             =   480
         Width           =   2655
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
         TabIndex        =   27
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
         TabIndex        =   26
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
         TabIndex        =   25
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
         TabIndex        =   24
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
         TabIndex        =   23
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
         TabIndex        =   22
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
         TabIndex        =   21
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
         TabIndex        =   20
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
         TabIndex        =   19
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
         TabIndex        =   18
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
         TabIndex        =   17
         Top             =   2280
         Visible         =   0   'False
         Width           =   375
      End
   End
   Begin XPFrame.FrameXp frmDetalle 
      Height          =   4575
      Left            =   120
      TabIndex        =   15
      Top             =   2880
      Width           =   13455
      _ExtentX        =   23733
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
         Height          =   4215
         Left            =   0
         TabIndex        =   9
         Top             =   315
         Width           =   13470
         _ExtentX        =   23760
         _ExtentY        =   7435
         Cols            =   5
         DefaultFontSize =   8.25
         Rows            =   30
      End
   End
   Begin VB.TextBox pivote 
      Height          =   285
      Left            =   120
      MaxLength       =   13
      TabIndex        =   14
      Top             =   2760
      Visible         =   0   'False
      Width           =   735
   End
   Begin MSAdodcLib.Adodc data 
      Height          =   330
      Left            =   120
      Top             =   7680
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
      Height          =   1695
      Left            =   8160
      TabIndex        =   42
      Top             =   7860
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   2990
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
         MaxLength       =   5
         TabIndex        =   12
         Tag             =   "proveedor"
         Text            =   "0"
         Top             =   1230
         Width           =   1215
      End
      Begin VB.TextBox dato10 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         MaxLength       =   5
         TabIndex        =   11
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
         TabIndex        =   52
         Top             =   1200
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
         TabIndex        =   50
         Top             =   1200
         Width           =   1215
      End
      Begin VB.Label lblIHA 
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
         Left            =   4080
         TabIndex        =   53
         Top             =   1200
         Visible         =   0   'False
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
         TabIndex        =   54
         Top             =   1200
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label lblIVA 
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
         Left            =   4080
         TabIndex        =   51
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
         TabIndex        =   49
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
         TabIndex        =   48
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
         TabIndex        =   47
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
         TabIndex        =   46
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
         TabIndex        =   45
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label lblNeto 
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
         Left            =   4080
         TabIndex        =   44
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label lblSub 
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
         Left            =   1440
         TabIndex        =   43
         Top             =   480
         Width           =   1215
      End
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Para Eliminar una linea presione la tecla -"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8100
      TabIndex        =   57
      Top             =   7560
      Width           =   5415
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
      TabIndex        =   56
      Top             =   7560
      Width           =   7935
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash opciones 
      Height          =   1455
      Left            =   120
      TabIndex        =   13
      Top             =   8040
      Width           =   7695
      _cx             =   13573
      _cy             =   2566
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
Attribute VB_Name = "PCotizacionesNotasPedido"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Private c As Cliente
    Private v As venta
    Private p As pagos
    Private formatoGrilla(10, 10) As String
    Private modifica As Boolean
    Private vacio As Boolean
    Private fila As Long
    Private columna As Long
    Private nula As Boolean
    Private lectura As Boolean
    Public imprimio As Boolean

Private Sub cmbTipo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        dato1.text = Left(cmbTipo.List(cmbTipo.ListIndex), 2)
        dato2.text = leerUltimoFolio(dato1.text)
        Call dato1_KeyPress(13)
        Call dato2_KeyPress(13)
    End If
End Sub

'============================================================
'MANEJO DE LOS CONTOLES
'============================================================
    '========================================================
    'GotFocus
    '========================================================
    Private Sub dato1_GotFocus()
        Call selecciona(dato1)
    End Sub
    
    Private Sub dato2_GotFocus()
        If dato1.text = "" Then
            cmbTipo_KeyPress (13)
        End If
        Call selecciona(dato2)
    End Sub
    
    Private Sub dato3_GotFocus()
        Call selecciona(dato3)
    End Sub
    
    Private Sub dato4_GotFocus()
        Call selecciona(dato4)
    End Sub
    
    Private Sub dato5_GotFocus()
        Call selecciona(dato5)
    End Sub
    
    Private Sub dato6_GotFocus()
        Call selecciona(dato6)
        Principal.barraEstado.Panels(2).text = "F2: Ayuda Cliente"
    End Sub
    
    Private Sub dato7_GotFocus()
        Call selecciona(dato7)
        Principal.barraEstado.Panels(2).text = "F2: Ayuda Sucursal"
    End Sub
    
    Private Sub dato8_GotFocus()
        Call selecciona(dato8)
        Principal.barraEstado.Panels(2).text = "F2: Ayuda Nota de Pedido"
    End Sub
    
    Private Sub dato9_GotFocus()
        Call selecciona(dato9)
        Principal.barraEstado.Panels(2).text = "F2: Ayuda Vendedor"
    End Sub
    
    Private Sub dato10_GotFocus()
        Call selecciona(dato10)
    End Sub
    
    Private Sub dato11_GotFocus()
        Call selecciona(dato11)
    End Sub
    '========================================================
    'GotFocus
    '========================================================
    '========================================================
    'KeyDown
    '========================================================
    Private Sub dato1_KeyDown(KeyCode As Integer, Shift As Integer)
        Call Flechas(KeyCode, dato1)
    End Sub
    
    Private Sub dato2_KeyDown(KeyCode As Integer, Shift As Integer)
        Call Flechas(KeyCode, cmbTipo)
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
            Call ayudaCliente(dato6)
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
        KeyAscii = esNumero(KeyAscii)
        If KeyAscii = 13 Then
            lblDocumento.Caption = leerNombreDocumento(dato1.text)
            If lblDocumento.Caption <> "" Then
                SendKeys "{Tab}"
            End If
        End If
    End Sub
    
    Private Sub dato2_KeyPress(KeyAscii As Integer)
        KeyAscii = esNumero(KeyAscii)
        If KeyAscii = 13 Then
            dato2.text = ceros(dato2)
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
                If detalle.Rows <= 1 Then
                    detalle.AddItem vbTab & vbTab & "1", True
                End If
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
            lblDV.Caption = rut(dato6.text)
            rut_cliente = dato6.text
            SendKeys "{Tab}"
        End If
    End Sub
    
    Private Sub dato7_KeyPress(KeyAscii As Integer)
        KeyAscii = esNumero(KeyAscii)
        If KeyAscii = 13 Then
            dato7.text = ceros(dato7)
            If leerCliente(c, dato6.text & lblDV.Caption, dato7.text, "=") = True Then
                structtoctrlCliente
                SendKeys "{Tab}"
            Else
                If MsgBox("El rut ingresado no se encuentra. ¿Desea crearlo?", vbYesNo, "Mensaje") = vbYes Then
                    Load MClientes
                    MClientes.dato1.text = dato6.text
                    MClientes.lblDV.Caption = lblDV.Caption
                    MClientes.dato2.text = dato7.text
                    MClientes.Show
                End If
            End If
        End If
    End Sub
    
    Private Sub dato8_KeyPress(KeyAscii As Integer)
        KeyAscii = esNumero(KeyAscii)
        If KeyAscii = 13 Then
            dato8.text = ceros(dato8)
            SendKeys "{Tab}"
        End If
    End Sub
    
    Private Sub dato9_KeyPress(KeyAscii As Integer)
        KeyAscii = esNumero(KeyAscii)
        If KeyAscii = 13 Then
            dato9.text = ceros(dato9)
            lblDVV.Caption = rut(dato9.text)
            lblVendedor.Caption = leerNombreVendedor(dato9.text)
            If lblVendedor.Caption <> "" Then
                detalle.Enabled = True
                If detalle.Rows > 1 And lectura = False Then
                    detalle.Cell(1, 1).SetFocus
                End If
            End If
        End If
    End Sub
    
    Private Sub dato10_KeyPress(KeyAscii As Integer)
        KeyAscii = esNumero(KeyAscii)
        If KeyAscii = 13 And dato10.text <> "" Then
            SendKeys "{Tab}"
        End If
        If dato10.text = "" Then
            dato10.text = "0"
        End If
    End Sub
    
    Private Sub dato11_KeyPress(KeyAscii As Integer)
        KeyAscii = esNumero(KeyAscii)
        If KeyAscii = 13 And dato11.text <> "" Then
            Call ctrltostruct(True)
            Call imprimir
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
    Private Sub dato10_KeyUp(KeyCode As Integer, Shift As Integer)
        Dim desc As Double
        Dim total As Double
        If lblNeto.Caption <> "" And lblIVA.Caption <> "" And lblIHA.Caption <> "" Then
            total = Round(CDbl(lblNeto.Caption) + CDbl(lblIVA.Caption) + CDbl(lblIHA.Caption), 0)
        End If
        If dato10.text <> "" And dato10.text <> "0" Then
            desc = total * CDbl(dato10.text) / 100
            dato11.text = desc
            lblTotal.Caption = Round(total - CDbl(dato11.text), 0)
        Else
            dato10.text = "0"
            dato11.text = "0"
            lblTotal.Caption = Round(total - CDbl(dato11.text), 0)
        End If
    End Sub

    Private Sub dato11_KeyUp(KeyCode As Integer, Shift As Integer)
        Dim desc As Double
        Dim neto As Double
        Dim total As Double
        If lblNeto.Caption <> "" And lblIVA.Caption <> "" And lblIHA.Caption <> "" Then
            total = Round(CDbl(lblNeto.Caption) + CDbl(lblIVA.Caption) + CDbl(lblIHA.Caption), 0)
        End If
        If dato11.text <> "" And dato11.text <> "0" Then
            desc = CDbl(dato11.text)
            desc = desc * 100 / total
            dato10.text = Round(desc, 5)
            lblTotal.Caption = Round(total - CDbl(dato11.text), 0)
        Else
            dato10.text = "0"
            dato11.text = "0"
            lblTotal.Caption = Round(total - CDbl(dato11.text), 0)
        End If
    End Sub
    '========================================================
    'KeyUp
    '========================================================

    '========================================================
    'LostFocus
    '========================================================
    Private Sub dato2_LostFocus()
        dato2.text = ceros(dato2)
    End Sub

    Private Sub dato6_LostFocus()
        dato6.text = ceros(dato6)
        Call limpiaBarra(2)
    End Sub
    
    Private Sub dato7_LostFocus()
        Call limpiaBarra(2)
    End Sub
    
    Private Sub dato8_LostFocus()
        Call limpiaBarra(2)
    End Sub
    
    Private Sub dato9_LostFocus()
        dato9.text = ceros(dato9)
        Call limpiaBarra(2)
    End Sub
    '========================================================
    'LostFocus
    '========================================================
'============================================================
'MANEJO DE LOS CONTOLES
'============================================================

    Private Sub Form_Activate()
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
        modifica = False
        nula = False
        imprimio = False
        lectura = False
        Call Centrar(Me)
        Call cargaGrilla(1, 6)
    End Sub
    
    Private Sub Form_Unload(Cancel As Integer)
        Principal.barraEstado.Panels(1).text = UCase(Principal.Caption)
        Call limpiaBarra(2)
    End Sub

'****************************************************************************
'Formato de la Grilla
'****************************************************************************
    Private Sub cargaGrilla(ByVal row As Integer, ByVal col As Integer)
        Dim i As Integer
        Rem DATOS DE LA COLUMNA
        formatoGrilla(1, 0) = "LN"
        formatoGrilla(1, 1) = "CODIGO"
        formatoGrilla(1, 2) = "DESCRIPCION"
        formatoGrilla(1, 3) = "CANTIDAD"
        formatoGrilla(1, 4) = "PRECIO"
        formatoGrilla(1, 5) = "TOTAL"
        
        Rem LARGO DE LOS DATOS
        formatoGrilla(2, 1) = "15"
        formatoGrilla(2, 2) = "40"
        formatoGrilla(2, 3) = "10"
        formatoGrilla(2, 4) = "12"
        formatoGrilla(2, 5) = "13"
        
        Rem TIPO DE DATOS C=CEROS,S=STRING,N=NUMERICO,D=DATE,
        formatoGrilla(3, 1) = "N"
        formatoGrilla(3, 2) = "S"
        formatoGrilla(3, 3) = "N"
        formatoGrilla(3, 4) = "N"
        formatoGrilla(3, 5) = "N"
        
        Rem FORMATO GRILLA
        formatoGrilla(4, 1) = "0000000000000"
        formatoGrilla(4, 2) = ""
        formatoGrilla(4, 3) = "#,###,##0"
        formatoGrilla(4, 4) = "$ ###,###,##0.00"
        formatoGrilla(4, 5) = "$ ###,###,##0.00"
        
        Rem LOCCKED
        'FormatoGrilla(5, 0) = "TRUE"
        formatoGrilla(5, 1) = "FALSE"
        formatoGrilla(5, 2) = "TRUE"
        formatoGrilla(5, 3) = "FALSE"
        formatoGrilla(5, 4) = "FALSE"
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
            
        detalle.Cols = col + 1
        detalle.Rows = row
        detalle.AllowUserResizing = False
        detalle.DisplayFocusRect = False
        detalle.ExtendLastCol = True
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
        
        detalle.Column(col).Width = 0
        detalle.Cell(0, col).text = "PCOSTO"
        detalle.Column(col).Locked = True
        
        detalle.Cell(0, 0).text = formatoGrilla(1, 0)
        For i = 1 To col - 1
            detalle.Cell(0, i).text = formatoGrilla(1, i)
            detalle.Column(i).Width = Val(formatoGrilla(2, i)) * (detalle.Cell(0, i).Font.Size + 1.25)
            detalle.Column(i).MaxLength = Val(formatoGrilla(2, i))
            detalle.Column(i).FormatString = formatoGrilla(4, i)
            detalle.Column(i).Locked = formatoGrilla(5, i)
            If formatoGrilla(3, i) = "N" Then
                detalle.Column(i).Alignment = cellRightCenter
                detalle.Column(i).Mask = cellNumeric
            Else
                detalle.Column(i).Alignment = cellLeftCenter
                detalle.Column(i).Mask = cellUpper
            End If
        Next i
        detalle.Range(0, 0, 0, detalle.Cols - 1).Alignment = cellCenterCenter
        detalle.AddItem vbTab & vbTab & "1", True
        detalle.Enabled = True
    End Sub
    Private Sub detalle_GotFocus()
        If detalle.ActiveCell.col = 1 Then
            Principal.barraEstado.Panels(2).text = "F2 Ayuda - Producto"
        Else
            Principal.barraEstado.Panels(2).text = ""
        End If
    End Sub
    
'    Private Sub Grid1_KeyPress(KeyAscii As Integer)
'        KeyAscii = Asc(UCase(Chr(KeyAscii)))
'
'        'If Grid1.ActiveCell.Col = 1 And Chr(KeyAscii) = "*" And SALDOPE = neto Then grabafactura
'
'        If FormatoGrilla(3, Grid1.ActiveCell.col) = "N" Then snum = 1: KeyAscii = esNumeroDecimal(Grid1.ActiveCell.text, KeyAscii)
'        If FormatoGrilla(3, Grid1.ActiveCell.col) = "C" Then snum = 1: KeyAscii = esNumero(KeyAscii, "N")
'    End Sub
    
    Private Sub detalle_KeyDown(KeyCode As Integer, ByVal Shift As Integer)
        fila = detalle.ActiveCell.row
        columna = detalle.ActiveCell.col
        If detalle.ActiveCell.col = 1 And KeyCode = vbKeyF2 Then Call ayudaProducto(detalle, pivote): detalle.Cell(fila, columna).SetFocus
        Select Case KeyCode
            Case 13, 37, 38, 39, 40
                If detalle.ActiveCell.text <> "" Then
                    vacio = False
                Else
                    vacio = True
                End If
            Case 106
                SendKeys "{Tab}"
            Case 109
                If fila < detalle.Rows - 1 Then
                    detalle.RemoveItem (fila)
                End If
        End Select
        
    End Sub
    
    Private Sub detalle_LeaveCell(ByVal row As Long, ByVal col As Long, NewRow As Long, NewCol As Long, Cancel As Boolean)
        Dim i As Integer
        Dim linea As String
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
            If col = detalle.Cols - 2 And NewCol = detalle.Cols - 1 Then
                If row = detalle.Rows - 1 Then
                    detalle.AddItem vbTab & vbTab & "1", True
                    NewRow = detalle.Rows - 1
                    NewCol = 1
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
            If col = 1 And NewCol <> col Then
                pivote.text = detalle.Cell(row, 1).text
                pivote.text = ceros(pivote)
                detalle.Cell(row, 1).text = pivote.text
                detalle.Cell(row, detalle.Cols - 1).text = leerCostoProducto(detalle.Cell(row, 1).text)
                detalle.Cell(row, 2).text = leerNombreProducto(detalle.Cell(row, 1).text)
                detalle.Cell(row, 4).text = leerPrecioProducto(detalle.Cell(row, 1).text, "1")
                If detalle.Cell(row, 3).text <> "" And detalle.Cell(row, 4).text <> "" Then
                    detalle.Cell(row, 5).text = detalle.Cell(row, 3).text * detalle.Cell(row, 4).text
                Else
                    detalle.Cell(row, 5).text = "0"
                End If
            End If
            If NewRow > row Then
                For i = 1 To detalle.Cols - 1
                    If detalle.Cell(row, i).text = "" Then
                        NewRow = row
                        NewCol = i
                        Exit For
                    End If
                Next i
                For i = 1 To detalle.Cols - 1
                    If detalle.Cell(NewRow, i).text = "" Then
                        NewCol = i
                        Exit For
                    End If
                Next i
            End If
            If row > 0 Then
                If detalle.Cell(row, 3).text <> "" And detalle.Cell(row, 4).text <> "" Then
                    detalle.Cell(row, 5).text = detalle.Cell(row, 3).text * detalle.Cell(row, 4).text
                End If
            End If
            Call sumaGrilla(detalle)
        End If
    End Sub
    
    Private Sub sumaGrilla(ByRef lista As Grid)
        Dim i As Integer
        Dim suma As Double
        Dim codigo As String
        Dim totIHA As Double
        Dim totIVA As Double
        Dim totPro As Double
        suma = 0
        totIVA = 0
        totIHA = 0
        For i = 1 To lista.Rows - 1
            If lista.Cell(i, 1).text <> "$ 0" And lista.Cell(i, 1).text <> "" Then
                codigo = lista.Cell(i, 1).text
                Select Case leerImpuestoProducto(codigo)
                    Case "IHA"
                        totPro = CDbl(lista.Cell(i, 5).text)
                        totIVA = totIVA + totPro - totPro / (1 + iva / 100)
                        totIHA = totIHA + totPro - totPro / (1 + iha / 100)
                    Case "IVA"
                        totPro = CDbl(lista.Cell(i, 5).text)
                        totIVA = totIVA + totPro - totPro / (1 + iva / 100)
                End Select
                suma = suma + CDbl(lista.Cell(i, 5).text)
            End If
        Next i
        lblSub.Caption = suma
        lblNeto.Caption = Round(suma - totIVA, 0)
        lblIVA.Caption = Round(totIVA, 0)
        If dato1.text = "FV" Then
            lblIHA.Caption = Round(totIHA, 0)
        Else
            lblIHA.Caption = "0"
        End If
        lblTotal.Caption = Round(CDbl(lblNeto.Caption) + CDbl(lblIVA.Caption) + CDbl(lblIHA.Caption), 0)
    End Sub

'****************************************************************************
'Formato de la Grilla
'****************************************************************************

'=============================================================================
'PASA LOS DATOS DE LOS CONTROLES A LA ESTRUCTURA
'=============================================================================
    Private Sub ctrltostruct(ByVal graba As Boolean)
        v.cabeza.loc = empresaActiva
        v.cabeza.tipo = dato1.text
        v.cabeza.numero = dato2.text
        v.cabeza.fecha = dato5.text & "-" & dato4.text & "-" & dato3.text
        v.cabeza.rut = dato6.text & lblDV.Caption
        v.cabeza.sucursal = dato7.text
        
        v.detalle.loc = v.cabeza.loc
        v.detalle.tipo = dato1.text
        v.detalle.numero = dato2.text
        v.detalle.fecha = v.cabeza.fecha
        v.detalle.rut = v.cabeza.rut
        
        'v.impuestos.loc = v.cabeza.loc
        'v.impuestos.tipo = dato1.text
        'v.impuestos.numero = dato2.text
        'v.impuestos.vencimiento = lblAño.Caption & "-" & lblMes.Caption & "-" & lblDia.Caption
        'v.impuestos.notapedido = dato8.text
        'v.impuestos.cajera = dato9.text & lblDVV.Caption
        'v.impuestos.subtotal = lblSub.Caption
        'v.impuestos.neto = lblNeto.Caption
        'v.impuestos.iva = lblIVA.Caption
        'v.impuestos.iha = lblIHA.Caption
        'v.impuestos.descuentoporcentaje = Replace(dato10.text, ",", ".")
        'v.impuestos.descuentopesos = dato11.text
        'v.impuestos.total = lblTotal.Caption
        
        p.tipodocumento = dato1.text
        p.numerodocumento = dato2.text
        p.fecha = v.cabeza.fecha
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
        Dim fecha As String
        lblRazon.Caption = c.nombre
        lblDireccion.Caption = c.direccion
        lblComuna.Caption = c.comuna
        lblCiudad.Caption = c.ciudad
        lblDias.Caption = c.plazo
        If lblDias.Caption <> "" Then
            fecha = DateAdd("d", lblDias.Caption, fechasistema)
        Else
            fecha = fechasistema
        End If
        lblDia.Caption = Format(fecha, "dd")
        lblMes.Caption = Format(fecha, "mm")
        lblAño.Caption = Format(fecha, "yyyy")
    End Sub
    
    Private Sub structtoctrl()
        Dim fecha As String
        dato1.text = v.cabeza.tipo
        dato2.text = v.cabeza.numero
        fecha = v.cabeza.fecha
        dato3.text = Format(fecha, "dd")
        dato4.text = Format(fecha, "mm")
        dato5.text = Format(fecha, "yyyy")
        dato6.text = v.cabeza.rut
        dato7.text = v.cabeza.sucursal
        'dato8.text = v.impuestos.notapedido
        'dato9.text = v.impuestos.cajera
        'dato10.text = v.impuestos.descuentoporcentaje
        'dato11.text = v.impuestos.descuentopesos
        'fecha = v.impuestos.vencimiento
        lblDia.Caption = Format(fecha, "dd")
        lblMes.Caption = Format(fecha, "mm")
        lblAño.Caption = Format(fecha, "yyyy")
        'lblDias.Caption = DateDiff("d", v.impuestos.vencimiento, v.cabeza.fecha)
        'lblSub.Caption = v.impuestos.subtotal
        'lblNeto.Caption = v.impuestos.neto
        'lblIVA.Caption = v.impuestos.iva
        'lblTotal.Caption = v.impuestos.total
        
        Call dato6_KeyPress(13)
        Call dato7_KeyPress(13)
        Call dato9_KeyPress(13)
        Call DeshabilitarCajas(Me)
       'Detalle.RemoveItem Detalle.Rows - 1
        detalle.Range(1, 1, detalle.Rows - 1, detalle.Cols - 1).Locked = True
    End Sub
'=============================================================================
'PASA LOS DATOS DE LA ESTRUCTURA A LOS CONTROLES
'=============================================================================



'=============================================================================
'OPCIONES
'=============================================================================
    Private Sub opciones_FSCommand(ByVal command As String, ByVal args As String)
        Select Case command
            Case "modifica"
                'Call modificar
            Case "elimina"
                Call eliminar
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
    
    Private Sub eliminar()
        Call ctrltostruct(False)
        Call eliminarVenta(v, detalle)
        Call eliminarPagos(p.tipodocumento, p.numerodocumento)
        Call retorno
        Call HabilitarCajas(Me, modifica)
        cmbTipo.SetFocus
    End Sub

    Private Sub imprimir()
        Call imprimeCotizacionNotaPedido(dato2.text, Impresion, data, dato1.text)
    End Sub
    
    Private Sub retorno()
        Call LimpiarCajas(PCotizacionesNotasPedido)
        Call LimpiarLabels(PCotizacionesNotasPedido)
        Call cargaGrilla(1, 6)
        modifica = False
        Call DeshabilitarCajas(Me)
        dato10.text = "0"
        dato11.text = "0"
        cmbTipo.SetFocus
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

