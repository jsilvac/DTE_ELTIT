VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash9d.ocx"
Object = "{4777436C-EB8C-4596-98A8-EBCF98683969}#1.0#0"; "FlexCell.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{ADBBAED6-B16F-40DC-80DF-B44910CBA76C}#1.0#0"; "Frame-Xp.ocx"
Begin VB.Form PVentas 
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
   MinButton       =   0   'False
   ScaleHeight     =   9660
   ScaleWidth      =   13710
   StartUpPosition =   2  'CenterScreen
   Begin XPFrame.FrameXp FRMPREVENTA 
      Height          =   1815
      Left            =   1440
      TabIndex        =   67
      Top             =   3600
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
         TabIndex        =   69
         Top             =   720
         Width           =   2535
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
         Left            =   720
         TabIndex        =   68
         Top             =   720
         Width           =   6015
      End
   End
   Begin XPFrame.FrameXp frmTipo 
      Height          =   1635
      Left            =   2580
      TabIndex        =   61
      Top             =   600
      Visible         =   0   'False
      Width           =   2715
      _ExtentX        =   4789
      _ExtentY        =   2884
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
      Begin VB.Label lbl25 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " 3 - Zeta"
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
         Top             =   1200
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
         TabIndex        =   63
         Top             =   840
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
         TabIndex        =   62
         Top             =   480
         Width           =   2475
      End
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
         Left            =   1920
         MaxLength       =   2
         TabIndex        =   7
         Tag             =   "proveedor"
         Top             =   1920
         Width           =   615
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
         Left            =   11520
         MaxLength       =   12
         TabIndex        =   8
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
         Left            =   1920
         MaxLength       =   2
         TabIndex        =   9
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
         TabIndex        =   28
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
         Left            =   2580
         TabIndex        =   23
         Top             =   1920
         Width           =   555
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
         Width           =   4785
      End
      Begin VB.Label lbl22 
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
         Height          =   4095
         Left            =   0
         TabIndex        =   10
         Top             =   360
         Width           =   13470
         _ExtentX        =   23760
         _ExtentY        =   7223
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
      Left            =   8460
      TabIndex        =   42
      Top             =   7620
      Width           =   5115
      _ExtentX        =   9022
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
         Locked          =   -1  'True
         MaxLength       =   9
         TabIndex        =   12
         Tag             =   "proveedor"
         Text            =   "0"
         Top             =   1200
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
         Locked          =   -1  'True
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
         Left            =   3780
         TabIndex        =   52
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
         TabIndex        =   50
         Top             =   1560
         Width           =   975
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
         Left            =   3780
         TabIndex        =   53
         Top             =   1200
         Width           =   1215
      End
      Begin VB.Label lbl18 
         Appearance      =   0  'Flat
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "EXENTO"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   2760
         TabIndex        =   54
         Top             =   1200
         Width           =   975
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
         Left            =   3780
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
         TabIndex        =   48
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
         Left            =   3780
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
         TabIndex        =   43
         Top             =   480
         Width           =   1215
      End
   End
   Begin FlexCell.Grid impresion 
      Height          =   495
      Left            =   1320
      TabIndex        =   55
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
      TabIndex        =   60
      Top             =   0
      Width           =   555
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
      Left            =   6360
      TabIndex        =   66
      Top             =   9180
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
      Left            =   6360
      TabIndex        =   65
      Top             =   8700
      Visible         =   0   'False
      Width           =   1935
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
      TabIndex        =   59
      Top             =   7860
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
      TabIndex        =   58
      Top             =   7860
      Visible         =   0   'False
      Width           =   2475
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
      Left            =   6360
      TabIndex        =   57
      Top             =   8220
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
      Left            =   90
      TabIndex        =   56
      Top             =   7200
      Width           =   7935
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash opciones 
      Height          =   1455
      Left            =   300
      TabIndex        =   13
      Top             =   8160
      Visible         =   0   'False
      Width           =   6375
      _cx             =   11245
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
Attribute VB_Name = "PVentas"
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
    Public desde As String
    Public hasta As String
    Public dire As Integer
    Private tipoprecio As String
    Private fecha As String
    Private cupo As Double
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
        If dato5.text + "-" + dato4.text + "-" + dato3.text <> fechasistema Then
        MsgBox ("imposible digitar facturas con fecha anterior")
        
        dato3.SetFocus
        
        End If
        
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
        Principal.barraEstado.Panels(2).text = "F2: Ayuda Nota de Pedido"
    End Sub
    
    Private Sub dato10_GotFocus()
        Call VerificarCajas(Me, dato10)
        Call selecciona(dato10)
        Principal.barraEstado.Panels(2).text = "F2: Ayuda Vendedor"
    End Sub
    
    Private Sub dato11_GotFocus()
        Static cont As Integer
        Call VerificarCajas(Me, dato11)
        Call selecciona(dato11)
        cont = cont + 1
        If descuento.Visible = False Then
            If cont = 1 Then
                Load descuento
                descuento.monto = CDbl(lblSub.Caption)
                descuento.Show vbModal
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
                dato1.text = "ZE"
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
            Call ayudaCliente(dato6, dato7, lblDV)
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
                Select Case dato1.text
                    Case "GM"
                        Unload Me
                        guiaMolienda.Show vbModal
                    Case Else
                        dato2.text = leerUltimoFolio(dato1.text)
                        dato2.text = ceros(dato2)
                End Select
                SendKeys "{Tab}"
            End If
            tipo_doc = dato1.text
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
                lbl30.Visible = True
                lbl20.Visible = False
                lbl26.Visible = False
                opciones.Visible = True
                opciones.SetFocus
                        
            Else
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
                SendKeys "{Tab}"
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
                If dato1 <> "ZE" Then
                    dato6.SetFocus
                Else
                    DesdeHasta.Show vbModal
                    dato6.text = "999999999"
                    lblDV.Caption = "6"
                    dato7.text = "0"
                    dato9.text = "0000000000"
                    dato10.text = "0000000000"
                    Call dato7_KeyPress(13)
                    'lblVendedor.Caption = leerNombreVendedor(dato10.text & lblDVV.Caption)
                    SendKeys "{Tab}"
                    SendKeys "{Tab}"
                    SendKeys "{Tab}"
                    SendKeys "{Tab}"
                    SendKeys "{Tab}"
                End If
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
                lblDV.Caption = "6"
                dato7.text = "0"
                dato9.text = "0000000000"
                dato10.text = "0000000000"
                Call dato7_KeyPress(13)
                'lblVendedor.Caption = leerNombreVendedor(dato10.text & lblDVV.Caption)
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
            lblDV.Caption = rut(dato6.text)
            rut_cliente = dato6.text
            SendKeys "{Tab}"
        End If
    End Sub
    
    Private Sub dato7_KeyPress(KeyAscii As Integer)
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If KeyAscii = 13 Then
            dato7.text = ceros(dato7)
            If leerCliente(c, dato6.text & lblDV.Caption, dato7.text, "=") = True Then
                structtoctrlCliente
                rut_cliente = dato6.text & lblDV.Caption
                sucursal_cliente = dato7.text
                If lectura = False Then
                    If autorizado = False Then
                        If verificarCupoCliente(dato6.text & lblDV.Caption, dato7.text) = False Then
                            Call enviarInformacion(rut_cliente, sucursal_cliente, dato1.text, dato2.text, "0", "CUPO INSUFICINTE")
                            'Call mensaje.mostrarMensaje("Información Crédito Cliente", "El cliente " & dato6.text & "-" & lblDV.Caption & " no posee cupo suficiente para realizar la compra.", "Solicite autorizaión")
                        End If
                    End If
                End If
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
    
    Private Sub dato9_KeyPress(KeyAscii As Integer)
        KeyAscii = esNumero(KeyAscii)
        If KeyAscii = 13 Then
            dato9.text = ceros(dato9)
            SendKeys "{Tab}"
        End If
    End Sub
    
    Private Sub dato10_KeyPress(KeyAscii As Integer)
        KeyAscii = esNumero(KeyAscii)
        If KeyAscii = 13 And dato10.text <> "" Then
            dato10.text = ceros(dato10)
            'lblDVV.Caption = rut(dato10.text)
            lblVendedor.Caption = leerNombreVendedor(dato10.text)
            If lblVendedor.Caption <> "" Then
                detalle.Enabled = True
                If detalle.Rows > 1 And (lectura = False Or modifica = True) Then
                    detalle.Cell(1, 1).SetFocus
                End If
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
            Call ctrltostruct(True)
            If dato1.text <> "GD" And dato1.text <> "GM" Then
                detallePagos.Show vbModal
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
        End If
        If dato12.text = "" Then
            dato11.text = "0"
            dato12.text = "0"
        End If
    End Sub
    
    Private Sub dato8_KeyPress(KeyAscii As Integer)
        Dim fechven As String
        KeyAscii = esNumero(KeyAscii)
        If KeyAscii = 13 Then
            If dato8.text <> "" Then
                If Val(dato8.text) > 90 Then
                    dato8.text = "90"
                End If
                fechven = DateAdd("d", dato8.text, fecha)
            Else
                dato8.text = "0"
                fechven = fecha
            End If
            lblDia.Caption = Format(fechven, "dd")
            lblMes.Caption = Format(fechven, "mm")
            lblAño.Caption = Format(fechven, "yyyy")
            SendKeys "{Tab}"
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
                detalle.Cell(i, 9).text = dato11.text
                Rem detalle.Cell(i, 7).text = CDbl(detalle.Cell(i, 3).text) * CDbl(detalle.Cell(i, 5).text) - (CDbl(detalle.Cell(i, 5).text) * CDbl(detalle.Cell(i, 6).text) / 100) * CDbl(detalle.Cell(i, 3).text)
            Next i
        Else
            dato11.text = "0"
            dato12.text = "0"
'            lblTotal.Caption = Round(total - CDbl(dato12.text), 0)
        End If
        Call sumaGrilla(detalle)
    End Sub

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
            dato11.text = Round(desc, 1)
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
    If dato1.text = "BV" Then
    CARGADATOSBOLETA
    
    End If
    
    End Sub

    Private Sub dato6_LostFocus()
        Call limpiaBarra(2)
    End Sub
    
    Private Sub dato7_LostFocus()
        Call limpiaBarra(2)
    End Sub
    
    Private Sub dato9_LostFocus()
        Call limpiaBarra(2)
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

    Private Sub Form_Activate()
        If segurity = True Then
            Seguridad.Show vbModal
            segurity = False
        End If
        Principal.barraEstado.Panels(1).text = UCase(Me.Caption)
    End Sub
    
    Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
        Dim tabla As String
        Dim tipo As String
        If KeyCode = 27 And Screen.ActiveForm.ActiveControl.Name = "dato1" Then
            If imprimio = False Then
                If lectura = False Then
                    Call ctrltostruct(False)
                    Call eliminarVenta(v, detalle)
                End If
            End If
            Unload Me
        End If
        If KeyCode = 80 Then
            If lectura = True Then
                Load detallePagos
                With detallePagos
                    .lectura = True
                    .pagos.Rows = 1
                    .pagos.AutoRedraw = False
                    tabla = "SELECT tipopago, monto, CONCAT('" & vbTab & "', monto, '" & vbTab & "', numerocheque, '" & vbTab & "', banco, '" & vbTab & "', cuentacorriente, '" & vbTab & "', IF(vencimiento <> '00-00-0000',CONCAT(DATE_FORMAT(vencimiento,'%d'), '" & vbTab & "', DATE_FORMAT(vencimiento,'%m'), '" & vbTab & "', DATE_FORMAT(vencimiento,'%Y')),'')) AS item "
                    tabla = tabla & "FROM sv_documento_pagos "
                    tabla = tabla & "WHERE local = '" & empresaActiva & "' AND tipo = '" & dato1.text & "' AND numero = '" & dato2.text & "'"
                    tabla = tabla & "ORDER BY lineapago ASC"
                    Call ConectarControlData(.data, servidor, baseVentas & rubro, usuario, password, tabla)
                    If .data.Recordset.RecordCount > 0 Then
                        .data.Recordset.MoveFirst
                        While Not .data.Recordset.EOF
                            tipo = .data.Recordset.Fields("tipopago")
                            Select Case tipo
                                Case "1"
                                    If Val(.data.Recordset.Fields("monto")) <= 0 Then
                                        tipo = "7 - Vuelto"
                                    Else
                                        tipo = "1 - Efectivo"
                                    End If
                                Case "2"
                                    tipo = "2 - Cheque Propio"
                                Case "3"
                                    tipo = "3 - Cheque Tercero"
                                Case "4"
                                    tipo = "4 - Crédito Directo"
                                'Case "5"
                                '    tipo = "5 - "
                                'Case "6"
                                '    tipo = "6 - "
                            End Select
                            .pagos.AddItem tipo & .data.Recordset.Fields("item"), True
                            .data.Recordset.MoveNext
                        Wend
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
        frmDatos.Enabled = True
        frmDatos.Enabled = False
        
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
        Call CargaGrilla(1, 8)
        dato1.text = "FV"
    End Sub
    
    Private Sub Form_Unload(Cancel As Integer)
        If imprimio = True Then
            Principal.barraEstado.Panels(1).text = UCase(Principal.Caption)
            Call limpiaBarra(2)
        Else
            If lectura = False Then
                Call ctrltostruct(False)
                Call eliminarVenta(v, detalle)
            End If
        End If
    End Sub

'****************************************************************************
'Formato de la Grilla
'****************************************************************************
    Private Sub CargaGrilla(ByVal row As Integer, ByVal col As Integer)
        Dim i As Integer
        Rem DATOS DE LA COLUMNA
        formatoGrilla(1, 0) = "LN"
        formatoGrilla(1, 1) = "CODIGO"
        formatoGrilla(1, 2) = "DESCRIPCION"
        formatoGrilla(1, 3) = "CANTIDAD"
        formatoGrilla(1, 4) = "     "
        formatoGrilla(1, 5) = "PRECIO"
        formatoGrilla(1, 6) = "DESC"
        formatoGrilla(1, 7) = "TOTAL"
        formatoGrilla(1, 8) = "PCOSTO"
        formatoGrilla(1, 9) = ""
        
        Rem LARGO DE LOS DATOS
        formatoGrilla(2, 1) = "15"
        formatoGrilla(2, 2) = "40"
        formatoGrilla(2, 3) = "10"
        formatoGrilla(2, 4) = "9"
        formatoGrilla(2, 5) = "9"
        formatoGrilla(2, 6) = "2"
        formatoGrilla(2, 7) = "9"
        formatoGrilla(2, 8) = "9"
        
        Rem TIPO DE DATOS C=CEROS,S=STRING,N=NUMERICO,D=DATE,
        formatoGrilla(3, 1) = "N"
        formatoGrilla(3, 2) = "S"
        formatoGrilla(3, 3) = "N"
        formatoGrilla(3, 4) = "N"
        formatoGrilla(3, 5) = "N"
        formatoGrilla(3, 6) = "N"
        formatoGrilla(3, 7) = "N"
        formatoGrilla(3, 8) = "N"
        formatoGrilla(3, 9) = "N"
        
        Rem FORMATO GRILLA
        formatoGrilla(4, 1) = "0000000000000"
        formatoGrilla(4, 2) = ""
        formatoGrilla(4, 3) = "#,###,##0.00"
        formatoGrilla(4, 4) = "###,###,##0"
        formatoGrilla(4, 5) = "$ ###,###,##0.00"
        formatoGrilla(4, 6) = "#0.00"
        formatoGrilla(4, 7) = "$ ###,###,##0"
        formatoGrilla(4, 8) = "########0"
        
        Rem LOCCKED
        'FormatoGrilla(5, 0) = "TRUE"
        formatoGrilla(5, 1) = "FALSE"
        formatoGrilla(5, 2) = "TRUE"
        formatoGrilla(5, 3) = "FALSE"
        formatoGrilla(5, 4) = "TRUE"
        formatoGrilla(5, 5) = "FALSE"
        formatoGrilla(5, 6) = "FALSE"
        formatoGrilla(5, 7) = "TRUE"
        formatoGrilla(5, 8) = "TRUE"
        
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
        formatoGrilla(8, 1) = "10"
        formatoGrilla(8, 2) = "35"
        formatoGrilla(8, 3) = "10"
        formatoGrilla(8, 4) = "6"
        formatoGrilla(8, 5) = "12"
        formatoGrilla(8, 6) = "5"
        formatoGrilla(8, 7) = "12"
        formatoGrilla(8, 8) = "0"
        formatoGrilla(8, 9) = "0"
            
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
        
        detalle.Cell(0, 0).text = formatoGrilla(1, 0)
        For i = 1 To col - 1
            detalle.Cell(0, i).text = formatoGrilla(1, i)
            detalle.Column(i).Width = Val(formatoGrilla(8, i)) * (detalle.Cell(0, i).Font.Size + 1.25)
            detalle.Column(i).MaxLength = Val(formatoGrilla(2, i))
            detalle.Column(i).FormatString = formatoGrilla(4, i)
            detalle.Column(i).Locked = formatoGrilla(5, i)
            If formatoGrilla(3, i) = "N" Then
                detalle.Column(i).Alignment = cellRightCenter
                If i <> 5 Then
                    detalle.Column(i).Mask = cellNumeric
                End If
            Else
                detalle.Column(i).Alignment = cellLeftCenter
                detalle.Column(i).Mask = cellUpper
            End If
        Next i
        detalle.Range(0, 0, 0, detalle.Cols - 1).Alignment = cellCenterCenter
        detalle.AddItem vbTab & vbTab & "1" & vbTab & "0" & vbTab & "0" & vbTab & "0" & vbTab & "0" & vbTab & "0", True
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
                Else
                    vacio = True
                End If
                If KeyCode = 38 And detalle.ActiveCell.row = 1 And detalle.ActiveCell.col = 1 Then
                    dato10.SetFocus
                End If
                If KeyCode = 40 And detalle.ActiveCell.row = detalle.Rows - 1 And detalle.ActiveCell.col = 1 Then
                    dato11.SetFocus
                End If
            Case 106
                'SendKeys "{Tab}"
                dato11.SetFocus
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
                            Case "FV", "FE", "GD"
                                limite = 35
                                
                            Case "BV"
                                limite = 13
                            Case Else
                                limite = 0
                        End Select
                        If limite > 0 Then
                            If limite > detalle.Rows - 1 Then
                                detalle.AddItem vbTab & vbTab & "1" & vbTab & "0" & vbTab & "0" & vbTab & "0" & vbTab & "0", True
                                
                                NewRow = detalle.Rows - 1
                                NewCol = 1
                            Else
                                dato11.SetFocus
                            End If
                        Else
                            detalle.AddItem vbTab & vbTab & "1" & vbTab & "0" & vbTab & "0" & vbTab & "0" & vbTab & "0" & vbTab & "0", True
                            
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
                pivote.text = ceros(pivote)
                If Val(pivote.text) = 0 Then
                    pivote.text = ""
                End If
                detalle.Cell(row, 1).text = pivote.text
                detalle.Cell(row, detalle.Cols - 1).text = leerCostoProducto(detalle.Cell(row, 1).text)
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
                For i = 1 To detalle.Cols - 2
                    If detalle.Cell(row, i).text = "" Then
                        NewRow = row
                        NewCol = i
                        Exit For
                    End If
                Next i
                For i = 1 To detalle.Cols - 2
                    If detalle.Cell(NewRow, i).text = "" Then
                        NewCol = i
                        Exit For
                    End If
                Next i
            End If
            If row > 0 Then
                If detalle.Cell(row, 3).text <> "" And detalle.Cell(row, 5).text <> "" Then
                    If dato1.text <> "ZE" Then
                        detalle.Cell(row, 7).text = Round(detalle.Cell(row, 3).text * detalle.Cell(row, 5).text + 0.1, 0)
                    Else
                        If Val(detalle.Cell(row, 7).text) <> Val(detalle.Cell(row, 3).text) * Val(detalle.Cell(row, 5).text) Then
                            detalle.Cell(row, 7).text = detalle.Cell(row, 5).text
                            detalle.Cell(row, 5).text = detalle.Cell(row, 7).text / detalle.Cell(row, 3).text
                        End If
                    End If
                    detalle.Cell(row, 4).text = CDbl(leerUnidadesProducto(detalle.Cell(row, 1).text)) * CDbl(detalle.Cell(row, 3).text)
                End If
            End If
'            If row > 0 Then
'                precio = Round(detalle.Cell(row, 3).text * detalle.Cell(row, 5).text + 0.1, 0)
'                descu = Int((detalle.Cell(row, 7).text * ((detalle.Cell(row, 6).text) / 100)) + 0.5)
'                detalle.Cell(row, 7).text = Str(precio - descu)
'            End If
            If col > 0 And row > 0 Then
                Call sumaGrilla(detalle)
            End If
        End If
    End Sub
    Private Sub sumaGrilla(ByRef Lista As Grid)
        Dim cad As String
        Dim i As Integer
        Dim suma As Double
        Dim sumaIVA As Double
        Dim sumaIHA As Double
        Dim sumaEXENTO As Double
        Dim codigo As String
        Dim deci As String
        Dim precio As Double
        Dim descu As Double
        Dim descu2 As Double
        
        suma = 0
        sumaIVA = 0
        sumaIHA = 0
        sumaEXENTO = 0
        descu2 = 0
        For i = 1 To Lista.Rows - 1
            codigo = Lista.Cell(i, 1).text
            If codigo <> "" Then
                If Val(Lista.Cell(i, 3).text) = 0 Then
                    Lista.Cell(i, 3).text = "0"
                End If
                If Val(Lista.Cell(i, 5).text) = 0 Then
                    Lista.Cell(i, 5).text = "0"
                End If
                precio = Int(Lista.Cell(i, 3).text * Lista.Cell(i, 5).text)
                If Lista.Cell(i, 6).text <> "" Then descu = Int((precio * ((Lista.Cell(i, 6).text) / 100)) + 0.5)
                Lista.Cell(i, 7).text = Str(precio - descu)
                precio = Lista.Cell(i, 7).text
                descu2 = descu2 + Int(((precio * dato11.text / 100)) + 0.5)
                sumaIVA = sumaIVA + CDbl(Lista.Cell(i, 7).text)
                
                                
                suma = suma + CDbl(Lista.Cell(i, 7).text)
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
        
        Select Case dato1.text
            Case "BV", "ZE"
                lblNeto.Caption = Round((CDbl(lblSub.Caption) - CDbl(dato12.text)) / (1 + iva / 100), 0)
                lblIVA.Caption = Round(CDbl(lblSub.Caption) - CDbl(dato12.text) - CDbl(lblNeto.Caption), 0)
                lblIHA.Caption = "0"
            Case "FE"
                lblNeto.Caption = Round(CDbl(lblSub.Caption) - CDbl(dato12.text), 0)
                lblIVA.Caption = "0"
                lblIHA.Caption = "0"
            Case "GD"
                lblNeto.Caption = Round(CDbl(lblSub.Caption) - CDbl(dato12.text), 0)
                lblIVA.Caption = Round(sumaIVA * iva / 100, 0)
                lblIHA.Caption = "0"
            Case Else
                lblNeto.Caption = Round(CDbl(lblSub.Caption) - CDbl(dato12.text), 0)
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
        v.cabeza.tipo = dato1.text
        v.cabeza.numero = dato2.text
        v.cabeza.fecha = dato5.text & "-" & dato4.text & "-" & dato3.text
        v.cabeza.plazo = dato8.text
        v.cabeza.vencimiento = lblAño.Caption & "-" & lblMes.Caption & "-" & lblDia.Caption
        v.cabeza.rut = dato6.text & lblDV.Caption
        v.cabeza.sucursal = dato7.text
        v.cabeza.cajera = dato10.text & lblDVV.Caption
        v.cabeza.notapedido = dato9.text
        v.cabeza.notaventas = ""
        v.cabeza.ordencompra = ""
        v.cabeza.subtotal = Format(lblSub.Caption, "########0")
        v.cabeza.neto = Format(lblNeto.Caption, "########0")
        v.cabeza.iva = Format(lblIVA.Caption, "########0")
        v.cabeza.impuestoharina = Format(lblIHA.Caption, "########0")
        v.cabeza.impuestoila = ""
        v.cabeza.impuestoespecifico = ""
        v.cabeza.exento = ""
        v.cabeza.retencionparcial = ""
        v.cabeza.retenciontotal = ""
        v.cabeza.total = Format(lblTotal.Caption, "########0")
        v.cabeza.abono = ""
        v.cabeza.descuento = Replace(dato12.text, ".", ",")
        v.cabeza.contabilizado = ""
        v.cabeza.PAGADO = ""
        v.cabeza.comision = ""
        v.cabeza.fechapagocomision = ""
        v.cabeza.nula = "N"
        v.cabeza.boletadesde = desde
        v.cabeza.boletahasta = hasta
        
        v.detalle.loc = empresaActiva
        v.detalle.tipo = dato1.text
        v.detalle.numero = dato2.text
        v.detalle.linea = ""
        v.detalle.fecha = dato5.text & "-" & dato4.text & "-" & dato3.text
        v.detalle.rut = dato6.text & lblDV.Caption
        v.detalle.sucursal = dato7.text
        v.detalle.codigo = ""
        v.detalle.descripcion = ""
        v.detalle.cantidad = ""
        v.detalle.unidades = ""
        v.detalle.precio = ""
        v.detalle.descuento = ""
        v.detalle.total = ""
        v.detalle.vendedor = dato10.text
        v.detalle.pcosto = ""
        v.detalle.bodega = bodega
        v.detalle.vencimiento = lblAño.Caption & "-" & lblMes.Caption & "-" & lblDia.Caption
        
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
        lblDireccion.Caption = c.direccion
        lblComuna.Caption = c.comuna
        lblCiudad.Caption = c.ciudad
        dato8.text = c.plazo
        If dato8.text <> "" Then
            fechven = DateAdd("d", dato8.text, fecha)
        Else
            fechven = fecha
        End If
        lblDia.Caption = Format(fechven, "dd")
        lblMes.Caption = Format(fechven, "mm")
        lblAño.Caption = Format(fechven, "yyyy")
    End Sub
    
    Private Sub structtoctrl()
        Dim subtotal As Double
        Dim descuento As Double
        
        dato1.text = v.cabeza.tipo
        dato2.text = v.cabeza.numero
        fecha = Format(v.cabeza.fecha, "dd-mm-yyyy")
        dato3.text = Format(v.cabeza.fecha, "dd")
        dato4.text = Format(v.cabeza.fecha, "mm")
        dato5.text = Format(v.cabeza.fecha, "yyyy")
        dato8.text = v.cabeza.plazo
        lblDia.Caption = Format(v.cabeza.vencimiento, "dd")
        lblMes.Caption = Format(v.cabeza.vencimiento, "mm")
        lblAño.Caption = Format(v.cabeza.vencimiento, "yyyy")
        dato6.text = v.cabeza.rut
        dato7.text = v.cabeza.sucursal
        dato10.text = v.cabeza.cajera
        dato9.text = v.cabeza.notapedido
        lblSub.Caption = v.cabeza.subtotal
        lblNeto.Caption = v.cabeza.neto
        lblIVA.Caption = v.cabeza.iva
        lblIHA.Caption = v.cabeza.impuestoharina
        lblTotal.Caption = v.cabeza.total
        If v.cabeza.descuento <> 0 Then
        descuento = CDbl(v.cabeza.descuento) / CDbl(v.cabeza.subtotal) * 100
        dato11.text = descuento
        End If
        dato12.text = Replace(v.cabeza.descuento, ",", ".")
        If dato1.text = "ZE" Then
            desde = String(10 - Len(v.cabeza.boletadesde), "0") & v.cabeza.boletadesde
            hasta = String(10 - Len(v.cabeza.boletahasta), "0") & v.cabeza.boletahasta
            lblDesde.Caption = "Desde: " & desde
            lblHasta.Caption = "Hasta: " & hasta
            lblDesde.Visible = True
            lblHasta.Visible = True
        End If
        
        Call dato6_KeyPress(13)
        Call dato7_KeyPress(13)
        Call dato10_KeyPress(13)
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
            Call eliminar
        Case Asc("i"), Asc("I")
            Call imprimir
        Case Asc("r"), Asc("R"), 46
            Call retorno
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
        If dato1.text = "ZE" Then
            modifica = True
            Call HabilitarCajas(Me, modifica)
            dato1.Enabled = False
            dato2.Enabled = False
            dato3.Enabled = True
            detalle.SelectionMode = cellSelectionFree
            dato3.SetFocus
        End If
    End Sub
    
    Private Sub eliminar()
        Select Case MsgBox("Si desea eliminar el documento presione SI" & vbCrLf & "Si desea anular el documento presione NO" & vbCrLf & "Presione CANCELAR para volver", vbYesNoCancel, "Alerta")
            Case vbYes
                Call ctrltostruct(False)
                Call eliminarVenta(v, detalle)
                Call eliminarPagos(p.tipodocumento, p.numeroDocumento)
                Call retorno
                Call HabilitarCajas(Me, modifica)
                dato1.SetFocus
            Case vbNo
                If lblNulo.Caption = "" Then
                    v.detalle.fecha = dato5.text & "-" & dato4.text & "-" & dato3.text
                    Call anularDocumento(dato1.text, dato2.text, detalle, v.detalle)
                    Call eliminarPagos(p.tipodocumento, p.numeroDocumento)
                    Call eliminarDocManual(dato1.text, dato2.text)
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
                    
                    ImpresionBoleta.Show
                Case "FV", "FE"
                    Call imprimeFactura(dato1.text, dato2.text, Impresion, data)
                Case "GD", "GM"
                    Call imprimeGuia(dato1.text, dato2.text, Impresion, data)
                Case "ZE"
            End Select
        Else
            Call MsgBox("Documento nulo no imprimible", vbOKOnly, "Mensaje")
        End If
    End Sub
    
    Private Sub retorno()
        Call LimpiarCajas(Me)
        'Call LimpiarLabels(Me)
        Call CargaGrilla(1, 8)
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
        lblSub.Caption = "0"
        lblNeto.Caption = "0"
        lblIVA.Caption = "0"
        lblIHA.Caption = "0"
        lblTotal.Caption = "0"
        Call HabilitarCajas(Me, modifica)
        lblDesde.Visible = False
        lblHasta.Visible = False
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

    Private Function leerPrecioEspecial(ByVal codigo As String) As String
        Dim condicion As String
        Dim campos(10, 3) As String
        Dim op As Integer
        Set sql = New CSQLUtil
        campos(0, 0) = "precioespecial"
        campos(1, 0) = ""
        
        campos(0, 2) = "sv_maestroclientes_especiales"
        
        condicion = "rut = '" & dato6.text & lblDV.Caption & "' AND sucursal = '" & dato7.text & "' AND codigo = '" & codigo & "'"
        op = 5
        sql.datos = campos
        Set sql.conexion = ventas
        Call sql.SQLUTIL(op, condicion)
        If sql.estado = 0 Then
            leerPrecioEspecial = sql.datos(0, 3)
        Else
            leerPrecioEspecial = ""
        End If
    End Function

    Private Sub opciones_GotFocus()
        manual.SetFocus
    End Sub


Private Sub PREVENTA_GotFocus()
PREVENTA.text = ""

End Sub

Private Sub PREVENTA_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
frmDatos.Enabled = True
frmDetalle.Enabled = True
FRMPREVENTA.Visible = False

dato3.SetFocus
PREVENTA.text = ceros(PREVENTA)

leepreventa
            
End If


End Sub
Sub leepreventa()
    Dim resultados As rdoResultset
    Dim cSql As New rdoQuery
    Dim rut As String
    Dim linea As Double
        Set cSql.ActiveConnection = ventasRubro
        cSql.sql = "SELECT codigo,descripcion,cantidad,precio,descuento,total,descuento2,vendedor "
        cSql.sql = cSql.sql + "FROM sv_documento_detalle "
        cSql.sql = cSql.sql + "WHERE local='" + empresaActiva + "' and tipo='PV' and numero='" + PREVENTA.text + "' "
        cSql.Execute
        linea = detalle.Rows - 2
        
        If cSql.RowsAffected > 0 Then
            
            Set resultados = cSql.OpenResultset
            dato10.text = resultados(6)
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
 ' borrapreventa

End Sub
Sub borrapreventa()
    Dim resultados As rdoResultset
    Dim cSql As New rdoQuery
    Dim rut As String
    Dim linea As Double
        Set cSql.ActiveConnection = ventasRubro
        cSql.sql = "DELETE "
        cSql.sql = cSql.sql + "FROM sv_documento_detalle "
        cSql.sql = cSql.sql + "WHERE local='" + empresaActiva + "' and tipo='PV' and numero='" + PREVENTA.text + "' "
        cSql.Execute
            
            
            Set cSql.ActiveConnection = ventasRubro
        cSql.sql = "DELETE "
        cSql.sql = cSql.sql + "FROM sv_documento_cabeza "
        cSql.sql = cSql.sql + "WHERE local='" + empresaActiva + "' and tipo='PV' and numero='" + PREVENTA.text + "' "
        cSql.Execute
    
        
End Sub

Sub CARGADATOSBOLETA()
dato3.text = Format(fechasistema, "dd")
dato4.text = Format(fechasistema, "mm")
dato5.text = Format(fechasistema, "yyyy")
dato6.text = "000000001"
lblDV.Caption = "9"
dato7.text = "0"
dato8.text = "0000000000"
dato9.text = "0000000000"
dato10.SetFocus



End Sub
