VERSION 5.00
Object = "{4777436C-EB8C-4596-98A8-EBCF98683969}#1.0#0"; "Flexcell.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{ADBBAED6-B16F-40DC-80DF-B44910CBA76C}#1.0#0"; "Frame-Xp.ocx"
Begin VB.Form DetalleDocumento 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   9240
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   13605
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   9240
   ScaleWidth      =   13605
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox manual 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   300
      ScaleHeight     =   375
      ScaleWidth      =   555
      TabIndex        =   0
      Top             =   7980
      Width           =   555
   End
   Begin XPFrame.FrameXp frmAnterior 
      Height          =   375
      Left            =   1380
      TabIndex        =   57
      Top             =   8340
      Width           =   2475
      _ExtentX        =   4366
      _ExtentY        =   661
      BackColor       =   49344
      Caption         =   "Anterior"
      CaptionEstilo3D =   1
      BackColor       =   49344
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
   End
   Begin XPFrame.FrameXp frmDatos 
      Height          =   2655
      Left            =   60
      TabIndex        =   1
      Top             =   60
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
      Begin XPFrame.FrameXp frmCerrar 
         Height          =   255
         Left            =   13140
         TabIndex        =   29
         Top             =   30
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   450
         BackColor       =   49344
         Caption         =   "X"
         CaptionEstilo3D =   1
         BackColor       =   49344
         ColorBarraArriba=   8388608
         ColorBarraAbajo =   16761024
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
      End
      Begin VB.Label lblNota 
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
         Left            =   11520
         TabIndex        =   38
         Top             =   1920
         Width           =   1815
      End
      Begin VB.Label lblRutVendedor 
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
         Left            =   1920
         TabIndex        =   37
         Top             =   2280
         Width           =   1455
      End
      Begin VB.Label lblSucursal 
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
         Left            =   5880
         TabIndex        =   36
         Top             =   840
         Width           =   675
      End
      Begin VB.Label lblRut 
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
         Left            =   1920
         TabIndex        =   35
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label lblDiaVenta 
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
         Left            =   11160
         TabIndex        =   34
         Top             =   480
         Width           =   495
      End
      Begin VB.Label lblMesVenta 
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
         Left            =   11760
         TabIndex        =   33
         Top             =   480
         Width           =   495
      End
      Begin VB.Label lblAñoVenta 
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
         Left            =   12360
         TabIndex        =   32
         Top             =   480
         Width           =   975
      End
      Begin VB.Label lblNumero 
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
         Left            =   7680
         TabIndex        =   31
         Top             =   480
         Width           =   1575
      End
      Begin VB.Label lblTipo 
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
         TabIndex        =   30
         Top             =   480
         Width           =   555
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
         TabIndex        =   26
         Top             =   2280
         Width           =   375
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
         TabIndex        =   25
         Top             =   1920
         Width           =   975
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
         TabIndex        =   24
         Top             =   1920
         Width           =   495
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Sucursal"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   4080
         TabIndex        =   23
         Top             =   840
         Width           =   1695
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
         Left            =   3870
         TabIndex        =   22
         Top             =   2280
         Width           =   9465
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
         TabIndex        =   20
         Top             =   1920
         Width           =   2295
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
         TabIndex        =   19
         Top             =   1560
         Width           =   4695
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
         TabIndex        =   18
         Top             =   1560
         Width           =   4695
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
         TabIndex        =   17
         Top             =   1200
         Width           =   11415
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
         TabIndex        =   16
         Top             =   840
         Width           =   4815
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
         Left            =   2520
         TabIndex        =   15
         Top             =   480
         Width           =   3255
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
         TabIndex        =   14
         Top             =   840
         Width           =   375
      End
      Begin VB.Label lbl8 
         Appearance      =   0  'Flat
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Ciudad"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   6840
         TabIndex        =   13
         Top             =   1560
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
         TabIndex        =   12
         Top             =   1560
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
         TabIndex        =   11
         Top             =   480
         Width           =   1695
      End
      Begin VB.Label lbl3 
         Appearance      =   0  'Flat
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Fecha"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   9360
         TabIndex        =   10
         Top             =   480
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
         TabIndex        =   9
         Top             =   840
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
         TabIndex        =   8
         Top             =   840
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
         TabIndex        =   7
         Top             =   1200
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
         TabIndex        =   6
         Top             =   1920
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
         TabIndex        =   5
         Top             =   1920
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
         TabIndex        =   4
         Top             =   2280
         Width           =   1695
      End
      Begin VB.Label lbl2 
         Appearance      =   0  'Flat
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Número"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   5880
         TabIndex        =   3
         Top             =   480
         Width           =   1695
      End
      Begin VB.Label lbl10 
         Appearance      =   0  'Flat
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Vencimiento"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   4920
         TabIndex        =   2
         Top             =   1920
         Width           =   1695
      End
   End
   Begin XPFrame.FrameXp frmDetalle 
      Height          =   4275
      Left            =   60
      TabIndex        =   27
      Top             =   2820
      Width           =   13455
      _ExtentX        =   23733
      _ExtentY        =   7541
      BackColor       =   16744576
      Caption         =   "Productos en la Venta"
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
         TabIndex        =   28
         Top             =   360
         Width           =   13470
         _ExtentX        =   23760
         _ExtentY        =   6800
         Cols            =   5
         DefaultFontSize =   8.25
         Rows            =   30
         SelectionMode   =   1
      End
   End
   Begin XPFrame.FrameXp frmResumen 
      Height          =   1935
      Left            =   8100
      TabIndex        =   39
      Top             =   7200
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
      Begin VB.Label lblDescuentoPesos 
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
         TabIndex        =   53
         Top             =   1200
         Width           =   1215
      End
      Begin VB.Label lblDescuentoPorcentaje 
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
         TabIndex        =   52
         Top             =   840
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
         TabIndex        =   51
         Top             =   480
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
         TabIndex        =   50
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
         TabIndex        =   49
         Top             =   840
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
      Begin VB.Label lbl17 
         Appearance      =   0  'Flat
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " IVA"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   2760
         TabIndex        =   45
         Top             =   840
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
         TabIndex        =   44
         Top             =   840
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
         TabIndex        =   43
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
         TabIndex        =   42
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
         TabIndex        =   41
         Top             =   1560
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
         TabIndex        =   40
         Top             =   1560
         Width           =   1215
      End
   End
   Begin MSAdodcLib.Adodc data 
      Height          =   330
      Left            =   6780
      Top             =   6660
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
   Begin XPFrame.FrameXp frmSiguiente 
      Height          =   375
      Left            =   4200
      TabIndex        =   58
      Top             =   8340
      Width           =   2475
      _ExtentX        =   4366
      _ExtentY        =   661
      BackColor       =   49344
      Caption         =   "Siguiente"
      CaptionEstilo3D =   1
      BackColor       =   49344
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
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "* o Fin para terminar venta"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   4680
      TabIndex        =   59
      Top             =   7245
      Width           =   3345
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
      Left            =   1380
      TabIndex        =   56
      Top             =   7740
      Visible         =   0   'False
      Width           =   2475
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
      Left            =   4200
      TabIndex        =   55
      Top             =   7740
      Visible         =   0   'False
      Width           =   2415
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
      Left            =   60
      TabIndex        =   54
      Top             =   7200
      Width           =   7935
   End
End
Attribute VB_Name = "DetalleDocumento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Private formatogrilla(10, 10) As String
    Private v As ventaDocumento
    Private c As Cliente
    Private nula As Boolean
    Public TIPO As String
    Public NUMERO As String
    Public fechaAudit As String
    Public cajaaudit As String
    
    
Private Sub Form_Activate()
    manual.SetFocus
    If leerVentaDocumento(v, TIPO, NUMERO, "=", data, detalle, cajaaudit) = True Then
        If v.cabeza.nula = "N" Then
            Call structtoctrl
        
            If LEERCLIENTE(c, lblRut.Caption & lblDV.Caption, lblSucursal.Caption, "=") = True Then
                structtoctrlCliente
            End If
            
            nula = leerDocumentoNulo(TIPO, NUMERO)
            If nula = True Then
                lblNulo.Caption = "DOCUMENTO ANULADO"
            Else
                lblNulo.Caption = ""
            End If
        End If
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 27
            Unload Me
        Case Else
            manual.SetFocus
    End Select
End Sub

Private Sub Form_Load()
    Call CARGAGRILLA(1, 8)
End Sub

    Private Sub frmAnterior_BarMouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
        Call cambiaColor(frmAnterior)
        frmAnterior.CaptionEstilo3D = Raised
    End Sub
    
    Private Sub frmAnterior_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
        Call cambiaColor(frmAnterior)
        frmAnterior.CaptionEstilo3D = Inserted
        If leerVentaDocumento(v, TIPO, NUMERO, "<", data, detalle) = True Then
            If v.cabeza.nula = "N" Then
                Call structtoctrl
                If LEERCLIENTE(c, lblRut.Caption & lblDV.Caption, lblSucursal.Caption, "=") = True Then
                    structtoctrlCliente
                End If
            Else
                NUMERO = v.cabeza.NUMERO
                Call frmAnterior_MouseUp(Button, Shift, x, Y)
            End If
        End If
        manual.SetFocus
    End Sub
    
    Private Sub frmSiguiente_BarMouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
        Call cambiaColor(frmSiguiente)
        frmSiguiente.CaptionEstilo3D = Raised
    End Sub
    
    Private Sub frmSiguiente_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
        Call cambiaColor(frmSiguiente)
        frmSiguiente.CaptionEstilo3D = Inserted
        If leerVentaDocumento(v, TIPO, NUMERO, ">", data, detalle) = True Then
            If v.cabeza.nula = "N" Then
                Call structtoctrl
                If LEERCLIENTE(c, lblRut.Caption & lblDV.Caption, lblSucursal.Caption, "=") = True Then
                    structtoctrlCliente
                End If
            Else
                NUMERO = v.cabeza.NUMERO
                Call frmSiguiente_MouseUp(Button, Shift, x, Y)
            End If
        End If
        manual.SetFocus
    End Sub

    Private Sub frmCerrar_BarMouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
        Call cambiaColor(frmCerrar)
        frmCerrar.CaptionEstilo3D = Raised
    End Sub

    Private Sub frmCerrar_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
        Call cambiaColor(frmCerrar)
        frmCerrar.CaptionEstilo3D = Inserted
        Unload Me
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
        formatogrilla(1, 4) = ""
        formatogrilla(1, 5) = "PRECIO"
        formatogrilla(1, 6) = "DESC"
        formatogrilla(1, 7) = "TOTAL"
        
        Rem LARGO DE LOS DATOS
        formatogrilla(2, 1) = "15"
        formatogrilla(2, 2) = "40"
        formatogrilla(2, 3) = "10"
        formatogrilla(2, 4) = "9"
        formatogrilla(2, 5) = "9"
        formatogrilla(2, 6) = "2"
        formatogrilla(2, 7) = "9"
        
        Rem TIPO DE DATOS C=CEROS,S=STRING,N=NUMERICO,D=DATE,
        formatogrilla(3, 1) = "N"
        formatogrilla(3, 2) = "S"
        formatogrilla(3, 3) = "N"
        formatogrilla(3, 4) = "N"
        formatogrilla(3, 5) = "N"
        formatogrilla(3, 6) = "N"
        formatogrilla(3, 7) = "N"
        
        Rem FORMATO GRILLA
        formatogrilla(4, 1) = "0000000000000"
        formatogrilla(4, 2) = ""
        formatogrilla(4, 3) = "#,###,##0.00"
        formatogrilla(4, 4) = "###,###,##0.00"
        formatogrilla(4, 5) = "$ ###,###,##0.00"
        formatogrilla(4, 6) = "#0.00"
        formatogrilla(4, 7) = "$ ###,###,##0"
        
        Rem LOCCKED
        'FormatoGrilla(5, 0) = "TRUE"
        formatogrilla(5, 1) = "TRUE"
        formatogrilla(5, 2) = "TRUE"
        formatogrilla(5, 3) = "TRUE"
        formatogrilla(5, 4) = "TRUE"
        formatogrilla(5, 5) = "TRUE"
        formatogrilla(5, 6) = "TRUE"
        formatogrilla(5, 7) = "TRUE"
        
        Rem VALOR MINIMO
        formatogrilla(6, 1) = ""
        formatogrilla(6, 2) = ""
        formatogrilla(6, 3) = ""
        formatogrilla(6, 4) = ""
        formatogrilla(6, 5) = ""
        formatogrilla(6, 6) = ""
        formatogrilla(6, 7) = ""
        
        Rem VALOR MAXIMO
        formatogrilla(7, 1) = ""
        formatogrilla(7, 2) = ""
        formatogrilla(7, 3) = ""
        formatogrilla(7, 4) = ""
        formatogrilla(7, 5) = ""
        formatogrilla(7, 6) = ""
        formatogrilla(7, 7) = ""
        
        Rem ANCHO
        formatogrilla(8, 1) = "10"
        formatogrilla(8, 2) = "35"
        formatogrilla(8, 3) = "10"
        formatogrilla(8, 4) = "6"
        formatogrilla(8, 5) = "12"
        formatogrilla(8, 6) = "5"
        formatogrilla(8, 7) = "12"
            
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
                If i <> 5 Then
                    detalle.Column(i).Mask = cellNumeric
                End If
            Else
                detalle.Column(i).Alignment = cellLeftCenter
                detalle.Column(i).Mask = cellUpper
            End If
        Next i
        detalle.Range(0, 0, 0, detalle.Cols - 1).Alignment = cellCenterCenter
        detalle.AddItem vbTab & vbTab & "1" & vbTab & "0" & vbTab & "0" & vbTab & "0" & vbTab & "0", True
        detalle.Enabled = True
    End Sub

'=============================================================================
'PASA LOS DATOS DE LA ESTRUCTURA A LOS CONTROLES
'=============================================================================
    Private Sub structtoctrlCliente()
        Dim fecha As String
        lblRazon.Caption = c.nombre
        lblDireccion.Caption = c.direccion
        lblComuna.Caption = c.comuna
        lblCiudad.Caption = c.ciudad
    End Sub
    
    Private Sub structtoctrl()
        Dim fecha As String
        lblTipo.Caption = v.cabeza.TIPO
        lblDocumento.Caption = leerNombreDocumento(lblTipo.Caption)
        lblNumero.Caption = v.cabeza.NUMERO
        NUMERO = v.cabeza.NUMERO
        fecha = v.cabeza.fecha
        lblDiaVenta.Caption = Format(fecha, "dd")
        lblMesVenta.Caption = Format(fecha, "mm")
        lblAñoVenta.Caption = Format(fecha, "yyyy")
        lblRut.Caption = Left(v.cabeza.rut, 9)
        lblDV.Caption = rut(lblRut.Caption)
        lblSucursal.Caption = v.cabeza.sucursal
        lblDias.Caption = v.cabeza.plazo
        fecha = v.cabeza.vencimiento
        lblDia.Caption = Format(fecha, "dd")
        lblMes.Caption = Format(fecha, "mm")
        lblAño.Caption = Format(fecha, "yyyy")
        lblNota.Caption = v.cabeza.notapedido
        lblRutVendedor.Caption = Left(v.cabeza.cajera, 9)
        lblDVV.Caption = rut(lblRutVendedor.Caption)
        lblVendedor.Caption = leerNombreVendedor(lblRutVendedor.Caption)
        lblSub.Caption = v.cabeza.subtotal
        lblDescuentoPesos.Caption = v.cabeza.Descuento
        lblDescuentoPorcentaje.Caption = Round(CDbl(lblDescuentoPesos.Caption) * 100 / CDbl(lblSub.Caption), 1)
        lblNeto.Caption = v.cabeza.neto
        lblIVA.Caption = v.cabeza.iva
        lblIHA.Caption = v.cabeza.impuestoHarina
        lblTotal.Caption = v.cabeza.total
        If lblTipo.Caption = "ZE" Then
            lblDesde.Caption = "Desde: " & String(10 - Len(v.cabeza.boletadesde), "0") & v.cabeza.boletadesde
            lblHasta.Caption = "Hasta: " & String(10 - Len(v.cabeza.boletahasta), "0") & v.cabeza.boletahasta
            lblDesde.Visible = True
            lblHasta.Visible = True
        End If
    End Sub
'=============================================================================
'PASA LOS DATOS DE LA ESTRUCTURA A LOS CONTROLES
'=============================================================================

Private Sub manual_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 27
            Unload Me
        Case Asc("a"), Asc("A"), 37
            Call frmAnterior_BarMouseDown(1, 0, 130, 15)
            Call frmAnterior_MouseUp(1, 0, 130, 15)
        Case Asc("s"), Asc("S"), 39
            Call frmSiguiente_BarMouseDown(1, 0, 60, 15)
            Call frmSiguiente_MouseUp(1, 0, 60, 15)
    End Select
End Sub
