VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash10d.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{4777436C-EB8C-4596-98A8-EBCF98683969}#1.0#0"; "FlexCell.ocx"
Object = "{ADBBAED6-B16F-40DC-80DF-B44910CBA76C}#1.0#0"; "Frame-Xp.ocx"
Begin VB.Form activofijotb01 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Maestro de Activo Fijos (TRIBUTARIO)"
   ClientHeight    =   8340
   ClientLeft      =   2235
   ClientTop       =   1305
   ClientWidth     =   13860
   DrawStyle       =   5  'Transparent
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   556
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   924
   ShowInTaskbar   =   0   'False
   Begin XPFrame.FrameXp FrameQuickMenu 
      Height          =   615
      Left            =   10560
      TabIndex        =   49
      Top             =   7560
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   1085
      BackColor       =   16744576
      Caption         =   " Mis Datos"
      BackColor       =   16744576
      BordeColor      =   4194304
      ColorBarraArriba=   4194304
      ColorBarraAbajo =   4194304
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.CommandButton botonmisaccesos 
         Caption         =   "Permisos Modulo"
         Height          =   255
         Left            =   120
         TabIndex        =   51
         Top             =   280
         Width           =   1455
      End
      Begin VB.CommandButton botonmisfavoritos 
         Caption         =   "Mis Favoritos"
         Height          =   255
         Left            =   1800
         TabIndex        =   50
         Top             =   280
         Width           =   1335
      End
   End
   Begin FlexCell.Grid impresion 
      Height          =   255
      Left            =   1320
      TabIndex        =   46
      Top             =   6960
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   450
      Cols            =   5
      DefaultFontSize =   8.25
      Rows            =   30
   End
   Begin XPFrame.FrameXp FrameXp1 
      Height          =   6735
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   13815
      _ExtentX        =   24368
      _ExtentY        =   11880
      BackColor       =   16761024
      Caption         =   "Datos"
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
      Begin VB.TextBox dato12 
         Alignment       =   1  'Right Justify
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
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   2640
         MaxLength       =   8
         TabIndex        =   39
         Tag             =   "nombre"
         Top             =   3360
         Width           =   1215
      End
      Begin VB.TextBox dato13 
         Alignment       =   1  'Right Justify
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
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   2640
         MaxLength       =   10
         TabIndex        =   38
         Tag             =   "nombre"
         Top             =   3720
         Width           =   1215
      End
      Begin VB.TextBox dato14 
         Alignment       =   1  'Right Justify
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
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   2640
         MaxLength       =   9
         TabIndex        =   37
         Tag             =   "nombre"
         Top             =   4080
         Width           =   1215
      End
      Begin XPFrame.FrameXp FrameXp2 
         Height          =   5295
         Left            =   8160
         TabIndex        =   35
         Top             =   240
         Width           =   5535
         _ExtentX        =   9763
         _ExtentY        =   9340
         BackColor       =   16744576
         Caption         =   "Datos"
         CaptionEstilo3D =   1
         BackColor       =   16744576
         BordeColor      =   4194304
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Begin FlexCell.Grid Grid1 
            Height          =   4815
            Left            =   120
            TabIndex        =   36
            Top             =   360
            Width           =   5295
            _ExtentX        =   9340
            _ExtentY        =   8493
            Cols            =   5
            DefaultFontSize =   8.25
            Rows            =   30
         End
      End
      Begin VB.TextBox dato11 
         Alignment       =   1  'Right Justify
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
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   2640
         MaxLength       =   50
         TabIndex        =   21
         Tag             =   "nombre"
         Top             =   3000
         Width           =   1215
      End
      Begin VB.TextBox dato10 
         Alignment       =   1  'Right Justify
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
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   2640
         MaxLength       =   50
         TabIndex        =   19
         Tag             =   "nombre"
         Top             =   2640
         Width           =   1215
      End
      Begin VB.TextBox dato9 
         Alignment       =   1  'Right Justify
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
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   2640
         MaxLength       =   50
         TabIndex        =   17
         Tag             =   "nombre"
         Top             =   2280
         Width           =   1215
      End
      Begin VB.TextBox dato8 
         Alignment       =   1  'Right Justify
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
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   2640
         MaxLength       =   50
         TabIndex        =   0
         Tag             =   "nombre"
         Top             =   1920
         Width           =   495
      End
      Begin VB.TextBox dato7 
         Alignment       =   1  'Right Justify
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
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   3120
         MaxLength       =   2
         TabIndex        =   14
         Tag             =   "nombre"
         Top             =   1560
         Width           =   375
      End
      Begin VB.TextBox dato6 
         Alignment       =   1  'Right Justify
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
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   2640
         MaxLength       =   2
         TabIndex        =   12
         Tag             =   "nombre"
         Top             =   1560
         Width           =   375
      End
      Begin VB.TextBox dato5 
         Alignment       =   1  'Right Justify
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
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   3600
         MaxLength       =   4
         TabIndex        =   11
         Tag             =   "nombre"
         Top             =   1200
         Width           =   855
      End
      Begin VB.TextBox dato4 
         Alignment       =   1  'Right Justify
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
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   3120
         MaxLength       =   2
         TabIndex        =   10
         Tag             =   "nombre"
         Top             =   1200
         Width           =   375
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
         Left            =   2640
         MaxLength       =   8
         TabIndex        =   6
         Tag             =   "codigo"
         Top             =   480
         Width           =   975
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
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   2640
         MaxLength       =   80
         TabIndex        =   5
         Top             =   840
         Width           =   5295
      End
      Begin VB.TextBox dato3 
         Alignment       =   1  'Right Justify
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
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   2640
         MaxLength       =   2
         TabIndex        =   4
         Tag             =   "nombre"
         Top             =   1200
         Width           =   375
      End
      Begin VB.Label Label19 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Año"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   255
         Left            =   5280
         TabIndex        =   48
         Top             =   480
         Width           =   855
      End
      Begin VB.Label lblaño 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   255
         Left            =   6240
         TabIndex        =   47
         Top             =   480
         Width           =   1695
      End
      Begin VB.Label lbldv 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
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
         Height          =   255
         Left            =   3960
         TabIndex        =   45
         Top             =   4080
         Width           =   255
      End
      Begin VB.Label lblnombreproveedor 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   255
         Left            =   4320
         TabIndex        =   44
         Top             =   4080
         Width           =   3735
      End
      Begin VB.Label lblfamilia 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   255
         Left            =   3960
         TabIndex        =   43
         Top             =   3360
         Width           =   4095
      End
      Begin VB.Label Label17 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Familia"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   255
         Left            =   240
         TabIndex        =   42
         Top             =   3360
         Width           =   2295
      End
      Begin VB.Label Label16 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Factura"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   255
         Left            =   240
         TabIndex        =   41
         Top             =   3720
         Width           =   2295
      End
      Begin VB.Label Label5 
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
         ForeColor       =   &H00400000&
         Height          =   255
         Left            =   240
         TabIndex        =   40
         Top             =   4080
         Width           =   2295
      End
      Begin VB.Label lbldepreciacionejercicio 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
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
         Height          =   255
         Left            =   2640
         TabIndex        =   34
         Top             =   6240
         Width           =   1215
      End
      Begin VB.Label lblvidausadaenejercicio 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
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
         Height          =   255
         Left            =   2640
         TabIndex        =   33
         Top             =   5880
         Width           =   1215
      End
      Begin VB.Label lblsaldovidautil 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
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
         Height          =   255
         Left            =   2640
         TabIndex        =   32
         Top             =   5520
         Width           =   1215
      End
      Begin VB.Label lblrevalorizado 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
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
         Height          =   255
         Left            =   2640
         TabIndex        =   31
         Top             =   5160
         Width           =   1215
      End
      Begin VB.Label lblcorreccion 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
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
         Height          =   255
         Left            =   2640
         TabIndex        =   30
         Top             =   4800
         Width           =   1215
      End
      Begin VB.Label lblfactor 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
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
         Height          =   255
         Left            =   2640
         TabIndex        =   29
         Top             =   4440
         Width           =   1215
      End
      Begin VB.Label Label15 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Depreciacion Ejercicio"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   255
         Left            =   240
         TabIndex        =   28
         Top             =   6240
         Width           =   2295
      End
      Begin VB.Label Label14 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Vida Usada En Ejercicio"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   255
         Left            =   240
         TabIndex        =   27
         Top             =   5880
         Width           =   2295
      End
      Begin VB.Label Label13 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Saldo Vida Util"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   255
         Left            =   240
         TabIndex        =   26
         Top             =   5520
         Width           =   2295
      End
      Begin VB.Label Label12 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Valor Revalorizado"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   255
         Left            =   240
         TabIndex        =   25
         Top             =   5160
         Width           =   2295
      End
      Begin VB.Label Label11 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Correccion Monetaria"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   255
         Left            =   240
         TabIndex        =   24
         Top             =   4800
         Width           =   2295
      End
      Begin VB.Label Label10 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Factor De Correccion"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   255
         Left            =   240
         TabIndex        =   23
         Top             =   4440
         Width           =   2295
      End
      Begin VB.Label Label9 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Valor Neto"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   255
         Left            =   240
         TabIndex        =   22
         Top             =   3000
         Width           =   2295
      End
      Begin VB.Label Label8 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Depreciacion Acumulada"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   255
         Left            =   240
         TabIndex        =   20
         Top             =   2640
         Width           =   2295
      End
      Begin VB.Label Label7 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Valor De Compra"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   255
         Left            =   240
         TabIndex        =   18
         Top             =   2280
         Width           =   2295
      End
      Begin VB.Label Label6 
         BackColor       =   &H00FF8080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Vida Util En Meses"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   255
         Left            =   240
         TabIndex        =   16
         Top             =   1920
         Width           =   2295
      End
      Begin VB.Label lblcrcc 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   255
         Left            =   3600
         TabIndex        =   15
         Top             =   1560
         Width           =   4335
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FF8080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " CRCC"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   1560
         Width           =   2295
      End
      Begin VB.Label Label2 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Fecha Puesta En Marcha"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   1200
         Width           =   2295
      End
      Begin VB.Label Label1 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Codigo Activo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   480
         Width           =   2295
      End
      Begin VB.Label Label3 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Nombre"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   840
         Width           =   2295
      End
   End
   Begin VB.TextBox pivote 
      Height          =   285
      Left            =   0
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   2760
      Visible         =   0   'False
      Width           =   615
   End
   Begin MSAdodcLib.Adodc mcm 
      Height          =   375
      Left            =   480
      Top             =   7440
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
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash opciones 
      Height          =   1215
      Left            =   120
      TabIndex        =   2
      Top             =   6840
      Width           =   6735
      _cx             =   11880
      _cy             =   2143
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
Attribute VB_Name = "activofijotb01"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public saldoglobal As Double
Public saldocuenta As Double

 
Private Sub dato1_GotFocus()
    dato1.text = LEERULTIMOACTIVO
    Call cargatexto(dato1)
    
End Sub
Private Sub dato2_GotFocus()
    Call leer
    Call cargatexto(dato2)
End Sub
Private Sub dato3_GotFocus()
    Call cargatexto(dato3)
End Sub
Private Sub dato4_GotFocus()
    Call cargatexto(dato4)
End Sub
Private Sub dato5_GotFocus()
    Call cargatexto(dato5)
End Sub
Private Sub dato6_GotFocus()
    Call cargatexto(dato6)
End Sub
Private Sub dato7_GotFocus()
    Call cargatexto(dato7)
End Sub
Private Sub dato8_GotFocus()
    Call cargatexto(dato8)
End Sub
Private Sub dato9_GotFocus()
    Call cargatexto(dato9)
End Sub
Private Sub dato10_GotFocus()
    Call cargatexto(dato10)
End Sub
Private Sub dato11_GotFocus()
    Call cargatexto(dato11)
End Sub
Private Sub dato12_GotFocus()
    Call cargatexto(dato9)
End Sub
Private Sub dato13_GotFocus()
    Call cargatexto(dato10)
End Sub
Private Sub dato14_GotFocus()
    Call cargatexto(dato11)
End Sub


Private Sub dato1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 38 Then Unload Me: GoTo no:
    Call flechas(dato1, dato2, KeyCode)
no:
End Sub
Private Sub dato2_KeyDown(KeyCode As Integer, Shift As Integer)
    Call flechas(dato1, dato3, KeyCode)
End Sub
Private Sub dato3_KeyDown(KeyCode As Integer, Shift As Integer)
    Call flechas(dato2, dato4, KeyCode)
End Sub
Private Sub dato4_KeyDown(KeyCode As Integer, Shift As Integer)
    Call flechas(dato3, dato5, KeyCode)
End Sub
Private Sub dato5_KeyDown(KeyCode As Integer, Shift As Integer)
    Call flechas(dato4, dato6, KeyCode)
End Sub
Private Sub dato6_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF2 Then Call ayudacentrocosto(dato6)
    Call flechas(dato5, dato7, KeyCode)
End Sub
Private Sub dato7_KeyDown(KeyCode As Integer, Shift As Integer)
    Call flechas(dato6, dato8, KeyCode)
End Sub
Private Sub dato8_KeyDown(KeyCode As Integer, Shift As Integer)
    Call flechas(dato7, dato9, KeyCode)
End Sub
Private Sub dato9_KeyDown(KeyCode As Integer, Shift As Integer)
    Call flechas(dato8, dato10, KeyCode)
End Sub
Private Sub dato10_KeyDown(KeyCode As Integer, Shift As Integer)
    Call flechas(dato9, dato11, KeyCode)
End Sub
Private Sub dato11_KeyDown(KeyCode As Integer, Shift As Integer)
    Call flechas(dato10, dato12, KeyCode)
End Sub
Private Sub dato12_KeyDown(KeyCode As Integer, Shift As Integer)
     If KeyCode = vbKeyF2 Then Call ayudafamilia(dato12)
    Call flechas(dato11, dato13, KeyCode)
End Sub
Private Sub dato13_KeyDown(KeyCode As Integer, Shift As Integer)
    Call flechas(dato12, dato14, KeyCode)
End Sub
Private Sub dato14_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF2 Then Call ayudactacte(dato14)
    Call flechas(dato13, dato14, KeyCode)
    
End Sub
Sub ayudactacte(ByRef caja As TextBox)
    Dim campos As Variant
    Dim cfijo As Variant
    Dim largo As Variant
    campos = Array("rut", "nombre")
    largo = Array("12n", "40s")
    cfijo = "tipo='" & CUENTAPROVEEDOR & "' and año='" + Format(fechasistema, "yyyy") + "'"
    cabezas = Array("rut", "nombre")
    mensajeAyuda = "Ayuda Proveedores"
    
    Call cargaAyudaT(servidor, basebus, Usuario, password, "cuentascorrientes", pivote, campos, cfijo, largo, 2)

    If Val(pivote.text) = 0 Then caja.SetFocus: GoTo no
     
    caja.text = Mid(pivote.text, 1, 9)
    lbldv.Caption = Mid(pivote.text, 10, 1)
    caja.Enabled = True
    caja.SetFocus

no:

End Sub

Sub ayudafamilia(ByRef caja As TextBox)
    Dim campos As Variant
    Dim cfijo As Variant
    Dim largo As Variant
    campos = Array("codigo", "nombre")
    largo = Array("12n", "40s")
    cfijo = "no"
    cabezas = Array("Codigo", "Nombre")
    mensajeAyuda = "Ayuda Familias"
    
    Call cargaAyudaT(servidor, clientesistema + "conta", Usuario, password, "maestro_familias", pivote, campos, cfijo, largo, 2)

    If Val(pivote.text) = 0 Then caja.SetFocus: GoTo no
     
    caja.text = pivote.text
    caja.Enabled = True
    caja.SetFocus

no:

End Sub


Private Sub Form_Activate()
    If dato7.text = "" Then
        dato1.SetFocus
    Else
        dato8.Enabled = True
        dato8.SetFocus
    End If
End Sub

Private Sub Form_Load()

    
    
    Call Conectar_BD
    Rem Call Funciones_Forms_M_Productos.Conecta_Maestro_Productos
    sc = 0
    opciones.Visible = False
DOCU(1) = "ACTIVO"
DOCU(2) = "PASIVO"
DOCU(3) = "RESULTADO"
CANDO = 3

Rem Call RECUPERAFECHA
Call CARGAPERMISO(Me.Name)
dato1.text = Format(LEERULTIMOACTIVO, "00000000")

End Sub
 
Private Sub dato1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then Unload Me
    snum = 0: KeyAscii = esNumero(KeyAscii)
    If KeyAscii = 13 Then
    Call ceros(dato1)
    
    Call Pregunta(dato1, dato2)
    End If
End Sub

Private Sub dato2_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 And dato2.text <> "" Then: Call Pregunta(dato2, dato3)
End Sub

Private Sub dato3_KeyPress(KeyAscii As Integer)
    snum = 0: KeyAscii = esNumero(KeyAscii)
    If KeyAscii = 13 Then
        Call ceros(dato3)
        If dato3.text = "00" Then dato3.text = Format(fechasistema, "dd")
        Call Pregunta(dato3, dato4)
    End If
End Sub
Private Sub dato4_KeyPress(KeyAscii As Integer)
    snum = 0: KeyAscii = esNumero(KeyAscii)
    If KeyAscii = 13 Then
        Call ceros(dato4)
        If dato4.text = "00" Then dato4.text = Format(fechasistema, "mm")
        Call Pregunta(dato4, dato5)
    End If
End Sub
Private Sub dato5_KeyPress(KeyAscii As Integer)
    snum = 0: KeyAscii = esNumero(KeyAscii)
    If KeyAscii = 13 Then
        Call ceros(dato5)
        If dato5.text = "0000" Then dato5.text = Format(fechasistema, "yyyy")
        If IsDate(dato3.text & "-" & dato4.text & "-" & dato5.text) = True Then
            Call Pregunta(dato5, dato6)
        Else
            MsgBox "FECHA INVALIDA, POR FAVOR VERIFICAR", vbCritical, "ATENCION"
            dato3.text = ""
            dato4.text = ""
            dato5.text = ""
            dato3.SetFocus
        End If
    End If
End Sub
 Private Sub dato6_KeyPress(KeyAscii As Integer)
    snum = 0: KeyAscii = esNumero(KeyAscii)
    If KeyAscii = 13 Then
        Call ceros(dato6)
        Call Pregunta(dato6, dato7)
    End If
End Sub
 Private Sub dato7_KeyPress(KeyAscii As Integer)
    snum = 0: KeyAscii = esNumero(KeyAscii)
    If KeyAscii = 13 Then
        Call ceros(dato7)
        lblcrcc.Caption = leerNOMBREcrcc(dato6.text & dato7.text)
        If lblcrcc.Caption <> "" Then
            Call Pregunta(dato7, dato8)
        Else
            MsgBox "CENTRO DE COSTO NO EXISTE,VERIFIQUE PORFAVOR", vbCritical, "ATENCION"
            dato7.text = ""
            dato6.text = ""
            dato6.SetFocus
        End If
    End If
End Sub
Private Sub dato8_KeyPress(KeyAscii As Integer)
    snum = 0: KeyAscii = esNumero(KeyAscii)
    If KeyAscii = 13 Then
       
        Call Pregunta(dato8, dato9)
    End If
End Sub
 Private Sub dato9_KeyPress(KeyAscii As Integer)
    snum = 0: KeyAscii = esNumero(KeyAscii)
    If KeyAscii = 13 Then
      
        Call Pregunta(dato9, dato10)
    End If
End Sub
Private Sub dato10_KeyPress(KeyAscii As Integer)
    snum = 0: KeyAscii = esNumero(KeyAscii)
    If KeyAscii = 13 Then
        If dato10.text = "" Then dato10.text = "0"
        dato11.text = dato9.text - dato10.text
        Call Pregunta(dato10, dato11)
    End If
End Sub
 Private Sub dato11_KeyPress(KeyAscii As Integer)
    snum = 0: KeyAscii = esNumero(KeyAscii)
    If KeyAscii = 13 Then
        Call Pregunta(dato11, dato12)
    End If
End Sub

Private Sub dato12_KeyPress(KeyAscii As Integer)
    snum = 0: KeyAscii = esNumero(KeyAscii)
    If KeyAscii = 13 Then
        Call ceros(dato12)
        lblfamilia.Caption = leerNombreMayor(dato12.text)
        If lblfamilia.Caption <> "" Then
            Call Pregunta(dato12, dato13)
        Else
            MsgBox "FAMILIA NO EXISTE POR FAVOR REVISAR", vbCritical, "ATENCION"
            dato12.text = ""
            dato12.SetFocus
        End If
    End If
End Sub
Private Sub dato13_KeyPress(KeyAscii As Integer)
    snum = 0: KeyAscii = esNumero(KeyAscii)
    If KeyAscii = 13 Then
        Call ceros(dato13)
        Call Pregunta(dato13, dato14)
    End If
End Sub
 Private Sub dato14_KeyPress(KeyAscii As Integer)
    snum = 0: KeyAscii = esNumero(KeyAscii)
    If KeyAscii = 13 Then
        Call ceros(dato14)
        lbldv.Caption = rut(dato14.text)
        lblnombreproveedor.Caption = LEERNOMBREPROVEEDOR(dato14.text & lbldv.Caption)
'        If lblnombreproveedor.Caption <> "" Then
            Call grabar(dato1.text, dato2.text, dato12.text, dato5.text & "-" & dato4.text & "-" & dato3.text, dato13.text, dato8.text, Val(dato9.text), dato10.text, dato11.text, lblvidausadaenejercicio.Caption, lblcorreccion.Caption, dato6.text & dato7.text, dato14.text & lbldv.Caption)
            Unload Me
'        End If
    End If
End Sub

Sub leer()
    campos(0, 0) = "codigo"
    campos(1, 0) = "nombre"
    campos(2, 0) = "familia"
    campos(3, 0) = "fechapuestaenmarcha"
    campos(4, 0) = "factura"
    campos(5, 0) = "proveedor"
    campos(6, 0) = "vidautil"
    campos(7, 0) = "valorcompra"
    campos(8, 0) = "depreciacion"
    campos(9, 0) = "valorreal"
    campos(10, 0) = "vidausada"
    campos(11, 0) = "correcionmonetaria"
    campos(12, 0) = "crcc"
    campos(13, 0) = "año"
    campos(14, 0) = ""
    
    campos(0, 2) = "activo_fijo_tributario"
    condicion = "codigo= '" + dato1.text + "' "

    op = 5
    sqlconta.response = campos
    Set sqlconta.conexion = db
    Call sqlconta.sqlconta(op, condicion)
    If sqlconta.status = 4 Then lblaño.Caption = Format(fechasistema, "yyyy"): GoTo no:
    carga
    opciones.Visible = True
    disponible (True)
    habilita (True)
    opciones.SetFocus
no:
End Sub
Sub leersiguiente()
    campos(0, 0) = "codigo"
    campos(1, 0) = "nombre"
    campos(2, 0) = "familia"
    campos(3, 0) = "fechapuestaenmarcha"
    campos(4, 0) = "factura"
    campos(5, 0) = "proveedor"
    campos(6, 0) = "vidautil"
    campos(7, 0) = "valorcompra"
    campos(8, 0) = "depreciacion"
    campos(9, 0) = "valorreal"
    campos(10, 0) = "vidausada"
    campos(11, 0) = "correcionmonetaria"
    campos(12, 0) = "crcc"
    campos(13, 0) = ""
    campos(0, 2) = "activo_fijo_tributario"
    condicion = "codigo> '" + dato1.text + "' order by codigo"
    op = 5
    sqlconta.response = campos
    Set sqlconta.conexion = db
    Call sqlconta.sqlconta(op, condicion)
    If sqlconta.status = 4 Then GoTo no:
    carga
    opciones.Visible = True
    disponible (True)
    habilita (True)
    opciones.SetFocus
    
no:
   
    
End Sub
Sub leeranterior()
    campos(0, 0) = "codigo"
    campos(1, 0) = "nombre"
    campos(2, 0) = "familia"
    campos(3, 0) = "fechapuestaenmarcha"
    campos(4, 0) = "factura"
    campos(5, 0) = "proveedor"
    campos(6, 0) = "vidautil"
    campos(7, 0) = "valorcompra"
    campos(8, 0) = "depreciacion"
    campos(9, 0) = "valorreal"
    campos(10, 0) = "vidausada"
    campos(11, 0) = "correcionmonetaria"
    campos(12, 0) = "crcc"
    campos(13, 0) = ""
    campos(0, 2) = "activo_fijo_tributario"
    condicion = "codigo< '" + dato1.text + "'  order by codigo desc"

    op = 5
    sqlconta.response = campos
    Set sqlconta.conexion = db
    Call sqlconta.sqlconta(op, condicion)
   If sqlconta.status = 4 Then GoTo no:
    carga
    opciones.Visible = True
    disponible (True)
    habilita (True)
    opciones.SetFocus
no:
   
    
End Sub

Sub carga()
Dim fecha1  As Date
Dim fecha2 As Date

    habilita (True)
    
    dato1.text = sqlconta.response(0, 3)
    dato2.text = sqlconta.response(1, 3)
    dato12.text = sqlconta.response(2, 3)
    dato3.text = Format(sqlconta.response(3, 3), "dd")
    dato4.text = Format(sqlconta.response(3, 3), "mm")
    dato5.text = Format(sqlconta.response(3, 3), "yyyy")
    dato13.text = sqlconta.response(4, 3)
    dato14.text = Mid(sqlconta.response(5, 3), 1, 9)
    lbldv.Caption = Mid(sqlconta.response(5, 3), 10, 1)
    dato8.text = sqlconta.response(6, 3)
    dato9.text = sqlconta.response(7, 3)
    dato10.text = sqlconta.response(8, 3)
    dato11.text = sqlconta.response(9, 3)
    lblvidausadaenejercicio.Caption = sqlconta.response(10, 3)
    lblcorreccion.Caption = sqlconta.response(11, 3)
    dato6.text = Mid(sqlconta.response(12, 3), 1, 2)
    dato7.text = Mid(sqlconta.response(12, 3), 3, 2)
    lblaño.Caption = sqlconta.response(13, 3)
    
    lblnombreproveedor.Caption = LEERNOMBREPROVEEDOR(dato14.text & lbldv.Caption)
    lblfamilia.Caption = leerNombreMayor(dato12.text)
    lblcrcc.Caption = leerNOMBREcrcc(dato6.text & dato7.text)
    
    If dato5.text < Format(fechasistema, "yyyy") Then
        lblfactor.Caption = Format(leeripc("00", Format(fechasistema, "yyyy")), "% ###,##0.0")
    Else
        lblfactor.Caption = Format(leeripc(dato4.text, dato5.text), "% ###,##0.0")
    End If
    

    
            lblcorreccion.Caption = Format((CDbl(dato11.text) * CDbl(Replace(lblfactor.Caption, "%", ""))) / 100, "$ ###,###,##0")
            lblrevalorizado.Caption = Format(CDbl(dato11.text) + CDbl(lblcorreccion.Caption), "$ ###,###,##0")
            
            fecha1 = dato3.text & "-" & dato4.text & "-" & dato5.text
            fecha2 = fechasistema
            
            lblvidausadaenejercicio.Caption = DateDiff("m", fecha1, fecha2)
            
            If lblvidausadaenejercicio.Caption > "12" Then lblvidausadaenejercicio.Caption = "12"
            lblsaldovidautil.Caption = CDbl(dato8.text) - CDbl(lblvidausadaenejercicio.Caption)
            
            lbldepreciacionejercicio.Caption = Format(Round((CDbl(dato11.text) / CDbl(dato8.text)) * CDbl(lblvidausadaenejercicio.Caption), 2), "$ ###,###,##0")

fin:
End Sub

Sub habilita(ByVal condicion As Boolean)
    
    dato1.Locked = condicion
    dato2.Locked = condicion
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
End Sub
Sub disponible(ByVal condicion As Boolean)
    
    dato1.Enabled = condicion
    dato2.Enabled = condicion
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
    
End Sub


Sub Pregunta(ByRef caja As TextBox, ByRef caja1 As TextBox)
    If caja.text = "" And sc = 0 Then caja.SetFocus
    If caja.text <> "" Or sc = 1 Then caja1.Enabled = True: caja1.SetFocus
    sc = 0
End Sub


Sub ayudacentrocosto(ByRef caja As TextBox)
    Dim campos As Variant
    Dim cfijo As Variant
    Dim largo As Variant
    campos = Array("codigo", "nombre")
    largo = Array("8s", "40s")
    cfijo = "año='" + Format(fechasistema, "yyyy") + "'"
    cabezas = Array("codigo", "nombre")
    mensajeAyuda = "Ayuda Centros de Costo"

    Call cargaAyudaT(servidor, basebus, Usuario, password, "centrosdecosto", pivote, campos, cfijo, largo, 2)
    If Val(pivote.text) = 0 Then caja.SetFocus: GoTo no
    dato7.Enabled = True
    dato6.text = Mid(pivote.text, 1, 2)
    dato7.text = Mid(pivote.text, 3, 2)
    
    caja.Enabled = True
    caja.SetFocus
no:
End Sub


Sub flechas(ByRef caja As TextBox, ByRef caja1 As TextBox, ByRef codigo As Integer)
    If codigo = 38 And caja.Enabled = True Then caja.SetFocus
    If codigo = 40 And caja1.Enabled = True Then caja1.SetFocus
End Sub
Sub grabar(codigo, NOMBRE, familia, fechapuestaenmarcha, factura, vidautil, valorcompra, depreciacion, valorreal, vidausada, correcionmonetaria, CRCC, proveedor)
        
    campos(0, 0) = "codigo"
    campos(1, 0) = "nombre"
    campos(2, 0) = "familia"
    campos(3, 0) = "fechapuestaenmarcha"
    campos(4, 0) = "factura"
    campos(5, 0) = "proveedor"
    campos(6, 0) = "vidautil"
    campos(7, 0) = "valorcompra"
    campos(8, 0) = "depreciacion"
    campos(9, 0) = "valorreal"
    campos(10, 0) = "vidausada"
    campos(11, 0) = "correcionmonetaria"
    campos(12, 0) = "crcc"
    campos(13, 0) = "año"
    campos(14, 0) = ""
    
    
    
    campos(0, 1) = codigo
    campos(1, 1) = NOMBRE
    campos(2, 1) = familia
    campos(3, 1) = fechapuestaenmarcha
    campos(4, 1) = factura
    campos(5, 1) = proveedor
    campos(6, 1) = vidautil
    campos(7, 1) = valorcompra
    campos(8, 1) = depreciacion
    campos(9, 1) = valorreal
    campos(10, 1) = vidausada
    campos(11, 1) = correcionmonetaria
    campos(12, 1) = CRCC
    campos(13, 1) = lblaño.Caption
    
    
    campos(0, 2) = "activo_fijo_tributario"
    If MODIFI = 1 Then condicion = "codigo='" & dato1.text & "'"
    If MODIFI = 1 Then op = 3 Else op = 2
    
    sqlconta.response = campos
    Set sqlconta.conexion = db
    Call sqlconta.sqlconta(op, condicion)
    MODIFI = 0
End Sub
 

Sub ELIMINAR()
    campos(0, 2) = "activo_fijo_tributario"
    condicion = "codigo= '" + dato1.text + "' "
    op = 4
    sqlconta.response = campos
    Set sqlconta.conexion = db
    Call sqlconta.sqlconta(op, condicion)
    
End Sub


Private Sub Label18_Click()

End Sub

Private Sub lblhistorico_Click(Index As Integer)

End Sub

Private Sub Frame2_DragDrop(Source As CONTROL, x As Single, Y As Single)

End Sub

 

Private Sub Form_Unload(Cancel As Integer)
    ingreso02.Grid1.Cell(ingreso02.Grid1.ActiveCell.row, 14).text = dato1.text
End Sub

Private Sub opciones_FSCommand(ByVal command As String, ByVal args As String)

If command = "retorno" Then retorno
If command = "modifica" Then disponible (True): habilita (False): dato1.Enabled = False: dato2.Enabled = False: dato3.SetFocus: MODIFI = 1
If command = "elimina" Then ELIMINAR: retorno
If command = "siguiente" Then leersiguiente
If command = "anterior" Then leeranterior
If command = "imprime" Then imprimir
 


End Sub
Sub retorno()
disponible (True)
habilita (False)
limpia
opciones.Visible = False
dato1.Enabled = True
dato1.SetFocus
 
End Sub
Sub limpia()
    dato1.text = ""
    dato2.text = ""
    dato3.text = ""
    dato4.text = ""
    dato5.text = ""
    dato6.text = ""
    dato7.text = ""
    dato8.text = ""
    dato9.text = ""
    dato10.text = ""
    dato11.text = ""
    dato12.text = ""
    dato13.text = ""
    dato14.text = ""
    lblnombreproveedor.Caption = ""
    lblfamilia.Caption = ""
    lblcrcc.Caption = ""
    lbldv.Caption = ""
    lblfactor.Caption = ""
    lblcorreccion.Caption = ""
    lblrevalorizado.Caption = ""
    lblsaldovidautil.Caption = ""
    lblvidausadaenejercicio.Caption = ""
    lbldepreciacionejercicio.Caption = ""
    lblaño.Caption = Format(fechasistema, "yyyy")
    
    
End Sub

Sub imprimir()
     
End Sub
Sub cabeza()
    
End Sub

Public Sub cargadeafueraactivofijo()
    Call dato1_KeyPress(13)
End Sub
Function LEERULTIMOACTIVO() As String
    Dim csql As New rdoQuery
    Dim resultados As rdoResultset
    
    Set csql.ActiveConnection = db
    csql.sql = "select ifnull(max(codigo)+1,1) from "
    csql.sql = csql.sql & "activo_fijo_tributario"
    csql.Execute
    LEERULTIMOACTIVO = "00000001"
    If csql.RowsAffected > 0 Then
        Set resultados = csql.OpenResultset
        LEERULTIMOACTIVO = String(8 - Len(resultados(0)), "0") & resultados(0)
    End If
    csql.Close
    Set csql = Nothing
    
    
End Function
Private Sub botonmisaccesos_Click()
programafiltro = Me.Caption
misaccesos.Show
End Sub


Private Sub botonmisfavoritos_Click()
Call AgregaFavoritos(USUARIOSISTEMA, App.EXEName, Me.Name, Me.Caption)

End Sub
