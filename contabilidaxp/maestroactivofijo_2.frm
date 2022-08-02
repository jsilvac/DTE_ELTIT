VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "msadodc.ocx"
Object = "{4777436C-EB8C-4596-98A8-EBCF98683969}#1.0#0"; "FlexCell.ocx"
Object = "{ADBBAED6-B16F-40DC-80DF-B44910CBA76C}#1.0#0"; "Frame-Xp.ocx"
Begin VB.Form activo02 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Maestro de Activos Fijo "
   ClientHeight    =   9300
   ClientLeft      =   2235
   ClientTop       =   1305
   ClientWidth     =   8415
   DrawStyle       =   5  'Transparent
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   620
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   561
   Begin FlexCell.Grid Grid1 
      Height          =   1455
      Left            =   10440
      TabIndex        =   52
      Top             =   5280
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   2566
      Cols            =   5
      DefaultFontSize =   8.25
      Rows            =   30
   End
   Begin VB.PictureBox CmdFavoritos 
      BackColor       =   &H0000FF00&
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
      Left            =   11040
      ScaleHeight     =   195
      ScaleWidth      =   2715
      TabIndex        =   49
      Top             =   9360
      Width           =   2775
   End
   Begin FlexCell.Grid impresion 
      Height          =   255
      Left            =   8280
      TabIndex        =   46
      Top             =   10320
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   450
      Cols            =   5
      DefaultFontSize =   8.25
      Rows            =   30
   End
   Begin XPFrame.FrameXp FrameXp1 
      Height          =   7935
      Left            =   120
      TabIndex        =   3
      Top             =   0
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   13996
      BackColor       =   16761024
      Caption         =   "Datos Del Activo"
      CaptionEstilo3D =   2
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
      Begin VB.CheckBox Check1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         Caption         =   "VENDER ACTIVO"
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   3960
         TabIndex        =   67
         Top             =   7440
         Width           =   2535
      End
      Begin XPFrame.FrameXp FrmVenta 
         Height          =   1575
         Left            =   4200
         TabIndex        =   55
         Top             =   4800
         Width           =   3735
         _ExtentX        =   6588
         _ExtentY        =   2778
         BackColor       =   16744576
         Caption         =   "VENDER ACTIVO"
         CaptionEstilo3D =   2
         BackColor       =   16744576
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
         Begin VB.CommandButton CmdGrabar2 
            BackColor       =   &H0000C000&
            Caption         =   "GRABAR"
            Height          =   375
            Left            =   1800
            Style           =   1  'Graphical
            TabIndex        =   66
            Top             =   1080
            Width           =   1815
         End
         Begin VB.TextBox dato18 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
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
            Left            =   1800
            TabIndex        =   64
            Tag             =   "nombre"
            Top             =   720
            Width           =   1815
         End
         Begin VB.TextBox dato17 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
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
            Left            =   2760
            MaxLength       =   4
            TabIndex        =   58
            Tag             =   "nombre"
            Top             =   360
            Width           =   855
         End
         Begin VB.TextBox dato16 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
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
            Left            =   2280
            MaxLength       =   2
            TabIndex        =   57
            Tag             =   "nombre"
            Top             =   360
            Width           =   375
         End
         Begin VB.TextBox dato15 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
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
            Left            =   1800
            MaxLength       =   2
            TabIndex        =   56
            Tag             =   "nombre"
            Top             =   360
            Width           =   375
         End
         Begin VB.Label Label23 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " Monto de Venta"
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
            Left            =   120
            TabIndex        =   65
            Top             =   720
            Width           =   1575
         End
         Begin VB.Label Label20 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " Fecha de Venta"
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
            Left            =   120
            TabIndex        =   59
            Top             =   360
            Width           =   1575
         End
      End
      Begin VB.TextBox txtcredito 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         TabIndex        =   53
         Tag             =   "nombre"
         Top             =   2280
         Width           =   1215
      End
      Begin VB.CommandButton CmdGrabar 
         BackColor       =   &H0000C000&
         Caption         =   "GRABAR"
         Height          =   375
         Left            =   2640
         Style           =   1  'Graphical
         TabIndex        =   51
         Top             =   7440
         Width           =   1215
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H000080FF&
         Caption         =   "VER FACTURA"
         Height          =   255
         Left            =   3960
         Style           =   1  'Graphical
         TabIndex        =   50
         Top             =   3720
         Width           =   1695
      End
      Begin VB.TextBox dato12 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         TabIndex        =   39
         Tag             =   "nombre"
         Top             =   3360
         Width           =   375
      End
      Begin VB.TextBox dato13 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Appearance      =   0  'Flat
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
         Left            =   9240
         TabIndex        =   35
         Top             =   360
         Visible         =   0   'False
         Width           =   5535
         _ExtentX        =   9763
         _ExtentY        =   9340
         BackColor       =   16744576
         Caption         =   "Datos"
         CaptionEstilo3D =   1
         BackColor       =   16744576
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Begin FlexCell.Grid Grid1_ 
            Height          =   4815
            Left            =   1
            TabIndex        =   36
            Top             =   1
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
         Appearance      =   0  'Flat
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
         Appearance      =   0  'Flat
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
         Appearance      =   0  'Flat
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
         Top             =   1920
         Width           =   1215
      End
      Begin VB.TextBox dato8 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Top             =   1560
         Width           =   495
      End
      Begin VB.TextBox dato7 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Left            =   3480
         MaxLength       =   4
         TabIndex        =   14
         Tag             =   "nombre"
         Top             =   9480
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox dato6 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Left            =   3000
         MaxLength       =   2
         TabIndex        =   12
         Tag             =   "nombre"
         Top             =   9480
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox dato5 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Appearance      =   0  'Flat
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
         Left            =   2640
         MaxLength       =   8
         TabIndex        =   6
         Tag             =   "codigo"
         Top             =   480
         Width           =   975
      End
      Begin VB.TextBox dato2 
         Appearance      =   0  'Flat
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
         Appearance      =   0  'Flat
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
      Begin VB.Label LblMontoVenta 
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
         TabIndex        =   63
         Top             =   6960
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label lblFechaVenta 
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
         TabIndex        =   62
         Top             =   6600
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label Label22 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Monto de La Venta"
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
         Left            =   240
         TabIndex        =   61
         Top             =   6960
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.Label Label21 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Fecha de Venta"
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
         Left            =   240
         TabIndex        =   60
         Top             =   6600
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.Label Label18 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Credito 4%"
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
         Left            =   240
         TabIndex        =   54
         Top             =   2280
         Width           =   2295
      End
      Begin VB.Label Label19 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
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
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   5280
         TabIndex        =   48
         Top             =   480
         Width           =   855
      End
      Begin VB.Label lblaño 
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
         Left            =   4320
         TabIndex        =   44
         Top             =   4080
         Width           =   3735
      End
      Begin VB.Label lblfamilia 
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
         Left            =   3120
         TabIndex        =   43
         Top             =   3360
         Width           =   4935
      End
      Begin VB.Label Label17 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
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
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   240
         TabIndex        =   42
         Top             =   3360
         Width           =   2295
      End
      Begin VB.Label Label16 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
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
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   240
         TabIndex        =   41
         Top             =   3720
         Width           =   2295
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
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
         ForeColor       =   &H80000008&
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
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
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
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   240
         TabIndex        =   28
         Top             =   6240
         Width           =   2295
      End
      Begin VB.Label Label14 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
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
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   240
         TabIndex        =   27
         Top             =   5880
         Width           =   2295
      End
      Begin VB.Label Label13 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
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
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   240
         TabIndex        =   26
         Top             =   5520
         Width           =   2295
      End
      Begin VB.Label Label12 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Valor Final"
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
         Left            =   240
         TabIndex        =   25
         Top             =   5160
         Width           =   2295
      End
      Begin VB.Label Label11 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
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
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   240
         TabIndex        =   24
         Top             =   4800
         Width           =   2295
      End
      Begin VB.Label Label10 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
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
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   240
         TabIndex        =   23
         Top             =   4440
         Width           =   2295
      End
      Begin VB.Label Label9 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Valor Inicial"
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
         Left            =   240
         TabIndex        =   22
         Top             =   3000
         Width           =   2295
      End
      Begin VB.Label Label8 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
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
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   240
         TabIndex        =   20
         Top             =   2640
         Width           =   2295
      End
      Begin VB.Label Label7 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Valor Libro"
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
         Left            =   240
         TabIndex        =   18
         Top             =   1920
         Width           =   2295
      End
      Begin VB.Label Label6 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
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
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   240
         TabIndex        =   16
         Top             =   1560
         Width           =   2295
      End
      Begin VB.Label lblcrcc 
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
         Left            =   4200
         TabIndex        =   15
         Top             =   9480
         Visible         =   0   'False
         Width           =   4095
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
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
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   600
         TabIndex        =   13
         Top             =   9480
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
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
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   1200
         Width           =   2295
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
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
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   480
         Width           =   2295
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
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
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   840
         Width           =   2295
      End
   End
   Begin VB.TextBox pivote 
      Height          =   285
      Left            =   9720
      TabIndex        =   1
      Top             =   9240
      Visible         =   0   'False
      Width           =   615
   End
   Begin MSAdodcLib.Adodc mcm 
      Height          =   375
      Left            =   7920
      Top             =   9120
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
      Top             =   8040
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
Attribute VB_Name = "activo02"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public saldoglobal As Double
Public saldocuenta As Double
Public depreciacionactual As Double
Public valorlibrofinal As Double
 
Private Sub Check1_Click()

If ExistemMovimientos(dato1, Format(fechasistema, "yyyy")) = False Then
    If Verifica_Permiso(Me.Caption, "autoriza") = True Then
        dato15.Enabled = True
        FrmVenta.Visible = True
        CmdGrabar.Visible = False
        Call Pregunta(dato14, dato15)
    Else
        MsgBox "USTED NO TIENE PRIVILEGIOS PARA MARCAR COMO VENDIDO UN ACTIVO" & vbNewLine & "SOLICITE ACCESO " & Chr(34) & "AUTORIZA" & Chr(34) & " Y VUELVA A INTENTAR"
        Exit Sub
    End If
        
Else
    MsgBox "NO PUEDE MODIFICAR DATOS SI HAY MOVIMIENTOS DE ESTE ACTIVO EN AÑOS SUPERIORES O INFERIORES AL ACTUAL", vbCritical
    Exit Sub
End If

End Sub

Private Sub CmdFavoritos_Click()
    Call AgregaFavoritos(USUARIOSISTEMA, App.EXEName, Me.Name, Me.Caption)
End Sub

Private Sub cmdgrabar_Click()
        If dato1.text = "" Then Exit Sub
        If dato2.text = "" Then Exit Sub
        If IsDate(dato3.text & "-" & dato4.text & "-" & DATO5.text) = False Then Exit Sub
        If Val(dato9) = 0 Then Exit Sub
        If dato10.text = "" Then Exit Sub
        If Val(dato11) = 0 Then dato11.SetFocus: Exit Sub
        If lblfamilia.Caption = "" Then Exit Sub
        If lblnombreproveedor.Caption = "" Then Exit Sub
        Call grabar(dato1.text, dato2.text, dato12.text, DATO5.text & "-" & dato4.text & "-" & dato3.text, dato13.text, dato8.text, Val(dato9.text), dato10.text, dato11.text, lblvidausadaenejercicio.Caption, lblcorreccion.Caption, dato6.text & dato7.text, dato14.text & LBLDV.Caption)
        Call dato1_KeyPress(13)
End Sub

Private Sub CmdGrabar2_Click()
        MODIFI = 1
        If dato1.text = "" Then Exit Sub
        If dato2.text = "" Then Exit Sub
        If IsDate(dato3.text & "-" & dato4.text & "-" & DATO5.text) = False Then Exit Sub
        If Val(dato9) = 0 Then Exit Sub
        If dato10.text = "" Then Exit Sub
        If Val(dato11) = 0 Then Exit Sub
        If lblfamilia.Caption = "" Then Exit Sub
        If lblnombreproveedor.Caption = "" Then Exit Sub
        Call grabar(dato1.text, dato2.text, dato12.text, DATO5.text & "-" & dato4.text & "-" & dato3.text, dato13.text, dato8.text, Val(dato9.text), dato10.text, dato11.text, lblvidausadaenejercicio.Caption, lblcorreccion.Caption, dato6.text & dato7.text, dato14.text & LBLDV.Caption)
        MODIFI = 0
        Call dato1_KeyPress(13)
End Sub

Private Sub Command1_Click()
Dim n As Double
Dim nombrefichero As String
Dim rutprove As String
Dim NUMERODOC As String
Dim tipodoc As String
Dim rutafinal As String
 
        NUMERODOC = dato13.text
        tipodoc = "1"
        rutprove = dato14.text
        nombrefichero = "TD" & tipodoc & NUMERODOC & ".JPG"
        rutafinal = "\\" & Servidor & "\" & RutaArchivos & "fc_" & empresaactiva & "\" & rutprove & "\" & nombrefichero
            If tipodoc <> "" Then
        If ExisteArchivo(rutafinal) = True Then
            ShellExecute Me.hwnd, "open", rutafinal, "", "", 4
        Else
        tipodoc = "4"
        nombrefichero = "TD" & tipodoc & NUMERODOC & ".JPG"
        rutafinal = "\\" & Servidor & "\" & RutaArchivos & "fc_" & empresaactiva & "\" & rutprove & "\" & nombrefichero
            If ExisteArchivo(rutafinal) = True Then
               ShellExecute Me.hwnd, "open", rutafinal, "", "", 4
            Else
            MsgBox "DOCUMENTO NO DIGITALIZADO"
            End If
        End If
        End If
 
End Sub

Private Sub dato1_GotFocus()
  If Me.Tag = "" Then
   ' dato1.text = LEERULTIMOACTIVO
  Else
    dato1.text = Me.Tag
    Call dato1_KeyPress(13)
  End If
    Call cargatexto(dato1)
    
End Sub

Private Sub dato12_Change()
lblfamilia.Caption = Empty
End Sub

Private Sub dato14_Change()
lblnombreproveedor.Caption = Empty
End Sub

Private Sub dato15_GotFocus()
Call cargatexto(dato15)
End Sub

Private Sub dato15_KeyPress(KeyAscii As Integer)
    snum = 0: KeyAscii = esNumero(KeyAscii)
    If KeyAscii = 13 Then
'        If dato15.text = "" Then
'
'
''            If MsgBox("DEJARA SIN FECHA DE VENTA?", vbYesNo, "A T E N C I O N") = vbYes Then
''                    dato15.text = "00"
''                    dato16.text = "00"
''                    dato17.text = "0000"
''                    CmdGrabar.SetFocus
''                    Exit Sub
''            End If
'        End If
    
        If dato15.text = "" Then
            dato15.text = Format(fechasistema, "dd")
        Else

        End If
    
        Call ceros(dato15)
        Call Pregunta(dato15, dato16)
    End If
End Sub

Private Sub DATO16_GotFocus()
Call cargatexto(dato16)
End Sub

Private Sub dato16_KeyPress(KeyAscii As Integer)
    snum = 0: KeyAscii = esNumero(KeyAscii)
    If KeyAscii = 13 Then
        If dato16.text = "" Then
            dato16.text = Format(fechasistema, "mm")
        End If

        Call ceros(dato16)
        Call Pregunta(dato16, dato17)
    End If
End Sub

Private Sub dato17_GotFocus()
Call cargatexto(dato17)
End Sub

Private Sub dato17_KeyPress(KeyAscii As Integer)
    snum = 0: KeyAscii = esNumero(KeyAscii)
    If KeyAscii = 13 Then
        If dato17.text = "" Then
            dato17.text = Format(fechasistema, "yyyy")
        End If
        Call ceros(dato17)
        Call Pregunta(dato17, dato18)
    End If
End Sub

Private Sub dato18_KeyPress(KeyAscii As Integer)
    snum = 0: KeyAscii = esNumero(KeyAscii)
    If KeyAscii = 13 Then
        If dato18.text = "" Then
            dato18.text = 0
        End If
        
        CmdGrabar2.SetFocus
    End If
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
    Call cargatexto(DATO5)
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
If KeyCode = vbKeyF2 Then Call AyudaActivoFijo(dato1)
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
    Call flechas(dato3, DATO5, KeyCode)
End Sub
Private Sub dato5_KeyDown(KeyCode As Integer, Shift As Integer)
    Call flechas(dato4, dato8, KeyCode)
End Sub
Private Sub dato6_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF2 Then Call ayudacentrocosto(dato6)
    Call flechas(DATO5, dato7, KeyCode)
End Sub
Private Sub dato7_KeyDown(KeyCode As Integer, Shift As Integer)
    Call flechas(dato6, dato8, KeyCode)
End Sub
Private Sub dato8_KeyDown(KeyCode As Integer, Shift As Integer)
    Call flechas(DATO5, dato9, KeyCode)
End Sub
Private Sub dato9_KeyDown(KeyCode As Integer, Shift As Integer)
    Call flechas(dato8, txtcredito, KeyCode)
End Sub
Private Sub dato10_KeyDown(KeyCode As Integer, Shift As Integer)
    Call flechas(txtcredito, dato11, KeyCode)
End Sub
Private Sub dato11_KeyDown(KeyCode As Integer, Shift As Integer)
    Call flechas(dato10, dato11, KeyCode)
End Sub
Private Sub dato12_KeyDown(KeyCode As Integer, Shift As Integer)
     If KeyCode = vbKeyF2 Then Call ayudafamilia(dato12)
    Call flechas(txtcredito, dato13, KeyCode)
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
    
    Call cargaAyudaT(Servidor, basebus, Usuario, password, "cuentascorrientes", pivote, campos, cfijo, largo, 2)

    If Val(pivote.text) = 0 Then caja.SetFocus: GoTo no
     
    caja.text = Mid(pivote.text, 1, 9)
    LBLDV.Caption = Mid(pivote.text, 10, 1)
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
    
    Call cargaAyudaT(Servidor, clientesistema + "conta", Usuario, password, "maestro_familias_nuevo", pivote, campos, cfijo, largo, 2)

    If Val(pivote.text) = 0 Then caja.SetFocus: GoTo no
     
    caja.text = pivote.text
    caja.Enabled = True
    caja.SetFocus

no:

End Sub


Private Sub dato9_LostFocus()
            Dim VALORCREDITO As Double
            If Val(dato9.text) > 0 Then
            VALORCREDITO = Format((Round(dato9 * 0.04)), "###,###,##0")
            If Val(txtcredito) <> VALORCREDITO Then
                If MsgBox("DESEA CALCULAR CREDITO 4%?", vbYesNo, "A T E N C I O N ** CREDITO " & Format((Round(dato9 * 0.04)), "###,###,##0")) = vbYes Then
                    
                    
                    txtcredito = VALORCREDITO
                Else
                    txtcredito.text = 0
                    Call Pregunta(dato9, txtcredito)
                End If
            End If
        End If
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
'dato1.text = Format(LEERULTIMOACTIVO, "00000000")
Rem Call RECUPERAFECHA
Call CARGAPERMISO(Me.Name)
If Me.Tag = "" Then
    dato1.text = Format(LEERULTIMOACTIVO, "00000000")
Else
    dato1.text = Me.Tag
End If
Call CENTRAR(Me)
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
        Call Pregunta(dato4, DATO5)
    End If
End Sub
Private Sub dato5_KeyPress(KeyAscii As Integer)
    snum = 0: KeyAscii = esNumero(KeyAscii)
    If KeyAscii = 13 Then
        Call ceros(DATO5)
        If DATO5.text = "0000" Then DATO5.text = Format(fechasistema, "yyyy")
        If IsDate(dato3.text & "-" & dato4.text & "-" & DATO5.text) = True Then
            Call Pregunta(DATO5, dato8)
        Else
            MsgBox "FECHA INVALIDA, POR FAVOR VERIFICAR", vbCritical, "ATENCION"
            dato3.text = ""
            dato4.text = ""
            DATO5.text = ""
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
      
        
        Call Pregunta(dato9, txtcredito)
        
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
        lblfamilia.Caption = LEERNOMBREFAMILIA(dato12.text)
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
        LBLDV.Caption = rut(dato14.text)
        lblnombreproveedor.Caption = LEERNOMBREPROVEEDOR(dato14.text & LBLDV.Caption)
        If lblnombreproveedor.Caption <> "" Then
                If CmdGrabar.Visible = False Then CmdGrabar.Visible = True
              CmdGrabar.SetFocus
          '  Call Pregunta(dato13, dato15)
        End If
    End If
End Sub

Sub leer()
Dim NOMBREACTIVO As String
NOMBREACTIVO = LeerNombreActivo(dato1)
If NOMBREACTIVO <> "" Then
    If LeerActivoFijoAño(dato1) = False Then
    MsgBox "ACTIVO " & NOMBREACTIVO & vbNewLine & " NO EXISTE PARA ESTE PERIODO" & vbNewLine & " DEBE HACER CIERRE U O UTLIZAR NUEVO CODIGO PARA NUEVO ACTIVO"
    dato1.text = Format(LEERULTIMOACTIVO, "00000000")
    dato1.SetFocus
    Exit Sub
    End If
End If

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
    campos(14, 0) = "valorcredito"
    campos(15, 0) = "ifnull(fechaventa,'')"
    campos(16, 0) = "montoventa"
    campos(17, 0) = ""
    
    campos(0, 2) = "activo_fijo_nuevo"
    condicion = "codigo= '" + dato1.text + "' and año='" & Format(fechasistema, "yyyy") & "' "

    op = 5
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
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
    campos(13, 0) = "año"
    campos(14, 0) = "valorcredito"
    campos(15, 0) = "ifnull(fechaventa,'')"
    campos(16, 0) = "montoventa"
    campos(17, 0) = ""
    
    campos(0, 2) = "activo_fijo_nuevo"
    condicion = "año='" & Format(fechasistema, "yyyy") & "' and codigo> '" + dato1.text + "' order by codigo"
    op = 5
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
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
    campos(13, 0) = "año"
    campos(14, 0) = "valorcredito"
    campos(15, 0) = "ifnull(fechaventa,'')"
    campos(16, 0) = "montoventa"
    campos(17, 0) = ""
    
    campos(0, 2) = "activo_fijo_nuevo"
    condicion = "año='" & Format(fechasistema, "yyyy") & "' and codigo< '" + dato1.text + "'  order by codigo desc"

    op = 5
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
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
    Check1.Value = 0
    dato1.text = sqlconta.response(0, 3)
    dato2.text = sqlconta.response(1, 3)
    dato12.text = sqlconta.response(2, 3)
    dato3.text = Format(sqlconta.response(3, 3), "dd")
    dato4.text = Format(sqlconta.response(3, 3), "mm")
    DATO5.text = Format(sqlconta.response(3, 3), "yyyy")
    dato13.text = sqlconta.response(4, 3)
    dato14.text = Mid(sqlconta.response(5, 3), 1, 9)
    LBLDV.Caption = Mid(sqlconta.response(5, 3), 10, 1)
    dato8.text = sqlconta.response(6, 3)
    
    
      Call LeerActivoFijo(dato1, dato12)
'    dato10.text = sqlconta.response(8, 3) '+ leercorreccionanterior(dato1.text)
    dato10.text = depreciacionactual
'    dato11.text = sqlconta.response(7, 3)
    dato11.text = valorlibrofinal
 
'    lblcorreccion.Caption = sqlconta.response(11, 3)
    dato6.text = Mid(sqlconta.response(12, 3), 1, 2)
    dato7.text = Mid(sqlconta.response(12, 3), 3, 2)
    lblaño.Caption = sqlconta.response(13, 3)
    txtcredito.text = sqlconta.response(14, 3)
    
    If sqlconta.response(3, 3) > fechasistema Then
'        Stop
    End If
    
    If IsDate(Format(sqlconta.response(15, 3), "dd-mm-yyyy")) = True Then
        FrmVenta.Visible = False
        LblMontoVenta.Visible = True
        lblFechaVenta.Visible = True
        Label21.Visible = True
        Label22.Visible = True
        Check1.Caption = "CAMBIAR DATOS DE LA VENTA"
        Check1.Visible = True
        dato15.text = Format(sqlconta.response(15, 3), "dd")
        dato16.text = Format(sqlconta.response(15, 3), "mm")
        dato17.text = Format(sqlconta.response(15, 3), "yyyy")
        dato18.text = sqlconta.response(16, 3)
        
        lblFechaVenta.Caption = Format(sqlconta.response(15, 3), "dd-mm-yyyy")
        LblMontoVenta.Caption = Format(sqlconta.response(16, 3), "$ ###,###,##0")
    Else
        Check1.Visible = True
        Check1.Caption = "VENDER ACTIVO"
        FrmVenta.Visible = False
    End If
    
    
    
    lblnombreproveedor.Caption = LEERNOMBREPROVEEDOR(dato14.text & LBLDV.Caption)
    
    lblfamilia.Caption = leerNombreMayor(dato12.text)
    lblcrcc.Caption = leerNOMBREcrcc(dato6.text & dato7.text)
    lblfamilia.Caption = LEERNOMBREFAMILIA(dato12.text)
    
   
    fecha1 = dato3.text & "-" & dato4.text & "-" & DATO5.text
    fecha2 = fechasistema
    CmdGrabar.Visible = False
  Call LeerActivoFijo(dato1, dato12)
 Rem Call LEERactivofijos2(dato1, dato12)
fin:
End Sub

Sub habilita(ByVal condicion As Boolean)
    
    dato1.Locked = condicion
    dato2.Locked = condicion
    dato3.Locked = condicion
    dato4.Locked = condicion
    DATO5.Locked = condicion
    dato6.Locked = condicion
    dato7.Locked = condicion
    dato8.Locked = condicion
    dato9.Locked = condicion
    dato10.Locked = condicion
    dato11.Locked = condicion
    dato12.Locked = condicion
    dato13.Locked = condicion
    dato14.Locked = condicion
    txtcredito.Locked = condicion
End Sub
Sub disponible(ByVal condicion As Boolean)
    
    dato1.Enabled = condicion
    dato2.Enabled = condicion
    dato3.Enabled = condicion
    dato4.Enabled = condicion
    DATO5.Enabled = condicion
    dato6.Enabled = condicion
    dato7.Enabled = condicion
    dato8.Enabled = condicion
    dato9.Enabled = condicion
    dato10.Enabled = condicion
    dato11.Enabled = condicion
    dato12.Enabled = condicion
    dato13.Enabled = condicion
    dato14.Enabled = condicion
    txtcredito.Enabled = condicion
    
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

    Call cargaAyudaT(Servidor, basebus, Usuario, password, "centrosdecosto", pivote, campos, cfijo, largo, 2)
    If Val(pivote.text) = 0 Then caja.SetFocus: GoTo no
    dato7.Enabled = True
    dato6.text = Mid(pivote.text, 1, 2)
    dato7.text = Mid(pivote.text, 3, 4)
    
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
    campos(14, 0) = "valorcredito"
    campos(15, 0) = "fechaventa"
    campos(16, 0) = "montoventa"
    campos(17, 0) = ""
    
    
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
    campos(14, 1) = txtcredito.text
    
    campos(15, 1) = Format(dato15.text & "-" & dato16.text & "-" & dato17.text, "yyyy-mm-dd")
    campos(16, 1) = dato18.text
    
    
    campos(0, 2) = "activo_fijo_nuevo"
    If MODIFI = 1 Then condicion = "codigo='" & dato1.text & "' and año='" & Format(fechasistema, "yyyy") & "' "
    If MODIFI = 1 Then op = 3 Else op = 2
    
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)
    MODIFI = 0
End Sub
 

Sub ELIMINAR()
    campos(0, 2) = "activo_fijo_nuevo"
    condicion = "codigo= '" + dato1.text + "' and año='" & Format(fechasistema, "yyyy") & "' "
    op = 4
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)
    
End Sub


Private Sub lblhistorico_Click(Index As Integer)

End Sub

Private Sub Frame2_DragDrop(Source As CONTROL, X As Single, Y As Single)

End Sub

 

Private Sub Form_Unload(Cancel As Integer)
'    ingreso02.Grid1.Cell(ingreso02.Grid1.ActiveCell.row, 14).text = dato1.text
If opciones.Visible = False And (dato1.text <> "" And dato2.text <> "") Or (MODIFI = 1) Then
    If MsgBox("SI CIERRA PERDERA LOS CAMBIOS" & vbNewLine & "        DESEA CONTINUAR?", vbYesNo, "ATENCION") = vbYes Then
    Cancel = 0
    Else
    Cancel = 1
    End If
End If
End Sub

Private Sub opciones_FSCommand(ByVal command As String, ByVal args As String)

If command = "retorno" Then retorno
If command = "modifica" Then
    If Verifica_Permiso(Me.Caption, "modifica") = True Then
        If ExistemMovimientos(dato1, Format(fechasistema, "yyyy")) = False Then
            disponible (True)
            habilita (False)
            dato1.Enabled = False
            dato2.Enabled = False
            dato3.SetFocus
            MODIFI = 1
            CmdGrabar.Visible = True
        Else
            MsgBox "NO PUEDE MODIFICAR DATOS SI HAY MOVIMIENTOS DE ESTE ACTIVO EN AÑOS SUPERIORES O INFERIORES AL ACTUAL", vbCritical
            Exit Sub
        End If
    Else
        MsgBox mensaje_nopermiso
    End If
End If
If command = "elimina" Then
    If Verifica_Permiso(Me.Caption, "elimina") = True Then
        If ExistemMovimientos2(dato1, Format(fechasistema, "yyyy")) = False Then
            ELIMINAR
            retorno
        Else
            MsgBox "NO PUEDE ELMINAR SI HAY MOVIMIENTOS DE ESTE ACTIVO EN AÑOS SUPERIORES" & vbNewLine & "DEBE BORRAR MOVIMIENTOS SUPERIORES PRIMERO"
        End If
    Else
        MsgBox mensaje_nopermiso
    End If
End If
If command = "siguiente" Then leersiguiente
If command = "anterior" Then leeranterior
If command = "imprime" Then imprimir
 


End Sub
Sub retorno()
Check1.Value = 0
FrmVenta.Visible = False
CmdGrabar.Visible = False
Check1.Caption = "VENDER ACTIVO"
Check1.Visible = False

Label21.Visible = False
Label22.Visible = False
lblFechaVenta.Visible = False
LblMontoVenta.Visible = False


disponible (True)
habilita (False)
Me.Tag = ""
Grid1.Rows = 1
limpia
opciones.Visible = False

 
    dato1.text = Format(LEERULTIMOACTIVO, "00000000")
 
dato1.Enabled = True
dato1.SetFocus
 
End Sub
Sub limpia()
    dato1.text = ""
    dato2.text = ""
    dato3.text = ""
    dato4.text = ""
    DATO5.text = ""
    dato6.text = ""
    dato7.text = ""
    dato8.text = ""
    dato9.text = ""
    dato10.text = ""
    dato11.text = ""
    dato12.text = ""
    dato13.text = ""
    dato14.text = ""
    dato15.text = ""
    dato16.text = ""
    dato17.text = ""
    dato18.text = ""
    
    lblnombreproveedor.Caption = ""
    lblfamilia.Caption = ""
    lblcrcc.Caption = ""
    LBLDV.Caption = ""
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
Sub CABEZA()
    
End Sub

Public Sub cargadeafueraactivofijo()
    Call dato1_KeyPress(13)
End Sub
Function LEERULTIMOACTIVO() As String
    Dim csql As New rdoQuery
    Dim resultados As rdoResultset
    
    Set csql.ActiveConnection = contadb
    csql.sql = "select ifnull(max(codigo)+1,1) from "
    csql.sql = csql.sql & "activo_fijo_nuevo "
    csql.Execute
    LEERULTIMOACTIVO = "00000001"
    If csql.RowsAffected > 0 Then
        Set resultados = csql.OpenResultset
        LEERULTIMOACTIVO = String(8 - Len(resultados(0)), "0") & resultados(0)
    End If
    csql.Close
    Set csql = Nothing
    
    
End Function




Function ExistemMovimientos(codigo, año) As Boolean
    Dim csql As New rdoQuery
    Dim resultado As rdoResultset
    
    Set csql.ActiveConnection = contadb
    csql.sql = "select count(codigo) from " & clientesistema & "conta" & empresaactiva & ".activo_fijo_nuevo"
    csql.sql = csql.sql & " where codigo='" & codigo & "' "
    csql.sql = csql.sql & " and (año >'" & año & "' or año<'" & año & "') "
    csql.Execute
    ExistemMovimientos = False
 If csql.RowsAffected > 0 Then
    Set resultado = csql.OpenResultset
    If resultado(0) > 0 Then
        ExistemMovimientos = True
    Else
        ExistemMovimientos = False
    End If
 End If
 csql.Close
 Set csql = Nothing
 Set resultado = Nothing
 
End Function


Function ExistemMovimientos2(codigo, año) As Boolean
    Dim csql As New rdoQuery
    Dim resultado As rdoResultset
    
    Set csql.ActiveConnection = contadb
    csql.sql = "select count(codigo) from " & clientesistema & "conta" & empresaactiva & ".activo_fijo_nuevo"
    csql.sql = csql.sql & " where codigo='" & codigo & "' "
    csql.sql = csql.sql & " and año >'" & año & "'  "
    csql.Execute
    ExistemMovimientos2 = False
 If csql.RowsAffected > 0 Then
    Set resultado = csql.OpenResultset
    If resultado(0) > 0 Then
        ExistemMovimientos2 = True
    Else
        ExistemMovimientos2 = False
    End If
 End If
 csql.Close
 Set csql = Nothing
 Set resultado = Nothing
 
End Function


 


Function TieneCorreccion(familia) As Boolean
    Dim csql As New rdoQuery
    Dim resultado As rdoResultset
    
    Set csql.ActiveConnection = conta
    csql.sql = "select correccion_monetaria from " & clientesistema & "conta.maestro_familias_nuevo"
    csql.sql = csql.sql & " where codigo='" & familia & "' "
    
    csql.Execute
    TieneCorreccion = False
 If csql.RowsAffected > 0 Then
    Set resultado = csql.OpenResultset
    If resultado(0) > 0 Then
        TieneCorreccion = True
    Else
        TieneCorreccion = False
    End If
 End If
 csql.Close
 Set csql = Nothing
 Set resultado = Nothing
 
End Function




Function TieneDepreciacion(familia) As Boolean
    Dim csql As New rdoQuery
    Dim resultado As rdoResultset
    
    Set csql.ActiveConnection = conta
    csql.sql = "select depreciacion from " & clientesistema & "conta.maestro_familias_nuevo"
    csql.sql = csql.sql & " where codigo='" & familia & "' "
    
    csql.Execute
    TieneDepreciacion = False
 If csql.RowsAffected > 0 Then
    Set resultado = csql.OpenResultset
    If resultado(0) > 0 Then
        TieneDepreciacion = True
    Else
        TieneDepreciacion = False
    End If
 End If
 csql.Close
 Set csql = Nothing
 Set resultado = Nothing
 
End Function


Sub LeerActivoFijo(codigo, familia)
Dim csql As New rdoQuery
Dim resultados As rdoResultset
Dim tipo As String
Dim filtro As String
Dim FILTRO2 As String
Dim totales(10) As Double
Dim totales2(10) As Double
Dim cuentapublicidad As String
Dim factor As Double
Dim CORRECCION As Double
Dim depreciacion As Double
Dim vidaanterior As Double
Dim valorlibro As Double
Dim depreciacionmes As Double
Dim vidaejercicio As Double
Dim libro As Double
'Call CargarFormatoGrilla
Grid1.Cols = 1
 Grid1.Cols = 18
 Grid1.Rows = 1
' For k = 1 To Grid1.Cols - 1
'    Grid1.Column(k).Locked = False
' Next k
    Set csql.ActiveConnection = contadb
'    csql.sql = "select codigo,nombre,crcc,fechapuestaenmarcha,valorcompra"
'    csql.sql = csql.sql & " ,depreciacion+ifnull((SELECT SUM(depreciacion_ejercicio) FROM " & clientesistema & "conta" & empresaactiva & ".activo_fijo_nuevo"
'    csql.sql = csql.sql & " WHERE año <c.año AND codigo=c.codigo),0) AS depreciacion "
'    csql.sql = csql.sql & " ,correcionmonetaria+ifnull((SELECT SUM(correccion_ejercicio) FROM " & clientesistema & "conta" & empresaactiva & ".activo_fijo_nuevo WHERE año <c.año AND codigo=c.codigo),0) AS correccionmonetaria"
'    csql.sql = csql.sql & ",valorcompra-depreciacion -valorcredito,vidautil"
'    csql.sql = csql.sql & ",ifnull((SELECT SUM(vida_ejercicio) FROM " & clientesistema & "conta" & empresaactiva & ".activo_fijo_nuevo WHERE año <c.año AND codigo=c.codigo),0) AS  vidausada,'0',familia"
'
'    csql.sql = csql.sql & ", valorcredito,fechaventa "
'    csql.sql = csql.sql & " from activo_fijo_nuevo  as c where año='" & Format(fechasistema, "yyyy") & "' "
'    If dato3.text <> "" Then
'        csql.sql = csql.sql & " and familia='" & familia & "' and codigo='" & codigo & "' "
'    End If
'
'    csql.sql = csql.sql & "order by familia,fechapuestaenmarcha "
'    csql.Execute
    
    
     csql.sql = "select codigo,nombre,crcc,fechapuestaenmarcha,valorcompra+correcionmonetaria+IFNULL((SELECT SUM(correccion_ejercicio) FROM " & clientesistema & "conta" & empresaactiva & ".activo_fijo_nuevo WHERE año <c.año AND codigo=c.codigo),0)"
    csql.sql = csql.sql & " ,depreciacion+ifnull((SELECT SUM(depreciacion_ejercicio) FROM " & clientesistema & "conta" & empresaactiva & ".activo_fijo_nuevo"
    csql.sql = csql.sql & " WHERE año <c.año AND codigo=c.codigo),0) AS depreciacion "
    csql.sql = csql.sql & " ,correcionmonetaria+ifnull((SELECT SUM(correccion_ejercicio) FROM " & clientesistema & "conta" & empresaactiva & ".activo_fijo_nuevo WHERE año <c.año AND codigo=c.codigo),0) AS correccionmonetaria"
    csql.sql = csql.sql & ",valorcompra+correcionmonetaria+IFNULL((SELECT SUM(correccion_ejercicio) FROM " & clientesistema & "conta" & empresaactiva & ".activo_fijo_nuevo WHERE año <c.año AND codigo=c.codigo),0)-(depreciacion+IFNULL((SELECT SUM(depreciacion_ejercicio) FROM " & clientesistema & "conta" & empresaactiva & ".activo_fijo_nuevo WHERE año <c.año AND codigo=c.codigo),0)) -valorcredito,vidautil"
    csql.sql = csql.sql & ",ifnull((SELECT SUM(vida_ejercicio) FROM " & clientesistema & "conta" & empresaactiva & ".activo_fijo_nuevo WHERE año <c.año AND codigo=c.codigo),0) AS  vidausada,'0',familia"
    csql.sql = csql.sql & ", valorcredito,fechaventa "
    csql.sql = csql.sql & " from activo_fijo_nuevo  as c where año='" & Format(fechasistema, "yyyy") & "' " 'AND (codigo='00001016' OR codigo='00000749' OR codigo='00000723') " 'and codigo='00000767' "
    If dato3.text <> "" Then
        csql.sql = csql.sql & " and familia='" & familia & "'  and codigo='" & codigo & "'   "
    End If
    
    
    csql.sql = csql.sql & "order by familia,fechapuestaenmarcha,codigo  "
    csql.Execute
    
    depreciacionactual = 0
    valorlibrofinal = 0
    Grid1.Rows = 1
    Grid1.AutoRedraw = False
    If csql.RowsAffected > 0 Then
        Set resultados = csql.OpenResultset
        filtro = resultados(11)
        FILTRO2 = filtro
        Grid1.AutoRedraw = False
        
        While Not resultados.EOF
        
            Grid1.Rows = Grid1.Rows + 1
            Grid1.Cell(Grid1.Rows - 1, 0).text = "1"
            Grid1.Cell(Grid1.Rows - 1, 1).text = resultados(0)
            Grid1.Cell(Grid1.Rows - 1, 2).text = resultados(1)
       '     Grid1.Cell(Grid1.Rows - 1, 3).text = leerNOMBREcrcc(resultados(2))
            Grid1.Cell(Grid1.Rows - 1, 4).text = Format(resultados(3), "dd-mm-yyyy")
            Grid1.Cell(Grid1.Rows - 1, 5).text = resultados(4)
            Grid1.Cell(Grid1.Rows - 1, 6).text = resultados(5)
             Grid1.Cell(Grid1.Rows - 1, 7).text = 0 'resultados(6)
            depreciacionactual = resultados(5)
            libro = resultados(4)
            dato9.text = libro
            
            CORRECCION = 0
            depreciacion = 0
            vidaejercicio = 0
            
            
            If libro < 0 Then libro = 1
            Grid1.Cell(Grid1.Rows - 1, 9).text = libro
            
            valorlibrofinal = resultados(7)
            
        If libro > 1 Then
          If Format(resultados(3), "yyyy") < lblaño Then
                Grid1.Cell(Grid1.Rows - 1, 13).text = leeripc("00", lblaño)
             Else
                Grid1.Cell(Grid1.Rows - 1, 13).text = leeripc(Format(resultados(3), "mm"), Format(resultados(3), "yyyy"))
            End If
            
            factor = Grid1.Cell(Grid1.Rows - 1, 13).text / 100
            CORRECCION = Round(valorlibrofinal * factor)
             
             If Format(resultados(3), "yyyy") < Format(fechasistema, "yyyy") Then
                vidaejercicio = 12
             Else
                If IsDate(resultados("fechaventa")) = False Then
                    vidaejercicio = Round(DateDiff("m", resultados(3), fechasistema)) + 1
                Else
                    vidaejercicio = Round(DateDiff("m", resultados(3), resultados("fechaventa"))) + 1
                End If
             End If
             
'             If resultados(8) < vidaejercicio Then
'                vidaejercicio = resultados(8)
'             End If
             
              If resultados(8) - resultados(9) < vidaejercicio Then
                vidaejercicio = resultados(8) - resultados(9)
             End If
             
            Grid1.Cell(Grid1.Rows - 1, 16).text = vidaejercicio
                If IsDate(resultados("fechaventa")) = False Then
                    Grid1.Cell(Grid1.Rows - 1, 17).text = resultados(8) - resultados(9) - vidaejercicio
                Else
                    Grid1.Range(Grid1.Rows - 1, 1, Grid1.Rows - 1, Grid1.Cols - 1).FontItalic = True
                    Grid1.Cell(Grid1.Rows - 1, 17).text = 0
                End If
            vidaanterior = resultados(8) - resultados(9)
            valorlibro = (libro)
            If vidaanterior = 0 Then vidaanterior = 1
            depreciacionmes = (valorlibrofinal + CORRECCION) / vidaanterior
             
            depreciacion = Round(depreciacionmes * vidaejercicio)
            If TieneDepreciacion(resultados("familia")) = False Then
                depreciacion = 0
            End If
            
            If TieneCorreccion(resultados("familia")) = False Then
                CORRECCION = 0
            End If
            
End If
            
            If libro = 0 Then libro = 1
            Grid1.Cell(Grid1.Rows - 1, 11).text = CORRECCION
            Grid1.Cell(Grid1.Rows - 1, 10).text = depreciacion
            Grid1.Cell(Grid1.Rows - 1, 12).text = libro - depreciacion + CORRECCION
            If Grid1.Cell(Grid1.Rows - 1, 12).text = 0 Then Grid1.Cell(Grid1.Rows - 1, 12).text = 1
            Grid1.Cell(Grid1.Rows - 1, 14).text = resultados(8)
            Grid1.Cell(Grid1.Rows - 1, 15).text = resultados(9)
            
            If Grid1.Cell(Grid1.Rows - 1, 17).text = "" Then Grid1.Cell(Grid1.Rows - 1, 17).text = 0
            
    
 
             
 '           Grid1.Cell(Grid1.Rows - 1, 14).text = 0
'            If Format(resultados(3), "yyyy") = Format(fechasistema, "yyyy") Then
                 Grid1.Cell(Grid1.Rows - 1, 8).text = Round(resultados("valorcredito"))
'            End If
            
'            If ExistemMovimientos(resultados(0), Format(fechasistema, "yyyy")) = False Then
'                Call ActualizaValores(resultados(0), Format(fechasistema, "yyyy"), depreciacion, CORRECCION, vidaejercicio)
'            End If
        
'            totales(1) = totales(1) + Grid1.Cell(Grid1.Rows - 1, 5).text
'            totales(2) = totales(2) + Grid1.Cell(Grid1.Rows - 1, 6).text
'            totales(3) = totales(3) + resultados(7)
'            totales(4) = totales(4) + Grid1.Cell(Grid1.Rows - 1, 8).text
'            totales(5) = totales(5) + Grid1.Cell(Grid1.Rows - 1, 9).text
'            totales(6) = totales(6) + Grid1.Cell(Grid1.Rows - 1, 10).text
'            totales(7) = totales(7) + Grid1.Cell(Grid1.Rows - 1, 11).text
'            totales(8) = totales(8) + Grid1.Cell(Grid1.Rows - 1, 17).text
'
'
'            totales2(1) = totales2(1) + Grid1.Cell(Grid1.Rows - 1, 5).text
'            totales2(2) = totales2(2) + Grid1.Cell(Grid1.Rows - 1, 6).text
'            totales2(3) = totales2(3) + resultados(7)
'            totales2(4) = totales2(4) + Grid1.Cell(Grid1.Rows - 1, 8).text
'            totales2(5) = totales2(5) + Grid1.Cell(Grid1.Rows - 1, 9).text
'            totales2(6) = totales2(6) + Grid1.Cell(Grid1.Rows - 1, 10).text
'            totales2(7) = totales2(7) + Grid1.Cell(Grid1.Rows - 1, 11).text
'            totales2(8) = totales2(8) + Grid1.Cell(Grid1.Rows - 1, 17).text
            
             totales(1) = totales(1) + Grid1.Cell(Grid1.Rows - 1, 5).text
            totales(2) = totales(2) + Grid1.Cell(Grid1.Rows - 1, 6).text
            totales(3) = totales(3) + Grid1.Cell(Grid1.Rows - 1, 7).text
            totales(4) = totales(4) + Grid1.Cell(Grid1.Rows - 1, 8).text
            totales(5) = totales(5) + resultados(7)
            totales(6) = totales(6) + Grid1.Cell(Grid1.Rows - 1, 10).text
            totales(7) = totales(7) + Grid1.Cell(Grid1.Rows - 1, 11).text
            totales(8) = totales(8) + Grid1.Cell(Grid1.Rows - 1, 17).text
            
            
            totales2(1) = totales2(1) + Grid1.Cell(Grid1.Rows - 1, 5).text
            totales2(2) = totales2(2) + Grid1.Cell(Grid1.Rows - 1, 6).text
            totales2(3) = totales2(3) + Grid1.Cell(Grid1.Rows - 1, 7).text
            totales2(4) = totales2(4) + Grid1.Cell(Grid1.Rows - 1, 8).text
            totales2(5) = totales2(5) + resultados(7)
            totales2(6) = totales2(6) + Grid1.Cell(Grid1.Rows - 1, 10).text
            totales2(7) = totales2(7) + Grid1.Cell(Grid1.Rows - 1, 11).text
            totales2(8) = totales2(8) + Grid1.Cell(Grid1.Rows - 1, 17).text
            
            
        
            resultados.MoveNext
            If Not resultados.EOF Then
                FILTRO2 = resultados("familia")
            End If
        Wend
                
                
    End If
            If Grid1.Cell(Grid1.Rows - 1, 16).text = "" Then Grid1.Cell(Grid1.Rows - 1, 16).text = 0
            If Grid1.Cell(Grid1.Rows - 1, 15).text = "" Then Grid1.Cell(Grid1.Rows - 1, 15).text = 0
            
      
            
            Me.lblfactor = factor * 100
            lblcorreccion = Format(CORRECCION, "###,###,##0")
            lblvidausadaenejercicio = vidaejercicio
            If lblvidausadaenejercicio < 0 Then lblvidausadaenejercicio = 0
            lblsaldovidautil = CDbl(Grid1.Cell(Grid1.Rows - 1, 14).text) - lblvidausadaenejercicio
             
            If IsDate(lblFechaVenta) = True Then
            lblsaldovidautil.Caption = 0
            End If
            If depreciacion < 0 Then depreciacion = 0
            lblrevalorizado = Format(Grid1.Cell(Grid1.Rows - 1, 12).text, "###,###,##0")
            lbldepreciacionejercicio = Format(depreciacion, "###,###,##0")
    
    Grid1.AutoRedraw = True
    Grid1.Refresh
    csql.Close
    Set csql = Nothing
    Set resultados = Nothing
End Sub


Sub LEERactivofijos2(codigo, familia)
Dim csql As New rdoQuery
Dim resultados As rdoResultset
Dim tipo As String
Dim filtro As String
Dim FILTRO2 As String
Dim totales(10) As Double
Dim totales2(10) As Double
Dim cuentapublicidad As String
Dim factor As Double
Dim CORRECCION As Double
Dim depreciacion As Double
Dim vidaanterior As Double
Dim valorlibro As Double
Dim depreciacionmes As Double
Dim vidaejercicio As Double
Grid1.Cols = 1
 Grid1.Cols = 18
 Grid1.Rows = 1
 
    Set csql.ActiveConnection = contadb
   csql.sql = "select codigo,nombre,crcc,fechapuestaenmarcha,valorcompra"
    csql.sql = csql.sql & " ,depreciacion+ifnull((SELECT SUM(depreciacion_ejercicio) FROM " & clientesistema & "conta" & empresaactiva & ".activo_fijo_nuevo"
    csql.sql = csql.sql & " WHERE año <c.año AND codigo=c.codigo),0) AS depreciacion "
    csql.sql = csql.sql & " ,correcionmonetaria+ifnull((SELECT SUM(correccion_ejercicio) FROM " & clientesistema & "conta" & empresaactiva & ".activo_fijo_nuevo WHERE año <c.año AND codigo=c.codigo),0) AS correccionmonetaria"
    csql.sql = csql.sql & ",valorcompra-depreciacion -valorcredito,vidautil"
    csql.sql = csql.sql & ",vidausada+ifnull((SELECT SUM(correccion_ejercicio) FROM " & clientesistema & "conta" & empresaactiva & ".activo_fijo_nuevo WHERE año <c.año AND codigo=c.codigo),0) AS  vidausada,'0',familia"
    csql.sql = csql.sql & ", valorcredito "
    csql.sql = csql.sql & " from activo_fijo_nuevo  as c where año='" & Format(fechasistema, "yyyy") & "' "
    If dato3.text <> "" Then
        csql.sql = csql.sql & " and familia='" & familia & "' and codigo='" & codigo & "' "
    End If
    
    csql.sql = csql.sql & "order by familia,fechapuestaenmarcha "
    csql.Execute
    Grid1.Rows = 1
    Grid1.AutoRedraw = False
    If csql.RowsAffected > 0 Then
        Set resultados = csql.OpenResultset
        filtro = resultados(11)
        FILTRO2 = filtro
        Grid1.AutoRedraw = False
        
        While Not resultados.EOF
             
            Grid1.Rows = Grid1.Rows + 1
            Grid1.Cell(Grid1.Rows - 1, 0).text = "1"
            Grid1.Cell(Grid1.Rows - 1, 1).text = resultados(0)
            Grid1.Cell(Grid1.Rows - 1, 2).text = resultados(1)
            Grid1.Cell(Grid1.Rows - 1, 3).text = leerNOMBREcrcc(resultados(2))
            Grid1.Cell(Grid1.Rows - 1, 4).text = Format(resultados(3), "dd-mm-yyyy")
            Grid1.Cell(Grid1.Rows - 1, 5).text = resultados(4)
            Grid1.Cell(Grid1.Rows - 1, 6).text = resultados(5)
            Grid1.Cell(Grid1.Rows - 1, 7).text = resultados(6)
            Grid1.Cell(Grid1.Rows - 1, 8).text = resultados(7)
            
          If Format(resultados(3), "yyyy") < Format(fechasistema, "yyyy") Then
                Grid1.Cell(Grid1.Rows - 1, 12).text = leeripc("00", Format(fechasistema, "Yyyy"))
             Else
                Grid1.Cell(Grid1.Rows - 1, 12).text = leeripc(Format(resultados(3), "mm"), Format(resultados(3), "yyyy"))
            End If
            factor = Grid1.Cell(Grid1.Rows - 1, 12).text / 100
            CORRECCION = Round(resultados(7) * factor)
             
             If Format(resultados(3), "yyyy") < Format(fechasistema, "Yyyy") Then
                vidaejercicio = 12
             Else
                vidaejercicio = DateDiff("m", resultados(3), DateAdd("m", 1, Format(fechasistema, "yyyy") & "-12-31"))
             End If
             
            Grid1.Cell(Grid1.Rows - 1, 15).text = vidaejercicio
            Grid1.Cell(Grid1.Rows - 1, 16).text = resultados(8) - resultados(9) - vidaejercicio
            vidaanterior = resultados(8) - resultados(9)
            valorlibro = (resultados(7))
            depreciacionmes = valorlibro / vidaanterior
             
            depreciacion = Round(depreciacionmes * vidaejercicio)
            If TieneDepreciacion(resultados("familia")) = False Then
                depreciacion = 0
            End If
            
            If TieneCorreccion(resultados("familia")) = False Then
                CORRECCION = 0
            End If
            
            Grid1.Cell(Grid1.Rows - 1, 10).text = CORRECCION
            Grid1.Cell(Grid1.Rows - 1, 9).text = depreciacion
            Grid1.Cell(Grid1.Rows - 1, 11).text = resultados(7) - depreciacion + CORRECCION
            Grid1.Cell(Grid1.Rows - 1, 13).text = resultados(8)
            Grid1.Cell(Grid1.Rows - 1, 14).text = resultados(9)
            
            If Grid1.Cell(Grid1.Rows - 1, 17).text = "" Then Grid1.Cell(Grid1.Rows - 1, 17).text = 0
            
    
             
            Grid1.Cell(Grid1.Rows - 1, 17).text = 0
            If Format(resultados(3), "yyyy") = Format(fechasistema, "yyyy") Then
                 Grid1.Cell(Grid1.Rows - 1, 17).text = Round(resultados("valorcompra") * 0.04)
            End If
            
            Me.lblfactor = factor * 100
            lblcorreccion = Format(CORRECCION, "###,###,##0")
            lblvidausadaenejercicio = Grid1.Cell(Grid1.Rows - 1, 14).text
            lblsaldovidautil = Grid1.Cell(Grid1.Rows - 1, 16).text
            lblrevalorizado = Format(Grid1.Cell(Grid1.Rows - 1, 11).text, "###,###,##0")
            lbldepreciacionejercicio = Format(depreciacion, "###,###,##0")
            resultados.MoveNext
 
        Wend
                 
    End If
    Grid1.AutoRedraw = True
 Grid1.Refresh
    csql.Close
    Set csql = Nothing
    Set resultados = Nothing
 
End Sub



Sub AyudaActivoFijo(ByRef caja As TextBox)
    Dim campos As Variant
    Dim cfijo As Variant
    Dim largo As Variant
    campos = Array("codigo", "nombre")
    largo = Array("8s", "40s")
    cfijo = "año='" + Format(fechasistema, "yyyy") + "'"
    cabezas = Array("codigo", "nombre")
    mensajeAyuda = "Ayuda Activos Fijos del año " & Format(fechasistema, "yyyy")

    Call cargaAyudaT(Servidor, basebus, Usuario, password, clientesistema & "conta" & empresaactiva & ".activo_fijo_nuevo", pivote, campos, cfijo, largo, 2)
    If Val(pivote.text) = 0 Then caja.SetFocus: GoTo no
    caja.Enabled = True
    caja.text = pivote.text
 
    
    caja.Enabled = True
    caja.SetFocus
no:
End Sub

Private Sub txtcredito_GotFocus()
    Call cargatexto(txtcredito)
End Sub

Private Sub txtcredito_KeyDown(KeyCode As Integer, Shift As Integer)
Call flechas(dato9, dato10, KeyCode)
End Sub

Private Sub txtcredito_KeyPress(KeyAscii As Integer)
 snum = 0: KeyAscii = esNumero(KeyAscii)
 Dim VALORCREDITO As Double
    If KeyAscii = 13 Then
      If txtcredito.text = "" Then txtcredito.text = 0
        If txtcredito.text > 0 Then
            
                VALORCREDITO = Format((Round(dato9 * 0.04)), "###,###,##0")
            If txtcredito <> VALORCREDITO Then
                MsgBox "ESTA INGRESANDO UN MONTO QUE NO CORRESPONDE ", vbExclamation, " CREDITO " & Format(VALORCREDITO, "###,###,##0")
                txtcredito.text = 0
            End If
            
        End If
        Call Pregunta(txtcredito, dato10)
        
    End If
End Sub
