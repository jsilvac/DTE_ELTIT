VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash.ocx"
Object = "{4777436C-EB8C-4596-98A8-EBCF98683969}#1.0#0"; "FlexCell.ocx"
Object = "{ADBBAED6-B16F-40DC-80DF-B44910CBA76C}#1.0#0"; "Frame-Xp.ocx"
Begin VB.Form arriendo06 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   9840
   ClientLeft      =   2040
   ClientTop       =   1425
   ClientWidth     =   15240
   DrawStyle       =   5  'Transparent
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   656
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1016
   StartUpPosition =   2  'CenterScreen
   Begin XPFrame.FrameXp FrameQuickMenu 
      Height          =   615
      Left            =   11880
      TabIndex        =   40
      Top             =   8880
      Width           =   3135
      _ExtentX        =   5530
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
      Begin VB.CommandButton botonmisfavoritos 
         Caption         =   "Mis Favoritos"
         Height          =   255
         Left            =   1680
         TabIndex        =   42
         Top             =   280
         Width           =   1335
      End
      Begin VB.CommandButton botonmisaccesos 
         Caption         =   "Permisos Modulo"
         Height          =   255
         Left            =   120
         TabIndex        =   41
         Top             =   280
         Width           =   1455
      End
   End
   Begin FlexCell.Grid Grid4 
      Height          =   375
      Left            =   7695
      TabIndex        =   39
      Top             =   9090
      Visible         =   0   'False
      Width           =   780
      _ExtentX        =   1376
      _ExtentY        =   661
      Cols            =   5
      DefaultFontSize =   8.25
      Rows            =   30
   End
   Begin XPFrame.FrameXp frmdetalle 
      Height          =   8295
      Left            =   7830
      TabIndex        =   18
      Top             =   180
      Width           =   7260
      _ExtentX        =   12806
      _ExtentY        =   14631
      BackColor       =   16761024
      Caption         =   "Detalle de Arriendos"
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
      Begin VB.CommandButton Command1 
         Caption         =   "Procesar  Pago"
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
         Left            =   855
         TabIndex        =   35
         Top             =   7920
         Width           =   2175
      End
      Begin VB.TextBox uf 
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
         Left            =   1845
         MaxLength       =   50
         TabIndex        =   24
         Tag             =   "montoarriendo"
         Text            =   "1"
         Top             =   7515
         Width           =   1395
      End
      Begin FlexCell.Grid Grid1 
         Height          =   7125
         Left            =   45
         TabIndex        =   19
         Top             =   270
         Width           =   7125
         _ExtentX        =   12568
         _ExtentY        =   12568
         Cols            =   5
         DefaultFontSize =   8.25
         Rows            =   30
      End
      Begin VB.Label lbltotalpesos 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000007&
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
         ForeColor       =   &H0000FFFF&
         Height          =   330
         Left            =   5400
         TabIndex        =   31
         Top             =   7920
         Width           =   1770
      End
      Begin VB.Label LBLtotal 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000007&
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
         ForeColor       =   &H0000FFFF&
         Height          =   330
         Left            =   5400
         TabIndex        =   30
         Top             =   7560
         Width           =   1770
      End
      Begin VB.Label Label7 
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Valor Arriendos"
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
         Height          =   300
         Left            =   3825
         TabIndex        =   27
         Top             =   7920
         Width           =   1530
      End
      Begin VB.Label Label5 
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Valor Arriendos"
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
         Height          =   300
         Left            =   3825
         TabIndex        =   26
         Top             =   7560
         Width           =   1530
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Valor U.F"
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
         Left            =   225
         TabIndex        =   25
         Top             =   7515
         Width           =   1530
      End
   End
   Begin XPFrame.FrameXp FrameXp1 
      Height          =   8325
      Left            =   180
      TabIndex        =   9
      Top             =   180
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   14684
      BackColor       =   16744576
      Caption         =   "DATOS  PAGO"
      CaptionEstilo3D =   1
      BackColor       =   16744576
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
      Begin VB.TextBox txttipo 
         Alignment       =   2  'Center
         BackColor       =   &H00FFC0C0&
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
         Left            =   1665
         MaxLength       =   2
         TabIndex        =   32
         Tag             =   "monedaarriendo"
         Top             =   2070
         Width           =   375
      End
      Begin VB.TextBox txtpropiedad 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFC0C0&
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
         Left            =   1620
         MaxLength       =   10
         TabIndex        =   29
         Tag             =   "numero"
         Top             =   1710
         Width           =   1215
      End
      Begin VB.TextBox dato2 
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
         Left            =   1620
         MaxLength       =   2
         TabIndex        =   1
         Tag             =   "Comuna"
         Top             =   630
         Width           =   375
      End
      Begin XPFrame.FrameXp FrameXp4 
         Height          =   2985
         Left            =   45
         TabIndex        =   22
         Top             =   5220
         Visible         =   0   'False
         Width           =   7485
         _ExtentX        =   13203
         _ExtentY        =   5265
         BackColor       =   16711680
         Caption         =   "HISTORICO DE PAGOS"
         CaptionEstilo3D =   1
         BackColor       =   16711680
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
         Begin FlexCell.Grid Grid3 
            Height          =   2625
            Left            =   45
            TabIndex        =   23
            Top             =   270
            Width           =   7395
            _ExtentX        =   13044
            _ExtentY        =   4630
            Cols            =   5
            DefaultFontSize =   8.25
            Rows            =   30
         End
      End
      Begin XPFrame.FrameXp frmdetallepago 
         Height          =   2355
         Left            =   45
         TabIndex        =   20
         Top             =   2835
         Width           =   7485
         _ExtentX        =   13203
         _ExtentY        =   4154
         BackColor       =   16761024
         Caption         =   "DETALLE DEL PAGO"
         CaptionEstilo3D =   1
         BackColor       =   16761024
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
         Begin VB.CommandButton Command2 
            Caption         =   "Grabar Comprobante"
            Height          =   285
            Left            =   4815
            TabIndex        =   38
            Top             =   2025
            Width           =   1995
         End
         Begin FlexCell.Grid Grid2 
            Height          =   1680
            Left            =   45
            TabIndex        =   21
            Top             =   315
            Width           =   7395
            _ExtentX        =   13044
            _ExtentY        =   2963
            Cols            =   5
            DefaultFontSize =   8.25
            Rows            =   30
         End
         Begin VB.Label lblpagado 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00000000&
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
            ForeColor       =   &H0000FFFF&
            Height          =   285
            Left            =   1665
            TabIndex        =   37
            Top             =   2025
            Width           =   1725
         End
         Begin VB.Label Label10 
            BackColor       =   &H00FFC0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Monto"
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
            Left            =   45
            TabIndex        =   36
            Top             =   2025
            Width           =   1530
         End
      End
      Begin VB.TextBox txtrut 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFC0C0&
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
         Left            =   1620
         MaxLength       =   9
         TabIndex        =   12
         Tag             =   "rutarrendatario"
         Top             =   1350
         Width           =   1215
      End
      Begin VB.TextBox dato1 
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
         Left            =   1620
         MaxLength       =   10
         TabIndex        =   0
         Tag             =   "numero"
         Top             =   270
         Width           =   1215
      End
      Begin VB.TextBox dato3 
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
         Left            =   2070
         MaxLength       =   2
         TabIndex        =   2
         Tag             =   "Comuna"
         Top             =   630
         Width           =   375
      End
      Begin VB.TextBox dato4 
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
         Left            =   2520
         MaxLength       =   4
         TabIndex        =   3
         Tag             =   "ciudad"
         Top             =   630
         Width           =   705
      End
      Begin VB.TextBox dato6 
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
         Height          =   330
         Left            =   1620
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   5
         Tag             =   "montoarriendo"
         Top             =   2430
         Width           =   1395
      End
      Begin VB.TextBox DATO5 
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
         Left            =   1620
         MaxLength       =   10
         TabIndex        =   4
         Tag             =   "numero"
         Top             =   990
         Width           =   1215
      End
      Begin VB.Label Label9 
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "moneda Arriendo"
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
         Height          =   300
         Left            =   45
         TabIndex        =   34
         Top             =   2070
         Width           =   1530
      End
      Begin VB.Label lblmoneda 
         BackColor       =   &H00FFC0C0&
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
         ForeColor       =   &H00400000&
         Height          =   300
         Left            =   2205
         TabIndex        =   33
         Top             =   2070
         Width           =   3930
      End
      Begin VB.Label Label8 
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Contrato"
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
         Height          =   300
         Left            =   45
         TabIndex        =   28
         Top             =   990
         Width           =   1530
      End
      Begin VB.Label lblpropiedad 
         BackColor       =   &H00FFC0C0&
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
         ForeColor       =   &H00400000&
         Height          =   300
         Left            =   2880
         TabIndex        =   17
         Top             =   1710
         Width           =   4605
      End
      Begin VB.Label Label6 
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Monto"
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
         Height          =   300
         Left            =   45
         TabIndex        =   16
         Top             =   2430
         Width           =   1530
      End
      Begin VB.Label lblarrendatario 
         BackColor       =   &H00FFC0C0&
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
         ForeColor       =   &H00400000&
         Height          =   300
         Left            =   3195
         TabIndex        =   15
         Top             =   1350
         Width           =   4290
      End
      Begin VB.Label dv2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFC0C0&
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
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   2880
         TabIndex        =   6
         Top             =   1350
         Width           =   255
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Numero"
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
         Height          =   300
         Left            =   45
         TabIndex        =   14
         Top             =   270
         Width           =   1530
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Propiedad"
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
         Height          =   300
         Left            =   45
         TabIndex        =   13
         Top             =   1710
         Width           =   1530
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Fecha Pago"
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
         Height          =   300
         Left            =   45
         TabIndex        =   11
         Top             =   630
         Width           =   1530
      End
      Begin VB.Label Label11 
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Rut Arrendatario"
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
         Height          =   300
         Index           =   0
         Left            =   45
         TabIndex        =   10
         Top             =   1350
         Width           =   1530
      End
   End
   Begin VB.PictureBox MANUAL 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H80000002&
      ForeColor       =   &H80000008&
      Height          =   0
      Left            =   0
      ScaleHeight     =   0
      ScaleWidth      =   15210
      TabIndex        =   8
      Top             =   9840
      Width           =   15240
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash opciones 
      Height          =   1215
      Left            =   270
      TabIndex        =   7
      Top             =   8550
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
Attribute VB_Name = "arriendo06"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public saldoglobal As Double
Private moneda As String
Private rutpropi As String

Private MODIFI As Integer

Private Sub codigo_Click()
    Call dato1_KeyDown(vbKeyF2, 0)
End Sub

 
Private Sub Command1_Click()
If dato6.text <> "" And dato6.text <> "0" Then
Grid2.Rows = 2
Grid2.Cell(1, 1).SetFocus
End If

End Sub

Private Sub COMMAND2_Click()
If lblpagado.Caption = "" Then lblpagado.Caption = "0"

If CDbl(lblpagado.Caption) = CDbl(dato6.text) Then
grabar

retorno

End If


End Sub

Private Sub dato1_GotFocus()
dato1.text = LEERULTIMOFOLIOcontrato
Call cargatexto(dato1)

End Sub

Private Sub dato1_KeyDown(KeyCode As Integer, Shift As Integer)
    Call flechas(dato1, dato2, KeyCode)
End Sub

Private Sub dato2_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF2 Then Call ayudapropiedades(dato2)
       Call flechas(dato1, dato3, KeyCode)
End Sub
Private Sub dato3_KeyDown(KeyCode As Integer, Shift As Integer)
      Call flechas(dato2, dato4, KeyCode)
End Sub
Private Sub dato4_KeyDown(KeyCode As Integer, Shift As Integer)
    Call flechas(dato3, DATO5, KeyCode)
End Sub

Private Sub dato6_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then

End If

End Sub

Private Sub Grid1_Click()
Dim suma As Double
Dim GASTOS As Double

suma = 0
GASTOS = 0
For k = 1 To Grid1.Rows - 1
If Grid1.Cell(k, 5).text = "1" Then
suma = suma + CDbl(Grid1.Cell(k, 2).text)
GASTOS = GASTOS + CDbl(Grid1.Cell(k, 4).text)
End If
Next k
lbltotal.Caption = Format(suma, "###,###,###,###")

If txttipo.text <> "01" Then
suma = suma * CDbl(uf.text)
End If


lbltotalpesos.Caption = Format(suma + GASTOS, "###,###,###,###")

dato6.text = lbltotalpesos.Caption

End Sub

Private Sub Grid2_DblClick()
dato1.text = Grid2.Cell(Grid2.ActiveCell.row, 4).text
Call dato1_KeyPress(13)
End Sub

Private Sub Label12_Click()

End Sub

Private Sub Grid2_LeaveCell(ByVal row As Long, ByVal col As Long, NewRow As Long, NewCol As Long, Cancel As Boolean)
Dim total As Double

If col = 3 And NewCol = 4 And Grid2.Cell(row, 1).text = "1" And Grid2.Rows - 1 = row Then
Grid2.Rows = Grid2.Rows + 1
NewRow = row + 1
NewCol = 1
End If
For k = 1 To Grid2.Rows - 1
If Grid2.Cell(k, 3).text <> "" And Grid2.Cell(k, 3).text <> "0" Then
total = total + CDbl(Grid2.Cell(k, 3).text)
End If
Next k
lblpagado.Caption = Format(total, "###,###,###,###")

End Sub

 Private Sub MANUAL_KeyPress(KeyAscii As Integer)
If UCase(Chr(KeyAscii)) = "M" Then Call opciones_FSCommand("modifica", "")
If UCase(Chr(KeyAscii)) = "E" Then Call opciones_FSCommand("elimina", "")
If UCase(Chr(KeyAscii)) = "S" Then Call opciones_FSCommand("siguiente", "")
If UCase(Chr(KeyAscii)) = "A" Then Call opciones_FSCommand("anterior", "")
If UCase(Chr(KeyAscii)) = "R" Then Call opciones_FSCommand("retorno", "")
If UCase(Chr(KeyAscii)) = "I" Then Call opciones_FSCommand("imprime", "")
End Sub

Private Sub Form_Load()
Call CENTRAR(Me)
    Call Conectar_BD
    Rem Call Funciones_Forms_M_Productos.Conecta_Maestro_Productos
    sc = 0
    opciones.Visible = False
 
Rem Call RECUPERAFECHA

Call CARGAPERMISO(Me.Name)
Call CARGAGRILLA
Call CARGAGRILLA2
Call CARGAGRILLA3
End Sub

Private Sub dato1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then Unload Me
    snum = 0: KeyAscii = esNumero(KeyAscii)
    If KeyAscii = 13 And Val(dato1.text) <> 0 Then
        Call ceros(dato1)
        leer
    End If
End Sub


Private Sub dato2_KeyPress(KeyAscii As Integer)
        snum = 0: KeyAscii = esNumero(KeyAscii)
    If KeyAscii = 13 Then
        Call ceros(dato2)
        If dato2.text = "00" Then dato2.text = Format(fechasistema, "dd")
        Call Pregunta(dato2, dato3)
    End If
End Sub
Private Sub dato3_KeyPress(KeyAscii As Integer)
        snum = 0: KeyAscii = esNumero(KeyAscii)
    If KeyAscii = 13 Then
        Call ceros(dato3)
        If dato3.text = "00" Then dato3.text = Format(fechasistema, "mm")
        Call Pregunta(dato3, dato4)
    End If
    End Sub
Private Sub dato4_KeyPress(KeyAscii As Integer)
        snum = 0: KeyAscii = esNumero(KeyAscii)
    If KeyAscii = 13 Then
        Call ceros(dato4)
        If dato4.text = "0000" Then dato4.text = Format(fechasistema, "yyyy")
        Call Pregunta(dato4, DATO5)
    End If
    End Sub
Private Sub dato5_KeyPress(KeyAscii As Integer)
        snum = 0: KeyAscii = esNumero(KeyAscii)
    If KeyAscii = 13 Then
        
        Call ceros(DATO5)
        If existecontrato(DATO5.text) = True Then
        dato6.Enabled = True
        Call leermensualidades(DATO5.text)
        
        
        dato6.SetFocus
        
        Else
        DATO5.SetFocus
        
        End If
        
    End If
    End Sub
    
    

Sub leer()
    campos(0, 0) = "numero"
    campos(1, 0) = "fecha"
    campos(2, 0) = "contrato"
    campos(3, 0) = "monto"
    campos(4, 0) = ""
    
    
    campos(0, 2) = clientesistema & "arriendos" & ".arriendos_pago_cabeza"
    condicion = "numero= '" & dato1.text & "' "

    op = 5
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)
    If sqlconta.status = 4 Then dato2.Enabled = True: dato2.SetFocus: GoTo no:
    carga
    opciones.Visible = True
    disponible (True)
    habilita (True)
    opciones.SetFocus
        
no:
End Sub
Sub leersiguiente()
    campos(0, 0) = "numero"
    campos(1, 0) = "fecha"
    campos(2, 0) = "contrato"
    campos(3, 0) = "monto"
    campos(4, 0) = ""
    
    
    campos(0, 2) = clientesistema & "arriendos" & ".arriendos_pago_cabeza"
    condicion = "numero > '" & dato1.text & "' order by numero "
    
    
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
    campos(0, 0) = "numero"
    campos(1, 0) = "fecha"
    campos(2, 0) = "contrato"
    campos(3, 0) = "monto"
    campos(4, 0) = ""
    
    campos(0, 2) = clientesistema & "arriendos" & ".contratos_arriendo"
    condicion = "numero < '" & dato1.text & "' order by numero desc "
    
    
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
    habilita (True)
    dato1.text = sqlconta.response(0, 3)
    
    dato2.text = Mid(sqlconta.response(1, 3), 1, 2)
    dato3.text = Mid(sqlconta.response(1, 3), 4, 2)
    dato4.text = Mid(sqlconta.response(1, 3), 7, 4)
    DATO5.text = sqlconta.response(2, 3)
    dato6.text = Format(sqlconta.response(3, 3), "###,###,###")
    Call existecontrato(DATO5.text)
    
    
    Call leerdetallepago(dato1.text)
    Call leermensualidadespagadas(dato1.text)
    
    frmdetallepago.Enabled = False
    

    
fin:
End Sub

Sub habilita(ByVal condicion As Boolean)
    
    dato1.Locked = condicion
    dato2.Locked = condicion
    dato3.Locked = condicion
    dato4.Locked = condicion
    DATO5.Locked = condicion
    
End Sub
Sub disponible(ByVal condicion As Boolean)
    
    dato1.Enabled = condicion
    dato2.Enabled = condicion
    dato3.Enabled = condicion
    dato4.Enabled = condicion
    DATO5.Enabled = condicion
    
End Sub


Sub Pregunta(ByRef caja As TextBox, ByRef caja1 As TextBox)
    If caja.text = "" And sc = 0 Then caja.SetFocus
    If caja.text <> "" Or sc = 1 Then caja1.Enabled = True: caja1.SetFocus
    sc = 0
End Sub


Sub ayudaarrendadores(ByRef caja As TextBox)
    Dim campos As Variant
    Dim cfijo As Variant
    Dim largo As Variant
    campos = Array("rut", "nombre")
    largo = Array("11s", "40s")
    cfijo = "rut like '%%'"
    cabezas = Array("Rut", "Nombre")
    mensajeAyuda = "Ayuda de Arrendadores"
       
    Call cargaAyudaT(Servidor, clientesistema & "arriendos", Usuario, password, ".maestro_arrendadores", caja, campos, cfijo, largo, 2)
    caja.Enabled = True
    caja.SetFocus
End Sub
Sub ayudaarrendatarios(ByRef caja As TextBox)
    Dim campos As Variant
    Dim cfijo As Variant
    Dim largo As Variant
    campos = Array("rut", "nombre")
    largo = Array("11s", "40s")
    cfijo = "rut like '%%'"
    cabezas = Array("Rut", "Nombre")
    mensajeAyuda = "Ayuda de Arrendatarios"
       
    Call cargaAyudaT(Servidor, clientesistema & "arriendos", Usuario, password, ".maestro_arrendatarios", caja, campos, cfijo, largo, 2)
    caja.Enabled = True
    caja.SetFocus
End Sub
Sub ayudapropiedades(ByRef caja As TextBox)
    Dim campos As Variant
    Dim cfijo As Variant
    Dim largo As Variant
    campos = Array("codigopropiedad", "direccion")
    largo = Array("11s", "40s")
    cfijo = "codigopropiedad like '%%'"
    cabezas = Array("Codigo", "Nombre")
    mensajeAyuda = "Ayuda de Propiedades"
       
    Call cargaAyudaT(Servidor, clientesistema & "arriendos", Usuario, password, ".maestro_propiedades", caja, campos, cfijo, largo, 2)
    caja.Enabled = True
    caja.SetFocus
End Sub


Sub flechas(ByRef caja As TextBox, ByRef caja1 As TextBox, ByRef codigo As Integer)
    If codigo = 38 And caja.Enabled = True Then caja.SetFocus
    If codigo = 40 And caja1.Enabled = True Then caja1.SetFocus
End Sub

Sub grabar()
    Dim LINEAS As Double
    
    campos(0, 0) = "numero"
    campos(1, 0) = "fecha"
    campos(2, 0) = "contrato"
    campos(3, 0) = "monto"
    campos(4, 0) = ""
    campos(0, 1) = dato1.text
    campos(1, 1) = dato4.text + "-" + dato3.text + "-" + dato2.text
    campos(2, 1) = DATO5.text
    campos(3, 1) = Replace(dato6.text, ".", "")
    
    
    campos(0, 2) = clientesistema & "arriendos" & ".arriendos_pago_cabeza "
    condicion = ""
    op = 2
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)
    LINEAS = 0
    For k = 1 To Grid2.Rows - 1
    If Grid2.Cell(k, 3).text <> "" Then
    LINEAS = LINEAS + 1
    Call grabardetallepagos(dato1.text, LINEAS, Grid2.Cell(k, 1).text, Grid2.Cell(k, 3).text, Grid2.Cell(k, 4).text, Grid2.Cell(k, 5).text, Grid2.Cell(k, 6).text, Grid2.Cell(k, 7).text)
    End If
    
    Next k
    
    For k = 1 To Grid1.Rows - 1
    If Grid1.Cell(k, 5).text = "1" Then
    Call modificapago(DATO5.text, Grid1.Cell(k, 1).text, "1", dato1.text)
    
    
    End If
    
    Next k
    
    End Sub
Sub grabardetallepagos(numero, LINEA, tipo, monto, banco, cuenta, numerocheque, vencimiento)
    campos(0, 0) = "numero"
    campos(1, 0) = "linea"
    campos(2, 0) = "tipo"
    campos(3, 0) = "monto"
    campos(4, 0) = "banco"
    campos(5, 0) = "cuenta"
    campos(6, 0) = "numerocheque"
    campos(7, 0) = "vencimiento"
    campos(8, 0) = ""
    campos(0, 1) = numero
    campos(1, 1) = LINEA
    campos(2, 1) = tipo
    campos(3, 1) = monto
    campos(4, 1) = banco
    campos(5, 1) = cuenta
    campos(6, 1) = numerocheque
    
    campos(7, 1) = Format(vencimiento, "yyyy-mm-dd")
    
    campos(0, 2) = clientesistema & "arriendos" & ".arriendos_pago_detalle "
    condicion = ""
    op = 2
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)
    
    End Sub
 
 
Sub ELIMINAR()
    campos(0, 2) = clientesistema & "arriendos" & ".arriendos_pago_cabeza "
    condicion = "numero=" + "'" + dato1.text + "' "
    op = 4
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)

    campos(0, 2) = clientesistema & "arriendos" & ".arriendos_pago_detalle "
    condicion = "numero=" + "'" + dato1.text + "' "
    op = 4
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)


End Sub
  

Private Sub opciones_FSCommand(ByVal command As String, ByVal args As String)

If command = "retorno" Then retorno
If command = "modifica" Then
    MsgBox "IMPOSIBLE MODIFICAR COMPROBANTES "
        
End If

If command = "elimina" Then
    If Verifica_Permiso(Me.Caption, "elimina") = True Then
        ELIMINA
    End If
End If


If command = "siguiente" Then leersiguiente
If command = "anterior" Then leeranterior
 
If command = "imprime" Then
imprimir

End If




End Sub
Sub ELIMINA()
 
disponible (True)
habilita (False)
ELIMINAR
limpia
opciones.Visible = False
dato1.SetFocus
 
End Sub

Sub modifica()
disponible (True)
habilita (False)
dato1.Enabled = False
dato2.SetFocus
MODIFI = 1

End Sub
Sub retorno()

disponible (True)
habilita (False)
limpia
frmdetalle.Enabled = True
frmdetallepago.Enabled = True

opciones.Visible = False
dato1.Enabled = True
dato1.SetFocus
MODIFI = 0
no:
 
 
    
End Sub
Sub limpia()
    dato1.text = ""
    dato2.text = ""
    lblpropiedad.Caption = ""
    lblarrendatario.Caption = ""
    LBLMONEDA.Caption = ""
    txtrut.text = ""
    txtpropiedad.text = ""
    txttipo.text = ""
    lbltotal.Caption = ""
    lbltotalpesos.Caption = ""
    lblpagado.Caption = ""
    
    
    
    dato3.text = ""
    dato4.text = ""
    DATO5.text = ""
    dato6.text = ""
    Grid2.Rows = 1
    Grid1.Rows = 1
    dv2.Caption = ""
End Sub
 
Sub cargatexto(ByRef caja As TextBox)
caja.SelStart = 0: caja.SelLength = Len(caja.text)
End Sub

Private Sub opciones_GotFocus()
MANUAL.SetFocus
End Sub

 Private Function leearrendador(rutarrendador) As String
 Dim csql As New rdoQuery
 Dim resultados As rdoResultset
 
 Set csql.ActiveConnection = contadb
 csql.sql = "select nombre from " & clientesistema & "arriendos" & ".maestro_arrendadores "
 csql.sql = csql.sql & "where rut='" & rutarrendador & "' "
 csql.Execute
 leearrendador = ""
 
 If csql.RowsAffected > 0 Then
    Set resultados = csql.OpenResultset
    leearrendador = resultados(0)
 Else
    leearrendador = ""
 End If
 
 csql.Close
 Set csql = Nothing
 Set resultados = Nothing
 
 End Function
 Private Function leearrendatario(rutarrendatario) As String
 Dim csql As New rdoQuery
 Dim resultados As rdoResultset
 
 Set csql.ActiveConnection = contadb
 csql.sql = "select nombre from " & clientesistema & "arriendos" & ".maestro_arrendatarios "
 csql.sql = csql.sql & "where rut='" & rutarrendatario & "' "
 csql.Execute
 leearrendatario = ""
 
 If csql.RowsAffected > 0 Then
    Set resultados = csql.OpenResultset
    leearrendatario = resultados(0)
    
 Else
    leearrendatario = ""
 End If
 
 csql.Close
 Set csql = Nothing
 Set resultados = Nothing
 
 End Function
 
 Private Function leepropiedad(codigopropiedad) As String
 Dim csql As New rdoQuery
 Dim resultados As rdoResultset
 
 Set csql.ActiveConnection = contadb
 csql.sql = "select direccion,monedaarriendo,rutpropietario from " & clientesistema & "arriendos" & ".maestro_propiedades "
 csql.sql = csql.sql & "where codigopropiedad='" & codigopropiedad & "' "
 csql.Execute
 leepropiedad = ""
 moneda = ""
 If csql.RowsAffected > 0 Then
    Set resultados = csql.OpenResultset
    leepropiedad = resultados(0)
    moneda = resultados(1)
    rutpropi = resultados(2)
 Else
    leepropiedad = ""

 End If
 
 csql.Close
 Set csql = Nothing
 Set resultados = Nothing
 
 End Function

Sub CARGAGRILLA()
    Dim formatogrilla2(10, 10)
    formatogrilla2(1, 1) = "FECHA"
    formatogrilla2(1, 2) = "MONTO"
    formatogrilla2(1, 3) = "MONEDA"
    formatogrilla2(1, 4) = "G/COMUNES "
    formatogrilla2(1, 5) = "PAGAR"
    
    Rem LARGO DE LOS DATOS
    formatogrilla2(2, 1) = "10"
    formatogrilla2(2, 2) = "10"
    formatogrilla2(2, 3) = "10"
    formatogrilla2(2, 4) = "10"
    formatogrilla2(2, 5) = "10"
    formatogrilla2(2, 6) = "10"
    formatogrilla2(2, 7) = "17"
    
    Rem TIPO DE DATOS C=CEROS,S=STRING,N=NUMERICO,D=DATE,
    formatogrilla2(3, 1) = "D"
    formatogrilla2(3, 2) = "N"
    formatogrilla2(3, 3) = "N"
    formatogrilla2(3, 4) = "N"
    formatogrilla2(3, 5) = "N"
    formatogrilla2(3, 6) = "N"
    formatogrilla2(3, 7) = "N"
    
    Rem FORMATO GRILLA
    
    formatogrilla2(4, 2) = " ###,###,##0.00"
    formatogrilla2(4, 3) = " ###,###,##0.00"
    formatogrilla2(4, 4) = " ###,###,##0"
    formatogrilla2(4, 5) = " ###,###,##0"
    
    Rem LOCCKED
    formatogrilla2(5, 1) = "TRUE"
    formatogrilla2(5, 2) = "TRUE"
    formatogrilla2(5, 3) = "TRUE"
    formatogrilla2(5, 4) = "TRUE"
    formatogrilla2(5, 5) = "FALSE"
    formatogrilla2(5, 6) = "TRUE"
    
    
    Rem VALOR MAXIMO
    
    Grid1.Cols = 6
    Grid1.Rows = 1
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
    Grid1.Column(0).Width = 0
    
    For k = 1 To Grid1.Cols - 1
        Grid1.Cell(0, k).text = formatogrilla2(1, k)
        Grid1.Column(k).Width = Val(formatogrilla2(2, k)) * 9
        Grid1.Column(k).MaxLength = Val(formatogrilla2(2, k))
        Grid1.Column(k).FormatString = formatogrilla2(4, k)
        Grid1.Column(k).Locked = formatogrilla2(5, k)
        If formatogrilla2(3, k) = "N" Then Grid1.Column(k).Alignment = cellRightCenter
        If formatogrilla2(3, k) = "S" Then Grid1.Column(k).Alignment = cellLeftTop
        If formatogrilla2(3, k) = "D" Then Grid1.Column(k).CellType = cellCalendar
        
    Next k
     Grid1.Column(5).CellType = cellCheckBox
     
     
    End Sub

Sub CARGAGRILLA2()
    Dim formatogrilla2(10, 12)
    formatogrilla2(1, 1) = "TIPO"
    formatogrilla2(1, 2) = "NOMBRE"
    formatogrilla2(1, 3) = "MONTO"
    formatogrilla2(1, 4) = "BANCO"
    formatogrilla2(1, 5) = "CUENTA"
    formatogrilla2(1, 6) = "NUMERO"
    formatogrilla2(1, 7) = "VENCIMIENTO"
    
    Rem LARGO DE LOS DATOS
    formatogrilla2(2, 1) = "4"
    formatogrilla2(2, 2) = "5"
    formatogrilla2(2, 3) = "8"
    formatogrilla2(2, 4) = "4"
    formatogrilla2(2, 5) = "8"
    formatogrilla2(2, 6) = "8"
    formatogrilla2(2, 7) = "8"
    
    Rem TIPO DE DATOS C=CEROS,S=STRING,N=NUMERICO,D=DATE,
    formatogrilla2(3, 1) = "N"
    formatogrilla2(3, 2) = "S"
    formatogrilla2(3, 3) = "N"
    formatogrilla2(3, 4) = "S"
    formatogrilla2(3, 5) = "S"
    formatogrilla2(3, 6) = "S"
    formatogrilla2(3, 7) = "D"
    Rem FORMATO GRILLA
    
    formatogrilla2(4, 3) = " ###,###,##0.00"
    
    
    Rem LOCCKED
    formatogrilla2(5, 1) = "FALSE"
    formatogrilla2(5, 2) = "TRUE"
    formatogrilla2(5, 3) = "FALSE"
    formatogrilla2(5, 4) = "FALSE"
    formatogrilla2(5, 5) = "FALSE"
    formatogrilla2(5, 6) = "FALSE"
    formatogrilla2(5, 7) = "FALSE"
    
    
    Rem VALOR MAXIMO
    
    Grid2.Cols = 8
    Grid2.Rows = 1
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
    Grid2.Column(0).Width = 0
    
    For k = 1 To Grid2.Cols - 1
        Grid2.Cell(0, k).text = formatogrilla2(1, k)
        Grid2.Column(k).Width = Val(formatogrilla2(2, k)) * 9
        Grid2.Column(k).MaxLength = Val(formatogrilla2(2, k))
        Grid2.Column(k).FormatString = formatogrilla2(4, k)
        Grid2.Column(k).Locked = formatogrilla2(5, k)
        If formatogrilla2(3, k) = "N" Then Grid2.Column(k).Alignment = cellRightCenter
        If formatogrilla2(3, k) = "S" Then Grid2.Column(k).Alignment = cellLeftTop
        
        
        If formatogrilla2(3, k) = "D" Then Grid2.Column(k).CellType = cellCalendar
        
    Next k
    
    
    Grid2.Column(3).Mask = cellNumeric
    
    
    End Sub
Sub CARGAGRILLA3()
    Dim formatogrilla2(10, 10)
    formatogrilla2(1, 1) = "FECHA"
    formatogrilla2(1, 2) = "MONTO"
    formatogrilla2(1, 3) = "TIPO"
    
    Rem LARGO DE LOS DATOS
    formatogrilla2(2, 1) = "10"
    formatogrilla2(2, 2) = "10"
    formatogrilla2(2, 3) = "10"
    
    Rem TIPO DE DATOS C=CEROS,S=STRING,N=NUMERICO,D=DATE,
    formatogrilla2(3, 1) = "D"
    formatogrilla2(3, 2) = "N"
    formatogrilla2(3, 3) = "N"
    
    Rem FORMATO GRILLA
    
    formatogrilla2(4, 2) = " ###,###,##0"
    
    Rem LOCCKED
    formatogrilla2(5, 1) = "TRUE"
    formatogrilla2(5, 2) = "TRUE"
    formatogrilla2(5, 3) = "TRUE"
    
    
    Rem VALOR MAXIMO
    
    Grid3.Cols = 6
    Grid3.Rows = 1
    Grid3.AllowUserResizing = False
    Grid3.DisplayFocusRect = False
    Grid3.ExtendLastCol = True
    Grid3.BoldFixedCell = False
    Grid3.DrawMode = cellOwnerDraw
    Grid3.Appearance = Flat
    Grid3.ScrollBarStyle = Flat
    Grid3.FixedRowColStyle = Flat
    Grid3.BackColorFixed = RGB(90, 158, 214)
    Grid3.BackColorFixedSel = RGB(110, 180, 230)
    Grid3.BackColorBkg = RGB(90, 158, 214)
    Grid3.BackColorScrollBar = RGB(231, 235, 247)
    Grid3.BackColor1 = RGB(231, 235, 247)
    Grid3.BackColor2 = RGB(239, 243, 255)
    Grid3.GridColor = RGB(148, 190, 231)
    Grid3.Column(0).Width = 0
    
    For k = 1 To Grid3.Cols - 1
        Grid3.Cell(0, k).text = formatogrilla2(1, k)
        Grid3.Column(k).Width = Val(formatogrilla2(2, k)) * 9
        Grid3.Column(k).MaxLength = Val(formatogrilla2(2, k))
        Grid3.Column(k).FormatString = formatogrilla2(4, k)
        Grid3.Column(k).Locked = formatogrilla2(5, k)
        If formatogrilla2(3, k) = "N" Then Grid3.Column(k).Alignment = cellRightCenter
        If formatogrilla2(3, k) = "S" Then Grid3.Column(k).Alignment = cellLeftTop
        If formatogrilla2(3, k) = "D" Then Grid3.Column(k).CellType = cellCalendar
        
    Next k
     Grid3.Column(5).CellType = cellCheckBox
     
     
    End Sub


 Public Sub leerpropiedades()
 Dim csql As New rdoQuery
 Dim resultados As rdoResultset
 
 Set csql.ActiveConnection = contadb
 csql.sql = "select mp.codigopropiedad,mp.nombrepropiedad,mp.direccion,ca.numero,ca.rutarrendatario,ca.fechainicio,ca.fechatermino,ca.montoarriendo,ca.monedaarriendo,ca.gastoscomunes  from " & clientesistema & "arriendos" & ".maestro_propiedades as mp left join " + clientesistema + "arriendos" & ".contratos_arriendo as ca on (mp.codigopropiedad = ca.propiedad) order by mp.direccion "
 csql.Execute
 Grid2.Rows = 1
 If csql.RowsAffected > 0 Then
    Set resultados = csql.OpenResultset
    While resultados.EOF = False
    Grid2.Rows = Grid2.Rows + 1
    For k = 1 To 3
    Grid2.Cell(Grid2.Rows - 1, k).text = resultados(k - 1)
    Next k
    If IsNull(resultados(3)) = False Then
    Grid2.Cell(Grid2.Rows - 1, 4).text = resultados(3)
    Grid2.Cell(Grid2.Rows - 1, 5).text = leearrendatario(resultados(4))
    Grid2.Cell(Grid2.Rows - 1, 6).text = resultados(5)
    Grid2.Cell(Grid2.Rows - 1, 7).text = resultados(6)
    Grid2.Cell(Grid2.Rows - 1, 8).text = resultados(7)
    Grid2.Cell(Grid2.Rows - 1, 9).text = leemonedas(resultados(8))
    Grid2.Cell(Grid2.Rows - 1, 10).text = resultados(9)
    Grid2.Cell(Grid2.Rows - 1, 11).text = arriendoatrasado(resultados(3), Format(fechasistema, "yyyy-mm-dd"))
    If Format(resultados(6), "yyyy-mm-dd") < Format(fechasistema, "yyyy-mm-dd") Then
    Grid2.Range(Grid2.Rows - 1, 1, Grid2.Rows - 1, 11).BackColor = &HFF&
    
    End If
    
    
    
    Else
    
    Grid2.Cell(Grid2.Rows - 1, 5).text = "*** DISPONIBLE **"
    
    End If
    
    resultados.MoveNext
    
    
    
    Wend
    
    
  End If
 csql.Close
 Set csql = Nothing
 Set resultados = Nothing
 
 End Sub

Private Function leemonedas(codigo) As String
Dim csql As New rdoQuery
Dim resultados As rdoResultset

Set csql.ActiveConnection = contadb

csql.sql = "select nombremoneda from " & clientesistema & "arriendos" & ".maestro_monedas where codigomoneda='" & codigo & "'"
csql.Execute
leemonedas = ""
If csql.RowsAffected > 0 Then
Set resultados = csql.OpenResultset
leemonedas = resultados(0)
End If
Set resultados = Nothing
csql.Close
Set csql = Nothing

End Function

Public Function LEERULTIMOFOLIOcontrato() As String
    Dim numero As Double
    
    
    Dim resultados As rdoResultset
    Dim csql As New rdoQuery
    
        Set csql.ActiveConnection = contadb

            csql.sql = "select IFNULL(max(numero)+1,1) from " + clientesistema + "arriendos.arriendos_pago_cabeza"
            
            csql.Execute
    If csql.RowsAffected > 0 Then
    Set resultados = csql.OpenResultset
        numero = resultados(0)
    
        LEERULTIMOFOLIOcontrato = Format(numero, "0000000000")
    End If
    
End Function


Public Sub leermensualidades(numero)
 Dim csql As New rdoQuery
 Dim resultados As rdoResultset
 
 Set csql.ActiveConnection = contadb
 csql.sql = "select * from " & clientesistema & "arriendos" & ".arriendos_mensuales as mp where numero='" + numero + "' and pagado='0' "
 csql.sql = csql.sql + "order by fecha "
 
 csql.Execute
 Grid1.Rows = 1
 If csql.RowsAffected > 0 Then
    Set resultados = csql.OpenResultset
    While resultados.EOF = False
    Grid1.Rows = Grid1.Rows + 1
    Grid1.Cell(Grid1.Rows - 1, 1).text = resultados(1)
    Grid1.Cell(Grid1.Rows - 1, 2).text = resultados(2)
    Grid1.Cell(Grid1.Rows - 1, 3).text = leemonedas(resultados(3))
    Grid1.Cell(Grid1.Rows - 1, 4).text = resultados(4)
    Grid1.Cell(Grid1.Rows - 1, 5).text = "0"
    resultados.MoveNext
    
    Wend
    
    
  End If
 csql.Close
 Set csql = Nothing
 Set resultados = Nothing
 
 End Sub

Public Sub leermensualidadespagadas(numero)
 Dim csql As New rdoQuery
 Dim resultados As rdoResultset
 
 Set csql.ActiveConnection = contadb
 csql.sql = "select * from " & clientesistema & "arriendos" & ".arriendos_mensuales as mp where comprobante_pago='" + numero + "' and pagado='1' "
 csql.sql = csql.sql + "order by fecha "
 
 csql.Execute
 Grid1.Rows = 1
 If csql.RowsAffected > 0 Then
    Set resultados = csql.OpenResultset
    While resultados.EOF = False
    Grid1.Rows = Grid1.Rows + 1
    Grid1.Cell(Grid1.Rows - 1, 1).text = resultados(1)
    Grid1.Cell(Grid1.Rows - 1, 2).text = resultados(2)
    Grid1.Cell(Grid1.Rows - 1, 3).text = leemonedas(resultados(3))
    Grid1.Cell(Grid1.Rows - 1, 4).text = resultados(4)
    Grid1.Cell(Grid1.Rows - 1, 5).text = "0"
    resultados.MoveNext
    
    Wend
    
    
  End If
 csql.Close
 Set csql = Nothing
 Set resultados = Nothing
 
 End Sub

Sub modificapago(contrato, fecha, pagado, numero)
    fecha = Format(fecha, "yyyy-mm-dd")
    campos(0, 0) = "pagado"
    campos(1, 0) = "comprobante_pago"
    campos(2, 0) = ""
    campos(0, 1) = pagado
    campos(1, 1) = numero
    
    campos(0, 2) = clientesistema & "arriendos" & ".arriendos_mensuales"
    condicion = "numero='" + contrato + "' and fecha='" + fecha + "' "
    op = 3
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)
    
    End Sub

 
Private Sub Option1_Click()
Call leermensualidades(dato1.text)

End Sub

Private Sub Option2_Click()
Call leermensualidades(dato1.text)

End Sub

Public Function existecontrato(numero) As Boolean


    Dim resultados As rdoResultset
    Dim csql As New rdoQuery
    
        Set csql.ActiveConnection = contadb

            csql.sql = "select rutarrendatario,propiedad,monedaarriendo from " + clientesistema + "arriendos.contratos_arriendo where numero='" + numero + "' "
            
            csql.Execute
    existecontrato = False
    
    If csql.RowsAffected > 0 Then
    Set resultados = csql.OpenResultset
    txtrut = Mid(resultados(0), 1, 9)
    dv2.Caption = Mid(resultados(0), 10, 1)
    txtpropiedad.text = resultados(1)
    lblpropiedad.Caption = leepropiedad(txtpropiedad.text)
    lblarrendatario.Caption = leearrendatario(resultados(0))
    txttipo.text = resultados(2)
    LBLMONEDA.Caption = leemonedas(resultados(2))
    existecontrato = True
    
    
    
    End If
    
End Function

Public Sub leerdetallepago(numero)
 Dim csql As New rdoQuery
 Dim resultados As rdoResultset
 
 Set csql.ActiveConnection = contadb
 csql.sql = "select * from " & clientesistema & "arriendos" & ".arriendos_pago_detalle where numero='" + numero + "' "
 csql.sql = csql.sql + "order by linea "
 
 csql.Execute
 Grid2.Rows = 1
 If csql.RowsAffected > 0 Then
    Set resultados = csql.OpenResultset
    While resultados.EOF = False
    Grid2.Rows = Grid2.Rows + 1
    Grid2.Cell(Grid2.Rows - 1, 1).text = resultados(2)
    Grid2.Cell(Grid2.Rows - 1, 2).text = ""
    Grid2.Cell(Grid2.Rows - 1, 3).text = resultados(3)
    Grid2.Cell(Grid2.Rows - 1, 4).text = resultados(4)
    Grid2.Cell(Grid2.Rows - 1, 5).text = resultados(5)
    If IsNull(resultados(6)) = False Then
    Grid2.Cell(Grid2.Rows - 1, 6).text = resultados(6)
    End If
    resultados.MoveNext
    
    Wend
    
    
  End If
 csql.Close
 Set csql = Nothing
 Set resultados = Nothing
 frmdetalle.Enabled = False
 
 End Sub

Sub imprimir()
Dim k As Double
Dim i As Double

Call CARGAGRILLA4
Call CABEZAS2("COMPROBANTE DE PAGO NUMERO " + dato1.text)

Grid4.Rows = 13

For k = 1 To Grid2.Rows - 1
    Grid4.Cell(k, 1).text = Grid2.Cell(k, 1).text
    Grid4.Cell(k, 2).text = Grid2.Cell(k, 2).text
    Grid4.Cell(k, 3).text = Grid2.Cell(k, 3).text
    Grid4.Cell(k, 4).text = Grid2.Cell(k, 4).text
    Grid4.Cell(k, 5).text = Grid2.Cell(k, 5).text
    Grid4.Cell(k, 6).text = Grid2.Cell(k, 6).text
    Grid4.Cell(k, 7).text = Grid2.Cell(k, 7).text
 
Next k

For k = 1 To Grid1.Rows - 1
    Grid4.Cell(k, 8).text = Grid1.Cell(k, 1).text
    Grid4.Cell(k, 9).text = Grid1.Cell(k, 2).text
    Grid4.Cell(k, 10).text = Grid1.Cell(k, 4).text
    Grid4.Cell(k, 11).text = CDbl(Grid1.Cell(k, 2).text) + CDbl(Grid1.Cell(k, 4).text)
Next k

    Grid4.Range(0, 1, 0, Grid4.Cols - 1).Borders(cellEdgeLeft) = cellThick
    Grid4.Range(0, 1, 0, Grid4.Cols - 1).Borders(cellEdgeTop) = cellThick
    Grid4.Range(0, 1, 0, Grid4.Cols - 1).Borders(cellEdgeBottom) = cellThick
    Grid4.Range(0, 1, 0, Grid4.Cols - 1).Borders(cellEdgeRight) = cellThick
    Grid4.Range(0, 1, 0, Grid4.Cols - 1).Borders(cellInsideHorizontal) = cellThick
    Grid4.Range(0, 1, 0, Grid4.Cols - 1).Borders(cellInsideVertical) = cellThick

    Grid4.Range(0, 1, Grid4.Rows - 1, Grid2.Cols - 1).Borders(cellEdgeLeft) = cellThick
    Grid4.Range(0, 1, Grid4.Rows - 1, Grid2.Cols - 1).Borders(cellEdgeTop) = cellThick
    Grid4.Range(0, 1, Grid4.Rows - 1, Grid2.Cols - 1).Borders(cellEdgeBottom) = cellThick
    Grid4.Range(0, 1, Grid4.Rows - 1, Grid2.Cols - 1).Borders(cellEdgeRight) = cellThick


    Grid4.Range(0, 8, Grid4.Rows - 1, Grid4.Cols - 1).Borders(cellEdgeLeft) = cellThick
    Grid4.Range(0, 8, Grid4.Rows - 1, Grid4.Cols - 1).Borders(cellEdgeTop) = cellThick
    Grid4.Range(0, 8, Grid4.Rows - 1, Grid4.Cols - 1).Borders(cellEdgeBottom) = cellThick
    Grid4.Range(0, 8, Grid4.Rows - 1, Grid4.Cols - 1).Borders(cellEdgeRight) = cellThick


    Grid4.Rows = Grid4.Rows + 2
    Grid4.Column(1).Locked = False
    Grid4.Column(2).Locked = False
    Grid4.Column(3).Locked = False
    Grid4.Column(4).Locked = False
    Grid4.Column(5).Locked = False
    Grid4.Column(6).Locked = False
    Grid4.Column(7).Locked = False
    Grid4.Column(8).Locked = False
    Grid4.Column(9).Locked = False
    Grid4.Column(10).Locked = False
    Grid4.Column(11).Locked = False
  
    
    
    Grid4.Range(Grid4.Rows - 1, 1, Grid4.Rows - 1, Grid4.Cols - 1).Merge
    Grid4.Range(Grid4.Rows - 1, 1, Grid4.Rows - 1, Grid4.Cols - 1).Alignment = cellCenterCenter
    Grid4.Range(Grid4.Rows - 1, 1, Grid4.Rows - 1, Grid4.Cols - 1).FontSize = 8
    Grid4.Range(Grid4.Rows - 1, 1, Grid4.Rows - 1, Grid4.Cols - 1).FontBold = True
    
    Grid4.Cell(Grid4.Rows - 1, 1).text = "ESTE COMPROBANTE ES VALIDO COMO RECIBO DE DINERO POR CONCEPTO DE PAGO DE ARRIENDOS DETALLADOS"
    Grid4.Rows = Grid4.Rows + 10
    
    Grid4.Range(Grid4.Rows - 1, 2, Grid4.Rows - 1, 4).Merge
    Grid4.Range(Grid4.Rows - 1, 2, Grid4.Rows - 1, 4).Alignment = cellCenterCenter
    Grid4.Range(Grid4.Rows - 1, 2, Grid4.Rows - 1, 4).FontSize = 8
    Grid4.Range(Grid4.Rows - 1, 2, Grid4.Rows - 1, 4).FontBold = True
    Grid4.Range(Grid4.Rows - 1, 2, Grid4.Rows - 1, 4).Borders(cellEdgeTop) = cellThick
    Grid4.Cell(Grid4.Rows - 1, 2).text = "FIRMA DEPTO. COBRANZA"
   
    Grid4.PrintPreview


End Sub
Sub CABEZAS2(titulo)
Dim objReportTitle As FlexCell.ReportTitle
Grid4.ReportTitles.Clear


    'Report Title 1
   Call leerdatosarrendatario(txtpropiedad.text)
        For k = 1 To 4
        Set objReportTitle = New FlexCell.ReportTitle
        objReportTitle.text = DATOSARRENDATARIO(k)
        objReportTitle.Font.Name = "arial"
        objReportTitle.Font.Size = 8
        objReportTitle.Font.Italic = True
        objReportTitle.PrintOnAllPages = True
        'objReportTitle.Color = RGB(128, 0, 0)
        objReportTitle.Align = CellLeft
        Grid4.ReportTitles.Add objReportTitle
    Next k
    
Set objReportTitle = New FlexCell.ReportTitle
    objReportTitle.text = titulo
    objReportTitle.Font.Name = "arial"
    objReportTitle.Font.Size = 12
    objReportTitle.PrintOnAllPages = True
    Grid4.ReportTitles.Add objReportTitle

    
Set objReportTitle = New FlexCell.ReportTitle
    objReportTitle.text = ""
    objReportTitle.Font.Name = "arial"
    objReportTitle.Font.Size = 12
    objReportTitle.PrintOnAllPages = True
    Grid4.ReportTitles.Add objReportTitle

Set objReportTitle = New FlexCell.ReportTitle
    objReportTitle.text = "FECHA :" + dato2.text + "-" + dato3.text + "-" + dato4.text
    objReportTitle.Font.Name = "arial"
    objReportTitle.Font.Size = 10
    objReportTitle.PrintOnAllPages = True
    objReportTitle.Align = CellLeft
    Grid4.ReportTitles.Add objReportTitle

Set objReportTitle = New FlexCell.ReportTitle
    objReportTitle.text = "CONTRATO :" + DATO5.text
    objReportTitle.Font.Name = "arial"
    objReportTitle.Font.Size = 10
    objReportTitle.PrintOnAllPages = True
    objReportTitle.Align = CellLeft
    
    Grid4.ReportTitles.Add objReportTitle
    
Set objReportTitle = New FlexCell.ReportTitle
    objReportTitle.text = "ARRENDATARIO :" + txtrut.text + "-" + dv2.Caption + "  " + lblarrendatario.Caption
    objReportTitle.Font.Name = "arial"
    objReportTitle.Font.Size = 10
    objReportTitle.PrintOnAllPages = True
    objReportTitle.Align = CellLeft
    
    Grid4.ReportTitles.Add objReportTitle

Set objReportTitle = New FlexCell.ReportTitle
    objReportTitle.text = "PROPIEDAD :" + txtpropiedad.text + "  " + lblpropiedad.Caption
    objReportTitle.Font.Name = "arial"
    objReportTitle.Font.Size = 10
    objReportTitle.PrintOnAllPages = True
    objReportTitle.Align = CellLeft
    
    Grid4.ReportTitles.Add objReportTitle

Set objReportTitle = New FlexCell.ReportTitle
    objReportTitle.text = "MONEDA :" + txttipo.text + "  " + LBLMONEDA.Caption
    objReportTitle.Font.Name = "arial"
    objReportTitle.Font.Size = 10
    objReportTitle.PrintOnAllPages = True
    objReportTitle.Align = CellLeft
    
    Grid4.ReportTitles.Add objReportTitle

Set objReportTitle = New FlexCell.ReportTitle
    objReportTitle.text = "CANCELADO :" + Format(dato6.text, " $ ###,###,###,###")
    objReportTitle.Font.Name = "arial"
    objReportTitle.Font.Size = 10
    objReportTitle.PrintOnAllPages = True
    objReportTitle.Align = CellLeft
    
    Grid4.ReportTitles.Add objReportTitle
    
Set objReportTitle = New FlexCell.ReportTitle
    objReportTitle.text = ""
    objReportTitle.Font.Name = "arial"
    objReportTitle.Font.Size = 10
    objReportTitle.PrintOnAllPages = True
    objReportTitle.Align = CellLeft
    Grid4.ReportTitles.Add objReportTitle
    
Set objReportTitle = New FlexCell.ReportTitle
    objReportTitle.text = "DETALLE DEL PAGO                                                         MESES PAGADOS        "
    objReportTitle.Font.Name = "arial"
    objReportTitle.Font.Size = 10
    objReportTitle.Font.Bold = True
    objReportTitle.PrintOnAllPages = True
    objReportTitle.Align = cellCenter
    
    Grid4.ReportTitles.Add objReportTitle

Set objReportTitle = New FlexCell.ReportTitle
    objReportTitle.text = ""
    objReportTitle.Font.Name = "arial"
    objReportTitle.Font.Size = 10
    objReportTitle.PrintOnAllPages = True
    objReportTitle.Align = CellLeft
    Grid4.ReportTitles.Add objReportTitle



With Grid4.PageSetup
        
        Rem If TIPO = "S" Then .Footer = "pagina &P"
        
        .HeaderAlignment = cellCenter
        .HeaderFont.Name = "Verdana"
        .HeaderFont.Size = 7
        .HeaderMargin = 2
        .TopMargin = 1
        .BottomMargin = 2
        .LeftMargin = 1
        .RightMargin = 0
        .PrintFixedRow = True
        .BlackAndWhite = True
        
End With

End Sub

Sub CARGAGRILLA4()
    Dim formatogrilla2(10, 12)
    formatogrilla2(1, 1) = "TIPO"
    formatogrilla2(1, 2) = "NOMBRE"
    formatogrilla2(1, 3) = "MONTO"
    formatogrilla2(1, 4) = "BANCO"
    formatogrilla2(1, 5) = "CUENTA"
    formatogrilla2(1, 6) = "NUMERO"
    formatogrilla2(1, 7) = "VENC"
    formatogrilla2(1, 8) = "F.PAGADOS"
    formatogrilla2(1, 9) = "MONTO"
    formatogrilla2(1, 10) = "G/COMUNES"
    formatogrilla2(1, 11) = "TOTAL"

    
    Rem LARGO DE LOS DATOS
    formatogrilla2(2, 1) = "4"
    formatogrilla2(2, 2) = "5"
    formatogrilla2(2, 3) = "8"
    formatogrilla2(2, 4) = "5"
    formatogrilla2(2, 5) = "8"
    formatogrilla2(2, 6) = "8"
    formatogrilla2(2, 7) = "8"
    formatogrilla2(2, 8) = "8"
    formatogrilla2(2, 9) = "8"
    formatogrilla2(2, 10) = "8"
    formatogrilla2(2, 11) = "8"
    
    Rem TIPO DE DATOS C=CEROS,S=STRING,N=NUMERICO,D=DATE,
    formatogrilla2(3, 1) = "N"
    formatogrilla2(3, 2) = "S"
    formatogrilla2(3, 3) = "N"
    formatogrilla2(3, 4) = "S"
    formatogrilla2(3, 5) = "S"
    formatogrilla2(3, 6) = "S"
    formatogrilla2(3, 7) = "D"
    formatogrilla2(3, 8) = "D"
    formatogrilla2(3, 9) = "N"
    formatogrilla2(3, 10) = "N"
    formatogrilla2(3, 11) = "N"
    Rem FORMATO GRILLA
    
    formatogrilla2(4, 3) = " ###,###,##0.00"
    formatogrilla2(4, 9) = "$ ###,###,##0"
    formatogrilla2(4, 10) = "$ ###,###,##0"
    formatogrilla2(4, 11) = "$ ###,###,##0"
    
    
    Rem LOCCKED
    formatogrilla2(5, 1) = "FALSE"
    formatogrilla2(5, 2) = "TRUE"
    formatogrilla2(5, 3) = "FALSE"
    formatogrilla2(5, 4) = "FALSE"
    formatogrilla2(5, 5) = "FALSE"
    formatogrilla2(5, 6) = "FALSE"
    formatogrilla2(5, 7) = "FALSE"
    formatogrilla2(5, 8) = "FALSE"
    formatogrilla2(5, 9) = "FALSE"
    formatogrilla2(5, 10) = "FALSE"
    formatogrilla2(5, 11) = "FALSE"
    
    Rem VALOR MAXIMO
    
    Grid4.Cols = 12
    Grid4.Rows = 2
    Grid4.AllowUserResizing = False
    Grid4.DisplayFocusRect = False
    Grid4.ExtendLastCol = True
    Grid4.BoldFixedCell = False
    Grid4.DrawMode = cellOwnerDraw
    Grid4.Appearance = Flat
    Grid4.ScrollBarStyle = Flat
    Grid4.FixedRowColStyle = Flat
    Grid4.BackColorFixed = RGB(90, 158, 214)
    Grid4.BackColorFixedSel = RGB(110, 180, 230)
    Grid4.BackColorBkg = RGB(90, 158, 214)
    Grid4.BackColorScrollBar = RGB(231, 235, 247)
    Grid4.BackColor1 = RGB(231, 235, 247)
    Grid4.BackColor2 = RGB(239, 243, 255)
    Grid4.GridColor = RGB(148, 190, 231)
    Grid4.Column(0).Width = 0
    
    For k = 1 To Grid4.Cols - 1
        Grid4.Cell(0, k).text = formatogrilla2(1, k)
        Grid4.Column(k).Width = Val(formatogrilla2(2, k)) * 9
        Grid4.Column(k).MaxLength = Val(formatogrilla2(2, k))
        Grid4.Column(k).FormatString = formatogrilla2(4, k)
        Grid4.Column(k).Locked = formatogrilla2(5, k)
        If formatogrilla2(3, k) = "N" Then Grid4.Column(k).Alignment = cellRightCenter
        If formatogrilla2(3, k) = "S" Then Grid4.Column(k).Alignment = cellLeftTop
        
        
        If formatogrilla2(3, k) = "D" Then Grid4.Column(k).CellType = cellCalendar
        
    Next k
    
    
    Grid4.Column(3).Mask = cellNumeric
    
    
    End Sub
Private Sub botonmisaccesos_Click()
programafiltro = Me.Caption
misaccesos.Show
End Sub


Private Sub botonmisfavoritos_Click()
Call AgregaFavoritos(USUARIOSISTEMA, App.EXEName, Me.Name, Me.Caption)
End Sub
