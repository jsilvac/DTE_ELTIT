VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash9b.ocx"
Object = "{ADBBAED6-B16F-40DC-80DF-B44910CBA76C}#1.0#0"; "Frame-Xp.ocx"
Begin VB.Form FactManCob 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Facturas Manuales a Cobranza"
   ClientHeight    =   5130
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9135
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5130
   ScaleWidth      =   9135
   Begin XPFrame.FrameXp FrameXp1 
      Height          =   3465
      Left            =   120
      TabIndex        =   15
      Top             =   120
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   6112
      BackColor       =   16744576
      Caption         =   "Datos Documento"
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
      Begin XPFrame.FrameXp frmTipo 
         Height          =   1635
         Left            =   1860
         TabIndex        =   40
         Top             =   420
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
            TabIndex        =   43
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
            TabIndex        =   42
            Top             =   840
            Width           =   2475
         End
         Begin VB.Label lbl25 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0FF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " 4 - Nota Crédito"
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
            TabIndex        =   41
            Top             =   1200
            Width           =   2475
         End
      End
      Begin VB.TextBox dato0 
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
         MaxLength       =   10
         TabIndex        =   0
         Tag             =   "proveedor"
         Top             =   420
         Width           =   360
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
         Left            =   4650
         MaxLength       =   1
         TabIndex        =   9
         Tag             =   "proveedor"
         Top             =   1140
         Width           =   390
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
         Left            =   1440
         MaxLength       =   60
         TabIndex        =   13
         Tag             =   "proveedor"
         Top             =   2640
         Width           =   7380
      End
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
         Left            =   6240
         MaxLength       =   10
         TabIndex        =   1
         Tag             =   "proveedor"
         Top             =   420
         Width           =   1380
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
         Left            =   1440
         MaxLength       =   2
         TabIndex        =   2
         Tag             =   "proveedor"
         Top             =   780
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
         Left            =   1860
         MaxLength       =   2
         TabIndex        =   3
         Tag             =   "proveedor"
         Top             =   780
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
         Left            =   2280
         MaxLength       =   4
         TabIndex        =   4
         Tag             =   "proveedor"
         Top             =   780
         Width           =   675
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
         Left            =   6240
         MaxLength       =   2
         TabIndex        =   5
         Tag             =   "proveedor"
         Top             =   780
         Width           =   375
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
         Left            =   6660
         MaxLength       =   2
         TabIndex        =   6
         Tag             =   "proveedor"
         Top             =   780
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
         Left            =   7080
         MaxLength       =   4
         TabIndex        =   7
         Tag             =   "proveedor"
         Top             =   780
         Width           =   675
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
         Left            =   1440
         MaxLength       =   9
         TabIndex        =   8
         Tag             =   "proveedor"
         Top             =   1140
         Width           =   1380
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
         MaxLength       =   2
         TabIndex        =   10
         Tag             =   "proveedor"
         Top             =   1920
         Width           =   1380
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
         MaxLength       =   9
         TabIndex        =   11
         Tag             =   "proveedor"
         Top             =   2280
         Width           =   1380
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
         Left            =   4440
         MaxLength       =   9
         TabIndex        =   12
         Tag             =   "proveedor"
         Top             =   2280
         Width           =   1380
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
         Left            =   1860
         TabIndex        =   39
         Top             =   420
         Width           =   2835
      End
      Begin VB.Label lbl0 
         Appearance      =   0  'Flat
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Tipo Doc."
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   120
         TabIndex        =   38
         Top             =   420
         Width           =   1215
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Sucursal"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   3330
         TabIndex        =   36
         Top             =   1140
         Width           =   1215
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
         Left            =   2790
         TabIndex        =   35
         Top             =   1920
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label lbl1 
         Appearance      =   0  'Flat
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Nº Factura"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   4920
         TabIndex        =   34
         Top             =   420
         Width           =   1215
      End
      Begin VB.Label lbl3 
         Appearance      =   0  'Flat
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Fecha Vcto."
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   4920
         TabIndex        =   33
         Top             =   780
         Width           =   1215
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
         Top             =   1140
         Width           =   1215
      End
      Begin VB.Label lbl5 
         Appearance      =   0  'Flat
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Vendedor"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   120
         TabIndex        =   31
         Top             =   1920
         Width           =   1215
      End
      Begin VB.Label lbl2 
         Appearance      =   0  'Flat
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Fecha Emisión"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   120
         TabIndex        =   30
         Top             =   780
         Width           =   1215
      End
      Begin VB.Label lbl6 
         Appearance      =   0  'Flat
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Monto"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   120
         TabIndex        =   29
         Top             =   2280
         Width           =   1215
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
         ForeColor       =   &H0080FFFF&
         Height          =   285
         Left            =   120
         TabIndex        =   28
         Top             =   1500
         Width           =   8655
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
         Left            =   2955
         TabIndex        =   27
         Top             =   1920
         Width           =   5820
      End
      Begin VB.Label lblDVC 
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
         Left            =   2880
         TabIndex        =   26
         Top             =   1140
         Width           =   315
      End
      Begin VB.Label lbl7 
         Appearance      =   0  'Flat
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Abono"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   3120
         TabIndex        =   25
         Top             =   2280
         Width           =   1215
      End
      Begin VB.Label lbl8 
         Appearance      =   0  'Flat
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Pendiente"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   6120
         TabIndex        =   24
         Top             =   2280
         Width           =   1215
      End
      Begin VB.Label lbl9 
         Appearance      =   0  'Flat
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Observaciones"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   120
         TabIndex        =   23
         Top             =   2640
         Width           =   1215
      End
      Begin VB.Label lblPendiente 
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
         Left            =   7440
         TabIndex        =   22
         Top             =   2280
         Width           =   1335
      End
      Begin VB.Label lbl10 
         Appearance      =   0  'Flat
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Liquidada"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   120
         TabIndex        =   21
         Top             =   3000
         Width           =   1215
      End
      Begin VB.Label lblLiquidada 
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
         TabIndex        =   20
         Top             =   3000
         Width           =   375
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
         Left            =   3360
         TabIndex        =   19
         Top             =   3000
         Width           =   375
      End
      Begin VB.Label lbl11 
         Appearance      =   0  'Flat
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Fecha Liq."
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   2040
         TabIndex        =   18
         Top             =   3000
         Width           =   1215
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
         Left            =   3780
         TabIndex        =   17
         Top             =   3000
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
         Left            =   4200
         TabIndex        =   16
         Top             =   3000
         Width           =   795
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
      TabIndex        =   37
      Top             =   0
      Width           =   555
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash opciones 
      Height          =   1335
      Left            =   180
      TabIndex        =   14
      Top             =   3780
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
Attribute VB_Name = "FactManCob"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Private modifica As Boolean
    Private f As factManuales

'============================================================
'MANEJO DE LOS CONTOLES
'============================================================
    '========================================================
    'GotFocus
    '========================================================
    Private Sub dato0_GotFocus()
        frmTipo.Visible = True
        Call VerificarCajas(Me, dato0)
        Call selecciona(dato0)
    End Sub
    
    Private Sub dato1_GotFocus()
        Call VerificarCajas(Me, dato1)
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
    End Sub
    
    Private Sub dato7_GotFocus()
        Call VerificarCajas(Me, dato7)
        Call selecciona(dato7)
    End Sub
    
    Private Sub dato8_GotFocus()
        Call VerificarCajas(Me, dato8)
        Call selecciona(dato8)
        Principal.barraEstado.Panels(2).text = "F2: Ayuda Cliente"
    End Sub
    
    Private Sub dato9_GotFocus()
        Call VerificarCajas(Me, dato9)
        Call selecciona(dato9)
    End Sub
    
    Private Sub dato10_GotFocus()
        Call VerificarCajas(Me, dato10)
        Call selecciona(dato10)
        Principal.barraEstado.Panels(2).text = "F2: Ayuda Vendedor"
    End Sub
    
    Private Sub dato11_GotFocus()
        Call VerificarCajas(Me, dato11)
        Call selecciona(dato11)
    End Sub
    
    Private Sub dato12_GotFocus()
        Call VerificarCajas(Me, dato12)
        Call selecciona(dato12)
    End Sub
    
    Private Sub dato13_GotFocus()
        Call VerificarCajas(Me, dato13)
        Call selecciona(dato13)
    End Sub
    '========================================================
    'GotFocus
    '========================================================
    
    '========================================================
    'KeyDown
    '========================================================
    Private Sub dato0_KeyDown(KeyCode As Integer, Shift As Integer)
        Select Case KeyCode
            Case 97, 49
                dato0.text = "FV"
            Case 98, 50
                dato0.text = "BV"
            Case 100, 52
                dato0.text = "NV"
            Case Else
                Call Flechas(KeyCode, dato0)
        End Select
    End Sub
    
    Private Sub dato1_KeyDown(KeyCode As Integer, Shift As Integer)
        Call Flechas(KeyCode, dato0)
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
        Call Flechas(KeyCode, dato5)
    End Sub
    
    Private Sub dato7_KeyDown(KeyCode As Integer, Shift As Integer)
        Call Flechas(KeyCode, dato6)
    End Sub
    
    Private Sub dato8_KeyDown(KeyCode As Integer, Shift As Integer)
        If KeyCode = vbKeyF2 Then
            Call ayudaCliente(dato8, dato9, lblDVC)
        Else
            Call Flechas(KeyCode, dato7)
        End If
    End Sub
    
    Private Sub dato9_KeyDown(KeyCode As Integer, Shift As Integer)
        Call Flechas(KeyCode, dato8)
    End Sub
    
    Private Sub dato10_KeyDown(KeyCode As Integer, Shift As Integer)
        If KeyCode = vbKeyF2 Then
            Call ayudaVendedores(dato9)
        Else
            Call Flechas(KeyCode, dato8)
        End If
    End Sub
    
    Private Sub dato11_KeyDown(KeyCode As Integer, Shift As Integer)
        Call Flechas(KeyCode, dato9)
    End Sub
    
    Private Sub dato12_KeyDown(KeyCode As Integer, Shift As Integer)
        Call Flechas(KeyCode, dato10)
    End Sub
    
    Private Sub dato13_KeyDown(KeyCode As Integer, Shift As Integer)
        Call Flechas(KeyCode, dato11)
    End Sub
    '========================================================
    'KeyDown
    '========================================================
    
    '========================================================
    'KeyPress
    '========================================================
    Private Sub dato0_KeyPress(KeyAscii As Integer)
        KeyAscii = esNumero(KeyAscii)
        If KeyAscii = 13 Then
            dato1.text = ceros(dato1)
            If leerFactManual(f, dato1.text, "=") = True Then
                Call structtoctrl
                opciones.SetFocus
            Else
                Call HabilitarCajas(Me, modifica)
                SendKeys "{Tab}"
            End If
        End If
    End Sub
    
    Private Sub dato1_KeyPress(KeyAscii As Integer)
        KeyAscii = esNumero(KeyAscii)
        If KeyAscii = 13 Then
            dato1.text = ceros(dato1)
            If leerFactManual(f, dato1.text, "=") = True Then
                Call structtoctrl
                opciones.SetFocus
            Else
                Call HabilitarCajas(Me, modifica)
                SendKeys "{Tab}"
            End If
        End If
    End Sub
    
    Private Sub dato2_KeyPress(KeyAscii As Integer)
        KeyAscii = esNumero(KeyAscii)
        If KeyAscii = 13 Then
            dato2.text = ceros(dato2)
            If dato2.text = "00" Then
                dato2.text = Format(fechasistema, "dd")
            End If
            SendKeys "{Tab}"
        End If
    End Sub
    
    Private Sub dato3_KeyPress(KeyAscii As Integer)
        KeyAscii = esNumero(KeyAscii)
        If KeyAscii = 13 Then
            dato3.text = ceros(dato3)
            If dato3.text = "00" Then
                dato3.text = Format(fechasistema, "mm")
            End If
            SendKeys "{Tab}"
        End If
    End Sub
    
    Private Sub dato4_KeyPress(KeyAscii As Integer)
        KeyAscii = esNumero(KeyAscii)
        If KeyAscii = 13 Then
            dato4.text = ceros(dato4)
            If dato4.text = "0000" Then
                dato4.text = Format(fechasistema, "yyyy")
            End If
            SendKeys "{Tab}"
        End If
    End Sub
    
    Private Sub dato5_KeyPress(KeyAscii As Integer)
        KeyAscii = esNumero(KeyAscii)
        If KeyAscii = 13 Then
            dato5.text = ceros(dato5)
            If dato5.text = "00" Then
                dato5.text = Format(fechasistema, "dd")
            End If
            SendKeys "{Tab}"
        End If
    End Sub
    
    Private Sub dato6_KeyPress(KeyAscii As Integer)
        KeyAscii = esNumero(KeyAscii)
        If KeyAscii = 13 Then
            dato6.text = ceros(dato6)
            If dato6.text = "00" Then
                dato6.text = Format(fechasistema, "mm")
            End If
            SendKeys "{Tab}"
        End If
    End Sub
    
    Private Sub dato7_KeyPress(KeyAscii As Integer)
        KeyAscii = esNumero(KeyAscii)
        If KeyAscii = 13 Then
            dato7.text = ceros(dato7)
            If dato7.text = "0000" Then
                dato7.text = Format(fechasistema, "yyyy")
            End If
            SendKeys "{Tab}"
        End If
    End Sub
    
    Private Sub dato8_KeyPress(KeyAscii As Integer)
        KeyAscii = esNumero(KeyAscii)
        If KeyAscii = 13 Then
            dato8.text = ceros(dato8)
            lblDVC.Caption = rut(dato8.text)
            SendKeys "{Tab}"
        End If
    End Sub
    
    Private Sub dato9_KeyPress(KeyAscii As Integer)
        KeyAscii = esNumero(KeyAscii)
        If KeyAscii = 13 Then
            dato9.text = ceros(dato9)
            lblNombre.Caption = leerNombreClienteSucursal(dato8.text & lblDVC.Caption, dato9.text)
            If lblNombre.Caption <> "" Then
                SendKeys "{Tab}"
            End If
        End If
    End Sub
    
    Private Sub dato10_KeyPress(KeyAscii As Integer)
        KeyAscii = esNumero(KeyAscii)
        If KeyAscii = 13 Then
            dato10.text = ceros(dato10)
            lblDVV.Caption = rut(dato10.text)
            lblVendedor.Caption = leerNombreVendedor(dato10.text)
            If lblVendedor.Caption <> "" Then
                SendKeys "{Tab}"
            End If
        End If
    End Sub
    
    Private Sub dato11_KeyPress(KeyAscii As Integer)
        KeyAscii = esNumero(KeyAscii)
        If KeyAscii = 13 Then
            SendKeys "{Tab}"
        End If
        If dato11.text = "" Then
            dato11.text = "0"
        End If
    End Sub
    
    Private Sub dato12_KeyPress(KeyAscii As Integer)
        KeyAscii = esNumero(KeyAscii)
        If KeyAscii = 13 Then
            SendKeys "{Tab}"
        End If
        If dato12.text = "" Then
            dato12.text = "0"
        End If
    End Sub
    
    Private Sub dato13_KeyPress(KeyAscii As Integer)
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If KeyAscii = 13 Then
            Call ctrltostruct
        End If
    End Sub
    '========================================================
    'KeyPress
    '========================================================
    
    '========================================================
    'KeyUp
    '========================================================
'    Private Sub dato2_KeyUp(KeyCode As Integer, Shift As Integer)
'        If Len(dato2.text) = dato2.MaxLength Then
'            Call dato2_KeyPress(13)
'        End If
'    End Sub
'
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
'    Private Sub dato6_KeyUp(KeyCode As Integer, Shift As Integer)
'        If Len(dato6.text) = dato6.MaxLength Then
'            Call dato6_KeyPress(13)
'        End If
'    End Sub
'
'    Private Sub dato7_KeyUp(KeyCode As Integer, Shift As Integer)
'        If Len(dato7.text) = dato7.MaxLength Then
'            Call dato7_KeyPress(13)
'        End If
'    End Sub
    
    Private Sub dato11_KeyUp(KeyCode As Integer, Shift As Integer)
        If dato11.text = "" Then
            dato11.text = "0"
        End If
        If dato12.text = "" Then
            dato12.text = "0"
        End If
        lblPendiente.Caption = CDbl(dato11.text) - CDbl(dato12.text)
    End Sub
    
    Private Sub dato12_KeyUp(KeyCode As Integer, Shift As Integer)
        If dato11.text = "" Then
            dato11.text = "0"
        End If
        If dato12.text = "" Then
            dato12.text = "0"
        End If
        lblPendiente.Caption = CDbl(dato11.text) - CDbl(dato12.text)
    End Sub
    '========================================================
    'KeyUp
    '========================================================
    
    '========================================================
    'LostFocus
    '========================================================
    Private Sub dato8_LostFocus()
        Call limpiaBarra(2)
    End Sub
    
    Private Sub dato9_LostFocus()
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
        Call Centrar(Me)
    End Sub
    
    Private Sub Form_Unload(Cancel As Integer)
        Principal.barraEstado.Panels(1).text = UCase(Principal.Caption)
        Call limpiaBarra(2)
    End Sub

'=============================================================================
'PASA LOS DATOS DE LOS CONTROLES A LA ESTRUCTURA
'=============================================================================
    Private Sub ctrltostruct()
        f.tipo = "FV"
        f.numero = dato1.text
        f.fEmision = dato4.text & "-" & dato3.text & "-" & dato2.text
        f.fVencimiento = dato7.text & "-" & dato6.text & "-" & dato5.text
        f.rut = dato8.text & lblDVC.Caption
        f.sucursal = dato9.text
        f.cajera = dato10.text
        f.monto = dato11.text
        f.abono = dato12.text
        f.obs = dato13.text
        Call grabarFactManCob(f, modifica)
        Call retorno
    End Sub
'=============================================================================
'PASA LOS DATOS DE LOS CONTROLES A LA ESTRUCTURA
'=============================================================================

'=============================================================================
'PASA LOS DATOS DE LA ESTRUCTURA A LOS CONTROLES
'=============================================================================
    Private Sub structtoctrl()
        Dim cad As String
        dato1.text = f.numero
        dato2.text = Format(f.fEmision, "dd")
        dato3.text = Format(f.fEmision, "mm")
        dato4.text = Format(f.fEmision, "yyyy")
        dato5.text = Format(f.fVencimiento, "dd")
        dato6.text = Format(f.fVencimiento, "mm")
        dato7.text = Format(f.fVencimiento, "yyyy")
        dato8.text = f.rut
        dato9.text = f.sucursal
        dato10.text = f.cajera
        dato11.text = f.monto
        dato12.text = f.abono
        dato13.text = f.obs
        Call dato8_KeyPress(13)
        Call dato9_KeyPress(13)
        Call dato10_KeyPress(13)
        Call DeshabilitarCajas(Me)
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
        Call eliminarFactManCob(f.tipo, f.numero)
        Call retorno
        Call HabilitarCajas(Me, modifica)
        dato1.SetFocus
    End Sub

    Private Sub retorno()
        Call LimpiarCajas(FactManCob)
        Call LimpiarLabels(FactManCob)
        modifica = False
        Call DeshabilitarCajas(Me)
        dato1.SetFocus
    End Sub
        
    Private Sub anterior()
        If leerFactManual(f, dato1.text, "<") = True Then
            structtoctrl
        End If
    End Sub
    
    Private Sub siguiente()
        If leerFactManual(f, dato1.text, ">") = True Then
            structtoctrl
        End If
    End Sub
'=============================================================================
'OPCIONES
'=============================================================================

Private Sub opciones_GotFocus()
    manual.SetFocus
End Sub
