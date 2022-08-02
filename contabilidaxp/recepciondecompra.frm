VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "msadodc.ocx"
Object = "{4777436C-EB8C-4596-98A8-EBCF98683969}#1.0#0"; "FlexCell.ocx"
Object = "{ADBBAED6-B16F-40DC-80DF-B44910CBA76C}#1.0#0"; "Frame-Xp.ocx"
Begin VB.Form Rcompra02 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Recepción orden de Compra"
   ClientHeight    =   10050
   ClientLeft      =   390
   ClientTop       =   330
   ClientWidth     =   14925
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   670
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   995
   StartUpPosition =   2  'CenterScreen
   Begin XPFrame.FrameXp FrameQuickMenu 
      Height          =   615
      Left            =   11520
      TabIndex        =   41
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
         TabIndex        =   43
         Top             =   280
         Width           =   1335
      End
      Begin VB.CommandButton botonmisaccesos 
         Caption         =   "Permisos Modulo"
         Height          =   255
         Left            =   120
         TabIndex        =   42
         Top             =   280
         Width           =   1455
      End
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H0080FF80&
      Caption         =   "MODIFICA DETALLE FACTURA"
      Height          =   420
      Left            =   7110
      Style           =   1  'Graphical
      TabIndex        =   39
      Top             =   9090
      Width           =   3210
   End
   Begin XPFrame.FrameXp bodega 
      Height          =   1230
      Left            =   5850
      TabIndex        =   21
      Top             =   3960
      Visible         =   0   'False
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   2170
      BackColor       =   16761024
      Caption         =   "Proveedor"
      CaptionEstilo3D =   1
      BackColor       =   16761024
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.TextBox dato7 
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
         Left            =   90
         MaxLength       =   9
         TabIndex        =   23
         Tag             =   "rut"
         Top             =   585
         Width           =   1455
      End
      Begin VB.CommandButton continuar 
         BackColor       =   &H00808080&
         Caption         =   "Continuar"
         Enabled         =   0   'False
         Height          =   375
         Left            =   3555
         TabIndex        =   8
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label dv2 
         BackColor       =   &H00C0FFFF&
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
         Height          =   285
         Left            =   1575
         TabIndex        =   26
         Top             =   585
         Width           =   285
      End
      Begin VB.Label Label2 
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "RUT CONTABLE"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   90
         TabIndex        =   25
         Top             =   270
         Width           =   1455
      End
      Begin VB.Label nombrecontable 
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
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   90
         TabIndex        =   24
         Top             =   900
         Width           =   5505
      End
   End
   Begin XPFrame.FrameXp FrameXp2 
      Height          =   4650
      Left            =   120
      TabIndex        =   20
      Top             =   1200
      Width           =   14505
      _ExtentX        =   25585
      _ExtentY        =   8202
      BackColor       =   16761024
      Caption         =   "Detalle de Recepcion"
      CaptionEstilo3D =   1
      BackColor       =   16761024
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.CommandButton Command2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "R&etorno"
         Height          =   345
         Left            =   7650
         Style           =   1  'Graphical
         TabIndex        =   32
         Top             =   4185
         Width           =   2280
      End
      Begin VB.CommandButton recepciona 
         BackColor       =   &H00C0C0C0&
         Caption         =   "&Recepcionar"
         Height          =   345
         Left            =   5040
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   4185
         Visible         =   0   'False
         Width           =   2280
      End
      Begin FlexCell.Grid Grid3 
         Height          =   3780
         Left            =   0
         TabIndex        =   6
         Top             =   315
         Width           =   14295
         _ExtentX        =   25215
         _ExtentY        =   6668
         Cols            =   5
         DefaultFontSize =   8.25
         Rows            =   30
      End
      Begin VB.Label lblrut 
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
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   45
         TabIndex        =   40
         Top             =   4230
         Width           =   4695
      End
      Begin VB.Label NETO 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00400000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "$ 0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FFFF&
         Height          =   420
         Left            =   12240
         TabIndex        =   33
         Top             =   4095
         Width           =   1785
      End
      Begin VB.Label total 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00400000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "$ 0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FFFF&
         Height          =   420
         Left            =   10350
         TabIndex        =   28
         Top             =   4095
         Width           =   1830
      End
   End
   Begin VB.TextBox lineas 
      Height          =   285
      Left            =   7320
      MaxLength       =   3
      TabIndex        =   29
      Top             =   8640
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox pivote 
      Height          =   330
      Left            =   6795
      MaxLength       =   10
      TabIndex        =   27
      Top             =   8190
      Visible         =   0   'False
      Width           =   1410
   End
   Begin XPFrame.FrameXp facturas 
      Height          =   2820
      Left            =   90
      TabIndex        =   22
      Top             =   5895
      Width           =   14505
      _ExtentX        =   25585
      _ExtentY        =   4974
      BackColor       =   16761024
      Caption         =   "Ingreso de Documentos"
      CaptionEstilo3D =   1
      BackColor       =   16761024
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.CommandButton Command3 
         Caption         =   "&Grabar Documentos"
         Height          =   255
         Left            =   4275
         TabIndex        =   10
         Top             =   2475
         Width           =   2280
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Eliminar Detalle"
         Height          =   255
         Left            =   6885
         TabIndex        =   11
         Top             =   2475
         Width           =   2505
      End
      Begin FlexCell.Grid Grid1 
         Height          =   2175
         Left            =   0
         TabIndex        =   9
         Top             =   225
         Width           =   14295
         _ExtentX        =   25215
         _ExtentY        =   3836
         Cols            =   5
         DefaultFontSize =   8.25
         Rows            =   30
         DateFormat      =   2
      End
   End
   Begin XPFrame.FrameXp FrameXp1 
      Height          =   1185
      Left            =   135
      TabIndex        =   15
      Top             =   0
      Width           =   14505
      _ExtentX        =   25585
      _ExtentY        =   2090
      BackColor       =   16761024
      Caption         =   "Datos"
      CaptionEstilo3D =   1
      BackColor       =   16761024
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.TextBox BODE 
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
         Left            =   2295
         MaxLength       =   2
         TabIndex        =   35
         Tag             =   "rut"
         Top             =   720
         Width           =   555
      End
      Begin VB.CommandButton BTMODIFICA 
         BackColor       =   &H00FF8080&
         Caption         =   "FINALIZA MODIFICACION"
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
         Left            =   6345
         Style           =   1  'Graphical
         TabIndex        =   34
         Top             =   720
         Width           =   2580
      End
      Begin VB.TextBox dato5 
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
         Left            =   6090
         Locked          =   -1  'True
         MaxLength       =   9
         TabIndex        =   4
         Tag             =   "rut"
         Top             =   360
         Width           =   1095
      End
      Begin VB.TextBox DATO1 
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
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   930
         MaxLength       =   10
         TabIndex        =   0
         Top             =   360
         Width           =   1380
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
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   4050
         MaxLength       =   4
         TabIndex        =   3
         Tag             =   "fecha"
         Top             =   360
         Width           =   615
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
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   3690
         MaxLength       =   2
         TabIndex        =   2
         Tag             =   "fecha"
         Top             =   360
         Width           =   375
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
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   3330
         MaxLength       =   2
         TabIndex        =   1
         Tag             =   "fecha"
         Top             =   360
         Width           =   375
      End
      Begin VB.Label autorizadopor 
         BackColor       =   &H00000000&
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
         ForeColor       =   &H0000FF00&
         Height          =   330
         Left            =   9090
         TabIndex        =   37
         Top             =   720
         Width           =   3390
      End
      Begin VB.Label fechaautorizacion 
         BackColor       =   &H00000000&
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
         ForeColor       =   &H0000FF00&
         Height          =   330
         Left            =   12600
         TabIndex        =   36
         Top             =   720
         Width           =   1410
      End
      Begin VB.Label Label5 
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " BODEGA RECEPCION"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   120
         TabIndex        =   31
         Top             =   720
         Width           =   2145
      End
      Begin VB.Label BODEGARECEPCION 
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
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   2895
         TabIndex        =   30
         Top             =   720
         Width           =   6000
      End
      Begin VB.Label dv 
         BackColor       =   &H00C0FFFF&
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
         Height          =   285
         Left            =   7170
         TabIndex        =   5
         Top             =   360
         Width           =   285
      End
      Begin VB.Label Label4 
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " PROVEEDOR"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   4725
         TabIndex        =   19
         Top             =   360
         Width           =   1305
      End
      Begin VB.Label Label3 
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " FOLIO:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   90
         TabIndex        =   18
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label6 
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " FECHA :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   2490
         TabIndex        =   17
         Top             =   360
         Width           =   735
      End
      Begin VB.Label nombreproveedor 
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
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   7530
         TabIndex        =   16
         Top             =   360
         Width           =   6495
      End
   End
   Begin VB.PictureBox MANUAL 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   135
      Left            =   0
      ScaleHeight     =   135
      ScaleWidth      =   375
      TabIndex        =   13
      Top             =   0
      Width           =   375
   End
   Begin MSAdodcLib.Adodc movi 
      Height          =   375
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   661
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
   Begin MSAdodcLib.Adodc ordenes 
      Height          =   330
      Left            =   0
      Top             =   5400
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
   Begin FlexCell.Grid impresion 
      Height          =   735
      Left            =   9000
      TabIndex        =   14
      Top             =   8055
      Visible         =   0   'False
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   1296
      Cols            =   5
      DefaultFontName =   "Arial"
      DefaultFontSize =   8.25
      DefaultFontBold =   -1  'True
      Rows            =   30
   End
   Begin VB.Label TIPOPAGO 
      Caption         =   "Label1"
      Height          =   285
      Left            =   90
      TabIndex        =   38
      Top             =   1035
      Visible         =   0   'False
      Width           =   600
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash opciones 
      Height          =   1215
      Left            =   135
      TabIndex        =   12
      Top             =   8730
      Visible         =   0   'False
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
Attribute VB_Name = "Rcompra02"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Private FORMATOGRILLA(10, 20) As String
    Private sg As String
    Private retienecarne As Boolean

    Private pasar As String
    Private PRECIOS(10, 4) As String
    Private cantidaddeprecios As Double
    Private PORCENTAJES(10) As Double
    Private NPRECIOS(10, 3) As Double
    Private pdescuento As Double
    Private impuesto As Double
    Private linea2 As Integer
    Private MODIFI As Integer
    Private existe As String
    Private vacio As Boolean
    Private fila As Integer
    Private columna As Integer
    Private direccion As String
    Private comuna As String
    Private ciudad As String
    Private ICA As Double
    Private IHA As Double
    Private descuento As Double
    Private ENCONTRADO As Boolean
    Private OBSERVA As String
    Private pasada As Boolean
    
Sub grabar()
    Dim lin As Double
    Dim total As Double
    Dim preciocosto As Double
    Dim margen As Double
    
        total = 0
    lin = 0
    For k = 1 To Grid3.Rows - 1
        Rem If CDbl(Grid3.Cell(k, 3).text) = 0 Then GoTo PASO:
        lin = lin + 1
        LINEAS.text = lin
        Call ceros(LINEAS)
        campos(0, 0) = "tipo"
        campos(1, 0) = "numero"
        campos(2, 0) = "linea"
        campos(3, 0) = "fecha"
        campos(4, 0) = "rut"
        campos(5, 0) = "codigo"
        campos(6, 0) = "cantidad"
        campos(7, 0) = "uxc"
        campos(8, 0) = "unidades"
        campos(9, 0) = "precio"
        campos(10, 0) = "total"
        campos(11, 0) = "costoventa"
        campos(12, 0) = "bodega"
        campos(13, 0) = "bodegatraspaso"
        campos(14, 0) = "descripcion"
        
        campos(15, 0) = ""
        campos(0, 1) = "OC"
        campos(1, 1) = dato1.text
        campos(2, 1) = LINEAS.text
        campos(3, 1) = Format(fechasistema, "yyyy-mm-dd")
        campos(4, 1) = DATO5.text & DV.Caption
        campos(5, 1) = Replace(Grid3.Cell(k, 1).text, ",", ".")
        campos(6, 1) = Replace(Grid3.Cell(k, 3).text, ",", ".")
        campos(7, 1) = Replace(Grid3.Cell(k, 4).text, ",", ".")
        campos(8, 1) = Replace(Grid3.Cell(k, 5).text, ",", ".")
        campos(9, 1) = Replace(Grid3.Cell(k, 6).text, ",", ".")
        campos(10, 1) = Replace(Grid3.Cell(k, 7).text, ",", ".")
        campos(11, 1) = "0"
        'If MODIFI = 0 Then campos(12, 1) = dato6.text
        If MODIFI = 1 Then campos(12, 1) = BODE.text
        Rem Call MODIFICAORDEN(campos(1, 1), campos(5, 1), campos(6, 1), campos(7, 1), campos(8, 1), campos(9, 1), campos(10, 1), campos(2, 1))
        campos(14, 1) = Grid3.Cell(k, 2).text
        
        campos(0, 2) = "l_movimientos_detalle_" + localorden
        condicion = ""
        op = 2
        sqlconta.response = campos
        Set sqlconta.conexion = gestionrubro
        Call sqlconta.sqlconta(op, condicion)
        total = total + Grid3.Cell(k, 7).text
        'all actualiza_stock("+", Grid3.Cell(k, 1).text, "N", "S", campos(12, 1), dato4.text, Grid3.Cell(k, 5).text, Grid3.Cell(k, 6).text, dato4.text & "-" & dato3.text & "-" & dato2.text, dato5.text & dv.Caption)
        margen = leermargen(Grid3.Cell(k, 1).text)
        If MODIFI = 0 Then Call modificacostos(Grid3.Cell(k, 1).text, Grid3.Cell(k, 6).text)
        If margen <> 0 Then
        If MODIFI = 0 Then
            If leernoactualiza(Grid3.Cell(k, 1).text) = 0 Then
                If campos(8, 1) <> 0 Then
'                        Call grabarPrecios(Grid3.Cell(k, 1).text, Grid3.Cell(k, 6).text, margen)
                End If
            End If
            End If
        
        End If
        
PASO:
    Next k
    
        'GRABA EN MOVIMIENTOS_CABEZA_
        campos(0, 0) = "tipo"
        campos(1, 0) = "numero"
        campos(2, 0) = "rut"
        campos(3, 0) = "fecha"
        campos(4, 0) = "monto"
        campos(5, 0) = "destino"
        campos(6, 0) = "localdestino"
        campos(7, 0) = ""
    
        campos(0, 1) = "OC"
        campos(1, 1) = dato1.text
        campos(2, 1) = DATO5.text & DV.Caption
        campos(3, 1) = dato4.text & dato3.text & dato2.text
        campos(4, 1) = total
        'f MODIFI = 0 Then campos(5, 1) = dato6.text
        If MODIFI = 1 Then campos(5, 1) = BODE.text
        campos(6, 1) = localorden
        
        campos(0, 2) = "l_movimientos_cabeza_" + localorden
        condicion = ""
        op = 2
        sqlconta.response = campos
        Set sqlconta.conexion = gestionrubro
        Call sqlconta.sqlconta(op, condicion)
        
        'ACTUALIZA EN ORDENDECOMPRA_CABEZA_
        campos(0, 0) = "montorecepcionado"
        campos(1, 0) = "fecharecepcion"
        campos(2, 0) = "destino"
        campos(3, 0) = ""
        
        campos(0, 1) = total
        campos(1, 1) = Format(fechasistema, "yyyy-mm-dd")
        'ampos(2, 1) = dato6.text
        
        campos(0, 2) = "l_ordendecompra_cabeza_" + localorden
        condicion = "numero='" & dato1.text & "'"
        op = 3
        sqlconta.response = campos
        Set sqlconta.conexion = gestionrubro
        Call sqlconta.sqlconta(op, condicion)
        
End Sub



Sub leeordendetalle()
    Dim suma As Double
    Dim lin As Integer
    Dim resultados As rdoResultset
    Dim csql As New rdoQuery
    Dim rut As String
    Dim netos As Double
    
    Set csql.ActiveConnection = gestionrubro
    csql.sql = "SELECT codigo,r_maestroproductos_fijo_" & rubro & ".descripcion,cantidad,uxc,unidades,precio,total "
    csql.sql = csql.sql + "FROM l_ordendecompra_detalle_" + localorden + ",r_maestroproductos_fijo_" & rubro & " "
    csql.sql = csql.sql + "WHERE numero='" + dato1.text + "' and r_maestroproductos_fijo_" & rubro & ".codigobarra=l_ordendecompra_detalle_" + localorden + ".codigo order by linea "
    csql.Execute
    suma = 0
    lin = 1
    Grid3.Rows = csql.RowsAffected + 1
    If csql.RowsAffected > 0 Then
        
        Set resultados = csql.OpenResultset
        LBLRUT.Caption = resultados(7)
        While Not resultados.EOF
            Grid3.Cell(lin, 0).text = lin
            For k = 0 To 6
                Grid3.Cell(lin, k + 1).text = resultados(k)
            Next k
            Grid3.Cell(lin, 8).text = Round((CDbl(resultados(6) / 1.19) + 0.5), 0)
            Grid3.Column(3).Locked = False
            suma = suma + resultados(6)
            netos = netos + Round((CDbl(resultados(6) / 1.19) + 0.5), 0)
            
            resultados.MoveNext
            lin = lin + 1
        Wend
        resultados.Close
        Set resultados = Nothing
    End If
    total.Caption = Format(suma, "$ ###,###,###,##0")
    NETO.Caption = Format(netos, "$ ###,###,###,##0")
    
    If csql.RowsAffected <> 0 Then Grid3.Enabled = True: recepciona.Visible = True: Command2.Visible = True
    
End Sub

Sub grabarPago()
    Dim i As Integer
    Dim cmps(5, 3) As String
    campos(0, 0) = "tipo"
    campos(1, 0) = "numero"
    campos(2, 0) = "fecha"
    campos(3, 0) = "vencimiento"
    campos(4, 0) = "neto"
    campos(5, 0) = "iva"
    campos(6, 0) = "exento"
    campos(7, 0) = "impuestos"
    campos(8, 0) = "total"
    campos(9, 0) = "categoria"
    campos(10, 0) = "bonificacion"
    campos(11, 0) = "ordendecompra"
    campos(12, 0) = "linea"
    campos(13, 0) = "rut"
    campos(14, 0) = ""
    cmps(0, 0) = "ordenconfactura"
    cmps(1, 0) = "ordenenlazada"
    cmps(2, 0) = ""
    For i = 1 To Grid1.Rows - 1
        If Grid1.Cell(i, 1).text = "OE ORDEN DE ENLACE" Then
            cmps(0, 1) = Grid1.Cell(i, 2).text
            cmps(1, 1) = dato1.text
            cmps(2, 1) = ""
            cmps(0, 2) = "l_ordendecompra_enlace_factura_" & localorden
            condicion = ""
            op = 2
            sqlconta.response = cmps
            Set sqlconta.conexion = gestionrubro
            Call sqlconta.sqlconta(op, condicion)
        End If
            If Grid1.Cell(i, 10).text <> "" Or Mid(Grid1.Cell(i, 1).text, 1, 2) = "OE" Then
                campos(0, 1) = Grid1.Cell(i, 1).text
                campos(1, 1) = Grid1.Cell(i, 2).text
                campos(2, 1) = Format(Grid1.Cell(i, 3).text, "yyyy-mm-dd")
                campos(3, 1) = Format(Grid1.Cell(i, 4).text, "yyyy-mm-dd")
                campos(4, 1) = Grid1.Cell(i, 5).text
                campos(5, 1) = Grid1.Cell(i, 6).text
                campos(6, 1) = Grid1.Cell(i, 7).text
                campos(7, 1) = Grid1.Cell(i, 8).text
                campos(8, 1) = Replace(Grid1.Cell(i, 9).text, ".", "")
                campos(9, 1) = Grid1.Cell(i, 10).text
                
                campos(12, 1) = Str(i)
                campos(13, 1) = dato7.text + dv2.Caption
                
                
                
                
                campos(10, 1) = Grid1.Cell(i, 11).text
                campos(11, 1) = dato1.text
                
                
                campos(0, 2) = "l_ordendecompra_detalle_facturas_" & localorden
                condicion = ""
                op = 2
                sqlconta.response = campos
                Set sqlconta.conexion = gestionrubro
                Call sqlconta.sqlconta(op, condicion)
            End If
      
    Next i
End Sub

Sub leeCabeza()
    campos(0, 0) = "numero"
    campos(1, 0) = "fecha"
    campos(2, 0) = "proveedor"
    campos(3, 0) = "fechaentrega"
    campos(4, 0) = "formadepago"
    campos(5, 0) = "observaciones"
    campos(6, 0) = "usuarioautorizacion"
    campos(7, 0) = "fechaautorizacion"
    campos(8, 0) = "fecharecepcion"
    campos(9, 0) = "montocomprado"
    campos(10, 0) = "montorecepcionado"
    campos(11, 0) = "prontopago"
    campos(12, 0) = ""
  
    campos(0, 2) = clientesistema & "gestion" & rubro & ".l_ordendecompra_cabeza_" + localorden
    condicion = "numero='" + dato1.text + "'"
    op = 5
    sqlconta.response = campos
    Set sqlconta.conexion = gestionrubro
    Call sqlconta.sqlconta(op, condicion)
    If sqlconta.status = 0 Then
       
       dato1.Enabled = True
       dato2.Enabled = True
       dato3.Enabled = True
       dato4.Enabled = True
       DATO5.Enabled = True
               autorizadopor.Caption = sqlconta.response(6, 3)
        fechaautorizacion.Caption = Format(sqlconta.response(7, 3))

       dato1.text = sqlconta.response(0, 3)
       If existe <> "S" Then
       dato2.text = Mid(sqlconta.response(1, 3), 1, 2)
       dato3.text = Mid(sqlconta.response(1, 3), 4, 2)
       dato4.text = Mid(sqlconta.response(1, 3), 7, 4)
       End If
       DATO5.text = Mid(sqlconta.response(2, 3), 1, 9)
       DV.Caption = Mid(sqlconta.response(2, 3), 10, 1)
       pdescuento = CDbl(sqlconta.response(11, 3))
       If sqlconta.response(5, 3) <> Null Then OBSERVA = sqlconta.response(5, 3)
        Call leeTIPOPAGO(sqlconta.response(4, 3))
        
       Call leeproveedor
       If existe <> "S" Then Call leeordendetalle
    Else
       dato2.text = ""
       dato3.text = ""
       dato4.text = ""
       MsgBox "Documento no existente, ingrese un número de documento válido.", vbInformation + vbOKOnly, "Error"
       pasar = "S"
       dato1.SetFocus
    End If
End Sub


Sub NUEVA()
'CARGAGRILLAbodegas
'CARGAGRILLA
'planillaproveedor
'planillaoc
'estadisticaoc
End Sub

Private Sub BODE_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = vbKeyF2 Then Call ayudaBodega2(BODE)
End Sub

Private Sub BODE_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Call ceros(BODE): Call leerBodega2(BODE.text)
    
End Sub

Private Sub botonmisaccesos_Click()
programafiltro = Me.Caption
misaccesos.Show
End Sub


Private Sub botonmisfavoritos_Click()
Call AgregaFavoritos(USUARIOSISTEMA, App.EXEName, Me.Name, Me.Caption)

End Sub

Private Sub BTMODIFICA_Click()
BTMODIFICA.Visible = False
grabar
grabarPago
MODIFI = 0
End Sub

Private Sub Command1_Click()
    Dim i As Integer
    
    Call eliminaPago
    Grid1.Rows = 2
    For i = 1 To Grid1.Cols - 1
        Grid1.Column(i).Locked = False
    Next i
    Grid1.AddItem ""
    Grid1.Cell(1, 1).SetFocus
End Sub

Private Sub COMMAND2_Click()
retorno
End Sub

Private Sub Command3_Click()
    Dim i As Integer
    
    If Grid1.Cell(1, 10).text <> "" Or Mid(Grid1.Cell(1, 1).text, 1, 2) = "OE" Then
        'COMPROBAR SI LOS NUMEROS DE DOCUMENTO HAN SIDO INGRESADOS ANTERIORMENTE.
'        For i = 1 To Grid1.Rows - 1
'            If Grid1.Cell(i, 10).text <> "" Then
'                If Comprueba_Numero_Facturas(Left(Grid1.Cell(i, 1).text, 2), Grid1.Cell(i, 2).text, dato5.text + dv.caption) = True Then
'                    MsgBox "Existe al menos un número de documento duplicado, revise su información e ingrese nuevamente.", vbExclamation + vbOKOnly, "Documento duplicado"
'                    Grid1.Cell(i, 2).SetFocus
'                    Exit Sub
'                End If
'            End If
'        Next i
'       pasada = True
'        If pasada = True Then
'        Call GRABAR
'        End If
'
        recepciona.Visible = False
        bodega.Visible = True
        Rem dato6.SetFocus
    
        
        Call grabarPago
        Call retorno
     
        Command3.Enabled = False
    End If
    pasada = True

End Sub

Function Comprueba_Numero_Facturas(tipo As String, numero As String, rut As String) As Boolean
    campos(0, 0) = "numero"
    campos(1, 0) = ""
    campos(0, 2) = "l_ordendecompra_detalle_facturas_" & localorden
    condicion = "tipo LIKE '" & tipo & "' AND numero LIKE '" & numero & "' AND rut ='" + rut + "' "
    op = 5
    Set sqlconta.conexion = gestionrubro
    sqlconta.response = campos
    Call sqlconta.sqlconta(op, condicion)

    If sqlconta.status = 0 Then
        Comprueba_Numero_Facturas = True
        Exit Function
    End If
    If sqlconta.status = 4 Then
        Comprueba_Numero_Facturas = False
        Exit Function
    End If
End Function



Private Sub Command4_Click()
bodega.Visible = True
dato7.SetFocus

End Sub

Private Sub continuar_Click()
        pasada = False
       Rem Call Elimina_Impuestos(dato1.text)
        eliminaPago
        
        facturas.Visible = True
        bodega.Visible = False
        Grid1.Enabled = True
        Grid1.Rows = Grid1.Rows + 1
        For k = 1 To Grid1.Cols - 1
        Grid1.Column(k).Locked = False
        Next k
        
        Grid1.Cell(1, 1).SetFocus
        Command3.Visible = True
        
        
        Command3.Enabled = True
        

End Sub

Private Sub dato5_KeyDown(KeyCode As Integer, Shift As Integer)
    'If KeyCode = vbKeyF2 Then Call ayudaProveedor(dato5)
    Call flechas(dato1, dato2, KeyCode)
End Sub

Private Sub Form_Load()
    Me.Left = (Screen.Width - Me.Width) / 2 - 50
    Me.Top = (Screen.Height - Me.Height) / 2 - 900
    'Call Conectargestion(servidor, basedatos, usuario, password)
     Rem  Call Conectargestionrubro(servidor, "eltit_gestion" + RUBRO, usuario, password)
  
    Call planillaoc
    Call planillafacturasdecompra
    Call Conectarconta(Servidor, clientesistema + "conta" & empresaactiva, Usuario, password)
    
   tiposdeprecios
   
    BODE.Enabled = False
    BTMODIFICA.Visible = False
    pasada = True
    Command4.Visible = False
    
   MODIFI = 0
    facturas.Visible = False
End Sub

Private Sub dato1_KeyDown(KeyCode As Integer, Shift As Integer)
    Call flechas(dato1, dato2, KeyCode)
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

'Private Sub dato6_KeyDown(KeyCode As Integer, Shift As Integer)
'    If KeyCode = vbKeyF2 Then Call ayudaBodega(dato6)
'    Call flechas(dato6, dato7, KeyCode)
'End Sub

Private Sub dato7_KeyDown(KeyCode As Integer, Shift As Integer)
    Call flechas(dato7, dato7, KeyCode)
End Sub

Private Sub dato1_GotFocus()
    Call cargatexto(dato1)
    Call dato1_KeyPress(13)
    
End Sub

Private Sub dato2_GotFocus()
    Dim i As Integer
    
    If MODIFI = 0 Then
    Call leerecepcion
    
    
    If existe = "S" Then
        
        Call leeCabeza
        facturas.Visible = True
        For i = 0 To Grid1.Cols - 1
            Grid1.Column(i).Locked = True
        Next i
        Call leerpagos
        recepciona.Visible = False
        Command2.Visible = False
        
        Command1.Visible = False
        Command3.Visible = False
        dato1.Enabled = False
        dato2.Enabled = False
        dato3.Enabled = False
        dato4.Enabled = False
        DATO5.Enabled = False
       
        
    Else
    MsgBox ("ORDEN NO EXISTE ")
    Unload Me
    
  
    End If
End If
    

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

Private Sub dato7_GotFocus()
    
    Call cargatexto(dato7)
End Sub

Private Sub dato1_KeyPress(KeyAscii As Integer)
    snum = 1: KeyAscii = esNumero(KeyAscii)
    If KeyAscii = 13 Then Call ceros(dato1): Call Pregunta(dato1, dato2)
End Sub

Private Sub dato2_KeyPress(KeyAscii As Integer)
    snum = 1: KeyAscii = esNumero(KeyAscii)
    If KeyAscii = 13 Then Call ceros(dato2): Call Pregunta(dato2, dato3)
End Sub

Private Sub dato3_KeyPress(KeyAscii As Integer)
    snum = 1: KeyAscii = esNumero(KeyAscii)
    If KeyAscii = 13 Then Call ceros(dato3): Call Pregunta(dato3, dato4)
End Sub

Private Sub dato4_KeyPress(KeyAscii As Integer)
    snum = 1: KeyAscii = esNumero(KeyAscii)
    If KeyAscii = 13 And MODIFI = 0 Then Call ceros(dato4): Call Pregunta(dato4, DATO5)
    If KeyAscii = 13 And MODIFI = 1 Then Call ceros(dato4): Call Pregunta(dato4, BODE)

End Sub

Private Sub dato5_KeyPress(KeyAscii As Integer)
    snum = 1: KeyAscii = esNumero(KeyAscii)
    If KeyAscii = 13 Then Call ceros(DATO5): DV.Caption = rut(DATO5): recepciona.SetFocus: Call leeproveedor
End Sub

'Private Sub dato6_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 27 Then
'    bodega.Visible = False
'    recepciona.Visible = True
'    Command2.Visible = True
'
'    recepciona.SetFocus
'    End If
'
'    KeyAscii = esNumero(KeyAscii)
'    If KeyAscii = 13 Then Call ceros(dato6): Call leerBodega(dato6.text)
'
'End Sub

Private Sub dato7_KeyPress(KeyAscii As Integer)
    snum = 1: KeyAscii = esNumero(KeyAscii)
    If KeyAscii = 13 Then Call ceros(dato7): dv2.Caption = rut(dato7): Call leeRutContable
End Sub

Sub leeproveedor()
    campos(0, 0) = "rut"
    campos(1, 0) = "nombre"
    campos(2, 0) = ""
    campos(0, 2) = clientesistema & "gestion" & rubro & ".r_maestroproveedores_" + rubro
    condicion = "rut='" & DATO5.text + DV.Caption & "'"
    op = 5
    Set sqlconta.conexion = gestionrubro
    sqlconta.response = campos
    Call sqlconta.sqlconta(op, condicion)

    Rem If sqlconta.status = 4 Then dato5.SetFocus: GoTo no:
    nombreproveedor.Caption = sqlconta.response(1, 3)

no:
End Sub

Sub leeRutContable()
    campos(0, 0) = "rut"
    campos(1, 0) = "nombre"
    campos(2, 0) = ""
    campos(0, 2) = "cuentascorrientes"
    condicion = "rut='" & dato7.text & dv2.Caption & "'"
    op = 5
    Set sqlconta.conexion = contadb
    sqlconta.response = campos
    Call sqlconta.sqlconta(op, condicion)

    'If sqlconta.status = 4 Then nombrecontable.Caption = "": dv2.Caption = "": continuar.Enabled = False: MsgBox "No existe una cuenta contable asociada a este proveedor.  Comunique este error al Departamento correspondiente.", vbCritical + vbOKOnly, "Error": dato6.SetFocus: dato7.SetFocus
    'If sqlconta.status = 0 Then
    nombrecontable.Caption = sqlconta.response(1, 3)
    continuar.Enabled = True
    continuar.SetFocus
    
End Sub


Sub planillaoc()
    Rem DATOS DE LA COLUMNA
    Grid3.DefaultFont.Size = 8
    Grid3.DefaultFont.Bold = True
    
    FORMATOGRILLA(1, 1) = "CODIGO"
    FORMATOGRILLA(1, 2) = "DESCRIPCION"
    FORMATOGRILLA(1, 3) = "CAJAS"
    FORMATOGRILLA(1, 4) = "UxC"
    FORMATOGRILLA(1, 5) = "UNIDADES"
    FORMATOGRILLA(1, 6) = "P.UNI."
    FORMATOGRILLA(1, 7) = "TOTAL C/IVA"
    FORMATOGRILLA(1, 8) = "TOTAL NETO "
    
    
    Rem LARGO DE LOS DATOS
    FORMATOGRILLA(2, 1) = "13"
    FORMATOGRILLA(2, 2) = "37"
    FORMATOGRILLA(2, 3) = "9"
    FORMATOGRILLA(2, 4) = "6"
    FORMATOGRILLA(2, 5) = "9"
    FORMATOGRILLA(2, 6) = "9"
    FORMATOGRILLA(2, 7) = "13"
    FORMATOGRILLA(2, 8) = "13"


    Rem TIPO DE DATOS C=CEROS,S=STRING,N=NUMERICO,D=DATE,
    FORMATOGRILLA(3, 1) = "S"
    FORMATOGRILLA(3, 2) = "S"
    FORMATOGRILLA(3, 3) = "N"
    FORMATOGRILLA(3, 4) = "N"
    FORMATOGRILLA(3, 5) = "N"
    FORMATOGRILLA(3, 6) = "N"
    FORMATOGRILLA(3, 7) = "N"
    FORMATOGRILLA(3, 8) = "N"
    
    Rem FORMATO GRILLA
    FORMATOGRILLA(4, 1) = ""
    FORMATOGRILLA(4, 2) = ""
    FORMATOGRILLA(4, 3) = "#,###,##0.0"
    FORMATOGRILLA(4, 4) = "#,###,##0.0"
    FORMATOGRILLA(4, 5) = "#,###,##0.0"
    FORMATOGRILLA(4, 6) = "#,###,##0.0"
    FORMATOGRILLA(4, 7) = "#,###,##0.0"
    FORMATOGRILLA(4, 8) = "#,###,##0.0"
    
    Rem LOCCKED
    FORMATOGRILLA(5, 1) = "FALSE"
    FORMATOGRILLA(5, 2) = "TRUE"
    FORMATOGRILLA(5, 3) = "FALSE"
    FORMATOGRILLA(5, 4) = "FALSE"
    FORMATOGRILLA(5, 5) = "TRUE"
       
     
'    If Verifica_Permiso(Me.Caption, "autoriza") = True Then
'
'    formatogrilla(5, 6) = "FALSE"
'    Else
    FORMATOGRILLA(5, 6) = "TRUE"
'
'    End If

    FORMATOGRILLA(5, 7) = "TRUE"
    FORMATOGRILLA(5, 8) = "TRUE"

Grid3.Cols = 9
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
    Grid3.BackColorFixedSel = RGB(110, 180, 214)
    Grid3.BackColorBkg = RGB(90, 158, 214)
    Grid3.BackColorScrollBar = RGB(231, 235, 247)
    Grid3.BackColor1 = RGB(231, 235, 247)
    Grid3.BackColor2 = RGB(239, 243, 255)
    Grid3.GridColor = RGB(148, 190, 231)
    For k = 1 To Grid3.Cols - 1
        Grid3.Cell(0, k).text = FORMATOGRILLA(1, k)
        Grid3.Column(k).Width = Val(FORMATOGRILLA(2, k)) * Grid3.DefaultFont.Size
        Grid3.Column(k).MaxLength = Val(FORMATOGRILLA(2, k))
        Grid3.Column(k).FormatString = FORMATOGRILLA(4, k)
        Grid3.Column(k).Locked = FORMATOGRILLA(5, k)
        If FORMATOGRILLA(3, k) = "N" Then Grid3.Column(k).Alignment = cellRightCenter
       
    Next k
Grid3.Column(0).Width = 30
Grid3.Cell(0, 0).text = "lin"
Grid3.Range(0, 0, 0, Grid3.Cols - 1).Alignment = cellCenterCenter
    

Grid3.Enabled = False
End Sub

Sub planillafacturasdecompra()
    Rem DATOS DE LA COLUMNA
    Grid1.DefaultFont.Size = 7
    Grid1.DefaultFont.Bold = True
    
    FORMATOGRILLA(1, 1) = "TIPO"
    FORMATOGRILLA(1, 2) = "NUMERO"
    FORMATOGRILLA(1, 3) = "FECHA"
    FORMATOGRILLA(1, 4) = "VENCIMIENTO"
    FORMATOGRILLA(1, 5) = "NETO"
    FORMATOGRILLA(1, 6) = "IVA"
    FORMATOGRILLA(1, 7) = "EXENTO"
    FORMATOGRILLA(1, 8) = "IMPUESTOS"
    FORMATOGRILLA(1, 9) = "TOTAL"
    FORMATOGRILLA(1, 10) = "TIPO"
    FORMATOGRILLA(1, 11) = "BONIFICACION"
    
    Rem LARGO DE LOS DATOS
    FORMATOGRILLA(2, 1) = "20"
    FORMATOGRILLA(2, 2) = "12"
    FORMATOGRILLA(2, 3) = "12"
    FORMATOGRILLA(2, 4) = "12"
    FORMATOGRILLA(2, 5) = "12"
    FORMATOGRILLA(2, 6) = "12"
    FORMATOGRILLA(2, 7) = "12"
    FORMATOGRILLA(2, 9) = "12"
    FORMATOGRILLA(2, 9) = "12"
    FORMATOGRILLA(2, 10) = "12"
    FORMATOGRILLA(2, 11) = "12"


    Rem TIPO DE DATOS C=CEROS,S=STRING,N=NUMERICO,D=DATE,
    FORMATOGRILLA(3, 1) = "C"
    FORMATOGRILLA(3, 2) = "N"
    FORMATOGRILLA(3, 3) = "D"
    FORMATOGRILLA(3, 4) = "D"
    FORMATOGRILLA(3, 5) = "N"
    FORMATOGRILLA(3, 6) = "N"
    FORMATOGRILLA(3, 7) = "N"
    FORMATOGRILLA(3, 9) = "N"
    FORMATOGRILLA(3, 9) = "N"
    FORMATOGRILLA(3, 10) = "C"
    FORMATOGRILLA(3, 11) = "CH"
    
    Rem FORMATO GRILLA
    FORMATOGRILLA(4, 1) = ""
    FORMATOGRILLA(4, 2) = ""
    FORMATOGRILLA(4, 3) = ""
    FORMATOGRILLA(4, 4) = ""
    FORMATOGRILLA(4, 6) = "#,###,##0.0"
    FORMATOGRILLA(4, 7) = "#,###,##0.0"
    FORMATOGRILLA(4, 9) = "#,###,##0.0"
    FORMATOGRILLA(4, 9) = "#,###,##0.0"
    
    Rem LOCCKED
    FORMATOGRILLA(5, 1) = "FALSE"
    FORMATOGRILLA(5, 2) = "FALSE"
    FORMATOGRILLA(5, 3) = "FALSE"
    FORMATOGRILLA(5, 4) = "FALSE"
    FORMATOGRILLA(5, 5) = "FALSE"
    FORMATOGRILLA(5, 6) = "FALSE"
    FORMATOGRILLA(5, 7) = "FALSE"
    FORMATOGRILLA(5, 8) = "TRUE"
    FORMATOGRILLA(5, 9) = "TRUE"
    FORMATOGRILLA(5, 10) = "FALSE"
    FORMATOGRILLA(5, 11) = "FALSE"

    Grid1.Cols = 12
    Grid1.Rows = 2
    
    Grid1.AllowUserResizing = False
    Grid1.DisplayFocusRect = False
    Grid1.ExtendLastCol = True
    Grid1.BoldFixedCell = False
    Grid1.DrawMode = cellOwnerDraw
    Grid1.Appearance = Flat
    Grid1.ScrollBarStyle = Flat
    Grid1.FixedRowColStyle = Flat
    Grid1.BackColorFixed = RGB(90, 158, 214)
    Grid1.BackColorFixedSel = RGB(110, 180, 214)
    Grid1.BackColorBkg = RGB(90, 158, 214)
    Grid1.BackColorScrollBar = RGB(231, 235, 247)
    Grid1.BackColor1 = RGB(231, 235, 247)
    Grid1.BackColor2 = RGB(239, 243, 255)
    Grid1.GridColor = RGB(148, 190, 231)
    For k = 1 To Grid1.Cols - 1
        Grid1.Cell(0, k).text = FORMATOGRILLA(1, k)
        Grid1.Column(k).Width = Val(FORMATOGRILLA(2, k)) * Grid3.DefaultFont.Size
        Grid1.Column(k).MaxLength = Val(FORMATOGRILLA(2, k))
        Grid1.Column(k).FormatString = FORMATOGRILLA(4, k)
        Grid1.Column(k).Locked = FORMATOGRILLA(5, k)
        If FORMATOGRILLA(3, k) = "N" Then Grid1.Column(k).Alignment = cellRightCenter
        If FORMATOGRILLA(3, k) = "D" Then Grid1.Column(k).CellType = cellCalendar
        If FORMATOGRILLA(3, k) = "C" Then Grid1.Column(k).CellType = cellComboBox
        If FORMATOGRILLA(3, k) = "CH" Then Grid1.Column(k).CellType = cellCheckBox
    Next k
    Grid1.Column(0).Width = 0
    Grid1.Range(0, 0, 0, Grid1.Cols - 1).Alignment = cellCenterCenter
    
    With Grid1.ComboBox(1)
        '.Locked = False
        .AutoComplete = True
        .Font.Name = "Courier New"
        .AddItem "FA FACTURA" '1
        .AddItem "ND NOTA DEBITO" '2
        .AddItem "NC NOTA CREDITO" '3
        .AddItem "FAE FACTURA ELECTRONICA" '1
        .AddItem "NDE NOTA DEBITO ELECTRONICA" '2
        .AddItem "NCE NOTA CREDITO ELECTRONICA" '3
        .AddItem "OE ORDEN DE ENLACE" '4
        .AddItem "GD DESPACHO" '4
        .AddItem "FC FACTURA DE COMPRA" '4
        .AddItem "FE FACTURA EXENTA MANUAL" '4
        .AddItem "FEE FACTURA EXENTA ELECTRONICA" '4
    
    
    End With
    With Grid1.ComboBox(10)
        '.Locked = True
        .AutoComplete = True
        .Font.Name = "Courier New"
        .AddItem "MERCADERIAS"
        .AddItem "CIGARRILLOS"
        .AddItem "FRUTAS Y VERDURAS"
        .AddItem "CARNICERIA"
        .AddItem "FIAMBRERIA"
        .AddItem "PANADERIA"
        .AddItem "EMPAQUE"
        .AddItem "DIARIOS"
        
    End With

Grid1.Enabled = False
End Sub

Sub ayudaProveedor(ByRef caja As TextBox)
    Dim campos As Variant
    Dim cfijo As Variant
    Dim largo As Variant
    campos = Array("rut", "nombre")
    largo = Array("13n", "30s")
    cfijo = "no"
    mensajeAyuda = "Ayuda Proveedores"
    cabezas = Array("rut", "nombre")

    Call cargaAyudaT(Servidor, basedatos & rubro, Usuario, password, "r_maestroproveedores_" & rubro, DATO5, campos, cfijo, largo, 2)
    caja.Enabled = True
    caja.SetFocus
    
End Sub

Sub flechas(ByRef caja As TextBox, ByRef caja1 As TextBox, ByRef codigo As Integer)
    If codigo = 38 And caja.Enabled = True Then caja.SetFocus
    If codigo = 40 And caja1.Enabled = True Then caja1.SetFocus
End Sub

Sub Pregunta(ByRef caja As TextBox, ByRef caja1 As TextBox)
    If caja.text = "" And sc = 0 Then caja.SetFocus
    If caja.text <> "" Or sc = 1 Then caja1.Enabled = True: caja1.SetFocus
    sc = 0
End Sub

Sub leelocal(codigo, LINEA)
    campos(0, 0) = "codigo"
    campos(1, 0) = "nombre"
    campos(2, 0) = ""
    campos(0, 2) = "g_maestroempresas"
    condicion = "codigo = '" & codigo & "'"
    op = 5
    Set sqlconta.conexion = gestion
    sqlconta.response = campos
    Call sqlconta.sqlconta(op, condicion)
    'Grid2.Cell(linea, 1).text = sqlconta.response(1, 3)
no:
End Sub

Sub calcularecepcion()
    Dim totales  As Double
    Dim netos As Double
    Dim unidades As Double
    Dim CAJAS As Double
    Dim Precio As Double
    Dim total2 As Double
    
    totales = 0
    netos = 0
    
    For k = 1 To Grid3.Rows - 1
    
    If Grid3.Cell(k, 3).text = "" Then Grid3.Cell(k, 3).text = "0"
    If Grid3.Cell(k, 4).text = "" Then Grid3.Cell(k, 4).text = "0"
    If Grid3.Cell(k, 6).text = "" Then Grid3.Cell(k, 6).text = "0"
    unidades = CDbl(Grid3.Cell(k, 3).text)
    CAJAS = CDbl(Grid3.Cell(k, 4).text)
    unidades = unidades * CAJAS
    Precio = CDbl(Grid3.Cell(k, 6).text)
    
    total2 = Round((unidades * Precio) + 0.5, 0)
    
    
    
    Grid3.Cell(k, 5).text = unidades
    Grid3.Cell(k, 7).text = total2
    Grid3.Cell(k, 8).text = Round(total2 / 1.19 + 0.5, 0)
    
    
    totales = totales + CDbl(Grid3.Cell(k, 7).text)
    netos = netos + CDbl(Grid3.Cell(k, 8).text)

    Next k
    total.Caption = Format(totales, "$ ###,###,##0")
    NETO.Caption = Format(netos, "$ ###,###,##0")
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
 Call Conectarconta(Servidor, clientesistema + "conta", Usuario, password)
End Sub

Private Sub Grid3_KeyDown(KeyCode As Integer, ByVal Shift As Integer)
    fila = Grid3.ActiveCell.row
    columna = Grid3.ActiveCell.col
    Select Case KeyCode
        Case 13, 37, 38, 39, 40
            If Grid3.ActiveCell.text <> "" Then
                vacio = False
            Else
                vacio = True
            End If
    End Select
 If Grid3.ActiveCell.col = 1 And KeyCode = vbKeyF2 Then Call ayudaProducto2(Grid3.ActiveCell.row, Grid3.ActiveCell.col)

End Sub
Private Sub ayudaProducto2(ByVal fila As Long, ByVal columna As Long)
        Dim campos As Variant
        Dim cfijo As Variant
        Dim largo As Variant
        pivote.MaxLength = 13
        
        campos = Array("codigo", "descripcion")
        largo = Array("15n", "30s")
        cfijo = "r_maestroproductos_stock_" & rubro & ".codigo=r_maestroproductos_fijo_" & rubro & ".codigobarra AND r_maestroproductos_stock_" & rubro & ".local='" & localorden & "' "
        mensajeAyuda = "Ayuda"
        cabezas = Array("Codigo", "Nombre")
        Call cargaAyudaT(Servidor, basedatos & rubro, Usuario, password, "r_maestroproductos_stock_" & rubro & ", r_maestroproductos_fijo_" & rubro, pivote, campos, cfijo, largo, 2)
        
        Grid3.Cell(fila, columna).text = pivote.text
    End Sub
'Private Sub Grid3_KeyPress(KeyAscii As Integer)
'    KeyAscii = esNumeroDecimal(Grid3.ActiveCell.text, KeyAscii)
'End Sub

Private Sub Grid3_LeaveCell(ByVal row As Long, ByVal col As Long, NewRow As Long, NewCol As Long, Cancel As Boolean)
    If vacio = True Then
        NewRow = fila
        NewCol = columna
    Else
              If col = 1 And (NewCol > 1 Or NewRow <> row) Then
               pivote.MaxLength = 13
               pivote.text = Grid3.Cell(row, 1).text
               Call ceros(pivote)
               Grid3.Cell(row, 1).text = pivote.text
               Call leerproducto(row, Grid3.Cell(row, 1).text)
               
               If row = Grid3.Rows - 1 And Grid3.Rows > 2 Then GoTo paso2:
               
               If col = 1 And ENCONTRADO = False Then NewCol = 1: NewRow = row
               If col = 1 And ENCONTRADO = True Then NewCol = 4
paso2:
               End If
        
        
        Grid3.Cell(row, 5).text = Grid3.Cell(row, 3).text * Grid3.Cell(row, 4).text
        Grid3.Cell(row, 7).text = Grid3.Cell(row, 5).text * Grid3.Cell(row, 6).text
    End If
    calcularecepcion
    
End Sub
Private Sub leerproducto(ByVal fila As Long, ByVal codigo As String)
        campos(0, 0) = "descripcion"
        campos(1, 0) = "pcosto"
        campos(2, 0) = "cantidadporembalaje"
        campos(3, 0) = ""
        
        campos(0, 2) = "r_maestroproductos_fijo_" & rubro
        condicion = "codigobarra = '" & codigo & "'"
        op = 5
        sqlconta.response = campos
        Set sqlconta.conexion = gestionrubro
        Call sqlconta.sqlconta(op, condicion)
        If sqlconta.status = 0 Then
            Call cargarProducto(fila)
            ENCONTRADO = True
        Else
            Grid3.Cell(fila, 1).text = ""
            ENCONTRADO = False
        End If
    End Sub
    

    Private Sub cargarProducto(ByVal fila As Long)
        Dim i As Integer
        Dim valor As Double
        Grid3.Cell(fila, 2).text = sqlconta.response(0, 3)
        Grid3.Cell(fila, 6).text = sqlconta.response(1, 3)
        Grid3.Cell(fila, 4).text = sqlconta.response(2, 3)
    
    End Sub

Private Sub Grid3_LostFocus()
Grid3.Cell(Grid3.ActiveCell.row, 5).text = Grid3.Cell(Grid3.ActiveCell.row, 3).text * Grid3.Cell(Grid3.ActiveCell.row, 4).text
        Grid3.Cell(Grid3.ActiveCell.row, 7).text = Grid3.Cell(Grid3.ActiveCell.row, 5).text * Grid3.Cell(Grid3.ActiveCell.row, 6).text


End Sub

Private Sub opciones_FSCommand(ByVal command As String, ByVal args As String)
    If command = "retorno" Then Call retorno
    
    
    If command = "imprime" Then Call imprime2
End Sub
Sub MODIFICAFECHA()
MODIFI = 1
dato2.Enabled = True
dato3.Enabled = True
dato4.Enabled = True
BODE.Enabled = True

BODE.Visible = True

BTMODIFICA.Visible = True
ELIMINA



End Sub

'Private Sub imprime()
'
'    Dim o As Integer
'    Dim compradas As Double
'    Dim total As Double
'    Dim totalfacturas As Double
'    Dim codigo As String
'    Dim Row As Integer
'    Dim FINROW As Integer
'    Dim i As Integer
'    Dim j As Integer
'    Dim objReportTitle As FlexCell.ReportTitle
'    Dim margen As Double
'    Dim pventa As Double
'    Dim pcosto As Double
'    Dim margenoriginal As Double
'    Dim MO2 As Double
'    Dim avance As Double
'    Dim codigoimpuesto As String
'
'    impresion.AutoRedraw = False
'    impresion.Cols = 13
'    impresion.Rows = 1
'    impresion.Range(0, 1, 0, 12).FontSize = 8
'    impresion.Column(0).Width = 0
'    impresion.Column(1).Width = 100
'    impresion.Column(2).Width = 230
'    impresion.Column(3).Width = 90
'    impresion.Column(4).Width = 35
'    impresion.Column(5).Width = 75
'    impresion.Column(6).Width = 75
'    impresion.Column(7).Width = 75
'    impresion.Column(8).Width = 75
'    impresion.Column(9).Width = 75
'    impresion.Column(10).Width = 60
'    impresion.Column(11).Width = 75
'    impresion.Column(12).Width = 50
'
'    impresion.Cell(0, 1).text = "CODIGO"
'    impresion.Cell(0, 2).text = "DESCRIPCION"
'    impresion.Cell(0, 3).text = "COMPRA"
'    impresion.Cell(0, 4).text = "UxC"
'    impresion.Cell(0, 5).text = "UNIDADES"
'    impresion.Cell(0, 6).text = "P.COSTO"
'
'
'    impresion.Cell(0, 7).text = "TOTAL"
'    impresion.Cell(0, 9).text = "P.VENTA"
'    impresion.Cell(0, 9).text = "%TEORICO"
'    impresion.Cell(0, 10).text = " %REAL  "
'    impresion.Cell(0, 11).text = "ANTERIOR"
'    impresion.Cell(0, 12).text = "% VAR. "
'
'    impresion.Cell(0, 1).Alignment = cellCenterGeneral
'    impresion.Cell(0, 2).Alignment = cellCenterGeneral
'    impresion.Cell(0, 3).Alignment = cellCenterGeneral
'    impresion.Cell(0, 4).Alignment = cellCenterGeneral
'    impresion.Cell(0, 5).Alignment = cellCenterGeneral
'    impresion.Cell(0, 6).Alignment = cellCenterGeneral
'    impresion.Cell(0, 7).Alignment = cellCenterGeneral
'    impresion.Cell(0, 9).Alignment = cellCenterGeneral
'    impresion.Cell(0, 9).Alignment = cellCenterGeneral
'    impresion.Cell(0, 10).Alignment = cellRightCenter
'    impresion.Cell(0, 11).Alignment = cellRightCenter
'    impresion.Cell(0, 12).Alignment = cellRightCenter
'
'    'Logo
'    'Grid4.Images.Add App.Path & "\Logo.gif", "Logo"
'    'Set objReportTitle = New FlexCell.ReportTitle
'    'objReportTitle.ImageKey = "Logo"
'    'Grid3.ReportTitles.Add objReportTitle
'    'impresion.PageSetup.PrintGridlines = True
'    impresion.PageSetup.BlackAndWhite = True
'    impresion.PageSetup.BottomMargin = 1
'    impresion.PageSetup.LeftMargin = 0.5
'    impresion.PageSetup.RightMargin = 0.5
'    impresion.PageSetup.TopMargin = 1
'    impresion.PageSetup.PrintFixedRow = True
'
'    impresion.DefaultFont.Size = 8
'    impresion.DefaultFont.Bold = True
'    impresion.PageSetup.PrintGridlines = False
'
'
'    ICA = 0
'    IHA = 0
'
'
'    Call cabeza
'    compradas = 0
'    total = 0
'    For i = 1 To Grid3.Rows - 1
'        impresion.AddItem ""
'        For j = 1 To Grid3.Cols - 1
'            impresion.Cell(impresion.Rows - 1, j).text = Grid3.Cell(i, j).text
'        Next j
'        impresion.Cell(impresion.Rows - 1, 6).text = Format(Grid3.Cell(i, 6).text, "#,###,##0.0")
'        codigo = Grid3.Cell(i, 1).text
'        Call leerprecioventa(codigo, i)
'        impresion.Cell(impresion.Rows - 1, 9).text = Format(NPRECIOS(1, 1), "$ ###,###,##0")
'        impresion.Cell(impresion.Rows - 1, 11).text = Format(NPRECIOS(1, 2), "$ ###,###,##0")
'
'        pcosto = CDbl(impresion.Cell(impresion.Rows - 1, 6).text)
'
'        If pcosto = 0 Then pcosto = 1
'        codigoimpuesto = leerimpuesto(codigo)
'        If codigoimpuesto = "00004" Then
'        IHA = IHA + Round((CDbl(Grid3.Cell(i, 7).text) * impuesto / 100) + 0.5)
'        End If
'        If codigoimpuesto = "00005" Then
'        ICA = ICA + Round((CDbl(Grid3.Cell(i, 7).text) * impuesto / 100) + 0.5)
'        End If
'
'
'        margenoriginal = leermargen(codigo)
'        impresion.Cell(impresion.Rows - 1, 9).text = Format(margenoriginal, "% ##0.00")
'
'        pventa = NPRECIOS(1, 1)
'        margen = (((pventa / pcosto) - 1) * 100)
'        impresion.Cell(impresion.Rows - 1, 10).text = Format(margen, "% ##0.00")
'
'        pventa = NPRECIOS(1, 2)
'
'        margen = (((NPRECIOS(1, 1) / NPRECIOS(1, 2)) - 1) * 100)
'        impresion.Cell(impresion.Rows - 1, 12).text = Format(margen, "% ##0.00")
'        impresion.Cell(impresion.Rows - 1, 3).text = Format(impresion.Cell(impresion.Rows - 1, 3).text, "#,###,###.0")
'        impresion.Cell(impresion.Rows - 1, 5).text = Format(impresion.Cell(impresion.Rows - 1, 5).text, "#,###,###.0")
'        impresion.Cell(impresion.Rows - 1, 6).text = Format(impresion.Cell(impresion.Rows - 1, 6).text, "#,###,###.0")
'        impresion.Cell(impresion.Rows - 1, 7).text = Format(impresion.Cell(impresion.Rows - 1, 7).text, "##,###,###")
'
'
'        If Option1.Value = True Then
'            linea2 = impresion.Rows - 1
'            For o = 2 To cantidaddeprecios
'            linea2 = linea2 + 1
'            impresion.AddItem ""
'            impresion.Cell(linea2, 7).text = "PRECIO X " + Str(PRECIOS(o, 4))
'            impresion.Cell(linea2, 9).text = Format(NPRECIOS(o, 1), "$ ###,###,##0")
'            impresion.Cell(linea2, 11).text = Format(NPRECIOS(o, 2), "$ ###,###,##0")
'            MO2 = margenoriginal * PORCENTAJES(o) / 100
'            impresion.Cell(linea2, 9).text = Format(MO2, "% ##0.00")
'            pventa = NPRECIOS(o, 1)
'            margen = (((pventa / pcosto) - 1) * 100)
'            impresion.Cell(linea2, 10).text = Format(margen, "% ##0.00")
'            pventa = NPRECIOS(o, 2)
'            margen = (((NPRECIOS(o, 1) / NPRECIOS(o, 2)) - 1) * 100)
'            impresion.Cell(linea2, 12).text = Format(margen, "% ##0.00")
'
'        Next o
'        Call leerOFERTAS(codigo, pcosto, margenoriginal)
'
'        End If
'        impresion.Range(1, 3, impresion.Rows - 1, impresion.Cols - 1).Alignment = cellRightGeneral
'        impresion.Range(1, 1, impresion.Rows - 1, impresion.Cols - 1).FontSize = 8
'        compradas = compradas + Grid3.Cell(i, 3).text
'        total = total + Grid3.Cell(i, 7).text
'    Next i
'
'
'    impresion.Range(0, 1, 0, impresion.Cols - 1).Borders(cellEdgeTop) = cellThin
'    impresion.Range(0, 1, 0, impresion.Cols - 1).Borders(cellEdgeBottom) = cellThin
'    impresion.Range(0, 1, 0, impresion.Cols - 1).Borders(cellEdgeLeft) = cellThin
'    impresion.Range(0, 1, 0, impresion.Cols - 1).Borders(cellEdgeRight) = cellThin
'    impresion.Range(0, 1, 0, impresion.Cols - 1).Borders(cellInsideVertical) = cellThin
'
'
'
'    FINROW = impresion.Rows - 1
'    impresion.Rows = impresion.Rows + 7
'     If ICA <> 0 Then
'        FINROW = FINROW + 1
'        impresion.Range(FINROW, 5, FINROW, 6).Merge
'        impresion.Cell(FINROW, 5).text = "IMPUESTO CARNE "
'        impresion.Range(FINROW, 5, FINROW, 6).Alignment = cellLeftCenter
'        impresion.Cell(FINROW, 7).text = Format(ICA, "$ ###,###,##0")
'        impresion.Cell(FINROW, 7).Alignment = cellRightGeneral
'    End If
'    If IHA <> 0 Then
'        FINROW = FINROW + 1
'        impresion.Range(FINROW, 5, FINROW, 6).Merge
'        impresion.Cell(FINROW, 5).text = "IMPUESTO HARINA"
'        impresion.Range(FINROW, 5, FINROW, 6).Alignment = cellLeftCenter
'        impresion.Cell(FINROW, 7).text = Format(IHA, "$ ###,###,##0")
'        impresion.Cell(FINROW, 7).Alignment = cellRightGeneral
'    End If
'
'    FINROW = FINROW + 1
'    impresion.Range(FINROW, 1, FINROW, 1).Borders(cellEdgeLeft) = cellThin
'    impresion.Range(FINROW, 9, FINROW, impresion.Cols - 1).Borders(cellEdgeRight) = cellThin
'    impresion.Range(FINROW, 1, FINROW, impresion.Cols - 1).Borders(cellEdgeTop) = cellThin
'    impresion.Cell(FINROW, 2).text = "UNIDADES COMPRADAS"
'    impresion.Cell(FINROW, 3).text = compradas
'    impresion.Cell(FINROW, 3).Alignment = cellRightGeneral
'    total = total + ICA + IHA
'        impresion.Range(FINROW, 5, FINROW, 6).Merge
'        impresion.Cell(FINROW, 5).text = "TOTAL RECEPCION"
'        impresion.Range(FINROW, 5, FINROW, 6).Alignment = cellLeftCenter
'        impresion.Cell(FINROW, 7).text = Format(total, "$ ###,###,##0")
'        impresion.Cell(FINROW, 7).Alignment = cellRightGeneral
'    descuento = Round(((total * pdescuento / 100) + 0.5), 0)
'
'    FINROW = FINROW + 1
'
'    impresion.Range(FINROW, 1, FINROW, 1).Borders(cellEdgeLeft) = cellThin
'    impresion.Range(FINROW, 9, FINROW, impresion.Cols - 1).Borders(cellEdgeRight) = cellThin
'    impresion.Range(FINROW, 1, FINROW, impresion.Cols - 1).Borders(cellEdgeTop) = cellThin
'    impresion.Range(FINROW, 5, FINROW, 6).Merge
'    impresion.Cell(FINROW, 5).text = "(-)D.FIN " + Format(pdescuento, "%##.#0")
'    impresion.Range(FINROW, 5, FINROW + 1, 6).Alignment = cellLeftCenter
'    impresion.Cell(FINROW, 7).text = Format(descuento, "$ ###,###,##0")
'    impresion.Cell(FINROW, 7).Alignment = cellRightGeneral
'
'    total = total - descuento
'
'    FINROW = FINROW + 1
'
'    impresion.Range(FINROW, 1, FINROW, 1).Borders(cellEdgeLeft) = cellThin
'    impresion.Range(FINROW, 9, FINROW, impresion.Cols - 1).Borders(cellEdgeRight) = cellThin
'    impresion.Range(FINROW, 1, FINROW, impresion.Cols - 1).Borders(cellEdgeTop) = cellThin
'    impresion.Range(FINROW, 5, FINROW, 6).Merge
'    impresion.Cell(FINROW, 5).text = "TOTAL (-) DESCUENTO    "
'    impresion.Range(FINROW, 5, FINROW + 1, 6).Alignment = cellLeftCenter
'
'    impresion.Cell(FINROW, 7).text = Format(total, "$ ###,###,##0")
'    impresion.Cell(FINROW, 7).Alignment = cellRightGeneral
'
'
'    FINROW = FINROW + 1
'
'    impresion.Range(FINROW, 1, FINROW, 1).Borders(cellEdgeLeft) = cellThin
'    impresion.Range(FINROW, 9, FINROW, impresion.Cols - 1).Borders(cellEdgeRight) = cellThin
'    impresion.Range(FINROW, 1, FINROW, impresion.Cols - 1).Borders(cellEdgeTop) = cellThin
'    impresion.Range(FINROW, 5, FINROW, 6).Merge
'    impresion.Cell(FINROW, 5).text = "TOTAL FACTURAS"
'    impresion.Range(FINROW, 5, FINROW, 6).Alignment = cellLeftCenter
'
'    totalfacturas = 0
'    For i = 1 To Grid1.Rows - 1
'        totalfacturas = totalfacturas + Grid1.Cell(i, 9).text
'    Next i
'    impresion.Cell(FINROW, 7).text = Format(totalfacturas, "$ ###,###,##0")
'    impresion.Cell(FINROW, 7).Alignment = cellRightCenter
'
'    FINROW = FINROW + 1
'    impresion.Range(FINROW, 1, FINROW, 1).Borders(cellEdgeLeft) = cellThin
'    impresion.Range(FINROW, 9, FINROW, impresion.Cols - 1).Borders(cellEdgeRight) = cellThin
'    impresion.Range(FINROW, 1, FINROW, impresion.Cols - 1).Borders(cellEdgeTop) = cellThin
'    impresion.Range(FINROW, 5, FINROW, 6).Merge
'    impresion.Cell(FINROW, 5).text = "DIFERENCIA"
'    impresion.Range(FINROW, 5, FINROW, 6).Alignment = cellLeftCenter
'    impresion.Cell(FINROW, 7).text = Format(total - totalfacturas, "$ ###,###,##0")
'    impresion.Cell(FINROW, 7).Alignment = cellRightCenter
'      FINROW = FINROW + 1
'
'    impresion.Range(FINROW, 1, FINROW, impresion.Cols - 1).Borders(cellEdgeTop) = cellThin
'    impresion.Range(FINROW, 1, FINROW, impresion.Cols - 1).FontSize = 8
'    impresion.AddItem "DETALLE DE FACTURAS"
'    impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, 9).Merge
'    impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, 9).Alignment = cellCenterGeneral
'    impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, 9).FontBold = True
'    impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, 9).FontSize = 8
'
'    impresion.AddItem "NUMERO" & vbTab & "PROVEEDOR" & vbTab & "FECHA" & vbTab & vbTab & "NETO" & vbTab & "IVA" & vbTab & "EXENTO" & vbTab & "IMPTOS" & vbTab & "TOTAL", False
'    impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, 9).FontSize = 8
'    impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, 9).FontBold = True
'    impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, 9).Alignment = cellCenterGeneral
'    impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, 9).Borders(cellEdgeTop) = cellThin
'    impresion.Cell(impresion.Rows - 1, 1).Border(cellEdgeLeft) = cellThin
'    impresion.Cell(impresion.Rows - 1, 2).Border(cellEdgeLeft) = cellThin
'    impresion.Cell(impresion.Rows - 1, 3).Border(cellEdgeLeft) = cellThin
'    impresion.Cell(impresion.Rows - 1, 4).Border(cellEdgeLeft) = cellThin
'    impresion.Cell(impresion.Rows - 1, 5).Border(cellEdgeLeft) = cellThin
'    impresion.Cell(impresion.Rows - 1, 6).Border(cellEdgeLeft) = cellThin
'    impresion.Cell(impresion.Rows - 1, 7).Border(cellEdgeLeft) = cellThin
'    impresion.Cell(impresion.Rows - 1, 9).Border(cellEdgeLeft) = cellThin
'    impresion.Cell(impresion.Rows - 1, 9).Border(cellEdgeLeft) = cellThin
'    impresion.Cell(impresion.Rows - 1, 9).Border(cellEdgeRight) = cellThin
'    impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, 9).Borders(cellEdgeBottom) = cellThin
'
'    For i = 1 To Grid1.Rows - 1
'        impresion.AddItem Grid1.Cell(i, 2).text & vbTab & nombreproveedor.caption & vbTab & Format(Grid1.Cell(i, 3).text, "dd-mm-yyyy") & vbTab & vbTab & Format(Grid1.Cell(i, 5).text, "###,###,###") & vbTab & Format(Grid1.Cell(i, 6).text, "###,###,###") & vbTab & Format(Grid1.Cell(i, 7).text, "###,###,###") & vbTab & Format(Grid1.Cell(i, 9).text, "###,###,###") & vbTab & Format(Grid1.Cell(i, 9).text, "###,###,###"), False
'        impresion.Cell(impresion.Rows - 1, 1).Alignment = cellCenterGeneral
'        impresion.Cell(impresion.Rows - 1, 2).Alignment = cellLeftCenter
'        impresion.Cell(impresion.Rows - 1, 3).Alignment = cellCenterGeneral
'        impresion.Cell(impresion.Rows - 1, 4).Alignment = cellCenterGeneral
'        impresion.Range(impresion.Rows - 1, 5, impresion.Rows - 1, 9).Alignment = cellRightCenter
'        impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, 9).FontSize = 8
'    Next i
'    impresion.AutoRedraw = True
'
'
'    impresion.PrintPreviewVirtualGrid
'
'
'End Sub
Private Sub imprime2()
    
    Dim o As Integer
    Dim compradas As Double
    Dim total As Double
    Dim totalfacturas As Double
    Dim codigo As String
    Dim row As Integer
    Dim FINROW As Integer
    Dim i As Integer
    Dim j As Integer
    Dim objReportTitle As FlexCell.ReportTitle
    Dim margen As Double
    Dim PVENTA As Double
    Dim pcosto As Double
    Dim margenoriginal As Double
    Dim MO2 As Double
    Dim avance As Double
    Dim codigoimpuesto As String
    Dim NETO As Double
    Dim totales(10) As Double
    Dim BONIFI As String
    
    impresion.AutoRedraw = False
    impresion.Cols = 10
    impresion.Rows = 1
    impresion.Range(0, 1, 0, 9).FontSize = 7
    impresion.Column(0).Width = 30
    impresion.Column(1).Width = 80
    impresion.Column(2).Width = 200
    impresion.Column(3).Width = 60
    impresion.Column(4).Width = 60
    impresion.Column(5).Width = 60
    impresion.Column(6).Width = 60
    impresion.Column(7).Width = 80
    impresion.Column(8).Width = 60
    impresion.Column(9).Width = 60
    
    impresion.Cell(0, 1).text = "CODIGO"
    impresion.Cell(0, 2).text = "DESCRIPCION"
    impresion.Cell(0, 3).text = "CANTIDAD"
    impresion.Cell(0, 4).text = "UXC"
    impresion.Cell(0, 5).text = "P.COMPRA"
    impresion.Cell(0, 6).text = " TOTAL"
    impresion.Cell(0, 7).text = "P.VENTA SISTEMA"
    impresion.Cell(0, 8).text = "%UTI"
    impresion.Cell(0, 9).text = "NETO"
    
    impresion.Cell(0, 1).Alignment = cellLeftCenter
    impresion.Cell(0, 2).Alignment = cellLeftCenter
    impresion.Cell(0, 3).Alignment = cellRightCenter
    impresion.Cell(0, 4).Alignment = cellRightCenter
    impresion.Cell(0, 5).Alignment = cellRightCenter
    impresion.Cell(0, 6).Alignment = cellRightCenter
    impresion.Cell(0, 7).Alignment = cellRightCenter
    impresion.Cell(0, 8).Alignment = cellRightCenter
    impresion.Cell(0, 9).Alignment = cellRightCenter
    For k = 3 To 9
    impresion.Column(k).Alignment = cellRightCenter
    Next k
    
    
    'Logo
    'Grid4.Images.Add App.Path & "\Logo.gif", "Logo"
    'Set objReportTitle = New FlexCell.ReportTitle
    'objReportTitle.ImageKey = "Logo"
    'Grid3.ReportTitles.Add objReportTitle
    'impresion.PageSetup.PrintGridlines = True
    impresion.PageSetup.BlackAndWhite = True
    impresion.PageSetup.BottomMargin = 1
    impresion.PageSetup.LeftMargin = 0.5
    impresion.PageSetup.RightMargin = 0.5
    impresion.PageSetup.TopMargin = 1
    impresion.PageSetup.PrintFixedRow = True
    impresion.PageSetup.PrintFixedColumn = True
    
    impresion.DefaultFont.Size = 8
    impresion.DefaultFont.Bold = False
    impresion.PageSetup.PrintGridlines = False
    
    ICA = 0
    IHA = 0
    Call cabeza
    compradas = 0
    total = 0
    For i = 1 To Grid3.Rows - 1
        impresion.AddItem ""
        impresion.Cell(impresion.Rows - 1, 0).text = i
        impresion.Cell(impresion.Rows - 1, 1).text = Grid3.Cell(i, 1).text
        impresion.Cell(impresion.Rows - 1, 2).text = Grid3.Cell(i, 2).text
        
        If CDbl(Grid3.Cell(i, 3).text) = 0 Then
        impresion.Range(impresion.Rows - 1, 2, impresion.Rows - 1, 4).Merge
        impresion.Cell(impresion.Rows - 1, 2).text = "*** NO LLEGO **"
        GoTo PASO:
        End If
        
        impresion.Cell(impresion.Rows - 1, 3).text = Format(Grid3.Cell(i, 3).text, "###,###,##0.0")
        impresion.Cell(impresion.Rows - 1, 4).text = Format(Grid3.Cell(i, 4).text, "###,###,##0.0")
        impresion.Cell(impresion.Rows - 1, 5).text = Format(Grid3.Cell(i, 6).text, "###,###,##0.0")
        impresion.Cell(impresion.Rows - 1, 6).text = Format(Grid3.Cell(i, 7).text, "###,###,##0")
        
        codigo = Grid3.Cell(i, 1).text
        NETO = Round(CDbl(Grid3.Cell(i, 8).text))
        Call leerprecioventa(codigo, i)
        impresion.Cell(impresion.Rows - 1, 7).text = Format(NPRECIOS(1, 1), "$ ###,###,##0")
        
        pcosto = CDbl(impresion.Cell(impresion.Rows - 1, 5).text)
        
        If pcosto = 0 Then pcosto = 1
        codigoimpuesto = leerimpuesto(codigo)
        If codigoimpuesto = "00006" Then
        IHA = IHA + Round((CDbl((Grid3.Cell(i, 7).text) / 1.19) * impuesto / 100) + 0.5)
        End If
        If codigoimpuesto = "00007" And retienecarne = False Then
        ICA = ICA + Round((CDbl((Grid3.Cell(i, 7).text) / 1.19) * impuesto / 100) + 0.5)
        End If
        
        
        margenoriginal = leermargen(codigo)
        'impresion.Cell(impresion.Rows - 1, 6).text = Format(margenoriginal, "  ##0.00")
        
        PVENTA = NPRECIOS(1, 1)
        margen = (((PVENTA / pcosto) - 1) * 100)
        impresion.Cell(impresion.Rows - 1, 8).text = Format(margen, "  ##0.00")
        
       impresion.Cell(impresion.Rows - 1, 9).text = Format(NETO, "###,###,##0")
        
'        If Option1.Value = True Then
'            linea2 = impresion.Rows - 1
'            For o = 2 To cantidaddeprecios
'            linea2 = linea2 + 1
'            impresion.AddItem ""
'            impresion.Cell(linea2, 7).text = "PRECIO X " + Str(PRECIOS(o, 4))
'            impresion.Cell(linea2, 9).text = Format(NPRECIOS(o, 1), "$ ###,###,##0")
'            impresion.Cell(linea2, 11).text = Format(NPRECIOS(o, 2), "$ ###,###,##0")
'            MO2 = margenoriginal * PORCENTAJES(o) / 100
'            impresion.Cell(linea2, 9).text = Format(MO2, "% ##0.00")
'            pventa = NPRECIOS(o, 1)
'            margen = (((pventa / pcosto) - 1) * 100)
'            impresion.Cell(linea2, 10).text = Format(margen, "% ##0.00")
'            pventa = NPRECIOS(o, 2)
'            margen = (((NPRECIOS(o, 1) / NPRECIOS(o, 2)) - 1) * 100)
'            impresion.Cell(linea2, 12).text = Format(margen, "% ##0.00")
'
'        Next o
        Call leerOFERTAS(codigo, pcosto, margenoriginal)
        
        impresion.Range(1, 0, impresion.Rows - 1, impresion.Cols - 1).Borders(cellEdgeTop) = cellThin
        impresion.Range(1, 0, impresion.Rows - 1, impresion.Cols - 1).Borders(cellEdgeBottom) = cellThin
        impresion.Range(1, 0, impresion.Rows - 1, impresion.Cols - 1).Borders(cellEdgeLeft) = cellThin
        impresion.Range(1, 0, impresion.Rows - 1, impresion.Cols - 1).Borders(cellEdgeRight) = cellThin
        impresion.Range(1, 0, impresion.Rows - 1, impresion.Cols - 1).Borders(cellInsideVertical) = cellThin
        impresion.Range(1, 0, impresion.Rows - 1, impresion.Cols - 1).Borders(cellInsideHorizontal) = cellThin
    
        
        'impresion.Range(1, 1, impresion.Rows - 1, 3).Alignment = cellLeftCenter
        
        ' impresion.Range(1, 4, impresion.Rows - 1, impresion.Cols - 1).Alignment = cellRightGeneral
        impresion.Range(1, 0, impresion.Rows - 1, impresion.Cols - 1).FontSize = 7
        compradas = compradas + Grid3.Cell(i, 3).text
        total = total + Grid3.Cell(i, 7).text
PASO:
    Next i
    
    
    impresion.Range(0, 0, 0, impresion.Cols - 1).Borders(cellEdgeTop) = cellThin
    impresion.Range(0, 0, 0, impresion.Cols - 1).Borders(cellEdgeBottom) = cellThin
    impresion.Range(0, 0, 0, impresion.Cols - 1).Borders(cellEdgeLeft) = cellThin
    impresion.Range(0, 0, 0, impresion.Cols - 1).Borders(cellEdgeRight) = cellThin
    impresion.Range(0, 0, 0, impresion.Cols - 1).Borders(cellInsideVertical) = cellThin
    FINROW = impresion.Rows - 1
    impresion.Rows = impresion.Rows + 9
     If ICA <> 0 Then
        FINROW = FINROW + 1
        impresion.Range(FINROW, 5, FINROW, 6).Merge
        impresion.Cell(FINROW, 5).text = "IMPUESTO CARNE "
        impresion.Range(FINROW, 5, FINROW, 6).Alignment = cellLeftCenter
        impresion.Cell(FINROW, 7).text = Format(ICA, "$ ###,###,##0")
        impresion.Cell(FINROW, 7).Alignment = cellRightGeneral
    End If
    If IHA <> 0 Then
        FINROW = FINROW + 1
        impresion.Range(FINROW, 5, FINROW, 6).Merge
        impresion.Cell(FINROW, 5).text = "IMPUESTO HARINA"
        impresion.Range(FINROW, 5, FINROW, 6).Alignment = cellLeftCenter
        impresion.Cell(FINROW, 7).text = Format(IHA, "$ ###,###,##0")
        impresion.Cell(FINROW, 7).Alignment = cellRightGeneral
    End If
    
    FINROW = FINROW + 1
    impresion.Range(FINROW, 1, FINROW, 1).Borders(cellEdgeLeft) = cellThin
    impresion.Range(FINROW, 8, FINROW, impresion.Cols - 1).Borders(cellEdgeRight) = cellThin
    impresion.Range(FINROW, 1, FINROW, impresion.Cols - 1).Borders(cellEdgeTop) = cellThin
    impresion.Cell(FINROW, 2).text = "UNIDADES"
    impresion.Cell(FINROW, 3).text = compradas
    impresion.Cell(FINROW, 3).Alignment = cellRightGeneral
    total = total + ICA + IHA
        impresion.Range(FINROW, 5, FINROW, 6).Merge
        impresion.Cell(FINROW, 5).text = "TOTAL RECEPCION"
        impresion.Range(FINROW, 5, FINROW, 6).Alignment = cellLeftCenter
        impresion.Cell(FINROW, 7).text = Format(total, "$ ###,###,##0")
        impresion.Cell(FINROW, 7).Alignment = cellRightGeneral
    descuento = Round(((total * pdescuento / 100) + 0.5), 0)
'
'    FINROW = FINROW + 1
'
'    impresion.Range(FINROW, 1, FINROW, 1).Borders(cellEdgeLeft) = cellThin
'    impresion.Range(FINROW, 8, FINROW, impresion.Cols - 1).Borders(cellEdgeRight) = cellThin
'    impresion.Range(FINROW, 1, FINROW, impresion.Cols - 1).Borders(cellEdgeTop) = cellThin
'    impresion.Range(FINROW, 5, FINROW, 6).Merge
'    impresion.Cell(FINROW, 5).text = "(-)D.FIN " + Format(pdescuento, "%##.#0")
'    impresion.Range(FINROW, 5, FINROW + 1, 6).Alignment = cellLeftCenter
'    impresion.Cell(FINROW, 7).text = Format(descuento, "$ ###,###,##0")
'    impresion.Cell(FINROW, 7).Alignment = cellRightGeneral
'
'    total = total - descuento
'
'    FINROW = FINROW + 1
'
'    impresion.Range(FINROW, 1, FINROW, 1).Borders(cellEdgeLeft) = cellThin
'    impresion.Range(FINROW, 8, FINROW, impresion.Cols - 1).Borders(cellEdgeRight) = cellThin
'    impresion.Range(FINROW, 1, FINROW, impresion.Cols - 1).Borders(cellEdgeTop) = cellThin
'    impresion.Range(FINROW, 5, FINROW, 6).Merge
'    impresion.Cell(FINROW, 5).text = "TOTAL (-) DESCUENTO    "
'    impresion.Range(FINROW, 5, FINROW + 1, 6).Alignment = cellLeftCenter
'
'    impresion.Cell(FINROW, 7).text = Format(total, "$ ###,###,##0")
'    impresion.Cell(FINROW, 7).Alignment = cellRightGeneral
'
'
    FINROW = FINROW + 1

    impresion.Range(FINROW, 1, FINROW, 1).Borders(cellEdgeLeft) = cellThin
    impresion.Range(FINROW, 8, FINROW, impresion.Cols - 1).Borders(cellEdgeRight) = cellThin
    impresion.Range(FINROW, 1, FINROW, impresion.Cols - 1).Borders(cellEdgeTop) = cellThin
    impresion.Range(FINROW, 5, FINROW, 6).Merge
    impresion.Cell(FINROW, 5).text = "TOTAL FACTURAS"
    impresion.Range(FINROW, 5, FINROW, 6).Alignment = cellLeftCenter
  
   
    totalfacturas = 0
    For i = 1 To Grid1.Rows - 1
        totalfacturas = totalfacturas + Grid1.Cell(i, 9).text
    Next i
    impresion.Cell(FINROW, 7).text = Format(totalfacturas, "$ ###,###,##0")
    impresion.Cell(FINROW, 7).Alignment = cellRightCenter

    FINROW = FINROW + 1
    impresion.Range(FINROW, 1, FINROW, 1).Borders(cellEdgeLeft) = cellThin
    impresion.Range(FINROW, 8, FINROW, impresion.Cols - 1).Borders(cellEdgeRight) = cellThin
    impresion.Range(FINROW, 1, FINROW, impresion.Cols - 1).Borders(cellEdgeTop) = cellThin
    impresion.Range(FINROW, 5, FINROW, 6).Merge
    impresion.Cell(FINROW, 5).text = "DIFERENCIA"
    impresion.Range(FINROW, 5, FINROW, 6).Alignment = cellLeftCenter
    impresion.Cell(FINROW, 7).text = Format(totalfacturas - total, "$ ###,###,##0")
    impresion.Cell(FINROW, 7).Alignment = cellRightCenter
    impresion.Range(FINROW, 1, FINROW, 1).Borders(cellEdgeLeft) = cellThin
    impresion.Range(FINROW, 8, FINROW, impresion.Cols - 1).Borders(cellEdgeRight) = cellThin
    FINROW = FINROW + 1
    impresion.Range(FINROW, 1, FINROW, impresion.Cols - 1).Borders(cellEdgeTop) = cellThin
    impresion.Range(FINROW, 1, FINROW, impresion.Cols - 1).FontSize = 7
    
    
    
    FINROW = FINROW + 2
    impresion.Range(FINROW, 1, FINROW, 5).Merge
    impresion.Cell(FINROW, 1).Font.Size = 10
    impresion.Cell(FINROW, 1).Font.Bold = True
    impresion.Cell(FINROW, 1).text = "EL VALOR A CANCELAR SERA : $ ________________________"
     
      
     FINROW = FINROW + 1
 
    
    impresion.AddItem "DETALLE DE FACTURAS"
    impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, 9).Merge
    impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, 9).Alignment = cellCenterGeneral
    impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, 9).FontBold = True
    impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, 9).FontSize = 7
    
    impresion.AddItem "PROVEEDOR" & vbTab & "NOMBRE " & vbTab & "NUMERO" & vbTab & "FECHA" & vbTab & "NETO" & vbTab & "IVA" & vbTab & "EXENTO" & vbTab & "IMPTOS" & vbTab & "TOTAL", False
    impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, 9).FontSize = 7
    impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, 9).FontBold = True
    impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, 9).Alignment = cellCenterGeneral
    impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, 9).Borders(cellEdgeTop) = cellThin
    impresion.Cell(impresion.Rows - 1, 1).Border(cellEdgeLeft) = cellThin
    impresion.Cell(impresion.Rows - 1, 2).Border(cellEdgeLeft) = cellThin
    impresion.Cell(impresion.Rows - 1, 3).Border(cellEdgeLeft) = cellThin
    impresion.Cell(impresion.Rows - 1, 4).Border(cellEdgeLeft) = cellThin
    impresion.Cell(impresion.Rows - 1, 5).Border(cellEdgeLeft) = cellThin
    impresion.Cell(impresion.Rows - 1, 6).Border(cellEdgeLeft) = cellThin
    impresion.Cell(impresion.Rows - 1, 7).Border(cellEdgeLeft) = cellThin
    impresion.Cell(impresion.Rows - 1, 8).Border(cellEdgeLeft) = cellThin
    impresion.Cell(impresion.Rows - 1, 8).Border(cellEdgeLeft) = cellThin
    impresion.Cell(impresion.Rows - 1, 8).Border(cellEdgeRight) = cellThin
    impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, 9).Borders(cellEdgeBottom) = cellThin

    For k = 1 To 5
    
    totales(k) = 0
    Next k
    For i = 1 To Grid1.Rows - 1
        If Grid1.Cell(i, 11).text = "0" Then BONIFI = "" Else BONIFI = " BONIFICADO"
        impresion.AddItem Grid1.Cell(i, 1).text & BONIFI & vbTab & nombreproveedor.Caption & vbTab & Grid1.Cell(i, 2).text & vbTab & Format(Grid1.Cell(i, 3).text, "dd-mm-yyyy") & vbTab & Format(Grid1.Cell(i, 5).text, "###,###,###") & vbTab & Format(Grid1.Cell(i, 6).text, "###,###,###") & vbTab & Format(Grid1.Cell(i, 7).text, "###,###,###") & vbTab & Format(Grid1.Cell(i, 8).text, "###,###,###") & vbTab & Format(Grid1.Cell(i, 9).text, "###,###,###"), False
        totales(1) = totales(1) + CDbl(Grid1.Cell(i, 5).text)
        totales(2) = totales(2) + CDbl(Grid1.Cell(i, 6).text)
        totales(3) = totales(3) + CDbl(Grid1.Cell(i, 7).text)
        totales(4) = totales(4) + CDbl(Grid1.Cell(i, 8).text)
        totales(5) = totales(5) + CDbl(Grid1.Cell(i, 9).text)
      
        
        
        impresion.Cell(impresion.Rows - 1, 1).Alignment = cellLeftCenter
        impresion.Cell(impresion.Rows - 1, 2).Alignment = cellLeftCenter
        impresion.Cell(impresion.Rows - 1, 3).Alignment = cellCenterGeneral
        impresion.Cell(impresion.Rows - 1, 4).Alignment = cellCenterGeneral
        impresion.Range(impresion.Rows - 1, 5, impresion.Rows - 1, 9).Alignment = cellRightCenter
        impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, 9).FontSize = 7
    Next i
    
    impresion.AddItem "" & vbTab & "" & vbTab & vbTab & vbTab & Format(totales(1), "###,###,###") & vbTab & Format(totales(2), "###,###,###") & vbTab & Format(totales(3), "###,###,###") & vbTab & Format(totales(4), "###,###,###") & vbTab & Format(totales(5), "###,###,###"), False
    impresion.Range(impresion.Rows - 1, 5, impresion.Rows - 1, 9).Borders(cellEdgeTop) = cellThick
    
    
    
    impresion.Range(impresion.Rows - 1, 5, impresion.Rows - 1, 9).Alignment = cellRightCenter
    impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, 9).FontSize = 7
    FINROW = impresion.Rows
    impresion.Rows = impresion.Rows + 13
    
    impresion.Cell(FINROW + 2, 1).text = "AUTORIZADA:"
    impresion.Range(FINROW + 2, 2, FINROW + 2, 6).Merge
    
    impresion.Cell(FINROW + 2, 2).text = autorizadopor.Caption + " EL " + fechaautorizacion.Caption
    
    impresion.Cell(FINROW + 3, 1).text = "FORMA PAGO:"
    impresion.Range(FINROW + 3, 2, FINROW + 3, 6).Merge
    impresion.Cell(FINROW + 3, 2).text = TIPOPAGO.Caption
    
    impresion.Cell(FINROW + 5, 1).text = "PRONTO PAGO:"
    impresion.Range(FINROW + 5, 2, FINROW + 5, 6).Merge
    
    impresion.Cell(FINROW + 5, 2).text = Format(pdescuento, "%#0.0")
    impresion.Range(FINROW + 6, 2, FINROW + 6, 5).Merge
    
    impresion.Cell(FINROW + 6, 2).text = "******** OBSERVACIONES ***************"

    impresion.Range(FINROW + 7, 2, FINROW + 7, 5).Merge
    impresion.Range(FINROW + 8, 2, FINROW + 8, 5).Merge
    impresion.Range(FINROW + 9, 2, FINROW + 9, 5).Merge
    impresion.Range(FINROW + 10, 2, FINROW + 10, 5).Merge
    
    
    impresion.Cell(FINROW + 7, 2).text = Mid(OBSERVA, 1, 70)
    impresion.Cell(FINROW + 8, 2).text = Mid(OBSERVA, 71, 70)
    impresion.Cell(FINROW + 9, 2).text = Mid(OBSERVA, 141, 70)
    impresion.Cell(FINROW + 10, 2).text = Mid(OBSERVA, 211, 30)
    impresion.Range(FINROW + 12, 2, FINROW + 12, 5).Merge
    
    impresion.Cell(FINROW + 12, 2).text = "FIRMA RECEPCIONISTA"
   

    
    
    
    
    
    impresion.AutoRedraw = True
    
    
    impresion.PrintPreviewVirtualGrid
    
    
End Sub

Sub cabeza()
    Dim objReportTitle As FlexCell.ReportTitle
    
    impresion.ReportTitles.Clear
    'Report Title 1
    For k = 1 To 1
        Set objReportTitle = New FlexCell.ReportTitle
        objReportTitle.text = DATOSEMPRESA(k)
        objReportTitle.Font.Name = "verdana"
        objReportTitle.Font.Size = 8
        objReportTitle.Font.Italic = True
        objReportTitle.PrintOnAllPages = True
        objReportTitle.color = RGB(128, 0, 0)
        objReportTitle.Align = CellLeft
        impresion.ReportTitles.Add objReportTitle
    Next k
    Set objReportTitle = New FlexCell.ReportTitle
    objReportTitle.text = "RECEPCION ORDEN DE COMPRA Nº : " + dato1.text
    objReportTitle.Font.Name = "verdana"
    objReportTitle.Font.Size = 10
    objReportTitle.Font.Bold = True
    objReportTitle.PrintOnAllPages = True
    impresion.ReportTitles.Add objReportTitle

    Set objReportTitle = New FlexCell.ReportTitle
    objReportTitle.text = "PROVEEDOR :" & "   " & DATO5.text + "-" + DV.Caption
    objReportTitle.Font.Name = "verdana"
    objReportTitle.Font.Size = 8
    objReportTitle.Font.Bold = True
    objReportTitle.PrintOnAllPages = True
    objReportTitle.Align = CellLeft
    impresion.ReportTitles.Add objReportTitle

    Set objReportTitle = New FlexCell.ReportTitle
    objReportTitle.text = "NOMBRE :" & "   " & nombreproveedor.Caption
    objReportTitle.Font.Name = "verdana"
    objReportTitle.Font.Size = 8
    objReportTitle.Font.Bold = True
    objReportTitle.PrintOnAllPages = True
    objReportTitle.Align = CellLeft
    impresion.ReportTitles.Add objReportTitle

    Set objReportTitle = New FlexCell.ReportTitle
    objReportTitle.text = "DIRECCION :" & "   " & direccion
    objReportTitle.Font.Name = "verdana"
    objReportTitle.Font.Size = 8
    objReportTitle.Font.Bold = True
    objReportTitle.PrintOnAllPages = True
    objReportTitle.Align = CellLeft
    impresion.ReportTitles.Add objReportTitle

    Set objReportTitle = New FlexCell.ReportTitle
    objReportTitle.text = "COMUNA :" & "   " & comuna
    objReportTitle.Font.Name = "verdana"
    objReportTitle.Font.Size = 8
    objReportTitle.Font.Bold = True
    objReportTitle.PrintOnAllPages = True
    objReportTitle.Align = CellLeft
    impresion.ReportTitles.Add objReportTitle

    Set objReportTitle = New FlexCell.ReportTitle
    objReportTitle.text = "CIUDAD :" & " " & ciudad + "                                                                     FECHA :" + dato2.text + "-" + dato3.text + "-" + dato4.text
    
    objReportTitle.Font.Name = "verdana"
    objReportTitle.Font.Size = 8
    objReportTitle.Font.Bold = True
    objReportTitle.PrintOnAllPages = True
    objReportTitle.Align = CellLeft
    impresion.ReportTitles.Add objReportTitle
    
    Set objReportTitle = New FlexCell.ReportTitle
    objReportTitle.text = "BODEGA :" & "   " & BODEGARECEPCION.Caption
    objReportTitle.Font.Name = "verdana"
    objReportTitle.Font.Size = 8
    objReportTitle.Font.Bold = True
    objReportTitle.PrintOnAllPages = True
    objReportTitle.Align = CellLeft
    impresion.ReportTitles.Add objReportTitle
    
    With impresion.PageSetup
        .HeaderFont.Size = 6
        .Header = "                                                                                                                   PAGINAS &P/&N EMITIDO:&D USUARIO " + USUARIOSISTEMA
        .HeaderAlignment = cellCenter
        .HeaderFont.Name = "Verdana"
        .HeaderMargin = 2
        ' .Orientation = cellLandscape
         .Orientation = cellPortrait
        
        
        
        
        
        
        
    End With
End Sub

Sub retorno()
    Dim i As Integer
    If MODIFI = 1 Then
    BTMODIFICA_Click
    
    End If
   
    opciones.Visible = False
    Grid3.Rows = 1
    Grid1.Rows = 1
    Grid1.AddItem ""
    facturas.Visible = False
    bodega.Visible = False
  
    DV.Caption = ""
    nombreproveedor.Caption = ""
    dv2.Caption = ""
    nombrecontable.Caption = ""
    recepciona.Visible = False
    Command2.Visible = False
    NETO.Caption = "$ 0"
    Rem nombrebodega.Caption = ""
    total.Caption = "$ 0"
    dato1.Enabled = True
   
    
    MODIFI = 0
    dato1.SetFocus
    Command4.Visible = False
    
    Unload Me
    
End Sub

Sub ELIMINA()
    
    Call eliminaPago
    '
    Call Elimina_Impuestos(dato1.text)
    
    campos(0, 0) = "tipo"
    campos(1, 0) = "numero"
    campos(2, 0) = ""
    campos(0, 2) = "l_movimientos_detalle_" & localorden
    condicion = "tipo='" & "OC" & "' AND numero='" & dato1.text & "'"
    op = 4
    sqlconta.response = campos
    Set sqlconta.conexion = gestionrubro
    Call sqlconta.sqlconta(op, condicion)
End Sub
Sub Elimina_Impuestos(numeroorden As String)
    Dim csql As New rdoQuery
    Dim resultados As rdoResultset

    Set csql.ActiveConnection = gestionrubro
    csql.sql = "DELETE FROM l_ordendecompra_impuestos_" & localorden & " "
    csql.sql = csql.sql + "WHERE numeroorden = '" & numeroorden & "' "
    csql.Execute
    Call sincronizadatos(csql.sql, gestionrubro, "")
    
End Sub

'Private Sub eliminaDetalle()
'    Dim k As Integer
'    Dim canti As Double
'    Dim codigo As String
'    Dim fecha As Date
'
'    For k = 1 To Grid3.Rows - 1
'        Call desactualiza_stock("+", Grid3.Cell(k, 1).text, "N", "S", dato6.text, dato4.text, Grid3.Cell(k, 5).text, Grid3.Cell(k, 6).text, dato4.text & "-" & dato3.text & "-" & dato2.text, dato5.text & dv.Caption)
'    Next k
'
'End Sub

Private Sub eliminaPago()
    Dim ORDEN As String
    Call ConectarControlData2(ordenes, Servidor, clientesistema + "gestion" & rubro, Usuario, password, "SELECT DISTINCT oef.ordenenlazada FROM l_ordendecompra_enlace_factura_" & localorden & " AS oef, l_ordendecompra_detalle_facturas_" & localorden & " AS opm WHERE oef.ordenconfactura = '" & dato1.text & "' AND oef.ordenconfactura = opm.ordendecompra ORDER BY oef.ordenenlazada ASC")
    campos(0, 2) = "l_ordendecompra_enlace_factura_" & localorden
    If ordenes.Recordset.RecordCount > 0 Then
        ordenes.Recordset.MoveFirst
        While Not ordenes.Recordset.EOF
            ORDEN = ordenes.Recordset.Fields("ordenenlazada")
            condicion = "ordenconfactura = '" & dato1.text & "' AND ordenenlazada = '" & ORDEN & "'"
            op = 4
            sqlconta.response = campos
            Set sqlconta.conexion = gestionrubro
            Call sqlconta.sqlconta(op, condicion)
            ordenes.Recordset.MoveNext
        Wend
    Else
        condicion = "ordenenlazada = '" & dato1.text & "'"
        op = 4
        sqlconta.response = campos
        Set sqlconta.conexion = gestionrubro
        Call sqlconta.sqlconta(op, condicion)
    End If
    campos(0, 2) = "l_ordendecompra_detalle_facturas_" & localorden
    condicion = "ordendecompra = '" & dato1.text & "'"
    op = 4
    sqlconta.response = campos
    Set sqlconta.conexion = gestionrubro
    Call sqlconta.sqlconta(op, condicion)
End Sub

Private Sub opciones_GotFocus()
    MANUAL.SetFocus
End Sub

    Private Sub ayudaBodega(ByRef caja As TextBox)
        
    End Sub
    Private Sub ayudaBodega2(ByRef caja As TextBox)
        Dim campos As Variant
        Dim cfijo As Variant
        Dim largo As Variant
        
        campos = Array("codigobodega", "nombre")
        largo = Array("5n", "20s")
        cfijo = "local=" + localorden
        mensajeAyuda = "Recepción Orden de Compra - Bodega"
        cabezas = Array("Codigo", "Nombre")
    
        Call cargaAyudaT(Servidor, basedatos & rubro, Usuario, password, "r_maestrobodegas_" & rubro, BODE, campos, cfijo, largo, 2)
        Call leerBodega2(BODE)
    End Sub
    
    Private Sub leerBodega(ByVal codigo As String)
        campos(0, 0) = "nombre"
        campos(1, 0) = ""
        campos(0, 2) = "r_maestrobodegas_" & rubro
        condicion = "codigobodega = '" & codigo & "' AND local= '" & localorden & "' AND rubro = '" & rubro & "'"
        op = 5
        sqlconta.response = campos
        Set sqlconta.conexion = gestionrubro
        Call sqlconta.sqlconta(op, condicion)
        If sqlconta.status = 0 Then
            Call cargarBodega
            dato7.SetFocus
        Else
        End If
    End Sub
 Private Sub leerBodega2(ByVal codigo As String)
        campos(0, 0) = "nombre"
        campos(1, 0) = ""
        campos(0, 2) = "r_maestrobodegas_" & rubro
        condicion = "codigobodega = '" & codigo & "' AND local= '" & localorden & "' AND rubro = '" & rubro & "'"
        op = 5
        sqlconta.response = campos
        Set sqlconta.conexion = gestionrubro
        Call sqlconta.sqlconta(op, condicion)
        If sqlconta.status = 0 Then
           BODEGARECEPCION.Caption = sqlconta.response(0, 3)
           
        Else
            BODEGARECEPCION.Caption = ""
            BODE.SelStart = 0
            BODE.SelLength = Len(BODE.text)
            BODE.SetFocus
        End If
    End Sub

    Private Sub cargarBodega()
        
    End Sub

Sub leerecepcion()
    Dim lin As Integer
    Dim suma As Double
    Dim sql As String
   
        
        Call leerProveedor
        
        sql = "SELECT codigo,r_maestroproductos_fijo_" & rubro & ".descripcion,linea,cantidad,uxc,unidades,precio,total,bodega,fecha "
        sql = sql + "FROM r_maestroproductos_fijo_" & rubro & ",l_movimientos_detalle_" + localorden + " "
        sql = sql + "WHERE codigobarra=codigo and tipo='OC' AND numero='" + dato1.text + "' order by linea "
    
        Call ConectarControlData2(movi, Servidor, clientesistema + "gestion" & rubro, Usuario, password, sql)
        
        If Not movi.Recordset.EOF Then movi.Recordset.MoveFirst
            Grid3.Rows = movi.Recordset.RecordCount + 1
            suma = 0: lin = 0
            If movi.Recordset.EOF = False Then
            Rem dato6.text = movi.Recordset.Fields(8)
            dato2.text = Format(movi.Recordset.Fields(9), "dd")
            dato3.text = Format(movi.Recordset.Fields(9), "mm")
            dato4.text = Format(movi.Recordset.Fields(9), "yyyy")
            
            Rem BODEGARECEPCION.Caption = leerNombreBodega(dato6.text)
            Rem BODE.text = dato6.text
            End If
            
            While Not movi.Recordset.EOF
                lin = lin + 1
               Grid3.Cell(lin, 0).text = lin
                Grid3.Cell(lin, 1).text = movi.Recordset.Fields(0)
                Grid3.Cell(lin, 2).text = movi.Recordset.Fields(1)
                Grid3.Cell(lin, 3).text = movi.Recordset.Fields(3)
                Grid3.Cell(lin, 4).text = movi.Recordset.Fields(4)
                Grid3.Cell(lin, 5).text = movi.Recordset.Fields(5)
                Grid3.Cell(lin, 6).text = movi.Recordset.Fields(6)
                Grid3.Cell(lin, 7).text = movi.Recordset.Fields(7)
                Grid3.Cell(lin, 8).text = CDbl(movi.Recordset.Fields(7) / 1.19)
                
                Grid3.Column(3).Locked = False
                suma = suma + movi.Recordset.Fields(7)
                movi.Recordset.MoveNext
            Wend
        If lin <> 0 Then
            existe = "S"
            
            Grid3.Enabled = False
            opciones.Visible = True
            opciones.SetFocus
            calcularecepcion
            Command4.Visible = True
            
        End If
        
        If lin = 0 Then existe = "N"
End Sub

Private Sub leerProveedor()
    campos(0, 0) = "mp.rut"
    campos(1, 0) = "mp.nombre"
    campos(2, 0) = "mp.direccion"
    campos(3, 0) = "mp.comuna"
    campos(4, 0) = "mp.ciudad"
    campos(5, 0) = ""
  
    campos(0, 2) = clientesistema & "gestion" & rubro & ".r_maestroproveedores_" & rubro & " AS mp, " & clientesistema & "gestion" & rubro & ".l_ordendecompra_cabeza_" & localorden & " AS oc"
    condicion = "mp.rut = oc.proveedor AND oc.numero = '" & dato1.text & "'"
    op = 5
    sqlconta.response = campos
    Set sqlconta.conexion = gestionrubro
    Call sqlconta.sqlconta(op, condicion)
    If sqlconta.status = 0 Then
        DATO5.text = Left(sqlconta.response(0, 3), 9)
        DV.Caption = Right(sqlconta.response(0, 3), 1)
        nombreproveedor.Caption = sqlconta.response(1, 3)
        direccion = sqlconta.response(2, 3)
        comuna = sqlconta.response(3, 3)
        ciudad = sqlconta.response(4, 3)
    End If
End Sub

Private Sub MANUAL_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then Call opciones_FSCommand("retorno", "")
    If UCase(Chr(KeyAscii)) = "M" Then Call opciones_FSCommand("modifica", "")
    If UCase(Chr(KeyAscii)) = "E" Then Call opciones_FSCommand("elimina", "")
    If UCase(Chr(KeyAscii)) = "S" Then Call opciones_FSCommand("siguiente", "")
    If UCase(Chr(KeyAscii)) = "A" Then Call opciones_FSCommand("anterior", "")
    If UCase(Chr(KeyAscii)) = "R" Then Call opciones_FSCommand("retorno", "")
    If UCase(Chr(KeyAscii)) = "I" Then Call opciones_FSCommand("imprime", "")
End Sub

Private Sub Grid1_KeyDown(KeyCode As Integer, ByVal Shift As Integer)
    Dim i As Integer
    
    If KeyCode = 27 Then
    facturas.Visible = False
    bodega.Visible = True
    
    
    End If
    
    If KeyCode = 38 And Grid1.ActiveCell.row = Grid1.Rows - 1 Then sg = "S" Else sg = "N"
    If Grid1.ActiveCell.col = 1 And Mid(Grid1.Cell(Grid1.ActiveCell.row, 1).text, 1, 3) = "OE " And KeyCode = 13 Then
        For i = 3 To Grid1.Cols - 1
            Grid1.Column(i).Locked = True
        Next i
        'Grid1.Cell(0, 2).text = "ORDEN DE COMPRA"
    End If
    If Grid1.ActiveCell.col = 1 And Mid(Grid1.Cell(Grid1.ActiveCell.row, 1).text, 1, 3) <> "GD " And KeyCode = 13 Then
        For i = 3 To Grid1.Cols - 1
            Grid1.Column(i).Locked = False
        Next i
        Grid1.Column(8).Locked = False
        
        Grid1.Column(9).Locked = True
        
    End If
End Sub

Private Sub Grid1_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 And Grid1.ActiveCell.col = 3 Then
  Grid1.Cell(Grid1.ActiveCell.row, Grid1.ActiveCell.col).text = Format(fechasistema, "dd") + "-" + Format(fechasistema, "mm") + "-" + Format(fechasistema, "yyyy")
   End If
   If KeyAscii = 13 And Grid1.ActiveCell.col = 4 Then
  Grid1.Cell(Grid1.ActiveCell.row, Grid1.ActiveCell.col).text = Format(fechasistema, "dd") + "-" + Format(fechasistema, "mm") + "-" + Format(fechasistema, "yyyy")
   End If
   
    If FORMATOGRILLA(3, Grid1.ActiveCell.col) = "S" Then Grid1.Cell(Grid1.ActiveCell.row, Grid1.ActiveCell.col).text = UCase(Grid1.ActiveCell.text)
    If FORMATOGRILLA(3, Grid1.ActiveCell.col) = "N" Then snum = 1: KeyAscii = esNumero(KeyAscii)
    If FORMATOGRILLA(3, Grid1.ActiveCell.col) = "C" Then snum = 1: 'KeyAscii = esnumero(keyascii)
End Sub

Private Sub Grid1_LeaveCell(ByVal row As Long, ByVal col As Long, NewRow As Long, NewCol As Long, Cancel As Boolean)
    Dim TEXTO As String
    Dim i As Integer
    Dim total As Double
    
    If NewCol = 8 And col = 8 Then
        NewCol = 10
    End If
    If NewCol = 8 And col = 10 Then
        NewCol = 8
    End If
    
    
        If row > 0 And row < Grid1.Rows And NewCol = 8 Then
        If Mid(Grid1.Cell(NewRow, 1).text, 1, 3) <> "OE " And Mid(Grid1.Cell(NewRow, 1).text, 1, 3) <> "GD " Then
        
        detalleimpuestos.Show vbModal
        End If
        
        End If
        
        
        If row = 0 And col = 0 Then NewRow = 1: NewCol = 1: GoTo no
        If row = Grid1.Rows - 1 And col = Grid1.Cols - 1 And NewCol = 1 Then Grid1.Rows = Grid1.Rows + 1: NewRow = Grid1.Rows - 1
        For k = 1 To NewCol - 1
            If Grid1.Cell(NewRow, k).text = "" Then NewCol = k:   Exit For
        Next k

    If col = 2 And Grid1.Cell(Grid1.ActiveCell.row, 2).text <> "" Then
    pivote.MaxLength = 10
    pivote.text = Grid1.Cell(Grid1.ActiveCell.row, 2).text: Call ceros(pivote): Grid1.Cell(Grid1.ActiveCell.row, 2).text = pivote.text
    End If
    If col > 4 And col < 10 Then
        For i = 5 To 8
            If Grid1.Cell(Grid1.ActiveCell.row, i).text = "" Then
                Grid1.Cell(Grid1.ActiveCell.row, i).text = "0"
            End If
            total = total + CDbl(Grid1.Cell(Grid1.ActiveCell.row, i).text)
        Next i
        Grid1.Cell(Grid1.ActiveCell.row, 9).text = total
    End If
    
    Rem If Grid1.Cell(row, col).text = "" And newcol > col Then newcol = col: newrow = row
no:
End Sub

Private Sub modificarOrden()
    Dim sumalin As Integer
    
'    For k = 1 To Grid3.Rows - 1
'        If Val(Grid4.Cell(k, 3).text) = 0 Then GoTo no:
'        sumalin = sumalin + 1
'        lineas.text = sumalin
'        Call ceros(lineas)
'
'        campos(0, 0) = "cantidad"
'        campos(1, 0) = "unidades"
'        campos(2, 0) = "total"
'        campos(3, 0) = ""
'
'        campos(0, 1) = Grid3.Cell(k, 3).text
'        campos(1, 1) = Grid3.Cell(k, 5).text
'        campos(2, 1) = Grid3.Cell(k, 7).text
'        campos(3, 1) = ""
'
'        campos(0, 2) = "ordendecompra_detalle_" + localorden
'        condicion = "numero = '" & DATO1.text & "' AND codigo = '" & Grid3.Cell(k, 1).text & "' AND linea = '" & sumline & "'"
'        op = 3
'        sqlconta.response = campos
'        Set sqlconta.conexion = GESTION
'        Call sqlconta.sqlconta(op, condicion)
'no:
'    Next k



End Sub

Sub formatoImpresion()
    Rem DATOS DE LA COLUMNA
    impresion.DefaultFont.Size = 8
    impresion.DefaultFont.Bold = True
    
    FORMATOGRILLA(1, 1) = "CODIGO"
    FORMATOGRILLA(1, 2) = "DESCRIPCION"
    FORMATOGRILLA(1, 3) = "COMPRA"
    FORMATOGRILLA(1, 4) = "UxC"
    FORMATOGRILLA(1, 5) = "UNI"
    FORMATOGRILLA(1, 6) = "P.COSTO"
    FORMATOGRILLA(1, 7) = "TOTAL"
    FORMATOGRILLA(1, 9) = "P.VENTA SISTEMA"
    FORMATOGRILLA(1, 9) = "MARGEN"
    
    
    Rem LARGO DE LOS DATOS
    FORMATOGRILLA(2, 1) = "13"
    FORMATOGRILLA(2, 2) = "60"
    FORMATOGRILLA(2, 3) = "10"
    FORMATOGRILLA(2, 4) = "10"
    FORMATOGRILLA(2, 5) = "10"
    FORMATOGRILLA(2, 6) = "10"
    FORMATOGRILLA(2, 7) = "10"
    FORMATOGRILLA(2, 9) = "10"
    FORMATOGRILLA(2, 9) = "10"


    Rem TIPO DE DATOS C=CEROS,S=STRING,N=NUMERICO,D=DATE,
    FORMATOGRILLA(3, 1) = "N"
    FORMATOGRILLA(3, 2) = "S"
    FORMATOGRILLA(3, 3) = "N"
    FORMATOGRILLA(3, 4) = "N"
    FORMATOGRILLA(3, 5) = "N"
    FORMATOGRILLA(3, 5) = "N"
    FORMATOGRILLA(3, 6) = "N"
    FORMATOGRILLA(3, 7) = "N"
    FORMATOGRILLA(3, 8) = "N"
    FORMATOGRILLA(3, 9) = "N"
    
    Rem FORMATO GRILLA
    FORMATOGRILLA(4, 1) = "0000000000000"
    FORMATOGRILLA(4, 2) = ""
    FORMATOGRILLA(4, 3) = "###,###,##0"
    FORMATOGRILLA(4, 4) = "###,###,##0"
    FORMATOGRILLA(4, 5) = "###,###,##0"
    FORMATOGRILLA(4, 6) = "$ ###,###,##0.0"
    FORMATOGRILLA(4, 7) = "$ ###,###,##0.0"
    FORMATOGRILLA(4, 8) = "$ ###,###,##0"
    FORMATOGRILLA(4, 9) = "% ###,###,##0.00"
    
    Rem LOCCKED
    FORMATOGRILLA(5, 1) = "FALSE"
    FORMATOGRILLA(5, 2) = "FALSE"
    FORMATOGRILLA(5, 3) = "FALSE"
    FORMATOGRILLA(5, 4) = "FALSE"
    FORMATOGRILLA(5, 5) = "FALSE"
    FORMATOGRILLA(5, 6) = "FALSE"
    FORMATOGRILLA(5, 7) = "FALSE"
    FORMATOGRILLA(5, 8) = "FALSE"
    FORMATOGRILLA(5, 9) = "FALSE"

    Rem ANCHO
    FORMATOGRILLA(6, 1) = "13"
    FORMATOGRILLA(6, 2) = "30"
    FORMATOGRILLA(6, 3) = "8"
    FORMATOGRILLA(6, 4) = "5"
    FORMATOGRILLA(6, 5) = "8"
    FORMATOGRILLA(6, 6) = "8"
    FORMATOGRILLA(6, 7) = "10"
    FORMATOGRILLA(6, 9) = "8"
    FORMATOGRILLA(6, 9) = "8"

    impresion.Cols = 10
    impresion.Rows = 1

    impresion.AllowUserResizing = False
    impresion.DisplayFocusRect = False
    impresion.ExtendLastCol = True
    impresion.BoldFixedCell = False
    impresion.DrawMode = cellOwnerDraw
    impresion.Appearance = Flat
    impresion.ScrollBarStyle = Flat
    impresion.FixedRowColStyle = Flat
    impresion.BackColorFixed = RGB(90, 158, 214)
    impresion.BackColorFixedSel = RGB(110, 180, 214)
    impresion.BackColorBkg = RGB(90, 158, 214)
    impresion.BackColorScrollBar = RGB(231, 235, 247)
    impresion.BackColor1 = RGB(231, 235, 247)
    impresion.BackColor2 = RGB(239, 243, 255)
    impresion.GridColor = RGB(148, 190, 231)
    For k = 1 To impresion.Cols - 1
        impresion.Cell(0, k).text = FORMATOGRILLA(1, k)
        impresion.Column(k).Width = Val(FORMATOGRILLA(6, k)) * Grid3.DefaultFont.Size
        impresion.Column(k).MaxLength = Val(FORMATOGRILLA(2, k))
        impresion.Column(k).FormatString = FORMATOGRILLA(4, k)
        impresion.Column(k).Locked = FORMATOGRILLA(5, k)
        If FORMATOGRILLA(3, k) = "N" Then
            impresion.Column(k).Alignment = cellRightCenter
        Else
            impresion.Column(k).Alignment = cellLeftCenter
        End If
    Next k
    impresion.Column(0).Width = 0
    impresion.Range(0, 0, 0, impresion.Cols - 1).Alignment = cellCenterCenter
    'Grid3.Enabled = False
End Sub
Sub leerprecioventa(codigo, LINEAS)
    Dim resultados As rdoResultset
    Dim csql As New rdoQuery
    Dim rut As String
    Dim LINEA As Double
    
        Set csql.ActiveConnection = gestionrubro
        csql.sql = "SELECT preciosistema,preciopuntoventa "
        csql.sql = csql.sql + "FROM r_maestroproductos_precios_" + rubro + " "
        csql.sql = csql.sql + "WHERE codigo ='" + codigo + "' order by codigoprecio"
        csql.Execute
        LINEA = 0
   
        If csql.RowsAffected > 0 Then
            Set resultados = csql.OpenResultset
            
            While Not resultados.EOF
                LINEA = LINEA + 1
                NPRECIOS(LINEA, 1) = resultados(0)
                NPRECIOS(LINEA, 2) = resultados(1)
               
                resultados.MoveNext
            Wend
            resultados.Close
        Set resultados = Nothing
  
    
        End If
End Sub

Private Function leerpagos()
    Call ConectarControlData2(ordenes, Servidor, clientesistema + "gestion" & rubro, Usuario, password, "SELECT CONCAT(tipo, '" & vbTab & "', numero, '" & vbTab & "', fecha, '" & vbTab & "', vencimiento, '" & vbTab & "', neto, '" & vbTab & "', iva, '" & vbTab & "', exento, '" & vbTab & "', impuestos, '" & vbTab & "', total, '" & vbTab & "', categoria, '" & vbTab & "', bonificacion) AS item ,rut FROM l_ordendecompra_detalle_facturas_" & localorden & " WHERE ordendecompra = '" & dato1.text & "' order by linea ")
    Grid1.Rows = 1
    Grid1.AutoRedraw = False
    If ordenes.Recordset.RecordCount > 0 Then
        ordenes.Recordset.MoveFirst
        LBLRUT.Caption = ordenes.Recordset.Fields("rut")
        
        While Not ordenes.Recordset.EOF
                
            Grid1.AddItem ordenes.Recordset.Fields("item"), True
            ordenes.Recordset.MoveNext
        Wend
    End If
    Grid1.AutoRedraw = True
    Grid1.Refresh
End Function


Sub grabarPrecios(codigo, costo, margen)
    Dim preciocosto As Double
    Dim precioventa As Double
    Dim MARGENDEFINITIVO As Double
    Dim precioventa2 As Double
    Dim compara As String
    Dim compara2 As Double
    Dim pesos As Double
    
    Dim i As Integer
    Rem Call eliminaPrecios(codigo)
    campos(0, 0) = "local"
    campos(1, 0) = "codigo"
    campos(2, 0) = "codigoprecio"
    campos(3, 0) = "preciosistema"
    campos(4, 0) = "preciocosto"
    campos(5, 0) = "margen"
    campos(6, 0) = "fechavigencia"
    campos(7, 0) = ""
    campos(0, 1) = "00"
    campos(1, 1) = codigo
    campos(6, 1) = fechasistema
    
    
    For i = 1 To cantidaddeprecios
       
       If leermarca(codigo) = "00000" Then
       campos(2, 1) = PRECIOS(i, 1)
       preciocosto = costo
       MARGENDEFINITIVO = margen * PORCENTAJES(i) / 100
       precioventa = Int((costo * (1 + (MARGENDEFINITIVO) / 100)) + 0.5)
       If costo > 50000 Then
       compara = Str(precioventa)
        pesos = CDbl(Right(precioventa, 1))
        If pesos = 0 Then precioventa = precioventa - 1
        If pesos < 3 And pesos <> 0 Then precioventa = precioventa - pesos - 1
        If pesos > 3 And pesos < 5 Then precioventa = precioventa - pesos + 5
        If pesos > 7 And pesos <= 9 Then precioventa = precioventa - pesos + 9
        If pesos > 5 And pesos < 8 Then precioventa = precioventa - (pesos - 5)
       End If
        compara = Str(precioventa)
        pesos = CDbl(Right(precioventa, 1))
        If pesos = 0 Then precioventa = precioventa - 1
        Rem If pesos < 3 And pesos <> 0 Then precioventa = precioventa - pesos - 1
        If pesos > 0 And pesos < 5 Then precioventa = precioventa - pesos + 5
        If pesos > 5 And pesos <= 9 Then precioventa = precioventa - pesos + 9
        Rem If pesos > 5 And pesos < 8 Then precioventa = precioventa - (pesos - 5)
       campos(3, 1) = precioventa
        
        
       
       
       campos(4, 1) = Str(costo)
       campos(5, 1) = Str(MARGENDEFINITIVO)
    
       campos(0, 2) = "r_maestroproductos_precios_" & rubro
       condicion = "codigo='" + codigo + "' and codigoprecio='" + PRECIOS(i, 1) + "' "
       op = 3
       Set sqlconta.conexion = gestionrubro
       sqlconta.response = campos
       Call sqlconta.sqlconta(op, condicion)
       
       End If
       
    Next i
End Sub
Sub modificacostos(codigo, costo)
    Rem Call eliminaPrecios(codigo)
        campos(0, 0) = "pcosto"
        campos(1, 0) = ""
        
        campos(0, 1) = Str(costo)
        campos(0, 2) = "r_maestroproductos_fijo_" & rubro
        condicion = "codigobarra='" + codigo + "' "
        op = 3
        Set sqlconta.conexion = gestionrubro
        sqlconta.response = campos
        Call sqlconta.sqlconta(op, condicion)
        
       
    
End Sub


Sub eliminaPrecios(codigo)
    campos(0, 2) = "r_maestroproductos_precios_" & rubro
    condicion = "codigo = '" & codigo & "'"
    op = 4
    sqlconta.response = campos
    Set sqlconta.conexion = gestionrubro
    Call sqlconta.sqlconta(op, condicion)
End Sub



Sub tiposdeprecios()
    Dim resultados As rdoResultset
    Dim csql As New rdoQuery
    Dim rut As String
    Dim LINEA As Double
    
        Set csql.ActiveConnection = gestion
        csql.sql = "SELECT codigo,nombre,porcentajedelmargen,unidades "
        csql.sql = csql.sql + "FROM g_maestrodetiposdeprecios ORDER BY codigo "
        'cSql.SQL = cSql.SQL + "WHERE local='" + codigoempresa + "' order by codigobodega"
        csql.Execute
        LINEA = 0
   
        If csql.RowsAffected > 0 Then
            Set resultados = csql.OpenResultset
            
            While Not resultados.EOF
                LINEA = LINEA + 1
                PRECIOS(LINEA, 1) = resultados(0)
                PRECIOS(LINEA, 2) = resultados(1)
                PRECIOS(LINEA, 3) = resultados(2)
                PRECIOS(LINEA, 4) = resultados(3)
                
                PORCENTAJES(LINEA) = resultados(2)
                resultados.MoveNext
            Wend
    cantidaddeprecios = LINEA
            resultados.Close
        Set resultados = Nothing
  
    
        End If
End Sub

Private Function leermargen(ByVal codigo As String) As String
    campos(0, 0) = "margen"
    campos(1, 0) = ""
  
    campos(0, 2) = "r_maestroproductos_fijo_" & rubro
    condicion = "codigobarra = '" & codigo & "' LIMIT 0,1"
    op = 5
    sqlconta.response = campos
    Set sqlconta.conexion = gestionrubro
    Call sqlconta.sqlconta(op, condicion)
    If sqlconta.status = 0 Then
        leermargen = sqlconta.response(0, 3)
    Else
        leermargen = "0"
    
    End If

End Function
Private Function leerimpuesto(ByVal codigo As String) As String
    campos(0, 0) = "codigoimpuesto"
    campos(1, 0) = ""
    campos(2, 0) = ""
  
    campos(0, 2) = "r_maestroproductos_fijo_" & rubro
    condicion = "codigobarra = '" & codigo & "' LIMIT 0,1"
    op = 5
    sqlconta.response = campos
    Set sqlconta.conexion = gestionrubro
    Call sqlconta.sqlconta(op, condicion)
    If sqlconta.status = 0 Then
        leerimpuesto = sqlconta.response(0, 3)
        
    Else
        leerimpuesto = "0"
    End If
    
    campos(0, 0) = "porcentaje"
    campos(1, 0) = ""
    
    campos(0, 2) = "g_maestroimpuestos"
    condicion = "codigo = '" & leerimpuesto & "' LIMIT 0,1"
    op = 5
    sqlconta.response = campos
    Set sqlconta.conexion = gestion
    Call sqlconta.sqlconta(op, condicion)
    If sqlconta.status = 0 Then
        impuesto = sqlconta.response(0, 3)
        
    Else
        impuesto = 0
    End If



End Function


Private Sub leerOFERTAS(codigo, pcosto, margenoriginal)
           Dim resultados As rdoResultset
            Dim csql As New rdoQuery
        Dim PVENTA As Double
        Dim margen As Double
        
        Set csql.ActiveConnection = gestionrubro
        csql.sql = "SELECT mpo.codigo,mpf.descripcion,mpo.cantidad,mpo.fechainicio,mpo.fechatermino,mpo.maximoxcliente,mpo.maximostockalaventa,mpo.preciooferta "
        csql.sql = csql.sql & "FROM r_maestroproductos_ofertas_" & rubro & " as mpo, r_maestroproductos_fijo_" & rubro & " as mpf "
        csql.sql = csql.sql & "where mpo.local='" + localorden + "' and mpf.codigobarra=mpo.codigo and mpo.fechainicio>='" + Format(fechasistema, "yyyy/mm/dd") + "' and mpo.fechatermino<='" + Format(fechasistema, "yyyy/mm/dd") + "' "
        csql.sql = csql.sql & "AND mpo.codigo='" + codigo + "'"
        
        csql.Execute
       
   
        If csql.RowsAffected > 0 Then
            Set resultados = csql.OpenResultset
            
            While Not resultados.EOF
               
            linea2 = impresion.Rows - 1
            
            linea2 = linea2 + 1
            impresion.AddItem ""
            
            impresion.Cell(linea2, 1).text = "OFERTA " + Format(resultados(3), "dd/mm/yyyy") + " HASTA " + Format(resultados(4), "dd/mm/yyyy")
            If pcosto = 0 Then pcosto = 1
            impresion.Cell(linea2, 5).text = "OFE x " + Str(resultados(2))
            impresion.Cell(linea2, 8).text = Format(resultados(7), "$ ###,###,##0")
            
            impresion.Cell(linea2, 7).text = Format(margenoriginal, "% ##0.00")
            PVENTA = resultados(7)
            margen = (((PVENTA / pcosto) - 1) * 100)
            impresion.Cell(linea2, 6).text = Format(margen, "% ##0.00")
                            
                
                resultados.MoveNext
            
            Wend
        End If
    End Sub
'===
Private Function leermarca(ByVal codigo As String) As String
    Dim campos2(10, 10)
    campos2(0, 0) = "codigomarca"
    campos2(1, 0) = ""
    campos2(2, 0) = ""
  
    campos2(0, 2) = "r_maestroproductos_fijo_" & rubro
    condicion = "codigobarra = '" & codigo & "' LIMIT 0,1"
    op = 5
    sqlconta.response = campos2
    Set sqlconta.conexion = gestionrubro
    Call sqlconta.sqlconta(op, condicion)
    If sqlconta.status = 0 Then
        leermarca = sqlconta.response(0, 3)
        
    Else
        leermarca = "00000"
    End If
    


End Function
Private Function leernoactualiza(ByVal codigo As String) As String
    Dim campos2(10, 10)
    campos2(0, 0) = "noactualiza"
    campos2(1, 0) = ""
    campos2(2, 0) = ""
  
    campos2(0, 2) = "r_maestroproductos_fijo_" & rubro
    condicion = "codigobarra = '" & codigo & "' LIMIT 0,1"
    op = 5
    sqlconta.response = campos2
    Set sqlconta.conexion = gestionrubro
    Call sqlconta.sqlconta(op, condicion)
    If sqlconta.status = 0 Then
        leernoactualiza = sqlconta.response(0, 3)
    Else
        leernoactualiza = 0
    End If
End Function

Sub MODIFICAORDEN(numero, codigo, Cantidad, UXC, unidades, Precio, total, LINEA)
    Dim campos2(10, 10)
    campos2(0, 0) = "codigo"
    campos2(1, 0) = "cantidad"
    campos2(2, 0) = "uxc"
    campos2(3, 0) = "unidades"
    campos2(4, 0) = "precio"
    campos2(5, 0) = "total"
    campos2(6, 0) = ""
  
    campos2(0, 1) = codigo
    campos2(1, 1) = Cantidad
    campos2(2, 1) = UXC
    campos2(3, 1) = unidades
    campos2(4, 1) = Precio
    campos2(5, 1) = total
    
    campos2(0, 2) = "l_ordendecompra_detalle_" & localorden
    condicion = "numero = '" & numero & "' and linea='" + LINEA + "' "
    op = 3
    sqlconta.response = campos2
    Set sqlconta.conexion = gestionrubro
    Call sqlconta.sqlconta(op, condicion)

End Sub
Sub leeTIPOPAGO(codigo)
           Dim resultados As rdoResultset
            Dim csql As New rdoQuery
        
        Set csql.ActiveConnection = gestion
        csql.sql = "SELECT * "
        csql.sql = csql.sql & "FROM g_tiposdepagoproveedor "
        csql.sql = csql.sql & "where codigo='" + codigo + "' "
        
        csql.Execute
       
   
        If csql.RowsAffected > 0 Then
        Set resultados = csql.OpenResultset
            
        TIPOPAGO.Caption = resultados(1)
        End If
End Sub

