VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "msadodc.ocx"
Object = "{4777436C-EB8C-4596-98A8-EBCF98683969}#1.0#0"; "FlexCell.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{ADBBAED6-B16F-40DC-80DF-B44910CBA76C}#1.0#0"; "Frame-Xp.ocx"
Begin VB.Form interno01 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Emision Documentos Electronicos"
   ClientHeight    =   10680
   ClientLeft      =   435
   ClientTop       =   825
   ClientWidth     =   14775
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10680
   ScaleWidth      =   14775
   StartUpPosition =   2  'CenterScreen
   Begin XPFrame.FrameXp FrameQuickMenu 
      Height          =   615
      Left            =   120
      TabIndex        =   88
      Top             =   9960
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
      Alignment       =   1
      Begin VB.CommandButton botonmisfavoritos 
         Caption         =   "Mis Favoritos"
         Height          =   255
         Left            =   1800
         TabIndex        =   90
         Top             =   280
         Width           =   1335
      End
      Begin VB.CommandButton botonmisaccesos 
         Caption         =   "Permisos Modulo"
         Height          =   255
         Left            =   120
         TabIndex        =   89
         Top             =   280
         Width           =   1455
      End
   End
   Begin XPFrame.FrameXp FrameXp5 
      Height          =   2295
      Left            =   11400
      TabIndex        =   47
      Top             =   7440
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   4048
      BackColor       =   16761024
      Caption         =   "Datos Vale Credito"
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
      Begin VB.TextBox FOLIO 
         BackColor       =   &H00FF8080&
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
         Left            =   240
         MaxLength       =   10
         TabIndex        =   50
         Top             =   600
         Width           =   1695
      End
      Begin VB.TextBox FECHA 
         BackColor       =   &H00FF8080&
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
         Left            =   240
         MaxLength       =   10
         TabIndex        =   49
         Top             =   1800
         Width           =   1695
      End
      Begin VB.CommandButton Command6 
         Caption         =   "leer"
         Height          =   255
         Left            =   600
         TabIndex        =   48
         Top             =   960
         Width           =   975
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "FOLIO"
         Height          =   255
         Left            =   240
         TabIndex        =   52
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label Label16 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "FECHA"
         Height          =   255
         Left            =   240
         TabIndex        =   51
         Top             =   1560
         Width           =   1695
      End
   End
   Begin FlexCell.Grid Grid1 
      Height          =   150
      Left            =   11400
      TabIndex        =   36
      Top             =   9960
      Visible         =   0   'False
      Width           =   3120
      _ExtentX        =   5503
      _ExtentY        =   265
      Cols            =   5
      DefaultFontName =   "Arial"
      DefaultFontSize =   8.25
      Rows            =   30
   End
   Begin XPFrame.FrameXp FrameXp2 
      Height          =   1410
      Left            =   135
      TabIndex        =   17
      Top             =   810
      Width           =   14550
      _ExtentX        =   25665
      _ExtentY        =   2487
      BackColor       =   16761024
      Caption         =   "Datos Cliente"
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
      HabilitarArrastre=   -1  'True
      Begin VB.Label LBLCIUDAD 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   8640
         TabIndex        =   38
         Top             =   1080
         Width           =   7620
      End
      Begin VB.Label LBLGIRO 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   945
         TabIndex        =   37
         Top             =   1080
         Width           =   7620
      End
      Begin VB.Label LBLRUT 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   945
         TabIndex        =   23
         Top             =   360
         Width           =   2175
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "RUT     :"
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
         Height          =   195
         Left            =   90
         TabIndex        =   22
         Top             =   405
         Width           =   765
      End
      Begin VB.Label LBLDIRECCION 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   1305
         TabIndex        =   21
         Top             =   720
         Width           =   7620
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DIRECCION :"
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
         Height          =   195
         Left            =   90
         TabIndex        =   20
         Top             =   720
         Width           =   1155
      End
      Begin VB.Label LBLNOMBRE 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   5040
         TabIndex        =   19
         Top             =   360
         Width           =   6675
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "NOMBRE :"
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
         Height          =   195
         Left            =   4005
         TabIndex        =   18
         Top             =   405
         Width           =   930
      End
   End
   Begin XPFrame.FrameXp FrameXp3 
      Height          =   4035
      Left            =   135
      TabIndex        =   8
      Top             =   2280
      Width           =   14580
      _ExtentX        =   25718
      _ExtentY        =   7117
      BackColor       =   16773879
      Caption         =   "Detalle Factura"
      CaptionEstilo3D =   1
      BackColor       =   16773879
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
      Begin FlexCell.Grid Informe 
         Height          =   3645
         Left            =   120
         TabIndex        =   7
         Top             =   315
         Width           =   14385
         _ExtentX        =   25374
         _ExtentY        =   6429
         BackColor1      =   14211288
         Cols            =   3
         DefaultFontSize =   8.25
         Rows            =   2
         SelectionMode   =   1
         DateFormat      =   2
      End
      Begin MSAdodcLib.Adodc data 
         Height          =   330
         Left            =   0
         Top             =   6840
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
   End
   Begin XPFrame.FrameXp FrameXp1 
      Height          =   780
      Left            =   135
      TabIndex        =   9
      Top             =   0
      Width           =   14670
      _ExtentX        =   25876
      _ExtentY        =   1376
      BackColor       =   16761024
      Caption         =   "Datos Documento a Imprimir"
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
      Begin VB.CommandButton Command1 
         Caption         =   "Leer Documento"
         Height          =   375
         Left            =   12000
         TabIndex        =   35
         Top             =   360
         Width           =   2475
      End
      Begin VB.TextBox DT4 
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
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   10245
         MaxLength       =   2
         TabIndex        =   4
         Tag             =   "fecha"
         Top             =   315
         Width           =   420
      End
      Begin VB.TextBox DT5 
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
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   10665
         MaxLength       =   2
         TabIndex        =   5
         Tag             =   "fecha"
         Top             =   315
         Width           =   375
      End
      Begin VB.TextBox DT6 
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
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   11070
         MaxLength       =   4
         TabIndex        =   6
         Tag             =   "fecha"
         Top             =   315
         Width           =   615
      End
      Begin VB.TextBox DT3 
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
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   8235
         MaxLength       =   10
         TabIndex        =   3
         Tag             =   "fecha"
         Top             =   315
         Width           =   1230
      End
      Begin VB.TextBox DT2 
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
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   6840
         MaxLength       =   2
         TabIndex        =   2
         Tag             =   "fecha"
         Top             =   315
         Width           =   375
      End
      Begin VB.TextBox DT1 
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
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   4275
         MaxLength       =   2
         TabIndex        =   1
         Tag             =   "fecha"
         Text            =   "FV"
         Top             =   315
         Width           =   375
      End
      Begin VB.TextBox DT0 
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
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   765
         MaxLength       =   2
         TabIndex        =   0
         Tag             =   "fecha"
         Top             =   315
         Width           =   375
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "FECHA"
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
         Height          =   195
         Left            =   9540
         TabIndex        =   16
         Top             =   405
         Width           =   615
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "NUMERO"
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
         Height          =   195
         Left            =   7380
         TabIndex        =   15
         Top             =   405
         Width           =   825
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CAJA"
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
         Height          =   195
         Left            =   6300
         TabIndex        =   14
         Top             =   405
         Width           =   465
      End
      Begin VB.Label LBLTIP0 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   4635
         TabIndex        =   13
         Top             =   315
         Width           =   1590
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "TIPO"
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
         Height          =   195
         Left            =   3780
         TabIndex        =   12
         Top             =   360
         Width           =   450
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "LOCAL "
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
         Height          =   195
         Left            =   45
         TabIndex        =   11
         Top             =   360
         Width           =   660
      End
      Begin VB.Label LBLLOCAL 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   1215
         TabIndex        =   10
         Top             =   315
         Width           =   2400
      End
   End
   Begin XPFrame.FrameXp FrameXp4 
      Height          =   4290
      Left            =   120
      TabIndex        =   24
      Top             =   6360
      Width           =   14550
      _ExtentX        =   25665
      _ExtentY        =   7567
      BackColor       =   16761024
      Caption         =   "Datos Factura"
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
      HabilitarArrastre=   -1  'True
      Begin TabDlg.SSTab SSTab1 
         Height          =   4215
         Left            =   3600
         TabIndex        =   53
         Top             =   0
         Width           =   10575
         _ExtentX        =   18653
         _ExtentY        =   7435
         _Version        =   393216
         Style           =   1
         Tabs            =   2
         TabsPerRow      =   2
         TabHeight       =   520
         BackColor       =   16761024
         TabCaption(0)   =   "CARGO A CREDITO INTERNO "
         TabPicture(0)   =   "cargafactura.frx":0000
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "frmgastos"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "CARGO A CREDITO EXTERNO"
         TabPicture(1)   =   "cargafactura.frx":001C
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "FrameXp6"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).ControlCount=   1
         Begin XPFrame.FrameXp frmgastos 
            Height          =   3855
            Left            =   0
            TabIndex        =   54
            Top             =   480
            Width           =   10335
            _ExtentX        =   18230
            _ExtentY        =   6800
            BackColor       =   16744576
            Caption         =   "Analisis de gastos"
            CaptionEstilo3D =   1
            BackColor       =   16744576
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
            Begin VB.TextBox dato7 
               BackColor       =   &H00C0E0FF&
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
               Left            =   1800
               MaxLength       =   4
               TabIndex        =   85
               Top             =   960
               Width           =   615
            End
            Begin VB.TextBox DATO6 
               BackColor       =   &H00C0E0FF&
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
               Left            =   1800
               MaxLength       =   30
               TabIndex        =   82
               Top             =   2880
               Width           =   5415
            End
            Begin VB.TextBox DATO5 
               BackColor       =   &H00C0E0FF&
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
               Left            =   1800
               MaxLength       =   30
               TabIndex        =   80
               Top             =   2520
               Width           =   5415
            End
            Begin VB.TextBox dato1 
               BackColor       =   &H00C0E0FF&
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
               Left            =   1800
               MaxLength       =   8
               TabIndex        =   64
               Top             =   600
               Width           =   975
            End
            Begin VB.TextBox dato2 
               BackColor       =   &H00C0E0FF&
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
               Left            =   1800
               MaxLength       =   4
               TabIndex        =   63
               Top             =   1320
               Width           =   615
            End
            Begin VB.TextBox dato3 
               BackColor       =   &H00C0E0FF&
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
               Left            =   1800
               MaxLength       =   4
               TabIndex        =   62
               Top             =   1680
               Width           =   615
            End
            Begin VB.TextBox DATO0 
               BackColor       =   &H00C0E0FF&
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
               Left            =   1800
               MaxLength       =   8
               TabIndex        =   61
               Top             =   240
               Width           =   375
            End
            Begin VB.CommandButton Command2 
               Caption         =   "Grabar"
               Height          =   300
               Left            =   120
               TabIndex        =   60
               Top             =   3360
               Width           =   1320
            End
            Begin VB.CommandButton Command4 
               Caption         =   "Eliminar"
               Height          =   300
               Left            =   2880
               TabIndex        =   59
               Top             =   3360
               Width           =   1200
            End
            Begin VB.CommandButton Command5 
               Caption         =   "Imprimir"
               Height          =   300
               Left            =   4200
               TabIndex        =   58
               Top             =   3360
               Width           =   1200
            End
            Begin VB.TextBox dato4 
               BackColor       =   &H00C0E0FF&
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
               Left            =   1800
               MaxLength       =   30
               TabIndex        =   57
               Top             =   2160
               Width           =   5415
            End
            Begin VB.CommandButton Command3 
               Caption         =   "Modificar"
               Height          =   300
               Left            =   1560
               TabIndex        =   56
               Top             =   3360
               Width           =   1200
            End
            Begin VB.CommandButton Command7 
               Caption         =   "Retorno"
               Height          =   300
               Left            =   5520
               TabIndex        =   55
               Top             =   3360
               Width           =   1200
            End
            Begin VB.Label lblcrcc 
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
               Height          =   300
               Left            =   2520
               TabIndex        =   87
               Top             =   960
               Width           =   4695
            End
            Begin VB.Label Label28 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "CENTRO COSTO"
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
               Height          =   195
               Left            =   120
               TabIndex        =   86
               Top             =   960
               Width           =   1470
            End
            Begin VB.Label Label25 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "GLOSA GASTO"
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
               Height          =   195
               Left            =   120
               TabIndex        =   83
               Top             =   2880
               Width           =   1320
            End
            Begin VB.Label Label13 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "SOLICITADO POR"
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
               Height          =   195
               Left            =   120
               TabIndex        =   81
               Top             =   2520
               Width           =   1575
            End
            Begin VB.Label Label4 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "CUENTA MAYOR"
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
               Height          =   195
               Left            =   120
               TabIndex        =   73
               Top             =   600
               Width           =   1485
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "CENTROS GASTO"
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
               Height          =   195
               Left            =   120
               TabIndex        =   72
               Top             =   1320
               Width           =   1590
            End
            Begin VB.Label Label10 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "DETALLE GASTO"
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
               Height          =   195
               Left            =   120
               TabIndex        =   71
               Top             =   1680
               Width           =   1530
            End
            Begin VB.Label lblmayor 
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
               Height          =   300
               Left            =   2880
               TabIndex        =   70
               Top             =   600
               Width           =   4335
            End
            Begin VB.Label lblcentro 
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
               Height          =   300
               Left            =   2520
               TabIndex        =   69
               Top             =   1320
               Width           =   4695
            End
            Begin VB.Label lblgasto 
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
               Height          =   300
               Left            =   2520
               TabIndex        =   68
               Top             =   1680
               Width           =   4695
            End
            Begin VB.Label lblempresa 
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
               Height          =   300
               Left            =   2280
               TabIndex        =   67
               Top             =   240
               Width           =   4935
            End
            Begin VB.Label Label14 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "EMPRESA"
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
               Height          =   195
               Left            =   120
               TabIndex        =   66
               Top             =   240
               Width           =   900
            End
            Begin VB.Label Label26 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "AUTORIZADO POR"
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
               Height          =   195
               Left            =   120
               TabIndex        =   65
               Top             =   2160
               Width           =   1680
            End
         End
         Begin XPFrame.FrameXp FrameXp6 
            Height          =   3735
            Left            =   -74880
            TabIndex        =   74
            Top             =   360
            Width           =   10335
            _ExtentX        =   18230
            _ExtentY        =   6588
            BackColor       =   16744576
            Caption         =   "Cargo Directo a Cliente"
            CaptionEstilo3D =   1
            BackColor       =   16744576
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
            Begin VB.CommandButton Command12 
               Caption         =   "Retorno"
               Height          =   300
               Left            =   5520
               TabIndex        =   79
               Top             =   2520
               Width           =   1200
            End
            Begin VB.CommandButton Command11 
               Caption         =   "Modificar"
               Height          =   300
               Left            =   1560
               TabIndex        =   78
               Top             =   2520
               Width           =   1200
            End
            Begin VB.CommandButton Command10 
               Caption         =   "Imprimir"
               Height          =   300
               Left            =   4200
               TabIndex        =   77
               Top             =   2520
               Width           =   1200
            End
            Begin VB.CommandButton Command9 
               Caption         =   "Eliminar"
               Height          =   300
               Left            =   2880
               TabIndex        =   76
               Top             =   2520
               Width           =   1200
            End
            Begin VB.CommandButton Command8 
               Caption         =   "Grabar"
               Height          =   300
               Left            =   120
               TabIndex        =   75
               Top             =   2520
               Width           =   1320
            End
            Begin VB.Label Label27 
               BackStyle       =   0  'Transparent
               Caption         =   "AL PRESIONAR GRABAR SE CARGARA AUTOMATICAMENTE EL CREDITO"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   1095
               Left            =   240
               TabIndex        =   84
               Top             =   840
               Width           =   4695
            End
         End
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "VINOS"
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
         Height          =   195
         Left            =   120
         TabIndex        =   46
         Top             =   1800
         Width           =   585
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "HARINA"
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
         Height          =   195
         Left            =   120
         TabIndex        =   45
         Top             =   2880
         Width           =   720
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CARNE"
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
         Height          =   195
         Left            =   120
         TabIndex        =   44
         Top             =   2520
         Width           =   645
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "REFRESCOS"
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
         Height          =   195
         Left            =   120
         TabIndex        =   43
         Top             =   2160
         Width           =   1125
      End
      Begin VB.Label lbltotal 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000007&
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
         ForeColor       =   &H0017DCEC&
         Height          =   300
         Left            =   2040
         TabIndex        =   42
         Top             =   3240
         Width           =   1455
      End
      Begin VB.Label lblica 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000007&
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
         ForeColor       =   &H0017DCEC&
         Height          =   300
         Left            =   2040
         TabIndex        =   41
         Top             =   2880
         Width           =   1455
      End
      Begin VB.Label lblvinos 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000007&
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
         ForeColor       =   &H0017DCEC&
         Height          =   300
         Left            =   2040
         TabIndex        =   40
         Top             =   1800
         Width           =   1455
      End
      Begin VB.Label lblrefrescos 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000007&
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
         ForeColor       =   &H0017DCEC&
         Height          =   300
         Left            =   2040
         TabIndex        =   39
         Top             =   2160
         Width           =   1455
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "TOTAL :"
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
         Height          =   195
         Left            =   150
         TabIndex        =   34
         Top             =   3240
         Width           =   735
      End
      Begin VB.Label lbliha 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000007&
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
         ForeColor       =   &H0017DCEC&
         Height          =   300
         Left            =   2040
         TabIndex        =   33
         Top             =   2520
         Width           =   1455
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "LICORES"
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
         Height          =   195
         Left            =   120
         TabIndex        =   32
         Top             =   1440
         Width           =   810
      End
      Begin VB.Label lbllicores 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000007&
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
         ForeColor       =   &H0017DCEC&
         Height          =   300
         Left            =   2040
         TabIndex        =   31
         Top             =   1440
         Width           =   1455
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "EXENTO :"
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
         Height          =   195
         Left            =   120
         TabIndex        =   30
         Top             =   1080
         Width           =   885
      End
      Begin VB.Label lblexento 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000007&
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
         ForeColor       =   &H0017DCEC&
         Height          =   300
         Left            =   2040
         TabIndex        =   29
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "I.V.A :"
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
         Height          =   195
         Left            =   120
         TabIndex        =   28
         Top             =   720
         Width           =   555
      End
      Begin VB.Label lbliva 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000007&
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
         ForeColor       =   &H0017DCEC&
         Height          =   300
         Left            =   2040
         TabIndex        =   27
         Top             =   720
         Width           =   1455
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "NETO :"
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
         Height          =   195
         Left            =   120
         TabIndex        =   26
         Top             =   405
         Width           =   645
      End
      Begin VB.Label lblneto 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000007&
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
         ForeColor       =   &H0017DCEC&
         Height          =   300
         Left            =   2040
         TabIndex        =   25
         Top             =   360
         Width           =   1455
      End
   End
End
Attribute VB_Name = "interno01"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Private ESSUPER As String
    
    Private FORMATOGRILLA(20, 20)
    Private fecha1 As String
    Private fecha2 As String
    Private rutemisos As String
    Private empresaemisor As String
    
Private Sub Command1_Click()
If LBLLOCAL.Caption <> "" Then
Call cargadocumento(DT0.text, DT1.text, DT3.text, DT2.text, DT6.text + "-" + DT5.text + "-" + DT4.text)

End If

End Sub

Private Sub Command10_Click()
IMPRIMEcreditoDIRECTO
retorno

End Sub

Private Sub COMMAND2_Click()

Dim EXENTO As Double
FOLIO.text = LEERULTIMOFOLIO
lblmayor.Caption = leerNombreMayor(dato1)
            

If LBLEMPRESA.Caption <> "" And lblmayor.Caption <> "" And lblmayor.Caption <> "" And lblcentro.Caption <> "" Then
ESSUPER = "0"
If MsgBox("CORRESPONDE A PRODUCTOS DE SUPERMERCADO", vbYesNo) = vbYes Then
ESSUPER = "1"
End If


Call grabar(FOLIO.text, fecha.text, DT1.text, DT3.text, DT6.text + "-" + DT5.text + "-" + DT4.text, DT0.text, dato2.text, dato1.text, dato3.text, dato0.text, lblneto.Caption, lbliva.Caption, lblexento.Caption, lbllicores.Caption, lblvinos.Caption, lblrefrescos.Caption, lbliha.Caption, lblica.Caption, lbltotal.Caption, DT2.text, LBLRUT.Caption, "S", dato4.text)
EXENTO = CDbl(lblrefrescos.Caption) + CDbl(lblvinos.Caption) + CDbl(lbllicores.Caption) + CDbl(lbliha.Caption) + CDbl(lblica.Caption)
Rem Call grabafactura(leetipofactura(DT0.text, DT1.text, DT3.text, DT2.text, DT6.text + "-" + DT5.text + "-" + DT4.text), DT3.text, DT6.text + "-" + DT5.text + "-" + DT4.text, DT6.text + "-" + DT5.text + "-" + DT4.text, LBLRUT.Caption, lblneto.Caption, lbliva.Caption, Str(EXENTO), "0", lbltotal.Caption, Format(Date, "yyyy"), Format(Date, "mm"), "", ESSUPER)
Call CARGACREDITO(FOLIO.text, lbltotal.Caption, Format(fecha.text, "yyyy-mm-dd"), Mid(LBLRUT.Caption, 1, 9) + Mid(LBLRUT.Caption, 11, 1), DT3.text, DT6.text + "-" + DT5.text + "-" + DT4.text, lblNOMBRE.Caption, Format(fechasistema, "mm"), Format(fechasistema, "yyyy"))

Call Command5_Click

retorno
Else
MsgBox "FALTA ALGUN DATO REVISAR"
End If


End Sub
Sub CARGACREDITO(numero, monto, fecha, rut, NUMERODOC, FECHADOC, NOMBRE, MES, año)
Dim CUENTA1 As String
Dim CUENTA2 As String
CUENTA1 = "11200029"
CUENTA2 = "11500190"
monto = Replace(monto, ".", "")
Call grabarcomprobante_lineas("CD", numero, "001", fecha, CUENTA1, "", rut, "", "CREDITO " + NOMBRE, "FV", NUMERODOC, FECHADOC, FECHADOC, monto, "D", USUARIOSISTEMA, Format(MES, "00"), año, Format(fechasistema, "yyyy-mm-dd"), Time, rut, "", "")
Call grabarcomprobante_lineas("CD", numero, "002", fecha, CUENTA2, "", rut, "", "CREDITO " + NOMBRE, "FV", NUMERODOC, FECHADOC, FECHADOC, monto, "H", USUARIOSISTEMA, Format(MES, "00"), año, Format(fechasistema, "yyyy-mm-dd"), Time, rut, "", "")
End Sub
Private Sub Command3_Click()
If Verifica_Permiso(Me.Caption, "Modifica") Then
Call ELIMINAR(leetipofactura(DT0.text, DT1.text, DT3.text, DT2.text, DT6.text + "-" + DT5.text + "-" + DT4.text), DT3.text, leerdatoslocal(DT0.text, "rut"))
dato1.text = ""
Else
MsgBox mensaje_nopermiso, vbCritical + vbOKOnly, "Permiso Denegado"
End If

End Sub

Private Sub Command4_Click()
If Verifica_Permiso(Me.Caption, "elimina") Then
Call ELIMINAR("1", DT3.text, leerdatoslocal(DT0.text, "rut"))
retorno
Else
 MsgBox mensaje_nopermiso, vbCritical + vbOKOnly, "Permiso Denegado"
End If


End Sub

Private Sub Command5_Click()
IMPRIMEcredito

End Sub

Private Sub Command6_Click()
Call leerfolio(FOLIO.text)
End Sub

Private Sub Command7_Click()
retorno

End Sub

Private Sub command8_Click()
Dim EXENTO As Double

Call grabar(FOLIO.text, fecha.text, DT1.text, DT3.text, DT6.text + "-" + DT5.text + "-" + DT4.text, DT0.text, dato2.text, dato1.text, dato3.text, dato0.text, lblneto.Caption, lbliva.Caption, lblexento.Caption, lbllicores.Caption, lblvinos.Caption, lblrefrescos.Caption, lbliha.Caption, lblica.Caption, lbltotal.Caption, DT2.text, LBLRUT.Caption, "N", dato4.text)
Call CARGACREDITO(FOLIO.text, lbltotal.Caption, Format(fecha.text, "yyyy-mm-dd"), Mid(LBLRUT.Caption, 1, 9) + Mid(LBLRUT.Caption, 11, 1), DT3.text, DT6.text + "-" + DT5.text + "-" + DT4.text, lblNOMBRE.Caption, DT5.text, DT6.text)
Call Command10_Click

retorno

End Sub

Private Sub Command9_Click()
Call Command4_Click
retorno

End Sub

'****************************************************************************
'Manejo de los Controles
'****************************************************************************
    '****************************************************************************
    'GOTFOCUS
    '****************************************************************************
    Private Sub dato1_GotFocus()
        Call cargatexto(dato1)
    End Sub

Private Sub dato1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF2 Then Call ayudamayor(dato1)
End Sub

    Private Sub dato2_GotFocus()
        Call cargatexto(dato2)
    End Sub
    
Private Sub dato2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF2 Then Call ayudatipoconsumo(dato2)
End Sub
Private Sub dato7_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF2 Then Call ayudacrcc(dato7)
End Sub

Sub ayudatipoconsumo(ByRef caja As TextBox)
    Dim campos As Variant
    Dim cfijo As Variant
    Dim largo As Variant
    campos = Array("codigo", "nombre")
    largo = Array("11s", "40s")
    cfijo = "codigo like '%%'"
    cabezas = Array("Codigo", "Nombre")
    mensajeAyuda = "Ayuda de Centros de Gastos"
       
    Call cargaAyudaT(Servidor, clientesistema & "conta", Usuario, password, ".presupuesto_centros", caja, campos, cfijo, largo, 2)
    caja.Enabled = True
    caja.SetFocus
End Sub
Sub ayudacrcc(ByRef caja As TextBox)
    Dim campos As Variant
    Dim cfijo As Variant
    Dim largo As Variant
    campos = Array("codigo", "nombre")
    largo = Array("11s", "40s")
    cfijo = "codigo like '%%' and año='" + Format(fechasistema, "yyyy") + "' "
    cabezas = Array("Codigo", "Nombre")
    mensajeAyuda = "Ayuda de Centros de Costos"
       
    Call cargaAyudaT(Servidor, clientesistema & "conta" + dato0.text, Usuario, password, ".centrosdecosto", caja, campos, cfijo, largo, 2)
    caja.Enabled = True
    caja.SetFocus
End Sub


    Private Sub dato3_GotFocus()
        Call cargatexto(dato3)
    End Sub
    
    Private Sub dato4_GotFocus()
        Call cargatexto(dato4)
    End Sub
    
    '****************************************************************************
    'GOTFOCUS
    '****************************************************************************
    
    '****************************************************************************
    'KEYDOWN
    '****************************************************************************
    '****************************************************************************
    'KEYDOWN
    '****************************************************************************
    
    '****************************************************************************
    'KEYPRESS
    '****************************************************************************
    Private Sub dato1_KeyPress(KeyAscii As Integer)
        KeyAscii = esNumero(KeyAscii)
        If KeyAscii = 13 Then
            Call ceros(dato1)
            If leerNombreMayor(dato1) <> "" Then
            lblmayor.Caption = leerNombreMayor(dato1)
            
            dato7.SetFocus
            Else
            dato1.SetFocus
            
            End If
            
        End If
    End Sub
    Private Sub dato7_KeyPress(KeyAscii As Integer)
        KeyAscii = esNumero(KeyAscii)
        If KeyAscii = 13 Then
            Call ceros(dato7)
            If leerNOMBREcrcc2(dato7, dato0.text) <> "" Then
            lblcrcc.Caption = leerNOMBREcrcc2(dato7, dato0.text)
            
            dato2.SetFocus
            Else
            dato7.SetFocus
            
            End If
            
        End If
    End Sub

Sub ayudamayor(ByRef caja As TextBox)
    Dim campos As Variant
    Dim cfijo As Variant
    Dim largo As Variant
    
    
    campos = Array("codigo", "nombre")
    largo = Array("12s", "40s")
    cfijo = "año='" + Format(fechasistema, "yyyy") + "' AND MID(codigo,5,4)<>'0000'"
    cabezas = Array("codigo", "nombre")
    mensajeAyuda = "Ayuda Cuentas del Mayor"
    basebus = clientesistema + "conta" + dato0.text
    
    Call cargaAyudaT(Servidor, basebus, Usuario, password, "cuentasdelmayor", caja, campos, cfijo, largo, 2)
  
    caja.SetFocus
    
no:
End Sub
    
    Private Sub dato2_KeyPress(KeyAscii As Integer)
        KeyAscii = esNumero(KeyAscii)
        If KeyAscii = 13 Then
            Call ceros(dato2)
            If leerNOMBREgastos(dato2) <> "" Then
            lblcentro.Caption = leerNOMBREgastos(dato2)
            dato3.SetFocus
            Else
            dato2.SetFocus
            
            
            End If
            
        End If
    End Sub
    
Private Sub dato3_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF2 Then Call ayudaanalisis(dato3)

End Sub
Sub ayudaanalisis(ByRef caja As TextBox)
    Dim campos As Variant
    Dim cfijo As Variant
    Dim largo As Variant
    campos = Array("codigo", "nombre")
    largo = Array("6n", "40s")
    cfijo = "cuenta='" & dato1.text & "' "
    cabezas = Array("codigo", "nombre")
    mensajeAyuda = "Analsis " + lblmayor.Caption
    Call cargaAyudaT(Servidor, clientesistema + "conta", Usuario, password, "presupuesto_detalle", caja, campos, cfijo, largo, 2)
    caja.Enabled = True
    caja.SetFocus

no:

End Sub


    Private Sub dato3_KeyPress(KeyAscii As Integer)
        KeyAscii = esNumero(KeyAscii)
        If KeyAscii = 13 Then
            Call ceros(dato3)
            If leernombreanalisis(dato1, dato3) <> "" Or Mid(dato1.text, 1, 1) < "3" Then
            lblgasto.Caption = leernombreanalisis(dato1, dato3)
            dato4.SetFocus
            
            
            Else
            dato3.SetFocus
            
            End If
            
            
        End If
    End Sub
    
Private Sub dato4_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
If KeyAscii = 13 Then
DATO5.SetFocus


End If

End Sub
Private Sub dato5_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
If KeyAscii = 13 Then
dato6.SetFocus


End If

End Sub
Private Sub dato6_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
If KeyAscii = 13 Then
Call Command2.SetFocus

End If

End Sub


Private Sub DT0_GotFocus()
FOLIO.text = LEERULTIMOFOLIO
End Sub

Private Sub DT0_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF2 Then
Call ayudalocales(DT0)
End If

End Sub

    '****************************************************************************
    'KEYPRESS
    '****************************************************************************

Private Sub DT0_KeyPress(KeyAscii As Integer)
KeyAscii = esNumero(KeyAscii)
If KeyAscii = 13 Then
DT0.text = Format(DT0.text, "00")
If leernombrelocal(DT0.text) <> "" Then
LBLLOCAL.Caption = leernombrelocal(DT0.text)
DT1.SetFocus
Else
MsgBox ("NUMERO DE LOCAL NO EXISTE ")
DT0.SetFocus

End If


End If

End Sub

Private Sub DT1_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))

If KeyAscii = 13 And (DT1.text = "FV" Or DT1.text = "BV") Then
DT2.SetFocus

End If

End Sub

Private Sub DT2_KeyPress(KeyAscii As Integer)
KeyAscii = esNumero(KeyAscii)
If KeyAscii = 13 And DT2.text <> "" And DT2.text <> "00" Then
DT2.text = Format(DT2.text, "00")
DT3.SetFocus
Else
DT2.SetFocus

End If

End Sub

Private Sub DT3_KeyPress(KeyAscii As Integer)
KeyAscii = esNumero(KeyAscii)
If KeyAscii = 13 And DT3.text <> "" And DT3.text <> "0000000000" Then
DT3.text = Format(DT3.text, "0000000000")
DT4.SetFocus

End If

End Sub
Private Sub DT4_KeyPress(KeyAscii As Integer)
KeyAscii = esNumero(KeyAscii)

If KeyAscii = 13 Then
DT4.text = Format(DT4.text, "00")
If DT4.text < "01" Or DT4.text > "30" Then
DT4.SetFocus
Else
DT5.SetFocus
End If
End If

End Sub

Private Sub DT5_KeyPress(KeyAscii As Integer)
KeyAscii = esNumero(KeyAscii)

If KeyAscii = 13 Then
DT5.text = Format(DT5.text, "00")
If DT5.text < "01" Or DT5.text > "12" Then
DT5.SetFocus
Else
DT6.SetFocus

End If
End If
End Sub


Private Sub DT6_KeyPress(KeyAscii As Integer)
KeyAscii = esNumero(KeyAscii)

If KeyAscii = 13 Then
DT6.text = Format(DT6.text, "0000")
Call Command1_Click

End If

End Sub

Private Sub FOLIO_KeyPress(KeyAscii As Integer)
KeyAscii = esNumero(KeyAscii)
If KeyAscii = 13 Then
DT5.text = Format(DT5.text, "00")
If DT5.text < "01" Or DT5.text > "12" Then
DT5.SetFocus
Else
DT6.SetFocus

End If
End If


End Sub

    Private Sub Form_Load()
        Call CENTRAR(Me)
        'cmbTipo.ListIndex = 0
        Call CARGAGRILLA(1, 7)
frmgastos.Enabled = False
FOLIO.text = LEERULTIMOFOLIO

fecha.text = Format(fechasistema, "dd-mm-yyyy")

    End Sub
    
    Private Sub cmdImprime_Click()
    End Sub
'****************************************************************************
'Manejo de los Controles
'****************************************************************************

Sub CARGAGRILLA(ByVal row As Long, ByVal col As Long)
    Dim i As Long
    Rem DATOS DE LA COLUMNA
'    Informe.DefaultFont.Size = 7.5
    FORMATOGRILLA(1, 1) = "CODIGO"
    FORMATOGRILLA(1, 2) = "CANTIDAD"
    FORMATOGRILLA(1, 3) = "DESCRIPCION"
    FORMATOGRILLA(1, 4) = "P/UNITARIO"
    FORMATOGRILLA(1, 5) = "DESCUENTO"
    FORMATOGRILLA(1, 6) = "TOTAL"
    
    
    Rem LARGO DE LOS DATOS
    FORMATOGRILLA(2, 1) = "10"
    FORMATOGRILLA(2, 2) = "50"
    FORMATOGRILLA(2, 3) = "10"
    FORMATOGRILLA(2, 4) = "7"
    FORMATOGRILLA(2, 5) = "9"
    FORMATOGRILLA(2, 6) = "8"
    FORMATOGRILLA(2, 7) = "8"
    FORMATOGRILLA(2, 8) = "8"
    FORMATOGRILLA(2, 9) = "8"
    
    Rem TIPO DE DATOS C=CEROS,S=STRING,N=NUMERICO,D=DATE,
    FORMATOGRILLA(3, 1) = "S"
    FORMATOGRILLA(3, 2) = "N"
    FORMATOGRILLA(3, 3) = "S"
    FORMATOGRILLA(3, 4) = "N"
    FORMATOGRILLA(3, 5) = "N"
    FORMATOGRILLA(3, 6) = "N"
    FORMATOGRILLA(3, 7) = "N"
    FORMATOGRILLA(3, 8) = "N"
    FORMATOGRILLA(3, 9) = "N"
    
    Rem FORMATO GRILLA
    FORMATOGRILLA(4, 1) = ""
    FORMATOGRILLA(4, 2) = ""
    FORMATOGRILLA(4, 3) = "###,##0.000"
    FORMATOGRILLA(4, 4) = "###,###,###"
    FORMATOGRILLA(4, 5) = "###,###,###"
    FORMATOGRILLA(4, 6) = "###,###,###"
    
    Rem LOCCKED
    FORMATOGRILLA(5, 1) = "TRUE"
    FORMATOGRILLA(5, 2) = "TRUE"
    FORMATOGRILLA(5, 3) = "TRUE"
    FORMATOGRILLA(5, 4) = "TRUE"
    FORMATOGRILLA(5, 5) = "TRUE"
    FORMATOGRILLA(5, 6) = "TRUE"
    FORMATOGRILLA(5, 7) = "FALSE"
    FORMATOGRILLA(5, 8) = "FALSE"
    FORMATOGRILLA(5, 9) = "FALSE"
    
    Rem ANCHO
    FORMATOGRILLA(6, 1) = "13"
    FORMATOGRILLA(6, 2) = "10"
    FORMATOGRILLA(6, 3) = "50"
    FORMATOGRILLA(6, 4) = "10"
    FORMATOGRILLA(6, 5) = "10"
    FORMATOGRILLA(6, 6) = "10"
    FORMATOGRILLA(6, 7) = "8"
    FORMATOGRILLA(6, 8) = "8"
    FORMATOGRILLA(6, 9) = "8"
    
    Informe.Cols = col
    Informe.Rows = row
    
    Informe.AllowUserResizing = False
    Informe.DisplayFocusRect = False
    Informe.ExtendLastCol = True
    Informe.BoldFixedCell = False
    Informe.DrawMode = cellOwnerDraw
    Informe.Appearance = Flat
    Informe.ScrollBarStyle = Flat
    Informe.FixedRowColStyle = Flat

    Informe.BackColorFixed = RGB(90, 158, 214)
    Informe.BackColorFixedSel = RGB(110, 180, 230)
    Informe.BackColorBkg = RGB(90, 158, 214)
    Informe.BackColorScrollBar = RGB(231, 235, 247)
    Informe.BackColor1 = RGB(231, 235, 247)
    Informe.BackColor2 = RGB(239, 243, 255)
    Informe.GridColor = RGB(148, 190, 231)
    Informe.Column(0).Width = 0
    
    For i = 1 To Informe.Cols - 1
        Informe.Cell(0, i).text = FORMATOGRILLA(1, i)
        Informe.Column(i).Width = Val(FORMATOGRILLA(6, i)) * Informe.DefaultFont.Size
        Informe.Column(i).MaxLength = Val(FORMATOGRILLA(2, i))
        Informe.Column(i).FormatString = FORMATOGRILLA(4, i)
        Informe.Column(i).Locked = FORMATOGRILLA(5, i)
        If FORMATOGRILLA(3, i) = "N" Then Informe.Column(i).Alignment = cellRightCenter
        If FORMATOGRILLA(3, i) = "S" Then Informe.Column(i).Alignment = cellLeftCenter
        If FORMATOGRILLA(3, i) = "D" Then Informe.Column(i).CellType = cellCalendar
    Next i
    Informe.Range(0, 0, 0, Informe.Cols - 1).Alignment = cellCenterCenter
End Sub

Private Sub cargadocumento(loc, tipo, numero, caja, fecha)
    Dim codigoempresa As String
    
    Dim resultados As rdoResultset
    Dim csql As New rdoQuery
    Set csql.ActiveConnection = contadb
    
    csql.sql = "SELECT dd.codigo, dd.cantidad,dd.descripcion,   dd.precio, dd.descuento,dd.total , dc.rut, dc.sucursal, dc.neto as neto, dc.iva, dc.impuestoharina , dc.impuestocarne , dc.impuestoilarefrescos , dc.impuestoilalicores , dc.impuestoilavinos , dc.total, dc.fecha ,dc.descuento,dc.foliosii,dc.donacion ,dc.numero,mc.nombre,mc.direccion "
    csql.sql = csql.sql & "from " + clientesistema + "ventas" + loc + ".sv_documento_cabeza_" + loc + " AS dc, " + clientesistema + "ventas" + loc + ".sv_documento_detalle_" + loc + " AS dd, " + clientesistema + "ventas.sv_maestroclientes as mc "
    csql.sql = csql.sql & "WHERE dc.caja='" + caja + "' and dc.fecha='" + Format(fecha, "yyyy-mm-dd") + "' and dc.local = '" & loc & "' AND dc.local = dd.local AND dd.tipo = '" + tipo + "' AND dc.foliosii = '" & numero & "' AND dd.caja=dc.caja and dd.tipo = dc.tipo AND dd.numero = dc.numero and dd.fecha=dc.fecha and dc.rut=mc.rut and mc.sucursal='0' ORDER BY dd.linea ASC "
    csql.Execute
    Informe.Rows = 1
        If csql.RowsAffected > 0 Then
        frmgastos.Enabled = True
        Command2.Enabled = True
        Command5.Enabled = True
        
        Set resultados = csql.OpenResultset
        LBLRUT.Caption = Mid(resultados(6), 1, 9) + "-" + Mid(resultados(6), 10, 1)
        lblNOMBRE.Caption = resultados(21)
        lbldireccion.Caption = resultados(22)
        codigoempresa = leerdatos(conta, "maestroempresas", "codigoempresa", "mid(rut,1,8)='" + Mid(LBLRUT.Caption, 2, 8) + "' ")
        LBLEMPRESA.Caption = leerempresa(codigoempresa)
        
        dato0.text = codigoempresa
        lblneto.Caption = Format(resultados("neto"), "###,###,##0")
        lbliva.Caption = Format(resultados("iva"), "###,###,##0")
        lbltotal.Caption = Format(resultados(15), "###,###,##0")
        lbliha.Caption = Format(resultados(10), "###,###,##0")
        lblica.Caption = Format(resultados(11), "###,###,##0")
        lblrefrescos.Caption = Format(resultados(12), "###,###,##0")
        lbllicores.Caption = Format(resultados(13), "###,###,##0")
        lblvinos.Caption = Format(resultados(14), "###,###,##0")
        
        While resultados.EOF = False
        Informe.Rows = Informe.Rows + 1
        Informe.Cell(Informe.Rows - 1, 1).text = resultados(0)
        Informe.Cell(Informe.Rows - 1, 2).text = resultados(1)
        Informe.Cell(Informe.Rows - 1, 3).text = resultados(2)
        Informe.Cell(Informe.Rows - 1, 4).text = resultados(3)
        Informe.Cell(Informe.Rows - 1, 5).text = resultados(4)
        Informe.Cell(Informe.Rows - 1, 6).text = resultados(5)
        
        resultados.MoveNext
        
        
        Wend
        
    End If
    Call leer(DT1.text, DT3.text, DT6.text + "-" + DT5.text + "-" + DT4.text, DT0.text, DT2.text)
   
    
End Sub



Private Sub Text4_Change()

End Sub

Sub grabar(FOLIO, fecha, tipo, numero, fechadocumento, loc, codigo_centro, cuentacontable, codigopresupuesto, codigocontable, NETO, iva, EXENTO, ilalicores, ilavinos, ilarefrescos, IHA, ICA, total, caja, rut, interno, autorizado)
    campos(0, 0) = "folio"
    campos(1, 0) = "fecha"
    campos(2, 0) = "tipo"
    campos(3, 0) = "numero"
    campos(4, 0) = "fechadocumento"
    campos(5, 0) = "local"
    campos(6, 0) = "codigo_centro"
    campos(7, 0) = "cuentacontable"
    campos(8, 0) = "codigopresupuesto"
    campos(9, 0) = "codigocontable"
    campos(10, 0) = "neto"
    campos(11, 0) = "iva"
    campos(12, 0) = "exento"
    campos(13, 0) = "ilalicores"
    campos(14, 0) = "ilavinos"
    campos(15, 0) = "ilarefrescos"
    campos(16, 0) = "iha"
    campos(17, 0) = "ica"
    campos(18, 0) = "total"
    campos(19, 0) = "caja"
    campos(20, 0) = "rut"
    campos(21, 0) = "interno"
    campos(22, 0) = "autorizado"
    campos(23, 0) = "solicitado"
    campos(24, 0) = "glosa"
    campos(25, 0) = "crcc"
    campos(26, 0) = ""
    campos(0, 1) = FOLIO
    campos(1, 1) = Format(fecha, "yyyy-mm-dd")
    campos(2, 1) = tipo
    campos(3, 1) = numero
    campos(4, 1) = Format(fechadocumento, "yyyy-mm-dd")
    campos(5, 1) = loc
    campos(6, 1) = codigo_centro
    campos(7, 1) = cuentacontable
    campos(8, 1) = codigopresupuesto
    campos(9, 1) = codigocontable
    campos(10, 1) = Replace(NETO, ".", "")
    campos(11, 1) = Replace(iva, ".", "")
    campos(12, 1) = Replace(EXENTO, ".", "")
    campos(13, 1) = Replace(ilalicores, ".", "")
    campos(14, 1) = Replace(ilavinos, ".", "")
    campos(15, 1) = Replace(ilarefrescos, ".", "")
    campos(16, 1) = Replace(IHA, ".", "")
    campos(17, 1) = Replace(ICA, ".", "")
    campos(18, 1) = Replace(total, ".", "")
    campos(19, 1) = caja
    campos(20, 1) = Mid(rut, 1, 9) + Mid(rut, 11, 1)
    campos(21, 1) = interno
    campos(22, 1) = autorizado
    campos(23, 1) = DATO5.text
    campos(24, 1) = dato6.text
    campos(25, 1) = dato7.text
    campos(0, 2) = "vales_credito"
    op = 2
    sqlconta.response = campos
    Set sqlconta.conexion = conta
    Call sqlconta.sqlconta(op, condicion)
    
    End Sub

Private Sub Text1_Change()

End Sub

Public Function LEERULTIMOFOLIO() As String

    Dim resultados As rdoResultset
    Dim csql As New rdoQuery
    
        Set csql.ActiveConnection = conta
            csql.sql = "select IFNULL(max(folio),0) from vales_credito "
            
            csql.Execute
    If csql.RowsAffected > 0 Then
    Set resultados = csql.OpenResultset
    
    
        LEERULTIMOFOLIO = Format(resultados(0) + 1, "0000000000")
    End If
    
End Function

Private Sub lblotros_Click()

End Sub

Public Sub IMPRIMEcredito()
    Dim k As Integer
    Dim NUMFIC As Integer
    Dim valecredito As String
    On Error GoTo no:
    
      NUMFIC = 20
    ''''''''''''''''''
    Open "COM1:4800,N,8,1,CD0,CS0,DS0,OP0,RS,TB100,RB100" For Output As #NUMFIC
    
    ''''''''''''''''''

    '''''''''''''''''''''
    'EMPAQUE
    '''''''''''''''''''''
    
    For k = 1 To 2
    Print #NUMFIC, Chr$(27); Chr$(64) '
    Print #NUMFIC, leerdatoslocal(DT0.text, "nombre")
    Print #NUMFIC, leerdatoslocal(DT0.text, "rut")
    Print #NUMFIC, ""
    Print #NUMFIC, "          VALE DE CREDITO         "
    Print #NUMFIC, "          ================         "
    Print #NUMFIC,
    Print #NUMFIC, "TIPO    :CONSUMO INTERNO "
    Print #NUMFIC, "FOLIO   :"; FOLIO.text
    Print #NUMFIC, "FECHA   :"; Format(fecha.text, "dd-mm-yyyy")
    Print #NUMFIC, "=========================================="
    Print #NUMFIC, "TIPO    :"; DT1.text
    Print #NUMFIC, "NUMERO  :"; DT3.text
    Print #NUMFIC, "FECHA   :"; DT4.text + "-" + DT5.text + "-" + DT6.text
    Print #NUMFIC, "CAJA    :"; DT2.text
    Print #NUMFIC, "CLIENTE :"; LBLRUT.Caption
    Print #NUMFIC, "NOMBRE  :"; lblNOMBRE.Caption
    Print #NUMFIC, "DIREC.  :"; lbldireccion.Caption
    Print #NUMFIC, "=========================================="
    Print #NUMFIC,
    Print #NUMFIC, "MONTO CREDITO :"; Format(lbltotal.Caption, " $ ###,###,###")
    Print #NUMFIC,
    
    Print #NUMFIC, "CONTABLE:"; dato0.text
    Print #NUMFIC, Mid(LBLEMPRESA.Caption, 1, 42)
    
    Print #NUMFIC, "MAYOR   :"; dato1.text
    Print #NUMFIC, Mid(lblmayor.Caption, 1, 42)
    
    Print #NUMFIC, "CENTRO  :"; dato2.text
    Print #NUMFIC, Mid(lblcentro.Caption, 1, 42)
    
    Print #NUMFIC, "GASTO   :"; dato3.text
    Print #NUMFIC, Mid(lblgasto.Caption, 1, 42)
    Print #NUMFIC,
    Print #NUMFIC,
    Print #NUMFIC,
    
    Print #NUMFIC,
    Print #NUMFIC, "              __________________             "
    Print #NUMFIC, "              FIRMA AUTORIZADORA             "
    Print #NUMFIC,
    Print #NUMFIC,
    
    Print #NUMFIC,
    Print #NUMFIC,
    Print #NUMFIC,
    Print #NUMFIC,
    
    Print #NUMFIC, Chr(27); "i"
    Next k
    Close #NUMFIC
    Exit Sub
    
no:
    MsgBox "IMPRESORA NO ESTA DISPONIBLE "

End Sub
Public Sub IMPRIMEcreditoDIRECTO()
    Dim k As Integer
    Dim NUMFIC As Integer
    Dim valecredito As String
    On Error GoTo no:
      NUMFIC = 20
    ''''''''''''''''''
    Open "COM1:4800,N,8,1,CD0,CS0,DS0,OP0,RS,TB100,RB100" For Output As #NUMFIC
    
    ''''''''''''''''''

    '''''''''''''''''''''
    'EMPAQUE
    '''''''''''''''''''''
    
    For k = 1 To 2
    Print #NUMFIC, Chr$(27); Chr$(64) '
    Print #NUMFIC, leerdatoslocal(DT0.text, "nombre")
    Print #NUMFIC, leerdatoslocal(DT0.text, "rut")
    Print #NUMFIC, ""
    Print #NUMFIC, "          VALE DE CREDITO         "
    Print #NUMFIC, "          ================         "
    Print #NUMFIC,
    Print #NUMFIC, "TIPO    :VALE CREDITO "
    Print #NUMFIC, "FOLIO   :"; FOLIO.text
    Print #NUMFIC, "FECHA   :"; Format(fecha.text, "dd-mm-yyyy")
    Print #NUMFIC, "=========================================="
    Print #NUMFIC, "TIPO    :"; DT1.text
    Print #NUMFIC, "NUMERO  :"; DT3.text
    Print #NUMFIC, "CAJA    :"; DT2.text
    Print #NUMFIC, "FECHA   :"; DT4.text + "-" + DT5.text + "-" + DT6.text
    Print #NUMFIC, "CLIENTE :"; LBLRUT.Caption
    Print #NUMFIC, "NOMBRE  :"; lblNOMBRE.Caption
    Print #NUMFIC, "DIREC.  :"; lbldireccion.Caption
    Print #NUMFIC, "=========================================="
    Print #NUMFIC,
    Print #NUMFIC, "MONTO CREDITO :"; Format(lbltotal.Caption, " $ ###,###,###")
    Print #NUMFIC,
    Print #NUMFIC,
    Print #NUMFIC,
    Print #NUMFIC,
    
    Print #NUMFIC,
    Print #NUMFIC, "              __________________             "
    Print #NUMFIC, "              FIRMA ACEPTA CARGO             "
    Print #NUMFIC,
    Print #NUMFIC,
    Print #NUMFIC,
    Print #NUMFIC, "              __________________             "
    Print #NUMFIC, "              FIRMA AUTORIZADORA             "
    Print #NUMFIC,
    Print #NUMFIC,
    
    Print #NUMFIC,
    Print #NUMFIC,
    Print #NUMFIC,
    Print #NUMFIC,
    
    Print #NUMFIC, Chr(27); "i"
    Next k
    Close #NUMFIC
    Exit Sub
no:
    MsgBox "IMPRESORA NO ESTA DISPONIBLE "
    

End Sub

Sub leer(tipo, numero, fechadocumento, loc, caja)
    campos(0, 0) = "folio"
    campos(1, 0) = "fecha"
    campos(2, 0) = "tipo"
    campos(3, 0) = "numero"
    campos(4, 0) = "fechadocumento"
    campos(5, 0) = "local"
    campos(6, 0) = "codigo_centro"
    campos(7, 0) = "cuentacontable"
    campos(8, 0) = "codigopresupuesto"
    campos(9, 0) = "codigocontable"
    campos(10, 0) = "autorizado"
    campos(11, 0) = "solicitado"
    campos(12, 0) = "glosa"
    campos(13, 0) = "interno"
    campos(14, 0) = "crcc"
    campos(15, 0) = ""
    condicion = "tipo='" + tipo + "' and numero='" + numero + "' and fechadocumento='" + Format(fechadocumento, "yyyy-mm-dd") + "' and local='" + loc + "' and caja='" + caja + "' "
    campos(0, 2) = "vales_credito"
    op = 5
    sqlconta.response = campos
    Set sqlconta.conexion = conta
    Call sqlconta.sqlconta(op, condicion)
    If sqlconta.status = 0 Then
    dato1.text = sqlconta.response(7, 3)
    dato2.text = sqlconta.response(6, 3)
    dato3.text = sqlconta.response(8, 3)
    FOLIO.text = sqlconta.response(0, 3)
    dato4.text = sqlconta.response(10, 3)
    DATO5.text = sqlconta.response(11, 3)
    dato6.text = sqlconta.response(12, 3)
    dato7.text = sqlconta.response(14, 3)
    
    If sqlconta.response(13, 3) = "S" Then
    SSTab1.Tab = 0
    Else
    SSTab1.Tab = 1
    
    End If
    
    fecha.text = Format(sqlconta.response(1, 3), "dd-mm-yyyy")
    lblmayor.Caption = leerNombreMayor(dato1)
    lblcentro.Caption = leerNOMBREgastos(dato2)
    lblgasto.Caption = leernombreanalisis(dato1, dato3)
    lblcrcc.Caption = leerNOMBREcrcc2(dato7, dato0.text)
    
    End If
    
    End Sub
Sub leerfolio(numero)
    campos(0, 0) = "folio"
    campos(1, 0) = "fecha"
    campos(2, 0) = "tipo"
    campos(3, 0) = "numero"
    campos(4, 0) = "fechadocumento"
    campos(5, 0) = "local"
    campos(6, 0) = "caja"
    campos(7, 0) = ""
    condicion = "folio='" + FOLIO + "' "
    campos(0, 2) = "vales_credito"
    op = 5
    sqlconta.response = campos
    Set sqlconta.conexion = conta
    Call sqlconta.sqlconta(op, condicion)
    If sqlconta.status = 0 Then
    DT0.text = sqlconta.response(5, 3)
    Call DT0_KeyPress(13)
    
    DT1.text = sqlconta.response(2, 3)
    DT2.text = sqlconta.response(6, 3)
    DT3.text = sqlconta.response(3, 3)
    DT4.text = Format(sqlconta.response(4, 3), "dd")
    DT5.text = Format(sqlconta.response(4, 3), "mm")
    DT6.text = Format(sqlconta.response(4, 3), "yyyy")
    Call Command1_Click
    
    
    
    End If
    
    End Sub


Sub grabafactura(tipo, numero, fecha, fechavencimiento, rut, NETO, iva, EXENTO, retencion, total, AÑOCONTABLE, MESCONTABLE, FOLIO, ESSUPER)
    Dim netos As Double
    Dim DH As String
    Dim DH2 As String
    Dim mesconta As String
    Dim añoconta As String
    Dim diaconta As String
    
    Dim exentos As Double
    Dim TIPOCON As String
    Dim CRCC As String
    Dim ELECTRONICA As String
    Dim tipodoc As String
  
    Dim fechacom As String
    
    rut = leerdatoslocal(DT0.text, "rut")
    
    
    
    campos(0, 0) = "tipo"
    campos(1, 0) = "numero"
    campos(2, 0) = "fecha"
    campos(3, 0) = "fechavencimiento"
    campos(4, 0) = "rut"
    campos(5, 0) = "neto"
    campos(6, 0) = "iva"
    campos(7, 0) = "exento"
    campos(8, 0) = "retencion"
    campos(9, 0) = "total"
    campos(10, 0) = "añocontable"
    campos(11, 0) = "mescontable"
    campos(12, 0) = "comentario"
    campos(13, 0) = "electronica"
    campos(14, 0) = "activo"
    campos(15, 0) = "fechadigitacion"
    campos(16, 0) = "folio"
    campos(17, 0) = "impuestoespecifico"
    campos(18, 0) = "comprasuper"
    campos(19, 0) = ""
 
    
    TIPOCON = tipo
    ELECTRONICA = "S"
    tipodoc = "FC"
    DH = "H"
    DH2 = "D"
    
    campos(0, 1) = TIPOCON
    campos(1, 1) = numero
    campos(2, 1) = Format(fecha, "yyyy-mm-dd")
    campos(3, 1) = Format(fechavencimiento, "yyyy-mm-dd")
    campos(4, 1) = rut
    campos(5, 1) = Replace(NETO, ".", "")
    campos(6, 1) = Replace(iva, ".", "")
    campos(7, 1) = Replace(EXENTO, ".", "")
    campos(8, 1) = "0"
    campos(9, 1) = Replace(total, ".", "")
    
    
    campos(10, 1) = AÑOCONTABLE
    campos(11, 1) = Format(MESCONTABLE, "00")
    campos(12, 1) = "CENTRALIZACION AUTOMATICA DE GASTOS"
        
    campos(13, 1) = ELECTRONICA
    campos(14, 1) = "N"
    campos(15, 1) = Format(fechasistema, "yyyy-mm-dd")
    
    campos(16, 1) = LEERULTIMOFOLIOCOMPRAS(campos(11, 1), campos(10, 1), dato0.text)
    campos(17, 1) = "0"
    campos(18, 1) = ESSUPER
    
    condicion = ""
    campos(0, 2) = clientesistema + "conta" + dato0.text + ".facturasdecompras"
    op = 2
    sqlconta.response = campos
    Set sqlconta.conexion = contadb

    Call sqlconta.sqlconta(op, condicion)
    k = sqlconta.status
    
    fecha = Format(campos(3, 1), "yyyy-mm-dd")
    fechacom = Format(fechasistema, "yyyy-mm") + "-" + "01"
    If fecha >= fechacom Then
    fechacom = fecha
    End If
    
    
    
    Call grabarcomprobante_lineas(tipodoc, campos(1, 1), "001", fechacom, CUENTAPROVEEDOR, "", campos(4, 1), "", "CENTRALIZA DOCUMENTO DE COMPRAS " + Grid1.Cell(LINEA, 1).text, tipodoc, campos(1, 1), campos(2, 1), campos(3, 1), campos(9, 1), DH, USUARIOSISTEMA, campos(11, 1), campos(10, 1), Format(fechasistema, "yyyy-mm-dd"), Time, campos(4, 1), "", "")
    Call grabarcomprobante_lineas(tipodoc, campos(1, 1), "002", fechacom, ivacredito, "", "", "", "CENTRALIZACION I.V.A", tipodoc, campos(1, 1), campos(2, 1), campos(3, 1), campos(6, 1), DH2, USUARIOSISTEMA, campos(11, 1), campos(10, 1), Format(fechasistema, "yyyy-mm-dd"), Time, campos(4, 1), "", "")
    Call grabardetallefactura(tipo, numero, rut)
    


End Sub

Sub grabardetallefactura(tipo, numero, rut)
    
    Dim TIPOCON As String
    Dim tipo2 As String
    Dim j As Integer
    Dim lin As Integer
    Dim ilas As Double
    Dim CRCC As String
    Dim cuenta As String
    Dim DH As String
    Dim NOMBRE As String
    Dim localfiltro As String
    
    Dim tipodoc As String
    
    
    campos(0, 0) = "tipo"
    campos(1, 0) = "numero"
    campos(2, 0) = "linea"
    campos(3, 0) = "rut"
    campos(4, 0) = "cuentadelmayor"
    campos(5, 0) = "glosa"
    campos(6, 0) = "monto"
    campos(7, 0) = "dh"
    campos(8, 0) = "centrodecosto"
    campos(9, 0) = "rutctacte"
    campos(10, 0) = "fechacreacion"
    campos(11, 0) = "cuenta_presupuesto"
    campos(12, 0) = "centro_gastos"
    campos(13, 0) = ""
    
    TIPOCON = tipo:
    tipodoc = "FC": DH = "D"
    
    cuenta = dato1.text
    NOMBRE = lblmayor.Caption
    
    
    
    
Rem CALCULA NETOS

    lin = 3
    campos(0, 1) = TIPOCON
    campos(1, 1) = numero
    campos(2, 1) = Format(lin, "000")
    campos(3, 1) = rut
    campos(4, 1) = cuenta
    campos(5, 1) = dato6.text
    campos(6, 1) = Replace(lblneto.Caption, ".", "")
    campos(7, 1) = DH
    campos(8, 1) = dato7.text
    campos(9, 1) = ""
    campos(10, 1) = Format(fechasistema, "yyyy-mm-dd")
    campos(11, 1) = dato3.text
    campos(12, 1) = dato2.text
    
    campos(0, 2) = clientesistema + "conta" + dato0.text + ".facturasdecompras_detalle"
    condicion = ""
    op% = 2
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)
    
    Call grabarcomprobante_lineas(tipodoc, campos(1, 1), campos(2, 1), fecha, campos(4, 1), "", campos(3, 1), campos(8, 1), campos(5, 1), tipodoc, campos(1, 1), fecha, fecha, campos(6, 1), DH, USUARIOSISTEMA, MES, año, Format(fechasistema, "yyyy-mm-dd"), Time, campos(3, 1), campos(12, 1), campos(11, 1))
    Rem CALCULA ILAS refrescos

    ilas = CDbl(lblrefrescos.Caption)
    If ilas <> 0 Then
    lin = lin + 1
    campos(0, 1) = TIPOCON
    campos(1, 1) = numero
    campos(2, 1) = Format(lin, "000")
    campos(3, 1) = rut
    campos(4, 1) = leerdatoslocal(localfiltro, "cuentailarefrescos")
    If campos(4, 1) = "" Then campos(4, 1) = cuenta
    campos(5, 1) = " IMPUESTO ILA REFRESCOS"
    campos(6, 1) = ilas
    campos(7, 1) = DH
    campos(8, 1) = ""
    campos(9, 1) = ""
    campos(10, 1) = Format(fechasistema, "yyyy-mm-dd")
    campos(11, 1) = ""
    campos(12, 1) = ""
    campos(0, 2) = clientesistema + "conta" + dato0.text + ".facturasdecompras_detalle"
    condicion = ""
    op% = 2
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)
    Call grabarcomprobante_lineas(tipodoc, campos(1, 1), campos(2, 1), fecha, campos(4, 1), "", "", "", campos(5, 1), tipodoc, campos(1, 1), fecha, fecha, campos(6, 1), DH, USUARIOSISTEMA, MES, año, Format(fechasistema, "yyyy-mm-dd"), Time, campos(3, 1), "", "")
    End If
Rem CALCULA ILAS vinos
    ilas = CDbl(lblvinos.Caption)
    If ilas <> 0 Then
    lin = lin + 1
    campos(0, 1) = TIPOCON
    campos(1, 1) = numero
    campos(2, 1) = Format(lin, "000")
    campos(3, 1) = rut
    campos(4, 1) = leerdatoslocal(localfiltro, "cuentailavinos")
    If campos(4, 1) = "" Then campos(4, 1) = cuenta
    campos(5, 1) = " IMPUESTO ILA VINOS"
    campos(6, 1) = ilas
    campos(7, 1) = DH
    campos(8, 1) = ""
    campos(9, 1) = ""
    campos(10, 1) = Format(fechasistema, "yyyy-mm-dd")
    campos(11, 1) = ""
    campos(12, 1) = ""
    campos(0, 2) = clientesistema + "conta" + dato0.text + ".facturasdecompras_detalle"
    condicion = ""
    op% = 2
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)
    Call grabarcomprobante_lineas(tipodoc, campos(1, 1), campos(2, 1), fecha, campos(4, 1), "", "", "", campos(5, 1), tipodoc, campos(1, 1), fecha, fecha, campos(6, 1), DH, USUARIOSISTEMA, MES, año, Format(fechasistema, "yyyy-mm-dd"), Time, campos(3, 1), "", "")
    End If
Rem CALCULA ILAS vinos
    ilas = CDbl(lbllicores.Caption)
    If ilas <> 0 Then
    lin = lin + 1
    

    campos(0, 1) = TIPOCON
    campos(1, 1) = numero
    campos(2, 1) = Format(lin, "000")
    campos(3, 1) = rut
    campos(4, 1) = leerdatoslocal(localfiltro, "cuentailalicores")
    If campos(4, 1) = "" Then campos(4, 1) = cuenta
    campos(5, 1) = " IMPUESTO ILA LICORES "
    campos(6, 1) = ilas
    campos(7, 1) = DH
    campos(8, 1) = ""
    campos(9, 1) = ""
    campos(10, 1) = Format(fechasistema, "yyyy-mm-dd")
    campos(11, 1) = ""
    campos(12, 1) = ""
    
    campos(0, 2) = clientesistema + "conta" + dato0.text + ".facturasdecompras_detalle"
    condicion = ""
    op% = 2
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)
    Call grabarcomprobante_lineas(tipodoc, campos(1, 1), campos(2, 1), fecha, campos(4, 1), "", "", "", campos(5, 1), tipodoc, campos(1, 1), fecha, fecha, campos(6, 1), DH, USUARIOSISTEMA, MES, año, Format(fechasistema, "yyyy-mm-dd"), Time, campos(3, 1), "", "")
    
    End If

Rem CALCULA HARINA
    ilas = CDbl(lbliha.Caption)
    If ilas <> 0 Then
    lin = lin + 1

    campos(0, 1) = TIPOCON
    campos(1, 1) = numero
    campos(2, 1) = Format(lin, "000")
    campos(3, 1) = rut
    campos(4, 1) = leerdatoslocal(localfiltro, "cuentaharina")
    If campos(4, 1) = "" Then campos(4, 1) = cuenta
    campos(5, 1) = " IMPUESTO HARINAS"
        campos(6, 1) = ilas
    campos(7, 1) = DH
    campos(8, 1) = ""
    campos(9, 1) = campos(3, 1)
    campos(10, 1) = Format(fechasistema, "yyyy-mm-dd")
    
    campos(0, 2) = clientesistema + "conta" + dato0.text + ".facturasdecompras_detalle"
    condicion = ""
    op% = 2
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)
    Call grabarcomprobante_lineas(tipodoc, campos(1, 1), campos(2, 1), fecha, campos(4, 1), "", "", "", campos(5, 1), tipodoc, campos(1, 1), fecha, fecha, campos(6, 1), DH, USUARIOSISTEMA, MES, año, Format(fechasistema, "yyyy-mm-dd"), Time, campos(3, 1), "", "")
    
    End If

Rem CALCULA carne
    ilas = CDbl(lblica.Caption)
    If ilas <> 0 Then
    lin = lin + 1
       campos(0, 1) = TIPOCON
    campos(1, 1) = numero
    campos(2, 1) = Format(lin, "000")
    campos(3, 1) = rut
    campos(4, 1) = leerdatoslocal(localfiltro, "cuentacarne")
    If campos(4, 1) = "" Then campos(4, 1) = cuenta
    campos(5, 1) = " IMPUESTO CARNE"
    campos(6, 1) = ilas
    campos(7, 1) = DH
    campos(8, 1) = ""
    campos(9, 1) = ""
    campos(10, 1) = Format(fechasistema, "yyyy-mm-dd")
    
    
    campos(0, 2) = clientesistema + "conta" + dato0.text + ".facturasdecompras_detalle"
    condicion = ""
    op% = 2
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)
    Call grabarcomprobante_lineas(tipodoc, campos(1, 1), campos(2, 1), fecha, campos(4, 1), "", "", "", campos(5, 1), tipodoc, campos(1, 1), fecha, fecha, campos(6, 1), DH, USUARIOSISTEMA, MES, año, Format(fechasistema, "yyyy-mm-dd"), Time, campos(3, 1), "", "")
    
    End If
    
   
    
    
End Sub

Public Function LEERULTIMOFOLIOCOMPRAS(mesconta, añoconta, empresa) As String

    Dim resultados As rdoResultset
    Dim csql As New rdoQuery
    
        Set csql.ActiveConnection = contadb

            csql.sql = "select max(folio) from " + clientesistema + "conta" + empresa + ".facturasdecompras where mescontable = '" & Format(mesconta, "00") & "' AND añocontable = '" & añoconta & "' "
            
            csql.Execute
    If csql.RowsAffected > 0 Then
    Set resultados = csql.OpenResultset
    
        If resultados(0) <> "NULO" Then
        LEERULTIMOFOLIOCOMPRAS = resultados(0) + 1
        Else
        LEERULTIMOFOLIOCOMPRAS = "0000000001"
        End If
        
    End If
    
End Function

Sub grabarcomprobante_lineas(tipo, numero, LINEA, fecha, codigocuenta, tipoctacte, rutctacte, centrocosto, glosacontable, tipodocumento, numerodocumento, fechadocumento, fechavencimiento, monto, DH, creadopor, MES, año, fechacreacion, horacreacion, rutproveedor, centro, detalle)
    Dim condicion As String
    Dim campos(40, 3) As String
    Dim op As Integer
    Dim TIPOCON As String
    Dim tipo2 As String
    Dim j As Integer
    Dim lin As String
    Dim lar As Integer
    campos(0, 0) = "tipo"
    campos(1, 0) = "numero"
    campos(2, 0) = "linea"
    campos(3, 0) = "fecha"
    campos(4, 0) = "codigocuenta"
    campos(5, 0) = "tipoctacte"
    campos(6, 0) = "rutctacte"
    campos(7, 0) = "centrocosto"
    campos(8, 0) = "glosacontable"
    campos(9, 0) = "tipodocumento"
    campos(10, 0) = "numerodocumento"
    campos(11, 0) = "fechadocumento"
    campos(12, 0) = "fechavencimiento"
    campos(13, 0) = "monto"
    campos(14, 0) = "dh"
    campos(15, 0) = "creadopor"
    campos(16, 0) = "mes"
    campos(17, 0) = "año"
    campos(18, 0) = "fechacreacion"
    campos(19, 0) = "horacreacion"
    campos(20, 0) = "rutproveedor"
    campos(21, 0) = "cuenta_presupuesto"
    campos(22, 0) = "centro_gastos"
    campos(23, 0) = ""
    
    campos(0, 1) = tipo
    campos(1, 1) = numero
    campos(2, 1) = LINEA
    campos(3, 1) = Format(fecha, "yyyy-mm-dd")
    campos(4, 1) = codigocuenta
    campos(5, 1) = tipoctacte
    campos(6, 1) = rutctacte
    campos(7, 1) = centrocosto
    campos(8, 1) = glosacontable
    campos(9, 1) = tipodocumento
    campos(10, 1) = numerodocumento
    campos(11, 1) = Format(fechadocumento, "yyyy-mm-dd")
    campos(12, 1) = Format(fechavencimiento, "yyyy-mm-dd")
    campos(13, 1) = monto

    campos(14, 1) = DH
    campos(15, 1) = creadopor
    campos(16, 1) = Format(MES, "00")
    campos(17, 1) = año
    
    campos(18, 1) = Format(fechacreacion, "yyyy-mm-dd")
    campos(19, 1) = horacreacion
    campos(20, 1) = rutproveedor
    campos(21, 1) = detalle
    campos(22, 1) = centro
If tipo = "CD" Then
    campos(0, 2) = clientesistema + "conta" + leerdatoslocal(DT0.text, "codigocontable") + ".movimientoscontables"
   Else
    campos(0, 2) = clientesistema + "conta" + dato0.text + ".movimientoscontables"
   
   End If
   
    op = 2
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    
    Call sqlconta.sqlconta(op, condicion)
   'Call ACTUALIZADOCUMENTO("+")
   
End Sub


Sub ELIMINAR(tipo, numero, rut)
    Dim TIPOCON As String
    
    Call eliminavale(DT1.text, DT3.text, DT6.text + "-" + DT5.text + "-" + DT4.text, DT0.text, DT2.text)
Rem elimina 1

    tipo = "4"
    campos(0, 2) = clientesistema + "conta" & dato0.text + ".facturasdecompras_impuestos"
    condicion = "tipo='" + tipo + "' and numero=" + "'" + numero + "'" + " and rut=" + "'" + rut + "'"
    op = 4
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)
    
Rem elimina 2

    campos(0, 2) = clientesistema + "conta" & dato0.text + ".facturasdecompras"
    condicion = "tipo='" + tipo + "' and numero=" + "'" + numero + "'" + " and rut=" + "'" + rut + "'"
    op = 4
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)

Rem elimina 3

    campos(0, 2) = clientesistema + "conta" & dato0.text + ".movimientoscontables"
    If tipo = "1" Then TIPOCON = "FC"
    If tipo = "4" Then TIPOCON = "FC"
    
    condicion = "tipo=" + "'" + TIPOCON + "'" + " and numero=" + "'" + numero + "' and rutproveedor='" + rut + "'"
    
    op = 4
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)
Rem ELIMINA CARGO
  
    campos(0, 2) = clientesistema + "conta" + leerdatoslocal(DT0.text, "codigocontable") + ".movimientoscontables"
    If tipo = "1" Then TIPOCON = "CD"
    If tipo = "4" Then TIPOCON = "CD"
    
    
    condicion = "tipo=" + "'" + TIPOCON + "'" + " and numero=" + "'" + FOLIO.text + "' "
    
    op = 4
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)
  
  
Rem elimina 4
    campos(0, 2) = clientesistema + "conta" & dato0.text + ".facturasdecompras_detalle"
    condicion = "tipo='" + tipo + "' and numero=" + "'" + numero + "'" + " and rut=" + "'" + rut + "'"
    
    op = 4
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)
  '  If sqlconta.status = 4 Then Stop

End Sub

Sub eliminavale(tipo, numero, fechadocumento, loc, caja)
    campos(0, 0) = "folio"
    campos(1, 0) = "fecha"
    campos(2, 0) = "tipo"
    campos(3, 0) = "numero"
    campos(4, 0) = "fechadocumento"
    campos(5, 0) = "local"
    campos(6, 0) = "codigo_centro"
    campos(7, 0) = "cuentacontable"
    campos(8, 0) = "codigopresupuesto"
    campos(9, 0) = "codigocontable"
    campos(10, 0) = ""
    condicion = "tipo='" + tipo + "' and numero='" + numero + "' and fechadocumento='" + Format(fechadocumento, "yyyy-mm-dd") + "' and local='" + loc + "' and caja='" + caja + "' "
    campos(0, 2) = "vales_credito"
    op = 4
    sqlconta.response = campos
    Set sqlconta.conexion = conta
    Call sqlconta.sqlconta(op, condicion)
    
    End Sub

Sub retorno()
frmgastos.Enabled = False
DT0.text = ""
DT1.text = "FV"
DT2.text = ""
DT3.text = ""
DT4.text = ""
DT5.text = ""
DT6.text = ""
LBLRUT.Caption = ""
lblNOMBRE.Caption = ""
lbldireccion.Caption = ""
lblgiro.Caption = ""
lblciudad.Caption = ""
Informe.Rows = 1
lblneto.Caption = ""
lbliva.Caption = ""
lblexento.Caption = ""
lbltotal.Caption = ""
lblrefrescos.Caption = ""
lbllicores.Caption = ""
lblvinos.Caption = ""
lblica.Caption = ""
lbliha.Caption = ""
dato1.text = ""
dato0.text = ""
dato2.text = ""
dato3.text = ""
dato4.text = ""
lblmayor.Caption = ""
lblcentro.Caption = ""
lblgasto.Caption = ""
LBLEMPRESA.Caption = ""
DT0.SetFocus
dato6.text = ""
DATO5.text = ""
SSTab1.Tab = 0
dato7.text = ""
lblcrcc.Caption = ""

End Sub

Sub ayudalocales(ByRef caja As TextBox)
    Dim campos As Variant
    Dim cfijo As Variant
    Dim largo As Variant
    campos = Array("codigo", "nombre")
    largo = Array("10s", "40s")
    cfijo = "no"
    cabezas = Array("Codigo", "Nombre")
    mensajeAyuda = "Ayuda Locales"
        
    Call cargaAyudaT(Servidor, basebus, Usuario, password, clientesistema & "gestion" & ".g_maestroempresas", caja, campos, cfijo, largo, 2)
    caja.Enabled = True
    caja.SetFocus
End Sub

Public Function leerNombreMayor(ByVal codigo As String) As String
    Dim condicion As String
    Dim op As Integer
    campos(0, 0) = "nombre"
    campos(1, 0) = ""
    
    campos(0, 2) = clientesistema + "conta" + dato0.text + ".cuentasdelmayor"
    
    condicion = "codigo = '" & codigo & "' and año='" + Format(fechasistema, "yyyy") + "'  "
    
    
    op = 5
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)
    
    If sqlconta.status = 0 Then
        leerNombreMayor = sqlconta.response(0, 3)
    Else
        leerNombreMayor = ""
    End If
End Function

Private Sub botonmisaccesos_Click()
programafiltro = Me.Caption
misaccesos.Show
End Sub


Private Sub botonmisfavoritos_Click()
Call AgregaFavoritos(USUARIOSISTEMA, App.EXEName, Me.Name, Me.Caption)
End Sub
Private Function leetipofactura(loc, tipo, numero, caja, fecha)
    Dim codigoempresa As String
    
    Dim resultados As rdoResultset
    Dim csql As New rdoQuery
    Set csql.ActiveConnection = contadb
    
    csql.sql = "SELECT dc.contabilizado "
    csql.sql = csql.sql & "from " + clientesistema + "ventas" + loc + ".sv_documento_cabeza_" + loc + " AS dc "
    csql.sql = csql.sql & "WHERE dc.caja='" + caja + "' and dc.fecha='" + Format(fecha, "yyyy-mm-dd") + "' and dc.local = '" & loc & "' anD dc.foliosii = '" & numero & "' "
    csql.Execute
    
    If csql.RowsAffected > 0 Then
        Set resultados = csql.OpenResultset
        If resultados(0) = "E" Then
        leetipofactura = "4"
        Else
        leetipofactura = "1"
        End If
    End If
        
    
    
    
End Function

