VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash10a.ocx"
Object = "{4777436C-EB8C-4596-98A8-EBCF98683969}#1.0#0"; "Flexcell.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{ADBBAED6-B16F-40DC-80DF-B44910CBA76C}#1.0#0"; "Frame-Xp.ocx"
Begin VB.Form MMGarantias 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Maestro de Garantias"
   ClientHeight    =   9495
   ClientLeft      =   1410
   ClientTop       =   1140
   ClientWidth     =   12585
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9495
   ScaleWidth      =   12585
   Begin XPFrame.FrameXp FRMIMPRESION 
      Height          =   1410
      Left            =   3960
      TabIndex        =   67
      Top             =   4440
      Visible         =   0   'False
      Width           =   5820
      _ExtentX        =   10266
      _ExtentY        =   2487
      BackColor       =   8454016
      Caption         =   "IMPRESION DE GUIAS"
      CaptionEstilo3D =   1
      BackColor       =   8454016
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
      Begin VB.CommandButton Command2 
         BackColor       =   &H00FF8080&
         Caption         =   "RETORNO"
         Height          =   330
         Left            =   4005
         Style           =   1  'Graphical
         TabIndex        =   71
         Top             =   495
         Width           =   1590
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FF8080&
         Caption         =   "IMPRIME"
         Height          =   330
         Left            =   4005
         Style           =   1  'Graphical
         TabIndex        =   70
         Top             =   900
         Width           =   1590
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H0080FF80&
         Caption         =   "IMPRIME GUIA LARGA"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   180
         TabIndex        =   69
         Top             =   855
         Width           =   3480
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H0080FF80&
         Caption         =   "IMPRIME GUIA CORTA"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Left            =   180
         TabIndex        =   68
         Top             =   405
         Value           =   -1  'True
         Width           =   3480
      End
   End
   Begin VB.CheckBox guia 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Generar Guia"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1440
      TabIndex        =   64
      Top             =   9000
      Width           =   1575
   End
   Begin VB.CommandButton cmd6 
      BackColor       =   &H000080FF&
      Caption         =   "INFORME"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   54
      Top             =   8880
      Width           =   2850
   End
   Begin VB.CheckBox entregados 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Entregar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   0
      TabIndex        =   42
      Top             =   9000
      Width           =   1095
   End
   Begin XPFrame.FrameXp frmDatos 
      Height          =   5520
      Left            =   6750
      TabIndex        =   37
      Top             =   0
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   9737
      BackColor       =   12648384
      Caption         =   "Movimientos de Garantias"
      CaptionEstilo3D =   1
      BackColor       =   12648384
      ColorBarraArriba=   12648384
      ColorBarraAbajo =   32768
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
      Begin XPFrame.FrameXp frmLista 
         Height          =   4875
         Left            =   135
         TabIndex        =   38
         Top             =   450
         Width           =   5580
         _ExtentX        =   9843
         _ExtentY        =   8599
         BackColor       =   12648447
         Caption         =   "Lista de Movimientos"
         BackColor       =   12648447
         ColorBarraArriba=   12648447
         ColorBarraAbajo =   32896
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
         Begin FlexCell.Grid lista 
            Height          =   4395
            Left            =   60
            TabIndex        =   39
            Top             =   405
            Width           =   5460
            _ExtentX        =   9631
            _ExtentY        =   7752
            Cols            =   5
            DefaultFontSize =   9.75
            Rows            =   1
            SelectionMode   =   1
         End
         Begin MSAdodcLib.Adodc data 
            Height          =   330
            Left            =   60
            Top             =   4560
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
      Begin XPFrame.FrameXp frmCerrar 
         Height          =   330
         Left            =   10260
         TabIndex        =   40
         Top             =   25
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   582
         BackColor       =   49344
         Caption         =   "X"
         CaptionEstilo3D =   1
         BackColor       =   49344
         ColorBarraArriba=   32896
         ColorBarraAbajo =   12648447
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
      End
   End
   Begin XPFrame.FrameXp FrameXp4 
      Height          =   2085
      Left            =   6930
      TabIndex        =   2
      Top             =   5625
      Width           =   5460
      _ExtentX        =   9631
      _ExtentY        =   3678
      BackColor       =   8454016
      Caption         =   "MOVIMIENTOS"
      BackColor       =   8454016
      ForeColor       =   8388608
      ColorBarraAbajo =   12648384
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BordeEstilo     =   6
      Alignment       =   1
      ColorTextShadow =   8454016
      Begin VB.CommandButton cmd4 
         BackColor       =   &H000080FF&
         Caption         =   "ELIMINAR MOVIMIENTO"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   3060
         Style           =   1  'Graphical
         TabIndex        =   41
         Top             =   1620
         Visible         =   0   'False
         Width           =   2310
      End
      Begin VB.CommandButton cmd3 
         BackColor       =   &H000080FF&
         Caption         =   "CREAR MOVIMIENTOS"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   135
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   1620
         Visible         =   0   'False
         Width           =   2850
      End
      Begin VB.TextBox observa 
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
         Height          =   1095
         Left            =   135
         MaxLength       =   200
         MultiLine       =   -1  'True
         TabIndex        =   3
         Top             =   450
         Width           =   5100
      End
   End
   Begin MSAdodcLib.Adodc Clientes 
      Height          =   330
      Left            =   675
      Top             =   0
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
   Begin VB.PictureBox manual 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   7440
      ScaleHeight     =   375
      ScaleWidth      =   555
      TabIndex        =   1
      Top             =   9120
      Visible         =   0   'False
      Width           =   555
   End
   Begin XPFrame.FrameXp FrameXp3 
      Height          =   8775
      Left            =   -90
      TabIndex        =   5
      Top             =   0
      Width           =   6720
      _ExtentX        =   11853
      _ExtentY        =   15478
      BackColor       =   16744576
      Caption         =   "Datos de la Garantia"
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
      Begin FlexCell.Grid Grid1 
         Height          =   255
         Left            =   120
         TabIndex        =   63
         Top             =   3360
         Visible         =   0   'False
         Width           =   135
         _ExtentX        =   238
         _ExtentY        =   450
         Cols            =   5
         DefaultFontSize =   8.25
         Rows            =   30
      End
      Begin VB.TextBox direccioncliente 
         Height          =   285
         Left            =   360
         TabIndex        =   62
         Top             =   0
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.TextBox pivote 
         Height          =   285
         Left            =   120
         TabIndex        =   55
         Text            =   "Text1"
         Top             =   0
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.TextBox rut_cliente 
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
         TabIndex        =   22
         Top             =   750
         Width           =   1500
      End
      Begin VB.TextBox desde3 
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
         Left            =   5985
         MaxLength       =   4
         TabIndex        =   20
         Tag             =   "proveedor"
         Top             =   390
         Width           =   615
      End
      Begin VB.TextBox desde2 
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
         Left            =   5625
         MaxLength       =   2
         TabIndex        =   19
         Tag             =   "proveedor"
         Top             =   390
         Width           =   375
      End
      Begin VB.TextBox desde1 
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
         Left            =   5265
         MaxLength       =   2
         TabIndex        =   18
         Tag             =   "proveedor"
         Top             =   390
         Width           =   375
      End
      Begin VB.TextBox celular 
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
         Left            =   4050
         MaxLength       =   10
         TabIndex        =   10
         Tag             =   "proveedor"
         Top             =   1425
         Width           =   2550
      End
      Begin VB.TextBox fono 
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
         TabIndex        =   9
         Tag             =   "proveedor"
         Top             =   1425
         Width           =   1215
      End
      Begin VB.TextBox nombre_cliente 
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
         MaxLength       =   50
         TabIndex        =   8
         Tag             =   "proveedor"
         Top             =   1110
         Width           =   5160
      End
      Begin VB.TextBox folio 
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
         TabIndex        =   7
         Top             =   390
         Width           =   1500
      End
      Begin XPFrame.FrameXp FrameXp5 
         Height          =   6915
         Left            =   135
         TabIndex        =   6
         Top             =   1830
         Width           =   6495
         _ExtentX        =   11456
         _ExtentY        =   12197
         BackColor       =   8454016
         Caption         =   "Datos del Producto"
         CaptionEstilo3D =   1
         BackColor       =   8454016
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
         Begin XPFrame.FrameXp frmentregar 
            Height          =   2310
            Left            =   225
            TabIndex        =   44
            Top             =   4560
            Visible         =   0   'False
            Width           =   6090
            _ExtentX        =   10742
            _ExtentY        =   4075
            BackColor       =   8454016
            Caption         =   "ENTREGA"
            CaptionEstilo3D =   1
            BackColor       =   8454016
            ColorBarraArriba=   0
            ColorBarraAbajo =   16711680
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BordeEstilo     =   5
            Alignment       =   1
            ColorTextShadow =   0
            Begin VB.TextBox final 
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
               Height          =   600
               Left            =   1350
               MaxLength       =   200
               MultiLine       =   -1  'True
               TabIndex        =   52
               Top             =   1170
               Width           =   4605
            End
            Begin VB.TextBox nombre2 
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
               Left            =   1350
               MaxLength       =   30
               TabIndex        =   49
               Tag             =   "proveedor"
               Top             =   810
               Width           =   4620
            End
            Begin VB.TextBox hasta3 
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
               Left            =   2670
               MaxLength       =   4
               TabIndex        =   48
               Tag             =   "proveedor"
               Top             =   405
               Width           =   615
            End
            Begin VB.TextBox hasta2 
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
               Left            =   2310
               MaxLength       =   2
               TabIndex        =   47
               Tag             =   "proveedor"
               Top             =   405
               Width           =   375
            End
            Begin VB.TextBox hasta1 
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
               Left            =   1950
               MaxLength       =   2
               TabIndex        =   46
               Tag             =   "proveedor"
               Top             =   405
               Width           =   375
            End
            Begin VB.CommandButton cmd5 
               BackColor       =   &H000080FF&
               Caption         =   "GUARDAR"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   120
               Style           =   1  'Graphical
               TabIndex        =   45
               Top             =   1890
               Width           =   2850
            End
            Begin VB.Label Label1 
               Appearance      =   0  'Flat
               BackColor       =   &H00F5C9B1&
               BorderStyle     =   1  'Fixed Single
               Caption         =   " Observacion"
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
               Left            =   90
               TabIndex        =   53
               Top             =   1170
               Width           =   1215
            End
            Begin VB.Label Label27 
               Appearance      =   0  'Flat
               BackColor       =   &H00F5C9B1&
               BorderStyle     =   1  'Fixed Single
               Caption         =   " Retirado por"
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
               Left            =   90
               TabIndex        =   51
               Top             =   810
               Width           =   1215
            End
            Begin VB.Label Label9 
               Appearance      =   0  'Flat
               BackColor       =   &H00F5C9B1&
               BorderStyle     =   1  'Fixed Single
               Caption         =   " Fecha Entrega"
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
               Left            =   90
               TabIndex        =   50
               Top             =   405
               Width           =   1695
            End
         End
         Begin XPFrame.FrameXp FrameXp9 
            Height          =   1065
            Left            =   135
            TabIndex        =   35
            Top             =   2295
            Width           =   6180
            _ExtentX        =   10901
            _ExtentY        =   1879
            BackColor       =   8454016
            Caption         =   "FALLA"
            CaptionEstilo3D =   1
            BackColor       =   8454016
            ColorBarraArriba=   0
            ColorBarraAbajo =   16711680
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BordeEstilo     =   5
            Alignment       =   1
            ColorTextShadow =   0
            Begin VB.TextBox falla 
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
               Height          =   645
               Left            =   45
               MaxLength       =   200
               MultiLine       =   -1  'True
               TabIndex        =   36
               Top             =   360
               Width           =   6090
            End
         End
         Begin VB.TextBox codigobarra 
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
            MaxLength       =   13
            TabIndex        =   33
            Tag             =   "proveedor"
            Top             =   450
            Width           =   1500
         End
         Begin VB.TextBox serie 
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
            MaxLength       =   50
            TabIndex        =   31
            Tag             =   "proveedor"
            Top             =   1170
            Width           =   4890
         End
         Begin VB.TextBox articulo 
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
            MaxLength       =   30
            TabIndex        =   27
            Tag             =   "proveedor"
            Top             =   1890
            Width           =   4890
         End
         Begin VB.TextBox marca 
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
            MaxLength       =   30
            TabIndex        =   26
            Tag             =   "proveedor"
            Top             =   1530
            Width           =   4890
         End
         Begin VB.TextBox descripcion 
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
            MaxLength       =   50
            TabIndex        =   25
            Tag             =   "proveedor"
            Top             =   810
            Width           =   4890
         End
         Begin XPFrame.FrameXp frmguia 
            Height          =   1185
            Left            =   120
            TabIndex        =   56
            Top             =   3360
            Visible         =   0   'False
            Width           =   6180
            _ExtentX        =   10901
            _ExtentY        =   2090
            BackColor       =   8454016
            Caption         =   "GENERAR GUIA"
            CaptionEstilo3D =   1
            BackColor       =   8454016
            ColorBarraArriba=   0
            ColorBarraAbajo =   16711680
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BordeEstilo     =   5
            Alignment       =   1
            ColorTextShadow =   0
            Begin VB.TextBox nguia 
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
               Left            =   1080
               MaxLength       =   10
               TabIndex        =   65
               Top             =   720
               Width           =   1500
            End
            Begin VB.CommandButton cmd7 
               BackColor       =   &H000080FF&
               Caption         =   "IMPRIMIR GUIA"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   3960
               Style           =   1  'Graphical
               TabIndex        =   61
               Top             =   720
               Visible         =   0   'False
               Width           =   2130
            End
            Begin VB.TextBox ruttecnico 
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
               Left            =   1080
               MaxLength       =   9
               TabIndex        =   58
               Top             =   360
               Width           =   1500
            End
            Begin VB.Label Label3 
               BackColor       =   &H0080FF80&
               Caption         =   "Nº GUIA"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   240
               TabIndex        =   66
               Top             =   720
               Width           =   855
            End
            Begin VB.Label lblnombretecnico 
               BackColor       =   &H0080FF80&
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
               Left            =   2880
               TabIndex        =   60
               Top             =   360
               Width           =   2775
            End
            Begin VB.Label lbldvtecnico 
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
               Left            =   2340
               TabIndex        =   59
               Top             =   360
               Width           =   495
            End
            Begin VB.Label Label2 
               BackColor       =   &H0080FF80&
               Caption         =   "TECNICO"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   240
               TabIndex        =   57
               Top             =   360
               Width           =   855
            End
         End
         Begin VB.Label Label8 
            Appearance      =   0  'Flat
            BackColor       =   &H00F5C9B1&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " Codigo"
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
            Left            =   135
            TabIndex        =   34
            Top             =   450
            Width           =   1215
         End
         Begin VB.Label Label7 
            Appearance      =   0  'Flat
            BackColor       =   &H00F5C9B1&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " Serie"
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
            Left            =   120
            TabIndex        =   32
            Top             =   1170
            Width           =   1215
         End
         Begin VB.Label Label6 
            Appearance      =   0  'Flat
            BackColor       =   &H00F5C9B1&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " Articulo"
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
            Left            =   120
            TabIndex        =   30
            Top             =   1890
            Width           =   1215
         End
         Begin VB.Label Label5 
            Appearance      =   0  'Flat
            BackColor       =   &H00F5C9B1&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " Marca"
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
            Left            =   120
            TabIndex        =   29
            Top             =   1530
            Width           =   1215
         End
         Begin VB.Label Label4 
            Appearance      =   0  'Flat
            BackColor       =   &H00F5C9B1&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " Descripcion"
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
            Left            =   135
            TabIndex        =   28
            Top             =   810
            Width           =   1215
         End
      End
      Begin XPFrame.FrameXp FrameXp6 
         Height          =   345
         Left            =   5085
         TabIndex        =   11
         Top             =   8865
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   609
         BackColor       =   49344
         Caption         =   "Datos Financieros"
         CaptionEstilo3D =   1
         BackColor       =   49344
         ColorBarraArriba=   12632319
         ColorBarraAbajo =   128
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
      End
      Begin XPFrame.FrameXp FrameXp7 
         Height          =   345
         Left            =   9000
         TabIndex        =   12
         Top             =   7335
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   609
         BackColor       =   49344
         Caption         =   "Datos Laborales"
         CaptionEstilo3D =   1
         BackColor       =   49344
         ColorBarraArriba=   12632319
         ColorBarraAbajo =   128
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
      End
      Begin XPFrame.FrameXp FrameXp8 
         Height          =   345
         Left            =   7380
         TabIndex        =   13
         Top             =   7755
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   609
         BackColor       =   49344
         Caption         =   "Crédito Directo"
         CaptionEstilo3D =   1
         BackColor       =   49344
         ColorBarraArriba=   12632319
         ColorBarraAbajo =   128
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
      End
      Begin VB.Label Label21 
         Appearance      =   0  'Flat
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Rut"
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
         Left            =   135
         TabIndex        =   24
         Top             =   750
         Width           =   1215
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
         Left            =   2700
         TabIndex        =   23
         Top             =   750
         Width           =   495
      End
      Begin VB.Label Label10 
         Appearance      =   0  'Flat
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Fecha"
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
         Left            =   3465
         TabIndex        =   21
         Top             =   390
         Width           =   1695
      End
      Begin VB.Label Label29 
         Appearance      =   0  'Flat
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Celular"
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
         Left            =   2760
         TabIndex        =   17
         Top             =   1425
         Width           =   1215
      End
      Begin VB.Label Label22 
         Appearance      =   0  'Flat
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Traido Por"
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
         Left            =   120
         TabIndex        =   16
         Top             =   1110
         Width           =   1215
      End
      Begin VB.Label Label20 
         Appearance      =   0  'Flat
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Fono"
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
         Left            =   120
         TabIndex        =   15
         Top             =   1425
         Width           =   1215
      End
      Begin VB.Label Label11 
         Appearance      =   0  'Flat
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Folio"
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
         Left            =   135
         TabIndex        =   14
         Top             =   390
         Width           =   1215
      End
   End
   Begin FlexCell.Grid Grid4 
      Height          =   4110
      Left            =   0
      TabIndex        =   43
      Top             =   0
      Visible         =   0   'False
      Width           =   14760
      _ExtentX        =   26035
      _ExtentY        =   7250
      AllowUserSort   =   -1  'True
      Cols            =   6
      DefaultFontSize =   8.25
      DefaultFontBold =   -1  'True
      Rows            =   30
      MultiSelect     =   0   'False
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash opciones 
      Height          =   1470
      Left            =   6840
      TabIndex        =   0
      Top             =   7875
      Width           =   6075
      _cx             =   10716
      _cy             =   2593
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
Attribute VB_Name = "MMGarantias"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Private c As Cliente
    Private modifica As Boolean
    Private cargo As Boolean
    Private formatogrilla(10, 10) As String
    Private fecha As String
    Private descripcionproducto As String
    Private filtro_rut As String
    Private folio_1 As String
    Private rutserviciotecnico As String
    Private nombreserviciotecnico As String
    Private direccionserviciotecnico As String
    Private fonoserviciotecnico As String
    Private comunaserviciotecnico As String
    Private ciudadserviciotecnico As String
Private Sub cmd3_Click()
If cmd3.Caption = "CREAR MOVIMIENTOS" Then
   observa.Enabled = True
   observa.SetFocus
   cmd3.Caption = "GRABAR MOVIMIENTO"
   cmd3.Enabled = True
Else
If observa.text <> "" Then
    grabarmovimientos
   End If
    cmd3.Caption = "CREAR MOVIMIENTOS"
    observa.Enabled = False
    observa.text = ""
End If

End Sub

Private Sub cmd4_Click()
Call eliminarEspeciales(lista.Cell(lista.ActiveCell.row, 1).text, lista.Cell(lista.ActiveCell.row, 3).text)
cmd4.Visible = False
observa.text = ""
leermovimientos
End Sub

Private Sub cmd5_Click()
modifica = True
grabar
End Sub

Private Sub cmd6_Click()
 titCaption = Replace(infogarantias.Caption, "&", "")
        Load infogarantias
        infogarantias.Caption = titCaption
        infogarantias.Show
End Sub



Private Sub cmd7_Click()
'CARGAGRILLA
 FRMIMPRESION.Visible = True


End Sub

Private Sub Command1_Click()
If Option1.Value = False Then
    Call imprime_guialarga
Else
    Call imprime_guiacorta
End If
modifica = True
If leermovimientosexiste = False Then
Call grabarmovimientosautomatico
End If
Call grabar
FRMIMPRESION.Visible = False
End Sub

Private Sub Command2_Click()
FRMIMPRESION.Visible = False
End Sub

Private Sub entregados_Click()
If entregados.Value = 1 Then
frmentregar.Visible = True
HASTA1.SetFocus
Else
frmentregar.Visible = False
End If


End Sub
    Private Sub Form_Activate()
        Principal.barraEstado.Panels(1).text = UCase(Me.Caption)
        FOLIO.SetFocus
    End Sub
   
Public Sub cargarfolio()
FOLIO_KeyPress (13)
End Sub
    Private Sub Form_Load()
      
        modifica = False
        cargo = False
        Call Centrar(Me)
        observa.Enabled = False
        grilla
         Grid4.Rows = 2
         Grid4.Enabled = False
         Call planillaproveedor
        
    End Sub
    
    Private Sub Form_Unload(Cancel As Integer)
        Principal.barraEstado.Panels(1).text = UCase(Principal.Caption)
        Call limpiaBarra(2)
    End Sub











Private Sub guia_Click()
If guia.Value = 1 And ruttecnico.text = "" Then
frmguia.Visible = True
ruttecnico.SetFocus
Else
frmguia.Visible = False
End If

End Sub

Private Sub Lista_DblClick()
observa.text = lista.Cell(lista.ActiveCell.row, 2).text
cmd4.Visible = True
End Sub

Private Sub lista_KeyDown(KeyCode As Integer, ByVal Shift As Integer)
  Select Case KeyCode
            Case 46
                If lista.ActiveCell.row > 0 Then
                    Call eliminarEspeciales(lista.Cell(lista.ActiveCell.row, 1).text, lista.Cell(lista.ActiveCell.row, 3).text)
                    lista.RemoveItem (lista.ActiveCell.row)
                End If
        End Select
        leermovimientos
End Sub



Private Sub nguia_KeyPress(KeyAscii As Integer)
 KeyAscii = esNumero(KeyAscii)
If nguia.text <> "" And lblnombretecnico.Caption <> "" Then
  If KeyAscii = 13 Then
   nguia.text = ceros(nguia)
   cmd7.Visible = True
   End If
End If
 
End Sub

  Private Sub opciones_FSCommand(ByVal command As String, ByVal args As String)
        Select Case command
            Case "modifica"
                Call modificar
            Case "elimina"
                If MsgBox("DESEA REALMENTE ELIMINAR Si / No", vbYesNo) = vbYes Then
                frmglosaeliminacion.Show vbModal
                Call ELIMINAR
                End If
            Case "imprime"
                EMPRESA
                 Grid4.Rows = 1
                Call IMPRIMIRgarantia
            Case "movimientos"
            Case "historico"
            Case "retorno"
                Call retorno
            Case "anterior"
'                Call anterior
            Case "siguiente"
'                Call siguiente
        End Select
    End Sub
Sub grilla()
Dim K As Integer
    
    lista.Rows = 1
    lista.Cols = 4
    lista.Column(0).Width = 0
    lista.Column(1).Width = 70
    lista.Column(2).Width = 290
    lista.Column(3).Width = 0
      
    lista.Cell(0, 1).text = "FECHA"
    lista.Cell(0, 2).text = "GLOSA"
    lista.Cell(0, 3).text = "HORA"
   
    lista.Column(1).Alignment = cellCenterCenter
    lista.Column(2).Alignment = cellCenterCenter
    lista.Column(3).Alignment = cellCenterCenter
  
   
        lista.AllowUserResizing = False
        lista.DisplayFocusRect = False
        lista.ExtendLastCol = False
        lista.BoldFixedCell = False
        lista.DisplayRowIndex = True
        lista.DrawMode = cellOwnerDraw
        lista.Appearance = Flat
        lista.ScrollBarStyle = Flat
        lista.FixedRowColStyle = Flat
        lista.BackColorFixed = RGB(90, 158, 214)
        lista.BackColorFixedSel = RGB(110, 180, 230)
        lista.BackColorBkg = RGB(90, 158, 214)
        lista.BackColorScrollBar = RGB(231, 235, 247)
        lista.BackColor1 = RGB(231, 235, 247)
        lista.BackColor2 = RGB(239, 243, 255)
        lista.GridColor = RGB(148, 190, 231)
 
    
For K = 1 To 3
lista.Range(0, K, 0, K).Borders(cellEdgeLeft) = cellThick
lista.Range(0, K, 0, K).Borders(cellEdgeTop) = cellThick
lista.Range(0, K, 0, K).Borders(cellEdgeRight) = cellThick
lista.Range(0, K, 0, K).Borders(cellEdgeBottom) = cellThick
Next K

    
End Sub
'gotfocus
'********
Private Sub FOLIO_GotFocus()

        FOLIO.text = leerUltimoFolio
        FOLIO.text = ceros(FOLIO)
        
        Call VerificarCajas(Me, FOLIO)
        Call selecciona(FOLIO)
        FOLIO.SetFocus
End Sub

Private Sub DESDE1_GotFocus()
        Call VerificarCajas(Me, DESDE1)
        Call selecciona(DESDE1)
End Sub
Private Sub DESDE2_GotFocus()
        Call VerificarCajas(Me, DESDE2)
        Call selecciona(DESDE2)
End Sub
Private Sub DESDE3_GotFocus()
 Call VerificarCajas(Me, DESDE3)
        Call selecciona(DESDE3)
End Sub

Private Sub HASTA1_GotFocus()
        Call VerificarCajas(Me, HASTA1)
        Call selecciona(HASTA1)
End Sub
Private Sub HASTA2_GotFocus()
        Call VerificarCajas(Me, HASTA2)
        Call selecciona(HASTA2)
End Sub
Private Sub HASTA3_GotFocus()
 Call VerificarCajas(Me, HASTA3)
        Call selecciona(HASTA3)
End Sub

Private Sub rut_cliente_GotFocus()
  Call VerificarCajas(Me, rut_cliente)
        Call selecciona(rut_cliente)
        Principal.barraEstado.Panels(2).text = "F2: Ayuda Cliente"
End Sub

Private Sub nombre_cliente_GotFocus()
        Call VerificarCajas(Me, nombre_cliente)
        Call selecciona(nombre_cliente)
End Sub

Private Sub nombre2_GotFocus()
        Call VerificarCajas(Me, nombre2)
        Call selecciona(nombre2)
End Sub

Private Sub fono_GotFocus()
        Call VerificarCajas(Me, fono)
        Call selecciona(fono)
End Sub

Private Sub celular_GotFocus()
        Call VerificarCajas(Me, celular)
        Call selecciona(celular)
End Sub

Private Sub codigobarra_GotFocus()
        Call VerificarCajas(Me, codigobarra)
        Call selecciona(codigobarra)
       
        Principal.barraEstado.Panels(2).text = "F2: Ayuda Producto "

End Sub


Private Sub descripcion_GotFocus()
        Call VerificarCajas(Me, descripcion)
        Call selecciona(descripcion)
End Sub




Private Sub ruttecnico_GotFocus()
 Principal.barraEstado.Panels(2).text = "F2: Ayuda Tecnico"
End Sub

Private Sub ruttecnico_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF2 Then
            Call ayudaTecnico(ruttecnico, lbldvtecnico)
End If
End Sub

Private Sub ruttecnico_KeyPress(KeyAscii As Integer)
        KeyAscii = esNumero(KeyAscii)
        If KeyAscii = 13 And ruttecnico.text <> "" Then
             ruttecnico.text = ceros(ruttecnico)
            lbldvtecnico.Caption = rut(ruttecnico.text)
            filtro_rut = (ruttecnico.text + lbldvtecnico.Caption)
            Call leerTecnico(filtro_rut)
          If lblnombretecnico.Caption = "" Then
             ruttecnico.SetFocus
          Else
            nguia.SetFocus
          End If
          
        End If
End Sub

Private Sub ruttecnico_LostFocus()
Principal.barraEstado.Panels(2).text = ""
End Sub

Private Sub serie_GotFocus()
        Call VerificarCajas(Me, serie)
        Call selecciona(serie)
End Sub

Private Sub marca_GotFocus()
        Call VerificarCajas(Me, marca)
        Call selecciona(serie)
End Sub
Private Sub articulo_GotFocus()
        Call VerificarCajas(Me, articulo)
        Call selecciona(articulo)
End Sub

Private Sub falla_GotFocus()
        Call VerificarCajas(Me, falla)
        Call selecciona(falla)
End Sub
Private Sub final_GotFocus()
 Call VerificarCajas(Me, final)
        Call selecciona(final)
End Sub
'fin gotfocus
'************


'************
'keydown

 
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
Private Sub rut_cliente_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = vbKeyF2 Then
            Call ayudaCliente(rut_cliente, nombre_cliente, lbldv)
        Else
            Call Flechas(KeyCode, FOLIO)
        End If
End Sub

Private Sub nombre_cliente_KeyDown(KeyCode As Integer, Shift As Integer)
 Call Flechas(KeyCode, rut_cliente)
End Sub

Private Sub nombre2_KeyDown(KeyCode As Integer, Shift As Integer)
 Call Flechas(KeyCode, HASTA3)
End Sub

Private Sub fono_KeyDown(KeyCode As Integer, Shift As Integer)
Call Flechas(KeyCode, nombre_cliente)
End Sub

Private Sub FOLIO_KeyDown(KeyCode As Integer, Shift As Integer)
Call Flechas(KeyCode, FOLIO)
End Sub

Private Sub celular_KeyDown(KeyCode As Integer, Shift As Integer)
  Call Flechas(KeyCode, fono)
End Sub

Private Sub codigobarra_KeyDown(KeyCode As Integer, Shift As Integer)
       If KeyCode = vbKeyF2 Then
            Call ayudaProductotxt(codigobarra)
        Else
             Call Flechas(KeyCode, nombre2)
        End If

End Sub

Private Sub descripcion_KeyDown(KeyCode As Integer, Shift As Integer)
  Call Flechas(KeyCode, codigobarra)
End Sub
Private Sub serie_KeyDown(KeyCode As Integer, Shift As Integer)
  Call Flechas(KeyCode, descripcion)
End Sub

Private Sub marca_KeyDown(KeyCode As Integer, Shift As Integer)
   Call Flechas(KeyCode, serie)
End Sub

Private Sub articulo_KeyDown(KeyCode As Integer, Shift As Integer)
 Call Flechas(KeyCode, marca)
End Sub

Private Sub falla_KeyDown(KeyCode As Integer, Shift As Integer)
Call Flechas(KeyCode, articulo)
End Sub

Private Sub DESDE1_KeyDown(KeyCode As Integer, Shift As Integer)
Call Flechas(KeyCode, FOLIO)
End Sub
Private Sub DESDE2_KeyDown(KeyCode As Integer, Shift As Integer)
Call Flechas(KeyCode, DESDE1)
End Sub
Private Sub DESDE3_KeyDown(KeyCode As Integer, Shift As Integer)
Call Flechas(KeyCode, DESDE2)
End Sub

Private Sub HASTA1_KeyDown(KeyCode As Integer, Shift As Integer)
Call Flechas(KeyCode, HASTA1)
End Sub
Private Sub HASTA2_KeyDown(KeyCode As Integer, Shift As Integer)
Call Flechas(KeyCode, HASTA1)
End Sub
Private Sub HASTA3_KeyDown(KeyCode As Integer, Shift As Integer)
Call Flechas(KeyCode, HASTA2)
End Sub
Private Sub final_KeyDown(KeyCode As Integer, Shift As Integer)
Call Flechas(KeyCode, nombre2)
End Sub
'fin keydown
'***********


'***********
'keypress

Private Sub FOLIO_KeyPress(KeyAscii As Integer)
 KeyAscii = esNumero(KeyAscii)
        If KeyAscii = 13 And FOLIO.text <> "" Then
            FOLIO.text = ceros(FOLIO)
            Call leergarantias(FOLIO.text)
            
        End If
End Sub
Private Sub rut_cliente_KeyPress(KeyAscii As Integer)
 KeyAscii = esNumero(KeyAscii)
        If KeyAscii = 13 And rut_cliente.text <> "" Then
            rut_cliente.text = ceros(rut_cliente)
            lbldv.Caption = rut(rut_cliente.text)
            filtro_rut = (rut_cliente.text + lbldv.Caption)
            Call LEERCLIENTE(filtro_rut)
           
        End If
End Sub

Private Sub nombre_cliente_KeyPress(KeyAscii As Integer)
  KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If KeyAscii = 13 Then
           fono.SetFocus
        End If
End Sub

Private Sub nombre2_KeyPress(KeyAscii As Integer)
  KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If KeyAscii = 13 And nombre2.text <> "" Then
         final.SetFocus
         
        End If
End Sub

Private Sub fono_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
           celular.SetFocus
        End If
End Sub

Private Sub celular_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then
 codigobarra.SetFocus
 End If
End Sub

Private Sub codigobarra_KeyPress(KeyAscii As Integer)
     KeyAscii = esNumero(KeyAscii)
        If KeyAscii = 13 And codigobarra.text <> "" Then
           codigobarra.text = ceros(codigobarra)
           descripcionproducto = leerNombreProducto(codigobarra.text)
           If descripcionproducto = "" Then
           codigobarra.SetFocus
           Else
           descripcion.text = descripcionproducto
           serie.SetFocus
           End If
           
        End If
End Sub

Private Sub descripcion_KeyPress(KeyAscii As Integer)
  KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If KeyAscii = 13 Then
           serie.SetFocus
        End If
End Sub


Private Sub serie_KeyPress(KeyAscii As Integer)
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If KeyAscii = 13 Then
          marca.SetFocus
        End If
End Sub
Private Sub marca_KeyPress(KeyAscii As Integer)
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If KeyAscii = 13 Then
          articulo.SetFocus
        End If
End Sub

Private Sub articulo_KeyPress(KeyAscii As Integer)
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If KeyAscii = 13 Then
            falla.SetFocus
        End If
End Sub

Private Sub falla_KeyPress(KeyAscii As Integer)

 KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If KeyAscii = 13 And falla.text <> "" Then
                 Call grabar
                
        End If
End Sub

Private Sub DESDE1_KeyPress(KeyAscii As Integer)
        
        KeyAscii = esNumero(KeyAscii)
        If KeyAscii = 13 Then
               DESDE1.text = ceros(DESDE1)
            If DESDE1.text = "00" Then
                DESDE1.text = Format(fechasistema, "dd")
                DESDE2.text = Format(fechasistema, "mm")
                DESDE3.text = Format(fechasistema, "yyyy")
                fecha = DESDE3.text & "-" & DESDE2.text & "-" & DESDE1.text
                 rut_cliente.SetFocus
                Else
                  DESDE2.SetFocus
              End If
        End If
End Sub
Private Sub DESDE2_KeyPress(KeyAscii As Integer)
   KeyAscii = esNumero(KeyAscii)
        If KeyAscii = 13 Then
            DESDE2.text = ceros(DESDE2)
            If DESDE2.text = "00" Then
               DESDE2.text = Format(fechasistema, "mm")
            End If
            DESDE3.SetFocus
        End If
End Sub

Private Sub DESDE3_KeyPress(KeyAscii As Integer)

  KeyAscii = esNumero(KeyAscii)
        If KeyAscii = 13 Then
           DESDE3.text = ceros(DESDE3)
            If DESDE3.text = "0000" Then
                DESDE3.text = Format(fechasistema, "yyyy")
            End If
            fecha = DESDE3.text & "-" & DESDE2.text & "-" & DESDE1.text
           
               rut_cliente.SetFocus
            
        End If
        End Sub
        
        
        Private Sub HASTA1_KeyPress(KeyAscii As Integer)
        
        KeyAscii = esNumero(KeyAscii)
        If KeyAscii = 13 Then
               HASTA1.text = ceros(HASTA1)
            If HASTA1.text = "00" Then
                HASTA1.text = Format(fechasistema, "dd")
                HASTA2.text = Format(fechasistema, "mm")
                HASTA3.text = Format(fechasistema, "yyyy")
                fecha = HASTA3.text & "-" & HASTA2.text & "-" & HASTA1.text
                nombre2.SetFocus
                Else
                  HASTA2.SetFocus
              End If
        End If
End Sub
Private Sub hasta2_KeyPress(KeyAscii As Integer)
   KeyAscii = esNumero(KeyAscii)
        If KeyAscii = 13 Then
            HASTA2.text = ceros(HASTA2)
            If HASTA2.text = "00" Then
               HASTA2.text = Format(fechasistema, "mm")
            End If
            HASTA3.SetFocus
        End If
End Sub

Private Sub hasta3_KeyPress(KeyAscii As Integer)

  KeyAscii = esNumero(KeyAscii)
        If KeyAscii = 13 Then
           HASTA3.text = ceros(HASTA3)
            If HASTA3.text = "0000" Then
                HASTA3.text = Format(fechasistema, "yyyy")
            End If
            fecha = HASTA3.text & "-" & HASTA2.text & "-" & HASTA1.text
           
             nombre2.SetFocus
            
        End If
        End Sub
  Private Sub final_KeyPress(KeyAscii As Integer)
 KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If KeyAscii = 13 And final.text <> "" Then
         cmd5_Click

        End If

End Sub
    
'fin keypress
'************

'************
'lostfocus


Private Sub DESDE1_LostFocus()
Call esfecha(DESDE1, DESDE2, DESDE3, "dd")
End Sub
Private Sub DESDE2_LostFocus()
Call esfecha(DESDE1, DESDE2, DESDE3, "mm")
End Sub
Private Sub DESDE3_LostFocus()
Call esfecha(DESDE1, DESDE2, DESDE3, "yyyy")
End Sub

Private Sub HASTA1_LostFocus()
Call esfecha(HASTA1, HASTA2, HASTA3, "dd")
End Sub
Private Sub HASTA2_LostFocus()
Call esfecha(HASTA1, HASTA2, HASTA3, "mm")
End Sub
Private Sub HASTA3_LostFocus()
Call esfecha(HASTA1, HASTA2, HASTA3, "yyyy")
End Sub

Private Sub rut_cliente_LostFocus()
 Principal.barraEstado.Panels(2).text = ""
End Sub
Private Sub codigobarra_LostFocus()
Principal.barraEstado.Panels(2).text = ""
End Sub

'fin lostfocus
'************



'funciones


Public Function leerUltimoFolio() As String
    
    Dim op As Integer
    Dim sql As New sqlventas.sqlventa
    Dim CAMPOS(3, 3) As String
    
    CAMPOS(0, 0) = "IFNULL(MAX(folio) + 1,'0000000001')"
    CAMPOS(1, 0) = ""
    CAMPOS(0, 2) = "sv_garantias_" + empresaActiva
    condicion = " local='00'"
    op = 5
    sql.response = CAMPOS
    Set sql.conexion = ventasRubro
    Call sql.sqlventas(op, condicion)
    If sql.Status = 0 Then
        If sql.response(0, 3) <> "" And sql.response(0, 3) <> "0" Then
            leerUltimoFolio = sql.response(0, 3)
        Else
            leerUltimoFolio = "0000000001"
        End If
    End If
End Function
'
'
'Public Sub grabar()
'        Dim di As Integer
'
'        Dim campos(20, 3) As String
'        Dim op As Integer
'        campos(0, 0) = "folio"
'        campos(1, 0) = "rut"
'        campos(2, 0) = "fecha"
'        campos(3, 0) = "codigo"
'        campos(4, 0) = "descripcion"
'        campos(5, 0) = "serie"
'        campos(6, 0) = "marca"
'        campos(7, 0) = "articulo"
'        campos(8, 0) = "falla"
'        campos(9, 0) = "fono_cliente"
'        campos(10, 0) = "traidopor"
'        campos(11, 0) = "observacion"
'        campos(12, 0) = "entregado"
'        campos(13, 0) = "retiradopor"
'        campos(14, 0) = "celular_cliente"
'        campos(15, 0) = "local"
'        campos(16, 0) = "estado"
'        campos(17, 0) = ""
'        campos(0, 1) = FOLIO.text
'        campos(1, 1) = rut_cliente.text + lblDV.Caption
'        campos(2, 1) = desde3.text & "-" & desde2.text & "-" & desde1.text
'        campos(3, 1) = codigobarra.text
'        campos(4, 1) = descripcion.text
'        campos(5, 1) = serie.text
'        campos(6, 1) = marca.text
'        campos(7, 1) = articulo.text
'        campos(8, 1) = falla.text
'        campos(9, 1) = fono.text
'        campos(10, 1) = nombre_cliente.text
'        campos(11, 1) = final.text
'        campos(12, 1) = hasta3.text + "-" + hasta2.text + "-" + hasta1.text
'        campos(13, 1) = nombre2.text
'        campos(14, 1) = celular.text
'        campos(15, 1) = "00"
'        campos(16, 1) = entregados.Value
'
'        campos(0, 2) = "sv_garantias_" + empresaActiva
'         If modifica = True Then condicion = "folio= '" + FOLIO.text + "'" Else condicion = ""
'         If modifica = True Then op = 3 Else op = 2
'
'
'        sqlventas.response = campos
'        Set sqlventas.conexion = ventasRubro
'        Call sqlventas.sqlventas(op, condicion)
'        condicion = sqlventas.Status
'
'    FOLIO.text = ""
'    rut_cliente.text = ""
'    lblDV.Caption = ""
'    desde1.text = ""
'    desde2.text = ""
'    desde3.text = ""
'    codigobarra.text = ""
'    descripcion.text = ""
'    serie.text = ""
'    marca.text = ""
'    articulo.text = ""
'    falla.text = ""
'    fono.text = ""
'    nombre_cliente = ""
'    observa.text = ""
'    hasta1.text = ""
'    hasta2.text = ""
'    hasta3.text = ""
'    nombre2.text = ""
'    celular.text = ""
'    fecha = ""
'    modifica = False
''    Call HabilitarCajas(Me, modifica)
'    FOLIO.Enabled = True
'    rut_cliente.Enabled = True
'    Call retorno
'    End Sub
    
    
    
    
Public Sub grabar()
        Dim di As Integer
        
        Dim CAMPOS(20, 3) As String
        Dim op As Integer
        CAMPOS(0, 0) = "folio"
        CAMPOS(1, 0) = "rut"
        CAMPOS(2, 0) = "fecha"
        CAMPOS(3, 0) = "codigo"
        CAMPOS(4, 0) = "descripcion"
        CAMPOS(5, 0) = "serie"
        CAMPOS(6, 0) = "marca"
        CAMPOS(7, 0) = "articulo"
        CAMPOS(8, 0) = "falla"
        CAMPOS(9, 0) = "fono_cliente"
        CAMPOS(10, 0) = "traidopor"
        CAMPOS(11, 0) = "observacion"
        CAMPOS(12, 0) = "entregado"
        CAMPOS(13, 0) = "retiradopor"
        CAMPOS(14, 0) = "celular_cliente"
        CAMPOS(15, 0) = "local"
        CAMPOS(16, 0) = "estado"
        CAMPOS(17, 0) = "tecnico"
        CAMPOS(18, 0) = "numeroguia"
        CAMPOS(19, 0) = ""
        
        pivote.MaxLength = 10
        pivote.text = leerUltimoFolio
        pivote.text = ceros(pivote)
        
'        campos(0, 1) = pivote.text
        CAMPOS(1, 1) = rut_cliente.text + lbldv.Caption
        CAMPOS(2, 1) = DESDE3.text & "-" & DESDE2.text & "-" & DESDE1.text
        CAMPOS(3, 1) = codigobarra.text
        CAMPOS(4, 1) = descripcion.text
        CAMPOS(5, 1) = serie.text
        CAMPOS(6, 1) = marca.text
        CAMPOS(7, 1) = articulo.text
        CAMPOS(8, 1) = falla.text
        CAMPOS(9, 1) = fono.text
        CAMPOS(10, 1) = nombre_cliente.text
        CAMPOS(11, 1) = final.text
        CAMPOS(12, 1) = HASTA3.text + "-" + HASTA2.text + "-" + HASTA1.text
        CAMPOS(13, 1) = nombre2.text
        CAMPOS(14, 1) = celular.text
        CAMPOS(15, 1) = "00"
        CAMPOS(16, 1) = entregados.Value
        CAMPOS(17, 1) = ruttecnico.text & lbldvtecnico.Caption
        CAMPOS(18, 1) = nguia.text
        CAMPOS(0, 2) = "sv_garantias_" + empresaActiva
         If modifica = True Then
         CAMPOS(0, 1) = FOLIO.text
         condicion = "folio= '" + FOLIO.text + "'"
         Else: condicion = ""
        pivote.MaxLength = 10
        pivote.text = leerUltimoFolio
        pivote.text = ceros(pivote)
        CAMPOS(0, 1) = pivote.text
         End If
         
         If modifica = True Then op = 3 Else op = 2
        
       
        sqlventas.response = CAMPOS
        Set sqlventas.conexion = ventasRubro
        sqlventas.audit = True: sqlventas.programaactivo = Me.Caption: sqlventas.programaactivo = Me.Caption
        Set sqlventas.conauditoria = conauditoria: sqlventas.usuarioauditoria = usuarioSistema
        Call sqlventas.sqlventas(op, condicion)
        condicion = sqlventas.Status
   
    FOLIO.text = ""
    rut_cliente.text = ""
    lbldv.Caption = ""
    DESDE1.text = ""
    DESDE2.text = ""
    DESDE3.text = ""
    codigobarra.text = ""
    descripcion.text = ""
    serie.text = ""
    marca.text = ""
    articulo.text = ""
    falla.text = ""
    fono.text = ""
    nombre_cliente = ""
    observa.text = ""
    HASTA1.text = ""
    HASTA2.text = ""
    HASTA3.text = ""
    nombre2.text = ""
    celular.text = ""
    ruttecnico.text = ""
    lbldvtecnico.Caption = ""
    lblnombretecnico.Caption = ""
    fecha = ""
    nguia.text = ""
    modifica = False
    frmguia.Visible = False
    guia.Value = 0
'    Call HabilitarCajas(Me, modifica)
    FOLIO.Enabled = True
    rut_cliente.Enabled = True
    frmguia.Visible = False
    Call retorno
    End Sub

Sub LEERCLIENTE(rut_cli)

Dim op As Integer
Dim CAMPOS(4, 4) As String

    CAMPOS(0, 0) = "nombre"
    CAMPOS(1, 0) = "fono1"
    CAMPOS(2, 0) = "celular"
    CAMPOS(3, 0) = "direccion"
    CAMPOS(4, 0) = ""
    CAMPOS(0, 2) = "sv_maestroclientes"
    condicion = "rut='" & rut_cli & "'"
    op = 5
    Set sqlventas.conexion = ventas
    sqlventas.response = CAMPOS
    Call sqlventas.sqlventas(op, condicion)

    If sqlventas.Status = 4 Then
    
    If MsgBox("El rut ingresado no se encuentra. ¿Desea crearlo?", vbYesNo, "Mensaje") = vbYes Then
                    Load MClientes
                    MClientes.dato1.text = rut_cliente.text
                    MClientes.lbldv.Caption = lbldv.Caption
                    MClientes.dato2.text = "0"
                    MClientes.Show
                    Else
                    rut_cliente.SetFocus
                End If
    
    Else
    
    nombre_cliente.text = sqlventas.response(0, 3)
    fono.text = sqlventas.response(1, 3)
    celular.text = sqlventas.response(2, 3)
    direccioncliente.text = sqlventas.response(3, 3)
    codigobarra.SetFocus
    
    End If
    

End Sub

 Private Sub leergarantias(folio1)
        Dim total As Double
        Dim tabla As String
        Dim empresabusca As String
        Dim resultados As rdoResultset
        Dim csql As New rdoQuery
        Set csql.ActiveConnection = ventasRubro

        
        csql.sql = "SELECT folio,rut,fecha,codigo,descripcion,serie,marca,articulo,falla,fono_cliente,traidopor,observacion,entregado,retiradopor,celular_cliente,estado,tecnico,numeroguia "
        csql.sql = csql.sql & "FROM sv_garantias_" + empresaActiva
        csql.sql = csql.sql & " WHERE folio= '" & folio1 & "' "
        csql.Execute
        
        If csql.RowsAffected > 0 Then
            Set resultados = csql.OpenResultset
            FOLIO.text = resultados(0)
            rut_cliente.text = Mid(resultados("rut"), 1, 9)
            lbldv.Caption = Mid(resultados("rut"), 10, 1)
            DESDE1.text = Format(resultados("fecha"), "dd")
            DESDE2.text = Format(resultados("fecha"), "mm")
            DESDE3.text = Format(resultados("fecha"), "yyyy")
            nombre_cliente.text = resultados(10)
            fono.text = resultados(9)
            celular.text = resultados(14)
            nombre2.text = resultados(13)
            codigobarra.text = resultados(3)
            descripcion.text = resultados(4)
            serie.text = resultados(5)
            marca.text = resultados(6)
            articulo.text = resultados(7)
            falla.text = resultados(8)
           
            If IsNull(resultados("entregado")) = True Then
                 HASTA1.text = ""
                 HASTA2.text = ""
                 HASTA3.text = ""
            Else
                 HASTA1.text = Format(resultados("entregado"), "dd")
                 HASTA2.text = Format(resultados("entregado"), "mm")
                 HASTA3.text = Format(resultados("entregado"), "yyyy")
            End If
            
            If resultados("estado") = 0 Then
                 entregados.Value = 0
                 entregados_Click
            Else
                 entregados.Value = 1
                 entregados_Click
             
            If IsNull(resultados("observacion")) = True Then
                final.text = ""
            Else
                final.text = resultados("observacion")
            End If
            
            End If
            
             If resultados("tecnico") = "" Then
               ruttecnico.text = ""
               lbldvtecnico.Caption = ""
               lblnombretecnico.Caption = ""
               frmguia.Visible = False
               cmd7.Visible = False
               guia.Enabled = True
               nguia.text = ""
            Else
              Call leerTecnico(resultados("tecnico"))
               ruttecnico.text = Mid(resultados("tecnico"), 1, 9)
               lbldvtecnico.Caption = Mid(resultados("tecnico"), 10, 1)
               lblnombretecnico.Caption = leerNombreTecnico(resultados("tecnico"))
               nguia.text = resultados("numeroguia")
               frmguia.Visible = True
               cmd7.Visible = True
               guia.Enabled = False
            End If
            
             resultados.Close
             cmd3.Visible = True
             leermovimientos
             'opciones.SetFocus
            Else
            DESDE1.SetFocus
        End If
    End Sub
    Private Sub grabarmovimientos()
        Dim di As Integer
        
        Dim CAMPOS(20, 3) As String
        Dim op As Integer
        
        CAMPOS(0, 0) = "folio"
        CAMPOS(1, 0) = "rut_cliente"
        CAMPOS(2, 0) = "hora"
        CAMPOS(3, 0) = "fecha"
        CAMPOS(4, 0) = "glosa"
        CAMPOS(5, 0) = ""
        CAMPOS(0, 1) = FOLIO.text
        CAMPOS(1, 1) = rut_cliente.text + lbldv.Caption
        CAMPOS(2, 1) = Time
        CAMPOS(3, 1) = Format(fechasistema, "yyyy-mm-dd")
        CAMPOS(4, 1) = observa.text
        
        CAMPOS(0, 2) = "sv_movimientos_garantias_" + empresaActiva
        condicion = ""
        op = 2
               
        sqlventas.response = CAMPOS
        Set sqlventas.conexion = ventasRubro
        sql.audit = True: sql.programaactivo = Me.Caption
        Set sql.conauditoria = conauditoria: sql.usuarioauditoria = usuarioSistema
        Call sqlventas.sqlventas(op, condicion)
        condicion = sqlventas.Status
        leermovimientos
    End Sub
    
    Private Sub grabarmovimientosautomatico()
        Dim di As Integer
        
        Dim CAMPOS(20, 3) As String
        Dim op As Integer
        
        CAMPOS(0, 0) = "folio"
        CAMPOS(1, 0) = "rut_cliente"
        CAMPOS(2, 0) = "hora"
        CAMPOS(3, 0) = "fecha"
        CAMPOS(4, 0) = "glosa"
        CAMPOS(5, 0) = ""
        CAMPOS(0, 1) = FOLIO.text
        CAMPOS(1, 1) = rut_cliente.text + lbldv.Caption
        CAMPOS(2, 1) = Time
        CAMPOS(3, 1) = Format(fechasistema, "yyyy-mm-dd")
        CAMPOS(4, 1) = "ENVIADO A SERVICIO TECNICO AUTOMATICO"
        
        CAMPOS(0, 2) = "sv_movimientos_garantias_" + empresaActiva
        condicion = ""
        op = 2
               
        sqlventas.response = CAMPOS
        Set sqlventas.conexion = ventasRubro
        Call sqlventas.sqlventas(op, condicion)
        condicion = sqlventas.Status
'        leermovimientos
    End Sub
 Private Function leermovimientosexiste() As Boolean
    Dim resultados As rdoResultset
    Dim csql As New rdoQuery
    Dim rut As String
        Set csql.ActiveConnection = ventasRubro
        csql.sql = "SELECT fecha,glosa,hora "
        csql.sql = csql.sql + "FROM sv_movimientos_garantias_" & empresaActiva & " "
        csql.sql = csql.sql + "WHERE folio='" & FOLIO.text & "' and glosa='ENVIADO A SERVICIO TECNICO AUTOMATICO' "
        csql.Execute
        lista.Rows = 1
        lista.AutoRedraw = False
        If csql.RowsAffected > 0 Then
         leermovimientosexiste = True
         Else
         leermovimientosexiste = False
        End If
      
       
      
End Function


Sub leermovimientos()
    Dim resultados As rdoResultset
    Dim csql As New rdoQuery
    Dim rut As String
    
                
        Set csql.ActiveConnection = ventasRubro
        csql.sql = "SELECT fecha,glosa,hora "
        csql.sql = csql.sql + "FROM sv_movimientos_garantias_" & empresaActiva & " "
        csql.sql = csql.sql + "WHERE folio='" + FOLIO.text + "' "
        csql.Execute
        lista.Rows = 1
        lista.AutoRedraw = False
        If csql.RowsAffected > 0 Then
        Set resultados = csql.OpenResultset
             
            
            While Not resultados.EOF
              lista.AddItem ""
             lista.Cell(lista.Rows - 1, 1).text = resultados(0)
             lista.Cell(lista.Rows - 1, 2).text = resultados(1)
             lista.Cell(lista.Rows - 1, 3).text = resultados(2)
          
                resultados.MoveNext
            Wend
             lista.AutoRedraw = True
             lista.Refresh
            resultados.Close
            Set resultados = Nothing
        End If
      
       
      
End Sub

  Private Sub eliminarEspeciales(ByVal fecha As String, ByVal HORA As String)
'        Dim resultados As rdoResultset
'        Dim cSql As New rdoQuery
'        Dim rut As String
'        Set cSql.ActiveConnection = ventasRubro
'        cSql.sql = "DELETE  "
'        cSql.sql = cSql.sql + "FROM sv_movimientos_garantias_" + empresaActiva + " "
'        cSql.sql = cSql.sql + "WHERE folio='" + folio.text + "' AND fecha='" + Format(fecha, "yyyy-mm-dd") + "' AND hora='" & HORA & "' "
'        cSql.Execute
        
        
        
        
        Dim op As Integer
        Dim CAMPOS(5, 5) As String
        Set sql = New sqlventas.sqlventa
        condicion = "folio='" + FOLIO.text + "' AND fecha='" + Format(fecha, "yyyy-mm-dd") + "' AND hora='" & HORA & "' "
        op = 4
        CAMPOS(0, 2) = "sv_movimientos_garantias_" + empresaActiva + " "
        sql.response = CAMPOS
        Set sql.conexion = ventasRubro
        sql.audit = True: sql.programaactivo = MServicio.Caption
        Set sql.conauditoria = conauditoria: sql.usuarioauditoria = usuarioSistema
        sql.glosaeliminacion = glosaeliminacionsistema
        sql.solicitoeliminacion = solicitaeliminacion
        Call sql.sqlventas(op, condicion)
        End Sub
        
        Private Sub eliminarmovimientos()
'        Dim resultados As rdoResultset
'        Dim cSql As New rdoQuery
'        Dim rut As String
'        Set cSql.ActiveConnection = ventasRubro
'        cSql.sql = "DELETE  "
'        cSql.sql = cSql.sql + "FROM sv_movimientos_garantias_" + empresaActiva + " "
'        cSql.sql = cSql.sql + "WHERE folio='" + folio.text + "' "
'        cSql.Execute
        
        
        
        
        
        Dim op As Integer
        Dim CAMPOS(6, 6) As String
        Set sql = New sqlventas.sqlventa
        condicion = " folio='" + FOLIO.text + "'  "
        op = 4
        CAMPOS(0, 2) = "sv_movimientos_garantias_" + empresaActiva + " "
        sql.response = CAMPOS
        Set sql.conexion = ventasRubro
        sql.audit = True: sql.programaactivo = MServicio.Caption
        Set sql.conauditoria = conauditoria: sql.usuarioauditoria = usuarioSistema
        sql.glosaeliminacion = glosaeliminacionsistema
        sql.solicitoeliminacion = solicitaeliminacion
        Call sql.sqlventas(op, condicion)
        End Sub
        
        
 'fin funciones
 '*************
 
' opciones
' *********

 Private Sub modificar()
        modifica = True
'        Call HabilitarCajas(Me, modifica)
        FOLIO.Enabled = False
        rut_cliente.Enabled = False
        nombre_cliente.SetFocus
        cmd3.Visible = True
        cmd4.Visible = False
                
    End Sub
Private Sub retorno()
        Call LimpiarCajas(MMGarantias)
        lbldv.Caption = ""
        modifica = False
        cargo = False
'        Call DeshabilitarCajas(Me)
        cmd3.Visible = False
        FOLIO.Enabled = True
        rut_cliente.Enabled = True
        frmentregar.Visible = False
        entregados.Value = 0
        cmd7.Visible = False
        lbldvtecnico.Caption = ""
        lblnombretecnico.Caption = ""
        grilla
        guia.Enabled = True
        guia.Value = 0
        frmguia.Visible = False
        FOLIO.SetFocus
    End Sub
    
     Private Sub ELIMINAR()
        
        Dim op As Integer
        Dim CAMPOS(5, 5) As String
        Set sql = New sqlventas.sqlventa
        condicion = " folio='" + FOLIO.text + "'  "
        op = 4
        CAMPOS(0, 2) = "sv_garantias_" + empresaActiva + " "
        sql.response = CAMPOS
        Set sql.conexion = ventasRubro
        sql.audit = True: sql.programaactivo = MServicio.Caption
        Set sql.conauditoria = conauditoria: sql.usuarioauditoria = usuarioSistema
        sql.glosaeliminacion = glosaeliminacionsistema
        sql.solicitoeliminacion = solicitaeliminacion
        Call sql.sqlventas(op, condicion)
      
        Call eliminarmovimientos
         Call LimpiarCajas(MMGarantias)
            grilla
            entregados.Value = 0
            entregados_Click
            lbldv.Caption = ""
            fecha = ""
            FOLIO.SetFocus
    End Sub
    
    Sub IMPRIMIRgarantia()
    Dim row As Integer
    Dim FINROW As Integer
    Dim i As Integer
    Dim objReportTitle As FlexCell.ReportTitle
    
    'Logo
    'lista.Images.Add App.Path & "\Logo.gif", "Logo"
    'Set objReportTitle = New FlexCell.ReportTitle
    'objReportTitle.ImageKey = "Logo"
    'Grid3.ReportTitles.Add objReportTitle
    Grid4.PageSetup.BlackAndWhite = False
    Grid4.PageSetup.BottomMargin = 1
    Grid4.PageSetup.LeftMargin = 1
    Grid4.PageSetup.RightMargin = 1
    Grid4.PageSetup.TopMargin = 1
    Grid4.PageSetup.PrintFixedRow = True
    Grid4.Column(1).Width = 13 * 8
    Call cabeza
'    Grid4.Range(0, 1, 0, 7).Borders(cellEdgeTop) = cellThin
'    Grid4.Range(0, 1, 0, 7).Borders(cellEdgeBottom) = cellThin
'    Grid4.Range(0, 1, 0, 7).Borders(cellEdgeLeft) = cellThin
'    Grid4.Range(0, 1, 0, 7).Borders(cellEdgeRight) = cellThin
'    Grid4.Range(0, 1, 0, 7).Borders(cellInsideVertical) = cellThin
    FINROW = Grid4.Rows
'    Grid4.Cell(FINROW + 1, 0).text = ""
    Grid4.Rows = Grid4.Rows + 19
'    Grid4.Range(FINROW + 1, 1, FINROW + 1, 7).Borders(cellEdgeTop) = cellThin
    Grid4.Range(FINROW + 1, 1, FINROW + 6, 7).Borders(cellInsideHorizontal) = cellThin
    Grid4.Cell(FINROW + 1, 1).text = "PRODUCTO :"
    Grid4.Cell(FINROW + 1, 2).text = codigobarra.text
    Grid4.Cell(FINROW + 2, 1).text = "DESCRIPCION :"
    Grid4.Cell(FINROW + 2, 2).text = descripcion.text
    Grid4.Cell(FINROW + 3, 1).text = "SERIE :"
    Grid4.Cell(FINROW + 3, 2).text = serie.text
    
    Grid4.Cell(FINROW + 4, 1).text = "MARCA :"
    Grid4.Cell(FINROW + 4, 2).text = marca.text
    Grid4.Cell(FINROW + 5, 1).text = "ARTICULO :"
    Grid4.Cell(FINROW + 5, 2).text = articulo.text
'    Grid4.Cell(FINROW + 5, 1).text = "TELEFONO"
'    Grid4.Cell(FINROW + 5, 2).text = fono.text
    Grid4.Cell(FINROW + 6, 1).text = "CELULAR"
    Grid4.Cell(FINROW + 6, 2).text = celular.text
    Grid4.Column(1).Locked = False
    Grid4.Column(2).Locked = False
    Grid4.Column(3).Locked = False
    Grid4.Column(4).Locked = False
    Grid4.Column(5).Locked = False
    Grid4.Column(6).Locked = False
    Grid4.Column(7).Locked = False
'    Grid4.Range(FINROW + 6, 1, FINROW + 9, 7).Merge
    Grid4.Cell(FINROW + 6, 1).text = "FALLA :"
    Grid4.Cell(FINROW + 6, 2).text = Mid(falla.text, 1, 70)
    Grid4.Cell(FINROW + 7, 2).text = Mid(falla.text, 71, 70)
    Grid4.Cell(FINROW + 8, 2).text = Mid(falla.text, 141, 70)
    Grid4.Cell(FINROW + 9, 2).text = Mid(falla.text, 211, 30)
    If entregados.Value = 1 Then
    Grid4.Cell(FINROW + 10, 1).text = "OBSERVACION :"
    Grid4.Cell(FINROW + 10, 2).text = final.text
    End If
    Grid4.Range(FINROW + 12, 1, FINROW + 12, 7).Borders(cellEdgeTop) = cellThin
    
    Grid4.Range(FINROW + 13, 1, FINROW + 13, 7).Merge
    Grid4.Range(FINROW + 13, 1, FINROW + 13, 7).FontSize = 7
    Grid4.Range(FINROW + 13, 1, FINROW + 13, 7).FontBold = True
    Grid4.Range(FINROW + 13, 1, FINROW + 13, 7).Alignment = cellCenterCenter
    Grid4.Cell(FINROW + 13, 1).text = " LA EMPRESA NO RESPONDE POR DAÑOS FISICOS NI GOLPES DE CORRIENTE "
    
    Grid4.Range(FINROW + 14, 1, FINROW + 14, 7).Merge
    Grid4.Range(FINROW + 14, 1, FINROW + 14, 7).FontSize = 7
    Grid4.Range(FINROW + 14, 1, FINROW + 14, 7).FontBold = True
    Grid4.Cell(FINROW + 14, 1).text = " "
    
    Grid4.Range(FINROW + 17, 2, FINROW + 17, 2).Merge
    Grid4.Range(FINROW + 17, 2, FINROW + 17, 2).Borders(cellEdgeTop) = cellThin
    Grid4.Cell(FINROW + 17, 2).text = "                   DEPTO GARANTIAS"
    
    Grid4.Range(FINROW + 17, 4, FINROW + 17, 6).Merge
    Grid4.Range(FINROW + 17, 4, FINROW + 17, 6).Alignment = cellLeftCenter
    Grid4.Range(FINROW + 17, 4, FINROW + 17, 6).Borders(cellEdgeTop) = cellThin
    Grid4.Cell(FINROW + 17, 4).text = "                   CLIENTE"
    
    Grid4.Column(2).Locked = True
    Grid4.Column(3).Locked = True
    Grid4.Column(4).Locked = True
    Grid4.Column(5).Locked = True
    Grid4.Column(6).Locked = True
    Grid4.Column(7).Locked = True
    Grid4.PageSetup.BlackAndWhite = True
    
    
    
    
    
    
     'Logo
    'lista.Images.Add App.Path & "\Logo.gif", "Logo"
    'Set objReportTitle = New FlexCell.ReportTitle
    'objReportTitle.ImageKey = "Logo"
    'Grid3.ReportTitles.Add objReportTitle
    Grid4.PageSetup.BlackAndWhite = False
    Grid4.PageSetup.BottomMargin = 1
    Grid4.PageSetup.LeftMargin = 1
    Grid4.PageSetup.RightMargin = 1
    Grid4.PageSetup.TopMargin = 1
    Grid4.PageSetup.PrintFixedRow = True
    Grid4.Column(1).Width = 13 * 8
'    Call cabeza
'    Grid4.Range(0, 1, 0, 7).Borders(cellEdgeTop) = cellThin
'    Grid4.Range(0, 1, 0, 7).Borders(cellEdgeBottom) = cellThin
'    Grid4.Range(0, 1, 0, 7).Borders(cellEdgeLeft) = cellThin
'    Grid4.Range(0, 1, 0, 7).Borders(cellEdgeRight) = cellThin
'    Grid4.Range(0, 1, 0, 7).Borders(cellInsideVertical) = cellThin
    FINROW = Grid4.Rows
'    Grid4.Cell(FINROW + 1, 0).text = ""
    Grid4.Rows = Grid4.Rows + 28
'    Grid4.Range(FINROW + 1, 1, FINROW + 1, 7).Borders(cellEdgeTop) = cellThin
    Grid4.Range(FINROW + 9, 1, FINROW + 15, 7).Borders(cellInsideHorizontal) = cellThin
    
    
    
    
    Grid4.Column(1).Locked = False
    Grid4.Column(2).Locked = False
    Grid4.Column(3).Locked = False
    Grid4.Column(4).Locked = False
    Grid4.Column(5).Locked = False
    Grid4.Column(6).Locked = False
    Grid4.Column(7).Locked = False
     If entregados.Value = 0 Then
  
    Grid4.Range(FINROW + 1, 1, FINROW + 1, 7).Merge
    Grid4.Range(FINROW + 1, 1, FINROW + 1, 7).FontSize = 10
    Grid4.Range(FINROW + 1, 1, FINROW + 1, 7).FontBold = True
    Grid4.Range(FINROW + 1, 1, FINROW + 1, 7).Alignment = cellCenterCenter
    Grid4.Cell(FINROW + 1, 1).text = "RECEPCION SERVICIO TECNICO "
    
    Grid4.Range(FINROW + 2, 1, FINROW + 2, 7).Merge
    Grid4.Range(FINROW + 2, 1, FINROW + 2, 7).FontSize = 10
    Grid4.Range(FINROW + 2, 1, FINROW + 2, 7).FontBold = True
    Grid4.Range(FINROW + 2, 1, FINROW + 2, 7).Alignment = cellCenterCenter
    Grid4.Cell(FINROW + 2, 1).text = " Nº : " + FOLIO.text

    
    Grid4.Range(FINROW + 3, 1, FINROW + 3, 7).Merge
    Grid4.Range(FINROW + 3, 1, FINROW + 3, 7).FontSize = 10
    Grid4.Range(FINROW + 3, 1, FINROW + 3, 7).FontBold = True
    Grid4.Range(FINROW + 3, 1, FINROW + 3, 7).Alignment = cellCenterCenter
    Grid4.Cell(FINROW + 3, 1).text = "FECHA RECEPCION : " + DESDE1.text + "-" + DESDE2.text + "-" + DESDE3.text
   
    Else
     Grid4.Range(FINROW + 1, 1, FINROW + 1, 7).Merge
    Grid4.Range(FINROW + 1, 1, FINROW + 1, 7).FontSize = 10
    Grid4.Range(FINROW + 1, 1, FINROW + 1, 7).FontBold = True
    Grid4.Range(FINROW + 1, 1, FINROW + 1, 7).Alignment = cellCenterCenter
    Grid4.Cell(FINROW + 1, 1).text = "RETIRO SERVICIO TECNICO "
     
    Grid4.Range(FINROW + 2, 1, FINROW + 2, 7).Merge
    Grid4.Range(FINROW + 2, 1, FINROW + 2, 7).FontSize = 10
    Grid4.Range(FINROW + 2, 1, FINROW + 2, 7).FontBold = True
    Grid4.Range(FINROW + 2, 1, FINROW + 2, 7).Alignment = cellCenterCenter
    Grid4.Cell(FINROW + 2, 1).text = " Nº : " + FOLIO.text
   
     
    Grid4.Range(FINROW + 3, 1, FINROW + 3, 7).Merge
    Grid4.Range(FINROW + 3, 1, FINROW + 3, 7).FontSize = 10
    Grid4.Range(FINROW + 3, 1, FINROW + 3, 7).FontBold = True
    Grid4.Range(FINROW + 3, 1, FINROW + 3, 7).Alignment = cellCenterCenter
    Grid4.Cell(FINROW + 3, 1).text = "FECHA RETIRO : " + DESDE1.text + "-" + DESDE2.text + "-" + DESDE3.text
   
 
    End If
 
    Grid4.Range(FINROW + 4, 1, FINROW + 4, 7).Merge
    Grid4.Range(FINROW + 4, 1, FINROW + 4, 7).FontSize = 10
    Grid4.Range(FINROW + 4, 1, FINROW + 4, 7).FontBold = True
    Grid4.Cell(FINROW + 4, 1).text = "CLIENTE : " + rut_cliente.text + "-" + lbldv.Caption

   
    Grid4.Range(FINROW + 5, 1, FINROW + 5, 7).Merge
    Grid4.Range(FINROW + 5, 1, FINROW + 5, 7).FontSize = 10
    Grid4.Range(FINROW + 5, 1, FINROW + 5, 7).FontBold = True
    Grid4.Cell(FINROW + 5, 1).text = "NOMBRE  : " + nombre_cliente.text
    
     
    Grid4.Range(FINROW + 6, 1, FINROW + 6, 7).Merge
    Grid4.Range(FINROW + 6, 1, FINROW + 6, 7).FontSize = 10
    Grid4.Range(FINROW + 6, 1, FINROW + 6, 7).FontBold = True
    Grid4.Cell(FINROW + 6, 1).text = "TELEFONO : " + fono.text

     
    Grid4.Range(FINROW + 7, 1, FINROW + 7, 7).Merge
    Grid4.Range(FINROW + 7, 1, FINROW + 7, 7).FontSize = 10
    Grid4.Range(FINROW + 7, 1, FINROW + 7, 7).FontBold = True
    Grid4.Cell(FINROW + 7, 1).text = "CELULAR : " + celular.text
   
    
    If entregados.Value = 1 Then
'    Set objReportTitle = New FlexCell.ReportTitle
'    objReportTitle.text = "FECHA RETIRO :" & " " & hasta1.text + "-" + hasta2.text + "-" + hasta3.text
'    objReportTitle.Font.Name = "verdana"
'    objReportTitle.Font.Size = 8
'    objReportTitle.Font.Bold = True
'    objReportTitle.PrintOnAllPages = True
'    objReportTitle.Align = cellLeft
'    Grid4.ReportTitles.Add objReportTitle
     
   Grid4.Range(FINROW + 8, 1, FINROW + 8, 7).Merge
   Grid4.Range(FINROW + 8, 1, FINROW + 8, 7).FontSize = 10
   Grid4.Range(FINROW + 8, 1, FINROW + 8, 7).FontBold = True
   Grid4.Cell(FINROW + 8, 1).text = "RETIRADO POR : " + nombre2.text


    End If
    Grid4.Column(2).Locked = True
    Grid4.Column(3).Locked = True
    Grid4.Column(4).Locked = True
    Grid4.Column(5).Locked = True
    Grid4.Column(6).Locked = True
    Grid4.Column(7).Locked = True
    
    Grid4.Cell(FINROW + 9, 1).text = "PRODUCTO :"
    Grid4.Cell(FINROW + 9, 2).text = codigobarra.text
    Grid4.Cell(FINROW + 10, 1).text = "DESCRIPCION :"
    Grid4.Cell(FINROW + 10, 2).text = descripcion.text
    Grid4.Cell(FINROW + 11, 1).text = "SERIE :"
    Grid4.Cell(FINROW + 11, 2).text = serie.text
    
    Grid4.Cell(FINROW + 12, 1).text = "MARCA :"
    Grid4.Cell(FINROW + 12, 2).text = marca.text
    Grid4.Cell(FINROW + 13, 1).text = "ARTICULO :"
    Grid4.Cell(FINROW + 13, 2).text = articulo.text
'    Grid4.Cell(FINROW + 5, 1).text = "TELEFONO"
'    Grid4.Cell(FINROW + 5, 2).text = fono.text
    Grid4.Cell(FINROW + 14, 1).text = "CELULAR"
    Grid4.Cell(FINROW + 14, 2).text = celular.text
    Grid4.Column(1).Locked = False
    Grid4.Column(2).Locked = False
    Grid4.Column(3).Locked = False
    Grid4.Column(4).Locked = False
    Grid4.Column(5).Locked = False
    Grid4.Column(6).Locked = False
    Grid4.Column(7).Locked = False
'    Grid4.Range(FINROW + 6, 1, FINROW + 9, 7).Merge
    Grid4.Cell(FINROW + 15, 1).text = "FALLA :"
    Grid4.Cell(FINROW + 15, 2).text = Mid(falla.text, 1, 70)
    Grid4.Cell(FINROW + 16, 2).text = Mid(falla.text, 71, 70)
    Grid4.Cell(FINROW + 17, 2).text = Mid(falla.text, 141, 70)
    Grid4.Cell(FINROW + 18, 2).text = Mid(falla.text, 211, 30)
    If entregados.Value = 1 Then
    Grid4.Cell(FINROW + 19, 1).text = "OBSERVACION :"
    Grid4.Cell(FINROW + 19, 2).text = final.text
    End If
    Grid4.Range(FINROW + 20, 1, FINROW + 20, 7).Borders(cellEdgeTop) = cellThin
    
    Grid4.Range(FINROW + 21, 1, FINROW + 21, 7).Merge
    Grid4.Range(FINROW + 21, 1, FINROW + 21, 7).FontSize = 7
    Grid4.Range(FINROW + 21, 1, FINROW + 21, 7).FontBold = True
    Grid4.Range(FINROW + 21, 1, FINROW + 21, 7).Alignment = cellCenterCenter
    Grid4.Cell(FINROW + 21, 1).text = " LA EMPRESA NO RESPONDE POR DAÑOS FISICOS NI GOLPES DE CORRIENTE "
    
    Grid4.Range(FINROW + 22, 1, FINROW + 22, 7).Merge
    Grid4.Range(FINROW + 22, 1, FINROW + 22, 7).FontSize = 7
    Grid4.Range(FINROW + 22, 1, FINROW + 22, 7).FontBold = True
    Grid4.Cell(FINROW + 22, 1).text = " "
    
    Grid4.Range(FINROW + 24, 2, FINROW + 24, 2).Merge
    Grid4.Range(FINROW + 24, 2, FINROW + 24, 2).Borders(cellEdgeTop) = cellThin
    Grid4.Cell(FINROW + 24, 2).text = "                   DEPTO GARANTIAS"
    
    Grid4.Range(FINROW + 24, 4, FINROW + 24, 6).Merge
    Grid4.Range(FINROW + 24, 4, FINROW + 24, 6).Alignment = cellLeftCenter
    Grid4.Range(FINROW + 24, 4, FINROW + 24, 6).Borders(cellEdgeTop) = cellThin
    Grid4.Cell(FINROW + 24, 4).text = "                   CLIENTE"
    
    Grid4.Column(2).Locked = True
    Grid4.Column(3).Locked = True
    Grid4.Column(4).Locked = True
    Grid4.Column(5).Locked = True
    Grid4.Column(6).Locked = True
    Grid4.Column(7).Locked = True
    Grid4.PageSetup.BlackAndWhite = True
    
    
    
    
    
    
    
    
    
    
    Grid4.PrintPreview
    Grid4.Rows = FINROW
  
End Sub
Sub cabeza()
Dim K As Integer
    Dim objReportTitle As FlexCell.ReportTitle
    
    Grid4.ReportTitles.Clear
    'Report Title 1
    For K = 1 To 5
        Set objReportTitle = New FlexCell.ReportTitle
        objReportTitle.text = DATOSEMPRESA(K)
        objReportTitle.Font.Name = "verdana"
        objReportTitle.Font.Size = 7
        objReportTitle.Font.Italic = True
        objReportTitle.PrintOnAllPages = True
        objReportTitle.color = RGB(128, 0, 0)
        objReportTitle.Align = cellLeft
        Grid4.ReportTitles.Add objReportTitle
    Next K
    If entregados.Value = 0 Then
    Set objReportTitle = New FlexCell.ReportTitle
    objReportTitle.text = "RECEPCION SERVICIO TECNICO "
    objReportTitle.Font.Name = "verdana"
    objReportTitle.Font.Size = 10
    objReportTitle.Font.Bold = True
    objReportTitle.PrintOnAllPages = True
    Grid4.ReportTitles.Add objReportTitle
    
    Set objReportTitle = New FlexCell.ReportTitle
    objReportTitle.text = " Nº : " + FOLIO.text
    objReportTitle.Font.Name = "verdana"
    objReportTitle.Font.Size = 10
    objReportTitle.Font.Bold = True
    objReportTitle.PrintOnAllPages = True
    Grid4.ReportTitles.Add objReportTitle
    
     Set objReportTitle = New FlexCell.ReportTitle
    objReportTitle.text = "FECHA RECEPCION : " + DESDE1.text + "-" + DESDE2.text + "-" + DESDE3.text
    objReportTitle.Font.Name = "verdana"
    objReportTitle.Font.Size = 8
    objReportTitle.Font.Bold = True
    objReportTitle.PrintOnAllPages = True
    Grid4.ReportTitles.Add objReportTitle
    Else
     Set objReportTitle = New FlexCell.ReportTitle
    objReportTitle.text = "RETIRO SERVICIO TECNICO "
    objReportTitle.Font.Name = "verdana"
    objReportTitle.Font.Size = 10
    objReportTitle.Font.Bold = True
    objReportTitle.PrintOnAllPages = True
    Grid4.ReportTitles.Add objReportTitle
    
     Set objReportTitle = New FlexCell.ReportTitle
    objReportTitle.text = " Nº : " + FOLIO.text
    objReportTitle.Font.Name = "verdana"
    objReportTitle.Font.Size = 10
    objReportTitle.Font.Bold = True
    objReportTitle.PrintOnAllPages = True
    Grid4.ReportTitles.Add objReportTitle
    
     Set objReportTitle = New FlexCell.ReportTitle
    objReportTitle.text = "FECHA RETIRO : " + DESDE1.text + "-" + DESDE2.text + "-" + DESDE3.text
    objReportTitle.Font.Name = "verdana"
    objReportTitle.Font.Size = 8
    objReportTitle.Font.Bold = True
    objReportTitle.PrintOnAllPages = True
    Grid4.ReportTitles.Add objReportTitle
    End If
    

    Set objReportTitle = New FlexCell.ReportTitle
    objReportTitle.text = "CLIENTE :" & "   " & rut_cliente.text + "-" + lbldv.Caption
    objReportTitle.Font.Name = "verdana"
    objReportTitle.Font.Size = 8
    objReportTitle.Font.Bold = True
    objReportTitle.PrintOnAllPages = True
    objReportTitle.Align = cellLeft
    Grid4.ReportTitles.Add objReportTitle
    
    Set objReportTitle = New FlexCell.ReportTitle
    objReportTitle.text = "NOMBRE  :" & "   " & nombre_cliente.text
    objReportTitle.Font.Name = "verdana"
    objReportTitle.Font.Size = 8
    objReportTitle.Font.Bold = True
    objReportTitle.PrintOnAllPages = True
    objReportTitle.Align = cellLeft
    Grid4.ReportTitles.Add objReportTitle
    
    Set objReportTitle = New FlexCell.ReportTitle
    objReportTitle.text = "TELEFONO :" & "  " & fono.text
    objReportTitle.Font.Name = "verdana"
    objReportTitle.Font.Size = 8
    objReportTitle.Font.Bold = True
    objReportTitle.PrintOnAllPages = True
    objReportTitle.Align = cellLeft
    Grid4.ReportTitles.Add objReportTitle
    
    Set objReportTitle = New FlexCell.ReportTitle
    objReportTitle.text = "CELULAR :" & " " & celular.text
    objReportTitle.Font.Name = "verdana"
    objReportTitle.Font.Size = 8
    objReportTitle.Font.Bold = True
    objReportTitle.PrintOnAllPages = True
    objReportTitle.Align = cellLeft
    Grid4.ReportTitles.Add objReportTitle
    
    If entregados.Value = 1 Then
'    Set objReportTitle = New FlexCell.ReportTitle
'    objReportTitle.text = "FECHA RETIRO :" & " " & hasta1.text + "-" + hasta2.text + "-" + hasta3.text
'    objReportTitle.Font.Name = "verdana"
'    objReportTitle.Font.Size = 8
'    objReportTitle.Font.Bold = True
'    objReportTitle.PrintOnAllPages = True
'    objReportTitle.Align = cellLeft
'    Grid4.ReportTitles.Add objReportTitle
    
    Set objReportTitle = New FlexCell.ReportTitle
    objReportTitle.text = "RETIRADO POR :" & " " & nombre2.text
    objReportTitle.Font.Name = "verdana"
    objReportTitle.Font.Size = 8
    objReportTitle.Font.Bold = True
    objReportTitle.PrintOnAllPages = True
    objReportTitle.Align = cellLeft
    Grid4.ReportTitles.Add objReportTitle
    End If
    
    
    Grid4.PageSetup.Footer = "Pág &P de &N" & vbCrLf & "Fecha: &D"
    Grid4.PageSetup.FooterAlignment = cellRight
    Grid4.PageSetup.FooterFont.Name = "Verdana"
    Grid4.PageSetup.FooterFont.Size = 7
    
    
    
    
    With Grid4.PageSetup
        .HeaderFont.Size = 6
        '.Header = "                                                                                                                   PAGINAS &P/&N EMITIDO:&D USUARIO " + USUARIOSISTEMA
        .HeaderAlignment = cellCenter
        .HeaderFont.Name = "Verdana"
        .HeaderMargin = 4
    End With
End Sub

Sub planillaproveedor()
Dim K As Integer
    Rem DATOS DE LA COLUMNA
    Grid4.DefaultFont.Size = 8
    Grid4.DefaultFont.Bold = False
    
    
    formatogrilla(1, 1) = ""
    formatogrilla(1, 2) = ""
    formatogrilla(1, 3) = ""
    formatogrilla(1, 4) = ""
    formatogrilla(1, 5) = ""
    formatogrilla(1, 6) = ""
    formatogrilla(1, 7) = ""
    
    Rem LARGO DE LOS DATOS
    formatogrilla(2, 1) = "13"
    formatogrilla(2, 2) = "30"
    formatogrilla(2, 3) = "8"
    formatogrilla(2, 4) = "8"
    formatogrilla(2, 5) = "8"
    formatogrilla(2, 6) = "8"
    formatogrilla(2, 7) = "8"
    
    Rem TIPO DE DATOS C=CEROS,S=STRING,N=NUMERICO,D=DATE,
    formatogrilla(3, 1) = "S"
    formatogrilla(3, 2) = "S"
    formatogrilla(3, 3) = "N"
    formatogrilla(3, 4) = "N"
    formatogrilla(3, 5) = "N"
    formatogrilla(3, 6) = "N"
    formatogrilla(3, 7) = "N"

    Rem FORMATO GRILLA
    formatogrilla(4, 1) = ""
    formatogrilla(4, 2) = ""
    formatogrilla(4, 3) = "###,##0.0"
    formatogrilla(4, 4) = "###,##0.0"
    formatogrilla(4, 5) = "###,##0.0"
    formatogrilla(4, 6) = "###,##0.0"
    formatogrilla(4, 7) = "###,##0.0"
    
    Rem LOCCKED
    formatogrilla(5, 1) = "TRUE"
    formatogrilla(5, 2) = "TRUE"
    formatogrilla(5, 3) = "FALSE"
    formatogrilla(5, 4) = "TRUE"
    formatogrilla(5, 5) = "TRUE"
    formatogrilla(5, 6) = "FALSE"
    formatogrilla(5, 7) = "TRUE"
    
    Grid4.FixedRows = 1
    Grid4.Cols = 8
    Grid4.Rows = 1
    
    Grid4.AllowUserResizing = False
    Grid4.DisplayFocusRect = False
    Grid4.ExtendLastCol = True
    Grid4.BoldFixedCell = False
    Grid4.DrawMode = cellOwnerDraw
    Grid4.Appearance = Flat
    Grid4.ScrollBarStyle = Flat
    Grid4.FixedRowColStyle = Flat
    Grid4.BackColorFixed = RGB(90, 158, 214)
    Grid4.BackColorFixedSel = RGB(110, 190, 230)
    Grid4.BackColorBkg = RGB(90, 158, 214)
    Grid4.BackColorScrollBar = RGB(231, 235, 247)
    Grid4.BackColor1 = RGB(231, 235, 247)
    Grid4.BackColor2 = RGB(239, 243, 255)
    Grid4.GridColor = RGB(148, 190, 231)
    For K = 1 To Grid4.Cols - 1
        Grid4.Cell(0, K).text = formatogrilla(1, K)
        'Grid4.Cell(1, k).text = FORMATOGRILLA(8, k)
        Grid4.Column(K).Width = Val(formatogrilla(2, K)) * Grid4.Cell(0, K).Font.Size
        Grid4.Column(K).MaxLength = Val(formatogrilla(2, K))
        Grid4.Column(K).FormatString = formatogrilla(4, K)
        Grid4.Column(K).Locked = formatogrilla(5, K)
        If formatogrilla(3, K) = "N" Then Grid4.Column(K).Alignment = cellRightCenter
    Next K
    Grid4.Column(0).Width = 0
    Grid4.Range(0, 0, 0, Grid4.Cols - 1).Alignment = cellCenterCenter
    Grid4.Column(3).UserSortIndicator = cellSortIndicatorDescending
    Rem Grid4.Enabled = False
End Sub

Public Sub imprime_guialarga()
    Dim SS As String
    Dim i As Integer
    Dim K As Integer
    Dim cad As String
    Dim totalprod As String
    Dim Descuento As String
    Dim neto As String
    Dim piva As String
    Dim piha As String
    Dim total As String
    Dim lineas As Integer
    Dim fecha As String
    Dim vencimiento As String
    Dim vendedor As String
    Dim notapedido As String
    Dim nombre As String
    Dim rut As String
    Dim direccion As String
    Dim ciudad As String
    Dim comuna As String
    Dim giro As String
    Dim fono1 As String
    Dim o As Integer
    Dim dia As String
    Dim mes As String
    Dim ano As String
    Dim nvalor As Long
    Dim CODIGO As String
    Dim tiposDePago As String
    Dim razon As String
    Dim dife As Double
    Dim impuesto(10) As Double
    Dim TAZAS(10) As Double
    
    
    Grid1.Rows = 1
    Grid1.Cols = 6
    Grid1.Rows = 50
   
    Grid1.DefaultFont.Bold = False
    
    Grid1.Column(0).Width = 0
    Grid1.Column(1).Width = 150
    Grid1.Column(2).Width = 90
    Grid1.Column(3).Width = 260
    Grid1.Column(4).Width = 100
    Grid1.Column(5).Width = 150
    
    Grid1.Column(1).Locked = False
    Grid1.Column(2).Locked = False
    Grid1.Column(3).Locked = False
    Grid1.Column(4).Locked = False
    Grid1.Column(5).Locked = False
    Grid1.PageSetup.BlackAndWhite = True
    
    Grid1.Column(1).Alignment = cellRightCenter
    Grid1.Column(2).Alignment = cellCenterCenter
    Grid1.Column(3).Alignment = cellLeftCenter '/**/
    Grid1.Column(4).Alignment = cellRightCenter
    Grid1.Column(5).Alignment = cellRightCenter

    'Grid1.Column(7).Alignment = cellRightCenter
    
    Grid1.PageSetup.PrintGridlines = False
    Grid1.AutoRedraw = False
    'CABEZA
    
       nombre = nombreserviciotecnico
       rut = rutserviciotecnico
       direccion = direccionserviciotecnico
       ciudad = ciudadserviciotecnico
       comuna = comunaserviciotecnico
       giro = ""
       razon = nombreserviciotecnico
        
        
    
    Grid1.Cell(5, 5).text = nguia.text
    Grid1.Range(5, 2, 5, 3).Merge
    Grid1.Range(5, 2, 5, 3).Alignment = cellCenterCenter
    Grid1.Cell(5, 2).text = leerNombreEmpresa(empresaActiva)
    
    'SEÑORES
    Grid1.Range(9, 2, 9, 3).Merge
    Grid1.Range(9, 2, 9, 3).Alignment = cellLeftCenter
    Grid1.Cell(9, 2).text = razon
    
    ' fecha
    fecha = DESDE1.text + "-" + DESDE2.text + "-" + DESDE3.text
    Grid1.Cell(9, 5).text = fecha
    
    
    'DIRECCION
    Grid1.Range(11, 2, 11, 3).Merge
    Grid1.Range(11, 2, 11, 3).Alignment = cellLeftCenter
    Grid1.Cell(11, 2).text = direccion
    
    'RUT
    'Grid1.Range(9, 2, 9, 3).Merge
    Grid1.Cell(11, 5).Alignment = cellLeftCenter
    Grid1.Cell(11, 5).text = "     " + Format(Left(rut, 9), "###,###,###") & "-" & Right(rut, 1)
    
    'GIRO
    Grid1.Range(13, 2, 13, 3).Merge
    Grid1.Range(13, 2, 13, 3).Alignment = cellLeftCenter
    Grid1.Cell(13, 2).text = razon
 
 
        lineas = 19
 
            lineas = lineas + 1
            Grid1.Cell(lineas, 1).text = codigobarra.text
            Grid1.Cell(lineas, 2).text = "1"
            Grid1.Cell(lineas, 3).text = descripcion.text
            Grid1.Cell(lineas, 4).text = "1"
            Grid1.Cell(lineas, 5).text = "1"
            
            If serie.text <> "" Then
            lineas = lineas + 1
            Grid1.Cell(lineas, 3).text = "Serie :" & serie.text
            End If
            
            If marca.text <> "" Then
            lineas = lineas + 1
            Grid1.Cell(lineas, 3).text = "Marca :" & marca.text
            End If
            
            If articulo.text <> "" Then
            lineas = lineas + 1
            Grid1.Cell(lineas, 3).text = "Articulo :" & articulo.text
            End If
            
            If falla.text <> "" Then
            lineas = lineas + 1
            Grid1.Cell(lineas, 3).text = "Falla :" & falla.text
            End If
            
            lineas = lineas + 2
            Grid1.Cell(lineas, 3).text = "Cliente :" & nombre_cliente.text
            
            If fono.text <> "" Then
               lineas = lineas + 1
               Grid1.Cell(lineas, 3).text = "Fono :" & fono.text
            End If
            
        Grid1.Range(47, 1, 47, 5).Merge
        Grid1.Range(47, 1, 47, 5).Alignment = cellCenterCenter
        Grid1.Range(47, 1, 47, 5).FontBold = True
        
        Grid1.Cell(47, 1).text = "NO CONSTITUYE VENTA, SOLO TRASLADO DE MERCADERIA."

    Grid1.AutoRedraw = True
    Grid1.Refresh
    
    Grid1.PageSetup.LeftMargin = 0.25
    Grid1.PageSetup.RightMargin = 0
    Grid1.PageSetup.TopMargin = 2.5
    Grid1.PageSetup.BottomMargin = 0
 
    Grid1.PageSetup.PrintGridlines = False
    'Grid1.DirectPrint
    Grid1.PrintPreview
End Sub


  
    Public Sub imprime_guiacorta()
    Dim SS As String
    Dim i As Integer
    Dim K As Integer
    Dim cad As String
    Dim totalprod As String
    Dim Descuento As String
    Dim neto As String
    Dim piva As String
    Dim piha As String
    Dim total As String
    Dim lineas As Integer
    Dim fecha As String
    Dim vencimiento As String
    Dim vendedor As String
    Dim notapedido As String
    Dim nombre As String
    Dim rut As String
    Dim direccion As String
    Dim ciudad As String
    Dim comuna As String
    Dim giro As String
    
    Dim o As Integer
    Dim dia As String
    Dim mes As String
    Dim ano As String
    Dim nvalor As Long
    Dim CODIGO As String
    Dim tiposDePago As String
    
    
    
    Grid1.Rows = 2
    Grid1.Cols = 7
    Grid1.Rows = 40
    Grid1.DefaultFont.Name = "Arial"
    Grid1.DefaultFont.Size = 8
    Grid1.DefaultFont.Bold = False
    
    Grid1.Column(0).Width = 0
    Grid1.Column(1).Width = 120
    Grid1.Column(2).Width = 110
    Grid1.Column(3).Width = 180
    Grid1.Column(4).Width = 105
    Grid1.Column(5).Width = 100
    Grid1.Column(6).Width = 130
    'Grid1.Column(7).Width = 125
    
    Grid1.Column(1).Alignment = cellRightCenter
    Grid1.Column(2).Alignment = cellCenterCenter
    Grid1.Column(3).Alignment = cellLeftCenter '/**/
    Grid1.Column(4).Alignment = cellLeftCenter '/**/
    Grid1.Column(5).Alignment = cellRightCenter
    Grid1.Column(6).Alignment = cellRightCenter
    'Grid1.Column(7).Alignment = cellRightCenter
    
    Grid1.DefaultRowHeight = 13
    
    Grid1.PageSetup.PrintGridlines = False
    Grid1.AutoRedraw = False
    
 
  
  
       nombre = nombreserviciotecnico
       rut = rutserviciotecnico
       direccion = direccionserviciotecnico
       ciudad = ciudadserviciotecnico
       comuna = comunaserviciotecnico
       giro = ""
       
       
        fecha = DESDE1.text + "-" + DESDE2.text + "-" + DESDE3.text
        fecha = DESDE1.text
        fecha = fecha & "                         "
        fecha = fecha & MonthName(DESDE2.text)
        fecha = fecha & "                                        "
        fecha = fecha & DESDE3.text
       

        
        
        
        Grid1.Cell(5, 4).Alignment = cellRightCenter
        Grid1.Cell(5, 6).text = nguia.text
        
   
            lineas = 21
            Grid1.Cell(lineas, 1).text = "1"
            Grid1.Cell(lineas, 2).text = codigobarra.text
            Grid1.Cell(lineas, 3).text = descripcion.text
            Grid1.Cell(lineas, 4).text = "1"
            Grid1.Cell(lineas, 5).text = "1"
            
            If serie.text <> "" Then
            lineas = lineas + 1
            Grid1.Cell(lineas, 3).text = "Serie :" & serie.text
            End If
            
            If marca.text <> "" Then
            lineas = lineas + 1
            Grid1.Cell(lineas, 3).text = "Marca :" & marca.text
            End If
            
            If articulo.text <> "" Then
            lineas = lineas + 1
            Grid1.Cell(lineas, 3).text = "Articulo :" & articulo.text
            End If
            
            If falla.text <> "" Then
            lineas = lineas + 1
            Grid1.Cell(lineas, 3).text = "Falla :" & falla.text
            End If
            
            lineas = lineas + 2
            Grid1.Cell(lineas, 3).text = "Cliente :" & nombre_cliente.text
            
            If fono.text <> "" Then
               lineas = lineas + 1
               Grid1.Cell(lineas, 3).text = "Fono :" & fono.text
            End If
            
            
            lineas = lineas + 5
            Grid1.Range(lineas, 1, lineas, 5).Merge
            Grid1.Range(lineas, 1, lineas, 5).Alignment = cellCenterCenter
            Grid1.Range(lineas, 1, lineas, 5).FontBold = True
            Grid1.Cell(lineas, 1).text = "NO CONSTITUYE VENTA, SOLO TRASLADO DE MERCADERIA."
    
    
    
    
    
    
    
    
        
    Grid1.Range(4, 2, 4, 3).Merge
    Grid1.Range(4, 2, 4, 3).Alignment = cellCenterCenter
    Grid1.Cell(4, 2).text = leerNombreEmpresa(empresaActiva)
    
    'FECHA
    Grid1.Range(8, 2, 8, 3).Merge
    Grid1.Range(8, 2, 8, 3).Alignment = cellLeftCenter
    Grid1.Cell(8, 2).text = fecha
    
    'SEÑORES
    Grid1.Range(11, 2, 11, 3).Merge
    Grid1.Range(11, 2, 11, 3).Alignment = cellLeftCenter
    Grid1.Cell(11, 2).text = nombre
    'RUT
    'Grid1.Range(11, 2, 11, 3).Merge
    Grid1.Cell(11, 6).Alignment = cellLeftCenter
    Grid1.Cell(11, 6).text = rut
    'DIRECCION
    Grid1.Range(12, 2, 12, 3).Merge
    Grid1.Range(12, 2, 12, 3).Alignment = cellLeftCenter
    Grid1.Cell(12, 2).text = direccion
    'CIUDAD
    'Grid1.Range(13, 2, 13, 3).Merge
    Grid1.Cell(12, 6).Alignment = cellLeftCenter
    Grid1.Cell(12, 6).text = ciudad
    'GIRO
    Grid1.Range(13, 2, 15, 3).Merge
    Grid1.Range(13, 2, 15, 3).Alignment = cellLeftCenter
    Grid1.Cell(13, 2).text = giro
    'COMUNA
    'Grid1.Range(15, 5, 15, 6).Merge
    Grid1.Cell(13, 6).Alignment = cellLeftCenter
    Grid1.Cell(13, 6).text = comuna
    
    
    For i = Grid1.Rows To 40
        Grid1.AddItem ""
    Next i
    
    
    Grid1.AddItem ""
    Grid1.AddItem ""
     
     
    Grid1.AutoRedraw = True
    Grid1.Refresh
    
    Grid1.PageSetup.LeftMargin = 0.25
    Grid1.PageSetup.RightMargin = 0
    Grid1.PageSetup.TopMargin = 2.5
    Grid1.PageSetup.BottomMargin = 0
    
    For i = 1 To Grid1.PageSetup.PaperSizes.Count
        If UCase(Grid1.PageSetup.PaperSizes.Item(i).PaperName) = "CARTA" Then
            Grid1.PageSetup.PaperSize = Grid1.PageSetup.PaperSizes.Item(i).Kind
            Exit For
        End If
    Next i
    
    'Grid1.DirectPrint
    Grid1.PrintPreview
End Sub










Sub leerTecnico(FILTRO)

Dim op As Integer
Dim CAMPOS(5, 5) As String
    CAMPOS(0, 0) = "nombre"
    CAMPOS(1, 0) = "direccion"
    CAMPOS(2, 0) = "comuna"
    CAMPOS(3, 0) = "ciudad"
    CAMPOS(4, 0) = "fono"
    CAMPOS(5, 0) = ""
    CAMPOS(0, 2) = "sv_maestroserviciotecnico"
    condicion = "rut='" & FILTRO & "'"
    op = 5
    Set sqlventas.conexion = ventas
    sqlventas.response = CAMPOS
    Call sqlventas.sqlventas(op, condicion)
    If sqlventas.Status = 4 Then
        ruttecnico.SetFocus
    Else
    rutserviciotecnico = FILTRO
    nombreserviciotecnico = sqlventas.response(0, 3)
    direccionserviciotecnico = sqlventas.response(1, 3)
    comunaserviciotecnico = sqlventas.response(2, 3)
    ciudadserviciotecnico = sqlventas.response(3, 3)
    fonoserviciotecnico = sqlventas.response(4, 3)
    lblnombretecnico.Caption = sqlventas.response(0, 3)
   
   End If
End Sub

Private Sub CARGAGRILLA()
        Dim col As Integer
        Dim row As Integer
        Dim K As Integer
        col = 8
        row = 2
        Dim formatogrilla(10, 10) As String
        Rem DATOS DE LA COLUMNA
        formatogrilla(1, 0) = "LN"
        
        formatogrilla(1, 1) = "CODIGO"
        formatogrilla(1, 2) = "DESCRIPCION"
        formatogrilla(1, 3) = "STOCK"
        formatogrilla(1, 4) = "CANTIDAD"
        formatogrilla(1, 5) = "PRECIO"
        formatogrilla(1, 6) = "DESC %"
        formatogrilla(1, 7) = "TOTAL"
        
        Rem LARGO DE LOS DATOS
        formatogrilla(2, 1) = "13"
        formatogrilla(2, 2) = "50"
        formatogrilla(2, 3) = "5"
        formatogrilla(2, 4) = "7"
        formatogrilla(2, 5) = "8"
        formatogrilla(2, 6) = "8"
        formatogrilla(2, 7) = "9"
        
        Rem TIPO DE DATOS C=CEROS,S=STRING,N=NUMERICO,D=DATE,
        formatogrilla(3, 1) = "C"
        formatogrilla(3, 2) = "S"
        formatogrilla(3, 3) = "N"
        formatogrilla(3, 4) = "N"
        formatogrilla(3, 5) = "N"
        formatogrilla(3, 6) = "N"
        formatogrilla(3, 7) = "N"
            
        Rem FORMATO GRILLA
        formatogrilla(4, 1) = "0000000000000"
        formatogrilla(4, 2) = ""
        formatogrilla(4, 3) = "###,##0.00"
        formatogrilla(4, 4) = "###,##0.00"
        formatogrilla(4, 5) = "$ ###,###,##0.0"
        formatogrilla(4, 6) = "##0.0"
        formatogrilla(4, 7) = "$ ###,###,##0"
        
        Rem LOCCKED
        formatogrilla(5, 1) = "FALSE"
        formatogrilla(5, 2) = "TRUE"
        formatogrilla(5, 3) = "TRUE"
        formatogrilla(5, 4) = "FALSE"
        formatogrilla(5, 5) = "FALSE"
        formatogrilla(5, 6) = "FALSE"
        formatogrilla(5, 7) = "TRUE"
        
        Rem VALOR MINIMO
        formatogrilla(6, 1) = ""
        formatogrilla(6, 2) = ""
        formatogrilla(6, 3) = ""
        formatogrilla(6, 4) = "1"
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
        Rem ANCHO DE LA COLUMNA
        formatogrilla(8, 1) = "100"
        formatogrilla(8, 2) = "300"
        formatogrilla(8, 3) = "75"
        formatogrilla(8, 4) = "75"
        formatogrilla(8, 5) = "85"
        formatogrilla(8, 6) = "85"
        formatogrilla(8, 7) = "85"
        
        Grid1.Cols = col
        Grid1.Rows = row
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
        
        Rem Asigna Valores a la Grilla
        Grid1.Cell(0, 0).text = formatogrilla(1, 0)
        For K = 1 To col - 1
            Grid1.Cell(0, K).text = formatogrilla(1, K)
            Grid1.Column(K).Width = Val(formatogrilla(8, K))
            Grid1.Column(K).MaxLength = Val(formatogrilla(2, K))
            Grid1.Column(K).FormatString = formatogrilla(4, K)
            Grid1.Column(K).Locked = formatogrilla(5, K)
            If formatogrilla(3, K) = "S" Then
                Grid1.Column(K).Alignment = cellLeftCenter
            Else
                Grid1.Column(K).Alignment = cellRightCenter
            End If
            Grid1.Cell(0, K).Alignment = cellCenterCenter
        Next K
    End Sub
Public Sub cargar_servicioafuera()
    Call leergarantias(FOLIO.text)
End Sub

