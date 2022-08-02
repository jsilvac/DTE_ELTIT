VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash9e.ocx"
Object = "{4777436C-EB8C-4596-98A8-EBCF98683969}#1.0#0"; "Flexcell.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{ADBBAED6-B16F-40DC-80DF-B44910CBA76C}#1.0#0"; "Frame-Xp.ocx"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form cotiza01 
   ClientHeight    =   9930
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15090
   LinkTopic       =   "Form1"
   ScaleHeight     =   9930
   ScaleWidth      =   15090
   StartUpPosition =   2  'CenterScreen
   Begin MSAdodcLib.Adodc data 
      Height          =   330
      Left            =   7200
      Top             =   315
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
   Begin XPFrame.FrameXp FrameXp1 
      Height          =   11295
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   15135
      _ExtentX        =   26696
      _ExtentY        =   19923
      BackColor       =   16744576
      Caption         =   ""
      CaptionEstilo3D =   1
      BackColor       =   16744576
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   1
      Begin VB.TextBox pivote 
         Height          =   285
         Left            =   0
         MaxLength       =   13
         TabIndex        =   53
         Text            =   "Text1"
         Top             =   3285
         Visible         =   0   'False
         Width           =   1140
      End
      Begin XPFrame.FrameXp FrameXp3 
         Height          =   2580
         Left            =   135
         TabIndex        =   25
         Top             =   675
         Width           =   8070
         _ExtentX        =   14235
         _ExtentY        =   4551
         BackColor       =   16761024
         Caption         =   "Datos Cliente"
         CaptionEstilo3D =   1
         BackColor       =   16761024
         ForeColor       =   65535
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Begin VB.CheckBox Check1 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
            Caption         =   "Autorizada"
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
            Height          =   285
            Left            =   5895
            TabIndex        =   52
            Top             =   2250
            Width           =   1995
         End
         Begin VB.TextBox DATO7 
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
            Left            =   6075
            MaxLength       =   9
            TabIndex        =   49
            Tag             =   "proveedor"
            Top             =   765
            Width           =   1455
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
            Left            =   1485
            MaxLength       =   10
            TabIndex        =   1
            Tag             =   "proveedor"
            Top             =   360
            Width           =   1455
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
            Left            =   5850
            MaxLength       =   2
            TabIndex        =   2
            Tag             =   "proveedor"
            Top             =   360
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
            Left            =   6210
            MaxLength       =   2
            TabIndex        =   3
            Tag             =   "proveedor"
            Top             =   360
            Width           =   375
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
            Left            =   6570
            MaxLength       =   4
            TabIndex        =   4
            Tag             =   "proveedor"
            Top             =   360
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
            Left            =   1485
            MaxLength       =   9
            TabIndex        =   5
            Tag             =   "proveedor"
            Top             =   765
            Width           =   1455
         End
         Begin VB.TextBox dato27 
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
            Left            =   10635
            MaxLength       =   10
            TabIndex        =   33
            Tag             =   "proveedor"
            Top             =   0
            Width           =   1455
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "VENDEDOR"
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
            Left            =   4590
            TabIndex        =   51
            Top             =   765
            Width           =   1275
         End
         Begin VB.Label LBLDV2 
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
            Height          =   285
            Left            =   7560
            TabIndex        =   50
            Top             =   765
            Width           =   330
         End
         Begin VB.Label lbldv 
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
            Height          =   285
            Left            =   2970
            TabIndex        =   45
            Top             =   765
            Width           =   330
         End
         Begin VB.Label LBLDIRECCION 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFC0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   330
            Left            =   1485
            TabIndex        =   37
            Top             =   1530
            Width           =   6450
         End
         Begin VB.Label LBLCIUDAD 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFC0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   330
            Left            =   1485
            TabIndex        =   36
            Top             =   1890
            Width           =   2940
         End
         Begin VB.Label LBLFONO 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFC0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   330
            Left            =   5850
            TabIndex        =   35
            Top             =   1890
            Width           =   2085
         End
         Begin VB.Label LBLNOMBRE 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFC0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   330
            Left            =   1485
            TabIndex        =   34
            Top             =   1170
            Width           =   6450
         End
         Begin VB.Label Label14 
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "FONO"
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
            Left            =   4500
            TabIndex        =   32
            Top             =   1890
            Width           =   1275
         End
         Begin VB.Label Label13 
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "NOMBRE"
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
            Left            =   135
            TabIndex        =   31
            Top             =   1170
            Width           =   1275
         End
         Begin VB.Label Label12 
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "CIUDAD"
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
            Left            =   135
            TabIndex        =   30
            Top             =   1890
            Width           =   1275
         End
         Begin VB.Label Label11 
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "DIRECCION"
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
            Left            =   135
            TabIndex        =   29
            Top             =   1530
            Width           =   1275
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "RUT"
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
            Left            =   135
            TabIndex        =   28
            Top             =   765
            Width           =   1275
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
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
            Height          =   285
            Left            =   4500
            TabIndex        =   27
            Top             =   360
            Width           =   1275
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
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
            Height          =   285
            Left            =   135
            TabIndex        =   26
            Top             =   360
            Width           =   1275
         End
      End
      Begin VB.Timer Timer2 
         Interval        =   1000
         Left            =   360
         Top             =   120
      End
      Begin XPFrame.FrameXp BARRAELIMINA 
         Height          =   285
         Left            =   180
         TabIndex        =   6
         Top             =   3330
         Width           =   14865
         _ExtentX        =   26220
         _ExtentY        =   503
         BackColor       =   0
         Caption         =   " MODO DE ELIMINACION DE PRODUCTOS SELECCIONE PRODUCTO Y PRESIONE  TECLA (SUPR) - (INICIO) PARA CONTINUAR"
         CaptionEstilo3D =   1
         BackColor       =   0
         ForeColor       =   65535
         BordeColor      =   0
         ColorBarraArriba=   0
         ColorBarraAbajo =   0
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
      End
      Begin XPFrame.FrameXp FrameXp5 
         Height          =   1500
         Left            =   10530
         TabIndex        =   7
         Top             =   8370
         Width           =   4335
         _ExtentX        =   7646
         _ExtentY        =   2646
         BackColor       =   16744576
         Caption         =   "TOTAL COTIZACION"
         CaptionEstilo3D =   1
         BackColor       =   16744576
         ColorBarraArriba=   12648384
         ColorBarraAbajo =   32768
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
         ColorTextShadow =   0
         Begin VB.Label Label9 
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Total"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   135
            TabIndex        =   42
            Top             =   1035
            Width           =   1440
         End
         Begin VB.Label Label8 
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Iva"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   135
            TabIndex        =   41
            Top             =   675
            Width           =   1440
         End
         Begin VB.Label Label7 
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Neto"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   135
            TabIndex        =   40
            Top             =   315
            Width           =   1440
         End
         Begin VB.Label lbltotal 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " 0"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000C000&
            Height          =   315
            Left            =   1665
            TabIndex        =   39
            Top             =   1035
            Width           =   2520
         End
         Begin VB.Label lbliva 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000C000&
            Height          =   315
            Left            =   1665
            TabIndex        =   38
            Top             =   675
            Width           =   2520
         End
         Begin VB.Label lblneto 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " 0"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000C000&
            Height          =   300
            Left            =   1665
            TabIndex        =   8
            Top             =   315
            Width           =   2520
         End
      End
      Begin XPFrame.FrameXp FrameXp6 
         Height          =   1605
         Left            =   8415
         TabIndex        =   9
         Top             =   1800
         Width           =   6615
         _ExtentX        =   11668
         _ExtentY        =   2831
         BackColor       =   16744576
         Caption         =   "Stock de Productos"
         CaptionEstilo3D =   1
         BackColor       =   16744576
         ColorBarraArriba=   16777152
         ColorBarraAbajo =   32768
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
         Begin FlexCell.Grid Grid2 
            Height          =   1335
            Left            =   60
            TabIndex        =   10
            Top             =   240
            Width           =   6540
            _ExtentX        =   11536
            _ExtentY        =   2355
            Cols            =   5
            DefaultFontSize =   8.25
            Rows            =   30
         End
      End
      Begin MSAdodcLib.Adodc rollo 
         Height          =   330
         Left            =   5400
         Top             =   11160
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
      Begin MSCommLib.MSComm MSComm2 
         Left            =   6360
         Top             =   120
         _ExtentX        =   1005
         _ExtentY        =   1005
         _Version        =   393216
         DTREnable       =   -1  'True
      End
      Begin XPFrame.FrameXp framePago 
         Height          =   30
         Left            =   120
         TabIndex        =   11
         Top             =   3240
         Visible         =   0   'False
         Width           =   6615
         _ExtentX        =   11668
         _ExtentY        =   53
         BackColor       =   16744576
         Caption         =   "CANCELADO CON"
         CaptionEstilo3D =   1
         BackColor       =   16744576
         ColorBarraArriba=   12648384
         ColorBarraAbajo =   32768
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
         ColorTextShadow =   0
         Begin VB.TextBox txtCancelado 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   48
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000C000&
            Height          =   1095
            Left            =   120
            Locked          =   -1  'True
            TabIndex        =   12
            Text            =   "$ 0"
            Top             =   360
            Width           =   6375
         End
         Begin XPFrame.FrameXp FrameXp7 
            Height          =   30
            Left            =   0
            TabIndex        =   13
            Top             =   3105
            Width           =   6615
            _ExtentX        =   11668
            _ExtentY        =   53
            BackColor       =   16744576
            Caption         =   "VUELTO"
            CaptionEstilo3D =   1
            BackColor       =   16744576
            ColorBarraArriba=   12648384
            ColorBarraAbajo =   32768
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
            ColorTextShadow =   0
            Begin VB.Label lblVuelto 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00000000&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "$ 0"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   48
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H0000C000&
               Height          =   1125
               Left            =   120
               TabIndex        =   14
               Top             =   360
               Width           =   6375
            End
         End
      End
      Begin XPFrame.FrameXp FrameXp2 
         Height          =   900
         Left            =   8415
         TabIndex        =   15
         Top             =   855
         Width           =   6615
         _ExtentX        =   11668
         _ExtentY        =   1588
         BackColor       =   16744576
         Caption         =   "VENDEDOR (A)"
         CaptionEstilo3D =   1
         BackColor       =   16744576
         ColorBarraArriba=   12648384
         ColorBarraAbajo =   32768
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
         Begin VB.Label lblvendedor 
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000C000&
            Height          =   375
            Left            =   120
            TabIndex        =   16
            Top             =   360
            Width           =   6375
         End
      End
      Begin XPFrame.FrameXp frameDescripcion 
         Height          =   135
         Left            =   240
         TabIndex        =   17
         Top             =   11400
         Visible         =   0   'False
         Width           =   135
         _ExtentX        =   238
         _ExtentY        =   238
         BackColor       =   16744576
         Caption         =   "DESCRIPCION DEL PRODUCTO"
         CaptionEstilo3D =   1
         BackColor       =   16744576
         ColorBarraArriba=   12648384
         ColorBarraAbajo =   32768
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
         ColorTextShadow =   0
         Begin XPFrame.FrameXp FrameXp8 
            Height          =   1575
            Left            =   0
            TabIndex        =   18
            Top             =   1560
            Width           =   6615
            _ExtentX        =   11668
            _ExtentY        =   2778
            BackColor       =   16744576
            Caption         =   "PRECIO"
            CaptionEstilo3D =   1
            BackColor       =   16744576
            ColorBarraArriba=   12648384
            ColorBarraAbajo =   32768
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
            ColorTextShadow =   0
            Begin VB.Label LBLPRECIO 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H000000FF&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "$ 0"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   48
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H0000FFFF&
               Height          =   1125
               Index           =   0
               Left            =   120
               TabIndex        =   19
               Top             =   360
               Width           =   6375
            End
         End
         Begin VB.Label lbldescripcion 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H000000FF&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   21.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   1125
            Index           =   0
            Left            =   120
            TabIndex        =   20
            Top             =   360
            Width           =   6375
         End
      End
      Begin XPFrame.FrameXp FrameXp4 
         Height          =   7740
         Left            =   0
         TabIndex        =   21
         Top             =   3330
         Width           =   15105
         _ExtentX        =   26644
         _ExtentY        =   13653
         BackColor       =   16744576
         Caption         =   "PRODUCTOS EN LA COMPRA"
         CaptionEstilo3D =   1
         BackColor       =   16744576
         ColorBarraArriba=   12648384
         ColorBarraAbajo =   32768
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
         Begin FlexCell.Grid Grid4 
            Height          =   330
            Left            =   8595
            TabIndex        =   48
            Top             =   6030
            Visible         =   0   'False
            Width           =   645
            _ExtentX        =   1138
            _ExtentY        =   582
            Cols            =   5
            DefaultFontSize =   8.25
            Rows            =   30
         End
         Begin VB.CommandButton grabar2 
            BackColor       =   &H00FFC0C0&
            Caption         =   "GRABAR"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Left            =   7650
            Style           =   1  'Graphical
            TabIndex        =   46
            Top             =   5085
            Width           =   1995
         End
         Begin XPFrame.FrameXp FrameXp9 
            Height          =   1995
            Left            =   9675
            TabIndex        =   43
            Top             =   4995
            Visible         =   0   'False
            Width           =   375
            _ExtentX        =   661
            _ExtentY        =   3519
            BackColor       =   49344
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
            Begin FlexCell.Grid grid3 
               Height          =   1620
               Left            =   180
               TabIndex        =   44
               Top             =   270
               Visible         =   0   'False
               Width           =   6330
               _ExtentX        =   11165
               _ExtentY        =   2858
               Cols            =   5
               DefaultFontSize =   8.25
               Rows            =   30
            End
         End
         Begin FlexCell.Grid GRID1 
            Height          =   4500
            Left            =   0
            TabIndex        =   22
            TabStop         =   0   'False
            Top             =   360
            Width           =   15090
            _ExtentX        =   26617
            _ExtentY        =   7938
            Cols            =   5
            DefaultFontSize =   8.25
            DefaultFontBold =   -1  'True
            DisplayRowIndex =   -1  'True
            Rows            =   1
            MultiSelect     =   0   'False
         End
         Begin ShockwaveFlashObjectsCtl.ShockwaveFlash opciones 
            Height          =   1650
            Left            =   270
            TabIndex        =   47
            Top             =   4995
            Width           =   7395
            _cx             =   13044
            _cy             =   2910
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
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "CANTIDAD"
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
         Left            =   9360
         TabIndex        =   24
         Top             =   10350
         Width           =   1815
      End
      Begin VB.Label HORA 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   225
         Width           =   3135
      End
   End
End
Attribute VB_Name = "cotiza01"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Private formatogrilla(20, 20) As String
    Private existe As Boolean
    Public comando As Integer
    Private color(1) As Variant
    Private numeros(20) As String
    Private ELIMINAINDIVIDUAL As Boolean
    Private FECHA As String
    Private modificando As Boolean
Private Sub Check1_Click()
Call autorizar(Check1.Value)
End Sub

Private Sub dato6_GotFocus()
Call cargatexto(dato6)

End Sub

Private Sub dato6_KeyDown(KeyCode As Integer, Shift As Integer)
        If KeyCode = vbKeyF2 Then
            Call ayudaCliente(dato6, DATO7, lbldv)
        Else
            Call Flechas(KeyCode, dato5)
        End If

End Sub

Private Sub dato7_GotFocus()
Call cargatexto(DATO7)

End Sub

Private Sub Form_Unload(Cancel As Integer)
 If modificando = True Then
  Call grabar2_Click
 End If
End Sub

Private Sub grabar2_Click()
If modificando = True Then
ELIMINAR
End If
If GRID1.Rows > 1 Then
grabar
End If
retorno

End Sub

Private Sub dato2_GotFocus()
If modificando = False Then
dato2.text = leerUltimoFoliocotizacion
End If

Call cargatexto(dato2)

End Sub

Private Sub dato2_KeyPress(KeyAscii As Integer)
KeyAscii = esNumero(KeyAscii)
If KeyAscii = 13 Then
dato2.text = ceros(dato2)
If leercotizacion = False Then
            If Verifica_Permiso(Me.Caption, "agrega") = True Then
                 dato3.SetFocus
            Else
                MsgBox mensaje_nopermiso, vbCritical + vbOKOnly, "Permiso Denegado"
                dato2.SelStart = 0
                dato2.SelLength = Len(dato2.text)
                dato2.SetFocus
            End If

Else
opciones.Visible = True
grabar2.Visible = False
GRID1.SelectionMode = cellSelectionByRow


End If


End If

End Sub

Private Sub dato6_KeyPress(KeyAscii As Integer)
KeyAscii = esNumero(KeyAscii)

If KeyAscii = 13 Then
    dato6.text = ceros(dato6)
    lbldv.Caption = rut(dato6)

            If leerNombreCliente(dato6.text + lbldv.Caption) <> "" Then
                LBLNOMBRE.Caption = leerNombreCliente(dato6.text + lbldv.Caption)
                LBLDIRECCION.Caption = leerDireccionCliente(dato6.text + lbldv.Caption, "0")
                LBLCIUDAD.Caption = leerCiudadCliente(dato6.text + lbldv.Caption, "0")
                GRID1.Rows = 2

                DATO7.SetFocus

            Else

                MsgBox ("CLIENTE NO ESTA CREADO ")

                dato6.SetFocus

            End If
End If

End Sub

Private Sub dato7_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then
DATO7.text = Format(DATO7.text, "000000000")
LBLDV2.Caption = rut(DATO7.text)
If leerNombreVendedor(DATO7.text + LBLDV2.Caption) <> "" Then
lblvendedor.Caption = leerNombreVendedor(DATO7.text + LBLDV2.Caption)
GRID1.Cell(1, 1).SetFocus
Else
MsgBox ("vendedor no esta creado")
DATO7.SetFocus
End If
End If
End Sub

Private Sub Form_Activate()
   
    
    BARRAELIMINA.Visible = False
    
End Sub

Private Sub Form_Load()
    color(0) = &HFFFF&  'amarillo
    color(1) = &HFF&          'rojo
    Call CargaGrillaGRID1(1, 12)
    Call CARGAGRILLAbodegas
'    If empresaactiva <> "" Then
'        caja = leeArchivo("CAJA", "C:\caja.txt")
'    End If
    Me.framePago.Visible = False
    Me.frameDescripcion.Visible = True
    FrameXp1.Caption = nombreempresa
    'Pelicula.Picture = LoadPicture(App.Path & "\fotoPeli.jpg")
    ''''''''''
    iva = 19 '
    ''''''''''
    opciones.Visible = False
    
    
End Sub

Sub CARGASTOCKBODEGAS(CODIGO)
    Dim a As Integer
    Dim resultados As rdoResultset
    Dim cSql As New rdoQuery
    Dim rut As String
    Dim saldo As Double

        Set cSql.ActiveConnection = gestionRubro
        cSql.sql = "SELECT local,bodega,ubicacion,stockactual "
        cSql.sql = cSql.sql + "FROM r_maestroproductos_stock_" & rubro & " "
        cSql.sql = cSql.sql + "WHERE año='" + Format(fechasistema, "yyyy") + "' AND codigo='" + CODIGO + "' order by bodega "
        cSql.Execute
        Grid2.Rows = 1
        Grid2.AutoRedraw = False
        If cSql.RowsAffected > 0 Then
            Set resultados = cSql.OpenResultset
            While Not resultados.EOF
                saldo = resultados(3)
                Grid2.AddItem leelocal(resultados(0)) & vbTab & leebodega(resultados(1)) & vbTab & resultados(2) & vbTab & saldo, False
                resultados.MoveNext
            Wend
            resultados.Close
            Set resultados = Nothing
        End If
        Grid2.AutoRedraw = True
        Grid2.Refresh
        Grid2.Enabled = True
End Sub
Function leelocal(CODIGO) As String

Dim op As Integer
Dim campos(3, 3) As String
    campos(0, 0) = "nombre"
    campos(1, 0) = ""
    campos(0, 2) = "g_maestroempresas"
    condicion = "codigo = '" & CODIGO & "'"
    op = 5
    Set sqlventas.conexion = gestion
    sqlventas.response = campos
    Call sqlventas.sqlventas(op, condicion)
    leelocal = sqlventas.response(0, 3)
End Function

Function leebodega(CODIGO) As String

Dim op As Integer
Dim campos(3, 3) As String
    campos(0, 0) = "nombre"
    campos(2, 0) = ""
    campos(0, 2) = "r_maestrobodegas_" & rubro
    condicion = "rubro = '" & rubro & "' AND codigobodega = '" & CODIGO & "'"
    op = 5
    Set sqlventas.conexion = gestionRubro
    sqlventas.response = campos
    Call sqlventas.sqlventas(op, condicion)
    leebodega = sqlventas.response(0, 3)
End Function
Sub CARGAGRILLAbodegas()
Dim K As Integer
    Rem DATOS DE LA COLUMNA
    formatogrilla(1, 1) = "LOCAL"
    formatogrilla(1, 2) = "BODEGA"
    formatogrilla(1, 3) = "UBICACION"
    formatogrilla(1, 4) = "STOCK "
    
    Rem LARGO DE LOS DATOS
    formatogrilla(2, 1) = "10"
    formatogrilla(2, 2) = "20"
    formatogrilla(2, 3) = "0"
    formatogrilla(2, 4) = "10"
    
    Rem TIPO DE DATOS C=CEROS,S=STRING,N=NUMERICO,D=DATE,
    formatogrilla(3, 1) = "S"
    formatogrilla(3, 2) = "S"
    formatogrilla(3, 3) = "S"
    formatogrilla(3, 4) = "N"
    formatogrilla(3, 5) = "N"
    
    Rem FORMATO GRILLA
    formatogrilla(4, 1) = ""
    formatogrilla(4, 2) = ""
    formatogrilla(4, 3) = ""
    formatogrilla(4, 4) = "#,###,##0.0"
    formatogrilla(4, 5) = "#,###,##0.0"
    
    Rem LOCCKED
    formatogrilla(5, 1) = "TRUE"
    formatogrilla(5, 2) = "TRUE"
    formatogrilla(5, 3) = "TRUE"
    formatogrilla(5, 4) = "TRUE"
    formatogrilla(5, 5) = "TRUE"
    Grid2.Cols = 5
    Grid2.Rows = 2
    
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
    For K = 1 To Grid2.Cols - 1
        Grid2.Cell(0, K).text = formatogrilla(1, K)
        Grid2.Column(K).Width = Val(formatogrilla(2, K)) * Grid2.DefaultFont.Size
        Grid2.Column(K).MaxLength = Val(formatogrilla(2, K))
        Grid2.Column(K).FormatString = formatogrilla(4, K)
        Grid2.Column(K).Locked = formatogrilla(5, K)
        If formatogrilla(3, K) = "N" Then Grid2.Column(K).Alignment = cellRightCenter
       
    Next K
    Grid2.Column(0).Width = 0
    Grid2.Range(0, 0, 0, Grid2.Cols - 1).Alignment = cellCenterCenter
    Grid2.Enabled = False
End Sub

'****************************************************************************
'Formato de la Grilla GRID1
'****************************************************************************
    Private Sub CargaGrillaGRID1(ByVal row As Integer, ByVal col As Integer)
        Dim i As Integer
        Rem DATOS DE LA COLUMNA
        formatogrilla(1, 1) = "CODIGO"
        formatogrilla(1, 2) = "DESCRIPCION"
        formatogrilla(1, 3) = "CANTIDAD"
        formatogrilla(1, 4) = "COSTO"
        formatogrilla(1, 5) = "(%)R"
        formatogrilla(1, 6) = "PUBLICO"
        formatogrilla(1, 7) = "COTIZA"
        formatogrilla(1, 8) = "(%)F"
        formatogrilla(1, 9) = "NETO"
        formatogrilla(1, 10) = "TOTAL C/IVA"
        formatogrilla(1, 11) = "BODEGA"
        formatogrilla(1, 12) = "OK"
        formatogrilla(1, 13) = "F.U.C"
        
        Rem LARGO DE LOS DATOS
        formatogrilla(2, 1) = "13"
        formatogrilla(2, 2) = "50"
        formatogrilla(2, 3) = "9"
        formatogrilla(2, 4) = "9"
        formatogrilla(2, 5) = "9"
        formatogrilla(2, 6) = "9"
        formatogrilla(2, 7) = "9"
        formatogrilla(2, 8) = "9"
        
        Rem TIPO DE DATOS C=CEROS,S=STRING,N=NUMERICO,D=DATE,
        formatogrilla(3, 1) = "N"
        formatogrilla(3, 2) = "S"
        formatogrilla(3, 3) = "N"
        formatogrilla(3, 4) = "N"
        formatogrilla(3, 5) = "N"
        formatogrilla(3, 6) = "N"
        formatogrilla(3, 7) = "N"
        formatogrilla(3, 8) = "N"
        formatogrilla(3, 9) = "N"
        formatogrilla(3, 10) = "N"
        formatogrilla(3, 11) = "N"
        formatogrilla(3, 12) = "N"
        formatogrilla(3, 13) = "D"
        
        
        Rem FORMATO GRILLA
        formatogrilla(4, 1) = "0000000000000"
        formatogrilla(4, 2) = ""
        formatogrilla(4, 3) = "###,###,##0.00"
        formatogrilla(4, 4) = "$ ###,###,##0.00"
        formatogrilla(4, 5) = "% #,##0.00"
        formatogrilla(4, 6) = "$ ###,###,##0"
        formatogrilla(4, 7) = "$ ###,###,##0"
        formatogrilla(4, 8) = "% #,##0.00"
        formatogrilla(4, 9) = "$ ###,###,##0.00"
        formatogrilla(4, 10) = "$ ###,###,##0"
        
        Rem LOCCKED
        formatogrilla(5, 1) = "FALSE"
        formatogrilla(5, 2) = "TRUE"
        formatogrilla(5, 3) = "FALSE"
        formatogrilla(5, 4) = "TRUE"
        formatogrilla(5, 5) = "TRUE"
        formatogrilla(5, 6) = "TRUE"
        formatogrilla(5, 7) = "FALSE"
        formatogrilla(5, 8) = "TRUE"
        formatogrilla(5, 9) = "TRUE"
        formatogrilla(5, 10) = "TRUE"
        formatogrilla(5, 11) = "FALSE"
        formatogrilla(5, 12) = "FALSE"
        formatogrilla(5, 13) = "FALSE"
        
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
        formatogrilla(8, 2) = "24"
        formatogrilla(8, 3) = "7"
        formatogrilla(8, 4) = "8"
        formatogrilla(8, 5) = "5"
        formatogrilla(8, 6) = "7"
        formatogrilla(8, 7) = "7"
        formatogrilla(8, 8) = "5"
        formatogrilla(8, 9) = "7"
        formatogrilla(8, 10) = "7"
        formatogrilla(8, 11) = "7"
        formatogrilla(8, 12) = "3"
        formatogrilla(8, 13) = "8"
        
        col = 14
        GRID1.Cols = col
        GRID1.Rows = row
        GRID1.AllowUserResizing = False
        GRID1.DisplayFocusRect = False
        GRID1.ExtendLastCol = True
        GRID1.BoldFixedCell = False
        GRID1.DrawMode = cellOwnerDraw
        GRID1.Appearance = Flat
        GRID1.ScrollBarStyle = Flat
        GRID1.FixedRowColStyle = Flat
        GRID1.BackColorFixed = RGB(90, 158, 214)
        GRID1.BackColorFixedSel = RGB(110, 180, 230)
        GRID1.BackColorBkg = RGB(90, 158, 214)
        GRID1.BackColorScrollBar = RGB(231, 235, 247)
        GRID1.BackColor1 = RGB(231, 235, 247)
        GRID1.BackColor2 = RGB(239, 243, 255)
        GRID1.GridColor = RGB(148, 190, 231)
        
        GRID1.Column(0).Width = 0
        For i = 1 To col - 1
            GRID1.Cell(0, i).text = formatogrilla(1, i)
            GRID1.Column(i).Width = Val(formatogrilla(8, i)) * (GRID1.Cell(0, i).Font.Size + 1.25)
            GRID1.Column(i).MaxLength = Val(formatogrilla(2, i))
            GRID1.Column(i).FormatString = formatogrilla(4, i)
            GRID1.Column(i).Locked = formatogrilla(5, i)
            If formatogrilla(3, i) = "N" Then
                GRID1.Column(i).Alignment = cellRightCenter
            Else
                GRID1.Column(i).Alignment = cellLeftCenter
            End If
        Next i
        GRID1.Cell(0, 1).Alignment = cellCenterCenter
        GRID1.Cell(0, 2).Alignment = cellCenterCenter
        GRID1.Cell(0, 3).Alignment = cellCenterCenter
        GRID1.Cell(0, 4).Alignment = cellCenterCenter
        GRID1.Cell(0, 5).Alignment = cellCenterCenter
        GRID1.Cell(0, 6).Alignment = cellCenterCenter
        GRID1.Cell(0, 7).Alignment = cellCenterCenter
        'GRID1.Enabled = True
    GRID1.Column(1).Mask = cellNumeric
    GRID1.Column(3).Mask = cellNumeric
    GRID1.Column(7).Mask = cellNumeric
    GRID1.Column(12).CellType = cellCheckBox
    
    
    
    
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
                FECHA = dato3.text & "-" & dato4.text & "-" & dato5.text
               
                    dato6.SetFocus
                
            End If
            
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
            FECHA = dato3.text & "-" & dato4.text & "-" & dato5.text
           
               dato6.SetFocus
               
            
        End If
    End Sub
'Private Sub detalle_LeaveCell(ByVal row As Long, ByVal col As Long, NewRow As Long, NewCol As Long, Cancel As Boolean)
'        Dim i As Integer
'        Dim linea As String
'        Dim limite As Integer
'        Dim descu As Double
'        Dim PRECIO As Double
'        Dim descu2 As Double
'
'        If detalle.Rows <= NewRow Then
'            NewRow = row
'        End If
'
'       If vacio = True Then
'            If NewRow <> row And row <> detalle.Rows - 1 Then
'                NewRow = fila
'                NewCol = columna
'            Else
'                If NewCol > col Then
'                    NewRow = fila
'                    NewCol = columna
'                End If
'            End If
'        Else
''         If (NewCol <> col Or NewRow <> row) And col = 5 And detalle.Cell(col, row).text = "0" Then
''         NewCol = col
''         NewRow = row
''
''
''         End If
'
'            If col = 6 And NewCol = 7 Then
'                If detalle.ActiveCell.text <> "" Then
'                    If row = detalle.Rows - 1 Then
'                        Select Case dato1.text
'                            Case "FV", "FE", "GD", "NP", "CO", "ZE"
'                                limite = 35
'
'                            Case "BV"
'                                limite = 100
'
'                            Case Else
'                                limite = 0
'                        End Select
'                        If limite > 0 Then
'                            If limite > detalle.Rows - 1 Then
'                                detalle.AddItem vbTab & vbTab & "1" & vbTab & "0" & vbTab & "0" & vbTab & "0" & vbTab & "0" & vbTab & "", True
'
'                                NewRow = detalle.Rows - 1
'                                NewCol = 1
'                            Else
'                                dato11.SetFocus
'                            End If
'                        Else
'                            detalle.AddItem vbTab & vbTab & "1" & vbTab & "0" & vbTab & "0" & vbTab & "0" & vbTab & "0" & vbTab & "0" & vbTab & "", True
'
'                            NewRow = detalle.Rows - 1
'                            NewCol = 1
'                        End If
'                    Else
'                        NewCol = 1
'                    End If
'                Else
'                    NewCol = col
'                End If
'            Else
'                If col = 5 And NewCol = 4 Then
'                    NewCol = 3
'                End If
'            End If
'            If col = 1 And NewCol = detalle.Cols - 1 Then
'                For i = 1 To detalle.Cols - 1
'                    If detalle.Cell(NewRow, i).text = "" Then
'                        NewCol = i
'                        Exit For
'                    End If
'                Next i
'            End If
'
'            If col = 1 And NewCol < detalle.Cols - 2 And NewCol > 1 Then
'                pivote.text = detalle.Cell(row, 1).text
'                If pivote.text <> "0000000000000" Then
'                pivote.text = ceros(pivote)
'                If Val(pivote.text) = "0" Then
'                    pivote.text = ""
'                End If
'                End If
'
'                detalle.Cell(row, 1).text = pivote.text
'                detalle.Cell(row, 1).text = leeralias(detalle.Cell(detalle.ActiveCell.row, 1).text)
'                If leerCodigoProducto(detalle.Cell(row, 1).text) = False Then
'                    detalle.Cell(row, 1).text = ""
'                    detalle.Cell(row, 1).SetFocus
'
'                End If
'
''                If leerstock(detalle.Cell(Row, 1).text) = False Then
'''                    detalle.Cell(Row, 1).text = ""
'''                      If MsgBox("PRODUCTO SIN STOCK", vbOKOnly, "ATENCION") = vbOK Then
''''                        glosas.Visible = True
''''                         dato24.SetFocus
'''                      End If
''               End If
'                           Rem detalle.Cell(row, detalle.Cols - 1).text = leerCostoProducto(detalle.Cell(row, 1).text)
'                           detalle.Cell(row, 2).text = leerNombreProducto(detalle.Cell(row, 1).text)
'                           If detalle.Cell(row, 2).text <> "" Then
'                              NewCol = 3
'                           End If
'                           If row > 0 Then
'                              detalle.Cell(row, 4).text = "1"
'                           End If
'                           Rem If Val(detalle.Cell(row, 5).text) = 0 Then
'                           detalle.Cell(row, 5).text = leerPrecioEspecial(detalle.Cell(row, 1).text)
'                           Rem If Val(detalle.Cell(row, 5).text) = 0 Then
'                           detalle.Cell(row, 5).text = leerPrecioProducto(detalle.Cell(row, 1).text, tipoprecio)
'                           If leerPrecioProducto(detalle.Cell(row, 1).text, tipoprecio) = "0" Then detalle.Cell(row, 5).text = ""
'                           Rem End If
'                           Rem End If
'                           If detalle.Cell(row, 3).text <> "" And detalle.Cell(row, 5).text <> "" And row > 0 Then
'
'                           Else
'                               detalle.Cell(row, 7).text = "0"
'                           End If
'               Else
'                    If col = 1 And NewCol = detalle.Cols - 2 Then
'                       NewCol = 5
'                    End If
'            End If
'            If col = 3 And NewCol <> col Then
'                If detalle.Cell(row, 3).text <> "" And CDbl(detalle.Cell(row, 3).text) > 0 Then
'                    If NewCol > col Then
'                        NewCol = 5
'                    End If
'                    If NewCol < col Then
'                        NewCol = 1
'                    End If
'                Else
'                    NewCol = col
'                End If
'            End If
'            If NewRow > row Then
'                For i = 1 To detalle.Cols - 3
'                    If detalle.Cell(row, i).text = "" Then
'                        NewRow = row
'                        NewCol = i
'                        Exit For
'                    End If
'                Next i
'                For i = 1 To detalle.Cols - 3
'                    If detalle.Cell(NewRow, i).text = "" Then
'                        NewCol = i
'                        Exit For
'                    End If
'                Next i
'            End If
'            If row > 0 Then
'                If detalle.Cell(row, 3).text <> "" And detalle.Cell(row, 5).text <> "" Then
'                    If dato1.text <> "WW" Then
'                        detalle.Cell(row, 7).text = Round(detalle.Cell(row, 3).text * detalle.Cell(row, 5).text + 0.1, 0)
'                    Else
'                        If Val(detalle.Cell(row, 7).text) <> Val(detalle.Cell(row, 3).text) * Val(detalle.Cell(row, 5).text) Then
'                            detalle.Cell(row, 7).text = detalle.Cell(row, 5).text
'                            detalle.Cell(row, 5).text = detalle.Cell(row, 7).text / detalle.Cell(row, 3).text
'                        End If
'                    End If
'                    detalle.Cell(row, 4).text = "1"
'                End If
'            End If
''            If row > 0 Then
''                precio = Round(detalle.Cell(row, 3).text * detalle.Cell(row, 5).text + 0.1, 0)
''                descu = Int((detalle.Cell(row, 7).text * ((detalle.Cell(row, 6).text) / 100)) + 0.5)
''                detalle.Cell(row, 7).text = Str(precio - descu)
''            End If
'            If NewCol = 5 Then
'            filaprecio = row
'
'            INGRESAPRECIO.Show vbModal
'
'            End If
'
'            If col > 0 And row > 0 Then
'                Call sumaGrilla(detalle)
'            End If
'        End If
''        End If
'
'
'    End Sub
        
Private Sub GRID1_KeyDown(KeyCode As Integer, ByVal Shift As Integer)
       If GRID1.ActiveCell.col = 1 And KeyCode = vbKeyF2 Then
       Call ayudaProducto2(GRID1, pivote): GRID1.Cell(GRID1.ActiveCell.row, GRID1.ActiveCell.col).SetFocus
       End If
End Sub

Private Sub GRID1_LeaveCell(ByVal row As Long, ByVal col As Long, NewRow As Long, NewCol As Long, Cancel As Boolean)
Dim venta As Double
Dim costo As Double
Dim margen As Double

If GRID1.Rows > 1 Then
            Rem busca el codigo
            
            If col = 1 And NewCol = 2 Then
            NewCol = 3
            End If
            
            If col = 3 And NewCol = 4 Then
            NewCol = 7
            End If
            
            If col = 7 And NewCol = 6 Then
            NewCol = 3
            End If
            
            If col = 3 And NewCol = 2 Then
            NewCol = 1
            End If
            If col = 7 And NewCol = 8 Then
            NewCol = 11
            End If
            
            
            If col = 1 And NewRow = row Then
                GRID1.Cell(row, 1).text = Format(GRID1.Cell(row, 1).text, "0000000000000")
                    If LEERPRODUCTO(GRID1.Cell(row, 1).text, row) = False Then
                    NewRow = row: NewCol = col
                    End If
            End If
            Rem cambia el precio
            If (col = 7 Or col = 3) And GRID1.Cell(row, col).text = "" Then
            GRID1.Cell(row, col).text = "0"
            
            End If
            If col = 3 And GRID1.Cell(row, 3).text = "0" Then
            MsgBox ("DEBE COTIZAR UNA CANTIDAD DISTINTA DE CERO")
            NewCol = col
            
            End If
            
            
            If (col = 7 Or col = 3) And GRID1.Cell(row, 7).text <> "" Then
            costo = CDbl(GRID1.Cell(row, 4).text)
            venta = CDbl(GRID1.Cell(row, 7).text)
            margen = venta / costo
            margen = (margen - 1) * 100
            GRID1.Cell(row, 8).text = margen
            GRID1.Cell(row, 9).text = Round(venta / 1.19, 2)
            GRID1.Cell(row, 10).text = CDbl(GRID1.Cell(row, 3).text) * CDbl(GRID1.Cell(row, 7).text)
            If margen < 0 Then
            MsgBox ("PRECIO DE VENTA BAJO EL COSTO")
            NewRow = row: NewCol = col
            End If
            End If
            
            
            If col = 11 And GRID1.Cell(row, 10).text <> "" And row = GRID1.Rows - 1 Then
            GRID1.Rows = GRID1.Rows + 1
            NewRow = row + 1: NewCol = 1
            End If
            If NewCol = 11 Then
            despacho.Show vbModal
            
            End If
            
            
            sumargrilla
            
            


End If

End Sub
Public Function LEERPRODUCTO(CODIGObarra, fila) As Boolean
        Dim campos(10, 10)
        
        Dim op As Integer
        Dim costo As Double
        Dim costo1 As Double
        Dim costo2 As Double
        Dim FECHA44 As Date
        
        Dim venta As Double
        Dim margen As Double
        
        Set sql = New sqlventas.sqlventa
        campos(0, 0) = "descripcion"
        campos(1, 0) = "pcosto"
        campos(2, 0) = ""
        campos(0, 2) = "r_maestroproductos_fijo_" & rubro
        
        condicion = "codigobarra = '" & CODIGObarra & "'"
        op = 5
        sql.response = campos
        Set sql.conexion = gestionRubro
        Call sql.sqlventas(op, condicion)
        If sql.Status = 0 Then
        LEERPRODUCTO = True
        GRID1.Cell(fila, 2).text = sql.response(0, 3)
        GRID1.Cell(fila, 4).text = sql.response(1, 3)
        
        Rem costo = sql.response(1, 3)
        If empresaActiva = "01" Or empresaActiva = "20" Then
        costo1 = leeultimacompra(CODIGObarra, "01")
        costo2 = leeultimacompra(CODIGObarra, "20")
        If fechaultimacompra1 > fechaultimacompra2 Then
        costo = costo1
        FECHA44 = fechaultimacompra1
        
        Else
        costo = costo2
        FECHA44 = fechaultimacompra2
        
        End If
        
        Else
        costo = leeultimacompra(CODIGObarra, empresaActiva)
        FECHA44 = fechaultimacompra1
        End If
        
        
       
        If costo = 0 Then costo = sql.response(1, 3)
        venta = leerPrecioProducto(CODIGObarra, "01")
        margen = venta / costo
        GRID1.Cell(fila, 5).text = (margen - 1) * 100
        GRID1.Cell(fila, 4).text = costo
        
        GRID1.Cell(fila, 6).text = venta
        GRID1.Cell(fila, 13).text = Format(FECHA44, "dd-mm-yyyy")
        
        End If
        
    End Function

Private Sub MSComm2_OnComm()
Dim str_rec As String

'Permanece escuchando al puerto COM y recibe la data de respuesta a los comandos enviados

    str_rec = MSComm2.Input
    
    If Len(str_rec) > 0 Then
        str_rec = Trim(str_rec)
     Rem   Timer1.Enabled = False
' Si recibe un Ascii 10, indica que la printer ha aceptado el comando y lo ha procesado correctamente
        If Left(str_rec, 1) = Chr$(10) Then
            'Lbl_status_IF.Caption = " Impresora OK !"
          
        Else
' Si recibio un Ascii distinto a 10, implica que la printer por alguna razon no pudo procesar el error
' Se pasa el caracter para analizar la respuesta y mostrar mensaje al operador
            'procesa_error (Mid(str_rec, 1, 3))
        End If
    End If
End Sub




Private Sub Grid1_Click()
Rem Call CARGASTOCKBODEGAS(GRID1.Cell(GRID1.ActiveCell.row, 2).text)
End Sub

Private Sub Timer1_Timer()
'    Static estado As Integer
'    estado = 1 - estado
'    lblMensaje.ForeColor = color(estado)
End Sub


    Private Sub opciones_FSCommand(ByVal command As String, ByVal args As String)
        Select Case command
            Case "modifica"
                If Verifica_Permiso(Me.Caption, "modifica") = True Then
                    Call modificar
                Else
                    MsgBox mensaje_nopermiso, vbCritical + vbOKOnly, "Permiso Denegado"
                End If
            Case "elimina"
            If Verifica_Permiso(Me.Caption, "elimina") = True Then
                If MsgBox("DESEA REALMENTE ELIMINAR Si / No", vbYesNo) = vbYes Then
                    frmglosaeliminacion.Show vbModal
                    Call ELIMINAR
                    retorno
                End If
            Else
                MsgBox mensaje_nopermiso, vbCritical + vbOKOnly, "Permiso Denegado"
            End If
                
            Case "imprime"
                Call imprimir
            Case "movimientos"
            Case "historico"
            Case "retorno"
                If modificando = True Then
                Call grabar2_Click
                
                End If
                
                Call retorno
            Case "anterior"
                Call anterior
            Case "siguiente"
                Call siguiente
        End Select
    End Sub
Sub modificar()
modificando = True
GRID1.SelectionMode = cellSelectionFree


grabar2.Visible = True


End Sub
Sub ELIMINAR()
    
    Dim op As Integer
    Dim sql As New sqlventas.sqlventa
    Dim campos(3, 3) As String
    
    campos(0, 2) = "sv_cotizaciones_" + empresaActiva
    condicion = "numero='" + dato2.text + "'"
    op = 4
    sql.response = campos
    sql.audit = True: sql.programaactivo = Me.Caption
    Set sql.conauditoria = conauditoria: sql.usuarioauditoria = usuarioSistema
    sql.glosaeliminacion = glosaeliminacionsistema
    sql.solicitoeliminacion = solicitaeliminacion
    Set sql.conexion = ventasRubro
    Call sql.sqlventas(op, condicion)
    Call consultaReplicas(generacadena(campos, op), clientesistema + "ventas" + empresaActiva)
    
    
End Sub
Sub autorizar(chequeo)
    
    Dim op As Integer
    Dim sql As New sqlventas.sqlventa
    Dim campos(3, 3) As String
    campos(0, 0) = "autorizada"
    campos(1, 0) = ""
    campos(0, 1) = chequeo
    
    campos(0, 2) = "sv_cotizaciones_" + empresaActiva
    condicion = "numero='" + dato2.text + "'"
    op = 3
    sql.response = campos
    Set sql.conexion = ventasRubro
    Call sql.sqlventas(op, condicion)
    Call consultaReplicas(generacadena(campos, op), clientesistema + "ventas" + empresaActiva)
    
End Sub


Sub siguiente()

End Sub
Sub anterior()

End Sub
Sub retorno()
dato2.text = ""
dato3.text = ""
dato4.text = ""
dato5.text = ""
dato6.text = ""
DATO7.text = ""
lbldv.Caption = ""
LBLNOMBRE.Caption = ""
LBLDIRECCION.Caption = ""
LBLCIUDAD.Caption = ""
LBLFONO.Caption = ""
lblneto.Caption = "0"
lbliva.Caption = "0"
lbltotal.Caption = "0"
GRID1.Rows = 1
opciones.Visible = False
grabar2.Visible = True
Check1.Value = 0
GRID1.SelectionMode = cellSelectionFree


dato2.SetFocus

End Sub
Sub imprimir()
    Dim row As Double
    Dim K As Integer
    
    Dim i As Integer
    Dim objReportTitle As FlexCell.ReportTitle
    GRID1.FixedRowColStyle = Fixed3D
    GRID1.CellBorderColorFixed = vbButtonShadow
    GRID1.ShowResizeTips = False
    GRID1.ReportTitles.Clear
    
    
    'ENCABEZADO DE PAGINA
    GRID1.PageSetup.Header = nombreempresa & vbCrLf & direccionempresa & vbCrLf & comunaempresa
    GRID1.PageSetup.HeaderAlignment = cellLeft
    GRID1.PageSetup.HeaderFont.Name = "Verdana"
    GRID1.PageSetup.HeaderFont.Size = 8
    
    'TITULOS DEL REPORTE
    Set objReportTitle = New FlexCell.ReportTitle
    objReportTitle.text = "COTIZACION NUMERO " + dato2.text
    objReportTitle.Font.Name = "Verdana" '"Times New Roman"
    objReportTitle.Font.Size = 10
    objReportTitle.Font.Bold = True
    objReportTitle.Align = cellCenter
    objReportTitle.PrintOnAllPages = True
    GRID1.ReportTitles.Add objReportTitle
    
    Set objReportTitle = New FlexCell.ReportTitle
    GRID1.ReportTitles.Add objReportTitle
    
    Set objReportTitle = New FlexCell.ReportTitle
    objReportTitle.text = "FECHA           :" + dato3.text + "-" + dato4.text + "-" + dato5.text
    objReportTitle.Font.Name = "Verdana" '"Times New Roman"
    objReportTitle.Font.Size = 8
    objReportTitle.Font.Bold = False
    objReportTitle.Align = cellLeft
    
    objReportTitle.PrintOnAllPages = True
    GRID1.ReportTitles.Add objReportTitle
    
    Set objReportTitle = New FlexCell.ReportTitle
    objReportTitle.text = "RUT              :" + dato6.text + "-" + lbldv.Caption
    objReportTitle.Font.Name = "Verdana" '"Times New Roman"
    objReportTitle.Font.Size = 8
    objReportTitle.Font.Bold = False
    objReportTitle.Align = cellLeft
    objReportTitle.PrintOnAllPages = True
    GRID1.ReportTitles.Add objReportTitle
    
    Set objReportTitle = New FlexCell.ReportTitle
    objReportTitle.text = "NOMBRE       :" + LBLNOMBRE.Caption
    objReportTitle.Font.Name = "Verdana" '"Times New Roman"
    objReportTitle.Font.Size = 8
    objReportTitle.Font.Bold = False
    objReportTitle.Align = cellLeft
    objReportTitle.PrintOnAllPages = True
    GRID1.ReportTitles.Add objReportTitle
    
    Set objReportTitle = New FlexCell.ReportTitle
    objReportTitle.text = "DIRECCION  : " + LBLDIRECCION.Caption
    objReportTitle.Font.Name = "Verdana" '"Times New Roman"
    objReportTitle.Font.Size = 8
    objReportTitle.Font.Bold = False
    objReportTitle.Align = cellLeft
    objReportTitle.PrintOnAllPages = True
    GRID1.ReportTitles.Add objReportTitle
    
    Set objReportTitle = New FlexCell.ReportTitle
    objReportTitle.text = "CIUDAD       : " + LBLCIUDAD.Caption
    objReportTitle.Font.Name = "Verdana" '"Times New Roman"
    objReportTitle.Font.Size = 8
    objReportTitle.Font.Bold = False
    objReportTitle.Align = cellLeft
    objReportTitle.PrintOnAllPages = True
    GRID1.ReportTitles.Add objReportTitle
    
    Set objReportTitle = New FlexCell.ReportTitle
    objReportTitle.text = "FONO      :" + LBLFONO.Caption
    objReportTitle.Font.Name = "Verdana" '"Times New Roman"
    objReportTitle.Font.Size = 8
    objReportTitle.Font.Bold = False
    objReportTitle.Align = cellLeft
    objReportTitle.PrintOnAllPages = True
    GRID1.ReportTitles.Add objReportTitle
    
    
    'PIE DE PAGINA
    
    
    
    
    GRID1.PageSetup.Footer = "Pág &P de &N" & vbCrLf & "Fecha: &D"
    GRID1.PageSetup.FooterAlignment = cellRight
    GRID1.PageSetup.FooterFont.Name = "Verdana"
    GRID1.PageSetup.FooterFont.Size = 7
    
    GRID1.PageSetup.LeftMargin = 1
    GRID1.PageSetup.RightMargin = 0.5
    GRID1.PageSetup.PrintGridlines = True
    GRID1.PageSetup.BlackAndWhite = True
    
    
    
    GRID1.PageSetup.PrintFixedRow = True
GRID1.AutoRedraw = False
    
    GRID1.Column(4).Width = 0
    GRID1.Column(5).Width = 0
    GRID1.Column(6).Width = 0
    GRID1.Column(8).Width = 0
    GRID1.Column(11).Width = 0
    GRID1.Column(12).Width = 0
    
    GRID1.Column(10).Width = 10 * (GRID1.DefaultFont.Size + 1.25)
    row = GRID1.Rows
    
    GRID1.Rows = GRID1.Rows + 1
    GRID1.Cell(GRID1.Rows - 1, 9).text = "NETO "
    GRID1.Cell(GRID1.Rows - 1, 10).text = lblneto.Caption
    
    
    GRID1.Rows = GRID1.Rows + 1
    GRID1.Cell(GRID1.Rows - 1, 9).text = "I.V.A "
    GRID1.Cell(GRID1.Rows - 1, 10).text = lbliva.Caption
    
    GRID1.Rows = GRID1.Rows + 1
    GRID1.Cell(GRID1.Rows - 1, 9).text = "TOTAL "
    GRID1.Cell(GRID1.Rows - 1, 10).text = lbltotal.Caption
    
    GRID1.Range(0, 0, row - 1, 12).Borders(cellEdgeTop) = cellThick
    GRID1.Range(0, 0, row - 1, 12).Borders(cellLeftTop) = cellThick
    GRID1.Range(0, 0, row - 1, 12).Borders(cellRightTop) = cellThick
    GRID1.Range(0, 0, row - 1, 12).Borders(cellEdgeBottom) = cellThick
    GRID1.Range(0, 0, row - 1, 12).Borders(cellInsideHorizontal) = cellThick
    GRID1.Range(0, 0, row - 1, 12).Borders(cellInsideVertical) = cellThick
    
    GRID1.Rows = GRID1.Rows + 4
    For K = 1 To 10
    GRID1.Column(K).Locked = False
    
    Next K
    
    GRID1.Range(GRID1.Rows - 1, 1, GRID1.Rows - 1, 10).Merge
    GRID1.Cell(GRID1.Rows - 1, 1).Alignment = cellCenterCenter
    GRID1.Cell(GRID1.Rows - 1, 1).text = "VALIDEZ DE LA COTIZACION : 3 DIAS A PARTIR "
    
    GRID1.Rows = GRID1.Rows + 1
    GRID1.Range(GRID1.Rows - 1, 1, GRID1.Rows - 1, 10).Merge
    GRID1.Cell(GRID1.Rows - 1, 1).Alignment = cellCenterCenter
    GRID1.Cell(GRID1.Rows - 1, 1).text = "DE LA FECHA DE EMISION"
    
    GRID1.Rows = GRID1.Rows + 4
    GRID1.Range(GRID1.Rows - 1, 1, GRID1.Rows - 1, 10).Merge
    GRID1.Cell(GRID1.Rows - 1, 1).Alignment = cellCenterCenter
    GRID1.Cell(GRID1.Rows - 1, 1).text = "   V B GERENCIA                                                                     V B " + lblvendedor.Caption
    
    
    
    
    For K = 1 To True
    GRID1.Column(K).Locked = False
    Next K
    
    

GRID1.PageSetup.PrintGridlines = False

GRID1.PrintPreview
    
    GRID1.Column(4).Width = 8 * (GRID1.DefaultFont.Size + 1.25)
    GRID1.Column(5).Width = 8 * (GRID1.DefaultFont.Size + 1.25)
    GRID1.Column(6).Width = 6 * (GRID1.DefaultFont.Size + 1.25)
    GRID1.Column(8).Width = 6 * (GRID1.DefaultFont.Size + 1.25)
    
GRID1.AutoRedraw = True


  GRID1.Range(0, 0, row - 1, 12).Borders(cellEdgeTop) = cellNone
  
    GRID1.Range(0, 0, row - 1, 12).Borders(cellLeftTop) = cellNone
    GRID1.Range(0, 0, row - 1, 12).Borders(cellRightTop) = cellNone
    GRID1.Range(0, 0, row - 1, 12).Borders(cellEdgeBottom) = cellNone
    GRID1.Range(0, 0, row - 1, 12).Borders(cellInsideHorizontal) = cellNone
    GRID1.Range(0, 0, row - 1, 12).Borders(cellInsideVertical) = cellNone
    

GRID1.Rows = row

 GRID1.Column(1).Locked = False
 GRID1.Column(3).Locked = False
 GRID1.Column(7).Locked = False
 GRID1.Column(11).Locked = False
 GRID1.Column(12).Locked = False
        
End Sub

Private Sub Timer2_Timer()
   Rem  sincronizarFechaHora
    fechasistema = Format(Now, "dd-mm-yyyy")
    HORA.Caption = Format(fechasistema, "dd-mm-yyyy") & " " & Time

End Sub


Private Sub txtCancelado_GotFocus()
    Call selecciona(txtCancelado)
End Sub


Public Sub Comprueba_Impresora()
Dim puerto As Integer
Dim cadena As String

puerto = 1


' Setea Puerto y configuracion de protocolo
    MSComm2.CommPort = puerto
    MSComm2.Settings = "9600,n,8,1"
    MSComm2.Handshaking = comRTS
    MSComm2.InputLen = 10
    MSComm2.RThreshold = 1
    MSComm2.SThreshold = 1
    MSComm2.PortOpen = True
   
    comando = 35
    cadena = Chr$(135)
    cadena = cadena & comando
    cadena = cadena & Chr$(136)
    MSComm2.Output = cadena
     Rem Timer1.Interval = 500
    Rem Timer1.Enabled = True
    'Sleep (70)
    Call MSComm2_OnComm


End Sub


Private Sub IMPRIMEPREVENTA(ByVal NUMERO As String, ByRef rollo As Adodc, ByRef lista As Grid)
    Dim tabla As String
    Dim i As Long
    Dim CODIGO As String
    Dim cantidad As String
    Dim precio As String
    Dim total As String
    Dim totalPreventa As Double
    Dim cadena As String
    Dim p As Printer
    Dim numfic As Integer
    Dim cSql As New rdoQuery
    Dim resultados As rdoResultset
    Dim caja As String
    
    Set cSql.ActiveConnection = ventasRubro
    
    tabla = "SELECT CURRENT_TIME AS hora, dc.numero, dc.fecha, dd.vendedor, dd.codigo, FORMAT(dd.cantidad,0) AS cantidad, FORMAT(dd.precio,0) AS precio, FORMAT(dd.total,0) AS total, dd.descripcion, dc.lugarretiro, dc.caja "
    tabla = tabla & "FROM sv_documento_cabeza_" + empresaActiva + " AS dc INNER JOIN sv_documento_detalle_" + empresaActiva + " AS dd ON dc.local = dd.local AND dc.numero = dd.numero "
    tabla = tabla & "WHERE dc.local = '00' AND dc.tipo = 'PV' AND dc.numero = '0000000001' "
    tabla = tabla & "ORDER BY dd.linea"
    cSql.sql = tabla
    cSql.Execute
    
    'Call ConectarControlData(rollo, servidor, baseventas &empresaactiva, usuario, password, tabla)
    lista.Rows = 1
    lista.Cols = 5
    lista.AutoRedraw = False
    
    lista.PageSetup.HeaderMargin = 1.25
    lista.PageSetup.TopMargin = 1.25
    lista.PageSetup.FooterMargin = 0.5
    lista.PageSetup.BottomMargin = 0.5
    lista.PageSetup.LeftMargin = 0.5
    lista.PageSetup.RightMargin = 0.5
    
    lista.Column(0).Width = 0
    lista.Column(1).Width = 90
    lista.Column(2).Width = 35
    lista.Column(3).Width = 60
    lista.Column(4).Width = 65
    
    'lista.Column(0).Width = 0
    'lista.Column(1).Width = 50
    'lista.Column(2).Width = 30
    'lista.Column(3).Width = 30
    'lista.Column(4).Width = 30
    
    If cSql.RowsAffected > 0 Then
        Set resultados = cSql.OpenResultset
        
        lista.AddItem leerNombreEmpresa(empresaActiva), True
        lista.Range(lista.Rows - 1, 1, lista.Rows - 1, lista.Cols - 1).Merge
        lista.Range(lista.Rows - 1, 1, lista.Rows - 1, lista.Cols - 1).Alignment = cellCenterCenter
        
        lista.AddItem leerNombreRubro(rubro), True
        lista.Range(lista.Rows - 1, 1, lista.Rows - 1, lista.Cols - 1).Merge
        lista.Range(lista.Rows - 1, 1, lista.Rows - 1, lista.Cols - 1).Alignment = cellCenterCenter
        
        lista.AddItem "", True
        
        lista.AddItem "PRE-VENTA          NRO:  " & NUMERO
        lista.Range(lista.Rows - 1, 1, lista.Rows - 1, lista.Cols - 1).Merge
        lista.Range(lista.Rows - 1, 1, lista.Rows - 1, lista.Cols - 1).Alignment = cellCenterCenter
        lista.AddItem "", True
        lista.AddItem "FECHA: " & resultados("fecha") & "     HORA: " & Format(resultados("hora"), "hh:mm:ss")
        lista.Range(lista.Rows - 1, 1, lista.Rows - 1, lista.Cols - 1).Merge
        'lista.Range(lista.Rows - 1, 1, lista.Rows - 1, lista.Cols - 1).Alignment = cellCenterCenter
        lista.AddItem "VENDEDOR(A): " & leerNombreCajera(resultados("vendedor")), True
        lista.Range(lista.Rows - 1, 1, lista.Rows - 1, lista.Cols - 1).Merge
        lista.AddItem "PREVENTA: " & leerNombreCajera(resultados("caja")), True
        lista.Range(lista.Rows - 1, 1, lista.Rows - 1, lista.Cols - 1).Merge
        
        lista.AddItem "", True
        
        lista.AddItem "==========================================", True
        lista.Range(lista.Rows - 1, 1, lista.Rows - 1, lista.Cols - 1).Merge
        
        lista.AddItem "   PRODUCTO   CANT.   P.UNIT      MONTO   "
        '             "           13|    5|        9|          12"
        lista.Range(lista.Rows - 1, 1, lista.Rows - 1, lista.Cols - 1).Alignment = cellCenterCenter
        lista.AddItem "==========================================", True
        lista.Range(lista.Rows - 1, 1, lista.Rows - 1, lista.Cols - 1).Merge
        i = 1
        totalPreventa = 0
        While Not resultados.EOF
            CODIGO = resultados("codigo")
            cantidad = Replace(resultados("cantidad"), ",", ".")
            precio = Replace(resultados("precio"), ",", ".")
            total = Replace(resultados("total"), ",", ".")
            totalPreventa = totalPreventa + CDbl(total)
            
            cadena = CODIGO & " "
            cantidad = String(5 - Len(cantidad), " ") & cantidad
            cadena = cadena & cantidad & " "
            
            precio = String(9 - Len(precio), " ") & precio
            cadena = cadena & precio & " "
            
            total = String(12 - Len(total), " ") & total
            cadena = cadena & total
            
            
            lista.AddItem cadena
            
            lista.Range(lista.Rows - 1, 1, lista.Rows - 1, lista.Cols - 1).Merge
            lista.AddItem resultados("descripcion")
            lista.Range(lista.Rows - 1, 1, lista.Rows - 1, lista.Cols - 1).Merge
            
            resultados.MoveNext
            i = i * 4
        Wend
        Set cSql = Nothing
        cSql.Close
        Set resultados = Nothing
        
        lista.AddItem "==========================================", True
        lista.Range(lista.Rows - 1, 1, lista.Rows - 1, lista.Cols - 1).Merge
        lista.AddItem "  EMPAQUE", True
        lista.Range(lista.Rows - 1, 1, lista.Rows - 1, lista.Cols - 1).FontBold = True
        lista.Range(lista.Rows - 1, 1, lista.Rows - 1, lista.Cols - 1).FontSize = 16
        lista.Range(lista.Rows - 1, 1, lista.Rows - 1, lista.Cols - 1).Merge
        
        lista.AddItem "", True
        
        cadena = Format(totalPreventa, "###,###,##0")
        cadena = String(15 - Len(cadena), " ") & cadena
        cadena = "                   TOTAL: " & cadena
        lista.AddItem cadena, True
        lista.Range(lista.Rows - 1, 1, lista.Rows - 1, lista.Cols - 1).Merge
        
    End If
    lista.AutoRedraw = True
    lista.Refresh
    
'
'
'    For i = 0 To Printers.Count
'        If UCase(Printers(i).DeviceName) = "SRP350 PARTIAL CUT" Then
'            lista.PageSetup.PrinterName = Printers(i).DeviceName
'            Exit For
'        End If
'    Next i
'
'    For i = 1 To lista.PageSetup.PaperSizes.Count
'        If UCase(lista.PageSetup.PaperSizes.Item(i).PaperName) = "A4 LENGTH" Then
'            lista.PageSetup.PaperSize = lista.PageSetup.PaperSizes.Item(i).Kind
'            Exit For
'        End If
'    Next i
'    ''''''''''''''''''
'    numfic = FreeFile
'    Open "COM1:9600,N,8,1,CD0,CS0,DS0,OP0,RS,TB100,RB100" For Output As #numfic
'    ''''''''''''''''''
'
'    '''''''''''''''''''''
'    'EMPAQUE
'    '''''''''''''''''''''
'    Print #numfic, Chr$(27); Chr$(64) '
'    For i = 1 To lista.Rows - 4
'        Print #numfic, lista.Cell(i, 1).text
'    Next i
'    Print #numfic, Chr(29); Chr(33); Chr(33); lista.Cell(lista.Rows - 3, 1).text
'    Print #numfic,
'    Print #numfic,
'    Print #numfic,
'    Print #numfic,
'    Print #numfic,
'    Print #numfic, Chr(27); "i"
'    '''''''''''''''''''''
'    'EMPAQUE
'    '''''''''''''''''''''
'
'    '''''''''''''''''''''
'    'CLIENTE
'    '''''''''''''''''''''
'    lista.Cell(lista.Rows - 3, 1).text = "  CLIENTE"
'    Print #numfic, Chr$(27); Chr$(64) '
'    For i = 1 To lista.Rows - 1
'        If lista.Cell(i, 1).text <> "  CLIENTE" Then
'            Print #numfic, lista.Cell(i, 1).text
'        Else
'            Print #numfic, Chr(29); Chr(33); Chr(33); lista.Cell(i, 1).text
'            Print #numfic, Chr$(27); Chr$(64)
'        End If
'    Next i
'    Print #numfic,
'    Print #numfic,
'    Print #numfic,
'    Print #numfic,
'    Print #numfic,
'    Print #numfic, Chr(27); "i"
'    '''''''''''''''''''''
'    'CLIENTE
'    '''''''''''''''''''''
'
'    '''''''''''''''''''''
'    'PRE-VENTA
'    '''''''''''''''''''''
'    Print #numfic, Chr$(27); Chr$(64)
'    Print #numfic, Chr(29); Chr(33); Chr(33); "  PRE-VENTA"
'    Print #numfic, Chr$(27) & Chr$(64)
'    Print #numfic, Chr$(29) & Chr$(104) & Chr(100)
'    Print #numfic, Chr$(29) & Chr$(119) & Chr(2)
'    Print #numfic, Chr$(29) & Chr$(72) & Chr(50)
'    Print #numfic, Chr$(29) & Chr$(107) & Chr(4) & numero & Chr(0) & "MONTO A PAGAR $ "; totalPreventa
'    Print #numfic, Chr$(10) & Chr$(10) & Chr(10) & Chr(10)
'    Print #numfic, Chr(27); "i"
'    '''''''''''''''''''''
'    'PRE-VENTA
'    '''''''''''''''''''''
'    Close #numfic
    
    lista.PrintPreview
End Sub


Public Sub eliminarPreventa()
'    Dim i As Integer
'    Dim cSql As rdoQuery
'    Dim NUMERO As String
'    For i = 0 To UBound(numeros)
'        NUMERO = numeros(i)
'        If NUMERO <> "" Then
'            Set cSql = New rdoQuery
'            Set cSql.ActiveConnection = ventasRubro
'            cSql.sql = "DELETE FROM sv_documento_cabeza_" + empresaactiva + " "
'            cSql.sql = cSql.sql & "WHERE local = '" & empresaactiva & "' AND tipo = 'PV' AND numero = '" & NUMERO & "' "
'            cSql.Execute
'            cSql.Close
'            Set cSql = Nothing
'
'            Set cSql = New rdoQuery
'            Set cSql.ActiveConnection = ventasRubro
'            cSql.sql = "DELETE FROM sv_documento_detalle_" + empresaactiva + " "
'            cSql.sql = cSql.sql & "WHERE local = '" & empresaactiva & "' AND tipo = 'PV' AND numero = '" & NUMERO & "' "
'            cSql.Execute
'            cSql.Close
'            Set cSql = Nothing
'        Else
'            Exit For
'        End If
'    Next i
End Sub
Public Sub cargadato()

End Sub
Public Sub eliminarlinea(linea)
            
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

 Private Sub dato3_LostFocus()
    Call esfecha(dato3, dato4, dato5, "dd")
    End Sub
    Private Sub dato4_LostFocus()
    Call esfecha(dato3, dato4, dato5, "mm")
    End Sub
    Private Sub dato5_LostFocus()
    Call esfecha(dato3, dato4, dato5, "yyyy")
    End Sub

Sub sumargrilla()
Dim neto As Double
Dim iva As Double
Dim total As Double

Dim K As Integer

total = 0

For K = 1 To GRID1.Rows - 1
If GRID1.Cell(K, 10).text <> "" Then
total = total + CDbl(GRID1.Cell(K, 10).text)
End If
Next K
neto = total / 1.19
iva = total - neto
lblneto.Caption = Format(neto, "###,###,###")
lbliva.Caption = Format(iva, "###,###,###")
lbltotal.Caption = Format(total, "###,###,###")

End Sub
Public Sub grabar()
        Dim campos(20, 10) As String
        
        
        Dim op As Integer
        Dim K As Integer
        
        Set sql = New sqlventas.sqlventa
        campos(0, 0) = "numero"
        campos(1, 0) = "rut"
        campos(2, 0) = "sucursal"
        campos(3, 0) = "fecha"
        campos(4, 0) = "linea"
        campos(5, 0) = "codigo"
        campos(6, 0) = "descripcion"
        campos(7, 0) = "cantidad"
        campos(8, 0) = "costo"
        campos(9, 0) = "porcentajereal"
        campos(10, 0) = "precioventa"
        campos(11, 0) = "preciofinal"
        campos(12, 0) = "porcentajefinal"
        campos(13, 0) = "precioneto"
        campos(14, 0) = "total"
        campos(15, 0) = "vendedor"
        campos(16, 0) = "autorizada"
        campos(17, 0) = "despacho"
        campos(18, 0) = "aceptada"
        campos(19, 0) = ""
        
        
        campos(20, 0) = ""
        campos(0, 1) = dato2.text
        campos(1, 1) = dato6.text + lbldv.Caption
        campos(2, 1) = "0"
        campos(3, 1) = dato5.text + "-" + dato4.text + "-" + dato3.text
        campos(15, 1) = DATO7.text + LBLDV2.Caption
        campos(16, 1) = Check1.Value
        
        For K = 1 To GRID1.Rows - 1
        If GRID1.Cell(K, 9).text <> "" Then
        campos(4, 1) = K
        campos(5, 1) = GRID1.Cell(K, 1).text
        campos(6, 1) = GRID1.Cell(K, 2).text
        campos(7, 1) = GRID1.Cell(K, 3).text
        campos(8, 1) = GRID1.Cell(K, 4).text
        campos(9, 1) = Replace(GRID1.Cell(K, 5).text, ",", ".")
        campos(10, 1) = GRID1.Cell(K, 6).text
        campos(11, 1) = GRID1.Cell(K, 7).text
        campos(12, 1) = Replace(GRID1.Cell(K, 8).text, ",", ".")
        campos(13, 1) = Replace(GRID1.Cell(K, 9).text, ",", ".")
        campos(14, 1) = GRID1.Cell(K, 10).text
        campos(17, 1) = GRID1.Cell(K, 11).text
        campos(18, 1) = GRID1.Cell(K, 12).text
        
        campos(0, 2) = "sv_cotizaciones_" + empresaActiva
        
        op = 2
        condicion = ""
        sql.response = campos
        Set sql.conexion = ventasRubro
        sql.audit = True: sql.programaactivo = Me.Caption
        Set sql.conauditoria = conauditoria: sql.usuarioauditoria = usuarioSistema
        Call sql.sqlventas(op, condicion)
        Call consultaReplicas(generacadena(campos, op), clientesistema + "ventas" + empresaActiva)
        End If
        Next K
        modificando = False
        
    End Sub

Public Function leerUltimoFoliocotizacion() As String
    
    Dim op As Integer
    Dim sql As New sqlventas.sqlventa
    Dim campos(3, 3) As String
    
    campos(0, 0) = "IFNULL(MAX(numero) + 1,'0000000001')"
    campos(1, 0) = ""
    campos(0, 2) = "sv_cotizaciones_" + empresaActiva
    condicion = "numero>'0'"
    op = 5
    sql.response = campos
    Set sql.conexion = ventasRubro
    Call sql.sqlventas(op, condicion)
    If sql.Status = 0 Then
        If sql.response(0, 3) <> "" And sql.response(0, 3) <> "0" Then
            leerUltimoFoliocotizacion = Format(sql.response(0, 3), "0000000000")
            
        Else
            leerUltimoFoliocotizacion = "0000000001"
        End If
    End If
End Function

Public Function leercotizacion() As Boolean

    Dim resultados As rdoResultset
    Dim cSql As New rdoQuery

        Set cSql.ActiveConnection = ventasRubro
        
        cSql.sql = "SELECT * "
        cSql.sql = cSql.sql + "FROM sv_cotizaciones_" & empresaActiva & " "
        cSql.sql = cSql.sql + "WHERE numero='" + dato2.text + "' order by linea "
        cSql.Execute
        leercotizacion = False
        
        If cSql.RowsAffected > 0 Then
            Set resultados = cSql.OpenResultset
        GRID1.Rows = 1
        dato3.text = Format(resultados(3), "dd")
        dato4.text = Format(resultados(3), "mm")
        dato5.text = Format(resultados(3), "yyyy")
        dato6.text = Mid(resultados(1), 1, 9)
        lbldv.Caption = Mid(resultados(1), 10, 1)
        DATO7.text = Mid(resultados(15), 1, 9)
        LBLDV2.Caption = Mid(resultados(15), 10, 1)
        lblvendedor.Caption = leerNombreVendedor(DATO7.text + LBLDV2.Caption)
        Check1.Value = resultados(16)
        LBLNOMBRE.Caption = leerNombreCliente(dato6.text + lbldv.Caption)
        LBLDIRECCION.Caption = leerDireccionCliente(dato6.text + lbldv.Caption, DATO7.text)
        LBLCIUDAD.Caption = leerCiudadCliente(dato6.text + lbldv.Caption, DATO7.text)
            While Not resultados.EOF
        leercotizacion = True
        GRID1.Rows = GRID1.Rows + 1
        GRID1.Cell(GRID1.Rows - 1, 1).text = resultados(5)
        GRID1.Cell(GRID1.Rows - 1, 2).text = resultados(6)
        GRID1.Cell(GRID1.Rows - 1, 3).text = resultados(7)
        GRID1.Cell(GRID1.Rows - 1, 4).text = resultados(8)
        GRID1.Cell(GRID1.Rows - 1, 5).text = resultados(9)
        GRID1.Cell(GRID1.Rows - 1, 6).text = resultados(10)
        GRID1.Cell(GRID1.Rows - 1, 7).text = resultados(11)
        GRID1.Cell(GRID1.Rows - 1, 8).text = resultados(12)
        GRID1.Cell(GRID1.Rows - 1, 9).text = resultados(13)
        GRID1.Cell(GRID1.Rows - 1, 10).text = resultados(14)
        GRID1.Cell(GRID1.Rows - 1, 11).text = resultados(17)
        GRID1.Cell(GRID1.Rows - 1, 12).text = resultados(18)
        
                resultados.MoveNext
                
            
            Wend
            resultados.Close
            Set resultados = Nothing
        End If
        sumargrilla
        
End Function

