VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form ventas13 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   12585
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   15210
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   12585
   ScaleWidth      =   15210
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame4 
      Caption         =   "Frame4"
      Height          =   1335
      Left            =   10080
      TabIndex        =   15
      Top             =   11160
      Width           =   3975
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Frame3"
      Height          =   5775
      Left            =   2520
      TabIndex        =   14
      Top             =   6000
      Width           =   11775
      Begin VB.TextBox Text11 
         Height          =   375
         Left            =   240
         TabIndex        =   24
         Text            =   "Text1"
         Top             =   240
         Width           =   1815
      End
      Begin VB.TextBox Text10 
         Height          =   375
         Left            =   2400
         TabIndex        =   23
         Text            =   "juan pere"
         Top             =   240
         Width           =   4455
      End
      Begin VB.TextBox Text9 
         Height          =   375
         Left            =   6960
         TabIndex        =   22
         Text            =   "Text3"
         Top             =   240
         Width           =   1335
      End
      Begin VB.TextBox Text8 
         Height          =   375
         Left            =   8520
         TabIndex        =   21
         Text            =   "Text4"
         Top             =   240
         Width           =   1215
      End
      Begin VB.TextBox Text7 
         Height          =   375
         Left            =   10080
         TabIndex        =   20
         Text            =   "Text6"
         Top             =   240
         Width           =   1335
      End
      Begin MSDataGridLib.DataGrid DataGrid2 
         Height          =   4455
         Left            =   240
         TabIndex        =   16
         Top             =   720
         Width           =   11295
         _ExtentX        =   19923
         _ExtentY        =   7858
         _Version        =   393216
         HeadLines       =   1
         RowHeight       =   15
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   13322
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   13322
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Frame2"
      Height          =   10935
      Left            =   120
      TabIndex        =   10
      Top             =   480
      Width           =   2295
      Begin MSDataGridLib.DataGrid DataGrid3 
         Height          =   4455
         Left            =   120
         TabIndex        =   19
         Top             =   5520
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   7858
         _Version        =   393216
         HeadLines       =   1
         RowHeight       =   15
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   13322
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   13322
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
      Begin VB.CommandButton Command1 
         Caption         =   "compras anteriores"
         Height          =   375
         Left            =   840
         TabIndex        =   12
         Top             =   10200
         Width           =   1095
      End
      Begin VB.TextBox Text5 
         Height          =   375
         Left            =   720
         TabIndex        =   11
         Text            =   "Text5"
         Top             =   360
         Width           =   1455
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H008080FF&
      Caption         =   "Frame1"
      Height          =   5295
      Left            =   2520
      TabIndex        =   5
      Top             =   480
      Width           =   11775
      Begin VB.CommandButton Command3 
         Caption         =   "depto"
         Height          =   375
         Left            =   6720
         TabIndex        =   18
         Top             =   5400
         Width           =   1215
      End
      Begin VB.TextBox Text3 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0FF&
         Height          =   375
         Left            =   6960
         TabIndex        =   17
         Text            =   "Text3"
         Top             =   4800
         Width           =   1335
      End
      Begin VB.CommandButton Command2 
         Caption         =   "mix cliente"
         Height          =   375
         Left            =   5280
         TabIndex        =   13
         Top             =   5400
         Width           =   1215
      End
      Begin VB.CheckBox Check1 
         Caption         =   "solo stock"
         Height          =   495
         Left            =   240
         TabIndex        =   9
         Top             =   5280
         Width           =   1215
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0FF&
         Height          =   375
         Left            =   2400
         TabIndex        =   8
         Text            =   "juan pere"
         Top             =   4800
         Width           =   4455
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0FF&
         Height          =   375
         Left            =   240
         TabIndex        =   7
         Text            =   "Text1"
         Top             =   4800
         Width           =   1815
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   4335
         Left            =   240
         TabIndex        =   6
         Top             =   240
         Width           =   11415
         _ExtentX        =   20135
         _ExtentY        =   7646
         _Version        =   393216
         HeadLines       =   1
         RowHeight       =   15
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   13322
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   13322
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
   End
   Begin MSAdodcLib.Adodc ControlData 
      Height          =   330
      Left            =   120
      Top             =   11880
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
   Begin VB.Label NumeroNota 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9600
      TabIndex        =   4
      Top             =   7200
      Width           =   1935
   End
   Begin VB.Label Label8 
      Caption         =   "Nº"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8880
      TabIndex        =   3
      Top             =   7200
      Width           =   495
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      Caption         =   "NOTA DE VENTA"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3600
      TabIndex        =   2
      Top             =   7200
      Width           =   5295
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Caption         =   "LISTA DE PRODUCTOS DISPONIBLES"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   2760
      Width           =   12855
   End
   Begin VB.Label Label1 
      Caption         =   "Creación de Notas de Venta"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3720
      TabIndex        =   0
      Top             =   120
      Width           =   5295
   End
End
Attribute VB_Name = "ventas13"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
