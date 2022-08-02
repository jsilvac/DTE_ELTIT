VERSION 5.00
Object = "{4777436C-EB8C-4596-98A8-EBCF98683969}#1.0#0"; "Flexcell.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{ADBBAED6-B16F-40DC-80DF-B44910CBA76C}#1.0#0"; "Frame-Xp.ocx"
Begin VB.Form electro01 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mestro de Cuentas del Mayor"
   ClientHeight    =   6960
   ClientLeft      =   1320
   ClientTop       =   2925
   ClientWidth     =   12690
   DrawStyle       =   5  'Transparent
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   464
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   846
   Begin XPFrame.FrameXp FrameXp1 
      Height          =   6855
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   12495
      _ExtentX        =   22040
      _ExtentY        =   12091
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
      Begin VB.ListBox List1 
         Height          =   6495
         Left            =   8040
         TabIndex        =   11
         Top             =   240
         Width           =   4215
      End
      Begin VB.Frame datospersonales 
         BackColor       =   &H00FFF2F7&
         BorderStyle     =   0  'None
         Caption         =   "Datos personales"
         Height          =   6495
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   7815
         Begin VB.DirListBox Dir1 
            Height          =   1440
            Left            =   240
            TabIndex        =   8
            Top             =   600
            Width           =   3855
         End
         Begin VB.TextBox ARCHIVO 
            Height          =   285
            Left            =   4320
            TabIndex        =   7
            Top             =   1680
            Width           =   3375
         End
         Begin VB.DriveListBox Drive1 
            Height          =   315
            Left            =   240
            TabIndex        =   6
            Top             =   240
            Width           =   3855
         End
         Begin VB.FileListBox File1 
            Height          =   1260
            Left            =   4320
            TabIndex        =   5
            Top             =   240
            Width           =   3375
         End
         Begin VB.CommandButton Command1 
            Caption         =   "INSERTA CERTIFICADO"
            Height          =   375
            Left            =   2280
            TabIndex        =   4
            Top             =   2040
            Width           =   3375
         End
         Begin FlexCell.Grid Grid1 
            Height          =   3735
            Left            =   240
            TabIndex        =   3
            Top             =   2520
            Width           =   7455
            _ExtentX        =   13150
            _ExtentY        =   6588
            Cols            =   5
            DefaultFontSize =   8.25
            Rows            =   30
         End
         Begin VB.Label NOMBRETIPO 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
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
            Height          =   255
            Left            =   1920
            TabIndex        =   10
            Top             =   1560
            Width           =   1695
         End
         Begin VB.Label NOMBRETIPO2 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
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
            Height          =   255
            Left            =   1800
            TabIndex        =   9
            Top             =   2640
            Width           =   2175
         End
      End
   End
   Begin VB.TextBox pivote 
      Height          =   285
      Left            =   1440
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   6720
      Visible         =   0   'False
      Width           =   615
   End
   Begin MSAdodcLib.Adodc mcm 
      Height          =   375
      Left            =   2400
      Top             =   6840
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
End
Attribute VB_Name = "electro01"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private PASO As Double
Private TIPO(50, 2) As String
Private NUM As Integer
Private INI As Integer
Private pasa As Integer
Private fintipo As Integer
Private rubro_transaccion As String
Private k As Integer


Private Sub Command1_Click()
Call leerxml(File1.Path + "\" + ARCHIVO.text)

End Sub

Private Sub Dir1_Change()
    Dir1.Path = Drive1.Drive
    File1.Path = Dir1.Path
    File1.Pattern = "foli*.xml"
End Sub

Private Sub Drive1_Change()
    Dir1.Path = Drive1.Drive
    File1.Path = Dir1.Path
    File1.Pattern = "foli*.xml"
End Sub

Private Sub File1_DblClick()
    k = File1.ListIndex
    ARCHIVO.text = File1.List(k)
End Sub

Private Sub Form_Load()
    Drive1.Drive = "c:\"
    
End Sub


