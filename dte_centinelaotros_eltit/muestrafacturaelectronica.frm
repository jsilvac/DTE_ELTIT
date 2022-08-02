VERSION 5.00
Object = "{4777436C-EB8C-4596-98A8-EBCF98683969}#1.0#0"; "FlexCell.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{ADBBAED6-B16F-40DC-80DF-B44910CBA76C}#1.0#0"; "Frame-Xp.ocx"
Begin VB.Form electro03 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Emision Documentos Electronicos"
   ClientHeight    =   9435
   ClientLeft      =   435
   ClientTop       =   825
   ClientWidth     =   14775
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9435
   ScaleWidth      =   14775
   Begin FlexCell.Grid Grid1 
      Height          =   870
      Left            =   8415
      TabIndex        =   39
      Top             =   8505
      Visible         =   0   'False
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   1535
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
         Left            =   9120
         TabIndex        =   41
         Top             =   720
         Width           =   4260
      End
      Begin VB.Label LBLGIRO 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   240
         TabIndex        =   40
         Top             =   1080
         Width           =   13140
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
      Height          =   5595
      Left            =   135
      TabIndex        =   8
      Top             =   2250
      Width           =   14580
      _ExtentX        =   25718
      _ExtentY        =   9869
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
         Height          =   5205
         Left            =   135
         TabIndex        =   7
         Top             =   315
         Width           =   14385
         _ExtentX        =   25374
         _ExtentY        =   9181
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
         Caption         =   "Leer"
         Height          =   375
         Left            =   14040
         TabIndex        =   37
         Top             =   315
         Width           =   555
      End
      Begin VB.TextBox Text5 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
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
         Height          =   285
         Left            =   12690
         MaxLength       =   10
         TabIndex        =   25
         Tag             =   "fecha"
         Top             =   315
         Width           =   1230
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
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "FOLIO SII"
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
         Left            =   11745
         TabIndex        =   24
         Top             =   360
         Width           =   855
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
      Height          =   1410
      Left            =   120
      TabIndex        =   26
      Top             =   7875
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
      Begin VB.CommandButton Command2 
         Caption         =   "IMPRIMIR"
         Height          =   420
         Left            =   12240
         TabIndex        =   38
         Top             =   840
         Width           =   2040
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "TOTAL :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   240
         Left            =   11790
         TabIndex        =   36
         Top             =   360
         Width           =   870
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
         Height          =   375
         Left            =   12825
         TabIndex        =   35
         Top             =   315
         Width           =   1455
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "OT.IMPUESTOS"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   240
         Left            =   7650
         TabIndex        =   34
         Top             =   405
         Width           =   1710
      End
      Begin VB.Label lblotros 
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
         Height          =   375
         Left            =   9720
         TabIndex        =   33
         Top             =   315
         Width           =   1455
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "EXENTO :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   240
         Left            =   4995
         TabIndex        =   32
         Top             =   405
         Width           =   1050
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
         Height          =   375
         Left            =   6030
         TabIndex        =   31
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "I.V.A :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   240
         Left            =   2565
         TabIndex        =   30
         Top             =   405
         Width           =   615
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
         Height          =   375
         Left            =   3330
         TabIndex        =   29
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "NETO :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   240
         Left            =   90
         TabIndex        =   28
         Top             =   405
         Width           =   765
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
         Height          =   375
         Left            =   945
         TabIndex        =   27
         Top             =   360
         Width           =   1455
      End
   End
End
Attribute VB_Name = "ELECTRO03"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Private formatogrilla(20, 20)
    Private fecha1 As String
    Private fecha2 As String
    Public GIROEMPRESA As String
    Public NOMBREEMPRESA As String
    Public DIRECCIONEMPRESA As String
    Public COMUNAEMPRESA As String
    Public CIUDADEMPRESA As String
    Public RUTEMPRESA As String
    
Private Sub Command1_Click()
Call cargadocumento(DT0.text, DT1.text, DT3.text, DT2.text, DT6.text + "-" + DT5.text + "-" + DT4.text)

End Sub

'****************************************************************************
'Manejo de los Controles
'****************************************************************************
    '****************************************************************************
    'GOTFOCUS
    '****************************************************************************
    '****************************************************************************
    'GOTFOCUS
    '****************************************************************************
    
    '****************************************************************************
    'KEYDOWN
    '****************************************************************************
    

Private Sub Command2_Click()
Dim TIMBRE As String
Dim rutenvia As String
Dim CODIGODEACTIVIDAD As String

CODIGODEACTIVIDAD = "111"
rutenvia = "11636809-9"
CARGAGRILLAFACTURA



Rem TIMBRE = generatimbre("1.0", "775753005", "33", DT3.text, DT6.text + "-" + DT5.text + "-" + DT4.text, LBLRUT.Caption, lblnombre.Caption, LBLTOTAL.Caption, Informe.Cell(1, 2).text, "rooodododododododoodododododododo", TimeStamp, "dasdasdasdasds")

Unload Me


End Sub
Sub CARGAGRILLAFACTURA()

NOMBREEMPRESA = "FERRETERIA ELTIT LIMITADA"
GIROEMPRESA = "GIRO: VENTA AL POR MAYOR DE VEHICULOS AUTOMOTORES EXCEPTO MOTOCICLETAS,"
GIROEMPRESA = GIROEMPRESA + "VENTA DE MOTOCICLETAS, VENTA AL POR MENOR DE ARTICULOS DE FERRETERIA Y"
GIROEMPRESA = GIROEMPRESA + " MATERIALES DE CONSTRUCCION"
DIRECCIONEMPRESA = "CASA MATRIZ : OHIGGINS 336 - FONO : 45-411103 - FAX 45-411186 CASILLA 4-D PUCON "
DIRECCIONEMPRESA = DIRECCIONEMPRESA + "SUCURSAL 1 : MIGUEL ANSORENA 365 : PUCON "
DIRECCIONEMPRESA = DIRECCIONEMPRESA + "SUCURSAL 2: CAMINO INTERNACIONAL 200 - 2do nivel"
COMUNAEMPRESA = "COMUNA PUCON"
RUTEMPRESA = "77575300-5"




End Sub

Private Sub DT1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
DT2.SetFocus

End If

End Sub

Private Sub DT2_KeyPress(KeyAscii As Integer)
KeyAscii = esNumero(KeyAscii)
If KeyAscii = 13 Then
DT2.text = Format(DT2.text, "00")
DT3.SetFocus

End If

End Sub

Private Sub DT3_KeyPress(KeyAscii As Integer)
KeyAscii = esNumero(KeyAscii)
If KeyAscii = 13 Then
DT3.text = Format(DT3.text, "0000000000")
DT4.SetFocus

End If

End Sub
Private Sub DT4_KeyPress(KeyAscii As Integer)
KeyAscii = esNumero(KeyAscii)

If KeyAscii = 13 Then
DT4.text = Format(DT4.text, "00")
DT5.SetFocus

End If

End Sub

Private Sub DT5_KeyPress(KeyAscii As Integer)
KeyAscii = esNumero(KeyAscii)

If KeyAscii = 13 Then
DT5.text = Format(DT5.text, "00")
DT6.SetFocus

End If

End Sub
Private Sub DT6_KeyPress(KeyAscii As Integer)
KeyAscii = esNumero(KeyAscii)

If KeyAscii = 13 Then
DT6.text = Format(DT6.text, "0000")

End If

End Sub

Private Sub Form_Activate()
Command1_Click
Command2_Click

End Sub

    Private Sub Form_Load()
        Call Centrar(Me)
        'cmbTipo.ListIndex = 0
        Call CARGAGRILLA(1, 7)
    End Sub
    
    
'****************************************************************************
'Manejo de los Controles
'****************************************************************************

Sub CARGAGRILLA(ByVal row As Long, ByVal col As Long)
    Dim i As Long
    Rem DATOS DE LA COLUMNA
'    Informe.DefaultFont.Size = 7.5
    formatogrilla(1, 1) = "CODIGO"
    formatogrilla(1, 2) = "CANTIDAD"
    formatogrilla(1, 3) = "DESCRIPCION"
    formatogrilla(1, 4) = "P/UNITARIO"
    formatogrilla(1, 5) = "DESCUENTO"
    formatogrilla(1, 6) = "TOTAL"
    
    
    Rem LARGO DE LOS DATOS
    formatogrilla(2, 1) = "10"
    formatogrilla(2, 2) = "50"
    formatogrilla(2, 3) = "10"
    formatogrilla(2, 4) = "7"
    formatogrilla(2, 5) = "9"
    formatogrilla(2, 6) = "8"
    formatogrilla(2, 7) = "8"
    formatogrilla(2, 8) = "8"
    formatogrilla(2, 9) = "8"
    
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
    
    Rem FORMATO GRILLA
    formatogrilla(4, 1) = ""
    formatogrilla(4, 2) = ""
    formatogrilla(4, 3) = "###,##0.000"
    formatogrilla(4, 4) = "###,###,###"
    formatogrilla(4, 5) = "###,###,###"
    formatogrilla(4, 6) = "###,###,###"
    
    Rem LOCCKED
    formatogrilla(5, 1) = "TRUE"
    formatogrilla(5, 2) = "TRUE"
    formatogrilla(5, 3) = "TRUE"
    formatogrilla(5, 4) = "TRUE"
    formatogrilla(5, 5) = "TRUE"
    formatogrilla(5, 6) = "TRUE"
    formatogrilla(5, 7) = "FALSE"
    formatogrilla(5, 8) = "FALSE"
    formatogrilla(5, 9) = "FALSE"
    
    Rem ANCHO
    formatogrilla(6, 1) = "13"
    formatogrilla(6, 2) = "10"
    formatogrilla(6, 3) = "50"
    formatogrilla(6, 4) = "10"
    formatogrilla(6, 5) = "10"
    formatogrilla(6, 6) = "10"
    formatogrilla(6, 7) = "8"
    formatogrilla(6, 8) = "8"
    formatogrilla(6, 9) = "8"
    
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
        Informe.Cell(0, i).text = formatogrilla(1, i)
        Informe.Column(i).Width = Val(formatogrilla(6, i)) * Informe.DefaultFont.Size
        Informe.Column(i).MaxLength = Val(formatogrilla(2, i))
        Informe.Column(i).FormatString = formatogrilla(4, i)
        Informe.Column(i).Locked = formatogrilla(5, i)
        If formatogrilla(3, i) = "N" Then Informe.Column(i).Alignment = cellRightCenter
        If formatogrilla(3, i) = "S" Then Informe.Column(i).Alignment = cellLeftCenter
        If formatogrilla(3, i) = "D" Then Informe.Column(i).CellType = cellCalendar
    Next i
    Informe.Range(0, 0, 0, Informe.Cols - 1).Alignment = cellCenterCenter
End Sub

Private Sub cargadocumento(loc, TIPO, NUMERO, caja, fecha)
    
    Dim resultados As rdoResultset
    Dim csql As New rdoQuery
    Set csql.ActiveConnection = ventasRubro
    
    
    csql.sql = "SELECT dd.codigo, dd.cantidad,dd.descripcion,   dd.precio, dd.descuento,dd.total , dc.rut, dc.sucursal, dc.neto as neto, dc.iva, dc.impuestoharina , dc.impuestocarne , dc.impuestoilarefrescos , dc.impuestoilalicores , dc.impuestoilavinos , dc.total, dc.fecha ,dc.descuento,dc.foliosii,dc.donacion ,dc.numero,mc.nombre,mc.direccion,mc.ciudad,mc.giro "
    csql.sql = csql.sql & "from " + clientesistema + "ventas" + loc + ".sv_documento_cabeza_" + loc + " AS dc, " + clientesistema + "ventas" + loc + ".sv_documento_detalle_" + loc + " AS dd, " + clientesistema + "ventas.sv_maestroclientes as mc "
    csql.sql = csql.sql & "WHERE dc.caja='" + caja + "' and dc.local = '" & loc & "' AND dc.local = dd.local AND dd.tipo = '" + TIPO + "' AND dc.numero = '" & NUMERO & "' AND dd.caja=dc.caja and dd.tipo = dc.tipo AND dd.numero = dc.numero and dc.rut=mc.rut and dd.fecha=dc.fecha and mc.sucursal='0'   ORDER BY dd.linea ASC "
    csql.Execute
Informe.Rows = 1
        If csql.RowsAffected > 0 Then
           
           
           
           Set resultados = csql.OpenResultset
        
        
        lblRut.Caption = Format(Mid(resultados(6), 1, 9), "#########") + "-" + Mid(resultados(6), 10, 1)
        lblnombre.Caption = resultados(21)
        LBLDIRECCION.Caption = resultados(22)
        LBLCIUDAD.Caption = resultados(23)
        LBLGIRO.Caption = resultados(24)
        lblneto.Caption = Format(resultados("neto"), "###,###,###")
        lbliva.Caption = Format(resultados("iva"), "###,###,###")
        lbltotal.Caption = Format(resultados(15), "###,###,###")
        lblotros.Caption = Format(resultados(14) + resultados(13) + resultados(12) + resultados(11) + resultados(10), "###,###,###")
        
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
    
    
End Sub



Private Sub Label4_Click()

End Sub

Private Sub Text4_Change()

End Sub

