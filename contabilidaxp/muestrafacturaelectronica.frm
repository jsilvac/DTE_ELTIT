VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{4777436C-EB8C-4596-98A8-EBCF98683969}#1.0#0"; "Flexcell.ocx"
Object = "{ADBBAED6-B16F-40DC-80DF-B44910CBA76C}#1.0#0"; "Frame-Xp.ocx"
Begin VB.Form factu01 
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
         Left            =   2790
         TabIndex        =   41
         Top             =   945
         Width           =   7620
      End
      Begin VB.Label LBLGIRO 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   945
         TabIndex        =   40
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
      Left            =   135
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
         Top             =   855
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
Attribute VB_Name = "factu01"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Private formatogrilla(20, 20)
    Private fecha1 As String
    Private fecha2 As String
    
Private Sub Command1_Click()
Call cargadocumento(DT0.text, DT1.text, DT3.text, DT2.text, DT6.text + "-" + DT5.text + "-" + DT4.text)

End Sub

'****************************************************************************
'Manejo de los Controles
'****************************************************************************
    '****************************************************************************
    'GOTFOCUS
    '****************************************************************************
    Private Sub DATO1_GotFocus()
        Call cargatexto(dato1)
    End Sub

    Private Sub dato2_GotFocus()
        Call cargatexto(dato2)
    End Sub
    
    Private Sub dato3_GotFocus()
        Call cargatexto(dato3)
    End Sub
    
    Private Sub dato4_GotFocus()
        Call cargatexto(dato4)
    End Sub
    
    Private Sub dato5_GotFocus()
        Call cargatexto(dato5)
    End Sub
    
    Private Sub dato6_GotFocus()
        Call cargatexto(dato6)
    End Sub
    '****************************************************************************
    'GOTFOCUS
    '****************************************************************************
    
    '****************************************************************************
    'KEYDOWN
    '****************************************************************************
    Private Sub DATO1_KeyDown(KeyCode As Integer, Shift As Integer)
        Call Flechas(dato1, KeyCode, dato1)
    End Sub

    Private Sub dato2_KeyDown(KeyCode As Integer, Shift As Integer)
        Call Flechas(dato2, KeyCode, dato1)
    End Sub
    
    Private Sub dato3_KeyDown(KeyCode As Integer, Shift As Integer)
        Call Flechas(dato3, KeyCode, dato2)
    End Sub
    
    Private Sub dato4_KeyDown(KeyCode As Integer, Shift As Integer)
        Call Flechas(dato4, KeyCode, dato3)
    End Sub
    
    Private Sub dato5_KeyDown(KeyCode As Integer, Shift As Integer)
        Call Flechas(dato5, KeyCode, dato4)
    End Sub
    
    Private Sub dato6_KeyDown(KeyCode As Integer, Shift As Integer)
        Call Flechas(dato6, KeyCode, dato5)
    End Sub
    '****************************************************************************
    'KEYDOWN
    '****************************************************************************
    
    '****************************************************************************
    'KEYPRESS
    '****************************************************************************
    Private Sub DATO1_KeyPress(KeyAscii As Integer)
        KeyAscii = esNumero(KeyAscii)
        If KeyAscii = 13 Then
            Call ceros(dato1)
            If dato1.text = "00" Then
                dato1.text = Format(fechasistema, "dd")
            End If
            SendKeys "{Tab}"
        End If
    End Sub
    
    Private Sub dato2_KeyPress(KeyAscii As Integer)
        KeyAscii = esNumero(KeyAscii)
        If KeyAscii = 13 Then
            Call ceros(dato2)
            If dato2.text = "00" Then
                dato2.text = Format(fechasistema, "mm")
            End If
            SendKeys "{Tab}"
        End If
    End Sub
    
    Private Sub dato3_KeyPress(KeyAscii As Integer)
        KeyAscii = esNumero(KeyAscii)
        If KeyAscii = 13 Then
            Call ceros(dato3)
            If dato3.text = "0000" Then
                dato3.text = Format(fechasistema, "yyyy")
            End If
            fecha1 = dato3.text & "-" & dato2.text & "-" & dato1.text
            SendKeys "{Tab}"
        End If
    End Sub
    
    Private Sub dato4_KeyPress(KeyAscii As Integer)
        KeyAscii = esNumero(KeyAscii)
        If KeyAscii = 13 Then
            Call ceros(dato4)
            SendKeys "{Tab}"
        End If
    End Sub
    
    Private Sub dato5_KeyPress(KeyAscii As Integer)
        KeyAscii = esNumero(KeyAscii)
        If KeyAscii = 13 Then
            Call ceros(dato5)
            SendKeys "{Tab}"
        End If
    End Sub
    
    Private Sub dato6_KeyPress(KeyAscii As Integer)
        KeyAscii = esNumero(KeyAscii)
        If KeyAscii = 13 Then
            Call ceros(dato6)
            fecha2 = dato6.text & "-" & dato5.text & "-" & dato4.text
            Call cargaInforme
            SendKeys "{Tab}"
        End If
    End Sub
    '****************************************************************************
    'KEYPRESS
    '****************************************************************************

Private Sub COMMAND2_Click()
Dim TIMBRE As String
TIMBRE = generatimbre("1.0", "775753005", "33", DT3.text, DT6.text + "-" + DT5.text + "-" + DT4.text, LBLRUT.Caption, lblnombre.Caption, LBLTOTAL.Caption, Informe.Cell(1, 2).text, "rooodododododododoodododododododo", TimeStamp, "dasdasdasdasds")
CARGAGRILLAFACTURA

End Sub
Sub CARGAGRILLAFACTURA()
Dim titulos(10, 2) As String

Grid1.Rows = 59
Grid1.Cols = 15
Grid1.Column(0).Width = 0

For K = 1 To 14
Grid1.Column(K).Width = 50
Next K
Grid1.PageSetup.PrintGridlines = True
Grid1.PageSetup.LeftMargin = 1
Grid1.PageSetup.TopMargin = 1
Grid1.PageSetup.RightMargin = 1
Grid1.PageSetup.BottomMargin = 1

titulos(1, 1) = "FERRETERIA ELTIT LIMITADA": titulos(1, 2) = 11
titulos(2, 1) = "GIRO: VENTA AL POR MAYOR DE VEHICULOS AUTOMOTORES EXCEPTO MOTOCICLETAS,": titulos(2, 2) = 6
titulos(3, 1) = "VENTA DE MOTOCICLETAS, VENTA AL POR MENOR DE ARTICULOS DE FERRETERIA Y": titulos(3, 2) = 6
titulos(4, 1) = "MATERIALES DE CONSTRUCCION": titulos(4, 2) = 6
titulos(5, 1) = "CASA MATRIX : OHIGGINS 336 - FONO : 45-411103 - FAX 45-411186 CASILLA 4-D PUCON ": titulos(5, 2) = 6
titulos(6, 1) = "SUCURSAL 1 : MIGUEL ANSORENA 365 : PUCON ": titulos(6, 2) = 6
titulos(7, 1) = "SUCURSAL 2: CAMINO INTERNACIONAL 200 - 2do nivel": titulos(7, 2) = 6
titulos(8, 1) = "COMUNA PUCON": titulos(8, 2) = 6
titulos(10, 1) = "RUT : 77.575.300-5"



For K = 1 To 8


Grid1.Range(K, 1, K, 7).Merge
Grid1.Range(K, 1, K, 7).Alignment = cellCenterCenter
Grid1.Range(K, 1, K, 7).FontSize = CDbl(titulos(K, 2))
Grid1.RowHeight(K) = 14


If K = 1 Then
Grid1.Cell(K, 1).Font.Bold = True
End If


Grid1.Cell(K, 1).text = titulos(K, 1)
Next K
For K = 1 To 10

Grid1.RowHeight(K) = 14

Next K
Grid1.Range(1, 9, 3, 14).Merge
Grid1.Cell(1, 9).Alignment = cellCenterCenter
Grid1.Cell(1, 9).Font.Size = 14
Grid1.Cell(1, 9).text = titulos(10, 1)


Grid1.Range(4, 9, 6, 14).Merge
Grid1.Cell(4, 9).Alignment = cellCenterCenter
Grid1.Cell(4, 9).Font.Size = 14
Grid1.Cell(4, 9).text = "FACTURA"

Grid1.Range(7, 9, 9, 14).Merge
Grid1.Cell(7, 9).Alignment = cellCenterCenter
Grid1.Cell(7, 9).Font.Size = 14
Grid1.Cell(7, 9).text = "N   " + DT3.text

Grid1.Range(1, 9, 1, 14).Borders(cellEdgeTop) = cellThick
For K = 2 To 7
Grid1.Range(K, 9, K, 14).Borders(cellEdgeLeft) = cellThick
Grid1.Range(K, 9, K, 14).Borders(cellEdgeRight) = cellThick

Next K

Grid1.Range(8, 9, 8, 14).Borders(cellEdgeBottom) = cellThick

Rem NOMBRE
Grid1.Range(9, 9, 9, 14).Merge



Grid1.Range(12, 1, 12, 2).Merge
Grid1.Range(12, 3, 12, 9).Merge
Grid1.Cell(12, 1).Alignment = cellLeftCenter
Grid1.Cell(12, 3).Alignment = cellLeftCenter
Grid1.Cell(12, 3).Font.Bold = True
Grid1.Cell(12, 1).text = "SEÑOR(ES)    :"
Grid1.Cell(12, 3).text = lblnombre.Caption
Rem DIRECCION

Grid1.Range(13, 1, 13, 2).Merge
Grid1.Range(13, 3, 13, 9).Merge
Grid1.Cell(13, 1).Alignment = cellLeftCenter
Grid1.Cell(13, 3).Alignment = cellLeftCenter
Grid1.Cell(13, 3).Font.Bold = True
Grid1.Cell(13, 1).text = "DIRECCION     :"
Grid1.Cell(13, 3).text = LBLDIRECCION.Caption

Rem GIRO

Grid1.Range(14, 1, 14, 2).Merge
Grid1.Range(14, 3, 14, 9).Merge
Grid1.Cell(14, 1).Alignment = cellLeftCenter
Grid1.Cell(14, 3).Alignment = cellLeftCenter
Grid1.Cell(14, 3).Font.Bold = True
Grid1.Cell(14, 1).text = "GIRO          :"
Grid1.Cell(14, 3).text = LBLGIRO.Caption

Rem FECHA

Grid1.Range(12, 10, 12, 11).Merge
Grid1.Range(12, 12, 12, 14).Merge
Grid1.Cell(12, 10).Alignment = cellLeftCenter
Grid1.Cell(12, 12).Alignment = cellLeftCenter
Grid1.Cell(12, 12).Font.Bold = True
Grid1.Cell(12, 10).text = "FECHA : "
Grid1.Cell(12, 12).text = DT4.text + "-" + DT5.text + "-" + DT6.text

Rem RUT

Grid1.Range(13, 10, 13, 11).Merge
Grid1.Range(13, 12, 13, 14).Merge
Grid1.Cell(13, 10).Alignment = cellLeftCenter
Grid1.Cell(13, 12).Alignment = cellLeftCenter
Grid1.Cell(13, 12).Font.Bold = True
Grid1.Cell(13, 10).text = "RUT:"
Grid1.Cell(13, 12).text = LBLRUT.Caption

Rem CIUDAD

Grid1.Range(14, 10, 14, 11).Merge
Grid1.Range(14, 12, 14, 14).Merge
Grid1.Cell(14, 10).Alignment = cellLeftCenter
Grid1.Cell(14, 12).Alignment = cellLeftCenter
Grid1.Cell(14, 12).Font.Bold = True
Grid1.Cell(14, 10).text = "CIUDAD :"
Grid1.Cell(14, 12).text = LBLCIUDAD.Caption


Rem DATOS


Grid1.Range(12, 1, 12, 14).Borders(cellEdgeTop) = cellThick
For K = 12 To 14
Grid1.Range(K, 1, K, 14).Borders(cellEdgeLeft) = cellThick
Grid1.Range(K, 1, K, 14).Borders(cellEdgeRight) = cellThick
Grid1.Range(K, 1, K, 14).BackColor = &HD8D8D8


Next K

Grid1.Range(14, 1, 14, 14).Borders(cellEdgeBottom) = cellThick


Rem DETALLE




For K = 16 To 50
Grid1.Range(K, 1, K, 2).Merge
Grid1.Range(K, 3, K, 3).Merge
Grid1.Range(K, 4, K, 9).Merge

Grid1.Range(K, 10, K, 11).Merge
Grid1.Range(K, 12, K, 12).Merge

Grid1.Range(K, 13, K, 14).Merge
Grid1.Range(K, 1, K, 14).Borders(cellEdgeLeft) = cellThin
Grid1.Range(K, 1, K, 14).Borders(cellEdgeRight) = cellThin







Next K

Grid1.Range(16, 1, 16, 14).Borders(cellEdgeTop) = cellThin
Grid1.Range(50, 1, 50, 14).Borders(cellEdgeBottom) = cellThin
Grid1.Range(16, 1, 50, 14).Borders(cellInsideHorizontal) = cellThin
Grid1.Range(16, 1, 50, 14).Borders(cellInsideVertical) = cellThin







Grid1.Cell(16, 1).Alignment = cellCenterCenter
Grid1.Cell(16, 1).Font.Bold = True
Grid1.Cell(16, 1).text = "CODIGO"

Grid1.Cell(16, 3).Alignment = cellCenterCenter
Grid1.Cell(16, 3).Font.Bold = True
Grid1.Cell(16, 3).text = "CANT."

Grid1.Cell(16, 4).Alignment = cellCenterCenter
Grid1.Cell(16, 4).Font.Bold = True
Grid1.Cell(16, 4).text = "DESCRIPCION"

Grid1.Cell(16, 10).Alignment = cellCenterCenter
Grid1.Cell(16, 10).Font.Bold = True
Grid1.Cell(16, 10).text = "P.UNITARIO"

Grid1.Cell(16, 12).Alignment = cellCenterCenter
Grid1.Cell(16, 12).Font.Bold = True
Grid1.Cell(16, 12).text = "% DCTO"

Grid1.Cell(16, 13).Alignment = cellCenterCenter
Grid1.Cell(16, 13).Font.Bold = True
Grid1.Cell(16, 13).text = "TOTAL "

Grid1.Range(16, 1, 16, 14).BackColor = &HE0E0E0



For K = 1 To Informe.Rows - 1
Grid1.Cell(K + 16, 1).Alignment = cellLeftCenter
Grid1.Cell(K + 16, 1).text = Informe.Cell(K, 1).text

Grid1.Cell(K + 16, 3).Alignment = cellRightCenter
Grid1.Cell(K + 16, 3).text = Informe.Cell(K, 2).text

Grid1.Cell(K + 16, 4).Alignment = cellLeftCenter
Grid1.Cell(K + 16, 4).text = Informe.Cell(K, 3).text

Grid1.Cell(K + 16, 10).Alignment = cellRightCenter
Grid1.Cell(K + 16, 10).text = Format(Informe.Cell(K, 4).text, "###,###,###,###")

Grid1.Cell(K + 16, 12).Alignment = cellRightCenter
Grid1.Cell(K + 16, 12).text = Informe.Cell(K, 5).text

Grid1.Cell(K + 16, 13).Alignment = cellRightCenter
Grid1.Cell(K + 16, 13).text = Format(Informe.Cell(K, 6).text, "###,###,###,###")

Next K
For K = 51 To 55
Grid1.Range(K, 10, K, 12).Merge
Grid1.Range(K, 10, K, 12).Alignment = cellLeftBottom
Grid1.Range(K, 10, K, 12).FontBold = True
Grid1.Range(K, 10, K, 12).Borders(cellEdgeLeft) = cellThin
Grid1.Range(K, 10, K, 12).Borders(cellEdgeRight) = cellThin



Grid1.Range(K, 13, K, 14).Merge
Grid1.Range(K, 13, K, 14).Alignment = cellRightBottom
Grid1.Range(K, 13, K, 14).FontBold = True
Grid1.Range(K, 13, K, 14).FontSize = 10
Grid1.Range(K, 13, K, 14).Borders(cellEdgeLeft) = cellThin
Grid1.Range(K, 13, K, 14).Borders(cellEdgeRight) = cellThin

Next K
Grid1.Range(51, 10, 51, 14).Borders(cellEdgeTop) = cellThin
Grid1.Range(55, 10, 55, 14).Borders(cellEdgeBottom) = cellThin


Grid1.Cell(51, 10).text = "NETO"
Grid1.Cell(52, 10).text = "IVA "
Grid1.Cell(53, 10).text = "EXENTO"
Grid1.Cell(54, 10).text = "OT.IMPTOS"
Grid1.Cell(55, 10).text = "TOTAL "

Grid1.Cell(51, 13).text = Format(lblneto.Caption, "###,###,###,###")
Grid1.Cell(52, 13).text = Format(lbliva.Caption, "###,###,###,###")
Grid1.Cell(53, 13).text = Format(lblexento.Caption, "###,###,###,###")
Grid1.Cell(54, 13).text = Format(lblotros.Caption, "###,###,###,###")
Grid1.Cell(55, 13).text = Format(LBLTOTAL.Caption, "###,###,###,###")

Rem inserta imagen
Grid1.Range(51, 1, 55, 9).Merge


Grid1.Images.Add "C:\TIMBRE.JPEG", "SAMPLE"
Grid1.Cell(51, 1).SetImage "SAMPLE"



Grid1.PageSetup.PrintGridlines = False


Grid1.PrintPreview




End Sub
Private Sub DT0_KeyPress(KeyAscii As Integer)
KeyAscii = esNumero(KeyAscii)
If KeyAscii = 13 Then
DT0.text = Format(DT0.text, "00")
If leerLocal(DT0.text) <> "" Then
lbllocal.Caption = leerLocal(DT0.text)
DT1.SetFocus
Else
MsgBox ("NUMERO DE LOCAL NO EXISTE ")
DT0.SetFocus

End If


End If

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

    Private Sub Form_Load()
        Call Centrar(Me)
        'cmbTipo.ListIndex = 0
        Call CARGAGRILLA(1, 7)
    End Sub
    
    Private Sub cmdImprime_Click()
        Call imprimeInforme
    End Sub
'****************************************************************************
'Manejo de los Controles
'****************************************************************************

Sub CARGAGRILLA(ByVal Row As Long, ByVal Col As Long)
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
    
    Informe.Cols = Col
    Informe.Rows = Row
    
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

Private Sub cargadocumento(loc, tipo, numero, caja, fecha)
    
    Dim resultados As rdoResultset
    Dim cSql As New rdoQuery
    Set cSql.ActiveConnection = DTE
    
    cSql.sql = "SELECT dd.codigo, dd.cantidad,dd.descripcion,   dd.precio, dd.descuento,dd.total , dc.rut, dc.sucursal, dc.neto as neto, dc.iva, dc.impuestoharina , dc.impuestocarne , dc.impuestoilarefrescos , dc.impuestoilalicores , dc.impuestoilavinos , dc.total, dc.fecha ,dc.descuento,dc.foliosii,dc.donacion ,dc.numero,mc.nombre,mc.direccion "
    cSql.sql = cSql.sql & "from " + clientesistema + "ventas" + loc + ".sv_documento_cabeza_" + loc + " AS dc, " + clientesistema + "ventas" + loc + ".sv_documento_detalle_" + loc + " AS dd, " + clientesistema + "ventas.sv_maestroclientes as mc "
    cSql.sql = cSql.sql & "WHERE dc.caja='" + caja + "' and dc.local = '" & loc & "' AND dc.local = dd.local AND dd.tipo = '" + tipo + "' AND dc.numero = '" & numero & "' AND dd.caja=dc.caja and dd.tipo = dc.tipo AND dd.numero = dc.numero and dc.rut=mc.rut ORDER BY dd.linea ASC "
    cSql.Execute
Informe.Rows = 1
        If cSql.RowsAffected > 0 Then
           
           
           
           Set resultados = cSql.OpenResultset
        LBLRUT.Caption = Mid(resultados(6), 1, 9) + "-" + Mid(resultados(6), 10, 1)
        lblnombre.Caption = resultados(21)
        LBLDIRECCION.Caption = resultados(22)
        lblneto.Caption = Format(resultados("neto"), "###,###,###")
        lbliva.Caption = Format(resultados("iva"), "###,###,###")
        LBLTOTAL.Caption = Format(resultados(15), "###,###,###")
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

