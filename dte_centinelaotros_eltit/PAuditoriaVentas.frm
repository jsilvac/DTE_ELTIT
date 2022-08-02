VERSION 5.00
Object = "{4777436C-EB8C-4596-98A8-EBCF98683969}#1.0#0"; "FlexCell.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{ADBBAED6-B16F-40DC-80DF-B44910CBA76C}#1.0#0"; "Frame-Xp.ocx"
Begin VB.Form PAuditoriaVentas 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Auditoria de Ventas"
   ClientHeight    =   9345
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11520
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9345
   ScaleWidth      =   11520
   Begin VB.OptionButton Option2 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Imprime Resumen Solamente"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   360
      TabIndex        =   31
      Top             =   9000
      Value           =   -1  'True
      Width           =   3975
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Imprime Detalle de Documentos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   360
      TabIndex        =   30
      Top             =   8640
      Width           =   3975
   End
   Begin XPFrame.FrameXp FrameXp2 
      Height          =   1320
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   11415
      _ExtentX        =   20135
      _ExtentY        =   2328
      BackColor       =   16744576
      Caption         =   "Periodo"
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
      Begin VB.CommandButton Command1 
         BackColor       =   &H0080FF80&
         Caption         =   "PROCESAR"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   3840
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   1080
         Width           =   3075
      End
      Begin VB.TextBox tcaja 
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
         Left            =   7470
         MaxLength       =   2
         TabIndex        =   25
         Top             =   720
         Width           =   420
      End
      Begin VB.TextBox tcajera 
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
         Left            =   1530
         MaxLength       =   9
         TabIndex        =   22
         Top             =   720
         Width           =   1320
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
         Left            =   1500
         MaxLength       =   2
         TabIndex        =   0
         Top             =   360
         Width           =   375
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
         Left            =   6720
         MaxLength       =   2
         TabIndex        =   1
         Tag             =   "proveedor"
         Top             =   360
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
         Left            =   7080
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
         Left            =   7440
         MaxLength       =   4
         TabIndex        =   3
         Tag             =   "proveedor"
         Top             =   360
         Width           =   735
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
         Left            =   10440
         MaxLength       =   4
         TabIndex        =   6
         Tag             =   "proveedor"
         Top             =   360
         Width           =   735
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
         Left            =   9720
         MaxLength       =   2
         TabIndex        =   4
         Tag             =   "proveedor"
         Top             =   360
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
         Left            =   10080
         MaxLength       =   2
         TabIndex        =   5
         Tag             =   "proveedor"
         Top             =   360
         Width           =   375
      End
      Begin VB.Label lbllocal 
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
         Left            =   1920
         TabIndex        =   32
         Top             =   360
         Width           =   3255
      End
      Begin VB.Label lbldv 
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
         TabIndex        =   29
         Top             =   720
         Width           =   240
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Caja"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   6615
         TabIndex        =   27
         Top             =   720
         Width           =   840
      End
      Begin VB.Label nombrecaja 
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
         Left            =   7965
         TabIndex        =   26
         Top             =   720
         Width           =   3255
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Cajera"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   45
         TabIndex        =   24
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label nombrecajera 
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
         Left            =   3285
         TabIndex        =   23
         Top             =   720
         Width           =   3255
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Local"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   60
         TabIndex        =   18
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label lbl3 
         Appearance      =   0  'Flat
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Hasta"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   8280
         TabIndex        =   15
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label lbl2 
         Appearance      =   0  'Flat
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Desde"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   5280
         TabIndex        =   14
         Top             =   360
         Width           =   1335
      End
   End
   Begin XPFrame.FrameXp FrameXp1 
      Height          =   2415
      Left            =   0
      TabIndex        =   11
      Top             =   1320
      Width           =   11415
      _ExtentX        =   20135
      _ExtentY        =   4260
      BackColor       =   16744576
      Caption         =   "Resumen Documentos"
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
      Begin FlexCell.Grid Documentos 
         Height          =   1950
         Left            =   120
         TabIndex        =   7
         Top             =   360
         Width           =   11175
         _ExtentX        =   19711
         _ExtentY        =   3440
         Cols            =   20
         DefaultFontSize =   8.25
         Rows            =   7
         SelectionMode   =   1
         DateFormat      =   2
      End
   End
   Begin XPFrame.FrameXp FrameXp3 
      Height          =   4785
      Left            =   0
      TabIndex        =   12
      Top             =   3720
      Width           =   11400
      _ExtentX        =   20108
      _ExtentY        =   8440
      BackColor       =   16744576
      Caption         =   "Resumen Ingresos"
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
      Begin FlexCell.Grid Ingresos 
         Height          =   4380
         Left            =   90
         TabIndex        =   8
         Top             =   360
         Width           =   11250
         _ExtentX        =   19844
         _ExtentY        =   7726
         Cols            =   2
         DefaultFontSize =   8.25
         Rows            =   7
         SelectionMode   =   1
         DateFormat      =   2
      End
   End
   Begin XPFrame.FrameXp FrameXp4 
      Height          =   120
      Left            =   9765
      TabIndex        =   13
      Top             =   8415
      Visible         =   0   'False
      Width           =   1665
      _ExtentX        =   2937
      _ExtentY        =   212
      BackColor       =   16744576
      Caption         =   "Resumen Egresos"
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
      Begin FlexCell.Grid Egresos 
         Height          =   4065
         Left            =   120
         TabIndex        =   9
         Top             =   360
         Width           =   5295
         _ExtentX        =   9340
         _ExtentY        =   7170
         Cols            =   2
         DefaultFontSize =   8.25
         Rows            =   9
         SelectionMode   =   1
         DateFormat      =   2
      End
   End
   Begin XPFrame.FrameXp frmImprimir 
      Height          =   375
      Left            =   4455
      TabIndex        =   16
      Top             =   8880
      Visible         =   0   'False
      Width           =   1965
      _ExtentX        =   3466
      _ExtentY        =   661
      BackColor       =   49344
      Caption         =   "Imprimir"
      CaptionEstilo3D =   1
      BackColor       =   49344
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   1
   End
   Begin MSAdodcLib.Adodc data 
      Height          =   345
      Left            =   3360
      Top             =   5430
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   609
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   -1
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
   Begin MSAdodcLib.Adodc data2 
      Height          =   375
      Left            =   1200
      Top             =   5400
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
      LockType        =   -1
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
   Begin FlexCell.Grid Impresion 
      Height          =   1035
      Left            =   60
      TabIndex        =   17
      Top             =   1620
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   1826
      Cols            =   5
      DefaultFontSize =   8.25
      Rows            =   30
   End
   Begin XPFrame.FrameXp FrmNegativo 
      Height          =   375
      Left            =   2475
      TabIndex        =   19
      Top             =   6840
      Visible         =   0   'False
      Width           =   3585
      _ExtentX        =   6324
      _ExtentY        =   661
      BackColor       =   49344
      Caption         =   "Ventas Stock Negativo"
      CaptionEstilo3D =   1
      BackColor       =   49344
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   1
   End
   Begin VB.TextBox pivote 
      Height          =   285
      Left            =   45
      TabIndex        =   20
      Top             =   7245
      Visible         =   0   'False
      Width           =   960
   End
   Begin XPFrame.FrameXp frmventas 
      Height          =   375
      Left            =   6885
      TabIndex        =   21
      Top             =   8880
      Width           =   2100
      _ExtentX        =   3704
      _ExtentY        =   661
      BackColor       =   49344
      Caption         =   "Ventas x Horas"
      CaptionEstilo3D =   1
      BackColor       =   49344
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   1
   End
End
Attribute VB_Name = "PAuditoriaVentas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Private a As auditoria
    Private formatogrilla(10, 20) As String
    Private fecha1 As String
    Private fecha2 As String
    Public tabla As String
    Private TIPO As String
    Private salto As Boolean
    
Private Sub Command1_Click()
   
   nombrecajera.Caption = leerNombreCajera(tcajera.text + lbldv.Caption)
   nombrecaja.Caption = leerNombreCaja(tcaja.text)
   
   If leerauditoria(a, fecha1, fecha2) = True Then
                                
                frmImprimir.Visible = True
                Call structtoctrl
            End If
End Sub

'============================================================
'MANEJO DE LOS CONTOLES
'============================================================
    '========================================================
    'GotFocus
    '========================================================
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
    '========================================================
    'GotFocus
    '========================================================
    
    '========================================================
    'KeyDown
    '========================================================
    Private Sub dato1_KeyDown(KeyCode As Integer, Shift As Integer)
        Call Flechas(KeyCode, dato1)
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
    '========================================================
    'KeyDown
    '========================================================
    
    '========================================================
    'KeyPress
    '========================================================
    Private Sub dato1_KeyPress(KeyAscii As Integer)
        KeyAscii = esNumero(KeyAscii)
        If KeyAscii = 13 Then
            dato1.text = ceros(dato1)
            lbllocal.Caption = leerNombreEmpresa(dato1.text)
            If lbllocal.Caption <> "" Then
                rubroAuditoria = leerRubro(dato1.text)
                Call ConectarAuditoria(servidor, rubroAuditoria, usuario, password, dato1.text)
                
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
            fecha1 = dato4.text & "-" & dato3.text & "-" & dato2.text
            fecha2 = dato7.text & "-" & dato6.text & "-" & dato5.text
            fechaAuditIni = fecha1
            fechaAuditFin = fecha2
            localAuditoria = dato1.text
            Rem Command1_Click
        
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
    '========================================================
    'KeyUp
    '========================================================
'============================================================
'MANEJO DE LOS CONTOLES
'============================================================
    
    Private Sub dato2_LostFocus()
  
    Call esfecha(dato2, dato3, dato4, "dd")
    End Sub
    Private Sub dato3_LostFocus()
    Call esfecha(dato2, dato3, dato4, "mm")
    End Sub
    Private Sub dato4_LostFocus()
    Call esfecha(dato2, dato3, dato4, "yyyy")
    End Sub
    Private Sub dato5_LostFocus()
    Call esfecha(dato5, dato6, dato7, "dd")
    End Sub
    Private Sub dato6_LostFocus()
    Call esfecha(dato5, dato6, dato7, "mm")
    End Sub
    Private Sub dato7_LostFocus()
    Call esfecha(dato5, dato6, dato7, "yyyy")
    End Sub
    Private Sub Documentos_DblClick()

If dato1.text = Empty Then Exit Sub 'SE CAIA AL NO TENER TODOS LOS DATOS

        Select Case Documentos.ActiveCell.row
            Case 1
                TIPO = "FV"
            Case 2
                TIPO = "BV"
            Case 3
                TIPO = "ZE"
            Case 4
                TIPO = "FE"
            Case 5
                TIPO = "NB' or tipo='NF"
            Case Else
                TIPO = ""
        End Select
        tabla = "SELECT CONCAT(CONCAT(dc.tipo, ' ', dc.foliosii), '" & vbTab & "', DATE_FORMAT(dc.fecha,'%d-%m-%Y'), '" & vbTab & "', dc.rut) AS item1, CONCAT(CONCAT('$ ', FORMAT(dc.descuento,0)), '" & vbTab & "', CONCAT('$ ', FORMAT(dc.neto,0)), '" & vbTab & "', CONCAT('$ ', FORMAT(dc.iva,0)), '" & vbTab & "', CONCAT('$ ', FORMAT(IF(dc.retencionparcial=0,dc.retenciontotal,dc.retencionparcial),0)), '" & vbTab & "', CONCAT('$ ', FORMAT(dc.total,0))) AS item2, dc.rut, dc.descuento, dc.neto, dc.iva, IF(dc.retencionparcial=0, dc.retenciontotal, dc.retencionparcial) AS retencion, dc.total, dc.tipo,dc.caja "
        tabla = tabla & "FROM sv_documento_cabeza_" + localAuditoria + " AS dc "
        tabla = tabla & "WHERE local = '" & localAuditoria & "' AND (tipo = '" & TIPO & "') AND fecha BETWEEN '" & fecha1 & "' AND '" & fecha2 & "' AND nula = 'N' and dc.cajera like '%" & tcajera.text & "%' and caja like '%" & tcaja.text & "%' ORDER BY dc.numero ASC"
        Load listaDocumentos
        listaDocumentos.formulario = "auditoria"
        listaDocumentos.datos = "ventas"
        listaDocumentos.tabla = tabla
        listaDocumentos.Show vbModal
    End Sub

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
        Call Centrar(Me)
        Call CargaGrillaResumenDoc(7, 7)
        Call CargaGrillaIngresos(15, 2)
        Call CargaGrillaEgresos(15, 2)
        Call cargartipospagos
    End Sub
    
    Private Sub Form_Unload(Cancel As Integer)
        Principal.barraEstado.Panels(1).text = UCase(Principal.Caption)
        Call limpiaBarra(2)
    End Sub

'****************************************************************************
'Formato de la Grilla Documentos
'****************************************************************************
    Private Sub CargaGrillaResumenDoc(ByVal row As Integer, ByVal col As Integer)
        Dim i As Integer
        Rem DATOS DE LA COLUMNA
        formatogrilla(1, 1) = "CANTIDAD"
        formatogrilla(1, 2) = "DESCUENTO"
        formatogrilla(1, 3) = "NULAS"
        formatogrilla(1, 4) = "TOTAL"
        formatogrilla(1, 5) = "FOL.INICIAL"
        formatogrilla(1, 6) = "FOL.FINAL"
        
        Rem LARGO DE LOS DATOS
        formatogrilla(2, 1) = "5"
        formatogrilla(2, 2) = "15"
        formatogrilla(2, 3) = "3"
        formatogrilla(2, 4) = "9"
        formatogrilla(2, 5) = "10"
        formatogrilla(2, 6) = "10"
        
        Rem TIPO DE DATOS C=CEROS,S=STRING,N=NUMERICO,D=DATE,
        formatogrilla(3, 1) = "N"
        formatogrilla(3, 2) = "N"
        formatogrilla(3, 3) = "N"
        formatogrilla(3, 4) = "N"
        formatogrilla(3, 5) = "N"
        formatogrilla(3, 6) = "N"
        
        Rem FORMATO GRILLA
        formatogrilla(4, 1) = "###,###,##0"
        formatogrilla(4, 2) = "$ ###,###,##0"
        formatogrilla(4, 3) = "###,###,##0"
        formatogrilla(4, 4) = "$ ###,###,##0"
        formatogrilla(4, 5) = "0000000000"
        formatogrilla(4, 6) = "0000000000"
        
        Rem LOCCKED
        formatogrilla(5, 1) = "TRUE"
        formatogrilla(5, 2) = "TRUE"
        formatogrilla(5, 3) = "TRUE"
        formatogrilla(5, 4) = "TRUE"
        formatogrilla(5, 5) = "TRUE"
        formatogrilla(5, 6) = "TRUE"
        
        Rem VALOR MINIMO
        formatogrilla(6, 1) = ""
        formatogrilla(6, 2) = ""
        formatogrilla(6, 3) = ""
        formatogrilla(6, 4) = ""
        formatogrilla(6, 5) = ""
        formatogrilla(6, 6) = ""
        
        Rem VALOR MAXIMO
        formatogrilla(7, 1) = ""
        formatogrilla(7, 2) = ""
        formatogrilla(7, 3) = ""
        formatogrilla(7, 4) = ""
        formatogrilla(7, 5) = ""
        formatogrilla(7, 6) = ""
        
        Rem ANCHO
        formatogrilla(8, 0) = "7"
        formatogrilla(8, 1) = "7"
        formatogrilla(8, 2) = "15"
        formatogrilla(8, 3) = "12"
        formatogrilla(8, 4) = "12"
        formatogrilla(8, 5) = "12"
        formatogrilla(8, 6) = "12"
            
        Documentos.Cols = col
        Documentos.Rows = row
        Documentos.AllowUserResizing = False
        Documentos.DisplayFocusRect = False
        Documentos.ExtendLastCol = True
        Documentos.BoldFixedCell = False
        Documentos.DrawMode = cellOwnerDraw
        Documentos.Appearance = Flat
        Documentos.ScrollBarStyle = Flat
        Documentos.FixedRowColStyle = Flat
        Documentos.BackColorFixed = RGB(90, 158, 214)
        Documentos.BackColorFixedSel = RGB(110, 180, 230)
        Documentos.BackColorBkg = RGB(90, 158, 214)
        Documentos.BackColorScrollBar = RGB(231, 235, 247)
        Documentos.BackColor1 = RGB(231, 235, 247)
        Documentos.BackColor2 = RGB(239, 243, 255)
        Documentos.GridColor = RGB(148, 190, 231)
        
        Documentos.Column(0).Width = Val(formatogrilla(8, 0)) * (Documentos.Cell(0, 0).Font.Size + 1.25)
        Documentos.Column(0).Alignment = cellLeftCenter
        Documentos.Cell(1, 0).text = "FACTURAS"
        Documentos.Cell(2, 0).text = "BOLETAS"
        Documentos.Cell(3, 0).text = "FACTURAS ELEC."
        Documentos.Cell(4, 0).text = "BOLETAS ELEC.."
        Documentos.Cell(5, 0).text = "NOTAS CREDITO"
        Documentos.Cell(6, 0).text = "TOTALES"
        
        For i = 1 To col - 1
            Documentos.Cell(0, i).text = formatogrilla(1, i)
            Documentos.Column(i).Width = Val(formatogrilla(8, i)) * (Documentos.Cell(0, i).Font.Size + 1.25)
            Documentos.Column(i).MaxLength = Val(formatogrilla(2, i))
            Documentos.Column(i).FormatString = formatogrilla(4, i)
            Documentos.Column(i).Locked = formatogrilla(5, i)
            If formatogrilla(3, i) = "N" Then
                Documentos.Column(i).Alignment = cellRightCenter
            Else
                Documentos.Column(i).Alignment = cellLeftCenter
            End If
        Next i
        Documentos.Range(0, 1, 0, Documentos.Cols - 1).Alignment = cellCenterCenter
        Documentos.Range(5, 0, 5, Documentos.Cols - 1).BackColor = RGB(254, 173, 175)
        Documentos.Range(6, 0, 6, Documentos.Cols - 1).BackColor = RGB(200, 225, 250)
        
        'Documentos.BackColor1 = RGB(231, 235, 247)
        'Documentos.BackColor2 = RGB(239, 243, 255)
        Documentos.Enabled = True
    End Sub
'****************************************************************************
'Formato de la Grilla Documentos
'****************************************************************************

'****************************************************************************
'Formato de la Grilla Ingresos
'****************************************************************************
    Private Sub CargaGrillaIngresos(ByVal row As Integer, ByVal col As Integer)
        Dim i As Integer
        Rem DATOS DE LA COLUMNA
        formatogrilla(1, 1) = "TOTALES"
'        formatoGrilla(1, 1) = "TOTAL EFECTIVO"
'        formatoGrilla(1, 2) = "TOTAL CHEQUES PROPIOS"
'        formatoGrilla(1, 3) = "TOTAL CHEQUES TERCERO"
'        formatoGrilla(1, 4) = "TOTAL TARJETA CREDITO"
'        formatoGrilla(1, 5) = "TOTAL TARJETA REDBANC"
'        formatoGrilla(1, 6) = "TOTAL CREDITO DIN-ABC"
'        formatoGrilla(1, 7) = "TOTAL CREDITO RIPLEY"
'        formatoGrilla(1, 8) = "TOTAL CREDITO SKORPIOS"
'        formatoGrilla(1, 9) = "TOTAL CREDITO DIRECTO  "
'        formatoGrilla(1, 10) = "TOTAL NOTAS DE CREDITO "
'        formatoGrilla(1, 11) = "TOTAL INGRESO X VENTA "
'        formatoGrilla(1, 12) = "TOTAL PAGOS C.DIRECTO  "
'        formatoGrilla(1, 13) = "TOTAL PAGOS C.SKORPIOS "
'        formatoGrilla(1, 14) = "TOTAL GENERAL INGRESOS "
        
        
        Rem LARGO DE LOS DATOS
        formatogrilla(2, 1) = "9"
        formatogrilla(2, 2) = "9"
        
        Rem TIPO DE DATOS C=CEROS,S=STRING,N=NUMERICO,D=DATE,
        formatogrilla(3, 1) = "N"
        formatogrilla(3, 2) = "N"
        formatogrilla(3, 3) = "N"
        formatogrilla(3, 4) = "N"
        formatogrilla(3, 5) = "N"
        formatogrilla(3, 6) = "N"
        formatogrilla(3, 7) = "N"
        formatogrilla(3, 8) = "N"
        
        Rem FORMATO GRILLA
        formatogrilla(4, 1) = "$ ###,###,##0"
        formatogrilla(4, 2) = "$ ###,###,##0"
        
        Rem LOCCKED
        formatogrilla(5, 1) = "TRUE"
        formatogrilla(5, 2) = "TRUE"
        
        Rem VALOR MINIMO
        formatogrilla(6, 1) = ""
        formatogrilla(6, 2) = ""
        formatogrilla(6, 3) = ""
        formatogrilla(6, 4) = ""
        formatogrilla(6, 5) = ""
        formatogrilla(6, 6) = ""
        
        Rem VALOR MAXIMO
        formatogrilla(7, 1) = ""
        formatogrilla(7, 2) = ""
        formatogrilla(7, 3) = ""
        formatogrilla(7, 4) = ""
        formatogrilla(7, 5) = ""
        formatogrilla(7, 6) = ""
        
        Rem ANCHO
        formatogrilla(8, 1) = "20"
        formatogrilla(8, 2) = "10"
                
        Ingresos.Cols = col
        Ingresos.Rows = row
        Ingresos.AllowUserResizing = False
        Ingresos.DisplayFocusRect = False
        Ingresos.ExtendLastCol = True
        Ingresos.BoldFixedCell = False
        Ingresos.DrawMode = cellOwnerDraw
        Ingresos.Appearance = Flat
        Ingresos.ScrollBarStyle = Flat
        Ingresos.FixedRowColStyle = Flat
        Ingresos.BackColorFixed = RGB(90, 158, 214)
        Ingresos.BackColorFixedSel = RGB(110, 180, 230)
        Ingresos.BackColorBkg = RGB(90, 158, 214)
        Ingresos.BackColorScrollBar = RGB(231, 235, 247)
        Ingresos.BackColor1 = RGB(231, 235, 247)
        Ingresos.BackColor2 = RGB(239, 243, 255)
        Ingresos.GridColor = RGB(148, 190, 231)
        
        Ingresos.RowHeight(0) = 0
        Ingresos.Column(0).Width = Val(formatogrilla(8, 1)) * (Ingresos.Cell(0, 0).Font.Size + 1.25)
        Ingresos.Column(1).Width = Val(formatogrilla(8, 2)) * (Ingresos.Cell(0, 1).Font.Size + 1.25)
        Ingresos.Column(0).MaxLength = Val(formatogrilla(2, 1))
        Ingresos.Column(1).MaxLength = Val(formatogrilla(2, 2))
        Ingresos.Column(0).FormatString = formatogrilla(4, 1)
        Ingresos.Column(1).FormatString = formatogrilla(4, 2)
        Ingresos.Column(0).Locked = formatogrilla(5, 1)
        Ingresos.Column(1).Locked = formatogrilla(5, 2)
        If formatogrilla(3, 1) = "N" Then
            Ingresos.Column(0).Alignment = cellRightCenter
        Else
            Ingresos.Column(0).Alignment = cellLeftCenter
        End If
        If formatogrilla(3, 2) = "N" Then
            Ingresos.Column(1).Alignment = cellRightCenter
        Else
            Ingresos.Column(1).Alignment = cellLeftCenter
        End If
        
        For i = 1 To row - 1
            Ingresos.Cell(i, 0).text = formatogrilla(1, i)
        Next i
        Ingresos.Cell(1, 0).Alignment = cellLeftCenter
        Ingresos.Cell(2, 0).Alignment = cellLeftCenter
        Ingresos.Cell(3, 0).Alignment = cellLeftCenter
        Ingresos.Cell(4, 0).Alignment = cellLeftCenter
        Ingresos.Cell(5, 0).Alignment = cellLeftCenter
        Ingresos.Cell(6, 0).Alignment = cellLeftCenter
        Ingresos.Cell(7, 0).Alignment = cellLeftCenter
        Ingresos.Cell(8, 0).Alignment = cellLeftCenter
        Ingresos.Cell(9, 0).Alignment = cellLeftCenter
        Ingresos.Cell(10, 0).Alignment = cellLeftCenter
        Ingresos.Cell(11, 0).Alignment = cellLeftCenter
        Ingresos.Cell(12, 0).Alignment = cellLeftCenter
        Ingresos.Cell(13, 0).Alignment = cellLeftCenter
        Ingresos.Cell(14, 0).Alignment = cellLeftCenter
        
        Ingresos.Range(11, 0, 11, 1).BackColor = RGB(200, 225, 250)
        Ingresos.Range(14, 0, 14, 1).BackColor = RGB(200, 225, 250)
        
        Ingresos.Enabled = True
    End Sub
'****************************************************************************
'Formato de la Grilla Ingresos
'****************************************************************************

'****************************************************************************
'Formato de la Grilla Egresos
'****************************************************************************
    Private Sub CargaGrillaEgresos(ByVal row As Integer, ByVal col As Integer)
        Dim i As Integer
        Rem DATOS DE LA COLUMNA
        formatogrilla(1, 1) = "TOTAL EGRESOS DE CAJA"
        formatogrilla(1, 2) = "TOTAL CHEQUES A FECHA"
        formatogrilla(1, 3) = "TOTAL VENTAS CON CREDITO"
        formatogrilla(1, 4) = "TOTAL DEPOSITOS BANCO"
        formatogrilla(1, 5) = "TOTAL EGRESOS"
        formatogrilla(1, 6) = "TOTAL EFECTIVO A RENDIR"
        formatogrilla(1, 7) = "TOTAL CHEQUES A RENDIR"
        formatogrilla(1, 8) = "TOTAL DEPOSITOS LOCAL"
        For i = 9 To 14
        formatogrilla(1, i) = ""
        Next i
        
        Rem LARGO DE LOS DATOS
        formatogrilla(2, 1) = "9"
        formatogrilla(2, 2) = "9"
        
        Rem TIPO DE DATOS C=CEROS,S=STRING,N=NUMERICO,D=DATE,
        formatogrilla(3, 1) = "N"
        formatogrilla(3, 2) = "N"
        formatogrilla(3, 3) = "N"
        formatogrilla(3, 4) = "N"
        formatogrilla(3, 5) = "N"
        formatogrilla(3, 6) = "N"
        
        Rem FORMATO GRILLA
        formatogrilla(4, 1) = "$ ###,###,##0"
        formatogrilla(4, 2) = "$ ###,###,##0"
        
        Rem LOCCKED
        formatogrilla(5, 1) = "TRUE"
        formatogrilla(5, 2) = "TRUE"
        
        Rem VALOR MINIMO
        formatogrilla(6, 1) = ""
        formatogrilla(6, 2) = ""
        formatogrilla(6, 3) = ""
        formatogrilla(6, 4) = ""
        formatogrilla(6, 5) = ""
        formatogrilla(6, 6) = ""
        formatogrilla(6, 7) = ""
        formatogrilla(6, 8) = ""
        
        Rem VALOR MAXIMO
        formatogrilla(7, 1) = ""
        formatogrilla(7, 2) = ""
        formatogrilla(7, 3) = ""
        formatogrilla(7, 4) = ""
        formatogrilla(7, 5) = ""
        formatogrilla(7, 6) = ""
        formatogrilla(7, 7) = ""
        formatogrilla(7, 8) = ""
        
        Rem ANCHO
        formatogrilla(8, 1) = "20"
        formatogrilla(8, 2) = "10"
                
        Egresos.Cols = col
        Egresos.Rows = row
        Egresos.AllowUserResizing = False
        Egresos.DisplayFocusRect = False
        Egresos.ExtendLastCol = True
        Egresos.BoldFixedCell = False
        Egresos.DrawMode = cellOwnerDraw
        Egresos.Appearance = Flat
        Egresos.ScrollBarStyle = Flat
        Egresos.FixedRowColStyle = Flat
        Egresos.BackColorFixed = RGB(90, 158, 214)
        Egresos.BackColorFixedSel = RGB(110, 180, 230)
        Egresos.BackColorBkg = RGB(90, 158, 214)
        Egresos.BackColorScrollBar = RGB(231, 235, 247)
        Egresos.BackColor1 = RGB(231, 235, 247)
        Egresos.BackColor2 = RGB(239, 243, 255)
        Egresos.GridColor = RGB(148, 190, 231)
        
        Egresos.RowHeight(0) = 0
        Egresos.Column(0).Width = Val(formatogrilla(8, 1)) * (Egresos.Cell(0, 0).Font.Size + 1.25)
        Egresos.Column(1).Width = Val(formatogrilla(8, 2)) * (Egresos.Cell(0, 1).Font.Size + 1.25)
        Egresos.Column(0).MaxLength = Val(formatogrilla(2, 1))
        Egresos.Column(1).MaxLength = Val(formatogrilla(2, 2))
        Egresos.Column(0).FormatString = formatogrilla(4, 1)
        Egresos.Column(1).FormatString = formatogrilla(4, 2)
        Egresos.Column(0).Locked = formatogrilla(5, 1)
        Egresos.Column(1).Locked = formatogrilla(5, 2)
        If formatogrilla(3, 1) = "N" Then
            Egresos.Column(0).Alignment = cellRightCenter
        Else
            Egresos.Column(0).Alignment = cellLeftCenter
        End If
        If formatogrilla(3, 2) = "N" Then
            Egresos.Column(1).Alignment = cellRightCenter
        Else
            Egresos.Column(1).Alignment = cellLeftCenter
        End If
        
        For i = 1 To row - 1
            Egresos.Cell(i, 0).text = formatogrilla(1, i)
        Next i
        Egresos.Cell(1, 0).Alignment = cellLeftCenter
        Egresos.Cell(2, 0).Alignment = cellLeftCenter
        Egresos.Cell(3, 0).Alignment = cellLeftCenter
        Egresos.Cell(4, 0).Alignment = cellLeftCenter
        Egresos.Cell(5, 0).Alignment = cellLeftCenter
        Egresos.Cell(6, 0).Alignment = cellLeftCenter
        Egresos.Cell(7, 0).Alignment = cellLeftCenter
        Egresos.Cell(8, 0).Alignment = cellLeftCenter
        Egresos.Range(5, 0, 5, 1).BackColor = RGB(200, 225, 250)
        
        Egresos.Enabled = True
    End Sub
'****************************************************************************
'Formato de la Grilla Egresos
'****************************************************************************

'****************************************************************************
'Formato de la Grilla Arqueo de Caja
'****************************************************************************
    Private Sub CargaGrillaArqueo(ByVal row As Integer, ByVal col As Integer)
        Dim i As Integer
        Rem DATOS DE LA COLUMNA
        formatogrilla(1, 1) = ""
        formatogrilla(1, 2) = "CANTIDAD"
        formatogrilla(1, 3) = "DESCUENTO"
        formatogrilla(1, 4) = "NULAS"
        formatogrilla(1, 5) = "TOTAL"
        formatogrilla(1, 6) = "FOL.INICIAL."
        formatogrilla(1, 7) = "FOL.FINAL."
        formatogrilla(1, 8) = ""
        
        Rem LARGO DE LOS DATOS
        formatogrilla(2, 1) = "30"
        formatogrilla(2, 2) = "30"
        formatogrilla(2, 3) = "30"
        formatogrilla(2, 4) = "30"
        formatogrilla(2, 5) = "30"
        formatogrilla(2, 6) = "30"
        formatogrilla(2, 7) = "30"
        formatogrilla(2, 8) = "30"
        
        Rem TIPO DE DATOS C=CEROS,S=STRING,N=NUMERICO,D=DATE,
        formatogrilla(3, 1) = "S"
        formatogrilla(3, 2) = "N"
        formatogrilla(3, 3) = "N"
        formatogrilla(3, 4) = "N"
        formatogrilla(3, 5) = "N"
        formatogrilla(3, 6) = "N"
        formatogrilla(3, 7) = "N"
        formatogrilla(3, 8) = "N"
        
        Rem FORMATO GRILLA
        formatogrilla(4, 1) = ""
        formatogrilla(4, 2) = ""
        formatogrilla(4, 3) = ""
        formatogrilla(4, 4) = ""
        formatogrilla(4, 5) = ""
        formatogrilla(4, 6) = ""
        formatogrilla(4, 7) = ""
        formatogrilla(4, 8) = ""
        
        Rem LOCCKED
        formatogrilla(5, 1) = "FALSE"
        formatogrilla(5, 2) = "FALSE"
        formatogrilla(5, 3) = "FALSE"
        formatogrilla(5, 4) = "FALSE"
        formatogrilla(5, 5) = "FALSE"
        formatogrilla(5, 6) = "FALSE"
        formatogrilla(5, 7) = "FALSE"
        formatogrilla(5, 8) = "FALSE"
        
        Rem VALOR MINIMO
        formatogrilla(6, 1) = ""
        formatogrilla(6, 2) = ""
        formatogrilla(6, 3) = ""
        formatogrilla(6, 4) = ""
        formatogrilla(6, 5) = ""
        formatogrilla(6, 6) = ""
        formatogrilla(6, 7) = ""
        formatogrilla(6, 8) = ""
        
        Rem VALOR MAXIMO
        formatogrilla(7, 1) = ""
        formatogrilla(7, 2) = ""
        formatogrilla(7, 3) = ""
        formatogrilla(7, 4) = ""
        formatogrilla(7, 5) = ""
        formatogrilla(7, 6) = ""
        formatogrilla(7, 7) = ""
        formatogrilla(7, 8) = ""
        
        Rem ANCHO
        formatogrilla(8, 1) = "10"
        formatogrilla(8, 2) = "7"
        formatogrilla(8, 3) = "10"
        formatogrilla(8, 4) = "10"
        formatogrilla(8, 5) = "10"
        formatogrilla(8, 6) = "10"
        formatogrilla(8, 7) = "10"
        formatogrilla(8, 8) = "10"
            
        impresion.Cols = col
        impresion.Rows = row
        impresion.Range(0, 0, impresion.Rows - 1, impresion.Cols - 1).Borders(cellEdgeBottom) = cellNone
        impresion.Range(0, 0, impresion.Rows - 1, impresion.Cols - 1).Borders(cellEdgeTop) = cellNone
        impresion.Range(0, 0, impresion.Rows - 1, impresion.Cols - 1).Borders(cellInsideVertical) = cellNone
        impresion.AllowUserResizing = False
        impresion.DisplayFocusRect = False
        impresion.ExtendLastCol = True
        impresion.BoldFixedCell = False
        impresion.DrawMode = cellOwnerDraw
        impresion.Appearance = Flat
        impresion.ScrollBarStyle = Flat
        impresion.FixedRowColStyle = Flat
        impresion.BackColorFixed = RGB(90, 158, 214)
        impresion.BackColorFixedSel = RGB(110, 180, 230)
        impresion.BackColorBkg = RGB(90, 158, 214)
        impresion.BackColorScrollBar = RGB(231, 235, 247)
        impresion.BackColor1 = RGB(231, 235, 247)
        impresion.BackColor2 = RGB(239, 243, 255)
        impresion.GridColor = RGB(148, 190, 231)
        
        impresion.Column(0).Width = 0
        'impresion.Cell(1, 1).text = "FACTURAS"
        'impresion.Cell(2, 1).text = "BOLETAS"
        'impresion.Cell(3, 1).text = "ZETAS"
        'impresion.Cell(4, 1).text = "TOTALES"
        
        For i = 1 To col - 1
            impresion.Cell(0, i).text = formatogrilla(1, i)
            impresion.Column(i).Width = Val(formatogrilla(8, i)) * (impresion.Cell(0, i).Font.Size + 1.25)
            impresion.Column(i).MaxLength = Val(formatogrilla(2, i))
            impresion.Column(i).FormatString = formatogrilla(4, i)
            impresion.Column(i).Locked = formatogrilla(5, i)
            If formatogrilla(3, i) = "N" Then
                impresion.Column(i).Alignment = cellRightCenter
            Else
                impresion.Column(i).Alignment = cellLeftCenter
            End If
        Next i
        impresion.Range(0, 1, 0, impresion.Cols - 1).Alignment = cellCenterCenter
        'impresion.Range(0, 1, 0, impresion.Cols - 2).Borders(cellEdgeBottom) = cellThin
        'impresion.Range(4, 0, 4, impresion.Cols - 1).BackColor = RGB(200, 225, 250)
        'impresion.Enabled = True
        
        
        
        
        
    End Sub
'****************************************************************************
'Formato de la Grilla Arqueo de Caja
'****************************************************************************

'****************************************************************************
'Formato de la Grilla Listado de Ventas
'****************************************************************************
    Private Sub CargaGrillaVentas(ByVal row As Integer, ByVal col As Integer)
        Dim i As Integer
        Rem DATOS DE LA COLUMNA
        formatogrilla(1, 1) = "DOCUMENTO"
        formatogrilla(1, 2) = "FECHA"
        formatogrilla(1, 3) = "RUT"
        formatogrilla(1, 4) = "CLINTE"
        formatogrilla(1, 5) = "DESCUENTO"
        formatogrilla(1, 6) = "NETO"
        formatogrilla(1, 7) = "IVA"
        formatogrilla(1, 8) = "RETENCION"
        formatogrilla(1, 9) = "TOTAL"
        
        Rem LARGO DE LOS DATOS
        formatogrilla(2, 1) = "30"
        formatogrilla(2, 2) = "30"
        formatogrilla(2, 3) = "30"
        formatogrilla(2, 4) = "50"
        formatogrilla(2, 5) = "30"
        formatogrilla(2, 6) = "30"
        formatogrilla(2, 7) = "30"
        formatogrilla(2, 8) = "30"
        formatogrilla(2, 9) = "30"
        
        Rem TIPO DE DATOS C=CEROS,S=STRING,N=NUMERICO,D=DATE,
        formatogrilla(3, 1) = "S"
        formatogrilla(3, 2) = "N"
        formatogrilla(3, 3) = "N"
        formatogrilla(3, 4) = "S"
        formatogrilla(3, 5) = "N"
        formatogrilla(3, 6) = "N"
        formatogrilla(3, 7) = "N"
        formatogrilla(3, 8) = "N"
        formatogrilla(3, 9) = "N"
        
        Rem FORMATO GRILLA
        formatogrilla(4, 1) = ""
        formatogrilla(4, 2) = ""
        formatogrilla(4, 3) = ""
        formatogrilla(4, 4) = ""
        formatogrilla(4, 5) = ""
        formatogrilla(4, 6) = ""
        formatogrilla(4, 7) = ""
        formatogrilla(4, 8) = ""
        formatogrilla(4, 9) = ""
        
        Rem LOCCKED
        formatogrilla(5, 1) = "FALSE"
        formatogrilla(5, 2) = "FALSE"
        formatogrilla(5, 3) = "FALSE"
        formatogrilla(5, 4) = "FALSE"
        formatogrilla(5, 5) = "FALSE"
        formatogrilla(5, 6) = "FALSE"
        formatogrilla(5, 7) = "FALSE"
        formatogrilla(5, 8) = "FALSE"
        formatogrilla(5, 9) = "FALSE"
        
        Rem VALOR MINIMO
        formatogrilla(6, 1) = ""
        formatogrilla(6, 2) = ""
        formatogrilla(6, 3) = ""
        formatogrilla(6, 4) = ""
        formatogrilla(6, 5) = ""
        formatogrilla(6, 6) = ""
        formatogrilla(6, 7) = ""
        formatogrilla(6, 8) = ""
        formatogrilla(6, 9) = ""
        
        Rem VALOR MAXIMO
        formatogrilla(7, 1) = ""
        formatogrilla(7, 2) = ""
        formatogrilla(7, 3) = ""
        formatogrilla(7, 4) = ""
        formatogrilla(7, 5) = ""
        formatogrilla(7, 6) = ""
        formatogrilla(7, 7) = ""
        formatogrilla(7, 8) = ""
        formatogrilla(7, 9) = ""
        
        Rem ANCHO
        formatogrilla(8, 1) = "9"
        formatogrilla(8, 2) = "7"
        formatogrilla(8, 3) = "7"
        formatogrilla(8, 4) = "15"
        formatogrilla(8, 5) = "8"
        formatogrilla(8, 6) = "8"
        formatogrilla(8, 7) = "8"
        formatogrilla(8, 8) = "8"
        formatogrilla(8, 9) = "8"
            
        impresion.Cols = col
        impresion.Rows = row
        impresion.Range(0, 0, impresion.Rows - 1, impresion.Cols - 1).Borders(cellEdgeBottom) = cellNone
        impresion.Range(0, 0, impresion.Rows - 1, impresion.Cols - 1).Borders(cellEdgeTop) = cellNone
        impresion.Range(0, 0, impresion.Rows - 1, impresion.Cols - 1).Borders(cellInsideVertical) = cellNone
        impresion.AllowUserResizing = False
        impresion.DisplayFocusRect = False
        impresion.ExtendLastCol = True
        impresion.BoldFixedCell = False
        impresion.DrawMode = cellOwnerDraw
        impresion.Appearance = Flat
        impresion.ScrollBarStyle = Flat
        impresion.FixedRowColStyle = Flat
        impresion.BackColorFixed = RGB(90, 158, 214)
        impresion.BackColorFixedSel = RGB(110, 180, 230)
        impresion.BackColorBkg = RGB(90, 158, 214)
        impresion.BackColorScrollBar = RGB(231, 235, 247)
        impresion.BackColor1 = RGB(231, 235, 247)
        impresion.BackColor2 = RGB(239, 243, 255)
        impresion.GridColor = RGB(148, 190, 231)
        
        impresion.Column(0).Width = 0
        'impresion.Cell(1, 1).text = "FACTURAS"
        'impresion.Cell(2, 1).text = "BOLETAS"
        'impresion.Cell(3, 1).text = "ZETAS"
        'impresion.Cell(4, 1).text = "TOTALES"
        
        For i = 1 To col - 1
            impresion.Cell(0, i).text = formatogrilla(1, i)
            impresion.Column(i).Width = Val(formatogrilla(8, i)) * (impresion.Cell(0, i).Font.Size + 1.25)
            impresion.Column(i).MaxLength = Val(formatogrilla(2, i))
            impresion.Column(i).FormatString = formatogrilla(4, i)
            impresion.Column(i).Locked = formatogrilla(5, i)
            If formatogrilla(3, i) = "N" Then
                impresion.Column(i).Alignment = cellRightCenter
            Else
                impresion.Column(i).Alignment = cellLeftCenter
            End If
        Next i
        impresion.Range(0, 1, 0, impresion.Cols - 1).Alignment = cellCenterCenter
        impresion.Range(0, 1, 0, impresion.Cols - 1).Borders(cellEdgeBottom) = cellThin
        'impresion.Range(4, 0, 4, impresion.Cols - 1).BackColor = RGB(200, 225, 250)
        'impresion.Enabled = True
        
    End Sub
'****************************************************************************
'Formato de la Grilla Listado de Ventas
'****************************************************************************

'****************************************************************************
'Formato de la Grilla Listado de Ventas
'****************************************************************************
    Private Sub CargaGrillaPClientes(ByVal row As Integer, ByVal col As Integer)
        Dim i As Integer
        Rem DATOS DE LA COLUMNA
        formatogrilla(1, 1) = "NUMERO"
        formatogrilla(1, 2) = "FECHA PAGO"
        formatogrilla(1, 3) = "RUT"
        formatogrilla(1, 4) = "NOMBRE"
        formatogrilla(1, 5) = ""
        formatogrilla(1, 6) = ""
        formatogrilla(1, 7) = "FORMA PAGO"
        formatogrilla(1, 8) = ""
        formatogrilla(1, 9) = "MONTO PAGO"
        formatogrilla(1, 10) = ""
        
        Rem LARGO DE LOS DATOS
        formatogrilla(2, 1) = "30"
        formatogrilla(2, 2) = "30"
        formatogrilla(2, 3) = "30"
        formatogrilla(2, 4) = "30"
        formatogrilla(2, 5) = "30"
        formatogrilla(2, 6) = "30"
        formatogrilla(2, 7) = "30"
        formatogrilla(2, 8) = "30"
        formatogrilla(2, 9) = "30"
        formatogrilla(2, 10) = "30"
        
        Rem TIPO DE DATOS C=CEROS,S=STRING,N=NUMERICO,D=DATE,
        formatogrilla(3, 1) = "N"
        formatogrilla(3, 2) = "N"
        formatogrilla(3, 3) = "N"
        formatogrilla(3, 4) = "S"
        formatogrilla(3, 5) = "S"
        formatogrilla(3, 6) = "S"
        formatogrilla(3, 7) = "S"
        formatogrilla(3, 8) = "N"
        formatogrilla(3, 9) = "N"
        formatogrilla(3, 10) = "S"
        
        Rem FORMATO GRILLA
        formatogrilla(4, 1) = ""
        formatogrilla(4, 2) = ""
        formatogrilla(4, 3) = ""
        formatogrilla(4, 4) = ""
        formatogrilla(4, 5) = ""
        formatogrilla(4, 6) = ""
        formatogrilla(4, 7) = ""
        formatogrilla(4, 8) = ""
        formatogrilla(4, 9) = ""
        formatogrilla(4, 10) = ""
        
        Rem LOCCKED
        formatogrilla(5, 1) = "FALSE"
        formatogrilla(5, 2) = "FALSE"
        formatogrilla(5, 3) = "FALSE"
        formatogrilla(5, 4) = "FALSE"
        formatogrilla(5, 5) = "FALSE"
        formatogrilla(5, 6) = "FALSE"
        formatogrilla(5, 7) = "FALSE"
        formatogrilla(5, 8) = "FALSE"
        formatogrilla(5, 9) = "FALSE"
        formatogrilla(5, 10) = "FALSE"
        
        Rem VALOR MINIMO
        formatogrilla(6, 1) = ""
        formatogrilla(6, 2) = ""
        formatogrilla(6, 3) = ""
        formatogrilla(6, 4) = ""
        formatogrilla(6, 5) = ""
        formatogrilla(6, 6) = ""
        formatogrilla(6, 7) = ""
        formatogrilla(6, 8) = ""
        formatogrilla(6, 9) = ""
        formatogrilla(6, 10) = ""
        
        Rem VALOR MAXIMO
        formatogrilla(7, 1) = ""
        formatogrilla(7, 2) = ""
        formatogrilla(7, 3) = ""
        formatogrilla(7, 4) = ""
        formatogrilla(7, 5) = ""
        formatogrilla(7, 6) = ""
        formatogrilla(7, 7) = ""
        formatogrilla(7, 8) = ""
        formatogrilla(7, 9) = ""
        formatogrilla(7, 10) = ""
        
        Rem ANCHO
        formatogrilla(8, 1) = "8"
        formatogrilla(8, 2) = "8"
        formatogrilla(8, 3) = "7"
        formatogrilla(8, 4) = "7"
        formatogrilla(8, 5) = "7"
        formatogrilla(8, 6) = "7"
        formatogrilla(8, 7) = "8"
        formatogrilla(8, 8) = "8"
        formatogrilla(8, 9) = "8"
        formatogrilla(8, 10) = "8"
            
        impresion.Cols = col
        impresion.Rows = row
        impresion.Range(0, 0, impresion.Rows - 1, impresion.Cols - 1).Borders(cellEdgeBottom) = cellNone
        impresion.Range(0, 0, impresion.Rows - 1, impresion.Cols - 1).Borders(cellEdgeTop) = cellNone
        impresion.Range(0, 0, impresion.Rows - 1, impresion.Cols - 1).Borders(cellInsideVertical) = cellNone
        impresion.AllowUserResizing = False
        impresion.DisplayFocusRect = False
        impresion.ExtendLastCol = True
        impresion.BoldFixedCell = False
        impresion.DrawMode = cellOwnerDraw
        impresion.Appearance = Flat
        impresion.ScrollBarStyle = Flat
        impresion.FixedRowColStyle = Flat
        impresion.BackColorFixed = RGB(90, 158, 214)
        impresion.BackColorFixedSel = RGB(110, 180, 230)
        impresion.BackColorBkg = RGB(90, 158, 214)
        impresion.BackColorScrollBar = RGB(231, 235, 247)
        impresion.BackColor1 = RGB(231, 235, 247)
        impresion.BackColor2 = RGB(239, 243, 255)
        impresion.GridColor = RGB(148, 190, 231)
        
        impresion.Column(0).Width = 0
        'impresion.Cell(1, 1).text = "FACTURAS"
        'impresion.Cell(2, 1).text = "BOLETAS"
        'impresion.Cell(3, 1).text = "ZETAS"
        'impresion.Cell(4, 1).text = "TOTALES"
        
        For i = 1 To col - 1
            impresion.Cell(0, i).text = formatogrilla(1, i)
            impresion.Column(i).Width = Val(formatogrilla(8, i)) * (impresion.Cell(0, i).Font.Size + 1.25)
            impresion.Column(i).MaxLength = Val(formatogrilla(2, i))
            impresion.Column(i).FormatString = formatogrilla(4, i)
            impresion.Column(i).Locked = formatogrilla(5, i)
            If formatogrilla(3, i) = "N" Then
                impresion.Column(i).Alignment = cellRightCenter
            Else
                impresion.Column(i).Alignment = cellLeftCenter
            End If
        Next i
        impresion.Range(0, 1, 0, impresion.Cols - 1).Alignment = cellCenterCenter
        impresion.Range(0, 1, 0, impresion.Cols - 1).Borders(cellEdgeBottom) = cellThin
        'impresion.Range(4, 0, 4, impresion.Cols - 1).BackColor = RGB(200, 225, 250)
        'impresion.Enabled = True
                
    End Sub
'****************************************************************************
'Formato de la Grilla Listado de Ventas
'****************************************************************************

'****************************************************************************
'Formato de la Grilla Impresion
'****************************************************************************
    Private Sub CargaGrillaImpresion(ByVal row As Integer, ByVal col As Integer)
        Dim i As Integer
        impresion.Cols = col
        impresion.Rows = row
        impresion.AllowUserResizing = False
        impresion.DisplayFocusRect = False
        impresion.ExtendLastCol = False
        impresion.BoldFixedCell = False
        impresion.DrawMode = cellOwnerDraw
        impresion.Appearance = Flat
        impresion.ScrollBarStyle = Flat
        impresion.FixedRowColStyle = Flat
        impresion.BackColorFixed = RGB(90, 158, 214)
        impresion.BackColorFixedSel = RGB(110, 180, 230)
        impresion.BackColorBkg = RGB(90, 158, 214)
        impresion.BackColorScrollBar = RGB(231, 235, 247)
        impresion.BackColor1 = RGB(231, 235, 247)
        impresion.BackColor2 = RGB(239, 243, 255)
        impresion.GridColor = RGB(148, 190, 231)
        
        impresion.Column(0).Width = 0
        
        For i = 1 To col - 1
            impresion.Column(i).Width = 2.25 * (impresion.Cell(0, i).Font.Size)
        Next i
    End Sub
'****************************************************************************
'Formato de la Grilla Impresion
'****************************************************************************

'=============================================================================
'PASA LOS DATOS DE LA ESTRUCTURA A LOS CONTROLES
'=============================================================================
    Private Sub structtoctrl()
        'FACTURAS
        Documentos.Cell(1, 1).text = a.factura.cantidad
        Documentos.Cell(1, 2).text = a.factura.Descuento
        Documentos.Cell(1, 3).text = a.factura.nulas
        Documentos.Cell(1, 4).text = a.factura.total
        Documentos.Cell(1, 5).text = a.factura.folini
        Documentos.Cell(1, 6).text = a.factura.folfin
        'BOLETAS
        Documentos.Cell(2, 1).text = a.boleta.cantidad
        Documentos.Cell(2, 2).text = a.boleta.Descuento
        Documentos.Cell(2, 3).text = a.boleta.nulas
        Documentos.Cell(2, 4).text = a.boleta.total
        Documentos.Cell(2, 5).text = a.boleta.folini
        Documentos.Cell(2, 6).text = a.boleta.folfin
        'ZETAS
        Documentos.Cell(3, 1).text = a.zeta.cantidad
        Documentos.Cell(3, 2).text = a.zeta.Descuento
        Documentos.Cell(3, 3).text = a.zeta.nulas
        Documentos.Cell(3, 4).text = a.zeta.total
        Documentos.Cell(3, 5).text = a.zeta.folini
        Documentos.Cell(3, 6).text = a.zeta.folfin
        'FACTURAS EXENTAS
        Documentos.Cell(4, 1).text = a.exenta.cantidad
        Documentos.Cell(4, 2).text = a.exenta.Descuento
        Documentos.Cell(4, 3).text = a.exenta.nulas
        Documentos.Cell(4, 4).text = a.exenta.total
        Documentos.Cell(4, 5).text = a.exenta.folini
        Documentos.Cell(4, 6).text = a.exenta.folfin
        'NOTAS DE CREDITO
        Documentos.Cell(5, 1).text = a.ncredito.cantidad
        Documentos.Cell(5, 2).text = a.ncredito.Descuento
        Documentos.Cell(5, 3).text = a.ncredito.nulas
        Documentos.Cell(5, 4).text = a.ncredito.total
        Documentos.Cell(5, 5).text = a.ncredito.folini
        Documentos.Cell(5, 6).text = a.ncredito.folfin
        
        'INGRESOS
'        Ingresos.Cell(1, 1).text = a.Ingreso.efectivo
'        Ingresos.Cell(2, 1).text = a.Ingreso.Cheques
'        Ingresos.Cell(3, 1).text = a.Ingreso.tarjetas
'        Ingresos.Cell(4, 1).text = a.Ingreso.Depositos
'        Ingresos.Cell(6, 1).text = a.Ingreso.pagoClientes
'
        'EGRESOS
        Egresos.Cell(1, 1).text = a.egreso.egresosCaja
        Egresos.Cell(2, 1).text = a.egreso.chequesFecha
        'Egresos.Cell(3, 1).text = a.Ingreso.tarjetas
        Egresos.Cell(4, 1).text = a.egreso.Depositos
        
        Call sumaGrilla
    End Sub
'=============================================================================
'PASA LOS DATOS DE LA ESTRUCTURA A LOS CONTROLES
'=============================================================================

    Private Sub sumaGrilla()
        Dim i As Long
        Dim efectivo As Double
        Dim total As Double
        
        Documentos.Cell(6, 1).text = "$ 0"
        Documentos.Cell(6, 2).text = "$ 0"
        Documentos.Cell(6, 3).text = "$ 0"
        Documentos.Cell(6, 4).text = "$ 0"
        For i = 1 To 5
            Documentos.Cell(6, 1).text = Format(CDbl(Documentos.Cell(6, 1).text) + CDbl(Documentos.Cell(i, 1).text), "$ ###,###,##0")
            Documentos.Cell(6, 3).text = Format(CDbl(Documentos.Cell(6, 3).text) + CDbl(Documentos.Cell(i, 3).text), "$ ###,###,##0")
            If i < 5 Then
                Documentos.Cell(6, 2).text = Format(CDbl(Documentos.Cell(6, 2).text) + CDbl(Documentos.Cell(i, 2).text), "$ ###,###,##0")
                Documentos.Cell(6, 4).text = Format(CDbl(Documentos.Cell(6, 4).text) + CDbl(Documentos.Cell(i, 4).text), "$ ###,###,##0")
            End If
        Next i
        total = CDbl(Documentos.Cell(6, 4).text)
        Ingresos.Cell(11, 1).text = Format(total, "$ ###,###,##0")
        If Ingresos.Cell(12, 1).text = "" Then Ingresos.Cell(12, 1).text = "0"
        If Ingresos.Cell(13, 1).text = "" Then Ingresos.Cell(13, 1).text = "0"
        
        total = total + CDbl(Ingresos.Cell(12, 1).text) + CDbl(Ingresos.Cell(13, 1).text)
        Ingresos.Cell(14, 1).text = Format(total, "$ ###,###,##0")
        efectivo = total
        For i = 2 To 10
        If Ingresos.Cell(i, 1).text = "" Then Ingresos.Cell(i, 1).text = "0"
        efectivo = efectivo - CDbl(Ingresos.Cell(i, 1).text)
        Next i
        
        
        Ingresos.Cell(1, 1).text = Format(efectivo, "$ ###,###,##0")
        
        'ZURITA
        Ingresos.Cell(15, 1).text = TotalEmpresaRelacionada(empresaActiva, fecha1, fecha2)
        Ingresos.Cell(16, 1).text = CDbl(Ingresos.Cell(14, 1).text) - CDbl(Ingresos.Cell(15, 1).text)
        
        
        
        
        
        Egresos.Cell(5, 1).text = "$ 0"
'        Ingresos.Cell(1, 1).text = "$ 0"
        For i = 1 To 4
'            If i > 1 Then
'                Ingresos.Cell(1, 1).text = Format(CDbl(Ingresos.Cell(1, 1).text) + CDbl(Ingresos.Cell(i, 1).text), "$ ###,###,##0")
'            End If
           If Egresos.Cell(i, 1).text = "" Then Egresos.Cell(i, 1).text = "0"
           Egresos.Cell(5, 1).text = Format(CDbl(Egresos.Cell(5, 1).text) + CDbl(Egresos.Cell(i, 1).text), "$ ###,###,##0")
        Next i
'        Ingresos.Cell(1, 1).text = Format(CDbl(Ingresos.Cell(5, 1).text) - CDbl(Ingresos.Cell(1, 1).text), "$ ###,###,##0")
        Egresos.Cell(6, 1).text = Format(CDbl(Ingresos.Cell(14, 1).text) - CDbl(Egresos.Cell(5, 1).text) - CDbl(Ingresos.Cell(2, 1).text) + CDbl(Egresos.Cell(2, 1).text), "$ ###,###,##0")
        Egresos.Cell(7, 1).text = Format(Ingresos.Cell(2, 1).text, "$ ###,###,##0")
        Egresos.Cell(8, 1).text = "$ 0"

    
    End Sub
    



    Private Sub frmImprimir_BarMouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
        frmImprimir.ColorBarraAbajo = &HFFC0C0
        frmImprimir.ColorBarraArriba = &H800000
        frmImprimir.CaptionEstilo3D = Raised
    End Sub

    Private Sub frmImprimir_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
        frmImprimir.ColorBarraAbajo = &H800000
        frmImprimir.ColorBarraArriba = &HFFC0C0
        frmImprimir.CaptionEstilo3D = Inserted
        Call imprimirUnaGrilla
        'Call imprimir
    End Sub
    
    Private Sub imprimirUnaGrilla()
        Dim i As Long
        
        salto = False
        Call CargaGrillaImpresion(1, 40)
        ''''''''''''''''
        'Impresion.Visible = True
        'Impresion.ZOrder
        ''''''''''''''''
        impresion.AutoRedraw = False
        
        impresion.DefaultFont.Name = "Arial"
        
        Call cargaCabeza("INFORMES DESDE EL " & Format(fecha1, "dd-mm-yyyy") & " HASTA EL " & Format(fecha2, "dd-mm-yyyy"), dato1.text, impresion)
        
        Call imprimirUnaGrillaArqueo
        If salto = True Then
            Call impresion.HPageBreaks.Add(impresion.Rows - 1)
        End If
       If Option1.Value = True Then
        Call imprimirUnaGrillaVentas
        If salto = True Then
            Call impresion.HPageBreaks.Add(impresion.Rows - 1)
        End If
        End If
        Call imprimirUnaGrillaEgresos
        If salto = True Then
            Call impresion.HPageBreaks.Add(impresion.Rows - 1)
        End If
        
        Call imprimirUnaGrillaPagos
        If salto = True Then
            Call impresion.HPageBreaks.Add(impresion.Rows - 1)
        End If
        
       ' Call imprimirUnaGrillaMovimientos
        
        impresion.AutoRedraw = True
        impresion.Refresh
        
        impresion.PageSetup.BlackAndWhite = True
        impresion.PageSetup.LeftMargin = 0.75
        impresion.PageSetup.RightMargin = 0.5
        impresion.PageSetup.TopMargin = 1.5
        impresion.PageSetup.HeaderMargin = 1.5
        impresion.PageSetup.BottomMargin = 1.5
        impresion.PageSetup.FooterMargin = 1.5
        
        For i = 1 To impresion.PageSetup.PaperSizes.Count
            If UCase(impresion.PageSetup.PaperSizes.Item(i).PaperName) = "CARTA" Then
                impresion.PageSetup.PaperSize = impresion.PageSetup.PaperSizes.Item(i).Kind
                Exit For
            End If
        Next i
        
        Call verificaImpresora(5, impresion)
        
    End Sub
    
    Private Sub imprimirUnaGrillaArqueo()
        Dim tabla As String
        Dim cadena As String
        Dim fila As Long
        Dim i As Long
        Dim j As Long
        Dim K As Long
        Dim fecha As Double
        Dim dia As Double
        Dim rut As String
        Dim tabulador As String
        Dim cajeras As String
        Dim caja As String
        'DOCUMENTOS RESUMEN
        'TITULO
        If nombrecajera.Caption <> "" Then
        impresion.AddItem "RESUMEN DE ARQUEO DE CAJA  CAJERA :" + tcajera.text + " " + nombrecajera.Caption, True
        End If
        If nombrecaja.Caption <> "" Then
        impresion.AddItem "RESUMEN DE ARQUEO DE CAJA  NUMERO:" + tcaja.text + " " + nombrecaja.Caption, True
        End If
        If nombrecaja.Caption = "" And nombrecajera.Caption = "" Then
        impresion.AddItem "RESUMEN DE ARQUEO DE CAJA  TODAS LAS CAJERAS Y CAJAS DEL LOCAL"
        
        End If
        If nombrecaja.Caption <> "" And nombrecajera.Caption <> "" Then
        impresion.AddItem "RESUMEN DE ARQUEO DE CAJA  CAJERA :" + tcajera.text + " " + nombrecajera.Caption + "  NUMERO CAJA:" + tcaja.text + " " + nombrecaja.Caption, True
        
        End If
        
        
        impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, impresion.Cols - 1).Merge
        impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, impresion.Cols - 1).Alignment = cellCenterCenter
        impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, impresion.Cols - 1).Borders(cellEdgeBottom) = cellThin
        impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, impresion.Cols - 1).FontBold = True
        
        tabulador = ""
        For i = 1 To 5
            tabulador = tabulador & vbTab
        Next i
        For i = 0 To Documentos.Rows - 1
            cadena = ""
            For j = 0 To Documentos.Cols - 1
                If j = 2 Or j = 4 Or j = 5 Or j = 6 Then
                    If j = 5 Or j = 6 Then
                        cadena = cadena & Format(Documentos.Cell(i, j).text, "0000000000") & tabulador & vbTab
                    Else
                        cadena = cadena & Documentos.Cell(i, j).text & tabulador & vbTab
                    End If
                Else
                    If j = 1 Or j = 3 Then
                        cadena = cadena & Format(Documentos.Cell(i, j).text, "########0") & tabulador
                    Else
                        cadena = cadena & Documentos.Cell(i, j).text & tabulador
                    End If
                End If
            Next j
            impresion.AddItem cadena, True
            'UNION DE CELDAS
            impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, 5).Merge
            impresion.Range(impresion.Rows - 1, 6, impresion.Rows - 1, 10).Merge
            impresion.Range(impresion.Rows - 1, 11, impresion.Rows - 1, 16).Merge
            impresion.Range(impresion.Rows - 1, 17, impresion.Rows - 1, 21).Merge
            impresion.Range(impresion.Rows - 1, 22, impresion.Rows - 1, 27).Merge
            impresion.Range(impresion.Rows - 1, 28, impresion.Rows - 1, 33).Merge
            impresion.Range(impresion.Rows - 1, 34, impresion.Rows - 1, 39).Merge
            'ALINEACION DE CELDAS
            If i = 0 Then
                impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, impresion.Cols - 1).Alignment = cellCenterCenter
            Else
                impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, 5).Alignment = cellLeftCenter
                impresion.Range(impresion.Rows - 1, 6, impresion.Rows - 1, 10).Alignment = cellRightCenter
                impresion.Range(impresion.Rows - 1, 11, impresion.Rows - 1, 16).Alignment = cellRightCenter
                impresion.Range(impresion.Rows - 1, 17, impresion.Rows - 1, 21).Alignment = cellRightCenter
                impresion.Range(impresion.Rows - 1, 22, impresion.Rows - 1, 27).Alignment = cellRightCenter
                impresion.Range(impresion.Rows - 1, 28, impresion.Rows - 1, 33).Alignment = cellRightCenter
                impresion.Range(impresion.Rows - 1, 34, impresion.Rows - 1, impresion.Cols - 1).Alignment = cellRightCenter
            End If
        Next i
        'BORDES
        impresion.Range(impresion.Rows - 5, 1, impresion.Rows - 5, impresion.Cols - 1).Borders(cellEdgeBottom) = cellThin
        impresion.Range(impresion.Rows - 2, 1, impresion.Rows - 2, impresion.Cols - 13).Borders(cellEdgeBottom) = cellThin
        impresion.Range(impresion.Rows - 5, 1, impresion.Rows - 1, 1).Borders(cellEdgeRight) = cellThin
        
        impresion.AddItem "", True
        'Impresion.AddItem "", True
        
        'INGRESOS Y EGRESOS
        tabulador = ""
        For i = 1 To 20
            tabulador = tabulador & vbTab
        Next i
        'TITULO
        impresion.AddItem "RESUMEN INGRESOS" & tabulador & "", True
        impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, 19).Merge
        impresion.Range(impresion.Rows - 1, 21, impresion.Rows - 1, impresion.Cols - 2).Merge
        impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, impresion.Cols - 1).Alignment = cellCenterCenter
        impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, 19).Borders(cellEdgeBottom) = cellThin
        impresion.Range(impresion.Rows - 1, 21, impresion.Rows - 1, impresion.Cols - 1).Borders(cellEdgeBottom) = cellThin
        impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, impresion.Cols - 1).FontBold = True
        'DATOS
        tabulador = ""
        For i = 1 To 9
            tabulador = tabulador & vbTab
        Next i
        For i = 1 To Egresos.Rows - 1
            cadena = ""
            If i <= Ingresos.Rows - 1 Then
                For j = 0 To Ingresos.Cols - 1
                    If j = 0 Then
                        cadena = cadena & Ingresos.Cell(i, j).text & tabulador & vbTab
                    Else
                        cadena = cadena & Ingresos.Cell(i, j).text & tabulador
                    End If
                Next j
                cadena = cadena & tabulador
            Else
                cadena = cadena & tabulador & tabulador & vbTab & vbTab
            End If
            
            For K = 0 To Ingresos.Cols - 1
                If K = 0 Then
                    ' cadena = cadena & Egresos.Cell(i, k).text & tabulador & vbTab
                    cadena = cadena & ""
                Else
                    cadena = cadena & ""
                End If
            Next K
            'UNION DE LAS CELDAS
            impresion.AddItem cadena, True
            impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, 10).Merge
            impresion.Range(impresion.Rows - 1, 11, impresion.Rows - 1, 19).Merge
           ' Impresion.Range(Impresion.Rows - 1, 21, Impresion.Rows - 1, 30).Merge
            'Impresion.Range(Impresion.Rows - 1, 31, Impresion.Rows - 1, 39).Merge
            'ALINEACION
            impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, 10).Alignment = cellLeftCenter
            impresion.Range(impresion.Rows - 1, 11, impresion.Rows - 1, 19).Alignment = cellRightCenter
            'Impresion.Range(Impresion.Rows - 1, 21, Impresion.Rows - 1, 30).Alignment = cellLeftCenter
            'Impresion.Range(Impresion.Rows - 1, 31, Impresion.Rows - 1, 39).Alignment = cellRightCenter
        Next i
        'BORDES
        impresion.Range(impresion.Rows - 5, 1, impresion.Rows - 5, 19).Borders(cellEdgeBottom) = cellThin
        'Impresion.Range(Impresion.Rows - 5, 21, Impresion.Rows - 5, Impresion.Cols - 1).Borders(cellEdgeBottom) = cellThin
        impresion.Range(impresion.Rows - 4, 1, impresion.Rows - 4, 19).Borders(cellEdgeBottom) = cellThin
        'Impresion.Range(Impresion.Rows - 4, 21, Impresion.Rows - 4, Impresion.Cols - 1).Borders(cellEdgeBottom) = cellThin
        'Impresion.Range(Impresion.Rows - 2, 21, Impresion.Rows - 2, Impresion.Cols - 1).Borders(cellEdgeBottom) = cellThin
        
        impresion.AddItem "", True
        salto = True
        'CHEQUES POR VENTAS
        tabulador = ""
        For i = 1 To 3
            tabulador = tabulador & vbTab
        Next i
        cajeras = tcajera.text
        caja = tcaja.text
        
'        tabla = "SELECT CONCAT(CONCAT(dp.tipo,' ', dp.numero), '" & tabulador & vbTab & vbTab & "', CONCAT(dp.tipopago, ' CH ')) AS item1, CONCAT('" & tabulador & tabulador & tabulador & tabulador & vbTab & vbTab & "', dp.banco, '" & tabulador & "', dp.plaza, '" & tabulador & "',dp.numerodocumento, '" & tabulador & vbTab & "') AS item2, CONCAT('$ ', FORMAT(dp.monto,0), '" & tabulador & tabulador & "', DATE_FORMAT(dp.vencimiento,'%d-%m-%Y')) AS item3, IFNULL(dp.vencimiento,'') AS vencimiento, dp.monto, dp.rut "
'        tabla = tabla & "FROM sv_documento_pagos_" + PAuditoriaVentas.dato1.text + " AS dp, sv_documento_cabeza_" + PAuditoriaVentas.dato1.text + " as dc "
'        tabla = tabla & "WHERE dc.foliosii=dp.foliofiscal and dp.local = '" & localAuditoria & "' AND dp.fecha BETWEEN '" & fecha1 & "' AND '" & fecha2 & "' AND dp.tipopago = '2' and dc.cajera like '%" + cajeras + "%' and dc.caja like '%" + caja + "%' ORDER BY dp.tipo, dp.numero ASC"
'        Call ConectarControlData(data, servidor, baseVentas & PAuditoriaVentas.dato1.text, usuario, password, tabla)
'        If data.Recordset.RecordCount > 0 Then
'            'TITULO
'            impresion.AddItem "CHEQUES RECIBIDOR POR VENTAS", True
'            impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, impresion.Cols - 1).Merge
'            impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, impresion.Cols - 1).Alignment = cellCenterCenter
'            impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, impresion.Cols - 1).Borders(cellEdgeBottom) = cellThin
'            impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, impresion.Cols - 1).FontBold = True
'
'            fecha = 0
'            dia = 0
'            'CABEZA
'            cadena = "DOCUMENTO" & tabulador & vbTab & vbTab
'            cadena = cadena & "DETALLE CUENTA" & tabulador & tabulador & tabulador & tabulador & vbTab & vbTab
'            cadena = cadena & "BANCO" & tabulador
'            cadena = cadena & "PLAZA" & tabulador
'            cadena = cadena & "NUMERO" & tabulador & vbTab
'            cadena = cadena & "MONTO" & tabulador & tabulador
'            cadena = cadena & "VENCIMIENTO"
'            impresion.AddItem cadena, True
'            impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, impresion.Cols - 1).Borders(cellEdgeBottom) = cellThin
'            'UNION DE CELDAS
'            impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, 5).Merge
'            impresion.Range(impresion.Rows - 1, 6, impresion.Rows - 1, 19).Merge
'            impresion.Range(impresion.Rows - 1, 20, impresion.Rows - 1, 22).Merge
'            impresion.Range(impresion.Rows - 1, 23, impresion.Rows - 1, 25).Merge
'            impresion.Range(impresion.Rows - 1, 26, impresion.Rows - 1, 29).Merge
'            impresion.Range(impresion.Rows - 1, 30, impresion.Rows - 1, 35).Merge
'            impresion.Range(impresion.Rows - 1, 36, impresion.Rows - 1, impresion.Cols - 1).Merge
'            impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, impresion.Cols - 1).Alignment = cellCenterCenter
'
'            data.Recordset.MoveFirst
'            While Not data.Recordset.EOF
'                rut = data.Recordset.Fields("rut")
'                If data.Recordset.Fields("vencimiento") > fechasistema Then
'                    fecha = fecha + CDbl(data.Recordset.Fields("monto"))
'                Else
'                    dia = dia + CDbl(data.Recordset.Fields("monto"))
'                End If
'                impresion.AddItem data.Recordset.Fields("item1") & "   " & leerNombreCliente(rut) & data.Recordset.Fields("item2") & Replace(data.Recordset.Fields("item3"), ",", "."), True
'                'UNION DE CELDAS
'                impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, 5).Merge
'                impresion.Range(impresion.Rows - 1, 6, impresion.Rows - 1, 19).Merge
'                impresion.Range(impresion.Rows - 1, 20, impresion.Rows - 1, 22).Merge
'                impresion.Range(impresion.Rows - 1, 23, impresion.Rows - 1, 25).Merge
'                impresion.Range(impresion.Rows - 1, 26, impresion.Rows - 1, 29).Merge
'                impresion.Range(impresion.Rows - 1, 30, impresion.Rows - 1, 35).Merge
'                impresion.Range(impresion.Rows - 1, 36, impresion.Rows - 1, impresion.Cols - 1).Merge
'                'ALINEACION
'                impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, 5).Alignment = cellCenterCenter
'                impresion.Range(impresion.Rows - 1, 6, impresion.Rows - 1, 19).Alignment = cellLeftCenter
'                impresion.Range(impresion.Rows - 1, 20, impresion.Rows - 1, 22).Alignment = cellRightCenter
'                impresion.Range(impresion.Rows - 1, 23, impresion.Rows - 1, 25).Alignment = cellRightCenter
'                impresion.Range(impresion.Rows - 1, 26, impresion.Rows - 1, 29).Alignment = cellRightCenter
'                impresion.Range(impresion.Rows - 1, 30, impresion.Rows - 1, 35).Alignment = cellRightCenter
'                impresion.Range(impresion.Rows - 1, 36, impresion.Rows - 1, impresion.Cols - 1).Alignment = cellCenterCenter
'                data.Recordset.MoveNext
'            Wend
'        End If
'
'        impresion.AddItem "", True
'
'        'CHEQUES POR PAGOS DE CLIENTES
'        tabulador = ""
'        For i = 1 To 3
'            tabulador = tabulador & vbTab
'        Next i
'
'        tabla = "SELECT CONCAT(CONCAT('PA', ' ', pc.numero), '" & tabulador & vbTab & vbTab & "', CONCAT(pc.tipopago, ' CH ')) AS item1, CONCAT('" & tabulador & tabulador & tabulador & tabulador & vbTab & vbTab & "', c.banco, '" & tabulador & "', c.plaza, '" & tabulador & "',c.numerocheque, '" & tabulador & vbTab & "') AS item2, CONCAT('$ ', FORMAT(c.monto,0), '" & tabulador & tabulador & "', DATE_FORMAT(c.fechavencimiento,'%d-%m-%Y')) AS item3, IFNULL(c.fechavencimiento,'') AS fechavencimiento, c.monto, pc.rut , pc.numero "
'        tabla = tabla & "FROM sv_pagos_cabeza_" & empresaActiva & " AS pc INNER JOIN sv_carteracheques AS c ON pc.local = c.local AND pc.numero = c.numero AND c.tipodocumento = 'PA' "
'        tabla = tabla & "WHERE pc.local = '" & localAuditoria & "' AND pc.fecha BETWEEN '" & fecha1 & "' AND '" & fecha2 & "' AND pc.tipopago = '2' "
'        tabla = tabla & "ORDER BY numero ASC"
'
'
'        'tabla = "SELECT CONCAT(CONCAT(pd.tipo,' ', pd.documento), '" & tabulador & vbTab & vbTab & "', CONCAT(pd.formapago, ' CH ')) AS item1, CONCAT('" & tabulador & tabulador & tabulador & tabulador & vbTab & vbTab & "', c.banco, '" & tabulador & "', c.plaza, '" & tabulador & "',c.numerocheque, '" & tabulador & vbTab & "') AS item2, CONCAT('$ ', FORMAT(c.monto,0), '" & tabulador & tabulador & "', DATE_FORMAT(c.fechavencimiento,'%d-%m-%Y')) AS item3, IFNULL(c.fechavencimiento,'') AS fechavencimiento, c.monto, pd.rut "
'        'tabla = tabla & "FROM sv_pagos_detalle AS pd INNER JOIN sv_carteracheques AS c ON pd.local = c.local AND pd.numero = c.numero /*INNER JOIN " & baseVentas & ".sv_maestroclientes AS mc ON pd.rut = mc.rut*/ LEFT JOIN " & baseVentas & ".sv_maestrobancos AS mb ON c.banco = mb.codigobanco INNER JOIN sv_pagos_cabeza AS pc ON pd.local = pc.local AND pd.numero = pc.numero "
'        'tabla = tabla & "WHERE pd.local = '" & localauditoria & "' AND pd.fecha BETWEEN '" & fecha1 & "' AND '" & fecha2 & "' AND pc.tipopago = '2' ORDER BY pd.numero ASC"
'        ' Call ConectarControlData(data, servidor, baseVentas & rubroAuditoria, usuario, password, tabla)
'
'        If data.Recordset.RecordCount < 0 Then
'            'TITULO
'            impresion.AddItem "CHEQUES RECIBIDOR POR PAGOS", True
'            impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, impresion.Cols - 1).Merge
'            impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, impresion.Cols - 1).Alignment = cellCenterCenter
'            impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, impresion.Cols - 1).Borders(cellEdgeBottom) = cellThin
'            impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, impresion.Cols - 1).FontBold = True
'
'            'fecha = 0
'            'dia = 0
'            'CABEZA
'            cadena = "DOCUMENTO" & tabulador & vbTab & vbTab
'            cadena = cadena & "DETALLE CUENTA" & tabulador & tabulador & tabulador & tabulador & vbTab & vbTab
'            cadena = cadena & "BANCO" & tabulador
'            cadena = cadena & "PLAZA" & tabulador
'            cadena = cadena & "NUMERO" & tabulador & vbTab
'            cadena = cadena & "MONTO" & tabulador & tabulador
'            cadena = cadena & "VENCIMIENTO"
'            impresion.AddItem cadena, True
'            impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, impresion.Cols - 1).Borders(cellEdgeBottom) = cellThin
'            'UNION DE CELDAS
'            impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, 5).Merge
'            impresion.Range(impresion.Rows - 1, 6, impresion.Rows - 1, 19).Merge
'            impresion.Range(impresion.Rows - 1, 20, impresion.Rows - 1, 22).Merge
'            impresion.Range(impresion.Rows - 1, 23, impresion.Rows - 1, 25).Merge
'            impresion.Range(impresion.Rows - 1, 26, impresion.Rows - 1, 29).Merge
'            impresion.Range(impresion.Rows - 1, 30, impresion.Rows - 1, 35).Merge
'            impresion.Range(impresion.Rows - 1, 36, impresion.Rows - 1, impresion.Cols - 1).Merge
'            impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, impresion.Cols - 1).Alignment = cellCenterCenter
'
'            data.Recordset.MoveFirst
'            While Not data.Recordset.EOF
'                rut = data.Recordset.Fields("rut")
'                If data.Recordset.Fields("fechavencimiento") > fechasistema Then
'                    fecha = fecha + CDbl(data.Recordset.Fields("monto"))
'                Else
'                    dia = dia + CDbl(data.Recordset.Fields("monto"))
'                End If
'                impresion.AddItem data.Recordset.Fields("item1") & "   " & leerNombreCliente(rut) & data.Recordset.Fields("item2") & Replace(data.Recordset.Fields("item3"), ",", "."), True
'                'UNION DE CELDAS
'                impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, 5).Merge
'                impresion.Range(impresion.Rows - 1, 6, impresion.Rows - 1, 19).Merge
'                impresion.Range(impresion.Rows - 1, 20, impresion.Rows - 1, 22).Merge
'                impresion.Range(impresion.Rows - 1, 23, impresion.Rows - 1, 25).Merge
'                impresion.Range(impresion.Rows - 1, 26, impresion.Rows - 1, 29).Merge
'                impresion.Range(impresion.Rows - 1, 30, impresion.Rows - 1, 35).Merge
'                impresion.Range(impresion.Rows - 1, 36, impresion.Rows - 1, impresion.Cols - 1).Merge
'                'ALINEACION
'                impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, 5).Alignment = cellCenterCenter
'                impresion.Range(impresion.Rows - 1, 6, impresion.Rows - 1, 19).Alignment = cellLeftCenter
'                impresion.Range(impresion.Rows - 1, 20, impresion.Rows - 1, 22).Alignment = cellRightCenter
'                impresion.Range(impresion.Rows - 1, 23, impresion.Rows - 1, 25).Alignment = cellRightCenter
'                impresion.Range(impresion.Rows - 1, 26, impresion.Rows - 1, 29).Alignment = cellRightCenter
'                impresion.Range(impresion.Rows - 1, 30, impresion.Rows - 1, 35).Alignment = cellRightCenter
'                impresion.Range(impresion.Rows - 1, 36, impresion.Rows - 1, impresion.Cols - 1).Alignment = cellCenterCenter
'                data.Recordset.MoveNext
'            Wend
'            impresion.AddItem "", True
'
'            tabulador = ""
'            For i = 1 To 5
'                tabulador = tabulador & vbTab
'            Next i
'            impresion.AddItem tabulador & tabulador & tabulador & tabulador & "TOTAL CHEQUES A FECHA" & tabulador & tabulador & Format(fecha, "$ ###,###,##0"), True
'            impresion.Range(impresion.Rows - 1, 20, impresion.Rows - 1, 29).Merge
'            impresion.Range(impresion.Rows - 1, 30, impresion.Rows - 1, 35).Merge
'            impresion.Range(impresion.Rows - 1, 20, impresion.Rows - 1, 29).Alignment = cellLeftCenter
'            impresion.Range(impresion.Rows - 1, 30, impresion.Rows - 1, 35).Alignment = cellRightCenter
'            impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, impresion.Cols - 1).FontBold = True
'
'            impresion.AddItem tabulador & tabulador & tabulador & tabulador & "TOTAL CHEQUES AL DIA" & tabulador & tabulador & Format(dia, "$ ###,###,##0"), True
'            impresion.Range(impresion.Rows - 1, 20, impresion.Rows - 1, 29).Merge
'            impresion.Range(impresion.Rows - 1, 30, impresion.Rows - 1, 35).Merge
'            impresion.Range(impresion.Rows - 1, 20, impresion.Rows - 1, 29).Alignment = cellLeftCenter
'            impresion.Range(impresion.Rows - 1, 30, impresion.Rows - 1, 35).Alignment = cellRightCenter
'            impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, impresion.Cols - 1).FontBold = True
'
'            impresion.AddItem tabulador & tabulador & tabulador & tabulador & "TOTAL CHEQUES RECIBIDOS" & tabulador & tabulador & Format(dia + fecha, "$ ###,###,##0"), True
'            impresion.Range(impresion.Rows - 2, 20, impresion.Rows - 2, 35).Borders(cellEdgeBottom) = cellThin
'            impresion.Range(impresion.Rows - 1, 20, impresion.Rows - 1, 29).Merge
'            impresion.Range(impresion.Rows - 1, 30, impresion.Rows - 1, 35).Merge
'            impresion.Range(impresion.Rows - 1, 20, impresion.Rows - 1, 29).Alignment = cellLeftCenter
'            impresion.Range(impresion.Rows - 1, 30, impresion.Rows - 1, 35).Alignment = cellRightCenter
'            impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, impresion.Cols - 1).FontBold = True
'        End If
        impresion.AddItem "", True
        impresion.RowHeight(impresion.Rows - 1) = 0
    End Sub
    
    Private Sub imprimirUnaGrillaVentas()
        Dim tabla As String
        Dim cadena As String
        Dim Descuento As Double
        Dim neto As Double
        Dim iva As Double
        Dim retencion As Double
        Dim total As Double
        Dim TIPO As String
        Dim NUMERO As String
        Dim i As Integer
        Dim Cliente As String
        Dim tabulador As String
        Dim cajeras As String
        Dim caja As String
        
        tabulador = ""
        For i = 1 To 2
            tabulador = tabulador & vbTab
        Next i
        'DATOS
        cajeras = tcajera.text
        caja = tcaja.text
        
        tabla = "SELECT DISTINCT IF(dc.tipo = 'FV', 1, IF(dc.tipo = 'BV', 2, IF(dc.tipo = 'FE', 3, IF(dc.tipo = 'ZE', 4, IF(dc.tipo = 'GD', 5, 0))))) AS orden, CONCAT(CONCAT(dc.tipo, ' ', dc.foliosii), '" & tabulador & tabulador & vbTab & "', DATE_FORMAT(dc.fecha,'%d-%m-%Y'),  '" & tabulador & tabulador & "') AS item1, IF(dc.nula='N',dc.rut,'NULA') as rut, IF(dc.nula='N',CONCAT('$ ', FORMAT(dc.descuento,0), '" & tabulador & vbTab & "', '$ ', FORMAT(dc.neto,0), '" & tabulador & tabulador & "', '$ ', FORMAT(dc.iva,0), '" & tabulador & tabulador & "', '$ ', FORMAT(dc.impuestoharina+dc.impuestocarne+dc.impuestoilalicores+dc.impuestoilarefrescos+dc.impuestoilavinos,0), "
        tabla = tabla + "'" & tabulador & tabulador & "', '$ ', FORMAT(dc.total,0)),'') AS item2, IF(dc.nula='N',dc.descuento,0) AS descuento, IF(dc.nula='N',dc.neto,0) AS neto, IF(dc.nula='N',dc.iva,0) AS iva, IF(dc.nula='N',dc.impuestoharina+dc.impuestocarne+dc.impuestoilalicores+dc.impuestoilarefrescos+dc.impuestoilavinos,0) AS retencion, IF(dc.nula='N',dc.total,0) AS total, dc.tipo, dc.numero,dc.fecha,dc.caja "
        tabla = tabla & "FROM sv_documento_cabeza_" + PAuditoriaVentas.dato1.text + " AS dc "
        tabla = tabla & "WHERE local = '" & localAuditoria & "' AND (tipo <> 'FV' OR tipo <> 'NV' OR tipo <> 'BV' OR tipo <> 'ZE' OR tipo <> 'FE') AND fecha BETWEEN '" & fecha1 & "' AND '" & fecha2 & "' and cajera like '%" + cajeras + "%' and caja like '%" + caja + "%' ORDER BY orden ASC"
        
        Call ConectarControlData(data, servidor, baseVentas & localAuditoria, usuario, password, tabla)
        
        If data.Recordset.RecordCount > 0 Then
            salto = True
            'LISTADO DE VENTAS
            'TITULO
            impresion.AddItem "LISTADO DE VENTAS ", True
            impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, impresion.Cols - 1).Merge
            impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, impresion.Cols - 1).Alignment = cellCenterCenter
            impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, impresion.Cols - 1).Borders(cellEdgeBottom) = cellThin
            impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, impresion.Cols - 1).FontBold = True
            
            'CABEZA
            cadena = "DOCUMENTO" & tabulador & tabulador & vbTab
            cadena = cadena & "FECHA" & tabulador & tabulador
            cadena = cadena & "RUT" & tabulador & tabulador
            cadena = cadena & "NOMBRE" & tabulador & tabulador & vbTab
            cadena = cadena & "T.V." & tabulador
            cadena = cadena & "DESC." & tabulador & vbTab
            cadena = cadena & "NETO" & tabulador & tabulador
            cadena = cadena & "IVA" & tabulador & tabulador
            cadena = cadena & "RETENCION" & tabulador & tabulador
            cadena = cadena & "TOTAL"
            impresion.AddItem cadena, True
            impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, impresion.Cols - 1).Borders(cellEdgeBottom) = cellThin
            'UNION DE CELDAS
            impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, 5).Merge
            impresion.Range(impresion.Rows - 1, 6, impresion.Rows - 1, 9).Merge
            impresion.Range(impresion.Rows - 1, 10, impresion.Rows - 1, 13).Merge
            impresion.Range(impresion.Rows - 1, 14, impresion.Rows - 1, 18).Merge
            impresion.Range(impresion.Rows - 1, 19, impresion.Rows - 1, 20).Merge
            impresion.Range(impresion.Rows - 1, 21, impresion.Rows - 1, 23).Merge
            impresion.Range(impresion.Rows - 1, 24, impresion.Rows - 1, 27).Merge
            impresion.Range(impresion.Rows - 1, 28, impresion.Rows - 1, 31).Merge
            impresion.Range(impresion.Rows - 1, 32, impresion.Rows - 1, 35).Merge
            impresion.Range(impresion.Rows - 1, 36, impresion.Rows - 1, impresion.Cols - 1).Merge
            impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, impresion.Cols - 1).Alignment = cellCenterCenter
            impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, impresion.Cols - 1).Borders(cellEdgeBottom) = cellThin
            
            data.Recordset.MoveFirst
            Descuento = 0
            neto = 0
            iva = 0
            retencion = 0
            total = 0
            TIPO = data.Recordset.Fields("tipo")
            NUMERO = data.Recordset.Fields("numero")
            
            While Not data.Recordset.EOF
                If TIPO = data.Recordset.Fields("tipo") Then
                    NUMERO = data.Recordset.Fields("numero")
                    If data.Recordset.Fields("rut") = "NULA" Then
                        Cliente = data.Recordset.Fields("rut") & tabulador & tabulador & vbTab & leerNombreCliente(data.Recordset.Fields("rut")) & tabulador & tabulador & vbTab & tabulador
                    Else
                        Cliente = data.Recordset.Fields("rut") & tabulador & tabulador & vbTab & leerNombreCliente(data.Recordset.Fields("rut")) & tabulador & tabulador & vbTab & leerTipoPago(TIPO, NUMERO, data.Recordset.Fields("caja"), data.Recordset.Fields("fecha")) & tabulador
                    End If
                    Descuento = Descuento + data.Recordset.Fields("descuento")
                    neto = neto + data.Recordset.Fields("neto")
                    iva = iva + data.Recordset.Fields("iva")
                    retencion = retencion + data.Recordset.Fields("retencion")
                    total = total + data.Recordset.Fields("total")
                    impresion.AddItem data.Recordset.Fields("item1") & Cliente & Replace(data.Recordset.Fields("item2"), ",", "."), True
                    'UNION DE CELDAS
                    impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, 5).Merge
                    impresion.Range(impresion.Rows - 1, 6, impresion.Rows - 1, 9).Merge
                    impresion.Range(impresion.Rows - 1, 10, impresion.Rows - 1, 13).Merge
                    impresion.Range(impresion.Rows - 1, 14, impresion.Rows - 1, 18).Merge
                    impresion.Range(impresion.Rows - 1, 19, impresion.Rows - 1, 20).Merge
                    impresion.Range(impresion.Rows - 1, 21, impresion.Rows - 1, 23).Merge
                    impresion.Range(impresion.Rows - 1, 24, impresion.Rows - 1, 27).Merge
                    impresion.Range(impresion.Rows - 1, 28, impresion.Rows - 1, 31).Merge
                    impresion.Range(impresion.Rows - 1, 32, impresion.Rows - 1, 35).Merge
                    impresion.Range(impresion.Rows - 1, 36, impresion.Rows - 1, impresion.Cols - 1).Merge
                    impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, impresion.Cols - 1).Alignment = cellCenterCenter
                    'ALINEACION
                    impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, 5).Alignment = cellCenterCenter
                    impresion.Range(impresion.Rows - 1, 6, impresion.Rows - 1, 9).Alignment = cellCenterCenter
                    impresion.Range(impresion.Rows - 1, 10, impresion.Rows - 1, 13).Alignment = cellRightCenter
                    impresion.Range(impresion.Rows - 1, 14, impresion.Rows - 1, 18).Alignment = cellLeftCenter
                    impresion.Range(impresion.Rows - 1, 19, impresion.Rows - 1, 20).Alignment = cellCenterCenter
                    impresion.Range(impresion.Rows - 1, 21, impresion.Rows - 1, 23).Alignment = cellRightCenter
                    impresion.Range(impresion.Rows - 1, 24, impresion.Rows - 1, 27).Alignment = cellRightCenter
                    impresion.Range(impresion.Rows - 1, 28, impresion.Rows - 1, 31).Alignment = cellRightCenter
                    impresion.Range(impresion.Rows - 1, 32, impresion.Rows - 1, 35).Alignment = cellRightCenter
                    impresion.Range(impresion.Rows - 1, 36, impresion.Rows - 1, impresion.Cols - 1).Alignment = cellRightCenter
                    
                Else
                    Select Case TIPO
                        Case "BV"
                            impresion.AddItem "TOTAL BOLETAS" & tabulador & tabulador & tabulador & tabulador & tabulador & tabulador & tabulador & tabulador & tabulador & tabulador & tabulador & Format(Descuento, "$ ###,###,##0") & tabulador & vbTab & Format(neto, "$ ###,###,##0") & tabulador & tabulador & Format(iva, "$ ###,###,##0") & tabulador & vbTab & Format(retencion, "$ ###,###,##0") & tabulador & vbTab & Format(total, "$ ###,###,##0"), True
                        Case "FV"
                            impresion.AddItem "TOTAL FACTURAS" & tabulador & tabulador & tabulador & tabulador & tabulador & tabulador & tabulador & tabulador & tabulador & tabulador & tabulador & Format(Descuento, "$ ###,###,##0") & tabulador & vbTab & Format(neto, "$ ###,###,##0") & tabulador & tabulador & Format(iva, "$ ###,###,##0") & tabulador & vbTab & Format(retencion, "$ ###,###,##0") & tabulador & vbTab & Format(total, "$ ###,###,##0"), True
                        Case "ZE"
                            impresion.AddItem "TOTAL ZETAS" & tabulador & tabulador & tabulador & tabulador & tabulador & tabulador & tabulador & tabulador & tabulador & tabulador & tabulador & Format(Descuento, "$ ###,###,##0") & tabulador & vbTab & Format(neto, "$ ###,###,##0") & tabulador & tabulador & Format(iva, "$ ###,###,##0") & tabulador & vbTab & Format(retencion, "$ ###,###,##0") & tabulador & vbTab & Format(total, "$ ###,###,##0"), True
                        Case "GD"
                            impresion.AddItem "TOTAL GUIAS DE DESPACHO" & tabulador & tabulador & tabulador & tabulador & tabulador & tabulador & tabulador & tabulador & tabulador & tabulador & tabulador & Format(Descuento, "$ ###,###,##0") & tabulador & vbTab & Format(neto, "$ ###,###,##0") & tabulador & tabulador & Format(iva, "$ ###,###,##0") & tabulador & vbTab & Format(retencion, "$ ###,###,##0") & tabulador & vbTab & Format(total, "$ ###,###,##0"), True
                        Case "GM"
                            impresion.AddItem "TOTAL GUIAS DE MOLIENDA" & tabulador & tabulador & tabulador & tabulador & tabulador & tabulador & tabulador & tabulador & tabulador & tabulador & tabulador & Format(Descuento, "$ ###,###,##0") & tabulador & vbTab & Format(neto, "$ ###,###,##0") & tabulador & tabulador & Format(iva, "$ ###,###,##0") & tabulador & vbTab & Format(retencion, "$ ###,###,##0") & tabulador & vbTab & Format(total, "$ ###,###,##0"), True
                        Case "FE"
                            impresion.AddItem "TOTAL FACTURAS EXENTAS" & tabulador & tabulador & tabulador & tabulador & tabulador & tabulador & tabulador & tabulador & tabulador & tabulador & tabulador & Format(Descuento, "$ ###,###,##0") & tabulador & vbTab & Format(neto, "$ ###,###,##0") & tabulador & tabulador & Format(iva, "$ ###,###,##0") & tabulador & vbTab & Format(retencion, "$ ###,###,##0") & tabulador & vbTab & Format(total, "$ ###,###,##0"), True
                        Case "NV"
                            impresion.AddItem "TOTAL NOTAS DE CREDITO" & tabulador & tabulador & tabulador & tabulador & tabulador & tabulador & tabulador & tabulador & tabulador & tabulador & tabulador & Format(Descuento, "$ ###,###,##0") & tabulador & vbTab & Format(neto, "$ ###,###,##0") & tabulador & tabulador & Format(iva, "$ ###,###,##0") & tabulador & vbTab & Format(retencion, "$ ###,###,##0") & tabulador & vbTab & Format(total, "$ ###,###,##0"), True
                    End Select
                    'UNION
                    impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, 22).Merge
                    impresion.Range(impresion.Rows - 1, 21, impresion.Rows - 1, 23).Merge
                    impresion.Range(impresion.Rows - 1, 24, impresion.Rows - 1, 27).Merge
                    impresion.Range(impresion.Rows - 1, 28, impresion.Rows - 1, 31).Merge
                    impresion.Range(impresion.Rows - 1, 32, impresion.Rows - 1, 35).Merge
                    impresion.Range(impresion.Rows - 1, 36, impresion.Rows - 1, impresion.Cols - 1).Merge
                    impresion.Range(impresion.Rows - 2, 1, impresion.Rows - 2, impresion.Cols - 1).Borders(cellEdgeBottom) = cellThin
                    'ALINEACION
                    impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, 22).Alignment = cellCenterCenter
                    impresion.Range(impresion.Rows - 1, 21, impresion.Rows - 1, 23).Alignment = cellRightCenter
                    impresion.Range(impresion.Rows - 1, 24, impresion.Rows - 1, 27).Alignment = cellRightCenter
                    impresion.Range(impresion.Rows - 1, 28, impresion.Rows - 1, 31).Alignment = cellRightCenter
                    impresion.Range(impresion.Rows - 1, 32, impresion.Rows - 1, 35).Alignment = cellRightCenter
                    impresion.Range(impresion.Rows - 1, 36, impresion.Rows - 1, impresion.Cols - 1).Alignment = cellRightCenter
                    impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, impresion.Cols - 1).FontBold = True
                    
                    TIPO = data.Recordset.Fields("tipo")
                    Descuento = 0
                    neto = 0
                    iva = 0
                    retencion = 0
                    total = 0
                    impresion.AddItem "", True
                    data.Recordset.MovePrevious
                End If
                data.Recordset.MoveNext
            Wend
            Select Case TIPO
                Case "BV"
                    impresion.AddItem "TOTAL BOLETAS" & tabulador & tabulador & tabulador & tabulador & tabulador & tabulador & tabulador & tabulador & tabulador & tabulador & tabulador & Format(Descuento, "$ ###,###,##0") & tabulador & vbTab & Format(neto, "$ ###,###,##0") & tabulador & tabulador & Format(iva, "$ ###,###,##0") & tabulador & vbTab & Format(retencion, "$ ###,###,##0") & tabulador & vbTab & Format(total, "$ ###,###,##0"), True
                Case "FV"
                    impresion.AddItem "TOTAL FACTURAS" & tabulador & tabulador & tabulador & tabulador & tabulador & tabulador & tabulador & tabulador & tabulador & tabulador & tabulador & Format(Descuento, "$ ###,###,##0") & tabulador & vbTab & Format(neto, "$ ###,###,##0") & tabulador & tabulador & Format(iva, "$ ###,###,##0") & tabulador & vbTab & Format(retencion, "$ ###,###,##0") & tabulador & vbTab & Format(total, "$ ###,###,##0"), True
                Case "ZE"
                    impresion.AddItem "TOTAL ZETAS" & tabulador & tabulador & tabulador & tabulador & tabulador & tabulador & tabulador & tabulador & tabulador & tabulador & tabulador & Format(Descuento, "$ ###,###,##0") & tabulador & vbTab & Format(neto, "$ ###,###,##0") & tabulador & tabulador & Format(iva, "$ ###,###,##0") & tabulador & vbTab & Format(retencion, "$ ###,###,##0") & tabulador & vbTab & Format(total, "$ ###,###,##0"), True
                Case "GD"
                    impresion.AddItem "TOTAL GUIAS DE DESPACHO" & tabulador & tabulador & tabulador & tabulador & tabulador & tabulador & tabulador & tabulador & tabulador & tabulador & tabulador & Format(Descuento, "$ ###,###,##0") & tabulador & vbTab & Format(neto, "$ ###,###,##0") & tabulador & tabulador & Format(iva, "$ ###,###,##0") & tabulador & vbTab & Format(retencion, "$ ###,###,##0") & tabulador & vbTab & Format(total, "$ ###,###,##0"), True
                Case "GM"
                    impresion.AddItem "TOTAL GUIAS DE MOLIENDA" & tabulador & tabulador & tabulador & tabulador & tabulador & tabulador & tabulador & tabulador & tabulador & tabulador & tabulador & Format(Descuento, "$ ###,###,##0") & tabulador & vbTab & Format(neto, "$ ###,###,##0") & tabulador & tabulador & Format(iva, "$ ###,###,##0") & tabulador & vbTab & Format(retencion, "$ ###,###,##0") & tabulador & vbTab & Format(total, "$ ###,###,##0"), True
                Case "FE"
                    impresion.AddItem "TOTAL FACTURAS EXENTAS" & tabulador & tabulador & tabulador & tabulador & tabulador & tabulador & tabulador & tabulador & tabulador & tabulador & tabulador & Format(Descuento, "$ ###,###,##0") & tabulador & vbTab & Format(neto, "$ ###,###,##0") & tabulador & tabulador & Format(iva, "$ ###,###,##0") & tabulador & vbTab & Format(retencion, "$ ###,###,##0") & tabulador & vbTab & Format(total, "$ ###,###,##0"), True
                Case "NV"
                    impresion.AddItem "TOTAL NOTAS DE CREDITO" & tabulador & tabulador & tabulador & tabulador & tabulador & tabulador & tabulador & tabulador & tabulador & tabulador & tabulador & Format(Descuento, "$ ###,###,##0") & tabulador & vbTab & Format(neto, "$ ###,###,##0") & tabulador & tabulador & Format(iva, "$ ###,###,##0") & tabulador & vbTab & Format(retencion, "$ ###,###,##0") & tabulador & vbTab & Format(total, "$ ###,###,##0"), True
            End Select
            'UNION
            impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, 22).Merge
            impresion.Range(impresion.Rows - 1, 21, impresion.Rows - 1, 23).Merge
            impresion.Range(impresion.Rows - 1, 24, impresion.Rows - 1, 27).Merge
            impresion.Range(impresion.Rows - 1, 28, impresion.Rows - 1, 31).Merge
            impresion.Range(impresion.Rows - 1, 32, impresion.Rows - 1, 35).Merge
            impresion.Range(impresion.Rows - 1, 36, impresion.Rows - 1, impresion.Cols - 1).Merge
            impresion.Range(impresion.Rows - 2, 1, impresion.Rows - 2, impresion.Cols - 1).Borders(cellEdgeBottom) = cellThin
            'ALINEACION
            impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, 22).Alignment = cellCenterCenter
            impresion.Range(impresion.Rows - 1, 21, impresion.Rows - 1, 23).Alignment = cellRightCenter
            impresion.Range(impresion.Rows - 1, 24, impresion.Rows - 1, 27).Alignment = cellRightCenter
            impresion.Range(impresion.Rows - 1, 28, impresion.Rows - 1, 31).Alignment = cellRightCenter
            impresion.Range(impresion.Rows - 1, 32, impresion.Rows - 1, 35).Alignment = cellRightCenter
            impresion.Range(impresion.Rows - 1, 36, impresion.Rows - 1, impresion.Cols - 1).Alignment = cellRightCenter
            impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, impresion.Cols - 1).FontBold = True
            impresion.AddItem "", True
        Else
            salto = False
        End If
        impresion.AddItem "", True
        impresion.RowHeight(impresion.Rows - 1) = 0
    End Sub
    
    Private Sub imprimirUnaGrillaEgresos()
        Dim tabla As String
        Dim cadena As String
        Dim i As Integer
        Dim total As Double
        Dim tabulador As String
                
        tabulador = ""
        For i = 1 To 5
            tabulador = tabulador & vbTab
        Next i
        'DATOS
        tabla = "SELECT CONCAT(numero, '" & tabulador & "', mte.codigo, ' ', mte.nombre, '" & tabulador & vbTab & vbTab & vbTab & "', DATE_FORMAT(fecha,'%d-%m-%Y'), '" & tabulador & "', glosa, '" & tabulador & tabulador & tabulador & "', CONCAT('$ ', FORMAT(monto,0))) AS item, monto "
        tabla = tabla & "FROM sv_egresoscaja_" + PAuditoriaVentas.dato1.text + " AS ec INNER JOIN " & baseVentas & ".sv_maestrotipoegresoscaja AS mte ON ec.tipo = mte.codigo "
        tabla = tabla & "WHERE local = '" & localAuditoria & "' AND fecha BETWEEN '" & fecha1 & "' AND '" & fecha2 & "' ORDER BY numero"
        Call ConectarControlData(data, servidor, baseVentas & localAuditoria, usuario, password, tabla)
        
        If data.Recordset.RecordCount > 0 Then
            salto = True
            'DETALLE DE EGRESOS
            'TITULO
            impresion.AddItem "DETALLE DE EGRESOS", True
            impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, impresion.Cols - 1).Merge
            impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, impresion.Cols - 1).Alignment = cellCenterCenter
            impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, impresion.Cols - 1).Borders(cellEdgeBottom) = cellThin
            impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, impresion.Cols - 1).FontBold = True
            'CABEZA
            cadena = "NUMERO" & tabulador
            cadena = cadena & "TIPO" & tabulador & vbTab & vbTab & vbTab
            cadena = cadena & "FECHA" & tabulador
            cadena = cadena & "GLOSA" & tabulador & tabulador & tabulador
            cadena = cadena & "MONTO"
            impresion.AddItem cadena, True
            'UNION
            impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, 5).Merge
            impresion.Range(impresion.Rows - 1, 6, impresion.Rows - 1, 13).Merge
            impresion.Range(impresion.Rows - 1, 14, impresion.Rows - 1, 18).Merge
            impresion.Range(impresion.Rows - 1, 19, impresion.Rows - 1, 33).Merge
            impresion.Range(impresion.Rows - 1, 34, impresion.Rows - 1, impresion.Cols - 1).Merge
            
            impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, impresion.Cols - 1).Alignment = cellCenterCenter
            impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, impresion.Cols - 1).Borders(cellEdgeBottom) = cellThin
        
            data.Recordset.MoveFirst
            total = 0
            While Not data.Recordset.EOF
                total = total + data.Recordset.Fields("monto")
                impresion.AddItem Replace(data.Recordset.Fields("item"), ",", "."), True
                'UNION
                impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, 5).Merge
                impresion.Range(impresion.Rows - 1, 6, impresion.Rows - 1, 13).Merge
                impresion.Range(impresion.Rows - 1, 14, impresion.Rows - 1, 18).Merge
                impresion.Range(impresion.Rows - 1, 19, impresion.Rows - 1, 33).Merge
                impresion.Range(impresion.Rows - 1, 34, impresion.Rows - 1, impresion.Cols - 1).Merge
                'ALINEACION
                impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, 5).Alignment = cellCenterCenter
                impresion.Range(impresion.Rows - 1, 6, impresion.Rows - 1, 13).Alignment = cellLeftCenter
                impresion.Range(impresion.Rows - 1, 14, impresion.Rows - 1, 18).Alignment = cellCenterCenter
                impresion.Range(impresion.Rows - 1, 19, impresion.Rows - 1, 33).Alignment = cellLeftCenter
                impresion.Range(impresion.Rows - 1, 34, impresion.Rows - 1, impresion.Cols - 1).Alignment = cellRightCenter
                data.Recordset.MoveNext
            Wend
            impresion.AddItem tabulador & tabulador & tabulador & vbTab & vbTab & vbTab & "TOTAL EGRESOS" & tabulador & tabulador & tabulador & Format(total, "$ ###,###,##0"), True
            impresion.Range(impresion.Rows - 1, 19, impresion.Rows - 1, 33).Merge
            impresion.Range(impresion.Rows - 1, 34, impresion.Rows - 1, impresion.Cols - 1).Merge
            impresion.Range(impresion.Rows - 1, 19, impresion.Rows - 1, 33).Alignment = cellCenterCenter
            impresion.Range(impresion.Rows - 1, 34, impresion.Rows - 1, impresion.Cols - 1).Alignment = cellRightCenter
            impresion.Range(impresion.Rows - 1, 19, impresion.Rows - 1, impresion.Cols - 1).Borders(cellEdgeTop) = cellThin
            impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, impresion.Cols - 1).FontBold = True
        Else
            salto = False
        End If
        impresion.AddItem "", True
        impresion.RowHeight(impresion.Rows - 1) = 0
    End Sub
    
    Private Sub imprimirUnaGrillaPagos()
        Dim tabla As String
        Dim NUMERO As String
        Dim rut As String
        Dim dias As String
        Dim fecha As String
        Dim i As Integer
        Dim TIPO As String
        Dim Cliente As String
        Dim cadena As String
        Dim tabulador As String
        Dim sumaDocumentos As Double
        Dim SUMAPAGOS As Double
                
        tabulador = ""
        For i = 1 To 4
            tabulador = tabulador & vbTab
        Next i
        'DATOS
        tabla = "SELECT CONCAT(pc.numero, '" & tabulador & "', DATE_FORMAT(pc.fecha,'%d-%m-%Y'), '" & tabulador & "') AS item1, pc.rut, CONCAT('" & tabulador & tabulador & tabulador & vbTab & "', CASE pc.tipopago WHEN '1' THEN '1 EFECTIVO' WHEN '2' THEN '2 CHEQUE' WHEN '3' THEN '3 DEPOSITO' ELSE '' END, '" & tabulador & vbTab & vbTab & "', CONCAT('$ ', FORMAT(pc.monto,0))) AS item2, pc.numero, pc.rut, DATE_FORMAT(pc.fecha,'%d-%m-%Y') AS fecha, pc.tipopago, pc.monto "
        tabla = tabla & "FROM sv_pagos_cabeza_" & PAuditoriaVentas.dato1.text & " AS pc "
        tabla = tabla & "WHERE pc.local = '" & localAuditoria & "' AND pc.fecha BETWEEN '" & fecha1 & "' AND '" & fecha2 & "' ORDER BY pc.numero ASC"
        Call ConectarControlData(data, servidor, baseVentas & localAuditoria, usuario, password, tabla)
        
        If data.Recordset.RecordCount > 0 Then
            salto = True
            'LISTADO DE CLIENTES CON PAGOS
            'TITULO
            impresion.AddItem "LISTADO DE PAGO DE CLIENTES", True
            impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, impresion.Cols - 1).Merge
            impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, impresion.Cols - 1).Alignment = cellCenterCenter
            impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, impresion.Cols - 1).Borders(cellEdgeBottom) = cellThin
            impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, impresion.Cols - 1).FontBold = True
            'CABEZA
            'PRIMERA LINEA
            cadena = "NUMERO" & tabulador
            cadena = cadena & "FECHA" & tabulador
            cadena = cadena & "RUT" & tabulador
            cadena = cadena & "NOMBRE" & tabulador & tabulador & tabulador & vbTab
            cadena = cadena & "FORMA PAGO" & tabulador & vbTab & vbTab
            cadena = cadena & "MONTO PAGO" & tabulador
            cadena = cadena & "VENCIMIENTO"
            impresion.AddItem cadena, True
            'UNION
            impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, 4).Merge
            impresion.Range(impresion.Rows - 1, 5, impresion.Rows - 1, 8).Merge
            impresion.Range(impresion.Rows - 1, 9, impresion.Rows - 1, 12).Merge
            impresion.Range(impresion.Rows - 1, 13, impresion.Rows - 1, 25).Merge
            impresion.Range(impresion.Rows - 1, 26, impresion.Rows - 1, 31).Merge
            impresion.Range(impresion.Rows - 1, 32, impresion.Rows - 1, 35).Merge
            impresion.Range(impresion.Rows - 1, 36, impresion.Rows - 1, impresion.Cols - 1).Merge
            
            impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, impresion.Cols - 1).Alignment = cellCenterCenter
            'SEGUNDA LINEA
            cadena = tabulador & tabulador & tabulador & tabulador & vbTab & vbTab & "DOCUMENTO" & tabulador & vbTab
            cadena = cadena & "FECHA" & tabulador
            cadena = cadena & "MONTO" & tabulador
            cadena = cadena & "PLAZO"
            impresion.AddItem cadena, True
            'UNION
            impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, 18).Merge
            impresion.Range(impresion.Rows - 1, 19, impresion.Rows - 1, 23).Merge
            impresion.Range(impresion.Rows - 1, 24, impresion.Rows - 1, 27).Merge
            impresion.Range(impresion.Rows - 1, 28, impresion.Rows - 1, 31).Merge
            impresion.Range(impresion.Rows - 1, 32, impresion.Rows - 1, impresion.Cols - 1).Merge
            
            impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, impresion.Cols - 1).Alignment = cellCenterCenter
            impresion.Range(impresion.Rows - 1, 19, impresion.Rows - 1, impresion.Cols - 1).Borders(cellEdgeBottom) = cellThin
            impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, 18).Borders(cellEdgeTop) = cellThin
            impresion.Range(impresion.Rows - 1, 18, impresion.Rows - 1, 18).Borders(cellEdgeRight) = cellThin
            
            sumaDocumentos = 0
            SUMAPAGOS = 0
            data.Recordset.MoveFirst
            While Not data.Recordset.EOF
                '''''''''''''''''''''''''''''
                SUMAPAGOS = SUMAPAGOS + CDbl(data.Recordset.Fields("monto"))
                NUMERO = data.Recordset.Fields("numero")
                rut = data.Recordset.Fields("rut")
                fecha = data.Recordset.Fields("fecha")
                TIPO = data.Recordset.Fields("tipopago")
                Cliente = rut & tabulador & leerNombreCliente(rut)
                If TIPO = "2" Then
                    fecha = leerVencimientoPago(NUMERO)
                End If
                If TIPO = "3" Then
                    fecha = leerDepositoPago(NUMERO)
                End If
                cadena = data.Recordset.Fields("item1") & Cliente & data.Recordset.Fields("item2") & tabulador & fecha
                impresion.AddItem Replace(cadena, ",", "."), True
                'UNION
                impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, 4).Merge
                impresion.Range(impresion.Rows - 1, 5, impresion.Rows - 1, 8).Merge
                impresion.Range(impresion.Rows - 1, 9, impresion.Rows - 1, 12).Merge
                impresion.Range(impresion.Rows - 1, 13, impresion.Rows - 1, 25).Merge
                impresion.Range(impresion.Rows - 1, 26, impresion.Rows - 1, 31).Merge
                impresion.Range(impresion.Rows - 1, 32, impresion.Rows - 1, 35).Merge
                impresion.Range(impresion.Rows - 1, 36, impresion.Rows - 1, impresion.Cols - 1).Merge
                'ALINEACION
                impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, 4).Alignment = cellCenterCenter
                impresion.Range(impresion.Rows - 1, 5, impresion.Rows - 1, 8).Alignment = cellCenterCenter
                impresion.Range(impresion.Rows - 1, 9, impresion.Rows - 1, 12).Alignment = cellRightCenter
                impresion.Range(impresion.Rows - 1, 13, impresion.Rows - 1, 25).Alignment = cellLeftCenter
                impresion.Range(impresion.Rows - 1, 26, impresion.Rows - 1, 31).Alignment = cellLeftCenter
                impresion.Range(impresion.Rows - 1, 32, impresion.Rows - 1, 35).Alignment = cellRightCenter
                impresion.Range(impresion.Rows - 1, 36, impresion.Rows - 1, impresion.Cols - 1).Alignment = cellCenterCenter
                
               'LISTADO DE DOCUMENTOS PAGADOS
                
                'tabla = "SELECT CONCAT(pd.tipo, ' ', pd.documento, '" & tabulador & vbTab & "', DATE_FORMAT(dc.fecha,'%d-%m-%Y'), '" & tabulador & "', '$ ', FORMAT(pd.monto,0), '" & tabulador & "') AS item, DATE_FORMAT(dc.fecha,'%d-%m-%Y') AS fecha "
                'tabla = tabla & "FROM sv_pagos_detalle AS pd INNER JOIN sv_documento_cabeza AS dc ON pd.tipo = dc.tipo AND pd.documento = dc.numero "
                'tabla = tabla & "WHERE pd.local = '" & localAuditoria & "' AND pd.fecha BETWEEN '" & fecha1 & "' AND '" & fecha2 & "' AND pd.numero = '" & numero & "' "
                'tabla = tabla & "UNION "
                tabla = "SELECT CONCAT(pd.tipo, ' ', pd.documento, '" & tabulador & vbTab & "', DATE_FORMAT(dc.fechaemision,'%d-%m-%Y'), '" & tabulador & "', '$ ', FORMAT(pd.monto,0), '" & tabulador & "') AS item, DATE_FORMAT(dc.fechaemision,'%d-%m-%Y') AS fecha, pd.monto "
                tabla = tabla & "FROM sv_pagos_detalle_" & empresaActiva & " AS pd INNER JOIN sv_documentos_cobranza_" & empresaActiva & " AS dc ON pd.tipo = dc.tipo AND pd.documento = dc.numero "
                tabla = tabla & "WHERE pd.local = '" & localAuditoria & "' AND pd.fecha BETWEEN '" & fecha1 & "' AND '" & fecha2 & "' AND pd.numero = '" & NUMERO & "' "
                tabla = tabla & "ORDER BY fecha ASC"
                Call ConectarControlData(data2, servidor, baseVentas & rubroAuditoria, usuario, password, tabla)
                If data2.Recordset.RecordCount > 0 Then
                    data2.Recordset.MoveFirst
                    While Not data2.Recordset.EOF
                        sumaDocumentos = sumaDocumentos + CDbl(data2.Recordset.Fields("monto"))
                        cadena = tabulador & tabulador & tabulador & tabulador & vbTab & vbTab
                        cadena = cadena & Replace(data2.Recordset.Fields("item"), ",", ".")
                        dias = Str(DateDiff("d", data2.Recordset.Fields("fecha"), fecha))
                        dias = String(4 - Len(dias), "  ") & dias
                        cadena = cadena & "                     DIAS PAGO " & dias
                        impresion.AddItem cadena, True
                        'UNION
                        impresion.Range(impresion.Rows - 1, 19, impresion.Rows - 1, 23).Merge
                        impresion.Range(impresion.Rows - 1, 24, impresion.Rows - 1, 27).Merge
                        impresion.Range(impresion.Rows - 1, 28, impresion.Rows - 1, 31).Merge
                        impresion.Range(impresion.Rows - 1, 32, impresion.Rows - 1, impresion.Cols - 1).Merge
                        'ALINEACION
                        impresion.Range(impresion.Rows - 1, 19, impresion.Rows - 1, 23).Alignment = cellCenterCenter
                        impresion.Range(impresion.Rows - 1, 24, impresion.Rows - 1, 27).Alignment = cellCenterCenter
                        impresion.Range(impresion.Rows - 1, 28, impresion.Rows - 1, 31).Alignment = cellRightCenter
                        impresion.Range(impresion.Rows - 1, 32, impresion.Rows - 1, impresion.Cols - 1).Alignment = cellLeftCenter
                        data2.Recordset.MoveNext
                    Wend
                End If
                impresion.Range(impresion.Rows - 1, 19, impresion.Rows - 1, impresion.Cols - 1).Borders(cellEdgeBottom) = cellThin
                impresion.AddItem "", True
                data.Recordset.MoveNext
            Wend
        Else
            salto = False
        End If
        If sumaDocumentos <> 0 Then
        impresion.AddItem "", True
        cadena = tabulador & tabulador & tabulador & tabulador & tabulador & vbTab & vbTab & vbTab
        cadena = cadena & "TOTALES" & tabulador & Format(sumaDocumentos, "$ ###,###,##0") & tabulador & vbTab & Format(SUMAPAGOS, "$ ###,###,##0")
        impresion.AddItem cadena, True
        'UNION
        impresion.Range(impresion.Rows - 1, 24, impresion.Rows - 1, 27).Merge
        impresion.Range(impresion.Rows - 1, 28, impresion.Rows - 1, 31).Merge
        impresion.Range(impresion.Rows - 1, 32, impresion.Rows - 1, 35).Merge
        'ALINEACION
        impresion.Range(impresion.Rows - 1, 24, impresion.Rows - 1, 27).Alignment = cellCenterCenter
        impresion.Range(impresion.Rows - 1, 28, impresion.Rows - 1, 31).Alignment = cellRightCenter
        impresion.Range(impresion.Rows - 1, 32, impresion.Rows - 1, 35).Alignment = cellRightCenter
        
        impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, impresion.Cols - 1).FontBold = True
        impresion.Range(impresion.Rows - 1, 24, impresion.Rows - 1, impresion.Cols - 1).Borders(cellEdgeTop) = cellThin
        
        impresion.AddItem ""
        impresion.RowHeight(impresion.Rows - 1) = 0
    End If
    End Sub
    
    Private Sub imprimirUnaGrillaMovimientos()
        Dim tabla As String
        Dim cadena As String
        Dim cad As String
        Dim cadaux As Variant
        Dim i As Integer
        Dim rut As String
        Dim tabulador As String
        Dim cupo As Double
        Dim Cheques As Double
        Dim facturas As Double
        
        tabulador = ""
        For i = 1 To 4
            tabulador = tabulador & vbTab
        Next i
        'DATOS
        tabla = "SELECT DISTINCT dc.rut "
        tabla = tabla & "FROM sv_documentos_cobranza_" & empresaActiva & " As dc "
        tabla = tabla & "WHERE dc.local = '" & localAuditoria & "' AND dc.fechaemision BETWEEN '" & fecha1 & "' AND '" & fecha2 & "' AND dc.rut <> '9999999996' AND dc.rut <> '0000000019' "
        tabla = tabla & "Union "
        tabla = tabla & "SELECT DISTINCT dc.rut "
        tabla = tabla & "FROM sv_pagos_cabeza_" & empresaActiva & " As dc "
        tabla = tabla & "WHERE dc.local = '" & localAuditoria & "' AND dc.fecha BETWEEN '" & fecha1 & "' AND '" & fecha2 & "' AND dc.rut <> '9999999996' AND dc.rut <> '0000000019' "
        tabla = tabla & "GROUP BY rut ORDER BY rut ASC"
        Call ConectarControlData(data, servidor, baseVentas & rubroAuditoria, usuario, password, tabla)
        
        If data.Recordset.RecordCount > 0 Then
            'DETALLE DE EGRESOS
            'TITULO
            impresion.AddItem "LISTADO DE CLIENTES CON MOVIMIENTOS", True
            impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, impresion.Cols - 1).Merge
            impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, impresion.Cols - 1).Alignment = cellCenterCenter
            impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, impresion.Cols - 1).Borders(cellEdgeBottom) = cellThin
            impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, impresion.Cols - 1).FontBold = True
            'CABEZA
            cadena = "RUT" & tabulador
            cadena = cadena & "NOMBRE" & tabulador & tabulador & tabulador & vbTab & vbTab
            cadena = cadena & "CUPO" & tabulador & vbTab
            cadena = cadena & "C.CHE" & vbTab & vbTab
            cadena = cadena & "CHEQUES" & tabulador
            cadena = cadena & "C.FAC" & vbTab & vbTab
            cadena = cadena & "FACTURAS" & tabulador
            cadena = cadena & "SALDO"
            impresion.AddItem cadena, True
            'UNION
            impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, 4).Merge
            impresion.Range(impresion.Rows - 1, 5, impresion.Rows - 1, 18).Merge
            impresion.Range(impresion.Rows - 1, 19, impresion.Rows - 1, 23).Merge
            impresion.Range(impresion.Rows - 1, 24, impresion.Rows - 1, 25).Merge
            impresion.Range(impresion.Rows - 1, 26, impresion.Rows - 1, 29).Merge
            impresion.Range(impresion.Rows - 1, 30, impresion.Rows - 1, 31).Merge
            impresion.Range(impresion.Rows - 1, 32, impresion.Rows - 1, 35).Merge
            impresion.Range(impresion.Rows - 1, 36, impresion.Rows - 1, impresion.Cols - 1).Merge
            
            impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, impresion.Cols - 1).Alignment = cellCenterCenter
            impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, impresion.Cols - 1).Borders(cellEdgeBottom) = cellThin
        
            data.Recordset.MoveFirst
            While Not data.Recordset.EOF
                rut = data.Recordset.Fields("rut")
                cadena = rut & tabulador
                cadena = cadena & leerNombreCliente(data.Recordset.Fields("rut")) & tabulador & tabulador & tabulador & vbTab & vbTab
                
                cupo = CDbl(leerCupoCliente(data.Recordset.Fields("rut")))
                cadena = cadena & Format(cupo, "###,###,##0") & tabulador & vbTab
                
                cad = LEERCHEQUES(rut, Format(fecha2, "yyyy-mm-dd"), vbTab & vbTab)
                cadaux = Split(cad, vbTab & vbTab)
                Cheques = CDbl(cadaux(1))
                cadena = cadena & cad & tabulador
                
                cad = leerFacturas(rut, vbTab & vbTab, data2)
                cadaux = Split(cad, vbTab & vbTab)
                facturas = CDbl(cadaux(1))
                cadena = cadena & cad & tabulador
                
                cupo = cupo - Cheques - facturas
                cadena = cadena & Format(cupo, " $ ###,###,##0")
                impresion.AddItem cadena, True
                
                'UNION
                impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, 4).Merge
                impresion.Range(impresion.Rows - 1, 5, impresion.Rows - 1, 18).Merge
                impresion.Range(impresion.Rows - 1, 19, impresion.Rows - 1, 23).Merge
                impresion.Range(impresion.Rows - 1, 24, impresion.Rows - 1, 25).Merge
                impresion.Range(impresion.Rows - 1, 26, impresion.Rows - 1, 29).Merge
                impresion.Range(impresion.Rows - 1, 30, impresion.Rows - 1, 31).Merge
                impresion.Range(impresion.Rows - 1, 32, impresion.Rows - 1, 35).Merge
                impresion.Range(impresion.Rows - 1, 36, impresion.Rows - 1, impresion.Cols - 1).Merge
                'ALINEACION
                impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, 4).Alignment = cellRightCenter
                impresion.Range(impresion.Rows - 1, 5, impresion.Rows - 1, 18).Alignment = cellLeftCenter
                impresion.Range(impresion.Rows - 1, 19, impresion.Rows - 1, 23).Alignment = cellRightCenter
                impresion.Range(impresion.Rows - 1, 24, impresion.Rows - 1, 25).Alignment = cellRightCenter
                impresion.Range(impresion.Rows - 1, 26, impresion.Rows - 1, 29).Alignment = cellRightCenter
                impresion.Range(impresion.Rows - 1, 30, impresion.Rows - 1, 31).Alignment = cellRightCenter
                impresion.Range(impresion.Rows - 1, 32, impresion.Rows - 1, 35).Alignment = cellRightCenter
                impresion.Range(impresion.Rows - 1, 36, impresion.Rows - 1, impresion.Cols - 1).Alignment = cellRightCenter
                data.Recordset.MoveNext
            Wend
        End If
        impresion.AddItem "", True
    End Sub
    
Private Sub FrmNegativo_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
        FrmNegativo.ColorBarraAbajo = &HFFC0C0
        FrmNegativo.ColorBarraArriba = &H800000
        FrmNegativo.CaptionEstilo3D = Raised
End Sub

Private Sub FrmNegativo_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
        FrmNegativo.ColorBarraAbajo = &H800000
        FrmNegativo.ColorBarraArriba = &HFFC0C0
        FrmNegativo.CaptionEstilo3D = Inserted
        Load ventasnegativas
        ventasnegativas.Show
        ventasnegativas.dato1.SetFocus
End Sub

Private Sub frmventas_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
        frmventas.ColorBarraAbajo = &HFFC0C0
        frmventas.ColorBarraArriba = &H800000
        frmventas.CaptionEstilo3D = Raised
End Sub

Private Sub frmventas_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
        frmventas.ColorBarraAbajo = &H800000
        frmventas.ColorBarraArriba = &HFFC0C0
        frmventas.CaptionEstilo3D = Inserted
        Load VentasxHoras
        With VentasxHoras
        .dato1.text = dato1.text
        .dato2.text = dato2.text
        .dato3.text = dato3.text
        .dato4.text = dato4.text
        .dato5.text = dato5.text
        .dato6.text = dato6.text
        .dato7.text = dato7.text
        
        .Show
        .dato1.SetFocus
        End With
        
End Sub

    Private Sub Ingresos_DblClick()
'        Load AuditoriaDetalles
'        AuditoriaDetalles.fechaini = fecha1
'        AuditoriaDetalles.fechafin = fecha2
'        AuditoriaDetalles.codLocal = localAuditoria
'        AuditoriaDetalles.codRubro = rubroAuditoria
'        AuditoriaDetalles.informe = "INGRESOS"
'        Select Case Ingresos.MouseRow
'            Case 1
'                titulo = "EFECTIVO"
'                AuditoriaDetalles.tipoInforme = Ingresos.MouseRow
'            Case 2
'                titulo = "LISTADO DE CHEQUES"
'                AuditoriaDetalles.tipoInforme = Ingresos.MouseRow
'            Case 3
'                titulo = "CREDITOS"
'                AuditoriaDetalles.tipoInforme = Ingresos.MouseRow
'            Case 4
'                titulo = "DEPOSITOS"
'                AuditoriaDetalles.tipoInforme = Ingresos.MouseRow
'            Case 5
'                titulo = "PAGO DE CLIENTES"
'                AuditoriaDetalles.tipoInforme = Ingresos.MouseRow
'            Case 6
'                titulo = "PAGO DE CLIENTES"
'                AuditoriaDetalles.tipoInforme = Ingresos.MouseRow
'            Case 7
'                titulo = "PAGO DE CLIENTES"
'                AuditoriaDetalles.tipoInforme = Ingresos.MouseRow
'            Case 8
'                titulo = "PAGO DE CLIENTES"
'                AuditoriaDetalles.tipoInforme = Ingresos.MouseRow
'            Case 9
'                titulo = "PAGO DE CLIENTES"
'                AuditoriaDetalles.tipoInforme = Ingresos.MouseRow
'            Case 10
'                titulo = "PAGO DE CLIENTES"
'                AuditoriaDetalles.tipoInforme = Ingresos.MouseRow
'            Case Else
'                AuditoriaDetalles.tipoInforme = 0
'        End Select
'        If AuditoriaDetalles.tipoInforme <> 0 Then
'            AuditoriaDetalles.Show vbModal
'        End If

'If Ingresos.Selection.FirstRow = 15 Then MsgBox "kj"

    End Sub
    
    Private Sub Egresos_DblClick()
        Load AuditoriaDetalles
        AuditoriaDetalles.fechaini = fecha1
        AuditoriaDetalles.fechafin = fecha2
        AuditoriaDetalles.codLocal = localAuditoria
        AuditoriaDetalles.codRubro = rubroAuditoria
        AuditoriaDetalles.informe = "EGRESOS"
        Select Case Egresos.MouseRow
            Case 1
                titulo = "EGRESOS DE CAJA"
                AuditoriaDetalles.tipoInforme = Egresos.MouseRow
            Case 2
                titulo = "CHEQUES A FECHA"
                AuditoriaDetalles.tipoInforme = Egresos.MouseRow
            Case 3
                titulo = "VENTAS CON CREDITO"
                AuditoriaDetalles.tipoInforme = Egresos.MouseRow
            Case 4
                titulo = "DEPOSITOS"
                AuditoriaDetalles.tipoInforme = Egresos.MouseRow
            'Case 6
            '    AuditoriaDetalles.tipoInforme = Egresos.MouseRow
            'Case 7
            '    AuditoriaDetalles.tipoInforme = Egresos.MouseRow
            'Case 8
            '    AuditoriaDetalles.tipoInforme = Egresos.MouseRow
            Case Else
                AuditoriaDetalles.tipoInforme = 0
        End Select
        If AuditoriaDetalles.tipoInforme <> 0 Then
            AuditoriaDetalles.Show vbModal
        End If
    End Sub
Sub cargartipospagos()
             
        Dim csql As rdoQuery
        Dim resultado As rdoResultset
        Dim i As Integer
        Dim diasmora As Double
        Dim saldo As Double
        Dim interes As Double
        Dim cSql2 As rdoQuery
        Dim resultado2 As rdoResultset
        Dim K As Double
        
        
        Set cSql2 = New rdoQuery
        Set cSql2.ActiveConnection = ventas

        cSql2.sql = "SELECT * "
        cSql2.sql = cSql2.sql & "FROM sv_tiposdepagoclientes "
        cSql2.Execute
        Ingresos.Rows = 1
        If cSql2.RowsAffected > 0 Then
        i = cSql2.RowsAffected
        End If
        cSql2.Close
        Set cSql2 = Nothing
        
'        Ingresos.Rows = i + 3
       For K = 1 To i
        
        Set csql = New rdoQuery
        Set csql.ActiveConnection = ventas
        pivote.MaxLength = 2
        pivote.text = K
        pivote.text = ceros(pivote)
        
        csql.sql = "SELECT codigo,nombre "
        csql.sql = csql.sql & "FROM sv_tiposdepagoclientes "
        csql.sql = csql.sql & "where codigo='" & pivote.text & "' "
        csql.Execute
        If csql.RowsAffected > 0 Then
        Set resultado = csql.OpenResultset
        While Not resultado.EOF
        
        If resultado(1) <> "VUELTO" Then
        Ingresos.Rows = Ingresos.Rows + 1
        Ingresos.Cell(K, 0).text = "TOTAL " + resultado(1)
        End If
        
            resultado.MoveNext
            Wend

        End If
        Set resultado = Nothing
        csql.Close
        Set csql = Nothing
        Next K
        Ingresos.Rows = Ingresos.Rows + 6
        Ingresos.Cell(i + 1, 0).text = "TOTAL INGRESO X VENTA"
        Ingresos.Cell(i + 2, 0).text = "TOTAL PAGOS C.DIRECTO"
        Ingresos.Cell(i + 3, 0).text = "TOTAL PAGOS C.TMP"
        Ingresos.Cell(i + 4, 0).text = "TOTAL GENERAL INGRESOS"
'ZURITA
        Ingresos.Cell(i + 5, 0).text = "TOTAL EMPRESAS RELACIONADAS"
        Ingresos.Cell(i + 6, 0).text = "TOTAL GENERAL "
        Ingresos.Range(i + 5, 0, i + 6, 1).BackColor = vbRed
        Ingresos.Range(i + 5, 0, i + 6, 1).FontBold = True

End Sub

  

Private Sub tcaja_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
If leerNombreCaja(tcaja.text) = "" Then
tcaja.SetFocus
Else
nombrecaja.Caption = leerNombreCaja(tcaja.text)
End If


End If

End Sub

Private Sub tcajera_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyF2 Then
    Call ayudacajera(tcajera)
   End If
End Sub

Private Sub tcajera_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
tcajera.text = ceros(tcajera)
lbldv.Caption = rut(tcajera.text)

nombrecajera.Caption = leerNombreCajera(tcajera.text + lbldv.Caption)
If nombrecajera.Caption = "" Then
tcajera.text = ""
tcajera.SetFocus

End If
End If
End Sub
Public Sub auditoriadeafuera()
    Call dato1_KeyPress(13)
    Call dato7_KeyPress(13)
    Command1_Click
End Sub


Private Function TotalEmpresaRelacionada(ByRef loc, fecha1, fecha2) As Long
On Error GoTo fin
Dim csql As rdoQuery
Dim resultado As rdoResultset
Set csql = New rdoQuery
Set csql.ActiveConnection = ventas
csql.sql = "select  sum(total) from "
csql.sql = csql.sql & clientesistema & "ventas" & loc & ".sv_documento_cabeza_" & loc & " AS dc "
csql.sql = csql.sql & " where dc.rut  IN(select concat('0',mid(rut,1,8),mid(rut,10,10)) from " & clientesistema & "conta.maestroempresas)"
csql.sql = csql.sql & " and dc.local = '" & loc & "' AND dc.fecha BETWEEN '" & fecha1 & "' AND '" & fecha2 & "'"
csql.sql = csql.sql & " AND  caja < '90' and cajera <> '' and (tipo = 'FV' OR TIPO = 'BV')"

csql.Execute
If csql.RowsAffected > 0 Then
    Set resultado = csql.OpenResultset
    TotalEmpresaRelacionada = resultado(0)
Else
    TotalEmpresaRelacionada = "0"
End If
resultado.Close
csql.Close
fin:
End Function
