VERSION 5.00
Object = "{4777436C-EB8C-4596-98A8-EBCF98683969}#1.0#0"; "FlexCell.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{ADBBAED6-B16F-40DC-80DF-B44910CBA76C}#1.0#0"; "Frame-Xp.ocx"
Begin VB.Form imprimeCuponera 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Impresión de Cuponera de Crédito"
   ClientHeight    =   6585
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10455
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6585
   ScaleWidth      =   10455
   Begin XPFrame.FrameXp frmTodo 
      Height          =   375
      Left            =   120
      TabIndex        =   30
      Top             =   6120
      Width           =   1665
      _ExtentX        =   2937
      _ExtentY        =   661
      BackColor       =   49344
      Caption         =   "Sel. Todo"
      CaptionEstilo3D =   1
      BackColor       =   49344
      ColorBarraArriba=   16777152
      ColorBarraAbajo =   8421376
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
      Height          =   330
      Left            =   120
      Top             =   6120
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
   Begin XPFrame.FrameXp FrameXp2 
      Height          =   1815
      Left            =   60
      TabIndex        =   9
      Top             =   60
      Width           =   10335
      _ExtentX        =   18230
      _ExtentY        =   3201
      BackColor       =   16744576
      Caption         =   "Datos de la Venta"
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
      Begin VB.TextBox dato6 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E1FFFD&
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
         Left            =   2460
         MaxLength       =   4
         TabIndex        =   5
         Tag             =   "proveedor"
         Top             =   1440
         Width           =   735
      End
      Begin VB.TextBox dato5 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E1FFFD&
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
         Left            =   1920
         MaxLength       =   2
         TabIndex        =   4
         Tag             =   "proveedor"
         Top             =   1440
         Width           =   555
      End
      Begin VB.TextBox dato4 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E1FFFD&
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
         Left            =   1380
         MaxLength       =   2
         TabIndex        =   3
         Tag             =   "proveedor"
         Top             =   1440
         Width           =   555
      End
      Begin VB.ComboBox cmbTipo 
         Appearance      =   0  'Flat
         BackColor       =   &H00E1FFFD&
         Height          =   315
         ItemData        =   "imprimeCuponera.frx":0000
         Left            =   1380
         List            =   "imprimeCuponera.frx":000D
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   720
         Width           =   795
      End
      Begin VB.TextBox dato3 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E1FFFD&
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
         Left            =   5520
         MaxLength       =   10
         TabIndex        =   2
         Tag             =   "proveedor"
         Top             =   720
         Width           =   1395
      End
      Begin VB.TextBox dato7 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E1FFFD&
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
         Left            =   4980
         MaxLength       =   3
         TabIndex        =   6
         Tag             =   "proveedor"
         Top             =   1440
         Width           =   555
      End
      Begin VB.TextBox dato2 
         Appearance      =   0  'Flat
         BackColor       =   &H00E1FFFD&
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
         Left            =   1380
         MaxLength       =   2
         TabIndex        =   8
         Tag             =   "proveedor"
         Top             =   720
         Visible         =   0   'False
         Width           =   435
      End
      Begin VB.TextBox dato1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E1FFFD&
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
         Left            =   1380
         MaxLength       =   9
         TabIndex        =   0
         Tag             =   "proveedor"
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Prim. Vencim."
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   120
         TabIndex        =   28
         Top             =   1440
         Width           =   1155
      End
      Begin VB.Label lblSaldo 
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
         Left            =   8580
         TabIndex        =   27
         Top             =   1080
         Width           =   1635
      End
      Begin VB.Label lbl7 
         Appearance      =   0  'Flat
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Saldo"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   7320
         TabIndex        =   26
         Top             =   1080
         Width           =   1155
      End
      Begin VB.Label lblAbono 
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
         Left            =   4980
         TabIndex        =   25
         Top             =   1080
         Width           =   1635
      End
      Begin VB.Label lbl6 
         Appearance      =   0  'Flat
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Abono"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   3720
         TabIndex        =   24
         Top             =   1080
         Width           =   1155
      End
      Begin VB.Label lblMonto 
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
         Left            =   1380
         TabIndex        =   23
         Top             =   1080
         Width           =   1635
      End
      Begin VB.Label lbl8 
         Appearance      =   0  'Flat
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Núm. Cuotas"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   3720
         TabIndex        =   21
         Top             =   1440
         Width           =   1155
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
         Left            =   9540
         TabIndex        =   20
         Top             =   720
         Width           =   675
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
         Left            =   9060
         TabIndex        =   19
         Top             =   720
         Width           =   495
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
         Left            =   8580
         TabIndex        =   18
         Top             =   720
         Width           =   495
      End
      Begin VB.Label lblTipo 
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
         Left            =   2220
         TabIndex        =   17
         Top             =   720
         Width           =   1935
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
         Left            =   3120
         TabIndex        =   16
         Top             =   360
         Width           =   7095
      End
      Begin VB.Label lbl3 
         Appearance      =   0  'Flat
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Número Doc."
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   4260
         TabIndex        =   15
         Top             =   720
         Width           =   1155
      End
      Begin VB.Label lbl2 
         Appearance      =   0  'Flat
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Tipo Doc"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   120
         TabIndex        =   14
         Top             =   720
         Width           =   1155
      End
      Begin VB.Label lbl5 
         Appearance      =   0  'Flat
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Monto"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   120
         TabIndex        =   13
         Top             =   1080
         Width           =   1155
      End
      Begin VB.Label lbl4 
         Appearance      =   0  'Flat
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Fecha Compra"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   7320
         TabIndex        =   12
         Top             =   720
         Width           =   1155
      End
      Begin VB.Label lbl1 
         Appearance      =   0  'Flat
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Rut Cliente"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   120
         TabIndex        =   11
         Top             =   360
         Width           =   1155
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
         Left            =   2580
         TabIndex        =   10
         Top             =   360
         Width           =   495
      End
   End
   Begin XPFrame.FrameXp FrameXp1 
      Height          =   4095
      Left            =   60
      TabIndex        =   22
      Top             =   1920
      Width           =   10335
      _ExtentX        =   18230
      _ExtentY        =   7223
      BackColor       =   16744576
      Caption         =   "Detalle de las Cuotas"
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
      Begin FlexCell.Grid detalle 
         Height          =   3570
         Left            =   120
         TabIndex        =   7
         Top             =   420
         Width           =   10095
         _ExtentX        =   17806
         _ExtentY        =   6297
         Cols            =   5
         DefaultFontSize =   8.25
         Rows            =   13
         SelectionMode   =   1
         DateFormat      =   2
      End
   End
   Begin XPFrame.FrameXp frmImprimir 
      Height          =   375
      Left            =   7080
      TabIndex        =   29
      Top             =   6120
      Width           =   3285
      _ExtentX        =   5794
      _ExtentY        =   661
      BackColor       =   49344
      Caption         =   "I   M   P   R   I   M   I   R"
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
   Begin XPFrame.FrameXp frmNada 
      Height          =   375
      Left            =   1980
      TabIndex        =   31
      Top             =   6120
      Width           =   1665
      _ExtentX        =   2937
      _ExtentY        =   661
      BackColor       =   49344
      Caption         =   "Sel. Nada"
      CaptionEstilo3D =   1
      BackColor       =   49344
      ColorBarraArriba=   16777152
      ColorBarraAbajo =   8421376
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
   Begin FlexCell.Grid impresion 
      Height          =   315
      Left            =   2400
      TabIndex        =   32
      Top             =   6180
      Visible         =   0   'False
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   556
      Cols            =   5
      DefaultFontSize =   8.25
      Rows            =   30
   End
End
Attribute VB_Name = "imprimeCuponera"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Private c As Cuponera
    Private formatoGrilla(10, 10) As String
    Private modifica As Boolean
    Private folio As String
    Private grabado As Boolean

Private Sub cmbTipo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        dato2.text = Left(cmbTipo.List(cmbTipo.ListIndex), 2)
        Call dato2_KeyPress(13)
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
        Principal.barraEstado.Panels(2).text = "F2: Ayuda Clientes"
    End Sub
    
    Private Sub dato2_GotFocus()
        Call VerificarCajas(Me, dato2)
        Call selecciona(dato2)
    End Sub
    
    Private Sub dato3_GotFocus()
        Call VerificarCajas(Me, dato3)
        Call selecciona(dato3)
        Principal.barraEstado.Panels(2).text = "F2: Ayuda Documentos x Cliente"
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
        If KeyCode = vbKeyF2 Then
            Call ayudaCliente(dato1)
        Else
            Call Flechas(KeyCode, dato1)
        End If
    End Sub
    
    Private Sub dato2_KeyDown(KeyCode As Integer, Shift As Integer)
        Call Flechas(KeyCode, dato1)
    End Sub
    
    Private Sub dato3_KeyDown(KeyCode As Integer, Shift As Integer)
        If KeyCode = vbKeyF2 Then
            Call ayudaDocumentoCliente(dato3, dato2.text, dato1.text & lblDV.Caption)
        Else
            Call Flechas(KeyCode, cmbTipo)
        End If
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
            lblDV.Caption = rut(dato1.text)
            lblNombre.Caption = leerNombreCliente(dato1.text & lblDV.Caption)
            If lblNombre.Caption <> "" Then
                Call HabilitarCajas(Me, modifica)
                SendKeys "{Tab}"
            End If
        End If
    End Sub
    
    Private Sub dato2_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
            lblTipo.Caption = leerNombreDocumento(dato2.text)
            If lblTipo.Caption <> "" Then
                SendKeys "{Tab}"
            End If
        End If
    End Sub
    
    Private Sub dato3_KeyPress(KeyAscii As Integer)
        KeyAscii = esNumero(KeyAscii)
        If KeyAscii = 13 Then
            dato3.text = ceros(dato3)
            If leerCuponera(c, dato1.text & lblDV.Caption, dato2.text, dato3.text, "=", detalle, data) = True Then
                Call structtoctrl
                grabado = True
            Else
                detalle.Rows = 1
                Call HabilitarCajas(Me, modifica)
                If leerDocumento(c.cabeza, dato1.text & lblDV.Caption, dato2.text, dato3.text) = True Then
                    lblMonto.Caption = Format(c.cabeza.total, "$ ###,###,##0")
                    lblAbono.Caption = Format(c.cabeza.abono, "$ ###,###,##0")
                    lblSaldo.Caption = Format(c.cabeza.total - c.cabeza.abono, "$ ###,###,##0")
                    lblDia.Caption = Format(c.cabeza.fechaCompra, "dd")
                    lblMes.Caption = Format(c.cabeza.fechaCompra, "mm")
                    lblAño.Caption = Format(c.cabeza.fechaCompra, "yyyy")
                Else
                    Call MsgBox("El Documento " & dato2.text & " " & dato3.text & " no existe", vbOKOnly, "Error")
                End If
            End If
            SendKeys "{Tab}"
        End If
    End Sub
    
    Private Sub dato4_KeyPress(KeyAscii As Integer)
        KeyAscii = esNumero(KeyAscii)
        If KeyAscii = 13 Then
            dato4.text = ceros(dato4)
            If dato4.text = "00" Then
                dato4.text = Format(Now, "dd")
            End If
            SendKeys "{Tab}"
        End If
    End Sub
    
    Private Sub dato5_KeyPress(KeyAscii As Integer)
        KeyAscii = esNumero(KeyAscii)
        If KeyAscii = 13 Then
            dato5.text = ceros(dato5)
            If dato5.text = "00" Then
                dato5.text = Format(Now, "mm")
            End If
            SendKeys "{Tab}"
        End If
    End Sub
    
    Private Sub dato6_KeyPress(KeyAscii As Integer)
        KeyAscii = esNumero(KeyAscii)
        If KeyAscii = 13 Then
            dato6.text = ceros(dato6)
            If dato6.text = "0000" Then
                dato6.text = Format(Now, "yyyy")
            End If
            SendKeys "{Tab}"
        End If
    End Sub
    
    Private Sub dato7_KeyPress(KeyAscii As Integer)
        KeyAscii = esNumero(KeyAscii)
        If KeyAscii = 13 And dato7.text <> "" Then
            Call calculaCuotas
            Call ctrltostruct
        End If
    End Sub
    '========================================================
    'KeyPress
    '========================================================
    
    '========================================================
    'LostFocus
    '========================================================
    Private Sub dato1_LostFocus()
        Call limpiaBarra(2)
    End Sub
    
    Private Sub dato3_LostFocus()
        Call limpiaBarra(2)
    End Sub
    '========================================================
    'LostFocus
    '========================================================
'============================================================
'MANEJO DE LOS CONTOLES
'============================================================

'****************************************************************************
'Formato de la Grilla Detalle
'****************************************************************************
    Private Sub CargaGrillaDetalle(ByVal row As Integer, ByVal col As Integer)
        Dim i As Integer
        Rem DATOS DE LA COLUMNA
        formatoGrilla(1, 1) = "CHECK"
        formatoGrilla(1, 2) = "FOLIO"
        formatoGrilla(1, 3) = "N° CUOTA"
        formatoGrilla(1, 4) = "VENCIMIENO"
        formatoGrilla(1, 5) = "MONTO"
        formatoGrilla(1, 6) = "ABONO"
        
        Rem LARGO DE LOS DATOS
        formatoGrilla(2, 1) = ""
        formatoGrilla(2, 2) = "10"
        formatoGrilla(2, 3) = "9"
        formatoGrilla(2, 4) = "10"
        formatoGrilla(2, 5) = "9"
        formatoGrilla(2, 6) = "9"
        
        Rem TIPO DE DATOS
        formatoGrilla(3, 1) = "C"
        formatoGrilla(3, 2) = "N"
        formatoGrilla(3, 3) = "N"
        formatoGrilla(3, 4) = "N"
        formatoGrilla(3, 5) = "N"
        formatoGrilla(3, 6) = "N"
        
        Rem FORMATO GRILLA
        formatoGrilla(4, 1) = ""
        formatoGrilla(4, 2) = "0000000000"
        formatoGrilla(4, 3) = "000"
        formatoGrilla(4, 4) = ""
        formatoGrilla(4, 5) = "$ ###,###,##0"
        formatoGrilla(4, 6) = "$ ###,###,##0"
        
        Rem LOCCKED
        formatoGrilla(5, 1) = "FALSE"
        formatoGrilla(5, 2) = "TRUE"
        formatoGrilla(5, 3) = "TRUE"
        formatoGrilla(5, 4) = "TRUE"
        formatoGrilla(5, 5) = "TRUE"
        formatoGrilla(5, 6) = "TRUE"
        
        Rem VALOR MINIMO
        formatoGrilla(6, 1) = ""
        formatoGrilla(6, 2) = ""
        formatoGrilla(6, 3) = ""
        formatoGrilla(6, 4) = ""
        formatoGrilla(6, 5) = ""
        formatoGrilla(6, 6) = ""
        
        Rem VALOR MAXIMO
        formatoGrilla(7, 1) = ""
        formatoGrilla(7, 2) = ""
        formatoGrilla(7, 3) = ""
        formatoGrilla(7, 4) = ""
        formatoGrilla(7, 5) = ""
        formatoGrilla(7, 6) = ""
        
        Rem ANCHO
        formatoGrilla(8, 1) = "7"
        formatoGrilla(8, 2) = "12"
        formatoGrilla(8, 3) = "10"
        formatoGrilla(8, 4) = "12"
        formatoGrilla(8, 5) = "12"
        formatoGrilla(8, 6) = "12"
            
        detalle.Cols = col
        detalle.Rows = row
        detalle.AllowUserResizing = False
        detalle.DisplayFocusRect = False
        detalle.ExtendLastCol = True
        detalle.BoldFixedCell = False
        detalle.DrawMode = cellOwnerDraw
        detalle.Appearance = Flat
        detalle.ScrollBarStyle = Flat
        detalle.FixedRowColStyle = Flat
        detalle.BackColorFixed = RGB(90, 158, 214)
        detalle.BackColorFixedSel = RGB(110, 180, 230)
        detalle.BackColorBkg = RGB(90, 158, 214)
        detalle.BackColorScrollBar = RGB(231, 235, 247)
        detalle.BackColor1 = RGB(231, 235, 247)
        detalle.BackColor2 = RGB(239, 243, 255)
        detalle.GridColor = RGB(148, 190, 231)
        
        detalle.Column(0).Width = 0
        For i = 1 To col - 1
            detalle.Cell(0, i).text = formatoGrilla(1, i)
            detalle.Column(i).Width = Val(formatoGrilla(8, i)) * (detalle.Cell(0, i).Font.Size + 1.25)
            detalle.Column(i).MaxLength = Val(formatoGrilla(2, i))
            detalle.Column(i).FormatString = formatoGrilla(4, i)
            detalle.Column(i).Locked = formatoGrilla(5, i)
            If formatoGrilla(3, i) = "N" Then
                detalle.Column(i).Alignment = cellRightCenter
            End If
            If formatoGrilla(3, i) = "S" Then
                detalle.Column(i).Alignment = cellLeftCenter
            End If
            If formatoGrilla(3, i) = "C" Then
                detalle.Column(i).Alignment = cellCenterCenter
                detalle.Column(i).CellType = cellCheckBox
            End If
        Next i
        detalle.Range(0, 0, 0, detalle.Cols - 1).Alignment = cellCenterCenter
        detalle.Enabled = True
    End Sub
'****************************************************************************
'Formato de la Grilla Detalle
'****************************************************************************

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
        grabado = False
        Call Centrar(Me)
        Call CargaGrillaDetalle(1, 7)
    End Sub
    
    Private Sub Form_Unload(Cancel As Integer)
        Principal.barraEstado.Panels(1).text = UCase(Principal.Caption)
        Call limpiaBarra(2)
    End Sub

'=============================================================================
'PASA LOS DATOS DE LOS CONTROLES A LA ESTRUCTURA
'=============================================================================
    Private Sub ctrltostruct()
        c.cabeza.local = empresaactiva
        c.cabeza.folio = folio
        c.cabeza.rut = dato1.text & lblDV.Caption
        c.cabeza.tipodocumento = dato2.text
        c.cabeza.numerodocumento = dato3.text
        c.cabeza.fechaCompra = lblAño.Caption & "-" & lblMes.Caption & "-" & lblDia.Caption
        c.cabeza.total = Format(lblMonto.Caption, "########0")
        c.cabeza.abono = Format(lblAbono.Caption, "########0")
        c.cabeza.cuotas = dato7.text
        
        c.detalle.local = c.cabeza.local
        c.detalle.folio = c.cabeza.folio
        
        'Call grabarBanco(b, modifica)
        'Call retorno
    End Sub
'=============================================================================
'PASA LOS DATOS DE LOS CONTROLES A LA ESTRUCTURA
'=============================================================================

'=============================================================================
'PASA LOS DATOS DE LA ESTRUCTURA A LOS CONTROLES
'=============================================================================
    Private Sub structtoctrl()
        Dim i As Integer
        dato1.text = c.cabeza.rut
        lblDV.Caption = rut(dato1.text)
        lblNombre.Caption = leerNombreCliente(c.cabeza.rut)
        For i = 0 To cmbTipo.ListCount - 1
            If cmbTipo.List(i) = c.cabeza.tipodocumento Then
                cmbTipo.ListIndex = i
                Exit For
            End If
        Next i
        cmbTipo_KeyPress (13)
        dato3.text = c.cabeza.numerodocumento
        dato7.text = c.cabeza.cuotas
        lblMonto.Caption = Format(c.cabeza.total, "$ ###,###,##0")
        lblAbono.Caption = Format(c.cabeza.abono, "$ ###,###,##0")
        lblSaldo.Caption = Format(c.cabeza.total - c.cabeza.abono, "$ ###,###,##0")
        lblDia.Caption = Format(c.cabeza.fechaCompra, "dd")
        lblMes.Caption = Format(c.cabeza.fechaCompra, "mm")
        lblAño.Caption = Format(c.cabeza.fechaCompra, "yyyy")
        dato4.text = Format(detalle.Cell(1, 4).text, "dd")
        dato5.text = Format(detalle.Cell(1, 4).text, "mm")
        dato6.text = Format(detalle.Cell(1, 4).text, "yyyy")
        Call DeshabilitarCajas(Me)
    End Sub
'=============================================================================
'PASA LOS DATOS DE LA ESTRUCTURA A LOS CONTROLES
'=============================================================================

    Private Sub calculaCuotas()
        Dim i As Integer
        Dim resto As Integer
        Dim fecha As String
        Dim vencimiento As Date
        Dim cadena As String
        Dim cadini As String
        Dim cadfin As String
        Dim monto As Double
        Dim cuotas As Double
        Dim cuota As Double
        detalle.Rows = 1
        detalle.AutoRedraw = False
        fecha = dato4.text & "-" & lblMes.Caption & "-" & lblAño.Caption
        monto = CDbl(lblSaldo.Caption)
        cuotas = CDbl(dato7.text)
        cadini = "1" & vbTab
        folio = leer_Ultimo_Folio("folio", "sv_credito_cabeza", 10, ventasRubro, "1")
        cadini = cadini & folio & vbTab
        resto = 0
        While (monto - resto) Mod cuotas <> 0
            resto = resto + 1
        Wend
        cuota = (monto - resto) / cuotas
        For i = 1 To cuotas
            vencimiento = DateAdd("m", i, fecha)
            If i = cuotas Then
                cuota = cuota + resto
            End If
            cadena = cadini & i & vbTab & vencimiento & vbTab & cuota & vbTab & "0"
            detalle.AddItem cadena
        Next i
        detalle.AutoRedraw = True
        detalle.Refresh
    End Sub

'=============================================================================
'OPCIONES
'=============================================================================
'    Private Sub opciones_FSCommand(ByVal command As String, ByVal args As String)
'        Select Case command
'            Case "modifica"
'                Call modificar
'            Case "elimina"
'                Call eliminar
'            Case "imprime"
'            Case "movimientos"
'            Case "historico"
'            Case "retorno"
'                Call retorno
'            Case "anterior"
'                Call anterior
'            Case "siguiente"
'                Call siguiente
'        End Select
'    End Sub
'
'    Private Sub modificar()
'        modifica = True
'        Call HabilitarCajas(Me, modifica)
'        dato1.Enabled = False
'        dato2.SetFocus
'    End Sub
'
'    Private Sub eliminar()
'        Call eliminarBanco(b)
'        Call retorno
'        Call HabilitarCajas(Me, modifica)
'        dato1.SetFocus
'    End Sub
'
'    Private Sub retorno()
'        Call LimpiarCajas(MBancos)
'        Call LimpiarLabels(MBancos)
'        modifica = False
'        Call DeshabilitarCajas(Me)
'        dato1.SetFocus
'    End Sub
'
'    Private Sub anterior()
'        If leerBanco(b, dato1.text, "<") = True Then
'            structtoctrl
'        End If
'    End Sub
'
'    Private Sub siguiente()
'        If leerBanco(b, dato1.text, ">") = True Then
'            structtoctrl
'        End If
'    End Sub
'=============================================================================
'OPCIONES
'=============================================================================

    Private Sub frmImprimir_BarMouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
        frmImprimir.ColorBarraAbajo = &HFFC0C0
        frmImprimir.ColorBarraArriba = &H800000
        frmImprimir.CaptionEstilo3D = Raised
    End Sub

    Private Sub frmImprimir_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
        frmImprimir.ColorBarraAbajo = &H800000
        frmImprimir.ColorBarraArriba = &HFFC0C0
        frmImprimir.CaptionEstilo3D = Inserted
        If grabado = False Then
            Call grabarCuponera(c, modifica, detalle)
        End If
        Call imprimir
    End Sub

    Private Sub frmTodo_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
        frmTodo.ColorBarraAbajo = &HFFFFC0
        frmTodo.ColorBarraArriba = &H808000
        frmTodo.CaptionEstilo3D = Raised
    End Sub

    Private Sub frmTodo_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
        frmTodo.ColorBarraAbajo = &H808000
        frmTodo.ColorBarraArriba = &HFFFFC0
        frmTodo.CaptionEstilo3D = Inserted
        Call Marca(1)
    End Sub
    
    Private Sub frmNada_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
        frmNada.ColorBarraAbajo = &HFFFFC0
        frmNada.ColorBarraArriba = &H808000
        frmNada.CaptionEstilo3D = Raised
    End Sub

    Private Sub frmNada_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
        frmNada.ColorBarraAbajo = &H808000
        frmNada.ColorBarraArriba = &HFFFFC0
        frmNada.CaptionEstilo3D = Inserted
        Call Marca(0)
    End Sub

    Private Sub Marca(ByVal check As Integer)
        Dim i As Long
        For i = 1 To detalle.Rows - 1
            detalle.Cell(i, 1).text = check
        Next i
    End Sub

    Private Sub imprimir()
        Dim i As Long
        Dim cadena As String
        Dim nombreEmpresa As String
        Dim giroEmpresa As String
        Dim direccionEmpresa As String
        Dim fonoEmpresa As String
        Dim rutEmpresa As String
        Dim rutCliente As String
        Dim nombreCliente As String
        
        For i = 1 To impresion.PageSetup.PaperSizes.Count
            If UCase(impresion.PageSetup.PaperSizes.Item(i).PaperName) = "CARTA" Or UCase(impresion.PageSetup.PaperSizes.Item(i).PaperName) = "LETTER" Then
                impresion.PageSetup.PaperSize = impresion.PageSetup.PaperSizes.Item(i).Kind
                Exit For
            End If
        Next i
        impresion.Rows = 1
        impresion.AutoRedraw = False
        impresion.PageSetup.HeaderMargin = 1
        impresion.PageSetup.TopMargin = 1
        impresion.PageSetup.LeftMargin = 2.5
        impresion.PageSetup.RightMargin = 1
        impresion.PageSetup.BottomMargin = 0.5
        
        impresion.Cols = 8
        impresion.Column(0).Width = 0
        impresion.Column(1).Width = 50
        impresion.Column(2).Width = 50
        impresion.Column(3).Width = 80
        impresion.Column(4).Width = 100
        impresion.Column(5).Width = 100
        impresion.Column(6).Width = 150
        impresion.Column(7).Width = 100
        
        rutCliente = dato1.text & "-" & lblDV.Caption
        nombreCliente = lblNombre.Caption
        nombreEmpresa = leerNombreEmpresa(empresaactiva)
        giroEmpresa = leerGiroEmpresa(empresaactiva)
        direccionEmpresa = leerDireccionEmpresa(empresaactiva)
        fonoEmpresa = leerFonoEmpresa(empresaactiva)
        rutEmpresa = leerRutEmpresa(empresaactiva)
        folio = detalle.Cell(1, 2).text
        
        impresion.AddItem nombreEmpresa, True
        impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, 7).Merge
        impresion.RowHeight(impresion.Rows - 1) = 10
        impresion.AddItem rutEmpresa, True
        impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, 7).Merge
        impresion.RowHeight(impresion.Rows - 1) = 10
        impresion.AddItem direccionEmpresa, True
        impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, 7).Merge
        impresion.RowHeight(impresion.Rows - 1) = 10
        impresion.AddItem fonoEmpresa, True
        impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, 7).Merge
        impresion.RowHeight(impresion.Rows - 1) = 10
        impresion.Range(impresion.Rows - 4, 1, impresion.Rows - 1, 7).FontSize = 7
        impresion.Range(impresion.Rows - 4, 1, impresion.Rows - 1, 7).FontBold = True
        
        impresion.AddItem nombreCliente, True
        impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, 7).Merge
        impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, 7).Alignment = cellCenterCenter
        impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, 7).FontBold = True
        impresion.AddItem rutCliente, True
        impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, 7).Merge
        impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, 7).Alignment = cellCenterCenter
        impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, 7).FontBold = True
        impresion.AddItem "", True
        impresion.AddItem "", True
        impresion.AddItem "", True
        impresion.AddItem "", True
        impresion.AddItem "", True
        impresion.AddItem "", True
        
        
        impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, 7).Borders(cellEdgeTop) = cellDot
        impresion.AddItem "", True
        impresion.RowHeight(impresion.Rows - 1) = 10
        impresion.AddItem "", True
        impresion.RowHeight(impresion.Rows - 1) = 10
        
        For i = 1 To detalle.Rows - 1
            If detalle.Cell(i, 1).text = 1 Then
                cadena = nombreEmpresa & vbTab & vbTab & folio & "   " & vbTab
                cadena = cadena & "   " & nombreEmpresa & vbTab & vbTab & giroEmpresa & vbTab & folio
                impresion.AddItem cadena
                impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, 2).Merge
                impresion.Range(impresion.Rows - 1, 4, impresion.Rows - 1, 5).Merge
                impresion.RowHeight(impresion.Rows - 1) = 10
                impresion.Cell(impresion.Rows - 1, 3).Alignment = cellRightCenter
                impresion.Cell(impresion.Rows - 1, 7).Alignment = cellRightCenter
                
                cadena = direccionEmpresa & vbTab & vbTab & vbTab
                cadena = cadena & "   " & direccionEmpresa & vbTab & vbTab & rutEmpresa & vbTab & vbTab
                impresion.AddItem cadena
                impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, 3).Merge
                impresion.Range(impresion.Rows - 1, 4, impresion.Rows - 1, 5).Merge
                impresion.RowHeight(impresion.Rows - 1) = 10
                
                cadena = "   " & fonoEmpresa
                impresion.AddItem cadena
                impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, 2).Merge
                impresion.Range(impresion.Rows - 1, 5, impresion.Rows - 1, 6).Merge
                impresion.RowHeight(impresion.Rows - 1) = 10
                
                impresion.Range(impresion.Rows - 3, 1, impresion.Rows - 1, 7).FontSize = 6
                impresion.Range(impresion.Rows - 3, 1, impresion.Rows - 1, 7).FontBold = True
                
                impresion.AddItem "", True
                impresion.RowHeight(impresion.Rows - 1) = 10
                
                cadena = "CLIENTE" & vbTab & vbTab & rutCliente & "   " & vbTab
                cadena = cadena & "   CLIENTE" & vbTab & nombreCliente
                impresion.AddItem cadena
                impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, 2).Merge
                impresion.Range(impresion.Rows - 1, 5, impresion.Rows - 1, 7).Merge
                impresion.Cell(impresion.Rows - 1, 3).Alignment = cellRightCenter
                impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, 2).FontBold = True
                impresion.Range(impresion.Rows - 1, 4, impresion.Rows - 1, 4).FontBold = True
                
                cadena = "CUOTA" & vbTab & vbTab & Val(detalle.Cell(i, 3).text) & " de " & dato7.text & vbTab
                cadena = cadena & "   RUT" & vbTab & rutCliente & "   " & vbTab & "CUOTA" & vbTab & Val(detalle.Cell(i, 3).text) & " de " & dato7.text & vbTab
                impresion.AddItem cadena
                impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, 2).Merge
                impresion.Cell(impresion.Rows - 1, 3).Alignment = cellCenterCenter
                impresion.Cell(impresion.Rows - 1, 5).Alignment = cellRightCenter
                impresion.Cell(impresion.Rows - 1, 6).Alignment = cellCenterCenter
                impresion.Cell(impresion.Rows - 1, 7).Alignment = cellCenterCenter
                impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, 2).FontBold = True
                impresion.Range(impresion.Rows - 1, 4, impresion.Rows - 1, 4).FontBold = True
                impresion.Range(impresion.Rows - 1, 6, impresion.Rows - 1, 6).FontBold = True
                
                cadena = "VENCIMIENTO" & vbTab & vbTab & detalle.Cell(i, 4).text & vbTab
                cadena = cadena & "   VENCIMIENTO" & vbTab & vbTab & detalle.Cell(i, 4).text & vbTab & vbTab
                impresion.AddItem cadena
                impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, 2).Merge
                impresion.Range(impresion.Rows - 1, 5, impresion.Rows - 1, 6).Merge
                impresion.Cell(impresion.Rows - 1, 3).Alignment = cellCenterCenter
                impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, 2).FontBold = True
                impresion.Range(impresion.Rows - 1, 4, impresion.Rows - 1, 4).FontBold = True
                
                cadena = "TOTAL A PAGAR" & vbTab & vbTab & Format(detalle.Cell(i, 5).text, "$ ###,###,##0") & "   " & vbTab
                cadena = cadena & vbTab & vbTab & "TOTAL A PAGAR" & vbTab & Format(detalle.Cell(i, 5).text, "$ ###,###,##0")
                impresion.AddItem cadena
                impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, 2).Merge
                impresion.Cell(impresion.Rows - 1, 3).Alignment = cellRightCenter
                impresion.Cell(impresion.Rows - 1, 6).Alignment = cellCenterCenter
                impresion.Cell(impresion.Rows - 1, 7).Alignment = cellRightCenter
                impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, 2).FontBold = True
                impresion.Range(impresion.Rows - 1, 6, impresion.Rows - 1, 6).FontBold = True
                
                cadena = "INTERES MORA" & vbTab & vbTab & vbTab
                cadena = cadena & vbTab & vbTab & "INTERES MORA"
                impresion.AddItem cadena
                impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, 2).Merge
                impresion.Cell(impresion.Rows - 1, 3).Alignment = cellRightCenter
                impresion.Cell(impresion.Rows - 1, 6).Alignment = cellCenterCenter
                impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, 2).FontBold = True
                impresion.Range(impresion.Rows - 1, 6, impresion.Rows - 1, 6).FontBold = True
                
                cadena = "TOTAL PAGADO" & vbTab & vbTab & vbTab
                cadena = cadena & vbTab & vbTab & "TOTAL PAGADO"
                impresion.AddItem cadena
                impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, 2).Merge
                impresion.Cell(impresion.Rows - 1, 3).Alignment = cellRightCenter
                impresion.Cell(impresion.Rows - 1, 6).Alignment = cellCenterCenter
                impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, 2).FontBold = True
                impresion.Range(impresion.Rows - 1, 6, impresion.Rows - 1, 6).FontBold = True
                
                impresion.Range(impresion.Rows - 6, 1, impresion.Rows - 6, 7).Borders(cellEdgeTop) = cellThin
                impresion.Range(impresion.Rows - 5, 1, impresion.Rows - 5, 7).Borders(cellEdgeTop) = cellThin
                impresion.Range(impresion.Rows - 4, 1, impresion.Rows - 4, 7).Borders(cellEdgeTop) = cellThin
                impresion.Range(impresion.Rows - 3, 1, impresion.Rows - 3, 7).Borders(cellEdgeTop) = cellThin
                
                impresion.Range(impresion.Rows - 2, 1, impresion.Rows - 2, 3).Borders(cellEdgeTop) = cellThin
                impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, 3).Borders(cellEdgeTop) = cellThin
                impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, 3).Borders(cellEdgeBottom) = cellThin
                
                impresion.Range(impresion.Rows - 5, 6, impresion.Rows - 5, 6).Borders(cellEdgeRight) = cellThin
                impresion.Range(impresion.Rows - 5, 6, impresion.Rows - 5, 6).Borders(cellEdgeLeft) = cellThin
                
                impresion.Range(impresion.Rows - 2, 6, impresion.Rows - 2, 7).Borders(cellEdgeTop) = cellThin
                impresion.Range(impresion.Rows - 1, 6, impresion.Rows - 1, 7).Borders(cellEdgeTop) = cellThin
                impresion.Range(impresion.Rows - 1, 6, impresion.Rows - 1, 7).Borders(cellEdgeBottom) = cellThin
                
                impresion.Range(impresion.Rows - 6, 1, impresion.Rows - 1, 1).Borders(cellEdgeLeft) = cellThin
                impresion.Range(impresion.Rows - 6, 1, impresion.Rows - 1, 1).Borders(cellEdgeRight) = cellThin
                
                impresion.Range(impresion.Rows - 6, 5, impresion.Rows - 4, 5).Borders(cellEdgeLeft) = cellThin
                impresion.Range(impresion.Rows - 3, 6, impresion.Rows - 1, 6).Borders(cellEdgeLeft) = cellThin
                impresion.Range(impresion.Rows - 3, 6, impresion.Rows - 1, 6).Borders(cellEdgeRight) = cellThin
                impresion.Range(impresion.Rows - 6, 7, impresion.Rows - 1, 7).Borders(cellEdgeRight) = cellThin
                
                
                If (i + 1) Mod 5 <> 0 Then
                    impresion.AddItem "", True
                    impresion.RowHeight(impresion.Rows - 1) = 10
                    impresion.AddItem "", True
                    impresion.RowHeight(impresion.Rows - 1) = 10
                    impresion.AddItem "", True
                    impresion.RowHeight(impresion.Rows - 1) = 10
                    impresion.AddItem "", True
                    impresion.RowHeight(impresion.Rows - 1) = 10
                    
                    impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, 7).Borders(cellEdgeTop) = cellDot
                    impresion.Range(impresion.Rows - 16, 3, impresion.Rows - 3, 3).Borders(cellEdgeRight) = cellDot
                    impresion.AddItem "", True
                    impresion.RowHeight(impresion.Rows - 1) = 10
                    impresion.AddItem "", True
                    impresion.RowHeight(impresion.Rows - 1) = 10
                Else
                    impresion.AddItem "", True
                    'impresion.RowHeight(impresion.Rows - 1) = 13
                    impresion.Range(impresion.Rows - 13, 3, impresion.Rows - 1, 3).Borders(cellEdgeRight) = cellDot
                End If
                
            End If
        Next i
        
        
        
        
        
        impresion.AutoRedraw = True
        impresion.Refresh
        impresion.PrintPreview
    End Sub





