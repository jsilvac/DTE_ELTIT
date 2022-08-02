VERSION 5.00
Object = "{4777436C-EB8C-4596-98A8-EBCF98683969}#1.0#0"; "FlexCell.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{ADBBAED6-B16F-40DC-80DF-B44910CBA76C}#1.0#0"; "Frame-Xp.ocx"
Begin VB.Form AnulaDocumentos 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Comisiones por Vendedor"
   ClientHeight    =   9000
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11175
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9000
   ScaleWidth      =   11175
   Begin VB.TextBox dato8 
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
      Height          =   315
      Left            =   300
      MaxLength       =   10
      TabIndex        =   7
      Tag             =   "proveedor"
      Top             =   8040
      Width           =   1575
   End
   Begin XPFrame.FrameXp frmImprimir 
      Height          =   375
      Left            =   7080
      TabIndex        =   8
      Top             =   8520
      Width           =   3375
      _ExtentX        =   5953
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
   Begin XPFrame.FrameXp frmLiquidar 
      Height          =   375
      Left            =   7080
      TabIndex        =   9
      Top             =   8040
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   661
      BackColor       =   49344
      Caption         =   "A   N   U   L   A   R"
      CaptionEstilo3D =   1
      BackColor       =   49344
      ColorBarraArriba=   12632319
      ColorBarraAbajo =   128
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
   Begin XPFrame.FrameXp frmSel 
      Height          =   375
      Left            =   5100
      TabIndex        =   11
      Top             =   8040
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   661
      BackColor       =   49344
      Caption         =   "Sel Todo"
      CaptionEstilo3D =   1
      BackColor       =   49344
      ColorBarraArriba=   12632319
      ColorBarraAbajo =   128
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
   Begin XPFrame.FrameXp FrameXp1 
      Height          =   1275
      Left            =   1320
      TabIndex        =   12
      Top             =   60
      Width           =   8535
      _ExtentX        =   15055
      _ExtentY        =   2249
      BackColor       =   16744576
      Caption         =   "Ingreso de Información"
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
         Height          =   315
         Left            =   6600
         MaxLength       =   2
         TabIndex        =   5
         Tag             =   "proveedor"
         Top             =   840
         Width           =   435
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
         Height          =   315
         Left            =   6120
         MaxLength       =   2
         TabIndex        =   4
         Tag             =   "proveedor"
         Top             =   840
         Width           =   435
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
         Height          =   315
         Left            =   7080
         MaxLength       =   4
         TabIndex        =   6
         Tag             =   "proveedor"
         Top             =   840
         Width           =   795
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
         Height          =   315
         Left            =   3120
         MaxLength       =   4
         TabIndex        =   3
         Tag             =   "proveedor"
         Top             =   840
         Width           =   795
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
         Height          =   315
         Left            =   2640
         MaxLength       =   2
         TabIndex        =   2
         Tag             =   "proveedor"
         Top             =   840
         Width           =   435
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
         Height          =   315
         Left            =   2160
         MaxLength       =   2
         TabIndex        =   1
         Tag             =   "proveedor"
         Top             =   840
         Width           =   435
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
         Height          =   315
         Left            =   1560
         MaxLength       =   2
         TabIndex        =   0
         Tag             =   "proveedor"
         Top             =   420
         Width           =   495
      End
      Begin VB.Label lbl2 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Desde"
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
         Height          =   315
         Left            =   720
         TabIndex        =   16
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label lbl3 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Hasta"
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
         Height          =   315
         Left            =   4680
         TabIndex        =   15
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Local"
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
         Height          =   315
         Left            =   120
         TabIndex        =   14
         Top             =   420
         Width           =   1335
      End
      Begin VB.Label lblLocal 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FFFF&
         Height          =   315
         Left            =   2160
         TabIndex        =   13
         Top             =   420
         Width           =   6255
      End
   End
   Begin XPFrame.FrameXp FrameXp2 
      Height          =   6375
      Left            =   60
      TabIndex        =   17
      Top             =   1440
      Width           =   11055
      _ExtentX        =   19500
      _ExtentY        =   11245
      BackColor       =   16744576
      Caption         =   "Informe"
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
      Begin FlexCell.Grid impresion 
         Height          =   5895
         Left            =   60
         TabIndex        =   18
         Top             =   420
         Width           =   10935
         _ExtentX        =   19288
         _ExtentY        =   10398
         Cols            =   5
         DefaultFontSize =   8.25
         Rows            =   30
         SelectionMode   =   1
         DateFormat      =   2
      End
   End
   Begin MSAdodcLib.Adodc data 
      Height          =   330
      Left            =   120
      Top             =   7620
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
   Begin VB.Label lblEncontrado 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   315
      Left            =   1920
      TabIndex        =   10
      Top             =   8040
      Width           =   2235
   End
End
Attribute VB_Name = "AnulaDocumentos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Private fecha1 As String
    Private fecha2 As String
        
'============================================================
'MANEJO DE LOS CONTOLES
'============================================================
    '========================================================
    'GotFocus
    '========================================================
    Private Sub dato1_GotFocus()
        Call VerificarCajas(Me, dato1)
        Call selecciona(dato1)
        Principal.barraEstado.Panels(2).text = "F2: Ayuda Locales"
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
    
    Private Sub dato8_GotFocus()
        Call selecciona(dato8)
    End Sub
    '========================================================
    'GotFocus
    '========================================================
    
    '========================================================
    'KeyDown
    '========================================================
    Private Sub dato1_KeyDown(KeyCode As Integer, Shift As Integer)
        Select Case KeyCode
            Case vbKeyF2
                Call ayudaVendedores(dato1)
            Case Else
                Call flechas(KeyCode, dato1)
        End Select
    End Sub
    
    Private Sub dato2_KeyDown(KeyCode As Integer, Shift As Integer)
        Call flechas(KeyCode, dato1)
    End Sub

    Private Sub dato3_KeyDown(KeyCode As Integer, Shift As Integer)
        Call flechas(KeyCode, dato2)
    End Sub
    
    Private Sub dato4_KeyDown(KeyCode As Integer, Shift As Integer)
        Call flechas(KeyCode, dato3)
    End Sub
    
    Private Sub dato5_KeyDown(KeyCode As Integer, Shift As Integer)
        Call flechas(KeyCode, dato4)
    End Sub
    
    Private Sub dato6_KeyDown(KeyCode As Integer, Shift As Integer)
        Call flechas(KeyCode, dato5)
    End Sub
    
    Private Sub dato7_KeyDown(KeyCode As Integer, Shift As Integer)
        Call flechas(KeyCode, dato6)
    End Sub
    
    Private Sub dato8_KeyDown(KeyCode As Integer, Shift As Integer)
        Call flechas(KeyCode, dato7)
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
            lblLocal.Caption = leerNombreEmpresa(dato1.text)
            If lblLocal.Caption <> "" Then
                SendKeys "{Tab}"
            Else
                Call selecciona(dato1)
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
            frmSel.Caption = "Sel Todos"
            lblEncontrado.Caption = ""
            dato8.text = ""
            dato7.text = ceros(dato7)
            If dato7.text = "0000" Then
                dato7.text = Format(fechasistema, "yyyy")
            End If
            fecha1 = dato4.text & "-" & dato3.text & "-" & dato2.text
            fecha2 = dato7.text & "-" & dato6.text & "-" & dato5.text
            Call listadoDocumentos(data, impresion, dato1.text, fecha1, fecha2)
        End If
    End Sub
    
    Private Sub dato8_KeyPress(KeyAscii As Integer)
        Dim i As Long
        Dim cad As String
        KeyAscii = esNumero(KeyAscii)
        If KeyAscii = 13 Then
            dato8.text = ceros(dato8)
            lblEncontrado.Caption = ""
            For i = 1 To impresion.Rows - 1
                cad = Right(impresion.Cell(i, 1).text, 10)
                If cad = dato8.text Then
                    impresion.Cell(i, impresion.Cols - 1).text = 1
                    lblEncontrado.Caption = "ENCONTRADO"
                    Call selecciona(dato8)
                    Exit For
                End If
            Next i
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
    '========================================================
    'LostFocus
    '========================================================
'============================================================
'MANEJO DE LOS CONTOLES
'============================================================
    
    Private Sub Form_Activate()
        Principal.barraEstado.Panels(1).text = UCase(Me.Caption)
    End Sub

    Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
        Select Case KeyCode
            Case 27
                Unload Me
            Case 38
                If Screen.ActiveForm.ActiveControl.Name = "dato1" Then
                    Unload Me
                End If
        End Select
    End Sub
    
    Private Sub Form_Load()
        Call Centrar(Me)
        Call CargaGrillaInforme(1, 10)
    End Sub

'****************************************************************************
'Formato de la Grilla Listado de Ventas
'****************************************************************************
    Private Sub CargaGrillaInforme(ByVal row As Integer, ByVal col As Integer)
        Dim formatoGrilla(10, 12) As String
        Dim i As Integer
        Rem DATOS DE LA COLUMNA
        formatoGrilla(1, 1) = "DOCUMENTO"
        formatoGrilla(1, 2) = "FECHA"
        formatoGrilla(1, 3) = "VENDEDOR"
        formatoGrilla(1, 4) = "CLIENTE"
        formatoGrilla(1, 5) = "T.PAGO"
        formatoGrilla(1, 6) = "MONTO"
        formatoGrilla(1, 7) = "DESCUENTO"
        formatoGrilla(1, 8) = "A PAGO"
        formatoGrilla(1, 9) = "NULA"
        
        Rem LARGO DE LOS DATOS
        formatoGrilla(2, 1) = "13"
        formatoGrilla(2, 2) = "10"
        formatoGrilla(2, 3) = "4"
        formatoGrilla(2, 4) = "70"
        formatoGrilla(2, 5) = "3"
        formatoGrilla(2, 6) = "9"
        formatoGrilla(2, 7) = "9"
        formatoGrilla(2, 8) = "9"
        formatoGrilla(2, 9) = "1"
        
        Rem TIPO DE DATOS C=CEROS,S=STRING,N=NUMERICO,D=DATE,
        formatoGrilla(3, 1) = "C"
        formatoGrilla(3, 2) = "C"
        formatoGrilla(3, 3) = "C"
        formatoGrilla(3, 4) = "S"
        formatoGrilla(3, 5) = "C"
        formatoGrilla(3, 6) = "N"
        formatoGrilla(3, 7) = "N"
        formatoGrilla(3, 8) = "N"
        formatoGrilla(3, 9) = "C"
        
        Rem FORMATO GRILLA
        formatoGrilla(4, 1) = ""
        formatoGrilla(4, 2) = ""
        formatoGrilla(4, 3) = ""
        formatoGrilla(4, 4) = ""
        formatoGrilla(4, 5) = ""
        formatoGrilla(4, 6) = "$ ###,###,##0"
        formatoGrilla(4, 7) = "$ ###,###,##0"
        formatoGrilla(4, 8) = "$ ###,###,##0"
        formatoGrilla(4, 9) = "0"
        
        Rem LOCCKED
        formatoGrilla(5, 1) = "FALSE"
        formatoGrilla(5, 2) = "FALSE"
        formatoGrilla(5, 3) = "FALSE"
        formatoGrilla(5, 4) = "FALSE"
        formatoGrilla(5, 5) = "FALSE"
        formatoGrilla(5, 6) = "FALSE"
        formatoGrilla(5, 7) = "FALSE"
        formatoGrilla(5, 8) = "FALSE"
        formatoGrilla(5, 9) = "FALSE"
        
        Rem ANCHO
        formatoGrilla(8, 1) = "9"
        formatoGrilla(8, 2) = "7"
        formatoGrilla(8, 3) = "5"
        formatoGrilla(8, 4) = "20"
        formatoGrilla(8, 5) = "5"
        formatoGrilla(8, 6) = "8"
        formatoGrilla(8, 7) = "7"
        formatoGrilla(8, 8) = "9"
        formatoGrilla(8, 9) = "3"
            
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
        
        For i = 1 To col - 1
            impresion.Cell(0, i).text = formatoGrilla(1, i)
            impresion.Column(i).Width = Val(formatoGrilla(8, i)) * (impresion.Cell(0, i).Font.Size + 1.25)
            impresion.Column(i).MaxLength = Val(formatoGrilla(2, i))
            impresion.Column(i).FormatString = formatoGrilla(4, i)
            impresion.Column(i).Locked = formatoGrilla(5, i)
            If formatoGrilla(3, i) = "N" Then
                impresion.Column(i).Alignment = cellRightCenter
            End If
            If formatoGrilla(3, i) = "S" Then
                impresion.Column(i).Alignment = cellLeftCenter
            End If
            If formatoGrilla(3, i) = "C" Then
                impresion.Column(i).Alignment = cellCenterCenter
            End If
        Next i
        impresion.Column(impresion.Cols - 1).CellType = cellCheckBox
        impresion.Range(0, 1, 0, impresion.Cols - 1).Alignment = cellCenterCenter
        impresion.Range(0, 1, 0, impresion.Cols - 1).Borders(cellEdgeBottom) = cellThin
        
        'impresion.Range(4, 0, 4, impresion.Cols - 1).BackColor = RGB(200, 225, 250)
        'impresion.Enabled = True
                
    End Sub
'****************************************************************************
'Formato de la Grilla Listado de Ventas
'****************************************************************************
    
    Private Sub frmSel_BarMouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
        Call cambiaColor(frmSel)
        frmSel.CaptionEstilo3D = Raised
    End Sub
    
    Private Sub frmSel_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
        Dim i As Long
        Dim estado As Integer
        Call cambiaColor(frmSel)
        frmSel.CaptionEstilo3D = Inserted
        If frmSel.Caption = "Sel Nada" Then
            frmSel.Caption = "Sel Todos"
            estado = 0
        Else
            frmSel.Caption = "Sel Nada"
            estado = 1
        End If
        For i = 1 To impresion.Rows - 1
            If impresion.Cell(i, impresion.Cols - 1).CellType = 6 Then
                impresion.Cell(i, impresion.Cols - 1).text = estado
            End If
        Next i
    End Sub

    Private Sub frmLiquidar_BarMouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
        Call cambiaColor(frmLiquidar)
        frmLiquidar.CaptionEstilo3D = Raised
    End Sub
    
    Private Sub frmLiquidar_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
        Call cambiaColor(frmLiquidar)
        frmLiquidar.CaptionEstilo3D = Inserted
        Call anular
    End Sub

    Private Sub frmImprimir_BarMouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
        Call cambiaColor(frmImprimir)
        frmImprimir.CaptionEstilo3D = Raised
    End Sub
    
    Private Sub frmImprimir_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
        Call cambiaColor(frmImprimir)
        frmImprimir.CaptionEstilo3D = Inserted
        Call imprimir
    End Sub
    
    Private Sub imprimir()
        Dim i As Long
        impresion.AutoRedraw = False
        
        For i = i To impresion.Rows - 1
            If impresion.Cell(i, impresion.Cols - 1).text <> "1" Then
                impresion.RowHeight(i) = 0
            End If
        Next i
        
        impresion.AutoRedraw = True
        impresion.Refresh
        impresion.PageSetup.HeaderMargin = 2
    
        impresion.PageSetup.TopMargin = 2
        impresion.PageSetup.LeftMargin = 1
        impresion.PageSetup.RightMargin = 1
        impresion.PageSetup.BottomMargin = 2
        
        impresion.PageSetup.FooterMargin = 2
        impresion.PageSetup.BlackAndWhite = True
        impresion.PageSetup.PrintFixedRow = True
        
        Call verificaImpresora(5, impresion)
        
        impresion.AutoRedraw = True
    End Sub
    
    Private Sub leerMeses()
        Dim i As Integer
        Dim fecha As String
        For i = 1 To 12
            fecha = Format(fechasistema, "yyyy") & "-" & i & "-01"
            meses(i) = UCase(Format(fecha, "mmmm"))
            If Val(Format(fechasistema, "mm")) = i Then
                cantMeses = i
            End If
        Next i
    End Sub

    Private Sub impresion_Click()
        If impresion.Cell(impresion.ActiveCell.row, 1).text <> "" Then
            If impresion.Cell(impresion.ActiveCell.row, impresion.Cols - 1).text = "1" Then
                impresion.Cell(impresion.ActiveCell.row, impresion.Cols - 1).text = "0"
            Else
                impresion.Cell(impresion.ActiveCell.row, impresion.Cols - 1).text = "1"
            End If
        End If
    End Sub

Private Sub anular()
    Dim i As Integer
    Dim codLoc As String
    Dim tipoDoc As String
    Dim numeroDoc As String
    
    For i = 1 To impresion.Rows - 1
        If impresion.Cell(i, impresion.Cols - 1).text = "1" Then
            codLoc = dato1.text
            tipoDoc = Left(impresion.Cell(i, 1).text, 2)
            numeroDoc = Right(impresion.Cell(i, 1).text, 10)
            Call anularDocumentoSin(codLoc, tipoDoc, numeroDoc)
        End If
    Next i
    Call imprimir
End Sub

Private Sub liquidarDocumento(ByVal codLoc As String, ByVal tipoDoc As String, ByVal numDoc As String, ByVal comision As String, ByVal porcentaje As String)
    Dim cSql As rdoQuery
    Set cSql = New rdoQuery
    Set cSql.ActiveConnection = ventasRubro
    
    cSql.sql = "UPDATE sv_documento_cabeza "
    cSql.sql = cSql.sql & "SET comision = '" & comision & "', fechapagocomision = '" & Format(fecha2, "yyyy-mm-dd") & "', numeroliquidacion = '" & dato7.text & dato6.text & "', porcentajecomision = " & porcentaje & " "
    cSql.sql = cSql.sql & "WHERE local = '" & codLoc & "' AND tipo = '" & tipoDoc & "' AND numero = '" & numDoc & "' "
    cSql.Execute
    
    cSql.Close
    Set cSql = Nothing
End Sub
