VERSION 5.00
Object = "{4777436C-EB8C-4596-98A8-EBCF98683969}#1.0#0"; "Flexcell.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{ADBBAED6-B16F-40DC-80DF-B44910CBA76C}#1.0#0"; "Frame-Xp.ocx"
Begin VB.Form PreciosPromedio 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Listado de Precios Promedio"
   ClientHeight    =   8385
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   15270
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8385
   ScaleWidth      =   15270
   Begin MSAdodcLib.Adodc data 
      Height          =   330
      Left            =   120
      Top             =   7920
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
   Begin XPFrame.FrameXp FrameXp2 
      Height          =   6375
      Left            =   60
      TabIndex        =   13
      Top             =   1440
      Width           =   15135
      _ExtentX        =   26696
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
         TabIndex        =   14
         Top             =   420
         Width           =   15015
         _ExtentX        =   26485
         _ExtentY        =   10398
         Cols            =   5
         DefaultFontSize =   8.25
         Rows            =   30
         SelectionMode   =   1
      End
   End
   Begin XPFrame.FrameXp FrameXp1 
      Height          =   1275
      Left            =   3360
      TabIndex        =   10
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
         Left            =   6600
         MaxLength       =   2
         TabIndex        =   4
         Tag             =   "proveedor"
         Top             =   420
         Width           =   435
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
         Left            =   6120
         MaxLength       =   2
         TabIndex        =   3
         Tag             =   "proveedor"
         Top             =   420
         Width           =   435
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
         Height          =   315
         Left            =   7080
         MaxLength       =   4
         TabIndex        =   5
         Tag             =   "proveedor"
         Top             =   420
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
         Left            =   3120
         MaxLength       =   4
         TabIndex        =   2
         Tag             =   "proveedor"
         Top             =   420
         Width           =   795
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
         Left            =   2640
         MaxLength       =   2
         TabIndex        =   1
         Tag             =   "proveedor"
         Top             =   420
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
         Left            =   2160
         MaxLength       =   2
         TabIndex        =   0
         Tag             =   "proveedor"
         Top             =   420
         Width           =   435
      End
      Begin VB.CheckBox chk1 
         BackColor       =   &H00FF8080&
         Caption         =   "Imprime Detalle"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   6240
         TabIndex        =   9
         Top             =   840
         Width           =   2235
      End
      Begin VB.OptionButton opt3 
         BackColor       =   &H00FF8080&
         Caption         =   "Todo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Left            =   3300
         TabIndex        =   8
         Top             =   900
         Value           =   -1  'True
         Width           =   975
      End
      Begin VB.OptionButton opt2 
         BackColor       =   &H00FF8080&
         Caption         =   "Boletas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Left            =   1800
         TabIndex        =   7
         Top             =   900
         Width           =   1275
      End
      Begin VB.OptionButton opt1 
         BackColor       =   &H00FF8080&
         Caption         =   "Facturas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Left            =   120
         TabIndex        =   6
         Top             =   900
         Width           =   1455
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
         TabIndex        =   12
         Top             =   420
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
         TabIndex        =   11
         Top             =   420
         Width           =   1335
      End
   End
   Begin XPFrame.FrameXp frmImprimir 
      Height          =   375
      Left            =   11820
      TabIndex        =   15
      Top             =   7920
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
End
Attribute VB_Name = "PreciosPromedio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Private TIPO As String
    Private detalle As Boolean
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
            If dato1.text = "00" Then
                dato1.text = Format(fechasistema, "dd")
            End If
            SendKeys "{Tab}"
        End If
    End Sub

    Private Sub dato2_KeyPress(KeyAscii As Integer)
        KeyAscii = esNumero(KeyAscii)
        If KeyAscii = 13 Then
            dato2.text = ceros(dato2)
            If dato2.text = "00" Then
                dato2.text = Format(fechasistema, "mm")
            End If
            SendKeys "{Tab}"
        End If
    End Sub
        
    Private Sub dato3_KeyPress(KeyAscii As Integer)
        KeyAscii = esNumero(KeyAscii)
        If KeyAscii = 13 Then
            dato3.text = ceros(dato3)
            If dato3.text = "0000" Then
                dato3.text = Format(fechasistema, "yyyy")
            End If
            SendKeys "{Tab}"
        End If
    End Sub
    
    Private Sub dato4_KeyPress(KeyAscii As Integer)
        KeyAscii = esNumero(KeyAscii)
        If KeyAscii = 13 Then
            dato4.text = ceros(dato4)
            If dato4.text = "00" Then
                dato4.text = Format(fechasistema, "dd")
            End If
            SendKeys "{Tab}"
        End If
    End Sub
    
    Private Sub dato5_KeyPress(KeyAscii As Integer)
        KeyAscii = esNumero(KeyAscii)
        If KeyAscii = 13 Then
            dato5.text = ceros(dato5)
            If dato5.text = "00" Then
                dato5.text = Format(fechasistema, "mm")
            End If
            SendKeys "{Tab}"
        End If
    End Sub
        
    Private Sub dato6_KeyPress(KeyAscii As Integer)
        KeyAscii = esNumero(KeyAscii)
        If KeyAscii = 13 Then
            dato6.text = ceros(dato6)
            If dato6.text = "0000" Then
                dato6.text = Format(fechasistema, "yyyy")
            End If
            fecha1 = dato3.text & "-" & dato2.text & "-" & dato1.text
            fecha2 = dato6.text & "-" & dato5.text & "-" & dato4.text
            SendKeys "{Tab}"
            Call generaInformePP(data, impresion, TIPO, detalle, dato1.text, fecha1, fecha2)
        End If
    End Sub
    '========================================================
    'KeyPress
    '========================================================

    '========================================================
    'KeyUp
    '========================================================
'    Private Sub dato1_KeyUp(KeyCode As Integer, Shift As Integer)
'        If Len(dato1.text) = dato1.MaxLength Then
'            Call dato1_KeyPress(13)
'        End If
'    End Sub
'
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
    '========================================================
    'KeyUp
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
        Call CargaGrillaInforme(1, 3)
        TIPO = "(dd.tipo = 'FV' OR dd.tipo = 'NV' OR dd.tipo = 'BV' OR dd.tipo = 'ZE' OR dd.tipo = 'FE')"
        detalle = False
    End Sub

'****************************************************************************
'Formato de la Grilla Listado de Ventas
'****************************************************************************
    Private Sub CargaGrillaInforme(ByVal row As Integer, ByVal col As Integer)
        Dim formatoGrilla(10, 20) As String
        Dim i As Integer
        
        Call leerEmpresas
        
        Rem DATOS DE LA COLUMNA
        formatoGrilla(1, 1) = "CODIGO"
        formatoGrilla(1, 2) = "DESCRIPCION"
        formatoGrilla(1, 3) = "ITEMS"
        
        Rem LARGO DE LOS DATOS
        formatoGrilla(2, 1) = "30"
        formatoGrilla(2, 2) = "30"
        formatoGrilla(2, 3) = "8"
        
        Rem TIPO DE DATOS C=CEROS,S=STRING,N=NUMERICO,D=DATE,
        formatoGrilla(3, 1) = "N"
        formatoGrilla(3, 2) = "S"
        formatoGrilla(3, 3) = "S"
        
        Rem FORMATO GRILLA
        formatoGrilla(4, 1) = ""
        formatoGrilla(4, 2) = ""
        formatoGrilla(4, 2) = ""
        
        Rem LOCCKED
        formatoGrilla(5, 1) = "FALSE"
        formatoGrilla(5, 2) = "FALSE"
        formatoGrilla(5, 3) = "FALSE"
        
        Rem VALOR MINIMO
        formatoGrilla(6, 1) = ""
        formatoGrilla(6, 2) = ""
        formatoGrilla(6, 3) = ""
        
        Rem VALOR MAXIMO
        formatoGrilla(7, 1) = ""
        formatoGrilla(7, 2) = ""
        formatoGrilla(7, 3) = ""
        
        Rem ANCHO
        formatoGrilla(8, 1) = "9"
        formatoGrilla(8, 2) = "10"
        formatoGrilla(8, 3) = "5"
                
        i = 4
        While Not data.Recordset.EOF
            If data.Recordset.Fields("codigo") = "00" Then
                formatoGrilla(1, i) = "CASA MATRIZ"
                cabezaLocales(0, i) = "CASA MATRIZ"
            Else
                formatoGrilla(1, i) = Replace(data.Recordset.Fields("nombre"), "MOLINO ", "")
                cabezaLocales(0, i) = Replace(data.Recordset.Fields("nombre"), "MOLINO ", "")
            End If
            cantLocales = i
            
            formatoGrilla(2, i) = "30"
            
            formatoGrilla(3, i) = "N"
            
            formatoGrilla(4, i) = ""
            
            formatoGrilla(5, i) = "FALSE"
            
            formatoGrilla(6, i) = ""
            
            formatoGrilla(7, i) = ""
            
            formatoGrilla(8, i) = "8"
            
            data.Recordset.MoveNext
            i = i + 1
        Wend
        Rem DATOS DE LA COLUMNA
        formatoGrilla(1, i) = "PROMEDIO"
        
        Rem LARGO DE LOS DATOS
        formatoGrilla(2, i) = "9"
        
        Rem TIPO DE DATOS C=CEROS,S=STRING,N=NUMERICO,D=DATE,
        formatoGrilla(3, i) = "N"
        
        Rem FORMATO GRILLA
        formatoGrilla(4, i) = ""
        
        Rem LOCCKED
        formatoGrilla(5, i) = "FALSE"
        
        Rem VALOR MINIMO
        formatoGrilla(6, i) = ""
        
        Rem VALOR MAXIMO
        formatoGrilla(7, i) = ""
        
        Rem ANCHO
        formatoGrilla(8, i) = "9"
        
        impresion.Cols = i + 1
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
        
        For i = 1 To impresion.Cols - 1
            impresion.Cell(0, i).text = formatoGrilla(1, i)
            impresion.Column(i).Width = Val(formatoGrilla(8, i)) * (impresion.Cell(0, i).Font.Size + 1.25)
            impresion.Column(i).MaxLength = Val(formatoGrilla(2, i))
            impresion.Column(i).FormatString = formatoGrilla(4, i)
            impresion.Column(i).Locked = formatoGrilla(5, i)
            If formatoGrilla(3, i) = "N" Then
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

    Private Sub frmImprimir_BarMouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
        Call cambiaColor(frmImprimir)
        frmImprimir.CaptionEstilo3D = Raised
    End Sub
    
    Private Sub frmImprimir_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
        Call cambiaColor(frmImprimir)
        frmImprimir.CaptionEstilo3D = Inserted
        Call imprimir
    End Sub
    
    Private Sub opt1_Click()
        If opt1.Value = True Then
            TIPO = "(dd.tipo = 'FV' OR dd.tipo = 'NV' OR dd.tipo = 'FE')"
            Call generaInformePP(data, impresion, TIPO, detalle, dato1.text, fecha1, fecha2)
        End If
    End Sub
    
    Private Sub opt2_Click()
        If opt2.Value = True Then
            TIPO = "(dd.tipo = 'BV' OR dd.tipo = 'ZE')"
            Call generaInformePP(data, impresion, TIPO, detalle, dato1.text, fecha1, fecha2)
        End If
    End Sub
    
    Private Sub opt3_Click()
        If opt3.Value = True Then
            TIPO = "(dd.tipo = 'FV' OR dd.tipo = 'NV' OR dd.tipo = 'BV' OR dd.tipo = 'ZE' OR dd.tipo = 'FE')"
            Call generaInformePP(data, impresion, TIPO, detalle, dato1.text, fecha1, fecha2)
        End If
    End Sub

    Private Sub chk1_Click()
        If chk1.Value = 1 Then
            detalle = True
        Else
            detalle = False
        End If
        Call generaInformePP(data, impresion, TIPO, detalle, dato1.text, fecha1, fecha2)
    End Sub
    
    Private Sub imprimir()
        Dim i As Long
        impresion.AutoRedraw = False
        If detalle = True Then
            For i = 1 To ultimaFila - 1
                impresion.RowHeight(i) = altoFila
            Next i
        End If
        impresion.PageSetup.HeaderMargin = 1.5
        
        impresion.PageSetup.TopMargin = 1.5
        impresion.PageSetup.LeftMargin = 1
        impresion.PageSetup.RightMargin = 0.5
        impresion.PageSetup.BottomMargin = 1.75
        
        impresion.PageSetup.FooterMargin = 1.75
        impresion.PageSetup.BlackAndWhite = True
        impresion.PageSetup.PrintFixedRow = True
        impresion.PageSetup.Orientation = cellLandscape
        
        Call verificaImpresora(5, impresion)
        
        If detalle = True Then
            For i = 1 To ultimaFila
                impresion.RowHeight(i) = 0
            Next i
        End If
        
        impresion.AutoRedraw = True
    End Sub
    
    Private Sub leerEmpresas()
        Dim tabla As String
        tabla = "SELECT codigo, nombre, rubro FROM g_maestroempresas WHERE rubro = '" & rubro & "' ORDER BY codigo ASC"
        Call ConectarControlData(data, servidor, basedatos, usuario, password, tabla)
        If data.Recordset.RecordCount = 0 Then
            MsgBox "No existen empresas del rubro", vbOKOnly, "Mensaje"
            Unload Me
        Else
            cantLocales = data.Recordset.RecordCount
            ReDim cabezaLocales(4, cantLocales + 3) As String
        End If
    End Sub
    
