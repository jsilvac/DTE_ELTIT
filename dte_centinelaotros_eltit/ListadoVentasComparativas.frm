VERSION 5.00
Object = "{4777436C-EB8C-4596-98A8-EBCF98683969}#1.0#0"; "FlexCell.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{ADBBAED6-B16F-40DC-80DF-B44910CBA76C}#1.0#0"; "Frame-Xp.ocx"
Begin VB.Form ListadoVentasComparativas 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Listado de Ventas Comparativas"
   ClientHeight    =   8460
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   14055
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8460
   ScaleWidth      =   14055
   Begin XPFrame.FrameXp FrameXp1 
      Height          =   855
      Left            =   2760
      TabIndex        =   1
      Top             =   60
      Width           =   8535
      _ExtentX        =   15055
      _ExtentY        =   1508
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
         Left            =   1800
         MaxLength       =   2
         TabIndex        =   0
         Tag             =   "proveedor"
         Top             =   420
         Width           =   435
      End
      Begin VB.Label lbl2 
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
         Left            =   360
         TabIndex        =   6
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
         Left            =   2280
         TabIndex        =   5
         Top             =   420
         Width           =   5955
      End
   End
   Begin MSAdodcLib.Adodc data1 
      Height          =   330
      Left            =   120
      Top             =   8040
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
   Begin XPFrame.FrameXp frmInforme 
      Height          =   6915
      Left            =   60
      TabIndex        =   2
      Top             =   1020
      Width           =   13935
      _ExtentX        =   24580
      _ExtentY        =   12197
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
         Height          =   6435
         Left            =   60
         TabIndex        =   3
         Top             =   420
         Width           =   13815
         _ExtentX        =   24368
         _ExtentY        =   11351
         Cols            =   5
         DefaultFontSize =   8.25
         Rows            =   30
      End
   End
   Begin XPFrame.FrameXp frmImprimir 
      Height          =   375
      Left            =   10620
      TabIndex        =   4
      Top             =   8040
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
   Begin MSAdodcLib.Adodc data2 
      Height          =   330
      Left            =   1500
      Top             =   8040
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
End
Attribute VB_Name = "ListadoVentasComparativas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Private TIPO As String
    Private posX As Single
    Private posY As Single
        
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
    '========================================================
    'GotFocus
    '========================================================
    
    '========================================================
    'KeyDown
    '========================================================
    Private Sub dato1_KeyDown(KeyCode As Integer, Shift As Integer)
        Select Case KeyCode
            Case vbKeyF2
                Call ayudaEmpresa(dato1)
            Case Else
                Call Flechas(KeyCode, dato1)
        End Select
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
                Call generaInformeLVC(data1, data2, impresion, TIPO, dato1.text)
            End If
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
        Call CargaGrillaInforme(2, 8)
        TIPO = "(dd.tipo = 'FV' OR dd.tipo = 'FE' OR dd.tipo = 'BV' OR dd.tipo = 'ZE')"
    End Sub

'****************************************************************************
'Formato de la Grilla Listado de Ventas
'****************************************************************************
    Private Sub CargaGrillaInforme(ByVal row As Integer, ByVal col As Integer)
        Dim formatoGrilla(10, 20) As String
        Dim i As Integer
        
        Call leerMeses
        impresion.FixedRows = 2
        
        Rem DATOS DE LA COLUMNA
        formatoGrilla(1, 1) = "DESCRIPCION"
        formatoGrilla(1, 2) = "INFORMACION DEL MES"
        formatoGrilla(1, 3) = ""
        formatoGrilla(1, 4) = ""
        formatoGrilla(1, 5) = "INFORMACION ACUMULADA"
        formatoGrilla(1, 6) = ""
        formatoGrilla(1, 7) = ""
        
        Rem LARGO DE LOS DATOS
        formatoGrilla(2, 1) = "50"
        formatoGrilla(2, 2) = "9"
        formatoGrilla(2, 3) = "9"
        formatoGrilla(2, 4) = "9"
        formatoGrilla(2, 5) = "9"
        formatoGrilla(2, 6) = "9"
        formatoGrilla(2, 7) = "9"
        
        Rem TIPO DE DATOS C=CEROS,S=STRING,N=NUMERICO,D=DATE,
        formatoGrilla(3, 1) = "S"
        formatoGrilla(3, 2) = "N"
        formatoGrilla(3, 3) = "N"
        formatoGrilla(3, 4) = "N"
        formatoGrilla(3, 5) = "N"
        formatoGrilla(3, 6) = "N"
        formatoGrilla(3, 7) = "N"
        
        Rem FORMATO GRILLA
        formatoGrilla(4, 1) = ""
        formatoGrilla(4, 2) = "###,###,##0"
        formatoGrilla(4, 3) = "$ ###,###,##0"
        formatoGrilla(4, 4) = "$ ###,###,##0.00"
        formatoGrilla(4, 5) = "###,###,##0"
        formatoGrilla(4, 6) = "$ ###,###,##0"
        formatoGrilla(4, 7) = "$ ###,###,##0.00"
        
        Rem LOCCKED
        formatoGrilla(5, 1) = "FALSE"
        formatoGrilla(5, 2) = "FALSE"
        formatoGrilla(5, 3) = "FALSE"
        formatoGrilla(5, 4) = "FALSE"
        formatoGrilla(5, 5) = "FALSE"
        formatoGrilla(5, 6) = "FALSE"
        formatoGrilla(5, 7) = "FALSE"
        
        Rem VALOR MINIMO
        formatoGrilla(6, 1) = ""
        formatoGrilla(6, 2) = ""
        formatoGrilla(6, 3) = ""
        formatoGrilla(6, 4) = ""
        formatoGrilla(6, 5) = ""
        formatoGrilla(6, 6) = ""
        formatoGrilla(6, 7) = ""
        
        Rem VALOR MAXIMO
        formatoGrilla(7, 1) = ""
        formatoGrilla(7, 2) = ""
        formatoGrilla(7, 3) = ""
        formatoGrilla(7, 4) = ""
        formatoGrilla(7, 5) = ""
        formatoGrilla(7, 6) = ""
        formatoGrilla(7, 7) = ""
        
        Rem ANCHO
        formatoGrilla(8, 1) = "23"
        formatoGrilla(8, 2) = "12"
        formatoGrilla(8, 3) = "12"
        formatoGrilla(8, 4) = "12"
        formatoGrilla(8, 5) = "12"
        formatoGrilla(8, 6) = "12"
        formatoGrilla(8, 7) = "12"
        
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
        impresion.Cell(1, 2).text = "         "
        impresion.Cell(1, 3).text = "NETO"
        impresion.Cell(1, 4).text = "PRECIO PROMEDIO"
        impresion.Cell(1, 5).text = "         "
        impresion.Cell(1, 6).text = "NETO"
        impresion.Cell(1, 7).text = "PRECIO PROMEDIO"
        impresion.Range(0, 1, 1, 1).Merge
        impresion.Range(0, 2, 0, 4).Merge
        impresion.Range(0, 5, 0, impresion.Cols - 1).Merge
        impresion.Range(0, 1, 1, impresion.Cols - 1).Alignment = cellCenterCenter
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
    
    Private Sub imprimir()
        Dim i As Long
        impresion.AutoRedraw = False
        
        impresion.PageSetup.HeaderMargin = 2
    
        impresion.PageSetup.TopMargin = 2
        impresion.PageSetup.LeftMargin = 2
        impresion.PageSetup.RightMargin = 1.5
        impresion.PageSetup.BottomMargin = 2
        
        impresion.PageSetup.FooterMargin = 2
        impresion.PageSetup.BlackAndWhite = True
        impresion.PageSetup.Orientation = cellLandscape
        
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
        Dim i As Integer
        Dim ancho As Single
        Dim alto As Single
        Static col As Single
        Static fila As Long
        
        Call GetCursorPos(mouse)
        posX = mouse.x
        posY = mouse.Y
        
        If fila = impresion.MouseRow And col = impresion.MouseCol Then
            Unload ComprasClienteGrafico
            fila1 = 0
            fila2 = 0
            fila = 0
            col1 = 0
            col2 = 0
        Else
            fila = impresion.MouseRow
            col = impresion.MouseCol
            If col > 1 And col < 5 Then
                col1 = col
                col2 = col1 + 3
            End If
            If col > 4 Then
                col1 = col - 3
                col2 = col
            End If
            If col = 1 Then
                col1 = 0
                col2 = 0
            End If
            If fila = -1 Then
                fila = 0
            Else
                fila1 = fila
                If fila1 <> 0 And impresion.Cell(fila1, col1).text <> "" And IsNumeric(impresion.Cell(fila1, col1).text) = True Then
                    tipoGrafico = 14
                    For i = 1 To 12
                        If impresion.Cell(fila1, 1).text = meses(i) Or impresion.Cell(fila1, 1).text = "TOTAL AÑO" Then
                            tipoGrafico = 1
                            Exit For
                        End If
                    Next i
                    If tipoGrafico = 14 Then
                        If col1 = 4 Then
                            tipoGrafico = 1
                        End If
                    End If
                    Unload ListadoVentasComparativasGrafico
                    Load ListadoVentasComparativasGrafico
                    With ListadoVentasComparativasGrafico
                        ancho = .Width
                        alto = .Height
                        .Left = posX * Screen.TwipsPerPixelX
                        .Top = posY * Screen.TwipsPerPixelY
                        If ancho + .Left > Screen.Width Then
                            .Left = .Left - ancho
                        End If
                        If alto + .Top > Screen.Height - 500 Then
                            .Top = .Top - alto
                        End If
                        .Show
                    End With
                End If
            End If
        End If
    End Sub

    








Private Sub impresion_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 And impresion.ActiveCell.col = 1 Then
impresion.Cell(impresion.ActiveCell.row, impresion.ActiveCell.col).Locked = False


End If

End Sub
