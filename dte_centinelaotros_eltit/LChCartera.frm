VERSION 5.00
Object = "{4777436C-EB8C-4596-98A8-EBCF98683969}#1.0#0"; "Flexcell.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{ADBBAED6-B16F-40DC-80DF-B44910CBA76C}#1.0#0"; "Frame-Xp.ocx"
Begin VB.Form LChCartera 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Listado Cheques en Cartera"
   ClientHeight    =   8430
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11610
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8430
   ScaleWidth      =   11610
   Begin MSAdodcLib.Adodc data 
      Height          =   330
      Left            =   240
      Top             =   7980
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
   Begin XPFrame.FrameXp FrameXp1 
      Height          =   6255
      Left            =   120
      TabIndex        =   9
      Top             =   1650
      Width           =   11385
      _ExtentX        =   20082
      _ExtentY        =   11033
      BackColor       =   16744576
      Caption         =   "Documentos"
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
         Height          =   5775
         Left            =   120
         TabIndex        =   8
         Top             =   360
         Width           =   11190
         _ExtentX        =   19738
         _ExtentY        =   10186
         DefaultFontSize =   8.25
         Rows            =   2
         SelectionMode   =   1
         DateFormat      =   2
      End
   End
   Begin XPFrame.FrameXp frmImprimir 
      Height          =   375
      Left            =   8220
      TabIndex        =   10
      Top             =   7980
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
   Begin XPFrame.FrameXp FrameXp2 
      Height          =   1275
      Left            =   1560
      TabIndex        =   11
      Top             =   120
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
      Begin VB.OptionButton opt1 
         BackColor       =   &H00FF8080&
         Caption         =   "Rango de fechas"
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
         Left            =   1020
         TabIndex        =   6
         Top             =   420
         Value           =   -1  'True
         Width           =   2715
      End
      Begin VB.OptionButton opt2 
         BackColor       =   &H00FF8080&
         Caption         =   "Todos los Cheques"
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
         Left            =   1020
         TabIndex        =   7
         Top             =   840
         Width           =   2715
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
         Left            =   5880
         MaxLength       =   2
         TabIndex        =   0
         Tag             =   "proveedor"
         Top             =   420
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
         Left            =   6360
         MaxLength       =   2
         TabIndex        =   1
         Tag             =   "proveedor"
         Top             =   420
         Width           =   435
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
         Left            =   6840
         MaxLength       =   4
         TabIndex        =   2
         Tag             =   "proveedor"
         Top             =   420
         Width           =   795
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
         Left            =   6840
         MaxLength       =   4
         TabIndex        =   5
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
         Left            =   5880
         MaxLength       =   2
         TabIndex        =   3
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
         Left            =   6360
         MaxLength       =   2
         TabIndex        =   4
         Tag             =   "proveedor"
         Top             =   840
         Width           =   435
      End
      Begin VB.Label lbl2 
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
         Left            =   4440
         TabIndex        =   13
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label lbl1 
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
         Left            =   4440
         TabIndex        =   12
         Top             =   420
         Width           =   1335
      End
   End
End
Attribute VB_Name = "LChCartera"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Private formatoGrilla(10, 10) As String
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
            fecha1 = dato3.text & "-" & dato2.text & "-" & dato1.text
            fecha2 = ""
            If opt2.Value = True Then
                Call cargaCarteraCheques
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
            Call cargaCarteraCheques
        End If
    End Sub
    '========================================================
    'KeyPress
    '========================================================

    '========================================================
    'KeyUp
    '========================================================
    Private Sub dato1_KeyUp(KeyCode As Integer, Shift As Integer)
        'If Len(dato1.text) = dato1.MaxLength Then
        '    Call dato1_KeyPress(13)
        'End If
    End Sub
    
    Private Sub dato2_KeyUp(KeyCode As Integer, Shift As Integer)
        'If Len(dato2.text) = dato2.MaxLength Then
        '    Call dato2_KeyPress(13)
        'End If
    End Sub
    
    Private Sub dato3_KeyUp(KeyCode As Integer, Shift As Integer)
        'If Len(dato3.text) = dato3.MaxLength Then
        '    Call dato3_KeyPress(13)
        'End If
    End Sub
    '========================================================
    'KeyUp
    '========================================================
    
    '========================================================
    'LostFocus
    '========================================================
 
    Private Sub dato1_LostFocus()
    Call limpiaBarra(2)
    Call esfecha(dato1, dato2, dato3, "dd")
    End Sub
    Private Sub dato2_LostFocus()
    Call esfecha(dato1, dato2, dato3, "mm")
    End Sub
    Private Sub dato3_LostFocus()
    Call esfecha(dato1, dato2, dato3, "yyyy")
    End Sub
    Private Sub dato4_LostFocus()
    Call esfecha(dato4, dato5, dato6, "dd")
    End Sub
    Private Sub dato5_LostFocus()
    Call esfecha(dato4, dato5, dato6, "mm")
    End Sub
    Private Sub dato6_LostFocus()
    Call esfecha(dato4, dato5, dato6, "yyyy")
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
    If KeyCode = 27 Then
        Unload Me
    End If
    If KeyCode = 38 Then
        If Screen.ActiveForm.ActiveControl.Name = "Documentos" Then
            Unload Me
        End If
    End If
End Sub

Private Sub Form_Load()
    Dim tabla As String
    Call Centrar(Me)
    Call CargaGrillaDocumentos(1, 7)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Principal.barraEstado.Panels(1).text = UCase(Principal.Caption)
    Call limpiaBarra(2)
End Sub


'****************************************************************************
'Formato de la Grilla Documentos
'****************************************************************************
    Private Sub CargaGrillaDocumentos(ByVal row As Integer, ByVal col As Integer)
        Dim i As Integer
        Rem DATOS DE LA COLUMNA
        formatoGrilla(1, 1) = "CLIENTE"
        formatoGrilla(1, 2) = "F.RECP."
        formatoGrilla(1, 3) = "N° CHEQUE"
        formatoGrilla(1, 4) = "F.VCTO."
        formatoGrilla(1, 5) = "DOCUMENTO"
        formatoGrilla(1, 6) = "MONTO"
        
        Rem LARGO DE LOS DATOS
        formatoGrilla(2, 1) = "35"
        formatoGrilla(2, 2) = "10"
        formatoGrilla(2, 3) = "10"
        formatoGrilla(2, 4) = "10"
        formatoGrilla(2, 5) = "12"
        formatoGrilla(2, 6) = "12"
        
        Rem TIPO DE DATOS C=CEROS,S=STRING,N=NUMERICO,D=DATE,
        formatoGrilla(3, 1) = "S"
        formatoGrilla(3, 2) = "C"
        formatoGrilla(3, 3) = "N"
        formatoGrilla(3, 4) = "C"
        formatoGrilla(3, 5) = "S"
        formatoGrilla(3, 6) = "N"
        
        Rem FORMATO GRILLA
        formatoGrilla(4, 1) = ""
        formatoGrilla(4, 2) = ""
        formatoGrilla(4, 3) = ""
        formatoGrilla(4, 4) = ""
        formatoGrilla(4, 5) = ""
        formatoGrilla(4, 6) = "$ ###,###,##0"
        
        Rem LOCCKED
        formatoGrilla(5, 1) = "FALSE"
        formatoGrilla(5, 2) = "FALSE"
        formatoGrilla(5, 3) = "FALSE"
        formatoGrilla(5, 4) = "fALSE"
        formatoGrilla(5, 5) = "FALSE"
        formatoGrilla(5, 6) = "FALSE"
        
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
        formatoGrilla(8, 1) = "30"
        formatoGrilla(8, 2) = "9"
        formatoGrilla(8, 3) = "9"
        formatoGrilla(8, 4) = "9"
        formatoGrilla(8, 5) = "9"
        formatoGrilla(8, 6) = "9"
            
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
        
        Documentos.Column(0).Width = 0
        For i = 1 To col - 1
            Documentos.Cell(0, i).text = formatoGrilla(1, i)
            Documentos.Column(i).Width = Val(formatoGrilla(8, i)) * (Documentos.Cell(0, i).Font.Size + 1.25)
            Documentos.Column(i).MaxLength = Val(formatoGrilla(2, i))
            Documentos.Column(i).FormatString = formatoGrilla(4, i)
            Documentos.Column(i).Locked = formatoGrilla(5, i)
            If formatoGrilla(3, i) = "N" Then
                Documentos.Column(i).Alignment = cellRightCenter
            End If
            If formatoGrilla(3, i) = "S" Then
                Documentos.Column(i).Alignment = cellLeftCenter
            End If
            If formatoGrilla(3, i) = "C" Then
                Documentos.Column(i).Alignment = cellCenterCenter
            End If
        Next i
        Documentos.Range(0, 0, 0, Documentos.Cols - 1).Alignment = cellCenterCenter
        Documentos.Enabled = True
    End Sub
'****************************************************************************
'Formato de la Grilla Documentos
'****************************************************************************

    Private Sub frmImprimir_BarMouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
        frmImprimir.ColorBarraAbajo = &HFFC0C0
        frmImprimir.ColorBarraArriba = &H800000
        frmImprimir.CaptionEstilo3D = Raised
    End Sub

    Private Sub frmImprimir_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
        frmImprimir.ColorBarraAbajo = &H800000
        frmImprimir.ColorBarraArriba = &HFFC0C0
        frmImprimir.CaptionEstilo3D = Inserted
        Call imprimir
    End Sub

    Private Sub imprimir()
        Call cabezaInforme("", Documentos, "LISTADO DE CHEQUES EN CARTERA", 1)
        Documentos.PageSetup.HeaderMargin = 1
        Documentos.PageSetup.TopMargin = 1
        Documentos.PageSetup.LeftMargin = 1.5
        Documentos.PageSetup.RightMargin = 1
        Documentos.PageSetup.PrintFixedRow = True
        Documentos.PageSetup.BlackAndWhite = True
        Documentos.Range(0, 0, 0, Documentos.Cols - 1).Borders(cellEdgeBottom) = cellThin
        
        Call verificaImpresora(5, Documentos)
        
    End Sub

    Private Sub cargaCarteraCheques()
        Dim tabla As String
        Dim totalFecha As Double
        Dim total As Double
        Dim fecha As String
        Dim dias As Integer
        Dim i As Integer
        
        fecha = Format(DateAdd("d", -1, fecha1), "yyyy-mm-dd")
        
        tabla = "SELECT CONCAT(cch.rut, ' ', mc.nombre, '" & vbTab & "', DATE_FORMAT(cch.fecharecepcion, '%d-%m-%Y'), '" & vbTab & "', cch.numerocheque, '" & vbTab & "', IF(cch.fecharecepcion = cch.fechavencimiento, DATE_FORMAT(DATE_ADD(cch.fechavencimiento, INTERVAL 1 DAY),'%d-%m-%Y'), DATE_FORMAT(cch.fechavencimiento,'%d-%m-%Y')), '" & vbTab & "', cch.tipodocumento, ' ', cch.numero, '" & vbTab & "', monto) AS item, IF(cch.fecharecepcion = cch.fechavencimiento, DATE_ADD(cch.fechavencimiento, INTERVAL 1 DAY), cch.fechavencimiento) AS vencimiento, monto "
        tabla = tabla & "FROM sv_carteracheques AS cch INNER JOIN " & baseVentas & ".sv_maestroclientes AS mc ON cch.rut = mc.rut AND cch.sucursal = mc.sucursal AND mc.sucursal = '0' "
        tabla = tabla & "WHERE local = '" & empresaActiva & "' AND cch.fecharecepcion = cch.fechavencimiento  AND cch.fechavencimiento =  '" & fecha & "' "
        If fecha2 <> "" Then
            tabla = tabla & "UNION "
            tabla = tabla & "SELECT CONCAT(cch.rut, ' ', mc.nombre, '" & vbTab & "', DATE_FORMAT(cch.fecharecepcion, '%d-%m-%Y'), '" & vbTab & "', cch.numerocheque, '" & vbTab & "', DATE_FORMAT(cch.fechavencimiento,'%d-%m-%Y'), '" & vbTab & "', cch.tipodocumento, ' ', cch.numero, '" & vbTab & "', monto) AS item, cch.fechavencimiento AS vencimiento, monto "
            tabla = tabla & "FROM sv_carteracheques AS cch INNER JOIN " & baseVentas & ".sv_maestroclientes AS mc ON cch.rut = mc.rut AND cch.sucursal = mc.sucursal AND mc.sucursal = '0' "
            tabla = tabla & "WHERE local = '" & empresaActiva & "' AND cch.fecharecepcion <> cch.fechavencimiento  AND cch.fechavencimiento BETWEEN '" & fecha1 & "' AND '" & fecha2 & "' "
            dias = DateDiff("d", fecha1, fecha2)
            fecha = Format(fecha1, "yyyy-mm-dd")
            'HACER UN FOR CON LA CANTIDAD DE DIAS
            For i = 1 To dias
                tabla = tabla & "UNION "
                tabla = tabla & "SELECT CONCAT(cch.rut, ' ', mc.nombre, '" & vbTab & "', DATE_FORMAT(cch.fecharecepcion, '%d-%m-%Y'), '" & vbTab & "', cch.numerocheque, '" & vbTab & "', IF(cch.fecharecepcion = cch.fechavencimiento, DATE_FORMAT(DATE_ADD(cch.fechavencimiento, INTERVAL 1 DAY),'%d-%m-%Y'), DATE_FORMAT(cch.fechavencimiento,'%d-%m-%Y')), '" & vbTab & "', cch.tipodocumento, ' ', cch.numero, '" & vbTab & "', monto) AS item, IF(cch.fecharecepcion = cch.fechavencimiento, DATE_ADD(cch.fechavencimiento, INTERVAL 1 DAY), cch.fechavencimiento) AS vencimiento, monto "
                tabla = tabla & "FROM sv_carteracheques AS cch INNER JOIN " & baseVentas & ".sv_maestroclientes AS mc ON cch.rut = mc.rut AND cch.sucursal = mc.sucursal AND mc.sucursal = '0' "
                tabla = tabla & "WHERE local = '" & empresaActiva & "' AND cch.fecharecepcion = cch.fechavencimiento  AND cch.fechavencimiento =  '" & fecha & "' "
                fecha = Format(DateAdd("d", 1, fecha), "yyyy-mm-dd")
            Next i
        Else
            tabla = tabla & "UNION "
            tabla = tabla & "SELECT CONCAT(cch.rut, ' ', mc.nombre, '" & vbTab & "', DATE_FORMAT(cch.fecharecepcion, '%d-%m-%Y'), '" & vbTab & "', cch.numerocheque, '" & vbTab & "', IF(cch.fecharecepcion = cch.fechavencimiento, DATE_FORMAT(DATE_ADD(cch.fechavencimiento, INTERVAL 1 DAY),'%d-%m-%Y'), DATE_FORMAT(cch.fechavencimiento,'%d-%m-%Y')), '" & vbTab & "', cch.tipodocumento, ' ', cch.numero, '" & vbTab & "', monto) AS item, IF(cch.fecharecepcion = cch.fechavencimiento, DATE_ADD(cch.fechavencimiento, INTERVAL 1 DAY), cch.fechavencimiento) AS vencimiento, monto "
            tabla = tabla & "FROM sv_carteracheques AS cch INNER JOIN " & baseVentas & ".sv_maestroclientes AS mc ON cch.rut = mc.rut AND cch.sucursal = mc.sucursal AND mc.sucursal = '0' "
            tabla = tabla & "WHERE local = '" & empresaActiva & "' AND cch.fechavencimiento >= '" & fecha1 & "' "
        End If
        tabla = tabla & "ORDER BY vencimiento ASC "
        
        'tabla = "SELECT CONCAT(cch.rut, ' ', mc.nombre, '" & vbTab & "', DATE_FORMAT(cch.fecharecepcion, '%d-%m-%Y'), '" & vbTab & "', cch.numerocheque, '" & vbTab & "', IF(cch.fecharecepcion = cch.fechavencimiento, DATE_FORMAT(DATE_ADD(cch.fechavencimiento, INTERVAL 1 DAY),'%d-%m-%Y'), DATE_FORMAT(cch.fechavencimiento,'%d-%m-%Y')), '" & vbTab & "', cch.tipodocumento, ' ', cch.numero, '" & vbTab & "', monto) AS item, IF(cch.fecharecepcion = cch.fechavencimiento, DATE_FORMAT(DATE_ADD(cch.fechavencimiento, INTERVAL 1 DAY),'%d-%m-%Y'), DATE_FORMAT(cch.fechavencimiento,'%d-%m-%Y')) AS vencimiento, monto "
        'tabla = tabla & "FROM sv_carteracheques AS cch INNER JOIN molino_ventas.sv_maestroclientes AS mc ON cch.rut = mc.rut AND cch.sucursal = mc.sucursal AND mc.sucursal = '0' "
        'If fecha2 = "" Then
        '    tabla = tabla & "WHERE local = '" & empresaActiva & "' AND fechavencimiento >= '" & Format(fecha1, "yyyy-mm-dd") & "' "
        'Else
        '    tabla = tabla & "WHERE local = '" & empresaActiva & "' AND fechavencimiento BETWEEN '" & Format(fecha1, "yyyy-mm-dd") & "' AND '" & Format(fecha2, "yyyy-mm-dd") & "' "
        'End If
        'tabla = tabla & "ORDER BY vencimiento ASC"
        
        Call ConectarControlData(data, servidor, baseVentas & empresaActiva, usuario, password, tabla)
        
        Documentos.Rows = 1
        Documentos.AutoRedraw = False
        
        If data.Recordset.RecordCount > 0 Then
            data.Recordset.MoveFirst
            totalFecha = 0
            total = 0
            fecha = Format(data.Recordset.Fields("vencimiento"), "dd-mm-yyyy")
            While Not data.Recordset.EOF
                If fecha <> Format(data.Recordset.Fields("vencimiento"), "dd-mm-yyyy") Then
                    Documentos.AddItem vbTab & "TOTAL A LA FECHA " & fecha & vbTab & vbTab & vbTab & vbTab & Format(totalFecha, "$ ###,###,##0"), True
                    Documentos.Range(Documentos.Rows - 1, 2, Documentos.Rows - 1, Documentos.Cols - 2).Merge
                    Documentos.Range(Documentos.Rows - 1, 1, Documentos.Rows - 1, Documentos.Cols - 2).Alignment = cellCenterCenter
                    Documentos.Range(Documentos.Rows - 1, 2, Documentos.Rows - 1, Documentos.Cols - 1).FontBold = True
                    Documentos.Range(Documentos.Rows - 1, 2, Documentos.Rows - 1, Documentos.Cols - 1).Borders(cellEdgeTop) = cellThin
                    
                    totalFecha = 0
                    fecha = Format(data.Recordset.Fields("vencimiento"), "dd-mm-yyyy")
                    Documentos.AddItem "", True
                End If
                Documentos.AddItem data.Recordset.Fields("item"), True
                totalFecha = totalFecha + CDbl(data.Recordset.Fields("monto"))
                total = total + CDbl(data.Recordset.Fields("monto"))
                data.Recordset.MoveNext
            Wend
            
            Documentos.AddItem vbTab & "TOTAL A LA FECHA " & fecha & vbTab & vbTab & vbTab & vbTab & Format(totalFecha, "$ ###,###,##0"), True
            Documentos.Range(Documentos.Rows - 1, 2, Documentos.Rows - 1, Documentos.Cols - 2).Merge
            Documentos.Range(Documentos.Rows - 1, 1, Documentos.Rows - 1, Documentos.Cols - 2).Alignment = cellCenterCenter
            Documentos.Range(Documentos.Rows - 1, 2, Documentos.Rows - 1, Documentos.Cols - 1).FontBold = True
            Documentos.Range(Documentos.Rows - 1, 2, Documentos.Rows - 1, Documentos.Cols - 1).Borders(cellEdgeTop) = cellThin
            
            Documentos.AddItem "", True
            
            Documentos.AddItem vbTab & "TOTAL GENERAL CHEQUES EN CARTERA" & vbTab & vbTab & vbTab & vbTab & Format(total, "$ ###,###,##0"), True
            Documentos.Range(Documentos.Rows - 1, 2, Documentos.Rows - 1, Documentos.Cols - 2).Merge
            Documentos.Range(Documentos.Rows - 1, 1, Documentos.Rows - 1, Documentos.Cols - 2).Alignment = cellCenterCenter
            Documentos.Range(Documentos.Rows - 1, 2, Documentos.Rows - 1, Documentos.Cols - 1).FontBold = True
            Documentos.Range(Documentos.Rows - 1, 2, Documentos.Rows - 1, Documentos.Cols - 1).Borders(cellEdgeTop) = cellThin
        End If
        Documentos.AutoRedraw = True
        Documentos.Refresh
    End Sub
    
Private Sub opt1_Click()
    lbl2.Visible = True
    dato4.Visible = True
    dato5.Visible = True
    dato6.Visible = True
    If dato1.text <> "" Then
        If dato2.text <> "" Then
            If dato3.text <> "" Then
                fecha1 = dato3.text & "-" & dato2.text & "-" & dato1.text
                dato4.SetFocus
            Else
                dato3.SetFocus
            End If
        Else
            dato2.SetFocus
        End If
    Else
        dato1.SetFocus
    End If
End Sub

Private Sub opt2_Click()
    lbl2.Visible = False
    dato4.Visible = False
    dato5.Visible = False
    dato6.Visible = False
    If dato1.text <> "" Then
        If dato2.text <> "" Then
            If dato3.text <> "" Then
                fecha1 = dato3.text & "-" & dato2.text & "-" & dato1.text
                fecha2 = ""
                Call cargaCarteraCheques
            Else
                dato3.SetFocus
            End If
        Else
            dato2.SetFocus
        End If
    Else
        dato1.SetFocus
    End If
End Sub
