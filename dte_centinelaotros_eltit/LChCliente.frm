VERSION 5.00
Object = "{4777436C-EB8C-4596-98A8-EBCF98683969}#1.0#0"; "Flexcell.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{ADBBAED6-B16F-40DC-80DF-B44910CBA76C}#1.0#0"; "Frame-Xp.ocx"
Begin VB.Form LChCliente 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Listado Cheques por Cliente"
   ClientHeight    =   7260
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10455
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7260
   ScaleWidth      =   10455
   Begin MSAdodcLib.Adodc data 
      Height          =   375
      Left            =   240
      Top             =   6780
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   2143
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
   Begin XPFrame.FrameXp FrameXp1 
      Height          =   1095
      Left            =   120
      TabIndex        =   1
      Top             =   135
      Width           =   10215
      _ExtentX        =   18018
      _ExtentY        =   1931
      BackColor       =   16744576
      Caption         =   "Cliente"
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
         Height          =   285
         Left            =   1560
         MaxLength       =   9
         TabIndex        =   0
         Tag             =   "proveedor"
         Top             =   360
         Width           =   1455
      End
      Begin VB.OptionButton opt1 
         BackColor       =   &H00FF8080&
         Caption         =   "Todos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   720
         Value           =   -1  'True
         Width           =   1815
      End
      Begin VB.OptionButton opt2 
         BackColor       =   &H00FF8080&
         Caption         =   "Cartola Pendientes"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   2400
         TabIndex        =   3
         Top             =   720
         Width           =   2295
      End
      Begin VB.Label lbl1 
         Appearance      =   0  'Flat
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Rut"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   120
         TabIndex        =   7
         Top             =   360
         Width           =   1335
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
         Left            =   3000
         TabIndex        =   6
         Top             =   360
         Width           =   615
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
         Left            =   3720
         TabIndex        =   5
         Top             =   360
         Width           =   6375
      End
   End
   Begin XPFrame.FrameXp FrameXp2 
      Height          =   5355
      Left            =   120
      TabIndex        =   2
      Top             =   1320
      Width           =   10215
      _ExtentX        =   18018
      _ExtentY        =   9446
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
         Height          =   4875
         Left            =   120
         TabIndex        =   8
         Top             =   360
         Width           =   9975
         _ExtentX        =   17595
         _ExtentY        =   8599
         DefaultFontSize =   8.25
         Rows            =   2
         SelectionMode   =   1
         DateFormat      =   2
      End
   End
   Begin XPFrame.FrameXp frmImprimir 
      Height          =   375
      Left            =   7020
      TabIndex        =   9
      Top             =   6780
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
End
Attribute VB_Name = "LChCliente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Private formatoGrilla(10, 10) As String

'============================================================
'MANEJO DE LOS CONTOLES
'============================================================
    '========================================================
    'GotFocus
    '========================================================
    Private Sub dato1_GotFocus()
        Principal.barraEstado.Panels(2).text = "F2: Ayuda Cliente"
    End Sub
    '========================================================
    'GotFocus
    '========================================================
    
    '========================================================
    'KeyDown
    '========================================================
    Private Sub dato1_KeyDown(KeyCode As Integer, Shift As Integer)
        If KeyCode = vbKeyF2 Then
            Call ayudaClienteSin(dato1, lblDV)
        Else
            Call Flechas(KeyCode, dato1)
        End If
    End Sub
    '========================================================
    'KeyDown
    '========================================================
    
    '========================================================
    'KeyPress
    '========================================================
    Private Sub dato1_KeyPress(KeyAscii As Integer)
        Dim tabla As String
        KeyAscii = esNumero(KeyAscii)
        If KeyAscii = 13 Then
            dato1.text = ceros(dato1)
            lblDV.Caption = rut(dato1.text)
            LBLNOMBRE.Caption = leerNombreCliente(dato1.text & lblDV.Caption)
            If LBLNOMBRE.Caption <> "" Then
                If opt1.Value = True Then
                    tabla = "SELECT CONCAT(DATE_FORMAT(fecharecepcion,'%d-%m-%Y'), '" & vbTab & "', numerocheque, '" & vbTab & "', DATE_FORMAT(fechavencimiento,'%d-%m-%Y'), '" & vbTab & "', CONCAT(tipodocumento, ' ', numero), '" & vbTab & "', monto) AS item "
                    tabla = tabla & "FROM sv_carteracheques "
                    tabla = tabla & "WHERE local = '" & empresaActiva & "' AND rut = '" & dato1.text & lblDV.Caption & "' "
                    tabla = tabla & "ORDER BY fechavencimiento, monto ASC"
                End If
                If opt2.Value = True Then
                    tabla = "SELECT CONCAT(DATE_FORMAT(fecharecepcion,'%d-%m-%Y'), '" & vbTab & "', numerocheque, '" & vbTab & "', DATE_FORMAT(fechavencimiento,'%d-%m-%Y'), '" & vbTab & "', CONCAT(tipodocumento, ' ', numero), '" & vbTab & "', monto) AS item "
                    tabla = tabla & "FROM sv_carteracheques "
                    tabla = tabla & "WHERE local = '" & empresaActiva & "' AND rut = '" & dato1.text & lblDV.Caption & "' AND fechavencimiento > '" & Format(fechasistema, "yyyy-mm-dd") & "'"
                    tabla = tabla & "ORDER BY fechavencimiento, monto ASC"
                End If
                Call ConectarControlData(data, servidor, baseVentas & empresaActiva, usuario, password, tabla)
                Call cargaInforme(data, Documentos)
                SendKeys "{Tab}"
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
        Call CargaGrillaDocumentos(1, 6)
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
        formatoGrilla(1, 1) = "RECIBIDO"
        formatoGrilla(1, 2) = "NÚMERO"
        formatoGrilla(1, 3) = "VENCIMIENTO"
        formatoGrilla(1, 4) = "DOCUMENTO"
        formatoGrilla(1, 5) = "MONTO"
        
        Rem LARGO DE LOS DATOS
        formatoGrilla(2, 1) = "10"
        formatoGrilla(2, 2) = "7"
        formatoGrilla(2, 3) = "10"
        formatoGrilla(2, 4) = "13"
        formatoGrilla(2, 5) = "9"
        
        Rem TIPO DE DATOS C=CEROS,S=STRING,N=NUMERICO,D=DATE,
        formatoGrilla(3, 1) = "N"
        formatoGrilla(3, 2) = "N"
        formatoGrilla(3, 3) = "N"
        formatoGrilla(3, 4) = "S"
        formatoGrilla(3, 5) = "N"
        
        Rem FORMATO GRILLA
        formatoGrilla(4, 1) = ""
        formatoGrilla(4, 2) = "0000000"
        formatoGrilla(4, 3) = ""
        formatoGrilla(4, 4) = ""
        formatoGrilla(4, 5) = "$ ###,###,##0"
        
        Rem LOCCKED
        formatoGrilla(5, 1) = "FALSE"
        formatoGrilla(5, 2) = "FALSE"
        formatoGrilla(5, 3) = "FALSE"
        formatoGrilla(5, 4) = "FALSE"
        formatoGrilla(5, 5) = "FALSE"
        
        Rem VALOR MINIMO
        formatoGrilla(6, 1) = ""
        formatoGrilla(6, 2) = ""
        formatoGrilla(6, 3) = ""
        formatoGrilla(6, 4) = ""
        formatoGrilla(6, 5) = ""
        
        Rem VALOR MAXIMO
        formatoGrilla(7, 1) = ""
        formatoGrilla(7, 2) = ""
        formatoGrilla(7, 3) = ""
        formatoGrilla(7, 4) = ""
        formatoGrilla(7, 5) = ""
        
        Rem ANCHO
        formatoGrilla(8, 1) = "10"
        formatoGrilla(8, 2) = "15"
        formatoGrilla(8, 3) = "10"
        formatoGrilla(8, 4) = "15"
        formatoGrilla(8, 5) = "9"
            
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
            Else
                Documentos.Column(i).Alignment = cellLeftCenter
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

    Private Sub opt1_Click()
        Dim tabla As String
        opt1.Value = True
        tabla = "SELECT CONCAT(DATE_FORMAT(fecharecepcion,'%d-%m-%Y'), '" & vbTab & "', numerocheque, '" & vbTab & "', DATE_FORMAT(fechavencimiento,'%d-%m-%Y'), '" & vbTab & "', CONCAT(tipodocumento, ' ', numero), '" & vbTab & "', monto) AS item "
        tabla = tabla & "FROM sv_carteracheques "
        tabla = tabla & "WHERE local = '" & empresaActiva & "' AND rut = '" & dato1.text & lblDV.Caption & "' "
        tabla = tabla & "ORDER BY fechavencimiento, monto ASC"
        Call ConectarControlData(data, servidor, baseVentas & empresaActiva, usuario, password, tabla)
        Call cargaInforme(data, Documentos)
    End Sub
    
    Private Sub opt2_Click()
        opt2.Value = True
        tabla = "SELECT CONCAT(DATE_FORMAT(fecharecepcion,'%d-%m-%Y'), '" & vbTab & "', numerocheque, '" & vbTab & "', DATE_FORMAT(fechavencimiento,'%d-%m-%Y'), '" & vbTab & "', CONCAT(tipodocumento, ' ', numero), '" & vbTab & "', monto) AS item "
        tabla = tabla & "FROM sv_carteracheques "
        tabla = tabla & "WHERE local = '" & empresaActiva & "' AND rut = '" & dato1.text & lblDV.Caption & "' AND fechavencimiento > '" & Format(fechasistema, "yyyy-mm-dd") & "'"
        tabla = tabla & "ORDER BY fechavencimiento, monto ASC"
        Call ConectarControlData(data, servidor, baseVentas & empresaActiva, usuario, password, tabla)
        Call cargaInforme(data, Documentos)
    End Sub
