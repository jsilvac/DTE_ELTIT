VERSION 5.00
Object = "{4777436C-EB8C-4596-98A8-EBCF98683969}#1.0#0"; "Flexcell.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{ADBBAED6-B16F-40DC-80DF-B44910CBA76C}#1.0#0"; "Frame-Xp.ocx"
Begin VB.Form LCCliente 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Listado Cartola por Cliente"
   ClientHeight    =   7020
   ClientLeft      =   2010
   ClientTop       =   2310
   ClientWidth     =   11055
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7020
   ScaleWidth      =   11055
   Begin MSAdodcLib.Adodc data 
      Height          =   330
      Left            =   240
      Top             =   6540
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
      Height          =   735
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   10800
      _ExtentX        =   19050
      _ExtentY        =   1296
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
         Left            =   1575
         MaxLength       =   9
         TabIndex        =   0
         Tag             =   "proveedor"
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label lbl1 
         Appearance      =   0  'Flat
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Rut"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   120
         TabIndex        =   6
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
         Left            =   3060
         TabIndex        =   5
         Top             =   360
         Width           =   375
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
         TabIndex        =   4
         Top             =   360
         Width           =   6960
      End
   End
   Begin XPFrame.FrameXp FrameXp2 
      Height          =   5475
      Left            =   120
      TabIndex        =   3
      Top             =   960
      Width           =   10800
      _ExtentX        =   19050
      _ExtentY        =   9657
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
         Height          =   4995
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   10560
         _ExtentX        =   18627
         _ExtentY        =   8811
         DefaultFontSize =   8.25
         Rows            =   2
         SelectionMode   =   1
         DateFormat      =   2
      End
   End
   Begin XPFrame.FrameXp frmImprimir 
      Height          =   375
      Left            =   7620
      TabIndex        =   7
      Top             =   6540
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
Attribute VB_Name = "LCCliente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Private formatogrilla(10, 10) As String

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
        KeyAscii = esNumero(KeyAscii)
        If KeyAscii = 13 Then
            dato1.text = ceros(dato1)
            lblDV.Caption = rut(dato1.text)
            LBLNOMBRE.Caption = leerNombreCliente(dato1.text & lblDV.Caption)
            If LBLNOMBRE.Caption <> "" Then
                Call leerCreditoCliente(dato1.text & lblDV.Caption)
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
        Call CargaGrillaDocumentos(2, 8)
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
        formatogrilla(1, 1) = "FECHA"
        formatogrilla(1, 2) = "TIPO"
        formatogrilla(1, 3) = "NUMERO"
        formatogrilla(1, 4) = "GLOSA"
        formatogrilla(1, 5) = "DEBE"
        formatogrilla(1, 6) = "HABER"
        formatogrilla(1, 7) = "SALDO"
        
        Rem LARGO DE LOS DATOS
        formatogrilla(2, 1) = "10"
        formatogrilla(2, 2) = "2"
        formatogrilla(2, 3) = "12"
        formatogrilla(2, 4) = "20"
        formatogrilla(2, 5) = "9"
        formatogrilla(2, 6) = "9"
        formatogrilla(2, 7) = "9"
        
        Rem TIPO DE DATOS C=CEROS,S=STRING,N=NUMERICO,D=DATE,
        formatogrilla(3, 1) = "N"
        formatogrilla(3, 2) = "S"
        formatogrilla(3, 3) = "N"
        formatogrilla(3, 4) = "S"
        formatogrilla(3, 5) = "N"
        formatogrilla(3, 6) = "N"
        formatogrilla(3, 7) = "N"
        
        Rem FORMATO GRILLA
        formatogrilla(4, 1) = ""
        formatogrilla(4, 2) = ""
        formatogrilla(4, 3) = "000000000000"
        formatogrilla(4, 4) = ""
        formatogrilla(4, 5) = "$ ###,###,##0"
        formatogrilla(4, 6) = "$ ###,###,##0"
        formatogrilla(4, 7) = "$ ###,###,##0"
        
        Rem LOCCKED
        formatogrilla(5, 1) = "TRUE"
        formatogrilla(5, 2) = "TRUE"
        formatogrilla(5, 3) = "TRUE"
        formatogrilla(5, 4) = "TRUE"
        formatogrilla(5, 5) = "TRUE"
        formatogrilla(5, 6) = "TRUE"
        formatogrilla(5, 7) = "TRUE"
        
        Rem VALOR MINIMO
        formatogrilla(6, 1) = ""
        formatogrilla(6, 2) = ""
        formatogrilla(6, 3) = ""
        formatogrilla(6, 4) = ""
        formatogrilla(6, 5) = ""
        formatogrilla(6, 6) = ""
        formatogrilla(6, 7) = ""
        
        Rem VALOR MAXIMO
        formatogrilla(7, 1) = ""
        formatogrilla(7, 2) = ""
        formatogrilla(7, 3) = ""
        formatogrilla(7, 4) = ""
        formatogrilla(7, 5) = ""
        formatogrilla(7, 6) = ""
        formatogrilla(7, 7) = ""
        
        Rem ANCHO
        formatogrilla(8, 1) = "8"
        formatogrilla(8, 2) = "3"
        formatogrilla(8, 3) = "9"
        formatogrilla(8, 4) = "25"
        formatogrilla(8, 5) = "9"
        formatogrilla(8, 6) = "9"
        formatogrilla(8, 7) = "9"
            
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
        Documentos.Range(0, 0, 0, Documentos.Cols - 1).Alignment = cellCenterCenter
        Documentos.Enabled = True
    End Sub
'****************************************************************************
'Formato de la Grilla Documentos
'****************************************************************************

'============================================================
'LEER CREDITOS CLIENTE
'============================================================
    Private Sub leerCreditoCliente(ByVal rut As String)
        Dim tabla As String
        tabla = "SELECT CONCAT(DATE_FORMAT(dc.fecha,'%d-%m-%Y'), '" & vbTab & "', dc.tipo, '" & vbTab & "', dc.numero, '" & vbTab & "', IF(dc.tipo = 'NV', CONCAT('ABONA NOTA DE CREDITO', '" & vbTab & "', ' $ 0', '" & vbTab & "', dc.total), CONCAT('DOC. VENCE EL ', DATE_FORMAT(dc.vencimiento,'%d-%m-%Y'), '" & vbTab & "', dc.total, '" & vbTab & "', ' $ 0'))) AS item, dc.numero, IF(dc.tipo = 'NV', 'H', 'D') AS tipo "
        tabla = tabla & "FROM sv_documento_cabeza_" + empresaActiva + " AS dc "
        tabla = tabla & "WHERE dc.local = '" & empresaActiva & "' AND dc.tipo <> 'GD' AND dc.rut = '" & rut & "' AND dc.nula = 'N' "
        tabla = tabla & "UNION "
        tabla = tabla & "SELECT CONCAT(DATE_FORMAT(dp.fecha,'%d-%m-%Y'), '" & vbTab & "', dp.tipo, '" & vbTab & "', dp.numero, '" & vbTab & "', IF(dp.tipopago = '1', 'DOC. CANCELADO EN EFECTIVO', CONCAT('DOC. CANCELADO CON CHEQUE ', cch.numerocheque)), '" & vbTab & "', ' $ 0', '" & vbTab & "', SUM(dp.monto)) AS item, dp.numero, 'H' AS tipo "
        tabla = tabla & "FROM sv_documento_pagos_" + empresaActiva + " AS dp LEFT JOIN sv_carteracheques AS cch ON dp.tipo = cch.tipodocumento AND dp.numero = cch.numero "
        tabla = tabla & "WHERE dp.local = '" & empresaActiva & "' AND dp.tipo <> 'GD' AND dp.rut = '" & rut & "' AND dp.tipopago <> '4' AND dp.monto > 0 "
        tabla = tabla & "GROUP BY dp.numero "
        tabla = tabla & "UNION "
        tabla = tabla & "SELECT CONCAT(DATEpagos_detalle_FORMAT(pd.fecha,'%d-%m-%Y'), '" & vbTab & "', pd.tipo, '" & vbTab & "', pd.documento, '" & vbTab & "', CONCAT('CANCELA C. I. ', pd.numero), '" & vbTab & "', ' $ 0', '" & vbTab & "', pd.monto) AS item, pd.documento, 'H' AS tipo "
        tabla = tabla & "FROM sv__" & empresaActiva & " AS pd "
        tabla = tabla & "WHERE pd.local = '" & empresaActiva & "' AND pd.tipo <> 'GD' AND pd.rut = '" & rut & "' "
        tabla = tabla & "ORDER BY numero, tipo ASC"
        Call ConectarControlData(data, servidor, baseVentas & empresaActiva, usuario, password, tabla)
        Documentos.AutoRedraw = False
        Documentos.Rows = 1
        If data.Recordset.RecordCount > 0 Then
            data.Recordset.MoveFirst
            While Not data.Recordset.EOF
                Documentos.AddItem data.Recordset.Fields("item"), True
                data.Recordset.MoveNext
            Wend
        End If
        Documentos.AutoRedraw = True
        Documentos.Refresh
        Call sumaSaldos
    End Sub
'============================================================
'LEER CREDITOS CLIENTE
'============================================================

'============================================================
'SUMA SALDOS
'============================================================
    Private Sub sumaSaldos()
        Dim saldo As Double
        Dim i As Long
        saldo = 0
        For i = 1 To Documentos.Rows - 1
            saldo = saldo + CDbl(Documentos.Cell(i, 5).text) - CDbl(Documentos.Cell(i, 6).text)
            Documentos.Cell(i, 7).text = saldo
        Next i
    End Sub
'============================================================
'SUMA SALDOS
'============================================================

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
        Call cabezaInforme("", Documentos, "CARTOLA CLIENTE " & dato1.text & "-" & lblDV.Caption, 1)
        Documentos.PageSetup.HeaderMargin = 1
        Documentos.PageSetup.TopMargin = 1
        Documentos.PageSetup.LeftMargin = 1.5
        Documentos.PageSetup.RightMargin = 1
        Documentos.PageSetup.PrintFixedRow = True
        Documentos.PageSetup.BlackAndWhite = True
        Documentos.Range(0, 0, 0, Documentos.Cols - 1).Borders(cellEdgeBottom) = cellThin
        
        Call verificaImpresora(5, Documentos)
        
    End Sub
