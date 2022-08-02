VERSION 5.00
Object = "{4777436C-EB8C-4596-98A8-EBCF98683969}#1.0#0"; "FlexCell.ocx"
Object = "{ADBBAED6-B16F-40DC-80DF-B44910CBA76C}#1.0#0"; "Frame-Xp.ocx"
Begin VB.Form prove0013 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "PUBLICIDAD"
   ClientHeight    =   9015
   ClientLeft      =   2040
   ClientTop       =   1305
   ClientWidth     =   15240
   DrawStyle       =   5  'Transparent
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   601
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1016
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin XPFrame.FrameXp FrameXp1 
      Height          =   1020
      Left            =   120
      TabIndex        =   2
      Top             =   45
      Width           =   15015
      _ExtentX        =   26485
      _ExtentY        =   1799
      BackColor       =   16744576
      Caption         =   "DATOS "
      CaptionEstilo3D =   1
      BackColor       =   16744576
      ForeColor       =   8438015
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   1
      Begin VB.CommandButton Command2 
         Caption         =   "PROCESAR"
         Height          =   330
         Left            =   6720
         TabIndex        =   6
         Top             =   360
         Width           =   2220
      End
   End
   Begin XPFrame.FrameXp FrameXp2 
      Height          =   8160
      Left            =   120
      TabIndex        =   3
      Top             =   1080
      Width           =   15105
      _ExtentX        =   26644
      _ExtentY        =   14393
      BackColor       =   16761024
      Caption         =   "Listado de Pagos Duplicados"
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
      Alignment       =   1
      Begin VB.CommandButton Command1 
         Caption         =   "IMPRIMIR"
         Height          =   375
         Left            =   6600
         TabIndex        =   5
         Top             =   6840
         Width           =   2085
      End
      Begin FlexCell.Grid Grid1 
         Height          =   6420
         Left            =   90
         TabIndex        =   4
         Top             =   315
         Width           =   14880
         _ExtentX        =   26247
         _ExtentY        =   11324
         Cols            =   5
         DefaultFontSize =   8.25
         Rows            =   30
      End
   End
   Begin VB.PictureBox MANUAL 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H80000002&
      ForeColor       =   &H80000008&
      Height          =   0
      Left            =   0
      ScaleHeight     =   0
      ScaleWidth      =   15210
      TabIndex        =   1
      Top             =   9015
      Width           =   15240
   End
   Begin VB.TextBox pivote 
      Height          =   285
      Left            =   8415
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   4230
      Visible         =   0   'False
      Width           =   615
   End
End
Attribute VB_Name = "prove0013"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public saldoglobal As Double
Private localfiltro As String

Private MODIFI As Integer

'Private Sub codigo_Click()
'    Call dato1_KeyDown(vbKeyF2, 0)
'End Sub
 Private Sub imprimir()
If Grid1.Rows > 1 Then
Call Titulos("LISTADO DE PUBLICIDAD POR COBRAR")
Grid1.PageSetup.Orientation = cellLandscape
Grid1.PageSetup.HeaderMargin = 0.5
Grid1.PageSetup.PrintFixedRow = True
Grid1.PageSetup.TopMargin = 1.5
Grid1.PageSetup.LeftMargin = 0.1
Grid1.PageSetup.RightMargin = 0.1
Grid1.PageSetup.BottomMargin = 1.5
Grid1.PageSetup.FooterMargin = 0.5
Grid1.PageSetup.BlackAndWhite = True

Grid1.Range(0, 1, 0, Grid1.Cols - 1).Borders(cellEdgeTop) = cellThin
Grid1.Range(0, 1, 0, Grid1.Cols - 1).Borders(cellEdgeBottom) = cellThin
Grid1.Range(0, 1, 0, Grid1.Cols - 1).Borders(cellEdgeLeft) = cellThin
Grid1.Range(0, 1, 0, Grid1.Cols - 1).Borders(cellEdgeRight) = cellThin
Grid1.Range(0, 1, 0, Grid1.Cols - 1).Borders(cellInsideVertical) = cellThin
Grid1.Range(0, 1, 0, Grid1.Cols - 1).Borders(cellInsideHorizontal) = cellThin
Grid1.PrintPreview
End If
End Sub
Sub Titulos(titulo1)

    Dim i As Integer
    Dim objReportTitle As FlexCell.ReportTitle
    
    Grid1.FixedRowColStyle = Fixed3D
    Grid1.CellBorderColorFixed = vbButtonShadow
    Grid1.ShowResizeTips = False
    Grid1.ReportTitles.Clear
    Grid1.PageSetup.CenterHorizontally = True
    Grid1.PageSetup.Orientation = cellLandscape
    Grid1.PageSetup.PrintTitleRows = 0
    
    'ENCABEZADO DE PAGINA
    Grid1.PageSetup.Header = nombreempresa & vbCrLf & direccionempresa & vbCrLf & comunaempresa & vbCrLf & rutempresa
    Grid1.PageSetup.HeaderAlignment = CellLeft
    Grid1.PageSetup.HeaderFont.Name = "Verdana"
    Grid1.PageSetup.HeaderFont.Size = 8
    
    'TITULOS DEL REPORTE
    
    Set objReportTitle = New FlexCell.ReportTitle
    objReportTitle.text = titulo1 & "  |  " & "EMITIDO  :  " & Format(fechasistema, "dd-MM-yyyy")
    objReportTitle.Font.Name = "Verdana" '"Times New Roman"
    objReportTitle.Font.Size = 8
    objReportTitle.Font.Bold = True
    objReportTitle.Align = cellCenter
    objReportTitle.PrintOnAllPages = True
    Grid1.ReportTitles.Add objReportTitle
    

    
    'PIE DE PAGINA
    Grid1.PageSetup.Footer = "Pág &P de &N" & vbCrLf & "Fecha: &D" & vbCrLf & "Usuario: " & USUARIOSISTEMA
    Grid1.PageSetup.FooterAlignment = cellRight
    Grid1.PageSetup.FooterFont.Name = "Verdana"
    Grid1.PageSetup.FooterFont.Size = 7
    
End Sub


'Private Sub dato2_KeyDown(KeyCode As Integer, Shift As Integer)
'    If KeyCode = vbKeyF2 Then Call ayudactacte(dato2)
'    Call flechas(dato1, dato4, KeyCode)
'End Sub
 

Private Sub Command1_Click()
imprimir

End Sub

Private Sub COMMAND2_Click()
Call LEERGUIAS


End Sub


Private Sub Form_Load()
Call CENTRAR(Me)

    Call Conectar_BD
    sc = 0
  
Call CARGAPERMISO(Me.Name)
 
 CARGAGRILLADETALLE

End Sub

Sub flechas(ByRef caja As TextBox, ByRef caja1 As TextBox, ByRef codigo As Integer)
    If codigo = 38 And caja.Enabled = True Then caja.SetFocus
    If codigo = 40 And caja1.Enabled = True Then caja1.SetFocus
End Sub

Sub cargatexto(ByRef caja As TextBox)
caja.SelStart = 0: caja.SelLength = Len(caja.text)
End Sub




 
Sub CARGAGRILLADETALLE()
    Dim formatogrilla2(10, 10)
    formatogrilla2(1, 1) = "RUT"
    formatogrilla2(1, 2) = "NOMBRE"
    formatogrilla2(1, 3) = "TIPO"
    formatogrilla2(1, 4) = "NUMERO"
    formatogrilla2(1, 5) = "MONTO"
    formatogrilla2(1, 6) = "FECHA"
    
    Rem LARGO DE LOS DATOS
    formatogrilla2(2, 1) = "10"
    formatogrilla2(2, 2) = "30"
    formatogrilla2(2, 3) = "10"
    formatogrilla2(2, 4) = "10"
    formatogrilla2(2, 5) = "10"
    formatogrilla2(2, 6) = "10"
    
    Rem TIPO DE DATOS C=CEROS,S=STRING,N=NUMERICO,D=DATE,
    formatogrilla2(3, 1) = "S"
    formatogrilla2(3, 2) = "S"
    formatogrilla2(3, 3) = "S"
    formatogrilla2(3, 4) = "S"
    formatogrilla2(3, 5) = "N"
    formatogrilla2(3, 6) = "D"
    
    Rem FORMATO GRILLA
    formatogrilla2(4, 5) = " ###,###,##0"
    
    Rem LOCCKED
    formatogrilla2(5, 1) = "TRUE"
    formatogrilla2(5, 2) = "TRUE"
    formatogrilla2(5, 3) = "TRUE"
    formatogrilla2(5, 4) = "TRUE"
    formatogrilla2(5, 5) = "TRUE"
    formatogrilla2(5, 6) = "TRUE"
    formatogrilla2(5, 7) = "TRUE"
    formatogrilla2(5, 8) = "TRUE"
    formatogrilla2(5, 9) = "TRUE"
    
    
    Rem VALOR MAXIMO
    
    Grid1.Cols = 7
    Grid1.Rows = 1
    Grid1.AllowUserResizing = False
    Grid1.DisplayFocusRect = False
    Grid1.ExtendLastCol = True
    Grid1.BoldFixedCell = False
    Grid1.DrawMode = cellOwnerDraw
    Grid1.Appearance = Flat
    Grid1.ScrollBarStyle = Flat
    Grid1.FixedRowColStyle = Flat
'    Grid1.BackColorFixed = RGB(90, 158, 214)
'    Grid1.BackColorFixedSel = RGB(110, 180, 230)
'    Grid1.BackColorBkg = RGB(90, 158, 214)
'    Grid1.BackColorScrollBar = RGB(231, 235, 247)
'    Grid1.BackColor1 = RGB(231, 235, 247)
'    Grid1.BackColor2 = RGB(239, 243, 255)
'    Grid1.GridColor = RGB(148, 190, 231)
    Grid1.Column(0).Width = 0
    
    For k = 1 To Grid1.Cols - 1
        Grid1.Cell(0, k).text = formatogrilla2(1, k)
        Grid1.Column(k).Width = Val(formatogrilla2(2, k)) * 8
        Grid1.Column(k).MaxLength = Val(formatogrilla2(2, k))
        Grid1.Column(k).FormatString = formatogrilla2(4, k)
        Grid1.Column(k).Locked = formatogrilla2(5, k)
        If formatogrilla2(3, k) = "N" Then Grid1.Column(k).Alignment = cellRightCenter
        If formatogrilla2(3, k) = "S" Then Grid1.Column(k).Alignment = cellLeftTop
        If formatogrilla2(3, k) = "D" Then Grid1.Column(k).CellType = cellCalendar
    Next k
 
 
    End Sub
 
Sub verdetalle(loc, numero)
'Dim cSql As New rdoQuery
'Dim resultados As rdoResultset
'Dim tipo As String
'tipo = "DM"
'
'Set cSql.ActiveConnection = contadb
'
'cSql.sql = "select linea,codigo,descripcion,cantidad,uxc,unidades,precio,descuento,total "
'cSql.sql = cSql.sql & "from " & clientesistema & "gestion" & leerrubro(dato1.text) & ".l_movimientos_detalle_" & loc & " where tipo='" & tipo & "' and numero='" & numero & "' order by linea"
'cSql.Execute
'
'If cSql.RowsAffected > 0 Then
'    Grid1.Rows = cSql.RowsAffected + 1
'    Set resultados = cSql.OpenResultset
'
'    While Not resultados.EOF
'        Grid1.Cell(resultados(0), 1).text = resultados(1)
'        Grid1.Cell(resultados(0), 2).text = resultados(2)
'        Grid1.Cell(resultados(0), 3).text = resultados(3)
'        Grid1.Cell(resultados(0), 4).text = resultados(4)
'        Grid1.Cell(resultados(0), 5).text = resultados(5)
'        Grid1.Cell(resultados(0), 6).text = resultados(6)
'        Grid1.Cell(resultados(0), 7).text = resultados(7)
'        Grid1.Cell(resultados(0), 8).text = resultados(8)
'        resultados.MoveNext
'    Wend
'End If
'
'cSql.Close
'Set cSql = Nothing
'Set resultados = Nothing
 
End Sub
Function leerrubro(loc) As String
    Dim csql As New rdoQuery
    Dim resultado As rdoResultset
    
    Set csql.ActiveConnection = contadb
    csql.sql = "select rubro from " & clientesistema & "gestion.g_maestroempresas where "
    csql.sql = csql.sql & "codigo='" & loc & "' "
    csql.Execute
    
 If csql.RowsAffected > 0 Then
    Set resultado = csql.OpenResultset
    leerrubro = resultado(0)
 End If
 csql.Close
 Set csql = Nothing
 Set resultado = Nothing
 
End Function


Sub LEERGUIAS()
Dim csql As New rdoQuery
Dim resultados As rdoResultset
Dim tipo As String
Dim rutpaso As String
Dim totales(2) As Double
Dim totales2(2) As Double
Dim cuentapublicidad As String
Dim TABONO1 As Double
Dim TABONO2 As Double
Dim TOTALGE1 As Double
Dim TOTALGE2 As Double
TABONO1 = 0
TABONO2 = 0
TOTALGE1 = 0
TOTALGE2 = 0
totales(1) = 0
totales(2) = 0
totales2(1) = 0
totales2(2) = 0
tipo = "DM"
Rem cuentapublicidad = leerdatos(conta, "maestroempresas", "cuentapublicidad", "codigoempresa='" + empresaactiva + "' ")

Set csql.ActiveConnection = contadb
csql.sql = "SELECT rutctacte,'',tipodocumento,numerodocumento,monto,fecha,COUNT(tipodocumento+numerodocumento+rutctacte) AS re,tipodocumento,rutctacte,numerodocumento,glosacontable "
csql.sql = csql.sql + "FROM movimientoscontables WHERE codigocuenta='23100026' AND DH='D' AND (año>'2012') AND tipodocumento='FC' AND (tipo='CE' OR tipo='DB' OR tipo='PA') AND monto<>0 "
csql.sql = csql.sql + " GROUP BY tipodocumento,numerodocumento,rutctacte "
csql.sql = csql.sql + " HAVING COUNT(tipodocumento+numerodocumento+rutctacte)>1 "
csql.Execute
  Grid1.Rows = 1
  Grid1.AutoRedraw = False
  
If csql.RowsAffected > 0 Then
  
    Set resultados = csql.OpenResultset
    rutpaso = resultados(1)
    While Not resultados.EOF
        Grid1.Rows = Grid1.Rows + 1
        Grid1.Cell(Grid1.Rows - 1, 1).text = resultados(0)
        Grid1.Cell(Grid1.Rows - 1, 2).text = LEERNOMBREPROVEEDOR(resultados(0))
        Grid1.Cell(Grid1.Rows - 1, 3).text = resultados(2)
        Grid1.Cell(Grid1.Rows - 1, 4).text = resultados(3)
        Grid1.Cell(Grid1.Rows - 1, 5).text = resultados(4)
        Grid1.Cell(Grid1.Rows - 1, 6).text = resultados(5)
        
        
        resultados.MoveNext
    
    Wend
        Grid1.AutoRedraw = True
        Grid1.Refresh
        
End If

csql.Close
Set csql = Nothing
Set resultados = Nothing
 
End Sub


