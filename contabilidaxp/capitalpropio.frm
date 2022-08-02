VERSION 5.00
Object = "{4777436C-EB8C-4596-98A8-EBCF98683969}#1.0#0"; "FlexCell.ocx"
Object = "{ADBBAED6-B16F-40DC-80DF-B44910CBA76C}#1.0#0"; "Frame-Xp.ocx"
Begin VB.Form auxiliar08 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Contratos De Arriendo De Propiedades"
   ClientHeight    =   10275
   ClientLeft      =   2040
   ClientTop       =   1425
   ClientWidth     =   15240
   DrawStyle       =   5  'Transparent
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   685
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1016
   StartUpPosition =   2  'CenterScreen
   Begin XPFrame.FrameXp FrameQuickMenu 
      Height          =   615
      Left            =   11760
      TabIndex        =   4
      Top             =   9360
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   1085
      BackColor       =   16744576
      Caption         =   " Mis Datos"
      BackColor       =   16744576
      BordeColor      =   4194304
      ColorBarraArriba=   4194304
      ColorBarraAbajo =   4194304
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   1
      Begin VB.CommandButton botonmisfavoritos 
         Caption         =   "Mis Favoritos"
         Height          =   255
         Left            =   1800
         TabIndex        =   6
         Top             =   280
         Width           =   1335
      End
      Begin VB.CommandButton botonmisaccesos 
         Caption         =   "Permisos Modulo"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   280
         Width           =   1455
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
      TabIndex        =   0
      Top             =   10275
      Width           =   15240
   End
   Begin XPFrame.FrameXp FrameXp3 
      Height          =   9915
      Left            =   180
      TabIndex        =   1
      Top             =   120
      Width           =   14910
      _ExtentX        =   26300
      _ExtentY        =   17489
      BackColor       =   16761024
      Caption         =   "Propiedades en Arriendo"
      CaptionEstilo3D =   1
      BackColor       =   16761024
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
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FF8080&
         Caption         =   "Imprimir"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5670
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   9315
         Width           =   2040
      End
      Begin FlexCell.Grid Grid2 
         Height          =   8835
         Left            =   90
         TabIndex        =   2
         Top             =   270
         Width           =   14775
         _ExtentX        =   26061
         _ExtentY        =   15584
         Cols            =   5
         DefaultFontSize =   8.25
         Rows            =   30
      End
   End
End
Attribute VB_Name = "auxiliar08"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public saldoglobal As Double
Private moneda As String
Private rutpropi As String
Private MODIFI As Integer


Private Sub Command1_Click()
    Call Titulos
    Grid2.PrintPreview
    
End Sub
Sub Titulos()
    

    Dim i As Integer
    Dim objReportTitle As FlexCell.ReportTitle
    
    Grid2.FixedRowColStyle = Fixed3D
    Grid2.CellBorderColorFixed = vbButtonShadow
    Grid2.ShowResizeTips = False
    Grid2.PageSetup.Orientation = cellPortrait
    Grid2.DefaultFont.Size = 7.5
    Grid2.Column(1).Width = 50
    Grid2.Column(2).Width = 150
    Grid2.Column(3).Width = 170
    Grid2.Column(4).Width = 60
    Grid2.Column(5).Width = 280
    Grid2.Column(6).Width = 55
    Grid2.Column(7).Width = 55
    Grid2.Column(8).Width = 70
    Grid2.Column(9).Width = 55
    Grid2.Column(10).Width = 55
    Grid2.Column(11).Width = 55
    Grid2.PageSetup.PrintFixedRow = True
    Grid2.ReportTitles.Clear
    Grid2.PageSetup.CenterHorizontally = True
    Grid2.PageSetup.PrintTitleRows = 0
    Grid2.PageSetup.BlackAndWhite = False
    Grid2.PageSetup.Orientation = cellLandscape
    'Logo
  
    'ENCABEZADO DE PAGINA
    Grid2.PageSetup.Header = nombreempresa & vbCrLf & direccionempresa & vbCrLf & comunaempresa
    Grid2.PageSetup.HeaderAlignment = CellLeft
    Grid2.PageSetup.HeaderFont.Name = "Verdana"
    Grid2.PageSetup.HeaderFont.Size = 8
    'TITULOS DEL REPORTE
    Set objReportTitle = New FlexCell.ReportTitle
    objReportTitle.text = "LISTADO DE ARRIENDOS Y SUS ESTADOS"
    objReportTitle.Font.Name = "Verdana" '"Times New Roman"
    objReportTitle.Font.Size = 10
    objReportTitle.Font.Bold = True
    objReportTitle.Align = cellCenter
    objReportTitle.PrintOnAllPages = True
    Grid2.ReportTitles.Add objReportTitle
        
    Set objReportTitle = New FlexCell.ReportTitle
    objReportTitle.Font.Name = "Verdana" '"Times New Roman"
    objReportTitle.Font.Size = 7
    objReportTitle.Font.Underline = True
    objReportTitle.Font.Bold = True
    objReportTitle.Align = cellCenter
    objReportTitle.PrintOnAllPages = True
    Grid2.ReportTitles.Add objReportTitle
    
    
    'PIE DE PAGINA
    Grid2.PageSetup.Footer = "Pág &P de &N" & vbCrLf & "Fecha: &D " & vbCrLf & "Usuario:" & USUARIOSISTEMA
    Grid2.PageSetup.FooterAlignment = cellRight
    Grid2.PageSetup.FooterFont.Name = "Verdana"
    Grid2.PageSetup.FooterFont.Size = 7
    Grid2.PageSetup.LeftMargin = 0.5
    Grid2.PageSetup.RightMargin = 0.5
    
    Grid2.Range(0, 1, 0, Grid2.Cols - 1).Borders(cellEdgeLeft) = cellThick
    Grid2.Range(0, 1, 0, Grid2.Cols - 1).Borders(cellEdgeTop) = cellThick
    Grid2.Range(0, 1, 0, Grid2.Cols - 1).Borders(cellEdgeBottom) = cellThick
    Grid2.Range(0, 1, 0, Grid2.Cols - 1).Borders(cellEdgeRight) = cellThick
    Grid2.Range(0, 1, 0, Grid2.Cols - 1).Borders(cellInsideHorizontal) = cellThick
    Grid2.Range(0, 1, 0, Grid2.Cols - 1).Borders(cellInsideVertical) = cellThick
    
    
    
    
    
End Sub

Private Sub Form_Load()
    Call CENTRAR(Me)
    Call Conectar_BD
    Rem Call Funciones_Forms_M_Productos.Conecta_Maestro_Productos
    sc = 0
    Rem Call RECUPERAFECHA
    Call CARGAPERMISO(Me.Name)
    Call CARGAGRILLA2
End Sub

Sub flechas(ByRef caja As TextBox, ByRef caja1 As TextBox, ByRef codigo As Integer)
    If codigo = 38 And caja.Enabled = True Then caja.SetFocus
    If codigo = 40 And caja1.Enabled = True Then caja1.SetFocus
End Sub

Sub cargatexto(ByRef caja As TextBox)
    caja.SelStart = 0: caja.SelLength = Len(caja.text)
End Sub

Private Sub opciones_GotFocus()
    MANUAL.SetFocus
End Sub

 Private Function leearrendador(rutarrendador) As String
    Dim csql As New rdoQuery
    Dim resultados As rdoResultset
 
    Set csql.ActiveConnection = contadb
    csql.sql = "select nombre from " & clientesistema & "arriendos" & ".maestro_arrendadores "
    csql.sql = csql.sql & "where rut='" & rutarrendador & "' "
    csql.Execute
    leearrendador = ""
 
    If csql.RowsAffected > 0 Then
        Set resultados = csql.OpenResultset
        leearrendador = resultados(0)
    Else
        leearrendador = ""
    End If
 
    csql.Close
    Set csql = Nothing
    Set resultados = Nothing
 
 End Function
 Private Function leearrendatario(rutarrendatario) As String
    Dim csql As New rdoQuery
    Dim resultados As rdoResultset
 
    Set csql.ActiveConnection = contadb
    csql.sql = "select nombre from " & clientesistema & "arriendos" & ".maestro_arrendatarios "
    csql.sql = csql.sql & "where rut='" & rutarrendatario & "' "
    csql.Execute
    leearrendatario = ""
 
    If csql.RowsAffected > 0 Then
        Set resultados = csql.OpenResultset
        leearrendatario = resultados(0)
    
    Else
        leearrendatario = ""
    End If
 
    csql.Close
    Set csql = Nothing
    Set resultados = Nothing
 
 End Function
 
 Private Function leepropiedad(codigopropiedad) As String
    Dim csql As New rdoQuery
    Dim resultados As rdoResultset
 
    Set csql.ActiveConnection = contadb
    csql.sql = "select direccion,monedaarriendo,rutpropietario from " & clientesistema & "arriendos" & ".maestro_propiedades "
    csql.sql = csql.sql & "where codigopropiedad='" & codigopropiedad & "' "
    csql.Execute
    leepropiedad = ""
    moneda = ""
    If csql.RowsAffected > 0 Then
        Set resultados = csql.OpenResultset
        leepropiedad = resultados(0)
        moneda = resultados(1)
        rutpropi = resultados(2)
    Else
    leepropiedad = ""

    End If
 
    csql.Close
    Set csql = Nothing
    Set resultados = Nothing
 
 End Function


Sub CARGAGRILLA2()
    Dim formatogrilla2(10, 12)
    formatogrilla2(1, 1) = "CODIGO"
    formatogrilla2(1, 2) = "PROPIEDAD"
    formatogrilla2(1, 3) = "DIRECCION"
    formatogrilla2(1, 4) = "CONTRATO"
    formatogrilla2(1, 5) = "ARRENDATARIO"
    formatogrilla2(1, 6) = "DESDE"
    formatogrilla2(1, 7) = "HASTA"
    formatogrilla2(1, 8) = "MONTO"
    formatogrilla2(1, 9) = "MONEDA"
    formatogrilla2(1, 10) = "G/COMUNES"
    formatogrilla2(1, 11) = "MOROSO"
    
    Rem LARGO DE LOS DATOS
    formatogrilla2(2, 1) = "5"
    formatogrilla2(2, 2) = "10"
    formatogrilla2(2, 3) = "20"
    formatogrilla2(2, 4) = "8"
    formatogrilla2(2, 5) = "20"
    formatogrilla2(2, 6) = "7"
    formatogrilla2(2, 7) = "7"
    formatogrilla2(2, 8) = "10"
    formatogrilla2(2, 9) = "5"
    formatogrilla2(2, 10) = "8"
    formatogrilla2(2, 11) = "5"
    
    Rem TIPO DE DATOS C=CEROS,S=STRING,N=NUMERICO,D=DATE,
    formatogrilla2(3, 1) = "N"
    formatogrilla2(3, 2) = "S"
    formatogrilla2(3, 3) = "S"
    formatogrilla2(3, 4) = "N"
    formatogrilla2(3, 5) = "S"
    formatogrilla2(3, 6) = "D"
    formatogrilla2(3, 7) = "D"
    formatogrilla2(3, 8) = "N"
    formatogrilla2(3, 9) = "N"
    formatogrilla2(3, 10) = "N"
    formatogrilla2(3, 11) = "N"
    
    Rem FORMATO GRILLA
    
    formatogrilla2(4, 8) = " ###,###,##0.00"
    formatogrilla2(4, 10) = " ###,###,##0.00"
    
    
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
    formatogrilla2(5, 10) = "TRUE"
    formatogrilla2(5, 11) = "TRUE"
    
    
    Rem VALOR MAXIMO
    
    Grid2.Cols = 12
    Grid2.Rows = 1
    Grid2.AllowUserResizing = False
    Grid2.DisplayFocusRect = False
    Grid2.ExtendLastCol = True
    Grid2.BoldFixedCell = False
    Grid2.DrawMode = cellOwnerDraw
    Grid2.Appearance = Flat
    Grid2.ScrollBarStyle = Flat
    Grid2.FixedRowColStyle = Flat
    Grid2.BackColorFixed = RGB(90, 158, 214)
    Grid2.BackColorFixedSel = RGB(110, 180, 230)
    Grid2.BackColorBkg = RGB(90, 158, 214)
    Grid2.BackColorScrollBar = RGB(231, 235, 247)
'    Grid2.BackColor1 = RGB(231, 235, 247)
'    Grid2.BackColor2 = RGB(239, 243, 255)
    Grid2.GridColor = RGB(148, 190, 231)
    Grid2.Column(0).Width = 0
    
    For k = 1 To Grid2.Cols - 1
        Grid2.Cell(0, k).text = formatogrilla2(1, k)
        Grid2.Column(k).Width = Val(formatogrilla2(2, k)) * 9
        Grid2.Column(k).MaxLength = Val(formatogrilla2(2, k))
        Grid2.Column(k).FormatString = formatogrilla2(4, k)
        Grid2.Column(k).Locked = formatogrilla2(5, k)
        If formatogrilla2(3, k) = "N" Then Grid2.Column(k).Alignment = cellRightCenter
        If formatogrilla2(3, k) = "S" Then Grid2.Column(k).Alignment = cellLeftTop
        
        
        If formatogrilla2(3, k) = "D" Then Grid2.Column(k).CellType = cellCalendar
        
    Next k
    Grid2.Column(11).CellType = cellCheckBox
     
    
    leerpropiedades
    
    End Sub


 Public Sub leerpropiedades()
 Dim csql As New rdoQuery
 Dim resultados As rdoResultset
 
 Set csql.ActiveConnection = contadb
 csql.sql = "select mp.codigopropiedad,mp.nombrepropiedad,mp.direccion, "
 csql.sql = csql.sql & "ca.numero,ca.rutarrendatario,ca.fechainicio,max(ca.fechatermino),ca.montoarriendo, "
 csql.sql = csql.sql & "ca.monedaarriendo,ca.gastoscomunes  from " & clientesistema & "arriendos "
 csql.sql = csql.sql & ".maestro_propiedades as mp left join " + clientesistema + "arriendos "
 csql.sql = csql.sql & ".contratos_arriendo as ca on (mp.codigopropiedad = ca.propiedad) group by ca.rutarrendatario,mp.codigopropiedad  order by mp.direccion "
 csql.Execute
 
 Grid2.Rows = 1
 If csql.RowsAffected > 0 Then
    Set resultados = csql.OpenResultset
    While resultados.EOF = False
    Grid2.Rows = Grid2.Rows + 1
    For k = 1 To 3
    Grid2.Cell(Grid2.Rows - 1, k).text = resultados(k - 1)
    Next k
    If IsNull(resultados(3)) = False Then
    Grid2.Cell(Grid2.Rows - 1, 4).text = resultados(3)
    Grid2.Cell(Grid2.Rows - 1, 5).text = leearrendatario(resultados(4))
    Grid2.Cell(Grid2.Rows - 1, 6).text = resultados(5)
    Grid2.Cell(Grid2.Rows - 1, 7).text = resultados(6)
    Grid2.Cell(Grid2.Rows - 1, 8).text = resultados(7)
    Grid2.Cell(Grid2.Rows - 1, 9).text = leemonedas(resultados(8))
    Grid2.Cell(Grid2.Rows - 1, 10).text = resultados(9)
    Grid2.Cell(Grid2.Rows - 1, 11).text = arriendoatrasado(resultados(3), Format(fechasistema, "yyyy-mm-dd"))
    If Format(resultados(6), "yyyy-mm-dd") < Format(fechasistema, "yyyy-mm-dd") Then
    Grid2.Range(Grid2.Rows - 1, 1, Grid2.Rows - 1, 11).BackColor = &HFF&
    
    End If
    
    
    
    Else
    
    Grid2.Cell(Grid2.Rows - 1, 5).text = "*** DISPONIBLE **"
    
    End If
    
    resultados.MoveNext
    
    
    
    Wend
    
     Grid2.Rows = Grid2.Rows + 2
    Grid2.Range(Grid2.Rows - 1, 1, Grid2.Rows - 1, 1).BackColor = &HFF&
    Grid2.Cell(Grid2.Rows - 1, 2).text = "* CONTRATOS VENCIDOS"
  End If
 csql.Close
 Set csql = Nothing
 Set resultados = Nothing

 End Sub

Private Function leemonedas(codigo) As String
Dim csql As New rdoQuery
Dim resultados As rdoResultset

Set csql.ActiveConnection = contadb

csql.sql = "select nombremoneda from " & clientesistema & "arriendos" & ".maestro_monedas where codigomoneda='" & codigo & "'"
csql.Execute
leemonedas = ""
If csql.RowsAffected > 0 Then
Set resultados = csql.OpenResultset
leemonedas = resultados(0)
End If
Set resultados = Nothing
csql.Close
Set csql = Nothing

End Function

Public Function LEERULTIMOFOLIOcontrato() As String

    Dim resultados As rdoResultset
    Dim csql As New rdoQuery
    
        Set csql.ActiveConnection = contadb

            csql.sql = "select IFNULL(max(numero),0) from " + clientesistema + "arriendos.contratos_arriendo"
            
            csql.Execute
    If csql.RowsAffected > 0 Then
    Set resultados = csql.OpenResultset
    
    
        LEERULTIMOFOLIOcontrato = Format(resultados(0) + 1, "0000000000")
    End If
    
End Function

Private Sub botonmisaccesos_Click()
programafiltro = Me.Caption
misaccesos.Show
End Sub


Private Sub botonmisfavoritos_Click()
Call AgregaFavoritos(USUARIOSISTEMA, App.EXEName, Me.Name, Me.Caption)
End Sub
