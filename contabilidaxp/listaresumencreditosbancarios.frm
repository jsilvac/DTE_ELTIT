VERSION 5.00
Object = "{4777436C-EB8C-4596-98A8-EBCF98683969}#1.0#0"; "FlexCell.ocx"
Object = "{ADBBAED6-B16F-40DC-80DF-B44910CBA76C}#1.0#0"; "Frame-Xp.ocx"
Begin VB.Form prestamo04 
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
      Left            =   11880
      TabIndex        =   4
      Top             =   9480
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
      Begin VB.CommandButton botonmisaccesos 
         Caption         =   "Permisos Modulo"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   280
         Width           =   1455
      End
      Begin VB.CommandButton botonmisfavoritos 
         Caption         =   "Mis Favoritos"
         Height          =   255
         Left            =   1800
         TabIndex        =   5
         Top             =   280
         Width           =   1335
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "imprimir"
      Height          =   375
      Left            =   5535
      TabIndex        =   3
      Top             =   9585
      Width           =   2220
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
      Height          =   8745
      Left            =   45
      TabIndex        =   1
      Top             =   225
      Width           =   15180
      _ExtentX        =   26776
      _ExtentY        =   15425
      BackColor       =   16761024
      Caption         =   "Resumen de creditos"
      CaptionEstilo3D =   1
      BackColor       =   16761024
      ForeColor       =   65535
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin FlexCell.Grid Grid2 
         Height          =   8250
         Left            =   90
         TabIndex        =   2
         Top             =   270
         Width           =   15045
         _ExtentX        =   26538
         _ExtentY        =   14552
         Cols            =   5
         DefaultFontSize =   8.25
         Rows            =   30
      End
   End
End
Attribute VB_Name = "prestamo04"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public saldoglobal As Double
Private moneda As String
Private rutpropi As String
Private cuotaspagadas As Double

Private MODIFI As Integer


Private Sub Command1_Click()
imprimir

End Sub





Private Sub Form_Load()
Call CENTRAR(Me)
    Call Conectar_BD
    Rem Call Funciones_Forms_M_Productos.Conecta_Maestro_Productos
    sc = 0
    
Rem Call RECUPERAFECHA

Call CARGAPERMISO(Me.Name)
Call CARGAGRILLA2
leerCREDITOS

End Sub


Sub flechas(ByRef caja As TextBox, ByRef caja1 As TextBox, ByRef codigo As Integer)
    If codigo = 38 And caja.Enabled = True Then caja.SetFocus
    If codigo = 40 And caja1.Enabled = True Then caja1.SetFocus
End Sub

Sub CARGAGRILLA2()
    Dim formatogrilla2(10, 12)
    formatogrilla2(1, 1) = "BANCO"
    formatogrilla2(1, 2) = "TIPO"
    formatogrilla2(1, 3) = "NUMERO"
    formatogrilla2(1, 4) = "EMPRESA"
    formatogrilla2(1, 5) = "GLOSA  "
    formatogrilla2(1, 6) = "EMISION"
    formatogrilla2(1, 7) = "CAPITAL"
    formatogrilla2(1, 8) = "TIPO "
    formatogrilla2(1, 9) = "TOTAL CREDITO"
    formatogrilla2(1, 10) = "PAGADO"
    formatogrilla2(1, 11) = "SALDO"
    formatogrilla2(1, 12) = "CUO/PAG"
    
    Rem LARGO DE LOS DATOS
    formatogrilla2(2, 1) = "10"
    formatogrilla2(2, 2) = "10"
    formatogrilla2(2, 3) = "7"
    formatogrilla2(2, 4) = "10"
    formatogrilla2(2, 5) = "10"
    formatogrilla2(2, 6) = "7"
    formatogrilla2(2, 7) = "10"
    formatogrilla2(2, 8) = "7"
    formatogrilla2(2, 9) = "10"
    formatogrilla2(2, 10) = "10"
    formatogrilla2(2, 11) = "10"
    formatogrilla2(2, 12) = "10"
    
    Rem TIPO DE DATOS C=CEROS,S=STRING,N=NUMERICO,D=DATE,
    formatogrilla2(3, 1) = "S"
    formatogrilla2(3, 2) = "S"
    formatogrilla2(3, 3) = "S"
    formatogrilla2(3, 4) = "S"
    formatogrilla2(3, 5) = "S"
    formatogrilla2(3, 6) = "S"
    formatogrilla2(3, 7) = "N"
    formatogrilla2(3, 8) = "S"
    formatogrilla2(3, 9) = "N"
    formatogrilla2(3, 10) = "N"
    formatogrilla2(3, 11) = "N"
    formatogrilla2(3, 12) = "S"
    
    Rem FORMATO GRILLA
    formatogrilla2(4, 7) = " ###,###,##0"
    
    formatogrilla2(4, 9) = " ###,###,##0.00"
    formatogrilla2(4, 10) = " ###,###,##0.00"
    formatogrilla2(4, 11) = " ###,###,##0.00"
    
    
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
    formatogrilla2(5, 12) = "TRUE"
    
    
    Rem VALOR MAXIMO
    
    Grid2.Cols = 13
    Grid2.Rows = 1
    Grid2.AllowUserResizing = True
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
    Grid2.BackColor1 = RGB(231, 235, 247)
    Grid2.BackColor2 = RGB(239, 243, 255)
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
   
  
    
    
    
    End Sub


 Public Sub leerCREDITOS()
 Dim csql As New rdoQuery
 Dim resultados As rdoResultset
 CARGAGRILLA2
 Dim monto As Double
 
 Set csql.ActiveConnection = contadb
 csql.sql = "select cb.banco,cb.tipo,cb.numero,cb.empresa,cb.glosa,cb.fecha,cb.capital,cb.moneda,cb.cuotas*cb.monto,cb.cuotas from " & clientesistema & "creditos_bancarios" & ".maestro_compromisos as cb "
  csql.sql = csql.sql + "order by cb.empresa "
 csql.Execute
 Grid2.Rows = 1
 If csql.RowsAffected > 0 Then
    Set resultados = csql.OpenResultset
    While resultados.EOF = False
    Grid2.Rows = Grid2.Rows + 1
    Grid2.Cell(Grid2.Rows - 1, 1).text = resultados(0) + "=" + leerbanco(resultados(0))
    Grid2.Cell(Grid2.Rows - 1, 2).text = resultados(1) + "=" + leertipocredito(resultados(1))
    Grid2.Cell(Grid2.Rows - 1, 3).text = resultados(2)
    Grid2.Cell(Grid2.Rows - 1, 4).text = resultados(3) + "=" + leerempresa(resultados(3))
    
    Grid2.Cell(Grid2.Rows - 1, 5).text = resultados(4)
    Grid2.Cell(Grid2.Rows - 1, 6).text = resultados(5)
    Grid2.Cell(Grid2.Rows - 1, 7).text = resultados(6)
    Grid2.Cell(Grid2.Rows - 1, 8).text = resultados(7) + "=" + leertipoMONEDA(resultados(7))
    Grid2.Cell(Grid2.Rows - 1, 9).text = resultados(8)
    monto = leerpagado(resultados(0), resultados(1), resultados(2), resultados(3))
    
    Grid2.Cell(Grid2.Rows - 1, 10).text = monto
    
    Grid2.Cell(Grid2.Rows - 1, 11).text = resultados(8) - monto
    Grid2.Cell(Grid2.Rows - 1, 12).text = Str(cuotaspagadas) & "/" & resultados(9)
    

'    If Format(resultados(6), "yyyy-mm-dd") < Format(fechasistema, "yyyy-mm-dd") Then
'    Grid2.Range(Grid2.Rows - 1, 1, Grid2.Rows - 1, 11).BackColor = &HFF&
    
'    End If
    
    
    
    
    resultados.MoveNext
    
    
    
    Wend
    
    
  End If
 csql.Close
 Set csql = Nothing
 Set resultados = Nothing
 
 End Sub
Public Function leerpagado(banco, tipo, numero, empresa) As Double

 Dim csql As New rdoQuery
 Dim resultados As rdoResultset
 
 Set csql.ActiveConnection = contadb
 csql.sql = "select IFNULL(sum(monto),0),count(monto) from " & clientesistema & "creditos_bancarios" & ".creditos_vencimientos as mp where tipo='" + tipo + "' and banco='" + banco + "' and empresa='" + empresa + "' and numero='" + numero + "' "
 csql.sql = csql.sql + " and pagado='1' "
 csql.sql = csql.sql + "order by fecha "
 
 csql.Execute
 cuotaspagadas = 0
 If csql.RowsAffected > 0 Then
    Set resultados = csql.OpenResultset
 leerpagado = resultados(0)
 cuotaspagadas = resultados(1)
 Else
 leerpagado = 0
 
  End If
 csql.Close
 Set csql = Nothing
 Set resultados = Nothing
 
 End Function

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




Sub imprimir()
Dim titulo As String

titulo = "LISTADO DE COMPROMISOS BANCARIOS"

Call CABEZAS2(titulo, "N", 1)
Grid2.DefaultFont.Size = 8
Grid2.PageSetup.Orientation = cellLandscape

Grid2.Range(0, 1, 0, Grid2.Cols - 1).Borders(cellEdgeBottom) = cellThick
Grid2.Range(0, 1, 0, Grid2.Cols - 1).Borders(cellEdgeLeft) = cellThick
Grid2.Range(0, 1, 0, Grid2.Cols - 1).Borders(cellEdgeTop) = cellThick
Grid2.Range(0, 1, 0, Grid2.Cols - 1).Borders(cellEdgeRight) = cellThick
Grid2.Range(0, 1, 0, Grid2.Cols - 1).Borders(cellInsideHorizontal) = cellThick
Grid2.Range(0, 1, 0, Grid2.Cols - 1).Borders(cellInsideVertical) = cellThick



Grid2.PageSetup.PrintFixedRow = True
Grid2.PageSetup.BottomMargin = 1
Grid2.PageSetup.TopMargin = 1
Grid2.PageSetup.LeftMargin = 1
Grid2.PageSetup.RightMargin = 0
Grid2.PageSetup.BlackAndWhite = True
Grid2.PageSetup.PrintGridlines = False
Grid2.PrintPreview 100

   
End Sub

Sub CABEZAS2(titulo, tipo, FOLIO)
Dim objReportTitle As FlexCell.ReportTitle
Grid2.ReportTitles.Clear


Set objReportTitle = New FlexCell.ReportTitle
    objReportTitle.text = titulo
    objReportTitle.Font.Name = "arial"
    objReportTitle.Font.Size = 12
    objReportTitle.PrintOnAllPages = True
    Grid2.ReportTitles.Add objReportTitle

Set objReportTitle = New FlexCell.ReportTitle
    objReportTitle.text = ""
    
    objReportTitle.Font.Name = "arial"
    objReportTitle.Font.Size = 12
    objReportTitle.PrintOnAllPages = True
    Grid2.ReportTitles.Add objReportTitle
    
    'Report Title 1
    If tipo = "N" Then
        For k = 1 To 5
        Set objReportTitle = New FlexCell.ReportTitle
        objReportTitle.text = DATOSEMPRESA(k)
        objReportTitle.Font.Name = "arial"
        objReportTitle.Font.Size = 8
        objReportTitle.Font.Italic = True
        objReportTitle.PrintOnAllPages = True
        'objReportTitle.Color = RGB(128, 0, 0)
        objReportTitle.Align = CellLeft
        Grid2.ReportTitles.Add objReportTitle
    Next k
    Else
        For k = 1 To 4
        Set objReportTitle = New FlexCell.ReportTitle
        objReportTitle.text = ""
        objReportTitle.Font.Name = "arial"
        objReportTitle.Font.Size = 8
        objReportTitle.Font.Italic = True
        objReportTitle.PrintOnAllPages = True
        'objReportTitle.Color = RGB(128, 0, 0)
        objReportTitle.Align = CellLeft
        Grid2.ReportTitles.Add objReportTitle
        
        Next k
    Set objReportTitle = New FlexCell.ReportTitle
        
        
        
        
        
        objReportTitle.text = ""
        objReportTitle.Font.Name = "arial"
        objReportTitle.Font.Size = 8
        objReportTitle.Font.Italic = True
        objReportTitle.PrintOnAllPages = True
        'objReportTitle.Color = RGB(128, 0, 0)
        objReportTitle.Align = CellLeft
        Grid2.ReportTitles.Add objReportTitle
        
    End If
    
With Grid2.PageSetup
        
        If tipo = "N" Then .Header = "Pagina &P de &N Emitido: &D Usuario:" + USUARIOSISTEMA
        Rem If TIPO = "S" Then .Footer = "pagina &P"
        
        .HeaderAlignment = cellCenter
        
        .HeaderFont.Name = "Verdana"
        .HeaderFont.Size = 7
        .HeaderMargin = 2
        .TopMargin = 1
        .BottomMargin = 2
        
        
        
End With

End Sub

Private Sub botonmisaccesos_Click()
programafiltro = Me.Caption
misaccesos.Show
End Sub


Private Sub botonmisfavoritos_Click()
Call AgregaFavoritos(USUARIOSISTEMA, App.EXEName, Me.Name, Me.Caption)
End Sub

