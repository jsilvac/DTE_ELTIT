VERSION 5.00
Object = "{4777436C-EB8C-4596-98A8-EBCF98683969}#1.0#0"; "FlexCell.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form auxiliar031 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Libro Mayor Analitico"
   ClientHeight    =   10230
   ClientLeft      =   435
   ClientTop       =   825
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10230
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ProgressBar barra 
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   9000
      Visible         =   0   'False
      Width           =   15015
      _ExtentX        =   26485
      _ExtentY        =   873
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   1
      Min             =   1
      Max             =   5000
      Scrolling       =   1
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Imprime Formato Grande"
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   9720
      Width           =   2055
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Exportar Excel"
      Height          =   375
      Left            =   2520
      TabIndex        =   2
      Top             =   9720
      Width           =   2055
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Exportar Html"
      Height          =   375
      Left            =   4800
      TabIndex        =   1
      Top             =   9720
      Width           =   2055
   End
   Begin FlexCell.Grid Grid1 
      Height          =   9495
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   15015
      _ExtentX        =   26485
      _ExtentY        =   16748
      Cols            =   5
      DefaultFontSize =   8.25
      Rows            =   30
   End
End
Attribute VB_Name = "auxiliar031"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private formatogrilla(20, 20)
Private lin As Double
Private saldo As Double


Private Sub busca_Click()

End Sub

Private Sub Command1_Click()


 Grid1.DefaultFont.Size = 6.5
For K = 1 To 15 - 1
Grid1.Column(K).Width = Val(formatogrilla(2, K)) * Grid1.DefaultFont.Size
Next K
Grid1.PageSetup.Orientation = cellPortrait

Grid1.PageSetup.PrintFixedRow = True


'Grid1.PageSetup.BlackAndWhite = True
Grid1.PageSetup.BottomMargin = 1
Grid1.PageSetup.TopMargin = 1
Grid1.PageSetup.LeftMargin = 1
Grid1.PageSetup.RightMargin = 0


CABEZA




Grid1.PrintPreview 75


End Sub
Sub CABEZA()
Dim objReportTitle As FlexCell.ReportTitle
Grid1.ReportTitles.Clear


Set objReportTitle = New FlexCell.ReportTitle
    objReportTitle.text = "Libro Diario"
    objReportTitle.Font.Name = "verdana"
    objReportTitle.Font.Size = 18
    objReportTitle.PrintOnAllPages = True
    Grid1.ReportTitles.Add objReportTitle
    
    'Report Title 1
    For K = 1 To 5
    Set objReportTitle = New FlexCell.ReportTitle
    objReportTitle.text = DATOSEMPRESA(K)
    objReportTitle.Font.Name = "verdana"
    objReportTitle.Font.Size = 8
    objReportTitle.Font.Italic = True
    objReportTitle.PrintOnAllPages = True
    objReportTitle.Color = RGB(128, 0, 0)
    objReportTitle.Align = CellLeft
    Grid1.ReportTitles.Add objReportTitle
    Next K
With Grid1.PageSetup
        
        .Footer = "Pagina &P de &N Emitido: &D Usuario:" + USUARIOSISTEMA
        .FooterAlignment = cellRight
        .FooterFont.Name = "Verdana"
        .FooterFont.Size = 7
        .FooterMargin = 0.1
        
End With

End Sub

Private Sub Command3_Click()
Grid1.ExportToExcel ("")


End Sub


Private Sub Form_Load()
    
    Call Conectar_BD
    Call Conectarconta(servidor, "conta", USUARIO, password)

CARGAGRILLA
leecuentas

End Sub

    
Sub LEERMOVIMIENTOS(cuenta, nombre)
Dim resultados As rdoResultset
    Dim cSql As New rdoQuery
    Dim rut As String
    Dim PASO As String
        Set cSql.ActiveConnection = db
        cSql.SQL = "SELECT fecha,tipo,numero,linea,codigocuenta,glosacontable,tipodocumento,numerodocumento,fechadocumento,fechavencimiento,monto,dh "
        cSql.SQL = cSql.SQL + "FROM movimientoscontables where codigocuenta='" + cuenta + "'"
        cSql.SQL = cSql.SQL + "order by fecha"
        cSql.Execute
       
        lin = lin + 1
        Grid1.Rows = Grid1.Rows + 1
        Call DATOSSALDOS(cuenta)
        For K = 1 To 6
        Grid1.Column(K).Locked = False
        Next K
        
        Grid1.Range(lin, 1, lin, 13).FontBold = True
        Grid1.Range(lin, 1, lin, 13).FontUnderline = True
        
        
        
        
        Grid1.Range(lin, 1, lin, 6).Merge
        
        Grid1.Cell(lin, 1).CellType = cellTextBox
        
        Grid1.Cell(lin, 10).CellType = cellTextBox
        
        Grid1.Cell(lin, 1).text = nombre
        Grid1.Cell(lin, 10).text = "SALDO-->"
        
        Grid1.Cell(lin, 13).text = saldo

        
        If cSql.RowsAffected > 0 Then
        Set resultados = cSql.OpenResultset
        
         While Not resultados.EOF
          lin = lin + 1
             Grid1.Rows = Grid1.Rows + 1
             For K = 0 To 9
             Grid1.Cell(lin, K + 1).text = resultados(K)
             Next K
             If resultados(11) = "D" Then Grid1.Cell(lin, 11).text = resultados(10): anted = anted + resultados(10): saldo = saldo + resultados(10)
             If resultados(11) = "H" Then Grid1.Cell(lin, 12).text = resultados(10): anteh = anteh + resultados(10): saldo = saldo - resultados(10)
             Grid1.Cell(lin, 13).text = saldo
        
             resultados.MoveNext
           
         Wend
          lin = lin + 1
             Grid1.Rows = Grid1.Rows + 1
         
         Call totalcomprobante(lin)
          resultados.Close
            Set resultados = Nothing

        End If

End Sub

Sub totalcomprobante(Row)
    Grid1.Range(Row, 1, Row, 12).FontBold = True
    Grid1.Range(Row, 1, Row, 12).FontUnderline = True
        
    
    Grid1.Range(Row, 11, Row, 12).Borders(cellEdgeTop) = cellThin
    Grid1.Cell(Row, 10).CellType = cellTextBox
    Grid1.Cell(Row, 10).text = "TOTAL "
    Grid1.Cell(Row, 11).text = anted
    Grid1.Cell(Row, 12).text = anteh
    lin = lin + 2
             Grid1.Rows = Grid1.Rows + 2
        
        anted = 0: anteh = 0: saldo = 0
    End Sub
    





Sub CARGAGRILLA()
Rem DATOS DE LA COLUMNA
    Grid1.DefaultFont.Size = 7.5
    
    
    formatogrilla(1, 1) = "FECHA"
    formatogrilla(1, 2) = "TP"
    formatogrilla(1, 3) = "NUMERO"
    formatogrilla(1, 4) = "LINEA"
    formatogrilla(1, 5) = "CUENTA"
    formatogrilla(1, 6) = "GLOSA"
    formatogrilla(1, 7) = "TP"
    formatogrilla(1, 8) = "NUMERO"
    formatogrilla(1, 9) = "EMISION"
    formatogrilla(1, 10) = "VENCIMIENTO"
    formatogrilla(1, 11) = "DEBE"
    formatogrilla(1, 12) = "HABER"
    formatogrilla(1, 13) = "SALDO"
    formatogrilla(1, 14) = "NOMBRE CUENTA"
    formatogrilla(1, 15) = "CUENTA CORRIENTE"
    formatogrilla(1, 16) = "CRCC"
     
    Rem LARGO DE LOS DATOS
    
    formatogrilla(2, 1) = "10"
    formatogrilla(2, 3) = "10"
    formatogrilla(2, 4) = "3"
    formatogrilla(2, 5) = "10"
    formatogrilla(2, 6) = "30"
    formatogrilla(2, 7) = "2"
    formatogrilla(2, 8) = "10"
    formatogrilla(2, 9) = "10"
    formatogrilla(2, 10) = "10"
    formatogrilla(2, 11) = "12"
    formatogrilla(2, 12) = "12"
    formatogrilla(2, 13) = "12"
    formatogrilla(2, 14) = "30"
    formatogrilla(2, 15) = "30"

    
    Rem TIPO DE DATOS C=CEROS,S=STRING,N=NUMERICO,D=DATE,
    formatogrilla(3, 1) = "D"
    formatogrilla(3, 2) = "S"
    formatogrilla(3, 3) = "S"
    formatogrilla(3, 4) = "S"
    formatogrilla(3, 5) = "S"
    formatogrilla(3, 6) = "S"
    formatogrilla(3, 7) = "S"
    formatogrilla(3, 8) = "S"
    formatogrilla(3, 9) = "D"
    formatogrilla(3, 10) = "D"
    formatogrilla(3, 11) = "N"
    formatogrilla(3, 12) = "N"
    formatogrilla(3, 13) = "N"
    formatogrilla(3, 14) = "S"
    formatogrilla(3, 15) = "S"
    
    
    Rem FORMATO GRILLA
    formatogrilla(4, 11) = "###,###,###,###"
    formatogrilla(4, 12) = "###,###,###,###"
    formatogrilla(4, 13) = "###,###,###,###"
    Rem LOCCKED
    formatogrilla(5, 1) = "TRUE"
    formatogrilla(5, 2) = "TRUE"
    formatogrilla(5, 3) = "TRUE"
    formatogrilla(5, 4) = "TRUE"
    formatogrilla(5, 5) = "TRUE"
    formatogrilla(5, 6) = "TRUE"
    formatogrilla(5, 7) = "TRUE"
    formatogrilla(5, 8) = "TRUE"
    formatogrilla(5, 9) = "TRUE"
    formatogrilla(5, 10) = "TRUE"
    formatogrilla(5, 11) = "TRUE"
    formatogrilla(5, 12) = "TRUE"
    formatogrilla(5, 13) = "TRUE"
    formatogrilla(5, 14) = "TRUE"
    formatogrilla(5, 15) = "TRUE"
    
    Grid1.Cols = 15
    Grid1.Rows = 2
    
     'Grid1.AllowUserResizing = False
    Grid1.DisplayFocusRect = False
    'Grid1.ExtendLastCol = True
    Grid1.BoldFixedCell = False
    
    Grid1.DrawMode = cellOwnerDraw
    
    Grid1.Appearance = Flat
    Grid1.ScrollBarStyle = Flat
    Grid1.FixedRowColStyle = Flat
    
   'Grid1.BackColorFixed = RGB(90, 158, 214)
   ' Grid1.BackColorFixedSel = RGB(110, 180, 230)
   ' Grid1.BackColorBkg = RGB(90, 158, 214)
   ' Grid1.BackColorScrollBar = RGB(231, 235, 247)
   ' Grid1.BackColor1 = RGB(231, 235, 247)
   ' Grid1.BackColor2 = RGB(239, 243, 255)
   ' Grid1.GridColor = RGB(148, 190, 231)
    Grid1.Column(0).Width = 0
    
    For K = 1 To Grid1.Cols - 1
        
        Grid1.Cell(0, K).text = formatogrilla(1, K)
        Grid1.Column(K).Width = Val(formatogrilla(2, K)) * Grid1.DefaultFont.Size
        
        
        Grid1.Column(K).MaxLength = Val(formatogrilla(2, K))
        Grid1.Column(K).FormatString = formatogrilla(4, K)
        Grid1.Column(K).Locked = formatogrilla(5, K)
        If formatogrilla(3, K) = "N" Then Grid1.Column(K).Alignment = cellRightCenter
        If formatogrilla(3, K) = "D" Then Grid1.Column(K).CellType = cellCalendar
        
    Next K
End Sub


Sub leecuentas()
barra.Visible = True
Dim resultados2 As rdoResultset
    Dim cSql2 As New rdoQuery
    Dim rut As String
    Dim PASO As String
    Dim LINEAS As Double
Grid1.AutoRedraw = False

        Set cSql2.ActiveConnection = db
        cSql2.SQL = "SELECT codigo,nombre "
        cSql2.SQL = cSql2.SQL + "FROM cuentasdelmayor "
        cSql2.SQL = cSql2.SQL + "order by codigo"
        cSql2.Execute
        lin = 0
         barra.Min = 0.01
        barra.Max = cSql2.RowsAffected + 4
        LINEAS = 0
        If cSql2.RowsAffected > 0 Then
        Set resultados2 = cSql2.OpenResultset
        While Not resultados2.EOF
        LINEAS = LINEAS + 1
        If Mid(resultados2(0), 5, 4) <> "0000" Then Call LEERMOVIMIENTOS(resultados2(0), resultados2(1))
        barra.Value = LINEAS
        resultados2.MoveNext
        Wend
          
          resultados2.Close
            Set resultados2 = Nothing

        End If
        Grid1.Column(8).Locked = True
        Grid1.Column(9).Locked = True
        Grid1.Column(10).Locked = True
  barra.Visible = False
  
Grid1.AutoRedraw = True
Grid1.Refresh


End Sub

Sub LEERSALDOS(cuenta)
    campos(0, 0) = "codigo"
    campos(1, 0) = "año"
    campos(2, 0) = "debeanterior"
    campos(3, 0) = "haberanterior"
    campos(4, 0) = "debe01"
    campos(5, 0) = "debe02"
    campos(6, 0) = "debe03"
    campos(7, 0) = "debe04"
    campos(8, 0) = "debe05"
    campos(9, 0) = "debe06"
    campos(10, 0) = "debe07"
    campos(11, 0) = "debe08"
    campos(12, 0) = "debe09"
    campos(13, 0) = "debe10"
    campos(14, 0) = "debe11"
    campos(15, 0) = "debe12"
    campos(16, 0) = "haber01"
    campos(17, 0) = "haber02"
    campos(18, 0) = "haber03"
    campos(19, 0) = "haber04"
    campos(20, 0) = "haber05"
    campos(21, 0) = "haber06"
    campos(22, 0) = "haber07"
    campos(23, 0) = "haber08"
    campos(24, 0) = "haber09"
    campos(25, 0) = "HABER10"
    campos(26, 0) = "HABER11"
    campos(27, 0) = "HABER12"
    campos(28, 0) = ""
    
    condicion = "codigo=" + "'" + cuenta + "' and año='" + Mid(fechasistema, 7, 4) + "' order by codigo"
    campos(0, 2) = "saldosdelmayor"
    op = 5
    SQLUTIL.datos = campos
    Set SQLUTIL.conexion = db
    Call SQLUTIL.SQLUTIL(op, condicion)
    
    If SQLUTIL.estado = 4 Then Stop

End Sub
Sub DATOSSALDOS(cuenta)

Call LEERSALDOS(cuenta)
SUMADOR = Val(SQLUTIL.datos(2, 3)) - Val(SQLUTIL.datos(3, 3))
For K = 1 To Val(Mid(fechasistema, 4, 2))
SUMADOR = SUMADOR + Val(SQLUTIL.datos(K + 3, 3)) - Val(SQLUTIL.datos(K + 15, 3))
Next K
saldo = SUMADOR
End Sub

