VERSION 5.00
Object = "{4777436C-EB8C-4596-98A8-EBCF98683969}#1.0#0"; "FlexCell.ocx"
Begin VB.Form auxiliar041 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Libro Diario"
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
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   15015
      _ExtentX        =   26485
      _ExtentY        =   16748
      Cols            =   5
      DefaultFontSize =   8.25
      Rows            =   30
   End
End
Attribute VB_Name = "auxiliar041"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private formatogrilla(20, 20)
Private lin As Double



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


cabeza




Grid1.PrintPreview 75


End Sub
Sub cabeza()
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

Private Sub Form_Load()
    
    Call Conectar_BD
    Call Conectarconta(servidor, "conta", USUARIO, password)
CARGAGRILLA
Consulta_Informe
End Sub


    
Sub Consulta_Informe()
Dim resultados As rdoResultset
    Dim cSql As New rdoQuery
    Dim rut As String
    Dim PASO As String
        Set cSql.ActiveConnection = db
        cSql.SQL = "SELECT fecha,tipo,numero,linea,codigocuenta,glosacontable,tipodocumento,numerodocumento,fechadocumento,fechavencimiento,monto,dh "
        cSql.SQL = cSql.SQL + "FROM movimientoscontables "
        cSql.SQL = cSql.SQL + "order by tipo,numero"
        cSql.Execute

        
        
        
        
        Grid1.AutoRedraw = False
        
        If cSql.RowsAffected > 0 Then
        Set resultados = cSql.OpenResultset
        lin = 0: PASO = resultados(1) + resultados(2)
         While Not resultados.EOF
          lin = lin + 1
             Grid1.Rows = Grid1.Rows + 1
             If resultados(1) + resultados(2) <> PASO Then Call totalcomprobante(lin)
             For K = 0 To 9
             Grid1.Cell(lin, K + 1).text = resultados(K)
             Next K
             Grid1.Cell(lin, 5).text = Mid(resultados(4), 1, 2) + "." + Mid(resultados(4), 3, 2) + "." + Mid(resultados(4), 5, 4)
             
             If resultados(11) = "D" Then Grid1.Cell(lin, 11).text = resultados(10): anted = anted + resultados(10)
             If resultados(11) = "H" Then Grid1.Cell(lin, 12).text = resultados(10): anteh = anteh + resultados(10)
             PASO = resultados(1) + resultados(2)
             resultados.MoveNext

           
         Wend
          
          resultados.Close
            Set resultados = Nothing

        End If

Grid1.AutoRedraw = True
Grid1.Refresh

End Sub

Sub totalcomprobante(Row)
    
    
    
    
    
    
    With Grid1.Range(Row, 11, Row, 12)
    
    .Borders(cellEdgeTop) = cellThin
    
    
    
     End With
   With Grid1.Range(Row, 1, Row, 12)
   .FontBold = True
    .FontUnderline = True
    End With
    
    
    
    Grid1.Cell(Row, 10).CellType = cellTextBox
    
    
    Grid1.Cell(Row, 10).text = "TOTAL "
    Grid1.Cell(Row, 11).text = anted
    Grid1.Cell(Row, 12).text = anteh
    lin = lin + 2
             Grid1.Rows = Grid1.Rows + 2
        
        anted = 0: anteh = 0
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
    formatogrilla(1, 13) = "NOMBRE CUENTA"
    formatogrilla(1, 14) = "CUENTA CORRIENTE"
    formatogrilla(1, 15) = "CRCC"
     
    Rem LARGO DE LOS DATOS
    
    formatogrilla(2, 1) = "8"
    formatogrilla(2, 3) = "10"
    formatogrilla(2, 4) = "5"
    formatogrilla(2, 5) = "10"
    formatogrilla(2, 6) = "30"
    formatogrilla(2, 7) = "3"
    formatogrilla(2, 8) = "10"
    formatogrilla(2, 9) = "10"
    formatogrilla(2, 10) = "10"
    formatogrilla(2, 11) = "12"
    formatogrilla(2, 12) = "12"
    formatogrilla(2, 13) = "30"
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
    formatogrilla(3, 13) = "S"
    formatogrilla(3, 14) = "S"
    formatogrilla(3, 15) = "S"
    
    
    Rem FORMATO GRILLA
    formatogrilla(4, 11) = "###,###,###,###"
    formatogrilla(4, 12) = "###,###,###,###"
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

