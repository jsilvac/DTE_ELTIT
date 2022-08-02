VERSION 5.00
Object = "{4777436C-EB8C-4596-98A8-EBCF98683969}#1.0#0"; "FlexCell.ocx"
Begin VB.Form informa01 
   BackColor       =   &H00FFC0C0&
   Caption         =   "Balance Tributario"
   ClientHeight    =   10185
   ClientLeft      =   255
   ClientTop       =   1425
   ClientWidth     =   14790
   LinkTopic       =   "Form1"
   ScaleHeight     =   10185
   ScaleWidth      =   14790
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command4 
      Caption         =   "Exportar Html"
      Height          =   375
      Left            =   4440
      TabIndex        =   3
      Top             =   9720
      Width           =   2055
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Exportar Excel"
      Height          =   375
      Left            =   2280
      TabIndex        =   2
      Top             =   9720
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Imprime Formato Grande"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   9720
      Width           =   2055
   End
   Begin FlexCell.Grid Grid1 
      Height          =   9495
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   15255
      _ExtentX        =   26908
      _ExtentY        =   16748
      AllowUserResizing=   0   'False
      BackColorBkg    =   -2147483639
      BackColorFixed  =   16761024
      BackColorFixedSel=   -2147483639
      Cols            =   5
      DefaultFontName =   "Verdana"
      DefaultFontSize =   6.75
      GridColor       =   -2147483635
      Rows            =   30
   End
End
Attribute VB_Name = "informa01"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
       Private formatogrilla(10, 20) As String
   
Private Sub Command1_Click()
Grid1.DefaultFont.Size = 7
For K = 1 To Grid1.Cols - 1

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
    objReportTitle.text = "Balance tributario"
    objReportTitle.Font.Name = "verdana"
    objReportTitle.Font.Size = 20
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

Private Sub Command4_Click()
Grid1.ExportToHTML ("")


End Sub


Private Sub Form_Load()
 Call Conectar_BD
 Call Conectarconta(servidor, "conta", USUARIO, password)

CARGAGRILLA
CARGAcuentas



End Sub



 



Sub CARGAGRILLA()
Rem DATOS DE LA COLUMNA
    formatogrilla(1, 1) = " CODIGO "
    formatogrilla(1, 2) = " NOMBRE CUENTA  "
    formatogrilla(1, 3) = "TIPO CUENTA"
    formatogrilla(1, 4) = "CUENTA CORRIENTE"
    formatogrilla(1, 5) = "ANALISIS"
    Rem LARGO DE LOS DATOS
    
    formatogrilla(2, 1) = "9"
    formatogrilla(2, 2) = "50"
    formatogrilla(2, 3) = "15"
    formatogrilla(2, 4) = "30"
    formatogrilla(2, 5) = "11"
    Rem TIPO DE DATOS C=CEROS,S=STRING,N=NUMERICO,D=DATE,
    formatogrilla(3, 1) = "S"
    formatogrilla(3, 2) = "S"
    formatogrilla(3, 3) = "S"
    formatogrilla(3, 4) = "S"
    formatogrilla(3, 5) = "S"
    Rem FORMATO GRILLA
    Rem LOCCKED
    formatogrilla(5, 1) = "TRUE"
    formatogrilla(5, 2) = "TRUE"
    formatogrilla(5, 3) = "TRUE"
    formatogrilla(5, 4) = "TRUE"
    formatogrilla(5, 5) = "TRUE"
    Grid1.Cols = 5
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
   

Sub CARGAcuentas()
    Dim resultados As rdoResultset
    Dim TIPOS(4) As String
    Dim cSql As New rdoQuery
    Dim rut As String
    Dim lin As Double
    TIPOS(1) = "ACTIVO"
    TIPOS(2) = "PASIVO"
    TIPOS(3) = "RESULTADO"
    
    With informes
        Set cSql.ActiveConnection = db
        cSql.SQL = "SELECT codigo,nombre,tipo,ctacte,glosa,centrocosto "
        cSql.SQL = cSql.SQL + "FROM cuentasdelmayor"
        cSql.SQL = cSql.SQL + " order by codigo"
        cSql.Execute
        lin = 0
        If cSql.RowsAffected > 0 Then
            Set resultados = cSql.OpenResultset
             While Not resultados.EOF
                lin = lin + 1
                Grid1.Rows = Grid1.Rows + 1
                Grid1.Cell(lin, 1).text = resultados(0)
                Grid1.Cell(lin, 2).text = resultados(1)
                Grid1.Cell(lin, 3).text = TIPOS(resultados(2))
                 Grid1.Cell(lin, 4).text = resultados(3) + " " + resultados(4)
'                Grid1.Cell(lin, 5).text = resultados(4)
                
                
            resultados.MoveNext
            Wend
           
        End If
    End With


End Sub

Sub Pregunta(ByRef caja As TextBox, ByRef caja1 As TextBox)
    If caja.text = "" And sc = 0 Then caja.SetFocus
    If caja.text <> "" Or sc = 1 Then caja1.Enabled = True: caja1.SetFocus
    sc = 0
End Sub

Private Sub NOIMPRIME_Click()
Unload Me
End Sub

