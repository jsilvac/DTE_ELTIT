VERSION 5.00
Object = "{4777436C-EB8C-4596-98A8-EBCF98683969}#1.0#0"; "FlexCell.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form auxiliar022 
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
   Begin MSComctlLib.ProgressBar barra 
      Height          =   495
      Left            =   0
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
   Begin VB.CommandButton Command4 
      Caption         =   "Exportar Html"
      Height          =   375
      Left            =   4680
      TabIndex        =   3
      Top             =   9720
      Width           =   2055
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Exportar Excel"
      Height          =   375
      Left            =   2400
      TabIndex        =   2
      Top             =   9720
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Imprime "
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   9720
      Width           =   2055
   End
   Begin FlexCell.Grid Grid1 
      Height          =   9015
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   15255
      _ExtentX        =   26908
      _ExtentY        =   15901
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
Attribute VB_Name = "auxiliar022"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
       Private formatogrilla(10, 20) As String
       Private montos(5) As Double
       
Private Sub Command1_Click()
Grid1.DefaultFont.Size = 7
For K = 1 To Grid1.Cols - 1
Grid1.Column(K).Width = Val(formatogrilla(2, K)) * Grid1.DefaultFont.Size
Next K
Grid1.Column(2).Width = 30 * Grid1.DefaultFont.Size

Grid1.PageSetup.Orientation = cellPortrait


Grid1.PageSetup.PrintFixedRow = True


'Grid1.PageSetup.BlackAndWhite = True
Grid1.PageSetup.BottomMargin = 1
Grid1.PageSetup.TopMargin = 1
Grid1.PageSetup.LeftMargin = 1
Grid1.PageSetup.RightMargin = 0


CABEZA ("BALANCE ANALITICO")




Grid1.PrintPreview 75
End Sub

Private Sub Command2_Click()

Grid1.DefaultFont.Size = 6

Grid1.Column(1).Width = 0

For K = 2 To Grid1.Cols - 1
        
        
        Grid1.Column(K).Width = Val(formatogrilla(2, K)) * Grid1.DefaultFont.Size
        
        
    Next K


'Grid1.PageSetup.Orientation = cellLandscape
Grid1.PageSetup.Orientation = cellPortrait



Grid1.PageSetup.PrintFixedRow = True


'Grid1.PageSetup.BlackAndWhite = True
Grid1.PageSetup.BottomMargin = 1
Grid1.PageSetup.TopMargin = 1
Grid1.PageSetup.LeftMargin = 1
Grid1.PageSetup.RightMargin = 0


CABEZA ("BALANCE ANALITICO")




Grid1.PrintPreview 75

End Sub
Sub CABEZA(titulo As String)
Dim objReportTitle As FlexCell.ReportTitle
Grid1.ReportTitles.Clear


Set objReportTitle = New FlexCell.ReportTitle
    objReportTitle.text = titulo
    objReportTitle.Font.Name = "verdana"
    objReportTitle.Font.Size = 12
    objReportTitle.PrintOnAllPages = True
    Grid1.ReportTitles.Add objReportTitle
    
    'Report Title 1
    For K = 1 To 5
    Set objReportTitle = New FlexCell.ReportTitle
    objReportTitle.text = DATOSEMPRESA(K)
    objReportTitle.Font.Name = "verdana"
    objReportTitle.Font.Size = 7
    objReportTitle.Font.Italic = True
    objReportTitle.PrintOnAllPages = True
    objReportTitle.Color = RGB(128, 0, 0)
    objReportTitle.Align = CellLeft
    Grid1.ReportTitles.Add objReportTitle
    Next K
With Grid1.PageSetup
        
        .HeaderFont.Size = 6
        .Header = "                                   PAGINAS &P/&N EMITIDO:&D USUARIO " + USUARIOSISTEMA
        .HeaderAlignment = cellCenter
        .HeaderFont.Name = "Verdana"
        .HeaderMargin = 3
        
End With

End Sub

Private Sub Command3_Click()
Grid1.ExportToExcel ("")
End Sub

Private Sub Command4_Click()
Grid1.ExportToHTML ("")


End Sub


Private Sub Command5_Click()
Grid1.DefaultFont.Size = Grid1.DefaultFont.Size + 0.5
For K = 1 To Grid1.Cols - 1
        Grid1.Column(K).Width = Val(formatogrilla(2, K)) * Grid1.DefaultFont.Size
        

Next K


Grid1.Refresh



End Sub
Private Sub Command6_Click()
Grid1.DefaultFont.Size = Grid1.DefaultFont.Size - 0.5
For K = 1 To Grid1.Cols - 1
        Grid1.Column(K).Width = Val(formatogrilla(2, K)) * Grid1.DefaultFont.Size
        

Next K


Grid1.Refresh



End Sub

Private Sub Form_Load()
 Call Conectar_BD
 Call Conectarconta(servidor, "conta", USUARIO, password)
'Call CONSULTAFECHAS("FECHA PARA IMPRIMIR BALANCE")
CARGAGRILLA
CARGABALANCE



End Sub



    Sub diferencia(Row)
    Grid1.Rows = Row + 1
     With Grid1.Range(Row, 1, Row, 10)
        .Borders(cellEdgeLeft) = cellThin
        .Borders(cellEdgeRight) = cellThin
        .Borders(cellEdgeTop) = cellThin
        .Borders(cellEdgeBottom) = cellThin
        .Borders(cellInsideHorizontal) = cellThin
        .Borders(cellInsideVertical) = cellThin
    End With
    
    Grid1.Cell(Row, 2).text = "RESULTADOS"
   
    For K = 1 To 8
    Grid1.Cell(Row, K + 2).text = difer(K - 1)
  
    Next K
    End Sub
    Sub totalfinal(Row)
    Grid1.Rows = Row + 1
    
     With Grid1.Range(Row, 1, Row, 10)
        .Borders(cellEdgeLeft) = cellThin
        .Borders(cellEdgeRight) = cellThin
        .Borders(cellEdgeTop) = cellThin
        .Borders(cellEdgeBottom) = cellThin
        .Borders(cellInsideHorizontal) = cellThin
        .Borders(cellInsideVertical) = cellThin
    End With
    
    Grid1.Cell(Row, 1).text = ""
    Grid1.Cell(Row, 2).text = "TOTALES"
                 
    For K = 1 To 8
    Grid1.Cell(Row, K + 2).text = sumast(K - 1)
    Next K
    
    End Sub
    




Sub total()
    
End Sub
Sub total1()
    
                
End Sub
Sub LEERSALDOS(LLAVE)
Dim SUMD As Double
Dim SUMH As Double
Dim tipo As String
Dim saldo As Double

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
    
    condicion = "codigo=" + "'" + LLAVE + "' and año ='" + Mid(fechasistema, 7, 4) + "' order by codigo"
    campos(0, 2) = "saldosdelmayor"
    op = 5
    SQLUTIL.datos = campos
    Set SQLUTIL.conexion = db
    Call SQLUTIL.SQLUTIL(op, condicion)
    If SQLUTIL.estado = 4 Then Stop
For K = 1 To 5: montos(K) = 0: Next K
saldo = SQLUTIL.datos(2, 3) - SQLUTIL.datos(3, 3)
SUMD = 0: SUMH = 0
For K = 1 To Val(Mid(fechasistema, 4, 2)) - 1
SUMD = SUMD + SQLUTIL.datos(K + 3, 3)
SUMH = SUMH + SQLUTIL.datos(K + 15, 3)
Next
saldo = saldo + SUMD - SUMH
montos(1) = saldo
SUMD = 0: SUMH = 0
For K = Val(Mid(fechasistema, 4, 2)) To Val(Mid(fechasistema, 4, 2))
SUMD = SUMD + SQLUTIL.datos(K + 3, 3)
SUMH = SUMH + SQLUTIL.datos(K + 15, 3)
Next
montos(2) = SUMD
montos(3) = SUMH
saldo = saldo + SUMD - SUMH
If saldo > 0 Then montos(4) = saldo
If saldo < 0 Then montos(5) = saldo

End Sub




Sub CARGAGRILLA()
Rem DATOS DE LA COLUMNA
    
    Grid1.DefaultFont.Size = 7
    
    formatogrilla(1, 1) = " CODIGO "
    formatogrilla(1, 2) = " CUENTA  "
    formatogrilla(1, 3) = " ANTERIOR"
    formatogrilla(1, 4) = " DEBE  "
    formatogrilla(1, 5) = " HABER "
    formatogrilla(1, 6) = "SALDO DEBE"
    formatogrilla(1, 7) = "SALDO HABER"
    Rem LARGO DE LOS DATOS
    
    formatogrilla(2, 1) = "9"
    formatogrilla(2, 2) = "60"
    formatogrilla(2, 3) = "12"
    formatogrilla(2, 4) = "12"
    formatogrilla(2, 5) = "12"
    formatogrilla(2, 6) = "12"
    formatogrilla(2, 7) = "12"
    Rem TIPO DE DATOS C=CEROS,S=STRING,N=NUMERICO,D=DATE,
    formatogrilla(3, 1) = "S"
    formatogrilla(3, 2) = "S"
    formatogrilla(3, 3) = "N"
    formatogrilla(3, 4) = "N"
    formatogrilla(3, 5) = "N"
    formatogrilla(3, 6) = "N"
    formatogrilla(3, 7) = "N"
    Rem FORMATO GRILLA
    formatogrilla(4, 1) = ""
    formatogrilla(4, 2) = ""
    formatogrilla(4, 3) = "###,###,###,###"
    formatogrilla(4, 4) = "###,###,###,###"
    formatogrilla(4, 5) = "###,###,###,###"
    formatogrilla(4, 6) = "###,###,###,###"
    formatogrilla(4, 7) = "###,###,###,###"
    Rem LOCCKED
    formatogrilla(5, 1) = "TRUE"
    formatogrilla(5, 2) = "TRUE"
    formatogrilla(5, 3) = "TRUE"
    formatogrilla(5, 4) = "TRUE"
    formatogrilla(5, 5) = "TRUE"
    formatogrilla(5, 6) = "TRUE"
    formatogrilla(5, 7) = "TRUE"
    
    
    Grid1.Cols = 8
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
   

Sub CARGABALANCE()
    barra.Visible = True
    
    Dim resultados As rdoResultset
    Dim cSql As New rdoQuery
    Dim rut As String
    Dim LIN As Double
    Dim contaDOR As Double
    With informes
        Set cSql.ActiveConnection = db
        cSql.SQL = "SELECT codigo,nombre,tipo "
        cSql.SQL = cSql.SQL + "FROM cuentasdelmayor"
        cSql.SQL = cSql.SQL + " order by codigo"
        cSql.Execute
        barra.Min = 0.1
        barra.Max = cSql.RowsAffected + 30
        
        
        LIN = 0: contaDOR = 0
        If cSql.RowsAffected > 0 Then
            Set resultados = cSql.OpenResultset
             While Not resultados.EOF
                 
                  Call LEERSALDOS(resultados(0))
                  contaDOR = contaDOR + 1
                  
                  LIN = LIN + 1
                  Grid1.Rows = LIN + 1
            barra.Value = contaDOR
                  If Mid(resultados(0), 5, 4) = "0000" Then
                  LIN = LIN + 1
                  Grid1.Rows = LIN + 1
            
                  With Grid1.Range(LIN, 1, LIN, 7)
                  .FontBold = True
                  .FontUnderline = True
                  End With
                  End If
                  
                  
                  
                  
                  
                  tipocue(0) = resultados(2)
                  
                  Grid1.Cell(LIN, 1).text = Mid(resultados(0), 1, 2) + "." + Mid(resultados(0), 3, 2) + "." + Mid(resultados(0), 5, 4)
                  Grid1.Cell(LIN, 2).text = resultados(1)
                  For K = 1 To 5
                  Grid1.Cell(LIN, K + 2).text = montos(K)
                  Next K
                  
            
            resultados.MoveNext
            Wend
            resultados.Close
            
            Set resultados = Nothing

        End If
    End With
barra.Visible = False


End Sub

Sub Pregunta(ByRef caja As TextBox, ByRef caja1 As TextBox)
    If caja.text = "" And sc = 0 Then caja.SetFocus
    If caja.text <> "" Or sc = 1 Then caja1.Enabled = True: caja1.SetFocus
    sc = 0
End Sub

Private Sub NOIMPRIME_Click()
Unload Me
End Sub

