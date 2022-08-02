VERSION 5.00
Object = "{4777436C-EB8C-4596-98A8-EBCF98683969}#1.0#0"; "FlexCell.ocx"
Begin VB.Form infoge03 
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
Attribute VB_Name = "infoge03"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private FORMATOGRILLA(20, 20)
Private lin As Double
Private tipoprove As String
Private plan(2000, 3) As Variant
Private canplan As Double
Private total(10) As Double




Private Sub Command1_Click()


 Grid1.DefaultFont.Size = 6.5
For K = 1 To Grid1.Cols - 1
Grid1.Column(K).Width = Val(FORMATOGRILLA(2, K)) * Grid1.DefaultFont.Size
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

leermayor
CARGAGRILLA
Consulta_Informe
totallibro
End Sub


    
Sub Consulta_Informe()
Dim resultados As rdoResultset
    Dim cSql As New rdoQuery
    Dim rut As String
    Dim mes As String
    Dim año As String
    Dim MULTI As Double
    Dim suma As Double
    Dim suma1 As Double
    Dim suma2 As Double
    
    año = Mid(fechasistema, 7, 4)
    mes = Mid(fechasistema, 4, 2)

    Dim PASO As String
        Set cSql.ActiveConnection = db
        cSql.SQL = "SELECT fecha,facturasdecompras.tipo,numero,fechavencimiento,facturasdecompras.rut,cuentascorrientes.nombre,total,abono "
        cSql.SQL = cSql.SQL + "FROM facturasdecompras,cuentascorrientes "
        cSql.SQL = cSql.SQL + "where facturasdecompras.rut=cuentascorrientes.rut and cuentascorrientes.tipo='" + tipoprove + "' and abono<total order by cuentascorrientes.nombre,fecha "
        cSql.Execute
        Grid1.AutoRedraw = False
        
        suma = 0: suma1 = 0: suma2 = 0
        If cSql.RowsAffected > 0 Then
        Set resultados = cSql.OpenResultset
         PASO = resultados(4)
        lin = 0
         While Not resultados.EOF
          lin = lin + 1
             
             Grid1.Rows = Grid1.Rows + 1
             If PASO <> resultados(4) Then Call totalrut(suma, suma1, suma2, lin): suma = 0: suma1 = 0: suma2 = 0: PASO = resultados(4): lin = lin + 1: Grid1.Rows = Grid1.Rows + 1

             Grid1.Cell(lin, 1).text = resultados(0)
             
             MULTI = 1
             If resultados(1) = "1" Then Grid1.Cell(lin, 2).text = "FA"
             If resultados(1) = "2" Then Grid1.Cell(lin, 2).text = "ND"
             If resultados(1) = "3" Then Grid1.Cell(lin, 2).text = "NC": MULTI = -1
             Grid1.Cell(lin, 3).text = resultados(2)
             Grid1.Cell(lin, 4).text = resultados(3)
             Grid1.Cell(lin, 5).text = resultados(4)
             Grid1.Cell(lin, 6).text = resultados(5)
             Grid1.Cell(lin, 7).text = resultados(6) * MULTI
             Grid1.Cell(lin, 8).text = resultados(7) * MULTI
             Grid1.Cell(lin, 9).text = Grid1.Cell(lin, 7).text - Grid1.Cell(lin, 8).text
             
             
             suma = suma + Grid1.Cell(lin, 7).text
             suma1 = suma1 + Grid1.Cell(lin, 8).text
             suma2 = suma2 + Grid1.Cell(lin, 9).text
             
             resultados.MoveNext
           
           
         Wend
          
          resultados.Close
            Set resultados = Nothing

        End If

Grid1.AutoRedraw = True
Grid1.Refresh

End Sub
Sub totalrut(suma, suma1, suma2, linea)
Grid1.Rows = Grid1.Rows + 1
Grid1.Range(linea, 1, linea, 9).FontBold = True
Grid1.Range(linea, 1, linea, 9).FontUnderline = True

Grid1.Cell(linea, 7).text = suma
Grid1.Cell(linea, 8).text = suma1
Grid1.Cell(linea, 9).text = suma2

End Sub
Sub totallibro()
    Dim TOTALge As Double
      lin = lin + 1
        Grid1.Rows = Grid1.Rows + 1
       ' Grid1.Range(lin, 7, lin, 10).Borders(cellEdgeTop) = cellThin
        Grid1.Cell(lin, 6).text = "TOTALES"
        Grid1.Cell(lin, 7).text = total(1)
        Grid1.Cell(lin, 8).text = total(2)
        Grid1.Cell(lin, 9).text = total(3)
'        Grid1.Cell(lin, 10).text = total(4)
    
    TOTALge = 0
    lin = lin + 2
    Grid1.Rows = Grid1.Rows + 2
    
    For K = 1 To canplan
    If plan(K, 3) <> 0 Then
             lin = lin + 1
             Grid1.Rows = Grid1.Rows + 1
        Grid1.Cell(lin, 5).text = plan(K, 1)
        Grid1.Cell(lin, 6).text = plan(K, 2)
        Grid1.Cell(lin, 7).text = plan(K, 3)
        TOTALge = TOTALge + plan(K, 3)
        End If
    Next K
        lin = lin + 1
             Grid1.Rows = Grid1.Rows + 1
        Grid1.Cell(lin, 6).text = "TOTAL DETALLE"
         Grid1.Cell(lin, 7).text = TOTALge
               
    End Sub
    





Sub CARGAGRILLA()
Rem DATOS DE LA COLUMNA
    Grid1.DefaultFont.Size = 7.5
    
    
    FORMATOGRILLA(1, 1) = "FECHA"
    FORMATOGRILLA(1, 2) = "TP"
    FORMATOGRILLA(1, 3) = "NUMERO"
    FORMATOGRILLA(1, 4) = "VENCI."
    FORMATOGRILLA(1, 5) = "RUT"
    FORMATOGRILLA(1, 6) = "PROVEEDOR"
    FORMATOGRILLA(1, 7) = "TOTAL"
    FORMATOGRILLA(1, 8) = "ABONO"
    FORMATOGRILLA(1, 9) = "SALDO"
    
    Rem LARGO DE LOS DATOS
    
    FORMATOGRILLA(2, 1) = "10"
    FORMATOGRILLA(2, 2) = "3"
    FORMATOGRILLA(2, 3) = "10"
    FORMATOGRILLA(2, 4) = "10"
    FORMATOGRILLA(2, 5) = "10"
    FORMATOGRILLA(2, 6) = "30"
    FORMATOGRILLA(2, 7) = "10"
    FORMATOGRILLA(2, 8) = "10"
    FORMATOGRILLA(2, 9) = "10"

    
    Rem TIPO DE DATOS C=CEROS,S=STRING,N=NUMERICO,D=DATE,
    FORMATOGRILLA(3, 1) = "D"
    FORMATOGRILLA(3, 2) = "S"
    FORMATOGRILLA(3, 3) = "S"
    FORMATOGRILLA(3, 4) = "D"
    FORMATOGRILLA(3, 5) = "S"
    FORMATOGRILLA(3, 6) = "S"
    FORMATOGRILLA(3, 7) = "N"
    FORMATOGRILLA(3, 8) = "N"
    FORMATOGRILLA(3, 9) = "N"
    
    Rem FORMATO GRILLA
    FORMATOGRILLA(4, 7) = "###,###,###"
    FORMATOGRILLA(4, 8) = "###,###,###"
    FORMATOGRILLA(4, 9) = "###,###,###"
    
    Rem LOCCKED
    Grid1.Cols = 10
    For K = 1 To Grid1.Cols - 1
    FORMATOGRILLA(5, K) = "TRUE"
    Next K
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
        
        Grid1.Cell(0, K).text = FORMATOGRILLA(1, K)
        Grid1.Column(K).Width = Val(FORMATOGRILLA(2, K)) * Grid1.DefaultFont.Size
        
        
        Grid1.Column(K).MaxLength = Val(FORMATOGRILLA(2, K))
        Grid1.Column(K).FormatString = FORMATOGRILLA(4, K)
        Grid1.Column(K).Locked = FORMATOGRILLA(5, K)
        If FORMATOGRILLA(3, K) = "N" Then Grid1.Column(K).Alignment = cellRightCenter
        If FORMATOGRILLA(3, K) = "D" Then Grid1.Column(K).CellType = cellCalendar
        
    Next K
End Sub

Sub leermayor()
    campos(0, 0) = "codigo"
    campos(1, 0) = "ctacte"
    campos(2, 0) = ""
    campos(0, 2) = "cuentasdelmayor"
    condicion = "codigo=" + "'" + cuentaproveedor + "'"
    op = 5
    SQLUTIL.datos = campos
    Set SQLUTIL.conexion = db
    Call SQLUTIL.SQLUTIL(op, condicion)
    tipoprove = SQLUTIL.datos(1, 3)
    
End Sub


