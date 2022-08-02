VERSION 5.00
Object = "{4777436C-EB8C-4596-98A8-EBCF98683969}#1.0#0"; "FlexCell.ocx"
Begin VB.Form auxiliar05 
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
Attribute VB_Name = "auxiliar05"
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


cabeza ("LIBRO DE COMPRAS")





Grid1.PrintPreview 75


End Sub
Sub cabeza(titulo As String)
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

Private Sub Form_Load()
    
    Call Conectar_BD
    Call Conectarconta(servidor, "conta", USUARIO, password)
CARGAmayor
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
    año = Mid(fechasistema, 7, 4)
    mes = Mid(fechasistema, 4, 2)

    Dim PASO As String
        Set cSql.ActiveConnection = db
        cSql.SQL = "SELECT folio,facturasdecompras.tipo,numero,fecha,facturasdecompras.rut,cuentascorrientes.nombre,neto,iva,exento,total "
        cSql.SQL = cSql.SQL + "FROM facturasdecompras,cuentascorrientes "
        cSql.SQL = cSql.SQL + "where facturasdecompras.rut=cuentascorrientes.rut and cuentascorrientes.tipo='" + tipoprove + "' and añocontable='" + año + "' and mescontable='" + mes + "' order by tipo,fecha "
        cSql.Execute
        Grid1.AutoRedraw = False
        total(1) = 0
        total(2) = 0
        total(3) = 0
        total(4) = 0
        If cSql.RowsAffected > 0 Then
        Set resultados = cSql.OpenResultset
        lin = 0
         While Not resultados.EOF
          lin = lin + 1
             Grid1.Rows = Grid1.Rows + 1
             For K = 0 To 9
             Grid1.Cell(lin, K + 1).text = resultados(K)
             
             Next K
             MULTI = 1
             If resultados(1) = "1" Then Grid1.Cell(lin, 2).text = "FA"
             If resultados(1) = "2" Then Grid1.Cell(lin, 2).text = "ND"
             If resultados(1) = "3" Then Grid1.Cell(lin, 2).text = "NC": MULTI = -1
             Grid1.Cell(lin, 7).text = resultados(6) * MULTI
             Grid1.Cell(lin, 8).text = resultados(7) * MULTI
             Grid1.Cell(lin, 9).text = resultados(8) * MULTI
             Grid1.Cell(lin, 10).text = resultados(9) * MULTI
             Grid1.Cell(lin, 5).text = Mid(resultados(4), 1, 9) + "-" + Mid(resultados(4), 10, 1)
             Call Consultadetalle(resultados(1), resultados(2), resultados(4))
             total(1) = total(1) + resultados(6) * MULTI
             total(2) = total(2) + resultados(7) * MULTI
             total(3) = total(3) + resultados(8) * MULTI
             total(4) = total(4) + resultados(9) * MULTI
             
             resultados.MoveNext

           
         Wend
          
          resultados.Close
            Set resultados = Nothing

        End If

Grid1.AutoRedraw = True
Grid1.Refresh

End Sub

Sub totallibro()
    Dim TOTALge As Double
      lin = lin + 1
        Grid1.Rows = Grid1.Rows + 1
        Grid1.Range(lin, 7, lin, 10).Borders(cellEdgeTop) = cellThin
        Grid1.Cell(lin, 6).text = "TOTALES"
        Grid1.Cell(lin, 7).text = total(1)
        Grid1.Cell(lin, 8).text = total(2)
        Grid1.Cell(lin, 9).text = total(3)
        Grid1.Cell(lin, 10).text = total(4)
    
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
    
    
    FORMATOGRILLA(1, 1) = "FOLIO"
    FORMATOGRILLA(1, 2) = "TP"
    FORMATOGRILLA(1, 3) = "NUMERO"
    FORMATOGRILLA(1, 4) = "FECHA"
    FORMATOGRILLA(1, 5) = "RUT"
    FORMATOGRILLA(1, 6) = "PROVEEDOR"
    FORMATOGRILLA(1, 7) = "NETO"
    FORMATOGRILLA(1, 8) = "IVA"
    FORMATOGRILLA(1, 9) = "EXENTO"
    FORMATOGRILLA(1, 10) = "TOTAL"
    FORMATOGRILLA(1, 11) = " CUENTA "
     
    Rem LARGO DE LOS DATOS
    
    FORMATOGRILLA(2, 1) = "10"
    FORMATOGRILLA(2, 2) = "3"
    FORMATOGRILLA(2, 3) = "10"
    FORMATOGRILLA(2, 4) = "10"
    FORMATOGRILLA(2, 5) = "10"
    FORMATOGRILLA(2, 6) = "30"
    FORMATOGRILLA(2, 7) = "9"
    FORMATOGRILLA(2, 8) = "9"
    FORMATOGRILLA(2, 9) = "9"
    FORMATOGRILLA(2, 10) = "9"
    FORMATOGRILLA(2, 11) = "30"

    
    Rem TIPO DE DATOS C=CEROS,S=STRING,N=NUMERICO,D=DATE,
    FORMATOGRILLA(3, 1) = "S"
    FORMATOGRILLA(3, 2) = "S"
    FORMATOGRILLA(3, 3) = "S"
    FORMATOGRILLA(3, 4) = "S"
    FORMATOGRILLA(3, 5) = "S"
    FORMATOGRILLA(3, 6) = "S"
    FORMATOGRILLA(3, 7) = "N"
    FORMATOGRILLA(3, 8) = "N"
    FORMATOGRILLA(3, 9) = "N"
    FORMATOGRILLA(3, 10) = "N"
    FORMATOGRILLA(3, 11) = "S"
    Rem FORMATO GRILLA
    FORMATOGRILLA(4, 7) = "###,###,###"
    FORMATOGRILLA(4, 8) = "###,###,###"
    FORMATOGRILLA(4, 9) = "###,###,###"
    FORMATOGRILLA(4, 10) = "###,###,###"
    
    Rem LOCCKED
    For K = 1 To 11
    FORMATOGRILLA(5, K) = "TRUE"
    Next K
    
    Grid1.Cols = 12
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

Sub Consultadetalle(tipo, numero, rut)

Dim resultados2 As rdoResultset
    Dim cSql2 As New rdoQuery
        Set cSql2.ActiveConnection = db
        cSql2.SQL = "SELECT tipo,numero,rut,linea,cuentadelmayor,glosa,monto,dh,centrodecosto "
        cSql2.SQL = cSql2.SQL + "FROM detallefacturasdecompra "
        cSql2.SQL = cSql2.SQL + "where tipo='" + tipo + "' and numero='" + numero + "' and rut='" + rut + "' order by linea "
        cSql2.Execute

        If cSql2.RowsAffected > 0 Then
        Set resultados2 = cSql2.OpenResultset
        
         While Not resultados2.EOF
          For K = 1 To canplan
          If resultados2(4) = plan(K, 1) Then plan(K, 3) = plan(K, 3) + resultados2(6): Grid1.Cell(lin, 11).text = plan(K, 2): K = canplan + 1
      
          Next K
          resultados2.MoveNext

           
         Wend
          
          resultados2.Close
          
        End If

End Sub
Sub CARGAmayor()
    Dim resultados As rdoResultset
    Dim cSql As New rdoQuery
    Dim lineas As Integer
    
    With informes
        Set cSql.ActiveConnection = db
        cSql.SQL = "SELECT codigo,nombre,tipo,ctacte,glosa,centrocosto "
        cSql.SQL = cSql.SQL + "FROM cuentasdelmayor"
        cSql.SQL = cSql.SQL + " order by codigo"
        cSql.Execute
        linea = 0
        If cSql.RowsAffected > 0 Then
            Set resultados = cSql.OpenResultset
             While Not resultados.EOF
             linea = linea + 1
             plan(linea, 1) = resultados(0)
             plan(linea, 2) = resultados(1)
             plan(linea, 3) = 0

            resultados.MoveNext
            Wend
        End If
canplan = linea
    End With


End Sub

