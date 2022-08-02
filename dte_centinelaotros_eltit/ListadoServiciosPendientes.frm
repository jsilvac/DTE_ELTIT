VERSION 5.00
Object = "{4777436C-EB8C-4596-98A8-EBCF98683969}#1.0#0"; "FlexCell.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{ADBBAED6-B16F-40DC-80DF-B44910CBA76C}#1.0#0"; "Frame-Xp.ocx"
Begin VB.Form ListadoServiciosPendientes 
   Caption         =   "Servicios Tecnicos Pendientes"
   ClientHeight    =   10425
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   18000
   LinkTopic       =   "Form1"
   ScaleHeight     =   10425
   ScaleWidth      =   18000
   StartUpPosition =   2  'CenterScreen
   Begin XPFrame.FrameXp FrameXp1 
      Height          =   10455
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   18015
      _ExtentX        =   31776
      _ExtentY        =   18441
      BackColor       =   16761024
      Caption         =   "DETALLE"
      CaptionEstilo3D =   1
      BackColor       =   16761024
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.CommandButton cmdlimpiar 
         BackColor       =   &H00FF8080&
         Caption         =   "LIMPIAR"
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
         Left            =   10080
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   9840
         Width           =   1815
      End
      Begin VB.OptionButton opt3 
         BackColor       =   &H00FFC0C0&
         Caption         =   "ENTREGADOS"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3360
         TabIndex        =   7
         Top             =   9840
         Width           =   1575
      End
      Begin VB.OptionButton opt2 
         BackColor       =   &H00FFC0C0&
         Caption         =   "POR ENTREGAR"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1560
         TabIndex        =   6
         Top             =   9840
         Width           =   1575
      End
      Begin VB.OptionButton opt1 
         BackColor       =   &H00FFC0C0&
         Caption         =   "TODOS"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   480
         TabIndex        =   5
         Top             =   9840
         Width           =   975
      End
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   9360
         Width           =   17775
         _ExtentX        =   31353
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.CommandButton cmdImprimir 
         BackColor       =   &H00FF8080&
         Caption         =   "IMPRIMIR"
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
         Left            =   7920
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   9840
         Width           =   1815
      End
      Begin VB.CommandButton cmdgenerar 
         BackColor       =   &H00FF8080&
         Caption         =   "GENERAR"
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
         Left            =   5760
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   9840
         Width           =   1815
      End
      Begin FlexCell.Grid Grid1 
         Height          =   9015
         Left            =   0
         TabIndex        =   2
         Top             =   240
         Width           =   17895
         _ExtentX        =   31565
         _ExtentY        =   15901
         Cols            =   5
         DefaultFontSize =   8.25
         Rows            =   30
      End
      Begin MSAdodcLib.Adodc data 
         Height          =   330
         Left            =   0
         Top             =   0
         Visible         =   0   'False
         Width           =   1200
         _ExtentX        =   2117
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
   End
End
Attribute VB_Name = "ListadoServiciosPendientes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Sub leerServiciosTecnicosPendientes()
    Dim cSql As New rdoQuery
    Dim resultados As rdoResultset
    Set cSql.ActiveConnection = ventasRubro
    cSql.sql = "select folio,rut,traidopor,ifnull(date_format(fecha,'%d-%m-%Y'),'0'),codigo,descripcion,falla,tecnico,numeroguia,ifnull(date_format(entregado,'%d-%m-%Y'),'0'),observacion "
    cSql.sql = cSql.sql & "from sv_garantias_" & empresaActiva & " "
    If opt1.Value = True Then
        cSql.sql = cSql.sql & "order by folio "
    End If
    If opt2.Value = True Then
        cSql.sql = cSql.sql & "where entregado = '0000-00-00' order by folio "
    End If
    If opt3.Value = True Then
        cSql.sql = cSql.sql & "where entregado <> '0000-00-00' order by folio "
    End If
    
    cSql.Execute
    If cSql.RowsAffected > 0 Then
        Grid1.Rows = 1
        Set resultados = cSql.OpenResultset
        While Not resultados.EOF
            Grid1.Rows = Grid1.Rows + 1
            Grid1.Cell(Grid1.Rows - 1, 1).text = resultados(0)
            Grid1.Cell(Grid1.Rows - 1, 2).text = resultados(1)
            Grid1.Cell(Grid1.Rows - 1, 3).text = resultados(2)
            Grid1.Cell(Grid1.Rows - 1, 4).text = resultados(3)
            Grid1.Cell(Grid1.Rows - 1, 5).text = resultados(4)
            Grid1.Cell(Grid1.Rows - 1, 6).text = resultados(5)
            Grid1.Cell(Grid1.Rows - 1, 7).text = resultados(6)
            Grid1.Cell(Grid1.Rows - 1, 8).text = leerultimomovimiento(resultados(0))
            Grid1.Cell(Grid1.Rows - 1, 9).text = resultados(7)
            Grid1.Cell(Grid1.Rows - 1, 10).text = leerNombreTecnico(resultados(7))
            Grid1.Cell(Grid1.Rows - 1, 11).text = resultados(8)
            Grid1.Cell(Grid1.Rows - 1, 12).text = resultados(9)
            Grid1.Cell(Grid1.Rows - 1, 13).text = resultados(10)
            
            resultados.MoveNext
        Wend
    End If
    
End Sub

Private Sub CargaGrillaInforme(ByVal row As Integer, ByVal col As Integer, ByVal impresion As Grid)
        Dim formatogrilla(20, 20) As String
        Dim i As Integer
        
        Rem DATOS DE LA COLUMNA
        formatogrilla(1, 1) = "NUMERO"
        formatogrilla(1, 2) = "RUT"
        formatogrilla(1, 3) = "NOMBRE"
        formatogrilla(1, 4) = "FECHA RECEP."
        formatogrilla(1, 5) = "CODIGO"
        formatogrilla(1, 6) = "DESCRIPCION"
        formatogrilla(1, 7) = "FALLA"
        formatogrilla(1, 8) = "ESTADO"
        formatogrilla(1, 9) = "TECNICO"
        formatogrilla(1, 10) = "NOMBRE TEC."
        formatogrilla(1, 11) = "NUM. GUIA"
        formatogrilla(1, 12) = "F. RETIRO"
        formatogrilla(1, 13) = "OBSERVACION"
        
        
        Rem LARGO DE LOS DATOS
        formatogrilla(2, 1) = "5"
        formatogrilla(2, 2) = "5"
        formatogrilla(2, 3) = "20"
        formatogrilla(2, 4) = "5"
        formatogrilla(2, 5) = "5"
        formatogrilla(2, 6) = "20"
        formatogrilla(2, 7) = "20"
        formatogrilla(2, 8) = "20"
        formatogrilla(2, 9) = "5"
        formatogrilla(2, 10) = "20"
        formatogrilla(2, 11) = "5"
        If opt2.Value = True Then
            formatogrilla(2, 12) = "0"
        Else
            formatogrilla(2, 12) = "5"
        End If
        formatogrilla(2, 13) = "20"
        
        
        Rem TIPO DE DATOS
        
        formatogrilla(3, 1) = "N"
        formatogrilla(3, 2) = "N"
        formatogrilla(3, 3) = "S"
        formatogrilla(3, 4) = "D"
        formatogrilla(3, 5) = "N"
        formatogrilla(3, 6) = "S"
        formatogrilla(3, 7) = "S"
        formatogrilla(3, 8) = "S"
        formatogrilla(3, 9) = "N"
        formatogrilla(3, 10) = "S"
        formatogrilla(3, 11) = "N"
        formatogrilla(3, 12) = "D"
        formatogrilla(3, 13) = "S"
        
        
        Rem FORMATO GRILLA
        formatogrilla(4, 1) = ""
        formatogrilla(4, 2) = ""
        formatogrilla(4, 3) = ""
        formatogrilla(4, 4) = ""
        formatogrilla(4, 5) = ""
        formatogrilla(4, 6) = ""
        formatogrilla(4, 7) = ""
        formatogrilla(4, 8) = ""
        formatogrilla(4, 9) = ""
        formatogrilla(4, 10) = ""
        formatogrilla(4, 11) = ""
        formatogrilla(4, 12) = ""
        formatogrilla(4, 13) = ""
        
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
        
        Rem VALOR MINIMO
        formatogrilla(6, 1) = ""
        formatogrilla(6, 2) = ""
        formatogrilla(6, 3) = ""
        formatogrilla(6, 4) = ""
        formatogrilla(6, 5) = ""
        formatogrilla(6, 6) = ""
        formatogrilla(6, 7) = ""
        formatogrilla(6, 8) = ""
        formatogrilla(6, 9) = ""
        formatogrilla(6, 10) = ""
        formatogrilla(6, 11) = ""
        formatogrilla(6, 12) = ""
        formatogrilla(6, 13) = ""
        
        Rem VALOR MAXIMO
        formatogrilla(7, 1) = ""
        formatogrilla(7, 2) = ""
        formatogrilla(7, 3) = ""
        formatogrilla(7, 4) = ""
        formatogrilla(7, 5) = ""
        formatogrilla(7, 6) = ""
        formatogrilla(7, 7) = ""
        formatogrilla(7, 8) = ""
        formatogrilla(7, 9) = ""
        formatogrilla(7, 10) = ""
        formatogrilla(7, 11) = ""
        formatogrilla(7, 12) = ""
        formatogrilla(7, 13) = ""
        Rem ANCHO
        formatogrilla(8, 1) = "7"
        formatogrilla(8, 2) = "7"
        formatogrilla(8, 3) = "20"
        formatogrilla(8, 4) = "8"
        formatogrilla(8, 5) = "9"
        formatogrilla(8, 6) = "30"
        formatogrilla(8, 7) = "30"
        formatogrilla(8, 8) = "30"
        formatogrilla(8, 9) = "7"
        formatogrilla(8, 10) = "20"
        formatogrilla(8, 11) = "7"
        If opt2.Value = True Then
            formatogrilla(8, 12) = "0"
        Else
            formatogrilla(8, 12) = "8"
        End If
        formatogrilla(8, 13) = "20"
        
                
        impresion.Cols = col
        impresion.Rows = row
        impresion.Range(0, 0, impresion.Rows - 1, impresion.Cols - 1).Borders(cellEdgeBottom) = cellNone
        impresion.Range(0, 0, impresion.Rows - 1, impresion.Cols - 1).Borders(cellEdgeTop) = cellNone
        impresion.Range(0, 0, impresion.Rows - 1, impresion.Cols - 1).Borders(cellInsideVertical) = cellNone
        impresion.AllowUserResizing = False
        impresion.DisplayFocusRect = False
        impresion.ExtendLastCol = True
        impresion.BoldFixedCell = False
        impresion.DrawMode = cellOwnerDraw
        impresion.Appearance = Flat
        impresion.ScrollBarStyle = Flat
        impresion.FixedRowColStyle = Flat
        impresion.BackColorFixed = RGB(90, 158, 214)
        impresion.BackColorFixedSel = RGB(110, 180, 230)
        impresion.BackColorBkg = RGB(90, 158, 214)
        impresion.BackColorScrollBar = RGB(231, 235, 247)
        impresion.BackColor1 = RGB(231, 235, 247)
        impresion.BackColor2 = RGB(239, 243, 255)
        impresion.GridColor = RGB(148, 190, 231)

        impresion.Column(0).Width = 0
        impresion.RowHeight(0) = impresion.DefaultRowHeight * 1
        impresion.Range(0, 1, 0, impresion.Cols - 1).WrapText = True
        impresion.DefaultFont.Size = 8
        
        For i = 1 To impresion.Cols - 1
            impresion.Cell(0, i).text = formatogrilla(1, i)
            impresion.Column(i).Width = Val(formatogrilla(8, i)) * (impresion.Cell(0, i).Font.Size + 1.25)
            impresion.Column(i).MaxLength = Val(formatogrilla(2, i))
            impresion.Column(i).FormatString = formatogrilla(4, i)
            impresion.Column(i).Locked = formatogrilla(5, i)
            If formatogrilla(3, i) = "N" Then
                impresion.Column(i).Alignment = cellRightCenter
            End If
            If formatogrilla(3, i) = "S" Then
                impresion.Column(i).Alignment = cellLeftCenter
            End If
            If formatogrilla(3, i) = "C" Then
                impresion.Column(i).Alignment = cellCenterCenter
            End If
        Next i
        impresion.Range(0, 1, 0, impresion.Cols - 1).Alignment = cellCenterCenter
        impresion.Range(0, 1, 0, impresion.Cols - 1).Borders(cellEdgeBottom) = cellThin
        impresion.SelectionMode = cellSelectionByRow
        impresion.AllowUserSort = True
    End Sub

Private Sub cmdgenerar_Click()
    leerServiciosTecnicosPendientes
End Sub
Private Sub cmdImprimir_Click()
If Grid1.Rows > 1 Then
        Call Titulos("LISTADO SERVICIOS TECNICOS", Grid1)
        Grid1.AutoRedraw = False
        Grid1.Range(1, 1, 1, Grid1.Cols - 1).Borders(cellEdgeTop) = cellThick
        Grid1.PageSetup.HeaderMargin = 0.5
        Grid1.PageSetup.TopMargin = 1
        Grid1.PageSetup.LeftMargin = 0.5
        Grid1.PageSetup.RightMargin = 0
        Grid1.PageSetup.BottomMargin = 2
        Grid1.PageSetup.FooterMargin = 1
        Grid1.PageSetup.BlackAndWhite = True
        Grid1.PageSetup.Orientation = cellLandscape
        Grid1.PageSetup.PrintFixedRow = True
        Grid1.PageSetup.BlackAndWhite = True
        Call verificaImpresora(5, Grid1)
        Grid1.AutoRedraw = True
End If
End Sub
Sub Titulos(titulo1 As String, impresion As Grid)

    Dim i As Integer
    Dim objReportTitle As FlexCell.ReportTitle
    
    impresion.FixedRowColStyle = Fixed3D
    impresion.CellBorderColorFixed = vbButtonShadow
    impresion.ShowResizeTips = False
    impresion.ReportTitles.Clear
    impresion.PageSetup.CenterHorizontally = True
    impresion.PageSetup.Orientation = cellPortrait
    
      
    impresion.PageSetup.PrintTitleRows = 1
    
    'Logo
'    Grid1.Images.Add App.path & "\Admin.gif", "Logo"
'    Set objReportTitle = New FlexCell.ReportTitle
'    objReportTitle.ImageKey = "Logo"
'    objReportTitle.Align = CellLeft
'    Grid1.ReportTitles.Add objReportTitle
    
    'ENCABEZADO DE PAGINA
    impresion.PageSetup.Header = nombreempresa & vbCrLf & direccionempresa & vbCrLf & comunaempresa
    impresion.PageSetup.HeaderAlignment = cellLeft
    impresion.PageSetup.HeaderFont.Name = "Verdana"
    impresion.PageSetup.HeaderFont.Size = 8
    impresion.PageSetup.HeaderFont.Italic = True
    
    'TITULOS DEL REPORTE
  
    Set objReportTitle = New FlexCell.ReportTitle
    objReportTitle.text = titulo1
    objReportTitle.Font.Name = "Verdana" '"Times New Roman"
    objReportTitle.Font.Size = 10
    objReportTitle.Font.Bold = True
    objReportTitle.Align = cellCenter
    objReportTitle.PrintOnAllPages = True
    impresion.ReportTitles.Add objReportTitle
    
    'PIE DE PAGINA
    impresion.PageSetup.Footer = "Pág &P de &N" & vbCrLf & "Fecha: &D" & vbCrLf & " usuario:" + usuarioSistema
    impresion.PageSetup.FooterAlignment = cellRight
    impresion.PageSetup.FooterFont.Name = "Verdana"
    impresion.PageSetup.FooterFont.Size = 7

    
End Sub

Private Sub cmdlimpiar_Click()
Grid1.Rows = 1
End Sub
 

Private Sub Grid1_DblClick()
    If Grid1.Rows > 1 Then
        If Grid1.ActiveCell.row <> 0 Then
            Load MMGarantias
            With MMGarantias
                .folio.text = Grid1.Cell(Grid1.ActiveCell.row, 1).text
                .Show
                .cargar_servicioafuera
            End With
        End If
    End If
    
End Sub

Private Sub opt1_Click()
    Call CargaGrillaInforme(1, 14, Grid1)
End Sub
Private Sub opt2_Click()
    Call CargaGrillaInforme(1, 14, Grid1)
End Sub
Private Sub opt3_Click()
    Call CargaGrillaInforme(1, 14, Grid1)
End Sub
Function leerultimomovimiento(folio) As String
    Dim tabla As String
        tabla = "select glosa from sv_movimientos_garantias_" & empresaActiva & " "
        tabla = tabla & "where folio='" & folio & "' order by folio desc limit 0,1 "
        Call ConectarControlData(data, servidor, baseVentas & rubro, usuario, password, tabla)
                leerultimomovimiento = ""
                If data.Recordset.RecordCount > 0 Then
                    data.Recordset.MoveFirst
                    While Not data.Recordset.EOF
                         leerultimomovimiento = data.Recordset.Fields("glosa")
                         data.Recordset.MoveNext
                    Wend
                End If
End Function
