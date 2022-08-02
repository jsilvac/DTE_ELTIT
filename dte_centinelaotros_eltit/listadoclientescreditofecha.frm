VERSION 5.00
Object = "{4777436C-EB8C-4596-98A8-EBCF98683969}#1.0#0"; "FlexCell.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{ADBBAED6-B16F-40DC-80DF-B44910CBA76C}#1.0#0"; "Frame-Xp.ocx"
Begin VB.Form tmplistado14 
   Caption         =   "LISTADO CLIENTES SEGURO"
   ClientHeight    =   8115
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14415
   LinkTopic       =   "Form1"
   ScaleHeight     =   8115
   ScaleWidth      =   14415
   StartUpPosition =   2  'CenterScreen
   Begin XPFrame.FrameXp FrameXp1 
      Height          =   8145
      Left            =   45
      TabIndex        =   0
      Top             =   0
      Width           =   14340
      _ExtentX        =   25294
      _ExtentY        =   14367
      BackColor       =   16761024
      Caption         =   ""
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
      Begin VB.CommandButton CMDgenerar 
         BackColor       =   &H00FF8080&
         Caption         =   "Generar Informe"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3480
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   7650
         Width           =   2760
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FF8080&
         Caption         =   "Imprimir"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6960
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   7650
         Width           =   2760
      End
      Begin MSComctlLib.ProgressBar BARRA 
         Height          =   240
         Left            =   45
         TabIndex        =   1
         Top             =   7380
         Width           =   14280
         _ExtentX        =   25188
         _ExtentY        =   423
         _Version        =   393216
         Appearance      =   1
      End
      Begin FlexCell.Grid GRID1 
         Height          =   7080
         Left            =   45
         TabIndex        =   3
         Top             =   270
         Width           =   14250
         _ExtentX        =   25135
         _ExtentY        =   12488
         Cols            =   5
         DefaultFontSize =   8.25
         Rows            =   30
      End
   End
End
Attribute VB_Name = "tmplistado14"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CMDgenerar_Click()
     LEErclientes
End Sub

Private Sub Command1_Click()

Call Titulos("LISTADO DE CLIENTES SEGUROS")
Grid1.PageSetup.Orientation = cellPortrait


Grid1.PageSetup.HeaderMargin = 0.5
Grid1.PageSetup.PrintFixedRow = True
Grid1.PageSetup.TopMargin = 1
Grid1.PageSetup.LeftMargin = 0.5
Grid1.PageSetup.RightMargin = 0.5
Grid1.PageSetup.BottomMargin = 2.5
Grid1.PageSetup.FooterMargin = 2
Grid1.PageSetup.BlackAndWhite = True

Grid1.Range(0, 1, 0, Grid1.Cols - 1).Borders(cellEdgeTop) = cellThick
Grid1.Range(0, 1, 0, Grid1.Cols - 1).Borders(cellEdgeBottom) = cellThick
Grid1.Range(0, 1, 0, Grid1.Cols - 1).Borders(cellEdgeLeft) = cellThick
Grid1.Range(0, 1, 0, Grid1.Cols - 1).Borders(cellEdgeRight) = cellThick
Grid1.Range(0, 1, 0, Grid1.Cols - 1).Borders(cellInsideHorizontal) = cellThick
Grid1.Range(0, 1, 0, Grid1.Cols - 1).Borders(cellInsideVertical) = cellThick



Grid1.PrintPreview
End Sub

Private Sub Form_Load()
Call CargaGrillaGRID1(1, 7)

End Sub

 Private Sub CargaGrillaGRID1(ByVal row As Integer, ByVal col As Integer)
        Dim i As Integer
       Dim formatogrilla(20, 20)
       Rem DATOS DE LA COLUMNA
        formatogrilla(1, 1) = "RUT"
        formatogrilla(1, 2) = "CLIENTE"
        formatogrilla(1, 3) = "F. NACIMIENTO"
        formatogrilla(1, 4) = "CUPO CREDITO"
        formatogrilla(1, 5) = "USADO"
        formatogrilla(1, 6) = "DISPONIBLE"
        
        Rem LARGO DE LOS DATOS
        formatogrilla(2, 1) = "10"
        formatogrilla(2, 2) = ""
        formatogrilla(2, 3) = ""
        formatogrilla(2, 4) = ""
        
        Rem TIPO DE DATOS
        formatogrilla(3, 1) = "N"
        formatogrilla(3, 2) = "S"
        formatogrilla(3, 3) = "N"
        formatogrilla(3, 4) = "N"
        formatogrilla(3, 5) = "N"
        formatogrilla(3, 6) = "N"
 
        
        Rem FORMATO GRILLA
        ''''''''''''''''''''''''
        formatogrilla(4, 1) = "0000000000"
        formatogrilla(4, 2) = ""
        formatogrilla(4, 3) = ""
        formatogrilla(4, 4) = "$ ###,###,##0"
        formatogrilla(4, 5) = "$ ###,###,##0"
        formatogrilla(4, 6) = "$ ###,###,##0"
 

        Rem LOCCKED
        formatogrilla(5, 1) = "TRUE"
        formatogrilla(5, 2) = "TRUE"
        formatogrilla(5, 3) = "TRUE"
        formatogrilla(5, 4) = "TRUE"
        formatogrilla(5, 5) = "TRUE"
        formatogrilla(5, 6) = "TRUE"
 
 
        Rem VALOR MINIMO
        formatogrilla(6, 1) = ""
        formatogrilla(6, 2) = ""
        formatogrilla(6, 3) = ""
        formatogrilla(6, 4) = ""
        
        Rem VALOR MAXIMO
        formatogrilla(7, 1) = ""
        formatogrilla(7, 2) = ""
        formatogrilla(7, 3) = ""
        formatogrilla(7, 4) = ""
        
        Rem ANCHO
        formatogrilla(8, 1) = "8"
        formatogrilla(8, 2) = "24"
        formatogrilla(8, 3) = "10"
        formatogrilla(8, 4) = "10"
        formatogrilla(8, 5) = "10"
        formatogrilla(8, 6) = "10"
 
        Grid1.Cols = col
        Grid1.Rows = row
        Grid1.AllowUserResizing = True
        Grid1.DisplayFocusRect = False
        Grid1.ExtendLastCol = True
        Grid1.BoldFixedCell = False
        Grid1.DrawMode = cellOwnerDraw
        Grid1.Appearance = Flat
        Grid1.ScrollBarStyle = Flat
        Grid1.FixedRowColStyle = Flat
        Grid1.BackColorFixed = RGB(90, 158, 214)
        Grid1.BackColorFixedSel = RGB(110, 180, 230)
        Grid1.BackColorBkg = RGB(90, 158, 214)
        Grid1.BackColorScrollBar = RGB(231, 235, 247)
        Grid1.BackColor1 = RGB(231, 235, 247)
        Grid1.BackColor2 = RGB(239, 243, 255)
        Grid1.GridColor = RGB(148, 190, 231)
        
        Grid1.Column(0).Width = 0
        For i = 1 To col - 1
            Grid1.Cell(0, i).text = formatogrilla(1, i)
            Grid1.Column(i).Width = Val(formatogrilla(8, i)) * (Grid1.Cell(0, i).Font.Size + 1.25)
            Grid1.Column(i).MaxLength = Val(formatogrilla(2, i))
            Grid1.Column(i).FormatString = formatogrilla(4, i)
            Grid1.Column(i).Locked = formatogrilla(5, i)
            If formatogrilla(3, i) = "N" Then
                Grid1.Column(i).Alignment = cellRightCenter
            End If
            If formatogrilla(3, i) = "S" Then
                Grid1.Column(i).Alignment = cellLeftCenter
            End If
            If formatogrilla(3, i) = "C" Then
                Grid1.Column(i).Alignment = cellCenterCenter
            End If
        Next i
        Grid1.Range(0, 0, 0, Grid1.Cols - 1).Alignment = cellCenterCenter
        Grid1.Enabled = True
'    Grid1.Column(9).CellType = cellCheckBox
'    Grid1.Column(10).CellType = cellCheckBox
'
    
    End Sub
'**
Sub LEErclientes()

        Dim cSql As rdoQuery
        Dim resultado As rdoResultset
        Dim i As Integer
        Dim CREDITO As Double
        Dim usado As Double
        Dim disponible As Double
        Dim mora As Double
        Dim total1 As Double
        Dim total2 As Double
        Dim total3 As Double
        Dim total4 As Double
        
        Set cSql = New rdoQuery
        Set cSql.ActiveConnection = ventas
        cSql.sql = "SELECT mc.rut,mc.nombre,mc.cupodirecto,sum(cd.montocuota-cd.abono),cupoutilizadodirecto,sum(case when cd.vencimientoactual<'" + Format(fechasistema, "yyyy-mm-dd") + "' then cd.montocuota-cd.abono else '0' end)  "
        cSql.sql = cSql.sql + "FROM sv_maestroclientes as mc inner join sv_cuotas_detalle as cd on (cd.rut=mc.rut) "
        cSql.sql = cSql.sql + "group by cd.rut order by mc.nombre "
        cSql.Execute
        
        If cSql.RowsAffected > 0 Then

            Set resultado = cSql.OpenResultset
            BARRA.Max = cSql.RowsAffected + 1
            BARRA.Value = 0
            
            Grid1.Rows = 2
            Grid1.AutoRedraw = False
        
            total1 = 0
            total2 = 0
            total3 = 0
            total4 = 0
        
            While Not resultado.EOF
                BARRA.Value = BARRA.Value + 1
                BARRA.Refresh
                CREDITO = resultado(2)
                usado = resultado(3)
                disponible = CREDITO - usado
        
                Grid1.Rows = Grid1.Rows + 1
                Grid1.Cell(Grid1.Rows - 1, 1).text = resultado(0)
                Grid1.Cell(Grid1.Rows - 1, 2).text = resultado(1)
                Grid1.Cell(Grid1.Rows - 1, 3).text = Format(leerfechanacimiento(resultado(0)), "dd-mm-yyyy")
                Grid1.Cell(Grid1.Rows - 1, 4).text = Format(resultado(2), "###,###,###")
                Grid1.Cell(Grid1.Rows - 1, 5).text = Format(usado, "###,###,###")
                Grid1.Cell(Grid1.Rows - 1, 6).text = Format(disponible, "###,###,###")

                total1 = total1 + CREDITO
                total2 = total2 + usado
                total3 = total3 + disponible
                resultado.MoveNext
            Wend
        Else
       
        End If
        Grid1.Rows = Grid1.Rows + 1
        Grid1.Range(Grid1.Rows - 1, 3, Grid1.Rows - 1, 6).Borders(cellEdgeTop) = cellThick
        
        Grid1.Cell(Grid1.Rows - 1, 2).text = "TOTALES GENERALES"
        
        Grid1.Cell(Grid1.Rows - 1, 4).text = Format(total1, "###,###,###")
        Grid1.Cell(Grid1.Rows - 1, 5).text = Format(total2, "###,###,###")
        Grid1.Cell(Grid1.Rows - 1, 6).text = Format(total3, "###,###,###")
'        GRID1.Cell(GRID1.Rows - 1, 6).text = Format(total4, "###,###,###")
        
        
        Set resultado = Nothing
        cSql.Close
        Set cSql = Nothing
        Grid1.AutoRedraw = True
        Grid1.Refresh
'
    End Sub

Private Sub GRID1_DblClick()
creditoPAGOSTMP.rut2.text = Mid(Grid1.Cell(Grid1.ActiveCell.row, 1).text, 1, 9)
creditoPAGOSTMP.lbldv.Caption = Mid(Grid1.Cell(Grid1.ActiveCell.row, 1).text, 10, 1)
creditoPAGOSTMP.Show

End Sub

Private Sub Option1_Click()
Call CargaGrillaGRID1(1, 7)
LEErclientes

End Sub

Private Sub Option2_Click()
Call CargaGrillaGRID1(1, 7)
LEErclientes

End Sub

Private Sub Option3_Click()
Call CargaGrillaGRID1(1, 7)
LEErclientes

End Sub
Sub Titulos(titulo1)

    Dim i As Integer
    Dim objReportTitle As FlexCell.ReportTitle
    
    Grid1.FixedRowColStyle = Fixed3D
    Grid1.CellBorderColorFixed = vbButtonShadow
    Grid1.ShowResizeTips = False
    Grid1.ReportTitles.Clear
    
  
    
    
    
      
    Grid1.PageSetup.PrintTitleRows = 1
    
    'Logo
'    Grid1.Images.Add App.path & "\Admin.gif", "Logo"
'    Set objReportTitle = New FlexCell.ReportTitle
'    objReportTitle.ImageKey = "Logo"
'    objReportTitle.Align = CellLeft
'    Grid1.ReportTitles.Add objReportTitle
    
    'ENCABEZADO DE PAGINA
    Grid1.PageSetup.Header = nombreempresa & vbCrLf & direccionempresa & vbCrLf & comunaempresa
    Grid1.PageSetup.HeaderAlignment = cellLeft
    Grid1.PageSetup.HeaderFont.Name = "Verdana"
    Grid1.PageSetup.HeaderFont.Size = 8
    
    'TITULOS DEL REPORTE
 
    Set objReportTitle = New FlexCell.ReportTitle
    objReportTitle.text = "LISTADO DE CLIENTES GENERAL"
    objReportTitle.Font.Name = "Verdana" '"Times New Roman"
    objReportTitle.Font.Size = 10
    objReportTitle.Font.Bold = True
    objReportTitle.Align = cellCenter
    objReportTitle.PrintOnAllPages = True
    Grid1.ReportTitles.Add objReportTitle
 
    
    Set objReportTitle = New FlexCell.ReportTitle
    objReportTitle.text = titulo1 & "  |  " & "AL DIA  :  " & Format(fechasistema, "dd-mm-yyyy")
    objReportTitle.Font.Name = "Verdana" '"Times New Roman"
    objReportTitle.Font.Size = 8
    objReportTitle.Font.Bold = False
    objReportTitle.Align = cellCenter
    objReportTitle.PrintOnAllPages = True
    Grid1.ReportTitles.Add objReportTitle
    
    'PIE DE PAGINA
    Grid1.PageSetup.Footer = "Pág &P de &N" & vbCrLf & "Fecha: &D"
    Grid1.PageSetup.FooterAlignment = cellRight
    Grid1.PageSetup.FooterFont.Name = "Verdana"
    Grid1.PageSetup.FooterFont.Size = 7
    
End Sub
Function leerfechanacimiento(rut) As String
    Dim cSql As New rdoQuery
    Dim resultados As rdoResultset
    Set cSql.ActiveConnection = ventas
    cSql.sql = "select fechanacimiento  from sv_maestroclientes_personales where rut='" & rut & "' "
    cSql.Execute
    leerfechanacimiento = "00-00-0000"
    If cSql.RowsAffected > 0 Then
        Set resultados = cSql.OpenResultset
        leerfechanacimiento = resultados(0)
    End If
    cSql.Close
    Set cSql = Nothing
    Set resultados = Nothing
    
End Function

