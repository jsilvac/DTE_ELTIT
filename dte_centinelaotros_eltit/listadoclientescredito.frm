VERSION 5.00
Object = "{4777436C-EB8C-4596-98A8-EBCF98683969}#1.0#0"; "Flexcell.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{ADBBAED6-B16F-40DC-80DF-B44910CBA76C}#1.0#0"; "Frame-Xp.ocx"
Begin VB.Form tmplistado5 
   Caption         =   "LISTADO DEUDA CLIENTES"
   ClientHeight    =   9285
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14415
   LinkTopic       =   "Form1"
   ScaleHeight     =   9285
   ScaleWidth      =   14415
   StartUpPosition =   2  'CenterScreen
   Begin XPFrame.FrameXp FrameXp2 
      Height          =   1095
      Left            =   90
      TabIndex        =   1
      Top             =   45
      Width           =   14325
      _ExtentX        =   25268
      _ExtentY        =   1931
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
      Begin VB.OptionButton Option3 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Al dia"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4635
         TabIndex        =   6
         Top             =   495
         Width           =   2535
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Morosos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2295
         TabIndex        =   5
         Top             =   495
         Width           =   2535
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Todos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   315
         TabIndex        =   4
         Top             =   495
         Width           =   2535
      End
      Begin XPFrame.FrameXp FrameXp5 
         Height          =   960
         Left            =   8055
         TabIndex        =   8
         Top             =   45
         Width           =   5010
         _ExtentX        =   8837
         _ExtentY        =   1693
         BackColor       =   12648384
         Caption         =   "TIPOS CLIENTES"
         CaptionEstilo3D =   1
         BackColor       =   12648384
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Begin VB.ComboBox Combotipos 
            Height          =   315
            Left            =   45
            TabIndex        =   9
            Text            =   "Combo1"
            Top             =   315
            Width           =   4875
         End
      End
   End
   Begin XPFrame.FrameXp FrameXp1 
      Height          =   8025
      Left            =   45
      TabIndex        =   0
      Top             =   1215
      Width           =   14340
      _ExtentX        =   25294
      _ExtentY        =   14155
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
      Begin MSComctlLib.ProgressBar BARRA 
         Height          =   240
         Left            =   45
         TabIndex        =   7
         Top             =   7380
         Width           =   14280
         _ExtentX        =   25188
         _ExtentY        =   423
         _Version        =   393216
         Appearance      =   1
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
         Left            =   5805
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   7650
         Width           =   2760
      End
      Begin FlexCell.Grid GRID1 
         Height          =   7080
         Left            =   45
         TabIndex        =   2
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
Attribute VB_Name = "tmplistado5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()

Call Titulos("LISTADO DE CLIENTES CON DEUDA")
GRID1.PageSetup.Orientation = cellLandscape


GRID1.PageSetup.HeaderMargin = 0.5
GRID1.PageSetup.PrintFixedRow = True
GRID1.PageSetup.TopMargin = 1
GRID1.PageSetup.LeftMargin = 0.5
GRID1.PageSetup.RightMargin = 0.5
GRID1.PageSetup.BottomMargin = 2
GRID1.PageSetup.FooterMargin = 2
GRID1.PageSetup.BlackAndWhite = True

GRID1.Range(0, 1, 0, GRID1.Cols - 1).Borders(cellEdgeTop) = cellThick
GRID1.Range(0, 1, 0, GRID1.Cols - 1).Borders(cellEdgeBottom) = cellThick
GRID1.Range(0, 1, 0, GRID1.Cols - 1).Borders(cellEdgeLeft) = cellThick
GRID1.Range(0, 1, 0, GRID1.Cols - 1).Borders(cellEdgeRight) = cellThick
GRID1.Range(0, 1, 0, GRID1.Cols - 1).Borders(cellInsideHorizontal) = cellThick
GRID1.Range(0, 1, 0, GRID1.Cols - 1).Borders(cellInsideVertical) = cellThick



GRID1.PrintPreview
End Sub

Private Sub Form_Load()
Call CargaGrillaGRID1(1, 11)

LEErTIPOSCLIENTES


End Sub

 Private Sub CargaGrillaGRID1(ByVal row As Integer, ByVal col As Integer)
        Dim i As Integer
       Dim formatogrilla(20, 20)
       Rem DATOS DE LA COLUMNA
        formatogrilla(1, 1) = "RUT"
        formatogrilla(1, 2) = "CLIENTE"
        formatogrilla(1, 3) = "CUPO CREDITO"
        formatogrilla(1, 4) = "USADO"
        formatogrilla(1, 5) = "DISPONIBLE"
        formatogrilla(1, 6) = "MORA"
        formatogrilla(1, 7) = "DIASMORA"
        formatogrilla(1, 8) = "F. MORA"
        formatogrilla(1, 9) = "F. EVENTO"
        formatogrilla(1, 10) = "EVENTO"
        
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
        formatogrilla(3, 7) = "N"
        
        Rem FORMATO GRILLA
        ''''''''''''''''''''''''
        formatogrilla(4, 1) = "0000000000"
        formatogrilla(4, 2) = ""
        formatogrilla(4, 3) = "$ ###,###,##0"
        formatogrilla(4, 4) = "$ ###,###,##0"
        formatogrilla(4, 5) = "$ ###,###,##0"
        formatogrilla(4, 6) = "$ ###,###,##0"
        formatogrilla(4, 7) = ""

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
        formatogrilla(8, 4) = "7"
        formatogrilla(8, 5) = "7"
        formatogrilla(8, 6) = "7"
        formatogrilla(8, 7) = "7"
        formatogrilla(8, 8) = "7"
        formatogrilla(8, 9) = "7"
        formatogrilla(8, 10) = "20"
            
        GRID1.Cols = col
        GRID1.Rows = row
        GRID1.AllowUserResizing = True
        GRID1.DisplayFocusRect = False
        GRID1.ExtendLastCol = True
        GRID1.BoldFixedCell = False
        GRID1.DrawMode = cellOwnerDraw
        GRID1.Appearance = Flat
        GRID1.ScrollBarStyle = Flat
        GRID1.FixedRowColStyle = Flat
        GRID1.BackColorFixed = RGB(90, 158, 214)
        GRID1.BackColorFixedSel = RGB(110, 180, 230)
        GRID1.BackColorBkg = RGB(90, 158, 214)
        GRID1.BackColorScrollBar = RGB(231, 235, 247)
        GRID1.BackColor1 = RGB(231, 235, 247)
        GRID1.BackColor2 = RGB(239, 243, 255)
        GRID1.GridColor = RGB(148, 190, 231)
        
        GRID1.Column(0).Width = 0
        For i = 1 To col - 1
            GRID1.Cell(0, i).text = formatogrilla(1, i)
            GRID1.Column(i).Width = Val(formatogrilla(8, i)) * (GRID1.Cell(0, i).Font.Size + 1.25)
            GRID1.Column(i).MaxLength = Val(formatogrilla(2, i))
            GRID1.Column(i).FormatString = formatogrilla(4, i)
            GRID1.Column(i).Locked = formatogrilla(5, i)
            If formatogrilla(3, i) = "N" Then
                GRID1.Column(i).Alignment = cellRightCenter
            End If
            If formatogrilla(3, i) = "S" Then
                GRID1.Column(i).Alignment = cellLeftCenter
            End If
            If formatogrilla(3, i) = "C" Then
                GRID1.Column(i).Alignment = cellCenterCenter
            End If
        Next i
        GRID1.Range(0, 0, 0, GRID1.Cols - 1).Alignment = cellCenterCenter
        GRID1.Enabled = True
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
        If Mid(Combotipos.text, 1, 2) <> "99" Then
        cSql.sql = cSql.sql + "and mc.tipocliente='" + Mid(Combotipos.text, 1, 2) + "' "
        End If
        
        cSql.sql = cSql.sql + "group by cd.rut order by mc.nombre "
        cSql.Execute
        
        If cSql.RowsAffected > 0 Then

            Set resultado = cSql.OpenResultset
            BARRA.Max = cSql.RowsAffected + 1
            BARRA.Value = 0
            
        GRID1.Rows = 1
       GRID1.AutoRedraw = False
        
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
        mora = resultado(5)
        If Option2.Value = True And mora <> 0 Then
        
        GRID1.Rows = GRID1.Rows + 1
        GRID1.Cell(GRID1.Rows - 1, 1).text = resultado(0)
        GRID1.Cell(GRID1.Rows - 1, 2).text = resultado(1)
        
        GRID1.Cell(GRID1.Rows - 1, 3).text = Format(resultado(2), "###,###,###")
        GRID1.Cell(GRID1.Rows - 1, 4).text = Format(usado, "###,###,###")
        GRID1.Cell(GRID1.Rows - 1, 5).text = Format(disponible, "###,###,###")
        
        GRID1.Cell(GRID1.Rows - 1, 6).text = Format(mora, "###,###,###")
        
        total1 = total1 + CREDITO
        total2 = total2 + usado
        total3 = total3 + disponible
        total4 = total4 + mora
        If mora <> 0 Then
        GRID1.Cell(GRID1.Rows - 1, 7).text = LEErdiasmora(resultado(0))
        GRID1.Cell(GRID1.Rows - 1, 8).text = FECHAMORA
        GRID1.Cell(GRID1.Rows - 1, 9).text = LEErultimoeventoFECHA(resultado(0), FECHAMORA)
        GRID1.Cell(GRID1.Rows - 1, 10).text = leernombreevento(LEErultimoevento(resultado(0), FECHAMORA))
        End If
        
        End If
        If Option3.Value = True And mora = 0 Then
        
        GRID1.Rows = GRID1.Rows + 1
        GRID1.Cell(GRID1.Rows - 1, 1).text = resultado(0)
        GRID1.Cell(GRID1.Rows - 1, 2).text = resultado(1)
        
        GRID1.Cell(GRID1.Rows - 1, 3).text = Format(resultado(2), "###,###,###")
        GRID1.Cell(GRID1.Rows - 1, 4).text = Format(usado, "###,###,###")
        GRID1.Cell(GRID1.Rows - 1, 5).text = Format(disponible, "###,###,###")
        
        GRID1.Cell(GRID1.Rows - 1, 6).text = Format(mora, "###,###,###")
        total1 = total1 + CREDITO
        total2 = total2 + usado
        total3 = total3 + disponible
        total4 = total4 + mora
                
        End If
        
        If Option1.Value = True Then
        
        GRID1.Rows = GRID1.Rows + 1
        GRID1.Cell(GRID1.Rows - 1, 1).text = resultado(0)
        GRID1.Cell(GRID1.Rows - 1, 2).text = resultado(1)
        
        GRID1.Cell(GRID1.Rows - 1, 3).text = Format(resultado(2), "###,###,###")
        GRID1.Cell(GRID1.Rows - 1, 4).text = Format(usado, "###,###,###")
        GRID1.Cell(GRID1.Rows - 1, 5).text = Format(disponible, "###,###,###")
        
        GRID1.Cell(GRID1.Rows - 1, 6).text = Format(mora, "###,###,###")
        total1 = total1 + CREDITO
        total2 = total2 + usado
        total3 = total3 + disponible
        total4 = total4 + mora
        If mora <> 0 Then
         GRID1.Cell(GRID1.Rows - 1, 7).text = LEErdiasmora(resultado(0))
        GRID1.Cell(GRID1.Rows - 1, 8).text = FECHAMORA
        GRID1.Cell(GRID1.Rows - 1, 9).text = LEErultimoeventoFECHA(resultado(0), FECHAMORA)
        GRID1.Cell(GRID1.Rows - 1, 10).text = leernombreevento(LEErultimoevento(resultado(0), FECHAMORA))
        End If
        End If
            
        
        
        

            
            resultado.MoveNext
            Wend
        Else
       
        End If
        GRID1.Rows = GRID1.Rows + 1
        GRID1.Range(GRID1.Rows - 1, 3, GRID1.Rows - 1, 6).Borders(cellEdgeTop) = cellThick
        
        
        GRID1.Cell(GRID1.Rows - 1, 2).text = "TOTALES GENERALES"
        
        GRID1.Cell(GRID1.Rows - 1, 3).text = Format(total1, "###,###,###")
        GRID1.Cell(GRID1.Rows - 1, 4).text = Format(total2, "###,###,###")
        GRID1.Cell(GRID1.Rows - 1, 5).text = Format(total3, "###,###,###")
        
        GRID1.Cell(GRID1.Rows - 1, 6).text = Format(total4, "###,###,###")
        
        
        Set resultado = Nothing
        cSql.Close
        Set cSql = Nothing
        GRID1.AutoRedraw = True
        GRID1.Refresh
'        lblUtilizado.Caption = Format(totalusado, "###,###,##0")
'        lblDisponible.Caption = Format(CDbl(lblCupo.Caption) - totalusado, "###,###,##0")
'        totaldeuda.Caption = Format(totalusado, "###,###,##0")
'        moroso.Caption = Format(moratotal, "###,###,##0")
'        totalcuotas.Caption = Format(t1, "###,###,##0")
'        totalmora.Caption = Format(t2, "###,###,##0")
'
    End Sub

Private Sub GRID1_DblClick()
creditoPAGOSTMP.rut2.text = Mid(GRID1.Cell(GRID1.ActiveCell.row, 1).text, 1, 9)
creditoPAGOSTMP.lblDV.Caption = Mid(GRID1.Cell(GRID1.ActiveCell.row, 1).text, 10, 1)
creditoPAGOSTMP.Show

End Sub

Private Sub Option1_Click()
Call CargaGrillaGRID1(1, 11)
LEErclientes

End Sub

Private Sub Option2_Click()
Call CargaGrillaGRID1(1, 11)
LEErclientes

End Sub

Private Sub Option3_Click()
Call CargaGrillaGRID1(1, 11)
LEErclientes

End Sub
Sub Titulos(titulo1)

    Dim i As Integer
    Dim objReportTitle As FlexCell.ReportTitle
    
    GRID1.FixedRowColStyle = Fixed3D
    GRID1.CellBorderColorFixed = vbButtonShadow
    GRID1.ShowResizeTips = False
    GRID1.ReportTitles.Clear
    
  
    
    
    
      
    GRID1.PageSetup.PrintTitleRows = 1
    
    'Logo
'    Grid1.Images.Add App.path & "\Admin.gif", "Logo"
'    Set objReportTitle = New FlexCell.ReportTitle
'    objReportTitle.ImageKey = "Logo"
'    objReportTitle.Align = CellLeft
'    Grid1.ReportTitles.Add objReportTitle
    
    'ENCABEZADO DE PAGINA
    GRID1.PageSetup.Header = nombreempresa & vbCrLf & direccionempresa & vbCrLf & comunaempresa
    GRID1.PageSetup.HeaderAlignment = cellLeft
    GRID1.PageSetup.HeaderFont.Name = "Verdana"
    GRID1.PageSetup.HeaderFont.Size = 8
    
    'TITULOS DEL REPORTE
    If Option1.Value = True Then
    Set objReportTitle = New FlexCell.ReportTitle
    objReportTitle.text = "LISTADO DE CLIENTES GENERAL"
    objReportTitle.Font.Name = "Verdana" '"Times New Roman"
    objReportTitle.Font.Size = 10
    objReportTitle.Font.Bold = True
    objReportTitle.Align = cellCenter
    objReportTitle.PrintOnAllPages = True
    GRID1.ReportTitles.Add objReportTitle
    End If
    If Option2.Value = True Then
    Set objReportTitle = New FlexCell.ReportTitle
    objReportTitle.text = "LISTADO DE CLIENTES MOROSOS"
    objReportTitle.Font.Name = "Verdana" '"Times New Roman"
    objReportTitle.Font.Size = 10
    objReportTitle.Font.Bold = True
    objReportTitle.Align = cellCenter
    objReportTitle.PrintOnAllPages = True
    GRID1.ReportTitles.Add objReportTitle
    End If
    If Option3.Value = True Then
    Set objReportTitle = New FlexCell.ReportTitle
    objReportTitle.text = "LISTADO DE CLIENTES AL DIA"
    objReportTitle.Font.Name = "Verdana" '"Times New Roman"
    objReportTitle.Font.Size = 10
    objReportTitle.Font.Bold = True
    objReportTitle.Align = cellCenter
    objReportTitle.PrintOnAllPages = True
    GRID1.ReportTitles.Add objReportTitle
    
    End If
    Set objReportTitle = New FlexCell.ReportTitle
    objReportTitle.text = titulo1 & "  |  " & "AL DIA  :  " & Format(fechasistema, "dd-mm-yyyy")
    objReportTitle.Font.Name = "Verdana" '"Times New Roman"
    objReportTitle.Font.Size = 8
    objReportTitle.Font.Bold = False
    objReportTitle.Align = cellCenter
    objReportTitle.PrintOnAllPages = True
    GRID1.ReportTitles.Add objReportTitle
    
    'PIE DE PAGINA
    GRID1.PageSetup.Footer = "Pág &P de &N" & vbCrLf & "Fecha: &D"
    GRID1.PageSetup.FooterAlignment = cellRight
    GRID1.PageSetup.FooterFont.Name = "Verdana"
    GRID1.PageSetup.FooterFont.Size = 7
    
End Sub


Sub LEErTIPOSCLIENTES()
    Dim resultados As rdoResultset
    Dim cSql As New rdoQuery
    Dim linea As Double
    
        Set cSql.ActiveConnection = ventas
        cSql.sql = "SELECT codigo,nombre "
        cSql.sql = cSql.sql + "FROM sv_tiposdeclientes "
        cSql.sql = cSql.sql + "ORDER BY codigo "
        cSql.Execute
        linea = 1
        If cSql.RowsAffected > 0 Then
            Set resultados = cSql.OpenResultset
            While Not resultados.EOF
                linea = linea + 1
                Combotipos.AddItem (resultados(0) + " " + resultados(1))
                resultados.MoveNext
            Wend
            resultados.Close
            Set resultados = Nothing
      Combotipos.AddItem ("99 TODOS")
      
      
                
        Combotipos.text = Combotipos.List(linea - 1)
        End If
        
End Sub

