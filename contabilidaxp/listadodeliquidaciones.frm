VERSION 5.00
Object = "{4777436C-EB8C-4596-98A8-EBCF98683969}#1.0#0"; "FlexCell.ocx"
Object = "{ADBBAED6-B16F-40DC-80DF-B44910CBA76C}#1.0#0"; "Frame-Xp.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form informe10 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " LISTADO DE LIQUIDACIONES DEL MES"
   ClientHeight    =   8805
   ClientLeft      =   3645
   ClientTop       =   1860
   ClientWidth     =   12165
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8805
   ScaleWidth      =   12165
   ShowInTaskbar   =   0   'False
   Begin FlexCell.Grid Grid1 
      Height          =   495
      Left            =   480
      TabIndex        =   6
      Top             =   8160
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      Cols            =   5
      DefaultFontSize =   8.25
      Rows            =   30
   End
   Begin XPFrame.FrameXp FrameXP1 
      Height          =   8655
      Index           =   2
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   11940
      _ExtentX        =   21061
      _ExtentY        =   15266
      BackColor       =   16773879
      Caption         =   " LISTA DE LIQUIDACIONES"
      CaptionEstilo3D =   1
      BackColor       =   16773879
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.CommandButton Command3 
         BackColor       =   &H0080FF80&
         Caption         =   "DIFERENCIA CAJA"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   9360
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   7200
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.CommandButton Command5 
         BackColor       =   &H0080FF80&
         Caption         =   "CONVENIOS"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   4800
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   7200
         Visible         =   0   'False
         Width           =   2175
      End
      Begin FlexCell.Grid Grid2 
         Height          =   735
         Left            =   240
         TabIndex        =   10
         Top             =   7560
         Visible         =   0   'False
         Width           =   135
         _ExtentX        =   238
         _ExtentY        =   1296
         Cols            =   5
         DefaultFontSize =   8.25
         Rows            =   30
      End
      Begin VB.CommandButton Command4 
         BackColor       =   &H0080FF80&
         Caption         =   "INFORME"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   7080
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   7200
         Width           =   2175
      End
      Begin MSComctlLib.ProgressBar BARRA 
         Height          =   255
         Left            =   360
         TabIndex        =   8
         Top             =   8160
         Width           =   11415
         _ExtentX        =   20135
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
         Min             =   1e-4
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H0080FF80&
         Caption         =   "GENERAR"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   7200
         Width           =   2175
      End
      Begin VB.CheckBox Check1 
         Caption         =   "IMPRESION DIRECTA IMPRESORA"
         Height          =   255
         Left            =   8400
         TabIndex        =   5
         Top             =   7800
         Width           =   3135
      End
      Begin FlexCell.Grid GRILLALIQUIDACION 
         Height          =   600
         Left            =   600
         TabIndex        =   4
         Top             =   7800
         Visible         =   0   'False
         Width           =   2580
         _ExtentX        =   4551
         _ExtentY        =   1058
         BackColor1      =   14737632
         BackColor2      =   14737632
         BackColorActiveCellSel=   16777088
         BackColorBkg    =   16761024
         BackColorFixedSel=   16761024
         BackColorScrollBar=   16744576
         BorderColor     =   16744576
         CellBorderColor =   16744576
         CellBorderColorFixed=   16744576
         SelectionBorderColor=   16744576
         DefaultFontName =   "Arial"
         DefaultFontSize =   9.75
         ForeColorFixed  =   8388608
         GridColor       =   16744576
         Rows            =   4
         DateFormat      =   2
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H0080FF80&
         Caption         =   "LIQUIDACIONES"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   2520
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   7200
         Visible         =   0   'False
         Width           =   2175
      End
      Begin FlexCell.Grid GridLiquida 
         Height          =   6690
         Left            =   240
         TabIndex        =   1
         Top             =   405
         Width           =   11565
         _ExtentX        =   20399
         _ExtentY        =   11800
         BackColor1      =   16761024
         BackColor2      =   16761024
         BackColorActiveCellSel=   16777088
         BackColorBkg    =   16761024
         BackColorFixedSel=   16761024
         BackColorScrollBar=   16744576
         BorderColor     =   16744576
         CellBorderColor =   16744576
         CellBorderColorFixed=   16744576
         SelectionBorderColor=   16744576
         Cols            =   3
         DefaultFontName =   "Arial"
         DefaultFontSize =   8.25
         DisplayRowIndex =   -1  'True
         ForeColorFixed  =   8388608
         GridColor       =   16744576
         Rows            =   4
         DateFormat      =   2
      End
   End
   Begin VB.PictureBox MANUAL 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   0
      ScaleHeight     =   315
      ScaleWidth      =   555
      TabIndex        =   2
      Top             =   0
      Width           =   555
   End
End
Attribute VB_Name = "informe10"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim ruttemporal As String
Dim cabezas1 As Variant
Dim fila As Integer
Dim columna As Integer
Dim i As Integer

Private Sub Command1_Click()
Dim labor As String
Dim CRCC As String
Dim U As Double

'For U = 1 To GridLiquida.Rows - 2
'' Call calculaliquidaciones.calculaliquidaciones(GridLiquida.Cell(U, 1).text, Format(fechasistema, "mm"), Format(fechasistema, "YYYY"), empresaactiva)
'liquidacion
'
'                  CRCC = leerdatostrabajador("codigog", "mt_semipermanente", "rut='" + GridLiquida.Cell(U, 1).text + "' and codigotg='0003' and mes='" + Format(fechasistema, "mm") + "' and año='" + Format(fechasistema, "yyyy") + "'", db)
'                   CRCC = leerglosa("0003", CRCC)
'                  labor = leerdatostrabajador("codigog", "mt_semipermanente", "rut='" + GridLiquida.Cell(U, 1).text + "' and codigotg='0002' and mes='" + Format(fechasistema, "mm") + "' and año='" + Format(fechasistema, "yyyy") + "'", db)
'
'                    labor = leerglosa("0002", labor)
'
'Call IMPRIMELIQUIDACION2(GRILLALIQUIDACION, GridLiquida.Cell(U, 1).text, GridLiquida.Cell(U, 2).text, Format(fechasistema, "mm"), Format(fechasistema, "yyyy"), labor, CRCC, Check1.Value, Grid1)
'
'Next U

End Sub
Sub IMPRIMELIQUIDACION2(ByRef grilla As Grid, rut, NOMBRE, MES, año, labor, CRCC, VISTA, impresion As Grid)
    Dim objReportTitle As FlexCell.ReportTitle
    Dim o As Double
    Dim i As Double
    Dim contador As Double
    Dim contadoraux As Double
    Dim CABEZAS2 As Variant
    
    impresion.Rows = 1
    CABEZAS2 = Array("HABERES", "BASE", "MONTO", "DESCUENTOS", "BASE", "MONTO", "")
    Call CargaGrilla23(1, 7, impresion, CABEZAS2)
    
    
    impresion.FixedRowColStyle = Fixed3D
    impresion.CellBorderColorFixed = vbButtonShadow
    impresion.ShowResizeTips = False
    impresion.PageSetup.Orientation = cellPortrait
    
    impresion.DefaultFont.Size = 9
    impresion.Column(1).Width = 27 * 9
    impresion.Column(2).Width = 6 * 9
    impresion.Column(3).Width = 9 * 9
    impresion.Column(4).Width = 27 * 9
    impresion.Column(5).Width = 6 * 9
    impresion.Column(6).Width = 9 * 9
    
    impresion.PageSetup.PrintFixedRow = True
    impresion.ReportTitles.Clear
    impresion.PageSetup.CenterHorizontally = False
    impresion.PageSetup.PrintTitleRows = 0
    impresion.PageSetup.BlackAndWhite = True
     
    
    'ENCABEZADO DE PAGINA
    
    impresion.PageSetup.Header = nombreempresa & vbCrLf & direccionempresa & vbCrLf & rutempresa
    impresion.PageSetup.HeaderAlignment = CellLeft
    impresion.PageSetup.HeaderFont.Name = "Verdana"
    impresion.PageSetup.HeaderFont.Size = 8
    impresion.PageSetup.HeaderFont.Bold = True
    impresion.PageSetup.HeaderMargin = 0.5
    
    'TITULOS DEL REPORTE
    Set objReportTitle = New FlexCell.ReportTitle
    objReportTitle.text = "LIQUIDACION DE SUELDOS  " + MonthName(MES) + " de " + año
    objReportTitle.Font.Name = "Verdana" '"Times New Roman"
    objReportTitle.Font.Size = 10
    objReportTitle.Font.Bold = True
    objReportTitle.Align = cellCenter
    objReportTitle.PrintOnAllPages = True
    impresion.ReportTitles.Add objReportTitle
    
    Set objReportTitle = New FlexCell.ReportTitle
    objReportTitle.text = "TRABAJADOR : " + Mid(rut, 1, 9) + "-" + Mid(rut, 10, 1) + " " + NOMBRE
    objReportTitle.Font.Name = "Verdana" '"Times New Roman"
    objReportTitle.Font.Size = 10
    objReportTitle.Font.Bold = True
    objReportTitle.Align = CellLeft
    objReportTitle.PrintOnAllPages = True
    impresion.ReportTitles.Add objReportTitle
        
    Set objReportTitle = New FlexCell.ReportTitle
    objReportTitle.text = "LABOR : " + labor
    objReportTitle.Font.Name = "Verdana" '"Times New Roman"
    objReportTitle.Font.Size = 10
    objReportTitle.Font.Bold = True
    objReportTitle.Align = CellLeft
    objReportTitle.PrintOnAllPages = True
    impresion.ReportTitles.Add objReportTitle
        
    Set objReportTitle = New FlexCell.ReportTitle
    objReportTitle.text = "CENTROCOSTO : " + CRCC
    objReportTitle.Font.Name = "Verdana" '"Times New Roman"
    objReportTitle.Font.Size = 10
    objReportTitle.Font.Bold = True
    objReportTitle.Align = CellLeft
    objReportTitle.PrintOnAllPages = True
    impresion.ReportTitles.Add objReportTitle
    
    
    'PIE DE PAGINA
    impresion.PageSetup.LeftMargin = 0.5
    impresion.PageSetup.RightMargin = 0.5
    impresion.PageSetup.TopMargin = 2
    impresion.PageSetup.BottomMargin = 0.5
    
    impresion.Range(0, 1, 0, grilla.Cols - 1).Borders(cellEdgeLeft) = cellThin
    impresion.Range(0, 1, 0, grilla.Cols - 1).Borders(cellEdgeTop) = cellThin
    impresion.Range(0, 1, 0, grilla.Cols - 1).Borders(cellEdgeBottom) = cellThin
    impresion.Range(0, 1, 0, grilla.Cols - 1).Borders(cellEdgeRight) = cellThin
    impresion.Range(0, 1, 0, grilla.Cols - 1).Borders(cellInsideHorizontal) = cellThin
    impresion.Range(0, 1, 0, grilla.Cols - 1).Borders(cellInsideVertical) = cellThin
    impresion.PageSetup.PaperWidth = 21.4
'    impresion.PageSetup.PaperHeight = 14
     contador = impresion.Rows - 1
     impresion.Rows = grilla.Rows
     For o = 1 To impresion.Cols - 1
         impresion.Column(o).Locked = False
     Next o
     
     For o = 0 To grilla.Rows - 1
        If contador = 16 Then
            impresion.Range(contador, 1, contador, 6).Merge
        End If
        For i = 0 To grilla.Cols - 1
            impresion.Cell(contador, i).text = grilla.Cell(o, i).text
        Next i
        contador = contador + 1
     Next o
     
     impresion.Rows = impresion.Rows + 6
        

        
        impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, impresion.Cols - 1).Merge
        impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, impresion.Cols - 1).Alignment = cellLeftTop
        impresion.Cell(impresion.Rows - 1, 1).Font.Size = 8
        impresion.Cell(impresion.Rows - 1, 1).Font.Bold = True
        impresion.Cell(impresion.Rows - 1, 1).Font.Name = "verdana"
        impresion.Cell(impresion.Rows - 1, 1).text = nombreempresa
        
        impresion.Rows = impresion.Rows + 1
        impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, impresion.Cols - 1).Merge
        impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, impresion.Cols - 1).Alignment = cellLeftTop
        impresion.Cell(impresion.Rows - 1, 1).Font.Size = 8
        impresion.Cell(impresion.Rows - 1, 1).Font.Bold = True
        impresion.Cell(impresion.Rows - 1, 1).text = direccionempresa


        
        impresion.Rows = impresion.Rows + 1
        impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, impresion.Cols - 1).Merge
        impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, impresion.Cols - 1).Alignment = cellLeftTop
        
        impresion.Cell(impresion.Rows - 1, 1).Font.Size = 8
        impresion.Cell(impresion.Rows - 1, 1).Font.Bold = True
        impresion.Cell(impresion.Rows - 1, 1).text = rutempresa
      
        impresion.Rows = impresion.Rows + 1
        impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, impresion.Cols - 1).Merge
        impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, impresion.Cols - 1).Alignment = cellCenterCenter
        impresion.Cell(impresion.Rows - 1, 1).Font.Size = 10
        impresion.Cell(impresion.Rows - 1, 1).Font.Bold = True
        impresion.Cell(impresion.Rows - 1, 1).text = "LIQUIDACION DE SUELDOS  " + MonthName(MES) + " de " + año
        
 
        impresion.Rows = impresion.Rows + 1
        impresion.Cell(impresion.Rows - 1, 1).Font.Size = 10
        impresion.Cell(impresion.Rows - 1, 1).Font.Bold = True
        impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, impresion.Cols - 1).Merge
        impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, impresion.Cols - 1).Alignment = cellLeftGeneral
        impresion.Cell(impresion.Rows - 1, 1).text = "TRABAJADOR : " + Mid(rut, 1, 9) + "-" + Mid(rut, 10, 1) + " " + NOMBRE
        
        impresion.Rows = impresion.Rows + 1
        impresion.Cell(impresion.Rows - 1, 1).Font.Size = 10
        impresion.Cell(impresion.Rows - 1, 1).Font.Bold = True
        impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, impresion.Cols - 1).Merge
        impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, impresion.Cols - 1).Alignment = cellLeftGeneral
        impresion.Cell(impresion.Rows - 1, 1).text = "LABOR : " + labor
        
        impresion.Rows = impresion.Rows + 1
        impresion.Cell(impresion.Rows - 1, 1).Font.Size = 10
        impresion.Cell(impresion.Rows - 1, 1).Font.Bold = True
        impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, impresion.Cols - 1).Merge
        impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, impresion.Cols - 1).Alignment = cellLeftGeneral
        impresion.Cell(impresion.Rows - 1, 1).text = "CENTROCOSTO : " + CRCC
        
        
        impresion.Rows = impresion.Rows + 1
        For o = 1 To impresion.Cols - 1
        impresion.Cell(impresion.Rows - 1, o).text = impresion.Cell(0, o).text
        impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, impresion.Cols - 1).Borders(cellEdgeTop) = cellThin
        impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, impresion.Cols - 1).Borders(cellEdgeLeft) = cellThin
        impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, impresion.Cols - 1).Borders(cellEdgeRight) = cellThin
        impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, impresion.Cols - 1).Borders(cellEdgeBottom) = cellThin
        impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, impresion.Cols - 1).Borders(cellInsideVertical) = cellThin
        impresion.Range(impresion.Rows - 1, 1, impresion.Rows - 1, impresion.Cols - 1).Borders(cellInsideHorizontal) = cellThin
          
        Next o
      contadoraux = 1
     contador = impresion.Rows
     impresion.Rows = impresion.Rows + grilla.Rows
     For o = 1 To grilla.Rows - 1
        If contadoraux = 16 Then
            impresion.Range(contador, 1, contador, 6).Merge
        End If
        For i = 0 To grilla.Cols - 1
            impresion.Cell(contador, i).text = grilla.Cell(o, i).text
        Next i
        contador = contador + 1
        contadoraux = contadoraux + 1
     Next o
     
impresion.Range(13, 1, 13, 6).Borders(cellEdgeBottom) = cellThin
impresion.Range(13, 1, 13, 6).Borders(cellEdgeLeft) = cellThin
impresion.Range(13, 1, 13, 6).Borders(cellEdgeTop) = cellThin
impresion.Range(13, 1, 13, 6).Borders(cellEdgeRight) = cellThin
impresion.Range(13, 3, 13, 3).Borders(cellEdgeRight) = cellThin

impresion.Range(44, 1, 44, 6).Borders(cellEdgeBottom) = cellThin
impresion.Range(44, 1, 44, 6).Borders(cellEdgeLeft) = cellThin
impresion.Range(44, 1, 44, 6).Borders(cellEdgeTop) = cellThin
impresion.Range(44, 1, 44, 6).Borders(cellEdgeRight) = cellThin
impresion.Range(44, 3, 44, 3).Borders(cellEdgeRight) = cellThin

    If VISTA = "1" Then
    impresion.DirectPrint
    Else
    
    impresion.PrintPreview
    End If
    
    impresion.DefaultFont.Size = 10
    impresion.Column(1).Width = 30 * 10
    impresion.Column(2).Width = 5 * 10
    impresion.Column(3).Width = 10 * 10
    impresion.Column(4).Width = 30 * 10
    impresion.Column(5).Width = 5 * 10
    impresion.Column(6).Width = 10 * 10
    
    
    
    
End Sub

Private Sub COMMAND2_Click()
Call CargaGridLiquida(GridLiquida)
End Sub

Private Sub Command3_Click()
   Dim U As Double
    IniciaGrid2
    For U = 1 To GridLiquida.Rows - 2
    If GridLiquida.Cell(U, GridLiquida.Cols - 1).text <> "" Then
    Call IMPRIMEVALEdiferencia(GridLiquida.Cell(U, 1).text, GridLiquida.Cell(U, 2).text, Grid2, Check1.Value, GridLiquida.Cell(U, GridLiquida.Cols - 1).text, "")
End If
Next U
End Sub
Sub imprimir(grilla As Grid)
Dim objReportTitle As FlexCell.ReportTitle
    
    
    grilla.FixedRowColStyle = Fixed3D
    grilla.CellBorderColorFixed = vbButtonShadow
    grilla.ShowResizeTips = False
    grilla.PageSetup.Orientation = cellPortrait
    
'    grilla.DefaultFont.Size = 7
    
    
    grilla.PageSetup.PrintFixedRow = True
    grilla.ReportTitles.Clear
    grilla.PageSetup.CenterHorizontally = False
    grilla.PageSetup.PrintTitleRows = 1
    grilla.PageSetup.BlackAndWhite = True
    
    
    'ENCABEZADO DE PAGINA
    
    grilla.PageSetup.Header = nombreempresa & vbCrLf & direccionempresa
    grilla.PageSetup.HeaderAlignment = CellLeft
    grilla.PageSetup.HeaderFont.Name = "Verdana"
    grilla.PageSetup.HeaderFont.Size = 12
    grilla.PageSetup.HeaderFont.Bold = True
    grilla.PageSetup.HeaderMargin = 0.5
    
    
    
    
    'TITULOS DEL REPORTE
    Set objReportTitle = New FlexCell.ReportTitle
    objReportTitle.text = "LISTADO DE LIQUIDACIONES EMITIDAS PARA PAGO"
    objReportTitle.Font.Name = "Verdana" '"Times New Roman"
    objReportTitle.Font.Size = 10
    objReportTitle.Font.Bold = True
    objReportTitle.Align = cellCenter
    objReportTitle.PrintOnAllPages = True
    grilla.ReportTitles.Add objReportTitle
    
   
    'PIE DE PAGINA
    grilla.PageSetup.LeftMargin = 0.5
    grilla.PageSetup.RightMargin = 0.1
    grilla.PageSetup.TopMargin = 3
    grilla.PageSetup.BottomMargin = 0.5
    
    
    grilla.Range(0, 1, 0, grilla.Cols - 1).Borders(cellEdgeLeft) = cellThin
    grilla.Range(0, 1, 0, grilla.Cols - 1).Borders(cellEdgeTop) = cellThin
    grilla.Range(0, 1, 0, grilla.Cols - 1).Borders(cellEdgeBottom) = cellThin
    grilla.Range(0, 1, 0, grilla.Cols - 1).Borders(cellEdgeRight) = cellThin
    grilla.Range(0, 1, 0, grilla.Cols - 1).Borders(cellInsideHorizontal) = cellNone

    grilla.Range(0, 1, 0, grilla.Cols - 1).Borders(cellInsideVertical) = cellNone
    
    grilla.PrintPreview
    
    
    
    
End Sub

Private Sub Command4_Click()
    Call imprimir(GridLiquida)
End Sub

Private Sub Command5_Click()
Dim U As Double
IniciaGrid2
For U = 1 To GridLiquida.Rows - 2
If GridLiquida.Cell(U, 4).text <> "" Then
    Call convenios(GridLiquida.Cell(U, 1).text, GridLiquida.Cell(U, 2).text, Grid2, Check1.Value)
End If
Next U

End Sub
Sub IMPRIMEVALEpago(rut, NOMBRE, impresion As Grid, VISTA, TOTAL)
    'ENCABEZADO DE PAGINA
    'TITULOS DEL REPORTE
    impresion.Rows = 50
    Rem pagina 1
    impresion.Range(1, 1, 1, impresion.Cols - 1).Merge
    impresion.Range(2, 1, 2, impresion.Cols - 1).Merge
    impresion.Range(3, 1, 3, impresion.Cols - 1).Merge
    impresion.Range(4, 1, 4, impresion.Cols - 1).Merge
    impresion.Range(5, 1, 5, impresion.Cols - 1).Merge
    impresion.Range(6, 1, 6, impresion.Cols - 1).Merge
    impresion.Range(7, 1, 7, impresion.Cols - 1).Merge
    
    impresion.DefaultFont.Size = 10
    impresion.DefaultFont.Bold = True
    impresion.Cell(1, 1).text = nombreempresa
    impresion.Cell(2, 1).Font.Size = "8"
    impresion.Cell(2, 1).text = Format(fechasistema, "dd") & " de " & MonthName(Format(fechasistema, "mm")) & " de " & Format(fechasistema, "yyyy")
    
    impresion.Cell(3, 1).Alignment = cellCenterCenter
    impresion.Cell(33, 1).Alignment = cellCenterCenter
    impresion.Cell(3, 1).text = "COMPROBANTE DE PAGO "
    impresion.Cell(5, 1).text = "           Se ha abonado a su Cta. Cte. Crédito con Promotora Palguin LTDA"
    impresion.Cell(6, 1).text = "la     suma    de " & WORDNUM(TOTAL, "PESO", "PESOS", 0)
    impresion.Range(10, 1, 10, impresion.Cols - 1).Merge
    impresion.Cell(10, 1).text = "Trabajador :" & Mid(rut, 1, 9) & "-" & Mid(rut, 10, 1) & "  " & NOMBRE
 
    
    impresion.Cell(13, 3).text = "_________________"
    impresion.Cell(14, 3).text = "FIRMA TRABAJADOR"
    impresion.Cell(15, 3).text = Mid(rut, 1, 9) & "-" & Mid(rut, 10, 1)
    
    
    For k = 1 To 15
    impresion.Cell(k + 30, 1).text = impresion.Cell(k, 1).text
    impresion.Cell(k + 30, 2).text = impresion.Cell(k, 2).text
    impresion.Cell(k + 30, 3).text = impresion.Cell(k, 3).text
    
    Next k
    impresion.Range(31, 1, 31, impresion.Cols - 1).Merge
    impresion.Range(32, 1, 32, impresion.Cols - 1).Merge
    impresion.Range(33, 1, 33, impresion.Cols - 1).Merge
    impresion.Range(34, 1, 34, impresion.Cols - 1).Merge
    impresion.Range(35, 1, 35, impresion.Cols - 1).Merge
    impresion.Range(36, 1, 36, impresion.Cols - 1).Merge
    impresion.Range(37, 1, 37, impresion.Cols - 1).Merge
    impresion.Range(40, 1, 40, impresion.Cols - 1).Merge
    Rem pagina 2
    
    If VISTA = 0 Then
    impresion.PrintPreview
    Else
    impresion.DirectPrint
    End If
End Sub
Sub IMPRIMEVALEprestamo(rut, NOMBRE, impresion As Grid, VISTA, TOTAL, glosa)
    'ENCABEZADO DE PAGINA
    'TITULOS DEL REPORTE
    impresion.Rows = 50
    Rem pagina 1
    impresion.Range(1, 1, 1, impresion.Cols - 1).Merge
    impresion.Range(2, 1, 2, impresion.Cols - 1).Merge
    impresion.Range(3, 1, 3, impresion.Cols - 1).Merge
    impresion.Range(4, 1, 4, impresion.Cols - 1).Merge
    impresion.Range(5, 1, 5, impresion.Cols - 1).Merge
    impresion.Range(6, 1, 6, impresion.Cols - 1).Merge
    impresion.Range(7, 1, 7, impresion.Cols - 1).Merge
    
    
    impresion.DefaultFont.Size = 10
    impresion.DefaultFont.Bold = True
    impresion.Cell(1, 1).text = nombreempresa
    impresion.Cell(2, 1).Font.Size = "8"
    impresion.Cell(2, 1).text = Format(fechasistema, "dd") & " de " & MonthName(Format(fechasistema, "mm")) & " de " & Format(fechasistema, "yyyy")

    impresion.Cell(3, 1).Alignment = cellCenterCenter
    impresion.Cell(33, 1).Alignment = cellCenterCenter
    impresion.Cell(3, 1).text = "COMPROBANTE DE PAGO PRESTAMO"
    impresion.Cell(5, 1).text = "           Se ha abonado a su Cta. Cte. Crédito con Promotora Palguin LTDA"
    impresion.Cell(6, 1).text = "la     suma    de  " & WORDNUM(TOTAL, "PESO", "PESOS", 0)
    impresion.Cell(7, 1).text = "por    concepto    de      pago    de  cuota   " & Replace(glosa, "PROMOTORA", "")
    impresion.Range(10, 1, 10, impresion.Cols - 1).Merge
    impresion.Cell(10, 1).text = "Trabajador :" & Mid(rut, 1, 9) & "-" & Mid(rut, 10, 1) & "  " & NOMBRE
 
    
    impresion.Cell(13, 3).text = "_________________"
    impresion.Cell(14, 3).text = "FIRMA TRABAJADOR"
    impresion.Cell(15, 3).text = Mid(rut, 1, 9) & "-" & Mid(rut, 10, 1)
    
    For k = 1 To 15
    impresion.Cell(k + 30, 1).text = impresion.Cell(k, 1).text
    impresion.Cell(k + 30, 2).text = impresion.Cell(k, 2).text
    impresion.Cell(k + 30, 3).text = impresion.Cell(k, 3).text
    
    Next k
    impresion.Range(31, 1, 31, impresion.Cols - 1).Merge
    impresion.Range(32, 1, 32, impresion.Cols - 1).Merge
    impresion.Range(33, 1, 33, impresion.Cols - 1).Merge
    impresion.Range(34, 1, 34, impresion.Cols - 1).Merge
    impresion.Range(35, 1, 35, impresion.Cols - 1).Merge
    impresion.Range(36, 1, 36, impresion.Cols - 1).Merge
    impresion.Range(37, 1, 37, impresion.Cols - 1).Merge
    impresion.Range(40, 1, 40, impresion.Cols - 1).Merge
    Rem pagina 2
    
    If VISTA = 0 Then
    impresion.PrintPreview
    Else
    impresion.DirectPrint
    End If
End Sub


Sub IMPRIMEVALEdiferencia(rut, NOMBRE, impresion As Grid, VISTA, TOTAL, glosa)
    'ENCABEZADO DE PAGINA
    'TITULOS DEL REPORTE
    impresion.Rows = 50
    Rem pagina 1
    impresion.Range(1, 1, 1, impresion.Cols - 1).Merge
    impresion.Range(2, 1, 2, impresion.Cols - 1).Merge
    impresion.Range(3, 1, 3, impresion.Cols - 1).Merge
    impresion.Range(4, 1, 4, impresion.Cols - 1).Merge
    impresion.Range(5, 1, 5, impresion.Cols - 1).Merge
    impresion.Range(6, 1, 6, impresion.Cols - 1).Merge
    impresion.Range(7, 1, 7, impresion.Cols - 1).Merge
    impresion.Range(8, 1, 8, impresion.Cols - 1).Merge
    
    
    impresion.DefaultFont.Size = 10
    impresion.DefaultFont.Bold = True
    impresion.Cell(1, 1).text = nombreempresa
    impresion.Cell(2, 1).Font.Size = "8"
    impresion.Cell(2, 1).text = Format(fechasistema, "dd") & " de " & MonthName(Format(fechasistema, "mm")) & " de " & Format(fechasistema, "yyyy")

    impresion.Cell(3, 1).Alignment = cellCenterCenter
    impresion.Cell(33, 1).Alignment = cellCenterCenter
    impresion.Cell(4, 1).text = "           Yo " & NOMBRE & " Cajero (a) "
    impresion.Cell(5, 1).text = "Reconosco  como    mi     responsabilidad , la     perdida d   e  dinero   en    mi"
    impresion.Cell(6, 1).text = "caja   durante     el   mes    de " & MonthName(Format(fechasistema, "mm")) & " razón    por     la    cual     autorizo"
    impresion.Cell(7, 1).text = "a     mi   empleador    descontar    del    total   de    mis   remuneraciones    la "
    impresion.Cell(8, 1).text = "cantidad de  $ " & TOTAL
  
    impresion.Cell(12, 2).text = "_________________"
    impresion.Cell(13, 2).text = "      FIRMA "
    
    For k = 1 To 15
    impresion.Cell(k + 30, 1).text = impresion.Cell(k, 1).text
    impresion.Cell(k + 30, 2).text = impresion.Cell(k, 2).text
    impresion.Cell(k + 30, 3).text = impresion.Cell(k, 3).text
    Next k
    
    impresion.Range(31, 1, 31, impresion.Cols - 1).Merge
    impresion.Range(32, 1, 32, impresion.Cols - 1).Merge
    impresion.Range(33, 1, 33, impresion.Cols - 1).Merge
    impresion.Range(34, 1, 34, impresion.Cols - 1).Merge
    impresion.Range(35, 1, 35, impresion.Cols - 1).Merge
    impresion.Range(36, 1, 36, impresion.Cols - 1).Merge
    impresion.Range(37, 1, 37, impresion.Cols - 1).Merge
    impresion.Range(38, 1, 38, impresion.Cols - 1).Merge
    Rem pagina 2
    
    If VISTA = 0 Then
    impresion.PrintPreview
    Else
    impresion.DirectPrint
    End If
End Sub



Sub Cargadatosconvenios(LINEA, rut, grilla As Grid)
    Dim resultados As rdoResultset
    Dim csql As New rdoQuery
    
    Set csql.ActiveConnection = contadb
    csql.sql = "SELECT hd.codigo, hd.glosa, hd.monto "
    csql.sql = csql.sql + " FROM calculoliquidaciones as hd inner join " + clientesistema + "remu.tabladecalculo as tc on tc.codigo=hd.codigo and tc.convenio='1' "
    csql.sql = csql.sql + " WHERE rut= '" & rut & "'"
    csql.sql = csql.sql + " AND mes= '" & Format(fechasistema, "mm") & "'"
    csql.sql = csql.sql + " AND año= '" & Format(fechasistema, "yyyy") & "'"
    csql.sql = csql.sql + " ORDER BY hd.codigo"
    csql.Execute
    If csql.RowsAffected > 0 Then
        Set resultados = csql.OpenResultset
        While Not resultados.EOF
            
            grilla.Cell(LINEA, 1).text = ""
            grilla.Cell(LINEA, 2).text = resultados(1) 'glosa
            grilla.Cell(LINEA, 3).text = Format(resultados(2), "###,###,###") 'monto
            resultados.MoveNext
            LINEA = LINEA + 1
        Wend
        resultados.Close
        Set resultados = Nothing
    End If
    '--------------------------------
    
End Sub

Sub convenios(rut, NOMBRE, impresion As Grid, VISTA)
    Dim resultados As rdoResultset
    Dim csql As New rdoQuery
    
    Set csql.ActiveConnection = contadb
    csql.sql = "SELECT hd.codigo, hd.glosa, sum(hd.monto) "
    csql.sql = csql.sql + " FROM calculoliquidaciones as hd inner join " + clientesistema + "remu.tabladecalculo as tc on tc.codigo=hd.codigo and tc.convenio='1' "
    csql.sql = csql.sql + " WHERE rut= '" & rut & "'"
    csql.sql = csql.sql + " AND mes= '" & Format(fechasistema, "mm") & "'"
    csql.sql = csql.sql + " AND año= '" & Format(fechasistema, "yyyy") & "'"
    csql.sql = csql.sql + " group by hd.codigo ORDER BY hd.codigo"
    csql.Execute
    If csql.RowsAffected > 0 Then
        Set resultados = csql.OpenResultset
        While Not resultados.EOF
        impresion.Rows = 1
        If resultados(0) = "00160" Then
            Call IMPRIMEVALEprestamo(rut, NOMBRE, impresion, VISTA, resultados(2), resultados(1))
        End If
        If resultados(0) = "00039" Then
            Call IMPRIMEVALEpago(rut, NOMBRE, impresion, VISTA, resultados(2))
        End If
          resultados.MoveNext
        Wend
        resultados.Close
        Set resultados = Nothing
    End If
    '--------------------------------
    
End Sub


Private Sub Form_Activate()
'sqlremu.audit = True
'sqlremu.programaactivo = Me.Caption
End Sub

'***********************************************************************
'***********************************************************************
Private Sub Form_Load()
    Call configuracion.Conectar_BD 'db
  '  Call configuracion.ConectarRemu(servidor, clientesistema + "remu", usuario, password) 'remu
    MODIFI = 0
    '----------------------------------------------
    IniciaGrid23
    
    Call Me.IniciaGrid1
End Sub

Private Sub GridLiquida_DblClick()
'ingreso1.dato(5).text = Mid(GridLiquida.Cell(GridLiquida.ActiveCell.Row, 1).text, 1, 10)
'ingreso1.Show

End Sub

Private Sub GridLiquida_LeaveRow(ByVal row As Long, Cancel As Boolean)
Rem Call Command2_Click

End Sub

'************************************************************************
'************************************************************************
Private Sub MANUAL_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 27 'esc
            Unload Me
    End Select
End Sub

'************************************************************************
'************************************************************************
Private Sub GridLiquida_Click()
  
End Sub

Private Sub GridLiquida_KeyDown(KeyCode As Integer, ByVal Shift As Integer)
    fila = GridLiquida.ActiveCell.row
    columna = GridLiquida.ActiveCell.col
    Select Case KeyCode
        Case 27 'esc
            MANUAL.SetFocus
        Case 77 'M:modificar
        Case 46 'suprimir/eliminar
    End Select
End Sub

'************************************************************************
'************************************************************************
Sub IniciaGrid1()
    cabezas1 = Array("RUT", "NOMBRE", "LIQUIDO", "CONVENIOS", "SALDO", "DIFERENCIA CAJA", "")
    Call CargaGrilla1(1, 7, GridLiquida, cabezas1)
    
End Sub

Sub IniciaGrid2()
    cabezas1 = Array("NADA", "")
    Call CARGAGRILLA2(1, 5, Grid2, cabezas1)
    
End Sub
Sub CARGAGRILLA2(numRow, numCol, grilla As Grid, camposgrid As Variant)
    Dim FORMATOGRILLA(50, 50) As String
    
    Dim i As Integer
    i = 0
    While (camposgrid(i) <> "")
        Rem DATOS DE LA COLUMNA
        FORMATOGRILLA(1, i + 1) = ""
        i = i + 1
    Wend
    Rem LARGO DE LOS DATOS
    FORMATOGRILLA(2, 1) = "6"
    'formatogrilla(2, 2) = "25"
    Rem TIPO DE DATOS C=CEROS,S=STRING,N=NUMERICO,D=DATE
    FORMATOGRILLA(3, 1) = "S"
    FORMATOGRILLA(3, 2) = "S"
    FORMATOGRILLA(3, 3) = "N"
    FORMATOGRILLA(3, 4) = "N"
    FORMATOGRILLA(3, 5) = "N"
    'formatogrilla(3, 2) = "S"
    Rem FORMATO GRILLA
    For i = 1 To 1
        FORMATOGRILLA(4, i) = ""
    Next i
    Rem LOCCKED
    For i = 1 To 6
        FORMATOGRILLA(5, i) = "FALSE"
    Next i
    Rem ancho
    FORMATOGRILLA(6, 1) = "9"
    FORMATOGRILLA(6, 2) = "28"
    FORMATOGRILLA(6, 3) = "19"
    FORMATOGRILLA(6, 4) = "5"
    FORMATOGRILLA(6, 5) = "5"
    With grilla
        .Cols = numCol
        .Rows = numRow
        .AllowUserResizing = False
        .DisplayFocusRect = False
        .ExtendLastCol = False
        .BoldFixedCell = False
        .DrawMode = cellOwnerDraw
        .Appearance = Flat
        .ScrollBarStyle = Flat
        .FixedRowColStyle = Flat
        For k = 1 To numCol - 1
            .Cell(0, k).text = FORMATOGRILLA(1, k)
            .Column(k).Width = Val(FORMATOGRILLA(6, k)) * .Cell(0, k).Font.Size + 1.25
         
            .Column(k).MaxLength = Val(FORMATOGRILLA(2, k))
            .Column(k).FormatString = FORMATOGRILLA(4, k)
            .Column(k).Locked = FORMATOGRILLA(5, k)
            If FORMATOGRILLA(3, k) = "N" Then
                .Column(k).Alignment = cellCenterCenter
                .Column(k).Mask = cellNumeric
            End If
            If FORMATOGRILLA(3, k) = "S" Then
                .Column(k).Alignment = cellLeftCenter
                .Column(k).Mask = cellUpper
            End If
        Next k
        .Range(0, 0, 0, .Cols - 1).Alignment = cellCenterCenter
    .Column(0).Width = 0
    
    
    End With '//grilla

End Sub


Sub CargaGridLiquida(grilla As Grid)
    Dim resultados As rdoResultset
    Dim csql As New rdoQuery
    Dim LINEA As Integer
    Dim LIQUIDO As Double
    Dim CONVENIO As Double
    Dim saldo As Double
    Dim TOTAL1 As Double
    Dim total2 As Double
    Dim total3 As Double
    Dim DIFERENCIAcaja As Double
    
    Set csql.ActiveConnection = contadb
    csql.sql = "SELECT rut,nombre,crcc "
    
    csql.sql = csql.sql + " FROM " & clientesistema & "remu" & empresaactiva & ".mt_fijo where mes='" + Format(fechasistema, "mm") + "' and año='" + Format(fechasistema, "yyyy") + "' and mid(fecharet,1,7)<>'" + Format(fechasistema, "yyyy-mm") + "' "
    csql.sql = csql.sql + " ORDER BY nombre "
    csql.Execute
    barra.Value = 1
    TOTAL1 = 0
    total2 = 0
    total3 = 0
    
    grilla.Rows = 1
    If csql.RowsAffected > 0 Then
    barra.Max = csql.RowsAffected + 1
    
        Set resultados = csql.OpenResultset
        
        LINEA = 1
        While Not resultados.EOF
'            Call calculaliquidaciones.calculaliquidaciones(resultados(0), Format(fechasistema, "mm"), Format(fechasistema, "YYYY"), empresaactiva)
            LIQUIDO = leercalculo(resultados(0), Format(fechasistema, "mm"), Format(fechasistema, "YYYY"), "LI001")
            If LIQUIDO <> 0 Then
            grilla.Rows = grilla.Rows + 1
            LINEA = grilla.Rows - 1
            grilla.Cell(LINEA, 1).text = resultados(0)
            grilla.Cell(LINEA, 2).text = resultados(1)
            CONVENIO = totalconvenios(resultados(0))
            saldo = LIQUIDO - CONVENIO
            DIFERENCIAcaja = leerdiferencia(resultados(0), Format(fechasistema, "mm"), Format(fechasistema, "YYYY"), "00019")
            grilla.Cell(LINEA, 3).text = Format(LIQUIDO, "###,###,###")
            grilla.Cell(LINEA, 4).text = Format(CONVENIO, "###,###,###")
            grilla.Cell(LINEA, 5).text = Format(saldo, "###,###,###")
            grilla.Cell(LINEA, 6).text = Format(DIFERENCIAcaja, "###,###,###")
            
            TOTAL1 = TOTAL1 + LIQUIDO
            total2 = total2 + CONVENIO
            total3 = total3 + saldo
            End If
            barra.Value = barra.Value + 1
            resultados.MoveNext
            LINEA = LINEA + 1
        Wend
        resultados.Close
        Set resultados = Nothing
    End If
    grilla.Rows = grilla.Rows + 1
    grilla.Range(grilla.Rows - 1, 1, grilla.Rows - 1, grilla.Cols - 1).Borders(cellEdgeTop) = cellThin
    grilla.Range(grilla.Rows - 1, 1, grilla.Rows - 1, grilla.Cols - 1).FontBold = True
    grilla.Cell(grilla.Rows - 1, 2).text = "TOTAL SUELDOS A PAGAR  "
    grilla.Cell(grilla.Rows - 1, 3).text = Format(TOTAL1, "#,###,###,###")
    grilla.Cell(grilla.Rows - 1, 4).text = Format(total2, "#,###,###,###")
    grilla.Cell(grilla.Rows - 1, 5).text = Format(total3, "#,###,###,###")
    
    
    grilla.AutoRedraw = True
    grilla.Refresh
    barra.Max = 1
    
End Sub

Sub CargaGrilla1(numRow, numCol, grilla As Grid, camposgrid As Variant)
    Dim FORMATOGRILLA(50, 50) As String
    
    Dim i As Integer
    i = 0
    While (camposgrid(i) <> "")
     
        Rem DATOS DE LA COLUMNA
        FORMATOGRILLA(1, i + 1) = camposgrid(i)
        i = i + 1
    Wend
    Rem LARGO DE LOS DATOS
    FORMATOGRILLA(2, 1) = "6"
    'formatogrilla(2, 2) = "25"
    Rem TIPO DE DATOS C=CEROS,S=STRING,N=NUMERICO,D=DATE
    FORMATOGRILLA(3, 1) = "S"
    FORMATOGRILLA(3, 2) = "S"
    FORMATOGRILLA(3, 3) = "N"
    FORMATOGRILLA(3, 4) = "N"
    FORMATOGRILLA(3, 5) = "N"
    FORMATOGRILLA(3, 6) = "N"
    Rem FORMATO GRILLA
    For i = 1 To 1
        FORMATOGRILLA(4, i) = ""
    Next i
    Rem LOCCKED
    For i = 1 To 7
        FORMATOGRILLA(5, i) = "TRUE"
    Next i
    Rem ancho
    FORMATOGRILLA(6, 1) = "10"
    FORMATOGRILLA(6, 2) = "33"
    FORMATOGRILLA(6, 3) = "10"
    FORMATOGRILLA(6, 4) = "10"
    FORMATOGRILLA(6, 5) = "10"
    FORMATOGRILLA(6, 6) = "15"
    With grilla
        .Cols = numCol
        .Rows = numRow
        .AllowUserResizing = False
        .DisplayFocusRect = False
        .ExtendLastCol = False
        .BoldFixedCell = False
        .DrawMode = cellOwnerDraw
        .Appearance = Flat
        .ScrollBarStyle = Flat
        .FixedRowColStyle = Flat
        For k = 1 To numCol - 1
            .Cell(0, k).text = FORMATOGRILLA(1, k)
            .Column(k).Width = Val(FORMATOGRILLA(6, k)) * .Cell(0, k).Font.Size + 1.25
         
            .Column(k).MaxLength = Val(FORMATOGRILLA(2, k))
            .Column(k).FormatString = FORMATOGRILLA(4, k)
            .Column(k).Locked = True 'formatogrilla(5, K)
            If FORMATOGRILLA(3, k) = "N" Then
                .Column(k).Alignment = cellCenterCenter
                .Column(k).Mask = cellNumeric
            End If
            If FORMATOGRILLA(3, k) = "S" Then
                .Column(k).Alignment = cellLeftCenter
                .Column(k).Mask = cellUpper
            End If
        Next k
        .Range(0, 0, 0, .Cols - 1).Alignment = cellCenterCenter
    .Column(0).Width = 0
    
    
    End With '//grilla

End Sub

Sub IniciaGrid23()
    Dim CABEZAS2 As Variant
    
    CABEZAS2 = Array("HABERES", "BASE", "MONTO", "DESCUENTOS", "BASE", "MONTO", "")
    Call CargaGrilla23(1, 7, GRILLALIQUIDACION, CABEZAS2)
    'cabezas2 = Array("DESCUENTOS", "BASE CAL.", "MONTO", "")
    'Call CargaGrilla23(1, 4, GridDescuentos, cabezas2)
End Sub

Sub CargaGrilla23(numRow, numCol, grilla As Grid, camposgrid As Variant)
    Dim i As Integer
    Dim FORMATOGRILLA(50, 50) As String
    
    i = 0
    While (camposgrid(i) <> "")
        FORMATOGRILLA(1, i + 1) = camposgrid(i) 'encabezados
        i = i + 1
    Wend
    Rem LARGO DE LOS DATOS
    For i = 1 To 6
        FORMATOGRILLA(2, i) = "10"
    Next i
    Rem TIPO DE DATOS C=CEROS,S=STRING,N=NUMERICO,D=DATE,
    FORMATOGRILLA(3, 1) = "S"
    FORMATOGRILLA(3, 2) = "N"
    FORMATOGRILLA(3, 3) = "N"
    FORMATOGRILLA(3, 4) = "S"
    FORMATOGRILLA(3, 5) = "N"
    FORMATOGRILLA(3, 6) = "N"
    Rem FORMATO GRILLA
    For i = 1 To 9
        FORMATOGRILLA(4, i) = ""
    Next i
    Rem LOCCKED
    For i = 1 To 9
        FORMATOGRILLA(5, i) = "FALSE"
    Next i
    Rem ancho
    FORMATOGRILLA(6, 1) = "30"
    FORMATOGRILLA(6, 2) = "7"
    FORMATOGRILLA(6, 3) = "10"
    FORMATOGRILLA(6, 4) = "30"
    FORMATOGRILLA(6, 5) = "7"
    FORMATOGRILLA(6, 6) = "10"
    grilla.Column(0).Width = 0
    
    With grilla
        .Cols = numCol
        .Rows = numRow
        .AllowUserResizing = False
        .DisplayFocusRect = False
        .ExtendLastCol = False
        .BoldFixedCell = False
        .DrawMode = cellOwnerDraw
        .Appearance = Flat
        .ScrollBarStyle = Flat
        .FixedRowColStyle = Flat
        .Column(0).Width = 0
        For k = 1 To numCol - 1
            .Cell(0, k).text = FORMATOGRILLA(1, k)
            .Column(k).Width = Val(FORMATOGRILLA(6, k)) * .Cell(0, k).Font.Size + 1.25
            .Column(0).Width = 0
            .Column(k).MaxLength = Val(FORMATOGRILLA(2, k))
            .Column(k).FormatString = FORMATOGRILLA(4, k)
            .Column(k).Locked = True 'formatogrilla(5, K)
            If FORMATOGRILLA(3, k) = "N" Then
                .Column(k).Alignment = cellRightCenter
                .Column(k).Mask = cellNumeric
            End If
            If FORMATOGRILLA(3, k) = "S" Then
                .Column(k).Alignment = cellLeftCenter
                .Column(k).Mask = cellUpper
            End If
            If FORMATOGRILLA(3, k) = "D" Then
                .Column(k).CellType = cellCalendar
                .Column(k).Mask = cellNumeric
            End If
            '.Column(7).CellType = cellComboBox
        Next k
        '.Range(0, 1, 0, 3).Merge
        '.Cell(0, 1).text = "CUENTA"
        .Range(0, 0, 0, .Cols - 1).Alignment = cellCenterCenter
    End With '//grilla
End Sub
'*esta es*/
'Sub liquidacion()
'    Dim W As Integer
'    Dim contadord As Integer
'    Dim contadorh As Integer
'    GRILLALIQUIDACION.Rows = 1
'
'    GRILLALIQUIDACION.Rows = 18
'    contadord = 0
'    contadorh = 0
'
'    For W = 1 To LINEAC
'        Select Case CALCULOS(W, 6)
'        Case "H"
'            If CALCULOS(W, 7) = "S" Or CALCULOS(W, 7) = "T" Then
'                contadord = contadord + 1
'                If CALCULOS(W, 1) = "TOTAL HABERES GENERALES " Then
'                contadord = 13
'                End If
'
'                GRILLALIQUIDACION.Cell(contadord, 1).text = CALCULOS(W, 1)
'                If CALCULOS(W, 2) <> 0 Then
'                GRILLALIQUIDACION.Cell(contadord, 2).text = CALCULOS(W, 2)
'                End If
'                GRILLALIQUIDACION.Cell(contadord, 3).text = Format(CALCULOS(W, 3), "###,###,###")
'                If CALCULOS(W, 7) = "T" Then
'                GRILLALIQUIDACION.Range(contadord, 1, contadord, 3).FontBold = True
'                End If
'
'
'            End If
'        Case "D"
'            If CALCULOS(W, 7) = "S" Or CALCULOS(W, 7) = "T" Or CALCULOS(W, 7) = "B" Then
'                contadorh = contadorh + 1
'                If CALCULOS(W, 1) = "TOTAL DESCUENTOS GENERALES " Then
'                contadorh = 13
'                End If
'                If Mid(CALCULOS(W, 1), 1, 7) = "LIQUIDO" Then
'                contadorh = 15
'                End If
'                GRILLALIQUIDACION.Cell(contadorh, 4).text = CALCULOS(W, 1)
'                If CALCULOS(W, 2) <> 0 Then
'
'                GRILLALIQUIDACION.Cell(contadorh, 5).text = CALCULOS(W, 2)
'                End If
'                GRILLALIQUIDACION.Cell(contadorh, 6).text = Format(CALCULOS(W, 3), "###,###,###")
'
'            If CALCULOS(W, 7) = "T" Then
'                GRILLALIQUIDACION.Range(contadorh, 4, contadorh, 6).FontBold = True
'                End If
'
'
'            End If
'        End Select
'    Next W
'GRILLALIQUIDACION.Range(13, 1, 13, 6).Borders(cellEdgeBottom) = cellThin
'GRILLALIQUIDACION.Range(13, 1, 13, 6).Borders(cellEdgeLeft) = cellThin
'GRILLALIQUIDACION.Range(13, 1, 13, 6).Borders(cellEdgeTop) = cellThin
'GRILLALIQUIDACION.Range(13, 1, 13, 6).Borders(cellEdgeRight) = cellThin
'GRILLALIQUIDACION.Range(13, 3, 13, 3).Borders(cellEdgeRight) = cellThin
'For k = 1 To 6
'GRILLALIQUIDACION.Column(k).Locked = False
'Next k
'
'GRILLALIQUIDACION.Range(16, 1, 16, 6).Merge
'GRILLALIQUIDACION.Cell(16, 1).text = "SON :" + WORDNUM(GRILLALIQUIDACION.Cell(15, 6).text, "PESO", "PESOS", 0)
'For k = 1 To 6
'GRILLALIQUIDACION.Column(k).Locked = True
'Next k
'GRILLALIQUIDACION.Cell(17, 4).text = "RECIBI CONFORME ____________________ "
'
'End Sub

'
'Sub liquidacion()
'    Dim W As Integer
'    Dim contadord As Integer
'    Dim contadorh As Integer
'    GRILLALIQUIDACION.Rows = 1
'
'    GRILLALIQUIDACION.Rows = 18
'    contadord = 0
'    contadorh = 0
'
'    For W = 1 To LINEAC
'        Select Case CALCULOS(W, 6)
'        Case "H"
'            If CALCULOS(W, 7) = "S" Or CALCULOS(W, 7) = "T" Then
'                contadord = contadord + 1
'                If CALCULOS(W, 1) = "TOTAL HABERES GENERALES " Then
'                contadord = 13
'                End If
'
'                GRILLALIQUIDACION.Cell(contadord, 1).text = CALCULOS(W, 1)
'                If CALCULOS(W, 2) <> 0 Then
'                GRILLALIQUIDACION.Cell(contadord, 2).text = CALCULOS(W, 2)
'                End If
'                GRILLALIQUIDACION.Cell(contadord, 3).text = Format(CALCULOS(W, 3), "###,###,###")
'                If CALCULOS(W, 7) = "T" Then
'                GRILLALIQUIDACION.Range(contadord, 1, contadord, 3).FontBold = True
'                End If
'
'
'            End If
'        Case "D"
'            If CALCULOS(W, 7) = "S" Or CALCULOS(W, 7) = "T" Then
'                contadorh = contadorh + 1
'                If CALCULOS(W, 1) = "TOTAL DESCUENTOS GENERALES " Then
'                contadorh = 13
'                End If
'
'                GRILLALIQUIDACION.Cell(contadorh, 4).text = CALCULOS(W, 1)
'                If CALCULOS(W, 2) <> 0 Then
'
'                GRILLALIQUIDACION.Cell(contadorh, 5).text = CALCULOS(W, 2)
'                End If
'                GRILLALIQUIDACION.Cell(contadorh, 6).text = Format(CALCULOS(W, 3), "###,###,###")
'
'            If CALCULOS(W, 7) = "T" Then
'                GRILLALIQUIDACION.Range(contadorh, 4, contadorh, 6).FontBold = True
'                End If
'
'
'            End If
'        End Select
'    Next W
'GRILLALIQUIDACION.Range(13, 1, 13, 6).Borders(cellEdgeBottom) = cellThin
'GRILLALIQUIDACION.Range(13, 1, 13, 6).Borders(cellEdgeLeft) = cellThin
'GRILLALIQUIDACION.Range(13, 1, 13, 6).Borders(cellEdgeTop) = cellThin
'GRILLALIQUIDACION.Range(13, 1, 13, 6).Borders(cellEdgeRight) = cellThin
'GRILLALIQUIDACION.Range(13, 3, 13, 3).Borders(cellEdgeRight) = cellThin
'For W = 1 To 6
'GRILLALIQUIDACION.Column(W).Locked = False
'Next W
'
'GRILLALIQUIDACION.Range(15, 1, 15, 6).Merge
'GRILLALIQUIDACION.Cell(15, 1).text = "SON :" + WORDNUM(GRILLALIQUIDACION.Cell(14, 6).text, "PESO", "PESOS", 0)
'For W = 1 To 6
'GRILLALIQUIDACION.Column(W).Locked = True
'Next W
'GRILLALIQUIDACION.Cell(17, 4).text = "RECIBI CONFORME ____________________ "
'
'End Sub
'
'Sub liquidacion()
'    Dim W As Integer
'    Dim contadord As Integer
'    Dim contadorh As Integer
'    GRILLALIQUIDACION.Rows = 1
'
'    GRILLALIQUIDACION.Rows = 15
'    contadord = 0
'    contadorh = 0
'
'    For W = 1 To LINEAC
'        Select Case CALCULOS(W, 6)
'        Case "H"
'            If CALCULOS(W, 7) = "S" Or CALCULOS(W, 7) = "T" Then
'                contadord = contadord + 1
'                If CALCULOS(W, 1) = "TOTAL HABERES GENERALES " Then
'                contadord = 13
'                End If
'
'                GRILLALIQUIDACION.Cell(contadord, 1).text = CALCULOS(W, 1)
'                If CALCULOS(W, 2) <> 0 Then
'                GRILLALIQUIDACION.Cell(contadord, 2).text = CALCULOS(W, 2)
'                End If
'                GRILLALIQUIDACION.Cell(contadord, 3).text = Format(CALCULOS(W, 3), "###,###,###")
'                If CALCULOS(W, 7) = "T" Then
'                GRILLALIQUIDACION.Range(contadord, 1, contadord, 3).FontBold = True
'                End If
'
'
'            End If
'        Case "D"
'            If CALCULOS(W, 7) = "S" Or CALCULOS(W, 7) = "T" Then
'                contadorh = contadorh + 1
'                If CALCULOS(W, 1) = "TOTAL DESCUENTOS GENERALES " Then
'                contadorh = 13
'                End If
'
'                GRILLALIQUIDACION.Cell(contadorh, 4).text = CALCULOS(W, 1)
'                If CALCULOS(W, 2) <> 0 Then
'
'                GRILLALIQUIDACION.Cell(contadorh, 5).text = CALCULOS(W, 2)
'                End If
'                GRILLALIQUIDACION.Cell(contadorh, 6).text = Format(CALCULOS(W, 3), "###,###,###")
'
'            If CALCULOS(W, 7) = "T" Then
'                GRILLALIQUIDACION.Range(contadorh, 4, contadorh, 6).FontBold = True
'                End If
'
'
'            End If
'        End Select
'    Next W
'GRILLALIQUIDACION.Range(13, 1, 13, 6).Borders(cellEdgeBottom) = cellThin
'
'GRILLALIQUIDACION.Range(13, 1, 13, 6).Borders(cellEdgeLeft) = cellThin
'GRILLALIQUIDACION.Range(13, 1, 13, 6).Borders(cellEdgeTop) = cellThin
'GRILLALIQUIDACION.Range(13, 1, 13, 6).Borders(cellEdgeRight) = cellThin
'GRILLALIQUIDACION.Range(13, 3, 13, 3).Borders(cellEdgeRight) = cellThin
'GRILLALIQUIDACION.Cell(14, 1).text = WORDNUM(Format(dato6.text, "########0"), "PESO", "PESOS", 0)
'
'
'
'End Sub

