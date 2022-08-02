VERSION 5.00
Object = "{4777436C-EB8C-4596-98A8-EBCF98683969}#1.0#0"; "FlexCell.ocx"
Object = "{ADBBAED6-B16F-40DC-80DF-B44910CBA76C}#1.0#0"; "Frame-Xp.ocx"
Begin VB.Form flujocaja 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "FLUJO DE CAJA"
   ClientHeight    =   9630
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   12765
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9630
   ScaleWidth      =   12765
   StartUpPosition =   2  'CenterScreen
   Begin XPFrame.FrameXp FrameXp2 
      Height          =   1215
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   12495
      _ExtentX        =   22040
      _ExtentY        =   2143
      BackColor       =   16761024
      Caption         =   "Datos Flujo"
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
      Begin VB.CommandButton Command1 
         Caption         =   "LEER FLUJO"
         Height          =   375
         Left            =   1920
         TabIndex        =   7
         Top             =   480
         Width           =   1575
      End
      Begin VB.TextBox dato3 
         BackColor       =   &H00E1FFFD&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   240
         MaxLength       =   2
         TabIndex        =   5
         Tag             =   "fecha"
         Top             =   600
         Width           =   375
      End
      Begin VB.TextBox dato4 
         BackColor       =   &H00E1FFFD&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   600
         MaxLength       =   2
         TabIndex        =   4
         Tag             =   "fecha"
         Top             =   600
         Width           =   375
      End
      Begin VB.TextBox dato5 
         BackColor       =   &H00E1FFFD&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   960
         MaxLength       =   4
         TabIndex        =   3
         Tag             =   "fecha"
         Top             =   600
         Width           =   615
      End
      Begin VB.Label lblempresa 
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4080
         TabIndex        =   8
         Top             =   480
         Width           =   8055
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "SEMANA"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   360
         Width           =   1335
      End
   End
   Begin XPFrame.FrameXp FrameXp1 
      Height          =   8055
      Left            =   120
      TabIndex        =   0
      Top             =   1440
      Width           =   12495
      _ExtentX        =   22040
      _ExtentY        =   14208
      BackColor       =   16761024
      Caption         =   "Flujo"
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
      Begin VB.CommandButton Command2 
         Caption         =   "imprimir"
         Height          =   375
         Left            =   5280
         TabIndex        =   9
         Top             =   7440
         Width           =   1455
      End
      Begin FlexCell.Grid Grid1 
         Height          =   7095
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   12255
         _ExtentX        =   21616
         _ExtentY        =   12515
         Cols            =   5
         DefaultFontSize =   8.25
         Rows            =   30
      End
   End
End
Attribute VB_Name = "flujocaja"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim totales(31) As Double
Dim totales2(31) As Double
Dim SALDOANTE(31) As Double
Dim fechalunes As String
Dim contador As Double
Dim fechaviernes As String




Private Sub Command1_Click()

fechaflujo = DATO5.text + "-" + dato4.text + "-" + dato3.text
dia = Weekday(fechaflujo, vbMonday)
If dia <> 1 Then
    contador = 0
    For k = dia To 7
    contador = contador + 1
    fechalunes = DateAdd("d", contador, fechaflujo)
    Next k
Else
    fechalunes = fechaflujo

End If

fechaviernes = DateAdd("d", 4, fechalunes)
dato3.text = Format(fechalunes, "dd")
dato4.text = Format(fechalunes, "mm")
DATO5.text = Format(fechalunes, "yyyy")




CARGAGRILLA

LEERFLUJO
End Sub

Private Sub COMMAND2_Click()
imprimir

End Sub
Sub imprimir()
Dim titulo As String
titulo = "LISTADO FLUJO DE CAJA"
Call CABEZAS2(titulo, "N", "000000000")
Grid1.Range(0, 1, 0, Grid1.Cols - 1).Borders(cellEdgeBottom) = cellThick
Grid1.Range(0, 1, 0, Grid1.Cols - 1).Borders(cellEdgeLeft) = cellThick
Grid1.Range(0, 1, 0, Grid1.Cols - 1).Borders(cellEdgeTop) = cellThick
Grid1.Range(0, 1, 0, Grid1.Cols - 1).Borders(cellEdgeRight) = cellThick
Grid1.Range(0, 1, 0, Grid1.Cols - 1).Borders(cellInsideHorizontal) = cellThick
Grid1.Range(0, 1, 0, Grid1.Cols - 1).Borders(cellInsideVertical) = cellThick
Grid1.DefaultFont.Size = 8
Grid1.PageSetup.Orientation = cellPortrait
Grid1.PageSetup.PrintFixedRow = True
Grid1.PageSetup.BottomMargin = 2
Grid1.PageSetup.TopMargin = 1
Grid1.PageSetup.LeftMargin = 1
Grid1.PageSetup.RightMargin = 0
Grid1.PageSetup.BlackAndWhite = True
Grid1.PageSetup.PrintGridlines = False
Grid1.PrintPreview 100

   
End Sub
Sub CABEZAS2(titulo, tipo, FOLIO)
Dim objReportTitle As FlexCell.ReportTitle
Grid1.ReportTitles.Clear


Set objReportTitle = New FlexCell.ReportTitle
    objReportTitle.text = titulo
    objReportTitle.Font.Name = "arial"
    objReportTitle.Font.Size = 12
    objReportTitle.PrintOnAllPages = True
    Grid1.ReportTitles.Add objReportTitle

Set objReportTitle = New FlexCell.ReportTitle
    objReportTitle.text = nombrebanco
    
    objReportTitle.Font.Name = "arial"
    objReportTitle.Font.Size = 12
    objReportTitle.PrintOnAllPages = True
    Grid1.ReportTitles.Add objReportTitle
    
    'Report Title 1
    If tipo = "N" Then
        For k = 1 To 1
        Set objReportTitle = New FlexCell.ReportTitle
        objReportTitle.text = LBLEMPRESA.Caption
        objReportTitle.Font.Name = "arial"
        objReportTitle.Font.Size = 8
        objReportTitle.Font.Italic = True
        objReportTitle.PrintOnAllPages = True
        'objReportTitle.Color = RGB(128, 0, 0)
        objReportTitle.Align = CellLeft
        Grid1.ReportTitles.Add objReportTitle
    Next k
    Else
        For k = 1 To 4
        Set objReportTitle = New FlexCell.ReportTitle
        objReportTitle.text = ""
        objReportTitle.Font.Name = "arial"
        objReportTitle.Font.Size = 8
        objReportTitle.Font.Italic = True
        objReportTitle.PrintOnAllPages = True
        'objReportTitle.Color = RGB(128, 0, 0)
        objReportTitle.Align = CellLeft
        Grid1.ReportTitles.Add objReportTitle
        
        Next k
    Set objReportTitle = New FlexCell.ReportTitle
        
        
        
        
        
        objReportTitle.text = ""
        objReportTitle.Font.Name = "arial"
        objReportTitle.Font.Size = 8
        objReportTitle.Font.Italic = True
        objReportTitle.PrintOnAllPages = True
        'objReportTitle.Color = RGB(128, 0, 0)
        objReportTitle.Align = CellLeft
        Grid1.ReportTitles.Add objReportTitle
        
    End If
    
With Grid1.PageSetup
        
        If tipo = "N" Then .Header = "Pagina &P de &N Emitido: &D Usuario:" + USUARIOSISTEMA
        Rem If TIPO = "S" Then .Footer = "pagina &P"
        
        .HeaderAlignment = cellCenter
        
        .HeaderFont.Name = "Verdana"
        .HeaderFont.Size = 7
        .HeaderMargin = 2
        .TopMargin = 1
        .BottomMargin = 2
        
        
        
End With

End Sub

Private Sub dato3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
dato4.SetFocus
End If
End Sub
Private Sub dato4_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
DATO5.SetFocus
End If
End Sub
Private Sub dato5_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Command1.SetFocus

End If
End Sub

Private Sub Form_Activate()
Command1_Click

End Sub

Private Sub Form_Load()

dia = Weekday(fechaflujo, vbMonday)
If dia <> 1 Then
    contador = 0
    For k = dia To 7
    contador = contador + 1
    fechalunes = DateAdd("d", contador, fechaflujo)
    Next k
Else
    fechalunes = fechaflujo

End If

fechaviernes = DateAdd("d", 4, fechalunes)
dato3.text = Format(fechalunes, "dd")
dato4.text = Format(fechalunes, "mm")
DATO5.text = Format(fechalunes, "yyyy")




CARGAGRILLA

LEERFLUJO

    
End Sub
Sub CARGAGRILLA()

Rem DATOS DE LA COLUMNA
    Dim FORMATOGRILLA(10, 40)
    Grid1.DefaultFont.Size = 8
       
    FORMATOGRILLA(1, 1) = "CODIGO"
    FORMATOGRILLA(1, 2) = "EMPRESA"
    
    For k = 0 To 30
    FORMATOGRILLA(1, 3 + k) = DateAdd("d", k, fechalunes)
    Next k
    Rem LARGO DE LOS DATOS
    FORMATOGRILLA(2, 1) = "6"
    FORMATOGRILLA(2, 2) = "30"
    For k = 3 To 31
    FORMATOGRILLA(2, k) = "10"
    Next k
    Rem TIPO DE DATOS C=CEROS,S=STRING,N=NUMERICO,D=DATE,
    FORMATOGRILLA(3, 1) = "S"
    FORMATOGRILLA(3, 2) = "S"
    FORMATOGRILLA(3, 3) = "N"
    FORMATOGRILLA(3, 4) = "N"
    FORMATOGRILLA(3, 5) = "N"
    FORMATOGRILLA(3, 6) = "N"
    FORMATOGRILLA(3, 7) = "N"
    For k = 1 To 31
    FORMATOGRILLA(3, k + 2) = "N"
    Next k
    
    FORMATOGRILLA(3, 8) = "N"
    Rem FORMATO GRILLA
    For k = 1 To 31
    FORMATOGRILLA(4, k + 2) = "###,###,##0"
    Next k
    Rem LOCCKED
    For k = 1 To 8
    FORMATOGRILLA(5, k) = "TRUE"
    Next k
    For k = 1 To 31
    FORMATOGRILLA(5, k + 2) = "true"
    Next k
    
    Grid1.Cols = 34
    Grid1.Rows = 1
    Grid1.AllowUserResizing = False
    Grid1.DisplayFocusRect = False
    Grid1.ExtendLastCol = True
    Grid1.BoldFixedCell = False
    Grid1.DrawMode = cellOwnerDraw
    
    Grid1.Appearance = Flat
    Grid1.ScrollBarStyle = Flat
    Grid1.FixedRowColStyle = Flat
    
'   Grid1.BackColorFixed = RGB(90, 158, 214)
'   Grid1.BackColorFixedSel = RGB(110, 180, 230)
'   Grid1.BackColorBkg = RGB(90, 158, 214)
'   Grid1.BackColorScrollBar = RGB(231, 235, 247)
'   Grid1.BackColor1 = RGB(231, 235, 247)
'   Grid1.BackColor2 = RGB(239, 243, 255)
'   Grid1.GridColor = RGB(148, 190, 231)
   Grid1.Column(0).Width = 0
    
    For k = 1 To Grid1.Cols - 1
        
        Grid1.Cell(0, k).text = FORMATOGRILLA(1, k)
        Grid1.Column(k).Width = Val(FORMATOGRILLA(2, k)) * Grid1.DefaultFont.Size
        Grid1.Column(k).MaxLength = Val(FORMATOGRILLA(2, k))
        Grid1.Column(k).FormatString = FORMATOGRILLA(4, k)
        Grid1.Column(k).Locked = FORMATOGRILLA(5, k)
        If FORMATOGRILLA(3, k) = "N" Then Grid1.Column(k).Alignment = cellRightCenter
        If FORMATOGRILLA(3, k) = "D" Then Grid1.Column(k).CellType = cellCalendar
        
    Next k
   
   Grid1.Column(3).Mask = cellNumeric
   Grid1.Column(4).Mask = cellNumeric
   Grid1.Column(5).Mask = cellNumeric
   Grid1.Column(6).Mask = cellNumeric
   Grid1.Column(7).Mask = cellNumeric
   
   
   
   
    
    
End Sub
Sub LEERFLUJO()
    Dim resultados As rdoResultset
    Dim csql As New rdoQuery
    
        Set csql.ActiveConnection = conta
        csql.sql = "SELECT * "
        csql.sql = csql.sql + "FROM flujo_caja_titulos "
        csql.sql = csql.sql + "ORDER BY codigo,subcuenta "
        csql.Execute
        Grid1.Rows = 1
        For k = 1 To 5
        totales(k) = 0
        totales2(k) = 0
        SALDOANTE(k) = 0
        Next k
         Rem SALDO ANTERIOR
           
        
        
        If csql.RowsAffected > 0 Then
            Set resultados = csql.OpenResultset
            While Not resultados.EOF
                Grid1.Rows = Grid1.Rows + 1
                Grid1.Cell(Grid1.Rows - 1, 0).text = resultados(0)
                Grid1.Cell(Grid1.Rows - 1, 1).text = resultados(1)
                Grid1.Cell(Grid1.Rows - 1, 2).text = resultados(2)
                If resultados(3) <> "T" Then
                Call LEERFLUJOsemanal(fechalunes, Format(DateAdd("D", 31, fechalunes), "yyyy-mm-dd"), resultados(0) + resultados(1), resultados(3))
                End If
                If resultados(3) = "T" Then
                
                Grid1.Range(Grid1.Rows - 1, 1, Grid1.Rows - 1, Grid1.Cols - 1).FontBold = True
                Grid1.Range(Grid1.Rows - 1, 1, Grid1.Rows - 1, Grid1.Cols - 1).Borders(cellEdgeTop) = cellThin
                Grid1.Range(Grid1.Rows - 1, 1, Grid1.Rows - 1, Grid1.Cols - 1).Borders(cellEdgeBottom) = cellThin
                Grid1.Range(Grid1.Rows - 1, 1, Grid1.Rows - 1, Grid1.Cols - 1).Borders(cellEdgeLeft) = cellThin
                Grid1.Range(Grid1.Rows - 1, 1, Grid1.Rows - 1, Grid1.Cols - 1).Borders(cellEdgeRight) = cellThin
                Grid1.Range(Grid1.Rows - 1, 1, Grid1.Rows - 1, Grid1.Cols - 1).Borders(cellInsideHorizontal) = cellThin
                Grid1.Range(Grid1.Rows - 1, 1, Grid1.Rows - 1, Grid1.Cols - 1).Borders(cellInsideVertical) = cellThin
                For k = 1 To 5
                Grid1.Cell(Grid1.Rows - 1, k + 2).text = totales(k)
                totales(k) = 0
                Next k
                
                
                
                End If
                
                resultados.MoveNext
            Wend
            
 Rem
            Grid1.Rows = Grid1.Rows + 1
            Grid1.Cell(Grid1.Rows - 1, 2).text = "SALDO ANTERIOR"
            Grid1.Range(Grid1.Rows - 1, 1, Grid1.Rows - 1, Grid1.Cols - 1).FontBold = True
            Grid1.Range(Grid1.Rows - 1, 1, Grid1.Rows - 1, Grid1.Cols - 1).Borders(cellEdgeTop) = cellThin
            Grid1.Range(Grid1.Rows - 1, 1, Grid1.Rows - 1, Grid1.Cols - 1).Borders(cellEdgeBottom) = cellThin
            Grid1.Range(Grid1.Rows - 1, 1, Grid1.Rows - 1, Grid1.Cols - 1).Borders(cellEdgeLeft) = cellThin
            Grid1.Range(Grid1.Rows - 1, 1, Grid1.Rows - 1, Grid1.Cols - 1).Borders(cellEdgeRight) = cellThin
            Grid1.Range(Grid1.Rows - 1, 1, Grid1.Rows - 1, Grid1.Cols - 1).Borders(cellInsideHorizontal) = cellThin
            Grid1.Range(Grid1.Rows - 1, 1, Grid1.Rows - 1, Grid1.Cols - 1).Borders(cellInsideVertical) = cellThin
            For k = 2 To 4
            Grid1.Cell(Grid1.Rows - 1, k + 2).text = SALDOANTE(k)
            
            Next k
 
            Grid1.Rows = Grid1.Rows + 1
            Grid1.Cell(Grid1.Rows - 1, 2).text = "SALDO "
            Grid1.Range(Grid1.Rows - 1, 1, Grid1.Rows - 1, Grid1.Cols - 1).FontBold = True
            Grid1.Range(Grid1.Rows - 1, 1, Grid1.Rows - 1, Grid1.Cols - 1).Borders(cellEdgeTop) = cellThin
            Grid1.Range(Grid1.Rows - 1, 1, Grid1.Rows - 1, Grid1.Cols - 1).Borders(cellEdgeBottom) = cellThin
            Grid1.Range(Grid1.Rows - 1, 1, Grid1.Rows - 1, Grid1.Cols - 1).Borders(cellEdgeLeft) = cellThin
            Grid1.Range(Grid1.Rows - 1, 1, Grid1.Rows - 1, Grid1.Cols - 1).Borders(cellEdgeRight) = cellThin
            Grid1.Range(Grid1.Rows - 1, 1, Grid1.Rows - 1, Grid1.Cols - 1).Borders(cellInsideHorizontal) = cellThin
            Grid1.Range(Grid1.Rows - 1, 1, Grid1.Rows - 1, Grid1.Cols - 1).Borders(cellInsideVertical) = cellThin
            For k = 1 To 31
            totales2(k) = SALDOANTE(k) + totales2(k)
            Grid1.Cell(Grid1.Rows - 1, k + 2).text = totales2(k)
            If k < 31 Then
            Grid1.Cell(1, k + 3).text = totales2(k)
            totales2(k + 1) = totales2(k + 1) + totales2(k)
            End If
            totales(k) = 0
            totales2(k) = 0
            Next k
            Grid1.Rows = Grid1.Rows + 1
            Grid1.Cell(Grid1.Rows - 1, 2).text = "VALOR DE LA UF"
            For k = 1 To 31
            Grid1.Cell(Grid1.Rows - 1, 2 + k).text = leerUF(Grid1.Cell(0, k + 2).text)
            Next k
            
            resultados.Close
            Set resultados = Nothing
        
        End If
        
End Sub

Sub LEERFLUJOsemanal(fecha1, fecha2, codigo, tipo)
    Dim resultados As rdoResultset
    Dim csql As New rdoQuery
    Dim dia As Double
    
        Set csql.ActiveConnection = conta
        csql.sql = "SELECT fecha,sum(monto) "
        csql.sql = csql.sql + "FROM flujo_caja where fecha between '" + Format(fecha1, "yyyy-mm-dd") + "' and '" + Format(fecha2, "yyyy-mm-dd") + "' and tipo='" + codigo + "' and empresa='" + empresaflujo + "' "
        csql.sql = csql.sql + "GROUP BY fecha ORDER BY tipo  "
        csql.Execute
      
        If csql.RowsAffected > 0 Then
            Set resultados = csql.OpenResultset
            While Not resultados.EOF
                For k = 2 To 33
                If Format(resultados(0)) = Grid1.Cell(0, k).text Then
                dia = k - 2: Exit For
                
                End If
                Next k
                totales(dia) = totales(dia) + resultados(1)
                

                Grid1.Cell(Grid1.Rows - 1, 2 + dia).text = resultados(1)
                totales(dia) = totales(dia) + resultados(1)
                If tipo = "+" Then
                totales2(dia) = totales2(dia) + resultados(1)
                Else
                totales2(dia) = totales2(dia) - resultados(1)
                
                End If
                
                resultados.MoveNext
            Wend
            resultados.Close
            Set resultados = Nothing
        
        End If
        
End Sub

Private Sub Grid1_DblClick()
If Grid1.ActiveCell.row = Grid1.Rows - 1 And Grid1.ActiveCell.col > 2 Then
maestro05.Show vbModal
End If

If Grid1.Cell(Grid1.ActiveCell.row, 1).text > "0000" And Grid1.Cell(Grid1.ActiveCell.row, 1).text < "9999" Then
If Grid1.ActiveCell.row > 1 And Grid1.ActiveCell.row < Grid1.Rows - 1 Then
If Grid1.ActiveCell.col > 2 Then
fechaflujo = Format(Grid1.Cell(0, Grid1.ActiveCell.col).text, "yyyy-mm-dd")
tipoflujo = Grid1.Cell(Grid1.ActiveCell.row, 0).text + Grid1.Cell(Grid1.ActiveCell.row, 1).text
glosaflujo = Grid1.Cell(Grid1.ActiveCell.row, 2).text
flujocaja2.Show vbModal
End If
End If
End If
End Sub
