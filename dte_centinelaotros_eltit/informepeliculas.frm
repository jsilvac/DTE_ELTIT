VERSION 5.00
Object = "{4777436C-EB8C-4596-98A8-EBCF98683969}#1.0#0"; "FlexCell.ocx"
Object = "{ADBBAED6-B16F-40DC-80DF-B44910CBA76C}#1.0#0"; "Frame-Xp.ocx"
Begin VB.Form infopeliculas 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Informe de Peliculas En Arriendo"
   ClientHeight    =   8595
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   14940
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   573
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   996
   Begin VB.Frame ventas100 
      BackColor       =   &H00FAF0E7&
      Height          =   8475
      Left            =   0
      TabIndex        =   0
      Top             =   45
      Width           =   14880
      Begin XPFrame.FrameXp FrameXp4 
         Height          =   645
         Left            =   2280
         TabIndex        =   8
         Top             =   7680
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   1138
         BackColor       =   16761024
         Caption         =   "Filtro"
         CaptionEstilo3D =   1
         BackColor       =   16761024
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Begin VB.OptionButton Orden3 
            BackColor       =   &H00FFC0C0&
            Caption         =   "todas"
            Height          =   195
            Left            =   2520
            TabIndex        =   11
            Top             =   290
            Value           =   -1  'True
            Width           =   780
         End
         Begin VB.OptionButton orden2 
            BackColor       =   &H00FFC0C0&
            Caption         =   "Sin atraso"
            Height          =   195
            Left            =   1200
            TabIndex        =   10
            Top             =   280
            Width           =   1140
         End
         Begin VB.OptionButton orden1 
            BackColor       =   &H00FFC0C0&
            Caption         =   "Atrasadas"
            Height          =   285
            Left            =   120
            TabIndex        =   9
            Top             =   240
            Width           =   1215
         End
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H00FFC0C0&
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
         Height          =   465
         Left            =   360
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   7800
         Width           =   1725
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00FFC0C0&
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
         Height          =   465
         Left            =   8370
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   7785
         Width           =   1680
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FFC0C0&
         Caption         =   "GENERA INFORME"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   6000
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   7800
         Width           =   1725
      End
      Begin XPFrame.FrameXp FrameXp2 
         Height          =   7635
         Left            =   0
         TabIndex        =   6
         Top             =   0
         Width           =   14865
         _ExtentX        =   26220
         _ExtentY        =   13467
         BackColor       =   16773879
         Caption         =   "Periodo"
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
         Begin FlexCell.Grid Grid2 
            Height          =   7050
            Left            =   90
            TabIndex        =   7
            Top             =   315
            Width           =   14685
            _ExtentX        =   25903
            _ExtentY        =   12435
            BackColorFixed  =   16744576
            BackColorFixedSel=   16744576
            Cols            =   7
            DefaultFontName =   "Arial"
            DefaultFontSize =   9.75
            Rows            =   30
            SelectionMode   =   1
         End
      End
      Begin VB.Label etiqueta 
         BackColor       =   &H00FAF0E7&
         Height          =   375
         Left            =   240
         TabIndex        =   1
         Top             =   8280
         Width           =   3255
      End
   End
   Begin VB.Label titulo 
      BackColor       =   &H00FFC0C0&
      Height          =   375
      Left            =   9990
      TabIndex        =   2
      Top             =   10125
      Visible         =   0   'False
      Width           =   4815
   End
End
Attribute VB_Name = "infopeliculas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Private MODIFI As Integer
Sub grilla()

    
    Grid2.Rows = 1
    Grid2.Cols = 10
    Grid2.Column(0).Width = 0
    Grid2.Column(1).Width = 30
    Grid2.Column(2).Width = 100
    Grid2.Column(3).Width = 100
    Grid2.Column(4).Width = 30
    Grid2.Column(5).Width = 130
    Grid2.Column(6).Width = 280
    Grid2.Column(7).Width = 80
    Grid2.Column(8).Width = 80
    Grid2.Column(9).Width = 80

    Grid2.Cell(0, 1).text = "TP"
    Grid2.Cell(0, 2).text = "NUMERO"
    Grid2.Cell(0, 3).text = "FECHA"
    Grid2.Cell(0, 4).text = "LIN"
    Grid2.Cell(0, 5).text = "CODIGO"
    Grid2.Cell(0, 6).text = "DESCRIPCION"
    Grid2.Cell(0, 7).text = "CANTI."
    Grid2.Cell(0, 8).text = "PRECIO"
    Grid2.Cell(0, 9).text = "TOTAL"
    
    For k = 1 To 9
    Grid2.Column(k).Alignment = cellLeftCenter
    Next k
    Grid2.Column(7).Alignment = cellRightCenter
    Grid2.Column(8).Alignment = cellRightCenter
    Grid2.Column(9).Alignment = cellRightCenter
    
    Grid2.Column(7).FormatString = "$##,###,###"
    Grid2.Column(8).FormatString = "$##,###,###"
    Grid2.Column(9).FormatString = "$##,###,###"
    
For k = 1 To 9
Grid2.Range(0, k, 0, k).Borders(cellEdgeLeft) = cellThick
Grid2.Range(0, k, 0, k).Borders(cellEdgeTop) = cellThick
Grid2.Range(0, k, 0, k).Borders(cellEdgeRight) = cellThick
Grid2.Range(0, k, 0, k).Borders(cellEdgeBottom) = cellThick



Next k

    
End Sub
Sub LISTAVENTAS()
    Dim fechade As String
    Dim fechaha As String
    Dim linea As Double
    Dim empresabusca As String
    Dim resultados As rdoResultset
    Dim cSql As New rdoQuery
    Dim total As Double
    Dim tipo As String
    Dim total2 As Double
    
    
'    fechade = dato4.text + "-" + dato3.text + "-" + dato2.text
'    fechaha = dato7.text + "-" + dato6.text + "-" + dato5.text
    Set cSql.ActiveConnection = GESTIONrubro
    cSql.sql = "SELECT vd.tipo,vd.numero,vd.fecha,vd.linea,vd.codigo,vd.descripcion,vd.cantidad,vd.precio,vd.total "
    cSql.sql = cSql.sql + "FROM sv_ventas as vd "
    cSql.sql = cSql.sql + "where vd.entregado='0' and  vd.tipo='AR'  "
    If orden1.Value = True Then
       cSql.sql = cSql.sql + " and  vd.fechaentrega < '" + Format(fechasistema, "yyyy,mm,dd") + "' "
    End If
    If orden2.Value = True Then
    cSql.sql = cSql.sql + " and  vd.fechaentrega  >= '" + Format(fechasistema, "yyyy,mm,dd") + "' "
   
    End If
    
    cSql.sql = cSql.sql + "ORDER BY vd.tipo,vd.fecha "
    cSql.Execute
    Grid2.Rows = 1
    If cSql.RowsAffected > 0 Then
        Set resultados = cSql.OpenResultset
        tipo = resultados(0)
        Grid2.AutoRedraw = False
        Rem TOTAL = cSql.RowsAffected
        Grid2.Rows = cSql.RowsAffected + 2
        linea = 1
        total = 0
        total2 = 0
        While Not resultados.EOF
            linea = linea + 1
            
            If tipo <> resultados(0) Then
            
            Grid2.Cell(linea, 9).text = total
            Grid2.Cell(linea, 6).text = "TOTAL VENTA " + tipo
            Grid2.Range(linea, 6, linea, 8).Merge
            Grid2.Range(linea, 6, linea, 9).Borders(cellEdgeBottom) = cellThick
            Grid2.Range(linea, 6, linea, 9).Borders(cellEdgeTop) = cellThick
            Grid2.Range(linea, 6, linea, 9).Borders(cellEdgeRight) = cellThick
            Grid2.Range(linea, 6, linea, 9).Borders(cellEdgeLeft) = cellThick
            Grid2.Range(linea, 6, linea, 9).Borders(cellInsideVertical) = cellThick
            Grid2.Range(linea, 6, linea, 9).Borders(cellInsideHorizontal) = cellThick
            
            linea = linea + 1
            tipo = resultados(0)
            Grid2.Rows = Grid2.Rows + 1
            total = 0
            End If
            
            Grid2.Cell(linea, 1).text = resultados(0)
            Grid2.Cell(linea, 2).text = resultados(1)
            Grid2.Cell(linea, 3).text = resultados(2)
            Grid2.Cell(linea, 4).text = resultados(3)
            Grid2.Cell(linea, 5).text = resultados(4)
            Grid2.Cell(linea, 6).text = resultados(5)
            Grid2.Cell(linea, 7).text = resultados(6)
            Grid2.Cell(linea, 8).text = resultados(7)
            Grid2.Cell(linea, 9).text = resultados(8)
            
            
            total = total + resultados(8)
            total2 = total2 + resultados(8)
            
            resultados.MoveNext
        Wend
         If resultados.EOF = True Then
            linea = linea + 1
            Grid2.Rows = Grid2.Rows + 1
            Grid2.Cell(linea, 9).text = total
            Grid2.Cell(linea, 6).text = "TOTAL VENTA " + tipo
            total = 0
            Grid2.Range(linea, 6, linea, 8).Merge
            Grid2.Range(linea, 6, linea, 9).Borders(cellEdgeBottom) = cellThick
            Grid2.Range(linea, 6, linea, 9).Borders(cellEdgeTop) = cellThick
            Grid2.Range(linea, 6, linea, 9).Borders(cellEdgeRight) = cellThick
            Grid2.Range(linea, 6, linea, 9).Borders(cellEdgeLeft) = cellThick
            Grid2.Range(linea, 6, linea, 9).Borders(cellInsideVertical) = cellThick
            Grid2.Range(linea, 6, linea, 9).Borders(cellInsideHorizontal) = cellThick
            
            
            
            
            
            End If
            
        resultados.Close
        
            
        
        Set resultados = Nothing
        linea = linea + 1
            Grid2.Rows = Grid2.Rows + 1
            Grid2.Cell(linea, 9).text = total2
            Grid2.Cell(linea, 6).text = "TOTAL GENERAL  "
            
            Grid2.Range(linea, 6, linea, 8).Merge
            Grid2.Range(linea, 6, linea, 9).Borders(cellEdgeBottom) = cellThick
            Grid2.Range(linea, 6, linea, 9).Borders(cellEdgeTop) = cellThick
            Grid2.Range(linea, 6, linea, 9).Borders(cellEdgeRight) = cellThick
            Grid2.Range(linea, 6, linea, 9).Borders(cellEdgeLeft) = cellThick
            Grid2.Range(linea, 6, linea, 9).Borders(cellInsideVertical) = cellThick
            Grid2.Range(linea, 6, linea, 9).Borders(cellInsideHorizontal) = cellThick
            
        Grid2.AutoRedraw = True
        Grid2.Refresh
       
    Else
        
    End If

End Sub

Private Sub Command1_Click()
LISTAVENTAS
End Sub

Private Sub Command2_Click()
grilla
'dato2.SetFocus

End Sub

Private Sub Command3_Click()

Call Titulos("INFORME PELICULAS EN ARRIENDO")


Grid2.PageSetup.Orientation = cellLandscape




  Grid2.PageSetup.HeaderMargin = 0.5
  
  
        Grid2.PageSetup.PrintFixedRow = True
        
        
        
        
        
        
        
        
        Grid2.PageSetup.TopMargin = 1
        Grid2.PageSetup.LeftMargin = 0.5
        Grid2.PageSetup.RightMargin = 0.5
        Grid2.PageSetup.BottomMargin = 2
        
        Grid2.PageSetup.FooterMargin = 2
        Grid2.PageSetup.BlackAndWhite = True


Grid2.PrintPreview

End Sub


'Private Sub dato2_GotFocus()
'    dato2.SelStart = 0
'    dato2.SelLength = Len(dato2.text)
'End Sub
'
'Private Sub dato3_GotFocus()
'    dato3.SelStart = 0
'    dato3.SelLength = Len(dato3.text)
'End Sub
'
'Private Sub dato4_GotFocus()
'    dato4.SelStart = 0
'    dato4.SelLength = Len(dato4.text)
'End Sub
'
'Private Sub dato5_GotFocus()
'    dato5.SelStart = 0
'    dato5.SelLength = Len(dato5.text)
'End Sub
'
'Private Sub dato6_GotFocus()
'    dato6.SelStart = 0
'    dato6.SelLength = Len(dato6.text)
'End Sub
'
'Private Sub dato7_GotFocus()
'    dato7.SelStart = 0
'    dato7.SelLength = Len(dato7.text)
'End Sub
Private Sub Form_Load()
grilla
'dato2.text = Format(fechasistema, "dd")
'dato3.text = Format(fechasistema, "mm")
'dato4.text = Format(fechasistema, "yyyy")
'dato5.text = Format(fechasistema, "dd")
'dato6.text = Format(fechasistema, "mm")
'dato7.text = Format(fechasistema, "yyyy")

End Sub
Sub Titulos(titulo1)

    Dim i As Integer
    Dim objReportTitle As FlexCell.ReportTitle
    
    Grid2.FixedRowColStyle = Fixed3D
    Grid2.CellBorderColorFixed = vbButtonShadow
    Grid2.ShowResizeTips = False
    Grid2.ReportTitles.Clear
    Grid2.PageSetup.CenterHorizontally = True
    Grid2.PageSetup.Orientation = cellLandscape
    
    Grid2.PageSetup.PrintTitleRows = 1
    
    'Logo
'    Grid1.Images.Add App.path & "\Admin.gif", "Logo"
'    Set objReportTitle = New FlexCell.ReportTitle
'    objReportTitle.ImageKey = "Logo"
'    objReportTitle.Align = CellLeft
'    Grid1.ReportTitles.Add objReportTitle
    
    'ENCABEZADO DE PAGINA
    Grid2.PageSetup.Header = nombreempresa & vbCrLf & direccionempresa & vbCrLf & comunaempresa
    Grid2.PageSetup.HeaderAlignment = CellLeft
    Grid2.PageSetup.HeaderFont.Name = "Verdana"
    Grid2.PageSetup.HeaderFont.Size = 8
    
    'TITULOS DEL REPORTE
    Set objReportTitle = New FlexCell.ReportTitle
    objReportTitle.text = "LISTADO DE VENTAS Y ARRIENDOS"
    objReportTitle.Font.Name = "Verdana" '"Times New Roman"
    objReportTitle.Font.Size = 10
    objReportTitle.Font.Bold = True
    objReportTitle.Align = cellCenter
    objReportTitle.PrintOnAllPages = True
    Grid2.ReportTitles.Add objReportTitle
    
    Set objReportTitle = New FlexCell.ReportTitle
    objReportTitle.text = titulo1 & "  |  " & "PERIODO  :  " & Format(infocaja.DATO2.text & "-" & infocaja.DATO3.text & "-" & infocaja.dato4.text, "dd-mm-yyyy") & " al " & Format(infocaja.dato5.text & "-" & infocaja.dato6.text & "-" & infocaja.dato7.text, "dd-mm-yyyy")
    objReportTitle.Font.Name = "Verdana" '"Times New Roman"
    objReportTitle.Font.Size = 6
    objReportTitle.Font.Bold = False
    objReportTitle.Align = cellCenter
    objReportTitle.PrintOnAllPages = True
    Grid2.ReportTitles.Add objReportTitle
    
    'PIE DE PAGINA
    Grid2.PageSetup.Footer = "Pág &P de &N" & vbCrLf & "Fecha: &D"
    Grid2.PageSetup.FooterAlignment = cellRight
    Grid2.PageSetup.FooterFont.Name = "Verdana"
    Grid2.PageSetup.FooterFont.Size = 7
    
End Sub



'Private Sub dato2_KeyDown(KeyCode As Integer, Shift As Integer)
'    Call flechas(dato2, dato3, KeyCode)
'End Sub
'
'Private Sub dato3_KeyDown(KeyCode As Integer, Shift As Integer)
'    Call flechas(dato2, dato4, KeyCode)
'End Sub
'
'Private Sub dato4_KeyDown(KeyCode As Integer, Shift As Integer)
'    Call flechas(dato3, dato5, KeyCode)
'End Sub
'
'Private Sub dato5_KeyDown(KeyCode As Integer, Shift As Integer)
'    Call flechas(dato4, dato6, KeyCode)
'End Sub
'
'Private Sub dato6_KeyDown(KeyCode As Integer, Shift As Integer)
'    Call flechas(dato5, dato7, KeyCode)
'End Sub
'
'Private Sub dato7_KeyDown(KeyCode As Integer, Shift As Integer)
'    Call flechas(dato6, dato7, KeyCode)
''End Sub
'
'Private Sub dato2_KeyPress(KeyAscii As Integer)
'    KeyAscii = esNumero(KeyAscii, "N")
'    If KeyAscii = 13 And dato2.text = "" And dato3.text = "" And dato4.text = "" Then
'        dato2.text = Format(Now(), "dd")
'        dato3.Enabled = True
'        dato3.text = Format(Now(), "mm")
'        dato4.Enabled = True
'        dato4.text = Format(Now(), "yyyy")
'        dato5.Enabled = True
'        dato5.SetFocus
'    Else
'        If KeyAscii = 13 Then Call ceros(dato2): Call Pregunta(dato2, dato3)
'    End If
'End Sub
'
'Private Sub dato3_KeyPress(KeyAscii As Integer)
'    KeyAscii = esNumero(KeyAscii, "N")
'    If KeyAscii = 13 Then Call ceros(dato3): Call Pregunta(dato3, dato4)
'End Sub
'
'Private Sub dato4_KeyPress(KeyAscii As Integer)
'    KeyAscii = esNumero(KeyAscii, "N")
'    If KeyAscii = 13 Then Call ceros(dato4): Call Pregunta(dato4, dato5)
'End Sub
'
'Private Sub dato5_KeyPress(KeyAscii As Integer)
'    KeyAscii = esNumero(KeyAscii, "N")
'    If KeyAscii = 13 And dato5.text = "" And dato6.text = "" And dato7.text = "" Then
'        dato5.text = Format(Now(), "dd")
'        dato5.Enabled = True
'        dato6.text = Format(Now(), "mm")
'        dato6.Enabled = True
'        dato7.text = Format(Now(), "yyyy")
'        dato7.Enabled = True
'        dato7.SetFocus
'    Else
'        If KeyAscii = 13 Then Call ceros(dato5): Call Pregunta(dato5, dato6)
'    End If
'End Sub
'
'Private Sub dato6_KeyPress(KeyAscii As Integer)
'    KeyAscii = esNumero(KeyAscii, "N")
'    If KeyAscii = 13 Then Call ceros(dato6): Call Pregunta(dato6, dato7)
'End Sub
'
'Private Sub dato7_KeyPress(KeyAscii As Integer)
'    KeyAscii = esNumero(KeyAscii, "N")
'    If KeyAscii = 13 Then Call ceros(dato7): Call Pregunta(dato7, dato7)
'End Sub

Sub Pregunta(ByRef caja As TextBox, ByRef caja1 As TextBox)
    If caja.text = "" And sc = 0 Then caja.SetFocus
    If caja.text <> "" Or sc = 1 Then caja1.Enabled = True: caja1.Locked = False: caja1.SetFocus
    sc = 0
End Sub

Sub CARGASTOCKBODEGAS(codigo As String, linea As Double)
    Dim a As Integer
    Dim resultados As rdoResultset
    Dim cSql As New rdoQuery
    Dim rut As String
    Dim saldo As Double
    Dim saldo2 As Double
    Dim bodega As String
    
        Set cSql.ActiveConnection = GESTIONrubro
        cSql.sql = "SELECT local,bodega,ubicacion,stockactual "
        cSql.sql = cSql.sql + "FROM r_maestroproductos_stock_" & rubro & " "
        cSql.sql = cSql.sql + "WHERE año='" + año + "' AND codigo='" + codigo + "' order by local,bodega limit 0,10"
        cSql.Execute
         saldo = 0
         saldo2 = 0
     
        If cSql.RowsAffected > 0 Then
            Set resultados = cSql.OpenResultset
            While Not resultados.EOF
                If resultados(1) = "01" Then
                saldo = resultados(3)
                bodega = resultados(2)
               Else
               saldo2 = saldo2 + resultados(3)
               End If
               
                resultados.MoveNext
            Wend
            resultados.Close
            Set resultados = Nothing
        End If
            Grid2.Cell(linea, 6).text = saldo
            Grid2.Cell(linea, 7).text = saldo2
            Grid2.Cell(linea, 8).text = bodega
        
End Sub

Private Sub Mayor_Click()
grilla

LISTAVENTAS

End Sub

Private Sub TIENDA_Click()
grilla
LISTAVENTAS

End Sub



