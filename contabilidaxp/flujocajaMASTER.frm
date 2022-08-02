VERSION 5.00
Object = "{4777436C-EB8C-4596-98A8-EBCF98683969}#1.0#0"; "FlexCell.ocx"
Object = "{ADBBAED6-B16F-40DC-80DF-B44910CBA76C}#1.0#0"; "Frame-Xp.ocx"
Begin VB.Form flujocajamaster 
   BackColor       =   &H00FFC0C0&
   Caption         =   "Form1"
   ClientHeight    =   10440
   ClientLeft      =   60
   ClientTop       =   390
   ClientWidth     =   12765
   LinkTopic       =   "Form1"
   ScaleHeight     =   10440
   ScaleWidth      =   12765
   StartUpPosition =   2  'CenterScreen
   Begin XPFrame.FrameXp FrameQuickMenu 
      Height          =   615
      Left            =   9360
      TabIndex        =   12
      Top             =   9600
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   1085
      BackColor       =   16744576
      Caption         =   " Mis Datos"
      BackColor       =   16744576
      BordeColor      =   4194304
      ColorBarraArriba=   4194304
      ColorBarraAbajo =   4194304
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.CommandButton botonmisaccesos 
         Caption         =   "Permisos Modulo"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   280
         Width           =   1455
      End
      Begin VB.CommandButton botonmisfavoritos 
         Caption         =   "Mis Favoritos"
         Height          =   255
         Left            =   1680
         TabIndex        =   13
         Top             =   280
         Width           =   1335
      End
   End
   Begin XPFrame.FrameXp FrameXp2 
      Height          =   855
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   12495
      _ExtentX        =   22040
      _ExtentY        =   1508
      BackColor       =   16744576
      Caption         =   "Datos"
      CaptionEstilo3D =   1
      BackColor       =   16744576
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
      Begin VB.CheckBox Check1 
         BackColor       =   &H00FF8080&
         Caption         =   "Flujo Mensual"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   9960
         TabIndex        =   11
         Top             =   360
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   6000
         Locked          =   -1  'True
         TabIndex        =   10
         Text            =   "Datos que Carga Automatico"
         Top             =   360
         Width           =   3495
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Cargar Automaticos"
         Height          =   375
         Left            =   3720
         TabIndex        =   9
         Top             =   360
         Width           =   1575
      End
      Begin VB.CommandButton Command1 
         Caption         =   "LEER FLUJO"
         Height          =   375
         Left            =   1680
         TabIndex        =   7
         Top             =   360
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
         Left            =   120
         MaxLength       =   2
         TabIndex        =   5
         Tag             =   "fecha"
         Top             =   480
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
         Left            =   480
         MaxLength       =   2
         TabIndex        =   4
         Tag             =   "fecha"
         Top             =   480
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
         Left            =   840
         MaxLength       =   4
         TabIndex        =   3
         Tag             =   "fecha"
         Top             =   480
         Width           =   615
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
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   1335
      End
   End
   Begin XPFrame.FrameXp FrameXp1 
      Height          =   9255
      Left            =   120
      TabIndex        =   0
      Top             =   1080
      Width           =   12495
      _ExtentX        =   22040
      _ExtentY        =   16325
      BackColor       =   16744576
      Caption         =   "Grilla"
      CaptionEstilo3D =   1
      BackColor       =   16744576
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
         Caption         =   "Imprimir"
         Height          =   375
         Left            =   5400
         TabIndex        =   8
         Top             =   8640
         Width           =   1575
      End
      Begin FlexCell.Grid Grid1 
         Height          =   8295
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   12255
         _ExtentX        =   21616
         _ExtentY        =   14631
         Cols            =   5
         DefaultFontSize =   8.25
         Rows            =   30
      End
   End
End
Attribute VB_Name = "flujocajamaster"
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
Dim glosacredito As String



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
        For k = 1 To 5
        Set objReportTitle = New FlexCell.ReportTitle
        objReportTitle.text = DATOSEMPRESA(k)
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



Private Sub Command3_Click()
cargaAUTOMATICO
Command1_Click

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
Call Command1_Click

End Sub

Private Sub Form_Load()
fechaflujo = fechasistema
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
    Grid1.Rows = 1
    Grid1.Cols = 1
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
        csql.sql = csql.sql + "FROM maestroempresas "
        csql.sql = csql.sql + "ORDER BY codigoempresa "
        csql.Execute
        Grid1.Rows = 1
        For k = 1 To 31
        totales(k) = 0
        totales2(k) = 0
        SALDOANTE(k) = 0
        Next k
         Rem SALDO ANTERIOR
           
        
        
        If csql.RowsAffected > 0 Then
            Set resultados = csql.OpenResultset
            While Not resultados.EOF
                Grid1.Rows = Grid1.Rows + 1
                Grid1.Cell(Grid1.Rows - 1, 1).text = resultados(0)
                Grid1.Cell(Grid1.Rows - 1, 2).text = resultados(1)
                Call LEERFLUJOsemanal(resultados(0), fechalunes, Format(DateAdd("d", 31, fechalunes), "yyyy-mm-dd"))
                
                For k = 1 To 31
                totales(k) = 0
                Next k
                
                

                
                resultados.MoveNext
            Wend
            
' Rem
           
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
            Grid1.Cell(Grid1.Rows - 1, k + 2).text = totales2(k)
            Next k
            
            
            resultados.Close
            Set resultados = Nothing
        
        End If
        
End Sub

Sub LEERFLUJOsemanal(empresa, fecha1, fecha2)
    Dim resultados As rdoResultset
    Dim csql As New rdoQuery
    Dim dia As Double
'    SELECT fc.fecha,sum(if(fct.tipo="+",monto,monto*-1)),fct.tipo
'FROM flujo_caja as fc left join flujo_caja_titulos as fct
'on (fct.codigo=mid(fc.tipo,1,2) and fct.subcuenta=mid(fc.tipo,3,4) )
'where fecha between '2009-05-11' and '2009-05-15' and empresa='08'
'GROUP BY fecha

        Set csql.ActiveConnection = conta
        csql.sql = "SELECT fc.fecha,sum(if(fct.tipo='+',monto,monto*-1)) "
        csql.sql = csql.sql + "FROM flujo_caja as fc left join flujo_caja_titulos as fct on (fct.codigo=mid(fc.tipo,1,2) and fct.subcuenta=mid(fc.tipo,3,4)) "
        csql.sql = csql.sql + "where fecha between '" + Format(fecha1, "yyyy-mm-dd") + "' and '" + Format(fecha2, "yyyy-mm-dd") + "' and empresa='" + empresa + "' "
        csql.sql = csql.sql + "GROUP BY fecha  "
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
                
                
                resultados.MoveNext
            Wend
            resultados.Close
            Set resultados = Nothing
        
        End If
       For k = 1 To 31
       Grid1.Cell(Grid1.Rows - 1, 2 + k).text = Format(totales(k), "###,###,###,##0.0")
       
       
       If k < 31 Then
       
       totales(k + 1) = totales(k + 1) + totales(k)
       End If
       Next k
       For k = 1 To 31
       totales2(k) = totales2(k) + totales(k)
       Next k
       
       
               
       Rem  Call grabarsaldos("010000", "2009-05-12", "SALDO ANTERIOR", totales2(1))
        
End Sub
Public Sub grabarTIPO(tipo, fecha, glosa, monto, empresa)
    Dim resultados As rdoResultset
    Dim csql As New rdoQuery
    Dim dia As Double
        Set csql.ActiveConnection = conta
        csql.sql = "insert into flujo_caja set tipo='" & tipo & "',fecha='" & Format(fecha, "yyyy-mm-dd") & "',glosa='" & glosa & "',monto='" & monto & "',empresa='" + empresa + "',automatico='S' "
        csql.Execute
        Call sincronizadatos(csql.sql, conta, "")
        
 
    End Sub

Private Sub Grid1_DblClick()
If Grid1.ActiveCell.col > 1 Then
empresaflujo = Grid1.Cell(Grid1.ActiveCell.row, 1).text
flujocaja.LBLEMPRESA.Caption = Grid1.Cell(Grid1.ActiveCell.row, 2).text
flujocaja.Show vbModal
End If
End Sub

Sub cargaAUTOMATICO()
Dim CHEQUEDIAMB As Double
Dim CHEQUEFECHAMB As Double
Dim CHEQUEDIAOB As Double
Dim CHEQUEFECHAOB As Double
Dim FECHACONSULTA As String
Dim TOTALCREDITOS As Double
Dim FECHADOMINGO As String


Call ELIMINAAUTOMATICOS("010003", fechalunes, Format(DateAdd("d", 31, fechalunes), "yyyy-mm-dd"))
Call ELIMINAAUTOMATICOS("010004", fechalunes, Format(DateAdd("d", 31, fechalunes), "yyyy-mm-dd"))
Call ELIMINAAUTOMATICOS("020004", fechalunes, Format(DateAdd("d", 31, fechalunes), "yyyy-mm-dd"))
Call ELIMINAAUTOMATICOS("020002", fechalunes, Format(DateAdd("d", 31, fechalunes), "yyyy-mm-dd"))
Call ELIMINAAUTOMATICOS("010008", fechalunes, Format(DateAdd("d", 31, fechalunes), "yyyy-mm-dd"))
Call ELIMINAAUTOMATICOS("020006", fechalunes, Format(DateAdd("d", 31, fechalunes), "yyyy-mm-dd"))


    Dim resultados As rdoResultset
    Dim csql As New rdoQuery
        Set csql.ActiveConnection = conta
        csql.sql = "SELECT * "
        csql.sql = csql.sql + "FROM maestroempresas "
        csql.sql = csql.sql + "ORDER BY codigoempresa "
        csql.Execute
    If csql.RowsAffected > 0 Then
    Set resultados = csql.OpenResultset
    While resultados.EOF = False
If Format(fechalunes, "dd") > "20" And Format(fechalunes, "dd") < "30" Then
Call cargaremuneraciones(resultados(0), Format(fechalunes, "yyyy-mm") + "-25", Format(fechalunes, "yyyy-mm") + "-25", "0")
End If
FECHADOMINGO = fechalunes
'For k = 2 To 1 Step -1
'FECHADOMINGO = DateAdd("d", -k, fechalunes)
'CHEQUEDIA = LEECHEQUES(resultados(0), FECHADOMINGO, FECHADOMINGO, "037")
'CHEQUEFECHA = LEEcartera(resultados(0), FECHADOMINGO, FECHADOMINGO, "037")
'CHEQUEDIAOB = LEECHEQUES(resultados(0), FECHADOMINGO, FECHADOMINGO, "001")
'CHEQUEFECHAOB = LEEcartera(resultados(0), FECHADOMINGO, FECHADOMINGO, "001")
'If CHEQUEDIA <> 0 Then
'Call grabarTIPO("010003", fechalunes, "CHEQUE " + Format(FECHADOMINGO) + "DIA M/B", CHEQUEDIA, resultados(0))
'End If
'If CHEQUEDIA <> 0 Then
'Call grabarTIPO("010004", fechalunes, "CHEQUE " + Format(FECHADOMINGO) + "DIA O/B", CHEQUEDIA, resultados(0))
'End If
'If CHEQUEFECHA <> 0 Then
'Call grabarTIPO("010003", fechalunes, "CHEQUE " + Format(FECHADOMINGO) + "FECHA M/B", CHEQUEFECHA, resultados(0))
'End If
'If CHEQUEDIAOB <> 0 Then
'Call grabarTIPO("010004", fechalunes, "CHEQUE " + Format(FECHADOMINGO) + "DIA M/B", CHEQUEDIAOB, resultados(0))
'End If
'If CHEQUEFECHAOB <> 0 Then
'Call grabarTIPO("010003", fechalunes, "CHEQUE " + Format(FECHADOMINGO) + "FECHA O/B", CHEQUEFECHAOB, resultados(0))
'End If
'TOTALCREDITOS = LEECREDITOS(resultados(0), FECHADOMINGO, FECHADOMINGO, "02")
'If TOTALCREDITOS <> 0 Then
'Call grabarTIPO("020004", fechalunes, "CREDITO EN U.F", TOTALCREDITOS, resultados(0))
'End If
'TOTALCREDITOS = LEECREDITOS(resultados(0), FECHADOMINGO, FECHADOMINGO, "01")
'If TOTALCREDITOS <> 0 Then
'Call grabarTIPO("020004", fechalunes, "CREDITO EN PESOS", TOTALCREDITOS, resultados(0))
'End If
Rem Call cargachequesafecha(resultados(0), FECHADOMINGO, fechalunes, "1")
Rem Call cargaventaproyectadaboletas(resultados(0), FECHADOMINGO, fechalunes, "1")
Rem Next k
Rem CARGA SEMANA

FECHACONSULTA = fechalunes
For k = 1 To 31
CHEQUEDIA = LEECHEQUES(resultados(0), FECHACONSULTA, FECHACONSULTA, "037")
CHEQUEFECHA = LEEcartera(resultados(0), FECHACONSULTA, FECHACONSULTA, "037")
CHEQUEDIAOB = LEECHEQUES(resultados(0), FECHACONSULTA, FECHACONSULTA, "001")
CHEQUEFECHAOB = LEEcartera(resultados(0), FECHACONSULTA, FECHACONSULTA, "001")
If CHEQUEDIA <> 0 Then
Call grabarTIPO("010003", FECHACONSULTA, "CHEQUE DIA MISMO BANCO", CHEQUEDIA, resultados(0))
End If
If CHEQUEDIA <> 0 Then
Call grabarTIPO("010004", FECHACONSULTA, "CHEQUE DIA MISMO BANCO", CHEQUEDIA, resultados(0))
End If
If CHEQUEFECHA <> 0 Then
Call grabarTIPO("010003", FECHACONSULTA, "CHEQUE FECHA MISMO BANCO", CHEQUEFECHA, resultados(0))
End If
If CHEQUEDIAOB <> 0 Then
Call grabarTIPO("010004", FECHACONSULTA, "CHEQUE DIA OTRO BANCO", CHEQUEDIAOB, resultados(0))
End If
If CHEQUEFECHAOB <> 0 Then
Call grabarTIPO("010003", FECHACONSULTA, "CHEQUE FECHA OTRO BANCO", CHEQUEFECHAOB, resultados(0))
End If
TOTALCREDITOS = LEECREDITOS(resultados(0), FECHACONSULTA, FECHACONSULTA, "02")
If TOTALCREDITOS <> 0 Then
Call grabarTIPO("020004", FECHACONSULTA, "CREDITO EN U.F", TOTALCREDITOS, resultados(0))
End If
TOTALCREDITOS = LEECREDITOS(resultados(0), FECHACONSULTA, FECHACONSULTA, "01")
If TOTALCREDITOS <> 0 Then
Call grabarTIPO("020004", FECHACONSULTA, "CREDITO EN PESOS", TOTALCREDITOS, resultados(0))
End If
Call cargachequesafecha(resultados(0), FECHACONSULTA, FECHACONSULTA, "2")
Call cargaventaproyectadaboletas(resultados(0), FECHACONSULTA, FECHACONSULTA, "1")




FECHACONSULTA = DateAdd("d", k, fechalunes)
Next k



    resultados.MoveNext
    Wend
End If
    


End Sub

Public Function LEECHEQUES(empresa, fecha1, fecha2, banco) As Double

    Dim resultados2 As rdoResultset
    Dim csql2 As New rdoQuery
    

    
        Set csql2.ActiveConnection = contadb
        csql2.sql = "select rc.local,me.codigocontable,sum(rc.monto) from " + clientesistema + "teso.rc_cartera as rc," + clientesistema + "gestion.g_maestroempresas as me "
        csql2.sql = csql2.sql + "where me.codigo=rc.local and me.codigocontable='" + empresa + "' and rc.fecha='" + Format(fecha1, "yyyy-mm-dd") + "' and rc.cartera='N' "
        If banco = "037" Then
        csql2.sql = csql2.sql + "and rc.banco='037' "
        Else
        csql2.sql = csql2.sql + "and rc.banco<>'037' "
        End If
        
        csql2.sql = csql2.sql + "group by me.codigocontable"
        csql2.Execute
        If csql2.RowsAffected > 0 Then
        Set resultados2 = csql2.OpenResultset
        
        LEECHEQUES = resultados2(2)
        Else
        
        LEECHEQUES = 0
        End If
        
    End Function
Public Function LEEcartera(empresa, fecha1, fecha2, banco) As Double

    Dim resultados2 As rdoResultset
    Dim csql2 As New rdoQuery
    
        Set csql2.ActiveConnection = contadb
        csql2.sql = "select rc.local,me.codigocontable,sum(rc.monto) from " + clientesistema + "teso.rc_cartera as rc," + clientesistema + "gestion.g_maestroempresas as me "
        csql2.sql = csql2.sql + "where me.codigo=rc.local and me.codigocontable='" + empresa + "' and rc.vencimiento='" + Format(fecha1, "yyyy-mm-dd") + "' and rc.cartera='S'  "
        If banco = "037" Then
        csql2.sql = csql2.sql + "and rc.banco='037' "
        Else
        csql2.sql = csql2.sql + "and rc.banco<>'037' "
        End If
        csql2.sql = csql2.sql + "group by me.codigocontable"
        csql2.Execute
        
        If csql2.RowsAffected > 0 Then
        Set resultados2 = csql2.OpenResultset
        
        LEEcartera = resultados2(2)
        Else
        
        LEEcartera = 0
        End If
        
    End Function

Public Sub ELIMINAAUTOMATICOS(tipo, fecha1, fecha2)
    Dim resultados As rdoResultset
    Dim csql As New rdoQuery
    Dim dia As Double
        Set csql.ActiveConnection = conta
        csql.sql = "delete from flujo_caja where fecha between '" + Format(fecha1, "yyyy-mm-dd") + "' and '" + Format(fecha2, "yyyy-mm-dd") + "' and automatico='S' "
        csql.Execute
        Call sincronizadatos(csql.sql, conta, "")
        
 
    End Sub

Public Function LEECREDITOS(empresa, fecha1, fecha2, tipo) As Double

    Dim resultados2 As rdoResultset
    Dim csql2 As New rdoQuery
    Dim FECHA3 As String
    FECHA3 = Format(fecha1, "yyyy-mm-dd")
    dia = Weekday(Format(fecha2, "yyyy-mm-dd"), vbUseSystemDayOfWeek)

If dia = 7 Then FECHA3 = DateAdd("d", 2, FECHA3)
If dia = 8 Then FECHA3 = DateAdd("d", 1, FECHA3)

 
        Set csql2.ActiveConnection = contadb
        csql2.sql = "select sum(cv.monto) from " + clientesistema + "creditos_bancarios.creditos_vencimientos as cv left join " + clientesistema + "creditos_bancarios.maestro_compromisos as mc on (cv.tipo=mc.tipo and cv.banco=mc.banco and cv.numero=mc.numero)"
        csql2.sql = csql2.sql + "where cv.empresa='" + empresa + "' and cv.fecha='" + Format(fecha1, "yyyy-mm-dd") + "' and cv.pagado='0' and mc.moneda='" + tipo + "' "
        csql2.sql = csql2.sql + "group by cv.empresa"
        csql2.Execute
        If csql2.RowsAffected > 0 Then
        Set resultados2 = csql2.OpenResultset
        If tipo = "02" Then
        LEECREDITOS = resultados2(0) * leerUF(FECHA3)
        Else
        LEECREDITOS = resultados2(0)
        
        End If
        
        Else
        
        LEECREDITOS = 0
        End If
        
    End Function
Public Sub cargachequesafecha(empresa, fecha1, fechagrabar, tipo)

    Dim resultados2 As rdoResultset
    Dim csql2 As New rdoQuery
    
        Set csql2.ActiveConnection = contadb
        csql2.sql = "select monto,giradoa from " + clientesistema + "conta" + empresa + ".chequesdocumento "
        csql2.sql = csql2.sql + "where vencimiento='" + Format(fecha1, "yyyy-mm-dd") + "' and fechacobro='0000-00-00' "
        csql2.Execute
        If csql2.RowsAffected > 0 Then
        Set resultados2 = csql2.OpenResultset
        While resultados2.EOF = False
       
        Call grabarTIPO("020002", fechagrabar, resultados2(1), resultados2(0), empresa)
       
        resultados2.MoveNext
        Wend
     
       End If
End Sub
Public Sub cargaventaproyectadaboletas(empresa, fecha1, fechagrabar, tipo)

    Dim resultados2 As rdoResultset
    Dim csql2 As New rdoQuery
    Dim fecha2 As String
    Dim dia As String
    
    fecha2 = DateAdd("yyyy", -1, fecha1)
    Rem fecha2 = DateAdd("d", 1, fecha2)
    

        Set csql2.ActiveConnection = contadb
        csql2.sql = "select sum(total) from " + clientesistema + "conta" + empresa + ".boletasdeventa "
        csql2.sql = csql2.sql + "where fecha='" + Format(fecha2, "yyyy-mm-dd") + "' group by fecha"
        csql2.Execute
        If csql2.RowsAffected > 0 Then
        Set resultados2 = csql2.OpenResultset
        While resultados2.EOF = False
        Call grabarTIPO("010008", fechagrabar, "BOLETAS " + dia + " " + Format(fecha2, "dd-mm-yyyy"), resultados2(0) * 0.8, empresa)
        resultados2.MoveNext
        Wend
     
       End If
End Sub
Public Sub cargaventaproyectadafacturas(empresa, fecha1, fechagrabar, tipo)

    Dim resultados2 As rdoResultset
    Dim csql2 As New rdoQuery
    Dim fecha2 As String
    Dim dia As String
    
    fecha2 = DateAdd("yyyy", -1, fecha1)
    Rem fecha2 = DateAdd("d", 1, fecha2)
    
dia = Weekday(Format(fecha2, "yyyy-mm-dd"), vbUseSystemDayOfWeek)
dia = WeekdayName(dia)

        Set csql2.ActiveConnection = contadb
        csql2.sql = "select sum(total) from " + clientesistema + "conta" + empresa + ".boletasdeventa "
        csql2.sql = csql2.sql + "where fecha='" + Format(fecha2, "yyyy-mm-dd") + "' group by fecha"
        csql2.Execute
        If csql2.RowsAffected > 0 Then
        Set resultados2 = csql2.OpenResultset
        While resultados2.EOF = False
        Call grabarTIPO("010008", fechagrabar, "FACTURAS " + dia + " " + Format(fecha2, "dd-mm-yyyy"), resultados2(0) * 0.8, empresa)
        resultados2.MoveNext
        Wend
     
       End If
End Sub
Public Sub cargaremuneraciones(empresa, fecha1, fechagrabar, tipo)

    Dim resultados2 As rdoResultset
    Dim csql2 As New rdoQuery
    Dim fecha2 As String
    Dim dia As String
    Dim MES As String
    
    fecha2 = DateAdd("m", -1, fecha1)
    
        Set csql2.ActiveConnection = contadb
        Rem SELECT SUM(monto),mes fROM movimientoscontables where codigocuenta='23100028' and dh='D' GROUP BY MES
        csql2.sql = "select sum(monto),mes from " + clientesistema + "conta" + empresa + ".movimientoscontables "
        csql2.sql = csql2.sql + "where codigocuenta='23100028' and MES='" + Format(fecha2, "mm") + "' and dh='D' group by mes"
        csql2.Execute
        If csql2.RowsAffected > 0 Then
        Set resultados2 = csql2.OpenResultset
        While resultados2.EOF = False
        Call grabarTIPO("020006", fechagrabar, "SUELDOS POR PAGAR", resultados2(0), empresa)
        resultados2.MoveNext
        Wend
     
       End If
End Sub
Private Sub botonmisaccesos_Click()
    programafiltro = Me.Caption
    misaccesos.Show
End Sub


Private Sub botonmisfavoritos_Click()
    Call AgregaFavoritos(USUARIOSISTEMA, App.EXEName, Me.Name, Me.Caption)
End Sub
