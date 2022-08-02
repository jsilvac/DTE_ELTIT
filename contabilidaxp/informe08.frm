VERSION 5.00
Object = "{4777436C-EB8C-4596-98A8-EBCF98683969}#1.0#0"; "Flexcell.ocx"
Object = "{ADBBAED6-B16F-40DC-80DF-B44910CBA76C}#1.0#0"; "Frame-Xp.ocx"
Begin VB.Form informe08 
   Caption         =   "Form1"
   ClientHeight    =   8265
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11280
   LinkTopic       =   "Form1"
   ScaleHeight     =   8265
   ScaleWidth      =   11280
   StartUpPosition =   3  'Windows Default
   Begin XPFrame.FrameXp frmxp 
      Height          =   11535
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   19095
      _ExtentX        =   33681
      _ExtentY        =   20346
      BackColor       =   16744576
      Caption         =   "Informes"
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
         BackColor       =   &H00FFC0C0&
         Caption         =   "Nuevo Informe"
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
         Left            =   6480
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   7560
         Width           =   2055
      End
      Begin VB.CommandButton cmdimprimir 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Imprimir"
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
         Left            =   2280
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   7560
         Width           =   2055
      End
      Begin XPFrame.FrameXp FrameXp1 
         Height          =   6135
         Left            =   120
         TabIndex        =   7
         Top             =   1320
         Width           =   10935
         _ExtentX        =   19288
         _ExtentY        =   10821
         BackColor       =   16761024
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
         Begin FlexCell.Grid Grid1 
            Height          =   5775
            Left            =   0
            TabIndex        =   8
            Top             =   240
            Width           =   10935
            _ExtentX        =   19288
            _ExtentY        =   10186
            Cols            =   5
            DefaultFontSize =   8.25
            Rows            =   30
         End
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Generar Informe"
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
         Left            =   9480
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   360
         Width           =   1575
      End
      Begin VB.TextBox dato1 
         BackColor       =   &H00C0FFFF&
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
         Left            =   1140
         MaxLength       =   9
         TabIndex        =   1
         Tag             =   "rut"
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Dejar en Blanco Para Ver Todos"
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
         Left            =   120
         TabIndex        =   11
         Top             =   720
         Width           =   2775
      End
      Begin VB.Label dv 
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
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
         Left            =   2340
         TabIndex        =   5
         Top             =   360
         Width           =   285
      End
      Begin VB.Label Label12 
         Appearance      =   0  'Flat
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Nombre"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   2760
         TabIndex        =   4
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " R.U.T."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   975
      End
      Begin VB.Label lblnombre 
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3780
         TabIndex        =   2
         Top             =   360
         Width           =   5655
      End
   End
End
Attribute VB_Name = "informe08"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdimprimir_Click()
    If Grid1.Rows > 1 Then
        Call cabeza("LISTADO CUOTAS PRESTAMO ")
        Grid1.PrintPreview
    End If
    
End Sub

Private Sub Command1_Click()
    Call leercuotasprestamos(dato1.text & dv.Caption)
End Sub

Private Sub Command2_Click()
    dato1.text = ""
    dv.Caption = ""
    Grid1.Rows = 1
    lblnombre.Caption = ""
    
    
End Sub

Private Sub dato1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF2 Then
        Call AyudaTrabajador(dato1)
    End If
End Sub

Private Sub dato1_KeyPress(KeyAscii As Integer)
    KeyAscii = esNumero(KeyAscii)
    If KeyAscii = 13 And dato1.text <> "" Then
        Call Ceros(dato1)
        dv.Caption = rut(dato1.text)
        If leertrabajador(dato1.text & dv.Caption) = True Then
            
        Else
            MsgBox "RUT NO INGRESADO EN NUESTRA BASE DE DATOS ", vbExclamation, "ATENCION"
        End If
    End If
End Sub

Function leertrabajador(rut) As Boolean
    Dim csql As New rdoQuery
    Dim resultados As rdoResultset
    Set csql.ActiveConnection = DB
    csql.sql = "select nombre  from mt_fijo where rut='" & rut & "' and mes='" + Format(fechasistema, "mm") + "' and año='" + Format(fechasistema, "yyyy") + "' "
    csql.Execute
    leertrabajador = False
    If csql.RowsAffected > 0 Then
        Set resultados = csql.OpenResultset
        lblnombre.Caption = resultados(0)
        leertrabajador = True
        Call leercuotasprestamos(rut)
    End If
    csql.Close
    Set csql = Nothing
    Set resultados = Nothing
End Function
Sub leercuotasprestamos(rut)
    Dim csql As New rdoQuery
    Dim resultados As rdoResultset
    Dim filtro As String
    Dim FILTRO2 As String
    Dim TOTAL1 As Double
    
    Set csql.ActiveConnection = DB
    
    
    csql.sql = "select comprobante,numerocuota,mes,año,monto,mesrebajado,añorebajado,rut "
    csql.sql = csql.sql & "from prestamo_cuota where rut like '%" & rut & "%' and mesrebajado='' order by comprobante,numerocuota "
    csql.Execute
    
    If csql.RowsAffected > 0 Then
        Grid1.Rows = 1
        Set resultados = csql.OpenResultset
        filtro = resultados(7)
        FILTRO2 = filtro
        While Not resultados.EOF
            If filtro <> FILTRO2 Then
                Grid1.Rows = Grid1.Rows + 1
                
                Grid1.Column(1).Locked = False
                Grid1.Column(2).Locked = False
                Grid1.Column(3).Locked = False
                Grid1.Column(4).Locked = False
                Grid1.Column(5).Locked = False
           
                Grid1.Range(Grid1.Rows - 1, 1, Grid1.Rows - 1, 4).Merge
                Grid1.Range(Grid1.Rows - 1, 1, Grid1.Rows - 1, Grid1.cols - 1).Borders(cellEdgeLeft) = cellThin
                Grid1.Range(Grid1.Rows - 1, 1, Grid1.Rows - 1, Grid1.cols - 1).Borders(cellEdgeRight) = cellThin
                Grid1.Range(Grid1.Rows - 1, 1, Grid1.Rows - 1, Grid1.cols - 1).Borders(cellEdgeTop) = cellThin
                Grid1.Range(Grid1.Rows - 1, 1, Grid1.Rows - 1, Grid1.cols - 1).Borders(cellEdgeBottom) = cellThin
                Grid1.Range(Grid1.Rows - 1, 1, Grid1.Rows - 1, Grid1.cols - 1).FontSize = 8
                Grid1.Range(Grid1.Rows - 1, 1, Grid1.Rows - 1, Grid1.cols - 1).FontBold = True
            
                Grid1.Cell(Grid1.Rows - 1, 1).text = filtro & "   " & leernombretrabajador(filtro)
                Grid1.Cell(Grid1.Rows - 1, 5).text = Format(TOTAL1, "$ ###,###,###")
              
                TOTAL1 = 0
                filtro = resultados(7)
                Grid1.Column(1).Locked = True
                Grid1.Column(2).Locked = True
                Grid1.Column(3).Locked = True
                Grid1.Column(4).Locked = True
                Grid1.Column(5).Locked = True
            End If
        
        
            Grid1.Rows = Grid1.Rows + 1
            Grid1.Cell(Grid1.Rows - 1, 1).text = resultados(0)
            Grid1.Cell(Grid1.Rows - 1, 2).text = resultados(1) & "   "
            Grid1.Cell(Grid1.Rows - 1, 3).text = resultados(2) & "-" & resultados(3)
            Grid1.Cell(Grid1.Rows - 1, 4).text = Format(resultados(4), "$ ###,###,###")
            Grid1.Cell(Grid1.Rows - 1, 5).text = resultados(2) & "-" & resultados(3)
            TOTAL1 = TOTAL1 + resultados(4)
            
            resultados.MoveNext
             If Not resultados.EOF Then
                FILTRO2 = resultados(7)
             End If
        Wend
                Grid1.Rows = Grid1.Rows + 1
                Grid1.Column(1).Locked = False
                Grid1.Column(2).Locked = False
                Grid1.Column(3).Locked = False
                Grid1.Column(4).Locked = False
                Grid1.Column(5).Locked = False
           
                Grid1.Range(Grid1.Rows - 1, 1, Grid1.Rows - 1, 4).Merge
                Grid1.Range(Grid1.Rows - 1, 1, Grid1.Rows - 1, Grid1.cols - 1).Borders(cellEdgeLeft) = cellThin
                Grid1.Range(Grid1.Rows - 1, 1, Grid1.Rows - 1, Grid1.cols - 1).Borders(cellEdgeRight) = cellThin
                Grid1.Range(Grid1.Rows - 1, 1, Grid1.Rows - 1, Grid1.cols - 1).Borders(cellEdgeTop) = cellThin
                Grid1.Range(Grid1.Rows - 1, 1, Grid1.Rows - 1, Grid1.cols - 1).Borders(cellEdgeBottom) = cellThin
                Grid1.Range(Grid1.Rows - 1, 1, Grid1.Rows - 1, Grid1.cols - 1).FontSize = 8
                Grid1.Range(Grid1.Rows - 1, 1, Grid1.Rows - 1, Grid1.cols - 1).FontBold = True
            
                Grid1.Cell(Grid1.Rows - 1, 1).text = filtro & "   " & leernombretrabajador(filtro)
                Grid1.Cell(Grid1.Rows - 1, 5).text = Format(TOTAL1, "$ ###,###,###")
              
                TOTAL1 = 0
                Grid1.Column(1).Locked = True
                Grid1.Column(2).Locked = True
                Grid1.Column(3).Locked = True
                Grid1.Column(4).Locked = True
                Grid1.Column(5).Locked = True
    End If
    
End Sub
Sub AyudaTrabajador(caja As TextBox)
    Dim campos As Variant
    Dim cfijo As Variant
    Dim largo As Variant
    campos = Array("rut", "nombre")
    largo = Array("10s", "25s")
    cfijo = " rut <> '00'"
    cfijo = "mes='" + Format(fechasistema, "mm") + "' and año='" + Format(fechasistema, "yyyy") + "' "
    cabezas = Array("RUT", "APELLIDOS", "NOMBRES")
    mensajeAyuda = "AYUDA TRABAJADOR"
    Call Mayuda.cargaAyudaT(servidor, clientesistema + "remu" & EMPRESAACTIVA, usuario, PASSWORD, "mt_fijo", caja, campos, cfijo, largo, 2)
    ruttemporal = caja.text
    If (ruttemporal <> "") Then
        dato1.text = Left(ruttemporal, 9) 'rut
        dv.Caption = Right(ruttemporal, 1)  'dv
    End If
End Sub
Sub CARGAGRILLA(Grid1 As Grid)
    Rem DATOS DE LA COLUMNA
    Grid1.DefaultFont.Size = 8
      
    formatogrilla(1, 1) = "NUMERO"
    formatogrilla(1, 2) = "CUOTA"
    formatogrilla(1, 3) = "FECHA INICIO"
    formatogrilla(1, 4) = "MONTO"
    formatogrilla(1, 5) = "FECHA DE COBRO "
    
    Rem LARGO DE LOS DATOS
    formatogrilla(2, 1) = "10"
    formatogrilla(2, 2) = "8"
    formatogrilla(2, 3) = "0"
    formatogrilla(2, 4) = "10"
    formatogrilla(2, 5) = "12"
 
    
    Rem TIPO DE DATOS C=CEROS,S=STRING,N=NUMERICO,D=DATE,
    formatogrilla(3, 1) = ""
    formatogrilla(3, 2) = "N"
    formatogrilla(3, 3) = ""
    formatogrilla(3, 4) = ""
    formatogrilla(3, 5) = ""
 
    Rem FORMATO GRILLA
    formatogrilla(4, 1) = ""
    formatogrilla(4, 2) = ""
    formatogrilla(4, 3) = ""
    formatogrilla(4, 4) = ""
 
    Rem LOCCKED
    formatogrilla(5, 1) = "TRUE"
    formatogrilla(5, 2) = "TRUE"
    formatogrilla(5, 3) = "TRUE"
    formatogrilla(5, 4) = "TRUE"
    formatogrilla(5, 5) = "TRUE"
    Grid1.cols = 6
    Grid1.Rows = 1
    
    Grid1.AllowUserResizing = False
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
    For k = 1 To Grid1.cols - 1
        Grid1.Cell(0, k).text = formatogrilla(1, k)
        Grid1.Column(k).width = CDbl(formatogrilla(2, k)) * Grid1.DefaultFont.Size
        Grid1.Column(k).MaxLength = CDbl(formatogrilla(2, k))
        Grid1.Column(k).FormatString = formatogrilla(4, k)
        Grid1.Column(k).Locked = formatogrilla(5, k)
        If formatogrilla(3, k) = "N" Then Grid1.Column(k).Alignment = cellRightCenter
        If formatogrilla(3, k) = "" Then Grid1.Column(k).Alignment = cellCenterCenter
       
    Next k
    Grid1.Column(0).width = 0
    Grid1.Range(0, 0, 0, Grid1.cols - 1).Alignment = cellCenterCenter
    Rem Grid1.Enabled = False
End Sub

Private Sub Form_Activate()
dato1.SetFocus
End Sub

Private Sub Form_Load()
    Call CARGAGRILLA(Grid1)
End Sub
Sub cabeza(titulos)
Dim k As Integer
    Dim objReportTitle As FlexCell.ReportTitle
    
    Grid1.ReportTitles.Clear
    'Report Title 1
 
        Set objReportTitle = New FlexCell.ReportTitle
        objReportTitle.text = nombreempresa
        objReportTitle.Font.Name = "verdana"
        objReportTitle.Font.Size = 7
        objReportTitle.Font.Italic = True
        objReportTitle.PrintOnAllPages = True
        objReportTitle.Color = RGB(128, 0, 0)
        objReportTitle.Align = CellLeft
        Grid1.ReportTitles.Add objReportTitle
 
 
    Set objReportTitle = New FlexCell.ReportTitle
    objReportTitle.text = titulos
    objReportTitle.Font.Name = "verdana"
    objReportTitle.Font.Size = 12
    objReportTitle.Font.Bold = True
    objReportTitle.PrintOnAllPages = True
    Grid1.ReportTitles.Add objReportTitle
    
    Grid1.PageSetup.Footer = "Pág &P de &N" & vbCrLf & "Fecha: &D"
    Grid1.PageSetup.FooterAlignment = cellRight
    Grid1.PageSetup.FooterFont.Name = "Verdana"
    Grid1.PageSetup.FooterFont.Size = 7
    
    With Grid1.PageSetup
        .HeaderFont.Size = 6
        '.Header = "                                                                                                                   PAGINAS &P/&N EMITIDO:&D USUARIO " + USUARIOSISTEMA
        .HeaderAlignment = cellCenter
        .HeaderFont.Name = "Verdana"
        .HeaderMargin = 4
    End With
     
    Grid1.PageSetup.BlackAndWhite = True
    Grid1.PageSetup.BottomMargin = 1
    Grid1.PageSetup.LeftMargin = 1
    Grid1.PageSetup.RightMargin = 1
    Grid1.PageSetup.TopMargin = 1
    Grid1.PageSetup.PrintFixedRow = True
    Grid1.Column(1).width = 13 * 8
    
 End Sub
Function leernombretrabajador(rut) As String
    Dim csql As New rdoQuery
    Dim resultados As rdoResultset
    Set csql.ActiveConnection = DB
    csql.sql = "select nombre from mt_fijo where rut='" & rut & "' "
    csql.Execute
    leernombretrabajador = ""
    
    If csql.RowsAffected > 0 Then
        Set resultados = csql.OpenResultset
        leernombretrabajador = resultados(0)
    End If
    
    
    
End Function



