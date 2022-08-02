VERSION 5.00
Object = "{4777436C-EB8C-4596-98A8-EBCF98683969}#1.0#0"; "Flexcell.ocx"
Object = "{ADBBAED6-B16F-40DC-80DF-B44910CBA76C}#1.0#0"; "Frame-Xp.ocx"
Begin VB.Form tmplistado9 
   Caption         =   "LISTADO IMPUESTO LETRAS"
   ClientHeight    =   9285
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14985
   LinkTopic       =   "Form1"
   ScaleHeight     =   9285
   ScaleWidth      =   14985
   StartUpPosition =   2  'CenterScreen
   Begin XPFrame.FrameXp FrameXp2 
      Height          =   1005
      Left            =   90
      TabIndex        =   1
      Top             =   120
      Width           =   14820
      _ExtentX        =   26141
      _ExtentY        =   1773
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
      Begin VB.OptionButton Option2 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Skorpios"
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
         Left            =   9450
         TabIndex        =   14
         Top             =   990
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Maat"
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
         Left            =   7155
         TabIndex        =   13
         Top             =   1035
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.CommandButton Command2 
         Caption         =   "GENERAR"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4545
         TabIndex        =   12
         Top             =   540
         Width           =   2175
      End
      Begin VB.TextBox dato5 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Height          =   315
         Left            =   2205
         Locked          =   -1  'True
         MaxLength       =   2
         TabIndex        =   9
         Tag             =   "proveedor"
         Top             =   645
         Width           =   435
      End
      Begin VB.TextBox dato4 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Height          =   315
         Left            =   1755
         Locked          =   -1  'True
         MaxLength       =   2
         TabIndex        =   8
         Tag             =   "proveedor"
         Top             =   645
         Width           =   435
      End
      Begin VB.TextBox dato6 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Height          =   315
         Left            =   2655
         Locked          =   -1  'True
         MaxLength       =   4
         TabIndex        =   7
         Tag             =   "proveedor"
         Top             =   645
         Width           =   795
      End
      Begin VB.TextBox dato3 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Height          =   315
         Left            =   2655
         Locked          =   -1  'True
         MaxLength       =   4
         TabIndex        =   6
         Tag             =   "proveedor"
         Top             =   270
         Width           =   795
      End
      Begin VB.TextBox dato2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Height          =   315
         Left            =   2205
         Locked          =   -1  'True
         MaxLength       =   2
         TabIndex        =   5
         Tag             =   "proveedor"
         Top             =   270
         Width           =   435
      End
      Begin VB.TextBox dato1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Height          =   315
         Left            =   1755
         Locked          =   -1  'True
         MaxLength       =   2
         TabIndex        =   4
         Tag             =   "proveedor"
         Top             =   270
         Width           =   435
      End
      Begin VB.Label lbl1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Desde"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   315
         TabIndex        =   11
         Top             =   270
         Width           =   1335
      End
      Begin VB.Label lbl2 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Hasta"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   315
         TabIndex        =   10
         Top             =   630
         Width           =   1335
      End
   End
   Begin XPFrame.FrameXp FrameXp1 
      Height          =   8025
      Left            =   45
      TabIndex        =   0
      Top             =   1215
      Width           =   14865
      _ExtentX        =   26220
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
         Top             =   7470
         Width           =   2760
      End
      Begin FlexCell.Grid GRID1 
         Height          =   7080
         Left            =   45
         TabIndex        =   2
         Top             =   315
         Width           =   14730
         _ExtentX        =   25982
         _ExtentY        =   12488
         Cols            =   5
         DefaultFontSize =   8.25
         Rows            =   30
      End
   End
End
Attribute VB_Name = "tmplistado9"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

Call Titulos("LISTADO IMPUESTO TIMBRES Y ESTAMPILLAS")
GRID1.PageSetup.Orientation = cellPortrait



GRID1.PageSetup.HeaderMargin = 0.5
GRID1.PageSetup.PrintFixedRow = True



GRID1.PageSetup.TopMargin = 1
GRID1.PageSetup.LeftMargin = 0.5
GRID1.PageSetup.RightMargin = 0.5
GRID1.PageSetup.BottomMargin = 1
GRID1.PageSetup.FooterMargin = 0.5
GRID1.PageSetup.BlackAndWhite = True

GRID1.Range(0, 1, 0, GRID1.Cols - 1).Borders(cellEdgeTop) = cellThin
GRID1.Range(0, 1, 0, GRID1.Cols - 1).Borders(cellEdgeBottom) = cellThin
GRID1.Range(0, 1, 0, GRID1.Cols - 1).Borders(cellEdgeLeft) = cellThin
GRID1.Range(0, 1, 0, GRID1.Cols - 1).Borders(cellEdgeRight) = cellThin
GRID1.Range(0, 1, 0, GRID1.Cols - 1).Borders(cellInsideVertical) = cellThin
GRID1.Range(0, 1, 0, GRID1.Cols - 1).Borders(cellInsideHorizontal) = cellThin




GRID1.PrintPreview
End Sub

Private Sub Command2_Click()
Call CargaGrillaGRID1(1, 9)
LEErCREDITOS

End Sub



Private Sub Form_Load()
Call CargaGrillaGRID1(1, 9)
dato1.text = "01"
dato2.text = Format(fechasistema, "MM")
dato3.text = Format(fechasistema, "YYYY")
If dato2.text = "01" Or dato2.text = "03" Or dato2.text = "05" Or dato2.text = "07" Or dato2.text = "08" Or dato2.text = "10" Or dato2.text = "12" Then
dato4.text = "31"
Else
dato4.text = "30"
End If
If dato2.text = "02" Then
dato4.text = "28"
End If

dato5.text = Format(fechasistema, "MM")
dato6.text = Format(fechasistema, "YYYY")



End Sub

 Private Sub CargaGrillaGRID1(ByVal row As Integer, ByVal col As Integer)
        Dim i As Integer
       Dim formatogrilla(20, 20)
       Rem DATOS DE LA COLUMNA
        formatogrilla(1, 1) = "FECHA"
        formatogrilla(1, 2) = "RUT"
        formatogrilla(1, 3) = "NOMBRE"
        formatogrilla(1, 4) = "CUOTA"
        formatogrilla(1, 5) = "VALOR"
        formatogrilla(1, 6) = "VENCIMIENTO"
        formatogrilla(1, 7) = "TAZA"
        formatogrilla(1, 8) = "IMPUESTO"
        formatogrilla(1, 9) = "TOTAL"
        
        
        Rem LARGO DE LOS DATOS
        formatogrilla(2, 1) = "10"
        formatogrilla(2, 2) = ""
        formatogrilla(2, 3) = ""
        formatogrilla(2, 4) = ""
        
        Rem TIPO DE DATOS
        formatogrilla(3, 1) = "S"
        formatogrilla(3, 2) = "N"
        formatogrilla(3, 3) = "S"
        formatogrilla(3, 4) = "N"
        formatogrilla(3, 5) = "N"
        formatogrilla(3, 6) = "D"
        formatogrilla(3, 7) = "N"
        formatogrilla(3, 8) = "N"
        formatogrilla(3, 9) = "N"
        
        Rem FORMATO GRILLA
        ''''''''''''''''''''''''
        formatogrilla(4, 1) = ""
        formatogrilla(4, 2) = ""

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
        formatogrilla(8, 2) = "8"
        formatogrilla(8, 3) = "20"
        formatogrilla(8, 4) = "8"
        formatogrilla(8, 5) = "8"
        formatogrilla(8, 6) = "8"
        formatogrilla(8, 7) = "8"
        formatogrilla(8, 8) = "8"
        formatogrilla(8, 9) = "8"
            
        GRID1.Cols = col
        GRID1.Rows = row
        GRID1.AllowUserResizing = False
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
    End Sub
'**
Sub LEErCREDITOS()

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
        Dim total5 As Double
        Dim total21 As Double
        Dim total22 As Double
        Dim total23 As Double
        Dim total24 As Double
        Dim total25 As Double
        Dim plazo As Double
        Dim impuesto As Double
        
        
        Dim fecha1 As String
        Dim fecha2 As String
        Dim loc As String
        Dim taza As Double
        Dim maximo As Double
        taza = 0.1
        maximo = 1.2
        
        fecha1 = dato3.text + "-" + dato2.text + "-" + dato1.text
        fecha2 = dato6.text + "-" + dato5.text + "-" + dato4.text
        
        Set cSql = New rdoQuery
        Set cSql.ActiveConnection = ventas

        cSql.sql = "SELECT cd.fechacompra,cd.rut,mc.nombre,cd.numerocuota,cd.montocuota,cd.vencimientoactual "
        cSql.sql = cSql.sql & "FROM sv_maestroclientes as mc,sv_cuotas_detalle as cd "
        cSql.sql = cSql.sql & "WHERE cd.rut=mc.rut and cd.fechacompra between '" + fecha1 + "' and '" + fecha2 + "' and cd.tipo<>'CA' "
        Rem     cSql.sql = cSql.sql & "order by cd.rut,fechacompra,cd.numero,cd.numerocuota "
        cSql.sql = cSql.sql + "Union "
        cSql.sql = cSql.sql + " select fa.fechaingreso,fa.rut,fc.nombre,'1',fa.monto,fa.fechavencimiento "
        cSql.sql = cSql.sql + "from " + clientesistema + "teso.factoring_cheques as fa ," + clientesistema + "teso.factoring_clientes as fc where fc.rut=fa.rut and fa.fechaingreso between '" + fecha1 + "' and '" + fecha2 + "'"

        
        cSql.Execute
        
        If cSql.RowsAffected > 0 Then

            Set resultado = cSql.OpenResultset
            
        GRID1.Rows = 1
        GRID1.AutoRedraw = False
        
        total1 = 0
        total2 = 0
        total3 = 0
        total4 = 0
        total5 = 0
        total21 = 0
        total22 = 0
        total23 = 0
        total24 = 0
        total25 = 0
        loc = resultado(0)
        
        While Not resultado.EOF
      
        GRID1.Rows = GRID1.Rows + 1
                
        GRID1.Cell(GRID1.Rows - 1, 1).text = Format(resultado(0), "dd-mm-yyyy")
        GRID1.Cell(GRID1.Rows - 1, 2).text = Mid(resultado(1), 1, 9) + "-" + Mid(resultado(1), 10, 1)
        GRID1.Cell(GRID1.Rows - 1, 3).text = resultado(2)
        GRID1.Cell(GRID1.Rows - 1, 4).text = Format(resultado(3), "###,###,###")
        GRID1.Cell(GRID1.Rows - 1, 5).text = Format(resultado(4), "###,###,###")
        
        GRID1.Cell(GRID1.Rows - 1, 6).text = Format(resultado(5), "dd-mm-yyyy")
        plazo = DateDiff("d", resultado(0), resultado(5))
        plazo = Int((plazo / 30) + 0.99) * taza
        If plazo > 1.2 Then plazo = 1.2
        impuesto = Round(resultado(4) * plazo / 100)
        totalimpuesto = resultado(4) + impuesto
        GRID1.Cell(GRID1.Rows - 1, 7).text = Format(plazo, "% ##0.000")
        GRID1.Cell(GRID1.Rows - 1, 8).text = Format(impuesto, "###,###,###")
        
        
        
        total1 = total1 + impuesto
            resultado.MoveNext
            Wend
        Else
       
        End If
        GRID1.Rows = GRID1.Rows + 1
        GRID1.Range(GRID1.Rows - 1, 3, GRID1.Rows - 1, 8).Borders(cellEdgeTop) = cellThick
        
        GRID1.Cell(GRID1.Rows - 1, 3).text = "TOTAL IMPUESTO"
        
        GRID1.Cell(GRID1.Rows - 1, 8).text = Format(total1, "###,###,###")
        GRID1.Rows = GRID1.Rows + 2
        
       
        
        Set resultado = Nothing
        cSql.Close
        Set cSql = Nothing
        GRID1.AutoRedraw = True
        GRID1.Refresh
    End Sub


Sub Titulos(titulo1)

    Dim i As Integer
    Dim objReportTitle As FlexCell.ReportTitle
    
    GRID1.FixedRowColStyle = Fixed3D
    GRID1.CellBorderColorFixed = vbButtonShadow
    GRID1.ShowResizeTips = False
    GRID1.ReportTitles.Clear
    GRID1.PageSetup.CenterHorizontally = True
    GRID1.PageSetup.Orientation = cellLandscape
    
      
    GRID1.PageSetup.PrintTitleRows = 0
    
    'Logo
'    Grid1.Images.Add App.path & "\Admin.gif", "Logo"
'    Set objReportTitle = New FlexCell.ReportTitle
'    objReportTitle.ImageKey = "Logo"
'    objReportTitle.Align = CellLeft
'    Grid1.ReportTitles.Add objReportTitle
    
    'ENCABEZADO DE PAGINA
    
    GRID1.PageSetup.Header = leerNombreEmpresa("07") & vbCrLf & leerDireccionEmpresa("07") & vbCrLf & leerRutEmpresa("07")
    
    GRID1.PageSetup.HeaderAlignment = cellLeft
    GRID1.PageSetup.HeaderFont.Name = "Verdana"
    GRID1.PageSetup.HeaderFont.Size = 8
    
    'TITULOS DEL REPORTE
    
    Set objReportTitle = New FlexCell.ReportTitle
    objReportTitle.text = titulo1 & "  |  " & " MES :  " & MonthName(dato2.text)
    objReportTitle.Font.Name = "Verdana" '"Times New Roman"
    objReportTitle.Font.Size = 8
    objReportTitle.Font.Bold = True
    objReportTitle.Align = cellCenter
    objReportTitle.PrintOnAllPages = True
    GRID1.ReportTitles.Add objReportTitle
    
    
    Set objReportTitle = New FlexCell.ReportTitle
    objReportTitle.text = tipoListado
    objReportTitle.Font.Name = "Verdana" '"Times New Roman"
    objReportTitle.Font.Size = 8
    objReportTitle.Font.Bold = True
    objReportTitle.Align = cellCenter
    objReportTitle.PrintOnAllPages = True
    GRID1.ReportTitles.Add objReportTitle
    
    
    'PIE DE PAGINA
    GRID1.PageSetup.Footer = "Pág &P de &N" & "Fecha: &D" & "Usuario: " & usuarioSistema
    GRID1.PageSetup.FooterAlignment = cellCenter
    
    GRID1.PageSetup.FooterFont.Name = "Verdana"
    GRID1.PageSetup.FooterFont.Size = 7

End Sub


 Private Sub dato1_GotFocus()
        Call VerificarCajas(Me, dato1)
        Call selecciona(dato1)
    End Sub

    Private Sub dato2_GotFocus()
        Call VerificarCajas(Me, dato2)
        Call selecciona(dato2)
    End Sub

    Private Sub dato3_GotFocus()
        Call VerificarCajas(Me, dato3)
        Call selecciona(dato3)
    End Sub
    
    Private Sub dato4_GotFocus()
        Call VerificarCajas(Me, dato4)
        Call selecciona(dato4)
    End Sub

    Private Sub dato5_GotFocus()
        Call VerificarCajas(Me, dato5)
        Call selecciona(dato5)
    End Sub
    
    Private Sub dato6_GotFocus()
        Call VerificarCajas(Me, dato6)
        Call selecciona(dato6)
    End Sub
    '========================================================
    'GotFocus
    '========================================================
    
    '========================================================
    'KeyDown
    '========================================================
    Private Sub dato1_KeyDown(KeyCode As Integer, Shift As Integer)
        Call Flechas(KeyCode, dato1)
    End Sub

    Private Sub dato2_KeyDown(KeyCode As Integer, Shift As Integer)
        Call Flechas(KeyCode, dato1)
    End Sub
    
    Private Sub dato3_KeyDown(KeyCode As Integer, Shift As Integer)
        Call Flechas(KeyCode, dato2)
    End Sub
    
    Private Sub dato4_KeyDown(KeyCode As Integer, Shift As Integer)
        Call Flechas(KeyCode, dato3)
    End Sub
    
    Private Sub dato5_KeyDown(KeyCode As Integer, Shift As Integer)
        Call Flechas(KeyCode, dato4)
    End Sub
    
    Private Sub dato6_KeyDown(KeyCode As Integer, Shift As Integer)
        Call Flechas(KeyCode, dato5)
    End Sub
    '========================================================
    'KeyDown
    '========================================================
    
    '========================================================
    'KeyPress
    '========================================================
    Private Sub dato1_KeyPress(KeyAscii As Integer)
        KeyAscii = esNumero(KeyAscii)
        If KeyAscii = 13 Then
            dato1.text = ceros(dato1)
            If dato1.text = "00" Then
                dato1.text = Format(fechasistema, "dd")
            End If
            SendKeys "{Tab}"
        End If
    End Sub

    Private Sub dato2_KeyPress(KeyAscii As Integer)
        KeyAscii = esNumero(KeyAscii)
        If KeyAscii = 13 Then
            dato2.text = ceros(dato2)
            If dato2.text = "00" Then
                dato2.text = Format(fechasistema, "mm")
            End If
            SendKeys "{Tab}"
        End If
    End Sub
        
    Private Sub dato3_KeyPress(KeyAscii As Integer)
        KeyAscii = esNumero(KeyAscii)
        If KeyAscii = 13 Then
            dato3.text = ceros(dato3)
            If dato3.text = "0000" Then
                dato3.text = Format(fechasistema, "yyyy")
            End If
            fecha1 = dato3.text & "-" & dato2.text & "-" & dato1.text
            fecha2 = ""
            
            SendKeys "{Tab}"
        End If
    End Sub
    
    Private Sub dato4_KeyPress(KeyAscii As Integer)
        KeyAscii = esNumero(KeyAscii)
        If KeyAscii = 13 Then
            dato4.text = ceros(dato4)
            If dato4.text = "00" Then
                dato4.text = Format(fechasistema, "dd")
            End If
            SendKeys "{Tab}"
        End If
    End Sub
    
    Private Sub dato5_KeyPress(KeyAscii As Integer)
        KeyAscii = esNumero(KeyAscii)
        If KeyAscii = 13 Then
            dato5.text = ceros(dato5)
            If dato5.text = "00" Then
                dato5.text = Format(fechasistema, "mm")
            End If
            SendKeys "{Tab}"
        End If
    End Sub
        
    Private Sub dato6_KeyPress(KeyAscii As Integer)
        KeyAscii = esNumero(KeyAscii)
        If KeyAscii = 13 Then
            dato6.text = ceros(dato6)
            If dato6.text = "0000" Then
                dato6.text = Format(fechasistema, "yyyy")
            End If
            fecha1 = dato3.text & "-" & dato2.text & "-" & dato1.text
            fecha2 = dato6.text & "-" & dato5.text & "-" & dato4.text
            SendKeys "{Tab}"
            
        End If
    End Sub
    '========================================================
    'KeyPress
    '========================================================

    '========================================================
    'KeyUp
    '========================================================
    Private Sub dato1_KeyUp(KeyCode As Integer, Shift As Integer)
        'If Len(dato1.text) = dato1.MaxLength Then
        '    Call dato1_KeyPress(13)
        'End If
    End Sub
    
    Private Sub dato2_KeyUp(KeyCode As Integer, Shift As Integer)
        'If Len(dato2.text) = dato2.MaxLength Then
        '    Call dato2_KeyPress(13)
        'End If
    End Sub
    
    Private Sub dato3_KeyUp(KeyCode As Integer, Shift As Integer)
        'If Len(dato3.text) = dato3.MaxLength Then
        '    Call dato3_KeyPress(13)
        'End If
    End Sub
    Private Sub dato1_LostFocus()
    Call esfecha(dato1, dato2, dato3, "dd")
    End Sub
    Private Sub dato2_LostFocus()
    Call esfecha(dato1, dato2, dato3, "mm")
    End Sub
    Private Sub dato3_LostFocus()
    Call esfecha(dato1, dato2, dato3, "yyyy")
    End Sub
    Private Sub dato4_LostFocus()
    Call esfecha(dato4, dato5, dato6, "dd")
    End Sub
    Private Sub dato5_LostFocus()
    Call esfecha(dato4, dato5, dato6, "mm")
    End Sub
    Private Sub dato6_LostFocus()
    Call esfecha(dato4, dato5, dato6, "yyyy")
    End Sub
   
Private Sub Option1_Click()
LEErCREDITOS

End Sub

Private Sub Option2_Click()
LEErCREDITOS

End Sub

Private Sub Option3_Click()
LEErCREDITOS

End Sub

Private Sub Option4_Click()
LEErCREDITOS


End Sub

Private Sub Option5_Click()
LEErCREDITOS


End Sub
