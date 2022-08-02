VERSION 5.00
Object = "{4777436C-EB8C-4596-98A8-EBCF98683969}#1.0#0"; "Flexcell.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{ADBBAED6-B16F-40DC-80DF-B44910CBA76C}#1.0#0"; "Frame-Xp.ocx"
Begin VB.Form infoilas 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Listado Resumen Ilas"
   ClientHeight    =   8595
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   15270
   DrawStyle       =   5  'Transparent
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   Moveable        =   0   'False
   ScaleHeight     =   573
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1018
   Begin XPFrame.FrameXp FrameQuickMenu 
      Height          =   615
      Left            =   11715
      TabIndex        =   17
      Top             =   45
      Width           =   3255
      _ExtentX        =   5741
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
      Alignment       =   1
      Begin VB.CommandButton botonmisaccesos 
         Caption         =   "Permisos Modulo"
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   280
         Width           =   1455
      End
      Begin VB.CommandButton botonmisfavoritos 
         Caption         =   "Mis Favoritos"
         Height          =   255
         Left            =   1800
         TabIndex        =   18
         Top             =   280
         Width           =   1335
      End
   End
   Begin VB.TextBox pivote 
      Height          =   285
      Left            =   6750
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   8865
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox MANUAL 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   135
      Left            =   120
      ScaleHeight     =   135
      ScaleWidth      =   135
      TabIndex        =   1
      Top             =   6120
      Width           =   135
   End
   Begin XPFrame.FrameXp FrameXp1 
      Height          =   8610
      Left            =   120
      TabIndex        =   2
      Top             =   45
      Width           =   14865
      _ExtentX        =   26220
      _ExtentY        =   15187
      BackColor       =   16744576
      Caption         =   "INFORME VENTAS CON ILA"
      CaptionEstilo3D =   1
      BackColor       =   16744576
      ForeColor       =   8438015
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   1
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FFC0C0&
         Caption         =   "IMPRIMIR"
         Height          =   330
         Left            =   2610
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   8145
         Width           =   2130
      End
      Begin XPFrame.FrameXp FrameXp3 
         Height          =   1050
         Left            =   135
         TabIndex        =   5
         Top             =   360
         Width           =   14640
         _ExtentX        =   25823
         _ExtentY        =   1852
         BackColor       =   16744576
         Caption         =   "DATOS DE FILTRADO"
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
         Begin VB.OptionButton Option2 
            BackColor       =   &H00FF8080&
            Caption         =   "Detallado"
            Height          =   330
            Left            =   12915
            TabIndex        =   15
            Top             =   270
            Width           =   1500
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00FF8080&
            Caption         =   "Acumulado"
            Height          =   285
            Left            =   11340
            TabIndex        =   14
            Top             =   270
            Width           =   1680
         End
         Begin VB.CommandButton Command2 
            Caption         =   "LISTAR"
            Height          =   285
            Left            =   12015
            TabIndex        =   7
            Top             =   720
            Width           =   1455
         End
         Begin XPFrame.FrameXp FrameXp6 
            Height          =   675
            Left            =   90
            TabIndex        =   8
            Top             =   270
            Width           =   3300
            _ExtentX        =   5821
            _ExtentY        =   1191
            BackColor       =   16744576
            Caption         =   "MES"
            CaptionEstilo3D =   1
            BackColor       =   16744576
            ForeColor       =   65535
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Begin VB.ComboBox COMBOMES 
               Height          =   315
               Left            =   45
               TabIndex        =   9
               Top             =   270
               Width           =   3180
            End
         End
         Begin XPFrame.FrameXp FrameXp7 
            Height          =   675
            Left            =   3510
            TabIndex        =   10
            Top             =   270
            Width           =   3075
            _ExtentX        =   5424
            _ExtentY        =   1191
            BackColor       =   16744576
            Caption         =   "AÑO"
            CaptionEstilo3D =   1
            BackColor       =   16744576
            ForeColor       =   65535
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Begin VB.ComboBox COMBOAÑO 
               Height          =   315
               Left            =   90
               TabIndex        =   11
               Top             =   270
               Width           =   2865
            End
         End
         Begin XPFrame.FrameXp FrameXp4 
            Height          =   675
            Left            =   6705
            TabIndex        =   12
            Top             =   270
            Width           =   4560
            _ExtentX        =   8043
            _ExtentY        =   1191
            BackColor       =   16744576
            Caption         =   "LOCAL"
            CaptionEstilo3D =   1
            BackColor       =   16744576
            ForeColor       =   65535
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Begin VB.ComboBox ComboLOCAL 
               Height          =   315
               Left            =   90
               TabIndex        =   13
               Top             =   270
               Width           =   4395
            End
         End
      End
      Begin XPFrame.FrameXp FrameXp2 
         Height          =   6675
         Left            =   135
         TabIndex        =   3
         Top             =   1485
         Width           =   14685
         _ExtentX        =   25903
         _ExtentY        =   11774
         BackColor       =   16744576
         Caption         =   "LISTADO DE FACTURAS DE VENTA EMITIDAS"
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
         Begin MSComctlLib.ProgressBar barra 
            Height          =   255
            Left            =   120
            TabIndex        =   16
            Top             =   6360
            Width           =   14415
            _ExtentX        =   25426
            _ExtentY        =   450
            _Version        =   393216
            Appearance      =   1
         End
         Begin FlexCell.Grid Grid1 
            Height          =   6000
            Left            =   45
            TabIndex        =   4
            Top             =   225
            Width           =   14595
            _ExtentX        =   25744
            _ExtentY        =   10583
            BackColorFixed  =   16744576
            Cols            =   5
            DefaultFontSize =   8.25
            GridColor       =   16711680
            Rows            =   30
         End
      End
   End
End
Attribute VB_Name = "infoilas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private localfiltro As String
Private COSTO1 As Double
Private COSTO2 As Double
Private COSTO3 As Double
Private COSTO10 As Double
Private COSTO20 As Double
Private COSTO30 As Double


Private Sub Command1_Click()
imprimir
End Sub



Private Sub Command2_Click()
localfiltro = Mid(ComboLOCAL.text, 1, 2)
año = COMBOAÑO.text
mes = COMBOMES.ListIndex + 1

leer


End Sub



Private Sub Command3_Click()
Dim k As Integer


End Sub


Private Sub Command4_Click()
End Sub

Private Sub Form_Load()
CENTRAR Me


    
    Call Conectar_BD

    sc = 0
CARGAGRILLA
Call Conectarventas(servidor, clientesistema + "ventas00", usuario, password)
Call Conectargestion(servidor, clientesistema + "gestion", usuario, password)
Call Conectargestionrubro(servidor, clientesistema + "gestion00", usuario, password)

For k = 1 To 12
COMBOMES.AddItem MonthName(k)
Next k
COMBOMES.ListIndex = CDbl(Format(fechasistema, "mm") - 1)
For k = 2000 To Val(Format(fechasistema, "yyyy"))
COMBOAÑO.AddItem k
Next k
COMBOAÑO.ListIndex = k - 2001
LEErlocales
Option1.Value = True


End Sub








Sub Pregunta(ByRef caja As TextBox, ByRef caja1 As TextBox)
    If caja.text = "" And sc = 0 Then caja.SetFocus
    If caja.text <> "" Or sc = 1 Then caja1.Enabled = True: caja1.SetFocus
    sc = 0
End Sub




Sub flechas(ByRef caja As TextBox, ByRef caja1 As TextBox, ByRef codigo As Integer)
    If codigo = 38 And caja.Enabled = True Then caja.SetFocus
    If codigo = 40 And caja1.Enabled = True Then caja1.SetFocus
End Sub


Private Sub lblhistorico_Click(Index As Integer)

End Sub




Private Sub Label16_Click()
End Sub

Sub limpia()
    
    
End Sub

Sub imprimir()
Dim titulo As String
titulo = "LISTADO DE FACTURAS CON ILA " + COMBOMES.text + " " + COMBOAÑO.text
Call CABEZAS2(titulo, "N", "000000000")
Grid1.Range(0, 1, 0, Grid1.Cols - 1).Borders(cellEdgeBottom) = cellThick
Grid1.Range(0, 1, 0, Grid1.Cols - 1).Borders(cellEdgeLeft) = cellThick
Grid1.Range(0, 1, 0, Grid1.Cols - 1).Borders(cellEdgeTop) = cellThick
Grid1.Range(0, 1, 0, Grid1.Cols - 1).Borders(cellEdgeRight) = cellThick
Grid1.Range(0, 1, 0, Grid1.Cols - 1).Borders(cellInsideHorizontal) = cellThick
Grid1.Range(0, 1, 0, Grid1.Cols - 1).Borders(cellInsideVertical) = cellThick
Grid1.DefaultFont.Size = 8
Grid1.PageSetup.Orientation = cellLandscape

Grid1.PageSetup.PrintFixedRow = True
Grid1.PageSetup.BottomMargin = 2
Grid1.PageSetup.TopMargin = 1
Grid1.PageSetup.LeftMargin = 1
Grid1.PageSetup.RightMargin = 0
Grid1.PageSetup.BlackAndWhite = True
Grid1.PageSetup.PrintGridlines = False
Grid1.PrintPreview 100

   
End Sub
Sub grilla()
    
End Sub




Private Sub opciones_GotFocus()

MANUAL.SetFocus

End Sub
Sub CARGAGRILLA()
Rem DATOS DE LA COLUMNA
    Dim FormatoGrilla(10, 20)
    Grid1.DefaultFont.Size = 8
       
    FormatoGrilla(1, 1) = "TP"
    FormatoGrilla(1, 2) = "NUMERO"
    FormatoGrilla(1, 3) = "RUT"
    FormatoGrilla(1, 4) = "CLIENTE"
    FormatoGrilla(1, 5) = "FECHA"
    FormatoGrilla(1, 6) = "NETO"
    FormatoGrilla(1, 7) = "IVA"
    FormatoGrilla(1, 8) = "I.REFRE"
    FormatoGrilla(1, 9) = "I.VINO "
    FormatoGrilla(1, 10) = "I.LICOR"
    FormatoGrilla(1, 11) = "T.ILAS"
    FormatoGrilla(1, 12) = "T.COSTO"
    FormatoGrilla(1, 13) = "TOTAL  "
    FormatoGrilla(1, 14) = "CONTA"
    
    Rem LARGO DE LOS DATOS
    FormatoGrilla(2, 1) = "3"
    FormatoGrilla(2, 2) = "8"
    FormatoGrilla(2, 3) = "10"
    FormatoGrilla(2, 4) = "28"
    FormatoGrilla(2, 5) = "8"
    FormatoGrilla(2, 6) = "8"
    FormatoGrilla(2, 7) = "8"
    FormatoGrilla(2, 8) = "8"
    FormatoGrilla(2, 9) = "8"
    FormatoGrilla(2, 10) = "6"
    FormatoGrilla(2, 11) = "6"
    FormatoGrilla(2, 12) = "6"
    FormatoGrilla(2, 13) = "8"
    FormatoGrilla(2, 14) = "0"

    
    Rem TIPO DE DATOS C=CEROS,S=STRING,N=NUMERICO,D=DATE,
    FormatoGrilla(3, 1) = "S"
    FormatoGrilla(3, 2) = "S"
    FormatoGrilla(3, 3) = "S"
    FormatoGrilla(3, 4) = "S"
    FormatoGrilla(3, 5) = "S"
    FormatoGrilla(3, 6) = "N"
    FormatoGrilla(3, 7) = "N"
    FormatoGrilla(3, 8) = "N"
    FormatoGrilla(3, 9) = "N"
    FormatoGrilla(3, 10) = "N"
    FormatoGrilla(3, 11) = "N"
    FormatoGrilla(3, 12) = "N"
    FormatoGrilla(3, 13) = "N"
   
    Rem FORMATO GRILLA
    FormatoGrilla(4, 6) = "##,###,##0"
    FormatoGrilla(4, 7) = "##,###,##0"
    FormatoGrilla(4, 8) = "##,###,##0"
    FormatoGrilla(4, 9) = "##,###,##0"
    FormatoGrilla(4, 10) = "##,###,##0"
    FormatoGrilla(4, 11) = "##,###,##0"
    FormatoGrilla(4, 12) = "##,###,##0"
    FormatoGrilla(4, 13) = "##,###,##0"
    Rem LOCCKED
    For k = 1 To 13
    FormatoGrilla(5, k) = "TRUE"
    
    Next k
        
    
    Grid1.Cols = 14
    Grid1.Rows = 2
    
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
        
        Grid1.Cell(0, k).text = FormatoGrilla(1, k)
        Grid1.Column(k).Width = Val(FormatoGrilla(2, k)) * Grid1.DefaultFont.Size
        
        
        Grid1.Column(k).MaxLength = Val(FormatoGrilla(2, k))
        Grid1.Column(k).FormatString = FormatoGrilla(4, k)
        Grid1.Column(k).Locked = FormatoGrilla(5, k)
        If FormatoGrilla(3, k) = "N" Then Grid1.Column(k).Alignment = cellRightCenter
        If FormatoGrilla(3, k) = "D" Then Grid1.Column(k).CellType = cellCalendar
        
    Next k
   
    
    
End Sub



Private Sub monto_Click()
End Sub

Private Sub leer()

Dim resultados As rdoResultset
    Dim csql As New rdoQuery
    Dim rut As String
    Dim PASO As String
    Dim fecha1 As String
    Dim fecha2 As String
    Dim linea As Double
    Dim TOTAL As Double
    Dim fec As Double
    Dim fec1 As Double
    Dim fechasum As String
    Dim total2 As Double
    Dim tila1 As Double
    Dim tila2 As Double
    Dim tila3 As Double
    
    linea = 0: fec = 0: fec1 = 0
    fecha1 = año + "-" + mes + "-" + "01"
    fecha2 = año + "-" + mes + "-" + "31"
    
        Set csql.ActiveConnection = ventaslocal
        csql.sql = "SELECT dc.tipo,dc.numero,dc.rut,mc.nombre,dc.fecha,dc.neto,dc.iva,dc.impuestoilarefrescos,dc.impuestoilavinos,dc.impuestoilalicores,dc.impuestoharina,dc.impuestocarne,dc.total,dc.caja,dc.fecha "
        csql.sql = csql.sql + "FROM " + clientesistema + "ventas" + localfiltro + ".sv_documento_cabeza_" + localfiltro + " as dc," + clientesistema + "ventas.sv_maestroclientes as mc "
        csql.sql = csql.sql + "where dc.rut=mc.rut and mc.sucursal='0' and dc.tipo='FV' and fecha between '" + fecha1 + "' and '" + fecha2 + "' "
        csql.sql = csql.sql + "and (dc.impuestoilarefrescos<>'0' or dc.impuestoilavinos<>'0' or dc.impuestoilalicores<>'0') "
        csql.sql = csql.sql + "order by dc.tipo,dc.numero "
        csql.Execute
        tila1 = 0
        tila2 = 0
        tila3 = 0
        COSTO1 = 0
        COSTO2 = 0
        COSTO3 = 0
        
        TOTAL = 0
        total2 = 0
        Grid1.Rows = 1
        barra.Value = 0
        If csql.RowsAffected > 0 Then
        Set resultados = csql.OpenResultset
        fechasum = Format(fechasistema, "yyyy") + "/" + Format(fechasistema, "mm") + "/" + Format(fechasistema, "dd")
         barra.Max = csql.RowsAffected + 1
         
         While Not resultados.EOF
            If Option1.Value = False Then
             
             linea = linea + 1
             Grid1.Rows = Grid1.Rows + 1
             Grid1.Cell(linea, 1).text = resultados(0)
             Grid1.Cell(linea, 2).text = resultados(1)
             Grid1.Cell(linea, 3).text = Mid(resultados(2), 1, 9) + "-" + Mid(resultados(2), 10, 1)
             Grid1.Cell(linea, 4).text = resultados(3)
             Grid1.Cell(linea, 5).text = resultados(4)
             Grid1.Cell(linea, 6).text = resultados(5)
             Grid1.Cell(linea, 7).text = resultados(6)
             Grid1.Cell(linea, 8).text = resultados(7)
             Grid1.Cell(linea, 9).text = resultados(8)
             Grid1.Cell(linea, 10).text = resultados(9)
             ' Grid1.Cell(linea, 11).text = resultados(7) + resultados(8) + resultados(9)
             Grid1.Cell(linea, 11).text = resultados(9)
             Grid1.Cell(linea, 12).text = ""
             Grid1.Cell(linea, 13).text = resultados(12)
             
        End If
             Call LEERDETALLEFACTURAS(resultados(0), resultados(1), resultados(13), resultados(14))
             If Option1.Value = False Then
             ' Grid1.Cell(linea, 12).text = COSTO10 + COSTO20 + COSTO30
             Grid1.Cell(linea, 12).text = COSTO30
             End If
             COSTO10 = 0
             COSTO20 = 0
             COSTO30 = 0
             tila1 = tila1 + resultados(7)
             tila2 = tila2 + resultados(8)
             tila3 = tila3 + resultados(9)
             barra.Value = barra.Value + 1
             resultados.MoveNext
            
            Wend
         resultados.Close
            Set resultados = Nothing

End If

Grid1.Rows = Grid1.Rows + 4
            
            
             Grid1.Cell(Grid1.Rows - 2, 6).text = "VENTA"
             Grid1.Cell(Grid1.Rows - 2, 7).text = "COSTO"
             Grid1.Cell(Grid1.Rows - 2, 8).text = "DIFEREN."
             
             Grid1.Cell(Grid1.Rows - 1, 4).text = "TOTAL ILA REFRESCOS"
             Grid1.Cell(Grid1.Rows - 1, 6).text = tila1
             Grid1.Cell(Grid1.Rows - 1, 7).text = COSTO1
             Grid1.Cell(Grid1.Rows - 1, 8).text = tila1 - COSTO1
Grid1.Rows = Grid1.Rows + 1
             Grid1.Cell(Grid1.Rows - 1, 4).text = "TOTAL ILA VINOS"
             Grid1.Cell(Grid1.Rows - 1, 6).text = tila2
             Grid1.Cell(Grid1.Rows - 1, 7).text = COSTO2
             Grid1.Cell(Grid1.Rows - 1, 8).text = tila2 - COSTO2
Grid1.Rows = Grid1.Rows + 1
             Grid1.Cell(Grid1.Rows - 1, 4).text = "TOTAL ILA LICORES"
             Grid1.Cell(Grid1.Rows - 1, 6).text = tila3
             Grid1.Cell(Grid1.Rows - 1, 7).text = COSTO3
             Grid1.Cell(Grid1.Rows - 1, 8).text = tila3 - COSTO3
             
      
      
End Sub
Sub limpiar()


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
    objReportTitle.text = ComboLOCAL.text
    objReportTitle.Font.Name = "arial"
    objReportTitle.Font.Size = 12
    objReportTitle.PrintOnAllPages = True
    Grid1.ReportTitles.Add objReportTitle

    'Report Title 1
    If tipo = "N" Then
        For k = 1 To 4
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
        .TopMargin = 2
        .BottomMargin = 1
        
        
        
End With

End Sub

Sub LEErlocales()
    Dim resultados As rdoResultset
    Dim csql As New rdoQuery
    
        Set csql.ActiveConnection = gestion
        csql.sql = "SELECT codigo,nombre "
        csql.sql = csql.sql + "FROM g_maestroempresas WHERE codigocontable='" + empresaactiva + "' "
        csql.sql = csql.sql + "ORDER BY codigo "
        csql.Execute
        
        If csql.RowsAffected > 0 Then
            Set resultados = csql.OpenResultset
            While Not resultados.EOF
                ComboLOCAL.AddItem (resultados(0) + " " + resultados(1))
                resultados.MoveNext
            Wend
            resultados.Close
            Set resultados = Nothing
        ComboLOCAL.text = ComboLOCAL.List(0)
        End If
        localfiltro = Mid(ComboLOCAL.List(0), 1, 2)
        
End Sub



Sub LEERDETALLEFACTURAS(tipo, numero, caja, FECHA)
Dim resultados As rdoResultset
    Dim csql As New rdoQuery
    Dim PASO As String
    Dim rubro As String
    
        Set csql.ActiveConnection = ventaslocal
       rubro = leerdatoslocal(localfiltro, "rubro")
        csql.sql = "SELECT mpf.codigoimpuesto,sum((mpf.pcosto/(1+.19+(mi.porcentaje/100))*mi.porcentaje/100)*dd.cantidad) as costo,mpf.pcosto,mi.porcentaje,mpf.codigobarra "
        csql.sql = csql.sql + "FROM " + clientesistema + "ventas" + localfiltro + ".sv_documento_detalle_" + localfiltro + " as dd," + clientesistema + "gestion" + rubro + ".r_maestroproductos_fijo_" + rubro + " as mpf," + clientesistema + "gestion.g_maestroimpuestos as mi "
        csql.sql = csql.sql + "where dd.fecha='" + Format(FECHA, "yyyy-mm-dd") + "' and dd.caja='" + caja + "' and dd.tipo='" + tipo + "' AND dd.numero='" + numero + "' and dd.codigo=mpf.codigobarra and mpf.codigoimpuesto=mi.codigo and mi.codigo<>'00000' "
        csql.sql = csql.sql + "group by mi.codigo"
        csql.Execute
        If csql.RowsAffected > 0 Then
            Set resultados = csql.OpenResultset
            While Not resultados.EOF
                If resultados(0) = "00001" Then COSTO1 = COSTO1 + resultados(1): COSTO10 = COSTO10 + resultados(1)
                If resultados(0) = "00002" Then COSTO2 = COSTO2 + resultados(1): COSTO20 = COSTO20 + resultados(1)
                If resultados(0) = "00003" Then COSTO3 = COSTO3 + resultados(1): COSTO30 = COSTO30 + resultados(1)
                resultados.MoveNext
            Wend
            resultados.Close
            Set resultados = Nothing
      
        End If
      

End Sub

Private Sub botonmisaccesos_Click()
programafiltro = Me.Caption
misaccesos.Show
End Sub


Private Sub botonmisfavoritos_Click()
Call AgregaFavoritos(USUARIOSISTEMA, App.EXEName, Me.Name, Me.Caption)
End Sub
