VERSION 5.00
Object = "{4777436C-EB8C-4596-98A8-EBCF98683969}#1.0#0"; "FlexCell.ocx"
Object = "{ADBBAED6-B16F-40DC-80DF-B44910CBA76C}#1.0#0"; "Frame-Xp.ocx"
Begin VB.Form form1846 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Formulario 1846"
   ClientHeight    =   9705
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   15270
   DrawStyle       =   5  'Transparent
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   647
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1018
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin XPFrame.FrameXp FrameQuickMenu 
      Height          =   615
      Left            =   11880
      TabIndex        =   12
      Top             =   9000
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
         TabIndex        =   14
         Top             =   280
         Width           =   1455
      End
      Begin VB.CommandButton botonmisfavoritos 
         Caption         =   "Mis Favoritos"
         Height          =   255
         Left            =   1800
         TabIndex        =   13
         Top             =   280
         Width           =   1335
      End
   End
   Begin XPFrame.FrameXp FrameXp3 
      Height          =   9495
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   16748
      BackColor       =   16761024
      Caption         =   "Plan de Cuentas"
      CaptionEstilo3D =   1
      BackColor       =   16761024
      ForeColor       =   8438015
      BordeColor      =   14737632
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   1
      ColorTextShadow =   16744576
      Begin FlexCell.Grid Grid2 
         Height          =   9135
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   7215
         _ExtentX        =   12726
         _ExtentY        =   16113
         BackColorActiveCellSel=   16761024
         BackColorBkg    =   -2147483644
         BackColorFixed  =   -2147483639
         BackColorFixedSel=   -2147483639
         BackColorScrollBar=   16761024
         BackColorSel    =   16761024
         BorderColor     =   12632256
         CellBorderColor =   12632256
         CellBorderColorFixed=   8421504
         SelectionBorderColor=   16711680
         Cols            =   5
         DefaultFontSize =   8.25
         GridColor       =   12632256
         Rows            =   30
         SelectionMode   =   1
      End
   End
   Begin VB.TextBox pivote 
      Height          =   285
      Left            =   240
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   5760
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
   Begin XPFrame.FrameXp frmbala 
      Height          =   9495
      Left            =   7440
      TabIndex        =   3
      Top             =   0
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   16748
      BackColor       =   16761024
      Caption         =   "Determinacion Renta Liquida"
      CaptionEstilo3D =   2
      BackColor       =   16761024
      BordeColor      =   4194304
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   1
      ColorTextShadow =   16744576
      Begin VB.CommandButton Command2 
         BackColor       =   &H00FF8080&
         Caption         =   "GENERA SII"
         Height          =   375
         Left            =   2400
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   9000
         Width           =   1575
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FF8080&
         Caption         =   "IMPRIME"
         Height          =   375
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   9000
         Width           =   1695
      End
      Begin FlexCell.Grid Grid1 
         Height          =   7455
         Left            =   120
         TabIndex        =   4
         Top             =   840
         Width           =   7455
         _ExtentX        =   13150
         _ExtentY        =   13150
         BackColorActiveCellSel=   16761024
         BackColorBkg    =   -2147483644
         BackColorFixed  =   -2147483639
         BackColorFixedSel=   -2147483639
         BackColorScrollBar=   -2147483639
         BackColorSel    =   16777215
         Cols            =   5
         DefaultFontSize =   8.25
         GridColor       =   -2147483641
         Rows            =   30
         EnterKeyMoveTo  =   1
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Base Imponible 1 Categoria"
         Height          =   255
         Left            =   3000
         TabIndex        =   10
         Top             =   8520
         Width           =   2055
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Resultado del Balance"
         Height          =   255
         Left            =   3120
         TabIndex        =   9
         Top             =   480
         Width           =   2055
      End
      Begin VB.Label lblbase 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label1"
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
         Left            =   5400
         TabIndex        =   8
         Top             =   8400
         Width           =   1695
      End
      Begin VB.Label lblresultado 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label1"
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
         Left            =   5400
         TabIndex        =   7
         Top             =   360
         Width           =   1695
      End
   End
End
Attribute VB_Name = "form1846"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public saldoglobal As Double
Public ROW1 As Double
Dim totales As Double
Dim totales2(20) As Double
Dim AÑOCONSULTA As String
Dim TOTALge As Double










Private Sub Command1_Click()
imprimir

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
    objReportTitle.text = ""
    
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



Private Sub COMMAND2_Click()
Dim LINEA As String
Dim NL As Double

Dim s As Integer

Close 20
NL = 0
Open "C:\SII\f1846" + empresaactiva + ".txt" For Output As #20
For k = 1 To Grid2.Rows - 1
If Grid2.Cell(k, 4).text <> "0" Then
LINEA = "1"
LINEA = LINEA + ";" + Grid2.Cell(k, 1).text
LINEA = LINEA + ";" + Mid(Grid2.Cell(k, 2).text, 1, 3)
LINEA = LINEA + ";" + Grid2.Cell(k, 3).text
LINEA = LINEA + ";" + Grid2.Cell(k, 4).text + ";;;;;"

LINEA = Replace(LINEA, "¢", "O")
Print #20, LINEA
End If

Next k
For k = 1 To Grid1.Rows - 1
If Grid1.Cell(k, 3).text <> "0" Then
LINEA = "2;;;;"
LINEA = LINEA + ";" + Grid1.Cell(k, 1).text
LINEA = LINEA + ";" + Grid1.Cell(k, 1).text
LINEA = LINEA + ";" + Grid1.Cell(k, 2).text
LINEA = LINEA + ";" + Grid1.Cell(k, 3).text
LINEA = LINEA + ";" + Grid1.Cell(k, 4).text
LINEA = Replace(LINEA, "¢", "O")
Print #20, LINEA
End If

Next k


Close 20

Shell "notepad C:\SII\f1846" + empresaactiva + ".txt"

End Sub

Private Sub Form_Activate()
sqlconta.audit = True
sqlconta.programaactivo = Me.Caption

End Sub

Private Sub Form_Load()
Call CENTRAR(Me)

'dibu1.FileName = App.path & "\archivo.gif"
'dibu2.FileName = App.path & "\archivo.gif"
Call Conectar_BD
AÑOCONSULTA = Format(fechasistema, "YYYY")
Call CARGAPERMISO(Me.Name)
CARGAGRILLA
CARGAGRILLA2
Rem frmbala.Caption = "DETERMINACION CAPITAL PROPIO " + "01-01-" + Format(fechasistema, "YYYY")
leeplan
leecapital
End Sub



Sub CARGAGRILLA()
    Dim formatogrilla2(10, 10)
    formatogrilla2(1, 1) = "CUENTA"
    formatogrilla2(1, 2) = "NOMBRE"
    formatogrilla2(1, 3) = "MONTO"
    formatogrilla2(1, 4) = "TIPO"
    formatogrilla2(1, 5) = "HABER"
    formatogrilla2(1, 6) = "SALDO ACTUAL"
    formatogrilla2(1, 7) = "EMPRESA"
    
    
    
    Rem LARGO DE LOS DATOS
    formatogrilla2(2, 1) = "4"
    formatogrilla2(2, 2) = "40"
    formatogrilla2(2, 3) = "15"
    formatogrilla2(2, 4) = "3"
    formatogrilla2(2, 5) = "10"
    formatogrilla2(2, 6) = "10"
    formatogrilla2(2, 7) = "17"
    
    Rem TIPO DE DATOS C=CEROS,S=STRING,N=NUMERICO,D=DATE,
    formatogrilla2(3, 1) = "S"
    formatogrilla2(3, 2) = "S"
    formatogrilla2(3, 3) = "S"
    formatogrilla2(3, 4) = "N"
    formatogrilla2(3, 5) = "N"
    formatogrilla2(3, 6) = "N"
    formatogrilla2(3, 7) = "S"
    
    Rem FORMATO GRILLA
    
    formatogrilla2(4, 3) = " ###,###,###,##0"
    
    Rem LOCCKED
    formatogrilla2(5, 1) = "TRUE"
    
    formatogrilla2(5, 2) = "TRUE"
    formatogrilla2(5, 3) = "FALSE"
    formatogrilla2(5, 4) = "FALSE"
    formatogrilla2(5, 5) = "TRUE"
    formatogrilla2(5, 6) = "TRUE"
    formatogrilla2(5, 7) = "TRUE"
    
    
    Rem VALOR MAXIMO
    
    Grid1.Cols = 5
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
'    Grid1.BackColorFixedSel = RGB(110, 180, 230)
'    Grid1.BackColorBkg = RGB(90, 158, 214)
'    Grid1.BackColorScrollBar = RGB(231, 235, 247)
'    Grid1.BackColor1 = RGB(231, 235, 247)
'    Grid1.BackColor2 = RGB(239, 243, 255)
'    Grid1.GridColor = RGB(148, 190, 231)
    Grid1.Column(0).Width = 0
    
    For k = 1 To Grid1.Cols - 1
        Grid1.Cell(0, k).text = formatogrilla2(1, k)
        
        
        Grid1.Column(k).Width = Val(formatogrilla2(2, k)) * 9
        Grid1.Column(k).MaxLength = Val(formatogrilla2(2, k))
        Grid1.Column(k).FormatString = formatogrilla2(4, k)
        Grid1.Column(k).Locked = formatogrilla2(5, k)
        If formatogrilla2(3, k) = "N" Then Grid1.Column(k).Alignment = cellRightCenter: Grid1.Column(k).Mask = cellNumeric
        
        
        
        If formatogrilla2(3, k) = "S" Then Grid1.Column(k).Alignment = cellLeftTop
        
        
        If formatogrilla2(3, k) = "D" Then Grid1.Column(k).CellType = cellCalendar
        
    Next k
 
    End Sub
Sub CARGAGRILLA2()
    Dim formatogrilla2(10, 10)
    formatogrilla2(1, 1) = "CUENTA"
    formatogrilla2(1, 2) = "F_1846"
    formatogrilla2(1, 3) = "NOMBRE"
    formatogrilla2(1, 4) = "MONTO"
    formatogrilla2(1, 5) = "TIPO"
    formatogrilla2(1, 6) = "SALDO ACTUAL"
    formatogrilla2(1, 7) = "EMPRESA"
    
    
    
    Rem LARGO DE LOS DATOS
    formatogrilla2(2, 1) = "8"
    formatogrilla2(2, 2) = "5"
    formatogrilla2(2, 3) = "25"
    formatogrilla2(2, 4) = "15"
    formatogrilla2(2, 5) = "10"
    formatogrilla2(2, 6) = "10"
    formatogrilla2(2, 7) = "17"
    
    Rem TIPO DE DATOS C=CEROS,S=STRING,N=NUMERICO,D=DATE,
    formatogrilla2(3, 1) = "S"
    formatogrilla2(3, 2) = "S"
    formatogrilla2(3, 3) = "S"
    formatogrilla2(3, 4) = "N"
    formatogrilla2(3, 5) = "S"
    formatogrilla2(3, 6) = "N"
    formatogrilla2(3, 7) = "S"
    
    Rem FORMATO GRILLA
    
    formatogrilla2(4, 3) = " ###,###,##0"
    formatogrilla2(4, 4) = " ###,###,##0"
    formatogrilla2(4, 5) = " "
    formatogrilla2(4, 6) = " ###,###,##0"
    
    Rem LOCCKED
    formatogrilla2(5, 1) = "TRUE"
    
    formatogrilla2(5, 2) = "TRUE"
    formatogrilla2(5, 3) = "TRUE"
    formatogrilla2(5, 4) = "TRUE"
    formatogrilla2(5, 5) = "TRUE"
    formatogrilla2(5, 6) = "TRUE"
    formatogrilla2(5, 7) = "TRUE"
    
    
    Rem VALOR MAXIMO
    
    Grid2.Cols = 5
    Grid2.Rows = 1
    Grid2.AllowUserResizing = False
    Grid2.DisplayFocusRect = False
    Grid2.ExtendLastCol = True
    Grid2.BoldFixedCell = False
    Grid2.DrawMode = cellOwnerDraw
    Grid2.Appearance = Flat
    Grid2.ScrollBarStyle = Flat
    Grid2.FixedRowColStyle = Flat
    Grid2.BackColorFixed = RGB(90, 158, 214)
'    grid2.BackColorFixedSel = RGB(110, 180, 230)
'    grid2.BackColorBkg = RGB(90, 158, 214)
'    grid2.BackColorScrollBar = RGB(231, 235, 247)
'    grid2.BackColor1 = RGB(231, 235, 247)
'    grid2.BackColor2 = RGB(239, 243, 255)
'    grid2.GridColor = RGB(148, 190, 231)
    Grid2.Column(0).Width = 0
    
    For k = 1 To Grid2.Cols - 1
        Grid2.Cell(0, k).text = formatogrilla2(1, k)
        
        
        Grid2.Column(k).Width = Val(formatogrilla2(2, k)) * 9
        Grid2.Column(k).MaxLength = Val(formatogrilla2(2, k))
        Grid2.Column(k).FormatString = formatogrilla2(4, k)
        Grid2.Column(k).Locked = formatogrilla2(5, k)
        If formatogrilla2(3, k) = "N" Then Grid2.Column(k).Alignment = cellRightCenter
        If formatogrilla2(3, k) = "S" Then Grid2.Column(k).Alignment = cellLeftTop
        
        
        If formatogrilla2(3, k) = "D" Then Grid2.Column(k).CellType = cellCalendar
        
    Next k
 
    End Sub


Sub leecapital()

Dim resultados2 As rdoResultset
    Dim csql2 As New rdoQuery
    Dim rut As String
    Dim PASO As String
    Dim LINEAS As Double
       totales = 0
        Set csql2.ActiveConnection = conta
        csql2.sql = "SELECT an.codigo,an.nombre,ifnull(ad.monto,'0'),ifnull(ad.tipo,'') "
        csql2.sql = csql2.sql + "FROM sii_1846_anexo as an left join sii_1846_datos as ad on ad.codigo=an.codigo and ad.año='" + Format(fechasistema, "yyyy") + "' and ad.empresa='" + empresaactiva + "' "
        csql2.sql = csql2.sql + "order by codigo"
        csql2.Execute
        Grid1.AutoRedraw = False
        Grid1.Rows = 1
        If csql2.RowsAffected > 0 Then
        Set resultados2 = csql2.OpenResultset
        While Not resultados2.EOF
        Grid1.Rows = Grid1.Rows + 1
        If Right(resultados2(0), 2) = "00" Then
        Grid1.Range(Grid1.Rows - 1, 1, Grid1.Rows - 1, 2).FontBold = True
        End If
        
        Grid1.Cell(Grid1.Rows - 1, 1).text = resultados2(0)
        Grid1.Cell(Grid1.Rows - 1, 2).text = resultados2(1)
        Grid1.Cell(Grid1.Rows - 1, 3).text = resultados2(2)
        Grid1.Cell(Grid1.Rows - 1, 4).text = resultados2(3)
        
        resultados2.MoveNext
        Wend
          
          resultados2.Close
            Set resultados2 = Nothing

        End If
sumargrilla
    
    Grid1.AutoRedraw = True
        Grid1.Refresh
        
    
    

End Sub

Sub leeCAPITALDETALLE(codigo, signo)

Dim resultados2 As rdoResultset
    Dim csql2 As New rdoQuery
    Dim rut As String
    Dim PASO As String
    Dim LINEAS As Double
    Dim saldo As Double
    
        
        Set csql2.ActiveConnection = contadb
        csql2.sql = "SELECT cpd.codigo,cm.nombre "
        csql2.sql = csql2.sql + "FROM capitalpropio_detalle as cpd left join  cuentasdelmayor as cm on cpd.codigo=cm.codigo and cm.año='" + Format(fechasistema, "yyyy") + "' "
        csql2.sql = csql2.sql + " where cpd.codigotitulo='" + codigo + "' "
        csql2.sql = csql2.sql + "order by cpd.codigo"
        csql2.Execute
        LINEAS = 0
        If csql2.RowsAffected > 0 Then
        Set resultados2 = csql2.OpenResultset
        While Not resultados2.EOF
        Grid1.Rows = Grid1.Rows + 1
        Grid1.Cell(Grid1.Rows - 1, 1).text = resultados2(0)
        Grid1.Cell(Grid1.Rows - 1, 2).text = resultados2(1)
        Grid1.Cell(Grid1.Rows - 1, 3).text = signo
        saldo = leersaldomayoranterior(resultados2(0))
        If saldo < 0 Then
        saldo = saldo * -1
        End If
        Grid1.Cell(Grid1.Rows - 1, 4).text = saldo
        
        
        If signo = "+" Then
        totales = totales + CDbl(Grid1.Cell(Grid1.Rows - 1, 4).text)
        
        Else
        totales = totales - CDbl(Grid1.Cell(Grid1.Rows - 1, 4).text)
        
        End If
        
        resultados2.MoveNext
        Wend
          
          resultados2.Close
            Set resultados2 = Nothing

        End If
    
    
    
    

End Sub


Sub leeplan()

Dim resultados2 As rdoResultset
    Dim csql2 As New rdoQuery
    Dim rut As String
    Dim PASO As String
    Dim LINEAS As Double
Dim saldo As Double
Dim resul As Double

        Set csql2.ActiveConnection = contadb
        csql2.sql = "SELECT codigo,f_1846,nombre  "
        csql2.sql = csql2.sql + "FROM cuentasdelmayor where año='" + AÑOCONSULTA + "' and tipo>'2' and mid(codigo,5,4)<>'0000' "
        csql2.sql = csql2.sql + "order by codigo"
        csql2.Execute
        LINEAS = 0: resul = 0
        Grid2.AutoRedraw = False
        
        
        Grid2.Rows = 1
        If csql2.RowsAffected > 0 Then
        Set resultados2 = csql2.OpenResultset
        While Not resultados2.EOF
        Grid2.Rows = Grid2.Rows + 1
        Grid2.Cell(Grid2.Rows - 1, 1).text = resultados2(0)
        Grid2.Cell(Grid2.Rows - 1, 2).text = resultados2(1)
        Grid2.Cell(Grid2.Rows - 1, 3).text = resultados2(2)
        saldo = leersaldomayor(resultados2(0), Format(fechasistema, "yyyy-mm-dd"))
        resul = resul + saldo
        
        If saldo < 0 Then saldo = saldo * -1
        Grid2.Cell(Grid2.Rows - 1, 4).text = saldo
        
        resultados2.MoveNext
        Wend
          
          resultados2.Close
           Set resultados2 = Nothing

        End If
        Grid2.AutoRedraw = True
        Grid2.Refresh
        lblresultado.Caption = Format(resul * -1, "###,###,###,###")
        
    End Sub

Sub grabar(codigo, codigotitulo)
    campos(0, 0) = "codigo"
    campos(1, 0) = "codigotitulo"
    campos(2, 0) = ""
   
    campos(0, 1) = codigo
    campos(1, 1) = codigotitulo
  
    campos(0, 2) = "capitalpropio_detalle"
    op = 2
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)
    If Mid(codigo, 5, 4) = "0000" Then
    Call eliminasubCAPITALDETALLE(codigo)
    
    End If
    
End Sub

Sub eliminaCAPITALDETALLE(codigo)

Dim resultados2 As rdoResultset
    Dim csql2 As New rdoQuery
    Dim rut As String
    Dim PASO As String
    Dim LINEAS As Double
        Set csql2.ActiveConnection = contadb
        csql2.sql = "delete FROM capitalpropio_detalle "
        csql2.sql = csql2.sql + " where codigo='" + codigo + "' "
        csql2.Execute
        Call sincronizadatos(csql2.sql, contadb, "")
        
        
End Sub

Sub eliminasubCAPITALDETALLE(codigo)

Dim resultados2 As rdoResultset
    Dim csql2 As New rdoQuery
    Dim rut As String
    Dim PASO As String
    Dim LINEAS As Double
        Set csql2.ActiveConnection = contadb
        csql2.sql = "delete FROM capitalpropio_detalle "
        csql2.sql = csql2.sql + " where mid(codigo,1,4)='" + Mid(codigo, 1, 4) + "' and mid(codigo,5,4)<>'0000'  "
        csql2.Execute
        Call sincronizadatos(csql2.sql, contadb, "")
        
        
End Sub

Public Function existeCAPITALDETALLE(codigo) As Boolean


Dim resultados2 As rdoResultset
    Dim csql2 As New rdoQuery
    Dim rut As String
    Dim PASO As String
    Dim LINEAS As Double
        Set csql2.ActiveConnection = contadb
        csql2.sql = "select * FROM capitalpropio_detalle "
        csql2.sql = csql2.sql + " where codigo='" + codigo + "' "
        csql2.Execute
        If csql2.RowsAffected > 0 Then
        existeCAPITALDETALLE = True
        Else
        existeCAPITALDETALLE = False
        End If
        
        
        Set csql2.ActiveConnection = contadb
        csql2.sql = "select * FROM capitalpropio_detalle "
        csql2.sql = csql2.sql + " where codigo='" + Mid(codigo, 1, 4) + "0000" + "' "
        csql2.Execute
        If csql2.RowsAffected > 0 Then
        existeCAPITALDETALLE = True
        End If
        
        
        
End Function


Sub imprimir()
Dim titulo As String


titulo = "DETERMINACION CAPITAL PROPIO INICIAL AL " + Format(fechasistema, "dd-mm-yyyy")
Call CABEZAS2(titulo, "N", 1)
Grid1.DefaultFont.Size = 8
Grid1.PageSetup.Orientation = cellPortrait
Grid1.Range(0, 1, 0, Grid1.Cols - 1).Borders(cellEdgeBottom) = cellThick
Grid1.Range(0, 1, 0, Grid1.Cols - 1).Borders(cellEdgeLeft) = cellThick
Grid1.Range(0, 1, 0, Grid1.Cols - 1).Borders(cellEdgeTop) = cellThick
Grid1.Range(0, 1, 0, Grid1.Cols - 1).Borders(cellEdgeRight) = cellThick
Grid1.Range(0, 1, 0, Grid1.Cols - 1).Borders(cellInsideHorizontal) = cellThick
Grid1.Range(0, 1, 0, Grid1.Cols - 1).Borders(cellInsideVertical) = cellThick


Grid1.PageSetup.CenterHorizontally = True


Grid1.PageSetup.PrintFixedRow = True
Grid1.PageSetup.BottomMargin = 1
Grid1.PageSetup.TopMargin = 1
Grid1.PageSetup.LeftMargin = 1
Grid1.PageSetup.RightMargin = 0
Grid1.PageSetup.BlackAndWhite = True
Grid1.PageSetup.PrintGridlines = False
Grid1.PrintPreview 100

   
End Sub

Private Sub Grid1_LeaveCell(ByVal row As Long, ByVal col As Long, NewRow As Long, NewCol As Long, Cancel As Boolean)

Call modifica_1846(Grid1.Cell(row, 1).text, Format(fechasistema, "yyyy"), Grid1.Cell(row, 3).text, Grid1.Cell(row, 4).text)
sumargrilla
End Sub
Sub sumargrilla()
TOTALge = CDbl(lblresultado.Caption)

For k = 1 To Grid1.Rows - 1
If Grid1.Cell(k, 4).text = "1" Then
TOTALge = TOTALge + Val(Grid1.Cell(k, 3).text)
End If
If Grid1.Cell(k, 4).text = "2" Then
TOTALge = TOTALge - Val(Grid1.Cell(k, 3).text)
End If
If Grid1.Cell(k, 4).text = "3" Then
TOTALge = TOTALge - Val(Grid1.Cell(k, 3).text)
End If



Next k
lblbase.Caption = Format(TOTALge, "###,###,###")

End Sub
Private Sub botonmisaccesos_Click()
programafiltro = Me.Caption
misaccesos.Show
End Sub


Private Sub botonmisfavoritos_Click()
Call AgregaFavoritos(USUARIOSISTEMA, App.EXEName, Me.Name, Me.Caption)
End Sub
