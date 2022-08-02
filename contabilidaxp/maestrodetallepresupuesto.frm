VERSION 5.00
Object = "{4777436C-EB8C-4596-98A8-EBCF98683969}#1.0#0"; "FlexCell.ocx"
Object = "{ADBBAED6-B16F-40DC-80DF-B44910CBA76C}#1.0#0"; "Frame-Xp.ocx"
Begin VB.Form presu01 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Lista Determinacion Capital Propio"
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
      Left            =   11850
      TabIndex        =   13
      Top             =   8850
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
      Begin VB.CommandButton botonmisfavoritos 
         Caption         =   "Mis Favoritos"
         Height          =   255
         Left            =   1800
         TabIndex        =   15
         Top             =   280
         Width           =   1335
      End
      Begin VB.CommandButton botonmisaccesos 
         Caption         =   "Permisos Modulo"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   280
         Width           =   1455
      End
   End
   Begin XPFrame.FrameXp FrameXp3 
      Height          =   9495
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   16748
      BackColor       =   16744576
      Caption         =   "Plan de Cuentas"
      CaptionEstilo3D =   1
      BackColor       =   16744576
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
         Width           =   5775
         _ExtentX        =   10186
         _ExtentY        =   16113
         BackColorBkg    =   -2147483644
         BackColorFixed  =   -2147483639
         BackColorFixedSel=   -2147483639
         BackColorScrollBar=   -2147483639
         BackColorSel    =   16777215
         Cols            =   5
         DefaultFontSize =   8.25
         GridColor       =   -2147483641
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
      Left            =   6120
      TabIndex        =   3
      Top             =   0
      Width           =   9015
      _ExtentX        =   15901
      _ExtentY        =   16748
      BackColor       =   16744576
      Caption         =   "Detalle de Gastos"
      CaptionEstilo3D =   1
      BackColor       =   16744576
      BordeColor      =   14737632
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
      Begin FlexCell.Grid Grid3 
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   9600
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   450
         Cols            =   5
         DefaultFontSize =   8.25
         Rows            =   30
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00FFC0C0&
         Caption         =   "IMPRIME TODO EL PLAN"
         Height          =   375
         Left            =   2760
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   8760
         Width           =   2055
      End
      Begin VB.TextBox dato2 
         Height          =   375
         Left            =   1080
         MaxLength       =   50
         TabIndex        =   10
         Top             =   8040
         Width           =   7695
      End
      Begin VB.TextBox dato1 
         Height          =   375
         Left            =   120
         MaxLength       =   5
         TabIndex        =   9
         Top             =   8040
         Width           =   855
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FFC0C0&
         Caption         =   "IMPRIME ESTE ITEM"
         Height          =   375
         Left            =   360
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   8760
         Width           =   2055
      End
      Begin FlexCell.Grid Grid1 
         Height          =   7095
         Left            =   120
         TabIndex        =   4
         Top             =   840
         Width           =   8775
         _ExtentX        =   15478
         _ExtentY        =   12515
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
         SelectionMode   =   1
      End
      Begin VB.Label LBLNIVEL 
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1680
         TabIndex        =   7
         Top             =   360
         Width           =   7215
      End
      Begin VB.Label nivel 
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   360
         Width           =   1455
      End
   End
End
Attribute VB_Name = "presu01"
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










 

Private Sub Command1_Click()
imprimir

End Sub

Sub CABEZAS2(titulo, tipo, FOLIO)
Dim objReportTitle As FlexCell.ReportTitle
Grid3.ReportTitles.Clear


Set objReportTitle = New FlexCell.ReportTitle
    objReportTitle.text = titulo
    objReportTitle.Font.Name = "arial"
    objReportTitle.Font.Size = 12
    objReportTitle.PrintOnAllPages = True
    Grid3.ReportTitles.Add objReportTitle

Set objReportTitle = New FlexCell.ReportTitle
    objReportTitle.text = ""
    
    objReportTitle.Font.Name = "arial"
    objReportTitle.Font.Size = 12
    objReportTitle.PrintOnAllPages = True
    Grid3.ReportTitles.Add objReportTitle
    
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
        Grid3.ReportTitles.Add objReportTitle
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
        Grid3.ReportTitles.Add objReportTitle
        
        Next k
    Set objReportTitle = New FlexCell.ReportTitle
        
        
        
        
        
        objReportTitle.text = ""
        objReportTitle.Font.Name = "arial"
        objReportTitle.Font.Size = 8
        objReportTitle.Font.Italic = True
        objReportTitle.PrintOnAllPages = True
        'objReportTitle.Color = RGB(128, 0, 0)
        objReportTitle.Align = CellLeft
        Grid3.ReportTitles.Add objReportTitle
        
    End If
    
With Grid3.PageSetup
        
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
CARGAGRILLA3
leeplan2
imprimir

End Sub

Private Sub dato2_KeyPress(KeyAscii As Integer)
 KeyAscii = Asc(UCase(Chr(KeyAscii)))
   
If KeyAscii = 13 Then
Call grabar(NIVEL.Caption, dato1.text, dato2.text)
Call leedetalle(NIVEL.Caption)
dato2.text = ""
End If

End Sub

Private Sub Form_Activate()
sqlconta.audit = True
sqlconta.programaactivo = Me.Caption

End Sub
Sub grabar(cuenta, codigo, NOMBRE)
    campos(0, 0) = "cuenta"
    campos(1, 0) = "codigo"
    campos(2, 0) = "nombre"
    campos(3, 0) = ""
    campos(0, 1) = cuenta
    campos(1, 1) = codigo
    campos(2, 1) = NOMBRE
  
    campos(0, 2) = "presupuesto_detalle"
    If MODIFI = 1 Then condicion = "cuenta='" + cuenta + "' and codigo='" & codigo & "'"
    If MODIFI = 1 Then op = 3 Else op = 2
    sqlconta.response = campos
    Set sqlconta.conexion = conta
    Call sqlconta.sqlconta(op, condicion)
    
    End Sub
 
Sub ELIMINAR(cuenta, codigo)
    campos(0, 2) = "presupuesto_detalle"
    condicion = "cuenta='" + cuenta + "' and codigo='" & codigo & "'"
    op = 4
    sqlconta.response = campos
    Set sqlconta.conexion = conta
    Call sqlconta.sqlconta(op, condicion)

    
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
Rem leecapital
End Sub



Sub CARGAGRILLA()
    Dim formatogrilla2(10, 10)
    formatogrilla2(1, 1) = "CUENTA"
    formatogrilla2(1, 2) = "NOMBRE"
    
    
    
    Rem LARGO DE LOS DATOS
    formatogrilla2(2, 1) = "8"
    formatogrilla2(2, 2) = "30"
    formatogrilla2(2, 3) = "8"
    formatogrilla2(2, 4) = "12"
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
    
    formatogrilla2(4, 4) = " ###,###,###,##0"
    
    Rem LOCCKED
    formatogrilla2(5, 1) = "TRUE"
    
    formatogrilla2(5, 2) = "TRUE"
    formatogrilla2(5, 3) = "TRUE"
    formatogrilla2(5, 4) = "TRUE"
    formatogrilla2(5, 5) = "TRUE"
    formatogrilla2(5, 6) = "TRUE"
    formatogrilla2(5, 7) = "TRUE"
    
    
    Rem VALOR MAXIMO
    
    Grid1.Cols = 3
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
        If formatogrilla2(3, k) = "N" Then Grid1.Column(k).Alignment = cellRightCenter
        If formatogrilla2(3, k) = "S" Then Grid1.Column(k).Alignment = cellLeftTop
        
        
        If formatogrilla2(3, k) = "D" Then Grid1.Column(k).CellType = cellCalendar
        
    Next k
 
    End Sub
Sub CARGAGRILLA2()
    Dim formatogrilla2(10, 10)
    formatogrilla2(1, 1) = "CODIGO"
    formatogrilla2(1, 2) = "NOMBRE"
    
    
    
    Rem LARGO DE LOS DATOS
    formatogrilla2(2, 1) = "8"
    formatogrilla2(2, 2) = "20"
    formatogrilla2(2, 3) = "8"
    formatogrilla2(2, 4) = "10"
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
    
    formatogrilla2(4, 3) = " ###,###,##0"
    formatogrilla2(4, 4) = " ###,###,##0"
    formatogrilla2(4, 5) = " ###,###,##0"
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
    
    Grid2.Cols = 3
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



Sub leedetalle(cuenta)

Dim resultados2 As rdoResultset
    Dim csql2 As New rdoQuery
    Dim rut As String
    Dim PASO As String
    Dim LINEAS As Double
    Dim saldo As Double
    Dim ultimo As String
    
        
        Set csql2.ActiveConnection = conta
        csql2.sql = "SELECT codigo,nombre "
        csql2.sql = csql2.sql + "FROM presupuesto_detalle "
        csql2.sql = csql2.sql + " where cuenta='" + cuenta + "' "
        csql2.sql = csql2.sql + "order by codigo"
        csql2.Execute
        LINEAS = 0
        ultimo = "0001"
        Grid1.Rows = 1
        If csql2.RowsAffected > 0 Then
        Set resultados2 = csql2.OpenResultset
        While Not resultados2.EOF
        Grid1.Rows = Grid1.Rows + 1
        Grid1.Cell(Grid1.Rows - 1, 1).text = resultados2(0)
        Grid1.Cell(Grid1.Rows - 1, 2).text = resultados2(1)
        ultimo = Format(resultados2(0) + 1, "0000")
        resultados2.MoveNext
        Wend
          
          resultados2.Close
            Set resultados2 = Nothing

        End If
    
    dato1.text = ultimo
    
    

End Sub
Sub leeplan()

Dim resultados2 As rdoResultset
    Dim csql2 As New rdoQuery
    Dim rut As String
    Dim PASO As String
    Dim LINEAS As Double
    
        Set csql2.ActiveConnection = contadb
        csql2.sql = "SELECT codigo,nombre "
        csql2.sql = csql2.sql + "FROM cuentasdelmayor where año='" + AÑOCONSULTA + "' and tipo>'3' and mid(codigo,3,6)<>'000000' "
        csql2.sql = csql2.sql + "order by codigo"
        csql2.Execute
        LINEAS = 0
        Grid2.AutoRedraw = False
        
        Grid2.Rows = 1
        If csql2.RowsAffected > 0 Then
        Set resultados2 = csql2.OpenResultset
        While Not resultados2.EOF
        Grid2.Rows = Grid2.Rows + 1
        If Mid(resultados2(0), 5, 4) = "0000" Then
        Grid2.Range(Grid2.Rows - 1, 1, Grid2.Rows - 1, 2).FontBold = True
        End If
        Grid2.Cell(Grid2.Rows - 1, 1).text = resultados2(0)
        Grid2.Cell(Grid2.Rows - 1, 2).text = resultados2(1)
        resultados2.MoveNext
        Wend
          
          resultados2.Close
           Set resultados2 = Nothing

        End If
        Grid2.AutoRedraw = True
        Grid2.Refresh
        
    End Sub

Private Sub Grid1_Click()
If Grid1.Cell(Grid1.ActiveCell.row, 0).text = "+" Or Grid1.Cell(Grid1.ActiveCell.row, 0).text = "-" Then
NIVEL.Caption = Grid1.Cell(Grid1.ActiveCell.row, 1).text
LBLNIVEL.Caption = Grid1.Cell(Grid1.ActiveCell.row, 2).text
End If

End Sub

Private Sub Grid1_KeyDown(KeyCode As Integer, ByVal Shift As Integer)
If KeyCode = 46 Then
Call ELIMINAR(NIVEL.Caption, Grid1.Cell(Grid1.ActiveCell.row, 1).text)
Call leedetalle(NIVEL.Caption)
End If

End Sub

Private Sub Grid2_DblClick()
If Mid(Grid2.Cell(Grid2.ActiveCell.row, 1).text, 5, 4) > "0000" Then
maestro01.dato1.text = Mid(Grid2.Cell(Grid2.ActiveCell.row, 1).text, 1, 2)
maestro01.dato2.text = Mid(Grid2.Cell(Grid2.ActiveCell.row, 1).text, 3, 2)
maestro01.dato3.text = Mid(Grid2.Cell(Grid2.ActiveCell.row, 1).text, 5, 4)

maestro01.Show
End If

End Sub

Private Sub Grid2_LeaveCell(ByVal row As Long, ByVal col As Long, NewRow As Long, NewCol As Long, Cancel As Boolean)

If Mid(Grid2.Cell(NewRow, 1).text, 5, 4) > "0000" Then
NIVEL.Caption = Grid2.Cell(NewRow, 1).text
LBLNIVEL.Caption = Grid2.Cell(NewRow, 2).text
Call leedetalle(NIVEL.Caption)
End If
End Sub
Sub imprimir()
Dim titulo As String


titulo = "LISTADO DE DETALLE DE GASTOS AL  " + Format(fechasistema, "dd-mm-yyyy")
Call CABEZAS2(titulo, "N", 1)
Grid3.DefaultFont.Size = 8
Grid3.PageSetup.Orientation = cellPortrait
Grid3.Range(0, 1, 0, Grid3.Cols - 1).Borders(cellEdgeBottom) = cellThick
Grid3.Range(0, 1, 0, Grid3.Cols - 1).Borders(cellEdgeLeft) = cellThick
Grid3.Range(0, 1, 0, Grid3.Cols - 1).Borders(cellEdgeTop) = cellThick
Grid3.Range(0, 1, 0, Grid3.Cols - 1).Borders(cellEdgeRight) = cellThick
Grid3.Range(0, 1, 0, Grid3.Cols - 1).Borders(cellInsideHorizontal) = cellThick
Grid3.Range(0, 1, 0, Grid3.Cols - 1).Borders(cellInsideVertical) = cellThick


Grid3.PageSetup.CenterHorizontally = False


Grid3.PageSetup.PrintFixedRow = True
Grid3.PageSetup.BottomMargin = 1
Grid3.PageSetup.TopMargin = 1
Grid3.PageSetup.LeftMargin = 1
Grid3.PageSetup.RightMargin = 0
Grid3.PageSetup.BlackAndWhite = True
Grid3.PageSetup.PrintGridlines = False
Grid3.PrintPreview 100

   
End Sub

Sub CARGAGRILLA3()
    Dim formatogrilla2(10, 10)
    formatogrilla2(1, 1) = "CODIGO"
    formatogrilla2(1, 2) = "NOMBRE"
    formatogrilla2(1, 3) = "NOMBRE GASTO"
    
    
    
    Rem LARGO DE LOS DATOS
    formatogrilla2(2, 1) = "8"
    formatogrilla2(2, 2) = "30"
    formatogrilla2(2, 3) = "30"
    formatogrilla2(2, 4) = "10"
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
    
    formatogrilla2(4, 3) = " ###,###,##0"
    formatogrilla2(4, 4) = " ###,###,##0"
    formatogrilla2(4, 5) = " ###,###,##0"
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
    
    Grid3.Cols = 4
    Grid3.Rows = 1
    Grid3.AllowUserResizing = False
    Grid3.DisplayFocusRect = False
    Grid3.ExtendLastCol = True
    Grid3.BoldFixedCell = False
    Grid3.DrawMode = cellOwnerDraw
    Grid3.Appearance = Flat
    Grid3.ScrollBarStyle = Flat
    Grid3.FixedRowColStyle = Flat
    Grid3.BackColorFixed = RGB(90, 158, 214)
'    GRID3.BackColorFixedSel = RGB(110, 180, 230)
'    GRID3.BackColorBkg = RGB(90, 158, 214)
'    GRID3.BackColorScrollBar = RGB(231, 235, 247)
'    GRID3.BackColor1 = RGB(231, 235, 247)
'    GRID3.BackColor2 = RGB(239, 243, 255)
'    GRID3.GridColor = RGB(148, 190, 231)
    Grid3.Column(0).Width = 0
    
    For k = 1 To Grid3.Cols - 1
        Grid3.Cell(0, k).text = formatogrilla2(1, k)
        
        
        Grid3.Column(k).Width = Val(formatogrilla2(2, k)) * 9
        Grid3.Column(k).MaxLength = Val(formatogrilla2(2, k))
        Grid3.Column(k).FormatString = formatogrilla2(4, k)
        Grid3.Column(k).Locked = formatogrilla2(5, k)
        If formatogrilla2(3, k) = "N" Then Grid3.Column(k).Alignment = cellRightCenter
        If formatogrilla2(3, k) = "S" Then Grid3.Column(k).Alignment = cellLeftTop
        
        
        If formatogrilla2(3, k) = "D" Then Grid3.Column(k).CellType = cellCalendar
        
    Next k
 
    End Sub

Sub leeplan2()

Dim resultados2 As rdoResultset
    Dim csql2 As New rdoQuery
    Dim rut As String
    Dim PASO As String
    Dim LINEAS As Double
    
        Set csql2.ActiveConnection = contadb
        csql2.sql = "SELECT codigo,nombre "
        csql2.sql = csql2.sql + "FROM cuentasdelmayor where año='" + AÑOCONSULTA + "' and tipo>'3' and mid(codigo,3,6)<>'000000' "
        csql2.sql = csql2.sql + "order by codigo"
        csql2.Execute
        LINEAS = 0
        Grid3.AutoRedraw = False
        
        Grid3.Rows = 1
        If csql2.RowsAffected > 0 Then
        Set resultados2 = csql2.OpenResultset
        While Not resultados2.EOF
        Grid3.Rows = Grid3.Rows + 1
        If Mid(resultados2(0), 5, 4) = "0000" Then
        Grid3.Range(Grid3.Rows - 1, 1, Grid3.Rows - 1, 2).FontBold = True
        End If
        
        Grid3.Cell(Grid3.Rows - 1, 1).text = resultados2(0)
        Grid3.Cell(Grid3.Rows - 1, 2).text = resultados2(1)
        Call leedetalle2(resultados2(0))
        resultados2.MoveNext
        Wend
          
          resultados2.Close
           Set resultados2 = Nothing

        End If
        Grid3.AutoRedraw = True
        Grid3.Refresh
        
    End Sub

Sub leedetalle2(cuenta)

Dim resultados2 As rdoResultset
    Dim csql2 As New rdoQuery
    Dim rut As String
    Dim PASO As String
    Dim LINEAS As Double
    Dim saldo As Double
    Dim ultimo As String
    
        
        Set csql2.ActiveConnection = conta
        csql2.sql = "SELECT codigo,nombre "
        csql2.sql = csql2.sql + "FROM presupuesto_detalle "
        csql2.sql = csql2.sql + " where cuenta='" + cuenta + "' "
        csql2.sql = csql2.sql + "order by codigo"
        csql2.Execute
        LINEAS = 0
        
        If csql2.RowsAffected > 0 Then
        Set resultados2 = csql2.OpenResultset
        While Not resultados2.EOF
        Grid3.Rows = Grid3.Rows + 1
        Grid3.Range(Grid3.Rows - 1, 1, Grid3.Rows - 1, 3).FontSize = 7
        
        Grid3.Cell(Grid3.Rows - 1, 2).text = resultados2(0)
        Grid3.Cell(Grid3.Rows - 1, 3).text = resultados2(1)
        
        resultados2.MoveNext
        Wend
          
          resultados2.Close
            Set resultados2 = Nothing

        End If
End Sub

Private Sub botonmisaccesos_Click()
programafiltro = Me.Caption
misaccesos.Show
End Sub


Private Sub botonmisfavoritos_Click()
Call AgregaFavoritos(USUARIOSISTEMA, App.EXEName, Me.Name, Me.Caption)
End Sub
