VERSION 5.00
Object = "{4777436C-EB8C-4596-98A8-EBCF98683969}#1.0#0"; "FlexCell.ocx"
Object = "{ADBBAED6-B16F-40DC-80DF-B44910CBA76C}#1.0#0"; "Frame-Xp.ocx"
Begin VB.Form maestro20 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Maestro de Cuentas del Mayor"
   ClientHeight    =   9705
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   14535
   DrawStyle       =   5  'Transparent
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   Moveable        =   0   'False
   ScaleHeight     =   647
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   969
   ShowInTaskbar   =   0   'False
   Begin XPFrame.FrameXp FrameXp3 
      Height          =   9255
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   16325
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
         Height          =   8895
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   6375
         _ExtentX        =   11245
         _ExtentY        =   15690
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
   Begin XPFrame.FrameXp FrameXp1 
      Height          =   9255
      Left            =   6840
      TabIndex        =   3
      Top             =   120
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   16325
      BackColor       =   16744576
      Caption         =   "Capital Propio"
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
      Begin FlexCell.Grid Grid1 
         Height          =   8295
         Left            =   120
         TabIndex        =   4
         Top             =   840
         Width           =   6375
         _ExtentX        =   11245
         _ExtentY        =   14631
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
         BackColor       =   &H80000009&
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
         Left            =   960
         TabIndex        =   7
         Top             =   360
         Width           =   5295
      End
      Begin VB.Label NIVEL 
         BackColor       =   &H80000009&
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
         Width           =   735
      End
   End
End
Attribute VB_Name = "MAESTRO20"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public saldoglobal As Double
Public ROW1 As Double









Private Sub Form_Activate()
sqlconta.audit = True
sqlconta.programaactivo = Me.Caption

End Sub

Private Sub Form_Load()
Call CENTRAR(Me)

'dibu1.FileName = App.path & "\archivo.gif"
'dibu2.FileName = App.path & "\archivo.gif"


    
    Call Conectar_BD

    
Call CARGAPERMISO(Me.Name)
CARGAGRILLA
CARGAGRILLA2

leeplan
leecapital
End Sub



Sub CARGAGRILLA()
    Dim formatogrilla2(10, 10)
    formatogrilla2(1, 1) = "CUENTA"
    formatogrilla2(1, 2) = "NOMBRE"
    formatogrilla2(1, 3) = "SALDO"
    formatogrilla2(1, 4) = "DEBE"
    formatogrilla2(1, 5) = "HABER"
    formatogrilla2(1, 6) = "SALDO ACTUAL"
    formatogrilla2(1, 7) = "EMPRESA"
    
    
    
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
    formatogrilla2(3, 3) = "N"
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
    
    Grid1.Cols = 4
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
    formatogrilla2(1, 1) = "CUENTA"
    formatogrilla2(1, 2) = "NOMBRE"
    formatogrilla2(1, 3) = "SALDO"
    formatogrilla2(1, 4) = "DEBE"
    formatogrilla2(1, 5) = "HABER"
    formatogrilla2(1, 6) = "SALDO ACTUAL"
    formatogrilla2(1, 7) = "EMPRESA"
    
    
    
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
    formatogrilla2(3, 3) = "N"
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
    
    Grid2.Cols = 4
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
    Dim cSql2 As New rdoQuery
    Dim rut As String
    Dim PASO As String
    Dim LINEAS As Double
        Set cSql2.ActiveConnection = conta
        cSql2.sql = "SELECT codigo,glosa "
        cSql2.sql = cSql2.sql + "FROM capitalpropio_titulos "
        cSql2.sql = cSql2.sql + "order by codigo"
        cSql2.Execute
        Grid1.AutoRedraw = False
        Grid1.Rows = 1
        If cSql2.RowsAffected > 0 Then
        Set resultados2 = cSql2.OpenResultset
        While Not resultados2.EOF
        Grid1.Rows = Grid1.Rows + 1
        Grid1.Range(Grid1.Rows - 1, 1, Grid1.Rows - 1, 2).FontBold = True
        Grid1.Cell(Grid1.Rows - 1, 1).text = resultados2(0)
        Grid1.Cell(Grid1.Rows - 1, 2).text = resultados2(1)
        Grid1.Cell(Grid1.Rows - 1, 0).text = "TITULO"
        Call leeCAPITALDETALLE(resultados2(0))
        ' Grid1.Cell(Grid1.Rows - 1, 3).text = leersaldomayor(resultados2(0), Format(fechasistema, "yyyy-mm-dd"))
        
        resultados2.MoveNext
        Wend
          
          resultados2.Close
            Set resultados2 = Nothing

        End If
    
    Grid1.AutoRedraw = True
        Grid1.Refresh
        
NIVEL.Caption = Grid1.Cell(1, 1).text
LBLNIVEL.Caption = Grid1.Cell(1, 2).text
    
    

End Sub

Sub leeCAPITALDETALLE(codigo)

Dim resultados2 As rdoResultset
    Dim cSql2 As New rdoQuery
    Dim rut As String
    Dim PASO As String
    Dim LINEAS As Double
        Set cSql2.ActiveConnection = db
        cSql2.sql = "SELECT cpd.codigo,cm.nombre "
        cSql2.sql = cSql2.sql + "FROM capitalpropio_detalle as cpd left join  cuentasdelmayor as cm on cpd.codigo=cm.codigo and cm.año='" + Format(fechasistema, "yyyy") + "' "
        cSql2.sql = cSql2.sql + " where cpd.codigotitulo='" + codigo + "' "
        cSql2.sql = cSql2.sql + "order by cpd.codigo"
        cSql2.Execute
        LINEAS = 0
        If cSql2.RowsAffected > 0 Then
        Set resultados2 = cSql2.OpenResultset
        While Not resultados2.EOF
        Grid1.Rows = Grid1.Rows + 1
        Grid1.Cell(Grid1.Rows - 1, 1).text = resultados2(0)
        Grid1.Cell(Grid1.Rows - 1, 2).text = resultados2(1)
        Rem Grid2.Cell(Grid2.Rows - 1, 3).text = leersaldomayor(resultados2(0), Format(fechasistema, "yyyy-mm-dd"))
        resultados2.MoveNext
        Wend
          
          resultados2.Close
            Set resultados2 = Nothing

        End If
    
    
    
    

End Sub


Sub leeplan()

Dim resultados2 As rdoResultset
    Dim cSql2 As New rdoQuery
    Dim rut As String
    Dim PASO As String
    Dim LINEAS As Double
        Set cSql2.ActiveConnection = db
        cSql2.sql = "SELECT codigo,nombre "
        cSql2.sql = cSql2.sql + "FROM cuentasdelmayor where año='" + Format(fechasistema, "yyyy") + "' and tipo<'3' and mid(codigo,3,6)<>'000000' "
        cSql2.sql = cSql2.sql + "order by codigo"
        cSql2.Execute
        LINEAS = 0
        Grid2.AutoRedraw = False
        
        
        Grid2.Rows = 1
        If cSql2.RowsAffected > 0 Then
        Set resultados2 = cSql2.OpenResultset
        While Not resultados2.EOF
        Grid2.Rows = Grid2.Rows + 1
        If Mid(resultados2(0), 5, 4) = "0000" Then
        Grid2.Range(Grid2.Rows - 1, 1, Grid2.Rows - 1, 2).FontBold = True
        
        End If
        
        Grid2.Cell(Grid2.Rows - 1, 1).text = resultados2(0)
        Grid2.Cell(Grid2.Rows - 1, 2).text = resultados2(1)
        Rem Grid2.Cell(Grid2.Rows - 1, 3).text = leersaldomayor(resultados2(0), Format(fechasistema, "yyyy-mm-dd"))
        Grid2.Cell(Grid2.Rows - 1, 0).text = "0"
        If existeCAPITALDETALLE(resultados2(0)) = True Then
        Grid2.Range(Grid2.Rows - 1, 1, Grid2.Rows - 1, Grid2.Cols - 1).BackColor = &HFFC0C0
        Grid2.Cell(Grid2.Rows - 1, 0).text = "1"
        
        End If
        
        resultados2.MoveNext
        Wend
          
          resultados2.Close
           Set resultados2 = Nothing

        End If
        Grid2.AutoRedraw = True
        Grid2.Refresh
        For k = 1 To Grid2.Rows - 1
        If Grid2.Cell(k, 0).text = "0" Then Grid2.Cell(k, 1).SetFocus: k = Grid2.Rows - 1
        Next k
        
    End Sub

Private Sub Grid1_Click()
If Grid1.Cell(Grid1.ActiveCell.row, 0).text = "TITULO" Then
NIVEL.Caption = Grid1.Cell(Grid1.ActiveCell.row, 1).text
LBLNIVEL.Caption = Grid1.Cell(Grid1.ActiveCell.row, 2).text
End If

End Sub

Private Sub Grid1_DblClick()
Call eliminaCAPITALDETALLE(Grid1.Cell(Grid1.ActiveCell.row, 1).text)
leecapital
leeplan

End Sub

Private Sub Grid2_DblClick()
If Mid(Grid2.Cell(Grid2.ActiveCell.row, 1).text, 3, 6) <> "000000" And Grid2.Cell(Grid2.ActiveCell.row, 0).text = "0" Then
If NIVEL.Caption <> "" Then
Call GRABAR(Grid2.Cell(Grid2.ActiveCell.row, 1).text, NIVEL.Caption)
ROW1 = Grid2.ActiveCell.row
leecapital
leeplan

End If

End If

End Sub
Sub GRABAR(codigo, codigotitulo)
    campos(0, 0) = "codigo"
    campos(1, 0) = "codigotitulo"
    campos(2, 0) = ""
   
    campos(0, 1) = codigo
    campos(1, 1) = codigotitulo
  
    campos(0, 2) = "capitalpropio_detalle"
    op = 2
    sqlconta.response = campos
    Set sqlconta.conexion = db
    Call sqlconta.sqlconta(op, condicion)
    If Mid(codigo, 5, 4) = "0000" Then
    Call eliminasubCAPITALDETALLE(codigo)
    
    End If
    
End Sub

Sub eliminaCAPITALDETALLE(codigo)

Dim resultados2 As rdoResultset
    Dim cSql2 As New rdoQuery
    Dim rut As String
    Dim PASO As String
    Dim LINEAS As Double
        Set cSql2.ActiveConnection = db
        cSql2.sql = "delete FROM capitalpropio_detalle "
        cSql2.sql = cSql2.sql + " where codigo='" + codigo + "' "
        cSql2.Execute
        
End Sub

Sub eliminasubCAPITALDETALLE(codigo)

Dim resultados2 As rdoResultset
    Dim cSql2 As New rdoQuery
    Dim rut As String
    Dim PASO As String
    Dim LINEAS As Double
        Set cSql2.ActiveConnection = db
        cSql2.sql = "delete FROM capitalpropio_detalle "
        cSql2.sql = cSql2.sql + " where mid(codigo,1,4)='" + Mid(codigo, 1, 4) + "' and mid(codigo,5,4)<>'0000'  "
        cSql2.Execute
        
End Sub

Public Function existeCAPITALDETALLE(codigo) As Boolean


Dim resultados2 As rdoResultset
    Dim cSql2 As New rdoQuery
    Dim rut As String
    Dim PASO As String
    Dim LINEAS As Double
        Set cSql2.ActiveConnection = db
        cSql2.sql = "select * FROM capitalpropio_detalle "
        cSql2.sql = cSql2.sql + " where codigo='" + codigo + "' "
        cSql2.Execute
        If cSql2.RowsAffected > 0 Then
        existeCAPITALDETALLE = True
        Else
        existeCAPITALDETALLE = False
        End If
        
        
        Set cSql2.ActiveConnection = db
        cSql2.sql = "select * FROM capitalpropio_detalle "
        cSql2.sql = cSql2.sql + " where codigo='" + Mid(codigo, 1, 4) + "0000" + "' "
        cSql2.Execute
        If cSql2.RowsAffected > 0 Then
        existeCAPITALDETALLE = True
        End If
        
        
        
End Function


Private Sub Grid2_LeaveCell(ByVal row As Long, ByVal col As Long, NewRow As Long, NewCol As Long, Cancel As Boolean)
ROW1 = NewRow
End Sub
