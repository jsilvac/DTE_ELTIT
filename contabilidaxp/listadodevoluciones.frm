VERSION 5.00
Object = "{4777436C-EB8C-4596-98A8-EBCF98683969}#1.0#0"; "Flexcell.ocx"
Object = "{ADBBAED6-B16F-40DC-80DF-B44910CBA76C}#1.0#0"; "Frame-Xp.ocx"
Begin VB.Form prove0005 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "DEVOLUCIONES"
   ClientHeight    =   9900
   ClientLeft      =   2040
   ClientTop       =   1305
   ClientWidth     =   13410
   DrawStyle       =   5  'Transparent
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   660
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   894
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin XPFrame.FrameXp FrameXp2 
      Height          =   8160
      Left            =   90
      TabIndex        =   3
      Top             =   1710
      Width           =   13065
      _ExtentX        =   23045
      _ExtentY        =   14393
      BackColor       =   16761024
      Caption         =   "Listado de Guias de Devolucion"
      CaptionEstilo3D =   1
      BackColor       =   16761024
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   1
      Begin VB.CommandButton Command1 
         Caption         =   "IMPRIMIR"
         Height          =   375
         Left            =   5670
         TabIndex        =   14
         Top             =   7695
         Width           =   2085
      End
      Begin FlexCell.Grid Grid1 
         Height          =   7260
         Left            =   90
         TabIndex        =   4
         Top             =   315
         Width           =   12840
         _ExtentX        =   22648
         _ExtentY        =   12806
         Cols            =   5
         DefaultFontSize =   8.25
         Rows            =   30
      End
   End
   Begin XPFrame.FrameXp FrameXp1 
      Height          =   1500
      Left            =   90
      TabIndex        =   2
      Top             =   45
      Width           =   13095
      _ExtentX        =   23098
      _ExtentY        =   2646
      BackColor       =   16744576
      Caption         =   "DATOS DEVOLUCION"
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
      Alignment       =   1
      Begin VB.CommandButton Command2 
         Caption         =   "PROCESAR"
         Height          =   330
         Left            =   7335
         TabIndex        =   15
         Top             =   1035
         Width           =   2220
      End
      Begin XPFrame.FrameXp frmrut 
         Height          =   600
         Left            =   4860
         TabIndex        =   9
         Top             =   270
         Width           =   8070
         _ExtentX        =   14235
         _ExtentY        =   1058
         BackColor       =   16761024
         Caption         =   "Datos Proveedor"
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
         Enabled         =   0   'False
         Begin VB.TextBox dato3 
            Alignment       =   1  'Right Justify
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
            Left            =   1800
            MaxLength       =   9
            TabIndex        =   10
            Tag             =   "rut"
            Top             =   270
            Width           =   1095
         End
         Begin VB.Label dv 
            Alignment       =   1  'Right Justify
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
            ForeColor       =   &H00FF0000&
            Height          =   285
            Left            =   3060
            TabIndex        =   13
            Top             =   270
            Width           =   255
         End
         Begin VB.Label Label2 
            BackColor       =   &H00FFC0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " Rut Proveedor"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   255
            Left            =   225
            TabIndex        =   12
            Top             =   270
            Width           =   1530
         End
         Begin VB.Label lblnombreproveedor 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   3375
            TabIndex        =   11
            Top             =   270
            Width           =   4455
         End
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00FF8080&
         Caption         =   "Individual"
         Height          =   240
         Left            =   2115
         TabIndex        =   6
         Top             =   1035
         Width           =   1680
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00FF8080&
         Caption         =   "Todos"
         Height          =   240
         Left            =   225
         TabIndex        =   5
         Top             =   1035
         Value           =   -1  'True
         Width           =   1725
      End
      Begin XPFrame.FrameXp FrameXp4 
         Height          =   675
         Left            =   135
         TabIndex        =   7
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
            TabIndex        =   8
            Top             =   270
            Width           =   4395
         End
      End
   End
   Begin VB.PictureBox MANUAL 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H80000002&
      ForeColor       =   &H80000008&
      Height          =   0
      Left            =   0
      ScaleHeight     =   0
      ScaleWidth      =   13380
      TabIndex        =   1
      Top             =   9900
      Width           =   13410
   End
   Begin VB.TextBox pivote 
      Height          =   285
      Left            =   8415
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   4230
      Visible         =   0   'False
      Width           =   615
   End
End
Attribute VB_Name = "prove0005"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public saldoglobal As Double
Private LOCALFILTRO As String

Private MODIFI As Integer

Private Sub codigo_Click()
    Call dato1_KeyDown(vbKeyF2, 0)
End Sub
 Private Sub impRimir()
If Grid1.Rows > 1 Then
Call Titulos("LISTADO DE SALDOS RELACIONADOS ")
Grid1.PageSetup.Orientation = cellPortrait
Grid1.PageSetup.HeaderMargin = 0.5
Grid1.PageSetup.PrintFixedRow = True
Grid1.PageSetup.TopMargin = 3
Grid1.PageSetup.LeftMargin = 0.1
Grid1.PageSetup.RightMargin = 0.1
Grid1.PageSetup.BottomMargin = 3
Grid1.PageSetup.FooterMargin = 2
Grid1.PageSetup.BlackAndWhite = True

Grid1.Range(0, 1, 0, Grid1.Cols - 1).Borders(cellEdgeTop) = cellThin
Grid1.Range(0, 1, 0, Grid1.Cols - 1).Borders(cellEdgeBottom) = cellThin
Grid1.Range(0, 1, 0, Grid1.Cols - 1).Borders(cellEdgeLeft) = cellThin
Grid1.Range(0, 1, 0, Grid1.Cols - 1).Borders(cellEdgeRight) = cellThin
Grid1.Range(0, 1, 0, Grid1.Cols - 1).Borders(cellInsideVertical) = cellThin
Grid1.Range(0, 1, 0, Grid1.Cols - 1).Borders(cellInsideHorizontal) = cellThin
Grid1.PrintPreview
End If
End Sub
Sub Titulos(titulo1)

    Dim i As Integer
    Dim objReportTitle As FlexCell.ReportTitle
    
    Grid1.FixedRowColStyle = Fixed3D
    Grid1.CellBorderColorFixed = vbButtonShadow
    Grid1.ShowResizeTips = False
    Grid1.ReportTitles.Clear
    Grid1.PageSetup.CenterHorizontally = True
    Grid1.PageSetup.Orientation = cellLandscape
    Grid1.PageSetup.PrintTitleRows = 0
    
    'ENCABEZADO DE PAGINA
    Grid1.PageSetup.Header = nombreempresa & vbCrLf & direccionempresa & vbCrLf & comunaempresa & vbCrLf & rutempresa
    Grid1.PageSetup.HeaderAlignment = CellLeft
    Grid1.PageSetup.HeaderFont.Name = "Verdana"
    Grid1.PageSetup.HeaderFont.Size = 8
    
    'TITULOS DEL REPORTE
    
    Set objReportTitle = New FlexCell.ReportTitle
    objReportTitle.text = titulo1 & "  |  " & "EMITIDO  :  " & Format(fechasistema, "dd-MM-yyyy")
    objReportTitle.Font.Name = "Verdana" '"Times New Roman"
    objReportTitle.Font.Size = 8
    objReportTitle.Font.Bold = True
    objReportTitle.Align = cellCenter
    objReportTitle.PrintOnAllPages = True
    Grid1.ReportTitles.Add objReportTitle
    
    Set objReportTitle = New FlexCell.ReportTitle
    objReportTitle.text = dato2.text + "-" + dv.Caption & "  " & dato4.text
    objReportTitle.Font.Name = "Verdana" '"Times New Roman"
    objReportTitle.Font.Size = 8
    objReportTitle.Font.Bold = True
    objReportTitle.Align = cellCenter
    objReportTitle.PrintOnAllPages = True
    Grid1.ReportTitles.Add objReportTitle

    
    'PIE DE PAGINA
    Grid1.PageSetup.Footer = "Pág &P de &N" & vbCrLf & "Fecha: &D" & vbCrLf & "Usuario: " & USUARIOSISTEMA
    Grid1.PageSetup.FooterAlignment = cellRight
    Grid1.PageSetup.FooterFont.Name = "Verdana"
    Grid1.PageSetup.FooterFont.Size = 7
    
End Sub


Private Sub dato2_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF2 Then Call ayudactacte(dato4)
    Call flechas(dato1, dato4, KeyCode)
End Sub
 

Private Sub Command1_Click()
impRimir

End Sub

Private Sub Command2_Click()
Call LEERGUIAS


End Sub

Private Sub dato3_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF2 Then Call ayudactacte(dato3)

End Sub

Private Sub dato3_KeyPress(KeyAscii As Integer)
KeyAscii = esNumero(KeyAscii)
If KeyAscii = 13 Then
Call ceros(dato3)
dv.Caption = rut(dato3)
lblnombreproveedor.Caption = leerdatos(db, "cuentascorrientes", "nombre", "tipo='" + cuentaproveedor + "' and rut='" + dato3.text + dv.Caption + "' ")
If lblnombreproveedor.Caption = "" Then
dato3.SetFocus
Else
LEERGUIAS


End If



End If

End Sub

Private Sub Form_Load()
Call CENTRAR(Me)

    Call Conectar_BD
    sc = 0
  
Call CARGAPERMISO(Me.Name)
 
 CARGAGRILLADETALLE
LEErlocales

End Sub
Sub LEErlocales()
    Dim resultados As rdoResultset
    Dim Csql As New rdoQuery
    
        Set Csql.ActiveConnection = conta
        Csql.sql = "SELECT codigo,nombre "
        Csql.sql = Csql.sql + "FROM " + clientesistema + "gestion.g_maestroempresas WHERE codigocontable='" + empresaactiva + "' "
        Csql.sql = Csql.sql + "ORDER BY codigo "
        Csql.Execute
        
        If Csql.RowsAffected > 0 Then
            Set resultados = Csql.OpenResultset
            While Not resultados.EOF
                ComboLOCAL.AddItem (resultados(0) + " " + resultados(1))
                resultados.MoveNext
            Wend
            resultados.Close
            Set resultados = Nothing
        ComboLOCAL.text = ComboLOCAL.List(0)
        End If
        LOCALFILTRO = Mid(ComboLOCAL.List(0), 1, 2)
        
End Sub



Sub flechas(ByRef caja As TextBox, ByRef caja1 As TextBox, ByRef codigo As Integer)
    If codigo = 38 And caja.Enabled = True Then caja.SetFocus
    If codigo = 40 And caja1.Enabled = True Then caja1.SetFocus
End Sub

Sub cargatexto(ByRef caja As TextBox)
caja.SelStart = 0: caja.SelLength = Len(caja.text)
End Sub


Sub ayudactacte(ByRef caja As TextBox)
    Dim CAMPOS As Variant
    Dim cfijo As Variant
    Dim largo As Variant
    CAMPOS = Array("rut", "nombre")
    largo = Array("12n", "40s")
    cfijo = "tipo='" & cuentaproveedor & "' and año='" + Format(fechasistema, "yyyy") + "'"
    cabezas = Array("rut", "nombre")
    mensajeAyuda = "Ayuda Cuentas Corrientes"
    
    Call cargaAyudaT(servidor, basebus, usuario, password, "cuentascorrientes", pivote, CAMPOS, cfijo, largo, 2)

    If Val(pivote.text) = 0 Then dato3.SetFocus: GoTo no
    dato3.text = Mid(pivote.text, 1, 9)
    dv.Caption = Mid(pivote.text, 10, 1)
    caja.Enabled = True
    caja.SetFocus
no:

End Sub

 
Sub CARGAGRILLADETALLE()
    Dim formatogrilla2(10, 10)
    formatogrilla2(1, 1) = "LOCAL"
    formatogrilla2(1, 2) = "RUT"
    formatogrilla2(1, 3) = "NOMBRE"
    formatogrilla2(1, 4) = "FECHA"
    formatogrilla2(1, 5) = "NUMERO"
    formatogrilla2(1, 6) = "MONTO"
    formatogrilla2(1, 7) = "REBAJADA"
    Rem LARGO DE LOS DATOS
    formatogrilla2(2, 1) = "3"
    formatogrilla2(2, 2) = "10"
    formatogrilla2(2, 3) = "30"
    formatogrilla2(2, 4) = "10"
    formatogrilla2(2, 5) = "10"
    formatogrilla2(2, 6) = "10"
    formatogrilla2(2, 7) = "10"
    
    Rem TIPO DE DATOS C=CEROS,S=STRING,N=NUMERICO,D=DATE,
    formatogrilla2(3, 1) = "N"
    formatogrilla2(3, 2) = "S"
    formatogrilla2(3, 3) = "S"
    formatogrilla2(3, 4) = "D"
    formatogrilla2(3, 5) = "N"
    formatogrilla2(3, 6) = "N"
    formatogrilla2(3, 7) = "N"
    
    Rem FORMATO GRILLA
    formatogrilla2(4, 6) = " ###,###,##0"
    formatogrilla2(4, 7) = " ###,###,##0"
    
    Rem LOCCKED
    formatogrilla2(5, 1) = "TRUE"
    formatogrilla2(5, 2) = "TRUE"
    formatogrilla2(5, 3) = "TRUE"
    formatogrilla2(5, 4) = "TRUE"
    formatogrilla2(5, 5) = "TRUE"
    formatogrilla2(5, 6) = "TRUE"
    formatogrilla2(5, 7) = "TRUE"
    
    
    Rem VALOR MAXIMO
    
    Grid1.Cols = 8
    Grid1.Rows = 1
    Grid1.AllowUserResizing = False
    Grid1.DisplayFocusRect = False
    Grid1.ExtendLastCol = True
    Grid1.BoldFixedCell = False
    Grid1.DrawMode = cellOwnerDraw
    Grid1.Appearance = Flat
    Grid1.ScrollBarStyle = Flat
    Grid1.FixedRowColStyle = Flat
'    Grid1.BackColorFixed = RGB(90, 158, 214)
'    Grid1.BackColorFixedSel = RGB(110, 180, 230)
'    Grid1.BackColorBkg = RGB(90, 158, 214)
'    Grid1.BackColorScrollBar = RGB(231, 235, 247)
'    Grid1.BackColor1 = RGB(231, 235, 247)
'    Grid1.BackColor2 = RGB(239, 243, 255)
'    Grid1.GridColor = RGB(148, 190, 231)
    Grid1.Column(0).Width = 0
    
    For K = 1 To Grid1.Cols - 1
        Grid1.Cell(0, K).text = formatogrilla2(1, K)
        Grid1.Column(K).Width = Val(formatogrilla2(2, K)) * 8
        Grid1.Column(K).MaxLength = Val(formatogrilla2(2, K))
        Grid1.Column(K).FormatString = formatogrilla2(4, K)
        Grid1.Column(K).Locked = formatogrilla2(5, K)
        If formatogrilla2(3, K) = "N" Then Grid1.Column(K).Alignment = cellRightCenter
        If formatogrilla2(3, K) = "S" Then Grid1.Column(K).Alignment = cellLeftTop
        If formatogrilla2(3, K) = "D" Then Grid1.Column(K).CellType = cellCalendar
    Next K
 
    End Sub
 
Sub verdetalle(loc, numero)
Dim Csql As New rdoQuery
Dim resultados As rdoResultset
Dim tipo As String
tipo = "DM"

Set Csql.ActiveConnection = db

Csql.sql = "select linea,codigo,descripcion,cantidad,uxc,unidades,precio,descuento,total "
Csql.sql = Csql.sql & "from " & clientesistema & "gestion" & leerrubro(dato1.text) & ".l_movimientos_detalle_" & loc & " where tipo='" & tipo & "' and numero='" & numero & "' order by linea"
Csql.Execute

If Csql.RowsAffected > 0 Then
    Grid1.Rows = Csql.RowsAffected + 1
    Set resultados = Csql.OpenResultset
    
    While Not resultados.EOF
        Grid1.Cell(resultados(0), 1).text = resultados(1)
        Grid1.Cell(resultados(0), 2).text = resultados(2)
        Grid1.Cell(resultados(0), 3).text = resultados(3)
        Grid1.Cell(resultados(0), 4).text = resultados(4)
        Grid1.Cell(resultados(0), 5).text = resultados(5)
        Grid1.Cell(resultados(0), 6).text = resultados(6)
        Grid1.Cell(resultados(0), 7).text = resultados(7)
        Grid1.Cell(resultados(0), 8).text = resultados(8)
        resultados.MoveNext
    Wend
End If

Csql.Close
Set Csql = Nothing
Set resultados = Nothing
 
End Sub
Function leerrubro(loc) As String
    Dim Csql As New rdoQuery
    Dim resultado As rdoResultset
    
    Set Csql.ActiveConnection = db
    Csql.sql = "select rubro from " & clientesistema & "gestion.g_maestroempresas where "
    Csql.sql = Csql.sql & "codigo='" & loc & "' "
    Csql.Execute
    
 If Csql.RowsAffected > 0 Then
    Set resultado = Csql.OpenResultset
    leerrubro = resultado(0)
 End If
 Csql.Close
 Set Csql = Nothing
 Set resultado = Nothing
 
End Function


Sub LEERGUIAS()
Dim Csql As New rdoQuery
Dim resultados As rdoResultset
Dim tipo As String
Dim rutpaso As String
Dim totales(2) As Double
Dim totales2(2) As Double

tipo = "DM"

Set Csql.ActiveConnection = db

Csql.sql = "select dp.local,dp.rut,cc.nombre,dp.fecha,dp.numero,dp.monto,dp.montoco "
Csql.sql = Csql.sql & "from devoluciones_proveedores as dp left join cuentascorrientes as cc on (dp.rut=cc.rut and cc.tipo='" + cuentaproveedor + "' AND cc.año='" + Format(fechasistema, "yyyy") + "') "
If Option2.Value = True Then
Csql.sql = Csql.sql & "where dp.rut='" & dato3.text + dv.Caption & "' "
End If
Csql.sql = Csql.sql & "order by cc.nombre "
Csql.Execute
  Grid1.Rows = 1
If Csql.RowsAffected > 0 Then
  
    Set resultados = Csql.OpenResultset
    rutpaso = resultados(1)
    While Not resultados.EOF
        Grid1.Rows = Grid1.Rows + 1
        
        If rutpaso <> resultados(1) Then
        Grid1.Range(Grid1.Rows - 1, 1, Grid1.Rows - 1, Grid1.Cols - 1).Borders(cellEdgeTop) = cellThin
        Grid1.Cell(Grid1.Rows - 1, 3).text = "TOTAL PROVEEDOR "
        Grid1.Cell(Grid1.Rows - 1, 6).text = totales(1)
        Grid1.Cell(Grid1.Rows - 1, 7).text = totales(2)
        rutpaso = resultados(1)
        Grid1.Rows = Grid1.Rows + 2
        totales(1) = 0
        totales(2) = 0
        
        End If
        
        Grid1.Cell(Grid1.Rows - 1, 1).text = resultados(0)
        Grid1.Cell(Grid1.Rows - 1, 2).text = resultados(1)
        If IsNull(resultados(2)) = False Then
        Grid1.Cell(Grid1.Rows - 1, 3).text = resultados(2)
        Else
        Grid1.Cell(Grid1.Rows - 1, 3).text = "**** RUT NO EXISTE *****"
        
            End If
        
        Grid1.Cell(Grid1.Rows - 1, 4).text = resultados(3)
        Grid1.Cell(Grid1.Rows - 1, 5).text = resultados(4)
        Grid1.Cell(Grid1.Rows - 1, 6).text = resultados(5)
        Grid1.Cell(Grid1.Rows - 1, 7).text = resultados(6)
        totales(1) = totales(1) + resultados(5)
        totales(2) = totales(2) + resultados(6)
        totales2(1) = totales2(1) + resultados(5)
        totales2(2) = totales2(2) + resultados(6)
        
        resultados.MoveNext
    Wend
        Grid1.Rows = Grid1.Rows + 1
        Grid1.Range(Grid1.Rows - 1, 1, Grid1.Rows - 1, Grid1.Cols - 1).Borders(cellEdgeTop) = cellThin
        
        
        Grid1.Cell(Grid1.Rows - 1, 3).text = "TOTAL PROVEEDOR "
        Grid1.Cell(Grid1.Rows - 1, 6).text = totales(1)
        Grid1.Cell(Grid1.Rows - 1, 7).text = totales(2)
        Grid1.Rows = Grid1.Rows + 2
        totales(1) = 0
        totales(2) = 0
        
        Grid1.Range(Grid1.Rows - 1, 1, Grid1.Rows - 1, Grid1.Cols - 1).Borders(cellEdgeTop) = cellThin
        
        Grid1.Cell(Grid1.Rows - 1, 3).text = "TOTAL GENERAL "
        Grid1.Cell(Grid1.Rows - 1, 6).text = totales2(1)
        Grid1.Cell(Grid1.Rows - 1, 7).text = totales2(2)
        totales2(1) = 0
        totales2(2) = 0
        
        Grid1.Rows = Grid1.Rows + 1
        
End If

Csql.Close
Set Csql = Nothing
Set resultados = Nothing
 
End Sub


Private Sub Option1_Click()
frmrut.Enabled = False
LEERGUIAS
End Sub

Private Sub Option2_Click()

frmrut.Enabled = True



dato3.SetFocus

End Sub
