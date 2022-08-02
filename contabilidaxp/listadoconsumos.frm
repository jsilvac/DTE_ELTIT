VERSION 5.00
Object = "{4777436C-EB8C-4596-98A8-EBCF98683969}#1.0#0"; "FlexCell.ocx"
Object = "{ADBBAED6-B16F-40DC-80DF-B44910CBA76C}#1.0#0"; "Frame-Xp.ocx"
Begin VB.Form consumo05 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Maestro Consumos Basicos"
   ClientHeight    =   10275
   ClientLeft      =   2040
   ClientTop       =   1425
   ClientWidth     =   15240
   DrawStyle       =   5  'Transparent
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   685
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1016
   StartUpPosition =   2  'CenterScreen
   Begin XPFrame.FrameXp FrameQuickMenu 
      Height          =   615
      Left            =   12240
      TabIndex        =   12
      Top             =   0
      Width           =   3015
      _ExtentX        =   5318
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
         Width           =   1335
      End
      Begin VB.CommandButton botonmisfavoritos 
         Caption         =   "Mis Favoritos"
         Height          =   255
         Left            =   1560
         TabIndex        =   13
         Top             =   280
         Width           =   1335
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
      ScaleWidth      =   15210
      TabIndex        =   1
      Top             =   10275
      Width           =   15240
   End
   Begin XPFrame.FrameXp FrameXp3 
      Height          =   10170
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   15135
      _ExtentX        =   26696
      _ExtentY        =   17939
      BackColor       =   16761024
      Caption         =   "Servicios Vigentes"
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
         Caption         =   "GENERA INFORME"
         Height          =   255
         Left            =   3480
         TabIndex        =   11
         Top             =   360
         Width           =   1935
      End
      Begin VB.TextBox dato1 
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
         Left            =   2100
         MaxLength       =   2
         TabIndex        =   0
         Tag             =   "fecha"
         Top             =   360
         Width           =   510
      End
      Begin VB.CommandButton Command1 
         Caption         =   "IMPRIMIR"
         Height          =   255
         Left            =   5520
         TabIndex        =   4
         Top             =   9840
         Width           =   2535
      End
      Begin FlexCell.Grid Grid2 
         Height          =   8640
         Left            =   0
         TabIndex        =   3
         Top             =   1080
         Width           =   14775
         _ExtentX        =   26061
         _ExtentY        =   15240
         Cols            =   5
         DefaultFontSize =   8.25
         Rows            =   30
      End
      Begin XPFrame.FrameXp FrameXp6 
         Height          =   735
         Left            =   5760
         TabIndex        =   7
         Top             =   240
         Width           =   3975
         _ExtentX        =   7011
         _ExtentY        =   1296
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
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   240
            TabIndex        =   8
            Top             =   240
            Width           =   3615
         End
      End
      Begin XPFrame.FrameXp FrameXp7 
         Height          =   735
         Left            =   9720
         TabIndex        =   9
         Top             =   240
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   1296
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
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   240
            TabIndex        =   10
            Top             =   240
            Width           =   2055
         End
      End
      Begin VB.Label LBLTIPO 
         BackColor       =   &H80000007&
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
         ForeColor       =   &H0000FFFF&
         Height          =   330
         Left            =   120
         TabIndex        =   6
         Top             =   720
         Width           =   5505
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   " TIPO CONSUMO"
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
         TabIndex        =   5
         Top             =   360
         Width           =   1905
      End
   End
End
Attribute VB_Name = "consumo05"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public saldoglobal As Double
Private moneda As String
Private rutpropi As String

Private MODIFI As Integer



Private Sub COMBOAÑO_Change()
leerCONSUMOS
End Sub

Private Sub COMBOMES_Change()
leerCONSUMOS
End Sub

Private Sub Command1_Click()
Titulos
Grid2.PrintPreview

End Sub
Sub Titulos()
    

    Dim i As Integer
    Dim objReportTitle As FlexCell.ReportTitle
    
    Grid2.FixedRowColStyle = Fixed3D
    Grid2.CellBorderColorFixed = vbButtonShadow
    Grid2.ShowResizeTips = False
    Grid2.PageSetup.Orientation = cellPortrait
    Grid2.DefaultFont.Size = 7.5
    Grid2.PageSetup.PrintFixedRow = True
    Grid2.ReportTitles.Clear
    Grid2.PageSetup.CenterHorizontally = True
    Grid2.PageSetup.PrintTitleRows = 0
    Grid2.PageSetup.BlackAndWhite = False
    Grid2.PageSetup.Orientation = cellLandscape
    'Logo
  
    'ENCABEZADO DE PAGINA
    Grid2.PageSetup.Header = nombreempresa & vbCrLf & direccionempresa & vbCrLf & comunaempresa
    Grid2.PageSetup.HeaderAlignment = CellLeft
    Grid2.PageSetup.HeaderFont.Name = "Verdana"
    Grid2.PageSetup.HeaderFont.Size = 8
    'TITULOS DEL REPORTE
    Set objReportTitle = New FlexCell.ReportTitle
    objReportTitle.text = "LISTADO DE CONSUMOS Y SUS ESTADOS"
    objReportTitle.Font.Name = "Verdana" '"Times New Roman"
    objReportTitle.Font.Size = 10
    objReportTitle.Font.Bold = True
    objReportTitle.Align = cellCenter
    objReportTitle.PrintOnAllPages = True
    Grid2.ReportTitles.Add objReportTitle
        
    Set objReportTitle = New FlexCell.ReportTitle
    objReportTitle.Font.Name = "Verdana" '"Times New Roman"
    objReportTitle.Font.Size = 7
    objReportTitle.Font.Underline = True
    objReportTitle.Font.Bold = True
    objReportTitle.Align = cellCenter
    objReportTitle.PrintOnAllPages = True
    Grid2.ReportTitles.Add objReportTitle
    
    
    'PIE DE PAGINA
    Grid2.PageSetup.Footer = "Pág &P de &N" & vbCrLf & "Fecha: &D " & vbCrLf & "Usuario:" & USUARIOSISTEMA
    Grid2.PageSetup.FooterAlignment = cellRight
    Grid2.PageSetup.FooterFont.Name = "Verdana"
    Grid2.PageSetup.FooterFont.Size = 7
    Grid2.PageSetup.LeftMargin = 0.5
    Grid2.PageSetup.RightMargin = 0.5
    
    Grid2.Range(0, 1, 0, Grid2.Cols - 1).Borders(cellEdgeLeft) = cellThick
    Grid2.Range(0, 1, 0, Grid2.Cols - 1).Borders(cellEdgeTop) = cellThick
    Grid2.Range(0, 1, 0, Grid2.Cols - 1).Borders(cellEdgeBottom) = cellThick
    Grid2.Range(0, 1, 0, Grid2.Cols - 1).Borders(cellEdgeRight) = cellThick
    Grid2.Range(0, 1, 0, Grid2.Cols - 1).Borders(cellInsideHorizontal) = cellThick
    Grid2.Range(0, 1, 0, Grid2.Cols - 1).Borders(cellInsideVertical) = cellThick
    
    
    
    
    
End Sub

Private Sub COMMAND2_Click()
leerCONSUMOS

End Sub

Private Sub Form_Load()
Call CENTRAR(Me)
    Call Conectar_BD
    Rem Call Funciones_Forms_M_Productos.Conecta_Maestro_Productos
    sc = 0
 
Rem Call RECUPERAFECHA
For k = 1 To 12
COMBOMES.AddItem MonthName(k)
Next k
COMBOMES.ListIndex = CDbl(Format(fechasistema, "mm") - 1)
For k = 2000 To Val(Format(fechasistema, "yyyy"))
COMBOAÑO.AddItem k
Next k
COMBOAÑO.ListIndex = k - 2001

Call CARGAPERMISO(Me.Name)

Call CARGAGRILLA2
leerCONSUMOS

End Sub



Sub flechas(ByRef caja As TextBox, ByRef caja1 As TextBox, ByRef codigo As Integer)
    If codigo = 38 And caja.Enabled = True Then caja.SetFocus
    If codigo = 40 And caja1.Enabled = True Then caja1.SetFocus
End Sub

 
Sub cargatexto(ByRef caja As TextBox)
caja.SelStart = 0: caja.SelLength = Len(caja.text)
End Sub

Private Sub opciones_GotFocus()
MANUAL.SetFocus
End Sub

 

Sub CARGAGRILLA2()
    Dim formatogrilla2(10, 12)
    formatogrilla2(1, 1) = "TIPO"
    formatogrilla2(1, 2) = "SERVICIO"
    formatogrilla2(1, 3) = "PROVEEDOR"
    formatogrilla2(1, 4) = "EMPRESA"
    formatogrilla2(1, 5) = "DIA/PAGO"
    formatogrilla2(1, 6) = "UBICACION"
    formatogrilla2(1, 7) = "TD"
    formatogrilla2(1, 8) = "NUMERO"
    formatogrilla2(1, 9) = "FECHA"
    formatogrilla2(1, 10) = "MONTO"
    formatogrilla2(1, 11) = "VENCE"
    formatogrilla2(1, 12) = "PAGADO"
    
    Rem LARGO DE LOS DATOS
    formatogrilla2(2, 1) = "5"
    formatogrilla2(2, 2) = "8"
    formatogrilla2(2, 3) = "15"
    formatogrilla2(2, 4) = "15"
    formatogrilla2(2, 5) = "8"
    formatogrilla2(2, 6) = "15"
    formatogrilla2(2, 7) = "2"
    formatogrilla2(2, 8) = "8"
    formatogrilla2(2, 9) = "8"
    formatogrilla2(2, 10) = "8"
    formatogrilla2(2, 11) = "8"
    formatogrilla2(2, 12) = "5"
    
    Rem TIPO DE DATOS C=CEROS,S=STRING,N=NUMERICO,D=DATE,
    formatogrilla2(3, 1) = "S"
    formatogrilla2(3, 2) = "S"
    formatogrilla2(3, 3) = "S"
    formatogrilla2(3, 4) = "S"
    formatogrilla2(3, 5) = "S"
    formatogrilla2(3, 6) = "S"
    formatogrilla2(3, 7) = "S"
    formatogrilla2(3, 8) = "S"
    formatogrilla2(3, 9) = "D"
    formatogrilla2(3, 10) = "N"
    formatogrilla2(3, 11) = "D"
    formatogrilla2(3, 12) = "S"
    
    Rem FORMATO GRILLA
    
    formatogrilla2(4, 10) = " ###,###,##0"
    
    
    Rem LOCCKED
    formatogrilla2(5, 1) = "TRUE"
    formatogrilla2(5, 2) = "TRUE"
    formatogrilla2(5, 3) = "TRUE"
    formatogrilla2(5, 4) = "TRUE"
    formatogrilla2(5, 5) = "TRUE"
    formatogrilla2(5, 6) = "TRUE"
    formatogrilla2(5, 7) = "TRUE"
    formatogrilla2(5, 8) = "TRUE"
    formatogrilla2(5, 9) = "TRUE"
    formatogrilla2(5, 10) = "TRUE"
    formatogrilla2(5, 11) = "TRUE"
    formatogrilla2(5, 12) = "TRUE"
    
    
    Rem VALOR MAXIMO
    
    Grid2.Cols = 13
    Grid2.Rows = 1
    Grid2.AllowUserResizing = True
    Grid2.DisplayFocusRect = False
    Grid2.ExtendLastCol = True
    Grid2.BoldFixedCell = False
    Grid2.DrawMode = cellOwnerDraw
    Grid2.Appearance = Flat
    Grid2.ScrollBarStyle = Flat
    Grid2.FixedRowColStyle = Flat
    Grid2.BackColorFixed = RGB(90, 158, 214)
    Grid2.BackColorFixedSel = RGB(110, 180, 230)
    Grid2.BackColorBkg = RGB(90, 158, 214)
    Grid2.BackColorScrollBar = RGB(231, 235, 247)
    Grid2.BackColor1 = RGB(231, 235, 247)
    Grid2.BackColor2 = RGB(239, 243, 255)
    Grid2.GridColor = RGB(148, 190, 231)
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
   
    
    Grid2.Column(12).CellType = cellCheckBox
    
    
    
    
    
    End Sub


 Public Sub leerCONSUMOS()
 Dim csql As New rdoQuery
 Dim resultados As rdoResultset
CARGAGRILLA2
 Set csql.ActiveConnection = contadb
 csql.sql = "select * from " & clientesistema & "consumos_basicos.maestro_unidades_consumo "
 csql.sql = csql.sql + "where tipo='" + dato1.text + "' "
 csql.sql = csql.sql + "order by empresacontable "
 csql.Execute
 Grid2.Rows = 1
 If csql.RowsAffected > 0 Then
    Set resultados = csql.OpenResultset
    While resultados.EOF = False
    Grid2.Rows = Grid2.Rows + 1
    Grid2.Cell(Grid2.Rows - 1, 1).text = resultados(0) + "=" + leetipoconsumo(resultados(0))
    Grid2.Cell(Grid2.Rows - 1, 2).text = resultados(1)
    Grid2.Cell(Grid2.Rows - 1, 3).text = leerProveedor(resultados(2))
    Grid2.Cell(Grid2.Rows - 1, 4).text = resultados(3) + "=" + leerempresa(resultados(3))
    Grid2.Cell(Grid2.Rows - 1, 5).text = Format(resultados(4), "00") + "/" + COMBOMES.text
    Grid2.Cell(Grid2.Rows - 1, 6).text = resultados(6)
    
    Call leerFACTURAS(resultados(0), resultados(1))




    resultados.MoveNext



    Wend


  End If
 csql.Close
 Set csql = Nothing
 Set resultados = Nothing

 End Sub
Public Sub leerFACTURAS(tipo, numero)
 Dim csql As New rdoQuery
 Dim resultados As rdoResultset
Dim mesa As String
mesa = Format(COMBOMES.ListIndex + 1, "00")
 Set csql.ActiveConnection = contadb
 csql.sql = "select tipodocumento,numerodocumento,fecha,montodocumento,fechacorte,tipocomprobante,rut from " & clientesistema & "consumos_basicos.detalle_servicios "
 csql.sql = csql.sql + "where tipo='" + tipo + "' and numeroservicio='" + numero + "' "
 csql.sql = csql.sql + "and mid(fecha,1,7)='" + COMBOAÑO.text + "-" + mesa + "' "
 
 csql.Execute
 If csql.RowsAffected > 0 Then
    Set resultados = csql.OpenResultset
    While resultados.EOF = False
    
    Grid2.Cell(Grid2.Rows - 1, 3).text = leerProveedor(resultados(6))
    Grid2.Cell(Grid2.Rows - 1, 7).text = resultados(0)
    Grid2.Cell(Grid2.Rows - 1, 8).text = resultados(1)
    
    Grid2.Cell(Grid2.Rows - 1, 9).text = resultados(2)
    Grid2.Cell(Grid2.Rows - 1, 10).text = resultados(3)
    If IsNull(resultados(4)) = False Then
    Grid2.Cell(Grid2.Rows - 1, 11).text = resultados(4)
    
    End If
    If resultados(5) <> "" Then
    Grid2.Cell(Grid2.Rows - 1, 12).text = "1"
    Else
    Grid2.Cell(Grid2.Rows - 1, 12).text = "0"
    
    End If
    
    




    resultados.MoveNext



    Wend


  End If
 csql.Close
 Set csql = Nothing
 Set resultados = Nothing

 End Sub


Private Function leemonedas(codigo) As String
Dim csql As New rdoQuery
Dim resultados As rdoResultset

Set csql.ActiveConnection = contadb

csql.sql = "select nombremoneda from " & clientesistema & "arriendos" & ".maestro_monedas where codigomoneda='" & codigo & "'"
csql.Execute
leemonedas = ""
If csql.RowsAffected > 0 Then
Set resultados = csql.OpenResultset
leemonedas = resultados(0)
End If
Set resultados = Nothing
csql.Close
Set csql = Nothing

End Function

Public Function LEERULTIMOFOLIOcontrato() As String

    Dim resultados As rdoResultset
    Dim csql As New rdoQuery
    
        Set csql.ActiveConnection = contadb

            csql.sql = "select IFNULL(max(numero),0) from " + clientesistema + "arriendos.contratos_arriendo"
            
            csql.Execute
    If csql.RowsAffected > 0 Then
    Set resultados = csql.OpenResultset
    
    
        LEERULTIMOFOLIOcontrato = Format(resultados(0) + 1, "0000000000")
    End If
    
End Function


Private Sub dato1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF2 Then Call ayudaTIPO(dato1)
    If KeyCode = 38 Then Unload Me
    
   
End Sub
Sub ayudaTIPO(ByRef caja As TextBox)
    Dim campos As Variant
    Dim cfijo As Variant
    Dim largo As Variant
    campos = Array("codigo", "nombre")
    cabezas = Array("CODIGO", "NOMBRE")
    largo = Array("4N", "40s")
    mensajeAyuda = "Ayuda Tipos de Consumos Basicos"
    cfijo = "no"
    
    Call cargaAyudaT(Servidor, clientesistema + "consumos_basicos", Usuario, password, "maestro_tipo_consumos", caja, campos, cfijo, largo, 2)
    If caja.text = "" Then caja.SetFocus: GoTo no
    caja.Enabled = True
    caja.SetFocus


no:

End Sub

Sub ayudaservicios(ByRef caja As TextBox)
    Dim campos As Variant
    Dim cfijo As Variant
    Dim largo As Variant
    campos = Array("numeroservicio", "ubicacion")
    cabezas = Array("CODIGO", "NOMBRE")
    largo = Array("20N", "40s")
    mensajeAyuda = "Ayuda tipo de creditos"
    cfijo = "tipo='" + dato1.text + "'"
    
    Call cargaAyudaT(Servidor, clientesistema + "consumos_basicos", Usuario, password, "maestro_unidades_consumo", caja, campos, cfijo, largo, 2)
    If caja.text = "" Then caja.SetFocus: GoTo no
    caja.Enabled = True
    caja.SetFocus


no:

End Sub


    
 
Sub ayudaempresa(ByRef caja As TextBox)
    Dim campos As Variant
    Dim cfijo As Variant
    Dim largo As Variant
    campos = Array("codigoempresa", "nombre")
    cabezas = Array("CODIGO", "NOMBRE")
    largo = Array("4N", "40s")
    mensajeAyuda = "Ayuda Empresas"
    cfijo = "no"
    
    Call cargaAyudaT(Servidor, clientesistema + "conta", Usuario, password, "maestroempresas", caja, campos, cfijo, largo, 2)
    If caja.text = "" Then caja.SetFocus: GoTo no
    caja.Enabled = True
    caja.SetFocus


no:

End Sub




Private Sub dato1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then Unload Me
    snum = 0: KeyAscii = esNumero(KeyAscii)
    If KeyAscii = 13 Then
        
    Call ceros(dato1)
    If leetipoconsumo(dato1.text) <> "" Then
    LBLTIPO.Caption = leetipoconsumo(dato1.text)
    leerCONSUMOS
    
    Else
    dato1.SetFocus
    End If
    End If
    
    
    
    
End Sub
Private Sub botonmisaccesos_Click()
programafiltro = Me.Caption
misaccesos.Show
End Sub


Private Sub botonmisfavoritos_Click()
Call AgregaFavoritos(USUARIOSISTEMA, App.EXEName, Me.Name, Me.Caption)
End Sub

