VERSION 5.00
Object = "{4777436C-EB8C-4596-98A8-EBCF98683969}#1.0#0"; "FlexCell.ocx"
Object = "{ADBBAED6-B16F-40DC-80DF-B44910CBA76C}#1.0#0"; "Frame-Xp.ocx"
Begin VB.Form activofijo03 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "LISTADO ACTIVO FIJO"
   ClientHeight    =   9900
   ClientLeft      =   2040
   ClientTop       =   1305
   ClientWidth     =   15270
   DrawStyle       =   5  'Transparent
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   660
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1018
   ShowInTaskbar   =   0   'False
   Begin XPFrame.FrameXp FrameQuickMenu 
      Height          =   615
      Left            =   12120
      TabIndex        =   16
      Top             =   0
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
         TabIndex        =   18
         Top             =   280
         Width           =   1455
      End
      Begin VB.CommandButton botonmisfavoritos 
         Caption         =   "Mis Favoritos"
         Height          =   255
         Left            =   1680
         TabIndex        =   17
         Top             =   280
         Width           =   1335
      End
   End
   Begin XPFrame.FrameXp FrameXp2 
      Height          =   8160
      Left            =   90
      TabIndex        =   3
      Top             =   1680
      Width           =   15105
      _ExtentX        =   26644
      _ExtentY        =   14393
      BackColor       =   16761024
      Caption         =   "Listado de Contratos de Publicidad  "
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
         BackColor       =   &H00FF8080&
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
         Height          =   375
         Left            =   5625
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   7650
         Width           =   2085
      End
      Begin FlexCell.Grid Grid1 
         Height          =   7260
         Left            =   90
         TabIndex        =   4
         Top             =   315
         Width           =   14880
         _ExtentX        =   26247
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
      Width           =   15135
      _ExtentX        =   26696
      _ExtentY        =   2646
      BackColor       =   16744576
      Caption         =   "DATOS "
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
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   4920
         TabIndex        =   14
         Top             =   960
         Width           =   2220
      End
      Begin XPFrame.FrameXp frmrut 
         Height          =   600
         Left            =   4860
         TabIndex        =   9
         Top             =   270
         Width           =   7110
         _ExtentX        =   12541
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
            MaxLength       =   2
            TabIndex        =   10
            Tag             =   "rut"
            Top             =   270
            Width           =   495
         End
         Begin VB.Label Label2 
            BackColor       =   &H00FFC0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " FAMILIA"
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
         Begin VB.Label lblnombreFAMILIA 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   2520
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
         Value           =   -1  'True
         Width           =   1680
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00FF8080&
         Caption         =   "Todos"
         Height          =   240
         Left            =   225
         TabIndex        =   5
         Top             =   1035
         Width           =   1725
      End
      Begin XPFrame.FrameXp FrameXp4 
         Height          =   675
         Left            =   135
         TabIndex        =   7
         Top             =   270
         Visible         =   0   'False
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
      Begin VB.Label Label1 
         BackColor       =   &H00FF8080&
         Caption         =   "DOBLE CLICK SOBRE GRILLA PARA VER DETALLE"
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
         Left            =   8760
         TabIndex        =   15
         Top             =   1200
         Width           =   3975
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
      ScaleWidth      =   15240
      TabIndex        =   1
      Top             =   9900
      Width           =   15270
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
Attribute VB_Name = "activofijo03"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public saldoglobal As Double
Private localfiltro As String

Private MODIFI As Integer

'Private Sub codigo_Click()
'    Call dato1_KeyDown(vbKeyF2, 0)
'End Sub
 Private Sub imprimir()
If Grid1.Rows > 1 Then
Call Titulos("LISTADO DE ACTIVOS FIJOS")
Grid1.PageSetup.Orientation = cellLandscape
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
    

    
    'PIE DE PAGINA
    Grid1.PageSetup.Footer = "Pág &P de &N" & vbCrLf & "Fecha: &D" & vbCrLf & "Usuario: " & USUARIOSISTEMA
    Grid1.PageSetup.FooterAlignment = cellRight
    Grid1.PageSetup.FooterFont.Name = "Verdana"
    Grid1.PageSetup.FooterFont.Size = 7
    
End Sub


'Private Sub dato2_KeyDown(KeyCode As Integer, Shift As Integer)
'    If KeyCode = vbKeyF2 Then Call ayudactacte(dato2)
'    Call flechas(dato1, dato4, KeyCode)
'End Sub
 

Private Sub Command1_Click()
imprimir

End Sub

Private Sub COMMAND2_Click()
Call LEERactivofijos
End Sub

Private Sub dato3_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF2 Then Call ayudafamilia(dato3)

End Sub

Private Sub dato3_KeyPress(KeyAscii As Integer)
KeyAscii = esNumero(KeyAscii)
If KeyAscii = 13 Then
Call ceros(dato3)
lblnombreFAMILIA.Caption = LEERNOMBREFAMILIA(dato3.text)
If lblnombreFAMILIA.Caption = "" Then
    dato3.SetFocus
Else
    LEERactivofijos
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
    Dim csql As New rdoQuery
    
        Set csql.ActiveConnection = contadb
        csql.sql = "SELECT codigo,nombre "
        csql.sql = csql.sql + "FROM " + clientesistema + "gestion.g_maestroempresas WHERE codigocontable='" + empresaactiva + "' "
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



Sub flechas(ByRef caja As TextBox, ByRef caja1 As TextBox, ByRef codigo As Integer)
    If codigo = 38 And caja.Enabled = True Then caja.SetFocus
    If codigo = 40 And caja1.Enabled = True Then caja1.SetFocus
End Sub

Sub cargatexto(ByRef caja As TextBox)
caja.SelStart = 0: caja.SelLength = Len(caja.text)
End Sub


Sub ayudafamilia(ByRef caja As TextBox)
    Dim campos As Variant
    Dim cfijo As Variant
    Dim largo As Variant
    campos = Array("codigo", "nombre")
    largo = Array("12n", "40s")
    cfijo = "no"
    cabezas = Array("Codigo", "Nombre")
    mensajeAyuda = "Ayuda Familias"
    
    Call cargaAyudaT(Servidor, clientesistema + "conta", Usuario, password, "maestro_familias", pivote, campos, cfijo, largo, 2)

    If Val(pivote.text) = 0 Then caja.SetFocus: GoTo no
     
    caja.text = pivote.text
    caja.Enabled = True
    caja.SetFocus

no:

End Sub


 
Sub CARGAGRILLADETALLE()
    Dim formatogrilla2(10, 40)
 
    
    formatogrilla2(1, 1) = "CODIGO"
    formatogrilla2(1, 2) = "NOMBRE"
    formatogrilla2(1, 3) = "CRCC"
    formatogrilla2(1, 4) = "FECHA P." & vbCrLf & "SERVICIO"
    formatogrilla2(1, 5) = "VALOR" & vbCrLf & "LIBRO"
    formatogrilla2(1, 6) = "DEPRECIACION" & vbCrLf & "ACUMULADA"
    formatogrilla2(1, 7) = "VALOR" & vbCrLf & "NETO"
    formatogrilla2(1, 8) = "FACTOR" & vbCrLf & "CORREC."
    formatogrilla2(1, 9) = "CORRECION" & vbCrLf & "MONETARIA"
    formatogrilla2(1, 10) = "VALOR" & vbCrLf & "ACTUAL."
    formatogrilla2(1, 11) = "VIDA" & vbCrLf & "UTIL"
    formatogrilla2(1, 12) = "VIDA" & vbCrLf & "USADA"
    formatogrilla2(1, 13) = "DEPRECIACION" & vbCrLf & "DEL EJERCICIO"
    
    
    Rem LARGO DE LOS DATOS
    formatogrilla2(2, 1) = "6"
    formatogrilla2(2, 2) = "15"
    formatogrilla2(2, 3) = "17"
    formatogrilla2(2, 4) = "7"
    formatogrilla2(2, 5) = "8"
    formatogrilla2(2, 6) = "10"
    formatogrilla2(2, 7) = "8"
    formatogrilla2(2, 8) = "8"
    formatogrilla2(2, 9) = "8"
    formatogrilla2(2, 10) = "8"
    formatogrilla2(2, 11) = "8"
    formatogrilla2(2, 12) = "8"
    formatogrilla2(2, 13) = "15"
    
    Rem TIPO DE DATOS C=CEROS,S=STRING,N=NUMERICO,D=DATE,
 
    
    formatogrilla2(3, 1) = "N"
    formatogrilla2(3, 2) = "S"
    formatogrilla2(3, 3) = "S"
    formatogrilla2(3, 4) = "D"
    formatogrilla2(3, 5) = "N"
    formatogrilla2(3, 6) = "N"
    formatogrilla2(3, 7) = "N"
    formatogrilla2(3, 8) = "N"
    formatogrilla2(3, 9) = "N"
    formatogrilla2(3, 10) = "N"
    formatogrilla2(3, 11) = "N"
    formatogrilla2(3, 12) = "N"
    formatogrilla2(3, 13) = "N"
    
    Rem FORMATO GRILLA
    formatogrilla2(4, 5) = "$ ###,###,##0"
    formatogrilla2(4, 6) = "$ ###,###,##0"
    formatogrilla2(4, 7) = "$ ###,###,##0"
    formatogrilla2(4, 8) = "% ###,###,##0.00"
    formatogrilla2(4, 9) = "$ ###,###,##0.00"
    formatogrilla2(4, 10) = "$ ###,###,##0"
    formatogrilla2(4, 13) = "$ ###,###,##0.00"
      
    
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
    formatogrilla2(5, 13) = "TRUE"
    
    
    Rem VALOR MAXIMO
    
    Grid1.Cols = 14
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
    
    For k = 1 To Grid1.Cols - 1
        Grid1.Cell(0, k).text = formatogrilla2(1, k)
        Grid1.Column(k).Width = Val(formatogrilla2(2, k)) * 8
        Grid1.Column(k).MaxLength = Val(formatogrilla2(2, k))
        Grid1.Column(k).FormatString = formatogrilla2(4, k)
        Grid1.Column(k).Locked = formatogrilla2(5, k)
        If formatogrilla2(3, k) = "N" Then Grid1.Column(k).Alignment = cellRightCenter
        If formatogrilla2(3, k) = "S" Then Grid1.Column(k).Alignment = cellLeftTop
        If formatogrilla2(3, k) = "D" Then Grid1.Column(k).CellType = cellCalendar
    Next k
  Grid1.DefaultFont.Size = 6.5
   Grid1.RowHeight(0) = 30
  Grid1.Column(0).Width = 0
  Grid1.Range(0, 0, 0, Grid1.Cols - 1).Alignment = cellCenterCenter
  Grid1.Range(0, 1, 0, Grid1.Cols - 1).WrapText = True
  
 
    End Sub
 
Sub verdetalle(loc, numero)
'Dim cSql As New rdoQuery
'Dim resultados As rdoResultset
'Dim tipo As String
'tipo = "DM"
'
'Set cSql.ActiveConnection = contadb
'
'cSql.sql = "select linea,codigo,descripcion,cantidad,uxc,unidades,precio,descuento,total "
'cSql.sql = cSql.sql & "from " & clientesistema & "gestion" & leerrubro(dato1.text) & ".l_movimientos_detalle_" & loc & " where tipo='" & tipo & "' and numero='" & numero & "' order by linea"
'cSql.Execute
'
'If cSql.RowsAffected > 0 Then
'    Grid1.Rows = cSql.RowsAffected + 1
'    Set resultados = cSql.OpenResultset
'
'    While Not resultados.EOF
'        Grid1.Cell(resultados(0), 1).text = resultados(1)
'        Grid1.Cell(resultados(0), 2).text = resultados(2)
'        Grid1.Cell(resultados(0), 3).text = resultados(3)
'        Grid1.Cell(resultados(0), 4).text = resultados(4)
'        Grid1.Cell(resultados(0), 5).text = resultados(5)
'        Grid1.Cell(resultados(0), 6).text = resultados(6)
'        Grid1.Cell(resultados(0), 7).text = resultados(7)
'        Grid1.Cell(resultados(0), 8).text = resultados(8)
'        resultados.MoveNext
'    Wend
'End If
'
'cSql.Close
'Set cSql = Nothing
'Set resultados = Nothing
 
End Sub
Function leerrubro(loc) As String
    Dim csql As New rdoQuery
    Dim resultado As rdoResultset
    
    Set csql.ActiveConnection = contadb
    csql.sql = "select rubro from " & clientesistema & "gestion.g_maestroempresas where "
    csql.sql = csql.sql & "codigo='" & loc & "' "
    csql.Execute
    
 If csql.RowsAffected > 0 Then
    Set resultado = csql.OpenResultset
    leerrubro = resultado(0)
 End If
 csql.Close
 Set csql = Nothing
 Set resultado = Nothing
 
End Function


Sub LEERactivofijos()
Dim csql As New rdoQuery
Dim resultados As rdoResultset
Dim tipo As String
Dim filtro As String
Dim FILTRO2 As String
Dim totales(3) As Double
Dim totales2(3) As Double
Dim cuentapublicidad As String

 
    Set csql.ActiveConnection = contadb
'    formatogrilla2(1, 1) = "CODIGO"
'    formatogrilla2(1, 2) = "NOMBRE"
'    formatogrilla2(1, 3) = "CRCC"
'    formatogrilla2(1, 4) = "FECHA P." & vbCrLf & "SERVICIO"
'    formatogrilla2(1, 5) = "VALOR" & vbCrLf & "LIBRO"
'    formatogrilla2(1, 6) = "DEPRECIAC." & vbCrLf & "ACUMULADA"
'    formatogrilla2(1, 7) = "VALOR" & vbCrLf & "NETO"
'    formatogrilla2(1, 8) = "FACTOR" & vbCrLf & "CORREC."
'    formatogrilla2(1, 9) = "CORRECION" & vbCrLf & "MONETARIA"
'    formatogrilla2(1, 10) = "VALOR" & vbCrLf & "ACTUAL."
'    formatogrilla2(1, 11) = "VIDA" & vbCrLf & "UTIL"
'    formatogrilla2(1, 12) = "VIDA" & vbCrLf & "USADA"
'    formatogrilla2(1, 13) = "DEPRECIACION" & vbCrLf & "DEL EJERCICIO"

    
    
    csql.sql = "select codigo,nombre,crcc,fechapuestaenmarcha,valorcompra,depreciacion,"
    csql.sql = csql.sql & "valorcompra-depreciacion,'',correcionmonetaria,'0',"
    csql.sql = csql.sql & "vidautil,vidausada,'0',familia"
    csql.sql = csql.sql & " from activo_fijo   "
    If Option2.Value = True Then
        csql.sql = csql.sql & "where familia='" & dato3.text & "' "
    End If
    csql.sql = csql.sql & "order by familia,nombre"
    csql.Execute
    Grid1.Rows = 1
    
    If csql.RowsAffected > 0 Then
        Set resultados = csql.OpenResultset
        filtro = resultados(13)
        FILTRO2 = filtro
        Grid1.AutoRedraw = False
        
        While Not resultados.EOF
            If filtro <> FILTRO2 Then
                Grid1.Rows = Grid1.Rows + 1
                Grid1.Column(1).Locked = False
                Grid1.Column(2).Locked = False
                Grid1.Column(3).Locked = False
                Grid1.Column(4).Locked = False
                Grid1.Range(Grid1.Rows - 1, 1, Grid1.Rows - 1, 4).FontBold = True
                
                Grid1.Cell(Grid1.Rows - 1, 4).Font.Bold = True
                Grid1.Cell(Grid1.Rows - 1, 5).Font.Bold = True
                Grid1.Cell(Grid1.Rows - 1, 7).Font.Bold = True
                Grid1.Cell(Grid1.Rows - 1, 10).Font.Bold = True
                Grid1.Range(Grid1.Rows - 1, 1, Grid1.Rows - 1, 3).Merge
                Grid1.Range(Grid1.Rows - 1, 1, Grid1.Rows - 1, Grid1.Cols - 1).Borders(cellEdgeTop) = cellThin
                Grid1.Range(Grid1.Rows - 1, 1, Grid1.Rows - 1, Grid1.Cols - 1).Borders(cellEdgeBottom) = cellThin
                Grid1.Range(Grid1.Rows - 1, 1, Grid1.Rows - 1, Grid1.Cols - 1).Borders(cellEdgeLeft) = cellThin
                Grid1.Range(Grid1.Rows - 1, 1, Grid1.Rows - 1, Grid1.Cols - 1).Borders(cellEdgeRight) = cellThin
                Grid1.Cell(Grid1.Rows - 1, 1).text = "TOTAL FAMILIA " & filtro & "-" & LEERNOMBREFAMILIA(filtro)
                Grid1.Range(Grid1.Rows - 1, 1, Grid1.Rows - 1, 3).Alignment = cellLeftGeneral
                Grid1.Cell(Grid1.Rows - 1, 5).text = totales(0)
                Grid1.Cell(Grid1.Rows - 1, 7).text = totales(1)
                Grid1.Cell(Grid1.Rows - 1, 10).text = totales(2)
                Grid1.Cell(Grid1.Rows - 1, 4).text = totales(3)
                
                For k = 0 To 3
                    totales(k) = 0
                Next k
                filtro = FILTRO2
                
            End If
            
            Grid1.Rows = Grid1.Rows + 1
             Grid1.Cell(Grid1.Rows - 1, 0).text = "1"
            Grid1.Cell(Grid1.Rows - 1, 1).text = resultados(0)
            Grid1.Cell(Grid1.Rows - 1, 2).text = resultados(1)
'            Grid1.Cell(Grid1.Rows - 1, 3).text = resultados(2) & "-" & LEERNOMBREFAMILIA(resultados(2))
            Grid1.Cell(Grid1.Rows - 1, 3).text = leerNOMBREcrcc(resultados(2))
            Grid1.Cell(Grid1.Rows - 1, 4).text = Format(resultados(3), "dd-mm-yyyy")
            Grid1.Cell(Grid1.Rows - 1, 5).text = resultados(4)
'            Grid1.Cell(Grid1.Rows - 1, 6).text = resultados(5) & "-" & LEERNOMBREPROVEEDOR(resultados(5))
             Grid1.Cell(Grid1.Rows - 1, 6).text = resultados(5)
            Grid1.Cell(Grid1.Rows - 1, 7).text = resultados(6)
            If Format(resultados(3), "yyyy") < Format(fechasistema, "yyyy") Then
                Grid1.Cell(Grid1.Rows - 1, 8).text = leeripc("00", Format(fechasistema, "yyyy"))
            Else
                Grid1.Cell(Grid1.Rows - 1, 8).text = leeripc(Format(resultados(3), "mm"), Format(resultados(3), "yyyy"))
            End If
            
            Grid1.Cell(Grid1.Rows - 1, 9).text = (CDbl(Grid1.Cell(Grid1.Rows - 1, 7).text) * CDbl(Grid1.Cell(Grid1.Rows - 1, 8).text)) / 100
            Grid1.Cell(Grid1.Rows - 1, 10).text = CDbl(Grid1.Cell(Grid1.Rows - 1, 7).text) + CDbl(Grid1.Cell(Grid1.Rows - 1, 9).text)
            Grid1.Cell(Grid1.Rows - 1, 11).text = resultados(10)
            
            Grid1.Cell(Grid1.Rows - 1, 12).text = DateDiff("m", resultados(3), DateAdd("m", 1, fechasistema))
            If Grid1.Cell(Grid1.Rows - 1, 12).text > "12" Then Grid1.Cell(Grid1.Rows - 1, 12).text = "12"
        
            
            Grid1.Cell(Grid1.Rows - 1, 13).text = Round((CDbl(Grid1.Cell(Grid1.Rows - 1, 7).text) / CDbl(Grid1.Cell(Grid1.Rows - 1, 11).text)) * CDbl(Grid1.Cell(Grid1.Rows - 1, 12).text), 2)
'            Grid1.Cell(Grid1.Rows - 1, 13).text = resultados(12) & "-" & leerNOMBREcrcc(resultados(12))
            totales(0) = totales(0) + resultados(4)
            totales(1) = totales(1) + resultados(6)
            totales(2) = totales(2) + Grid1.Cell(Grid1.Rows - 1, 10).text
            totales(3) = totales(3) + 1
            
            totales2(0) = totales2(0) + resultados(4)
            totales2(1) = totales2(1) + resultados(6)
            totales2(2) = totales2(2) + Grid1.Cell(Grid1.Rows - 1, 10).text
            totales2(3) = totales2(3) + 1
            
            resultados.MoveNext
            If Not resultados.EOF Then
                FILTRO2 = resultados(13)
            End If
        Wend
                Grid1.Rows = Grid1.Rows + 1
                Grid1.Cell(Grid1.Rows - 1, 1).text = "TOTAL FAMILIA " & filtro & "-" & LEERNOMBREFAMILIA(filtro)
                Grid1.Column(1).Locked = False
                Grid1.Column(2).Locked = False
                Grid1.Column(3).Locked = False
                Grid1.Column(4).Locked = False
                Grid1.Range(Grid1.Rows - 1, 1, Grid1.Rows - 1, 3).Merge
                Grid1.Range(Grid1.Rows - 1, 1, Grid1.Rows - 1, 3).Alignment = cellLeftGeneral
                Grid1.Range(Grid1.Rows - 1, 1, Grid1.Rows - 1, 3).FontBold = True
                Grid1.Range(Grid1.Rows - 1, 1, Grid1.Rows - 1, Grid1.Cols - 1).Borders(cellEdgeTop) = cellThin
                Grid1.Range(Grid1.Rows - 1, 1, Grid1.Rows - 1, Grid1.Cols - 1).Borders(cellEdgeBottom) = cellThin
                Grid1.Range(Grid1.Rows - 1, 1, Grid1.Rows - 1, Grid1.Cols - 1).Borders(cellEdgeLeft) = cellThin
                Grid1.Range(Grid1.Rows - 1, 1, Grid1.Rows - 1, Grid1.Cols - 1).Borders(cellEdgeRight) = cellThin
                Grid1.Cell(Grid1.Rows - 1, 4).Font.Bold = True
                Grid1.Cell(Grid1.Rows - 1, 5).Font.Bold = True
                Grid1.Cell(Grid1.Rows - 1, 7).Font.Bold = True
                Grid1.Cell(Grid1.Rows - 1, 10).Font.Bold = True
                Grid1.Cell(Grid1.Rows - 1, 5).text = totales(0)
                Grid1.Cell(Grid1.Rows - 1, 7).text = totales(1)
                Grid1.Cell(Grid1.Rows - 1, 10).text = totales(2)
                Grid1.Cell(Grid1.Rows - 1, 4).text = totales(3)
                For k = 0 To 3
                    totales(k) = 0
                Next k
                
                Grid1.Rows = Grid1.Rows + 2
                Grid1.Cell(Grid1.Rows - 1, 1).text = "TOTAL GENERAL"
                Grid1.Column(1).Locked = False
                Grid1.Column(2).Locked = False
                Grid1.Column(3).Locked = False
                Grid1.Column(4).Locked = False
                Grid1.Range(Grid1.Rows - 1, 1, Grid1.Rows - 1, 3).Merge
                Grid1.Range(Grid1.Rows - 1, 1, Grid1.Rows - 1, 3).Alignment = cellLeftGeneral
                Grid1.Range(Grid1.Rows - 1, 1, Grid1.Rows - 1, 3).FontBold = True
                Grid1.Range(Grid1.Rows - 1, 1, Grid1.Rows - 1, Grid1.Cols - 1).Borders(cellEdgeTop) = cellThin
                Grid1.Range(Grid1.Rows - 1, 1, Grid1.Rows - 1, Grid1.Cols - 1).Borders(cellEdgeBottom) = cellThin
                Grid1.Range(Grid1.Rows - 1, 1, Grid1.Rows - 1, Grid1.Cols - 1).Borders(cellEdgeLeft) = cellThin
                Grid1.Range(Grid1.Rows - 1, 1, Grid1.Rows - 1, Grid1.Cols - 1).Borders(cellEdgeRight) = cellThin
                Grid1.Cell(Grid1.Rows - 1, 4).Font.Bold = True
                Grid1.Cell(Grid1.Rows - 1, 5).Font.Bold = True
                Grid1.Cell(Grid1.Rows - 1, 7).Font.Bold = True
                Grid1.Cell(Grid1.Rows - 1, 10).Font.Bold = True
                Grid1.Cell(Grid1.Rows - 1, 4).text = totales2(3)
                Grid1.Cell(Grid1.Rows - 1, 5).text = totales2(0)
                Grid1.Cell(Grid1.Rows - 1, 7).text = totales2(1)
                Grid1.Cell(Grid1.Rows - 1, 10).text = totales2(2)
                
                
                For k = 1 To 3
                    totales2(k) = 0
                Next k
                Grid1.AutoRedraw = True
                Grid1.Refresh
                Grid1.Column(1).Locked = True
                Grid1.Column(2).Locked = True
                Grid1.Column(3).Locked = True
                Grid1.Column(4).Locked = True
    End If
    
    csql.Close
    Set csql = Nothing
    Set resultados = Nothing
 
End Sub

 

Private Sub Grid1_DblClick()
    If Grid1.Rows > 1 Then
        If Grid1.Cell(Grid1.ActiveCell.row, 0).text = "1" Then
            maestro06.dato1.text = Grid1.Cell(Grid1.ActiveCell.row, 1).text
            Load maestro06
            maestro06.Show
            maestro06.cargadeafueraactivofijo
        End If
    End If
    
End Sub

Private Sub Option1_Click()
    dato3.text = ""
    lblnombreFAMILIA.Caption = ""
    frmrut.Enabled = False
    LEERactivofijos
End Sub

Private Sub Option2_Click()
    frmrut.Enabled = True
    dato3.SetFocus
End Sub
Private Sub botonmisaccesos_Click()
programafiltro = Me.Caption
misaccesos.Show
End Sub


Private Sub botonmisfavoritos_Click()
Call AgregaFavoritos(USUARIOSISTEMA, App.EXEName, Me.Name, Me.Caption)
End Sub
