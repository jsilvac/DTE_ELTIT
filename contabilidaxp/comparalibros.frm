VERSION 5.00
Object = "{4777436C-EB8C-4596-98A8-EBCF98683969}#1.0#0"; "FlexCell.ocx"
Object = "{ADBBAED6-B16F-40DC-80DF-B44910CBA76C}#1.0#0"; "Frame-Xp.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form control09 
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
      TabIndex        =   9
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
         TabIndex        =   11
         Top             =   280
         Width           =   1335
      End
      Begin VB.CommandButton botonmisfavoritos 
         Caption         =   "Mis Favoritos"
         Height          =   255
         Left            =   1560
         TabIndex        =   10
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
      TabIndex        =   0
      Top             =   10275
      Width           =   15240
   End
   Begin XPFrame.FrameXp FrameXp3 
      Height          =   10170
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   15135
      _ExtentX        =   26696
      _ExtentY        =   17939
      BackColor       =   16761024
      Caption         =   "LISTADO"
      CaptionEstilo3D =   1
      BackColor       =   16761024
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
      Begin XPFrame.FrameXp FrameXp1 
         Height          =   615
         Left            =   2280
         TabIndex        =   19
         Top             =   0
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   1085
         BackColor       =   16761024
         Caption         =   "LIBROS"
         CaptionEstilo3D =   1
         BackColor       =   16761024
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
         Begin VB.OptionButton opt2 
            BackColor       =   &H00FFC0C0&
            Caption         =   "VENTAS"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1440
            TabIndex        =   21
            Top             =   240
            Width           =   1335
         End
         Begin VB.OptionButton opt1 
            BackColor       =   &H00FFC0C0&
            Caption         =   "COMPRAS"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   20
            Top             =   240
            Value           =   -1  'True
            Width           =   1335
         End
      End
      Begin MSComctlLib.ProgressBar barra 
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   9480
         Width           =   14775
         _ExtentX        =   26061
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.CommandButton Command3 
         Caption         =   "EXCEL"
         Height          =   255
         Left            =   3120
         TabIndex        =   13
         Top             =   9840
         Width           =   2535
      End
      Begin VB.CheckBox CHK1 
         BackColor       =   &H00FFC0C0&
         Caption         =   "INCONSISTENCIAS"
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
         Left            =   12840
         TabIndex        =   12
         Top             =   600
         Width           =   2055
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00C0FFC0&
         Caption         =   "GENERA INFORME"
         Height          =   255
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   360
         Width           =   1935
      End
      Begin VB.CommandButton Command1 
         Caption         =   "IMPRIMIR"
         Height          =   255
         Left            =   7800
         TabIndex        =   3
         Top             =   9840
         Width           =   2535
      End
      Begin FlexCell.Grid Grid2 
         Height          =   8400
         Left            =   240
         TabIndex        =   2
         Top             =   1080
         Width           =   14775
         _ExtentX        =   26061
         _ExtentY        =   14817
         Cols            =   5
         DefaultFontSize =   8.25
         Rows            =   30
      End
      Begin XPFrame.FrameXp FrameXp6 
         Height          =   735
         Left            =   5640
         TabIndex        =   4
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
            TabIndex        =   5
            Top             =   240
            Width           =   3615
         End
      End
      Begin XPFrame.FrameXp FrameXp7 
         Height          =   735
         Left            =   9600
         TabIndex        =   6
         Top             =   240
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   1296
         BackColor       =   16744576
         Caption         =   " AÑO"
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
            TabIndex        =   7
            Top             =   240
            Width           =   2055
         End
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "EN ERP NO EN SII o MONTOS <>"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2760
         TabIndex        =   18
         Top             =   720
         Width           =   2655
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         BackColor       =   &H000000FF&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2400
         TabIndex        =   17
         Top             =   720
         Width           =   375
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "EN SII NO EN ERP"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   600
         TabIndex        =   16
         Top             =   720
         Width           =   1455
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H0080FFFF&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   240
         TabIndex        =   15
         Top             =   720
         Width           =   375
      End
   End
End
Attribute VB_Name = "control09"
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
'leerCONSUMOS
End Sub

Private Sub COMBOMES_Change()
'leerCONSUMOS
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
If opt1.Value = True Then
    LEERcompras
End If
If opt2.Value = True Then
    LEERVENTAS
End If
End Sub

Private Sub Command3_Click()
    If Grid2.Rows > 0 Then
        Call Grid2.ExportToExcel("", True, True)
    End If
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
    formatogrilla2(1, 2) = "RUT"
    formatogrilla2(1, 3) = "PROVEEDOR"
    formatogrilla2(1, 4) = "NUMERO"
    formatogrilla2(1, 5) = "FECHA"
    formatogrilla2(1, 6) = "IVA ERP"
    formatogrilla2(1, 7) = "TOTAL ERP"
    formatogrilla2(1, 8) = "IVA SII"
    formatogrilla2(1, 9) = "TOTAL SII"
    
    
    Rem LARGO DE LOS DATOS
    formatogrilla2(2, 1) = "5"
    formatogrilla2(2, 2) = "8"
    formatogrilla2(2, 3) = "25"
    formatogrilla2(2, 4) = "10"
    formatogrilla2(2, 5) = "8"
    formatogrilla2(2, 6) = "10"
    formatogrilla2(2, 7) = "10"
    formatogrilla2(2, 8) = "10"
    formatogrilla2(2, 9) = "10"
   
    
    Rem TIPO DE DATOS C=CEROS,S=STRING,N=NUMERICO,D=DATE,
    formatogrilla2(3, 1) = "N"
    formatogrilla2(3, 2) = "N"
    formatogrilla2(3, 3) = "S"
    formatogrilla2(3, 4) = "N"
    formatogrilla2(3, 5) = "D"
    formatogrilla2(3, 6) = "N"
    formatogrilla2(3, 7) = "N"
    formatogrilla2(3, 8) = "N"
    formatogrilla2(3, 9) = "N"
   
    
    Rem FORMATO GRILLA
    formatogrilla2(4, 6) = " ###,###,###,##0"
    formatogrilla2(4, 7) = " ###,###,###,##0"
    formatogrilla2(4, 8) = " ###,###,###,##0"
    formatogrilla2(4, 9) = " ###,###,###,##0"
    
    
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
   
    
    
    Rem VALOR MAXIMO
    
    Grid2.Cols = 10
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
    
    
    
    
    End Sub


 Public Sub LEERcompras()
 Dim csql As New rdoQuery
 Dim resultados As rdoResultset
 
 
 CARGAGRILLA2
 Set csql.ActiveConnection = contadb
 empresa_fae = leerdatos(conta, "maestroempresas", "empresafae", "codigoempresa='" + empresaactiva + "' ")
 csql.sql = "SELECT fc.tipo,fc.rut,cc.nombre,fc.numero,fc.fecha,fc.iva,fc.total,ifnull(f.iva,0) as iva,ifnull(f.total,0) as total,f.tipo FROM "
 csql.sql = csql.sql & "facturasdecompras AS fc LEFT JOIN cuentascorrientes AS cc ON  fc.rut=cc.rut LEFT JOIN "
 csql.sql = csql.sql & cliente_sql & "fae" & empresa_fae & ".sv_dte_libros_sii_compras AS f "
 csql.sql = csql.sql & "ON f.rut=fc.rut AND f.numero=fc.numero AND f.fecha=fc.fecha "
 csql.sql = csql.sql & "WHERE fc.tipo<>'' AND cc.año='" & COMBOAÑO.text & "' AND cc.tipo='23100026' AND "
 csql.sql = csql.sql & "fc.añocontable='" & COMBOAÑO.text & "' AND fc.mescontable='" & Format(COMBOMES.ListIndex + 1, "00") & "' "
 If chk1.Value = 1 Then
    csql.sql = csql.sql & " having fc.total<>total or fc.iva<>iva "
 End If
 csql.sql = csql.sql & " ORDER BY fc.fecha "
 
 csql.Execute
 
 
 Grid2.Rows = 1
 Grid2.AutoRedraw = False
 
 If csql.RowsAffected > 0 Then
    barra.Max = csql.RowsAffected + 1
    Set resultados = csql.OpenResultset
    While resultados.EOF = False
    Grid2.Rows = Grid2.Rows + 1
    barra.Value = Grid2.Rows
    
                If resultados(0) = "1" Then Grid2.Cell(Grid2.Rows - 1, 1).text = "FA"
                If resultados(0) = "2" Then Grid2.Cell(Grid2.Rows - 1, 1).text = "ND"
                If resultados(0) = "3" Then Grid2.Cell(Grid2.Rows - 1, 1).text = "NC"
                If resultados(0) = "4" Then Grid2.Cell(Grid2.Rows - 1, 1).text = "FAE"
                If resultados(0) = "5" Then Grid2.Cell(Grid2.Rows - 1, 1).text = "NDE"
                If resultados(0) = "6" Then Grid2.Cell(Grid2.Rows - 1, 1).text = "NCE"
                If resultados(0) = "7" Then Grid2.Cell(Grid2.Rows - 1, 1).text = "FC"
                If resultados(0) = "8" Then Grid2.Cell(Grid2.Rows - 1, 1).text = "IM"
                If resultados(0) = "9" Then Grid2.Cell(Grid2.Rows - 1, 1).text = "FE"
                If resultados(0) = "0" Then Grid2.Cell(Grid2.Rows - 1, 1).text = "FEE"
                If resultados(0) = "L" Then Grid2.Cell(Grid2.Rows - 1, 1).text = "LFE"
    
 
    Grid2.Cell(Grid2.Rows - 1, 2).text = resultados(1)
    Grid2.Cell(Grid2.Rows - 1, 3).text = resultados(2)
    Grid2.Cell(Grid2.Rows - 1, 4).text = resultados(3)
    Grid2.Cell(Grid2.Rows - 1, 5).text = resultados(4)
    Grid2.Cell(Grid2.Rows - 1, 6).text = resultados(5)
    Grid2.Cell(Grid2.Rows - 1, 7).text = resultados(6)
    Grid2.Cell(Grid2.Rows - 1, 8).text = resultados(7)
    Grid2.Cell(Grid2.Rows - 1, 9).text = resultados(8)
    
   If resultados(6) <> resultados(8) Or resultados(5) <> resultados(7) Then
     Grid2.Range(Grid2.Rows - 1, 1, Grid2.Rows - 1, Grid2.Cols - 1).BackColor = vbRed
   End If

    resultados.MoveNext



    Wend


  End If
  Call buscanoencontradoscompras(COMBOAÑO.text, Format(COMBOMES.ListIndex + 1, "00"), empresa_fae)
  
  Grid2.AutoRedraw = True
  Grid2.Refresh
  
 csql.Close
 Set csql = Nothing
 Set resultados = Nothing

 End Sub
 
Public Sub LEERVENTAS()
 Dim csql As New rdoQuery
 Dim resultados As rdoResultset
 
 
 CARGAGRILLA2
 Set csql.ActiveConnection = contadb
 empresa_fae = leerdatos(conta, "maestroempresas", "empresafae", "codigoempresa='" + empresaactiva + "' ")
 csql.sql = "SELECT fc.tipo,fc.rut,cc.nombre,fc.numero,fc.fecha,fc.iva,fc.total,ifnull(f.iva,0) as iva,ifnull(f.total,0) as total,f.tipo FROM "
 csql.sql = csql.sql & "facturasdeventas AS fc LEFT JOIN cuentascorrientes AS cc ON  fc.rut=cc.rut LEFT JOIN "
 csql.sql = csql.sql & cliente_sql & "fae" & empresa_fae & ".sv_dte_libros_sii_ventas AS f "
 csql.sql = csql.sql & "ON f.rut=fc.rut AND f.numero=fc.numero AND f.fecha=fc.fecha "
 csql.sql = csql.sql & "WHERE fc.tipo<>'' AND cc.año='" & COMBOAÑO.text & "' AND cc.tipo='11200027' AND "
 csql.sql = csql.sql & "fc.fecha LIKE '" & COMBOAÑO.text & "-" & Format(COMBOMES.ListIndex + 1, "00") & "%' "
 If chk1.Value = 1 Then
'    csql.sql = csql.sql & " having fc.total=total or fc.iva<>iva "
    csql.sql = csql.sql & " and f.total=0 "
 End If
 csql.sql = csql.sql & " ORDER BY fc.fecha "
 
 csql.Execute
 
 
 Grid2.Rows = 1
 Grid2.AutoRedraw = False
 
 If csql.RowsAffected > 0 Then
    barra.Max = csql.RowsAffected + 1
    Set resultados = csql.OpenResultset
    While resultados.EOF = False
    Grid2.Rows = Grid2.Rows + 1
    barra.Value = Grid2.Rows
    
                If resultados(0) = "1" Then Grid2.Cell(Grid2.Rows - 1, 1).text = "FA"
                If resultados(0) = "2" Then Grid2.Cell(Grid2.Rows - 1, 1).text = "ND"
                If resultados(0) = "3" Then Grid2.Cell(Grid2.Rows - 1, 1).text = "NC"
                If resultados(0) = "4" Then Grid2.Cell(Grid2.Rows - 1, 1).text = "NC"
                If resultados(0) = "5" Then Grid2.Cell(Grid2.Rows - 1, 1).text = "FEX"
                If resultados(0) = "6" Then Grid2.Cell(Grid2.Rows - 1, 1).text = "FAE"
                If resultados(0) = "7" Then Grid2.Cell(Grid2.Rows - 1, 1).text = "NDE"
                If resultados(0) = "8" Then Grid2.Cell(Grid2.Rows - 1, 1).text = "NCE"
                If resultados(0) = "9" Then Grid2.Cell(Grid2.Rows - 1, 1).text = "FE"
                If resultados(0) = "0" Then Grid2.Cell(Grid2.Rows - 1, 1).text = "FEE"
                If resultados(0) = "L" Then Grid2.Cell(Grid2.Rows - 1, 1).text = "LFE"
    
 
    Grid2.Cell(Grid2.Rows - 1, 2).text = resultados(1)
    Grid2.Cell(Grid2.Rows - 1, 3).text = resultados(2)
    Grid2.Cell(Grid2.Rows - 1, 4).text = resultados(3)
    Grid2.Cell(Grid2.Rows - 1, 5).text = resultados(4)
    Grid2.Cell(Grid2.Rows - 1, 6).text = resultados(5)
    Grid2.Cell(Grid2.Rows - 1, 7).text = resultados(6)
    Grid2.Cell(Grid2.Rows - 1, 8).text = resultados(7)
    Grid2.Cell(Grid2.Rows - 1, 9).text = resultados(8)
    
   If resultados(8) = 0 Then
     Grid2.Range(Grid2.Rows - 1, 1, Grid2.Rows - 1, Grid2.Cols - 1).BackColor = vbRed
   End If

    resultados.MoveNext



    Wend


  End If
  Call buscanoencontradosventas(COMBOAÑO.text, Format(COMBOMES.ListIndex + 1, "00"), empresa_fae)
  
  Grid2.AutoRedraw = True
  Grid2.Refresh
  
 csql.Close
 Set csql = Nothing
 Set resultados = Nothing

 End Sub
 
 Sub buscanoencontradoscompras(año, MES, loc)
    Dim csql As New rdoQuery
    Dim resultados As rdoResultset
    Set csql.ActiveConnection = contadb
    
     
    csql.sql = " SELECT tipo,rut,numero,fecha,iva,total FROM " & cliente_sql & "fae" & loc & ".sv_dte_libros_sii_compras "
    csql.sql = csql.sql & "WHERE  mescontable='" & MES & "' AND añocontable='" & año & "' "
    csql.sql = csql.sql & "AND numero NOT IN (SELECT numero FROM  facturasdecompras AS fc  "
    csql.sql = csql.sql & "WHERE fc.añocontable='" & año & "' AND fc.mescontable='" & MES & "')"
    csql.Execute
'    Grid2.Rows = Grid2.Rows + 1
    If csql.RowsAffected > 0 Then
        Grid2.AutoRedraw = False
        Set resultados = csql.OpenResultset
        While Not resultados.EOF
            Grid2.Rows = Grid2.Rows + 1
            
                If resultados(0) = "30" Then Grid2.Cell(Grid2.Rows - 1, 1).text = "FA"
                If resultados(0) = "55" Then Grid2.Cell(Grid2.Rows - 1, 1).text = "ND"
                If resultados(0) = "60" Then Grid2.Cell(Grid2.Rows - 1, 1).text = "NC"
                If resultados(0) = "33" Then Grid2.Cell(Grid2.Rows - 1, 1).text = "FAE"
                If resultados(0) = "56" Then Grid2.Cell(Grid2.Rows - 1, 1).text = "NDE"
                If resultados(0) = "61" Then Grid2.Cell(Grid2.Rows - 1, 1).text = "NCE"
                If resultados(0) = "46" Then Grid2.Cell(Grid2.Rows - 1, 1).text = "FC"
                If resultados(0) = "914" Then Grid2.Cell(Grid2.Rows - 1, 1).text = "IM"
                If resultados(0) = "32" Then Grid2.Cell(Grid2.Rows - 1, 1).text = "FE"
                If resultados(0) = "34" Then Grid2.Cell(Grid2.Rows - 1, 1).text = "FEE"
                If resultados(0) = "43" Then Grid2.Cell(Grid2.Rows - 1, 1).text = "LFE"
                
                
         
            Grid2.Cell(Grid2.Rows - 1, 2).text = resultados(1)
            Grid2.Cell(Grid2.Rows - 1, 3).text = LEERNOMBREPROVEEDOR(resultados(1))
            Grid2.Cell(Grid2.Rows - 1, 4).text = resultados(2)
            Grid2.Cell(Grid2.Rows - 1, 5).text = resultados(3)
            Grid2.Cell(Grid2.Rows - 1, 8).text = resultados(4)
            Grid2.Cell(Grid2.Rows - 1, 9).text = resultados(5)
            Grid2.Cell(Grid2.Rows - 1, 6).text = "0"
            Grid2.Cell(Grid2.Rows - 1, 7).text = "0"
            Grid2.Range(Grid2.Rows - 1, 1, Grid2.Rows - 1, Grid2.Cols - 1).BackColor = vbYellow
            resultados.MoveNext
        Wend
        Grid2.AutoRedraw = True
        Grid2.Refresh
        
    End If
    
    
 End Sub
 
 Sub buscanoencontradosventas(año, MES, loc)
    Dim csql As New rdoQuery
    Dim resultados As rdoResultset
    Set csql.ActiveConnection = contadb
    
     
    csql.sql = " SELECT tipo,rut,numero,fecha,iva,total FROM " & cliente_sql & "fae" & loc & ".sv_dte_libros_sii_ventas "
    csql.sql = csql.sql & "WHERE  mescontable='" & MES & "' AND añocontable='" & año & "' "
    csql.sql = csql.sql & "AND numero NOT IN (SELECT numero FROM  facturasdeventas AS fc  "
    csql.sql = csql.sql & "WHERE fc.fecha like '" & año & "-" & MES & "%') "
    csql.Execute
'    Grid2.Rows = Grid2.Rows + 1
    If csql.RowsAffected > 0 Then
        Grid2.AutoRedraw = False
        Set resultados = csql.OpenResultset
        While Not resultados.EOF
            Grid2.Rows = Grid2.Rows + 1
            
                If resultados(0) = "30" Then Grid2.Cell(Grid2.Rows - 1, 1).text = "FA"
                If resultados(0) = "55" Then Grid2.Cell(Grid2.Rows - 1, 1).text = "ND"
                If resultados(0) = "60" Then Grid2.Cell(Grid2.Rows - 1, 1).text = "NC"
                If resultados(0) = "33" Then Grid2.Cell(Grid2.Rows - 1, 1).text = "FAE"
                If resultados(0) = "56" Then Grid2.Cell(Grid2.Rows - 1, 1).text = "NDE"
                If resultados(0) = "61" Then Grid2.Cell(Grid2.Rows - 1, 1).text = "NCE"
                If resultados(0) = "46" Then Grid2.Cell(Grid2.Rows - 1, 1).text = "FC"
                If resultados(0) = "914" Then Grid2.Cell(Grid2.Rows - 1, 1).text = "IM"
                If resultados(0) = "32" Then Grid2.Cell(Grid2.Rows - 1, 1).text = "FE"
                If resultados(0) = "34" Then Grid2.Cell(Grid2.Rows - 1, 1).text = "FEE"
                If resultados(0) = "43" Then Grid2.Cell(Grid2.Rows - 1, 1).text = "LFE"
                
                
         
            Grid2.Cell(Grid2.Rows - 1, 2).text = resultados(1)
            Grid2.Cell(Grid2.Rows - 1, 3).text = LEERNOMBREPROVEEDOR(resultados(1))
            Grid2.Cell(Grid2.Rows - 1, 4).text = resultados(2)
            Grid2.Cell(Grid2.Rows - 1, 5).text = resultados(3)
            Grid2.Cell(Grid2.Rows - 1, 8).text = resultados(4)
            Grid2.Cell(Grid2.Rows - 1, 9).text = resultados(5)
            Grid2.Cell(Grid2.Rows - 1, 6).text = "0"
            Grid2.Cell(Grid2.Rows - 1, 7).text = "0"
            Grid2.Range(Grid2.Rows - 1, 1, Grid2.Rows - 1, Grid2.Cols - 1).BackColor = vbYellow
            resultados.MoveNext
        Wend
        Grid2.AutoRedraw = True
        Grid2.Refresh
        
    End If
    
    
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




'Private Sub dato1_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 27 Then Unload Me
'    snum = 0: KeyAscii = esNumero(KeyAscii)
'    If KeyAscii = 13 Then
'
'    Call ceros(dato1)
'    If leetipoconsumo(dato1.text) <> "" Then
'    LBLTIPO.Caption = leetipoconsumo(dato1.text)
'    leerCONSUMOS
'
'    Else
'    dato1.SetFocus
'    End If
'    End If
'
'
'
'
'End Sub
Private Sub botonmisaccesos_Click()
programafiltro = Me.Caption
misaccesos.Show
End Sub


Private Sub botonmisfavoritos_Click()
Call AgregaFavoritos(USUARIOSISTEMA, App.EXEName, Me.Name, Me.Caption)
End Sub

