VERSION 5.00
Object = "{4777436C-EB8C-4596-98A8-EBCF98683969}#1.0#0"; "Flexcell.ocx"
Object = "{ADBBAED6-B16F-40DC-80DF-B44910CBA76C}#1.0#0"; "Frame-Xp.ocx"
Begin VB.Form tmpvistagestion 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "LISTADO GESTION COBRANZA"
   ClientHeight    =   7995
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9180
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7995
   ScaleWidth      =   9180
   StartUpPosition =   2  'CenterScreen
   Begin XPFrame.FrameXp FrameXp1 
      Height          =   7935
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9255
      _ExtentX        =   16325
      _ExtentY        =   13996
      BackColor       =   16761024
      Caption         =   "Gestión Cobranza"
      CaptionEstilo3D =   1
      BackColor       =   16761024
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   1
      Begin VB.CommandButton CMDimprimir 
         BackColor       =   &H00FF8080&
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
         Left            =   1080
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   7320
         Width           =   2055
      End
      Begin FlexCell.Grid Grid2 
         Height          =   375
         Left            =   240
         TabIndex        =   9
         Top             =   5760
         Visible         =   0   'False
         Width           =   30
         _ExtentX        =   53
         _ExtentY        =   661
         Cols            =   5
         DefaultFontSize =   8.25
         Rows            =   30
      End
      Begin VB.TextBox pivote 
         Height          =   285
         Left            =   9120
         TabIndex        =   7
         Text            =   "Text1"
         Top             =   1920
         Visible         =   0   'False
         Width           =   150
      End
      Begin XPFrame.FrameXp FrameXp2 
         Height          =   6255
         Left            =   240
         TabIndex        =   4
         Top             =   960
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   11033
         BackColor       =   16744576
         Caption         =   "Eventos"
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
         Begin FlexCell.Grid Grid1 
            Height          =   5895
            Left            =   120
            TabIndex        =   5
            Top             =   240
            Width           =   8775
            _ExtentX        =   15478
            _ExtentY        =   10398
            Cols            =   5
            DefaultFontSize =   8.25
            Rows            =   30
         End
      End
      Begin VB.CommandButton cmdgrabar 
         BackColor       =   &H00FF8080&
         Caption         =   "RETORNO"
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
         Left            =   5520
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   7320
         Width           =   2055
      End
      Begin VB.Label lblnombreevento 
         BackColor       =   &H00FFC0C0&
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
         Left            =   2880
         TabIndex        =   8
         Top             =   2400
         Width           =   3735
      End
      Begin VB.Label lblrut 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
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
         ForeColor       =   &H0000FF00&
         Height          =   255
         Left            =   2400
         TabIndex        =   6
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label Label8 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "RUT"
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
         Left            =   240
         TabIndex        =   3
         Top             =   240
         Width           =   1935
      End
      Begin VB.Label lblnombre 
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "NOMBRE"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   600
         Width           =   8775
      End
   End
End
Attribute VB_Name = "tmpvistagestion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdgrabar_Click()
 Unload Me
End Sub

Private Sub CMDimprimir_Click()
If Grid1.Rows > 1 Then
Call Titulos("LISTADO DE GESTION COBRANZA CREDITOS MOROSOS")
Grid1.PageSetup.Orientation = cellPortrait
Grid1.PageSetup.HeaderMargin = 0.5
Grid1.PageSetup.PrintFixedRow = True

Grid1.PageSetup.TopMargin = 3
Grid1.PageSetup.LeftMargin = 0.5
Grid1.PageSetup.RightMargin = 0.5
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
      
    Grid1.PageSetup.PrintTitleRows = 0
    
    'Logo
'    Grid1.Images.Add App.path & "\Admin.gif", "Logo"
'    Set objReportTitle = New FlexCell.ReportTitle
'    objReportTitle.ImageKey = "Logo"
'    objReportTitle.Align = CellLeft
'    Grid1.ReportTitles.Add objReportTitle
    
    'ENCABEZADO DE PAGINA
    Grid1.PageSetup.Header = nombreempresa & vbCrLf & direccionempresa & vbCrLf & comunaempresa & vbCrLf & rutempresa
    Grid1.PageSetup.HeaderAlignment = cellLeft
    Grid1.PageSetup.HeaderFont.Name = "Verdana"
    Grid1.PageSetup.HeaderFont.Size = 8
    
    'TITULOS DEL REPORTE
    
    Set objReportTitle = New FlexCell.ReportTitle
    objReportTitle.text = titulo1
    objReportTitle.Font.Name = "Verdana" '"Times New Roman"
    objReportTitle.Font.Size = 10
    objReportTitle.Font.Bold = True
    objReportTitle.Align = cellCenter
    objReportTitle.PrintOnAllPages = True
    Grid1.ReportTitles.Add objReportTitle
    
    Set objReportTitle = New FlexCell.ReportTitle
    objReportTitle.text = "CLIENTE : " & lblnombre.Caption
    objReportTitle.Font.Name = "Verdana" '"Times New Roman"
    objReportTitle.Font.Size = 8
    objReportTitle.Font.Bold = True
    objReportTitle.Align = cellCenter
    objReportTitle.PrintOnAllPages = True
    Grid1.ReportTitles.Add objReportTitle
    
    Set objReportTitle = New FlexCell.ReportTitle
    Grid1.ReportTitles.Add objReportTitle
    
    
    'PIE DE PAGINA
    Grid1.PageSetup.Footer = "Pág &P de &N" & vbCrLf & "Fecha: &D" & vbCrLf & "Usuario: " & usuarioSistema
    Grid1.PageSetup.FooterAlignment = cellRight
    Grid1.PageSetup.FooterFont.Name = "Verdana"
    Grid1.PageSetup.FooterFont.Size = 7
    
End Sub


Private Sub Form_Load()
    Call CargaGrillaGRID1(10, 5)
    Call leer(lblrut.Caption)
 End Sub
 Private Sub CargaGrillaGRID1(ByVal row As Integer, ByVal col As Integer)
        Dim i As Integer
       Dim formatogrilla(20, 20)
       Rem DATOS DE LA COLUMNA
        formatogrilla(1, 1) = "FECHA"
        formatogrilla(1, 2) = "HORA"
        formatogrilla(1, 3) = "EVENTO"
        formatogrilla(1, 4) = "GLOSA"
        
        Rem LARGO DE LOS DATOS
        formatogrilla(2, 1) = "10"
        formatogrilla(2, 2) = ""
        formatogrilla(2, 3) = ""
        formatogrilla(2, 4) = ""
        
        Rem TIPO DE DATOS
        formatogrilla(3, 1) = "D"
        formatogrilla(3, 2) = "S"
        formatogrilla(3, 3) = "S"
        formatogrilla(3, 4) = "S"
 
        
        Rem FORMATO GRILLA
        ''''''''''''''''''''''''
        formatogrilla(4, 1) = ""
        formatogrilla(4, 2) = ""
        formatogrilla(4, 3) = ""
        formatogrilla(4, 4) = ""

        Rem LOCCKED
        formatogrilla(5, 1) = "TRUE"
        formatogrilla(5, 2) = "TRUE"
        formatogrilla(5, 3) = "TRUE"
        formatogrilla(5, 4) = "TRUE"

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
        formatogrilla(8, 1) = "7"
        formatogrilla(8, 2) = "6"
        formatogrilla(8, 3) = "11"
        formatogrilla(8, 4) = "39"
 

            
        Grid1.Cols = col
        Grid1.Rows = row
        Grid1.AllowUserResizing = True
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
        
        Grid1.Column(0).Width = 0
        For i = 1 To col - 1
            Grid1.Cell(0, i).text = formatogrilla(1, i)
            Grid1.Column(i).Width = Val(formatogrilla(8, i)) * (Grid1.Cell(0, i).Font.Size + 1.25)
            Grid1.Column(i).MaxLength = Val(formatogrilla(2, i))
            Grid1.Column(i).FormatString = formatogrilla(4, i)
            Grid1.Column(i).Locked = formatogrilla(5, i)
            If formatogrilla(3, i) = "N" Then
                Grid1.Column(i).Alignment = cellRightCenter
            End If
            If formatogrilla(3, i) = "S" Then
                Grid1.Column(i).Alignment = cellLeftCenter
            End If
            If formatogrilla(3, i) = "C" Then
                Grid1.Column(i).Alignment = cellCenterCenter
            End If
        Next i
        Grid1.Range(0, 0, 0, Grid1.Cols - 1).Alignment = cellCenterCenter
        Grid1.Enabled = True
    
    
    End Sub
Sub leer(rut)
Dim cSql As New rdoQuery
Dim resultados As rdoResultset
rut = Replace(rut, "-", "")
rut = Replace(rut, ".", "")
pivote.MaxLength = 10
pivote.text = rut
pivote.text = ceros(pivote)

rut = pivote.text

Set cSql.ActiveConnection = ventas
cSql.sql = "select fecha,hora,evento,glosa from "
cSql.sql = cSql.sql & "sv_cobranza_gestion "
cSql.sql = cSql.sql & "where rut='" & rut & "' "
cSql.Execute
If cSql.RowsAffected > 0 Then
Grid1.Rows = 1
    Set resultados = cSql.OpenResultset
    While Not resultados.EOF
        Grid1.Rows = Grid1.Rows + 1
        Grid1.Cell(Grid1.Rows - 1, 1).text = resultados(0)
        Grid1.Cell(Grid1.Rows - 1, 2).text = resultados(1)
        Grid1.Cell(Grid1.Rows - 1, 3).text = leernombreevento(resultados(2))
        Grid1.Cell(Grid1.Rows - 1, 4).text = resultados(3)
        resultados.MoveNext
    Wend
End If
End Sub
Public Sub cargavista()
leer (lblrut.Caption)
End Sub

