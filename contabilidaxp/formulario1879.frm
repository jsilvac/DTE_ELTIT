VERSION 5.00
Object = "{4777436C-EB8C-4596-98A8-EBCF98683969}#1.0#0"; "FlexCell.ocx"
Object = "{ADBBAED6-B16F-40DC-80DF-B44910CBA76C}#1.0#0"; "Frame-Xp.ocx"
Begin VB.Form form1879 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "LISTA CERTIFICADOS DE HONORARIO"
   ClientHeight    =   8910
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   14880
   DrawStyle       =   5  'Transparent
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   Moveable        =   0   'False
   ScaleHeight     =   594
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   992
   Begin XPFrame.FrameXp FrameQuickMenu 
      Height          =   615
      Left            =   11520
      TabIndex        =   16
      Top             =   120
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
         TabIndex        =   18
         Top             =   280
         Width           =   1455
      End
      Begin VB.CommandButton botonmisfavoritos 
         Caption         =   "Mis Favoritos"
         Height          =   255
         Left            =   1800
         TabIndex        =   17
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
      Left            =   -90
      ScaleHeight     =   135
      ScaleWidth      =   135
      TabIndex        =   1
      Top             =   6120
      Width           =   135
   End
   Begin XPFrame.FrameXp FrameXp1 
      Height          =   8925
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   14865
      _ExtentX        =   26220
      _ExtentY        =   15743
      BackColor       =   16744576
      Caption         =   "INFORME CERTIFICADOS DE HONORARIOS"
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
      Begin VB.CommandButton Command3 
         BackColor       =   &H00FFC0C0&
         Caption         =   "GENERA SII"
         Height          =   330
         Left            =   3420
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   8370
         Width           =   1365
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00FF8080&
         Caption         =   "Vista Previa"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   315
         TabIndex        =   14
         Top             =   8370
         Width           =   1365
      End
      Begin VB.TextBox FIRMA 
         Height          =   330
         Left            =   5355
         TabIndex        =   10
         Top             =   8460
         Width           =   5460
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00FF8080&
         Caption         =   "TODOS"
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
         Left            =   12015
         TabIndex        =   9
         Top             =   8235
         Value           =   1  'Checked
         Width           =   2085
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FFC0C0&
         Caption         =   "IMPRIMIR"
         Height          =   330
         Left            =   1890
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   8370
         Width           =   1365
      End
      Begin XPFrame.FrameXp FrameXp3 
         Height          =   1050
         Left            =   135
         TabIndex        =   4
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
         Begin VB.CommandButton Command4 
            Caption         =   "IMPRIME PLANILLA"
            Height          =   495
            Left            =   5520
            TabIndex        =   19
            Top             =   360
            Width           =   3375
         End
         Begin VB.CommandButton Command2 
            Caption         =   "LISTAR"
            Height          =   285
            Left            =   12915
            TabIndex        =   6
            Top             =   630
            Width           =   1455
         End
         Begin XPFrame.FrameXp FrameXp7 
            Height          =   675
            Left            =   135
            TabIndex        =   7
            Top             =   315
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
               Style           =   2  'Dropdown List
               TabIndex        =   8
               Top             =   270
               Width           =   2865
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
         Begin FlexCell.Grid GRID1 
            Height          =   6330
            Left            =   90
            TabIndex        =   12
            Top             =   225
            Width           =   14550
            _ExtentX        =   25665
            _ExtentY        =   11165
            Cols            =   5
            DefaultFontName =   "Arial"
            DefaultFontSize =   8.25
            FixedRowColStyle=   0
            Rows            =   30
         End
      End
      Begin FlexCell.Grid Grid2 
         Height          =   5565
         Left            =   10395
         TabIndex        =   13
         Top             =   3330
         Visible         =   0   'False
         Width           =   4785
         _ExtentX        =   8440
         _ExtentY        =   9816
         Cols            =   5
         DefaultFontName =   "Arial"
         DefaultFontSize =   8.25
         FixedRowColStyle=   0
         Rows            =   30
         SelectionMode   =   1
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00FF8080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "RUT Y NOMBRE REPRESENTANTE LEGAL"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   5355
         TabIndex        =   11
         Top             =   8190
         Width           =   5505
      End
   End
End
Attribute VB_Name = "form1879"
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
Private rea1 As Double
Private rea2 As Double
Private MESES(12) As String


Private Sub Check1_Click()
For k = 1 To Grid1.Rows - 2
If Check1.Value = "0" Then
Grid1.Cell(k, 8).text = "0"
Else
Grid1.Cell(k, 8).text = "1"
End If

Next k

End Sub

Private Sub Command1_Click()
Dim s As Integer

CARGAGRILLA2

For s = 1 To Grid1.Rows - 2
If Grid1.Cell(s, 8).text = "1" Then
Call leercertificado(Grid1.Cell(s, 1).text, Grid1.Cell(s, 2).text, Grid1.Cell(s, 7).text)
Call IMPRIMIR2(Grid1.Cell(s, 1).text, Grid1.Cell(s, 2).text, Grid1.Cell(s, 7).text)


End If

Next s


End Sub






Private Sub Command3_Click()
Dim D1 As String
Dim D2 As String
Dim D3 As String
Dim D4 As String
Dim D5 As String
Dim D8 As String
Dim i As Double

Close 10

Open "F1879_" + empresaactiva + ".TXT" For Output As #10
For k = 1 To Grid1.Rows - 2
D1 = CDbl(Mid(Grid1.Cell(k, 1).text, 2, 8)) & Mid(Grid1.Cell(k, 1).text, 10, 1)

D2 = Format(Grid1.Cell(k, 6).text, "000000000000")
D3 = "000000000000"
D4 = "000000000000"
D8 = ""
For i = 1 To 12
D8 = D8 + Grid1.Cell(k, 8 + i).text + ";"
Next i
D8 = D8 & 0 & ";"
D8 = D8 & 0 & ";"
D5 = Format(CDbl(Grid1.Cell(k, 7).text), "0000000")
Print #10, D1 + ";" + D2 + ";" + D3 + ";" + D4 + ";" + D8 + D5
Next k
Close #10
Shell "NOTEPAD " + "F1879_" + empresaactiva + ".TXT"




End Sub


Private Sub Command4_Click()
    Call cabezas4("INFORME CERTIFICADO 1879", "N", 0)
    
    Grid1.PrintPreview
    
    End Sub
 
Sub cabezas4(titulo, tipo, FOLIO)
Dim objReportTitle As FlexCell.ReportTitle
Grid1.ReportTitles.Clear


Set objReportTitle = New FlexCell.ReportTitle
    objReportTitle.text = titulo
    objReportTitle.Font.Name = "arial"
    objReportTitle.Font.Size = 12
    objReportTitle.PrintOnAllPages = True
    Grid1.ReportTitles.Add objReportTitle

Set objReportTitle = New FlexCell.ReportTitle

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
        .HeaderMargin = 4
        .TopMargin = 2
        .BottomMargin = 1
        .LeftMargin = 1
        .RightMargin = 1
        .Orientation = cellLandscape
        .PrintFixedRow = True
        
        
        
        
End With

End Sub

Private Sub COMMAND2_Click()
leer
End Sub

Private Sub Form_Load()
CENTRAR Me


    
    Call Conectar_BD

    sc = 0
CARGAGRILLA
Call Conectarventas(Servidor, clientesistema + "ventas00", Usuario, password)
Call Conectargestion(Servidor, clientesistema + "gestion", Usuario, password)
Call Conectargestionrubro(Servidor, clientesistema + "gestion00", Usuario, password)

For k = 2000 To Val(Format(fechasistema, "yyyy"))
COMBOAÑO.AddItem k
Next k
COMBOAÑO.ListIndex = k - 2001


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
Sub IMPRIMIR2(rut, NOMBRE, numero)
Dim titulo As String
Call cabezas3(rut, NOMBRE, numero)
Grid2.Range(0, 1, 0, Grid2.Cols - 1).Borders(cellEdgeBottom) = cellThick
Grid2.Range(0, 1, 0, Grid2.Cols - 1).Borders(cellEdgeLeft) = cellThick
Grid2.Range(0, 1, 0, Grid2.Cols - 1).Borders(cellEdgeTop) = cellThick
Grid2.Range(0, 1, 0, Grid2.Cols - 1).Borders(cellEdgeRight) = cellThick
Grid2.Range(0, 1, 0, Grid2.Cols - 1).Borders(cellInsideHorizontal) = cellThick
Grid2.Range(0, 1, 0, Grid2.Cols - 1).Borders(cellInsideVertical) = cellThick
Grid2.DefaultFont.Size = 8
Grid2.PageSetup.Orientation = cellPortrait

Grid2.PageSetup.PrintFixedRow = True
Grid2.PageSetup.BottomMargin = 2
Grid2.PageSetup.TopMargin = 1
Grid2.PageSetup.LeftMargin = 0.5
Grid2.PageSetup.RightMargin = 0.5
Grid2.PageSetup.BlackAndWhite = True
Grid2.PageSetup.PrintGridlines = False

Grid2.Range(1, 1, 13, Grid2.Cols - 1).Borders(cellEdgeBottom) = cellThin
Grid2.Range(1, 1, 13, Grid2.Cols - 1).Borders(cellEdgeLeft) = cellThin
Grid2.Range(1, 1, 13, Grid2.Cols - 1).Borders(cellEdgeTop) = cellThin
Grid2.Range(1, 1, 13, Grid2.Cols - 1).Borders(cellEdgeRight) = cellThin
Grid2.Range(1, 1, 13, Grid2.Cols - 1).Borders(cellInsideHorizontal) = cellThin
Grid2.Range(1, 1, 13, Grid2.Cols - 1).Borders(cellInsideVertical) = cellThin

If Check2.Value = "1" Then

Grid2.PrintPreview 100
Else
Grid2.DirectPrint

End If

   
End Sub


Sub grilla()
    
End Sub




Private Sub opciones_GotFocus()

MANUAL.SetFocus

End Sub
Sub CARGAGRILLA()
Rem DATOS DE LA COLUMNA
    Dim FORMATOGRILLA(10, 20)
    Grid1.DefaultFont.Size = 8
    FORMATOGRILLA(1, 1) = "RUT"
    FORMATOGRILLA(1, 2) = "NOMBRE"
    FORMATOGRILLA(1, 3) = "HONORARIOS"
    FORMATOGRILLA(1, 4) = "RETENCION"
    FORMATOGRILLA(1, 5) = "HONORARIOS"
    FORMATOGRILLA(1, 6) = "RETENCION"
    FORMATOGRILLA(1, 7) = "NUMERO"
    FORMATOGRILLA(1, 8) = "IMPRIMIR"
    FORMATOGRILLA(1, 9) = "Ene"
    FORMATOGRILLA(1, 10) = "Feb"
    FORMATOGRILLA(1, 11) = "Mar"
    FORMATOGRILLA(1, 12) = "Abr"
    FORMATOGRILLA(1, 13) = "May"
    FORMATOGRILLA(1, 14) = "Jun"
    FORMATOGRILLA(1, 15) = "Jul"
    FORMATOGRILLA(1, 16) = "Ago"
    FORMATOGRILLA(1, 17) = "Sep"
    FORMATOGRILLA(1, 18) = "Oct"
    FORMATOGRILLA(1, 19) = "Nov"
    FORMATOGRILLA(1, 20) = "Dic"
    
    Rem LARGO DE LOS DATOS
    FORMATOGRILLA(2, 1) = "8"
    FORMATOGRILLA(2, 2) = "30"
    FORMATOGRILLA(2, 3) = "9"
    FORMATOGRILLA(2, 4) = "9"
    FORMATOGRILLA(2, 5) = "9"
    FORMATOGRILLA(2, 6) = "9"
    FORMATOGRILLA(2, 7) = "8"
    FORMATOGRILLA(2, 8) = "6"
    FORMATOGRILLA(2, 9) = "3"
    FORMATOGRILLA(2, 10) = "3"
    FORMATOGRILLA(2, 11) = "3"
    FORMATOGRILLA(2, 12) = "3"
    FORMATOGRILLA(2, 13) = "3"
    FORMATOGRILLA(2, 14) = "3"
    FORMATOGRILLA(2, 15) = "3"
    FORMATOGRILLA(2, 16) = "3"
    FORMATOGRILLA(2, 17) = "3"
    FORMATOGRILLA(2, 18) = "3"
    FORMATOGRILLA(2, 19) = "3"
    FORMATOGRILLA(2, 20) = "3"
    
    
    Rem TIPO DE DATOS C=CEROS,S=STRING,N=NUMERICO,D=DATE,
    FORMATOGRILLA(3, 1) = "S"
    FORMATOGRILLA(3, 2) = "S"
    FORMATOGRILLA(3, 3) = "S"
    FORMATOGRILLA(3, 4) = "N"
    FORMATOGRILLA(3, 5) = "N"
    FORMATOGRILLA(3, 6) = "N"
    FORMATOGRILLA(3, 7) = "N"
    FORMATOGRILLA(3, 8) = "N"
    FORMATOGRILLA(3, 9) = "S"
    FORMATOGRILLA(3, 10) = "S"
    FORMATOGRILLA(3, 11) = "S"
    FORMATOGRILLA(3, 12) = "S"
    FORMATOGRILLA(3, 13) = "S"
    FORMATOGRILLA(3, 14) = "S"
    FORMATOGRILLA(3, 15) = "S"
    FORMATOGRILLA(3, 16) = "S"
    FORMATOGRILLA(3, 17) = "S"
    FORMATOGRILLA(3, 18) = "S"
    FORMATOGRILLA(3, 19) = "S"
    FORMATOGRILLA(3, 20) = "S"
    
    Rem FORMATO GRILLA
    FORMATOGRILLA(4, 3) = "##,###,##0"
    FORMATOGRILLA(4, 4) = "##,###,##0"
    FORMATOGRILLA(4, 5) = "##,###,##0"
    FORMATOGRILLA(4, 6) = "##,###,##0"
    
    
    Rem LOCCKED
    For k = 1 To 20
    FORMATOGRILLA(5, k) = "TRUE"
    
    Next k
        
    
    Grid1.Cols = 21
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
        
        Grid1.Cell(0, k).text = FORMATOGRILLA(1, k)
        Grid1.Column(k).Width = Val(FORMATOGRILLA(2, k)) * Grid1.DefaultFont.Size
        Grid1.Column(k).MaxLength = Val(FORMATOGRILLA(2, k))
        Grid1.Column(k).FormatString = FORMATOGRILLA(4, k)
        Grid1.Column(k).Locked = FORMATOGRILLA(5, k)
        If FORMATOGRILLA(3, k) = "N" Then Grid1.Column(k).Alignment = cellRightCenter
        If FORMATOGRILLA(3, k) = "D" Then Grid1.Column(k).CellType = cellCalendar
        
    Next k
   Grid1.Column(8).CellType = cellCheckBox
   
   
    
    
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
    Dim LINEA As Double
    Dim total As Double
    Dim fec As Double
    Dim fec1 As Double
    Dim fechasum As String
    Dim total2 As Double
    Dim tila1 As Double
    Dim tila2 As Double
    Dim tila3 As Double
    Dim total3 As Double
    Dim total4 As Double
    
    LINEA = 0: fec = 0: fec1 = 0
    fecha1 = año + "-" + MES + "-" + "01"
    fecha2 = año + "-" + MES + "-" + "31"
    
        Set csql.ActiveConnection = contadb
        csql.sql = "select rut,sum(monto),sum(retencion) from boletasdehonorarios where retencion<>'0' and añocontable='" + COMBOAÑO.text + "' group by rut order by rut "
        
        csql.Execute
        Grid1.Rows = 1
        If csql.RowsAffected > 0 Then
        Set resultados = csql.OpenResultset
         While Not resultados.EOF
             LINEA = LINEA + 1
             Grid1.Rows = Grid1.Rows + 1
             Grid1.Cell(LINEA, 1).text = resultados(0)
             Grid1.Cell(LINEA, 2).text = leerdatos(contadb, "cuentascorrientes", "nombre", "rut='" + resultados(0) + "' and tipo='" + cuentahonorarios + "' ")
             COSTO1 = calculacertificado(resultados(0))
             Grid1.Cell(LINEA, 3).text = resultados(1)
             Grid1.Cell(LINEA, 4).text = resultados(2)
             Grid1.Cell(LINEA, 5).text = rea1
             Grid1.Cell(LINEA, 6).text = rea2
             Grid1.Cell(LINEA, 7).text = Format(LINEA, "0000000000")
             Grid1.Cell(LINEA, 8).text = "1"
             For k = 1 To 12
             Grid1.Cell(LINEA, 8 + k).text = MESES(k)
             
             Next k
             
             total = total + resultados(1)
             total2 = total2 + resultados(2)
             total3 = total3 + rea1
             total4 = total4 + rea2
             
             resultados.MoveNext
            
            Wend
             LINEA = LINEA + 1
             Grid1.Rows = Grid1.Rows + 1
             Grid1.Range(LINEA, 1, LINEA, 6).FontBold = True
             Grid1.Range(LINEA, 1, LINEA, 6).Borders(cellEdgeTop) = cellThin
             Grid1.Cell(LINEA, 3).text = total
             Grid1.Cell(LINEA, 4).text = total2
             Grid1.Cell(LINEA, 5).text = total3
             Grid1.Cell(LINEA, 6).text = total4
             
         
         resultados.Close
            Set resultados = Nothing

End If

      
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


Sub CARGAGRILLA2()
Rem DATOS DE LA COLUMNA
    Dim FORMATOGRILLA(10, 20)
    Grid2.DefaultFont.Size = 8
       
    FORMATOGRILLA(1, 1) = "PERIODO"
    FORMATOGRILLA(1, 2) = "HONORARIO " + vbCrLf + "BRUTO"
    FORMATOGRILLA(1, 3) = "RETENCION " + vbCrLf + " IMPUESTO"
    FORMATOGRILLA(1, 4) = "FACTOR " + vbCrLf + "ACTUALIZACION"
    FORMATOGRILLA(1, 5) = "HONORARIO " + vbCrLf + "BRUTO ACTUALIZADO"
    FORMATOGRILLA(1, 6) = "RETENCION " + vbCrLf + "IMPTO ACTUALIZADO"
    
    Rem LARGO DE LOS DATOS
    FORMATOGRILLA(2, 1) = "15"
    FORMATOGRILLA(2, 2) = "15"
    FORMATOGRILLA(2, 3) = "15"
    FORMATOGRILLA(2, 4) = "15"
    FORMATOGRILLA(2, 5) = "15"
    FORMATOGRILLA(2, 6) = "15"
    
    
    Rem TIPO DE DATOS C=CEROS,S=STRING,N=NUMERICO,D=DATE,
    FORMATOGRILLA(3, 1) = "S"
    FORMATOGRILLA(3, 2) = "N"
    FORMATOGRILLA(3, 3) = "N"
    FORMATOGRILLA(3, 4) = "N"
    FORMATOGRILLA(3, 5) = "N"
    FORMATOGRILLA(3, 6) = "N"
    
    Rem FORMATO GRILLA
    FORMATOGRILLA(4, 2) = "##,###,##0"
    FORMATOGRILLA(4, 3) = "##,###,##0"
    FORMATOGRILLA(4, 4) = "##,###,##0.000"
    FORMATOGRILLA(4, 5) = "##,###,##0"
    FORMATOGRILLA(4, 6) = "##,###,##0"
    
    
    Rem LOCCKED
    For k = 1 To 6
    FORMATOGRILLA(5, k) = "false"
    
    Next k
        
    
    Grid2.Cols = 7
    Grid2.Rows = 2
    
    Grid2.AllowUserResizing = False
    Grid2.DisplayFocusRect = False
    Grid2.ExtendLastCol = True
    Grid2.BoldFixedCell = False
    Grid2.DrawMode = cellOwnerDraw
    
    Grid2.Appearance = Flat
    Grid2.ScrollBarStyle = Flat
    Grid2.FixedRowColStyle = Flat
    
   
    
    
'   grid2.BackColorFixed = RGB(90, 158, 214)
'   grid2.BackColorFixedSel = RGB(110, 180, 230)
'   grid2.BackColorBkg = RGB(90, 158, 214)
'   grid2.BackColorScrollBar = RGB(231, 235, 247)
'   grid2.BackColor1 = RGB(231, 235, 247)
'   grid2.BackColor2 = RGB(239, 243, 255)
'   grid2.GridColor = RGB(148, 190, 231)
    Grid2.Column(0).Width = 0
    
    For k = 1 To Grid2.Cols - 1
        
        Grid2.Cell(0, k).text = FORMATOGRILLA(1, k)
        Grid2.Column(k).Width = Val(FORMATOGRILLA(2, k)) * Grid2.DefaultFont.Size
        Grid2.Column(k).MaxLength = Val(FORMATOGRILLA(2, k))
        Grid2.Column(k).FormatString = FORMATOGRILLA(4, k)
        Grid2.Column(k).Locked = FORMATOGRILLA(5, k)
        If FORMATOGRILLA(3, k) = "N" Then Grid2.Column(k).Alignment = cellRightCenter
        If FORMATOGRILLA(3, k) = "D" Then Grid2.Column(k).CellType = cellCalendar
        
    Next k
        Grid2.RowHeight(0) = 30
   
      Grid2.Range(0, 1, 0, Grid2.Cols - 1).WrapText = True
      Grid2.Cell(0, 1).Alignment = cellCenterCenter
      Grid2.Cell(0, 2).Alignment = cellCenterCenter
      Grid2.Cell(0, 3).Alignment = cellCenterCenter
      Grid2.Cell(0, 4).Alignment = cellCenterCenter
      Grid2.Cell(0, 5).Alignment = cellCenterCenter
      Grid2.Cell(0, 6).Alignment = cellCenterCenter
      
      
      
   
    
    
End Sub

Private Sub leercertificado(rut, NOMBRE, numero)

Dim resultados As rdoResultset
    Dim csql As New rdoQuery
    Dim j As Double
    
    Dim PASO As String
    Dim fecha1 As String
    Dim fecha2 As String
    Dim LINEA As Double
    Dim total As Double
    Dim fec As Double
    Dim fec1 As Double
    Dim fechasum As String
    Dim total2 As Double
    Dim total3 As Double
    Dim total4 As Double
    Dim tila3 As Double
    Dim ipc As Double
    Dim corre1 As Double
    Dim corre2 As Double
    Dim total5 As Double
    Dim total6 As Double
    
    LINEA = 0: fec = 0: fec1 = 0
    fecha1 = año + "-" + MES + "-" + "01"
    fecha2 = año + "-" + MES + "-" + "31"
        
        Set csql.ActiveConnection = contadb
        csql.sql = "select mescontable,sum(monto),sum(retencion) from boletasdehonorarios where retencion<>'0' and añocontable='" + COMBOAÑO.text + "' and rut='" + rut + "' group by mescontable order by mescontable "
        
        csql.Execute
        Grid2.Rows = 32
        For j = 1 To 12
        Grid2.Cell(j, 1).text = MonthName(j)
        Grid2.Cell(j, 2).text = "0"
        Grid2.Cell(j, 3).text = "0"
        Grid2.Cell(j, 4).text = 1 + (leeripc(Format(j, "00"), COMBOAÑO.text) / 100)
        Grid2.Cell(j, 5).text = "0"
        Grid2.Cell(j, 6).text = "0"
        
        Next j
        
        If csql.RowsAffected > 0 Then
        Set resultados = csql.OpenResultset
         While Not resultados.EOF
             LINEA = resultados(0)
             ipc = 1 + (leeripc(resultados(0), COMBOAÑO.text) / 100)
             corre1 = Round(resultados(1) * ipc, 0)
             corre2 = Round(resultados(2) * ipc, 0)
             Grid2.Cell(LINEA, 1).text = MonthName(resultados(0))
             Grid2.Cell(LINEA, 2).text = resultados(1)
             Grid2.Cell(LINEA, 3).text = resultados(2)
             Grid2.Cell(LINEA, 4).text = ipc
             Grid2.Cell(LINEA, 5).text = corre1
             Grid2.Cell(LINEA, 6).text = corre2
             total = total + resultados(1)
             total2 = total2 + resultados(2)
             total3 = total3 + corre1
             total4 = total4 + corre2
             resultados.MoveNext
            Wend
             LINEA = 13
             
             Grid2.Range(LINEA, 1, LINEA, 6).FontBold = True
             Grid2.Range(LINEA, 1, LINEA, 6).Borders(cellEdgeTop) = cellThin
             Grid2.Cell(LINEA, 1).text = "TOTALES"
             Grid2.Cell(LINEA, 2).text = total
             Grid2.Cell(LINEA, 3).text = total2
             Grid2.Cell(LINEA, 5).text = total3
             Grid2.Cell(LINEA, 6).text = total4
             
         
            resultados.Close
            Set resultados = Nothing
            
            Grid2.Range(15, 1, 15, 6).Merge
            Grid2.Range(15, 1, 15, 6).FontSize = 12
            
            
            Grid2.Range(16, 1, 16, 6).Merge
            Grid2.Range(16, 1, 16, 6).FontSize = 12
            
            Grid2.Range(17, 1, 17, 6).Merge
            Grid2.Range(17, 1, 17, 6).FontSize = 12
            
            Grid2.Range(18, 1, 18, 6).Merge
            Grid2.Range(18, 1, 18, 6).FontSize = 12
            
            Grid2.Cell(16, 1).text = "Se extiende el presente certificado en cumplimiento de lo dispuesto en la Resolucion Ex Nro 6509 del"
            Grid2.Cell(17, 1).text = "Servicio de Impuestos Internos , publicada en el Diario Oficial de fecha 20 de Diciembre de 1993,y "
            Grid2.Cell(18, 1).text = "sus modificaciones posteriores "
            
            Grid2.Range(30, 3, 30, 5).Merge
            Grid2.Range(30, 3, 30, 5).Borders(cellEdgeTop) = cellThin
            Grid2.Cell(30, 3).Alignment = cellCenterCenter
            
            
            
            Grid2.Cell(30, 3).text = FIRMA.text
            


End If

      
End Sub

Public Function calculacertificado(rut) As Double


    Dim resultados As rdoResultset
    Dim csql As New rdoQuery
    Dim j As Double
    Dim ipc As Double
    Dim corre1 As Double
    Dim corre2 As Double
    
        
        Set csql.ActiveConnection = contadb
        csql.sql = "select mescontable,sum(monto),sum(retencion) from boletasdehonorarios where retencion<>'0' and añocontable='" + COMBOAÑO.text + "' and rut='" + rut + "' group by mescontable order by mescontable "
        csql.Execute
        rea1 = 0
        rea2 = 0
       For k = 1 To 12
       MESES(k) = ""
       Next k
       
        If csql.RowsAffected > 0 Then
        Set resultados = csql.OpenResultset
         While Not resultados.EOF
             MESES(resultados(0)) = "X"
             ipc = 1 + (leeripc(resultados(0), COMBOAÑO.text) / 100)
             corre1 = Round(resultados(1) * ipc, 0)
             corre2 = Round(resultados(2) * ipc, 0)
             rea1 = rea1 + corre1
             rea2 = rea2 + corre2
             
             resultados.MoveNext
            Wend
             
         
            resultados.Close
            Set resultados = Nothing
            
        End If
calculacertificado = 0

      
End Function


Sub cabezas3(rut, NOMBRE, numero)
Dim objReportTitle As FlexCell.ReportTitle
Grid2.ReportTitles.Clear



    'Report Title 1
        For k = 1 To 4
        Set objReportTitle = New FlexCell.ReportTitle
        objReportTitle.text = DATOSEMPRESA(k)
        objReportTitle.Font.Name = "arial"
        objReportTitle.Font.Size = 8
        objReportTitle.Font.Italic = True
        objReportTitle.PrintOnAllPages = True
        'objReportTitle.Color = RGB(128, 0, 0)
        objReportTitle.Align = CellLeft
        Grid2.ReportTitles.Add objReportTitle
    Next k

Set objReportTitle = New FlexCell.ReportTitle
    objReportTitle.text = "CERTIFICADO SOBRE HONORARIOS NUMERO " + numero
    objReportTitle.Font.Name = "arial"
    objReportTitle.Font.Size = 12
    objReportTitle.PrintOnAllPages = True
    Grid2.ReportTitles.Add objReportTitle

Set objReportTitle = New FlexCell.ReportTitle
    objReportTitle.Font.Name = "arial"
    objReportTitle.Font.Size = 12
    objReportTitle.PrintOnAllPages = True
    Grid2.ReportTitles.Add objReportTitle


Set objReportTitle = New FlexCell.ReportTitle
    objReportTitle.text = DATOSEMPRESA(3) + "   fecha :" & fechasistema
    objReportTitle.Font.Name = "arial"
    objReportTitle.Font.Size = 10
    objReportTitle.PrintOnAllPages = True
    objReportTitle.Align = cellRight
    
    Grid2.ReportTitles.Add objReportTitle


Set objReportTitle = New FlexCell.ReportTitle
    objReportTitle.text = "La empresa Certifica Que Don :" + NOMBRE
    objReportTitle.Font.Name = "arial"
    objReportTitle.Font.Size = 12
    objReportTitle.PrintOnAllPages = True
    objReportTitle.Align = CellLeft
    
    Grid2.ReportTitles.Add objReportTitle


Set objReportTitle = New FlexCell.ReportTitle
    objReportTitle.text = "Rut :" + Mid(rut, 1, 9) + "-" + Mid(rut, 10, 1) + " durante el año " + COMBOAÑO.text + " se le han pagado las siguientes rentas por concepto de honorarios "
    objReportTitle.Font.Name = "arial"
    objReportTitle.Font.Size = 12
    objReportTitle.PrintOnAllPages = True
    objReportTitle.Align = CellLeft
    Grid2.ReportTitles.Add objReportTitle

Set objReportTitle = New FlexCell.ReportTitle
    objReportTitle.text = "sobre los cuales se le practicaron las retenciones de impuestos que se señalan "
    objReportTitle.Font.Name = "arial"
    objReportTitle.Font.Size = 12
    objReportTitle.PrintOnAllPages = True
    objReportTitle.Align = CellLeft
    Grid2.ReportTitles.Add objReportTitle

Set objReportTitle = New FlexCell.ReportTitle
    objReportTitle.text = ""
    objReportTitle.Font.Name = "arial"
    objReportTitle.Font.Size = 12
    objReportTitle.PrintOnAllPages = True
    objReportTitle.Align = CellLeft
    Grid2.ReportTitles.Add objReportTitle


With Grid2.PageSetup
        
        Rem If tipo = "N" Then .Header = "Pagina &P de &N Emitido: &D Usuario:" + USUARIOSISTEMA
        Rem If TIPO = "S" Then .Footer = "pagina &P"
        
        .HeaderAlignment = cellCenter
        
        .HeaderFont.Name = "Verdana"
        .HeaderFont.Size = 7
        .HeaderMargin = 2
        .TopMargin = 2
        .BottomMargin = 1
        
        
        
End With

End Sub

Private Sub Grid1_Click()
If Grid1.ActiveCell.col = 8 Then
        If Grid1.Cell(Grid1.ActiveCell.row, 8).text = "1" Then
            Grid1.Cell(Grid1.ActiveCell.row, 8).text = "0"
        Else
            Grid1.Cell(Grid1.ActiveCell.row, 8).text = "1"
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
