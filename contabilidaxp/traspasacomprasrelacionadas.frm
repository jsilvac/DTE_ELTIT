VERSION 5.00
Object = "{4777436C-EB8C-4596-98A8-EBCF98683969}#1.0#0"; "FlexCell.ocx"
Object = "{ADBBAED6-B16F-40DC-80DF-B44910CBA76C}#1.0#0"; "Frame-Xp.ocx"
Begin VB.Form prove0003 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Traspaso de Facturas de Compras Relacionadas"
   ClientHeight    =   8745
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   15270
   DrawStyle       =   5  'Transparent
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   Moveable        =   0   'False
   ScaleHeight     =   583
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1018
   Begin XPFrame.FrameXp FrameQuickMenu 
      Height          =   615
      Left            =   12120
      TabIndex        =   17
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
         TabIndex        =   19
         Top             =   280
         Width           =   1455
      End
      Begin VB.CommandButton botonmisfavoritos 
         Caption         =   "Mis Favoritos"
         Height          =   255
         Left            =   1680
         TabIndex        =   18
         Top             =   280
         Width           =   1335
      End
   End
   Begin XPFrame.FrameXp FrameXp1 
      Height          =   8610
      Left            =   90
      TabIndex        =   2
      Top             =   45
      Width           =   14865
      _ExtentX        =   26220
      _ExtentY        =   15187
      BackColor       =   16761024
      Caption         =   "CENTRALIZACION DE FACTURAS DE COMPRAS RELACIONADAS"
      CaptionEstilo3D =   1
      BackColor       =   16761024
      ForeColor       =   65535
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
      Begin VB.TextBox ORDEN 
         BackColor       =   &H00FFC0C0&
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
         Left            =   12825
         MaxLength       =   10
         TabIndex        =   16
         Top             =   8235
         Width           =   1500
      End
      Begin VB.CommandButton BUSCAR 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Busca Orden"
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
         Left            =   11340
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   8235
         Width           =   1320
      End
      Begin VB.CommandButton Command4 
         BackColor       =   &H00FF8080&
         Caption         =   "TRASPASA CONTABILIDAD"
         Height          =   330
         Left            =   7200
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   8190
         Width           =   2535
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FF8080&
         Caption         =   "IMPRIMIR"
         Height          =   330
         Left            =   3240
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   8190
         Width           =   2130
      End
      Begin XPFrame.FrameXp FrameXp3 
         Height          =   1050
         Left            =   90
         TabIndex        =   5
         Top             =   360
         Width           =   14640
         _ExtentX        =   25823
         _ExtentY        =   1852
         BackColor       =   16761024
         Caption         =   "DATOS DE FILTRADO"
         CaptionEstilo3D =   1
         BackColor       =   16761024
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
         Begin VB.CommandButton Command2 
            Caption         =   "LISTAR"
            Height          =   285
            Left            =   11970
            TabIndex        =   7
            Top             =   675
            Width           =   1455
         End
         Begin XPFrame.FrameXp FrameXp6 
            Height          =   675
            Left            =   90
            TabIndex        =   9
            Top             =   270
            Width           =   3300
            _ExtentX        =   5821
            _ExtentY        =   1191
            BackColor       =   16761024
            Caption         =   "MES"
            CaptionEstilo3D =   1
            BackColor       =   16761024
            ForeColor       =   65535
            ColorBarraArriba=   4194304
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
               TabIndex        =   10
               Top             =   270
               Width           =   3180
            End
         End
         Begin XPFrame.FrameXp FrameXp7 
            Height          =   675
            Left            =   3510
            TabIndex        =   11
            Top             =   270
            Width           =   3075
            _ExtentX        =   5424
            _ExtentY        =   1191
            BackColor       =   16761024
            Caption         =   "A?O"
            CaptionEstilo3D =   1
            BackColor       =   16761024
            ForeColor       =   65535
            ColorBarraArriba=   4194304
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Begin VB.ComboBox COMBOA?O 
               Height          =   315
               Left            =   90
               TabIndex        =   12
               Top             =   270
               Width           =   2865
            End
         End
         Begin XPFrame.FrameXp FrameXp4 
            Height          =   675
            Left            =   6705
            TabIndex        =   13
            Top             =   270
            Width           =   4560
            _ExtentX        =   8043
            _ExtentY        =   1191
            BackColor       =   16761024
            Caption         =   "LOCAL"
            CaptionEstilo3D =   1
            BackColor       =   16761024
            ForeColor       =   65535
            ColorBarraArriba=   4194304
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
               TabIndex        =   14
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
         BackColor       =   16761024
         Caption         =   "LISTADO DE FACTURAS DE FACTURAS RECIBIDAS RELACIONADAS"
         CaptionEstilo3D =   1
         BackColor       =   16761024
         ForeColor       =   65535
         ColorBarraArriba=   8388608
         ColorBarraAbajo =   4194304
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
            Height          =   6360
            Left            =   0
            TabIndex        =   4
            Top             =   270
            Width           =   14595
            _ExtentX        =   25744
            _ExtentY        =   11218
            BackColorFixed  =   16761024
            BackColorSel    =   16761024
            Cols            =   5
            DefaultFontSize =   8.25
            GridColor       =   12640511
            Rows            =   30
            DateFormat      =   2
         End
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
End
Attribute VB_Name = "prove0003"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private localfiltro As String



 

Private Sub BUSCAR_Click()
 Dim i As Integer
 
  For i = 1 To Grid1.Rows - 1
            If Grid1.Cell(i, 2).text = ORDEN.text Then
                Grid1.Range(i, 1, i, Grid1.Cols - 1).Selected
                Grid1.Cell(i, 1).EnsureVisible
                Exit For
            End If
        Next i
End Sub

Private Sub Command1_Click()
imprimir
End Sub



Private Sub COMMAND2_Click()
localfiltro = Mid(ComboLOCAL.text, 1, 2)
a?o = COMBOA?O.text
MES = COMBOMES.ListIndex + 1

leer


End Sub



Private Sub Command3_Click()
Dim k As Integer


End Sub


Private Sub Command4_Click()

Dim MES As String
Dim a?o As String

Dim k As Double
a?o = COMBOA?O.text
MES = Format(COMBOMES.ListIndex + 1, "00")

If estacerrado(a?o + "-" + MES + "-" + Format(fechasistema, "dd")) <> True Then
        

For k = 1 To Grid1.Rows - 1
    If Grid1.Cell(k, 14).text = "1" Then
        Call grabafactura(k, Grid1.Cell(k, 15).text, Grid1.Cell(k, 2).text)
        
    End If
Next k
leer

Else
MsgBox "mes ya cerrado imposible procesar"

End If

End Sub

Private Sub Form_Load()
CENTRAR Me
    Call Conectar_BD
    sc = 0
CARGAGRILLA
Call Conectarventas(Servidor, clientesistema + "ventas00", Usuario, password)
Call Conectargestion(Servidor, clientesistema + "gestion", Usuario, password)
Call Conectargestionrubro(Servidor, clientesistema + "gestion" + rubro, Usuario, password)

For k = 1 To 12
COMBOMES.AddItem MonthName(k)
Next k
COMBOMES.ListIndex = CDbl(Format(fechasistema, "mm") - 1)
For k = 2000 To Val(Format(fechasistema, "yyyy"))
COMBOA?O.AddItem k
Next k
COMBOA?O.ListIndex = k - 2001
LEErlocales


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
titulo = "LISTADO DE FACTURAS EMITIDAS " + COMBOMES.text + " " + COMBOA?O.text
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
    Dim FORMATOGRILLA(10, 20)
    Grid1.DefaultFont.Size = 8
    Grid1.DefaultFont.Bold = False
    
    
    FORMATOGRILLA(1, 1) = "TP"
    FORMATOGRILLA(1, 2) = "NUMERO"
    FORMATOGRILLA(1, 3) = "RUT"
    FORMATOGRILLA(1, 4) = "PROVEEDOR"
    FORMATOGRILLA(1, 5) = "FECHA"
    FORMATOGRILLA(1, 6) = "NETO"
    FORMATOGRILLA(1, 7) = "IVA"
    FORMATOGRILLA(1, 8) = "I.REFRE"
    FORMATOGRILLA(1, 9) = "I.VINO "
    FORMATOGRILLA(1, 10) = "I.LICOR"
    FORMATOGRILLA(1, 11) = "I.HARINA"
    FORMATOGRILLA(1, 12) = "I.CARNE"
    FORMATOGRILLA(1, 13) = "TOTAL  "
    FORMATOGRILLA(1, 14) = "CO"
    FORMATOGRILLA(1, 15) = "TP"
    FORMATOGRILLA(1, 16) = "ORDEN"
    FORMATOGRILLA(1, 17) = "MES"
    FORMATOGRILLA(1, 18) = "A?O"
    
    Rem LARGO DE LOS DATOS
    FORMATOGRILLA(2, 1) = "2"
    FORMATOGRILLA(2, 2) = "10"
    FORMATOGRILLA(2, 3) = "11"
    FORMATOGRILLA(2, 4) = "20"
    FORMATOGRILLA(2, 5) = "10"
    FORMATOGRILLA(2, 6) = "8"
    FORMATOGRILLA(2, 7) = "7"
    FORMATOGRILLA(2, 8) = "7"
    FORMATOGRILLA(2, 9) = "7"
    FORMATOGRILLA(2, 10) = "7"
    FORMATOGRILLA(2, 11) = "7"
    FORMATOGRILLA(2, 12) = "8"
    FORMATOGRILLA(2, 13) = "8"
    FORMATOGRILLA(2, 14) = "3"
    FORMATOGRILLA(2, 15) = "3"
    FORMATOGRILLA(2, 16) = "0"
    FORMATOGRILLA(2, 17) = "4"
    FORMATOGRILLA(2, 18) = "4"

    
    Rem TIPO DE DATOS C=CEROS,S=STRING,N=NUMERICO,D=DATE,
    FORMATOGRILLA(3, 1) = "S"
    FORMATOGRILLA(3, 2) = "N"
    FORMATOGRILLA(3, 3) = "N"
    FORMATOGRILLA(3, 4) = "S"
    FORMATOGRILLA(3, 5) = "D"
    FORMATOGRILLA(3, 6) = "N"
    FORMATOGRILLA(3, 7) = "N"
    FORMATOGRILLA(3, 8) = "N"
    FORMATOGRILLA(3, 9) = "N"
    FORMATOGRILLA(3, 10) = "N"
    FORMATOGRILLA(3, 11) = "N"
    FORMATOGRILLA(3, 12) = "N"
    FORMATOGRILLA(3, 13) = "N"
   
    Rem FORMATO GRILLA
    FORMATOGRILLA(4, 2) = "0000000000"
    FORMATOGRILLA(4, 3) = "0000000000"
    
    FORMATOGRILLA(4, 4) = ""
    FORMATOGRILLA(4, 5) = ""
    
    
    FORMATOGRILLA(4, 6) = "##,###,##0"
    FORMATOGRILLA(4, 7) = "##,###,##0"
    FORMATOGRILLA(4, 8) = "##,###,##0"
    FORMATOGRILLA(4, 9) = "##,###,##0"
    FORMATOGRILLA(4, 10) = "##,###,##0"
    FORMATOGRILLA(4, 11) = "##,###,##0"
    FORMATOGRILLA(4, 12) = "##,###,##0"
    FORMATOGRILLA(4, 13) = "##,###,##0"
    FORMATOGRILLA(4, 16) = "0000000000"
    
    Rem LOCCKED
    For k = 1 To 18
    FORMATOGRILLA(5, k) = "TRUE"
    
    Next k
    
    FORMATOGRILLA(5, 14) = "FALSE"
    
    
    Grid1.Cols = 19
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
        Grid1.Column(k).Width = Val(FORMATOGRILLA(2, k)) * (Grid1.DefaultFont.Size - 1)
        Grid1.Column(k).MaxLength = Val(FORMATOGRILLA(2, k))
        Grid1.Column(k).FormatString = FORMATOGRILLA(4, k)
        Grid1.Column(k).Locked = FORMATOGRILLA(5, k)
        If FORMATOGRILLA(3, k) = "N" Then Grid1.Column(k).Alignment = cellRightCenter
        If FORMATOGRILLA(3, k) = "D" Then Grid1.Column(k).CellType = cellCalendar
        
    Next k
    Grid1.Column(15).Width = 30
    Grid1.Column(1).Width = 30
    Grid1.Column(14).CellType = cellCheckBox
    Grid1.Column(2).Mask = cellNumeric
    Grid1.Column(6).Mask = cellNumeric
    Grid1.Column(7).Mask = cellNumeric
    Grid1.Column(8).Mask = cellNumeric
    Grid1.Column(9).Mask = cellNumeric
    Grid1.Column(10).Mask = cellNumeric
    Grid1.Column(11).Mask = cellNumeric
    Grid1.Column(1).CellType = cellComboBox
    
    Grid1.Column(15).CellType = cellComboBox
    
    
    
    With Grid1.ComboBox(1)
        
        '.Locked = False
        .AutoComplete = True
        
        .AddItem "FA FACTURA" '1
        .AddItem "ND NOTA DEBITO" '2
        .AddItem "NC NOTA CREDITO" '3
        .AddItem "FAE FACTURA ELECTRONICA" '1
        .AddItem "NDE NOTA DEBITO ELECTRONICA" '2
        .AddItem "NCE NOTA CREDITO ELECTRONICA" '3
        .AddItem "OE ORDEN DE ENLACE" '4
        .AddItem "GD DESPACHO" '4
    
    
    End With
    
    With Grid1.ComboBox(15)
        '.Locked = True
        .AutoComplete = True
        .AddItem "MERCADERIAS"
        .AddItem "CIGARRILLOS"
        .AddItem "FRUTAS Y VERDURAS"
        .AddItem "CARNICERIA"
        .AddItem "FIAMBRERIA"
        .AddItem "PANADERIA"
        .AddItem "EMPAQUE"
        .AddItem "DIARIOS"
        
    End With

    
    
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
    Dim MESCONTABLE As Double
    Dim numerodo As String
    
    Dim A?OCONTABLE As Double
    
    
    LINEA = 0: fec = 0: fec1 = 0
    fecha1 = a?o + "-" + MES + "-" + "01"
    fecha2 = a?o + "-" + MES + "-" + "31"
    
        Set csql.ActiveConnection = gestionrubro
        
        csql.sql = "SELECT tipo,numero,localorigen,fecha,monto "
        csql.sql = csql.sql + "FROM l_movimientos_cabeza_" + localfiltro + " "
        csql.sql = csql.sql + "where fecha>='" + fecha1 + "' AND fecha<='" + fecha2 + "' and (tipo='RL' or tipo='RF') "
        csql.sql = csql.sql + "ORDER BY localorigen,numero "
        csql.Execute
        total = 0
        total2 = 0
        Grid1.Rows = 1
        Grid1.AutoRedraw = False
        
        
        If csql.RowsAffected > 0 Then
        Set resultados = csql.OpenResultset
         
         While Not resultados.EOF
'         If resultados(1) = "0000143046" Then Stop
numerodo = resultados(1)
If resultados(0) = "RF" Then
numerodo = leerFOLIOSIIDTE(resultados(2), "FV", resultados(1), resultados(3), "99", resultados(2))
        End If
        
        Call leerfacturaventacontabilidad(numerodo, resultados(2), resultados(0), resultados(3))
        resultados.MoveNext
        
            Wend
End If
      Grid1.AutoRedraw = True
      Grid1.Refresh
      
      
      
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
Sub eliminafactura(tipo, numero)
    Dim resultados As rdoResultset
    Dim csql As New rdoQuery
    
        Set csql.ActiveConnection = ventaslocal
        csql.sql = "delete "
        csql.sql = csql.sql + "FROM sv_documento_cabeza_" + localfiltro + " "
        csql.sql = csql.sql + "where tipo='" + tipo + "' and numero='" + numero + "' "
        csql.Execute
        Call sincronizadatos(csql.sql, ventaslocal, "")
        
        csql.sql = "delete "
        csql.sql = csql.sql + "FROM sv_documento_detalle_" + localfiltro + " "
        csql.sql = csql.sql + "where tipo='" + tipo + "' and numero='" + numero + "' "
        csql.Execute
        Call sincronizadatos(csql.sql, ventaslocal, "")
        
        
        csql.sql = "delete "
        csql.sql = csql.sql + "FROM sv_documento_pagos_" + localfiltro + " "
        csql.sql = csql.sql + "where tipo='" + tipo + "' and numero='" + numero + "' "
        csql.Execute
        Call sincronizadatos(csql.sql, ventaslocal, "")
        
        Set csql.ActiveConnection = gestionrubro
        csql.sql = "delete "
        csql.sql = csql.sql + "FROM l_movimientos_detalle_" + localfiltro + " "
        csql.sql = csql.sql + "where tipo='" + tipo + "' and numero='" + numero + "' "
        csql.Execute
        Call sincronizadatos(csql.sql, gestionrubro, "")
        
        
End Sub


Private Sub Grid1_KeyDown(KeyCode As Integer, ByVal Shift As Integer)

'If KeyCode = 46 Then
'Call eliminafactura(Grid1.Cell(Grid1.ActiveCell.Row, 1).text, Grid1.Cell(Grid1.ActiveCell.Row, 2).text)
'End If
'leer
End Sub

Sub grabafactura(LINEA, tipo, ORDEN)
    Dim netos As Double
    Dim DH As String
    Dim DH2 As String
    Dim mesconta As String
    Dim a?oconta As String
    Dim diaconta As String
    
    Dim exentos As Double
    Dim TIPOCON As String
    Dim CRCC As String
    Dim ELECTRONICA As String
    Dim tipodoc As String
    Dim fecha As Date
    
    campos(0, 0) = "tipo"
    campos(1, 0) = "numero"
    campos(2, 0) = "fecha"
    campos(3, 0) = "fechavencimiento"
    campos(4, 0) = "rut"
    campos(5, 0) = "neto"
    campos(6, 0) = "iva"
    campos(7, 0) = "exento"
    campos(8, 0) = "retencion"
    campos(9, 0) = "total"
    campos(10, 0) = "a?ocontable"
    campos(11, 0) = "mescontable"
    campos(12, 0) = "comentario"
    campos(13, 0) = "electronica"
    campos(14, 0) = "activo"
    campos(15, 0) = "fechadigitacion"
    campos(16, 0) = "folio"
    campos(17, 0) = "impuestoespecifico"
    campos(18, 0) = ""
 
    If Grid1.Cell(LINEA, 1).text = "FA" Then TIPOCON = "1": ELECTRONICA = "N": tipodoc = "FC": DH = "H": DH2 = "D"
    If Grid1.Cell(LINEA, 1).text = "NC" Then TIPOCON = "2": ELECTRONICA = "N": tipodoc = "DC": DH = "H": DH2 = "D"
    If Grid1.Cell(LINEA, 1).text = "ND" Then TIPOCON = "3": ELECTRONICA = "N": tipodoc = "NC": DH = "D": DH2 = "H"
    If Grid1.Cell(LINEA, 1).text = "FAE" Then TIPOCON = "4": ELECTRONICA = "S": tipodoc = "FC": DH = "H": DH2 = "D"
    If Grid1.Cell(LINEA, 1).text = "NDE" Then TIPOCON = "5": ELECTRONICA = "S": tipodoc = "DC": DH = "H": DH2 = "D"
    If Grid1.Cell(LINEA, 1).text = "NCE" Then TIPOCON = "6": ELECTRONICA = "S": tipodoc = "NC": DH = "D": DH2 = "H"
    
    campos(0, 1) = TIPOCON
    campos(1, 1) = Grid1.Cell(LINEA, 2).text
    campos(2, 1) = Format(Grid1.Cell(LINEA, 5).text, "yyyy-mm-dd")
    campos(3, 1) = Format(Grid1.Cell(LINEA, 5).text, "yyyy-mm-dd")
    campos(4, 1) = Mid(Grid1.Cell(LINEA, 3).text, 1, 9) + Mid(Grid1.Cell(LINEA, 3).text, 10, 1)
    campos(5, 1) = Replace(Grid1.Cell(LINEA, 6).text, ",", ".")
    campos(6, 1) = Replace(Grid1.Cell(LINEA, 7).text, ",", ".")
    exentos = CDbl(Grid1.Cell(LINEA, 8).text) + CDbl(Grid1.Cell(LINEA, 9).text) + CDbl(Grid1.Cell(LINEA, 10).text) + CDbl(Grid1.Cell(LINEA, 11).text) + CDbl(Grid1.Cell(LINEA, 12).text)
    campos(7, 1) = Str(exentos)
    campos(8, 1) = "0"
    campos(9, 1) = Replace(Grid1.Cell(LINEA, 13).text, ",", ".")
    
    
    campos(10, 1) = Grid1.Cell(LINEA, 18).text
    campos(11, 1) = Grid1.Cell(LINEA, 17).text
    campos(12, 1) = "CENTRALIZACION AUTOMATICA"
        
    campos(13, 1) = ELECTRONICA
    campos(14, 1) = "N"
    campos(15, 1) = Format(fechasistema, "yyyy-mm-dd")
    
    campos(16, 1) = LEERULTIMOFOLIO(campos(11, 1), campos(10, 1))
    campos(17, 1) = "0"
    
    condicion = ""
    campos(0, 2) = "facturasdecompras"
    op = 2
    sqlconta.response = campos
    Set sqlconta.conexion = contadb

    Call sqlconta.sqlconta(op, condicion)
    k = sqlconta.status
    fecha = Format(campos(3, 1), "yyyy-mm-dd")
    Call grabarcomprobante_lineas(tipodoc, campos(1, 1), "001", fecha, CUENTAPROVEEDOR, "", campos(4, 1), "", "CENTRALIZA DOCUMENTO DE COMPRAS " + Grid1.Cell(LINEA, 1).text, tipodoc, campos(1, 1), campos(2, 1), campos(3, 1), campos(9, 1), DH, USUARIOSISTEMA, campos(11, 1), campos(10, 1), Format(fechasistema, "yyyy-mm-dd"), Time, campos(4, 1))
    Call grabarcomprobante_lineas(tipodoc, campos(1, 1), "002", fecha, ivacredito, "", "", "", "CENTRALIZACION I.V.A", tipodoc, campos(1, 1), campos(2, 1), campos(3, 1), campos(6, 1), DH2, USUARIOSISTEMA, campos(11, 1), campos(10, 1), Format(fechasistema, "yyyy-mm-dd"), Time, campos(4, 1))
    Call grabardetallefactura(LINEA, tipo, ORDEN, fecha, campos(11, 1), campos(10, 1))
     
End Sub

Sub grabardetallefactura(LINEA, tipo, ORDEN, fecha, MES, a?o)
    
    Dim TIPOCON As String
    Dim tipo2 As String
    Dim j As Integer
    Dim lin As Integer
    Dim ilas As Double
    Dim CRCC As String
    Dim cuenta As String
    Dim DH As String
    Dim NOMBRE As String
    Dim tipodoc As String
    
    campos(0, 0) = "tipo"
    campos(1, 0) = "numero"
    campos(2, 0) = "linea"
    campos(3, 0) = "rut"
    campos(4, 0) = "cuentadelmayor"
    campos(5, 0) = "glosa"
    campos(6, 0) = "monto"
    campos(7, 0) = "dh"
    campos(8, 0) = "centrodecosto"
    campos(9, 0) = "rutctacte"
    campos(10, 0) = "fechacreacion"
    campos(11, 0) = ""
    If localfiltro = "00" Then CRCC = "0101"
    If localfiltro = "41" Then CRCC = "0104"
    If localfiltro = "17" Then CRCC = "0101"
    If localfiltro = "42" Then CRCC = "0101"
    
    
    If Grid1.Cell(LINEA, 1).text = "FA" Then TIPOCON = "1": tipodoc = "FC": DH = "D"
    If Grid1.Cell(LINEA, 1).text = "NC" Then TIPOCON = "2": tipodoc = "DC": DH = "D"
    If Grid1.Cell(LINEA, 1).text = "ND" Then TIPOCON = "3": tipodoc = "NC": DH = "H"
    If Grid1.Cell(LINEA, 1).text = "FAE" Then TIPOCON = "4": tipodoc = "FC": DH = "D"
    If Grid1.Cell(LINEA, 1).text = "NDE" Then TIPOCON = "5": tipodoc = "DC": DH = "D"
    If Grid1.Cell(LINEA, 1).text = "NCE" Then TIPOCON = "6": tipodoc = "NC": DH = "H"
    
    If tipo = "DI" Then cuenta = "11350008": NOMBRE = "DIARIOS"
    If tipo = "ME" Then cuenta = "11350001": NOMBRE = "MERCADERIAS"
    If tipo = "CI" Then cuenta = "11350007": NOMBRE = "CIGARRILLOS"
    If tipo = "FR" Then cuenta = "11350002": NOMBRE = "FRUTAS"
    If tipo = "CA" Then cuenta = "11350003": NOMBRE = "CARNICERIA"
    If tipo = "FI" Then cuenta = "11350004": NOMBRE = "FIAMBRERIA"
    If tipo = "PA" Then cuenta = "11350007": NOMBRE = "PANADERIA"
    If tipo = "EM" Then cuenta = "11350006": NOMBRE = "MATERIAL EMPAQUE"
    

Rem CALCULA NETOS

    lin = 3
    pivote.MaxLength = 3
    pivote.text = lin
    
    Call ceros(pivote)
    campos(0, 1) = TIPOCON
    campos(1, 1) = Grid1.Cell(LINEA, 2).text
    campos(2, 1) = pivote.text
    campos(3, 1) = Mid(Grid1.Cell(LINEA, 3).text, 1, 9) + Mid(Grid1.Cell(LINEA, 3).text, 10, 1)
    campos(4, 1) = cuenta
    campos(5, 1) = "F/ERR " + ORDEN + " " + NOMBRE
    campos(6, 1) = Replace(Grid1.Cell(LINEA, 6).text, ",", ".")
    campos(7, 1) = DH
    campos(8, 1) = leerdatoslocal(localfiltro, "codigocrcc")
    campos(9, 1) = campos(3, 1)
    campos(10, 1) = Format(fechasistema, "yyyy-mm-dd")
  
    campos(0, 2) = "facturasdecompras_detalle"
    condicion = ""
    op% = 2
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)
    
    Call grabarcomprobante_lineas(tipodoc, campos(1, 1), campos(2, 1), fecha, campos(4, 1), "", campos(3, 1), "", campos(5, 1), tipodoc, campos(1, 1), fecha, fecha, campos(6, 1), DH, USUARIOSISTEMA, MES, a?o, Format(fechasistema, "yyyy-mm-dd"), Time, campos(3, 1))
    
    
    
Rem CALCULA ILAS refrescos

    ilas = CDbl(Grid1.Cell(LINEA, 8).text)
    If ilas <> 0 Then
    lin = lin + 1
    pivote.MaxLength = 3
    pivote.text = lin
    
    Call ceros(pivote)
    campos(0, 1) = TIPOCON
    campos(1, 1) = Grid1.Cell(LINEA, 2).text
    campos(2, 1) = pivote.text
    campos(3, 1) = Mid(Grid1.Cell(LINEA, 3).text, 1, 9) + Mid(Grid1.Cell(LINEA, 3).text, 10, 1)
    campos(4, 1) = leerdatoslocal(localfiltro, "cuentailarefrescos")
    campos(5, 1) = "F/ERR " + ORDEN + " IMPUESTO ILA REFRESCOS"
    campos(6, 1) = ilas
    campos(7, 1) = DH
    campos(8, 1) = ""
    campos(9, 1) = campos(3, 1)
    campos(10, 1) = Format(fechasistema, "yyyy-mm-dd")
    campos(0, 2) = "facturasdecompras_detalle"
    condicion = ""
    op% = 2
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)
    Call grabarcomprobante_lineas(tipodoc, campos(1, 1), campos(2, 1), fecha, campos(4, 1), "", "", "", campos(5, 1), tipodoc, campos(1, 1), fecha, fecha, campos(6, 1), DH, USUARIOSISTEMA, MES, a?o, Format(fechasistema, "yyyy-mm-dd"), Time, campos(3, 1))
    
    End If
Rem CALCULA ILAS vinos

    ilas = CDbl(Grid1.Cell(LINEA, 9).text)
    If ilas <> 0 Then
    lin = lin + 1
    pivote.MaxLength = 3
    pivote.text = lin
    
    Call ceros(pivote)
    campos(0, 1) = TIPOCON
    campos(1, 1) = Grid1.Cell(LINEA, 2).text
    campos(2, 1) = pivote.text
    campos(3, 1) = Mid(Grid1.Cell(LINEA, 3).text, 1, 9) + Mid(Grid1.Cell(LINEA, 3).text, 10, 1)
    campos(4, 1) = leerdatoslocal(localfiltro, "cuentailavinos")
    campos(5, 1) = "F/ERR " + ORDEN + " IMPUESTO ILA VINOS "
    campos(6, 1) = ilas
    campos(7, 1) = DH
    campos(8, 1) = ""
    campos(9, 1) = campos(3, 1)
    campos(10, 1) = Format(fechasistema, "yyyy-mm-dd")
    campos(0, 2) = "facturasdecompras_detalle"
    condicion = ""
    op% = 2
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)
    Call grabarcomprobante_lineas(tipodoc, campos(1, 1), campos(2, 1), fecha, campos(4, 1), "", "", "", campos(5, 1), tipodoc, campos(1, 1), fecha, fecha, campos(6, 1), DH, USUARIOSISTEMA, MES, a?o, Format(fechasistema, "yyyy-mm-dd"), Time, campos(3, 1))
    
    End If

Rem CALCULA ILAS vinos

    ilas = CDbl(Grid1.Cell(LINEA, 10).text)
    If ilas <> 0 Then
    lin = lin + 1
    pivote.MaxLength = 3
    pivote.text = lin
    
    Call ceros(pivote)
    campos(0, 1) = TIPOCON
    campos(1, 1) = Grid1.Cell(LINEA, 2).text
    campos(2, 1) = pivote.text
    campos(3, 1) = Mid(Grid1.Cell(LINEA, 3).text, 1, 9) + Mid(Grid1.Cell(LINEA, 3).text, 10, 1)
    campos(4, 1) = leerdatoslocal(localfiltro, "cuentailalicores")
    campos(5, 1) = "F/ERR " + ORDEN + " IMPUESTO ILA LICORES "
    campos(6, 1) = ilas
    campos(7, 1) = DH
    campos(8, 1) = ""
    campos(9, 1) = campos(3, 1)
    campos(10, 1) = Format(fechasistema, "yyyy-mm-dd")
    campos(0, 2) = "facturasdecompras_detalle"
    condicion = ""
    op% = 2
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)
    Call grabarcomprobante_lineas(tipodoc, campos(1, 1), campos(2, 1), fecha, campos(4, 1), "", "", "", campos(5, 1), tipodoc, campos(1, 1), fecha, fecha, campos(6, 1), DH, USUARIOSISTEMA, MES, a?o, Format(fechasistema, "yyyy-mm-dd"), Time, campos(3, 1))
    
    End If

Rem CALCULA HARINA
    ilas = CDbl(Grid1.Cell(LINEA, 11).text)
    If ilas <> 0 Then
    lin = lin + 1
    pivote.MaxLength = 3
    pivote.text = lin
    
    Call ceros(pivote)
    campos(0, 1) = TIPOCON
    campos(1, 1) = Grid1.Cell(LINEA, 2).text
    campos(2, 1) = pivote.text
    campos(3, 1) = Mid(Grid1.Cell(LINEA, 3).text, 1, 9) + Mid(Grid1.Cell(LINEA, 3).text, 10, 1)
    campos(4, 1) = leerdatoslocal(localfiltro, "cuentaharina")
    campos(5, 1) = "F/ERR " + ORDEN + " IMPUESTO HARINAS"
    campos(6, 1) = ilas
    campos(7, 1) = DH
    campos(8, 1) = ""
    campos(9, 1) = campos(3, 1)
    campos(10, 1) = Format(fechasistema, "yyyy-mm-dd")
    
    campos(0, 2) = "facturasdecompras_detalle"
    condicion = ""
    op% = 2
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)
    Call grabarcomprobante_lineas(tipodoc, campos(1, 1), campos(2, 1), fecha, campos(4, 1), "", "", "", campos(5, 1), tipodoc, campos(1, 1), fecha, fecha, campos(6, 1), DH, USUARIOSISTEMA, MES, a?o, Format(fechasistema, "yyyy-mm-dd"), Time, campos(3, 1))
    
    End If

Rem CALCULA carne
    ilas = CDbl(Grid1.Cell(LINEA, 12).text)
    If ilas <> 0 Then
    lin = lin + 1
    pivote.MaxLength = 3
    pivote.text = lin
    
    Call ceros(pivote)
    campos(0, 1) = TIPOCON
    campos(1, 1) = Grid1.Cell(LINEA, 2).text
    campos(2, 1) = pivote.text
    campos(3, 1) = Mid(Grid1.Cell(LINEA, 3).text, 1, 9) + Mid(Grid1.Cell(LINEA, 3).text, 10, 1)
    campos(4, 1) = leerdatoslocal(localfiltro, "cuentacarne")
    campos(5, 1) = "F/ERR " + ORDEN + " IMPUESTO CARNE"
    campos(6, 1) = ilas
    campos(7, 1) = DH
    campos(8, 1) = ""
    campos(9, 1) = campos(3, 1)
    campos(10, 1) = Format(fechasistema, "yyyy-mm-dd")
    
    
    campos(0, 2) = "facturasdeventas_detalle"
    condicion = ""
    op% = 2
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)
    Call grabarcomprobante_lineas(tipodoc, campos(1, 1), campos(2, 1), fecha, campos(4, 1), "", "", "", campos(5, 1), tipodoc, campos(1, 1), fecha, fecha, campos(6, 1), DH, USUARIOSISTEMA, MES, a?o, Format(fechasistema, "yyyy-mm-dd"), Time, campos(3, 1))
    
    End If
    
   
    
    
End Sub

Public Function leefactura(tipo, numero, rut) As String

    Dim TIPODO As String
    
    campos(0, 0) = "tipo"
    campos(1, 0) = "numero"
    campos(2, 0) = ""
    If tipo = "FA" Then TIPODO = "1"
    If tipo = "ND" Then TIPODO = "2"
    If tipo = "NC" Then TIPODO = "3"
    If tipo = "FAE" Then TIPODO = "4"
    If tipo = "NDE" Then TIPODO = "5"
    If tipo = "NCE" Then TIPODO = "6"
    
    Rem  condicion = "tipo='" + TIPODO + "' and numero='" + numero + "' and rut='" + rut + "' "
    condicion = "numero='" + numero + "' and rut='" + rut + "' "
    campos(0, 2) = "facturasdecompras"
    op = 5
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    
    Call sqlconta.sqlconta(op, condicion)
    If sqlconta.status = 0 Then
    leefactura = "1"
    
    Else
    leefactura = "0"
    
    End If
    
    

End Function

Sub crearcuentacorriente(rut)
    Dim resultados As rdoResultset
    Dim csql As New rdoQuery
    
        Set csql.ActiveConnection = gestion

            csql.sql = "INSERT INTO " + clientesistema + "conta" + empresaactiva + ".cuentascorrientes "
            csql.sql = csql.sql & "(a?o,tipo,rut,nombre,direccion,comuna,ciudad,giro,fono) "
            csql.sql = csql.sql & "SELECT '" + a?o + "','" + cuentacliente + "',mc.rut,mc.nombre,mc.direccion,mc.comuna,mc.ciudad,mc.giro,mc.fono1 "
            csql.sql = csql.sql & "FROM " & clientesistema & "ventas.sv_maestroclientes as mc "
            csql.sql = csql.sql & "WHERE mc.rut = '" & rut & "' AND mc.sucursal ='0'"
            
            csql.Execute
            Call sincronizadatos(csql.sql, gestion, "")
            
            
            csql.sql = "INSERT INTO " + clientesistema + "conta" + empresaactiva + ".saldosctacte "
            csql.sql = csql.sql & "(a?o,tipo,rut) "
            csql.sql = csql.sql & "SELECT '" + a?o + "','" + cuentacliente + "',mc.rut "
            csql.sql = csql.sql & "FROM " & clientesistema & "ventas.sv_maestroclientes as mc "
            csql.sql = csql.sql & "WHERE mc.rut = '" & rut & "' AND mc.sucursal ='0'"
            
            csql.Execute
            Call sincronizadatos(csql.sql, gestion, "")
            


End Sub
'cSql.SQL = "INSERT INTO l_movimientos_detalle_" & empresaactiva & " "
'            cSql.SQL = cSql.SQL & "(tipo, numero, linea, fecha, rut, codigo, descripcion, cantidad, unidades, precio, total, costoventa, bodega, bodegatraspaso, uxc) "
'            cSql.SQL = cSql.SQL & "SELECT dd.tipo, dd.numero, dd.linea, dd.fecha, dd.rut, dd.codigo, dd.descripcion, dd.cantidad, dd.unidades, dd.precio, dd.total, dd.pcosto, dd.bodega, dd.bodega, ROUND(dd.unidades / dd.cantidad, 0) "
'            cSql.SQL = cSql.SQL & "FROM " & baseVentas & rubro & ".sv_documento_detalle_" + empresaactiva + " as dd "
'            cSql.SQL = cSql.SQL & "WHERE dd.local = '" & empresaactiva & "' AND dd.tipo = '" & v.detalle.tipo & "' AND dd.numero = '" & v.detalle.numero & "'"
'            cSql.Execute

Public Function LEERULTIMOFOLIO(mesconta, a?oconta) As String

    Dim resultados As rdoResultset
    Dim csql As New rdoQuery
    
        Set csql.ActiveConnection = contadb

            csql.sql = "select max(folio) from facturasdecompras where mescontable = '" & Format(mesconta, "00") & "' AND a?ocontable = '" & a?oconta & "' "
            
            csql.Execute
    If csql.RowsAffected > 0 Then
    Set resultados = csql.OpenResultset
    
        If resultados(0) <> "NULO" Then
        LEERULTIMOFOLIO = resultados(0) + 1
        Else
        LEERULTIMOFOLIO = "0000000001"
        End If
        
    End If
    
End Function
Public Function LEERMONTOIMPUESTO(tipo, numero, rut, cuenta) As Double

    Dim resultados As rdoResultset
    Dim csql As New rdoQuery
    
        Set csql.ActiveConnection = gestionrubro

            csql.sql = "select monto from l_ordendecompra_impuestos_" + localfiltro + " where cuenta = '" & cuenta & "' and tipo='" + tipo + "' and numero='" + numero + "' and rut='" + rut + "' "
            
            csql.Execute
    LEERMONTOIMPUESTO = 0
    If csql.RowsAffected > 0 Then
    
    Set resultados = csql.OpenResultset
    LEERMONTOIMPUESTO = resultados(0)
    
    End If
    
End Function
Sub grabarcomprobante_lineas(tipo, numero, LINEA, fecha, codigocuenta, tipoctacte, rutctacte, centrocosto, glosacontable, tipodocumento, numerodocumento, fechadocumento, fechavencimiento, monto, DH, creadopor, MES, a?o, fechacreacion, horacreacion, rutproveedor)
    Dim condicion As String
    Dim campos(40, 3) As String
    Dim op As Integer
    Dim TIPOCON As String
    Dim tipo2 As String
    Dim j As Integer
    Dim lin As String
    Dim lar As Integer
    campos(0, 0) = "tipo"
    campos(1, 0) = "numero"
    campos(2, 0) = "linea"
    campos(3, 0) = "fecha"
    campos(4, 0) = "codigocuenta"
    campos(5, 0) = "tipoctacte"
    campos(6, 0) = "rutctacte"
    campos(7, 0) = "centrocosto"
    campos(8, 0) = "glosacontable"
    campos(9, 0) = "tipodocumento"
    campos(10, 0) = "numerodocumento"
    campos(11, 0) = "fechadocumento"
    campos(12, 0) = "fechavencimiento"
    campos(13, 0) = "monto"
    campos(14, 0) = "dh"
    campos(15, 0) = "creadopor"
    campos(16, 0) = "mes"
    campos(17, 0) = "a?o"
    campos(18, 0) = "fechacreacion"
    campos(19, 0) = "horacreacion"
    campos(20, 0) = "rutproveedor"
    campos(21, 0) = ""
    
    campos(0, 1) = tipo
    campos(1, 1) = numero
    campos(2, 1) = LINEA
    campos(3, 1) = Format(fecha, "yyyy-mm-dd")
    campos(4, 1) = codigocuenta
    campos(5, 1) = tipoctacte
    campos(6, 1) = rutctacte
    campos(7, 1) = centrocosto
    campos(8, 1) = glosacontable
    campos(9, 1) = tipodocumento
    campos(10, 1) = numerodocumento
    campos(11, 1) = Format(fechadocumento, "yyyy-mm-dd")
    campos(12, 1) = Format(fechavencimiento, "yyyy-mm-dd")
    campos(13, 1) = monto

    campos(14, 1) = DH
    campos(15, 1) = creadopor
    campos(16, 1) = MES
    campos(17, 1) = a?o
    
    campos(18, 1) = Format(fechacreacion, "yyyy-mm-dd")
    campos(19, 1) = horacreacion
    campos(20, 1) = rutproveedor

    campos(0, 2) = "movimientoscontables"
   

    op = 2
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    
    Call sqlconta.sqlconta(op, condicion)
   'Call ACTUALIZADOCUMENTO("+")
   
End Sub



Private Sub Grid1_KeyPress(KeyAscii As Integer)
    Static palabra As String
    Dim i As Integer
    Dim largo As Integer
    If KeyAscii = 13 Then
        palabra = ""
    Else
        palabra = palabra + UCase(Chr(KeyAscii))
        largo = Len(palabra)
        For i = 1 To Grid1.Rows - 1
            If Mid(Grid1.Cell(i, 16).text, 1, largo) = palabra Then
                Grid1.Range(i, 1, i, Grid1.Cols - 1).Selected
                Grid1.Cell(i, 1).EnsureVisible
                Exit For
            End If
        Next i
    End If
    
End Sub

Private Sub leerfacturaventa(numero, loc)

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
    Dim tipodoc As String
    Dim MESCONTABLE As Double
    Dim A?OCONTABLE As Double
    
    
        Set csql.ActiveConnection = ventaslocal
        csql.sql = "SELECT dc.tipo,dc.foliosii,dc.rut,dc.fecha,dc.neto,dc.iva,dc.impuestoilarefrescos,dc.impuestoilavinos,dc.impuestoilalicores,dc.impuestoharina,dc.impuestocarne,dc.total,dc.caja "
        csql.sql = csql.sql + "FROM " + clientesistema + "ventas" + loc + ".sv_documento_cabeza_" + loc + " as dc "
        csql.sql = csql.sql + "where dc.tipo='FV' and numero='" + numero + "' "
        csql.sql = csql.sql + "order by dc.foliosii,dc.tipo "
        csql.Execute
        total = 0
        total2 = 0
       
        
        If csql.RowsAffected > 0 Then
        Set resultados = csql.OpenResultset
        fechasum = Format(fechasistema, "yyyy") + "/" + Format(fechasistema, "mm") + "/" + Format(fechasistema, "dd")
        
         While Not resultados.EOF
                     
             tipodoc = "FA"
             If leefactura(tipodoc, resultados(1), leerdatoslocal(loc, "rut")) = "0" Then
             Grid1.Rows = Grid1.Rows + 1
             LINEA = Grid1.Rows - 1
             
             Grid1.Cell(LINEA, 1).text = "FA"
             Grid1.Cell(LINEA, 2).text = resultados(1)
             Grid1.Cell(LINEA, 3).text = leerdatoslocal(loc, "rut")
             Grid1.Cell(LINEA, 4).text = leerdatoslocal(loc, "nombre")
             Grid1.Cell(LINEA, 5).text = resultados(3)
             Grid1.Cell(LINEA, 6).text = resultados(4)
             Grid1.Cell(LINEA, 7).text = resultados(5)
             Grid1.Cell(LINEA, 8).text = resultados(6)
             Grid1.Cell(LINEA, 9).text = resultados(7)
             Grid1.Cell(LINEA, 10).text = resultados(8)
             Grid1.Cell(LINEA, 11).text = resultados(9)
             Grid1.Cell(LINEA, 12).text = resultados(10)
             Grid1.Cell(LINEA, 13).text = resultados(11)
             Grid1.Cell(LINEA, 15).text = "ME"
                         
             MESCONTABLE = CDbl(Format(fechasistema, "mm"))
             A?OCONTABLE = CDbl(Format(fechasistema, "yyyy"))
             If Format(resultados(3), "yyyy-mm") < Format(fechasistema, "yyyy-mm") And Format(fechasistema, "dd") <= diacierrecompra Then
             MESCONTABLE = MESCONTABLE - 1
             If MESCONTABLE = 0 Then MESCONTABLE = 12: A?OCONTABLE = A?OCONTABLE - 1
             
             End If
             
             Grid1.Cell(LINEA, 17).text = Format(MESCONTABLE, "00")
             Grid1.Cell(LINEA, 18).text = A?OCONTABLE
           
           End If
            
             
            
            resultados.MoveNext
       
            Wend
End If
      
End Sub
Private Sub leerfacturaventacontabilidad(numero, loc, TIPODO, fecha)

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
    Dim tipodoc As String
    Dim MESCONTABLE As Double
    Dim A?OCONTABLE As Double
    Dim emprecon As String
    Dim TIPODOCU As String
    
    
    emprecon = leerdatoslocal(loc, "codigocontable")
    
        Rem If TIPODO = "RF" Then numero = leerFOLIOSIIDTE()
        
        Set csql.ActiveConnection = contadb
        csql.sql = "SELECT dc.tipo,dc.numero,dc.rut,dc.fecha,dc.neto,dc.iva,dc.total,dc.caja "
        csql.sql = csql.sql + "FROM " + clientesistema + "conta" + emprecon + ".facturasdeventas as dc "
        csql.sql = csql.sql + "where (dc.tipo='1' or dc.tipo='6') and dc.numero='" + numero + "' and dc.fecha='" + Format(fecha, "yyyy-mm-dd") + "' "
'        If TIPODO = "RL" Then
'        csql.sql = csql.sql + "where dc.tipo='1' and dc.numero='" + numero + "' and dc.fecha='" + Format(fecha, "yyyy-mm-dd") + "' "
'        End If
'        If TIPODO = "RF" Then
'        csql.sql = csql.sql + "where dc.tipo='6' and dc.numero='" + numero + "' and dc.fecha='" + Format(fecha, "yyyy-mm-dd") + "' "
'        End If
'
        csql.sql = csql.sql + ""
        csql.Execute
        total = 0
        total2 = 0
       
        
        If csql.RowsAffected > 0 Then
        Set resultados = csql.OpenResultset
        fechasum = Format(fechasistema, "yyyy") + "/" + Format(fechasistema, "mm") + "/" + Format(fechasistema, "dd")
        
         While Not resultados.EOF
                     
             TIPODOCU = "FA"
             If resultados(0) = "6" Then TIPODOCU = "FAE"
             If leefactura(TIPODOCU, resultados(1), leerdatoslocal(loc, "rut")) = "0" Then
             Grid1.Rows = Grid1.Rows + 1
             LINEA = Grid1.Rows - 1
             
             Grid1.Cell(LINEA, 1).text = TIPODOCU
             Grid1.Cell(LINEA, 2).text = resultados(1)
             Grid1.Cell(LINEA, 3).text = leerdatoslocal(loc, "rut")
             Grid1.Cell(LINEA, 4).text = leerdatoslocal(loc, "nombre")
             Grid1.Cell(LINEA, 5).text = resultados(3)
             Grid1.Cell(LINEA, 6).text = resultados(4)
             Grid1.Cell(LINEA, 7).text = resultados(5)
                
'                formatogrilla(1, 8) = "I.REFRE"
'                formatogrilla(1, 9) = "I.VINO "
'                formatogrilla(1, 10) = "I.LICOR"
'                formatogrilla(1, 11) = "I.HARINA"
'                formatogrilla(1, 12) = "I.CARNE"
    
    
             Grid1.Cell(LINEA, 8).text = leerimpuestoFACTURA("11400010", TIPODOCU, resultados(1), "", emprecon) 'refresco
             Grid1.Cell(LINEA, 9).text = leerimpuestoFACTURA("11400011", TIPODOCU, resultados(1), "", emprecon) 'vino
             Grid1.Cell(LINEA, 10).text = leerimpuestoFACTURA("11400013", TIPODOCU, resultados(1), "", emprecon) 'licor
             Grid1.Cell(LINEA, 11).text = leerimpuestoFACTURA("23200005", TIPODOCU, resultados(1), "", emprecon) 'harina
             Grid1.Cell(LINEA, 12).text = leerimpuestoFACTURA("23200009", TIPODOCU, resultados(1), "", emprecon) 'carne
             Grid1.Cell(LINEA, 13).text = resultados(6)
             Grid1.Cell(LINEA, 15).text = "ME"
                         
             MESCONTABLE = CDbl(Format(fechasistema, "mm"))
             A?OCONTABLE = CDbl(Format(fechasistema, "yyyy"))
             If Format(resultados(3), "yyyy-mm") < Format(fechasistema, "yyyy-mm") And Format(fechasistema, "dd") <= diacierrecompra Then
             MESCONTABLE = MESCONTABLE - 1
             If MESCONTABLE = 0 Then MESCONTABLE = 12: A?OCONTABLE = A?OCONTABLE - 1
             
             End If
             
             Grid1.Cell(LINEA, 17).text = Format(MESCONTABLE, "00")
             Grid1.Cell(LINEA, 18).text = A?OCONTABLE
           
           End If
            
             
            
            resultados.MoveNext
       
            Wend
End If
      
End Sub
Private Sub ORDEN_GotFocus()
Call cargatexto(ORDEN)
End Sub

Private Sub ORDEN_KeyPress(KeyAscii As Integer)
KeyAscii = esNumero(KeyAscii)
If KeyAscii = 13 Then
Call ceros(ORDEN)
Call BUSCAR_Click
'Grid1.Cell(Grid1.ActiveCell.Row, 14).SetFocus

End If

End Sub
Private Sub botonmisaccesos_Click()
programafiltro = Me.Caption
misaccesos.Show
End Sub


Private Sub botonmisfavoritos_Click()
Call AgregaFavoritos(USUARIOSISTEMA, App.EXEName, Me.Name, Me.Caption)
End Sub
