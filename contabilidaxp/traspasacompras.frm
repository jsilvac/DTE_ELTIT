VERSION 5.00
Object = "{4777436C-EB8C-4596-98A8-EBCF98683969}#1.0#0"; "FlexCell.ocx"
Object = "{ADBBAED6-B16F-40DC-80DF-B44910CBA76C}#1.0#0"; "Frame-Xp.ocx"
Begin VB.Form prove0001 
   Appearance      =   0  'Flat
   BackColor       =   &H0080C0FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Traspaso de Facturas de Compras"
   ClientHeight    =   8595
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   15270
   DrawStyle       =   5  'Transparent
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   Moveable        =   0   'False
   ScaleHeight     =   573
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1018
   Begin XPFrame.FrameXp FrameQuickMenu 
      Height          =   615
      Left            =   12120
      TabIndex        =   19
      Top             =   0
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   1085
      BackColor       =   8438015
      Caption         =   " Mis Datos"
      BackColor       =   8438015
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
      Begin VB.CommandButton botonmisfavoritos 
         Caption         =   "Mis Favoritos"
         Height          =   255
         Left            =   1680
         TabIndex        =   21
         Top             =   280
         Width           =   1335
      End
      Begin VB.CommandButton botonmisaccesos 
         Caption         =   "Permisos Modulo"
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   280
         Width           =   1455
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
      Visible         =   0   'False
      Width           =   135
   End
   Begin XPFrame.FrameXp FrameXp1 
      Height          =   8610
      Left            =   75
      TabIndex        =   2
      Top             =   0
      Width           =   15135
      _ExtentX        =   26696
      _ExtentY        =   15187
      BackColor       =   12648447
      Caption         =   "CENTRALIZACION DE FACTURAS DE COMPRAS"
      CaptionEstilo3D =   1
      BackColor       =   12648447
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
         TabIndex        =   16
         Top             =   8190
         Width           =   1320
      End
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
         TabIndex        =   15
         Top             =   8190
         Width           =   1500
      End
      Begin VB.CommandButton Command4 
         BackColor       =   &H0080FF80&
         Caption         =   "TRASPASA CONTABILIDAD"
         Height          =   330
         Left            =   8160
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   8190
         Width           =   2535
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H0080FF80&
         Caption         =   "IMPRIMIR"
         Height          =   330
         Left            =   5520
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   8190
         Width           =   2130
      End
      Begin XPFrame.FrameXp FrameXp3 
         Height          =   1050
         Left            =   135
         TabIndex        =   5
         Top             =   360
         Width           =   14910
         _ExtentX        =   26300
         _ExtentY        =   1852
         BackColor       =   8438015
         Caption         =   "DATOS DE FILTRADO"
         CaptionEstilo3D =   1
         BackColor       =   8438015
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
         Begin VB.OptionButton Option2 
            BackColor       =   &H0080C0FF&
            Caption         =   "Mensual"
            Height          =   255
            Left            =   13200
            TabIndex        =   18
            Top             =   360
            Width           =   1575
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H0080C0FF&
            Caption         =   "Diario"
            Height          =   255
            Left            =   11640
            TabIndex        =   17
            Top             =   360
            Value           =   -1  'True
            Width           =   975
         End
         Begin VB.CommandButton Command2 
            Caption         =   "LISTAR"
            Height          =   285
            Left            =   12120
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
            BackColor       =   8438015
            Caption         =   "MES"
            CaptionEstilo3D =   1
            BackColor       =   8438015
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
            BackColor       =   8438015
            Caption         =   "AÑO"
            CaptionEstilo3D =   1
            BackColor       =   8438015
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
            Begin VB.ComboBox COMBOAÑO 
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
            BackColor       =   8438015
            Caption         =   "LOCAL"
            CaptionEstilo3D =   1
            BackColor       =   8438015
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
         Width           =   14910
         _ExtentX        =   26300
         _ExtentY        =   11774
         BackColor       =   8438015
         Caption         =   "LISTADO DE FACTURAS DE FACTURAS RECIBIDAS"
         CaptionEstilo3D =   1
         BackColor       =   8438015
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
            Width           =   14865
            _ExtentX        =   26220
            _ExtentY        =   11218
            BackColorFixed  =   12640511
            BackColorSel    =   12648447
            Cols            =   5
            DefaultFontSize =   8.25
            GridColor       =   12640511
            Rows            =   30
            DateFormat      =   2
         End
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Proveedor Electronico, Documento Normal, Revisar"
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
         Left            =   840
         TabIndex        =   23
         Top             =   8190
         Width           =   4515
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H000000FF&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   8190
         Width           =   615
      End
   End
End
Attribute VB_Name = "prove0001"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private localfiltro As String


Private Sub BUSCAR_Click()
 Dim i As Integer
 
  For i = 1 To Grid1.Rows - 1
            If Mid(Grid1.Cell(i, 18).text, 1, 10) = ORDEN.text Then
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
rubro = leerrubrocomercio(localfiltro)
Call Conectargestionrubro(Servidor, clientesistema + "gestion" + rubro, Usuario, password)


año = COMBOAÑO.text
MES = COMBOMES.ListIndex + 1

leer


End Sub



Private Sub Command3_Click()
Dim k As Integer


End Sub


Private Sub Command4_Click()

Dim MES As String
Dim año As String

Dim k As Double
año = COMBOAÑO.text
MES = Format(COMBOMES.ListIndex + 1, "00")

If estacerrado(año + "-" + MES + "-" + Format(fechasistema, "dd")) <> True Then
        

For k = 1 To Grid1.Rows - 1
    If Grid1.Cell(k, 16).text = "1" Then
        Call grabafactura(k, Grid1.Cell(k, 17).text, Grid1.Cell(k, 18).text)
        
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
'rubro = leerrubrocomercio(ComboLOCAL.text)
'Call Conectargestionrubro(servidor, clientesistema + "gestion" + rubro, Usuario, password)

For k = 1 To 12
COMBOMES.AddItem MonthName(k)
Next k
COMBOMES.ListIndex = CDbl(Format(fechasistema, "mm") - 1)
For k = 2000 To Val(Format(fechasistema, "yyyy"))
COMBOAÑO.AddItem k
Next k
COMBOAÑO.ListIndex = k - 2001
LEErlocales
rubro = leerrubrocomercio(localfiltro)
Call Conectargestionrubro(Servidor, clientesistema + "gestion" + rubro, Usuario, password)


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
titulo = "LISTADO DE FACTURAS EMITIDAS " + COMBOMES.text + " " + COMBOAÑO.text
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
    Dim FORMATOGRILLA(10, 22)
    Grid1.DefaultFont.Size = 8
    Grid1.DefaultFont.Bold = False
    
    
    FORMATOGRILLA(1, 1) = "TP"
    FORMATOGRILLA(1, 2) = "NUMERO"
    FORMATOGRILLA(1, 3) = "RUT"
    FORMATOGRILLA(1, 4) = "PROVEEDOR"
    FORMATOGRILLA(1, 5) = "FECHA"
    FORMATOGRILLA(1, 6) = "NETO"
    FORMATOGRILLA(1, 7) = "IVA"
    FORMATOGRILLA(1, 8) = "I.CERV"
    FORMATOGRILLA(1, 9) = "I.AZUCA"
    FORMATOGRILLA(1, 10) = "I.VINO "
    FORMATOGRILLA(1, 11) = "I.LICOR"
    FORMATOGRILLA(1, 12) = "I.HARINA"
    FORMATOGRILLA(1, 13) = "I.CARNE"
    FORMATOGRILLA(1, 14) = "I.NO AZU"
    
    FORMATOGRILLA(1, 15) = "TOTAL  "
    FORMATOGRILLA(1, 16) = "CO"
    FORMATOGRILLA(1, 17) = "TP"
    FORMATOGRILLA(1, 18) = "ORDEN"
    FORMATOGRILLA(1, 19) = "MES"
    FORMATOGRILLA(1, 20) = "AÑO"
    FORMATOGRILLA(1, 21) = "RECEPCION"
    
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
    FORMATOGRILLA(2, 12) = "7"
    FORMATOGRILLA(2, 13) = "8"
    FORMATOGRILLA(2, 14) = "8"
    FORMATOGRILLA(2, 15) = "8"
    FORMATOGRILLA(2, 16) = "3"
    FORMATOGRILLA(2, 17) = "2"
    FORMATOGRILLA(2, 18) = "10"
    FORMATOGRILLA(2, 19) = "4"
    FORMATOGRILLA(2, 20) = "4"
    FORMATOGRILLA(2, 21) = "10"
    
    
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
    FORMATOGRILLA(3, 14) = "N"
    FORMATOGRILLA(3, 15) = "N"
    FORMATOGRILLA(3, 16) = "D"
   
    Rem FORMATO GRILLA
    FORMATOGRILLA(4, 2) = "0000000000"
    FORMATOGRILLA(4, 3) = ""
    
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
    FORMATOGRILLA(4, 14) = "##,###,##0"
    FORMATOGRILLA(4, 15) = "##,###,##0"
    
    FORMATOGRILLA(4, 18) = "0000000000"
    
    Rem LOCCKED
    For k = 1 To 21
    FORMATOGRILLA(5, k) = "TRUE"
    
    Next k
    
    FORMATOGRILLA(5, 16) = "FALSE"
    
    
    Grid1.Cols = 22
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
    Grid1.Column(17).Width = 30
    Grid1.Column(1).Width = 30
    Grid1.Column(16).CellType = cellCheckBox
    Grid1.Column(2).Mask = cellNumeric
    Grid1.Column(6).Mask = cellNumeric
    Grid1.Column(7).Mask = cellNumeric
    Grid1.Column(8).Mask = cellNumeric
    Grid1.Column(9).Mask = cellNumeric
    Grid1.Column(10).Mask = cellNumeric
    Grid1.Column(11).Mask = cellNumeric
    Grid1.Column(12).Mask = cellNumeric
    
    Grid1.Column(1).CellType = cellComboBox
    
    Grid1.Column(17).CellType = cellComboBox
    
    
    
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
    
    With Grid1.ComboBox(17)
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
    
    Dim AÑOCONTABLE As Double
    
    
    LINEA = 0: fec = 0: fec1 = 0
    fecha1 = año + "-" + MES + "-" + "01"
    fecha2 = año + "-" + MES + "-" + "31"
    
        Set csql.ActiveConnection = gestionrubro
        csql.sql = "SELECT lof.tipo,lof.numero,lof.rut,lof.fecha,lof.neto,lof.iva,lof.total,lof.categoria,lof.bonificacion,loc.fecharecepcion,lof.ordendecompra "
        csql.sql = csql.sql + "FROM l_ordendecompra_detalle_facturas_" + localfiltro + " as lof,l_ordendecompra_cabeza_" + localfiltro + " as loc "
        csql.sql = csql.sql + "where lof.ordendecompra=loc.numero and "
        If Option1.Value = False Then
        csql.sql = csql.sql + "loc.fecharecepcion>='" + fecha1 + "' AND loc.fecharecepcion<='" + fecha2 + "' "
        Else
        csql.sql = csql.sql + "loc.fecharecepcion='" & Format(fechasistema, "yyyy-mm-dd") + "' "
        End If
        
        csql.sql = csql.sql + "and (lof.tipo='FEE' OR lof.tipo='FE' OR lof.tipo='FA' or lof.tipo='NC' or lof.tipo='ND' or lof.tipo='FAE' or lof.tipo='NCE' or lof.tipo='NDE') order by loc.fecharecepcion,lof.ordendecompra "
        csql.sql = csql.sql + ""
        csql.Execute
        total = 0
        total2 = 0
        Grid1.Rows = 1
        Grid1.AutoRedraw = False
        
        
        If csql.RowsAffected > 0 Then
        Set resultados = csql.OpenResultset
        fechasum = Format(fechasistema, "yyyy") + "/" + Format(fechasistema, "mm") + "/" + Format(fechasistema, "dd")
        
         While Not resultados.EOF
             
             If leefactura(resultados(0), resultados(1), resultados(2)) = "0" Then
             Grid1.Rows = Grid1.Rows + 1
             
             LINEA = LINEA + 1
             Grid1.Cell(LINEA, 1).text = resultados(0)
             Grid1.Cell(LINEA, 2).text = Format(resultados(1), "0000000000")
             Grid1.Cell(LINEA, 3).text = Mid(resultados(2), 1, 9) + "-" + Mid(resultados(2), 10, 1)
             Grid1.Cell(LINEA, 4).text = nombrectacte(resultados(2))
             If IsNull(resultados(3)) = False Then
             Grid1.Cell(LINEA, 5).text = resultados(3)
             End If
             Grid1.Cell(LINEA, 6).text = resultados(4)
             Grid1.Cell(LINEA, 7).text = resultados(5)
             Grid1.Cell(LINEA, 8).text = LEERMONTOIMPUESTO(resultados(0), resultados(1), resultados(10), "11400014")
             Grid1.Cell(LINEA, 9).text = LEERMONTOIMPUESTO(resultados(0), resultados(1), resultados(10), "11400010")
             Grid1.Cell(LINEA, 10).text = LEERMONTOIMPUESTO(resultados(0), resultados(1), resultados(10), "11400011")
             Grid1.Cell(LINEA, 11).text = LEERMONTOIMPUESTO(resultados(0), resultados(1), resultados(10), "11400013") + LEERMONTOIMPUESTO(resultados(0), resultados(1), resultados(2), "11400014") + LEERMONTOIMPUESTO(resultados(0), resultados(1), resultados(2), "11400015")
             Grid1.Cell(LINEA, 12).text = LEERMONTOIMPUESTO(resultados(0), resultados(1), resultados(10), "11400005")
             Grid1.Cell(LINEA, 13).text = LEERMONTOIMPUESTO(resultados(0), resultados(1), resultados(10), "11400012")
             Grid1.Cell(LINEA, 14).text = LEERMONTOIMPUESTO(resultados(0), resultados(1), resultados(10), "11400017")
             Grid1.Cell(LINEA, 15).text = resultados(6)
             Grid1.Cell(LINEA, 16).text = "0"
             Grid1.Cell(LINEA, 17).text = resultados(7)
             Grid1.Cell(LINEA, 18).text = resultados(10)
             MESCONTABLE = CDbl(Format(fechasistema, "mm"))
             AÑOCONTABLE = CDbl(Format(fechasistema, "yyyy"))
             If Format(resultados(3), "yyyy-mm") < Format(fechasistema, "yyyy-mm") And Format(fechasistema, "dd") <= diacierrecompra Then
             MESCONTABLE = MESCONTABLE - 1
             If MESCONTABLE = 0 Then MESCONTABLE = 12: AÑOCONTABLE = AÑOCONTABLE - 1
             
             End If
             
             Grid1.Cell(LINEA, 19).text = Format(MESCONTABLE, "00")
             Grid1.Cell(LINEA, 20).text = AÑOCONTABLE
             Grid1.Cell(LINEA, 21).text = resultados(9)
             If proveedorelectronico(resultados(2)) = True Then
                If resultados(0) <> "FAE" And resultados(0) <> "NDE" And resultados(0) <> "NCE" And resultados(0) <> "FEE" Then
                    Grid1.Range(LINEA, 1, LINEA, Grid1.Cols - 1).BackColor = vbRed
                End If
             End If
             
            End If
            resultados.MoveNext
       
            Wend
End If
      Grid1.AutoRedraw = True
      Grid1.Refresh
      
      
      
End Sub

Function eselectronico(rutconableprove) As Boolean
     Dim csql As New rdoQuery
     Dim resultados As rdoResultset
     Set csql.ActiveConnection = conta
     csql.sql = "select rut from " & cliente_sql & "conta.proveedores_cuenta "
     csql.sql = csql.sql & " where rut='" & rutconableprove & "' "
     
     
     csql.Execute
        eselectronico = False
        
     If csql.RowsAffected > 0 Then
        eselectronico = True
     End If
     csql.Close
     Set csql = Nothing
     
End Function
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


Private Sub Grid1_Click()
If Grid1.ActiveCell.col = 16 Then
If Mid(Grid1.Cell(Grid1.ActiveCell.row, 4).text, 1, 3) = "***" Then
Grid1.Cell(Grid1.ActiveCell.row, 16).text = "0"
End If
End If


End Sub

Private Sub Grid1_DblClick()
If Grid1.ActiveCell.col = 16 Then
If Mid(Grid1.Cell(Grid1.ActiveCell.row, 4).text, 1, 3) = "***" Then
Grid1.Cell(Grid1.ActiveCell.row, 16).text = "0"
End If
End If

localorden = localfiltro
Rcompra02.dato1.text = Grid1.Cell(Grid1.ActiveCell.row, 18).text

Rcompra02.Show vbModal


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
    Dim añoconta As String
    Dim diaconta As String
    Dim CUENTA2 As String
    
    Dim exentos As Double
    Dim TIPOCON As String
    Dim CRCC As String
    Dim ELECTRONICA As String
    Dim tipodoc As String
    Dim fecha As Date
    Dim fechacom As Date
    
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
    campos(10, 0) = "añocontable"
    campos(11, 0) = "mescontable"
    campos(12, 0) = "comentario"
    campos(13, 0) = "electronica"
    campos(14, 0) = "activo"
    campos(15, 0) = "fechadigitacion"
    campos(16, 0) = "folio"
    campos(17, 0) = "impuestoespecifico"
    campos(18, 0) = "usuario"
    campos(19, 0) = "fechatraspaso"
    campos(20, 0) = "horatraspaso"
    campos(21, 0) = ""
    
    
    
 
    If Grid1.Cell(LINEA, 1).text = "FA" Then TIPOCON = "1": ELECTRONICA = "N": tipodoc = "FC": DH = "H": DH2 = "D"
    If Grid1.Cell(LINEA, 1).text = "ND" Then TIPOCON = "2": ELECTRONICA = "N": tipodoc = "DC": DH = "H": DH2 = "D"
    If Grid1.Cell(LINEA, 1).text = "NC" Then TIPOCON = "3": ELECTRONICA = "N": tipodoc = "NC": DH = "D": DH2 = "H"
    If Grid1.Cell(LINEA, 1).text = "FAE" Then TIPOCON = "4": ELECTRONICA = "S": tipodoc = "FC": DH = "H": DH2 = "D"
    If Grid1.Cell(LINEA, 1).text = "NDE" Then TIPOCON = "5": ELECTRONICA = "S": tipodoc = "DC": DH = "H": DH2 = "D"
    If Grid1.Cell(LINEA, 1).text = "NCE" Then TIPOCON = "6": ELECTRONICA = "S": tipodoc = "NC": DH = "D": DH2 = "H"
    If Grid1.Cell(LINEA, 1).text = "FEE" Then TIPOCON = "0": ELECTRONICA = "S": tipodoc = "EE": DH = "H": DH2 = "D": exentos = Replace(Grid1.Cell(LINEA, 15).text, ",", ".")
    If Grid1.Cell(LINEA, 1).text = "FE" Then TIPOCON = "9": ELECTRONICA = "N": tipodoc = "EN": DH = "H": DH2 = "D"
    
    
    campos(0, 1) = TIPOCON
    campos(1, 1) = Format(Grid1.Cell(LINEA, 2).text, "0000000000")
    campos(2, 1) = Format(Grid1.Cell(LINEA, 5).text, "yyyy-mm-dd")
    campos(3, 1) = Format(Grid1.Cell(LINEA, 5).text, "yyyy-mm-dd")
    campos(4, 1) = Mid(Grid1.Cell(LINEA, 3).text, 1, 9) + Mid(Grid1.Cell(LINEA, 3).text, 11, 1)
    campos(5, 1) = Replace(Grid1.Cell(LINEA, 6).text, ",", ".")
    campos(6, 1) = Replace(Grid1.Cell(LINEA, 7).text, ",", ".")
    exentos = CDbl(Grid1.Cell(LINEA, 8).text) + CDbl(Grid1.Cell(LINEA, 9).text) + CDbl(Grid1.Cell(LINEA, 10).text) + CDbl(Grid1.Cell(LINEA, 11).text) + CDbl(Grid1.Cell(LINEA, 12).text) + CDbl(Grid1.Cell(LINEA, 13).text) + CDbl(Grid1.Cell(LINEA, 14).text)
    If Grid1.Cell(LINEA, 1).text = "FEE" Then
    exentos = Replace(Grid1.Cell(LINEA, 15).text, ",", ".")
    End If
    
    
    
    campos(7, 1) = Str(exentos)
    campos(8, 1) = "0"
    campos(9, 1) = Replace(Grid1.Cell(LINEA, 15).text, ",", ".")
    
    
    campos(10, 1) = Grid1.Cell(LINEA, 20).text
    campos(11, 1) = Grid1.Cell(LINEA, 19).text
    campos(12, 1) = "CENTRALIZACION AUTOMATICA"
        
    campos(13, 1) = ELECTRONICA
    campos(14, 1) = "N"
    campos(15, 1) = Format(fechasistema, "yyyy-mm-dd")
    
    campos(16, 1) = LEERULTIMOFOLIO(campos(11, 1), campos(10, 1))
    campos(17, 1) = "0"
    campos(18, 1) = USUARIOSISTEMA
    campos(19, 1) = Format(fechasistema, "yyyy-mm-dd")
    campos(20, 1) = Time
    
    
    condicion = ""
    campos(0, 2) = "facturasdecompras"
    op = 2
    sqlconta.response = campos
    Set sqlconta.conexion = contadb

    Call sqlconta.sqlconta(op, condicion)
    k = sqlconta.status
    fecha = Format(campos(3, 1), "yyyy-mm-dd")
    
    
    fechacom = Format(fechasistema, "yyyy-mm") + "-" + "01"
    If fecha >= fechacom Then
    fechacom = fecha
    End If
    
    If TIPOCON = "3" Or TIPOCON = "6" Then
    CUENTA2 = "11200044"
    Else
    CUENTA2 = CUENTAPROVEEDOR
    
    End If
    
    
    Call grabarcomprobante_lineas(tipodoc, campos(1, 1), "001", fechacom, CUENTA2, "", campos(4, 1), "", "CENTRALIZA DOCUMENTO DE COMPRAS " + Grid1.Cell(LINEA, 1).text, tipodoc, campos(1, 1), campos(2, 1), campos(3, 1), campos(9, 1), DH, USUARIOSISTEMA, campos(11, 1), campos(10, 1), Format(fechasistema, "yyyy-mm-dd"), Time, campos(4, 1))
    If Grid1.Cell(LINEA, 1).text <> "FEE" And Grid1.Cell(LINEA, 1).text <> "FE" Then
    Call grabarcomprobante_lineas(tipodoc, campos(1, 1), "002", fechacom, ivacredito, "", campos(4, 1), "", "CENTRALIZACION I.V.A", tipodoc, campos(1, 1), campos(2, 1), campos(3, 1), campos(6, 1), DH2, USUARIOSISTEMA, campos(11, 1), campos(10, 1), Format(fechasistema, "yyyy-mm-dd"), Time, campos(4, 1))
    End If
    
    Call grabardetallefactura(LINEA, tipo, ORDEN, fechacom, campos(11, 1), campos(10, 1))


End Sub

Sub grabardetallefactura(LINEA, tipo, ORDEN, fecha, MES, año)
    
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
    If Grid1.Cell(LINEA, 1).text = "ND" Then TIPOCON = "2": tipodoc = "DC": DH = "D"
    If Grid1.Cell(LINEA, 1).text = "NC" Then TIPOCON = "3": tipodoc = "NC": DH = "H"
    If Grid1.Cell(LINEA, 1).text = "FAE" Then TIPOCON = "4": tipodoc = "FC": DH = "D"
    If Grid1.Cell(LINEA, 1).text = "NDE" Then TIPOCON = "5": tipodoc = "DC": DH = "D"
    If Grid1.Cell(LINEA, 1).text = "NCE" Then TIPOCON = "6": tipodoc = "NC": DH = "H"
    If Grid1.Cell(LINEA, 1).text = "FEE" Then TIPOCON = "0": tipodoc = "EE": DH = "D"
    If Grid1.Cell(LINEA, 1).text = "FE" Then TIPOCON = "9": tipodoc = "EN": DH = "D"
    
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
    campos(3, 1) = Mid(Grid1.Cell(LINEA, 3).text, 1, 9) + Mid(Grid1.Cell(LINEA, 3).text, 11, 1)
    campos(4, 1) = cuenta
    campos(5, 1) = "O/C " + ORDEN + " " + NOMBRE
    If TIPOCON = "9" Or TIPOCON = "0" Then
    campos(6, 1) = Replace(Grid1.Cell(LINEA, 15).text, ",", ".")
    Else
    campos(6, 1) = Replace(Grid1.Cell(LINEA, 6).text, ",", ".")
    
    End If
    
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
    
    Call grabarcomprobante_lineas(tipodoc, campos(1, 1), campos(2, 1), fecha, campos(4, 1), "", campos(3, 1), "", campos(5, 1), tipodoc, campos(1, 1), fecha, fecha, campos(6, 1), DH, USUARIOSISTEMA, MES, año, Format(fechasistema, "yyyy-mm-dd"), Time, campos(3, 1))
    
    
    
Rem CALCULA ILAS CERVEZAS

    ilas = CDbl(Grid1.Cell(LINEA, 8).text)
    If ilas <> 0 Then
    lin = lin + 1
    pivote.MaxLength = 3
    pivote.text = lin
    
    Call ceros(pivote)
    campos(0, 1) = TIPOCON
    campos(1, 1) = Grid1.Cell(LINEA, 2).text
    campos(2, 1) = pivote.text
    campos(3, 1) = Mid(Grid1.Cell(LINEA, 3).text, 1, 9) + Mid(Grid1.Cell(LINEA, 3).text, 11, 1)
    
   
             campos(4, 1) = leerdatoslocal(localfiltro, "cuentailacervezas")
 
        

   
    
    campos(5, 1) = "O/C " + ORDEN + " IMPUESTO ILA CERVEZAS"
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
    Call grabarcomprobante_lineas(tipodoc, campos(1, 1), campos(2, 1), fecha, campos(4, 1), "", "", "", campos(5, 1), tipodoc, campos(1, 1), fecha, fecha, campos(6, 1), DH, USUARIOSISTEMA, MES, año, Format(fechasistema, "yyyy-mm-dd"), Time, campos(3, 1))
    
    End If
Rem CALCULA ILAS refrescos

    ilas = CDbl(Grid1.Cell(LINEA, 9).text)
    If ilas <> 0 Then
    lin = lin + 1
    pivote.MaxLength = 3
    pivote.text = lin
    
    Call ceros(pivote)
    campos(0, 1) = TIPOCON
    campos(1, 1) = Grid1.Cell(LINEA, 2).text
    campos(2, 1) = pivote.text
    campos(3, 1) = Mid(Grid1.Cell(LINEA, 3).text, 1, 9) + Mid(Grid1.Cell(LINEA, 3).text, 11, 1)
    
     
             campos(4, 1) = leerdatoslocal(localfiltro, "cuentailarefrescos")
        
        
        
    campos(5, 1) = "O/C " + ORDEN + " IMPUESTO ILA REFRESCOS"
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
    Call grabarcomprobante_lineas(tipodoc, campos(1, 1), campos(2, 1), fecha, campos(4, 1), "", "", "", campos(5, 1), tipodoc, campos(1, 1), fecha, fecha, campos(6, 1), DH, USUARIOSISTEMA, MES, año, Format(fechasistema, "yyyy-mm-dd"), Time, campos(3, 1))
    
    End If

Rem CALCULA ILAS AZUCAR

    ilas = CDbl(Grid1.Cell(LINEA, 14).text)
    If ilas <> 0 Then
    lin = lin + 1
    pivote.MaxLength = 3
    pivote.text = lin
    
    Call ceros(pivote)
    campos(0, 1) = TIPOCON
    campos(1, 1) = Grid1.Cell(LINEA, 2).text
    campos(2, 1) = pivote.text
    campos(3, 1) = Mid(Grid1.Cell(LINEA, 3).text, 1, 9) + Mid(Grid1.Cell(LINEA, 3).text, 11, 1)
    
    
    
            campos(4, 1) = leerdatoslocal(localfiltro, "cuentailanoazucar")
       
        
        
    campos(5, 1) = "O/C " + ORDEN + " IMPUESTO ILA NO AZUCAR"
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
    Call grabarcomprobante_lineas(tipodoc, campos(1, 1), campos(2, 1), fecha, campos(4, 1), "", "", "", campos(5, 1), tipodoc, campos(1, 1), fecha, fecha, campos(6, 1), DH, USUARIOSISTEMA, MES, año, Format(fechasistema, "yyyy-mm-dd"), Time, campos(3, 1))
    
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
    campos(3, 1) = Mid(Grid1.Cell(LINEA, 3).text, 1, 9) + Mid(Grid1.Cell(LINEA, 3).text, 11, 1)
   
    
  
            campos(4, 1) = leerdatoslocal(localfiltro, "cuentailavinos")
    
        
        
    campos(5, 1) = "O/C " + ORDEN + " IMPUESTO ILA VINOS "
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
    Call grabarcomprobante_lineas(tipodoc, campos(1, 1), campos(2, 1), fecha, campos(4, 1), "", "", "", campos(5, 1), tipodoc, campos(1, 1), fecha, fecha, campos(6, 1), DH, USUARIOSISTEMA, MES, año, Format(fechasistema, "yyyy-mm-dd"), Time, campos(3, 1))
    
    End If

Rem CALCULA ILAS licores

    ilas = CDbl(Grid1.Cell(LINEA, 11).text)
    If ilas <> 0 Then
    lin = lin + 1
    pivote.MaxLength = 3
    pivote.text = lin
    
    Call ceros(pivote)
    campos(0, 1) = TIPOCON
    campos(1, 1) = Grid1.Cell(LINEA, 2).text
    campos(2, 1) = pivote.text
    campos(3, 1) = Mid(Grid1.Cell(LINEA, 3).text, 1, 9) + Mid(Grid1.Cell(LINEA, 3).text, 11, 1)
    
    
    
    
            campos(4, 1) = leerdatoslocal(localfiltro, "cuentailalicores")
     
        
        
    campos(5, 1) = "O/C " + ORDEN + " IMPUESTO ILA LICORES "
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
    Call grabarcomprobante_lineas(tipodoc, campos(1, 1), campos(2, 1), fecha, campos(4, 1), "", "", "", campos(5, 1), tipodoc, campos(1, 1), fecha, fecha, campos(6, 1), DH, USUARIOSISTEMA, MES, año, Format(fechasistema, "yyyy-mm-dd"), Time, campos(3, 1))
    
    End If

Rem CALCULA HARINA
    ilas = CDbl(Grid1.Cell(LINEA, 12).text)
    If ilas <> 0 Then
    lin = lin + 1
    pivote.MaxLength = 3
    pivote.text = lin
    
    Call ceros(pivote)
    campos(0, 1) = TIPOCON
    campos(1, 1) = Grid1.Cell(LINEA, 2).text
    campos(2, 1) = pivote.text
    campos(3, 1) = Mid(Grid1.Cell(LINEA, 3).text, 1, 9) + Mid(Grid1.Cell(LINEA, 3).text, 11, 1)
    campos(4, 1) = leerdatoslocal(localfiltro, "cuentaharina")
    campos(5, 1) = "O/C " + ORDEN + " IMPUESTO HARINAS"
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
    Call grabarcomprobante_lineas(tipodoc, campos(1, 1), campos(2, 1), fecha, campos(4, 1), "", "", "", campos(5, 1), tipodoc, campos(1, 1), fecha, fecha, campos(6, 1), DH, USUARIOSISTEMA, MES, año, Format(fechasistema, "yyyy-mm-dd"), Time, campos(3, 1))
    
    End If

Rem CALCULA carne
    ilas = CDbl(Grid1.Cell(LINEA, 13).text)
    If ilas <> 0 Then
    lin = lin + 1
    pivote.MaxLength = 3
    pivote.text = lin
    
    Call ceros(pivote)
    campos(0, 1) = TIPOCON
    campos(1, 1) = Grid1.Cell(LINEA, 2).text
    campos(2, 1) = pivote.text
    campos(3, 1) = Mid(Grid1.Cell(LINEA, 3).text, 1, 9) + Mid(Grid1.Cell(LINEA, 3).text, 11, 1)
    campos(4, 1) = leerdatoslocal(localfiltro, "cuentacarne")
    campos(5, 1) = "O/C " + ORDEN + " IMPUESTO CARNE"
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
    Call grabarcomprobante_lineas(tipodoc, campos(1, 1), campos(2, 1), fecha, campos(4, 1), "", "", "", campos(5, 1), tipodoc, campos(1, 1), fecha, fecha, campos(6, 1), DH, USUARIOSISTEMA, MES, año, Format(fechasistema, "yyyy-mm-dd"), Time, campos(3, 1))
    
    End If
    
   
    
    
End Sub

Public Function leefactura(tipo, numero, rut) As String

    
    campos(0, 0) = "tipo"
    campos(1, 0) = "numero"
    campos(2, 0) = ""
    If tipo = "FA" Then tipo = "1"
    If tipo = "ND" Then tipo = "2"
    If tipo = "NC" Then tipo = "3"
    If tipo = "FAE" Then tipo = "4"
    If tipo = "NDE" Then tipo = "5"
    If tipo = "NCE" Then tipo = "6"
    If tipo = "FEE" Then tipo = "0"
    If tipo = "FE" Then tipo = "9"
    
    condicion = "tipo='" + tipo + "' and numero='" + numero + "' and rut='" + rut + "' "
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
            csql.sql = csql.sql & "(año,tipo,rut,nombre,direccion,comuna,ciudad,giro,fono) "
            csql.sql = csql.sql & "SELECT '" + año + "','" + cuentacliente + "',mc.rut,mc.nombre,mc.direccion,mc.comuna,mc.ciudad,mc.giro,mc.fono1 "
            csql.sql = csql.sql & "FROM " & clientesistema & "ventas.sv_maestroclientes as mc "
            csql.sql = csql.sql & "WHERE mc.rut = '" & rut & "' AND mc.sucursal ='0'"
            
            csql.Execute
            Call sincronizadatos(csql.sql, gestion, "")
            
            
            csql.sql = "INSERT INTO " + clientesistema + "conta" + empresaactiva + ".saldosctacte "
            csql.sql = csql.sql & "(año,tipo,rut) "
            csql.sql = csql.sql & "SELECT '" + año + "','" + cuentacliente + "',mc.rut "
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

Public Function LEERULTIMOFOLIO(mesconta, añoconta) As String

    Dim resultados As rdoResultset
    Dim csql As New rdoQuery
    
        Set csql.ActiveConnection = contadb

            csql.sql = "select max(folio) from facturasdecompras where mescontable = '" & Format(mesconta, "00") & "' AND añocontable = '" & añoconta & "' "
            
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
Public Function LEERMONTOIMPUESTO(tipo, numero, ORDEN, cuenta) As Double

    Dim resultados As rdoResultset
    Dim csql As New rdoQuery
 
        Set csql.ActiveConnection = gestionrubro

            csql.sql = "select monto from l_ordendecompra_impuestos_" + localfiltro + " where cuenta = '" & cuenta & "' and tipo='" + tipo + "' and numero='" + numero + "' and numeroorden='" + ORDEN + "' "
            
            csql.Execute
    LEERMONTOIMPUESTO = 0
    If csql.RowsAffected > 0 Then
    
    Set resultados = csql.OpenResultset
    LEERMONTOIMPUESTO = resultados(0)
    
    End If
    
End Function
Sub grabarcomprobante_lineas(tipo, numero, LINEA, fecha, codigocuenta, tipoctacte, rutctacte, centrocosto, glosacontable, tipodocumento, numerodocumento, fechadocumento, fechavencimiento, monto, DH, creadopor, MES, año, fechacreacion, horacreacion, rutproveedor)
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
    campos(17, 0) = "año"
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
    campos(17, 1) = año
    
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

Private Sub ORDEN_GotFocus()
Call cargatexto(ORDEN)
End Sub

Private Sub ORDEN_KeyPress(KeyAscii As Integer)
KeyAscii = esNumero(KeyAscii)
If KeyAscii = 13 Then
Call ceros(ORDEN)
Call BUSCAR_Click


End If

End Sub
Private Sub botonmisaccesos_Click()
programafiltro = Me.Caption
misaccesos.Show
End Sub


Private Sub botonmisfavoritos_Click()
Call AgregaFavoritos(USUARIOSISTEMA, App.EXEName, Me.Name, Me.Caption)
End Sub
