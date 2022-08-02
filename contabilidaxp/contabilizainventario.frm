VERSION 5.00
Object = "{4777436C-EB8C-4596-98A8-EBCF98683969}#1.0#0"; "FlexCell.ocx"
Object = "{ADBBAED6-B16F-40DC-80DF-B44910CBA76C}#1.0#0"; "Frame-Xp.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form contabilizainventario 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Centralizacion de inventarios"
   ClientHeight    =   8715
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   14940
   DrawStyle       =   5  'Transparent
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   Moveable        =   0   'False
   ScaleHeight     =   581
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   996
   Begin XPFrame.FrameXp FrameQuickMenu 
      Height          =   615
      Left            =   11760
      TabIndex        =   21
      Top             =   45
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
         TabIndex        =   23
         Top             =   280
         Width           =   1455
      End
      Begin VB.CommandButton botonmisfavoritos 
         Caption         =   "Mis Favoritos"
         Height          =   255
         Left            =   1680
         TabIndex        =   22
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
   Begin XPFrame.FrameXp FrameXp1 
      Height          =   8730
      Left            =   45
      TabIndex        =   1
      Top             =   45
      Width           =   14865
      _ExtentX        =   26220
      _ExtentY        =   15399
      BackColor       =   16744576
      Caption         =   "CENTRALIZACION DE  INVENTARIOS"
      CaptionEstilo3D =   1
      BackColor       =   16744576
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
      Begin VB.CommandButton Command3 
         BackColor       =   &H0080FF80&
         Caption         =   "IMPRIME GRILLA2"
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
         Left            =   7680
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   8190
         Width           =   1815
      End
      Begin VB.CommandButton BTELIMINA 
         BackColor       =   &H000000FF&
         Caption         =   "ELIMINA COMPROBANTES"
         Height          =   330
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   8280
         Width           =   2535
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H0080FF80&
         Caption         =   "IMPRIME GRILLA1"
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
         Left            =   2925
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   8190
         Width           =   1815
      End
      Begin XPFrame.FrameXp FrameXp3 
         Height          =   1050
         Left            =   90
         TabIndex        =   4
         Top             =   360
         Width           =   14640
         _ExtentX        =   25823
         _ExtentY        =   1852
         BackColor       =   16744576
         Caption         =   "DATOS DE FILTRADO"
         CaptionEstilo3D =   1
         BackColor       =   16744576
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Begin VB.CheckBox Check1 
            BackColor       =   &H00FF8080&
            Caption         =   "Click Todos"
            Height          =   240
            Left            =   13050
            TabIndex        =   13
            Top             =   765
            Visible         =   0   'False
            Width           =   1410
         End
         Begin VB.CommandButton Command2 
            Caption         =   "LISTAR"
            Height          =   285
            Left            =   12960
            TabIndex        =   6
            Top             =   360
            Width           =   1455
         End
         Begin XPFrame.FrameXp FrameXp6 
            Height          =   675
            Left            =   90
            TabIndex        =   8
            Top             =   270
            Width           =   3300
            _ExtentX        =   5821
            _ExtentY        =   1191
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
               Height          =   315
               Left            =   45
               TabIndex        =   9
               Top             =   270
               Width           =   3180
            End
         End
         Begin XPFrame.FrameXp FrameXp7 
            Height          =   675
            Left            =   3510
            TabIndex        =   10
            Top             =   270
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
               TabIndex        =   11
               Top             =   270
               Width           =   2865
            End
         End
         Begin XPFrame.FrameXp FrameXp4 
            Height          =   675
            Left            =   6660
            TabIndex        =   14
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
               TabIndex        =   15
               Top             =   270
               Width           =   4395
            End
         End
      End
      Begin XPFrame.FrameXp FrameXp2 
         Height          =   6675
         Left            =   135
         TabIndex        =   2
         Top             =   1485
         Width           =   14685
         _ExtentX        =   25903
         _ExtentY        =   11774
         BackColor       =   16744576
         Caption         =   "COMPRAS A CENTRALIZAR"
         CaptionEstilo3D =   1
         BackColor       =   16744576
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Begin FlexCell.Grid Grid2 
            Height          =   6000
            Left            =   3960
            TabIndex        =   16
            Top             =   270
            Width           =   10590
            _ExtentX        =   18680
            _ExtentY        =   10583
            Cols            =   5
            DefaultFontSize =   8.25
            Rows            =   30
         End
         Begin MSComctlLib.ProgressBar BARRA 
            Height          =   285
            Left            =   90
            TabIndex        =   12
            Top             =   6300
            Width           =   14460
            _ExtentX        =   25506
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   1
         End
         Begin FlexCell.Grid Grid1 
            Height          =   5145
            Left            =   45
            TabIndex        =   3
            Top             =   270
            Width           =   3840
            _ExtentX        =   6773
            _ExtentY        =   9075
            BackColorFixed  =   16744576
            Cols            =   5
            DefaultFontSize =   8.25
            GridColor       =   16711680
            Rows            =   30
            DateFormat      =   2
         End
         Begin VB.Label LBLCOMPRAS 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00000000&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   465
            Left            =   180
            TabIndex        =   18
            Top             =   5715
            Width           =   3435
         End
      End
      Begin VB.CommandButton Command4 
         BackColor       =   &H0080FF80&
         Caption         =   "TRASPASA CONTABILIDAD"
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
         Left            =   4980
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   8190
         Width           =   2535
      End
      Begin VB.Label PROCESO 
         BackColor       =   &H80000007&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   375
         Left            =   10080
         TabIndex        =   17
         Top             =   8190
         Width           =   4335
      End
   End
End
Attribute VB_Name = "contabilizainventario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private codigoenconta As String

Private localfiltro As String
Private cuentas(20) As String
Private lineascomprobante As Double
Private TIPODO As String
Private fecha2 As String
Private TIPODO2 As String
Private numero As String
Private dias As String

Private cuenta As String
Private CRCC As String
Private monto As Double
Private DH As String
Private glosa As String
Private numerorut As String

 

Private Sub BTELIMINA_Click()

sqlconta.audit = True
sqlconta.programaactivo = Me.Caption

año = COMBOAÑO.text
MES = Format(COMBOMES.ListIndex + 1, "00")


If estacerrado(año + "-" + MES + "-" + Format(fechasistema, "dd")) <> True Then

If Verifica_Permiso(ingreso01.Caption, "elimina") = True Then

Call eliminacomprobantesmasivos("MI", MES, año, empresaactiva)
Else
MsgBox mensaje_nopermiso
End If
COMMAND2_Click
Else
MsgBox "mes ya cerrado"
End If

End Sub

Private Sub Check1_Click()

For k = 1 To Grid1.Rows - 1
Grid1.Cell(k, 14).text = Check1.Value

Next k

End Sub

Private Sub Command1_Click()
imprimir
End Sub



Private Sub COMMAND2_Click()
localfiltro = Mid(ComboLOCAL.text, 1, 2)
año = COMBOAÑO.text
MES = COMBOMES.ListIndex + 1
Rem Call Conectar(servidor, clientesistema + "conta" + localfiltro, usuario, password)
Call Conectargestion(Servidor, clientesistema + "gestion", Usuario, password)
Call Conectargestionrubro(Servidor, clientesistema + "gestion" + leerdatoslocal(localfiltro, "rubro"), Usuario, password)

LEERcompras
CRCC = leerdatoslocal(localfiltro, "codigocrcc")

If CDbl(MES) = 2 Then dias = "28" Else dias = "30"


'ariel agrega que lea el codienconta antes de leer dodumentos, codigoenconta venia vacio
codigoenconta = leerdatoslocal(localfiltro, "codigocontable")

If LEEDOCUMENTOMES("MI", año + "-" + MES + "-" + dias, CRCC) = True Then
PROCESO.Caption = "CONTABILIZADO"
Else
PROCESO.Caption = "SIN CONTABILIZAR "
End If

             

End Sub



Private Sub Command3_Click()
IMPRIMIR2

End Sub


Private Sub Command4_Click()

sqlconta.audit = True
sqlconta.programaactivo = Me.Caption

Dim CRCC As String
Dim rut As String
Dim dias As String

año = COMBOAÑO.text
MES = Format(COMBOMES.ListIndex + 1, "00")
codigoenconta = leerdatoslocal(localfiltro, "codigocontable")
If estacerrado(año + "-" + MES + "-" + Format(fechasistema, "dd")) <> True Then
    If PROCESO.Caption = "SIN CONTABILIZAR " Then
        numero = LEERULTIMOFOLIO("MI")
        MES = Format(MES, "00")
        lineascomprobante = 0
        For k = 1 To Grid1.Rows - 1
            If MES = "02" Then dias = "28" Else dias = "30"
            Call grabarinventario(Format(año + "-" + MES + "-" + dias, "yyyy-mm-dd"), Grid1.Cell(k, 1).text, Grid1.Cell(k, 2).text, Grid1.Cell(k, 3).text, Grid1.Cell(k, 4).text)
        Next k
        For k = 1 To Grid2.Rows - 1
            lineascomprobante = lineascomprobante + 1
            If Grid2.Cell(k, 1).text <> "OC" Then
            CRCC = ""
            If leerdatos(contadb, "cuentasdelmayor", "crcc", "codigo='" + Grid2.Cell(k, 3).text + "' and año='" + Format(fechasistema, "yyyy") + "'") = "1" Then
                CRCC = leerdatoslocal(localfiltro, "codigocrcc")
            End If
            Rem If CRCC = "" Then Stop
            rut = ""
            If leertiene(Grid2.Cell(k, 3).text, "1") = True Then
                rut = "0000000019"
                End If
                If MES = "02" Then dias = "28" Else dias = "30"
                Call grabarcomprobante_lineas("MI", numero, lineascomprobante, año + "-" + MES + "-" + dias, Grid2.Cell(k, 3).text, "", rut, CRCC, "CENTRALIZACION " + Grid2.Cell(k, 2).text, "MI", numero, Format(año + "-" + MES + "-" + dias, "YYYY-MM-DD"), Format(año + "-" + MES + "-" + dias, "YYYY-MM-DD"), Grid2.Cell(k, 4).text, Grid2.Cell(k, 5).text, USUARIOSISTEMA, Format(año + "-" + MES + "-" + dias, "MM"), Format(año + "-" + MES + "-" + dias, "YYYY"), Date, Time, "")
            End If

        Next k
    Else
        MsgBox ("MES YA ESTA CONTABILIZADO")
    End If
    COMMAND2_Click
Else
    MsgBox "MES YA CERRADO "
End If

End Sub

Private Sub Form_Activate()

sqlconta.audit = True
sqlconta.programaactivo = Me.Caption

End Sub

Private Sub Form_Load()
CENTRAR Me
 
   
    Call Conectargestion(Servidor, clientesistema + "gestion", Usuario, password)
    sc = 0
CARGAGRILLA



For k = 1 To 12
COMBOMES.AddItem MonthName(k)
Next k
COMBOMES.ListIndex = CDbl(Format(fechasistema, "mm") - 1)
For k = 2000 To Val(Format(fechasistema, "yyyy"))
COMBOAÑO.AddItem k
Next k
COMBOAÑO.ListIndex = k - 2001
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
titulo = "LISTADO DE CENTRALIZACIONES " + COMBOMES.text + " " + COMBOAÑO.text
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

Sub IMPRIMIR2()
Dim titulo As String
titulo = "LISTADO DE CENTRALIZACIONES " + COMBOMES.text + " " + COMBOAÑO.text
Call CABEZAS2(titulo, "N", "000000000")
Grid2.Range(0, 1, 0, Grid2.Cols - 1).Borders(cellEdgeBottom) = cellThick
Grid2.Range(0, 1, 0, Grid2.Cols - 1).Borders(cellEdgeLeft) = cellThick
Grid2.Range(0, 1, 0, Grid2.Cols - 1).Borders(cellEdgeTop) = cellThick
Grid2.Range(0, 1, 0, Grid2.Cols - 1).Borders(cellEdgeRight) = cellThick
Grid2.Range(0, 1, 0, Grid2.Cols - 1).Borders(cellInsideHorizontal) = cellThick
Grid2.Range(0, 1, 0, Grid2.Cols - 1).Borders(cellInsideVertical) = cellThick
Grid2.DefaultFont.Size = 8
Grid2.PageSetup.Orientation = cellLandscape

Grid2.PageSetup.PrintFixedRow = True
Grid2.PageSetup.BottomMargin = 2
Grid2.PageSetup.TopMargin = 1
Grid2.PageSetup.LeftMargin = 1
Grid2.PageSetup.RightMargin = 0
Grid2.PageSetup.BlackAndWhite = True
Grid2.PageSetup.PrintGridlines = False
Grid2.PrintPreview 100

   
End Sub


Sub grilla()
    
End Sub




Private Sub opciones_GotFocus()



End Sub
Sub CARGAGRILLA()
Rem DATOS DE LA COLUMNA
    Dim FORMATOGRILLA(10, 20)
    Grid1.DefaultFont.Size = 8
       
    FORMATOGRILLA(1, 1) = "FECHA"
    FORMATOGRILLA(1, 2) = "RUT"
    FORMATOGRILLA(1, 3) = "ORDEN"
    FORMATOGRILLA(1, 4) = "COMPRA"
    FORMATOGRILLA(1, 5) = "CONTABILIZAR"
    FORMATOGRILLA(1, 6) = "CONTABILIZADA"
    
    Rem LARGO DE LOS DATOS
    FORMATOGRILLA(2, 1) = "10"
    FORMATOGRILLA(2, 2) = "10"
    FORMATOGRILLA(2, 3) = "10"
    FORMATOGRILLA(2, 4) = "10"
    FORMATOGRILLA(2, 5) = "10"
    FORMATOGRILLA(2, 6) = "10"
    
    Rem TIPO DE DATOS C=CEROS,S=STRING,N=NUMERICO,D=DATE,
    FORMATOGRILLA(3, 1) = "D"
    FORMATOGRILLA(3, 2) = "N"
    FORMATOGRILLA(3, 3) = "N"
    FORMATOGRILLA(3, 4) = "N"
    FORMATOGRILLA(3, 5) = "S"
    FORMATOGRILLA(3, 6) = "S"
    
    Rem FORMATO GRILLA
    FORMATOGRILLA(4, 2) = "##,###,##0"
    FORMATOGRILLA(4, 3) = "##,###,##0"
    FORMATOGRILLA(4, 4) = "##,###,##0"
    
    Rem LOCCKED
    For k = 1 To 4
    FORMATOGRILLA(5, k) = "TRUE"
    Next k
    
    Grid1.Cols = 5
    Grid1.Rows = 1
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
  
    
    
End Sub
Sub CARGAGRILLA2()
Rem DATOS DE LA COLUMNA
    Dim FORMATOGRILLA(10, 20)
    Grid2.DefaultFont.Size = 8
       
    FORMATOGRILLA(1, 1) = "TIPO"
    FORMATOGRILLA(1, 2) = "GLOSA"
    FORMATOGRILLA(1, 3) = "CUENTA"
    FORMATOGRILLA(1, 4) = "MONTO"
    FORMATOGRILLA(1, 5) = "D/H"
    FORMATOGRILLA(1, 6) = "ORIGINAL NETO"
    FORMATOGRILLA(1, 7) = "TOTAL"
    
    Rem LARGO DE LOS DATOS
    FORMATOGRILLA(2, 1) = "10"
    FORMATOGRILLA(2, 2) = "20"
    FORMATOGRILLA(2, 3) = "10"
    FORMATOGRILLA(2, 4) = "10"
    FORMATOGRILLA(2, 5) = "10"
    FORMATOGRILLA(2, 6) = "10"
    FORMATOGRILLA(2, 7) = "10"
    
    Rem TIPO DE DATOS C=CEROS,S=STRING,N=NUMERICO,D=DATE,
    FORMATOGRILLA(3, 1) = "S"
    FORMATOGRILLA(3, 2) = "S"
    FORMATOGRILLA(3, 3) = "S"
    FORMATOGRILLA(3, 4) = "N"
    FORMATOGRILLA(3, 5) = "S"
    FORMATOGRILLA(3, 6) = "N"
    FORMATOGRILLA(3, 7) = "N"
    
    Rem FORMATO GRILLA
    FORMATOGRILLA(4, 4) = "##,###,###,##0"
    FORMATOGRILLA(4, 6) = "##,###,###,##0"
    FORMATOGRILLA(4, 7) = "     %###0.00"
    
    Rem LOCCKED
    For k = 1 To 7
    FORMATOGRILLA(5, k) = "TRUE"
    Next k
    
    Grid2.Cols = 8
    Grid2.Rows = 1
    Grid2.AllowUserResizing = False
    Grid2.DisplayFocusRect = False
    Grid2.ExtendLastCol = True
    Grid2.BoldFixedCell = False
    Grid2.DrawMode = cellOwnerDraw
    
    Grid2.Appearance = Flat
    Grid2.ScrollBarStyle = Flat
    Grid2.FixedRowColStyle = Flat
    
'   GRID2.BackColorFixed = RGB(90, 158, 214)
'   GRID2.BackColorFixedSel = RGB(110, 180, 230)
'   GRID2.BackColorBkg = RGB(90, 158, 214)
'   GRID2.BackColorScrollBar = RGB(231, 235, 247)
'   GRID2.BackColor1 = RGB(231, 235, 247)
'   GRID2.BackColor2 = RGB(239, 243, 255)
'   GRID2.GridColor = RGB(148, 190, 231)
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
  
   
    
    
End Sub



Private Sub monto_Click()
End Sub

Private Sub LEERcompras()

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
    Dim tipo As String
    Dim DH As String
    Dim CRCC As String
    Dim rutcompara As String
    
    Call CARGAGRILLA
    Call CARGAGRILLA2
    
    LINEA = 0: fec = 0: fec1 = 0
    fecha1 = año + "-" + MES + "-" + "01"
    fecha2 = año + "-" + MES + "-" + "31"
    rutcompara = "0" + Mid(rutempresa, 1, 8) + Mid(rutempresa, 10, 1)
    
        Set csql.ActiveConnection = gestionrubro
        csql.sql = "SELECT fecha,rut,numero,SUM(total)/1.19 "
        csql.sql = csql.sql + "from l_movimientos_detalle_" + localfiltro + " "
        csql.sql = csql.sql + "where fecha>='" + fecha1 + "' AND fecha<='" + fecha2 + "' and tipo='OC' and rut<>'" + rutcompara + "' GROUP BY TIPO,NUMERO"
        csql.Execute
        total = 0
        total2 = 0
        Grid1.Rows = 1
        Grid1.AutoRedraw = False
        barra.Max = csql.RowsAffected + 1
        
        barra.Value = 0
        
        If csql.RowsAffected > 0 Then
        Set resultados = csql.OpenResultset
        
         While Not resultados.EOF
             barra.Value = barra.Value + 1
             Grid1.Rows = Grid1.Rows + 1
             LINEA = LINEA + 1
             
             Grid1.Cell(LINEA, 1).text = resultados(0)
             Grid1.Cell(LINEA, 2).text = resultados(1)
             Grid1.Cell(LINEA, 3).text = resultados(2)
             Grid1.Cell(LINEA, 4).text = resultados(3)
            
            total2 = total2 + resultados(3)
            resultados.MoveNext
       
            Wend
End If
      LBLCOMPRAS.Caption = Format(total2, "###,###,###,###")
      Grid1.AutoRedraw = True
   Grid1.Refresh
   
      leerresumen
      
End Sub
   
Private Sub leerresumen()

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
    Dim tipo As String
    Dim DH As String
    Dim CRCC As String
    Dim margen As Double
    Dim destino As String
    Dim rutcompara As String
    Dim costo As Double
    Dim totalfacturas As Double
    Dim totalboletas As Double
    Dim totalnc As Double
    Dim rubroloc As String
    
    
    destino = "20"
    LINEA = 0: fec = 0: fec1 = 0
    rutcompara = "0" + Mid(rutempresa, 1, 8) + Mid(rutempresa, 10, 1)
    
    fecha1 = año + "-" + MES + "-" + "01"
    fecha2 = año + "-" + MES + "-" + "31"
    Set csql.ActiveConnection = gestionrubro
    rubroloc = leerrubrocomercio(localfiltro)
     
     
      If fechasistema < "2018-01-01" Then
      
        csql.sql = "SELECT tipo,sum(total)/1.19,sum(costoventa)/1.19  "
        csql.sql = csql.sql + "from l_movimientos_detalle_" + localfiltro + " "
        csql.sql = csql.sql + "where fecha>='" + fecha1 + "' AND fecha<='" + fecha2 + "' and tipo<>'EL' AND tipo<>'RL'  and tipo<>'IB' AND TIPO<>'GV' and tipo<>'OC' AND tipo<>'AP' AND tipo<>'VP' and tipo<>'EG' GROUP BY TIPO "
        csql.sql = csql.sql + "union "
        csql.sql = csql.sql + "SELECT lmd.tipo,sum(lmd.total)/1.19,sum(lmd.costoventa)/1.19  "
        csql.sql = csql.sql + "from l_movimientos_detalle_" + localfiltro + " as lmd,l_movimientos_cabeza_" + localfiltro + " as lmc," + clientesistema + "gestion.g_maestroempresas as me "
        csql.sql = csql.sql + " where me.codigo=lmc.localdestino and lmc.tipo=lmd.tipo and lmd.fecha=lmc.fecha and lmd.numero=lmc.numero "
        csql.sql = csql.sql + "and me.codigocontable<>'" + empresaactiva + "' and lmd.fecha>='" + fecha1 + "' AND lmd.fecha<='" + fecha2 + "' and lmd.tipo='EL'  GROUP BY lmd.TIPO "
        csql.sql = csql.sql + "union "
        csql.sql = csql.sql + "SELECT lmd.tipo,sum(lmd.total)/1.19,sum(lmd.costoventa)/1.19  "
        csql.sql = csql.sql + "from l_movimientos_detalle_" + localfiltro + " as lmd,l_movimientos_cabeza_" + localfiltro + " as lmc," + clientesistema + "gestion.g_maestroempresas as me "
        csql.sql = csql.sql + " where me.codigo=lmc.localorigen and lmc.tipo=lmd.tipo and lmd.fecha=lmc.fecha and lmd.numero=lmc.numero "
        csql.sql = csql.sql + "and me.codigocontable<>'" + empresaactiva + "' and lmd.fecha>='" + fecha1 + "' AND lmd.fecha<='" + fecha2 + "' and lmd.tipo='RL'  GROUP BY lmd.TIPO "
        csql.sql = csql.sql + "union "
        csql.sql = csql.sql + "SELECT tipo,sum(total)/1.19,sum(costoventa)/1.19  "
        csql.sql = csql.sql + "from l_movimientos_detalle_" + localfiltro + " "
        csql.sql = csql.sql + "where fecha>='" + fecha1 + "' AND fecha<='" + fecha2 + "' and tipo='OC' and rut<>'" + rutcompara + "'  GROUP BY TIPO "
      Else
             csql.sql = "SELECT tipo,sum(total)/1.19,sum(costoventa)/1.19  "
            csql.sql = csql.sql + "from l_movimientos_detalle_" + localfiltro + " "
            csql.sql = csql.sql + "where fecha>='" + fecha1 + "' AND fecha<='" + fecha2 + "' and tipo<>'EL' AND tipo<>'RL'  and tipo<>'IB' and tipo<>'EF' and tipo<>'FV'  AND TIPO<>'GV' and tipo<>'OC' AND tipo<>'AP' AND tipo<>'VP' and tipo<>'EG' GROUP BY TIPO "
            csql.sql = csql.sql + "union "
            
            csql.sql = csql.sql & "SELECT tipo,ROUND(SUM(lmd.total/(1.19+(IF(mi.codigo<>'00000',mi.porcentaje/100,0))))) AS pc,ROUND(SUM(lmd.costoventa/(1.19+(IF(mi.codigo<>'00000',mi.porcentaje/100,0))))) AS cc "
            csql.sql = csql.sql + "from l_movimientos_detalle_" + localfiltro + " AS lmd LEFT JOIN " & clientesistema & "gestion" & rubroloc & ".r_maestroproductos_fijo_" & rubroloc & " AS mpf "
            csql.sql = csql.sql & "ON lmd.codigo=mpf.codigobarra INNER JOIN " & clientesistema & "gestion.g_maestroimpuestos AS mi ON mpf.codigoimpuesto=mi.codigo "
            csql.sql = csql.sql + "where fecha>='" + fecha1 + "' AND fecha<='" + fecha2 + "' and (tipo='EF' or tipo='FV')  GROUP BY TIPO "
            
            csql.sql = csql.sql + "union "
            csql.sql = csql.sql + "SELECT lmd.tipo,sum(lmd.total)/1.19,sum(lmd.costoventa)/1.19  "
            csql.sql = csql.sql + "from l_movimientos_detalle_" + localfiltro + " as lmd,l_movimientos_cabeza_" + localfiltro + " as lmc," + clientesistema + "gestion.g_maestroempresas as me "
            csql.sql = csql.sql + " where me.codigo=lmc.localdestino and lmc.tipo=lmd.tipo and lmd.fecha=lmc.fecha and lmd.numero=lmc.numero "
            csql.sql = csql.sql + "and me.codigocontable<>'" + empresaactiva + "' and lmd.fecha>='" + fecha1 + "' AND lmd.fecha<='" + fecha2 + "' and lmd.tipo='EL'  GROUP BY lmd.TIPO "
            csql.sql = csql.sql + "union "
            csql.sql = csql.sql + "SELECT lmd.tipo,sum(lmd.total)/1.19,sum(lmd.costoventa)/1.19  "
            csql.sql = csql.sql + "from l_movimientos_detalle_" + localfiltro + " as lmd,l_movimientos_cabeza_" + localfiltro + " as lmc," + clientesistema + "gestion.g_maestroempresas as me "
            csql.sql = csql.sql + " where me.codigo=lmc.localorigen and lmc.tipo=lmd.tipo and lmd.fecha=lmc.fecha and lmd.numero=lmc.numero "
            csql.sql = csql.sql + "and me.codigocontable<>'" + empresaactiva + "' and lmd.fecha>='" + fecha1 + "' AND lmd.fecha<='" + fecha2 + "' and lmd.tipo='RL'  GROUP BY lmd.TIPO "
            csql.sql = csql.sql + "union "
            csql.sql = csql.sql + "SELECT tipo,sum(total)/1.19,sum(costoventa)/1.19  "
            csql.sql = csql.sql + "from l_movimientos_detalle_" + localfiltro + " "
            csql.sql = csql.sql + "where fecha>='" + fecha1 + "' AND fecha<='" + fecha2 + "' and tipo='OC' and rut<>'" + rutcompara + "'  GROUP BY TIPO "
      End If
        
        csql.Execute
        total = 0
        total2 = 0
        Grid2.Rows = 1
        Grid2.AutoRedraw = False
        totalfacturas = 0
        totalboletas = 0
        totalnc = 0
        
        If csql.RowsAffected > 0 Then
        Set resultados = csql.OpenResultset
        
         While Not resultados.EOF
             Call leercuentascomprobante(resultados(0))
             Grid2.Rows = Grid2.Rows + 2
             Grid2.Cell(Grid2.Rows - 2, 1).text = resultados(0)
             Grid2.Cell(Grid2.Rows - 2, 2).text = NOMBREDOCUMENTO
             Grid2.Cell(Grid2.Rows - 1, 1).text = resultados(0)
             Grid2.Cell(Grid2.Rows - 1, 2).text = NOMBREDOCUMENTO
             
             
             Grid2.Cell(Grid2.Rows - 2, 3).text = cuentadebe
             Grid2.Cell(Grid2.Rows - 1, 3).text = cuentahaber
             
             If tienecosto = "S" Then
             Grid2.Cell(Grid2.Rows - 2, 4).text = resultados(2)
             Grid2.Cell(Grid2.Rows - 1, 4).text = resultados(2)
             If resultados(2) <> 0 Then
             margen = ((resultados(1) / resultados(2)) - 1) * 100
             Grid2.Cell(Grid2.Rows - 1, 7).text = margen
             End If
             Else
             Grid2.Cell(Grid2.Rows - 2, 4).text = resultados(1)
             Grid2.Cell(Grid2.Rows - 1, 4).text = resultados(1)
             End If
             Grid2.Cell(Grid2.Rows - 2, 5).text = "D"
             Grid2.Cell(Grid2.Rows - 1, 5).text = "H"
             Grid2.Cell(Grid2.Rows - 1, 6).text = resultados(1)
             Select Case resultados(0)
                Case "FV"
                    totalfacturas = totalfacturas + resultados(2)
                Case "BV"
                    totalboletas = totalboletas + resultados(2)
                Case "NB"
                    totalnc = totalnc + resultados(2)
                Case "NF"
                    totalnc = totalnc + resultados(2)
             End Select
            resultados.MoveNext
       
            Wend
End If
      
      Rem LINEA 2
      
         Grid2.Rows = Grid2.Rows + 2
             Grid2.Cell(Grid2.Rows - 2, 1).text = "PM"
             Grid2.Cell(Grid2.Rows - 2, 2).text = "PROVISION MERMA"
             Grid2.Cell(Grid2.Rows - 1, 1).text = "PM"
             Grid2.Cell(Grid2.Rows - 1, 2).text = "PROVISION MERMA"
             
             
             Grid2.Cell(Grid2.Rows - 2, 3).text = "23400006"
             Grid2.Cell(Grid2.Rows - 1, 3).text = "47100006"
             costo = ((totalfacturas + totalboletas) - totalnc) * (2 / 100)
             
             
             
             
             Grid2.Cell(Grid2.Rows - 2, 4).text = costo
             Grid2.Cell(Grid2.Rows - 1, 4).text = costo
       
             Grid2.Cell(Grid2.Rows - 2, 5).text = "H"
             Grid2.Cell(Grid2.Rows - 1, 5).text = "D"
             Grid2.Cell(Grid2.Rows - 1, 6).text = 0
      
      
      
      Grid2.AutoRedraw = True
      Grid2.Refresh
  

End Sub
   
      
   
Private Sub grabarinventario(fechacomprobante, fecha, rut, ORDEN, monto)

Dim resultados As rdoResultset
    Dim csql As New rdoQuery
    Dim csql2 As New rdoQuery
    Dim dias As String
    
    lineascomprobante = lineascomprobante + 1
    LINEA = Format(lineascomprobante, "000")
    fecha = Format(fecha, "yyyy-mm-dd")
    cuenta = "11550001"
    CRCC = ""
    monto = monto
    glosa = "O/C " + ORDEN
    If MES = "02" Then dias = "28" Else dias = "30"

    Call existerut(Format(año + "-" + MES + "-" + dias, "yyyy"), cuenta, rut)
    
    Call grabarcomprobante_lineas("MI", numero, LINEA, fechacomprobante, cuenta, "", rut, CRCC, glosa, "OC", numero, fecha, fecha, monto, "D", USUARIOSISTEMA, Format(fecha, "MM"), Format(fecha, "yyyy"), Format(año + "-" + MES + "-" + dias, "yyyy-mm-dd"), Time, rut)
    lineascomprobante = lineascomprobante + 1
    LINEA = Format(lineascomprobante, "000")
    
    cuenta = "11350001"
    CRCC = ""
    monto = monto
    glosa = "O/C " + ORDEN
    Call existerut(Format(año + "-" + MES + "-" + dias, "yyyy"), cuenta, rut)
    
    Call grabarcomprobante_lineas("MI", numero, LINEA, fechacomprobante, cuenta, "", rut, CRCC, glosa, "OC", numero, fecha, fecha, monto, "H", USUARIOSISTEMA, Format(fecha, "MM"), Format(fecha, "yyyy"), Format(año + "-" + MES + "-" + dias, "yyyy-mm-dd"), Time, rut)
    
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



Sub leercrcc()
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



Private Sub Grid1_KeyDown(KeyCode As Integer, ByVal Shift As Integer)

'If KeyCode = 46 Then
'Call eliminafactura(Grid1.Cell(Grid1.ActiveCell.Row, 1).text, Grid1.Cell(Grid1.ActiveCell.Row, 2).text)
'End If
'leer
End Sub



Public Function LEEDOCUMENTO(tipo, fecha) As Boolean

    
    campos(0, 0) = "tipo"
    campos(1, 0) = ""
    campos(2, 0) = ""
    
    condicion = "tipo='" + tipo + "' and fecha='" & Format(fecha, "yyyy-mm-dd") & "' and linea='1' "
    campos(0, 2) = clientesistema + "conta" + codigoenconta + ".movimientoscontables"
    op = 5
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    
    Call sqlconta.sqlconta(op, condicion)
    If sqlconta.status = 0 Then
    LEEDOCUMENTO = True
    
    Else
    LEEDOCUMENTO = False
    
    
    End If
End Function

Public Function LEEDOCUMENTOMES(tipo, fecha, CRCC) As Boolean

    
    campos(0, 0) = "tipo"
    campos(1, 0) = ""
    campos(2, 0) = ""
    
    condicion = "tipo='" + tipo + "' and MID(fecha,1,7)='" & Format(fecha, "yyyy-mm") & "' and centrocosto='" + CRCC + "' and (codigocuenta='47100001' or codigocuenta='47100002') limit 0,1"
    campos(0, 2) = clientesistema + "conta" + codigoenconta + ".movimientoscontables"
    op = 5
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    
    Call sqlconta.sqlconta(op, condicion)
    If sqlconta.status = 0 Then
    
    LEEDOCUMENTOMES = True
    
    Else
    LEEDOCUMENTOMES = False
    
    
    End If
    
    

End Function


Public Function LEERULTIMOFOLIO(tipo) As String

    Dim resultados As rdoResultset
    Dim csql As New rdoQuery
    
        Set csql.ActiveConnection = contadb

            csql.sql = "select IFNULL(max(numero),0) from movimientoscontables where mes = '" & Format(MES, "00") & "' AND año = '" & año & "' and tipo='" + tipo + "' "
            
            csql.Execute
    If csql.RowsAffected > 0 Then
    Set resultados = csql.OpenResultset
    
    
        LEERULTIMOFOLIO = Format(resultados(0) + 1, "0000000000")
    End If
    
End Function
Public Function LEERMONTOIMPUESTO(tipo, desde, hasta, cuenta, CRCC) As Double

    Dim resultados As rdoResultset
    Dim csql As New rdoQuery
'    cSql.sql = cSql.sql + "where fecha>='" + fecha1 + "' AND fecha<='" + fecha2 + "' and (caja<'90' or caja='') and crcc='" + CRCC + "'  group by tipo,fecha "
'        cSql.sql = cSql.sql + "Union "
'        cSql.sql = cSql.sql + "SELECT 'BV',fecha,ROUND(sum(total)/1.19),sum(total)-ROUND((sum(total)/1.19)),sum(exento),sum(total) "
'        cSql.sql = cSql.sql + "from boletasdeventa "
'        cSql.sql = cSql.sql + "where fecha>='" + fecha1 + "' AND fecha<='" + fecha2 + "' and centrocosto='" + CRCC + "' group by fecha order by fecha,tipo"
'
        Set csql.ActiveConnection = contadb
            desde = Format(desde, "yyyy-mm-dd")
            hasta = Format(hasta, "yyyy-mm-dd")
            
            csql.sql = "select ifnull(sum(fvd.monto),0) "
            csql.sql = csql.sql + "from facturasdeventas_detalle as fvd ,facturasdeventas as fv "
            csql.sql = csql.sql + "where fvd.tipo=fv.tipo and fvd.numero=fv.numero and cuentadelmayor= '" & cuenta & "' and fecha>='" + desde + "' and fecha<='" + hasta + "' and fvd.tipo='" + tipo + "' and (fv.caja<'90' or fv.caja='') and crcc='" + CRCC + "' "
            
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

    campos(0, 2) = clientesistema + "conta" + codigoenconta + ".movimientoscontables"
   

    op = 2
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    
    Call sqlconta.sqlconta(op, condicion)
   'Call ACTUALIZADOCUMENTO("+")
   
End Sub


Public Function existerut(año, tipo, rut) As Boolean
   
    campos(0, 0) = "año"
    campos(1, 0) = "tipo"
    campos(2, 0) = "rut"
    campos(3, 0) = "nombre"
    campos(4, 0) = ""
    campos(0, 1) = Format(año + "-" + MES + "-" + dias, "yyyy")
    campos(1, 1) = tipo
    campos(2, 1) = rut
    campos(3, 1) = nombreproveedor(rut)
    condicion = "tipo='" + tipo + "' and rut='" + rut + "' "
    campos(0, 2) = clientesistema + "conta" + codigoenconta + ".cuentascorrientes"
    op = 5
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)
    If sqlconta.status = 0 Then
    existerut = True
    Else
    Call grabar(año, tipo, rut, nombreproveedor(rut))
    
    End If

    
    End Function
Public Function nombreproveedor(rut) As String
    campos(0, 0) = "nombre"
    campos(1, 0) = ""
    condicion = "rut='" + rut + "' and tipo='" + CUENTAPROVEEDOR + "'  "
    campos(0, 2) = clientesistema + "conta" + codigoenconta + ".cuentascorrientes"
    op = 5
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)
    If sqlconta.status = 0 Then
    nombreproveedor = sqlconta.response(0, 3)
    Else
    nombreproveedor = ""
    End If
    End Function

Sub grabar(año, tipo, rut, NOMBRE)
    campos(0, 0) = "año"
    campos(1, 0) = "tipo"
    campos(2, 0) = "rut"
    campos(3, 0) = "nombre"
    campos(4, 0) = ""
    campos(0, 1) = Format(año + "-" + MES + "-" + dias, "yyyy")
    campos(1, 1) = tipo
    campos(2, 1) = rut
    campos(3, 1) = NOMBRE
    
    campos(0, 2) = clientesistema + "conta" + codigoenconta + ".cuentascorrientes"
    condicion = ""
    op = 2
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)
    
     Call grabar2(año, tipo, rut)
    
    End Sub
Sub grabar2(año, tipo, rut)
      
    campos(0, 0) = "año"
    campos(1, 0) = "tipo"
    campos(2, 0) = "rut"
    campos(3, 0) = ""
    
    campos(0, 1) = año
    campos(1, 1) = tipo
    campos(2, 1) = rut
    
    campos(0, 2) = clientesistema + "conta" + codigoenconta + ".saldosctacte"
    op = 2
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)
    

End Sub

Private Sub botonmisaccesos_Click()
programafiltro = Me.Caption
misaccesos.Show
End Sub


Private Sub botonmisfavoritos_Click()
Call AgregaFavoritos(USUARIOSISTEMA, App.EXEName, Me.Name, Me.Caption)
End Sub
