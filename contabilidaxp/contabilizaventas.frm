VERSION 5.00
Object = "{4777436C-EB8C-4596-98A8-EBCF98683969}#1.0#0"; "FlexCell.ocx"
Object = "{ADBBAED6-B16F-40DC-80DF-B44910CBA76C}#1.0#0"; "Frame-Xp.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form proceso05 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Centralizacion de Ventas"
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
      Left            =   12000
      TabIndex        =   17
      Top             =   0
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
         TabIndex        =   19
         Top             =   280
         Width           =   1455
      End
      Begin VB.CommandButton botonmisfavoritos 
         Caption         =   "Mis Favoritos"
         Height          =   255
         Left            =   1800
         TabIndex        =   18
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
      Height          =   8610
      Left            =   45
      TabIndex        =   1
      Top             =   45
      Width           =   14865
      _ExtentX        =   26220
      _ExtentY        =   15187
      BackColor       =   16744576
      Caption         =   "CENTRALIZACION DE  VENTAS"
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
      Begin VB.CommandButton BTELIMINA 
         BackColor       =   &H000000FF&
         Caption         =   "ELIMINA COMPROBANTES"
         Height          =   330
         Left            =   480
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   8160
         Width           =   2535
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H0080FF80&
         Caption         =   "IMPRIMIR"
         Height          =   330
         Left            =   3735
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   8190
         Width           =   2130
      End
      Begin VB.CommandButton Command4 
         BackColor       =   &H0080FF80&
         Caption         =   "TRASPASA CONTABILIDAD"
         Height          =   330
         Left            =   6720
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   8190
         Width           =   2535
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
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Begin VB.CommandButton Command2 
            Caption         =   "LISTAR"
            Height          =   285
            Left            =   11745
            TabIndex        =   6
            Top             =   405
            Width           =   1455
         End
         Begin VB.CheckBox Check1 
            BackColor       =   &H00FF8080&
            Caption         =   "Click Todos"
            Height          =   240
            Left            =   13050
            TabIndex        =   5
            Top             =   765
            Width           =   1410
         End
         Begin XPFrame.FrameXp FrameXp6 
            Height          =   675
            Left            =   90
            TabIndex        =   7
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
               TabIndex        =   8
               Top             =   270
               Width           =   3180
            End
         End
         Begin XPFrame.FrameXp FrameXp7 
            Height          =   675
            Left            =   3510
            TabIndex        =   9
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
               TabIndex        =   10
               Top             =   270
               Width           =   2865
            End
         End
         Begin XPFrame.FrameXp FrameXp4 
            Height          =   675
            Left            =   6705
            TabIndex        =   11
            Top             =   270
            Width           =   4560
            _ExtentX        =   8043
            _ExtentY        =   1191
            BackColor       =   16744576
            Caption         =   "CRCC"
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
               TabIndex        =   12
               Top             =   270
               Width           =   4395
            End
         End
      End
      Begin XPFrame.FrameXp FrameXp2 
         Height          =   6675
         Left            =   135
         TabIndex        =   13
         Top             =   1485
         Width           =   14685
         _ExtentX        =   25903
         _ExtentY        =   11774
         BackColor       =   16744576
         Caption         =   "DIAS A CENTRALIZAR"
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
         Begin MSComctlLib.ProgressBar BARRA 
            Height          =   285
            Left            =   180
            TabIndex        =   14
            Top             =   6300
            Width           =   14325
            _ExtentX        =   25268
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   1
         End
         Begin FlexCell.Grid Grid1 
            Height          =   5955
            Left            =   45
            TabIndex        =   15
            Top             =   270
            Width           =   14595
            _ExtentX        =   25744
            _ExtentY        =   10504
            BackColorFixed  =   16744576
            Cols            =   5
            DefaultFontSize =   8.25
            GridColor       =   16711680
            Rows            =   30
            DateFormat      =   2
         End
      End
   End
End
Attribute VB_Name = "proceso05"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private localfiltro As String
Private cuentas(20) As String


Private Sub BTELIMINA_Click()
localfiltro = Mid(ComboLOCAL.text, 1, 2)
año = COMBOAÑO.text
MES = Format(COMBOMES.ListIndex + 1, "00")

If estacerrado(año + "-" + MES + "-" + Format(fechasistema, "dd")) <> True Then
    If Verifica_Permiso(ingreso01.Caption, "elimina") = True Then
        Call eliminacomprobantesmasivos("CV", MES, año, empresaactiva)
    Else
        MsgBox mensaje_nopermiso
    End If
Else
    MsgBox "MES YA CERRADO"
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
Rem Call Conectarventas(servidor, clientesistema + "ventas" + localfiltro, usuario, password)
leer


End Sub



Private Sub Command3_Click()
Dim k As Integer


End Sub


Private Sub Command4_Click()
Dim k As Double
Dim numero As String
Dim fecha As Date
Dim cuenta As String
Dim CRCC As String
Dim monto As Double
Dim DH As String
Dim glosa As String
Dim TIPODO As String
Dim fecha2 As String
Dim TIPODO2 As String

Dim LINEA As String
Dim lin As Double
Dim DH2 As String
TIPODO = "CV"
lin = 0
año = COMBOAÑO.text
MES = Format(COMBOMES.ListIndex + 1, "00")

If estacerrado(año + "-" + MES + "-" + Format(fechasistema, "dd")) <> True Then

For k = 1 To Grid1.Rows - 1
    If Grid1.Cell(k, 14).text = "1" Then
    TIPODO2 = Grid1.Cell(k, 1).text
    If fecha2 <> Format(Grid1.Cell(k, 2).text, "yyyy-mm-dd") Then
    numero = LEERULTIMOFOLIO
    lin = 0
    fecha2 = Format(Grid1.Cell(k, 2).text, "yyyy-mm-dd")
    End If
    
Rem RENDICION
    If Grid1.Cell(k, 1).text = "VC" Then DH = "H": DH2 = "D"
    
    If TIPODO2 <> "VC" Then
    lin = lin + 1
    LINEA = Format(lin, "000")
    fecha = Format(Grid1.Cell(k, 2).text, "yyyy-mm-dd")
    cuenta = leerdatoslocal(localfiltro, "cuentarendicion")
    CRCC = ""
   
    monto = Replace(Grid1.Cell(k, 6).text, ",", ".")
    
    glosa = "CONTABILIZACION VENTAS " + Grid1.Cell(k, 1).text
    If Grid1.Cell(k, 1).text = "FE" Then DH = "D": DH2 = "H"
    If Grid1.Cell(k, 1).text = "FA" Then DH = "D": DH2 = "H"
    If Grid1.Cell(k, 1).text = "NF" Then DH = "H": DH2 = "D"
    If Grid1.Cell(k, 1).text = "NB" Then DH = "H": DH2 = "D"
    If Grid1.Cell(k, 1).text = "BV" Then DH = "D": DH2 = "H"
    If Grid1.Cell(k, 1).text = "EF" Then DH = "D": DH2 = "H"
    If Grid1.Cell(k, 1).text = "ED" Then DH = "D": DH2 = "H"
    If Grid1.Cell(k, 1).text = "EC" Then DH = "H": DH2 = "D"
    
    
    Call grabarcomprobante_lineas(TIPODO, numero, LINEA, fecha, cuenta, "", "", CRCC, glosa, TIPODO2, numero, fecha, fecha, monto, DH, USUARIOSISTEMA, Format(fecha, "MM"), Format(fecha, "yyyy"), Format(fechasistema, "yyyy-mm-dd"), Time, "")
    End If
Rem IVA
    If CDbl(Grid1.Cell(k, 4).text) <> "0" Then
    lin = lin + 1
    LINEA = Format(lin, "000")
    fecha = Format(Grid1.Cell(k, 2).text, "yyyy-mm-dd")
    cuenta = ivadebito
    CRCC = ""
    monto = Replace(Grid1.Cell(k, 4).text, ",", ".")
    glosa = "CONTABILIZACION IVAS " + Grid1.Cell(k, 1).text
    Call grabarcomprobante_lineas(TIPODO, numero, LINEA, fecha, cuenta, "", "", CRCC, glosa, TIPODO2, numero, fecha, fecha, monto, DH2, USUARIOSISTEMA, Format(fecha, "MM"), Format(fecha, "yyyy"), Format(fechasistema, "yyyy-mm-dd"), Time, "")
    End If
Rem ingresoventa
    lin = lin + 1
    LINEA = Format(lin, "000")
    fecha = Format(Grid1.Cell(k, 2).text, "yyyy-mm-dd")
    cuenta = leerdatoslocal(localfiltro, "cuentaingresoventa")
    CRCC = leerdatoslocal(localfiltro, "codigocrcc")
    monto = Replace(Grid1.Cell(k, 3).text, ",", ".")
    monto = monto + Replace(Grid1.Cell(k, 5).text, ",", ".")
     If TIPODO2 = "VC" Then
         monto = Replace(Grid1.Cell(k, 4).text, ",", ".")
        DH2 = "H"
    End If
    glosa = "CONTABILIZACION VENTAS " + Grid1.Cell(k, 1).text
    Call grabarcomprobante_lineas(TIPODO, numero, LINEA, fecha, cuenta, "", "", CRCC, glosa, TIPODO2, numero, fecha, fecha, monto, DH2, USUARIOSISTEMA, Format(fecha, "MM"), Format(fecha, "yyyy"), Format(fechasistema, "yyyy-mm-dd"), Time, "")
        
Rem ILAREFRESCOS
    If CDbl(Grid1.Cell(k, 7).text) <> 0 Then
    lin = lin + 1
    LINEA = Format(lin, "000")
    fecha = Format(Grid1.Cell(k, 2).text, "yyyy-mm-dd")
    cuenta = "23300010"
    CRCC = ""
    monto = Replace(Grid1.Cell(k, 7).text, ",", ".")
    glosa = "CONTABILIZACION VENTAS " + Grid1.Cell(k, 1).text
    Call grabarcomprobante_lineas(TIPODO, numero, LINEA, fecha, cuenta, "", "", CRCC, glosa, TIPODO2, numero, fecha, fecha, monto, DH2, USUARIOSISTEMA, Format(fecha, "MM"), Format(fecha, "yyyy"), Format(fechasistema, "yyyy-mm-dd"), Time, "")
    End If
Rem ILAvinos
    If CDbl(Grid1.Cell(k, 8).text) <> 0 Then
    lin = lin + 1
    LINEA = Format(lin, "000")
    fecha = Format(Grid1.Cell(k, 2).text, "yyyy-mm-dd")
    cuenta = "23300011"
    CRCC = ""
    monto = Replace(Grid1.Cell(k, 8).text, ",", ".")
    glosa = "CONTABILIZACION VENTAS " + Grid1.Cell(k, 1).text
    Call grabarcomprobante_lineas(TIPODO, numero, LINEA, fecha, cuenta, "", "", CRCC, glosa, TIPODO2, numero, fecha, fecha, monto, DH2, USUARIOSISTEMA, Format(fecha, "MM"), Format(fecha, "yyyy"), Format(fechasistema, "yyyy-mm-dd"), Time, "")
    End If
Rem ILALICORES
    If CDbl(Grid1.Cell(k, 9).text) <> 0 Then
    lin = lin + 1
    LINEA = Format(lin, "000")
    fecha = Format(Grid1.Cell(k, 2).text, "yyyy-mm-dd")
    cuenta = "23300013"
    CRCC = ""
    monto = Replace(Grid1.Cell(k, 9).text, ",", ".")
    glosa = "CONTABILIZACION VENTAS " + Grid1.Cell(k, 1).text
    Call grabarcomprobante_lineas(TIPODO, numero, LINEA, fecha, cuenta, "", "", CRCC, glosa, TIPODO2, numero, fecha, fecha, monto, DH2, USUARIOSISTEMA, Format(fecha, "MM"), Format(fecha, "yyyy"), Format(fechasistema, "yyyy-mm-dd"), Time, "")
    End If
Rem ILALIGHT
    If CDbl(Grid1.Cell(k, 10).text) <> 0 Then
    lin = lin + 1
    LINEA = Format(lin, "000")
    fecha = Format(Grid1.Cell(k, 2).text, "yyyy-mm-dd")
    cuenta = "23300017"
    CRCC = ""
    monto = Replace(Grid1.Cell(k, 10).text, ",", ".")
    glosa = "CONTABILIZACION VENTAS " + Grid1.Cell(k, 1).text
    Call grabarcomprobante_lineas(TIPODO, numero, LINEA, fecha, cuenta, "", "", CRCC, glosa, TIPODO2, numero, fecha, fecha, monto, DH2, USUARIOSISTEMA, Format(fecha, "MM"), Format(fecha, "yyyy"), Format(fechasistema, "yyyy-mm-dd"), Time, "")
    End If
Rem harinas
    If CDbl(Grid1.Cell(k, 11).text) <> 0 Then
    lin = lin + 1
    LINEA = Format(lin, "000")
    fecha = Format(Grid1.Cell(k, 2).text, "yyyy-mm-dd")
    cuenta = leerdatoslocal(localfiltro, "harinaventa")
    CRCC = ""
    monto = Replace(Grid1.Cell(k, 11).text, ",", ".")
    glosa = "CONTABILIZACION VENTAS " + Grid1.Cell(k, 1).text
    Call grabarcomprobante_lineas(TIPODO, numero, LINEA, fecha, cuenta, "", "", CRCC, glosa, TIPODO2, numero, fecha, fecha, monto, DH2, USUARIOSISTEMA, Format(fecha, "MM"), Format(fecha, "yyyy"), Format(fechasistema, "yyyy-mm-dd"), Time, "")
    End If
Rem carne
    If CDbl(Grid1.Cell(k, 12).text) <> 0 Then
    lin = lin + 1
    LINEA = Format(lin, "000")
    fecha = Format(Grid1.Cell(k, 2).text, "yyyy-mm-dd")
    cuenta = leerdatoslocal(localfiltro, "carneventas")
    CRCC = ""
    monto = Replace(Grid1.Cell(k, 12).text, ",", ".")
    glosa = "CONTABILIZACION VENTAS " + Grid1.Cell(k, 1).text
    Call grabarcomprobante_lineas(TIPODO, numero, LINEA, fecha, cuenta, "", "", CRCC, glosa, TIPODO2, numero, fecha, fecha, monto, DH2, USUARIOSISTEMA, Format(fecha, "MM"), Format(fecha, "yyyy"), Format(fechasistema, "yyyy-mm-dd"), Time, "")
    End If
Rem ILA CERVEZA
    If CDbl(Grid1.Cell(k, 13).text) <> 0 Then
    lin = lin + 1
    LINEA = Format(lin, "000")
    fecha = Format(Grid1.Cell(k, 2).text, "yyyy-mm-dd")
    cuenta = "23300014"
    CRCC = ""
    monto = Replace(Grid1.Cell(k, 13).text, ",", ".")
    glosa = "CONTABILIZACION VENTAS " + Grid1.Cell(k, 1).text
    Call grabarcomprobante_lineas(TIPODO, numero, LINEA, fecha, cuenta, "", "", CRCC, glosa, TIPODO2, numero, fecha, fecha, monto, DH2, USUARIOSISTEMA, Format(fecha, "MM"), Format(fecha, "yyyy"), Format(fechasistema, "yyyy-mm-dd"), Time, "")
    End If
    
    End If
Next k
leer
Else
MsgBox "MES YA CERRADO"
End If

End Sub

Private Sub Form_Activate()
sqlconta.audit = True
sqlconta.programaactivo = Me.Caption

End Sub

Private Sub Form_Load()
CENTRAR Me
    Call Conectar_BD
    sc = 0
CARGAGRILLA

Call Conectargestion(Servidor, clientesistema + "gestion", Usuario, password)
Call Conectargestionrubro(Servidor, clientesistema + "gestion00", Usuario, password)

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
Sub grilla()
    
End Sub




Private Sub opciones_GotFocus()



End Sub
Sub CARGAGRILLA()
Rem DATOS DE LA COLUMNA
    Dim FORMATOGRILLA(10, 20)
    Grid1.DefaultFont.Size = 8
       
    FORMATOGRILLA(1, 1) = "TP"
    FORMATOGRILLA(1, 2) = "FECHA"
    FORMATOGRILLA(1, 3) = "NETO"
    FORMATOGRILLA(1, 4) = "IVA"
    FORMATOGRILLA(1, 5) = "EXENTO"
    FORMATOGRILLA(1, 6) = "TOTAL"
    FORMATOGRILLA(1, 7) = "I.REFRESCOS"
    FORMATOGRILLA(1, 8) = "I.VINOS    "
    FORMATOGRILLA(1, 9) = "I.LICORES"
    FORMATOGRILLA(1, 10) = "I.LIGHT"
    FORMATOGRILLA(1, 11) = "I.HARINA   "
    FORMATOGRILLA(1, 12) = "I.CARNES"
    FORMATOGRILLA(1, 13) = "I.CERVEZAS"
    FORMATOGRILLA(1, 14) = "CONTABILIZA"
    
    Rem LARGO DE LOS DATOS
    FORMATOGRILLA(2, 1) = "3"
    FORMATOGRILLA(2, 2) = "8"
    FORMATOGRILLA(2, 3) = "8"
    FORMATOGRILLA(2, 4) = "8"
    FORMATOGRILLA(2, 5) = "8"
    FORMATOGRILLA(2, 6) = "8"
    FORMATOGRILLA(2, 7) = "8"
    FORMATOGRILLA(2, 8) = "8"
    FORMATOGRILLA(2, 9) = "8"
    FORMATOGRILLA(2, 10) = "8"
    FORMATOGRILLA(2, 11) = "8"
    FORMATOGRILLA(2, 12) = "8"
    FORMATOGRILLA(2, 13) = "8"
    FORMATOGRILLA(2, 14) = "5"

    
    Rem TIPO DE DATOS C=CEROS,S=STRING,N=NUMERICO,D=DATE,
    FORMATOGRILLA(3, 1) = "S"
    FORMATOGRILLA(3, 2) = "D"
    FORMATOGRILLA(3, 3) = "N"
    FORMATOGRILLA(3, 4) = "N"
    FORMATOGRILLA(3, 5) = "N"
    FORMATOGRILLA(3, 6) = "N"
    FORMATOGRILLA(3, 7) = "N"
    FORMATOGRILLA(3, 8) = "N"
    FORMATOGRILLA(3, 9) = "N"
    FORMATOGRILLA(3, 10) = "N"
    FORMATOGRILLA(3, 11) = "N"
    FORMATOGRILLA(3, 12) = "N"
    FORMATOGRILLA(3, 13) = "N"
    
    FORMATOGRILLA(3, 14) = "S"
    
    Rem FORMATO GRILLA
    FORMATOGRILLA(4, 3) = "##,###,##0"
    FORMATOGRILLA(4, 4) = "##,###,##0"
    FORMATOGRILLA(4, 5) = "##,###,##0"
    FORMATOGRILLA(4, 6) = "##,###,##0"
    FORMATOGRILLA(4, 7) = "##,###,##0"
    FORMATOGRILLA(4, 8) = "##,###,##0"
    FORMATOGRILLA(4, 9) = "##,###,##0"
    FORMATOGRILLA(4, 10) = "##,###,##0"
    FORMATOGRILLA(4, 11) = "##,###,##0"
    FORMATOGRILLA(4, 12) = "##,###,##0"
    FORMATOGRILLA(4, 13) = "##,###,##0"
    
    
    Rem LOCCKED
    For k = 1 To 13
    FORMATOGRILLA(5, k) = "TRUE"
    Next k
    FORMATOGRILLA(5, 14) = "FALSE"
    
    Grid1.Cols = 15
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
    Dim tipo As String
    Dim DH As String
    Dim CRCC As String
    
    LINEA = 0: fec = 0: fec1 = 0
    fecha1 = año + "-" + MES + "-" + "01"
    fecha2 = año + "-" + MES + "-" + "31"
    CRCC = leerdatoslocal(localfiltro, "codigocrcc")
    
        Set csql.ActiveConnection = contadb
        csql.sql = "SELECT tipo,fecha,sum(neto),ROUND(sum(iva)),'0',ROUND(sum(total)) "
        csql.sql = csql.sql + "from facturasdeventas "
        csql.sql = csql.sql + "where fecha>='" + fecha1 + "' AND fecha<='" + fecha2 + "' and (caja<'90' or caja='') and crcc='" + CRCC + "' AND (neto<>'0' or exento<>'0')  group by tipo,fecha "
        csql.sql = csql.sql + "Union "
        csql.sql = csql.sql + "SELECT if(cigarro='0','BV','VC'),fecha,ROUND(sum(monto)/1.19),sum(monto)-ROUND((sum(monto)/1.19)),sum(exento),sum(total) "
        csql.sql = csql.sql + "from boletasdeventa "
        csql.sql = csql.sql + "where fecha>='" + fecha1 + "' AND fecha<='" + fecha2 + "' and centrocosto='" + CRCC + "' group by fecha,cigarro order by fecha,tipo"
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
             tipo = resultados(0): DH = "D"
             If tipo = "1" Then tipo = "FA": DH = "D"
             If tipo = "2" Then tipo = "ND": DH = "D"
             If tipo = "3" Then tipo = "NB": DH = "H"
             If tipo = "4" Then tipo = "NF": DH = "H"
             If tipo = "5" Then tipo = "FE": DH = "D"
             If tipo = "6" Then tipo = "EF": DH = "D"
             If tipo = "7" Then tipo = "ED": DH = "D"
             If tipo = "8" Then tipo = "EC": DH = "H"
             If tipo = "9" Then tipo = "EX": DH = "D"
             
             If LEEDOCUMENTO(leerdatoslocal(localfiltro, "cuentarendicion"), resultados(1), tipo, resultados(5), DH) = False Then
             Grid1.Rows = Grid1.Rows + 1
             LINEA = LINEA + 1
             
             Grid1.Cell(LINEA, 1).text = tipo
             Grid1.Cell(LINEA, 2).text = resultados(1)
            
             Grid1.Cell(LINEA, 3).text = resultados(2)
             Grid1.Cell(LINEA, 4).text = resultados(3)
             Grid1.Cell(LINEA, 5).text = resultados(4)
             Grid1.Cell(LINEA, 6).text = resultados(5)
             Grid1.Cell(LINEA, 7).text = LEERMONTOIMPUESTO(resultados(0), resultados(1), resultados(1), leerdatoslocal(localfiltro, "23300010"), CRCC)
             Grid1.Cell(LINEA, 8).text = LEERMONTOIMPUESTO(resultados(0), resultados(1), resultados(1), leerdatoslocal(localfiltro, "23300011"), CRCC)
             Grid1.Cell(LINEA, 9).text = LEERMONTOIMPUESTO(resultados(0), resultados(1), resultados(1), leerdatoslocal(localfiltro, "23300013"), CRCC)
             Grid1.Cell(LINEA, 10).text = LEERMONTOIMPUESTO(resultados(0), resultados(1), resultados(1), leerdatoslocal(localfiltro, "23300017"), CRCC)
             Grid1.Cell(LINEA, 11).text = LEERMONTOIMPUESTO(resultados(0), resultados(1), resultados(1), leerdatoslocal(localfiltro, "harinaventa"), CRCC)
             Grid1.Cell(LINEA, 12).text = LEERMONTOIMPUESTO(resultados(0), resultados(1), resultados(1), leerdatoslocal(localfiltro, "carneventas"), CRCC)
             Grid1.Cell(LINEA, 13).text = LEERMONTOIMPUESTO(resultados(0), resultados(1), resultados(1), leerdatoslocal(localfiltro, "cuentailas"), CRCC) + LEERMONTOIMPUESTO(resultados(0), resultados(1), resultados(1), leerdatoslocal(localfiltro, "23300014"), CRCC)
             If tipo <> "BV" And tipo <> "VC" Then
                Grid1.Cell(LINEA, 3).text = (resultados(5) - resultados(3) - CDbl(Grid1.Cell(LINEA, 7).text) - CDbl(Grid1.Cell(LINEA, 8).text) - CDbl(Grid1.Cell(LINEA, 9).text) - CDbl(Grid1.Cell(LINEA, 10).text) - CDbl(Grid1.Cell(LINEA, 11).text) - CDbl(Grid1.Cell(LINEA, 12).text) - CDbl(Grid1.Cell(LINEA, 13).text))
             End If
             If tipo = "VC" Then
                Grid1.Cell(LINEA, 3).text = "0"
                Grid1.Cell(LINEA, 4).text = resultados(4) - Round(resultados(4) / 1.19)
                Grid1.Cell(LINEA, 5).text = "0"
                Grid1.Cell(LINEA, 6).text = "0"
             End If
             
            End If
            resultados.MoveNext
       
            Wend
End If
      Grid1.Column(14).CellType = cellCheckBox
      
      
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



Private Sub Grid1_DblClick()
localorden = localfiltro
Rcompra02.dato1.text = Grid1.Cell(Grid1.ActiveCell.row, 16).text

Rcompra02.Show vbModal


End Sub

Private Sub Grid1_KeyDown(KeyCode As Integer, ByVal Shift As Integer)

'If KeyCode = 46 Then
'Call eliminafactura(Grid1.Cell(Grid1.ActiveCell.Row, 1).text, Grid1.Cell(Grid1.ActiveCell.Row, 2).text)
'End If
'leer
End Sub



Public Function LEEDOCUMENTO(cuenta, fecha, tipo, monto, DH) As Boolean

    
    campos(0, 0) = "tipo"
    campos(1, 0) = ""
    campos(2, 0) = ""
    If tipo <> "VC" Then
        condicion = "codigocuenta='" + cuenta + "' and fecha='" & Format(fecha, "yyyy-mm-dd") & "' and tipodocumento='" + tipo + "' and monto='" & monto & "' and dh='" + DH + "' and tipo='CV' "
    Else
        monto = monto - Round(monto / 1.19)
        condicion = "codigocuenta='23200001' and fecha='" & Format(fecha, "yyyy-mm-dd") & "' and tipodocumento='" + tipo + "' and monto='" & monto & "' and dh='D' and tipo='CV' "

    End If
    campos(0, 2) = "movimientoscontables"
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


Public Function LEERULTIMOFOLIO() As String

    Dim resultados As rdoResultset
    Dim csql As New rdoQuery
    
        Set csql.ActiveConnection = contadb

            csql.sql = "select IFNULL(max(numero),0) from movimientoscontables where mes = '" & Format(MES, "00") & "' AND año = '" & año & "' and tipo='CV' "
            
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

    campos(0, 2) = "movimientoscontables"
   

    op = 2
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    
    Call sqlconta.sqlconta(op, condicion)
   'Call ACTUALIZADOCUMENTO("+")
   
End Sub

Private Sub botonmisaccesos_Click()
    programafiltro = Me.Caption
    misaccesos.Show
End Sub


Private Sub botonmisfavoritos_Click()
Call AgregaFavoritos(USUARIOSISTEMA, App.EXEName, Me.Name, Me.Caption)
End Sub

