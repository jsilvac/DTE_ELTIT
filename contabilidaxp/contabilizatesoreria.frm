VERSION 5.00
Object = "{4777436C-EB8C-4596-98A8-EBCF98683969}#1.0#0"; "FlexCell.ocx"
Object = "{ADBBAED6-B16F-40DC-80DF-B44910CBA76C}#1.0#0"; "Frame-Xp.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form contabilizatesoreria 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Centralizacion de tesoreria"
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
      TabIndex        =   18
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
      Begin VB.CommandButton botonmisfavoritos 
         Caption         =   "Mis Favoritos"
         Height          =   255
         Left            =   1680
         TabIndex        =   20
         Top             =   280
         Width           =   1335
      End
      Begin VB.CommandButton botonmisaccesos 
         Caption         =   "Permisos Modulo"
         Height          =   255
         Left            =   120
         TabIndex        =   19
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
   Begin XPFrame.FrameXp FrameXp1 
      Height          =   8610
      Left            =   45
      TabIndex        =   1
      Top             =   45
      Width           =   14865
      _ExtentX        =   26220
      _ExtentY        =   15187
      BackColor       =   16744576
      Caption         =   "CENTRALIZACION DE TESORERIA"
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
      Alignment       =   1
      Begin VB.CommandButton BTELIMINA 
         BackColor       =   &H000000FF&
         Caption         =   "ELIMINA TODOS LOS COMPROBANTES"
         Height          =   330
         Left            =   600
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   8160
         Width           =   3375
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
         Left            =   7560
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   8160
         Width           =   2535
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H0080FF80&
         Caption         =   "IMPRIMIR"
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
         Left            =   4560
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   8160
         Width           =   2535
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
            Height          =   5955
            Left            =   45
            TabIndex        =   3
            Top             =   270
            Width           =   7125
            _ExtentX        =   12568
            _ExtentY        =   10504
            BackColorFixed  =   16744576
            Cols            =   5
            DefaultFontSize =   8.25
            GridColor       =   16711680
            Rows            =   30
            DateFormat      =   2
         End
         Begin FlexCell.Grid Ingresadas 
            Height          =   5970
            Left            =   7200
            TabIndex        =   16
            Top             =   270
            Width           =   7365
            _ExtentX        =   12991
            _ExtentY        =   10530
            BackColorFixed  =   0
            Cols            =   3
            DefaultFontSize =   9.75
            DefaultFontBold =   -1  'True
            ForeColorFixed  =   65535
            Rows            =   7
            DateFormat      =   2
         End
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FF8080&
         Caption         =   "DOBLE CLICK SOBRE FECHA PARA ELIMINAR ESE COMPROBANTE"
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
         Left            =   10320
         TabIndex        =   21
         Top             =   8160
         Width           =   4335
      End
   End
End
Attribute VB_Name = "contabilizatesoreria"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private localfiltro As String
Private cuentas(20) As String
Private lineascomprobante As Double
Private TIPODO As String
Private fecha2 As String
Private TIPODO2 As String
Private numero As String

Private cuenta As String
Private CRCC As String
Private monto As Double
Private DH As String
Private glosa As String
Private numerorut As String

Private Sub BTELIMINA_Click()


sqlconta.audit = True
sqlconta.programaactivo = Me.Caption
localfiltro = Mid(ComboLOCAL.text, 1, 2)
año = COMBOAÑO.text
MES = Format(COMBOMES.ListIndex + 1, "00")
If estacerrado(año + "-" + MES + "-" + Format(fechasistema, "dd")) <> True Then

If Verifica_Permiso(ingreso01.Caption, "elimina") = True Then
Call eliminacomprobantesmasivos("CC", MES, año, localfiltro)
Else
MsgBox mensaje_nopermiso
End If
COMMAND2_Click
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
Call Conectar(Servidor, clientesistema + "conta" + localfiltro, Usuario, password)
Call Conectargestion(Servidor, clientesistema + "gestion", Usuario, password)

leer


End Sub



Private Sub Command3_Click()
Dim k As Integer


End Sub


Private Sub Command4_Click()


sqlconta.audit = True
sqlconta.programaactivo = Me.Caption

Dim k As Double
Dim numero As String
Dim fecha As Date
Dim cuenta As String
Dim CRCC As String
Dim monto As Double
Dim DH As String
Dim glosa As String

Dim LINEA As String
Dim lin As Double
Dim DH2 As String
TIPODO = "CV"
lin = 0
lineascomprobante = 0

año = COMBOAÑO.text
MES = Format(COMBOMES.ListIndex + 1, "00")

If estacerrado(año + "-" + MES + "-" + Format(fechasistema, "dd")) <> True Then
        


For k = 1 To Grid1.Rows - 1
    If Grid1.Cell(k, 5).text = "1" And Grid1.Cell(k, 6).text <> "1" Then
    Call cargaIngresadas(Format(Grid1.Cell(k, 1).text, "yyyy-mm-dd"))
    Call contabilizacajas(Format(Grid1.Cell(k, 1).text, "yyyy-mm-dd"))
    End If
    
   
Next k
leer
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
 
    Call Conectar_BD
    Call Conectarteso(Servidor, clientesistema + "teso", Usuario, password)
    
    sc = 0
CARGAGRILLA
Call CargaGrillaIngresadas(1, 23)


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
       
    FORMATOGRILLA(1, 1) = "FECHA"
    FORMATOGRILLA(1, 2) = "RENDIDO"
    FORMATOGRILLA(1, 3) = "POR RENDIR"
    FORMATOGRILLA(1, 4) = "DIFERENCIA"
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
    
    Grid1.Cols = 7
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
   Grid1.Column(5).CellType = cellCheckBox
   Grid1.Column(6).CellType = cellCheckBox
   
   
   
    
    
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
    Call CARGAGRILLA
    
    LINEA = 0: fec = 0: fec1 = 0
    fecha1 = año + "-" + MES + "-" + "01"
    fecha2 = año + "-" + MES + "-" + "31"
  
    
        Set csql.ActiveConnection = teso
        csql.sql = "SELECT rc.fecha,sum(rc.totalrendido),sum(rc.totalarendir),sum(rc.totalrendido)-sum(rc.totalarendir) "
        csql.sql = csql.sql + "from rc_rendicionesdecaja as rc inner join " + clientesistema + "gestion.g_maestroempresas as me on (rc.local=me.codigo) "
        csql.sql = csql.sql + "where me.codigocontable='" + localfiltro + "' and rc.fecha>='" + fecha1 + "' AND rc.fecha<='" + fecha2 + "'  group by rc.fecha order by rc.fecha "
        csql.Execute
        total = 0
        total2 = 0
        Grid1.Rows = 1
        Grid1.AutoRedraw = False
        Barra.Max = csql.RowsAffected + 1
        
        Barra.Value = 0
        
        If csql.RowsAffected > 0 Then
        Set resultados = csql.OpenResultset
        
         While Not resultados.EOF
             Barra.Value = Barra.Value + 1
             Grid1.Rows = Grid1.Rows + 1
             LINEA = LINEA + 1
             
             Grid1.Cell(LINEA, 1).text = resultados(0)
             Grid1.Cell(LINEA, 2).text = resultados(1)
             Grid1.Cell(LINEA, 3).text = resultados(2)
             Grid1.Cell(LINEA, 4).text = resultados(3)
             Grid1.Cell(LINEA, 5).text = "0"
             Grid1.Cell(LINEA, 6).text = "0"
          
            If LEEDOCUMENTO("CC", resultados(0)) = True Then
             Grid1.Cell(LINEA, 6).text = "1"
            End If
             
            
            resultados.MoveNext
       
            Wend
End If
      
      Grid1.AutoRedraw = True
      Grid1.Refresh
      
      
      
End Sub
Private Sub contabilizacajas(fecha)

Dim resultados As rdoResultset
    Dim csql As New rdoQuery
    Dim csql2 As New rdoQuery
    
    numero = LEERULTIMOFOLIO("CC")
    lineascomprobante = 0
    For k = 1 To Ingresadas.Rows - 2
    
    If Ingresadas.Cell(k, 15).text <> "0" Then
    
    lineascomprobante = lineascomprobante + 1
    LINEA = Format(lineascomprobante, "000")
    fecha = Format(fecha, "yyyy-mm-dd")
    cuenta = leerdatoslocal(Ingresadas.Cell(k, 1).text, "cuentarendicion")
   
    CRCC = ""
    monto = CDbl(Ingresadas.Cell(k, 15).text) - CDbl(Ingresadas.Cell(k, 22).text)
    glosa = "CENTRALIZACION DE CAJAS " + leerdatoslocal(Ingresadas.Cell(k, 1).text, "nombre")
   
    Call grabarcomprobante_lineas("CC", numero, LINEA, fecha, cuenta, "", "", CRCC, glosa, "CC", numero, fecha, fecha, monto, "H", USUARIOSISTEMA, Format(fecha, "MM"), Format(fecha, "yyyy"), Format(fechasistema, "yyyy-mm-dd"), Time, "")
    
    End If
    
    If Ingresadas.Cell(k, 17).text <> "0" Then
    
    lineascomprobante = lineascomprobante + 1
    LINEA = Format(lineascomprobante, "000")
    fecha = Format(fecha, "yyyy-mm-dd")
    cuenta = "11200030"
   
    CRCC = ""
    If empresaactiva = "09" Then
    monto = CDbl(Ingresadas.Cell(k, 17).text) - Leerpagosdiagas(fecha, Ingresadas.Cell(k, 1).text)
    Else
    monto = CDbl(Ingresadas.Cell(k, 17).text)
    
    End If
    
    glosa = "CENTRALIZACION PAGOS " + leerdatoslocal(Ingresadas.Cell(k, 1).text, "nombre")
   
    Call grabarcomprobante_lineas("CC", numero, LINEA, fecha, cuenta, "", "", CRCC, glosa, "CC", numero, fecha, fecha, monto, "H", USUARIOSISTEMA, Format(fecha, "MM"), Format(fecha, "yyyy"), Format(fechasistema, "yyyy-mm-dd"), Time, "")
    
    End If
    
    If Ingresadas.Cell(k, 22).text <> "0" Then
    
    lineascomprobante = lineascomprobante + 1
    LINEA = Format(lineascomprobante, "000")
    fecha = Format(fecha, "yyyy-mm-dd")
    cuenta = leerdatoslocal(Ingresadas.Cell(k, 1).text, "cuentadonaciones")
    CRCC = ""
    monto = CDbl(Ingresadas.Cell(k, 22).text)
    glosa = "CENTRALIZACION DE CAJAS " + leerdatoslocal(Ingresadas.Cell(k, 1).text, "nombre")
    Call grabarcomprobante_lineas("CC", numero, LINEA, fecha, cuenta, "", "", CRCC, glosa, "CC", numero, fecha, fecha, monto, "H", USUARIOSISTEMA, Format(fecha, "MM"), Format(fecha, "yyyy"), Format(fechasistema, "yyyy-mm-dd"), Time, "")
    End If
    
    
    
    Next k
    Rem EFECTIVO
    If Ingresadas.Cell(Ingresadas.Rows - 1, 3).text <> "0" Then
    lineascomprobante = lineascomprobante + 1
    LINEA = Format(lineascomprobante, "000")
    fecha = Format(fecha, "yyyy-mm-dd")
    cuenta = "11100001"
    CRCC = ""
    monto = Ingresadas.Cell(Ingresadas.Rows - 1, 3).text
    glosa = "CONTABILIZA EFECTIVO "
    Call grabarcomprobante_lineas("CC", numero, LINEA, fecha, cuenta, "", "", CRCC, glosa, "CC", numero, fecha, fecha, monto, "D", USUARIOSISTEMA, Format(fecha, "MM"), Format(fecha, "yyyy"), Format(fechasistema, "yyyy-mm-dd"), Time, "")
    End If
    
    Rem CHEQUES DIA
    If Ingresadas.Cell(Ingresadas.Rows - 1, 4).text <> "0" Then
    lineascomprobante = lineascomprobante + 1
    LINEA = Format(lineascomprobante, "000")
    fecha = Format(fecha, "yyyy-mm-dd")
    cuenta = "11100002"
    CRCC = ""
    monto = Ingresadas.Cell(Ingresadas.Rows - 1, 4).text
    glosa = "CONTABILIZA CHEQUES DIA "
    Call grabarcomprobante_lineas("CC", numero, LINEA, fecha, cuenta, "", "", CRCC, glosa, "CC", numero, fecha, fecha, monto, "D", USUARIOSISTEMA, Format(fecha, "MM"), Format(fecha, "yyyy"), Format(fechasistema, "yyyy-mm-dd"), Time, "")
    End If
          
    Rem CHEQUES FECHA
    If Ingresadas.Cell(Ingresadas.Rows - 1, 5).text <> "0" Then
    lineascomprobante = lineascomprobante + 1
    LINEA = Format(lineascomprobante, "000")
    fecha = Format(fecha, "yyyy-mm-dd")
    cuenta = "11100003"
    CRCC = ""
    monto = Ingresadas.Cell(Ingresadas.Rows - 1, 5).text
    glosa = "CONTABILIZA CHEQUES FECHA "
    Call grabarcomprobante_lineas("CC", numero, LINEA, fecha, cuenta, "", "", CRCC, glosa, "CC", numero, fecha, fecha, monto, "D", USUARIOSISTEMA, Format(fecha, "MM"), Format(fecha, "yyyy"), Format(fechasistema, "yyyy-mm-dd"), Time, "")
    End If
          
    Rem CHEQUES TARJETA CREDITO
    If Ingresadas.Cell(Ingresadas.Rows - 1, 6).text <> "0" Then
    lineascomprobante = lineascomprobante + 1
    LINEA = Format(lineascomprobante, "000")
    fecha = Format(fecha, "yyyy-mm-dd")
    cuenta = "11100004"
    CRCC = ""
    monto = Ingresadas.Cell(Ingresadas.Rows - 1, 6).text
    glosa = "CONTABILIZA TARJETAS DE CREDITO  "
    Call grabarcomprobante_lineas("CC", numero, LINEA, fecha, cuenta, "", "", CRCC, glosa, "CC", numero, fecha, fecha, monto, "D", USUARIOSISTEMA, Format(fecha, "MM"), Format(fecha, "yyyy"), Format(fechasistema, "yyyy-mm-dd"), Time, "")
    End If
    
    Rem CHEQUES TARJETA DEBITO
    If Ingresadas.Cell(Ingresadas.Rows - 1, 7).text <> "0" Then
    lineascomprobante = lineascomprobante + 1
    LINEA = Format(lineascomprobante, "000")
    fecha = Format(fecha, "yyyy-mm-dd")
    cuenta = "11100008"
    CRCC = ""
    monto = Ingresadas.Cell(Ingresadas.Rows - 1, 7).text
    glosa = "CONTABILIZA TARJETAS DE DEBITO  "
    Call grabarcomprobante_lineas("CC", numero, LINEA, fecha, cuenta, "", "", CRCC, glosa, "CC", numero, fecha, fecha, monto, "D", USUARIOSISTEMA, Format(fecha, "MM"), Format(fecha, "yyyy"), Format(fechasistema, "yyyy-mm-dd"), Time, "")
    End If
    
    Rem CHEQUES TARJETA GIFTCARD
    
    If Ingresadas.Cell(Ingresadas.Rows - 1, 23).text <> "0" And Ingresadas.Cell(Ingresadas.Rows - 1, 23).text <> "" Then
    lineascomprobante = lineascomprobante + 1
    LINEA = Format(lineascomprobante, "000")
    fecha = Format(fecha, "yyyy-mm-dd")
    cuenta = "11100009"
    CRCC = ""
    monto = Ingresadas.Cell(Ingresadas.Rows - 1, 23).text
    glosa = "CONTABILIZA GIFT CARD  "
    Call grabarcomprobante_lineas("CC", numero, LINEA, fecha, cuenta, "", "", CRCC, glosa, "CC", numero, fecha, fecha, monto, "D", USUARIOSISTEMA, Format(fecha, "MM"), Format(fecha, "yyyy"), Format(fechasistema, "yyyy-mm-dd"), Time, "")
    End If
    
    
'     Rem CARGAR REDONDEO
'
'    If Ingresadas.Cell(Ingresadas.Rows - 1, 23).text <> "0" And Ingresadas.Cell(Ingresadas.Rows - 1, 23).text <> "" Then
'    lineascomprobante = lineascomprobante + 1
'    LINEA = Format(lineascomprobante, "000")
'    fecha = Format(fecha, "yyyy-mm-dd")
'    cuenta = "11100009"
'    CRCC = ""
'    monto = Ingresadas.Cell(Ingresadas.Rows - 1, 23).text
'    glosa = "CONTABILIZA REDONDEO  "
'    Call grabarcomprobante_lineas("CC", numero, LINEA, fecha, cuenta, "", "", CRCC, glosa, "CC", numero, fecha, fecha, monto, "D", USUARIOSISTEMA, Format(fecha, "MM"), Format(fecha, "yyyy"), Format(fechasistema, "yyyy-mm-dd"), Time, "")
'    End If
    
    
    Call CARGACASAS(fecha)
    
    Call CARGACREDITOSTESO(fecha)
    Call CARGANCTESO(fecha)
    Call CARGAVALESDEcompra(fecha)
    Call cargavarios(fecha)
    Call CARGADIFERENCIASDECAJA(fecha)
    Call CARGASOBRANTECAJA(fecha)
    Call CARGAREDONDEO(fecha)
    Rem Call CARGADIFERENCIASDECAJAPOSITIVA(fecha)
    Call CARGAVALESDEcredito(fecha)
    If empresaactiva = "09" Then
    Call cargapagosdiagas(fecha)
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
    Dim k As Double
    
        Set csql.ActiveConnection = conta
        csql.sql = "SELECT codigoempresa,nombre "
        csql.sql = csql.sql + "FROM maestroempresas "
        csql.sql = csql.sql + "ORDER BY codigoempresa "
        csql.Execute
        
        If csql.RowsAffected > 0 Then
            Set resultados = csql.OpenResultset
            While Not resultados.EOF
                ComboLOCAL.AddItem (resultados(0) + " " + resultados(1))
                resultados.MoveNext
            Wend
            resultados.Close
            Set resultados = Nothing
            For k = 0 To ComboLOCAL.ListCount - 1
                If Mid(ComboLOCAL.List(k), 1, 2) = empresaactiva Then
                    ComboLOCAL.text = ComboLOCAL.List(k)
                    Exit For
                End If
            Next k
        
        End If
        localfiltro = Mid(ComboLOCAL.text, 1, 2)
        
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



Private Sub Form_Unload(Cancel As Integer)
'empresaactiva = localfiltro

End Sub

Private Sub Grid1_DblClick()
    
    If Grid1.Rows > 1 And Grid1.ActiveCell.col = 2 Then
        Call cargaIngresadas(Grid1.Cell(Grid1.ActiveCell.row, 1).text)
    End If
    If Grid1.Rows > 1 And Grid1.ActiveCell.col = 1 And Grid1.Cell(Grid1.ActiveCell.row, 6).text = 1 Then
         If MsgBox("DESDEA ELIMINAR FECHA SELECCIONADA", vbYesNo, "ATENCION") = vbYes Then
            Call eliminacomprobante(Grid1.Cell(Grid1.ActiveCell.row, 1).text, empresaactiva, "CC")
         End If
    End If
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

    campos(0, 2) = "movimientoscontables"
   

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
    campos(0, 1) = Format(fechasistema, "yyyy")
    campos(1, 1) = tipo
    campos(2, 1) = rut
    campos(3, 1) = nombrecliente(rut)
    condicion = "tipo='" + tipo + "' and rut='" + rut + "' and año='" + año + "' "
    campos(0, 2) = "cuentascorrientes"
    op = 5
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)
    If sqlconta.status = 0 Then
    existerut = True
    Else
    Call grabar(año, tipo, rut, nombrecliente(rut))
    
    End If

    
    End Function
   Public Function existerut2(año, tipo, rut, NOMBRE) As Boolean
   
    campos(0, 0) = "año"
    campos(1, 0) = "tipo"
    campos(2, 0) = "rut"
    campos(3, 0) = "nombre"
    campos(4, 0) = ""
    campos(0, 1) = Format(fechasistema, "yyyy")
    campos(1, 1) = tipo
    campos(2, 1) = rut
    campos(3, 1) = nombrecliente(rut)
    condicion = "tipo='" + tipo + "' and rut='" + rut + "' and año='" + año + "' "
    campos(0, 2) = "cuentascorrientes"
    op = 5
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)
    If sqlconta.status = 0 Then
    existerut2 = True
    Else
    Call grabar(año, tipo, rut, NOMBRE)
    
    End If

    
    End Function
    
Public Function nombrecliente(rut) As String
    campos(0, 0) = "nombre"
    campos(1, 0) = ""
    condicion = "rut='" + rut + "' "
    campos(0, 2) = clientesistema + "ventas.sv_maestroclientes"
    op = 5
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)
    If sqlconta.status = 0 Then
    nombrecliente = sqlconta.response(0, 3)
    Else
    nombrecliente = ""
    End If
    End Function

Sub grabar(año, tipo, rut, NOMBRE)
    campos(0, 0) = "año"
    campos(1, 0) = "tipo"
    campos(2, 0) = "rut"
    campos(3, 0) = "nombre"
    campos(4, 0) = ""
    campos(0, 1) = Format(fechasistema, "yyyy")
    campos(1, 1) = tipo
    campos(2, 1) = rut
    campos(3, 1) = NOMBRE
    
    campos(0, 2) = "cuentascorrientes"
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
    
    campos(0, 2) = "saldosctacte"
    op = 2
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)
    

End Sub
Private Function Leerpagosdiagas(fecha, loc) As Double
 Dim sql As New rdoQuery
    Dim resultados As rdoResultset
        Set sql.ActiveConnection = contadb
  
        sql.sql = "SELECT ifnull(sum(dd.total),'0')  "
        sql.sql = sql.sql & "FROM eltit_ventas.sv_pagos_gas_cabeza dc inner join  "
        sql.sql = sql.sql & "eltit_ventas.sv_pagos_gas_detalle AS dd on dc.local=dd.local "
        sql.sql = sql.sql & "and dc.numero=dd.numero and dc.fecha=dd.fecha and dc.tipo=dd.tipo "
        sql.sql = sql.sql & "and dc.caja=dd.caja  "
        sql.sql = sql.sql & "where dc.fecha='" & fecha & "' and dc.caja<'80' and dc.tipo='VG' and dc.atencion='ACTIVA' and dc.local='" + loc + "' "
        
        sql.Execute
            If sql.RowsAffected > 0 Then
            Set resultados = sql.OpenResultset
            While Not resultados.EOF
                If IsNull(resultados(0)) = False Then
                    Leerpagosdiagas = resultados(0)
                End If
                resultados.MoveNext
            Wend
            Else
            Leerpagosdiagas = 0
            End If
        
End Function



Private Sub cargaIngresadas(fecha)
Dim totales(30) As Double
Dim chequesdia As Double
Dim chequesfecha As Double

 Dim sql As New rdoQuery
    Dim resultados As rdoResultset
    Set sql.ActiveConnection = teso
    chequesdia = 0
    chequesfecha = 0
    Dim suma As Double
    sql.sql = "SELECT rc.local,rc.codigocajera,sum(rc.totalefectivo),sum(rc.montocheques),sum(rc.montochequefecha),sum(rc.montotcredito),sum(rc.montotdebito),sum(rc.montoextranjera),sum(rc.montootorgado),sum(rc.montovcompra),sum(rc.montoocredito),sum(rc.montoncredito),sum(rc.montovarios),sum(rc.totalrendido),sum(rc.totalventas)+sum(rc.totaldonaciones),'0',sum(rc.totalpagos),sum(rc.totalarendir),sum(rc.totalrendido),sum(rc.totalrendido)-sum(rc.totalarendir),rc.glosa ,sum(rc.donacion),sum(rc.montogiftcard),sum(rc.redondeo) "
    sql.sql = sql.sql + "FROM rc_rendicionesdecaja as rc inner join " + clientesistema + "gestion.g_maestroempresas as me on (me.codigo=rc.local) "
    sql.sql = sql.sql + "WHERE fecha = '" & Format(fecha, "yyyy-mm-dd") & "' and me.codigocontable='" + localfiltro + "' group by rc.local "
    sql.sql = sql.sql + " order by local,codigocajera "
    sql.Execute
    For k = 1 To 30
    totales(k) = 0
    Next k
    chequesdia = 0
    chequesfecha = 0
   Ingresadas.Rows = 1
    Ingresadas.AutoRedraw = False
    
  
    If sql.RowsAffected > 0 Then
        Set resultados = sql.OpenResultset
        Ingresadas.Rows = 1
        While resultados.EOF = False
      
        Ingresadas.Rows = Ingresadas.Rows + 1
        
            Ingresadas.Cell(Ingresadas.Rows - 1, 1).text = resultados(0)
            Ingresadas.Cell(Ingresadas.Rows - 1, 2).text = resultados(1)
            For k = 2 To 22
            Ingresadas.Cell(Ingresadas.Rows - 1, k + 1).text = resultados(k)
            If k = 3 Then
            Ingresadas.Cell(Ingresadas.Rows - 1, k + 1).text = montocheques(Format(fecha, "yyyy-mm-dd"), resultados(0), Mid(resultados(1), 1, 9), fecha, "1")
            End If
            If k = 4 Then
            Ingresadas.Cell(Ingresadas.Rows - 1, k + 1).text = montocheques(Format(fecha, "yyyy-mm-dd"), resultados(0), Mid(resultados(1), 1, 9), fecha, "2")
            End If
            Ingresadas.Cell(Ingresadas.Rows - 1, 21).text = resultados(20)
            Ingresadas.Cell(Ingresadas.Rows - 1, 22).text = resultados(21)
            Ingresadas.Cell(Ingresadas.Rows - 1, 23).text = resultados(22)
            Ingresadas.Cell(Ingresadas.Rows - 1, 24).text = resultados(23)
            If k <> 20 Then
            totales(k) = totales(k) + CDbl(Ingresadas.Cell(Ingresadas.Rows - 1, k + 1).text)
            End If
            
            Next k
            resultados.MoveNext
        Wend
           
            
            Ingresadas.Rows = Ingresadas.Rows + 1
            Ingresadas.Cell(Ingresadas.Rows - 1, 2).text = "TOTALES "
            Ingresadas.Range(Ingresadas.Rows - 1, 2, Ingresadas.Rows - 1, Ingresadas.Cols - 1).Borders(cellEdgeTop) = cellThin
            Ingresadas.Range(Ingresadas.Rows - 1, 2, Ingresadas.Rows - 1, Ingresadas.Cols - 1).Borders(cellEdgeBottom) = cellThin
            Ingresadas.Range(Ingresadas.Rows - 1, 2, Ingresadas.Rows - 1, Ingresadas.Cols - 1).Borders(cellEdgeLeft) = cellThin
            Ingresadas.Range(Ingresadas.Rows - 1, 2, Ingresadas.Rows - 1, Ingresadas.Cols - 1).Borders(cellEdgeRight) = cellThin
            Ingresadas.Range(Ingresadas.Rows - 1, 2, Ingresadas.Rows - 1, Ingresadas.Cols - 1).Borders(cellInsideHorizontal) = cellThin
            Ingresadas.Range(Ingresadas.Rows - 1, 2, Ingresadas.Rows - 1, Ingresadas.Cols - 1).Borders(cellInsideVertical) = cellThin
    
            For k = 2 To 22
            Ingresadas.Cell(Ingresadas.Rows - 1, k + 1).text = totales(k)
            Next k
            
            
            
            
    End If
Ingresadas.AutoRedraw = True
            Ingresadas.Refresh

    


End Sub

Private Sub CargaGrillaIngresadas(ByVal row As Integer, ByVal col As Integer)
        Dim i As Integer
        Rem DATOS DE LA COLUMNA
        col = 25
        Dim FORMATOGRILLA(20, 30) As String
        
        FORMATOGRILLA(1, 1) = "LO": FORMATOGRILLA(8, 1) = "10"
        FORMATOGRILLA(1, 2) = "CAJERA": FORMATOGRILLA(8, 2) = "0"
        FORMATOGRILLA(1, 3) = "EFECTIVO": FORMATOGRILLA(8, 3) = "8"
        FORMATOGRILLA(1, 4) = "CHE.DIAS": FORMATOGRILLA(8, 4) = "8"
        FORMATOGRILLA(1, 5) = "CHE.FEC ": FORMATOGRILLA(8, 5) = "8"
        FORMATOGRILLA(1, 6) = "T.CREDITO": FORMATOGRILLA(8, 6) = "0"
        FORMATOGRILLA(1, 7) = "T.DEBITO": FORMATOGRILLA(8, 7) = "0"
        FORMATOGRILLA(1, 8) = "M.EXT": FORMATOGRILLA(8, 8) = "8"
        FORMATOGRILLA(1, 9) = "CREDITOS": FORMATOGRILLA(8, 9) = "8"
        FORMATOGRILLA(1, 10) = "V.COMPRA": FORMATOGRILLA(8, 10) = "0"
        FORMATOGRILLA(1, 11) = "O.CRED.": FORMATOGRILLA(8, 11) = "0"
        FORMATOGRILLA(1, 12) = "N.CRED.": FORMATOGRILLA(8, 12) = "8"
        FORMATOGRILLA(1, 13) = "VARIOS": FORMATOGRILLA(8, 13) = "0"
        FORMATOGRILLA(1, 14) = "T.RENDIDO": FORMATOGRILLA(8, 14) = "0"
        FORMATOGRILLA(1, 15) = "T. VENTA": FORMATOGRILLA(8, 15) = "8"
        FORMATOGRILLA(1, 16) = "DONAC.": FORMATOGRILLA(8, 16) = "0"
        FORMATOGRILLA(1, 17) = "T.PAGOS": FORMATOGRILLA(8, 17) = "8"
        FORMATOGRILLA(1, 18) = "T.A RENDIR": FORMATOGRILLA(8, 18) = "8"
        FORMATOGRILLA(1, 19) = "T.RENDIDO": FORMATOGRILLA(8, 19) = "8"
        FORMATOGRILLA(1, 20) = "D.CAJA": FORMATOGRILLA(8, 20) = "8"
        FORMATOGRILLA(1, 21) = "GIFT.CARD": FORMATOGRILLA(8, 21) = "8"
        FORMATOGRILLA(1, 22) = "": FORMATOGRILLA(8, 22) = "0"
        FORMATOGRILLA(1, 23) = "": FORMATOGRILLA(8, 22) = "8"
        FORMATOGRILLA(1, 24) = "": FORMATOGRILLA(8, 22) = "8"
        
        Rem LARGO DE LOS DATOS
        FORMATOGRILLA(2, 1) = "15"
        FORMATOGRILLA(2, 2) = "20"
        For k = 3 To 24
        FORMATOGRILLA(2, k) = "8"
        Next k
        Rem TIPO DE DATOS C=CEROS,S=STRING,N=NUMERICO,D=DATE,
        FORMATOGRILLA(3, 1) = "S"
        FORMATOGRILLA(3, 2) = "S"
        For k = 3 To 24
        FORMATOGRILLA(3, k) = "N"
        Next k
        
        Rem FORMATO GRILLA
        FORMATOGRILLA(4, 1) = ""
        FORMATOGRILLA(4, 2) = ""
        For k = 3 To 24
        FORMATOGRILLA(4, k) = "###,###,##0"
        Next k
        
        Rem LOCCKED
        FORMATOGRILLA(5, 1) = "TRUE"
        FORMATOGRILLA(5, 2) = "TRUE"
        For k = 3 To 24
        FORMATOGRILLA(5, k) = "TRUE"
        Next k
        
        Rem VALOR MINIMO
        FORMATOGRILLA(6, 1) = ""
        FORMATOGRILLA(6, 2) = ""
        FORMATOGRILLA(6, 3) = ""
        
        Rem VALOR MAXIMO
        FORMATOGRILLA(7, 1) = ""
        FORMATOGRILLA(7, 2) = ""
        FORMATOGRILLA(7, 3) = ""
        
        
        Ingresadas.Cols = col
        Ingresadas.Rows = row
        Ingresadas.AllowUserResizing = False
        Ingresadas.DisplayFocusRect = False
        Ingresadas.ExtendLastCol = False
        
        Ingresadas.BoldFixedCell = False
        Ingresadas.DrawMode = cellOwnerDraw
        Ingresadas.Appearance = Flat
        Ingresadas.ScrollBarStyle = Flat
        Ingresadas.FixedRowColStyle = Flat
'        Ingresadas.BackColorFixed = RGB(90, 158, 214)
'        Ingresadas.BackColorFixedSel = RGB(110, 180, 230)
'        Ingresadas.BackColorBkg = RGB(90, 158, 214)
'        Ingresadas.BackColorScrollBar = RGB(231, 235, 247)
'        Ingresadas.BackColor1 = RGB(231, 235, 247)
'        Ingresadas.BackColor2 = RGB(239, 243, 255)
'        Ingresadas.GridColor = RGB(148, 190, 231)
        Ingresadas.Column(0).Width = 0
        
        For i = 1 To col - 1
            Ingresadas.Cell(0, i).text = FORMATOGRILLA(1, i)
            Ingresadas.Column(i).Width = Val(FORMATOGRILLA(8, i)) * (Ingresadas.Cell(0, i).Font.Size + 1.25)
            Ingresadas.Column(i).MaxLength = Val(FORMATOGRILLA(2, i))
            Ingresadas.Column(i).FormatString = FORMATOGRILLA(4, i)
            Ingresadas.Column(i).Locked = FORMATOGRILLA(5, i)
            If FORMATOGRILLA(3, i) = "N" Then
                Ingresadas.Column(i).Alignment = cellRightCenter
                Ingresadas.Column(i).Mask = cellNumeric
            Else
                Ingresadas.Column(i).Alignment = cellLeftCenter
                Ingresadas.Column(i).Mask = cellUpper
            End If
        Next i
        Ingresadas.Range(0, 0, 0, col - 1).Alignment = cellCenterCenter
        
        Ingresadas.Enabled = True
    End Sub
Public Function montocheques(fecha, loc, rut, vencimiento, tipo) As Double

 Dim sql As New rdoQuery
    Dim resultados As rdoResultset
        Set sql.ActiveConnection = teso
        If tipo = "1" Then
        sql.sql = "SELECT sum(monto),count(monto) "
        sql.sql = sql.sql + "FROM rc_cartera "
        sql.sql = sql.sql + "WHERE fecha = '" & fecha & "' and local='" + loc + "' "
        sql.sql = sql.sql + "and cartera='N' "
        
        sql.sql = sql.sql + "group by fecha "
        
        
        sql.Execute
        montocheques = 0
            If sql.RowsAffected > 0 Then
            Set resultados = sql.OpenResultset
            montocheques = resultados(0)
                
            End If
        
        
        End If
        
        If tipo = "2" Then
        sql.sql = "SELECT sum(monto),count(monto) "
        sql.sql = sql.sql + "FROM rc_cartera "
        sql.sql = sql.sql + "WHERE fecha = '" & fecha & "' and local='" + loc + "' "
        sql.sql = sql.sql + "and cartera='S' "
        sql.sql = sql.sql + "group by fecha "
        sql.Execute
        montocheques = 0
    
            If sql.RowsAffected > 0 Then
            Set resultados = sql.OpenResultset
            montocheques = resultados(0)
            
            End If
        
        End If
        
    
End Function



Sub CARGACREDITOSTESO(fecha)
    Dim sql As New rdoQuery
    Dim resultados As rdoResultset
    Set sql.ActiveConnection = teso
    Dim suma As Double
    
    
    
    
    '****************************
    'TARJETA CREDITO
    '***************************
   
    
    sql.sql = "SELECT co.localcredito,co.rut,sum(co.total) "
    sql.sql = sql.sql + "FROM rc_creditosotorgados as co inner join " + clientesistema + "gestion.g_maestroempresas as me on (me.codigo=co.local) "
    sql.sql = sql.sql + "WHERE co.fecha = '" & Format(fecha, "yyyy-mm-dd") & "' AND me.codigocontable = '" & localfiltro & "' "
    sql.sql = sql.sql + "group by localcredito order by localcredito"
    sql.Execute
    
    If sql.RowsAffected > 0 Then
        Set resultados = sql.OpenResultset
        While resultados.EOF = False
    lineascomprobante = lineascomprobante + 1
    LINEA = Format(lineascomprobante, "000")
    fecha = Format(fecha, "yyyy-mm-dd")
    cuenta = "11200030"
    CRCC = ""
    monto = resultados(2)
    glosa = "CONTABILIZA CREDITOS " + leerdatoslocal(resultados(0), "nombre")
    Call grabarcomprobante_lineas("CC", numero, LINEA, fecha, cuenta, "", "", CRCC, glosa, "CC", numero, fecha, fecha, monto, "D", USUARIOSISTEMA, Format(fecha, "MM"), Format(fecha, "yyyy"), Format(fechasistema, "yyyy-mm-dd"), Time, "")
            
        resultados.MoveNext
        Wend
     
    End If
End Sub

Sub CARGANCTESO(fecha)
    Dim sql As New rdoQuery
    Dim resultados As rdoResultset
    Set sql.ActiveConnection = teso
    Dim suma As Double
    Dim montorelacionada As Double
    
    Dim rut2 As String 'no dejaba compilar rz 12-02-2020
    
    
    '****************************
    'TARJETA CREDITO eltit
    '***************************
   
    
    sql.sql = "SELECT co.local,co.numero,co.monto,co.rut "
    sql.sql = sql.sql + "FROM rc_notasdecreditos as co inner join " + clientesistema + "gestion.g_maestroempresas as me on (me.codigo=co.local) "
    sql.sql = sql.sql + "WHERE co.fecha = '" & Format(fecha, "yyyy-mm-dd") & "' AND me.codigocontable = '" & localfiltro & "' "
    sql.sql = sql.sql + "order by local"
    sql.Execute
    montorelacionada = 0
    If sql.RowsAffected > 0 Then
        Set resultados = sql.OpenResultset
        While resultados.EOF = False
            lineascomprobante = lineascomprobante + 1
            LINEA = Format(lineascomprobante, "000")
            fecha = Format(fecha, "yyyy-mm-dd")
            cuenta = leerdatoslocal(resultados(0), "cuentarendicion")
            CRCC = ""
            monto = resultados(2)
            rut2 = ""
            glosa = "CONTABILIZA N.CREDITO " + leerdatoslocal(resultados(0), "nombre")
            Call grabarcomprobante_lineas("CC", numero, LINEA, fecha, cuenta, "", rut2, CRCC, glosa, "NC", resultados(1), fecha, fecha, monto, "D", USUARIOSISTEMA, Format(fecha, "MM"), Format(fecha, "yyyy"), Format(fechasistema, "yyyy-mm-dd"), Time, "")
          
            
            If esempresarelacionada(resultados(3)) = True Then
            lineascomprobante = lineascomprobante + 1
            LINEA = Format(lineascomprobante, "000")
            fecha = Format(fecha, "yyyy-mm-dd")
            cuenta = "11200029"
            rut2 = resultados(3)
            
            Call grabarcomprobante_lineas("CC", numero, LINEA, fecha, cuenta, "", rut2, CRCC, glosa, "NC", resultados(1), fecha, fecha, monto, "H", USUARIOSISTEMA, Format(fecha, "MM"), Format(fecha, "yyyy"), Format(fechasistema, "yyyy-mm-dd"), Time, "")
            
            
            End If
            
          
        resultados.MoveNext
        Wend
        
            
        
        
    End If
End Sub
Function esempresarelacionada(RUTEMPRE) As Boolean
    Dim csql As New rdoQuery
    Dim resultados As rdoResultset
    Set csql.ActiveConnection = teso
     csql.sql = "select rut from " & cliente_sql & "conta.maestroempresas "
     csql.sql = csql.sql & "WHERE  LPAD(MID(rut,1,LENGTH(rut)-2),9,0)='" & Mid(RUTEMPRE, 1, 9) & "' "
     csql.Execute
        esempresarelacionada = False
     If csql.RowsAffected > 0 Then
        Set resultados = csql.OpenResultset
        While Not resultados.EOF
        
            esempresarelacionada = True
            resultados.MoveNext
        Wend
     End If
     csql.Close
     Set csql = Nothing
     
     
End Function
Sub cargapagosdiagas(fecha)
 Dim sql As New rdoQuery
    Dim resultados As rdoResultset
        Set sql.ActiveConnection = contadb
  
        sql.sql = "SELECT dc.numero,dc.rut,sum(dd.total),dc.nombre "
        sql.sql = sql.sql & "FROM eltit_ventas.sv_pagos_gas_cabeza dc inner join  "
        sql.sql = sql.sql & "eltit_ventas.sv_pagos_gas_detalle AS dd on dc.local=dd.local "
        sql.sql = sql.sql & "and dc.numero=dd.numero and dc.fecha=dd.fecha and dc.tipo=dd.tipo "
        sql.sql = sql.sql & "and dc.caja=dd.caja  "
        sql.sql = sql.sql & "where dc.fecha='" & fecha & "' and dc.caja<'80' and dc.tipo='VG' and dc.atencion='ACTIVA' and (dc.local='01' or dc.local='20') group by dc.numero "
        
        sql.Execute
            If sql.RowsAffected > 0 Then
            Set resultados = sql.OpenResultset
            While Not resultados.EOF
                If IsNull(resultados(0)) = False Then
    
    lineascomprobante = lineascomprobante + 1
    LINEA = Format(lineascomprobante, "000")
    fecha = Format(fecha, "yyyy-mm-dd")
    cuenta = "11500200"
    CRCC = ""
    monto = resultados(2)
    glosa = "PAGOS GAS "
    Call existerut2(Format(fecha, "yyyy"), cuenta, resultados(1), resultados(3))
    Call grabarcomprobante_lineas("CC", numero, LINEA, fecha, cuenta, cuenta, resultados(1), CRCC, glosa, "VG", resultados(0), fecha, fecha, monto, "H", USUARIOSISTEMA, Format(fecha, "MM"), Format(fecha, "yyyy"), Format(fechasistema, "yyyy-mm-dd"), Time, "")
                    
                End If
                resultados.MoveNext
            Wend
          End If
          
          
          
        
End Sub


Sub CARGAVALESDEcompra(fecha)
    Dim sql As New rdoQuery
    Dim resultados As rdoResultset
    Set sql.ActiveConnection = teso
    Dim suma As Double
    Dim montoweb As Double
    
    
    
    
    
    '****************************
    'TARJETA vales eltit
    '***************************
   
    
    sql.sql = "SELECT co.local,co.numero,co.monto,co.rut,co.caja "
    sql.sql = sql.sql + "FROM rc_valesdecompra as co inner join " + clientesistema + "gestion.g_maestroempresas as me on (me.codigo=co.local) "
    sql.sql = sql.sql + "WHERE co.fecha = '" & Format(fecha, "yyyy-mm-dd") & "' AND me.codigocontable = '" & localfiltro & "' "
    sql.sql = sql.sql + "order by co.local"
    sql.Execute
    
    If sql.RowsAffected > 0 Then
        Set resultados = sql.OpenResultset
        While resultados.EOF = False
            lineascomprobante = lineascomprobante + 1
            LINEA = Format(lineascomprobante, "000")
            fecha = Format(fecha, "yyyy-mm-dd")
            If resultados(4) = "60" Then
                    cuenta = "11100008"
                    glosa = "CONTABILIZA VENTA WEB " + leerdatoslocal(resultados(0), "nombre")
            Else
                 cuenta = "11500190"
                 glosa = "CONTABILIZA VALE COMPRA " + leerdatoslocal(resultados(0), "nombre")
            End If
            CRCC = ""
            monto = resultados(2)
            
            Call grabarcomprobante_lineas("CC", numero, LINEA, fecha, cuenta, cuenta, resultados(3), CRCC, glosa, "VC", resultados(1), fecha, fecha, monto, "D", USUARIOSISTEMA, Format(fecha, "MM"), Format(fecha, "yyyy"), Format(fechasistema, "yyyy-mm-dd"), Time, "")
            
        resultados.MoveNext
        Wend
        
        If montoweb > 0 Then
            
        End If
     
    End If
End Sub
Sub cargavarios(fecha)
    Dim sql As New rdoQuery
    Dim resultados As rdoResultset
    Set sql.ActiveConnection = teso
    Dim suma As Double
    Dim CUENTA2 As String
    Dim rut2 As String
    
    
    
    
    '****************************
    'TARJETA vales eltit
    '***************************
   
    
    sql.sql = "SELECT co.local,co.cuenta,co.rut,co.crcc,co.glosa,co.tipo,co.numero,co.monto,co.dh "
    sql.sql = sql.sql + "FROM rc_contabilidad as co inner join " + clientesistema + "gestion.g_maestroempresas as me on (me.codigo=co.local) "
    sql.sql = sql.sql + "WHERE co.fecha = '" & Format(fecha, "yyyy-mm-dd") & "' AND me.codigocontable = '" & localfiltro & "' "
    sql.sql = sql.sql + "order by co.local"
    sql.Execute
    
    If sql.RowsAffected > 0 Then
        Set resultados = sql.OpenResultset
        While resultados.EOF = False
    lineascomprobante = lineascomprobante + 1
    LINEA = Format(lineascomprobante, "000")
    fecha = Format(fecha, "yyyy-mm-dd")
    cuenta = resultados(1)
    CRCC = resultados(3)
    monto = resultados(7)
    glosa = resultados(4)
    If resultados(2) <> "" Then
    CUENTA2 = cuenta
    End If
    
    Call grabarcomprobante_lineas("CC", numero, LINEA, fecha, cuenta, CUENTA2, resultados(2), CRCC, glosa, resultados(5), resultados(6), fecha, fecha, monto, resultados(8), USUARIOSISTEMA, Format(fecha, "MM"), Format(fecha, "yyyy"), Format(fechasistema, "yyyy-mm-dd"), Time, "")
            
        resultados.MoveNext
        Wend
     
    End If
End Sub


Sub CARGADIFERENCIASDECAJA(fecha)
    Dim sql As New rdoQuery
    Dim resultados As rdoResultset
    Set sql.ActiveConnection = teso
    Dim suma As Double
    
    
    
    
    '****************************
    'TARJETA CREDITO
    '***************************
   
    
    sql.sql = "SELECT co.local,co.totalarendir-co.totalrendido,co.codigocajera,co.glosa "
    sql.sql = sql.sql + "FROM rc_rendicionesdecaja as co inner join " + clientesistema + "gestion.g_maestroempresas as me on (me.codigo=co.local) "
    sql.sql = sql.sql + "WHERE co.fecha = '" & Format(fecha, "yyyy-mm-dd") & "' AND me.codigocontable = '" & localfiltro & "' "
    sql.sql = sql.sql + "and co.totalarendir-co.totalrendido>0  order by local"
    sql.Execute
    
    If sql.RowsAffected > 0 Then
        Set resultados = sql.OpenResultset
        While resultados.EOF = False
    lineascomprobante = lineascomprobante + 1
    LINEA = Format(lineascomprobante, "000")
    fecha = Format(fecha, "yyyy-mm-dd")
    cuenta = "11250005"
    CRCC = ""
    monto = resultados(1)
    glosa = "D/C" + resultados(3)
    Call crearcajera(resultados(2))
    Call grabarcomprobante_lineas("CC", numero, LINEA, fecha, cuenta, cuenta, resultados(2), CRCC, glosa, "CC", numero, fecha, fecha, monto, "D", USUARIOSISTEMA, Format(fecha, "MM"), Format(fecha, "yyyy"), Format(fechasistema, "yyyy-mm-dd"), Time, "")
            
        resultados.MoveNext
        Wend
     
    End If
End Sub

Sub CARGAREDONDEO(fecha)
    Dim sql As New rdoQuery
    Dim resultados As rdoResultset
    Set sql.ActiveConnection = teso
    Dim suma As Double
    Dim DH As String
    
    
    
    
    '****************************
    'TARJETA CREDITO
    '***************************
   
    
    sql.sql = "SELECT co.local,sum(co.redondeo),co.codigocajera,co.glosa "
    sql.sql = sql.sql + "FROM rc_rendicionesdecaja as co inner join " + clientesistema + "gestion.g_maestroempresas as me on (me.codigo=co.local) "
    sql.sql = sql.sql + "WHERE co.fecha = '" & Format(fecha, "yyyy-mm-dd") & "' AND me.codigocontable = '" & localfiltro & "' "
    sql.sql = sql.sql + "and co.redondeo<>0  group by co.local order by local"
    sql.Execute
    
    If sql.RowsAffected > 0 Then
        Set resultados = sql.OpenResultset
        While resultados.EOF = False
    lineascomprobante = lineascomprobante + 1
    LINEA = Format(lineascomprobante, "000")
    fecha = Format(fecha, "yyyy-mm-dd")
    cuenta = "47600015"
    CRCC = leerdatoslocal(resultados(0), "codigocrcc")
    
    monto = resultados(1)
    glosa = "REDONDEO "
    DH = "D"
    If monto > 0 Then
        DH = "H"
    Else
        DH = "D"
        monto = monto * -1
    End If
'    Call crearcajera(resultados(2))
    Call grabarcomprobante_lineas("CC", numero, LINEA, fecha, cuenta, cuenta, "", CRCC, glosa, "CC", numero, fecha, fecha, monto, DH, USUARIOSISTEMA, Format(fecha, "MM"), Format(fecha, "yyyy"), Format(fechasistema, "yyyy-mm-dd"), Time, "")
            
        resultados.MoveNext
        Wend
     
    End If
End Sub
Sub CARGASOBRANTECAJA(fecha)
    Dim sql As New rdoQuery
    Dim resultados As rdoResultset
    Set sql.ActiveConnection = teso
    Dim suma As Double
    Dim cajera As String
    
    
    
    '****************************
    'TARJETA CREDITO
    '***************************
   
    
    sql.sql = "SELECT co.local,co.totalarendir-co.totalrendido,co.codigocajera,co.glosa,redondeo "
    sql.sql = sql.sql + "FROM rc_rendicionesdecaja as co inner join " + clientesistema + "gestion.g_maestroempresas as me on (me.codigo=co.local) "
    sql.sql = sql.sql + "WHERE co.fecha = '" & Format(fecha, "yyyy-mm-dd") & "' AND me.codigocontable = '" & localfiltro & "' "
    sql.sql = sql.sql + "and co.totalarendir-co.totalrendido<0  order by local"
    sql.Execute
    
    If sql.RowsAffected > 0 Then
        Set resultados = sql.OpenResultset
        While resultados.EOF = False
    lineascomprobante = lineascomprobante + 1
    LINEA = Format(lineascomprobante, "000")
    fecha = Format(fecha, "yyyy-mm-dd")
    cuenta = "11250005"
    CRCC = ""
    monto = resultados(1) * -1
    glosa = "D/C" + resultados(3)
    Call crearcajera(resultados(2))
    
    Select Case resultados(0)
        Case "41"
            If resultados(2) = "0775753404" Then
                cuenta = "11100007"
                cajera = "0001040006"
                glosa = "D/C VUELTOS C.A."
                Call grabarcomprobante_lineas("CC", numero, LINEA, fecha, cuenta, cuenta, cajera, CRCC, glosa, "CC", numero, fecha, fecha, monto + resultados(4), "H", USUARIOSISTEMA, Format(fecha, "MM"), Format(fecha, "yyyy"), Format(fechasistema, "yyyy-mm-dd"), Time, "")
            Else
                Call grabarcomprobante_lineas("CC", numero, LINEA, fecha, cuenta, cuenta, resultados(2), CRCC, glosa, "CC", numero, fecha, fecha, monto, "H", USUARIOSISTEMA, Format(fecha, "MM"), Format(fecha, "yyyy"), Format(fechasistema, "yyyy-mm-dd"), Time, "")
            End If
        Case "60"
            If resultados(2) = "0775753404" Then
                cuenta = "11100007"
                cajera = "0001060007"
                 glosa = "D/C VUELTOS C.A."
                Call grabarcomprobante_lineas("CC", numero, LINEA, fecha, cuenta, cuenta, cajera, CRCC, glosa, "CC", numero, fecha, fecha, monto + resultados(4), "H", USUARIOSISTEMA, Format(fecha, "MM"), Format(fecha, "yyyy"), Format(fechasistema, "yyyy-mm-dd"), Time, "")
            Else
                Call grabarcomprobante_lineas("CC", numero, LINEA, fecha, cuenta, cuenta, resultados(2), CRCC, glosa, "CC", numero, fecha, fecha, monto, "H", USUARIOSISTEMA, Format(fecha, "MM"), Format(fecha, "yyyy"), Format(fechasistema, "yyyy-mm-dd"), Time, "")
            End If
        Case Else
            Call grabarcomprobante_lineas("CC", numero, LINEA, fecha, cuenta, cuenta, resultados(2), CRCC, glosa, "CC", numero, fecha, fecha, monto, "H", USUARIOSISTEMA, Format(fecha, "MM"), Format(fecha, "yyyy"), Format(fechasistema, "yyyy-mm-dd"), Time, "")
        End Select
        
        resultados.MoveNext
        Wend
     
    End If
End Sub


Sub CARGADIFERENCIASDECAJAPOSITIVA(fecha)
    Dim sql As New rdoQuery
    Dim resultados As rdoResultset
    Set sql.ActiveConnection = teso
    Dim suma As Double
    
    
    
    
    '****************************
    'TARJETA CREDITO
    '***************************
   
    
    sql.sql = "SELECT co.local,sum(co.totalarendir)-sum(co.totalrendido),co.codigocajera,co.glosa "
    sql.sql = sql.sql + "FROM rc_rendicionesdecaja as co inner join " + clientesistema + "gestion.g_maestroempresas as me on (me.codigo=co.local) "
    sql.sql = sql.sql + "WHERE co.fecha = '" & Format(fecha, "yyyy-mm-dd") & "' AND me.codigocontable = '" & localfiltro & "' "
    sql.sql = sql.sql + "and co.totalarendir-co.totalrendido<0  group by local order by local"
    sql.Execute
    
    If sql.RowsAffected > 0 Then
        Set resultados = sql.OpenResultset
        While resultados.EOF = False
    lineascomprobante = lineascomprobante + 1
    LINEA = Format(lineascomprobante, "000")
    fecha = Format(fecha, "yyyy-mm-dd")
    cuenta = "35200011"
    CRCC = leerdatoslocal(resultados(0), "codigocrcc")
    monto = resultados(1) * -1
    glosa = "INGRESOS X DIFERENCIA CAJA "
    Call grabarcomprobante_lineas("CC", numero, LINEA, fecha, cuenta, "", "", CRCC, glosa, "CC", numero, fecha, fecha, monto, "H", USUARIOSISTEMA, Format(fecha, "MM"), Format(fecha, "yyyy"), Format(fechasistema, "yyyy-mm-dd"), Time, "")
            
        resultados.MoveNext
        Wend
     
    End If
End Sub

Sub CARGAVALESDEcredito(fecha)
    Dim sql As New rdoQuery
    Dim resultados As rdoResultset
    Set sql.ActiveConnection = teso
    Dim suma As Double
    
    
    
    
    '****************************
    'TARJETA vales eltit
    '***************************
   
    
    sql.sql = "SELECT co.local,co.numero,co.monto,co.rut "
    sql.sql = sql.sql + "FROM rc_otroscreditos as co inner join " + clientesistema + "gestion.g_maestroempresas as me on (me.codigo=co.local) "
    sql.sql = sql.sql + "WHERE co.fecha = '" & Format(fecha, "yyyy-mm-dd") & "' AND me.codigocontable = '" & localfiltro & "' "
    sql.sql = sql.sql + "order by co.local"
    sql.Execute
    suma = 0
    If sql.RowsAffected > 0 Then
        Set resultados = sql.OpenResultset
        While resultados.EOF = False
    lineascomprobante = lineascomprobante + 1
    LINEA = Format(lineascomprobante, "000")
    fecha = Format(fecha, "yyyy-mm-dd")
    cuenta = "11500190"
    CRCC = ""
    
    monto = resultados(2)
    suma = suma + monto
    glosa = "CONTABILIZA CREDITO " + leerdatoslocal(resultados(0), "nombre")
    Call grabarcomprobante_lineas("CC", numero, LINEA, fecha, cuenta, cuenta, resultados(3), CRCC, glosa, "VC", resultados(1), fecha, fecha, monto, "D", USUARIOSISTEMA, Format(fecha, "MM"), Format(fecha, "yyyy"), Format(fechasistema, "yyyy-mm-dd"), Time, "")
        
        resultados.MoveNext
        Wend
     
    End If
    
End Sub

Sub CARGACASAS(fecha)
    Dim sql As New rdoQuery
    Dim resultados As rdoResultset
    Set sql.ActiveConnection = teso
    Dim suma As Double
    Dim ru As String
    
    
    
    
    '****************************
    'TARJETA vales eltit
    '***************************
   
    
    sql.sql = "SELECT co.local,co.tipo,co.monto "
    sql.sql = sql.sql + "FROM rc_tarjetascasascomerciales as co inner join " + clientesistema + "gestion.g_maestroempresas as me on (me.codigo=co.local) "
    sql.sql = sql.sql + "WHERE co.fecha = '" & Format(fecha, "yyyy-mm-dd") & "' AND me.codigocontable = '" & localfiltro & "' "
    sql.sql = sql.sql + "order by co.local"
    sql.Execute
    
    If sql.RowsAffected > 0 Then
        Set resultados = sql.OpenResultset
        While resultados.EOF = False
    lineascomprobante = lineascomprobante + 1
    LINEA = Format(lineascomprobante, "000")
    fecha = Format(fecha, "yyyy-mm-dd")
    cuenta = "11100011"
    CRCC = ""
    monto = resultados(2)
    ru = leerdatos(teso, clientesistema + "ventas.sv_tiposdepagoclientes", "rut", "codigo='" + resultados(1) + "'")
    glosa = "CONTABILIZA CASA COMERCIAL " & leerdatos(teso, clientesistema + "ventas.sv_tiposdepagoclientes", "nombre", "codigo='" + resultados(1) + "'")
    
    If fecha >= "2017-03-01" Then
        If resultados(1) = "18" Then
            cuenta = "47150008"
            glosa = "VALE COLACION " & resultados(0)
            CRCC = leerdatoslocal(resultados(0), "codigocrcc")
        Else
            cuenta = "11100011"
        End If
    End If
    
    Call grabarcomprobante_lineas("CC", numero, LINEA, fecha, cuenta, cuenta, ru, CRCC, glosa, "TC", numero, fecha, fecha, monto, "D", USUARIOSISTEMA, Format(fecha, "MM"), Format(fecha, "yyyy"), Format(fechasistema, "yyyy-mm-dd"), Time, "")
            
        resultados.MoveNext
        Wend
     
    End If
End Sub

Private Sub botonmisaccesos_Click()
programafiltro = Me.Caption
misaccesos.Show
End Sub


Private Sub botonmisfavoritos_Click()
Call AgregaFavoritos(USUARIOSISTEMA, App.EXEName, Me.Name, Me.Caption)
End Sub

