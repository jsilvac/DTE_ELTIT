VERSION 5.00
Object = "{4777436C-EB8C-4596-98A8-EBCF98683969}#1.0#0"; "FlexCell.ocx"
Object = "{ADBBAED6-B16F-40DC-80DF-B44910CBA76C}#1.0#0"; "Frame-Xp.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form contabilizapromotora_castigos 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Centralizacion de creditos y pagos"
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
      Left            =   11760
      TabIndex        =   20
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
         TabIndex        =   14
         Top             =   280
         Width           =   1455
      End
      Begin VB.CommandButton botonmisfavoritos 
         Caption         =   "Mis Favoritos"
         Height          =   255
         Left            =   1680
         TabIndex        =   21
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
      Caption         =   "CENTRALIZACION DE CREDITOS Y PAGOS"
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
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   8160
         Width           =   2535
      End
      Begin VB.CommandButton Command4 
         BackColor       =   &H0080FF80&
         Caption         =   "TRASPASA CONTABILIDAD"
         Height          =   330
         Left            =   6720
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   8160
         Width           =   2535
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H0080FF80&
         Caption         =   "IMPRIMIR"
         Height          =   330
         Left            =   3735
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   8190
         Width           =   2130
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
         Begin XPFrame.FrameXp FrameXp5 
            Height          =   645
            Left            =   7515
            TabIndex        =   16
            Top             =   270
            Width           =   3975
            _ExtentX        =   7011
            _ExtentY        =   1138
            BackColor       =   16761024
            Caption         =   "CONTABILIZA"
            CaptionEstilo3D =   1
            BackColor       =   16761024
            ForeColor       =   65535
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
            Begin VB.OptionButton Option2 
               BackColor       =   &H00FFC0C0&
               Caption         =   "PAGOS"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Left            =   2160
               TabIndex        =   18
               Top             =   315
               Width           =   1320
            End
            Begin VB.OptionButton Option1 
               BackColor       =   &H00FFC0C0&
               Caption         =   "CREDITOS"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Left            =   135
               TabIndex        =   17
               Top             =   315
               Value           =   -1  'True
               Width           =   1635
            End
         End
         Begin VB.CheckBox Check1 
            BackColor       =   &H00FF8080&
            Caption         =   "Click Todos"
            Height          =   240
            Left            =   13050
            TabIndex        =   15
            Top             =   765
            Width           =   1410
         End
         Begin VB.CommandButton Command2 
            Caption         =   "LISTAR"
            Height          =   285
            Left            =   11760
            TabIndex        =   6
            Top             =   405
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
            Left            =   6705
            TabIndex        =   12
            Top             =   270
            Visible         =   0   'False
            Width           =   150
            _ExtentX        =   265
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
               TabIndex        =   13
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
         Begin FlexCell.Grid Grid1 
            Height          =   5955
            Left            =   45
            TabIndex        =   3
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
         Begin MSComctlLib.ProgressBar Barra 
            Height          =   375
            Left            =   0
            TabIndex        =   22
            Top             =   6240
            Width           =   14535
            _ExtentX        =   25638
            _ExtentY        =   661
            _Version        =   393216
            Appearance      =   1
         End
      End
   End
End
Attribute VB_Name = "contabilizapromotora_castigos"
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


Private Sub BTELIMINA_Click()

sqlconta.audit = True
sqlconta.programaactivo = Me.Caption

año = COMBOAÑO.text
MES = Format(COMBOMES.ListIndex + 1, "00")
If estacerrado(año + "-" + MES + "-" + Format(fechasistema, "dd")) <> True Then

If Verifica_Permiso(ingreso01.Caption, "elimina") = True Then
            If Option1.Value = True Then
                Call eliminacomprobantesmasivos("PC", MES, año, empresaactiva)
            Else
                Call eliminacomprobantesmasivos("PP", MES, año, empresaactiva)
            End If

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
If Grid1.Cell(k, 6).text <> "1" Then
Grid1.Cell(k, 5).text = Check1.Value
End If

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
TIPODO = "PC"
lin = 0
lineascomprobante = 0
año = COMBOAÑO.text
MES = Format(COMBOMES.ListIndex + 1, "00")

If estacerrado(año + "-" + MES + "-" + Format(fechasistema, "dd")) <> True Then

For k = 1 To Grid1.Rows - 1
    If Grid1.Cell(k, 5).text = "1" And Grid1.Cell(k, 6).text <> "1" Then
    If Option1.Value = True Then
    Call leerCREDITOS(Grid1.Cell(k, 1).text)
    Else
    Call leerpagos(Grid1.Cell(k, 1).text)
    
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
    empresaactiva = "28"
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
       
    FORMATOGRILLA(1, 1) = "FECHA"
    If Option1.Value = True Then
    FORMATOGRILLA(1, 2) = "CREDITOS"
    FORMATOGRILLA(1, 3) = "CAPITAL"
    FORMATOGRILLA(1, 4) = "INTERESES"
    Else
    FORMATOGRILLA(1, 2) = "PAGOS"
    FORMATOGRILLA(1, 3) = "INTERES MORA"
    FORMATOGRILLA(1, 4) = "TOTAL"
    
    End If
    
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
  
    
        Set csql.ActiveConnection = contadb
        If Option1.Value = True Then
        csql.sql = "SELECT fechacompra,SUM(montocuota*cantidadcuotas),sum(montocredito),sum((montocuota*cantidadcuotas)-montocredito) "
        csql.sql = csql.sql + "from " + clientesistema + "ventas.pl_cuotas_detalle "
        csql.sql = csql.sql + "where fechacompra>='" + fecha1 + "' AND fechacompra<='" + fecha2 + "' and numerocuota='1' and antiguo<>'1' group by fechacompra order by fechacompra "
        Else
        csql.sql = "SELECT fecha,sum(montocuotas),sum(interesesmora),sum(montocuotas+interesesmora) "
        csql.sql = csql.sql + "from " + clientesistema + "ventas.pl_cuotas_pago_cabeza "
        csql.sql = csql.sql + "where fecha>='" + fecha1 + "' AND fecha<='" + fecha2 + "'  group by fecha order by fecha "

        End If
        
        
        
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
             Grid1.Cell(LINEA, 5).text = "0"
             Grid1.Cell(LINEA, 6).text = "0"
            If Option1.Value = True Then
            If LEEDOCUMENTO("11200046", resultados(0), "D", "PC") = True Then
             Grid1.Cell(LINEA, 6).text = "1"
            End If
            End If
            
            If Option1.Value = False Then
            If LEEDOCUMENTO("11200046", resultados(0), "H", "PP") = True Then
             Grid1.Cell(LINEA, 6).text = "1"
            End If
            End If
             
            
            resultados.MoveNext
       
            Wend
End If
      
      Grid1.AutoRedraw = True
      Grid1.Refresh
      
      
      
End Sub
Private Sub leerCREDITOS(fecha)

Dim resultados As rdoResultset
    Dim csql As New rdoQuery
    Dim csql2 As New rdoQuery
    
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
    Dim monto As String
    Dim numero As String
    Dim numerorut As String
    Dim INTERESES As Double
    
    Dim cuenta As String
    
    LINEA = 0: fec = 0: fec1 = 0
    fecha1 = Format(fecha, "yyyy-mm-dd")
    fecha2 = Format(fecha, "yyyy-mm-dd")
    lineascomprobante = 0
    numero = LEERULTIMOFOLIO("PC")
        INTERESES = 0
        Set csql.ActiveConnection = contadb
        csql.sql = "SELECT montocuota,cantidadcuotas,rut,montocuota*cantidadcuotas,local,montocredito,tipo,numero "
        csql.sql = csql.sql + "from " + clientesistema + "ventas.pl_cuotas_detalle "
        csql.sql = csql.sql + "where fechacompra>='" + fecha1 + "' AND fechacompra<='" + fecha2 + "' and numerocuota='1' and antiguo<>'1' order by rut "

        csql.Execute


        If csql.RowsAffected > 0 Then
        Set resultados = csql.OpenResultset

         While Not resultados.EOF
                lineascomprobante = lineascomprobante + 1
                LINEA = Format(lineascomprobante, "000")
                fecha = Format(fecha, "yyyy-mm-dd")
                cuenta = "11200046"
                CRCC = ""
                monto = resultados(3)
                glosa = "CARGA T.PLUS  " & resultados(1) & " x " & Format(resultados(0), "###,###") & " x compra " + leerdatoslocal(resultados(4), "nombre")
                Call existerut(Format(fechasistema, "yyyy"), "11200046", resultados(2))
                Call grabarcomprobante_lineas(TIPODO, numero, LINEA, fecha, cuenta, cuenta, resultados(2), CRCC, glosa, resultados(6), resultados(7), fecha, fecha, monto, "D", USUARIOSISTEMA, Format(fecha, "MM"), Format(fecha, "yyyy"), Format(fechasistema, "yyyy-mm-dd"), Time, "")
                INTERESES = INTERESES + resultados(3)
                resultados.MoveNext

            Wend
End If
          Rem carga separacion locales

csql.Close
        
        Set csql2.ActiveConnection = contadb

        csql2.sql = "SELECT sum(montocredito),local,tipo "
        csql2.sql = csql2.sql + "from " + clientesistema + "ventas.pl_cuotas_detalle "
        csql2.sql = csql2.sql + "where fechacompra>='" + fecha1 + "' AND fechacompra<='" + fecha2 + "' and numerocuota='1' AND antiguo<>'1' group by local,tipo order by local "

        csql2.Execute


        If csql2.RowsAffected > 0 Then
        Set resultados = csql2.OpenResultset
        
         While Not resultados.EOF
       INTERESES = INTERESES - resultados(0)
       lineascomprobante = lineascomprobante + 1
    LINEA = Format(lineascomprobante, "000")
    fecha = Format(fecha, "yyyy-mm-dd")
    
    If resultados(2) = "VR" Then
    cuenta = "11200030"
    numerorut = ""
    monto = resultados(0)
    glosa = "CARGA REPACTACIONES PLUS " + leerdatoslocal(resultados(1), "nombre")
    End If
    
    If resultados(2) <> "VR" And resultados(2) <> "CA" Then
    cuenta = "23100001"
    CRCC = ""
    numerorut = leerdatoslocal(resultados(1), "rut")
    monto = resultados(0)
    glosa = "CARGA DEUDA POR VENTAS PLUS " + leerdatoslocal(resultados(1), "nombre")
    End If
    
    If resultados(2) = "CA" Then
    cuenta = "35200004"
    CRCC = "0103"
    numerorut = ""
    monto = resultados(0)
    glosa = "CARGA COBROS POR SEGURO PLUS " + leerdatoslocal(resultados(1), "nombre")
    End If
    
    If resultados(2) <> "VR" And (resultados(1) = "07" Or resultados(1) = "77") And resultados(2) <> "CA" Then
    cuenta = "11200030"
    numerorut = ""
    monto = resultados(0)
    glosa = "CARGA REPACTACIONES PLUS " + leerdatoslocal(resultados(1), "nombre")
    End If
    
    Call grabarcomprobante_lineas(TIPODO, numero, LINEA, fecha, cuenta, cuenta, numerorut, CRCC, glosa, "PC", numero, fecha, fecha, monto, "H", USUARIOSISTEMA, Format(fecha, "MM"), Format(fecha, "yyyy"), Format(fechasistema, "yyyy-mm-dd"), Time, "")

            resultados.MoveNext

            Wend
End If
'
Rem GRABA LINEA INTERES
If INTERESES <> 0 Then
    lineascomprobante = lineascomprobante + 1
    LINEA = Format(lineascomprobante, "000")
    fecha = Format(fecha, "yyyy-mm-dd")
    cuenta = "35200001"
    CRCC = "0103"
    numerorut = ""
    monto = INTERESES
    glosa = "INTERES POR COMPRA PLUS "
    Call grabarcomprobante_lineas(TIPODO, numero, LINEA, fecha, cuenta, "", "", CRCC, glosa, "PC", numero, fecha, fecha, monto, "H", USUARIOSISTEMA, Format(fecha, "MM"), Format(fecha, "yyyy"), Format(fechasistema, "yyyy-mm-dd"), Time, "")
End If
End Sub
Private Sub leerpagos(fecha)

Dim resultados As rdoResultset
    Dim csql As New rdoQuery
    Dim csql2 As New rdoQuery
    
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
    Dim monto As String
    Dim numero As String
    Dim numerorut As String
    Dim INTERESES As Double
    Dim pagos As Double
    Dim sumador(100) As Double
    Dim codigocontable As String
    
    Dim cuenta As String
    
    LINEA = 0: fec = 0: fec1 = 0
    fecha1 = Format(fecha, "yyyy-mm-dd")
    fecha2 = Format(fecha, "yyyy-mm-dd")
    lineascomprobante = 0
    numero = LEERULTIMOFOLIO("PP")
        For k = 0 To 100
        sumador(k) = 0
        Next k
        
        INTERESES = 0
        Set csql.ActiveConnection = contadb
        csql.sql = "SELECT fecha,rut,sum(montocuotas),sum(interesesmora),numero,local "
        csql.sql = csql.sql + "from " + clientesistema + "ventas.pl_cuotas_pago_cabeza "
        csql.sql = csql.sql + "where fecha>='" + fecha1 + "' AND fecha<='" + fecha2 + "' group by rut order by rut "

        
        csql.Execute


        If csql.RowsAffected > 0 Then
        Set resultados = csql.OpenResultset

         While Not resultados.EOF
    lineascomprobante = lineascomprobante + 1
    LINEA = Format(lineascomprobante, "000")
    fecha = Format(fecha, "yyyy-mm-dd")
    cuenta = "11200046"
    CRCC = ""
    monto = resultados(2)
    glosa = "CANCELA CUOTAS PLUS "
    sumador(resultados(5)) = sumador(resultados(5)) + (resultados(2) + resultados(3))
    Call existerut(Format(fechasistema, "yyyy"), "11200046", resultados(1))
    Call grabarcomprobante_lineas("PP", numero, LINEA, fecha, cuenta, cuenta, resultados(1), CRCC, glosa, "PA", resultados(4), fecha, fecha, monto, "H", USUARIOSISTEMA, Format(fecha, "MM"), Format(fecha, "yyyy"), Format(fechasistema, "yyyy-mm-dd"), Time, "")
    INTERESES = INTERESES + resultados(3)
    pagos = pagos + resultados(2) + resultados(3)
    resultados.MoveNext
    Wend
End If
          Rem carga separacion locales

csql.Close
                
    For k = 0 To 100
  If sumador(k) <> 0 Then
    
    lineascomprobante = lineascomprobante + 1
    LINEA = Format(lineascomprobante, "000")
    fecha = Format(fecha, "yyyy-mm-dd")
    If k = 7 Or k = 77 Then
    cuenta = "11200030"
    CRCC = ""
    numerorut = ""
    Else
    cuenta = "23100001"
    CRCC = ""
    numerorut = leerdatoslocal(Format(k, "00"), "rut")
    
    End If
    
    monto = sumador(k)
    glosa = "CARGA PAGOS CLIENTES PLUS"
    Call grabarcomprobante_lineas("PP", numero, LINEA, fecha, cuenta, "", numerorut, "", glosa, "PC", numero, fecha, fecha, monto, "D", USUARIOSISTEMA, Format(fecha, "MM"), Format(fecha, "yyyy"), Format(fechasistema, "yyyy-mm-dd"), Time, "")

  
  End If
  Next k

'
Rem GRABA LINEA INTERES
If INTERESES <> 0 Then
    lineascomprobante = lineascomprobante + 1
    LINEA = Format(lineascomprobante, "000")
    fecha = Format(fecha, "yyyy-mm-dd")
    cuenta = "35200002"
    CRCC = "0103"
    numerorut = ""
    monto = INTERESES
    glosa = "INTERES POR MORA PLUS "
    Call grabarcomprobante_lineas("PP", numero, LINEA, fecha, cuenta, "", "", CRCC, glosa, "PC", numero, fecha, fecha, monto, "H", USUARIOSISTEMA, Format(fecha, "MM"), Format(fecha, "yyyy"), Format(fechasistema, "yyyy-mm-dd"), Time, "")
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



Private Sub Form_Unload(Cancel As Integer)
'empresaactiva = localfiltro

End Sub

Private Sub Grid1_KeyDown(KeyCode As Integer, ByVal Shift As Integer)

'If KeyCode = 46 Then
'Call eliminafactura(Grid1.Cell(Grid1.ActiveCell.Row, 1).text, Grid1.Cell(Grid1.ActiveCell.Row, 2).text)
'End If
'leer
End Sub



Public Function LEEDOCUMENTO(cuenta, fecha, DH, tipo) As Boolean

    
    campos(0, 0) = "tipo"
    campos(1, 0) = ""
    campos(2, 0) = ""
    
    condicion = "codigocuenta='" + cuenta + "' and fecha='" & Format(fecha, "yyyy-mm-dd") & "'  and dh='" + DH + "' and tipo='" + tipo + "' and linea='1' "
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
    condicion = "tipo='" + tipo + "' and rut='" + rut + "' and año='" + año + "'  "
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
Public Function nombrecliente(rut) As String
    campos(0, 0) = "nombre"
    campos(1, 0) = ""
    condicion = "rut='" + rut + "' "
    campos(0, 2) = clientesistema + "ventas.pl_maestroclientes"
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

Private Sub botonmisaccesos_Click()
programafiltro = Me.Caption
misaccesos.Show
End Sub


Private Sub botonmisfavoritos_Click()
Call AgregaFavoritos(USUARIOSISTEMA, App.EXEName, Me.Name, Me.Caption)
End Sub
