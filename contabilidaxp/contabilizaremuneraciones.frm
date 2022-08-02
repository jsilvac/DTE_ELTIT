VERSION 5.00
Object = "{4777436C-EB8C-4596-98A8-EBCF98683969}#1.0#0"; "FlexCell.ocx"
Object = "{ADBBAED6-B16F-40DC-80DF-B44910CBA76C}#1.0#0"; "Frame-Xp.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form proceso06 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Centralizacion de Remuneraciones"
   ClientHeight    =   8850
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   14970
   DrawStyle       =   5  'Transparent
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   Moveable        =   0   'False
   ScaleHeight     =   590
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   998
   Begin XPFrame.FrameXp FrameQuickMenu 
      Height          =   615
      Left            =   11760
      TabIndex        =   15
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
      Begin VB.CommandButton botonmisfavoritos 
         Caption         =   "Mis Favoritos"
         Height          =   255
         Left            =   1800
         TabIndex        =   17
         Top             =   280
         Width           =   1335
      End
      Begin VB.CommandButton botonmisaccesos 
         Caption         =   "Permisos Modulo"
         Height          =   255
         Left            =   120
         TabIndex        =   16
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
      Height          =   8850
      Left            =   45
      TabIndex        =   1
      Top             =   45
      Width           =   14865
      _ExtentX        =   26220
      _ExtentY        =   15610
      BackColor       =   16744576
      Caption         =   "CENTRALIZACION DE  REMUNERACIONES"
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
         Left            =   720
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   8280
         Width           =   2535
      End
      Begin VB.CommandButton Command5 
         BackColor       =   &H0080FF80&
         Caption         =   "GENERA CONTABILIZACION"
         Height          =   330
         Left            =   6960
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   8280
         Width           =   2535
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H0080FF80&
         Caption         =   "IMPRIMIR"
         Height          =   330
         Left            =   4080
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   8280
         Width           =   2130
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
         Begin VB.CheckBox Check1 
            BackColor       =   &H00FF8080&
            Caption         =   "Click Todos"
            Height          =   240
            Left            =   13050
            TabIndex        =   13
            Top             =   765
            Width           =   1410
         End
         Begin VB.CommandButton Command2 
            Caption         =   "LISTAR"
            Height          =   285
            Left            =   11745
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
         Caption         =   "Empresas a Centralizar"
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
            Left            =   120
            TabIndex        =   12
            Top             =   6300
            Width           =   14445
            _ExtentX        =   25479
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   1
         End
         Begin FlexCell.Grid Grid1 
            Height          =   5955
            Left            =   0
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
      End
   End
End
Attribute VB_Name = "proceso06"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private totalmutual As Double

Private localfiltro As String
Private cuentas(20) As String
Private LINEAS As Double
Private nombretrabajador As String
Private MONTOINVALIDEZ As Double
Private CCAF As Double
Private MUTUAL As Double
Private PORCENTAJEMUTUAL As Double
Private montodebe As Double
Private montohaber As Double
Private numero As String
Private DIFE As Double



Private Sub BTELIMINA_Click()



año = COMBOAÑO.text
MES = Format(COMBOMES.ListIndex + 1, "00")

If estacerrado(año + "-" + MES + "-" + Format(fechasistema, "dd")) <> True Then
    If Verifica_Permiso(ingreso01.Caption, "elimina") = True Then
        For k = 1 To Grid1.Rows - 1
            If Val(Grid1.Cell(k, 7).text) = 1 Then
            Call eliminacomprobantesmasivos("CR", MES, año, Grid1.Cell(k, 1).text)
            End If
        Next k
        COMMAND2_Click
    Else
        MsgBox mensaje_nopermiso
    End If
Else
    MsgBox "MES YA CERRADO "
End If




End Sub

Private Sub Command5_Click()
Dim DH As String

año = COMBOAÑO.text
MES = Format(COMBOMES.ListIndex + 1, "00")

If estacerrado(año + "-" + MES + "-" + Format(fechasistema, "dd")) <> True Then
    For k = 1 To Grid1.Rows - 1
    If Grid1.Cell(k, 6).text = "1" And Grid1.Cell(k, 5).text = "0" Then
    Call CONTABILIZAEMPRESAS(Grid1.Cell(k, 1).text, Format(COMBOMES.ListIndex + 1, "00"), COMBOAÑO.text)
    DIFE = montodebe - montohaber
    If DIFE > -100 And DIFE < 100 And DIFE <> 0 Then
      If DIFE < 0 Then
      DIFE = DIFE * -1
        DH = "D"
      Else
        DH = "H"
      
      End If
      
      LINEAS = LINEAS + 1
                        glosa = "AJUSTE CENTRALIZACION "
                        Call grabarcomprobante_lineas("CR", numero, LINEAS, Format(fechasistema, "yyyy-mm-dd"), "47150016", "", "", "0101", glosa, "CR", numero, fecha, fecha, DIFE, DH, USUARIOSISTEMA, MES, año, Date, Time, "", Grid1.Cell(k, 1).text, "", "")
    
    
    End If
    
    
    
    End If
    
    Next k
    
    
    leeempresas
Else
    MsgBox "MES YA CERRADO "

End If


End Sub

Private Sub Check1_Click()
For k = 1 To Grid1.Rows - 1
If Grid1.Cell(k, 5).text <> "1" Then
If Grid1.Cell(k, 4).text <> "" Then
Grid1.Cell(k, 6).text = Check1.Value
End If

End If

Next k

End Sub

Private Sub Command1_Click()
imprimir
End Sub



Private Sub COMMAND2_Click()
año = COMBOAÑO.text
MES = COMBOMES.ListIndex + 1
leeempresas


End Sub



Private Sub Command3_Click()
Dim k As Integer


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
       
    FORMATOGRILLA(1, 1) = "LOCAL"
    FORMATOGRILLA(1, 2) = "EMPRESA"
    FORMATOGRILLA(1, 3) = "TRABAJADORES"
    FORMATOGRILLA(1, 4) = "LIQUIDO PAGO"
    FORMATOGRILLA(1, 5) = "CONTABILIZADA"
    FORMATOGRILLA(1, 6) = "CONTABILIZAR"
    FORMATOGRILLA(1, 7) = "ELIMINAR"
    FORMATOGRILLA(1, 8) = "DIFE"
    
    Rem LARGO DE LOS DATOS
    FORMATOGRILLA(2, 1) = "3"
    FORMATOGRILLA(2, 2) = "30"
    FORMATOGRILLA(2, 3) = "10"
    FORMATOGRILLA(2, 4) = "10"
    FORMATOGRILLA(2, 5) = "12"
    FORMATOGRILLA(2, 6) = "12"
    FORMATOGRILLA(2, 7) = "12"
    FORMATOGRILLA(2, 8) = "10"
    
    
    Rem TIPO DE DATOS C=CEROS,S=STRING,N=NUMERICO,D=DATE,
    FORMATOGRILLA(3, 1) = "S"
    FORMATOGRILLA(3, 2) = "S"
    FORMATOGRILLA(3, 3) = "N"
    FORMATOGRILLA(3, 4) = "N"
    FORMATOGRILLA(3, 5) = "S"
    FORMATOGRILLA(3, 6) = "S"
    FORMATOGRILLA(3, 7) = "S"
    FORMATOGRILLA(3, 8) = "N"
    
    Rem FORMATO GRILLA
    FORMATOGRILLA(4, 1) = ""
    FORMATOGRILLA(4, 2) = ""
    FORMATOGRILLA(4, 3) = ""
    FORMATOGRILLA(4, 4) = ""
    FORMATOGRILLA(4, 5) = ""
    FORMATOGRILLA(4, 6) = ""
    FORMATOGRILLA(4, 7) = ""
    FORMATOGRILLA(4, 8) = ""
    
    Rem LOCCKED
    For k = 1 To 5
    FORMATOGRILLA(5, k) = "TRUE"
    Next k
    FORMATOGRILLA(6, k) = "FALSE"
    
    Grid1.Cols = 9
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
    Grid1.Column(7).CellType = cellCheckBox
    
    
    
End Sub



Private Sub monto_Click()
End Sub

Sub limpiar()


End Sub







'
'Private Sub Grid1_Click(ByVal Row As Long, ByVal Col As Long)
'If Grid1.Cell(GRDI1.ActiveCell.Row, 5).text <> "1" Then
'If Grid1.Cell(Grid1.ActiveCell.Row, 4).text <> "" Then
'If Grid1.Cell(Grid1.ActiveCell.Row, 6).text = "0" Then
'Grid1.Cell(Grid1.ActiveCell.Row, 6).text = "1"
'Else
'Grid1.Cell(Grid1.ActiveCell.Row, 6).text = "0"
'
'End If
'
'End If
'
'End If
'
'End Sub

Public Function LEEDOCUMENTO(cuenta, fecha, tipo, monto, DH) As Boolean

    
    campos(0, 0) = "tipo"
    campos(1, 0) = ""
    campos(2, 0) = ""
    condicion = "codigocuenta='" + cuenta + "' and fecha='" & Format(fecha, "yyyy-mm-dd") & "' and tipodocumento='" + tipo + "' and monto='" & monto & "' and dh='" + DH + "' and tipo='CV' "
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


Public Function LEERULTIMOFOLIO(MES, año) As String

    Dim resultados As rdoResultset
    Dim csql As New rdoQuery
    
        Set csql.ActiveConnection = contadb

            csql.sql = "select IFNULL(max(numero),0) from movimientoscontables where mes = '" & Format(MES, "00") & "' AND año = '" & año & "' and tipo='CR' "
            
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
Sub grabarcomprobante_lineas(tipo, numero, LINEA, fecha, codigocuenta, tipoctacte, rutctacte, centrocosto, glosacontable, tipodocumento, numerodocumento, fechadocumento, fechavencimiento, monto, DH, creadopor, MES, año, fechacreacion, horacreacion, rutproveedor, empresa, cuenta_presupuesto, centro_gastos)
    Dim condicion As String
    Dim campos(40, 3) As String
    Dim op As Integer
    Dim TIPOCON As String
    Dim tipo2 As String
    Dim j As Integer
    Dim lin As String
    Dim lar As Integer
    
    If monto < 0 Then
    monto = monto * -1
    If DH = "D" Then DH = "H" Else DH = "D"
    End If
    
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
    campos(21, 0) = "cuenta_presupuesto"
    campos(22, 0) = "centro_gastos"
    campos(23, 0) = ""
    
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
    campos(21, 1) = cuenta_presupuesto
    campos(22, 1) = centro_gastos

    campos(0, 2) = clientesistema + "conta" + empresa + ".movimientoscontables"
   

    op = 2
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    
    Call sqlconta.sqlconta(op, condicion)
   'Call ACTUALIZADOCUMENTO("+")
   If rutctacte <> "" Then
   Call existerut(año, codigocuenta, rutctacte, empresa)
   
   End If
   
   If DH = "D" Then
   montodebe = montodebe + monto
   Else
   montohaber = montohaber + monto
   
   End If
   
End Sub



Sub CONTABILIZAEMPRESAS(empresa, MES, año)
    Dim resultados As rdoResultset
    Dim csql As New rdoQuery
    Dim rut As String
    Dim cuenta As String
    Dim montoremuneracion As Double
    Dim aportepatronal As Double
    Dim MONTOSINISAPRE As Double
    Dim rutfiltro As String
    
    
    Dim fechas As String
    
    rutfiltro = ""
    
    
    montoremuneracion = LEERMONTO(empresa, MES, año, "IRE02", rutfiltro)
    MONTOSINISAPRE = LEERMONTO(empresa, MES, año, "SINIS", rutfiltro)
    montodebe = 0: montohaber = 0
    MUTUAL = LEETablaempresa(empresa, "mutual")
    CCAF = LEETablaempresa(empresa, "ccaf")
If Format(fechasistema, "yyyy-mm-dd") > "2017-12-31" Then
    PORCENTAJEMUTUAL = LEETablaempresa(empresa, "porcentaje_mutual") + 0.93
Else
    PORCENTAJEMUTUAL = LEETablaempresa(empresa, "porcentaje_mutual") + 0.95
End If
        Set csql.ActiveConnection = contadb
        
        'ariel agrega elimina basura de calculoliquidaciones y liquidacionhd de trabajadores ELIMINADOS
        csql.sql = "DELETE FROM " + clientesistema + "remu" + empresa + ".calculoliquidaciones WHERE mes='" + MES + "' AND año='" + año + "' AND rut NOT IN "
        csql.sql = csql.sql + "(SELECT rut FROM " + clientesistema + "remu" + empresa + ".mt_fijo WHERE mes='" + MES + "' AND año='" + año + "' )"
        csql.Execute
        csql.sql = "DELETE FROM " + clientesistema + "remu" + empresa + ".liquidacionhd WHERE mes='" + MES + "' AND año='" + año + "' AND rut NOT IN "
        csql.sql = csql.sql + "(SELECT rut FROM " + clientesistema + "remu" + empresa + ".mt_fijo WHERE mes='" + MES + "' AND año='" + año + "' )"
        csql.Execute
        

        
        csql.sql = "select IfNULL(ac.contable,'23200010'),sum(rs.monto),IfNULL(ac.nombre,''),IfNULL(ac.dh,'H'),rs.codigohd,rs.glosa "
        csql.sql = csql.sql + "from " + clientesistema + "remu" + empresa + ".calculoliquidaciones as rs left join "
        csql.sql = csql.sql + clientesistema + "remu.asiento_contable as ac on ac.codigo=rs.codigohd "
        csql.sql = csql.sql + "where rs.mes='" + MES + "' and rs.año='" + año + "' and (ac.contable<>'' or mid(rs.codigohd,1,3)='I00' OR mid(rs.codigohd,1,3)='P00' or mid(rs.codigohd,1,2)='A0' or mid(rs.codigohd,1,2)='ST' OR mid(rs.codigohd,1,2)='SE' or mid(rs.codigohd,1,2)='B0' or mid(rs.codigohd,1,2)='CA' or mid(rs.codigohd,1,1)='W' or mid(rs.codigohd,1,2)='BP') "
'        csql.sql = csql.sql & " and mid(rs.codigohd,1,2)='BP' "
        If rutfiltro <> "" Then
            csql.sql = csql.sql & " and rs.rut='" & rutfiltro & "' "
        End If
        
        csql.sql = csql.sql + "group by ac.contable,rs.codigohd "
        csql.Execute
        numero = LEERULTIMOFOLIO(MES, año)
        LINEAS = 0
        MONTOINVALIDEZ = 0
        If csql.RowsAffected > 0 Then
            Set resultados = csql.OpenResultset
            While Not resultados.EOF
                fecha = Format(fechasistema, "yyyy-mm-dd")
         If IsNull(resultados(0)) = True Then
         cuenta = ""
         Else
         
         cuenta = resultados(0)
         
         End If
         If IsNull(resultados(2)) = True Then
         glosa = ""
         Else
         
                  glosa = resultados(2)
         
         
         End If
         
        
        If leerNombreCuentaMayorempresa(cuenta, 1, empresa) = "" Or cuenta = "23200010" Or cuenta = "" Then
                  LINEAS = LINEAS + 1
                  rut = ""
                            If cuenta = "" Then
                            cuenta = "23200010"
                            End If
                    If Mid(resultados(4), 1, 1) = "A" Then
                    rut = leerdatos(contadb, clientesistema + "remu.glosas", "rut2", "codigotg='0004' and codigo='" + Mid(resultados(4), 2, 4) + "'")
                            If glosa = "" Then
                            glosa = "IMPOSICIONES " + resultados(5)
                            End If
                    End If
                    If Mid(resultados(4), 1, 1) = "P" Then
                    rut = leerdatos(contadb, clientesistema + "remu.glosas", "rut2", "codigotg='0004' and codigo='" + Mid(resultados(4), 2, 4) + "'")
                            If glosa = "" Then
                            glosa = "A.P.V " + resultados(5)
                            End If
                    End If
                    
                    
                    If Mid(resultados(4), 1, 1) = "W" Then
                    rut = leerdatos(contadb, clientesistema + "remu.glosas", "rut2", "codigotg='0004' and codigo='" + Mid(resultados(4), 2, 4) + "'")
                            If glosa = "" Then
                            glosa = resultados(5)
                            End If
                    End If
                    
                    If Mid(resultados(4), 1, 2) = "CA" Or Mid(resultados(4), 1, 2) = "VC" Then
                    rut = "0818268009"
                            If glosa = "" Then
                            glosa = resultados(5)
                            End If
                    End If
                    
                    If Mid(resultados(4), 1, 1) = "I" Then
                            If glosa = "" Then
                            glosa = "IMPOSICIONES " + resultados(5)
                            
                            End If
                            
                            rut = leerdatos(contadb, clientesistema + "remu.glosas", "rut2", "codigotg='0008' and codigo='" + Mid(resultados(4), 2, 4) + "'")
                    End If
                    If Mid(resultados(4), 1, 2) = "ST" Then
                            If glosa = "" Then
                            glosa = resultados(5)
                            End If
                            
                            rut = leerdatos(contadb, clientesistema + "remu.glosas", "rut2", "codigotg='0004' and codigo='0" + Mid(resultados(4), 3, 3) + "'")
                            If resultados(4) = "SE13" Then
                            rut = "0980004007"
                            End If
                            
                    End If
              
                    If Mid(resultados(4), 1, 2) = "SE" Then
                            If glosa = "" Then
                            glosa = resultados(5)
                            End If
                            
                            rut = leerdatos(contadb, clientesistema + "remu.glosas", "rut2", "codigotg='0004' and codigo='00" + Mid(resultados(4), 3, 2) + "'")
                            If resultados(4) = "SE13" Then
                            rut = "0980004007"
                            End If
                            
                    End If
              
                    If Mid(resultados(4), 1, 2) = "B0" Then
                            If glosa = "" Then
                            glosa = "INVALIDEZ " + resultados(5)
                            
                            End If
                            
                            rut = leerdatos(contadb, clientesistema + "remu.glosas", "rut2", "codigotg='0004' and codigo='" + Mid(resultados(4), 2, 4) + "'")
                            MONTOINVALIDEZ = MONTOINVALIDEZ + resultados(1)
                    End If
              
                    
              
                If Mid(cuenta, 1, 1) <> "4" Then
                    Call grabarcomprobante_lineas("CR", numero, LINEAS, fecha, cuenta, "", rut, "", glosa, "CR", numero, fecha, fecha, resultados(1), resultados(3), USUARIOSISTEMA, MES, año, Date, Time, "", empresa, "", "")
                    Else
                    Call CONTABILIZAgastos(empresa, MES, año, cuenta, numero, resultados(4), rutfiltro)
                End If
                
        Else
        
            Call CONTABILIZArut(empresa, MES, año, resultados(0), numero, resultados(4), rutfiltro)
            
            
        End If
        
                
                resultados.MoveNext
            Wend
            resultados.Close
            Set resultados = Nothing
        End If
  
    Call CONTABILIZAgastos2(empresa, MES, año, cuenta, numero, "", rutfiltro)
PASO3:

  
  aportepatronal = Round(montoremuneracion * PORCENTAJEMUTUAL / 100)
  
  
  If aportepatronal <> 0 And Format(fechasistema, "yyyy-mm-dd") < "2018-08-01" Then
  LINEAS = LINEAS + 1
                    
                    glosa = "APORTE PATRONAL MUTUAL"
                    
                    Call grabarcomprobante_lineas("CR", numero, LINEAS, fecha, "47150016", "", "", "0101", glosa, "CR", numero, fecha, fecha, aportepatronal, "D", USUARIOSISTEMA, MES, año, Date, Time, "", empresa, "", "")
  
  LINEAS = LINEAS + 1
                    Call grabarcomprobante_lineas("CR", numero, LINEAS, fecha, "23200010", "", "0702851009", "", glosa, "CR", numero, fecha, fecha, aportepatronal, "H", USUARIOSISTEMA, MES, año, Date, Time, "", empresa, "", "")
  
  End If
  
  Rem nueva formula
  
  If aportepatronal <> 0 And Format(fechasistema, "yyyy-mm-dd") > "2018-07-31" Then
  LINEAS = LINEAS + 1
                    
                    Call CONTABILIZAgastosmutual(empresa, MES, año, cuenta, numero, "", rutfiltro)
                    
                    
                    glosa = "APORTE PATRONAL MUTUAL"
  LINEAS = LINEAS + 1
                    Call grabarcomprobante_lineas("CR", numero, LINEAS, fecha, "23200010", "", "0702851009", "", glosa, "CR", numero, fecha, fecha, totalmutual, "H", USUARIOSISTEMA, MES, año, Date, Time, "", empresa, "", "")
  
  End If
  
  
  
  
  
  If CCAF <> 0 Then
  LINEAS = LINEAS + 1
  aportepatronal = Round(MONTOSINISAPRE * 0.6 / 100)
                    glosa = "CCAF APORTE "
                    Call grabarcomprobante_lineas("CR", numero, LINEAS, fecha, "23200010", "", "0818268009", "", glosa, "CR", numero, fecha, fecha, aportepatronal, "H", USUARIOSISTEMA, MES, año, Date, Time, "", empresa, "", "")
  
  End If
  
  If CCAF <> 0 Then
  LINEAS = LINEAS + 1
  aportepatronal = Round(MONTOSINISAPRE * 6.4 / 100)
                    glosa = "SALUD FONASA "
                    Call grabarcomprobante_lineas("CR", numero, LINEAS, fecha, "23200010", "", "0616030000", "", glosa, "CR", numero, fecha, fecha, aportepatronal, "H", USUARIOSISTEMA, MES, año, Date, Time, "", empresa, "", "")
  
  End If
  If CCAF = 0 Then
  LINEAS = LINEAS + 1
  aportepatronal = Round(MONTOSINISAPRE * 7 / 100)
                    glosa = "SALUD FONASA "
                    Call grabarcomprobante_lineas("CR", numero, LINEAS, fecha, "23200010", "", "0616030000", "", glosa, "CR", numero, fecha, fecha, aportepatronal, "H", USUARIOSISTEMA, MES, año, Date, Time, "", empresa, "", "")
  
  End If
   
   Call CONTABILIZArut2(empresa, MES, año, "11250008", numero, "SE", rutfiltro)
           
End Sub

Sub leeempresas()
    Dim resultados As rdoResultset
    Dim csql As New rdoQuery
    
        Set csql.ActiveConnection = conta
        csql.sql = "SELECT codigoempresa,nombre "
        csql.sql = csql.sql + "FROM maestroempresas where codigoempresa<'70' "
        csql.sql = csql.sql + "ORDER BY codigoempresa "
        csql.Execute
        Grid1.Rows = 1
        If csql.RowsAffected > 0 Then
            Set resultados = csql.OpenResultset
            While Not resultados.EOF
                If leertotalremuneraciones(resultados(0), "CT001", Format(COMBOMES.ListIndex + 1, "00"), COMBOAÑO.text, "") <> 0 Then
                Grid1.Rows = Grid1.Rows + 1
                Grid1.Cell(Grid1.Rows - 1, 1).text = resultados(0)
                Grid1.Cell(Grid1.Rows - 1, 2).text = resultados(1)
                Grid1.Cell(Grid1.Rows - 1, 3).text = Format(leertotalremuneraciones(resultados(0), "CT001", Format(COMBOMES.ListIndex + 1, "00"), COMBOAÑO.text, ""), "###,###,###,###")
                Grid1.Cell(Grid1.Rows - 1, 4).text = Format(leertotalremuneraciones(resultados(0), "LI001", Format(COMBOMES.ListIndex + 1, "00"), COMBOAÑO.text, "1"), "###,###,###,###")
                Grid1.Cell(Grid1.Rows - 1, 5).text = estacontabilizado("CR", Format(COMBOMES.ListIndex + 1, "00"), COMBOAÑO.text, resultados(0))
                
             
                    Grid1.Cell(Grid1.Rows - 1, 8).text = Format(Cuadrado(Mid(resultados(0), 1, 2), Format(fechasistema, "YYYY"), Format(fechasistema, "MM")), "###,###,###,##0")
                    If CDbl(Grid1.Cell(Grid1.Rows - 1, 8).text) <> 0 Then
                        Grid1.Range(Grid1.Rows - 1, 1, Grid1.Rows - 1, Grid1.Cols - 1).BackColor = vbRed
                    End If
            End If
                resultados.MoveNext
            Wend
            resultados.Close
            Set resultados = Nothing
        End If
        
End Sub

Sub CONTABILIZArut(empresa, MES, año, cuenta, numero, codigo, rutfiltro)
    Dim resultados As rdoResultset
    Dim csql As New rdoQuery
    Dim MONTOSEGURO As Double
    
    Dim fechas As String
    If cuenta = "11200004" Then Stop
        Set csql.ActiveConnection = contadb
        csql.sql = "select ac.contable,sum(rs.monto),ac.nombre,ac.dh,rs.codigohd,rs.rut,rs.origen,rs.codigo "
        csql.sql = csql.sql + "from " + clientesistema + "remu" + empresa + ".calculoliquidaciones as rs left join "
        csql.sql = csql.sql + clientesistema + "remu.asiento_contable as ac on ac.codigo=rs.codigohd "
        csql.sql = csql.sql + "where  rs.mes='" + MES + "' and rs.año='" + año + "' and ac.contable='" + cuenta + "' and rs.codigohd='" + codigo + "'  "
        If rutfiltro <> "" Then
            csql.sql = csql.sql & " and rs.rut='" & rutfiltro & "' "
        End If
        csql.sql = csql.sql + "group by rs.rut,rs.origen "
        csql.Execute
        MONTOSEGURO = 0
        If csql.RowsAffected > 0 Then
            Set resultados = csql.OpenResultset
            While Not resultados.EOF
                fecha = Format(fechasistema, "yyyy-mm-dd")
                LINEAS = LINEAS + 1
                nombretrabajador = leerdatostrabajador("nombre", clientesistema + "remu" + empresa + ".mt_fijo", "rut='" + resultados(5) + "'", db)
                Call grabarcomprobante_lineas("CR", numero, LINEAS, fecha, resultados(0), "", resultados(5), "", resultados(2) & " " & nombretrabajador, "CR", numero, fecha, fecha, resultados(1), resultados(3), USUARIOSISTEMA, MES, año, Date, Time, "", empresa, "", "")
                Rem If cuenta = "11250004" Or cuenta = "11250006" Then
                If cuenta = "11250004" And resultados("codigo") = "00039" Then
                LINEAS = LINEAS + 1
                Call grabarcomprobante_lineas("CR", numero, LINEAS, fecha, "23100028", "", resultados(5), "", resultados(2) & " " & nombretrabajador, "CR", numero, fecha, fecha, resultados(1), "D", USUARIOSISTEMA, MES, año, Date, Time, "", empresa, "", "")
                
                End If
                
                resultados.MoveNext
            Wend
                
            resultados.Close
            Set resultados = Nothing
        End If
        
End Sub
Sub CONTABILIZArut2(empresa, MES, año, cuenta, numero, codigo, rutfiltro)
    Dim resultados As rdoResultset
    Dim csql As New rdoQuery
    Dim MONTOSEGURO As Double
    
    Dim fechas As String
    
        Set csql.ActiveConnection = contadb
        csql.sql = "select '11250008',sum(rs.monto),'CONTABILIZA SEGURO CESANTIA' ,'D',rs.codigohd,rs.rut "
        csql.sql = csql.sql + "from " + clientesistema + "remu" + empresa + ".calculoliquidaciones as rs "
        csql.sql = csql.sql + "where rs.mes='" + MES + "' and rs.año='" + año + "' And MID(rs.codigohd,1,2)='SE' "
        If rutfiltro <> "" Then
            csql.sql = csql.sql & " and rs.rut='" & rutfiltro & "' "
        End If
        csql.sql = csql.sql + "group by rs.rut "
        csql.Execute
        MONTOSEGURO = 0
        If csql.RowsAffected > 0 Then
            Set resultados = csql.OpenResultset
            While Not resultados.EOF
                fecha = Format(fechasistema, "yyyy-mm-dd")
                LINEAS = LINEAS + 1
                nombretrabajador = leerdatostrabajador("nombre", clientesistema + "remu" + empresa + ".mt_fijo", "rut='" + resultados(5) + "'", db)
                Call grabarcomprobante_lineas("CR", numero, LINEAS, fecha, resultados(0), "", resultados(5), "", nombretrabajador, "CR", numero, fecha, fecha, resultados(1), resultados(3), USUARIOSISTEMA, MES, año, Date, Time, "", empresa, "", "")
                
                resultados.MoveNext
            Wend
                
            resultados.Close
            Set resultados = Nothing
        End If
        
End Sub

Sub CONTABILIZAgastos(empresa, MES, año, cuenta, numero, codigohd, rutfiltro)
    Dim resultados As rdoResultset
    Dim csql As New rdoQuery
    Dim CRCC As String
    Dim cuenta_resultado As String
    Dim centro_gastos As String
    
    Dim fechas As String
    
        Set csql.ActiveConnection = contadb
        csql.sql = "select ac.contable,sum(rs.monto),ac.nombre,ac.dh,rs.codigohd,rs.rut "
        csql.sql = csql.sql + "from " + clientesistema + "remu" + empresa + ".calculoliquidaciones as rs left join "
        csql.sql = csql.sql + clientesistema + "remu.asiento_contable as ac on ac.codigo=rs.codigohd "
        csql.sql = csql.sql + "where rs.mes='" + MES + "' and rs.año='" + año + "' and ac.contable='" + cuenta + "' and ac.codigo='" + codigohd + " '"
        If rutfiltro <> "" Then
            csql.sql = csql.sql & " and rs.rut='" & rutfiltro & "' "
        End If
        csql.sql = csql.sql + "group by rs.rut "
        csql.Execute
        MONTOINVALIDEZ = 0
        If csql.RowsAffected > 0 Then
            Set resultados = csql.OpenResultset
            While Not resultados.EOF
                fecha = Format(fechasistema, "yyyy-mm-dd")
                LINEAS = LINEAS + 1
                CRCC = leerdatostrabajador("codigog", clientesistema + "remu" + empresa + ".mt_semipermanente", "rut='" + resultados(5) + "' and codigotg='0011' and mes='" + MES + "' and año='" + año + "'", db)
                
                cuenta_resultado = leerdatostrabajador("codigog", clientesistema + "remu" + empresa + ".mt_semipermanente", "rut='" + resultados(5) + "' and codigotg='0002' and mes='" + MES + "' and año='" + año + "'", db)
                centro_gastos = leerdatostrabajador("codigog", clientesistema + "remu" + empresa + ".mt_semipermanente", "rut='" + resultados(5) + "' and codigotg='0003' and mes='" + MES + "' and año='" + año + "'", db)
                nombretrabajador = leerdatostrabajador("nombre", clientesistema + "remu" + empresa + ".mt_fijo", "rut='" + resultados(5) + "'", db)
                                
                        
                If CRCC = "0" Then
                CRCC = "0101"
                Else
                CRCC = leerdatoslocal(Mid(CRCC, 3, 2), "codigocrcc")
                
                End If
                
                Call grabarcomprobante_lineas("CR", numero, LINEAS, fecha, resultados(0), "", "", CRCC, nombretrabajador, "CR", numero, fecha, fecha, resultados(1), resultados(3), USUARIOSISTEMA, MES, año, Date, Time, "", empresa, cuenta_resultado, centro_gastos)
                
                resultados.MoveNext
            Wend
            resultados.Close
            Set resultados = Nothing
        End If
            
End Sub
Sub CONTABILIZAgastos2(empresa, MES, año, cuenta, numero, codigohd, rutfiltro)
    Dim resultados As rdoResultset
    Dim csql As New rdoQuery
    Dim CRCC As String
    Dim cuenta_resultado As String
    Dim centro_gastos As String
    
    Dim fechas As String
    
        Set csql.ActiveConnection = contadb
        csql.sql = "select '47150021',sum(rs.monto),'SEGURO INVALIDEZ','D',rs.codigohd,rs.rut "
        csql.sql = csql.sql + "from " + clientesistema + "remu" + empresa + ".calculoliquidaciones as rs "
        csql.sql = csql.sql + "where rs.mes='" + MES + "' and rs.año='" + año + "' and mid(codigohd,1,2)='B0' "
        If rutfiltro <> "" Then
            csql.sql = csql.sql & " and rs.rut='" & rutfiltro & "' "
        End If
        csql.sql = csql.sql + "group by rs.rut "
        csql.Execute
        MONTOINVALIDEZ = 0
        If csql.RowsAffected > 0 Then
            Set resultados = csql.OpenResultset
            While Not resultados.EOF
                fecha = Format(fechasistema, "yyyy-mm-dd")
                LINEAS = LINEAS + 1
                CRCC = leerdatostrabajador("codigog", clientesistema + "remu" + empresa + ".mt_semipermanente", "rut='" + resultados(5) + "' and codigotg='0011' and mes='" + MES + "' and año='" + año + "'", db)
                
                cuenta_resultado = leerdatostrabajador("codigog", clientesistema + "remu" + empresa + ".mt_semipermanente", "rut='" + resultados(5) + "' and codigotg='0002' and mes='" + MES + "' and año='" + año + "'", db)
                centro_gastos = leerdatostrabajador("codigog", clientesistema + "remu" + empresa + ".mt_semipermanente", "rut='" + resultados(5) + "' and codigotg='0003' and mes='" + MES + "' and año='" + año + "'", db)
                nombretrabajador = leerdatostrabajador("nombre", clientesistema + "remu" + empresa + ".mt_fijo", "rut='" + resultados(5) + "'", db)
                                
                        
                If CRCC = "0" Then
                CRCC = "0101"
                Else
                CRCC = leerdatoslocal(Mid(CRCC, 3, 2), "codigocrcc")
                
                End If
                
                Call grabarcomprobante_lineas("CR", numero, LINEAS, fecha, resultados(0), "", "", CRCC, nombretrabajador, "CR", numero, fecha, fecha, resultados(1), resultados(3), USUARIOSISTEMA, MES, año, Date, Time, "", empresa, cuenta_resultado, centro_gastos)
                
                resultados.MoveNext
            Wend
            resultados.Close
            Set resultados = Nothing
        End If
            
End Sub

Sub CONTABILIZAgastosmutual(empresa, MES, año, cuenta, numero, codigohd, rutfiltro)
    Dim resultados As rdoResultset
    Dim csql As New rdoQuery
    Dim CRCC As String
    Dim cuenta_resultado As String
    Dim centro_gastos As String
    
    Dim fechas As String
                totalmutual = 0
    
        Set csql.ActiveConnection = contadb
        csql.sql = "select '47150016',sum(rs.monto),'APORTE MUTUAL','D',rs.codigohd,rs.rut "
        csql.sql = csql.sql + "from " + clientesistema + "remu" + empresa + ".calculoliquidaciones as rs "
        csql.sql = csql.sql + "where rs.mes='" + MES + "' and rs.año='" + año + "' and codigohd='MUT00' "
        If rutfiltro <> "" Then
            csql.sql = csql.sql & " and rs.rut='" & rutfiltro & "' "
        End If
        csql.sql = csql.sql + "group by rs.rut "
        csql.Execute
        MONTOINVALIDEZ = 0
        If csql.RowsAffected > 0 Then
            Set resultados = csql.OpenResultset
            While Not resultados.EOF
                fecha = Format(fechasistema, "yyyy-mm-dd")
                LINEAS = LINEAS + 1
                CRCC = leerdatostrabajador("codigog", clientesistema + "remu" + empresa + ".mt_semipermanente", "rut='" + resultados(5) + "' and codigotg='0011' and mes='" + MES + "' and año='" + año + "'", db)
                
                cuenta_resultado = leerdatostrabajador("codigog", clientesistema + "remu" + empresa + ".mt_semipermanente", "rut='" + resultados(5) + "' and codigotg='0002' and mes='" + MES + "' and año='" + año + "'", db)
                centro_gastos = leerdatostrabajador("codigog", clientesistema + "remu" + empresa + ".mt_semipermanente", "rut='" + resultados(5) + "' and codigotg='0003' and mes='" + MES + "' and año='" + año + "'", db)
                nombretrabajador = leerdatostrabajador("nombre", clientesistema + "remu" + empresa + ".mt_fijo", "rut='" + resultados(5) + "'", db)
                                
                        
                If CRCC = "0" Then
                CRCC = "0101"
                Else
                CRCC = leerdatoslocal(Mid(CRCC, 3, 2), "codigocrcc")
                
                End If
                totalmutual = totalmutual + resultados(1)
                Call grabarcomprobante_lineas("CR", numero, LINEAS, fecha, resultados(0), "", "", CRCC, nombretrabajador, "CR", numero, fecha, fecha, resultados(1), resultados(3), USUARIOSISTEMA, MES, año, Date, Time, "", empresa, cuenta_resultado, centro_gastos)
                
                resultados.MoveNext
            Wend
            resultados.Close
            Set resultados = Nothing
        End If
            
End Sub



Public Sub existerut(año, tipo, rut, empresa)
 
    campos(0, 0) = "año"
    campos(1, 0) = "tipo"
    campos(2, 0) = "rut"
    campos(3, 0) = "nombre"
    campos(4, 0) = ""
    campos(0, 1) = Format(fechasistema, "yyyy")
    campos(1, 1) = tipo
    campos(2, 1) = rut
    campos(3, 1) = nombretraba(rut, empresa)
    condicion = "tipo='" + tipo + "' and rut='" + rut + "' and año='" + año + "'  "
    campos(0, 2) = clientesistema + "conta" + empresa + ".cuentascorrientes"
    op = 5
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)
    If sqlconta.status = 0 Then
    Else
    
    Call grabar(año, tipo, rut, nombretraba(rut, empresa), empresa)
    
    End If

    
    End Sub
Public Function nombretraba(rut, empresa) As String
    campos(0, 0) = "nombre"
    campos(1, 0) = ""
    condicion = "rut='" + rut + "' "
    campos(0, 2) = clientesistema + "remu" + empresa + ".mt_fijo "
    op = 5
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)
    If sqlconta.status = 0 Then
    nombretraba = sqlconta.response(0, 3)
    Else
    nombretraba = ""
    End If
    End Function

Sub grabar(año, tipo, rut, NOMBRE, empresa)
    campos(0, 0) = "año"
    campos(1, 0) = "tipo"
    campos(2, 0) = "rut"
    campos(3, 0) = "nombre"
    campos(4, 0) = ""
    campos(0, 1) = Format(fechasistema, "yyyy")
    campos(1, 1) = tipo
    campos(2, 1) = rut
    campos(3, 1) = NOMBRE
    
    campos(0, 2) = clientesistema + "conta" + empresa + ".cuentascorrientes"
    condicion = ""
    op = 2
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)
    
     Call grabar2(año, tipo, rut, empresa)
    
    End Sub
Sub grabar2(año, tipo, rut, empresa)
      
    campos(0, 0) = "año"
    campos(1, 0) = "tipo"
    campos(2, 0) = "rut"
    campos(3, 0) = ""
    
    campos(0, 1) = año
    campos(1, 1) = tipo
    campos(2, 1) = rut
    
    campos(0, 2) = clientesistema + "conta" + empresa + ".saldosctacte"
    op = 2
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)
    

End Sub

Public Function trabajadorfiniquitado(rut, MES, año, empresa) As Boolean

    campos(0, 0) = "fecha"
    campos(1, 0) = ""
    condicion = "rut='" + rut + "' and mes='" + MES + "' and año='" + año + "' "
    campos(0, 2) = clientesistema + "remu" + empresa + ".finiquitos "
    op = 5
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)
    If sqlconta.status = 0 Then
    trabajadorfiniquitado = True
    Else
    trabajadorfiniquitado = False
    End If
    End Function

Public Function LEERMONTO(empresa, MES, año, codigohd, rutfiltro) As Double

    Dim resultados As rdoResultset
    Dim csql As New rdoQuery
    
        Set csql.ActiveConnection = contadb
        

        csql.sql = "select SUM(MONTO) from " + clientesistema + "remu" + empresa + ".calculoliquidaciones "
        csql.sql = csql.sql & "where codigohd='" + codigohd + "' and mes='" + MES + "' and año='" + año + "' "
        If rutfiltro <> "" Then
            csql.sql = csql.sql & " and rut='" & rutfiltro & "' "
        End If
        csql.sql = csql.sql & "group by codigohd "
        csql.Execute
        LEERMONTO = 0
        If csql.RowsAffected > 0 Then
            Set resultados = csql.OpenResultset
            LEERMONTO = resultados(0)
            resultados.Close
            Set resultados = Nothing
        End If
            
End Function

Private Sub botonmisaccesos_Click()
programafiltro = Me.Caption
misaccesos.Show
End Sub


Private Sub botonmisfavoritos_Click()
Call AgregaFavoritos(USUARIOSISTEMA, App.EXEName, Me.Name, Me.Caption)
End Sub



Private Function Cuadrado(empresa, año, MES) As Double
        Dim resultados As rdoResultset
        Dim sql As New rdoQuery
        Dim multi As Double
        Dim total As Double
        Dim totalD As Double
        Dim totalH As Double
        Dim tabla As String
        Set sql.ActiveConnection = contadb
        
        tabla = "SELECT codigocuenta,IFNULL(SUM(monto),0) AS monto,dh FROM "
        tabla = tabla & clientesistema & "conta" & empresa & ".movimientoscontables "
        tabla = tabla & " WHERE mes='" & MES & "' AND tipo='CR' AND año='" & año & "'"
        tabla = tabla & " GROUP BY dh "
        sql.sql = tabla
        sql.Execute
       
        Cuadrado = 0
        If sql.RowsAffected > 0 Then
        Set resultados = sql.OpenResultset
            While resultados.EOF = False
                If resultados("dh") = "H" Then totalH = resultados("monto")
                If resultados("dh") = "D" Then totalD = resultados("monto")
                
       
            resultados.MoveNext
            Wend
            
             
                Cuadrado = totalD - totalH
            
        End If
End Function

Private Sub Grid1_DblClick()
Dim row As Double
Dim empresa As String
Dim año As String
Dim MES As String

row = Grid1.ActiveCell.row
año = Format(fechasistema, "yyyy")
MES = Format(fechasistema, "mm")
    If Grid1.ActiveCell.BackColor = vbRed Then
        empresa = Mid(Grid1.Cell(row, 1).text, 1, 2)
        If empresa <> "" Then
            Load listadoCodigosSinCuentas
            Call listadoCodigosSinCuentas.buscaCuentas(empresa, año, MES)
            listadoCodigosSinCuentas.Show 1
        End If
        
    End If

End Sub





