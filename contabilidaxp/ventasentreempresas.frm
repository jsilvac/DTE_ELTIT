VERSION 5.00
Object = "{4777436C-EB8C-4596-98A8-EBCF98683969}#1.0#0"; "FlexCell.ocx"
Object = "{085FBDF3-00C5-421B-B762-1D57299A2B89}#1.0#0"; "CLBUTN.OCX"
Object = "{ADBBAED6-B16F-40DC-80DF-B44910CBA76C}#1.0#0"; "Frame-Xp.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form ventaentreempresas 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Documentos Emitidos"
   ClientHeight    =   10185
   ClientLeft      =   2130
   ClientTop       =   435
   ClientWidth     =   11610
   DrawStyle       =   5  'Transparent
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   679
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   774
   Begin XPFrame.FrameXp FrameQuickMenu 
      Height          =   615
      Left            =   8160
      TabIndex        =   7
      Top             =   8760
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
         TabIndex        =   9
         Top             =   280
         Width           =   1455
      End
      Begin VB.CommandButton botonmisfavoritos 
         Caption         =   "Mis Favoritos"
         Height          =   255
         Left            =   1800
         TabIndex        =   8
         Top             =   280
         Width           =   1335
      End
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
      TabIndex        =   0
      Top             =   6120
      Width           =   135
   End
   Begin XPFrame.FrameXp FrameXp1 
      Height          =   10185
      Left            =   45
      TabIndex        =   1
      Top             =   0
      Width           =   14865
      _ExtentX        =   26220
      _ExtentY        =   17965
      BackColor       =   16744576
      Caption         =   "REVISOR EMPRESA RELACIONADA"
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
      Begin VB.CommandButton Command2 
         BackColor       =   &H0080FF80&
         Caption         =   "Imprimir"
         Height          =   330
         Left            =   3000
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   8400
         Width           =   2130
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H0080FF80&
         Caption         =   "Generar"
         Height          =   330
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   8400
         Width           =   2130
      End
      Begin XPFrame.FrameXp FrameXp2 
         Height          =   7995
         Left            =   135
         TabIndex        =   2
         Top             =   285
         Width           =   11205
         _ExtentX        =   19764
         _ExtentY        =   14102
         BackColor       =   16744576
         Caption         =   "DOCUMENTOS EMITIDOS"
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
            Height          =   7560
            Left            =   120
            TabIndex        =   3
            Top             =   360
            Width           =   10995
            _ExtentX        =   19394
            _ExtentY        =   13335
            BackColorFixed  =   16744576
            Cols            =   5
            DefaultFontSize =   8.25
            GridColor       =   16711680
            Rows            =   30
         End
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H0080FF80&
         Caption         =   "Excel"
         Height          =   330
         Left            =   5400
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   8400
         Width           =   2130
      End
      Begin XPFrame.FrameXp fechas 
         Height          =   1050
         Left            =   480
         TabIndex        =   10
         Top             =   9120
         Width           =   7005
         _ExtentX        =   12356
         _ExtentY        =   1852
         BackColor       =   14737632
         Caption         =   "Rangos de Fecha"
         CaptionEstilo3D =   1
         BackColor       =   14737632
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
         Begin CoolButtons.cool_Button command8 
            Height          =   375
            Left            =   4950
            TabIndex        =   11
            Top             =   555
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   661
            SkinId          =   "13"
            Caption         =   "Cambia Fecha"
         End
         Begin VB.Label hastafecha 
            BackColor       =   &H00FFC0C0&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   375
            Left            =   2520
            TabIndex        =   15
            Top             =   600
            Width           =   1935
         End
         Begin VB.Label desdefecha 
            BackColor       =   &H00FFC0C0&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   375
            Left            =   360
            TabIndex        =   14
            Top             =   600
            Width           =   1935
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Hasta Fecha"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   2520
            TabIndex        =   13
            Top             =   360
            Width           =   1935
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Desde Fecha"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   360
            TabIndex        =   12
            Top             =   360
            Width           =   1935
         End
      End
      Begin MSComctlLib.ProgressBar BARRA 
         Height          =   285
         Left            =   240
         TabIndex        =   16
         Top             =   8760
         Width           =   7575
         _ExtentX        =   13361
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   1
      End
   End
End
Attribute VB_Name = "ventaentreempresas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private localfiltro As String


Private Sub Command1_Click()
Call leerventaempresa
End Sub



Private Sub COMMAND2_Click()
Grid1.PrintPreview


End Sub



Private Sub Command3_Click()
Grid1.ExportToExcel ("")


End Sub




Private Sub command8_Click()
Call retornofecha(desdefecha, hastafecha)
End Sub

Private Sub Form_Load()
CENTRAR Me
 
Call Conectar_BD


CARGAGRILLA
desdefecha.Caption = "01-" & Format(fechasistema, "mm-yyyy")


hastafecha.Caption = DateSerial(Format(fechasistema, "yyyy"), Format(fechasistema, "mm") + 1, 1 - 1)


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



Private Sub opciones_GotFocus()

MANUAL.SetFocus

End Sub
Sub CARGAGRILLA()
Rem DATOS DE LA COLUMNA
    Dim FORMATOGRILLA(10, 20)
    Grid1.DefaultFont.Size = 8
       
    FORMATOGRILLA(1, 1) = "TIPO"
    FORMATOGRILLA(1, 2) = "NUMERO"
    FORMATOGRILLA(1, 3) = "RUT"
    FORMATOGRILLA(1, 4) = "RAZON SOCIAL"
    FORMATOGRILLA(1, 5) = "TOTAL"
    FORMATOGRILLA(1, 6) = "FECHA EMISION"
    FORMATOGRILLA(1, 7) = "CONTABILIZADA"
    FORMATOGRILLA(1, 8) = "MES/AÑO"
    
    Rem LARGO DE LOS DATOS
    FORMATOGRILLA(2, 1) = "3"
    FORMATOGRILLA(2, 2) = "10"
    FORMATOGRILLA(2, 3) = "13"
    FORMATOGRILLA(2, 4) = "25"
    FORMATOGRILLA(2, 5) = "10"
    FORMATOGRILLA(2, 6) = "10"
    FORMATOGRILLA(2, 7) = "5"
    FORMATOGRILLA(2, 8) = "8"
    

    
    Rem TIPO DE DATOS C=CEROS,S=STRING,N=NUMERICO,D=DATE,
    FORMATOGRILLA(3, 1) = "S"
    FORMATOGRILLA(3, 2) = "S"
    FORMATOGRILLA(3, 3) = "S"
    FORMATOGRILLA(3, 4) = "S"
    FORMATOGRILLA(3, 5) = "N"
    FORMATOGRILLA(3, 6) = "S"
    FORMATOGRILLA(3, 7) = "S"
    FORMATOGRILLA(3, 8) = "S"
    
    Rem FORMATO GRILLA
    FORMATOGRILLA(4, 3) = ""
    FORMATOGRILLA(4, 4) = ""
    FORMATOGRILLA(4, 5) = "##,###,##0"
    FORMATOGRILLA(4, 6) = ""
    FORMATOGRILLA(4, 7) = ""
    
    
    Rem LOCCKED
    For k = 1 To 8
    FORMATOGRILLA(5, k) = "TRUE"
    Next k
        
    
    Grid1.Cols = 9
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
    'Grid1.Column(10).CellType = cellCheckBox

    
End Sub








Sub grababoleta(LINEA)
    Dim netos As Double
    Dim DH As String
    Dim exentos As Double
    Dim TIPOCON As String
    Dim CRCC As String
    
    campos(0, 0) = "fecha"
    campos(1, 0) = "caja"
    campos(2, 0) = "boletainicial"
    campos(3, 0) = "boletafinal"
    campos(4, 0) = "monto"
    campos(5, 0) = "exento"
    campos(6, 0) = "total"
    campos(7, 0) = "centrocosto"
    campos(8, 0) = ""
    
    campos(0, 1) = Format(Grid1.Cell(LINEA, 1).text, "yyyy-mm-dd")
    campos(1, 1) = Grid1.Cell(LINEA, 2).text
    campos(2, 1) = Replace(Grid1.Cell(LINEA, 3).text, ",", ".")
    campos(3, 1) = Replace(Grid1.Cell(LINEA, 4).text, ",", ".")
    campos(4, 1) = Replace(Grid1.Cell(LINEA, 5).text, ",", ".")
    campos(5, 1) = Replace(Grid1.Cell(LINEA, 6).text, ",", ".")
    campos(6, 1) = Replace(Grid1.Cell(LINEA, 7).text, ",", ".")
    campos(7, 1) = Grid1.Cell(LINEA, 8).text
   
    
    
    condicion = ""
    campos(0, 2) = "boletasdeventa"
    op = 2
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    
    
    Call sqlconta.sqlconta(op, condicion)
    

End Sub


Public Function leeboleta(LINEA) As String

    
    campos(0, 0) = "fecha"
    campos(1, 0) = ""
    condicion = "fecha='" + Format(Grid1.Cell(LINEA, 1).text, "yyyy-mm-dd") + "' and caja='" + Grid1.Cell(LINEA, 2).text + "'"
    campos(0, 2) = "boletasdeventa"
    op = 5
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    
    Call sqlconta.sqlconta(op, condicion)
    If sqlconta.status = 0 Then
    leeboleta = "1"
    
    Else
    leeboleta = "0"
    
    End If
    
    

End Function

Sub leectacte(rut)
    campos(0, 0) = "rut"
    campos(1, 0) = "nombre"
    campos(2, 0) = ""
    campos(0, 2) = "cuentascorrientes"
    condicion = "tipo=" + "'" + cuentacliente + "' and rut=" + "'" + rut + "' and año='" + año + "'"
    op = 5
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)
    If sqlconta.status = 4 Then
    Call crearcuentacorriente(rut)
    End If
    
End Sub
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
Private Sub Grid1_AfterReorderColumn(ByVal OriginalPosition As Long, ByVal NewPosition As Long)

End Sub


Private Sub leerventaempresa()
    Dim resultados As rdoResultset
    Dim csql As New rdoQuery
    Set csql.ActiveConnection = contadb
    
    csql.sql = "SELECT tipo,numero,rut,total,fecha FROM " & clientesistema & "conta" & empresaactiva & ". facturasdeventas WHERE " & _
               "FECHA Between '" & Format(desdefecha.Caption, "yyyy-mm-dd") & "' AND '" & Format(hastafecha.Caption, "yyyy-mm-dd") & "' AND rut IN(SELECT " & _
               "CONCAT('0',MID(rut,1,8),MID(rut,10,1)) AS rut FROM eltit_conta.maestroempresas) order by rut,fecha,numero"
        csql.Execute
        LINEA = 0
        Grid1.AutoRedraw = False
        barra.Value = 0
        barra.Max = csql.RowsAffected + 1
        Grid1.Rows = csql.RowsAffected + 1
        If csql.RowsAffected > 0 Then
        Set resultados = csql.OpenResultset
         While Not resultados.EOF
             LINEA = LINEA + 1
        barra.Value = LINEA
     
             Grid1.Cell(LINEA, 1).text = resultados(0)
             Grid1.Cell(LINEA, 2).text = resultados(1)
             Grid1.Cell(LINEA, 3).text = resultados(2)
             Grid1.Cell(LINEA, 4).text = leerdatos(conta, "maestroempresas", "nombre", "mid(rut,1,8)='" + Mid(resultados(2), 2, 8) + "' ")
             Grid1.Cell(LINEA, 5).text = resultados(3)
             Grid1.Cell(LINEA, 6).text = resultados(4)
             Grid1.Cell(LINEA, 8).text = leeregistrocompras(leerdatos(conta, "maestroempresas", "codigoempresa", "mid(rut,1,8)='" + Mid(resultados(2), 2, 8) + "' "), resultados(0), resultados(1), resultados(2))
             If Grid1.Cell(LINEA, 8).text <> "0000-00" Then
                Grid1.Cell(LINEA, 7).text = "SI"
                Grid1.Range(LINEA, 7, LINEA, 7).BackColor = vbGreen
             Else
                Grid1.Cell(LINEA, 7).text = "NO"
                Grid1.Range(LINEA, 7, LINEA, 7).BackColor = vbRed
             End If
             
             
            resultados.MoveNext
         Wend
        End If
      Grid1.AutoRedraw = True
      Grid1.Refresh
        
End Sub

'codigoempresa = leerdatos(conta, "maestroempresas", "codigoempresa", "mid(rut,1,8)='" + Mid(LBLRUT.Caption, 2, 8) + "' ")


Private Function leeregistrocompras(empresa, tipo, numero, rut) As String

    Dim resultados As rdoResultset
    Dim csql As New rdoQuery
    Set csql.ActiveConnection = contadb
    If tipo = "6" Then tipo = "4"
   If tipo = "7" Then tipo = "5"
    If tipo = "8" Then tipo = "6"
     
    csql.sql = "SELECT CONCAT(añocontable,'-', mescontable) FROM " & clientesistema & "conta" & empresa & ".facturasdecompras WHERE numero='" & numero & "' AND rut='" & "0" & Mid(rutempresa, 1, 8) & Mid(rutempresa, 10, 1) & "' AND tipo ='" & tipo & "'"
    csql.Execute
        
        If csql.RowsAffected > 0 Then
        Set resultados = csql.OpenResultset
        leeregistrocompras = resultados(0)
        Else
        leeregistrocompras = "0000-00"
        End If
      
End Function



Private Sub botonmisaccesos_Click()
programafiltro = Me.Caption
misaccesos.Show
End Sub


Private Sub botonmisfavoritos_Click()
Call AgregaFavoritos(USUARIOSISTEMA, App.EXEName, Me.Name, Me.Caption)
End Sub
