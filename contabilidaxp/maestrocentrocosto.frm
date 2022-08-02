VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "msadodc.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{4777436C-EB8C-4596-98A8-EBCF98683969}#1.0#0"; "FlexCell.ocx"
Object = "{ADBBAED6-B16F-40DC-80DF-B44910CBA76C}#1.0#0"; "Frame-Xp.ocx"
Begin VB.Form maestro03 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Maestro de Centros de Costo"
   ClientHeight    =   8340
   ClientLeft      =   2235
   ClientTop       =   1305
   ClientWidth     =   12855
   DrawStyle       =   5  'Transparent
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   556
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   857
   ShowInTaskbar   =   0   'False
   Begin XPFrame.FrameXp FrameXp3 
      Height          =   7455
      Left            =   6600
      TabIndex        =   10
      Top             =   120
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   13150
      BackColor       =   16744576
      Caption         =   "Cuentas"
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
         Height          =   7095
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   12515
         Cols            =   5
         DefaultFontSize =   8.25
         Rows            =   30
         SelectionMode   =   1
      End
   End
   Begin XPFrame.FrameXp FrameXp2 
      Height          =   4215
      Left            =   120
      TabIndex        =   8
      Top             =   2160
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   7435
      BackColor       =   16744576
      Caption         =   "Saldos"
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
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid SALDOS 
         Height          =   3495
         Left            =   360
         TabIndex        =   9
         Top             =   360
         Width           =   4455
         _ExtentX        =   7858
         _ExtentY        =   6165
         _Version        =   393216
         BackColor       =   16776436
         ForeColor       =   12582912
         Rows            =   13
         Cols            =   4
         FixedCols       =   0
         BackColorFixed  =   16107953
         BackColorSel    =   16777215
         ForeColorSel    =   16744576
         BackColorBkg    =   16776436
         GridColor       =   -2147483635
         GridColorFixed  =   12582912
         GridLinesFixed  =   1
         BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   4
         _Band(0).GridLinesBand=   1
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
      End
   End
   Begin XPFrame.FrameXp FrameXp1 
      Height          =   1695
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   2990
      BackColor       =   16761024
      Caption         =   "Datos"
      CaptionEstilo3D =   1
      BackColor       =   16761024
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
      Begin VB.TextBox dato1 
         BackColor       =   &H00E1FFFD&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   1560
         MaxLength       =   2
         TabIndex        =   5
         Tag             =   "codigo"
         Top             =   480
         Width           =   375
      End
      Begin VB.TextBox dato2 
         BackColor       =   &H00E1FFFD&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   2040
         MaxLength       =   2
         TabIndex        =   4
         Top             =   480
         Width           =   375
      End
      Begin VB.TextBox dato3 
         BackColor       =   &H00E1FFFD&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   1560
         MaxLength       =   50
         TabIndex        =   3
         Tag             =   "nombre"
         Top             =   840
         Width           =   4575
      End
      Begin VB.Label Label1 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Codigo Cuenta"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label Label3 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Nombre"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   840
         Width           =   1095
      End
   End
   Begin VB.TextBox pivote 
      Height          =   285
      Left            =   2160
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   6000
      Visible         =   0   'False
      Width           =   615
   End
   Begin MSAdodcLib.Adodc mcm 
      Height          =   375
      Left            =   2400
      Top             =   6840
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   1
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   3
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin XPFrame.FrameXp FrameQuickMenu 
      Height          =   615
      Left            =   9480
      TabIndex        =   12
      Top             =   7680
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
         Left            =   1800
         TabIndex        =   13
         Top             =   280
         Width           =   1335
      End
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash opciones 
      Height          =   1215
      Left            =   120
      TabIndex        =   1
      Top             =   7080
      Width           =   6735
      _cx             =   11880
      _cy             =   2143
      FlashVars       =   ""
      Movie           =   "c:\barra_opciones.swf"
      Src             =   "c:\barra_opciones.swf"
      WMode           =   "Transparent"
      Play            =   "0"
      Loop            =   "-1"
      Quality         =   "High"
      SAlign          =   ""
      Menu            =   "-1"
      Base            =   ""
      AllowScriptAccess=   "always"
      Scale           =   "ShowAll"
      DeviceFont      =   "0"
      EmbedMovie      =   "0"
      BGColor         =   ""
      SWRemote        =   ""
      MovieData       =   ""
      SeamlessTabbing =   "1"
      Profile         =   "0"
      ProfileAddress  =   ""
      ProfilePort     =   "0"
      AllowNetworking =   "all"
      AllowFullScreen =   "false"
   End
End
Attribute VB_Name = "maestro03"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public saldoglobal As Double
Public saldocuenta As Double

Private Sub codigo_Click()
    Call dato1_KeyDown(vbKeyF2, 0)
End Sub

Private Sub Command1_Click()
imprimir
End Sub


Private Sub dato1_GotFocus()
grillasaldos
Call cargatexto(dato1)
End Sub
Private Sub dato2_GotFocus()
Call cargatexto(dato2)
End Sub
Private Sub dato3_GotFocus()
    If MODIFI = 0 Then leer
    Call cargatexto(dato3)
End Sub

Private Sub dato1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 38 Then Unload Me: GoTo no:
    If KeyCode = vbKeyF2 Then Call ayudacentrocosto(dato3)
    Call flechas(dato1, dato2, KeyCode)
no:
End Sub

Private Sub dato2_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Call flechas(dato1, dato3, KeyCode)
End Sub
Private Sub dato3_KeyDown(KeyCode As Integer, Shift As Integer)
    Call flechas(dato2, dato3, KeyCode)
End Sub

Private Sub Form_Load()

    
    
    Call Conectar_BD
    Rem Call Funciones_Forms_M_Productos.Conecta_Maestro_Productos
    sc = 0
    opciones.Visible = False
DOCU(1) = "ACTIVO"
DOCU(2) = "PASIVO"
DOCU(3) = "RESULTADO"
CANDO = 3

Rem Call RECUPERAFECHA
Call CARGAPERMISO(Me.Name)
Call CARGAGRILLA(1, 4)


End Sub
Sub CARGAGRILLA(row, col)
    Dim FORMATOGRILLA(10, 10) As String
    Rem DATOS DE LA COLUMNA
    FORMATOGRILLA(1, 1) = "CODIGO"
    FORMATOGRILLA(1, 2) = "NOMBRE"
    FORMATOGRILLA(1, 3) = "SALDO"

    Rem LARGO DE LOS DATOS
    
    FORMATOGRILLA(2, 1) = "6"
    FORMATOGRILLA(2, 2) = "20"
    FORMATOGRILLA(2, 3) = "8"
    Rem TIPO DE DATOS C=CEROS,S=STRING,N=NUMERICO,D=DATE,
    FORMATOGRILLA(3, 1) = "C"
    FORMATOGRILLA(3, 2) = "C"
    FORMATOGRILLA(3, 3) = "N"
    Rem FORMATO GRILLA
    FORMATOGRILLA(4, 1) = ""
    FORMATOGRILLA(4, 2) = ""
    FORMATOGRILLA(4, 3) = " ###,###,##0"
    Rem LOCCKED
    FORMATOGRILLA(5, 1) = "true"
    FORMATOGRILLA(5, 2) = "true"
    FORMATOGRILLA(5, 3) = "true"
    Rem VALOR MINIMO
    FORMATOGRILLA(6, 1) = ""
    FORMATOGRILLA(6, 2) = ""
    FORMATOGRILLA(6, 3) = ""
    Rem VALOR MAXIMO
    FORMATOGRILLA(7, 1) = ""
    FORMATOGRILLA(7, 2) = ""
    FORMATOGRILLA(7, 3) = ""
    Grid1.Cols = col
    Grid1.Rows = row
    Grid1.AllowUserResizing = False
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
    For k = 1 To col - 1
        Grid1.Cell(0, k).text = FORMATOGRILLA(1, k)
        Grid1.Column(k).Width = Val(FORMATOGRILLA(2, k)) * 10.5
        Grid1.Column(k).MaxLength = Val(FORMATOGRILLA(2, k))
        Grid1.Column(k).FormatString = FORMATOGRILLA(4, k)
        Grid1.Column(k).Locked = FORMATOGRILLA(5, k)
        If FORMATOGRILLA(3, k) = "N" Then Grid1.Column(k).Alignment = cellRightCenter
        If FORMATOGRILLA(3, k) = "D" Then Grid1.Column(k).CellType = cellCalendar
        
        
    Next k
End Sub

Private Sub dato1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then Unload Me
    snum = 0: KeyAscii = esNumero(KeyAscii)
    If KeyAscii = 13 Then Call ceros(dato1): Call Pregunta(dato1, dato2)
End Sub

Private Sub dato2_KeyPress(KeyAscii As Integer)
    snum = 0: KeyAscii = esNumero(KeyAscii)
    If KeyAscii = 13 Then Call ceros(dato2): Call Pregunta(dato2, dato3)
End Sub

Private Sub dato3_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then grabar: retorno
End Sub


Sub leer()
    campos(0, 0) = dato1.Tag
    campos(1, 0) = dato3.Tag
    campos(2, 0) = ""
    campos(0, 2) = "centrosdecosto"
    condicion = "codigo=" + "'" + dato1.text + dato2.text + "' and año='" + Format(fechasistema, "yyyy") + "'"

    op = 5
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)
    If sqlconta.status = 4 Then GoTo no:
    carga
    opciones.Visible = True
    disponible (True)
    habilita (True)
    opciones.SetFocus
    
no:
End Sub
Sub leersiguiente()
    campos(0, 0) = dato1.Tag
    campos(1, 0) = dato3.Tag
    campos(2, 0) = ""
    campos(0, 2) = "centrosdecosto"
    condicion = "codigo>" + "'" + dato1.text + dato2.text + "' and año='" + Format(fechasistema, "yyyy") + "' order by codigo"
    op = 5
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)
    If sqlconta.status = 4 Then GoTo no:
    carga
    opciones.Visible = True
    disponible (True)
    habilita (True)
    opciones.SetFocus
    
no:
   
    
End Sub
Sub leeranterior()
    campos(0, 0) = dato1.Tag
    campos(1, 0) = dato3.Tag
    campos(2, 0) = ""
    campos(0, 2) = "centrosdecosto"
    condicion = "codigo<" + "'" + dato1.text + dato2.text + "' and año='" + Format(fechasistema, "yyyy") + "' order by codigo desc"

    op = 5
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)
   If sqlconta.status = 4 Then GoTo no:
    carga
    opciones.Visible = True
    disponible (True)
    habilita (True)
    opciones.SetFocus

    
no:
   
    
End Sub

Sub carga()
    habilita (True)
    dato1.text = Mid(sqlconta.response(0, 3), 1, 2)
    dato2.text = Mid(sqlconta.response(0, 3), 3, 2)
    dato3.text = sqlconta.response(1, 3)
    leercrcc
fin:
End Sub

Sub habilita(ByVal condicion As Boolean)
    
    dato1.Locked = condicion
    dato2.Locked = condicion
    dato3.Locked = condicion
    
End Sub
Sub disponible(ByVal condicion As Boolean)
    
    dato1.Enabled = condicion
    dato2.Enabled = condicion
    dato3.Enabled = condicion
    
End Sub


Sub Pregunta(ByRef caja As TextBox, ByRef caja1 As TextBox)
    If caja.text = "" And sc = 0 Then caja.SetFocus
    If caja.text <> "" Or sc = 1 Then caja1.Enabled = True: caja1.SetFocus
    sc = 0
End Sub


Sub ayudacentrocosto(ByRef caja As TextBox)
    Dim campos As Variant
    Dim cfijo As Variant
    Dim largo As Variant
    campos = Array("codigo", "nombre")
    largo = Array("8s", "40s")
    cfijo = "año='" + Format(fechasistema, "yyyy") + "'"
    cabezas = Array("codigo", "nombre")
    mensajeAyuda = "Ayuda Centros de Costo"

    Call cargaAyudaT(Servidor, basebus, Usuario, password, "centrosdecosto", pivote, campos, cfijo, largo, 2)
    If Val(pivote.text) = 0 Then dato1.SetFocus: GoTo no
    dato2.Enabled = True
    dato1.text = Mid(pivote.text, 1, 2)
    dato2.text = Mid(pivote.text, 3, 2)
    
    caja.Enabled = True
    caja.SetFocus
no:
End Sub


Sub flechas(ByRef caja As TextBox, ByRef caja1 As TextBox, ByRef codigo As Integer)
    If codigo = 38 And caja.Enabled = True Then caja.SetFocus
    If codigo = 40 And caja1.Enabled = True Then caja1.SetFocus
End Sub
Sub grabar()
    
    
    campos(0, 0) = dato1.Tag
    campos(1, 0) = dato3.Tag
    campos(2, 0) = "año"
    
    campos(3, 0) = ""
    
    campos(0, 1) = dato1.text + dato2.text
    campos(1, 1) = dato3.text
    campos(2, 1) = Format(fechasistema, "yyyy")
    
    campos(0, 2) = "centrosdecosto"
    If MODIFI = 1 Then condicion = "codigo=" + "'" + dato1.text + dato2.text + "' and año='" + Format(fechasistema, "yyyy") + "'"
    If MODIFI = 1 Then op = 3 Else op = 2
    
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)
    If MODIFI = 0 Then grabarcuentas
    MODIFI = 0
End Sub
Sub grabar2(cuenta, año)
    
    campos(0, 0) = "codigo"
    campos(1, 0) = "año"
    campos(2, 0) = "cuenta"
    campos(3, 0) = ""
    campos(0, 1) = dato1.text + dato2.text
    campos(1, 1) = año
    campos(2, 1) = cuenta
    campos(0, 2) = "saldoscentrosdecosto"
    op = 2
    
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)
    

End Sub

Sub ELIMINAR()
    campos(0, 2) = "centrosdecosto"
    condicion = "codigo=" + "'" + dato1.text + dato2.text + "' and año='" + Format(fechasistema, "yyyy") + "'"
    op = 4
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)
    campos(0, 2) = "saldoscentrosdecosto"
    condicion = "codigo=" + "'" + dato1.text + dato2.text + "' and año='" + Format(fechasistema, "yyyy") + "'"
    op = 4
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)

    
End Sub


Private Sub Label18_Click()

End Sub

Private Sub lblhistorico_Click(Index As Integer)

End Sub




Private Sub Frame2_DragDrop(Source As CONTROL, X As Single, Y As Single)

End Sub

Private Sub Grid1_LeaveCell(ByVal row As Long, ByVal col As Long, NewRow As Long, NewCol As Long, Cancel As Boolean)
DATOSSALDOS (Grid1.Cell(NewRow, 1).text)

End Sub

Private Sub opciones_FSCommand(ByVal command As String, ByVal args As String)

If command = "retorno" Then retorno

If command = "modifica" Then disponible (True): habilita (False): dato1.Enabled = False: dato2.Enabled = False: dato3.SetFocus: MODIFI = 1

If command = "elimina" Then ELIMINAR: retorno
If command = "siguiente" Then leersiguiente
If command = "anterior" Then leeranterior
If command = "imprime" Then imprimir
If command = "movimientos" Then informa04.Show


End Sub
Sub retorno()
disponible (True)
habilita (False)
limpia
opciones.Visible = False
dato1.Enabled = True
dato1.SetFocus
Grid1.Rows = 1
End Sub
Sub limpia()
    dato1.text = ""
    dato2.text = ""
    dato3.text = ""
End Sub

Sub imprimir()
    
End Sub
Sub cabeza()
    
End Sub


Sub Consulta_Informe()
    Dim resultados As rdoResultset
    Dim csql As New rdoQuery
    Dim rut As String
    
    
        Set csql.ActiveConnection = contadb
        csql.sql = "SELECT codigo,nombre,tipo,ctacte,glosa,centrocosto "
        csql.sql = csql.sql + "FROM cuentasdelmayor where año='" + Format(fechasistema, "yyyy") + "' "
        csql.sql = csql.sql + " order by codigo"
        csql.Execute
        If csql.RowsAffected > 0 Then
            Set resultados = csql.OpenResultset
            While Not resultados.EOF
                
                dato(1) = Mid(resultados(0), 1, 2) + "." + Mid(resultados(0), 3, 2) + "." + Mid(resultados(0), 5, 4): colu(1) = 15: tipodato(1) = "s"
                dato(2) = resultados(1): colu(2) = 52: tipodato(2) = "s"
                dato(3) = resultados(2) + " " + DOCU$(Val(resultados(2)))
                dato(4) = resultados(3)
                dato(5) = resultados(4)
                dato(6) = resultados(5) + " " + DOCU2$(Val(resultados(5)))
                colu(3) = 10: tipodato(3) = "s"
                colu(4) = 10: tipodato(4) = "s"
                colu(5) = 20: tipodato(5) = "s"
                colu(6) = 10: tipodato(6) = "s"
                 cancolu = 6
                
                resultados.MoveNext
            Wend
            resultados.Close
            
            Set resultados = Nothing

        End If
    

End Sub

Sub DATOSSALDOS(cuenta)

Call LEERSALDOS(cuenta)
sumador = Val(sqlconta.response(3, 3)) - Val(sqlconta.response(4, 3))
SALDOS.TextMatrix(1, 1) = Format(sqlconta.response(3, 3), "###,###,##0")
SALDOS.TextMatrix(1, 2) = Format(sqlconta.response(4, 3), "###,###,##0")
SALDOS.TextMatrix(1, 3) = Format(sumador, "###,###,##0")
For k = 5 To 16
SALDOS.TextMatrix(k - 3, 1) = Format(sqlconta.response(k, 3), "###,###,##0")
SALDOS.TextMatrix(k - 3, 2) = Format(sqlconta.response(k + 12, 3), "###,###,##0")
sumador = sumador + Val(sqlconta.response(k, 3)) - Val(sqlconta.response(k + 12, 3))
SALDOS.TextMatrix(k - 3, 3) = Format(sumador, "###,###,##0")
Next k
saldocuenta = sumador

End Sub
Sub grillasaldos()
SALDOS.Cols = 4
SALDOS.Rows = 14
SALDOS.ColWidth(0) = 120 * 12
SALDOS.ColWidth(1) = 120 * 8
SALDOS.ColWidth(2) = 120 * 8
SALDOS.ColWidth(3) = 120 * 8
SALDOS.TextMatrix(0, 0) = "MESES   "
SALDOS.TextMatrix(0, 1) = "DEBE    "
SALDOS.TextMatrix(0, 2) = "HABER   "
SALDOS.TextMatrix(0, 3) = "SALDO   "
SALDOS.TextMatrix(1, 0) = "AÑO ANTERIOR"
SALDOS.TextMatrix(2, 0) = "ENERO"
SALDOS.TextMatrix(3, 0) = "FEBRERO"
SALDOS.TextMatrix(4, 0) = "MARZO"
SALDOS.TextMatrix(5, 0) = "ABRIL"
SALDOS.TextMatrix(6, 0) = "MAYO"
SALDOS.TextMatrix(7, 0) = "JUNIO"
SALDOS.TextMatrix(8, 0) = "JULIO"
SALDOS.TextMatrix(9, 0) = "AGOSTO"
SALDOS.TextMatrix(10, 0) = "SEPTIEMBRE"
SALDOS.TextMatrix(11, 0) = "OCTUBRE"
SALDOS.TextMatrix(12, 0) = "NOVIEMBRE "
SALDOS.TextMatrix(13, 0) = "DICIEMBRE "
For k = 1 To 13
SALDOS.TextMatrix(k, 1) = "0"
SALDOS.TextMatrix(k, 2) = "0"
SALDOS.TextMatrix(k, 3) = "0"
Next k
End Sub

Sub LEERSALDOS(cuenta)
    
    campos(0, 0) = "codigo"
    campos(1, 0) = "año"
    campos(2, 0) = "cuenta"
    
    campos(3, 0) = "debeanterior"
    campos(4, 0) = "haberanterior"
    campos(5, 0) = "debe01"
    campos(6, 0) = "debe02"
    campos(7, 0) = "debe03"
    campos(8, 0) = "debe04"
    campos(9, 0) = "debe05"
    campos(10, 0) = "debe06"
    campos(11, 0) = "debe07"
    campos(12, 0) = "debe08"
    campos(13, 0) = "debe09"
    campos(14, 0) = "debe10"
    campos(15, 0) = "debe11"
    campos(16, 0) = "debe12"
    campos(17, 0) = "haber01"
    campos(18, 0) = "haber02"
    campos(19, 0) = "haber03"
    campos(20, 0) = "haber04"
    campos(21, 0) = "haber05"
    campos(22, 0) = "haber06"
    campos(23, 0) = "haber07"
    campos(24, 0) = "haber08"
    campos(25, 0) = "haber09"
    campos(26, 0) = "HABER10"
    campos(27, 0) = "HABER11"
    campos(28, 0) = "HABER12"
    campos(29, 0) = ""
    condicion = "codigo=" + "'" + dato1.text + dato2.text + "' and año='" + Mid(fechasistema, 7, 4) + "' and cuenta='" + cuenta + "'"
    campos(0, 2) = "saldoscentrosdecosto"
    op = 5
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)
    
    End Sub

Sub leercrcc()
    Dim resultados2 As rdoResultset
    Dim csql2 As New rdoQuery
    Dim rut As String
    Dim PASO As String
    Dim LINEAS As Double
     Set csql2.ActiveConnection = contadb
     csql2.sql = "SELECT scr.cuenta,cm.nombre "
     csql2.sql = csql2.sql + "FROM saldoscentrosdecosto as scr,cuentasdelmayor as cm "
     csql2.sql = csql2.sql + "WHERE scr.codigo='" + dato1.text + dato2.text + "' and scr.cuenta=cm.codigo AND scr.año=cm.año "
     csql2.sql = csql2.sql + "and scr.año='" + Format(fechasistema, "yyyy") + "' "
     csql2.sql = csql2.sql + "order by cm.codigo"
     csql2.Execute
     LINEAS = 0
     'Grid1.Rows = cSql2.RowsAffected + 1
     Grid1.Rows = 1
     Grid1.AutoRedraw = False
     If csql2.RowsAffected > 0 Then
         Set resultados2 = csql2.OpenResultset
         While Not resultados2.EOF
             LINEAS = LINEAS + 1
             'Grid1.Cell(LINEAS, 1).text = resultados2(0)
             'Grid1.Cell(LINEAS, 2).text = resultados2(1)
             Call DATOSSALDOS(resultados2(0))
             Grid1.AddItem resultados2(0) & vbTab & resultados2(1) & vbTab & saldocuenta, True
             'Grid1.Cell(LINEAS, 3).text = saldocuenta
             resultados2.MoveNext
         Wend
         resultados2.Close
         Set resultados2 = Nothing
     End If
     Grid1.AutoRedraw = True
     Grid1.Refresh
End Sub

Sub grabarcuentas()

Dim resultados2 As rdoResultset
    Dim csql2 As New rdoQuery
    Dim rut As String
    Dim PASO As String
    Dim LINEAS As Double
        Set csql2.ActiveConnection = contadb
        csql2.sql = "SELECT codigo,nombre "
        csql2.sql = csql2.sql + "FROM cuentasdelmayor where crcc='1' and año='" + Format(fechasistema, "yyyy") + "' "
        
        csql2.sql = csql2.sql + "order by codigo"
        csql2.Execute
        LINEAS = 0
        
        If csql2.RowsAffected > 0 Then
         
        
        Set resultados2 = csql2.OpenResultset
        While Not resultados2.EOF
        Call grabar2(resultados2(0), Mid(fechasistema, 7, 4))
        
        resultados2.MoveNext
        Wend
          
          resultados2.Close
            Set resultados2 = Nothing

        End If
    
End Sub


Private Sub botonmisaccesos_Click()
programafiltro = Me.Caption
misaccesos.Show
End Sub


Private Sub botonmisfavoritos_Click()
Call AgregaFavoritos(USUARIOSISTEMA, App.EXEName, Me.Name, Me.Caption)

End Sub
