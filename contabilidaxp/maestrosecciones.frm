VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8b.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form maestro02 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mestro de Secciones"
   ClientHeight    =   4785
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7965
   DrawStyle       =   5  'Transparent
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   319
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   531
   Begin MSAdodcLib.Adodc ms 
      Height          =   375
      Left            =   8040
      Top             =   3120
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
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFF2F7&
      BorderStyle     =   0  'None
      Caption         =   "Datos personales"
      Height          =   2175
      Left            =   840
      TabIndex        =   0
      Top             =   480
      Width           =   6015
      Begin VB.TextBox dato3 
         BackColor       =   &H00FBEDE6&
         Enabled         =   0   'False
         ForeColor       =   &H00000080&
         Height          =   285
         Left            =   1320
         MaxLength       =   5
         TabIndex        =   12
         Tag             =   "descuento"
         Top             =   960
         Width           =   615
      End
      Begin VB.TextBox dato4 
         BackColor       =   &H00FBEDE6&
         Enabled         =   0   'False
         ForeColor       =   &H00000080&
         Height          =   285
         Left            =   1320
         MaxLength       =   5
         TabIndex        =   13
         Tag             =   "porcentaje"
         Top             =   1320
         Width           =   615
      End
      Begin VB.TextBox dato2 
         BackColor       =   &H00FBEDE6&
         Enabled         =   0   'False
         ForeColor       =   &H00000080&
         Height          =   285
         Left            =   1320
         MaxLength       =   50
         TabIndex        =   11
         Tag             =   "nombre"
         Top             =   600
         Width           =   4335
      End
      Begin VB.TextBox dato1 
         BackColor       =   &H00FBEDE6&
         ForeColor       =   &H00000080&
         Height          =   285
         Left            =   1320
         MaxLength       =   5
         TabIndex        =   10
         Tag             =   "codigoseccion"
         Top             =   240
         Width           =   615
      End
      Begin VB.Label label 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Index           =   3
         Left            =   2040
         TabIndex        =   9
         Top             =   1200
         Width           =   3255
      End
      Begin VB.Label label 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Index           =   2
         Left            =   2040
         TabIndex        =   8
         Top             =   960
         Width           =   3255
      End
      Begin VB.Label label 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Index           =   1
         Left            =   2040
         TabIndex        =   7
         Top             =   600
         Width           =   3255
      End
      Begin VB.Label label 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Index           =   0
         Left            =   2040
         TabIndex        =   6
         Top             =   240
         Width           =   3255
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "margen"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   1320
         Width           =   855
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H00000080&
         BorderWidth     =   3
         FillColor       =   &H00FFC0C0&
         Height          =   2175
         Left            =   0
         Top             =   0
         Width           =   6015
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "descuento"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   960
         Width           =   975
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Nombre"
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Seccion"
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   240
         Width           =   855
      End
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash opciones 
      Height          =   1215
      Left            =   720
      TabIndex        =   5
      Top             =   3120
      Visible         =   0   'False
      Width           =   6735
      _cx             =   11880
      _cy             =   2143
      FlashVars       =   ""
      Movie           =   "\\servidor\e\gestion comercial\barra_opciones.swf"
      Src             =   "\\servidor\e\gestion comercial\barra_opciones.swf"
      WMode           =   "Transparent"
      Play            =   0   'False
      Loop            =   -1  'True
      Quality         =   "High"
      SAlign          =   ""
      Menu            =   -1  'True
      Base            =   ""
      AllowScriptAccess=   "always"
      Scale           =   "ShowAll"
      DeviceFont      =   0   'False
      EmbedMovie      =   0   'False
      BGColor         =   ""
      SWRemote        =   ""
      MovieData       =   ""
      SeamlessTabbing =   -1  'True
      Profile         =   0   'False
      ProfileAddress  =   ""
      ProfilePort     =   0
   End
   Begin VB.Shape Shape5 
      BackColor       =   &H00808080&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00404040&
      FillColor       =   &H00FF8080&
      Height          =   2175
      Left            =   960
      Top             =   600
      Width           =   6015
   End
End
Attribute VB_Name = "maestro02"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub dato1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF2 Then Call ayudaseccion(dato2)
    Call flechas(dato1, dato2, KeyCode)
End Sub

Private Sub dato1_LostFocus()
    If sl = 0 Then leer
    sl = 0
End Sub

Private Sub dato2_KeyDown(KeyCode As Integer, Shift As Integer)
    Call flechas(dato1, dato3, KeyCode)
End Sub

Private Sub dato3_KeyDown(KeyCode As Integer, Shift As Integer)
    Call flechas(dato2, dato4, KeyCode)
End Sub

Private Sub Form_Load()
LARGOPAGINA = 60
    Call Conectar_BD
    Call Funciones_Forms_M_Productos.Conecta_Maestro_Productos
    sc = 0
    opciones.Visible = False
End Sub

Private Sub dato1_KeyPress(KeyAscii As Integer)
    snum = 1: KeyAscii = esNumero(KeyAscii)
    If KeyAscii = 13 Then Call ceros(dato1): Call Pregunta(dato1, dato2)
End Sub

Private Sub dato2_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then sc = 1: Call Pregunta(dato2, dato3)
End Sub

Private Sub dato3_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then Call Pregunta(dato3, dato4)
End Sub


Private Sub dato4_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then GRABAR: leer:
End Sub

Private Sub foto_DblClick()
    cargaFoto.Show vbModal
End Sub

Sub leer()
    campos(0, 0) = dato1.Tag 'CODIGO
    campos(1, 0) = dato2.Tag 'DESCRIPCION
    campos(2, 0) = dato3.Tag 'PORCENTAJE
    campos(3, 0) = dato4.Tag 'DESCUENTO
    campos(4, 0) = ""
    campos(0, 2) = "maestrosecciones"
    condicion = "codigoseccion=" + "'" + dato1.text + "'"
    op = 5
    SQLUTIL.datos = campos
    Set SQLUTIL.conexion = db
    Call SQLUTIL.SQLUTIL(op, condicion)
    If SQLUTIL.estado = 0 Then carga: opciones.Visible = True: disponible (True): habilita (True): opciones.SetFocus
End Sub
Sub leersiguiente()
    campos(0, 0) = dato1.Tag 'CODIGO
    campos(1, 0) = dato2.Tag 'DESCRIPCION
    campos(2, 0) = dato3.Tag 'PORCENTAJE
    campos(3, 0) = dato4.Tag 'DESCUENTO
    campos(0, 2) = "maestrosecciones"
    condicion = "codigoseccion>" + "'" + dato1.text + "' order by codigoseccion"
    op = 5
    SQLUTIL.datos = campos
    Set SQLUTIL.conexion = db
    Call SQLUTIL.SQLUTIL(op, condicion)
    If SQLUTIL.estado = 0 Then carga: opciones.Visible = True: disponible (True): habilita (True): opciones.SetFocus
End Sub
Sub leeranterior()
    campos(0, 0) = dato1.Tag 'CODIGO
    campos(1, 0) = dato2.Tag 'DESCRIPCION
    campos(2, 0) = dato3.Tag 'PORCENTAJE
    campos(3, 0) = dato4.Tag 'DESCUENTO
    campos(0, 2) = "maestrosecciones"
    condicion = "codigoseccion<" + "'" + dato1.text + "' order by codigoseccion"
    op = 5
    SQLUTIL.datos = campos
    Set SQLUTIL.conexion = db
    Call SQLUTIL.SQLUTIL(op, condicion)
    If SQLUTIL.estado = 0 Then carga: opciones.Visible = True: disponible (True): habilita (True): opciones.SetFocus
End Sub

Sub carga()
    habilita (True)
    dato1.text = SQLUTIL.datos(0, 3)
    dato2.text = SQLUTIL.datos(1, 3)
    dato3.text = SQLUTIL.datos(2, 3)
    dato4.text = SQLUTIL.datos(3, 3)
fin:
End Sub

Sub habilita(ByVal condicion As Boolean)
    dato1.Locked = condicion
    dato2.Locked = condicion
    dato3.Locked = condicion
    dato4.Locked = condicion
End Sub
Sub disponible(ByVal condicion As Boolean)
    dato1.Enabled = condicion
    dato2.Enabled = condicion
    dato3.Enabled = condicion
    dato4.Enabled = condicion
End Sub
Sub Conecta_Maestro_Secciones()
    'GENERA LA CONEXION Y LA CONSULTA DEL DATA CONTROL.
    With maestro02
        .ms.ConnectionString = "DRIVER={MySQL ODBC 3.51 Driver};server=servidor;uid=root;pwd=123;database=gestioncomercial"
    End With
End Sub

Sub Pregunta(ByRef caja As TextBox, ByRef caja1 As TextBox)
    If caja.text = "" And sc = 0 Then caja.SetFocus
    If caja.text <> "" Or sc = 1 Then caja1.Enabled = True: caja1.SetFocus
    sc = 0
End Sub

Sub ayudaseccion(ByRef caja As TextBox)
    Dim campos As Variant
    Dim cfijo As Variant
    campos = Array("codigoseccion", "nombre")
    cfijo = Array("no")
    Call cargaAyudaT("servidor", "gestioncomercial", "root", "123", "maestrosecciones", dato1, campos, cfijo, 2)
    caja.Enabled = True
    caja.SetFocus
End Sub

Sub flechas(ByRef caja As TextBox, ByRef caja1 As TextBox, ByRef codigo As Integer)
    If codigo = 38 And caja.Enabled = True Then caja.SetFocus
    If codigo = 40 And caja1.Enabled = True Then caja1.SetFocus
End Sub

Sub GRABAR()
    campos(0, 0) = dato1.Tag 'CODIGO
    campos(1, 0) = dato2.Tag 'DESCRIPCION
    campos(2, 0) = dato3.Tag 'PORCENTAJE
    campos(3, 0) = dato4.Tag 'DESCUENTO
    campos(0, 1) = dato1.text 'CODIGO
    campos(1, 1) = dato2.text 'DESCRIPCION
    campos(2, 1) = dato3.text 'UNIDAD MEDIDA
    campos(3, 1) = dato4.text 'SECCION
    campos(0, 2) = "maestrosecciones"
    If modifi = 1 Then condicion = "codigoseccion=" + "'" + dato1.text + "'"
    If modifi = 1 Then op = 3 Else op = 2
    modifi = 0
    SQLUTIL.datos = campos
    Set SQLUTIL.conexion = db
    Call SQLUTIL.SQLUTIL(op, condicion)
    status = SQLUTIL.estado
End Sub
Sub ELIMINAR()
    campos(0, 2) = "maestrosecciones"
    condicion = "codigoseccion=" + "'" + dato1.text + "'"
    op = 4
    SQLUTIL.datos = campos
    Set SQLUTIL.conexion = db
    Call SQLUTIL.SQLUTIL(op, condicion)
End Sub

Private Sub opciones_FSCommand(ByVal command As String, ByVal args As String)
    If command = "retorno" Then disponible (True): habilita (False): limpia: opciones.Visible = False: dato1.SetFocus
    If command = "modifica" Then disponible (True): habilita (False): dato1.Enabled = False: dato2.SetFocus: modifi = 1
    If command = "elimina" Then disponible (True): habilita (False): ELIMINAR: limpia: opciones.Visible = False: dato1.SetFocus
    If command = "siguiente" Then leersiguiente
    If command = "anterior" Then leeranterior
    If command = "imprime" Then imprimir
End Sub

Sub limpia()
    dato1.text = ""
    dato2.text = ""
    dato3.text = ""
    dato4.text = ""
End Sub
Sub imprimir()
    LINEAS = 65
    Consulta_Informe
    informes.Show
End Sub
Sub grilla()
    palabra = ""
    For K = 1 To cancolu
    If tipodato(K) = "s" Or tipodato(K) = "S" Then dato(K) = dato(K) & String(colu(K) - Len(dato(K)), 32)
    If tipodato(K) = "n" Or tipodato(K) = "N" Then dato(K) = String(colu(K) - Len(dato(K)), 32) & dato(K)
    palabra = palabra & dato(K)
    Next K
    If LINEAS > LARGOPAGINA Then Call cabeza
    informes.info.AddItem (palabra)
    LINEAS = LINEAS + 1
End Sub
Sub cabeza()
    informes.info.AddItem ("")
    informes.info.AddItem ("")
    pagina = pagina + 1
    informes.info.AddItem ("NOMBRE EMPRESA                                                                                 PAGINA " + Str$(pagina))
    informes.info.AddItem ("DIRECCION EMPRESA                                                                              ")
    informes.info.AddItem ("                                " + tituloinforme)
    informes.info.AddItem String(180, "=")
    palabra = ""
    For K = 1 To cancolu
    titu(K) = titu(K) & String(colu(K) - Len(titu(K)), 32)
    palabra = palabra & dato(K)
    Next K
    informes.info.AddItem (palabra)
    informes.info.AddItem String(180, "=")
LINEAS = 8
End Sub


Sub Consulta_Informe()
    Dim resultados As rdoResultset
    Dim cSql As New rdoQuery
    Dim rut As String
    
    With informes
        Set cSql.ActiveConnection = db
        cSql.SQL = "SELECT codigoseccion,nombre "
        cSql.SQL = cSql.SQL + "FROM maestrosecciones"
        'cSql.SQL = cSql.SQL + "WHERE rut='" & rut & "'"
        cSql.Execute
        If cSql.RowsAffected > 0 Then
            Set resultados = cSql.OpenResultset
            While Not resultados.EOF
                dato(1) = resultados(0): colu(1) = 15: tipodato(1) = "s"
                dato(2) = resultados(1): colu(2) = 52: tipodato(2) = "s"
                dato(3) = 12363: colu(3) = 15: tipodato(3) = "s"
                dato(4) = 12345: colu(4) = 52: tipodato(4) = "s"
                
                cancolu = 4
                grilla
                resultados.MoveNext
            Wend
            resultados.Close
            Set resultados = Nothing

        End If
    End With

End Sub
