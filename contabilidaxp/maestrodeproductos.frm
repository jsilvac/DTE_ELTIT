VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form maestro01 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mestro de Cuentas del Mayor"
   ClientHeight    =   8895
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11790
   DrawStyle       =   5  'Transparent
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   593
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   786
   Begin MSAdodcLib.Adodc mcm 
      Height          =   375
      Left            =   8520
      Top             =   8520
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
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFF2F7&
      BorderStyle     =   0  'None
      Caption         =   "Datos personales"
      Height          =   3135
      Left            =   6720
      TabIndex        =   6
      Top             =   3720
      Width           =   5055
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Historico de Cuentas "
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
         Left            =   120
         TabIndex        =   7
         Top             =   120
         Width           =   2055
      End
      Begin VB.Shape Shape3 
         BorderColor     =   &H00000080&
         BorderWidth     =   3
         FillColor       =   &H00FFC0C0&
         Height          =   3135
         Left            =   0
         Top             =   0
         Width           =   5055
      End
   End
   Begin VB.Frame datospersonales 
      BackColor       =   &H00FFF2F7&
      BorderStyle     =   0  'None
      Caption         =   "Datos personales"
      Height          =   5055
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   6135
      Begin VB.TextBox dato6 
         BackColor       =   &H00FBEDE6&
         Enabled         =   0   'False
         ForeColor       =   &H00000080&
         Height          =   285
         Left            =   1200
         MaxLength       =   1
         TabIndex        =   16
         Tag             =   "centrocosto"
         Top             =   3120
         Width           =   375
      End
      Begin VB.TextBox dato5 
         BackColor       =   &H00FBEDE6&
         Enabled         =   0   'False
         ForeColor       =   &H00000080&
         Height          =   285
         Left            =   1200
         MaxLength       =   100
         TabIndex        =   14
         Tag             =   "glosa"
         Top             =   2640
         Width           =   4335
      End
      Begin VB.TextBox dato4 
         BackColor       =   &H00FBEDE6&
         Enabled         =   0   'False
         ForeColor       =   &H00000080&
         Height          =   285
         Left            =   1200
         MaxLength       =   10
         TabIndex        =   12
         Tag             =   "ctacte"
         Top             =   2160
         Width           =   1215
      End
      Begin VB.TextBox dato3 
         BackColor       =   &H00FBEDE6&
         Enabled         =   0   'False
         ForeColor       =   &H00000080&
         Height          =   285
         Left            =   1200
         MaxLength       =   3
         TabIndex        =   11
         Tag             =   "tipo"
         Top             =   1680
         Width           =   375
      End
      Begin VB.TextBox dato2 
         BackColor       =   &H00FBEDE6&
         Enabled         =   0   'False
         ForeColor       =   &H00000080&
         Height          =   285
         Left            =   1200
         MaxLength       =   40
         TabIndex        =   10
         Tag             =   "nombre"
         Top             =   1200
         Width           =   4335
      End
      Begin VB.TextBox dato1 
         BackColor       =   &H00FBEDE6&
         ForeColor       =   &H00000080&
         Height          =   285
         Left            =   3360
         MaxLength       =   8
         TabIndex        =   9
         Tag             =   "codigo"
         Top             =   600
         Width           =   855
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "C. Costo"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   3240
         Width           =   975
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Glosa"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   2640
         Width           =   855
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Cta. Corriente"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   2160
         Width           =   975
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   1680
         Width           =   495
      End
      Begin VB.Label codigo 
         BackStyle       =   0  'Transparent
         Caption         =   "F2 ( ? )"
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
         Left            =   480
         MouseIcon       =   "maestrodeproductos.frx":0000
         MousePointer    =   99  'Custom
         TabIndex        =   8
         Top             =   480
         Width           =   735
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Codigo Cuenta"
         Height          =   255
         Left            =   1680
         TabIndex        =   5
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Nombre"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   1200
         Width           =   855
      End
      Begin VB.Shape Shape4 
         BorderColor     =   &H00000080&
         BorderWidth     =   3
         FillColor       =   &H00FFC0C0&
         Height          =   5055
         Left            =   0
         Top             =   0
         Width           =   6135
      End
      Begin VB.Label Label16 
         BackStyle       =   0  'Transparent
         Caption         =   "Cuentas del Mayor"
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
         Left            =   240
         TabIndex        =   2
         Top             =   120
         Width           =   2055
      End
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash opciones 
      Height          =   1215
      Left            =   480
      TabIndex        =   1
      Top             =   7440
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
   Begin VB.Shape Shape6 
      BackColor       =   &H00808080&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00404040&
      FillColor       =   &H00FF8080&
      Height          =   3135
      Left            =   6720
      Top             =   3720
      Width           =   4935
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00808080&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00404040&
      FillColor       =   &H00FF8080&
      Height          =   5055
      Left            =   240
      Top             =   360
      Width           =   6135
   End
End
Attribute VB_Name = "maestro01"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub codigo_Click()
    Call dato1_KeyDown(vbKeyF2, 0)
End Sub

Private Sub dato1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF2 Then Call ayudaproducto(dato2)
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

Private Sub dato4_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF2 Then Call ayudaseccion(dato5)
    Call flechas(dato3, dato5, KeyCode)
End Sub

Private Sub dato5_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF2 Then Call ayudadepto(dato6)
    Call flechas(dato4, dato6, KeyCode)
End Sub

Private Sub dato6_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF2 Then Call ayudalinea(DATO7)
    Call flechas(dato5, DATO7, KeyCode)
End Sub

Private Sub Form_Load()
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
    If KeyAscii = 13 Then Call ceros(dato4): Call Pregunta(dato4, dato5)
End Sub

Private Sub dato4_LostFocus()
        
    campos(0, 0) = "codigo"
    campos(1, 0) = "nombre"
    campos(2, 0) = "tipo"
    campos(3, 0) = "ctacte"
    campos(4, 0) = "glosa"
    campos(5, 0) = "centrocosto"
    campos(6, 0) = ""
    campos(0, 2) = "cuentasdelmayor"
    
    condicion = "codigo= '" & dato4.text & "'"
    op = 5
    Set SQLUTIL.conexion = db
    SQLUTIL.datos = campos
    Call SQLUTIL.SQLUTIL(op, condicion)
    status = SQLUTIL.estado
    label(0) = SQLUTIL.datos(1, 3)
    If status <> 0 Then dato4.SetFocus
End Sub

Private Sub dato5_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Call ceros(dato5): Call Pregunta(dato5, dato6)
End Sub
Private Sub dato5_LostFocus()
    campos(0, 0) = "codigodepto"
    campos(1, 0) = "nombre"
    campos(2, 0) = ""
    campos(0, 2) = "maestrodepartamentos"
    condicion = "codigodepto = '" & dato5.text & "'"
    op = 5
    Set SQLUTIL.conexion = db
    SQLUTIL.datos = campos
    Call SQLUTIL.SQLUTIL(op, condicion)
    status = SQLUTIL.estado
    label(1) = SQLUTIL.datos(1, 3)
    If status <> 0 Then dato5.SetFocus
End Sub

Private Sub dato6_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Call Pregunta(dato6, DATO7)
End Sub
Private Sub dato6_LostFocus()
    campos(0, 0) = "codigolinea"
    campos(1, 0) = "nombre"
    campos(2, 0) = ""
    campos(0, 2) = "maestrolineas"
    condicion = "codigodepto = '" & dato5.text & "' AND codigolinea = '" & dato6.text & "'"
    op = 5
    Set SQLUTIL.conexion = db
    SQLUTIL.datos = campos
    Call SQLUTIL.SQLUTIL(op, condicion)
    status = SQLUTIL.estado
    label(2) = SQLUTIL.datos(1, 3)
    If status <> 0 Then dato6.SetFocus
End Sub
Private Sub dato7_LostFocus()
    campos(0, 0) = "codigoimpuesto"
    campos(1, 0) = "nombre"
    campos(2, 0) = ""
    campos(0, 2) = "maestroimpuestos"
    condicion = "codigoimpuesto = '" & DATO7.text & "'"
    op = 5
    Set SQLUTIL.conexion = db
    SQLUTIL.datos = campos
    Call SQLUTIL.SQLUTIL(op, condicion)
    status = SQLUTIL.estado
    label(3) = SQLUTIL.datos(1, 3)
    If status <> 0 Then DATO7.SetFocus
End Sub

Private Sub dato7_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Call ceros(DATO7): Call Pregunta(DATO7, DATO8)

End Sub

Private Sub dato8_KeyPress(KeyAscii As Integer)
    KeyAscii = esNumero(KeyAscii)
    If KeyAscii = 13 Then Call formato(DATO8): Call Pregunta(DATO8, DATO9)
End Sub

Private Sub dato9_KeyPress(KeyAscii As Integer)
        KeyAscii = esNumero(KeyAscii)
    If KeyAscii = 13 Then Call formato(DATO9): Call Pregunta(DATO9, DATO10)
End Sub

Private Sub dato10_KeyPress(KeyAscii As Integer)
    KeyAscii = esNumero(KeyAscii)
    If KeyAscii = 13 Then Call formato(DATO10): Call Pregunta(DATO10, DATO11)
End Sub

Private Sub dato11_KeyPress(KeyAscii As Integer)
    KeyAscii = esNumero(KeyAscii)
    If KeyAscii = 13 Then Call formato(DATO11): Call Pregunta(DATO11, DATO12)
End Sub

Private Sub dato12_KeyPress(KeyAscii As Integer)
    KeyAscii = esNumero(KeyAscii)
    If KeyAscii = 13 Then Call Pregunta(DATO12, DATO13)
End Sub

Private Sub dato13_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then Call Pregunta(DATO13, dato14)
End Sub

Private Sub dato14_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then GRABAR: leer:
End Sub

Private Sub foto_DblClick()
    cargaFoto.Show vbModal
End Sub

Sub leer()
    campos(0, 0) = dato1.Tag 'CODIGO
    campos(1, 0) = dato2.Tag 'DESCRIPCION
    campos(2, 0) = dato3.Tag 'tipo
    campos(3, 0) = dato4.Tag 'cta.cte
    campos(4, 0) = dato5.Tag 'glosa
    campos(5, 0) = dato6.Tag 'centro costo
    campos(0, 2) = "cuentasdelmayor"
    condicion = "codigo=" + "'" + dato1.text + "'"
    op = 5
    SQLUTIL.datos = campos
    Set SQLUTIL.conexion = db
    Call SQLUTIL.SQLUTIL(op, condicion)
    If SQLUTIL.estado = 0 Then carga: opciones.Visible = True: disponible (True): habilita (True): opciones.SetFocus
    
End Sub
Sub leersiguiente()
    campos(0, 0) = dato1.Tag
    campos(1, 0) = dato2.Tag
    campos(2, 0) = dato3.Tag
    campos(3, 0) = dato4.Tag
    campos(4, 0) = dato5.Tag
    campos(5, 0) = dato6.Tag
    campos(0, 2) = "cuentasdelmayor"
    condicion = "codigo>" + "'" + dato1.text + "' order by codigo"
    op = 5
    SQLUTIL.datos = campos
    Set SQLUTIL.conexion = db
    Call SQLUTIL.SQLUTIL(op, condicion)
    If SQLUTIL.estado = 0 Then carga: opciones.Visible = True: disponible (True): habilita (True): opciones.SetFocus
    
End Sub
Sub leeranterior()
    campos(0, 0) = dato1.Tag
    campos(1, 0) = dato2.Tag
    campos(2, 0) = dato3.Tag
    campos(3, 0) = dato4.Tag
    campos(4, 0) = dato5.Tag
    campos(5, 0) = dato6.Tag
    campos(0, 2) = "cuentasdelmayor"
    condicion = "codigo<" + "'" + dato1.text + "' order by codigo"
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
    dato5.text = SQLUTIL.datos(4, 3)
    dato6.text = SQLUTIL.datos(5, 3)
fin:
End Sub

Sub habilita(ByVal condicion As Boolean)
    
    dato1.Locked = condicion
    dato2.Locked = condicion
    dato3.Locked = condicion
    dato4.Locked = condicion
    dato5.Locked = condicion
    dato6.Locked = condicion
    
End Sub
Sub disponible(ByVal condicion As Boolean)
    
    dato1.Enabled = condicion
    dato2.Enabled = condicion
    dato3.Enabled = condicion
    dato4.Enabled = condicion
    dato5.Enabled = condicion
    dato6.Enabled = condicion
    
End Sub
Sub Conecta_Maestro_Productos()
    'GENERA LA CONEXION Y LA CONSULTA DEL DATA CONTROL.
    With maestro01
        .mcm.ConnectionString = "DRIVER={MySQL ODBC 3.51 Driver};server=localhost;uid=root;pwd=;database=conta01"
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
    Call cargaAyudaT("servidor", "conta01", "root", "123", "maestrosecciones", dato4, campos, cfijo, 2)
    caja.Enabled = True
    caja.SetFocus
    
End Sub

Sub ayudaimpuesto(ByRef caja As TextBox)
    Dim campos As Variant
    Dim cfijo As Variant
    campos = Array("codigoimpuesto", "nombre")
    cfijo = Array("no")
    Call cargaAyudaT("servidor", "conta01", "root", "123", "maestroimpuestos", DATO7, campos, cfijo, 2)
    caja.Enabled = True
    caja.SetFocus
End Sub
Sub ayudadepto(ByRef caja As TextBox)
    Dim campos As Variant
    Dim cfijo As Variant
    campos = Array("codigodepto", "nombre")
    cfijo = Array("codigoseccion", dato4.text)
    Call cargaAyudaT("servidor", "conta01", "root", "123", "maestrodepartamentos", dato5, campos, cfijo, 2)
    caja.Enabled = True
    caja.SetFocus
    

End Sub

Sub ayudalinea(ByRef caja As TextBox)
    Dim campos As Variant
    Dim cfijo As Variant
    campos = Array("codigolinea", "nombre")
    cfijo = Array("codigodepto", dato5.text)
    Call cargaAyudaT("servidor", "conta01", "root", "123", "maestrolineas", dato6, campos, cfijo, 2)
    caja.Enabled = True
    caja.SetFocus
    
End Sub

Sub ayudaproducto(ByRef caja As TextBox)
    Dim campos As Variant
    Dim cfijo As Variant
    campos = Array("codigo", "nombre")
    cfijo = Array("no")
    Call cargaAyudaT("servidor", "conta01", "root", "123", "cuentasdelmayor", dato1, campos, cfijo, 2)
    sl = 0: leer
    caja.Enabled = True
    caja.SetFocus
End Sub

Sub flechas(ByRef caja As TextBox, ByRef caja1 As TextBox, ByRef codigo As Integer)
    If codigo = 38 And caja.Enabled = True Then caja.SetFocus
    If codigo = 40 And caja1.Enabled = True Then caja1.SetFocus
End Sub
Sub GRABAR()
    campos(0, 0) = dato1.Tag
    campos(1, 0) = dato2.Tag
    campos(2, 0) = dato3.Tag
    campos(3, 0) = dato4.Tag
    campos(4, 0) = dato5.Tag
    campos(5, 0) = dato6.Tag
    campos(6, 0) = DATO7.Tag
   
    
    campos(0, 2) = "cuentasdelmayor"
    If modifi = 1 Then condicion = "codigo=" + "'" + dato1.text + "'"
    If modifi = 1 Then op = 3 Else op = 2
    
    SQLUTIL.datos = campos
    Set SQLUTIL.conexion = db
    Call SQLUTIL.SQLUTIL(op, condicion)
modifi = 0

End Sub
Sub ELIMINAR()
    
    campos(0, 2) = "cuentasdelmayor"
    condicion = "codigo=" + "'" + dato1.text + "'"
    op = 4
    SQLUTIL.datos = campos
    Set SQLUTIL.conexion = db
    Call SQLUTIL.SQLUTIL(op, condicion)

    
End Sub


Private Sub Label11_Click()

End Sub

Private Sub Label18_Click()

End Sub

Private Sub lblhistorico_Click(Index As Integer)

End Sub

Private Sub opciones_FSCommand(ByVal command As String, ByVal args As String)
If command = "retorno" Then disponible (True): habilita (False): limpia: opciones.Visible = False: dato1.SetFocus
If command = "modifica" Then disponible (True): habilita (False): dato1.Enabled = False: dato2.SetFocus: modifi = 1
If command = "elimina" Then disponible (True): habilita (False): ELIMINAR: limpia: opciones.Visible = False: dato1.SetFocus
If command = "siguiente" Then leersiguiente
If command = "anterior" Then leeranterior
End Sub

Sub limpia()


    dato1.text = ""
    dato2.text = ""
    dato3.text = ""
    dato4.text = ""
    dato5.text = ""
    dato6.text = ""
   
End Sub
