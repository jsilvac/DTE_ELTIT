VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8b.ocx"
Begin VB.Form maestro04 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Maestro Linea de Departamentos"
   ClientHeight    =   5550
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8370
   DrawStyle       =   5  'Transparent
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   370
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   558
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFF2F7&
      BorderStyle     =   0  'None
      Caption         =   "Datos personales"
      Height          =   2895
      Left            =   720
      TabIndex        =   5
      Top             =   480
      Width           =   6735
      Begin VB.TextBox dato5 
         BackColor       =   &H00FBEDE6&
         ForeColor       =   &H00000080&
         Height          =   285
         Left            =   2160
         TabIndex        =   4
         Tag             =   "utilidad"
         Top             =   2160
         Width           =   975
      End
      Begin VB.TextBox dato4 
         BackColor       =   &H00FBEDE6&
         ForeColor       =   &H00000080&
         Height          =   285
         Left            =   2160
         TabIndex        =   3
         Tag             =   "descuento"
         Top             =   1800
         Width           =   975
      End
      Begin VB.TextBox dato3 
         BackColor       =   &H00FBEDE6&
         ForeColor       =   &H00000080&
         Height          =   285
         Left            =   2160
         TabIndex        =   2
         Tag             =   "nombre"
         Top             =   1440
         Width           =   3615
      End
      Begin VB.TextBox dato2 
         BackColor       =   &H00FBEDE6&
         ForeColor       =   &H00000080&
         Height          =   285
         Left            =   2160
         TabIndex        =   1
         Tag             =   "codigolinea"
         Top             =   1080
         Width           =   975
      End
      Begin VB.TextBox dato1 
         BackColor       =   &H00FBEDE6&
         ForeColor       =   &H00000080&
         Height          =   285
         Left            =   2160
         TabIndex        =   0
         Tag             =   "codigodepto"
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Porcentaje Descuento"
         Height          =   255
         Left            =   480
         TabIndex        =   13
         Top             =   1800
         Width           =   1695
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Porcentaje Utilidad"
         Height          =   255
         Left            =   480
         TabIndex        =   12
         Top             =   2160
         Width           =   1455
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Creacion de Lineas de Departamentos"
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
         TabIndex        =   10
         Top             =   120
         Width           =   3375
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Codigo Departamento"
         Height          =   255
         Left            =   480
         TabIndex        =   9
         Top             =   720
         Width           =   1575
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Codigo Linea"
         Height          =   255
         Left            =   480
         TabIndex        =   8
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Nombre"
         Height          =   255
         Left            =   480
         TabIndex        =   7
         Top             =   1440
         Width           =   1095
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H00000080&
         BorderWidth     =   3
         FillColor       =   &H00FFC0C0&
         Height          =   2895
         Left            =   0
         Top             =   0
         Width           =   6735
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
         Left            =   3240
         TabIndex        =   6
         Top             =   720
         Width           =   735
      End
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash opciones 
      Height          =   1215
      Left            =   840
      TabIndex        =   11
      Top             =   3840
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
      Height          =   2895
      Left            =   840
      Top             =   600
      Width           =   6735
   End
End
Attribute VB_Name = "maestro04"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub dato1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF2 Then Call ayudadepto(dato2)
    Call flechas(dato1, dato2, KeyCode)
End Sub

Private Sub dato2_LostFocus()
    If sl = 0 Then leer
    sl = 0
End Sub

Private Sub dato2_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF2 Then Call ayudalinea(dato3)
    Call flechas(dato1, dato3, KeyCode)
End Sub

Private Sub dato3_KeyDown(KeyCode As Integer, Shift As Integer)
    Call flechas(dato2, dato4, KeyCode)
End Sub

Private Sub Form_Activate()
    dato1.SetFocus
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
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then Call Pregunta(dato4, dato5)
End Sub

Private Sub dato5_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then GRABAR: leer:
End Sub

Private Sub foto_DblClick()
    cargaFoto.Show vbModal
End Sub

Sub leer()
    campos(0, 0) = dato1.Tag
    campos(1, 0) = dato2.Tag
    campos(2, 0) = dato3.Tag
    campos(3, 0) = dato4.Tag
    campos(4, 0) = dato5.Tag
    campos(0, 2) = "maestrolineas"
    condicion = "codigodepto=" + "'" + dato2.text + "'"
    op = 5
    SQLUTIL.datos = campos
    Set SQLUTIL.conexion = db
    Call SQLUTIL.SQLUTIL(op, condicion)
    If SQLUTIL.estado = 0 Then carga: opciones.Visible = True: disponible (True): habilita (True): opciones.SetFocus
End Sub
Sub leersiguiente()
    campos(0, 0) = dato1.Tag 'CODIGODEPTO
    campos(1, 0) = dato2.Tag 'CODIGOLINEA
    campos(2, 0) = dato3.Tag 'NOMBRE
    campos(3, 0) = dato4.Tag 'DESCUENTO
    campos(4, 0) = dato5.Tag 'MARGEN
    campos(0, 2) = "maestrolineas"
    condicion = "codigolinea>" + "'" + dato2.text + "' order by codigolinea"
    op = 5
    SQLUTIL.datos = campos
    Set SQLUTIL.conexion = db
    Call SQLUTIL.SQLUTIL(op, condicion)
    If SQLUTIL.estado = 0 Then carga: opciones.Visible = True: disponible (True): habilita (True): opciones.SetFocus
End Sub
Sub leeranterior()
    campos(0, 0) = dato1.Tag 'CODIGODEPTO
    campos(1, 0) = dato2.Tag 'CODIGOLINEA
    campos(2, 0) = dato3.Tag 'NOMBRE
    campos(3, 0) = dato4.Tag 'DESCUENTO
    campos(4, 0) = dato5.Tag 'MARGEN
    campos(0, 2) = "maestrolineas"
    condicion = "codigolinea<" + "'" + dato2.text + "' order by codigolinea"
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
        .ms.ConnectionString = "DRIVER={MySQL ODBC 3.51 Driver};server=servidor;uid=root;pwd=123;database=conta01"
    End With
End Sub

Sub Pregunta(ByRef caja As TextBox, ByRef caja1 As TextBox)
    If caja.text = "" And sc = 0 Then caja.SetFocus
    If caja.text <> "" Or sc = 1 Then caja1.Enabled = True: caja1.SetFocus
    sc = 0
End Sub

Sub ayudadepto(ByRef caja As TextBox)
    Dim campos As Variant
    Dim cfijo As Variant
    campos = Array("codigodepto", "nombre")
    cfijo = Array("no")
    Call cargaAyudaT("servidor", "conta01", "root", "123", "maestrodepartamentos", dato1, campos, cfijo, 2)
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
    campos(4, 0) = dato5.Tag 'DESCUENTO
    campos(0, 1) = dato1.text 'CODIGO
    campos(1, 1) = dato2.text 'DESCRIPCION
    campos(2, 1) = dato3.text 'UNIDAD MEDIDA
    campos(3, 1) = dato4.text 'SECCION
    campos(4, 1) = dato5.text 'SECCION
    campos(0, 2) = "maestrolineas"
    If modifi = 1 Then condicion = "codigodepto=" + "'" + dato1.text + "'"
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
End Sub

Sub limpia()
    dato1.text = ""
    dato2.text = ""
    dato3.text = ""
    dato4.text = ""
End Sub
Sub ayudalinea(ByRef caja As TextBox)
    Dim campos As Variant
    Dim cfijo As Variant
    campos = Array("codigolinea", "nombre")
    cfijo = Array("codigodepto", dato1.text)
    Call cargaAyudaT("servidor", "conta01", "root", "123", "maestrolineas", dato2, campos, cfijo, 2)
    caja.Enabled = True
    caja.SetFocus
End Sub

