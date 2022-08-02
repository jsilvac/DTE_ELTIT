VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form ventas03 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mestro de Zonas"
   ClientHeight    =   4245
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7965
   DrawStyle       =   5  'Transparent
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   283
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
      Height          =   1575
      Left            =   840
      TabIndex        =   0
      Top             =   480
      Width           =   6015
      Begin VB.TextBox dato2 
         BackColor       =   &H00FBEDE6&
         Enabled         =   0   'False
         ForeColor       =   &H00000080&
         Height          =   285
         Left            =   1320
         MaxLength       =   30
         TabIndex        =   9
         Tag             =   "nombre"
         Top             =   720
         Width           =   2895
      End
      Begin VB.TextBox dato1 
         BackColor       =   &H00FBEDE6&
         ForeColor       =   &H00000080&
         Height          =   285
         Left            =   1320
         MaxLength       =   4
         TabIndex        =   8
         Tag             =   "codigozona"
         Top             =   360
         Width           =   495
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
         Left            =   4800
         TabIndex        =   7
         Top             =   720
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
         TabIndex        =   6
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
         TabIndex        =   5
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
         TabIndex        =   4
         Top             =   240
         Width           =   3255
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H00000080&
         BorderWidth     =   3
         FillColor       =   &H00FFC0C0&
         Height          =   1575
         Left            =   0
         Top             =   0
         Width           =   6015
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Nombre"
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Codigo"
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   855
      End
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash opciones 
      Height          =   1215
      Left            =   720
      TabIndex        =   3
      Top             =   2640
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
      Height          =   1575
      Left            =   960
      Top             =   600
      Width           =   6015
   End
End
Attribute VB_Name = "ventas03"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub dato1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF2 Then Call ayudazona(dato2)
    Call flechas(dato1, dato2, KeyCode)
End Sub

Private Sub dato1_LostFocus()
    If sl = 0 Then leer
    sl = 0
End Sub

Private Sub Form_Activate()
    dato1.SetFocus
End Sub

Private Sub Form_Load()
    Call Conectar_BD
    sc = 0
    opciones.Visible = False
End Sub

Private Sub DATO1_KeyPress(KeyAscii As Integer)
    snum = 1: KeyAscii = esNumero(KeyAscii)
    If KeyAscii = 13 Then Call ceros(dato1): Call Pregunta(dato1, dato2)
End Sub

Private Sub dato2_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then GRABAR: leer:
End Sub

Private Sub foto_DblClick()
    cargaFoto.Show vbModal
End Sub

Sub leer()
    campos(0, 0) = dato1.Tag 'CODIGO
    campos(1, 0) = dato2.Tag 'NOMBRE
    
    
    campos(0, 2) = "maestrozonas"
    condicion = "codigozona=" + "'" + dato1.text + "'"
    op = 5
    SQLUTIL.datos = campos
    Set SQLUTIL.conexion = db
    Call SQLUTIL.SQLUTIL(op, condicion)
    If SQLUTIL.ESTADO = 0 Then carga: opciones.Visible = True: disponible (True): habilita (True): opciones.SetFocus
End Sub
Sub leersiguiente()
    campos(0, 0) = dato1.Tag 'CODIGO
    campos(1, 0) = dato2.Tag 'NOMBRE
    
    
    campos(0, 2) = "maestrozonas"
    condicion = "codigozona>" + "'" + dato1.text + "' order by codigozona"
    op = 5
    SQLUTIL.datos = campos
    Set SQLUTIL.conexion = db
    Call SQLUTIL.SQLUTIL(op, condicion)
    If SQLUTIL.ESTADO = 0 Then carga: opciones.Visible = True: disponible (True): habilita (True): opciones.SetFocus
End Sub
Sub leeranterior()
    campos(0, 0) = dato1.Tag 'CODIGO
    campos(1, 0) = dato2.Tag 'NOMBRE
    
    
    campos(0, 2) = "maestrozonas"
    condicion = "codigozona<" + "'" + dato1.text + "' order by codigozona"
    op = 5
    SQLUTIL.datos = campos
    Set SQLUTIL.conexion = db
    Call SQLUTIL.SQLUTIL(op, condicion)
    If SQLUTIL.ESTADO = 0 Then carga: opciones.Visible = True: disponible (True): habilita (True): opciones.SetFocus
End Sub

Sub carga()
    habilita (True)
    dato1.text = SQLUTIL.datos(0, 3)
    dato2.text = SQLUTIL.datos(1, 3)
    
    
fin:
End Sub

Sub habilita(ByVal condicion As Boolean)
    dato1.Locked = condicion
    dato2.Locked = condicion
    
    
End Sub
Sub disponible(ByVal condicion As Boolean)
    dato1.Enabled = condicion
    dato2.Enabled = condicion
    
    
End Sub

Sub Pregunta(ByRef caja As TextBox, ByRef caja1 As TextBox)
    If caja.text = "" And sc = 0 Then caja.SetFocus
    If caja.text <> "" Or sc = 1 Then caja1.Enabled = True: caja1.SetFocus
    sc = 0
End Sub

Sub ayudazona(ByRef caja As TextBox)
    Dim campos As Variant
    Dim cfijo As Variant
    campos = Array("codigozona", "nombre")
    cfijo = Array("no")
    Call cargaAyudaT(servidor, basebus, USUARIO, password, "maestrozonas", dato1, campos, cfijo, 2)
    caja.Enabled = True
    caja.SetFocus
End Sub

Sub flechas(ByRef caja As TextBox, ByRef caja1 As TextBox, ByRef codigo As Integer)
    If codigo = 38 And caja.Enabled = True Then caja.SetFocus
    If codigo = 40 And caja1.Enabled = True Then caja1.SetFocus
End Sub

Sub GRABAR()
    campos(0, 0) = dato1.Tag 'CODIGO
    campos(1, 0) = dato2.Tag 'NOMRE
    
    
    campos(0, 1) = dato1.text 'CODIGO
    campos(1, 1) = dato2.text 'NOMBRE
    
    
    campos(0, 2) = "maestrozonas"
    If modifi = 1 Then condicion = "codigozona=" + "'" + dato1.text + "'"
    If modifi = 1 Then op = 3 Else op = 2
    modifi = 0
    SQLUTIL.datos = campos
    Set SQLUTIL.conexion = db
    Call SQLUTIL.SQLUTIL(op, condicion)
    status = SQLUTIL.ESTADO
End Sub
Sub ELIMINAR()
    campos(0, 2) = "maestrozonas"
    condicion = "codigozona=" + "'" + dato1.text + "'"
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
    

End Sub
