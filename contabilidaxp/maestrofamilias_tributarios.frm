VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash10d.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{ADBBAED6-B16F-40DC-80DF-B44910CBA76C}#1.0#0"; "Frame-Xp.ocx"
Begin VB.Form activofijotb03 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Maestro de Familias IPC Tributario"
   ClientHeight    =   3945
   ClientLeft      =   2235
   ClientTop       =   1305
   ClientWidth     =   8250
   DrawStyle       =   5  'Transparent
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   263
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   550
   ShowInTaskbar   =   0   'False
   Begin XPFrame.FrameXp FrameXp1 
      Height          =   1815
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   3201
      BackColor       =   16761024
      Caption         =   "Datos"
      CaptionEstilo3D =   1
      BackColor       =   16761024
      ForeColor       =   8438015
      BordeColor      =   -2147483635
      ColorBarraArriba=   4194304
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.CheckBox chk2 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Depreciación"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5160
         TabIndex        =   8
         Top             =   1320
         Width           =   1455
      End
      Begin VB.CheckBox chk1 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Corrección Monetaria"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2640
         TabIndex        =   7
         Top             =   1320
         Width           =   2175
      End
      Begin VB.TextBox dato1 
         Alignment       =   1  'Right Justify
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
         Left            =   2640
         MaxLength       =   2
         TabIndex        =   4
         Tag             =   "codigo"
         Top             =   480
         Width           =   495
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
         Left            =   2640
         MaxLength       =   80
         TabIndex        =   3
         Top             =   840
         Width           =   5295
      End
      Begin VB.Label Label1 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Codigo "
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
         Top             =   480
         Width           =   2295
      End
      Begin VB.Label Label3 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Nombre"
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
         TabIndex        =   5
         Top             =   840
         Width           =   2295
      End
   End
   Begin VB.TextBox pivote 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   3000
      Visible         =   0   'False
      Width           =   615
   End
   Begin MSAdodcLib.Adodc mcm 
      Height          =   375
      Left            =   480
      Top             =   7440
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
      Left            =   4920
      TabIndex        =   9
      Top             =   1920
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
      Begin VB.CommandButton botonmisfavoritos 
         Caption         =   "Mis Favoritos"
         Height          =   255
         Left            =   1800
         TabIndex        =   11
         Top             =   280
         Width           =   1335
      End
      Begin VB.CommandButton botonmisaccesos 
         Caption         =   "Permisos Modulo"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   280
         Width           =   1455
      End
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash opciones 
      Height          =   1215
      Left            =   120
      TabIndex        =   1
      Top             =   2640
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
Attribute VB_Name = "activofijotb03"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public saldoglobal As Double
Public saldocuenta As Double
Public leido As Boolean

 
Private Sub chk1_Click()
   If leido = True Then
         dato2.SetFocus
   End If
End Sub
Private Sub chk2_Click()
    If leido = True Then
        dato2.SetFocus
    End If
End Sub

Private Sub dato1_GotFocus()
    Call cargatexto(dato1)
End Sub
Private Sub dato2_GotFocus()
   If MODIFI = 0 Then
    Call leer
    End If
    Call cargatexto(dato2)
End Sub

Private Sub dato1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 38 Then Unload Me: GoTo no:
     If KeyCode = vbKeyF2 Then Call ayudafamilia(dato1)
    Call flechas(dato1, dato2, KeyCode)
no:
End Sub
Private Sub dato2_KeyDown(KeyCode As Integer, Shift As Integer)
    Call flechas(dato1, dato2, KeyCode)
End Sub
 
Sub ayudafamilia(ByRef caja As TextBox)
    Dim campos As Variant
    Dim cfijo As Variant
    Dim largo As Variant
    campos = Array("codigo", "nombre")
    largo = Array("12n", "40s")
    cfijo = "no"
    cabezas = Array("Codigo", "Nombre")
    mensajeAyuda = "Ayuda Familias"
    
    Call cargaAyudaT(servidor, clientesistema + "conta", Usuario, password, "maestro_familias_tributario", pivote, campos, cfijo, largo, 2)

    If Val(pivote.text) = 0 Then caja.SetFocus: GoTo no
     
    caja.text = pivote.text
    caja.Enabled = True
    caja.SetFocus

no:

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

End Sub
 
Private Sub dato1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then Unload Me
    snum = 0: KeyAscii = esNumero(KeyAscii)
    If KeyAscii = 13 Then
    Call ceros(dato1)
    Call Pregunta(dato1, dato2)
    End If
End Sub

Private Sub dato2_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 And dato2.text <> "" Then: Call grabar: retorno: leer
 End Sub

  

Sub leer()
    campos(0, 0) = "codigo"
    campos(1, 0) = "nombre"
    campos(2, 0) = "correccion_monetaria"
    campos(3, 0) = "depreciacion"
    campos(4, 0) = ""
    campos(0, 2) = "maestro_familias_tributario"
    condicion = "codigo= '" + dato1.text + "' "

    op = 5
    sqlconta.response = campos
    Set sqlconta.conexion = conta
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
    campos(0, 0) = "codigo"
    campos(1, 0) = "nombre"
    campos(2, 0) = "correccion_monetaria"
    campos(3, 0) = "depreciacion"
    campos(4, 0) = ""
    
    campos(0, 2) = "maestro_familias_tributario"
    condicion = "codigo> '" + dato1.text + "' order by codigo"
    op = 5
    sqlconta.response = campos
    Set sqlconta.conexion = db
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
    campos(0, 0) = "codigo"
    campos(1, 0) = "nombre"
    campos(2, 0) = "correccion_monetaria"
    campos(3, 0) = "depreciacion"
    campos(4, 0) = ""
    
    campos(0, 2) = "maestro_familias_tributario"
    condicion = "codigo< '" + dato1.text + "'  order by codigo desc"

    op = 5
    sqlconta.response = campos
    Set sqlconta.conexion = db
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
    dato1.text = sqlconta.response(0, 3)
    dato2.text = sqlconta.response(1, 3)
    leido = False
    chk1.Value = sqlconta.response(2, 3)
    chk2.Value = sqlconta.response(3, 3)
    leido = True
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

Sub flechas(ByRef caja As TextBox, ByRef caja1 As TextBox, ByRef codigo As Integer)
    If codigo = 38 And caja.Enabled = True Then caja.SetFocus
    If codigo = 40 And caja1.Enabled = True Then caja1.SetFocus
End Sub
Sub grabar()

    campos(0, 0) = "codigo"
    campos(1, 0) = "nombre"
    campos(2, 0) = "correccion_monetaria"
    campos(3, 0) = "depreciacion"
    campos(4, 0) = ""
    
    campos(0, 1) = dato1.text
    campos(1, 1) = dato2.text
    campos(2, 1) = chk1.Value
    campos(3, 1) = chk2.Value
    
    
    campos(0, 2) = "maestro_familias_tributario"
    If MODIFI = 1 Then condicion = "codigo='" & dato1.text & "'"
    If MODIFI = 1 Then op = 3 Else op = 2
    
    sqlconta.response = campos
    Set sqlconta.conexion = conta
    Call sqlconta.sqlconta(op, condicion)
    MODIFI = 0
End Sub
 

Sub ELIMINAR()
    campos(0, 2) = "maestro_familias_tributario"
    condicion = "codigo= '" + dato1.text + "' "
    op = 4
    sqlconta.response = campos
    Set sqlconta.conexion = conta
    Call sqlconta.sqlconta(op, condicion)
    
End Sub


Private Sub Label18_Click()

End Sub

Private Sub lblhistorico_Click(Index As Integer)

End Sub

Private Sub Frame2_DragDrop(Source As CONTROL, x As Single, Y As Single)

End Sub

 

Private Sub opciones_FSCommand(ByVal command As String, ByVal args As String)

If command = "retorno" Then retorno
    If command = "modifica" Then
        Call disponible(True)
        Call habilita(False)
        dato1.Enabled = False
        dato2.Enabled = True
        chk1.Enabled = True
        chk2.Enabled = True
        dato2.SetFocus
        MODIFI = 1
    End If
If command = "elimina" Then
    If Verifica_Permiso(Me.Caption, "elimina") = True Then
        ELIMINAR
        retorno
    End If
End If
If command = "siguiente" Then leersiguiente
If command = "anterior" Then leeranterior
If command = "imprime" Then imprimir

End Sub
Sub retorno()
disponible (True)
habilita (False)
limpia
opciones.Visible = False
dato1.Enabled = True
dato1.SetFocus
 
End Sub
Sub limpia()
    dato1.text = ""
    dato2.text = ""
    chk1.Value = 0
    chk2.Value = 0
End Sub

Sub imprimir()
    
End Sub
Sub cabeza()
    
End Sub


Private Sub botonmisaccesos_Click()
programafiltro = Me.Caption
misaccesos.Show
End Sub


Private Sub botonmisfavoritos_Click()
Call AgregaFavoritos(USUARIOSISTEMA, App.EXEName, Me.Name, Me.Caption)

End Sub
