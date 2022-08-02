VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash9b.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{ADBBAED6-B16F-40DC-80DF-B44910CBA76C}#1.0#0"; "Frame-Xp.ocx"
Begin VB.Form arriendo04 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Maestro de Propiedades"
   ClientHeight    =   3015
   ClientLeft      =   2040
   ClientTop       =   1305
   ClientWidth     =   5820
   DrawStyle       =   5  'Transparent
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   201
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   388
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin XPFrame.FrameXp FrameXp1 
      Height          =   1170
      Left            =   120
      TabIndex        =   7
      Top             =   240
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   2064
      BackColor       =   16744576
      Caption         =   "DATOS  "
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
      Alignment       =   1
      Begin VB.TextBox dato1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
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
         Left            =   1725
         MaxLength       =   2
         TabIndex        =   0
         Tag             =   "codigomoneda"
         Top             =   315
         Width           =   495
      End
      Begin VB.TextBox dato2 
         BackColor       =   &H00C0FFFF&
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
         Left            =   1725
         MaxLength       =   50
         TabIndex        =   1
         Tag             =   "nombremoneda"
         Top             =   675
         Width           =   3615
      End
      Begin VB.Label lblmoneda 
         BackColor       =   &H00FF8080&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2160
         TabIndex        =   10
         Top             =   3555
         Width           =   4815
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Codigo"
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
         Left            =   100
         TabIndex        =   9
         Top             =   315
         Width           =   1530
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFC0C0&
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
         Left            =   100
         TabIndex        =   8
         Top             =   675
         Width           =   1530
      End
   End
   Begin VB.PictureBox MANUAL 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H80000002&
      ForeColor       =   &H80000008&
      Height          =   0
      Left            =   0
      ScaleHeight     =   0
      ScaleWidth      =   5790
      TabIndex        =   6
      Top             =   3015
      Width           =   5820
   End
   Begin VB.TextBox pivote 
      Height          =   285
      Left            =   8415
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   4230
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFF2F7&
      BorderStyle     =   0  'None
      Caption         =   "Datos personales"
      Height          =   3735
      Left            =   8400
      TabIndex        =   3
      Top             =   240
      Width           =   4695
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid SALDOS 
         Height          =   3495
         Left            =   120
         TabIndex        =   5
         Top             =   120
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
      Begin VB.Shape Shape3 
         BorderColor     =   &H00FF8080&
         BorderWidth     =   3
         FillColor       =   &H00FFC0C0&
         Height          =   3735
         Left            =   0
         Top             =   0
         Width           =   4695
      End
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash opciones 
      Height          =   1215
      Left            =   120
      TabIndex        =   2
      Top             =   1560
      Width           =   5655
      _cx             =   9975
      _cy             =   2143
      FlashVars       =   ""
      Movie           =   "c:\barra_opciones.swf"
      Src             =   "c:\barra_opciones.swf"
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
      AllowNetworking =   "all"
      AllowFullScreen =   "false"
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H00808080&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00E0E0E0&
      FillColor       =   &H00FF8080&
      Height          =   3735
      Left            =   8520
      Top             =   360
      Width           =   4695
   End
End
Attribute VB_Name = "arriendo04"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public saldoglobal As Double
Private MODIFI As Integer

Private Sub dato1_GotFocus()
Call cargatexto(dato1)
End Sub

Private Sub dato2_GotFocus()
If MODIFI = 0 Then Call leer
Call cargatexto(dato2)
End Sub
  
Private Sub dato1_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyF2 Then Call ayudamonedas(dato1)
    Call flechas(dato1, dato2, KeyCode)
End Sub
Private Sub dato2_KeyDown(KeyCode As Integer, Shift As Integer)
       Call flechas(dato1, dato2, KeyCode)
End Sub
  
 Private Sub MANUAL_KeyPress(KeyAscii As Integer)
If UCase(Chr(KeyAscii)) = "M" Then Call opciones_FSCommand("modifica", "")
If UCase(Chr(KeyAscii)) = "E" Then Call opciones_FSCommand("elimina", "")
If UCase(Chr(KeyAscii)) = "S" Then Call opciones_FSCommand("siguiente", "")
If UCase(Chr(KeyAscii)) = "A" Then Call opciones_FSCommand("anterior", "")
If UCase(Chr(KeyAscii)) = "R" Then Call opciones_FSCommand("retorno", "")
If UCase(Chr(KeyAscii)) = "I" Then Call opciones_FSCommand("imprime", "")
End Sub

Private Sub Form_Load()
Call CENTRAR(Me)

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
    If KeyAscii = 13 And dato1.text <> "" Then
        Call ceros(dato1)
        Call Pregunta(dato1, dato2)
    End If
End Sub
Private Sub dato2_KeyPress(KeyAscii As Integer)
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If KeyAscii = 13 And dato2.text <> "" Then
            grabar
            retorno
        End If
End Sub

Sub leer()
    CAMPOS(0, 0) = dato1.Tag
    CAMPOS(1, 0) = dato2.Tag
    CAMPOS(2, 0) = ""
    CAMPOS(0, 2) = clientesistema & "arriendos" & ".maestro_monedas"
    condicion = "codigomoneda = '" & dato1.text & "' "

    op = 5
    sqlconta.response = CAMPOS
    Set sqlconta.conexion = db
    Call sqlconta.sqlconta(op, condicion)
    If sqlconta.status = 4 Then dato2.SetFocus: GoTo no:
    carga
    opciones.Visible = True
    disponible (True)
    habilita (True)
    opciones.SetFocus
        
no:
End Sub
Sub leersiguiente()
    CAMPOS(0, 0) = dato1.Tag
    CAMPOS(1, 0) = dato2.Tag
    CAMPOS(2, 0) = ""
    CAMPOS(0, 2) = clientesistema & "arriendos" & ".maestro_monedas"
    condicion = " codigomoneda > '" & dato1.text & "' order by codigomoneda asc "

    op = 5
    sqlconta.response = CAMPOS
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
    CAMPOS(0, 0) = dato1.Tag
    CAMPOS(1, 0) = dato2.Tag
    CAMPOS(2, 0) = ""
    CAMPOS(0, 2) = clientesistema & "arriendos" & ".maestro_monedas"
    condicion = " codigomoneda < '" & dato1.text & "' order by codigomoneda desc "
    
    op = 5
    sqlconta.response = CAMPOS
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

 
Sub ayudamonedas(ByRef caja As TextBox)
    Dim CAMPOS As Variant
    Dim cfijo As Variant
    Dim largo As Variant
    CAMPOS = Array("codigomoneda", "nombremoneda")
    largo = Array("11s", "40s")
    cfijo = "codigomoneda like '%%'"
    cabezas = Array("Codigo", "Nombre")
    mensajeAyuda = "Ayuda de Monedas"
    Call cargaAyudaT(servidor, clientesistema & "arriendos", usuario, password, ".maestro_monedas", caja, CAMPOS, cfijo, largo, 2)
    caja.Enabled = True
    caja.SetFocus
End Sub
 

Sub flechas(ByRef caja As TextBox, ByRef caja1 As TextBox, ByRef codigo As Integer)
    If codigo = 38 And caja.Enabled = True Then caja.SetFocus
    If codigo = 40 And caja1.Enabled = True Then caja1.SetFocus
End Sub

Sub grabar()
    CAMPOS(0, 0) = dato1.Tag
    CAMPOS(1, 0) = dato2.Tag
    CAMPOS(2, 0) = ""
   
    CAMPOS(0, 1) = dato1.text
    CAMPOS(1, 1) = dato2.text
   
    CAMPOS(0, 2) = clientesistema & "arriendos" & ".maestro_monedas"
    If MODIFI = 1 Then condicion = "codigomoneda ='" & dato1.text & "'"
    If MODIFI = 1 Then op = 3 Else op = 2
    sqlconta.response = CAMPOS
    Set sqlconta.conexion = db
    Call sqlconta.sqlconta(op, condicion)
    
    End Sub
 
Sub ELIMINAR()
    CAMPOS(0, 2) = clientesistema & "arriendos" & ".maestro_monedas"
    condicion = "codigomoneda='" & dato1.text & "'"
    op = 4
    sqlconta.response = CAMPOS
    Set sqlconta.conexion = db
    Call sqlconta.sqlconta(op, condicion)

    
End Sub
  

Private Sub opciones_FSCommand(ByVal command As String, ByVal args As String)

If command = "retorno" Then retorno
If command = "modifica" Then modifica

If command = "elimina" Then
    If Verifica_Permiso(Me.Caption, "elimina") = True Then
        elimina
    End If
End If
If command = "siguiente" Then leersiguiente
If command = "anterior" Then leeranterior
 



End Sub
Sub elimina()
 
disponible (True)
habilita (False)
ELIMINAR
limpia
opciones.Visible = False
dato1.SetFocus
 
End Sub

Sub modifica()
disponible (True)
habilita (False)
dato1.Enabled = False
dato2.SetFocus
MODIFI = 1

End Sub
Sub retorno()

disponible (True)
habilita (False)
limpia
opciones.Visible = False
dato1.Enabled = True
dato1.SetFocus
MODIFI = 0
no:
 
End Sub
Sub limpia()
    dato1.text = ""
    dato2.text = ""
End Sub
 
Sub cargatexto(ByRef caja As TextBox)
caja.SelStart = 0: caja.SelLength = Len(caja.text)
End Sub

Private Sub opciones_GotFocus()
MANUAL.SetFocus
End Sub
