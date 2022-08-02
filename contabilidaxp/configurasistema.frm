VERSION 5.00
Object = "{085FBDF3-00C5-421B-B762-1D57299A2B89}#1.0#0"; "CLBUTN.OCX"
Object = "{ADBBAED6-B16F-40DC-80DF-B44910CBA76C}#1.0#0"; "Frame-Xp.ocx"
Begin VB.Form confi01 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   3360
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   5985
   DrawStyle       =   5  'Transparent
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   224
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   399
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox CONTROL 
      Height          =   15
      Left            =   4440
      ScaleHeight     =   15
      ScaleWidth      =   1095
      TabIndex        =   7
      Top             =   2640
      Width           =   1095
   End
   Begin XPFrame.FrameXp FrameXp1 
      Height          =   3015
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   5318
      BackColor       =   16761024
      Caption         =   "CONFIGURA EMPRESA A UTILIZAR"
      CaptionEstilo3D =   1
      BackColor       =   16761024
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
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   1320
         MaxLength       =   2
         TabIndex        =   0
         Tag             =   "codigoempresa"
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
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   1320
         MaxLength       =   50
         TabIndex        =   1
         Tag             =   "nombre"
         Top             =   840
         Width           =   4335
      End
      Begin VB.TextBox dato4 
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
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   1320
         MaxLength       =   50
         TabIndex        =   3
         Tag             =   "comuna"
         Top             =   1560
         Width           =   3615
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
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   1320
         MaxLength       =   50
         TabIndex        =   2
         Tag             =   "direccion"
         Top             =   1200
         Width           =   4335
      End
      Begin VB.TextBox dato5 
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
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   1320
         MaxLength       =   50
         TabIndex        =   4
         Tag             =   "ciudad"
         Top             =   1920
         Width           =   3615
      End
      Begin VB.TextBox dato6 
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
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   1320
         MaxLength       =   10
         TabIndex        =   5
         Tag             =   "rut"
         Top             =   2280
         Width           =   2175
      End
      Begin CoolButtons.cool_Button opciones 
         Height          =   375
         Left            =   4230
         TabIndex        =   14
         Top             =   2520
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
         Caption         =   "Retorno"
         ForeColor       =   16711680
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FF8080&
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
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   480
         Width           =   975
      End
      Begin VB.Label Label5 
         BackColor       =   &H00FF8080&
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
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   840
         Width           =   975
      End
      Begin VB.Label Label6 
         BackColor       =   &H00FF8080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Direccion"
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
         TabIndex        =   11
         Top             =   1200
         Width           =   975
      End
      Begin VB.Label Label10 
         BackColor       =   &H00FF8080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Comuna"
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
         Top             =   1560
         Width           =   975
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FF8080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Ciudad"
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
         TabIndex        =   9
         Top             =   1920
         Width           =   975
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FF8080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Rut"
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
         TabIndex        =   8
         Top             =   2280
         Width           =   975
      End
   End
End
Attribute VB_Name = "confi01"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CONTROL_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Salir
End Sub

Private Sub dato1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF2 Then Call ayudaempresa(dato2)
        Call flechas(dato1, dato1, KeyCode)
End Sub

Private Sub dato1_LostFocus()
    If sl = 0 Then leer
sl = 0
End Sub





Private Sub Form_Activate()
If FormularioActivo("prove0001") = True Then
MsgBox "IMPOSIBLE CAMBIAR EMPRESA CON " + prove0001.Caption + " ABIERTO "
Unload Me
End If

If FormularioActivo("prove0002") = True Then
MsgBox "IMPOSIBLE CAMBIAR EMPRESA CON " + prove0002.Caption + " ABIERTO "
Unload Me
End If

If FormularioActivo("prove0003") = True Then
MsgBox "IMPOSIBLE CAMBIAR EMPRESA CON " + prove0003.Caption + " ABIERTO "
Unload Me
End If

If FormularioActivo("prove0003") = True Then
MsgBox "IMPOSIBLE CAMBIAR EMPRESA CON " + prove0003.Caption + " ABIERTO "
Unload Me
End If

If FormularioActivo("prove0003") = True Then
MsgBox "IMPOSIBLE CAMBIAR EMPRESA CON " + prove0003.Caption + " ABIERTO "
Unload Me
End If

If FormularioActivo("proceso03") = True Then
MsgBox "IMPOSIBLE CAMBIAR EMPRESA CON " + proceso03.Caption + " ABIERTO "
Unload Me
End If

If FormularioActivo("proceso04") = True Then
MsgBox "IMPOSIBLE CAMBIAR EMPRESA CON " + proceso04.Caption + " ABIERTO "
Unload Me
End If
If FormularioActivo("proceso05") = True Then
MsgBox "IMPOSIBLE CAMBIAR EMPRESA CON " + proceso05.Caption + " ABIERTO "
Unload Me
End If
If FormularioActivo("proceso06") = True Then
MsgBox "IMPOSIBLE CAMBIAR EMPRESA CON " + proceso06.Caption + " ABIERTO "
Unload Me
End If



'If prove0002.Visible = True Then
'MsgBox "IMPOSIBLE CAMBIAR EMPRESA CON " + prove0002.Caption + " ABIERTO "
'Unload Me
'End If

'If prove0001.Visible = True Then
'MsgBox "IMPOSIBLE CAMBIAR EMPRESA CON " + prove0001.Caption + " ABIERTO "
'Unload Me
'End If
'If prove0003.Visible = True Then
'MsgBox "IMPOSIBLE CAMBIAR EMPRESA CON " + prove0003.Caption + " ABIERTO "
'Unload Me
'End If
'If proceso03.Visible = True Then
'MsgBox "IMPOSIBLE CAMBIAR EMPRESA CON " + proceso03.Caption + " ABIERTO "
'Unload Me
'End If
'If proceso04.Visible = True Then
'MsgBox "IMPOSIBLE CAMBIAR EMPRESA CON " + proceso04.Caption + " ABIERTO "
'Unload Me
'End If
'If proceso05.Visible = True Then
'MsgBox "IMPOSIBLE CAMBIAR EMPRESA CON " + proceso05.Caption + " ABIERTO "
''Unload Me
'End If
'If proceso06.Visible = True Then
'MsgBox "IMPOSIBLE CAMBIAR EMPRESA CON " + proceso06.Caption + " ABIERTO "
'Unload Me
'End If


End Sub

Private Sub Form_Load()
Close 20

    Call Conectarconta(Servidor, clientesistema + "conta", Usuario, password)
    sqlconta.area = clientesistema + "conta"
    SQlcontabilidad.cliente_sql = clientesistema
      Call Conectar_Auditoria
    Set sqlconta.conAuditoria = conexionauditoria
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
    If KeyAscii = 13 Then Call Pregunta(dato4, DATO5)
End Sub
Private Sub dato5_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then Call Pregunta(DATO5, dato6)
End Sub
Private Sub dato6_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then Call Pregunta(dato6, dato6)
End Sub



Sub leer()
    campos(0, 0) = dato1.Tag 'CODIGO
    campos(1, 0) = dato2.Tag 'BODEGA
    campos(2, 0) = dato3.Tag 'DIRECCION
    campos(3, 0) = dato4.Tag 'CIUDAD
    campos(4, 0) = DATO5.Tag 'OTROS
    campos(5, 0) = dato6.Tag
    campos(6, 0) = ""
    campos(0, 2) = "maestroempresas"
    condicion = "codigoempresa=" + "'" + dato1.text + "'"
    op = 5
    sqlconta.response = campos
    Set sqlconta.conexion = conta
    
    
    
    Call sqlconta.sqlconta(op, condicion)
    If sqlconta.status = 0 Then carga: opciones.Visible = True: disponible (True): habilita (True): opciones.SetFocus
    If sqlconta.status = 4 Then dato1.SetFocus
    
End Sub
Sub carga()
    habilita (True)
    dato1.text = sqlconta.response(0, 3)
    dato2.text = sqlconta.response(1, 3)
    dato3.text = sqlconta.response(2, 3)
    dato4.text = sqlconta.response(3, 3)
    DATO5.text = sqlconta.response(4, 3)
    dato6.text = sqlconta.response(5, 3)
    
fin:
End Sub

Sub habilita(ByVal condicion As Boolean)
    
    dato1.Locked = condicion
    dato2.Locked = condicion
    dato3.Locked = condicion
    dato4.Locked = condicion
    DATO5.Locked = condicion
    dato6.Locked = condicion
    
    End Sub
Sub disponible(ByVal condicion As Boolean)
    
    dato1.Enabled = condicion
    dato2.Enabled = condicion
    dato3.Enabled = condicion
    dato4.Enabled = condicion
    DATO5.Enabled = condicion
    dato6.Enabled = condicion
    
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

    campos(0, 0) = dato1.Tag 'CODIGO
    campos(1, 0) = dato2.Tag 'BODEGA
    campos(2, 0) = dato3.Tag 'DIRECCION
    campos(3, 0) = dato4.Tag 'CIUDAD
    campos(4, 0) = DATO5.Tag 'OTROS
    campos(5, 0) = dato6.Tag
    
    
    
  
    campos(0, 1) = dato1.text 'CODIGO
    campos(1, 1) = dato2.text 'BODEGA
    campos(2, 1) = dato3.text 'DIRECCION
    campos(3, 1) = dato4.text 'CIUDAD
    campos(4, 1) = DATO5.text 'OTROS
    campos(5, 1) = dato6.text
    
        
    campos(0, 2) = "maestroempresas"
    If MODIFI = 1 Then condicion = "codigoempresa=" + "'" + dato1.text + "'"
    If MODIFI = 1 Then op = 3 Else op = 2
    MODIFI = 0
    sqlconta.response = campos
    Set sqlconta.conexion = conta
    
    Call sqlconta.sqlconta(op, condicion)
    status = sqlconta.status




End Sub

Sub limpia()

    dato1.text = ""
    dato2.text = ""
    dato3.text = ""
    dato4.text = ""
    DATO5.text = ""
    dato6.text = ""
    End Sub

Private Sub opciones_Click()
    Salir


End Sub
Sub Salir()

grabasegurity
empresa

End Sub
Sub grabasegurity()
    campos(0, 0) = "usuario"
    campos(1, 0) = "fechaingreso"
    campos(2, 0) = "horaingreso"
    campos(3, 0) = "programaactivo"
    campos(4, 0) = "nombreprogramaactivo"
    campos(5, 0) = "empresaactiva"
    campos(6, 0) = ""
    campos(0, 1) = USUARIOSISTEMA
    campos(1, 1) = Mid(Date$, 7, 4) + Mid(Date$, 4, 2) + Mid(Date$, 1, 2)
    campos(2, 1) = Time$
    campos(3, 1) = "PRINCIPAL"
    campos(4, 1) = "MENU PRINCIPAL "
    campos(5, 1) = dato1.text
    campos(0, 2) = "usuariosactivos"
    condicion = "usuario = '" + USUARIOSISTEMA + "'"
    op = 3
    sqlconta.response = campos
    Set sqlconta.conexion = conta
    Call sqlconta.sqlconta(op, condicion)
End Sub

Sub empresa()
    empresaactiva = dato1.text
    
    campos(0, 0) = "codigoempresa"
    campos(1, 0) = "nombre"
    campos(2, 0) = "direccion"
    campos(3, 0) = "comuna"
    campos(4, 0) = "ciudad"
    campos(5, 0) = "rut"
    campos(6, 0) = "cuentaproveedor"
    campos(7, 0) = "cuentahonorarios"
    campos(8, 0) = "cuentaclientes"
    campos(9, 0) = "ivacredito"
    campos(10, 0) = "ivadebito"
    campos(11, 0) = "retencionhonorarios"
    campos(12, 0) = "cuentaperdida"
    campos(13, 0) = "cuentaganancia"
    campos(14, 0) = "auditoria"
    campos(15, 0) = "codigoae"
    campos(16, 0) = "representantelegal"
    campos(17, 0) = "rutenviasii"
    campos(18, 0) = "emailcontable"
    campos(19, 0) = "emailcontable"
    campos(20, 0) = "mercaderias"
    campos(21, 0) = "ingresosporventa"
    campos(22, 0) = "ivaretenido"
    campos(23, 0) = "codigosii"
    campos(24, 0) = "rubro"
    campos(25, 0) = "cuentadiferencia"
    campos(26, 0) = "diacierrecompras"
    campos(27, 0) = "fechacierre"
    campos(28, 0) = "rutenviasii"
    campos(29, 0) = "fecharesolucion"
    campos(30, 0) = "numeroresolucion"
    campos(31, 0) = "empresafae"
    campos(32, 0) = "cuentaprovisiones"
    campos(33, 0) = "cuentagastoferia"
    campos(34, 0) = "certificado"
    campos(35, 0) = "clave_certificado"
    
    campos(36, 0) = ""
    
    campos(0, 2) = "maestroempresas"
  
    condicion = "codigoempresa=" + "'" + empresaactiva + "'" + " ORDER BY codigoempresa"
    op = 5
    sqlconta.response = campos
    Set sqlconta.conexion = conta
    
    Call sqlconta.sqlconta(op, condicion)
    cuentadiferencia = sqlconta.response(25, 3)
    CONFI_EMPRESAFAE = sqlconta.response(31, 3)
    
    rubro = sqlconta.response(24, 3)
    
    rutempresa = sqlconta.response(5, 3)
    direccionempresa = sqlconta.response(2, 3)
    comunaempresa = sqlconta.response(3, 3)
    codigosii = sqlconta.response(23, 3)
    codigoempresa = sqlconta.response(0, 3)
    nombreempresa = sqlconta.response(1, 3)
    CUENTAPROVEEDOR = sqlconta.response(6, 3)
    cuentahonorarios = sqlconta.response(7, 3)
    cuentacliente = sqlconta.response(8, 3)
    cuentaperdida = sqlconta.response(12, 3)
    cuentaganancia = sqlconta.response(13, 3)
    codigoae = sqlconta.response(15, 3)
    ivadebito = sqlconta.response(10, 3)
    ivacredito = sqlconta.response(9, 3)
    retencion = sqlconta.response(11, 3)
    
    mercaderias = sqlconta.response(20, 3)
    ingresosporventa = sqlconta.response(21, 3)
    
    ivaretenido = sqlconta.response(22, 3)
    
    diacierrecompra = sqlconta.response(26, 3)
    fechacierre = sqlconta.response(27, 3)
    rut_representante = sqlconta.response(17, 3)
    nombre_representante = sqlconta.response(16, 3)
    cuentaprovisiones = sqlconta.response(32, 3)
    cuentagastoferia = sqlconta.response(33, 3)
    certificado_sii = sqlconta.response(34, 3)
    clave_certificado_sii = sqlconta.response(35, 3)
    
    For k = 0 To 22
    DATOSEMPRESA(k) = sqlconta.response(k, 3)
    Next k
    
    rut_enviasii = sqlconta.response(28, 3)
    fecharesolucion = Format(sqlconta.response(29, 3), "yyyy-mm-dd")
    numeroresolucion = sqlconta.response(30, 3)
    If numeroresolucion = 0 Then
    f3327 = True
    f3328 = True
    Else
    f3327 = False
    f3328 = False
    End If
    
    
    
   Rem  sqlconta.basededatos = clientesistema + empresaactiva
    
    leercorreospagoproveedores
    
    Call configurabasededatos
Unload segurity
Unload Me


PRINCIPAL.Show
        
        End Sub



Sub leercorreospagoproveedores()
    Dim csql As New rdoQuery
    Dim resultados As rdoResultset
    
    Set csql.ActiveConnection = conta
    csql.sql = "select correo,servidor,clave,ruta from correo_pagoproveedores "
    csql.Execute
    
        UsuarioCorreo = ""
        ServerCorreo = ""
        ClaveCorreo = ""
        
    If csql.RowsAffected > 0 Then
        Set resultados = csql.OpenResultset
        UsuarioCorreo = resultados(0)
        ServerCorreo = resultados(1)
        ClaveCorreo = resultados(2)
        RUTACOMPROBANTE = resultados(3)
    End If
      
      csql.Close
      Set csql = Nothing
      
      
 
    
End Sub
Private Sub opciones_GotFocus()
CONTROL.SetFocus
'CONTROL.SetFocus
'opciones.SetFocus


End Sub
Sub ayudaempresa(ByRef caja As TextBox)
    Dim campos As Variant
    Dim cfijo As Variant
    Dim largo As Variant
    campos = Array("codigoempresa", "nombre")
    largo = Array("6s", "40s")
    cfijo = "no"
    basebus = clientesistema + "conta"
    Call cargaAyudaT(Servidor, basebus, Usuario, password, "maestroempresas", dato1, campos, cfijo, largo, 2)
    caja.Enabled = True
    caja.SetFocus
    
End Sub

