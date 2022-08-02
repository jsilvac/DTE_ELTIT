VERSION 5.00
Object = "{ADBBAED6-B16F-40DC-80DF-B44910CBA76C}#1.0#0"; "Frame-Xp.ocx"
Begin VB.Form segurity 
   Caption         =   "SISTEMA DE CONTABILIDAD"
   ClientHeight    =   10680
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   11400
   Icon            =   "segurity.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   712
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   760
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin XPFrame.FrameXp seguridad 
      Height          =   2175
      Left            =   4080
      TabIndex        =   0
      Top             =   1920
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   3836
      BackColor       =   192
      Caption         =   "SISTEMA DE SEGURIDAD ACTIVADO"
      CaptionEstilo3D =   1
      BackColor       =   192
      BordeColor      =   -2147483638
      ColorBarraAbajo =   12582912
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
      Begin VB.TextBox DATO1 
         BackColor       =   &H00C0E0FF&
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
         Left            =   3240
         TabIndex        =   1
         Top             =   840
         Width           =   3375
      End
      Begin VB.TextBox CLAVE 
         BackColor       =   &H00C0E0FF&
         ForeColor       =   &H00800000&
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   3240
         PasswordChar    =   "*"
         TabIndex        =   2
         Top             =   1200
         Width           =   1575
      End
      Begin VB.Image Image2 
         Height          =   1290
         Left            =   240
         Picture         =   "segurity.frx":5A4A
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label Label1 
         BackColor       =   &H000000C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "USUARIO"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   2040
         TabIndex        =   4
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label Label2 
         BackColor       =   &H000000C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "CLAVE"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   2040
         TabIndex        =   3
         Top             =   1200
         Width           =   1095
      End
   End
   Begin VB.Image Image1 
      Height          =   10620
      Left            =   0
      Picture         =   "segurity.frx":612E
      Stretch         =   -1  'True
      Top             =   0
      Width           =   15360
   End
End
Attribute VB_Name = "segurity"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub confsistema_Click()

'maestro05.Show

End Sub

Private Sub CLAVE_LostFocus()
If CLAVE.text = "" Then GoTo no:
If CLAVE.text = clavesistema Then enviar
no:
End Sub

Private Sub Command1_Click()
Unload Me

End Sub

Private Sub dato3_Change()

End Sub

Private Sub Form_Load()
    '==================================
    'PERMITE UNA INSTANCIA DEL SISTEMA
    '==================================
Dim pat1 As String
Dim pat2 As String
Dim pater(10) As String
Close 20
    Open App.path + "\configu.txt" For Input As #20
12:    Input #20, SS
    If EOF(20) Then GoTo fin:
    If Mid(SS, 1, 8) = "SERVIDOR" Then
        Servidor = Mid(SS, 10, Len(SS) - 9)
    End If
     
   
    
    GoTo 12
fin:
        pat1 = "_licencia"
        pat2 = "mifranchitaflan"
    pater(1) = "erp_"
    pater(2) = "licencia_"
    pater(3) = "775753404"
    
'     Usuario = "erp_licencia"
'        password = "erp_licencia_775753404"
'
        

Call Conectartemporal(Servidor, "adminerp_inicio", "erp" + pat1, pater(1) + pater(2) + pater(3))


'  Usuario = "adminerp" + pat1
'  password = pater(1) + pater(2) + pater(3)
Call leerdatosconeccion("contabilidadxp.exe")
'   Usuario = "adminerp" + pat1
'  password = pater(1) + pater(2) + pater(3)
  
    basedatos = clientesistema + "conta"
'    Call Conectarconta(Servidor, basedatos, Usuario, password)
     mensaje_nopermiso = "Usted no tiene privilegios suficientes para realizar esta operación."
    cabezas = Array("codigo", "nombre")
    mensajeAyuda = "Ayuda General"
    Call COMPARAPROGRAMAS

dwlen = MAX_COMPUTERNAME_LENGTH + 1
strstring = String(dwlen, "X")
GetComputerName strstring, dwlen
strstring = Left(strstring, dwlen)
nombremaquina = strstring
ipfija = GetWanIP
ipusada = ipfija
'ipusada = DateAdd("m", -6, "2019-11-30")
'     ipfija = "192.168.4.50"
 If VerificaAplicacion("VB6.EXE") = True Then USUARIOSISTEMA = "CESAR": enviar: Exit Sub

   If estaactivo(ipfija, "contabilidadxp.exe") = True Then
    enviar
    
    Else
    MsgBox "INTENTO DE ENTRADA FALSO DEBE SER POR MENU PRINCIPAL "
    
     End
    End If
    
    
End Sub



Private Sub dato1_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 13 Then CLAVE.SetFocus
   
End Sub
Private Sub clave_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then

   End If
   
End Sub

Sub leerUsuario()
    campos(0, 0) = "usuario"
    campos(1, 0) = "clave"
    campos(2, 0) = "nombre"
    campos(3, 0) = ""
    campos(0, 2) = clientesistema + "auditoria.segu_usuarios"
    condicion = "usuario=" + "'" + dato1.text + "'"
    op = 5
    
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)
    If sqlconta.status = 4 Then dato1.text = "": dato1.SetFocus: GoTo no:
    USUARIOSISTEMA = sqlconta.response(0, 3)
    clavesistema = sqlconta.response(1, 3)
    
no:
End Sub
Sub enviar()
fechasistema = Date
MES = Format(fechasistema, "mm")
año = Format(fechasistema, "yyyy")
dia = Format(fechasistema, "dd")
sqlconta.usuarioauditoria = USUARIOSISTEMA
confi01.Show vbModal
End Sub

Public Sub COMPARAPROGRAMAS()
Dim origen As String
Dim destino As String
Dim File As String
Dim Temp As String
Dim Attrib As Integer
Dim FPA As String
Dim HPA As String
Dim FPO As String
Dim HPO As String
Dim Tamaño As String
On Error GoTo controlerror
File = RUTA + "contabilidadxp.EXE"
FPA = Mid(FileDateTime(File), 1, 10)
HPA = Mid(FileDateTime(File), 12, 10)
origen = File
rutadestino = App.path + "\"
File = rutadestino + "contabilidadxp.EXE"
FPO = Mid(FileDateTime(File), 1, 10)
HPO = Mid(FileDateTime(File), 12, 10)
destino = File
If FPA <> FPO Or HPA <> HPO Then
            actualizar
            
End If
Exit Sub
controlerror:
Rem MsgBox "EL SISTEMA NO ENCONTRO LA RUTA DE ACTUALIZACIONES", vbCritical, "ATENCION"
 End Sub
Public Sub actualizar()
     Call escribeArchivoRuta("SISTEMA", App.path & "\" & App.EXEName & ".exe", "C:\UPDATE.TXT")
     Call escribeArchivoRuta("UPDATE", RUTA & App.EXEName & ".exe", "C:\UPDATE.TXT")
     Call Shell(RUTA & "\Update.exe", vbNormalFocus)
    
   End Sub

Sub escribeArchivoRuta(ByVal tipo As String, ByVal cadena As String, ByVal ARCHIVO As String)
        Dim NUMFIC As Integer
        NUMFIC = FreeFile
        If tipo = "SISTEMA" Then
            Open ARCHIVO For Output As #NUMFIC
            Close #NUMFIC
        End If
        NUMFIC = FreeFile
        Open ARCHIVO For Append As #NUMFIC
        Print #NUMFIC, tipo & "=" & cadena
        Close #NUMFIC
    End Sub


Function estaactivo(ip, sistema) As Boolean
    Dim csql As New rdoQuery
    Dim resultados As rdoResultset
    
    Set csql.ActiveConnection = conta
    csql.sql = "select usuario,password from " & clientesistema & "auditoria.usuarios_activos where ip='" & ip & "' and sistema='" + sistema + "' limit 0,1 "
    csql.Execute
    estaactivo = False
    
    If csql.RowsAffected > 0 Then
         Set resultados = csql.OpenResultset
         USUARIOSISTEMA = resultados(0)
         dato1.text = USUARIOSISTEMA
         estaactivo = True
    End If
    csql.Close
    Set csql = Nothing
    
End Function

Sub leerdatosconeccion(NOMBRE)

    Call leerdatos_Certificado
    campos(0, 0) = "usuariomysql"
    campos(1, 0) = "passwordmysql"
    campos(2, 0) = "cliente"
    campos(3, 0) = "rutaactualizaciones"
    campos(4, 0) = "rutaarchivos"
    campos(5, 0) = ""
    
    campos(0, 2) = "admin_confi.clientes_admin "
    condicion = "sistema=" + "'" + NOMBRE + "'"
    op = 5
    
    sqlconta.response = campos
    Set sqlconta.conexion = conta
    Call sqlconta.sqlconta(op, condicion)
    If sqlconta.status = 0 Then
'        Usuario = sqlconta.response(0, 3)
'        password = sqlconta.response(1, 3)
        clientesistema = sqlconta.response(2, 3)
        RUTA = sqlconta.response(3, 3)
        RutaArchivos = sqlconta.response(4, 3)
    Else
    MsgBox ("NO EXISTE CONFIGURACION NI LICENCIA PARA ESTE SOFTWARE")
    Unload Me
    End If
    
no:
End Sub

