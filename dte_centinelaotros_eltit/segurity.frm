VERSION 5.00
Object = "{ADBBAED6-B16F-40DC-80DF-B44910CBA76C}#1.0#0"; "Frame-Xp.ocx"
Begin VB.Form seguridad 
   Caption         =   "SISTEMA DE VENTAS"
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
         Locked          =   -1  'True
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
         Locked          =   -1  'True
         PasswordChar    =   "*"
         TabIndex        =   2
         Top             =   1200
         Width           =   1575
      End
      Begin VB.Image Image2 
         Height          =   1290
         Left            =   240
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
      Stretch         =   -1  'True
      Top             =   0
      Width           =   15360
   End
End
Attribute VB_Name = "seguridad"
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
    
    Dim ss As String
    Dim K As Integer
    
    Close 20
        Open App.Path + "\confiventas.txt" For Input As #20
    While EOF(20) = False
    Input #20, ss
    
    If Mid(ss, 1, 8) = "SERVIDOR" Then
        Servidor = Mid(ss, 10, Len(ss) - 9)
    End If
    If Mid(ss, 1, 9) = "BASEDATOS" Then
        basedatos = Mid(ss, 11, Len(ss) - 10)
    End If
    If Mid(ss, 1, 10) = "BASEVENTAS" Then
        baseVentas = Mid(ss, 12, Len(ss) - 11)
    End If
    If Mid(ss, 1, 7) = "EMPRESA" Then
        empresaActiva = Mid(ss, 9, Len(ss) - 8)
    End If
    If Mid(ss, 1, 6) = "BODEGA" Then
        bodega = Mid(ss, 8, Len(ss) - 7)
    End If
    If Mid(ss, 1, 4) = "CAJA" Then
        idCaja = Mid(ss, 6, Len(ss) - 5)
    End If
    If Mid(ss, 1, 4) = "RUTA" Then
        rutaUpdate = Mid(ss, 6, Len(ss) - 5)
    End If
    If Mid(ss, 1, 13) = "IMPRESORAPAGO" Then
        IMPRESORAPAGO = Mid(ss, 15, Len(ss) - 5)
    End If
    If Mid(ss, 1, 12) = "BODEGARETIRO" Then
        BODEGARETIRO = Mid(ss, 14, Len(ss) - 5)
    End If
    If Mid(ss, 1, 16) = "IMPRESORACREDITO" Then
        impresoracredito = Mid(ss, 18, Len(ss) - 5)
    End If
    If Mid(ss, 1, 11) = "ASEGURADORA" Then
        ASEGURADORA = Mid(ss, 13, Len(ss) - 12)
    End If
    If Mid(ss, 1, 14) = "IMPRIMEDIRECTO" Then
    If Mid(ss, 16, Len(ss) - 14) = "S" Then
        imprimeDirecto = True
        Else
        imprimeDirecto = False
    End If
    End If
    If Mid(ss, 1, 11) = "IMPRIMETIPO" Then
        imprIMETIPO = Mid(ss, 13, Len(ss) - 11)
    End If
    
    
    Wend
        Close 20
 
 
        usuario = "admixp"
        password = "1"
 

    Call Conectartemporal(Servidor, "mysql", usuario, password)

    Call leerdatosconeccion("facturaelectronica.exe")
    basedatos = clientesistema + "gestion"
            
        baseteso = clientesistema & "teso"
        baseauditoria = clientesistema
        segundosespera = "60"


        Call Conectar(Servidor, basedatos, usuario, password)
        
        rubro = leerRubro(empresaActiva)
        Call ConectarRubro(Servidor, basedatos, usuario, password)
        Call Conectarventas(Servidor, baseVentas & empresaActiva, usuario, password)
        iva = leerImpuesto("IVA")
        iha = leerImpuesto("IHA")
        fechasistema = Format(Now, "yyyy-mm-dd")
        empresa
        
    
    envia = False
    mensaje_nopermiso = "Usted no tiene privilegios suficientes para realizar esta operación."
    cabezas = Array("codigo", "nombre")
    mensajeAyuda = "Ayuda General"
                Call Conectar_Auditoria
            Set sqlventas.conauditoria = conexionauditoria
'    Call COMPARAPROGRAMAS

    dwlen = MAX_COMPUTERNAME_LENGTH + 1
    strstring = String(dwlen, "X")
    GetComputerName strstring, dwlen
    strstring = Left(strstring, dwlen)
    nombremaquina = strstring
    IPFIJA = GetWanIP
    electro04.Show
    
End Sub



Private Sub dato1_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 13 Then CLAVE.SetFocus
   
End Sub
Private Sub clave_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then dato1.SetFocus
   
End Sub

Sub leerUsuario()
    campos(0, 0) = "usuario"
    campos(1, 0) = "clave"
    campos(2, 0) = "nombre"
    campos(3, 0) = ""
    campos(0, 2) = clientesistema & "auditoria.segu_usuarios"
    condicion = "usuario=" + "'" + dato1.text + "'"
    op = 5
    
    sqlgesti.response = campos
    Set sqlgesti.conexion = conta
    Call sqlgesti.sqlgesti(op, condicion)
    If sqlgesti.Status = 4 Then dato1.text = "": dato1.SetFocus: GoTo no:
    usuarioSistema = sqlgestion.response(0, 3)
    clavesistema = sqlgestion.response(1, 3)
    
no:
End Sub
Sub enviar()
fechasistema = Date
mes = Format(fechasistema, "mm")
año = Format(fechasistema, "yyyy")
dia = Format(fechasistema, "dd")
sqlventas.usuarioauditoria = usuarioSistema
 Call revisarmenus(Principal)
 Principal.Show
 Unload Me
End Sub

 

Sub escribeArchivoRuta(ByVal tipo As String, ByVal cadena As String, ByVal ARCHIVO As String)
        Dim numfic As Integer
        numfic = FreeFile
        If tipo = "SISTEMA" Then
            Open ARCHIVO For Output As #numfic
            Close #numfic
        End If
        numfic = FreeFile
        Open ARCHIVO For Append As #numfic
        Print #numfic, tipo & "=" & cadena
        Close #numfic
    End Sub


Function estaactivo(IP, sistema) As Boolean
    Dim csql As New rdoQuery
    Dim resultados As rdoResultset
    
    Set csql.ActiveConnection = temporal
    csql.sql = "select usuario,password from " & clientesistema & "auditoria.usuarios_activos where ip='" & IP & "' and sistema='" + sistema + "' limit 0,1 "
    csql.Execute
    estaactivo = False
    
    If csql.RowsAffected > 0 Then
         Set resultados = csql.OpenResultset
         usuarioSistema = resultados(0)
         dato1.text = usuarioSistema
         estaactivo = True
    End If
    csql.Close
    Set csql = Nothing
    
End Function

Sub leerdatosconeccion(nombre)
     Dim campos(20, 3) As String
     Dim op As Integer
     
     Dim sql As New sqlventas.sqlventa
    campos(0, 0) = "usuariomysql"
    campos(1, 0) = "passwordmysql"
    campos(2, 0) = "cliente"
    campos(3, 0) = "rutaactualizaciones"
    campos(0, 2) = "admin_confi.clientes_admin "
    condicion = "sistema=" + "'" + nombre + "'"
    op = 5
    
    sql.response = campos
    Set sql.conexion = temporal
    Call sql.sqlventas(op, condicion)
    If sql.Status = 0 Then
    usuario = sql.response(0, 3)
    password = sql.response(1, 3)
    clientesistema = sql.response(2, 3)
    rutaUpdate = sql.response(3, 3)
    Else
    MsgBox ("NO EXISTE CONFIGURACION NI LICENCIA PARA ESTE SOFTWARE")
    Unload Me
    End If
    
no:
End Sub

