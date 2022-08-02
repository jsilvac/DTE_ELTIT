VERSION 5.00
Object = "{ADBBAED6-B16F-40DC-80DF-B44910CBA76C}#1.0#0"; "Frame-Xp.ocx"
Begin VB.Form MUsuarios 
   BackColor       =   &H00008000&
   BorderStyle     =   0  'None
   Caption         =   "Configurar Puesto de Trabajo"
   ClientHeight    =   4110
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7575
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4110
   ScaleWidth      =   7575
   ShowInTaskbar   =   0   'False
   Begin XPFrame.FrameXp FrameXp1 
      Height          =   3915
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   6906
      BackColor       =   12648384
      Caption         =   " Maestro de Usuarios"
      CaptionEstilo3D =   1
      BackColor       =   12648384
      ColorBarraArriba=   12648384
      ColorBarraAbajo =   32768
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.TextBox dato6 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   2160
         MaxLength       =   20
         PasswordChar    =   "*"
         TabIndex        =   16
         Top             =   1440
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.TextBox dato2 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   2160
         MaxLength       =   20
         PasswordChar    =   "*"
         TabIndex        =   1
         Top             =   1020
         Width           =   2415
      End
      Begin VB.TextBox dato5 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   2160
         MaxLength       =   50
         TabIndex        =   4
         Top             =   2760
         Width           =   4995
      End
      Begin VB.TextBox dato1 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   2160
         MaxLength       =   18
         TabIndex        =   0
         Top             =   600
         Width           =   2415
      End
      Begin VB.TextBox dato3 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   2160
         MaxLength       =   50
         TabIndex        =   2
         Top             =   1920
         Width           =   4995
      End
      Begin VB.TextBox dato4 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   2160
         MaxLength       =   50
         TabIndex        =   3
         Top             =   2340
         Width           =   4995
      End
      Begin VB.CheckBox Check3 
         BackColor       =   &H00C0FFC0&
         Caption         =   " Ver Password"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   315
         Left            =   5040
         TabIndex        =   5
         Top             =   1020
         Visible         =   0   'False
         Width           =   2055
      End
      Begin XPFrame.FrameXp frmCerrar 
         Height          =   330
         Left            =   6960
         TabIndex        =   7
         Top             =   30
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   582
         BackColor       =   49344
         Caption         =   "X"
         CaptionEstilo3D =   1
         BackColor       =   49344
         ColorBarraArriba=   32896
         ColorBarraAbajo =   12648447
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
      End
      Begin XPFrame.FrameXp frmEliminar 
         Height          =   375
         Left            =   5580
         TabIndex        =   13
         Top             =   3240
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         BackColor       =   49344
         Caption         =   "Eliminar"
         CaptionEstilo3D =   1
         BackColor       =   49344
         ColorBarraArriba=   12648384
         ColorBarraAbajo =   32768
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
      End
      Begin XPFrame.FrameXp frmModificar 
         Height          =   375
         Left            =   3900
         TabIndex        =   14
         Top             =   3240
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         BackColor       =   49344
         Caption         =   "Modificar"
         CaptionEstilo3D =   1
         BackColor       =   49344
         ColorBarraArriba=   12648384
         ColorBarraAbajo =   32768
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
      End
      Begin XPFrame.FrameXp frmRetorno 
         Height          =   375
         Left            =   2160
         TabIndex        =   15
         Top             =   3240
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         BackColor       =   49344
         Caption         =   "Retorno"
         CaptionEstilo3D =   1
         BackColor       =   49344
         ColorBarraArriba=   12648384
         ColorBarraAbajo =   32768
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
      End
      Begin VB.Label lbl6 
         Appearance      =   0  'Flat
         BackColor       =   &H00008080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Clave Nueva"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   180
         TabIndex        =   17
         Top             =   1440
         Visible         =   0   'False
         Width           =   1875
      End
      Begin VB.Label lbl5 
         Appearance      =   0  'Flat
         BackColor       =   &H00008080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Email"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   180
         TabIndex        =   12
         Top             =   2760
         Width           =   1875
      End
      Begin VB.Label lbl4 
         Appearance      =   0  'Flat
         BackColor       =   &H00008080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Labor"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   180
         TabIndex        =   11
         Top             =   2340
         Width           =   1875
      End
      Begin VB.Label lbl1 
         Appearance      =   0  'Flat
         BackColor       =   &H00008080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Nombre Usuario"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   180
         TabIndex        =   10
         Top             =   600
         Width           =   1875
      End
      Begin VB.Label lbl2 
         Appearance      =   0  'Flat
         BackColor       =   &H00008080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Clave"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   180
         TabIndex        =   9
         Top             =   1020
         Width           =   1875
      End
      Begin VB.Label lbl3 
         Appearance      =   0  'Flat
         BackColor       =   &H00008080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Nombre"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   180
         TabIndex        =   8
         Top             =   1920
         Width           =   1875
      End
   End
End
Attribute VB_Name = "MUsuarios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Private campos(10, 3) As String
    Private modificar As Boolean

Private Sub Check3_Click()
    If Check3.Value = 0 Then
        dato2.PasswordChar = "*"
    Else
        dato2.PasswordChar = ""
    End If
End Sub

'============================================================
'MANEJO DE LOS CONTOLES
'============================================================
    '========================================================
    'GotFocus
    '========================================================
    Private Sub dato1_GotFocus()
        Call cargatexto(dato1)
       
    End Sub
    
    Private Sub dato2_GotFocus()
        Call cargatexto(dato2)
    End Sub
    
    Private Sub dato3_GotFocus()
         Call cargatexto(dato3)
    End Sub
    
    Private Sub dato4_GotFocus()
         Call cargatexto(dato4)
    End Sub
    
    Private Sub dato5_GotFocus()
        Call cargatexto(dato5)
    End Sub
    '========================================================
    'GotFocus
    '========================================================
    
    '========================================================
    'KeyDown
    '========================================================
    Private Sub dato1_KeyDown(KeyCode As Integer, Shift As Integer)
          Call Flechas(KeyCode, dato1)
    End Sub
    
    Private Sub dato2_KeyDown(KeyCode As Integer, Shift As Integer)
       Call Flechas(KeyCode, dato2)
    End Sub
    
    Private Sub dato3_KeyDown(KeyCode As Integer, Shift As Integer)
         Call Flechas(KeyCode, dato3)
    End Sub
    
    Private Sub dato4_KeyDown(KeyCode As Integer, Shift As Integer)
        Call Flechas(KeyCode, dato4)
    End Sub
    
    Private Sub dato5_KeyDown(KeyCode As Integer, Shift As Integer)
       Call Flechas(KeyCode, dato5)
    End Sub
    '========================================================
    'KeyDown
    '========================================================
    
    '========================================================
    'KeyPress
    '========================================================
    Private Sub dato1_KeyPress(KeyAscii As Integer)
        If KeyAscii = 27 Then Unload Me
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If KeyAscii = 13 And dato1.text <> "" Then
            If leerUsuario("=") = False Then
                SendKeys "{Tab}"
            Else
                Call DeshabilitarCajas(Me)
            End If
        End If
    End Sub
    
    Private Sub dato2_KeyPress(KeyAscii As Integer)
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If KeyAscii = 13 And dato2.text <> "" And modificar = False Then
            dato3.SetFocus
        End If
         If KeyAscii = 13 And dato2.text <> "" And modificar = True Then
            dato6.SetFocus
        End If
    End Sub
    
    Private Sub dato3_KeyPress(KeyAscii As Integer)
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If KeyAscii = 13 And dato3.text <> "" Then
            SendKeys "{Tab}"
        End If
    End Sub
    
    Private Sub dato4_KeyPress(KeyAscii As Integer)
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If KeyAscii = 13 And dato4.text <> "" Then
            SendKeys "{Tab}"
        End If
    End Sub
    
    Private Sub dato5_KeyPress(KeyAscii As Integer)
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If KeyAscii = 13 And dato5.text <> "" And dato2.text <> "" Then
        
         If dato6.text <> "" And modificar = True Then
            
            If leerpassUsuario(dato1.text, dato2.text) = True Then
               Call grabarUsuario(modificar)
               Call retorno
            Else
               MsgBox ("Contraseña Actual Erronea")
               dato2.text = ""
               dato2.SetFocus
           End If
           
         Else
        If modificar = False Then
         Call grabarUsuario(modificar)
         Call retorno
         Else
         MsgBox ("Clave Nueva Esta en Blanco")
         dato6.SetFocus
         End If
         
        
        End If
       End If
    End Sub
    '========================================================
    'KeyPress
    '========================================================
'============================================================
'MANEJO DE LOS CONTOLES
'============================================================



Private Sub dato6_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 And dato6.text <> "" Then

If leerpassUsuario(dato1.text, dato2.text) = True Then
dato3.SetFocus
Else
MsgBox ("Contraseña Actual Erronea")
dato2.text = ""
dato2.SetFocus
End If
End If

End Sub

    Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
        If KeyCode = 27 Then
            Unload Me
        End If
        If KeyCode = 38 Then
            If Screen.ActiveForm.ActiveControl.Name = "dato1" Then
                Unload Me
            End If
        End If
    End Sub

    Private Sub Form_Load()
        Call Centrar(Me)
        modificar = False
    End Sub



    Private Sub frmCerrar_BarClick()
        Call cambiaColor(frmCerrar)
        frmCerrar.CaptionEstilo3D = Inserted
        Unload Me
    End Sub
    
    Private Sub frmCerrar_BarMouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
        Call cambiaColor(frmCerrar)
        frmCerrar.CaptionEstilo3D = Raised
    End Sub

    Private Sub frmModificar_BarClick()
        Call cambiaColor(frmModificar)
        frmModificar.CaptionEstilo3D = Inserted
        modificar = True
        Call HabilitarCajas(Me, modificar)
        lbl6.Visible = True
        dato6.Visible = True
        
        dato2.SetFocus
    End Sub
    
    Private Sub frmModificar_BarMouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
        Call cambiaColor(frmModificar)
        frmModificar.CaptionEstilo3D = Raised
    End Sub
    
    Private Sub frmEliminar_BarClick()
        Call cambiaColor(frmEliminar)
        frmEliminar.CaptionEstilo3D = Inserted
        If MsgBox("DESEA REALMENTE ELIMINAR Si / No", vbYesNo) = vbYes Then
        
        Call eliminarUsuario
        Call retorno
        
        End If
        
    End Sub
    
    Private Sub frmEliminar_BarMouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
        Call cambiaColor(frmEliminar)
        frmEliminar.CaptionEstilo3D = Raised
    End Sub

'=============================================================================
'LEER USUARIO
'=============================================================================
    Private Function leerUsuario(ByVal operador As String) As Boolean
        Dim condicion As String
        Dim op As Integer
     
        campos(0, 0) = "usuario"
        campos(1, 0) = "clave"
        campos(2, 0) = "nombre"
        campos(3, 0) = "labor"
        campos(4, 0) = "email"
        campos(5, 0) = ""
        
        campos(0, 2) = "segu_usuarios"
        
        condicion = "usuario " & operador & " '" & dato1.text & "'"
        If operador = "<" Then
            condicion = condicion & "ORDER BY usuario DESC"
        Else
            condicion = condicion & "ORDER BY usuario ASC"
        End If
        op = 5
        SQLUTIL.datos = campos
        Set SQLUTIL.conexion = ventas
        Call SQLUTIL.SQLUTIL(op, condicion)
        If SQLUTIL.estado = 0 Then
            leerUsuario = True
            dato1.text = SQLUTIL.datos(0, 3)
            dato2.text = "***"
            dato3.text = SQLUTIL.datos(2, 3)
            dato4.text = SQLUTIL.datos(3, 3)
            dato5.text = SQLUTIL.datos(4, 3)
        Else
            leerUsuario = False
        End If
    End Function
    
     Private Function leerpassUsuario(ByVal usuario As String, ByVal CLAVE As String) As Boolean
        Dim condicion As String
        Dim op As Integer
     
        campos(0, 0) = "clave"
        campos(1, 0) = ""
        campos(0, 2) = "segu_usuarios"
        
        condicion = "usuario ='" & usuario & "'"
        op = 5
        SQLUTIL.datos = campos
        Set SQLUTIL.conexion = ventas
        Call SQLUTIL.SQLUTIL(op, condicion)
        If SQLUTIL.estado = 0 Then
        If CLAVE = SQLUTIL.datos(0, 3) Then
            leerpassUsuario = True
        Else
            leerpassUsuario = False
        End If
        End If
        
        
    End Function
 
'=============================================================================
'LEER USUARIO
'=============================================================================

'=============================================================================
'GRABAR USUARIO
'=============================================================================
    Private Sub grabarUsuario(ByVal modifica As Boolean)
        Dim cSql As rdoQuery
        
        Set cSql = New rdoQuery
        Set cSql.ActiveConnection = ventas
        
        cSql.sql = "INSERT INTO segu_usuarios (usuario, clave, nombre, labor, email) "
        cSql.sql = cSql.sql & "VALUES('" & dato1.text & "','" & dato2.text & "', '" & dato3.text & "', '" & dato4.text & "', '" & dato5.text & "') "
        
        If modifica = True Then
            cSql.sql = "UPDATE segu_usuarios SET clave = '" & dato6.text & "', nombre = '" & dato3.text & "', labor = '" & dato4.text & "', email = '" & dato5.text & "' "
            cSql.sql = cSql.sql & "WHERE usuario = '" & dato1.text & "' "
        End If
        cSql.Execute
        cSql.Close
        Set cSql = Nothing
    End Sub
'=============================================================================
'GRABAR USUARIO
'=============================================================================

'=============================================================================
'ELIMINAR USUARIO
'=============================================================================
    Private Sub eliminarUsuario()
        Dim condicion As String
        Dim op As Integer
        
        condicion = "usuario = '" & dato1.text & "'"
        op = 4
        campos(0, 2) = "segu_usuarios"
        SQLUTIL.datos = campos
        Set SQLUTIL.conexion = ventas
        Call SQLUTIL.SQLUTIL(op, condicion)
    End Sub
'=============================================================================
'ELIMINAR USUARIO
'=============================================================================

    Private Sub retorno()
        modificar = False
        Call LimpiarCajas(Me)
        Call HabilitarCajas(Me, modificar)
        lbl6.Visible = False
        dato6.Visible = False
        
        dato1.SetFocus
    End Sub

Private Sub frmRetorno_BarClick()
 Call cambiaColor(frmRetorno)
 frmRetorno.CaptionEstilo3D = Inserted
 retorno
End Sub

Private Sub frmRetorno_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call cambiaColor(frmRetorno)
   frmRetorno.CaptionEstilo3D = Raised
End Sub


