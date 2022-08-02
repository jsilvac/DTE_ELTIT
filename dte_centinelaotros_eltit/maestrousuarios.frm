VERSION 5.00
Object = "{ADBBAED6-B16F-40DC-80DF-B44910CBA76C}#1.0#0"; "Frame-Xp.ocx"
Begin VB.Form maestro15 
   BackColor       =   &H00008000&
   BorderStyle     =   0  'None
   Caption         =   "Configurar Puesto de Trabajo"
   ClientHeight    =   5115
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7575
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5115
   ScaleWidth      =   7575
   ShowInTaskbar   =   0   'False
   Begin XPFrame.FrameXp FrameXp1 
      Height          =   4770
      Left            =   135
      TabIndex        =   2
      Top             =   120
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   8414
      BackColor       =   12648384
      Caption         =   " Cambiar Claves de Seguridad"
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
      Begin VB.TextBox dato4 
         Appearance      =   0  'Flat
         BackColor       =   &H80000007&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   2160
         MaxLength       =   8
         PasswordChar    =   "*"
         TabIndex        =   15
         Top             =   1980
         Width           =   2415
      End
      Begin VB.TextBox dato3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000007&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   2160
         MaxLength       =   8
         PasswordChar    =   "*"
         TabIndex        =   10
         Top             =   1530
         Width           =   2415
      End
      Begin VB.TextBox dato2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000007&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   2160
         MaxLength       =   8
         PasswordChar    =   "*"
         TabIndex        =   1
         Top             =   1035
         Width           =   2415
      End
      Begin VB.TextBox dato1 
         Appearance      =   0  'Flat
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
         Height          =   300
         Left            =   2160
         MaxLength       =   18
         TabIndex        =   0
         Top             =   630
         Width           =   2415
      End
      Begin XPFrame.FrameXp frmCerrar 
         Height          =   330
         Left            =   6960
         TabIndex        =   3
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
      Begin XPFrame.FrameXp frmRetorno 
         Height          =   375
         Left            =   2655
         TabIndex        =   9
         Top             =   4320
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
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Max (8 digitos Letras o Numeros)"
         Height          =   330
         Left            =   4680
         TabIndex        =   19
         Top             =   1980
         Width           =   2445
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Max (8 digitos Letras o Numeros)"
         Height          =   330
         Left            =   4680
         TabIndex        =   18
         Top             =   1530
         Width           =   2445
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Max (8 digitos Letras o Numeros)"
         Height          =   330
         Left            =   4680
         TabIndex        =   17
         Top             =   1080
         Width           =   2445
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00008080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Redigite Clave"
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
         TabIndex        =   16
         Top             =   1980
         Width           =   1875
      End
      Begin VB.Label lblmail 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   315
         Left            =   2205
         TabIndex        =   14
         Top             =   3825
         Width           =   4935
      End
      Begin VB.Label lbllabor 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   315
         Left            =   2205
         TabIndex        =   13
         Top             =   3420
         Width           =   4935
      End
      Begin VB.Label lblnombre 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   315
         Left            =   2205
         TabIndex        =   12
         Top             =   3015
         Width           =   4935
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
         TabIndex        =   11
         Top             =   1530
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
         Left            =   225
         TabIndex        =   8
         Top             =   3840
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
         Left            =   225
         TabIndex        =   7
         Top             =   3420
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
         TabIndex        =   6
         Top             =   600
         Width           =   1875
      End
      Begin VB.Label lbl2 
         Appearance      =   0  'Flat
         BackColor       =   &H00008080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Clave Actual"
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
         TabIndex        =   5
         Top             =   1065
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
         Left            =   225
         TabIndex        =   4
         Top             =   3000
         Width           =   1875
      End
   End
End
Attribute VB_Name = "maestro15"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Private CAMPOS(10, 3) As String
    Private modificar As Boolean
    Private CLAVE As String
    

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
      Call Flechas(KeyCode, dato1)
    End Sub
    
    Private Sub dato3_KeyDown(KeyCode As Integer, Shift As Integer)
        Call Flechas(KeyCode, dato2)
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
              
            End If
        End If
    End Sub
    
    Private Sub dato2_KeyPress(KeyAscii As Integer)
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If KeyAscii = 13 Then
               If dato2.text = CLAVE Then
                dato3.Enabled = True
                
                dato3.SetFocus
                Else
                 MsgBox ("CLAVE INGRESADA NO CORRESPONDE A LA ORIGINAL")
                dato2.SetFocus
                End If
        End If
       
    End Sub
    
    Private Sub dato3_KeyPress(KeyAscii As Integer)
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If KeyAscii = 13 And dato3.text <> "" Then
            dato4.Enabled = True
            
            SendKeys "{Tab}"
        End If
    End Sub
    
Private Sub dato4_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If dato3.text <> dato4.text Then
    MsgBox ("CLAVE NUEVA NO CONCIDE CON REDIGITACION ")
    dato3.text = ""
    dato4.text = ""
    dato3.SetFocus
    Else
    If dato2.text = CLAVE Then
        Call modificaclave(dato1.text, dato4.text)
        MsgBox ("CLAVE HA SIDO MODIFICADA CON EXITO CERRAR SESION ")
        Unload Me
    Else
        MsgBox ("CLAVE NO COINCIDE CON LA ORIGINAL ")
        dato2.SetFocus
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
        dato1.text = usuarioSistema
        Call leerUsuario("=")
   
   End Sub



    Private Sub frmCerrar_BarClick()
        Call cambiaColor(frmCerrar)
        frmCerrar.CaptionEstilo3D = Inserted
        Unload Me
    End Sub
    
    Private Sub frmCerrar_BarMouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
        Call cambiaColor(frmCerrar)
        frmCerrar.CaptionEstilo3D = Raised
    End Sub

'=============================================================================
'LEER USUARIO
'=============================================================================
    Private Function leerUsuario(ByVal operador As String) As Boolean
        
        Dim op As Integer
     
        CAMPOS(0, 0) = "usuario"
        CAMPOS(1, 0) = "clave"
        CAMPOS(2, 0) = "nombre"
        CAMPOS(3, 0) = "labor"
        CAMPOS(4, 0) = "email"
        CAMPOS(5, 0) = ""
        
        CAMPOS(0, 2) = clientesistema & "auditoria.segu_usuarios"
        
        condicion = "usuario " & operador & " '" & dato1.text & "'"
        If operador = "<" Then
            condicion = condicion & "ORDER BY usuario DESC"
        Else
            condicion = condicion & "ORDER BY usuario ASC"
        End If
        op = 5
        sqlventas.response = CAMPOS
        Set sqlventas.conexion = ventas
        Call sqlventas.sqlventas(op, condicion)
        If sqlventas.Status = 0 Then
            leerUsuario = True
            dato1.text = sqlventas.response(0, 3)
            CLAVE = sqlventas.response(1, 3)
            lblnombre.Caption = sqlventas.response(2, 3)
            lbllabor.Caption = sqlventas.response(3, 3)
            lblmail.Caption = sqlventas.response(4, 3)
        Else
            leerUsuario = False
        End If
    End Function
    
     Sub modificaclave(ByVal usuario As String, ByVal CLAVENUEVA As String)
        
        Dim op As Integer
     
        CAMPOS(0, 0) = "clave"
        CAMPOS(1, 0) = ""
        CAMPOS(0, 1) = CLAVENUEVA
        
        CAMPOS(0, 2) = clientesistema & "auditoria.segu_usuarios"
        
        condicion = "usuario ='" & usuario & "'"
        op = 3
        sqlventas.response = CAMPOS
        Set sqlventas.conexion = ventas
        Call sqlventas.sqlventas(op, condicion)
        
        
    End Sub
 
'=============================================================================
'LEER USUARIO
'=============================================================================

'=============================================================================
'GRABAR USUARIO
'=============================================================================
   
'=============================================================================
'GRABAR USUARIO
'=============================================================================

'=============================================================================
'ELIMINAR USUARIO
'=============================================================================
    
'=============================================================================
'ELIMINAR USUARIO
'=============================================================================

    Private Sub retorno()
        Unload Me
        
    End Sub

Private Sub frmRetorno_BarClick()
 Call cambiaColor(frmRetorno)
 frmRetorno.CaptionEstilo3D = Inserted
 retorno
End Sub

Private Sub frmRetorno_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
   Call cambiaColor(frmRetorno)
   frmRetorno.CaptionEstilo3D = Raised
End Sub


