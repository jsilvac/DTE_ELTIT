VERSION 5.00
Object = "{ADBBAED6-B16F-40DC-80DF-B44910CBA76C}#1.0#0"; "Frame-Xp.ocx"
Begin VB.Form logincajera 
   Caption         =   "INGRESO DE CAJERA"
   ClientHeight    =   2175
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5445
   LinkTopic       =   "Form1"
   ScaleHeight     =   2175
   ScaleWidth      =   5445
   StartUpPosition =   3  'Windows Default
   Begin XPFrame.FrameXp FrameXp1 
      Height          =   2160
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   3810
      BackColor       =   16744576
      Caption         =   "Inicio de Sesión"
      CaptionEstilo3D =   1
      BackColor       =   16744576
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
      Begin VB.TextBox txtUserName 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   2880
         MaxLength       =   9
         TabIndex        =   0
         Top             =   585
         Width           =   1920
      End
      Begin VB.TextBox txtPassword 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         IMEMode         =   3  'DISABLE
         Left            =   2880
         MaxLength       =   10
         PasswordChar    =   "*"
         TabIndex        =   1
         Top             =   1080
         Width           =   2325
      End
      Begin VB.CommandButton cmdOK 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Aceptar"
         Height          =   390
         Left            =   1200
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   1680
         Width           =   1140
      End
      Begin VB.CommandButton cmdCancel 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Cancelar"
         Height          =   390
         Left            =   3120
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   1680
         Width           =   1140
      End
      Begin VB.Label dv 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4815
         TabIndex        =   8
         Top             =   585
         Width           =   375
      End
      Begin VB.Label lblDV 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   4860
         TabIndex        =   7
         Top             =   600
         Visible         =   0   'False
         Width           =   420
      End
      Begin VB.Label lblLabels 
         BackStyle       =   0  'Transparent
         Caption         =   "Rut Cajera"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   0
         Left            =   120
         TabIndex        =   6
         Top             =   600
         Width           =   2640
      End
      Begin VB.Label lblLabels 
         BackStyle       =   0  'Transparent
         Caption         =   "Contraseña"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   1
         Left            =   120
         TabIndex        =   5
         Top             =   1080
         Width           =   1920
      End
   End
End
Attribute VB_Name = "LOGINCAJERA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
AUTORIZASISTEMA = False
Unload Me

End Sub

Private Sub cmdOK_Click()
    
    If leerUsuario(Me.txtUserName.text, Me.txtPassword.text) = True Then
        cajero = txtUserName.text
        cajero1 = txtUserName.text
        AUTORIZASISTEMA = True
        
        Unload Me
        
    Else
        MsgBox ("CODIGO DE CAJERA NO EXISTE ")
        Me.txtUserName.SetFocus
    End If
End Sub

Private Sub txtUserName_GotFocus()
    Call selecciona(txtUserName)
End Sub

Private Sub txtUserName_KeyPress(KeyAscii As Integer)
    KeyAscii = esNumero(KeyAscii)
    If KeyAscii = 13 Then
        txtUserName.text = ceros(txtUserName)
        rutCajera = txtUserName.text
        dv.Caption = rut(txtUserName.text)
        txtPassword.SetFocus
        
        
    End If
End Sub

Private Function leerUsuario(ByVal usuario As String, ByVal pass As String) As Boolean
    
    Dim op As Integer
    Dim CAMPOS(10, 10) As String
    
    Dim cad As String
    Dim p As String
    CAMPOS(0, 0) = "nombre"
    CAMPOS(1, 0) = ""
    CAMPOS(0, 2) = "sv_maestrocajeras"
    condicion = "rut = '" & txtUserName.text & dv.Caption & "' AND password = '" & txtPassword.text & "'"
    op = 5
    sqlventas.response = CAMPOS
    Set sqlventas.conexion = ventas
    Call sqlventas.sqlventas(op, condicion)
    If sqlventas.Status = 0 Then
        leerUsuario = True
        nombrecajero = sqlventas.response(0, 3)
        codigocajero = txtUserName.text
        
    Else
        leerUsuario = False
    End If
End Function

Private Sub txtPassword_GotFocus()
    Call selecciona(txtPassword)
End Sub

Private Sub txtPassword_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        If txtPassword.text = "SALIR" Then
            End
        Else
        Call cmdOK_Click
        End If
    End If
End Sub

