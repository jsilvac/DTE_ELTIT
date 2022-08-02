VERSION 5.00
Object = "{ADBBAED6-B16F-40DC-80DF-B44910CBA76C}#1.0#0"; "Frame-Xp.ocx"
Begin VB.Form Seguridad4 
   BackColor       =   &H00000080&
   BorderStyle     =   0  'None
   Caption         =   "Sistema de Seguridad Activado"
   ClientHeight    =   2700
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6795
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   2700
   ScaleWidth      =   6795
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin XPFrame.FrameXp FrameXp1 
      Height          =   2535
      Left            =   60
      TabIndex        =   2
      Top             =   60
      Width           =   6675
      _ExtentX        =   11774
      _ExtentY        =   4471
      BackColor       =   192
      Caption         =   "SISTEMA DE SEGURIDAD ACTIVADO"
      CaptionEstilo3D =   1
      BackColor       =   192
      ColorBarraArriba=   12632319
      ColorBarraAbajo =   128
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
      Begin VB.TextBox txtPassword 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   3780
         PasswordChar    =   "*"
         TabIndex        =   1
         Top             =   1560
         Width           =   2715
      End
      Begin VB.TextBox txtUsuario 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3780
         TabIndex        =   0
         Top             =   720
         Width           =   2715
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "CONTRASEÑA"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   240
         TabIndex        =   4
         Top             =   1560
         Width           =   3375
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "NOMBRE DE USUARIO"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   240
         TabIndex        =   3
         Top             =   720
         Width           =   3375
      End
   End
End
Attribute VB_Name = "Seguridad4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    
Private Sub Form_Activate()
    Principal.barraEstado.Panels(1).text = UCase(Me.Caption)
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then
        If titCaption = "PRINCIPAL" Then
            End
        Else
            Call UnloadHijo(Principal, titCaption)
            Unload Me
        End If
    End If
    If KeyCode = 38 Then
        If Screen.ActiveForm.ActiveControl.Name = "txtUsuario" Then
            If titCaption = "PRINCIPAL" Then
                End
            Else
                Unload Me
                Call UnloadHijo(Principal, titCaption)
            End If
        End If
    End If
End Sub
Private Sub Form_Load()
    Call leerDatosConectar
    COMPARAPROGRAMAS
    EMPRESA
End Sub

Private Sub txtPassword_GotFocus()
    Call selecciona(txtPassword)
End Sub

Private Sub txtPassword_KeyDown(KeyCode As Integer, Shift As Integer)
    Call Flechas(KeyCode, txtUsuario)
End Sub

Private Sub txtPassword_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 And txtUsuario.text <> "" Then
        If Verificar(txtUsuario.text, txtPassword.text) = True Then
            If Principal.barraEstado.Panels(3).text = "" Then
                usuarioSistema = txtUsuario.text
                passwordSistema = txtPassword.text
            End If
            Principal.barraEstado.Panels(3).text = txtUsuario.text
            Unload Me
            Principal.PASO = True
        Else
            txtUsuario.text = ""
            txtPassword.text = ""
            txtUsuario.SetFocus
        End If
    End If
End Sub

Private Sub txtUsuario_GotFocus()
    Call selecciona(txtUsuario)
End Sub

Private Sub txtUsuario_KeyDown(KeyCode As Integer, Shift As Integer)
    Call Flechas(KeyCode, txtUsuario)
End Sub

Private Sub txtUsuario_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 And txtUsuario.text <> "" Then
    txtUsuario.text = Replace(txtUsuario.text, "'", "")
    
        SendKeys "{Tab}"
    End If
End Sub
Sub COMPARAPROGRAMAS()
Dim ORIGEN As String
Dim destino As String
Dim File As String
Dim Temp As String
Dim Attrib As Integer
Dim FPA As String
Dim HPA As String
Dim FPO As String
Dim HPO As String
Dim rutadestino As String

 On Error GoTo controlerror
    rutadestino = App.Path + "\"

    File = rutaUpdate + "\sistemadeventas.exe"
    FPA = Mid(FileDateTime(File), 1, 10)
    HPA = Mid(FileDateTime(File), 12, 10)
    ORIGEN = File
    File = rutadestino + "sistemadeventas.exe"
    FPO = Mid(FileDateTime(File), 1, 10)
    HPO = Mid(FileDateTime(File), 12, 10)
    destino = File
    If FPA <> FPO Or HPA <> HPO Then
        actualizar
        End
    End If
    mensaje_nopermiso = "Usted no tiene privilegios suficientes para realizar esta operación."
    mensaje_noelimina = "Imposible eliminar el producto, presenta movimientos de inventario asociados"
    Exit Sub
controlerror:
MsgBox "EL SISTEMA NO ENCONTRO LA RUTA DE ACTUALIZACIONES", vbCritical, "ATENCION"
End Sub
Public Sub actualizar()
     Call escribeArchivoRuta("SISTEMA", App.Path & "\" & App.EXEName & ".exe", "C:\UPDATE.TXT")
    Call escribeArchivoRuta("UPDATE", rutaUpdate & "\" & App.EXEName & ".exe", "C:\UPDATE.TXT")
    
    
        Call Shell(rutaUpdate & "\Update.exe", vbNormalFocus)
    
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

