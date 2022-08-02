VERSION 5.00
Object = "{ADBBAED6-B16F-40DC-80DF-B44910CBA76C}#1.0#0"; "Frame-Xp.ocx"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form permiso 
   BackColor       =   &H0080FF80&
   BorderStyle     =   0  'None
   Caption         =   "Agregar Pre-Venta"
   ClientHeight    =   1485
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6390
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1485
   ScaleWidth      =   6390
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSCommLib.MSComm scanner 
      Left            =   0
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin VB.TextBox PIVOTE1 
      Height          =   195
      Left            =   0
      MaxLength       =   10
      TabIndex        =   3
      Top             =   0
      Visible         =   0   'False
      Width           =   150
   End
   Begin XPFrame.FrameXp FrameXp1 
      Height          =   1215
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   2143
      BackColor       =   65280
      Caption         =   "Permiso Supervisor"
      CaptionEstilo3D =   1
      BackColor       =   65280
      ForeColor       =   65535
      ColorBarraArriba=   0
      ColorBarraAbajo =   255
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
      ColorTextShadow =   0
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         IMEMode         =   3  'DISABLE
         Left            =   3480
         MaxLength       =   13
         PasswordChar    =   "*"
         TabIndex        =   0
         Top             =   600
         Width           =   2565
      End
      Begin VB.Label lblLabels 
         BackStyle       =   0  'Transparent
         Caption         =   "Código de Supervisor"
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
         Left            =   120
         TabIndex        =   2
         Top             =   600
         Width           =   3240
      End
   End
End
Attribute VB_Name = "permiso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then
        Unload Me
    End If
End Sub

Private Sub txtCodigo_GotFocus()
    Call selecciona(txtCodigo)
End Sub

Private Sub txtCodigo_KeyPress(KeyAscii As Integer)
    KeyAscii = esNumero(KeyAscii)
    autorizador = False
    tercera_Edad = False
    
    
    If KeyAscii = 13 Then
        'verificar permisos
        PIVOTE1.text = txtCodigo.text
        PIVOTE1.MaxLength = 10
        PIVOTE1.text = ceros(PIVOTE1)
        If leeautorizacionterceraedad(PIVOTE1.text) = False Then
        tercera_Edad = False
        If leeautorizacion(Me.txtCodigo.text) = True Then
           autorizador = True
           claveautorizadora = Me.txtCodigo.text
           
            Unload Me
'            Select Case code
'                Case vbKeyF4
'                    eliminaCodigo.Show vbModal
'                Case vbKeyF5
'                    Call borrarRollo
'                Case vbKeyF9
'                    PuntoVenta.dato3.Locked = False
'                    PuntoVenta.dato3.SetFocus'
'            End Select
        Else
              MsgBox "EL CODIGO DE AUTORIZACION ES INCORRECTO CODIGO: " & Me.txtCodigo.text & "", vbCritical, "ERROR "
            autorizador = False
            Unload Me
            
        End If
        Else
          autorizador = True
          tercera_Edad = True
          Unload Me
        End If
    End If
    
End Sub

 
