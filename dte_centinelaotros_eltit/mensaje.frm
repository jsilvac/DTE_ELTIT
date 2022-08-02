VERSION 5.00
Object = "{ADBBAED6-B16F-40DC-80DF-B44910CBA76C}#1.0#0"; "Frame-Xp.ocx"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form mensaje 
   BackColor       =   &H000B11FB&
   BorderStyle     =   0  'None
   Caption         =   "Mensaje"
   ClientHeight    =   6045
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10860
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   6045
   ScaleWidth      =   10860
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin XPFrame.FrameXp FrameXp1 
      Height          =   5775
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   10575
      _ExtentX        =   18653
      _ExtentY        =   10186
      BackColor       =   192
      Caption         =   ""
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
      Begin MSWinsockLib.Winsock ws 
         Left            =   120
         Top             =   4620
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin VB.Timer tmrBlink 
         Interval        =   500
         Left            =   120
         Top             =   5220
      End
      Begin VB.TextBox dato1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   10200
         MaxLength       =   6
         TabIndex        =   4
         Top             =   4500
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00000080&
         Caption         =   " Presione 'E' para solicitar autorización, de lo contrario presione 'Esc'."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   1155
         Left            =   660
         TabIndex        =   3
         Top             =   4500
         Width           =   9255
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblMensaje 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   1935
         Left            =   240
         TabIndex        =   2
         Top             =   480
         Width           =   10095
      End
      Begin VB.Label lblCodigo 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   36
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   1695
         Left            =   240
         TabIndex        =   1
         Top             =   2640
         Width           =   10095
      End
   End
End
Attribute VB_Name = "mensaje"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Private color(1) As Variant

Private Sub dato1_GotFocus()
    Call selecciona(dato1)
End Sub

'Private Sub dato1_KeyPress(KeyAscii As Integer)
'    KeyAscii = esNumero(KeyAscii)
'    If KeyAscii = 13 Then
'        dato1.text = ceros(dato1)
'        'VERIFICAR CODIGO
'        If dato1.text = "000123" Then
'            Unload Me
'            autorizado = True
'        Else
'            Call selecciona(dato1)
'        End If
'    End If
'End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then
        Unload Me
        If detallePagos.Visible = True Then
            Unload detallePagos
        End If
        Unload PVentas
    End If
    If KeyCode = Asc("e") Or KeyCode = Asc("E") Then
        Call enviarDatos
    End If
End Sub

Private Sub Form_Load()
    Beep
    color(0) = &HFFFF&  'amarillo
    color(1) = &HC0&    'rojo
    If PVentas.Visible = True Then
        'If envia = True Then
            'ENVIAR POR RED
            Call conectarWs
        'End If
    End If
End Sub

Private Sub tmrBlink_Timer()
    Static estado As Integer
    estado = 1 - estado
    lblMensaje.ForeColor = color(estado)
    If PVentas.Visible = True Then
        'If verificarCupoCliente(rut_cliente, sucursal_cliente) = True Then
            Unload Me
        'End If
    End If
End Sub

Public Sub mostrarMensaje(ByVal titulo As String, ByVal msj1 As String, ByVal msj2 As String)
    FrameXp1.Caption = titulo
    lblMensaje.Caption = msj1
    lblCodigo.Caption = msj2
    mensaje.Show vbModal
End Sub

Private Sub conectarWs()
    ws.LocalPort = 1000
    ws.RemoteHost = "127.0.0.1"
    ws.RemotePort = 1001
    ws.Connect
End Sub

Private Sub enviarDatos()
    ws.SendData empresaActiva & "/" & rut_cliente & "/" & sucursal_cliente & "/" & tipo_doc & "/" & numero_doc
End Sub

Private Sub ws_ConnectionRequest(ByVal requestID As Long)
    ws.Close
    ws.Accept requestID
End Sub
