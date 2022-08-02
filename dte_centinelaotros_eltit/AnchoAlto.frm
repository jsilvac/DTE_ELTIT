VERSION 5.00
Object = "{ADBBAED6-B16F-40DC-80DF-B44910CBA76C}#1.0#0"; "Frame-Xp.ocx"
Begin VB.Form AnchoAlto 
   BackColor       =   &H00BC4A36&
   BorderStyle     =   0  'None
   Caption         =   "Agregar Pre-Venta"
   ClientHeight    =   1935
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3690
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1935
   ScaleWidth      =   3690
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin XPFrame.FrameXp FrameXp1 
      Height          =   1695
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   3435
      _ExtentX        =   6059
      _ExtentY        =   2990
      BackColor       =   16744576
      Caption         =   "Tamaño del Papel"
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
      Begin VB.TextBox txtAlto 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         IMEMode         =   3  'DISABLE
         Left            =   2280
         MaxLength       =   10
         TabIndex        =   1
         Top             =   1140
         Width           =   1005
      End
      Begin VB.TextBox txtAncho 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         IMEMode         =   3  'DISABLE
         Left            =   2280
         MaxLength       =   10
         TabIndex        =   0
         Top             =   600
         Width           =   1005
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Alto (cm.)"
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
         Height          =   390
         Left            =   120
         TabIndex        =   4
         Top             =   1140
         Width           =   2040
      End
      Begin VB.Label lbl1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Ancho (cm.)"
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
         Height          =   390
         Left            =   120
         TabIndex        =   3
         Top             =   600
         Width           =   2040
      End
   End
End
Attribute VB_Name = "AnchoAlto"
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

Private Sub txtAncho_GotFocus()
    Call VerificarCajas(Me, txtAncho)
    Call selecciona(txtAncho)
End Sub

Private Sub txtAlto_GotFocus()
    Call VerificarCajas(Me, txtAlto)
    Call selecciona(txtAlto)
End Sub

Private Sub txtAncho_KeyDown(KeyCode As Integer, Shift As Integer)
    Call Flechas(KeyCode, txtAncho)
End Sub

Private Sub txtAlto_KeyDown(KeyCode As Integer, Shift As Integer)
    Call Flechas(KeyCode, txtAlto)
End Sub

Private Sub txtAncho_KeyPress(KeyAscii As Integer)
    KeyAscii = esNumeroDecimal(txtAncho, KeyAscii)
    If KeyAscii = 13 And txtAncho.text <> "" Then
        SendKeys "{Tab}"
    End If
End Sub

Private Sub txtAlto_KeyPress(KeyAscii As Integer)
    KeyAscii = esNumeroDecimal(txtAlto, KeyAscii)
    If KeyAscii = 13 And txtAlto.text <> "" Then
        MCajas.ancho = txtAncho.text
        MCajas.alto = txtAlto.text
        Unload Me
    End If
End Sub

