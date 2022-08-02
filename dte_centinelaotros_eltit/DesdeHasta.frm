VERSION 5.00
Object = "{ADBBAED6-B16F-40DC-80DF-B44910CBA76C}#1.0#0"; "Frame-Xp.ocx"
Begin VB.Form DesdeHasta 
   BackColor       =   &H00BC4A36&
   BorderStyle     =   0  'None
   Caption         =   "Agregar Pre-Venta"
   ClientHeight    =   1935
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6390
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1935
   ScaleWidth      =   6390
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin XPFrame.FrameXp FrameXp1 
      Height          =   1695
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   2990
      BackColor       =   16744576
      Caption         =   "Rango de Zetas"
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
      Begin VB.TextBox txtFinal 
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
         Height          =   345
         IMEMode         =   3  'DISABLE
         Left            =   3480
         MaxLength       =   10
         TabIndex        =   1
         Top             =   1140
         Width           =   2565
      End
      Begin VB.TextBox txtInicial 
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
         Height          =   345
         IMEMode         =   3  'DISABLE
         Left            =   3480
         MaxLength       =   10
         TabIndex        =   0
         Top             =   600
         Width           =   2565
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " N° Boleta Final"
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
         Width           =   3240
      End
      Begin VB.Label lbl1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " N° Boleta Inicial"
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
         Width           =   3240
      End
   End
End
Attribute VB_Name = "DesdeHasta"
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

Private Sub Form_Load()
    txtInicial.text = leer_Ultimo_Folio("boletahasta", "sv_documento_cabeza_" + empresaActiva, 10, ventasRubro, "local = '" & empresaActiva & "' AND tipo = 'ZE'")
End Sub

Private Sub txtInicial_GotFocus()
    Call selecciona(txtInicial)
End Sub

Private Sub txtFinal_GotFocus()
    Call selecciona(txtFinal)
End Sub


Private Sub txtInicial_KeyDown(KeyCode As Integer, Shift As Integer)
    Call Flechas(KeyCode, txtInicial)
End Sub

Private Sub txtFinal_KeyDown(KeyCode As Integer, Shift As Integer)
    Call Flechas(KeyCode, txtInicial)
End Sub


Private Sub txtInicial_KeyPress(KeyAscii As Integer)
    KeyAscii = esNumero(KeyAscii)
    If KeyAscii = 13 And txtInicial.text <> "" Then
        txtInicial.text = ceros(txtInicial)
        SendKeys "{Tab}"
    End If
End Sub

Private Sub txtFinal_KeyPress(KeyAscii As Integer)
    KeyAscii = esNumero(KeyAscii)
    If KeyAscii = 13 And txtFinal.text <> "" Then
        txtFinal.text = ceros(txtFinal)
       
        Unload Me
    End If
End Sub

