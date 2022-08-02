VERSION 5.00
Object = "{ADBBAED6-B16F-40DC-80DF-B44910CBA76C}#1.0#0"; "Frame-Xp.ocx"
Begin VB.Form digitacrcc 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Digita Centro de Costos"
   ClientHeight    =   1410
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5175
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1410
   ScaleWidth      =   5175
   StartUpPosition =   2  'CenterScreen
   Begin XPFrame.FrameXp CRCC 
      Height          =   1455
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   2566
      BackColor       =   16744576
      Caption         =   ""
      CaptionEstilo3D =   1
      BackColor       =   16744576
      ForeColor       =   8438015
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.TextBox pivote2 
         Enabled         =   0   'False
         Height          =   285
         Left            =   360
         TabIndex        =   6
         Text            =   "Text1"
         Top             =   1200
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox DATO21 
         BackColor       =   &H00E1FFFD&
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
         Left            =   1320
         MaxLength       =   2
         TabIndex        =   0
         Tag             =   "codigo"
         Top             =   240
         Width           =   375
      End
      Begin VB.TextBox DATO22 
         BackColor       =   &H00E1FFFD&
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
         Left            =   1755
         MaxLength       =   2
         TabIndex        =   1
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton crccok 
         Caption         =   "OK"
         Height          =   375
         Left            =   2400
         TabIndex        =   3
         Top             =   960
         Width           =   1815
      End
      Begin VB.Label nombrecrcc 
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
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
         Left            =   120
         TabIndex        =   5
         Top             =   600
         Width           =   4815
      End
      Begin VB.Label Label7 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Codigo Cuenta"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   1095
      End
   End
End
Attribute VB_Name = "digitacrcc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub crccok_Click()
If nombrecrcc.Caption <> "" Then
DIGITA_CRCC_CODIGO = DATO21.text + DATO22.text
DIGITA_CRCC_NOMBRE = nombrecrcc.Caption
Unload Me


Else
DATO21.SetFocus
End If

End Sub

Private Sub DATO21_GotFocus()
    Call cargatexto(DATO21)
End Sub

Private Sub dato22_GotFocus()
    Call cargatexto(DATO22)
End Sub

Private Sub dato21_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF2 Then Call ayudacrcc
    Call flechas(DATO21, DATO22, KeyCode)
no:
End Sub

Private Sub dato22_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Call flechas(DATO21, DATO22, KeyCode)
End Sub

Private Sub dato21_KeyPress(KeyAscii As Integer)
    snum = 0: KeyAscii = esNumero(KeyAscii)
    If KeyAscii = 13 Then
    Call ceros(DATO21)
    DATO22.SetFocus
    
    End If
    
End Sub

Private Sub dato22_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 13 Then
   Call ceros(DATO22)
  If leerNOMBREcrcc(DATO21.text + DATO22.text) <> "" Then
  nombrecrcc.Caption = leerNOMBREcrcc(DATO21.text + DATO22.text)
  crccok.SetFocus
  Else
  DATO21.text = ""
  DATO22.text = ""
  DATO21.SetFocus
  
  End If
  
   
   End If
   
End Sub


Sub ayudacrcc()
    Dim campos As Variant
    Dim cfijo As Variant
    Dim largo As Variant
    campos = Array("codigo", "nombre")
    cabezas = Array("codigo", "nombre")
    largo = Array("8n", "40s")
    mensajeAyuda = "Ayuda Centros de costo"
    cfijo = "año='" + Format(fechasistema, "yyyy") + "'"
    pivote2.MaxLength = 4
    Call cargaAyudaT(Servidor, basebus, Usuario, password, "centrosdecosto", pivote2, campos, cfijo, largo, 2)
    DATO21.text = Mid(pivote2.text, 1, 2)
    DATO22.text = Mid(pivote2.text, 3, 2)
    
    pivote2.text = ""
End Sub


    

Sub leercrcc(row As Long, col As Long)
    campos(0, 0) = "codigo"
    campos(1, 0) = "nombre"
    campos(2, 0) = ""
    campos(0, 2) = "centrosdecosto"
    condicion = "codigo=" + "'" + DATO21.text + DATO22.text + "' and año='" + Format(fechasistema, "yyyy") + "'"
    op = 5
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)
    Rem If sqlconta.status = 4 Or DATO22.text = "00" Then DATO21.Enabled = True: DATO21.text = "": DATO22.text = "": DATO21.SetFocus: GoTo no:
    Grid1.Cell(row, 12).text = sqlconta.response(1, 3)
    Grid1.Cell(row, 14).text = sqlconta.response(0, 3)
    If col <> 9999 Then
    nombrecrcc.Caption = sqlconta.response(0, 3)
   
  End If

no:
End Sub

