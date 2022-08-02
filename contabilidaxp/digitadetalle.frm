VERSION 5.00
Object = "{ADBBAED6-B16F-40DC-80DF-B44910CBA76C}#1.0#0"; "Frame-Xp.ocx"
Begin VB.Form digitarut 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Digita Cuenta Corriente"
   ClientHeight    =   1965
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7215
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1965
   ScaleWidth      =   7215
   StartUpPosition =   2  'CenterScreen
   Begin XPFrame.FrameXp frmrut 
      Height          =   1935
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   3413
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
      Begin VB.TextBox pivote 
         Height          =   285
         Left            =   600
         MaxLength       =   10
         TabIndex        =   8
         Text            =   "Text1"
         Top             =   1680
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox DATO20 
         BackColor       =   &H00C0FFFF&
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
         Left            =   1395
         MaxLength       =   9
         TabIndex        =   0
         Tag             =   "rut"
         Top             =   840
         Width           =   1095
      End
      Begin VB.CommandButton RUTOK 
         Caption         =   "OK"
         Height          =   315
         Left            =   2520
         TabIndex        =   2
         Top             =   1560
         Width           =   1695
      End
      Begin VB.Label lblglosa 
         BackColor       =   &H00C0FFC0&
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
         Left            =   1440
         TabIndex        =   7
         Top             =   360
         Width           =   5535
      End
      Begin VB.Label lblcuenta 
         BackColor       =   &H00C0FFC0&
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
         ForeColor       =   &H00400000&
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label nombrectacte 
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
         Top             =   1200
         Width           =   6855
      End
      Begin VB.Label Label9 
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Rut"
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
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label dv 
         BackColor       =   &H00C0FFFF&
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
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   2520
         TabIndex        =   3
         Top             =   840
         Width           =   255
      End
   End
End
Attribute VB_Name = "digitarut"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub dato20_KeyPress(KeyAscii As Integer)
    snum = 0: KeyAscii = esNumero(KeyAscii)
    If KeyAscii = 13 Then
    Call ceros(DATO20)
    DV.Caption = rut(DATO20.text)
    If leerNombrerut(lblcuenta.Caption, DATO20.text + DV.Caption) <> "" Then
    nombrectacte.Caption = leerNombrerut(lblcuenta.Caption, DATO20.text + DV.Caption)
    RUTOK.SetFocus
    
    End If
    
    
    
    End If
End Sub
Private Sub dato20_KeyDown(KeyCode As Integer, Shift As Integer)
    
    
    
    If KeyCode = vbKeyF2 Then Call ayudactacte(DATO20)
    End Sub


Private Sub Label1_Click()

End Sub

Private Sub Label2_Click()

End Sub

Sub ayudactacte(ByRef caja As TextBox)
    Dim CAMPOS As Variant
    Dim cfijo As Variant
    Dim largo As Variant
    CAMPOS = Array("rut", "nombre")
    largo = Array("12n", "40s")
    cfijo = "tipo='" & lblcuenta.Caption & "' and año='" + Format(fechasistema, "yyyy") + "'"
    cabezas = Array("rut", "nombre")
    mensajeAyuda = "Ayuda Cuentas Corrientes"
    pivote.MaxLength = 10
    Call cargaAyudaT(servidor, basebus, usuario, password, "cuentascorrientes", pivote, CAMPOS, cfijo, largo, 2)
    If Val(pivote.text) = 0 Then caja.SetFocus: GoTo no
    DATO20.text = Mid(pivote.text, 1, 9)
    DV.Caption = Mid(pivote.text, 10, 1)
    caja.Enabled = True
    caja.SetFocus

no:

End Sub




Private Sub Form_Unload(Cancel As Integer)
If nombrectacte.Caption = "" Then
Cancel = 1

End If
End Sub

Private Sub RUTOK_Click()
If nombrectacte.Caption <> "" Then
DIGITA_RUT_RUT = DATO20.text + DV.Caption
DIGITA_RUT_NOMBRE = nombrectacte.Caption
Unload Me


Else
DATO20.SetFocus
End If
End Sub
