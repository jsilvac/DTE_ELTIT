VERSION 5.00
Object = "{ADBBAED6-B16F-40DC-80DF-B44910CBA76C}#1.0#0"; "Frame-Xp.ocx"
Begin VB.Form cambioFecha 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cambiar Fecha Sistema"
   ClientHeight    =   2070
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3555
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2070
   ScaleWidth      =   3555
   StartUpPosition =   2  'CenterScreen
   Begin XPFrame.FrameXp FrameXp1 
      Height          =   1395
      Left            =   60
      TabIndex        =   2
      Top             =   60
      Width           =   3435
      _ExtentX        =   6059
      _ExtentY        =   2461
      BackColor       =   16744576
      Caption         =   "Fecha Sistema"
      CaptionEstilo3D =   1
      BackColor       =   16744576
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
      Begin VB.TextBox dato3 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Left            =   2400
         MaxLength       =   4
         TabIndex        =   9
         Tag             =   "proveedor"
         Top             =   900
         Width           =   915
      End
      Begin VB.TextBox dato2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Left            =   1920
         MaxLength       =   2
         TabIndex        =   8
         Tag             =   "proveedor"
         Top             =   900
         Width           =   435
      End
      Begin VB.TextBox dato1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Left            =   1440
         MaxLength       =   2
         TabIndex        =   0
         Tag             =   "proveedor"
         Top             =   900
         Width           =   435
      End
      Begin VB.Label lblAño 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
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
         ForeColor       =   &H0080FFFF&
         Height          =   285
         Left            =   2400
         TabIndex        =   7
         Top             =   420
         Width           =   915
      End
      Begin VB.Label lblMes 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
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
         ForeColor       =   &H0080FFFF&
         Height          =   285
         Left            =   1920
         TabIndex        =   6
         Top             =   420
         Width           =   435
      End
      Begin VB.Label lbl1 
         Appearance      =   0  'Flat
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Fecha Actual"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   120
         TabIndex        =   5
         Top             =   420
         Width           =   1215
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Nueva Fecha"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   120
         TabIndex        =   4
         Top             =   900
         Width           =   1215
      End
      Begin VB.Label lblDia 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
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
         ForeColor       =   &H0080FFFF&
         Height          =   285
         Left            =   1440
         TabIndex        =   3
         Top             =   420
         Width           =   435
      End
   End
   Begin XPFrame.FrameXp frmAceptar 
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   1560
      Width           =   3315
      _ExtentX        =   5847
      _ExtentY        =   661
      BackColor       =   49344
      Caption         =   "A   C   E   P   T   A   R"
      CaptionEstilo3D =   1
      BackColor       =   49344
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   1
   End
End
Attribute VB_Name = "cambioFecha"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Private contador As Integer

Private Sub Cambiar()
    If dato1.text <> "" And dato2.text <> "" And dato3.text <> "" Then
        fechasistema = dato3.text & "-" & dato2.text & "-" & dato1.text
        Unload Me
    End If
End Sub

Private Sub dato1_GotFocus()
    Call selecciona(dato1)
End Sub

Private Sub dato2_GotFocus()
    Call selecciona(dato2)
End Sub

Private Sub dato3_GotFocus()
    Call selecciona(dato3)
End Sub

Private Sub dato1_KeyDown(KeyCode As Integer, Shift As Integer)
    Call VerificarCajas(Me, dato1)
    Call Flechas(KeyCode, dato1)
End Sub

Private Sub dato2_KeyDown(KeyCode As Integer, Shift As Integer)
    Call VerificarCajas(Me, dato2)
    Call Flechas(KeyCode, dato1)
End Sub

Private Sub dato3_KeyDown(KeyCode As Integer, Shift As Integer)
    Call VerificarCajas(Me, dato3)
    Call Flechas(KeyCode, dato2)
End Sub

Private Sub dato1_KeyPress(KeyAscii As Integer)
    KeyAscii = esNumero(KeyAscii)
    If KeyAscii = 13 Then
        dato1.text = ceros(dato1)
        If dato1.text = "00" Then
            dato1.text = Format(fechasistema, "dd")
        End If
        SendKeys "{Tab}"
    End If
End Sub

Private Sub dato2_KeyPress(KeyAscii As Integer)
    KeyAscii = esNumero(KeyAscii)
    If KeyAscii = 13 Then
        dato2.text = ceros(dato2)
        If dato2.text = "00" Then
            dato2.text = Format(fechasistema, "mm")
        End If
        SendKeys "{Tab}"
    End If
End Sub

Private Sub dato3_KeyPress(KeyAscii As Integer)
    KeyAscii = esNumero(KeyAscii)
    If KeyAscii = 13 Then
        If contador = 0 Then
            dato3.text = ceros(dato3)
            If dato3.text = "0000" Then
                dato3.text = Format(fechasistema, "yyyy")
                contador = 1
            End If
        Else
            Call Cambiar
        End If
    End If
End Sub

'Private Sub dato1_KeyUp(KeyCode As Integer, Shift As Integer)
'    If Len(dato1.text) = dato1.MaxLength Then
'        Call dato1_KeyPress(13)
'    End If
'End Sub
'
'Private Sub dato2_KeyUp(KeyCode As Integer, Shift As Integer)
'    If Len(dato2.text) = dato2.MaxLength Then
'        Call dato2_KeyPress(13)
'    End If
'End Sub
'
'Private Sub dato3_KeyUp(KeyCode As Integer, Shift As Integer)
'    If Len(dato3.text) = dato3.MaxLength Then
'        Call dato3_KeyPress(13)
'    End If
'End Sub

Private Sub Form_Activate()
    Principal.barraEstado.Panels(1).text = UCase(Me.Caption)
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
    contador = 0
    lblDia.Caption = Format(fechasistema, "dd")
    lblMes.Caption = Format(fechasistema, "mm")
    lblAño.Caption = Format(fechasistema, "yyyy")
End Sub

    Private Sub frmAceptar_BarMouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
        Call cambiaColor(frmAceptar)
        frmAceptar.CaptionEstilo3D = Raised
    End Sub
    
    Private Sub frmAceptar_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
        Call cambiaColor(frmAceptar)
        frmAceptar.CaptionEstilo3D = Inserted
        Call Cambiar
    End Sub
