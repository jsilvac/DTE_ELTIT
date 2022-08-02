VERSION 5.00
Begin VB.Form cambioLocal 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cambio de Local"
   ClientHeight    =   1635
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6045
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1635
   ScaleWidth      =   6045
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Cambiar 
      Caption         =   "&Cambiar"
      Height          =   315
      Left            =   1680
      TabIndex        =   1
      Top             =   1200
      Width           =   2355
   End
   Begin VB.TextBox dato1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
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
      Left            =   1380
      MaxLength       =   2
      TabIndex        =   0
      Tag             =   "proveedor"
      Top             =   600
      Width           =   795
   End
   Begin VB.Label lblNuevo 
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
      Left            =   2220
      TabIndex        =   5
      Top             =   600
      Width           =   3675
   End
   Begin VB.Label lblActual 
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
      Left            =   1380
      TabIndex        =   4
      Top             =   120
      Width           =   4515
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00F5C9B1&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Nuevo Local"
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   60
      TabIndex        =   3
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label lbl1 
      Appearance      =   0  'Flat
      BackColor       =   &H00F5C9B1&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Local Actual"
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   60
      TabIndex        =   2
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "cambioLocal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Cambiar_Click()
    If dato1.text <> "" And lblNuevo.Caption <> "" Then
        empresaactiva = dato1.text
        rubro = leerRubro(empresaactiva)
        Call ConectarRubro(servidor, basedatos, usuario, password)
        Unload Me
    End If
End Sub

Private Sub dato1_GotFocus()
    Call selecciona(dato1)
End Sub

Private Sub dato1_KeyPress(KeyAscii As Integer)
    KeyAscii = esNumero(KeyAscii)
    If KeyAscii = 13 Then
        dato1.text = ceros(dato1)
        lblNuevo.Caption = leerNombreEmpresa(dato1.text)
        If lblNuevo.Caption <> "" Then
            SendKeys "{Tab}"
        End If
    End If
End Sub

Private Sub Form_Load()
    lblActual = leerNombreEmpresa(empresaactiva)
End Sub
