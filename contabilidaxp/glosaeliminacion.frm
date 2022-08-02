VERSION 5.00
Object = "{ADBBAED6-B16F-40DC-80DF-B44910CBA76C}#1.0#0"; "Frame-Xp.ocx"
Begin VB.Form frmglosaeliminacion 
   Caption         =   "GLOSA ELIMINACION"
   ClientHeight    =   3480
   ClientLeft      =   60
   ClientTop       =   465
   ClientWidth     =   7575
   LinkTopic       =   "Form1"
   ScaleHeight     =   3480
   ScaleWidth      =   7575
   StartUpPosition =   3  'Windows Default
   Begin XPFrame.FrameXp Detalleflete 
      Height          =   3495
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   7530
      _ExtentX        =   13282
      _ExtentY        =   6165
      BackColor       =   255
      Caption         =   "Motivo Eliminacion"
      CaptionEstilo3D =   1
      BackColor       =   255
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
      Alignment       =   1
      Begin VB.TextBox SOLICITADO 
         Height          =   285
         Left            =   2070
         MaxLength       =   50
         TabIndex        =   0
         Top             =   360
         Width           =   5280
      End
      Begin VB.CommandButton Finalizar 
         BackColor       =   &H00FF8080&
         Caption         =   "Continuar"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2745
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   3120
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.TextBox glosa2 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1860
         Left            =   90
         MaxLength       =   200
         MultiLine       =   -1  'True
         TabIndex        =   2
         Top             =   1155
         Width           =   7380
      End
      Begin VB.Label lbl3 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   0
         TabIndex        =   6
         Top             =   3120
         Width           =   2655
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H0080C0FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "RAZON DE ELIMINACION SEA CLARO EN SU DETALLE INFORMACION GERENCIA"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   90
         TabIndex        =   5
         Top             =   900
         Width           =   7350
      End
      Begin VB.Label Label1 
         BackColor       =   &H0080C0FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "SOLICITADO POR"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   135
         TabIndex        =   4
         Top             =   360
         Width           =   1860
      End
   End
End
Attribute VB_Name = "frmglosaeliminacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Finalizar_Click()

Unload Me

End Sub


Private Sub Form_Unload(Cancel As Integer)
If existeusuario(SOLICITADO.text) = False Then
SOLICITADO.text = ""
End If
If glosa2.text = "" Or SOLICITADO.text = "" Then
Cancel = 1
SOLICITADO.SetFocus
End If

End Sub




Private Sub glosa2_KeyPress(KeyAscii As Integer)
  KeyAscii = Asc(UCase(Chr(KeyAscii)))
If KeyAscii = 13 And glosa2.text <> "" Then
glosaeliminacionsistema = glosa2.text
Finalizar_Click

End If


End Sub


Private Sub SOLICITADO_GotFocus()
lbl3.Caption = "F2 - Ayuda Usuarios"
End Sub

Private Sub SOLICITADO_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF2 Then
Call ayudausuarios(SOLICITADO)
End If
End Sub

Private Sub SOLICITADO_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
If KeyAscii = 13 And SOLICITADO.text <> "" Then
If existeusuario(SOLICITADO.text) = True Then
solicitaeliminacion = SOLICITADO.text
glosa2.SetFocus
End If
End If


End Sub
 Sub ayudausuarios(ByRef caja As TextBox)
    Dim campos As Variant
    Dim cfijo As Variant
    Dim largo As Variant
    campos = Array("usuario", "nombre")
    cabezas = Array("usuario", "nombre")
    largo = Array("30s", "30s")
    mensajeAyuda = "Ayuda usuarios"
    cfijo = "no"
    
    Call cargaAyudaT(Servidor, clientesistema + "auditoria", Usuario, password, "segu_usuarios", caja, campos, cfijo, largo, 2)
    If caja.text = "" Then caja.SetFocus: GoTo no
    caja.Enabled = True
    caja.SetFocus


no:

End Sub

 
Private Sub SOLICITADO_LostFocus()
lbl3.Caption = ""
End Sub
Function existeusuario(Usuario) As Boolean
Dim csql As New rdoQuery
Dim resultado As rdoResultset
Dim tabla As String
Set csql.ActiveConnection = contadb
tabla = "select nombre from " + clientesistema + "auditoria.segu_usuarios where usuario='" & Usuario & "' "
csql.sql = tabla
csql.Execute
existeusuario = False
If csql.RowsAffected > 0 Then
existeusuario = True
Else
MsgBox "USUARIO NO EXISTE PORFAVOR VERIFIQUE INFORMACION", vbCritical, "ATENCION"
existeusuario = False
End If

End Function

