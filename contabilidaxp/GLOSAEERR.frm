VERSION 5.00
Begin VB.Form GLOSAEERR 
   BackColor       =   &H00FF8080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "OBSERVACION"
   ClientHeight    =   7635
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8805
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7635
   ScaleWidth      =   8805
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      BackColor       =   &H000000FF&
      Caption         =   "ELIMINAR GLOSA"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   4560
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   6960
      Width           =   2670
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "GRABAR GLOSA"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   840
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6960
      Width           =   2670
   End
   Begin VB.TextBox glosa 
      Height          =   5790
      Left            =   45
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   1065
      Width           =   8655
   End
   Begin VB.Label lblcentrocosto 
      Height          =   375
      Left            =   480
      TabIndex        =   5
      Top             =   7560
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label lblempresa 
      Height          =   375
      Left            =   -120
      TabIndex        =   4
      Top             =   7560
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label lblnombre 
      BackColor       =   &H00FF8080&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   600
      Width           =   8535
   End
   Begin VB.Label lblfecha 
      BackColor       =   &H00FF8080&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   8535
   End
End
Attribute VB_Name = "GLOSAEERR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
    If glosa.text <> "" Then
        Call grabar
        Unload Me
    Else
        MsgBox "DEBE AGREGAR UNA OBSERVACION ANTES DE GRABAR", vbCritical, "ATENCION"
        glosa.SetFocus
    End If
End Sub
Sub grabar()
    campos(0, 0) = "empresa"
    campos(1, 0) = "crcc"
    campos(2, 0) = "nombre"
    campos(3, 0) = "fecha"
    campos(4, 0) = "glosa"
    campos(5, 0) = ""
    
    campos(0, 1) = LBLEMPRESA.Caption
    campos(1, 1) = lblcentrocosto.Caption
    campos(2, 1) = lblNOMBRE.Caption
    campos(3, 1) = lblfecha.Caption
    campos(4, 1) = glosa.text
    
    
    campos(0, 2) = "analisis_eerr"
    
      condicion = "empresa='" & LBLEMPRESA.Caption & "' and crcc='" & lblcentrocosto.Caption & "' and nombre='" & lblNOMBRE.Caption & "' and fecha='" & lblfecha.Caption & "' "
    op = 5
    sqlconta.response = campos
    Set sqlconta.conexion = conta
    Call sqlconta.sqlconta(op, condicion)
    If sqlconta.status = 4 Then
        condicion = ""
        op = 2
        sqlconta.response = campos
        Set sqlconta.conexion = conta
        Call sqlconta.sqlconta(op, condicion)
    Else
        condicion = "empresa='" & LBLEMPRESA.Caption & "' and crcc='" & lblcentrocosto.Caption & "' and nombre='" & lblNOMBRE.Caption & "' and fecha='" & lblfecha.Caption & "' "

        op = 3
        sqlconta.response = campos
        Set sqlconta.conexion = conta
        Call sqlconta.sqlconta(op, condicion)
    End If
End Sub
Sub ELIMINAR()
    campos(0, 0) = ""
    campos(0, 2) = "analisis_eerr"
    condicion = "empresa='" & LBLEMPRESA.Caption & "' and crcc='" & lblcentrocosto.Caption & "' and nombre='" & lblNOMBRE.Caption & "' and fecha='" & lblfecha.Caption & "' "
    op = 4
    sqlconta.response = campos
    Set sqlconta.conexion = conta
    Call sqlconta.sqlconta(op, condicion)
End Sub
Private Sub COMMAND2_Click()
'    If Verifica_Permiso(infoge05.Caption, "elimina") = True Then
    If MsgBox("¿ESTA SEGURO QUE DESEA ELIMINAR GLOSA?", vbYesNo, "ATENCION") = vbYes Then
        frmglosaeliminacion.Show vbModal
        sqlconta.glosaeliminacion = glosaeliminacionsistema
        sqlconta.solicitoeliminacion = solicitaeliminacion
        Call ELIMINAR
        lblfecha.Caption = ""
        lblNOMBRE.Caption = ""
        glosa.text = ""
        LBLEMPRESA.Caption = ""
        lblcentrocosto.Caption = ""
        Unload Me
    End If
End Sub

Private Sub Form_Activate()
    leer
End Sub
 

Public Sub leer()
    
    campos(0, 0) = "empresa"
    campos(1, 0) = "crcc"
    campos(2, 0) = "nombre"
    campos(3, 0) = "fecha"
    campos(4, 0) = "glosa"
    campos(5, 0) = ""
    
    campos(0, 2) = "analisis_eerr"
    condicion = "empresa='" & LBLEMPRESA.Caption & "' and crcc='" & lblcentrocosto.Caption & "' and nombre='" & lblNOMBRE.Caption & "' and fecha='" & lblfecha.Caption & "' "
    op = 5
    sqlconta.response = campos
    Set sqlconta.conexion = conta
    Call sqlconta.sqlconta(op, condicion)
    If sqlconta.status = 0 Then
    glosa.text = sqlconta.response(4, 3)
    
    End If

End Sub
 

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        Unload Me
    End If
    
End Sub

Private Sub glosa_GotFocus()
    Call cargatexto(glosa)
End Sub

Private Sub glosa_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    
End Sub
