VERSION 5.00
Begin VB.Form GLOSARECEPCION 
   BackColor       =   &H00FF8080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "glosa"
   ClientHeight    =   7635
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8805
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7635
   ScaleWidth      =   8805
   StartUpPosition =   3  'Windows Default
   Begin VB.OptionButton Option2 
      Caption         =   "Faltante en Factura"
      Height          =   285
      Left            =   6345
      TabIndex        =   4
      Top             =   135
      Width           =   1770
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Diferencia Precios"
      Height          =   285
      Left            =   4320
      TabIndex        =   3
      Top             =   135
      Width           =   1770
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
      Left            =   3105
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   7020
      Width           =   2670
   End
   Begin VB.TextBox glosa 
      Height          =   6270
      Left            =   45
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   585
      Width           =   8655
   End
   Begin VB.Label noc 
      BackColor       =   &H00FF8080&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   180
      TabIndex        =   2
      Top             =   90
      Width           =   3615
   End
End
Attribute VB_Name = "GLOSARECEPCION"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()

    CAMPOS(0, 0) = "glosa"
    CAMPOS(1, 0) = "tipoglosa"
    CAMPOS(2, 0) = ""
    
    CAMPOS(0, 1) = glosa.text
    If Option1.Value = True Then
    CAMPOS(1, 1) = "1"
    Else
    CAMPOS(1, 1) = "2"
    End If
    
    
    CAMPOS(0, 2) = "l_movimientos_cabeza_" + localorden
    
    condicion = "tipo='OC' and numero='" + noc.Caption + "' "
    op = 3
    sqlconta.response = CAMPOS
    Set sqlconta.conexion = gestionrubro
    Call sqlconta.sqlconta(op, condicion)
   Unload Me
   
End Sub

Private Sub Form_Load()

noc.Caption = prove0002.GRID1.Cell(prove0002.GRID1.ActiveCell.row, 2).text

Option1.Value = True
leer
End Sub

Sub leer()
 
    
    CAMPOS(0, 0) = "glosa"
    CAMPOS(1, 0) = "tipoglosa"
    CAMPOS(2, 0) = ""
    CAMPOS(0, 2) = "l_movimientos_cabeza_" + localorden
    
    condicion = "tipo='OC' and numero='" + noc.Caption + "' "
    op = 5
    sqlconta.response = CAMPOS
    Set sqlconta.conexion = gestionrubro
    Call sqlconta.sqlconta(op, condicion)
   If sqlconta.status = 0 Then
   glosa.text = sqlconta.response(0, 1)
If sqlconta.response(1, 1) = "1" Then

    Option1.Value = "1"
    Else
    Option2.Value = "1"
    End If
   End If

End Sub

