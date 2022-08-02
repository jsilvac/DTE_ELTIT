VERSION 5.00
Object = "{085FBDF3-00C5-421B-B762-1D57299A2B89}#1.0#0"; "CLBUTN.OCX"
Begin VB.Form preguntar 
   Caption         =   "Modulo Pregunta"
   ClientHeight    =   2655
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5040
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   2655
   ScaleWidth      =   5040
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame opcionelimina 
      BackColor       =   &H00FF8080&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   4335
      Begin CoolButtons.cool_Button si 
         Height          =   375
         Left            =   240
         TabIndex        =   1
         Top             =   1680
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         Caption         =   "Si"
         ForeColor       =   16711680
      End
      Begin CoolButtons.cool_Button no 
         Height          =   375
         Left            =   2280
         TabIndex        =   2
         Top             =   1680
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         Caption         =   "No "
         ForeColor       =   16711680
      End
      Begin VB.Label MENSAJE 
         Alignment       =   2  'Center
         BackColor       =   &H8000000A&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   1935
         Left            =   120
         TabIndex        =   3
         Top             =   120
         Width           =   4095
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00F5C9B1&
         BorderWidth     =   2
         Height          =   2175
         Left            =   0
         Top             =   0
         Width           =   4335
      End
   End
End
Attribute VB_Name = "preguntar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_KeyPress(KeyAscii As Integer)
If UCase(Chr(KeyAscii)) = "S" Then RESPUESTA = "S": Unload Me
If UCase(Chr(KeyAscii)) = "N" Then RESPUESTA = "N": Unload Me
End Sub

Private Sub Form_Load()
RESPUESTA = "N"
End Sub

Private Sub no_Click()
RESPUESTA = "N"
Unload Me

End Sub

Private Sub si_Click()
RESPUESTA = "S"
Unload Me

End Sub
