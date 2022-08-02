VERSION 5.00
Object = "{2B5A7812-71D1-4C51-B59B-AA38CD8D6BA3}#6.0#0"; "VB2_SkinControlLt.ocx"
Object = "{085FBDF3-00C5-421B-B762-1D57299A2B89}#1.0#0"; "CLBUTN.OCX"
Begin VB.Form MENSAJES 
   BackColor       =   &H00C0FFC0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mensajes Al Operador"
   ClientHeight    =   2985
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5175
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2985
   ScaleWidth      =   5175
   StartUpPosition =   2  'CenterScreen
   Begin VB2_SkinControlLt.VB2_SkinCtrlLt VB2_SkinCtrlLt1 
      Left            =   0
      Top             =   0
      _ExtentX        =   1111
      _ExtentY        =   953
      SkinPicture     =   "MENSAJES.frx":0000
      Skin            =   2
   End
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
      Begin CoolButtons.cool_Button command1 
         Height          =   375
         Left            =   840
         TabIndex        =   1
         Top             =   1560
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   661
         Caption         =   "Continuar"
         ForeColor       =   16711680
      End
      Begin VB.Label TEXTOMENSAJE 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Desea Realmente Eliminar El Documento "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   855
         Left            =   360
         TabIndex        =   4
         Top             =   480
         Width           =   3615
      End
      Begin VB.Label Label8 
         BackColor       =   &H8000000A&
         BorderStyle     =   1  'Fixed Single
         Height          =   1935
         Left            =   120
         TabIndex        =   3
         Top             =   120
         Width           =   4095
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00E0E0E0&
         BorderWidth     =   2
         Height          =   2175
         Left            =   0
         Top             =   0
         Width           =   4335
      End
      Begin VB.Label mensaje 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Desea Realmente Eliminar El Documento "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   855
         Left            =   360
         TabIndex        =   2
         Top             =   480
         Width           =   3615
      End
   End
End
Attribute VB_Name = "MENSAJES"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Unload Me
End Sub

Private Sub Form_Load()
VB2_SkinCtrlLt1.ActivateSkin
End Sub
