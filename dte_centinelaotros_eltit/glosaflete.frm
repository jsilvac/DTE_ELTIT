VERSION 5.00
Object = "{ADBBAED6-B16F-40DC-80DF-B44910CBA76C}#1.0#0"; "Frame-Xp.ocx"
Begin VB.Form glosaflete 
   Caption         =   "GLOSA"
   ClientHeight    =   3480
   ClientLeft      =   60
   ClientTop       =   465
   ClientWidth     =   6885
   LinkTopic       =   "Form1"
   ScaleHeight     =   3480
   ScaleWidth      =   6885
   StartUpPosition =   3  'Windows Default
   Begin XPFrame.FrameXp Detalleflete 
      Height          =   3495
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   6165
      BackColor       =   12632256
      Caption         =   "Detalle Flete"
      CaptionEstilo3D =   1
      BackColor       =   12632256
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
      Begin VB.CommandButton Finalizar 
         BackColor       =   &H00FF8080&
         Caption         =   "Finalizar"
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
         Left            =   5160
         Style           =   1  'Graphical
         TabIndex        =   7
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
         Height          =   2175
         Left            =   120
         MaxLength       =   200
         MultiLine       =   -1  'True
         TabIndex        =   6
         Top             =   840
         Width           =   6615
      End
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4320
         MaxLength       =   2
         TabIndex        =   4
         Text            =   "20"
         Top             =   450
         Width           =   495
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2400
         MaxLength       =   2
         TabIndex        =   2
         Text            =   "08"
         Top             =   450
         Width           =   495
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   ":00 HRS  "
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
         Left            =   4850
         TabIndex        =   5
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   ":00 HRS  Y LAS"
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
         Left            =   2950
         TabIndex        =   3
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "ENTREGAR ENTRE LAS"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   480
         TabIndex        =   1
         Top             =   480
         Width           =   1935
      End
   End
End
Attribute VB_Name = "glosaflete"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Finalizar_Click()
glosa1flete = Label1.Caption & " " & Text1.text & Label2.Caption & " " & Text2.text & Label3.Caption
glosa2flete = glosa2.text

Unload Me

End Sub


Private Sub glosa2_KeyPress(KeyAscii As Integer)
  KeyAscii = Asc(UCase(Chr(KeyAscii)))
If KeyAscii = 13 Then
Finalizar.Visible = True
End If


End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
KeyAscii = esNumero(KeyAscii)
If KeyAscii = 13 And Text1.text <> "" And Text1.text < "23" Then
Text2.SetFocus
End If
End Sub


Private Sub Text2_KeyPress(KeyAscii As Integer)
KeyAscii = esNumero(KeyAscii)
If KeyAscii = 13 And Text2.text <> "" And Text2.text < 23 Then
glosa2.SetFocus
End If
End Sub
