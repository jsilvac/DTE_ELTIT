VERSION 5.00
Object = "{ADBBAED6-B16F-40DC-80DF-B44910CBA76C}#1.0#0"; "Frame-Xp.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4995
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9030
   DrawMode        =   6  'Mask Pen Not
   LinkTopic       =   "Form1"
   ScaleHeight     =   4995
   ScaleWidth      =   9030
   StartUpPosition =   3  'Windows Default
   Begin XPFrame.FrameXp FrameXp5 
      Height          =   1095
      Left            =   1080
      TabIndex        =   0
      Top             =   1080
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   1931
      BackColor       =   16761024
      Caption         =   "Barra Opciones"
      CaptionEstilo3D =   1
      BackColor       =   16761024
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.CommandButton Command8 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Modifica"
         Height          =   855
         Left            =   5880
         Picture         =   "Form1.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton Command7 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Modifica"
         Height          =   855
         Left            =   5040
         Picture         =   "Form1.frx":36A1
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton Command6 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Modifica"
         Height          =   855
         Left            =   4200
         Picture         =   "Form1.frx":6D42
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton Command5 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Modifica"
         Height          =   855
         Left            =   3360
         Picture         =   "Form1.frx":A3E3
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton Command4 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Modifica"
         Height          =   855
         Left            =   2520
         Picture         =   "Form1.frx":DA84
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Modifica"
         Height          =   855
         Left            =   1680
         Picture         =   "Form1.frx":11125
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Modifica"
         Height          =   855
         Left            =   840
         Picture         =   "Form1.frx":147C6
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Modifica"
         Height          =   855
         Left            =   0
         Picture         =   "Form1.frx":17E67
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   240
         Width           =   855
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub modifica1_Click()
If modifica1.Visible = True Then
modifica1.Visible = False
modifica2.Visible = True

End If

End Sub
Private Sub modifica2_Click()
If modifica2.Visible = True Then
modifica1.Visible = True
modifica2.Visible = False

End If

End Sub

Private Sub Label1_Click()

End Sub

Private Sub Image7_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image7.Visible = True

End Sub

Private Sub Image13_Click()
End Sub

Private Sub Image8_Click()
Image8.

Stop
End Sub
