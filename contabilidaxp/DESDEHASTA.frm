VERSION 5.00
Begin VB.Form DESDEHASTA 
   Caption         =   "Form1"
   ClientHeight    =   2310
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4485
   LinkTopic       =   "Form1"
   ScaleHeight     =   2310
   ScaleWidth      =   4485
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame FECHAS 
      BackColor       =   &H00FF8080&
      Height          =   2295
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4455
      Begin VB.CommandButton Command1 
         Caption         =   "RETORNO"
         Height          =   495
         Left            =   960
         TabIndex        =   9
         Top             =   1320
         Width           =   2295
      End
      Begin VB.TextBox HASTA3 
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
         Left            =   2880
         MaxLength       =   4
         TabIndex        =   6
         Tag             =   "fecha"
         Top             =   480
         Width           =   615
      End
      Begin VB.TextBox HASTA2 
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
         Left            =   2520
         MaxLength       =   2
         TabIndex        =   5
         Tag             =   "fecha"
         Top             =   480
         Width           =   375
      End
      Begin VB.TextBox HASTA1 
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
         Left            =   2160
         MaxLength       =   2
         TabIndex        =   4
         Tag             =   "fechavencimiento"
         Top             =   480
         Width           =   375
      End
      Begin VB.TextBox DESDE1 
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
         Left            =   600
         MaxLength       =   2
         TabIndex        =   3
         Tag             =   "fecha"
         Top             =   480
         Width           =   375
      End
      Begin VB.TextBox DESDE2 
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
         Left            =   960
         MaxLength       =   2
         TabIndex        =   2
         Tag             =   "fecha"
         Top             =   480
         Width           =   375
      End
      Begin VB.TextBox DESDE3 
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
         Left            =   1320
         MaxLength       =   4
         TabIndex        =   1
         Tag             =   "fecha"
         Top             =   480
         Width           =   615
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "DESDE"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   600
         TabIndex        =   8
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "HASTA"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2160
         TabIndex        =   7
         Top             =   240
         Width           =   1335
      End
   End
End
Attribute VB_Name = "DESDEHASTA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Not IsDate(DESDE1.text + DESDE2.text + DESDE3.text) Then DESDE1.SetFocus

Unload Me


End Sub

