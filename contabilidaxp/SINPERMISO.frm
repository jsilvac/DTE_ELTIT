VERSION 5.00
Begin VB.Form SINPERMISO 
   BackColor       =   &H00F5C9B1&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   1950
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5670
   LinkTopic       =   "Form1"
   ScaleHeight     =   1950
   ScaleWidth      =   5670
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H000040C0&
      BorderStyle     =   0  'None
      Caption         =   "SISTEMA DE SEGURIDAD ACTIVADO"
      Height          =   1815
      Left            =   80
      TabIndex        =   0
      Top             =   70
      Width           =   5535
      Begin VB.CommandButton Command1 
         BackColor       =   &H0080C0FF&
         Caption         =   "Retorno"
         Height          =   255
         Left            =   1680
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   1320
         Width           =   2415
      End
      Begin VB.Shape Shape2 
         Height          =   735
         Left            =   360
         Top             =   480
         Width           =   4935
      End
      Begin VB.Label MENSAJESEGURIDAD 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   " USUARIO NO PUEDE CREAR REGISTROS"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   495
         Left            =   0
         TabIndex        =   3
         Top             =   720
         Width           =   5655
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0080C0FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "MODULO DE SEGURIDAD ACTIVADO"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   240
         Width           =   5175
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00FFC0C0&
         BorderWidth     =   3
         Height          =   1815
         Left            =   0
         Top             =   0
         Width           =   5535
      End
   End
End
Attribute VB_Name = "SINPERMISO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
Unload Me

End Sub

