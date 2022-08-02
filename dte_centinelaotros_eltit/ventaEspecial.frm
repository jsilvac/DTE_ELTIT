VERSION 5.00
Object = "{ADBBAED6-B16F-40DC-80DF-B44910CBA76C}#1.0#0"; "Frame-Xp.ocx"
Begin VB.Form ventaEspecial 
   BackColor       =   &H00BC4A36&
   BorderStyle     =   0  'None
   ClientHeight    =   2655
   ClientLeft      =   3105
   ClientTop       =   2685
   ClientWidth     =   4515
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2655
   ScaleWidth      =   4515
   ShowInTaskbar   =   0   'False
   Begin XPFrame.FrameXp FrameXp1 
      Height          =   2430
      Left            =   90
      TabIndex        =   0
      Top             =   120
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   4286
      BackColor       =   16744576
      Caption         =   "Configurar Impresora"
      CaptionEstilo3D =   1
      BackColor       =   16744576
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   1
      Begin VB.CommandButton cmdF1 
         BackColor       =   &H00FFC0C0&
         Caption         =   " F1"
         Enabled         =   0   'False
         Height          =   285
         Left            =   225
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   1575
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.CommandButton cmdF2 
         BackColor       =   &H00FFC0C0&
         Caption         =   "I"
         Height          =   285
         Left            =   225
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   495
         Width           =   495
      End
      Begin VB.CommandButton cmdF3 
         BackColor       =   &H00FFC0C0&
         Caption         =   "X"
         Height          =   285
         Left            =   225
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   855
         Width           =   495
      End
      Begin VB.CommandButton cmdF4 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Z"
         Height          =   285
         Left            =   225
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   1215
         Width           =   495
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "ESC PARA VOLVER"
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
         Left            =   825
         TabIndex        =   9
         Top             =   2025
         Width           =   2625
      End
      Begin VB.Label Label11 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Inicializa Caja"
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
         Height          =   285
         Left            =   820
         TabIndex        =   8
         Top             =   495
         Width           =   2895
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Guardar/Recuperar Venta"
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
         Height          =   285
         Left            =   820
         TabIndex        =   7
         Top             =   1575
         Visible         =   0   'False
         Width           =   2895
      End
      Begin VB.Image Image16 
         Height          =   285
         Left            =   3840
         Stretch         =   -1  'True
         Top             =   480
         Width           =   285
      End
      Begin VB.Image Image15 
         Height          =   285
         Left            =   3840
         Stretch         =   -1  'True
         Top             =   840
         Width           =   285
      End
      Begin VB.Label Label31 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Imprime Informe Z"
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
         Height          =   285
         Left            =   820
         TabIndex        =   6
         Top             =   1215
         Width           =   2895
      End
      Begin VB.Label Label32 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Imprime Informe X"
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
         Height          =   285
         Left            =   820
         TabIndex        =   5
         Top             =   855
         Width           =   2895
      End
      Begin VB.Image Image1 
         Height          =   285
         Left            =   3840
         Stretch         =   -1  'True
         Top             =   1200
         Width           =   285
      End
      Begin VB.Image Image2 
         Height          =   285
         Left            =   3840
         Stretch         =   -1  'True
         Top             =   1560
         Width           =   285
      End
   End
End
Attribute VB_Name = "ventaEspecial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    
Private Sub cmdF1_Click()
    Call Form_KeyDown(vbKeyF1, 0)
End Sub

Private Sub cmdF2_Click()
    Call Form_KeyDown(Asc("I"), 0)
End Sub

Private Sub cmdF3_Click()
    Call Form_KeyDown(Asc("X"), 0)
End Sub

Private Sub cmdF4_Click()
    Call Form_KeyDown(Asc("Z"), 0)
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyF1
            Unload Me
            Principal.Show vbModal
        Case 73
            sw = True
            ImpresionBoleta.InicializaCaja
        Case 88
            sw = True
            ImpresionBoleta.imprimeX
        Case 90
            sw = True
            ImpresionBoleta.imprimeZ
       
        Case 27
            Unload Me
           
            
    End Select
End Sub

