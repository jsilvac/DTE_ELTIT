VERSION 5.00
Object = "{ADBBAED6-B16F-40DC-80DF-B44910CBA76C}#1.0#0"; "Frame-Xp.ocx"
Begin VB.Form Ruta 
   BackColor       =   &H00008000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   5850
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6240
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5850
   ScaleWidth      =   6240
   StartUpPosition =   3  'Windows Default
   Begin XPFrame.FrameXp FrameXp1 
      Height          =   5595
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   9869
      BackColor       =   12648384
      Caption         =   "Seleccione una Ruta para Actualizaciones"
      CaptionEstilo3D =   1
      BackColor       =   12648384
      ColorBarraArriba=   12648384
      ColorBarraAbajo =   32768
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.Frame frmRuta 
         BackColor       =   &H00C0FFC0&
         Caption         =   " Buscar Carpeta "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   4455
         Left            =   180
         TabIndex        =   1
         Top             =   540
         Width           =   5655
         Begin VB.DriveListBox disco 
            BackColor       =   &H00008080&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0FFFF&
            Height          =   360
            Left            =   240
            TabIndex        =   5
            Top             =   300
            Width           =   5175
         End
         Begin VB.DirListBox dir 
            Appearance      =   0  'Flat
            BackColor       =   &H00008080&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0FFFF&
            Height          =   3600
            Left            =   240
            TabIndex        =   4
            Top             =   660
            Width           =   5175
         End
      End
      Begin XPFrame.FrameXp frmAceptar 
         Height          =   375
         Left            =   4140
         TabIndex        =   2
         Top             =   5100
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         BackColor       =   49344
         Caption         =   "Aceptar"
         CaptionEstilo3D =   1
         BackColor       =   49344
         ColorBarraArriba=   12648384
         ColorBarraAbajo =   32768
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
      End
      Begin XPFrame.FrameXp frmCancelar 
         Height          =   375
         Left            =   2280
         TabIndex        =   3
         Top             =   5100
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         BackColor       =   49344
         Caption         =   "Cancelar"
         CaptionEstilo3D =   1
         BackColor       =   49344
         ColorBarraArriba=   12648384
         ColorBarraAbajo =   32768
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
      End
   End
End
Attribute VB_Name = "Ruta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub lblCaja_Click()
End Sub

Private Sub Label5_Click()

End Sub

Private Sub disco_Change()
    dir.Path = disco.Drive
End Sub

Private Sub Form_Load()
    disco.Drive = "C:"
    dir.Path = disco.Drive
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If KeyCode = 27 Then
        Unload Me
    End If
End Sub

    Private Sub frmCancelar_BarClick()
        Call cambiaColor(frmCancelar)
        frmCancelar.CaptionEstilo3D = Inserted
        Unload Me
    End Sub
    
    Private Sub frmCancelar_BarMouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
        Call cambiaColor(frmCancelar)
        frmCancelar.CaptionEstilo3D = Raised
    End Sub
    
    Private Sub frmAceptar_BarClick()
        Call cambiaColor(frmAceptar)
        frmAceptar.CaptionEstilo3D = Inserted
        rutaUpdate = UCase(dir.Path)
        Unload Me
    End Sub
    
    Private Sub frmAceptar_BarMouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
        Call cambiaColor(frmAceptar)
        frmAceptar.CaptionEstilo3D = Raised
    End Sub

