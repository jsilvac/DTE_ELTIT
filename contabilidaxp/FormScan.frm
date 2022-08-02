VERSION 5.00
Object = "{ADBBAED6-B16F-40DC-80DF-B44910CBA76C}#1.0#0"; "Frame-Xp.ocx"
Begin VB.Form escaneo 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Pantalla de escaneo de Facturas"
   ClientHeight    =   10440
   ClientLeft      =   2370
   ClientTop       =   315
   ClientWidth     =   8190
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10440
   ScaleWidth      =   8190
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin XPFrame.FrameXp FrameXp1 
      Height          =   8775
      Left            =   120
      TabIndex        =   10
      Top             =   1680
      Width           =   8055
      _ExtentX        =   14208
      _ExtentY        =   15478
      BackColor       =   16744576
      Caption         =   "IMAGEN"
      CaptionEstilo3D =   1
      BackColor       =   16744576
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
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FF8080&
      Height          =   735
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   8055
      Begin VB.CommandButton Command4 
         Caption         =   "ZOOM -"
         Height          =   375
         Left            =   6840
         TabIndex        =   9
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton Command1 
         Caption         =   "ZOOM +"
         Height          =   375
         Left            =   5640
         TabIndex        =   8
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton Command3 
         Caption         =   "ROTAR 180º"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4560
         TabIndex        =   7
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton Command5 
         Caption         =   "ROTAR 90º"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   1095
      End
      Begin VB.HScrollBar HScroll1 
         Height          =   375
         Left            =   1320
         TabIndex        =   5
         Top             =   240
         Width           =   3135
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FF8080&
      Caption         =   "SELECCION DE SCANNER A UTILIZAR "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   1
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8055
      Begin VB.CommandButton Command2 
         Caption         =   "USAR IMAGEN"
         Height          =   375
         Left            =   6120
         TabIndex        =   3
         Top             =   240
         Width           =   1575
      End
      Begin VB.ComboBox cboimagesource 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   240
         Width           =   3495
      End
      Begin VB.CommandButton CmdScan 
         Caption         =   "ESCANEAR"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3720
         TabIndex        =   1
         Top             =   240
         Width           =   2175
      End
   End
End
Attribute VB_Name = "escaneo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim StrFile, StrType As String
Private Sub CmdScan_Click()

    cboimagesource.AddItem Scanner1.GetImageSourceName(i)
    Scanner1.SelectImageSourceByIndex cboimagesource.ListIndex
    Scanner1.DuplexEnabled = False
    Scanner1.FeederEnabled = False
    Scanner1.DPI = 96
    Scanner1.PixelType = -1
    Scanner1.SetCaptureArea 0, 0, 0, 0
    Scanner1.View = 0
    Scanner1.ApplyChange
    Scanner1.Scan
    Call GuardaIMG
    Scanner1.View = 10
If MsgBox("DESEA USAR ESTA IMAGEN?", vbYesNo, "ATENCION") = vbYes Then UsarImg
End Sub

Sub GuardaIMG()
'result = Me.Scanner1.Save(StrFile, StrType)
End Sub
Sub UsarImg()
Call GuardaIMG
        With IngresoImgFactura
        .dialogo.FileName = StrFile & ".JPG"
        .imagen.Picture = LoadPicture(.dialogo.FileName)
        KbImagen = Mid(Str(FileLen(.dialogo.FileName)), 1, Len(Str(FileLen(.dialogo.FileName))) - 3)
        .kb = KbImagen
    End With
    Unload escaneo

End Sub

Private Sub Command1_Click()
Scanner1.View = Scanner1.View + 1
Scanner1.SetFocus
Call GuardaIMG
End Sub

Private Sub COMMAND2_Click()
'Scanner1.View = -700
'    result = Me.Scanner1.Save(StrFile, StrType)
'        With IngresoImgFactura
'        .dialogo.FileName = StrFile & ".JPG"
'        .imagen.Picture = LoadPicture(.dialogo.FileName)
'        KbImagen = Mid(Str(FileLen(.dialogo.FileName)), 1, Len(Str(FileLen(.dialogo.FileName))) - 3)
'        .kb = KbImagen
'    End With
'Unload escaneo
End Sub
Private Sub Command3_Click()
'Me.Scanner1.Rotate180
Call GuardaIMG
End Sub

Private Sub Command4_Click()
Scanner1.View = Scanner1.View - 1
Scanner1.SetFocus
Call GuardaIMG
End Sub

Private Sub Command5_Click()
'Me.Scanner1.Rotate90
Call GuardaIMG
End Sub
Private Sub Form_Load()
StrFile = "C:\TMP"
StrType = "jpg"
iCount = Scanner1.GetNumImageSources
For i = 0 To iCount - 1
cboimagesource.AddItem Scanner1.GetImageSourceName(i)
Next
If iCount > 0 Then
    cboimagesource.ListIndex = 0
End If
End Sub
Private Sub HScroll1_Change()
Scanner1.RotateAt HScroll1.Value
End Sub
