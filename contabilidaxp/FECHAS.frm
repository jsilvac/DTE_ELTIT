VERSION 5.00
Object = "{4777436C-EB8C-4596-98A8-EBCF98683969}#1.0#0"; "FlexCell.ocx"
Object = "{085FBDF3-00C5-421B-B762-1D57299A2B89}#1.0#0"; "CLBUTN.OCX"
Object = "{ADBBAED6-B16F-40DC-80DF-B44910CBA76C}#1.0#0"; "Frame-Xp.ocx"
Begin VB.Form FECHAS 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " "
   ClientHeight    =   3795
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   5340
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3795
   ScaleWidth      =   5340
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin CoolButtons.cool_Button COMMAND1 
      Height          =   495
      Left            =   1680
      TabIndex        =   1
      Top             =   2640
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   873
      SkinId          =   "3"
      Caption         =   "ACEPTAR"
      ForeColor       =   8388608
   End
   Begin XPFrame.FrameXp FrameXp1 
      Height          =   1455
      Left            =   240
      TabIndex        =   0
      Top             =   480
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   2566
      BackColor       =   16744576
      Caption         =   "RANGO DE FECHAS"
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
      Alignment       =   1
      Begin FlexCell.Grid calendario 
         Height          =   855
         Left            =   240
         TabIndex        =   2
         Top             =   360
         Width           =   4335
         _ExtentX        =   7646
         _ExtentY        =   1508
         Cols            =   5
         DefaultFontSize =   8.25
         Rows            =   30
         DateFormat      =   2
      End
   End
End
Attribute VB_Name = "FECHAS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public fecha1 As label
Public fecha2 As label

Private Sub Command1_Click()
If calendario.Cell(1, 1).text = "" Then calendario.Cell(1, 1).SetFocus: GoTo no:
If calendario.Cell(1, 2).text = "" Then calendario.Cell(1, 2).SetFocus: GoTo no:
If DateDiff("d", calendario.Cell(1, 1).text, calendario.Cell(1, 2).text) < 0 Then calendario.Cell(1, 2).SetFocus: GoTo no:
Rem validar

fecha1.Caption = calendario.Cell(1, 1).text
fecha2.Caption = calendario.Cell(1, 2).text

Unload Me
no:
End Sub

Private Sub Form_Load()
Call FORMATOCALENDARIO(calendario)

End Sub
