VERSION 5.00
Object = "{4777436C-EB8C-4596-98A8-EBCF98683969}#1.0#0"; "FlexCell.ocx"
Begin VB.Form cartolas 
   ClientHeight    =   9600
   ClientLeft      =   255
   ClientTop       =   450
   ClientWidth     =   15105
   LinkTopic       =   "Form1"
   ScaleHeight     =   9600
   ScaleWidth      =   15105
   WindowState     =   2  'Maximized
   Begin FlexCell.Grid Grid1 
      Height          =   7935
      Left            =   0
      TabIndex        =   1
      Top             =   600
      Width           =   15135
      _ExtentX        =   26696
      _ExtentY        =   13996
      Cols            =   5
      DefaultFontSize =   8.25
      Rows            =   30
   End
   Begin VB.Label titulocartola 
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      TabIndex        =   0
      Top             =   0
      Width           =   9015
   End
End
Attribute VB_Name = "cartolas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Grid1_DblClick()
If Grid1.ActiveCell.col = 3 Then muestra
End Sub
Sub muestra()

PASO(1) = Grid1.Cell(Grid1.ActiveCell.row, 2).text
PASO(2) = Grid1.Cell(Grid1.ActiveCell.row, 3).text
PASO(3) = Mid(Grid1.Cell(Grid1.ActiveCell.row, 1).text, 1, 2)
PASO(4) = Mid(Grid1.Cell(Grid1.ActiveCell.row, 1).text, 4, 2)
PASO(5) = Mid(Grid1.Cell(Grid1.ActiveCell.row, 1).text, 7, 4)

muestracomprobantes.Show vbModal

End Sub
