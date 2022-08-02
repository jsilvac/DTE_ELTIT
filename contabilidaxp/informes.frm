VERSION 5.00
Begin VB.Form informes 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mestro de Secciones"
   ClientHeight    =   9990
   ClientLeft      =   45
   ClientTop       =   420
   ClientWidth     =   15225
   DrawStyle       =   5  'Transparent
   BeginProperty Font 
      Name            =   "Courier New"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   666
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1015
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command3 
      BackColor       =   &H00FFC0C0&
      Caption         =   "+"
      Height          =   255
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   0
      Width           =   615
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00FFC0C0&
      Caption         =   "-"
      Height          =   255
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   0
      Width           =   615
   End
   Begin VB.ListBox info 
      BackColor       =   &H00FFFFFF&
      Height          =   9660
      IntegralHeight  =   0   'False
      Left            =   0
      TabIndex        =   0
      Top             =   240
      Width           =   15135
   End
End
Attribute VB_Name = "informes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
impresora.ShowPrinter
Printer.FontName = "Courier new"
Printer.FontSize = 7
pagina = 0
For K = 0 To info.ListCount
For i = 1 To Len(info.List(K))
If Mid(info.List(K), i, 6) = "PAGINA" And pagina > 0 Then Printer.NewPage
Next i
pagina = 1
Printer.Print info.List(K)
Next K

Printer.EndDoc


End Sub

Private Sub Command2_Click()
Open "ARCHIVO.TXT" For Output As #20
For K = 0 To info.ListCount
palabra = info.List(K)
Print #20, palabra
Next K
Close 20
Shell "WORDPAD.EXE ARCHIVO.TXT"
End Sub

Private Sub Dir1_Change()

End Sub

Private Sub Command3_Click()
info.FontSize = info.FontSize + 1
End Sub

Private Sub Command4_Click()
info.FontSize = info.FontSize - 1
End Sub

