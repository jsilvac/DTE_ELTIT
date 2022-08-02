VERSION 5.00
Object = "{ADBBAED6-B16F-40DC-80DF-B44910CBA76C}#1.0#0"; "Frame-Xp.ocx"
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCHRT20.OCX"
Begin VB.Form ComprasClienteGrafico 
   BackColor       =   &H00C0FFC0&
   BorderStyle     =   0  'None
   Caption         =   "Grafico"
   ClientHeight    =   5445
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6690
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   363
   ScaleMode       =   0  'User
   ScaleWidth      =   446
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin XPFrame.FrameXp FrameXp1 
      Height          =   4875
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   6675
      _ExtentX        =   11774
      _ExtentY        =   8599
      BackColor       =   49152
      Caption         =   "GRAFICO"
      CaptionEstilo3D =   1
      BackColor       =   49152
      ColorBarraArriba=   12648384
      ColorBarraAbajo =   32768
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   1
      Begin MSChart20Lib.MSChart Graf 
         Height          =   4395
         Left            =   60
         OleObjectBlob   =   "ComprasClienteGrafico.frx":0000
         TabIndex        =   4
         Top             =   360
         Width           =   6555
      End
   End
   Begin VB.CommandButton cmdGuardar 
      BackColor       =   &H0000FF00&
      Caption         =   "G - Guardar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1440
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4980
      Width           =   1155
   End
   Begin VB.CommandButton cmdimprimir 
      BackColor       =   &H0000FF00&
      Caption         =   "I - Imprimir"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4980
      Width           =   1155
   End
   Begin VB.CommandButton cmdSalir 
      BackColor       =   &H0000FF00&
      Caption         =   "Esc - Salir"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5400
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4980
      Width           =   1155
   End
End
Attribute VB_Name = "ComprasClienteGrafico"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Private datos() As Variant

Private Sub cmdGuardar_Click()
    Call guardar
End Sub

Private Sub cmdimprimir_Click()
    Call imprimir
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 27
            Unload Me
        Case Asc("i"), Asc("I")
            Call imprimir
        Case Asc("g"), Asc("G")
            Call guardar
    End Select
End Sub

Private Sub Form_Load()
    Dim i As Integer
    With ComprasCliente
        ReDim datos(1 To cantMeses, 1 To 2) As Variant
        Graf.ColumnCount = cantMeses
        Graf.RowCount = 2
        For i = 1 To cantMeses
            Graf.Column = i
            Graf.ColumnLabel = meses(i)
            datos(i, 1) = Format(.impresion.Cell(fila1, i + 2).text, "########0")
            datos(i, 2) = Format(.impresion.Cell(fila2, i + 2).text, "########0")
        Next i
        Graf.ChartData = datos
        For i = 1 To cantMeses
            Graf.row = i
            Graf.RowLabel = Left(meses(i), 3)
        Next i
        Graf.Column = 1
        Graf.ColumnLabel = primerAño
        Graf.Column = 2
        Graf.ColumnLabel = segundoAño
        Graf.TitleText = "COMPRAS CLIENTE " & nombrecliente
    End With
    
    Graf.ShowLegend = True
    Call TranslucentForm(Me, 200)
End Sub

Private Sub guardar()
    Dim strArchivoGuardar As String
    
    strArchivoGuardar = App.Path & "\[" & primerAño & " " & segundoAño & "] - " & nombrecliente & ".bmp"
    Graf.EditCopy
    Call SavePicture(Clipboard.GetData, strArchivoGuardar)
    MsgBox "El gráfico ha sido guardado en " & strArchivoGuardar, vbInformation, "Guardar Gráfico"
End Sub

Private Sub imprimir()
    Set Printer = Printers(impOtros(1))
    Graf.EditCopy
    Printer.PaintPicture Clipboard.GetData, 0, 0
    Printer.NewPage
    Printer.EndDoc
    
    MsgBox "El gráfico ha sido enviado para su impresión.", vbInformation, "Imprimir gráfico"
    
End Sub

