VERSION 5.00
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCHRT20.OCX"
Begin VB.Form VentasUnidadRangoGrafico 
   BackColor       =   &H00C0FFC0&
   BorderStyle     =   0  'None
   Caption         =   "Grafico"
   ClientHeight    =   5025
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6600
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   335
   ScaleMode       =   0  'User
   ScaleWidth      =   440
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
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
      TabIndex        =   3
      Top             =   4500
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
      TabIndex        =   2
      Top             =   4500
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
      Left            =   5340
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4500
      Width           =   1155
   End
   Begin MSChart20Lib.MSChart Graf 
      Height          =   4395
      Left            =   0
      OleObjectBlob   =   "VentasUnidadRangoGrafico.frx":0000
      TabIndex        =   1
      Top             =   0
      Width           =   6555
   End
End
Attribute VB_Name = "VentasUnidadRangoGrafico"
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
    With VentasUnidadRango
        ReDim datos(1 To .impresion.Cols - 3) As Variant
        For i = 1 To .impresion.Cols - 3
            datos(i) = Format(.impresion.Cell(fila1, i + 2).text, "########0")
        Next i
        Graf.ChartData = datos
        For i = 1 To .impresion.Cols - 3
            Graf.Column = i
            Graf.ColumnLabel = .impresion.Cell(0, i + 2).text
        Next i
        Graf.row = 1
        If .impresion.Cell(fila1, 2).text <> "" Then
            Graf.RowLabel = .impresion.Cell(fila1, 2).text
        Else
            Graf.RowLabel = .impresion.Cell(fila1, 1).text
        End If
        Graf.TitleText = "DISTRIBUCION DE VENTAS POR LOCAL POR KILOS" & primerAño & " - " & segundoAño
    End With
    
    Graf.ShowLegend = True
    Call TranslucentForm(Me, 200)
End Sub

Private Sub guardar()
    Dim strArchivoGuardar As String
    
    strArchivoGuardar = App.Path & "\[" & primerAño & " " & segundoAño & "] - VENTAS POR KILOS" & ".bmp"
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

