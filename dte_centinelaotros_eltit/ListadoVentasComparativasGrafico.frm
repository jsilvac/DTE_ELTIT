VERSION 5.00
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCHRT20.OCX"
Begin VB.Form ListadoVentasComparativasGrafico 
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
   ScaleMode       =   3  'Pixel
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
      OleObjectBlob   =   "ListadoVentasComparativasGrafico.frx":0000
      TabIndex        =   1
      Top             =   0
      Width           =   6555
   End
End
Attribute VB_Name = "ListadoVentasComparativasGrafico"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Private datos() As Variant

Private Sub cmdGuardar_Click()
    Call guardar
End Sub

Private Sub cmdImprimir_Click()
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
    Dim cadena1 As String
    Dim cadena2 As String
    With ListadoVentasComparativas
        Graf.ChartType = tipoGrafico
        
        ReDim datos(1 To 2) As Variant
        datos(1) = Format(.impresion.Cell(fila1, col1).text, "########0")
        If tipoGrafico = 14 Then
            datos(2) = Format(CDbl(.impresion.Cell(fila1, col2).text) - CDbl(.impresion.Cell(fila1, col1).text), "########0")
        Else
            datos(2) = Format(.impresion.Cell(fila1, col2).text, "########0")
        End If
        Graf.ChartData = datos
        Graf.row = 1
        Graf.RowLabel = .impresion.Cell(1, col1).text & " - " & .impresion.Cell(fila1, 1).text
        
        If tipoGrafico = 14 Then
            Graf.Column = 1
            cadena1 = "MES ACTUAL"
            Graf.ColumnLabel = cadena1
            Graf.Column = 2
            cadena2 = "ACUMULADO"
            Graf.ColumnLabel = cadena2
        Else
            Graf.Column = 1
            cadena1 = segundoAño
            Graf.ColumnLabel = cadena1
            Graf.Column = 2
            cadena2 = primerAño
            Graf.ColumnLabel = cadena2
        End If

        Graf.TitleText = "COMPARATIVA " & cadena1 & " VS " & cadena2
    End With
    
    Graf.ShowLegend = True
    Call TranslucentForm(Me, 200)
End Sub

Private Sub guardar()
    Dim strArchivoGuardar As String
    
    strArchivoGuardar = App.Path & "\[" & primerAño & " " & segundoAño & "] - " & NOMBRECLIENTE & ".bmp"
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

