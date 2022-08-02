VERSION 5.00
Object = "{4777436C-EB8C-4596-98A8-EBCF98683969}#1.0#0"; "Flexcell.ocx"
Object = "{ADBBAED6-B16F-40DC-80DF-B44910CBA76C}#1.0#0"; "Frame-Xp.ocx"
Begin VB.Form detalleauditoria 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   9120
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9180
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   9120
   ScaleWidth      =   9180
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin XPFrame.FrameXp FrameXp1 
      Height          =   9105
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9150
      _ExtentX        =   16140
      _ExtentY        =   16060
      BackColor       =   16761024
      Caption         =   "DETALLE AUDITORIA DE EVENTOS"
      CaptionEstilo3D =   1
      BackColor       =   16761024
      ForeColor       =   1563884
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   1
      Begin VB.CommandButton Command1 
         Caption         =   "RETORNO (ESC)"
         Height          =   285
         Left            =   3555
         TabIndex        =   2
         Top             =   8730
         Width           =   1995
      End
      Begin FlexCell.Grid Grid1 
         Height          =   8295
         Left            =   45
         TabIndex        =   1
         Top             =   405
         Width           =   9060
         _ExtentX        =   15981
         _ExtentY        =   14631
         Cols            =   5
         DefaultFontSize =   8.25
         Rows            =   30
      End
   End
End
Attribute VB_Name = "detalleauditoria"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private formatogrilla(20, 20)
Private CAMPOS As String
Private ORIGINALES As String
Private MODIFICADOS As String

Private Sub Command1_Click()
Unload Me

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
Unload Me

End If


End Sub

Private Sub Form_Load()
Call CARGAGRILLA(4, 1)
CARGADATOS


End Sub
Sub CARGADATOS()
Dim K As Integer
Dim PASO As Integer
Dim INICIO As Integer
Grid1.Rows = 10

Grid1.Cell(1, 1).text = moduloseguridad2.Grid1.Cell(0, 1).text
Grid1.Cell(2, 1).text = moduloseguridad2.Grid1.Cell(0, 2).text
Grid1.Cell(3, 1).text = moduloseguridad2.Grid1.Cell(0, 3).text
Grid1.Cell(4, 1).text = moduloseguridad2.Grid1.Cell(0, 4).text
Grid1.Cell(5, 1).text = moduloseguridad2.Grid1.Cell(0, 5).text
Grid1.Cell(6, 1).text = moduloseguridad2.Grid1.Cell(0, 6).text
Grid1.Cell(7, 1).text = moduloseguridad2.Grid1.Cell(0, 7).text
Grid1.Cell(8, 1).text = moduloseguridad2.Grid1.Cell(0, 8).text
Grid1.Cell(9, 1).text = moduloseguridad2.Grid1.Cell(0, 9).text

Grid1.Cell(1, 2).text = moduloseguridad2.Grid1.Cell(moduloseguridad2.Grid1.ActiveCell.Row, 1).text
Grid1.Cell(2, 2).text = moduloseguridad2.Grid1.Cell(moduloseguridad2.Grid1.ActiveCell.Row, 2).text
Grid1.Cell(3, 2).text = moduloseguridad2.Grid1.Cell(moduloseguridad2.Grid1.ActiveCell.Row, 3).text
Grid1.Cell(4, 2).text = moduloseguridad2.Grid1.Cell(moduloseguridad2.Grid1.ActiveCell.Row, 4).text
Grid1.Cell(5, 2).text = moduloseguridad2.Grid1.Cell(moduloseguridad2.Grid1.ActiveCell.Row, 5).text
Grid1.Cell(6, 2).text = moduloseguridad2.Grid1.Cell(moduloseguridad2.Grid1.ActiveCell.Row, 6).text
Grid1.Cell(7, 2).text = moduloseguridad2.Grid1.Cell(moduloseguridad2.Grid1.ActiveCell.Row, 7).text
Grid1.Cell(8, 2).text = moduloseguridad2.Grid1.Cell(moduloseguridad2.Grid1.ActiveCell.Row, 8).text
Grid1.Cell(9, 2).text = moduloseguridad2.Grid1.Cell(moduloseguridad2.Grid1.ActiveCell.Row, 9).text




CAMPOS = moduloseguridad2.Grid1.Cell(moduloseguridad2.Grid1.ActiveCell.Row, 10).text
ORIGINALES = moduloseguridad2.Grid1.Cell(moduloseguridad2.Grid1.ActiveCell.Row, 11).text
MODIFICADOS = moduloseguridad2.Grid1.Cell(moduloseguridad2.Grid1.ActiveCell.Row, 12).text
PASO = 0
INICIO = 0
For K = 1 To Len(CAMPOS)
If Mid(CAMPOS, K, 1) = "]" Then
PASO = PASO + 1
Grid1.Rows = Grid1.Rows + 1
Grid1.Cell(Grid1.Rows - 1, 1).text = Mid(CAMPOS, INICIO + 1, K - INICIO)
INICIO = K
End If




Next K
INICIO = 0
PASO = 9
For K = 1 To Len(ORIGINALES)
If Mid(ORIGINALES, K, 1) = "]" Then
PASO = PASO + 1

Grid1.Cell(PASO, 2).text = Mid(ORIGINALES, INICIO + 1, K - INICIO)
INICIO = K
End If
Next K

INICIO = 0
PASO = 9
For K = 1 To Len(MODIFICADOS)
If Mid(MODIFICADOS, K, 1) = "]" Then
PASO = PASO + 1

Grid1.Cell(PASO, 3).text = Mid(MODIFICADOS, INICIO + 1, K - INICIO)
INICIO = K
End If
Next K


End Sub
Sub CARGAGRILLA(Row, Col)
    Rem DATOS DE LA COLUMNA
    Col = 4
    Row = 1
    formatogrilla(1, 1) = "CAMPOS"
    formatogrilla(1, 2) = "ORIGINALES"
    formatogrilla(1, 3) = "MODIFICADOS"
    
    
    Rem LARGO DE LOS DATOS
    
    formatogrilla(2, 1) = "20"
    formatogrilla(2, 2) = "20"
    formatogrilla(2, 3) = "20"
    
    Rem TIPO DE DATOS C=CEROS,S=STRING,N=NUMERICO,D=DATE,
    formatogrilla(3, 1) = "S"
    formatogrilla(3, 2) = "S"
    formatogrilla(3, 3) = "S"
    
    Rem FORMATO GRILLA
    formatogrilla(4, 1) = ""
    formatogrilla(4, 2) = ""
    Rem LOCCKED
    For K = 1 To 3
    formatogrilla(5, K) = "true"
    Next K
    
    
    Grid1.Cols = Col
    Grid1.Rows = Row
    Grid1.AllowUserResizing = True
    Grid1.DisplayFocusRect = False
    Grid1.AllowUserSort = True
    Grid1.ExtendLastCol = True
    Grid1.BoldFixedCell = False
    Grid1.DrawMode = cellOwnerDraw
    Grid1.Appearance = Flat
    Grid1.ScrollBarStyle = Flat
    Grid1.FixedRowColStyle = Flat
    
    Grid1.Column(0).Width = 0
    
    
            For K = 1 To Col - 1
            Grid1.Cell(0, K).text = formatogrilla(1, K)
            Grid1.Column(K).Width = Val(formatogrilla(2, K)) * 10
            Grid1.Column(K).MaxLength = Val(formatogrilla(2, K))
            Rem GRID1.Column(k).FormatString = formatoGrilla(4, k)
            Grid1.Column(K).Locked = formatogrilla(5, K)
            If formatogrilla(3, K) = "S" Then
                Grid1.Column(K).Alignment = cellLeftCenter
            Else
                
                Grid1.Column(K).Alignment = cellRightCenter
            End If
            Grid1.Cell(0, K).Alignment = cellCenterCenter
        Next K
    
  
    
    
End Sub

