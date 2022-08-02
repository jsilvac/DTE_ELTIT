VERSION 5.00
Object = "{4777436C-EB8C-4596-98A8-EBCF98683969}#1.0#0"; "FlexCell.ocx"
Object = "{ADBBAED6-B16F-40DC-80DF-B44910CBA76C}#1.0#0"; "Frame-Xp.ocx"
Begin VB.Form despacho 
   BackColor       =   &H00008000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   4335
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4515
   LinkTopic       =   "Form1"
   ScaleHeight     =   4335
   ScaleWidth      =   4515
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin XPFrame.FrameXp frmGlosa 
      Height          =   4095
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   7223
      BackColor       =   8454016
      Caption         =   "TIPO DESPACHO"
      CaptionEstilo3D =   1
      BackColor       =   8454016
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
      ColorTextShadow =   0
      Begin VB.TextBox tipodespacho 
         Height          =   375
         Left            =   120
         MaxLength       =   2
         TabIndex        =   2
         Top             =   3600
         Width           =   495
      End
      Begin FlexCell.Grid grilladespacho 
         Height          =   3015
         Left            =   120
         TabIndex        =   1
         Top             =   480
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   5318
         Cols            =   5
         DefaultFontSize =   8.25
         Rows            =   30
      End
      Begin VB.Label lblnombredespacho 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   720
         TabIndex        =   3
         Top             =   3600
         Width           =   3495
      End
   End
End
Attribute VB_Name = "despacho"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private formatogrilla(20, 20) As String
Private ctd As Double

Sub CARGAGRILLATIPOS()
Dim K As Integer
    Rem DATOS DE LA COLUMNA
    formatogrilla(1, 1) = "CODIGO"
    formatogrilla(1, 2) = "NOMBRE"
    
    Rem LARGO DE LOS DATOS
    formatogrilla(2, 1) = "10"
    formatogrilla(2, 2) = "20"
    
    Rem TIPO DE DATOS C=CEROS,S=STRING,N=NUMERICO,D=DATE,
    formatogrilla(3, 1) = "S"
    formatogrilla(3, 2) = "S"

    Rem FORMATO GRILLA
    formatogrilla(4, 1) = ""
    formatogrilla(4, 2) = ""
    
    Rem LOCCKED
    formatogrilla(5, 1) = "TRUE"
    formatogrilla(5, 2) = "TRUE"

    grilladespacho.Cols = 3
    grilladespacho.Rows = 2
    
    grilladespacho.AllowUserResizing = False
    grilladespacho.DisplayFocusRect = False
    grilladespacho.ExtendLastCol = True
    grilladespacho.BoldFixedCell = False
    grilladespacho.DrawMode = cellOwnerDraw
    grilladespacho.Appearance = Flat
    grilladespacho.ScrollBarStyle = Flat
    grilladespacho.FixedRowColStyle = Flat
    grilladespacho.BackColorFixed = RGB(90, 158, 214)
    grilladespacho.BackColorFixedSel = RGB(110, 180, 230)
    grilladespacho.BackColorBkg = RGB(90, 158, 214)
    grilladespacho.BackColorScrollBar = RGB(231, 235, 247)
    grilladespacho.BackColor1 = RGB(231, 235, 247)
    grilladespacho.BackColor2 = RGB(239, 243, 255)
    grilladespacho.GridColor = RGB(148, 190, 231)
    For K = 1 To grilladespacho.Cols - 1
        grilladespacho.Cell(0, K).text = formatogrilla(1, K)
        grilladespacho.Column(K).Width = Val(formatogrilla(2, K)) * grilladespacho.DefaultFont.Size
        grilladespacho.Column(K).MaxLength = Val(formatogrilla(2, K))
        grilladespacho.Column(K).FormatString = formatogrilla(4, K)
        grilladespacho.Column(K).Locked = formatogrilla(5, K)
        If formatogrilla(3, K) = "N" Then grilladespacho.Column(K).Alignment = cellRightCenter
       
    Next K
    grilladespacho.Column(0).Width = 0
    grilladespacho.Range(0, 0, 0, grilladespacho.Cols - 1).Alignment = cellCenterCenter
    grilladespacho.Enabled = False
End Sub

Private Sub Form_Load()
Call CARGAGRILLATIPOS
Call leertiposdespacho
End Sub
Sub leertiposdespacho()
Dim cSql As New rdoQuery
Dim resultados As rdoResultset
Dim linea As Double

Set cSql.ActiveConnection = ventas

cSql.sql = "select * "
cSql.sql = cSql.sql & "from sv_tipo_despacho "
cSql.sql = cSql.sql & "where local='" & empresaActiva & "' "
cSql.Execute
grilladespacho.Rows = cSql.RowsAffected + 1
    
linea = 0
If cSql.RowsAffected > 0 Then
ctd = cSql.RowsAffected
Set resultados = cSql.OpenResultset
        grilladespacho.AutoRedraw = False
        While Not resultados.EOF
           linea = linea + 1
           grilladespacho.Cell(linea, 1).text = resultados(0)
           grilladespacho.Cell(linea, 2).text = resultados(1)
          
            resultados.MoveNext
        Wend
        resultados.Close
        Set resultados = Nothing
        grilladespacho.AutoRedraw = True
        grilladespacho.Refresh

End If


End Sub

Private Sub tipodespacho_GotFocus()
If ctd = 1 Then
cotiza01.GRID1.Cell(cotiza01.GRID1.ActiveCell.Row, 11).text = "01"
Unload Me
End If

End Sub

Private Sub tipodespacho_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
cotiza01.GRID1.Cell(cotiza01.GRID1.ActiveCell.Row, 11).text = "RET"
Unload Me

End If

End Sub

Private Sub tipodespacho_KeyPress(KeyAscii As Integer)
KeyAscii = esNumero(KeyAscii)
Dim K As Integer

If KeyAscii = 13 And tipodespacho.text <> "" Then
For K = 1 To grilladespacho.Rows - 1
If CDbl(tipodespacho.text) = CDbl(grilladespacho.Cell(K, 1).text) Then
lblnombredespacho.Caption = grilladespacho.Cell(K, 2).text
tipodespacho.text = ceros(tipodespacho)
cotiza01.GRID1.Cell(cotiza01.GRID1.ActiveCell.Row, 11).text = tipodespacho.text + " " + lblnombredespacho.Caption

Unload Me
Exit For
Else
lblnombredespacho.Caption = ""
End If
Next K

End If


End Sub
