VERSION 5.00
Object = "{4777436C-EB8C-4596-98A8-EBCF98683969}#1.0#0"; "FlexCell.ocx"
Object = "{ADBBAED6-B16F-40DC-80DF-B44910CBA76C}#1.0#0"; "Frame-Xp.ocx"
Begin VB.Form misaccesos 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Accesos "
   ClientHeight    =   5955
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11835
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5955
   ScaleWidth      =   11835
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin XPFrame.FrameXp impuestos 
      Height          =   5970
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   10530
      BackColor       =   16761024
      Caption         =   ""
      CaptionEstilo3D =   1
      BackColor       =   16761024
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.CommandButton Command2 
         Caption         =   "Con Acceso"
         Height          =   375
         Left            =   240
         TabIndex        =   4
         Top             =   240
         Width           =   1695
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Sin acceso"
         Height          =   375
         Left            =   9840
         TabIndex        =   3
         Top             =   240
         Width           =   1695
      End
      Begin VB.TextBox filtro 
         Alignment       =   2  'Center
         Height          =   375
         Left            =   2280
         TabIndex        =   2
         Top             =   240
         Width           =   7335
      End
      Begin FlexCell.Grid Grid1 
         Height          =   5250
         Left            =   45
         TabIndex        =   1
         Top             =   630
         Width           =   11610
         _ExtentX        =   20479
         _ExtentY        =   9260
         Cols            =   5
         DefaultFontName =   "Arial"
         DefaultFontSize =   8.25
         ExtendLastCol   =   -1  'True
         Rows            =   30
      End
   End
End
Attribute VB_Name = "misaccesos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Call leepermisos(filtro.text, 2)
End Sub

Private Sub COMMAND2_Click()
Call leepermisos(filtro.text, 1)
End Sub

Private Sub Form_Activate()
If Verifica_Permiso(programafiltro, "autoriza") = False Then

MsgBox "ATRIBUTOS INSUFICIENTES PARA VER ESTE MODULO"
Unload Me
Else

Call leepermisos(filtro.text, 1)
End If

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
Unload Me
End If
End Sub

Private Sub Form_Load()
'FormatoImpuestos
'CargaImpuestos
filtro.text = programafiltro
Call CARGAGRILLAPERMISOS(1, 8)
End Sub

Function leepermisos(programa As String, tipo As Integer) As Double

    Dim resultados As rdoResultset
    Dim csql As New rdoQuery
    Dim conta As Integer
    Set csql.ActiveConnection = contadb
    If tipo = 1 Then
    csql.sql = "SELECT * FROM " & clientesistema & "conta.segu_permisos WHERE programa ='" & programa & "' AND ingresa+agrega+modifica+elimina+autoriza+todas > 0 ORDER BY usuario "
    Else
    csql.sql = "SELECT usu.usuario,'','',0,0,0,0,0,'',0 FROM eltit_auditoria.segu_usuarios AS usu LEFT JOIN eltit_conta.segu_permisos AS per ON (usu.usuario=per.usuario) WHERE per.programa IS NULL OR ingresa+agrega+modifica+elimina+autoriza+todas = 0"
    End If
    csql.Execute
    conta = 0
    Grid1.Rows = 1
    If csql.RowsAffected > 0 Then
        Set resultados = csql.OpenResultset
        While Not resultados.EOF
        Grid1.Rows = Grid1.Rows + 1
        Grid1.Cell(Grid1.Rows - 1, 1).text = resultados(0)
        Grid1.Cell(Grid1.Rows - 1, 2).text = resultados(3)
        Grid1.Cell(Grid1.Rows - 1, 3).text = resultados(4)
        Grid1.Cell(Grid1.Rows - 1, 4).text = resultados(5)
        Grid1.Cell(Grid1.Rows - 1, 5).text = resultados(6)
        Grid1.Cell(Grid1.Rows - 1, 6).text = resultados(8)
        Grid1.Cell(Grid1.Rows - 1, 7).text = resultados(7)
        
        resultados.MoveNext
        Wend
        resultados.Close
        Set resultados = Nothing
    End If
    
End Function



Private Sub Grid2_LeaveCell(ByVal row As Long, ByVal col As Long, NewRow As Long, NewCol As Long, Cancel As Boolean)
   
    If NewCol <> 3 Then NewCol = 3

End Sub

Sub CARGAGRILLAPERMISOS(row, col)
    Dim FORMATOGRILLA(10, 12)
    Rem DATOS DE LA COLUMNA
    
    Rem LARGO DE LOS DATOS
    
    FORMATOGRILLA(2, 1) = "30"
    FORMATOGRILLA(2, 2) = "2"
    
    Rem TIPO DE DATOS C=CEROS,S=STRING,N=NUMERICO,D=DATE,
    FORMATOGRILLA(3, 1) = "S"
    FORMATOGRILLA(3, 2) = "S"
    
    Rem FORMATO GRILLA
    FORMATOGRILLA(4, 1) = ""
    FORMATOGRILLA(4, 2) = ""
    Rem LOCCKED
    FORMATOGRILLA(5, 1) = "FALSE"
    FORMATOGRILLA(5, 2) = "FALSE"
    Grid1.Cols = col
    Grid1.Rows = row
    Grid1.AllowUserResizing = False
    Grid1.DisplayFocusRect = False
    Grid1.ExtendLastCol = True
    Grid1.BoldFixedCell = False
    Grid1.DrawMode = cellOwnerDraw
    Grid1.Appearance = Flat
    Grid1.ScrollBarStyle = Flat
    Grid1.FixedRowColStyle = Flat
    Grid1.Column(0).Width = 0
    Grid1.Column(1).Width = 280
    Grid1.Column(1).Locked = True
    For k = 2 To Grid1.Cols - 1
    Grid1.Column(k).Width = 75
    Grid1.Column(k).CellType = cellCheckBox
   Next k
   Grid1.Cell(0, 1).text = "USUARIO"
   Grid1.Cell(0, 2).text = "INGRESAR"
   Grid1.Cell(0, 3).text = "AGREGAR"
   Grid1.Cell(0, 4).text = "MODIFICAR"
   Grid1.Cell(0, 5).text = "ELIMINAR"
   Grid1.Cell(0, 6).text = "SUPERVISOR"
   Grid1.Cell(0, 7).text = "AUTORIZAR"
    
End Sub

Sub GrabarPermiso(nombreprograma As String, Usuario As String)
    
    campos(0, 0) = "usuario"
    campos(1, 0) = "empresa"
    campos(2, 0) = "programa"
    campos(3, 0) = "ingresa"
    campos(4, 0) = "modifica"
    campos(5, 0) = "elimina"
    campos(6, 0) = "agrega"
    campos(7, 0) = "todas"
    campos(8, 0) = "menu"
    campos(9, 0) = "autoriza"
    campos(10, 0) = ""
  
    campos(0, 1) = Usuario
    campos(1, 1) = ""
    campos(2, 1) = nombreprograma
    campos(3, 1) = Grid1.Cell(Grid1.ActiveCell.row, 2).text 'ingresa
    campos(4, 1) = Grid1.Cell(Grid1.ActiveCell.row, 4).text 'modificar
    campos(5, 1) = Grid1.Cell(Grid1.ActiveCell.row, 5).text 'eliminar
    campos(6, 1) = Grid1.Cell(Grid1.ActiveCell.row, 3).text 'agregar
    campos(7, 1) = Grid1.Cell(Grid1.ActiveCell.row, 6).text 'supervisor
    campos(8, 1) = ""
    campos(9, 1) = Grid1.Cell(Grid1.ActiveCell.row, 7).text 'autoriza
    
    campos(0, 2) = "segu_permisos"
    condicion = "usuario=" + "'" + Usuario + "' and programa='" + nombreprograma + "'"
    
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    If ELIMINA = False Then
    op = 5
    Call sqlconta.sqlconta(op, condicion)
  
  
  If sqlconta.status = 4 Then
  op = 2
  condicion = ""
  Else
  op = 3
  End If
  Call sqlconta.sqlconta(op, condicion)
  Else
  op = 4
  Call sqlconta.sqlconta(op, condicion)
End If
     
End Sub

Private Sub Grid1_Click()
If Grid1.Cell(Grid1.ActiveCell.row, 1).text <> "" Then
Call GrabarPermiso(filtro.text, Grid1.Cell(Grid1.ActiveCell.row, 1).text)
End If

End Sub
