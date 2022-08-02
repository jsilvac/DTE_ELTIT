VERSION 5.00
Object = "{4777436C-EB8C-4596-98A8-EBCF98683969}#1.0#0"; "FlexCell.ocx"
Object = "{ADBBAED6-B16F-40DC-80DF-B44910CBA76C}#1.0#0"; "Frame-Xp.ocx"
Begin VB.Form maestro04 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Form1"
   ClientHeight    =   5700
   ClientLeft      =   45
   ClientTop       =   255
   ClientWidth     =   7815
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5700
   ScaleWidth      =   7815
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin XPFrame.FrameXp FrameXp1 
      Height          =   5775
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   10186
      BackColor       =   16761024
      CaptionEstilo3D =   1
      BackColor       =   16761024
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
      Begin VB.TextBox pivote 
         Height          =   285
         Left            =   480
         TabIndex        =   7
         Top             =   5160
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FF8080&
         Caption         =   "GUARDAR"
         BeginProperty Font 
            Name            =   "Tahoma"
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
         TabIndex        =   3
         Top             =   4320
         Width           =   1095
      End
      Begin FlexCell.Grid Grid1 
         Height          =   3375
         Left            =   120
         TabIndex        =   2
         Top             =   840
         Width           =   7575
         _ExtentX        =   13361
         _ExtentY        =   5953
         Cols            =   5
         DefaultFontSize =   8.25
         Rows            =   12
      End
      Begin VB.ComboBox comboa�o 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   360
         Width           =   5415
      End
      Begin XPFrame.FrameXp FrameQuickMenu 
         Height          =   615
         Left            =   4440
         TabIndex        =   4
         Top             =   5040
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   1085
         BackColor       =   16744576
         Caption         =   " Mis Datos"
         BackColor       =   16744576
         BordeColor      =   4194304
         ColorBarraArriba=   4194304
         ColorBarraAbajo =   4194304
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Begin VB.CommandButton botonmisaccesos 
            Caption         =   "Permisos Modulo"
            Height          =   255
            Left            =   120
            TabIndex        =   6
            Top             =   280
            Width           =   1455
         End
         Begin VB.CommandButton botonmisfavoritos 
            Caption         =   "Mis Favoritos"
            Height          =   255
            Left            =   1800
            TabIndex        =   5
            Top             =   280
            Width           =   1335
         End
      End
   End
End
Attribute VB_Name = "maestro04"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub CARGAGRILLA(row, col)
    Dim FORMATOGRILLA(10, 10) As String
    Rem DATOS DE LA COLUMNA
    FORMATOGRILLA(1, 1) = "MES"
    FORMATOGRILLA(1, 2) = "% DE REAJUSTE"
    FORMATOGRILLA(1, 3) = "FACTOR"

    Rem LARGO DE LOS DATOS
    
    FORMATOGRILLA(2, 1) = "10"
    FORMATOGRILLA(2, 2) = "10"
    FORMATOGRILLA(2, 3) = "10"
    Rem TIPO DE DATOS C=CEROS,S=STRING,N=NUMERICO,D=DATE,
    FORMATOGRILLA(3, 1) = "S"
    FORMATOGRILLA(3, 2) = "N"
    FORMATOGRILLA(3, 3) = "N"
    Rem FORMATO GRILLA
    FORMATOGRILLA(4, 1) = ""
    FORMATOGRILLA(4, 2) = ""
    FORMATOGRILLA(4, 3) = ""
    Rem LOCCKED
    FORMATOGRILLA(5, 1) = "TRUE"
    FORMATOGRILLA(5, 2) = "FALSE"
    FORMATOGRILLA(5, 3) = "FALSE"
    Rem VALOR MINIMO
    FORMATOGRILLA(6, 1) = ""
    FORMATOGRILLA(6, 2) = ""
    FORMATOGRILLA(6, 3) = ""
    Rem VALOR MAXIMO
    FORMATOGRILLA(7, 1) = ""
    FORMATOGRILLA(7, 2) = ""
    FORMATOGRILLA(7, 3) = ""
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
    Grid1.BackColorFixed = RGB(90, 158, 214)
    Grid1.BackColorFixedSel = RGB(110, 180, 230)
    Grid1.BackColorBkg = RGB(90, 158, 214)
    Grid1.BackColorScrollBar = RGB(231, 235, 247)
    Grid1.BackColor1 = RGB(231, 235, 247)
    Grid1.BackColor2 = RGB(239, 243, 255)
    Grid1.GridColor = RGB(148, 190, 231)
    Grid1.Column(0).Width = 0
    For k = 1 To col - 1
        Grid1.Cell(0, k).text = FORMATOGRILLA(1, k)
        Grid1.Column(k).Width = Val(FORMATOGRILLA(2, k)) * 10.5
        Grid1.Column(k).MaxLength = Val(FORMATOGRILLA(2, k))
        Grid1.Column(k).FormatString = FORMATOGRILLA(4, k)
        Grid1.Column(k).Locked = FORMATOGRILLA(5, k)
        If FORMATOGRILLA(3, k) = "N" Then Grid1.Column(k).Alignment = cellRightCenter
        If FORMATOGRILLA(3, k) = "D" Then Grid1.Column(k).CellType = cellCalendar
         
    Next k
End Sub


Private Sub comboa�o_Click()
    Call leeripc(COMBOA�O.text)
End Sub

Private Sub Form_Load()
    Call CARGAGRILLA(1, 4)
    For k = Val(Format(fechasistema, "yyyy")) - 5 To Val(Format(fechasistema, "yyyy"))
        COMBOA�O.AddItem k
        COMBOA�O.text = COMBOA�O.List(0)
    Next k
    COMBOA�O.Refresh
End Sub
Sub leeripc(a�o)
Dim csql As New rdoQuery
Dim resultados As rdoResultset
Dim k As Double

Set csql.ActiveConnection = conta

csql.sql = "select mes,porcentaje from ipc where a�o='" & a�o & "' order by mes"
csql.Execute
Grid1.Rows = 14
If csql.RowsAffected > 0 Then
    Set resultados = csql.OpenResultset
    While Not resultados.EOF
    If resultados(0) <> "00" Then
        Grid1.Cell(resultados(0), 1).text = MonthName(resultados(0))
        Grid1.Cell(resultados(0), 2).text = resultados(1)
        Grid1.Cell(resultados(0), 3).text = (resultados(1) / 100) + 1
    Else
        Grid1.Cell(13, 1).text = "ANTERIOR"
        Grid1.Cell(13, 2).text = resultados(1)
        Grid1.Cell(13, 3).text = (resultados(1) / 100) + 1
    End If
    resultados.MoveNext
    Wend
Else
    For k = 1 To 13
        If k < 13 Then
            Grid1.Cell(k, 1).text = MonthName(k)
        End If
        Grid1.Cell(k, 2).text = 0
        Grid1.Cell(k, 3).text = 0
    Next k
    
End If
   
End Sub

Private Sub Grid1_KeyPress(KeyAscii As Integer)
     pivote.text = Grid1.ActiveCell.text
    If pivote.text <> "" Then
                KeyAscii = esNumeroDecimal(pivote, KeyAscii)
    End If
 

End Sub

Private Sub botonmisaccesos_Click()
programafiltro = Me.Caption
misaccesos.Show
End Sub


Private Sub botonmisfavoritos_Click()
Call AgregaFavoritos(USUARIOSISTEMA, App.EXEName, Me.Name, Me.Caption)

End Sub
Private Sub Command1_Click()
    Dim i As Integer
    
 For k = 1 To Grid1.Rows - 1
    i = k
    With Grid1
        Call grabar(i, COMBOA�O.text, .Cell(k, 2).text, .Cell(k, 3).text)
    End With
    Next k
leeripc (COMBOA�O.text)
End Sub

Sub grabar(ByRef dato0, dato1, dato2, dato3)
    campos(0, 0) = "mes"
    campos(1, 0) = "a�o"
    campos(2, 0) = "porcentaje"
    campos(3, 0) = "porcentaje_dolar"
    campos(4, 0) = ""
    
    If dato0 = "13" Then dato0 = "00"
    
    campos(0, 1) = Format(dato0, "00")
    campos(1, 1) = dato1
    campos(2, 1) = Replace(dato2, ",", ".")
    campos(3, 1) = Replace(dato3, ",", ".")
    campos(4, 1) = dato4
    
    campos(0, 2) = "ipc"
    opcion = 4
    condicion = "a�o='" & COMBOA�O.text & "'" & " and mes = '" & Format(dato0, "00") & "'"
    sqlconta.response = campos
    Set sqlconta.conexion = conta
    Call sqlconta.sqlconta(opcion, condicion)
   
    
   
    opcion = 2
    condicion = ""
    sqlconta.response = campos
    Set sqlconta.conexion = conta
    Call sqlconta.sqlconta(opcion, condicion)
    status = sqlconta.status
   
End Sub


