VERSION 5.00
Object = "{4777436C-EB8C-4596-98A8-EBCF98683969}#1.0#0"; "FlexCell.ocx"
Object = "{ADBBAED6-B16F-40DC-80DF-B44910CBA76C}#1.0#0"; "Frame-Xp.ocx"
Begin VB.Form control08 
   Caption         =   "Form1"
   ClientHeight    =   5775
   ClientLeft      =   60
   ClientTop       =   360
   ClientWidth     =   11160
   LinkTopic       =   "Form1"
   ScaleHeight     =   5775
   ScaleWidth      =   11160
   StartUpPosition =   3  'Windows Default
   Begin XPFrame.FrameXp FrameXp1 
      Height          =   5775
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11055
      _ExtentX        =   19500
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
      Begin VB.CheckBox chk1 
         BackColor       =   &H00FFC0C0&
         Caption         =   "TODOS"
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
         Left            =   7680
         TabIndex        =   9
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton Command2 
         Caption         =   "BUSCAR"
         Height          =   375
         Left            =   4800
         TabIndex        =   8
         Top             =   360
         Width           =   2175
      End
      Begin VB.ComboBox comboimpuesto 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   360
         Width           =   4335
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
         Left            =   360
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   5280
         Width           =   1095
      End
      Begin VB.TextBox pivote 
         Height          =   285
         Left            =   360
         TabIndex        =   1
         Top             =   6600
         Visible         =   0   'False
         Width           =   615
      End
      Begin FlexCell.Grid Grid1 
         Height          =   4215
         Left            =   120
         TabIndex        =   3
         Top             =   840
         Width           =   10815
         _ExtentX        =   19076
         _ExtentY        =   7435
         Cols            =   5
         DefaultFontSize =   8.25
         Rows            =   12
      End
      Begin XPFrame.FrameXp FrameQuickMenu 
         Height          =   615
         Left            =   7800
         TabIndex        =   5
         Top             =   5160
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
         Begin VB.CommandButton botonmisfavoritos 
            Caption         =   "Mis Favoritos"
            Height          =   255
            Left            =   1800
            TabIndex        =   7
            Top             =   280
            Width           =   1335
         End
         Begin VB.CommandButton botonmisaccesos 
            Caption         =   "Permisos Modulo"
            Height          =   255
            Left            =   120
            TabIndex        =   6
            Top             =   280
            Width           =   1455
         End
      End
   End
End
Attribute VB_Name = "control08"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Dim k As Double
    For k = 1 To Grid1.Rows - 1
        Call grabar(Mid(Grid1.Cell(k, 1).text, 1, 5), Mid(Grid1.Cell(k, 2).text, 1, 5), Mid(Grid1.Cell(k, 3).text, 1, 5), Grid1.Cell(k, 4).text)
    Next k
    Call buscalineas(Mid(comboimpuesto.text, 1, 5))
End Sub

Private Sub COMMAND2_Click()
    Call buscalineas(Mid(comboimpuesto.text, 1, 5))
End Sub
Sub CARGAGRILLA(row, col)
    Dim FORMATOGRILLA(10, 10) As String
    Rem DATOS DE LA COLUMNA
    FORMATOGRILLA(1, 1) = "SECCION"
    FORMATOGRILLA(1, 2) = "DEPTO"
    FORMATOGRILLA(1, 3) = "LINEA"
    FORMATOGRILLA(1, 4) = "CODIGO SII"

    Rem LARGO DE LOS DATOS
    
    FORMATOGRILLA(2, 1) = "17"
    FORMATOGRILLA(2, 2) = "17"
    FORMATOGRILLA(2, 3) = "17"
    FORMATOGRILLA(2, 4) = "10"
 
    
    Rem TIPO DE DATOS C=CEROS,S=STRING,N=NUMERICO,D=DATE,
    FORMATOGRILLA(3, 1) = "S"
    FORMATOGRILLA(3, 2) = "S"
    FORMATOGRILLA(3, 3) = "S"
    FORMATOGRILLA(3, 4) = "N"
     
    
    
    
    Rem FORMATO GRILLA
    FORMATOGRILLA(4, 1) = ""
    FORMATOGRILLA(4, 2) = ""
    FORMATOGRILLA(4, 3) = ""
    Rem LOCCKED
    FORMATOGRILLA(5, 1) = "TRUE"
    FORMATOGRILLA(5, 2) = "TRUE"
    FORMATOGRILLA(5, 3) = "TRUE"
    FORMATOGRILLA(5, 4) = "FALSE"
   
    
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

Private Sub Form_Load()
    Call CARGAGRILLA(1, 5)
    Call cargacomboimpuesto
End Sub
Sub cargacomboimpuesto()
    Dim csql As New rdoQuery
    Dim resultados As rdoResultset
    Set csql.ActiveConnection = contadb
    csql.sql = "select codigo,nombre from " & clientesistema & "gestion.g_maestroimpuestos "
    csql.sql = csql.sql & " where codigo between '00004' and '00005' "
    csql.Execute
   
    If csql.RowsAffected > 0 Then
        Set resultados = csql.OpenResultset
        While Not resultados.EOF
             comboimpuesto.AddItem resultados(0) & " - " & resultados(1)
            resultados.MoveNext
            
        Wend
    End If
    csql.Close
    Set csql = Nothing
   
    
  
        comboimpuesto.text = comboimpuesto.List(0)
        comboimpuesto.Refresh
    
End Sub
Sub buscalineas(codigoimpuesto)
    Dim csql As New rdoQuery
    Dim resultados As rdoResultset
    Set csql.ActiveConnection = contadb
    csql.sql = "SELECT ml.codigoseccion,ml.codigodepto,ml.codigolinea,ml.nombre,ml.codigosii "
    csql.sql = csql.sql & " FROM " & clientesistema & "gestion" & rubro & ".r_maestrolineas_" & rubro & " AS ml "
    csql.sql = csql.sql & "INNER JOIN " & clientesistema & "gestion" & rubro & ".r_maestroproductos_fijo_" & rubro & " AS mpf "
    csql.sql = csql.sql & "ON ml.codigoseccion=mpf.codigoseccion AND ml.codigodepto=mpf.codigodepto AND ml.codigolinea=mpf.codigolinea "
    csql.sql = csql.sql & " WHERE mpf.codigoimpuesto='" & codigoimpuesto & "' "
    If chk1.Value = 0 Then
        csql.sql = csql.sql & "AND ml.codigosii='' "
    End If
    
    csql.sql = csql.sql & "GROUP BY ml.codigoseccion,ml.codigodepto,ml.codigolinea"
    csql.Execute
     Grid1.Rows = 1
    Grid1.AutoRedraw = False
    If csql.RowsAffected > 0 Then
        Set resultados = csql.OpenResultset
        While Not resultados.EOF
            Grid1.Rows = Grid1.Rows + 1
            Grid1.Cell(Grid1.Rows - 1, 1).text = resultados(0) & " - " & leerdatos(conta, clientesistema & "gestion" & rubro & ".r_maestrosecciones_" & rubro, "nombre", "codigo='" & resultados(0) & "'")
            Grid1.Cell(Grid1.Rows - 1, 2).text = resultados(1) & " - " & leerdatos(conta, clientesistema & "gestion" & rubro & ".r_maestrodepartamentos_" & rubro, "nombre", "codigoseccion='" & resultados(0) & "' and codigodepto='" & resultados(1) & "'")
            Grid1.Cell(Grid1.Rows - 1, 3).text = resultados(2) & " - " & resultados(3)
            Grid1.Cell(Grid1.Rows - 1, 4).text = resultados(4)
            resultados.MoveNext
        Wend
    End If
    csql.Close
    Set csql = Nothing
    
    
     Grid1.AutoRedraw = True
    Grid1.Refresh
    
End Sub
Sub grabar(seccion, depto, LINEA, sii)
    Dim csql As New rdoQuery
    Dim resultados As rdoResultset
    Set csql.ActiveConnection = contadb
    
    
    csql.sql = "update " & clientesistema & "gestion" & rubro & ".r_maestrolineas_" & rubro
    csql.sql = csql.sql & " set codigosii='" & sii & "' "
    csql.sql = csql.sql & " where codigoseccion='" & Format(seccion, "00000") & "' and codigodepto='" & Format(depto, "00000") & "' and codigolinea='" & Format(LINEA, "00000") & "' "
    csql.Execute
    Call sincronizadatos(csql.sql, conta, Servidor)
    csql.Close
    Set csql = Nothing
    
End Sub

Private Sub Grid1_KeyPress(KeyAscii As Integer)
    KeyAscii = esNumero(KeyAscii)
    
End Sub
