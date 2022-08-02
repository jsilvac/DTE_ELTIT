VERSION 5.00
Object = "{4777436C-EB8C-4596-98A8-EBCF98683969}#1.0#0"; "FlexCell.ocx"
Object = "{ADBBAED6-B16F-40DC-80DF-B44910CBA76C}#1.0#0"; "Frame-Xp.ocx"
Begin VB.Form permisosUsuario 
   BackColor       =   &H00008000&
   BorderStyle     =   0  'None
   Caption         =   "Configurar Puesto de Trabajo"
   ClientHeight    =   7905
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8655
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7905
   ScaleWidth      =   8655
   ShowInTaskbar   =   0   'False
   Begin XPFrame.FrameXp FrameXp1 
      Height          =   7515
      Left            =   180
      TabIndex        =   2
      Top             =   180
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   13256
      BackColor       =   12648384
      Caption         =   " Permisos de Usuario"
      CaptionEstilo3D =   1
      BackColor       =   12648384
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
      Begin FlexCell.Grid Lista 
         Height          =   6255
         Left            =   180
         TabIndex        =   1
         Top             =   1020
         Width           =   7935
         _ExtentX        =   13996
         _ExtentY        =   11033
         Cols            =   5
         DefaultFontSize =   8.25
         Rows            =   30
      End
      Begin VB.TextBox dato1 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1380
         MaxLength       =   18
         TabIndex        =   0
         Top             =   600
         Width           =   1875
      End
      Begin XPFrame.FrameXp frmCerrar 
         Height          =   330
         Left            =   7920
         TabIndex        =   3
         Top             =   30
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   582
         BackColor       =   49344
         Caption         =   "X"
         CaptionEstilo3D =   1
         BackColor       =   49344
         ColorBarraArriba=   32896
         ColorBarraAbajo =   12648447
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
      End
      Begin VB.Label lblNombre 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   3360
         TabIndex        =   5
         Top             =   600
         Width           =   4695
      End
      Begin VB.Label lbl1 
         Appearance      =   0  'Flat
         BackColor       =   &H00008080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Usuario"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   180
         TabIndex        =   4
         Top             =   600
         Width           =   1095
      End
   End
End
Attribute VB_Name = "permisosUsuario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Private CAMPOS(10, 3) As String
    Private existe As Boolean

'============================================================
'MANEJO DE LOS CONTOLES
'============================================================
    '========================================================
    'GotFocus
    '========================================================
    Private Sub dato1_GotFocus()
        Call VerificarCajas(Me, dato1)
        Call selecciona(dato1)
    End Sub
    '========================================================
    'GotFocus
    '========================================================
    
    '========================================================
    'KeyDown
    '========================================================
    Private Sub dato1_KeyDown(KeyCode As Integer, Shift As Integer)
        Call Flechas(KeyCode, dato1)
    End Sub
    '========================================================
    'KeyDown
    '========================================================
    
    '========================================================
    'KeyPress
    '========================================================
    Private Sub dato1_KeyPress(KeyAscii As Integer)
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If KeyAscii = 13 And dato1.text <> "" Then
            lista.AutoRedraw = False
            Call cargaMenus
            lblnombre.Caption = leerUsuario
            If lblnombre.Caption = "" Then
                Call selecciona(dato1)
            Else
                Call cargarPermisos
                SendKeys "{Tab}"
            End If
        End If
    End Sub
    '========================================================
    'KeyPress
    '========================================================
'============================================================
'MANEJO DE LOS CONTOLES
'============================================================

    Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
        If KeyCode = 27 Then
            Unload Me
        End If
        If KeyCode = 38 Then
            If Screen.ActiveForm.ActiveControl.Name = "dato1" Then
                Unload Me
            End If
        End If
    End Sub

    Private Sub Form_Load()
        existe = False
        Call Centrar(Me)
        Call CARGAGRILLA(1, 7)
        Call cargaMenus
    End Sub

    Private Sub frmCerrar_BarClick()
        Call cambiaColor(frmCerrar)
        frmCerrar.CaptionEstilo3D = Inserted
        Unload Me
    End Sub
    
    Private Sub frmCerrar_BarMouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
        Call cambiaColor(frmCerrar)
        frmCerrar.CaptionEstilo3D = Raised
    End Sub

'=============================================================================
'LEER USUARIO
'=============================================================================
    Private Function leerUsuario() As String
        
        Dim op As Integer
        Set sql = New sqlventas.sqlventa
        CAMPOS(0, 0) = "nombre"
        CAMPOS(1, 0) = ""
        
        CAMPOS(0, 2) = "g_usuarios"
        
        condicion = "usuario = '" & dato1.text & "'"
        op = 5
        sql.response = CAMPOS
        Set sql.conexion = gestion
        Call sql.sqlventas(op, condicion)
        If sql.Status = 0 Then
            existe = True
            leerUsuario = sql.response(0, 3)
        Else
            existe = False
            leerUsuario = ""
        End If
    End Function
'=============================================================================
'LEER USUARIO
'=============================================================================

'****************************************************************************
'Formato de la Grilla
'****************************************************************************
    Private Sub CARGAGRILLA(ByVal row As Integer, ByVal col As Integer)
        Dim i As Integer
        Dim formatogrilla(10, 6) As String
        Rem DATOS DE LA COLUMNA
        formatogrilla(1, 1) = ""
        formatogrilla(1, 2) = "NOMBRE MENU"
        formatogrilla(1, 3) = "INGRESA"
        formatogrilla(1, 4) = "MODIFICA"
        formatogrilla(1, 5) = "ELIMINA"
        formatogrilla(1, 6) = "TODAS"
        
        Rem LARGO DE LOS DATOS
        formatogrilla(2, 1) = "1"
        formatogrilla(2, 2) = "50"
        formatogrilla(2, 3) = "1"
        formatogrilla(2, 4) = "1"
        formatogrilla(2, 5) = "1"
        formatogrilla(2, 6) = "1"
        
        Rem TIPO DE DATOS C=CEROS,S=STRING,N=NUMERICO,D=DATE,
        formatogrilla(3, 1) = "C"
        formatogrilla(3, 2) = "S"
        formatogrilla(3, 3) = "C"
        formatogrilla(3, 4) = "C"
        formatogrilla(3, 5) = "C"
        formatogrilla(3, 6) = "C"
        
        Rem FORMATO GRILLA
        formatogrilla(4, 1) = ""
        formatogrilla(4, 2) = ""
        formatogrilla(4, 3) = ""
        formatogrilla(4, 4) = ""
        formatogrilla(4, 5) = ""
        formatogrilla(4, 6) = ""
        
        Rem LOCCKED
        formatogrilla(5, 1) = "TRUE"
        formatogrilla(5, 2) = "TRUE"
        formatogrilla(5, 3) = "TRUE"
        formatogrilla(5, 4) = "TRUE"
        formatogrilla(5, 5) = "TRUE"
        formatogrilla(5, 6) = "TRUE"
        
        Rem ANCHO
        formatogrilla(8, 1) = "1.5"
        formatogrilla(8, 2) = "20"
        formatogrilla(8, 3) = "8"
        formatogrilla(8, 4) = "8"
        formatogrilla(8, 5) = "8"
        formatogrilla(8, 6) = "8"
            
        lista.Cols = col
        lista.Rows = row
        lista.AllowUserResizing = False
        lista.DisplayFocusRect = False
        lista.ExtendLastCol = False
        lista.BoldFixedCell = False
        lista.DisplayRowIndex = True
        lista.DrawMode = cellOwnerDraw
        lista.Appearance = Flat
        lista.ScrollBarStyle = Flat
        lista.FixedRowColStyle = Flat
        lista.BackColorFixed = RGB(90, 214, 158)
        lista.BackColorFixedSel = RGB(110, 230, 180)
        lista.BackColorBkg = RGB(90, 214, 158)
        lista.BackColorScrollBar = RGB(231, 247, 235)
        lista.BackColor1 = RGB(231, 247, 235)
        lista.BackColor2 = RGB(239, 255, 243)
        lista.GridColor = RGB(148, 231, 190)
        
        lista.Column(0).Width = 0
        For i = 1 To col - 1
            lista.Cell(0, i).text = formatogrilla(1, i)
            lista.Column(i).Width = Val(formatogrilla(8, i)) * (lista.Cell(0, i).Font.Size + 1.25)
            lista.Column(i).MaxLength = Val(formatogrilla(2, i))
            lista.Column(i).FormatString = formatogrilla(4, i)
            lista.Column(i).Locked = formatogrilla(5, i)
            If formatogrilla(3, i) = "N" Then
                lista.Column(i).Alignment = cellRightCenter
            End If
            If formatogrilla(3, i) = "S" Then
                lista.Column(i).Alignment = cellLeftCenter
            End If
            If formatogrilla(3, i) = "C" Then
                lista.Column(i).Alignment = cellCenterCenter
            End If
        Next i
        lista.Range(0, 1, 0, lista.Cols - 1).Alignment = cellCenterCenter
    End Sub
'****************************************************************************
'Formato de la Grilla
'****************************************************************************

    Private Sub cargaMenus()
        Dim obj As Object
        Dim mnu As MENU
        Dim cadena As String
        Dim cad1 As String
        Dim cad2 As String
        Dim espacios As String
        Dim NIVEL As Integer
        lista.Rows = 1
        lista.AutoRedraw = False
        espacios = ""
        NIVEL = 0
        For Each obj In Principal
            If TypeOf obj Is MENU Then
                Set mnu = obj
                cadena = Replace(mnu.Caption, "&", "")
                cad1 = Right(cadena, 1)
                cad2 = UCase(cad1)
                If cad1 <> cad2 Then
                    espacios = vbTab & "     "
                Else
                    espacios = "-" & vbTab
                    If lista.Rows > 1 And cadena <> "-" And cadena <> "VENTANAS" And cadena <> "SALIR" And cadena <> "Cerrar Sesión..." And cadena <> "Salir" And cadena <> "Acerca de..." Then
                        lista.AddItem "", True
                    End If
                End If
                If cadena <> "-" And cadena <> "VENTANAS" And cadena <> "SALIR" And cadena <> "Cerrar Sesión..." And cadena <> "Salir" And cadena <> "Acerca de..." Then
                    lista.AddItem espacios & cadena & vbTab & "N" & vbTab & "N" & vbTab & "N" & vbTab & "N", True
                    If espacios = "-" & vbTab Then
                        lista.Range(lista.Rows - 1, 1, lista.Rows - 1, lista.Cols - 1).FontBold = True
                    End If
                End If
            End If
        Next obj
        lista.AutoRedraw = True
        lista.Refresh
    End Sub

    Private Sub cargarPermisos()
        Dim csql As rdoQuery
        Dim resultado As rdoResultset
        Dim i As Integer
        Set csql = New rdoQuery
        Set csql.ActiveConnection = gestion
        
        csql.sql = "SELECT glosa, ingresa, modifica, elimina, todas "
        csql.sql = csql.sql & "FROM g_permisos "
        csql.sql = csql.sql & "WHERE usuario = '" & dato1.text & "' "
        csql.Execute
        
        If csql.RowsAffected > 0 Then
            lista.AutoRedraw = False
            Set resultado = csql.OpenResultset
            While Not resultado.EOF
                For i = 1 To lista.Rows - 1
                    If InStr(1, lista.Cell(i, 2).text, resultado("glosa"), vbBinaryCompare) > 0 Then
                        lista.Cell(i, 3).text = resultado("ingresa")
                        lista.Cell(i, 4).text = resultado("modifica")
                        lista.Cell(i, 5).text = resultado("elimina")
                        lista.Cell(i, 6).text = resultado("todas")
                        Exit For
                    End If
                Next i
                resultado.MoveNext
            Wend
            lista.AutoRedraw = True
            lista.Refresh
        End If
        Set resultado = Nothing
        csql.Close
        Set csql = Nothing
    End Sub
    
Private Sub Lista_Click()
    Dim col As Integer
    Dim fil As Integer
    Dim i As Integer
    Dim signo As String
    col = lista.ActiveCell.col
    fil = lista.ActiveCell.row
    If col = 1 Then
        If lista.Cell(fil, col).text = "-" Then
            lista.AutoRedraw = False
            For i = fil + 1 To lista.Rows - 1
                If lista.Cell(i, col + 1).text <> "" Then
                    lista.RowHeight(i) = 0
                Else
                    Exit For
                End If
            Next i
            signo = "+"
            lista.AutoRedraw = True
            lista.Refresh
        End If
        If lista.Cell(fil, col).text = "+" Then
            lista.AutoRedraw = False
            For i = fil + 1 To lista.Rows - 1
                If lista.Cell(i, col + 1).text <> "" Then
                    lista.RowHeight(i) = lista.DefaultRowHeight
                Else
                    Exit For
                End If
            Next i
            signo = "-"
            lista.AutoRedraw = True
            lista.Refresh
        End If
        lista.Cell(fil, col).text = signo
    End If
End Sub

Private Sub Lista_DblClick()
    Dim col As Integer
    Dim fil As Integer
    Dim i As Integer
    Dim estados(0 To 4) As String
    col = lista.ActiveCell.col
    fil = lista.ActiveCell.row
    If existe = True Then
        If col >= 3 Then
            If col = lista.Cols - 1 Then
                If lista.Cell(fil, col).text = "S" Then
                    For i = 3 To col
                        lista.Cell(fil, i).text = "N"
                        estados(i - 2) = "N"
                    Next i
                Else
                    For i = 3 To col
                        lista.Cell(fil, i).text = "S"
                        estados(i - 2) = "S"
                    Next i
                End If
            Else
                If lista.Cell(fil, col).text = "S" Then
                    lista.Cell(fil, col).text = "N"
                    estados(col - 2) = "N"
                Else
                    lista.Cell(fil, col).text = "S"
                    estados(col - 2) = "S"
                End If
            End If
            estados(0) = Trim(lista.Cell(fil, 2).text)
        End If
        Call modificaPermiso(estados)
    End If
End Sub

    Private Sub modificaPermiso(ByVal estados As Variant)
        Dim csql As rdoQuery
        
        Set csql = New rdoQuery
        Set csql.ActiveConnection = gestion
        
        csql.sql = "INSERT INTO g_permisos (usuario, glosa, ingresa, modifica, elimina, todas) "
        csql.sql = csql.sql & "VALUES('" & dato1.text & "', '" & estados(0) & "', '" & estados(1) & "', '" & estados(2) & "', '" & estados(3) & "', '" & estados(4) & "') "
        csql.sql = csql.sql & "ON DUPLICATE KEY UPDATE ingresa = '" & estados(1) & "', modifica = '" & estados(2) & "', elimina = '" & estados(3) & "', todas = '" & estados(4) & "' "
        
        csql.Execute
        csql.Close
        Set csql = Nothing
    End Sub
