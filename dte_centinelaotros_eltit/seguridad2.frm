VERSION 5.00
Object = "{4777436C-EB8C-4596-98A8-EBCF98683969}#1.0#0"; "Flexcell.ocx"
Object = "{ADBBAED6-B16F-40DC-80DF-B44910CBA76C}#1.0#0"; "Frame-Xp.ocx"
Begin VB.Form seguridad2 
   BackColor       =   &H00FFC0C0&
   Caption         =   "MODULO DE SEGURIDAD"
   ClientHeight    =   10065
   ClientLeft      =   1260
   ClientTop       =   450
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   NegotiateMenus  =   0   'False
   ScaleHeight     =   671
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1016
   WindowState     =   2  'Maximized
   Begin XPFrame.FrameXp FrameXp2 
      Height          =   9855
      Left            =   12000
      TabIndex        =   21
      Top             =   120
      Width           =   3195
      _ExtentX        =   5636
      _ExtentY        =   17383
      BackColor       =   8454016
      Caption         =   "PROGRAMAS PROHIBIDOS"
      CaptionEstilo3D =   1
      BackColor       =   8454016
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin XPFrame.FrameXp FrameXp4 
         Height          =   4455
         Left            =   0
         TabIndex        =   26
         Top             =   5280
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   7858
         BackColor       =   8454016
         Caption         =   "PROGRAMAS USUARIO"
         CaptionEstilo3D =   1
         BackColor       =   8454016
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Begin FlexCell.Grid Grid6 
            Height          =   4125
            Left            =   0
            TabIndex        =   27
            Top             =   240
            Width           =   3090
            _ExtentX        =   5450
            _ExtentY        =   7276
            Cols            =   5
            DefaultFontSize =   8.25
            Rows            =   30
         End
      End
      Begin VB.TextBox dato1 
         Height          =   285
         Left            =   120
         TabIndex        =   24
         Top             =   4200
         Width           =   2895
      End
      Begin VB.CommandButton cmdagregarprograma 
         BackColor       =   &H0000FF00&
         Caption         =   "AGREGAR A LA LISTA"
         Enabled         =   0   'False
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
         Left            =   600
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   4560
         Width           =   2355
      End
      Begin FlexCell.Grid Grid5 
         Height          =   3165
         Left            =   45
         TabIndex        =   23
         Top             =   315
         Width           =   3090
         _ExtentX        =   5450
         _ExtentY        =   5583
         Cols            =   5
         DefaultFontSize =   8.25
         Rows            =   30
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Escriba ""ejemplo.exe"" sin comillas y presione ""AGREGAR A LA LISTA"""
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   120
         TabIndex        =   25
         Top             =   3600
         Width           =   3015
      End
   End
   Begin VB.TextBox pivote 
      Height          =   285
      Left            =   120
      TabIndex        =   20
      Top             =   120
      Visible         =   0   'False
      Width           =   375
   End
   Begin XPFrame.FrameXp FrameXp3 
      Height          =   1905
      Left            =   0
      TabIndex        =   11
      Top             =   6120
      Visible         =   0   'False
      Width           =   4875
      _ExtentX        =   8599
      _ExtentY        =   3360
      BackColor       =   8454016
      Caption         =   "MODULO DE COPIA"
      CaptionEstilo3D =   1
      BackColor       =   8454016
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.TextBox dato4 
         Appearance      =   0  'Flat
         BackColor       =   &H00E1FFFD&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   1665
         MaxLength       =   30
         TabIndex        =   15
         Tag             =   "proveedor"
         Top             =   810
         Width           =   3045
      End
      Begin VB.TextBox dato3 
         Appearance      =   0  'Flat
         BackColor       =   &H00E1FFFD&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   1665
         MaxLength       =   30
         TabIndex        =   14
         Tag             =   "proveedor"
         Top             =   450
         Width           =   3045
      End
      Begin VB.CommandButton CANCELAR 
         BackColor       =   &H0000FF00&
         Caption         =   "CANCELAR"
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
         Left            =   2565
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   1350
         Width           =   1635
      End
      Begin VB.CommandButton COPIAR 
         BackColor       =   &H0000FF00&
         Caption         =   "ACEPTAR"
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
         Left            =   315
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   1350
         Width           =   1680
      End
      Begin VB.Label lbl4 
         Appearance      =   0  'Flat
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Usuario Destino"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   135
         TabIndex        =   17
         Top             =   810
         Width           =   1395
      End
      Begin VB.Label lbl3 
         Appearance      =   0  'Flat
         BackColor       =   &H00F5C9B1&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Usuario Origen"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   135
         TabIndex        =   16
         Top             =   450
         Width           =   1425
      End
   End
   Begin XPFrame.FrameXp MENU 
      Height          =   5325
      Left            =   45
      TabIndex        =   7
      Top             =   90
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   9393
      BackColor       =   16761024
      Caption         =   "MENU"
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
      Begin VB.CommandButton MENU1 
         Appearance      =   0  'Flat
         Caption         =   "INGRESOS"
         Height          =   240
         Left            =   0
         TabIndex        =   8
         Top             =   0
         Visible         =   0   'False
         Width           =   1455
      End
      Begin FlexCell.Grid Grid1 
         Height          =   4875
         Left            =   90
         TabIndex        =   9
         Top             =   270
         Width           =   11685
         _ExtentX        =   20611
         _ExtentY        =   8599
         Cols            =   5
         DefaultFontSize =   8.25
         Rows            =   30
      End
   End
   Begin XPFrame.FrameXp FRMUSUARIO 
      Height          =   4380
      Left            =   135
      TabIndex        =   0
      Top             =   5625
      Width           =   3720
      _ExtentX        =   6562
      _ExtentY        =   7726
      BackColor       =   8454016
      Caption         =   "USUARIOS"
      CaptionEstilo3D =   1
      BackColor       =   8454016
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.CommandButton Command3 
         BackColor       =   &H0000FF00&
         Caption         =   "COPIAR PERMISOS"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1920
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   3645
         Width           =   1635
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H0000FF00&
         Caption         =   "AGREGAR USUARIO"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   3645
         Width           =   1755
      End
      Begin FlexCell.Grid Grid2 
         Height          =   3255
         Left            =   90
         TabIndex        =   1
         Top             =   315
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   5741
         Cols            =   5
         DefaultFontSize =   8.25
         Rows            =   30
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00F5C9B1&
         BackStyle       =   0  'Transparent
         Caption         =   "Supr - Eliminar Usuario"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   1920
         TabIndex        =   18
         Top             =   0
         Width           =   1740
      End
   End
   Begin XPFrame.FrameXp FrameXp1 
      Height          =   4335
      Left            =   3840
      TabIndex        =   2
      Top             =   5640
      Width           =   8115
      _ExtentX        =   14314
      _ExtentY        =   7646
      BackColor       =   8454016
      Caption         =   "DATOS USUARIOS"
      CaptionEstilo3D =   1
      BackColor       =   8454016
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.CommandButton Command4 
         BackColor       =   &H0000FF00&
         Caption         =   "MODIFICAR"
         Enabled         =   0   'False
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
         Left            =   3840
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   3780
         Width           =   2355
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H0000FF00&
         Caption         =   "GRABAR"
         Enabled         =   0   'False
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
         Left            =   480
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   3780
         Width           =   2355
      End
      Begin FlexCell.Grid Grid3 
         Height          =   3165
         Left            =   45
         TabIndex        =   3
         Top             =   315
         Width           =   3450
         _ExtentX        =   6085
         _ExtentY        =   5583
         Cols            =   5
         DefaultFontSize =   8.25
         Rows            =   30
      End
      Begin FlexCell.Grid Grid4 
         Height          =   3165
         Left            =   3600
         TabIndex        =   5
         Top             =   315
         Width           =   4410
         _ExtentX        =   7779
         _ExtentY        =   5583
         Cols            =   5
         DefaultFontSize =   8.25
         Rows            =   30
      End
   End
End
Attribute VB_Name = "seguridad2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private formatogrilla(12, 12)
Private VARIABLE As String
Private USUARIOSELECCIONADO As String
Private menuseleccion As String
Private modifo As Double
Private eli As Boolean

Private Sub CANCELAR_Click()
dato3.text = ""
dato4.text = ""
FrameXp3.Visible = False
End Sub

Private Sub Command1_Click()
If Grid3.Cell(1, 2).text <> "" Then
grabarusuario
Command1.Enabled = False
Command4.Enabled = False
LEERUSUARIOS
Else
Grid3.Cell(1, 2).SetFocus
End If
End Sub

Private Sub Command2_Click()
Grid3.Cell(1, 2).text = ""
Grid3.Cell(2, 2).text = ""
Grid3.Cell(3, 2).text = ""
Grid3.Cell(4, 2).text = ""
Grid3.Cell(5, 2).text = ""
Grid3.Cell(6, 2).text = ""
Grid3.Cell(1, 2).SetFocus
Command1.Enabled = True
End Sub

Private Sub Command3_Click()
FrameXp3.Visible = True
If Grid2.Cell(Grid2.ActiveCell.row, 1).text = "USUARIOS" Then
dato3.text = ""
Else

dato3.text = Grid2.Cell(Grid2.ActiveCell.row, 1).text
End If

dato3.SetFocus

End Sub

Private Sub Command4_Click()
Grid3.Cell(1, 2).SetFocus
MODIFI = True

Command1.Enabled = True
End Sub

Private Sub cmdagregarprograma_Click()
Dim i As Double
If dato1.text <> "" Then
    Call grabarprogramaprohibido(dato1.text)
    Call leerprogramasprohibidos(Grid5)
    Call leerprogramasprohibidos(Grid6)
    
    For i = 1 To Grid6.Rows - 1
        Call leerprogramasprohibidosusuario(USUARIOSELECCIONADO, Grid6.Cell(i, 1).text, i, Grid6)
    Next i
    
    cmdagregarprograma.Enabled = False
    dato1.text = ""
    dato1.SetFocus
    
End If
End Sub

Private Sub COPIAR_Click()
 Dim resultados2 As rdoResultset
    Dim cSql2 As New rdoQuery
    Dim saldodebe As String
    Dim saldohaber As String
    Dim lineas As Double
    
If MsgBox("ESTA SEGURO QUE DESEA COPIAR", vbOKCancel, "ADVERTENCIA") = vbOK Then
        Set cSql2.ActiveConnection = ventas
        cSql2.sql = "DELETE "
        cSql2.sql = cSql2.sql + "FROM segu_permisos "
        cSql2.sql = cSql2.sql + "where usuario='" + dato4.text + "' "
        cSql2.Execute
        Call copiarpermisos(dato3.text, dato4.text)
        dato4.text = ""
        
        dato4.SetFocus
        
End If

End Sub

 

Private Sub dato1_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(Chr(KeyAscii))
If KeyAscii = 13 And dato1.text <> "" Then
    cmdagregarprograma.Enabled = True
    cmdagregarprograma.SetFocus
End If

End Sub

 Private Sub dato3_GotFocus()
        Call cargatexto(dato3)
 End Sub
Private Sub dato3_KeyDown(KeyCode As Integer, Shift As Integer)
        Call Flechas(KeyCode, dato3)
End Sub
Private Sub dato3_KeyPress(KeyAscii As Integer)
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If KeyAscii = 13 Then
           dato4.SetFocus
        End If
End Sub

Private Sub dato4_GotFocus()
        Call cargatexto(dato4)
End Sub
Private Sub dato4_KeyDown(KeyCode As Integer, Shift As Integer)
        Call Flechas(KeyCode, dato3)
End Sub
 Private Sub dato4_KeyPress(KeyAscii As Integer)
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If KeyAscii = 13 Then
           COPIAR_Click
        End If
    End Sub

Private Sub Form_Activate()
sqlventas.audit = True
sqlventas.programaactivo = Me.Caption

End Sub

Private Sub Form_Load()
Dim K As Integer

  '==================================
    'PERMITE UNA INSTANCIA DEL SISTEMA
'    '==================================
'    Dim saveTitle$
'    If App.PrevInstance Then
'        saveTitle$ = App.Title
'        App.Title = "... duplicate instance."
'        Me.Caption = "... duplicate instance."
'        AppActivate saveTitle$
'        SendKeys "% R", True
'        End
'    End If
'    ==================================
'    PERMITE UNA INSTANCIA DEL SISTEMA
'    ==================================
''
'Close 20
'Open "c:\configu.txt" For Input As #20
'Input #20, SS
'    servidor = SS
'Input #20, SS
'
'    USUARIO = SS
'Input #20, SS
'    password = SS
'Input #20, SS
'    clientesistema = SS


'servidor = "localhost"
'USUARIO = "root"
'password = "123"

'servidor = "164.77.237.204"
'USUARIO = "prueba"
'password = ""
'If Format(fechasistema, "yyyy") > "2007" Then clientesistema = "molino2_"

    
'    basedatos = clientesistema + "conta"
'    Call Conectarconta(servidor, basedatos, USUARIO, password)
'
    cabezas = Array("codigo", "nombre")
    mensajeAyuda = "Ayuda General"
    rutaUpdate = "i:\Actualizaciones"
    'Call verificarUpdate


Call CARGAGRILLAPERMISOS(6, 8)
Call CARGAGRILLAUSUARIOS(2, 2)
Call CARGAGRILLADATOS(7, 3)
Call CARGAGRILLAEMPRESA(10, 4)
Call CARGAGRILLAPROGRAMAPROHIBIDOS(1, 2)
Call CARGAGRILLAPROGRAMASUSUARIO(1, 3)
leerempresa2
Call leerprogramasprohibidos(Grid5)
Call leerprogramasprohibidos(Grid6)
'For K = 1 To ingresos.Count
'ingresos(K).Checked = False
'
'Next K
LEERUSUARIOS
Call MENU1_Click


 

End Sub

Sub CARGAGRILLADATOS(row, col)
    Rem DATOS DE LA COLUMNA
    formatogrilla(1, 1) = "DATOS  "
    formatogrilla(1, 2) = "INGRESAR"
    
    Rem LARGO DE LOS DATOS
    
    formatogrilla(2, 1) = "10"
    formatogrilla(2, 2) = "10"
    
    Rem TIPO DE DATOS C=CEROS,S=STRING,N=NUMERICO,D=DATE,
    formatogrilla(3, 1) = "S"
    formatogrilla(3, 2) = "S"
    
    Rem FORMATO GRILLA
    formatogrilla(4, 1) = ""
    formatogrilla(4, 2) = ""
    Rem LOCCKED
    formatogrilla(5, 1) = "FALSE"
    formatogrilla(5, 2) = "FALSE"
    
    Grid3.Cols = col
    Grid3.Rows = row
    Grid3.AllowUserResizing = False
    Grid3.DisplayFocusRect = False
    Grid3.ExtendLastCol = True
    Grid3.BoldFixedCell = False
    Grid3.DrawMode = cellOwnerDraw
    Grid3.Appearance = Flat
    Grid3.ScrollBarStyle = Flat
    Grid3.FixedRowColStyle = Flat
    Grid3.Column(0).Width = 0
    
    Grid3.Column(1).Width = 10 * 10
    Grid3.Column(2).Width = 10 * 10
    
    Grid3.Cell(1, 1).text = "RUT"
    Grid3.Cell(2, 1).text = "USUARIO"
    Grid3.Cell(3, 1).text = "CLAVE"
    Grid3.Cell(4, 1).text = "NOMBRE"
    Grid3.Cell(5, 1).text = "LABOR"
    Grid3.Cell(6, 1).text = "EMAIL"
    
    Grid3.Column(1).Locked = True
    
    
    
    
End Sub

'Sub CARGAGRILLAPERMISOS(Row, Col)
'    Rem DATOS DE LA COLUMNA
'
'    Rem LARGO DE LOS DATOS
'
'    formatoGrilla(2, 1) = "40"
'    formatoGrilla(2, 2) = "2"
'
'    Rem TIPO DE DATOS C=CEROS,S=STRING,N=NUMERICO,D=DATE,
'    formatoGrilla(3, 1) = "S"
'    formatoGrilla(3, 2) = "S"
'
'    Rem FORMATO GRILLA
'    formatoGrilla(4, 1) = ""
'    formatoGrilla(4, 2) = ""
'    Rem LOCCKED
'    formatoGrilla(5, 1) = "FALSE"
'    formatoGrilla(5, 2) = "FALSE"
'
'    Grid1.Cols = Col
'    Grid1.Rows = Row
'    Grid1.AllowUserResizing = False
'    Grid1.DisplayFocusRect = False
'    Grid1.ExtendLastCol = True
'    Grid1.BoldFixedCell = False
'    Grid1.DrawMode = cellOwnerDraw
'    Grid1.Appearance = Flat
'    Grid1.ScrollBarStyle = Flat
'    Grid1.FixedRowColStyle = Flat
'    Grid1.Column(0).Width = 0
'
'    Grid1.Column(1).Width = 50 * 10
'    Grid1.Column(1).Locked = True
'
'    For k = 2 To Grid1.Cols - 1
'    Grid1.Column(k).Width = 9 * 10
'    Grid1.Column(k).CellType = cellCheckBox
'   Next k
'   Grid1.Cell(0, 1).text = "MODULO DEL SISTEMA"
'   Grid1.Cell(0, 2).text = "INGRESAR"
'   Grid1.Cell(0, 3).text = "AGREGAR"
'   Grid1.Cell(0, 4).text = "MODIFICAR"
'   Grid1.Cell(0, 5).text = "ELIMINAR"
'   Grid1.Cell(0, 6).text = "SUPERVISOR"
'
'
'End Sub

Sub CARGAGRILLAPERMISOS(row, col)
    Rem DATOS DE LA COLUMNA
    
    Rem LARGO DE LOS DATOS
    
    formatogrilla(2, 1) = "50"
    formatogrilla(2, 2) = "2"
    
    Rem TIPO DE DATOS C=CEROS,S=STRING,N=NUMERICO,D=DATE,
    formatogrilla(3, 1) = "S"
    formatogrilla(3, 2) = "S"
    
    Rem FORMATO GRILLA
    formatogrilla(4, 1) = ""
    formatogrilla(4, 2) = ""
    Rem LOCCKED
    formatogrilla(5, 1) = "FALSE"
    formatogrilla(5, 2) = "FALSE"
    
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
    
    Grid1.Column(1).Width = 34 * 10
    Grid1.Column(1).Locked = True
    
    For K = 2 To Grid1.Cols - 1
    Grid1.Column(K).Width = 7 * 10
    Grid1.Column(K).CellType = cellCheckBox
   Next K
   Grid1.Cell(0, 1).text = "MODULO DEL SISTEMA"
   Grid1.Cell(0, 2).text = "INGRESAR"
   Grid1.Cell(0, 3).text = "AGREGAR"
   Grid1.Cell(0, 4).text = "MODIFICAR"
   Grid1.Cell(0, 5).text = "ELIMINAR"
   Grid1.Cell(0, 6).text = "SUPERVISOR"
   Grid1.Cell(0, 7).text = "AUTORIZAR"
   
    
End Sub

Sub CARGAGRILLAEMPRESA(row, col)
    Rem DATOS DE LA COLUMNA
    formatogrilla(1, 1) = "CODIGO"
    formatogrilla(1, 2) = "EMPRESA"
    formatogrilla(1, 3) = "ACTIVO"
    
    Rem LARGO DE LOS DATOS
    
    formatogrilla(2, 1) = "10"
    formatogrilla(2, 2) = "10"
    formatogrilla(2, 3) = "10"
    
    
    Rem TIPO DE DATOS C=CEROS,S=STRING,N=NUMERICO,D=DATE,
    formatogrilla(3, 1) = "S"
    formatogrilla(3, 2) = "S"
    
    Rem FORMATO GRILLA
    formatogrilla(4, 1) = ""
    formatogrilla(4, 2) = ""
    Rem LOCCKED
    formatogrilla(5, 1) = "FALSE"
    formatogrilla(5, 2) = "FALSE"
    formatogrilla(5, 3) = "FALSE"
    
    
    Grid4.Cols = col
    Grid4.Rows = row
    Grid4.AllowUserResizing = False
    Grid4.DisplayFocusRect = False
    Grid4.ExtendLastCol = True
    Grid4.BoldFixedCell = False
    Grid4.DrawMode = cellOwnerDraw
    Grid4.Appearance = Flat
    Grid4.ScrollBarStyle = Flat
    Grid4.FixedRowColStyle = Flat
    Grid4.Column(0).Width = 0
    
    Grid4.Column(1).Width = 4 * 10
    Grid4.Column(2).Width = 15 * 10
    
    Grid4.Cell(0, 1).text = "CODIGO"
    Grid4.Cell(0, 2).text = "NOMBRE"
 

    Grid4.Column(3).Width = 2 * 10
    Grid4.Column(3).CellType = cellCheckBox
    

    
    
End Sub

Sub CARGAGRILLAPROGRAMASUSUARIO(row, col)
    Rem DATOS DE LA COLUMNA
    formatogrilla(1, 1) = "CODIGO"
 
    
    Rem LARGO DE LOS DATOS
    
    formatogrilla(2, 1) = "10"
    formatogrilla(2, 2) = "10"
 
    
    
    Rem TIPO DE DATOS C=CEROS,S=STRING,N=NUMERICO,D=DATE,
    formatogrilla(3, 1) = "S"
    formatogrilla(3, 2) = "S"
    
    Rem FORMATO GRILLA
    formatogrilla(4, 1) = ""
    formatogrilla(4, 2) = ""
    Rem LOCCKED
    formatogrilla(5, 1) = "FALSE"
    formatogrilla(5, 2) = "FALSE"
    formatogrilla(5, 3) = "FALSE"
    
    
    Grid6.Cols = col
    Grid6.Rows = row
    Grid6.AllowUserResizing = False
    Grid6.DisplayFocusRect = False
    Grid6.ExtendLastCol = True
    Grid6.BoldFixedCell = False
    Grid6.DrawMode = cellOwnerDraw
    Grid6.Appearance = Flat
    Grid6.ScrollBarStyle = Flat
    Grid6.FixedRowColStyle = Flat
    Grid6.Column(0).Width = 0
    
    Grid6.Column(1).Width = 10 * 15
    Grid6.Column(2).Width = 2 * 10
    Grid6.Column(2).CellType = cellCheckBox
    

    
    
End Sub
Sub CARGAGRILLAPROGRAMAPROHIBIDOS(row, col)
    Rem DATOS DE LA COLUMNA
    formatogrilla(1, 1) = "PROGRAMA"
    Rem LARGO DE LOS DATOS
    formatogrilla(2, 1) = "10"
    
    Rem TIPO DE DATOS C=CEROS,S=STRING,N=NUMERICO,D=DATE,
    formatogrilla(3, 1) = "S"
    Rem FORMATO GRILLA
    formatogrilla(4, 1) = ""
 
    Rem LOCCKED
    formatogrilla(5, 1) = "TRUE"
 
    Grid5.Cols = col
    Grid5.Rows = row
    Grid5.AllowUserResizing = False
    Grid5.DisplayFocusRect = False
    Grid5.ExtendLastCol = True
    Grid5.BoldFixedCell = False
    Grid5.DrawMode = cellOwnerDraw
    Grid5.Appearance = Flat
    Grid5.ScrollBarStyle = Flat
    Grid5.FixedRowColStyle = Flat
    Grid5.Column(0).Width = 0
    Grid5.Column(1).Width = 4 * 10
    Grid5.Column(1).Locked = True
    
End Sub



Sub CARGAGRILLAUSUARIOS(row, col)
    Rem DATOS DE LA COLUMNA
    formatogrilla(1, 1) = ""
    
    Rem LARGO DE LOS DATOS
    
    formatogrilla(2, 1) = "20"
    
    Rem TIPO DE DATOS C=CEROS,S=STRING,N=NUMERICO,D=DATE,
    formatogrilla(3, 1) = "S"
    formatogrilla(3, 2) = "S"
    
    Rem FORMATO GRILLA
    formatogrilla(4, 1) = ""
    formatogrilla(4, 2) = ""
    Rem LOCCKED
    formatogrilla(5, 1) = "FALSE"
    formatogrilla(5, 2) = "FALSE"
    
    Grid2.Cols = col
    Grid2.Rows = row
    Grid2.AllowUserResizing = False
    Grid2.DisplayFocusRect = False
    Grid2.ExtendLastCol = True
    Grid2.BoldFixedCell = False
    Grid2.DrawMode = cellOwnerDraw
    Grid2.Appearance = Flat
    Grid2.ScrollBarStyle = Flat
    Grid2.FixedRowColStyle = Flat
    Grid2.Column(0).Width = 0
    
    Grid2.Column(1).Width = 10 * 10
    Grid2.Cell(0, 1).text = "USUARIOS"
    
    Grid2.Column(1).Locked = True
    
    
    
    
    
End Sub




Private Sub Grid1_Click()
If Grid2.Cell(Grid2.ActiveCell.row, 1).text <> "" Then

Call grabarpermiso(Grid1.Cell(Grid1.ActiveCell.row, 1).text)
End If


End Sub

Private Sub Grid2_Click()
Dim o As Integer
Dim i As Double

seguridad2.Caption = "MODULO DE SEGURIDAD USUARIO ACTIVO =" + Grid2.Cell(Grid2.ActiveCell.row, 1).text
Call LEERUSUARIOindividual(Grid2.Cell(Grid2.ActiveCell.row, 1).text)
USUARIOSELECCIONADO = Grid2.Cell(Grid2.ActiveCell.row, 1).text
MENU1_Click

For o = 1 To Grid1.Rows - 1
    Call leerpermisos2(USUARIOSELECCIONADO, Grid1.Cell(o, 1).text, o)
Next o
For i = 1 To Grid6.Rows - 1
    Call leerprogramasprohibidosusuario(USUARIOSELECCIONADO, Grid6.Cell(i, 1).text, i, Grid6)
Next i

Command4.Enabled = True
'Call leerpermisos(USUARIOSELECCIONADO)
End Sub

Private Sub Grid2_KeyDown(KeyCode As Integer, ByVal Shift As Integer)
Dim cSql2 As New rdoQuery
    Dim csql As New rdoQuery
   
If KeyCode = 46 Then
If MsgBox("ESTA SEGURO QUE DESEA ELIMINAR A " + Grid2.Cell(Grid2.ActiveCell.row, 1).text + " Y SUS PERMISOS", vbOKCancel, "ATENCION") = vbOK Then

        Set cSql2.ActiveConnection = ventas
        cSql2.sql = "DELETE "
        cSql2.sql = cSql2.sql + "FROM segu_permisos "
        cSql2.sql = cSql2.sql + "where usuario='" + Grid2.Cell(Grid2.ActiveCell.row, 1).text + "' "
        cSql2.Execute
        Set csql.ActiveConnection = ventas
        csql.sql = "DELETE "
        csql.sql = csql.sql + "FROM " & clientesistema & "auditoria.segu_usuarios "
        csql.sql = csql.sql + "where usuario='" + Grid2.Cell(Grid2.ActiveCell.row, 1).text + "' "
        csql.Execute
        
        Set csql.ActiveConnection = ventas
        csql.sql = "DELETE "
        csql.sql = csql.sql + "FROM " & clientesistema & "auditoria.usuarios_prohibidos "
        csql.sql = csql.sql + "where usuario='" + Grid2.Cell(Grid2.ActiveCell.row, 1).text + "' "
        csql.Execute
End If
End If
LEERUSUARIOS
MENU1_Click

End Sub


Private Sub Grid3_KeyPress(KeyAscii As Integer)
  KeyAscii = Asc(UCase(Chr(KeyAscii)))
  
  If KeyAscii = 13 And Grid3.Cell(1, 2).text <> "" Then
  pivote.MaxLength = 10
  pivote.text = Grid3.Cell(1, 2).text
  Call ceros(pivote)
  Grid3.Cell(1, 2).text = pivote.text
  End If
End Sub

Private Sub Grid3_LeaveCell(ByVal row As Long, ByVal col As Long, NewRow As Long, NewCol As Long, Cancel As Boolean)
NewCol = 2
End Sub



Private Sub ingresos_Click(Index As Integer)
Dim VARIABLE As String
VARIABLE = Ingresos.Count



permiso.Caption = Ingresos(Index).Caption
VARIABLE = Ingresos(Index).Caption


Ingresos(Index).Checked = True
menuseleccion = "ingresos(" & Index & ")"
eli = False
Command3.Visible = True

'Call grabarpermiso("ingresos(" & Index & ")", VARIABLE, False)
'Call leerpermisos(USUARIOSELECCIONADO)






End Sub
Sub grabarpermiso(nombreprograma As String)
Dim campos(20, 3) As String

Dim op As Integer
Set sql = New sqlventas.sqlventa
    nombreprograma = achica(nombreprograma)
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
  
    campos(0, 1) = USUARIOSELECCIONADO
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
    condicion = "usuario=" + "'" + USUARIOSELECCIONADO + "' and programa='" + nombreprograma + "'"
    
    sql.response = campos
    Set sql.conexion = ventas
    
    If ELIMINA = False Then
    op = 5
    Call sql.sqlventas(op, condicion)
  
  
  If sql.Status = 4 Then
  op = 2
  condicion = ""
  Else
  op = 3
  End If
  Call sql.sqlventas(op, condicion)
  Else
  op = 4
  Call sql.sqlventas(op, condicion)


End If


     
End Sub
Sub LEERUSUARIOS()
    Dim resultados2 As rdoResultset
    Dim cSql2 As New rdoQuery
    Dim saldodebe As String
    Dim saldohaber As String
    Dim lineas As Double
    

        Set cSql2.ActiveConnection = ventas
        cSql2.sql = "SELECT * "
        cSql2.sql = cSql2.sql + "FROM " & clientesistema & "auditoria.segu_usuarios "
        cSql2.sql = cSql2.sql + "order by usuario "
        cSql2.Execute
        Grid2.Rows = cSql2.RowsAffected + 1
        
        If cSql2.RowsAffected > 0 Then
        Set resultados2 = cSql2.OpenResultset
        lineas = 0
        While Not resultados2.EOF
        lineas = lineas + 1
        Grid2.Cell(lineas, 1).text = resultados2(1)
        
        resultados2.MoveNext
        Wend
          
          resultados2.Close
            Set resultados2 = Nothing

        End If
      
 
  

End Sub
Sub LEERUSUARIOindividual(usuario)
    Dim resultados2 As rdoResultset
    Dim cSql2 As New rdoQuery
    Dim saldodebe As String
    Dim saldohaber As String
    Dim lineas As Double
    Dim INICIO As Double
    
    

        Set cSql2.ActiveConnection = ventas
        cSql2.sql = "SELECT * "
        cSql2.sql = cSql2.sql + "FROM " & clientesistema & "auditoria.segu_usuarios where usuario='" + usuario + "' "
        cSql2.Execute
        
        If cSql2.RowsAffected > 0 Then
        Set resultados2 = cSql2.OpenResultset
        lineas = 1
        While Not resultados2.EOF
        
        Grid3.Cell(1, 2).text = resultados2(0)
        Grid3.Cell(2, 2).text = resultados2(1)
        Grid3.Cell(3, 2).text = resultados2(2)
        Grid3.Cell(4, 2).text = resultados2(3)
        Grid3.Cell(5, 2).text = resultados2(4)
        Grid3.Cell(6, 2).text = resultados2(5)
        resultados2.MoveNext
        Wend
          
          resultados2.Close
            Set resultados2 = Nothing

        End If
       leerempresa2
       
    
    
        Set cSql2.ActiveConnection = ventas
        cSql2.sql = "SELECT * "
        cSql2.sql = cSql2.sql + "FROM segu_empresas where usuario='" + usuario + "' "
        cSql2.sql = cSql2.sql + "order by empresa "
        
        cSql2.Execute
        
        
        If cSql2.RowsAffected > 0 Then
        
        Set resultados2 = cSql2.OpenResultset
        While Not resultados2.EOF
        For INICIO = 1 To Grid4.Rows - 1
        If resultados2(1) = Grid4.Cell(INICIO, 1).text Then
            Grid4.Cell(INICIO, 3).text = resultados2(2)
        End If
        Next INICIO
        
        
        resultados2.MoveNext
        lineas = lineas + 1
        Wend
          
          resultados2.Close
            Set resultados2 = Nothing

        End If
 
    For K = 1 To 5
    Grid1.Cell(K, 2).text = 0
    Next K

End Sub

Private Function leerempresa2()
    Dim resultados2 As rdoResultset
    Dim cSql2 As New rdoQuery
    Dim lineas As Double
    

        Set cSql2.ActiveConnection = gestion
        cSql2.sql = "SELECT * "
        cSql2.sql = cSql2.sql + "FROM g_maestroempresas order by codigo "
        cSql2.Execute
        
        Grid4.Rows = cSql2.RowsAffected + 1
        
        If cSql2.RowsAffected > 0 Then
        Set resultados2 = cSql2.OpenResultset
        lineas = 0
        While Not resultados2.EOF
        lineas = lineas + 1
        Grid4.Cell(lineas, 1).text = resultados2(0)
        Grid4.Cell(lineas, 2).text = resultados2(1)
        Grid4.Cell(lineas, 3).text = 0
        
        resultados2.MoveNext
       
        Wend
        resultados2.Close
        Set resultados2 = Nothing
        End If
End Function


Private Function leerpermisos2(usuario, MENU, linea)
    Dim resultados2 As rdoResultset
    Dim cSql2 As New rdoQuery
    Dim lineas As Double
    Dim final As Double
    MENU = achica(MENU)

        Set cSql2.ActiveConnection = ventas
        cSql2.sql = "SELECT * "
        cSql2.sql = cSql2.sql + "FROM segu_permisos "
        cSql2.sql = cSql2.sql + "where usuario='" + usuario + "' and programa='" + MENU + "'"
        cSql2.Execute
       
        
        If cSql2.RowsAffected > 0 Then
        Set resultados2 = cSql2.OpenResultset
         
        Grid1.AutoRedraw = False
   
        While Not resultados2.EOF
      
        Grid1.Cell(linea, 2).text = resultados2(3)
        Grid1.Cell(linea, 3).text = resultados2(4)
        Grid1.Cell(linea, 4).text = resultados2(5)
        Grid1.Cell(linea, 5).text = resultados2(6)
        Grid1.Cell(linea, 6).text = resultados2(8)
        Grid1.Cell(linea, 7).text = resultados2(7)
        resultados2.MoveNext
       
        Wend
        resultados2.Close
        Set resultados2 = Nothing
        Grid1.AutoRedraw = True
        Grid1.Refresh
  
        
       
        End If
End Function
Private Function achica(palabra) As String
Dim INICIO As Double
Dim final As Double
For K = 1 To Len(palabra)
If Mid(palabra, K, 1) <> Chr(32) Then INICIO = K: Exit For

Next K

achica = Mid(palabra, INICIO, Len(palabra) - INICIO)

End Function

Sub ACTIVAMENU(ByVal Opcion As String)


'For K = 1 To ingresos.Count
'
'
'If ingresos(K).caption = Opcion Then ingresos(K).Checked = True
'
'
'Next K
'

End Sub

Private Sub Grid4_Click()
If Grid2.Cell(Grid2.ActiveCell.row, 1).text <> "" Then
Call grabarempresa
End If

End Sub
 
 

Private Sub Grid5_KeyDown(KeyCode As Integer, ByVal Shift As Integer)
    Dim i As Double
    If KeyCode = 46 Then
        If MsgBox("ESTA SEGURO QUE DESEA ELIMINAR  " & Grid5.Cell(Grid5.ActiveCell.row, 1).text & "", vbOKCancel, "ATENCION") = vbOK Then
            eli = True
            Call grabarprogramaprohibido(Grid5.Cell(Grid5.ActiveCell.row, 1).text)
            Call leerprogramasprohibidos(Grid5)
            Call leerprogramasprohibidos(Grid6)
            
            For i = 1 To Grid6.Rows - 1
                Call leerprogramasprohibidosusuario(USUARIOSELECCIONADO, Grid6.Cell(i, 1).text, i, Grid6)
            Next i
            eli = False
        End If
    End If
End Sub

Private Sub Grid6_Click()
If Grid6.Cell(Grid6.ActiveCell.row, 1).text <> "" Then
    Call grabarprogramausuario(Grid6.Cell(Grid6.ActiveCell.row, 1).text, Grid6.Cell(Grid6.ActiveCell.row, 2).text, USUARIOSELECCIONADO)
End If

End Sub

'Private Sub MENU1_Click()
'Dim contador As Double
'Dim INICIO As Double
'Dim final As Double
'Dim pasar As Double
'Dim NIVEL As String
'Dim NIVELBANDERA As String
'Call CARGAGRILLAPERMISOS(6, 8)
'
'Close 20
'
'Open App.Path + "\principal.txt" For Input As #20
'Grid1.Rows = 1
'pasar = 0
'While Not EOF(20)
'Line Input #20, VARIPASO
'If contador = 1 Then
'For K = 1 To Len(VARIPASO)
'
'If Mid(VARIPASO, K, 1) = Chr(34) Then
'VARIPASO = Mid(VARIPASO, K + 1, 80)
'K = Len(VARIPASO) + 1
'End If
'
'Next K
'For K = 1 To Len(VARIPASO)
'
'If Mid(VARIPASO, K, 1) = Chr(34) Then
'VARIPASO = Mid(VARIPASO, 1, K)
'K = Len(VARIPASO) + 1
'End If
'
'Next K
'VARIPASO = Replace(VARIPASO, Chr(34), " ")
'Grid1.Rows = Grid1.Rows + 1
'If NIVELBANDERA = "0" Then
'
'Rem Grid1.Range(Grid1.Rows - 1, 3, Grid1.Rows - 1, 7).Merge
'Grid1.Cell(Grid1.Rows - 1, 1).Font.Bold = True
'
'End If
'
'
'
'Grid1.Cell(Grid1.Rows - 1, 1).text = NIVEL + VARIPASO
'contador = 0
'End If
'
'For K = 1 To Len(VARIPASO) - 13
'If UCase(Mid(VARIPASO, K, 13)) = "BEGIN VB.MENU" Then
'
'        If K = 4 Then
'        NIVEL = ""
'        NIVELBANDERA = "0"
'        contador = 1
'        End If
'
'
'
'        If K = 7 Then
'        NIVELBANDERA = "1"
'        NIVEL = "       "
'        contador = 1
'        End If
'
'        If K = 10 Then
'        NIVELBANDERA = "3"
'        NIVEL = "               "
'        contador = 1
'        End If
'
'Exit For
'Else
'contador = 0
'
'End If
'
'
'
'Next K
'
''
'
'Wend
'
''OPCIONES1.Clear
''For K = 1 To PRINCIPAL.ingresos.Count
''OPCIONES1.AddItem (PRINCIPAL.ingresos(K).Caption)
''Next K
'
'End Sub
Private Sub MENU1_Click()
Dim contador As Double
Dim INICIO As Double
Dim final As Double
Dim pasar As Double
Dim NIVEL As String
Dim NIVELBANDERA As String
Call CARGAGRILLAPERMISOS(6, 8)

Close 20

Open App.Path + "\principal.txt" For Input As #20
Grid1.Rows = 1
pasar = 0
While Not EOF(20)
Line Input #20, VARIPASO
If contador = 1 Then
For K = 1 To Len(VARIPASO)

If Mid(VARIPASO, K, 1) = Chr(34) Then
VARIPASO = Mid(VARIPASO, K + 1, 50)
K = Len(VARIPASO) + 1
End If

Next K
For K = 1 To Len(VARIPASO)

If Mid(VARIPASO, K, 1) = Chr(34) Then
VARIPASO = Mid(VARIPASO, 1, K)
K = Len(VARIPASO) + 1
End If

Next K
VARIPASO = Replace(VARIPASO, Chr(34), " ")
Grid1.Rows = Grid1.Rows + 1
If NIVELBANDERA = "0" Then

Rem Grid1.Range(Grid1.Rows - 1, 3, Grid1.Rows - 1, 7).Merge
Grid1.Cell(Grid1.Rows - 1, 1).Font.Bold = True



End If

VARIPASO = Replace(VARIPASO, "&", "")


Grid1.Cell(Grid1.Rows - 1, 1).text = NIVEL + VARIPASO
contador = 0
End If

For K = 1 To Len(VARIPASO) - 13
If UCase(Mid(VARIPASO, K, 13)) = "BEGIN VB.MENU" Then
        
        If K = 4 Then
        NIVEL = ""
        NIVELBANDERA = "0"
        contador = 1
        End If
        
        
        
        If K = 7 Then
        NIVELBANDERA = "1"
        NIVEL = "       "
        contador = 1
        End If
        
        If K = 10 Then
        NIVELBANDERA = "3"
        NIVEL = "               "
        contador = 1
        End If

Exit For
Else
contador = 0

End If



Next K

'

Wend
End Sub
Sub grabarusuario()
    Dim campos(30, 3) As String
    
    Dim op As Integer
    Set sql = New sqlventas.sqlventa
    campos(0, 0) = "rut"
    campos(1, 0) = "usuario"
    campos(2, 0) = "clave"
    campos(3, 0) = "nombre"
    campos(4, 0) = "labor"
    campos(5, 0) = "email"
    campos(6, 0) = ""
  
    campos(0, 1) = Grid3.Cell(1, 2).text
    campos(1, 1) = Grid3.Cell(2, 2).text
    campos(2, 1) = Grid3.Cell(3, 2).text
    campos(3, 1) = Grid3.Cell(4, 2).text
    campos(4, 1) = Grid3.Cell(5, 2).text
    campos(5, 1) = Grid3.Cell(6, 2).text
    
   
    
    campos(0, 2) = clientesistema & "auditoria.segu_usuarios"
    condicion = "usuario=" + "'" + Grid3.Cell(2, 2).text + "' "
    
    sql.response = campos
    Set sql.conexion = ventas
   

    op = 5
    Call sql.sqlventas(op, condicion)
  
  
  If sql.Status = 4 Then
  op = 2
  condicion = ""
  Else
  op = 3
  End If
  Call sql.sqlventas(op, condicion)
    
End Sub
Sub copiarpermisos(usuarioorigen, usuariodestino)
    Dim resultados2 As rdoResultset
    Dim cSql2 As New rdoQuery

        Set cSql2.ActiveConnection = ventas
        cSql2.sql = "INSERT INTO segu_permisos (usuario,programa,ingresa,agrega,modifica,elimina,autoriza,todas) "
        cSql2.sql = cSql2.sql + "SELECT  '" + usuariodestino + "',programa,ingresa,agrega,modifica,elimina,autoriza,todas "
        cSql2.sql = cSql2.sql + "FROM segu_permisos "
        cSql2.sql = cSql2.sql + "where usuario='" + usuarioorigen + "'"
        cSql2.Execute
         
End Sub

Sub grabarempresa()
    Dim campos(5, 5) As String
    Dim op As Integer
    
    Set sql = New sqlventas.sqlventa
    campos(0, 0) = "usuario"
    campos(1, 0) = "empresa"
    campos(2, 0) = "permiso"
    campos(3, 0) = ""
  
    campos(0, 1) = USUARIOSELECCIONADO
    campos(1, 1) = Grid4.Cell(Grid4.ActiveCell.row, 1).text
    campos(2, 1) = Grid4.Cell(Grid4.ActiveCell.row, 3).text
  
    
    campos(0, 2) = "segu_empresas"
    condicion = "usuario=" + "'" + USUARIOSELECCIONADO + "' and empresa='" + Grid4.Cell(Grid4.ActiveCell.row, 1).text + "'"
    
    sql.response = campos
    Set sql.conexion = ventas
    If ELIMINA = False Then
    op = 5
    Call sql.sqlventas(op, condicion)
  
  
  If sql.Status = 4 Then
  op = 2
  condicion = ""
  Else
  op = 3
  End If
  Call sql.sqlventas(op, condicion)
  Else
  op = 4
  Call sql.sqlventas(op, condicion)
End If
End Sub
Sub leerprogramasprohibidos(grilla)
    Dim csql As New rdoQuery
    Dim resultados As rdoResultset
    
    Set csql.ActiveConnection = ventas
    
    csql.sql = "select programa from " & clientesistema & "auditoria" & ".programas_prohibidos order by programa "
    csql.Execute
    If csql.RowsAffected > 0 Then
        Set resultados = csql.OpenResultset
        grilla.Rows = 1
        While Not resultados.EOF
            grilla.Rows = grilla.Rows + 1
            grilla.Cell(grilla.Rows - 1, 1).text = resultados(0)
            resultados.MoveNext
        Wend
    End If
    csql.Close
    Set csql = Nothing
    Set resultados = Nothing
    
End Sub
Sub grabarprogramaprohibido(programa)
Dim campos(4, 3) As String

Dim op As Integer
 Set sql = New sqlventas.sqlventa
    campos(0, 0) = "programa"
    campos(1, 0) = ""
    campos(0, 1) = programa
    
    campos(0, 2) = clientesistema & "auditoria.programas_prohibidos"
    condicion = "programa='" & programa & "'"
    
    sql.response = campos
    Set sql.conexion = ventas
     If eli = False Then
        op = 5
        Call sql.sqlventas(op, condicion)
        If sql.Status = 4 Then
            op = 2
            condicion = ""
        Else
            op = 3
        End If
        Call sql.sqlventas(op, condicion)
    Else
        op = 4
        Call sql.sqlventas(op, condicion)
    End If
End Sub

Private Function leerprogramasprohibidosusuario(usuario, programa, linea, grilla)
    Dim resultados2 As rdoResultset
    Dim cSql2 As New rdoQuery
    Dim lineas As Double
    Dim final As Double

        Set cSql2.ActiveConnection = ventas
        cSql2.sql = "SELECT usuario,programa "
        cSql2.sql = cSql2.sql + "FROM " & clientesistema & "auditoria.usuarios_prohibidos "
        cSql2.sql = cSql2.sql + "where usuario='" & usuario & "' and programa='" & programa & "'"
        cSql2.Execute
        If cSql2.RowsAffected > 0 Then
            grilla.AutoRedraw = False
            grilla.Cell(linea, 2).text = "1"
            
        Else
            grilla.Cell(linea, 2).text = "0"
        End If
        
     
        grilla.AutoRedraw = True
        grilla.Refresh
End Function
Sub grabarprogramausuario(programa, estado, usuario)
Dim campos(5, 3) As String

Dim op As Integer
Set sql = New sqlventas.sqlventa
    campos(0, 0) = "programa"
    campos(1, 0) = "usuario"
    campos(2, 0) = ""
    campos(0, 1) = programa
    campos(1, 1) = usuario
    
    campos(0, 2) = clientesistema & "auditoria.usuarios_prohibidos"
    condicion = "programa='" & programa & "' and usuario='" & usuario & "' "
    
    sql.response = campos
    Set sql.conexion = ventas
    
     If estado = "1" Then
        op = 5
        Call sql.sqlventas(op, condicion)
        If sql.Status = 4 Then
            op = 2
            condicion = ""
            Call sql.sqlventas(op, condicion)
        End If
    Else
        op = 4
        Call sql.sqlventas(op, condicion)
    End If
End Sub
