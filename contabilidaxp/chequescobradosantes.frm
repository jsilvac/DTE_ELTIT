VERSION 5.00
Object = "{4777436C-EB8C-4596-98A8-EBCF98683969}#1.0#0"; "FlexCell.ocx"
Object = "{085FBDF3-00C5-421B-B762-1D57299A2B89}#1.0#0"; "CLBUTN.OCX"
Object = "{ADBBAED6-B16F-40DC-80DF-B44910CBA76C}#1.0#0"; "Frame-Xp.ocx"
Begin VB.Form banco04 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Listado de Cheques Cobrados Antes"
   ClientHeight    =   8595
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   15210
   DrawStyle       =   5  'Transparent
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   Moveable        =   0   'False
   ScaleHeight     =   573
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1014
   Begin XPFrame.FrameXp FrameQuickMenu 
      Height          =   615
      Left            =   11760
      TabIndex        =   20
      Top             =   0
      Width           =   3135
      _ExtentX        =   5530
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
         Left            =   1680
         TabIndex        =   22
         Top             =   280
         Width           =   1335
      End
      Begin VB.CommandButton botonmisaccesos 
         Caption         =   "Permisos Modulo"
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   280
         Width           =   1455
      End
   End
   Begin VB.TextBox pivote 
      Height          =   285
      Left            =   6750
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   8865
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox MANUAL 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   135
      Left            =   120
      ScaleHeight     =   135
      ScaleWidth      =   135
      TabIndex        =   4
      Top             =   6120
      Width           =   135
   End
   Begin XPFrame.FrameXp FrameXp1 
      Height          =   8610
      Left            =   240
      TabIndex        =   5
      Top             =   0
      Width           =   14655
      _ExtentX        =   25850
      _ExtentY        =   15187
      BackColor       =   16744576
      Caption         =   " L I S T A D O  D E  C H E Q U E S    COBRADOS ANTES "
      CaptionEstilo3D =   1
      BackColor       =   16744576
      ForeColor       =   8438015
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   1
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FFC0C0&
         Caption         =   "IMPRIMIR"
         Height          =   330
         Left            =   5085
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   8190
         Width           =   2130
      End
      Begin XPFrame.FrameXp FrameXp3 
         Height          =   1290
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   14145
         _ExtentX        =   24950
         _ExtentY        =   2275
         BackColor       =   16744576
         Caption         =   "Cuenta Bancaria"
         CaptionEstilo3D =   1
         BackColor       =   16744576
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
         Begin VB.CommandButton ver 
            Caption         =   "Ver"
            Height          =   285
            Left            =   13200
            TabIndex        =   11
            Top             =   960
            Visible         =   0   'False
            Width           =   780
         End
         Begin VB.TextBox dato3 
            BackColor       =   &H00FFFFFF&
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
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   900
            MaxLength       =   4
            TabIndex        =   3
            Top             =   720
            Width           =   735
         End
         Begin VB.TextBox dato2 
            BackColor       =   &H00FFFFFF&
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
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   495
            MaxLength       =   2
            TabIndex        =   2
            Top             =   720
            Width           =   375
         End
         Begin VB.TextBox dato1 
            BackColor       =   &H00FFFFFF&
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
            Left            =   90
            MaxLength       =   2
            TabIndex        =   1
            Tag             =   "codigo"
            Top             =   720
            Width           =   375
         End
         Begin XPFrame.FrameXp FrameXp8 
            Height          =   1215
            Left            =   8160
            TabIndex        =   13
            Top             =   240
            Width           =   6375
            _ExtentX        =   11245
            _ExtentY        =   2143
            BackColor       =   14737632
            Caption         =   "            Rangos de Fecha"
            CaptionEstilo3D =   1
            BackColor       =   14737632
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
            Begin CoolButtons.cool_Button cool_Button3 
               Height          =   375
               Left            =   4320
               TabIndex        =   14
               Top             =   360
               Width           =   1335
               _ExtentX        =   2355
               _ExtentY        =   661
               SkinId          =   "13"
               Caption         =   "Cambia Fecha"
            End
            Begin VB.Label Label16 
               BackStyle       =   0  'Transparent
               BorderStyle     =   1  'Fixed Single
               Caption         =   "Desde Fecha"
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
               Height          =   255
               Left            =   120
               TabIndex        =   18
               Top             =   240
               Width           =   1935
            End
            Begin VB.Label Label17 
               BackStyle       =   0  'Transparent
               BorderStyle     =   1  'Fixed Single
               Caption         =   "Hasta Fecha"
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
               Height          =   255
               Left            =   2160
               TabIndex        =   17
               Top             =   240
               Width           =   1935
            End
            Begin VB.Label desdefecha 
               BackColor       =   &H00FFC0C0&
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
               Height          =   375
               Left            =   120
               TabIndex        =   16
               Top             =   480
               Width           =   1935
            End
            Begin VB.Label hastafecha 
               BackColor       =   &H00FFC0C0&
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
               Height          =   375
               Left            =   2160
               TabIndex        =   15
               Top             =   480
               Width           =   1935
            End
         End
         Begin VB.Label Label8 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Cuenta Bancaria"
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
            TabIndex        =   10
            Top             =   360
            Width           =   6450
         End
         Begin VB.Label nombrebanco 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
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
            TabIndex        =   9
            Top             =   720
            Width           =   6450
         End
         Begin VB.Label Label1 
            BorderStyle     =   1  'Fixed Single
            Caption         =   " Banco"
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
            Left            =   90
            TabIndex        =   8
            Top             =   360
            Width           =   1545
         End
      End
      Begin XPFrame.FrameXp FrameXp2 
         Height          =   6630
         Left            =   135
         TabIndex        =   6
         Top             =   1530
         Width           =   14070
         _ExtentX        =   24818
         _ExtentY        =   11695
         BackColor       =   16744576
         Caption         =   "Cheques pendientes de Cobro"
         CaptionEstilo3D =   1
         BackColor       =   16744576
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
         Begin FlexCell.Grid Grid1 
            Height          =   6405
            Left            =   0
            TabIndex        =   19
            Top             =   225
            Width           =   14025
            _ExtentX        =   24739
            _ExtentY        =   11298
            Cols            =   5
            DefaultFontSize =   8.25
            Rows            =   30
         End
      End
   End
End
Attribute VB_Name = "banco04"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub codigo_Click()
    Call dato1_KeyDown(vbKeyF2, 0)
End Sub

Private Sub botonmisaccesos_Click()
programafiltro = Me.Caption
misaccesos.Show
End Sub


Private Sub botonmisfavoritos_Click()
Call AgregaFavoritos(USUARIOSISTEMA, App.EXEName, Me.Name, Me.Caption)

End Sub

Private Sub Command1_Click()
imprimir

End Sub

Private Sub cool_Button3_Click()
Call retornofecha(desdefecha, hastafecha)
End Sub



Private Sub dato2_GotFocus()
Call cargatexto(dato2)
End Sub
Private Sub dato3_GotFocus()
Call cargatexto(dato3)
End Sub
Sub leercuenta()
    Rem lee cuenta madre
  
lee2:    Rem lee cuenta madre
    campos(0, 0) = dato1.Tag 'CODIGO
    campos(1, 0) = "nombre"
    campos(2, 0) = ""
    campos(0, 2) = "cuentasdelmayor"
    condicion = "codigo=" + "'" + dato1.text + dato2.text + dato3.text + "' and año='" + Format(fechasistema, "yyyy") + "'"
    op = 5
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)
    If sqlconta.status = 4 Then dato1.SetFocus: GoTo no:
    nombrebanco.Caption = sqlconta.response(1, 3)
    leer
    
no:
   
End Sub


Private Sub dato1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 38 Then Unload Me: GoTo no:
    If KeyCode = vbKeyF2 Then Call ayudamayor(dato3)
    Call flechas(dato1, dato2, KeyCode)
no:
End Sub

Private Sub dato2_KeyDown(KeyCode As Integer, Shift As Integer)
    Call flechas(dato1, dato3, KeyCode)
End Sub

Private Sub dato3_KeyDown(KeyCode As Integer, Shift As Integer)
    Call flechas(dato2, dato3, KeyCode)
End Sub







Private Sub dato3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Call ceros(dato3)
leercuenta
leer
Rem desdefecha.Caption = "01-01-" + Format(fechasistema, "yyyy")
Rem hastafecha.Caption = fechasistema
End If

End Sub


Private Sub Form_Load()
CENTRAR Me


    
    Call Conectar_BD

    sc = 0
CARGAGRILLA


desdefecha.Caption = "01" + "-" + Format(fechasistema, "mm") + "-" + Format(fechasistema, "yyyy")
hastafecha.Caption = fechasistema

End Sub

Private Sub dato1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then Unload Me
    snum = 0: KeyAscii = esNumero(KeyAscii)
    If KeyAscii = 13 Then Call ceros(dato1): Call Pregunta(dato1, dato2)
End Sub

Private Sub dato2_KeyPress(KeyAscii As Integer)
    snum = 0: KeyAscii = esNumero(KeyAscii)

    If KeyAscii = 13 Then Call ceros(dato2): Call Pregunta(dato2, dato3)
End Sub



Sub carga()
    habilita (True)
    dato1.text = Mid(sqlconta.response(0, 3), 1, 2)
    dato2.text = Mid(sqlconta.response(0, 3), 3, 2)
    dato3.text = Mid(sqlconta.response(0, 3), 5, 4)
fin:
End Sub

Sub habilita(ByVal condicion As Boolean)
    
    dato1.Locked = condicion
    dato2.Locked = condicion
    dato3.Locked = condicion
    
    
End Sub
Sub disponible(ByVal condicion As Boolean)
    
    dato1.Enabled = condicion
    dato2.Enabled = condicion
    dato3.Enabled = condicion
    

End Sub


Sub Pregunta(ByRef caja As TextBox, ByRef caja1 As TextBox)
    If caja.text = "" And sc = 0 Then caja.SetFocus
    If caja.text <> "" Or sc = 1 Then caja1.Enabled = True: caja1.SetFocus
    sc = 0
End Sub


Sub ayudamayor(ByRef caja As TextBox)
    Dim campos As Variant
    Dim cfijo As Variant
    Dim largo As Variant
    campos = Array("codigo", "nombre")
    largo = Array("12s", "40s")
    cfijo = "banco='1'"
    cabezas = Array("codigo", "nombre")
    mensajeAyuda = "Ayuda Cuentas del Mayor"
    
    Call cargaAyudaT(Servidor, basebus, Usuario, password, "cuentasdelmayor", pivote, campos, cfijo, largo, 2)
    If Val(pivote.text) = 0 Then dato1.SetFocus: GoTo no
    dato2.Enabled = True
    dato3.Enabled = True
    dato1.text = Mid(pivote.text, 1, 2)
    dato2.text = Mid(pivote.text, 3, 2)
    dato3.text = Mid(pivote.text, 5, 4)
    caja.Enabled = True
    caja.SetFocus
    Call dato3_KeyPress(13)
no:
End Sub



Sub flechas(ByRef caja As TextBox, ByRef caja1 As TextBox, ByRef codigo As Integer)
    If codigo = 38 And caja.Enabled = True Then caja.SetFocus
    If codigo = 40 And caja1.Enabled = True Then caja1.SetFocus
End Sub

Sub ELIMINAR()
    
    campos(0, 2) = "cuentasdelmayor"
    condicion = "codigo=" + "'" + dato1.text + dato2.text + dato3.text + "'"
    op = 4
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)

    
End Sub


Private Sub lblhistorico_Click(Index As Integer)

End Sub






Sub limpia()
    dato1.text = ""
    dato2.text = ""
    dato3.text = ""
    
    
End Sub

Sub imprimir()
Dim titulo As String

titulo = "CHEQUES A FECHA COBRADOS ANTES "
Call CABEZAS2(titulo, "N", 1)
Grid1.DefaultFont.Size = 8
Grid1.PageSetup.Orientation = cellPortrait
Grid1.Range(0, 1, 0, Grid1.Cols - 1).Borders(cellEdgeBottom) = cellThick
Grid1.Range(0, 1, 0, Grid1.Cols - 1).Borders(cellEdgeLeft) = cellThick
Grid1.Range(0, 1, 0, Grid1.Cols - 1).Borders(cellEdgeTop) = cellThick
Grid1.Range(0, 1, 0, Grid1.Cols - 1).Borders(cellEdgeRight) = cellThick
Grid1.Range(0, 1, 0, Grid1.Cols - 1).Borders(cellInsideHorizontal) = cellThick
Grid1.Range(0, 1, 0, Grid1.Cols - 1).Borders(cellInsideVertical) = cellThick



Grid1.PageSetup.PrintFixedRow = True
Grid1.PageSetup.BottomMargin = 1
Grid1.PageSetup.TopMargin = 1
Grid1.PageSetup.LeftMargin = 1
Grid1.PageSetup.RightMargin = 0
Grid1.PageSetup.BlackAndWhite = True
Grid1.PageSetup.PrintGridlines = False
Grid1.PrintPreview 100

   
End Sub
Sub grilla()
    
End Sub
Sub cabeza()
    

End Sub




Private Sub opciones_GotFocus()

MANUAL.SetFocus

End Sub
Sub CARGAGRILLA()
Rem DATOS DE LA COLUMNA
    Dim FORMATOGRILLA(10, 10)
    Grid1.DefaultFont.Size = 8
       
    FORMATOGRILLA(1, 1) = "EMISION"
    FORMATOGRILLA(1, 2) = "TIPO"
    FORMATOGRILLA(1, 3) = "NUMERO"
    FORMATOGRILLA(1, 4) = "CHEQUE"
    FORMATOGRILLA(1, 5) = "GIRADO A"
    FORMATOGRILLA(1, 6) = "VENCIMIENTO"
    FORMATOGRILLA(1, 7) = "MONTO"
    FORMATOGRILLA(1, 8) = "ACUMULADO"
    FORMATOGRILLA(1, 9) = "FECHA COBRO"
    
     
    Rem LARGO DE LOS DATOS
    FORMATOGRILLA(2, 1) = "8"
    FORMATOGRILLA(2, 2) = "0"
    FORMATOGRILLA(2, 3) = "0"
    
    FORMATOGRILLA(2, 4) = "10"
    FORMATOGRILLA(2, 5) = "30"
    FORMATOGRILLA(2, 6) = "10"
    FORMATOGRILLA(2, 7) = "10"
    FORMATOGRILLA(2, 8) = "10"
    FORMATOGRILLA(2, 9) = "10"

    
    Rem TIPO DE DATOS C=CEROS,S=STRING,N=NUMERICO,D=DATE,
    FORMATOGRILLA(3, 1) = "D"
    FORMATOGRILLA(3, 2) = "S"
    FORMATOGRILLA(3, 3) = "N"
    
    FORMATOGRILLA(3, 4) = "S"
    FORMATOGRILLA(3, 5) = "S"
    FORMATOGRILLA(3, 6) = "D"
    FORMATOGRILLA(3, 7) = "N"
    FORMATOGRILLA(3, 8) = "N"
    FORMATOGRILLA(3, 9) = "D"
   
    
    Rem FORMATO GRILLA
    FORMATOGRILLA(4, 7) = "###,###,###,##0"
    FORMATOGRILLA(4, 8) = "###,###,###,##0"
  
    
    Rem LOCCKED
    FORMATOGRILLA(5, 1) = "TRUE"
    FORMATOGRILLA(5, 2) = "TRUE"
    FORMATOGRILLA(5, 3) = "TRUE"
    FORMATOGRILLA(5, 4) = "TRUE"
    FORMATOGRILLA(5, 5) = "TRUE"
    FORMATOGRILLA(5, 6) = "TRUE"
    FORMATOGRILLA(5, 7) = "TRUE"
    FORMATOGRILLA(5, 8) = "TRUE"
    FORMATOGRILLA(5, 9) = "TRUE"
    
    Grid1.Cols = 10
    Grid1.Rows = 2
    
 Grid1.AllowUserResizing = False
    Grid1.DisplayFocusRect = False
Grid1.ExtendLastCol = True
    Grid1.BoldFixedCell = False
    Grid1.DrawMode = cellOwnerDraw
    
    Grid1.Appearance = Flat
    Grid1.ScrollBarStyle = Flat
    Grid1.FixedRowColStyle = Flat
    
'   Grid1.BackColorFixed = RGB(90, 158, 214)
'   Grid1.BackColorFixedSel = RGB(110, 180, 230)
'   Grid1.BackColorBkg = RGB(90, 158, 214)
'   Grid1.BackColorScrollBar = RGB(231, 235, 247)
'   Grid1.BackColor1 = RGB(231, 235, 247)
'   Grid1.BackColor2 = RGB(239, 243, 255)
'   Grid1.GridColor = RGB(148, 190, 231)
   Grid1.Column(0).Width = 0
    
    For k = 1 To Grid1.Cols - 1
        
        Grid1.Cell(0, k).text = FORMATOGRILLA(1, k)
        Grid1.Column(k).Width = Val(FORMATOGRILLA(2, k)) * Grid1.DefaultFont.Size
        
        
        Grid1.Column(k).MaxLength = Val(FORMATOGRILLA(2, k))
        Grid1.Column(k).FormatString = FORMATOGRILLA(4, k)
        Grid1.Column(k).Locked = FORMATOGRILLA(5, k)
        If FORMATOGRILLA(3, k) = "N" Then Grid1.Column(k).Alignment = cellRightCenter
        If FORMATOGRILLA(3, k) = "D" Then Grid1.Column(k).CellType = cellCalendar
        
    Next k
End Sub



Private Sub monto_Click()
End Sub

Private Sub leer()

Dim resultados As rdoResultset
    
    Dim csql As New rdoQuery
    Dim rut As String
    Dim PASO As String
    Dim fecha1 As String
    Dim fecha2 As String
    Dim LINEA As Double
    Dim total As Double
    Dim fechasum As String
    Dim total2 As Double
    
    LINEA = 0
 
        Set csql.ActiveConnection = contadb
        'cSql.SQL = "SELECT fecha,tipo,numero,linea,codigocuenta,glosacontable,tipodocumento,numerodocumento,fechadocumento,fechavencimiento,monto,dh,centrocosto,tipoctacte,rutctacte "
'        dia = 1
'        MES = 1
'        año = 2005
      fecha1 = Mid(desdefecha.Caption, 7, 4) + "-" + Mid(desdefecha.Caption, 4, 2) + "-" + Mid(desdefecha.Caption, 1, 2)
      fecha2 = Mid(hastafecha.Caption, 7, 4) + "-" + Mid(hastafecha.Caption, 4, 2) + "-" + Mid(hastafecha.Caption, 1, 2)
        csql.sql = "SELECT emision,tipocomprobante,numerocomprobante,numero,giradoa,vencimiento,fechacobro,monto "
        csql.sql = csql.sql + "FROM chequesdocumento where cuenta='" + dato1.text + dato2.text + dato3.text + "'and fechacobro >0 and vencimiento > 0 and fechacobro < vencimiento and emision >='" + fecha1 + "' and emision <='" + fecha2 + "' "
        csql.sql = csql.sql + "order by cuenta,vencimiento"
        csql.Execute
        total = 0
        total2 = 0
        Grid1.Rows = csql.RowsAffected + 1
        
        If csql.RowsAffected > 0 Then
        Set resultados = csql.OpenResultset
        fechasum = Format(resultados(2), "yyyy") + "/" + Format(resultados(2), "mm") + "/" + Format(resultados(2), "dd")
        
         While Not resultados.EOF
          LINEA = LINEA + 1
             
             Grid1.Cell(LINEA, 1).text = resultados(0)
             Grid1.Cell(LINEA, 2).text = resultados(1)
             Grid1.Cell(LINEA, 3).text = resultados(2)
             Grid1.Cell(LINEA, 4).text = resultados(3)
             Grid1.Cell(LINEA, 5).text = resultados(4)
             Grid1.Cell(LINEA, 6).text = resultados(5)
             Grid1.Cell(LINEA, 7).text = resultados(7)
             total = total + resultados(7)
             Grid1.Cell(LINEA, 8).text = total
             If Not IsNull(resultados(6)) = True Then
             Grid1.Cell(LINEA, 9).text = resultados(6)
             End If
             resultados.MoveNext
          If resultados.EOF = False Then
       
          End If
   
                   Wend
End If
            
     
 
End Sub
Sub limpiar()


End Sub


Sub CABEZAS2(titulo, tipo, FOLIO)
Dim objReportTitle As FlexCell.ReportTitle
Grid1.ReportTitles.Clear


Set objReportTitle = New FlexCell.ReportTitle
    objReportTitle.text = titulo
    objReportTitle.Font.Name = "arial"
    objReportTitle.Font.Size = 12
    objReportTitle.PrintOnAllPages = True
    Grid1.ReportTitles.Add objReportTitle

Set objReportTitle = New FlexCell.ReportTitle
    objReportTitle.text = nombrebanco
    
    objReportTitle.Font.Name = "arial"
    objReportTitle.Font.Size = 12
    objReportTitle.PrintOnAllPages = True
    Grid1.ReportTitles.Add objReportTitle
    
    'Report Title 1
    If tipo = "N" Then
        For k = 1 To 5
        Set objReportTitle = New FlexCell.ReportTitle
        objReportTitle.text = DATOSEMPRESA(k)
        objReportTitle.Font.Name = "arial"
        objReportTitle.Font.Size = 8
        objReportTitle.Font.Italic = True
        objReportTitle.PrintOnAllPages = True
        'objReportTitle.Color = RGB(128, 0, 0)
        objReportTitle.Align = CellLeft
        Grid1.ReportTitles.Add objReportTitle
    Next k
    Else
        For k = 1 To 4
        Set objReportTitle = New FlexCell.ReportTitle
        objReportTitle.text = ""
        objReportTitle.Font.Name = "arial"
        objReportTitle.Font.Size = 8
        objReportTitle.Font.Italic = True
        objReportTitle.PrintOnAllPages = True
        'objReportTitle.Color = RGB(128, 0, 0)
        objReportTitle.Align = CellLeft
        Grid1.ReportTitles.Add objReportTitle
        
        Next k
    Set objReportTitle = New FlexCell.ReportTitle
        
        
        
        
        
        objReportTitle.text = ""
        objReportTitle.Font.Name = "arial"
        objReportTitle.Font.Size = 8
        objReportTitle.Font.Italic = True
        objReportTitle.PrintOnAllPages = True
        'objReportTitle.Color = RGB(128, 0, 0)
        objReportTitle.Align = CellLeft
        Grid1.ReportTitles.Add objReportTitle
        
    End If
    
With Grid1.PageSetup
        
        If tipo = "N" Then .Header = "Pagina &P de &N Emitido: &D Usuario:" + USUARIOSISTEMA
        Rem If TIPO = "S" Then .Footer = "pagina &P"
        
        .HeaderAlignment = cellCenter
        
        .HeaderFont.Name = "Verdana"
        .HeaderFont.Size = 7
        .HeaderMargin = 2
        .TopMargin = 1
        .BottomMargin = 2
        
        
        
End With

End Sub


