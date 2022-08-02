VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{4777436C-EB8C-4596-98A8-EBCF98683969}#1.0#0"; "FlexCell.ocx"
Object = "{085FBDF3-00C5-421B-B762-1D57299A2B89}#1.0#0"; "CLBUTN.OCX"
Object = "{ADBBAED6-B16F-40DC-80DF-B44910CBA76C}#1.0#0"; "Frame-Xp.ocx"
Begin VB.Form banco07 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Listado Distribucion de Cheques"
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
   Begin XPFrame.FrameXp TIPOS 
      Height          =   2880
      Left            =   2475
      TabIndex        =   20
      Top             =   90
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   5080
      BackColor       =   16761024
      Caption         =   "DISTRIBUCION DE CHEQUES"
      CaptionEstilo3D =   1
      BackColor       =   16761024
      ForeColor       =   65535
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid GRILLATIPO 
         Height          =   2520
         Left            =   120
         TabIndex        =   21
         Top             =   240
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   4445
         _Version        =   393216
         BackColor       =   16107953
         ForeColor       =   16711680
         Rows            =   3
         FixedRows       =   0
         FixedCols       =   0
         ForeColorFixed  =   16777152
         BackColorBkg    =   16761024
         GridColor       =   16744576
         GridColorFixed  =   14282751
         GridColorUnpopulated=   14282751
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
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
      Left            =   0
      ScaleHeight     =   135
      ScaleWidth      =   135
      TabIndex        =   4
      Top             =   6120
      Visible         =   0   'False
      Width           =   135
   End
   Begin XPFrame.FrameXp FrameXp1 
      Height          =   8610
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   15150
      _ExtentX        =   26723
      _ExtentY        =   15187
      BackColor       =   16744576
      Caption         =   "LISTADO DISTRIBUCION DE CHEQUES"
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
      Begin XPFrame.FrameXp FrameQuickMenu 
         Height          =   615
         Left            =   12000
         TabIndex        =   27
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
            TabIndex        =   29
            Top             =   280
            Width           =   1335
         End
         Begin VB.CommandButton botonmisaccesos 
            Caption         =   "Permisos Modulo"
            Height          =   255
            Left            =   120
            TabIndex        =   28
            Top             =   280
            Width           =   1455
         End
      End
      Begin FlexCell.Grid Grid2 
         Height          =   240
         Left            =   360
         TabIndex        =   26
         Top             =   8190
         Visible         =   0   'False
         Width           =   1230
         _ExtentX        =   2170
         _ExtentY        =   423
         Cols            =   5
         DefaultFontSize =   8.25
         Rows            =   30
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00FFC0C0&
         Caption         =   "ETIQUETAS"
         Height          =   330
         Left            =   7020
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   8190
         Width           =   2130
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FFC0C0&
         Caption         =   "IMPRIMIR"
         Height          =   330
         Left            =   4320
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   8190
         Width           =   2130
      End
      Begin XPFrame.FrameXp FrameXp3 
         Height          =   1290
         Left            =   135
         TabIndex        =   7
         Top             =   225
         Width           =   14910
         _ExtentX        =   26300
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
         Begin VB.TextBox DATO5 
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
            Left            =   1710
            MaxLength       =   1
            TabIndex        =   22
            Tag             =   "nombre"
            Top             =   900
            Width           =   270
         End
         Begin VB.CommandButton ver 
            Caption         =   "Ver"
            Height          =   120
            Left            =   13440
            TabIndex        =   11
            Top             =   1125
            Visible         =   0   'False
            Width           =   105
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
            Left            =   945
            MaxLength       =   4
            TabIndex        =   3
            Top             =   540
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
            Left            =   540
            MaxLength       =   2
            TabIndex        =   2
            Top             =   540
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
            Left            =   135
            MaxLength       =   2
            TabIndex        =   1
            Tag             =   "codigo"
            Top             =   540
            Width           =   375
         End
         Begin XPFrame.FrameXp FrameXp8 
            Height          =   975
            Left            =   8280
            TabIndex        =   13
            Top             =   360
            Width           =   6375
            _ExtentX        =   11245
            _ExtentY        =   1720
            BackColor       =   14737632
            Caption         =   "Rangos de Fecha"
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
            Alignment       =   1
            Begin CoolButtons.cool_Button cool_Button3 
               Height          =   375
               Left            =   4320
               TabIndex        =   14
               Top             =   360
               Width           =   1875
               _ExtentX        =   3307
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
         Begin VB.Label lbldis 
            BackColor       =   &H00000000&
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
            ForeColor       =   &H0000FFFF&
            Height          =   330
            Left            =   2115
            TabIndex        =   24
            Top             =   900
            Width           =   5670
         End
         Begin VB.Label Label7 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Distribucion"
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
            Left            =   135
            TabIndex        =   23
            Top             =   900
            Width           =   1455
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
            Left            =   1710
            TabIndex        =   10
            Top             =   270
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
            Left            =   1710
            TabIndex        =   9
            Top             =   540
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
            Left            =   135
            TabIndex        =   8
            Top             =   270
            Width           =   1545
         End
      End
      Begin XPFrame.FrameXp FrameXp2 
         Height          =   6630
         Left            =   135
         TabIndex        =   6
         Top             =   1530
         Width           =   15000
         _ExtentX        =   26458
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
            Top             =   240
            Width           =   14955
            _ExtentX        =   26379
            _ExtentY        =   11298
            Cols            =   5
            DefaultFontSize =   8.25
            Rows            =   30
         End
      End
   End
End
Attribute VB_Name = "banco07"
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

Private Sub COMMAND2_Click()
Dim contador As Double
Dim rutctacte As String
Dim fechacheque As String
Dim csql As New rdoQuery
Dim resultados As rdoResultset
CARGAGRILLA2
Grid2.PageSetup.PrintFixedRow = True
Grid2.PageSetup.BottomMargin = 1.5
Grid2.PageSetup.TopMargin = 0
Grid2.PageSetup.LeftMargin = 0.5
Grid2.PageSetup.RightMargin = 0
Grid2.PageSetup.BlackAndWhite = True
Grid2.PageSetup.PrintGridlines = False


 Grid2.Rows = Grid2.Rows + 4
 
 contador = 0
For k = 1 To Grid1.Rows - 1
    Set csql.ActiveConnection = contadb
    csql.sql = "select tipo,numero,fecha "
    csql.sql = csql.sql & "From movimientoscontables "
    csql.sql = csql.sql & "where numerodocumento='" & Grid1.Cell(k, 4).text & "' and codigocuenta='" & dato1.text & dato2.text & dato3.text & "' "
    csql.Execute

    If csql.RowsAffected > 0 Then
        Set resultados = csql.OpenResultset
        fechacheque = resultados(2)
        rutctacte = leerrutctacte(resultados(0), resultados(1), resultados(2))
    End If
    csql.Close
    Set csql = Nothing
    Set resultados = Nothing
   
    Set csql.ActiveConnection = contadb
   
    csql.sql = "select nombre,direccion,comuna,ciudad from cuentascorrientes where "
    csql.sql = csql.sql & "tipo='" & CUENTAPROVEEDOR & "' and rut='" & rutctacte & "' and año='" & Format(fechacheque, "yyyy") & "'"
    csql.Execute
    If csql.RowsAffected > 0 Then
        Set resultados = csql.OpenResultset
        Grid2.Rows = Grid2.Rows + 1
        Grid2.Cell(Grid2.Rows - 1, 1).text = "Nombre    : " & resultados(0)
        Grid2.Rows = Grid2.Rows + 1
        Grid2.Cell(Grid2.Rows - 1, 1).text = "Direccion : " & resultados(1)
        Grid2.Rows = Grid2.Rows + 1
        Grid2.Cell(Grid2.Rows - 1, 1).text = "Comuna    : " & resultados(2)
        Grid2.Rows = Grid2.Rows + 1
        Grid2.Cell(Grid2.Rows - 1, 1).text = "Ciudad    : " & resultados(3)
       
        Grid2.Rows = Grid2.Rows + 1
        Grid2.RowHeight(Grid2.Rows - 1) = 24
    End If
    csql.Close
    Set csql = Nothing
    Set resultados = Nothing
    contador = contador + 1
    If contador = 10 Then
        Grid2.Rows = Grid2.Rows + 4
        contador = 0
    End If
Next k

'Grid2.PageSetup.PaperWidth = 10
'Grid2.PageSetup.PaperSize = 10
Grid2.PrintPreview





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
dato5.SetFocus

End If

End Sub


Private Sub dato5_GotFocus()
TIPOS.Visible = True

End Sub

Private Sub dato5_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 And dato5.text < "9" And dato5.text <> "" Then
leer
lbldis.Caption = GRILLATIPO.TextMatrix(CDbl(dato5.text), 1)
TIPOS.Visible = False
Command1.SetFocus

End If

End Sub

Private Sub Form_Load()
CENTRAR Me
TIPOS.Visible = False


    
    Call Conectar_BD

    sc = 0
CARGAGRILLA
GRILLATIPOS


desdefecha.Caption = fechasistema
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

titulo = lbldis.Caption


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
    
    Grid1.Cols = 8
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
Sub CARGAGRILLA2()
Rem DATOS DE LA COLUMNA
    Dim FORMATOGRILLA(10, 10)
    Grid1.DefaultFont.Size = 9
       
    
    Grid2.Cols = 2
    Grid2.Rows = 1
    
    Grid2.AllowUserResizing = False
    Grid2.DisplayFocusRect = False
    Grid2.ExtendLastCol = True
    Grid2.BoldFixedCell = False
    Grid2.DrawMode = cellOwnerDraw
    
    Grid2.Appearance = Flat
    Grid2.ScrollBarStyle = Flat
    Grid2.FixedRowColStyle = Flat
    
   Grid2.Column(0).Width = 0
   Grid2.Column(1).Width = 400
   
    
    
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
    Dim TOTAL As Double
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
'        csql.sql = csql.sql + "FROM chequesdocumento where cuenta='" + dato1.text + dato2.text + dato3.text + "'and emision >='" + fecha1 + "' and emision <='" + fecha2 + "' and ubicacion='" + DATO5.text + "' "
        csql.sql = csql.sql + "FROM chequesdocumento where cuenta='" + dato1.text + dato2.text + dato3.text + "'and fechamovimiento >='" + fecha1 + "' and fechamovimiento <='" + fecha2 + "' and ubicacion='" + dato5.text + "' "
 
        csql.sql = csql.sql + "order by cuenta,fechamovimiento"
        csql.Execute
        TOTAL = 0
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


Sub GRILLATIPOS()
GRILLATIPO.Cols = 2
GRILLATIPO.Rows = 10
GRILLATIPO.ColWidth(0) = 200 * 2
GRILLATIPO.ColWidth(1) = 200 * 10

GRILLATIPO.TextMatrix(0, 0) = "0"
GRILLATIPO.TextMatrix(1, 0) = "1"
GRILLATIPO.TextMatrix(2, 0) = "2"
GRILLATIPO.TextMatrix(3, 0) = "3"
GRILLATIPO.TextMatrix(4, 0) = "4"
GRILLATIPO.TextMatrix(5, 0) = "5"
GRILLATIPO.TextMatrix(6, 0) = "6"
GRILLATIPO.TextMatrix(7, 0) = "7"
GRILLATIPO.TextMatrix(8, 0) = "8"
GRILLATIPO.TextMatrix(9, 0) = "9"

GRILLATIPO.TextMatrix(0, 1) = "EMITIDO SIN FIRMA"
GRILLATIPO.TextMatrix(1, 1) = "ENVIADO POR CORREO"
GRILLATIPO.TextMatrix(2, 1) = "EN PODER SECRETARIA"
GRILLATIPO.TextMatrix(3, 1) = "RETENIDO POR PUBLICIDAD"
GRILLATIPO.TextMatrix(4, 1) = "RETENIDO POR GERENCIA"
GRILLATIPO.TextMatrix(5, 1) = "ENTREGADO A PROVEEDOR"
GRILLATIPO.TextMatrix(6, 1) = "GIRADO A FECHA   "
GRILLATIPO.TextMatrix(7, 1) = "ENVIADO BUSES JAC"
GRILLATIPO.TextMatrix(8, 1) = "ENVIADO BUSES TURBUS"
GRILLATIPO.TextMatrix(9, 1) = "DEPOSITO EN CTA CTE "

CANDO = 9



End Sub

