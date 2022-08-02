VERSION 5.00
Object = "{4777436C-EB8C-4596-98A8-EBCF98683969}#1.0#0"; "FlexCell.ocx"
Object = "{ADBBAED6-B16F-40DC-80DF-B44910CBA76C}#1.0#0"; "Frame-Xp.ocx"
Begin VB.Form factoring02 
   Appearance      =   0  'Flat
   BackColor       =   &H0000FF00&
   Caption         =   "Cancelacion Ordenes de Compra"
   ClientHeight    =   8655
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14895
   DrawStyle       =   5  'Transparent
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   Moveable        =   0   'False
   ScaleHeight     =   577
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   993
   StartUpPosition =   2  'CenterScreen
   Begin XPFrame.FrameXp FrameQuickMenu 
      Height          =   615
      Left            =   11760
      TabIndex        =   27
      Top             =   0
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   1085
      BackColor       =   8454016
      Caption         =   " Mis Datos"
      BackColor       =   8454016
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
         TabIndex        =   29
         Top             =   280
         Width           =   1455
      End
      Begin VB.CommandButton botonmisfavoritos 
         Caption         =   "Mis Favoritos"
         Height          =   255
         Left            =   1680
         TabIndex        =   28
         Top             =   280
         Width           =   1335
      End
   End
   Begin XPFrame.FrameXp frmcheque 
      Height          =   2535
      Left            =   3960
      TabIndex        =   16
      Top             =   4320
      Width           =   5325
      _ExtentX        =   9393
      _ExtentY        =   4471
      BackColor       =   16761024
      Caption         =   "PANTALLA DATOS DEL CHEQUE"
      CaptionEstilo3D =   1
      BackColor       =   16761024
      ForeColor       =   65535
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   1
      Begin VB.TextBox dato4 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   1800
         MaxLength       =   10
         TabIndex        =   23
         Top             =   1530
         Width           =   1815
      End
      Begin VB.CommandButton Command5 
         BackColor       =   &H0080FFFF&
         Caption         =   "INICIAR PROCESO"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1395
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   2070
         Width           =   2535
      End
      Begin VB.TextBox dato3 
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
         Left            =   3420
         MaxLength       =   4
         TabIndex        =   19
         Top             =   405
         Width           =   735
      End
      Begin VB.TextBox dato2 
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
         Left            =   2940
         MaxLength       =   2
         TabIndex        =   18
         Top             =   405
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
         Left            =   2520
         MaxLength       =   2
         TabIndex        =   17
         Tag             =   "codigo"
         Top             =   405
         Width           =   375
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "NUMERO INICIAL"
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
         Height          =   240
         Left            =   1800
         TabIndex        =   24
         Top             =   1260
         Width           =   1815
      End
      Begin VB.Label lblBanco 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
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
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   90
         TabIndex        =   22
         Top             =   810
         Width           =   5145
      End
      Begin VB.Label Label1 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " BANCO"
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
         Left            =   990
         TabIndex        =   20
         Top             =   390
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
   Begin XPFrame.FrameXp FrameXp1 
      Height          =   8610
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   14865
      _ExtentX        =   26220
      _ExtentY        =   15187
      BackColor       =   8454016
      Caption         =   "PANTALLA PAGO DE ORDENES DECOMPRA"
      CaptionEstilo3D =   1
      BackColor       =   8454016
      ForeColor       =   65535
      ColorBarraArriba=   12648384
      ColorBarraAbajo =   16384
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
      Begin VB.TextBox ORDEN 
         BackColor       =   &H00FFC0C0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   13050
         MaxLength       =   10
         TabIndex        =   26
         Top             =   8235
         Width           =   1500
      End
      Begin VB.CommandButton BUSCAR 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Busca Orden"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   11565
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   8235
         Width           =   1320
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H0080FFFF&
         Caption         =   "Generar Pagos Electronicos"
         Height          =   330
         Left            =   8190
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   8190
         Visible         =   0   'False
         Width           =   2535
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H0080FFFF&
         Caption         =   "IMPRIMIR"
         Height          =   330
         Left            =   2205
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   8190
         Width           =   2130
      End
      Begin XPFrame.FrameXp FrameXp3 
         Height          =   1050
         Left            =   135
         TabIndex        =   5
         Top             =   360
         Width           =   14640
         _ExtentX        =   25823
         _ExtentY        =   1852
         BackColor       =   8454016
         Caption         =   "DATOS DE FILTRADO"
         CaptionEstilo3D =   1
         BackColor       =   8454016
         ForeColor       =   8438015
         ColorBarraArriba=   12648384
         ColorBarraAbajo =   16384
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Begin VB.CommandButton Command2 
            Caption         =   "LISTAR"
            Height          =   285
            Left            =   11970
            TabIndex        =   7
            Top             =   675
            Width           =   1455
         End
         Begin XPFrame.FrameXp FrameXp6 
            Height          =   675
            Left            =   90
            TabIndex        =   9
            Top             =   270
            Width           =   3300
            _ExtentX        =   5821
            _ExtentY        =   1191
            BackColor       =   8454016
            Caption         =   "MES"
            CaptionEstilo3D =   1
            BackColor       =   8454016
            ForeColor       =   65535
            ColorBarraArriba=   16384
            ColorBarraAbajo =   8454016
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Begin VB.ComboBox COMBOMES 
               Height          =   315
               Left            =   45
               TabIndex        =   10
               Top             =   270
               Width           =   3180
            End
         End
         Begin XPFrame.FrameXp FrameXp7 
            Height          =   675
            Left            =   3510
            TabIndex        =   11
            Top             =   270
            Width           =   3075
            _ExtentX        =   5424
            _ExtentY        =   1191
            BackColor       =   8454016
            Caption         =   "AÑO"
            CaptionEstilo3D =   1
            BackColor       =   8454016
            ForeColor       =   65535
            ColorBarraArriba=   16384
            ColorBarraAbajo =   8454016
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Begin VB.ComboBox COMBOAÑO 
               Height          =   315
               Left            =   90
               TabIndex        =   12
               Top             =   270
               Width           =   2865
            End
         End
         Begin XPFrame.FrameXp FrameXp4 
            Height          =   675
            Left            =   6705
            TabIndex        =   13
            Top             =   270
            Visible         =   0   'False
            Width           =   4560
            _ExtentX        =   8043
            _ExtentY        =   1191
            BackColor       =   8454016
            Caption         =   "LOCAL"
            CaptionEstilo3D =   1
            BackColor       =   8454016
            ForeColor       =   65535
            ColorBarraArriba=   16384
            ColorBarraAbajo =   8454016
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Begin VB.ComboBox ComboLOCAL 
               Height          =   315
               Left            =   90
               TabIndex        =   14
               Top             =   270
               Width           =   4395
            End
         End
      End
      Begin XPFrame.FrameXp FrameXp2 
         Height          =   6675
         Left            =   135
         TabIndex        =   3
         Top             =   1485
         Width           =   14685
         _ExtentX        =   25903
         _ExtentY        =   11774
         BackColor       =   8454016
         Caption         =   "LISTADO ORDENES A CANCELAR"
         CaptionEstilo3D =   1
         BackColor       =   8454016
         ForeColor       =   8438015
         ColorBarraArriba=   12648384
         ColorBarraAbajo =   16384
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
            Height          =   6360
            Left            =   0
            TabIndex        =   4
            Top             =   225
            Width           =   14595
            _ExtentX        =   25744
            _ExtentY        =   11218
            BackColorFixed  =   8454016
            Cols            =   5
            DefaultFontSize =   8.25
            GridColor       =   32768
            Rows            =   30
            DateFormat      =   2
         End
      End
      Begin VB.CommandButton Command4 
         BackColor       =   &H0080FFFF&
         Caption         =   "Generar Pagos Cheques"
         Height          =   330
         Left            =   4995
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   8190
         Width           =   2535
      End
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
      TabIndex        =   1
      Top             =   6120
      Visible         =   0   'False
      Width           =   135
   End
End
Attribute VB_Name = "factoring02"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private lineassiguientes As Double

Private localfiltro As String
Private pagada As Boolean
Private NCHEQUE As Double
Private rutreal As String
Private DIFERENCIA As Double
Private contabilizada As String
Private lineafinal As Double
Private numerocontable As String
Private FECHACONTABLE As String
Private lineacontable As Double
Private rutcontable As String
Private tipocontable As String
Private TOTALCheque As Double
Private fechacheque As String
Private NOMBREGIRADO As String
Private montoharina As Double
Private montocarne As Double





Private Sub Command1_Click()
imprimir
End Sub



Private Sub COMMAND2_Click()
localfiltro = Mid(ComboLOCAL.text, 1, 2)
año = COMBOAÑO.text
MES = COMBOMES.ListIndex + 1

leer


End Sub



Private Sub Command3_Click()
Dim k As Double
Dim rutprove As String
Dim tipo As String
Dim ordenes As Double
lineacontable = 15
For k = 1 To Grid1.Rows - 1
    If Grid1.Cell(k, 14).text = "1" Then
        If rutprove <> Mid(Grid1.Cell(k, 3).text, 1, 9) + Mid(Grid1.Cell(k, 3).text, 11, 1) Or lineacontable > 10 Then
            If TOTALCheque <> 0 Then
                NOMBREGIRADO = Grid1.Cell(k, 4).text
            Call grabarcheque(TOTALCheque)
            End If
        fechacheque = Format(Grid1.Cell(k, 15).text, "yyyy-mm-dd")
        NOMBREGIRADO = Grid1.Cell(k, 4).text
        FECHACONTABLE = Format(fechasistema, "yyyy-mm-dd")
        tipocontable = "NG"
        numerocontable = LEERFOLIOCE("NG")
        lineacontable = 0
        rutprove = Mid(Grid1.Cell(k, 3).text, 1, 9) + Mid(Grid1.Cell(k, 3).text, 11, 1)
        rutcontable = rutprove
        TOTALCheque = 0
        End If
        
        Call GRABARCOMPROBANTE(Grid1.Cell(k, 2).text, Grid1.Cell(k, 9).text, Grid1.Cell(k, 10).text, Format(Grid1.Cell(k, 15).text, "yyyy-mm-dd"), Grid1.Cell(k, 1).text, Grid1.Cell(k, 16).text, rutcontable, Format(Grid1.Cell(k, 5).text, "yyyy-mm-dd"))
        
    End If
Next k
        If TOTALCheque <> 0 Then
        Call grabarcheque(TOTALCheque)
        TOTALCheque = 0
        End If
        
leer


End Sub

Private Sub Command4_Click()
frmcheque.Visible = True
dato1.SetFocus

End Sub

Private Sub Command5_Click()


Dim k As Double
Dim rutprove As String
Dim tipo As String

If lblBanco.Caption <> "" Then

NCHEQUE = CDbl(dato4.text) - 1



For k = 1 To Grid1.Rows - 1
    If Grid1.Cell(k, 14).text = "1" Then
        If rutprove <> Mid(Grid1.Cell(k, 3).text, 1, 9) + Mid(Grid1.Cell(k, 3).text, 11, 1) Then
        If TOTALCheque <> 0 Then
        Call grabarcheque(TOTALCheque)
        End If
        fechacheque = Format(Grid1.Cell(k, 15).text, "yyyy-mm-dd")
        NOMBREGIRADO = Grid1.Cell(k, 17).text
        FECHACONTABLE = Format(fechasistema, "yyyy-mm-dd")
        tipocontable = "PF"
        numerocontable = LEERFOLIOCE("PF")
        lineacontable = 0
        rutprove = Mid(Grid1.Cell(k, 3).text, 1, 9) + Mid(Grid1.Cell(k, 3).text, 11, 1)
        rutcontable = rutprove
        TOTALCheque = 0
        End If
        
        Call GRABARCOMPROBANTE(Grid1.Cell(k, 2).text, Grid1.Cell(k, 9).text, Grid1.Cell(k, 10).text, Format(Grid1.Cell(k, 15).text, "yyyy-mm-dd"), Grid1.Cell(k, 1).text, "CHEQUE", rutprove, Format(Grid1.Cell(k, 5).text, "yyyy-mm-dd"))
        
    End If
Next k
        If TOTALCheque <> 0 Then
        Call grabarcheque(TOTALCheque)
        TOTALCheque = 0
        End If
        
leer

End If
frmcheque.Visible = False
dato4.text = ""
End Sub

Private Sub dato4_KeyPress(KeyAscii As Integer)
KeyAscii = esNumero(KeyAscii)
If KeyAscii = 13 Then
Call ceros(dato4)
If leercheque(dato1.text + dato2.text + dato3.text, dato4.text) = True Then
MsgBox ("EL NUMERO DE CHEQUE YA ESTA EMITIDO")
dato4.text = ""
dato4.SetFocus
Else

Command5.SetFocus
End If

End If

End Sub

Private Sub Form_Load()
'CENTRAR Me
    Call Conectar_BD
    sc = 0
CARGAGRILLA
Call Conectarventas(Servidor, clientesistema + "ventas00", Usuario, password)
Call Conectargestion(Servidor, clientesistema + "gestion", Usuario, password)
Call Conectargestionrubro(Servidor, clientesistema + "gestion00", Usuario, password)

For k = 1 To 12
COMBOMES.AddItem MonthName(k)
Next k
COMBOMES.ListIndex = CDbl(Format(fechasistema, "mm") - 1)
For k = 2000 To Val(Format(fechasistema, "yyyy"))
COMBOAÑO.AddItem k
Next k
COMBOAÑO.ListIndex = k - 2001
LEErlocales
frmcheque.Visible = False


End Sub

Sub Pregunta(ByRef caja As TextBox, ByRef caja1 As TextBox)
    If caja.text = "" And sc = 0 Then caja.SetFocus
    If caja.text <> "" Or sc = 1 Then caja1.Enabled = True: caja1.SetFocus
    sc = 0
End Sub
Sub flechas(ByRef caja As TextBox, ByRef caja1 As TextBox, ByRef codigo As Integer)
    If codigo = 38 And caja.Enabled = True Then caja.SetFocus
    If codigo = 40 And caja1.Enabled = True Then caja1.SetFocus
End Sub

Private Sub lblhistorico_Click(Index As Integer)

End Sub

Private Sub Label16_Click()
End Sub

Sub limpia()
    
    
End Sub

Sub imprimir()
Dim titulo As String
titulo = "ORDENES PENDIENTES DE PAGO " + COMBOMES.text + " " + COMBOAÑO.text
Call CABEZAS2(titulo, "N", "000000000")
Grid1.Range(0, 1, 0, Grid1.Cols - 1).Borders(cellEdgeBottom) = cellThick
Grid1.Range(0, 1, 0, Grid1.Cols - 1).Borders(cellEdgeLeft) = cellThick
Grid1.Range(0, 1, 0, Grid1.Cols - 1).Borders(cellEdgeTop) = cellThick
Grid1.Range(0, 1, 0, Grid1.Cols - 1).Borders(cellEdgeRight) = cellThick
Grid1.Range(0, 1, 0, Grid1.Cols - 1).Borders(cellInsideHorizontal) = cellThick
Grid1.Range(0, 1, 0, Grid1.Cols - 1).Borders(cellInsideVertical) = cellThick
Grid1.DefaultFont.Size = 8
Grid1.PageSetup.Orientation = cellLandscape

Grid1.PageSetup.PrintFixedRow = True
Grid1.PageSetup.BottomMargin = 2
Grid1.PageSetup.TopMargin = 1
Grid1.PageSetup.LeftMargin = 1
Grid1.PageSetup.RightMargin = 0
Grid1.PageSetup.BlackAndWhite = True
Grid1.PageSetup.PrintGridlines = False
Grid1.PrintPreview 100

   
End Sub
Sub grilla()
    
End Sub




Private Sub opciones_GotFocus()

MANUAL.SetFocus

End Sub
Sub CARGAGRILLA()
Rem DATOS DE LA COLUMNA
    Dim FORMATOGRILLA(10, 20)
    Grid1.DefaultFont.Size = 8
      
    FORMATOGRILLA(1, 1) = "TP"
    FORMATOGRILLA(1, 2) = "NUMERO"
    FORMATOGRILLA(1, 3) = "RUT"
    FORMATOGRILLA(1, 4) = "PROVEEDOR"
    FORMATOGRILLA(1, 5) = "FECHA"
    FORMATOGRILLA(1, 6) = "TOTAL"
    FORMATOGRILLA(1, 7) = ""
    FORMATOGRILLA(1, 8) = ""
    FORMATOGRILLA(1, 9) = "PAGAR"
    FORMATOGRILLA(1, 10) = ""
    FORMATOGRILLA(1, 11) = ""
    FORMATOGRILLA(1, 12) = ""
    FORMATOGRILLA(1, 13) = "OK"
    FORMATOGRILLA(1, 14) = "PAGA"
    FORMATOGRILLA(1, 15) = "FECHA PAGO"
    FORMATOGRILLA(1, 16) = "CESIONARIO"
    FORMATOGRILLA(1, 17) = "NOMBRE"
  
    
    Rem LARGO DE LOS DATOS
    FORMATOGRILLA(2, 1) = "6"
    FORMATOGRILLA(2, 2) = "9"
    FORMATOGRILLA(2, 3) = "9"
    FORMATOGRILLA(2, 4) = "20"
    FORMATOGRILLA(2, 5) = "8"
    FORMATOGRILLA(2, 6) = "8"
    FORMATOGRILLA(2, 7) = "0"
    FORMATOGRILLA(2, 8) = "0"
    FORMATOGRILLA(2, 9) = "8"
    FORMATOGRILLA(2, 10) = "0"
    FORMATOGRILLA(2, 11) = "0"
    FORMATOGRILLA(2, 12) = "0"
    FORMATOGRILLA(2, 13) = "3"
    FORMATOGRILLA(2, 14) = "5"
    FORMATOGRILLA(2, 15) = "10"
    FORMATOGRILLA(2, 16) = "9"
    FORMATOGRILLA(2, 17) = "20"
    
    Rem TIPO DE DATOS C=CEROS,S=STRING,N=NUMERICO,D=DATE,
    FORMATOGRILLA(3, 1) = "S"
    FORMATOGRILLA(3, 2) = "S"
    FORMATOGRILLA(3, 3) = "S"
    FORMATOGRILLA(3, 4) = "S"
    FORMATOGRILLA(3, 5) = "S"
    FORMATOGRILLA(3, 6) = "N"
    FORMATOGRILLA(3, 7) = "N"
    FORMATOGRILLA(3, 8) = "N"
    FORMATOGRILLA(3, 9) = "N"
    FORMATOGRILLA(3, 10) = "N"
    FORMATOGRILLA(3, 11) = "S"
    FORMATOGRILLA(3, 12) = "S"
    FORMATOGRILLA(3, 13) = "S"
    FORMATOGRILLA(3, 14) = "S"
    FORMATOGRILLA(3, 15) = "S"
    FORMATOGRILLA(3, 16) = "S"
    FORMATOGRILLA(3, 17) = "S"
    
    Rem FORMATO GRILLA
    FORMATOGRILLA(4, 6) = "##,###,##0"
    FORMATOGRILLA(4, 7) = "##,###,##0"
    FORMATOGRILLA(4, 8) = "##,###,##0"
    FORMATOGRILLA(4, 9) = "##,###,##0"
    FORMATOGRILLA(4, 10) = "##,###,##0"
    
    Rem LOCCKED
    For k = 1 To 17
    FORMATOGRILLA(5, k) = "TRUE"
    
    Next k
    
  
    FORMATOGRILLA(5, 14) = "FALSE"
    FORMATOGRILLA(5, 15) = "FALSE"
    FORMATOGRILLA(5, 16) = "FALSE"
    
    Grid1.Cols = 18
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
    
    Grid1.Column(11).CellType = cellCheckBox
    Grid1.Column(12).CellType = cellCheckBox
    Grid1.Column(13).CellType = cellCheckBox
    Grid1.Column(14).CellType = cellCheckBox
    Grid1.Column(15).CellType = cellCalendar
    Grid1.Column(17).CellType = cellTextBox
    Grid1.Column(1).Locked = False
    Grid1.Column(1).CellType = cellComboBox
    
    
    
    With Grid1.ComboBox(1)
        '.Locked = True
        .AutoComplete = True
        .Font.Name = "Courier New"
        .AddItem "CHEQUES"
'        .AddItem "TRANSFERENCIA"
        
    End With
    
    
End Sub



Private Sub monto_Click()
End Sub

Public Sub leer()

Dim resultados As rdoResultset
    Dim csql As New rdoQuery
    Dim rut As String
    Dim PASO As String
    Dim fecha1 As String
    Dim fecha2 As String
    Dim LINEA As Double
    Dim TOTAL As Double
    Dim fec As Double
    Dim fec1 As Double
    Dim fechasum As String
    Dim total2 As Double
    Dim montofacturas As Double
    Dim OTROS As Double
    Dim apagar As Double
    Dim saldoctacte As String
    Dim sipu As String
    Dim DEVOLU As Boolean
    Dim glosa As String
    Dim RUTEMPRE As String
    Dim tipo As String
    Dim loc As String
    
    
    LINEA = 0: fec = 0: fec1 = 0
    fecha1 = año + "-" + MES + "-" + "01"
    fecha2 = año + "-" + MES + "-" + "31"
    
'      empresa_fae
      loc = leerdatos(conta, "maestroempresas", "empresafae", "codigoempresa='" + empresaactiva + "' ")
      
      
        Set csql.ActiveConnection = contadb
'        csql.sql = "SELECT mc.tipo,mc.numero,mc.fecha,mc.total,mc.rut,autorizacancelacion,fechaautorizacionpago "
'        csql.sql = csql.sql + "FROM facturasdecompras as mc "
''        csql.sql = csql.sql + " where (mc.tipo='1' OR mc.tipo='4') AND mc.fecha between '" + fecha1 + "' AND '" + fecha2 + "' and autorizacancelacion='1' "
'         csql.sql = csql.sql + " where   mc.fecha between '" + fecha1 + "' AND '" + fecha2 + "' and autorizacancelacion='1' "
''        csql.sql = csql.sql + " order by mc.rut,mc.fecha "
        
        csql.sql = "SELECT  tipo,LPAD(numero,10,0),LPAD(REPLACE(cedente_rut,'-',''),10,0) AS rut,cedente_nombre,fecha,cesion_monto,"
        csql.sql = csql.sql & "cesion_monto,LPAD(REPLACE(cesionario_rut,'-',''),10,0) AS cesionario_rut,UCASE(cesionario_nombre),autorizacionpago,fechaautorizacionpago "
        csql.sql = csql.sql & "FROM " & clientesistema & "fae" & loc & ".sv_dte_cedidos_" & loc & " as dte WHERE fecha BETWEEN '" & fecha1 & "' AND '" & fecha2 & "' "
        
        
        csql.Execute
        TOTAL = 0
        total2 = 0
        Grid1.Rows = 1
        Grid1.AutoRedraw = False
        
        
        If csql.RowsAffected > 0 Then
        Set resultados = csql.OpenResultset
        fechasum = Format(fechasistema, "yyyy") + "/" + Format(fechasistema, "mm") + "/" + Format(fechasistema, "dd")
        
         While Not resultados.EOF
            
            rutreal = resultados(2)
            pagada = False
            
            
            If resultados(0) = "1" Or resultados(0) = "4" Then tipo = "FC"
            If resultados(0) = "2" Or resultados(0) = "5" Then tipo = "ND"
            If resultados(0) = "3" Or resultados(0) = "6" Then tipo = "NC"
            If resultados(0) = "0" Or resultados(0) = "0" Then tipo = "EE"
            If resultados(0) = "9" Or resultados(0) = "9" Then tipo = "EX"
            If resultados(0) = "33" Then tipo = "FC"
            
            

           montofacturas = leerabonofacturaFactoring(tipo, resultados(1), resultados(2), CUENTAPROVEEDOR, "D", resultados(4))
            If montofacturas = 0 Then
            
             Grid1.Rows = Grid1.Rows + 1
             LINEA = LINEA + 1
             Grid1.Cell(LINEA, 1).text = resultados(0)
             Rem Grid1.Cell(linea, 1).text = "CHEQUES"
             Grid1.Cell(LINEA, 2).text = resultados(1)
             
             Grid1.Cell(LINEA, 3).text = Mid(rutreal, 1, 9) + "-" + Mid(rutreal, 10, 1)
             If resultados(0) <> "BH" Then
             Grid1.Cell(LINEA, 4).text = leerdatos(contadb, "cuentascorrientes", "nombre", "rut='" + rutreal + "' and tipo='" + CUENTAPROVEEDOR + "' ")
             Else
             Grid1.Cell(LINEA, 4).text = leerdatos(contadb, "cuentascorrientes", "nombre", "rut='" + rutreal + "' and tipo='23100029' and año='" & Format(fechasistema, "yyyy") & "'")
             End If
             Grid1.Cell(LINEA, 5).text = resultados(4)
             Grid1.Cell(LINEA, 6).text = resultados(5)
'             OTROS = totalpagosotros(Mid(resultados(0), 1, 1), resultados(1))
              OTROS = 0
             Grid1.Cell(LINEA, 8).text = OTROS
'             Grid1.Cell(LINEA, 9).text = resultados(3) - OTROS
             
            
             
             apagar = resultados(5) + OTROS
             DIFERENCIA = 0
             
             
             
             Grid1.Cell(LINEA, 9).text = apagar
             
             Grid1.Cell(LINEA, 10).text = DIFERENCIA
             
             Grid1.Cell(LINEA, 11).text = "0"
             
             Dim ANTICIPO As Double
    
'             ANTICIPO = leersaldoctacte("23100027", rutreal, fechasistema)
'             If ANTICIPO = 0 Then
             Grid1.Cell(LINEA, 12).text = "0"
'            Else
'             Grid1.Cell(LINEA, 12).text = "1"
'
'            End If
            Grid1.Cell(LINEA, 16).text = Mid(resultados(7), 1, 9) & "-" & Mid(resultados(7), 10, 1)
            Grid1.Cell(LINEA, 17).text = resultados(8)
            
             
            
               
             Grid1.Cell(LINEA, 13).text = "1"
             Grid1.Cell(LINEA, 14).text = resultados(9)
                If IsNull(resultados(10)) = False Then
                    Grid1.Cell(LINEA, 15).text = resultados(10)
                End If
                 If IsNull(resultados(10)) = True And resultados(5) = "1" Then
                    Grid1.Cell(LINEA, 15).text = fechasistema
                 End If
                 
                 
                 
'            If ANTICIPO <> 0 And OTROS = 0 Then
'                Grid1.Cell(LINEA, 13).text = "0"
'                Grid1.Cell(LINEA, 14).text = "0"
'                Grid1.Cell(LINEA, 15).text = ""
'            End If
            
            
            End If
            resultados.MoveNext
                        
            Wend
      End If
      Grid1.AutoRedraw = True
      Grid1.Refresh
      
      
      
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
    objReportTitle.text = ComboLOCAL.text
    objReportTitle.Font.Name = "arial"
    objReportTitle.Font.Size = 12
    objReportTitle.PrintOnAllPages = True
    Grid1.ReportTitles.Add objReportTitle

    'Report Title 1
    If tipo = "N" Then
        For k = 1 To 4
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
        .TopMargin = 2
        .BottomMargin = 1
        
        
        
End With

End Sub


Sub LEErlocales()
    Dim resultados As rdoResultset
    Dim csql As New rdoQuery
    
        Set csql.ActiveConnection = gestion
        csql.sql = "SELECT codigo,nombre "
        csql.sql = csql.sql + "FROM g_maestroempresas WHERE codigocontable='" + empresaactiva + "' "
        csql.sql = csql.sql + "ORDER BY codigo "
        csql.Execute
        
        If csql.RowsAffected > 0 Then
            Set resultados = csql.OpenResultset
            While Not resultados.EOF
                ComboLOCAL.AddItem (resultados(0) + " " + resultados(1))
                resultados.MoveNext
            Wend
            resultados.Close
            Set resultados = Nothing
        ComboLOCAL.text = ComboLOCAL.List(0)
        End If
        localfiltro = Mid(ComboLOCAL.List(0), 1, 2)
        
End Sub



Private Sub Grid1_Click()
Dim fila As Double
fila = Grid1.ActiveCell.row

If Grid1.ActiveCell.col = 14 Then

If Grid1.Cell(Grid1.ActiveCell.row, 13).text = "0" Then
Grid1.Cell(Grid1.ActiveCell.row, Grid1.ActiveCell.col).text = "0"
End If

End If
If Grid1.ActiveCell.col = 14 And Grid1.Cell(fila, 12).text = "1" Then
Call MsgBox("debe rebajar anticipos antes cancelar")
Grid1.Cell(fila, 14).text = "0"
End If


If Grid1.ActiveCell.col = 14 And Grid1.Cell(fila, 15).text = "" Then
Call MsgBox("debe colocar fecha de pago antes de autorizar")
Grid1.Cell(fila, 14).text = "0"
End If


If Grid1.ActiveCell.col = 14 Or Grid1.ActiveCell.col = 15 Or Grid1.ActiveCell.col = 16 Then
If Grid1.Cell(fila, 14).text = "0" Then
Grid1.Cell(fila, 15).text = ""
End If
'
'    Call modificaordenpagoentre(Grid1.Cell(fila, 2).text, localfiltro, Grid1.Cell(fila, 14).text, Format(Grid1.Cell(fila, 15).text, "yyyy-mm-dd"), Grid1.Cell(fila, 0).text, Grid1.Cell(fila, 5).text)
End If

If Grid1.ActiveCell.col = 15 And Grid1.Cell(fila, 15).text <> "" Then
    
 Call modificaordenpago(Grid1.Cell(fila, 2).text, localfiltro, Grid1.Cell(fila, 14).text, Format(Grid1.Cell(fila, 15).text, "yyyy-mm-dd"), Grid1.Cell(fila, 1).text, Grid1.Cell(fila, 5).text, Grid1.Cell(fila, 3).text)
End If


End Sub
Public Sub modificaordenpago(numero, loc, pago, fecha, tipo, fechadocu, rutprove)

        Dim resultados As rdoResultset
        Dim sql As New rdoQuery
        Dim multi As Double
        Dim TOTAL As Double
        
        Dim tabla As String
        Set sql.ActiveConnection = contadb
        rutprove = Replace(rutprove, "-", "")
        tabla = "update  " & clientesistema & "fae" & loc & ".sv_dte_cedidos_" & loc & "  set  autorizacionpago='" + pago + "',fechaautorizacionpago='" + Format(fecha, "yyyy-mm-dd") + "' "
        tabla = tabla & "WHERE numero= '" & Val(numero) & "' and cedente_rut LIKE '" & Val(Mid(rutprove, 1, 9)) & "%' and tipo='" & tipo & "' "
        sql.sql = tabla
        sql.Execute
       
        
    
    End Sub
Public Function leefactura(tipo, numero, rut) As Boolean
    campos(0, 0) = "tipo"
    campos(1, 0) = "numero"
    campos(2, 0) = "total"
    campos(3, 0) = "abono"
    campos(4, 0) = ""
    If tipo = "FA" Then tipo = "1"
    If tipo = "ND" Then tipo = "2"
    If tipo = "NC" Then tipo = "3"
    If tipo = "FAE" Then tipo = "4"
    If tipo = "NDE" Then tipo = "5"
    If tipo = "NCE" Then tipo = "6"
    
    condicion = "tipo='" + tipo + "' and numero='" + numero + "' and rut='" + rut + "' "
    campos(0, 2) = "facturasdecompras"
    op = 5
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    contabilizada = False
    pagada = False
    
    Call sqlconta.sqlconta(op, condicion)
    If sqlconta.status = 0 Then
    leefactura = True
    If sqlconta.response(3, 3) >= sqlconta.response(2, 3) Then
    pagada = True
    End If
    
    End If
    

End Function
Public Function leemontofactura(tipo, numero, rut) As Double

    campos(0, 0) = "tipo"
    campos(1, 0) = "numero"
    campos(2, 0) = "total"
    campos(3, 0) = "abono"
    campos(4, 0) = ""
    If tipo = "FA" Then tipo = "1"
    If tipo = "ND" Then tipo = "2"
    If tipo = "NC" Then tipo = "3"
    If tipo = "FAE" Then tipo = "4"
    If tipo = "NDE" Then tipo = "5"
    If tipo = "NCE" Then tipo = "6"
    
    condicion = "tipo='" + tipo + "' and numero='" + numero + "' and rut='" + rut + "' "
    campos(0, 2) = "facturasdecompras"
    op = 5
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    leemontofactura = 0
    pagada = False
    
    Call sqlconta.sqlconta(op, condicion)
    If sqlconta.status = 0 Then
    leemontofactura = sqlconta.response(2, 3)
    If sqlconta.response(3, 3) >= sqlconta.response(2, 3) Then
    pagada = True
    End If
    
    End If
    

End Function


Public Function LEERFOLIOCE(tipo) As String
    Dim resultados As rdoResultset
    Dim csql As New rdoQuery
        Set csql.ActiveConnection = contadb
            csql.sql = "select max(numero) from movimientoscontables where mes = '" & Format(Format(fechasistema, "mm"), "00") & "' AND año = '" & Format(fechasistema, "yyyy") & "' and tipo='" + tipo + "' "
            csql.Execute
    If csql.RowsAffected > 0 Then
    Set resultados = csql.OpenResultset
        If IsNull(resultados(0)) = False Then
        LEERFOLIOCE = Format(resultados(0) + 1, "0000000000")
        Else
        LEERFOLIOCE = Format(1, "0000000000")
        End If
        
    End If
    
End Function
Sub grabarcomprobante_lineas(tipo, numero, LINEA, fecha, codigocuenta, tipoctacte, rutctacte, centrocosto, glosacontable, tipodocumento, numerodocumento, fechadocumento, fechavencimiento, monto, DH, creadopor, MES, año, fechacreacion, horacreacion, rutproveedor)
    Dim condicion As String
    Dim campos(40, 3) As String
    Dim op As Integer
    Dim TIPOCON As String
    Dim tipo2 As String
    Dim j As Integer
    Dim lin As String
    Dim lar As Integer
    campos(0, 0) = "tipo"
    campos(1, 0) = "numero"
    campos(2, 0) = "linea"
    campos(3, 0) = "fecha"
    campos(4, 0) = "codigocuenta"
    campos(5, 0) = "tipoctacte"
    campos(6, 0) = "rutctacte"
    campos(7, 0) = "centrocosto"
    campos(8, 0) = "glosacontable"
    campos(9, 0) = "tipodocumento"
    campos(10, 0) = "numerodocumento"
    campos(11, 0) = "fechadocumento"
    campos(12, 0) = "fechavencimiento"
    campos(13, 0) = "monto"
    campos(14, 0) = "dh"
    campos(15, 0) = "creadopor"
    campos(16, 0) = "mes"
    campos(17, 0) = "año"
    campos(18, 0) = "fechacreacion"
    campos(19, 0) = "horacreacion"
    campos(20, 0) = "rutproveedor"
    campos(21, 0) = ""
    
    campos(0, 1) = tipo
    campos(1, 1) = numero
    campos(2, 1) = LINEA
    campos(3, 1) = Format(fecha, "yyyy-mm-dd")
    campos(4, 1) = codigocuenta
    campos(5, 1) = tipoctacte
    campos(6, 1) = rutctacte
    campos(7, 1) = centrocosto
    campos(8, 1) = glosacontable
    campos(9, 1) = tipodocumento
    campos(10, 1) = numerodocumento
    campos(11, 1) = Format(fechadocumento, "yyyy-mm-dd")
    campos(12, 1) = Format(fechavencimiento, "yyyy-mm-dd")
    campos(13, 1) = monto

    campos(14, 1) = DH
    campos(15, 1) = creadopor
    campos(16, 1) = MES
    campos(17, 1) = año
    
    campos(18, 1) = Format(fechacreacion, "yyyy-mm-dd")
    campos(19, 1) = horacreacion
    campos(20, 1) = rutproveedor

    campos(0, 2) = "movimientoscontables"
   

    op = 2
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    
    Call sqlconta.sqlconta(op, condicion)
   'Call ACTUALIZADOCUMENTO("+")
   
End Sub

Public Function LEERcompras(ORDEN) As Double
Dim resultados As rdoResultset
    Dim csql As New rdoQuery
    Dim TOTAL As Double
    Dim multi As Double
    Dim pasada As String
    
    
        Set csql.ActiveConnection = gestionrubro
        csql.sql = "SELECT tipo,total,rut,numero,ordendecompra "
        csql.sql = csql.sql + "FROM l_ordendecompra_detalle_facturas_" + localfiltro + " WHERE ordendecompra='" + ORDEN + "' "
        csql.sql = csql.sql + "ORDER BY ordendecompra "
        csql.Execute
        TOTAL = 0
        contabilizada = "0"
        montocarne = 0
        montoharina = 0
       
        If csql.RowsAffected > 0 Then
            
            Set resultados = csql.OpenResultset
            rutreal = resultados(2)
            While Not resultados.EOF
               
               If resultados(0) = "NCE" Or resultados(0) = "NC" Then multi = -1 Else multi = 1
               TOTAL = TOTAL + (resultados(1) * multi)
               
               montocarne = montocarne + LEERMONTOIMPUESTO(resultados(0), resultados(3), ORDEN, "11400005")
               montoharina = montoharina + LEERMONTOIMPUESTO(resultados(0), resultados(3), ORDEN, "11400012")
               
               If leefactura(resultados(0), resultados(3), resultados(2)) = False Then
               
               pasada = "1"
               Else
               contabilizada = "1"
               End If
               
                
                resultados.MoveNext
            Wend
            resultados.Close
            Set resultados = Nothing
        End If
        LEERcompras = TOTAL
        If pasada = "1" Then
        pasada = ""
        contabilizada = "0"
        End If
        
End Function
Sub GRABARCOMPROBANTE(ORDEN, montocheque, DIFERENCIA, fechacheque, TIPOPAGO, glosadiferencia, rutprove, fechado)
    Dim DH As String
    Dim numero As String
    Dim LINEA As Double
    Dim fecha As Date
    Dim rut As String
    Dim tipodocumento As String
    Dim numerodocumento As String
    Dim fechadocumento As String
    Dim fechavencimiento As String
    Dim MES As String
    Dim año As String
    Dim monto As Double
    Dim CUENTABANCO As String
    Dim montofactura2 As Double
    
    Dim tipo2 As String
    Dim TIPO3 As String
    Dim DOCUMENTOPAGO As String
    Dim resultados As rdoResultset
    Dim csql As New rdoQuery
    Dim rut2 As String
    
        Set csql.ActiveConnection = contadb
        
        If TIPOPAGO <> "BH" Then
            csql.sql = "SELECT tipo,total,rut,numero,fecha "
            csql.sql = csql.sql + "FROM  facturasdecompras  WHERE numero='" + ORDEN + "' and rut='" & rutprove & "' and fecha='" + fechado + "'  "
            csql.sql = csql.sql + "ORDER BY numero "
            csql.Execute
        Else
             csql.sql = "SELECT 'BH',liquido,rut,numero,fecha "
            csql.sql = csql.sql + "FROM  boletasdehonorarios  WHERE numero='" + ORDEN + "' and rut='" & rutprove & "' and fecha='" + fechado + "'  "
            csql.sql = csql.sql + "ORDER BY numero "
            csql.Execute
        End If
        
        If csql.RowsAffected > 0 Then
            Set resultados = csql.OpenResultset
            While Not resultados.EOF
            DH = "D"
            If resultados(0) = "NC" Or resultados(0) = "NCE" Then
            DH = "H"
            End If
            fecha = Format(fechasistema, "yyyy-mm-dd")
            rut = resultados(2)
            tipodocumento = resultados(0)
            TIPO3 = resultados(0)
            If tipodocumento = "1" Then tipo2 = "1": TIPO3 = "FC"
            If tipodocumento = "2" Then tipo2 = "2": TIPO3 = "ND"
            If tipodocumento = "3" Then tipo2 = "3": TIPO3 = "NC"
            If tipodocumento = "4" Then tipo2 = "4": TIPO3 = "FC"
            If tipodocumento = "5" Then tipo2 = "5": TIPO3 = "ND"
            If tipodocumento = "6" Then tipo2 = "6": TIPO3 = "NC"
            If tipodocumento = "9" Then tipo2 = "9": TIPO3 = "EX"
            If tipodocumento = "0" Then tipo2 = "0": TIPO3 = "EE"
            
            If tipodocumento = "BH" Then tipo2 = "1": TIPO3 = "BH"
            
            
            
            
            
            If glosadiferencia = "CHEQUE" Then
                glosadiferencia = ""
                DOCUMENTOPAGO = "PF"
            Else
                DOCUMENTOPAGO = "NG"
            End If
            
            tipodocumento = TIPO3
            numerodocumento = resultados(3)
            fechadocumento = resultados(4)
            fechavencimiento = fechadocumento
            MES = Format(fechasistema, "mm")
            año = Format(fechasistema, "yyyy")
            monto = resultados(1)
            
            If DH = "D" Then
            TOTALCheque = TOTALCheque + monto
            Else
            TOTALCheque = TOTALCheque - monto
            End If
            
             lineacontable = lineacontable + 1
            rut2 = rut
            Call grabarcomprobante_lineas(DOCUMENTOPAGO, numerocontable, lineacontable, fecha, CUENTAPROVEEDOR, " ", rut2, " ", "CANCELA DOCUMENTO", tipodocumento, numerodocumento, fechadocumento, fechavencimiento, monto, DH, USUARIOSISTEMA, MES, año, Format(Date, "yyyy-mm-dd"), Time, rutreal)
            
            montofactura2 = resultados(1)
            If tipodocumento <> "BH" Then
                Call abonofactura(tipo2, numerodocumento, rut2, montofactura2)
            Else
                Call abonoBH(tipo2, numerodocumento, rut2, montofactura2)
            End If
            resultados.MoveNext
            Wend
            resultados.Close
            Set resultados = Nothing
        End If
   Rem graba linea  diferencia
        
        If DIFERENCIA <> 0 Then
           
        lineacontable = lineacontable + 1
            
            monto = DIFERENCIA
            DH = "H"
            If DH = "D" Then
            TOTALCheque = TOTALCheque + monto
            Else
            TOTALCheque = TOTALCheque - monto
            End If
            
            Call grabarcomprobante_lineas(DOCUMENTOPAGO, numerocontable, lineacontable, fecha, cuentadiferencia, " ", rut, " ", glosadiferencia, "OC", ORDEN, fechadocumento, fechavencimiento, monto, DH, USUARIOSISTEMA, MES, año, Format(Date, "yyyy-mm-dd"), Time, rut)
        End If
        lineafinal = lineacontable
        If TIPOPAGO <> "BH" Then
            Call leerotros(DOCUMENTOPAGO, numerocontable, fecha, localfiltro, rut, ORDEN, lineacontable, MES, año, tipo2)
        Else
            Call leerotros(DOCUMENTOPAGO, numerocontable, fecha, localfiltro, rut, ORDEN, lineacontable, MES, año, TIPO3)
        End If
        lineacontable = lineafinal
End Sub
Sub grabarcheque(montocheque As Double)
Dim tipodocumento As String
Dim numerodocumento As String
Dim CUENTABANCO As String
Dim fechavencimiento As String
Dim monto As Double
Dim DH As String



    Rem graba cheque
        
        NCHEQUE = NCHEQUE + 1
        lineacontable = lineacontable + 1
        
        If tipocontable = "PF" Then
            tipodocumento = "CH"
            numerodocumento = Format(NCHEQUE, "0000000000")
            CUENTABANCO = dato1.text + dato2.text + dato3.text
            fechavencimiento = fechacheque
            monto = montocheque
            
            
            Else
            
            tipodocumento = "NG"
            numerodocumento = Format(numerocontable, "0000000000")
            CUENTABANCO = "11130001"
            fechavencimiento = fechacheque
            monto = montocheque
            
        End If
        
        DH = "H"
        Call grabarcomprobante_lineas(tipocontable, numerocontable, lineacontable, FECHACONTABLE, CUENTABANCO, " ", "", " ", NOMBREGIRADO, tipodocumento, numerodocumento, FECHACONTABLE, fechavencimiento, monto, DH, USUARIOSISTEMA, Format(fechasistema, "mm"), Format(fechasistema, "yyyy"), Format(Date, "yyyy-mm-dd"), Time, rutcontable)
        If tipocontable = "PF" Then
        fecha = Format(fechasistema, "yyyy-mm-dd")
        Call grabacheque(CUENTABANCO, numerodocumento, fecha, monto, fechavencimiento, "PF", numerocontable, NOMBREGIRADO, "0")
        End If
End Sub

Public Sub leerotros(tipo, numero, fecha, loc, rut, ORDEN, LINEA, MES, año, TIPODO)
        Dim tipocontable As String
        Dim numerocontable As String
        
        Dim resultados As rdoResultset
        Dim sql As New rdoQuery
        Dim multi As Double
        Dim TOTAL As Double
        
        Dim tabla As String
        Set sql.ActiveConnection = gestionrubro
        
        tabla = "SELECT cuenta,glosa,monto,dh,tipodo,numerodo "
        tabla = tabla & "FROM " + clientesistema + "conta" + empresaactiva + ".facturasdecompras_anexospagos "
        tabla = tabla & "WHERE tipo='" + Mid(TIPODO, 1, 1) + "' and numero= '" & ORDEN & "' ORDER BY linea asc "
        sql.sql = tabla
        sql.Execute
        
        If sql.RowsAffected > 0 Then
        
            Set resultados = sql.OpenResultset
            While Not resultados.EOF
                LINEA = LINEA + 1
                If leerNombreCuentaMayor(resultados(0), 1) = "" Then
                rut = ""
                End If
                
                If resultados(3) = "D" Then
                TOTALCheque = TOTALCheque + resultados(2)
                Else
                TOTALCheque = TOTALCheque - resultados(2)
                End If
                tipocontable = tipo
                numerocontable = numero
                If resultados(4) = "DM" Or resultados(4) = "D1" Then
                tipocontable = resultados(4)
                numerocontable = resultados(5)
                Call abonoGUIADEVOLUCION(tipocontable, numerocontable, tipo, numero, fecha, resultados(2))
                
                End If
                
                Call grabarcomprobante_lineas(tipo, numero, LINEA, fecha, resultados(0), " ", rut, " ", resultados(1), tipocontable, numerocontable, fecha, fecha, resultados(2), resultados(3), USUARIOSISTEMA, MES, año, Format(Date, "yyyy-mm-dd"), Time, rut)
                
                resultados.MoveNext
            Wend
        
        End If
    lineafinal = LINEA
    
    
    End Sub



Sub grabacheque(cuenta, numero, emision, monto, vencimiento, tipocomprobante, numerocomprobante, giradoa, ubicacion)
    campos(0, 0) = "cuenta"
    campos(1, 0) = "numero"
    campos(2, 0) = "emision"
    campos(3, 0) = "monto"
    campos(4, 0) = "vencimiento"
    campos(5, 0) = "tipocomprobante"
    campos(6, 0) = "numerocomprobante"
    campos(7, 0) = "giradoa"
    campos(8, 0) = "ubicacion"
    campos(9, 0) = ""
    
    campos(0, 1) = cuenta
    campos(1, 1) = numero
    campos(2, 1) = emision
    campos(3, 1) = monto
    campos(4, 1) = vencimiento
    campos(5, 1) = tipocomprobante
    campos(6, 1) = numerocomprobante
    campos(7, 1) = giradoa
    campos(8, 1) = "0"
    campos(0, 2) = "chequesdocumento"
       
    op = 2
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)
End Sub


Sub abonofactura(tipo, numero, rut, monto)
    Dim csql As rdoQuery
    Set csql = New rdoQuery
    Set csql.ActiveConnection = contadb
    csql.sql = "update facturasdecompras set abono = abono + " & monto & " "
    csql.sql = csql.sql & "where tipo='" + tipo + "' and rut='" + rut + "' and numero='" + numero + "'"
    csql.Execute
'    Call sincronizadatos(csql.sql, contadb, "")
    
    csql.Close
    Set csql = Nothing
End Sub

Sub abonoBH(tipo, numero, rut, monto)
    Dim csql As rdoQuery
    Set csql = New rdoQuery
    Set csql.ActiveConnection = contadb
    csql.sql = "update boletasdehonorarios set abono = abono + " & monto & " "
    csql.sql = csql.sql & "where tipo='" + tipo + "' and rut='" + rut + "' and numero='" + numero + "'"
    csql.Execute
'    Call sincronizadatos(csql.sql, contadb, "")
    
    csql.Close
    Set csql = Nothing
End Sub
Sub abonoGUIADEVOLUCION(tipo, numero, TIPOCO, NUMEROCO, fechaco, montoco)
    Dim csql As rdoQuery
    Set csql = New rdoQuery
    Set csql.ActiveConnection = contadb
    csql.sql = "update devoluciones_proveedores set tipoco='" & TIPOCO & "',numeroco='" & NUMEROCO & "',fechaco='" & Format(fechaco, "yyyy-mm-dd") & "',montoco='" & montoco & "' "
    csql.sql = csql.sql & "where tipo='" + tipo + "' and numero='" + numero + "'"
    csql.Execute
    Call sincronizadatos(csql.sql, contadb, "")
    
    csql.Close
    Set csql = Nothing
End Sub


    Private Sub dato1_GotFocus()
        Call cargatexto(dato1)
    End Sub
    
    Private Sub dato2_GotFocus()
        Call cargatexto(dato2)
    End Sub
    
    Private Sub dato3_GotFocus()
        Call cargatexto(dato3)
    End Sub
'****************************************************************************
'GOTFOCUS
'****************************************************************************

'****************************************************************************
'KEYDOWN
'****************************************************************************
    Private Sub dato1_KeyDown(KeyCode As Integer, Shift As Integer)
        If KeyCode = 38 Then Unload Me: GoTo no:
        If KeyCode = vbKeyF2 Then Call ayudamayor(dato1)
        Call flechas(dato1, dato2, KeyCode)
no:
    End Sub
    
    Private Sub dato2_KeyDown(KeyCode As Integer, Shift As Integer)
        Call flechas(dato1, dato3, KeyCode)
    End Sub
    
    Private Sub dato3_KeyDown(KeyCode As Integer, Shift As Integer)
        Call flechas(dato2, dato4, KeyCode)
    End Sub
'*********************************************
'KEYDOWN
'****************************************************************************

'****************************************************************************
'KEYPRESS
'****************************************************************************
    Private Sub dato1_KeyPress(KeyAscii As Integer)
        KeyAscii = esNumero(KeyAscii)
        If KeyAscii = 13 Then
            Call ceros(dato1)
           
          dato2.SetFocus
          
           
        End If
    End Sub
    
    Private Sub dato2_KeyPress(KeyAscii As Integer)
        KeyAscii = esNumero(KeyAscii)
        If KeyAscii = 13 Then
            Call ceros(dato2)
           dato3.SetFocus
        End If
    End Sub
    
    Private Sub dato3_KeyPress(KeyAscii As Integer)
        KeyAscii = esNumero(KeyAscii)
        If KeyAscii = 13 Then
            Call ceros(dato3)
            lblBanco.Caption = leerNombreCuentaMayor(dato1.text & dato2.text & dato3.text, 3)
            If lblBanco.Caption <> "" Then
                
            dato4.SetFocus
            End If
        
        End If
    End Sub

Sub ayudamayor(ByRef caja As TextBox)
    Dim campos As Variant
    Dim cfijo As Variant
    Dim largo As Variant
    campos = Array("codigo", "nombre")
    largo = Array("12s", "40s")
    cfijo = "año='" + Format(fechasistema, "yyyy") + "' AND banco='1'"
    cabezas = Array("codigo", "nombre")
    mensajeAyuda = "Ayuda Cuentas del Mayor"
    basebus = clientesistema + "conta" + empresaactiva
    Call cargaAyudaT(Servidor, basebus, Usuario, password, "cuentasdelmayor", pivote, campos, cfijo, largo, 2)
    If Val(pivote.text) = 0 Then dato1.SetFocus: GoTo no
    dato2.Enabled = True
    dato3.Enabled = True
    dato1.text = Mid(pivote.text, 1, 2)
    dato2.text = Mid(pivote.text, 3, 2)
    dato3.text = Mid(pivote.text, 5, 4)
    caja.Enabled = True
    caja.SetFocus
no:
End Sub



Private Sub Grid1_DblClick()
If Grid1.ActiveCell.col = 8 Then
localorden = localfiltro
pagosotros.tipo.text = Grid1.Cell(Grid1.ActiveCell.row, 1).text

pagosotros.montoorden.Caption = Format(CDbl(Grid1.Cell(Grid1.ActiveCell.row, 9).text), "###,###,###,###")
pagosotros.numero = Grid1.Cell(Grid1.ActiveCell.row, 2).text
pagosotros.LBLPROVEEDOR.Caption = Grid1.Cell(Grid1.ActiveCell.row, 3).text + " " + Grid1.Cell(Grid1.ActiveCell.row, 4).text
pagosotros.Show vbModal
End If
'If Grid1.ActiveCell.col = 10 Then
'localorden = localfiltro
'
'GLOSARECEPCION.Show vbModal
'
'End If
'
'

End Sub

Private Sub Grid1_LeaveCell(ByVal row As Long, ByVal col As Long, NewRow As Long, NewCol As Long, Cancel As Boolean)
If (col = 14 Or col = 15 Or col = 16) And row <> NewRow Then
Call modificaordenpago(Grid1.Cell(row, 2).text, localfiltro, Grid1.Cell(row, 14).text, Grid1.Cell(row, 15).text, Grid1.Cell(row, 1).text, Grid1.Cell(row, 5).text, Grid1.Cell(row, 3).text)

End If

End Sub

'Sub ayudamayor(ByRef caja As TextBox)
'    Dim campos As Variant
'    Dim cfijo As Variant
'    Dim largo As Variant
'    campos = Array("codigo", "nombre")
'    largo = Array("8s", "40s")
'    cfijo = "año='" + Format(fechasistema, "yyyy") + "'"
'    cabezas = Array("cuenta", "nombre")
'    mensajeAyuda = "Ayuda tipo de Cuentas mayor"
'
'    Call cargaAyudaT(servidor, basebus, usuario, password, "cuentasdelmayor", dato1, campos, cfijo, largo, 2)
'    caja.Enabled = True
'    caja.SetFocus
'End Sub
Private Sub ORDEN_GotFocus()
Call cargatexto(ORDEN)
End Sub

Private Sub ORDEN_KeyPress(KeyAscii As Integer)
KeyAscii = esNumero(KeyAscii)
If KeyAscii = 13 Then
Call ceros(ORDEN)
Call BUSCAR_Click


End If

End Sub

Private Sub BUSCAR_Click()
 Dim i As Integer
 
  For i = 1 To Grid1.Rows - 1
            If Mid(Grid1.Cell(i, 2).text, 1, 10) = ORDEN.text Then
                Grid1.Range(i, 1, i, Grid1.Cols - 1).Selected
                Grid1.Cell(i, 1).EnsureVisible
                Exit For
            End If
        Next i
End Sub

Private Function LEERMONTOIMPUESTO(tipo, numero, ORDEN, cuenta) As Double

    Dim resultados As rdoResultset
    Dim csql As New rdoQuery
 
        Set csql.ActiveConnection = gestionrubro

            csql.sql = "select monto from l_ordendecompra_impuestos_" + localfiltro + " where cuenta = '" & cuenta & "' and tipo='" + tipo + "' and numero='" + numero + "' and numeroorden='" + ORDEN + "' "
            
            csql.Execute
    LEERMONTOIMPUESTO = 0
    If csql.RowsAffected > 0 Then
    
    Set resultados = csql.OpenResultset
    LEERMONTOIMPUESTO = resultados(0)
    
    End If
    
End Function

Public Function leerlineas(ORDEN) As Double

    Dim resultados As rdoResultset
    Dim csql As New rdoQuery

        Set csql.ActiveConnection = gestionrubro
        csql.sql = "SELECT tipo,total,rut,numero,fecha "
        csql.sql = csql.sql + "FROM l_ordendecompra_detalle_facturas_" + localfiltro + " WHERE ordendecompra='" + ORDEN + "' "
        csql.sql = csql.sql + "ORDER BY ordendecompra "
        csql.Execute
        leerlineas = 0
        If csql.RowsAffected > 0 Then
        leerlineas = csql.RowsAffected + leerlineasotros(ORDEN) + 1
        End If
        
        
End Function

Public Function leerlineasotros(ORDEN) As Double


        Dim resultados As rdoResultset
        Dim sql As New rdoQuery
        Dim multi As Double
        Dim TOTAL As Double
        
        Dim tabla As String
        Set sql.ActiveConnection = gestionrubro
        
        tabla = "SELECT cuenta,glosa,monto,dh "
        tabla = tabla & "FROM l_ordendecompra_anexopagos_" + localfiltro + " "
        tabla = tabla & "WHERE numero= '" & ORDEN & "' ORDER BY linea asc "
        sql.sql = tabla
        sql.Execute
        leerlineasotros = 0
        If sql.RowsAffected > 0 Then
        leerlineasotros = sql.RowsAffected
        
                End If
    
    
    End Function

Private Sub leerfacturaventacontabilidad(numero, loc)

Dim resultados As rdoResultset
    Dim csql As New rdoQuery
    Dim rut As String
    Dim PASO As String
    Dim fecha1 As String
    Dim fecha2 As String
    Dim LINEA As Double
    Dim TOTAL As Double
    Dim fec As Double
    Dim fec1 As Double
    Dim fechasum As String
    Dim total2 As Double
    Dim tipodoc As String
    Dim MESCONTABLE As Double
    Dim AÑOCONTABLE As Double
    Dim emprecon As String
    
    emprecon = leerdatoslocal(loc, "codigocontable")
    
    
        Set csql.ActiveConnection = contadb
        csql.sql = "SELECT dc.tipo,dc.numero,dc.rut,dc.fecha,dc.neto,dc.iva,dc.total,dc.caja "
        csql.sql = csql.sql + "FROM " + clientesistema + "conta" + emprecon + ".facturasdeventas as dc "
        csql.sql = csql.sql + "where dc.tipo='1' and dc.numero='" + numero + "' "
        csql.sql = csql.sql + ""
        csql.Execute
        TOTAL = 0
        total2 = 0
       
        
        If csql.RowsAffected > 0 Then
        Set resultados = csql.OpenResultset
        fechasum = Format(fechasistema, "yyyy") + "/" + Format(fechasistema, "mm") + "/" + Format(fechasistema, "dd")
        
         While Not resultados.EOF
                     
             If leefactura(tipodoc, resultados(1), leerdatoslocal(loc, "rut")) = "0" Then
             Grid1.Cell(LINEA, 3).text = leerdatoslocal(loc, "rut")
             Grid1.Cell(LINEA, 4).text = leerdatoslocal(loc, "nombre")
             Grid1.Cell(LINEA, 5).text = resultados(3)
             Grid1.Cell(LINEA, 6).text = resultados(4)
             Grid1.Cell(LINEA, 7).text = resultados(5)
             Grid1.Cell(LINEA, 8).text = "0"
             Grid1.Cell(LINEA, 9).text = "0"
             Grid1.Cell(LINEA, 10).text = "0"
             Grid1.Cell(LINEA, 11).text = "0"
             Grid1.Cell(LINEA, 12).text = "0"
             Grid1.Cell(LINEA, 13).text = resultados(6)
             Grid1.Cell(LINEA, 15).text = "ME"
                         
             MESCONTABLE = CDbl(Format(fechasistema, "mm"))
             AÑOCONTABLE = CDbl(Format(fechasistema, "yyyy"))
             If Format(resultados(3), "yyyy-mm") < Format(fechasistema, "yyyy-mm") And Format(fechasistema, "dd") <= diacierrecompra Then
             MESCONTABLE = MESCONTABLE - 1
             If MESCONTABLE = 0 Then MESCONTABLE = 12: AÑOCONTABLE = AÑOCONTABLE - 1
             
             End If
             
             Grid1.Cell(LINEA, 17).text = Format(MESCONTABLE, "00")
             Grid1.Cell(LINEA, 18).text = AÑOCONTABLE
           
           End If
            
             
            
            resultados.MoveNext
       
            Wend
End If
      
End Sub

Public Function LEERmercaderia(rut, numero) As Boolean
Dim resultados As rdoResultset
    Dim csql As New rdoQuery
    
    
        Set csql.ActiveConnection = gestionrubro
        csql.sql = "SELECT tipo,total,rut,numero,ordendecompra "
        csql.sql = csql.sql + "FROM l_ordendecompra_detalle_facturas_00 WHERE rut='" + rut + "' and numero='" + numero + "' "
        
        
        csql.Execute
        
        LEERmercaderia = False
       
        If csql.RowsAffected > 0 Then
            LEERmercaderia = True
            resultados.Close
            Set resultados = Nothing
        End If
        
End Function

Private Sub botonmisaccesos_Click()
programafiltro = Me.Caption
misaccesos.Show
End Sub


Private Sub botonmisfavoritos_Click()
Call AgregaFavoritos(USUARIOSISTEMA, App.EXEName, Me.Name, Me.Caption)
End Sub
