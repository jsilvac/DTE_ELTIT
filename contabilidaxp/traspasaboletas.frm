VERSION 5.00
Object = "{4777436C-EB8C-4596-98A8-EBCF98683969}#1.0#0"; "FlexCell.ocx"
Object = "{085FBDF3-00C5-421B-B762-1D57299A2B89}#1.0#0"; "CLBUTN.OCX"
Object = "{ADBBAED6-B16F-40DC-80DF-B44910CBA76C}#1.0#0"; "Frame-Xp.ocx"
Begin VB.Form proceso04 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Traspaso de Boletas"
   ClientHeight    =   9870
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   14985
   DrawStyle       =   5  'Transparent
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   658
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   999
   Begin XPFrame.FrameXp FrameQuickMenu 
      Height          =   615
      Left            =   11640
      TabIndex        =   24
      Top             =   0
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
      Alignment       =   1
      Begin VB.CommandButton botonmisfavoritos 
         Caption         =   "Mis Favoritos"
         Height          =   255
         Left            =   1800
         TabIndex        =   26
         Top             =   280
         Width           =   1335
      End
      Begin VB.CommandButton botonmisaccesos 
         Caption         =   "Permisos Modulo"
         Height          =   255
         Left            =   120
         TabIndex        =   25
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
      TabIndex        =   1
      Top             =   6120
      Width           =   135
   End
   Begin XPFrame.FrameXp FrameXp1 
      Height          =   9825
      Left            =   45
      TabIndex        =   2
      Top             =   0
      Width           =   14865
      _ExtentX        =   26220
      _ExtentY        =   17330
      BackColor       =   16744576
      Caption         =   "TRASPASO DE BOLETAS"
      CaptionEstilo3D =   1
      BackColor       =   16744576
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
      Begin VB.CommandButton Command4 
         BackColor       =   &H0080FF80&
         Caption         =   "TRASPASA CONTABILIDAD"
         Height          =   330
         Left            =   6210
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   8190
         Width           =   2535
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H0080FF80&
         Caption         =   "IMPRIMIR"
         Height          =   330
         Left            =   3735
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
         BackColor       =   16744576
         Caption         =   "DATOS DE FILTRADO"
         CaptionEstilo3D =   1
         BackColor       =   16744576
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Begin VB.OptionButton Option3 
            BackColor       =   &H00FF8080&
            Caption         =   "Boletas Manuales"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   240
            Left            =   11340
            TabIndex        =   17
            Top             =   480
            Width           =   2235
         End
         Begin VB.OptionButton Option2 
            BackColor       =   &H00FF8080&
            Caption         =   "Sistema Ibm"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   240
            Left            =   11880
            TabIndex        =   16
            Top             =   1080
            Width           =   1995
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00FF8080&
            Caption         =   "Sistema Admin"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   240
            Left            =   11340
            TabIndex        =   15
            Top             =   270
            Value           =   -1  'True
            Width           =   1995
         End
         Begin VB.CommandButton Command2 
            Caption         =   "LISTAR"
            Height          =   285
            Left            =   13440
            TabIndex        =   7
            Top             =   720
            Width           =   1125
         End
         Begin XPFrame.FrameXp FrameXp6 
            Height          =   675
            Left            =   90
            TabIndex        =   9
            Top             =   270
            Width           =   3300
            _ExtentX        =   5821
            _ExtentY        =   1191
            BackColor       =   16744576
            Caption         =   "MES"
            CaptionEstilo3D =   1
            BackColor       =   16744576
            ForeColor       =   65535
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
            BackColor       =   16744576
            Caption         =   "AÑO"
            CaptionEstilo3D =   1
            BackColor       =   16744576
            ForeColor       =   65535
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
            Left            =   6720
            TabIndex        =   13
            Top             =   240
            Width           =   4560
            _ExtentX        =   8043
            _ExtentY        =   1191
            BackColor       =   16744576
            Caption         =   "LOCAL"
            CaptionEstilo3D =   1
            BackColor       =   16744576
            ForeColor       =   65535
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
         BackColor       =   16744576
         Caption         =   "LISTADO DE BOLETAS EMITIDAS"
         CaptionEstilo3D =   1
         BackColor       =   16744576
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Begin FlexCell.Grid Grid1 
            Height          =   6360
            Left            =   0
            TabIndex        =   4
            Top             =   240
            Width           =   14595
            _ExtentX        =   25744
            _ExtentY        =   11218
            BackColorFixed  =   16744576
            Cols            =   5
            DefaultFontSize =   8.25
            GridColor       =   16711680
            Rows            =   30
         End
      End
      Begin XPFrame.FrameXp fechas 
         Height          =   1170
         Left            =   7785
         TabIndex        =   18
         Top             =   8595
         Width           =   7005
         _ExtentX        =   12356
         _ExtentY        =   2064
         BackColor       =   14737632
         Caption         =   "Rangos de Fecha"
         CaptionEstilo3D =   1
         BackColor       =   14737632
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
         Begin CoolButtons.cool_Button command8 
            Height          =   375
            Left            =   4950
            TabIndex        =   19
            Top             =   675
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   661
            SkinId          =   "13"
            Caption         =   "Cambia Fecha"
         End
         Begin VB.Label Label4 
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
            Height          =   375
            Left            =   360
            TabIndex        =   23
            Top             =   360
            Width           =   1935
         End
         Begin VB.Label Label3 
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
            Height          =   375
            Left            =   2520
            TabIndex        =   22
            Top             =   360
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
            Left            =   360
            TabIndex        =   21
            Top             =   720
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
            Left            =   2520
            TabIndex        =   20
            Top             =   720
            Width           =   1935
         End
      End
   End
End
Attribute VB_Name = "proceso04"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private localfiltro As String


Private Sub Command1_Click()
imprimir
End Sub



Private Sub COMMAND2_Click()
localfiltro = Mid(ComboLOCAL.text, 1, 2)
año = COMBOAÑO.text
MES = COMBOMES.ListIndex + 1
Call Conectarventas(Servidor, clientesistema + "ventas" + localfiltro, Usuario, password)

Grid1.Rows = 1
If Option1.Value = True Then
leeradmin

End If
If Option2.Value = True Then
leeribm

End If

'If Option4.Value = True Then
'    Call leertbk
'End If

End Sub



Private Sub Command3_Click()
Dim k As Integer


End Sub


Private Sub Command4_Click()
Dim k As Double
 If Verifica_FORM29(COMBOAÑO.text & "-" & COMBOMES.ListIndex + 1 & "-01", empresaactiva) = False Then
    For k = 1 To Grid1.Rows - 1
        If Grid1.Cell(k, 12).text = "0" Then
            Call grababoleta(k)
        End If
    Next k
 Else
    MsgBox mensaje_nopermiso, vbCritical, "ATENCION"
 End If


If Option1.Value = True Then
leeradmin

End If
If Option2.Value = True Then
leeribm

End If


End Sub

Private Sub command8_Click()
Call retornofecha(desdefecha, hastafecha)
End Sub

Private Sub Form_Load()
CENTRAR Me


    
    Call Conectar_BD

    sc = 0
CARGAGRILLA
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
Option1.Value = True
desdefecha.Caption = "01-" + Format(fechasistema, "mm-yyyy")
hastafecha.Caption = fechasistema

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
titulo = "LISTADO DE BOLETAS EMITIDAS " + COMBOMES.text + " " + COMBOAÑO.text
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
       
    FORMATOGRILLA(1, 1) = "FECHA"
    FORMATOGRILLA(1, 2) = "CAJA"
    FORMATOGRILLA(1, 3) = "B.INICIAL"
    FORMATOGRILLA(1, 4) = "B.FINAL"
    FORMATOGRILLA(1, 5) = "MONTO"
    FORMATOGRILLA(1, 6) = "EXENTO"
    FORMATOGRILLA(1, 7) = "TOTAL"
    FORMATOGRILLA(1, 8) = "CRCC"
    FORMATOGRILLA(1, 9) = "NOMBRE"
    FORMATOGRILLA(1, 10) = "TIPO"
    FORMATOGRILLA(1, 11) = "TBK"
    FORMATOGRILLA(1, 12) = "CONTA"
    
    Rem LARGO DE LOS DATOS
    FORMATOGRILLA(2, 1) = "8"
    FORMATOGRILLA(2, 2) = "8"
    FORMATOGRILLA(2, 3) = "10"
    FORMATOGRILLA(2, 4) = "10"
    FORMATOGRILLA(2, 5) = "10"
    FORMATOGRILLA(2, 6) = "10"
    FORMATOGRILLA(2, 7) = "10"
    FORMATOGRILLA(2, 8) = "4"
    FORMATOGRILLA(2, 9) = "25"
    FORMATOGRILLA(2, 10) = "5"
    FORMATOGRILLA(2, 11) = "5"
    FORMATOGRILLA(2, 12) = "5"

    
    Rem TIPO DE DATOS C=CEROS,S=STRING,N=NUMERICO,D=DATE,
    FORMATOGRILLA(3, 1) = "D"
    FORMATOGRILLA(3, 2) = "S"
    FORMATOGRILLA(3, 3) = "N"
    FORMATOGRILLA(3, 4) = "N"
    FORMATOGRILLA(3, 5) = "N"
    FORMATOGRILLA(3, 6) = "N"
    FORMATOGRILLA(3, 7) = "N"
    FORMATOGRILLA(3, 8) = "S"
    FORMATOGRILLA(3, 9) = "S"
    FORMATOGRILLA(3, 10) = "S"
    FORMATOGRILLA(3, 11) = "S"
    FORMATOGRILLA(3, 12) = "S"
    
    Rem FORMATO GRILLA
    FORMATOGRILLA(4, 3) = "##,###,##0"
    FORMATOGRILLA(4, 4) = "##,###,##0"
    FORMATOGRILLA(4, 5) = "##,###,##0"
    FORMATOGRILLA(4, 6) = "##,###,##0"
    FORMATOGRILLA(4, 7) = "##,###,##0"
    FORMATOGRILLA(4, 10) = "##,###,##0"
    
    Rem LOCCKED
    For k = 1 To 11
    FORMATOGRILLA(5, k) = "TRUE"
    Next k
        
    
    Grid1.Cols = 13
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
    Grid1.Column(12).CellType = cellCheckBox

    
End Sub



Private Sub monto_Click()
End Sub

Private Sub leeradmin()

Dim resultados As rdoResultset
    Dim csql As New rdoQuery
    Dim rut As String
    Dim PASO As String
    Dim fecha1 As String
    Dim fecha2 As String
    Dim LINEA As Double
    Dim total As Double
    Dim fec As Double
    Dim fec1 As Double
    Dim fechasum As String
    Dim total2 As Double
    Dim CAJAS As Double
    
    LINEA = 0: fec = 0: fec1 = 0
    fecha1 = Format(desdefecha.Caption, "yyyy-mm-dd")
    fecha2 = Format(hastafecha.Caption, "yyyy-mm-dd")
        Set csql.ActiveConnection = ventaslocal
        CAJAS = CDbl(localfiltro) + 20
        If CDbl(localfiltro) = 25 Then CAJAS = CDbl(localfiltro) + 40
        If CDbl(localfiltro) = 41 Then CAJAS = CDbl(localfiltro) + 50
        
'        If localfiltro = "52" Then
'            csql.sql = "SELECT fecha,CONCAT('0',local,caja),min(foliosii),max(foliosii),sum(if(rut<>'0888888888',total,0)),sum(if(rut<>'0888888888',exento,0)),"
'            csql.sql = csql.sql + "(select dp.tipopago from sv_documento_pagos_" + localfiltro + " as dp where dp.tipo=dc.tipo and dp.numero=dc.numero and dc.fecha=dp.fecha and dc.caja=dp.caja) as tipopago,dc.contabilizado "
'            csql.sql = csql.sql + "FROM sv_documento_cabeza_" + localfiltro + " as dc "
'            csql.sql = csql.sql + "where (tipo='BV' or tipo='BE') and fecha between '" + fecha1 + "' and '" + fecha2 + "' "
'            csql.sql = csql.sql + "group by dc.fecha,dc.caja,tipopago order by fecha "
'
'            'ariel cambia por inner join consulta original arriba
'            csql.sql = "SELECT  dc.fecha,  CONCAT('0',dc.local,dc.caja),  MIN(dc.foliosii),  MAX(dc.foliosii),  SUM(IF(dc.rut<>'0888888888',total,0)),  SUM(IF(dc.rut<>'0888888888',exento,0)),  dp.tipopago AS tipopago,dc.contabilizado "
'            csql.sql = csql.sql + "FROM sv_documento_cabeza_" + localfiltro + " AS dc INNER JOIN  sv_documento_pagos_" + localfiltro + " AS dp ON (dp.tipo = dc.tipo AND dp.numero = dc.numero  AND dc.fecha = dp.fecha AND dc.caja = dp.caja) "
'            csql.sql = csql.sql + "where dc.tipo='BV' and dc.fecha between '" + fecha1 + "' and '" + fecha2 + "' "
'            csql.sql = csql.sql + "GROUP BY dc.fecha,dc.caja,tipopago ORDER BY dc.fecha "
'        Else
            csql.sql = "SELECT fecha,CONCAT('0',local,caja),min(foliosii),max(foliosii),sum(if(rut<>'0888888888',total,0)),sum(if(rut<>'0888888888',exento,0)),contabilizado,contabilizado "
            csql.sql = csql.sql & " FROM sv_documento_cabeza_" + localfiltro + " "
            csql.sql = csql.sql & " where (tipo='BV' or tipo='BE') "
            csql.sql = csql.sql & " and fecha between '" & fecha1 & "' and '" & fecha2 & "' "
            csql.sql = csql.sql & " group by fecha,caja,contabilizado"
            csql.sql = csql.sql & " order by fecha "
        
'        End If
        
        csql.Execute
        total = 0
        total2 = 0
        Grid1.Rows = csql.RowsAffected + 1
        Grid1.AutoRedraw = False
        If csql.RowsAffected > 0 Then
        Set resultados = csql.OpenResultset
        fechasum = Format(fechasistema, "yyyy") + "/" + Format(fechasistema, "mm") + "/" + Format(fechasistema, "dd")
        
         While Not resultados.EOF
                     
             LINEA = LINEA + 1
            
            Grid1.Cell(LINEA, 1).text = resultados(0)
             Grid1.Cell(LINEA, 2).text = resultados(1)
             Grid1.Cell(LINEA, 3).text = resultados(2)
             Grid1.Cell(LINEA, 4).text = resultados(3)
             
             If localfiltro = "50" Or localfiltro = "51" Then
             Grid1.Cell(LINEA, 5).text = resultados(4) - resultados(5)
             Grid1.Cell(LINEA, 6).text = resultados(5)
             Grid1.Cell(LINEA, 7).text = resultados(4)
             
             Else
             Grid1.Cell(LINEA, 5).text = resultados(4)
             Grid1.Cell(LINEA, 6).text = "0"
             Grid1.Cell(LINEA, 7).text = resultados(4)
             
             End If
             If localfiltro = "52" Then
             If resultados(6) = "4" Or resultados(6) = "3" Then
            Grid1.Cell(LINEA, 2).text = "05202"
            
            End If
            End If
             Grid1.Cell(LINEA, 8).text = leerdatoslocal(localfiltro, "codigocrcc")
             Grid1.Cell(LINEA, 9).text = leerdatos(contadb, "centrosdecosto", "nombre", "codigo='" + leerdatoslocal(localfiltro, "codigocrcc") + "'") + " (SISTEMA ADMIN)"
             
             Grid1.Cell(LINEA, 11).text = ""
             If localfiltro = "52" Then
                If resultados(6) = "4" Or resultados(6) = "3" Then
                Grid1.Cell(LINEA, 11).text = "TBK"
                End If
             End If
             
            Grid1.Cell(LINEA, 10).text = resultados("contabilizado")
            Grid1.Cell(LINEA, 12).text = leeboleta(LINEA)
            resultados.MoveNext
       
            Wend
End If
      Grid1.AutoRedraw = False
      Grid1.Refresh
      
End Sub

Private Sub leertbk()

Dim resultados As rdoResultset
    Dim csql As New rdoQuery
    Dim rut As String
    Dim PASO As String
    Dim fecha1 As String
    Dim fecha2 As String
    Dim LINEA As Double
    Dim total As Double
    Dim fec As Double
    Dim fec1 As Double
    Dim fechasum As String
    Dim total2 As Double
    Dim CAJAS As Double
    
    LINEA = 0: fec = 0: fec1 = 0
    fecha1 = Format(desdefecha.Caption, "yyyy-mm-dd")
    fecha2 = Format(hastafecha.Caption, "yyyy-mm-dd")
        Set csql.ActiveConnection = ventaslocal
        CAJAS = CDbl(localfiltro) + 20
        If CDbl(localfiltro) = 25 Then CAJAS = CDbl(localfiltro) + 40
        If CDbl(localfiltro) = 41 Then CAJAS = CDbl(localfiltro) + 50
        
        
        csql.sql = "SELECT fecha,CONCAT('0',local,caja),min(foliosii),max(foliosii),sum(if(rut<>'0888888888',total,0)),sum(if(rut<>'0888888888',exento,0)) "
        csql.sql = csql.sql + "FROM sv_documento_cabeza_" + localfiltro + " "
        csql.sql = csql.sql + "where tipo='BV' and fecha between '" + fecha1 + "' and '" + fecha2 + "' "
        csql.sql = csql.sql + "group by fecha,caja order by fecha "
        
        csql.Execute
        total = 0
        total2 = 0
        Grid1.Rows = csql.RowsAffected + 1
        
        If csql.RowsAffected > 0 Then
        Set resultados = csql.OpenResultset
        fechasum = Format(fechasistema, "yyyy") + "/" + Format(fechasistema, "mm") + "/" + Format(fechasistema, "dd")
        
         While Not resultados.EOF
                     
             LINEA = LINEA + 1
            
            Grid1.Cell(LINEA, 1).text = resultados(0)
             Grid1.Cell(LINEA, 2).text = resultados(1)
             Grid1.Cell(LINEA, 3).text = resultados(2)
             Grid1.Cell(LINEA, 4).text = resultados(3)
             
             If localfiltro = "50" Or localfiltro = "51" Then
             Grid1.Cell(LINEA, 5).text = resultados(4) - resultados(5)
             Grid1.Cell(LINEA, 6).text = resultados(5)
             Grid1.Cell(LINEA, 7).text = resultados(4)
             
             Else
             Grid1.Cell(LINEA, 5).text = resultados(4)
             Grid1.Cell(LINEA, 6).text = "0"
             Grid1.Cell(LINEA, 7).text = resultados(4)
             
             End If
             
             Grid1.Cell(LINEA, 8).text = leerdatoslocal(localfiltro, "codigocrcc")
             Grid1.Cell(LINEA, 9).text = leerdatos(contadb, "centrosdecosto", "nombre", "codigo='" + leerdatoslocal(localfiltro, "codigocrcc") + "'") + " (SISTEMA ADMIN)"
             Grid1.Cell(LINEA, 10).text = leeboleta(LINEA)

            resultados.MoveNext
       
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



Sub grababoleta(LINEA)
    Dim netos As Double
    Dim DH As String
    Dim exentos As Double
    Dim TIPOCON As String
    Dim CRCC As String
    
    campos(0, 0) = "fecha"
    campos(1, 0) = "caja"
    campos(2, 0) = "boletainicial"
    campos(3, 0) = "boletafinal"
    campos(4, 0) = "monto"
    campos(5, 0) = "exento"
    campos(6, 0) = "total"
    campos(7, 0) = "centrocosto"
    campos(8, 0) = "estbk"
    campos(9, 0) = "tipodocumento"
    campos(10, 0) = ""
    
    campos(0, 1) = Format(Grid1.Cell(LINEA, 1).text, "yyyy-mm-dd")
    campos(1, 1) = Grid1.Cell(LINEA, 2).text
    campos(2, 1) = Replace(Grid1.Cell(LINEA, 3).text, ",", ".")
    campos(3, 1) = Replace(Grid1.Cell(LINEA, 4).text, ",", ".")
    campos(4, 1) = Replace(Grid1.Cell(LINEA, 5).text, ",", ".")
    campos(5, 1) = Replace(Grid1.Cell(LINEA, 6).text, ",", ".")
    campos(6, 1) = Replace(Grid1.Cell(LINEA, 7).text, ",", ".")
    campos(7, 1) = Grid1.Cell(LINEA, 8).text
    If Grid1.Cell(LINEA, 11).text = "TBK" Then
        campos(8, 1) = "1"
    Else
        campos(8, 1) = "0"
    End If
   
    campos(9, 1) = Grid1.Cell(LINEA, 10).text
    
    condicion = ""
    Rem condicion = "CAJA='" & Grid1.Cell(LINEA, 2).text & "' AND FECHA='" & Format(Grid1.Cell(LINEA, 1).text, "yyyy-mm-dd") & "'"
    
    campos(0, 2) = "boletasdeventa"
    op = 2
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    
    
    Call sqlconta.sqlconta(op, condicion)
    

End Sub


Public Function leeboleta(LINEA) As String

    
    campos(0, 0) = "fecha"
    campos(1, 0) = ""
    condicion = "fecha='" & Format(Grid1.Cell(LINEA, 1).text, "yyyy-mm-dd") & "'"
    condicion = condicion & " and caja='" & Grid1.Cell(LINEA, 2).text & "'"
    condicion = condicion & " and tipodocumento='" & Grid1.Cell(LINEA, 10).text & "' "
    campos(0, 2) = "boletasdeventa"
    op = 5
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    
    Call sqlconta.sqlconta(op, condicion)
    If sqlconta.status = 0 Then
    leeboleta = "1"
    
    Else
    leeboleta = "0"
    
    End If
    
    

End Function

Sub leectacte(rut)
    campos(0, 0) = "rut"
    campos(1, 0) = "nombre"
    campos(2, 0) = ""
    campos(0, 2) = "cuentascorrientes"
    condicion = "tipo=" + "'" + cuentacliente + "' and rut=" + "'" + rut + "' and año='" + año + "'"
    op = 5
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)
    If sqlconta.status = 4 Then
    Call crearcuentacorriente(rut)
    End If
    
End Sub
Sub crearcuentacorriente(rut)
    Dim resultados As rdoResultset
    Dim csql As New rdoQuery
    
        Set csql.ActiveConnection = gestion

            csql.sql = "INSERT INTO " + clientesistema + "conta" + empresaactiva + ".cuentascorrientes "
            csql.sql = csql.sql & "(año,tipo,rut,nombre,direccion,comuna,ciudad,giro,fono) "
            csql.sql = csql.sql & "SELECT '" + año + "','" + cuentacliente + "',mc.rut,mc.nombre,mc.direccion,mc.comuna,mc.ciudad,mc.giro,mc.fono1 "
            csql.sql = csql.sql & "FROM " & clientesistema & "ventas.sv_maestroclientes as mc "
            csql.sql = csql.sql & "WHERE mc.rut = '" & rut & "' AND mc.sucursal ='0'"
            
            csql.Execute
            Call sincronizadatos(csql.sql, gestion, "")
            
            
            
            csql.sql = "INSERT INTO " + clientesistema + "conta" + empresaactiva + ".saldosctacte "
            csql.sql = csql.sql & "(año,tipo,rut) "
            csql.sql = csql.sql & "SELECT '" + año + "','" + cuentacliente + "',mc.rut "
            csql.sql = csql.sql & "FROM " & clientesistema & "ventas.sv_maestroclientes as mc "
            csql.sql = csql.sql & "WHERE mc.rut = '" & rut & "' AND mc.sucursal ='0'"
            
            csql.Execute
            Call sincronizadatos(csql.sql, gestion, "")
            


End Sub

'cSql.SQL = "INSERT INTO l_movimientos_detalle_" & empresaactiva & " "
'            cSql.SQL = cSql.SQL & "(tipo, numero, linea, fecha, rut, codigo, descripcion, cantidad, unidades, precio, total, costoventa, bodega, bodegatraspaso, uxc) "
'            cSql.SQL = cSql.SQL & "SELECT dd.tipo, dd.numero, dd.linea, dd.fecha, dd.rut, dd.codigo, dd.descripcion, dd.cantidad, dd.unidades, dd.precio, dd.total, dd.pcosto, dd.bodega, dd.bodega, ROUND(dd.unidades / dd.cantidad, 0) "
'            cSql.SQL = cSql.SQL & "FROM " & baseVentas & rubro & ".sv_documento_detalle_" + empresaactiva + " as dd "
'            cSql.SQL = cSql.SQL & "WHERE dd.local = '" & empresaactiva & "' AND dd.tipo = '" & v.detalle.tipo & "' AND dd.numero = '" & v.detalle.numero & "'"
'            cSql.Execute
Private Sub Grid1_AfterReorderColumn(ByVal OriginalPosition As Long, ByVal NewPosition As Long)

End Sub
Private Sub leeribm()

    Dim resultados As rdoResultset
    Dim csql As New rdoQuery
    Dim rut As String
    Dim PASO As String
    Dim fecha1 As String
    Dim fecha2 As String
    Dim LINEA As Double
    Dim total As Double
    Dim fec As Double
    Dim fec1 As Double
    Dim fechasum As String
    Dim total2 As Double
    Dim ARCHIVO As String
    Dim caja As String
    Dim FOLIOINICIAL As String
    Dim FOLIOFINAL As String
    Dim FECHACAJA As String
    Dim monto As String
    
    LINEA = 0: fec = 0: fec1 = 0
    fecha1 = Format(desdefecha.Caption, "yyyy-mm-dd")
    fecha2 = Format(hastafecha.Caption, "yyyy-mm-dd")
    For k = 1 To 31
         fechasum = Format(k, "00") + Format(MES, "00") + Mid(año, 3, 2)
        ARCHIVO = "z:\informes\zetas\su" + fechasum + ".TXT"
         If ExisteArchivo(ARCHIVO) = True Then
         Close 20
         Open ARCHIVO For Input As #20
         
         While EOF(20) = False
                     
             Line Input #20, PASO
             If Val(Mid(PASO, 1, 10)) <> 0 Then
             caja = Format(Mid(PASO, 9, 2), "00")
             FOLIOFINAL = Mid(PASO, 48, 10)
             FOLIOINICIAL = Mid(PASO, 35, 10)
             FECHACAJA = Mid(PASO, 13, 8)
             FECHACAJA = Mid(FECHACAJA, 5, 4) + "-" + Mid(FECHACAJA, 3, 2) + "-" + Format(Mid(FECHACAJA, 1, 2), "00")
             monto = Mid(PASO, 70, 10)
             
             LINEA = LINEA + 1

             Grid1.Rows = Grid1.Rows + 1
             Grid1.Cell(LINEA, 1).text = FECHACAJA
             Grid1.Cell(LINEA, 2).text = caja
             Grid1.Cell(LINEA, 3).text = FOLIOINICIAL
             Grid1.Cell(LINEA, 4).text = FOLIOFINAL
             Grid1.Cell(LINEA, 5).text = monto
             Grid1.Cell(LINEA, 6).text = "0"
             Grid1.Cell(LINEA, 7).text = monto
             Grid1.Cell(LINEA, 8).text = leerdatoslocal(localfiltro, "codigocrcc")
             Grid1.Cell(LINEA, 9).text = leerdatos(contadb, "centrosdecosto", "nombre", "codigo='" + leerdatoslocal(localfiltro, "codigocrcc") + "'") + " (SISTEMA IBM)"
             Grid1.Cell(LINEA, 10).text = leeboleta(LINEA)

            End If
            
            Wend
End If
Next k


      
End Sub

Private Sub botonmisaccesos_Click()
programafiltro = Me.Caption
misaccesos.Show
End Sub


Private Sub botonmisfavoritos_Click()
Call AgregaFavoritos(USUARIOSISTEMA, App.EXEName, Me.Name, Me.Caption)
End Sub
