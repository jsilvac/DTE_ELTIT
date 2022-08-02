VERSION 5.00
Object = "{4777436C-EB8C-4596-98A8-EBCF98683969}#1.0#0"; "FlexCell.ocx"
Object = "{085FBDF3-00C5-421B-B762-1D57299A2B89}#1.0#0"; "CLBUTN.OCX"
Object = "{ADBBAED6-B16F-40DC-80DF-B44910CBA76C}#1.0#0"; "Frame-Xp.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form proceso21 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Traspaso de Facturas"
   ClientHeight    =   10455
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   15000
   DrawStyle       =   5  'Transparent
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   697
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1000
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox pivote 
      Height          =   285
      Left            =   6600
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   9480
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
      Height          =   10350
      Left            =   90
      TabIndex        =   2
      Top             =   45
      Width           =   14865
      _ExtentX        =   26220
      _ExtentY        =   18256
      BackColor       =   12632256
      Caption         =   "CENTRALIZACION DE CAJAS"
      CaptionEstilo3D =   1
      BackColor       =   12632256
      ColorBarraArriba=   4210752
      ColorBarraAbajo =   4210752
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
         Left            =   4920
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   8520
         Width           =   2535
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H0080FF80&
         Caption         =   "IMPRIMIR"
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
         Left            =   2520
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   8520
         Width           =   2130
      End
      Begin XPFrame.FrameXp FrameXp3 
         Height          =   1050
         Left            =   240
         TabIndex        =   5
         Top             =   360
         Width           =   14640
         _ExtentX        =   25823
         _ExtentY        =   1852
         BackColor       =   12632256
         Caption         =   "DATOS DE FILTRADO"
         CaptionEstilo3D =   1
         BackColor       =   12632256
         ColorBarraArriba=   4210752
         ColorBarraAbajo =   4210752
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
            Caption         =   "LISTAR"
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
            Left            =   12000
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
            BackColor       =   8421504
            Caption         =   "MES"
            CaptionEstilo3D =   1
            BackColor       =   8421504
            ForeColor       =   65535
            ColorBarraArriba=   12632256
            ColorBarraAbajo =   4210752
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
            BackColor       =   8421504
            Caption         =   "AÑO"
            CaptionEstilo3D =   1
            BackColor       =   8421504
            ForeColor       =   65535
            ColorBarraArriba=   12632256
            ColorBarraAbajo =   4210752
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
            Width           =   4560
            _ExtentX        =   8043
            _ExtentY        =   1191
            BackColor       =   8421504
            Caption         =   "LOCAL"
            CaptionEstilo3D =   1
            BackColor       =   8421504
            ForeColor       =   65535
            ColorBarraArriba=   12632256
            ColorBarraAbajo =   4210752
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
               Left            =   45
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
         BackColor       =   14737632
         Caption         =   "LISTADO DE PAGOS RECIBIDOS"
         CaptionEstilo3D =   1
         BackColor       =   14737632
         ColorBarraArriba=   8421504
         ColorBarraAbajo =   4210752
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
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
            BackColorFixed  =   4210752
            Cols            =   5
            DefaultFontSize =   8.25
            ForeColorFixed  =   16777215
            GridColor       =   16711680
            Rows            =   30
         End
      End
      Begin XPFrame.FrameXp fechas 
         Height          =   1170
         Left            =   7560
         TabIndex        =   15
         Top             =   8520
         Width           =   7005
         _ExtentX        =   12356
         _ExtentY        =   2064
         BackColor       =   14737632
         Caption         =   "Rangos de Fecha"
         CaptionEstilo3D =   1
         BackColor       =   14737632
         ColorBarraArriba=   8421504
         ColorBarraAbajo =   4210752
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
            TabIndex        =   16
            Top             =   675
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   661
            SkinId          =   "13"
            Caption         =   "Cambia Fecha"
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
            TabIndex        =   19
            Top             =   720
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
            TabIndex        =   18
            Top             =   360
            Width           =   1935
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
            TabIndex        =   17
            Top             =   360
            Width           =   1935
         End
      End
      Begin XPFrame.FrameXp FrameXp5 
         Height          =   930
         Left            =   120
         TabIndex        =   22
         Top             =   9000
         Visible         =   0   'False
         Width           =   5085
         _ExtentX        =   8969
         _ExtentY        =   1640
         BackColor       =   14737632
         Caption         =   "TIPOS"
         CaptionEstilo3D =   1
         BackColor       =   14737632
         ColorBarraArriba=   8421504
         ColorBarraAbajo =   4210752
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
         Begin VB.OptionButton optElectronicas 
            Caption         =   "ELETRONICAS"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   3120
            TabIndex        =   25
            Top             =   480
            Width           =   1815
         End
         Begin VB.OptionButton optNormales 
            Caption         =   "NORMALES"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1440
            TabIndex        =   24
            Top             =   480
            Width           =   1575
         End
         Begin VB.OptionButton optTodas 
            Caption         =   "TODAS"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   23
            Top             =   480
            Value           =   -1  'True
            Width           =   1095
         End
      End
      Begin MSComctlLib.ProgressBar barra 
         Height          =   255
         Left            =   120
         TabIndex        =   26
         Top             =   8160
         Width           =   14655
         _ExtentX        =   25850
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Factura Electronica"
         Height          =   255
         Left            =   240
         TabIndex        =   21
         Top             =   8520
         Visible         =   0   'False
         Width           =   1455
      End
   End
End
Attribute VB_Name = "proceso21"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private localfiltro As String
Private cr_crcc As String
Private cr_cuenta As String
Private linea2 As Double



Private Sub CmdFavoritos_Click()
    Call AgregaFavoritos(USUARIOSISTEMA, App.EXEName, Me.Name, Me.Caption)
End Sub

Private Sub Command1_Click()
imprimir
End Sub



Private Sub COMMAND2_Click()
localfiltro = Mid(ComboLOCAL.text, 1, 2)
año = COMBOAÑO.text
MES = COMBOMES.ListIndex + 1
Call Conectarventas(Servidor, Replace(clientesistema, "_", "ip_") + "ventas" + localfiltro, Usuario, password)
leer


End Sub



Private Sub Command3_Click()
Dim k As Integer


End Sub


Private Sub Command4_Click()
    
    
    Dim k As Double
        For k = 1 To Grid1.Rows - 1
            If Grid1.Cell(k, 14).text <> "1" Then
                Call grabafactura(Format(Grid1.Cell(k, 2).text, "0000000000"), Format(Grid1.Cell(k, 5).text, "yyyy-mm-dd"), Grid1.Cell(k, 3).text, Grid1.Cell(k, 13).text, Grid1.Cell(k, 1).text)
            End If
        Next k
    leer
End Sub

Private Sub command8_Click()
Call retornofecha(desdefecha, hastafecha)
End Sub

Private Sub Form_Load()
CENTRAR Me


    
    Call Conectar_BD

    sc = 0
CARGAGRILLA

Call Conectargestion(Servidor, Replace(clientesistema, "_", "ip_") + "gestion", Usuario, password)
Call Conectargestionrubro(Servidor, Replace(clientesistema, "_", "ip_") + "gestion00", Usuario, password)

For k = 1 To 12
COMBOMES.AddItem MonthName(k)
Next k
COMBOMES.ListIndex = CDbl(Format(fechasistema, "mm") - 1)
For k = 2000 To Val(Format(fechasistema, "yyyy"))
COMBOAÑO.AddItem k
Next k
COMBOAÑO.ListIndex = k - 2001
LEErlocales
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
titulo = "LISTADO DE FACTURAS EMITIDAS " + COMBOMES.text + " " + COMBOAÑO.text
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
    FORMATOGRILLA(1, 4) = "CLIENTE"
    FORMATOGRILLA(1, 5) = "FECHA"
    FORMATOGRILLA(1, 6) = "NETO"
    FORMATOGRILLA(1, 7) = "IVA"
    FORMATOGRILLA(1, 8) = "I.REFRE"
    FORMATOGRILLA(1, 9) = "I.VINO "
    FORMATOGRILLA(1, 10) = "I.LICOR"
    FORMATOGRILLA(1, 11) = "I.HARINA"
    FORMATOGRILLA(1, 12) = "I.CARNE"
    FORMATOGRILLA(1, 13) = "TOTAL  "
    FORMATOGRILLA(1, 14) = "CONTA"
    FORMATOGRILLA(1, 15) = "TIPO"
    FORMATOGRILLA(1, 16) = "CUENTA"
    FORMATOGRILLA(1, 17) = "CRCC"
    
    Rem LARGO DE LOS DATOS
    FORMATOGRILLA(2, 1) = "3"
    FORMATOGRILLA(2, 2) = "8"
    FORMATOGRILLA(2, 3) = "10"
    FORMATOGRILLA(2, 4) = "23"
    FORMATOGRILLA(2, 5) = "8"
    FORMATOGRILLA(2, 6) = "0"
    FORMATOGRILLA(2, 7) = "0"
    FORMATOGRILLA(2, 8) = "0"
    FORMATOGRILLA(2, 9) = "0"
    FORMATOGRILLA(2, 10) = "0"
    FORMATOGRILLA(2, 11) = "0"
    FORMATOGRILLA(2, 12) = "0"
    FORMATOGRILLA(2, 13) = "8"
    FORMATOGRILLA(2, 14) = "5"
    FORMATOGRILLA(2, 15) = "0"
    FORMATOGRILLA(2, 16) = "0"
    FORMATOGRILLA(2, 17) = "0"

    
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
    FORMATOGRILLA(3, 11) = "N"
    FORMATOGRILLA(3, 12) = "N"
    FORMATOGRILLA(3, 13) = "N"
   
    Rem FORMATO GRILLA
    FORMATOGRILLA(4, 6) = "##,###,##0"
    FORMATOGRILLA(4, 7) = "##,###,##0"
    FORMATOGRILLA(4, 8) = "##,###,##0"
    FORMATOGRILLA(4, 9) = "##,###,##0"
    FORMATOGRILLA(4, 10) = "##,###,##0"
    FORMATOGRILLA(4, 11) = "##,###,##0"
    FORMATOGRILLA(4, 12) = "##,###,##0"
    FORMATOGRILLA(4, 13) = "##,###,##0"
    Rem LOCCKED
    For k = 1 To 15
    FORMATOGRILLA(5, k) = "TRUE"
    
    Next k
        
    
    Grid1.Cols = 18
    Grid1.Rows = 2
    
 Grid1.AllowUserResizing = False
    Grid1.DisplayFocusRect = False
Grid1.ExtendLastCol = False
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
    Grid1.Column(14).CellType = cellCheckBox
    
    
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
    Dim fec As Double
    Dim fec1 As Double
    Dim fechasum As String
    Dim total2 As Double
    Dim tipodoc As String
    
    LINEA = 0: fec = 0: fec1 = 0
    fecha1 = Format(desdefecha.Caption, "yyyy-mm-dd")
    fecha2 = Format(hastafecha.Caption, "yyyy-mm-dd")
        Set csql.ActiveConnection = ventaslocal
    csql.sql = "update sv_pagos_detalle_" + localfiltro + " as pd inner join sv_documento_cabeza_" + localfiltro + " as dc "
    csql.sql = csql.sql + "on pd.documento=dc.folio_ref  and pd.tipo=dc.tipo_ref "
    csql.sql = csql.sql + "set pd.tipofactura=dc.tipo,pd.numerofactura=dc.numero "
    csql.sql = csql.sql + "where pd.fecha like '" + Mid(fecha1, 1, 7) + "%'  and numerofactura='' ;"
        csql.Execute
    
    
        Set csql.ActiveConnection = ventaslocal
        csql.sql = "SELECT 'PC',dc.numero,dc.rut,mc.nombre,dc.fecha,dc.monto "
        csql.sql = csql.sql + "FROM sv_pagos_cabeza_" + localfiltro + " as dc," + Replace(clientesistema, "_", "ip_") + "ventas.sv_maestroclientes as mc "
        csql.sql = csql.sql + "where dc.rut=mc.rut and mc.sucursal='0' and "
        csql.sql = csql.sql & "dc.fecha between '" + fecha1 + "' and '" + fecha2 + "' "
        csql.sql = csql.sql + "Union "
        csql.sql = csql.sql + "select 'ZE',concat(mid(fecha,1,4),mid(fecha,6,2),mid(fecha,9,2)),'0000000019','BOLETAS',fecha,sum(monto) from boletasdeventa "
        csql.sql = csql.sql + "where fecha between '" & fecha1 & "' and '" & fecha2 & "' group by fecha "
        csql.sql = csql.sql + " order by numero"
        
        csql.Execute
        total = 0
        total2 = 0
        
        Grid1.Rows = 1
        
        If csql.RowsAffected > 0 Then
            Grid1.AutoRedraw = False
            Grid1.Rows = csql.RowsAffected + 1
            barra.Max = csql.RowsAffected + 1
            barra.Value = 0
            Set resultados = csql.OpenResultset
            fechasum = Format(fechasistema, "yyyy") + "/" + Format(fechasistema, "mm") + "/" + Format(fechasistema, "dd")
        
         While Not resultados.EOF
                     
             LINEA = LINEA + 1
             tipodoc = resultados(0)
             barra.Value = barra.Value + 1
            
             
             Grid1.Cell(LINEA, 1).text = resultados(0)
             Grid1.Cell(LINEA, 2).text = resultados(1)
             Grid1.Cell(LINEA, 3).text = Mid(resultados(2), 1, 9) + "-" + Mid(resultados(2), 10, 1)
             Grid1.Cell(LINEA, 4).text = resultados(3)
             Grid1.Cell(LINEA, 5).text = resultados(4)
'             Grid1.Cell(linea, 6).text = resultados(5)
'             Grid1.Cell(linea, 7).text = resultados(6)
'             Grid1.Cell(linea, 8).text = resultados(7)
'             Grid1.Cell(linea, 9).text = resultados(8)
'             Grid1.Cell(linea, 10).text = resultados(9)
'             Grid1.Cell(linea, 11).text = resultados(10)
'             Grid1.Cell(linea, 12).text = resultados(11)
             Grid1.Cell(LINEA, 13).text = resultados(5)
             Grid1.Cell(LINEA, 14).text = leefactura(LINEA)
'             Grid1.Cell(linea, 15).text = resultados("tipoconcepto")
'             Call leerdatoscontrato(resultados("contrato"))
'             If resultados("codigocontable") <> "" Then
'             cr_cuenta = resultados("codigocontable")
'             cr_crcc = "010000"
'             End If
'             If resultados("tipoconcepto") = "MT" Then
'             cr_cuenta = "35100022"
'             cr_crcc = "010000"
'             End If
'             If resultados("tipoconcepto") = "OC" Then
'             cr_cuenta = "35100026"
'             cr_crcc = "010000"
'             End If
'             If resultados("tipoconcepto") = "GC" Then
'             cr_cuenta = "35100027"
'
'             End If
'
'
'
'             Grid1.Cell(linea, 16).text = cr_cuenta
'             Grid1.Cell(linea, 17).text = cr_crcc
'
'
             
             Call leectacte(resultados(2))
            resultados.MoveNext
       
            Wend
            Grid1.AutoRedraw = True
            Grid1.Refresh
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

Sub eliminafactura(tipo, numero)
    Dim resultados As rdoResultset
    Dim csql As New rdoQuery
    
        Set csql.ActiveConnection = ventaslocal
        csql.sql = "delete "
        csql.sql = csql.sql + "FROM sv_documento_cabeza_" + localfiltro + " "
        csql.sql = csql.sql + "where tipo='" + tipo + "' and numero='" + numero + "' "
        csql.Execute
        Call sincronizadatos(csql.sql, ventaslocal, "")
        
        csql.sql = "delete "
        csql.sql = csql.sql + "FROM sv_documento_detalle_" + localfiltro + " "
        csql.sql = csql.sql + "where tipo='" + tipo + "' and numero='" + numero + "' "
        csql.Execute
        Call sincronizadatos(csql.sql, ventaslocal, "")
        
        csql.sql = "delete "
        csql.sql = csql.sql + "FROM sv_documento_pagos_" + localfiltro + " "
        csql.sql = csql.sql + "where tipo='" + tipo + "' and numero='" + numero + "' "
        csql.Execute
        Call sincronizadatos(csql.sql, ventaslocal, "")
        
        
        Set csql.ActiveConnection = gestionrubro
        csql.sql = "delete "
        csql.sql = csql.sql + "FROM l_movimientos_detalle_" + localfiltro + " "
        csql.sql = csql.sql + "where tipo='" + tipo + "' and numero='" + numero + "' "
        csql.Execute
        Call sincronizadatos(csql.sql, gestionrubro, "")
        

        
End Sub

Private Sub Grid1_KeyDown(KeyCode As Integer, ByVal Shift As Integer)
'If KeyCode = 46 Then
'Call eliminafactura(Grid1.Cell(Grid1.ActiveCell.Row, 1).text, Grid1.Cell(Grid1.ActiveCell.Row, 2).text)
'
'End If
'leer
End Sub

Sub grabafactura(numero, fecha, rut, monto, TIPOTR)
    Dim netos As Double
    Dim DH As String
    Dim exentos As Double
    Dim TIPOCON As String
    Dim CRCC As String
    Dim cuenta As String
    Dim DH2 As String
    Dim tipodoc As String
    Dim numerofolio As String
    
    cuenta = "11200029"
    
    Call grabardetallepagos(numero, rut, numero, TIPOTR, monto, fecha)
End Sub
Public Function LEERULTIMOFOLIO(tipo) As String

    Dim resultados As rdoResultset
    Dim csql As New rdoQuery
    
        Set csql.ActiveConnection = contadb

            csql.sql = "select IFNULL(max(numero),0) from movimientoscontables where mes = '" & Format(MES, "00") & "' AND año = '" & año & "' and tipo='" + tipo + "' "
            
            csql.Execute
    If csql.RowsAffected > 0 Then
    Set resultados = csql.OpenResultset
    
    
        LEERULTIMOFOLIO = Format(resultados(0) + 1, "0000000000")
    End If
    
End Function

Sub grabardetallepagos(numero, rut, FOLIO, TIPOTR, monto, fecha)
    Dim resultados As rdoResultset
    Dim csql As New rdoQuery
    Dim TIPODO As String
    Dim numerodo As String
    Dim fechado As String
    Dim cuentacontable As String
    Dim glosa As String
   If TIPOTR <> "ZE" Then
    
    Call grabardetallepagos_PAGOS(numero, rut, FOLIO, TIPOTR)
        
        Set csql.ActiveConnection = ventaslocal
        csql.sql = "SELECT tipo,documento,tipofactura,numerofactura,monto,fecha,rut,ifnull(fechafactura,'2010-01-01'),documento "
        csql.sql = csql.sql + "FROM sv_pagos_detalle_" + localfiltro + " WHERE numero='" + numero + "'  "
        csql.sql = csql.sql + "ORDER BY numerofactura "
        csql.Execute
    
        If csql.RowsAffected > 0 Then
            Set resultados = csql.OpenResultset
            While Not resultados.EOF
            linea2 = linea2 + 1
            TIPODO = resultados("tipofactura")
            
            numerodo = Format(resultados("numerofactura"), "0000000000")
            glosa = "CANCELA FACTURA" + leerNombrerut("11200029", resultados("rut"))
            fechado = resultados(5)
            If numerodo = "" Or fechado = "0000-00-00" Then
            TIPODO = resultados("tipo")
            numerodo = Format(resultados("documento"), "0000000000")
            fechado = resultados("fecha")
            glosa = "CANCELA FACTURA " + leerNombrerut("11200029", resultados("rut"))
            End If
            cuentacontable = cuentacliente
            
            If TIPODO = "PR" Then
            cuentacontable = "11200007"
            glosa = "CANCELA PROTESTO " + leerNombrerut("11200029", resultados("rut"))
            End If
            If TIPODO = "GR" Then
            cuentacontable = "23200014"
            glosa = "CHEQUE EN GARANTIA " + leerNombrerut("11200029", resultados("rut"))
            End If
            
            Call grabarcomprobante_lineas("PC", FOLIO, Format(linea2, "000"), resultados("fecha"), cuentacontable, cuentacliente, resultados("rut"), "", glosa, TIPODO, numerodo, fechado, fechado, resultados("monto"), "H", USUARIOSISTEMA, Format(resultados("fecha"), "MM"), Format(resultados("fecha"), "YYYY"), Format(fechasistema, "yyyy-mm-dd"), Time, resultados("rut"))
                resultados.MoveNext
            Wend
            resultados.Close
            Set resultados = Nothing
        ComboLOCAL.text = ComboLOCAL.List(0)
        End If
        localfiltro = Mid(ComboLOCAL.List(0), 1, 2)
        End If
        If TIPOTR = "ZE" Then
            TIPODO = "ZE"
            numerodo = numero
            fechado = fecha
            glosa = "CONTABILIZA ZETA"
            cuentacontable = "11500001"
            If empresaactiva = "01" Or empresaactiva = "29" Then
                cuentacontable = "11500200"
            End If
            
            Call grabarcomprobante_lineas("ZE", numero, Format(1, "000"), fecha, cuentacontable, "", "", "", glosa, TIPODO, numerodo, fechado, fechado, monto, "D", USUARIOSISTEMA, Format(fecha, "MM"), Format(fecha, "YYYY"), Format(fechasistema, "yyyy-mm-dd"), Time, "")
            TIPODO = "ZE"
            numerodo = numero
            fechado = fecha
            glosa = "CONTABILIZA ZETA"
            cuentacontable = "35100041"
            Call grabarcomprobante_lineas("ZE", numero, Format(1, "000"), fecha, cuentacontable, "", "", "", glosa, TIPODO, numerodo, fechado, fechado, monto, "H", USUARIOSISTEMA, Format(fecha, "MM"), Format(fecha, "YYYY"), Format(fechasistema, "yyyy-mm-dd"), Time, "")
            
        
        
        End If
End Sub
Sub grabardetallepagos_PAGOS(numero, rut, FOLIO, TIPOTR)
    Dim resultados2 As rdoResultset
    Dim csql2 As New rdoQuery
    Dim TIPODO As String
    Dim numerodo As String
    Dim fechado As String
    Dim cuentacontable As String
    Dim TIPOPAGO As String
    Dim monto As Double
    Dim glosa As String
    
    
        Set csql2.ActiveConnection = ventaslocal
        csql2.sql = "SELECT formapago,numerodocumento,banco,vencimiento,sum(monto),fecha,rut "
        csql2.sql = csql2.sql + "FROM sv_pagos_detalle_" + localfiltro + " WHERE numero='" + numero + "'  "
        csql2.sql = csql2.sql + "group by formapago,numerodocumento ORDER BY numerofactura "
        csql2.Execute
    linea2 = 0
        If csql2.RowsAffected > 0 Then
            Set resultados2 = csql2.OpenResultset
            While Not resultados2.EOF
            linea2 = linea2 + 1
            TIPOPAGO = resultados2(0)
            If TIPOPAGO = "1" Then
            cuentacontable = "11500001"
             If empresaactiva = "01" Or empresaactiva = "29" Then
                cuentacontable = "11500200"
            End If
            monto = resultados2(4)
            TIPODO = "CI"
            numerodo = numero
            fechado = resultados2(5)
            glosa = leerNombrerut("11200029", resultados2("rut"))
            End If
            If TIPOPAGO = "2" Then
            cuentacontable = "11500001"
             If empresaactiva = "01" Or empresaactiva = "29" Then
                cuentacontable = "11500200"
            End If
            
            monto = resultados2(4)
            TIPODO = "CH"
            numerodo = resultados2(1)
            fechado = resultados2(3)
            glosa = leerNombrerut("11200029", resultados2("rut"))
            End If
            If (TIPOPAGO <> "2" Or TIPOPAGO <> "1") Then
            cuentacontable = "11500001"
             If empresaactiva = "01" Or empresaactiva = "29" Then
                cuentacontable = "11500200"
            End If
            monto = resultados2(4)
            TIPODO = "PC"
            numerodo = resultados2(1)
            If numerodo = "" Then numerodo = numero
            fechado = resultados2(3)
            glosa = leerNombrerut("11200029", resultados2("rut"))
            End If
            Call grabarcomprobante_lineas("PC", FOLIO, Format(linea2, "000"), resultados2("fecha"), cuentacontable, cuentacliente, resultados2("rut"), "", glosa, TIPODO, numerodo, fechado, fechado, monto, "D", USUARIOSISTEMA, Format(resultados2("fecha"), "MM"), Format(resultados2("fecha"), "YYYY"), Format(fechasistema, "yyyy-mm-dd"), Time, resultados2("rut"))
                resultados2.MoveNext
            Wend
            resultados2.Close
            Set resultados2 = Nothing
        ComboLOCAL.text = ComboLOCAL.List(0)
        End If
        localfiltro = Mid(ComboLOCAL.List(0), 1, 2)
        
End Sub

Sub grabardetallefactura(LINEA, fecha2, caja)
    Dim TIPOCON As String
    Dim tipo2 As String
    Dim j As Integer
    Dim lin As Integer
    Dim ilas As Double
    Dim CRCC As String
    Dim DH As String
    Dim DH2 As String
    Dim tipodoc As String
    Dim fecha As Date
    Dim GLOSAS As String
    GLOSAS = leerNombreMayor(Grid1.Cell(LINEA, 16).text)
    
    fecha = Format(fecha2, "yyyy-mm-dd")
    
    
    campos(0, 0) = "tipo"
    campos(1, 0) = "numero"
    campos(2, 0) = "linea"
    campos(3, 0) = "rut"
    campos(4, 0) = "cuentadelmayor"
    campos(5, 0) = "glosa"
    campos(6, 0) = "monto"
    campos(7, 0) = "dh"
    campos(8, 0) = "centrodecosto"
    campos(9, 0) = "rutctacte"
    campos(10, 0) = "fechacreacion"
    campos(11, 0) = ""
    If Grid1.Cell(LINEA, 1).text = "FE" Then TIPOCON = "9": DH = "D": DH2 = "H": tipodoc = "FE"
    If Grid1.Cell(LINEA, 1).text = "FV" Then TIPOCON = "1": DH = "D": DH2 = "H": tipodoc = "FA"
    If Grid1.Cell(LINEA, 1).text = "NB" Then TIPOCON = "3": DH = "H": DH2 = "D": tipodoc = "NF"
    If Grid1.Cell(LINEA, 1).text = "NF" Then TIPOCON = "4": DH = "H": DH2 = "D": tipodoc = "NB"
    
    If Grid1.Cell(LINEA, 1).text = "FEE" And Grid1.Cell(LINEA, 0).text = "E" Then TIPOCON = "9": DH = "D": DH2 = "H": tipodoc = "FE"
    If Grid1.Cell(LINEA, 1).text = "FAE" And Grid1.Cell(LINEA, 0).text = "E" Then TIPOCON = "6": DH = "D": DH2 = "H": tipodoc = "FA"
    If Grid1.Cell(LINEA, 1).text = "NDE" And Grid1.Cell(LINEA, 0).text = "E" Then TIPOCON = "7": DH = "D": DH2 = "H": tipodoc = "ND"
    If Grid1.Cell(LINEA, 1).text = "NCE" And Grid1.Cell(LINEA, 0).text = "E" Then TIPOCON = "8": DH = "H": DH2 = "D": tipodoc = "NC"
    
    
    Rem  CALCULA netos
    
    lin = linea2
    pivote.MaxLength = 3
    pivote.text = lin
    
    Call ceros(pivote)
    campos(0, 1) = TIPOCON
    campos(1, 1) = Grid1.Cell(LINEA, 2).text
    campos(2, 1) = pivote.text
    campos(3, 1) = Mid(Grid1.Cell(LINEA, 3).text, 1, 9) + Mid(Grid1.Cell(LINEA, 3).text, 11, 1)
    campos(4, 1) = Grid1.Cell(LINEA, 16).text
    campos(5, 1) = GLOSAS
    campos(6, 1) = Replace(Grid1.Cell(LINEA, 6).text, ",", ".")
    campos(7, 1) = DH2
    campos(8, 1) = Grid1.Cell(LINEA, 17).text
    campos(9, 1) = ""
    campos(10, 1) = Format(fechasistema, "yyyy-mm-dd")
    campos(0, 2) = "facturasdeventas_detalle"
    condicion = ""
    op% = 2
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)
    Rem If Grid1.Cell(linea, 15).text = "99" Then
 
    Call grabarcomprobante_lineas(tipodoc, campos(1, 1), campos(2, 1), fecha, campos(4, 1), "", "", campos(8, 1), campos(5, 1), tipodoc, campos(1, 1), fecha, fecha, campos(6, 1), DH2, USUARIOSISTEMA, Format(fecha, "MM"), Format(fecha, "YYYY"), Format(fechasistema, "yyyy-mm-dd"), Time, campos(3, 1))
    Rem End If
Rem CALCULA ILAS refrescos

    
    
End Sub

Public Function leefactura(LINEA) As String

    Dim TIPOCON As String
    
    campos(0, 0) = "tipo"
    campos(1, 0) = "numero"
    campos(2, 0) = ""
    
    condicion = "tipo='" + Grid1.Cell(LINEA, 1).text + "' and numero='" + Format(Grid1.Cell(LINEA, 2).text, "0000000000") + "'"
    campos(0, 2) = "movimientoscontables"
    op = 5
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    
    Call sqlconta.sqlconta(op, condicion)
    If sqlconta.status = 0 Then
    leefactura = "1"
    Else
    leefactura = "0"
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
            csql.sql = csql.sql & "FROM " & Replace(clientesistema, "_", "ip_") & "ventas.sv_maestroclientes as mc "
            csql.sql = csql.sql & "WHERE mc.rut = '" & rut & "' AND mc.sucursal ='0'"
            
            csql.Execute
            Call sincronizadatos(csql.sql, gestion, "")
            
            
            csql.sql = "INSERT INTO " + clientesistema + "conta" + empresaactiva + ".saldosctacte "
            csql.sql = csql.sql & "(año,tipo,rut) "
            csql.sql = csql.sql & "SELECT '" + año + "','" + cuentacliente + "',mc.rut "
            csql.sql = csql.sql & "FROM " & Replace(clientesistema, "_", "ip_") & "ventas.sv_maestroclientes as mc "
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
Sub leerdatoscontrato(contrato)
    Dim resultados As rdoResultset
    Dim csql As New rdoQuery
    
        Set csql.ActiveConnection = gestion
        Rem select msp.tipolocal,msp.codigopropiedad,cd.contrato,mtp.cuenta,mp.crcc from ruta_arriendos00.contratos_arriendo_detalle as cd
        Rem left join ruta_arriendos00.maestro_subpropiedades as msp on msp.codigopropiedad=cd.propiedad and msp.codigo=cd.subpropiedad
        Rem left join ruta_arriendos.maestro_tipo_propiedad as mtp on mtp.codigo=msp.tipolocal
        Rem left join ruta_arriendos00.maestro_propiedades as mp on cd.propiedad=mp.codigopropiedad
        Rem where cd.contrato='0000000271' limit 0,1;

            csql.sql = "select mtp.cuenta,mp.crcc from ruta_arriendos" + empresaactiva + ".contratos_arriendo_detalle as cd "
            csql.sql = csql.sql & "left join ruta_arriendos" + empresaactiva + ".maestro_subpropiedades as msp on msp.codigopropiedad=cd.propiedad and msp.codigo=cd.subpropiedad "
            csql.sql = csql.sql & "left join ruta_arriendos.maestro_tipo_propiedad as mtp on mtp.codigo=msp.tipolocal "
            csql.sql = csql.sql & "left join ruta_arriendos" + empresaactiva + ".maestro_propiedades as mp on cd.propiedad=mp.codigopropiedad "
            csql.sql = csql.sql & "where cd.contrato='" + contrato + "' limit 0,1 "
            csql.Execute
cr_cuenta = ""
cr_crcc = ""
        If csql.RowsAffected > 0 Then
            Set resultados = csql.OpenResultset
            cr_cuenta = resultados(0)
            cr_crcc = resultados(1)
        resultados.Close
        Set resultados = Nothing
    End If

End Sub

