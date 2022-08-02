VERSION 5.00
Object = "{4777436C-EB8C-4596-98A8-EBCF98683969}#1.0#0"; "FlexCell.ocx"
Object = "{085FBDF3-00C5-421B-B762-1D57299A2B89}#1.0#0"; "CLBUTN.OCX"
Object = "{ADBBAED6-B16F-40DC-80DF-B44910CBA76C}#1.0#0"; "Frame-Xp.ocx"
Begin VB.Form factoring01 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Control Documentos Electronicos Recibidos Factorizados"
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
   Begin XPFrame.FrameXp FrameXp4 
      Height          =   615
      Left            =   11880
      TabIndex        =   14
      Top             =   8040
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
         TabIndex        =   16
         Top             =   280
         Width           =   1455
      End
      Begin VB.CommandButton botonmisfavoritos 
         Caption         =   "Mis Favoritos"
         Height          =   255
         Left            =   1800
         TabIndex        =   15
         Top             =   280
         Width           =   1335
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
      TabIndex        =   1
      Top             =   6120
      Visible         =   0   'False
      Width           =   135
   End
   Begin XPFrame.FrameXp FrameXp1 
      Height          =   8610
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   15150
      _ExtentX        =   26723
      _ExtentY        =   15187
      BackColor       =   16761024
      Caption         =   ""
      CaptionEstilo3D =   1
      BackColor       =   16761024
      ForeColor       =   8438015
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
      Begin FlexCell.Grid Grid2 
         Height          =   240
         Left            =   720
         TabIndex        =   12
         Top             =   9480
         Visible         =   0   'False
         Width           =   1230
         _ExtentX        =   2170
         _ExtentY        =   423
         Cols            =   5
         DefaultFontSize =   8.25
         Rows            =   30
      End
      Begin XPFrame.FrameXp FrameXp3 
         Height          =   1530
         Left            =   135
         TabIndex        =   4
         Top             =   225
         Width           =   14910
         _ExtentX        =   26300
         _ExtentY        =   2699
         BackColor       =   16744576
         Caption         =   ""
         CaptionEstilo3D =   1
         BackColor       =   16744576
         ForeColor       =   8438015
         ColorBarraArriba=   4194304
         ColorBarraAbajo =   4194304
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Begin VB.CommandButton Command4 
            BackColor       =   &H00FFC0C0&
            Caption         =   "EXCEL"
            Height          =   330
            Left            =   12840
            Style           =   1  'Graphical
            TabIndex        =   26
            Top             =   720
            Width           =   1890
         End
         Begin VB.CommandButton Command1 
            BackColor       =   &H00FFC0C0&
            Caption         =   "IMPRIMIR"
            Height          =   330
            Left            =   12840
            Style           =   1  'Graphical
            TabIndex        =   25
            Top             =   1080
            Width           =   1890
         End
         Begin VB.OptionButton Option5 
            BackColor       =   &H00FFC0C0&
            Caption         =   "Factura No Acuse Recibo"
            Height          =   255
            Left            =   9960
            TabIndex        =   23
            Top             =   1920
            Visible         =   0   'False
            Width           =   2415
         End
         Begin VB.OptionButton Option4 
            BackColor       =   &H00FFC0C0&
            Caption         =   "Factura No Recepcionada"
            Height          =   255
            Left            =   9960
            TabIndex        =   22
            Top             =   1680
            Visible         =   0   'False
            Width           =   2415
         End
         Begin VB.CommandButton Command3 
            Caption         =   "Recontabiliza automatico"
            Height          =   375
            Left            =   12480
            TabIndex        =   21
            Top             =   2640
            Visible         =   0   'False
            Width           =   1935
         End
         Begin VB.TextBox tipotxt 
            Height          =   285
            Left            =   8040
            TabIndex        =   20
            Top             =   840
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.OptionButton Option3 
            BackColor       =   &H00FFC0C0&
            Caption         =   "xml no Recibido"
            Height          =   255
            Left            =   9960
            TabIndex        =   19
            Top             =   1440
            Visible         =   0   'False
            Width           =   2415
         End
         Begin VB.OptionButton Option2 
            BackColor       =   &H00FFC0C0&
            Caption         =   "Pagados"
            Height          =   255
            Left            =   9600
            TabIndex        =   18
            Top             =   480
            Value           =   -1  'True
            Width           =   2415
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00FFC0C0&
            Caption         =   "Todos"
            Height          =   255
            Left            =   9600
            TabIndex        =   17
            Top             =   240
            Width           =   2415
         End
         Begin VB.CommandButton Command2 
            BackColor       =   &H00FFC0C0&
            Caption         =   "Generar Informe"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   12840
            Style           =   1  'Graphical
            TabIndex        =   13
            Top             =   360
            Width           =   1905
         End
         Begin XPFrame.FrameXp FrameXp8 
            Height          =   975
            Left            =   135
            TabIndex        =   5
            Top             =   240
            Width           =   6375
            _ExtentX        =   11245
            _ExtentY        =   1720
            BackColor       =   14737632
            Caption         =   "Rangos de Fecha"
            CaptionEstilo3D =   1
            BackColor       =   14737632
            ForeColor       =   8438015
            ColorBarraArriba=   4194304
            ColorBarraAbajo =   4194304
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
               TabIndex        =   6
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
               TabIndex        =   10
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
               TabIndex        =   9
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
               TabIndex        =   8
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
               TabIndex        =   7
               Top             =   480
               Width           =   1935
            End
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00FF8080&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "x Tipo"
            Height          =   255
            Left            =   7200
            TabIndex        =   24
            Top             =   840
            Visible         =   0   'False
            Width           =   735
         End
      End
      Begin XPFrame.FrameXp FrameXp2 
         Height          =   6270
         Left            =   135
         TabIndex        =   3
         Top             =   1800
         Width           =   15000
         _ExtentX        =   26458
         _ExtentY        =   11060
         BackColor       =   16744576
         Caption         =   ""
         CaptionEstilo3D =   1
         BackColor       =   16744576
         ForeColor       =   8438015
         BordeColor      =   4194304
         ColorBarraArriba=   4194304
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
            Height          =   5925
            Left            =   120
            TabIndex        =   11
            Top             =   240
            Width           =   14715
            _ExtentX        =   25956
            _ExtentY        =   10451
            Cols            =   5
            DefaultFontSize =   8.25
            Rows            =   30
         End
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " DOBLE CLIC EN EL NUMERO PARA VER CONTENIDO DEL  DTE"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   27
         Top             =   8160
         Width           =   11535
      End
   End
End
Attribute VB_Name = "factoring01"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub Check1_Click()
'If Check1.Value = True Then
'Timer1.Enabled = True
'Else
'Timer1.Enabled = False
'
'End If

End Sub

Private Sub Command1_Click()
imprimir

End Sub

Private Sub COMMAND2_Click()
CARGAGRILLA
leer



End Sub

Private Sub Command3_Click()
'Option2.Value = True
'
'For k = 1 To Grid1.Rows - 1
'If Grid1.Cell(k, 7).text = "0" Then
'Call eliminaerrores(Grid1.Cell(k, 11).text)
'Grid1.Cell(k, 7).SetFocus
'Grid1.Refresh
'
'End If
'
'Next k
'leer

End Sub

Private Sub Command4_Click()
Call Grid1.ExportToExcel("", True, True)
End Sub

Private Sub cool_Button3_Click()
Call retornofecha(desdefecha, hastafecha)
End Sub


 











Private Sub Form_Load()
CENTRAR Me


    
    Call Conectar_BD

    sc = 0
CARGAGRILLA


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







Sub imprimir()
Dim titulo As String



Call CABEZAS2(titulo, "N", 1)
Grid1.DefaultFont.Size = 8
Grid1.PageSetup.Orientation = cellLandscape
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
Sub CABEZA()
    

End Sub




Private Sub opciones_GotFocus()

MANUAL.SetFocus

End Sub
Sub CARGAGRILLA()
Rem DATOS DE LA COLUMNA
    Dim FORMATOGRILLA(10, 20)
    Grid1.DefaultFont.Size = 8
       
    FORMATOGRILLA(1, 1) = "TD"
    FORMATOGRILLA(1, 2) = "NUMERO"
    FORMATOGRILLA(1, 3) = "FECHA"
    FORMATOGRILLA(1, 4) = "TOTAL"
    FORMATOGRILLA(1, 5) = "EMISOR"
    FORMATOGRILLA(1, 6) = "NOMBRE EMISOR"
    FORMATOGRILLA(1, 7) = "EMAIL EMISOR"
    FORMATOGRILLA(1, 8) = "CESIONARIO"
    FORMATOGRILLA(1, 9) = "NOMBRE CESIONARIO"
    FORMATOGRILLA(1, 10) = "FECHA CESION"
    FORMATOGRILLA(1, 11) = "ESTADO CESION"
    FORMATOGRILLA(1, 12) = "ABONO"
    FORMATOGRILLA(1, 13) = "DATOS PAGO"
    
     
    Rem LARGO DE LOS DATOS
    FORMATOGRILLA(2, 1) = "4"
    FORMATOGRILLA(2, 2) = "9"
    FORMATOGRILLA(2, 3) = "9"
    FORMATOGRILLA(2, 4) = "9"
    FORMATOGRILLA(2, 5) = "9"
    FORMATOGRILLA(2, 6) = "20"
    FORMATOGRILLA(2, 7) = "10"
    FORMATOGRILLA(2, 8) = "10"
    FORMATOGRILLA(2, 9) = "15"
    FORMATOGRILLA(2, 10) = "10"
    FORMATOGRILLA(2, 11) = "15"
    FORMATOGRILLA(2, 12) = "8"
    FORMATOGRILLA(2, 13) = "15"
    
    
    Rem TIPO DE DATOS C=CEROS,S=STRING,N=NUMERICO,D=DATE,
    FORMATOGRILLA(3, 1) = "S"
    FORMATOGRILLA(3, 2) = "N"
    FORMATOGRILLA(3, 3) = "D"
    FORMATOGRILLA(3, 4) = "N"
    FORMATOGRILLA(3, 5) = "S"
    FORMATOGRILLA(3, 6) = "S"
    FORMATOGRILLA(3, 7) = "S"
    FORMATOGRILLA(3, 8) = "S"
    FORMATOGRILLA(3, 9) = "S"
    FORMATOGRILLA(3, 10) = "D"
    FORMATOGRILLA(3, 11) = "S"
    FORMATOGRILLA(3, 12) = "N"
    FORMATOGRILLA(3, 13) = "S"
    
    Rem FORMATO GRILLA
    FORMATOGRILLA(4, 6) = "###,###,###,##0"
    FORMATOGRILLA(4, 4) = "###,###,###,##0"
    
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
    FORMATOGRILLA(5, 10) = "TRUE"
    FORMATOGRILLA(5, 11) = "TRUE"
    FORMATOGRILLA(5, 12) = "TRUE"
    FORMATOGRILLA(5, 13) = "TRUE"
    
    Grid1.Cols = 14
    Grid1.Rows = 1
    
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
        If FORMATOGRILLA(3, k) = "N" Then
            Grid1.Column(k).Alignment = cellRightCenter
            Grid1.Column(k).Mask = cellNumeric
        End If
        If FORMATOGRILLA(3, k) = "D" Then Grid1.Column(k).CellType = cellCalendar
        
    Next k
  Grid1.PageSetup.BlackAndWhite = False
  
End Sub




Private Sub monto_Click()
End Sub

Private Sub leermutuos()

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
 
        Set csql.ActiveConnection = conta
        'cSql.SQL = "SELECT fecha,tipo,numero,linea,codigocuenta,glosacontable,tipodocumento,numerodocumento,fechadocumento,fechavencimiento,monto,dh,centrocosto,tipoctacte,rutctacte "
'        dia = 1
'        MES = 1
'        año = 2005
        csql.sql = "SELECT banco,empresa,tipo,sum(if(evento='1',monto,monto*-1)),evento "
        csql.sql = csql.sql + "FROM inver_fondosmutuos group by banco,empresa,tipo "
        csql.Execute
        total = 0
        total2 = 0
        LINEA = Grid1.Rows - 1
        Grid1.Rows = Grid1.Rows + 1
        
        Grid1.Cell(Grid1.Rows - 1, 1).text = "INVERSIONES FONDOS MUTUOS "
        
        If csql.RowsAffected > 0 Then
        Set resultados = csql.OpenResultset
        
         While Not resultados.EOF
          
          
             Grid1.Rows = Grid1.Rows + 1
             Grid1.Cell(Grid1.Rows - 1, 1).text = leerbanco(resultados(0))
             Grid1.Cell(Grid1.Rows - 1, 2).text = leerempresa(resultados(1))
             Grid1.Cell(Grid1.Rows - 1, 3).text = leerdeposito(resultados(2))
             Grid1.Cell(Grid1.Rows - 1, 7).text = resultados(3)
             
             
             
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
    objReportTitle.text = ""
    
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
Private Sub botonmisaccesos_Click()
programafiltro = Me.Caption
misaccesos.Show
End Sub


Private Sub botonmisfavoritos_Click()
Call AgregaFavoritos(USUARIOSISTEMA, App.EXEName, Me.Name, Me.Caption)
End Sub

Private Sub Grid1_DblClick()
If Grid1.ActiveCell.col = 2 Then
electro88.tipo.text = Grid1.Cell(Grid1.ActiveCell.row, 1).text

electro88.FOLIO.text = Grid1.Cell(Grid1.ActiveCell.row, 2).text
electro88.cliente.text = Format(Replace(Grid1.Cell(Grid1.ActiveCell.row, 5).text, "-", ""), "0000000000")
electro88.Show vbModal
End If

'If Grid1.ActiveCell.col = 8 Then
'If Grid1.Cell(Grid1.ActiveCell.row, 7).text = "0" And Grid1.Cell(Grid1.ActiveCell.row, 11).text <> "0" Then
'MsgBox "SE REPROCESARA EL DOCUMENTO "
'Call eliminaerrores(Grid1.Cell(Grid1.ActiveCell.row, 11).text)
'Leer
'Else
'MsgBox "DOCUMENTO NO HA LLEGADO POR CORREO "
'
'End If


 

End Sub

Private Sub Option1_Click()
COMMAND2_Click

End Sub

Private Sub Option2_Click()
COMMAND2_Click

End Sub

Private Sub Option3_Click()
COMMAND2_Click

End Sub
Private Function estarecepcionada(tipo, numero, rut, monto) As String


Dim resultados As rdoResultset
    
    Dim csql As New rdoQuery
    If tipo <> "33" Then Exit Function
    
    tipo = "FAE"
 
        Set csql.ActiveConnection = contadb
        
        If CONFI_EMPRESAFAE = "00" Then
        csql.sql = "SELECT '1' "
        csql.sql = csql.sql + "FROM " + clientesistema + "gestion" + "00" + ".l_ordendecompra_detalle_facturas_" + "00 "
        csql.sql = csql.sql + "where tipo='" + tipo + "' and numero='" + Format(numero, "0000000000") + "' and total='" & monto & "' and rut='" + rut + "' "
        csql.sql = csql.sql + "Union "
        csql.sql = csql.sql + "SELECT '1' "
        csql.sql = csql.sql + "FROM " + clientesistema + "gestion" + "00" + ".l_ordendecompra_detalle_facturas_" + "25 "
        csql.sql = csql.sql + "where tipo='" + tipo + "' and numero='" + Format(numero, "0000000000") + "' and total='" & monto & "' and rut='" + rut + "' "
        csql.sql = csql.sql + "Union "
        csql.sql = csql.sql + "SELECT '1' "
        csql.sql = csql.sql + "FROM " + clientesistema + "gestion" + "00" + ".l_ordendecompra_detalle_facturas_" + "41 "
        csql.sql = csql.sql + "where tipo='" + tipo + "' and numero='" + Format(numero, "0000000000") + "' and total='" & monto & "' and rut='" + rut + "' "
        
        End If
        
        If CONFI_EMPRESAFAE = "01" Then
        csql.sql = "SELECT '1' "
        csql.sql = csql.sql + "FROM " + clientesistema + "gestion" + "01" + ".l_ordendecompra_detalle_facturas_" + "01 "
        csql.sql = csql.sql + "where tipo='" + tipo + "' and numero='" + Format(numero, "0000000000") + "' and total='" & monto & "' and rut='" + rut + "' "
        csql.sql = csql.sql + "Union "
        csql.sql = csql.sql + "SELECT '1' "
        csql.sql = csql.sql + "FROM " + clientesistema + "gestion" + "01" + ".l_ordendecompra_detalle_facturas_" + "20 "
        csql.sql = csql.sql + "where tipo='" + tipo + "' and numero='" + Format(numero, "0000000000") + "' and total='" & monto & "' and rut='" + rut + "' "
        csql.sql = csql.sql + "Union "
        csql.sql = csql.sql + "SELECT '1' "
        csql.sql = csql.sql + "FROM " + clientesistema + "gestion" + "01" + ".l_ordendecompra_detalle_facturas_" + "39 "
        csql.sql = csql.sql + "where tipo='" + tipo + "' and numero='" + Format(numero, "0000000000") + "' and total='" & monto & "' and rut='" + rut + "' "
        
        End If
        If CONFI_EMPRESAFAE = "03" Then
        csql.sql = "SELECT '1' "
        csql.sql = csql.sql + "FROM " + clientesistema + "gestion" + "03" + ".l_ordendecompra_detalle_facturas_" + "03 "
        csql.sql = csql.sql + "where tipo='" + tipo + "' and numero='" + Format(numero, "0000000000") + "' and total='" & monto & "' and rut='" + rut + "' "
        
        End If
        
        If CONFI_EMPRESAFAE = "02" Then
        csql.sql = "SELECT '1' "
        csql.sql = csql.sql + "FROM " + clientesistema + "gestion" + "02" + ".l_ordendecompra_detalle_facturas_" + "02 "
        csql.sql = csql.sql + "where tipo='" + tipo + "' and numero='" + Format(numero, "0000000000") + "' and total='" & monto & "' and rut='" + rut + "' "
        
        End If
        If CONFI_EMPRESAFAE = "17" Then
        csql.sql = "SELECT '1' "
        csql.sql = csql.sql + "FROM " + clientesistema + "gestion" + "00" + ".l_ordendecompra_detalle_facturas_" + "17 "
        csql.sql = csql.sql + "where tipo='" + tipo + "' and numero='" + Format(numero, "0000000000") + "' and total='" & monto & "' and rut='" + rut + "' "
        csql.sql = csql.sql + "Union "
        csql.sql = csql.sql + "SELECT '1' "
        csql.sql = csql.sql + "FROM " + clientesistema + "gestion" + "00" + ".l_ordendecompra_detalle_facturas_" + "18 "
        csql.sql = csql.sql + "where tipo='" + tipo + "' and numero='" + Format(numero, "0000000000") + "' and total='" & monto & "' and rut='" + rut + "' "
        
        End If
        If CONFI_EMPRESAFAE = "42" Then
        csql.sql = "SELECT '1' "
        csql.sql = csql.sql + "FROM " + clientesistema + "gestion" + "00" + ".l_ordendecompra_detalle_facturas_" + "42 "
        csql.sql = csql.sql + "where tipo='" + tipo + "' and numero='" + Format(numero, "0000000000") + "' and total='" & monto & "' and rut='" + rut + "' "
        csql.sql = csql.sql + "Union "
        csql.sql = csql.sql + "SELECT '1' "
        csql.sql = csql.sql + "FROM " + clientesistema + "gestion" + "00" + ".l_ordendecompra_detalle_facturas_" + "44 "
        csql.sql = csql.sql + "where tipo='" + tipo + "' and numero='" + Format(numero, "0000000000") + "' and total='" & monto & "' and rut='" + rut + "' "
        csql.sql = csql.sql + "Union "
        csql.sql = csql.sql + "SELECT '1' "
        csql.sql = csql.sql + "FROM " + clientesistema + "gestion" + "00" + ".l_ordendecompra_detalle_facturas_" + "45 "
        csql.sql = csql.sql + "where tipo='" + tipo + "' and numero='" + Format(numero, "0000000000") + "' and total='" & monto & "' and rut='" + rut + "' "
        
        End If
        If CONFI_EMPRESAFAE = "29" Then
        csql.sql = "SELECT '1' "
        csql.sql = csql.sql + "FROM " + clientesistema + "gestion" + "15" + ".l_ordendecompra_detalle_facturas_" + "55 "
        csql.sql = csql.sql + "where tipo='" + tipo + "' and numero='" + Format(numero, "0000000000") + "' and total='" & monto & "' and rut='" + rut + "' "
        
        End If
        
        
        csql.Execute
        
        estarecepcionada = "0"
        If csql.RowsAffected > 0 Then
        estarecepcionada = "1"
        End If
            
     If estarecepcionada = "0" Then
     If leerempresaproveedor(rut) <> "" Then
     estarecepcionada = "1"
     End If
     
     End If
     
     
 
End Function

Private Sub eliminaerrores(ARCHIVO)

Dim resultados As rdoResultset
    
    Dim csql As New rdoQuery
    
        Set csql.ActiveConnection = contadb
        csql.sql = "delete from " + clientesistema + "fae" + CONFI_EMPRESAFAE + ".sv_dte" + CONFI_EMPRESAFAE + "_recibidos where nombrearchivo='" + ARCHIVO + "' "
        csql.Execute
        
        Set csql.ActiveConnection = contadb
        csql.sql = "update " + clientesistema + "fae" + CONFI_EMPRESAFAE + ".sv_recepcion_dte" + CONFI_EMPRESAFAE + " set archivo_respuesta='' where archivo='" + ARCHIVO + "' "
        csql.Execute
        
            
     
 
End Sub
Public Function lee_factura_de_compra(tipo, numero, rut) As Boolean
    Dim csql As New rdoQuery
    Dim CUENTA2 As String
    Rem On Error GoTo no:
    If tipo = "33" Then tipo = "4"
    If tipo = "61" Then tipo = "6"
    Set csql.ActiveConnection = contadb
    csql.sql = "select numero from " & clientesistema & "conta" & empresaactiva & ".facturasdecompras "
    csql.sql = csql.sql & "where tipo='" + tipo + "' and numero='" & Format(numero, "0000000000") & "' and rut='" + rut + "' "
    csql.Execute
    lee_factura_de_compra = False
    If csql.RowsAffected > 0 Then
    lee_factura_de_compra = True
    End If
    Exit Function
no:
   
End Function


Sub leer()
Dim NOMBRE_ARCHIVO As String
Dim cadena_mensaje As String
    Dim resultados As rdoResultset
    Dim csql As New rdoQuery
    Dim csql1 As New rdoQuery
    Dim resultados1 As rdoResultset
    Dim LINEA As Double
    Dim destinatarioNombre As String
    Dim destinatarioCorreo As String
    Dim CadenaMensaje As String
    Dim nombrearchivo As String
    Dim empresacontable As String
    Dim loc As String
    Set csql1.ActiveConnection = conta
    Set csql.ActiveConnection = contadb
    
    Dim desde As String
    Dim hasta As String
    desde = Format(desdefecha, "yyyy-mm-dd")
    hasta = Format(hastafecha, "yyyy-mm-dd")
    
    csql1.sql = "SELECT codigoempresa,nombre,empresafae,rut,nombre "
    csql1.sql = csql1.sql & " from " & clientesistema & "conta.maestroempresas "
    csql1.sql = csql1.sql & " where  codigoempresa='" & empresaactiva & "' and  empresafae<>'' "
    csql1.Execute
            
    If csql1.RowsAffected > 0 Then
    Set resultados1 = csql1.OpenResultset
    While Not resultados1.EOF
                   
       
        
        empresacontable = resultados1(0)
        loc = resultados1(2)
        
        
        csql.sql = "SELECT *  "
        csql.sql = csql.sql & " ,IFNULL((SELECT IFNULL(SUM(monto),'0') FROM "
        csql.sql = csql.sql & clientesistema & "conta" & empresacontable & ".movimientoscontables"
        csql.sql = csql.sql & " WHERE (tipodocumento='FC'"
        csql.sql = csql.sql & " AND numerodocumento=LPAD(dte.numero,10,0)"
        csql.sql = csql.sql & " AND codigocuenta='" & CUENTAPROVEEDOR & "' "
        csql.sql = csql.sql & " AND rutctacte=LPAD(REPLACE(dte.cedente_rut,'-',''),10,0)"
        csql.sql = csql.sql & " AND dh='D')OR ((glosacontable LIKE CONCAT('FACTURA PUBLICIDAD %',dte.numero,'%'))"
        csql.sql = csql.sql & " AND codigocuenta='23100026'"
        csql.sql = csql.sql & " AND rutctacte=LPAD(REPLACE(dte.cedente_rut,'-',''),10,0)"
        csql.sql = csql.sql & " AND dh='D') GROUP BY numerodocumento"
        csql.sql = csql.sql & " ),0) AS abono"
        
        csql.sql = csql.sql & " ,IFNULL((SELECT CONCAT(tipo,' ',numero,' ',fecha) FROM "
        csql.sql = csql.sql & clientesistema & "conta" & empresacontable & ".movimientoscontables"
        csql.sql = csql.sql & " WHERE (tipodocumento='FC'"
        csql.sql = csql.sql & " AND numerodocumento=LPAD(dte.numero,10,0)"
        csql.sql = csql.sql & " AND codigocuenta='23100026'"
        csql.sql = csql.sql & " AND rutctacte=LPAD(REPLACE(dte.cedente_rut,'-',''),10,0)"
        csql.sql = csql.sql & " AND dh='D')"
        csql.sql = csql.sql & " OR ((glosacontable LIKE CONCAT('FACTURA PUBLICIDAD %',dte.numero,'%'))"
        csql.sql = csql.sql & " AND codigocuenta='23100026'"
        csql.sql = csql.sql & " AND rutctacte=LPAD(REPLACE(dte.cedente_rut,'-',''),10,0)"
        csql.sql = csql.sql & " AND dh='D') GROUP BY numerodocumento"
        csql.sql = csql.sql & " ),0) AS pago"
        csql.sql = csql.sql & " FROM " & clientesistema & "fae" & loc & ".sv_dte_cedidos_" & loc & " as dte"
        
        csql.sql = csql.sql & " where fecha between '" & desde & "' "
        csql.sql = csql.sql & " and '" & hasta & "' "
         If Option2.Value = True Then
            csql.sql = csql.sql & " having abono<>0 "
        End If
        
        csql.sql = csql.sql & " ORDER BY abono, dte.cesion_fecha ASC"
        csql.sql = csql.sql & " ,RIGHT(pago,10) DESC,cesion_fecha"
        
       
        
        csql.Execute
        Grid1.Rows = 1
        Grid1.AutoRedraw = False
        If csql.RowsAffected > 0 Then
            Set resultados = csql.OpenResultset
            While Not resultados.EOF
                Grid1.Rows = Grid1.Rows + 1
                Grid1.Cell(Grid1.Rows - 1, 1).text = resultados(0)
                Grid1.Cell(Grid1.Rows - 1, 2).text = resultados(1)
                Grid1.Cell(Grid1.Rows - 1, 3).text = resultados(2)
                Grid1.Cell(Grid1.Rows - 1, 4).text = resultados(3)
                Grid1.Cell(Grid1.Rows - 1, 5).text = resultados(5)
                Grid1.Cell(Grid1.Rows - 1, 6).text = resultados(6)
                Grid1.Cell(Grid1.Rows - 1, 7).text = resultados(7)
                Grid1.Cell(Grid1.Rows - 1, 8).text = resultados(8)
                Grid1.Cell(Grid1.Rows - 1, 9).text = resultados(9)
                Grid1.Cell(Grid1.Rows - 1, 10).text = resultados("cesion_fecha")
                Grid1.Cell(Grid1.Rows - 1, 11).text = resultados(14)
                Grid1.Cell(Grid1.Rows - 1, 12).text = resultados("abono")
                Grid1.Cell(Grid1.Rows - 1, 13).text = resultados("pago")
                If Grid1.Cell(Grid1.Rows - 1, 13).text <> "0" Then
                    Grid1.Range(Grid1.Rows - 1, 1, Grid1.Rows - 1, Grid1.Cols - 1).BackColor = vbRed
                    Grid1.Range(Grid1.Rows - 1, 1, Grid1.Rows - 1, Grid1.Cols - 1).FontBold = True
                
                End If
                
                resultados.MoveNext
            Wend
            resultados.Close
            Set resultados = Nothing
                    
            
           End If
         
        
        
                Grid1.AutoRedraw = True
                Grid1.Refresh
                resultados1.MoveNext
                 Wend
                resultados1.Close
                Set resultados1 = Nothing
            
            End If
             
         
       
End Sub
