VERSION 5.00
Object = "{4777436C-EB8C-4596-98A8-EBCF98683969}#1.0#0"; "FlexCell.ocx"
Object = "{ADBBAED6-B16F-40DC-80DF-B44910CBA76C}#1.0#0"; "Frame-Xp.ocx"
Begin VB.Form IMPUESTOS 
   Caption         =   "Form1"
   ClientHeight    =   4335
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6780
   LinkTopic       =   "Form1"
   ScaleHeight     =   4335
   ScaleWidth      =   6780
   StartUpPosition =   3  'Windows Default
   Begin XPFrame.FrameXp impuestos 
      Height          =   4290
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6780
      _ExtentX        =   11959
      _ExtentY        =   7567
      BackColor       =   49344
      Caption         =   "IMPUESTOS"
      CaptionEstilo3D =   1
      BackColor       =   49344
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
      Begin FlexCell.Grid Grid2 
         Height          =   3975
         Left            =   0
         TabIndex        =   1
         Top             =   240
         Width           =   6735
         _ExtentX        =   11880
         _ExtentY        =   7011
         Cols            =   5
         DefaultFontName =   "Arial"
         DefaultFontSize =   8.25
         ExtendLastCol   =   -1  'True
         Rows            =   30
      End
   End
End
Attribute VB_Name = "IMPUESTOS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub CARGAGRILLAexento()
    formatogrilla2(1, 1) = "CODIGO"
    formatogrilla2(1, 2) = "IMPUESTO"
    formatogrilla2(1, 3) = "MONTO"
    Rem LARGO DE LOS DATOS
    formatogrilla2(2, 1) = "10"
    
    formatogrilla2(2, 2) = "30"
    formatogrilla2(2, 3) = "10"
    
    Rem TIPO DE DATOS C=CEROS,S=STRING,N=NUMERICO,D=DATE,
    formatogrilla2(3, 1) = "C"
    formatogrilla2(3, 2) = "C"
    formatogrilla2(3, 3) = "N"
    Rem FORMATO GRILLA
    
    formatogrilla2(4, 3) = " ###,###,##0"
    Rem LOCCKED
    formatogrilla2(5, 1) = "TRUE"
    
    formatogrilla2(5, 2) = "TRUE"
    formatogrilla2(5, 3) = "FALSE"
    Rem VALOR MAXIMO
    formatogrilla2(7, 3) = "999999999"
    Grid2.Cols = 4
    Grid2.Rows = 1
    Grid2.AllowUserResizing = False
    Grid2.DisplayFocusRect = False
    Grid2.ExtendLastCol = True
    Grid2.BoldFixedCell = False
    Grid2.DrawMode = cellOwnerDraw
    Grid2.Appearance = Flat
    Grid2.ScrollBarStyle = Flat
    Grid2.FixedRowColStyle = Flat
    Grid2.BackColorFixed = RGB(90, 158, 214)
    Grid2.BackColorFixedSel = RGB(110, 180, 230)
    Grid2.BackColorBkg = RGB(90, 158, 214)
    Grid2.BackColorScrollBar = RGB(231, 235, 247)
    Grid2.BackColor1 = RGB(231, 235, 247)
    Grid2.BackColor2 = RGB(239, 243, 255)
    Grid2.GridColor = RGB(148, 190, 231)
    Grid2.Column(0).Width = 0
    
    For K = 1 To Grid2.Cols - 1
        Grid2.Cell(0, K).text = formatogrilla2(1, K)
        If K < 5 Then Grid2.Column(K).Width = Val(formatogrilla2(2, K)) * 8
        
        
        Rem Grid1.Column(K).Width = Val(formatogrilla(2, K)) * 9
        Grid2.Column(K).MaxLength = Val(formatogrilla2(2, K))
        Grid2.Column(K).FormatString = formatogrilla2(4, K)
        Grid2.Column(K).Locked = formatogrilla2(5, K)
        If formatogrilla2(3, K) = "N" Then Grid2.Column(K).Alignment = cellRightCenter
        If formatogrilla2(3, K) = "D" Then Grid2.Column(K).CellType = cellCalendar
        
    Next K
   Call leecuentas
    
    End Sub

