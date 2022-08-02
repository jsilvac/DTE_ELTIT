VERSION 5.00
Object = "{4777436C-EB8C-4596-98A8-EBCF98683969}#1.0#0"; "Flexcell.ocx"
Object = "{ADBBAED6-B16F-40DC-80DF-B44910CBA76C}#1.0#0"; "Frame-Xp.ocx"
Begin VB.Form flujocaja2 
   Caption         =   "Form1"
   ClientHeight    =   8100
   ClientLeft      =   60
   ClientTop       =   390
   ClientWidth     =   12765
   LinkTopic       =   "Form1"
   ScaleHeight     =   8100
   ScaleWidth      =   12765
   StartUpPosition =   3  'Windows Default
   Begin XPFrame.FrameXp FrameXp2 
      Height          =   1215
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   12495
      _ExtentX        =   22040
      _ExtentY        =   2143
      BackColor       =   49344
      CaptionEstilo3D =   1
      BackColor       =   49344
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
      Begin VB.CommandButton Command1 
         Caption         =   "LEER FLUJO"
         Height          =   495
         Left            =   6120
         TabIndex        =   7
         Top             =   600
         Width           =   1575
      End
      Begin VB.TextBox dato3 
         BackColor       =   &H00E1FFFD&
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
         Left            =   3240
         MaxLength       =   2
         TabIndex        =   5
         Tag             =   "fecha"
         Top             =   600
         Width           =   375
      End
      Begin VB.TextBox dato4 
         BackColor       =   &H00E1FFFD&
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
         Left            =   3600
         MaxLength       =   2
         TabIndex        =   4
         Tag             =   "fecha"
         Top             =   600
         Width           =   375
      End
      Begin VB.TextBox dato5 
         BackColor       =   &H00E1FFFD&
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
         Left            =   3960
         MaxLength       =   4
         TabIndex        =   3
         Tag             =   "fecha"
         Top             =   600
         Width           =   615
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "SEMANA"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   3240
         TabIndex        =   6
         Top             =   360
         Width           =   1335
      End
   End
   Begin XPFrame.FrameXp FrameXp1 
      Height          =   6015
      Left            =   120
      TabIndex        =   0
      Top             =   1440
      Width           =   12495
      _ExtentX        =   22040
      _ExtentY        =   10610
      BackColor       =   49344
      CaptionEstilo3D =   1
      BackColor       =   49344
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
         Height          =   5655
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   12255
         _ExtentX        =   21616
         _ExtentY        =   9975
         Cols            =   5
         DefaultFontSize =   8.25
         Rows            =   30
      End
   End
End
Attribute VB_Name = "flujocaja2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Call LEERFLUJO
End Sub

Private Sub Form_Load()
CARGAGRILLA
LEERFLUJO
End Sub
Sub CARGAGRILLA()
Rem DATOS DE LA COLUMNA
    Dim FORMATOGRILLA(10, 20)
    Grid1.DefaultFont.Size = 8
       
    FORMATOGRILLA(1, 1) = "FLUJO"
    FORMATOGRILLA(1, 2) = "GLOSA"
    FORMATOGRILLA(1, 3) = "ANTERIOR"
    FORMATOGRILLA(1, 4) = "LUNES"
    FORMATOGRILLA(1, 5) = "MARTES"
    FORMATOGRILLA(1, 6) = "MIERCOLES"
    FORMATOGRILLA(1, 7) = "JUEVES"
    FORMATOGRILLA(1, 8) = "VIERNES"
    Rem LARGO DE LOS DATOS
    FORMATOGRILLA(2, 1) = "10"
    FORMATOGRILLA(2, 2) = "20"
    FORMATOGRILLA(2, 3) = "10"
    FORMATOGRILLA(2, 4) = "10"
    FORMATOGRILLA(2, 5) = "10"
    FORMATOGRILLA(2, 6) = "10"
    FORMATOGRILLA(2, 7) = "10"
    FORMATOGRILLA(2, 8) = "10"
    
    Rem TIPO DE DATOS C=CEROS,S=STRING,N=NUMERICO,D=DATE,
    FORMATOGRILLA(3, 1) = "S"
    FORMATOGRILLA(3, 2) = "S"
    FORMATOGRILLA(3, 3) = "N"
    FORMATOGRILLA(3, 4) = "N"
    FORMATOGRILLA(3, 5) = "N"
    FORMATOGRILLA(3, 6) = "N"
    FORMATOGRILLA(3, 7) = "N"
    FORMATOGRILLA(3, 8) = "N"
    Rem FORMATO GRILLA
    FORMATOGRILLA(4, 3) = "###,###,##0"
    FORMATOGRILLA(4, 4) = "###,###,##0"
    FORMATOGRILLA(4, 5) = "###,###,##0"
    FORMATOGRILLA(4, 6) = "###,###,##0"
    FORMATOGRILLA(4, 7) = "###,###,##0"
    FORMATOGRILLA(4, 8) = "###,###,##0"
    Rem LOCCKED
    For k = 1 To 8
    FORMATOGRILLA(5, k) = "TRUE"
    Next k
    
    Grid1.Cols = 9
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
        If FORMATOGRILLA(3, k) = "N" Then Grid1.Column(k).Alignment = cellRightCenter
        If FORMATOGRILLA(3, k) = "D" Then Grid1.Column(k).CellType = cellCalendar
        
    Next k
   
   
   
    
    
End Sub

Sub LEERFLUJO()
    Dim resultados As rdoResultset
    Dim csql As New rdoQuery
    
        Set csql.ActiveConnection = conta
        csql.sql = "SELECT * "
        csql.sql = csql.sql + "FROM flujo_caja_titulos "
        csql.sql = csql.sql + "ORDER BY codigo,subcuenta "
        csql.Execute
        Grid1.Rows = 1
        If csql.RowsAffected > 0 Then
            Set resultados = csql.OpenResultset
            While Not resultados.EOF
                Grid1.Rows = Grid1.Rows + 1
                Grid1.Cell(Grid1.Rows - 1, 1).text = resultados(1)
                Grid1.Cell(Grid1.Rows - 1, 2).text = resultados(2)
                Call LEERFLUJOsemanal("2009-05-04", "2009-05-08", resultados(0) + resultados(1))
                
                resultados.MoveNext
            Wend
            resultados.Close
            Set resultados = Nothing
        
        End If
        
End Sub

Sub LEERFLUJOsemanal(fecha1, fecha2, codigo)
    Dim resultados As rdoResultset
    Dim csql As New rdoQuery
    Dim dia As Double
    
        Set csql.ActiveConnection = db
        csql.sql = "SELECT fecha,sum(monto) "
        csql.sql = csql.sql + "FROM flujo_caja where fecha between '" + fecha1 + "' and '" + fecha2 + "' and tipo='" + codigo + "' "
        csql.sql = csql.sql + "GROUP BY fecha ORDER BY tipo  "
        csql.Execute
      
        If csql.RowsAffected > 0 Then
            Set resultados = csql.OpenResultset
            While Not resultados.EOF
                
                dia = Weekday(Format(resultados(0), "yyyy-mm-dd"), vbUseSystemDayOfWeek) - 1
                Grid1.Cell(Grid1.Rows - 1, 3 + dia).text = resultados(1)
                resultados.MoveNext
            Wend
            resultados.Close
            Set resultados = Nothing
        
        End If
        
End Sub

