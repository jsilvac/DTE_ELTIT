VERSION 5.00
Object = "{4777436C-EB8C-4596-98A8-EBCF98683969}#1.0#0"; "FlexCell.ocx"
Object = "{ADBBAED6-B16F-40DC-80DF-B44910CBA76C}#1.0#0"; "Frame-Xp.ocx"
Begin VB.Form presu03 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Lista Determinacion Capital Propio"
   ClientHeight    =   9885
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   15270
   DrawStyle       =   5  'Transparent
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   659
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1018
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin XPFrame.FrameXp FrameQuickMenu 
      Height          =   615
      Left            =   600
      TabIndex        =   17
      Top             =   2160
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
         TabIndex        =   19
         Top             =   280
         Width           =   1335
      End
      Begin VB.CommandButton botonmisaccesos 
         Caption         =   "Permisos Modulo"
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   280
         Width           =   1455
      End
   End
   Begin XPFrame.FrameXp FrameXp3 
      Height          =   3015
      Left            =   120
      TabIndex        =   2
      Top             =   0
      Width           =   15015
      _ExtentX        =   26485
      _ExtentY        =   5318
      BackColor       =   16744576
      Caption         =   "Plan de Cuentas"
      CaptionEstilo3D =   1
      BackColor       =   16744576
      ForeColor       =   8438015
      BordeColor      =   14737632
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
      ColorTextShadow =   16744576
      Begin VB.OptionButton Option2 
         BackColor       =   &H00FF8080&
         Caption         =   "Acumulado"
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
         Left            =   2160
         TabIndex        =   16
         Top             =   1800
         Width           =   1575
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00FF8080&
         Caption         =   "Detallado"
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
         Left            =   240
         TabIndex        =   15
         Top             =   1800
         Value           =   -1  'True
         Width           =   1695
      End
      Begin VB.TextBox dato1 
         BackColor       =   &H00C0FFFF&
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
         Left            =   960
         Locked          =   -1  'True
         MaxLength       =   4
         TabIndex        =   14
         Top             =   240
         Width           =   735
      End
      Begin VB.TextBox dato2 
         BackColor       =   &H00C0FFFF&
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
         Left            =   960
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   6
         Tag             =   "nombre"
         Top             =   600
         Width           =   3015
      End
      Begin XPFrame.FrameXp FrameXp1 
         Height          =   2655
         Left            =   4080
         TabIndex        =   7
         Top             =   240
         Width           =   10815
         _ExtentX        =   19076
         _ExtentY        =   4683
         BackColor       =   16761024
         Caption         =   "Centros de Consumos"
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
         Begin FlexCell.Grid Grid2 
            Height          =   2385
            Left            =   120
            TabIndex        =   8
            Top             =   240
            Width           =   10575
            _ExtentX        =   18653
            _ExtentY        =   4207
            BackColor1      =   16573154
            BackColor2      =   16573154
            BackColorBkg    =   16761024
            Cols            =   5
            DefaultFontSize =   8.25
            GridColor       =   13193780
            Rows            =   30
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "DBclick Selecciona "
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
            Left            =   480
            TabIndex        =   9
            Top             =   6840
            Width           =   2535
         End
      End
      Begin XPFrame.FrameXp FrameXp7 
         Height          =   675
         Left            =   480
         TabIndex        =   12
         Top             =   960
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
            TabIndex        =   13
            Top             =   270
            Width           =   2865
         End
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Nombre"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   600
         Width           =   810
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Codigo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   810
      End
   End
   Begin VB.TextBox pivote 
      Height          =   285
      Left            =   240
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   2400
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
      Visible         =   0   'False
      Width           =   135
   End
   Begin XPFrame.FrameXp frmbala 
      Height          =   6855
      Left            =   120
      TabIndex        =   3
      Top             =   3000
      Width           =   15015
      _ExtentX        =   26485
      _ExtentY        =   12091
      BackColor       =   16744576
      Caption         =   "Presupuesto de Gastos"
      CaptionEstilo3D =   1
      BackColor       =   16744576
      BordeColor      =   14737632
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
      ColorTextShadow =   16744576
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FFC0C0&
         Caption         =   "IMPRIME"
         Height          =   375
         Left            =   2640
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   9000
         Width           =   2055
      End
      Begin FlexCell.Grid Grid1 
         Height          =   6375
         Left            =   120
         TabIndex        =   4
         Top             =   360
         Width           =   14775
         _ExtentX        =   26061
         _ExtentY        =   11245
         BackColor1      =   16573154
         BackColor2      =   16573154
         BackColorActiveCellSel=   16761024
         BackColorBkg    =   -2147483647
         BackColorFixed  =   16053492
         BackColorFixedSel=   -2147483639
         BackColorScrollBar=   -2147483639
         BackColorSel    =   16777215
         BorderColor     =   16761024
         CellBorderColor =   16744576
         CellBorderColorFixed=   16744576
         Cols            =   5
         DefaultFontSize =   8.25
         ForeColorFixed  =   16761024
         GridColor       =   -2147483645
         Rows            =   30
      End
   End
End
Attribute VB_Name = "presu03"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public saldoglobal As Double
Public ROW1 As Double
Dim totales(20) As Double
Dim totales2(20) As Double
Dim AÑOCONSULTA As String
Dim TOTALES20(20) As Double
Dim TOTALES30(20) As Double










Private Sub Command1_Click()
imprimir

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



Private Sub Form_Activate()
sqlconta.audit = True
sqlconta.programaactivo = Me.Caption

End Sub
Sub grabar(centro, CUENTAMAYOR, codigo, año)
    campos(0, 0) = "centro"
    campos(1, 0) = "cuentamayor"
    campos(2, 0) = "codigo"
    campos(3, 0) = "año"
    campos(4, 0) = ""
    campos(0, 1) = centro
    campos(1, 1) = CUENTAMAYOR
    campos(2, 1) = codigo
    campos(3, 1) = año
  
    campos(0, 2) = clientesistema + "conta" + empresaactiva + ".presupuestos_anuales"
    op = 2
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)
    
    End Sub
Sub modificar(centro, CUENTAMAYOR, codigo, año, MES, monto)
    CUENTAMAYOR = Mid(CUENTAMAYOR, 3, 8)
    campos(0, 0) = "ga" + MES
    campos(1, 0) = ""
    campos(0, 1) = monto
    condicion = "centro='" + centro + "' and cuentamayor='" + CUENTAMAYOR + "' and codigo='" & codigo & "' and año='" + año + "' "
    
    campos(0, 2) = clientesistema + "conta" + empresaactiva + ".presupuestos_anuales"
    op = 3
    sqlconta.response = campos
    Set sqlconta.conexion = contadb
    Call sqlconta.sqlconta(op, condicion)
    
    End Sub
 
Sub ELIMINAR(cuenta, codigo)
    campos(0, 2) = "presupuesto_detalle"
    condicion = "cuenta='" + cuenta + "' and codigo='" & codigo & "'"
    op = 4
    sqlconta.response = campos
    Set sqlconta.conexion = conta
    Call sqlconta.sqlconta(op, condicion)

    
End Sub
  

Private Sub Form_Load()
Call CENTRAR(Me)

'dibu1.FileName = App.path & "\archivo.gif"
'dibu2.FileName = App.path & "\archivo.gif"


    
    Call Conectar_BD

AÑOCONSULTA = Format(fechasistema, "YYYY")
Call CARGAPERMISO(Me.Name)
CARGAGRILLA
CARGAGRILLA2
Rem frmbala.Caption = "DETERMINACION CAPITAL PROPIO " + "01-01-" + Format(fechasistema, "YYYY")
Rem lan
Rem leecapital


LEERCODIGOS

For k = 2000 To Val(Format(fechasistema, "yyyy") + 10)
COMBOAÑO.AddItem k
Next k
COMBOAÑO.ListIndex = k - Format(fechasistema, "yyyy") - 2

End Sub



Sub CARGAGRILLA()
    Dim formatogrilla2(10, 20)
    formatogrilla2(1, 1) = "CODIGO"
    formatogrilla2(1, 2) = "NOMBRE"
    For k = 1 To 12
    formatogrilla2(1, k + 2) = MonthName(k)
    formatogrilla2(2, k + 2) = "8"
    formatogrilla2(3, k + 2) = "N"
    formatogrilla2(4, k + 2) = " ###,###,###,##0"
    formatogrilla2(5, k + 2) = "true"
    
    Next k
    
    
    
    Rem LARGO DE LOS DATOS
    formatogrilla2(2, 1) = "0"
    formatogrilla2(2, 2) = "15"
    formatogrilla2(2, 15) = "0"
    
    Rem TIPO DE DATOS C=CEROS,S=STRING,N=NUMERICO,D=DATE,
    formatogrilla2(3, 1) = "S"
    formatogrilla2(3, 2) = "S"
    formatogrilla2(3, 15) = "S"
    
    Rem FORMATO GRILLA
    
    
    Rem LOCCKED
    formatogrilla2(5, 1) = "TRUE"
    
    formatogrilla2(5, 2) = "TRUE"
    
    
    Rem VALOR MAXIMO
    
    Grid1.Cols = 16
    Grid1.Rows = 1
    Grid1.AllowUserResizing = False
    Grid1.DisplayFocusRect = False
    Grid1.ExtendLastCol = False
    Grid1.BoldFixedCell = False
    Grid1.DrawMode = cellOwnerDraw
    Grid1.Appearance = Flat
    Grid1.ScrollBarStyle = Flat
    Grid1.FixedRowColStyle = Flat
 '   Grid1.BackColorFixed = RGB(90, 158, 214)
'    Grid1.BackColorFixedSel = RGB(110, 180, 230)
'    Grid1.BackColorBkg = RGB(90, 158, 214)
'    Grid1.BackColorScrollBar = RGB(231, 235, 247)
'    Grid1.BackColor1 = RGB(231, 235, 247)
'    Grid1.BackColor2 = RGB(239, 243, 255)
'    Grid1.GridColor = RGB(148, 190, 231)
    Grid1.Column(0).Width = 0
    
    For k = 1 To Grid1.Cols - 1
        Grid1.Cell(0, k).text = formatogrilla2(1, k)
        
        
        Grid1.Column(k).Width = Val(formatogrilla2(2, k)) * 9
        Grid1.Column(k).MaxLength = Val(formatogrilla2(2, k)) + 2
        Grid1.Column(k).FormatString = formatogrilla2(4, k)
        Grid1.Column(k).Locked = formatogrilla2(5, k)
        If formatogrilla2(3, k) = "N" Then Grid1.Column(k).Alignment = cellRightCenter: Grid1.Column(k).Mask = cellNumeric
        If formatogrilla2(3, k) = "S" Then Grid1.Column(k).Alignment = cellLeftTop
        
        
        
        
        If formatogrilla2(3, k) = "D" Then Grid1.Column(k).CellType = cellCalendar
        
    Next k
 
 Grid1.Column(15).Width = 0
 
    End Sub
Sub CARGAGRILLA2()
    Dim formatogrilla2(10, 20)
    
    formatogrilla2(1, 1) = "CODIGO"
    formatogrilla2(1, 2) = "NOMBRE"
    For k = 1 To 12
    formatogrilla2(1, k + 2) = MonthName(k)
    formatogrilla2(2, k + 2) = "8"
    formatogrilla2(3, k + 2) = "N"
    formatogrilla2(4, k + 2) = " ###,###,###,##0"
    formatogrilla2(5, k + 2) = "TRUE"
    
    Next k
    
    
    
    Rem LARGO DE LOS DATOS
    formatogrilla2(2, 1) = "5"
    formatogrilla2(2, 2) = "20"
    
    Rem TIPO DE DATOS C=CEROS,S=STRING,N=NUMERICO,D=DATE,
    formatogrilla2(3, 1) = "S"
    formatogrilla2(3, 2) = "S"
    
    Rem FORMATO GRILLA
    
    
    Rem LOCCKED
    formatogrilla2(5, 1) = "TRUE"
    formatogrilla2(5, 2) = "TRUE"
    
    
    Rem VALOR MAXIMO
    
    Grid2.Cols = 15
    Grid2.Rows = 1
    Grid2.AllowUserResizing = False
    Grid2.DisplayFocusRect = False
    Grid2.ExtendLastCol = True
    Grid2.BoldFixedCell = False
    Grid2.DrawMode = cellOwnerDraw
    Grid2.Appearance = Flat
    Grid2.ScrollBarStyle = Flat
    Grid2.FixedRowColStyle = Flat
 '   Grid2.BackColorFixed = RGB(90, 158, 214)
'    grid2.BackColorFixedSel = RGB(110, 180, 230)
'    grid2.BackColorBkg = RGB(90, 158, 214)
'    grid2.BackColorScrollBar = RGB(231, 235, 247)
'    grid2.BackColor1 = RGB(231, 235, 247)
'    grid2.BackColor2 = RGB(239, 243, 255)
'    grid2.GridColor = RGB(148, 190, 231)
    Grid2.Column(0).Width = 0
    
    For k = 1 To Grid2.Cols - 1
        Grid2.Cell(0, k).text = formatogrilla2(1, k)
        
        
        Grid2.Column(k).Width = Val(formatogrilla2(2, k)) * 9
        Grid2.Column(k).MaxLength = Val(formatogrilla2(2, k))
        Grid2.Column(k).FormatString = formatogrilla2(4, k)
        Grid2.Column(k).Locked = formatogrilla2(5, k)
        If formatogrilla2(3, k) = "N" Then Grid2.Column(k).Alignment = cellRightCenter
        If formatogrilla2(3, k) = "S" Then Grid2.Column(k).Alignment = cellLeftTop
        
        
        If formatogrilla2(3, k) = "D" Then Grid2.Column(k).CellType = cellCalendar
        
    Next k
 
    End Sub



Sub leedetalle()

Dim resultados2 As rdoResultset
    Dim csql2 As New rdoQuery
    Dim rut As String
    Dim PASO As String
    Dim LINEAS As Double
    Dim saldo As Double
    Dim ultimo As String
    Dim cuenta As String
    Dim NOMBRE As String
    Dim DIFE As Double
    
    CARGAGRILLA
        
        Set csql2.ActiveConnection = conta
        csql2.sql = "SELECT pd.codigo,pd.nombre,pd.cuenta,cm.nombre "
        csql2.sql = csql2.sql + "FROM presupuesto_detalle as pd left join " + clientesistema + "conta" + empresaactiva + ".cuentasdelmayor as cm "
        csql2.sql = csql2.sql + "on (cm.año='" + COMBOAÑO.text + "' and cm.codigo=pd.cuenta) "
        
        csql2.sql = csql2.sql + "order by pd.cuenta,pd.nombre "
        csql2.Execute
        LINEAS = 0
        Grid1.AutoRedraw = False
        
        
        Grid1.Rows = 1
        If csql2.RowsAffected > 0 Then
        Set resultados2 = csql2.OpenResultset
        cuenta = resultados2(2)
        If IsNull(resultados2(3)) = False Then
        NOMBRE = resultados2(3)
        End If
        While Not resultados2.EOF
        
        If cuenta <> resultados2(2) Then
        Grid1.Rows = Grid1.Rows + 1
        Grid1.Cell(Grid1.Rows - 1, 1).text = cuenta
        Grid1.Cell(Grid1.Rows - 1, 2).text = NOMBRE
        Grid1.Range(Grid1.Rows - 1, 1, Grid1.Rows - 1, 14).FontBold = True
        Grid1.Range(Grid1.Rows - 1, 1, Grid1.Rows - 1, 14).BackColor = &HC0FFFF
        Grid1.Range(Grid1.Rows - 1, 1, Grid1.Rows - 1, 14).Locked = True
        
        
        For k = 1 To 12
        Grid1.Cell(Grid1.Rows - 1, k + 2).text = totales(k)
        Next k
        
        Rem REALES
        
        Grid1.Rows = Grid1.Rows + 1
        Grid1.Cell(Grid1.Rows - 1, 1).text = "99999999"
        Grid1.Cell(Grid1.Rows - 1, 2).text = " REALES"
        Grid1.Range(Grid1.Rows - 1, 1, Grid1.Rows - 1, 14).FontBold = True
        Grid1.Range(Grid1.Rows - 1, 1, Grid1.Rows - 1, 14).BackColor = &HC0FFFF
        For k = 1 To 12
        Grid1.Cell(Grid1.Rows - 1, k + 2).text = TOTALES20(k)
        
        Next k
        
        Rem DIFERENCIAS
        Grid1.Rows = Grid1.Rows + 1
        
        Grid1.Cell(Grid1.Rows - 1, 1).text = "99999999"
        Grid1.Cell(Grid1.Rows - 1, 2).text = " DIFERENCIAS "
        Grid1.Range(Grid1.Rows - 1, 1, Grid1.Rows - 1, 14).BackColor = &HC0FFFF
        
        
        For k = 1 To 12
        DIFE = CDbl(Grid1.Cell(Grid1.Rows - 3, k + 2).text) - CDbl(Grid1.Cell(Grid1.Rows - 2, k + 2).text)
        If DIFE < 0 Then
        Grid1.Cell(Grid1.Rows - 1, k + 2).ForeColor = &HFF&
        End If
        TOTALES20(k) = 0
        totales(k) = 0
        Grid1.Cell(Grid1.Rows - 1, k + 2).text = DIFE * -1
        Next k
        
        
        
        cuenta = resultados2(2)
         If IsNull(resultados2(3)) = False Then
        NOMBRE = resultados2(3)
        End If
        
        End If
Rem detalle
If Option1.Value = True Then
        
        Grid1.Rows = Grid1.Rows + 1
        
        Grid1.Cell(Grid1.Rows - 1, 1).text = resultados2(0)
        Grid1.Cell(Grid1.Rows - 1, 2).text = "  " + resultados2(1)
        Grid1.Cell(Grid1.Rows - 1, 15).text = "  " + resultados2(2)
       
        Rem presupuesto
End If
        Call cargapresupuestos(Grid1, Grid1.Rows - 1, dato1.text, resultados2(2), resultados2(0), Format(fechasistema, "yyyy"))
        
       Rem gastos
If Option1.Value = True Then
        Grid1.Rows = Grid1.Rows + 1
End If

        Call cargagastos(Grid1, Grid1.Rows - 1, dato1.text, resultados2(2), resultados2(0), Format(fechasistema, "yyyy"))
             
        resultados2.MoveNext
        Wend
     Rem presupuesto
        
        Grid1.Rows = Grid1.Rows + 1
        Grid1.Cell(Grid1.Rows - 1, 1).text = cuenta
        Grid1.Cell(Grid1.Rows - 1, 2).text = leerNombreMayor(cuenta)
        Grid1.Range(Grid1.Rows - 1, 1, Grid1.Rows - 1, 2).FontBold = True
        Grid1.Range(Grid1.Rows - 1, 1, Grid1.Rows - 1, 14).FontBold = True
        Grid1.Range(Grid1.Rows - 1, 1, Grid1.Rows - 1, 14).BackColor = &HC0FFFF
        Grid1.Range(Grid1.Rows - 1, 1, Grid1.Rows - 1, 14).Locked = True
        
        For k = 1 To 12
        Grid1.Cell(Grid1.Rows - 1, k + 2).text = totales(k)
        totales(k) = 0
        Next k
        
        Grid1.Rows = Grid1.Rows + 1
        Grid1.Cell(Grid1.Rows - 1, 1).text = ""
        Grid1.Cell(Grid1.Rows - 1, 2).text = "TOTAL PRESUPUESTO"
        Grid1.Range(Grid1.Rows - 1, 1, Grid1.Rows - 1, 2).FontBold = True
        Grid1.Range(Grid1.Rows - 1, 1, Grid1.Rows - 1, 14).FontBold = True
        Grid1.Range(Grid1.Rows - 1, 1, Grid1.Rows - 1, 14).BackColor = &HC0FFFF
        Grid1.Range(Grid1.Rows - 1, 1, Grid1.Rows - 1, 14).Locked = True
        
        
        
        For k = 1 To 12
        Grid1.Cell(Grid1.Rows - 1, k + 2).text = totales2(k)
        totales2(k) = 0
        Next k
        
        resultados2.Close
        Set resultados2 = Nothing

        End If
 
    Grid1.AutoRedraw = True
    Grid1.Refresh
    
    
     
End Sub

Sub cargapresupuestos(grilla As Grid, LINEA, centro, CUENTAMAYOR, codigo, año)
    Dim resultados As rdoResultset
    Dim csql As New rdoQuery
    Dim PASO As Integer
    Dim i As Integer
    
        Set csql.ActiveConnection = contadb
        csql.sql = "select ga01,ga02,ga03,ga04,ga05,ga06,ga07,ga08,ga09,ga10,ga11,ga12 from " + clientesistema + "conta" + empresaactiva + ".presupuestos_anuales "
        csql.sql = csql.sql + "where centro='" + centro + "' and cuentamayor='" + CUENTAMAYOR + "' and codigo='" + codigo + "' and año='" + año + "' "
        csql.Execute
       
        If csql.RowsAffected > 0 Then
            Set resultados = csql.OpenResultset
            
            While Not resultados.EOF
If Option1.Value = True Then
            grilla.Cell(LINEA, 3).text = resultados(0)
            grilla.Cell(LINEA, 4).text = resultados(1)
            grilla.Cell(LINEA, 5).text = resultados(2)
            grilla.Cell(LINEA, 6).text = resultados(3)
            grilla.Cell(LINEA, 7).text = resultados(4)
            grilla.Cell(LINEA, 8).text = resultados(5)
            grilla.Cell(LINEA, 9).text = resultados(6)
            grilla.Cell(LINEA, 10).text = resultados(7)
            grilla.Cell(LINEA, 11).text = resultados(8)
            grilla.Cell(LINEA, 12).text = resultados(9)
            grilla.Cell(LINEA, 13).text = resultados(10)
            grilla.Cell(LINEA, 14).text = resultados(11)
    
End If
        
        totales(1) = totales(1) + resultados(0)
        totales(2) = totales(2) + resultados(1)
        totales(3) = totales(3) + resultados(2)
        totales(4) = totales(4) + resultados(3)
        totales(5) = totales(5) + resultados(4)
        totales(6) = totales(6) + resultados(5)
        totales(7) = totales(7) + resultados(6)
        totales(8) = totales(8) + resultados(7)
        totales(9) = totales(9) + resultados(8)
        totales(10) = totales(10) + resultados(9)
        totales(11) = totales(11) + resultados(10)
        totales(12) = totales(12) + resultados(11)
        
        totales2(1) = totales2(1) + resultados(0)
        totales2(2) = totales2(2) + resultados(1)
        totales2(3) = totales2(3) + resultados(2)
        totales2(4) = totales2(4) + resultados(3)
        totales2(5) = totales2(5) + resultados(4)
        totales2(6) = totales2(6) + resultados(5)
        totales2(7) = totales2(7) + resultados(6)
        totales2(8) = totales2(8) + resultados(7)
        totales2(9) = totales2(9) + resultados(8)
        totales2(10) = totales2(10) + resultados(9)
        totales2(11) = totales2(11) + resultados(10)
        totales2(12) = totales2(12) + resultados(11)
        
            
            resultados.MoveNext
                
            Wend
            resultados.Close
            Set resultados = Nothing
        Else
      
        End If
 
        
End Sub
Sub cargagastos(grilla As Grid, LINEA, centro, CUENTAMAYOR, codigo, año)
    Dim resultados As rdoResultset
    Dim csql As New rdoQuery
    Dim PASO As Integer
    Dim i As Integer
    Dim DIFE As Double
    
    
        Set csql.ActiveConnection = contadb
        
'       select sum(monto),cuenta_presupuesto,centro_gastos,mes,año
'        from movimientoscontables where codigocuenta='47150001' and año='2009'
'        group by cuenta_presupuesto,centro_gastos,mes,año
'
        csql.sql = "select sum(monto),mes from " + clientesistema + "conta" + empresaactiva + ".movimientoscontables "
        csql.sql = csql.sql + "where centro_gastos='" + centro + "' and codigocuenta='" + CUENTAMAYOR + "' and cuenta_presupuesto='" + codigo + "' and año='" + año + "' "
        csql.sql = csql.sql + "group by mes "
        csql.Execute
If Option1.Value = True Then
        For k = 1 To 12
        grilla.Cell(LINEA, k + 2).text = "0"
        Next k
        grilla.Cell(LINEA, 1).text = "99999999"
        grilla.Cell(LINEA, 2).text = " REALES"
        grilla.Range(LINEA, 1, LINEA, 13).BackColor = &HCAE8CE
End If

        If csql.RowsAffected > 0 Then
            
            Set resultados = csql.OpenResultset
            
            While Not resultados.EOF
            
If Option1.Value = True Then
            grilla.Cell(LINEA, CDbl(resultados(1)) + 2).text = resultados(0)
End If
           TOTALES20(resultados(1)) = TOTALES20(resultados(1)) + resultados(0)
            
            resultados.MoveNext
                
            Wend
            resultados.Close
            Set resultados = Nothing
        Else
    
      
        End If
If Option1.Value = True Then
        Grid1.Rows = Grid1.Rows + 1
        grilla.Cell(Grid1.Rows - 1, 1).text = "99999999"
        grilla.Cell(Grid1.Rows - 1, 2).text = " DIFERENCIAS "
        grilla.Range(Grid1.Rows - 1, 1, Grid1.Rows - 1, 14).BackColor = &HF4F4F4
        
End If
        For k = 1 To 12
If Option1.Value = True Then
        DIFE = CDbl(grilla.Cell(LINEA - 1, k + 2).text) - CDbl(grilla.Cell(LINEA, k + 2).text)
        If DIFE < 0 Then
        grilla.Cell(Grid1.Rows - 1, k + 2).ForeColor = &HFF&
               
        End If
End If
If Option1.Value = True Then

        grilla.Cell(Grid1.Rows - 1, k + 2).text = CDbl(grilla.Cell(LINEA, k + 2).text) - CDbl(grilla.Cell(LINEA - 1, k + 2).text)
End If
        Next k
        
End Sub



Private Sub Grid2_DblClick()

dato1.text = Grid2.Cell(Grid2.ActiveCell.row, 1).text
dato2.text = Grid2.Cell(Grid2.ActiveCell.row, 2).text
Call leedetalle

End Sub

Private Sub Grid2_LeaveCell(ByVal row As Long, ByVal col As Long, NewRow As Long, NewCol As Long, Cancel As Boolean)
dato1.text = Grid2.Cell(NewRow, 1).text
dato2.text = Grid2.Cell(NewRow, 2).text


'
'If Mid(Grid2.Cell(NewRow, 1).text, 5, 4) > "0000" Then
'nivel.Caption = Grid2.Cell(NewRow, 1).text
'LBLNIVEL.Caption = Grid2.Cell(NewRow, 2).text
Call leedetalle
'End If
End Sub
Sub imprimir()
Dim titulo As String


titulo = "DETERMINACION CAPITAL PROPIO INICIAL AL " + Format(fechasistema, "dd-mm-yyyy")
Call CABEZAS2(titulo, "N", 1)
Grid1.DefaultFont.Size = 8
Grid1.PageSetup.Orientation = cellPortrait
Grid1.Range(0, 1, 0, Grid1.Cols - 1).Borders(cellEdgeBottom) = cellThick
Grid1.Range(0, 1, 0, Grid1.Cols - 1).Borders(cellEdgeLeft) = cellThick
Grid1.Range(0, 1, 0, Grid1.Cols - 1).Borders(cellEdgeTop) = cellThick
Grid1.Range(0, 1, 0, Grid1.Cols - 1).Borders(cellEdgeRight) = cellThick
Grid1.Range(0, 1, 0, Grid1.Cols - 1).Borders(cellInsideHorizontal) = cellThick
Grid1.Range(0, 1, 0, Grid1.Cols - 1).Borders(cellInsideVertical) = cellThick


Grid1.PageSetup.CenterHorizontally = True


Grid1.PageSetup.PrintFixedRow = True
Grid1.PageSetup.BottomMargin = 1
Grid1.PageSetup.TopMargin = 1
Grid1.PageSetup.LeftMargin = 1
Grid1.PageSetup.RightMargin = 0
Grid1.PageSetup.BlackAndWhite = True
Grid1.PageSetup.PrintGridlines = False
Grid1.PrintPreview 100

   
End Sub

Sub LEERCODIGOS()
    Dim resultados As rdoResultset
    Dim csql As New rdoQuery
    Dim PASO As Integer
    Dim i As Integer
    
        Set csql.ActiveConnection = conta
        csql.sql = "SELECT codigo,nombre "
        csql.sql = csql.sql + "FROM presupuesto_centros "
        csql.sql = csql.sql + "order by nombre "
        csql.Execute
        Grid2.Rows = 1
        If csql.RowsAffected > 0 Then
            Set resultados = csql.OpenResultset
            
            While Not resultados.EOF
            Grid2.Rows = Grid2.Rows + 1
            Grid2.Cell(Grid2.Rows - 1, 1).text = resultados(0)
            Grid2.Cell(Grid2.Rows - 1, 2).text = resultados(1)
            
            resultados.MoveNext
                
            Wend
            resultados.Close
            Set resultados = Nothing
 
        End If
 
        
End Sub
Private Sub botonmisaccesos_Click()
programafiltro = Me.Caption
misaccesos.Show
End Sub


Private Sub botonmisfavoritos_Click()
Call AgregaFavoritos(USUARIOSISTEMA, App.EXEName, Me.Name, Me.Caption)
End Sub
