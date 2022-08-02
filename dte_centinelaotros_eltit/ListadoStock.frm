VERSION 5.00
Object = "{4777436C-EB8C-4596-98A8-EBCF98683969}#1.0#0"; "FlexCell.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{ADBBAED6-B16F-40DC-80DF-B44910CBA76C}#1.0#0"; "Frame-Xp.ocx"
Begin VB.Form ListadoStock 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Listado Retencion SII"
   ClientHeight    =   8430
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10245
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   8430
   ScaleWidth      =   10245
   Begin MSAdodcLib.Adodc data 
      Height          =   330
      Left            =   180
      Top             =   7980
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin XPFrame.FrameXp FrameXp2 
      Height          =   6675
      Left            =   60
      TabIndex        =   1
      Top             =   1200
      Width           =   10095
      _ExtentX        =   17806
      _ExtentY        =   11774
      BackColor       =   16744576
      Caption         =   "Informe"
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
      Begin FlexCell.Grid impresion 
         Height          =   6135
         Left            =   60
         TabIndex        =   0
         Top             =   420
         Width           =   9915
         _ExtentX        =   17489
         _ExtentY        =   10821
         Cols            =   5
         DefaultFontSize =   8.25
         Rows            =   30
         SelectionMode   =   1
      End
   End
   Begin XPFrame.FrameXp frmImprimir 
      Height          =   375
      Left            =   6720
      TabIndex        =   2
      Top             =   7980
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   661
      BackColor       =   49344
      Caption         =   "I   M   P   R   I   M   I   R"
      CaptionEstilo3D =   1
      BackColor       =   49344
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   1
   End
   Begin XPFrame.FrameXp frmMensual 
      Height          =   975
      Left            =   1620
      TabIndex        =   3
      Top             =   120
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   1720
      BackColor       =   16744576
      Caption         =   "Informe Mensual"
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
      Begin VB.ComboBox cmbMeses 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
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
         Height          =   315
         Left            =   3540
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   420
         Width           =   2175
      End
      Begin VB.Label lbl1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Seleccione Mes"
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
         Height          =   315
         Left            =   1380
         TabIndex        =   5
         Top             =   420
         Width           =   1875
      End
   End
End
Attribute VB_Name = "ListadoStock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Private fecha1 As String
    Private fecha2 As String

Private Sub cmbMeses_Click()
    Call cmbMeses_KeyPress(13)
End Sub

Private Sub cmbMeses_KeyPress(KeyAscii As Integer)
    Dim ultimo As String
    'Primero = DateSerial(Year(Now), Month(Now) + 0, 1)
    If KeyAscii = 13 Then
        fecha1 = Format(fechasistema, "yyyy") & "-" & Left(cmbMeses.List(cmbMeses.ListIndex), 2) & "-01"
        ultimo = DateSerial(Year(fecha1), Month(fecha1) + 1, 0)
        fecha2 = Format(fechasistema, "yyyy") & "-" & Left(cmbMeses.List(cmbMeses.ListIndex), 2) & "-" & Format(ultimo, "dd")
        Call generaInformeLS(data, impresion, fecha1, fecha2)
        SendKeys "{Tab}"
    End If
End Sub

Private Sub Form_Activate()
    Principal.barraEstado.Panels(1).text = UCase(Me.Caption)
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 27, 38
            Unload Me
    End Select
End Sub

Private Sub Form_Load()
    Call Centrar(Me)
    Call cargaMeses
    Call CargaGrillaInforme(1, 6)
End Sub

Private Sub cargaMeses()
    Dim i As Integer
    Dim fecha As String
    Dim cad As String
    For i = 1 To 12
        cad = Format(i, "00")
        fecha = "01-" & cad & "-" & Format(fechasistema, "yyyy")
        cmbMeses.AddItem cad & " - " & Format(fecha, "mmmm")
    Next i
End Sub

'****************************************************************************
'Formato de la Grilla Listado de Ventas
'****************************************************************************
    Private Sub CargaGrillaInforme(ByVal row As Integer, ByVal col As Integer)
        Dim formatoGrilla(10, 12) As String
        Dim i As Integer
        Rem DATOS DE LA COLUMNA
        formatoGrilla(1, 1) = "FECHA"
        formatoGrilla(1, 2) = "MOVIMIENTO"
        formatoGrilla(1, 3) = "INGRESOS"
        formatoGrilla(1, 4) = "EGRESOS"
        formatoGrilla(1, 5) = "SALDO"
        
        Rem LARGO DE LOS DATOS
        formatoGrilla(2, 1) = "10"
        formatoGrilla(2, 2) = "20"
        formatoGrilla(2, 3) = "12"
        formatoGrilla(2, 4) = "12"
        formatoGrilla(2, 5) = "12"
        
        Rem TIPO DE DATOS
        formatoGrilla(3, 1) = "C"
        formatoGrilla(3, 2) = "S"
        formatoGrilla(3, 3) = "N"
        formatoGrilla(3, 4) = "N"
        formatoGrilla(3, 5) = "N"
        
        Rem FORMATO GRILLA
        formatoGrilla(4, 1) = ""
        formatoGrilla(4, 2) = ""
        formatoGrilla(4, 3) = ""
        formatoGrilla(4, 4) = ""
        formatoGrilla(4, 5) = ""
        
        Rem LOCCKED
        formatoGrilla(5, 1) = "FALSE"
        formatoGrilla(5, 2) = "FALSE"
        formatoGrilla(5, 3) = "FALSE"
        formatoGrilla(5, 4) = "FALSE"
        formatoGrilla(5, 5) = "FALSE"
        
        Rem VALOR MINIMO
        formatoGrilla(6, 1) = ""
        formatoGrilla(6, 2) = ""
        formatoGrilla(6, 3) = ""
        formatoGrilla(6, 4) = ""
        formatoGrilla(6, 5) = ""
        
        Rem VALOR MAXIMO
        formatoGrilla(7, 1) = ""
        formatoGrilla(7, 2) = ""
        formatoGrilla(7, 3) = ""
        formatoGrilla(7, 4) = ""
        formatoGrilla(7, 5) = ""
        
        Rem ANCHO
        formatoGrilla(8, 1) = "10"
        formatoGrilla(8, 2) = "20"
        formatoGrilla(8, 3) = "12"
        formatoGrilla(8, 4) = "12"
        formatoGrilla(8, 5) = "12"
            
        impresion.Cols = col
        impresion.Rows = row
        impresion.Range(0, 0, impresion.Rows - 1, impresion.Cols - 1).Borders(cellEdgeBottom) = cellNone
        impresion.Range(0, 0, impresion.Rows - 1, impresion.Cols - 1).Borders(cellEdgeTop) = cellNone
        impresion.Range(0, 0, impresion.Rows - 1, impresion.Cols - 1).Borders(cellInsideVertical) = cellNone
        impresion.AllowUserResizing = False
        impresion.DisplayFocusRect = False
        impresion.ExtendLastCol = True
        impresion.BoldFixedCell = False
        impresion.DrawMode = cellOwnerDraw
        impresion.Appearance = Flat
        impresion.ScrollBarStyle = Flat
        impresion.FixedRowColStyle = Flat
        impresion.BackColorFixed = RGB(90, 158, 214)
        impresion.BackColorFixedSel = RGB(110, 180, 230)
        impresion.BackColorBkg = RGB(90, 158, 214)
        impresion.BackColorScrollBar = RGB(231, 235, 247)
        impresion.BackColor1 = RGB(231, 235, 247)
        impresion.BackColor2 = RGB(239, 243, 255)
        impresion.GridColor = RGB(148, 190, 231)
        
        impresion.Column(0).Width = 0
        
        For i = 1 To col - 1
            impresion.Cell(0, i).text = formatoGrilla(1, i)
            impresion.Column(i).Width = Val(formatoGrilla(8, i)) * (impresion.Cell(0, i).Font.Size + 1.25)
            impresion.Column(i).MaxLength = Val(formatoGrilla(2, i))
            impresion.Column(i).FormatString = formatoGrilla(4, i)
            impresion.Column(i).Locked = formatoGrilla(5, i)
            If formatoGrilla(3, i) = "N" Then
                impresion.Column(i).Alignment = cellRightCenter
            End If
            If formatoGrilla(3, i) = "S" Then
                impresion.Column(i).Alignment = cellLeftCenter
            End If
            If formatoGrilla(3, i) = "C" Then
                impresion.Column(i).Alignment = cellCenterCenter
            End If
        Next i
        
        impresion.Range(0, 1, 0, impresion.Cols - 1).Alignment = cellCenterCenter
        impresion.Range(0, 1, 0, impresion.Cols - 1).FontBold = True
        impresion.Range(0, 1, 0, impresion.Cols - 1).Borders(cellEdgeTop) = cellThin
        impresion.Range(0, 1, 0, impresion.Cols - 1).Borders(cellEdgeBottom) = cellThin
    End Sub
'****************************************************************************
'Formato de la Grilla Listado de Ventas
'****************************************************************************

    Private Sub frmImprimir_BarMouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
        Call cambiaColor(frmImprimir)
        frmImprimir.CaptionEstilo3D = Raised
    End Sub
    
    Private Sub frmImprimir_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
        Call cambiaColor(frmImprimir)
        frmImprimir.CaptionEstilo3D = Inserted
        Call imprimir
    End Sub
    
    Private Sub imprimir()
        Dim i As Long
        impresion.AutoRedraw = False
        impresion.PageSetup.HeaderMargin = 2
    
        impresion.PageSetup.TopMargin = 2
        impresion.PageSetup.LeftMargin = 2
        impresion.PageSetup.RightMargin = 2
        impresion.PageSetup.BottomMargin = 2
        
        impresion.PageSetup.FooterMargin = 2
        impresion.PageSetup.BlackAndWhite = True
        impresion.PageSetup.PrintFixedRow = True
        'impresion.PageSetup.Orientation = cellLandscape
        
        Call verificaImpresora(5, impresion)
        
        impresion.AutoRedraw = True
    End Sub

