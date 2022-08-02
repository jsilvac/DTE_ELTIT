VERSION 5.00
Object = "{4777436C-EB8C-4596-98A8-EBCF98683969}#1.0#0"; "FlexCell.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{ADBBAED6-B16F-40DC-80DF-B44910CBA76C}#1.0#0"; "Frame-Xp.ocx"
Begin VB.Form LCupSalAut 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Listado Cupos, Saldos, Autorización"
   ClientHeight    =   7260
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10455
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7260
   ScaleWidth      =   10455
   Begin MSAdodcLib.Adodc data 
      Height          =   330
      Left            =   240
      Top             =   6840
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   2143
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
   Begin XPFrame.FrameXp FrameXp1 
      Height          =   6555
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   10215
      _ExtentX        =   18018
      _ExtentY        =   11562
      BackColor       =   16744576
      Caption         =   "Resumen"
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
      Begin FlexCell.Grid Documentos 
         Height          =   6075
         Left            =   120
         TabIndex        =   0
         Top             =   360
         Width           =   9975
         _ExtentX        =   17595
         _ExtentY        =   10716
         DefaultFontSize =   8.25
         Rows            =   2
         SelectionMode   =   1
         DateFormat      =   2
      End
   End
   Begin XPFrame.FrameXp frmImprimir 
      Height          =   375
      Left            =   7020
      TabIndex        =   2
      Top             =   6780
      Width           =   3285
      _ExtentX        =   5794
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
End
Attribute VB_Name = "LCupSalAut"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Private FORMATOGRILLA(10, 10) As String

Private Sub Form_Activate()
    Principal.barraEstado.Panels(1).text = UCase(Me.Caption)
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then
        Unload Me
    End If
    If KeyCode = 38 Then
        If Screen.ActiveForm.ActiveControl.Name = "Documentos" Then
            Unload Me
        End If
    End If
End Sub

Private Sub Form_Load()
    Dim tabla As String
    Call Centrar(Me)
    Call CargaGrillaDocumentos(1, 5)
    tabla = "SELECT CONCAT(rut, '" & vbTab & "', nombre, '" & vbTab & "', credito, '" & vbTab & "', cupodirecto) AS item "
    tabla = tabla & "FROM sv_maestroclientes "
    tabla = tabla & "ORDER BY rut ASC"
    Call ConectarControlData(data, servidor, baseVentas, usuario, password, tabla)
    Call cargaInforme(data, Documentos)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Principal.barraEstado.Panels(1).text = UCase(Principal.Caption)
    Call limpiaBarra(2)
End Sub


'****************************************************************************
'Formato de la Grilla Documentos
'****************************************************************************
    Private Sub CargaGrillaDocumentos(ByVal Row As Integer, ByVal Col As Integer)
        Dim i As Integer
        Rem DATOS DE LA COLUMNA
        FORMATOGRILLA(1, 1) = "RUT"
        FORMATOGRILLA(1, 2) = "RAZÓN SOCIAL"
        FORMATOGRILLA(1, 3) = "SIT.COMERCIAL"
        FORMATOGRILLA(1, 4) = "CRÉDITO"
        
        Rem LARGO DE LOS DATOS
        FORMATOGRILLA(2, 1) = "10"
        FORMATOGRILLA(2, 2) = ""
        FORMATOGRILLA(2, 3) = ""
        FORMATOGRILLA(2, 4) = "9"
        
        Rem TIPO DE DATOS
        FORMATOGRILLA(3, 1) = "N"
        FORMATOGRILLA(3, 2) = "S"
        FORMATOGRILLA(3, 3) = "C"
        FORMATOGRILLA(3, 4) = "N"
        
        Rem FORMATO GRILLA
        ''''''''''''''''''''''''
        FORMATOGRILLA(4, 1) = "0000000000"
        FORMATOGRILLA(4, 2) = ""
        FORMATOGRILLA(4, 3) = ""
        FORMATOGRILLA(4, 4) = "$ ###,###,##0"
        
        Rem LOCCKED
        FORMATOGRILLA(5, 1) = "TRUE"
        FORMATOGRILLA(5, 2) = "TRUE"
        FORMATOGRILLA(5, 3) = "TRUE"
        FORMATOGRILLA(5, 4) = "TRUE"
        
        Rem VALOR MINIMO
        FORMATOGRILLA(6, 1) = ""
        FORMATOGRILLA(6, 2) = ""
        FORMATOGRILLA(6, 3) = ""
        FORMATOGRILLA(6, 4) = ""
        
        Rem VALOR MAXIMO
        FORMATOGRILLA(7, 1) = ""
        FORMATOGRILLA(7, 2) = ""
        FORMATOGRILLA(7, 3) = ""
        FORMATOGRILLA(7, 4) = ""
        
        Rem ANCHO
        FORMATOGRILLA(8, 1) = "10"
        FORMATOGRILLA(8, 2) = "30"
        FORMATOGRILLA(8, 3) = "10"
        FORMATOGRILLA(8, 4) = "9"
            
        Documentos.Cols = Col
        Documentos.Rows = Row
        Documentos.AllowUserResizing = False
        Documentos.DisplayFocusRect = False
        Documentos.ExtendLastCol = True
        Documentos.BoldFixedCell = False
        Documentos.DrawMode = cellOwnerDraw
        Documentos.Appearance = Flat
        Documentos.ScrollBarStyle = Flat
        Documentos.FixedRowColStyle = Flat
        Documentos.BackColorFixed = RGB(90, 158, 214)
        Documentos.BackColorFixedSel = RGB(110, 180, 230)
        Documentos.BackColorBkg = RGB(90, 158, 214)
        Documentos.BackColorScrollBar = RGB(231, 235, 247)
        Documentos.BackColor1 = RGB(231, 235, 247)
        Documentos.BackColor2 = RGB(239, 243, 255)
        Documentos.GridColor = RGB(148, 190, 231)
        
        Documentos.Column(0).Width = 0
        For i = 1 To Col - 1
            Documentos.Cell(0, i).text = FORMATOGRILLA(1, i)
            Documentos.Column(i).Width = Val(FORMATOGRILLA(8, i)) * (Documentos.Cell(0, i).Font.Size + 1.25)
            Documentos.Column(i).MaxLength = Val(FORMATOGRILLA(2, i))
            Documentos.Column(i).FormatString = FORMATOGRILLA(4, i)
            Documentos.Column(i).Locked = FORMATOGRILLA(5, i)
            If FORMATOGRILLA(3, i) = "N" Then
                Documentos.Column(i).Alignment = cellRightCenter
            End If
            If FORMATOGRILLA(3, i) = "S" Then
                Documentos.Column(i).Alignment = cellLeftCenter
            End If
            If FORMATOGRILLA(3, i) = "C" Then
                Documentos.Column(i).Alignment = cellCenterCenter
            End If
        Next i
        Documentos.Range(0, 0, 0, Documentos.Cols - 1).Alignment = cellCenterCenter
        Documentos.Enabled = True
    End Sub
'****************************************************************************
'Formato de la Grilla Documentos
'****************************************************************************

    Private Sub frmImprimir_BarMouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
        frmImprimir.ColorBarraAbajo = &HFFC0C0
        frmImprimir.ColorBarraArriba = &H800000
        frmImprimir.CaptionEstilo3D = Raised
    End Sub

    Private Sub frmImprimir_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
        frmImprimir.ColorBarraAbajo = &H800000
        frmImprimir.ColorBarraArriba = &HFFC0C0
        frmImprimir.CaptionEstilo3D = Inserted
        Call imprimir
    End Sub

    Private Sub imprimir()
        Call cabezaInforme("", Documentos, "LISTADO DE CUPOS, SALDOS, AUTORIZACION", 1)
        Documentos.PageSetup.HeaderMargin = 1
        Documentos.PageSetup.TopMargin = 1
        Documentos.PageSetup.LeftMargin = 1.5
        Documentos.PageSetup.RightMargin = 1
        Documentos.PageSetup.PrintFixedRow = True
        Documentos.PageSetup.BlackAndWhite = True
        Documentos.Range(0, 0, 0, Documentos.Cols - 1).Borders(cellEdgeBottom) = cellThin
        
        Call verificaImpresora(5, Documentos)
        
    End Sub

