VERSION 5.00
Object = "{4777436C-EB8C-4596-98A8-EBCF98683969}#1.0#0"; "Flexcell.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{ADBBAED6-B16F-40DC-80DF-B44910CBA76C}#1.0#0"; "Frame-Xp.ocx"
Begin VB.Form LVendedores 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Listado de Vendedores"
   ClientHeight    =   7335
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8610
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7335
   ScaleWidth      =   8610
   Begin MSAdodcLib.Adodc data 
      Height          =   375
      Left            =   360
      Top             =   6840
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   -1
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
      Height          =   6615
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   8355
      _ExtentX        =   14737
      _ExtentY        =   11668
      BackColor       =   16744576
      Caption         =   "Listado de Vendedores"
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
      Begin FlexCell.Grid Vendedores 
         Height          =   6195
         Left            =   60
         TabIndex        =   0
         Top             =   360
         Width           =   8235
         _ExtentX        =   14526
         _ExtentY        =   10927
         DefaultFontSize =   8.25
         Rows            =   2
         SelectionMode   =   1
         DateFormat      =   2
      End
   End
   Begin XPFrame.FrameXp frmImprimir 
      Height          =   375
      Left            =   5220
      TabIndex        =   2
      Top             =   6840
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
Attribute VB_Name = "LVendedores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Private formatogrilla(10, 10) As String

Private Sub Form_Activate()
    Principal.barraEstado.Panels(1).text = UCase(Me.Caption)
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then
        Unload Me
    End If
    If KeyCode = 38 Then
        If Screen.ActiveForm.ActiveControl.Name = "Vendedores" Then
            Unload Me
        End If
    End If
End Sub

Private Sub Form_Load()
    Dim tabla As String
    Call Centrar(Me)
    Call CargaGrillaVendedores(1, 4)
    tabla = "SELECT CONCAT( rut, '" & vbTab & "', nombre, '" & vbTab & "', comision) AS item "
    tabla = tabla & "FROM sv_maestrovendedores "
    tabla = tabla & "ORDER BY nombre ASC"
    Call ConectarControlData(data, servidor, baseVentas, usuario, password, tabla)
    Call cargaInforme(data, vendedores)
    
    vendedores.AddItem "", True
    vendedores.AddItem "CANTIDAD DE VENDEDORES      " & data.Recordset.RecordCount, True
    vendedores.Range(vendedores.Rows - 1, 1, vendedores.Rows - 1, vendedores.Cols - 1).Merge
    vendedores.Range(vendedores.Rows - 1, 1, vendedores.Rows - 1, vendedores.Cols - 1).Alignment = cellCenterCenter
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Principal.barraEstado.Panels(1).text = UCase(Principal.Caption)
    Call limpiaBarra(2)
End Sub


'****************************************************************************
'Formato de la Grilla Documentos
'****************************************************************************
    Private Sub CargaGrillaVendedores(ByVal row As Integer, ByVal col As Integer)
        Dim i As Integer
        Rem DATOS DE LA COLUMNA
        formatogrilla(1, 1) = "RUT"
        formatogrilla(1, 2) = "NOMBRE"
        formatogrilla(1, 3) = "COMISION"
      
        
        Rem LARGO DE LOS DATOS
        formatogrilla(2, 1) = "10"
        formatogrilla(2, 2) = "50"
        formatogrilla(2, 3) = "10"
    
        
        Rem TIPO DE DATOS C=CEROS,S=STRING,N=NUMERICO,D=DATE,
        formatogrilla(3, 1) = "N"
        formatogrilla(3, 2) = "S"
        formatogrilla(3, 3) = "N"
   
        
        Rem FORMATO GRILLA
        formatogrilla(4, 1) = "000000000"
        formatogrilla(4, 2) = ""
        formatogrilla(4, 3) = "##0.0"
        
        
        Rem LOCCKED
        formatogrilla(5, 1) = "FALSE"
        formatogrilla(5, 2) = "FALSE"
        formatogrilla(5, 3) = "FALSE"
       
        
        Rem VALOR MINIMO
        formatogrilla(6, 1) = ""
        formatogrilla(6, 2) = ""
        formatogrilla(6, 3) = ""
    
        
        Rem VALOR MAXIMO
        formatogrilla(7, 1) = ""
        formatogrilla(7, 2) = ""
        formatogrilla(7, 3) = ""
        
        
        Rem ANCHO
        formatogrilla(8, 1) = "5"
        formatogrilla(8, 2) = "10"
        formatogrilla(8, 3) = "30"
        formatogrilla(8, 4) = "8"
            
        vendedores.Cols = col
        vendedores.Rows = row
        vendedores.AllowUserResizing = False
        vendedores.DisplayFocusRect = False
        vendedores.ExtendLastCol = True
        vendedores.BoldFixedCell = False
        vendedores.DrawMode = cellOwnerDraw
        vendedores.Appearance = Flat
        vendedores.ScrollBarStyle = Flat
        vendedores.FixedRowColStyle = Flat
        vendedores.BackColorFixed = RGB(90, 158, 214)
        vendedores.BackColorFixedSel = RGB(110, 180, 230)
        vendedores.BackColorBkg = RGB(90, 158, 214)
        vendedores.BackColorScrollBar = RGB(231, 235, 247)
        vendedores.BackColor1 = RGB(231, 235, 247)
        vendedores.BackColor2 = RGB(239, 243, 255)
        vendedores.GridColor = RGB(148, 190, 231)
        
        vendedores.Column(0).Width = 0
        For i = 1 To col - 1
            vendedores.Cell(0, i).text = formatogrilla(1, i)
            vendedores.Column(i).Width = Val(formatogrilla(8, i)) * (vendedores.Cell(0, i).Font.Size + 1.25)
            vendedores.Column(i).MaxLength = Val(formatogrilla(2, i))
            vendedores.Column(i).FormatString = formatogrilla(4, i)
            vendedores.Column(i).Locked = formatogrilla(5, i)
            If formatogrilla(3, i) = "N" Then
                vendedores.Column(i).Alignment = cellRightCenter
            Else
                vendedores.Column(i).Alignment = cellLeftCenter
            End If
        Next i
        vendedores.Range(0, 0, 0, vendedores.Cols - 1).Alignment = cellCenterCenter
        vendedores.Enabled = True
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
        Call cabezaInforme("", vendedores, "LISTADO DE CUPOS, SALDOS, AUTORIZACION", 1)
        vendedores.PageSetup.HeaderMargin = 1
        vendedores.PageSetup.TopMargin = 1
        vendedores.PageSetup.LeftMargin = 1.5
        vendedores.PageSetup.RightMargin = 1
        vendedores.PageSetup.PrintFixedRow = True
        vendedores.PageSetup.BlackAndWhite = True
        vendedores.Range(0, 0, 0, vendedores.Cols - 1).Borders(cellEdgeBottom) = cellThin
        
        Call verificaImpresora(5, vendedores)
        
    End Sub

