VERSION 5.00
Object = "{4777436C-EB8C-4596-98A8-EBCF98683969}#1.0#0"; "Flexcell.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{ADBBAED6-B16F-40DC-80DF-B44910CBA76C}#1.0#0"; "Frame-Xp.ocx"
Begin VB.Form listaDocumentos 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   6840
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12090
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   ScaleHeight     =   6840
   ScaleWidth      =   12090
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin XPFrame.FrameXp frmDocumentos 
      Height          =   6675
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   11955
      _ExtentX        =   21087
      _ExtentY        =   11774
      BackColor       =   16744576
      Caption         =   "Lista Documentos"
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
      Begin FlexCell.Grid listaVentas 
         Height          =   6255
         Left            =   45
         TabIndex        =   1
         Top             =   360
         Width           =   11775
         _ExtentX        =   20770
         _ExtentY        =   11033
         Cols            =   5
         DefaultFontSize =   8.25
         Rows            =   30
         SelectionMode   =   1
      End
      Begin XPFrame.FrameXp frmCerrar 
         Height          =   255
         Left            =   11640
         TabIndex        =   2
         Top             =   30
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   450
         BackColor       =   49344
         Caption         =   "X"
         CaptionEstilo3D =   1
         BackColor       =   49344
         ColorBarraArriba=   8388608
         ColorBarraAbajo =   16761024
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
      End
      Begin MSAdodcLib.Adodc data 
         Height          =   330
         Left            =   180
         Top             =   6300
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
   End
End
Attribute VB_Name = "listaDocumentos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Private formatogrilla(10, 10) As String
    Private TIPO As String
    Private NUMERO As String
    Private fecha As String
    Public tabla As String
    Public formulario As String
    Public datos As String

Private Sub Form_Activate()
    Select Case datos
        Case "ventas"
            Call CargaGrillaDocumentos(1, 11)
        Case "protestos"
            Call CargaGrillaProtestos(1, 6)
        Case "prorrogas"
            Call CargaGrillaProrrogas(1, 6)
        Case "cheques"
            Call CargaGrillaCheques(1, 7)
    End Select
    Select Case formulario
        Case "auditoria"
            Call ConectarControlData(data, servidor, baseVentas & localAuditoria, usuario, password, tabla)
        Case Else
            Call ConectarControlData(data, servidor, baseVentas & empresaActiva, usuario, password, tabla)
    End Select
    listaVentas.Rows = 1
    listaVentas.AutoRedraw = False
    If data.Recordset.RecordCount > 0 Then
        data.Recordset.MoveFirst
        While Not data.Recordset.EOF
            Select Case datos
                Case "ventas"
                    listaVentas.AddItem data.Recordset.Fields("item1") & vbTab & leerNombreCliente(data.Recordset.Fields("rut")) & vbTab & Replace(data.Recordset.Fields("item2"), ",", ".") & vbTab & data.Recordset.Fields("caja"), True
                Case Else
                    listaVentas.AddItem Replace(data.Recordset.Fields("item"), ",", "."), True
            End Select
            data.Recordset.MoveNext
        Wend
    End If
    listaVentas.AutoRedraw = True
    listaVentas.Refresh
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    'Call CargaGrillaDocumentos(1, 10)
End Sub

    Private Sub frmCerrar_BarMouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
        frmCerrar.ColorBarraAbajo = &H800000
        frmCerrar.ColorBarraArriba = &HFFC0C0
        frmCerrar.CaptionEstilo3D = Raised
    End Sub

    Private Sub frmCerrar_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
        frmCerrar.ColorBarraAbajo = &HFFC0C0
        frmCerrar.ColorBarraArriba = &H800000
        frmCerrar.CaptionEstilo3D = Inserted
        Unload Me
    End Sub

'****************************************************************************
'Formato de la Grilla Listado de Ventas
'****************************************************************************
    Private Sub CargaGrillaDocumentos(ByVal row As Integer, ByVal col As Integer)
        Dim i As Integer
        Rem DATOS DE LA COLUMNA
        formatogrilla(1, 1) = "DOCUMENTO"
        formatogrilla(1, 2) = "FECHA"
        formatogrilla(1, 3) = "RUT"
        formatogrilla(1, 4) = "CLINTE"
        formatogrilla(1, 5) = "DESCUENTO"
        formatogrilla(1, 6) = "NETO"
        formatogrilla(1, 7) = "IVA"
        formatogrilla(1, 8) = "RETENCION"
        formatogrilla(1, 9) = "TOTAL"
        formatogrilla(1, 10) = "CAJA"
        
        Rem LARGO DE LOS DATOS
        formatogrilla(2, 1) = "30"
        formatogrilla(2, 2) = "30"
        formatogrilla(2, 3) = "30"
        formatogrilla(2, 4) = "50"
        formatogrilla(2, 5) = "30"
        formatogrilla(2, 6) = "30"
        formatogrilla(2, 7) = "30"
        formatogrilla(2, 8) = "30"
        formatogrilla(2, 9) = "30"
        formatogrilla(2, 10) = "10"
        
        Rem TIPO DE DATOS C=CEROS,S=STRING,N=NUMERICO,D=DATE,
        formatogrilla(3, 1) = "S"
        formatogrilla(3, 2) = "N"
        formatogrilla(3, 3) = "N"
        formatogrilla(3, 4) = "S"
        formatogrilla(3, 5) = "N"
        formatogrilla(3, 6) = "N"
        formatogrilla(3, 7) = "N"
        formatogrilla(3, 8) = "N"
        formatogrilla(3, 9) = "N"
        formatogrilla(3, 10) = "N"
        
        Rem FORMATO GRILLA
        formatogrilla(4, 1) = ""
        formatogrilla(4, 2) = ""
        formatogrilla(4, 3) = ""
        formatogrilla(4, 4) = ""
        formatogrilla(4, 5) = ""
        formatogrilla(4, 6) = ""
        formatogrilla(4, 7) = ""
        formatogrilla(4, 8) = ""
        formatogrilla(4, 9) = ""
        
        Rem LOCCKED
        formatogrilla(5, 1) = "FALSE"
        formatogrilla(5, 2) = "FALSE"
        formatogrilla(5, 3) = "FALSE"
        formatogrilla(5, 4) = "FALSE"
        formatogrilla(5, 5) = "FALSE"
        formatogrilla(5, 6) = "FALSE"
        formatogrilla(5, 7) = "FALSE"
        formatogrilla(5, 8) = "FALSE"
        formatogrilla(5, 9) = "FALSE"
        formatogrilla(5, 10) = "FALSE"
        
        Rem VALOR MINIMO
        formatogrilla(6, 1) = ""
        formatogrilla(6, 2) = ""
        formatogrilla(6, 3) = ""
        formatogrilla(6, 4) = ""
        formatogrilla(6, 5) = ""
        formatogrilla(6, 6) = ""
        formatogrilla(6, 7) = ""
        formatogrilla(6, 8) = ""
        formatogrilla(6, 9) = ""
        
        Rem VALOR MAXIMO
        formatogrilla(7, 1) = ""
        formatogrilla(7, 2) = ""
        formatogrilla(7, 3) = ""
        formatogrilla(7, 4) = ""
        formatogrilla(7, 5) = ""
        formatogrilla(7, 6) = ""
        formatogrilla(7, 7) = ""
        formatogrilla(7, 8) = ""
        formatogrilla(7, 9) = ""
        
        Rem ANCHO
        formatogrilla(8, 1) = "9"
        formatogrilla(8, 2) = "7"
        formatogrilla(8, 3) = "7"
        formatogrilla(8, 4) = "15"
        formatogrilla(8, 5) = "8"
        formatogrilla(8, 6) = "8"
        formatogrilla(8, 7) = "8"
        formatogrilla(8, 8) = "8"
        formatogrilla(8, 9) = "8"
        formatogrilla(8, 10) = "5"
            
        listaVentas.Cols = col
        listaVentas.Rows = row
        listaVentas.Range(0, 0, listaVentas.Rows - 1, listaVentas.Cols - 1).Borders(cellEdgeBottom) = cellNone
        listaVentas.Range(0, 0, listaVentas.Rows - 1, listaVentas.Cols - 1).Borders(cellEdgeTop) = cellNone
        listaVentas.Range(0, 0, listaVentas.Rows - 1, listaVentas.Cols - 1).Borders(cellInsideVertical) = cellNone
        listaVentas.AutoRedraw = False
        listaVentas.AllowUserResizing = False
        listaVentas.DisplayFocusRect = False
        listaVentas.ExtendLastCol = True
        listaVentas.BoldFixedCell = False
        listaVentas.DrawMode = cellOwnerDraw
        listaVentas.Appearance = Flat
        listaVentas.ScrollBarStyle = Flat
        listaVentas.FixedRowColStyle = Flat
        listaVentas.BackColorFixed = RGB(90, 158, 214)
        listaVentas.BackColorFixedSel = RGB(110, 180, 230)
        listaVentas.BackColorBkg = RGB(90, 158, 214)
        listaVentas.BackColorScrollBar = RGB(231, 235, 247)
        listaVentas.BackColor1 = RGB(231, 235, 247)
        listaVentas.BackColor2 = RGB(239, 243, 255)
        listaVentas.GridColor = RGB(148, 190, 231)
        
        listaVentas.Column(0).Width = 0
        For i = 1 To col - 1
            listaVentas.Cell(0, i).text = formatogrilla(1, i)
            listaVentas.Column(i).Width = Val(formatogrilla(8, i)) * (listaVentas.Cell(0, i).Font.Size + 1.25)
            listaVentas.Column(i).MaxLength = Val(formatogrilla(2, i))
            listaVentas.Column(i).FormatString = formatogrilla(4, i)
            listaVentas.Column(i).Locked = formatogrilla(5, i)
            If formatogrilla(3, i) = "N" Then
                listaVentas.Column(i).Alignment = cellRightCenter
            Else
                listaVentas.Column(i).Alignment = cellLeftCenter
            End If
        Next i
        listaVentas.Range(0, 1, 0, listaVentas.Cols - 1).Alignment = cellCenterCenter
        listaVentas.Range(0, 1, 0, listaVentas.Cols - 1).Borders(cellEdgeBottom) = cellThin
        listaVentas.AutoRedraw = True
        listaVentas.Refresh
    End Sub
'****************************************************************************
'Formato de la Grilla Listado de Ventas
'****************************************************************************

'****************************************************************************
'Formato de la Grilla Listado de Protestos
'****************************************************************************
    Private Sub CargaGrillaProtestos(ByVal row As Integer, ByVal col As Integer)
        Dim i As Integer
        Rem DATOS DE LA COLUMNA
        formatogrilla(1, 1) = "NUMERO CHEQUE"
        formatogrilla(1, 2) = "VENCIMIENTO"
        formatogrilla(1, 3) = "FECHA PROTESTO"
        formatogrilla(1, 4) = "MONTO"
        formatogrilla(1, 5) = "FECHA CANCELACION"
        
        Rem LARGO DE LOS DATOS
        formatogrilla(2, 1) = "7"
        formatogrilla(2, 2) = "10"
        formatogrilla(2, 3) = "10"
        formatogrilla(2, 4) = "9"
        formatogrilla(2, 5) = "10"
        
        Rem TIPO DE DATOS C=CEROS,S=STRING,N=NUMERICO,D=DATE,
        formatogrilla(3, 1) = "N"
        formatogrilla(3, 2) = "N"
        formatogrilla(3, 3) = "N"
        formatogrilla(3, 4) = "N"
        formatogrilla(3, 5) = "N"
        
        Rem FORMATO GRILLA
        formatogrilla(4, 1) = ""
        formatogrilla(4, 2) = ""
        formatogrilla(4, 3) = ""
        formatogrilla(4, 4) = ""
        formatogrilla(4, 5) = ""
        
        Rem LOCCKED
        formatogrilla(5, 1) = "FALSE"
        formatogrilla(5, 2) = "FALSE"
        formatogrilla(5, 3) = "FALSE"
        formatogrilla(5, 4) = "FALSE"
        formatogrilla(5, 5) = "FALSE"
        
        Rem VALOR MINIMO
        formatogrilla(6, 1) = ""
        formatogrilla(6, 2) = ""
        formatogrilla(6, 3) = ""
        formatogrilla(6, 4) = ""
        formatogrilla(6, 5) = ""
        
        Rem VALOR MAXIMO
        formatogrilla(7, 1) = ""
        formatogrilla(7, 2) = ""
        formatogrilla(7, 3) = ""
        formatogrilla(7, 4) = ""
        formatogrilla(7, 5) = ""
        
        Rem ANCHO
        formatogrilla(8, 1) = "15"
        formatogrilla(8, 2) = "15"
        formatogrilla(8, 3) = "15"
        formatogrilla(8, 4) = "18"
        formatogrilla(8, 5) = "15"
            
        listaVentas.Cols = col
        listaVentas.Rows = row
        listaVentas.Range(0, 0, listaVentas.Rows - 1, listaVentas.Cols - 1).Borders(cellEdgeBottom) = cellNone
        listaVentas.Range(0, 0, listaVentas.Rows - 1, listaVentas.Cols - 1).Borders(cellEdgeTop) = cellNone
        listaVentas.Range(0, 0, listaVentas.Rows - 1, listaVentas.Cols - 1).Borders(cellInsideVertical) = cellNone
        listaVentas.AutoRedraw = False
        listaVentas.AllowUserResizing = False
        listaVentas.DisplayFocusRect = False
        listaVentas.ExtendLastCol = True
        listaVentas.BoldFixedCell = False
        listaVentas.DrawMode = cellOwnerDraw
        listaVentas.Appearance = Flat
        listaVentas.ScrollBarStyle = Flat
        listaVentas.FixedRowColStyle = Flat
        listaVentas.BackColorFixed = RGB(90, 158, 214)
        listaVentas.BackColorFixedSel = RGB(110, 180, 230)
        listaVentas.BackColorBkg = RGB(90, 158, 214)
        listaVentas.BackColorScrollBar = RGB(231, 235, 247)
        listaVentas.BackColor1 = RGB(231, 235, 247)
        listaVentas.BackColor2 = RGB(239, 243, 255)
        listaVentas.GridColor = RGB(148, 190, 231)
        
        listaVentas.Column(0).Width = 0
        
        For i = 1 To col - 1
            listaVentas.Cell(0, i).text = formatogrilla(1, i)
            listaVentas.Column(i).Width = Val(formatogrilla(8, i)) * (listaVentas.Cell(0, i).Font.Size + 1.25)
            listaVentas.Column(i).MaxLength = Val(formatogrilla(2, i))
            listaVentas.Column(i).FormatString = formatogrilla(4, i)
            listaVentas.Column(i).Locked = formatogrilla(5, i)
            If formatogrilla(3, i) = "N" Then
                listaVentas.Column(i).Alignment = cellRightCenter
            Else
                listaVentas.Column(i).Alignment = cellLeftCenter
            End If
        Next i
        listaVentas.Range(0, 1, 0, listaVentas.Cols - 1).Alignment = cellCenterCenter
        listaVentas.Range(0, 1, 0, listaVentas.Cols - 1).Borders(cellEdgeBottom) = cellThin
        listaVentas.AutoRedraw = True
        listaVentas.Refresh
    End Sub
'****************************************************************************
'Formato de la Grilla Listado de Protestos
'****************************************************************************

'****************************************************************************
'Formato de la Grilla Listado de Prorrogas
'****************************************************************************
    Private Sub CargaGrillaProrrogas(ByVal row As Integer, ByVal col As Integer)
        Dim i As Integer
        Rem DATOS DE LA COLUMNA
        formatogrilla(1, 1) = "FOLIO"
        formatogrilla(1, 2) = "NUMERO CHEQUE"
        formatogrilla(1, 3) = "VENCIMIENTO"
        formatogrilla(1, 4) = "MONTO"
        formatogrilla(1, 5) = "FECHA PRORROGA"
        
        Rem LARGO DE LOS DATOS
        formatogrilla(2, 1) = "10"
        formatogrilla(2, 2) = "7"
        formatogrilla(2, 3) = "10"
        formatogrilla(2, 4) = "9"
        formatogrilla(2, 5) = "10"
        
        Rem TIPO DE DATOS C=CEROS,S=STRING,N=NUMERICO,D=DATE,
        formatogrilla(3, 1) = "N"
        formatogrilla(3, 2) = "N"
        formatogrilla(3, 3) = "N"
        formatogrilla(3, 4) = "N"
        formatogrilla(3, 5) = "N"
        
        Rem FORMATO GRILLA
        formatogrilla(4, 1) = ""
        formatogrilla(4, 2) = ""
        formatogrilla(4, 3) = ""
        formatogrilla(4, 4) = ""
        formatogrilla(4, 5) = ""
        
        Rem LOCCKED
        formatogrilla(5, 1) = "FALSE"
        formatogrilla(5, 2) = "FALSE"
        formatogrilla(5, 3) = "FALSE"
        formatogrilla(5, 4) = "FALSE"
        formatogrilla(5, 5) = "FALSE"
        
        Rem VALOR MINIMO
        formatogrilla(6, 1) = ""
        formatogrilla(6, 2) = ""
        formatogrilla(6, 3) = ""
        formatogrilla(6, 4) = ""
        formatogrilla(6, 5) = ""
        
        Rem VALOR MAXIMO
        formatogrilla(7, 1) = ""
        formatogrilla(7, 2) = ""
        formatogrilla(7, 3) = ""
        formatogrilla(7, 4) = ""
        formatogrilla(7, 5) = ""
        
        Rem ANCHO
        formatogrilla(8, 1) = "15"
        formatogrilla(8, 2) = "15"
        formatogrilla(8, 3) = "15"
        formatogrilla(8, 4) = "18"
        formatogrilla(8, 5) = "15"
            
        listaVentas.Cols = col
        listaVentas.Rows = row
        listaVentas.Range(0, 0, listaVentas.Rows - 1, listaVentas.Cols - 1).Borders(cellEdgeBottom) = cellNone
        listaVentas.Range(0, 0, listaVentas.Rows - 1, listaVentas.Cols - 1).Borders(cellEdgeTop) = cellNone
        listaVentas.Range(0, 0, listaVentas.Rows - 1, listaVentas.Cols - 1).Borders(cellInsideVertical) = cellNone
        listaVentas.AutoRedraw = False
        listaVentas.AllowUserResizing = False
        listaVentas.DisplayFocusRect = False
        listaVentas.ExtendLastCol = True
        listaVentas.BoldFixedCell = False
        listaVentas.DrawMode = cellOwnerDraw
        listaVentas.Appearance = Flat
        listaVentas.ScrollBarStyle = Flat
        listaVentas.FixedRowColStyle = Flat
        listaVentas.BackColorFixed = RGB(90, 158, 214)
        listaVentas.BackColorFixedSel = RGB(110, 180, 230)
        listaVentas.BackColorBkg = RGB(90, 158, 214)
        listaVentas.BackColorScrollBar = RGB(231, 235, 247)
        listaVentas.BackColor1 = RGB(231, 235, 247)
        listaVentas.BackColor2 = RGB(239, 243, 255)
        listaVentas.GridColor = RGB(148, 190, 231)
        
        listaVentas.Column(0).Width = 0
        
        For i = 1 To col - 1
            listaVentas.Cell(0, i).text = formatogrilla(1, i)
            listaVentas.Column(i).Width = Val(formatogrilla(8, i)) * (listaVentas.Cell(0, i).Font.Size + 1.25)
            listaVentas.Column(i).MaxLength = Val(formatogrilla(2, i))
            listaVentas.Column(i).FormatString = formatogrilla(4, i)
            listaVentas.Column(i).Locked = formatogrilla(5, i)
            If formatogrilla(3, i) = "N" Then
                listaVentas.Column(i).Alignment = cellRightCenter
            Else
                listaVentas.Column(i).Alignment = cellLeftCenter
            End If
        Next i
        listaVentas.Range(0, 1, 0, listaVentas.Cols - 1).Alignment = cellCenterCenter
        listaVentas.Range(0, 1, 0, listaVentas.Cols - 1).Borders(cellEdgeBottom) = cellThin
        listaVentas.AutoRedraw = True
        listaVentas.Refresh
    End Sub
'****************************************************************************
'Formato de la Grilla Listado de Prorrogas
'****************************************************************************

'****************************************************************************
'Formato de la Grilla Listado de Cheques en Cartera
'****************************************************************************
    Private Sub CargaGrillaCheques(ByVal row As Integer, ByVal col As Integer)
        Dim i As Integer
        Rem DATOS DE LA COLUMNA
        formatogrilla(1, 1) = "NUMERO CHEQUE"
        formatogrilla(1, 2) = "BANCO"
        formatogrilla(1, 3) = "RECEPCION"
        formatogrilla(1, 4) = "MONTO"
        formatogrilla(1, 5) = "VENCIMIENTO"
        formatogrilla(1, 6) = "DOC.CANCELADO"
        
        Rem LARGO DE LOS DATOS
        formatogrilla(2, 1) = "7"
        formatogrilla(2, 2) = "30"
        formatogrilla(2, 3) = "10"
        formatogrilla(2, 4) = "9"
        formatogrilla(2, 5) = "10"
        formatogrilla(2, 6) = "13"
        
        Rem TIPO DE DATOS C=CEROS,S=STRING,N=NUMERICO,D=DATE,
        formatogrilla(3, 1) = "N"
        formatogrilla(3, 2) = "S"
        formatogrilla(3, 3) = "N"
        formatogrilla(3, 4) = "N"
        formatogrilla(3, 5) = "N"
        formatogrilla(3, 6) = "S"
        
        Rem FORMATO GRILLA
        formatogrilla(4, 1) = ""
        formatogrilla(4, 2) = ""
        formatogrilla(4, 3) = ""
        formatogrilla(4, 4) = ""
        formatogrilla(4, 5) = ""
        formatogrilla(4, 6) = ""
        
        Rem LOCCKED
        formatogrilla(5, 1) = "FALSE"
        formatogrilla(5, 2) = "FALSE"
        formatogrilla(5, 3) = "FALSE"
        formatogrilla(5, 4) = "FALSE"
        formatogrilla(5, 5) = "FALSE"
        formatogrilla(5, 6) = "FALSE"
        
        Rem VALOR MINIMO
        formatogrilla(6, 1) = ""
        formatogrilla(6, 2) = ""
        formatogrilla(6, 3) = ""
        formatogrilla(6, 4) = ""
        formatogrilla(6, 5) = ""
        formatogrilla(6, 6) = ""
        
        Rem VALOR MAXIMO
        formatogrilla(7, 1) = ""
        formatogrilla(7, 2) = ""
        formatogrilla(7, 3) = ""
        formatogrilla(7, 4) = ""
        formatogrilla(7, 5) = ""
        formatogrilla(7, 6) = ""
        
        Rem ANCHO
        formatogrilla(8, 1) = "10"
        formatogrilla(8, 2) = "25"
        formatogrilla(8, 3) = "10"
        formatogrilla(8, 4) = "10"
        formatogrilla(8, 5) = "10"
        formatogrilla(8, 6) = "10"
            
        listaVentas.Cols = col
        listaVentas.Rows = row
        listaVentas.Range(0, 0, listaVentas.Rows - 1, listaVentas.Cols - 1).Borders(cellEdgeBottom) = cellNone
        listaVentas.Range(0, 0, listaVentas.Rows - 1, listaVentas.Cols - 1).Borders(cellEdgeTop) = cellNone
        listaVentas.Range(0, 0, listaVentas.Rows - 1, listaVentas.Cols - 1).Borders(cellInsideVertical) = cellNone
        listaVentas.AutoRedraw = False
        listaVentas.AllowUserResizing = False
        listaVentas.DisplayFocusRect = False
        listaVentas.ExtendLastCol = True
        listaVentas.BoldFixedCell = False
        listaVentas.DrawMode = cellOwnerDraw
        listaVentas.Appearance = Flat
        listaVentas.ScrollBarStyle = Flat
        listaVentas.FixedRowColStyle = Flat
        listaVentas.BackColorFixed = RGB(90, 158, 214)
        listaVentas.BackColorFixedSel = RGB(110, 180, 230)
        listaVentas.BackColorBkg = RGB(90, 158, 214)
        listaVentas.BackColorScrollBar = RGB(231, 235, 247)
        listaVentas.BackColor1 = RGB(231, 235, 247)
        listaVentas.BackColor2 = RGB(239, 243, 255)
        listaVentas.GridColor = RGB(148, 190, 231)
        
        listaVentas.Column(0).Width = 0
        
        For i = 1 To col - 1
            listaVentas.Cell(0, i).text = formatogrilla(1, i)
            listaVentas.Column(i).Width = Val(formatogrilla(8, i)) * (listaVentas.Cell(0, i).Font.Size + 1.25)
            listaVentas.Column(i).MaxLength = Val(formatogrilla(2, i))
            listaVentas.Column(i).FormatString = formatogrilla(4, i)
            listaVentas.Column(i).Locked = formatogrilla(5, i)
            If formatogrilla(3, i) = "N" Then
                listaVentas.Column(i).Alignment = cellRightCenter
            Else
                listaVentas.Column(i).Alignment = cellLeftCenter
            End If
        Next i
        listaVentas.Range(0, 1, 0, listaVentas.Cols - 1).Alignment = cellCenterCenter
        listaVentas.Range(0, 1, 0, listaVentas.Cols - 1).Borders(cellEdgeBottom) = cellThin
        listaVentas.AutoRedraw = True
        listaVentas.Refresh
    End Sub
'****************************************************************************
'Formato de la Grilla Listado de Cheques en Cartera
'****************************************************************************

Private Sub listaVentas_DblClick()
    If datos = "ventas" Then
        If 0 < listaVentas.ActiveCell.row And listaVentas.ActiveCell.row < listaVentas.Rows Then
            TIPO = Left(listaVentas.Cell(listaVentas.ActiveCell.row, 1).text, 2)
            NUMERO = Right(listaVentas.Cell(listaVentas.ActiveCell.row, 1).text, 10)
            fecha = listaVentas.Cell(listaVentas.ActiveCell.row, 2).text
            Load DetalleDocumento
            DetalleDocumento.TIPO = TIPO
            DetalleDocumento.NUMERO = NUMERO
            DetalleDocumento.cajaaudit = listaVentas.Cell(listaVentas.ActiveCell.row, 10).text
            DetalleDocumento.fechaAudit = Format(fecha, "yyyy-mm-dd")
            DetalleDocumento.Show vbModal
        End If
    End If
End Sub









