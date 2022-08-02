VERSION 5.00
Object = "{4777436C-EB8C-4596-98A8-EBCF98683969}#1.0#0"; "Flexcell.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{ADBBAED6-B16F-40DC-80DF-B44910CBA76C}#1.0#0"; "Frame-Xp.ocx"
Begin VB.Form LClientes 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Listado de Clientes"
   ClientHeight    =   8250
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   15210
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8250
   ScaleWidth      =   15210
   Begin MSAdodcLib.Adodc data 
      Height          =   375
      Left            =   360
      Top             =   6780
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
      Height          =   7320
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   14970
      _ExtentX        =   26405
      _ExtentY        =   12912
      BackColor       =   16744576
      Caption         =   "Listado de Clientes"
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
      Begin VB.CommandButton Command1 
         Caption         =   "GENERA INFORME"
         Height          =   375
         Left            =   1350
         TabIndex        =   5
         Top             =   585
         Width           =   2175
      End
      Begin FlexCell.Grid Clientes 
         Height          =   6075
         Left            =   120
         TabIndex        =   0
         Top             =   1080
         Width           =   14775
         _ExtentX        =   26061
         _ExtentY        =   10716
         DefaultFontSize =   8.25
         Rows            =   2
         SelectionMode   =   1
         DateFormat      =   2
      End
      Begin XPFrame.FrameXp FrameXp5 
         Height          =   960
         Left            =   9855
         TabIndex        =   3
         Top             =   90
         Width           =   5010
         _ExtentX        =   8837
         _ExtentY        =   1693
         BackColor       =   12648384
         Caption         =   "TIPOS CLIENTES"
         CaptionEstilo3D =   1
         BackColor       =   12648384
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Begin VB.ComboBox Combotipos 
            Height          =   315
            Left            =   45
            TabIndex        =   4
            Text            =   "Combo1"
            Top             =   315
            Width           =   4875
         End
      End
   End
   Begin XPFrame.FrameXp frmImprimir 
      Height          =   375
      Left            =   10035
      TabIndex        =   2
      Top             =   7650
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
Attribute VB_Name = "LClientes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Private formatoGrilla(10, 10) As String

Private Sub Command1_Click()

    tabla = "SELECT CONCAT(LEFT(rut, 9), '-', RIGHT(rut, 1), '" & vbTab & "', sucursal, '" & vbTab & "', nombre, '" & vbTab & "', '  ', giro, '" & vbTab & "', '  ', direccion, '" & vbTab & "', '  ', comuna, '" & vbTab & "', fono1, '" & vbTab & "', celular, '" & vbTab & "', cupodirecto, '" & vbTab & "',descuento ) AS item "
    tabla = tabla & "FROM sv_maestroclientes "
    If Mid(Combotipos.text, 1, 2) <> "99" Then
    tabla = tabla & "WHERE tipocliente='" + Mid(Combotipos.text, 1, 2) + "' "
    End If
    
    tabla = tabla & "ORDER BY nombre ASC"
    Call ConectarControlData(data, servidor, baseVentas, usuario, password, tabla)
    Call cargaInforme(data, Clientes)
    
    Clientes.AddItem "", True
    Clientes.AddItem "CANTIDAD DE CLIENTES      " & data.Recordset.RecordCount, True
    Clientes.Range(Clientes.Rows - 1, 1, Clientes.Rows - 1, Clientes.Cols - 1).Merge
    Clientes.Range(Clientes.Rows - 1, 1, Clientes.Rows - 1, Clientes.Cols - 1).Alignment = cellCenterCenter
End Sub

Private Sub Form_Activate()
    Principal.barraEstado.Panels(1).text = UCase(Me.Caption)
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then
        Unload Me
    End If
    If KeyCode = 38 Then
        If Screen.ActiveForm.ActiveControl.Name = "Clientes" Then
            Unload Me
        End If
    End If
End Sub

Private Sub Form_Load()
    Dim tabla As String
    Call Centrar(Me)
    Call CargaGrillaClientes(2, 11)
    LEErTIPOSCLIENTES
    
End Sub
Sub LEErTIPOSCLIENTES()
    Dim resultados As rdoResultset
    Dim cSql As New rdoQuery
    Dim linea As Double
    
        Set cSql.ActiveConnection = ventas
        cSql.sql = "SELECT codigo,nombre "
        cSql.sql = cSql.sql + "FROM sv_tiposdeclientes "
        cSql.sql = cSql.sql + "ORDER BY codigo "
        cSql.Execute
        linea = 1
        If cSql.RowsAffected > 0 Then
            Set resultados = cSql.OpenResultset
            While Not resultados.EOF
                linea = linea + 1
                Combotipos.AddItem (resultados(0) + " " + resultados(1))
                resultados.MoveNext
            Wend
            resultados.Close
            Set resultados = Nothing
      Combotipos.AddItem ("99 TODOS")
      
      
                
        Combotipos.text = Combotipos.List(linea - 1)
        End If
        
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Principal.barraEstado.Panels(1).text = UCase(Principal.Caption)
    Call limpiaBarra(2)
End Sub


'****************************************************************************
'Formato de la Grilla Documentos
'****************************************************************************
    Private Sub CargaGrillaClientes(ByVal row As Integer, ByVal col As Integer)
        Dim i As Integer
        Rem DATOS DE LA COLUMNA
        formatoGrilla(1, 1) = "RUT"
        formatoGrilla(1, 2) = "SUC"
        formatoGrilla(1, 3) = "NOMBRE"
        formatoGrilla(1, 4) = "GIRO"
        formatoGrilla(1, 5) = "DIRECCION"
        formatoGrilla(1, 6) = "COMUNA/CIUDAD"
        formatoGrilla(1, 7) = "FONO"
        formatoGrilla(1, 8) = "CELULAR"
        formatoGrilla(1, 9) = "CUPO"
        formatoGrilla(1, 10) = "DESCUENTO"
        
        Rem LARGO DE LOS DATOS
        formatoGrilla(2, 1) = "10"
        formatoGrilla(2, 2) = "1"
        formatoGrilla(2, 3) = "50"
        formatoGrilla(2, 4) = "30"
        formatoGrilla(2, 5) = "30"
        formatoGrilla(2, 6) = "30"
        formatoGrilla(2, 7) = "10"
        formatoGrilla(2, 8) = "10"
        formatoGrilla(2, 9) = "10"
        formatoGrilla(2, 10) = "10"
        
        Rem TIPO DE DATOS C=CEROS,S=STRING,N=NUMERICO,D=DATE,
        formatoGrilla(3, 1) = "N"
        formatoGrilla(3, 2) = "C"
        formatoGrilla(3, 3) = "S"
        formatoGrilla(3, 4) = "S"
        formatoGrilla(3, 5) = "S"
        formatoGrilla(3, 6) = "S"
        formatoGrilla(3, 7) = "N"
        formatoGrilla(3, 8) = "N"
        formatoGrilla(3, 9) = "N"
        formatoGrilla(3, 10) = "N"
        
        Rem FORMATO GRILLA
        formatoGrilla(4, 1) = ""
        formatoGrilla(4, 2) = "0"
        formatoGrilla(4, 3) = ""
        formatoGrilla(4, 4) = ""
        formatoGrilla(4, 5) = ""
        formatoGrilla(4, 6) = ""
        formatoGrilla(4, 7) = ""
        formatoGrilla(4, 8) = ""
        formatoGrilla(4, 9) = "$ ###,###,##0"
        formatoGrilla(4, 10) = "% ##0.0"
    
        
        Rem LOCCKED
        formatoGrilla(5, 1) = "FALSE"
        formatoGrilla(5, 2) = "FALSE"
        formatoGrilla(5, 3) = "FALSE"
        formatoGrilla(5, 4) = "FALSE"
        formatoGrilla(5, 5) = "FALSE"
        formatoGrilla(5, 6) = "FALSE"
        formatoGrilla(5, 7) = "FALSE"
        formatoGrilla(5, 8) = "FALSE"
        formatoGrilla(5, 9) = "FALSE"
        formatoGrilla(5, 10) = "FALSE"
        
        Rem VALOR MINIMO
        formatoGrilla(6, 1) = ""
        formatoGrilla(6, 2) = ""
        formatoGrilla(6, 3) = ""
        formatoGrilla(6, 4) = ""
        formatoGrilla(6, 5) = ""
        formatoGrilla(6, 6) = ""
        formatoGrilla(6, 7) = ""
        formatoGrilla(6, 8) = ""
        formatoGrilla(6, 9) = ""
        formatoGrilla(6, 10) = ""
        
        Rem VALOR MAXIMO
        formatoGrilla(7, 1) = ""
        formatoGrilla(7, 2) = ""
        formatoGrilla(7, 3) = ""
        formatoGrilla(7, 4) = ""
        formatoGrilla(7, 5) = ""
        formatoGrilla(7, 6) = ""
        formatoGrilla(7, 7) = ""
        formatoGrilla(7, 8) = ""
        formatoGrilla(7, 9) = ""
        formatoGrilla(7, 10) = ""
        
        Rem ANCHO
        formatoGrilla(8, 1) = "8"
        formatoGrilla(8, 2) = "3"
        formatoGrilla(8, 3) = "15"
        formatoGrilla(8, 4) = "10"
        formatoGrilla(8, 5) = "20"
        formatoGrilla(8, 6) = "10"
        formatoGrilla(8, 7) = "7"
        formatoGrilla(8, 8) = "7"
        formatoGrilla(8, 9) = "9"
        formatoGrilla(8, 10) = "9"
            
        Clientes.Cols = col
        Clientes.Rows = row
        Clientes.AllowUserResizing = False
        Clientes.DisplayFocusRect = False
        Clientes.ExtendLastCol = True
        Clientes.BoldFixedCell = False
        Clientes.DrawMode = cellOwnerDraw
        Clientes.Appearance = Flat
        Clientes.ScrollBarStyle = Flat
        Clientes.FixedRowColStyle = Flat
        Clientes.BackColorFixed = RGB(90, 158, 214)
        Clientes.BackColorFixedSel = RGB(110, 180, 230)
        Clientes.BackColorBkg = RGB(90, 158, 214)
        Clientes.BackColorScrollBar = RGB(231, 235, 247)
        Clientes.BackColor1 = RGB(231, 235, 247)
        Clientes.BackColor2 = RGB(239, 243, 255)
        Clientes.GridColor = RGB(148, 190, 231)
        
        Clientes.Column(0).Width = 0
        For i = 1 To col - 1
            Clientes.Cell(0, i).text = formatoGrilla(1, i)
            Clientes.Column(i).Width = Val(formatoGrilla(8, i)) * (Clientes.Cell(0, i).Font.Size + 1.25)
            Clientes.Column(i).MaxLength = Val(formatoGrilla(2, i))
            Clientes.Column(i).FormatString = formatoGrilla(4, i)
            Clientes.Column(i).Locked = formatoGrilla(5, i)
            If formatoGrilla(3, i) = "N" Then
                Clientes.Column(i).Alignment = cellRightCenter
            End If
            If formatoGrilla(3, i) = "S" Then
                Clientes.Column(i).Alignment = cellLeftCenter
            End If
            If formatoGrilla(3, i) = "C" Then
                Clientes.Column(i).Alignment = cellCenterCenter
            End If
        Next i
        Clientes.Range(0, 0, 0, Clientes.Cols - 1).Alignment = cellCenterCenter
        Clientes.Enabled = True
    End Sub
'****************************************************************************
'Formato de la Grilla Documentos
'****************************************************************************

    Private Sub frmImprimir_BarMouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
        frmImprimir.ColorBarraAbajo = &HFFC0C0
        frmImprimir.ColorBarraArriba = &H800000
        frmImprimir.CaptionEstilo3D = Raised
    End Sub

    Private Sub frmImprimir_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
        frmImprimir.ColorBarraAbajo = &H800000
        frmImprimir.ColorBarraArriba = &HFFC0C0
        frmImprimir.CaptionEstilo3D = Inserted
        Call imprimir
    End Sub

    Private Sub imprimir()
        Call cabezaInforme("", Clientes, "LISTADO DE CLIENTES", 1)
        Clientes.PageSetup.HeaderMargin = 1
        Clientes.PageSetup.TopMargin = 1
        Clientes.PageSetup.LeftMargin = 1.5
        Clientes.PageSetup.RightMargin = 1
        Clientes.PageSetup.PrintFixedRow = True
        Clientes.PageSetup.BlackAndWhite = True
        Clientes.PageSetup.Orientation = cellLandscape
        Clientes.Range(0, 0, 0, Clientes.Cols - 1).Borders(cellEdgeBottom) = cellThin
        
        Call verificaImpresora(5, Clientes)
        
    End Sub





