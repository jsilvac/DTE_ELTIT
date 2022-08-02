VERSION 5.00
Object = "{4777436C-EB8C-4596-98A8-EBCF98683969}#1.0#0"; "FlexCell.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{ADBBAED6-B16F-40DC-80DF-B44910CBA76C}#1.0#0"; "Frame-Xp.ocx"
Begin VB.Form Especiales 
   BackColor       =   &H00008000&
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   6840
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10770
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   ScaleHeight     =   6840
   ScaleWidth      =   10770
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin XPFrame.FrameXp frmDatos 
      Height          =   6675
      Left            =   60
      TabIndex        =   3
      Top             =   60
      Width           =   10635
      _ExtentX        =   18759
      _ExtentY        =   11774
      BackColor       =   12648384
      Caption         =   "Precios Especiales para Clientes"
      CaptionEstilo3D =   1
      BackColor       =   12648384
      ColorBarraArriba=   12648384
      ColorBarraAbajo =   32768
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   1
      Begin VB.TextBox dato2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   8475
         MaxLength       =   9
         TabIndex        =   1
         Top             =   6240
         Width           =   1965
      End
      Begin VB.TextBox dato1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   180
         MaxLength       =   13
         TabIndex        =   0
         Top             =   6210
         Width           =   1635
      End
      Begin XPFrame.FrameXp frmLista 
         Height          =   4875
         Left            =   120
         TabIndex        =   9
         Top             =   960
         Width           =   10395
         _ExtentX        =   18336
         _ExtentY        =   8599
         BackColor       =   12648447
         Caption         =   "Lista de Productos"
         CaptionEstilo3D =   1
         BackColor       =   12648447
         ColorBarraArriba=   12648447
         ColorBarraAbajo =   32896
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
         Begin FlexCell.Grid lista 
            Height          =   4395
            Left            =   60
            TabIndex        =   2
            Top             =   420
            Width           =   10275
            _ExtentX        =   18124
            _ExtentY        =   7752
            Cols            =   5
            DefaultFontSize =   9.75
            Rows            =   1
            SelectionMode   =   1
         End
         Begin MSAdodcLib.Adodc data 
            Height          =   330
            Left            =   60
            Top             =   4560
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
      Begin XPFrame.FrameXp frmCerrar 
         Height          =   330
         Left            =   10260
         TabIndex        =   10
         Top             =   25
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   582
         BackColor       =   49344
         Caption         =   "X"
         CaptionEstilo3D =   1
         BackColor       =   49344
         ColorBarraArriba=   32896
         ColorBarraAbajo =   12648447
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
      End
      Begin VB.Label lblPrecio 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   6795
         TabIndex        =   16
         Top             =   6240
         Width           =   1650
      End
      Begin VB.Label lblProducto 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   1860
         TabIndex        =   15
         Top             =   6240
         Width           =   4905
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00008080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "P.Especial"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   8475
         TabIndex        =   14
         Top             =   5880
         Width           =   1965
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00008080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "P.Normal"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   6795
         TabIndex        =   13
         Top             =   5880
         Width           =   1650
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00008080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Descripcion"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   1860
         TabIndex        =   12
         Top             =   5880
         Width           =   4905
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00008080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Codigo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   180
         TabIndex        =   11
         Top             =   5880
         Width           =   1650
      End
      Begin VB.Label lblDV 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   3240
         TabIndex        =   8
         Top             =   540
         Width           =   435
      End
      Begin VB.Label lblNombre 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   5280
         TabIndex        =   7
         Top             =   540
         Width           =   5235
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         BackColor       =   &H00008080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Nombre"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   3780
         TabIndex        =   6
         Top             =   540
         Width           =   1335
      End
      Begin VB.Label lblRut 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   1620
         TabIndex        =   5
         Top             =   540
         Width           =   1575
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00008080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Rut"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   120
         TabIndex        =   4
         Top             =   540
         Width           =   1335
      End
   End
End
Attribute VB_Name = "Especiales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Private formatogrilla(10, 10) As String
    Public suc As String
    Private modifica As Boolean

'============================================================
'MANEJO DE LOS CONTOLES
'============================================================
    '========================================================
    'GotFocus
    '========================================================
    Private Sub dato1_GotFocus()
        Call VerificarCajas(Me, dato1)
        Call selecciona(dato1)
        Principal.barraEstado.Panels(2).text = "F2: Ayuda Productos"
    End Sub
    
    Private Sub dato2_GotFocus()
        Call VerificarCajas(Me, dato2)
        Call selecciona(dato2)
    End Sub
    '========================================================
    'GotFocus
    '========================================================
    
    '========================================================
    'KeyDown
    '========================================================
    Private Sub dato1_KeyDown(KeyCode As Integer, Shift As Integer)
        If KeyCode = vbKeyF2 Then
            Call ayudaProductotxt(dato1)
        Else
            Call Flechas(KeyCode, dato1)
        End If
    End Sub
    
    Private Sub dato2_KeyDown(KeyCode As Integer, Shift As Integer)
        Call Flechas(KeyCode, dato1)
    End Sub
    '========================================================
    'KeyDown
    '========================================================
    
    '========================================================
    'KeyPress
    '========================================================
    Private Sub dato1_KeyPress(KeyAscii As Integer)
        KeyAscii = esNumero(KeyAscii)
        If KeyAscii = 13 And dato1.text <> "" Then
            dato1.text = ceros(dato1)
            If revisaCodigo = False Then
                lblProducto.Caption = leerNombreProducto(dato1.text)
                lblPrecio.Caption = leerPrecioProducto(dato1.text, "01")
                If lblProducto.Caption <> "" Then
                    dato2.text = lblPrecio.Caption
                    SendKeys "{Tab}"
                Else
                    Call selecciona(dato1)
                End If
            Else
                If MsgBox("El Codigo ya tiene un precio especial." & vbCrLf & "¿Desea cambiarlo?", vbYesNo, "Messaje") = vbYes Then
                    Call leerEspecial
                    modifica = True
                    dato2.SetFocus
                Else
                    Call selecciona(dato1)
                End If
            End If
        End If
    End Sub
    
    Private Sub dato2_KeyPress(KeyAscii As Integer)
        KeyAscii = esNumero(KeyAscii)
        If KeyAscii = 13 And dato2.text <> "" Then
            If modifica = False Then
                Call grabarEspeciales
            Else
                Call modificaEspeciales
            End If
            dato1.text = ""
            dato2.text = ""
            lblProducto.Caption = ""
            lblPrecio.Caption = ""
            dato1.SetFocus
            leerEspeciales
        End If
    End Sub
    '========================================================
    'KeyPress
    '========================================================
    
    '========================================================
    'LostFocus
    '========================================================
    Private Sub dato1_LostFocus()
        Call limpiaBarra(2)
    End Sub
    '========================================================
    'LostFocus
    '========================================================
'============================================================
'MANEJO DE LOS CONTOLES
'============================================================

'****************************************************************************
'Formato de la Grilla Documentos
'****************************************************************************
    Private Sub CargaGrillaLista(ByVal row As Integer, ByVal col As Integer)
        Dim i As Integer
        Rem DATOS DE LA COLUMNA
        formatogrilla(1, 1) = "CODIGO"
        formatogrilla(1, 2) = "DESCRIPCION"
        formatogrilla(1, 3) = "P.NORMAL"
        formatogrilla(1, 4) = "P.ESPECIAL"
        
        Rem LARGO DE LOS DATOS
        formatogrilla(2, 1) = "13"
        formatogrilla(2, 2) = "50"
        formatogrilla(2, 3) = "9"
        formatogrilla(2, 4) = "9"
        
        Rem TIPO DE DATOS C=CEROS,S=STRING,N=NUMERICO,D=DATE,
        formatogrilla(3, 1) = "N"
        formatogrilla(3, 2) = "S"
        formatogrilla(3, 3) = "N"
        formatogrilla(3, 4) = "N"
        
        Rem FORMATO GRILLA
        formatogrilla(4, 1) = "0000000000000"
        formatogrilla(4, 2) = ""
        formatogrilla(4, 3) = "$ ###,###,##0"
        formatogrilla(4, 4) = "$ ###,###,##0"
        
        Rem LOCCKED
        formatogrilla(5, 1) = "TRUE"
        formatogrilla(5, 2) = "TRUE"
        formatogrilla(5, 3) = "TRUE"
        formatogrilla(5, 4) = "TRUE"
        
        Rem VALOR MINIMO
        formatogrilla(6, 1) = ""
        formatogrilla(6, 2) = ""
        formatogrilla(6, 3) = ""
        formatogrilla(6, 4) = ""
        
        Rem VALOR MAXIMO
        formatogrilla(7, 1) = ""
        formatogrilla(7, 2) = ""
        formatogrilla(7, 3) = ""
        formatogrilla(7, 4) = ""
        
        Rem ANCHO
        formatogrilla(8, 1) = "10"
        formatogrilla(8, 2) = "30"
        formatogrilla(8, 3) = "10"
        formatogrilla(8, 4) = "10"
            
        lista.Cols = col
        lista.Rows = row
        lista.AllowUserResizing = False
        lista.DisplayFocusRect = False
        lista.ExtendLastCol = True
        lista.BoldFixedCell = False
        lista.DrawMode = cellOwnerDraw
        lista.Appearance = Flat
        lista.ScrollBarStyle = Flat
        lista.FixedRowColStyle = Flat
        lista.BackColorFixed = RGB(90, 214, 158)
        lista.BackColorFixedSel = RGB(110, 230, 180)
        lista.BackColorBkg = RGB(90, 214, 158)
        lista.BackColorScrollBar = RGB(231, 247, 235)
        lista.BackColor1 = RGB(231, 247, 235)
        lista.BackColor2 = RGB(239, 255, 243)
        lista.GridColor = RGB(148, 231, 190)
        
        lista.Column(0).Width = 0
        For i = 1 To col - 1
            lista.Cell(0, i).text = formatogrilla(1, i)
            lista.Column(i).Width = Val(formatogrilla(8, i)) * (lista.Cell(0, i).Font.Size + 1.25)
            lista.Column(i).MaxLength = Val(formatogrilla(2, i))
            lista.Column(i).FormatString = formatogrilla(4, i)
            lista.Column(i).Locked = formatogrilla(5, i)
            If formatogrilla(3, i) = "N" Then
                lista.Column(i).Alignment = cellRightCenter
            Else
                lista.Column(i).Alignment = cellLeftCenter
            End If
        Next i
        lista.Range(0, 1, 0, lista.Cols - 1).Alignment = cellCenterCenter
        lista.Enabled = True
    End Sub
'****************************************************************************
'Formato de la Grilla Documentos
'****************************************************************************

'=============================================================================
'LEER PRECIOS ESPECIALES
'=============================================================================
    Private Sub leerEspeciales()
        Dim tabla As String
        tabla = "SELECT CONCAT(codigo, '" & vbTab & "', descripcion, '" & vbTab & "', precionormal, '" & vbTab & "', precioespecial) AS item "
        tabla = tabla & "FROM sv_maestroclientes_especiales "
        tabla = tabla & "WHERE rut = '" & lblrut.Caption & lbldv.Caption & "' AND sucursal = '" & suc & "' ORDER BY codigo ASC"
        Call ConectarControlData(data, servidor, baseVentas, usuario, password, tabla)
        lista.Rows = 1
        lista.AutoRedraw = False
        If data.Recordset.RecordCount > 0 Then
            data.Recordset.MoveFirst
            While Not data.Recordset.EOF
                lista.AddItem data.Recordset.Fields("item"), True
                data.Recordset.MoveNext
            Wend
        lista.AutoRedraw = True
        lista.Refresh
        End If
    End Sub
'=============================================================================
'LEER PRECIOS ESPECIALES
'=============================================================================

'=============================================================================
'LEER ESPECIAL
'=============================================================================
    Public Sub leerEspecial()
        
        Dim CAMPOS(10, 3) As String
        Dim op As Integer
        Set sql = New sqlventas.sqlventa
        CAMPOS(0, 0) = "codigo"
        CAMPOS(1, 0) = "descripcion"
        CAMPOS(2, 0) = "precionormal"
        CAMPOS(3, 0) = "precioespecial"
        CAMPOS(4, 0) = ""
        
        CAMPOS(0, 2) = "sv_maestroclientes_especiales"
        
        condicion = "rut = '" & lblrut.Caption & lbldv.Caption & "' AND sucursal = '" & suc & "' AND codigo = '" & dato1.text & "'"
        op = 5
        sql.response = CAMPOS
        Set sql.conexion = ventas
        Call sql.sqlventas(op, condicion)
        dato1.text = sql.response(0, 3)
        lblProducto.Caption = sql.response(1, 3)
        lblPrecio.Caption = sql.response(2, 3)
        dato2.text = sql.response(3, 3)
    End Sub
'=============================================================================
'LEER ESPECIAL
'=============================================================================

'=============================================================================
'GRABAR PRECIOS ESPECIALES
'=============================================================================
    Public Sub grabarEspeciales()
        
        Dim CAMPOS(10, 3) As String
        Dim op As Integer
        Set sql = New sqlventas.sqlventa
        CAMPOS(0, 0) = "rut"
        CAMPOS(1, 0) = "sucursal"
        CAMPOS(2, 0) = "codigo"
        CAMPOS(3, 0) = "descripcion"
        CAMPOS(4, 0) = "precionormal"
        CAMPOS(5, 0) = "precioespecial"
        CAMPOS(6, 0) = ""
        
        CAMPOS(0, 1) = lblrut.Caption & lbldv.Caption
        CAMPOS(1, 1) = suc
        CAMPOS(2, 1) = dato1.text
        CAMPOS(3, 1) = lblProducto.Caption
        CAMPOS(4, 1) = lblPrecio.Caption
        CAMPOS(5, 1) = dato2.text
        CAMPOS(6, 1) = ""
        
        CAMPOS(0, 2) = "sv_maestroclientes_especiales"
        
        condicion = ""
        op = 2
        sql.response = CAMPOS
        Set sql.conexion = ventas
        Call sql.sqlventas(op, condicion)
    End Sub
'=============================================================================
'GRABAR PRECIOS ESPECIALES
'=============================================================================

'=============================================================================
'MODIFICAR PRECIOS ESPECIALES
'=============================================================================
    Public Sub modificaEspeciales()
        
        Dim CAMPOS(10, 3) As String
        Dim op As Integer
        Set sql = New sqlventas.sqlventa
        CAMPOS(0, 0) = "rut"
        CAMPOS(1, 0) = "sucursal"
        CAMPOS(2, 0) = "codigo"
        CAMPOS(3, 0) = "descripcion"
        CAMPOS(4, 0) = "precionormal"
        CAMPOS(5, 0) = "precioespecial"
        CAMPOS(6, 0) = ""
        
        CAMPOS(0, 1) = lblrut.Caption & lbldv.Caption
        CAMPOS(1, 1) = suc
        CAMPOS(2, 1) = dato1.text
        CAMPOS(3, 1) = lblProducto.Caption
        CAMPOS(4, 1) = lblPrecio.Caption
        CAMPOS(5, 1) = dato2.text
        CAMPOS(6, 1) = ""
        
        CAMPOS(0, 2) = "sv_maestroclientes_especiales"
        
        condicion = "rut = '" & lblrut.Caption & lbldv.Caption & "' AND sucursal = '" & suc & "' AND codigo = '" & dato1.text & "'"
        op = 3
        sql.response = CAMPOS
        Set sql.conexion = ventas
        Call sql.sqlventas(op, condicion)
        modifica = False
    End Sub
'=============================================================================
'MODIFICAR PRECIOS ESPECIALES
'=============================================================================

'=============================================================================
'ELIMINAR PRECIOS ESPECIALES
'=============================================================================
    Private Sub eliminarEspeciales(ByVal CODIGO As String)
        
        Dim CAMPOS(1, 3) As String
        Dim op As Integer
        Set sql = New sqlventas.sqlventa
        condicion = "rut = '" & lblrut.Caption & lbldv.Caption & "' AND sucursal = '" & suc & "' AND codigo = '" & CODIGO & "'"
        op = 4
        CAMPOS(0, 2) = "sv_maestroclientes_especiales"
        sql.response = CAMPOS
        Set sql.conexion = ventas
        Call sql.sqlventas(op, condicion)
    End Sub
'=============================================================================
'ELIMINAR PRECIOS ESPECIALES
'=============================================================================

    Private Sub Form_Activate()
        Principal.barraEstado.Panels(1).text = UCase(Me.frmDatos.Caption)
        Call leerEspeciales
    End Sub
    
    Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
        If KeyCode = 27 Then
            Unload Me
        End If
        If KeyCode = 38 Then
            If Screen.ActiveForm.ActiveControl.Name = "dato1" Then
                Unload Me
            End If
        End If
    End Sub
    
    Private Sub Form_Load()
        Call CargaGrillaLista(1, 5)
        modifica = False
    End Sub
    
    Private Sub Form_Unload(Cancel As Integer)
        Principal.barraEstado.Panels(1).text = UCase(MClientes.Caption)
        Call limpiaBarra(2)
    End Sub

    Private Sub frmCerrar_BarClick()
        Call cambiaColor(frmCerrar)
        frmCerrar.CaptionEstilo3D = Inserted
        Unload Me
    End Sub
    
    Private Sub frmCerrar_BarMouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
        Call cambiaColor(frmCerrar)
        frmCerrar.CaptionEstilo3D = Raised
    End Sub

    Private Sub Lista_DblClick()
        dato1.text = lista.Cell(lista.ActiveCell.row, 1).text
        lblProducto.Caption = lista.Cell(lista.ActiveCell.row, 2).text
        lblPrecio.Caption = lista.Cell(lista.ActiveCell.row, 3).text
        dato2.text = lista.Cell(lista.ActiveCell.row, 4).text
        modifica = True
        dato2.SetFocus
    End Sub

    Private Sub lista_KeyDown(KeyCode As Integer, ByVal Shift As Integer)
        Select Case KeyCode
            Case 46
                If lista.ActiveCell.row > 0 Then
                    Call eliminarEspeciales(lista.Cell(lista.ActiveCell.row, 1).text)
                    lista.RemoveItem (lista.ActiveCell.row)
                End If
        End Select
    End Sub

    Private Function revisaCodigo() As Boolean
        Dim i As Long
        revisaCodigo = False
        For i = 1 To lista.Rows - 1
            If lista.Cell(i, 1).text = dato1.text Then
                revisaCodigo = True
                Exit For
            End If
        Next i
    End Function
















