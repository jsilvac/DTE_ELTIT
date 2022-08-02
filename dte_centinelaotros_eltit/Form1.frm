VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   7800
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10305
   LinkTopic       =   "Form1"
   ScaleHeight     =   7800
   ScaleWidth      =   10305
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Procesar para precios"
      Height          =   495
      Left            =   7320
      TabIndex        =   7
      Top             =   7200
      Width           =   2415
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Procesar"
      Height          =   495
      Left            =   3840
      TabIndex        =   6
      Top             =   7200
      Width           =   2415
   End
   Begin VB.ComboBox cmbAño 
      Height          =   315
      Left            =   7200
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   240
      Width           =   1935
   End
   Begin MSAdodcLib.Adodc productos 
      Height          =   375
      Left            =   1680
      Top             =   7200
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
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
   Begin MSAdodcLib.Adodc data 
      Height          =   375
      Left            =   120
      Top             =   7200
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
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
   Begin VB.ComboBox cmbLocal 
      Height          =   315
      Left            =   2400
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   240
      Width           =   1935
   End
   Begin VB.Label lblDescripcion 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4455
      Left            =   360
      TabIndex        =   5
      Top             =   2520
      Width           =   9615
   End
   Begin VB.Label lblCodigo 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   1200
      TabIndex        =   4
      Top             =   1200
      Width           =   7695
   End
   Begin VB.Label Label2 
      Caption         =   "AÑO"
      Height          =   255
      Left            =   5160
      TabIndex        =   3
      Top             =   360
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "LOCAL"
      Height          =   255
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   1695
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Private precios(5, 2) As String
    Private campos(25, 10) As String

Private Sub Command1_Click()
    Dim condicion As String
    Dim op As Integer
    Dim codigo As String
    Dim i As Long
    Call cargaProductos
    Set sql = New CSQLUtil
    If productos.Recordset.RecordCount > 0 Then
        productos.Recordset.MoveFirst
        i = 1
        While Not productos.Recordset.EOF
            codigo = Format(i, "#############")
            codigo = String(13 - Len(codigo), "0") & codigo
            
            lblCodigo.Caption = productos.Recordset.Fields(0)
            lblDescripcion.Caption = productos.Recordset.Fields(1)
            lblCodigo.Refresh
            lblDescripcion.Refresh
            
            'ESTADISTICA
            campos(0, 0) = "local"
            campos(1, 0) = "codigo"
            campos(2, 0) = "año"
            campos(3, 0) = ""
            
            campos(0, 1) = cmbLocal.List(cmbLocal.ListIndex)
            campos(1, 1) = codigo
            campos(2, 1) = cmbAño.List(cmbAño.ListIndex)
            campos(3, 1) = ""
            
            campos(0, 2) = "r_maestroproductos_estadistica_" & rubro & "_copy"
            
            condicion = ""
            op = 2
            sql.datos = campos
            Set sql.conexion = gestionRubro
            Call sql.SQLUTIL(op, condicion)
            
            'FIJO
           ' campos(0, 0) = "codigobarra"
           ' campos(1, 0) = "descripcion"
           ' campos(2, 0) = "codigoseccion"
           ' campos(3, 0) = "codigodepto"
           ' campos(4, 0) = "codigolinea"
           ' campos(5, 0) = "codigoimpuesto"
           ' campos(6, 0) = "codigomarca"
           ' campos(7, 0) = "codigotemporada"
           ' campos(8, 0) = "proveedor"
           ' campos(9, 0) = "unidadmedida"
           ' campos(10, 0) = "contenido"
           ' campos(11, 0) = "tipoembalaje"
           ' campos(12, 0) = "cantidadporembalaje"
           ' campos(13, 0) = "pesable"
           ' campos(14, 0) = "vendesinstock"
           ' campos(15, 0) = "glosaflejes"
           ' campos(16, 0) = "imprimefleje"
           ' campos(17, 0) = "glosaregistradoras"
           ' campos(18, 0) = "codigoproveedor"
           ' campos(19, 0) = "dun14"
           ' campos(20, 0) = "pcosto"
           ' campos(21, 0) = "margen"
           ' campos(22, 0) = "reponedor"
           ' campos(23, 0) = ""
           '
           ' campos(0, 1) = codigo
           ' campos(1, 1) = productos.Recordset.Fields(1)
           ' campos(2, 1) = productos.Recordset.Fields(2)
           ' campos(3, 1) = productos.Recordset.Fields(3)
           ' campos(4, 1) = productos.Recordset.Fields(4)
           ' campos(5, 1) = productos.Recordset.Fields(5)
           ' campos(6, 1) = productos.Recordset.Fields(6)
           ' campos(7, 1) = productos.Recordset.Fields(7)
           ' If IsNumeric(productos.Recordset.Fields(8)) = True Then
           '     campos(8, 1) = productos.Recordset.Fields(8)
           '     campos(9, 1) = productos.Recordset.Fields(9)
           '     campos(10, 1) = productos.Recordset.Fields(10)
           '     campos(11, 1) = productos.Recordset.Fields(11)
           '     campos(12, 1) = productos.Recordset.Fields(12)
           '     campos(13, 1) = productos.Recordset.Fields(13)
           '     campos(14, 1) = productos.Recordset.Fields(14)
           '     campos(15, 1) = productos.Recordset.Fields(15)
           '     campos(16, 1) = productos.Recordset.Fields(16)
           '     campos(17, 1) = productos.Recordset.Fields(17)
           '     campos(18, 1) = productos.Recordset.Fields(18)
           '     campos(19, 1) = productos.Recordset.Fields(19)
           ' Else
           '     campos(8, 1) = ""
           '     campos(9, 1) = productos.Recordset.Fields(8)
           '     campos(10, 1) = productos.Recordset.Fields(9)
           '     campos(11, 1) = productos.Recordset.Fields(10)
           '     campos(12, 1) = productos.Recordset.Fields(11)
           '     campos(13, 1) = productos.Recordset.Fields(12)
           '     campos(14, 1) = productos.Recordset.Fields(13)
           '     campos(15, 1) = productos.Recordset.Fields(14)
           '     campos(16, 1) = productos.Recordset.Fields(15)
           '     campos(17, 1) = productos.Recordset.Fields(16)
           '     campos(18, 1) = productos.Recordset.Fields(17)
           '     campos(19, 1) = productos.Recordset.Fields(18)
           ' End If
           ' campos(20, 1) = productos.Recordset.Fields(20)
           ' campos(21, 1) = productos.Recordset.Fields(21)
           ' campos(22, 1) = productos.Recordset.Fields(22)
           ' campos(23, 1) = ""
          '
          '  campos(0, 2) = "r_maestroproductos_fijo_" & rubro & "_copy"
          '
            'condicion = ""
            'op = 2
            'sql.datos = campos
            'Set sql.conexion = gestionRubro
            'Call sql.SQLUTIL(op, condicion)
            
            'STOCK
            campos(0, 0) = "local"
            campos(1, 0) = "codigo"
            campos(2, 0) = "año"
            campos(3, 0) = "bodega"
            campos(4, 0) = ""
            
            campos(0, 1) = cmbLocal.List(cmbLocal.ListIndex)
            campos(1, 1) = codigo
            campos(2, 1) = cmbAño.List(cmbAño.ListIndex)
            campos(3, 1) = "00"
            campos(4, 1) = ""
            
            campos(0, 2) = "r_maestroproductos_stock_" & rubro & "_copy"
            
            condicion = ""
            op = 2
            sql.datos = campos
            Set sql.conexion = gestionRubro
            Call sql.SQLUTIL(op, condicion)
            i = i + 1
            productos.Recordset.MoveNext
        Wend
    End If
End Sub

Private Sub Command2_Click()
    Dim tabla As String
    Dim condicion As String
    Dim op As Integer
    Dim codigon As String
    Dim codigov As String
    Set sql = New CSQLUtil
    tabla = "SELECT DISTINCT g01.codigobarra AS codnuevo, g01.descripcion, g00.codigobarra AS codviejo FROM gestion01.r_maestroproductos_fijo_00 AS g01 INNER JOIN r_maestroproductos_fijo_00 AS g00 ON g01.descripcion = g00.descripcion ORDER BY g01.codigobarra ASC"
    Call ConectarControlData(data, servidor, basedatos & rubro, usuario, password, tabla)
    If data.Recordset.RecordCount > 0 Then
        data.Recordset.MoveFirst
        While Not data.Recordset.EOF
            lblCodigo.Caption = data.Recordset.Fields("codnuevo")
            lblDescripcion.Caption = data.Recordset.Fields("descripcion")
            lblCodigo.Refresh
            lblDescripcion.Refresh
            codigon = data.Recordset.Fields("codnuevo")
            codigov = data.Recordset.Fields("codviejo")
            tabla = "SELECT local, codigoprecio, precioautomatico,preciosistema, preciopuntoventa, preciocosto, margen, precioactivo FROM r_maestroproductos_precios_" & rubro & " WHERE codigo = '" & codigov & "'"
            Call ConectarControlData(productos, servidor, basedatos & rubro, usuario, password, tabla)
            If productos.Recordset.RecordCount > 0 Then
                productos.Recordset.MoveFirst
                While Not productos.Recordset.EOF
                    campos(0, 0) = "local"
                    campos(1, 0) = "codigo"
                    campos(2, 0) = "codigoprecio"
                    campos(3, 0) = "precioautomatico"
                    campos(4, 0) = "preciosistema"
                    campos(5, 0) = "preciopuntoventa"
                    campos(6, 0) = "preciocosto"
                    campos(7, 0) = "margen"
                    campos(8, 0) = "precioactivo"
                    campos(9, 0) = ""
                    
                    campos(0, 1) = productos.Recordset.Fields("local")
                    campos(1, 1) = codigon
                    campos(2, 1) = productos.Recordset.Fields("codigoprecio")
                    campos(3, 1) = productos.Recordset.Fields("precioautomatico")
                    campos(4, 1) = productos.Recordset.Fields("preciosistema")
                    campos(5, 1) = productos.Recordset.Fields("preciopuntoventa")
                    campos(6, 1) = productos.Recordset.Fields("preciocosto")
                    campos(7, 1) = productos.Recordset.Fields("margen")
                    campos(8, 1) = productos.Recordset.Fields("precioactivo")
                    campos(9, 1) = ""
                    
                    campos(0, 2) = "r_maestroproductos_precios_" & rubro & "_copy"
                    
                    condicion = "codigo = '" & codigov & "'"
                    op = 2
                    sql.datos = campos
                    Set sql.conexion = gestionRubro
                    Call sql.SQLUTIL(op, condicion)
                    
                    productos.Recordset.MoveNext
                Wend
            End If
            data.Recordset.MoveNext
        Wend
    End If
End Sub

Private Sub Form_Load()
    Call cargaLocales
    Call cargaAños
    Call cargaPrecios
End Sub

Private Sub cargaLocales()
    Dim tabla As String
    tabla = "select codigo from g_maestroempresas order by codigo asc"
    Call ConectarControlData(data, servidor, basedatos, usuario, password, tabla)
    If data.Recordset.RecordCount > 0 Then
        data.Recordset.MoveFirst
        While Not data.Recordset.EOF
            cmbLocal.AddItem data.Recordset.Fields("codigo")
            data.Recordset.MoveNext
        Wend
    End If
End Sub

Private Sub cargaAños()
    Dim i As Integer
    Dim cad As String
    For i = 2000 To Val(Format(Now, "yyyy"))
        cad = "01-01-"
        cad = cad & i
        cmbAño.AddItem Format(cad, "yyyy")
    Next i
End Sub

Private Sub cargaProductos()
    Dim tabla As String
    tabla = "select * from r_maestroproductos_fijo_" & rubro & " order by codigobarra asc"
    Call ConectarControlData(productos, servidor, basedatos & rubro, usuario, password, tabla)
End Sub

Private Sub cargaPrecios()
    Dim tabla As String
    Dim i As Integer
    tabla = "select codigo, porcentajedelmargen AS margen from g_maestrodetiposdeprecios order by codigo asc"
    Call ConectarControlData(data, servidor, basedatos, usuario, password, tabla)
    If data.Recordset.RecordCount > 0 Then
        data.Recordset.MoveFirst
        i = 0
        While Not data.Recordset.EOF
            precios(i, 0) = data.Recordset.Fields("codigo")
            precios(i, 1) = data.Recordset.Fields("margen")
            data.Recordset.MoveNext
        Wend
    End If
End Sub

