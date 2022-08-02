VERSION 5.00
Object = "{4777436C-EB8C-4596-98A8-EBCF98683969}#1.0#0"; "Flexcell.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{ADBBAED6-B16F-40DC-80DF-B44910CBA76C}#1.0#0"; "Frame-Xp.ocx"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form ImpresionBoleta 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Imprimiendo boleta......."
   ClientHeight    =   1530
   ClientLeft      =   4365
   ClientTop       =   8235
   ClientWidth     =   6420
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1530
   ScaleWidth      =   6420
   ShowInTaskbar   =   0   'False
   Begin VB.ListBox List1 
      Height          =   3375
      Left            =   240
      TabIndex        =   8
      Top             =   3360
      Width           =   9495
   End
   Begin VB.Timer Timer4 
      Left            =   1890
      Top             =   2565
   End
   Begin VB.Timer Timer2 
      Left            =   960
      Top             =   2565
   End
   Begin VB.Timer Timer3 
      Left            =   1440
      Top             =   2565
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Z"
      Height          =   240
      Left            =   4995
      TabIndex        =   6
      Top             =   2880
      Width           =   600
   End
   Begin VB.CommandButton Command3 
      Caption         =   "salir"
      Height          =   330
      Left            =   7080
      TabIndex        =   5
      Top             =   1200
      Width           =   1140
   End
   Begin VB.TextBox numeroboleta 
      Height          =   330
      Left            =   6480
      MaxLength       =   10
      TabIndex        =   3
      Top             =   1800
      Width           =   2400
   End
   Begin VB.Timer Timer1 
      Interval        =   5000
      Left            =   540
      Top             =   2565
   End
   Begin MSCommLib.MSComm FISCAL 
      Left            =   0
      Top             =   2640
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin XPFrame.FrameXp FrameXp1 
      Height          =   1545
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   2725
      BackColor       =   16744576
      Caption         =   ""
      CaptionEstilo3D =   1
      BackColor       =   16744576
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
         Caption         =   "Salir"
         Height          =   330
         Left            =   4680
         TabIndex        =   4
         Top             =   960
         Width           =   1230
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Height          =   555
         Left            =   135
         TabIndex        =   9
         Top             =   855
         Width           =   4875
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "IMPRIMIENDO BOLETA.........."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   345
         Left            =   630
         TabIndex        =   1
         Top             =   540
         Width           =   4275
      End
   End
   Begin FlexCell.Grid Impresion 
      Height          =   735
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Visible         =   0   'False
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   1296
      Cols            =   5
      DefaultFontSize =   8.25
      Rows            =   30
   End
   Begin MSAdodcLib.Adodc data 
      Height          =   330
      Left            =   90
      Top             =   855
      Visible         =   0   'False
      Width           =   1275
      _ExtentX        =   2249
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
   Begin VB.Label lbl_respuesta 
      Height          =   330
      Left            =   1260
      TabIndex        =   7
      Top             =   1530
      Width           =   2985
   End
End
Attribute VB_Name = "ImpresionBoleta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

    Private comandofiscal As String
    Private glosapago As String
    Private boleta As String
    Private str_rec As String
    Private pide_numero_boleta As Boolean
    Private unidades As String
    Private preciounitario As String
    Private totaldescuento As String
    Private total As String
    Private decimales As String
    Private enteros As String
    Private puertoif As Integer
    Private puerto As Integer
    Private numero_venta As String
    Private i As Integer
    Private car As String
    Private montofinal As Double
    Private abono As String
    Private montocredito As Double
    Private cantidaddecuotas As Double
    Private montocuotas As Double
    Private clientecuotas As String
    Private primervencimiento As String
    
    Private nombreclientecuotas As String
    
    
    
Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Command2_Click()
    Label2 = "Enviando solicitud a la impresora..." & vbCrLf & "Espere por favor"
    Command1.Enabled = False
    Command2.Enabled = False
    Timer3.Interval = 10000
    Timer3.Enabled = True
    imprime_comando ("")
End Sub


Private Sub Form_Load()
    Dim x As Long
    Dim nBufferLength As Long
    Dim LpBuffer As String
    Dim linea As String
    
    
    'Call Centrar(Me)
    nBufferLength = 255
    LpBuffer = String$(255, 32)

    decimales = 0
    enteros = 0
    Timer1.Interval = 1000
    Timer1.Enabled = True
    
' Configura Numero de puerto COM y asigna parametros
    puertoif = 1
    puerto = puertoif
    FISCAL.CommPort = puerto
    'imprime_log ("Asigna puerto COM" & puerto)
    FISCAL.Settings = "19200,n,8,1"
    'imprime_log ("Setea puerto COM" & puerto)
    FISCAL.Handshaking = comRTS
   ' imprime_log ("Asigna Handshaking de puerto COM" & puerto)
    FISCAL.InputLen = 10
    FISCAL.RThreshold = 1
    FISCAL.SThreshold = 1
    If FISCAL.PortOpen = True Then
       FISCAL.PortOpen = False
    End If
    
    FISCAL.PortOpen = True
   ' imprime_log ("Abre puerto COM" & puerto)
    pide_numero_boleta = False
   numeroboleta = PVentas.dato2.text
   If sw = False Then
        Call Imprime_Boletas(numeroboleta, data)
    End If
    If imprimepago = "si" Then
    Call IMPRIMEPAGOFISCAL
    End If
    If repactacionimprimir = "si" Then
    Call IMPRIMErepactacion
    End If
    
End Sub

Private Sub Form_LostFocus()
    Me.SetFocus
End Sub
Sub InicializaCaja()
  
    Call imprime_comando("32")
    sw = False
    Call imprime_comando("60")
    Call imprime_comando("111----------------------------------------")
    Call imprime_comando("111                                        ")
    Call imprime_comando("116    CAJA INICIALIZADA CORRECTAMENTE     ")
       
    Call imprime_comando("111----------------------------------------")
       
    Call imprime_comando("61")
    Call imprime_comando("99")
    
    
End Sub
Sub imprimeX()
Call imprime_comando("01")
sw = False
End Sub
Sub imprimeZ()
  Call imprime_comando("021")
  sw = False
End Sub

Private Sub FISCAL_OnComm()
Dim NUMERO As String
Dim K As Integer
Dim final As Integer

'Permanece escuchando al puerto COM y recibe la data de respuesta a los comandos enviados
'If comandofiscal = "482" Then
'
'numero = FISCAL.Input
'For k = 1 To Len(numero)
'If Mid(numero, k, 1) = ":" Then final = k - 1: Exit For
'
'Next k
'numero = Mid(numero, 1, final)
'
'numeroboleta = String(10 - final, "0") + Mid(Str(Val(numero) + 1), 2, final)
'End If

     str_rec = FISCAL.Input
  
    
    If Len(str_rec) > 0 Then
        str_rec = Trim(str_rec)
        Timer3.Enabled = False
' Si recibe un Ascii 10, indica que la printer ha aceptado el comando y lo ha procesado correctamente

        If Left(str_rec, 1) = Chr$(10) Then
            Label2 = "Imprimiendo informe...." & vbCrLf & "Espere por favor"
            Command1.Enabled = False
            Command2.Enabled = False
            Timer2.Interval = 200
            Timer2.Enabled = True
        Else
' Si recibio un Ascii distinto a 10, implica que la printer por alguna razon no pudo procesar el error
' Se pasa el caracter para analizar la respuesta y mostrar mensaje al operador
            lbl_respuesta.Caption = str_rec
            pide_numero_boleta = True
            If pide_numero_boleta = True Then
                For i = 1 To Len(str_rec)
                  car = Mid(str_rec, i, 1)
                    If car = ":" Then
                        numero_venta = Mid(str_rec, 1, i - 1)
                        Exit For
                    End If
                  Next i
            End If
            procesa_error (Mid(str_rec, 1, 3))
        End If
    End If
End Sub
    


Private Sub Timer1_Timer()
    Timer1.Enabled = False
End Sub
Public Function procesa_error(resp As String)
    Command1.Enabled = True
    Command2.Enabled = True
' Se procesa la respuesta al comando enviado, en caso que se distinto a OK (Ascii 10)
' Y se indica al operador.
    Label2.Caption = ""
    Select Case resp
    Case "013"
        Label2 = "Impresora no responde" & vbCrLf & "Apague y encienda la impresora y reintente"
    Case "040"
        Label2 = "Impresora ocupada" & vbCrLf & "Espere unos segundos y reintente"
    Case "043"
        Label2 = "Falta Papel o la tapa de la impresora esta abierta" & vbCrLf & "Verifique y luego reintente"
    Case "059"
        Label2 = "Impresora en periodo de Venta" & vbCrLf & "Debe emitir reporte Z"
    Case "060"
        Label2 = "Impresora en periodo de cierre" & vbCrLf & "Debe emitir una boleta"
    Case Else
        Label2 = "Codigo de retorno no catalogado" & vbCrLf & "Codigo: " & resp
    End Select
End Function

Public Sub Imprime_Boletas(ByVal numero_venta As String, ByRef rollo As Adodc)
    Dim tabla As String
    Dim descripcion As String
    Dim j, K, i, O As Integer
    Dim comand As String
    Dim CODIGO As String
    Dim montototal As Double
    Dim montoefectivo As String
    Dim totalventa As String
    Dim FPAGO As String
    Dim LAR As Integer
    Dim PAGADO As Double
    Dim MONTOPAGADO As String
    Dim CODIPAGO As Integer
    Dim Descuento As Double
    Dim PORCENTAJE As Double
    Dim porcentaje2 As Double
    Dim descuento2 As Double
    Dim descuento3 As Double
    Dim montodescuento As String
    Dim s As String
    Dim totallinea As Double
    Dim cantidad As String
    Dim rut_cli As String
    Dim dv As String
    Dim rutdv As String
    Dim cajero As String
    Dim descu As Double
    
    Call imprime_comando("482")
    

    Call imprime_comando("001")
    Call imprime_comando("34255255")
    Call imprime_comando("111   FONOS : 2851284 - 2851287                ")
    Call imprime_comando("111--------------------------------------------")
    Call imprime_comando("111CAJERO(A) : " & cajero & " ")
    Call imprime_comando("111VEN.(A) : " & PVentas.lblVendedor & " ")
    Call imprime_comando("111NUMERO OFFSET: " & numero_venta & "")
    Call imprime_comando("111--------------------------------------------")
    Call imprime_comando("12")
    pide_numero_boleta = True

    If numero_venta = "" Then
        numero_venta = "0"
    End If
    If Mid(numero_venta, 1, 1) < Chr$(14) Then
        numero_venta = Mid(numero_venta, 2, Len(numero_venta))
    End If

    pide_numero_boleta = False

    tabla = "SELECT dd.numero, DATE_FORMAT(dd.fecha,'%d-%m-%Y') AS fecha, dd.codigo,dd.descripcion,dd.precio, dd.cantidad,dd.total, dd.descuento2 , dd.total, dd.descuento " '-- , mv.nombre "
    tabla = tabla & "FROM sv_documento_detalle_" + empresaActiva + " as dd " '-- INNER JOIN " & baseventas & ".sv_maestrovendedores as mv ON dc.rut=mv.rut "
    tabla = tabla & "WHERE dd.local = '" & empresaActiva & "' AND dd.numero = '" & numero_venta & "' AND dd.tipo = 'BV' "
    tabla = tabla & "ORDER BY dd.linea ASC"
    Descuento = 0
    descuento2 = 0
    descuento3 = 0
    montototal = 0
    Call ConectarControlData(rollo, servidor, baseVentas & empresaActiva, usuario, password, tabla)
    If rollo.Recordset.RecordCount > 0 Then

    While Not rollo.Recordset.EOF
                  
           descripcion = rollo.Recordset.Fields("descripcion")
           PORCENTAJE = rollo.Recordset.Fields("descuento")
           cantidad = rollo.Recordset.Fields("cantidad")
           porcentaje2 = rollo.Recordset.Fields("descuento2")
          
           
           
           preciounitario = rollo.Recordset.Fields("precio")
           
           If cantidad <> 0 Then
           total = cantidad * preciounitario
           totallinea = total
           
           descuento2 = PVentas.dato12.text
         
          ' descu = Int((total * porcentaje2 / 100) + 0.5)
           Descuento = Descuento + Int(((total) * PORCENTAJE / 100) + 0.5)
           

            decimales = (cantidad) - Int(cantidad)
            decimales = decimales * 10000
            decimales = Int(decimales)
            decimales = decimales / 10000
            enteros = Int(Val(cantidad))
            decimales = Mid(decimales, 1, 5)
            cantidad = Val(enteros) + (decimales)
            decimales = Mid(decimales, 3, 3)

'            If total_productos <> Int(precio * litros + 0.5) Then
'                     litros = 1
'                     precio = total_productos
'                     enteros = "000001"
'                     decimales = "000"
'            End If
' rem cuadra con ceros


            For j = 1 To (6 - Len(enteros))
                     enteros = "0" & enteros
            Next j
            For K = 1 To (3 - Len(decimales))
                   decimales = decimales & "0"
            Next K
            For i = 1 To (9 - Len(preciounitario))
                 preciounitario = " " & preciounitario
            Next i
            CODIGO = rollo.Recordset.Fields("Codigo")
            CODIGO = Trim(CODIGO)
            For i = 1 To (13 - Len(CODIGO))
                    CODIGO = CODIGO & "a"
            Next i

            For i = 1 To (9 - Len(total))
                  total = " " & total
            Next i
            unidades = enteros & decimales
            
            If cantidad > 1 Then comand = "13111" & CODIGO & unidades & preciounitario & total & descripcion & ""
            If cantidad = 1 Then comand = "13112" & CODIGO & unidades & preciounitario & total & descripcion & ""
            Call imprime_comando(comand)

            montototal = montototal + totallinea

        End If
        rollo.Recordset.MoveNext
        
        Wend
        Descuento = Descuento + descuento2
        If Descuento <> 0 Then
        K = Len(Str(Descuento)) - 1
        montodescuento = String(9 - K, 32) + Mid(Str(Descuento), 2, K)
            
        Call imprime_comando("1711" & montodescuento & "Descuento Total               ")
            
     
        End If
        
        totalventa = Str(montototal - Descuento)

        K = Len(Str(totalventa)) - 1
        totalventa = String(10 - K, "0") + Mid(Str(totalventa), 2, K)
        

        Call imprime_comando("19" & montototal & "")
        Call imprime_comando("20" & totalventa & "")
    For O = 1 To detallePagos.pagos.Rows - 2
        If CDbl(detallePagos.pagos.Cell(O, 2).text) <> 0 Then
            PAGADO = CDbl(detallePagos.pagos.Cell(O, 2).text)
            K = Len(Str(PAGADO)) - 1
            MONTOPAGADO = String(10 - K, "0") + Mid(Str(PAGADO), 2, K)
            
            If Val(Mid(detallePagos.pagos.Cell(O, 1).text, 1, 2)) < 10 Then
            FPAGO = "0" + Mid(detallePagos.pagos.Cell(O, 1).text, 1, 1)
            Else
            FPAGO = Mid(detallePagos.pagos.Cell(O, 1).text, 1, 2)
           
            End If
            glosapago = Mid(detallePagos.pagos.Cell(O, 1).text, 4, 20)
            If FPAGO > "10" Then FPAGO = "10"
            Call imprime_comando("2611" & FPAGO & MONTOPAGADO & "")
        End If
     
    Next O
       Call imprime_comando("271000000000") 'fin venta
       Call imprime_comando("28")
       Call Comentarios_Boleta
        Call imprime_comando("99") '  cortar
    
    If FPAGO = "08" Then
    IMPRIMEVALECREDITO
    
    Else
       
    End If
    s = numeroboleta
End If


End Sub

Sub IMPRIMEVALECREDITO()
        Dim comand As String
        Dim j As Integer
        
       Call LEErcuotas(PVentas.dato1.text, PVentas.dato2.text)
        For j = 1 To 2
        comand = "60"
        Call imprime_comando(comand)
        Call imprime_comando("111.")
        Call imprime_comando(30)
        
        comand = "114" + leerNombreEmpresa(empresaActiva)
        Call imprime_comando(comand)
        Call imprime_comando(30)
        
        comand = "114    COMPROBANTE DE CREDITO"
        Call imprime_comando(comand)
        
        comand = "114TIPO    : " & PVentas.dato1.text
        Call imprime_comando(comand)
        
        comand = "114NUMERO  : " & PVentas.dato2.text
        Call imprime_comando(comand)
        
        comand = "114FECHA  : " & Format(fechasistema, "dd-mm-yyyy")
        Call imprime_comando(comand)
        
        comand = "114RUT     : " & Mid(clientecuotas, 1, 9) & "-" & Mid(clientecuotas, 10, 1)
        Call imprime_comando(comand)
        
        comand = "114NOMBRE  : " & nombreclientecuotas
        Call imprime_comando(comand)
        Call imprime_comando(30)  'enter
        comand = "114CREDITO  : " & Format(montocredito, "$ ###,###,###")
        Call imprime_comando(comand)
        Call imprime_comando(30)  'enter



        comand = "112 Segun Contrato Autorizo Cargar a mi "
        Call imprime_comando(comand)
        comand = "112 cuenta"
        Call imprime_comando(comand)
        
        Call imprime_comando(30)
        comand = "111TOTAL DE CUOTAS : " & cantidaddecuotas
        Call imprime_comando(comand)
        comand = "111MONTO CUOTAS    : " & Format(montocuotas, "$ ###,###")
        Call imprime_comando(comand)
        comand = "111PRIMER VENCIMIENTO : " & primervencimiento & ""
        Call imprime_comando(comand)
        Call imprime_comando(30)
        comand = "111IMPORTANTE:  " & Chr$(10)
        Call imprime_comando(comand)
        comand = "111GUARDE ESTE DCTO. COMO COMPROBANTE"
        Call imprime_comando(comand)
        comand = "114    ***** GRACIAS POR SU COMPRA *****"
        Call imprime_comando(comand)
        Call imprime_comando(30)
        Call imprime_comando(30)
        Call imprime_comando(30)
        comand = "114         FIRMA CLIENTE ______________"
        Call imprime_comando(comand)
        Call imprime_comando(99)
        Call imprime_comando(61)
        Next j
        
End Sub
Public Function imprime_comando(comando As String) As String
    Dim str_envio As String
    Dim STR_RESPUESTA As String
    str_envio = Chr$(135)
    str_envio = str_envio & comando
    str_envio = str_envio & Chr$(136)


    List1.AddItem (str_envio)
    comandofiscal = comando


    FISCAL.Output = str_envio

    Sleep (30)
    FISCAL_OnComm





  Timer1.Enabled = True
End Function


Public Sub Comentarios_Boleta()
       Call imprime_comando("111----------------------------------------")
      'Call imprime_comando("114       Agradecemos su Preferencia")
      'Call imprime_comando("111                                        ")
      'Call imprime_comando("111----------------------------------------")
      
      Call imprime_comando("114" + glosapago + " ")
      Call imprime_comando("111                                        ")
      Call imprime_comando("116       ALMACENES ELTIT               ")
      Call imprime_comando("114       GRACIAS POR PREFERIRNOS....       ")
      
      Call imprime_comando("111-----------------------------------------")
End Sub

Private Sub Timer2_Timer()
     Timer2.Enabled = False
    Command1.Enabled = True
    Command2.Enabled = True
    Label2 = "Informe Emitido"
    Unload Me
End Sub

Private Sub Timer3_Timer()
    Command1.Enabled = True
    Command2.Enabled = True
    Timer3.Enabled = False
    Label2 = "No se ha recibido respuesta de la impresora" & vbCrLf & "Verifique que este encendida y conectada y reintente"
End Sub
Sub LEErcuotas(TIPO, NUMERO)

        Dim cSql As rdoQuery
        Dim resultado As rdoResultset
        Dim i As Integer
        Dim diasmora As Double
        Dim saldo As Double
        Dim interes As Double
        
        Set cSql = New rdoQuery
        Set cSql.ActiveConnection = ventas

        cSql.sql = "SELECT dc.rut,dc.montocredito,dc.cantidadcuotas,dc.montocuota,dc.vencimientooriginal,mc.nombre "
        cSql.sql = cSql.sql & "FROM sv_cuotas_detalle as dc," + baseVentas + ".sv_maestroclientes as mc "
        cSql.sql = cSql.sql & "WHERE tipo = '" & TIPO & "' and numero='" & NUMERO & "' and dc.rut=mc.rut  "
        cSql.sql = cSql.sql & "order by vencimientooriginal limit 0,1"
        cSql.Execute
        
        If cSql.RowsAffected > 0 Then

            Set resultado = cSql.OpenResultset
            
        
        While Not resultado.EOF
            clientecuotas = resultado(0)
            montocredito = resultado(1)
            cantidaddecuotas = resultado(2)
            montocuotas = resultado(3)
            primervencimiento = Format(resultado(4), "dd-mm-yyyy")
            nombreclientecuotas = resultado(5)
            
            
            resultado.MoveNext
            Wend
        
        End If
        Set resultado = Nothing
        cSql.Close
        Set cSql = Nothing
    End Sub


Sub IMPRIMEPAGOFISCAL()
        Dim comand As String
        Dim j As Integer
        Dim K As Integer
        
'       Call LEErcuotas(PVentas.dato1.text, PVentas.dato2.text)
        
 With creditotmppago
        For j = 1 To 2
        comand = "60"
       
        Call imprime_comando(comand)
        Call imprime_comando("111.")
        Call imprime_comando(30)
        
        comand = "114" + leerNombreEmpresa(empresaActiva)
        Call imprime_comando(comand)
        Call imprime_comando(30)
        
        comand = "114    COMPROBANTE DE PAGO - CREDITO"
        Call imprime_comando(comand)
       
        
        comand = "114NUMERO  : " & .FOLIO.Caption
        Call imprime_comando(comand)
        
        comand = "114FECHA  : " & .dia.Caption & "-" & .mes.Caption & "-" & .año.Caption
        Call imprime_comando(comand)
        
        comand = "114RUT     : " & .rut2.Caption & "-" & .lblDV.Caption
        Call imprime_comando(comand)
        
        comand = "114NOMBRE  : " & .LBLNOMBRE.Caption
        Call imprime_comando(comand)
        Call imprime_comando(30)  'enter
        comand = "114PAGADO  : " & Format(.total.Caption, "$ ###,###,###")
        Call imprime_comando(comand)
        Call imprime_comando(30)  'enter



'        comand = "112 Segun Contrato Autorizo Cargar a mi "
'        Call imprime_comando(comand)
'        comand = "112 cuenta"
'        Call imprime_comando(comand)
        
        Call imprime_comando(30)
'        comand = "111TOTAL DE CUOTAS PAGADAS : " & .GRID1.Rows - 1
'        Call imprime_comando(comand)
'
'        For k = 1 To .GRID1.Rows - 1
'        comand = "111NUMERO DOC : " & .GRID1.Cell(k, 4).text & ""
'        Call imprime_comando(comand)
'        comand = "111NUMERO CUOTA : " & .GRID1.Cell(k, 5).text & ""
'        Call imprime_comando(comand)
'        comand = "111MONTO CUOTA    : " & Format(.GRID1.Cell(k, 6).text, "$ ###,###")
'        Call imprime_comando(comand)
'        Next k
'
        Call imprime_comando(30)
        comand = "111IMPORTANTE:  " & Chr$(10)
        Call imprime_comando(comand)
        comand = "111GUARDE ESTE DCTO. COMO COMPROBANTE"
        Call imprime_comando(comand)
        comand = "114ALMACENES ELTIT AGRADECE SU PREFERENCIA"
        Call imprime_comando(comand)
        Call imprime_comando(30)
        Call imprime_comando(30)
        Call imprime_comando(30)
'        comand = "114         FIRMA CLIENTE ______________"
'        Call imprime_comando(comand)
        Call imprime_comando(99)
        Call imprime_comando(61)
        Next j
        End With
        imprimepago = "no"
        sw = False
        
End Sub
Sub IMPRIMErepactacion()
        Dim comand As String
        Dim j As Integer
        With repactacion
        For j = 1 To 2
        comand = "60"
        
        Call imprime_comando(comand)
        Call imprime_comando("111.")
        Call imprime_comando(30)
        
        comand = "114" + leerNombreEmpresa(empresaActiva)
        Call imprime_comando(comand)
        Call imprime_comando(30)
        
        comand = "114    COMPROBANTE DE CREDITO"
        Call imprime_comando(comand)
        
        comand = "114TIPO    : IM "
        Call imprime_comando(comand)
        
        comand = "114NUMERO  : " & .numerorepa
        Call imprime_comando(comand)
        
        comand = "114FECHA  : " & Format(fechasistema, "dd-mm-yyyy")
        Call imprime_comando(comand)
        
        comand = "114RUT     : " & .rut2.text & "-" & .lblDV.Caption
        Call imprime_comando(comand)
        
        comand = "114NOMBRE  : " & .LBLNOMBRE.Caption
        Call imprime_comando(comand)
        Call imprime_comando(30)  'enter
        
        comand = "114ABONO REPACTACION  : " & .MONTO.text
        Call imprime_comando(comand)
        Call imprime_comando(30)  'enter
               
        
        comand = "114CREDITO  : " & Format(.TOTALREPACTACION.Caption, "$ ###,###,###")
        Call imprime_comando(comand)
        Call imprime_comando(30)  'enter



        comand = "112 Segun Contrato Autorizo Cargar a mi "
        Call imprime_comando(comand)
        comand = "112 cuenta"
        Call imprime_comando(comand)
        
        Call imprime_comando(30)
        comand = "111TOTAL DE CUOTAS : " & .numerocuota.text
        Call imprime_comando(comand)
        comand = "111MONTO CUOTAS    : " & Format(.VALORCUOTA.text, "$ ###,###")
        Call imprime_comando(comand)
        comand = "111PRIMER VENCIMIENTO : " & .DIAC.text & "-" & .MESC.text & "-" & .AÑOC.text & ""
        Call imprime_comando(comand)
        Call imprime_comando(30)
        comand = "111IMPORTANTE:  " & Chr$(10)
        Call imprime_comando(comand)
        comand = "111GUARDE ESTE DCTO. COMO COMPROBANTE"
        Call imprime_comando(comand)
        comand = "114    ***** GRACIAS POR SU COMPRA *****"
        Call imprime_comando(comand)
        Call imprime_comando(30)
        Call imprime_comando(30)
        Call imprime_comando(30)
        comand = "114         FIRMA CLIENTE ______________"
        Call imprime_comando(comand)
        Call imprime_comando(99)
        Call imprime_comando(61)
        Next j
        repactacionimprimir = "no"
        sw = False
        End With
End Sub


